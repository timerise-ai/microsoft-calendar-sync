# HTTP routes

Seven handlers on the Web `Request`/`Response` API, so they mount unchanged in Next.js route handlers,
Hono, Bun or Deno, and through a small adapter in Express.

| Route | Method | Auth | Answers |
|---|---|---|---|
| `/calendar/microsoft/connect` | GET | user | `{ url }` to navigate to |
| `/calendar/microsoft/callback` | GET | none (Microsoft redirect; the state is the credential) | 303 to `appReturnUrl?calendar=connected\|denied\|admin-consent-required\|<error code>` |
| `/calendar/microsoft/calendars` | GET | user | `{ connection, calendars[] }` with health |
| `/calendar/microsoft/calendars/:calendarId` | PATCH | user | the calendar view; 404, 409 on guard failures |
| `/calendar/microsoft/disconnect` | POST | user | 204 |
| `/calendar/microsoft/notifications` | POST | none (signed clientState) | 200 text/plain for validation, 202 for notifications, 400 for a malformed body |
| `/calendar/microsoft/cron` | POST | `Bearer CRON_SECRET` | the renewal report |

- **Validation first.** Creating a subscription makes Graph POST `?validationToken=...` to the URL and
  wait up to 10 seconds for the token back, decoded, as `text/plain` with 200. The subscription is
  only created if that succeeds, so the route must be deployed and public before anyone turns on busy
  sync. Lifecycle URLs are validated the same way.
- **Answer within 3 seconds; sync afterwards.** Graph treats a slower answer as a failure. Past 10 % of
  slow answers in ten minutes it delays every notification by ten minutes; past 15 % over 10 seconds it
  drops them, unrecoverably. The handler verifies, collects calendar ids, answers 202, and hands the
  sync to `defer`, which the host must supply: `task => void task()` on a long-running server, `after`
  in Next.js 15+, `waitUntil` on Vercel or Cloudflare. A sync that fails afterwards is recorded on the
  row and repaired by the cron.
- **202 for anything that parses.** Graph retries any other status for up to four hours. A forged or
  stale notification is logged and acknowledged; retrying it helps nobody.
- **The callback always redirects.** A browser is on the other end; a JSON error there is a dead page.
  Cancelling arrives as `?error=access_denied` (`calendar=denied`); a tenant that needs admin approval
  as `consent_required` (`calendar=admin-consent-required`), which the UI should explain.
- **The cron secret fails closed.** An empty `cronSecret` rejects everything.
- **Mount the notification route on the public HTTPS host Graph can reach**, with TLS 1.2 or later.

## Handlers

```ts
// microsoft-calendar/webhook.ts
import { syncBusy } from './busySync';
import { verifyClientState } from './crypto';
import { ensureSubscription } from './subscriptions';
import type { CalendarSyncDeps } from './types';

/** One entry of a notification POST's `value`: a change, or a lifecycle event when `lifecycleEvent` is set. */
export interface GraphNotification {
  subscriptionId?: string;
  clientState?: string;
  changeType?: string;
  lifecycleEvent?: 'subscriptionRemoved' | 'reauthorizationRequired' | 'missed' | string;
}

/** Runs the work after the response. Required: Graph counts any answer slower than 3 s as a failure. */
export type Defer = (task: () => Promise<void>) => Promise<void> | void;

export interface NotificationResult {
  status: number;
  /** The validation handshake's plain-text answer. */
  text?: string;
  /** Syncs to run after answering. */
  work?: () => Promise<void>;
}

/** Graph's validation tokens are short opaque strings; anything else is not worth echoing. */
const VALIDATION_TOKEN_MAX = 1024;

/**
 * Graph change and lifecycle notifications. Validates, collects which calendars to sync, and
 * returns at once with the work to run afterwards: Graph wants a 2xx within 3 seconds, marks a
 * slower endpoint "slow" and then drops its notifications for good. Answers 202 for everything that
 * parses, forged or stale entries included: any other status makes Graph retry for four hours,
 * which helps nobody.
 */
export async function handleNotification(
  deps: CalendarSyncDeps,
  input: { validationToken: string | null; body: unknown }
): Promise<NotificationResult> {
  // Subscription creation: echo the token as text/plain within 10 seconds.
  if (input.validationToken !== null) {
    if (!input.validationToken || input.validationToken.length > VALIDATION_TOKEN_MAX) return { status: 400 };
    return { status: 200, text: input.validationToken };
  }

  const value = (input.body as { value?: unknown } | null)?.value;
  if (!Array.isArray(value)) return { status: 400 };

  const resync = new Set<string>();
  const renew = new Set<string>();
  for (const entry of value as (GraphNotification | null)[]) {
    // A public endpoint gets junk: an entry that is not an object, or has no string clientState, is
    // skipped like a forgery, so it neither answers 500 nor stops the entries beside it.
    if (typeof entry !== 'object' || entry === null || typeof entry.clientState !== 'string') {
      deps.log.warn('microsoft-calendar-webhook-bad-client-state', {});
      continue;
    }
    // Verified before any lookup: junk traffic costs an HMAC, not a database read.
    const subscriptionId = verifyClientState(entry.clientState, deps.config.clientStateSecret);
    if (!subscriptionId) {
      deps.log.warn('microsoft-calendar-webhook-bad-client-state', { graphId: entry.subscriptionId });
      continue;
    }
    const subscription = await deps.store.getSubscription(subscriptionId);
    if (!subscription) {
      deps.log.info('microsoft-calendar-webhook-unknown-subscription', { subscriptionId });
      continue;
    }
    // A clientState is only valid on the subscription it was issued for.
    if (subscription.graphId !== entry.subscriptionId) {
      deps.log.warn('microsoft-calendar-webhook-subscription-mismatch', { subscriptionId });
      continue;
    }
    // Removed or about to need reauthorization: renew or replace it. Missed: notifications were
    // dropped, so resync. Every other entry is a change: resync.
    if (entry.lifecycleEvent === 'subscriptionRemoved' || entry.lifecycleEvent === 'reauthorizationRequired') {
      renew.add(subscription.calendarId);
    }
    resync.add(subscription.calendarId);
  }

  if (resync.size === 0) return { status: 202 };
  return {
    status: 202,
    work: async () => {
      for (const calendarId of resync) {
        try {
          if (renew.has(calendarId)) {
            const calendar = await deps.store.getCalendar(calendarId);
            if (calendar?.syncBusy) await ensureSubscription(deps, calendar, { force: true });
          }
          await syncBusy(deps, calendarId);
        } catch (error) {
          // Recorded on the row by syncBusy; the cron retries.
          deps.log.error('microsoft-calendar-webhook-sync-failed', {
            calendarId,
            message: error instanceof Error ? error.message : String(error),
          });
        }
      }
    },
  };
}
```

```ts
// microsoft-calendar/routes.ts
import { refreshCalendarList, updateCalendarSettings } from './calendars';
import { completeConnect, disconnect, startConnect } from './connect';
import { safeEqual } from './crypto';
import { deriveCalendarHealth, type CalendarHealth } from './health';
import { renewSubscriptions } from './subscriptions';
import { handleNotification, type Defer } from './webhook';
import {
  CalendarSyncError,
  type CalendarErrorCode,
  type CalendarSettingsPatch,
  type CalendarSyncDeps,
  type LinkedCalendar,
} from './types';

export interface CalendarRouteOptions {
  /** The host's auth guard: the signed-in user's id, or null. */
  authenticate: (request: Request) => Promise<string | null>;
  /** Where the browser lands after the callback; `?calendar=<result>` is appended. */
  appReturnUrl: string;
  cronSecret: string;
  /**
   * Runs notification syncs after the response. Long-running server: `task => void task()`.
   * Next.js: `after`. Vercel/Cloudflare: `waitUntil`. Never inline: Graph needs an answer in 3 s.
   */
  defer: Defer;
  /** Older than this, a synced calendar reports `stale`. Default 26 h for a 12-hourly cron. */
  staleAfterMs?: number;
}

const STATUS: Record<CalendarErrorCode, number> = {
  'not-connected': 409,
  'reauth-required': 409,
  'calendar-not-found': 404,
  'calendar-not-writable': 409,
  'invalid-or-expired-state': 400,
  'no-refresh-token': 400,
  'no-account-id': 400,
  'missing-calendar-scope': 400,
};

const json = (body: unknown, status = 200) => Response.json(body, { status });

function errorResponse(error: unknown): Response {
  if (error instanceof CalendarSyncError) return json({ error: error.code }, STATUS[error.code]);
  throw error;
}

function parsePatch(body: unknown): CalendarSettingsPatch | null {
  if (typeof body !== 'object' || body === null) return null;
  const input = body as Record<string, unknown>;
  const patch: CalendarSettingsPatch = {};
  for (const key of Object.keys(input)) {
    const value = input[key];
    if (key === 'title' && typeof value === 'string' && value.trim() && value.length <= 200) patch.title = value.trim();
    else if ((key === 'syncBusy' || key === 'syncBookings') && typeof value === 'boolean') patch[key] = value;
    else return null;
  }
  return patch;
}

export interface CalendarView {
  id: string;
  title: string;
  isDefault: boolean;
  isWritable: boolean;
  syncBusy: boolean;
  syncBookings: boolean;
  health: CalendarHealth;
  lastSyncedAt: string | null;
  lastSyncError: string | null;
}

/** Framework-free handlers on the Web Request/Response API. Mount them in the host's router. */
export function createCalendarRoutes(deps: CalendarSyncDeps, options: CalendarRouteOptions) {
  const staleAfterMs = options.staleAfterMs ?? 26 * 60 * 60 * 1000;

  async function view(userId: string, calendars: LinkedCalendar[]) {
    const connection = await deps.store.getConnection(userId);
    if (!connection) return { connection: null, calendars: [] };
    const now = deps.now();
    const rows: CalendarView[] = [];
    for (const calendar of calendars) {
      const subscription = calendar.subscriptionId ? await deps.store.getSubscription(calendar.subscriptionId) : null;
      rows.push({
        id: calendar.id,
        title: calendar.title,
        isDefault: calendar.isDefault,
        isWritable: calendar.isWritable,
        syncBusy: calendar.syncBusy,
        syncBookings: calendar.syncBookings,
        health: deriveCalendarHealth({ connection, calendar, subscription, now, staleAfterMs }),
        lastSyncedAt: calendar.lastSyncedAt?.toISOString() ?? null,
        lastSyncError: calendar.lastSyncError,
      });
    }
    return {
      connection: {
        accountEmail: connection.accountEmail,
        status: connection.status,
        canWrite: connection.canWrite,
        mailboxTimeZone: connection.mailboxTimeZone,
      },
      calendars: rows,
    };
  }

  return {
    /** GET returns `{ url }`; the client navigates to it. */
    async connect(request: Request): Promise<Response> {
      const userId = await options.authenticate(request);
      if (!userId) return json({ error: 'unauthenticated' }, 401);
      return json({ url: await startConnect(deps, userId) });
    },

    /** GET, from Microsoft. A browser is on the other end: always redirect, never answer with JSON. */
    async callback(request: Request): Promise<Response> {
      const params = new URL(request.url).searchParams;
      const back = (result: string) => {
        const url = new URL(options.appReturnUrl);
        url.searchParams.set('calendar', result);
        return Response.redirect(url.toString(), 303);
      };
      const code = params.get('code');
      const state = params.get('state');
      // error=access_denied when the user cancels; consent_required when an admin must approve.
      const error = params.get('error');
      if (error) return back(error === 'consent_required' ? 'admin-consent-required' : 'denied');
      if (!code || !state) return back('denied');
      try {
        await completeConnect(deps, { state, code });
        return back('connected');
      } catch (failure) {
        deps.log.error('microsoft-calendar-callback-failed', {
          message: failure instanceof Error ? failure.message : String(failure),
        });
        return back(failure instanceof CalendarSyncError ? failure.code : 'failed');
      }
    },

    /** GET returns the connection, calendars and their health. Refreshes from Graph; falls back to stored rows. */
    async list(request: Request): Promise<Response> {
      const userId = await options.authenticate(request);
      if (!userId) return json({ error: 'unauthenticated' }, 401);
      if (!(await deps.store.getConnection(userId))) return json({ connection: null, calendars: [] });
      let calendars: LinkedCalendar[];
      try {
        calendars = await refreshCalendarList(deps, userId);
      } catch (error) {
        if (!(error instanceof CalendarSyncError)) {
          deps.log.warn('microsoft-calendar-list-refresh-failed', {
            userId,
            message: error instanceof Error ? error.message : String(error),
          });
        }
        calendars = await deps.store.listCalendars(userId);
      }
      return json(await view(userId, calendars));
    },

    /** PATCH `{ title?, syncBusy?, syncBookings? }`. */
    async update(request: Request, calendarId: string): Promise<Response> {
      const userId = await options.authenticate(request);
      if (!userId) return json({ error: 'unauthenticated' }, 401);
      const patch = parsePatch(await request.json().catch(() => null));
      if (!patch) return json({ error: 'invalid-body' }, 400);
      try {
        const calendar = await updateCalendarSettings(deps, userId, calendarId, patch);
        const result = await view(userId, [calendar]);
        return json(result.calendars[0] ?? null);
      } catch (error) {
        return errorResponse(error);
      }
    },

    /** POST. */
    async disconnect(request: Request): Promise<Response> {
      const userId = await options.authenticate(request);
      if (!userId) return json({ error: 'unauthenticated' }, 401);
      await disconnect(deps, userId);
      return new Response(null, { status: 204 });
    },

    /** POST, from Graph. Unauthenticated by design; the signed clientState is the credential. */
    async notifications(request: Request): Promise<Response> {
      const validationToken = new URL(request.url).searchParams.get('validationToken');
      const body = validationToken === null ? await request.json().catch(() => null) : null;
      const result = await handleNotification(deps, { validationToken, body });
      if (result.work) await options.defer(result.work);
      if (result.text !== undefined) {
        return new Response(result.text, { status: result.status, headers: { 'Content-Type': 'text/plain' } });
      }
      return new Response(null, { status: result.status });
    },

    /** POST with `Authorization: Bearer <cronSecret>`. Every 12 h with the default 48 h renewal window. */
    async cron(request: Request): Promise<Response> {
      const header = request.headers.get('authorization') ?? '';
      // Fails closed: an unset secret matches nothing.
      if (!options.cronSecret || !safeEqual(header, `Bearer ${options.cronSecret}`)) {
        return json({ error: 'unauthorized' }, 401);
      }
      return json(await renewSubscriptions(deps));
    },
  };
}
```

## Wiring

One module builds the dependencies. This example uses Firestore and environment variables; swap the
store and the declared host functions for your own. On Supabase the store line becomes
`createSupabaseCalendarStore(createClient(url, serviceRoleKey))` from
[supabase-backend.md](supabase-backend.md), with `@supabase/supabase-js` installed from the registry.

Everything is built on the first request, never at import: `next build` imports every route module
without the secrets, and a missing variable then fails the one request that needs it, by name. This is
the whole job of the module; it needs no fallback values, no `503` of its own and no other laziness.

```ts
// microsoft-calendar/setup.ts: the one place that builds the deps (example: Firestore, env vars)
import { getFirestore } from 'firebase-admin/firestore';

import { createGraphApi } from './graphApi';
import { createCalendarRoutes } from './routes';
import { createFirestoreCalendarStore } from './store/firestore';
import { DEFAULT_SYNC_SETTINGS, type BookingForCalendar, type CalendarSyncDeps } from './types';

function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`${name} is not set`);
  return value;
}

// Host seams, declared here so the example compiles: replace with your own.
declare function verifySession(request: Request): Promise<string | null>;
declare function listUpcomingBookingsForHost(userId: string, from: Date): Promise<BookingForCalendar[]>;
declare function saveMeetingUrl(bookingId: string, url: string): Promise<void>;
declare function businessTimeZoneOfHost(userId: string): Promise<string | null>;

export type CalendarRoutes = ReturnType<typeof createCalendarRoutes>;

function build(): { deps: CalendarSyncDeps; routes: CalendarRoutes } {
  const config = {
    ...DEFAULT_SYNC_SETTINGS,
    clientId: required('MICROSOFT_CALENDAR_CLIENT_ID'),
    clientSecret: required('MICROSOFT_CALENDAR_CLIENT_SECRET'),
    tenant: process.env.MICROSOFT_CALENDAR_TENANT || DEFAULT_SYNC_SETTINGS.tenant,
    redirectUri: required('MICROSOFT_CALENDAR_REDIRECT_URI'),
    notificationUrl: required('MICROSOFT_CALENDAR_NOTIFICATION_URL'),
    clientStateSecret: required('MICROSOFT_CALENDAR_CLIENT_STATE_SECRET'),
    encryptionKey: required('CALENDAR_TOKEN_KEY'),
    markerPropertySetId: required('MICROSOFT_CALENDAR_MARKER_GUID'),
    defaultTimeZone: required('MICROSOFT_CALENDAR_DEFAULT_TIME_ZONE'),
  };
  const deps: CalendarSyncDeps = {
    store: createFirestoreCalendarStore(getFirestore()),
    graph: createGraphApi(config),
    config,
    bookings: { listUpcomingForPrimaryHost: listUpcomingBookingsForHost },
    log: {
      info: (kind, data) => console.log(JSON.stringify({ severity: 'INFO', kind, ...data })),
      warn: (kind, data) => console.warn(JSON.stringify({ severity: 'WARNING', kind, ...data })),
      error: (kind, data) => console.error(JSON.stringify({ severity: 'ERROR', kind, ...data })),
    },
    now: () => new Date(),
    onConferenceLink: saveMeetingUrl,
    hostTimeZone: businessTimeZoneOfHost,
  };
  const routes = createCalendarRoutes(deps, {
    authenticate: verifySession,
    appReturnUrl: required('CALENDAR_RETURN_URL'),
    cronSecret: required('CRON_SECRET'),
    // A long-running server keeps working after the response. On serverless use the platform's
    // after-response primitive instead (Next.js `after`, Vercel `waitUntil`).
    defer: task => {
      void task();
    },
  });
  return { deps, routes };
}

let built: ReturnType<typeof build> | undefined;

/** The deps for `pushBookingEvent` and `removeBookingEvent` in the host's booking flow. */
export function calendarDeps(): CalendarSyncDeps {
  return (built ??= build()).deps;
}

export function calendarRoutes(): CalendarRoutes {
  return (built ??= build()).routes;
}
```

`authenticate` is the host's session check, and only that: it returns the user id of a verified session
or token, or `null`. It never reads a user id the client sends, in a header, the query or the body,
not even "for tests": the routes' tests pass their own `authenticate`, and a header any caller can set
attaches anyone's calendar to anyone's account.

Next.js App Router (pass `defer: after` from `next/server` in `setup.ts` there). Mount these two files as
they are: `notifications` echoes Graph's `validationToken` and verifies before it schedules anything, so
the route never parses or answers a notification itself. `cron` is on `GET` as well as `POST` because
Vercel Cron sends `GET` with `Authorization: Bearer $CRON_SECRET`; the secret check is the same.

```ts
// app/api/calendar/microsoft/[action]/route.ts: Next.js App Router
import { calendarRoutes } from '@/microsoft-calendar/setup';

export async function GET(request: Request, { params }: { params: Promise<{ action: string }> }) {
  const { action } = await params;
  if (action === 'connect') return calendarRoutes().connect(request);
  if (action === 'callback') return calendarRoutes().callback(request);
  if (action === 'calendars') return calendarRoutes().list(request);
  if (action === 'cron') return calendarRoutes().cron(request);
  return new Response(null, { status: 404 });
}

export async function POST(request: Request, { params }: { params: Promise<{ action: string }> }) {
  const { action } = await params;
  if (action === 'notifications') return calendarRoutes().notifications(request);
  if (action === 'disconnect') return calendarRoutes().disconnect(request);
  if (action === 'cron') return calendarRoutes().cron(request);
  return new Response(null, { status: 404 });
}
```

```ts
// app/api/calendar/microsoft/calendars/[calendarId]/route.ts
import { calendarRoutes } from '@/microsoft-calendar/setup';

export async function PATCH(request: Request, { params }: { params: Promise<{ calendarId: string }> }) {
  return calendarRoutes().update(request, (await params).calendarId);
}
```

Express (the adapter turns each request into a Web `Request` once):

```ts
// server/calendarRouter.ts: Express: adapt each request to the Web API once, and send the Response back.
import express, { type Request as ExpressRequest, type Response as ExpressResponse } from 'express';

import { calendarRoutes, type CalendarRoutes } from '@/microsoft-calendar/setup';

function toWebRequest(req: ExpressRequest): Request {
  const headers = new Headers();
  for (const [name, value] of Object.entries(req.headers)) {
    if (typeof value === 'string') headers.set(name, value);
    else if (Array.isArray(value)) for (const item of value) headers.append(name, item);
  }
  // Graph's validation POST has a text/plain body that express.json() leaves unparsed; the handler
  // reads only the query string for it.
  const hasBody = req.method !== 'GET' && req.method !== 'HEAD' && req.body !== undefined;
  return new Request(`${req.protocol}://${req.get('host')}${req.originalUrl}`, {
    method: req.method,
    headers,
    body: hasBody ? JSON.stringify(req.body) : undefined,
  });
}

async function send(res: ExpressResponse, response: Response): Promise<void> {
  res.status(response.status);
  response.headers.forEach((value, name) => res.setHeader(name, value));
  res.send(Buffer.from(await response.arrayBuffer()));
}

// The routes are built on the first request, like the Next.js handlers: pick the handler then.
const route = (pick: (routes: CalendarRoutes) => (request: Request) => Promise<Response>) =>
  (req: ExpressRequest, res: ExpressResponse) =>
    pick(calendarRoutes())(toWebRequest(req)).then(response => send(res, response));

export const calendarRouter = express.Router();
calendarRouter.use(express.json());
calendarRouter.get('/calendar/microsoft/connect', route(r => r.connect));
calendarRouter.get('/calendar/microsoft/callback', route(r => r.callback));
calendarRouter.get('/calendar/microsoft/calendars', route(r => r.list));
calendarRouter.patch('/calendar/microsoft/calendars/:calendarId', (req, res) =>
  calendarRoutes().update(toWebRequest(req), String(req.params.calendarId)).then(response => send(res, response))
);
calendarRouter.post('/calendar/microsoft/disconnect', route(r => r.disconnect));
calendarRouter.post('/calendar/microsoft/notifications', route(r => r.notifications));
calendarRouter.post('/calendar/microsoft/cron', route(r => r.cron));
```

## Checklist

- [ ] Notification URL public and HTTPS before busy sync is offered; same value as
  `MICROSOFT_CALENDAR_NOTIFICATION_URL`
- [ ] Callback URL registered on the app; same value as `MICROSOFT_CALENDAR_REDIRECT_URI`
- [ ] `defer` runs the work after the response on this platform; never awaited inline
- [ ] `authenticate` is the host's real session check, returning the user id
- [ ] Cron route scheduled with the bearer secret
