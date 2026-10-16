# Graph API client

Fourteen calls over `fetch`, behind one interface. No MSAL, no Graph SDK: the token endpoint is two
form POSTs, and an interface over `fetch` is what the tests replace with a fake.

## Calls

| Method | HTTP | Notes |
|---|---|---|
| `authUrl` | — | `/{tenant}/oauth2/v2.0/authorize`, PKCE S256, `prompt=select_account` |
| `exchangeCode` | `POST /{tenant}/oauth2/v2.0/token` | `authorization_code` with `code_verifier`; returns `id_token` |
| `refresh` | `POST …/token` | `refresh_token`; **returns a new refresh token** |
| `mailboxTimeZone` | `GET /me/mailboxSettings/timeZone` | Windows name or IANA; 403 without `MailboxSettings.Read` |
| `listCalendars` | `GET /me/calendars` | Paged with `@odata.nextLink` |
| `listCalendarView` | `GET /me/calendars/{id}/calendarView` | Expands recurrences; `Prefer: outlook.timezone="UTC"` on every page |
| `findEventsByMarker` | `GET /me/calendars/{id}/events?$filter=…` | Extended-property filter |
| `createEvent` | `POST /me/calendars/{id}/events` | With `transactionId` |
| `updateEvent` | `PATCH /me/calendars/{id}/events/{eventId}` | Partial update |
| `deleteEvent` | `DELETE /me/calendars/{id}/events/{eventId}` | 404 when already gone |
| `createSubscription` | `POST /subscriptions` | Validation handshake first; 409 if the resource already has one |
| `renewSubscription` | `PATCH /subscriptions/{id}` | New `expirationDateTime`; 404 when Graph dropped it |
| `deleteSubscription` | `DELETE /subscriptions/{id}` | 404 when already gone |
| `listSubscriptions` | `GET /subscriptions` | This app's subscriptions for the signed-in user |

## Facts that shape the code

- **`calendarView` is the right list, not `events`.** It expands recurring series into occurrences
  inside `startDateTime`–`endDateTime`, each with its own id and times. `/events` returns series
  masters with a recurrence pattern you would have to expand yourself.
- **Times come back as wall clock plus a zone name**, not as instants: `{ dateTime:
  '2026-10-20T07:00:00.0000000', timeZone: 'UTC' }`, with seven fractional digits. The default zone is
  UTC, but the client sends `Prefer: outlook.timezone="UTC"` on every request so nothing depends on a
  default, and `utcInstant` refuses any other zone rather than misread it.
- **`@odata.nextLink` is an absolute URL carrying the query.** Follow it as given, with the same
  `Prefer` header. The client refuses a next link that is not on `graph.microsoft.com`, so a bug or a
  tampered response cannot send the bearer token elsewhere.
- **Calendar and event ids are long base64 strings** (`AAMkAGI2…AAA=`) that can contain `/`, `+` and
  `=`. Every id in a path goes through `encodeURIComponent`. Outlook calendar ids run to about 150
  characters, which matters for `clientState` ([subscriptions.md](subscriptions.md)).
- **Extended properties are filtered with OData.** The value is a string literal in which `'` is
  doubled. Graph matches the filter in Exchange; the client checks the value again on the results.
- **`singleValueExtendedProperties` are only returned when expanded**, filtered to the one property
  id, which is why `listCalendarView` takes `markerId`.
- **Throttling** is a 429 (sometimes 503/504) with `Retry-After` in seconds. Outlook allows each app 10,000
  requests per mailbox per 10 minutes, and 4 concurrent requests per mailbox. The client waits out a
  `Retry-After` of up to 10 s, twice; longer waits fail the call, the sync records the error, and the
  next notification or cron run retries.
- **Errors** come back as `{ error: { code, message } }` from Graph and `{ error, error_description }`
  from the token endpoint. `GraphApiError.code` normalises both: `invalid_grant`,
  `interaction_required`, `ErrorItemNotFound`, `ErrorAccessDenied`, `TooManyRequests`.

## Code

```ts
// microsoft-calendar/graphApi.ts
const GRAPH = 'https://graph.microsoft.com/v1.0';
const LOGIN = 'https://login.microsoftonline.com';

/** Requested on connect. `openid profile` return the id token that identifies the account. */
export const REQUESTED_SCOPES = [
  'openid',
  'profile',
  'offline_access',
  'Calendars.ReadWrite',
  'MailboxSettings.Read',
] as const;

export class GraphApiError extends Error {
  readonly status: number;
  /** Graph's `error.code` (`ErrorItemNotFound`…), or the token endpoint's `error` (`invalid_grant`…). */
  readonly code: string;

  constructor(status: number, code: string, message: string) {
    super(message);
    this.name = 'GraphApiError';
    this.status = status;
    this.code = code;
  }
}

export interface TokenResponse {
  access_token: string;
  expires_in: number;
  /** Microsoft rotates refresh tokens: store the new one every time. */
  refresh_token?: string;
  scope?: string;
  id_token?: string;
  token_type: string;
}

export interface GraphCalendar {
  id: string;
  name?: string;
  canEdit?: boolean;
  isDefaultCalendar?: boolean;
}

/** Wall clock in `timeZone`, without an offset: `2026-10-20T09:00:00.0000000`. */
export interface GraphDateTime {
  dateTime?: string;
  timeZone?: string;
}

export interface GraphAttendee {
  emailAddress: { address: string; name?: string };
  type: 'required' | 'optional';
}

export interface GraphEvent {
  id?: string;
  subject?: string;
  body?: { contentType: 'text' | 'html'; content: string };
  start?: GraphDateTime;
  end?: GraphDateTime;
  isAllDay?: boolean;
  isCancelled?: boolean;
  showAs?: 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown';
  responseStatus?: { response?: string };
  attendees?: GraphAttendee[];
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: 'teamsForBusiness' | 'skypeForBusiness' | 'skypeForConsumer' | 'unknown';
  onlineMeeting?: { joinUrl?: string } | null;
  singleValueExtendedProperties?: { id: string; value: string }[];
  /** Graph de-duplicates a retried POST that carries the same value. */
  transactionId?: string;
  lastModifiedDateTime?: string;
}

export interface EventsPage {
  items: GraphEvent[];
  /** Absolute URL of the next page, exactly as Graph returned it. */
  nextLink?: string;
}

export interface GraphSubscriptionResource {
  id: string;
  resource: string;
  expirationDateTime: string;
}

export interface NewSubscription {
  changeType: string;
  resource: string;
  notificationUrl: string;
  lifecycleNotificationUrl?: string;
  expirationDateTime: string;
  clientState: string;
}

export interface GraphApi {
  authUrl(params: { state: string; codeChallenge: string }): string;
  exchangeCode(code: string, codeVerifier: string): Promise<TokenResponse>;
  refresh(refreshToken: string): Promise<TokenResponse>;
  /** The mailbox's zone as Outlook stores it: a Windows name ("W. Europe Standard Time") or IANA. */
  mailboxTimeZone(accessToken: string): Promise<string>;
  listCalendars(accessToken: string): Promise<GraphCalendar[]>;
  /** One page of the expanded calendar view, times in UTC. Pass the previous page's `nextLink` to continue. */
  listCalendarView(
    accessToken: string,
    calendarId: string,
    query: { start: Date; end: Date; markerId: string; nextLink?: string }
  ): Promise<EventsPage>;
  findEventsByMarker(
    accessToken: string,
    calendarId: string,
    markerId: string,
    value: string
  ): Promise<GraphEvent[]>;
  createEvent(accessToken: string, calendarId: string, event: GraphEvent): Promise<GraphEvent>;
  updateEvent(accessToken: string, calendarId: string, eventId: string, event: GraphEvent): Promise<GraphEvent>;
  deleteEvent(accessToken: string, calendarId: string, eventId: string): Promise<void>;
  createSubscription(accessToken: string, subscription: NewSubscription): Promise<GraphSubscriptionResource>;
  renewSubscription(accessToken: string, graphId: string, expiresAt: Date): Promise<GraphSubscriptionResource>;
  deleteSubscription(accessToken: string, graphId: string): Promise<void>;
  listSubscriptions(accessToken: string): Promise<GraphSubscriptionResource[]>;
}

interface ErrorBody {
  error?: string | { code?: string; message?: string };
  error_description?: string;
}

async function toError(response: Response): Promise<GraphApiError> {
  let body: ErrorBody = {};
  try {
    body = (await response.json()) as ErrorBody;
  } catch {
    // Not JSON: keep the status.
  }
  if (typeof body.error === 'string') {
    return new GraphApiError(response.status, body.error, body.error_description ?? body.error);
  }
  const code = body.error?.code ?? `http_${response.status}`;
  return new GraphApiError(response.status, code, body.error?.message ?? code);
}

/** OData string literal: single quotes are doubled. */
function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

const RETRYABLE = new Set([429, 503, 504]);

export function createGraphApi(options: {
  clientId: string;
  clientSecret: string;
  tenant: string;
  redirectUri: string;
  fetch?: typeof fetch;
  /** Retries of a throttled call (429/503/504), honouring Retry-After. Default 2. */
  maxRetries?: number;
  /** Longest Retry-After waited for in-process; longer ones fail and the next sync retries. Default 10 s. */
  maxRetryWaitMs?: number;
  sleep?: (ms: number) => Promise<void>;
}): GraphApi {
  const http = options.fetch ?? fetch;
  const maxRetries = options.maxRetries ?? 2;
  const maxRetryWaitMs = options.maxRetryWaitMs ?? 10_000;
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>(resolve => setTimeout(resolve, ms)));
  const tokenUrl = `${LOGIN}/${encodeURIComponent(options.tenant)}/oauth2/v2.0/token`;

  async function call<T>(
    accessToken: string,
    method: string,
    url: string,
    init: { body?: unknown; headers?: Record<string, string> } = {}
  ): Promise<T> {
    // A nextLink is followed as given; refuse one that would send the token anywhere but Graph.
    if (!url.startsWith(`${GRAPH}/`)) throw new Error(`Refusing to call ${new URL(url).origin}`);
    for (let attempt = 0; ; attempt++) {
      const response = await http(url, {
        method,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          ...(init.body === undefined ? {} : { 'Content-Type': 'application/json' }),
          ...init.headers,
        },
        body: init.body === undefined ? undefined : JSON.stringify(init.body),
      });
      if (RETRYABLE.has(response.status) && attempt < maxRetries) {
        const waitMs = Number(response.headers.get('retry-after') ?? '2') * 1000;
        if (Number.isFinite(waitMs) && waitMs <= maxRetryWaitMs) {
          await sleep(Math.max(waitMs, 0));
          continue;
        }
      }
      if (!response.ok) throw await toError(response);
      if (response.status === 204) return undefined as T;
      const text = await response.text();
      return (text ? JSON.parse(text) : undefined) as T;
    }
  }

  async function token(params: Record<string, string>): Promise<TokenResponse> {
    const response = await http(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: options.clientId,
        client_secret: options.clientSecret,
        scope: REQUESTED_SCOPES.join(' '),
        ...params,
      }),
    });
    if (!response.ok) throw await toError(response);
    return (await response.json()) as TokenResponse;
  }

  // Calendar and event ids are base64 and may contain '/', '+' and '='.
  const calendarUrl = (calendarId: string) => `${GRAPH}/me/calendars/${encodeURIComponent(calendarId)}`;
  const eventUrl = (calendarId: string, eventId: string) =>
    `${calendarUrl(calendarId)}/events/${encodeURIComponent(eventId)}`;

  async function pages<T>(accessToken: string, first: string): Promise<T[]> {
    const items: T[] = [];
    let next: string | undefined = first;
    while (next) {
      const page: { value?: T[]; '@odata.nextLink'?: string } = await call(accessToken, 'GET', next);
      items.push(...(page.value ?? []));
      next = page['@odata.nextLink'];
    }
    return items;
  }

  return {
    authUrl({ state, codeChallenge }) {
      const url = new URL(`${LOGIN}/${encodeURIComponent(options.tenant)}/oauth2/v2.0/authorize`);
      url.search = new URLSearchParams({
        client_id: options.clientId,
        response_type: 'code',
        response_mode: 'query',
        redirect_uri: options.redirectUri,
        scope: REQUESTED_SCOPES.join(' '),
        state,
        code_challenge: codeChallenge,
        code_challenge_method: 'S256',
        // Not 'consent': forcing the consent screen shows "Need admin approval" to users of a
        // tenant that disabled user consent, even after an administrator granted it.
        prompt: 'select_account',
      }).toString();
      return url.toString();
    },

    exchangeCode: (code, codeVerifier) =>
      token({
        grant_type: 'authorization_code',
        code,
        code_verifier: codeVerifier,
        redirect_uri: options.redirectUri,
      }),

    refresh: refreshToken => token({ grant_type: 'refresh_token', refresh_token: refreshToken }),

    async mailboxTimeZone(accessToken) {
      const result = await call<{ value?: string }>(accessToken, 'GET', `${GRAPH}/me/mailboxSettings/timeZone`);
      return result.value ?? '';
    },

    listCalendars: accessToken =>
      pages<GraphCalendar>(
        accessToken,
        `${GRAPH}/me/calendars?$top=100&$select=id,name,canEdit,isDefaultCalendar`
      ),

    async listCalendarView(accessToken, calendarId, { start, end, markerId, nextLink }) {
      let url = nextLink;
      if (!url) {
        const query = new URLSearchParams({
          startDateTime: start.toISOString(),
          endDateTime: end.toISOString(),
          $top: '500',
          $select: 'id,isAllDay,isCancelled,showAs,responseStatus,start,end,lastModifiedDateTime',
          $expand: `singleValueExtendedProperties($filter=id eq ${odataString(markerId)})`,
        });
        url = `${calendarUrl(calendarId)}/calendarView?${query.toString()}`;
      }
      const page = await call<{ value?: GraphEvent[]; '@odata.nextLink'?: string }>(accessToken, 'GET', url, {
        // Timed events come back as UTC wall clock; every page needs the header, not only the first.
        headers: { Prefer: 'outlook.timezone="UTC"' },
      });
      const next = page['@odata.nextLink'];
      return { items: page.value ?? [], ...(next ? { nextLink: next } : {}) };
    },

    async findEventsByMarker(accessToken, calendarId, markerId, value) {
      const query = new URLSearchParams({
        $filter: `singleValueExtendedProperties/Any(ep: ep/id eq ${odataString(markerId)} and ep/value eq ${odataString(value)})`,
        $select: 'id,subject,start,end,onlineMeeting',
        $expand: `singleValueExtendedProperties($filter=id eq ${odataString(markerId)})`,
      });
      const events = await call<{ value?: GraphEvent[] }>(
        accessToken,
        'GET',
        `${calendarUrl(calendarId)}/events?${query.toString()}`,
        { headers: { Prefer: 'outlook.timezone="UTC"' } }
      );
      // The filter is evaluated by Exchange; check the value here too rather than trusting it.
      return (events.value ?? []).filter(event =>
        event.singleValueExtendedProperties?.some(p => p.id.toLowerCase() === markerId.toLowerCase() && p.value === value)
      );
    },

    createEvent: (accessToken, calendarId, event) =>
      call<GraphEvent>(accessToken, 'POST', `${calendarUrl(calendarId)}/events`, { body: event }),

    updateEvent: (accessToken, calendarId, eventId, event) =>
      call<GraphEvent>(accessToken, 'PATCH', eventUrl(calendarId, eventId), { body: event }),

    async deleteEvent(accessToken, calendarId, eventId) {
      await call<void>(accessToken, 'DELETE', eventUrl(calendarId, eventId));
    },

    createSubscription: (accessToken, subscription) =>
      call<GraphSubscriptionResource>(accessToken, 'POST', `${GRAPH}/subscriptions`, { body: subscription }),

    renewSubscription: (accessToken, graphId, expiresAt) =>
      call<GraphSubscriptionResource>(accessToken, 'PATCH', `${GRAPH}/subscriptions/${encodeURIComponent(graphId)}`, {
        body: { expirationDateTime: expiresAt.toISOString() },
      }),

    async deleteSubscription(accessToken, graphId) {
      await call<void>(accessToken, 'DELETE', `${GRAPH}/subscriptions/${encodeURIComponent(graphId)}`);
    },

    listSubscriptions: accessToken => pages<GraphSubscriptionResource>(accessToken, `${GRAPH}/subscriptions`),
  };
}
```

## Checklist

- [ ] Every calendar and event id in a path goes through `encodeURIComponent`
- [ ] `listCalendarView` callers loop on `nextLink` and abort on any failed page
- [ ] `Prefer: outlook.timezone="UTC"` on every calendar-view and marker request
- [ ] Errors branch on `GraphApiError.status` / `.code`, never on message text
