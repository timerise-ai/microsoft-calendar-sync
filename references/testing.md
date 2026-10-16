# Testing

The suite runs the real flows against two doubles: a store with the contract's compare-and-swap
semantics, and a Graph that pages, finds events by marker, de-duplicates by `transactionId`, refuses a
`clientState` over 128 characters, answers 409 on a second subscription for a resource, and 404 on a
subscription it no longer has. A third file tests the HTTP client itself against a scripted `fetch`.
No network, no framework: `node:test` and `node:assert`. The tests are in
[test-suite.md](test-suite.md) and [test-flows.md](test-flows.md).

## Running

Node 22.18+ runs TypeScript directly by stripping types; the templates avoid syntax that needs
transpiling (no enums, no constructor parameter properties). Node's loader needs file extensions, so
rewrite the extensionless imports in a copy:

```bash
cp -R microsoft-calendar /tmp/mc-run
find /tmp/mc-run -name '*.ts' -exec perl -pi -e "s/from '(\.{1,2}\/[^']+)'/from '\$1.ts'/g" {} +
for tz in UTC Europe/Warsaw Pacific/Auckland America/Los_Angeles; do
  TZ=$tz node --test --experimental-strip-types /tmp/mc-run/rules.test.ts /tmp/mc-run/graphApi.test.ts /tmp/mc-run/flows.test.ts
done
```

In a host with vitest or jest, the files run unchanged once the runner's TypeScript transform is on.

**Run under at least two process time zones.** All-day anchoring and DST handling must not depend on
the server's `TZ`; a test that passes only in UTC is how "works in production, fails on my laptop" starts.

## What the suite pins

| Test | Guards against |
|---|---|
| UTC wall clock read as UTC; other zones refused | Busy time shifted by an offset |
| all-day vacation at mailbox midnights; DST day 23 h | Vacations not blocking, or blocking the wrong hours |
| own bookings not imported | Group sessions blocked after the first booking |
| free / working elsewhere / declined / cancelled | Busy time the user does not have |
| clientState ≤ 128 and bound to its subscription | Graph refusing every subscription; replayed notifications |
| Windows zone names mapped | All-day events falling back to the wrong zone |
| event times: zone wall clock, UTC near DST | Bookings an hour off on the night clocks change |
| invite modes; hosts without e-mail dropped | Unwanted e-mails; Graph rejecting the event |
| no forced consent; Prefer UTC on every page; nextLink stays on Graph | Admin-approval loops; shifted pages; token leaks |
| Retry-After honoured, long waits fail | Throttled syncs failing at once, or hanging |
| connect: single use, expiry, account `tid:oid`, mailbox zone | Calendar attached to another account |
| Calendars.Read only: nothing writable | Booking writes failing with 403 |
| every page imported; failed page deletes nothing | Truncated calendars; mass deletion on a transient error |
| busy time imported when no subscription can be opened | No busy time at all while Graph refuses subscriptions |
| webhook answers before syncing; a burst converges | Graph dropping notifications; duplicate blocks |
| renew in place; 404 replaces; 409 clears | Subscriptions silently lost, forever |
| lifecycle `subscriptionRemoved` replaces | Hours without notifications |
| pushed twice is one event, no second write | Duplicate events; duplicate e-mails to attendees |
| duplicates collapse; cancel twice; re-confirm | Leftover events; errors on repeat |
| rotated refresh token stored | Connections dying after an hour |
| `invalid_grant` / `interaction_required` → reauth; reconnect keeps settings | Empty calendar lists; re-configuring after every revoke |
| other account tears down; calendar removed tears down | Stale busy time from calendars the user left |
| routes: validation text, `defer`, redirects, cron closed | Sync inline in the route; JSON to a browser |

## Proving the tests can fail

Mutate the code back toward each defect and confirm a test fails. Each of these turns the suite red:

| Mutation | File |
|---|---|
| remove the `own-booking` skip | `busyRules.ts` |
| all-day events not re-anchored | `busyRules.ts` |
| `free` blocks | `busyRules.ts` |
| never delete rows missing from `desired` | `busyRules.ts` |
| read any zone as UTC | `busyRules.ts` |
| stop after the first page | `busySync.ts` |
| accept a clientState without comparing its subscription | `webhook.ts` |
| sync before answering | `webhook.ts`, `routes.ts` |
| rethrow a 404 on renewal | `subscriptions.ts` |
| no 409 handling on create | `subscriptions.ts` |
| skip the sync when the subscription fails (toggle or cron) | `subscriptions.ts` |
| always create; rewrite unchanged; keep duplicates | `bookingEvents.ts` |
| no UTC fallback near DST | `bookingEvents.ts` |
| keep the old refresh token | `tokens.ts` |
| `interaction_required` not reauth | `tokens.ts` |
| follow any nextLink; `prompt=consent`; no throttling retry | `graphApi.ts` |
| writable without `canEdit` | `calendars.ts` |
| reconnect never the same account | `connect.ts` |
| auth request not deleted on consume | `testing/memoryStore.ts` |
| calendar data in the clientState | `crypto.ts` |

## Doubles

```ts
// microsoft-calendar/testing/memoryStore.ts
import type {
  AuthRequest,
  BusyBlock,
  CalendarConnection,
  CalendarStore,
  LinkedCalendar,
  WatchSubscription,
} from '../types';

/** Yields to the event loop, so concurrent callers interleave the way they would against a database. */
const tick = () => new Promise<void>(resolve => setImmediate(resolve));

/** The store contract in memory: a test double, and the executable statement of each method's semantics. */
export function createMemoryStore() {
  const authRequests = new Map<string, AuthRequest>();
  const connections = new Map<string, CalendarConnection>();
  const calendars = new Map<string, LinkedCalendar>();
  const subscriptions = new Map<string, WatchSubscription>();
  const busy = new Map<string, BusyBlock>();

  const store: CalendarStore = {
    async saveAuthRequest(request) {
      authRequests.set(request.state, { ...request });
    },
    async consumeAuthRequest(state) {
      const request = authRequests.get(state) ?? null;
      authRequests.delete(state);
      return request;
    },
    async deleteExpiredAuthRequests(before) {
      let count = 0;
      for (const [state, request] of authRequests) {
        if (request.expiresAt < before) {
          authRequests.delete(state);
          count++;
        }
      }
      return count;
    },

    async getConnection(userId) {
      const connection = connections.get(userId);
      return connection ? { ...connection } : null;
    },
    async saveConnection(connection) {
      connections.set(connection.userId, { ...connection });
    },
    async patchConnection(userId, patch) {
      const current = connections.get(userId);
      if (current) connections.set(userId, { ...current, ...patch });
    },
    async deleteConnection(userId) {
      connections.delete(userId);
    },

    async listCalendars(userId) {
      return [...calendars.values()].filter(c => c.userId === userId).map(c => ({ ...c }));
    },
    async getCalendar(calendarId) {
      const calendar = calendars.get(calendarId);
      return calendar ? { ...calendar } : null;
    },
    async mutateCalendar(calendarId, fn) {
      for (;;) {
        const snapshot = calendars.get(calendarId);
        const before = snapshot ? { ...snapshot } : null;
        await tick();
        const next = fn(before ? { ...before } : null);
        await tick();
        // Compare-and-swap: someone wrote in between, run `fn` again on the new state.
        if ((calendars.get(calendarId)?.version ?? 0) !== (before?.version ?? 0)) continue;
        if (!next) return { before, after: before };
        const after = { ...next, id: calendarId, version: (before?.version ?? 0) + 1 };
        calendars.set(calendarId, after);
        return { before, after: { ...after } };
      }
    },
    async deleteCalendar(calendarId) {
      calendars.delete(calendarId);
    },
    async listBusySyncCalendars(cursor, limit) {
      return [...calendars.values()]
        .filter(c => c.syncBusy && (cursor === null || c.id > cursor))
        .sort((a, b) => (a.id < b.id ? -1 : 1))
        .slice(0, limit)
        .map(c => ({ ...c }));
    },

    async saveSubscription(subscription) {
      subscriptions.set(subscription.id, { ...subscription });
    },
    async getSubscription(subscriptionId) {
      const subscription = subscriptions.get(subscriptionId);
      return subscription ? { ...subscription } : null;
    },
    async deleteSubscription(subscriptionId) {
      subscriptions.delete(subscriptionId);
    },
    async listSubscriptionsExpiringBefore(before, limit) {
      return [...subscriptions.values()].filter(s => s.expiresAt < before).slice(0, limit);
    },

    async listBusyBlocks(calendarId) {
      await tick();
      return [...busy.values()].filter(b => b.calendarId === calendarId).map(b => ({ ...b }));
    },
    async applyBusyDiff(_calendarId, upserts, deleteIds) {
      await tick();
      for (const block of upserts) busy.set(block.id, { ...block });
      for (const id of deleteIds) busy.delete(id);
    },
    async deleteBusyBlocks(calendarId) {
      for (const [id, block] of busy) if (block.calendarId === calendarId) busy.delete(id);
    },
  };

  return { store, authRequests, connections, calendars, subscriptions, busy };
}
```

```ts
// microsoft-calendar/testing/fakeGraph.ts
import {
  GraphApiError,
  type GraphApi,
  type GraphCalendar,
  type GraphDateTime,
  type GraphEvent,
  type GraphSubscriptionResource,
} from '../graphApi';

const MAX_SUBSCRIPTION_MS = 10_080 * 60_000;

/** An unsigned JWT with the claims the connect flow reads. */
export function fakeIdToken(claims: Record<string, string>): string {
  const part = (value: object) => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${part({ alg: 'none' })}.${part(claims)}.`;
}

/** Wall clock in `timeZone` → instant, the way Exchange reads an event's start and end. */
function toInstant(value: GraphDateTime | undefined): Date {
  const wall = Date.parse(`${value?.dateTime ?? ''}Z`);
  const zone = value?.timeZone ?? 'UTC';
  if (zone === 'UTC') return new Date(wall);
  const offset = (at: number) => {
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone: zone,
      hourCycle: 'h23',
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
    }).formatToParts(new Date(at));
    const get = (type: string) => Number(parts.find(p => p.type === type)?.value);
    return Date.UTC(get('year'), get('month') - 1, get('day'), get('hour'), get('minute'), get('second')) - at;
  };
  return new Date(wall - offset(wall - offset(wall)));
}

/** The UTC wall-clock form calendar views return under `Prefer: outlook.timezone="UTC"`. */
export function utcGraph(iso: string): GraphDateTime {
  return { dateTime: `${new Date(iso).toISOString().slice(0, 19)}.0000000`, timeZone: 'UTC' };
}

/** Microsoft Graph in memory: paging, marker search, subscriptions with Graph's 404/409 behaviour. */
export function createFakeGraph(options: { pageSize?: number; now?: () => Date } = {}) {
  const now = options.now ?? (() => new Date());
  const state = {
    tenantId: 'tenant-1',
    objectId: 'object-1',
    username: 'host@contoso.com',
    // Graph returns some scopes as full resource URIs.
    grantedScope:
      'https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/MailboxSettings.Read openid profile offline_access',
    mailboxZone: 'W. Europe Standard Time',
    calendars: [] as GraphCalendar[],
    events: new Map<string, Map<string, GraphEvent>>(),
    pageSize: options.pageSize ?? 500,
    subscriptions: new Map<string, GraphSubscriptionResource & { clientState: string; notificationUrl: string }>(),
    refreshFails: null as null | 'invalid_grant' | 'interaction_required',
    subscribeFails: false,
    failPage: null as null | number,
    calls: [] as string[],
    issued: 0,
  };

  const eventsOf = (calendarId: string) => {
    let events = state.events.get(calendarId);
    if (!events) state.events.set(calendarId, (events = new Map()));
    return events;
  };
  const tokens = () => ({
    access_token: `access-${++state.issued}`,
    refresh_token: `refresh-${state.issued}`,
    expires_in: 3600,
    scope: state.grantedScope,
    id_token: fakeIdToken({ tid: state.tenantId, oid: state.objectId, preferred_username: state.username }),
    token_type: 'Bearer',
  });
  /** Stored the way Exchange keeps it, returned the way a UTC calendar view shows it. */
  const normalise = (event: GraphEvent): GraphEvent => ({
    ...event,
    start: utcGraph(toInstant(event.start).toISOString()),
    end: utcGraph(toInstant(event.end).toISOString()),
  });

  const api: GraphApi = {
    authUrl: ({ state: s, codeChallenge }) => `https://login.example/authorize?state=${s}&cc=${codeChallenge}`,
    async exchangeCode() {
      state.calls.push('exchangeCode');
      return tokens();
    },
    async refresh() {
      state.calls.push('refresh');
      if (state.refreshFails) throw new GraphApiError(400, state.refreshFails, 'AADSTS70000: The grant is expired.');
      return tokens();
    },
    async mailboxTimeZone() {
      if (!/mailboxsettings\.read/i.test(state.grantedScope)) {
        throw new GraphApiError(403, 'ErrorAccessDenied', 'Access is denied. Check credentials and try again.');
      }
      return state.mailboxZone;
    },
    async listCalendars() {
      state.calls.push('listCalendars');
      return state.calendars.map(calendar => ({ ...calendar }));
    },
    async listCalendarView(_token, calendarId, { nextLink }) {
      const page = nextLink ? Number(new URL(nextLink).searchParams.get('page')) : 0;
      state.calls.push(`calendarView:${calendarId}:${page}`);
      if (state.failPage === page) throw new GraphApiError(503, 'ServiceUnavailable', 'Service Unavailable');
      const all = [...eventsOf(calendarId).values()];
      const items = all.slice(page * state.pageSize, (page + 1) * state.pageSize);
      const more = (page + 1) * state.pageSize < all.length;
      return more
        ? { items, nextLink: `https://graph.microsoft.com/v1.0/me/calendarView?page=${page + 1}` }
        : { items };
    },
    async findEventsByMarker(_token, calendarId, markerId, value) {
      return [...eventsOf(calendarId).values()].filter(event =>
        event.singleValueExtendedProperties?.some(p => p.id === markerId && p.value === value)
      );
    },
    async createEvent(_token, calendarId, event) {
      state.calls.push('createEvent');
      const events = eventsOf(calendarId);
      const repeat = [...events.values()].find(e => event.transactionId && e.transactionId === event.transactionId);
      if (repeat) return repeat;
      const id = `AAMkAGI2/${++state.issued}+event=`;
      const stored: GraphEvent = {
        ...normalise(event),
        id,
        ...(event.isOnlineMeeting ? { onlineMeeting: { joinUrl: `https://teams.example/l/meetup/${id}` } } : {}),
      };
      events.set(id, stored);
      return stored;
    },
    async updateEvent(_token, calendarId, eventId, event) {
      state.calls.push(`updateEvent:${eventId}`);
      const events = eventsOf(calendarId);
      const current = events.get(eventId);
      if (!current) throw new GraphApiError(404, 'ErrorItemNotFound', 'The specified object was not found in the store.');
      const stored = { ...current, ...normalise({ ...current, ...event }) };
      events.set(eventId, stored);
      return stored;
    },
    async deleteEvent(_token, calendarId, eventId) {
      state.calls.push(`deleteEvent:${eventId}`);
      if (!eventsOf(calendarId).delete(eventId)) {
        throw new GraphApiError(404, 'ErrorItemNotFound', 'The specified object was not found in the store.');
      }
    },
    async createSubscription(_token, body) {
      state.calls.push(`subscribe:${body.resource}`);
      if (state.subscribeFails) throw new GraphApiError(503, 'ServiceUnavailable', 'Service Unavailable');
      if (body.clientState.length > 128) throw new GraphApiError(400, 'InvalidRequest', 'clientState is too long');
      for (const existing of state.subscriptions.values()) {
        if (existing.resource === body.resource) {
          throw new GraphApiError(409, 'Conflict', `Subscription Id ${existing.id} already exists for the requested combination`);
        }
      }
      if (Date.parse(body.expirationDateTime) - now().getTime() > MAX_SUBSCRIPTION_MS) {
        throw new GraphApiError(400, 'InvalidRequest', 'Subscription expiration exceeds the maximum');
      }
      const id = `graph-sub-${++state.issued}`;
      const subscription = {
        id,
        resource: body.resource,
        expirationDateTime: body.expirationDateTime,
        clientState: body.clientState,
        notificationUrl: body.notificationUrl,
      };
      state.subscriptions.set(id, subscription);
      return { id, resource: subscription.resource, expirationDateTime: subscription.expirationDateTime };
    },
    async renewSubscription(_token, graphId, expiresAt) {
      state.calls.push(`renew:${graphId}`);
      const current = state.subscriptions.get(graphId);
      if (!current) throw new GraphApiError(404, 'ResourceNotFound', 'The object was not found.');
      current.expirationDateTime = expiresAt.toISOString();
      return { id: graphId, resource: current.resource, expirationDateTime: current.expirationDateTime };
    },
    async deleteSubscription(_token, graphId) {
      state.calls.push(`unsubscribe:${graphId}`);
      if (!state.subscriptions.delete(graphId)) throw new GraphApiError(404, 'ResourceNotFound', 'The object was not found.');
    },
    async listSubscriptions() {
      return [...state.subscriptions.values()].map(({ id, resource, expirationDateTime }) => ({
        id,
        resource,
        expirationDateTime,
      }));
    },
  };

  return { api, state, eventsOf };
}
```

## Checklist

- [ ] Suite green under at least two `TZ` values
- [ ] Every mutation in the table turns it red
- [ ] A host-specific store passes the flow tests when swapped in for the memory store
