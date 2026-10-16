# Test suite: rules and client

Pure functions and the HTTP client, 22 tests in two files. The flows are in
[test-flows.md](test-flows.md); how to run them, what each pins and the doubles they use:
[testing.md](testing.md).

## Rules (pure functions)

```ts
// microsoft-calendar/rules.test.ts
import assert from 'node:assert/strict';
import { test } from 'node:test';

import { attendeesFor, buildBookingEvent, defaultDescribeBooking, graphDateTime } from './bookingEvents';
import { bookingMarkerId, diffBusy, localMidnight, toBusyBlock, type BusyContext } from './busyRules';
import { applySettings, mergeCalendarRow, planTransitions } from './calendars';
import { accountFromIdToken } from './connect';
import { CLIENT_STATE_MAX_LENGTH, decrypt, encrypt, signClientState, verifyClientState } from './crypto';
import { deriveCalendarHealth } from './health';
import { fakeIdToken, utcGraph } from './testing/fakeGraph';
import { toIanaTimeZone } from './timeZones';
import { grantedScopes } from './tokens';
import type { BookingForCalendar, BusyBlock, CalendarConnection, LinkedCalendar } from './types';

const MARKER = bookingMarkerId('0f4e8a52-3c1d-4b7a-9e26-5d8c1b2a7f90');
const ctx: BusyContext = { userId: 'u1', calendarId: 'cal1', markerId: MARKER, allDayTimeZone: 'Europe/Warsaw' };
const block = (decision: ReturnType<typeof toBusyBlock>) => {
  assert.ok('block' in decision, `expected a block, got ${JSON.stringify(decision)}`);
  return decision.block;
};
const timed = (startIso: string, endIso: string) => ({ start: utcGraph(startIso), end: utcGraph(endIso) });

test('timed event: UTC wall clock with seven fractional digits becomes the right instants', () => {
  const b = block(toBusyBlock({ id: 'e1', ...timed('2026-10-20T07:00:00Z', '2026-10-20T08:30:00Z') }, ctx));
  assert.equal(b.from.toISOString(), '2026-10-20T07:00:00.000Z');
  assert.equal(b.to.toISOString(), '2026-10-20T08:30:00.000Z');
  assert.equal(b.allDay, false);
});

test('a time in any zone but UTC is refused, never read in the wrong zone', () => {
  const event = {
    id: 'e1',
    start: { dateTime: '2026-10-20T09:00:00.0000000', timeZone: 'W. Europe Standard Time' },
    end: { dateTime: '2026-10-20T10:00:00.0000000', timeZone: 'W. Europe Standard Time' },
  };
  assert.deepEqual(toBusyBlock(event, ctx), { skip: 'invalid-time' });
});

test('an all-day vacation blocks every day, at the mailbox zone midnights', () => {
  const event = {
    id: 'vacation',
    isAllDay: true,
    showAs: 'oof' as const,
    start: { dateTime: '2026-08-03T00:00:00.0000000', timeZone: 'UTC' },
    end: { dateTime: '2026-08-08T00:00:00.0000000', timeZone: 'UTC' },
  };
  const b = block(toBusyBlock(event, ctx));
  assert.equal(b.from.toISOString(), '2026-08-02T22:00:00.000Z');
  assert.equal(b.to.toISOString(), '2026-08-07T22:00:00.000Z');
  assert.equal(b.allDay, true);
});

test('an all-day event on the spring-forward day is 23 hours, midnight to midnight', () => {
  const event = {
    id: 'dst',
    isAllDay: true,
    start: { dateTime: '2026-03-29T00:00:00.0000000', timeZone: 'UTC' },
    end: { dateTime: '2026-03-30T00:00:00.0000000', timeZone: 'UTC' },
  };
  const b = block(toBusyBlock(event, ctx));
  assert.equal(b.from.toISOString(), '2026-03-28T23:00:00.000Z');
  assert.equal(b.to.toISOString(), '2026-03-29T22:00:00.000Z');
});

test('an all-day event that arrives as real instants is used as it is', () => {
  const event = { id: 'x', isAllDay: true, ...timed('2026-08-02T22:00:00Z', '2026-08-03T22:00:00Z') };
  const b = block(toBusyBlock(event, ctx));
  assert.equal(b.from.toISOString(), '2026-08-02T22:00:00.000Z');
  assert.equal(b.to.toISOString(), '2026-08-03T22:00:00.000Z');
});

test('localMidnight falls on the first instant of a day whose midnight does not exist', () => {
  assert.equal(localMidnight('2026-09-06', 'America/Santiago').toISOString(), '2026-09-06T04:00:00.000Z');
  assert.equal(localMidnight('2026-10-17', 'Pacific/Auckland').toISOString(), '2026-10-16T11:00:00.000Z');
});

test('events this integration wrote for bookings are not imported back as busy time', () => {
  const event = {
    id: 'mine',
    ...timed('2026-10-20T07:00:00Z', '2026-10-20T08:00:00Z'),
    // Graph may return the property id with different letter case.
    singleValueExtendedProperties: [{ id: MARKER.toUpperCase(), value: 'booking-1' }],
  };
  assert.deepEqual(toBusyBlock(event, ctx), { skip: 'own-booking' });
});

test('free, working elsewhere, declined and cancelled do not block; tentative and away do', () => {
  const t = timed('2026-10-20T07:00:00Z', '2026-10-20T08:00:00Z');
  assert.deepEqual(toBusyBlock({ id: 'a', ...t, showAs: 'free' }, ctx), { skip: 'free' });
  assert.deepEqual(toBusyBlock({ id: 'b', ...t, showAs: 'workingElsewhere' }, ctx), { skip: 'working-elsewhere' });
  assert.deepEqual(toBusyBlock({ id: 'c', ...t, responseStatus: { response: 'declined' } }, ctx), { skip: 'declined' });
  assert.deepEqual(toBusyBlock({ id: 'd', ...t, isCancelled: true }, ctx), { skip: 'cancelled' });
  assert.ok('block' in toBusyBlock({ id: 'e', ...t, showAs: 'tentative' }, ctx));
  assert.ok('block' in toBusyBlock({ id: 'f', ...t, showAs: 'oof' }, ctx));
  assert.deepEqual(toBusyBlock({ id: 'g', start: { dateTime: 'garbage', timeZone: 'UTC' } }, ctx), {
    skip: 'invalid-time',
  });
});

test('diffBusy deletes duplicates left by older code and writes only what changed', () => {
  const at = (h: number) => new Date(Date.UTC(2026, 9, 20, h));
  const desired: BusyBlock[] = [
    { id: 'k1', userId: 'u1', calendarId: 'cal1', eventId: 'e1', from: at(9), to: at(10), allDay: false },
    { id: 'k2', userId: 'u1', calendarId: 'cal1', eventId: 'e2', from: at(12), to: at(13), allDay: false },
  ];
  const existing: BusyBlock[] = [
    { ...desired[0]! },
    { ...desired[1]!, from: at(11) },
    { ...desired[0]!, id: 'legacy-random-id' },
  ];
  const { upserts, deletes } = diffBusy(existing, desired);
  assert.deepEqual(upserts.map(b => b.id), ['k2']);
  assert.deepEqual(deletes, ['legacy-random-id']);
});

test('tokens: encryption round-trips; clientState fits 128 characters and binds to its subscription', () => {
  const key = 'a'.repeat(64);
  assert.equal(decrypt(encrypt('{"refreshToken":"r"}', key), key), '{"refreshToken":"r"}');

  const id = '6f1c2c4e-8a3b-4d0e-9a51-1f2e3d4c5b6a';
  const clientState = signClientState(id, 'secret');
  assert.ok(clientState.length <= CLIENT_STATE_MAX_LENGTH);
  assert.equal(verifyClientState(clientState, 'secret'), id);
  assert.equal(verifyClientState(clientState, 'other-secret'), null);
  assert.equal(verifyClientState(clientState.replace('6f1c', '0000'), 'secret'), null);
  assert.equal(verifyClientState(clientState, ''), null);
  assert.throws(() => signClientState(id, ''));
});

test('time zones: Windows names map to IANA, IANA passes, junk is null', () => {
  assert.equal(toIanaTimeZone('W. Europe Standard Time'), 'Europe/Berlin');
  assert.equal(toIanaTimeZone('Pacific Standard Time'), 'America/Los_Angeles');
  assert.equal(toIanaTimeZone('Europe/Warsaw'), 'Europe/Warsaw');
  assert.equal(toIanaTimeZone('Not A Zone'), null);
  assert.equal(toIanaTimeZone(null), null);
});

test('account and scopes are read from the token response', () => {
  assert.deepEqual(accountFromIdToken(fakeIdToken({ tid: 't', oid: 'o', preferred_username: 'a@b.c' })), {
    accountId: 't:o',
    accountEmail: 'a@b.c',
  });
  assert.equal(accountFromIdToken(fakeIdToken({ oid: 'o' })), null);
  assert.equal(accountFromIdToken('not-a-jwt'), null);
  const scopes = grantedScopes('https://graph.microsoft.com/Calendars.ReadWrite openid Calendars.Read');
  assert.ok(scopes.has('calendars.readwrite') && scopes.has('calendars.read') && scopes.has('openid'));
});

const booking: BookingForCalendar = {
  bookingId: "b'1",
  start: new Date('2026-10-20T07:00:00Z'),
  end: new Date('2026-10-20T08:00:00Z'),
  timeZone: 'Europe/Warsaw',
  title: 'Haircut',
  guestName: 'Ada',
  guestEmail: 'ada@example.com',
  hostIds: ['h1', 'h2', 'h3'],
  hostEmails: { h1: 'h1@example.com', h2: 'h2@example.com' },
};

test('event times: wall clock in the booking zone, UTC next to a DST change', () => {
  assert.deepEqual(graphDateTime(new Date('2026-10-20T07:00:00Z'), 'Europe/Warsaw'), {
    dateTime: '2026-10-20T09:00:00',
    timeZone: 'Europe/Warsaw',
  });
  // 2026-10-25 02:30 happens twice in Warsaw; the instant must not depend on Exchange's guess.
  assert.deepEqual(graphDateTime(new Date('2026-10-25T00:30:00Z'), 'Europe/Warsaw'), {
    dateTime: '2026-10-25T00:30:00',
    timeZone: 'UTC',
  });
});

test('booking event: attendees per invite mode, hosts without e-mail left out, marker set', () => {
  const emails = (mode: 'all' | 'hosts' | 'none') => attendeesFor(booking, 'h1', mode).map(a => a.emailAddress.address);
  assert.deepEqual(emails('all'), ['ada@example.com', 'h2@example.com']);
  assert.deepEqual(emails('hosts'), ['h2@example.com']);
  assert.deepEqual(emails('none'), []);

  const event = buildBookingEvent(booking, 'h1', defaultDescribeBooking(booking), {
    markerId: MARKER,
    inviteAttendees: 'none',
  });
  assert.deepEqual(event.singleValueExtendedProperties, [{ id: MARKER, value: "b'1" }]);
  assert.match(event.body?.content ?? '', /Guest: Ada <ada@example.com>/);
  assert.equal(event.isOnlineMeeting, undefined);
  const teams = buildBookingEvent({ ...booking, wantsTeamsMeeting: true }, 'h1', defaultDescribeBooking(booking), {
    markerId: MARKER,
    inviteAttendees: 'all',
  });
  assert.equal(teams.onlineMeetingProvider, 'teamsForBusiness');
});

const outlook = (canEdit: boolean) => ({ id: 'AAMkAGI2=', name: 'Calendar', canEdit, isDefaultCalendar: true });

test('calendar rows: writable when Outlook says canEdit and the connection may write', () => {
  assert.equal(mergeCalendarRow('u1', outlook(true), null, true).isWritable, true);
  assert.equal(mergeCalendarRow('u1', outlook(false), null, true).isWritable, false);
  assert.equal(mergeCalendarRow('u1', outlook(true), null, false).isWritable, false);
  const current = { ...mergeCalendarRow('u1', outlook(true), null, true), title: 'Mine', syncBusy: true, syncBookings: true };
  const next = mergeCalendarRow('u1', outlook(false), current, true);
  assert.equal(next.syncBookings, false);
  assert.equal(next.syncBusy, true);
  assert.equal(next.title, 'Mine');
});

test('settings: transitions and guards', () => {
  const readOnly = mergeCalendarRow('u1', outlook(false), null, true);
  assert.throws(() => applySettings('u1', readOnly, { syncBookings: true }), /calendar-not-writable/);
  assert.throws(() => applySettings('someone-else', readOnly, { syncBusy: true }), /calendar-not-found/);
  const on = applySettings('u1', readOnly, { syncBusy: true });
  assert.deepEqual(planTransitions(readOnly, on), ['start-busy']);
  assert.deepEqual(planTransitions(on, readOnly), ['stop-busy']);
});

test('health is derived, problems first', () => {
  const now = new Date('2026-10-20T12:00:00Z');
  const connection = { status: 'connected' } as CalendarConnection;
  const calendar: LinkedCalendar = {
    ...mergeCalendarRow('u1', outlook(true), null, true),
    syncBusy: true,
    subscriptionId: 's1',
    lastSyncedAt: new Date('2026-10-20T06:00:00Z'),
  };
  const subscription = { id: 's1', graphId: 'g', userId: 'u1', calendarId: calendar.id, createdAt: now, expiresAt: new Date('2026-10-25T00:00:00Z') };
  const day = 26 * 3_600_000;
  assert.equal(deriveCalendarHealth({ connection, calendar, subscription, now, staleAfterMs: day }), 'ok');
  assert.equal(deriveCalendarHealth({ connection, calendar, subscription: null, now, staleAfterMs: day }), 'no-subscription');
  assert.equal(
    deriveCalendarHealth({ connection, calendar, subscription: { ...subscription, expiresAt: now }, now, staleAfterMs: day }),
    'subscription-expired'
  );
  assert.equal(
    deriveCalendarHealth({ connection: { ...connection, status: 'reauth_required' }, calendar, subscription, now, staleAfterMs: day }),
    'reauth-required'
  );
});
```

## Graph client (scripted fetch)

```ts
// microsoft-calendar/graphApi.test.ts
import assert from 'node:assert/strict';
import { test } from 'node:test';

import { createGraphApi, GraphApiError } from './graphApi';

type Reply = { status: number; body?: unknown; headers?: Record<string, string> };

/** A `fetch` that records each request and answers from a script. */
function scriptedFetch(replies: Reply[]) {
  const requests: { url: string; method: string; headers: Headers; body: string }[] = [];
  const fetchFn = (async (input: string | URL | Request, init?: RequestInit) => {
    requests.push({
      url: String(input),
      method: init?.method ?? 'GET',
      headers: new Headers(init?.headers),
      body: typeof init?.body === 'string' ? init.body : String(init?.body ?? ''),
    });
    const reply = replies.shift() ?? { status: 200, body: {} };
    return new Response(reply.body === undefined ? null : JSON.stringify(reply.body), {
      status: reply.status,
      headers: reply.headers,
    });
  }) as typeof fetch;
  return { fetchFn, requests };
}

const options = { clientId: 'client', clientSecret: 'secret', tenant: 'organizations', redirectUri: 'https://api.example/cb' };

test('authUrl: tenant authority, scopes, PKCE, and no forced consent screen', () => {
  const url = new URL(createGraphApi(options).authUrl({ state: 's', codeChallenge: 'c' }));
  assert.equal(url.origin + url.pathname, 'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize');
  assert.match(url.searchParams.get('scope') ?? '', /offline_access/);
  assert.match(url.searchParams.get('scope') ?? '', /Calendars\.ReadWrite/);
  assert.equal(url.searchParams.get('code_challenge_method'), 'S256');
  assert.equal(url.searchParams.get('prompt'), 'select_account');
});

test('calendar view: Prefer UTC on every page, follows nextLink, refuses one that leaves Graph', async () => {
  const next = 'https://graph.microsoft.com/v1.0/me/calendars/x/calendarView?$skiptoken=abc';
  const { fetchFn, requests } = scriptedFetch([
    { status: 200, body: { value: [{ id: 'a' }], '@odata.nextLink': next } },
    { status: 200, body: { value: [{ id: 'b' }] } },
  ]);
  const graph = createGraphApi({ ...options, fetch: fetchFn });
  const query = { start: new Date('2026-10-16T00:00:00Z'), end: new Date('2026-10-17T00:00:00Z'), markerId: "String {g} Name bookingId" };
  const first = await graph.listCalendarView('token', 'AAMk/x+y=', query);
  const second = await graph.listCalendarView('token', 'AAMk/x+y=', { ...query, nextLink: first.nextLink });
  assert.deepEqual([...first.items, ...second.items].map(e => e.id), ['a', 'b']);
  assert.ok(requests[0]?.url.includes('/me/calendars/AAMk%2Fx%2By%3D/calendarView?'));
  assert.equal(requests[1]?.url, next);
  assert.ok(requests.every(r => r.headers.get('prefer') === 'outlook.timezone="UTC"'));
  await assert.rejects(
    graph.listCalendarView('token', 'x', { ...query, nextLink: 'https://evil.example/steal' }),
    /Refusing to call https:\/\/evil\.example/
  );
});

test('throttling: waits Retry-After and retries; a wait too long fails instead', async () => {
  const waits: number[] = [];
  const sleep = async (ms: number) => void waits.push(ms);
  const ok = scriptedFetch([
    { status: 429, headers: { 'Retry-After': '3' }, body: { error: { code: 'TooManyRequests', message: 'slow down' } } },
    { status: 200, body: { value: 'Europe/Warsaw' } },
  ]);
  assert.equal(await createGraphApi({ ...options, fetch: ok.fetchFn, sleep }).mailboxTimeZone('t'), 'Europe/Warsaw');
  assert.deepEqual(waits, [3000]);

  const tooLong = scriptedFetch([
    { status: 429, headers: { 'Retry-After': '120' }, body: { error: { code: 'TooManyRequests', message: 'slow down' } } },
  ]);
  await assert.rejects(createGraphApi({ ...options, fetch: tooLong.fetchFn, sleep }).mailboxTimeZone('t'), (error: unknown) =>
    error instanceof GraphApiError && error.status === 429
  );
});

test('errors: token endpoint and Graph shapes both become GraphApiError codes', async () => {
  const { fetchFn } = scriptedFetch([
    { status: 400, body: { error: 'invalid_grant', error_description: 'AADSTS700082: expired' } },
    { status: 404, body: { error: { code: 'ErrorItemNotFound', message: 'not found' } } },
  ]);
  const graph = createGraphApi({ ...options, fetch: fetchFn });
  await assert.rejects(graph.refresh('r'), (e: unknown) => e instanceof GraphApiError && e.code === 'invalid_grant');
  await assert.rejects(graph.deleteEvent('t', 'c', 'e'), (e: unknown) => e instanceof GraphApiError && e.code === 'ErrorItemNotFound');
});

test('marker search escapes quotes and re-checks the value Exchange matched', async () => {
  const marker = 'String {g} Name bookingId';
  const { fetchFn, requests } = scriptedFetch([
    {
      status: 200,
      body: {
        value: [
          { id: 'mine', singleValueExtendedProperties: [{ id: marker, value: "o'neil" }] },
          { id: 'other', singleValueExtendedProperties: [{ id: marker, value: 'someone else' }] },
        ],
      },
    },
  ]);
  const found = await createGraphApi({ ...options, fetch: fetchFn }).findEventsByMarker('t', 'c', marker, "o'neil");
  assert.deepEqual(found.map(e => e.id), ['mine']);
  assert.match(new URL(requests[0]?.url ?? '').searchParams.get('$filter') ?? '', /ep\/value eq 'o''neil'/);
});
```
