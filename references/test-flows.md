# Test suite: flows

The connect, sync, subscription, webhook, outbound and route flows, 21 tests against the memory store
and the fake Graph. The pure-function and client tests are in [test-suite.md](test-suite.md); the
doubles and how to run everything in [testing.md](testing.md).

```ts
// microsoft-calendar/flows.test.ts
import assert from 'node:assert/strict';
import { test } from 'node:test';

import { pushBookingEvent, removeBookingEvent } from './bookingEvents';
import { bookingMarkerId } from './busyRules';
import { syncBusy } from './busySync';
import { refreshCalendarList, updateCalendarSettings } from './calendars';
import { completeConnect, disconnect, startConnect } from './connect';
import { signClientState } from './crypto';
import { createFakeGraph, utcGraph } from './testing/fakeGraph';
import { createMemoryStore } from './testing/memoryStore';
import { createCalendarRoutes } from './routes';
import { renewSubscriptions, subscriptionResource } from './subscriptions';
import { openTokens, withAccessToken } from './tokens';
import { handleNotification } from './webhook';
import { DEFAULT_SYNC_SETTINGS, type BookingForCalendar, type CalendarSyncDeps } from './types';

const USER = 'user-1';
const PRIMARY = 'AAMkAGI2/primary+calendar=';
const SHARED = 'AAMkAGI2/shared+calendar=';
const MARKER_SET = '0f4e8a52-3c1d-4b7a-9e26-5d8c1b2a7f90';

function setup(options: { pageSize?: number } = {}) {
  const memory = createMemoryStore();
  const clock = { now: new Date('2026-10-16T08:00:00Z') };
  const fake = createFakeGraph({ ...options, now: () => clock.now });
  fake.state.calendars = [
    { id: PRIMARY, name: 'Calendar', canEdit: true, isDefaultCalendar: true },
    { id: SHARED, name: 'Team (shared)', canEdit: false },
  ];
  const upcoming: BookingForCalendar[] = [];
  const links: [string, string][] = [];
  const logs: string[] = [];
  const deps: CalendarSyncDeps = {
    store: memory.store,
    graph: fake.api,
    config: {
      ...DEFAULT_SYNC_SETTINGS,
      clientId: 'id',
      clientSecret: 'secret',
      redirectUri: 'https://api.example/calendar/microsoft/callback',
      notificationUrl: 'https://api.example/calendar/microsoft/notifications',
      clientStateSecret: 'client-state-secret',
      encryptionKey: 'b'.repeat(64),
      markerPropertySetId: MARKER_SET,
      defaultTimeZone: 'UTC',
    },
    bookings: { listUpcomingForPrimaryHost: async () => upcoming },
    log: { info: kind => logs.push(kind), warn: kind => logs.push(kind), error: kind => logs.push(kind) },
    now: () => clock.now,
    onConferenceLink: async (bookingId, url) => void links.push([bookingId, url]),
  };
  return { deps, memory, fake, clock, upcoming, links, logs };
}

type Env = ReturnType<typeof setup>;

async function connect(env: Env) {
  const url = await startConnect(env.deps, USER);
  const state = new URL(url).searchParams.get('state') ?? '';
  return completeConnect(env.deps, { state, code: 'code' });
}

async function primaryRow(env: Env) {
  const [row] = (await refreshCalendarList(env.deps, USER)).filter(c => c.externalId === PRIMARY);
  assert.ok(row);
  return row;
}

async function connectWithBusySync(env: Env) {
  await connect(env);
  return updateCalendarSettings(env.deps, USER, (await primaryRow(env)).id, { syncBusy: true });
}

/** A notification as Graph would send it for the calendar's current subscription. */
function notificationFor(env: Env, calendarId: string, extra: Record<string, string> = {}) {
  const subscriptionId = env.memory.calendars.get(calendarId)?.subscriptionId ?? '';
  const graphId = env.memory.subscriptions.get(subscriptionId)?.graphId ?? '';
  return {
    subscriptionId: graphId,
    clientState: signClientState(subscriptionId, 'client-state-secret'),
    changeType: 'updated',
    ...extra,
  };
}

const event = (id: string, startIso: string, endIso: string) => ({ id, start: utcGraph(startIso), end: utcGraph(endIso) });

test('connect: single-use state, user from the server, encrypted tokens, account and mailbox zone', async () => {
  const env = setup();
  const url = await startConnect(env.deps, USER);
  const state = new URL(url).searchParams.get('state') ?? '';
  assert.equal((await completeConnect(env.deps, { state, code: 'code' })).userId, USER);
  await assert.rejects(completeConnect(env.deps, { state, code: 'code' }), /invalid-or-expired-state/);
  await assert.rejects(completeConnect(env.deps, { state: '../users/x', code: 'code' }), /invalid-or-expired-state/);

  const stored = env.memory.connections.get(USER);
  assert.ok(stored && !stored.encryptedTokens.includes('refresh-'));
  assert.equal(stored.accountId, 'tenant-1:object-1');
  assert.equal(stored.mailboxTimeZone, 'Europe/Berlin');
  assert.equal(stored.canWrite, true);
});

test('connect: an expired state is refused; Calendars.Read only makes nothing writable; no calendar scope is refused', async () => {
  const env = setup();
  const url = await startConnect(env.deps, USER);
  env.clock.now = new Date(env.clock.now.getTime() + 11 * 60_000);
  const state = new URL(url).searchParams.get('state') ?? '';
  await assert.rejects(completeConnect(env.deps, { state, code: 'code' }), /invalid-or-expired-state/);

  env.fake.state.grantedScope = 'Calendars.Read openid offline_access';
  await connect(env);
  assert.ok((await refreshCalendarList(env.deps, USER)).every(c => !c.isWritable));
  assert.equal(env.memory.connections.get(USER)?.mailboxTimeZone, null);

  env.fake.state.grantedScope = 'openid offline_access';
  await assert.rejects(connect(env), /missing-calendar-scope/);
});

test('busy sync imports every page; a failed page deletes nothing', async () => {
  const env = setup({ pageSize: 2 });
  const events = env.fake.eventsOf(PRIMARY);
  for (let i = 0; i < 5; i++) events.set(`ev${i}`, event(`ev${i}`, `2026-10-2${i}T09:00:00Z`, `2026-10-2${i}T10:00:00Z`));
  const calendar = await connectWithBusySync(env);
  assert.equal(env.memory.busy.size, 5);

  env.fake.state.failPage = 1;
  await assert.rejects(syncBusy(env.deps, calendar.id));
  assert.equal(env.memory.busy.size, 5);
  assert.match(env.memory.calendars.get(calendar.id)?.lastSyncError ?? '', /Service Unavailable/);
});

test('all-day events: mailbox zone first, the host zone without MailboxSettings.Read', async () => {
  const vacation = {
    id: 'vacation',
    isAllDay: true,
    showAs: 'oof' as const,
    start: { dateTime: '2026-11-02T00:00:00.0000000', timeZone: 'UTC' },
    end: { dateTime: '2026-11-04T00:00:00.0000000', timeZone: 'UTC' },
  };
  const env = setup();
  env.fake.eventsOf(PRIMARY).set('vacation', vacation);
  await connectWithBusySync(env);
  assert.equal([...env.memory.busy.values()][0]?.from.toISOString(), '2026-11-01T23:00:00.000Z'); // Europe/Berlin

  const noMailbox = setup();
  noMailbox.fake.state.grantedScope = 'Calendars.ReadWrite openid offline_access';
  noMailbox.deps.hostTimeZone = async () => 'America/New_York';
  noMailbox.fake.eventsOf(PRIMARY).set('vacation', vacation);
  await connectWithBusySync(noMailbox);
  assert.equal([...noMailbox.memory.busy.values()][0]?.from.toISOString(), '2026-11-02T05:00:00.000Z');
});

test('webhook: validation echo, answers before syncing, and a burst converges', async () => {
  const env = setup();
  const events = env.fake.eventsOf(PRIMARY);
  events.set('a', event('a', '2026-10-20T09:00:00Z', '2026-10-20T10:00:00Z'));
  const calendar = await connectWithBusySync(env);

  assert.deepEqual(await handleNotification(env.deps, { validationToken: 'Validation: Token', body: null }), {
    status: 200,
    text: 'Validation: Token',
  });

  events.set('b', event('b', '2026-10-21T09:00:00Z', '2026-10-21T10:00:00Z'));
  const viewsBefore = env.fake.state.calls.filter(c => c.startsWith('calendarView')).length;
  const notification = notificationFor(env, calendar.id);
  const results = await Promise.all(
    [1, 2, 3].map(() => handleNotification(env.deps, { validationToken: null, body: { value: [notification, notification] } }))
  );
  assert.deepEqual(results.map(r => r.status), [202, 202, 202]);
  // Nothing was synced while answering: Graph gets its 202 within the 3 seconds.
  assert.equal(env.fake.state.calls.filter(c => c.startsWith('calendarView')).length, viewsBefore);

  const running = Promise.all(results.map(r => r.work?.()));
  events.set('c', event('c', '2026-10-22T09:00:00Z', '2026-10-22T10:00:00Z'));
  await running;
  assert.deepEqual([...env.memory.busy.values()].map(b => b.eventId).sort(), ['a', 'b', 'c']);
  assert.equal(env.memory.calendars.get(calendar.id)?.syncLeaseUntil, null);
});

test('webhook: forged, foreign, unknown and malformed notifications schedule nothing', async () => {
  const env = setup();
  const calendar = await connectWithBusySync(env);
  const good = notificationFor(env, calendar.id);
  const forged = { ...good, clientState: signClientState(env.memory.calendars.get(calendar.id)?.subscriptionId ?? '', 'guess') };
  const foreign = { ...good, subscriptionId: 'graph-sub-of-someone-else' };
  const unknown = { ...good, clientState: signClientState('f00dfeed-0000-4000-8000-000000000000', 'client-state-secret') };
  const result = await handleNotification(env.deps, { validationToken: null, body: { value: [forged, foreign, unknown] } });
  assert.deepEqual(result, { status: 202 });
  assert.equal((await handleNotification(env.deps, { validationToken: null, body: { nope: 1 } })).status, 400);
  // Junk beside a good entry is skipped; the good one is still scheduled.
  const mixed = await handleNotification(env.deps, { validationToken: null, body: { value: [null, { clientState: 7 }, good] } });
  assert.equal(mixed.status, 202);
  assert.equal(typeof mixed.work, 'function');
});

test('busy time is imported even when Graph will not open a subscription', async () => {
  const env = setup();
  env.fake.eventsOf(PRIMARY).set('a', event('a', '2026-10-20T09:00:00Z', '2026-10-20T10:00:00Z'));
  env.fake.state.subscribeFails = true;
  await connect(env);
  const row = await primaryRow(env);
  await updateCalendarSettings(env.deps, USER, row.id, { syncBusy: true });
  assert.equal(env.memory.busy.size, 1);
  assert.match(env.memory.calendars.get(row.id)?.lastSyncError ?? '', /Service Unavailable/);

  env.fake.eventsOf(PRIMARY).set('b', event('b', '2026-10-21T09:00:00Z', '2026-10-21T10:00:00Z'));
  const report = await renewSubscriptions(env.deps);
  assert.equal(report.failed, 1);
  assert.equal(report.synced, 1);
  assert.equal(env.memory.busy.size, 2);
});

test('renewal: in place when expiring, replaced when Graph lost it, kept with an error when Graph fails', async () => {
  const env = setup();
  const calendar = await connectWithBusySync(env);
  const first = env.memory.calendars.get(calendar.id)?.subscriptionId ?? '';
  const firstGraphId = env.memory.subscriptions.get(first)?.graphId ?? '';

  env.clock.now = new Date('2026-10-21T08:00:00Z'); // expires 2026-10-22, inside the 48 h window
  assert.equal((await renewSubscriptions(env.deps)).renewed, 1);
  assert.equal(env.memory.calendars.get(calendar.id)?.subscriptionId, first);
  assert.ok(env.fake.state.calls.includes(`renew:${firstGraphId}`));

  env.fake.state.subscriptions.delete(firstGraphId); // expired or removed at Graph
  env.clock.now = new Date('2026-10-26T08:00:00Z');
  env.fake.state.subscribeFails = true;
  const failed = await renewSubscriptions(env.deps);
  assert.equal(failed.failed, 1);
  assert.match(env.memory.calendars.get(calendar.id)?.lastSyncError ?? '', /Service Unavailable/);

  env.fake.state.subscribeFails = false;
  assert.equal((await renewSubscriptions(env.deps)).opened, 1);
  const second = env.memory.calendars.get(calendar.id)?.subscriptionId ?? '';
  assert.notEqual(second, first);
  assert.equal(env.fake.state.subscriptions.size, 1);
});

test('renewal heals a lost record, deleting the subscription Graph still holds (409)', async () => {
  const env = setup();
  const calendar = await connectWithBusySync(env);
  const lost = env.memory.calendars.get(calendar.id)?.subscriptionId ?? '';
  env.memory.subscriptions.delete(lost);
  await renewSubscriptions(env.deps);
  const healed = env.memory.calendars.get(calendar.id)?.subscriptionId ?? '';
  assert.ok(healed && healed !== lost && env.memory.subscriptions.has(healed));
  assert.equal(env.fake.state.subscriptions.size, 1);
  assert.equal([...env.fake.state.subscriptions.values()][0]?.resource, subscriptionResource(PRIMARY));
});

test('lifecycle: subscriptionRemoved replaces the subscription and resyncs', async () => {
  const env = setup();
  const calendar = await connectWithBusySync(env);
  const removed = notificationFor(env, calendar.id, { lifecycleEvent: 'subscriptionRemoved' });
  env.fake.state.subscriptions.delete(removed.subscriptionId);
  const result = await handleNotification(env.deps, { validationToken: null, body: { value: [removed] } });
  await result.work?.();
  const current = env.memory.subscriptions.get(env.memory.calendars.get(calendar.id)?.subscriptionId ?? '');
  assert.ok(current && current.graphId !== removed.subscriptionId);
  assert.ok(env.fake.state.subscriptions.has(current.graphId));
});

test('turning busy sync off deletes the subscription and the blocks', async () => {
  const env = setup();
  env.fake.eventsOf(PRIMARY).set('a', event('a', '2026-10-20T09:00:00Z', '2026-10-20T10:00:00Z'));
  const calendar = await connectWithBusySync(env);
  await updateCalendarSettings(env.deps, USER, calendar.id, { syncBusy: false });
  assert.equal(env.memory.busy.size, 0);
  assert.equal(env.memory.subscriptions.size, 0);
  assert.equal(env.fake.state.subscriptions.size, 0);
});

const booking: BookingForCalendar = {
  bookingId: 'b-1',
  start: new Date('2026-10-20T07:00:00Z'),
  end: new Date('2026-10-20T08:00:00Z'),
  timeZone: 'Europe/Warsaw',
  title: 'Yoga class',
  guestEmail: 'guest@example.com',
  hostIds: [USER],
  hostEmails: { [USER]: 'host@contoso.com' },
  wantsTeamsMeeting: true,
};

test('outbound: pushed twice is one event and no second write; its mirror never becomes busy time', async () => {
  const env = setup();
  const calendar = await connectWithBusySync(env);
  await updateCalendarSettings(env.deps, USER, calendar.id, { syncBookings: true });

  assert.deepEqual((await pushBookingEvent(env.deps, booking)).map(r => r.ok), [true]);
  assert.deepEqual((await pushBookingEvent(env.deps, booking)).map(r => r.ok), [true]);
  assert.equal(env.fake.eventsOf(PRIMARY).size, 1);
  assert.equal(env.fake.state.calls.filter(c => c.startsWith('createEvent') || c.startsWith('updateEvent')).length, 1);
  assert.equal(env.links.length, 2);
  assert.match(env.links[0]?.[1] ?? '', /^https:\/\/teams\.example\//);

  await syncBusy(env.deps, calendar.id);
  assert.equal(env.memory.busy.size, 0);

  // Moved in place: the same event is updated, not a second one created.
  await pushBookingEvent(env.deps, { ...booking, start: new Date('2026-10-20T09:00:00Z'), end: new Date('2026-10-20T10:00:00Z') });
  const [moved] = [...env.fake.eventsOf(PRIMARY).values()];
  assert.equal(env.fake.eventsOf(PRIMARY).size, 1);
  assert.equal(moved?.start?.dateTime, '2026-10-20T09:00:00.0000000');
});

test('outbound: duplicates left by racing triggers collapse to one', async () => {
  const env = setup();
  await connect(env);
  await updateCalendarSettings(env.deps, USER, (await primaryRow(env)).id, { syncBookings: true });
  const marker = bookingMarkerId(MARKER_SET);
  for (const id of ['dup-1', 'dup-2']) {
    env.fake.eventsOf(PRIMARY).set(id, {
      ...event(id, '2026-10-20T07:00:00Z', '2026-10-20T08:00:00Z'),
      subject: 'Guest - Yoga class',
      singleValueExtendedProperties: [{ id: marker, value: 'b-1' }],
    });
  }
  await pushBookingEvent(env.deps, booking);
  assert.deepEqual([...env.fake.eventsOf(PRIMARY).keys()], ['dup-1']);
});

test('outbound: cancelling deletes, twice is harmless, re-confirming writes it again, sync off still deletes', async () => {
  const env = setup();
  await connect(env);
  const primary = await primaryRow(env);
  await updateCalendarSettings(env.deps, USER, primary.id, { syncBookings: true });
  await pushBookingEvent(env.deps, booking);
  assert.deepEqual((await removeBookingEvent(env.deps, booking)).map(r => r.ok), [true]);
  assert.deepEqual((await removeBookingEvent(env.deps, booking)).map(r => r.ok), [true]);
  assert.equal(env.fake.eventsOf(PRIMARY).size, 0);
  await pushBookingEvent(env.deps, booking);
  assert.equal(env.fake.eventsOf(PRIMARY).size, 1);
  // Booking sync turned off after the event was written: the toggle leaves it, the cancellation removes it.
  await updateCalendarSettings(env.deps, USER, primary.id, { syncBookings: false });
  assert.equal(env.fake.eventsOf(PRIMARY).size, 1);
  await removeBookingEvent(env.deps, booking);
  assert.equal(env.fake.eventsOf(PRIMARY).size, 0);
});

test('outbound: turning booking sync on backfills upcoming bookings; read-only calendars refuse it', async () => {
  const env = setup();
  env.upcoming.push(booking, { ...booking, bookingId: 'b-2' });
  await connect(env);
  const rows = await refreshCalendarList(env.deps, USER);
  const shared = rows.find(c => c.externalId === SHARED);
  assert.ok(shared);
  await assert.rejects(updateCalendarSettings(env.deps, USER, shared.id, { syncBookings: true }), /calendar-not-writable/);
  await updateCalendarSettings(env.deps, USER, (await primaryRow(env)).id, { syncBookings: true });
  assert.equal(env.fake.eventsOf(PRIMARY).size, 2);
});

test('refresh stores the rotated refresh token', async () => {
  const env = setup();
  await connect(env);
  const before = openTokens(env.deps, env.memory.connections.get(USER)?.encryptedTokens ?? '');
  env.clock.now = new Date(env.clock.now.getTime() + 2 * 3_600_000);
  await withAccessToken(env.deps, USER, async () => 'ok');
  const after = openTokens(env.deps, env.memory.connections.get(USER)?.encryptedTokens ?? '');
  assert.notEqual(after.refreshToken, before.refreshToken);
});

test('invalid_grant and interaction_required mark the connection; reconnecting keeps settings', async () => {
  for (const code of ['invalid_grant', 'interaction_required'] as const) {
    const env = setup();
    const calendar = await connectWithBusySync(env);
    env.clock.now = new Date(env.clock.now.getTime() + 2 * 3_600_000);
    env.fake.state.refreshFails = code;
    await assert.rejects(withAccessToken(env.deps, USER, async () => 'never'), /reauth-required/);
    assert.equal(env.memory.connections.get(USER)?.status, 'reauth_required');
    await assert.rejects(refreshCalendarList(env.deps, USER), /reauth-required/);

    env.fake.state.refreshFails = null;
    assert.equal((await connect(env)).sameAccount, true);
    const after = env.memory.calendars.get(calendar.id);
    assert.equal(after?.syncBusy, true);
    assert.equal(env.memory.connections.get(USER)?.status, 'connected');
    assert.equal(env.fake.state.subscriptions.size, 1);
  }
});

test('connecting a different Microsoft account tears the old one down', async () => {
  const env = setup();
  env.fake.eventsOf(PRIMARY).set('a', event('a', '2026-10-20T09:00:00Z', '2026-10-20T10:00:00Z'));
  await connectWithBusySync(env);
  env.fake.state.objectId = 'object-2';
  assert.equal((await connect(env)).sameAccount, false);
  assert.equal(env.memory.calendars.size, 0);
  assert.equal(env.memory.busy.size, 0);
  assert.equal(env.fake.state.subscriptions.size, 0);
});

test('disconnect deletes subscriptions and forgets everything', async () => {
  const env = setup();
  env.fake.eventsOf(PRIMARY).set('a', event('a', '2026-10-20T09:00:00Z', '2026-10-20T10:00:00Z'));
  await connectWithBusySync(env);
  await disconnect(env.deps, USER);
  assert.equal(env.memory.connections.size, 0);
  assert.equal(env.memory.calendars.size, 0);
  assert.equal(env.memory.busy.size, 0);
  assert.equal(env.fake.state.subscriptions.size, 0);
});

test('a calendar removed in Outlook is torn down on the next listing', async () => {
  const env = setup();
  await connectWithBusySync(env);
  env.fake.state.calendars = env.fake.state.calendars.filter(c => c.id !== PRIMARY);
  const rows = await refreshCalendarList(env.deps, USER);
  assert.deepEqual(rows.map(r => r.externalId), [SHARED]);
  assert.equal(env.fake.state.subscriptions.size, 0);
});

test('routes: validation as text/plain, syncs handed to defer, callback always redirects, cron fails closed', async () => {
  const env = setup();
  const calendar = await connectWithBusySync(env);
  const deferred: (() => Promise<void>)[] = [];
  const routes = createCalendarRoutes(env.deps, {
    authenticate: async () => USER,
    appReturnUrl: 'https://app.example/settings',
    cronSecret: '',
    defer: task => void deferred.push(task),
  });
  const base = 'https://api.example/calendar/microsoft';

  const validation = await routes.notifications(new Request(`${base}/notifications?validationToken=abc%20123`, { method: 'POST' }));
  assert.equal(validation.status, 200);
  assert.equal(validation.headers.get('content-type'), 'text/plain');
  assert.equal(await validation.text(), 'abc 123');

  const views = () => env.fake.state.calls.filter(c => c.startsWith('calendarView')).length;
  const before = views();
  const notified = await routes.notifications(
    new Request(`${base}/notifications`, { method: 'POST', body: JSON.stringify({ value: [notificationFor(env, calendar.id)] }) })
  );
  assert.equal(notified.status, 202);
  assert.equal(deferred.length, 1);
  assert.equal(views(), before);
  await deferred[0]?.();
  assert.equal(views(), before + 1);

  const denied = await routes.callback(new Request(`${base}/callback?error=consent_required&state=x`));
  assert.equal(denied.status, 303);
  assert.equal(denied.headers.get('location'), 'https://app.example/settings?calendar=admin-consent-required');
  const replayed = await routes.callback(new Request(`${base}/callback?code=c&state=${'a'.repeat(43)}`));
  assert.equal(replayed.headers.get('location'), 'https://app.example/settings?calendar=invalid-or-expired-state');

  assert.equal((await routes.cron(new Request(`${base}/cron`, { method: 'POST', headers: { authorization: 'Bearer ' } }))).status, 401);
});
```
