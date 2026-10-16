# Busy sync (inbound)

A host's Outlook events become busy blocks the availability engine subtracts. The blocks are a mirror
of one window, `[now, now + horizonDays)`, rebuilt by listing every event in it and diffing against
what is stored. Nothing edits a block in place, so nothing can drift.

## Which events block

Evaluated in this order by `toBusyBlock`; the first match decides.

| Event | Blocks? | Why |
|---|---|---|
| no `id` | no (`no-id`) | Cannot be diffed |
| `isCancelled` | no | A cancelled meeting stays on attendees' calendars until they remove it |
| booking marker present | no (`own-booking`) | Written by this integration; the booking store already counts it |
| `showAs: 'free'` | no (`free`) | Outlook's own free/busy ignores it |
| `showAs: 'workingElsewhere'` | no (`working-elsewhere`) | A location, not a commitment |
| `responseStatus.response: 'declined'` | no | The user is not going |
| `showAs: 'tentative'`, `'oof'`, `'busy'`, `'unknown'` | yes | Tentative is a likely yes; out of office is the clearest busy there is |
| `isAllDay` | yes, its dates at the mailbox zone's midnights | See [time-zones.md](time-zones.md) |
| timed | yes, the UTC instants | Requested in UTC, refused in any other zone |
| unparseable or `end <= start` | no (`invalid-time`) | Never write a block the engine cannot reason about |

**All-day events are where Outlook and Google differ most.** Google defaults all-day events to *Free*;
Outlook defaults them to *Free* too, except *Out of office* ones, which is how vacations are entered.
Those must block every day they cover.

**The own-booking skip is what keeps group sessions bookable.** The engine already subtracts each
booking's seats. If the event written for the first booking of a ten-seat class came back as a busy
block, the host would be busy for the whole slot and the other nine seats would vanish.

**Limit: co-hosts' copies carry no marker.** Extended properties stay on the organizer's item; an
attendee's copy of the meeting does not have them. With `inviteAttendees` set to `all` or `hosts`, a
co-host who also syncs busy time from that calendar gets the booking back as a busy block. If the
engine treats that as a conflict for a multi-host, multi-seat service, set `inviteAttendees: 'none'`
or have the engine ignore a host's busy blocks that exactly cover a booking the host is on.

## The sync run

```
syncBusy(calendarId)
  lease: row.syncBusy? lease free? -- no lease --> set resyncRequested, return 'queued'
     | acquired
     v
  calendarView every page of [now, now + horizon) -- any page fails --> record lastSyncError, release, throw (nothing deleted)
     v
  zone for all-day events -> toBusyBlock each -> diffBusy(stored, desired) -> applyBusyDiff -> lastSyncedAt
     v
  release: resyncRequested? -- yes (max 3 runs) --> run again
```

- **Single flight per calendar.** Graph often sends several notifications for one edit (created, then
  updated). Without the lease each would list the calendar in parallel and the slowest, with the
  oldest data, would write last. With it, one run holds the lease and the rest set `resyncRequested`;
  the holder runs once more, so a change that arrived mid-run is still picked up.
- **Deterministic block ids make the diff idempotent.** Even if two runs overlapped (an expired lease,
  two deployments), both write the same ids and the result converges.
- **The diff iterates stored rows, not a map keyed by event id.** Rows with ids from an older scheme,
  or duplicates of one event, are not in the desired set and are deleted.
- **Writes only on change.** A subject edit in Outlook does not rewrite the block; a moved event does.
- **The horizon only moves when a sync runs.** A calendar nobody edits gets no notifications, so the
  cron resyncs every busy calendar ([subscriptions.md](subscriptions.md)). Set `horizonDays` at least as
  far out as bookings can be made.

## Feeding availability

Busy blocks belong to a person. The availability engine asks for a host's blocks overlapping a range
and treats them as host busy time (`kind: 'BLOCK'` in `time-slot-generator`):

```sql
select starts_at, ends_at
from microsoft_calendar_busy_blocks
where user_id = $1 and ends_at > $2 and starts_at < $3;
```

Firestore: `listHostBusy` in [firestore-backend.md](firestore-backend.md). Blocks are UTC instants;
never convert them with a business or calendar zone on the way in.

## Code

```ts
// microsoft-calendar/busyRules.ts
import { busyBlockId } from './crypto';
import type { GraphDateTime, GraphEvent } from './graphApi';
import type { BusyBlock } from './types';

/**
 * The single-value extended property that marks an event this integration wrote for a booking.
 * The GUID names your property set: generate one for your app, keep it forever, and never copy one
 * from documentation, where other apps copy it too.
 */
export function bookingMarkerId(propertySetId: string): string {
  return `String {${propertySetId}} Name bookingId`;
}

export type SkipReason =
  | 'no-id'
  | 'cancelled'
  | 'own-booking'
  | 'free'
  | 'working-elsewhere'
  | 'declined'
  | 'invalid-time';

export type BusyDecision = { block: BusyBlock } | { skip: SkipReason };

export interface BusyContext {
  userId: string;
  calendarId: string;
  /** `bookingMarkerId(config.markerPropertySetId)`. */
  markerId: string;
  /** IANA zone all-day events are anchored to: the mailbox's, else the host's. */
  allDayTimeZone: string;
}

const UTC_NAMES = new Set(['UTC', 'Etc/UTC', 'Coordinated Universal Time']);
const WALL_CLOCK = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})(\.\d+)?$/;

function offsetMs(instant: Date, timeZone: string): number {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(instant);
  const get = (type: Intl.DateTimeFormatPartTypes) => Number(parts.find(p => p.type === type)?.value);
  const wallAsUtc = Date.UTC(get('year'), get('month') - 1, get('day'), get('hour'), get('minute'), get('second'));
  return wallAsUtc - (instant.getTime() - instant.getUTCMilliseconds());
}

function localDateOf(instant: Date, timeZone: string): string {
  return new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(
    instant
  );
}

/**
 * The first instant of a local calendar date. Independent of the process TZ. Where midnight does
 * not exist (DST starting at 00:00, e.g. America/Santiago), it is the first instant that does.
 */
export function localMidnight(date: string, timeZone: string): Date {
  const [year = NaN, month = NaN, day = NaN] = date.split('-').map(Number);
  const guess = Date.UTC(year, month - 1, day);
  const first = guess - offsetMs(new Date(guess), timeZone);
  const second = guess - offsetMs(new Date(first), timeZone);
  return localDateOf(new Date(second), timeZone) === date ? new Date(second) : new Date(first);
}

export function addDays(date: string, days: number): string {
  const [year = NaN, month = NaN, day = NaN] = date.split('-').map(Number);
  return new Date(Date.UTC(year, month - 1, day + days)).toISOString().slice(0, 10);
}

/**
 * A Graph `dateTime` read as UTC. Calendar views are requested with `Prefer: outlook.timezone="UTC"`;
 * anything else is refused (Invalid Date) rather than read in the wrong zone. Graph sends seven
 * fractional digits, which `Date` accepts.
 */
export function utcInstant(value: GraphDateTime | undefined): Date {
  const match = WALL_CLOCK.exec(value?.dateTime ?? '');
  if (!match || !UTC_NAMES.has(value?.timeZone ?? '')) return new Date(NaN);
  return new Date(`${match[1]}T${match[2]}${match[3] ?? ''}Z`);
}

/** Decide whether one Outlook event blocks the user, and when. Pure. */
export function toBusyBlock(event: GraphEvent, ctx: BusyContext): BusyDecision {
  if (!event.id) return { skip: 'no-id' };
  // Cancelled meetings stay on attendees' calendars until they remove them.
  if (event.isCancelled) return { skip: 'cancelled' };
  // Our own booking, mirrored out. The booking store already counts it; importing it again would
  // block the remaining seats of a group booking.
  const marker = ctx.markerId.toLowerCase();
  if (event.singleValueExtendedProperties?.some(p => p.id.toLowerCase() === marker && p.value)) {
    return { skip: 'own-booking' };
  }
  if (event.showAs === 'free') return { skip: 'free' };
  if (event.showAs === 'workingElsewhere') return { skip: 'working-elsewhere' };
  if (event.responseStatus?.response === 'declined') return { skip: 'declined' };

  let from = utcInstant(event.start);
  let to = utcInstant(event.end);
  const allDay = event.isAllDay === true;
  if (allDay) {
    // All-day events float: Graph returns them as midnight to midnight in whatever zone the request
    // asked for, so in UTC they are the right dates at the wrong instants. Re-anchor the dates at
    // the mailbox's midnights; `end` is already the exclusive next day.
    const start = WALL_CLOCK.exec(event.start?.dateTime ?? '');
    const end = WALL_CLOCK.exec(event.end?.dateTime ?? '');
    if (start?.[2] === '00:00:00' && start[1]) {
      from = localMidnight(start[1], ctx.allDayTimeZone);
      to = localMidnight(end?.[2] === '00:00:00' && end[1] ? end[1] : addDays(start[1], 1), ctx.allDayTimeZone);
    }
    // Not at midnight: Graph sent real instants after all; they are used as they are.
  }
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime()) || to <= from) {
    return { skip: 'invalid-time' };
  }

  return {
    block: {
      id: busyBlockId(ctx.calendarId, event.id),
      userId: ctx.userId,
      calendarId: ctx.calendarId,
      eventId: event.id,
      from,
      to,
      allDay,
    },
  };
}

/**
 * What to write so the stored blocks equal `desired`. Every stored row not desired is deleted,
 * iterating rows rather than a map keyed by event, so duplicates left by older code are removed too.
 */
export function diffBusy(
  existing: BusyBlock[],
  desired: BusyBlock[]
): { upserts: BusyBlock[]; deletes: string[] } {
  const wanted = new Map(desired.map(block => [block.id, block]));
  const stored = new Map(existing.map(block => [block.id, block]));

  const deletes = existing.filter(block => !wanted.has(block.id)).map(block => block.id);
  const upserts = [...wanted.values()].filter(block => {
    const current = stored.get(block.id);
    return (
      !current ||
      current.from.getTime() !== block.from.getTime() ||
      current.to.getTime() !== block.to.getTime() ||
      current.allDay !== block.allDay
    );
  });
  return { upserts, deletes };
}
```

```ts
// microsoft-calendar/busySync.ts
import { bookingMarkerId, diffBusy, toBusyBlock, type SkipReason } from './busyRules';
import type { GraphEvent } from './graphApi';
import { isIanaTimeZone } from './timeZones';
import { withAccessToken } from './tokens';
import type { BusyBlock, CalendarSyncDeps, LinkedCalendar } from './types';

export type SyncOutcome =
  | { status: 'synced'; upserted: number; deleted: number; skipped: Partial<Record<SkipReason, number>> }
  | { status: 'queued' }
  | { status: 'disabled' };

/** Notifications that arrive during a sync queue one more run; this caps the reruns per lease. */
const MAX_RUNS_PER_LEASE = 3;

/**
 * Replace the calendar's busy blocks with what Outlook has now, inside [now, now + horizonDays].
 * Single-flight per calendar: a lease on the row, and a notification that arrives while it is held
 * sets `resyncRequested` so the holder runs once more instead of a second sync running alongside.
 */
export async function syncBusy(deps: CalendarSyncDeps, calendarId: string): Promise<SyncOutcome> {
  const now = deps.now();
  let acquired = false;
  const { after } = await deps.store.mutateCalendar(calendarId, current => {
    acquired = false;
    if (!current?.syncBusy) return null;
    if (current.syncLeaseUntil && current.syncLeaseUntil > now) {
      return current.resyncRequested ? null : { ...current, resyncRequested: true };
    }
    acquired = true;
    return {
      ...current,
      syncLeaseUntil: new Date(now.getTime() + deps.config.leaseMs),
      resyncRequested: false,
    };
  });
  if (!after?.syncBusy) return { status: 'disabled' };
  if (!acquired) return { status: 'queued' };

  let calendar: LinkedCalendar = after;
  let result: SyncOutcome = { status: 'disabled' };
  for (let run = 1; ; run++) {
    try {
      result = await runOnce(deps, calendar);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      await deps.store.mutateCalendar(calendarId, current =>
        current ? { ...current, syncLeaseUntil: null, lastSyncError: message } : null
      );
      deps.log.error('microsoft-calendar-sync-failed', { calendarId, userId: calendar.userId, message });
      throw error;
    }

    let rerun = false;
    const released = await deps.store.mutateCalendar(calendarId, current => {
      if (!current) return null;
      rerun = current.resyncRequested && current.syncBusy && run < MAX_RUNS_PER_LEASE;
      return {
        ...current,
        lastSyncedAt: deps.now(),
        lastSyncError: null,
        resyncRequested: false,
        syncLeaseUntil: rerun ? new Date(deps.now().getTime() + deps.config.leaseMs) : null,
      };
    });
    if (!rerun || !released.after) return result;
    calendar = released.after;
  }
}

/** The zone all-day events are anchored to: the mailbox's, else the host's, else the configured default. */
export async function allDayTimeZoneFor(deps: CalendarSyncDeps, userId: string): Promise<string> {
  const connection = await deps.store.getConnection(userId);
  if (connection?.mailboxTimeZone) return connection.mailboxTimeZone;
  const hostZone = deps.hostTimeZone ? await deps.hostTimeZone(userId) : null;
  return hostZone && isIanaTimeZone(hostZone) ? hostZone : deps.config.defaultTimeZone;
}

async function runOnce(deps: CalendarSyncDeps, calendar: LinkedCalendar): Promise<SyncOutcome> {
  const start = deps.now();
  const end = new Date(start.getTime() + deps.config.horizonDays * 86_400_000);
  const markerId = bookingMarkerId(deps.config.markerPropertySetId);

  // Every page, or nothing: a partial list would delete the blocks of every event it missed.
  const events = await withAccessToken(deps, calendar.userId, async accessToken => {
    const all: GraphEvent[] = [];
    let nextLink: string | undefined;
    do {
      const page = await deps.graph.listCalendarView(accessToken, calendar.externalId, {
        start,
        end,
        markerId,
        nextLink,
      });
      all.push(...page.items);
      nextLink = page.nextLink;
    } while (nextLink);
    return all;
  });

  const allDayTimeZone = await allDayTimeZoneFor(deps, calendar.userId);
  const desired: BusyBlock[] = [];
  const skipped: Partial<Record<SkipReason, number>> = {};
  for (const event of events) {
    const decision = toBusyBlock(event, { userId: calendar.userId, calendarId: calendar.id, markerId, allDayTimeZone });
    if ('block' in decision) desired.push(decision.block);
    else skipped[decision.skip] = (skipped[decision.skip] ?? 0) + 1;
  }

  const existing = await deps.store.listBusyBlocks(calendar.id);
  const { upserts, deletes } = diffBusy(existing, desired);
  if (upserts.length > 0 || deletes.length > 0) {
    await deps.store.applyBusyDiff(calendar.id, upserts, deletes);
  }
  deps.log.info('microsoft-calendar-synced', {
    calendarId: calendar.id,
    events: events.length,
    upserted: upserts.length,
    deleted: deletes.length,
    skipped,
  });
  return { status: 'synced', upserted: upserts.length, deleted: deletes.length, skipped };
}
```

## Checklist

- [ ] Every page read before any delete
- [ ] Events carrying the booking marker skipped
- [ ] All-day events anchored to the mailbox zone, or the host's
- [ ] `horizonDays` at least the booking window
- [ ] Busy blocks read by user id and treated as host busy time, in UTC
