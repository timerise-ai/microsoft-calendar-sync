# Calendars and settings

The calendar list mirrors what the account can see in Outlook; the two toggles on each row decide
what the integration does with it.

## Calendar list

`refreshCalendarList` runs whenever the user opens the settings screen:

- **Every page.** `/me/calendars` is paged; the client follows `@odata.nextLink`. Users with shared
  team calendars, other people's calendars and *Birthdays* pass the first page quickly.
- **New calendars** get rows with both syncs off.
- **Existing rows** keep the user's settings and title, and take Outlook's current `canEdit`.
- **`canEdit` decides writability**, and covers calendars other people shared with edit rights. A
  connection with `Calendars.Read` only makes every calendar read-only. Losing write access turns
  booking sync off.
- **The mailbox zone is refreshed** at the same time ([time-zones.md](time-zones.md)).
- **Calendars Outlook no longer lists are torn down**: subscription deleted, blocks deleted, row
  removed. Otherwise a calendar the user removed keeps blocking their time.
- **Errors propagate.** A revoked grant must reach the UI as `reauth-required`, not as an empty list
  that looks like "you have no calendars".

## Settings

`updateCalendarSettings` writes the row (one compare-and-swap, so concurrent toggles serialize), then
runs what the change implies:

| Before → after | Transition | Effect |
|---|---|---|
| `syncBusy` off → on | `start-busy` | open subscription, first sync |
| `syncBusy` on → off | `stop-busy` | delete subscription, delete blocks |
| `syncBookings` off → on | `backfill-bookings` | write upcoming bookings of this host |
| `syncBookings` on → off | — | existing events stay |
| `title` | — | display only |

Guards: the calendar must belong to the caller (`calendar-not-found` otherwise, never "forbidden", so
ids cannot be probed), and booking sync needs `isWritable` (`calendar-not-writable`). A transition that
fails leaves the setting on and records the error; the cron completes it. Reverting the toggle instead
would turn a transient Graph error into a setting the user has to find and fix.

## Code

```ts
// microsoft-calendar/calendars.ts
import { backfillBookings } from './bookingEvents';
import { calendarRowId } from './crypto';
import type { GraphCalendar } from './graphApi';
import { disableBusySync, enableBusySync, teardownCalendar } from './subscriptions';
import { toIanaTimeZone } from './timeZones';
import { withAccessToken } from './tokens';
import {
  CalendarSyncError,
  type CalendarSettingsPatch,
  type CalendarSyncDeps,
  type LinkedCalendar,
} from './types';

/**
 * The row for one Outlook calendar: new rows start with both syncs off; existing rows keep the
 * user's settings and title, take the current edit permission, and lose booking sync if no longer writable.
 */
export function mergeCalendarRow(
  userId: string,
  entry: GraphCalendar,
  current: LinkedCalendar | null,
  connectionCanWrite: boolean
): LinkedCalendar {
  // `canEdit` covers the user's own calendars and calendars shared with them with edit rights.
  const isWritable = connectionCanWrite && entry.canEdit === true;
  const isDefault = entry.isDefaultCalendar === true;
  if (current) {
    return { ...current, isDefault, isWritable, syncBookings: current.syncBookings && isWritable };
  }
  return {
    id: calendarRowId(userId, entry.id),
    userId,
    externalId: entry.id,
    title: entry.name ?? 'Calendar',
    isDefault,
    isWritable,
    syncBusy: false,
    syncBookings: false,
    subscriptionId: null,
    lastSyncedAt: null,
    lastSyncError: null,
    syncLeaseUntil: null,
    resyncRequested: false,
    version: 0,
  };
}

function sameRow(a: LinkedCalendar, b: LinkedCalendar): boolean {
  return a.isDefault === b.isDefault && a.isWritable === b.isWritable && a.syncBookings === b.syncBookings;
}

/**
 * Read the mailbox's zone again (the user may have moved). Best effort: without
 * `MailboxSettings.Read` it stays as it was and all-day events use the host's zone.
 */
export async function refreshMailboxTimeZone(deps: CalendarSyncDeps, userId: string, accessToken: string): Promise<void> {
  try {
    const zone = toIanaTimeZone(await deps.graph.mailboxTimeZone(accessToken));
    const connection = await deps.store.getConnection(userId);
    if (zone && connection && connection.mailboxTimeZone !== zone) {
      await deps.store.patchConnection(userId, { mailboxTimeZone: zone, updatedAt: deps.now() });
    }
  } catch (error) {
    deps.log.info('microsoft-calendar-mailbox-zone-unavailable', {
      userId,
      message: error instanceof Error ? error.message : String(error),
    });
  }
}

/**
 * Bring the stored rows in line with the account's calendar list. Errors propagate: a revoked
 * connection must surface as `reauth-required`, never as an empty list.
 */
export async function refreshCalendarList(deps: CalendarSyncDeps, userId: string): Promise<LinkedCalendar[]> {
  const connection = await deps.store.getConnection(userId);
  if (!connection) throw new CalendarSyncError('not-connected');
  const entries = await withAccessToken(deps, userId, async accessToken => {
    const calendars = await deps.graph.listCalendars(accessToken);
    await refreshMailboxTimeZone(deps, userId, accessToken);
    return calendars;
  });

  const rows: LinkedCalendar[] = [];
  for (const entry of entries) {
    const { after } = await deps.store.mutateCalendar(calendarRowId(userId, entry.id), current => {
      const next = mergeCalendarRow(userId, entry, current, connection.canWrite);
      return current && sameRow(current, next) ? null : next;
    });
    if (after) rows.push(after);
  }

  const listed = new Set(rows.map(row => row.id));
  for (const stale of await deps.store.listCalendars(userId)) {
    if (!listed.has(stale.id)) await teardownCalendar(deps, stale);
  }
  return rows;
}

export type Transition = 'start-busy' | 'stop-busy' | 'backfill-bookings';

/** What a settings change sets in motion. Turning booking sync off leaves existing events alone. */
export function planTransitions(before: LinkedCalendar, after: LinkedCalendar): Transition[] {
  const transitions: Transition[] = [];
  if (!before.syncBusy && after.syncBusy) transitions.push('start-busy');
  if (before.syncBusy && !after.syncBusy) transitions.push('stop-busy');
  if (!before.syncBookings && after.syncBookings) transitions.push('backfill-bookings');
  return transitions;
}

export function applySettings(
  userId: string,
  current: LinkedCalendar | null,
  patch: CalendarSettingsPatch
): LinkedCalendar {
  if (!current || current.userId !== userId) throw new CalendarSyncError('calendar-not-found');
  if (patch.syncBookings === true && !current.isWritable) {
    throw new CalendarSyncError('calendar-not-writable');
  }
  return {
    ...current,
    ...(patch.title === undefined ? {} : { title: patch.title }),
    ...(patch.syncBusy === undefined ? {} : { syncBusy: patch.syncBusy }),
    ...(patch.syncBookings === undefined ? {} : { syncBookings: patch.syncBookings }),
  };
}

/**
 * Change one calendar's settings and run what the change implies. The row is written first and the
 * side effects follow; if one fails, the error is recorded and the cron finishes the job (it opens
 * missing subscriptions and resyncs every calendar with busy sync).
 */
export async function updateCalendarSettings(
  deps: CalendarSyncDeps,
  userId: string,
  calendarId: string,
  patch: CalendarSettingsPatch
): Promise<LinkedCalendar> {
  const { before, after } = await deps.store.mutateCalendar(calendarId, current =>
    applySettings(userId, current, patch)
  );
  if (!before || !after) throw new CalendarSyncError('calendar-not-found');

  for (const transition of planTransitions(before, after)) {
    try {
      if (transition === 'start-busy') await enableBusySync(deps, after);
      if (transition === 'stop-busy') await disableBusySync(deps, after);
      if (transition === 'backfill-bookings') await backfillBookings(deps, after);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      deps.log.error('microsoft-calendar-transition-failed', { calendarId, transition, message });
      await deps.store.mutateCalendar(calendarId, row => (row ? { ...row, lastSyncError: message } : null));
      if (error instanceof CalendarSyncError) throw error;
    }
  }
  return (await deps.store.getCalendar(calendarId)) ?? after;
}
```

## Checklist

- [ ] Calendar list refresh follows every page and tears down calendars Outlook dropped
- [ ] Settings writes go through `mutateCalendar`, with the ownership check inside `fn`
- [ ] The UI offers booking sync only where `isWritable`
