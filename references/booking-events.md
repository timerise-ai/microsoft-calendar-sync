# Booking events (outbound)

A confirmed booking becomes an event in its primary host's booking-sync calendars; a cancelled one is
deleted. Graph assigns event ids, so the event carries the booking id in a private extended property,
and every write starts by finding the event through it.

## Hooks into the host's booking flow

| Booking change | Call | Notes |
|---|---|---|
| Confirmed (or paid, whichever the host calls "booked") | `pushBookingEvent(deps, booking)` | Primary host's calendars with `syncBookings` and `isWritable` |
| Cancelled, expired, refunded | `removeBookingEvent(deps, booking)` | Every host's `syncBookings` calendars |
| Rescheduled into a new booking | `removeBookingEvent(old)` then `pushBookingEvent(new)` | A new booking id is a new event |
| Time or title edited in place | `pushBookingEvent(deps, booking)` again | Found by the marker, PATCHed |
| Host turns booking sync on | `backfillBookings` (automatic, via settings) | `BookingSource.listUpcomingForPrimaryHost` |

Neither call throws. An outage, a revoked grant or a deleted calendar is logged and reported in the
returned results, and the booking flow carries on. Run them after the booking's own transaction commits,
never inside it: a retried transaction would call Graph twice, and a slow call holds it open.

## Writing one event

```
findEventsByMarker(bookingId)
  ├─ several ─▶ delete all but the first (two triggers raced and both created one)
  ├─ one ─────▶ subject and times unchanged? ── yes ─▶ done, nothing sent
  │                                           └ no ──▶ PATCH
  └─ none ────▶ POST with a fresh transactionId
```

- **Find before create.** There is no client-chosen id and no 409 to fall back on, so the marker search
  is what makes a repeated trigger, a retry or a backfill write one event.
- **`transactionId`** is Graph's guard against a POST retried after a lost response: the retry returns
  the event the first attempt created. It is fresh per call, and the client's own throttling retries
  reuse the same request body, so they carry the same value.
- **Skip unchanged writes.** Every PATCH to a meeting with attendees sends them an update. A push that
  would change nothing (a webhook retried, a status flow run twice) must send nothing. The comparison is
  on subject and times; edits to notes alone are not propagated.
- **Racing triggers can still create two events** between the search and the POST. The next push
  deletes all but one. A host that needs a hard guarantee serialises pushes per booking.

## The event

| Field | Value | Why |
|---|---|---|
| `subject` / `body` | `describeBooking(booking)`, plain text | The host's wording and language |
| `start` / `end` | wall clock in `booking.timeZone`, or UTC next to a DST change | Displays in the business's zone; never ambiguous |
| `attendees` | per `inviteAttendees`; hosts without an address left out | Graph rejects the whole event for one missing address |
| `singleValueExtendedProperties` | `String {markerGuid} Name bookingId` = booking id | Marks it as ours; how it is found again |
| `isOnlineMeeting` / `onlineMeetingProvider` | `true` / `teamsForBusiness` with `wantsTeamsMeeting` | Work and school accounts only |
| `transactionId` | random UUID, create only | Retried POSTs do not duplicate |

- **Who gets e-mailed.** Graph sends invitations to everyone in `attendees`, from the host's mailbox,
  and a cancellation when the organizer deletes the event. There is no "send no updates" switch.

  | `inviteAttendees` | Guest | Co-hosts | Use when |
  |---|---|---|---|
  | `all` (default) | invited | invited | Guests should have the meeting in their own calendar |
  | `hosts` | listed in the body | invited | The app's own confirmation is enough for guests |
  | `none` | listed in the body | not invited | Silent events; co-hosts sync from their own calendars |

- **Why the time can switch to UTC.** Graph takes a wall-clock time and a zone. In the hour that repeats
  when clocks go back, `02:30 Europe/Warsaw` names two instants, and Exchange picks one. Within three
  hours of a DST change the event is sent as UTC instead: exact, displayed in each viewer's zone anyway.
- **Teams links.** Graph creates the meeting with the event and usually returns `onlineMeeting.joinUrl`
  in the response; `onConferenceLink(bookingId, url)` lets the host store it on the booking for its own
  e-mails. If it is missing, the next push returns it.
- **Only the primary host's calendars get the event.** Other hosts are attendees on it. Deletion still
  walks every host, which covers a primary host that changed after the event was written.
- **Turning booking sync off leaves existing events in place.** Deleting a calendar's worth of events
  on a toggle is not reversible.

## Code

```ts
// microsoft-calendar/bookingEvents.ts
import { randomUUID } from 'node:crypto';

import { bookingMarkerId, utcInstant } from './busyRules';
import { GraphApiError, type GraphAttendee, type GraphDateTime, type GraphEvent } from './graphApi';
import { withAccessToken } from './tokens';
import type { BookingForCalendar, CalendarSyncDeps, InviteAttendees, LinkedCalendar } from './types';

export function defaultDescribeBooking(booking: BookingForCalendar): { summary: string; description: string } {
  const guest = [booking.guestName, booking.guestEmail ? `<${booking.guestEmail}>` : ''].filter(Boolean).join(' ');
  const lines = [
    ...(guest ? [`Guest: ${guest}`] : []),
    ...(booking.links ?? []).map(link => `${link.label}: ${link.url}`),
    ...(booking.detailsUrl ? [`Booking details: ${booking.detailsUrl}`] : []),
    ...(booking.notes ? ['', booking.notes] : []),
  ];
  return {
    summary: `${booking.guestName || 'Guest'} - ${booking.title}`,
    description: lines.join('\n').trim(),
  };
}

function wallClock(instant: Date, timeZone: string): string {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(instant);
  const get = (type: Intl.DateTimeFormatPartTypes) => parts.find(p => p.type === type)?.value ?? '00';
  return `${get('year')}-${get('month')}-${get('day')}T${get('hour')}:${get('minute')}:${get('second')}`;
}

function sameOffset(a: Date, b: Date, timeZone: string): boolean {
  const offset = (d: Date) => Date.parse(`${wallClock(d, timeZone)}Z`) - Math.floor(d.getTime() / 1000) * 1000;
  return offset(a) === offset(b);
}

/**
 * A Graph start/end for an instant. Graph takes wall clock plus a zone, which shows the event in the
 * booking's zone. Near a DST change a wall-clock time can be ambiguous (the repeated hour), and
 * Exchange picks one of the two readings: there, send the instant as UTC wall clock instead.
 */
export function graphDateTime(instant: Date, timeZone: string): GraphDateTime {
  const window = 3 * 60 * 60 * 1000;
  const before = new Date(instant.getTime() - window);
  const after = new Date(instant.getTime() + window);
  if (sameOffset(before, after, timeZone)) return { dateTime: wallClock(instant, timeZone), timeZone };
  return { dateTime: instant.toISOString().slice(0, 19), timeZone: 'UTC' };
}

/** The attendee list Outlook will e-mail. Hosts without an address are left out: Graph rejects the event. */
export function attendeesFor(booking: BookingForCalendar, hostId: string, invite: InviteAttendees): GraphAttendee[] {
  const attendees: GraphAttendee[] = [];
  if (invite === 'all' && booking.guestEmail) {
    attendees.push({ emailAddress: { address: booking.guestEmail, name: booking.guestName ?? '' }, type: 'required' });
  }
  if (invite !== 'none') {
    for (const otherHost of booking.hostIds) {
      const email = booking.hostEmails[otherHost];
      if (otherHost !== hostId && email) attendees.push({ emailAddress: { address: email }, type: 'required' });
    }
  }
  return attendees;
}

/** The event body for one host's calendar. Pure. */
export function buildBookingEvent(
  booking: BookingForCalendar,
  hostId: string,
  text: { summary: string; description: string },
  options: { markerId: string; inviteAttendees: InviteAttendees }
): GraphEvent {
  return {
    subject: text.summary,
    body: { contentType: 'text', content: text.description },
    start: graphDateTime(booking.start, booking.timeZone),
    end: graphDateTime(booking.end, booking.timeZone),
    attendees: attendeesFor(booking, hostId, options.inviteAttendees),
    // Marks the event as ours, so busy sync skips it. Graph cannot take an id we choose, so this
    // marker is also how the event is found again.
    singleValueExtendedProperties: [{ id: options.markerId, value: booking.bookingId }],
    ...(booking.wantsTeamsMeeting ? { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' as const } : {}),
  };
}

/** Whether the stored event already shows what we would write; an unchanged PATCH would e-mail attendees again. */
function unchanged(existing: GraphEvent, event: GraphEvent, booking: BookingForCalendar): boolean {
  return (
    existing.subject === event.subject &&
    utcInstant(existing.start).getTime() === booking.start.getTime() &&
    utcInstant(existing.end).getTime() === booking.end.getTime()
  );
}

export interface EventWriteResult {
  calendarId: string;
  ok: boolean;
  error?: string;
}

async function writeToCalendar(
  deps: CalendarSyncDeps,
  booking: BookingForCalendar,
  calendar: LinkedCalendar
): Promise<EventWriteResult> {
  const describe = deps.describeBooking ?? defaultDescribeBooking;
  const markerId = bookingMarkerId(deps.config.markerPropertySetId);
  const event = buildBookingEvent(booking, calendar.userId, describe(booking), {
    markerId,
    inviteAttendees: deps.config.inviteAttendees,
  });
  try {
    const written = await withAccessToken(deps, calendar.userId, async accessToken => {
      // Find before create: Graph assigns ids, so the marker is the only way to recognise a booking
      // already written by an earlier attempt or trigger.
      const [first, ...duplicates] = await deps.graph.findEventsByMarker(
        accessToken,
        calendar.externalId,
        markerId,
        booking.bookingId
      );
      // Two triggers that raced both created an event: keep one, delete the rest.
      for (const duplicate of duplicates) {
        if (duplicate.id) await deps.graph.deleteEvent(accessToken, calendar.externalId, duplicate.id);
      }
      if (first?.id) {
        if (unchanged(first, event, booking)) return first;
        return deps.graph.updateEvent(accessToken, calendar.externalId, first.id, event);
      }
      // transactionId makes a POST that is retried after a lost response create one event, not two.
      return deps.graph.createEvent(accessToken, calendar.externalId, { ...event, transactionId: randomUUID() });
    });
    const joinUrl = written.onlineMeeting?.joinUrl;
    if (booking.wantsTeamsMeeting && joinUrl && deps.onConferenceLink) {
      await deps.onConferenceLink(booking.bookingId, joinUrl);
    }
    return { calendarId: calendar.id, ok: true };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    deps.log.error('microsoft-calendar-event-write-failed', {
      bookingId: booking.bookingId,
      calendarId: calendar.id,
      message,
    });
    return { calendarId: calendar.id, ok: false, error: message };
  }
}

/**
 * Booking confirmed (or rescheduled into a new booking): write it to every booking-sync calendar
 * of the primary host. Never throws — an Outlook outage must not fail the booking.
 */
export async function pushBookingEvent(
  deps: CalendarSyncDeps,
  booking: BookingForCalendar
): Promise<EventWriteResult[]> {
  const primaryHostId = booking.hostIds[0];
  if (!primaryHostId) return [];
  try {
    const calendars = (await deps.store.listCalendars(primaryHostId)).filter(c => c.syncBookings && c.isWritable);
    return await Promise.all(calendars.map(calendar => writeToCalendar(deps, booking, calendar)));
  } catch (error) {
    deps.log.error('microsoft-calendar-push-failed', {
      bookingId: booking.bookingId,
      message: error instanceof Error ? error.message : String(error),
    });
    return [];
  }
}

/**
 * Booking cancelled (or the old half of a reschedule): delete its events from every host's booking
 * calendars. Outlook sends attendees a cancellation, as it does for any meeting the organizer deletes.
 */
export async function removeBookingEvent(
  deps: CalendarSyncDeps,
  booking: Pick<BookingForCalendar, 'bookingId' | 'hostIds'>
): Promise<EventWriteResult[]> {
  const markerId = bookingMarkerId(deps.config.markerPropertySetId);
  const results: EventWriteResult[] = [];
  for (const hostId of booking.hostIds) {
    let calendars: LinkedCalendar[];
    try {
      calendars = (await deps.store.listCalendars(hostId)).filter(c => c.syncBookings);
    } catch {
      continue;
    }
    for (const calendar of calendars) {
      try {
        await withAccessToken(deps, hostId, async accessToken => {
          const events = await deps.graph.findEventsByMarker(accessToken, calendar.externalId, markerId, booking.bookingId);
          for (const event of events) {
            if (!event.id) continue;
            try {
              await deps.graph.deleteEvent(accessToken, calendar.externalId, event.id);
            } catch (error) {
              // Deleted meanwhile, by the host or a concurrent cancel: the state we want.
              if (!(error instanceof GraphApiError) || error.status !== 404) throw error;
            }
          }
        });
        results.push({ calendarId: calendar.id, ok: true });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        deps.log.error('microsoft-calendar-event-delete-failed', { bookingId: booking.bookingId, message });
        results.push({ calendarId: calendar.id, ok: false, error: message });
      }
    }
  }
  return results;
}

/** Booking sync switched on: write the host's upcoming bookings. Idempotent through the marker. */
export async function backfillBookings(deps: CalendarSyncDeps, calendar: LinkedCalendar): Promise<number> {
  const bookings = await deps.bookings.listUpcomingForPrimaryHost(calendar.userId, deps.now());
  let written = 0;
  for (const booking of bookings) {
    const result = await writeToCalendar(deps, booking, calendar);
    if (result.ok) written++;
  }
  return written;
}
```

## Checklist

- [ ] `pushBookingEvent` / `removeBookingEvent` called after the booking transaction commits
- [ ] Rescheduling removes the old booking's event before pushing the new one
- [ ] `inviteAttendees` chosen deliberately; the app's own e-mails adjusted to match
- [ ] `describeBooking` supplied in the host's language
- [ ] `onConferenceLink` stores the Teams URL where guests' e-mails read it
- [ ] The marker GUID never changes after the first event is written
