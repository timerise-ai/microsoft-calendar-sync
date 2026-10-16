# Operations

What an operator and a host must be able to see and do, the configuration, and the runbook for the
failures this integration actually has.

## Health, derived

Each calendar's health is computed on read from the connection, the row and its subscription. Nothing
stores a status, so nothing can report "fine" because the job that would have updated it never ran.

| Health | Meaning | What fixes it |
|---|---|---|
| `reauth-required` | The token endpoint refused the refresh token | The user reconnects (same account keeps settings) |
| `sync-failing` | The last sync or transition recorded an error | Read `lastSyncError`; the next notification or cron run retries |
| `subscription-expired` | The subscription's expiry passed without renewal | The cron is not running or is failing |
| `no-subscription` | Busy sync is on but no subscription exists | Next cron run opens one |
| `stale` | Last successful sync older than `staleAfterMs` | The cron is not reaching this calendar |
| `never-synced` | Busy sync on, first sync not finished | Wait, or look at the logs |
| `ok` / `off` | Working / both syncs off | Nothing |

```ts
// microsoft-calendar/health.ts
import type { CalendarConnection, LinkedCalendar, WatchSubscription } from './types';

export type CalendarHealth =
  | 'off'
  | 'ok'
  | 'reauth-required'
  | 'never-synced'
  | 'sync-failing'
  | 'no-subscription'
  | 'subscription-expired'
  | 'stale';

/** Problems first, in the order an operator should fix them. */
export const HEALTH_SEVERITY: Record<CalendarHealth, number> = {
  'reauth-required': 0,
  'sync-failing': 1,
  'subscription-expired': 2,
  'no-subscription': 3,
  stale: 4,
  'never-synced': 5,
  ok: 6,
  off: 7,
};

/**
 * Derived, never stored: a stored status needs a job to keep it true, and a failed job makes
 * everything look healthy. `staleAfterMs` should exceed the cron interval with margin.
 */
export function deriveCalendarHealth(input: {
  connection: CalendarConnection;
  calendar: LinkedCalendar;
  subscription: WatchSubscription | null;
  now: Date;
  staleAfterMs: number;
}): CalendarHealth {
  const { connection, calendar, subscription, now, staleAfterMs } = input;
  if (connection.status === 'reauth_required') return 'reauth-required';
  if (!calendar.syncBusy) return calendar.syncBookings ? 'ok' : 'off';
  if (calendar.lastSyncError) return 'sync-failing';
  if (!subscription) return 'no-subscription';
  if (subscription.expiresAt <= now) return 'subscription-expired';
  if (!calendar.lastSyncedAt) return 'never-synced';
  if (now.getTime() - calendar.lastSyncedAt.getTime() > staleAfterMs) return 'stale';
  return 'ok';
}
```

## What each surface should show

| Surface | Shows | Lets you |
|---|---|---|
| Host's settings screen | account e-mail, connection status, each calendar with its two toggles, health, last synced | connect, reconnect, disconnect, toggle, rename |
| Admin view (per business) | every host's connection status and worst calendar health, sorted by `HEALTH_SEVERITY` | see who must reconnect before their availability is wrong |
| Logs | the `kind` values below | filter one failure class |

The reconnect prompt is the most important element on the host's screen: a `reauth_required`
connection silently stops blocking busy time, and the host's next double booking is how they find out.
Consider e-mailing the host when their connection flips to `reauth_required`.

## Log kinds

| Kind | Level | Meaning |
|---|---|---|
| `microsoft-calendar-connected` / `-disconnected` | info | Connection lifecycle |
| `microsoft-calendar-reauth-required` | warn | `invalid_grant` or `interaction_required` on refresh |
| `microsoft-calendar-mailbox-zone-unavailable` | info | No `MailboxSettings.Read`; all-day events use the host's zone |
| `microsoft-calendar-synced` | info | Per sync: events listed, upserted, deleted, skipped by reason |
| `microsoft-calendar-sync-failed` | error | Sync aborted; nothing was deleted |
| `microsoft-calendar-renewed` | info | Cron report: renewed, opened, synced, failed, orphans deleted |
| `microsoft-calendar-renew-failed` | error | One calendar's renewal or resync failed |
| `microsoft-calendar-subscription-delete-failed` | warn | Usually harmless: Graph already forgot it |
| `microsoft-calendar-webhook-bad-client-state` / `-subscription-mismatch` | warn | Forged or misrouted notification |
| `microsoft-calendar-webhook-unknown-subscription` | info | A subscription deleted here that Graph still delivers for |
| `microsoft-calendar-webhook-sync-failed` | error | A deferred sync failed; the row records it |
| `microsoft-calendar-event-write-failed` / `-delete-failed` | error | A booking's event was not written / not removed |
| `microsoft-calendar-transition-failed` | error | A toggle's side effect failed; the cron completes it |

The `skipped` counts on `microsoft-calendar-synced` answer most "why is this event not blocking?"
questions: `free`, `working-elsewhere`, `declined`, `own-booking`.

## Configuration

| Variable | Example | Notes |
|---|---|---|
| `MICROSOFT_CALENDAR_CLIENT_ID` | GUID | Application (client) ID |
| `MICROSOFT_CALENDAR_CLIENT_SECRET` | | Secret manager; expires, rotate before it does |
| `MICROSOFT_CALENDAR_TENANT` | `organizations` | `common` to allow personal accounts; a tenant id for one organisation |
| `MICROSOFT_CALENDAR_REDIRECT_URI` | `https://api.example.com/calendar/microsoft/callback` | Registered on the app, per environment |
| `MICROSOFT_CALENDAR_NOTIFICATION_URL` | `https://api.example.com/calendar/microsoft/notifications` | Public HTTPS, deployed before busy sync is offered |
| `MICROSOFT_CALENDAR_CLIENT_STATE_SECRET` | 32+ random bytes | Rotating it orphans every subscription: delete the subscription records, then run the cron, which clears Graph's copies (409) and opens fresh ones |
| `MICROSOFT_CALENDAR_MARKER_GUID` | `crypto.randomUUID()` | Never changes; existing events are recognised by it |
| `MICROSOFT_CALENDAR_DEFAULT_TIME_ZONE` | `Europe/Warsaw` | All-day fallback; the business's zone |
| `CALENDAR_TOKEN_KEY` | `openssl rand -hex 32` | Losing it means every user reconnects |
| `CALENDAR_RETURN_URL` | `https://app.example.com/settings/calendar` | Where the callback lands |
| `CRON_SECRET` | | Bearer for the cron route |

Commit a `.env.example` with these eleven names and empty values, and un-ignore it when the app ignores
`.env*`. The host's own store variables (the Supabase URL and service-role key, or the Firebase
credentials) sit beside them; nothing else is added for the calendar. Invitation mode, horizon and the
other `DEFAULT_SYNC_SETTINGS` are code, not variables, and no value falls back to a made-up default:
`setup.ts` names the missing variable on the first request that needs it.

The operator needs four things that no code can do for them: the Entra app registration (redirect URI,
the delegated permissions in [oauth-and-tokens.md](oauth-and-tokens.md)), the variables of `.env.example`
filled in (the marker GUID and the token key generated once and never changed), the notification URL served on public
HTTPS before busy sync is offered, and the cron below every 12 hours. An agent that builds this module
ends its summary with those four, each with its detail: the redirect URI and permissions, the GUID and key
generated once, the notification URL, the 12-hour schedule.

Schedule the cron every 12 hours (Cloud Scheduler, Vercel Cron, Kubernetes CronJob):

```bash
curl -X POST -H "Authorization: Bearer $CRON_SECRET" https://api.example.com/calendar/microsoft/cron
```

Its run time grows with the number of busy calendars: one subscription check and one full sync each,
four at a time (Outlook allows four concurrent requests per mailbox). With thousands of calendars,
run it as a job rather than a request, or give the HTTP timeout room.

## Runbook

| Symptom | Check | Fix |
|---|---|---|
| Every connection `reauth_required` about an hour after connecting | Was the refresh token stored? | Store `refresh_token` from every token response; users reconnect once |
| Busy sync on, no notifications ever | Subscription creation errors; clientState length; is the notification URL reachable and answering the validation? | Fix the URL or the clientState; the cron then opens subscriptions |
| Notifications arrive ten minutes late, or stop | Notification route latency | Answer before syncing (`defer`); Graph leaves the slow state on its own |
| Busy time stopped updating for everyone | `microsoft-calendar-renewed` in the last 12 h | Restore the cron; the next run renews, replaces and resyncs |
| A vacation does not block | Was it `showAs: oof`? `skipped.free` count; connection `mailboxTimeZone` | Free all-day events do not block by design; set the event to Out of office |
| Group session full after one booking | `own-booking` count is 0 while bookings exist | Outbound events lack the marker, or the marker GUID changed |
| Guests receive Outlook invitations they did not expect | `inviteAttendees` | `hosts` or `none` |
| Users of one company cannot connect | The callback lands with `admin-consent-required` | Their admin opens the admin-consent link once |
| Sandbox and production steal each other's subscriptions | One app registration shared | One registration per environment |

## Development

Graph cannot deliver notifications or validation requests to `localhost`, and refuses to create a
subscription whose URL fails validation. Either expose the API through an HTTPS tunnel and set
`MICROSOFT_CALENDAR_NOTIFICATION_URL` to the tunnel's URL, or leave busy sync off locally and test
with the cron route, which resyncs without notifications (opening the subscription fails and is
recorded). The OAuth redirect works on `http://localhost` when that exact URI is registered.

## Optional extensions

- **Retry queue for outbound writes.** Failed `pushBookingEvent` results are logged and returned, not
  retried. A host that needs guarantees can store failed `(bookingId, calendarId)` pairs and replay them
  from the cron; find-before-create makes replays safe.
- **Delta queries.** `calendarView/delta` cuts traffic for very large calendars, at the cost of keeping
  a delta token per calendar and handling its expiry with a full resync. The full-window resync here is
  simpler and repairs itself.
