---
name: microsoft-calendar-sync
description: >
  Build two-way Microsoft 365 / Outlook calendar sync for a booking or scheduling backend over
  Microsoft Graph: hosts connect their Microsoft account with OAuth, events on chosen calendars block
  their availability (change-notification subscriptions plus a renewal cron), and confirmed bookings
  are written to their calendar as events, with Teams meeting links. Use when: (1) an app with
  bookings, appointments, classes or shifts must respect staff's Outlook busy time or put bookings into
  their Outlook calendar, (2) an existing Outlook or Graph calendar integration stops working an hour
  after connecting, never receives notifications, ignores vacations, duplicates events or loses its
  subscriptions, (3) the user mentions: Outlook calendar integration, Microsoft 365 calendar, Office 365
  calendar sync, Microsoft Graph calendar, Entra ID app registration, Calendars.ReadWrite, calendarView,
  change notifications, Graph subscription, clientState, validationToken, lifecycleNotificationUrl,
  subscriptionRemoved, MSAL refresh token, AADSTS700082, invalid_grant, interaction_required,
  singleValueExtendedProperties, teamsForBusiness, showAs, isAllDay, mailboxSettings timeZone.
  Carries a dependency-free Graph client on fetch, a pure busy-time mapper that anchors all-day events
  to the mailbox zone, single-flight sync, booking events found again by an extended-property marker,
  in-place subscription renewal that replaces lost subscriptions, derived health per calendar,
  Firestore and Supabase stores, and a 43-test node:test suite run under four process time zones.
  TypeScript for Next.js App Router route handlers, with an Express adapter for any other Node server;
  storage, auth and the booking source are seams behind CalendarStore, authenticate and BookingSource.
  Not Google Calendar, not Exchange on-premises (EWS), not a calendar UI.
---

# Microsoft Calendar Sync

Two directions, one rule each. **Inbound**: Outlook is the truth about a host's other commitments, so
busy blocks are a disposable mirror of a time window, rebuilt by a full resync and never edited by
hand. **Outbound**: the booking store is the truth about bookings, so an event written to Outlook
carries a private marker with the booking id, is found again by it, and is never read back as busy
time. Unlike Google, Graph picks event ids, wants answers within 3 s and allows one subscription per calendar.

## When to use

- Hosts connect Microsoft 365 or Outlook.com and pick calendars that block their availability.
- Bookings appear in the host's Outlook calendar, optionally as Teams meetings, and vanish on cancel.
- Auditing an existing Graph integration: dead after an hour, silent subscriptions, missed vacations.
- Moving that sync between Firestore and Supabase/Postgres.

## When NOT to use

- **Google Calendar**: use `google-calendar-sync`. Same architecture and record shapes, different API.
- **Generating bookable slots, subtracting busy time**: use `time-slot-generator`. This skill *produces*
  the busy blocks it consumes.
- **Signing users in with Microsoft (SSO)** is an identity concern; this skill asks for calendar access
  of an already signed-in user.
- **Room and resource mailboxes, or organisation-wide access without per-user consent** need
  application permissions with admin consent, a different design.
- **Exchange Server on-premises** speaks EWS, not Graph.

## Architecture

```
          connect (PKCE, single-use state) -> callback -> id token tid:oid + encrypted tokens + mailbox zone -> rows
                                                                                        |
 inbound   Graph -- POST {value[]} -> /notifications -> verify clientState -> 202 -> defer(syncBusy) -> calendarView pages
              ^                                                                       -> toBusyBlock -> diffBusy -> busy blocks
              +-- subscription <-- cron: PATCH renew, 404 replace, 409 clear, resync every busy calendar
 outbound  booking confirmed -> find by marker -> unchanged? skip : PATCH - none? POST (transactionId) -> Teams link back
           booking cancelled -> find by marker -> DELETE each (404 = done)
```

## Critical facts

1. **Store the refresh token from every token response.** Microsoft rotates it on each refresh.
   MSAL for Node keeps it in its own cache and never returns it, so an integration that stores MSAL's
   result has no refresh token and dies when the first access token expires, about an hour in.
2. **`clientState` is at most 128 characters.** Graph refuses the subscription otherwise. An Outlook
   calendar id alone is around 150. Put only your subscription id in it, signed.
3. **Answer notifications within 3 seconds, then sync.** Graph marks a slower endpoint *slow*, delays
   it, then drops its notifications for good. Verify, return 202, run the sync after the response.
4. **One subscription per calendar; renew in place.** A second POST for the same resource is a 409, so
   Google's new-before-old renewal does not apply. PATCH the expiry; a 404 means Graph dropped it, and
   only then is a new one created.
5. **All-day events float.** Graph returns them as midnight to midnight in the zone you requested, so
   in UTC they are the right dates at the wrong instants. Anchor them to the mailbox's zone, which
   Outlook usually reports as a Windows name.
6. **Graph assigns event ids.** A booking's event is found again through a single-value extended
   property. Find before create; collapse duplicates; skip a write that changes nothing, because
   every PATCH to a meeting with attendees e-mails them.
7. **Attendees are e-mailed; there is no switch.** Inviting the guest sends them an Outlook invitation
   from the host, and deleting the event sends a cancellation. `inviteAttendees` decides who is listed.

## Hard rules

> **Never take the user id from the OAuth callback URL.** The state is a key to a server-side,
> single-use, ten-minute record. Otherwise anyone attaches their calendar to someone else's account.

> **Never sync before answering a notification.** Hand the work to `defer`; the cron is the safety net.

> **Never trust a notification before its clientState verifies and names the subscription it arrived on.**
> The endpoint is public; the signed clientState is the only credential.

> **Never store tokens in clear text.** A refresh token is long-lived read/write access to a calendar.

> **Never let an Outlook failure fail a booking.** Outbound writes log and return; the booking stands.

## Adaptation Contract

This table is the seam contract; the skill has no `adaptation.md`. The rename table is in
[data-model.md](references/data-model.md), the order of work is the quick start below.

| Seam | The skill ships | The host supplies |
|---|---|---|
| Domain entities | Connection, LinkedCalendar, WatchSubscription, BusyBlock, BookingForCalendar | Its names (staff, resource, appointment) |
| Tenant scope | Per user: one Microsoft account per `userId` | Its user id; org scoping stays in its guard |
| Auth guard | `authenticate(request)`, returning `userId \| null` | Session, Clerk, NextAuth, Firebase, Supabase |
| Data access | `CalendarStore` + Firestore and Supabase implementations | Its ORM or SDK |
| Bookings | `BookingSource` + `pushBookingEvent` / `removeBookingEvent` | Calls them from its booking status flow |
| Availability | Busy blocks per user, `listHostBusy` query | Its slot engine (`time-slot-generator`) |
| Time zone | Mailbox zone; `hostTimeZone(userId)`; `defaultTimeZone` | The host's business zone |
| HTTP | Web `Request`/`Response` handlers | Next.js route handlers, Express adapter |
| Background work | `defer` for notifications; `renewSubscriptions` every 12 h | `after()`, `waitUntil`, Cloud Scheduler, Vercel Cron |
| Strings | `describeBooking` (event title/body), error codes | Its i18n |
| Secrets | Client id/secret, clientState secret, 32-byte token key, marker GUID | Its secret manager |
| Tests | `node:test` suite, memory store, fake Graph | Its runner |

## Quick start

1. Register the app in Entra ID and generate the marker GUID:
   [oauth-and-tokens.md](references/oauth-and-tokens.md).
2. Copy the types and decide the rename: [data-model.md](references/data-model.md).
3. Copy the Graph client: [graph-api.md](references/graph-api.md); and the zone map:
   [time-zones.md](references/time-zones.md).
4. Copy inbound sync: [busy-sync.md](references/busy-sync.md); outbound:
   [booking-events.md](references/booking-events.md); subscriptions:
   [subscriptions.md](references/subscriptions.md); the calendar list and toggles:
   [calendars-and-settings.md](references/calendars-and-settings.md).
5. Pick a store: [firestore-backend.md](references/firestore-backend.md) or
   [supabase-backend.md](references/supabase-backend.md).
6. Mount the routes, `defer` and the cron: [http-routes.md](references/http-routes.md).
7. Copy the suite and run it under two `TZ` values: [testing.md](references/testing.md),
   [test-suite.md](references/test-suite.md), [test-flows.md](references/test-flows.md).
8. Give operators the health view: [operations.md](references/operations.md).
9. Repairing an existing Graph integration? The fix order is in [provenance.md](references/provenance.md).

## Reference directory

| Scenario | Trigger keywords | Reference |
|---|---|---|
| Entities, store contract, rename | LinkedCalendar, WatchSubscription, BusyBlock, CalendarStore, mutateCalendar, version | [data-model.md](references/data-model.md) |
| Connecting an account | Entra ID, app registration, redirect URI, tenant, organizations, common, scopes, PKCE, id_token, tid, oid, refresh token rotation, invalid_grant, interaction_required, admin consent, AES-256-GCM | [oauth-and-tokens.md](references/oauth-and-tokens.md) |
| Talking to Graph | fetch, calendarView, Prefer outlook.timezone, @odata.nextLink, $filter extended property, 429, Retry-After, subscriptions, MSAL | [graph-api.md](references/graph-api.md) |
| Mailbox zones | mailboxSettings, Windows time zone, "W. Europe Standard Time", IANA, CLDR windowsZones | [time-zones.md](references/time-zones.md) |
| Busy time in | toBusyBlock, showAs, free, workingElsewhere, oof, tentative, declined, isCancelled, isAllDay, diffBusy, lease, horizon | [busy-sync.md](references/busy-sync.md) |
| Bookings out | singleValueExtendedProperties, marker GUID, transactionId, attendees, invitations, Teams, isOnlineMeeting, joinUrl, backfill | [booking-events.md](references/booking-events.md) |
| Subscriptions | clientState, 128 characters, expirationDateTime, 10080 minutes, PATCH renew, 404, 409 Conflict, lifecycle, subscriptionRemoved, missed | [subscriptions.md](references/subscriptions.md) |
| Calendar list and toggles | canEdit, isDefaultCalendar, shared calendar, syncBusy, syncBookings, planTransitions | [calendars-and-settings.md](references/calendars-and-settings.md) |
| Routes and webhook | validationToken, text/plain, 202, 3 seconds, defer, after, waitUntil, callback redirect, consent_required, cron secret, Express | [http-routes.md](references/http-routes.md) |
| Firestore | firebase-admin, runTransaction, batch 500, composite index, security rules | [firestore-backend.md](references/firestore-backend.md) |
| Supabase / Postgres | schema.sql, RLS, service role, optimistic concurrency, 23505, PostgREST 1000 rows | [supabase-backend.md](references/supabase-backend.md) |
| Proving it | node:test, memory store, fake Graph, TZ=, mutation check | [testing.md](references/testing.md), [test-suite.md](references/test-suite.md), [test-flows.md](references/test-flows.md) |
| Running it | health, stale, subscription-expired, env vars, throttling, runbook, tunnel, localhost | [operations.md](references/operations.md) |
| Why it is built this way | provenance, defects, kept deliberately, fix order | [provenance.md](references/provenance.md) |

Part of the [Timerise Skills](https://github.com/timerise-ai/skills) index, which lists the sibling skills.
