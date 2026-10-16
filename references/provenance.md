# Provenance

This module is the Microsoft 365 calendar integration of a booking platform, written down by the engineer
who builds it and audited against the earlier implementation: a Node API on a document database that offers
hosts Google and Microsoft calendars side by side. Its Microsoft side shared the busy-time machinery of its
Google side, which had already been hardened: single-use OAuth state with PKCE, encrypted tokens at rest,
deterministic busy-block ids with a row-based diff, own bookings skipped by their marker, the free, declined
and cancelled rules, every page of the calendar view, and a cron that resyncs every busy calendar. Those are
carried over. The architecture, record shapes and test method follow the sibling skill
`google-calendar-sync`.

The audit read the Microsoft integration end to end against Graph's documentation and the MSAL library it
used. This file is the ledger for the person editing the skill: what the audit fixed and how the templates
verify it, what it changed by design, what was kept deliberately and why it is safe, and what was designed
here and has never run in production. Each defect below was confirmed against the earlier implementation;
each is fixed in the templates and, where it can be, pinned by a test in [test-suite.md](test-suite.md) or
[test-flows.md](test-flows.md).

## Fixed in the templates

### 1. Every connection stopped working about an hour after it was made

The integration used `@azure/msal-node` and read the refresh token from MSAL's result. MSAL keeps
refresh tokens in its own in-memory cache and does not put them on the result (confirmed in MSAL 5.1's
response handler), so the code stored an empty refresh token and logged a warning. When the first
access token expired, every call failed with "refresh token is missing, user must re-authenticate":
busy time froze and bookings stopped appearing. **Shipped:** the token endpoint over `fetch`, and the
refresh token from every response stored, rotation included, in [oauth-and-tokens.md](oauth-and-tokens.md).

### 2. Graph could not accept a single subscription, so busy time was never imported

The `clientState` held the calendar id, our subscription id, the user id and a signature. Graph caps
`clientState` at 128 characters, and an Outlook calendar id alone runs to about 150. And because the
toggle and the cron both opened the subscription *before* the first sync in the same step, the failure
also skipped the sync: a calendar with busy sync on never imported anything. **Shipped:** the
clientState carries only a UUID and its HMAC (80 characters), and a subscription failure is recorded
without skipping the sync, in [subscriptions.md](subscriptions.md).

### 3. Notifications were answered only after a full sync

The webhook listed the whole calendar and wrote its busy blocks before replying, once per notification.
Graph wants a 2xx within 3 seconds; slower endpoints are delayed by ten minutes and then have their
notifications dropped. **Shipped:** verify, answer 202, sync through the host's `defer`, with a
single-flight lease collapsing the burst, in [http-routes.md](http-routes.md).

### 4. A subscription Graph dropped was never replaced

Renewal was a PATCH; when Graph answered 404 (expired after a few days without the cron, or removed),
the error was logged and the PATCH retried on every run, forever. The cron's repair step only looked
for a missing local record, which still existed. **Shipped:** 404 replaces the subscription; 409 on
create clears the orphan Graph still holds; lifecycle notifications trigger the repair at once.

### 5. Vacations did not block

All-day events were skipped entirely, because Graph returns them as floating midnights and anchoring
them needs the mailbox's zone. Out-of-office days, the way most people enter a vacation, left the host
bookable. **Shipped:** the mailbox zone (Windows names mapped to IANA), else the host's zone, else a
configured default, in [time-zones.md](time-zones.md).

### 6. Only the first page of calendars was listed

`/me/calendars` was read once and `@odata.nextLink` ignored; calendars past the first page could not be
chosen. **Shipped:** every page.

### 7. A revoked grant looked like "no calendars", and reconnecting wiped every setting

Listing errors returned an empty list, and nothing recorded `invalid_grant`. Reconnecting deleted every
Microsoft calendar row and its busy blocks, even for the same account. **Shipped:** `reauth_required`
(for `interaction_required` too), surfaced through health; the account is identified by the id token's
`tid:oid`, and the same account keeps its rows and gets its subscriptions renewed.

### 8. Booking events were not idempotent

Each push POSTed a new event. A repeated trigger created a second one; the backfill de-duplicated only
against the 180-day window. **Shipped:** find by marker first, PATCH or skip, POST with a
`transactionId`, and collapse duplicates, in [booking-events.md](booking-events.md).

### 9. Smaller ones

- A co-host without an e-mail address (or missing from the booking's host map) made the event fail.
- Calendar ids went into URL paths unencoded; they contain `/`, `+` and `=`.
- The booking marker used the property-set GUID from Microsoft's documentation examples, which other
  apps copy too. The GUID is now configuration.
- The notification's binding to its subscription was checked only when Graph sent a subscription id.
- A legacy path for unsigned clientStates outlived its own sunset date.
- MSAL's in-memory cache accumulated every user's tokens for the life of the process.
- Booking times near a DST change were sent as ambiguous wall-clock times.

## Kept deliberately

- **Full-window resync per notification**, not delta queries: more traffic, self-repairing.
- **180-day horizon, 12-hour cron, 48-hour renewal window.**
- **`organizations` as the default tenant**: work and school accounts, where Teams meetings exist.
  `common` adds personal accounts.
- **Attendees invited by default** (`inviteAttendees: 'all'`), as the earlier implementation did: guests
  and co-hosts receive the Outlook invitation. `hosts` and `none` are options.
- **Teams join URLs written back to the booking** through `onConferenceLink`.
- **Only the primary host's calendars receive the event**; deletion walks every host.
- **`canEdit` decides writability.**
- **Turning booking sync off leaves existing events**; cancelling one of their bookings still removes it.
- **One Microsoft account per user.**

## Changed by design

- **`fetch` instead of MSAL and axios**, so the refresh token is visible and the client can be faked.
- **`prompt=select_account` instead of `prompt=consent`**, which brings the admin-approval page back
  for users in tenants that disabled user consent, even after an administrator consented.
- **Six-day subscriptions** instead of two: Graph's maximum for Outlook events is 10,080 minutes.
- **Busy blocks hold UTC instants only**, without a wall-clock copy.
- **Top-level, prefixed collections and tables**, so the cron, webhook and availability queries are
  single queries, and Google and Microsoft records never collide.
- **Web `Request`/`Response` handlers** instead of Express routes plus GraphQL resolvers.

## Designed here, not yet run in production

None of these ran in the earlier implementation. Each is covered by the suite where a test can reach it,
and none has carried production traffic yet.

- Lifecycle notifications (`subscriptionRemoved`, `reauthorizationRequired`, `missed`).
- Throttling retries honouring `Retry-After`, and a guard that keeps next links on Graph.
- Derived per-calendar health, in [operations.md](operations.md).
- Single-flight sync lease with a queued rerun.
- Disconnect.
- Granted-scope check on connect (`canWrite`, mailbox zone).
- Orphan subscription clean-up in the cron.
- `inviteAttendees`, unchanged-write skip, duplicate collapse.
- Supabase/Postgres store and schema, in [supabase-backend.md](supabase-backend.md).
- The test suite, memory store and fake Graph, in [testing.md](testing.md).
- From the 0.1.0 agent evals: a notification entry that is not an object, or has no string
  `clientState`, is skipped instead of throwing a 500 that dropped the whole batch.
- From the 0.1.0 agent evals: cancelling a booking deletes its event from every writable calendar, so an
  event written before booking sync was turned off no longer outlives its booking.
- From the 0.1.0 agent evals: `setup.ts` builds the deps on the first request, so `next build` runs
  without the secrets and no agent adds fallbacks of its own; the Next.js cron route answers `GET` for
  Vercel Cron.

## Known limit

Co-hosts' copies of a booking's meeting carry no marker, because extended properties stay on the
organizer's item. A co-host who syncs busy time from the calendar the invitation lands in gets the
booking back as a busy block. See [busy-sync.md](busy-sync.md) for the two ways around it.

## Repairing an existing integration with these defects

Fix in this order, most damaging first:

1. The refresh token (#1): nothing works after the first hour. Every existing connection has an empty
   refresh token and must be reconnected once; there is nothing to recover.
2. The clientState and the sync-after-subscribe order (#2): without them no busy time is ever imported.
   No subscriptions exist to migrate; the first cron run after the fix opens them.
3. Answering before syncing (#3) and replacing lost subscriptions (#4): notifications that arrive late,
   then not at all.
4. All-day events (#5): bookable vacations.
5. Reauth state, same-account reconnect, calendar paging (#6, #7).
6. Idempotent events and the smaller ones (#8, #9).

**Keep the marker GUID the existing integration wrote** (set `MICROSOFT_CALENDAR_MARKER_GUID` to
it). Events already in hosts' calendars carry it; a new GUID would make busy sync import them as busy time
and the outbound side create a second copy of each.
