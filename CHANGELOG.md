# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.3] - 2026-10-16

From scoring the prompt-1 agent eval runs against 0.1.2. Wording only; no template changed.

### Changed

- Personal Outlook.com accounts are `MICROSOFT_CALENDAR_TENANT=common` in the operator handover, never a
  tenant fallback written into `setup.ts`; the rule moved into the quick start in `SKILL.md`, and
  `http-routes.md` names the case.

## [0.1.2] - 2026-10-16

From scoring the prompt-1 agent eval runs against 0.1.1. Wording only; no template changed.

### Changed

- `setup.ts` changes only the store line, the host seams, `log` and `defer`; the `config` block stays as
  shipped, and a different default is set through its variable and named in the summary
  (`SKILL.md`, `http-routes.md`).
- The shipped route handlers are mounted unwrapped; the host's own background work gets its own route
  (`http-routes.md`).
- The handover names each step's detail: redirect URI and permissions, the marker GUID and token key
  generated once, the public HTTPS notification URL, the 12-hour cron (`SKILL.md`, `operations.md`).

## [0.1.1] - 2026-10-16

Fix release, from scoring the prompt-1 agent eval runs against 0.1.0.

### Fixed

- A notification entry that is not an object, or carries no string `clientState`, is skipped like a
  forgery; it threw before, answering 500 and dropping the valid entries in the same batch. Apps built
  from 0.1.0 copy the new guard in `webhook.ts`.
- Cancelling a booking deletes its event from every writable calendar, so an event written before booking
  sync was turned off no longer outlives its booking. Apps built from 0.1.0 copy the new filter in
  `removeBookingEvent`.
- The Firestore store imports its types from `firebase-admin/firestore`, so it compiles against
  `firebase-admin` 14 as well as 13.

### Changed

- `setup.ts` builds the deps on the first request (`calendarRoutes()`, `calendarDeps()`), so `next build`
  runs without secrets; the Next.js cron route also answers `GET`, for Vercel Cron (`http-routes.md`).
- The quick start in `SKILL.md` says to copy templates verbatim, install the store SDK from the registry,
  run the three test files unchanged through `npm test`, keep client-sent user ids out of `authenticate`,
  write `.env.example` and hand the operator four steps, which `operations.md` and `testing.md` detail.
- The suite still has 43 tests; two of them assert more.

## [0.1.0] - 2026-10-16

First release: two-way Microsoft 365 calendar sync over Graph, written by the engineer who builds the
calendar integrations of a booking platform and audited against the earlier implementation.

### Added

- OAuth connection with single-use server-side state, PKCE, `prompt=select_account`, granted-scope
  check, account identity from the id token's `tid:oid`, AES-256-GCM token storage, refresh-token
  rotation, `invalid_grant` / `interaction_required` mark `reauth_required`, same-account reconnect that
  keeps settings, and disconnect.
- A dependency-free Microsoft Graph client on `fetch`, with throttling retries and a next-link guard.
- Mailbox time zones, with the CLDR Windows-to-IANA table.
- Busy sync: event rules (free, working elsewhere, declined, cancelled, own bookings, all-day events
  anchored to the mailbox zone), every page of the calendar view, a row-based diff over deterministic
  ids, and a single-flight lease with a queued rerun.
- Booking events found by an extended-property marker: find before create, `transactionId`, unchanged
  writes skipped, duplicates collapsed, Teams meetings, `inviteAttendees`, backfill.
- Subscriptions with an 80-character signed `clientState`, in-place renewal, 404 replacement, 409
  orphan clean-up, lifecycle notifications, and a cron that resyncs every busy calendar.
- Web `Request`/`Response` route handlers that answer Graph before syncing, with Next.js and Express wiring.
- Firestore and Supabase/Postgres stores, the Postgres schema with RLS.
- Derived per-calendar health and an operations runbook.
- A 43-test `node:test` suite with a memory store, a fake Graph and a scripted `fetch`, run under
  four process time zones.
