# microsoft-calendar-sync

[![Agent Skills](https://img.shields.io/badge/Agent_Skills-open_format-059669)](https://agentskills.io)
[![skills.sh](https://img.shields.io/badge/skills.sh-npx_skills_add-059669)](https://www.skills.sh)
[![Claude Code](https://img.shields.io/badge/Claude_Code-compatible-059669)](https://docs.claude.com/en/docs/claude-code/skills)
[![Codex CLI](https://img.shields.io/badge/Codex_CLI-compatible-059669)](https://developers.openai.com/codex/skills)
[![Gemini CLI](https://img.shields.io/badge/Gemini_CLI-compatible-059669)](https://github.com/google-gemini/gemini-cli/blob/main/docs/cli/skills.md)

An [Agent Skill](https://agentskills.io) that teaches an agent to build **two-way Microsoft 365 and Outlook
calendar sync** over Microsoft Graph, in a **Next.js App Router** app or any other Node server through an
Express adapter. Hosts connect their Microsoft account. Events on the calendars they choose block their
availability, kept current by Graph change notifications and a renewal cron. Confirmed bookings are written
to their calendar as events, as Teams meetings when asked. Firestore and Supabase/Postgres are both carried.

The two directions never meet. **Busy time is a disposable mirror of a time window, rebuilt by a full resync
and diffed by deterministic ids; a booking's event carries a private marker, is found again by it, and is
never read back as busy time.** Graph adds its own rules on top: refresh tokens rotate, `clientState` is
capped at 128 characters, notifications must be answered within 3 seconds, a calendar can have one
subscription, and all-day events float. The skill is built around each of them.

This skill was written by the engineer who builds the calendar integrations of a booking platform, and
audited against the earlier implementation of its Microsoft side. The templates hold the properties a Graph
calendar sync has to hold: a connection keeps working past its first access token because every rotated
refresh token is stored; every subscription fits Graph's `clientState` limit and is bound to the
notifications it signs; a notification is answered before any sync starts; a subscription Graph dropped is
replaced on the next cron run; an all-day vacation blocks the host's local days in the mailbox zone; a
booking pushed any number of times is one event. A 43-test `node:test` suite, run under four process time
zones, states each one, and [`references/provenance.md`](references/provenance.md) has the record.

## Install

One command, via the [skills.sh](https://www.skills.sh) CLI, which installs the skill into every
skills-compatible agent it detects, including Claude Code, Codex CLI and Gemini CLI:

```bash
npx skills add timerise-ai/microsoft-calendar-sync
```

Name the agents instead with `-a`, for example
`npx skills add timerise-ai/microsoft-calendar-sync -a claude-code -a codex`.

### Manual install

Nothing here is Claude-specific: the skill is a plain [Agent Skills](https://agentskills.io) folder,
`SKILL.md` plus markdown references with no file that calls a model, so cloning it into an agent's skills
directory is all an install is. For Claude Code:

```bash
git clone https://github.com/timerise-ai/microsoft-calendar-sync.git ~/.claude/skills/microsoft-calendar-sync
```

To scope it to a single project instead, clone it into that project's `.claude/skills/` directory. For another
agent, clone into that agent's skills directory, or symlink the Claude Code copy so one `git pull` updates
every agent:

```bash
mkdir -p ~/.agents/skills
ln -s ~/.claude/skills/microsoft-calendar-sync ~/.agents/skills/microsoft-calendar-sync
```

Update with `git pull` in the clone. The current release is **0.1.1**. See [CHANGELOG.md](CHANGELOG.md).

## Activation

The skill activates on tasks such as "connect our hosts' Outlook calendars", "block booked times from
Microsoft 365", "put every booking into the host's Outlook calendar as a Teams meeting", or "our Graph
calendar integration stops working after an hour". Each host matches skills its own way, so on a first run
invoke it explicitly: `/microsoft-calendar-sync` in Claude Code, `$microsoft-calendar-sync` in Codex CLI, and
through `/skills` in Gemini CLI.

## What's inside

| File | Contents |
|---|---|
| `SKILL.md` | Entry point: trigger, architecture, critical facts, hard rules, adaptation contract, quick start, reference directory |
| `references/data-model.md` | Records, types, the store contract, the rename table |
| `references/oauth-and-tokens.md` | Entra ID registration, scopes, the OAuth round trip, token rotation, encryption |
| `references/graph-api.md` | The fetch-based client and the Graph facts that shape it |
| `references/time-zones.md` | Mailbox zones, Windows names, all-day anchoring |
| `references/busy-sync.md` | Which events block, the sync run, feeding availability |
| `references/booking-events.md` | Writing and deleting booking events, invitations, Teams |
| `references/subscriptions.md` | Opening, renewing, replacing subscriptions; the cron |
| `references/calendars-and-settings.md` | The calendar list and the two toggles |
| `references/http-routes.md` | Route handlers, the webhook, Next.js and Express wiring |
| `references/firestore-backend.md` | Collections, index, rules, store |
| `references/supabase-backend.md` | Schema, RLS, optimistic concurrency, store |
| `references/operations.md` | Health, logs, configuration, runbook |
| `references/testing.md` | Running the suite, what it pins, mutation check, doubles |
| `references/test-suite.md` | The rules and client tests |
| `references/test-flows.md` | The flow tests |
| `references/provenance.md` | The audit ledger: fixed, changed by design, kept deliberately, designed here |
| `evals/` | `prompts.md` with what an operator types, and one file per agent eval run |
| `.github/workflows/agent-eval.yml` | Runs the agent evals on every published release |
| `CHANGELOG.md` | One section per release, newest first |
| `CLAUDE.md` | Editing conventions for an agent working on this repository |
| `LICENSE` | MIT |

The seam is the adaptation contract table in `SKILL.md`: the `CalendarStore` interface with its Firestore and
Supabase implementations, `authenticate(request)`, the `BookingSource` with `pushBookingEvent` and
`removeBookingEvent`, `defer` for background work and `hostTimeZone` for the business zone. They bound what
the host app supplies: its database, its auth, its booking status flow, its scheduler and its slot engine.
The calendar settings screen is the host app's too.

## The five non-negotiables

1. **Never take the user id from the OAuth callback URL.** The state is a key to a server-side, single-use,
   ten-minute record that holds the user id, so nobody can attach their calendar to someone else's account.
   The connect tests consume the state, refuse it a second time and after expiry, and read the user from
   the server.
2. **Never sync before answering a notification.** Graph wants a reply within 3 seconds and drops the
   notifications of a slow endpoint; the handler verifies, answers 202 and hands the sync to `defer`, with
   the cron as the safety net. The webhook test checks that nothing has synced when the response returns,
   and the routes test that syncs go to `defer`.
3. **Never trust a notification before its clientState verifies and names the subscription it arrived on.**
   The endpoint is public, so the signed clientState is the only credential. The rules suite checks the
   signature binds to its subscription, and the webhook test sends forged, foreign and unknown notifications
   that must schedule nothing.
4. **Never store tokens in clear text.** A refresh token is long-lived read and write access to a calendar,
   so both tokens are sealed with AES-256-GCM before they reach the store. The rules suite round-trips the
   encryption and the connect test checks the stored tokens are sealed.
5. **Never let an Outlook failure fail a booking.** The booking store is the truth about bookings; outbound
   writes log the error and return, and the booking stands. `pushBookingEvent` and `removeBookingEvent`
   catch at their boundary; this rule is checked by reading, as no test yet throws through them.

Everything else, the names, the screens, the strings, the store and the scheduler, is the host app's.

## Requirements

TypeScript on Node.js, with Next.js App Router route handlers or Express. `firebase-admin` or
`@supabase/supabase-js`; the Graph client is plain `fetch` with no Microsoft SDK. An Entra ID app
registration with delegated `Calendars.ReadWrite`, `MailboxSettings.Read` and `offline_access`, a public HTTPS
URL Graph can post notifications to, and a scheduler that calls the cron route every 12 hours. Running the
suite as written needs Node 22.18 or later.

## Verification

Every TypeScript block, written out at the path in its header comment, compiles as one project under
`strict` and `noUncheckedIndexedAccess`. The 43 tests in [test-suite.md](references/test-suite.md) and
[test-flows.md](references/test-flows.md) pass under `node --test` with the process time zone set to UTC,
Europe/Warsaw, Pacific/Auckland and America/Los_Angeles, and every mutation in the table in
[testing.md](references/testing.md) turns the suite red. `schema.sql` was applied to a scratch Postgres.

## Not this

| Problem | Go to |
|---|---|
| Google Calendar busy time and booking events | `google-calendar-sync` |
| Generating bookable slots and subtracting busy time | `time-slot-generator` |
| Public booking pages per host, Calendly-style | `calendar-scheduling-links` |
| Signing users in with Microsoft (SSO) | The host app's identity provider |
| Room and resource mailboxes, organisation-wide access without per-user consent | Application permissions with admin consent, a different design |
| Exchange Server on-premises | EWS, not Graph |

## Contributing

Issues and pull requests are welcome here. Pure markdown, with no build step, but the code blocks are checked:
every block names its destination on the first line, every TypeScript block is written to compile as one
project under `strict` and `noUncheckedIndexedAccess`, and the suite passes under `node --test` in four time
zones. Claims in this skill are meant to be verifiable: if you change a factual claim, say how you verified
it, whether against the Microsoft Graph reference, the Microsoft identity platform documentation, the MSAL
for Node source, the CLDR `windowsZones` table, the Firestore or Supabase documentation, or a reproduction.

Adding, removing or renaming a file in `references/` means updating the quick start and the reference
directory table in `SKILL.md`, the file table above, and any relative cross-links. Every odd-looking part of
the templates is there for a reason, and `references/provenance.md` is the ledger that must stay truthful:
read it before simplifying anything, and add an entry for anything you change. Commits follow Conventional
Commits and releases follow [STANDARD.md](https://github.com/timerise-ai/skills/blob/main/STANDARD.md) in the
index; `CLAUDE.md` carries the full editing conventions.

## Part of the Timerise Skills

This is one of the [Timerise Skills](https://github.com/timerise-ai/skills): modules for **Next.js App
Router** apps written by our own senior engineers from the modules they have shipped, not synthetic, each
published as its own repository and indexed there. They share one layout, so an agent that has read one knows
how to read the next: a `SKILL.md` entry point, `references/` loaded on demand, and a seam contract carrying
the module's non-negotiables.

## Author

Built and maintained by [Timerise](https://timerise.ai).

## License

MIT. See [LICENSE](LICENSE).
