# microsoft-calendar-sync

[![Agent Skills](https://img.shields.io/badge/Agent_Skills-open_format-059669)](https://agentskills.io)
[![skills.sh](https://img.shields.io/badge/skills.sh-npx_skills_add-059669)](https://www.skills.sh)
[![Claude Code](https://img.shields.io/badge/Claude_Code-compatible-059669)](https://docs.claude.com/en/docs/claude-code/skills)
[![Codex CLI](https://img.shields.io/badge/Codex_CLI-compatible-059669)](https://developers.openai.com/codex/skills)
[![Gemini CLI](https://img.shields.io/badge/Gemini_CLI-compatible-059669)](https://github.com/google-gemini/gemini-cli/blob/main/docs/cli/skills.md)

An [Agent Skill](https://agentskills.io) that teaches an agent to build two-way Microsoft 365 / Outlook
calendar sync in a TypeScript backend, over Microsoft Graph. Hosts connect their Microsoft account. Events
on the calendars they choose block their availability, kept current by Graph change notifications and a
renewal cron. Confirmed bookings are written to their calendar as events, as Teams meetings when asked.
Firestore and Supabase/Postgres are both carried.

Busy time is a disposable mirror of a time window, rebuilt by a full resync and diffed by deterministic
ids. Booking events carry a private marker, are found again by it, and are never read back as busy time.
Graph adds its own rules on top: refresh tokens rotate, `clientState` is capped at 128 characters,
notifications must be answered within 3 seconds, a calendar can have one subscription, and all-day events
float. The skill is built around each of them.

It was extracted from the Microsoft calendar integration of a production booking API and audited before
extraction. The audit found that connections stopped working an hour after they were made and that
Graph could not accept a single subscription, so busy time was never imported. It also found
notifications answered too slowly, lost subscriptions never replaced, vacations ignored, and duplicate
events. All are fixed in the templates, most of them pinned by a test.
[`references/provenance.md`](references/provenance.md) records what changed, what was kept deliberately
and why. Its sibling for Google Calendar is `google-calendar-sync`.

## Install

```bash
npx skills add <owner>/microsoft-calendar-sync
```

Or copy this folder into your agent's skills directory (`~/.claude/skills/`, `~/.codex/skills/`,
`~/.gemini/skills/`).

## Contents

| File | What it answers |
|---|---|
| [`SKILL.md`](SKILL.md) | When to use it, critical facts, hard rules, the seam contract |
| [`references/data-model.md`](references/data-model.md) | Records, types, the store contract, the rename table |
| [`references/oauth-and-tokens.md`](references/oauth-and-tokens.md) | Entra ID registration, scopes, the OAuth round trip, token rotation, encryption |
| [`references/graph-api.md`](references/graph-api.md) | The fetch-based client and the Graph facts that shape it |
| [`references/time-zones.md`](references/time-zones.md) | Mailbox zones, Windows names, all-day anchoring |
| [`references/busy-sync.md`](references/busy-sync.md) | Which events block, the sync run, feeding availability |
| [`references/booking-events.md`](references/booking-events.md) | Writing and deleting booking events, invitations, Teams |
| [`references/subscriptions.md`](references/subscriptions.md) | Opening, renewing, replacing subscriptions; the cron |
| [`references/calendars-and-settings.md`](references/calendars-and-settings.md) | The calendar list and the two toggles |
| [`references/http-routes.md`](references/http-routes.md) | Route handlers, the webhook, Next.js and Express wiring |
| [`references/firestore-backend.md`](references/firestore-backend.md) | Collections, index, rules, store |
| [`references/supabase-backend.md`](references/supabase-backend.md) | Schema, RLS, optimistic concurrency, store |
| [`references/operations.md`](references/operations.md) | Health, logs, configuration, runbook |
| [`references/testing.md`](references/testing.md) | Running the suite, what it pins, mutation check, doubles |
| [`references/test-suite.md`](references/test-suite.md) | The rules and client tests |
| [`references/test-flows.md`](references/test-flows.md) | The flow tests |
| [`references/provenance.md`](references/provenance.md) | Defects fixed, choices kept, additions |

## License

MIT
