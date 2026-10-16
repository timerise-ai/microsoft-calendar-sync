# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

An [Agent Skill](https://agentskills.io) package: markdown only. There is no `package.json`, no build and
nothing here executes. It teaches an agent to build two-way Microsoft 365 / Outlook calendar sync over
Microsoft Graph in *someone else's* TypeScript backend, on Firestore or Supabase/Postgres. Its sibling,
`google-calendar-sync`, shares the architecture and record shapes; keep the two consistent.

The code blocks in `references/` are the templates an agent copies. They are kept as complete files (one
fenced block per file, headed by a `// microsoft-calendar/<file>.ts` comment, or `// app/…` and
`// server/…` for the wiring examples) so they can be extracted verbatim and checked. Do not add a
`package.json` or test config to make them runnable here.

## Structure

- `SKILL.md`: the router, loaded on every activation. Keep it at or under 150 lines. The `description`
  frontmatter is the activation trigger.
- `references/`: one question per file. `provenance.md` is the rationale layer; read it before
  "simplifying" anything, because most odd-looking choices are listed under "Kept deliberately".

## Verifying a change

Template edits must be re-verified the way 0.1.0 was:

1. Extract every `// microsoft-calendar/*.ts` block, the wiring blocks and `schema.sql` into a scratch
   project, at the paths in their header comments (the wiring imports `@/microsoft-calendar/setup`).
2. `tsc --strict` passes, and also with `--noUncheckedIndexedAccess`. Borrow `firebase-admin`,
   `@supabase/supabase-js`, `express`, `@types/express` and `@types/node` from any project.
3. `node --test` passes under `TZ=UTC`, `Europe/Warsaw`, `Pacific/Auckland` and `America/Los_Angeles`.
   Rewrite the extensionless imports to `.ts` for Node's type stripping.
4. Every mutation in the table in `references/testing.md` still turns the suite red.
5. `schema.sql` applies to a scratch Postgres.

Never change `bookingMarkerId`, `busyBlockId` or `calendarRowId` inputs. Every consumer's existing
events, blocks and rows are addressed by them.

## Releases

Bump `CHANGELOG.md` with Keep a Changelog headings. Patch for template fixes, minor for new capability.
