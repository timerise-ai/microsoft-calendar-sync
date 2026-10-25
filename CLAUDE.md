# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

An [Agent Skill](https://agentskills.io) package: markdown only. There is no `package.json`, build or dev
server here, and nothing in this repository executes. It teaches an agent to build two-way Microsoft 365 and
Outlook calendar sync over Microsoft Graph in *someone else's* TypeScript backend, on Next.js App Router route
handlers or Express, with Firestore or Supabase/Postgres. The version of record is the git tag.

The commands and code in `references/` describe the app the agent will generate, not this repository. The
one thing checked here is that the templates compile and their tests pass, in a scratch project; the recipe
is below. Do not add a `package.json` or test config to make them runnable here.

The skill was written by the engineer who owns the module, and audited against the earlier implementation.
`references/provenance.md` is the rationale layer and the ledger of that audit: what it fixed, what it
changed by design, what was kept deliberately with the reason it is safe, and what was designed here and has
never run in production. Read it before "simplifying" anything.

Sibling directories under `../` are other skills, not dependencies. `google-calendar-sync` shares this
skill's architecture and record shapes; keep the two consistent. `time-slot-generator` and
`calendar-scheduling-links` are named from `SKILL.md` and `README.md`. The standard every skill follows is
`../skills/STANDARD.md`.

## Structure

- `SKILL.md`: entry point, loaded whole on every activation, so it stays between 130 and 160 lines, the
  closing index line aside. The frontmatter `description` is the trigger surface. Its adaptation contract
  table is the seam contract; there is no `references/adaptation.md`, and the rename table is in
  `data-model.md`.
- `README.md`: the human-facing front door, in the section order every Timerise skill shares.
- `CHANGELOG.md`: Keep a Changelog, newest release first. The version lives here, in the README's
  current-release line, and in the git tag, and the three agree.
- `LICENSE`: MIT.
- `references/*.md`: one question per file. Each template file is one fenced block in the reference that
  explains it:

  | Reference | Templates |
  |---|---|
  | `data-model.md` | `types.ts` (records and the store contract every other file depends on) |
  | `oauth-and-tokens.md` | `crypto.ts`, `tokens.ts`, `connect.ts` |
  | `graph-api.md`, `time-zones.md` | `graphApi.ts`, `timeZones.ts` |
  | `busy-sync.md` | `busyRules.ts` (pure rules), `busySync.ts` |
  | `booking-events.md`, `subscriptions.md`, `calendars-and-settings.md` | `bookingEvents.ts`, `subscriptions.ts`, `calendars.ts` |
  | `http-routes.md` | `webhook.ts`, `routes.ts`, `setup.ts`, plus the `app/` and `server/` wiring |
  | `firestore-backend.md`, `supabase-backend.md` | `store/firestore.ts`; `store/schema.sql`, `store/supabase.ts` |
  | `operations.md` | `health.ts` |
  | `testing.md` | the doubles `testing/memoryStore.ts`, `testing/fakeGraph.ts` |
  | `test-suite.md`, `test-flows.md` | `rules.test.ts`, `graphApi.test.ts`; `flows.test.ts` |

- `evals/`: `prompts.md` holds what an operator types after installing, in their words; the first prompt
  is the agent eval run before every release. Every other file there is one eval run: measured frontmatter
  that is never edited, then the notes of the person who ran it. Add a prompt rather than rewording one that
  has results. The procedure is section 10 of the index's STANDARD.md.
- `.github/workflows/agent-eval.yml`: the caller of the index's reusable eval workflow, copied verbatim from
  section 10 of the standard and run on every published release and on a maintainer's dispatch. It is the
  same in every skill; never edit it, and never add a trigger on `push` or `pull_request`.
- `.github/cover.svg`: the README cover, rendered from the skill's cover motif on timerise.ai. It is not
  skill content; redraw it there, not here.

## Editing conventions

- **Code blocks name their destination on the first line**: `// microsoft-calendar/<file>.ts`, or
  `// app/...` and `// server/...` for the wiring, optionally followed by `: description`. The SQL uses
  `-- microsoft-calendar/store/schema.sql`. A block that continues a file already introduced omits it.
- **The code blocks are compiled and run.** A template edit is re-verified this way:

  1. Write every block out to a scratch project (never this repo) at the path in its header. The wiring
     imports `@/microsoft-calendar/setup`. This writes them all into `$OUT`:

     ~~~bash
     awk -v out="$OUT" '
     /^```/ { if (f) { close(f); f=""; next }; start=1; next }
     start { start=0; if (match($0, /^(\/\/|--) (microsoft-calendar|app|server)\/[^ ]+/)) {
       p=$0; sub(/^(\/\/|--) /,"",p); sub(/[ :].*/,"",p); f=out"/"p; d=f; sub(/\/[^\/]+$/,"",d)
       system("mkdir -p \"" d "\"") } }
     f { print > f }' references/*.md
     ~~~

  2. `tsc --strict` passes, and also with `--noUncheckedIndexedAccess`. Borrow `firebase-admin`,
     `@supabase/supabase-js`, `express`, `@types/express` and `@types/node` from any project.
  3. `node --test` passes under `TZ=UTC`, `Europe/Warsaw`, `Pacific/Auckland` and `America/Los_Angeles`,
     with the extensionless imports rewritten to `.ts` for Node's type stripping, as
     `references/testing.md` section Running shows (Node 22.18+, `--experimental-strip-types`, all three
     test files). For one test, add `--test-name-pattern='<name>'`.
  4. Every mutation in the table in `references/testing.md` still turns the suite red.
  5. `schema.sql` applies to a scratch Postgres.

- **Test counts are claims.** The suite has 43 tests; `SKILL.md`, `README.md` and `CHANGELOG.md` repeat the
  number. Change a suite, change the numbers.
- **Keep the three indexes in sync** with `references/`: the quick start and the reference directory in
  `SKILL.md`, and the file table in `README.md`.
- **Identifiers are shared across files.** `CalendarStore`, `mutateCalendar`, `BookingSource`,
  `pushBookingEvent`, `removeBookingEvent`, `syncBusy`, `toBusyBlock`, `diffBusy`, `renewSubscriptions`,
  `handleNotification`, `signClientState`, `withAccessToken`, `authenticate`, `defer`, `hostTimeZone` and the
  record names. Rename in all files or none.
- **Two kinds of name.** The domain words the host renames are the rename table in `data-model.md`. The
  identifiers above are the authoring contract of this repository.
- **Never change `bookingMarkerId`, `busyBlockId` or `calendarRowId` inputs.** Every consumer's existing
  events, blocks and rows are addressed by them.
- **The non-negotiables are never presented as optional.** The five hard rules in `SKILL.md` and the five
  non-negotiables in `README.md` are one list, in one order. Changing one is a MAJOR release.
- **Measured and vendor numbers are load-bearing.** `clientState` at most 128 characters and ours at 80, the
  3-second notification deadline, 10,080-minute Outlook subscriptions and our six-day ones, the 180-day
  horizon, the 12-hour cron, the 48-hour renewal window, the ten-minute OAuth state, the 5-minute early
  refresh. Change one, change every file that states it.
- **Do not remove the odd-looking parts.** Each is a ledger entry in `provenance.md`, most under "Kept
  deliberately".
- **Mark additions as additions.** Anything designed in the skill and never run in the earlier implementation
  belongs in the "Designed here, not yet run in production" section of `provenance.md`.
- **Evals are not skill content.** A new prompt or an eval result is committed as `chore(evals): ...`, never
  causes a version bump and never rides in a release commit. After each eval, the README's eval badges point
  at the newest prompt 1 result per agent, updated in the commit that adds the run's notes.
- **Origin wording.** Call the codebase the audit ran against "the earlier implementation", and follow
  section 2 of the standard: none of its banned words, no denials, no corpus figures, and no defect counts or
  failure stories on the front door (README intro, `SKILL.md`, the description, this file).
- **Plain punctuation.** No em-dashes, en-dashes, arrows in prose, middle dots, ellipsis characters or smart
  quotes anywhere in this repository's markdown; diagrams are plain ASCII. `LC_ALL=C grep -rn '[^ -~	]'
  --include='*.md' .` must print nothing. Prose wraps at 110 columns; table rows and commands stay on one
  line.
- **Claims are verifiable.** A changed factual claim says how it was verified. Never from memory.
- **Releases** follow section 9 of the standard: SemVer, a `chore(release): X.Y.Z` commit holding only
  `CHANGELOG.md` and the README version line, tag `vX.Y.Z`, a GitHub Release with that changelog section.
  Commits follow Conventional Commits, and no file or commit message names a tool or a model as author.
