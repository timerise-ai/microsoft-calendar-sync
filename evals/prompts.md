---
prompts:
  - prompt: "Let hosts in this Next.js app on Supabase connect their Microsoft 365 or Outlook calendar: the times they are busy in Outlook must stop being bookable, and every confirmed booking should show up in their Outlook calendar and disappear again when it is cancelled."
    stack: Supabase
  - prompt: "Add Outlook calendar sync to our Firestore backend, and put each booking into the host's calendar as a Teams meeting with the guest invited."
    stack: Firestore
  - prompt: "Our Outlook calendar integration stops working about an hour after a host connects, and their vacations never block anything. Find out why and fix it."
---

# Prompts

What an operator types after installing this skill, in their own words. An agent eval installs the skill
into an empty Next.js app, gives the agent one of these prompts and no further help, then type-checks, builds
and tests the result; the first prompt runs before every release. The results are the other files in this
folder. Section 10 of [STANDARD.md](https://github.com/timerise-ai/skills/blob/main/STANDARD.md) says how a
run is made. The prompts and the newest runs are on
[the skill's page](https://timerise.ai/skills/microsoft-calendar-sync) on timerise.ai.
