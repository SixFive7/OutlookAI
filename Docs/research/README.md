# Research notes

**Why these are tracked rather than left in scratch.** Each of these answered a question that a
decision now rests on, and each was written in `.work/`, which is gitignored. That is not a safe
home: `Testbed/README.md` opens by recording that *"three times in one week, material this project
depended on was kept only in a scratch directory on one machine, and that directory was cleared
without warning."* It happened again on 2026-09-15 — `profile-automation-research.md` lived in an
agent worktree and was recovered by chance minutes before the worktree would have been cleaned up.

**What makes them worth keeping, as opposed to folding into the docs.** Every claim carries a
label — `[MS-DOC]`, `[COMMUNITY]`, `[MEASURED]`, `[INFERRED]` — with its source. The conclusions
are already in the runbook and `MEDIA.md`; what only lives here is **which claims are documented,
which are inferred, and what could not be established at all**. When an Office build changes and
something stops working, that is the part you need, and it is the part a summary always loses.

**They are snapshots, not maintained documents.** Each is true as of its date and about the builds
it names. Do not edit them to stay current — if a finding is superseded, record that where the
decision lives and leave the note as the historical record of what was known when.

| File | Question it answered | Decision resting on it |
| --- | --- | --- |
| `com-release-audit.md` | Does the shipped code release Outlook COM references on exit, and can killing the COM host poison a shared Outlook? | Making the clean-exit path reachable. The audit's central premise was later **measured false** — a killed holder does not poison the instance — so read it together with the kill-path measurement in `Docs/autonomous-session-log.md`. |
| `pop3-account-routes.md` | Three routes to creating a POP3 account without a paid component: PRF, requirement analysis, UI Automation. | That the requirement is far smaller than assumed: one account of any type reaches 26 of 34 blocked tests, and exactly **one** test needs mail on the wire. |
| `profile-automation-research.md` | Can Outlook profiles, PST stores and per-account signatures be created programmatically? | Extended MAPI for profiles and stores, with `PR_DISPLAY_NAME` set at creation. |

**Not included, deliberately:** the licence-state research and the reduced-functionality
measurement. Their conclusions are fully absorbed into `Testbed/MEDIA.md`, including the detection
query and its traps, so the working notes add nothing a reader would come back for.
