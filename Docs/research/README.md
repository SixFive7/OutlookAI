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
| `q149-draft-text/` | Q149 (d): should OutlookAI keep writing draft text through Word in a hidden compose window, with a lifetime pin while drafting (b), or write it as HTML with no window and no pin (c)? The merged measurement plan, its decision rules and its results, and the recovered 2026-07-31 probe D49 rests on, with the corrections to how D49 recorded it. | The choice between (b) and (c), which is the maintainer's. |
| `outlook-start-hang-2026-11-15.md` | Why did Outlook not start on 2026-11-15, and what did OutlookAI's server have to do with it? The timeline, the deadlock as far as it was established, what was ruled out, and what the settle wait costs. | Q175 to Q182: the settle wait and the start lock, the 120 s limit, the hang report and the start log, a cancel that keeps its deadline, a retry that only attaches, and no warm-up at Outlook start. |
| `q118-pin-measure/` | The raw logs of the Q118 multi-session pin measurement, 2026-10-03. | The pin rules summarised in `McpServer/Docs/com-host.md`. |

**Not included, deliberately:** the licence-state research and the reduced-functionality
measurement. Their conclusions are fully absorbed into `Testbed/MEDIA.md`, including the detection
query and its traps, so the working notes add nothing a reader would come back for.
