# Project Instructions

## Subfolder instructions

Before working in a subfolder, read any AGENTS.md from that folder up to the repo root that you haven't seen yet. Claude Code attaches them only when a file there is read: https://github.com/anthropics/claude-code/tree/main/mods/agents-md#where-it-still-differs-from-claudemd

## Changelog

When committing changes, **always update `CHANGELOG.md`** under the `## Unreleased` section.

Rules:
- Write entries as **user-facing summaries**, not developer jargon. Describe what changed from the user's or project's perspective.
- Keep entries concise — one line per change, starting with a verb (Add, Fix, Remove, Update, Improve).
- Group related commits into a single entry when they are part of the same logical change.
- Do not include CI fixes, typo fixes, or internal refactoring unless they affect user-visible behavior.
- Never modify released sections (any `## v...` heading). Only add to `## Unreleased`.
- If the Unreleased section already has entries from earlier in the session, add to it rather than replacing it.

## TODO.md

**An item is either open or gone.** Decided by the maintainer 2026-10-03. When an item is done -
or decided, if it was a decision - delete it from `TODO.md`. Never tick it (`- [x]`), and never
leave a DONE, CLOSED or RESOLVED note behind. The record of what was done lives in git history,
the CHANGELOG and the runbook, not in the TODO list.

## Pushing

**Push every commit individually, as it is made.** Decided 2026-09-17, standing. Do not
accumulate a local backlog and push it in one go.

**One by one means the FIRST-PARENT MAINLINE, not `git rev-list` order.** A repository with
merges is a DAG, and the commits on a merged side branch are not fast-forwards from one
another - pushing them in `rev-list --reverse` order fails on the first side-branch commit
with *"the tip of your current branch is behind its remote counterpart"*. The sequence that
works is:

```
git rev-list --reverse --first-parent origin/master..master
git push origin <sha>:refs/heads/master      # for each, in that order
```

Each step advances `master` by exactly one mainline commit, and a merge carries its whole
branch with it. Measured 2026-09-17 on a 16-commit backlog: 12 mainline steps, all clean.

**Check CI after every push, and treat a failed run as a bug.** Decided by the maintainer
2026-10-03 (Q95). Once the last push of a batch is done, read the runs it triggered
(`gh run list --limit 5`) when they finish, and fix or report any failure before moving on - a
local green suite does not stand in for it. The `McpServer` workflow last passed on 2026-09-15
and failed on every push checked from 2026-09-24 to 2026-09-27, on a test that only passed outside
UTC, and nobody looked: every merge had been verified locally, never on the runner.

## Build and Release

- **Build** runs automatically on every pull request (`.github/workflows/build.yml`), and can be triggered on demand. It only compiles — no releases, no tags, no changelog changes. On pull requests it also runs a dependency review.
- **Release** is triggered on demand (`.github/workflows/release.yml`) via `gh workflow run release`. It extracts the Unreleased changelog section, builds, creates an installer, publishes a GitHub Release, and stamps the changelog.
- The release workflow **fails if the Unreleased section is empty** — you must have release notes before creating a release.
- Version is derived from the latest GitHub release tag (base version) + commit count. No hardcoded version in the repo.
- The release workflow requires a `version_bump` input in `major.minor.patch` format (e.g. `1.0.0` for major bump, `0.1.0` for minor, `0.0.1` for patch). This input is **required** — the workflow will not run without it. `0.0.0` is rejected — every release must bump at least one version component.
- **After committing, ALWAYS ask the user if they want to create a release.** If yes:
  0. **Ask whether he has run `Docs/release-manual-checks.md` on this release candidate** (Q74 D2,
     decided on his behalf 2026-10-03 - see the overnight review). Those checks write to his real
     mailbox, so he runs them himself and an agent never does; if he has not, say in one line what they
     cover and let him decide whether to release anyway.
  1. **ALWAYS ask the version bump question.** Get the current version from the latest release tag via `gh release view --json tagName -q .tagName` and present options in A/B/C format showing current → new version. Example with latest tag v2.1.0.103:
     - A) Patch — 2.1.0 → 2.1.1
     - B) Minor — 2.1.0 → 2.2.0
     - C) Major — 2.1.0 → 3.0.0
  2. Run: `gh workflow run release -f version_bump=X.X.X` with the user's chosen bump value.
  3. Monitor with `gh run watch`.
- After a release, pull the stamped changelog commit before continuing work: `git pull --rebase`.

## MCP Server (`McpServer/`)

- `McpServer/` holds the MCP server projects (`OutlookAI.Core`, `OutlookAI.McpServer`, `OutlookAI.McpServer.Tests`). Build them with `dotnet build` **by explicit csproj path** — never via `OutlookAI.slnx`, which only contains the VSTO add-in (MSBuild-only).
- Their CI is `.github/workflows/mcpserver.yml` (windows runner, dotnet only; runs `dotnet test --filter "Category!=Live"`). Tests marked `Category=Live` need the real Windows Search index plus Outlook and only run on a configured dev machine.
- Developer documentation: `McpServer/README.md`.

## Dependencies

**No external applications and no licensed components. Ever.** Decided 2026-09-15, standing.

This is not a preference about tidiness — it is a hard constraint on every design decision, and
it has already excluded an otherwise-ideal answer. When the question "how do we create an Outlook
POP3 account programmatically" was researched, the only off-the-shelf component that can do it
(Redemption, $299.99 distributable / $899.99 with the profile library) was ruled out on this rule
alone. Its distributable tier also excludes open-source projects, and this repository is public.

**What this permits.** Anything that ships with Windows or with the .NET Framework already
present on the machine: Extended MAPI through the stub `mapi32.dll`, `System.Windows.Automation`
(UIAutomation), `Add-Type` (the Framework's own `csc.exe`, which needs no SDK), the Office
Deployment Tool, `oscdimg` from the Windows ADK. None of these is an install the project owns.

**What this forbids.** Purchased libraries, third-party COM components, anything requiring a
licence key, and anything a rebuilder would have to download and install beyond the media
`Testbed/MEDIA.md` already names as preconditions.

**Do not re-litigate this per task.** If a route appears blocked without a paid component, the
answer is to question the requirement, not the rule — see `Testbed/README.md` and `Docs/research/profile-automation-research.md` for how the POP3 account
question was reframed rather than bought.

**One exception, decided by the maintainer 2026-09-24 (Q71): a loopback mail sink for the test
VMs.** A ready-made open-source mail server is permitted for this one job, so that the live tests
that send mail can run on the VMs. Conditions, all of them: free and open source under a
permissive licence; no licence key, account or telemetry; staged offline as media in
`Testbed/MEDIA.md` and pinned by a hash its own maintainers publish; installed only on the test
guests, never on the maintainer's workstation. `Testbed/MEDIA.md` names the tool and version.
This is the only exception; it does not generalise to "open source is fine".

## The add-in on the maintainer's workstation (Q81)

**Outlook on the maintainer's workstation loads the add-in only from a folder no build writes
into.** Decided by the maintainer 2026-09-27 (Q81), in his words: *"I want to be able to ask you to
put the dev build on my machine if I want to test it for a release. I do not want it on my machine
whilst agents are still actively building it."* Until then every plain build of `OutlookAI.csproj`
registered itself with the Outlook on the machine that built it, so his Outlook quietly loaded
whichever build folder had been built last.

1. **Agents never register or install the add-in on the maintainer's workstation** — not with the
   installer, not with a hand-written registry value, not with a build. The one exception: he asks
   for it **in the current message**, and then only through `Tools/Switch-AddInBuild.ps1` —
   `-Commit <rev> -Execute` puts a dev build on, `-Restore -Execute` puts the installed release
   back, `-Status` says what Outlook will load. A request in an earlier message, from another
   agent or in a plan is not a request.
2. **Agents build the add-in only through `Testbed/host/Publish-AddInPayload.ps1` or
   `Tools/Switch-AddInBuild.ps1`** (`-BuildOnly` when a build is all you need). Both build a commit
   from a `git archive`, stand in the VSTO tasks that write the registry, and prove the host
   unchanged afterwards. `OutlookAI.csproj` no longer registers anything when built outside Visual
   Studio, but a commit from before that change still does — so no plain `msbuild` of the add-in,
   and never `BuildingInsideVisualStudio=true` on a command line, which turns registration back on.
3. **Outlook on the workstation never loads the add-in from a folder any build writes into** — not
   `bin\Debug` or `bin\Release`, not a worktree, not `.work`. A dev build is copied to
   `%LOCALAPPDATA%\OutlookAI\DevBuilds\<folder>` first, and only that copy is registered. If
   `-Status` shows Outlook pointed at a build folder, say so and offer `-Restore`; do not change it
   unasked.

Visual Studio is the maintainer's own tool and keeps registering: F5 builds `bin\Debug`, points
Outlook at it, and leaves it pointed there after the debugging session. `-Status` shows that, and
`-Restore` puts the release back. The script never touches the MCP server or its disabled
executable, never runs elevated, and never starts, quits or kills Outlook — a change takes effect at
the next Outlook restart.

## Mailbox Safety (MANDATORY — live tests touch REAL mailboxes)

**THE MAINTAINER'S WORKSTATION IS READ-ONLY FOR LIVE TESTS — ALWAYS.** Decided by the maintainer
2026-09-24 (Q69, Q72), in his words: *"run read-only and always only read-only!"* Every live test
that can run on the test VMs runs **only** there. The only live tests that may run on the
workstation are the fundamentally immovable ones — those that need Exchange (delegate and shared
mailboxes, cached mode), which no test VM can have under the Dependencies rule — and they run
**read-only**. **Never run a write-capable live test on the workstation, and never select a
workstation run by a filter that could include one.** Since Q74 (2026-10-03) code enforces this:
the workstation runs only the derived filter in `Testbed/README.md` section 4d - live tests that need
Exchange AND carry `Writes=None`, which T1 proves read-only from the compiled code - its settings'
profile makes every in-process write throw, the test hub included, and the test-side MCP client
refuses every tool not classified read-only. Those gates are a floor, not a licence: never edit the
workstation's settings file, never re-declare it `Portable`, and if you cannot show a workstation
run is read-only, do not start it. The rules below still bind every live run, on the workstation
and on the VMs alike.

`Category=Live` tests run against the developer's **real production Outlook profile**: real mail accounts plus delegate/shared mailboxes **to which the profile has full write access**. Treat every live run as an operation on production data. A past incident mass-deleted real mail (fully recovered) because an agent improvised a cleanup script — these rules exist so that never repeats. They are non-negotiable and apply to every agent, every session, whether or not live tests are the task:

1. **Never mutate mailbox items from ad-hoc shell code.** No PowerShell, no raw COM one-liners, no throwaway scripts. Creating, deleting, moving or editing an item happens **only** through the project's tested helper code (`LiveOutlookTestMailer` and the live fixtures) or the shipped MCP tools. If cleanup needs something the helpers cannot do, extend the helpers with tests — do not improvise.
2. **Never pattern-match subjects shell-side.** PowerShell's `-like "*[tag]*"` treats `[...]` as a character-class wildcard, so it matches nearly every subject — that is exactly how real mail was destroyed. Deletion selection is **EntryID allowlist AND ordinal tag match, both required**; every test-created item carries a subject tag matched **ordinally** — `[OutlookAI-McpTest]` for live-tier artifacts, `[OutlookAI-Corpus]` for measurement-corpus items. **The two are deliberately different strings so that an artifact sweep can never select a corpus item**, and `T1/CorpusTagSeparationTests` enforces that they stay different (not merely unequal — neither may contain the other).
3. **Writes only in the designated test mailbox** (see the gitignored live-test settings). Every other account and **all** delegate/shared mailboxes are read-only for tests — no exceptions. The `StoreWriteAllowlist` guard enforces this in code: a write aimed anywhere else throws instead of running. Logs and failure messages never print other stores' subjects or bodies.
4. **Every live run must end with zero tagged artifacts**, proven by the post-run sweep — which covers Drafts, Inbox, Sent Items, **Outbox**, Deleted Items and the **Sync Issues subtree** (Conflicts / Local Failures / Server Failures), plus test folders removed deepest-first.
5. **A live run may not lose mail anywhere.** The per-store count tripwire snapshots every store's mail folders before and after; any item-count **decrease**, or any folder added/removed, outside the test mailbox fails the suite loudly. No snapshot ⇒ the live tier refuses to run.
6. **Signatures are user data.** Tests may only create/update/delete signatures prefixed `OutlookAI-McpTest-`; the `SignatureDirectorySnapshot` guard (SHA-256 before/after) must run and the suite must leave the user's real signatures bit-identical. `manage_signature` tests restore any registry defaults they touch.
7. **Outlook lifecycle:** never `taskkill` OUTLOOK.EXE. Graceful `Application.Quit()` only when no unsent compose windows are open and the Outbox is empty — and release COM references BEFORE quitting (quitting while refs are held zombifies the process). Prefer leaving Outlook headless.
8. **Run live tests only via the suite**, and only as `Testbed/README.md` section 4c describes: on a test guest, through `guest/Register-InteractiveTask.ps1`, with the per-run opt-in it gives - or, on the maintainer's workstation, only the read-only run section 4d describes. Every live test refuses to start without that opt-in (decided 2026-09-24): an accidental name filter once selected live tests, and in a checkout holding a real settings file it would have run them. **Never set the opt-in on the maintainer's workstation to run a test that can write** - see the paragraph above these rules. The suite's fixtures enforce the snapshots, allowlists, tripwire and zero-artifact sweeps. Never perform mailbox operations outside it during testing.
9. If a gitignored `v3.MD` exists at the repo root, read its §0 safety envelope before any live-test or mailbox-touching work — it is the authoritative, more detailed contract.
