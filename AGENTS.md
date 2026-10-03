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

## Build and Release

**There is no CI.** Decided by the maintainer 2026-10-03, in his words: *"Remove the git CI
pipeline. I want only the build and test suite to run on my pc. No CI pipelines on github."*
Nothing runs on GitHub: no build on a pull request, no CodeQL scan, no dependency review, no
Dependabot version updates, no release workflow. Tests and self-tests run on the build VM (the
section after next); builds and the four guards run on this workstation; and a release is
`Tools/Publish-Release.ps1`, run here.

**The release script** is the old release workflow, ported step for step, with the gates a
workstation needs added. `pwsh -File Tools/Publish-Release.ps1 -VersionBump X.Y.Z` is a DRY RUN -
everything except publishing; `-Execute` publishes. In order, it:

1. refuses unless the working tree is clean and HEAD is `origin/master`'s tip (a dry run only
   notes the second);
2. derives the version: the latest GitHub release tag's base plus the bump, and HEAD's commit count
   plus one - the stamp commit - as the fourth part. No hardcoded version in the repo. The bump is
   **required**, in `major.minor.patch` form (`1.0.0` major, `0.1.0` minor, `0.0.1` patch), and
   `0.0.0` is rejected - every release bumps at least one component;
3. takes the release notes from the CHANGELOG's `## Unreleased` section - and **refuses if it is
   empty**: you must have release notes before creating a release. It also refuses notes longer than
   the 125,000 characters GitHub accepts as a release body (a dry run only notes that);
4. checks that the certificate `OutlookAI.csproj` pins is in `Cert:\CurrentUser\My` with its private
   key and not expired;
5. runs the four guards under `pwsh` and `powershell.exe`, then `check-pinned-constants.ps1` against
   that certificate;
6. runs **D7 (c)** (below), and refuses to release if either comparison fails;
7. builds through `Testbed/host/Publish-AddInPayload.ps1 -ReleaseSigningThumbprint` - so the Q81
   guards hold for a release too - with the MCP server and the VSTO runtime in the installer;
8. signs the installer by thumbprint with an RFC 3161 timestamp, reads the signature back, and
   refuses an installer over the updater's 50 MB cap;
9. runs the whole non-live suite and every self-test of HEAD on the build VM - anything but exit 0
   refuses;
10. makes the stamp commit - `## Unreleased`, then `## v<version> - <date>` - with git plumbing, so
    neither the working tree nor any branch moves;
11. with `-Execute` only: pushes that one commit to master (a fast-forward), runs `gh release create`
    with the signed installer, and fast-forwards a local master that sat on the released commit.

Everything lands in `.work\release\v<version>\`; `release.json` there is the record. It needs
Visual Studio with the Office workload, Inno Setup 6, the Windows SDK's signtool, the .NET 10 SDK,
`gh` logged in, and the staged VSTO runtime (`Testbed/MEDIA.md`). Run it in the background or with a
timeout of 30 minutes or more: the build VM's run is inside it.

**D7 (c): the one piece of a self-test that runs on this workstation.** Decided by the maintainer
2026-10-03. The self-tests of `Testbed/host/Publish-AddInPayload.ps1` and
`Tools/Switch-AddInBuild.ps1` compare their stand-ins for Visual Studio's VSTO build tasks with
Visual Studio's real targets file, and on the build VM, which has no Visual Studio, they skip that
comparison. The stand-ins are what keep a build from registering the add-in in his Outlook (Q81),
and a Visual Studio update could silently reopen that hole. So before every release the release
script runs both comparisons - each script's read-only `-CompareInstalledTargets` - and refuses to
release if either fails or finds no Visual Studio. That comparison is the whole exception: never
run either script's `-SelfTest` here.

- **After committing, ALWAYS ask the user if they want to create a release.** If yes:
  0. **Ask whether he has run `Docs/release-manual-checks.md` on this release candidate** (Q74 D2,
     decided on his behalf 2026-10-03 - see the overnight review). Those checks write to his real
     mailbox, so he runs them himself and an agent never does; if he has not, say in one line what they
     cover and let him decide whether to release anyway.
  1. **ALWAYS ask the version bump question.** Get the current version from the latest release tag via `gh release view --json tagName -q .tagName` and present options in A/B/C format showing current → new version. Example with latest tag v2.1.0.103:
     - A) Patch — 2.1.0 → 2.1.1
     - B) Minor — 2.1.0 → 2.2.0
     - C) Major — 2.1.0 → 3.0.0
  2. Run `pwsh -File Tools/Publish-Release.ps1 -VersionBump X.X.X -Execute` with the user's chosen
     bump, from a clean checkout of master with everything pushed. Without `-Execute` the same
     command is a dry run, for when he wants to see it first.
  3. Read its verdict: exit 0 and `RELEASED v<version>`, then `gh release view v<version>`. Any
     refusal names the step and its log under `.work\release\v<version>\logs\`.
- After a release, a local master that sat on the released commit has been fast-forwarded by the
  script; anywhere else, pull the stamped changelog commit before continuing work: `git pull --rebase`.

## MCP Server (`McpServer/`)

- `McpServer/` holds the MCP server projects (`OutlookAI.Core`, `OutlookAI.McpServer`, `OutlookAI.McpServer.Tests`). Build them with `dotnet build` **by explicit csproj path** — never via `OutlookAI.slnx`, which only contains the VSTO add-in (MSBuild-only).
- The non-live suite (`dotnet test --filter "Category!=Live"`) runs only on the build VM - next section. Tests marked `Category=Live` need Outlook and a mailbox: they run on the test VMs, and on this workstation only the Exchange-only read-only subset (Q74, Mailbox Safety below).
- Developer documentation: `McpServer/README.md`.

## Tests run on the build VM, never on this workstation (Q94, Q102)

**Non-live tests and script self-tests run only through `Testbed/host/Invoke-TestsOnBuildVm.ps1`,
on the build VM `OutlookAI-Build` - never on the maintainer's workstation.** Decided by the
maintainer 2026-10-03, in his words: *"Move everything (except for the exchange tests because I
do not have exchange in the vms) to the VMs"* (Q94), onto a small dedicated build-and-test VM
(Q102 (b)). Proven the same day: master's non-live suite - 3,346 tests at `3e7b861`, every one
passing - and every self-test, in under four minutes (`Testbed/README.md` section 1c).

```
pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1                    # this checkout's HEAD: the whole non-live suite and every -SelfTest
pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 <commit-or-branch>  # any revision: a merge to verify, another agent's branch
pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -Filter 'FullyQualifiedName~T1.SomeTests' -SkipSelfTests
pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -SkipSuite -SelfTestInclude 'Testbed/guest/*'
```

- **It tests a commit.** Uncommitted changes are not in the run, which says how many it left out:
  commit first - a work-in-progress commit is fine.
- **The exit code is the verdict:** 0 pass; 1 a test or self-test failed, the suite hung, or the
  filter selected nothing; 2 the revision does not build; 3 not tested (the VM, the lock, a time
  limit); 4 refused. Results land in `.work\build-vm-runs\<run>\` of the checkout it ran from:
  `summary.txt`, `summary.json`, `vm\trx\suite.trx` and every log.
- **One run at a time; callers queue.** Run it in the background or with a timeout of 20 minutes
  or more. Never use the VM by hand while runs may happen; to hold runs off, take a lease on it
  with `Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Build` and release it after.

**What stays on this workstation - this, and nothing else:**

- the four static guards, `Tools/Checks/check-*.ps1`, which only read files - run them under
  both `powershell.exe` and `pwsh`;
- builds - `dotnet build` by csproj path, and the add-in build, which needs Visual Studio,
  through the two scripts "The add-in on the maintainer's workstation (Q81)" below names;
- releases - `Tools/Publish-Release.ps1` (Build and Release, above), which builds here and tests on
  the build VM, and runs D7 (c)'s two read-only `-CompareInstalledTargets` comparisons here;
- the Exchange-only read-only live tests (Q74) - the derived filter of `Testbed/README.md`
  section 4d, under Mailbox Safety below.

So no `dotnet test` runs a test on the workstation except that last live run, and no script's
`-SelfTest` runs here at all: the runner finds and runs every one of them on the VM. D7 (c) runs
one comparison out of two self-tests here, and only that, before a release.
`dotnet test --list-tests`, which builds and discovers and executes no test, stays usable here -
it is how a workstation live run's selection is checked before it starts.

## Test VMs stay saved unless in use

**Every test VM is saved to disk whenever nothing is using it - never left running, never
paused.** Decided by the maintainer 2026-10-03, in his words: *"Keep them off (and saved to disk
not standby in ram) if you do not need them. Make sure they are off when we end this session or
I will forget and lose performance the coming months."*

- Take a lease (`Testbed/host/Set-TestbedLease.ps1`) before starting or using a VM, and hold it
  for as long as you use it.
- `Save-VM` it when more than about ten minutes without guest work lie ahead, and when your task
  ends - then release the lease.
- Before a session ends, every test VM is saved: `Get-VM OutlookAI-*` shows none running unless
  a live lease says something is still using it.
- Two backstops, not a licence to leave one running: the `OutlookAI-TestbedIdleSave` task saves
  any running testbed VM that has no live lease and has been up ten minutes, every fifteen
  minutes (`Testbed/README.md` section 5b); and no test VM starts with the host
  (`AutomaticStartAction Nothing`, `AutomaticStopAction Save`).
- **No scheduled task on this workstation may start `powershell.exe`, `pwsh.exe` or any other
  console program directly.** Each run opens a console window - a Windows Terminal window here -
  that takes the maintainer's keyboard focus; the idle-save task did that every 15 minutes until
  2026-10-03. Start it through a GUI-subsystem launcher instead, as
  `Testbed/host/Invoke-TestbedIdleSave.vbs` does (`wscript.exe`, the window hidden from creation).
  The same goes for any background process a script starts on this workstation: hide it AT
  CREATION (`Start-Process -WindowStyle Hidden`, `-NoNewWindow`, `CREATE_NO_WINDOW`, or `SW_HIDE`
  in the startup info, as the build-VM runner's janitor does) - never `powershell.exe
  -WindowStyle Hidden`, which hides a console window only after it has already taken focus.

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
   `Tools/Switch-AddInBuild.ps1`** (`-BuildOnly` when a build is all you need). A release builds
   through the first: `Tools/Publish-Release.ps1` calls it with `-ReleaseSigningThumbprint`. Both build a commit
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
Exchange AND carry `Writes=Nothing`, which T1 proves read-only from the compiled code - its settings'
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
