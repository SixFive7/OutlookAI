# Overnight 2026-10-03 - decisions taken on the maintainer's behalf, and deviations

**For review.** At 01:59Z on 2026-10-03 the maintainer handed the work over for the night:
continue autonomously, answer any question with the option that would be recommended, and record
every such answer and every deviation from the plan here. If an answer below differs from the one
he would have given, that item is backtracked; nothing else is.

Each **decision** names the question, the options, what was chosen, why, and how to undo it.
Each **deviation** names what the plan said, what happened instead, and why.

## The plan as it stood at hand-over

- **Q74** - enforce a read-only workstation: layers 1+2+3, plus A1, B1, C1 with C3, D1 + D2.
- **Q86 + Q93** - unit tests stop writing the real audit log (option 1, concurrency-safe), the
  existing log is renamed untouched afterwards (b), then a read-only MCP tool reads the log.
- **Q94 + Q102** - every test except the Exchange-only live tests moves off the workstation, onto
  a new small build-and-test VM (`OutlookAI-Build`).
- **Q95** - the time-zone test fixed; CI checked after every push (now in `AGENTS.md`).
- **Q96** - the four loose ends of folder-creation reporting.
- **Q98** - populations built now without the undated items (a); the index measurement (f) later.
- **Q99** - store-to-index matching: implement if a correct, deterministic, not-too-complex method
  exists; otherwise come back with options.
- **Q100** - split the add-in installer: install with admin rights, first Outlook start without.
- **Q101** - the remaining live tests that could pass without checking.
- Then: the first live-tier run on guest two, guest one's populations and live run, and the report
  on what cannot move off the workstation.

## Decisions taken on the maintainer's behalf

### D1 - Q101: the "no index" return in `SearchAlwaysAnswers` goes through the shared helper too
- **Question.** Q101 narrowed `SearchAlwaysAnswers`' early return to the one legitimate case (the
  index is unreachable). Should that remaining return pass silently, or say so?
- **Chosen.** Through the shared Production/Portable helper: it FAILS on the workstation profile and
  prints `PROVED NOTHING:` on the VMs. Same pattern as every other Q101 item.
- **Alternative.** A plain return, as before (silent).
- **Undo.** One call in `T3/OutlookAvailabilityLiveTests.cs`.

### D2 - Q95: no CHANGELOG entry for the time-zone fix
- **Why.** `AGENTS.md` excludes CI fixes and changes with no user-visible behaviour; the product's
  conversion is unchanged (verified across every Windows time zone).
- **Undo.** Add a line under Unreleased.

### D3 - Q102: no Office on the build-and-test VM
- **Question.** Does `OutlookAI-Build` need Office installed to run the non-live tests?
- **Evidence.** None of the 3,005 non-live tests (150 classes) needs Office or Outlook installed or
  running: GitHub CI's `windows-latest` has no Office and passed 3,004 of 3,005, the one failure
  being the time-zone test since fixed (Q95). About 200 tests take an "Outlook is not running"
  branch and a handful a "no Office" branch; all pass either way.
- **Chosen.** (a) No Office. **Alternatives:** (b) Office installed but never started - a licence
  and an Outlook the dormant attach paths could reach; (c) none, plus an occasional run on a test
  guest to cover the "Office present" branches; (d) none, plus code hooks so both branches are
  tested everywhere - the later fix if those branches matter.
- **Undo.** Add the existing Office Deployment Tool step to the build VM.

### D4 - Q102: the build VM's time zone is W. Europe Standard Time
- Matches the workstation and the test guests; GitHub CI keeps covering UTC.
  **Alternatives:** UTC; an odd-offset zone; both on every run.

### D5 - Q102: source reaches the build VM as `git archive` plus an empty `.git` marker
- No git on the VM - it would be new media under the Dependencies rule. One self-test
  (`New-LiveTestSettings.ps1`) only checks that a `.git` exists, hence the marker. The four
  `.github/scripts` checks need a real clone (`git ls-files`), so they stay on the workstation and
  in GitHub CI; they read files only and touch nothing of Outlook.
  **Alternative:** stage git as new media and clone from a bundle.

### D6 - Q102: Windows PowerShell 5.1 only on the build VM
- Built in, and the only PowerShell the test guests have (Q78). GitHub CI still runs the checks
  under PowerShell 7. **Alternative:** stage PowerShell 7 too (new media, a Dependencies question).

### D7 - Q102: the Visual Studio targets comparison in two self-tests is skipped on the VM
- `Publish-AddInPayload.ps1` and `Tools/Switch-AddInBuild.ps1` compare their build stand-ins with
  Visual Studio's real VSTO targets only where Visual Studio is installed. On the VM that check
  skips; nothing else covers it except a run on a machine with Visual Studio.
  **Alternatives:** install Visual Studio with the Office workload on the VM; or keep running those
  two read-only self-tests on the workstation.

## Deviations from the plan

### V1 - Q96 started after one agent finished, not two
The plan said Q96 and Q100 would start when two running agents had finished. Q96 started when the
first (Q95 + Q101) finished, because that agent's subject was done and nothing overlapped. Q100
still waits for the next one.

### V2 - The TODO rule was applied to the whole file, not only to the item that prompted it
The maintainer's rule ("items are either open or removed") came with one item. It was applied to
all of `TODO.md`: 38 ticked items, the finished MCP-server restore, and the decided undated-rows
question were removed (2,597 to about 1,090 lines, 62 open items intact). Six references in other
documents that pointed at removed items were re-pointed, and one of them - the MCP server README -
turned out to describe a sweep-cache bug fixed on 2026-08-24 as current; that was corrected too.

## Notes (no decision needed)

- **Q94's research, for the record.** Four non-live test classes wrote the real audit log
  (`AtomicityClaimsTests`, `DraftUpdateReentrancyTests`, one test in `DraftValidationTests`, and
  `SoakBatchCCiToolShapeTests` through the server) - fixed by Q86 and moot once tests leave the
  workstation. `WritingRulesGateCiTests` printed the workstation user's own writing rules into test
  output, and a signature test listed every signature name - harmless on a VM. Several test paths
  hold a real in-process COM gateway behind a guard (`DraftValidation`, `Phase3Validation`,
  `SendValidation`, two tripwire tests, `ComHostSupervisionCiTests`): on the workstation a broken
  guard there could attach to or start the real Outlook; on a VM with no Office it fails loudly.
  Under the maintainer's first Q94 rule, 40 classes (761 tests) would have moved and 110 stayed.

- **Q101 consequence.** On the workstation profile, a healthy Outlook now always fails the
  transient-state availability test - that is what "fail on Production" means for it. With every
  non-Exchange test moving to the VMs it should not run there at all.
- **Risk to watch at the first live run.** Q101's Outbox/Inspector checks now run before the
  window branch, so a hidden Inspector left behind by the headless test's `display:false` draft
  would now fail `LiveDisconnectRecoveryTests` on a guest.
