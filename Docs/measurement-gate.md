# The pre-release measurement gate

**What this is.** A release gate that reads every performance and coverage number this project
produces, compares each one against *this machine's own history*, and refuses the release when
anything has moved beyond tolerance — **in either direction**.

**Why it exists.** Every timing constant in this codebase was derived from a measurement taken
once, on one machine, under conditions nobody wrote down. That has already gone wrong here:
`MailService.SweepBudgetMs` was set to 180 s from a measurement taken while `Table.Sort` was
silently failing, so the number described *broken* behaviour and nothing recorded that. The same
shape produced `comHost.largestFrameBytes` at 432 KB — read as 152× headroom, actually bounded by
a timeout, and later measured at 10,734,599 bytes on the VM corpus. A number whose bound is
another defect reads exactly like a healthy number.

This is the mechanism that stops the third one. `Docs/vm-coverage-analysis.md` §6.4 argued for
it (direction C); §7 Question 5 asked who reads the result. Both halves are built here.

---

## The three rules it is built on

**1. Movement in either direction counts.** A sweep that got 40 % *faster* is not good news by
default. In this codebase the usual cause is that something stopped being done: a sort that
started failing, a folder set that shrank, a date window that stopped matching, a cache that
started answering. So the default verdict class is symmetric and a faster number fails exactly
as a slower one does. The report names the direction and does not congratulate.

**2. The bias is to fail.** The maintainer's words: *"fail aggressively over leaking slow
performance degradations."* A borderline call fails. A metric that has history and is missing
from this run fails. A cold start does not silently pass — it exits `2` and has to be accepted by
hand.

**3. The numbers never reach the repository, and never reach release notes.** They are
statistics about one person's machine. They are meaningful *only* relative to older values from
that same machine, they are not representative of anyone else's system, and this repository is
public. This is a hard privacy requirement, and it is enforced structurally rather than by
discipline — see [Why the numbers are local-only](#why-the-numbers-are-local-only).

> **Note for anyone reading `Docs/vm-coverage-analysis.md`:** its Question 5 recommended putting
> the raw measurement table *into the release notes* so that somebody would read it. **That half
> is superseded.** The numbers may not be published. The role it was meant to play — making sure
> a human or agent actually looks at the table — is filled instead by printing the full table to
> the console on every release run, and by the reader step below.

---

## Where the baseline lives

```
%LOCALAPPDATA%\OutlookAI\Measurements\
    history.jsonl                    append-only, one JSON record per run   <- authoritative
    annotations.jsonl                append-only, one record per deliberate baseline move
    reports\<runId>.txt              the human-readable verdict for that run
    reports\<runId>.comparison.json  the same comparison, for the reader step
    README.txt                       written on first use: what this is, why it is not in the repo
```

On this machine that resolves to `C:\Users\jori\AppData\Local\OutlookAI\Measurements\`. The
directory is created on first use. Override with `-StoreRoot`, but see below — the gate will
refuse a great many paths on purpose.

**Format: JSON Lines.** One complete record per line, appended, never rewritten. A script can
diff it (`Get-Content history.jsonl | ConvertFrom-Json`), a human can read it, and a corrupt line
is one bad run rather than a lost history. The gate refuses to start on a line it cannot parse
rather than skipping it, because a record it cannot compare against is a baseline that has
quietly narrowed.

### What one record carries, and why the provenance is not optional

A number without its conditions is not comparable — that is the whole lesson of the 180 s
episode. Every record therefore carries:

| Field | Why |
| --- | --- |
| `runId`, `recordedAtUtc` | identity and ordering |
| `provenance.gitCommit`, `gitShort`, `gitBranch`, `gitDirty` | *what code* was measured. `gitDirty: true` raises a notice: the commit beside the numbers does not fully describe them |
| `provenance.machine`, `user`, `os`, `powershell`, `dotnetSdk` | *what machine*. A record from a different machine is not a baseline, it is noise |
| `conditions.profile` | `production` or `vm`. **Part of the comparison scope** — a VM number and a production number are never compared, because the same operation is 50×–90× faster on a local PST than on a delegate Exchange store |
| `conditions.indexed` | `indexed` / `unindexed` / `mixed`. **Decisive**: an indexed store takes the index path, an unindexed one takes the 7-day sweep fallback. A change here *fails* rather than being compared — the two runs measure different code |
| `conditions.corpusId` | which synthetic corpus, if any. Also part of the comparison scope |
| `conditions.corpusAnchor`, `corpusAgeDays` | the `--anchor` the corpus was built from, and how far past it the run is. See [The aged-corpus refusal](#the-aged-corpus-refusal) |
| `conditions.storeSet` | e.g. `5 stores / 159 folders / 2044 items`. Every elapsed figure is a rate over this. A change raises a notice (a failure under `-StrictConditions`) |
| `conditions.notes` | free text: warm or cold host, what was running, anything that would change the reading |
| `metrics[].conditions` | per-metric override, for numbers taken under conditions of their own — `comHost.largestFrameBytes` read *after* the sweep is not the same reading as one taken before it |

---

## Running it before a release

The release run is one command, and it needs the suite log and a live-run file:

```powershell
# 1. Standing verification, output kept so the gate can read the suite numbers.
dotnet build McpServer/OutlookAI.Core/OutlookAI.Core.csproj
dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj `
    --filter "Category!=Live&FullyQualifiedName!~Tests.T3." *> .work/test.log

# 2. Take the live measurements (see below) into a run file.
pwsh -File .github/scripts/measurement-gate.ps1 -Template > .work/live-run.json   # skeleton
#    ... fill it in from the live run ...

# 3. The gate.
pwsh -File .github/scripts/measurement-gate.ps1 `
    -Run .work/live-run.json `
    -Collect -TestLog .work/test.log `
    -ProfileKind production -Indexed indexed `
    -StoreSet "5 stores / 159 folders / 2044 items" `
    -Require All `
    -Label "pre-release 2.2.0"
```

`-Require All` is what makes it a *release* run: every catalogued metric must be present, so a
partial run cannot pass as a full one. Without it (`-Require Present`, the default) the gate
compares whatever it was given — useful mid-development, not sufficient before a release.

### What `-Collect` gathers on its own

No Outlook, no mailbox, no network. It reads the repository and the suite log:

- the **suite numbers** from `-TestLog` (passed / total / failed / skipped / duration);
- the **invariant counts** by running `check-pinned-constants.ps1` and parsing its summary;
- the **36 budget constants** straight out of the sources.

### What has to be measured by hand, and where each field comes from

The live numbers need Outlook and a real profile, so the gate ingests them from a JSON file
rather than taking them. `-Template` prints the skeleton. The mapping:

| Run-file key | Read from |
| --- | --- |
| `sweep.*` | the `search` payload's `sweep` object: `elapsedMs`, `itemsSeen`, `foldersSwept`, `sortRefusedFolders`, `itemCappedFolders`, `itemsBodyCapped`; the per-store breakdown for `perStore.*` |
| `scan.*` | the `search` payload's `exhaustive` object with `exhaustive: true`: `elapsedMs`, `foldersScanned` |
| `census.*` | the live tier's tripwire census line (`[tripwire] post-run census in N ms (...)`) and the `CensusIdentityPlan` counters |
| `index.*`, `search.indexQueryMs.median` | `outlook_health`'s `index.perStore[]` frontier ages, sampled repeatedly; the index query wall clock |
| `comHost.largestFrameBytes` | `outlook_health`'s `comHost.largestFrameBytes`, read **after** the sweep |
| `connect.*`, `move.*`, `transport.*` | the connect timings, the hub-only move batch, and the live inbox-arrival round trip |

**Three traps that will otherwise produce a wrong number that looks right** (all from
`Docs/corpus-measurement-plan.md` step 3, and all still true):

1. **`SweepCache` has a 10-second TTL.** Two identical searches inside 10 s give the second one
   `sweep.cached: true` and `elapsedMs: 0`. Vary the query or wait.
2. **A cold COM host gets a 90 s connect floor on top of the budget.** The first sweep after a
   host start is not comparable to the rest. Record it as `connect.coldSearchMs` and discard it
   from the sweep figures.
3. **Read the frame high-water after the sweep, never before.** That ordering is the difference
   between the 432 KB reading and a real one.

### The reader step — required, and not optional theatre

**Tolerances catch drift. They structurally cannot catch the stable-but-wrong shape**, because
that shape does not move: a number bounded by a timeout, a window that selects nothing, a cached
value, a code path that no longer runs. Each of those produces a stable number that a tolerance
check passes forever. This project has already been bitten by exactly that once.

So every run writes a machine-readable comparison — per metric, the old value, the new value, the
delta, the tolerance, the verdict, and **the conditions each contributing run was taken under** —
and the release procedure hands it to an agent:

```
%LOCALAPPDATA%\OutlookAI\Measurements\reports\<runId>.comparison.json
```

> Read this measurement comparison. Is there cause for alarm or a course correction?

The gate prints that path and that question at the end of every run. The file carries a
`readerQuestion` field saying the same thing, and each metric carries `baselineSamples`: the
individual earlier runs behind the median, each with its own value, commit and conditions — which
is what lets a reader notice that a number is suspiciously *stable*, or that its stated conditions
do not match what it claims to measure.

The gate helps where it mechanically can: a `both`-class continuous metric (ms, s, bytes, rates)
that has been **byte-identical across three or more runs and again now** raises a
`SUSPICIOUSLY STABLE` notice. Wall-clock measurements do not repeat exactly, so when one does it
is saying something.

**The console print is the floor.** The full table — every catalogued metric, baseline, now,
delta, run count, verdict — goes to the console on every run, pass or fail. That is the guarantee
the raw numbers are in front of whoever triggered the release even if nobody invokes the reader.
It is what replaces the release-notes table that Question 5 asked for and privacy rules out.

---

## Reading the report

```
METRIC                                          BASELINE           NOW     DELTA  RUNS  VERDICT
--------------------------------------------------------------------------------------------------
sweep.wholeStore7Day.elapsedMs                    36,600        12,000    -67.2%     5  FAIL
    smaller/faster by 67.2% (tolerance 10.0%): 36,600 ms -> 12,000 ms. A number that got
    smaller is not good news by default - the usual cause here is that something stopped being
    done (a sort that started failing, a folder set that shrank, a filter that stopped
    matching). Prove which before accepting it.
```

- **BASELINE** is the **median of the last 5 runs in the same scope** (`-BaselineRuns` changes
  the window). A median resists one noisy run without letting an ancient value anchor the
  comparison forever. **RUNS** says how many actually contributed.
- **Scope** is `profile|corpusId`. Runs in different scopes are never compared.

### Verdicts

| Verdict | Meaning |
| --- | --- |
| `OK` | inside tolerance for its class |
| `FAIL` | beyond tolerance, in either direction; or a class rule broken |
| `NEW` | no baseline for this metric in this scope — nothing to compare |
| `MISSING` | this metric has history but was **not collected now**. Always a failure: coverage cannot be dropped silently |
| `ABSENT` | never collected, and `-Require All` was asked for |
| `-` | never collected and not collected now. Not a failure under `-Require Present`; it *is* a gap |

### Verdict classes, and why there are five

Each exists because the others get something wrong.

| Class | Rule | Used for |
| --- | --- | --- |
| `both` | beyond tolerance in **either** direction fails | the default. Every timing, size and rate. Catches "it got faster because it stopped doing something" |
| `coverage` | **any** decrease fails; an increase passes and is reported | test counts, folders scanned, samples taken. Less coverage is never acceptable; more is never the alarm |
| `noIncrease` | **any** increase fails | counters of things going wrong (`census.foldersDegradedToCount`, `suite.testsSkipped`) |
| `mustBeZero` | any non-zero fails, baseline or not | claims the codebase makes about itself that a run can check (`sweep.sortRefusedFolders`, `suite.testsFailed`) |
| `pinned` | **any change at all** fails | the 36 constants read out of the sources. A constant is not a noisy measurement — it either moved or it did not, and a budget constant moving between releases is exactly the course correction this gate is for |

### Tolerances

Default **±10 %**, symmetric, for `both`-class metrics. Change it for a run with `-Tolerance`;
change it permanently by editing the catalogue in `measurement-gate.ps1`. Per-metric overrides
are in the table below and each one has a stated reason — `suite.durationMs` is 0.35 because it
shares a machine with whatever else is running, the index frontier ages are 0.50 because they
measure a race between an indexer and arriving mail rather than a code path.

Two metrics carry an **absolute floor** (`MinAbsolute`) as well: `storeIndexProbe.*` are measured
at 9–30 ms, where a 10 % relative test is millisecond jitter. A change smaller than the floor
never fails.

### Notices

Not failures, and not safe to skip. They are what a tolerance cannot see: a dirty working tree, a
changed `storeSet`, an `unknown` profile or index condition, a collector that could not run, a
corpus anchor's age, and the `SUSPICIOUSLY STABLE` flags. They are the reader's job.

### Exit codes

| Code | Meaning |
| --- | --- |
| `0` | pass |
| `1` | something moved beyond tolerance, a class rule broke, or a metric with history went missing |
| `2` | **no baseline** — nothing was compared. Recorded, but not a pass. See below |
| `3` | the gate refused to run: store path inside a git tree, CI detected, bad input, unknown metric id |

### The cold-start case

A first run has nothing to compare against. The gate **records the run, reports
`NO BASELINE - nothing to compare` per metric, and exits 2.** It does not pass.

Accepting the new baseline is deliberate and explicit:

```powershell
pwsh -File .github/scripts/measurement-gate.ps1 ... -AcceptNewBaseline
```

The reason it is not automatic: *a first measurement can just as easily be a first measurement of
something already broken.* Read the values before you accept them.

### The aged-corpus refusal

`corpus-build` is deterministic from `(corpusId, seed, anchor)` and **every date band in the
corpus is relative to the `--anchor`, not to the clock**. So a corpus eventually ages past its
own windows: once the run date is more than 7 days past the anchor, a 7-day sweep over it selects
*nothing* — and a sweep over nothing is fast, stable and completely healthy-looking.

When `conditions.corpusAnchor` is set, the gate computes `corpusAgeDays` and **fails any
window-scoped metric whose window the corpus has outlived** (7 days for the sweep metrics, 60 for
the whole-store scan):

```
sweep.wholeStore7Day.elapsedMs: the corpus anchor is 53.9 days old and this metric is measured
over a 7-day window, so the window selects NOTHING. The number is a measurement of an empty
selection - fast, stable and meaningless. Re-anchor and rebuild the corpus, or drop this metric
from the run.
```

Re-anchor and rebuild the corpus, or leave those metrics out of the run and say so.

---

## When a change legitimately moves a number

**This path has to exist**, or the first real improvement is treated as a fault forever. It is
deliberately not a tolerance widening — widening the tolerance hides every *future* move as well,
which is the opposite of what a legitimate improvement deserves.

**Move one metric's baseline:**

```powershell
pwsh -File .github/scripts/measurement-gate.ps1 `
    -Annotate -Metric sweep.wholeStore7Day.elapsedMs `
    -Reason "table-read rewrite (a1b2c3d): sweep no longer opens each item, re-measured at 11 s over the same 5 stores with sortRefusedFolders still 0"
```

**Move every baseline** (a new machine, a rebuilt profile, a corpus rebuild):

```powershell
pwsh -File .github/scripts/measurement-gate.ps1 -ResetBaseline -Reason "profile rebuilt: 5 stores -> 4, Archive re-created"
```

`-Reason` is **required** and the gate refuses without one: an unexplained baseline reset is
indistinguishable from hiding a regression, which is the one thing this gate exists to prevent.
Each annotation is appended to `annotations.jsonl` with the reason, the author, the commit, and
the run it came after. From then on, that metric is compared only against runs recorded *after*
the annotation — earlier runs stay in the history as a record, and stop being a baseline.

A good reason names **what moved and what measurement justifies it**. `"faster now"` is not one.

**Deleting the store is also safe** and is the blunt version of a reset: the next run reports
`no baseline` for everything and refuses to pass until `-AcceptNewBaseline`. You lose the history
and the reasons, which is why the annotation path exists.

---

## Why the numbers are local-only

The maintainer's reasoning, and it is settled: measurements from their machine are meaningful
only relative to older values from the same machine, they are not representative of anyone else's
system, and they do not want statistics about their machine leaked. This repository is public.

That rule is enforced structurally, in five places, so it does not depend on anyone remembering
it:

1. **The store lives under `%LOCALAPPDATA%`**, outside every checkout.
2. **The gate refuses to write inside the repository.** Any `-StoreRoot` resolving under the repo
   root is refused with exit 3.
3. **The gate refuses to write inside *any* git working tree.** It walks up from the resolved
   store path looking for a `.git` directory *or file* (a worktree — which is what this repo uses
   for agent branches, and which a naive directory check walks straight past). A store one
   `git add` from publication is refused even in a repository the gate has never heard of.
4. **The gate refuses to run its comparing modes under CI at all** (`CI`, `GITHUB_ACTIONS`,
   `TF_BUILD`), because it prints real measurements and a CI log on this repository is public.
   `-AllowCi` exists for a private runner and says so out loud. `-SelfTest`, `-ListMetrics` and
   `-Template` are CI-safe and are not gated: they touch no store and print no measurement.
5. **Every write goes through one function** that asserts its target is under the validated store
   root. There is no second path that could drift.

And `.gitignore` covers `Measurements/`, `history.jsonl`, `annotations.jsonl`,
`*.comparison.json`, `*.measurement-run.json` and `measurement-history*`, so a store copied or
redirected into the tree does not become tracked by accident.

### The check that catches a future accidental commit

```
pwsh -File .github/scripts/check-measurement-privacy.ps1
```

CI-safe: it needs no measurements, touches no store and prints no values. Three checks.

1. **Nothing measurement-shaped is tracked.** Every record the gate writes carries a marker
   string, and that marker is **assembled from two halves in both scripts so that it never
   appears contiguously in any repository file**. That is what makes the check exact rather than
   heuristic: any tracked file containing the assembled marker is a real measurement record,
   whatever it is named and wherever it landed — and there is no allowlist to poke a hole in.
   Path shapes are checked too, for a record reformatted past its marker.
2. **The `.gitignore` rules are still there.** A deleted ignore rule is silent until the day
   something lands.
3. **The gate still describes what it gates** (`measurement-gate.ps1 -SelfTest`): the catalogue is
   internally consistent, every source-constant collector still matches its file, and the
   catalogue and this document name the same metrics **in both directions**. A metric added to
   the gate and left out of this document is a gate nobody can read; a metric documented but not
   gated is a promise the gate does not keep.

Run it alongside `check-pinned-constants.ps1`. **It is not yet wired into a workflow** — see
[Follow-ups](#follow-ups).

---

---

## All options

| Option | What it does |
| --- | --- |
| `-Run <path>` | ingest a measurement-run JSON (the live half). `-Template` prints a skeleton |
| `-Collect` | also gather everything that needs no mailbox: source constants, invariant counts, and — with `-TestLog` — the suite numbers |
| `-TestLog <path>` | `dotnet test` output or a `.trx`. Without it, `-Collect` records a notice that the suite numbers were not taken |
| `-Require Present\|All` | `Present` (default) compares what the run carries; `All` demands every catalogued metric. **`All` is the release invocation** |
| `-Tolerance <fraction>` | default symmetric tolerance for `both`-class metrics. Default `0.10` |
| `-BaselineRuns <n>` | how many prior runs the median is taken over. Default `5` |
| `-AcceptNewBaseline` | accept `NO BASELINE` metrics as the new baseline instead of exiting 2 |
| `-AllowUnknownMetrics` | permit metric ids the catalogue does not know. Off by default: a typo'd id would otherwise report `no baseline` forever and never fail |
| `-StrictConditions` | promote the `storeSet` change notice to a failure |
| `-DryRun` | compare and report, append nothing, write no artifacts |
| `-ProfileKind production\|vm\|unknown` | the profile condition. Part of the comparison scope |
| `-Indexed indexed\|unindexed\|mixed\|unknown` | the index condition. A change fails rather than being compared |
| `-CorpusId <id>` | which synthetic corpus. Part of the comparison scope |
| `-CorpusAnchor <yyyy-MM-dd>` | the `--anchor` the corpus was built from. Drives `corpusAgeDays` and the aged-corpus refusal |
| `-StoreSet <text>`, `-Notes <text>`, `-Label <text>` | free-text conditions and a label for the run |
| `-StoreRoot <path>` | override the baseline store. Refused if it resolves inside any git working tree |
| `-Show [-Last <n>]` | print the recorded history: run id, time, commit, scope, metric count, conditions |
| `-Annotate -Metric <id> -Reason <text>` | move one metric's baseline. `-Reason` is required |
| `-ResetBaseline -Reason <text>` | move every metric's baseline |
| `-Template` | print a run-file skeleton. CI-safe |
| `-ListMetrics` | print the catalogue as the markdown table below. CI-safe |
| `-SelfTest` | check the catalogue against itself, its collectors and this document. CI-safe |
| `-ReleaseNoteSummary` | print a one-line verdict carrying **no measurements at all**, and withhold the failure detail. For the one place the numbers may not go |
| `-AllowCi` | run the comparing modes despite `CI`/`GITHUB_ACTIONS`/`TF_BUILD` being set. For a private runner only |

## The 74 gated measurements

Generated from the gate's own catalogue. `pwsh -File .github/scripts/measurement-gate.ps1
-ListMetrics` reprints this table; `-SelfTest` fails if it and the catalogue stop agreeing, so
adding a metric means editing this section.

#### COM host (3)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `comHost.largestFrameBytes` | bytes | both | default | outlook_health comHost.largestFrameBytes high-water. THE cautionary tale: 432 KB read as 152x headroom was really a number bounded by a timeout, and the VM corpus later produced 10,734,599 bytes. Take it after the sweep, never before. |
| `connect.attachAndHealthMs` | ms | both | default | attach to a running Outlook plus one health probe. Measured 1.0 s; ConnectDeadlineMs is 180 s against it. |
| `connect.coldSearchMs` | ms | both | default | first search after a cold COM host start. Measured 6.2 s. Not comparable to any warm figure and must never be recorded as one. |

#### exhaustive scan (5)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `scan.inboxOnly.elapsedMs` | ms | both | default | exhaustive.elapsedMs, Inbox only, no subfolders. |
| `scan.inboxWithSubfolders.elapsedMs` | ms | both | default | exhaustive.elapsedMs, Inbox with subfolders. 66.5 s here is the other half of the OperationDeadlineMs derivation. |
| `scan.wholeStore60Day.elapsedMs` | ms | both | default | exhaustive.elapsedMs for a 60-day whole-store scan. Use a term that matches nothing, so the scan runs to the end of its budget instead of stopping at SearchTopCap. |
| `scan.wholeStore60Day.foldersScanned` | folders | coverage | n/a | exhaustive.foldersScanned for that scan. 3 of 32 in 105 s is why ExhaustiveScanDeadlineMs is 615 s; this is the number that says whether it is enough now. |
| `scan.wholeStore60Day.itemsPerSecond` | items/s | both | default | throughput for that scan. Step 5 of corpus-measurement-plan.md, which has never been run on either machine - so ExhaustiveTimeBudgetMs has no throughput measurement behind it at all. Expect NO BASELINE until somebody runs it. |

#### freshness sweep (8)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `sweep.foldersSwept` | folders | coverage | n/a | sweep.foldersSwept. Fewer folders swept is lost coverage however fast the sweep got. |
| `sweep.itemCappedFolders` | folders | both | default | count of folders truncated by SweepPerFolderCap. Documented as "never fires in steady state" on a real profile; movement off zero is that claim breaking. |
| `sweep.itemsBodyCapped` | items | both | default | items whose body was cut at SweepBodyCharsCap. Moves with the correspondents, not with the code, which is why it is reported rather than pinned. |
| `sweep.perStore.elapsedMs.max` | ms | both | default | slowest single store in the per-store breakdown. The number that decides whether one bad store can spend the whole budget. |
| `sweep.perStore.elapsedMs.total` | ms | both | default | sum over stores. Compared against sweep.wholeStore7Day.elapsedMs it says how much of the sweep is not per-store work. |
| `sweep.sortRefusedFolders` | folders | mustBeZero | n/a | sweep.sortRefusedFolders. 03a0857 made the claim "the received-date sort applies" checkable and nothing has checked it on a real profile since. Non-zero means capped folders kept an ARBITRARY slice of the window, which is the exact shape that produced a wrong 180 s budget. |
| `sweep.wholeStore7Day.elapsedMs` | ms | both | default | search payload sweep.elapsedMs for a whole-profile 7-day sweep, COM host already warm (a cold host adds the 90 s connect floor and is not comparable). |
| `sweep.wholeStore7Day.itemsSeen` | items | both | default | sweep.itemsSeen for the same sweep. The denominator every elapsed figure above is a rate over; a drop here with elapsed flat means the sweep got slower per item. |

#### invariants (3)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `invariants.comHostFilesScanned` | files | coverage | n/a | COM-host source files that scan reached. A drop means the scan stopped seeing a directory, which turns the check into one that always passes. |
| `invariants.comHostThrownTypes` | types | coverage | n/a | exception types raised behind the IOutlookSession contract, all of which must be modelled by ComHostErrorMapper. |
| `invariants.pinnedConstantChecks` | checks | coverage | n/a | cross-file invariants check-pinned-constants.ps1 asserted. Baseline 11. A check deleted is a check that stops proving anything, and nothing else notices. |

#### search index (6)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `index.frontierAgeMinutes.median` | min | both | 0.5 | median index frontier age over the sampled probes. Measured ~6 min. Tolerance is deliberately wide: this is a race between an indexer and arriving mail, not a code path. |
| `index.frontierAgeMinutes.p90` | min | both | 0.5 | p90 of the same samples. StaleIndexNoticeMinutes (30) IS this number; if it has moved, the constant is stale. |
| `index.frontierSampleCount` | samples | coverage | n/a | how many probes the two figures above are computed from. A p90 over 5 samples is not a p90; fewer samples than last time is a weaker statistic and fails. |
| `search.indexQueryMs.median` | ms | both | default | median wall clock of the index query alone. Measured healthy at 60-550 ms, which is the whole justification for SearchIndexTimeoutSeconds = 60. |
| `storeIndexProbe.delegateMissMs` | ms | both | default | per-store index probe, delegate-subtree miss. Measured 9-10 ms, which is why StoreIndexProbeBudgetMs is 1,500. Absolute floor of 5 ms so single-millisecond jitter cannot fail it. |
| `storeIndexProbe.discoveryMissMs` | ms | both | default | per-store index probe, @-discovery miss. Measured 27-30 ms. |

#### source constants (36)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `source.AutostartCooldownMs` | ms | pinned | n/a | cooldown between autostart attempts. |
| `source.CensusPerFolderLimit` | items | pinned | n/a | largest folder the tripwire baseline will identify item by item. |
| `source.CensusPerStoreItemBudget` | items | pinned | n/a | identity budget per store. Bounds the whole profile at stores x this. |
| `source.CensusRepeatGrowthHeadroom` | x | pinned | n/a | growth a folder may show between the two censuses and still be walked. |
| `source.CleanExitGraceMs` | ms | pinned | n/a | grace given to a COM host asked to exit. |
| `source.ConnectDeadlineMs` | ms | pinned | n/a | COM session establishment, cold Outlook start included. |
| `source.ExhaustiveScanDeadlineMs` | ms | pinned | n/a | exhaustive scan hard deadline, its own class. |
| `source.FreshnessSweepDeadlineMs` | ms | pinned | n/a | freshness class hard deadline - the threshold the sweep budget is judged against. Derived as SearchBudgetMs + ResultReturnHeadroomMs; narrowing the sweep budget must move it too. |
| `source.ExplorerFolderSettleDelayMs` | ms | pinned | n/a | wait for an Explorer to settle on a folder change. |
| `source.HandshakeBudgetMs` | ms | pinned | n/a | COM host pipe handshake, both ends. |
| `source.HealthIndexTimeoutSeconds` | s | pinned | n/a | index query timeout inside outlook_health. |
| `source.HealthPerStoreIndexBudgetMs` | ms | pinned | n/a | per-store index probe budget inside outlook_health. |
| `source.HealthProbeDeadlineMs` | ms | pinned | n/a | health probe. The instrument, not the work - deliberately short. |
| `source.LiveInboxArrivalDeadlineSeconds` | s | pinned | n/a | live transport arrival deadline. Compare against transport.inboxArrivalSeconds. |
| `source.MaxFrameBytesMiB` | MiB | pinned | n/a | protocol frame ceiling. comHost.largestFrameBytes is measured against it. |
| `source.MinimumDispatchDeadlineMs` | ms | pinned | n/a | floor under any dispatch deadline. |
| `source.MinimumItemBudgetMs` | ms | pinned | n/a | floor per item inside a batch move. |
| `source.MoveBatchBudgetMs` | ms | pinned | n/a | whole move/archive batch. Compare against move.batch50.elapsedMs. |
| `source.OperationDeadlineMs` | ms | pinned | n/a | shared COM operation deadline; 4.5x the slowest healthy operation measured. |
| `source.RecipientResnapshotDelayMs` | ms | pinned | n/a | wait before re-reading resolved recipients. |
| `source.ResultReturnHeadroomMs` | ms | pinned | n/a | reserved for handing the result back rather than for doing work. |
| `source.ScopedSweepTimeBudgetMs` | ms | pinned | n/a | subtree walk budget for a scoped sweep. |
| `source.SearchIndexTimeoutSeconds` | s | pinned | n/a | index half of the search budget. |
| `source.SearchTopCap` | items | pinned | n/a | largest result set that crosses the pipe; ResultReturnHeadroomMs is sized against it. |
| `source.StaGiveUpAfterMs` | ms | pinned | n/a | how long the STA runner retries a busy Outlook before giving up. |
| `source.StaleIndexNoticeMinutes` | min | pinned | n/a | p90 of the dev profile index frontier age over 177 probes. Compare against index.frontierAgeMinutes.p90. |
| `source.StaRetryAfterMs` | ms | pinned | n/a | SERVERCALL_RETRYLATER retry interval. |
| `source.StartBackoffMs` | ms | pinned | n/a | backoff after repeated start failures. |
| `source.StartFailureBackoffThreshold` | failures | pinned | n/a | start failures before backing off. |
| `source.StoreIndexProbeBudgetMs` | ms | pinned | n/a | per-store index probe budget on the search path. |
| `source.SweepBodyBytesBudgetMiB` | MiB | pinned | n/a | accumulated body bytes one sweep may return. Load-bearing since the 10,734,599-byte corpus high-water. |
| `source.SweepBodyCharsCap` | chars | pinned | n/a | per-body character cut in the sweep. |
| `source.SweepBudgetMs` | ms | pinned | n/a | freshness sweep budget. The one that was derived from a measurement taken while the sort was failing; 600 s since 2026-08-24 is a CEILING awaiting that re-measurement, and it moved to ComOperationBudgets with the sweep's own deadline class. |
| `source.SweepPerFolderCap` | items | pinned | n/a | items per folder the sweep will open. |
| `source.UnresponsiveCooldownMs` | ms | pinned | n/a | how long the breaker stays open. |
| `source.UnresponsiveTimeoutThreshold` | timeouts | pinned | n/a | consecutive timeouts before the breaker opens. |
| `source.VeryStaleAdviceMinutes` | min | pinned | n/a | upper rung of the staleness ladder. |

#### suite (5)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `suite.durationMs` | ms | both | 0.35 | suite wall clock. Tolerance 0.35 rather than the default 0.10 because this shares a machine with whatever else is running; tighter produced false failures in trial. |
| `suite.testsFailed` | tests | mustBeZero | n/a | failures. Any is a failure, baseline or not. |
| `suite.testsPassed` | tests | coverage | n/a | passing tests under the standing verification filter. Baseline 1,936. |
| `suite.testsSkipped` | tests | noIncrease | n/a | skips. A test that starts skipping is coverage lost without the count dropping, which is the quiet version of the same failure. |
| `suite.testsTotal` | tests | coverage | n/a | total discovered under the same filter. |

#### tripwire census (6)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `census.elapsedMs.maxStore` | ms | both | default | slowest single store. This is what the 3-minute STA join in LiveOutlookTestMailer actually bounds, and the number a live run refuses to start on. |
| `census.elapsedMs.total` | ms | both | default | whole-profile baseline census wall clock. The 2026-08-20 table-read rewrite put this at 16.9 s for 5 stores / 159 folders / 2,044 items; before it, one store alone exceeded the 3-minute STA budget. |
| `census.foldersDegradedToCount` | folders | noIncrease | n/a | folders the plan chose to walk and could not, so identity was lost for them. A table missing its columns on every folder would disable half the tripwire, and it must not do that quietly. |
| `census.foldersWalked` | folders | coverage | n/a | folders the census reached. |
| `census.itemsWalked` | items | coverage | n/a | items the census identified. The denominator of census.elapsedMs.total. |
| `census.storesWalked` | stores | coverage | n/a | stores the census reached. Recorded as a metric and not only as a condition, so a shrinking profile shows up in the diff instead of silently making every elapsed figure look better. |

#### write paths (2)

| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |
| --- | --- | --- | --- | --- |
| `move.batch50.elapsedMs` | ms | both | default | the hub-only 50-item move/archive batch (6.4(3)). A PST move is a local file operation and an Exchange move is a server round trip; MoveBatchBudgetMs (240 s) has never been measured against the second. |
| `transport.inboxArrivalSeconds` | s | both | 0.5 | send-to-self round trip. LiveInboxArrival.DeadlineSeconds is 180 because a real round trip once exceeded 120 s and failed a 17-minute run. Wide tolerance: this is the mail system, not the code. |
---

## What is NOT gated, and why

Ten numbers this project produces or has produced are **not** in the catalogue. They are listed
because a gate that covers a sample and implies it covers everything is worse than one that says
where it stops.

| Not gated | Why not | What it would take |
| --- | --- | --- |
| **Corpus build throughput** (50.9 items/s, 40 000 items in 12m27s) | needs a full corpus rebuild on the VM to observe — it is a number about *making* the fixture, not about the product | a `corpus-build` run that emits machine-readable timings, plus a decision that rebuilding is part of a release |
| **Per-call COM cost** (~0.3 ms local vs 12–15 ms delegate Exchange) | derived by hand from two *different* operations. `Docs/vm-coverage-analysis.md` §3.1 says explicitly that the "1,200 vs 12" pair is not a measured pair, and warns against quoting it as one | an instrument that times one COM call on each store class. Nothing emits this today |
| **Sweep cost-model coefficients** (~15 ms/item opened, ~19 ms/folder fixed) | fitted offline over 215 sweeps. Re-fitting needs `Docs/v3-probes/soakfix13-probe-sweep-cost.ps1`, which is **gitignored** and exists only on the dev machine | commit a synthetic-safe probe, or emit per-folder timings from the sweep itself. Today the server reports one clock (the `sweep` payload's `elapsedMs`) and no per-folder breakdown at all |
| **Exhaustive-scan items/second on the corpus** (step 5 of `corpus-measurement-plan.md`) | *the id exists* (`scan.wholeStore60Day.itemsPerSecond`) but it has **never been run on either machine**, so `ExhaustiveTimeBudgetMs` at 600 s has no throughput measurement behind it. It will read `NEW`/`-` until somebody runs it | the step-5 procedure: a search term that matches nothing, so the scan runs to the end of its budget |
| **Build wall clock** (Core net48 + net10) | dominated by NuGet restore and OS disk cache, not by the code. Gating it at a tight tolerance produces constant false failures; gating it loosely gates nothing | a warm-cache, no-restore measurement taken the same way every time |
| **Live-tier method inventory** (Portable / ProfileBound counts) | **already pinned** by `T1/LiveTierInventoryTests` by reflection over the assembly. Re-counting it here would create a second copy of a list, which is the exact failure `check-pinned-constants.ps1` exists to prevent | nothing — this is deliberate, not a gap |
| **`SweepPerFolderCap` "never fires in steady state"** | that is a claim about *many* sweeps (0 of 215 hit the cap), not about one. `sweep.itemCappedFolders` gates the single-run half; the distribution is not expressible in a per-run gate | an aggregate the server keeps across sweeps, or a soak harness |
| **`PumpedStaRunner` retry behaviour** (`RetryAfterMs` 250, `GiveUpAfterMs` 30 000) | the constants are pinned, but the *behaviour* is not measured: nothing counts `SERVERCALL_RETRYLATER` retries. `Docs/vm-coverage-analysis.md` §3.2 notes the retry path is unreachable on the VM entirely | a retry counter in `PumpedStaRunner`, surfaced through `outlook_health`. This is a production-code change and therefore a maintainer's call |
| **Signature-snapshot and zero-artifact sweep timings** | safety machinery, not performance. Their cost is bounded by the census metrics that already are gated | — |
| **Memory, handle and process counts** | the project emits none, on either side of the COM boundary | an instrument that does not exist yet |

---

## Follow-ups

**1. Wire `check-measurement-privacy.ps1` into CI.** It is written, tested and CI-safe, but
`.github/workflows/` was outside the territory of the change that added it, so nothing invokes it
yet. It belongs beside the existing pinned-constants step in `build.yml` and `release.yml`:

```yaml
      - name: Check no measurement data reached the repo
        run: .github/scripts/check-measurement-privacy.ps1
```

Until that lands, the check has to be run by hand — which means the accidental-commit guard is
only as good as somebody remembering it.

**2. `scan.wholeStore60Day.itemsPerSecond` has never been measured.** It is catalogued so its
absence is visible rather than silent. `Docs/corpus-measurement-plan.md` step 5 says exactly how,
it is cheap, and it is read-only on the VM. It is the one measurement that would settle whether
`ExhaustiveTimeBudgetMs` at 600 s is sized correctly.

**3. The sweep emits one clock and no per-store breakdown.** `sweep.perStore.elapsedMs.max` and
`.total` are catalogued because they are the numbers that say whether one slow store can spend
the whole budget — but the server does not report them, so today they have to be taken by running
the sweep once per store. Emitting a per-store breakdown in the `sweep` payload would make three
catalogued metrics collectable directly. Production-code change, maintainer's call.

---

## See also

- `Docs/vm-coverage-analysis.md` §6.4 — the argument for this gate (direction C), and §7
  Question 5 — who reads the result. The release-notes half of Question 5's recommendation is
  superseded here.
- `Docs/corpus-measurement-plan.md` — how to take the sweep and scan measurements, and the three
  traps that produce a wrong number that looks right.
- `Docs/magic-numbers.md` — every constant in the repository and where its value came from.
- `.github/scripts/check-pinned-constants.ps1` — the sibling mechanism, for values that exist
  twice in two languages.
