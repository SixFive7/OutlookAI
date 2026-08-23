# RESUME HERE - state of play at 2026-08-23, late

**Read this first after any context loss.** Everything below it is history and reasoning.

## Position - 2026-08-24, after the VM-infrastructure merge

`HEAD` = `e7384dc`, pushed, tree clean. **2,024 tests in 8 seconds with no mailbox contact**
(1,936 + 41 + 47 from the two merges). `OutlookAI.Core` builds clean for net48 and net10 with
zero warnings. `check-pinned-constants.ps1` 11/11. The test VM is running, checkpoints intact.

**A CREDENTIAL WAS LEAKED FROM THIS FILE AND MUST BE ROTATED.** This repository is PUBLIC
(`SixFive7/OutlookAI`). This file recorded the VM guest password in plain text; it was pushed and
is in **32 commits of history**, first `e1d8c6c`. Redacted at `HEAD` in `77df4e4`, which stops it
spreading and does NOT un-publish it. **Rotation is the fix and it has not been done** - it needs
the maintainer's word, because changing a credential is a real mutation, and the scripts under
`C:\Users\jori\Downloads\tmp-outlookai-vm\` hard-code the old value. Never write the new one into
a tracked file; the gitignored live-test settings are the only place it belongs.

**Verification command - the standing bar:**

    dotnet build McpServer/OutlookAI.Core/OutlookAI.Core.csproj
    dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj --filter "Category!=Live&FullyQualifiedName!~Tests.T3."
    pwsh -File .github/scripts/check-pinned-constants.ps1

`Category!=Live` is now honest (commit `7a38458`) - measured, no COM host spawned across 481
process samples. The narrow filter is still used because tier 3 spawns server processes and is slow,
not because it is unsafe.

## IN FLIGHT - 2026-08-24

| Agent | Branch | State |
| --- | --- | --- |
| Fix the VM test infrastructure | `worktree-agent-ad5951c0e2020cddf` | **MERGED** `4d7efbc` - corpus freshness + re-anchor, census, placement probe, mail sink, runbook |
| Clear the product gap map | `worktree-agent-ab7461aa27a49e30a` | **MERGED** `e7384dc` - C5, E3, B4, B5, the two `snippet_chars` clamps; A5 and F2 verified already closed |
| Mutation-verify the sort fix | (worktree) | running - `bea7fc9`, the queued verification that never happened |
| Build the measurement gate | (worktree) | running - local-only baselines under `%LOCALAPPDATA%`, fail-biased tolerances |

**The completeness gap map now has one row left: H3**, which needs a live corpus re-run rather
than a fix.

Branches are real refs and survive any conversation loss; `git branch --list` finds them. Agents
commit on their own branch, never push, and never edit this file. **To finish one:** merge its
branch into `master`, then run the standing verification command above.

**A heartbeat monitor is armed** (`bash ~/.claude/scripts/heartbeat.sh`, persistent).
`state=all-finished` is the only safe signal to stop it, and `bgroot=` must be read before
stopping because root-owned background jobs are never killed for you.

## NINE OPEN QUESTIONS PUT TO THE MAINTAINER - 2026-08-24, UNANSWERED

Asked in full, each with a primer, directions and a recommendation. **Nothing below is being
implemented until they are answered.** Recorded here so a compaction cannot lose them.

| # | Question | My recommendation |
| --- | --- | --- |
| 1 | The 16 mislabelled tier-3 files - 8 of 100 methods reach real Outlook, and the interim filter discards the other 92 locally | Move the 8 into the live tier, implemented as a file split |
| 2 | How far to push `Requires` from class to method - all ~30 classes, or the 6 straddlers | The 6 straddlers now, the rest lazily as each is enabled |
| 3 | Who reads the measurement table, now that release notes are ruled out | An agent reads it against the previous run; console print as the floor |
| 4 | The four unmeasured atomicity residuals | Measure the RPC HRESULT question and the soft-delete survival; accept the other two as documented |
| 5 | The tripwire's re-run bound | Two re-censuses ~30 s apart, then one bounded re-run of implicated tests |
| 6 | `SweepBudgetMs` 180 s (derived while the sort was broken) and the census identity budget (16.9 s, one trial) | Ceilings now - 600 s and 120 s - narrowed later from VM data |
| 7 | `ExhaustiveScanDeadlineMs` 615 s has never been measured on either machine | Run `corpus-measurement-plan.md` step 5 on the VM; read-only |
| 8 | Four `Open - needs a decision` rows in `magic-numbers.md` | Fix the update-service backoff and the row constant; accept the tint; close the registry row |
| 9 | **The leaked VM password** - public repo, 32 commits of history | **Rotate the credential.** History rewrite is optional hygiene, not the fix |
| 10 | The freshness-sweep cache is unreachable for every UNSCOPED search - a cost regression since `c515565`, verified against history. Directions in `TODO.md` | Re-key the unscoped cache on the profile frontier |
| 11 | Should `thread` apply a store scope it DERIVED? C5 is closed on reporting; the behaviour is untouched, so a member in a second account is still absent - now named rather than unmentioned. Directions in `TODO.md` | Stop scoping when the store was derived; keep it when the caller named one |

## Everything outstanding, in one list

**Decided, not built:**
1. Corpus freshness assertion, then re-anchor on restore. A fixed anchor means every "last N days"
   window selects nothing after ~6 weeks **and every test still passes**.
2. Four tests that return early rather than assert must fail instead - including the one asserting
   that search always answers, which on an indexless machine asserted nothing.
3. A local SMTP sink that delivers back. An unroutable dummy account leaves a permanent tagged
   Outbox artifact and the mandatory zero-artifact sweep fails on it forever.
4. A third tier value `VmCapable`, and push `Requires` from class level to method level. Class-level
   attribution is why 96 tests looked impossible when only 15 are.
5. Fault hooks for shapes fixed blind - a folder that throws on open, a store whose display name
   cannot be read (**no way to produce this on any machine, so that fix has never once executed**),
   an item with no delivery time.
6. The pre-release measurement gate. An agent reads all the numbers; compares against THIS machine's
   own history; movement in either direction counts; bias is to fail. **Numbers stay local - never
   in the repo, never in release notes.** Baseline history belongs under `%LOCALAPPDATA%`.
7. **The VM build**, now two Windows ACCOUNTS rather than two data files - see the section on why a
   profile cannot split the index. Plus profiles, the sink, three stores, and a build-from-nothing
   runbook with seed instructions.

**Bug queue:**
8. **Mutation-verify `bea7fc9`** (the sort fix). Committed and green but its pass never ran; the
   killed agent left one mutation applied which had to be found from the failures alone.
9. **Re-measure the sweep budget.** 180 s was derived while the sort was silently failing, so it
   rests on a measurement of broken behaviour.
10. Corpus generator's two defects: ~5,500 duplicate Outbox items on a large build, deterministic;
    and the placement probe's folder-table check failing against a folder with many items.
11. Live move-batch exercise before release - making that batch a real aggregate changed behaviour
    on a mutating path.
12. Tripwire re-census-then-re-run, bounded by a maximum.
13. **All remaining gap-map rows** - the maintainer said clear every one.
14. `thread`'s store asymmetry plus a scan for the same shape elsewhere.
15. Restore the installed MCP server - deferred to the release.

**Reported, not fixed, needing a decision at some point:** `list_accounts` starts Outlook when it is
not running; `send` catches only `TimeoutException`; `SendUsingAccount` is written before the
identity readback and never restored.

## Send-path and startup decisions - 2026-08-23, late

**1. `list_accounts` starting Outlook is ACCEPTED as-is (option a).** The question was whether a
call that reads like a query should launch a mail client. It turns out it does not launch anything
visible: the server starts Outlook **headless** - the code says so in as many words, and
`outlook_health` carries a `headless` field which read `true` on the test VM. No window, and no tray
icon either, so it is less visible than the tray-only case the maintainer was willing to accept.
The declaration and the health-style guard were therefore dropped as unnecessary.
**Carried caveat:** headless is not free. That Outlook holds the data-file locks and keeps running,
and this project's own history records a wedge involving precisely the COM-activated
`OUTLOOK.EXE -Embedding` form. Invisible is not the same as harmless.

**2. The interrupted-send message keys on whether `Send()` was reached, not on which exception
arrived (option c).** Today the careful wording - naming the Outbox and Sent Items, stating the
outcome is unknown, saying explicitly not to send again - hangs off a `catch` on
`TimeoutException`, so it covers one failure mode of several while a generic path answers the rest.
Invert it: once `Send()` has been issued, ANY failure gets the specific message. What the caller
needs to know was never determined by the exception type; it is determined by whether the mail may
already be on its way.

**3. The send path must verify identity BEFORE writing the account pin, then restore if a window
remains.** Today it writes `SendUsingAccount` onto the draft and only then verifies, so a refused
send leaves the draft altered - contradicting a refusal that says nothing happened, and possibly
sending later from the wrong account. Reordering removes the problem rather than compensating for
it, and it matters that the alternative - restoring on the failure path - means adding a MUTATING
call to a path that is currently refusing to mutate, which is exactly why this was deferred when it
was first found.

## Standing rules

- **Completeness outranks performance, whatever the cost.**
- **Never ask about release timing.** Settled. They will say.
- **Never run the live tier or touch the mailbox from a subagent.** I run live tests; agents write
  them. Agents must not touch Hyper-V or the VM scratch folder.
- **Verify both frameworks.** A net48 break reached master because only net10 was checked.
- **Only a build proves a tree clean after a mutation pass.** Restore by index, never by matching
  replacement text; a killed process skips cleanup; passes must be serial.
- **Commit research into the repo.** The scratch folder was deleted once and took every long-form
  analysis with it.
- **No windows, no focus theft.** `-WindowStyle Hidden` belongs on `Start-Process`, never inside
  `-ArgumentList` where it is a no-op. Output must go to a file: an elevated process's stdout cannot
  reach the caller.
- **PowerShell cannot drive these COM objects by late binding** - `Table.GetRows`, `Table.Sort`,
  `CSearchManager` all failed in a way that reads like the API refusing. Use
  `$obj.GetType().InvokeMember(...)` or write the probe in C#.
- **The guest runs Windows PowerShell 5.1** via PowerShell Direct: no `??`, no ternary.

## The biggest findings, for context

- **The freshness sweep never sorted**, for the life of the feature, on any store - it passed a
  namespace-qualified property name that `Table.Sort` refuses, and the failure was swallowed. So its
  200-item cap always cut an arbitrary slice. Measured 5/5 stores. Fixed in `bea7fc9`.
- **Budgets were about half the measured work**, which is why the COM host was being killed during
  ordinary searches.
- **Sixteen atomicity claims were false** - the product said nothing had changed when nobody could
  know. Fixed in `7b4cfd9`.
- **The tripwire could not take a baseline at all** on the real profile; its census now reads a
  table instead of opening every message - 5 stores, 159 folders, 2,044 items in 16.9 s.
- **`Category!=Live` read the real mailbox on every verification run** for the whole session.

---


# Autonomous session log - 2026-08-18/19

**What this file is.** The maintainer went to sleep mid-session and asked two questions to be
answerable on return: *why is it not done*, and *what did you decide on your own so I can review
it*. This file answers both. It is updated as work lands, and it is the first thing to read after
a context compaction.

**Working rule while the maintainer is asleep:** decisions already given are executed; anything
genuinely new is decided with the best available reasoning, recorded in section 3 with the
reasoning, and flagged for review rather than buried.

---

## 1. Decisions the maintainer gave, and where each stands

| # | Decision | State |
| --- | --- | --- |
| Tripwire strictness | Detect complications and re-run the affected tests, up to a maximum, rather than failing outright | **queued** |
| Frame size | Cap bodies at the COM layer | **shipped** `b46cf8a` |
| Second PST on the testbed | Add one via the tested helpers | **queued - and now load-bearing, see section 3.36** |
| Live tier | Move the intermediate tier to the VM; keep the ability to run everything against the real system before a release | **shipped** (uncommitted, 2026-08-19) - traits, guards, settings and runbook; 19 of 115 tests move, and section 3.37 says why the other 96 cannot |
| `Stick-Test` VM + scratch | Delete both | **done** |
| Installed MCP server | Leave disabled until the release | **done, nothing to do** |
| Exhaustive scan | Resumable walk with a continuation token | **shipped** (uncommitted, 2026-08-19) - gap F2 closed; design in `tmp-aitrace/resumable-scan-design.md`, record in `Docs/completeness-gaps.md` F2 |
| `Table.Sort` namespace reference | Settle whether the sweep's sort has ever worked; write it as a read-only T2 test and do NOT run it | **probe written, NOT run** (`T2/LiveTableSortProbeTests`); the split `catch` and `sweep.sortRefusedFolders` shipped alongside |
| `thread` store asymmetry | Derive the warning from Outlook's store list; also scan for the same asymmetry elsewhere | **queued** |
| Timeout defects | Fold all three into the timeout-raising pass | **shipped** `4502c92` |
| COM host kill | Keep the hard kill, document it, add a brief wait before killing, make the kill outcome-aware | **shipped** `4502c92` |
| `top` ceiling | Leave at 100; rely on resumption | **decided, no work** |
| Remaining gap-map rows | Clear **all** of them before the release | **queued** |
| Work order | Infrastructure first: corpus, second PST, live tier on the VM | **in progress** |
| `update_draft` | **(d)** make it re-entrant: record intent first, so a retry completes rather than repeats | **shipped** `db34923` |
| Sweep timeout | **(d)** make expiry graceful **and** distinguish budget expiry from unresponsiveness at the supervisor | **shipped** `4502c92` |
| H3 (undated mail invisible to the sweep) | Check whether DASL can express "or the property is absent" first; failing that, report it; full fallback enumeration only if it proves common | **answered by measurement - NOT fixed, see section 3** |

## 1b. Decisions given 2026-08-19, after the overnight run

| Question | Answer | State |
| --- | --- | --- |
| The two corpus-generator defects (Outbox duplication; placement probe failing on a large folder) | **Fix both** | queued |
| If the sort probe confirms the sweep has never sorted | **Fix immediately, AND re-measure the corpus** - the 180 s sweep budget was measured with the sort failing, so a working sort may change the cost | conditional on the probe |
| `discard_draft`'s "Nothing was changed" wording | **Audit every such claim in the product**, not just that one - the phrase asserts atomicity and the product has been wrong about it once already | **audit delivered** (`tmp-aitrace/atomicity-claims.md`, 31 claims, 16 wrong); **fix shipped** (uncommitted, 2026-08-20) |
| Authorising the `claude.ai` and `VF Dev` MCP servers | **Leave them** - nothing here depends on either | closed |

## 1c. Decisions given 2026-08-20

**STANDING INSTRUCTION: do not ask about release timing.** The maintainer will say when it is
time. Asking again is noise, not diligence. This survives compaction; do not reintroduce the
question.

| Question | Answer |
| --- | --- |
| Atomicity audit | **Fix all of it now** - every row, not the three prose-only ones |
| `outcome: unchanged / applied / unknown` payload field | **Add it** |
| Orphaned draft after a failed `new_draft` | **Both fixes**, with registering the id before the post-save steps as the substance |
| Folders created before a refused move | **Fix now**, accepting the result-shape change |
| Wording vs measurement for the disconnect claims | **Fix the wording regardless** - a claim about what did not happen must not rest on an unmeasured probability |

**All five atomicity answers are SHIPPED (uncommitted, 2026-08-20).** All 31 claims were
enumerated, 16 were wrong, 16 are fixed. The `outcome` field is on the error object and on
`MoveItemView`; the two behavioural fixes landed (a failed create registers its draft's id before
the call is judged, a refused move reports the folders it created); the shared opening sentence is
`Core/Com/MutationOutcome`, keyed on `ComSessionOperations.IsRetryable`, with every site keeping its
own remedy clause. Pinned by T1 `AtomicityClaimsTests` (33 tests), including the assertion that was
missing when `db34923` shipped - the tool layer's `advice`, not only the `message`. The rows and
what each now says are in `Docs/completeness-gaps.md` section 7b; what reading could not settle is
in `TODO.md`.
| `HealthProbeDeadlineMs` at 5 s | **Keep** |
| Move batch as a real aggregate | **Keep, with a live batch exercise before release** |
| `scan_resumed` on every resumed page | **Keep** |
| Corpus generator's no-accounts guard, no override | **Keep** |
| Sort probe | **Run on both** the production profile and the VM |
| Second PST vs profiles | **Both.** Profiles solve the account constraint; the second PST is the absent-arrival-folders shape AND the tripwire's bystander - with one store that is also the hub, the tripwire censuses it, identifies nothing and CANNOT fail |
| Boundary of "all tests possible on the VM" | **Two stores, one indexed and one not**, so both search tiers are exercisable there. Delegate stores and real transport stay production-only. **Explicitly NOT** a faked delegate store: delegate mailboxes are indexed without folder nesting, which is real Outlook behaviour a local PST cannot reproduce, and faking it would give false confidence in the area this product has most often been surprised by |

**The VM goal is now "all tests possible there", not the 19 that move today**, and the VM must be
**reproducible from nothing**: how to build it - Windows, Office, the add-in, the profiles, the
store layout - and how to generate the seed corpus, so it can be rebuilt when deleted or moved to
another machine. That matters more than usual because two of this session's measurements only mean
anything against a corpus of known shape.

## 1d. VM build shape - decided 2026-08-20

**Three stores, not two.** The earlier pair of answers did not compose: a plain near-empty data
file cannot serve as an index-tier fixture, and indexing the corpus would destroy the unindexed
testbed that made this session's sweep measurements possible. So:

| Store | Indexed | Purpose |
| --- | --- | --- |
| Corpus A | **yes** | index-tier tests, and the shape 96 of 115 live tests want |
| Corpus B | **no** | the degraded path: no index frontier, seven-day fallback window, the sweep and frame measurements |
| Plain bystander | n/a | the absent-arrival-folders shape (Q5), and the store the tripwire watches - it must be one the tests never touch, which is why it is not a fixture |

**Dummy account: unroutable server, send enabled.** Drafts, updates and discards become reachable;
a send queues in the Outbox and can never leave, so delivery is physically impossible on a machine
whose purpose is running destructive tests unattended. **Open sub-task, to be checked rather than
assumed:** whether any live test asserts on Sent Items *after* a successful send. If some do, those
need a local SMTP sink; if none do, the unroutable account is sufficient and the sink is not built.

**Tripwire on a suspected loss: re-census first, re-run only if it persists.** A person reading
their mail produces a one-off delta; a test that deletes something reproduces it. A second census
costs seconds against a 27-minute tier run and separates ambient activity from a real fault without
running a single test. Re-running the plausibly-implicated tests is the fallback when the delta
survives the second census, bounded by a maximum.

## 1e. Test-execution policy - decided 2026-08-23

**The whole cycle moves to the test VM. The maintainer's mailbox comes out of the loop.**

The trigger: `Category!=Live` does NOT mean "does not touch Outlook". Sixteen tier-3 files spawn
the real server, which spawns a COM host, which attaches to the maintainer's production Outlook,
and they call `outlook_health`, `search` and `list_accounts` through it. Several are named
`...CiToolShapeTests`. Every verification run this session has been reading their real mailbox.
Reads only - nothing was created, moved, edited or deleted - but it was neither intended nor
declared, and the label is a lie.

**Interim policy, effective immediately, until the VM is ready:**

    dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
      --filter "Category!=Live&FullyQualifiedName!~Tests.T3."

Measured: **1,927 tests in 9 seconds, 0.05 s of Outlook CPU** (noise). The full non-live suite is
2,081 tests in about two minutes and hammers Outlook. So **92% of the tests need nine seconds and
no mailbox**; the remaining 8% cost two minutes and a real Outlook. Verifying with the narrow filter
plus a both-framework build of `OutlookAI.Core` is the standing bar until the VM can run the rest.

**Target arrangement: VM by default, plus a pre-release gate against the real profile.** The gate is
not "run everything"; a gate that repeats what the VM already proved is a slow ritual. It must
specifically cover what only a real profile can show - the latency-sensitive constants, the
delegate-store paths, real transport, and real message-class diversity. What belongs in it is being
derived; the analysis lives in the session trace folder as `vm-coverage-analysis.md`.

**Known losses, to be quantified by that analysis rather than assumed:** a local data file served
~1,200 items/second in this session's measurements where Exchange served ~12, so every budget,
timeout, breaker and hang detector validated on the VM is validated against latency two orders of
magnitude wrong - the 180 s sweep budget exists BECAUSE of the 12/s figure and the VM would have
said 5 s was ample. Delegate stores are indexed without folder nesting, which no local data file
reproduces. And at least four of this session's findings came only from real data: the census
timing out on a slow delegate store, the sweep budget needing a real store count, the frame
high-water measured from real bodies, and H3 answered by counting 43,048 real items.

**The corpus must grow to match**: message-class diversity (NDRs, read receipts, meeting requests,
sharing invitations, posts), items with no delivery time, very large bodies, deep folder trees, a
folder that fails to open, a store whose display name cannot be read. Several of those are shapes
this project has FIXED BLIND in the last few days, with no test able to produce them.

## 2. Timeout values - SHIPPED in `4502c92`

| Constant | Was | Now | Derivation |
| --- | --- | --- | --- |
| `ExhaustiveTimeBudgetMs` | 105 s | **600 s** | `ExhaustiveScanDeadlineMs - ResultReturnHeadroomMs` |
| `ExhaustiveScanDeadlineMs` | (did not exist) | **615 s** | new `ComHostOperationClass.ExhaustiveScan` |
| `SweepBudgetMs` | 30 s | **180 s** | measured: ~12 s per store x 5 stores, x3 headroom |
| `SweepWorkBudgetMs` | (did not exist) | **165 s** | `SweepBudgetMs - ResultReturnHeadroomMs` |
| `ThreadWalkBudgetMs` | 30 s | **180 s** | `= SweepBudgetMs`, unchanged expression |
| `SearchIndexTimeoutSeconds` | 15 s | **60 s** | with `OleDbIndexClient.DefaultCommandTimeoutSeconds` 30 -> 60 to stay its ceiling |
| `SearchBudgetMs` | 45 s | **240 s** | `SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs` |
| `OperationDeadlineMs` | 120 s | **300 s** | the hang detector; 4.5x the slowest healthy operation measured |
| `ConnectDeadlineMs` | 90 s | **180 s** | 60% of the operation deadline, the ratio it already had |
| `MoveBatchBudgetMs` | 120 s | **240 s** | 80% of the operation deadline, strictly below it |
| `MinimumItemBudgetMs` | (did not exist) | **1 s** | floor below which a batch item is not attempted |
| `ResultReturnHeadroomMs` | 15 s | **15 s** | deliberately NOT scaled - it covers one answer's size, not the budget's length |
| `HealthProbeDeadlineMs` | 5 s | **5 s** | unchanged, deliberately |
| `HandshakeBudgetMs` / floor | 30 s / 10 s | **30 s / 10 s** | unchanged as values; the floor now yields to a caller-declared budget |

**The blocker is gone**, and it is why this had to be one pass: `BudgetCompositionTests` asserts
`SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs` fits inside `OperationDeadlineMs`. The sum is now
240 s inside 300 s. Moving the sweep alone would have failed that test before anything reached a
mailbox.

**The sweep budget is measured, not preferred** (measurement delivered mid-pass, 2026-08-19). One
PST outside the local index, 20,000 items across the four arrival-path folders with real received
dates, 1,612 of them inside the seven-day fallback window so the 200-per-folder cap engages: four
sweeps of that ONE store took 13.6 s, 11.8 s, 10.7 s and 11.9 s. The maintainer's profile mounts
five stores, so the extrapolation is **~60 s against a 30 s budget** - the direct explanation for
the sweep timeout seen on the real profile. 180 s is 3x that, and the margin is headroom rather
than luxury because the corpus is a fast local PST and Exchange is slower per item. A second figure
from the same run matters to a constant this pass did NOT change: the frame high-water from one
store's sweep was 10,734,599 bytes (~13.5 KB per item over 758 items), so five unindexed stores
extrapolates to ~54 MB against the 64 MB frame limit - `SweepBodyBytesBudget` (32 MiB) bites first,
which is exactly its design intent. The 432 KB previously measured on the real profile was bounded
by the old 30 s timeout, not by any item cap, so raising the budget is what makes the body budget
load-bearing. Both figures are now in `Docs/magic-numbers.md`.

**Why `HealthProbeDeadlineMs` stays at 5 s** (an autonomous call, per the maintainer's "you decide
in context"): it is the diagnostic run precisely when Outlook is wedged. A health check that also
takes minutes turns every generous budget elsewhere into an unbounded wait with no way to find out
why.

**A verification gap of mine, found and closed.** `4502c92` shipped a type that does not exist on
net48, leaving CI red on master, and my verification did not catch it because I only built the test
project - which targets net10 alone, while CI builds `OutlookAI.Core` explicitly for both. Fixed in
`db34923`, and my own check now builds Core for both frameworks before any commit. Worth knowing
because it means every commit before `db34923` this session was verified against a weaker bar than
I stated at the time.

**Did NOT run the sort probe against the production mailbox, though it is written and ready.**
`T2/LiveTableSortProbeTests` would settle the largest open question of the session - whether the
freshness sweep's sort has ever worked - and the probe itself is strictly read-only. What stopped
me is its fixture: it builds the store-count tripwire baseline, and that tripwire was **rewritten
this session and has never once executed**. Exercising a rewritten safety guard against a real
mailbox, unsupervised, while its owner is asleep, to gain a measurement obtainable later at no
risk, is the wrong trade. Run it with:

    dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj --filter "FullyQualifiedName~LiveTableSortProbeTests"

Its final ANSWER line covers all four outcomes, and refuting the hypothesis is as useful as
confirming it. The same reasoning applies to `T2/LiveResumableScanTests`, which is the acceptance
no stand-in can give. **Both are why the live-tier-on-the-VM work is now the unblocking item**
rather than one queued item among several: on the VM there is no production mail to protect and no
maintainer activity to confuse the tripwire with.

## 3. Decisions taken autonomously - REVIEW THESE

**H3 measured to zero, and deliberately not fixed.** DASL *can* express absence: Microsoft documents
`IS NULL` and documents it specifically as the way to test whether a date property has been set. The
negation route that looked trivial is documented NOT to work - MAPI says a restriction over a
property that does not exist has *undefined* results, so `NOT (x >= floor)` negates an undefined
value. So the fix was available. It was not taken, because the population it would rescue was
measured at **zero**: a read-only probe over the maintainer's five stores and twenty arrival-path
folders counted **43,048 items, none lacking a usable date**. The fix is not free either - an item
with no date never leaves the sweep window, so it would be re-selected on every sweep forever, 200
items per folder per search. Paying a permanent cost to rescue an empty set is the wrong trade.
**What this does not establish:** that profile is entirely Exchange-delivered mail and mounts no
PST, and H3's hypothesised shape is imported, copied or restored mail. The row stays open,
downgraded, with the measurement attached. Evidence and the reusable probe are in the session trace
folder under Downloads (`tmp-aitrace/h3-measurement.md`, `h3-probe.ps1`).

**Two corrections the repo owed, from the same research - BOTH APPLIED 2026-08-19.**
`ExhaustiveDaslFilter` and `QUESTIONS.md` now say MAPI documents the result of a restriction over an
absent property as **undefined** rather than "silently excludes", and both name the inference that
wording invited - that `NOT (...)` therefore admits such a row - because that is the reasoning that
made a broken fix look viable. The gap map's H3 row now cites `SweepFolder` by symbol. **Checking
the rest turned the second correction into a systemic finding:** *every one* of the map's nine
`OutlookComSession.cs` line references had drifted off its subject, by between 60 and 900 lines,
and several `MailService.cs` ones had too (`:220` lands in an unrelated comment, `:612` and `:3995`
on bare `</summary>` lines). All nine `OutlookComSession.cs` citations are converted to method
names, the map carries a standing "cite by symbol" rule, and the remainder is in `TODO.md` to be
converted as each row is next touched.

1. **Overrode the corpus generator's date-fidelity refusal** with `--allow-undated`, reasoning that
   an all-recent corpus is still the sweep's worst case. **Wrong, twice over.** Undated items are
   invisible to a date-restricted sweep rather than recent, and separately the items were not in
   swept folders at all. Both faults are fixed in `6490ba9`; the refusal message that invited the
   bad inference is corrected and pinned. Cost: one wasted 12-minute build. No lasting damage.
2. **Reverted the VM to `CP-06-PRE-CORPUS` and deleted `CP-07-CORPUS-40K`.** The first corpus had
   40,000 items in Drafts with no dates. Reverting discards them in seconds where teardown would
   take a long time and would exercise a path unrelated to the measurement. `CP-07` snapshotted that
   broken corpus, so keeping it would only invite a later measurement against it.
3. **Built the corpus at 40,000 items** rather than a larger figure. The sweep caps at 200 items per
   folder, so beyond a few thousand per folder the sweep cost stops changing; scan throughput
   benefits from more, and the generator is additive, so growing it later is cheap.
4. **Ran read-only measurements against the maintainer's real production profile** (searches and
   health only, nothing created, moved or deleted) to size the budgets. Justified as read-only under
   the standing rules. It produced the finding that a 60-day exhaustive scan reaches 3 folders of 32
   before expiring.
5. **Kept `HealthProbeDeadlineMs` at 5 s** while raising everything else - see section 2.
6. **Did not investigate the 5,532 Outbox items** beyond confirming the profile could not send. The
   guard now makes the mechanism unreachable; understanding it is archaeology on a broken build.
7. **Accepted the corpus generator's broader safety rule** - refusing unless the profile has **no
   accounts at all**, with no override flag - rather than the narrower "no account can send from
   this store" I originally asked for. The object model cannot express the narrower rule, so the
   narrower version would read as a proof without being one.

8. **Gave the exhaustive scan its own deadline CLASS, and then passed that deadline explicitly
   from `MailService` as well.** The class alone is not enough and the reason is mechanical:
   `RemoteComGateway.Run<T>(operation)` bounds the whole lambda with an AGGREGATE equal to the
   ordinary operation deadline, and `EffectiveDeadlineMilliseconds` clamps each call to what is
   left of it - so a 615 s class deadline would have been clipped straight back to 300 s. The
   exhaustive lambda makes exactly one contract call, so the aggregate and the call deadline are
   the same number, and the call site passes `ComOperationBudgets.ExhaustiveScanDeadlineMs` with
   `allowConnectFloor: true`. The class is still load-bearing rather than decorative: it is where
   the number lives, `DeadlineFor` returns it, T1 pins that the two agree and that the class
   deadline exceeds the ordinary one. **The cost:** with an explicit budget, the
   `OUTLOOKAI_COMHOST_DEADLINE_MS` test override no longer shortens the exhaustive path. No CI
   test uses that override on an exhaustive scan today, so nothing broke; a future one would have
   to pass its own budget.

9. **Made the move/archive batch budget a real aggregate, not just a smaller number.** The decided
   change was 240 s "strictly below its deadline", which fixes the T1 assertion. It does not fix
   the arithmetic: the check runs BEFORE each item and each item was a fresh gateway call with a
   full deadline of its own, so the batch could still overshoot by a whole extra deadline (240 +
   300 = 540 s). Each item is now dispatched with what is LEFT of the batch budget, with
   `MinimumItemBudgetMs = 1 s` as the floor below which the item is reported "not attempted"
   rather than refused by the COM host's dispatch floor as a bare timeout, and a per-item
   `TimeoutException` is caught and reported per item saying the move's outcome is UNKNOWN. The
   items opt into `allowConnectFloor: true`, the same way the sweep does, so the first move on a
   fresh host still gets its cold-start allowance. **Review this one**: it touches the move path,
   which is a mutating path, and it was not literally what was asked for.

10. **Fixed the handshake floor by rule rather than by class.** The inventory suggested letting a
    `HealthProbe`-class operation use its own deadline. That would not have fixed the actual
    defect: `outlook_health`'s gateway calls are `GetStoreDetails`, which classifies as
    `Operation` with an explicit 5 s budget, so the class is the wrong key. The rule used instead
    is the one already written two lines away for the cold-start CONNECT floor: an explicit
    caller budget outranks a floor, with `allowConnectFloor` as the opt-out for the sweep. Cost:
    on a genuinely cold host, `outlook_health`'s first call now has 5 s to start the child and
    may report the COM half as unreachable. That is health degrading, which is health working;
    the next call (no explicit budget) starts the child with the full floor.

11. **Raised the two live-tier bounds that this pass would otherwise have broken**, though neither
    was in the task. `LiveDisconnectRecoveryTests.ScenarioClock`'s per-step budget was 180 s with
    a remark that this was "1.5x `OperationDeadlineMs` (120 s)"; at 300 s the remark was false and
    the bound would have abandoned COM calls the shipped product is still waiting on, so the step
    is now `max(LiveInboxArrival.DeadlineSeconds, 1.5x OperationDeadlineMs)` and the prose says
    so. `Phase3LiveMcpToolShapeTests` ran a whole stdio session, ending in an exhaustive scan, on
    a flat six-minute budget; it now derives that from `ExhaustiveScanDeadlineMs`. Neither is in
    the CI suite, so neither shows up in the test count.

12. **Split the sweep's budget expiry into its own coverage code (`sweep_time_budget`) rather than
    reusing `time_budget`.** The existing code belongs to the folder-scoped subtree walk's 2 s
    bound and its advice says "scope narrower, or pass include_subfolders:false" - advice that
    cannot be acted on over a default-folder sweep, which is shallow by construction. Same rule
    the codebase already applied when it split `item_cap` from `item_cap_unsorted`: two bounds
    with different remedies need two codes. The flag is deliberately NOT attributed per store,
    because the stores the sweep never reached are exactly the ones with no per-store entry to
    attribute it to.

13. **Wrote the in-process budget as a `DispatchProxy` over `IOutlookSession`.** It is the only way
    to check a clock between contract calls without hand-writing 26 forwarding methods. The known
    hazard is that reflection wraps everything in `TargetInvocationException`, which this
    repository has already paid for once - a reflective layer on the COM-host path flattened every
    deliberate error into "Exception has been thrown by the target of an invocation", breaking both
    the tool layer's advice and `ComGateway`'s disconnect rebuild. Failures are re-thrown through
    `ExceptionDispatchInfo`, and T1 asserts a `COMException` crosses the proxy with its type,
    message and HRESULT intact.

14. **Wrote the `send_outcome_unknown` audit line as best-effort**, unlike every other write on the
    send path, where a failed append refuses the operation (D4: no send without its line). The
    operation this line describes has already happened or already failed; throwing an audit error
    over it would replace the one message the caller most needs with a message about a log file.

15. **Did not add a stop-request protocol**, per the decision - and the reason is worth keeping
    where the next reader will find it: `ComHostServer.ServeAsync` calls `Invoke` synchronously
    inside its read loop, so while wedged the child is not reading the pipe and a stop frame is
    structurally undeliverable rather than merely slow. That is now written at `KillChild` and in
    `McpServer/Docs/com-host.md`.

16. **Two things this pass changed are provably unguarded by any non-live test**, established by
    mutation rather than assumed: reverting `ComGateway`'s budget overload to a pass-through, and
    removing the 250 ms clean-exit wait, both leave the whole suite green. Recorded in `TODO.md`
    with the options rather than papered over.

17. **Put the `update_draft` intent record in the PARENT PROCESS, in memory, and not on the draft.**
    The decision the maintainer gave was (d) - record intent first - and it left open where. Three
    candidates were weighed. A property on the draft travels with the item and a second server
    instance can see it, which is genuinely more than parent-side state offers; it was rejected on
    three grounds, of which the first is decisive. **It can only be written by the process that
    dies**, so it is exactly as unreliable as the thing it records - and the failure being defended
    against is precisely that this process is terminated between two COM calls. It also needs a
    SECOND mutation to clear it, which is one more window for the same kill, and it writes server
    bookkeeping onto the user's own mail, which this product does not do anywhere else (the same
    instinct behind soft-delete-only discard and bit-identical signatures). The audit log was
    rejected as the *mechanism* for the opposite reason - it is append-only and its documented
    ordering is mutate-then-record, so it can state that an outcome is unknown but cannot be read
    back as resumable state; it now gets an `update_draft_outcome_unknown` line anyway, on the same
    reasoning as `send_outcome_unknown`. `ServerDraftRegistry` was rejected as the *home* because it
    is an allowlist of ids and would have had to grow a second meaning. So: a new
    `Services/DraftUpdateIntents`, per-process and unpersisted like `ServerDraftRegistry` and
    `SendConfirmationTokens`, on their stated rule - **a restarted process never observed the draft
    and must not claim it can complete an update it cannot vouch for**. When the record is gone the
    caller gets the pre-existing answer (outcome unknown, look before acting), which is a smaller
    guarantee and never a wrong one.

18. **The idempotence key is DERIVED BY THE SERVER, not supplied by the caller** - deliberately the
    opposite choice from the send path's confirm token, and the reason is that the two mechanisms
    want opposite things. The send token is caller-supplied *because* its purpose is friction: a
    human has to say yes. Re-entrancy has to work when the caller does nothing special, because the
    caller that most needs it is an agent re-issuing the call that just failed - and a caller-supplied
    key would additionally let two DIFFERENT requests claim one identity, a worse failure than the
    one it prevents. The key is SHA-256 over the draft id plus every argument that reaches Outlook,
    canonicalised presence-first so an OMITTED list and an EMPTY list cannot hash the same (they mean
    "leave alone" and "clear"). **Identical arguments alone are NOT a retry**, which was the sharpest
    question in the task: only a request whose outcome is still unknown is resumable, a call that
    ANSWERED settles its record, and any other update to the same draft drops it. So the only way
    into resume mode is to re-issue exactly the call that was interrupted, before anything else
    intervened.

19. **The replayability classification, established from the code, and what it forced.** Reading the
    ~20-call sequence rather than assuming: body/signature replace (replayable - it rewrites the
    draft region), recipients replace per class (replayable - clear then set), subject, importance
    and read-receipt (replayable - assignment), `Save`/`Display` (replayable). **Attachment ADD is
    the only accumulating step**; attachment REMOVE by name is idempotent in itself and destructive
    only in sequence. And one row nobody had named: **the conversation-index restore is
    ORDER-COUPLED and NOT replayable**, because its input is captured live and is destroyed by the
    subject write it compensates for - a repeat reading it live would faithfully restore the value
    the interrupted attempt had already regenerated, report `conversationTopicPreserved: true`, and
    leave the draft out of its thread. That row is what makes the pre-image load-bearing beyond
    attachments.
    **The design that follows: converge on the END STATE, do not replay the STEPS.** Progress cannot
    be recorded, because only the process that died knew it; the pre-image plus the draft's current
    contents can always be compared. `Com/DraftAttachmentPlan` does that as pure logic, and a FIRST
    attempt is its identity case (before == now reduces it to remove-every-match plus add-everything,
    which is exactly the two loops it replaced) - so the re-entrant path is not a second mode with
    its own semantics. **It is also indifferent to whether a partial attempt persisted at all**,
    which matters because Microsoft documents neither outcome for an unsaved change when the
    automation client dies: both are just states the draft can be observed in.
    **One case is deliberately NOT resolved and is redone instead:** a name that is both removed and
    added (a replace). The old copy and the new copy are the same name, so no pre-image can tell
    "nothing applied" from "both halves applied". The plan deletes every current copy and re-attaches
    every requested path, which converges from all three reachable states because the source is a
    file on disk. Cost: repeated work in a rare path. Reviewed and accepted.

20. **Did the reorder too (the (b) that (d) does not exclude), because it was close to free.**
    Additions now run BEFORE removals, and removals delete the N LOWEST-INDEXED instances of each
    name, counted before anything was attached. Replace semantics survive because `Attachments.Add`
    appends, so the pre-existing copies are exactly the low-indexed ones. The gain is the shape of
    the danger window: with removals first it holds a draft that has LOST the user's file; with
    additions first it holds a duplicate, which is visible and undoable. **No intermediate `Save()`
    was added** - that would deliberately CREATE durable partial states that today may not exist.

21. **Corrected a message that was simply untrue, for `update_draft` only.** `BuildDraftRefusal`'s
    catch-all `com_failure` branch told the caller "Nothing was changed". That holds for every NAMED
    refusal beside it - each is decided before anything is written - and does not hold for the
    catch-all, which wraps the whole ~20-call sequence and can fire after the body has been committed
    through the inspector or after an attachment has gone. It now says the outcome is unknown and
    points at the repeat, and its intent is kept pending rather than settled. **`discard_draft`'s half
    of the same branch was left alone**: it is a different sequence, the item is recoverable from
    Deleted Items either way, and changing a deletion path's wording without a live run to check it is
    not a trade worth making unsupervised. Recorded in `TODO.md`.

22. **Fixed a build break that was already on master, because CI builds the file it is in.**
    `dotnet build McpServer/OutlookAI.Core/OutlookAI.Core.csproj` fails for the **net48** target at
    HEAD `23dca4f`: `BudgetedSessionProxy` (added the previous day) uses `DispatchProxy`, which is not
    in net48's default reference set, and `Stopwatch.GetElapsedTime`, which is net7+. The
    `mcpserver.yml` workflow builds that csproj explicitly, so the branch is red regardless of this
    work; the test project only builds net10, which is why the suite stayed green and nothing noticed.
    Fixed with the `System.Reflection.DispatchProxy` package for net48 and the elapsed-time arithmetic
    written out. **This is the net48 gate doing its job** - it exists so Core cannot acquire a
    dependency the v3.1 event host could not take. The new code in this pass was written net48-clean
    for the same reason (no `SHA256.HashData`, no `ArgumentNullException.ThrowIfNull`, no `Math.Clamp`).

23. **The pre-image read doubles as the store resolver, which removes a mutating fan-out.** Taking it
    costs at most two extra READS and only when the request contains something a blind repeat could
    get wrong - a subject change (conversation index/topic) or files to attach (attachment names); a
    body-and-recipients update pays nothing. Because that read resolves which store holds the draft,
    the mutating cross-store loop below it no longer fires in the ordinary case: a bare EntryID in a
    non-default store used to be found by offering `TryUpdateDraft` to each store in turn. The loop is
    kept as the fallback for a draft no read could find. **If the pre-image cannot be taken at all, no
    intent is recorded** and the caller gets the old "do not retry" advice - a resume this server
    cannot vouch for is worse than no resume.

24. **The gap map's line numbers are systemically stale, not just H3's.** Checked all 40 citations:
    every one of the nine `OutlookComSession.cs` references had drifted off its subject, by between
    60 and 900 lines, and `MailService.cs:220` / `:612` / `:3995` land in unrelated comments or on
    bare `</summary>` lines. All nine are converted to method names, the map gained a standing
    "cite by symbol" rule with the measurement attached, and the remainder is in `TODO.md`.

25. **Method note, recorded because it cost real time and would repeat.** Two mutation passes were
    run concurrently by mistake. Each saves the file it is about to mutate and restores that copy
    afterwards, so the second pass captured the FIRST pass's mutation as its "original" and restored
    it permanently - re-applying a reverted decision line silently, in a tree that still built and
    still passed 1,919 tests because the mutation it re-applied is one only a resumed call can
    observe. Caught by re-checking every anchor before the third run rather than by any test.
    Mutation passes are serial by construction; the script now verifies its anchors first.
    Separately, **six NUL bytes** had been written into `DraftUpdateIntents.cs` where single-space
    string literals were intended, and the canonicalisation was rewritten presence-first so no
    sentinel value exists to be corrupted in the first place.

26. **Mutation check: 15 decision lines reverted, 14 caught, 1 not.** Each was reverted, built,
    run against the whole non-live suite, and restored. Caught (with the test that noticed):
    ignoring the pre-image on a repeat; preferring the live conversation index over the recorded
    one; recording the intent AFTER the mutating call (2 tests); a successful update not settling
    its intent; a new intent not dropping the draft's other pending ones; an unclassified COM
    failure settling the intent; the killed-update message never offering the repeat; the plan not
    reporting what an interrupted attempt already removed; a replaced name planned as an ordinary
    addition; a stale intent never expiring; an omitted list hashing like an empty one; the
    pre-image read taken for every update; a draft no store could open treated as an unknown
    outcome; a discard not dropping its pending pre-image. **NOT caught: the guard that skips
    enumerating attachments on an update that touches none** - a COM-side cost guard with no
    observable payload; recorded in `TODO.md` rather than papered over. Four of the fifteen had to
    be re-expressed because `if (false)` is constant-folded and `TreatWarningsAsErrors` turns the
    resulting CS0162 into a build failure - a mutation that does not compile proves nothing, so
    they were rewritten as runtime-false conditions.

27. **F2's four open questions, decided as the maintainer recommended - and a fifth they did not
    ask about, which I added.** The design document's section 8 left four open; the maintainer gave
    their preference on each and invited me to overrule with reasons. I overruled none, and the
    reasoning below is mine rather than a restatement, because "I agreed" is not a record.
    - **Q1 - `top` is PER PAGE, with `itemsReturnedTotal` as the chain total.** Per chain was the
      only real alternative and it defeats the thing it appears to serve: a five-page scan that
      returns 100 items in TOTAL cannot deliver completeness at all, which is the whole point of
      resuming. Per page matches `list_folders`' `offset`, and the accumulated context - the actual
      cost `top` = 100 exists to bound - is made VISIBLE rather than bounded, so an agent stops
      deliberately instead of discovering the bill afterwards. The `scan_resumed` sentence states
      the total in words as well as in a field.
    - **Q2 - a superseded token is REFUSED, and the refusal carries the position.** Keeping the last
      K tokens live was the tempting option and it is quietly wrong: the chain's suppression set has
      already advanced past the replayed position, so honouring the replay would suppress exactly
      the rows the replay exists to return. Fixing that needs a snapshot of the suppression set per
      position, which is a lot of memory for a rare case. Refusing costs nothing ONLY because the
      refusal names the folder and the date to continue from, in `folder` and `before` - parameters
      that already exist - so a lost response costs one round trip rather than the whole scan.
    - **Q3 - a resumable stop still sets `degraded: true`.** The answer in hand is incomplete, and
      `degraded` is the flag the tool description tells an agent to relay to a human. Dropping it
      because a remedy exists is the "looks complete and quietly is not" failure the whole
      coverage-code system was built against, and it would be invisible - the payload would gain a
      field and lose a flag. No `resumable: true` was added either: a second boolean before anything
      branches on it is how a payload grows fields nobody reads. `nextToken` present or absent
      already says it.
    - **Q4 (the document's) - the `Table.Sort` fix is NOT taken, only the probe and the split.** The
      document recommended probe-then-fix and the maintainer's Part 2 said the same in stronger
      terms: write the test, do not run it. Changing the sort call before the probe runs would
      destroy the evidence, so the call is untouched.
    - **Q4 (the maintainer's) - `stopReason` is RECORDED where the walk stops, not derived.** This
      is the one they stated instead of the document's Q4, and it is right for a mechanical reason:
      `truncated` and `timedOut` are independent and both can be true, so any derivation picks one
      by accident of which `if` came first - and the two remedies point in opposite directions (a
      budget stop means "keep resuming, there is no cheaper route", a cap stop means "keep resuming,
      or narrow, and narrowing is cheaper"). `depthLimitReached` is excluded from the vocabulary
      entirely, because the depth guard never ends a walk.
    - **The fifth, mine: `scan_resumed`.** Not in the design and not asked for. Without it the LAST
      page of a chain - which by itself covered everything it was asked for - reports
      `stopReason: complete`, no coverage code, `freshness: "live"` and no `degraded` flag, and an
      agent relaying that page tells the user the search was complete when it saw a hundred of
      several thousand matches. **Q3's decision is unenforceable without it**, because `freshness` is
      recomputed from the codes rather than from a boolean. It also gives the other four resumption
      codes their footing: each is only reachable on a page that already carries `scan_resumed`, so
      none of them weakens `degraded` on its own.

28. **Where the walk state lives, and the two places it deliberately does not.** Server parent, not
    the COM child and not the wire. **Not the child**, for a mechanical reason: the failure this
    design exists for is a scan running past its deadline, which ends with the supervisor killing
    that child - state kept there would be destroyed by exactly the event that makes resumption
    necessary. **Not the wire**, because proving no folder was skipped requires the SET of finished
    folders, and at ~140 characters per EntryID that is kilobytes per page in each direction,
    against a standing decision that payload is context and context is the scarce resource. A
    self-describing token would additionally have to give up duplicate suppression and
    added-folder detection, which are the two failures the brief names. The one thing parent-side
    state cannot survive - an MCP-server restart - is covered without any state at all:
    `exhaustive.position` carries the resume folder and date in plain fields, so the caller
    continues with `folder` and `before`.

29. **Two passes over the folder tree per page, decided against one.** The scan enumerates its scope
    in the stable order first (structure reads only, no `GetTable`), then walks that list, re-opening
    each folder by EntryID. A single combined walk is one fewer traversal and cannot produce three
    things: an honest `foldersTotal` when the walk stops after four folders of thirty-two, index
    arithmetic instead of recursion state for resumption, and any vantage point from which "a folder
    appeared BEFORE the cursor" is visible - which is the detection the whole server-side-state
    argument rests on. **The cost is one `GetFolderFromID` per scanned folder** (32 on the
    maintainer's store) plus a second structure traversal, against a per-folder `GetTable` that
    dominates both. Recorded because it is a real cost nobody asked me to pay.

30. **One number in this design is unmeasured and it is the one that decides whether the design is
    cheap.** Re-opening a resumed folder's table - `Folder.GetTable(filter)` with a date restriction
    over 108,144 items - is measured nowhere. Sensitivity: at 10 s it is roughly 8% overhead across
    the Archive's ~15 pages; at 60 s it is roughly 50%, and the design would then want fewer, longer
    pages. `T2/LiveResumableScanTests` records wall clock per page, so one live run measures it as a
    side effect. I did not reopen the 600 s budget or the `top` ceiling over it; I recorded the
    sensitivity instead.

31. **The sweep's refused-sort counter is my call, and it is more than the brief asked for.**
    Splitting the `catch` alone makes the two failures distinguishable IN THE CODE and answerable by
    nobody, since neither half reaches a payload - so "answerable from telemetry rather than from a
    probe" would still have needed the probe. `sweep.sortRefusedFolders` counts folders where the
    column WAS added and `Sort` then threw; equal to the folders swept, on every store, is the
    namespace-reference hypothesis confirmed from an ordinary search. It raises no coverage code,
    changes no advice and never degrades an answer - a diagnostic beside `rowsDropped`, not a hole.
    The sort CALL is untouched, deliberately.

32. **A third copy of the sibling comparator, found by the mutation pass's own anchor check.**
    `ListFolders` sorted STORES with a hand-written copy of the same name-then-index comparison the
    sibling sort used, so "stable order leg 1" and "leg 2" were two copies of one rule. Both now call
    the shared `OutlookComSession.CompareSiblings`, which is also what T1 pins. Worth recording for
    how it surfaced: the mutation script refuses to run unless every anchor is unique, and the
    duplicate is precisely what that check reported.

33. **Two live tests written and NOT run**, per instruction, both read-only.
    `T2/LiveResumableScanTests` holds F2's real acceptance - a scan paged at `top: 2` must return the
    same EntryID set as one unpaged run, with no duplicates - which no stand-in can prove, because a
    stand-in returns whatever the test tells it to. `T2/LiveTableSortProbeTests` settles the sort
    question: one table per store, both property spellings, column add and sort caught separately,
    the first row's date read in each case, a printed per-store verdict and one ANSWER line. It
    asserts only that it RAN, because every other outcome answers the question - including "neither
    spelling sorts", which refutes the hypothesis as usefully as confirming it would.

34. **Mutation check: 18 decision lines reverted, 16 caught on the first pass, 2 not - and one of
    those two was a real hole in my own tests, now closed.** Each was reverted, built clean, run
    against the whole non-live suite, and restored; the pass is serial by construction and refuses
    to start unless every anchor is unique.
    - **NOT CAUGHT, and fixed: `M01` - the request-fingerprint comparison.** Making a mismatched
      resume resolve as `Valid` left all 1,949 tests green. That is the worst shape a coverage gap
      can have: the fingerprint itself, the argument diff and all five refusal messages were pinned,
      so the area LOOKED covered while the one line that decides whether a continuation answers the
      question it started was unprotected. Three tests now cover it - the store-level comparison, the
      same refusal end to end through `MailService` (including that the chain survives it, so a
      mistyped argument does not throw away the minutes page one cost), and a `resume_token` passed
      WITHOUT `exhaustive`, which was being silently ignored and is now refused. Re-run after: CAUGHT.
    - **NOT CAUGHT, and recorded rather than papered over: `M18` - the sweep's refused-sort
      counter.** It lives in `SweepFolder`, behind a COM call no non-live test can execute. Exactly
      the same class as `sortApplied` itself, which `TODO.md` already carries as live-only, and the
      same cheap substitute applies (a temporary build that forces the branch). Re-run after the new
      tests: NOT CAUGHT, as expected.
    - Caught, with the guard that noticed: superseded token refused; expired refused AS expired
      rather than as unknown; malformed handle its own answer; a resumable stop still degrading the
      answer; `stopReason` recorded rather than derived; a resumed page saying so; `scan_resumed`
      raised; a vanished cursor folder raising `tree_changed`; the resume date bound inclusive;
      sibling order breaking ties by collection position; `folder` inside the fingerprint; the chain
      total accumulating; the store evicting at capacity; a stop with no token saying so instead of
      looking complete; a finished chain releasing its state; the final page reporting the whole
      chain's total.

35. **Method note, and it is the same hazard section 25 records wearing different clothes.** To
    re-run two mutations I wrote `import mutate` - which EXECUTES the module, so it restarted the
    whole 18-mutation pass. The tool call was killed at its 10-minute limit, and a killed process
    does not run its `finally`, so **one mutation (`M05`, suppressing `degraded` when a token
    exists) was left APPLIED to the working tree**. It was found by scanning the tree for every
    mutation's replacement text rather than by any test - the suite was green at the time, because
    the pass had not reached the test step. Two things now guard it: the subset runner is a separate
    script that never imports the full one, and the anchor check is run in BOTH directions
    afterwards (every original present exactly once, every replacement absent). The general rule
    worth keeping: a mutation script must be safe to KILL, and the only way to know a tree is clean
    after one is to check it rather than to trust the `finally`.

36. **The live tier's split is two TRAITS, not a filter string somebody has to remember - and
    the classification is checked in CI.** The brief said not to narrow the tier quietly, and a
    hand-maintained list of "tests that work on the VM" narrows it the first time somebody adds
    a test and forgets. So every live test now carries `LiveTier` (`Portable` or `ProfileBound`,
    exactly one) and, when ProfileBound, at least one `Requires` value naming a capability a
    test machine cannot have: `SearchIndex`, `MailAccount`, `Transport`, `MultipleStores`,
    `DelegateStore`, `SmallHubStore`, `ProbePopulation`. Two further values,
    `InteractiveDesktop` and `AddInRegistry`, describe things a test machine CAN have and
    constrain only how a run is launched. `T1/LiveTierInventoryTests` reads the assembly's own
    attributes and fails when a live test has no tier, an unknown tier, a ProfileBound test with
    no reason, a Portable test claiming a production-only capability, or a `Requires` value
    outside the vocabulary. **The point is that reclassifying a test to make the VM subset look
    bigger now requires deleting its reason, which is a visible act.** Counts today: 115 live
    methods, 19 Portable, 96 ProfileBound, verified by `--list-tests` (discovery only, nothing
    executed).

37. **Only 19 of 115 can move, and the reason is structural rather than a matter of
    configuration.** A profile with no mail accounts cannot create a draft at all -
    `NewDraft` resolves an Account object by SMTP address and refuses when none matches - which
    puts the whole draft, update/discard, HTML-draft and send families out of reach whatever
    else is arranged. On top of that, `testHubStoreDisplayName` doubles as an SMTP address in
    several tests (`to: Hub`, `FindAccountBySmtp(Hub)`), which a PST display name can never
    satisfy; and several tests assume a hub small enough to page in one request, which a
    20,000-item corpus is not. **This is the open question I could not decide** - see the end of
    this section.

38. **Found by reading, and fixed: a filtered live run took a store census and never compared
    it.** `LiveStoreCountTripwire.Verify` was called from exactly one place -
    `LiveLifecycleFixture.Dispose` - on the strength of `SuiteCollectionOrderer` forcing that
    collection last. True for a whole-tier run; false for every filtered one, because a filter
    that selects no LiveLifecycle test never constructs that fixture. So
    `--filter "FullyQualifiedName~LiveTableSortProbeTests"` - the exact command section 2 of this
    log hands the maintainer - would have taken the baseline, paid for it, connected to Outlook,
    and thrown the census away, reporting green with the guard silently absent. Filtered runs are
    not an edge case in the new world; they are the whole point of the VM.
    **The fix uses the one vantage point that can see the run's shape:** xunit hands the
    collection orderer the collections that survived the filter, before any fixture is
    constructed. It now publishes that list to a new `LiveTierRunPlan`, every guarded fixture
    calls `LiveStoreCountTripwire.CollectionFinished(...)` in its dispose, and the tripwire
    verifies when nothing guarded remains. When no plan was published at all the answer is
    `Unknown`, which verifies AND stays armed - a census per collection boundary, deliberately,
    because an unverified run is the one outcome that must not happen.

39. **Found by reading, and fixed: three live classes ran outside every guard.**
    `LiveMcpToolShapeTests`, `Phase3LiveMcpToolShapeTests` and `Phase7LiveMcpToolShapeTests`
    carried `Category=Live` and belonged to no collection, so xunit gave each an implicit one
    with no fixture - no census, no health preflight, no verification. Nine live tests against
    real mailboxes, one of which drives Outlook's UI and one of which runs a full exhaustive
    scan. They now share a `LiveMcpToolShape` collection whose fixture exists only to be that
    hook, and the collection orderer ranks it where those tests already ran (their implicit
    names sorted after every "Live" name), so bringing them under the guard does not silently
    move nine live tests to the front of an ordering that was arrived at by being bitten.

40. **Found by reading, and fixed: the tripwire's confirmation census could dismiss a real
    deletion as noise.** A suspected loss is re-censused and only what fails BOTH times is
    reported - and the comparison was `verdict.Failures.Intersect(second.Failures)` over the
    RENDERED message. Those messages carry the folder's before and after counts and how many
    items arrived, so a single mail landing between the two censuses changes the string and the
    failure drops out as "enumeration noise". On a real profile during a 27-minute run that is
    ordinary. Failures now carry a stable `Key` (kind, store, folder - never a tally) beside the
    message, and the intersection is on the key. **This is strictly stronger, never weaker:**
    equal strings imply equal keys, so everything confirmed before is still confirmed, plus the
    cases an arrival used to hide. The comment claiming the intersection required "the SAME
    items" is corrected to say what it now does.

41. **`machineProfile` in the live-test settings, so a second machine can be configured
    honestly.** `LiveTestSettings.Load` required `probeTerm` and a complete `subjectOnlyProbe`
    on every machine. Both name real mail: a word proven to hit this machine's search index, and
    a population whose term is in the subject and not the body. A test machine has neither, and
    a requirement that cannot be met honestly gets met dishonestly - somebody types a plausible
    value and the tests that read it fail somewhere far away from the mistake. The file now
    declares `Production` or `Portable`; the hub and the watched stores are required on both, the
    two real-mail blocks only on Production. **`Production` is the default**, so the maintainer's
    existing settings file keeps exactly the validation it was written under. A block that is
    PRESENT must still be complete on either profile: three fields out of four reads as
    configured and behaves as absent.

42. **Two tests that reported success having proven nothing now refuse to, on a Production
    profile.** `LiveStaleIndexRowTests` returns early when no `delegateNestedFolderProbe` is
    configured and `LiveManageSignatureTests.DefaultAssignment` when the hub account has no row
    in the signature registry - each writing a line to the test output and passing green. On the
    machine they were written for, absent means the machine or the settings have drifted, and
    green hides it. They now call `RequireProductionPopulation`, which throws on Production and
    no-ops on Portable, and the Portable path prints `PROVED NOTHING:` rather than `SKIP:`.
    **The two `ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain` tests were left alone** although
    they degrade the same way (the "three accounts" loop collapses to whatever the settings
    list): they are ProfileBound, so they do not run on a test machine, and changing a sweep
    assertion without a live run to check it is not a trade worth making unsupervised.

43. **Mutation check: 13 decision lines reverted.** Results in the report; the pass is serial by
    construction, refuses to start unless every anchor is unique AND no replacement is already
    present, and is verified afterwards by a SEPARATE script that checks both directions - which
    is the only thing that can be trusted after a kill, since a killed process skips its
    `finally`. The verifier caught its own false positive on the first attempt (a mutation that
    DELETES a line has a replacement that is a substring of the original, so it looks applied in
    a clean tree) and refused to start rather than proceeding, which is the right failure.
    **Deliberately not mutation-checked, because no non-live test can reach them:**
    `CollectionFinished`'s early return on `NotLast`, the keep-alive release and the latch in
    `Verify(final)`, and the key-based intersection inside `Verify` itself - all of them behind a
    COM census. The KEY is pinned; the intersection that uses it is not.

44. **The count tripwire's first run is predicted in `Docs/live-tier-on-the-vm.md` section 6, and
    the headline is that the VM as it stands cannot exercise it.** With one PST that is also the
    hub, `PlanFor` gives the hub a count-only plan and `Evaluate` exempts it, so the guard
    censuses, identifies nothing, and cannot fail. A second store is what gives it something to
    watch - and the numbers say what SHAPE that store wants. A 20,000-item corpus is the wrong
    shape: all four populated folders (Inbox 10,912 / Sent 4,964 / Deleted 2,467 / Junk 1,663) are
    above the 500-item per-folder limit, so every one falls back to counts and the 3,000-item
    per-store budget goes almost entirely unspent. **A small store of a few hundred items
    exercises the identity path completely**, which is the half that was rewritten.

45. **Which store is the hub is not a free choice, and I got it wrong first time.** The obvious
    layout - corpus as a watched non-hub store, small empty store as the hub - is backwards. The
    Portable subset is mostly scans and sweeps and every one of them targets the HUB:
    `LiveResumableScanTests` pages through it, `LiveExhaustiveSearchTests` bounds a scan to one of
    its folders, `LiveSweepScopeTests` sweeps its arrival-path folders. Against an empty hub they
    all take their "corpus too small" early return and report green having proven nothing - which
    is precisely the failure mode this whole pass exists to remove, reintroduced by the store
    layout. **So: corpus IS the hub, and a small bystander store is what the tripwire watches.**
    Caught by re-reading my own runbook against the test list rather than by anything failing.

46. **Where the shared helper lives, and why only the OPENING sentence is shared.** The audit
    recommended Direction C plus a scoped Direction A, and I took it without changing the shape.
    The helper is `Core/Com/MutationOutcome` - Core rather than ComHost, because `MailService`
    needs it as much as `OutlookTools` and `ComHostServer` do, and Core is what all three can
    see. It exposes the three-value vocabulary and four functions: `ForInterrupted` /
    `ForCompleted` (the field) and `DescribeInterrupted` / `DescribeAnswerLost` (the sentence).
    **Two sentences rather than one with a flag**, deliberately: they say opposite things - one
    means "nobody can tell", the other means "it SUCCEEDED and repeating it would do it twice" -
    and folding them into one function with a boolean is how that distinction gets lost by the
    third caller. **Only the opening is shared**: every site appends its own remedy, because
    update means repeat the identical call, send means check the Outbox, create means check
    Drafts, discard means check Deleted Items and move means find the item first. A shared
    sentence that tried to carry all five would be worse than the per-site prose it replaced,
    and specific remedies are exactly what makes `DescribeSendOutcomeUnknown` the good one.
    Each site calls the classification ONCE and branches its remedy on the returned value, so
    there is one decision per site rather than two that could disagree.

47. **`ComSessionOperations.IsRetryable` needed an operation NAME at three sites that had none,
    and I added an ambient trace rather than guessing.** Rows 3, 7 and 15 all key on the same
    classification, but only row 7 (`ComHostTimeoutException`) already carried the name.
    - Row 3: `ComHostResponseTooLargeException` gained an `Operation`, filled by the supervisor,
      which has it in scope where it rebuilds the child's error. The child-side message is fixed
      in `ComHostServer.TooLarge` at the same time, so both halves of the same claim agree.
    - Row 15: a bare `COMException` reaching the tool layer carries nothing. `ComHostRequestContext`
      now records the last dispatched contract operation for the current request, stamped by
      `RemoteSessionProxy` before each round trip. **The recorder is a mutable object held by the
      `AsyncLocal`, not a string**, and that is load-bearing: an `AsyncLocal` ASSIGNMENT made
      inside a `Task.Run` body does not flow back out to the awaiting caller, so a string would
      have read as null in exactly the `catch` that needs it. The reference flows in, the proxy
      mutates the shared object, `GuardAsync` reads it after the await.
    - **A null trace states NO outcome at all**, rather than defaulting either way. A request that
      dispatched no COM call may still have failed for a reason this server cannot classify, and
      answering `unchanged` there would be the same habit the whole audit is about.
    **Review this one**: it is new ambient state on a hot path, and the cost is one field write
    per contract call.

48. **A new exception type, `OperationOutcomeException`, so the SERVICE layer can state an
    outcome too.** The tool layer can classify what it catches; it cannot classify an
    `InvalidOperationException` that `MailService` threw deliberately with an accurate message of
    its own. Rather than sniff messages, the eleven such throws now carry the value: derived from
    `InvalidOperationException` so every existing `catch` still matches (the move batch catches
    that type per item), never crosses the COM pipe because every site is parent-side, and
    invisible to invariant 10 for the same reason. **The alternative I rejected** was attaching
    the outcome only where the tool layer could derive it, which would have left row 5 - the row
    with the real user harm - with no field at all.

49. **A failed create reports `unknown`, not `applied`, even when the draft's id is known.** The
    draft exists, so `applied` is arguable. I took `unknown` because the requested operation is
    "a complete draft, filed in Drafts, optionally on screen" and a failure part-way through
    demonstrably did not deliver that - and because an agent reading `applied` will read it as
    success. The message carries the id either way, which is what makes the difference
    actionable; the field only has to be safe.

50. **The `outcome` clause went on all 13 mutating tools' descriptions, and on none of the
    read-only ones.** One shared `OutcomeHint` const of ~190 units, appended rather than
    paraphrased per tool, because a paraphrased vocabulary drifts. The client's cut is per
    string with no per-tool bucket (measured 2026-08-18), so 13 copies cost nothing against any
    single budget: the largest description after the change is `update_draft`, and `search` -
    the one already near the cut at 1791 of 2048 - is read-only and untouched. **Not put in the
    `initialize` instructions**, although that would have been one copy instead of thirteen:
    that string is injected into EVERY session including ones where the tools are deferred
    name-only, and it is pinned verbatim by T3, so it is the wrong place for a contract that
    only matters once a mutating tool is actually called.

51. **`ComHostSupervisor.DescribeInterruption` was NOT rewritten.** It is the site the audit
    called the model, its remedy is the specific one, and T1 pins both halves. It gained the
    outcome VALUE (carried on `ComHostUnavailableException`) and nothing else. Reusing its shape
    was the instruction; rewriting its words would have been a regression dressed as
    consistency.

52. **Read `BuildNavigationError`, which the audit listed as the one function it left unread.**
    It is three branches. `StoreNotFound` and `FolderNotFound` are resolution failures decided
    before any window is touched and keep their claim. The catch-all is NOT: by the time it can
    fire, `EnsureVisibleExplorer` may have created and shown an Explorer, `CurrentFolder` may
    have been set, and for `show_search_results` `Explorer.Search` may already be running - the
    only thing left to throw is the state snapshot afterwards. So "Outlook could not show the
    requested view" was the same defect as the rest of the sweep, in the two tools nobody had
    checked. It is now "Outlook could not CONFIRM the requested view ... THE WINDOW MAY HAVE
    MOVED ANYWAY".

53. **The `FileInfo.Length` fix is a behaviour change nobody asked for, and it is the right one.**
    `TrySaveAttachment` measured the saved file inside the same `try` as the save, so an
    `IOException` from the size read answered "Attachment could not be saved" over a file that
    was saved in full - a false negative on a completed write. Measuring is reporting; it must
    not be able to fail the thing it reports on. The size now falls back to 0 when unreadable,
    which is a smaller lie than the one it replaces and is visible beside a path that exists.

54. **Two tests I changed rather than added.** Four `DraftUpdateReentrancyTests` cases used
    `Assert.Throws<InvalidOperationException>`, which is an EXACT type match in xunit, so the new
    subclass broke them. They now assert the subclass, which is strictly stronger - the same
    cases additionally pin that the outcome field travels.

55. **Mutation check: 37 decision lines reverted, 30 caught, 7 not - and EIGHT of the thirty
    were caught only after the gap they exposed was closed.** Each was reverted, built, run
    against the whole non-live suite and restored. The table, with what noticed each one, is in
    `tmp-aitrace/mutation-table.md`; the seven that remain are in `TODO.md` with what each would
    need. The eight closed gaps are the useful half of the exercise and are worth naming, because
    every one of them was a test that LOOKED like it covered the area:
    - the per-item timeout outcome was pinned for `move_mail` and not for `archive_mail`, which
      keeps its own copy of the same arm on its own COM call;
    - the send path's two messages (rows 9 and 10) had no test at all until one drove the real
      two-step token flow against a stand-in;
    - the framing refusal was only ever exercised over a READ (`GetProfileName`), so the branch
      the whole row is about - what it says over a MUTATION - was unreachable; it now runs over
      `TrySaveAttachment`, which is classified mutating and returns a string, so an oversized
      answer is one line of setup rather than a fabricated 64 MB draft result;
    - the tool layer's timeout and too-large advice were inline in `catch` arms, so no test could
      reach them; both are now pure internal helpers beside `ComFailureAdvice`, which is the
      shape that made row 15 assertable in the first place;
    - the derived-draft failure had no test distinguishing "the source was never opened" from
      "a draft may already be sitting in Drafts", which is the difference between the two
      answers.

56. **Method note, and it is the same hazard as sections 25 and 35 wearing a third disguise.**
    The mutation script restored each revert with `text.replace(new, old, 1)`, which is correct
    only while the REPLACEMENT is unique - and a mutation that DELETES a block has a replacement
    made of ordinary code. M12's was `if (moved == null)`, which occurs in both `ArchiveOne` and
    `MoveOne`, so the restore inserted MoveOne's block into ArchiveOne against names that do not
    exist there. **M13 to M36 then all failed to BUILD and proved nothing**, and the verifier
    PASSED on that tree, because M12's original text really was present exactly once - in the
    wrong place. The splice is now by INDEX in both directions, which cannot be ambiguous.
    **The lesson that generalises: a two-directional text check is not enough after a mutation
    pass; only a build is.** Separately, one background run was killed mid-mutation and left M30
    applied - the documented kill hazard, caught by the verifier and restored by hand - so the
    remaining mutations were run in small chunks to bound what a kill could cost.

57. **Row 18 was verdict TRUE and got the optional half anyway.** The signature backup's
    "the operation was ABORTED and nothing was modified" is true of the USER'S signatures, and it
    is true for the reason that makes any such claim assertable: the backup runs before anything
    touches them. What it never mentioned is that a half-written BACKUP directory can survive -
    `CreateDirectory` succeeds and a later `File.Copy` fails. It now names that directory when
    one exists and says nothing when it does not, which is the distinction the two tests pin.

58. **The census cost driver was established by COUNTING CALLS, not by measuring, and that is
    said out loud because nothing here may run against the mailbox.** Per store per pass the
    old census cost ~15 setup calls, 12 for the six `GetDefaultFolder` volatile probes, ~8 per
    folder for the tree walk, and **five cross-process calls per item walked** - `Items[i]`
    (which OPENS the message) plus `EntryID`, `ReceivedTime`, `Size` and `Subject`. At the
    3,000-item budget the last term is **15,000 round trips, ~94% of the census**, and the
    next largest term is smaller by two orders of magnitude. It needs 12 ms per call, or ~60
    ms per item opened, to exceed the 3-minute budget - the range a non-cached Exchange
    delegate store sits in, and `info@voipfabric.com` is a delegate store. This is arithmetic
    over the code plus a per-call cost model; it is not a measurement, and the first run after
    the fix is what will turn it into one.

59. **`Table.GetArray`, not `Table.GetNextRow`, and the difference is the whole point.**
    `GetNextRow` + `Row.GetValues` is the idiom shipped in four places in Core and is
    therefore the better-proven one, but it is two round trips per ROW. `GetArray(n)` is one
    per n rows: at a batch of 200 the per-store budget goes from 15,000 calls to about
    fifteen. The cost is that a 2-D variant array carries no labels, so nothing in it
    distinguishes it from its own transpose. That is handled by checking BOTH dimensions
    against numbers already known - the rows requested and the column count the table
    reported - which leaves a transpose acceptable only when a folder happens to have exactly
    as many rows left as the table has columns, and the EntryID and duplicate checks then
    reject that. Every shape failure abandons the folder to a COUNT, so the worst case of
    being wrong about `GetArray` is a weaker reading, never a false one.

60. **The 3-minute STA timeout was NOT raised, and "make it per store" was already true.**
    `CaptureMailFolderCensus` is one `RunSta` per store, so the budget has always been per
    store; the direction that suggested changing that buys nothing. Raising it was rejected on
    its own terms: it moves a silent failure later without evidence, and the term that could
    plausibly exceed it has been removed. What was added instead is the evidence - each
    store's census now prints its folder count, what identifying it cost and its elapsed ms,
    and a refusal names how far the census had got. `CensusIdentityPlan` doubles as that
    progress record, which is the only reading still available when the STA thread has not
    come back; the counters are plain int writes read from another thread, so a post-timeout
    reading may be stale but cannot be torn. It is offered to the maintainer as an open
    question rather than decided (`TODO.md`, 2026-08-20 entry, question 2).

61. **The identity budget was deliberately left at 500/3,000.** With the walk 250x cheaper the
    obvious move is to raise it, and raising it changes WHAT THE GUARD PROVES - a 4,918-item
    Sent Items would go from counted to identified. That is a decision about the guard, not a
    side effect of making it affordable, so the numbers did not move and the question is
    written up instead.

62. **One thing the census no longer does, and it is a deliberate trade.** A folder whose
    table will not carry all four columns is now recorded as a COUNT, where the old per-item
    walk would have produced identity with a null fingerprint. Identity without a fingerprint
    cannot prove a filing, so it turns a person filing mail during a run into a suite failure;
    identity without a subject cannot tell the suite's own mail from anyone else's, so a
    removal the suite caused would be attributed "undecidable". Both are worse than a count.
    The alternative I rejected was keeping the item-by-item walk as a per-folder fallback:
    that would mix two fingerprint derivations inside one census, and a departure read one way
    could then not be matched to its arrival read the other - a false "ITEMS REMOVED" line
    manufactured by the fallback itself. How many folders fell back is now printed, so the
    degenerate case (no store carries the columns at all) cannot happen quietly.

63. **The census reads an unspecified `DateTimeKind` from a table as UTC, and Core reads the
    same variant as LOCAL. Exactly one of them is wrong, and it is not the census that pays.**
    `CensusTableRow.ReadUtc` takes `Unspecified` as already-UTC (what Microsoft documents for
    the `Table` object, and the contract `DaslDateLiteral` states for this solution).
    `OutlookComSession.ReadRowDate` does the opposite: `Kind != Utc ? ToUniversalTime()`, which
    on a COM-marshalled `VT_DATE` (always `Unspecified`) subtracts the local offset. If the
    census's reading is right, then in the census only the instant PRINTED beside a departed
    item can ever be off, because both ends of every comparison come through the one method -
    but `ReadRowDate` feeds `_lastAdmittedUtc`, which becomes the resumed exhaustive scan's
    inclusive `at or before` date bound, so a bound two hours early would SKIP the mail
    received in those two hours and the scan would report itself complete. **Not fixed here:
    it is shipped Core behaviour, outside this task, and which way it is wrong is a
    measurement rather than a reading.** It is cheap to settle: `T2/LiveTableSortProbeTests`
    already reports `FirstRowReceivedUtc` through `ReadRowDate`, so comparing it against the
    same item's `MailItem.ReceivedTime` on one live read decides it for both call sites.

64. **The projection, the column map and the bulk-read shape check were written as PURE
    functions on purpose (`T2/CensusTableRow`).** Everything in `LiveOutlookTestMailer`'s
    census is unreachable from CI - no non-live test can open a folder - so the decisions that
    matter were moved to the one side a test can reach: what makes a row identify an item,
    which column spellings are accepted, and what shape of block may be believed. That is why
    16 of the mutations below could be caught at all; the ones inside the COM walk could not
    be, and are listed as such rather than left implied.

65. **Mutation check: 26 decision lines reverted, 15 caught, 11 not - and the split is exactly
    the seam between the pure half and the COM half.** Every one of the 15 lines a non-live test
    can reach was caught; every one of the 11 that were not caught is inside
    `CaptureMailFolderCensus`, which no non-live test can execute a line of. That is a
    structural gap, already recorded in `TODO.md` for the census generally, and it is why the
    projection, the column map and the bulk-read shape check were extracted as pure functions in
    the first place - before that extraction the number reachable from CI would have been about
    four. Full table in `tmp-aitrace/mutation-table.md`; the 11 are listed in `TODO.md` with what
    each would need. **Two of the 26 are worth naming.** M09 (`a fingerprint needs both halves`)
    was caught by the COMPILER rather than a test, which is the cheapest possible catch and worth
    knowing about. M01 (`the bulk read rejects anything that is not a 2-D block`) came back NOT
    CAUGHT on the first pass and exposed a real hole in my own test: the rank case used a 1-D
    array of four entries while asking for two rows, so the ROW-COUNT check rejected it and the
    rank check was never exercised. A rank-1 array of exactly the requested length was added and
    M01 was re-run and caught. That is the second time in this repository a mutation has found a
    test that looked like it covered the area and did not.

66. **A pre-existing non-live test now fails on this machine, and it is not this change.**
    `T3/OutlookAvailabilityCiTests.SearchAlwaysAnswers_AndSaysWhetherItIsComplete` drives the
    SHIPPED MCP server executable over stdio against the real Windows Search index and the real
    Outlook, and asserts that a `search` answers within 100 s. It measured **139.1 s**. It passed
    at the start of this session (2028 of 2028) and failed 4 of 4 afterwards, including against a
    mutation of a constant in the TEST assembly, which that server process cannot see. Outlook has
    been up for 40 hours with 10,791 s of CPU on it. So the final count is **2045 passed, 1
    failed, 2046 total**, and the one failure is the machine, not the change - which is stated
    rather than rounded off, because a suite reported as green when it is not is how a real
    failure gets waved through later. Worth a look separately: a test whose pass depends on the
    developer's Outlook being un-busy will keep doing this.

## 4. VM state (`OutlookAI-TestVM`)

- Guest credentials for PowerShell Direct are **not recorded here**. This repository is
  public, so they live only in the gitignored live-test settings on the maintainer's
  machine. An earlier revision of this file printed them in plain text; that value is
  therefore compromised and must be treated as burned wherever it still appears.
- **PowerShell Direct lands in session 0, where Outlook can never finish starting.** Anything
  needing COM runs as a scheduled task with `-LogonType Interactive`, which lands in session 1.
- Checkpoints: `CP-01-WIN-CLEAN`, `CP-02-INSTALLER-STAGED`, `CP-03-OUTLOOKAI-INSTALLED`,
  `CP-04-OFFICE-GOLD`, `CP-05-ADDIN-TRUSTED`, `CP-06-PRE-CORPUS`.
- Store: one PST, `Outlook Data File`, **not in the local index** - which is the property the whole
  testbed depends on.
- **Placement and dates are both VERIFIED** on this store: items must be created in Drafts, saved,
  have the unsent flag cleared, saved again, then moved (`DraftsThenMoveWithSentFlag`); received
  dates written through the PropertyAccessor **do** drive DASL selection. The earlier "dates are
  impossible" verdict was an artefact of the placement bug.
- Host-side scripts live in `C:\Users\jori\Downloads\tmp-outlookai-vm\`; the corpus manifest is
  copied out to `corpus-vm1.jsonl` there, and **without it the corpus cannot be torn down**.

## 5. Measurements taken

- **Real profile, 5 stores** (`jori@huisman.io` alone: Archive 108,144 items, Sent 10,024): Inbox-only
  exhaustive scan 30.5 s; with subfolders 66.5 s; whole store 7-day window 36.6 s over 32 folders;
  whole store 60-day window **timed out at 105 s having scanned 3 folders of 32**.
- **Frame high-water on the real profile: 432 KB against a 64 MB limit**, zero refusals. The ceiling
  was set by the 30-second sweep timeout, not by any size limit.
- **The Windows Search provider does not put undated rows at the top of an ordered result** in
  either direction, and accepts the `1601-01-01` floor literal - so the displacement guard's
  recovery statement should essentially never fire. Not taken under a mapi scope, so the `T2` probe
  still stands.
- **Sweep cost, MEASURED 2026-08-19 on the VM corpus** (one PST outside the local index, 20,000
  items across the four arrival-path folders with real received dates: Inbox 10,912, Sent 4,964,
  Deleted 2,467, Junk 1,663; 1,612 inside the seven-day fallback window, so the 200-per-folder cap
  engages in at least two folders). Four sweeps, same corpus: **13,624 / 11,818 / 10,652 / 11,889
  ms**, `itemsSeen` 758 in all four. Codes: `no_index_frontier` + `item_cap_unsorted` on every one,
  plus `body_cap` with `itemsBodyCapped=2` on the term-matching-nothing run. **~12 s for ONE store
  with the cap engaged; five stores extrapolates to ~60 s against the old 30 s budget** - the direct
  explanation for the sweep timeout on the real profile. This is a fast LOCAL PST; Exchange is
  slower per item.
- **Frame high-water from that same corpus: 10,734,599 bytes** - 10.2 MB over 758 items, ~13.5 KB
  per item, from a SINGLE store. Five unindexed stores extrapolates to ~54 MB against the 64 MB
  frame limit, so `SweepBodyBytesBudget` (32 MiB) bites first, which is its design intent. This
  supersedes the reading of the 432 KB measured on the real profile: that was bounded by the 30 s
  sweep timeout, not by the item caps, so raising the sweep budget is what makes the body budget
  load-bearing.
- **`item_cap_unsorted` fired on every sweep of that corpus**, and the answer correctly reported the
  cap as having cut arbitrarily rather than claiming the oldest mail is what is missing. The
  2026-08-18 H2 fix, observed working against real data rather than in a test. **DOWNGRADED
  2026-08-19:** this entry originally read "i.e. `Table.Sort` genuinely does not apply on that
  store", which is a stronger claim than the evidence supports. What was observed is that the sort
  CALL was refused; the shipped `catch` wrapped `Columns.Add` and `Sort` together, so it could not
  even say which of the two failed. Microsoft documents that a sort property may be referenced "by
  their explicit string names only; cannot reference properties by their namespaces" and the call
  passes a namespace - a store-INDEPENDENT explanation that fits "every sweep" better than a
  per-store one does. The `catch` is now split and `sweep.sortRefusedFolders` counts the folders
  where the column was added and the sort still threw; `T2/LiveTableSortProbeTests` settles it.
- **Corpus build throughput:** 50.9 items/s without the move rung. With the move rung each item is
  written twice, so budget roughly double.
- **Tripwire census call budget, COUNTED FROM THE CODE rather than timed** (2026-08-20, and the
  distinction matters - nothing in this session was allowed to touch the mailbox). Per store, per
  census pass, BEFORE the change: ~15 setup calls, 12 for the six `GetDefaultFolder` volatile
  probes, ~8 per folder for the tree walk, and 5 cross-process calls per item walked, so
  **~15,000 round trips at the 3,000-item budget - about 94% of the whole census**. AFTER: the
  item term becomes ~1 call per 200 rows plus ~25 fixed per identified folder, so **about 15 calls
  for the same 3,000 items**, and the folder tree walk (~8 calls x up to 165 folders on a delegate
  store) becomes the largest remaining term at ~1,300 calls. The threshold that broke it: 12 ms
  per call, or ~60 ms per item opened, exceeds the 3-minute STA budget - which is ordinary for a
  shared Exchange mailbox that is not cached locally.

## 5b. Two research outputs worth reading before the next work starts

Both live in the session trace folder under Downloads (`tmp-aitrace/`), because they were produced
while the repo was held by another agent.

**`resumable-scan-design.md`** - the design for the maintainer's chosen continuation-token scan.
Recommends a three-tier resumption ladder behind a short opaque handle, with walk state in the
server parent rather than the COM child: per folder, try a date cursor, then a validated ordinal
skip, then a full folder restart with EntryID dedup, reporting which tier paid. Key constraints it
establishes: the folder walk has no stable order until this code imposes one (the `list_folders`
walk already sorts siblings for exactly that reason and its comparator should be shared); within a
folder MAPI documents that nothing is stable, so a token cannot name a row; resuming only at folder
boundaries is fatal here because the measured failure IS one 108,144-item folder; and server-side
state is forced rather than preferred, because proving no folder was skipped needs the set of
completed folders, which is too big for the wire. Four open questions are listed in its section 8.

**`sort-hypothesis.md`** - UNRESOLVED, and potentially the largest single defect found this session.
Microsoft documents that `Table.Sort` accepts property names "by their explicit string names only;
cannot reference properties by their namespaces". The sweep passes
`urn:schemas:httpmail:datereceived`, which is a namespace reference. If the documentation holds for
this call then the freshness sweep has NEVER sorted on any store for any user, its 200-item cap has
always cut arbitrarily, and the tier whose entire purpose is recent mail has been taking an
arbitrary 200 instead of the newest 200. It would also mean this session's reading of
`item_cap_unsorted` - "the sort genuinely does not apply on that store" - is wrong: it would not
apply anywhere. **Four read-only PowerShell probes failed to settle it**, all with the same error
across every property form and the no-argument form, which is PowerShell late binding against the
`Table` COM object rather than Outlook's verdict. **It needs a few lines of C# in the existing live
harness**, where binding is not a problem. One cheap fix falls out regardless: the shipped `catch`
wraps the column add and the sort together, so `sortApplied: false` cannot say which failed.

## 6. Still unverified, and why

- The sweep budget itself. The corpus exists to settle it and the measurement has not been taken.
- Whether the sweep recognises the substitute Junk folder the generator creates in a PST. If it does
  not, sweep measurements cover ~92% of the corpus rather than all of it.
- Everything in `TODO.md` marked as needing a live profile.
- **The count tripwire has still never executed**, and now neither has the run-plan machinery that
  decides when it verifies. The pure logic on both sides is pinned in CI; the COM census and the
  `CollectionFinished` -> `Verify` path are not, and cannot be from a non-live test.
- **Nobody has run the `LiveTier=Portable` subset anywhere.** The classification is enforced by
  `T1/LiveTierInventoryTests` and the filter expression is verified by discovery
  (`--list-tests --filter "Category=Live&LiveTier=Portable"` selects 19 of 115), but no test in that
  subset has been executed on a machine other than the maintainer's, so "Portable" means "reads as
  runnable there" and not yet "ran there".
- **Bringing the three stdio shape classes into a collection changes when they run.** Their implicit
  collection names sorted after every "Live" name, so `SuiteCollectionOrderer` now ranks the new
  `LiveMcpToolShape` collection late on purpose, to keep them where they were. That reasoning is
  from reading xunit's ordering, not from a live run.
- ~~**The census cost on the maintainer's own profile.** Up to 3,000 items walked per non-hub
  store, four late-bound property reads each, at least twice per run, over five stores. Never
  measured. On a local PST it should be milliseconds; Exchange is a different question.~~
  **ANSWERED 2026-08-20, by the census failing:** the baseline for `info@voipfabric.com` exceeded
  the 3-minute STA budget and refused the whole live tier. Exchange was indeed a different
  question. The walk is a bulk table read now (section 3.58-3.64); what remains unverified is
  everything below.
- **What the census costs AFTER the table change has still never been timed** - the arithmetic
  says the item term drops from ~15,000 round trips per store to ~15 and that the folder tree
  walk (~8 calls per folder, up to 165 folders on a delegate store) becomes the dominant term,
  but no live run has happened. Every store now prints its own elapsed ms, so the next run is
  the measurement.
- **Whether `Columns.Add` accepts `Size` and `ReceivedTime` on a real Exchange folder table.**
  Both have a fallback spelling and a folder whose columns will not land degrades to a count, so
  being wrong is safe - but if it happens on every folder, the identity half of the guard is
  effectively off. That is exactly what the new `N folder(s) fell back to counting` line in the
  census log is for; read it on the first run.
- **Whether `Table.GetArray` hands back rows-by-columns on this provider.** The repository's own
  `TrySkipRows` already reads it that way, both dimensions are checked against known numbers
  before a block is believed, and a mismatch counts the folder rather than inventing items. Still
  a reading of documentation and of one untested code path, not an observation.
- **Whether an Outlook `Table` reports date-time values in UTC or in local time** - see section
  3.63. The census is safe either way; `OutlookComSession.ReadRowDate` is not, because its value
  becomes a resumed scan's date bound.
