# Measuring the sweep and scan budgets against a known corpus

**What this file is for.** Two budgets in the MCP server are set by argument rather than by
measurement, and one of them is about to be changed. This is the plan for replacing both
arguments with numbers, using the synthetic corpus that
`OutlookAI.RemediationTools corpus-build` puts into a local PST on the Hyper-V test VM. It
says what to run, in what order, and what each number would settle.

**Why a corpus is needed at all.** The two questions are about volume, and the developer
profile cannot ask them. Every store on it is indexed, so the freshness sweep's window comes
from each store's index frontier and the seven-day fallback never engages; and the only
unindexed store to hand is the test VM's, which is empty. A corpus is the missing half: a
store that is local, unindexed, and holds a known number of items of known sizes at known
ages.

**The state of the two numbers, and where they come from.**

| Constant | Value now | Where | How it was arrived at |
| --- | --- | --- | --- |
| `MailService.SweepBudgetMs` | `30_000` | `McpServer/OutlookAI.Core/Services/MailService.cs` | Proposed to become `180_000`. The only support so far is a model: four folders x 200 items x five stores at roughly 15 ms per item opened. |
| `MailService.ExhaustiveTimeBudgetMs` | `105_000` (derived from `ComOperationBudgets.ChildWorkBudgetMs`) | `McpServer/OutlookAI.Core/Services/MailService.cs`, `McpServer/OutlookAI.Core/Com/ComOperationBudgets.cs` | Derived from the COM operation deadline minus the result-return headroom. Never measured against volume: on the real profile a sixty-day scan reached three folders of thirty-two in 105 s and stopped there. |

**The one model that already exists**, from `Docs/magic-numbers.md`: roughly **19 ms per
folder** fixed plus **15 ms per item opened**, fitted over 215 sweeps on the real profile. It
predicts a seven-day empty-index sweep at 6.0 s capped and 22.9 s uncapped for 1501 items.
Every number below should be compared against that model, because if the model holds on a
PST the budget question becomes arithmetic, and if it does not, the model is the thing that
was wrong.

---

## Read this before proposing 180 s

**180 s does not fit.** `MailService.SearchBudgetMs` is
`(SearchIndexTimeoutSeconds * 1000) + SweepBudgetMs`, and T1
`BudgetCompositionTests.SearchBudget_IsComposedFromItsPartsAndFitsTheOperationDeadline`
asserts that sum is at most `ComOperationBudgets.OperationDeadlineMs`, which is `120_000`.
With `SweepBudgetMs = 180_000` the sum is 195 s against a 120 s deadline, so that test fails
before anything reaches a mailbox. Raising the sweep budget past about **105 s** is therefore
not a one-constant change: it moves `OperationDeadlineMs`, and with it `ChildWorkBudgetMs`,
`ExhaustiveTimeBudgetMs` and the supervisor's own timing. **Decide the shape of that change
before the measurement, so the measurement can be aimed at the right question.**

**The sweep's timeout is not a graceful stop, and the scan's is.** When the sweep budget is
exceeded the gateway raises a `TimeoutException`, the supervisor kills and replaces the COM
host, and the search degrades to index-only. The exhaustive scan instead latches a
`TimedOut` flag and returns partial results, which is what `ResultReturnHeadroomMs` buys. So
"how long may the sweep take" is really "how long may a user wait before the COM host is
destroyed", and that is a different question from "how long may a scan run". Any number
proposed for the sweep should be justified against the harsher consequence.

---

## Step 0 - build the expectation sheet before touching Outlook

`corpus-plan` is pure and runs anywhere:

```
OutlookAI.RemediationTools corpus-plan --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000
```

It prints, for the exact corpus that will be built: item count per folder, per size class and
per age band; total and mean body bytes; how many bodies exceed 24 KB, 96 KB and
`OutlookComSession.SweepBodyCharsCap`; and how many items fall inside 1, 7, 30, 60, 90 and 365
days of the anchor. **Save this output beside the results.** Every measurement below is a
ratio against one of these numbers, and computing them after the fact from the store is both
slower and less trustworthy than reading them off the plan that produced it.

Two of those numbers deserve attention before the build starts. The **total body bytes** is
the lower bound on how much the PST will hold, and at the default mixture the mean is roughly
10 KB, so 40 000 items is on the order of 400 MB of body text before Outlook's own overhead.
And the count inside **7 days** versus inside **60 days** is the whole point: if those two are
close, change the date bands before building, not after.

## Step 1 - build, and record what the build itself cost

```
OutlookAI.RemediationTools corpus-build --store "<PST display name>" --allow-store "<PST display name>" \
    --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000 \
    --manifest D:\corpus\vm1.jsonl --execute
```

The build reports items per second and total body bytes. That is not one of the two numbers
being chased, but write it down anyway: it is the only figure on record for how long a rebuild
after a VM rollback will take, and the plan is deterministic, so it will be the same next time.

**Confirm before continuing:** the date probe printed at the start says a rung verified. If it
says `NOT ACHIEVABLE`, every step below that mentions a window is void - see "If the dates do
not stick" at the end.

**Then let the machine settle.** Windows Search must be given the chance to either index the
PST or not, and which one happened must be established rather than assumed: run
`outlook_health` and read `index.perStore[]` for the corpus store. If it is indexed, the
seven-day fallback will not engage and the sweep measurements will be measuring the wrong
path. Excluding the PST from indexing is the intended state of this VM.

## Step 2 - per-folder sweep cost, out of band

The server reports one clock for the whole sweep (`sweep.elapsedMs`) and no per-folder or
per-store timings at all. The per-folder shape has to come from the existing read-only probe,
`Docs/v3-probes/soakfix13-probe-sweep-cost.ps1`, which mimics the real sweep closely: the same
`GetTable` date filter, the same `Columns.Add` and `Sort`, the same row walk, the same
200-item cap.

**That probe is gitignored** (`.gitignore` covers `Docs/v3-probes/`), so it exists on the
developer machine only and has to be copied to the VM by hand. If it has been lost, it is
reconstructible from the description above plus one detail that is not optional: the date
literal must be the year-first `yyyy-MM-dd HH:mm:ss` form that `DaslDateLiteral` emits.
Outlook parses a DASL date literal in the MACHINE locale, and a `MM/dd/yyyy` literal on a
Dutch-locale box silently selected the wrong rows in both directions - it is what made a
seven-day sweep select nothing while still reporting four folders swept.

Run it against the corpus store with the window set to 7 days rather than the 10/60/1440
minutes it currently uses, and with `$openItems` both false and true. That pair is the
measurement that matters most:

- **`openItems = false`** is the table walk alone: how long it takes to find the rows.
- **`openItems = true`** is the table walk plus one `GetItemFromID` per row, which is what the
  real sweep does before reading nine properties off each item.

The difference between them, divided by the row count, is the **per-item cost** the 15 ms
model claims. Getting that number on a PST, at a row count the plan predicted, is the single
most useful measurement in this document, because the sweep budget is per-item cost times
items times folders times stores and nothing else.

## Step 3 - the seven-day sweep, through the shipped server

Now measure the thing itself. A `search` against the corpus store with no index behind it
takes the fallback window, so the sweep selects the plan's `7d=` count.

Record from the payload: `sweep.elapsedMs`, `sweep.performed`, `sweep.itemsSeen`,
`sweep.folderCapReached`, `sweep.itemCappedFolders`, `sweep.itemsBodyCapped`,
`sweep.bodyBudgetExhausted`, `sweep.coverageGaps`, and from `outlook_health`,
`comHost.largestFrameBytes`.

Three traps to avoid, all of which will otherwise produce a wrong number that looks right:

1. **`SweepCache` has a 10-second TTL.** Two identical searches inside 10 s give the second one
   `sweep.cached: true` and `elapsedMs: 0`. Vary the query or wait.
2. **A cold COM host gets a connect floor of 90 s on top of the budget**
   (`allowConnectFloor: true`). The first sweep after a host start is not comparable to the
   rest. Discard it, or record it separately and label it.
3. **The 200-items-per-folder cap may or may not bite.** With the default corpus the seven-day
   window holds roughly 8 % of the corpus, spread over four folders, so at 40 000 items each
   folder is well past 200 and `folderCapReached` should be true. That is deliberate: it is
   the first measurement of what a capped sweep costs when the cap is actually reached, which
   the developer profile has never done (0 of 215 sweeps hit it).

**Then repeat with the cap raised.** The cap is what keeps the sweep bounded, so the honest
budget question is "how long does an uncapped seven-day sweep take on a store of this size",
and that is the number the 180 s proposal has to beat. Raising `SweepPerFolderCap` for one
build is the cheapest way to get it; the alternative is to compute it from the per-item cost
in step 2 and the plan's `7d=` count, which is the model this whole exercise is meant to
replace.

## Step 4 - the same sweep with the body cap in play

The corpus deliberately contains bodies past `OutlookComSession.SweepBodyCharsCap` (500 000
chars) - `corpus-plan` prints how many. Any sweep whose window reaches them must cut them and
say so.

Compare `sweep.itemsBodyCapped` against the plan's `bodies over sweep cap` figure for the same
window, and watch `sweep.bodyBudgetExhausted` and `comHost.largestFrameBytes`. The measured
high-water frame on the real profile is 432 KB against a 64 MB limit, about 152x headroom; the
question this corpus answers is what that number becomes when the sweep is dominated by
quoted threads rather than ordinary mail. `SweepBodyBytesBudget` is 32 MiB, so a sweep that
reaches it will report `bodyBudgetExhausted` and the frame will stop growing - confirming the
bound works is as valuable as the timing.

## Step 5 - exhaustive scan throughput

The exhaustive scan has no default window; it refuses to run unbounded and requires `folder`
and/or `after`. So the sixty-day figure from the real profile is a test parameter, not a
constant, and it is reproduced here by passing `after` sixty days before the anchor.

Run the scan store-scoped with `exhaustive: true`, and record `exhaustive.elapsedMs`,
`foldersScanned`, `foldersSkipped`, `timedOut`, `truncated`, `rowsDropped`, `depthLimitReached`.

The number wanted is **items examined per second**, and it needs care: the scan stops at
`maxItems` (`top`, capped at 100), so on a corpus this dense it will hit that cap almost
immediately and the elapsed time will say nothing about throughput. To measure throughput,
either use a search term that matches nothing - so the scan runs to the end of its budget and
`foldersScanned` and the folder item counts give the denominator - or scope it to one folder
at a time and use the plan's per-folder counts. **A term that matches nothing is the better
run**: it is the worst case, it is what the 105 s budget actually has to survive, and it maps
directly onto "three folders of thirty-two in 105 s" from the real profile.

Report the result as folders per second and items per second, then multiply by the real
profile's 32 folders to say whether 105 s was ever going to be enough there.

## What each number settles

| Measurement | Settles |
| --- | --- |
| Per-item cost from step 2 (`openItems` true minus false, over rows) | Whether the 15 ms/item model holds on a PST. Every budget arithmetic downstream depends on it. |
| Per-folder fixed cost from step 2 (intercept over folders with no matching rows) | Whether the 19 ms/folder model holds. This is what multiplies by store count on a real profile. |
| `sweep.elapsedMs` capped, step 3 | What a seven-day sweep costs when the 200-item cap is reached. Directly comparable to 30 s. |
| `sweep.elapsedMs` uncapped, step 3 | The worst case the budget must survive on one unindexed store, and the number 180 s is really being proposed against. |
| Capped vs uncapped ratio | Whether `SweepPerFolderCap` is doing the work, or whether the budget is. |
| `itemsBodyCapped` vs the plan, step 4 | That the body cap fires when it should, and what it costs. |
| `largestFrameBytes` under a big-body sweep, step 4 | How much of the 152x frame headroom is real once bodies dominate. |
| `bodyBudgetExhausted`, step 4 | That the 32 MiB accumulation bound is reachable and works. |
| Folders/second and items/second, step 5 | Whether 105 s can scan a real profile's 32 folders, and what window size it can afford. |

## Deciding the sweep budget from the numbers

The budget has to cover the worst realistic sweep, and the worst realistic sweep is
`per-folder cost x folders x stores` plus `per-item cost x items opened`. The corpus gives
both coefficients; the profile gives the multipliers (four folders per store, however many
stores). **Compute the number, then add margin, then check it against the operation deadline
before proposing it** - the deadline is the constraint that 180 s failed, and any candidate
above roughly 105 s needs the deadline moved as part of the same change.

If the computed number is comfortably under 105 s, the change is one constant and one test
update. If it is not, the honest conclusion is that the per-folder cap and not the budget is
the thing to change, because a budget that must exceed the operation deadline is saying the
work does not fit in one operation.

## Repeating it

The corpus is deterministic: same `--corpus-id`, `--seed`, `--anchor` and shape produce the
same items, so a VM rollback plus a rebuild reproduces the exact population a number was
measured against. That is the point of forbidding a defaulted anchor - a corpus anchored on
the clock would drift between runs and quietly invalidate comparisons.

Snapshot the VM **after** the build and **before** any measurement, so every run starts from
the same store. `corpus-teardown --manifest ... --execute` removes the corpus if the PST is
to be reused for something else; it deletes only what the manifest records, by EntryID
allowlist and subject tag together.

## If the dates do not stick

`corpus-build` probes date fidelity before it writes anything, and refuses to build a corpus
with unusable dates unless `--allow-undated` is passed. If it comes to that, the resulting
corpus is still worth having, and it is worth being precise about what changes:

- **Still measurable:** per-item and per-folder sweep cost at a known item count; the body cap
  and the byte budget; frame size; exhaustive scan throughput scoped to a folder; everything
  in steps 2, 4 and 5 that does not name a window.
- **Not measurable:** anything comparing a 7-day window against a 60-day one, because every
  item would carry a received time of roughly "now" and both windows would select the whole
  corpus. Step 3 as written is void, and the uncapped-versus-capped comparison would be a
  comparison against the entire store rather than against a window.

Do not paper over this case. A corpus whose items are all dated "now" looks exactly like a
good one from the outside, and a window measurement taken against it would be wrong in a way
nothing downstream could detect.
