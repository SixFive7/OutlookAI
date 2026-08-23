# Measuring the sweep and scan budgets against a known corpus

> **SUPERSEDED IN PART, 2026-08-19.** Both budgets have since been measured and changed, so
> everything below that reads as "what to decide" is now a record of how the question was
> framed rather than an open one. What actually happened: **the sweep was measured** on this
> corpus at ~12 s for ONE store with the 200-per-folder cap engaged (13.6 / 11.8 / 10.7 /
> 11.9 s over four runs, 758 items each), which extrapolates to ~60 s on the maintainer's
> five-store profile and is the direct explanation for the 30 s timeout seen there.
> `SweepBudgetMs` is now **180 s** (3x that) with a derived inner budget
> `SweepWorkBudgetMs` of 165 s that stops the walk gracefully at a folder boundary; the scan
> got its **own deadline class** (`ComHostOperationClass.ExhaustiveScan`, 615 s) rather than
> dragging the shared deadline up, so `ExhaustiveTimeBudgetMs` is **600 s** while every other
> tool keeps a 300 s hang detector. The same run also measured a frame high-water of
> 10,734,599 bytes from one store, which makes `SweepBodyBytesBudget` load-bearing rather than
> insurance. Current values and their derivations: `Docs/magic-numbers.md` and section 2 of
> `Docs/autonomous-session-log.md`. Two symbols named below no longer exist under those names:
> `ComOperationBudgets.ChildWorkBudgetMs` is now `ExhaustiveScanWorkBudgetMs`, and it derives
> from `ExhaustiveScanDeadlineMs` rather than from the shared operation deadline.

**What this file is for.** Two budgets in the MCP server are set by argument rather than by
measurement, and one of them is about to be changed. This is the plan for replacing both
arguments with numbers, using the synthetic corpus that
`OutlookAI.RemediationTools corpus-build` puts into a local PST on the Hyper-V test VM. It
says what to run, in what order, and what each number would settle.

**What the first real run established (2026-08-19).** A 40 000 item corpus was built into the
VM's PST in 12m27s at 50.9 items/sec with zero failures; resumability and determinism both
worked (a second run skipped the 2 000 items an earlier timing run had already made). Three
things went wrong, and all three are now guarded rather than remembered:

1. **Items were queued for delivery.** 5 532 landed in the target store's Outbox. Inert on
   that VM because the profile has no mail account, and 5 532 queued messages on any profile
   that has one. The build now refuses unless the profile has **no accounts at all** - see
   "Before you run anything" below.

   **The population is now identified, 2026-08-24.** The plan for that exact shape - corpus
   `vm1`, seed 4242, 40 000 items - marks **5,532 items unread**, which `corpus-plan` now
   prints as its own line. Not approximately: 5,532. So the queued items are precisely the
   ones the plan wanted left unread, and the read state was the only thing the builder did
   differently to them. It set `MailItem.UnRead` and then wrote `PR_MESSAGE_FLAGS` WHOLESALE,
   as `MSGFLAG_READ` for a read item and as **0** for an unread one. Both are gone: the read
   state now travels through a single read-modify-write that clears `MSGFLAG_SUBMIT` - the bit
   that means "queued for delivery" - on every item, clears `MSGFLAG_UNSENT` only when the
   placement rung calls for it, and preserves every bit it does not own.

   **What is still unproved is the mechanism inside Outlook**, and only a build can prove it.
   So `corpus-census` reports Outbox strays SPLIT BY THE PLAN'S INTENDED READ STATE. A small
   build settles it: all-unread confirms the identity, an even split kills it.
2. **Every item was filed as a draft**, because `Items.Add` + `Save` produces an UNSENT item
   and Outlook files unsent items in Drafts whatever folder they were added to. The sweep
   covers Inbox, Sent Items, Deleted Items and Junk Email and **not** Drafts, so a sweep over
   40 000 items selected **6**, in 234-367 ms. The measurement could not be taken.
3. **The date guard's refusal message was wrong**, and the wrong sentence is what caused the
   run to proceed. It said the items would be dated "roughly now", from which an all-recent
   corpus looks like the sweep's worst case - a reasonable inference from a false premise.
   The truth is that an item the folder table carries no delivery time for is selected by
   **no** window, so the sweep sees fewer items, not more. The message now states the
   consequence as a count.

**A corpus expires, and it does so silently (2026-08-24).** Everything above is measured
against windows counted back from the corpus ANCHOR. Every test asks its question against the
CLOCK. The two diverge from the moment the corpus is written, and roughly six weeks later a
seven-day window selects nothing at all - while every test asking about that window still
PASSES, because selecting nothing is a valid answer about an empty window. The measurement
stops happening and nothing goes red.

Two commands close that, and the live tier runs the first of them fail-closed at fixture time:

* `corpus-verify` is PURE - no Outlook, no store, runnable on the host - and refuses when any
  window under test has emptied. It derives the shift the store already carries from the
  manifest, and prints each window as `now/at-anchor`.
* `corpus-reanchor --to now --execute` is the repair. It shifts every item's received and
  submit instants to an ABSOLUTE target, so it is idempotent and resumable; it never creates,
  moves or removes an item; and it leaves the seed, the shape and the manifest header
  untouched, so the corpus is still the corpus every earlier measurement was taken against.

Regenerating instead was considered and rejected: the numbers above are held against THIS
snapshot, and a regenerated corpus is a different population wearing the same figures.

**Run `corpus-verify` before quoting any number on this page.** A measurement taken against a
stale corpus is a measurement of an empty window.

**Why a corpus is needed at all.** The two questions are about volume, and the developer
profile cannot ask them. Every store on it is indexed, so the freshness sweep's window comes
from each store's index frontier and the seven-day fallback never engages; and the only
unindexed store to hand is the test VM's, which is empty. A corpus is the missing half: a
store that is local, unindexed, and holds a known number of items of known sizes at known
ages.

**The state of the two numbers, and where they come from.**

| Constant | Value now | Where | How it was arrived at |
| --- | --- | --- | --- |
| `MailService.SweepBudgetMs` | ~~`30_000`~~ **`180_000` since 2026-08-19** | `McpServer/OutlookAI.Core/Services/MailService.cs` | Was a model (four folders x 200 items x five stores at ~15 ms per item opened). Now MEASURED on this corpus: ~12 s per store with the cap engaged, ~60 s extrapolated to five stores, and the budget is 3x that. |
| `MailService.ExhaustiveTimeBudgetMs` | ~~`105_000`~~ **`600_000` since 2026-08-19** (derived from `ComOperationBudgets.ExhaustiveScanWorkBudgetMs`, itself `ExhaustiveScanDeadlineMs` less the return trip) | `McpServer/OutlookAI.Core/Services/MailService.cs`, `McpServer/OutlookAI.Core/Com/ComOperationBudgets.cs` | Derived from the COM operation deadline minus the result-return headroom. Never measured against volume: on the real profile a sixty-day scan reached three folders of thirty-two in 105 s and stopped there. |

**The one model that already exists**, from `Docs/magic-numbers.md`: roughly **19 ms per
folder** fixed plus **15 ms per item opened**, fitted over 215 sweeps on the real profile. It
predicts a seven-day empty-index sweep at 6.0 s capped and 22.9 s uncapped for 1501 items.
Every number below should be compared against that model, because if the model holds on a
PST the budget question becomes arithmetic, and if it does not, the model is the thing that
was wrong.

---

## Read this before proposing 180 s

> **RESOLVED 2026-08-19: 180 s fits now.** The composition below was the real blocker and it was cleared exactly as this section says it must be - in one pass, moving `OperationDeadlineMs` with the sweep. `SearchBudgetMs` is now 240 s (60 s index + 180 s sweep) inside a 300 s operation deadline, and the exhaustive scan was taken OUT of that composition entirely by giving it its own deadline class. The paragraph below is kept because the constraint it describes is permanent, even though this particular instance of it is settled.

**180 s did not fit.** `MailService.SearchBudgetMs` is
`(SearchIndexTimeoutSeconds * 1000) + SweepBudgetMs`, and T1
`BudgetCompositionTests.SearchBudget_IsComposedFromItsPartsAndFitsTheOperationDeadline`
asserts that sum is at most `ComOperationBudgets.OperationDeadlineMs`, which is `120_000`.
With `SweepBudgetMs = 180_000` the sum is 195 s against a 120 s deadline, so that test fails
before anything reaches a mailbox. Raising the sweep budget past about **105 s** is therefore
not a one-constant change: it moves `OperationDeadlineMs`, and with it the child work budget
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

## Before you run anything - the guards that gate a build

**The profile must have no mail accounts.** `corpus-build` reads `Session.Accounts`, compares
each account's `DeliveryStore` to the target by `StoreID`, and refuses if there is any account
at all. That is stricter than "no account delivers into the target", and the strictness is
forced: `SendUsingAccount` is per item, so any account may send a message that lives anywhere,
and "no account can send from this store" is only provable through the object model as "no
account can send". There is **no override flag** for this one.

**Placement is probed before the build and the build refuses if it fails.** `corpus-probe`
now walks a placement ladder - create in place with MSGFLAG_UNSENT cleared; create in Drafts,
clear the flag, then `Move`; create in Drafts and `Move` without the flag; and a plain saved
item as a control - and a rung passes only when the item's `Parent` **is** the target folder
**and** the target folder's `GetTable` returns it. The second half is decisive: the sweep
enumerates a folder through its table, so an item the table does not carry does not exist as
far as this measurement is concerned. `--allow-drafts-placement` overrides it and says in the
same breath that the sweep will select 0 of N items.

**The probe's table check now names the item it is looking for.** It used to ask the folder
for every corpus subject in it and walk at most 2,000 rows. Against a folder that already held
~22,000 corpus items it never reached the item it had just created, gave up, and reported it
ABSENT from the folder's table - so the build refused a placement that works. The lookup is
now filtered on the probe's own reserved ordinal, so it selects roughly one row; the row cap is
a bound on a runaway rather than a search budget; and reaching it is reported as INCONCLUSIVE
with its own refusal text, which blames the measurement instead of the store. The date probe's
exclusion half had the same defect and is fixed the same way.

**Every build censuses itself.** `corpus-census` re-reads the store and compares it against the
plan: right count, right folders, one copy each, and nothing stranded in Drafts or the Outbox.
It sets the build's exit code alongside the failure count, and it can be run on its own at any
time. It exists because the first real build reported 40 000 items created and zero failures
while every one of them sat in Drafts.

**Run `corpus-probe` on its own first.** It is cheap, it creates and deletes a handful of
throwaway items, and it answers both questions - placement and dates - before any long build.

## Step 0 - build the expectation sheet before touching Outlook

`corpus-plan` is pure and runs anywhere:

```
OutlookAI.RemediationTools corpus-plan --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000
```

It prints, for the exact corpus that will be built: item count per folder, per size class and
per age band; the **unread** count; total and mean body bytes; how many bodies exceed 24 KB,
96 KB and `OutlookComSession.SweepBodyCharsCap`; and how many items fall inside 1, 7, 30, 60,
90 and 365 days of the anchor. **Save this output beside the results.** Every measurement below is a
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

**Confirm before continuing, in this order:**

1. The store and profile lines both say accepted, and `profile accounts: 0`.
2. The **placement** probe named a verified rung. If it says `NOT ACHIEVABLE`, steps 3 and 4
   are void and step 2 becomes the primary route - see "If placement fails" at the end.
3. The **date** probe named a verified rung. If it says `NOT ACHIEVABLE`, every step that
   mentions a window is void - see "If the dates do not stick".

Placement is settled before dates on purpose. Probing a date against an item that was filed
somewhere other than the folder being queried cannot tell "the date does not drive selection"
from "the item is not in this folder", and the first run's date verdict was taken under
exactly that confusion, so it proves nothing either way.

**Then let the machine settle.** Windows Search must be given the chance to either index the
PST or not, and which one happened must be established rather than assumed: run
`outlook_health` and read `index.perStore[]` for the corpus store. If it is indexed, the
seven-day fallback will not engage and the sweep measurements will be measuring the wrong
path. Excluding the PST from indexing is the intended state of this VM.

## Step 2 - per-folder sweep cost, out of band

**This step is the fallback route for the whole document.** If placement cannot be made to
work, steps 3 and 4 are impossible and everything the sweep budget needs has to come from
here: per-item and per-folder cost measured directly, multiplied out by the folder and store
counts of a real profile. It is a weaker answer than measuring the shipped sweep - it measures
the cost model's coefficients rather than the thing itself - but it is a measured answer, and
it is the one the current 30 s value never had.

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

## If placement fails

If no placement rung verifies, items can only be created as drafts, and that changes the shape
of this document rather than merely annotating it:

- **Steps 3 and 4 are void.** The sweep does not read Drafts. It will report `foldersSwept: 4`
  and `freshness: "live"` and select nothing, which looks identical to a quiet mailbox.
- **Step 2 becomes the primary measurement**, not a supporting one. Time the table walk and
  the per-row `GetItemFromID` directly against the Drafts folder, which holds the whole corpus
  and is a perfectly good folder for measuring per-item cost. The coefficients are what the
  budget arithmetic needs; the folder they were measured in does not matter.
- **Step 5 still works** if the scan is scoped to the folder the corpus is in. The exhaustive
  scan walks named folders rather than the sweep's fixed four.
- **Say so in the results.** A sweep number taken against a drafts-placed corpus is a number
  about an empty folder set, and it will read like a fast sweep.

If it comes to that, record what each rung reported - the probe prints `landedIn` for every
one - because "which rung failed and where the item went" is the whole diagnosis, and it is
the evidence anyone revisiting this needs.

## If the dates do not stick

`corpus-build` probes date fidelity before it writes anything, and refuses to build a corpus
with unusable dates unless `--allow-undated` is passed. If it comes to that, the resulting
corpus is still worth having, and it is worth being precise about what changes:

- **Still measurable:** per-item and per-folder sweep cost at a known item count; the body cap
  and the byte budget; frame size; exhaustive scan throughput scoped to a folder; everything
  in steps 2, 4 and 5 that does not name a window.
- **Not measurable:** anything involving a date window at all. This is worse than it sounds
  and the first run got it wrong: undated items are **not** "all recent", they are *absent*
  from the sweep. The restriction is `(datereceived >= X) OR (date >= X)`, so an item exposing
  neither property matches no window however wide - a 7-day sweep and a 60-day sweep would
  both select zero. Step 3 is void, and so is the capped-versus-uncapped comparison.

Do not paper over this case, and do not reason around it the way the first run did. The
guard's message now states the consequence as a count for exactly that reason.

**This one is also a product finding, not only a tooling one.** Mail that reaches a PST
without transport - imported, copied between stores, restored from backup - can carry no
usable delivery time, and it is then invisible to the freshness sweep on a real user's
machine while the payload still reports the folder as swept. Recorded as **H3** in
`Docs/completeness-gaps.md`, with the evidence and, importantly, with the confound: the
observation that motivated it was taken while every item was also in the wrong folder, so it
does not yet prove the date half. A re-run with placement verified is what would settle it.
