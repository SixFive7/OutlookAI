# What the test VM can and cannot prove

**Date:** 2026-08-23. **Tree:** `1f6a491`, clean. **Method:** read-only. Nothing was built, no
test was run, no mailbox, Outlook, Hyper-V guest or registry hive was touched. Every count below
comes from parsing the test assembly's attributes and from reading the code and the repo's own
measurement records.

**How to read the evidence markers.** `[V]` = verified this session by reading the code or by
parsing it mechanically. `[R]` = recorded in this repository as a past measurement (I did not
re-take it). `[I]` = inference, reasoned from the code but not observed. Anything unmarked in a
list inherits the marker of its heading.

---

## 0. The shape of the suite, measured

`[V]` Parsed from `McpServer/OutlookAI.McpServer.Tests` by reading every `[Fact]`/`[Theory]` with
its class-level and method-level traits (1,607 test **methods**; xunit expands theories to 2,081
non-live **cases**, which matches the session log's own count).

| Tier | Files | Methods | `Category=Live` | Non-live |
|---|---|---|---|---|
| T1 (in-process unit) | 90 | 1,391 | 0 | 1,391 |
| T2 (live COM) | 53 | 107 | 107 | 0 |
| T3 (stdio against the built server exe) | 24 | 109 | 9 | 100 |

`Category=Live`: **116** methods (not 115). **20** `LiveTier=Portable`, **96** `ProfileBound`.
`[V]` The 115/19 in `Docs/live-tier-on-the-vm.md` is one behind: `03a0857` (today) added
`LiveTableSortProbeTests.ATableDate_IsEitherUtcOrLocal_AndTheRunSaysWhich` as a Portable test.

**Supporting data beside this file:** `live-inventory.txt` is the full 116-method live table, one
line per test, with its `LiveTier` and its effective `Requires` set. `inv.json` is the machine-readable
inventory of all 1,607 methods (tier, file, class, method, line, class traits, method traits,
collection, Fact/Theory), so any count here can be re-derived without re-parsing the assembly.

`Requires` totals across live methods `[V]`: `SearchIndex` 47, `MailAccount` 42, `Transport` 41,
`MultipleStores` 41, `DelegateStore` 23, `InteractiveDesktop` 16, `ProbePopulation` 7,
`SmallHubStore` 1, `AddInRegistry` 1.

**`Requires` is declared per CLASS in 30 of 36 live classes** `[V]`, so it is the union of what any
one test in the class needs. That over-attributes, and it is why the "96 cannot move" figure is an
upper bound on the impossible set rather than a measurement of it. Six classes push traits down to
the method (`LiveExhaustiveSearchTests`, `LiveFolderScopeTests`, `LiveManageSignatureTests`,
`LiveShowMeTests`, `LiveSignatureTests`, `LiveSweepScopeTests`, plus `Phase7LiveMcpToolShapeTests`);
those are the only classes where the trait means what it says per test.

---

## 1. The tier-3 mislabelling, quantified

The claim under examination: sixteen T3 files spawn the real server and reach the maintainer's
production Outlook while sitting outside `Category=Live`. **It is true, and it is narrower and
worse than it looks.**

`[V]` The sixteen non-live T3 files are `ComHostSupervisionCiTests`, `DescriptionBudgetCiTests`,
`DraftOptionsCiToolShapeTests`, `HtmlBodyCiToolShapeTests`, `McpStdioConformanceTests`,
`MoveArchiveCiToolShapeTests`, `OutlookAvailabilityCiTests`, `Phase2CiToolShapeTests`,
`Phase3CiToolShapeTests`, `Phase4CiToolShapeTests`, `Phase5CiToolShapeTests`,
`Phase7CiToolShapeTests`, `SearchSchemaCiTests`, `SoakBatchCCiToolShapeTests`,
`SoakToolSurfaceCiTests`, `WritingRulesGateCiTests`. 100 test methods, ~154 cases.

`[V]` I read every `CallToolAsync` site in all sixteen. **Eight methods across five files actually
reach the machine's real Outlook, real Windows Search index or real user data.** The other 92 are
`tools/list` schema reads or argument-validation refusals aimed at deliberately unresolvable inputs
(`id = "h999999"`, `account = " "`, `folder = "  "`, `name = "OutlookAI-NoSuchSignature-424242"`),
and each of those returns `InvalidArgument` before any COM host is started. That is verified from
the call sites and the assertions, not assumed from the file names.

The eight:

| Test | What it does to the real machine |
|---|---|
| `OutlookAvailabilityCiTests.ATransientOutlookState_...` | one real `list_accounts` through COM |
| `OutlookAvailabilityCiTests.RepeatedCalls_NeverEachPayAFullBudget` | five more real `list_accounts` |
| `OutlookAvailabilityCiTests.SearchAlwaysAnswers_AndSaysWhetherItIsComplete` | real `search` for `"invoice"` against the real index **plus a freshness sweep of every mounted store** |
| `OutlookAvailabilityCiTests.HealthAlwaysAnswersQuickly_...` | real `outlook_health` |
| `Phase2CiToolShapeTests.OutlookHealth_...` | real `outlook_health` |
| `Phase7CiToolShapeTests.OutlookHealth_OnAnyMachine_ReturnsWellFormedReport` | real `outlook_health` |
| `McpStdioConformanceTests` | real `outlook_health` |
| `SoakToolSurfaceCiTests.ListSignatures_IsCallableOnAnyMachine_...` | reads the user's real signature directory |

All eight are read-only. Nothing in the non-live tier writes to a mailbox `[V]`.

**Three things make this a real finding rather than a naming complaint.**

1. **`--filter "Category!=Live"` is the CI command in `.github/workflows/mcpserver.yml`** `[V]`. The
   filter is exactly as safe as the machine it runs on. On GitHub's `windows-latest` there is no
   Outlook, so the same eight tests take the degraded branch and the filter looks correct. On the
   maintainer's machine it is a production read. **The filter never expressed a safety property; it
   expressed the absence of Outlook on the runner.**
2. **None of the sixteen sits in a `[Collection]`** `[V]`. `T1/LiveTierInventoryTests` enforces that
   every `Category=Live` class belongs to a guarded collection, precisely because three T3 classes
   once fell through that hole. The sixteen are outside the check by construction: they are not
   `Category=Live`, so the inventory test never looks at them. **They run with no health preflight,
   no store-count tripwire, no census and no signature snapshot** - the four guards that exist
   because a live run once left seven tagged items in a real mailbox. They are read-only today, so
   nothing is lost today; the point is that a future edit to any of the sixteen has no guard in
   front of it and no test that would notice.
3. **One of them already fails because of the real mailbox.** `[R]`
   `OutlookAvailabilityCiTests.SearchAlwaysAnswers_...` measured **139.1 s** against an Outlook that
   had been up 40 hours, passed at session start (2028/2028) and failed 4 of 4 afterwards, including
   against a mutation of a test-assembly constant the server process cannot see. Its wall-clock
   assertion has since been removed for exactly that reason, and the file's own header now says so.

**Correctly named, the tier is:** 92 wire-shape and validation tests that are genuinely
machine-independent (they need only the built exe), and 8 environment probes that belong in the live
tier under a `LiveTier` trait. `ComHostSupervisionCiTests` is worth calling out on the other side: it
injects faults with `OUTLOOKAI_COMHOST_FAULT` and forces `OUTLOOKAI_COMHOST_LIVENESS=Responsive` with
a 4 s deadline override, so the whole timeout/kill/respawn/breaker path runs with no Outlook at all
`[V]`. That one is honestly labelled.

---

## 2. Question 1: what cannot run on the VM at all

Re-derived against the **decided** VM, not the one the traits were written for: two profiles (one
with no accounts for corpus generation, one with a dummy account on an unroutable server), three
stores (Corpus A indexed, Corpus B unindexed, plain bystander), ~20,000 items per corpus store with
real received dates spread over years, Windows 11 + Office + the add-in + checkpoints.

### 2.1 T1: nothing is lost

`[V]` All 1,391 T1 methods are in-process and machine-independent. I checked the four categories
that could have been otherwise:

- No T1 test starts Outlook or opens a mailbox. The Outlook interop references in fourteen T1 files
  are type references satisfied by fakes and by pure snapshot records.
- `AuditLogTests` writes to a temp directory and says so explicitly, never `%LOCALAPPDATA%`.
- `OfficeVersionDetectionTests` does read the real registry (`OutlookProfileRegistry.OfficeVersion`),
  and is written to hold both with and without Office present. On a VM with Office installed it
  proves what it proves on the dev box.
- `LiveOutlookPreflightTests`, despite the name, injects its liveness probe and its clock.

**Conclusion: the VM runs the entire T1 tier at full strength. That is 87% of the methods and about
93% of the non-live cases, and it is the largest single fact in this analysis.** `[R]` The session's
own interim filter measured 1,927 cases in 9 seconds with 0.05 s of Outlook CPU.

### 2.2 T3 non-live: nothing is lost, and something is gained

`[V]` All 100 need only the built server exe. The eight environment probes become *more* meaningful
on the VM, not less, because they currently assert against a mailbox nobody controls. The one caveat
is that `Phase7LiveMcpToolShapeTests.Health_OverStdio_...` carries `Requires=AddInRegistry`, so the
add-in must have run at least once on the VM to populate its tuning state.

### 2.3 The genuinely impossible live set

These cannot be made to work on the VM by configuration or by seeding. Each is impossible for a
reason, and the reason is not the same in each group.

**(a) Delegate and shared-mailbox semantics - 6 methods, and the decision to exclude them is right.**

`[V]` `LiveFolderScopeTests.DelegateFirstLevelFolders_StillResolve_...`,
`LiveFolderScopeTests.DelegateSubfolders_AreReachableAgain_...`,
`LiveIndexSearchTests.DelegateStoreSubtree_ReturnsRowsUnder2s`,
`LiveStaleIndexRowTests.DelegateHitsInANestedFolder_AreReadable_ViaTheFlatLeafName`,
`LiveMailServiceTests.ListAccounts_ExactAccountsDelegatesAndFlags`,
`T3/LiveMcpToolShapeTests.Status_Accounts_Folders_GoldenShapes_OverRealStdio`.

The property they need is not "a second mailbox". It is that **Windows Search publishes a delegate
mailbox's folders FLAT while Outlook nests them**, so the index scope is `<storePrefix>/1` and a
folder that Outlook calls `Team/2026/Invoices` appears in the index under its leaf name alone. That
is why `HitLocator` has a delegate leaf-walk with its own retry ladder (`DelegateLeafWalkRetryMs`
400 x 3) and why `IndexSearchService` probes `SCOPE='<prefix>/1'` for existence at all. `[I]` A local
PST cannot produce that shape: there is no `/1` store-type segment and no flattening, so a faked
delegate store would make the leaf-walk code path unreachable while the tests around it went green.
The 2026-08-20 decision not to fake it is the correct one, and this analysis endorses it without
qualification: **faking it would give false confidence in the single area this product has been
surprised by most often.**

**(b) Real transport - 6 methods that genuinely need mail to arrive.**

`[V]` `LiveOutlookTestMailer.SendSelfMail` finds the profile account whose `SmtpAddress` equals the
hub name, sets `SendUsingAccount` through the PROPERTYPUTREF accessor, hard-verifies the identity,
sends, and then `LiveInboxArrival.WaitFor` sweeps the hub Inbox for up to **180 s** until the item
lands. On an unroutable account the item queues in the Outbox and never arrives, so every such test
burns 180 s and then throws `TimeoutException`.

Files calling it `[V]`: `LiveDraftOptionsTests`, `LiveDraftTests`, `LiveFreshModeTests`,
`LiveHtmlDraftTests`, `LiveMoveArchiveTests`, `LiveSignatureTests`, `LiveSweepScopeTests`,
`LiveUpdateDiscardTests`. The ones where the arrival is the point rather than a seeding convenience:

- `LiveFreshModeTests.FreshSearch_FindsSelfSentMail_BeforeIndexCatchesUp_ThenCleansUp` - the whole
  test is "mail arrives, and search finds it before the index does". No arrival, no test.
- `LiveMoveArchiveTests.MoveChain_TestFolderRoundTrip_...` - seeds arrived mail to move.
- `LiveSweepScopeTests.ControlledCorpus_CrossColumnTermsMatch_...` - seeds a controlled population.
- `LiveDraftOptionsTests.DerivedDrafts_...` and `LiveHtmlDraftTests.ReplyDraft_...` and
  `LiveUpdateDiscardTests.UpdateDraft_OnAReply_...` - a reply needs an arrived original to reply to.
- `T3/Phase5LiveMcpToolShapeTests.SendTool_TwoStepFlow_RoundTrip_...` and
  `T3/MoveArchiveLiveMcpToolTests` and `T3/Phase4LiveMcpToolShapeTests` carry their own
  `ArrivalSeconds`/`SeedVisibleSeconds` waits `[R]`.

**Answering the maintainer's open sub-task directly: does any live test assert on Sent Items after a
successful send?** `[V]` Not in the sense they feared, and yes in a sense that matters more.
`LiveSendTests` (4 methods) is explicitly "all WITHOUT any transport - every path here refuses BEFORE
`Send()`, and the audit log is asserted to contain NO send line", so nothing asserts on a delivered
copy. But the artifact sweeps do cover Sent Items and the Outbox
(`LiveOutlookTestMailer.SweepFolderIds` is `{16, 6, 5, 4, 20, 19, 21, 22, 3}`, and folder id 4 is the
Outbox), and the mailbox-safety contract requires every run to end with zero tagged artifacts in the
**Outbox**. **A queued send on
an unroutable account leaves a permanent tagged artifact in the Outbox that the zero-artifact sweep
will find and fail on.** So the unroutable account is not sufficient on its own: either the sweep has
to be taught that an Outbox item on a Portable machine is expected and deletable, or a local SMTP
sink is needed. This is a decision, and it is question 3 in section 7.

**(c) Exchange EntryID and store semantics - 3 methods that are about Exchange itself.**

`[V]` `LiveDecodeVerifyTests` (3 methods). Its own header records the Phase-1 live finding: on cached
Exchange stores, `GetItemFromID` rejects the 24-byte OST-internal id decoded from the index URL with
`0x80040107`, in every store, so hit mapping must go through `HitLocator`'s folder probe and
`ItemPathDisplay` fallback. `[I]` A PST does not have that split: the id in the index URL and the id
`GetItemFromID` accepts are not related the same way. Run on the VM this test would either pass
through the primary path trivially or fail on an assertion about a byte layout that only Exchange
produces. Either outcome is noise. **This test is not "delegate-bound" or "index-bound"; it is
Exchange-bound, and no trait currently says so.**

**(d) Hardcoded profile arity - 2 methods.**

`[V]` `LiveMailServiceTests.ListAccounts_ExactAccountsDelegatesAndFlags` asserts
`Assert.Equal(3, outcome.Accounts.Count)` and that every account's SMTP address is one of the
configured store names. `T3/LiveMcpToolShapeTests.Status_Accounts_Folders_...` is its stdio twin.
A one-account VM fails both. These are relaxable (read the count from settings) but relaxing them
removes exactly the assertion that catches a misconfigured profile, which is the trade the TODO
already names.

**(e) The hand-curated real-mail probes - 7 methods.**

`[V]` `LiveSearchInTests` (6) and `LiveStaleIndexRowTests` (1) carry `Requires=ProbePopulation`. They
need `subjectOnlyProbe`: a real population whose term appears in the SUBJECT and the sender address
but **not** in the body stream, which is the exact shape the pre-D40 unqualified `CONTAINS('term')`
could never match. `[I]` This one is *not* impossible on the VM - it is a seeding gap, and section 5
says how to close it. It is listed here because it is impossible against the corpus **as it exists
today**, where every item's body is generated from its subject's ordinal and no item has a sender at
all.

**(f) One paging assertion that contradicts the corpus hub - 1 method.**

`[V]` `Phase7LiveMcpToolShapeTests.Search_TopOne_OnHubStore_SetsTruncated_AndTopHundredDoesNot`
carries `Requires=SmallHubStore` and asserts the hub holds between 2 and 99 items. The Portable scans
and sweeps all need the corpus to BE the hub (they take a "corpus too small" early return otherwise;
see 3.6). The two requirements are mutually exclusive on one machine. With three stores there is now
a way out that did not exist before: make the **bystander** the small store this test points at, and
give the test a settings-driven store name instead of the hub. That is a code change, not a
configuration.

**(g) Tests that need a real sender - 1 method today, more once the index tier moves.**

`[V]` `LiveIndexSearchTests.SenderFilter_PerColumnContains_IndexBackedUnder2s` collects
`hit.FromAddress` from recent hits and asserts `candidates.Count > 0`. `[V]` Corpus items are created
with `items.Add(0)` and have only `Subject`, `Body`, `UnRead`, `PR_MESSAGE_FLAGS`,
`PR_MESSAGE_DELIVERY_TIME` and `PR_CLIENT_SUBMIT_TIME` set. **No sender, no recipients.** So this
test fails on a corpus store.

### 2.4 The size of the impossible set

`[V]` Counting method by method rather than by class trait:

| Category | Methods | Fixable by seeding? | Fixable by code? |
|---|---|---|---|
| Delegate/shared semantics | 6 | no | no |
| Real transport arrival | 6 | no (needs an SMTP sink) | no |
| Exchange EntryID semantics | 3 | no | no |
| Hardcoded 3-account arity | 2 | no | yes, at a cost |
| `ProbePopulation` shapes | 7 | **yes** | no |
| `SmallHubStore` conflict | 1 | no | yes |
| Needs a real sender | 1 | **yes** | no |
| **Hard floor (first three rows)** | **15** | | |

**15 of 116 live methods are genuinely impossible on any VM.** Another ~11 are impossible today but
reachable through seeding or a bounded code change. **The remaining ~90 are a configuration and
seeding problem, not a capability problem** - which is a very different answer from "96 cannot move",
and it is different because the 96 came from class-level traits that were written before the dummy
account and the three-store layout were decided.

### 2.5 Six blockers that are not tests, and will stop the VM before any test runs

These are the things I expect to break first. All `[V]` from the code unless marked.

1. **`testHubStoreDisplayName` is used as an SMTP address.** `LiveUpdateDiscardTests` calls
   `NewDraft(Hub, Hub, ...)`; `SendSelfMail(smtpAddress: Hub, ...)` matches it against
   `Account.SmtpAddress`; `OutlookComSession.FindAccountBySmtp` resolves the account by that string
   and throws `ArgumentException` on blank. **The hub PST must be named exactly the dummy account's
   SMTP address** (for example a store literally called `test@vm.invalid`). `[I]` Whether Outlook
   accepts an `@` in a store display name is untested. This is cheap to try and it gates the entire
   draft family.
2. **`expectedStoreDisplayNames` is overloaded and the three-store layout splits its two meanings.**
   It is simultaneously the tripwire's watch list (must include all three stores) and the index
   tier's "these must be discoverable in the index" list. `LiveIndexSearchTests.ProbeParity_AllThreeStores`,
   `ProbeParity_McpShapedQuery` and `StoreDiscovery_FindsAllExpectedStores` call
   `LivePhase1Fixture.GetScope(name)`, which **throws** when a store is not among the discovered index
   scopes. Corpus B is unindexed by design and the bystander is near-empty, so all three throw. The
   settings file needs a fourth list (indexed stores) or those tests need to read a subset.
3. **The corpus generator and the live tier need different profiles, and switching costs an Outlook
   restart.** `CorpusSafety.EvaluateProfile` refuses unless `Session.Accounts.Count == 0`, with no
   override, deliberately. So: boot the no-account profile, build or tear down the corpus, checkpoint,
   switch profiles, run tests. **Every corpus regeneration is a profile switch.**
4. **`CorpusSafety.EvaluateStore` requires `Store.IsDataFileStore == true`.** `[I]` The repo's own
   comment says that property is "true only for a store not tied to an account". If the dummy account
   delivers into the hub PST, `corpus-build`, `corpus-teardown` and `corpus-reindex` may refuse that
   store from then on, even under the no-account profile if the binding persists in the profile. Point
   the dummy account's delivery store at a **fourth, throwaway PST**, not at the corpus. Worth a
   five-minute check before building anything.
5. **`machineProfile: "Portable"` converts "found nothing to test" from a failure into a pass.**
   `LiveTestSettings.RequireProductionPopulation` throws on Production and **no-ops on Portable**. Two
   call sites today, but the pattern is wider: `LiveResumableScanTests` early-returns green four times
   with `"hub corpus fits one page - no rung was exercised"`, and
   `LiveIndexSearchTests.FilterShapes_...` asserts `Assert.All(withAttachments.Hits, ...)` over what
   will be an empty list. **Running the tier on the VM without a "how many assertions actually fired"
   check trades a mailbox risk for a silent-green risk.**
6. **A fixed-anchor corpus decays.** `CorpusPlanOptions.AnchorUtc` is deliberately required rather
   than defaulted, so the same seed always means the same corpus. The consequence is that every
   "last N days" assertion measures the age of the checkpoint.
   `LiveIndexSearchTests.ProbeParity_DateRangeQuery_HitsUnder2s` asserts rows exist within the last
   30 days; `LiveSweepScopeTests.DefaultSweep_...` sweeps the last 15 minutes; the sweep's 7-day
   fallback window selects the `last-24h` and `1d-7d` bands. **Six weeks after the anchor, the corpus
   silently stops exercising the windows it was built for**, and every one of those tests still
   passes. The corpus needs either a re-anchor-and-rebuild step in the runbook or a small
   continuously-refreshed "recent" slice.

---

## 3. Question 2: what runs but proves less

### 3.1 The latency gap, stated precisely

The claim to quantify: a local data file served ~1,200 items/second where Exchange served ~12.

**What the repo actually contains** `[R]`:

| Measurement | Where | Number |
|---|---|---|
| Census per-item COM cost model | log 3.58 | **5 cross-process calls per item** (`Items[i]` which OPENS the message, then `EntryID`, `ReceivedTime`, `Size`, `Subject`) |
| Cost that breaks the 3-minute STA budget on a delegate Exchange store | log 3.58 | **12 ms per call, ~60 ms per item** = **~17 items/s** |
| Older sweep model | `magic-numbers.md` | ~15 ms per item opened, ~19 ms per folder (5 calls x 15 ms = 75 ms/item = **~13 items/s**) |
| Sweep of one 20,000-item local unindexed PST, bodies included | log 5 | 758 items in 10.7-13.6 s = **~64 items/s** |
| Corpus build into a local PST | log 5 | **50.9 items/s** written (two COM writes per item with the move rung) |
| Tripwire baseline illustration | `live-tier-on-the-vm.md` | 312 identified items in 431 ms = **~724 items/s**, i.e. ~0.28 ms/call |

`[I]` So the exact pair "1,200 vs 12" is not written down in this repository as a measured pair. What
**is** supported is: the same five-call-per-item COM path costs **~0.3 ms per call on a local PST**
and **12-15 ms per call on a non-cached delegate Exchange store**, a ratio of **40x to 50x per call**
and, at the observed per-item rates, **~700-1,200 items/s local against ~13-17 items/s Exchange, a
ratio of roughly 50x to 90x.** That is one and a half to two orders of magnitude, and it is the right
order for the maintainer's argument. I would state it in the record as **"roughly 50x to 90x, from
the repo's own per-call figures"** rather than as 1,200:12, because the smaller number is derived
from a call-cost model and the larger from a different operation, and a precise-looking pair invites
someone to divide them later.

**The direction that matters is not in the ratio, though; it is in what the VM can be wrong about.**
A VM can only ever say a budget is **generous**. It can never say a budget is **too small**, because
the operation it measures is 50x-90x faster than the one the budget exists for. Every constant below
therefore has a one-sided validation on the VM.

### 3.2 Which constants that actually undermines

`[V]` values read from `ComOperationBudgets.cs`, `MailService.cs`, `ComHostPolicy.cs`. `[R]`
derivations from `Docs/magic-numbers.md` and the session log.

**Group 1 - derived on the VM already, so the VM cannot re-validate them; only the real profile can.**

| Constant | Value | Why the VM cannot validate it |
|---|---|---|
| `SweepBudgetMs` | 180,000 | Derived AS ~12 s x 5 stores x 3. The 12 s came from the VM corpus. `magic-numbers.md` says in as many words that "the margin is headroom rather than luxury **because the corpus is a fast local PST and Exchange is slower per item**". The VM will report ~12 s and conclude 180 s is 15x generous, forever. |
| `SweepWorkBudgetMs` | 165,000 | `SweepBudgetMs - 15,000`. Inherits the above. |
| `ThreadWalkBudgetMs` | 180,000 | `= SweepBudgetMs`. Inherits the above. |
| `SweepPerFolderCap` | 200 | `[R]` "never fires in steady state, always fires on an unindexed store". The VM exercises only the always-fires half; the real profile is the only place the never-fires half is observed. |
| `SweepBodyBytesBudget` | 32 MiB | The 10,734,599-byte high-water that made this load-bearing was measured on the VM corpus, whose body-size mixture is a design parameter. On a real profile the mixture is whatever the user's correspondents send. The budget was validated against a distribution we chose. |

**Group 2 - derived on the real profile, so the VM cannot notice if they become wrong.**

| Constant | Value | Real-profile evidence behind it |
|---|---|---|
| `OperationDeadlineMs` | 300,000 | `[R]` "4.5x the slowest **healthy** operation measured": a whole-store 7-day sweep at 36.6 s and an Inbox-with-subfolders exhaustive scan at 66.5 s, both on the five-store profile. **A hang detector is calibrated by the slowest healthy operation, and only a real profile produces one.** On the VM the slowest healthy operation is ~12 s, so 300 s reads as 25x and a regression that doubled real-world cost would be invisible. |
| `ExhaustiveTimeBudgetMs` / `ExhaustiveScanDeadlineMs` | 600,000 / 615,000 | `[R]` Exists because a 60-day whole-store scan **reached 3 folders of 32 in 105 s** on the real profile. The VM's three stores of four folders each finish far inside it. `[V]` Step 5 of `corpus-measurement-plan.md` (items/s for the scan) has **never been run on either machine**, so this budget has no throughput measurement behind it at all. |
| `ConnectDeadlineMs` | 180,000 | `[R]` Sized for "a large OST on a slow disk". The VM has no OST. |
| `SearchIndexTimeoutSeconds` | 60 | `[R]` Measured healthy at 60-550 ms against a real, busy SystemIndex. A 20k-item single-store index answers faster still. |
| `SearchBudgetMs` | 240,000 | Composed from the two above. |
| `MoveBatchBudgetMs` / `MinimumItemBudgetMs` | 240,000 / 1,000 | `[I]` A PST move is a local file operation; an Exchange move is a server round trip. A 50-item batch is where the difference compounds. `[R]` The maintainer has already asked for "a live batch exercise before release" - this row is why that is the right instinct. |
| `HealthProbeDeadlineMs`, `HealthIndexTimeoutSeconds`, `HealthPerStoreIndexBudgetMs` | 5,000 / 4 / 8,000 | The per-store index probes run once per store. Three local stores against five real ones, two of them delegates on a saturated indexer, is not the same measurement. |
| `StoreIndexProbeBudgetMs` | 1,500 | `[R]` Measured on the dev machine: 9-10 ms for a delegate-subtree miss, 27-30 ms for an `@` discovery miss. Both probes are delegate-shaped. |
| `StaleIndexNoticeMinutes` | 30 | `[R]` Set at the **p90 of the index frontier age on the dev profile, from 177 probes** (median ~6 min). This is a statistic about a live Exchange mailbox being indexed while mail arrives. A VM corpus is indexed once and then the frontier never lags again. **The whole staleness ladder - this, `VeryStaleAdviceMinutes` 720, `SweepSafetyMargin` 10 min, `EmptyIndexSweepWindow` 7 days - describes an index racing arriving mail, and a VM has no arriving mail.** |
| `PumpedStaRunner.RetryAfterMs` / `GiveUpAfterMs` | 250 / 30,000 | `[R]` Justified by a Phase-3 live run. `[I]` These handle COM `SERVERCALL_RETRYLATER` from `IMessageFilter`, which Outlook raises when it is busy with a server or a user. A PST-only Outlook on an idle VM with no user essentially never raises it, so **the retry path is unreachable code on the VM**. |
| `RecipientResnapshotDelayMs` 1,500, `ExplorerFolderSettleDelayMs` 250 x6 | | `[I]` Recipient resolution on Exchange is a directory lookup. With a dummy unroutable account there is no directory, so recipients resolve as SMTP strings immediately and the resnapshot path never has anything to wait for. |
| `LiveInboxArrival.DeadlineSeconds` 180 | | `[R]` Exists because a real round trip once exceeded 120 s and failed a 17-minute live run. Not applicable on a VM at all. |

**Group 3 - genuinely machine-independent, and the VM proves them fully.**

`[V]` The whole breaker and supervision surface: `UnresponsiveTimeoutThreshold` 2,
`UnresponsiveCooldownMilliseconds` 30,000, `StartFailureBackoffThreshold` 3,
`StartBackoffMilliseconds` 30,000, `AutostartCooldownMilliseconds` 20,000,
`MinimumDispatchDeadlineMilliseconds` 1,000, `CleanExitGraceMilliseconds` 250,
`HandshakeBudgetMs` 30,000 / floor 10,000. `ComHostSupervisionCiTests` drives all of them with
injected faults and a 4 s deadline override, so no real wedge is needed. `T1/BudgetCompositionTests`
pins the arithmetic (`SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs <= OperationDeadlineMs`, now
240 s inside 300 s) with no mailbox involved.

**The honest summary of group 3 is worth stating, because it changes the shape of the gate:** the
*logic* of the breaker, the kill, the respawn, the deadline composition and the graceful expiry is
fully provable on the VM. What the VM cannot say is whether the *thresholds* are the right numbers -
whether 2 timeouts is the right count when each timeout costs 300 s, whether a 30 s cooldown is right
when a real Outlook takes minutes to unwedge. Those are judgements about real wedges, and this repo
has observed exactly one (2026-08-15, two searches hanging for the full 1800 s client timeout).

### 3.3 The index tier: runs, and proves a different thing

`[V]` With Corpus A indexed, most of `LiveIndexSearchTests` becomes runnable. What changes:

- `MaxQueryMs = 2000` is asserted on eight queries. Against one 20k-item store on an idle VM this is
  trivially met; against ~160k+ items across five stores with a working indexer it is the actual test.
  **The assertion survives and stops being a test.**
- `ProbeParity_Top5Email_HitsUnder2s` needs 5 mail rows: fine.
- `FilterShapes_ReadAndAttachmentFlags_WorkUnder2s`: the unread half works (the corpus carries a read
  mixture). The attachment half asserts over an empty collection and passes vacuously.
- `SenderFilter_...` fails outright (no senders).
- `Staleness_SelfReportsPlausibleFrontier` works but only ever reports a frontier that is exactly as
  old as the corpus anchor. It can never observe the lag it exists to describe.
- `[I]` Whether one PST can be kept OUT of the Windows Search index while another in the same profile
  stays in it is the single load-bearing assumption of the whole three-store design, and **it is
  untested**. Indexing Options exposes Outlook as a selectable tree, so per-store exclusion looks
  available, but "looks available in the UI" and "the indexer honours it and does not silently
  re-crawl after a profile change" are different claims. **Verify this before building anything
  else**; if it does not hold, the three-store layout collapses back to two machines or two profiles.

### 3.4 The completeness oracle degrades in a way worth naming

`[V]` `LiveCompletenessOracleTests` builds ground truth by a full COM walk of the hub and compares it
against index results. Two problems on the VM:

1. Its header records a live-bitten finding: **`System.Search.Contents` indexes more than the COM
   plain-text subject and body** - HTML-only tokens (link URLs, alt text), attachment text, address
   fields - verified on a real mail where the term was absent from subject+body and present for the
   index. The oracle *tolerates* these as precision-positive extras. **On a corpus of plain-text
   bodies with no attachments, no HTML and no addresses, the extras cannot occur, so the tolerance
   branch is never exercised and the oracle's hardest case disappears.**
2. It was written for a "tiny designated test-hub store". A full COM walk of a 20,000-item hub is a
   different cost class. `[I]` At the local rate of ~700-1,200 items/s that is 17-30 s per walk, which
   is survivable, but it is no longer the cheap ground truth it was designed as.

### 3.5 Thread and conversation coverage collapses

`[V]` Corpus subjects each carry a unique ordinal, so every conversation has exactly one member.
`LiveMailServiceTests.Thread_IndexPath_AndComFallback` walks a real production conversation
specifically because "any single one of them can be unwalkable right now" - it needs several
candidates and a genuinely multi-member thread. On the corpus it will find one-member threads and
prove nothing about the conversation-graph fallback. **The corpus needs deliberate conversation
families** (see section 5).

### 3.6 The sweep tests: the interesting one is not the one you would expect

`[V]` `LiveSweepScopeTests.DefaultSweep_CoversTheArrivalPathFolders_WithinBudget` is `Portable` and
asserts `< 5000 ms`. But it sweeps a **15-minute** window with `perFolderCap: 50`, so it finds almost
nothing on either machine. What it actually measures is the **fixed per-folder cost multiplied by
store count**, not the per-item cost. That makes it a good VM test (3 stores x 4 folders) and a good
real-profile test (5 stores x 4 folders plus delegates), and the two numbers are directly comparable.
It is the one latency assertion in the live tier that survives the move intact, and it is worth
knowing that, because it means **the per-folder half of the sweep model is VM-provable and the
per-item half is not**.

`[V]` `LiveResumableScanTests` (4 methods, the acceptance the project is blocked on) is fully
VM-runnable **and needs the corpus to be the hub** - all four early-return green with "hub corpus fits
one page" otherwise.

`[V]` `LiveTableSortProbeTests` (2 methods) is fully VM-runnable. Note that its subject, the sort
defect, was settled today by `03a0857` on the real profile, and see 4.3 for why that matters.

### 3.7 The store-count tripwire proves less, and knows it

`[R]` `live-tier-on-the-vm.md` section 6 already works this out: a corpus store is the **wrong shape**
for the identity half of the guard. All four populated corpus folders (Inbox 10,912 / Sent 4,964 /
Deleted 2,467 / Junk 1,663) exceed the 500-items-per-folder identity budget and fall back to counts,
and Deleted Items and Junk are excluded from identity anyway. **The bystander store is the only place
the rewritten half of the guard runs at all, so it must have a few hundred items in it, not zero.**
`[R]` And what broke the guard in the first place - a per-store census exceeding the 3-minute STA
budget - is an Exchange delegate-store cost that the VM cannot reproduce even in principle.

---

## 4. Question 3: insight that is not a test

This is the part I would push back hardest on if the plan were "move everything to the VM and stop".

### 4.1 The class of defect that only a real profile produces

Five mechanisms, each with at least one defect this project has already paid for:

**(a) Cost that is invisible until it is multiplied by real scale.** Not "slow", but "crosses a
threshold". `[R]` The census: five cross-process calls per item is a fine design at 0.3 ms per call
and a suite-refusing failure at 12 ms per call, and nothing about the code changes in between. The
defect lives in the product of a per-call cost the code does not control and an item count the code
does not choose. **A VM fixes both factors at values that hide it.**

**(b) Aggregate cost that needs a real store count.** `[R]` The sweep budget: 12 s per store was
measured on the VM and was not the problem; 12 s x 5 stores against a 30 s budget was. The VM
supplied the coefficient and the real profile supplied the multiplier, and neither alone was the
finding.

**(c) Payload size driven by real content.** `[R]` The frame high-water. The 432 KB measured on the
real profile turned out to be bounded by the 30 s timeout rather than by any item cap, which is
itself a finding a single measurement could not produce - it needed the real number, the corpus
number and the knowledge that a timeout sat between them. **A measurement whose bound is another
defect reads exactly like a healthy measurement.**

**(d) Population questions.** `[R]` H3: whether any real mail lacks a usable delivery time. Answered
by counting 43,048 items across five stores and twenty arrival-path folders and finding zero. A
generated corpus can only answer "does the code handle this shape", never "does this shape occur".
**The decision not to fix H3 rests entirely on a population count, and no VM can ever supply one.**
The session log is careful about this and says so: that profile is entirely Exchange-delivered mail
and mounts no PST, and H3's hypothesised shape is imported or restored mail, so the row stays open
and downgraded rather than closed.

**(b2) Cost that is invisible until it meets a real folder.** `[V]` A fifth mechanism sits between
(a) and (b) and deserves its own line, because it is the cheapest kind to hit: the teardown sweep's
folder set deliberately EXCLUDES the Archive folder (id 39) because "business-store archives hold
~100k items and a LIKE count over them takes minutes - adding 39 here made the cross-account sweeps
time out (live-bitten 2026-07-26)". A restriction that is instant on a 10,000-item folder takes
minutes on a 108,144-item one, and nothing about the query changes. **A generated corpus of 20,000
items per store is, by construction, below every threshold of this kind that this project has hit.**

**(e) Provider and object-model behaviour that differs by store type.** `[R]` Four already paid for:
`SendUsingAccount` is a PROPERTYPUTREF whose plain dynamic assignment silently no-ops, so the
Phase-2/3 seeds went out from the wrong account; `GetItemFromID` rejects the index-decoded 24-byte id
on cached Exchange; delegate folders are nested in Outlook and flat in the index; and
`System.Search.Contents` indexes strictly more than the COM plain text. **All four are cases where
the code was correct against the documentation and wrong against the machine.**

### 4.2 Which of this session's findings would have been missed

Honest, one line each.

| Finding | Would the VM have found it? |
|---|---|
| Per-store census exceeding the 3-minute STA budget on a delegate Exchange store | **No.** It needs 12 ms per COM call. A local PST is ~0.3 ms. |
| Sweep budget needing a real store count (12 s x 5 = 60 s against a 30 s budget) | **Half.** The VM produced the 12 s. Only the real profile produced the 5, and only the real profile showed the timeout. |
| Frame high-water and the discovery that the 432 KB figure was timeout-bounded | **No, as a pair.** The VM produced 10.7 MB; the reinterpretation needed the real-profile number beside it. |
| H3 population = 0 across 43,048 items | **No, and never.** |
| Exhaustive scan reaching 3 folders of 32 in 105 s | **No.** The VM's three stores of four folders finish easily. |
| `OutlookAvailabilityCiTests` search at 139.1 s against a 40-hour-old Outlook | **No.** It needs a long-running, busy, real Outlook. |
| **The `Table.Sort` namespace defect** (`03a0857`, today) | **Partly, and this is the important row.** The probe is `LiveTier=Portable`, so the VM could have run it and would have shown that the explicit name applies and the namespace form is refused. What the VM could **not** have produced is the evidence that made it a shipped fix on the same day: five real stores, five of five refusing the namespace form, and the top row coming back as **2022, 2024 and 2025 depending on the store** with the namespace spelling versus today's mail with the explicit one. A single synthetic store of known dates would have shown "the sort was refused"; five real stores of years-deep history showed *what the user was actually getting instead*. |

**Six of seven would have been missed or halved.** That is the number the arrangement has to answer
to, and it is why "VM by default plus a gate" is the right frame and "VM only" is not.

### 4.3 A seventh thing, which is not a defect but is the reason the gate must be more than a smoke test

`[R]` `StaleIndexNoticeMinutes` is set at the p90 of 177 index-frontier probes on a live mailbox.
That is not a bug found; it is a **constant calibrated from a distribution that only exists on a real
machine**. If the maintainer's mail volume changes, or Microsoft changes the indexer, that constant
goes wrong silently and no test anywhere will say so. The gate's job includes **re-taking
measurements**, not only running assertions. A gate that only runs tests cannot notice a constant
drifting away from its evidence.

---

## 5. The seeding question: what the corpus must contain that it does not today

`[V]` What `corpus-build` produces today, read from `ComCorpusMailbox` and `CorpusPlan`:

- `items.Add(0)` = `IPM.Note`, always. **One message class.**
- `Subject` (tag + corpus id + ordinal), `Body` (plain text, generated), `UnRead`, then
  `PR_MESSAGE_FLAGS` (clears MSGFLAG_UNSENT, sets MSGFLAG_READ per plan), `PR_MESSAGE_DELIVERY_TIME`,
  `PR_CLIENT_SUBMIT_TIME`. **Nothing else.**
- **No sender, no recipients, no attachments, no HTML body, no categories, no flags, no importance.**
- Four folders only: Inbox (6), Sent Items (5), Deleted Items (3), Junk Email (23), by weight. **No
  subfolders and no CLI option to create any**; the only folder the builder creates is a substitute
  when a PST refuses `GetDefaultFolder` for Junk.
- Body sizes 200 B to 640 KB in five weighted classes, deliberately running past
  `SweepBodyCharsCap` (500,000 chars) so ~1 in 5 of the top class trips the per-item cap. **This part
  is good and should not change.**
- Dates in five bands cut at the 1-day, 7-day and 60-day marks the sweep and scan are measured on.
  **Also good, subject to the anchor-decay problem in 2.5.6.**

### 5.1 What to add, in priority order

**P0 - blocks tests that would otherwise run.**

1. **Senders and recipients.** Set `SenderName`/`SenderEmailAddress` (via `PropertyAccessor` on
   `PR_SENDER_*` and `PR_SENT_REPRESENTING_*`) and `To`/`CC` from a small synthetic directory of, say,
   40 correspondents with a Zipfian distribution. Unblocks
   `LiveIndexSearchTests.SenderFilter_...`, the `System.Message.FromAddress` index column, per-column
   `CONTAINS`, `TryDiscoverStoreScopeByAddress`, and the sender half of `subjectOnlyProbe`.
2. **Attachments, of several kinds.** `[R]` The attachment-recall fix exists because 709 of 3,139
   attachment rows on the real profile (22.6%) were images, embedded messages and `.ics` invites that
   the old `System.Kind IN ('email','document')` predicate dropped. The corpus needs at least: a PDF,
   a DOCX, a PNG or JPG, an embedded `.msg`, an `.ics`, and one multi-attachment mail. Unblocks
   `LiveAttachmentKindRecallTests` (5), `LiveMailServiceTests.AttachmentHit_ReadParent_SaveToScratch`,
   the `HasAttachments` filter, and `save_attachment` end to end.
3. **HTML bodies.** A share of items with `HTMLBody` set, including quoted-reply structure, an inline
   `cid:` image, and a link whose URL contains a term absent from the visible text. That last one is
   what reproduces the `System.Search.Contents`-indexes-more-than-COM behaviour the completeness
   oracle tolerates and would otherwise never see.
4. **Conversation families.** Deliberate threads: N items sharing a `ConversationTopic` with
   `Re:`/`Fw:` prefixes, spread across Inbox and Sent Items, some with a member in a second store.
   Unblocks `Thread_IndexPath_AndComFallback` and the `unwalked_store` coverage code.
5. **The `subjectOnlyProbe` shape as a first-class corpus feature.** A population whose term is in the
   subject and the sender address and provably **not** in the body stream. Then
   `machineProfile: "Portable"` can carry a real `subjectOnlyProbe` block instead of omitting it, and
   the seven `Requires=ProbePopulation` tests move.

**P1 - the shapes this project has fixed blind.** Every one of these is code that shipped with no
test able to produce its input.

6. **Message-class diversity.** `MailItemAdmission.ClassesTheOldFiltersDropped` names ten classes:
   `REPORT.IPM.Note.NDR`, `REPORT.IPM.Note.IPNRN`, `REPORT.IPM.Note.DR`,
   `IPM.Schedule.Meeting.Request`, `.Canceled`, `.Resp.Pos`, `.Resp.Neg`, `.Resp.Tent`, `IPM.Post`,
   `IPM.Sharing`. `[V]` The only test that touches them is `T1/ItemClassAdmissionTests`, which drives
   the strings through `MailItemAdmission.Admits`, **a method that returns `true` unconditionally**.
   That is a pin on a decision, not a proof of behaviour. **Nothing anywhere proves that a real NDR
   sitting in a real Inbox is returned by all three tiers, carries a legible `itemClass`, and produces
   the "not ordinary mail" advice sentence.** `[I]` Most of these can be manufactured with
   `PropertyAccessor.SetProperty` on `PR_MESSAGE_CLASS` after `Save()`; a genuine NDR cannot be
   generated but a `REPORT.IPM.Note.NDR`-classed item with the right properties is enough for every
   code path this product has.
7. **Items with no delivery time.** `[R]` H3. The generator already knows how to make them (that is
   what `--allow-undated` produces) and refuses to by default, correctly, because an undated corpus
   makes every window select the same population. What is needed is a **small, deliberate, tagged
   slice** of undated items, not an undated corpus: enough that the DASL restriction's behaviour over
   an absent property can be observed rather than reasoned about from MAPI's "undefined".
8. **Deep folder trees.** `[V]` `FolderWalkAbsoluteCap` is 10,000 and
   `OutlookComSession.FolderWalkDepthGuard` is 64. `[R]` TODO already records that proving
   `depthLimitReached` needs a temporary build with the guard lowered, because no real profile is
   64 deep either. **A generator that can build an arbitrary tree makes the real guard testable at its
   real value**, which no lowered-guard build can do. Also needed for `list_folders` paging past 1,000
   folders and for its stable-order comparator.
9. **A folder that fails to open.** `[V]` The sweep counts `FoldersSkipped` separately from
   `FoldersAbsent` precisely because conflating them made every search on such a profile report itself
   degraded, and `LiveSweepScopeTests.FolderScopedSweep_UnknownFolder_DegradesWithAdvice_NeverThrows`
   covers only the *named-but-nonexistent* case. `[I]` A folder that exists and throws on open is
   producible: a folder whose `DefaultItemType` is not `olMailItem`, or one made unreadable through
   permissions, or the `wedgedEmpty` shape `LiveOutlookTestMailer.CountLiveTestFolders` already has a
   counter for. **This one needs a design decision, not just a generator flag** - see question 4.
10. **A store whose display name cannot be read.** `[V]` `StoreNaming` exists entirely for gap G2:
    every scope, bucket, label and refusal message is keyed by `DisplayName`, so a store whose
    `DisplayName` read throws was dropped from `list_folders`, from `list_accounts`, and counted as
    skipped in the sweep total while landing in no per-store bucket. The fix labels it
    `(unnamed store N)` and refuses to let that label round-trip as a scope.
    `[I]` **I could find no way to produce this shape deliberately, on any machine.** It is the one
    item on this list where the honest answer may be "the code is correct by inspection and stays
    that way", or where a test seam (an injectable store enumerator in the COM session) is cheaper
    than a mailbox that can produce it.

**P2 - fidelity.**

11. **Non-ASCII bodies.** `[V]` `SweepBodyBytesBudget` is counted in BYTES because
    `JavaScriptEncoder.Default` escapes every non-ASCII character as a six-byte `\uXXXX`, so the same
    text costs 1 byte or 6 depending on language. The whole justification for a byte budget rather
    than a character budget rests on CJK and Cyrillic mail, and the corpus is ASCII.
    `EncodedBodyByteCeiling` is an over-estimate that T1 checks against the real serializer, so the
    logic is pinned; what is unmeasured is a real frame built from real multi-byte bodies.
12. **A read/unread and importance/flag mixture wide enough for the filter shapes.** Partly present
    (read state is planned); importance, follow-up flags and categories are not.
13. **Items in the Outbox and in Drafts on purpose.** `[R]` The first corpus build put 5,532 items in
    the Outbox and all 40,000 in Drafts by accident, and the sweep covers neither, which is how a
    20,000-item corpus yielded 6 swept items. Both folders are now in the teardown scan. Having a
    small deliberate population in each would make the "a folder nothing sweeps is a folder items can
    be stranded in" lesson a test rather than a comment.

### 5.2 One thing the corpus should NOT try to be

Do not synthesise a delegate store, and do not synthesise Exchange EntryID semantics. Both were
already decided against for delegates and the reasoning extends to EntryIDs: a fake that satisfies the
test while missing the real behaviour is worse than a documented gap, because the gap is visible in a
list and the fake is visible to nobody.

---

## 6. The arrangement

### 6.1 Primer

The question is not "VM or real profile". It is **what a pre-release gate against the real profile
has to contain to be worth the risk and the wall clock**, given that the VM will, on the numbers
above, prove roughly 90 of 116 live methods, all 1,391 T1 methods and all 100 non-live T3 methods,
and will prove them all against latency one and a half to two orders of magnitude too fast.

A gate that re-runs what the VM already proved is a slow ritual that will be skipped under time
pressure, which is the failure mode that matters: a gate nobody runs is worse than no gate, because
the plan still claims it exists. A gate that is too narrow is a gate that misses the class of defect
in section 4.1, which is the class this project keeps finding.

Four directions, then a recommendation.

### 6.2 Direction A - VM only, no real-profile gate

Move everything, accept the 15-method hard floor as permanently unproven, and rely on user reports
for the rest.

- **For:** simplest. Zero production risk. The maintainer's mailbox is out of the loop completely and
  provably. No profile-switching runbook to maintain beyond the corpus one.
- **Against:** six of this session's seven findings would have been missed. Every latency constant in
  group 2 of section 3.2 becomes uncalibrated the moment anything changes. The delegate paths - which
  is where this product has been surprised most - become dead code that still ships.
- **Cost:** near zero to run; the cost is entirely deferred and lands on users.

### 6.3 Direction B - VM by default, full `Category=Live` on the real profile before each release

The current documented plan: VM for the cycle, `--filter "Category=Live"` on the real profile as the
gate.

- **For:** nothing is left unproven that can be proven. Simple to describe. No new trait vocabulary.
- **Against:** it is 116 methods, `[R]` a ~27-minute tier run plus fixture setup, and it re-proves on
  production the ~90 the VM already proved that morning. It requires nobody to touch five mailboxes
  for the duration, or the tripwire fails and hands over EntryIDs. It writes to a real mailbox and
  sends real mail for the sake of assertions the VM already made. **The re-proved 90 carry all of the
  risk and add none of the information.**
- **Cost:** ~30-40 minutes of exclusive mailbox use per release, plus the standing risk of any write
  path regressing while pointed at production.

### 6.4 Direction C - VM by default, plus a narrow **measurement** gate on the real profile

A named, versioned subset that runs on the real profile and does two things the VM cannot: exercise
the five impossible mechanisms, and **re-take the measurements the constants are derived from**.

Concretely, the gate is:

1. **The 15 hard-floor methods** (delegate semantics, transport arrival, Exchange EntryID) plus the
   2 arity methods. Read-only wherever possible.
2. **A measurement pass**, which is not a test tier at all but a script that records and diffs against
   the last release: whole-store 7-day sweep elapsed and per-store breakdown; Inbox-only and
   Inbox-with-subfolders exhaustive scan elapsed; a 60-day whole-store scan's `foldersScanned` and
   `elapsedMs`; the census elapsed per store and `folders fell back to counting`; index frontier age
   sampled N times; `largestFrameBytes`; `sweep.sortRefusedFolders` (which should now read **zero** -
   `03a0857` made that claim checkable and nothing has checked it yet).
3. **One write exercise**: the move/archive batch the maintainer already asked for, hub-only, under
   the existing guards.
4. **A drift check on the four Exchange-behaviour findings**: does `GetItemFromID` still reject the
   decoded id; is the delegate index still flat; does `SendUsingAccount` still need the putref
   accessor; does `System.Search.Contents` still over-index. Each is three lines in the existing live
   harness and each has burned this project once.

- **For:** it covers exactly what the VM cannot, it is short enough to actually run, and the
  measurement pass turns "these constants were right once" into "these constants are still right".
  It also makes the constants' evidence a release artifact rather than a doc paragraph.
- **Against:** it needs a third trait value (or a `RealProfileGate` category), it needs the
  measurement script written, and someone must read the diff rather than a red/green. **A measurement
  gate is only as good as the person reading it**, which is a real weakness.
- **Cost:** `[I]` roughly 10-15 minutes of exclusive mailbox use, versus 30-40 for direction B, plus
  a one-off build cost for the measurement script and the trait.

### 6.5 Direction D - two VMs, one of them joined to a real Exchange test tenant

Add a second guest joined to a throwaway Microsoft 365 tenant with two mailboxes, one granted
delegate access to the other, and real transport between them.

- **For:** delegate semantics, real transport arrival, cached-Exchange EntryIDs, index-flattening and
  real server latency all become reproducible with **zero production risk**. It is the only direction
  that closes the 15-method hard floor. It also makes the sends real, which retires the Outbox
  problem in 2.3(b) entirely.
- **Against:** a tenant costs money and administration; cached-mode OST behaviour on a small fresh
  mailbox is not the same as on a 108,144-item Archive, so it closes the *semantics* gap but only
  narrows the *scale* gap; and it is the largest amount of new infrastructure on this list.
- **Cost:** `[I]` two M365 Business Basic seats is on the order of tens of euros a year, plus initial
  setup and a second VM's disk and checkpoints, plus keeping a second guest patched.

### 6.6 Recommendation

**C now, D later, and never B.**

**Adopt direction C**, because it is the only one that answers the actual finding of section 4:
six of seven defects came from real latency, real scale, real population and real provider behaviour,
and only two of those four are addressed by running more *tests*. The other two are addressed by
taking *measurements*, and a gate built as a test filter cannot take them. C is also the only
direction whose cost goes **down** relative to today: it replaces a full-tier production run with a
short one, and it replaces the current situation - where every ordinary verification run silently
reads the production mailbox - with a deliberate, scheduled, bounded one.

**Plan for direction D as the next infrastructure item after the VM is working**, not instead of it.
It is the only way to retire the 15-method hard floor, and the delegate paths are where this product
has been wrong most often. But it is strictly a second step: a delegate mailbox on a fresh tenant
with nothing in it proves the semantics and not the scale, so it complements C rather than replacing
it. Do not let it block the VM.

**Reject B** on the specific ground that re-running 90 already-proven methods against a production
mailbox is pure risk with no information, and that the size of it is what will cause it to be skipped.

**Reject A** on the ground that this session alone produced six findings it would have missed.

**What the gate must include to be worth running,** stated as a checklist rather than prose, because
this is the part the maintainer asked for directly:

- [ ] The 15 impossible methods, plus the 2 arity ones, selectable by a single filter and pinned by a
      reflection test the way `LiveTierInventoryTests` pins the Portable set today. **If it is a
      filter string in a runbook, it will drift.**
- [ ] The seven measurements in 6.4(2), written to a file, diffed against the previous release, with
      a stated tolerance per measurement so the diff has a verdict and not just numbers.
- [ ] `sweep.sortRefusedFolders == 0` asserted explicitly. It is the newest claim in the codebase and
      nothing has checked it on a real profile since the fix.
- [ ] One hub-only move/archive batch exercise.
- [ ] The four Exchange-behaviour drift probes.
- [ ] A count of assertions that actually fired, so a `PROVED NOTHING` or an early return cannot pass
      as coverage. On a Production profile `RequireProductionPopulation` already throws; the early
      returns in `LiveResumableScanTests` and the vacuous `Assert.All` in
      `LiveIndexSearchTests.FilterShapes_...` do not.
- [ ] The store-count tripwire and the signature snapshot, unchanged. They are the reason a
      production run is survivable at all.

**And the VM side must include** (these are prerequisites, not extras):

- [ ] Verify that one PST can be excluded from the Windows Search index while another in the same
      profile is included, and that it stays excluded across a profile switch. **Everything else in
      the three-store design rests on this.**
- [ ] Name the hub PST as the dummy account's SMTP address, and confirm Outlook accepts it.
- [ ] Point the dummy account's delivery store at a fourth throwaway PST, not at the corpus.
- [ ] Give `expectedStoreDisplayNames` a companion list for "expected to be in the index", or the
      index tier cannot coexist with a deliberately unindexed store.
- [ ] Put a few hundred items in the bystander so the tripwire's identity half runs.
- [ ] Decide the Outbox question (section 7, question 3) before the first send-path test runs.
- [ ] Add a re-anchor step to the corpus runbook, or the date windows expire silently.

---

## 7. Open questions, each with its own primer, directions and recommendation

### Question 1 - what to do with the sixteen mislabelled T3 files

**Primer.** 92 of their 100 methods are honestly machine-independent; 8 reach the real Outlook, real
index or real user data. Today all 100 run under `--filter "Category!=Live"`, which is the CI command,
and none of them sits in a guarded collection. The interim policy excludes the whole tier by name
(`FullyQualifiedName!~Tests.T3.`), which throws away 92 good tests to avoid 8.

1. **Keep the interim filter.** Cheapest, and wrong: it silently drops 92 wire-conformance tests,
   including every schema and description pin, which is exactly the coverage that catches an
   accidental tool-surface change.
2. **Move the 8 into the live tier** with `Category=Live` and a `LiveTier` value, put them in a guarded
   collection, and let `Category!=Live` mean what it says again. The 92 stay in CI.
3. **Split the files.** Move the 8 methods into a new `T3/LiveEnvironmentProbeTests` and leave the 16
   files otherwise intact, so the `...CiToolShapeTests` names become true.
4. **Gate them on an environment variable** (`OUTLOOKAI_ALLOW_REAL_OUTLOOK=1`), skipping otherwise.
   Keeps them in CI where there is no Outlook and skips them on a developer box.

**Recommendation: 2, implemented as 3.** The trait is what makes the claim checkable by
`LiveTierInventoryTests`; the file split is what makes the names stop lying. Option 4 is tempting and
should be rejected: a test that skips itself based on the environment is how the tier ended up
undifferentiated in the first place.

### Question 2 - how much class-level `Requires` over-attribution to fix

**Primer.** 30 of 36 live classes declare `Requires` at the class level, so it reads as the union of
what any test in the class needs. `LiveIndexSearchTests` carries `DelegateStore` for all ten of its
methods because one of them needs a delegate store. That is what turns a genuine floor of ~15 methods
into a reported 96.

1. **Push every `Requires` down to the method.** Most accurate, ~30 classes to edit, and it makes the
   VM subset grow by roughly 75 methods on paper before a single one is made to pass.
2. **Push it down only for the classes that straddle the boundary** (`LiveIndexSearchTests`,
   `LiveMailServiceTests`, `LiveSendTests`, `LiveDraftOptionsTests`, `LiveUpdateDiscardTests`,
   `T3/LiveMcpToolShapeTests`). Six classes, covers most of the distortion.
3. **Leave it and track the true floor in a document.** Zero code change, and the document drifts,
   which is the failure `LiveTierInventoryTests` was built to prevent.
4. **Add a third `LiveTier` value, `VmCapable`,** meaning "runs on the VM once it is seeded", and keep
   `Portable` meaning "runs on a bare test machine today". The two questions are genuinely different
   and one trait is currently answering both.

**Recommendation: 4, then 2.** The trait vocabulary is the distortion, not the class-level
declaration: `Portable` was defined as "PST stores, **no mail accounts**, nothing indexed", and the
decided VM has an account and an index, so the word no longer describes the machine being built.
Adding `VmCapable` lets the inventory test keep enforcing "say why" while the subset grows honestly.
Then push `Requires` down for the six straddling classes.

### Question 3 - the Outbox, and whether a local SMTP sink is needed

**Primer.** The decided VM enables send on an unroutable account so a send can never leave the
machine. `[V]` But the mailbox-safety contract requires every live run to end with **zero tagged
artifacts including in the Outbox**, and `LiveOutlookTestMailer`'s sweep folder list includes the
Outbox for exactly that reason. A queued send that can never leave is a permanent tagged Outbox item.
Separately, six tests need mail to actually **arrive** (section 2.3(b)), which an unroutable account
can never deliver.

1. **Unroutable account, and teach the sweep that an Outbox artifact on a Portable machine is expected
   and deletable.** Cheapest. Weakens the one guard that catches a real stranded send, on the machine
   where sends are most likely to be exercised.
2. **Unroutable account, and never enable any test that reaches `Send()`.** Keeps the guard intact.
   Means the send path is only ever proven on production, which is the highest-consequence path in the
   product.
3. **Local SMTP sink** (a loopback listener that accepts and discards, or writes to a maildir).
   Sends complete, the Outbox drains, the zero-artifact sweep stays strict. **But mail still does not
   arrive back**, so `LiveInboxArrival` still times out: a sink alone unblocks the *send* half and not
   the *round-trip* half.
4. **Local SMTP sink plus a local IMAP/POP delivery back to the same account**, so a self-addressed
   send genuinely round-trips. Unblocks all six arrival-dependent tests and makes the fresh-mode proof
   runnable on the VM. It is a real mail server on the VM (hMailServer, Mailpit with delivery, or
   similar).
5. **Defer to direction D** and let the real test tenant provide transport.

**Recommendation: 4, scoped small, unless direction D is happening soon - then 3 as a stopgap.**
`LiveFreshModeTests` is the acceptance for the single most distinctive thing this product does
(finding mail before the index has it), and it is currently provable only on production. A loopback
mail server is a well-understood piece of infrastructure and it is entirely inside the VM. Option 1
should be rejected outright: weakening a zero-artifact guard on the machine where destructive tests
run unattended inverts the whole reason the guard exists.

### Question 4 - how to produce a folder that fails to open, and a store with no readable name

**Primer.** Both shapes have shipped fixes with no test able to produce them (`FoldersSkipped` versus
`FoldersAbsent`; `StoreNaming`'s `(unnamed store N)` label and its refusal to round-trip as a scope).
`[I]` The folder case looks producible; the store-name case I could not find any way to produce on any
machine.

1. **Produce them for real.** A permissions-denied folder, or a folder whose `DefaultItemType` is not
   mail, gets the folder half. The store half has no known recipe.
2. **Add a test seam**: an injectable store/folder enumerator in `OutlookComSession` so a fake can
   throw on `DisplayName`. Makes both testable in T1 at full fidelity. Costs a seam in the COM layer,
   which is the layer this project has deliberately kept thin.
3. **A fault-injection environment variable in the COM host**, the way `OUTLOOKAI_COMHOST_FAULT`
   already works for the supervision tests. Precedent exists, it is already CI-safe, and it needs no
   production-code seam beyond the one that is there.
4. **Accept both as documented gaps**, correct by inspection.

**Recommendation: 3 for the store-name case, 1 for the folder case.** The fault-injection hook is
already a shipped, tested pattern in this repository and extending it costs almost nothing; a real
unreadable folder is cheap and is more faithful than any fake. Option 4 is the status quo and it is
what "fixed blind" means.

### Question 5 - does the measurement gate need a human reader, and what if there is not one

**Primer.** Direction C's value depends on someone comparing this release's measurements against the
last. `[R]` This project has already been bitten by a measurement whose bound was another defect (the
432 KB frame high-water, bounded by a timeout and read as headroom), which is precisely the kind of
thing a tolerance check would pass and a reader would catch.

1. **Human reads the diff.** Highest quality, lowest reliability.
2. **Tolerances in code**, failing the gate on a >2x move in any measurement. Reliable, and it will
   miss the "bounded by another defect" shape entirely, because that shape is *stable*.
3. **Both: tolerances fail the gate, and the raw table goes into the release notes** so it is read
   whether or not anyone intended to.
4. **Record only, never fail.** A log nobody opens.

**Recommendation: 3.** The tolerance catches the drift; putting the table where it will be seen
catches the stable-but-wrong case. `[R]` The precedent is good: `sweep.sortRefusedFolders` was added
specifically so a claim becomes something a run can check, and this is the same move applied to the
constants.

---

## 8. Things found along the way that are not part of the three questions

1. **`Docs/live-tier-on-the-vm.md` is one test behind** (115/19 versus the actual 116/20) and predates
   both the dummy-account and three-store decisions. Its section 2.3 recommends a **two**-PST layout
   with the corpus as hub, which the 2026-08-20 three-store decision supersedes. Its section 7
   "Known limits" is still accurate and still useful.
2. **`C:\Users\jori\Downloads\tmp-aitrace\failing-now.txt`** holds stale output from before `03a0857`:
   `T1/SweepSortPropertyTests` failing on the namespace-reference sort. That defect is fixed; the file
   is misleading if read now and should be deleted with the rest of the scratch folder.
3. **`corpus-measurement-plan.md` step 5 (exhaustive scan throughput) has never been run**, on either
   machine. It is the one measurement that would settle whether `ExhaustiveTimeBudgetMs` at 600 s is
   sized correctly, and it is cheap and read-only on the VM. Worth doing first, since the plan already
   says exactly how (use a term that matches nothing so the scan runs to the end of its budget).
4. **`Assert.All(withAttachments.Hits, ...)` in `LiveIndexSearchTests.FilterShapes_...` passes
   vacuously on an empty collection**, on any machine with no attachment-bearing indexed mail in the
   first configured store. That is a pre-existing weak assertion, not a VM problem, and it is a
   one-line fix (`Assert.NotEmpty` first, or an explicit `PROVED NOTHING`).
5. **`ProbeParity_DateRangeQuery_HitsUnder2s` asserts "no mail indexed in the last 30 days" as a
   failure.** On the real profile that is a real health assertion. On a checkpoint-restored VM it is
   an assertion about the age of the checkpoint. It needs to move to a corpus-relative window.
