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
| Second PST on the testbed | Add one via the tested helpers | **queued** |
| Live tier | Move the intermediate tier to the VM; keep the ability to run everything against the real system before a release | **queued** |
| `Stick-Test` VM + scratch | Delete both | **done** |
| Installed MCP server | Leave disabled until the release | **done, nothing to do** |
| Exhaustive scan | Resumable walk with a continuation token | **queued** |
| `thread` store asymmetry | Derive the warning from Outlook's store list; also scan for the same asymmetry elsewhere | **queued** |
| Timeout defects | Fold all three into the timeout-raising pass | **shipped** (uncommitted) |
| COM host kill | Keep the hard kill, document it, add a brief wait before killing, make the kill outcome-aware | **shipped** (uncommitted) |
| `top` ceiling | Leave at 100; rely on resumption | **decided, no work** |
| Remaining gap-map rows | Clear **all** of them before the release | **queued** |
| Work order | Infrastructure first: corpus, second PST, live tier on the VM | **in progress** |
| `update_draft` | **(d)** make it re-entrant: record intent first, so a retry completes rather than repeats | **shipped** (uncommitted) |
| Sweep timeout | **(d)** make expiry graceful **and** distinguish budget expiry from unresponsiveness at the supervisor | **shipped** (uncommitted) |
| H3 (undated mail invisible to the sweep) | Check whether DASL can express "or the property is absent" first; failing that, report it; full fallback enumeration only if it proves common | **answered by measurement - NOT fixed, see section 3** |

## 2. Timeout values - SHIPPED (uncommitted) on 2026-08-19

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

## 4. VM state (`OutlookAI-TestVM`)

- Guest credentials for PowerShell Direct: `vmadmin` / `***REDACTED-CREDENTIAL***`.
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
- **`item_cap_unsorted` fired on every sweep of that corpus**, i.e. `Table.Sort` genuinely does not
  apply on that store and the answer correctly reports the cap cut arbitrarily rather than claiming
  the oldest mail is what is missing. The 2026-08-18 H2 fix, observed working against real data
  rather than in a test.
- **Corpus build throughput:** 50.9 items/s without the move rung. With the move rung each item is
  written twice, so budget roughly double.

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
