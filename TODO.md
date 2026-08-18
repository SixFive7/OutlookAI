# TODO

- [ ] **Restore the installed MCP server — it is deliberately disabled right now.**
  On 2026-08-16, while developing the COM-host work, the installed executable was moved
  aside so that no Claude Code session could start a stale build and re-activate Outlook
  behind the tests. **Until this is done, every Claude Code session fails to start
  `outlookai`** — that is expected, not a fault.

  Current state on the affected machine:
  - `%LOCALAPPDATA%\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe` — **absent**
  - `%LOCALAPPDATA%\OutlookAI\Setup\McpServer\OutlookAI.McpServer.exe.disabled` — the old
    **v3.0.1.321** binary, kept only so the move was reversible

  Steps, in this order:
  - [ ] Cut the release first (the intent is to restore to the NEW version, not to move
        the old binary back). Version bump is a judgement call; the reading at the time
        was **minor** (3.1 → 3.2): the architecture changed and tool failures now set MCP
        `isError`, which is wire-visible to clients, but no tool name, parameter or
        success payload changed.
  - [ ] Install that release, so the executable exists again at the registered path.
        The registration in `~/.claude.json` still points at
        `${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe`, so nothing
        needs re-registering — the file simply has to be there.
  - [ ] **Install BEFORE restarting Claude Code sessions.** A session started while the
        exe is missing keeps a dead server registration for its whole lifetime.
  - [ ] Delete the leftover `OutlookAI.McpServer.exe.disabled`. The installer will not
        remove it: its `[Files]` rule only copies `publish\*` and never cleans unknown
        files, so it would linger and confuse the next person to look in that folder.
  - [ ] Verify: a fresh Claude Code session starts the server, and `outlook_health`
        reports the new version — the report carries `runningFrom`, which shows plainly
        which build is actually live.

  Worth knowing for the upgrade itself: the installer now stops **both**
  `OutlookAI.McpServer.exe` and `OutlookAI.ComHost.exe` (child first) before replacing
  files. That gap was fixed in the same work — an in-place upgrade will not leave a stale
  COM host holding its own image open.

- [x] **DONE (2026-08-18) - Audit the codebase for other timers and time-based behaviour.**
  The sweep is finished. The three halves, and what each one turned up:

  - **Add-in (`Services\`, `TaskPane\`, root `*.cs`) - clean, and was already clean.**
    `UpdateService` measures "checked 4m ago" from the process-lifetime `Stopwatch`
    (`_sinceStart` / `_checkedAtMs`); `AITaskPane`'s debug-click window is a `Stopwatch`.
    The three remaining wall-clock reads are ABSOLUTE INSTANTS and are correct as they are:
    `UpdateService.LastChecked`, the debug log's `HH:mm:ss` stamps, and the `LastReconcileUtc`
    stamps written by `McpRegistrationService` and `OutlookTuningService` - which are only
    ever displayed, never subtracted from anything. The WinForms `Timer`s
    (`_versionTimer` x2, `McpRegistrationPrompt`) count ticks, not clock time, and are all
    stopped and disposed.
  - **MCP server (`Core\`, `ComHost\`, `McpServer\`, `RemediationTools\`) - clean, and
    already documented.** Six TTLs went to `MonotonicClock` earlier; the five wall-clock
    reads left are absolute on both sides (`AuditLog` line stamps, `SignatureManager` backup
    filenames, `IndexSearchService`'s staleness "as of", `MailService`'s `baseGapStart` DASL
    base, `OutlookComSession`'s temp-file purge against `File.GetLastWriteTimeUtc`). The COM
    host's supervision, breaker and deadline paths hold no `DateTime` at all -
    `Stopwatch.GetTimestamp` throughout. Row in `Docs/magic-numbers.md`.
  - **Live test tier (`T2\`, `T3\`) - THIS is where the defect still lived, and it is now
    fixed.** Thirteen polling loops measured their own timeout as
    `DateTime deadline = DateTime.UtcNow + timeout; while (DateTime.UtcNow < deadline)`, with
    waits of 10 s to 240 s each - including every "wait for the mail to come back" and the
    cleanup loop that decides a mailbox is free of test artifacts. A backwards clock jump
    lengthens all of them by the size of the jump, which on this suite is indistinguishable
    from the hang that already cost a night; a forwards jump ends the artifact-cleanup loop
    with tagged items still in a real mailbox. All thirteen now use `T2/LiveWaitBudget`
    (`Stopwatch.GetTimestamp`). Four hardcoded `180`/`120` second literals became named
    constants that the timeout MESSAGES are built from, which is the same defect
    `LiveSweepScopeTests` had already been bitten by.
  - **Drift guard:** `T1/LiveTierClockDriftTests` scans every `T2`/`T3` source for the shape
    and fails on the fourteenth copy. It asserts that it found files to scan and that its own
    detector still recognises the exact lines it replaced, so it cannot silently switch off.

  What is deliberately still on the wall clock in the live tier, and why: the send instant
  returned by `LiveOutlookTestMailer.SendSelfMail` and its equivalent in
  `Phase5LiveMcpToolShapeTests` (both become DASL sweep-window bases compared against
  Outlook's own `DateReceived`), the `ReceivedOnOrAfterUtc` filters, and the screenshot
  filenames. Each is commented in place.

  ~~The update poll above was found by accident while adding the manual check; nothing has
  ever looked at the set as a whole. Worth one pass to find polls that should be events,
  intervals that no longer match what they wait for, and anything that assumes wall-clock
  time moves forward smoothly.~~

  Known starting points — not a complete list, which is the point of the audit:
  - `Services/UpdateService.cs` — the 10-minute `System.Threading.Timer` poll, the 5-minute
    `HttpClient.Timeout`, and the `Get-Process outlook | Wait-Process; Start-Sleep -Seconds 2`
    in the handed-off installer script.
  - `McpServer/` — the COM-host supervision and health paths carry several timeouts and
    back-off intervals (the wedged-Outlook work added re-check and cool-down periods).

  Still open from the original list, and NOT covered by the clock sweep: the installer
  hand-off script's `Get-Process outlook | Wait-Process; Start-Sleep -Seconds 2` still waits
  a fixed grace period rather than for a condition, and nobody has revisited the 10-minute
  update poll or the 5-minute API timeout against what they actually wait for. Those are
  interval-choice questions, not clock-correctness ones.

- [ ] **Live tier: a test hangs, and an aborted run leaves artifacts behind. Read this before
      the next live run.** (2026-08-18, ~03:00-03:45)

  **State left on the machine:** 7 items tagged `[OutlookAI-McpTest]` remain in the
  `telefonie@xxlnet.nl` hub - **6 in Drafts, 1 in Outbox** - found by a read-only `search` after
  the run was stopped. They are inert: drafts sit there, and the Outbox item is a self-addressed
  test seed, so the worst case is a test mail arriving in the test mailbox. **The next successful
  live run's post-run sweep covers Drafts and Outbox and will delete them** - that is the
  sanctioned cleanup path and it needs no special handling. They were NOT removed by hand: the
  shipped tools cannot (`discard_draft` only touches drafts from its own session, `move_mail`
  refuses Outbox and Deleted Items), and the safety envelope forbids ad-hoc deletion.

  **What happened.** `LiveDisconnectRecoveryTests.OutlookExit_ReleasesHeldRefsInBackground_
  HealthProbes_GatewayReattaches` ran for **22.5 minutes** and was still going. Outlook had been
  quit and restarted (uptime confirmed it), so the test got past the exit it drives and then
  waited on a condition that never became true. The run was stopped rather than left overnight.
  Stopping it is what left the artifacts: the teardown sweep never ran. **That is the lesson worth
  keeping - an aborted live run has no sweep, so the mailbox state it leaves is whatever the last
  test created.**

  **Then the cleanup attempt hung too.** A short subset (`FullyQualifiedName~LiveSweepScope`,
  about a minute earlier the same night) sat for **10 minutes at fixture setup**, before any test
  ran and without ever spawning a COM host. So the second hang is in the fixture's own COM path,
  not in the disconnect test.

  **What is NOT yet known, and should not be guessed at:** whether either hang is a regression
  from that night's work (the coverage attribution, per-store staleness scoping, the not-needed
  sweep verdict, the DASL date fix) or fallout from Outlook's automation being left in a bad state
  by the first abort. Evidence against a regression: the same fixture ran clean at 106/107 earlier
  that night, after the budget-composition work. Evidence for caution: nothing else changed on the
  machine, and the second hang is in a path the first abort could plausibly have poisoned.

  **UPDATE, 04:40 - the fixture-setup hang is REPRODUCIBLE.** A third attempt (the full live tier,
  after the disconnect test was given bounds) sat **15 minutes** at `A total of 1 test files
  matched`, again without ever spawning a COM host. Two independent attempts, same symptom, same
  place - so this is not a one-off.

  **What that does and does not tell us.** Outlook has been up continuously since 03:05 UTC, which
  is BEFORE the abort - so the poisoned-COM-state hypothesis is still the leading one and is NOT
  ruled out by the reproduction. What IS ruled out: the disconnect test itself (this run never
  reached any test) and the new bounds (they cannot fire before setup).

  **Why it was not chased further tonight.** Diagnosing it needs Outlook restarted cleanly, and the
  safety envelope allows a graceful quit only when the Outbox is empty - it is not, because of the
  artifacts above. That is a genuine deadlock for an unattended session: the mailbox cannot be swept
  without a live run, the live run needs a clean COM state, and the clean COM state needs a restart
  the Outbox forbids. It needs a human at the machine, which is where it now sits.

  **Suggested first step, in this order:** close Outlook by hand (or let it send the one Outbox
  item, then close it), start it fresh, then run the short subset alone. If it passes, both hangs
  were abort fallout and the sweep will clear the artifacts on the next full run. If it still hangs
  at fixture setup on a freshly started Outlook, it is a real defect in the COM attach path - and
  the diff to bisect is small, since the tier ran 106/107 clean earlier the same night.

  Diagnostic logs: `C:\Users\jori\Downloads\tmp-aitrace\live-run4.txt` (the 22-minute test, with
  the long-running-test diagnostics that named it) and `cleanup-sweep.txt` (the fixture-setup hang).

- [x] **RESOLVED 2026-08-18 11:45 - the live tier was never hanging. I misread it, twice.**

  The full tier takes **26.8 minutes** and passes: 107 tests, 107 passed, under
  `--blame-hang --blame-hang-timeout 4m --logger "console;verbosity=normal"`. The two runs called
  hangs on 2026-08-18 (aborted at 10 and 15 minutes) were healthy runs killed early.

  **Both pieces of evidence behind that diagnosis were worthless, and it is worth knowing why:**
  - **Silence.** `dotnet test` at default verbosity prints NOTHING for a passing test. A healthy
    27-minute run and a wedged one produce byte-identical output until the summary.
  - **"No COM host process."** The live fixtures use the IN-PROCESS `ComGateway`, so they never spawn
    `OutlookAI.ComHost.exe` at all. Its absence was normal, and I read it as proof of a stall.

  The cost was not theoretical: aborting the first of those runs skipped the artifact sweep and left
  7 tagged items in a real mailbox, and I then wrote a reproducible-hang finding and a maintainer
  question around a defect that did not exist.

  **What was real, and stays:** the 22.5-minute `LiveDisconnectRecoveryTests` stall on 2026-08-17,
  which the long-running-test diagnostics named. The bounds added to it and the preflight gate are
  both worth having on their own merits - the preflight now reports
  `[preflight] Outlook responsive (0 of 5 UI windows hung)` per collection.

  **Rule for the next person, including me: never diagnose a live-tier hang without
  `--logger "console;verbosity=normal"`.** Run it verbose and read what it is doing before touching
  anything.

- [x] **DONE (2026-08-18) - The store-count tripwire could not tell the user apart from the
  tests, and said so loudly.** The 26.8-minute run passed all 107 tests and then FAILED at
  teardown:

  ```
  STORE COUNT TRIPWIRE: the live tier changed mailboxes it may not touch.
    ITEMS LOST: store 'info@voipfabric.com' folder 'Inbox' 168 -> 161 (-7).
    ITEMS LOST: store 'Jan van Linge' folder 'Ongewenste e-mail' 1 -> 0 (-1).
    ITEMS LOST: store 'Jan van Linge' folder 'Postvak IN' 52 -> 50 (-2).
  ```

  **CONFIRMED: the maintainer deleted that mail himself.** The alarm was a false positive and
  no mail was lost by the suite.

  **What was actually wrong, and what could not be fixed.** A before/after reading of a
  mailbox CANNOT name the actor - Outlook records that an item is gone, never who removed it.
  So the guard was not made quieter, and a removal it cannot explain still fails the suite.
  What changed is that it now measures enough to be CHECKED, and enough to stop being blind
  in the direction nobody had looked at:

  - **The census records item identities, not just counts**, for every ordinary folder it can
    afford (`T2/CensusIdentityPlan`: 500 items per folder, 3,000 per store, the hub and the
    self-pruning folders skipped because a shrink there is not evidence of anything). The hub
    keeps its existing treatment - exempt, policed by the zero-artifact sweep instead.
  - **A firing now names WHICH items left**, by EntryID plus a metadata fingerprint (received
    instant and size), and says where each one ended up. `-7` was unfalsifiable; seven ids
    sitting in Deleted Items can be confirmed or refuted in seconds. **No subject and no body
    is ever stored or printed** (S3) - the subject is read in-process only to set a boolean.
  - **Mail that was FILED is no longer reported as lost.** An item that left one folder and
    turned up in another ordinary folder of the same store is still there, and the census can
    prove it. That is the one exoneration the evidence actually supports, and a count could
    never see it. Rule-filing on a busy machine was a false-alarm class of its own.
  - **The blind half is closed: a removal masked by an arrival now fires.** A count that
    starts and ends at 168 while an item is destroyed and another arrives was invisible.
    That was the half "nobody has looked at" and it is the one real strengthening here.
  - **Every failure carries an ATTRIBUTION line.** It says "THE SUITE" only when a departed
    item carried `[OutlookAI-McpTest]` in a mailbox the write allowlist forbids - the one
    thing the evidence supports outright - and otherwise says plainly that the actor is
    undecidable and shows its working.
  - **Junk mail (folder id 23) joined the self-pruning set.** Junk expires on a server policy
    nobody here controls; `Ongewenste e-mail` going 1 -> 0 during a run is expiry, not loss.
    That is one of the three lines above gone on its own.
  - Fail-closed is unchanged: no baseline still means the live tier REFUSES to run, and a
    folder whose walk fails degrades to a count rather than to nothing.

  Pinned by 18 new T1 tests (`StoreCountTripwireTests`, `CensusIdentityPlanTests`).

  **STILL OPEN - one decision for the maintainer, and one unknown that needs a live run:**

  1. **The 2026-08-18 reading would still fail today.** Deleting your own mail during a run
     and a runaway test deleting it produce the same census. The remaining options are all
     policy rather than measurement: leave it strict (today), add an explicit
     "the mailbox was in use" declaration to the live-test settings that downgrades
     non-attributable removals to a loud warning, or require an idle machine. None was
     implemented, because a switch that can silence this guard is the maintainer's call.
  2. **What the identity census costs on this profile has never been measured.** The
     estimate is seconds, not minutes, because the budget bounds it at 3,000 items per store
     per census and the walk is late-bound COM. `[tripwire] baseline: ... identified N
     folder(s)/M item(s), T ms` now prints the real number on every run - read it on the
     first live run and raise or lower `CensusIdentityPlan.DefaultPerStoreItemBudget`
     accordingly.

- [ ] **PENDING TASK - process `C:\Source\SixFive7\BrowserAI\.work	runcation-prompt-for-sibling-project.md`.**
  The maintainer asked for this at 09:00 on 2026-08-18. It is expected to be the portable
  description-budget prompt written for another project; read it and act on what it asks for. Recorded
  here because auto-compaction was imminent when it was requested.

- [x] **DONE (`eee02f2`) - `TryCreateDerivedDraft`'s cross-store retry is unguarded.** It re-attempted
  draft creation across every store on `r == null` alone, while its sibling loop in `TryUpdateDraft`
  only retried on `error == "ItemNotFound"`. Tracing the token before changing it turned up the
  opposite bug in that sibling: `TryUpdateDraft` and `TryDiscardDraft` never SET `"ItemNotFound"`, so
  their cross-store retries were dead code and a draft in a non-default store answered with an opaque
  COM code - `BuildDraftRefusal` even carries a written-out `case "ItemNotFound"` that could never
  fire. All three draft paths now set the token at their `GetItemFromID` and nowhere else, which
  tightens the first and revives the other two, and the rule is single-sourced as pure
  `MailService.ShouldSearchOtherStores` / `KeepSearchingStores` because the two loops disagreed only
  because it was written twice. ~~**Still open, deliberately:** four other loops (`TryReadItem`,
  `TryDisplayItem`, `TrySaveAttachment`, `TryGetSendableDraftState`, `TryGetMailInfo`) still retry on
  `r == null` alone. Three are read-only; `TryDisplayItem` opens a window and `TrySaveAttachment`
  writes a file, and both are classified MUTATING. Their COM layers do not set the token either, so
  tightening them without that groundwork would break them exactly as `TryUpdateDraft` was broken.~~

  **CLOSED 2026-08-18 - the remaining five are done, groundwork first.** Their COM layers now set
  `ItemNotFound` at their own `GetItemFromID` and nowhere else, so the token exists before anything
  reads it: `TryReadItem`, `TrySaveAttachment`, `TryDisplayItem`, `TryGetMailInfo` and
  `TryGetSendableDraftState`. All ten by-EntryID operations in `OutlookComSession` now set it at the
  same place, and the word itself is a shared constant (`ComErrorTokens.ItemNotFound`, aliased by
  `MailService.ItemNotFoundToken`) rather than two literals a compiler cannot relate - which is how
  the original pair came apart. The five service-layer loops then went over to
  `ShouldSearchOtherStores` / `KeepSearchingStores`, so nine of nine now decide the same way.
  The inner catch in each COM method is deliberately the SAME filter that method already captured,
  so this renames a failure and never changes which failures escape. Pinned by
  `T1/CrossStoreRetryScopeTests.cs`, which drives all five through the real service against a
  counting stand-in session: a non-open failure must stop at the first store, an absent reason must
  stop too (fail closed), and a genuine not-found must still reach every store - the last of those
  being the half that breaks silently.

- [x] **DONE (`eee02f2`) - One door back to the cross-store attribution defect.**
  `MailService.ApplySweepCounters` fell back to whole-sweep totals when a store was named but
  `result.PerStore` was empty. Guarded rather than commented: `store != null` alone now decides, a
  missing entry answers zero, and zeroes make `DescribeCoverageGaps` raise `nothing_swept` with
  `degraded: true`. That is the loud, safe direction - "no coverage attributable to this store"
  instead of lending it another account's - and a future third construction site that forgets
  `PerStore` fails visibly instead of silently reopening `c515565`. Pinned in T1.

- [x] **DONE (`1460ddc`, audited 2026-08-18) - A useful error message never reaches the caller.**
  An `exhaustive` search naming a folder that does not exist came back as an opaque
  `ComHostRemoteException: "Exception has been thrown by the target of an invocation."`
  instead of the message the code takes trouble to write:
  `"Folder 'X' was not found in store 'Y' (...). Use list_folders for paths."`
  (`OutlookComSession.cs`, the exhaustive scan's folder resolution). Reproduced 2026-08-18 on
  a primary store and on a delegate store, before and after the DASL date fix, so it is
  independent of that work.

  It is a reflection-invocation wrapper swallowing the inner exception's message somewhere on
  the COM-host pipe, which means the diagnosis is "where does the wrapper lose it", not "write
  a better message" - the good message already exists. Worth checking whether OTHER deliberate,
  actionable errors from inside the COM host reach callers intact, because this one does not
  and nothing noticed until a probe happened to hit it. An agent that cannot tell "no such
  folder" from "something went wrong" retries blindly instead of calling `list_folders`.

  **Where it was lost:** two reflective hops in series - `GatewayRoutingProxy.Invoke`'s
  `targetMethod.Invoke` and `ComHostServer.Invoke`'s `method.Invoke` - wrapped the session's
  exception twice, and the server peeled exactly one layer, so the INNER wrapper is what went
  on the wire: type `TargetInvocationException`, that one sentence as its message, and null
  for both HResult and Reason. `1460ddc` fixed it in two independent places (the routing proxy
  unwraps its own hop with `ExceptionDispatchInfo`; the server's peel is now a loop), and this
  TODO entry was simply never marked.

  **The audit the entry asked for is now done**, over all 26 contract methods and both
  channels, in `T1/ComHostErrorFidelityTests.cs` - the whole child side runs for real over an
  in-process pipe, with the parent's own mapper rebuilding what comes back. Result: every
  operation delivers its own message intact; every one of the 14 exception types the mapper
  models arrives as itself; COM HRESULTs and refusal reasons survive; and the `out string?
  error` channel (which the exception path says nothing about) delivers its token for all 17
  operations that use one. Proved by reinstating the defect - with both unwraps removed, 19 of
  the 20 audit cases fail; the survivor is the token channel, which never crossed the
  exception path. **Two residual losses were found and fixed:** an unmodelled child-side type
  reached the agent labelled `ComHostRemoteException` (the name of the pipe, not of the
  failure) because `RemoteType` was carried on a property nothing read - the tool layer now
  reports it; and the mapper's own comment claimed that name went into the message, which it
  never did. Invariant 10 in `check-pinned-constants.ps1` now fails the build when the COM
  layer raises a type the mapper does not model.

- [ ] **Run the index-collation probe on the live profile** - `T2 LiveOrderKeyCollationTests`
  (read-only, index statements only, no COM and no mailbox writes). It answers two things the
  B3 follow-up could only reason about, both recorded in `QUESTIONS.md` under Q8 and in
  `Docs/magic-numbers.md` beside `WsSqlBuilder.OrderKeyFloorUtc`:
  - [ ] **Where the provider sorts a NULL under `ORDER BY System.Message.DateReceived DESC`.**
        If last, the displacement refetch never fires and the guard is free; if first, it fires
        on every truncated search and each one costs a second index statement. The guarantee
        holds either way - this decides only what it costs, and it is the number that belongs
        in the magic-numbers row, which currently says "not measured".
  - [ ] **Whether the provider accepts the `1601-01-01 00:00:00` floor literal** and treats the
        comparison as "has a value". If it does not, the refetch fails and searches that need it
        return a short answer flagged with `index.candidatesExhausted` - loud, but the guarantee
        then rests on a query that never runs.

  **PARTIAL ANSWER, measured 2026-08-18 on this machine, directly against `Search.CollatorDSO`
  (three read-only SELECTs, no Outlook, no mailbox).** Under `ORDER BY System.Message.DateReceived
  DESC` over a predicate matching the whole index, the first 25 rows were **all dated** - and on a
  developer machine files vastly outnumber mail, so had undated rows sorted FIRST the block would
  have been entirely undated. Under `ASC` the first 25 were the oldest mail rather than undated
  rows, so they are not sorting lowest either. **The `1601-01-01 00:00:00` floor literal was
  accepted and returned rows.** So the displacement refetch should essentially never fire, and the
  guard is free in practice. Two readings fit the data and it cannot separate them: the provider
  may exclude rows lacking the ORDER BY property from an ordered result, or place them last in both
  directions - both give the same answer here, but they are different facts. **This does NOT close
  the item:** the statements carried no `SCOPE='mapi...'`, so they ran over the general SystemIndex
  namespace rather than the one the product uses. Full write-up and the exact statements are in the
  session trace folder under Downloads (`tmp-aitrace/nullorder-finding.md`).

- [x] **DONE 2026-08-18 15:12 against `2d28957` - Re-run the ten store-scope probes on the unindexed-PST machine.** The A4 fix (a `store`
  scope resolved against the profile Outlook has rather than against the index) is pinned in T1
  against stand-in index clients whose store catalog is empty, and against one that omits a store
  the profile has - both fixtures the suite had never had, which is why the defect shipped. What
  T1 cannot exercise is the real chain that produced the empty catalog in the first place: a live
  `DiscoverStoreScopes` over a SystemIndex that holds nothing for the profile, and the COM store
  list underneath it. Re-run the same ten probes that found this (the Hyper-V VM: one PST, not
  indexed, `WSearch` running, Outlook connected) and confirm probe 14 now answers with
  `index.storeNotIndexed: true` plus `no_index_frontier`, probe 15 still refuses and names the
  real store, and probes 11/12/13/17/18/19 are unchanged. Read-only: `search`, `list_folders` and
  `list_accounts` only, no mailbox writes.

  **Result: verified.** Probe 14 answers `degraded:true` `freshness:"partial"`
  `coverageGaps:["no_index_frontier"]` `sweep.storesWithoutIndex:["Outlook Data File"]`
  `index.storeNotIndexed:true` `scope.shape:"store_not_indexed"`. Probe 15 still refuses, with the
  new message naming the real store: *"Store 'no-such-store-xyz' was not found in Outlook. Known
  stores: Outlook Data File."* Probes 10-13 and 17-19 byte-identical to the pre-fix run. **One
  criterion the machine cannot test:** "does not widen" is unobservable on a single-store profile,
  since a widened search and a correctly scoped one return the same set; T1 covers it with an index
  stand-in that answers with a foreign store's mail. Evidence, and the pass criteria written before
  the fix landed, are in the VM working folder under Downloads (`tmp-outlookai-vm`), with the
  before/after transcripts beside them.

- [ ] **An answer too big to frame: refused (done), but not yet prevented (open).** Found by the
  boundary audit on 2026-08-18. **The refusal and the measurement landed 2026-08-18; the four
  responses below are still the maintainer's to choose between, and nothing here pre-empts them.**

  **Done - (a) the refusal.** `ComHostServer.ServeAsync` now catches the framing refusal around
  its write and answers with a `ComHostResponseTooLargeException` frame naming the operation, the
  size and the limit; the serve loop keeps serving, so one oversized answer costs one call instead
  of the session. `ComHostErrorMapper` rebuilds the type, and `OutlookTools.GuardAsync` returns it
  as `ResponseTooLarge` with advice to NARROW the request rather than retry it. The testing
  objection below is answered by the fourth option it did not list: `ComHostServer` takes its
  ceiling as a constructor parameter defaulting to `ComHostProtocol.MaxFrameBytes`, so T1 reaches
  the branch with a 64 KB answer against a 4 KB ceiling - real serve loop, real framing, real
  mapper, no 64 MB allocation. Measured 2026-08-18: the two refusal tests run in 73 ms and all six
  new tests in 99 ms, against a T1 tier of 1 m 55 s.

  **Done - the measurement.** `ComHostFrameMeter` keeps a per-process high-water mark of the
  largest payload seen (both directions: encode and read, because the big frames are built in the
  child whose counters die with it) plus a refusal count. `outlook_health` reports both beside the
  limit, as `comHost.largestFrameBytes`, `frameLimitBytes` and `framesRefusedTooLarge`, and a
  refusal also raises a `problems` line. Lifetime is one SERVER process and survives child
  restarts, deliberately: the child is restartable and the question is about the product. So the
  number that says whether 64 MB is right is now collectable from any running install - it was
  previously unmeasured, which is why the entry below argues from the caps instead.

  **Still open - what to do about the size itself.** The refusal turns a dead host into a legible
  failure; it does not stop the answer being too big. Four responses, unchanged and still
  unanswered: **(b)** raise or lower `MaxFrameBytes`; **(c)** cap bodies at the COM layer so an
  unsendable frame cannot be built; **(d)** chunk the sweep result across frames; or accept the
  refusal as the whole answer and let callers narrow. The high-water mark is the evidence any of
  those choices should rest on.

  **The limit is reachable by ordinary use - derived from the caps 2026-08-18, not measured.** One
  `SweepFoldersNewerThan` answer is a single frame and `MailService` calls it with
  `includeBodies: true`, so a frame carries 4 arrival-path folders x `SweepPerFolderCap` (200)
  items per store, times every store in the profile. **The bodies are not capped at the COM
  layer** - `SnapshotBrief` takes `item.Body` whole, and `BodyCharsDefault`/`BodyCharsCap` are
  applied in `MailService`, on the FAR side of the frame: they bound what the agent sees, not what
  crosses the pipe. That puts 64 MB at ~80 KB average body on a one-store profile, ~27 KB on three
  stores, ~16 KB on five. An 80 KB body is an ordinary long quoted thread. **And the path there is
  the unindexed-store case**: the sweep window is normally minutes wide, so 200-per-folder never
  fills, EXCEPT when a store is missing from the index and the window falls back to seven days. So
  the more degraded the index, the larger the frame - and before the refusal landed, the frame
  bursting killed the subsystem that was compensating for the degraded index. It now refuses that
  one call instead, which is why (c) and (d) are still worth choosing between: the sweep on a
  degraded profile is exactly the caller with nothing left to narrow. This is where options (c)
  and (d) come from. Write-up in the session trace folder under Downloads
  (`tmp-aitrace/frame-size-analysis.md`).

  **The defect, as found (kept for the record; the refusal half is fixed).**
  `ComHostProtocol.EncodeFrame` refuses a payload over `MaxFrameBytes` (64 MB) by throwing
  `ComHostProtocolException` - a deliberate, specific, actionable failure. But it was thrown from
  `ComHostServer.WriteAsync`, which `ServeAsync` guarded with `catch (IOException)` only. So the
  exception left the serve loop, `Program.Main` printed it to stderr and the child exited with 1.
  The caller learned "the COM host went away", which is the one fact that says nothing about what
  to do next; the sentence naming the size and the limit reached only the child's stderr. This is
  the same species as the wrapper defect above - a good message that does not survive the process
  boundary - and it is the only other instance the audit found.

  **How likely is it?** Low - `Docs/com-host.md` calls 64 MB "far above any real payload" and a
  `read` returns ~0.5 MB. The candidates are `SweepFoldersNewerThan(includeBodies: true)` and
  `ExhaustiveScan` over a large window. That "low" is still a derivation rather than an
  observation; the high-water mark now accumulating in `outlook_health` is what will replace it,
  and until an install has been read it says nothing on its own.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
