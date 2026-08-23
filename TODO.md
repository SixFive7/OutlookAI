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

- [ ] **Residual questions from the 2026-08-19/20 atomicity-claims sweep.** All 31 claims of
  non-effect were enumerated, 16 were wrong and all 16 are fixed (`Docs/completeness-gaps.md`
  section 7b, T1 `AtomicityClaimsTests`). These are the things reading could not settle, and the
  two wrinkles beside the send path that are NOT the audited defect.

  - [ ] **Which HRESULTs Outlook actually raises from `MailItem.Delete()`, `Move()` and
        `Display()` when the RPC channel breaks mid-call.** The may-or-may-not reading of
        `RPC_S_CALL_FAILED` is documented and is this repo's own stated basis for the retry
        classification, but whether these three calls produce it in practice is unmeasured. It
        decides whether the unknown-outcome wording on those three paths describes a rare event or
        a theoretical one. **The wording was deliberately fixed WITHOUT waiting for it**, on the
        maintainer's reasoning: a claim about what did not happen should not rest on an unmeasured
        probability. Measurable on the VM with `OUTLOOKAI_COMHOST_FAULT` plus a read afterwards.
  - [ ] **Whether a MUTATING response can actually exceed the 64 MB frame.** Row J8's verdict is
        "false by construction": `ComDraftUpdateResult` and `ComDraftCreateResult` carry no body,
        only recipients and attachment metadata, and no realistic 64 MB case could be constructed
        by reading - a recipient count in the hundreds of thousands would be needed. The serialised
        size of a `ComDraftUpdateResult` is measured nowhere. The fix is one `if` either way, so
        this is a curiosity rather than a blocker.
  - [ ] **Whether `TryDiscardDraft` can reach its catch-all through a NON-`IsComCallFailure`
        exception after `Delete()`.** A `NullReferenceException` or `InvalidOperationException`
        from `collection.Count` or the indexer inside `TryFindDiscardedCopy` would escape
        `_runner.Run` entirely rather than becoming `com_failure`, and what `PumpedStaRunner` does
        with such an escape was not established. It changes which message the caller gets, not
        whether the delete happened.
  - [ ] **Whether soft-deleted drafts actually survive on the maintainer's profile.** The discard
        fix leans on recoverability from Deleted Items. Outlook's "empty Deleted Items on exit"
        option and Exchange retention tags can both remove the item, and neither is visible from
        this codebase. The message says to look in Deleted Items, which is right in either case;
        what is unknown is how often looking will find anything.
  - [ ] **`send` catches `TimeoutException` only** - noted by the audit and NOT part of the defect
        it was auditing, so deliberately left alone in that pass. A child that dies for any other
        reason raises `ComHostUnavailableException`, which is not a `TimeoutException`, so that
        path is saved by `ComHostSupervisor.DescribeInterruption` instead. The outcome is right by
        a different route rather than by this `catch`, which means the send path's own
        `send_outcome_unknown` audit line is NOT written for it. Worth widening the filter, or at
        least writing down that the audit line is conditional on which way the child died.
  - [ ] **`SendUsingAccount` is written to the user's draft before the identity readback and is
        never restored.** The "Nothing was sent" claims stay true, and the messages now SAY the pin
        may have been rewritten (rows 9 and 10) - but saying so is a smaller fix than restoring it.
        Restoring needs the previous value captured before the putref and put back on every failure
        path, which is one more mutating call on a path that is currently refusing to mutate
        anything, so it wants a decision rather than a quiet fix.
  - [ ] **Seven decision lines from this pass are provably unguarded, established by mutation
        rather than assumed.** Each was reverted, built, run against the whole non-live suite and
        restored; 30 of 37 were caught, and eight of those thirty only after the gap they exposed
        was closed with a new test. The full table is in `tmp-aitrace/mutation-table.md`.
        - **Five live in `OutlookComSession`, behind a COM call no non-live test can execute**,
          which is the same class as `sortApplied` and the attachment-plan execution already
          recorded above: the saved-draft id being published on the failure path and re-read after
          the relocate (row 5's COM half), `TryMoveItemToPath` reporting created folders on the
          failure path (row 6's COM half), `TrySaveAttachment` reporting the attempted path, and
          the size read being best-effort (row 11's two halves). Every T1 test substitutes the
          session, so a test can only prove that the SERVICE layer uses what it is given. What
          each is worth is not in doubt - the shapes are read off the code - but they are
          unexercised until a live run. The cheap substitute is the one already used elsewhere: a
          temporary build that forces the branch.
        - **`CompleteMove`'s audit-failure branch** (`Ok=false` over an item that MOVED, reported
          as `outcome: applied`) needs `AuditLog.Append` to FAIL, and it writes to
          `%LOCALAPPDATA%\OutlookAI` through a path that is not injectable. `AppendTo` takes a
          directory and is used by `AuditLogTests`; wiring the service layer to it would make this
          reachable, and is a bigger change than the row it guards.
        - **The supervisor's own wiring for the interrupted-request outcome** needs a child that
          dies while holding a request. Both ends are pinned separately - the value
          (`MutationOutcome.ForInterrupted`) and the carrier
          (`ComHostUnavailableException.Outcome`) - so what is unguarded is the line that joins
          them.

  - [ ] **The `outcome` field is not consumed by anything yet.** It is additive and null-omitted,
        the tool descriptions teach it, and T1 asserts a `com_failure` never carries `unchanged` -
        but no agent behaviour depends on it, so its value is currently "the claim is testable"
        rather than "the claim is acted on". If it turns out nothing ever branches on it, the
        honest conclusion is that the prose was the whole fix and the field is cost.

- [ ] **Residual gaps left by the 2026-08-19 timeout pass.** The values, the three inventory
  defects, the graceful sweep expiry and the kill work all landed; these did not, and each one
  is written down because a mutation check proved it unguarded rather than because it was
  guessed at.

  - [ ] **Nothing in the non-live tier notices if `ComGateway`'s budget overload goes back to
        being a pass-through.** Reverting `BudgetedSessionProxy.Wrap(session, budget)` to
        `session` leaves all 1,887 tests green. The MECHANISM is pinned (T1
        `InProcessBudgetTests` drives the proxy directly, and removing its dispatch check
        fails); the WIRING is not, because exercising it needs a real COM session. The live
        tier is where it is exercised, which is the tier that had no budget at all until this
        pass. Options: accept and rely on T2; add `InternalsVisibleTo` to `OutlookAI.Core` and
        pin the wiring through an internal seam; or a structural IL assertion, which is
        fragile and unlike anything else here.
  - [ ] **Same for `ComHostSupervisor.CleanExitGraceMilliseconds`.** Replacing the
        `WaitForExit(250)` with a no-op leaves the suite green. It is a process-lifecycle
        behaviour with no observable payload; proving it needs a child that logs its own
        clean exit, which is a T3-shaped test nobody has written.
  - [ ] **`MailService.SearchIndexTimeoutSeconds` is pinned only from above.** T1 asserts it
        never exceeds `OleDbIndexClient.DefaultCommandTimeoutSeconds`, so reverting 60 to 15
        fails nothing. A lower bound would need a measurement constant for "how long a
        statement on a saturated indexer legitimately takes", and no such measurement exists.
  - [ ] **`T3/McpStdioClient` still has one budget for a whole session.** One
        `CancellationTokenSource` bounds every read and write for the client's lifetime, so it
        is a session budget masquerading as a per-call one. The exhaustive-scan live test now
        passes an explicitly derived budget, which is the case that would have broken first;
        the general split (session budget plus a per-`RoundTripAsync` budget, both named)
        is still open. Raising the DEFAULT is deliberately not the fix - it is CI's only
        safety net against a hung stdio test, and CI's job timeout is 20 minutes.
  - [ ] **`T2/LiveAttachmentKindRecallTests` (~line 364) keeps two bare literals** (90 s wait,
        5 s poll) and is the last product-shaped index call in the suite relying on the client's
        default command timeout. Name them and pass a `commandTimeoutSeconds`; one slow
        statement currently eats a third of the wait.
  - [ ] **Claude Code's 30-minute stdio idle abort is now the nearest client-side limit, and
        nobody owns it.** A 600 s exhaustive scan is 600 s of complete silence on the pipe -
        this server sends no progress notifications. It fits (600 s < 1800 s idle < the
        ~27.8 h per-call hard ceiling), but the idle limit is a client default nobody here
        chose, no test watches it, and a user who sets a per-server `timeout` in `.mcp.json`
        for a sensible reason will land far below 600 s. Re-measure instruction: `grep -a` the
        shipped `~/.local/share/claude/versions/<ver>` binary for
        `CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT` and resolve the stdio branch's constant.
  - [ ] **The graceful sweep expiry has never fired against real mail.** It is pinned at both
        ends in T1 - the pure boundary (`OutlookComSession.SweepBudgetSpent`) and the whole
        reporting chain from `ComSweepResult` to the advice sentence - but the walk that stops
        exists only over live COM folders. With the budget now at 165 s inner / 180 s outer and
        a measured whole-profile sweep of ~60 s, it should never fire in ordinary use, which is
        the point and also the reason it will stay unexercised.

- [ ] **Residual gaps left by the 2026-08-19 re-entrant `update_draft` pass.** The intent record,
  the idempotence key, the attachment reconciler and the add-before-remove reorder all landed and
  are pinned in T1 `DraftUpdateReentrancyTests`. These did not, and each is here because a mutation
  check proved it unguarded rather than because it was guessed at.

  - [ ] **The attachment enumeration guard is provably unguarded** (mutation M15, 2026-08-19):
        making `TryUpdateDraft` enumerate the draft's attachments even when the request touches
        none leaves all 1,921 tests green. It is a cost guard - one COM call per attachment on
        every body-only revision - with no observable payload, so nothing outside a live profile
        can see it. Options: accept and rely on T2; or count contract-level COM calls in a live
        fixture, which nothing does today.
  - [ ] **The COM-side EXECUTION of the plan is unguarded by any non-live test.** `DraftAttachmentPlan`
        is pinned state by state, and the two decisions the COM sequence makes were lifted out so they
        could be reverted (`BuildForAttempt`, `ComDraftUpdateResume.ThreadIndexFor`). What no T1 can
        reach is `RemoveAttachments` deleting the N lowest-indexed instances of a name, and the fact
        that additions now run before removals - both are `dynamic` COM against a live
        `Attachments` collection. Reversing the order back leaves the suite green. T2
        `LiveUpdateDiscardTests` is where it would be exercised; extending it to assert the ORDER
        needs a way to observe a mid-sequence state, which nothing has today.
  - [x] **DONE (2026-08-20) - `discard_draft`'s `com_failure` no longer claims "Nothing was
        changed".** Fixed as part of the full atomicity sweep below. **The reason recorded here was
        wrong and is corrected rather than inherited:** this entry said the post-delete re-locate
        reaches the catch-all. It does not - `TryFindDiscardedCopy` has its own
        `catch when (IsComCallFailure(ex))` covering all six COM-failure types the session
        recognises, so that route is closed, and between `Delete()` and the return there is nothing
        but a null check and a managed constructor. The route that IS open is `Delete()` itself,
        sitting bare in `TryDiscardDraft`'s outer `try`, whose disconnect family includes
        `RPC_S_CALL_FAILED` - which this repository already documents as MAY OR MAY NOT have
        executed on the server. One mutating call with that semantic is enough, and it is a stronger
        argument than the one it replaces because it does not depend on a second failure after the
        first.
  - [ ] **Nothing verifies that an interrupted attempt's partial writes are DURABLE.** The whole
        design is deliberately indifferent to it - it converges on the end state from whatever the
        draft is observed to hold - but the question is still unanswered and worth answering, because
        it decides how often the resume does anything at all. Whether Outlook keeps an unsaved
        `Attachments.Add` / `Attachment.Delete` when the automation client dies is undocumented by
        Microsoft (`tmp-aitrace/kill-safety.md` section 2.3). Measurable on the dev machine with
        `OUTLOOKAI_COMHOST_FAULT=hang:TryUpdateDraft` plus a read of the draft afterwards.
  - [ ] **The gap map's remaining line-number citations are stale.** Checked 2026-08-19: all nine
        `OutlookComSession.cs` references had drifted (converted to symbols in this pass), and the
        `MailService.cs`, `MailModels.cs` and `OutlookTools.cs` ones are stale too - `MailService.cs:220`
        lands in an unrelated comment, `:612` and `:3995` on bare `</summary>` lines. Convert each as
        its row is next touched; a stale number reads as evidence and points the next reader at
        unrelated code.

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

- [x] **DONE (2026-08-20) - The count tripwire's census could not finish on a real Exchange
  profile, so the live tier refused to run at all.** The first run since the census began
  capturing identities ended before a single test executed:

  ```
  REFUSING to run the live tier: the baseline per-store census for 'info@voipfabric.com'
  could not be taken (TimeoutException: Test mailer STA operation timed out.).
  ```

  **The refusal itself is correct and is unchanged** - an unmeasured mailbox cannot be proven
  untouched, and fail-closed is the whole point of the guard. What was wrong is that the
  census was too expensive to complete.

  **The cost driver, established by counting the calls the census makes rather than by
  guessing.** Per store, per census pass, the old census cost roughly: 15 calls of setup,
  12 for the six `GetDefaultFolder` volatile probes, about 8 per folder for the tree walk
  (`Name`, `EntryID`, `DefaultItemType`, `Folders`, `Folders.Count`, the child fetch, `Items`,
  `Items.Count`) - and then **five cross-process calls for every item walked**: `Items[i]`,
  which OPENS the message, plus `EntryID`, `ReceivedTime`, `Size` and `Subject`. At the
  3,000-item per-store budget that last term alone is **15,000 round trips, about 94% of the
  whole census**, and everything else is two orders of magnitude smaller. It needs only 12 ms
  per call - or ~60 ms per item opened - to exceed the 3-minute STA budget, and a shared or
  delegate Exchange mailbox that is not cached locally is squarely in that range. The
  prediction of "well under a second" was made about two local PSTs, where it was right.

  **The fix: the walk is a bulk table read.** `Folder.GetTable` with `ReceivedTime` and `Size`
  added as columns (`EntryID` and `Subject` are already default columns), read
  `CensusTableRowBatch` = 200 rows at a time through `Table.GetArray`, which is one round trip
  per 200 items instead of five per item. The same 3,000 items now cost about **fifteen calls
  instead of 15,000**, and the folder tree walk becomes the dominant term. Nothing about the
  budget moved: it now bounds bytes and memory rather than round trips.

  **What it still proves - all of it.** Same counts (still `Items.Count`), same per-item
  EntryID identity, same move-stable fingerprint (received instant plus size, no subject and
  no body), same tag flag, same fail-closed refusal, same "a folder whose walk fails degrades
  to a count, never to nothing". Three cross-checks now have to agree before a walk is
  accepted: the table hands back exactly as many rows as the count promised, no EntryID
  repeats, and the count is unchanged afterwards.

  **What changed, precisely, and it is two things.**
  1. A folder whose table will not carry all four columns is recorded as a COUNT, where the
     old code would have produced identity without a fingerprint. That is deliberate:
     identity without a fingerprint cannot prove a filing, so it would turn a person filing
     mail during a run into a suite failure, and identity without a subject would say
     "undecidable" over a removal the suite itself caused. The number of folders that fell
     back is now printed, so this cannot happen quietly.
  2. The fingerprint's received instant comes from the table's date column, and an
     unspecified `DateTimeKind` is read as UTC (what Microsoft documents for the `Table`
     object, and the contract `DaslDateLiteral` already states for this solution). If that
     reading were wrong the tripwire's DECISIONS would not change - every value at both ends
     of a comparison comes through the same method - and only the instant printed beside a
     departed item would be offset by the machine's UTC offset.

  **A silent timeout is now a diagnosis.** Each store's census prints its own folder count,
  what identifying it cost and its elapsed ms, and a refusal names how far the census had got
  when it stopped (`CensusIdentityPlan` doubles as the progress record). The 2026-08-20
  failure could not distinguish a slow folder tree from a slow item walk; the next one will.

  Pinned by 16 new T1 tests (`CensusTableRowTests`, plus two in `CensusIdentityPlanTests`).
  The projection, the column map and the bulk-read shape check were deliberately written as
  PURE functions in `T2\CensusTableRow` so CI can reach them - the COM walk around them still
  cannot be executed by any non-live test.

  **Mutation-checked: 26 decision lines, 15 caught, 11 not** (table:
  `tmp-aitrace/mutation-table.md`). The split is clean and structural - every line a non-live
  test can reach was caught, and **all 11 that were not are inside `CaptureMailFolderCensus`,
  which no non-live test can execute a line of**. They are, with what each would need:
  `plan.NoteFolderMeasured()` and `plan.NoteDegradedToCount()` at their call sites; the negative
  and empty `expectedCount` guards; the `!columns.IsUsable` degradation; the per-batch
  `Math.Min` that asks only for the rows still owed; the `EndOfTable` check after the walk; the
  `ConfirmUnchanged` re-read; `AddCensusColumn` trying both spellings; and the two log lines in
  `LiveStoreCountTripwire.Capture`. Each needs the same thing: a seam that lets a fake stand in
  for `Folder`, `Items` and `Table`, which is the cheap substitute already proposed for the
  census generally elsewhere in this file. Until then they are covered by a live run and by
  nothing else.

  **STILL OPEN - three questions for the maintainer:**

  1. **Should the per-folder identity limit rise now that identity is nearly free?** At 500 a
     4,918-item Sent Items and a 108,144-item Archive are counted, not identified, so a
     deletion there is detected but unnamed, and a filing OUT of them cannot be proven. The
     old objection was cost and it has largely gone: 5,000 items is 25 `GetArray` calls.
     Raising it changes what the guard proves, so it is not being changed as a side effect of
     making the census affordable. Options: leave at 500; raise to ~5,000 per folder with the
     per-store budget raised to match; or raise only for non-delegate stores.
  2. **Should the 3-minute STA timeout move?** It is already per store (one `RunSta` per
     store), so a "per store rather than per operation" change buys nothing. It was left
     alone: the fix removes the term that could plausibly exceed it, and raising a timeout
     without evidence moves a silent failure later. If the next run fails again it will now
     say where, which is the cheaper way to buy the evidence.
  3. **`AnUnspecifiedKindIsReadAsUtc_SoTwoCensusesAgree` cannot fail on a UTC machine.** It
     catches the mutation here (nl-NL, UTC+2 in August) and would not catch it on a CI runner
     set to UTC, because the two readings coincide there. No way was found to pin it
     machine-independently; it is recorded rather than papered over.

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
  2. ~~**What the identity census costs on this profile has never been measured.** The
     estimate is seconds, not minutes, because the budget bounds it at 3,000 items per store
     per census and the walk is late-bound COM.~~ **ANSWERED THE HARD WAY, 2026-08-20: that
     estimate was wrong and the census timed out, refusing the whole tier.** The reasoning
     held for a local PST and not for an Exchange mailbox, where five late-bound calls per
     item are five server round trips. See the 2026-08-20 entry above; the walk is a bulk
     table read now, and the budget stayed where it was.

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

- [ ] **Re-run the unindexed-store probes on a MIXED profile - the one shape no machine here has.**
  Group A and E of `Docs/completeness-gaps.md` are now all closed (A1-A5, E1). Everything about
  them has been verified on two profile shapes: the fully-indexed developer profile, and the
  Hyper-V VM whose only store is an unindexed PST. **The shape that carries A1's residue is
  neither of those** - one INDEXED mailbox plus one UNINDEXED data file, so that the profile-wide
  frontier probe succeeds while one store is still absent from the index catalog. That is the
  ordinary "Exchange account plus archive.pst" desktop, it is what T1
  `UnindexedStoreReportingTests` models with a stand-in index client, and no machine to hand can
  produce it live. What T1 cannot exercise is the real chain: a live `DiscoverStoreScopes` whose
  sample returns one store's prefixes and not the other's, and `StoreHasIndexRows` answering
  false against a real SystemIndex for a mounted PST.

  Mount a PST on a profile that already has an indexed account (or add an account to the VM), do
  not add it to Indexing Options, then confirm READ-ONLY, with `search` only:
  - a plain unscoped search names it in `sweep.storesWithoutIndex` with `no_index_frontier`
    (this already worked - it is the regression guard);
  - an unscoped search with `before` older than 7 days reports `sweep.notNeeded:true` **and**
    `indexFrontierMissing:true`, `freshness:"partial"`, `degraded:true` - it used to say
    `freshness:"live"` with nothing else;
  - an attachment-only search reports `freshness:"index-only"` **and** `no_index_frontier`;
  - `outlook_health` shows the PST as `index.perStore[].inLocalIndex:false` with a problem line,
    while the indexed account still reports a frontier.

  Also worth measuring on that profile: the cost of the one COM store-list read the three
  no-sweep paths now pay (5-minute cached, so expected to be a pipe round trip), and whether
  `StoreIndexProbeBudgetMs` (1.5 s) is ever the thing that cuts the probe short.

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

  **Done - (c), chosen by the maintainer 2026-08-18: cap bodies at the COM layer, so a frame that
  cannot be sent cannot be built.** (b) and (d) were the alternatives and were NOT chosen -
  `MaxFrameBytes` is unchanged and the sweep result is still one frame. Two bounds, both in
  `OutlookComSession`, both applied in `SnapshotBrief` where the body is read:
  `SweepBodyCharsCap` (500 000 chars, = `MailService.BodyCharsCap`, the largest body window `read`
  will ever return in one call) per ITEM, and `SweepBodyBytesBudget` (32 MiB, = `MaxFrameBytes / 2`)
  across the whole sweep. The per-item cap alone provably cannot do the job: the item count is
  4 folders x `SweepPerFolderCap` x every store, and the store count is unbounded, so 800 items at
  500 000 chars is already ~400 MB of theoretical worst case on ONE store. What it does instead is
  keep the budget FAIR - one enormous mail cannot spend it all and blind the sweep to everything
  behind it. The budget is counted in encoded BYTES rather than characters because the pipe escapes
  every non-ASCII character to six bytes: a character budget would have to be sized for that case
  and would then bite at ~5.5 M characters, which an ordinary unindexed-PST sweep really reaches.

  **The cut is never silent, and it is not a display truncation.** These bodies are matched against
  (`FreshMerge.MatchesTerms`) and shown to nobody, so unlike `read`'s windowing - which pages and
  loses nothing - a cut here can make a search MISS a real match. `ComMailBrief.BodyTruncated`
  carries the fact per item; `sweep.itemsBodyCapped` and `sweep.itemsBodyCappedUnmatched` reach the
  payload; the `body_cap` coverage code is raised on the INTERSECTION alone (cut AND unmatched),
  since a cut body on an item that matched anyway cost nothing; `sweep.bodyBudgetExhausted` says
  which bound cut, because the remedies differ. The two facts CAN be told apart and are; what
  cannot be settled is whether the term really sat past the cut, since that needs the text the
  bound refused to carry, so the sentence says "may be" and never "is". T1 `SweepBodyCapTests`
  (18 tests) plus two wire round-trips in `ComHostProtocolTests`, mutation-checked with 13 separate
  reverts, each disabling one decision.

  **What remains reachable, stated rather than implied.** (i) The NON-body half of a frame is still
  unbounded in the store count: per-item EntryIDs, StoreIDs, subjects and folder names are ~1-2 KB
  each, so ~40 mounted stores each holding 200 items in all four arrival-path folders inside the
  window would build an oversized frame carrying no body text at all. Never observed; the typed
  refusal is the backstop. (ii) The byte budget is enforced against an OVER-estimate of the encoded
  size (`EncodedBodyByteCeiling`), so it errs toward cutting early, never toward a frame that will
  not send. (iii) The bound has only ever been exercised in T1 - see the live-profile item below.
  (iv) `read` is DELIBERATELY not capped at the COM layer: `TryReadItem` still returns the whole
  body, plus the whole `HTMLBody` when `include_html` is set, so one pathological mail could in
  principle build an oversized `read` frame. Capping it there would break the one contract that
  makes `read` lossless - `bodyTotalChars` is measured from the full body and `body_offset` pages
  the whole of it, so a COM-side cut would make the total a lie and the tail unreachable. The
  measured `read` payload is ~0.5 MB, 0.8% of the limit. The sweep was the case worth closing
  because its size is driven by MAIL VOLUME rather than by one mail, and because its bodies are
  never shown, so a cut there costs matching rather than reading.

  **The other two frames that carry mail were checked and need nothing.** `ExhaustiveScan` and the
  `thread` walk both snapshot briefs with `includeBody: false` - the exhaustive tier matches
  server-side through DASL and the thread walk needs no body at all - so the sweep is the only
  frame in this server that ever carried body text in bulk.

  **MEASURED 2026-08-18 on the real 5-store profile, with that high-water mark** (read-only:
  `outlook_health`, `list_accounts`, four searches; nothing created, moved or deleted). Largest
  frame **441,930 bytes - 432 KB, 0.66% of the limit, about 152x headroom**; zero refusals.
  **This corrects the derived worst case below:** "reachable by ordinary use" is too strong. The
  filter-only search, the one that should have swept hardest, **timed out** - `Outlook did not
  respond to 'SweepFoldersNewerThan' within 30000 ms` - the supervisor replaced the COM host, and
  the search degraded to `index-only` and still answered with 100 hits. So **the 30-second sweep
  budget bites long before the 200-items-per-folder cap does**, and the cap arithmetic is the wrong
  worst case for an Exchange store. **SUPERSEDED IN PART, 2026-08-19: the 30 s budget is now 180 s,
  and with the bound lifted the frame measurement changes completely.** On a purpose-built corpus
  (one unindexed PST, 20,000 items across the four arrival-path folders, the 200-per-folder cap
  engaged) a single store's sweep produced a frame high-water of **10,734,599 bytes - 10.2 MB over
  758 items, ~13.5 KB per item**, and five such stores extrapolates to **~54 MB against the 64 MB
  limit**. So the 432 KB measured on the real profile was bounded by the TIMEOUT, not by the item
  caps, and `SweepBodyBytesBudget` (32 MiB) is load-bearing rather than insurance: it bites before
  the frame limit does, which is exactly its design intent. The residual PST case below is no longer
  the one this machine cannot produce - it has now been produced, on the test VM, and the bounds
  held. The residual risk narrows, and stays real: a fast LOCAL store
  absent from the index, where the window falls back to seven days, holding a lot of recent large
  mail - the archive/PST shape, and the one case this machine cannot produce, because the only
  unindexed store to hand is the test VM's and it is empty. **Bearing on the options:** (b) was not
  urgent on this evidence, and (c) was the one that would close the residual case outright - which
  is why (c) is what the maintainer chose (see above). The residual PST case is still the one this
  machine cannot produce, so the new bounds have never been exercised against real mail.
  Incidentally, the timeout path was observed working on a real profile for the first time: no
  hang, host replaced, honest degraded answer naming the reason.

  **The limit is reachable by ordinary use - derived from the caps 2026-08-18, not measured.** One
  `SweepFoldersNewerThan` answer is a single frame and `MailService` calls it with
  `includeBodies: true`, so a frame carries 4 arrival-path folders x `SweepPerFolderCap` (200)
  items per store, times every store in the profile. **The bodies were not capped at the COM
  layer** - `SnapshotBrief` took `item.Body` whole, and `BodyCharsDefault`/`BodyCharsCap` are
  applied in `MailService`, on the FAR side of the frame: they bound what the agent sees, not what
  crosses the pipe. **(They are capped there now - `SweepBodyCharsCap` / `SweepBodyBytesBudget`,
  2026-08-18 - so the arithmetic below is the worst case as it WAS.)** That puts 64 MB at ~80 KB average body on a one-store profile, ~27 KB on three
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

- [ ] **Verify the sweep's new body bounds against a live profile - T1 owns every decision above the COM call and none of the COM half.**
  `SweepBodyCharsCap` / `SweepBodyBytesBudget` landed 2026-08-18 and are pinned by T1
  `SweepBodyCapTests` (18 tests, mutation-checked). What T1 owns is the pure cut, the byte ceiling
  against the real serializer, the frame-half invariant, the payload fields, the coverage code, the
  advice split and the body-cache guarantee. What it cannot produce is a real `item.Body` big enough
  to cut, so **no swept body has ever actually been truncated on this machine.**

  - **What to confirm read-only, and it needs the shape the whole frame-size item is about:** a
    LOCAL store (PST/archive) absent from the index, so the sweep window falls back to
    `EmptyIndexSweepWindow`, holding enough recent large mail to move real body volume. Then an
    ordinary `search` naming that store should report `comHost.largestFrameBytes` (in
    `outlook_health`) well under `frameLimitBytes` with `framesRefusedTooLarge: 0`, and - only if
    something really was cut - `sweep.itemsBodyCapped` with its advice sentence.
  - **The two numbers worth measuring while that store is mounted**, because both are predictions
    rather than observations: the largest frame such a sweep actually produces (to see how much of
    the 32 MiB budget real mail uses), and the per-item cost of the body pass, which is one linear
    scan of at most `SweepBodyCharsCap` characters per item on top of the COM `.Body` read that
    produced it. The scan is expected to be noise next to the read; nothing has timed it.
  - **Unverifiable without a mailbox that has one:** whether any real mail body exceeds 500 000
    characters at all. If none ever does, the per-item cap is pure insurance and only the budget
    can ever bite - which would be the good outcome and should be recorded as such.

- [ ] **Measure the sweep and scan budgets against a known corpus - the tooling exists, the corpus does not yet.**
  The store shape the item above asks for cannot be borrowed from anywhere: every store on the real
  profile is indexed, so `EmptyIndexSweepWindow` never engages, and the Hyper-V VM's PST is the only
  unindexed store to hand and it is empty. So the corpus is built rather than found.
  `McpServer/OutlookAI.RemediationTools` gained five commands for it - `corpus-plan`, `corpus-probe`,
  `corpus-build`, `corpus-teardown`, `corpus-reindex` - and `Docs/corpus-measurement-plan.md` is the
  plan for what to run against the result and what each number would settle. T1
  `CorpusGeneratorTests` pins the size distribution, the date spread, the seeding, the store
  refusals, the teardown rule, the manifest format and the date-fidelity verdicts; only the COM
  calls are outside that tier, and they carry no decisions.

  - **Built once, 2026-08-19, and the measurement still has not been taken.** 40 000 items into
    the VM's PST in 12m27s at 50.9 items/sec, zero failures; resumability and determinism both
    demonstrated (a re-run skipped the 2 000 items an earlier timing run had made). Three faults
    came out of it, all now guarded in code rather than written down as cautions:
    - **Items were queued for delivery.** 5 532 landed in the target store's **Outbox** - inert
      on that VM only because its profile has no mail account. The store guard could not catch
      it: "local .pst" and "an account's delivery store" are not mutually exclusive.
      `CorpusSafety.EvaluateProfile` now refuses unless `Session.Accounts` is EMPTY, with no
      override. Stricter than "no account delivers here" because the object model cannot express
      the narrower rule - `SendUsingAccount` is per item, so any account may send a message that
      lives anywhere.
    - **Every item was filed as a draft**, so the sweep saw 6 of 40 000 in 234-367 ms.
      `Items.Add` + `Save` produces an UNSENT item and Outlook files those in Drafts whatever
      folder they were added to; the sweep covers Inbox/Sent/Deleted/Junk and not Drafts. The
      root cause was a design fault of mine: the MSGFLAG_UNSENT write lived as a rung of the
      DATE ladder, so `--allow-undated` silently disabled placement too. `CorpusPlacement` is now
      its own probed ladder, and a rung passes only when the item's `Parent` is the target folder
      AND that folder's `GetTable` returns it.
    - **`corpus-reindex` and the post-teardown count looked only at Inbox/Sent/Junk/Deleted**, so
      with all 40 000 items in Drafts the recovery path would have reported ZERO and teardown
      would have claimed a clean store. Drafts (16) and Outbox (4) are in the scan set now. Same
      lesson as the Outbox omission `ComMailbox.SweepFolderIds` already records.
  - **Still to do: tear down `CP-07-CORPUS-40K` and re-run.** The 40 000 drafts are still in the
    PST; the manifest is 40 002 lines and a copy is outside the guest. Teardown deletes by
    EntryID allowlist AND tag, and now reaches Drafts.
  - **The date verdict from that run proves nothing and must not be quoted.** The probe reported
    `readBack` correct with `daslIn=False`, which reads as "the date does not drive selection" -
    but the item was in Drafts while the probe queried the Inbox's table, so "not in this folder"
    explains it equally well. The two failures were indistinguishable in the output. The probe now
    settles placement first and builds its date probe with the placement that verified, so a
    re-run isolates the question.
  - **The one thing that must be settled first, before any of it is worth doing:** whether that PST
    accepts back-dated mail at all. `MailItem.SentOn` is read-only in the object model, and an item
    created straight into a folder is UNSENT, which some stores date themselves. `corpus-probe`
    settles it empirically - it writes one throwaway item per method, re-opens it by EntryID, reads
    `ReceivedTime` back, and then asks a DASL date restriction on either side of the instant whether
    it selects the item - and `corpus-build` refuses to build an undated corpus unless
    `--allow-undated` says so in as many words. **An undated corpus would make both windows select
    the same population while looking exactly like a good corpus**, which is why the refusal is the
    default rather than a warning.
  - **A product gap fell out of it, recorded as H3 in `Docs/completeness-gaps.md` and OPEN.** The
    sweep restricts on `(datereceived >= X) OR (date >= X)`, so mail carrying neither property is
    selected by NO window - absent from the freshness tier rather than mis-dated, and absent
    however wide the window is opened - while `sweep.foldersSwept` still counts the folder and
    `freshness` still says `live`. Real users hit this with imported, copied or restored mail.
    The filter is code and is not in doubt; the supporting observation is confounded by the
    placement fault above, and the row says so.
  - **The blocker for the 180 s proposal, found while writing the plan and not yet acted on:**
    `SearchBudgetMs` is `SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs`, and T1
    `BudgetCompositionTests.SearchBudget_IsComposedFromItsPartsAndFitsTheOperationDeadline` asserts
    that sum fits inside `ComOperationBudgets.OperationDeadlineMs` (120 s). At 180 s the sum is
    195 s and that test fails before anything reaches a mailbox. Anything above roughly 105 s moves
    the operation deadline too, and with it the child work budget and `ExhaustiveTimeBudgetMs`. Decide
    the shape of that change before measuring, so the measurement is aimed at the right question.

- [ ] **Verify the three folder-walk reporting fixes against a live profile - the COM half none of them can reach from T1.**
  G2, G3 and G4 of `Docs/completeness-gaps.md` were closed on 2026-08-18 and are pinned by T1
  `FolderWalkReportingTests` (21 tests, driving the real `MailService` through a stand-in session
  and index client; mutation-checked - removing any one of the three fix lines fails 4 of them).
  What T1 owns is everything ABOVE the COM call, which is where the whole defect lived, since all
  three drops were decisions taken there. What it cannot produce is the COM failure itself:

  - **G2 needs a store whose `DisplayName` read throws.** Never observed on any machine here; it
    was found by reading the `catch` that swallowed it. Nothing in the repo knows what actually
    provokes it - a damaged profile entry, a data file that will not open, a store mid-removal are
    guesses. Worth trying: mount a PST, then rename or delete the file underneath Outlook while a
    session holds it. What to confirm read-only: the store appears in `list_folders` under
    `(unnamed store N)` with `nameUnreadable: true`, an unscoped `search` reports
    `sweep.storesUnnamed`, hits from it carry the label as their `store`, `outlook_health` lists it
    (via the same COM store list, so it should now reach `index.perStore[].inLocalIndex: false`),
    and `search(store: "(unnamed store N)")` is refused with the placeholder message rather than
    the typo one. **Also unverified in the real world:** whether a store that will not name itself
    will still answer `GetDefaultFolder` and `GetTable`, i.e. whether the sweep of it actually
    returns mail rather than four skips. The code handles both; only a live case can say which
    happens.
  - **G3 needs 10 000 folders in one profile, or a 65-level-deep tree.** Both are constructible in
    a test PST and neither exists here. The cheap partial check is a temporary build with the cap
    lowered (it is a `public const` on `MailService`, read at the call site) against the real
    profile: confirm `truncated: true`, `walkCapReached: true`, NO `nextOffset`, and the advice
    sentence - and that a store-by-store listing then returns the tree the capped call could not.
  - **G4 needs a delegate mailbox whose folder walk hits a bound**, i.e. G3's condition inside a
    shared mailbox. The dev profile has two delegate mailboxes indexing 11 and 23 folder paths, so
    the honest statement is that this flag has never fired on real data and cannot until the walk
    cap is reachable. The lowered-cap build covers it in the same pass: a delegate folder search
    should then report `scope.folderNamesTruncated: true`, `degraded: true`, and the INCOMPLETE
    SCOPE sentence.

  Read-only throughout - `list_folders`, `search`, `outlook_health`, `list_accounts`. No mailbox
  writes are needed for any of it.

- [ ] **Verify the sweep's sort-failure detection against a live profile - the one half of H2 that T1 cannot reach.**
  H2, G5, B2 and F3 of `Docs/completeness-gaps.md` were closed on 2026-08-18 and are pinned by T1
  `SearchCoverageClaimTests` (20 tests driving the real `MailService` through a stand-in session
  and index client, mutation-checked: five separate revert-one-line runs each fail exactly the
  tests that own that line). Three of the four are settled by that, because the whole defect lived
  above the COM call. H2 is not, and the gap is narrow and specific:

  - **`SweepFolder`'s new `out bool sortApplied` is set inside the `catch` around
    `Table.Sort`, and that catch has never been observed to fire on any machine here** - the
    defect was found by reading the swallowed exception, not by hitting it. So the value the whole
    row turns on is produced by a line no test executes. What IS pinned: the flag's journey across
    the process boundary (`ComHostProtocolTests.ComSweepResult_CarriesTheUnsortedCappedFoldersAcrossTheWire`),
    its per-store filtering in `ApplySweepCounters`, the code split, and both advice sentences.
  - **What would provoke it is unknown.** `Table.Sort` needs the property present as a column, and
    the code adds `urn:schemas:httpmail:datereceived` to `Columns` before sorting, so the ordinary
    path cannot fail. Guesses worth trying, none verified: a folder whose `DefaultItemType` is mail
    but whose contents are not, a folder on a store mid-reconnect, or a search folder. A cheaper
    substitute is a temporary build that forces `sortApplied = false` and confirms the payload
    end-to-end on the real profile.
  - **What to confirm read-only**, once a folder can be made to both refuse the sort and exceed
    `SweepPerFolderCap` (200 items in the freshness window - see the `EmptyIndexSweepWindow` row in
    `Docs/magic-numbers.md` for how wide the window has to get): `sweep.itemCappedFolders` names the
    folder, `sweep.itemCappedFoldersUnsorted` names it too, `sweep.coverageGaps` carries
    `item_cap_unsorted` and NOT `item_cap`, and the advice contains the word ARBITRARY and neither
    "newest-first" nor "OLDEST".

  Also unmeasured, and cheap to settle on the dev profile while the above is open: how often the
  G5 probes now run. The trigger widened from "the merged answer was empty" to "the index tier
  returned no rows", which costs two TOP-1 statements per folder-scoped search that the index did
  not answer. It was accepted on the standing rule that completeness outranks speed, so the number
  is worth having rather than worth acting on.

- [x] **F2 DECIDED AND SHIPPED (uncommitted, 2026-08-19): the exhaustive scan is resumable.**
  The maintainer chose **(c) make it resumable**, and **(a) order the walk** came with it as its
  stated prerequisite rather than as an alternative. They also chose to leave `top` at 100 and rely
  on resumption, because payload is context and context is the scarce resource. What shipped:
  `ExhaustiveScan` enumerates its scope in the shared sibling order first (so `foldersTotal` is
  honest even when the walk covered four of thirty-two), then walks that list; a stop returns
  `exhaustive.nextToken` over walk state in the SERVER PARENT (`ExhaustiveScanCursors`); per folder
  the ladder is date cursor, validated ordinal, folder restart with EntryID suppression, and
  `position.resumeTier` says which paid. `Docs/completeness-gaps.md` F2 carries the full record.
  **Still needs a live profile** - the whole of `T2/LiveResumableScanTests`, which is the only tier
  that can prove a paged scan returns exactly what an unpaged one returns.

- [ ] **Settle whether `Table.Sort` has EVER applied - run `T2/LiveTableSortProbeTests` and act on the answer.**
  Potentially the largest single defect found on 2026-08-19, and it is unresolved rather than fixed.
  Microsoft's `Table.Sort` reference says a sort property may be referenced "by their explicit string
  names only; cannot reference properties by their namespaces". `SweepFolder` passes
  `urn:schemas:httpmail:datereceived`, which is a namespace. If the documentation holds for this call
  then the freshness sweep has **never sorted on any store for any user**, its 200-item cap has always
  cut arbitrarily, and the tier whose entire purpose is recent mail has been returning an arbitrary
  200 rather than the newest 200 - a completeness defect, not only a reporting one. It would also mean
  this session's reading of `item_cap_unsorted` ("the sort genuinely does not apply on that store") is
  wrong: it would not apply anywhere.
  - **Four read-only PowerShell probes could not settle it.** Every property form - the namespace
    form, `ReceivedTime`, the DASL proptag form, `SentOn` - and every argument shape, including the
    no-argument form, failed identically with `DISP_E_PARAMNOTOPTIONAL`. That uniformity is the tell:
    it is PowerShell late binding against the `Table` COM object, not Outlook's verdict.
  - **The probe is written and NOT run:** `T2/LiveTableSortProbeTests`, read-only, one table per
    store, both spellings, each in its own try/catch, printing a per-store verdict and a single
    ANSWER line. `Category=Live`, so it needs the configured dev machine.
  - **Two cheap things shipped alongside it, and one deliberately did not.** The `catch` that wrapped
    `Columns.Add` and `Sort` together is split, so `sortApplied: false` can say which failed; the
    folders where the column WAS added and `Sort` still threw are counted into
    `sweep.sortRefusedFolders`, which answers the question from ordinary telemetry (equal to the
    folders swept, everywhere, means the property name is the cause). **The sort call itself was NOT
    changed**, because changing it before the probe runs destroys the evidence.
  - **If the hypothesis holds:** change `SweepFolder`'s `Sort` to the explicit property name, re-check
    H2's advice sentence (it becomes true for the first time), and note that the resumable scan's date
    rung becomes the normal path rather than the lucky one. **If it does not:** `item_cap_unsorted` is
    correct as it stands and the scan will live on its ordinal and restart rungs on this profile.

- [ ] **Verify the exhaustive scan's depth guard against a live profile - the half of F4 that T1 cannot reach.**
  F4 was closed on 2026-08-18 and is pinned by T1 `ScanDepthAndSweepScopeTests` end to end from
  `ComExhaustiveResult` to the payload, plus the process-boundary round trip in
  `ComHostProtocolTests`. What no test executes is the guard itself: `if (depth > FolderWalkDepthGuard)`
  needs a folder tree more than 64 levels deep, which no CI runner and no real mailbox has. The same
  shape as H2's `sortApplied`, and the same cheap substitute applies - a temporary build with the
  guard lowered to 2 or 3, then one read-only `exhaustive: true` search of a store with any nesting
  at all, confirming `exhaustive.depthLimitReached: true`, `depth_limit` in
  `exhaustive.coverageGaps`, `freshness: "partial"`, `degraded: true`, and an advice sentence
  naming the guard's value. Worth pairing with the lowered-cap build the G3/G4 item above already
  asks for; both are read-only and neither needs a mailbox write.

- [x] **Settle C5 (`Docs/completeness-gaps.md`) - `thread`'s `store` narrows the evidence its own coverage code is computed from.**
  Found while closing C4 and recorded rather than fixed, because the fix changes which rows come
  back. A `store` on `thread` scopes the index query, and the store is auto-derived from the
  referenced hit when the caller passes `id` without `conversation_id` - so on those call shapes the
  index rows can only ever name one store and `unwalked_store` cannot fire, leaving a member in a
  second INDEXED account both absent and unreported. C4's fix covers the unindexed half of this
  (it reads Outlook's store list rather than the index rows). The candidate directions and their
  trade-offs are on the C5 row; the coherent one is to stop scoping the conversation query at all,
  since `thread`'s `store` is documented as a speed hint rather than a filter - which is exactly
  C3's own reasoning for allowing it to widen.
  **DONE 2026-08-23 on the maintainer's decision, which was none of the three listed: derive the
  warning from Outlook's store list on the scoped shapes too, exactly as C4's fix does.** So the
  scope is still applied and the same rows still come back - it is a reporting fix, which is why it
  could be taken without the acceptance the "changes which rows come back" note was waiting for.
  `live.storesNotQueried` + `unqueried_store` + `scopeStore` + `scopeStoreDerived`; T1
  `ThreadScopedStoreTests`. See the C5 row for what each field says and why the remedy differs by
  how the store arrived.

- [ ] **Decide whether `thread` should apply a store scope it DERIVED (C5's remaining half).**
  C5 is closed on reporting, and the underlying behaviour is untouched: a caller who passes `id`
  alone still gets a lookup narrowed to that hit's store, so a member in a second indexed account is
  still ABSENT - now named as unqueried rather than unmentioned. Three directions:
  - **(a) Leave it.** Defensible: the narrowing is a real speed win on a large profile, the payload
        now says it happened, and the remedy (pass `conversation_id` beside `id`) is one argument.
        The cost is that the default call shape returns a partial conversation on every multi-store
        profile, and `unqueried_store` therefore fires often enough to blunt it.
  - **(b) Stop scoping when the store was DERIVED; keep scoping when the caller named one.** The
        distinction the code already draws for the advice (`scopeStoreDerived`), applied to the
        behaviour: nobody asked for the narrowing, so do not apply it. `unwalked_store` then works
        as designed on that shape - it makes the STRONGER claim off index rows that can finally name
        the other store - and `unqueried_store` goes quiet unless a caller chose a scope. Costs one
        unscoped `ConversationID` query per `thread` call.
  - **(c) Stop scoping the conversation query entirely**, using `store` only to order results. The
        original TODO's preference and C3's own reasoning (`thread`'s `store` is a speed hint, not a
        filter). Widest and simplest to explain; pays the unscoped query on every call.
  Recommendation: **(b)**. It removes the defect exactly where nobody asked for it, keeps the hint
  a caller explicitly chose, and leaves the reporting that now exists to cover the case it keeps.

- [ ] **Decide what the test VM is FOR, because 96 of 115 live tests cannot move to it as it
      stands.** The live tier is now split by trait (`LiveTier=Portable` vs `ProfileBound`, see
      `Docs/live-tier-on-the-vm.md`), and the Portable subset is 19 tests. The limit is not
      configuration; it is that a profile with no mail accounts cannot create a draft at all
      (`NewDraft` resolves an Account by SMTP address and refuses when none matches), which takes
      the draft, update/discard, HTML-draft and send families off the table whatever else is
      arranged. Four directions, none of them started:
  - [ ] **Add one dummy mail account to the VM** (POP/IMAP pointing nowhere, send disabled). Would
        unblock the draft families, and it is the single highest-yield change. **The catch:** the
        corpus generator refuses to run at all unless the profile has no accounts whatsoever, so
        the order is corpus first, checkpoint, then account - and re-generating later means
        removing the account again.
  - [ ] **Give the hub store an SMTP-shaped display name.** Several tests use
        `testHubStoreDisplayName` as an address (`to: Hub`, `FindAccountBySmtp(Hub)`), so a PST
        called `Outlook Data File` fails them before anything else does. Cheap to try; unknown
        whether Outlook tolerates it.
  - [ ] **Accept 19 and stop.** Defensible: the 19 include both acceptances the project is blocked
        on, the sweep-scope and sweep-cache behaviour, and the signature lifecycle. It means the
        index, accounts, delegate stores and send path are only ever proven on the maintainer's
        own profile, before a release.
  - [ ] **Relax the tests instead of the machine** - make the account-count and store-count
        assertions read from the settings rather than hardcoding three. Would move more tests, and
        would weaken exactly the assertions that catch a misconfigured profile.

- [ ] **Add the second PST to the test VM - it is what makes the count tripwire mean anything
      there.** The tripwire exempts the hub store, so a machine whose only store IS the hub gives
      it nothing to watch: it will census, report zero failures, and be structurally incapable of
      reporting anything else. Recommended layout and the reason are in
      `Docs/live-tier-on-the-vm.md` section 2.3. Add it through Outlook's own UI (File > Account
      Settings > Data Files > Add) rather than a script - creating stores is not something the
      tested helpers do, and mailbox mutation from ad-hoc shell code is the thing that once
      destroyed real mail.

- [ ] **Put a few hundred items in the SECOND store, not the corpus, or the identity half of the
      count tripwire is never exercised.** The identity budget is 500 items per folder and 3,000
      per store, so a small store is walked item by item and a corpus is not: all four populated
      corpus folders (Inbox 10,912 / Sent 4,964 / Deleted 2,467 / Junk 1,663) are above the
      per-folder limit and fall back to counts. Note also that the corpus store must be the HUB -
      the Portable scans and sweeps all target the hub and take a "corpus too small" early return
      against an empty one - so the second store is the only one the tripwire can watch anyway.

- [ ] **Live-only, and unguarded by any non-live test: three decisions inside the count tripwire's
      verification.** Established by construction rather than by mutation, because they sit behind
      a COM census that no CI test can execute: `CollectionFinished`'s early return on `NotLast`;
      the keep-alive release and the `_verified` latch in `Verify(final)`; and the key-based
      intersection in the confirmation census. The KEY itself is pinned in
      `T1/StoreCountTripwireTests`; the code that uses it is not. The cheap substitute is the same
      one this file already records for `sortApplied`: a temporary build that forces the branch.

- [ ] **Watch for the count tripwire firing on a PST's Junk folder.** The census marks self-pruning
      folders by asking the store for its default Deleted Items, Junk and sync-issue folders. A PST
      may refuse `GetDefaultFolder` for Junk, in which case a generator-made "Junk Email" folder is
      an ordinary folder and a decrease in it FAILS rather than being noted. Nothing prunes it on a
      machine with no accounts, so this should stay theoretical - but if the tripwire ever fires on
      Junk, this is why, and the fix is to mark volatility by folder NAME as well as by default-folder
      identity.

- [ ] **Two live tests still degrade silently, and were left that way deliberately.**
      `LiveDraftTests.ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain` and its `LiveSendTests` twin
      loop over `expectedStoreDisplayNames`, so on a machine with fewer stores they sweep fewer and
      still pass. Both are `ProfileBound`, so they do not run on a test machine; changing a sweep
      assertion without a live run to check it was not a trade worth making unsupervised. The other
      two of the four (`LiveStaleIndexRowTests`, `LiveManageSignatureTests.DefaultAssignment`) now
      refuse on a Production profile.

- [ ] **Residual gaps left by the 2026-08-23 tier-3 classification pass.** `Category!=Live`
      no longer reaches a mailbox: eleven T3 tests that called `outlook_health`,
      `list_accounts` or `search` moved into `ComHostSupervisionLiveTests`,
      `OutlookAvailabilityLiveTests` and `OutlookHealthLiveToolShapeTests`
      (`Category=Live` + `LiveTier=Portable` + `Requires=OutlookInstance`), `McpStdioClient`
      now refuses those three tools unless a test declares them, and
      `T1/LiveTierInventoryTests.EveryStdioTestReachingOutlook_DeclaresIt` reads the
      declaration back out of the IL. What that pass found and did NOT fix:

  - [ ] **`list_accounts` starts Outlook, and nothing in the tool layer stops it.** The
        supervisor's liveness verdict for `NotRunning` is `MayStart`, which calls
        `BeginWarmUp` and connects to `Outlook.Application` - so on a machine with Outlook
        installed but closed, a bare `list_accounts` launches it. `outlook_health` guards its
        own probe with `if (outlookRunning)`; `list_accounts`, `list_folders`, `read`,
        `search`'s sweep and every draft path do not. Correct for a shipped tool
        (S7/D17 permits the cold start); a hazard for a test tier, and it was reachable from
        the default run until this change. Decide whether the live tier should force
        `allowStartingOutlook: false` for the T3 stdio tests, which would need a server-side
        switch it does not have.
  - [ ] **Four T3 tests pass while asserting almost nothing on a machine with no Outlook**,
        each by an early `return` that is documented where it sits. They are not new and none
        is wrong, but together they are why "the CI tier is green" said less than it looked:
        `OutlookAvailabilityLiveTests.SearchAlwaysAnswers_AndSaysWhetherItIsComplete` (returns
        the moment `search` reports an error, which on an indexless machine is every run),
        `...ATransientOutlookState_AnswersFastAndCarriesRetryGuidance` (returns when Outlook
        is healthy, keeping only the timing assertion),
        `ComHostSupervisionLiveTests.NoComHostSurvivesTheServer` (returns when no COM host was
        spawned, keeping only the stdin-close assertion) and
        `OutlookHealthLiveToolShapeTests.OutlookHealth_CarriesTheFreshnessBlock_WithOrWithoutAnIndex`
        (skips the advice assertion when the index provider is unavailable). All four are now
        `Category=Live`, so the question is what the VM run should assert INSTEAD of returning.
  - [ ] **The pin reads tool NAMES, not arguments.** `search`, `read`, `thread`, the draft
        tools, `move_mail` and the show-me tools all have a refusal that fires before any COM
        work, which is what the protocol-only half of T3 is built on - so they cannot be
        blanket-guarded, and a future test that calls one with arguments that DO reach Outlook
        would not be caught. A bounded exhaustive `search` is the realistic case.
  - [ ] **`McpServer/README.md` said `LiveTier=Portable` was "19 of 115 methods".** The
        measured figure after adding 11 methods is 31 of 127, so the documented number had
        already drifted by one before this pass. Updated to the measured value; worth a
        thought about whether that line should be generated rather than typed.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
