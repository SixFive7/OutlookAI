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

- [ ] **Audit the codebase for other timers and time-based behaviour.** The update poll above
  was found by accident while adding the manual check; nothing has ever looked at the set as
  a whole. Worth one pass to find polls that should be events, intervals that no longer match
  what they wait for, and anything that assumes wall-clock time moves forward smoothly.

  Known starting points — not a complete list, which is the point of the audit:
  - `Services/UpdateService.cs` — the 10-minute `System.Threading.Timer` poll, the 5-minute
    `HttpClient.Timeout`, and the `Get-Process outlook | Wait-Process; Start-Sleep -Seconds 2`
    in the handed-off installer script.
  - `McpServer/` — the COM-host supervision and health paths carry several timeouts and
    back-off intervals (the wedged-Outlook work added re-check and cool-down periods).

  Things to check for specifically:
  - **`DateTime.Now` used for elapsed time.** It is wall-clock and can jump backwards over a
    DST change or a clock sync. `Stopwatch` is the right tool where a duration is meant.
    **Done for the add-in:** `UpdateService` measures "checked 4m ago" from a process-wide
    `Stopwatch` (`_sinceStart` / `_checkedAtMs`) and the negative-interval clamp that was the
    symptom is gone; `AITaskPane`'s debug-click window is a `Stopwatch` too. `LastChecked`
    and the debug log's `HH:mm:ss` stamps stay wall-clock on purpose — those are absolute
    instants, not durations. `McpServer/` has not been swept for this.
  - **Timers that outlive what they poll**, and any that are never disposed.
  - **Intervals chosen once and never revisited** against what they are actually waiting for.

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

- [ ] **The store-count tripwire cannot tell the user apart from the tests, and says so loudly.**
  (2026-08-18) The 26.8-minute run above passed all 107 tests and then FAILED at teardown:

  ```
  STORE COUNT TRIPWIRE: the live tier changed mailboxes it may not touch.
    ITEMS LOST: store 'info@voipfabric.com' folder 'Inbox' 168 -> 161 (-7).
    ITEMS LOST: store 'Jan van Linge' folder 'Ongewenste e-mail' 1 -> 0 (-1).
    ITEMS LOST: store 'Jan van Linge' folder 'Postvak IN' 52 -> 50 (-2).
  ```

  **Evidence that the tests did not do it:** `StoreWriteAllowlist` throws in code on any write aimed
  outside the designated hub and never fired; all 107 tests passed; the artifact sweep reported
  `taggedArtifacts=0` for all three accounts. **Evidence for ordinary activity:** this was the first
  live run taken while the maintainer was actively using the machine, a 27-minute window covers mail
  read, deleted and rule-filed, a junk folder going 1 to 0 is what junk expiry looks like, and the
  tripwire separately noted a `Deleted Items (self-pruning)` folder appearing, which is Outlook's own
  auto-prune.

  **CONFIRMED 2026-08-18: the maintainer deleted that mail himself.** The alarm was a false positive
  and no mail was lost by the suite. That does not close the item - it promotes it from theory to a
  measured instance.

  **The design gap, now demonstrated rather than suspected:** the tripwire compares raw counts and
  cannot distinguish the user from the suite. It is deliberately fail-closed, which is right - but it
  means the live tier cannot be run on an actively used machine without a false alarm, and a real
  alarm would look identical. Directions: record per-item EntryIDs for the small stores rather than
  counts; or capture the count deltas the suite itself could plausibly cause and subtract them; or
  require the machine to be idle, which is what the S7 quit-when-safe guard already does for a
  different purpose. **Nobody should trust a green tripwire on a busy machine either** - that is the
  half of this nobody has looked at.

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

- [ ] **An answer too big to frame kills the COM host instead of being refused.** Found by the
  boundary audit on 2026-08-18; not fixed, because the fix is not the small one it looks like.

  `ComHostProtocol.EncodeFrame` refuses a payload over `MaxFrameBytes` (64 MB) by throwing
  `ComHostProtocolException` - a deliberate, specific, actionable failure. But it is thrown from
  `ComHostServer.WriteAsync`, which `ServeAsync` guards with `catch (IOException)` only. So the
  exception leaves the serve loop, `Program.Main` prints it to stderr and the child exits with 1.
  The caller learns "the COM host went away", which is the one fact that says nothing about what
  to do next; the sentence naming the size and the limit reaches only the child's stderr. This is
  the same species as the wrapper defect above - a good message that does not survive the process
  boundary - and it is the only other instance the audit found.

  **Why it is not just a catch clause.** A refusal frame is small, so answering with one is easy;
  what is hard is testing it. Provoking it needs a genuinely oversized response, which means
  allocating well over 64 MB in a T1 test that currently runs the whole tier in two minutes. The
  alternatives each cost something: measuring the encoded size inside `Invoke` (where the existing
  catch would turn it into a proper error frame for free) serializes every payload twice; making
  `MaxFrameBytes` settable adds a production seam for a test; a new fault-injection kind is more
  production-code-for-tests. **Pick one before writing it.**

  **How likely is it?** Low - `Docs/com-host.md` calls 64 MB "far above any real payload" and a
  `read` returns ~0.5 MB. The candidates are `SweepFoldersNewerThan(includeBodies: true)` and
  `ExhaustiveScan` over a large window. Nobody has measured the largest frame the product actually
  produces, and that measurement is worth having on its own: it is also the number that says
  whether 64 MB is the right limit.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
