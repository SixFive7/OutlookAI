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

- [ ] **IN FLIGHT 2026-08-18 09:00 - live tier: subset PASSES, full tier does not. Read this first.**

  **New evidence, all from 2026-08-18 08:45-09:00, which overturns yesterday's hypothesis:**
  - **Outlook COM is healthy.** A bounded out-of-process probe attached and read `Stores.Count = 5`
    instantly. The "poisoned COM state" theory from the overnight notes is WRONG.
  - **The short subset now passes end to end.** `--filter FullyQualifiedName~LiveSweepScope` under
    `--blame-hang`: `[preflight] Outlook responsive (0 of 5 UI windows hung) - live tier may run.`,
    `[tripwire] baseline: 5 stores, 148 mail folders, 12495 ms.`, tests passed, `hub reconciled: 5
    baseline folders, all item counts back at pre-run values`, exit 0.
  - **The FULL tier still hung** at `A total of 1 test files matched` for 19 minutes earlier the same
    hour, with the preflight in place and Outlook responding.

  **So the hang is NOT the shared tripwire and NOT a wedged Outlook.** It is something in a
  collection the subset does not touch. The next diagnostic is to run the tier collection by
  collection under `--blame-hang --blame-hang-timeout 4m` (that flag works and produces a sequence
  file when a hang is real) and find which fixture stalls. Do NOT run the whole tier blind again -
  it costs 20 minutes and produces no information.

  **Artifacts:** 7 tagged items (6 Drafts, 1 Outbox) were left by an aborted run. The subset run above
  reported the hub reconciled, so they may now be swept - VERIFY with a read-only `search` for
  `OutlookAI-McpTest` before assuming either way.

- [ ] **PENDING TASK - process `C:\Source\SixFive7\BrowserAI\.work	runcation-prompt-for-sibling-project.md`.**
  The maintainer asked for this at 09:00 on 2026-08-18. It is expected to be the portable
  description-budget prompt written for another project; read it and act on what it asks for. Recorded
  here because auto-compaction was imminent when it was requested.

- [ ] **`TryCreateDerivedDraft`'s cross-store retry is unguarded.** `MailService.cs:2169` re-attempts
  draft creation across every store on `r == null` alone, while its sibling loop in `TryUpdateDraft`
  (`MailService.cs:2501`) only retries on `error == "ItemNotFound"`. So a creation that failed for
  some other reason is retried against store after store. Found 2026-08-18 while classifying which
  session operations are safe to retry after a disconnect; left alone because tightening it silently
  is the wrong move - it needs a decision about which failures should fan out and which should stop.

- [ ] **One door back to the cross-store attribution defect.** `MailService.ApplySweepCounters` falls
  back to whole-sweep totals when a store is named but `result.PerStore` is empty - which is the
  pre-`c515565` behaviour that let one account's unreadable folder degrade another account's search.
  Unreachable today (both `ComSweepResult` construction sites populate `PerStore`, and an empty
  bucket list means nothing was walked at all), so it is a latent seam rather than a live defect -
  but it is the one path that reopens a closed defect, and it deserves either a guard or a comment
  saying why it cannot happen.

- [ ] **A useful error message never reaches the caller.** An `exhaustive` search naming a
  folder that does not exist comes back as an opaque
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

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
