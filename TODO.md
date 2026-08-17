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

- [ ] **Make the update indicators event-driven instead of polled.**
  `UpdateService` is a static class whose background timer writes `volatile` fields, and it
  raises nothing when they change. So both version indicators — `lblVersion` in
  `TaskPane/AITaskPane.cs` and the "Version and updates" group in `TaskPane/SettingsDialog.cs`
  — poll it on a 1-second `Windows.Forms.Timer` instead.

  That tick is doing two jobs at once, and only one of them wants to be fast:
  - **State changes** (`checking…` → `downloading v3.2…` → `ready - installs on close`)
    should appear promptly, and today only do so because the poll is quick.
  - **The relative time** ("checked 4m ago") changes at most once a minute, so for that the
    poll runs roughly 60× more often than the resolution it displays.

  Steps:
  - [ ] Add a `StateChanged` event to `UpdateService`, raised whenever `LastChecked`,
        `LastError` or `Status` changes. Note it can fire from a thread-pool thread (the
        poll timer and `CheckNowAsync` both run off the UI thread), so subscribers must
        marshal — `AITaskPane.InvokeOnUI` and the `BeginInvoke` pattern in
        `SettingsDialog.OnThemeChanged` are the two existing examples.
  - [ ] Subscribe both indicators, and unsubscribe on dispose. `SettingsDialog` is modeless
        and single-instance; `AITaskPane` is created per inspector, so several can be live
        at once and every one of them must detach.
  - [ ] Keep a slow timer (30–60s) purely for the "4m ago → 5m ago" rollover. This part
        genuinely cannot be event-driven: nothing in the service changes when a minute
        passes. It is the reason the tick cannot be removed outright.
  - [ ] Preserve the property `SettingsDialog.RefreshVersionLine` already has — it returns
        whether the *geometry* changed, not whether the text did, so a repaint that needs no
        re-layout cannot reset a scrolled dialog to the top. Any rework must keep that,
        because `PerformDialogLayout` scrolls to the top by design (see the note below).

  Not blocking anything: the current cost is a string compare and no layout, so this is
  tidiness rather than a fix.

- [ ] **Audit the codebase for other timers and time-based behaviour.** The update poll above
  was found by accident while adding the manual check; nothing has ever looked at the set as
  a whole. Worth one pass to find polls that should be events, intervals that no longer match
  what they wait for, and anything that assumes wall-clock time moves forward smoothly.

  Known starting points — not a complete list, which is the point of the audit:
  - `Services/UpdateService.cs` — the 10-minute `System.Threading.Timer` poll, the 5-minute
    `HttpClient.Timeout`, and the `Get-Process outlook | Wait-Process; Start-Sleep -Seconds 2`
    in the handed-off installer script.
  - `TaskPane/AITaskPane.cs` and `TaskPane/SettingsDialog.cs` — the 1-second version ticks.
  - `McpServer/` — the COM-host supervision and health paths carry several timeouts and
    back-off intervals (the wedged-Outlook work added re-check and cool-down periods).

  Things to check for specifically:
  - **`DateTime.Now` used for elapsed time.** It is wall-clock and can jump backwards over a
    DST change or a clock sync. `UpdateService.DescribeState` already clamps a negative
    interval to zero, which is a symptom of exactly this; `Stopwatch` is the right tool where
    a duration is what is meant.
  - **Timers that outlive what they poll**, and any that are never disposed.
  - **Intervals chosen once and never revisited** against what they are actually waiting for.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
