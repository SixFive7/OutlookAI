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

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)
