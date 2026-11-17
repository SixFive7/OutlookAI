# The Q118 pin measurement of 2026-10-03, raw

The raw logs of the multi-session measurement of D49's lifetime pin: real MCP server processes, each
with its own COM host, against one Outlook on `OutlookAI-Unindexed` (Office LTSC 2024,
16.0.17932.20996), driven by `Testbed/guest/Measure-PinWindows.ps1`. The summary is
`McpServer/Docs/com-host.md` (the Q118 section) and the header of `T2/LivePinWindowTests.cs`; these
are the logs that summary was written from.

They lived only in an agent worktree's scratch folder until 2026-11-17, which is not a durable home,
so they are committed as they were. Nothing personal was in them: the guest's accounts are
`*@vm.invalid` and its user is the test machine's own. The two host drivers are archived as `.txt`
so nothing treats them as scripts to run.

| Path | What it is |
| --- | --- |
| `host-<scenario>-<time>.txt` | The host's capture of each scenario: Control, Visible, Started, Sequential, Killed, KilledVisible, ShowMe, ControlClose, Headless, StarterLeaves, End. |
| `guest/q118-<scenario>-<time>.jsonl` | The guest's record of each step of each scenario. |
| `guest/child-*.json` | Each child probe's reading - State, Quit, CloseHidden, ShowActive. |
| `guest-jobs-*.txt` | The interactive task's job log for each scenario. |
| `guest-live-test-*.txt`, `guest-live-test/q118-live.trx` | The run of `LivePinWindowTests` on the guest. |
| `Run-Q118Scenario.ps1.txt`, `Run-Q118LiveTest.ps1.txt` | The host drivers. |
