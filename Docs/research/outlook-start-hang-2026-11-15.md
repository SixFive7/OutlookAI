# The Outlook start hang of 2026-11-15

**What this is.** The evidence behind Q175 to Q182 (`QUESTIONS.md`, "The Outlook hang of
2026-11-15"): what happened on the maintainer's workstation, what caused it as far as that was
established, and what was measured. It condenses the diagnosis made that day and the investigation
item that stood in `TODO.md` until 2026-11-18, whose full text is in git history
(`git show 3dc85ea^:TODO.md`). The full report, with the machine's real names, stays in the
maintainer's gitignored `.work`; nothing here names a person, a mailbox or a machine. Times are
local (CET).

**Labels.** `[MEASURED]`: read off the machine - its event logs, process and thread records, wait
chains and Claude Code transcripts - or measured on the test machines. `[INFERRED]`: the diagnosing
agent's reading of those, not proven. Office was the maintainer's Office 2021, Click-to-Run, with a
cached Exchange profile and the Teams Meeting Add-in.

## Timeline

| Time | What happened | |
| --- | --- | --- |
| 16:21:10 | An OutlookAI MCP server starts for a Claude Code session doing unrelated research. Six servers from earlier days are still running; eight were alive during the hang. | `[MEASURED]` |
| 16:27:39.467 | A sub-agent of that session calls `list_accounts`. Outlook is closed. | `[MEASURED]` |
| 16:27:39 | `OUTLOOK.EXE -Embedding` starts with DcomLaunch as its parent: Outlook started over COM, with no window. | `[MEASURED]` |
| 16:27:41 | Outlook loads its seven add-ins: OutlookAI in 656 ms, the Teams Meeting Add-in in 16 ms. | `[MEASURED]` |
| about 16:27:44.7 | Outlook's main thread enters the wait it never leaves. | `[INFERRED]` from the wait-chain tick |
| 16:27:44.8 to 16:27:48.1 | A Claude Code session with the prompt `-` runs in `C:\WINDOWS\system32` and exits after 3.3 s: the add-in's start-up warm-up, as run from an Outlook started over COM. | `[MEASURED]`; its origin `[INFERRED]` |
| 16:44:15 | The user asks a new Claude session why Outlook does not start. A launch from the shell hands itself to the running instance, which is deadlocked and has no window, so nothing appears. | `[MEASURED]`; the hand-off `[INFERRED]` |
| 16:50:00 | The user interrupts the `list_accounts` call, 22 minutes in. The server's call into Outlook stays outstanding. | `[MEASURED]`; the outstanding call `[INFERRED]` |
| 17:02:02 | The diagnosing agent ends the hung Outlook, 34 minutes after it started. 67 ms later the same server, with no request pending, starts another `OUTLOOK.EXE -Embedding`, which starts normally. | `[MEASURED]` |

## The cause

`[INFERRED]`, from the wait chains and the main thread's stack, with Office's public symbols only:

1. Serving a late-bound object-model call from the server's connect sequence - most likely its
   `Logon`, or the hidden Explorer it opened to keep Outlook running - Outlook opened the Exchange
   offline store on its main thread and took a lock in `OLMAPI32!HrCreateOfflineObj`.
2. Inside that it made an outgoing COM call to the Network List Service (`netprofm`, `GetCost`). A
   single-threaded apartment waiting on an outgoing call keeps dispatching incoming ones.
3. In that window the Teams Meeting Add-in's start-up telemetry, on a CLR thread-pool thread, asked
   the object model for the current user's SMTP address. The call ran on the main thread on top of
   the first one, needed the MAPI async-task singleton, and waited for it without pumping messages,
   on something the outer frame held.
4. Neither frame could finish. `[MEASURED]`: 41 threads queued behind them, 31 on that one lock; the
   process used 0.000 s of CPU in 6 s; the instance never showed a window.

It takes three things at once - a call that opens the stores on the main thread, an outgoing call
inside it, and an add-in's call landing in that window - so it depends on timing: the second cold
start, at 17:02:02, did not deadlock. In this start over COM, Outlook had loaded its add-ins before
anything opened the stores, and the first call from outside then opened them while the add-ins' own
start-up work was calling in; a user's start opens the profile first (general knowledge, not
verified here). A hang on 2026-11-13 may have been the same; Windows kept no dump.

**Ruled out** `[MEASURED]`: elevation (every Outlook and server process ran at Medium integrity),
new Outlook, an Office update, a crash, a busy loop, `netprofm` hanging for everyone (a fresh client
got an answer in 108 ms), an add-in Outlook had disabled or flagged, no network, and another program
calling into Outlook (the wait chains of every MCP server and Explorer found only the one server).
**Not established:** which object-model call the outer frame was, what the inner wait waited for,
and what inside the server started Outlook again at 17:02:02.

## What it led to

| Decision | What in the hang it answers | What it became |
| --- | --- | --- |
| Q175 (a), (e) | Mailboxes were opened within a second of a start over COM, and each of eight servers could start Outlook. | The settle wait - Outlook 20 s old, the add-in's start-up stamp, the main thread idle - and the start lock, one connection at a time (`McpServer/Docs/com-host.md`, "Starting Outlook without racing its add-ins"). |
| Q176 (c) | An Outlook with no window stayed "starting" for 34 minutes with nothing saying so. | A 120 s limit on starting, and a per-user record of the Outlook OutlookAI started and the step it is in. |
| Q177 (a), Q178 (e) | Only ending OUTLOOK.EXE helped, and the agent's advice - restart Outlook, it clears by itself - was wrong for an Outlook nobody could see. | The report: what is stuck, the process ID, the Task Manager steps and what ending it risks; health's `outlook.start`; the start log `%LOCALAPPDATA%\OutlookAI\outlook-start.log`; the README's troubleshooting entry. |
| Q179 (e) | The interrupted request's call stayed outstanding, and Outlook was started again 67 ms after it was ended, for a request abandoned 12 minutes before. | A cancelled request keeps its deadline, and the COM host's retry only attaches. |
| Q181 (d) | The test machines cannot reproduce it: none has Teams, and only the Exchange VM has an Exchange cache. | No reproduction; the prevention and the start log ship, and hangs are watched (`TODO.md`). |
| Q182 (a) | The add-in ran a Claude Code session at every Outlook start, from `system32` when a program started Outlook; it played no part in the hang. | Removed. |

## What the settle wait costs

`[MEASURED]` on 2026-11-18 on the two Outlook test machines (Office LTSC 2024, no Teams), by the
disconnect test, which stops Outlook and has OutlookAI start it, and the start test:

| | Unindexed machine, three runs | Indexed machine, one run |
| --- | --- | --- |
| Cold start: the settle wait | 18.0 to 18.6 s, six starts | 18.3 s and 18.6 s |
| Cold start: Outlook's age when it settled | 20.2 to 20.5 s | 20.3 s and 20.0 s |
| Cold start: the whole connection | 21.7 to 22.2 s | 21.7 s and 21.2 s |
| Attach to an Outlook minutes old: the settle wait | 20 to 37 ms | 21 and 25 ms |
| An Outlook opened with its window: the add-in started | 6.2 s, 14.2 s and 27.9 s in | 171 s in |

Before the wait, a cold start connected in about 4 s. In every cold start the 20 s age was the last
condition met. The indexed machine's 171 s came on a day its Outlook started slowly - preparing that
run took 3 min 22 s - and a settle wait ends at the two-minute limit whatever the add-in has done.
The add-in installed on both machines predates the start-up stamp, so every wait read its tuning
reconcile instead (`Services/AddInServerContract.cs`).
