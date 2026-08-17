# The COM host: why the MCP server runs Outlook in a child process

## The incident this exists to prevent

On 2026-08-15 two `search` calls from a Claude Code session each hung for the full
1800-second client idle timeout and were aborted by the client. No error, no partial
result, no server-side timeout — silence.

Reproduced against the installed 3.0.1.321 build with a bounded stdio probe:

| call | result |
| --- | --- |
| `initialize` + `tools/list` | OK, 19 ms / 8 ms |
| `outlook_health` | **no response, 120 s** |
| `search` | **no response, 300 s** |
| `list_signatures` | **OK, 0.1 s** |

`list_signatures` answering normally is the important one: the process, the stdio
transport, JSON-RPC framing and MCP dispatch were all healthy. Only the tools that
touch the COM gateway were dead.

Managed stacks from two wedged processes — one caught wedged *in the wild*, mid-`search`:

```
# in-the-wild process, search
IDispatchInvoke  <-  ComposeSurface.TryPinProcess  <-  EnsureComposeSurfacePin  <-  Connect
    <- PumpedStaRunner.DrainWorkQueue <- PumpLoop

# probe process, outlook_health
RuntimeTypeHandle.AllocateComObject  <-  Activator.CreateInstance("Outlook.Application")
    <-  OutlookComSession.Connect
```

Both inside `OutlookComSession.Connect`, before any search logic ran — which is why
`outlook_health`, the natural diagnostic, hung identically.

The trigger was environmental: Outlook had stopped servicing COM. Verified independently
of this codebase — a plain PowerShell `New-Object -ComObject Outlook.Application` blocked
for over 60 s, with **no `#32770` dialog present** (so not a modal prompt) and Outlook's
UI thread failing `SendMessageTimeout(WM_NULL, SMTO_ABORTIFHUNG)` in 0–6 ms.

A control run against a healthy Outlook returned `outlook_health` in 1.0 s and `search`
in 6.2 s. So the hang is **conditional on Outlook's state** — but the defect is ours: we
turned a bad Outlook state into permanent, product-wide silence.

## The load-bearing constraint

**A blocked outbound COM call cannot be cancelled.** There is no timeout, no token, no
abort. Worse, the RCWs cannot be released either — `Marshal.ReleaseComObject` marshals
into the same wedged apartment and blocks in turn.

So an in-process timeout can only make the *caller* give up. The STA thread stays blocked
forever holding Outlook references. Every such timeout permanently leaks a thread and a
COM reference; after a few, the process is poisoned and must be recycled anyway.

**The only way to reclaim a wedged COM call is to end the process that made it.** That
single fact is the whole architecture.

## Shape

| Process | Owns | Can block on Outlook? |
| --- | --- | --- |
| `OutlookAI.McpServer.exe` (parent) | stdio/JSON-RPC, Windows Search index queries, result merging, supervision, signature tools | **Never** |
| `OutlookAI.ComHost.exe` (child) | the pumped STA thread, `Outlook.Application`, the pin Explorer, every late-bound call | Yes — and that is now safe |

The split follows a seam that already existed: the index tier queries Windows Search, not
Outlook, so the parent can compute index results with no COM at all.

Communication is a length-prefixed JSON frame over a named pipe. Length-prefixed rather
than newline-delimited because payloads carry mail bodies full of newlines, and a framing
desync would reproduce this very class of silent hang one layer down.

## Why the child has no internal timeouts

Deliberate. `PumpedStaRunner.Run` still blocks indefinitely on its work item, and that is
now correct: the child is not trying to survive a wedge, it is trying to *be disposable*
during one. Bounding the wait inside the child would recreate the leaked-thread problem
the process boundary exists to solve.

The bound lives in the parent, and its enforcement is `Process.Kill`.

## Supervision

`ComHostPolicy` holds the decisions as pure, total functions:

- `DecideDispatch` — dispatch / start-then-dispatch / refuse-backoff / refuse-unavailable
- `DecideInFlight` — keep-waiting / timeout-kill-child / fail-child-died / abandon-client-cancelled

They are pure so T1 can pin every branch with a synthetic clock, the same idiom
`SweepWalkBoundsTests` already uses for the sweep bounds. Otherwise the behaviour that
matters most would only be observable by wedging a real Outlook — neither reproducible
nor CI-safe.

Deadlines: 120 s for an ordinary operation, 90 s for connect (which may cold-start
Outlook), **5 s for the health probe**. Health must answer while Outlook is wedged,
because that is exactly when it is asked; exceeding its probe budget degrades the report
rather than failing it.

Failure of a request is not failure of the server: the child is killed, the caller gets a
structured `Timeout` error, and the next call spawns a fresh child. Repeated start
failures trigger a backoff so a broken machine does not become a spawn loop.

## Remembering, not just bounding

Bounding each call is necessary but not sufficient. Measured against a genuinely wedged
Outlook on 2026-08-16, with per-call bounds already in place:

| tool | latency |
| --- | --- |
| `list_signatures` (no COM) | 0.0 s |
| `outlook_health` | 5.8 s |
| `list_accounts` | **120.2 s** |
| `search` | **120.2 s** |

Every request independently paid its full budget and spawned a child to rediscover what
the previous one had already established. Fifteen times better than 1800 s, and still bad:
the tenth search in a row cost two minutes to learn nothing new.

So the supervisor remembers. After two consecutive timeouts it refuses COM requests
immediately for 30 s, then allows one **cheap** liveness probe — `GetProfileName` on the
5 s health budget, not the caller's real request, because re-probing with a full operation
would make every cooldown expiry cost two minutes again. Any success closes it, so a user
who restarts Outlook is picked up automatically.

The freshness sweep also got its own, much shorter budget (30 s rather than 120 s). It is
an *enhancement*: search already holds its indexed answer before the sweep runs, and the
tool's own description promises "sub-second and cheap". Healthy sweeps measure 0.5–6 s.

Same machine, same wedge, after:

| tool | before | after |
| --- | --- | --- |
| `outlook_health` | 5.8 s | 6.1 s, then **0.1 s** |
| `search` | 120.2 s | 30.3 s, then **0.1 s** (3 indexed hits, `sweep.performed: false`) |
| `list_accounts` | 120.2 s | **0.0 s** |
| `list_folders` | — | **0.0 s** |

Verified not to latch: after the cooldown a 5.2 s probe ran, failed, and re-opened —
rather than either giving up permanently or paying a full budget again.

## Asking Windows first

Bounding calls, then remembering across them, still left every *first* encounter with an
unavailable Outlook paying real time. The cheapest fix turned out to need no COM at all.

Windows already tracks whether a window's owning thread is servicing its message queue,
and answers in microseconds through `IsHungAppWindow`. `OutlookLiveness` reads it and
classifies Outlook as **not running / starting / responsive / not responding**, before any
COM is attempted. Three things fall out of that one free check:

- **Hung** - refuse instantly. No child spawned, no budget spent, no kill needed.
- **Starting** - return retry guidance (`OutlookStarting`, `retryAfterSeconds`) instead of
  blocking a caller for a cold start that can take a minute and a half. Outlook is warmed
  up in the background so the retry lands on a ready session.
- **Not running** - start Outlook at most once per 20 s cooldown. This is the **anti-churn
  guard**, and it exists because of the root cause below.

Measured against the same wedged Outlook, first call, no warm-up:

| tool | originally | per-call bounds | + breaker | + liveness |
| --- | --- | --- | --- | --- |
| `outlook_health` | 126 s | 5.8 s | 5.8 s | **0.6 s** |
| `search` | 120 s | 120 s | 30 s | **0.2 s** |
| `list_accounts` | 120 s | 120 s | 120 s | **0.04 s** |

Over 70 sequential and 12 concurrent calls against that wedge, **no COM host was spawned
at all** and Outlook was never restarted.

## Root cause: we were wedging Outlook ourselves

Analysis of a live wedge on 2026-08-16 found the process was `OUTLOOK.EXE -Embedding`
with parent `svchost.exe` - i.e. **COM-activated by us**, not launched by the user. It had
burned 1.34 s of CPU, made no network calls at all, and stopped immediately after loading
add-ins; the event log showed **two Outlook starts 39 seconds apart**, the second never
completing startup.

The suspected mechanism is our own doing: a timeout kills the COM host, that host was the
last COM client of an Outlook *we* had started headlessly, so Outlook begins shutting down
- and the very next request activates it again while the previous instance is still
exiting. Starting Outlook on top of an exiting one is a well-known way to get a
half-initialised, wedged process.

Hence the cooldown. The breaker helps here too, by removing the rapid kill/respawn cycle
that produced the churn. This is inference from strong evidence, not proof: confirming it
needs a native stack, and no debugger is installed on that machine.

## Degraded results are stated, not implied

When the live check cannot run, `search` still **succeeds** and returns its indexed answer
- discarding results we already hold would be the worse failure, and `isError` would
invite clients to throw them away. It carries `degraded: true` and
`freshness: "index-only"` alongside prose opening with `INCOMPLETE RESULTS - TELL THE
USER`. `degraded: true` means "not fully fresh" and now has two shapes: the sweep did not
run at all (`freshness: "index-only"`, above), or it ran and covered only part of its
scope (`freshness: "partial"`, with `sweep.coverageGaps` naming which of the seven
coverage holes fired - failed folders, the per-folder item cap, the folder cap, the time
budget, the depth limit, skipped folders, or a sweep that swept nothing at all). Read the
flag, not either value: a partial sweep used to report `freshness: "live"` with no
degradation, so a caller reading fields rather than prose was told a partial answer was
complete. A result that looks complete and quietly is not is the one failure mode here that
misleads a reader rather than merely inconveniencing them.

## Lifetime

Two independent guards, because the failure they prevent — an orphaned process holding
Outlook COM — is the one that started all this. On 2026-08-15 this machine had **18
orphaned `OutlookAI.McpServer` processes**, one of them wedged.

1. A Windows **Job Object** with `KILL_ON_JOB_CLOSE`. The kernel terminates the child when
   the parent's handle closes, including on a hard kill where no cleanup code runs.
2. The child watches the parent PID and exits if it disappears, covering the case where
   the job could not be created or the handle outlived the process.

## Error contract

Failures now set MCP `isError: true` and carry a machine-readable body, rather than the
previous convention of a protocol-level *success* whose text happened to contain an
`error` object.

One trap, worth stating because it is invisible and would silently reintroduce the
original symptom: the SDK rethrows `OperationCanceledException` only when the **request**
token fired. A server-side deadline is a different token, so it falls through to the
SDK's generic handler and the client receives the message-redacted
`"An error occurred invoking 'search'."` — silence-adjacent, which is exactly what this
work exists to eliminate. `GuardAsync` therefore catches `OperationCanceledException`
itself and distinguishes deadline expiry from client cancellation explicitly.

## What this does not fix

It does not make Outlook responsive. If Outlook is wedged, fresh mail is still
unreachable. What changes is that the failure is **fast, structured, per-request and
recoverable**, and that index results — already computed before the sweep, and previously
thrown away by the hang — are returned with a `coverage` block saying what was and was not
covered.
