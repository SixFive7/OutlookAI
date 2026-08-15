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
