# COM reference-release audit

Read-only audit of the shipped code, 2026-09-15. No file changed but this one. No Outlook,
mailbox, VM or live test touched.

Every line reference is to the working tree at the time of the audit. Claims are marked
**READ** (stated by the source) or **INFERRED** (my reasoning from the source plus Windows
semantics) throughout.

---

## Direct answer

**Yes. The shipped code has the exposure, and it is the architecture's deliberate central
trade-off rather than an oversight.**

The COM host child is designed to be terminated with `TerminateProcess` while it holds an
`Application`, a `NameSpace`, a non-displayed pin `Explorer`, an **advised event-sink
registration inside Outlook**, and whatever the in-flight call had bound. On that path
nothing releases any of them. The code says so in its own words
(`OutlookComSession.cs:5730`: *"TRUE ONLY IF THIS PROCESS SURVIVES. A killed COM host runs
no finally"*).

What is *not* wrong: the per-call COM discipline. 249 `Release(...)` sites in
`OutlookComSession.cs`, essentially all in `finally`, with no leak found on an exception
path. If today's incident is caused by a process dying while holding references, the defect
is not in how this code releases references — it is in the premise that dying is a safe way
to stop holding them.

That premise is stated once, in a comment, with no measurement behind it in this repository:

> `ComHostSupervisor.cs:842` — *"`TerminateProcess` destroys the whole address space at
> once, and Windows tears down the LRPC endpoints and Outlook's references to the dead
> client with it."*

Every neighbouring claim in this codebase carries a date and a number. This one does not. It
is the single load-bearing assumption of the whole COM-host split, and today's incident is
evidence against it.

---

## 1. Where the project binds Outlook COM objects

### Shipped product — exactly one activation site

`McpServer/OutlookAI.Core/Com/OutlookComSession.cs:334-337` (READ)

    Type.GetTypeFromProgID("Outlook.Application") -> Activator.CreateInstance(...)
       -> app.GetNamespace("MAPI") -> ns.Logon(...)

Everything the product does against Outlook funnels through `OutlookComSession`. Confirmed
by grepping the whole repo for `GetTypeFromProgID`, `GetActiveObject`, `BindToMoniker`,
`CreateObject` and `Outlook.Application`: the only other Outlook activations are in the test
mailer and the remediation console (below), and the only other COM activation of any kind is
`ADODB.Connection` / `ADODB.Recordset` in `IndexSearch/IndexClients.cs:120-122` (Windows
Search, not Outlook; released in `finally` at `:159-163`).

**Long-lived references a session holds** (`OutlookComSession.cs:121-133`, READ):

| field | what it is | released where |
| --- | --- | --- |
| `_application` | `Outlook.Application` | `Dispose` only |
| `_namespace` | `NameSpace` (MAPI) | `Dispose` only |
| `_composeSurfacePin` | non-displayed `Explorer` from `Explorers.Add(folder, 0)` | `Dispose` only |
| `_quitSinkRegistration` | `IConnectionPoint` + an **advise cookie held by Outlook** | `Dispose` only |
| `_quitSink` | our managed sink object, marshalled into Outlook | (implicit, with the above) |

Everything else — `Store`, `Folder`, `Items`, `Table`, `PropertyAccessor`, `MailItem`,
`Explorer`, `Inspector`, `Account`, `Recipients`, `Document` — is transient: bound inside one
`_runner.Run(() => { ... })` work item and released in that item's `finally`.

**Threading shape** (READ): one dedicated STA thread per session with a real Win32 message
pump (`PumpedStaRunner`), plus a `CoRegisterMessageFilter` retry filter that turns
`RPC_E_CALL_REJECTED` into retries for up to 30 s. All Outlook objects are created and used
only on that thread.

**Process shape** (READ, `McpServer/Docs/com-host.md`): the parent `OutlookAI.McpServer.exe` holds no
Outlook COM at all; the child `OutlookAI.ComHost.exe` holds all of it. Verified against the
code — the parent's session is a `DispatchProxy` (`RemoteSessionProxy`) that serialises calls
over a named pipe, and the signature tools (`SignatureManager`, `SignatureCatalog`,
`OutlookProfileRegistry`) contain no `Marshal.` and no COM at all. They are file and registry
work, which is why `list_signatures` answered in 0.1 s while everything else was wedged.

### VSTO add-in — a different exposure class

`ThisAddIn.Designer.cs:26,42` gets the `Application` as a VSTO host item; `ThisAddIn.cs`
keeps `Inspectors`, `Explorers` and a list of hooked `Explorer`s, and releases via
`ThisAddIn.ReleaseCom` (`:44-51`). This code runs **inside OUTLOOK.EXE**: same apartment, no
LRPC, no cross-process reference. It cannot produce the incident (INFERRED, from the hosting
model). `AddInAutomation.cs` exposes an `IDispatch` surface *into* the add-in via
`COMAddIn.Object` — the reverse direction, also in-process.

### Remediation tools — explicitly not product

`McpServer/OutlookAI.RemediationTools/` — the csproj says *"One-shot operator console… NOT
part of the product surface"* (READ). `ComMailbox.cs:749-752` and
`ComCorpusMailbox.cs:1586-1589` each activate `Outlook.Application` on their own STA thread.
Their release discipline is the same `finally`-per-fetch pattern and looks sound
(`ComMailbox.cs:99-105` is representative: `store`, `stores`, `ns`, `app` all released in one
`finally`). `ComMailbox.RunSta` has a 3-minute `Thread.Join` timeout that **abandons** the STA
thread on expiry, which would leave a live process holding references — a real but narrow
version of the same exposure, in a tool run by hand.

### Tests

`T2/LiveOutlookTestMailer.cs:1801-1804` activates Outlook directly for the live tier. In
scope only because the live tier runs on the maintainer's machine.

---

## 2. Is release deterministic?

**Yes, on every path the process survives.** (READ)

- Mechanism is explicit `Marshal.ReleaseComObject`, funnelled through one helper,
  `OutlookComSession.Release` (`:10870`), guarded by `Marshal.IsComObject`.
- **No `FinalReleaseComObject` anywhere in the repo.** No RCW wrapper type, no `using`-based
  COM handle, no `SafeHandle`-style ownership. Release is by hand and by convention.
- Structural check: 249 `Release(` call sites in `OutlookComSession.cs`; 270 `try` blocks;
  171 `finally` blocks. 230 of the 249 sit within 12 lines after a `finally`. I inspected
  every one of the ~19 that do not: they are release-as-you-go inside walk loops
  (`:756, 3357, 3381, 4562, 4576, 4654, 5530-5556, 7517, 10757, 10766`), one hand-off release
  after a successful `MailItem.Move` where the enclosing `finally` still releases `parent`
  (`:7182`), and `Dispose` itself (`:10955-10958`). **I found no site where a reference would
  survive an exception.**
- `Dispose` (`:10893`) releases in a deliberate order — unadvise the sink, close-or-release
  the pin, release namespace, release application — inside a `_runner.Run(...)` on the STA,
  then stops the runner, then `GC.Collect()` + `GC.WaitForPendingFinalizers()` as a backstop.
- The other COM surfaces are equally disciplined: `IndexClients.TryCloseAndRelease` in a
  `finally`; `OutlookQuitSink.TryAdvise`'s failure path releases the connection point;
  `ComposeSurface.TryPinProcess` releases `folder` and `explorers` in a `finally`;
  `ThisAddIn.ComEqual` balances `GetIUnknownForObject` with `Marshal.Release` in a `finally`.

**This is a clean result and should be read as one.** There is no "the code forgets to
release" finding here. The finding is entirely about paths on which no code runs at all.

---

## 3. Abnormal exit — the question that matters

### Stated unambiguously

**On the deadline path the child is terminated with `TerminateProcess` while holding live
Outlook references, and nothing — in that process or in the parent — releases them. There is
no cleanup on that path. It is intentional.** (READ)

Evidence, in order:

1. `ComHostSupervisor.ArmDeadline` (`:420`) → on expiry calls `KillChild` (`:870`) →
   `child.Kill(entireProcessTree: true)` (`:895`). No stop frame, no signal, no grace. The
   design explains why a polite request is impossible: `ComHostServer.ServeAsync` calls
   `Invoke` synchronously inside its read loop, so a wedged child is not reading the pipe at
   all (READ, `:862-867` and `McpServer/Docs/com-host.md`).
2. `KillChild` then calls `TearDownChild`, whose 250 ms `CleanExitGraceMilliseconds` (`:830`)
   is by its own comment *"free on the deadline path, because `KillChild` has already
   terminated the process and `HasExited` is true by the time the wait is reached"* (READ).
   The grace does not apply to the kill it was added beside.
3. The child's **only** release path is a `using` declaration — `ComHost/Program.cs:56`,
   `using ComGateway gateway = new ComGateway(allowStartingOutlook)`. `TerminateProcess` runs
   no `finally`, so this never executes (READ + INFERRED from Windows semantics).
4. The child holds no other cleanup hook: no `AppDomain.CurrentDomain.ProcessExit`, no
   console control handler, no finaliser that releases. Grepped; there are none anywhere in
   `OutlookAI.ComHost`, `OutlookAI.Core` or `OutlookAI.McpServer`.

### The clean-exit path exists and, in the shipped server, is not reachable with a live session

This is the sharper half of the finding and I did not expect it.

`CleanExitGraceMilliseconds` was added on 2026-08-19 specifically so the child's own COM
release could run (READ, `:815-829`: *"an orderly exit runs its finally blocks and releases
its COM references, which a kill skips"*). I enumerated every `TearDownChild` call site
(`:529, 614, 634, 905, 1091`):

| site | context | does the child hold an Outlook session? |
| --- | --- | --- |
| `:529` | top of `StartChildAsync`, clearing a previous non-Ready child | No — the child connects to Outlook lazily on first use, and a non-Ready child has served nothing |
| `:614` | pipe handshake failed | No |
| `:634` | child connected but never reported ready | No |
| `:905` | called *from* `KillChild` | Already terminated |
| `:1091` | `ComHostSupervisor.Dispose` | **Never called — see below** |

`ComHostSupervisor.Dispose` runs only from `RemoteComGateway.Dispose`. The gateway is a
`static Lazy<RemoteComGateway>` in `OutlookTools.cs:30`, is not registered in DI, and
`McpServer/OutlookAI.McpServer/Program.cs` is sixteen lines with no lifetime hook, no `IHostApplicationLifetime`
subscription and no `ProcessExit` handler. **Nothing in the shipped MCP server ever disposes
its gateway.** (READ)

So in production the child's orderly COM release is reachable only through the child's own
`WatchParent` (`Program.cs:97-133`): parent exits → `Process.Exited` → cancel → `ServeAsync`
returns → `Main` returns → `using` disposes → refs released.

### …and that path races a kernel kill it will normally lose

`ChildJobObject` is `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE` (`ChildJobObject.cs:57`), assigned
immediately after `Process.Start` (`ComHostSupervisor.cs:572-573`). When the parent dies by
any means its handles close and the kernel terminates every process in the job *immediately*.
The child's managed route to the same outcome is a thread-pool `Exited` callback, a
cancellation, a loop exit, an async unwind and a cross-thread STA dispatch.

**INFERRED, not measured: the job object wins, so at MCP-server shutdown the child is
terminated rather than unwound, and its references are not released.** The design is explicit
that this was the intended trade — orphan prevention over cleanup (*"a killed or crashed
parent cannot run cleanup code, so process-tree lifetime has to be enforced by the kernel"*,
READ) — but the consequence for COM release is not stated anywhere.

*Settles it:* have the child write a line naming its exit route (EOF / parent-watch / never)
and read it after a normal server shutdown. `TODO.md:229` already asks for exactly this test
and records that it does not exist.

### Does that reproduce today's incident?

I cannot answer that from source, and I want to be precise about why.

- The shipped kill path leaves **the same reference set the probes left**: a dead process
  that last held `Application` and `NameSpace` against a shared Outlook. So if a dead client
  holding those is sufficient to block later clients, the shipped code reproduces it on every
  deadline kill. (INFERRED)
- The repo contains one measurement pointing the *other* way, and it is dated:
  `ComGateway.cs:44-50` and `OutlookComSession.cs:10887-10890` both record, from a 2026-07-23
  probe, that a Quit-parked Outlook *"unsticks ~6 s after our refs release (e.g. **the server
  process ends**)"*. That says process death did release the references, at least for that
  failure shape, on that build. (READ)
- Those two cannot both be the whole truth. Either the 2026-07-23 measurement does not
  generalise beyond the Quit-park shape, or something other than plain held references
  produced today's nine minutes. **This tension is the most important thing the audit found**,
  because the entire kill-based architecture rests on the optimistic half of it and only the
  pessimistic half has been observed recently.

Three shipped-code details bear on which it is, all **INFERRED**:

1. **The advised event sink is a reference pointing the wrong way.**
   `OutlookQuitSink.TryAdvise` (`:47-85`) calls `IConnectionPoint.Advise` on **every**
   `Connect`, including in the COM host child. That hands *Outlook* a marshalled pointer to
   an object in our address space. It is the one reference a client can never release from
   its own side, and `Unadvise` runs only in `Dispose`. A killed child therefore leaves
   Outlook holding a sink registration to a dead process, and Outlook firing any
   `Application` event must then marshal to an endpoint that is gone. This is the most
   plausible mechanism I can see for "a process that exits holding references blocks every
   later client".
2. **The pin `Explorer` is leaked into Outlook on every kill.**
   `ComposeSurface.TryPinProcess` (`:316-352`) does `Explorers.Add(folder, 0)` and never
   `Display()`. `OutlookComSession.Dispose` (`:10925-10940`) deliberately **`Close()`s** it
   rather than merely releasing, and says why, from measurement: *"an Explorer released but
   left in Outlook's collection outlives the session that made it and keeps Outlook up for
   good — a later `Application.Quit` (the user choosing Exit) then does NOT terminate the
   process, and repeated sessions accumulate invisible Explorers"* (READ). A kill runs no
   `Close`. So the measured consequence is already written down in this repo, and the kill
   path produces it every time. The pin is created only when `Explorers.Count == 0` — i.e. on
   a headless Outlook, which is exactly the test guest's state.
3. **`Connect` itself calls `GetDefaultFolder`.** `EnsureComposeSurfacePin` →
   `TryPinProcess` → `namespaceObject.GetDefaultFolder(6)` (`ComposeSurface.cs:338`). The
   product's connect path runs the very call that blocked for nine minutes today. That does
   not cause the exposure, but it means `outlook_health` and every first COM call are exposed
   to the same block, behind a 180 s connect deadline.

One honest counterweight: today's probes were PowerShell and almost certainly held no advised
sink and created no pin. If the incident reproduced without either, then a plain held
`Application` / `NameSpace` in a dead process is sufficient on its own — which makes the sink
and the pin *additional* liabilities rather than the required mechanism, and makes the shipped
kill path strictly worse than the probes rather than merely equivalent.

### Other abnormal exits

- `Environment.FailFast` in `ComHostFaultInjection.cs:108` (the `crash:` fault kind) is a
  second no-`finally` exit. Environment-variable gated (`OUTLOOKAI_COMHOST_FAULT`), test only,
  but compiled into the shipped binary with no `#if`.
- `OutlookComSession.Dispose` can block indefinitely on a wedged STA: it calls
  `_runner.Run(...)` (`:10917`), which enqueues onto the pumped thread and waits on a
  `TaskCompletionSource` with **no timeout**. On the supervised path that only means the
  250 ms grace expires and the kill lands. But if `ChildJobObject.CreateOrInert` returned an
  inert job (restricted token, or an incompatible existing job — `:44-84` degrades silently)
  **and** the parent dies, the child's parent-watch unwind blocks forever in `Dispose` and the
  process becomes an orphan holding Outlook COM — the precise leak the architecture exists to
  prevent, reached through its own cleanup path. Narrow, but it is the failure mode that
  produced the 18 orphaned processes on 2026-08-15. (INFERRED)

---

## 4. Do the timeouts protect the right thing?

**Yes — for the shipped server. Not for the live test tier, which says so itself.**

### The mechanism, concretely

A deadline here does not interrupt a COM call. Nothing can; the design states the constraint
plainly and names `Marshal.ReleaseComObject` as blocking too, because it marshals into the
same wedged apartment (READ). What the deadline does is bound a **managed wait in a healthy
process** and then destroy the process that made the call.

Chain, all READ:

1. `RemoteSessionProxy.Invoke` classifies the operation (`ComOperationClasses.ClassOf`),
   computes `ComHostPolicy.DeadlineFor(class, override)` and shrinks it against the remaining
   aggregate (`EffectiveDeadlineMilliseconds`), then blocks on
   `_supervisor.InvokeAsync(...).GetAwaiter().GetResult()`.
2. That blocking wait is on a `TaskCompletionSource` **in the parent**, not on COM.
3. `ArmDeadline` (`:420`) arms a `Task.Delay(...).ContinueWith(...)` on
   `TaskScheduler.Default`, deliberately **not** linked to the caller's token — so it fires
   even after the MCP client has cancelled and stopped listening (`:12-20`).
4. On expiry the order is: record the replacement decision, **complete the caller's TCS with
   `ComHostTimeoutException`**, then kill. The comment at `:436-462` records that this
   ordering was wrong twice, in opposite directions.

So the deadline fires against something that is not blocked, which is exactly why it fires
where the 2026-08-15 one did not. **This is the right shape.** Granularity is per contract
call (one `IOutlookSession` method = one pipe round trip), with a second aggregate budget
across a multi-call lambda, so a `Run(session => ...)` making N calls is bounded as a whole
rather than N times over.

Budgets (`ComOperationBudgets`, `ComHostPolicy`, raised 2026-08-19, READ): 300 s ordinary,
180 s connect, 615 s exhaustive scan, 180 s freshness sweep, 5 s health probe, 30 s handshake
with a 10 s floor, 1 s minimum dispatch floor.

### Where it does not hold

**The in-process gateway cannot bound a call and documents that it cannot.**
`ComGateway.Run(operation, budgetMilliseconds, ...)` wraps the session in
`BudgetedSessionProxy`, which checks the clock **between** contract calls and refuses to
*start* the next one. It has no power over one in flight (`BudgetedSessionProxy.cs:20-27`,
READ: *"It CANNOT bound one call… It CAN bound a SEQUENCE"*). That gateway is what
`MailService.CreateDefault()` builds (`MailService.cs:456-458`), and every T2 live fixture is
built on `CreateDefault()`. So **in the live test tier a single blocked COM call is still
unbounded.** The cost is recorded in `T2/LiveDisconnectRecoveryTests.cs:34-45`: a run on
2026-08-18 went 22.5 minutes, was stopped by hand, skipped fixture teardown, and left 7 tagged
items in a real mailbox.

Overshoot on the shipped path is one call: the aggregate stops the *next* dispatch; the
per-call deadline plus the kill stops the current one.

---

## 5. Existing tests

**Supervision is well pinned. Release is not pinned at all.**

| test | tier | what it pins |
| --- | --- | --- |
| `T3/ComHostSupervisionCiTests.cs` (7 tests) | CI, no Outlook | The whole timeout → kill → respawn → breaker path, via `OUTLOOKAI_COMHOST_FAULT` injected **above** the routing proxy so no session is ever created. Explicitly the regression test for 2026-08-15. |
| `T3/ComHostSupervisionLiveTests.cs` | `Category=Live`, `Requires=OutlookInstance` | `TheWedgedHostProcessIsEnded_NotReused` (restart count rises), `NoComHostSurvivesTheServer` (child pid is dead after the server exits), health reporting with the breaker open |
| `T1/ComHostPolicyTests.cs` | CI | Every policy branch, synthetic clock |
| `T1/InProcessBudgetTests.cs` | CI | `BudgetedSessionProxy` dispatch floor, driven directly |
| `T1/ComHostErrorFidelityTests.cs` | CI | Error survives the process boundary, per operation, enumerated from `IOutlookSession` by reflection |
| `T2/LiveDisconnectRecoveryTests.cs` | Live | **The opposite direction**: Outlook exits, our watcher releases our refs, health re-probes, the gateway re-attaches |

**Nothing anywhere asserts that a killed child leaves Outlook usable for the next client.**
That is the incident's question and it is unpinned. The nearest adjacent hole is already
recorded at `TODO.md:229`: *"Same for `ComHostSupervisor.CleanExitGraceMilliseconds`.
Replacing the `WaitForExit(250)` with a no-op leaves the suite green… proving it needs a child
that logs its own clean exit, which is a T3-shaped test nobody has written."*

**Where such a test would live, and can it run without a mailbox?**

- A *release* test cannot run without Outlook — it is about a real COM server's behaviour. It
  can run without **mail**: connect, kill, then time `GetDefaultFolder` from a fresh client.
  It reads nothing, writes nothing, and needs no profile beyond a working one.
- That is exactly the shape and trait of `T3/ComHostSupervisionLiveTests`
  (`Requires=OutlookInstance`, not `Requires=MailAccount`), whose own header says any Outlook
  profile satisfies it and the dedicated test VM runs it unchanged. That class is where it
  belongs.
- The supervision half — that the kill happens, in the right order, with the right error — is
  already CI-safe and already covered.

---

## What I could not establish

1. **Whether `TerminateProcess` actually causes Outlook to release a dead client's references
   promptly.** Asserted at `ComHostSupervisor.cs:842`; no measurement behind it in this repo;
   contradicted in spirit by today's incident and supported by the 2026-07-23 Quit-park probe.
   *Settles it:* on the test guest, from a freshly restarted Outlook — (a) connect a child and
   confirm `GetDefaultFolder` is fast from a second client; (b) `TerminateProcess` the child;
   (c) time `GetDefaultFolder` from a third client at intervals; (d) repeat with an orderly
   EOF exit as the control. Two processes, no mailbox writes.
2. **Whether the advised quit sink is the mechanism.** Needs the same experiment with the
   advise suppressed; there is no flag for that today, so it needs a build.
3. **Whether the job object or the child's parent-watch wins at server shutdown.** Needs the
   child to record its own exit route — the test `TODO.md:229` already asks for.
4. **Whether `ReleaseComObject` (rather than `FinalReleaseComObject`) balances every fetch.**
   The CLR keeps one RCW per COM identity per context with a reference count, so two fetches
   of the same identity share one RCW and need two releases. Every use in this code is
   fetch-once / release-once, so it balances by construction — but I verified that by reading
   the shape, not by instrumenting the counts, and a nested fetch of an already-held identity
   would not be obvious in review.
5. **Whether today's probes held anything beyond `Application` and `NameSpace`.** I read only
   enough of `.work/wedge-repro.ps1` to confirm it binds `$ol` / `$ns` and calls
   `GetDefaultFolder`; the audit was scoped to product code. This matters, because if the
   incident reproduced with nothing but those two references then the sink and the pin are
   aggravating factors rather than the cause.

---

## Adjacent problems noticed while reading

1. **The advised quit sink is, on the shipped path, pure liability.** Its own doc comment
   records that on this build the Quit event *"does NOT reach out-of-process sinks… Outlook
   parks instead"* and that the process-exit watcher is the load-bearing signal (READ,
   `OutlookQuitSink.cs:15-21`). Its only production consumer forwards an event the parent
   describes as *"Advisory: the parent re-probes anyway, this only makes it prompt"* (READ,
   `ComHost/Program.cs:60-63`). So it buys nothing measured, and it is the one reference whose
   cleanup is Outlook's problem rather than ours after a kill.
2. **Every deadline kill leaks an invisible `Explorer` into the shared Outlook**, with a
   consequence this repo has already measured and written down: the user can no longer exit
   Outlook, and they accumulate. See §3. Worth checking `Explorers.Count` on the test guest.
3. **The MCP server never disposes its gateway**, so the clean-exit grace added on 2026-08-19
   has no reachable production caller that holds a live session. See §3.
4. **`Dispose` closes the pin only when `StartedOutlook` is true** (`:10938`) — a session that
   merely *attached* releases-without-closing even on the orderly path. Deliberate, with a good
   reason in the comment (a fixture that owned the pin and disposed mid-run cost 16 cascading
   `RPC_S_SERVER_UNAVAILABLE` failures), but it means the leak is not unique to the kill path.
5. **`OutlookComSession.Dispose` can block forever on a wedged STA**, which combined with an
   inert job object is a route to the orphan the architecture exists to prevent. See §3.
6. **`ComMailbox.RunSta` abandons its STA thread after 3 minutes** rather than failing the
   process — an operator-tool process left alive holding references.
7. **The live (T2) tier still cannot bound a blocked COM call**, and has already lost a run and
   left artifacts in a real mailbox because of it. Known and recorded; noted here because it is
   the tier most likely to reproduce today's incident by accident.
