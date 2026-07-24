# OutlookAI MCP Server — Developer Documentation

The `McpServer/` folder contains the v3 layer of OutlookAI: a local **MCP (Model Context Protocol) server** that lets AI agents (Claude Code or any MCP client) search, read, show, and draft mail in classic Outlook — fast, offline, and without any cloud mail API.

This document is for contributors. User-facing documentation is in the [repository README](../README.md#mcp-server-mail-search-reading-and-drafting-for-ai-agents).

## Architecture

Two independent data paths, one process:

```
MCP client (stdio JSON-RPC)
   │
   ▼
OutlookAI.McpServer.exe          net10.0-windows x64 console host
   ├─ IndexSearch  ── OLE DB `Search.CollatorDSO` → Windows Search SystemIndex
   │                  (WS-SQL; ~10–100 ms per query; works with Outlook closed)
   ├─ EntryIdCodec ── decodes index item URLs (store UID, folder segments, /at= attachment parents)
   ├─ HitLocator   ── maps an index hit to a real Outlook item (see "load-bearing facts")
   ├─ ComGateway   ── ONE dedicated pumped STA thread owns ALL Outlook COM (late-bound dynamic)
   │                  └──► OUTLOOK.EXE  (read, show-me, drafts, send, freshness sweeps)
   └─ Audit log    ── structured line per write op, %LOCALAPPDATA%\OutlookAI\audit.log
```

Three projects:

| Project | Target | Contents |
|---|---|---|
| `OutlookAI.Core` | **net48 + net10.0-windows** | ALL domain logic: index search, EntryID codec, COM gateway/session, mail service, audit, health. Host-neutral — no MCP types, no console assumptions. The net48 target exists so a future in-Outlook (VSTO add-in) host can reference Core without a rewrite; it is CI-gated from day one. |
| `OutlookAI.McpServer` | net10.0-windows x64 | Thin host: [`ModelContextProtocol`](https://www.nuget.org/packages/ModelContextProtocol) SDK (pinned 1.4.1) + stdio wiring + the tool attribute surface. stdout carries JSON-RPC; all logging goes to stderr. |
| `OutlookAI.McpServer.Tests` | net10.0-windows | Three test tiers (below). |

Design rules (binding — see the code comments for the reasoning):

- **One STA thread owns all Outlook COM.** It runs a real message pump and a COM `IMessageFilter` that retries `RPC_E_CALL_REJECTED` (a busy Outlook rejects automation calls). Index queries run on MTA threadpool threads and are parallel-safe.
- **Core stays host-neutral** and keeps shared state under `%LOCALAPPDATA%\OutlookAI\` (audit log, attachment scratch dir) so future hosts can share it.
- **The server never kills or restarts Outlook.** It may *start* Outlook when a COM-requiring tool needs it — unless the add-in installer's `OutlookAISetup` mutex is held (returns a clear retry-later error instead). Always run non-elevated: an elevation mismatch breaks out-of-process COM attach.
- **Compact payloads everywhere.** Every list has a cap, every cap has a truncation/has-more flag, and the cap values are public constants on `MailService` pinned by T1 tests. Search over-fetches by one row so `truncated=true` means more matches definitely exist.
- **No destructive surface.** The server has zero delete/move/modify tools for existing mail. The only writes: draft creation, attachment saves to the scratch dir, audit lines. `send` is a deliberate two-step flow (single-use content-bound confirm token, ~2 min TTL) with the sending identity hard-verified before transport.

## Tools

L1 search: `search` (always fresh; `exhaustive` flag), `thread`, `index_status`, `list_accounts`, `list_folders` · L2 read: `read`, `save_attachment` · L3 show-me: `open_in_outlook`, `goto_folder`, `show_search_results` · L4 drafts: `new_draft`, `reply_draft`, `replyall_draft`, `forward_draft` · L5: `send` (high-friction, two-step) · diagnostics: `health`, `echo`.

Search behavior (there is no mode parameter): every search merges the index results with a bounded COM sweep of mail newer than the index frontier, so just-arrived mail is always found. The sweep may autostart Outlook headless; its result is cached for ~10 s so rapid-fire iterative searches pay it once and then run at index speed. When the sweep cannot run (installer mutex held, COM attach failed, Outlook cannot start) the search still succeeds with index-only results plus a freshness warning in `advice` — a search never fails because of the sweep. `exhaustive: true` instead runs a folder/date-bounded COM scan that bypasses the index entirely (requires `store` plus `folder` and/or `after`; whole-word matching on subject+body) for when the index is stale/broken or correctness beats speed.

## Build and test

Requires the .NET 10 SDK. **Always build by explicit csproj path — never via `OutlookAI.slnx`** (the solution only carries the VSTO add-in, which needs MSBuild + VSTO tooling):

```
dotnet build McpServer/OutlookAI.Core/OutlookAI.Core.csproj -c Release
dotnet build McpServer/OutlookAI.McpServer/OutlookAI.McpServer.csproj -c Release
dotnet test  McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj --filter "Category!=Live"   # CI-safe tier
dotnet test  McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj                             # full suite (dev machine)
```

`McpServer/Directory.Build.props` enforces `TreatWarningsAsErrors`, nullable, and latest C# for all three projects. Building Core standalone gates **both** targets; a net48 break fails the build. CI is `.github/workflows/mcpserver.yml` (windows runner, dotnet only, live tier excluded).

### Test tiers

| Tier | What | Where it runs |
|---|---|---|
| **T1** unit | Pure logic: WS-SQL builder shapes (anti-patterns as negative tests), EntryID codec against recorded hex fixtures, payload caps/truncation, token store, validation | Anywhere, incl. CI |
| **T2** live integration (`[Trait("Category", "Live")]`) | Against the real SystemIndex and a real Outlook profile | Dev machine only |
| **T3** MCP conformance | Spawns the built exe and speaks real JSON-RPC over stdio (`initialize`, `tools/list`, `tools/call`); CI-safe subset + live subset | Both |

Live-test conventions:

- Machine-local identifiers (store display names, a probe term) live in the **gitignored** `OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json` (`testHubStoreDisplayName`, `expectedStoreDisplayNames[]`, `expectedDelegateStoreDisplayNames[]`, `probeTerm`). The repo is public: account identifiers, live fixtures, audit logs, and screenshots are never committed.
- All test writes target one designated low-value **test-hub mailbox** (configured in the settings file); other stores are read-only for tests except property-asserted, never-displayed identity checks.
- Every test artifact carries the subject tag `[OutlookAI-McpTest]`; deletion requires BOTH the tag and a this-run EntryID/marker match. Cleanup loops until the zero count is **stable** (self-send copies can materialize after a first pass returns zero) and the suite ends with a cross-account sweep asserting zero tagged items.
- Tests may start Outlook (headless COM start) and drive its UI, but never kill or restart it. Test output never prints subjects/bodies from business stores — counts, ids, hashes, booleans only.
- Tests run sequentially (`xunit.runner.json` disables collection parallelism) so live timing and Outlook interactions stay deterministic. The T3 client locates the built server exe via an `AssemblyMetadata` attribute baked into the tests csproj.

## Registration

Register the built server user-globally for Claude Code (any MCP client works equivalently):

```
claude mcp add --scope user outlookai <absolute-path-to>\McpServer\OutlookAI.McpServer\bin\Release\net10.0-windows\OutlookAI.McpServer.exe
```

One server process is spawned per agent session and exits on stdin EOF. Hit ids (`h1`, `h2`, …) and send-confirmation tokens are per-process; a restart invalidates them (the error text tells the agent to re-search). During development/soak the registration points at build output; installer/updater integration is a later productization step.

## Load-bearing facts (learned in Phases 1–6; the code comments reference these)

1. **Decoded 24-byte index EntryIDs are NOT openable on cached-Exchange stores** — `GetItemFromID` returns 0x80040107 for every store of this shape. The index URL's value is the store UID + folder segments + timestamp; `HitLocator` maps a hit to the real (~70-byte) EntryID by walking to the folder and probing via `Folder.GetTable` with a DASL subject + received-time window (5 s tolerance for mail, 120 s for attachment rows), tiering down to `Items.Restrict`. Located EntryIDs are cached per process; reads average well under 100 ms.
2. **Index `System.Message.DateReceived` is UTC**; COM `ReceivedTime` is local wall time. Convert deliberately, never compare raw.
3. **Delegate-store hits are indexed under the OWNER's `/1/<delegate name>` subtree** and the URL short id carries the OWNER's store UID — route delegate hits by folder segments, never by UID. Index URL folder segments use localized display names and match COM folder names exactly.
4. **Only per-column `CONTAINS` is index-backed for sender/recipient filters** — `=`/`LIKE` on `System.Message.FromAddress`/`ToAddress` are multi-second property scans. `SCOPE` needs the exact `($hash)` store segment; discover small stores via `CONTAINS(To/From, '"address"')`, not big URL samples.
5. **Late-bound dynamic COM maps failure HRESULTs to plain .NET exceptions** (E_INVALIDARG → `ArgumentException`, binder failures → `RuntimeBinderException`). Catching only `COMException` misses real failures — use `OutlookComSession.IsComCallFailure` for every optional COM path.
6. **A busy Outlook rejects incoming automation calls (`RPC_E_CALL_REJECTED`)** — e.g. right after `Explorer.Search`. The pumped STA registers an `IMessageFilter` that retries every 250 ms for up to 30 s; this protects every gateway call.
7. **`MailItem.SendUsingAccount` is a PROPERTYPUTREF property: a dynamic assignment silently no-ops** and the DEFAULT account would send. Set it via `InvokeMember(..., BindingFlags.PutRefDispProperty, ...)` and verify by getter readback **in the same session** before any `Send()`. Cross-session readback of saved drafts returns null — verify at send time, not after reopen.
8. **Per-account draft filing:** a plain `CreateItem` draft saves into the DEFAULT store's Drafts — create per-account drafts via the target store's `Drafts.Items.Add(0)`. `Reply()`/`ReplyAll()`/`Forward()` + `Save()` land in the source store's Drafts. Threading only ever via those three methods.
9. **The signature `GetInspector` touch leaves a hidden Inspector** inside Outlook — close it (`Close(olDiscard)`) after `Save()` when not displaying.
10. **Sent items get a NEW EntryID** — capture outcome snapshots before `Send()`; verify arrival via the Inbox side (the Sent-copy lags sweep windows).
11. **The audit log is load-bearing:** draft/save/send operations THROW when their audit line cannot be written (the EntryID is preserved in the error text). `send` refuses fail-closed on any token/content/identity mismatch, and recomputes the content hash inside the STA right before sending.
12. **Deleted items keep index rows** (with `IncludeDeletedItems=1`) — live samplers must exclude the test tag; agents should expect the occasional just-deleted hit to fail location with a re-run-search error.

## Health & diagnostics

`health` reports everything the server depends on (Outlook process/version, store reachability, index freshness, WSearch service state, audit writability, tuning state, installer mutex) without starting Outlook. `index_status` is the lighter freshness-only self-report. The audit log at `%LOCALAPPDATA%\OutlookAI\audit.log` records every write operation with timestamps and EntryIDs.

The tuning block of `health` also carries `uiSearchBackend` — the EFFECTIVE Outlook UI search backend read from the live registry (`DisableServerAssistedSearch`, policy hive authoritative over the user hive; absent/0 = Outlook's server-assisted default). `"local"` means the Outlook search box queries the same Windows Search index the agent's `search` uses; `"server-assisted"` means UI results are server-capped and differently ranked, so what the user sees can diverge from what the agent finds — `show_search_results` then appends a strong advice note recommending the Search tuning group be re-enabled. The registry value stays deliberately user-togglable: disabling it is the sanctioned mitigation for the documented replies-stick-in-Outbox side effect of local-search mode.
