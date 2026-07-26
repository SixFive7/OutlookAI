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
- **No destructive surface for mail.** The server has zero delete/move/modify tools for existing mail. The only writes: draft creation, attachment saves to the scratch dir, audit lines, and signature-file management (`manage_signature` — the one deliberately destructive-capable tool, guarded by an ALWAYS-ON pre-modification backup, see below). `send` is a deliberate two-step flow (single-use content-bound confirm token, ~2 min TTL) with the sending identity hard-verified before transport.

## Tools

L1 search: `search` (always fresh; `exhaustive` flag), `thread`, `list_accounts`, `list_folders`, `list_signatures`, `manage_signature` · L2 read: `read` (body windows page via `body_offset`), `save_attachment` · L2.5 organize: `move_mail`, `archive_mail` (same-store, audited, reversible) · L3 show-me: `open_in_outlook`, `goto_folder`, `show_search_results` · L4 drafts: `new_draft`, `reply_draft`, `replyall_draft`, `forward_draft` (each with an optional `signature` override) · L5: `send` (high-friction, two-step) · diagnostics: `outlook_health` (19 tools total).

`manage_signature` creates, updates and deletes signatures by writing the signature file set under `%APPDATA%\Microsoft\Signatures` — always all three renditions (`.htm` UTF-8 + charset meta, `.txt` UTF-16 LE, `.rtf` derived ASCII), because each Outlook mail format reads only its own rendition and silently omits the signature when that file is missing. For create/update the caller supplies `body_text` and/or `body_html`; the missing renditions are derived. ALWAYS-ON safety: before ANY update or delete the signature's full current file set is copied to `%LOCALAPPDATA%\OutlookAI\signature-backups\<utc-timestamp>-<name>\` and the backup path is returned — a failing backup aborts the operation untouched. Deleting a signature also clears per-account default assignments referencing it, and the optional `set_default_for {account, scope}` records the signature as an account default (Outlook picks default changes up at its next start). Roaming-signatures caveat: on Microsoft 365 Apps 2303+ the cloud copy of signatures can overrule or revert local file writes unless `DisableRoamingSignatures=1` is set (`HKCU\Software\Microsoft\Office\16.0\Outlook\Setup`); on Office LTSC (no roaming) the local files are authoritative. Every operation is audit-logged.

`list_folders` always returns the FULL folder tree in a stable traversal order (stores by display name, then depth-first with siblings by name), paged at 1000 folders per call via `offset`/`nextOffset` — real profiles fit in one page. `read` serves long bodies as windows: `body_offset` continues from the cached one-time extraction instead of re-transferring the body from the start. `list_signatures` enumerates `%APPDATA%\Microsoft\Signatures` (names + short excerpts for language detection) plus per-account defaults where the profile registry records them — absent values are reported as unknown, never guessed. The draft tools' `signature` parameter overrides the applied signature by name: the swap runs through the hidden Inspector's WordEditor (`_MailAutoSig` bookmark replace/insert via `InsertFile`, bookmark recreated — the same mechanism as Outlook's own signature switcher and the add-in's proven bookmark dance), preserving threading, the quoted history and the body-above-signature contract; on any failure the draft still stands with the account default and the outcome reports `signatureApplied:false`.

`move_mail` moves 1–50 items (hit ids or EntryIDs) to a store-relative folder path, **same-store only** (v1 restriction: each item moves within the store it lives in; a requested `store` that differs from an item's own store fails per-item — archive semantics are same-store and EntryIDs are store-scoped). `create_folder: true` creates missing target segments as mail folders. Refused targets: Deleted Items and its subtree (moving there is deletion semantics — the server has no delete surface), the Outbox, non-mail folders, and the item's current folder. `archive_mail` moves each item to ITS OWN store's **designated Archive folder** — the folder Outlook's own Archive action (Backspace), mobile swipe-archive and OWA use. Resolution is localization-proof and never guesses by name: primary = `Store.GetDefaultFolder(39)` (an undocumented `OlDefaultFolders` value the live build honors; the resolved folder is verified — same store, mail folder, not a core default folder — before use), fallback = `PR_IPM_ARCHIVE_ENTRYID` (0x35FF0102) on the store object; a store with neither fails per-item and nothing is created. Both tools are content-preserving, per-item audited (`op=move_mail`/`op=archive_mail` with from→to) and reversible: every result carries `fromFolder` plus `oldEntryId`/`newEntryId` — undo = `move_mail` with `newEntryId` and `folder = fromFolder`. **EntryIDs change on ANY move**: the moved hit id keeps resolving within the session (its cache is refreshed to the new EntryID), other stale ids/index rows surface the standard re-run-search error until the index catches up.

Search behavior (there is no mode parameter): every search merges the index results with a bounded COM sweep for mail newer than the index frontier, so just-arrived mail is always found. **The sweep follows the search scope** (soak fix 13): with `folder` set it covers that folder *and its subfolders* (the index tier's `SCOPE=` is recursive, so the sweep matches it; bounded at 40 folders), otherwise it covers the four arrival-path default folders — Inbox, Sent Items, Deleted Items, Junk Email — of the store(s) in scope. Measured on the dev machine across 5 stores: 86 ms for the old Inbox+Sent pair, 135 ms for all four, versus ~10 ms per folder for a full walk of 41-46 folders per store (which is why the default set is fixed, and why mail a rule files into a custom folder needs `store` + `folder` until the index catches up). The `sweep` block of the response reports the scope descriptor, the folder count and — while the list is short — the folder names. The sweep may autostart Outlook headless; its result is cached for ~10 s, keyed on the index frontier **plus the store and folder scope** so a narrow sweep can never answer a broader query. When the sweep cannot run (installer mutex held, COM attach failed, Outlook cannot start, folder not openable) the search still succeeds with index-only results plus a freshness warning in `advice` — a search never fails because of the sweep. `exhaustive: true` instead runs a folder/date-bounded COM scan that bypasses the index entirely (requires `store` plus `folder` and/or `after`; whole-word matching on subject+body, no attachment content; a `folder` bound scans that folder only, without subfolders) for when the index is stale/broken or correctness beats speed.

Term matching (D40, renamed 2026-07-26; cross-column AND added in soak fix 13): `query` terms are whitespace-separated and ANDed, and `search_in` chooses where they must match — `subject_and_body` (default), `subject` or `body`; all three tiers honor it. In the default scope every tier ANDs the terms **across** the parts, one subject-OR-body pair per term (`(CONTAINS(System.Subject, '"a"') OR CONTAINS(System.Search.Contents, '"a"')) AND (CONTAINS(System.Subject, '"b"') OR CONTAINS(System.Search.Contents, '"b"'))`), so a mail carrying one term in the subject and another in the body matches — the in-column shape that shipped until then missed exactly those. Narrowed scopes stay single-column, where an in-column AND is equivalent and cheaper. Measured cost of the pairs on the dev machine (warm best-of-3, agent-sized TOP 26 + ORDER BY): +0-2 ms at 1-3 terms; recall gain on real corpora: `invoice payment` 3817 → 3857 rows, `offerte klant` 1058 → 1093. `System.Search.Contents` is body **plus attachment content**, which is why attachment-content matches only exist in the index tier; the sweep matches `item.Body` and exhaustive matches `urn:schemas:httpmail:textdescription`. Whole-word matching is exact in the index tier; the sweep's re-match is substring by design (deliberate over-match on the freshness window) and exhaustive falls back from `ci_phrasematch` to substring `LIKE` when instant search is off, when `GetTable` rejects the filter, or for prefix (`*`) terms — the runtime surfaces that in `advice`.

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

### Server instructions (MCP `initialize`)

`ServerMetadata.Instructions` is returned as the `instructions` field of the MCP `initialize` result. Claude Code injects it verbatim into every session's context at start (a "# MCP Server Instructions" section) — including sessions where tool search defers the tool schemas to name-only — so the agent knows the mail capability exists before any tool is loaded. Keep it short (it is a passive per-session cost; Claude Code truncates at 2 KB) and keyword-rich (tool search matches against it). T3 pins the exact wire string; T1 pins the length budget, the discovery keywords, and that no test canary ever ships.

## Load-bearing facts (learned in Phases 1–6; the code comments reference these)

1. **Decoded 24-byte index EntryIDs are NOT openable on cached-Exchange stores** — `GetItemFromID` returns 0x80040107 for every store of this shape. The index URL's value is the store UID + folder segments + timestamp; `HitLocator` maps a hit to the real (~70-byte) EntryID by walking to the folder and probing via `Folder.GetTable` with a DASL subject + received-time window (5 s tolerance for mail, 120 s for attachment rows), tiering down to `Items.Restrict`. Located EntryIDs are cached per process; reads average well under 100 ms.
2. **Index `System.Message.DateReceived` is UTC**; COM `ReceivedTime` is local wall time. Convert deliberately, never compare raw.
3. **Delegate-store hits are indexed under the OWNER's `/1/<delegate name>` subtree** and the URL short id carries the OWNER's store UID — route delegate hits by folder segments, never by UID. Index URL folder segments use localized display names and match COM folder names exactly.
4. **Only per-column `CONTAINS` is index-backed for sender/recipient filters** — `=`/`LIKE` on `System.Message.FromAddress`/`ToAddress` are multi-second property scans. `SCOPE` needs the exact `($hash)` store segment; discover small stores via `CONTAINS(To/From, '"address"')`, not big URL samples.
5. **Late-bound dynamic COM maps failure HRESULTs to plain .NET exceptions** (E_INVALIDARG → `ArgumentException`, binder failures → `RuntimeBinderException`). Catching only `COMException` misses real failures — use `OutlookComSession.IsComCallFailure` for every optional COM path.
6. **A busy Outlook rejects incoming automation calls (`RPC_E_CALL_REJECTED`)** — e.g. right after `Explorer.Search`. The pumped STA registers an `IMessageFilter` that retries every 250 ms for up to 30 s; this protects every gateway call.
7. **`MailItem.SendUsingAccount` is a PROPERTYPUTREF property: a dynamic assignment silently no-ops** and the DEFAULT account would send. Set it via `InvokeMember(..., BindingFlags.PutRefDispProperty, ...)` and verify by getter readback **in the same session** before any `Send()`. Cross-session readback of saved drafts returns null — verify at send time, not after reopen.
8. **Per-account draft filing:** a plain `CreateItem` draft saves into the DEFAULT store's Drafts — create per-account drafts via the target store's `Drafts.Items.Add(0)`. `Reply()`/`ReplyAll()`/`Forward()` + `Save()` land in the source store's Drafts. Threading only ever via those three methods.
9. **The signature `GetInspector` touch leaves a hidden Inspector** inside Outlook — close it (`Close(olDiscard)`) after `Save()` when not displaying. **Word-document edits (the signature-override path) reach the item ONLY via `Inspector.Close(olSave)` on the SAME held inspector** — `item.Save()` does not flush them, an `item.Save()` between the edits and the close re-renders the document from the item and silently wipes them, and closing via a re-acquired inspector reference loses them too (probe-proven).
10. **Sent items get a NEW EntryID** — capture outcome snapshots before `Send()`; verify arrival via the Inbox side (the Sent-copy lags sweep windows).
11. **The audit log is load-bearing:** draft/save/send operations THROW when their audit line cannot be written (the EntryID is preserved in the error text). `send` refuses fail-closed on any token/content/identity mismatch, and recomputes the content hash inside the STA right before sending.
12. **Deleted items keep index rows** (with `IncludeDeletedItems=1`) — live samplers must exclude the test tag; agents should expect the occasional just-deleted hit to fail location with a re-run-search error.
13. **Items get a NEW EntryID on ANY move** (`MailItem.Move` returns the moved item — snapshot its EntryID; the old id goes stale, though cached Exchange may keep answering it briefly). **Folder sync-wedge (live-probed):** deleting items while they sit INSIDE a folder marks that folder "synchronizing local changes" on cached Exchange, after which the folder cannot be hard-removed from Deleted Items for the rest of the Outlook session — move items OUT to a permanent folder and delete them there instead; a folder with only move-history removes cleanly (`Folders.Remove` on the Deleted Items copy).

## Health & diagnostics

`outlook_health` reports everything the server depends on (Outlook process/version/headless state, probed COM liveness, store reachability, index freshness globally AND per store with actionable advice, WSearch service state, audit writability, tuning state, installer mutex) without ever starting Outlook — it merges the former `health` and `index_status` tools into one call. The audit log at `%LOCALAPPDATA%\OutlookAI\audit.log` records every write operation with timestamps and EntryIDs.

The tuning block of `outlook_health` also carries `uiSearchBackend` — the EFFECTIVE Outlook UI search backend read from the live registry (`DisableServerAssistedSearch`, policy hive authoritative over the user hive; absent/0 = Outlook's server-assisted default). `"local"` means the Outlook search box queries the same Windows Search index the agent's `search` uses; `"server-assisted"` means UI results are server-capped and differently ranked, so what the user sees can diverge from what the agent finds — `show_search_results` then appends a strong advice note recommending the Search tuning group be re-enabled. The registry value stays deliberately user-togglable: disabling it is the sanctioned mitigation for the documented replies-stick-in-Outbox side effect of local-search mode.
