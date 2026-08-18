# OutlookAI MCP - completeness gap map

**Why this file exists.** The product's worst possible failure is a search that quietly returns less
than it should - worse than being slow, worse than erroring, because nothing downstream can tell.
Several such defects were found on 2026-08-17/18 (date literals parsed in the wrong locale, a
freshness sweep anchored to the wrong account's frontier, coverage counters pooled across accounts),
so the whole surface was audited read-only for the same species.

**How to read the severity column.** **SILENT** means nothing in the payload says the answer is
partial - an agent cannot know. **PROSE** means only `advice` or the tool description mentions it, so
an agent that reads fields rather than sentences is not told. **REPORTED** means a field an agent can
branch on. The standing rule is that a partial answer must be reported in a form software can act on,
so PROSE is a defect too, merely a lesser one.

**Status.** This is the map as audited. Items get closed against it; a row is not removed when fixed,
it is marked, so the history of what was silent stays legible.

**How a closed row is marked.** The Gap cell keeps its original wording and gains a
`**CLOSED <date>** — <what the payload says now>` sentence; the severity cell becomes
`~~OLD~~ → **NEW**`. Nothing is rewritten, so the row still reads as the defect it was.

Closed so far: **B1**, **C1**, **C2**, **C3**, **F1** (2026-08-18), and **C4** downgraded from SILENT
to REPORTED by C1's work.

Working tree was NOT clean at audit start (see notes at end). No file in the repo was modified by this audit.

Severity key:
- **SILENT** — nothing in the payload says the answer is partial.
- **PROSE** — only `advice` (or a tool description) mentions it.
- **REPORTED** — a field an agent can branch on.

All paths relative to `c:\Source\SixFive7\OutlookAI\McpServer\`.

## 1. Empty / partial index for a whole store (the flagged Hyper-V PST case)

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| A1 | No index rows in scope, so the sweep window silently becomes "last 7 days". Everything older in that store is reachable by neither tier. | `OutlookAI.Core/Services/MailService.cs:880` (`?? DateTime.UtcNow - EmptyIndexSweepWindow`), constant at `:112` | `freshness:"live"`, `degraded` **absent**, `sweep.performed:true`, `sweep.foldersSwept:4`, `sweep.coverageGaps` absent, `staleness.newestIndexedUtc` **absent** (nulls omitted, `OutlookTools.cs:54`), `staleness.ageMinutes` absent, no advice (`DescribeStaleIndex` returns null on a null age, `MailService.cs:627-631`) | **SILENT** |
| A2 | `outlook_health` reports "Index is current; searches run at index speed." and `status:"ok"` when the index holds zero mail rows. | `MailService.cs:3768-3772` (fires when `advice.Count==0 && ageMinutes==null`) | `index.newestIndexedUtc` absent, `index.perStore: []`, `status:"ok"`, advice says current | **SILENT (worse: actively wrong prose)** |
| A3 | `index.perStore` is built from the index-derived catalog, so a mounted store with no index rows never appears; nothing compares it against `outlook.stores` (the COM list). | `MailService.cs:3727-3747`, catalog at `:4403-4409` | Two lists that disagree, with no field naming the disagreement | **SILENT** |
| A4 | Store-scoped search of a store absent from the index catalog throws instead of degrading to a COM-only answer. Also fires for a small-but-indexed store the unordered 2000-row discovery sample missed. | `MailService.cs:4345-4348`; sample SQL `IndexSearch/WsSqlBuilder.cs:279-289` (no ORDER BY); at-sign-only rescue at `MailService.cs:4307-4314` | `ArgumentException`: "Store 'X' was not found in the local index. Known stores: ..." (empty list on an all-PST profile) | **REPORTED (error), but misleading** |
| A5 | The fact is machine-readable only in a different tool: `list_accounts` returns `inLocalIndex:false`, `locallySearchable:false`. `search` never reads or echoes it. | `MailService.cs:3874-3893`, `ProbeStoreInIndex` `:4243-4277` | Correct field — in a tool the agent may never call | **REPORTED (wrong tool)** |

## 2. Freshness scope

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| E1 | An UNSCOPED search opens one sweep window from the PROFILE-wide frontier. A store whose own frontier lags by hours gets a window of minutes; its recent mail falls outside both tiers. | `MailService.cs:435-437` plus `StalenessScopeFor` `:607-609` (null when no store is named); window at `:880` | `staleness.newestIndexedUtc` = the profile value, `freshness:"live"`, `degraded` absent | **SILENT** (documented in `McpServer/README.md`, invisible in the payload) |
| E2 | The default sweep set is four arrival-path folders only, and shallow (no subfolders). Mail a rule files elsewhere before indexing is in neither tier. | `Com/OutlookComSession.cs:1268-1271` (`DefaultSweepFolderKinds`), shallow `SweepFolder` call at `:1633-1636` | `sweep.scope:"default folders (Inbox, Sent Items, Deleted Items, Junk Email)"` — a human string, not a flag; `freshness:"live"`, `degraded` absent, no advice | **PROSE** |
| E3 | A cached sweep (up to 10 s old) can miss an arrival that happened inside the TTL. | `Services/SweepCache.cs:31`, use at `MailService.cs:951-960` | `sweep.cached:true`, `sweep.cacheAgeSeconds` — but `freshness:"live"`, not degraded | **REPORTED (weak)** |

## 3. Tier asymmetries (index vs freshness sweep vs exhaustive)

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| B1 | `from` matches the ADDRESS only in the index tier. `System.Message.FromName` is SELECTed and never used in a predicate; the sweep and exhaustive tiers both match name OR address. The tool description promises "address or name fragment". **CLOSED 2026-08-18** — the predicate is now `(CONTAINS(FromAddress,…) OR CONTAINS(FromName,…))`; `IndexQuery.FromAddressContains` renamed `SenderContains` so the field cannot re-acquire the old meaning. Measured on this index: FromName alone 18 ms vs FromAddress alone 42 ms for the same fragment, the OR pair -19 to +12 ms on agent-sized TOP 26 statements; recall `vleuterweide` 0 → 18 rows, `subsidiewijzer` 0 → 9. 218 of 419 distinct senders (52%) in a 3 000-row sample carry a display-name token absent from their address. | `IndexSearch/WsSqlBuilder.cs:192-196` (only `CONTAINS(System.Message.FromAddress...)`), `:47` (FromName in SELECT); sweep `MailService.cs:1027-1031`; exhaustive `MailService.cs:1122-1126`; description `OutlookTools.cs:133` | A full-looking result set built from the sweep window alone; `freshness:"live"`, `degraded` absent | ~~SILENT~~ → **fixed** (all three tiers agree) |
| B3 | Item-class admission differs per tier and is never counted. Index: message rows need `System.Kind` to contain `email` (meeting requests index as `calendar`, so they are dropped). Sweep: **no class filter at all**. Exhaustive: `PR_MESSAGE_CLASS like 'IPM.Note%'` AND `Class==43`, dropping meeting requests/responses, NDRs and read receipts (`REPORT.IPM.Note.*`), `IPM.Post`, `IPM.Sharing`. | `IndexSearch/IndexRowFilter.cs:100-124`, `WsSqlBuilder.cs:157-161`; `Com/OutlookComSession.cs:7005-7101` (no class filter); `Com/ExhaustiveDaslFilter.cs:75` plus `Com/OutlookComSession.cs:6899-6913` | Nothing. `IndexSearchResult.RowsDropped` exists (`IndexSearch/IndexSearchService.cs:47`) but is never copied into `SearchOutcome` | **SILENT** |
| B2 | Attachment text is index-only. The refused attachment-ONLY case is properly flagged; the DEFAULT case is not — a term inside an attachment of just-arrived mail is invisible while the answer says `freshness:"live"`. | `Services/FreshMerge.cs:256-272`; default path `MailService.cs:464-505` | attachment-only: `sweep.error`, `freshness:"index-only"`, `degraded:true` (**REPORTED**). Default: nothing | **PROSE** (default case) |
| B4 | `search_in:"body"` means body **plus attachment content** in the index tier, `MailItem.Body` in the sweep, `urn:schemas:httpmail:textdescription` in exhaustive. | `WsSqlBuilder.cs:448-462`; `FreshMerge.cs:300-330`; `ExhaustiveDaslFilter.cs:44` | Nothing in the payload | **PROSE** (README only) |
| B5 | Whole-word (index) vs substring (sweep, and exhaustive's LIKE fallback). | `FreshMerge.cs:300-330`; `ExhaustiveDaslFilter.cs:113-130` | exhaustive: `exhaustive.engine`, `exhaustive.instantSearchEnabled` plus advice (**REPORTED**). The sweep's over-match is the safe direction | **REPORTED / benign** |

## 4. `thread`

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| C1 | `thread` runs NO freshness sweep and carries NO freshness fields. If the index has one or more rows for the conversation the COM walk never runs, so replies newer than the frontier are omitted. **CLOSED 2026-08-18** — the COM conversation walk is no longer a fallback: it runs on every call that has an anchor `id`, and its members are merged and deduped with the index rows (`FreshMerge.SelectFreshOnly`). `thread` now carries `degraded`, `freshness`, `staleness`, a `live` block and `advice`, on the same three-value contract `search` has. Sweeping was chosen over reporting-only because a conversation is a scope Outlook enumerates directly — one `GetConversation` plus one table walk, no window, no folder set, no per-folder cap — bounded by `ThreadWalkBudgetMs` (= `SweepBudgetMs`, 30 s), plus one ~2 s `HitLocator` probe the first time a given hit id is used (cached after). | `MailService.cs:1495-1511`; model `Services/MailModels.cs:620-639` | `source:"index"`, `truncated:false`, and **no** `degraded`/`freshness`/`staleness` at all. Tool description: "Fetch the full conversation of a mail" (`OutlookTools.cs:175-180`) | ~~SILENT~~ → **REPORTED** (`degraded`, `freshness`, `live.coverageGaps`) |
| C2 | `thread`'s index path is `KindFilter.EmailOnly`, so a meeting-request member of the thread is dropped. **CLOSED 2026-08-18** — new `KindFilter.MessagesAnyClass`: message-level rows of every item class, no attachment rows. `search` keeps `EmailOnly` (there the caller asked for mail); a thread member Outlook puts in the conversation belongs in the conversation. | `MailService.cs:1489` | Nothing | ~~SILENT~~ → **fixed** (members no longer dropped) |
| C3 | An unresolvable `store` silently widens the conversation lookup to the whole profile. **CLOSED 2026-08-18** — the widening stays (over-return is the safe direction) and is now reported as `scopeWidened:true` plus an advice sentence naming the store. | `MailService.cs:1473-1483` (`catch (ArgumentException) { scope = null; }`) | Nothing (over-return: the safe direction) | ~~SILENT~~ → **REPORTED** (`scopeWidened`) |
| C4 | The COM fallback walks one item's `Conversation.GetTable()` — same-store members only. **DOWNGRADED 2026-08-18** — still same-store, but no longer silent: when the index rows for the conversation name a store the walk did not cover, `live.coverageGaps` raises `unwalked_store`, `freshness` is `"partial"` and advice says to pass an `id` from that account. Undetectable when the OTHER store has no index rows either, so this row stays open. | `Com/OutlookComSession.cs:1993-2060` | `source:"com"` only | ~~SILENT~~ → **REPORTED** when the index can see the other store; SILENT when it cannot |

## 5. `exhaustive:true`

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| F1 | `degraded` and `freshness` are never set on an exhaustive result, so a scan that timed out and skipped folders looks undegraded on the two fields the tool description teaches agents to branch on. **CLOSED 2026-08-18** — `FreshMerge.ClassifyExhaustiveFreshness` sets both from the scan's own counters. The scan IS the live check, so `"index-only"` is unreachable there by construction (it names the state the mode exists to avoid) and no fourth value was invented; `"live"` when it covered its scope, `"partial"` when it timed out, skipped folders **or** truncated. Truncation counts here and does not in the merged path: an indexed `truncated` caps a date-sorted complete match set and pages, while the exhaustive cap stops the folder walk in arbitrary order (F2) and cannot be paged past. | `MailService.cs:1198-1220` (the exhaustive `return new SearchOutcome` sets neither; contrast the indexed path at `:562-580`) | `exhaustive.timedOut/foldersSkipped/truncated` present (**REPORTED**), `degraded` and `freshness` **absent** | ~~PROSE / inconsistent contract~~ → **REPORTED** (same two flags as a merged search) |
| F2 | The scan never sorts the folder table, so the `maxItems` cap truncates in arbitrary order and there is no way to page past it (top is capped at 100). Contrast the sweep, which sorts newest-first. | `Com/OutlookComSession.cs:6873-6923` (no `t.Sort`) vs `:7045` | `truncated:true` plus "raise top" advice, which cannot help beyond 100 | **REPORTED but unactionable** |
| F3 | `from` / `unread_only` / `has_attachments` are applied AFTER the scan cap, so an exhaustive search can return 2 rows with `truncated:true` while thousands match. | `MailService.cs:1118-1140` | `truncated:true`; the advice does not say the filter ran post-cap | **PROSE** |
| F5 | Rows dropped inside the scan are never counted: `GetItemFromID` failure, `Class` read failure, non-`IPM.Note` class. | `Com/OutlookComSession.cs:6890-6913` | Nothing | **SILENT** |
| F4 | `ScanFolderTree` has no depth guard (unlike `SweepFolderTree` and `CollectFolders`) — a cyclic tree is an uncatchable StackOverflow. | `Com/OutlookComSession.cs:6763-6822` | n/a — robustness, not completeness | tangential |

## 6. Scope resolution that yields zero instead of an error

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| G1 | `list_folders(store:"typo")` returns an empty tree with no error. | `Com/OutlookComSession.cs:1186-1205`; `MailService.cs:3995-4004` | `stores: []`, `folderTotal: 0`, `truncated:false` | **SILENT** |
| G2 | A store whose `DisplayName` COM read fails is silently dropped from `list_folders`; in the sweep its four default folders count as skipped in the total but land in nobody's per-store bucket. | `OutlookComSession.cs:1197-1200`; sweep `:1580-1592` | list_folders: nothing. Sweep: unscoped searches see `foldersSkipped` (**REPORTED**); store-scoped ones cannot | **SILENT / partial** |
| G3 | `CollectFolders` / `CollectFolderPaths` truncate at `FolderWalkAbsoluteCap = 10 000` and depth 64 with no flag; `FoldersOutcome.truncated` is computed against the already-truncated list. | `OutlookComSession.cs:7611,7625,7669-7720`; cap at `MailService.cs:220` | `truncated:false` on a truncated tree | **SILENT** (low probability) |
| G4 | The delegate folder-NAME list comes from the same truncatable walk; a short list under-returns the delegate folder scope. | `MailService.cs:4373` into `OutlookComSession.ListFolderPaths:7451` | `scope.folderNamesMatched` (a count, not a completeness flag) | **SILENT** (low probability) |
| G5 | A folder path that resolves in COM but not in the index (localized/renamed folder, non-recursive `ItemFolderPathDisplay` mismatch) yields zero index rows. The C7 guard catches it — but only when the MERGED answer is completely empty, and only as advice. | `MailService.cs:697-720`, `Services/FolderScopeResolver.cs:296-306` | advice only; one swept item suppresses the guard entirely | **PROSE, conditional** |
| G6 | Index-tier candidate exhaustion (the post-filter ran out of over-fetched rows) is advice-only; `IndexSearchResult.CandidatesExhausted` is never copied into `SearchOutcome`. | `MailService.cs:546-552` | an advice sentence; no field, and `degraded` stays absent | **PROSE** |

## 7. Row- and item-level silent drops in the freshness sweep

| # | Gap | Code path | What the caller sees | Severity |
|---|---|---|---|---|
| H1 | A table row whose EntryID column is missing, or whose `GetItemFromID` fails, is skipped and does not count toward the per-folder cap. A folder where every row fails returns `SweepOutcome.Complete` with zero items. | `Com/OutlookComSession.cs:7060-7089` | `foldersSwept` counts that folder as fully covered | **SILENT** |
| H2 | `t.Sort(...)` failure is swallowed; if the item cap then fires, the `item_cap` advice claims "reads newest-first, so the OLDEST ... is not covered", which is then false. | `OutlookComSession.cs:7045-7050`; advice `MailService.cs:801-807` | `item_cap` gap code (**REPORTED**) with a wrong explanation | **REPORTED, prose wrong** |
| I1 | `unread_only`, `has_attachments`, `after`, `before` all DROP a swept item when the underlying COM property could not be read (`IsRead`/`HasAttachments`/`ReceivedTime` null). | `MailService.cs:1033-1055` | Nothing | **SILENT** |

## 8. Caps that ARE properly reported (for map completeness)

| Cap | Value | Field |
|---|---|---|
| `top` | 1-100, default 25 | `truncated:true` (over-fetch-by-one, so definite) plus clamp advice (`MailService.cs:646-659`) |
| `SweepPerFolderCap` | 200 | `sweep.itemCappedFolders` plus `item_cap` in `sweep.coverageGaps` |
| `MaxScopedSweepFolders` / time budget / depth | 40 / 2000 ms / 64 | `sweep.folderCapReached` / `timeBudgetExceeded` / `depthLimitReached` plus gap codes |
| `SweptFolderListCap` | 12 | `sweep.folderListOmitted:true` |
| Sweep could not run | — | `freshness:"index-only"`, `degraded:true`, `sweep.error` |
| Sweep not needed | — | `sweep.notNeeded:true`, `freshness:"live"` |
| Delegate widening / leaf collision | 40 paths | `scope.shape`, `scope.widened`, `scope.folderNamesMatched` plus advice |
| `read` body/html/headers/recipients/attachments | 500k/100k/64k/100/100 | `bodyTruncated`+`bodyTotalChars`, `bodyHtmlTruncated`+`bodyHtmlTotalChars`, `headersTruncated`, `recipientsTruncated`+`recipientsTotal`, `attachmentsTruncated`+`attachmentsTotal` |
| `list_folders` page | 1000 | `truncated` plus `nextOffset` plus `folderTotal` |
| `thread` top | 1-200 | `truncated`; it caps the LIVE walk too, so a capped walk also raises `member_cap` in `live.coverageGaps` with `freshness:"partial"` (C1) |
| `thread` live walk could not run | — | `freshness:"index-only"`, `degraded:true`, `live.error` (`NoAnchorItem` / `UnknownAnchorId` / `AnchorNotLocatable` / COM token) plus advice (C1) |
| Exhaustive scan cap / time budget / skipped folders | `top` / 105 s / — | `exhaustive.truncated` / `timedOut` / `foldersSkipped`, and since 2026-08-18 `degraded:true` + `freshness:"partial"` (F1) |
| `move_mail` ids | 50 | error |
| `UnresolvedRecipientsCap` (drafts) | 20 | **no flag** — `Take(20)` at `MailService.cs:2374`, tangential to search |
| `snippet_chars` | clamped 0-1000 | silent clamp, cosmetic |

## Notes on the working tree

At audit start `git status` reported (contrary to the session snapshot):

    M  McpServer/OutlookAI.McpServer.Tests/T2/LiveStoreCountTripwire.cs   (+14 lines)
    ?? McpServer/OutlookAI.McpServer.Tests/T1/LiveOutlookPreflightTests.cs
    ?? McpServer/OutlookAI.McpServer.Tests/T2/LiveOutlookPreflight.cs

mtimes 08:43-08:52 local, i.e. minutes before this audit began; all three were already present in the first file listing this audit made. They account for the 1393 to 1416 test-count delta. Nothing here was written by this audit.
