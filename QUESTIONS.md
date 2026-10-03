# Open questions for the maintainer

**What this file is.** Questions that need a human decision, written down instead of blocking. Each
one states what is being asked, why it is open, the options, a recommendation, and - importantly -
**what happens by default if nobody answers**, so an unanswered question never stalls the work and is
never silently decided either.

**How to use it.** Answer inline under a question, or delete it and say what you chose. Anything
answered moves to the decision log at the bottom so the reasoning survives.

**Where decisions already made live.** `Docs/magic-numbers.md` carries every constant with a status
of Fixed, Kept - defensible, or Open - needs a decision. `CHANGELOG.md` carries the user-visible
half. This file is only for things that need *you*.

---

## Q1 - When to cut a release

The `## Unreleased` section has grown large: editable prompts and quick buttons, the tabbed settings
window, the writing-rules gate, model selection, Office version detection, the search truncation fix,
settings surviving uninstall, and the timing and drift-guard work. You said "no release, we have more
things to build" and that still holds as far as I know.

**Options.** Cut a minor release now and start a fresh Unreleased section; keep accumulating; or cut
a patch release purely to get the search-guidance truncation fix out, since that one silently
degrades every agent session today.

**Recommendation.** Keep accumulating while the work is this dense - a release mid-stream costs a
version bump and a changelog stamp for no user benefit, and nothing currently in Unreleased is a
field emergency. Revisit when the audit follow-ups are done.

**ANSWERED 2026-08-18: no release yet; the maintainer will say when.** I will not trigger the release
workflow autonomously regardless - the project's own rules make that an explicit-word action.

---

## Q5 - I reversed a decision that was made deliberately hours earlier

`e706315` established that a default folder a store does not HAVE is not a coverage gap, and its
test said so in as many words: *"absence is not a gap, but a sweep that ended up covering NOTHING
is - whatever the reason"*. That "whatever the reason" clause was deliberate.

It stopped being right when `c515565` made the coverage counters per store. Before, a sweep covering
nothing needed a whole profile with no arrival-path folder anywhere - vanishingly rare, so treating
it as a gap cost nothing. After, it describes an everyday PST or archive-only store, whose four
default folders are all absent: `foldersSwept: 0`, so every search naming that store reported itself
degraded. A review proved it.

So in `687929f` I reversed the clause: absence suppresses `nothing_swept` when it is the whole story,
while one absent folder beside one unreadable folder is still a hole, and a scope the sweep never
reached still degrades.

**Why this is a question and not just a fix.** I overrode a judgement someone made explicitly, with
its reasoning written down, a few hours after they made it. That is exactly the kind of change worth
a second opinion - the reasoning may have covered a case I did not see.

**Recommendation.** Keep the reversal. The original clause was correct for the shape of the data it
was written against and wrong for the shape that existed six commits later; the test now records
both readings so the history is legible.

**ANSWERED 2026-08-18: keep it, and CONFIRM IT ON THE PST-ONLY TESTBED.** That machine - Outlook with
no accounts and only local PSTs - is the shape where all four arrival-path folders are legitimately
absent. What to look for there: a search naming a PST store must come back complete and correct, NOT
flagged degraded. If it also reports `no_index_frontier`, that is the separate and expected finding
that the PST is not in the Windows Search index. **This verification has not been done** - it needs a
machine this session cannot reach.

## Q7 - Three follow-ups the freshness work deliberately did not decide

Raised by the agent that added the `no_index_frontier` state; none blocks anything.

**(a) `staleness.newestIndexedUtc` on an unscoped search** is still the profile-wide maximum. It is no
longer the sweep's window base, and the advice age now uses the widest per-store frontier, but the
field itself still reports the maximum - because narrowing it would make `search` and `outlook_health`
report different numbers for the same profile. Options: leave it; add
`staleness.oldestStoreFrontierUtc`; or change the field's meaning and update health to match.
*Recommendation: add the second field.* It answers the question without making two tools disagree.
**ANSWERED 2026-08-18: do this. SHIPPED** - `staleness.oldestStoreFrontierUtc`, beside the unchanged
`staleness.newestIndexedUtc`. Store-scoped search: the same value as the existing field, because one
store is in scope and its frontier is both the newest and the oldest - emitted rather than omitted so a
caller reading only the new field gets a true answer on every search shape. Unscoped: the earliest of
the per-store frontiers the sweep planner already measures, which is the figure the freshness advice has
been quoting all along. Absent when no per-store frontier was measured at all (an exhaustive search, or
an unscoped one whose store catalog could not be read); absence means "not measured", never "no lag" -
substituting the profile maximum there would put a number in the field that no store's index stands at.
Decided by the pure `MailService.OldestStoreFrontier`, all three branches pinned in T1.

**(b) The unindexed-store list is uncapped.** Every other list in this server has a cap and a has-more
flag. A profile with many unindexed PSTs would list them all, in the payload and in an advice
sentence. *Recommendation: cap it like the others* - the principle is already settled here, this is
just an omission. **ANSWERED 2026-08-18: do this. SHIPPED** - `MailService.UnindexedStoreListCap`,
derived from `SweptFolderListCap` (12) rather than written as a second 12, since both bound a name list
in the same sweep block for the same reason. The list is TRUNCATED rather than dropped, which is where
it differs from the swept-folder list: a folder list is a legibility aid and is worth nothing in part,
while each unindexed store NAME is separately actionable. Reported as `sweep.storesWithoutIndexTruncated`
and `sweep.storesWithoutIndexTotal`, and the `no_index_frontier` advice sentence names the cap and the
remainder too - capping the payload alone would have left the whole list in the prose an agent relays to
the user. T1 pins the cap, both flags, and both wordings of the sentence.

**(c) `notNeeded` now costs one ordinary sweep** in a narrow case: an unscoped search bounded to mail
older than the frontier but newer than the fallback runs a sweep where it previously did no COM work.
That is the price of `notNeeded` no longer lying on unindexed stores. *Recommendation: accept it.* Any
mitigation trades a bounded window of completeness for latency, which is the trade the standing rule
forbids. **ANSWERED 2026-08-18: accepted.**

## Q8 - The three search tiers disagree about what counts as "mail"

Audit gap B3, and the last item in the top ten I have not touched, because it is a product decision
rather than a defect.

**Primer.** A search can be answered by three different engines and they admit different item classes:

- **Index tier**: requires `System.Kind` to include `email`. Meeting requests index as `calendar`, so
  they are excluded.
- **Freshness sweep**: no class filter at all. It returns whatever is in the folder.
- **Exhaustive scan**: `PR_MESSAGE_CLASS like 'IPM.Note%'`, so no meeting requests, and **no NDRs or
  read receipts** (`REPORT.IPM.Note.*`), no `IPM.Post`, no `IPM.Sharing`.

**Why it matters.** The same query gives different item sets depending on which tier answered, and
nothing in the payload says so. A meeting request found by the sweep today vanishes once it is
indexed. The mode that exists for correctness - exhaustive - is the one blind to bounce messages, so
"did my mail bounce?" is unanswerable exactly where a user would go looking hardest.

**Options.** *(a)* Make all three admit the same set, whatever it is - one rule, one place.
*(b)* Keep the tiers different but REPORT the difference, so an agent knows a result came from a tier
that excludes meeting requests. *(c)* Define "mail" narrowly and consistently (`IPM.Note` plus
reports) and exclude calendar items everywhere. *(d)* Leave it.

**Recommendation.** *(a)*, with the set including NDRs and read receipts, because those are mail a
user asks about by name. But which classes count is your call, not mine - it changes what every
search returns, and I would rather ask than pick. `RowsDropped` already exists in the index layer and
reaches no payload, so whatever is decided, the count of what a tier refused should surface.

**ANSWERED 2026-08-18: unify all three, and prefer returning EVERYTHING where possible.** So the
admission rule is one rule in one place, as inclusive as each tier can be made - NDRs, read receipts,
meeting requests and post items included - rather than three different narrowings. Where a tier
physically cannot reach a class, that is a coverage fact to report, not a filter to leave implicit.

**SHIPPED 2026-08-18.** The rule lives in `McpServer/OutlookAI.Core/Mapi/MailItemAdmission.cs` and it is
that **an item's class never excludes it**; what bounds a search is the folder it looks in. It is
written as a method that cannot return false, so a future narrowing has to delete a call site and a T1
assertion, both of which say what is being given up - rather than quietly adding a class test next to an
item loop, which is how the three tiers drifted apart in the first place.

*An allowlist of "mail-ish" classes was considered and rejected*, and one fact decides it: the
SystemIndex carries no message-class column at all, so an allowlist could only ever be enforced in the
COM tiers - replacing one asymmetry with another, in the same payload, for the same query. Unifying
UPWARDS to the widest of the three (the sweep, which never filtered) is the only shape that leaves the
tiers agreeing.

- **Freshness sweep**: unchanged. It is the tier the other two were unified to.
- **Exhaustive scan**: the `PR_MESSAGE_CLASS like 'IPM.Note%'` clause and the `Class == 43` gate are
  both gone. It now returns bounce reports, read receipts, meeting requests and responses, posts and
  sharing invitations - the mode chosen BECAUSE completeness matters is no longer the one blind to "did
  my mail bounce?". Where a scan has no terms and no dates to restrict on it emits a predicate that
  matches every class, because `@SQL=` with no predicate is not a restriction Outlook accepts.
- **Index tier**: message-level rows are admitted whatever their `System.Kind`. `KindFilter` was renamed
  with the rule (`MessagesAndAttachments` / `MessagesOnly` / `AttachmentsOnly`, plus `MailKindOnly`
  which only store discovery uses), because names carrying the old narrowing would be the same defect
  one level down.

**What each tier still cannot reach, reported rather than implicit.** The COM tiers only enter folders
whose `DefaultItemType` is `olMailItem` - unchanged, and not a class filter: it is where mail lives. The
index tier has no folder-type column and no message-class column, so it cannot draw that same line: its
widening also admits the calendar and contact items of folders the COM tiers never open. That is
over-return rather than under-return, which is the direction the standing rule prefers, and it is
visible - every hit that is not ordinary mail carries `itemClass`, and one advice sentence names the
count and the classes when an answer holds any.

**The counts of what a tier refused now surface.** `index.rowsScanned` / `index.rowsDropped` /
`index.candidatesExhausted` are a new block on `SearchOutcome` (the last of those also closes audit gap
G6). Adding it had previously been declined on the `search` description budget; that cost does not
exist - the client cap is per string, a payload block needs no description text, and `search` measures
1791 units before and after the change. On the exhaustive side, `rowsDropped` minus `rowsUnreadable` was
exactly this item-class filter, so that difference is now **zero by construction** - which is the
machine-checkable statement that the tier admits every class.

**Not verified here**: that a real meeting request, NDR or read receipt comes back from all three tiers
on a live profile. That needs the live tier, which this work did not run.

### 2026-08-18 follow-up: the widening's one open risk, closed by construction

The commit above flagged a risk it could not settle. It is real, and it is worth stating precisely,
because the precise version is not quite the one the flag described.

Undated rows compete for `SELECT TOP n ... ORDER BY System.Message.DateReceived DESC` on terms nobody
here has measured: an appointment or a contact carries no `System.Message.DateReceived`, so where the
provider sorts a NULL under `DESC` decides whether they fill the `n`. The server-side sort that puts
them last runs on rows the provider already truncated, so it cannot recover any of it. What the commit
changed is **what happens next**:

- `include_attachment_hits: true` (THE DEFAULT) already emitted no kind predicate under a SCOPE before
  the commit, so those rows could always take slots. The post-filter then dropped them, so a
  NULLs-first provider produced a SHORT answer - and `candidatesExhausted` fired, which is the whole
  point of that counter. Loud.
- `include_attachment_hits: false` used `KindFilter.EmailOnly`, which put `System.Kind='email'` **in
  the SQL**. That shape was immune, and is not any more.
- After the commit both shapes ADMIT the undated rows, so the answer is full length and can contain no
  mail at all. **The loss stopped being visible.** That is the regression: not that displacement
  became possible, but that the one signal which would have shown it went quiet.

**The guarantee now shipped, and why it holds.** *A row the index cannot date can never reduce the
number of dated rows a search returns.* Two things could take a slot from mail and both are closed:

- **The client-side trim.** The service took the provider's first `Top` admitted rows. It now orders
  rankable rows first, by their key, with unrankable rows after them
  (`IndexOrderGuard.RankableFirst`), and trims after that - the same "undated last" convention
  `MailService` already applies when it merges sweep hits into the same list. This alone fixes every
  case where the statement was not truncated, because then every matching row is already in hand.
- **The provider-side cut.** Where the statement WAS cut off and an unrankable row came back in it,
  rankable rows may never have left the provider, and no client-side ordering can recover them. The
  service then re-runs the same statement with one added predicate that admits only rows carrying the
  ordering column, and unions the two answers. The union can only add.

The trigger is `truncated AND at least one unrankable row in the block`, and it is sound under **any**
collation, including an interleaved one: below the TOP nothing was displaced, and an unrankable row
that sorted above the cut would be IN the block by definition. Both halves are pure functions with a
T1 suite (`IndexOrderDisplacementTests`, 15 tests) that drives the real service through a scripted
provider in both collations - NULLs-first must return the mail, NULLs-last must not pay for a second
statement.

**What it costs.** Nothing when the provider sorts NULLs last (no second statement is ever issued) or
when a search carries an `after`/`before` bound (a date predicate already excludes undated rows). One
extra index statement per truncated search if NULLs sort first, on the order of 40-100 ms by the
measured shapes in `Docs/magic-numbers.md`. Plus mapping every returned row rather than the first
`Top` of them, which is pure CPU on at most 5000 rows and is what makes `index.rowsScanned` /
`index.rowsDropped` finally mean what their names say. **Measured 2026-10-03 on `OutlookAI-Indexed`,
inside a mapi `SCOPE`: NULLs sort LAST** and the floor literal is accepted - so on that provider the
second statement is never issued (`Docs/live-tier-on-the-vm.md` section 4.2e).

**What was deliberately NOT protected.** A dated meeting request, bounce report or read receipt can
still push an older mail off the end of a `Top n` list. That is the B3 decision working: under
`MailItemAdmission` those ARE mail, and they compete on the same axis as mail. Ruling that out would
mean re-narrowing the tier this decision widened.

**Rejected alternatives.** Making the ordering explicit (WS-SQL has no `NULLS LAST` and no `COALESCE`
in `ORDER BY`); a bigger over-fetch (a Calendar folder can hold more undated rows than any bounded
factor, so a safe factor does not exist); excluding undated rows outright (works, and re-narrows the
tier - it would also drop unsent items, which are mail); and running the mail-only statement as a
floor on every search (an unconditional second query that recovers less than the date-floor shape,
since a dated meeting request is not `kind='email'`).

**Two smaller items from the same commit, settled.**

1. **The `PR_MESSAGE_CLASS like '%'` predicate stays.** The question was whether a different
   always-true predicate is more clearly correct. None is, and the reason is structural: MAPI
   documents the result of a restriction over a property the message does NOT have as **undefined**,
   not false, so a row whose property is absent may be admitted or dropped at the provider's
   discretion and no predicate over that property can promise either. (Corrected 2026-08-19 - this
   read "excludes", which invites the inference that `NOT (...)` therefore admits such a row. It does
   not: negating an undefined value leaves it undefined, and that inference is what made a broken fix
   for the sweep's date restriction look viable.) A different predicate moves that risk rather than
   removing it, onto syntax this codebase has never emitted. The only construction that removes it is
   no restriction at all
   (`Folder.GetTable()` with no argument), which was considered and not taken: PR_MESSAGE_CLASS is
   required on every MAPI message and is what Outlook itself reads to choose the item type it hands
   back, so the absent case is unreachable through the object model, while dropping the filter changes
   a COM call site nothing outside a live profile can exercise and makes the reported scan engine
   (`"like"`) a claim about matching that never happened. The residual doubt is only whether the
   provider reads `%` as "any string", and that failure is loud: `GetTable` throws, the folder is
   counted skipped and a coverage gap is raised.
2. **The extra `MessageClass` read is plainly fine; no measurement needed.** The sweep's fitted cost
   model is ~19 ms per folder plus ~15 ms per item opened (215 sweeps, `Docs/magic-numbers.md`), so
   one read out of nine is ~1.7 ms per item. Steady state opens 0-5 items across all 20 arrival-path
   folders, so the whole addition is under 10 ms against a 30 s budget; the empty-index path's 377
   capped items add ~0.6 s to a predicted 6.0 s sweep. It cannot decide that budget either way,
   because the eight reads already there blow it first - the 200 x 4 x N worst case is ~60 s at eight
   reads before this one is counted. On the exhaustive tier it is not an addition at all: that loop
   used to read `item.Class` on every item in order to drop non-mail, so an admitted item paid nine
   reads then and pays nine now.

**Still needs a live profile** (`T2 LiveOrderKeyCollationTests`, read-only, written and NOT run):
where the provider sorts a NULL under `DESC` - which decides whether the refetch fires on every
truncated search or on none of them - and whether it accepts the `1601-01-01 00:00:00` floor literal.
Until that runs, the guarantee rests on construction rather than on measurement; the failure mode of
an unaccepted literal is a flagged short answer (`index.candidatesExhausted`), never a silent one.

## Q9 - How much of a mailbox the count tripwire should identify, now that identifying is cheap

**Primer.** The tripwire censuses every store before and after a live run. Every folder is counted;
folders inside a budget are also walked item by item, which is what lets a firing say WHICH items
left and lets it prove that an item was FILED rather than deleted. The budget is 500 items per
folder and 3,000 per store. It was set when the walk cost five cross-process calls per item, and on
2026-08-20 that cost refused the whole live tier - one delegate store's census exceeded the
3-minute STA budget. The walk is now a bulk `Table.GetArray` read: the same 3,000 items cost about
fifteen calls instead of fifteen thousand. So the reason the budget is where it is has largely gone,
and the budget was deliberately left alone because moving it changes what the guard proves.

**Why it matters.** At 500 per folder, a 4,918-item Sent Items and a 108,144-item Archive are
counted and not identified. In those folders the guard can see that items left but not which ones,
cannot tell a deletion from a filing, and is blind to a deletion masked by an arrival. Those are
exactly the folders a runaway test would do the most damage in.

**Options.** *(a)* Leave 500/3,000 - the guard covers the folders a person works in, and everything
else is still count-guarded. *(b)* Raise to about 5,000 per folder and 25,000 per store, which
covers Sent Items and most working folders and costs roughly 125 table calls per store. *(c)* Raise
the per-folder limit but keep a tight per-store budget, so one huge folder cannot consume the whole
allowance. *(d)* Raise only for non-delegate stores, on the reasoning that delegate stores are the
slow ones and the least likely to be written to. *(e)* Identify everything, no budget - honest but
unbounded, and a 108,144-item Archive is roughly 550 table calls and tens of MB of EntryIDs held
twice.

**Recommendation.** *(b)*, once one live run has printed what the census actually costs. The whole
argument for the old number was cost, the cost measurement has never been taken, and the first run
after this change prints per-store timings precisely so the next move is made on a number rather
than an argument.

**Default if nobody answers.** 500/3,000 stays. The guard is not weaker than it was yesterday; it
is simply not stronger than it could be.

## Q10 - Whether to raise the census STA timeout, and what to do about a store that is still slow

**Primer.** Every census runs on its own short-lived STA thread with a 3-minute join. That budget is
already per store - `CaptureMailFolderCensus` is one `RunSta` call per store - so the sometimes
suggested "make it per store rather than per operation" is already true. On 2026-08-20 one store
exceeded it and the live tier refused to run, correctly, but the refusal could not say whether the
time went into the folder tree or the item walk.

**Why it matters.** A live run costs about half an hour of a real mailbox's day. Failing at the
census wastes the whole slot. Raising the timeout risks turning a wedged Outlook into a longer wait
instead of a faster answer.

**Options.** *(a)* Leave it at 3 minutes; the term that blew it has been removed and a repeat
failure will now say where the time went. *(b)* Raise it to 6-10 minutes so the first post-fix run
completes and produces the measurement even if a store is slower than expected. *(c)* Make it
adaptive: a short budget for the first store, extended for later ones once one store's real cost is
known. *(d)* Keep 3 minutes but let a single store's census FAIL SOFT into a count-only census for
that store, rather than refusing the tier - explicitly weaker, and it would need saying in the
verdict.

**Recommendation.** *(a)*. Diagnosis was the thing missing, not headroom, and it has been added: the
per-store log line and the progress in the refusal message make one more failed run cheap and
informative. *(d)* is the one to avoid - it converts a fail-closed guard into a fail-quiet one.

**Default if nobody answers.** 3 minutes stays.

## Q11 - Does an Outlook `Table` report date-time values in UTC or in local time

**Primer.** This repository now contains two opposite readings of the same variant.
`CensusTableRow.ReadUtc` (the tripwire census) takes a `DateTimeKind.Unspecified` value from a table
as already-UTC. `OutlookComSession.ReadRowDate` (the shipped exhaustive scan and freshness sweep)
takes the same value and calls `ToUniversalTime()` on it, which treats it as local. A COM-marshalled
`VT_DATE` always arrives as `Unspecified`, so exactly one of the two is wrong on any given machine,
by the size of the local UTC offset.

**Why it matters, and it is not symmetric.** In the census, being wrong is cosmetic: every value at
both ends of every comparison comes through the one method, so items still match each other, and
only the instant printed beside a departed item would be offset. In `ReadRowDate` it is not
cosmetic: the value becomes `_lastAdmittedUtc`, which becomes a resumed exhaustive scan's inclusive
"at or before" date bound. A bound two hours early skips the mail received in those two hours, and
the scan reports itself complete - in the one search mode a caller chooses BECAUSE completeness
matters.

**Options.** *(a)* Settle it with one live read and fix whichever side is wrong.
*(b)* Fix `ReadRowDate` to match the census on the documentation alone. *(c)* Leave both; the census
is self-consistent and the scan's error is bounded by the local offset.

**Recommendation.** *(a)*, and it is nearly free: `T2/LiveTableSortProbeTests` already reports
`FirstRowReceivedUtc` through `ReadRowDate`. Reading the same item's `MailItem.ReceivedTime` beside
it in that probe answers the question for both call sites in one live read. Note that the probe has
been written and deliberately not run since 2026-08-19, so this rides along with whatever run
settles the `Table.Sort` question.

**Default if nobody answers.** Both stay as they are. The census is safe; the scan carries an
unquantified resume gap of at most one UTC offset.

**ANSWERED 2026-10-03 by measurement - option (a), and the answer is "it depends on the column's
spelling".** See the decision log below.

## Decision log

### 2026-10-04, autonomous - Q125 carried out: the SDK's security analysers, CodeQL locally, a package gate

**Primer.** Removing the GitHub CI also removed CodeQL's scan and the pull-request dependency review.
The maintainer chose (b) - turn on the security analysers the .NET SDK ships, in the builds of both
code bases - plus an exception to the Dependencies rule letting CodeQL run locally, and leaned
towards `dotnet list package --vulnerable` as a release gate. AGENTS.md (Build and Release,
Dependencies) now holds the rules; this entry holds the choices made on his behalf.

**What the analysers found, and what became of it.** MCP server, 78 sites: CA5392 (71 P/Invokes
with no DLL search path) FIXED - every assembly now loads its DLLs from System32 only; CA2100 (the
index client's OLE DB command) suppressed - every statement comes from `WsSqlBuilder`'s allow-list
and quote-doubling, the provider takes no parameters, and WS-SQL is read-only; CA5351 (MD5 in a
test, reproducing Outlook's ConversationId) and CA5394 (a fixed-seed `Random` in a test) suppressed.
Add-in: the same System32 fix covers its 16 P/Invokes and the elevated helper's 12 - for the helper,
which runs as administrator from the per-user install folder, a planted `wtsapi32.dll` (not a
KnownDLL) would have run elevated: the one real security fix here. CA5386 (the updater's
`|= Tls12`) suppressed, measured: in the legacy TLS state the add-in runs in, the analyser's own fix
(`SystemDefault`) fails against GitHub outright, and `|= Tls12` is the only setting that reached it.

**Decided on his behalf, each with what it beat:**
1. *The level: Security only, `10.0-all`, every finding an error.* Beat "recommended" (misses the
   injection, deserialisation and DLL-search rules), every category at "all" (thousands of style
   findings), `latest` (an SDK update would change the rules under a passing build) and warnings
   (unread in a scripted build).
2. *Every project, tests included* - the tests run against real mailboxes. Beat shipped-only.
3. *The add-in imports the SDK's own analyser targets* through an SDK-resolved import
   (`NetSecurityAnalyzers.targets`). Beat a NuGet package (its build has no restore), a hard-coded
   SDK path (breaks on every SDK update) and a second, SDK-style project over the same sources.
4. *CodeQL on the workstation, build-mode none, the `security-extended` suite, CodeQL's default
   threat model.* Beat the build VM (offline: 700 MB of staged media and VM time per run), a traced
   build of the add-in (full VSTO type resolution, but a build through the Q81 path for every
   scan) and the narrower default suite the old workflow ran. Whether the release should also count
   local input as untrusted is left open (TODO.md): measured, it is 0, 6 or 174 more results to
   triage depending on how much of "local" is on.
5. *The gate: every CodeQL result fixed, or triaged in `Tools/codeql-accepted.json` by rule, file and
   line hash with a reason; an entry that matches nothing is reported, not fatal.* Beat a
   severity threshold (a medium finding would pass untriaged) and inline suppression comments
   (scattered through the code, where the one reviewed list keeps every triage in sight).
6. *Release step 7, before the build: packages, then CodeQL; anything but exit 0 refuses - and "could
   not ask nuget.org" refuses too.* Beat a report-only step and a place after signing.

**First runs (2026-10-04):** CodeQL 2.27.1 over 04327b7 and 0ccf5df - 432 of 432 C# files, 55
security queries, no result; a probe commit with a DES/ECB cipher failed the gate and passed once
triaged.
The package check: six projects clean; Newtonsoft.Json 12.0.1 in a scratch project failed it.
**Not measured:** a release dry run with step 7 in it. **Undo:** revert the commit.

### 2026-10-03, autonomous - the indexed guest's first live runs: sixteen failures were the tests, one the product

**Primer.** The first full live run on `OutlookAI-Indexed` (d4e31fe, runbook 4.2f) failed 17 of 123
tests. Thirteen stopped at "Store 'Corpus A' not found among 3 discovered index scopes", and four
failed on their own. Sixteen were tests that had never met an indexed guest with a 160,000-item
store, or an Office 2024 behaviour; the seventeenth found a product defect (item 5). Each is decided
here, with three failures another agent saw on the same guest (item 6), the crash in run 4 (item 7) and where
the guest rests (item 8).

**1. A store the discovery sample cannot reach.** The live tier found a store's index scope in a
2000-row unordered sample of mail rows, then, for an address-named store the sample missed, by the
mail addressed to it. The sample held 1,941 rows of another profile's 20,000-item store, 59 of the
hub's, 1 of the bystander's and none of Corpus A's, and no address names Corpus A. *Options:* (a) add
the index's own store-root listing as a third step - the listing the product's store map already
reads - taking the one root of that name with anything indexed below it; (b) resolve the scope
through the product's store map (Microsoft's store hash) in the fixture; (c) a bigger or ordered
sample; (d) take Corpus A out of the indexed stores. **Decided: (a)** (`T2/LiveIndexScopes`,
`T1/LiveIndexScopesTests`): it finds every store the index holds, names its failure, never guesses
between two roots of one name, and refuses a root with nothing below it; (b) needs the COM store
details and the hash inputs in a fixture that measures the index on its own, (c) only moves the
cliff, and (d) would stop measuring the one store the latency bounds exist for. **Undo:** revert.

**2. The fresh-mode frontier at the index's precision.** `LiveFreshModeTests` accepts an index hit
only when the index frontier covers the send. The index served the hit with its frontier at
16:02:34.0000000Z for a send at 16:02:34.3804579Z: it had indexed the arrival within the second, and
keeps whole seconds. *Options:* (a) compare at whole seconds; (b) compare with the arrived item's own
indexed time; (c) a fixed tolerance; (d) demand the live sweep win. **Decided: (a)** - it is the
precision of the data; a frontier a whole second behind still fails. (b) needs a second query for the
same answer, (c) is arbitrary, (d) would race a fast indexer. **Undo:** revert.

**3. The pin release must not reconnect.** On Office LTSC 2024 closing the last visible window
raises Quit (the D49 entry below), so in `LiveDisconnectRecoveryTests` the parked windows' close
ended the pre-existing Outlook, and the step that releases its pin reconnected - into an Outlook
shutting down (RPC_S_SERVER_UNAVAILABLE); against one already gone it would have STARTED Outlook.
*Options:* (a) release the pin only on a session that still answers, and read the two RPC
disconnect codes in that race as "already quitting"; (b) a gateway call that never connects; (c)
catch every exception. **Decided: (a)**: the wait that follows still proves Outlook exited; (b) is
product surface for one test, (c) would hide a real failure. **Undo:** revert.

**4. The apostrophe test asserted the zero-row guard's old contract.** It demanded "matched NOTHING
in the index" stay quiet whenever the merged answer was non-empty. Gap G5 changed the product on
purpose to judge the INDEX tier's own rows (`T1/SearchCoverageClaimTests`), and no indexed hub had run
the test since; on this guest the folder, created a second earlier, was new to the index and the
guard said so. *Options:* (a) hold the test to G5 - flag and sentence agree, an index row for the
item keeps the guard quiet, an unindexed hub never trips it; (b) revert G5; (c) wait for the index to
reach the folder, then demand silence; (d) drop the assertion. **Decided: (a)**; (b) would undo a
decision for a stale test, (c) is the stronger proof of the escaping and is left as a possible
follow-up, (d) would be loosening. **Undo:** revert.

**5. A folder's own index row came back as a search hit - a product fix.** `search` on the hub
with no query returned 100 hits for a store of 68 items. Counted by store, tier, folder and item
class (two runs): 67 mail items, 12 contact cards and 21 rows of `System.Kind = folder` - one for
every folder of the store, its root, Calendar, Quick Step Settings and the emptied subfolders in
Deleted Items included - undated, with no item segment in their URL, and nothing a caller can open.
Gap B3 dropped the kind predicate under a mapi scope on purpose, so an appointment or a contact card
is admitted as over-return a caller can see; a folder was never meant to be a hit (the overnight
review recorded folder rows as rows "no search returns"). *Options:* (a) drop a row that is of kind
`folder` AND addresses no item (no EntryID decoded from its URL) in the index tier's admission;
(b) drop every message-level row whose URL addresses no item, on the URL alone; (c) put a kind
predicate back under a scope; (d) leave it, and count only items in the test. **Decided: (a)**
(`IndexRowFilter.IsFolderRow`, T1 `IndexRowFilterTests`): both halves are required, so an item is
never taken for a folder on its kind and a zipped-folder attachment stays an attachment; (b) would
also drop the synthetic rows a large part of T1 builds its searches from and rests on a decode the
product never needed for admission, (c) would undo B3, (d) would hide a defect users see. **Undo:**
revert the commit.

**6. Three failures another agent saw on this guest, not reproduced under the procedure.** A
folder-path run from `CP-17C` (coordinator heads-up) failed `LiveMoveArchiveTests.MoveChain`
("hub archive resolution failed after archiving: NoDesignatedArchiveFolder", which left an item in
the hub's Archive and failed five later tests' hub check), `LiveSweepScopeTests.ControlledCorpus`
(its self-sent mail never arrived through the mail sink) and the apostrophe test (item 4). That run
staged the suite onto the checkpoint's running Outlook: no graceful restart, no step 9a or 9a-ii.
Here MoveChain passed in all five full runs and in a narrowed run that restarted the guest but
skipped 9a; ControlledCorpus passed in all five. *Options:* (a) no code change - the one
condition every pass shares and the failing run lacked is the graceful restart, which
`Invoke-LiveTierOnGuest.ps1` makes unconditional - and keep the question open until it recurs under
the procedure; (b) re-read the Archive designation from a fresh store object in the test's verify
session, unproven against a failure nobody can reproduce; (c) reproduce it with a run that skips the
restart. **Decided: (a)**, with the question kept in `TODO.md`; (c) is the way to close it.

**7. Outlook crashed inside Word once, after master's all-kinds data was merged.** Run 4 (the merged
master on `CP-18C-ALL-KINDS`) lost `OUTLOOK.EXE` in `LiveDraftOptionsTests.NewDraft_Hub_SignatureOverride_*`:
`wwlib.dll`, `0xc0000005`, and the class's next two tests failed while it went. Run 5, the same commit
and checkpoint, passed 122 of 122 with no crash. The signature path that test drives captures and
releases every Word object it touches, unlike the guest-two crash, which stopped once every COM child
object was released; and the guests keep no crash dumps, so there is nothing to read. *Options:* (a)
take the green run, and make the next crash leave a dump - Windows Error Reporting's LocalDumps for
`OUTLOOK.EXE` on both guests - before chasing it; (b) rerun the draft-options class in a loop until it
reproduces; (c) review Word's threading in the signature path now, without evidence of where it
faulted; (d) call it Office's. **Decided: (a)** (`TODO.md`); (b) costs guest hours for a 1-in-5 event
with nothing to read when it hits, (c) has no fault site to start from, (d) is unearned.

**8. Where guest one rests.** The green run's checkpoint is `CP-19C-LIVE-GREEN` (child of
`CP-18C-ALL-KINDS`). *Options:* rest on it, or on `CP-18C-ALL-KINDS` with the 30/60 settings.
**Decided: `CP-18C-ALL-KINDS`** - the base the coordinator named and every agent's phase restores; the
green state is a run's end, with a hub the next run rebuilds anyway, and stays beside it as evidence.

### 2026-10-03, autonomous - D49 on Office LTSC 2024: the show-me window was the lifetime pin itself

**Primer.** D49: a live session holds a non-displayed Explorer - the lifetime pin - so that an
Outlook OutlookAI started without a window does not exit when the last window closes. On the test
guest (Office LTSC 2024, 16.0.17932) `LiveDisconnectRecoveryTests` failed in every run (runbook 4.1e,
F9): Outlook started headless, `goto_folder` put one window on screen, the window was closed, and
Outlook exited, with the session still reporting itself pinned. The question set: which kinds of
window Office 2024 counts as keeping Outlook open - then fix, or scope the test to what is measured.

**Measured** (two probes on the guest, Explorer windows only, nothing in any mailbox touched; the
maintainer's own Office was not touched). Probe v2: whatever was kept open first - nothing, a hidden
Explorer, a displayed Explorer parked off-screen and hidden - closing the shown window ended Outlook;
but `Explorers.Count` stayed 1 after a second Explorer was added, and with a first Explorer on
screen no second window appeared at all - a confound. Probe v3 resolved it: `Explorers.Add` on the
folder a non-displayed Explorer already shows returns THAT Explorer (the same COM object, the count
unchanged); on another folder it makes a new one; and closing the new window leaves Outlook running,
held by the non-displayed Explorer, even after the client released every reference. So the pin
works on Office 2024 - and the show-me path was displaying the pin itself, because the pin sits on
the default Inbox and `goto_folder` on that Inbox was handed the pin by `Explorers.Add`.

**Options.** *(a)* Never let the show-me path return an Explorer that already existed: when `Add`
hands one back, open the window on the store's top folder and navigate it to the folder asked for.
*(b)* Pin on a folder no show-me call asks for. *(c)* Accept that Outlook ends with the window on
Office 2024, scope the test to that and rely on the re-attach. *(d)* Re-pin after every show-me call.

**Decided: (a)** - a fix, because the measurement shows Office 2024 can keep the promise. It changes
nothing where `Add` makes a new Explorer, as the maintainer's older Office presumably does (T1 pins
both shapes), and it does not depend on which folder a caller asks for, as *(b)* would; *(c)* would
drop a promise the build can keep, and *(d)* would still have shown the pin. An Explorer counts as
existing when it is a registered pin OR the count did not go up, because a pin another session made
is reached through another apartment's proxy, which the registry does not hold. `LiveDisconnectRecoveryTests`
is unchanged. **Not measured:** the maintainer's Office build, and whether `ActiveExplorer()` can
return another session's hidden pin (`TODO.md`). **Undo:** revert the commit.

**Then measured, and decided a second time (same day).** With the fix in, the test still failed in
every full-suite run - two Explorers present, the pin's reference held - and passed when run alone.
Five more probes never reproduced the exit; a ninety-second subset run (the show-me tests plus this
one) did, and two scratch builds bisected it: without the session's Application Quit sink the test
passed, with it it failed. **Office LTSC 2024 raises `Quit` when the user closes the last VISIBLE
window, even though the hidden Explorer then keeps Outlook running.** The session hears it (SF-2),
the gateway drops it, and its `Dispose` - because that session STARTED this Outlook - closes the pin
("leave Outlook as you found it"), so Outlook ends. Alone, the test's pin belonged to another session,
which kept it; in the suite the session that re-started Outlook pinned and promoted itself.
*Options:* (i) keep honouring Quit - the user closing Outlook's last window is quitting it on Office
2024, and closing the pin on Quit is also the only thing that lets the user's own Exit end an Outlook
OutlookAI started (a pin left in place keeps it running, measured for the D49 dispose rule); (ii)
treat Quit as a hint and keep the pin until the process really exits - then the user's Exit would
leave Outlook running headless for as long as the server lives; (iii) drop the Quit sink - the same
cost as (ii) on Exit, plus SF-2's early release. **Decided: (i)**, and the test is scoped to it, not
loosened: when Outlook ends on the close, it passes only if the promoting session started Outlook,
the window was not the pin (two Explorers before the close), and the session was ended by the QUIT
EVENT - a process exit first would be a crash; the reattach that follows must still bring Outlook back
headless. Otherwise Outlook must survive, as before. Made observable by `OutlookComSession.GoneSignal`
and `ComGateway.LastSessionGoneSignal` (diagnostics, like `QuitSinkActive`). **What a user sees:**
closing the window OutlookAI showed ends an Outlook that OutlookAI itself started, as closing Outlook
would; OutlookAI starts it again, without a window, on its next request. **Undo:** revert the scoping
commit; the test then holds every build to "survives" and fails on Office 2024 when OutlookAI started
Outlook.

### 2026-10-03, autonomous - a subject override's conversation id outside Exchange: the promise is scoped, not dropped

**Primer.** `reply_draft`, `replyall_draft` and `forward_draft` take a `subject` override. Assigning
a subject makes Outlook regenerate the draft's conversation index, so since A3 the product restores
the child index and the source's topic after the rename (`conversationTopicPreserved`), and
`T2/LiveDraftOptionsTests` also held the renamed reply to its SOURCE's ConversationId. On the first
guest runs - a POP3 data file under Office LTSC 2024 - that one assertion failed every time (runbook
4.1e, F8), and three attempts to make the product keep the id were taken out again: restoring the
index-tracking flag, writing `PR_CONVERSATION_ID` back (refused: "does not support this operation")
and setting the subject as `PR_SUBJECT`.

**The measurement** (bisect run E4b, 2026-10-03, through `T2/ConversationIdHashes`): the renamed
reply's id is MD5 over the upper-cased KEPT topic in UTF-16LE - and not a hash of the new subject in
any of the four encodings tried; the seed's and the plain reply's ids are their index header's GUID
(bytes 6-21). So outside Exchange it is the topic the product restores that decides the id: the
override does not start a conversation of its own, and it does not keep the original's either.

**Options.** *(a)* Scope the same-id promise to Exchange and hold every other store to the measured
derivation. *(b)* Drop the id promise everywhere. *(c)* Keep it everywhere and leave the test red on
every data-file store. *(d)* Refuse the override outside Exchange.

**Decided: (a)**, by the coordinator on the maintainer's behalf. The id is the hash of the topic the
product keeps, not of the new subject, so outside Exchange the promise is not void - it is a
different, exact one, and *(b)* would throw away an assertion that pins it. **What changes for a
user** of a POP3/IMAP mailbox or a data file: a renamed derived draft keeps its index thread
(recipients' clients thread it as before) and its topic, but its `conversationId` is not its
source's, so a lookup by the source's id does not find it. The subject hint of the three tools says
so; Exchange is unchanged - its half was proven on the maintainer's workstation and no test guest can
re-measure it, so a store whose type cannot be read is held to that stricter promise.

**What was done** (on the maintainer's behalf, for review): `DerivedSubjectHint`, the result models'
comments and `McpServer/README.md` state it; `LiveDraftOptionsTests` asserts per store kind - the
source's id on Exchange, the kept topic's hash elsewhere - for the renamed reply and, newly, the
renamed forward. **Not measured:** `update_draft` renaming a reply draft takes the same restore path,
but no live test renames a derived draft through it. **Undo:** revert the commit; the test goes back to
one promise for every store and is red again on the guests.

### 2026-10-03, autonomous - Q11 ANSWERED by measurement: a table reports a date in the zone its column spelling asks for

**The measurement.** The first live runs on a test guest (`OutlookAI-Unindexed`, Office LTSC 2024
16.0.17932, W. Europe at UTC+2) ran `T2/LiveTableSortProbeTests`, which reads the same rows through a
table and through the opened items. Under the EXPLICIT built-in name `ReceivedTime` the raw table
value equalled the opened item's own LOCAL `ReceivedTime` (13:44:18 for an item received 11:44:18Z),
on both stores read; under the NAMESPACE reference `urn:schemas:httpmail:datereceived` the same
folder's rows read in UTC. So neither "UTC" nor "local" was right for every column.

**What it cost.** The exhaustive scan adds the explicit name first, and `ComDateValue.FromTableValue`
took every table value as UTC, so its resume cursor sat one offset LATE at UTC+2: run 1's
`LiveResumableScanTests.APagedScan` got ordinal 24 again on a page after one that had ended at
ordinal 7. West of UTC the same misreading moves the cursor EARLY, which skips mail and reports the
scan complete - the failure the question was raised about.

**What was done** (commit `b09041b`, on the maintainer's behalf, for review):
`ComDateValue.FromTableValue(value, columnProperty)` decides by the spelling - a namespace reference
(`urn:`, `http://`, `https://`) reads as UTC, an explicit name converts as an item value - and the
scan, the sort probe and the date-kind probe pass the spelling they added. The census
(`CensusTableRow.ReadUtc`) still reads every value as UTC: it compares its own readings with each
other, so a fingerprint taken through the explicit column is offset but consistent, and only the
instant it would print is wrong (`TODO.md`). Undo: revert `b09041b`; the old one-argument reading is
still there for namespace columns.


Answers move here with the date and the reasoning, so a future reader sees not just what was chosen
but why, and what the alternative was.

### 2026-08-18 - Q3 ANSWERED: no warn tier at all, fail only on what actually truncates

The question was whether `search` should sit at 87% of the cap. The maintainer rejected the framing,
and rightly: **"I want to fail the build the instant a change means something becomes too big. I want
to allow everything that fits without getting truncated. I want no warnings for something
approaching a limit."**

That kills the 75% warn tier outright. The argument for it was early notice on a silent cliff - the
server cannot detect its own truncation, so noticing before crossing is the only defence, and
`search` had once reached 3912 characters precisely because nothing flagged the growth. The argument
against is stronger: a warning that fires on three strings every single run, none of which will ever
change, is wallpaper. It trains everyone to ignore the channel, which makes it worse than nothing on
the day it matters.

Consequences, both following from "allow everything that fits":

- **The 75% warn tier is removed.** No approaching-a-limit output at all.
- **The house cap on parameter descriptions goes too.** Measurement established the client does not
  truncate them at any length - 20,000 characters arrive intact - so they always fit, and a rule that
  rejects text the client delivers whole is exactly what the maintainer ruled out. Sizes are still
  REPORTED, because that is the number a future per-tool bucket would be judged against, and the
  re-measure trigger stays documented.
- **What still fails the build:** a tool description or the server instructions exceeding 2048 UTF-16
  code units, which is the boundary the client was measured to cut at.

`search` at 1791 is therefore simply fine, and stops being flagged forever.

### 2026-08-18 - Q4 ANSWERED and shipped: per-store sweep windows for unscoped searches

Asked whether an unscoped search should pay roughly five extra index queries so every account gets a
window sized to its own index frontier, rather than one window from the profile-wide frontier.
Answered "proceed, completeness outranks cost", and shipped in `79c1827`.

Measured after the fact: **33-39 ms added per unscoped search** on a two-store catalog, against store
frontiers that sat **11 minutes 19 seconds apart** - so the single window really was eleven minutes
short for one store, on every search, silently. The larger half of the fix was not the map but the
fallback: a store missing from the index catalog used to inherit the profile frontier, the narrowest
window on the machine handed to the one store whose gap nobody had measured. It now gets the widest.

### 2026-08-18 - Q2 and Q6 ANSWERED by measurement, not by reasoning

**The questions.** Q2 asked whether Claude Code's documented "truncates tool descriptions and server
instructions at 2KB each" could be trusted at all, and in which unit - characters or UTF-8 bytes -
since the guardrail hedged by measuring both and failing on the larger. Q6 asked two things the same
sentence leaves open: whether `inputSchema.properties[*].description` is capped at all, and whether
the 2 KB is per STRING or per serialized tool. Q6's second half was the one that mattered, because
under the per-tool reading the 2026-08-17 trim - moving `search` detail out of the description and
onto its arguments - would have moved text from one capped bucket into the same capped bucket, and a
fix reported as solving the problem would have solved nothing.

**The evidence.** An interception experiment against Claude Code `2.1.234` on Windows 11, run
2026-08-18: a local HTTP endpoint stood in for the API (`ANTHROPIC_BASE_URL` plus a throwaway
`ANTHROPIC_AUTH_TOKEN`), captured the client's outbound `POST /v1/messages`, and the `tools` array
the model actually receives was read byte for byte. Reproduced against two models, byte-identical,
because the cut is client-side. Not a model's recollection of what it received - the wire.

**The answers.**

1. **Per string. There is no per-tool bucket at all.** A probe entry of 17,411 bytes and another of
   20,172 bytes both arrived intact, as did 202 tools totalling 348,314 bytes of serialized entries
   in one request. **So Q6's dangerous reading is disproved and the `search` trim was valid** - it
   moved text out of a capped string into an uncapped one. (348 KB establishes no cap at 348 KB, not
   that none exists above it.)
2. **UTF-16 code units, never bytes.** A 2,048-character description weighing 6,004 UTF-8 bytes
   arrived whole, and two strings of very different byte lengths were cut at the same CHARACTER
   offset. Units rather than code points, and the cut is surrogate-aware (2,047 rather than splitting
   a pair).
3. **Parameter descriptions are not capped at any length.** 20,000 characters through intact. The
   documentation's silence about `inputSchema` is accurate, not an omission.
4. **Boundary:** cut when `length > 2048`, so exactly 2,048 passes and 2,049 does not - measured as a
   triple in one run. A cut string reaches the model as its exact prefix plus a 13-unit marker
   (U+2026 HORIZONTAL ELLIPSIS, space, `[truncated]`), so 2,061 units in total.
5. **The marker is invisible to us.** It is appended after our JSON-RPC response has left, with no
   error, notification or re-request: **a server cannot detect its own truncation**, and no test in
   this repo ever will. A model can, so "did that arrive whole?" is answerable by asking and
   unanswerable by logging.

**What changed as a result.** `DescriptionBudgetCiTests` now measures UTF-16 code units alone
(`string.Length`) instead of `max(chars, UTF-8 bytes)`. Failing on bytes could only ever produce
FALSE failures - it rejected text the client delivers whole - and on today's surface it changed
nothing at all: all 132 wire strings are pure ASCII, so no measured size moved and the warn tier is
unchanged (`search` 1791, `update_draft` 1593). The guard is now right rather than accidentally
harmless. The 2048 applied to parameter descriptions was **kept but relabelled** as
`HouseParameterBudget`, a separate constant with its own reasoning: it floats with a client version
we do not control and get no signal about, and `BodyHtmlHint` is one constant reused across five
drafting tools, so one over-long shared parameter description would be five silent truncations the
day a release starts cutting schemas. **That last part is superseded by Q3 above** - the maintainer
ruled that a limit which rejects text the client delivers whole is a false failure whatever the
future risk, so `HouseParameterBudget` is gone and parameter sizes are reported without a budget;
the re-measure trigger it was guarding is documented rather than enforced. The rest of this record
stands. The marker is recorded as `ClientTruncationMarker` - not as a
detector, which is impossible, but because it is the string a human greps for in a transcript when
something looks cut - and the guard now also fails if a shipped description ever CONTAINS it, which
would mean already-truncated text was copied back into source.

**The caveat that replaces the old uncertainty.** This is one client at one version, and nothing
watches it: no version header, no notification, no server-side signal when it changes. The number is
only as current as its date. Re-measure at client-bump time; the change worth re-measuring for is a
release that introduces a per-tool bucket, which would cut large schemas on day one and silently.

### 2026-08-18, autonomous - a measured defect jumped the queue

The overnight sweep measurements found that DASL date literals are emitted as `MM/dd/yyyy` while
Outlook parses them in the machine locale, which here is day-first. On any date whose day is 12 or
lower - about 40% of days - the day and month swap silently. Measured consequences: an `exhaustive`
search for 1-5 August returned 48 items from April and May; a sweep window starting 5 September was
read as 9 May, blew the 30 s budget and killed the COM host; and a 7-day empty-index window opened
today would be read as a future date, so the sweep selects nothing while reporting `foldersSwept: 4`
and `freshness: "live"`.

I moved this ahead of the four fixes you approved, without asking, because it produces silently wrong
search results and the alternative was leaving it in place for hours. If you would have sequenced it
differently, that is the call to correct.

### 2026-08-18, autonomous - three sweep constants kept, with evidence

`SweepSafetyMargin` (10 min), `EmptyIndexSweepWindow` (7 days) and `SweepPerFolderCap` (200) were all
marked "Open - needs measurement". All three are now **Kept - defensible**, measured over 43 sweep
samples and 177 index-frontier probes on the real profile; the numbers and their spread are in
`Docs/magic-numbers.md`. Two honest gaps are recorded there rather than papered over: the 7-day
window's cost is a prediction from a measured cost model rather than an observed sweep, because the
window cannot be widened through the shipped tools; and indexing latency could only be sampled during
one overnight hour, so its spread is a floor rather than the whole picture.
