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

**Why it matters.** At 500 per folder, a 4,918-item Sent Items and a 100,000-item Archive are
counted and not identified. In those folders the guard can see that items left but not which ones,
cannot tell a deletion from a filing, and is blind to a deletion masked by an arrival. Those are
exactly the folders a runaway test would do the most damage in.

**Options.** *(a)* Leave 500/3,000 - the guard covers the folders a person works in, and everything
else is still count-guarded. *(b)* Raise to about 5,000 per folder and 25,000 per store, which
covers Sent Items and most working folders and costs roughly 125 table calls per store. *(c)* Raise
the per-folder limit but keep a tight per-store budget, so one huge folder cannot consume the whole
allowance. *(d)* Raise only for non-delegate stores, on the reasoning that delegate stores are the
slow ones and the least likely to be written to. *(e)* Identify everything, no budget - honest but
unbounded, and a 100,000-item Archive is roughly 550 table calls and tens of MB of EntryIDs held
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

### 2026-11-18 - the maintainer's answers to Q138, Q163 to Q182 and Q194 to Q215

**What this is.** Forty-two questions went to the maintainer on 2026-11-18, and he sent 41 one-line
answers the same day; two more were asked and answered that day too: Q214, raised by his answer to
Q173, and Q215, raised by the work on Q166. The questions as asked, each with its primer, context,
directions and recommendation, are in `.work\open-questions-2026-11-18.md` of his checkout, which
is gitignored, so not in a clone; what the decisions rest on is restated here, with placeholders
where the questions named real mailboxes. Words in quotes are his, verbatim, spelling included.
Where he answered "rec", the recommendation he had been given was taken, and it is spelled out.
Five of the 42 did not end in a plain one-line decision:
- **Q172 and Q173** drew questions back from him, and were decided the same day after follow-ups;
  the exchanges are quoted below.
- **Q197 and Q199**, which concern CodeQL alone, are superseded by Q173 (b). Q197's line, which
  arrived as *"ec"*, needs no confirming now.
- **Q183** asked for an action of his, not a decision; it was done the same day.

One decision was taken on his behalf the same day, under his answer to Q131 and Q132; it is
recorded with Q207.

#### Code sharing - Q163 to Q172, Q215

**Primer.** Q154 (2026-11-17) made "no code duplicated" a rule and let the add-in reference
OutlookAI.Core's net48 build "behind a small interface". A read-only survey of master 9f4646d found
73 groups of duplicated code (`.work\dup-survey-2026-11\` of his checkout). These ten asked how the
refactor goes, and Q215 followed from one of them; `TODO.md` holds the refactor.

- **Q163 - logic only the add-in runs.** Some pure code is used by the add-in alone: the update
  schedule, the decision whether to register the MCP server, the prompt checks, the settings
  reconciler. The tests cannot reference the add-in, so they compile those source files in.
  *"163 a"*. **Decided: (a)**: it moves into Core, in a section of its own for the add-in - one
  reference and no shared source files, though Core is then no longer purely neutral between the
  server and the add-in.
- **Q164 - what "behind a small interface" means.** Core also holds code that must never run inside
  Outlook's own process, where the add-in runs - its Outlook connection, its COM thread, the index
  search - because a crash there takes Outlook down. *"164 a"*. **Decided: (a)**: one public part of
  Core that the add-in may use, held to that by a text check; nothing in it connects to Outlook,
  starts a COM thread or opens the index. No C# interface and factory, and no bridge DLL.
- **Q165 - which copy wins where the copies have drifted apart.** Core's copies are newer and
  live-tested: they release COM objects the add-in leaks, and embed signature pictures the add-in
  probably leaves as file links. One difference shows: Core writes a line break as a soft break, the
  add-in's AI pane as a new paragraph. *"165 c"*. **Decided: (c)**: Core wins for robustness fixes;
  a difference a person can see is measured on a test machine first.
- **Q166 - how far the rule reaches into PowerShell.** The scripts copy helpers on purpose - the
  worst case is one wrapper, copied 15 times - because some scripts travel alone to a VM (the setup
  disk carries one file), and the Windows PowerShell 5.1 check sees only helpers in the same file.
  *"166 b"*. **Decided: (b)**: a module for the scripts that travel with others, and a check that
  fails when a copy in a script that travels alone differs from the module. The "module" became a
  file the scripts dot-source: stream R2 measured that under Windows PowerShell 5.1 the helpers
  cannot be a .psm1 module, because the module's scope loses the callers' error settings. Where the
  file lives is Q215.
- **Q215 - where Q166 (b)'s shared file lives.** A new question, raised by stream R2's work on
  Q166 (b) and answered the same day. Directions: (a) `Testbed/guest/` - the routes that copy
  scripts to the VMs stay as they are, and `Tools/` starts depending on `Testbed/`; (b) `Tools/`,
  with every route that copies scripts to a guest also copying the file; (c) a new top-level folder,
  delivered as in (b); (d) only the host's scripts share it, for now; (e) two files kept identical.
  *"215 rec"*. **Decided: the recommendation, (a)**: the file lives in `Testbed/guest/`.
- **Q167 - how the add-in's build gets Core.** The add-in is an old-style Office project, built only
  by the two scripts that keep it from registering itself in his Outlook (Q81). *"167 a"*.
  **Decided: (a)**: a project reference, as the add-in had to PolicyWriter until Q145 (a), with
  package restore added to both scripts, and Core added to the Visual Studio solution so that F5
  still works. That changes `AGENTS.md`'s rule that the solution holds only the add-in; the stream
  that builds it changes the rule.
- **Q168 - testing the add-in's AI pane.** Nothing tests the pane today - the test VMs only prove
  that the add-in loads - and its signature and draft code is moving into Core. *"168 b"*.
  **Decided: (b)**, over the recommended (c) plus (a): the pane is automated on a test VM - real
  coverage, at the cost of a new automation entry point in the shipped add-in.
- **Q169 - where shared test code lives.** One test helper exists 18 times, 17 of them identical,
  and another 13 times; the repair tools (`OutlookAI.RemediationTools`) share code with the tests
  too. *"169 rec (your recommendation)"*. **Decided: the recommendation**: (c), a shared folder in
  the test project for code only the tests use and a test-support project for what the tools share
  with them - or only (a), the folder, if the tools share no code with the tests once Q170 (a) has
  deleted their incident commands.
- **Q170 - the repair tools' incident commands.** Beside the corpus commands the test VMs use,
  `OutlookAI.RemediationTools` still has `audit`, `refile`, `purge` and `dedupe` from the July 2026
  incident, and `purge` and `dedupe` delete mail. They run on an old copy of the test mailer that has
  drifted: its purge skips the Sync Issues folders the safety rules require. The recommendation
  rested on his confirming that the incident is closed. *"170 rec"*. **Decided: the recommendation,
  (a)**, with that confirmation: the four commands and their code are deleted, and the corpus
  tooling stays. The personal-data clean-up carried it out the same day (`CHANGELOG.md`, Removed).
- **Q171 - copies across languages.** PowerShell restates rules the C# holds - account
  classification, the test tags, the updater's signature check - and the installer and some
  documents repeat constants that an existing check keeps equal. *"171 rec"*. **Decided: the
  recommendation**: (c), C# source files both sides compile, as one file already is, for the Windows
  interop code and the signature check; (a), checked copies, for the installer, the documents and
  the build files.
- **Q172 - the add-in's C# version, and then every project's .NET.** The add-in targets .NET
  Framework 4.8 and compiles as C# 7.3 by default; the server's projects are on .NET 10, and Core
  is built for both. Core's public code works with 7.3. Directions: (a) keep 7.3 during the
  refactor and revisit after, recommended; (b) the latest version; (c) a newer pinned version with
  null checks on. His answer reached past the add-in: *"172 move everything to the latest possible
  .net or .net core version (beware that .net 11 just released)"*. A follow-up went to him, and he
  asked back: *"172 I tought we had a aot bianry with the framework (or at least parts of it)
  included so that we would not have to install a framework on the machines?"* The answer was no:
  the server has been framework-dependent since 2026-08-08, because of the updater's 50 MB download
  cap, which a self-contained server would pass (`McpServer/README.md`). *"172 why is there a 50 MB
  limit? That is silly. Double check this."* It is an arbitrary sanity limit from the bug-fix batch
  of 2026-06-01 (0021ca1), compiled into every installed copy (`Services/UpdateService.cs`). A
  second follow-up went to him, and he answered: *"172 can we go for option c but instead of
  shipping it along and breaking our stupid file size limit we have the installer download it on
  demand?"* **Decided**, all on 2026-11-18: everything that can moves to .NET 11 and C# 15, and the
  VSTO add-in, which cannot run on modern .NET, stays on .NET Framework 4.8 but compiles as C# 15;
  the shipped MCP server runs on .NET 11 with a private copy of the runtime, which the installer
  downloads on demand - Microsoft's official runtime package, checked against its published hash -
  and unpacks into OutlookAI's own per-user folder, so it needs no administrator rights and a silent
  update does it too; and the installer itself stays under the 50 MB cap, which is unchanged.
  Stream U builds it.

#### The licence and CodeQL - Q173, Q214, Q197 to Q199

**Primer.** On 2026-10-01 the maintainer decided to move OutlookAI to the FSL-based licence BrowserAI
carries; it is an open `TODO.md` item. On 2026-10-03 (Q125) he allowed CodeQL in the release
checks, and CodeQL's licence permits it only on code under an OSI-approved open-source licence,
which FSL is not: `Tools/Invoke-CodeQL.ps1` refuses unless `LICENSE` is MIT (`AGENTS.md`,
Dependencies, the third exception). Since Q146 the scan treats every MCP tool's arguments as
untrusted input.

- **Q173 - the licence change rules out CodeQL.** Directions: (a) stay on MIT and keep CodeQL; (b)
  move to FSL and drop CodeQL from the release checks, keeping the .NET security analysers on every
  build and the vulnerable-package check - recommended, if FSL still stands; (c) move to FSL and find
  a free scanner whose licence allows it, which needs research and an exception to the Dependencies
  rule; (d) decide when the move actually happens. His first answer was a question back: *"173 what
  does CodeQL bring us? I'd rather not let something steer my license but I have no clue what it's
  loss will cost us. Why was it added?"* Why it was added is in this log's Q125 record (2026-10-04);
  what it found, in the Q146 record (2026-10-11) and under Q198 below. A follow-up went to him the
  same day, and he answered: *"173 b"*. **Decided: (b)**: OutlookAI moves to the FSL licence, and
  CodeQL leaves the release checks. In its place a build rule keeps the risky calls - file writes,
  the search query, program starts - inside a few vetted wrapper classes that check their input; the
  .NET security analysers on every build and the vulnerable-package check stay. Stream S builds it,
  and the rules it changes in `AGENTS.md` change with it. It supersedes Q197 and Q199, and raised
  Q214.
- **Q214 - from when the FSL licence applies.** A new question, raised by Q173 (b) and answered the
  same day. Every release so far went out under MIT, and what was published under MIT stays MIT for
  whoever has it. Directions: (a) from the commit that changes the licence; (b) from the next
  release; (c) from a fixed date. *"214 a"*. **Decided: (a)**: everything after the licence commit,
  the next release included, is under the FSL, and the release notes say so. Stream S builds it.
- **Q197 - should CodeQL follow an AI's input past the tools themselves. Superseded by Q173 (b).**
  CodeQL loses track of an AI's arguments in two places, so what happens to them after those is not
  scanned: where the work crosses into the separate process that talks to Outlook, through a channel
  CodeQL cannot follow, and at list arguments such as item IDs and attachments. Directions: (a) one
  CodeQL model entry per operation of that process, 27 today, counted by the census that already
  guards the tools; (b) a small custom CodeQL library, which also covers the list arguments; (c)
  leave it, with the gaps written down. Recommended: (a), after Q198 - a trial run with those
  entries found 3 new results, all in `save_attachment`, which would block every release until Q198
  was decided. His line arrived as *"ec"*, most likely "197 rec" with its start lost. With CodeQL
  leaving the release checks there is nothing left to decide, and the line needs no confirming.
- **Q198 - where `save_attachment` may save files.** It saves an attachment into any folder the AI
  names. The sender chooses the file and its name, and the AI, which reads the sender's mail, chooses
  the folder, so a malicious mail could talk an AI into saving its attachment into the Startup
  folder, where it would run at the next sign-in. Whether Outlook's blocking of dangerous attachment
  types stops that save is not measured. *"198 rec"*. **Decided: the recommendation, (a)**: only
  OutlookAI's own download folder, `%LOCALAPPDATA%\OutlookAI\scratch\attachments`, and folders under
  it. An AI can still copy the file elsewhere with its own tools, under their own permission rules.
  It stands without CodeQL, whose scan found the hole.
- **Q199 - how strictly an accepted CodeQL finding is matched. Superseded by Q173 (b).** A release
  passes CodeQL only when every finding is fixed or accepted with a written reason in
  `Tools/codeql-accepted.json`, and an acceptance matches on the rule, the file and the line - so a
  new source of untrusted input reaching the same line is accepted silently, its reason never
  re-checked. The accepted search-query finding already covers 7 sources. *"199 c"*: each acceptance
  would have recorded every source's exact location, computed by the script, over the recommended
  (b), which records only how many sources a finding lists. The acceptance file goes with CodeQL, so
  nothing of it is built.

#### How drafts are written - Q208 to Q211

**Primer.** Q149 (d), 2026-11-17: measure first, then choose between keeping Word for drafts, with
pins only while drafting, and writing them as HTML with no compose window and no pins - HTML only if
its drafts are indistinguishable from today's. The measurement ran that day on the four Outlook test
guests (`Docs/research/q149-draft-text`): no HTML-only draft looked the same as today's, so drafts
keep going through Word, and a window-less Outlook needs a pin for as long as they do - without one,
the compose window's close ended Outlook in 40 of 40 drafts. Q175 to Q179 and Q118's pin clean-up
depend on Q208.

- **Q208 - how Word-based drafting works from now on.** Directions: (A) as approved, a pin only while
  a draft is being written - closing it after each draft ends Outlook, so every draft starts Outlook
  from cold and the compose window flashes every time; (B) as measured, one pin for the whole
  session, with Q118's clean-up so that pins are not left behind, one visible moment per Outlook
  start instead of per draft, and no per-draft activation at all while his Outlook window is open;
  (C) HTML-only after all, which fails his "identical" rule today; (D) B now, then retry HTML-only
  with a third version that copies today's markup, once Q209 fixes replies - about an hour of
  test-machine time; (E) judge HTML-only against what Outlook's own compose window produces instead
  of today's drafts, not measured. *"208 rec"*. **Decided: the recommendation, B**, together with
  Q209's fix. B carries Q118 (a)'s pin clean-up, approved 2026-10-11 and on hold until now.
- **Q209 - fixing replies and forwards.** A new reply starts with two empty paragraphs above the
  quoted message's header block; today's code writes into the header block itself, so the text and
  the signature end up inside the quoted header - 12 of 12 measured. *"209 rec"*. **Decided: the
  recommendation, (a)**: the text and the signature go into those top paragraphs, with a test that
  the first paragraph has no top border.
- **Q210 - the measurement suite.** The tests that produced Q149's numbers run only when selected on
  their own, so they cost nothing in normal runs. *"210 rec"*. **Decided: the recommendation, (a)**:
  all of it stays in the repository. Results already differ between Office 2021 and 2024, so it will
  be needed again.
- **Q211 - the spacing in drafts.** Today's draft paragraphs use Word's "Normal (Web)" style, with
  5 pt of space above and below. *"211 rec"*. **Decided: the recommendation, (b)**: first measure
  what Outlook's own compose window does with typed text.

#### The Outlook hang of 2026-11-15 - Q174 to Q182

**Primer.** On 2026-11-15 an OutlookAI MCP server on the maintainer's workstation started Outlook
through COM, with no window, while Outlook was closed. About 5 seconds into its start that Outlook
deadlocked - inferred: OutlookAI's connect sequence and the Teams Meeting add-in's start-up call met
on Outlook's main thread - never showed a window, and Outlook would not start for him until the
process was ended 34 minutes later (`TODO.md`, the item on the hang). What master has done since
reaches his PC only with a release.

- **Q174 - his installed OutlookAI until his next release.** *"174 a"*. **Decided: (a)**: it stays
  as it is, and the installed MCP server is not disabled. The recommended (c) would also have put the
  prevention (Q175, Q178, Q179) first, so that his next release carries it. Releases stay his call.
- **Q175 - how OutlookAI starts Outlook, or attaches to one that is still starting.** Today it starts
  Outlook with no window (D17, D33) and opens the mailboxes within a second - exactly when the
  deadlock struck. Q150 (a), restarting Outlook on the next call after he closes it, stands.
  Directions: (a) wait for Outlook to settle before opening any mailbox - Outlook at least about 20 s
  old, the add-in's start-up stamp present, its main thread idle - which narrows the race, is
  unproven until measured, and makes cold starts on the test VMs about 20 s slower; (b) (a), plus no
  hidden pin window at connect, which removes two more mailbox-opening calls and depends on how Q208
  decides the pin; (c) start Outlook minimised, as he does, reversing D33; (d) never start Outlook,
  refuse instead, reversing D17 and Q150 (a); (e) only one OutlookAI server at a time may start
  Outlook, and the others wait - eight were running during the hang. *"175 rec"*. **Decided: the
  recommendation**: (a) and (e) now, and (b) once Q208 decides the pin - Q208 is B, above. D33 stays
  unless the measurement shows that only (c) avoids the deadlock.
- **Q176 - noticing an Outlook that starts and never finishes, or stops answering.** Master asks
  Windows whether Outlook's windows are frozen. That would have caught this hang, but not one that
  strikes before Outlook creates its windows, which reads as "starting" forever. *"176 c - we can
  then report back to the calling llm that outlook is stuck and why. Allowing for the user to report
  back to us or ask the llm to fix it."* **Decided: (c)**: a time limit on "starting", about 120 s,
  and a per-user record of the Outlook that OutlookAI started - its process, its start time, which
  server started it, and which start-up step it has been stuck in, since when. With his requirement:
  the calling AI is told that Outlook is stuck and why, so that the user can report it to the
  project or ask the AI to fix it.
- **Q177 - may OutlookAI ever end a hung Outlook.** A deadlocked Outlook ignores Quit and window
  closes; only ending the process helps. OutlookAI never stops Outlook (S7, D17), and the mailbox
  safety rules forbid killing OUTLOOK.EXE. *"177 a - ensure to communicate back the risks of killing
  it when communicating back the hang."* **Decided: (a)**: never; OutlookAI reports the hang clearly,
  and the user ends Outlook in Task Manager. With his requirement: the report of a hang also says
  what ending Outlook that way risks.
- **Q178 - what the agent and the user are told about a hung Outlook.** For a hung Outlook, master
  tells the agent "restarting Outlook fixes it now ... this also clears by itself". For an Outlook
  with no window that is untrue: the user can neither see nor close it, launching Outlook does
  nothing, and a deadlock never clears by itself. *"178 rec"*. **Decided: the recommendation, (e)**:
  different text when Outlook is hung and has no window - the Task Manager steps, and no "clears by
  itself"; health fields for the process, its start time, who started it, whether it ever showed a
  window, frozen since when and which start-up step; a start log under `%LOCALAPPDATA%\OutlookAI`,
  with no mail content, which would have named the stuck call here; and a troubleshooting entry in
  the README.
- **Q179 - only a live request may start Outlook.** After his Outlook was ended, the old version
  started it again for a request abandoned 12 minutes earlier. Master can do the same another way: a
  cancelled request switches off its own hang timer, the process that makes the Outlook calls (the
  COM host) then retries once, and that retry starts Outlook. *"179 rec"*. **Decided: the
  recommendation, (e)**: the hang timer keeps running after a cancel, as the code's own description
  already says it does, and the COM host's retry may only attach to a running Outlook, never start
  one.
- **Q180 - recording how Outlook starts during measurement runs.** The offline test VMs cannot
  reproduce this deadlock - no Exchange mailbox cache, no Teams - but every measurement run there
  cold-starts Outlook many times. *"180 rec"*. **Decided: the recommendation, (b)**: the next
  measurement run there logs every cold start's time to ready, its "alive" verdicts and any hang -
  about 100 data points for nothing.
- **Q181 - reproducing the hang.** Only the Exchange VM has an Exchange mailbox cache, and it runs
  Office 2024; his own Office 2021 runs only on his workstation, where no live test may run (Q116),
  and no test VM has Teams. *"181 d"*. **Decided: (d)**, over the recommended (a) plus (d): no
  reproduction - the prevention and the start log ship, and what follows is watched. No stand-in
  add-in on the Exchange VM making Teams' kind of call.
- **Q182 - the add-in's Claude warm-up.** At every Outlook start the add-in runs a full Claude Code
  session from inside OUTLOOK.EXE to "warm up" the command line: one model request with the prompt
  "-", under the user's global Claude settings, run from `C:\Windows\system32` when a program started
  Outlook. It does not stay warm - it runs and exits - and about 656 such sessions were on disk. It
  played no part in the hang. *"182 a"*. **Decided: (a)**: the warm-up is removed.

#### The Exchange VM - Q206, Q207, Q212, Q194, Q138

**Primer.** `OutlookAI-Exchange` is the one test VM signed in to Microsoft 365 (Q108 to Q111): the
dedicated test account, its writable hub (Q137), and a real mailbox, a read-only bystander. Runbook
section 4.4a has its runs.

- **Q206 - getting back Office's ability to create items there.** Its Office LTSC 2024 was installed
  on 2026-10-03 and never activated, because no activation server answers it, and its grace ran out
  about 25 days later. Since then Outlook there refuses to create any item through its object model:
  41 tests of the run of 2026-11-17 failed at their first draft or mail, and every read test passed.
  On 2026-10-08, inside the grace, the same tests created and removed their items. The licence is the
  most likely cause, not proven: proving it means creating an item, which only the tested suite may
  do. Directions: (a) reset Office's grace with its own reset tool, `ospprearm.exe` - free, but only a
  limited number of times, its licence terms to be checked, and the cheapest proof of the cause; (b)
  install Microsoft 365 Apps there, activated by the test account's own licence if the tenant has one
  for it - a licensed component, so an exception to the Dependencies rule, and it would let real
  roaming signatures be measured, which the LTSC version cannot; (c) activate the current Office with
  a volume key, MAK or KMS - a licence key; (d) only the tests that write nothing, until one of these
  lands. Recommended: (a) once, to prove the cause, then (b) if the test account has a Microsoft 365
  Apps licence, else (d). *"206 a with fallback to b with fallback to reporting to me for spla mak
  keys"*. **Decided:** (a) first; if the reset does not bring item creation back, (b); and if (b) is
  not possible, an agent reports to him, and he supplies a volume key (a SPLA MAK). `AGENTS.md`
  (Dependencies, the second exception) carries the fallbacks. Until one of them works, the VM runs
  only the tests that write nothing, as it has since 2026-11-17.
- **Q207 - a mailbox shown twice makes `outlook_health` report "degraded".** The test account's
  mailbox is in the VM's Outlook twice: as its own cached mailbox, which Windows Search indexes, and
  as an online-only entry Exchange added automatically because the bystander mailbox has Full Access
  to it, which is never indexed. So `outlook_health` reports "degraded" - a mailbox with nothing in
  the index, and searches falling back to a slower 7-day sweep - which is true of that entry and
  misleading, because the mailbox is fully indexed through its cached entry. Both entries carry the
  same internal ID. Any user whose account has full access to another mailbox in the same profile
  gets this, and the health test fails on the VM because of it. Directions: (a) switch the automatic
  adding off in Microsoft 365, or remove the Full Access - his tenant's administrator, a colleague -
  which also rules that entry out as a stand-in delegate mailbox in tests; (b) in the product,
  `outlook_health` and search recognise an entry with the same ID as an indexed mailbox as that
  mailbox: health calls it a second view of an indexed mailbox, not a degradation, and a search
  covers it once, through the indexed entry, instead of also sweeping the same mail live; (c) in the
  tests only, the health test accepts "degraded" exactly when every unindexed entry is such a second
  view; (d) leave the test failing there, or exclude it. Recommended: (b), and (c) meanwhile if (b)
  has to wait. *"207 b"*. **Decided: (b).** The stopgap (c) was not part of his answer. Stream K
  builds (b) together with the decision that follows.
- **Decided on his behalf, 2026-11-18: one mailbox, one store ID** - under his answer to Q131 and
  Q132, *"what you think is the most stable"*. Run 6 on the VM (runbook section 4.4a) found the
  duplicate breaking folder IDs too: the test account's `store_id` named both entries, so
  `list_folders` listed both folder trees, 74 folders, and `find_folder` found two Inboxes - a defect
  under Q131 and Q132 whatever Q207 decided. **Decided:** a mailbox that appears as two store
  entries - its cached store and an automapped online entry - gets one store ID and one folder tree,
  resolved through its cached (indexed) entry, and no ID resolves to two stores. Stream K builds it,
  with Q207 (b).
- **Q212 - may a folder search on an Exchange mailbox start Outlook while it is closed.** For a local
  mail file the search index finds a folder by its ID without Outlook; for Exchange it cannot
  (measured, runbook section 4.4a), so only Outlook can find the folder. As built, OutlookAI starts
  Outlook without a window to answer, as it does for other tools under D17 and Q150 (a), and refuses
  only when Outlook does not respond. Q132 could be read as "always refuse while Outlook is closed",
  and a start from closed is what preceded the 2026-11-15 hang. *"212 rec"*. **Decided: the
  recommendation, (a)**, as built: start Outlook, answer correctly, and refuse only when Outlook does
  not respond. Q175's start protections apply to it too.
- **Q194 - the roaming-signature test with Office's online features on.** R3, on 2026-11-17, made a
  test signature the test account's default in Outlook on the web and watched whether the VM's
  Outlook, Office LTSC 2024, picked it up. It never did - not after 1 minute, 15 minutes or an hour,
  nor after an Outlook restart - and no local copy appeared; the signature was deleted afterwards,
  and no real mailbox's signatures were touched. Two things could have hidden roaming: the testbed's
  policy switching Office's online features off, and File > Options greyed out there (probably the
  expired activation, Q206), so that "store my settings in the cloud" could not be checked.
  *"194 rec"*. **Decided: the recommendation, (b)**: R3 runs again with the four policy values
  removed - a change to that VM's registry only, undone by the checkpoint R3 restores anyway. While
  it lasts, Office there talks more to Microsoft's online services.
- **Q138 - the one old test message in the bystander mailbox's Sent Items.** It is left from the
  years the live tests ran on his workstation: `[OutlookAI-McpTest] d3994a2ad25309442 d39 move seed`,
  sent 2026-07-26 03:46, and a read-only scan found it the only tagged item in either mailbox on the
  VM. Under Q138 (c) (2026-10-11) a tested helper deletes exactly this item - its item ID and its
  test tag both matched - then removes it from Deleted Items, but only after he confirms.
  *"138 rec"*. **Decided: the recommendation, (a)**: delete it - his confirmation, given on the item
  as the question described it, by its subject and the time it was sent. The deletion is a write, so
  it waits until Q206 gives Office there back the ability to change items.

#### Signatures - Q200 to Q205

**Primer.** Q147 and Q152 to Q162 (2026-11-17) shaped how OutlookAI handles signatures, and
building it raised these six (this log, "Q147 carried out", 2026-11-17).

- **Q200 - finding out what Outlook does when an administrator forces a signature.** Q156 allows
  another signature, or none, under a forced default only if Outlook allows it; until that is
  measured, OutlookAI refuses both. The test machines run without administrator rights, so they
  cannot set such a policy today. *"200 a"*. **Decided: (a)**: a runner step sets the policy as
  administrator before Outlook starts, and the VM's checkpoint undoes it afterwards; a live test then
  checks what Outlook inserts. It uses the kind of new-mail window that crashed Outlook on
  2026-10-04, so it runs with crash dumps on. It answers Q201 too.
- **Q201 - where a default signature set in Outlook's own mail settings ranks.** Besides an
  administrator's policy, Outlook's general mail settings can name a default signature, and the
  sources disagree on whether Outlook puts it above or below an account's own default. Directions:
  (a) above the account's default, as built - default changes refused, a per-message choice
  allowed; (b) below it; (c) ignore it; (d) show it, but act as in (b). Recommended: (a) until Q200
  measures it, then follow the measurement. *"201 meassure it"*. **Decided:** measure it, with
  Q200 (a)'s runner step; (a), as built, stays until the measurement is in, and the measurement then
  decides.
- **Q202 - reading signatures stored in the mailbox itself.** Microsoft 365 can keep signatures
  inside the mailbox ("roaming"); OutlookAI uses Outlook's local copies and detects roaming (Q152).
  R3 (Q194, above) found nothing arriving on Office LTSC 2024, and no hidden local copy. *"202 rec"*.
  **Decided: the recommendation, (a)**: the mailbox is never read, as built - not through Outlook's
  hidden local copy, whose format is undocumented, and not through Outlook on the web's private
  interface.
- **Q203 - a named signature on a reply from a shared or delegate mailbox.** Q160 said no signature
  there for now, and the result says so. As built, no default goes in, but a signature the AI names
  explicitly does. *"203 rec"*. **Decided: the recommendation**: (a), as built, now; then (d) - once
  the shared test mailbox (Q109) exists, measure what Outlook does there and match it. `TODO.md`
  holds (d).
- **Q204 - how far an administrator's "no signatures" policy stops signature editing.** What is
  built is stricter than Q156 asked: while policy disables signatures, every signature change
  through OutlookAI is refused, and a signature the policy forces cannot be deleted. *"204 a"*.
  **Decided: (a)**, as built: Outlook's own signature editor is closed under that policy too, and
  deleting a forced signature would leave the forced default naming nothing.
- **Q205 - updating one format of a signature.** Each signature has three files: HTML, Rich Text and
  plain text. As built, an update writes only the formats it is given, and the answer names the
  formats it left unchanged - nothing is converted, but the formats can disagree until the AI sends
  the others. *"205 d"*. **Decided: (d)**, over the recommended (a): an update is refused unless it
  sends every format the signature already has.

#### Personal data and the leaked signing key - Q196, Q213, Q183

- **Q196 - committing the personal-data check's fingerprints.** The check, on the personal-data
  branch, finds the maintainer's identifiers, and third parties', without storing the values: it
  stores salted fingerprints (hashes) of them in a committed file. The salt is public, so anyone can
  test guesses against them - the agent recovered his computer name from 1,000 guesses, a million
  take 2 seconds, and word lists of Dutch words, first names or .nl domains would expose the third
  parties' values the same way. *"196 c"*. **Decided: the recommendation, (c)**: no fingerprint is
  committed; the build-VM runner copies the private list from his workstation into each run - the VM
  has no internet and is reset after every run. Nothing guessable is published, and the build VM
  keeps checking every commit. The personal-data clean-up built it, with its rule in `AGENTS.md`
  (Personal data); the next entry records it too.
- **Q213 - silent updates and the leaked certificate in a user's Trusted Root store.** Windows makes
  the user confirm every removal from their own Trusted Root store, with No as the default answer.
  The installer explains and asks on an interactive install or uninstall; auto-updates run silently
  right after Outlook closes, so they cannot ask without surprising the user. Only machines where
  someone added the leaked certificate to Trusted Root by hand are affected - no release ever did
  that. Measured on a test VM (runbook section 4.8): every removal asked, 13 of 13, even with the
  removal tool hidden, and an unanswered question still waited after 120 s; deleting the
  certificate's registry key directly skips the question, but Windows keeps its record of the
  certificate as one the user confirmed, so writing the key back trusts it again with no question.
  Directions: (a) skip it on silent runs, and the next interactive install or uninstall asks - as
  built, on master since e8ee8d0; (b) delete the registry key silently; (c) ask without waiting; (d)
  the add-in explains and asks at the next Outlook start; (e) (a), plus a line in the release notes
  on how to check for the certificate and remove it - recommended. *"213 a only"*. **Decided: (a)
  only**, as built: no release-notes line, and no question from the add-in. A machine that only
  auto-updates keeps trusting the leaked root until an interactive install or uninstall.
- **Q183 - the GitHub Support request to purge the old commits that hold the leaked signing key.**
  Not a decision but an action: he chose (a), "send it for me", on 2026-11-17 (the next entry), and
  the request went to GitHub Support on 2026-11-18 as ticket #5020876, signed in as the repository
  owner. A second request follows the personal-data history rewrite.

#### The test VMs - Q195

- **Q195 - how much of his PC's processing power the test VMs may take.** He had asked for as much
  parallel work as possible, extra VMs included (2026-11-17). Memory is fine - 35 to 64 GB stays
  free - but the processors are not: with everything running at once his own work slows down, and
  each test run takes twice as long, so going wider gains less than it looks. *"195 rec"*.
  **Decided: the recommendation, (e)**: a load check - a new run starts only while the PC's
  processors are below about 70 % busy, and otherwise waits - plus a pause switch he can flip when
  he needs the PC: running test VMs finish their run, and no new one starts until he flips it back.
  The rule this brings to `AGENTS.md` (Test VMs stay saved unless in use) comes with the stream that
  builds it.

### 2026-11-17 - the maintainer's answers to Q183 to Q193: personal data in the public history

**What this is.** Questions asked and answered in conversation on 2026-11-17, after a read-only
audit of everything this public repository has ever published: the 923 commits of master, the 198
commits GitHub still serves by ID from the history rewrites of 2026-03-24 and 2026-08-24, every
blob, message and author field, and the GitHub side - releases, Actions runs, the activity feed.
The audit, with the values, stays in `.work\history-audit-2026-11\` of the maintainer's checkout;
this record holds none of them, and nor may any other committed file (Q190 (c)). Words in quotes
are the maintainer's; "rec" took the recommendation he had been given, spelled out below. In chat
these are numbered 183 to 193 (Q192).

**Primer.** The audit found two credentials, both only in the 198 left-over commits: the
code-signing key of releases v1.0.0.1 to v2.3.1.111, which opens with no password, and an old
test-VM password, changed since. In master it found personal data, most of it written by agents
into documents, tests, TODO items and commit messages: a third party's shared mailbox, its mail
domain and its use for the tests; a delegate's full name beside an incident in that person's
mailbox; a colleague's first name; one possibly real mail subject; and the maintainer's own e-mail
in file contents, his Windows user name, his computer name and facts about his mailboxes.

- **Q183 - the left-over commits, now.** *"183 a - send it for me."* **Decided: (a)**: GitHub Support
  is asked to purge the 198 commits and their cached pages. GitHub takes such a request only through
  its support form, from the repository owner's account; it is being sent from his account on
  2026-11-17. Not chosen: making the repository private first (installed copies could not update,
  and the CodeQL licence of Q125 needs a public repository), deleting and re-creating it (that loses
  issue #1, the release dates and the download counts), and waiting for the rewrite.
- **Q184 - the old signing key.** *"184 a"*. **Decided: (a)**, treat it as compromised: it was
  deleted from his certificate store and put on Windows' untrusted list for his user, both on
  2026-11-17. Q193 replaced the half that concerned the installer.
- **Q185 - what comes out.** *"185 b"*. **Decided: (b)**: the third parties' data - the
  organisation's shared mailbox, its domain, its local part and display name, the planned test
  mailbox, the delegate's name and the incident line, the colleague's first name and the possibly
  real subject - and his own machine and mailbox details: his e-mail in file contents, his Windows
  user name, his computer name, the facts about his mailboxes, and his name outside LICENSE. It comes
  out of every commit since it first appeared; the audit counted 764 of the 923 commits and 3
  release tags. The upstream author's MIT attribution and his own already-public data stay.
- **Q186 - his name and e-mail as author.** *"186 a"*. **Decided: (a)**: both stay as author and
  committer, and his name stays in LICENSE. That drops Q190 (d), GitHub's setting that blocks a push
  exposing his e-mail: it would block his own pushes.
- **Q187 - the tool.** *"187 a"*. **Decided: (a)**: `git filter-repo`, under a written one-off
  exception, now in `AGENTS.md` (Dependencies). Not chosen: `git filter-branch`, which git's own
  documentation warns against; git's export and import with a transformer of our own; BFG.
- **Q188 - tags and releases.** *"188 a - we will provide a new install and update after this all
  anyway. the updater of existing installations needs to keep working however."* **Decided: (a)**:
  the rewritten tags replace the old ones under the same names; the releases and their installers
  stay. Checked the same day: every release with an updater, v1.1.0.61 onwards, asks GitHub for the
  latest release by owner and repository name, reads its tag name and downloads the asset named
  after it, and none reads a commit ID. GitHub's API orders releases by the date of the tagged
  commit, which the rewrite keeps.
- **Q189 - where tests get the real values.** *"189 rec"*. **Decided: (a) + (b), with (d) as a
  cross-check**: tests and self-tests use placeholders; live and testbed runs read the Exchange VM's
  mailbox from the gitignored live-test settings file and check it against the guest's Outlook at
  run time; `Testbed/testbed.json` holds a placeholder; `AGENTS.md` says "the Exchange VM's
  mailbox". Until that is done, the baseline in `Tools/Checks/personal-data.json` counts what HEAD
  still holds.
- **Q190 - keeping it out.** *"190 rec"*. **Decided: (a) + (b) + (c) + (e)**, built the same day:
  (a) `Tools/Checks/check-personal-data.ps1`, the fifth guard, which the release runs and whose
  self-test runs on the build VM. It matched the identifiers by their salted SHA-256, committed in
  `Tools/Checks/personal-data.json` - until Q196 (c) kept them in the private list only - plus
  generic shapes: e-mail addresses outside
  the reserved domains, profile paths with a real-looking user name, private keys, key files,
  Windows default computer names, and the names of the machine it runs on. (b) Git hooks -
  pre-commit, commit-msg and pre-push - installed by `Tools/Install-GitHooks.ps1` into the main
  repository's `.git/hooks`, which every worktree shares; they also read the private list of real
  values, `.work/personal-data/identifiers.json` in the maintainer's checkout. (c) The rule in
  `AGENTS.md` (Personal data): placeholders only, in every tracked file, commit message, TODO item
  and release note. (e) `*.p12`, `*.snk`, `*.pem` and `*.key` in `.gitignore`, beside `*.pfx`.
- **Q191 - when.** *"191 rec"*. **Decided: (d), then (a)**: the purge of Q183 now; the content
  rewrite at a quiet point - the in-flight branches landed, work frozen, the rewrite, its proof,
  the force-push, every clone made again - followed by one commit that repoints the commit IDs
  cited in tracked files, from filter-repo's commit map.
- **Q192 - question numbers.** The maintainer asked for numbers without the "Q" everywhere; asked how
  far that should reach, *"a"*. **Decided: (a)**: in chat a question is "183"; the repository keeps
  "Q183", because a bare number cannot be searched - it also matches every count, version and date.
- **Q193 - the installer half of Q184.** Putting the old certificate on users' untrusted lists needs
  the installer to carry a copy, and the certificate's subject names his computer and his user -
  details Q185 removes. *"b"*. **Decided: (b)**: the next installer removes the old certificate from
  every trust store by its thumbprint, `E205060633DD7062D4F90033130542948A69D068`, and ships no copy;
  the untrusted-list entry stays on his own PC, where it is done.
- **Q196 - the committed hashes, 2026-11-18.** Q190 (a) committed the salted SHA-256 of each
  identifier so that the build VM could check without the values. A short value's hash is found
  again by hashing a dictionary: the computer name came back from its thousand candidates at once.
  **Decided: (c)**: commit no fingerprints. The private list is the only place the identifiers
  are; the build-VM runner copies it into each run, onto a VM that has no internet and is reverted
  after every run, and never fetches it back; a run without it says so in its summary.
  `Tools/Checks/personal-data.json` keeps the baseline alone, and the guard refuses it if it carries
  an identifier or a hash. The hooks and a release read the list where it lies.

### 2026-11-17 - Q131-Q136 carried out: folders and stores addressed by ID

The answers, with their primers, are in the next entry ("Folders and stores by id"). What was built
on them (`McpServer/README.md` load-bearing fact 18):

- **Q131 (b), the ID format:** the folder's whole entry id, re-encoded shorter with a check character -
  `Core/Mapi/OutlookIds.cs`: a kind letter, Crockford base32 without the constant bytes, and a GF(32) check
  that refuses every single wrong character and every swap of two characters fewer than 31 apart. 34
  characters for a folder in a data file, 66 for an Exchange folder. Nothing stored, no table.
- **Q132 (i), one folder while Outlook is closed:** found through the index by `System.ProviderItemID` for a
  data file's folder. Exchange measured by stream A on `OutlookAI-Exchange` (2026-11-17): the route does not
  work there - the row's id is the .ost's own node id, not in the 46-byte folder id, and it changes on an
  .ost rebuild - so such a search is refused with that reason. No learned maps.
- **Q133 (i), stores:** a 28-character `store_id` (`S` and the provider UID every entry id in the store
  carries), or `T` and the store's own entry id where the root was not read; display names beside it.
- **Q134 (i):** `create_folder(parent_id, name)`, a `/` part of the name; `move_mail` creates nothing.
- **Q135 (i):** a clean break: names and paths are refused with a pointer to `find_folder`, and the retired
  arguments by a call-tool filter, since the SDK would otherwise ignore them.
- **Q136 (ii):** `find_folder(name)`, read-only, candidates only.

Decided on the maintainer's behalf while building it: the target of `move_mail` is checked once, before
any item, so a refused folder refuses the whole call; a store whose name several index stores carry is
refused when nothing ties it by hash (stream A's stale-.ost finding); a folder whose path a namesake
shares is reported (`scope.pathShared`), not refused; a data file's folder ID is looked up in the index
first even while Outlook runs, so a search goes where the index files the folder now; an Exchange
folder ID is resolved through Outlook, which is started for it as for the search's sweep (D17), and
refused only when Outlook does not answer; `find_folder` carries the read-only tool annotations
`audit_log` carries, and `create_folder`, like every other write tool, none. Open follow-ups are in
`TODO.md`.

### 2026-11-17, autonomous - Q147 carried out: the signature model, its edits and its refusals

**Primer.** The maintainer decided the shape (Q147 (a)+(b), Q152 (a), Q153, Q155 (c), Q156, Q157 (a),
Q159, Q160; the add-in's language pick, Q158 + Q162 (b), later). Carrying it out took the choices below
on his behalf. The read model, the edits and the draft decision live in `McpServer/OutlookAI.Core/Services/`
(`SignatureEnvironment`, `SignatureManager`, `SignatureSelection`, `SignatureRtf`, `SignatureHtml`,
`SignatureText`, `SignatureListing`); `McpServer/README.md` describes them.

**1. Reading one signature in full.** *Options:* (a) a `name` argument on `list_signatures`; (b) a new
`get_signature` tool; (c) a `read` action on `manage_signature`. **Decided: (a)** - the tool count stays
22, and the read stays on the tool classified read-only; (c) would put a read behind a write tool's
gates. **Undo:** a tool and its pins.

**2. Editing the RTF rendition without shipping its pictures through the agent.** Outlook's RTF embeds
every picture as hex, often twice (a PNG and a WMF fallback). *Options:* (a) show each picture as a
`[[picture:N]]` token and put the original back byte for byte wherever the token stays, with
`[[picture-file:<file>]]` for a new PNG or JPEG; (b) return and accept the full hex; (c) RTF edits by
find-and-replace only. **Decided: (a)**, with (c) as well: a 50 KB logo is 100 KB of hex for nothing,
and (c) alone cannot add or remove a picture. **Undo:** `SignatureRtf.Elide`/`Expand`.

**3. A text edit that formatting splits.** *Options:* (a) refuse the whole call and say to send that
rendition edited; (b) replace across the formatting, dropping it; (c) skip that rendition silently.
**Decided: (a)** - (b) changes formatting nobody asked to change, (c) leaves the renditions saying
different things. **Undo:** `ReplaceText`'s split branch.

**4. What the policy refusals cover (Q156).** *Options:* (a) under `DisableSignatures`, no
`manage_signature` write at all, and a policy-forced signature cannot be deleted; (b) refuse only
default changes, as Q156 names. **Decided: (a)** - Outlook's own signature editor is closed under that
policy, and deleting a forced signature leaves Outlook's forced default naming nothing. **Undo:** two
refusals in `SignatureManager.Manage`.

**5. A default deployed in Outlook's own settings (`Common\MailSettings` `NewSignature`/`ReplySignature`,
not a policy).** Its effect is disputed in the sources (research section 2). *Options:* (a) every
account's default, above the account's own, default changes refused, the per-message choice free;
(b) below the account's own; (c) ignored. **Decided: (a)**, the research's precedence model, until G2's
live test (`LiveSignatureSettingsTests.G2_...`) records what Outlook's own signature editor reports. Its
first run (2026-11-17, OAI-UNINDEXED) could not tell: Outlook's signature settings reported no default
with the deployed one or without it, written while Outlook ran; `TODO.md` holds the measurement.
**Undo:** `SignaturePolicy.Read`'s preference source.

**6. Roaming on perpetual Office (Q152 (a)), after stream A's R1 (Office LTSC 2024 showed no sign of
roaming) and the coordinator's instruction to keep the detection generic.** *Options:* (a) off unless
Outlook's own `Name (address)` copies of an Exchange account are on the computer, then evident for
those accounts; (b) treated as possible, as first built; (c) off. **Decided: (a)** - right whichever
way R3 comes out, and it leaves the testbed's LTSC 2024 and the maintainer's LTSC 2021 their local
defaults. Microsoft 365 Apps 2303+ stays possible on Microsoft's word. **Undo:** `SignatureRoaming.Decide`.

**7. A named signature on a reply from a delegate or shared mailbox (Q160).** *Options:* (a) a
signature of this computer goes in, no default does; (b) refuse any signature there. **Decided: (a)** -
Q160's "no signature for now" read as the default; the agent's explicit choice is honoured and the
result says which went in. **Undo:** the delegate branch of `SignatureSelector`.

**8. A chosen signature that fails to go in.** Since the crash fix nothing else puts one in, and a
second inspector on the same item is the shape that wedged Outlook (`ReviseHeldDocument`). *Options:*
(a) report it - the draft has no signature, and `update_draft` with the same signature retries; (b)
retry with the account default in a second inspector. **Decided: (a).** **Undo:** none needed.

**9. Smaller ones.** `"none"` is reserved, so `manage_signature` refuses a signature of that name.
`clear_default` takes the name of the default it clears, so nothing is cleared by mistake. An Exchange
account added under an alias is matched by `Account.DisplayName` against its entry's `Account Name`. The
server instructions' budget is 480 characters, up from 400, for the one signature sentence Q157 asked
for. The add-in's signature button reads the folder through the one linked rule
(`Services/SignatureFolderLocation.cs`) - the add-in otherwise unchanged, as Q158 asks - because the
folder defect was named in all three places. **Undo:** each is one change.

**10. What the add-in's own pick will call (Q158 + Q162, built later).** *Options:* (a) the code the
server decides with, as it stands in Core; (b) a separate facade for the add-in. **Decided: (a)** -
one copy (Q154), and the server exercises it on every draft: `SignatureEnvironment.ReadCurrent()` for
the folder, profiles, policy and roaming; `SignatureSelector.Candidates(environment, account,
runningProfile)` for the signatures that account may use, with excerpts to judge the language by, its
effective defaults, and whether policy disables or forces them; `SignatureSelector.Decide(...)` for
the default and the reason. The add-in keeps one thing for itself: whether the signature in the open
draft was put there by a person. **Undo:** nothing is built on it yet.

### 2026-10-11 and 2026-11-17 - the maintainer's answers to Q118, Q131 to Q136, Q138 to Q140, Q142 to Q147 and Q149 to Q162

**What this is.** These questions were asked and answered in conversation, not in this file, and
none was recorded in the repository until now: answers of 2026-10-11, in two messages, and of
2026-11-17. Words in quotes are the maintainer's. Where he answered "your recommendation" or
"whatever you think is best", the recommendation he had been given was taken, and it is spelled out.
The research behind several answers is in `.work\research-2026-10-11\` (q118-tray, q145-admin,
q147-signatures, q149-pins, q149-wordeditor) and the duplicated-code survey in
`.work\dup-survey-2026-11\`, both in the maintainer's checkout. `.work` is gitignored, so neither is
in a clone; what the decisions rest on is restated here. The survey's questions, Q163 to Q172, were
answered on 2026-11-18 (the 2026-11-18 entry).

#### Folders and stores by id - Q131 to Q136, 2026-10-11

**Primer.** On 2026-10-03 the maintainer decided that tools address a folder only by a unique id,
its name kept for display (Q114), and asked whether a search of one folder can be answered from the
search index alone while Outlook is closed (Q115). `Docs/live-tier-on-the-vm.md` section 8 item 31
measured it: a local mail file's folder id is its 24-byte EntryID, which survives a rename, a move
within its store, a soft delete and restarts, and the index keeps every row's node id, so the id
finds the folder's index scope with no Outlook call, in 59 to 114 ms. Exchange's 46-byte ids were
not measured. Q131 to Q136 asked what the id looks like and how the tools change around it.

- **Q131 - the id's form.** *"what you think is the most stable"*. **Decided: (b)**, the folder's
  full EntryID, re-encoded shorter, with a check character, and nothing stored. It is stateless: no
  table to go stale, the same id in every session and process, and usable while Outlook is closed.
  About 34 characters for a local mail file's folder and 66 for an Exchange one.
- **Q132 - searching one folder while Outlook is closed.** *"what you think is the most stable"*.
  **Decided: (i)**: the Windows Search index resolves the folder id - measured for local mail files,
  correct after renames and moves, about 0.1 s. Exchange takes the same route once it is measured on
  the Exchange VM; until then OutlookAI refuses such a search with a clear message. No learned map
  from ids to paths.
- **Q133 - addressing a store.** *"i + display names for reading along as a human"*. **Decided:** a
  store is addressed by a `store_id`, with display names shown beside it for a person reading along.
- **Q134 - making a folder.** *"i"*. **Decided:** a separate `create_folder(parent_id, name)`, in
  which a `/` is part of the name; `move_mail` no longer creates folders.
- **Q135 - folder paths in existing calls.** *"i"*. **Decided:** a clean break in one release: path
  input is refused, with a pointer to `find_folder`.
- **Q136 - finding a folder by name.** *"ii"*. **Decided:** a read-only `find_folder(name)` that
  returns candidates and never acts.

Together they make `TODO.md`'s item on a folder whose name holds `/` moot: such a folder is reached
by its id, and created with the `/` in its name.

#### The Exchange VM - Q138 to Q140, Q161

`OutlookAI-Exchange` holds one real Microsoft 365 mailbox (Q108) and stays read-only for tests until
the maintainer approves the Phase 2 write-safety design (`Docs/live-tier-on-the-vm.md` section 4.4).

- **Q138 - the old tagged item in that mailbox's Sent Items, 2026-10-11.** The VM's read-only count
  found one item carrying `[OutlookAI-McpTest]` there, left from the years the live tests ran on the
  workstation; it fails `T2/LiveExchangeHubArtifactTests` on every run until it is gone. *"c"*.
  **Decided:** a one-off, TESTED helper deletes it, only after the maintainer has confirmed its
  EntryID, selecting by EntryID allowlist AND ordinal tag match. It is the only write that mailbox
  ever gets.
- **Q139 - Outlook's "a program is trying to access" prompt, 2026-10-11.** Outlook's Object Model
  Guard stalls a run with it whenever Defender's signatures are stale, and
  `Invoke-ExchangeSignIn.ps1 -Mode Preflight` updates them before every run. The offline guests also
  carry Q80's auto-approve policy, which on this online VM would let any process read addresses
  without a prompt. *"a"*. **Decided:** keep only the signature-update preflight; no auto-approve
  policy.
- **Q140 - the add-in on the Exchange VM, 2026-10-11.** Not installed, because Phase 1 needed no
  write to the profile's configuration; `T2/LiveHealthTests` and the `AddInRegistry` tests need it.
  *"b"*. **Decided:** install the OutlookAI add-in there. Under Q145 (a) below, the settings it
  manages are written at user level, never as policies.
- **Q161 - signing in to Outlook on the web there, 2026-11-17.** Measuring roaming signatures (Q147
  (b), below) takes a test signature kept in the mailbox itself, which Outlook on the web makes; a
  browser sign-in on the VM is a route Q111 does not cover. *"c"*. **Decided:** an agent may sign in
  to Outlook on the web on the Exchange VM, to measure roaming signatures. The agent that runs the
  Exchange VM writes the amendment to Q111 in `AGENTS.md`, with its conditions.

#### Cached Mode settings without administrator rights - Q142, Q143, Q145, Q151

**Primer.** OutlookAI keeps five of Outlook's Cached Mode settings so that the mail an agent
searches is offline and indexed: how much mail is kept offline (a months value and a days value),
whether shared and delegate folders are cached (two values), and whether shared mailboxes follow
the offline period. It writes them as Outlook POLICIES, under `HKCU\Software\Policies`, which
Windows normally lets a user read but not write; since Q128 (2026-10-03) OutlookAI Settings shows
them and writes the ones the user cannot through an elevated helper, `OutlookAI.PolicyWriter.exe`,
behind one UAC prompt. Q145 asked whether that helper may run from the per-user install folder,
where any process of the user can replace it before the prompt; Q142, whether a standard user
approving the prompt with an administrator's password works - never tested.

- **Q142 and Q145 - why administrator rights at all; asked 2026-10-11, decided 2026-11-17.** On
  2026-10-11: *"Why do we even need admin for all these user level settings?"* The research's
  answer: only because the values were written as policies. Each has a user-level home Outlook
  honours: the offline period and the "Download shared folders" box are kept per account in the
  user's own Outlook profile, where Account Settings changes them; the shared-mailbox switch has a
  documented value under the user's own Cached Mode key; and the two shared-folder switches are on
  by default. On 2026-11-17: *"Q145 a"*. **Decided: (a), no administrator rights:**
  - the elevated helper, its UAC path and its part of the installer are deleted;
  - the settings are written at user level;
  - Microsoft 365 Cloud Policy and Group Policy keys are only read, to show a value as "Set by your
    organization";
  - an existing account's offline period: the user is guided to Account Settings now, and OutlookAI
    writes the account's own profile value (`0x6649`) once the Exchange VM shows that Outlook
    honours it. That writer goes into OutlookAI.Core, under the Q154 rule;
  - Q142 dissolves with the prompt.
- **Q143 - which offline periods, 2026-10-11.** OutlookAI Settings took months 0, 1, 3, 6, 12, 24
  and days 0, 3, 7, 14 as two independent lists. *"unsure, do what you think is best to be complete
  and correct"*. **Decided:** exactly the eleven periods Microsoft documents, each written as one
  (months, days) pair: 3 days (0,3), 1 week (0,7), 2 weeks (0,14), 1, 3 and 6 months (1,0), (3,0),
  (6,0), 1, 2, 3 and 5 years (12,0), (24,0), (36,0), (60,0), and All (0,0). Nothing else is
  accepted, because Microsoft says an invalid value silently becomes 12 months (KB 3115009). 3 and 5
  years, which the old lists lacked, are confirmed on the Exchange VM.
- **Q151 - the policy values earlier versions wrote, 2026-11-17.** OutlookAI 3.0.0.319 to 3.1.0.325
  wrote the five as policies wherever the user's token could - the maintainer's workstation since
  2026-07-23 - and a policy locks Outlook's own setting. *"Q151 a - So I can see how it looks in our
  gui. Then after I will tell you to remove them."* **Decided: (a)**: they stay in place, labelled
  "Set by an earlier OutlookAI", and are removed on his word. Before they are, each account's own
  period is set to All: one account on his workstation reads 12 months under the policy's All, and
  removing the policy alone would make Outlook drop its offline mail older than that (inferred from
  Microsoft's KB 2733062). `TODO.md` holds the removal.

#### Pins, and how drafts are written - Q118, Q149, Q150

**Primer.** When a program asks for Outlook while it is closed, Outlook starts with no window - only
its icon by the clock, *"Another program is using Outlook..."* - and OutlookAI starts it that way on
purpose (the v3 plan's D33). In that mode Outlook ends when its last window closes, a draft's hidden
compose window included, so a session holds a never-shown Explorer, the pin, to keep it up (D49).
Q118 measured pins left behind (`McpServer/Docs/com-host.md`, "Pins left behind, measured"): a
session that did not start Outlook never closes its pin, and a pin left behind keeps a window-less
Outlook up that `Application.Quit()` does not end.

- **Q118 - pins left behind, 2026-10-11.** *"Q118 a - what about the special mode with the special
  icon in the systray?"* **Approved: (a)**, a pin lifecycle fix: a shared list of sessions; the last
  session out closes the pin; pins of dead sessions are closed; `outlook_health` lists the sessions;
  and a lock, so sessions that start Outlook at once make one pin. His question was researched the
  same day: the tray mode is Outlook's own "Another program is using Outlook" state, OutlookAI has no
  tray icon of its own, and pins exist only in that mode. **(a) is on hold** until Q149's
  measurement decides between pins only while drafting and no pins at all; `TODO.md` holds it.
- **Q149 - do drafts need a window, and pins; asked 2026-10-11, decided 2026-11-17.** On 2026-10-11:
  *"Q149 What happens if we drop the pin system? Why is it needed?"*, then *"I have serious doubts
  that the word editor needs a real window. Research this in more depth. I also have serious doubts
  that we need pins. Otherwise the special hidden outlook mode microsoft programmed would be
  hobbled. Research if this really is the case. Then ask the question again in a similar fashion
  updated with the research results."* What the research found:
  - the window-less mode is not hobbled: everything works there but the Word editor and what needs a
    compose window, such as Outlook's own signature insertion;
  - closing a compose window ends a window-less Outlook - Microsoft's rule that closing the last
    window ends Outlook - measured 7 of 8 on Office 2021, never on Office 2024;
  - the Word editor is refused only while the Outlook process has no window shown or activated
    (Office 2021); the draft's own window never has to be shown;
  - edits to a re-opened draft are kept only after activation (3 of 3 on Office 2021; on Office 2024
    400 of 400 bodies were lost without it).

  On 2026-11-17: *"Q149 d"*. **Decided: (d)**: measure first, on the test VMs, then choose between
  (b), keep Word, with pins only while drafting plus Q118's lifecycle work, and (c), write the
  draft's text as HTML through Outlook's object model, with no compose window and no pins. (c) is
  chosen only if its drafts are indistinguishable from today's.
- **Q150 - after the user closes Outlook, 2026-11-17.** While an agent is still working, OutlookAI's
  next call that needs Outlook starts it again without a window (the v3 plan's D17), so the icon
  comes back within seconds. *"a"*. **Decided:** OutlookAI restarts it on the next call, as today.

#### Signatures - Q147, Q152 to Q160, Q162

**Primer.** Outlook keeps signatures as files - .htm, .rtf and .txt, plus a `<name>_files` folder
for pictures - in one folder per Windows user, and each mail account's defaults for new mail and for
replies and forwards in its Outlook profile. The MCP server lists them (`list_signatures`), creates,
updates and deletes them (`manage_signature`), the draft tools take a `signature`, and since the
compose-crash fix of 2026-10-04 OutlookAI inserts an account's default itself. All of it read the
default profile and one fixed folder; nothing read the profile Outlook runs on, administrator
policy, or roaming signatures, the ones Microsoft 365 keeps in the mailbox. Classic Outlook uses
those from version 2303, and only Microsoft 365 Apps is known to: never the maintainer's Office LTSC
2021, and Office LTSC 2024 is unmeasured. The add-in's task pane has one action, "Select the best
signature", over the add-in's own copy of the server's signature catalog. Research: q147-signatures.

- **Q147 - signatures, 2026-10-11.** *"Q147 a+b - OutlookAI should be able to handle
  creating/editing/deleting signatures and selecting signatures per message based on all signatures
  in the relevant account. This signature selection and management should be possible from both the
  in app gui as well as the mcp server."* **Decided: (a) and (b)**: (a) read the running Outlook
  profile and the Group Policy setting; (b) measure roaming signatures on the Exchange VM and handle
  them. Q158 narrowed the GUI half on 2026-11-17. Its research raised Q152 to Q161.
- **Q152 - how far roaming signatures are handled, 2026-11-17.** No supported interface reaches a
  mailbox's signatures, and Microsoft has no plans for one in Microsoft Graph. *"your
  recommendation"*. **Decided:** use Outlook's local copies, detect roaming, and refuse a local edit
  of a signature the mailbox manages. Whether OutlookAI also reads the mailbox's own signature
  folder is decided after a read-only measurement. It is never written.
- **Q153 - testing classic Outlook's roaming, 2026-11-17.** Only Microsoft 365 Apps is known to
  roam, and the testbed runs perpetual-licence Office by decision. *"your recommendation"*.
  **Decided:** no Microsoft 365 Apps in the testbed. Measure the mailbox side and what Office 2024
  does; classic Outlook's roaming is a known limit.
- **Q154 - one copy of the code, 2026-11-17.** The add-in and the server need the same signature
  logic, and the add-in kept its own copies of it, already drifting apart, because the "Phase-6
  no-coupling choice" kept it from referencing OutlookAI.Core. That choice is recorded only in the
  gitignored v3.MD (lines 64, 79-80, 144, 184 and 264) and cited in a comment of the add-in's
  signature catalog; Core builds for net48 too, so that the add-in could reference it. *"Q154 b -
  refactor everything to bring this pattern everywhere relevant. I want as much code shared as
  possible. And as a rule no code duplicated."* **Decided: (b)**: the add-in references Core's
  net48 build behind a small interface, superseding the no-coupling choice - and, as a standing rule
  now in `AGENTS.md` (Shared code), no code is duplicated. A survey of master 9f4646d found 73 groups
  of duplicated code (`.work\dup-survey-2026-11\`: README.txt, inventory.tsv, members.tsv); its
  follow-up questions, Q163 to Q172, were answered on 2026-11-18 (the 2026-11-18 entry).
  `TODO.md` holds the refactor.
- **Q155 - which Outlook profile, 2026-11-17.** Outlook uses the profile it runs on, which is not
  the default one after `/profile` or the profile picker; OutlookAI read only the default profile,
  and with none recorded it merged every profile into one list. *"Q155 c - Always be explicit. A
  draft that needs a signature already is written in one of the accounts and send by that email
  which is a choice the agent made. It can match the correct signature. When adding the signature to
  a draft maybe suggest what signature matches with the used account."* **Decided: (c)**: every
  Outlook profile is listed, and every answer names its profile; changing a default needs an
  explicit profile and account; a draft's result says which signature went in, and suggests the one
  that matches the sending account.
- **Q156 - administrator policy, 2026-11-17.** Policy can forbid signatures or force one default on
  every account, and Cloud Policy overrides Group Policy. *"your recommendation"*. **Decided:**
  respect it: no signature where policy forbids them; a forced default is used, and a change of
  default is refused. A per-message choice under a forced default is allowed only if a measurement
  confirms that Outlook allows one.
- **Q157 - the draft tools' `signature`, 2026-11-17.** There was no way to ask for none,
  `update_draft` could not remove one, a result did not say which default went in, and a name was
  checked against every signature on the machine. *"Q157 a - make sure the agent understands from
  the mcp operation instructions"*. **Decided: (a)**: `signature` left out means the account's
  effective default; `"none"` means no signature, and in `update_draft` also removes the one there;
  only a name the sending account can use is accepted; the result says which signature went in and
  why. The tool descriptions and the server instructions spell this out.
- **Q158 - the add-in's part, 2026-11-17.** In full: *"I mean, in Outlook there is already an
  interface for signature management. So we do not have to implement that interface. Only for the
  mcp server. But inside of the outlook addon gui I do want the ai to auto select the correct
  signature. For example based on the language of the email so the difference between my kind
  regards en vriendelijke groeten signatures get set correctly. No need for an interface or
  dropdown. That is already part of outlook itself. Just a checkbox "Allow the AI to automatically
  choose the correct signature" (on by default) or something."* **Decided:** signatures are managed
  through the MCP server only; the add-in gets no editor, list or dropdown. In the add-in the AI
  chooses the signature itself - by the mail's language, for one - under one checkbox, "Allow the AI
  to automatically choose the correct signature", on by default. When it chooses is Q162.
- **Q159 - pictures and formats, 2026-11-17.** `manage_signature`'s update deleted a signature's
  `<name>_files` folder and rebuilt its .rtf from plain text, so editing a signature with a logo
  broke it. *"Have full support for existing signatures with pictures and new ones created by the
  user or the mcp server in all formats. The goal is that the mcp server can also tweak existing
  signatures without breaking or converting stuff."* **Decided:** full support, in every format, for
  existing signatures with pictures and for new ones the user or the MCP server makes; the MCP
  server edits an existing one without breaking or converting it.
- **Q160 - replies from a delegate or shared mailbox, 2026-11-17.** For mail in a store no account
  delivers to, OutlookAI inserts no default, and what Outlook itself does there is unmeasured.
  *"your recommendation"*. **Decided:** no signature for now, and the result says so; once the
  shared test mailbox exists (Q109), measure what Outlook does there and match it. `TODO.md` holds
  it, with Q109.
- **Q162 - when the add-in's AI chooses, 2026-11-17.** *"b"*. **Decided:** when it writes or
  rewrites the body, and when a reply or forward opens, matching the language of the mail being
  answered. It never replaces a signature someone chose by hand, and never changes a mail as it is
  sent. It waits for the add-in's reference to Core (Q154); `TODO.md` holds it.

#### Tests and scans - Q144, Q146, 2026-10-11

- **Q144 - keeping a time zone that is not UTC.** Since Q126 (a) the non-live suite runs only in
  UTC, so a test that uses the machine's own zone as the "other" zone proves only the identity there;
  `TODO.md`'s item names the tests found and five directions. *"whatever you think is best"*.
  **Decided: (b) and (c)**, that item's directions (2) and (3): the build VM stays on UTC; a short
  second pass in W. Europe runs the tests marked as zone-sensitive, and a check fails a
  zone-sensitive test that is not marked; and tests are rewritten to set their own zone as they are
  next touched.
- **Q146 - the threat model of the release's CodeQL scan.** The release runs CodeQL's default threat
  model, under which only network input is untrusted, and the product reads none.
  `local,!environment` gave 6 results over 0ccf5df, one of them real - the elevated helper's path
  read from HKCU, which goes with Q145 (a). The MCP server's real input, an agent's tool arguments,
  is read by the MCP SDK outside the analysed source, so CodeQL never sees it (inferred). *"whatever
  you think is best"*. **Decided: (b), then (d)**, the `TODO.md` item's directions (2) then (4): the
  release's scan runs with the threat model `local,!environment`, every result fixed or recorded with
  its reason; then the MCP tools' parameters are modelled as untrusted input.

#### Working in parallel - 2026-11-17

*"try to parallalize as much as possible. Even create extra virtual machines if you think that can
speed things up."* Since then several agents work at once, each in its own worktree, and one adds
build VMs and clones of the Outlook guests. `AGENTS.md`'s rules for test VMs cover the new ones too.

### 2026-10-04, autonomous - the compose crash: dumps on the guests, the cause, and a compose that never hands Outlook that close

**Primer.** OUTLOOK.EXE crashed now and then in the live compose tests on both Outlook guests (runbook
4.1e, 4.2f, 4.5), and the guests kept no dump. They do now, and the 30 dumps of this work all read the
same: an access violation at `OLMAPI32.DLL+0x2E411`, on Outlook's main thread, inside the client's
`Inspector.Close` on the hidden compose inspector of an item `Items.Add` had just returned (runbook 4.6).
Five decisions were taken on the maintainer's behalf.

**1. How the guests keep dumps.** *Options:* (a) the runner arms Windows Error Reporting's LocalDumps key
after every restore, with its own checkout's script; (b) new frozen checkpoints with the key in them;
(c) a scheduled task in the guest; (d) dumps only when asked for. **Decided: (a)**
(`Testbed/guest/Set-OutlookCrashDumps.ps1`, `Invoke-LiveTierOnGuest.ps1`). No checkpoint changes, so the
frozen instants of 4.5 and their T1 pins stand; any commit is dumped, an older one too. (b) means two new
frozen instants, new pins and minutes off the 90-minute window; (c) never runs, because a restore resumes
a running guest with no logon; (d) loses the next crash the way the last ones were lost. **Undo:** revert.

**2. How the dumps are read.** *Options:* (a) a minidump reader written here, unwinding with Windows' own
`dbghelp.dll` (`StackWalk64`) and naming frames by export, no symbols; (b) the Debugging Tools for Windows,
which need a Dependencies exception; (c) Microsoft's public symbols. **Decided: (a)**
(`Testbed/host/Read-CrashDump.ps1`). Nothing outside Windows, and enough here: the faulting function's own
bytes, the broken holder it was resetting, and the `Invoke` frame's arguments named the call. (c) was
checked and has nothing to give: Office publishes no PDB for this build (404 for `outlook.pdb` and
`olmapi32.pdb`), only Windows' own modules answer. No exception was asked for. **Undo:** none needed.

**3. The fix.** A scratch soak - `new_draft` 400 times in a row - crashed about once in 40 (4 crashes in
171 new drafts), and twelve variations of the compose changed nothing that mattered (runbook 4.6, item 5).
*Options:* (a) never open an inspector on the object a creator returned: save it, re-open it by EntryID and
compose there the way update_draft revises a draft; (b) a pause before the close (2 s: one crash in 237);
(c) keep Outlook's own signature injection and end that inspector with `MailItem.Close` (one in 309); (d)
compose by HTML alone, without Word, which drops the placement contracts of D46 and A1; (e) leave it, as
Outlook's fault. **Decided: (a)** (`ComposeReopened`, `T1/ComposeOnReopenedDraftTests`): no crash in 2,400
new drafts, nor in 20 compose-class and 4 full-tier runs. (b) and (c) only make it rarer; (d) gives up what the Word path exists for; (e) leaves
users' Outlook crashing about once in 40 drafts. *Its cost:* new drafts now take update_draft's off-screen editor
promotion (`Activate`) on a windowed Outlook too, as revisions always have; nothing reaches the screen.
**Undo:** revert `c4595b1` and `c31a9d9`.

**4. The account's default signature.** Outlook adds it only in the inspector the fix may not open, so
drafts for an account with one came out without it. *Options:* (a) OutlookAI inserts the configured
default itself - the profile's "New Signature" or "Reply-Forward Signature", in Outlook's default profile,
the view `list_signatures` and `manage_signature` use - only when the agent named none and Outlook injected
nothing; (b) no default signature in agent drafts; (c) item 3 (c) for accounts that have one; (d) read the
running session's profile instead of the default one. **Decided: (a)**, through the path a named signature
already takes. (b) loses what users get today, (c) is the crash, and (d) would make the draft disagree with
`list_signatures`. **Not covered, and not measured:** a signature Outlook picks by other means - roaming
signatures kept in an Exchange Online mailbox, a Group Policy MailSettings value - and an Outlook running on
a profile that is not the default. **Undo:** revert `c31a9d9`.

**5. The soak test stays scratch.** *Options:* (a) build it as scratch commits outside the branch, with
git plumbing, and describe it in the runbook; (b) commit it as an opt-in live class. **Decided: (a)** - a
400-draft loop is a measuring instrument, not a contract, and the compose classes plus the T1 pin carry the
contract. Runbook 4.6 describes it well enough to rebuild. **Undo:** none needed.

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
   triage depending on how much of "local" is on. *Decided 2026-10-11 as Q146 and carried out
   2026-11-17:* local input except environment variables, then every MCP tool's parameters as
   untrusted input through the model in `Tools/CodeQL` (`Tools/Invoke-CodeQL.ps1` says how).
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
