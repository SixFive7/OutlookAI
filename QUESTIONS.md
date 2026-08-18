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

**Default if unanswered.** No release. I will not trigger the release workflow autonomously; the
project's own rules make that an explicit-word-only action, and I am treating that as unchanged by
the general autonomy grant.

---

## Q2 - The 2 KB description cap is trusted, not verified

Every description-budget decision - including cutting `search` from 3912 to 1798 characters - rests
on Claude Code's documentation saying it "truncates tool descriptions and server instructions at 2KB
each". That is the vendor documenting its own client. I have never observed it truncating our
strings, and the documentation does not say whether the 2 KB is characters or UTF-8 bytes, which is
why the guardrail measures both and fails on the larger.

**Options.** Trust the documentation and move on; verify empirically by shipping a deliberately
over-long description and inspecting what the model actually receives; or ask Anthropic.

**Recommendation.** Trust it. The failure mode of being wrong in the conservative direction is a
slightly terser description; the failure mode of being wrong the other way is silent loss of
instructions, which we have already seen produce a real defect.

**Default if unanswered.** Trust the documentation, keep the guardrail measuring both units.

---

## Q3 - Whether `search` should stay this close to the cap

`search` is at 1791 of 2048 - 87%, inside the guardrail's warn tier. It is the most-used tool and
its description is doing real work. Two others sit in the tier as well: `update_draft` at 1593 and
`outlook_health`, which was brought down to 1337 today.

**Options.** Leave it and accept that any future addition to `search` must first remove something;
trim it further now toward a comfortable margin; or raise the warn threshold so it stops flagging.

**Recommendation.** Leave it. The remaining text is what a caller needs before the call plus the one
instruction it must act on afterwards, and the warn tier doing its job is not a reason to silence
it. Raising the threshold to stop a true warning would be the wrong move.

**Default if unanswered.** Leave it, and let the warn tier keep flagging it on every CI run.

---

## Q4 - An unscoped search still sweeps every store from one window base

Freshness is now measured per store when a search names a store. An **unscoped** search - one that
covers every account - still opens a single window from the profile-wide newest indexed mail, so a
quiet account's gap is only covered when the search names that account. Measured spread between
stores on this profile: 45.4 hours.

**Options.** *(a)* Give the sweep a per-store window - one `sinceUtc` each - which needs a frontier
probe per store on every unscoped search: about five extra TOP-1 index queries on a path that
currently costs 85-185 ms. *(b)* Use the profile-wide MINIMUM frontier, so the window is as wide as
the slowest store for everyone - cheap, but it would routinely trip the 200-item per-folder cap on
busy stores and report `partial` constantly, trading a silent miss for a permanent false alarm.
*(c)* Leave it, documented: name the account when freshness in a quiet account matters.

**Recommendation.** (a) is the correct answer and the cost is probably acceptable - five index
queries against 85-185 ms is noise - but it is a hot path and I would want it measured rather than
assumed, which needs a decision about spending that latency at all. (b) I would rule out: it makes
every busy-store search lie in the other direction.

**Default if unanswered.** Leave it at (c) and keep it documented. Nothing regresses; the behaviour
is what shipped before tonight, now merely understood and written down.

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

**Default if unanswered.** The reversal stands, and the test carries the explanation.

## Q6 - Does the 2 KB cap apply to parameter descriptions, and is it per string or per tool?

The documented sentence covers two surfaces: *"Claude Code truncates tool descriptions and server
instructions at 2KB each."* Our guardrail applies 2048 to a third - every
`inputSchema.properties[*].description` - which is an extrapolation I made without flagging it.

Two separate unknowns sit here, and only one is cheap to be wrong about.

**Do parameter descriptions have a cap at all?** If not, we are merely terser than necessary. Fine.

**Is the cap per STRING or per serialized tool?** This one matters. Our guard assumes per string. If
the 2 KB actually applies to a tool's whole entry in `tools/list` - description plus schema - then
`search` at 1791 characters plus 831 on `folder`, 609 on `exhaustive` and 569 on
`include_attachment_hits` is far over, and the trim that moved detail from the description onto the
parameters moved it from one capped bucket into the same capped bucket. That would mean a fix I
reported as solving the problem solved nothing.

**Options.** Keep the conservative per-string guard and accept the unknown; verify empirically by
overshooting deliberately and observing what the model receives; ask Anthropic.

**Recommendation.** Keep the per-string guard - it costs nothing and errs safe. But settle the
per-tool question, because unlike Q2 its downside is not terser text, it is work that did not do what
it claimed. The same overshoot experiment answers both.

**Default if unanswered.** Guard stays as is; the uncertainty stays recorded here rather than
implied to be settled.

## Q7 - Three follow-ups the freshness work deliberately did not decide

Raised by the agent that added the `no_index_frontier` state; none blocks anything.

**(a) `staleness.newestIndexedUtc` on an unscoped search** is still the profile-wide maximum. It is no
longer the sweep's window base, and the advice age now uses the widest per-store frontier, but the
field itself still reports the maximum - because narrowing it would make `search` and `outlook_health`
report different numbers for the same profile. Options: leave it; add
`staleness.oldestStoreFrontierUtc`; or change the field's meaning and update health to match.
*Recommendation: add the second field.* It answers the question without making two tools disagree.

**(b) The unindexed-store list is uncapped.** Every other list in this server has a cap and a has-more
flag. A profile with many unindexed PSTs would list them all, in the payload and in an advice
sentence. *Recommendation: cap it like the others* - the principle is already settled here, this is
just an omission.

**(c) `notNeeded` now costs one ordinary sweep** in a narrow case: an unscoped search bounded to mail
older than the frontier but newer than the fallback runs a sweep where it previously did no COM work.
That is the price of `notNeeded` no longer lying on unindexed stores. *Recommendation: accept it.* Any
mitigation trades a bounded window of completeness for latency, which is the trade the standing rule
forbids.

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

**Default if unanswered.** Leave it. Recorded in `Docs/completeness-gaps.md` as B3, unclosed.

## Decision log

Answers move here with the date and the reasoning, so a future reader sees not just what was chosen
but why, and what the alternative was.

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
