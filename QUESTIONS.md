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
**ANSWERED 2026-08-18: do this.**

**(b) The unindexed-store list is uncapped.** Every other list in this server has a cap and a has-more
flag. A profile with many unindexed PSTs would list them all, in the payload and in an advice
sentence. *Recommendation: cap it like the others* - the principle is already settled here, this is
just an omission. **ANSWERED 2026-08-18: do this.**

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

## Decision log

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
