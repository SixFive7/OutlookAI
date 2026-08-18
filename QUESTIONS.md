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

## Decision log

Answers move here with the date and the reasoning, so a future reader sees not just what was chosen
but why, and what the alternative was.
