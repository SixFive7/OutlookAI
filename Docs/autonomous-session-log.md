# Autonomous session log - 2026-08-18/19

**What this file is.** The maintainer went to sleep mid-session and asked two questions to be
answerable on return: *why is it not done*, and *what did you decide on your own so I can review
it*. This file answers both. It is updated as work lands, and it is the first thing to read after
a context compaction.

**Working rule while the maintainer is asleep:** decisions already given are executed; anything
genuinely new is decided with the best available reasoning, recorded in section 3 with the
reasoning, and flagged for review rather than buried.

---

## 1. Decisions the maintainer gave, and where each stands

| # | Decision | State |
| --- | --- | --- |
| Tripwire strictness | Detect complications and re-run the affected tests, up to a maximum, rather than failing outright | **queued** |
| Frame size | Cap bodies at the COM layer | **shipped** `b46cf8a` |
| Second PST on the testbed | Add one via the tested helpers | **queued** |
| Live tier | Move the intermediate tier to the VM; keep the ability to run everything against the real system before a release | **queued** |
| `Stick-Test` VM + scratch | Delete both | **done** |
| Installed MCP server | Leave disabled until the release | **done, nothing to do** |
| Exhaustive scan | Resumable walk with a continuation token | **queued** |
| `thread` store asymmetry | Derive the warning from Outlook's store list; also scan for the same asymmetry elsewhere | **queued** |
| Timeout defects | Fold all three into the timeout-raising pass | **queued** |
| COM host kill | Keep the hard kill, document it, add a brief wait before killing, make the kill outcome-aware | **queued** |
| `top` ceiling | Leave at 100; rely on resumption | **decided, no work** |
| Remaining gap-map rows | Clear **all** of them before the release | **queued** |
| Work order | Infrastructure first: corpus, second PST, live tier on the VM | **in progress** |
| `update_draft` | **(d)** make it re-entrant: record intent first, so a retry completes rather than repeats | **queued** |
| Sweep timeout | **(d)** make expiry graceful **and** distinguish budget expiry from unresponsiveness at the supervisor | **queued** |
| H3 (undated mail invisible to the sweep) | Check whether DASL can express "or the property is absent" first; failing that, report it; full fallback enumeration only if it proves common | **queued** |

## 2. Timeout values agreed (measurement still pending for the sweep)

| Constant | Now | Target |
| --- | --- | --- |
| `ExhaustiveTimeBudgetMs` | 105 s | 600 s |
| `SweepBudgetMs` | 30 s | 180 s **(blocked, see below)** |
| `ThreadWalkBudgetMs` | 30 s | 180 s |
| `SearchIndexTimeoutSeconds` | 15 s | 60 s |
| `OperationDeadlineMs` | 120 s | 300 s |
| `ConnectDeadlineMs` | 90 s | 180 s |
| `MoveBatchBudgetMs` | 120 s | 240 s, strictly below its deadline |
| `HealthProbeDeadlineMs` | 5 s | **unchanged, deliberately** |

**The blocker:** `BudgetCompositionTests` asserts `SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs`
fits inside `OperationDeadlineMs`. At 180 s the sum is 195 s and the test fails before anything
reaches a mailbox, so the sweep cannot move alone - the operation deadline moves with it. That is
consistent with the maintainer's instruction ("serious headroom everywhere") but it is why the two
must be changed in one pass.

**Why `HealthProbeDeadlineMs` stays at 5 s** (an autonomous call, per the maintainer's "you decide
in context"): it is the diagnostic run precisely when Outlook is wedged. A health check that also
takes minutes turns every generous budget elsewhere into an unbounded wait with no way to find out
why.

## 3. Decisions taken autonomously - REVIEW THESE

1. **Overrode the corpus generator's date-fidelity refusal** with `--allow-undated`, reasoning that
   an all-recent corpus is still the sweep's worst case. **Wrong, twice over.** Undated items are
   invisible to a date-restricted sweep rather than recent, and separately the items were not in
   swept folders at all. Both faults are fixed in `6490ba9`; the refusal message that invited the
   bad inference is corrected and pinned. Cost: one wasted 12-minute build. No lasting damage.
2. **Reverted the VM to `CP-06-PRE-CORPUS` and deleted `CP-07-CORPUS-40K`.** The first corpus had
   40,000 items in Drafts with no dates. Reverting discards them in seconds where teardown would
   take a long time and would exercise a path unrelated to the measurement. `CP-07` snapshotted that
   broken corpus, so keeping it would only invite a later measurement against it.
3. **Built the corpus at 40,000 items** rather than a larger figure. The sweep caps at 200 items per
   folder, so beyond a few thousand per folder the sweep cost stops changing; scan throughput
   benefits from more, and the generator is additive, so growing it later is cheap.
4. **Ran read-only measurements against the maintainer's real production profile** (searches and
   health only, nothing created, moved or deleted) to size the budgets. Justified as read-only under
   the standing rules. It produced the finding that a 60-day exhaustive scan reaches 3 folders of 32
   before expiring.
5. **Kept `HealthProbeDeadlineMs` at 5 s** while raising everything else - see section 2.
6. **Did not investigate the 5,532 Outbox items** beyond confirming the profile could not send. The
   guard now makes the mechanism unreachable; understanding it is archaeology on a broken build.
7. **Accepted the corpus generator's broader safety rule** - refusing unless the profile has **no
   accounts at all**, with no override flag - rather than the narrower "no account can send from
   this store" I originally asked for. The object model cannot express the narrower rule, so the
   narrower version would read as a proof without being one.

## 4. VM state (`OutlookAI-TestVM`)

- Guest credentials for PowerShell Direct: `vmadmin` / `***REDACTED-CREDENTIAL***`.
- **PowerShell Direct lands in session 0, where Outlook can never finish starting.** Anything
  needing COM runs as a scheduled task with `-LogonType Interactive`, which lands in session 1.
- Checkpoints: `CP-01-WIN-CLEAN`, `CP-02-INSTALLER-STAGED`, `CP-03-OUTLOOKAI-INSTALLED`,
  `CP-04-OFFICE-GOLD`, `CP-05-ADDIN-TRUSTED`, `CP-06-PRE-CORPUS`.
- Store: one PST, `Outlook Data File`, **not in the local index** - which is the property the whole
  testbed depends on.
- **Placement and dates are both VERIFIED** on this store: items must be created in Drafts, saved,
  have the unsent flag cleared, saved again, then moved (`DraftsThenMoveWithSentFlag`); received
  dates written through the PropertyAccessor **do** drive DASL selection. The earlier "dates are
  impossible" verdict was an artefact of the placement bug.
- Host-side scripts live in `C:\Users\jori\Downloads\tmp-outlookai-vm\`; the corpus manifest is
  copied out to `corpus-vm1.jsonl` there, and **without it the corpus cannot be torn down**.

## 5. Measurements taken

- **Real profile, 5 stores** (`jori@huisman.io` alone: Archive 108,144 items, Sent 10,024): Inbox-only
  exhaustive scan 30.5 s; with subfolders 66.5 s; whole store 7-day window 36.6 s over 32 folders;
  whole store 60-day window **timed out at 105 s having scanned 3 folders of 32**.
- **Frame high-water on the real profile: 432 KB against a 64 MB limit**, zero refusals. The ceiling
  was set by the 30-second sweep timeout, not by any size limit.
- **The Windows Search provider does not put undated rows at the top of an ordered result** in
  either direction, and accepts the `1601-01-01` floor literal - so the displacement guard's
  recovery statement should essentially never fire. Not taken under a mapi scope, so the `T2` probe
  still stands.
- **Corpus build throughput:** 50.9 items/s without the move rung. With the move rung each item is
  written twice, so budget roughly double.

## 6. Still unverified, and why

- The sweep budget itself. The corpus exists to settle it and the measurement has not been taken.
- Whether the sweep recognises the substitute Junk folder the generator creates in a PST. If it does
  not, sweep measurements cover ~92% of the corpus rather than all of it.
- Everything in `TODO.md` marked as needing a live profile.
