# TODO

- [ ] **Three decided jobs, held until the agents now running have merged (decided by the
  maintainer 2026-10-03).** Each one touches files every open branch also touches, or stops the
  build VM they all share, so each waits for a quiet moment.
  - **Q123 - release notes in BrowserAI's style.** Copy the release-notes style, system and rules
    from the BrowserAI repository (`C:\Source\SixFive7\BrowserAI`) into this one: `CHANGELOG.md`
    conventions, `AGENTS.md` rules, and `Tools/Publish-Release.ps1`. The current Unreleased section
    (about 152,000 characters) is over GitHub's 125,000-character limit for a release body, so the
    next release cannot publish until this lands. Do it last, because it rewrites the Unreleased
    section every branch adds to.
  - **Q125 - security scanning without GitHub.** (b) Turn on the security analysers that ship with
    the .NET SDK in the builds, and triage what they find. Plus an exception to the Dependencies rule,
    granted by the maintainer: CodeQL may be run locally. Record the exception in `AGENTS.md` beside
    Q71/Q111, pin the CodeQL bundle by version and published hash, and add a script to run it.
  - **Q126 - the build VM in UTC.** Set `OutlookAI-Build` to UTC and take a new base checkpoint, so
    the non-live suite runs in a zone other than the workstation's - GitHub's runner used to catch
    zone bugs that way (Q95). Update the runner, its records and its pins. Hold the build VM's lease
    while switching.

- [ ] **On or after 2026-10-05, ask the maintainer whether the shared test mailbox exists (Q109).**
  He requested a free shared mailbox in his Microsoft 365 tenant on 2026-10-03 (for example
  `outlookai-test@xxlnet.nl`, with full access for `telefonie@xxlnet.nl`); creating it takes a few
  days, and he asked to be reminded after 48 hours. Until it exists, the six `Requires=DelegateStore`
  Exchange tests stay disabled on the Exchange test VM. Once it does: enable them there, and move
  test writes from telefonie into the shared mailbox wherever a test allows it (Q110).

- [ ] **Approve, amend or refuse the Exchange VM's Phase 2 write-safety design** (proposed
  2026-10-03, `Docs/live-tier-on-the-vm.md` section 4.4). Until then `OutlookAI-Exchange` is
  read-only - profile `ExchangeGuest` - and runs only `Writes=Nothing` tests (`Testbed/README.md`
  section 4e). Nothing in the design is built.

- [ ] **Decide how the one tagged leftover in telefonie's Sent Items goes.** The Exchange VM's
  read-only count (`T2/LiveExchangeHubArtifactTests`) found one item whose subject carries
  `OutlookAI-McpTest` in the mailbox's Sent Items on 2026-10-03 - from the workstation years, not from
  the VM. It fails that test on every run until it is gone, and only a tested sweep on an approved
  write run may remove it. Directions: (1) remove it with the existing tested sweep on the first
  approved Phase 2 run - it carries the tag, so `DeleteTaggedArtifactsUntilStableZero` takes it, but
  no run marker of a recorded run, so that would be a deletion by tag alone; (2) a one-off tested
  helper that deletes exactly the one EntryID the maintainer confirms; (3) he deletes it himself in
  Outlook on the web. Recommended: (3) - one item, his mailbox, and no deletion rule loosened for it.

- [ ] **Decide whether the Exchange VM gets the add-in.** Not installed (Phase 1 needed no write to the
  profile's configuration): `T2/LiveHealthTests` and the `AddInRegistry` tests need it, and on an
  Exchange profile its tuning reconcile writes Cached Mode policy values that change how the mailbox
  syncs - and an unelevated Outlook cannot finish that reconcile at all (the open item above). Decide
  with that item.

- [ ] **Four things the Exchange VM's first runs left open.** (4) Outlook's Object Model Guard
  stalls a run whenever Defender's signatures are stale (`Testbed/README.md` section 1d, item 4):
  `host/Invoke-ExchangeSignIn.ps1 -Mode Preflight` updates them before every run. Decide whether the
  VM also gets Q80's auto-approve policy (`guest/Set-OutlookProgrammaticAccess.ps1`), which removes
  the dependency but lets any process on that online VM read addresses without a prompt - so far
  only the offline guests have it. (1) The sign-in's verification-code
  step has never met a code page: Microsoft asked for no MFA code on either sign-in of 2026-10-03, so
  only the RFC test vectors stand behind the TOTP generator - the first code page will be its first
  real proof. (2) `T2/LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail` reads the add-in's
  tuning state without declaring `Requires=AddInRegistry`, so it fails on any machine without the add-in
  instead of being filtered out. (3) The five `T2/LiveSearchInTests` throw NullReference when the
  settings carry no `subjectOnlyProbe` instead of refusing with the remedy.

- [ ] **Decide what the add-in's tuning reconcile does with the five Cached Mode values it writes
  under `HKCU\Software\Policies` (found 2026-10-03 by the first guest run of the two-phase add-in
  install).** `OutlookTuningService.Reconcile` writes D25's five `caching.policy.*` values there, and
  that key is read-only to a NOT elevated token - so in the Outlook a user runs, the first of them
  throws, the reconcile's catch-all swallows it, and nothing after it runs: not the two user Cached
  Mode values, not the two OST size values, not `LastReconcileUtc` (`outlook_health` then reports
  `tuning.lastReconcileUtc` null). Only a machine where an administrator, a GPO or an earlier
  elevated Outlook already set the five values escapes it; the maintainer's workstation is one.
  Measured, with a control that isolates it, in `Docs/live-tier-on-the-vm.md` section 2.3. Until it
  is decided, a guest rebuilt by `Testbed/README.md` section 1 stops at step 7c `BROKEN`, where
  `T2/LiveHealthTests` would fail. Directions: (1) walk on past a value that cannot be written,
  record it - in `PolicyConflicts`, or a new "needs an administrator" list the settings dialog and
  `outlook_health` show - and always write `LastReconcileUtc`; (2) stop writing the Policies hive and
  keep only the user-hive values; (3) write the policy values from an elevated step - the installer,
  per-user today, or a one-time elevated helper; (4) change nothing in the product and set the five
  values in the testbed's elevated install phase, which hides the defect the way the old elevated
  `-Execute` did. Recommended: (1), then the proof again from `CP-08`, for `ADDIN-READY` with no
  control.
- [ ] **Recognise ANOTHER server session's lifetime pin on the show-me path's `ActiveExplorer()`
  branch (D49, found 2026-10-03, not measured).** `EnsureVisibleExplorer` refuses to display an
  Explorer `ActiveExplorer()` hands back only when `ComposeSurface.IsPin` knows it, and the pin
  registry holds IUnknown pointers - which for an out-of-process server are per-apartment proxies, so
  a pin another session made on another STA thread is very likely not recognised, whatever the
  registry's remarks say about being process-wide. The `Explorers.Add` branch no longer depends on it
  (the count check in `ComposeSurface.AddShowMeExplorer`, after the D49 probes of 2026-10-03), but if
  `ActiveExplorer()` can return a hidden Explorer at all, a second session would display that pin and
  the user's close would end Outlook again. Directions: (1) measure on a guest whether
  `ActiveExplorer()` ever returns a non-displayed Explorer; (2) if it does, recognise a pin by its
  window instead (`IOleWindow`, `IsWindowVisible`); (3) keep one pin per process, owned by the
  gateway rather than by a session. Recommended: (1) first - it is one probe on a guest.

- [ ] **Run the PST half of Q74 C3 on the indexed guest.** `LiveDecodeVerifyTests.ShortDecodedId_OpensAsTheItemItself_OnAPstStore`
  carries `Requires=SearchIndex`, so `OutlookAI-Unindexed`'s filter never selects it and the first
  guest live runs (2026-10-03) could not confirm it. It needs `OutlookAI-Indexed`, which was busy with
  Q99 that night.

- [ ] **Let the count tripwire's census read a table date by its column spelling too.**
  `CensusTableRow.ReadUtc` still calls the one-argument `ComDateValue.FromTableValue`, which takes every
  value as UTC; Q11's measurement (2026-10-03) showed the explicit `ReceivedTime` column - the census's
  first spelling - reports LOCAL time. The census only compares its own readings with each other, so
  the fingerprints stay consistent and nothing is mis-judged; only the instant it would print beside a
  departed item is off by the UTC offset. Pass the spelling (`CensusColumnMap` knows the index, the
  names list the spelling) when the census is next touched.


- [ ] **Decide what an UNSCOPED search does with hits from another Outlook profile's stores (Q99
  finding).** One Windows user has one search index for all of their Outlook profiles, so a search
  without `store` returns index hits from every profile - measured on `OutlookAI-Indexed`, where the
  tier profile's unscoped search returned the corpus store of `CorpusProfile`, under the name
  `Outlook Data File` that the tier profile's own store also had (`Docs/live-tier-on-the-vm.md`
  section 8 item 24). Such a hit cannot be opened from the open profile. Unchanged by Q99, and harmless
  on a one-profile machine. Directions: (1) leave it, and document it; (2) FLAG such hits from the
  store map (a hit whose store root no store of this profile claimed); (3) DROP them; (4) scope an
  unscoped search to this profile's roots. Recommended: (2). The Exchange half it waited for is
  measured (2026-10-03, Q113 (b), on the Exchange test VM): a cached Exchange mailbox is tied by its
  store hash, input `profileMappingSignature`, so a root the hash did not tie is not this profile's own
  cached Exchange store. Delegate roots stay `delegateFolder`; a delegate's row is not measured yet.

- [ ] **Let a folder whose name holds `/` be named in a `folder` argument (Q99 folder finding,
  2026-10-03).** Outlook accepts `/` in a folder name and the index spells it `%2F`, so the folder is
  searchable - but every tool's `folder` argument is a `/`-separated path, so `Parent/a/b` means three
  folders and the one called `a/b` cannot be named. Today it is reached through its parent
  (`include_subfolders`, the default), its hits report `Parent/a/b` and open, and `list_folders` lists
  it as `Parent/a/b` - a path that reads as nesting and, passed back, finds nothing (the zero-row guard
  then says so). Measured: `Docs/live-tier-on-the-vm.md` section 8 item 26. Directions: (a) leave it,
  documented (`McpServer/README.md` fact 17); (b) an escape inside a segment (`\/`, or `%2F` itself),
  parsed by `search`, `move_mail` and the exhaustive scan and written by `list_folders` - one parser
  and one renderer; (c) a segment-array argument beside `folder`; (d) when a path does not resolve,
  retry with adjacent segments joined by `/`. **(a) is in place, decided on the maintainer's behalf
  2026-10-03** - it changes no tool's contract overnight, and the parent route works. **Recommended
  next: (b)**, which is what keeps this item open; undo (a) by doing it.

- [ ] **Measure a STORE name holding `% / \ * ?` in a display path (Q99 folder finding).** The index
  writes FOLDER names into `System.ItemFolderPathDisplay` and `System.ItemPathDisplay` as names, not
  URL spellings (section 8 item 26), and the product now derives a non-recursive folder search's
  display path the same way for the STORE part too (`MapiItemUrl.TryBuildFolderPathDisplay` decodes
  it). That half is inferred: no guest store has one of the five in its name since the Q99 store
  went with its checkpoint. The next time a PST named like `q99 50% off*?x` is attached (item 24's
  route), read one item row's `System.ItemFolderPathDisplay` and run a folder search in it with
  `include_subfolders: false`; if the store part is the URL spelling, stop decoding it there.

- [ ] **Put a release candidate's MCP server in front of one session on the workstation, without installing it.**
  The manual pre-release checks (`Docs/release-manual-checks.md`, Q74 D2) exercise the Exchange-only
  write paths through the MCP server, by hand, on the maintainer's own profile - and nothing can put a
  candidate's server there today: `Tools/Switch-AddInBuild.ps1` copies and registers the ADD-IN only and
  never touches the server. Until this exists the list runs right after a release is installed, with the
  previous installer kept for rollback. Recommended shape: the same copy-then-register discipline as
  Q81 - a server build copied out of any build folder and registered for one Claude Code session, never
  globally - so it can never become the server every other session starts.

- [ ] **What still stops a rebuilder rebuilding the test VM from this repository alone.**
  `Testbed/` is the entry point and holds the runnable half - parameter set, host and guest
  scripts, the settings template, the credential contract - and `Tools/Checks/check-testbed-references.ps1`
  fails the build when a document names something the repository does not contain. The corpus
  parameters are now recorded and verified: **`vm2` / seed `7777` / anchor `2026-08-19` /
  20,000 items, default shape**, recovered from the manifest header on the guest and confirmed by
  re-running `corpus-plan`, which reproduced the per-folder and seven-day counts the measurement
  docs quote. Every doc that used `vm1 / 4242 / 2026-08-01 / 40000` was quoting an EXAMPLE.

  What is left, in the order it blocks a rebuild:

  - [ ] **Answer the open questions in `Testbed/README.md` section 6.** They are the facts that are
        genuinely not recorded anywhere - Hyper-V spec, Windows edition, Office version and bitness,
        which Outlook profile is default and how the switch is automated, whether the three-store
        layout or the mail sink exist at all. Each needs a guest or the maintainer; none can be
        derived from the repository. Since 2026-10-03 only the questions that apply to the
        script-built guests are left to answer: those section 6 asks only about the original
        guest, `OutlookAI-TestVM`, went unanswerable when it was deleted that day.

        **Two corrections, 2026-09-15, because this line had drifted from the section it points at.**
        (a) It said **eleven** questions; section 6 now numbers **16** items - 13 questions plus 3
        things nobody can put in a repository - and several carry their own "half-answered" notes.
        Quote the section, not this count. (b) **"The second Windows account" is no longer one of
        them.** Section 1.1 of the runbook used to need two Windows logons per machine, because
        Windows Search indexes a `mapi16://{SID}/` scope per ACCOUNT; the build is now two GUESTS
        (`OutlookAI-Indexed` / `OutlookAI-Unindexed`), so index state is a property of the machine
        and Indexing Options controls it unambiguously (`49e563e`). Stated carefully, because these
        are different facts and only the second happened: the per-account index assumption was NOT
        disproved - the design stopped depending on it, which also retired the riskiest unverified
        assumption in the whole layout.
  - [ ] **Fold the recovered facts into `Docs/live-tier-on-the-vm.md`.** Its section 8 lists ~20
        open items; the corpus parameters (item 15), the PST path and display name (item 11), the
        scheduled-task recipe (item 10) and how results leave the guest (item 13) are now answered
        elsewhere in the repository, and its worked examples still use the `vm1` corpus.
  - [ ] **The history rewrite for the leaked guest password is still outstanding.** The value is
        dead - rotated 2026-08-24, and the current one is provably absent from HEAD and from every
        commit - but the old username-and-password line remains reachable in history from
        `d499bf1` to `54ecd26^`. Tracked in `Docs/autonomous-session-log.md`; noted here because a
        rebuilder reading the testbed docs will find the reference to it.

- [ ] **Residual questions from the 2026-08-19/20 atomicity-claims sweep.** All 31 claims of
  non-effect were enumerated, 16 were wrong and all 16 are fixed (`Docs/completeness-gaps.md`
  section 7b, T1 `AtomicityClaimsTests`). These are the things reading could not settle, and the
  two wrinkles beside the send path that are NOT the audited defect.

  - [ ] **Which HRESULTs Outlook actually raises from `MailItem.Delete()`, `Move()` and
        `Display()` when the RPC channel breaks mid-call.** The may-or-may-not reading of
        `RPC_S_CALL_FAILED` is documented and is this repo's own stated basis for the retry
        classification, but whether these three calls produce it in practice is unmeasured. It
        decides whether the unknown-outcome wording on those three paths describes a rare event or
        a theoretical one. **The wording was deliberately fixed WITHOUT waiting for it**, on the
        maintainer's reasoning: a claim about what did not happen should not rest on an unmeasured
        probability. Measurable on the VM with `OUTLOOKAI_COMHOST_FAULT` plus a read afterwards.
  - [ ] **Whether a MUTATING response can actually exceed the 64 MB frame.** Row J8's verdict is
        "false by construction": `ComDraftUpdateResult` and `ComDraftCreateResult` carry no body,
        only recipients and attachment metadata, and no realistic 64 MB case could be constructed
        by reading - a recipient count in the hundreds of thousands would be needed. The serialised
        size of a `ComDraftUpdateResult` is measured nowhere. The fix is one `if` either way, so
        this is a curiosity rather than a blocker.
  - [ ] **Whether `TryDiscardDraft` can reach its catch-all through a NON-`IsComCallFailure`
        exception after `Delete()`.** A `NullReferenceException` or `InvalidOperationException`
        from `collection.Count` or the indexer inside `TryFindDiscardedCopy` would escape
        `_runner.Run` entirely rather than becoming `com_failure`, and what `PumpedStaRunner` does
        with such an escape was not established. It changes which message the caller gets, not
        whether the delete happened.
  - [ ] **Whether soft-deleted drafts actually survive on the maintainer's profile.** The discard
        fix leans on recoverability from Deleted Items. Outlook's "empty Deleted Items on exit"
        option and Exchange retention tags can both remove the item, and neither is visible from
        this codebase. The message says to look in Deleted Items, which is right in either case;
        what is unknown is how often looking will find anything.
  - [ ] **`send` catches `TimeoutException` only** - noted by the audit and NOT part of the defect
        it was auditing, so deliberately left alone in that pass. A child that dies for any other
        reason raises `ComHostUnavailableException`, which is not a `TimeoutException`, so that
        path is saved by `ComHostSupervisor.DescribeInterruption` instead. The outcome is right by
        a different route rather than by this `catch`, which means the send path's own
        `send_outcome_unknown` audit line is NOT written for it. Worth widening the filter, or at
        least writing down that the audit line is conditional on which way the child died.
  - [ ] **`SendUsingAccount` is written to the user's draft before the identity readback and is
        never restored.** The "Nothing was sent" claims stay true, and the messages now SAY the pin
        may have been rewritten (rows 9 and 10) - but saying so is a smaller fix than restoring it.
        Restoring needs the previous value captured before the putref and put back on every failure
        path, which is one more mutating call on a path that is currently refusing to mutate
        anything, so it wants a decision rather than a quiet fix.
  - [ ] **Seven decision lines from this pass are provably unguarded, established by mutation
        rather than assumed.** Each was reverted, built, run against the whole non-live suite and
        restored; 30 of 37 were caught, and eight of those thirty only after the gap they exposed
        was closed with a new test. The full table is in `tmp-aitrace/mutation-table.md`.
        - **Five live in `OutlookComSession`, behind a COM call no non-live test can execute**,
          which is the same class as `sortApplied` and the attachment-plan execution already
          recorded above: the saved-draft id being published on the failure path and re-read after
          the relocate (row 5's COM half), `TryMoveItemToPath` reporting created folders on the
          failure path (row 6's COM half), `TrySaveAttachment` reporting the attempted path, and
          the size read being best-effort (row 11's two halves). Every T1 test substitutes the
          session, so a test can only prove that the SERVICE layer uses what it is given. What
          each is worth is not in doubt - the shapes are read off the code - but they are
          unexercised until a live run. The cheap substitute is the one already used elsewhere: a
          temporary build that forces the branch.
        - **`CompleteMove`'s audit-failure branch** (`Ok=false` over an item that MOVED, reported
          as `outcome: applied`) needs `AuditLog.Append` to FAIL, and it writes to
          `%LOCALAPPDATA%\OutlookAI` through a path that is not injectable. `AppendTo` takes a
          directory and is used by `AuditLogTests`; wiring the service layer to it would make this
          reachable, and is a bigger change than the row it guards. (2026-10-03: no wiring is
          needed any more. Since Q86 every test process appends to its own throwaway log,
          `AuditLog.EffectiveLogPath`, so a T1 test can make `Append` fail by holding that file
          open without sharing for the duration of the call - `T1/AuditLogToolTests` already
          locks it that way.)
        - **The supervisor's own wiring for the interrupted-request outcome** needs a child that
          dies while holding a request. Both ends are pinned separately - the value
          (`MutationOutcome.ForInterrupted`) and the carrier
          (`ComHostUnavailableException.Outcome`) - so what is unguarded is the line that joins
          them.

  - [ ] **The `outcome` field is not consumed by anything yet.** It is additive and null-omitted,
        the tool descriptions teach it, and T1 asserts a `com_failure` never carries `unchanged` -
        but no agent behaviour depends on it, so its value is currently "the claim is testable"
        rather than "the claim is acted on". If it turns out nothing ever branches on it, the
        honest conclusion is that the prose was the whole fix and the field is cost.

- [ ] **Residual gaps left by the 2026-08-19 timeout pass.** The values, the three inventory
  defects, the graceful sweep expiry and the kill work all landed; these did not, and each one
  is written down because a mutation check proved it unguarded rather than because it was
  guessed at.

  - [ ] **Nothing in the non-live tier notices if `ComGateway`'s budget overload goes back to
        being a pass-through.** Reverting `BudgetedSessionProxy.Wrap(session, budget)` to
        `session` leaves all 1,887 tests green. The MECHANISM is pinned (T1
        `InProcessBudgetTests` drives the proxy directly, and removing its dispatch check
        fails); the WIRING is not, because exercising it needs a real COM session. The live
        tier is where it is exercised, which is the tier that had no budget at all until this
        pass. Options: accept and rely on T2; add `InternalsVisibleTo` to `OutlookAI.Core` and
        pin the wiring through an internal seam; or a structural IL assertion, which is
        fragile and unlike anything else here. (2026-10-03: the `InternalsVisibleTo` now
        exists - Q86 added it for the audit-log redirect - so the second option costs only
        the seam.)
  - [ ] **The grace values themselves are unmeasured.** `CleanExitGraceMilliseconds` (250)
        and `ShutdownExitGraceMilliseconds` (2000) are judgements: nobody has timed how long
        `OutlookComSession.Dispose` takes against a real Outlook, so nobody knows whether
        2000 ms is generous or short. Measuring it needs a live-tier run that times the
        child's exit with a session attached. A wrong value degrades to today's behaviour
        (the child is terminated and releases nothing), which is why it shipped unmeasured.
  - [ ] **`MailService.SearchIndexTimeoutSeconds` is pinned only from above.** T1 asserts it
        never exceeds `OleDbIndexClient.DefaultCommandTimeoutSeconds`, so reverting 60 to 15
        fails nothing. A lower bound would need a measurement constant for "how long a
        statement on a saturated indexer legitimately takes", and no such measurement exists.
  - [ ] **`T3/McpStdioClient` still has one budget for a whole session.** One
        `CancellationTokenSource` bounds every read and write for the client's lifetime, so it
        is a session budget masquerading as a per-call one. The exhaustive-scan live test now
        passes an explicitly derived budget, which is the case that would have broken first;
        the general split (session budget plus a per-`RoundTripAsync` budget, both named)
        is still open. Raising the DEFAULT is deliberately not the fix - it is the non-live
        run's only safety net against a hung stdio test, and the build VM gives a whole run 60
        minutes (Testbed/host/Invoke-TestsOnBuildVm.ps1 -RunTimeoutMinutes).
  - [ ] **Claude Code's 30-minute stdio idle abort is now the nearest client-side limit, and
        nobody owns it.** A 600 s exhaustive scan is 600 s of complete silence on the pipe -
        this server sends no progress notifications. It fits (600 s < 1800 s idle < the
        ~27.8 h per-call hard ceiling), but the idle limit is a client default nobody here
        chose, no test watches it, and a user who sets a per-server `timeout` in `.mcp.json`
        for a sensible reason will land far below 600 s. Re-measure instruction: `grep -a` the
        shipped `~/.local/share/claude/versions/<ver>` binary for
        `CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT` and resolve the stdio branch's constant.
  - [ ] **The graceful sweep expiry has never fired against real mail.** It is pinned at both
        ends in T1 - the pure boundary (`OutlookComSession.SweepBudgetSpent`) and the whole
        reporting chain from `ComSweepResult` to the advice sentence - but the walk that stops
        exists only over live COM folders. With the budget now at 165 s inner / 180 s outer and
        a measured whole-profile sweep of ~60 s, it should never fire in ordinary use, which is
        the point and also the reason it will stay unexercised.

- [ ] **Residual gaps left by the 2026-08-19 re-entrant `update_draft` pass.** The intent record,
  the idempotence key, the attachment reconciler and the add-before-remove reorder all landed and
  are pinned in T1 `DraftUpdateReentrancyTests`. These did not, and each is here because a mutation
  check proved it unguarded rather than because it was guessed at.

  - [ ] **The attachment enumeration guard is provably unguarded** (mutation M15, 2026-08-19):
        making `TryUpdateDraft` enumerate the draft's attachments even when the request touches
        none leaves all 1,921 tests green. It is a cost guard - one COM call per attachment on
        every body-only revision - with no observable payload, so nothing outside a live profile
        can see it. Options: accept and rely on T2; or count contract-level COM calls in a live
        fixture, which nothing does today.
  - [ ] **The COM-side EXECUTION of the plan is unguarded by any non-live test.** `DraftAttachmentPlan`
        is pinned state by state, and the two decisions the COM sequence makes were lifted out so they
        could be reverted (`BuildForAttempt`, `ComDraftUpdateResume.ThreadIndexFor`). What no T1 can
        reach is `RemoveAttachments` deleting the N lowest-indexed instances of a name, and the fact
        that additions now run before removals - both are `dynamic` COM against a live
        `Attachments` collection. Reversing the order back leaves the suite green. T2
        `LiveUpdateDiscardTests` is where it would be exercised; extending it to assert the ORDER
        needs a way to observe a mid-sequence state, which nothing has today.
  - [ ] **Nothing verifies that an interrupted attempt's partial writes are DURABLE.** The whole
        design is deliberately indifferent to it - it converges on the end state from whatever the
        draft is observed to hold - but the question is still unanswered and worth answering, because
        it decides how often the resume does anything at all. Whether Outlook keeps an unsaved
        `Attachments.Add` / `Attachment.Delete` when the automation client dies is undocumented by
        Microsoft (`tmp-aitrace/kill-safety.md` section 2.3). Measurable on the dev machine with
        `OUTLOOKAI_COMHOST_FAULT=hang:TryUpdateDraft` plus a read of the draft afterwards.
  - [ ] **The gap map's remaining line-number citations are stale.** Checked 2026-08-19: all nine
        `OutlookComSession.cs` references had drifted (converted to symbols in this pass), and the
        `MailService.cs`, `MailModels.cs` and `OutlookTools.cs` ones are stale too - `MailService.cs:220`
        lands in an unrelated comment, `:612` and `:3995` on bare `</summary>` lines. Convert each as
        its row is next touched; a stale number reads as evidence and points the next reader at
        unrelated code.

- [ ] **PARTLY FIXED 2026-09-15 - option (2) shipped; option (1) is the COM half and still needs a
  live run.** Fix the census's identity-to-count degradation, which is why the tripwire needs a
  re-census at all on a stable mailbox. Found by reading during the 2026-08-24 investigation.

  **What shipped (option 2, "carry the reason").** `FolderCensus.CountOnly` now REQUIRES a
  `CensusCountReason`, and refuses `Walked` - so a count-only reading whose reason says it was
  walked cannot be constructed, and "reason not recorded" cannot quietly become the commonest
  answer. `CensusIdentityPlan.TryIdentify` reports which of the six refusals applied, the walk
  failure path reports `TableUnusable`, and `T2/CensusReadingStrength` - pure - turns a PAIR of
  readings into the clause the verdict carries. Three things follow:
  - **The wrong sentence is gone.** `EvaluateByCount` said `(folder above the identity budget)`
    whatever the cause. That was right in one case of six and pointed the reader at the wrong
    number in the other five - in an emergency, in a message read once, by somebody who believes
    mail has just been deleted.
  - **A degraded pair says it is degraded.** An `ITEMS LOST` over a folder the baseline identified
    and this pass could only count now says the POST-RUN reading was WEAKER, names why, and says
    the same mailbox state can read as a loss on one pass and as filed on another for that reason
    alone - so "this folder lost items" and "this reading was weaker than the one it is compared
    against" stop printing identically.
  - **The silent case is noted.** A degraded pair whose count did NOT move is where the guard
    quietly loses its teeth: the baseline could tell a filing from a deletion there and this pass
    cannot, so an item removed while another arrived is invisible. Noted, never failed - nothing
    was observed to leave.

  Pinned by 22 new T1 tests (`T1/CensusReadingStrengthTests`), including a source read of the live
  census's call site, which is the half no CI test can execute. **Nothing about which runs FAIL
  changed**: the same deltas fail, with a verdict that now says what it is made of.

  **Still open: option (1), re-walk the degraded folder before concluding anything.** That is the
  only one that removes the false failure at its SOURCE rather than explaining it, and it is one
  COM table read inside `CaptureFolder` - so it cannot land without a live run to check it.
  Option (3), refusing to compare a degraded pair at all, stays rejected: noisier by a lot, and it
  would fail runs for a property of the census rather than of the mailbox.

  <details><summary>Original entry</summary>

  **The defect.** Whether a folder is compared BY IDENTITY or BY COUNT is decided independently
  on each pass, and the post-run decision is timing-dependent. `CensusIdentityPlan.Repeating`
  refuses a folder that grew past `500 x 4`; `ShouldIdentify` refuses one once the 120 s
  identity budget has expired for that store; and `WalkFolderItems` returns null - meaning
  "count this folder" - on any COM call failure, a missing column, a row with no EntryID, a
  duplicate EntryID, `EndOfTable` false, or a count that moved under the read. Any of those on
  the post-run pass but not the baseline silently switches the folder from
  `EvaluateByIdentity` to `EvaluateByCount`.

  **Why that matters.** The count rule cannot exonerate a filing and cannot see a departure
  masked by an arrival. So the SAME mailbox state yields `note: filed (not loss)` on one reading
  and `ITEMS LOST` on the next, purely because the second reading was weaker. That is a census
  that reports differently twice with nothing having changed - the thing the original "treating
  as enumeration noise" message blamed on COM - and it is exactly the shape a re-census 30 s
  later would clear, which is how it has been hiding.

  **It also lies about itself.** `EvaluateByCount`'s failure text says
  `(folder above the identity budget)`, which is one of at least five possible reasons and is
  the wrong one whenever the cause was the clock, a transient COM failure or an unusable table.

  Three ways to fix it, cheapest first:

  1. **Re-walk the folder, not the run.** When a folder the baseline identified degrades on the
     post-run pass, walk that one folder again immediately - one table read - before concluding
     anything. Removes the timing dependence at its source and is the only option that makes the
     re-census rung unnecessary for this class.
  2. **Carry the reason.** `FolderCensus.CountOnly(count, reason)`, and a failure that names it:
     "the post-run reading of this folder was WEAKER than the baseline's (identity time budget
     expired)". Does not fix the false failure, but makes it self-explaining and removes the
     wrong sentence.
  3. **Refuse to compare a degraded pair at all** - treat baseline-identified plus
     post-run-counted as an unmeasured folder and fail the run for that reason instead. Strictly
     honest, and noisier than (1) by a lot.

  (1) and (2) compose and are the recommendation. Nothing here was implemented: it changes what
  the guard proves, which is the maintainer's call.

  </details>

- [ ] **Two things `Docs/live-tier-on-the-vm.md` should say, found while answering "can the
  tripwire be always correct?" (2026-08-24).** Not edited here - another agent owns that file.
  - **AutoArchive must be OFF in the VM's Outlook profile**, and it belongs in section 2 beside
    the other profile settings. It is a client-side actor that moves items out of a PST on a
    schedule, which is indistinguishable from mail loss to a before/after census, and it is the
    one such actor a test machine can have.
  - **Section 7's prediction is now wrong in one respect:** on a `Portable` profile the retry
    ladder does not run at all (`TripwireRetryPolicy.None`), so a suspected loss on the VM fails
    on the first reading with `NO RE-CENSUS IS CONFIGURED` rather than after two re-censuses and
    a re-run. Open item 18 ("the re-census-then-re-run policy is decided and not built") is now
    closed in both halves.

  **Four more, from the bystander work (2026-08-24).** Same reason - not edited here.
  - **Section 1.3 and section 2.6 must say the bystander is DECLARED, not merely listed.**
    "The bystander is listed in `expectedStoreDisplayNames` and is never the hub" is now half
    the instruction: it also has to be named in the new `bystanderStoreDisplayNames`, and it
    must stay in `expectedStoreDisplayNames` (that is what censuses it and what
    `list_accounts` exactness counts). Listing it in only one of the two is a refusal, not a
    warning.
    **CORRECTED AND DONE 2026-09-24 (Q77): that last sentence was wrong.** The census adds every
    declared bystander back in, so one left out of `expectedStoreDisplayNames` is still watched and
    is not refused, and a store only in `expectedStoreDisplayNames` is the identity account's
    legitimate shape; the tier refuses only when no watched store is both non-hub and write-denied.
    The maintainer chose to correct the documents rather than the loader, and sections 1.3, 2.6 and
    2.10 now say what the code does. The 2.10 example and field table below are done too - with the
    split into a watched and an indexed list (Q70) on top.
  - **The section 2.10 settings example needs the new key**, matching
    `Testbed/live-test-settings.example.json`: `"bystanderStoreDisplayNames": [ "OutlookAI
    Bystander" ]`. Without it a rebuilder follows the runbook and gets a bystander the suite
    may still write to, which is the defect this closed.
  - **Section 6's "what to read in the output" gains a line**, printed straight after the
    settings line: `[tripwire] watch soundness: N declared bystander(s), M store(s) this census
    can fail on, K watched store(s) the suite may still write to`. `M=0` is the machine-readable
    form of section 7's "the guard runs and proves nothing", and it says so in the same words.
  - **Open behaviour item 20 is closed** and should say how: not by narrowing the grant - two
    live tests use it - but by a declaration the allowlist honours and the tripwire verifies.

  **Four more, from the refuse-a-vacuous-census work (2026-08-24).** Same reason.
  - **Section 1.3 must say the bystander is MANDATORY, not advisable.** "The bystander is the one
    people leave out, and the tripwire is useless without it" now understates it: without one the
    tier refuses to start, and no setting turns that off. Same for the sentence in section 2.6.
  - **Section 7's "with one PST that is also the hub" prediction no longer describes a run.**
    That configuration never reaches a census - `EnsureBaseline` throws `NO STORE THIS CENSUS
    WATCHES CAN PRODUCE A FAILURE` before the health gate and before any COM call. The paragraph
    should say so, because as written it tells a rebuilder what to expect from a run that cannot
    happen.
  - **Section 6's `[tripwire] watch soundness:` line can no longer report `M=0`** on a run that
    got as far as printing anything, so the "M=0 means the guard proves nothing" reading belongs
    with the refusal instead.
  - **The section 2.10 settings example and field table need the corpus store in
    `bystanderStoreDisplayNames`**, matching `Testbed/live-test-settings.example.json` and
    `Testbed/README.md` §3b: `[ "OutlookAI Bystander", "Corpus A" ]`. The table's
    `expectedStoreDisplayNames` row should also say that a declared bystander the profile does
    not mount refuses the tier - which is why `Corpus B`, ~~in the other Windows account's
    profile~~ **on the other GUEST** (corrected 2026-09-15: the two-Windows-accounts layout was
    replaced by two guests, `OutlookAI-Indexed` with `Corpus A` and `OutlookAI-Unindexed` with
    `Corpus B` - `49e563e`), goes in that machine's settings file and not this one. The reasoning
    is unchanged and now easier to state: it is a different MACHINE, so it is a different settings
    file.

- [ ] **PENDING TASK - process `C:/Source/SixFive7/BrowserAI/.work/truncation-prompt-for-sibling-project.md`.**
  *(Path corrected 2026-09-15: it was written with a Windows backslash before `truncation`, which this
  file stored as a literal TAB - so the path as printed named a file that cannot exist and the `t` was
  missing from the name. Forward slashes here so it cannot happen again. The file DOES still exist at
  the corrected path, checked 2026-09-15, so this item is genuinely outstanding rather than moot.)*
  The maintainer asked for this at 09:00 on 2026-08-18. It is expected to be the portable
  description-budget prompt written for another project; read it and act on what it asks for. Recorded
  here because auto-compaction was imminent when it was requested.

- [ ] **Re-run the unindexed-store probes on a MIXED profile - the one shape no machine here has.**
  Group A and E of `Docs/completeness-gaps.md` are now all closed (A1-A5, E1). Everything about
  them has been verified on two profile shapes: the fully-indexed developer profile, and the
  Hyper-V VM whose only store is an unindexed PST. **The shape that carries A1's residue is
  neither of those** - one INDEXED mailbox plus one UNINDEXED data file, so that the profile-wide
  frontier probe succeeds while one store is still absent from the index catalog. That is the
  ordinary "Exchange account plus archive.pst" desktop, it is what T1
  `UnindexedStoreReportingTests` models with a stand-in index client, and no machine to hand can
  produce it live. What T1 cannot exercise is the real chain: a live `DiscoverStoreScopes` whose
  sample returns one store's prefixes and not the other's, and `StoreHasIndexRows` answering
  false against a real SystemIndex for a mounted PST.

  Mount a PST on a profile that already has an indexed account (or add an account to the VM), do
  not add it to Indexing Options, then confirm READ-ONLY, with `search` only:
  - a plain unscoped search names it in `sweep.storesWithoutIndex` with `no_index_frontier`
    (this already worked - it is the regression guard);
  - an unscoped search with `before` older than 7 days reports `sweep.notNeeded:true` **and**
    `indexFrontierMissing:true`, `freshness:"partial"`, `degraded:true` - it used to say
    `freshness:"live"` with nothing else;
  - an attachment-only search reports `freshness:"index-only"` **and** `no_index_frontier`;
  - `outlook_health` shows the PST as `index.perStore[].inLocalIndex:false` with a problem line,
    while the indexed account still reports a frontier.

  Also worth measuring on that profile: the cost of the one COM store-list read the three
  no-sweep paths now pay (5-minute cached, so expected to be a pipe round trip), and whether
  `StoreIndexProbeBudgetMs` (1.5 s) is ever the thing that cuts the probe short.

- [ ] **An answer too big to frame: refused (done), but not yet prevented (open).** Found by the
  boundary audit on 2026-08-18. **The refusal and the measurement landed 2026-08-18; the four
  responses below are still the maintainer's to choose between, and nothing here pre-empts them.**

  **Done - (a) the refusal.** `ComHostServer.ServeAsync` now catches the framing refusal around
  its write and answers with a `ComHostResponseTooLargeException` frame naming the operation, the
  size and the limit; the serve loop keeps serving, so one oversized answer costs one call instead
  of the session. `ComHostErrorMapper` rebuilds the type, and `OutlookTools.GuardAsync` returns it
  as `ResponseTooLarge` with advice to NARROW the request rather than retry it. The testing
  objection below is answered by the fourth option it did not list: `ComHostServer` takes its
  ceiling as a constructor parameter defaulting to `ComHostProtocol.MaxFrameBytes`, so T1 reaches
  the branch with a 64 KB answer against a 4 KB ceiling - real serve loop, real framing, real
  mapper, no 64 MB allocation. Measured 2026-08-18: the two refusal tests run in 73 ms and all six
  new tests in 99 ms, against a T1 tier of 1 m 55 s.

  **Done - the measurement.** `ComHostFrameMeter` keeps a per-process high-water mark of the
  largest payload seen (both directions: encode and read, because the big frames are built in the
  child whose counters die with it) plus a refusal count. `outlook_health` reports both beside the
  limit, as `comHost.largestFrameBytes`, `frameLimitBytes` and `framesRefusedTooLarge`, and a
  refusal also raises a `problems` line. Lifetime is one SERVER process and survives child
  restarts, deliberately: the child is restartable and the question is about the product. So the
  number that says whether 64 MB is right is now collectable from any running install - it was
  previously unmeasured, which is why the entry below argues from the caps instead.

  **Done - (c), chosen by the maintainer 2026-08-18: cap bodies at the COM layer, so a frame that
  cannot be sent cannot be built.** (b) and (d) were the alternatives and were NOT chosen -
  `MaxFrameBytes` is unchanged and the sweep result is still one frame. Two bounds, both in
  `OutlookComSession`, both applied in `SnapshotBrief` where the body is read:
  `SweepBodyCharsCap` (500 000 chars, = `MailService.BodyCharsCap`, the largest body window `read`
  will ever return in one call) per ITEM, and `SweepBodyBytesBudget` (32 MiB, = `MaxFrameBytes / 2`)
  across the whole sweep. The per-item cap alone provably cannot do the job: the item count is
  4 folders x `SweepPerFolderCap` x every store, and the store count is unbounded, so 800 items at
  500 000 chars is already ~400 MB of theoretical worst case on ONE store. What it does instead is
  keep the budget FAIR - one enormous mail cannot spend it all and blind the sweep to everything
  behind it. The budget is counted in encoded BYTES rather than characters because the pipe escapes
  every non-ASCII character to six bytes: a character budget would have to be sized for that case
  and would then bite at ~5.5 M characters, which an ordinary unindexed-PST sweep really reaches.

  **The cut is never silent, and it is not a display truncation.** These bodies are matched against
  (`FreshMerge.MatchesTerms`) and shown to nobody, so unlike `read`'s windowing - which pages and
  loses nothing - a cut here can make a search MISS a real match. `ComMailBrief.BodyTruncated`
  carries the fact per item; `sweep.itemsBodyCapped` and `sweep.itemsBodyCappedUnmatched` reach the
  payload; the `body_cap` coverage code is raised on the INTERSECTION alone (cut AND unmatched),
  since a cut body on an item that matched anyway cost nothing; `sweep.bodyBudgetExhausted` says
  which bound cut, because the remedies differ. The two facts CAN be told apart and are; what
  cannot be settled is whether the term really sat past the cut, since that needs the text the
  bound refused to carry, so the sentence says "may be" and never "is". T1 `SweepBodyCapTests`
  (18 tests) plus two wire round-trips in `ComHostProtocolTests`, mutation-checked with 13 separate
  reverts, each disabling one decision.

  **What remains reachable, stated rather than implied.** (i) The NON-body half of a frame is still
  unbounded in the store count: per-item EntryIDs, StoreIDs, subjects and folder names are ~1-2 KB
  each, so ~40 mounted stores each holding 200 items in all four arrival-path folders inside the
  window would build an oversized frame carrying no body text at all. Never observed; the typed
  refusal is the backstop. (ii) The byte budget is enforced against an OVER-estimate of the encoded
  size (`EncodedBodyByteCeiling`), so it errs toward cutting early, never toward a frame that will
  not send. (iii) The bound has only ever been exercised in T1 - see the live-profile item below.
  (iv) `read` is DELIBERATELY not capped at the COM layer: `TryReadItem` still returns the whole
  body, plus the whole `HTMLBody` when `include_html` is set, so one pathological mail could in
  principle build an oversized `read` frame. Capping it there would break the one contract that
  makes `read` lossless - `bodyTotalChars` is measured from the full body and `body_offset` pages
  the whole of it, so a COM-side cut would make the total a lie and the tail unreachable. The
  measured `read` payload is ~0.5 MB, 0.8% of the limit. The sweep was the case worth closing
  because its size is driven by MAIL VOLUME rather than by one mail, and because its bodies are
  never shown, so a cut there costs matching rather than reading.

  **The other two frames that carry mail were checked and need nothing.** `ExhaustiveScan` and the
  `thread` walk both snapshot briefs with `includeBody: false` - the exhaustive tier matches
  server-side through DASL and the thread walk needs no body at all - so the sweep is the only
  frame in this server that ever carried body text in bulk.

  **MEASURED 2026-08-18 on the real 5-store profile, with that high-water mark** (read-only:
  `outlook_health`, `list_accounts`, four searches; nothing created, moved or deleted). Largest
  frame **441,930 bytes - 432 KB, 0.66% of the limit, about 152x headroom**; zero refusals.
  **This corrects the derived worst case below:** "reachable by ordinary use" is too strong. The
  filter-only search, the one that should have swept hardest, **timed out** - `Outlook did not
  respond to 'SweepFoldersNewerThan' within 30000 ms` - the supervisor replaced the COM host, and
  the search degraded to `index-only` and still answered with 100 hits. So **the 30-second sweep
  budget bites long before the 200-items-per-folder cap does**, and the cap arithmetic is the wrong
  worst case for an Exchange store. **SUPERSEDED IN PART, 2026-08-19: the 30 s budget is now 180 s,
  and with the bound lifted the frame measurement changes completely.** On a purpose-built corpus
  (one unindexed PST, 20,000 items across the four arrival-path folders, the 200-per-folder cap
  engaged) a single store's sweep produced a frame high-water of **10,734,599 bytes - 10.2 MB over
  758 items, ~13.5 KB per item**, and five such stores extrapolates to **~54 MB against the 64 MB
  limit**. So the 432 KB measured on the real profile was bounded by the TIMEOUT, not by the item
  caps, and `SweepBodyBytesBudget` (32 MiB) is load-bearing rather than insurance: it bites before
  the frame limit does, which is exactly its design intent. The residual PST case below is no longer
  the one this machine cannot produce - it has now been produced, on the test VM, and the bounds
  held. The residual risk narrows, and stays real: a fast LOCAL store
  absent from the index, where the window falls back to seven days, holding a lot of recent large
  mail - the archive/PST shape, and the one case this machine cannot produce, because the only
  unindexed store to hand is the test VM's and it is empty. **Bearing on the options:** (b) was not
  urgent on this evidence, and (c) was the one that would close the residual case outright - which
  is why (c) is what the maintainer chose (see above). The residual PST case is still the one this
  machine cannot produce, so the new bounds have never been exercised against real mail.
  Incidentally, the timeout path was observed working on a real profile for the first time: no
  hang, host replaced, honest degraded answer naming the reason.

  **The limit is reachable by ordinary use - derived from the caps 2026-08-18, not measured.** One
  `SweepFoldersNewerThan` answer is a single frame and `MailService` calls it with
  `includeBodies: true`, so a frame carries 4 arrival-path folders x `SweepPerFolderCap` (200)
  items per store, times every store in the profile. **The bodies were not capped at the COM
  layer** - `SnapshotBrief` took `item.Body` whole, and `BodyCharsDefault`/`BodyCharsCap` are
  applied in `MailService`, on the FAR side of the frame: they bound what the agent sees, not what
  crosses the pipe. **(They are capped there now - `SweepBodyCharsCap` / `SweepBodyBytesBudget`,
  2026-08-18 - so the arithmetic below is the worst case as it WAS.)** That puts 64 MB at ~80 KB average body on a one-store profile, ~27 KB on three
  stores, ~16 KB on five. An 80 KB body is an ordinary long quoted thread. **And the path there is
  the unindexed-store case**: the sweep window is normally minutes wide, so 200-per-folder never
  fills, EXCEPT when a store is missing from the index and the window falls back to seven days. So
  the more degraded the index, the larger the frame - and before the refusal landed, the frame
  bursting killed the subsystem that was compensating for the degraded index. It now refuses that
  one call instead, which is why (c) and (d) are still worth choosing between: the sweep on a
  degraded profile is exactly the caller with nothing left to narrow. This is where options (c)
  and (d) come from. Write-up in the session trace folder under Downloads
  (`tmp-aitrace/frame-size-analysis.md`).

  **The defect, as found (kept for the record; the refusal half is fixed).**
  `ComHostProtocol.EncodeFrame` refuses a payload over `MaxFrameBytes` (64 MB) by throwing
  `ComHostProtocolException` - a deliberate, specific, actionable failure. But it was thrown from
  `ComHostServer.WriteAsync`, which `ServeAsync` guarded with `catch (IOException)` only. So the
  exception left the serve loop, `Program.Main` printed it to stderr and the child exited with 1.
  The caller learned "the COM host went away", which is the one fact that says nothing about what
  to do next; the sentence naming the size and the limit reached only the child's stderr. This is
  the same species as the wrapper defect above - a good message that does not survive the process
  boundary - and it is the only other instance the audit found.

  **How likely is it?** Low - `McpServer/Docs/com-host.md` calls 64 MB "far above any real payload" and a
  `read` returns ~0.5 MB. The candidates are `SweepFoldersNewerThan(includeBodies: true)` and
  `ExhaustiveScan` over a large window. That "low" is still a derivation rather than an
  observation; the high-water mark now accumulating in `outlook_health` is what will replace it,
  and until an install has been read it says nothing on its own.

- [ ] **Verify the sweep's new body bounds against a live profile - T1 owns every decision above the COM call and none of the COM half.**
  `SweepBodyCharsCap` / `SweepBodyBytesBudget` landed 2026-08-18 and are pinned by T1
  `SweepBodyCapTests` (18 tests, mutation-checked). What T1 owns is the pure cut, the byte ceiling
  against the real serializer, the frame-half invariant, the payload fields, the coverage code, the
  advice split and the body-cache guarantee. What it cannot produce is a real `item.Body` big enough
  to cut, so **no swept body has ever actually been truncated on this machine.**

  - **What to confirm read-only, and it needs the shape the whole frame-size item is about:** a
    LOCAL store (PST/archive) absent from the index, so the sweep window falls back to
    `EmptyIndexSweepWindow`, holding enough recent large mail to move real body volume. Then an
    ordinary `search` naming that store should report `comHost.largestFrameBytes` (in
    `outlook_health`) well under `frameLimitBytes` with `framesRefusedTooLarge: 0`, and - only if
    something really was cut - `sweep.itemsBodyCapped` with its advice sentence.
  - **The two numbers worth measuring while that store is mounted**, because both are predictions
    rather than observations: the largest frame such a sweep actually produces (to see how much of
    the 32 MiB budget real mail uses), and the per-item cost of the body pass, which is one linear
    scan of at most `SweepBodyCharsCap` characters per item on top of the COM `.Body` read that
    produced it. The scan is expected to be noise next to the read; nothing has timed it.
  - **Unverifiable without a mailbox that has one:** whether any real mail body exceeds 500 000
    characters at all. If none ever does, the per-item cap is pure insurance and only the budget
    can ever bite - which would be the good outcome and should be recorded as such.

- [ ] **Measure the sweep and scan budgets against a known corpus - the tooling exists, the corpus does not yet.**
  The store shape the item above asks for cannot be borrowed from anywhere: every store on the real
  profile is indexed, so `EmptyIndexSweepWindow` never engages, and the Hyper-V VM's PST is the only
  unindexed store to hand and it is empty. So the corpus is built rather than found.
  `McpServer/OutlookAI.RemediationTools` gained five commands for it - `corpus-plan`, `corpus-probe`,
  `corpus-build`, `corpus-teardown`, `corpus-reindex` - and `Docs/corpus-measurement-plan.md` is the
  plan for what to run against the result and what each number would settle. T1
  `CorpusGeneratorTests` pins the size distribution, the date spread, the seeding, the store
  refusals, the teardown rule, the manifest format and the date-fidelity verdicts; only the COM
  calls are outside that tier, and they carry no decisions.

  - **Built once, 2026-08-19, and the measurement still has not been taken.** 40 000 items into
    the VM's PST in 12m27s at 50.9 items/sec, zero failures; resumability and determinism both
    demonstrated (a re-run skipped the 2 000 items an earlier timing run had made). Three faults
    came out of it, all now guarded in code rather than written down as cautions:
    - **Items were queued for delivery.** 5 532 landed in the target store's **Outbox** - inert
      on that VM only because its profile has no mail account. The store guard could not catch
      it: "local .pst" and "an account's delivery store" are not mutually exclusive.
      `CorpusSafety.EvaluateProfile` now refuses unless `Session.Accounts` is EMPTY, with no
      override. Stricter than "no account delivers here" because the object model cannot express
      the narrower rule - `SendUsingAccount` is per item, so any account may send a message that
      lives anywhere.

      **The population is identified, 2026-08-24.** The plan for that shape marks **5,532 items
      unread**, exactly - `corpus-plan` prints it now. So the queued items are precisely the ones
      the plan wanted left unread, and the read state was the only thing the builder treated
      differently: it set `MailItem.UnRead` and then wrote `PR_MESSAGE_FLAGS` WHOLESALE, as
      `MSGFLAG_READ` for a read item and as **0** for an unread one. Both are gone. One
      read-modify-write now carries the read state, clears `MSGFLAG_SUBMIT` on every item -
      that is the bit meaning "queued for delivery" - clears `MSGFLAG_UNSENT` only when the
      placement rung calls for it, and preserves every bit it does not own. **The mechanism
      inside Outlook is still unproved and only a build can prove it**, so `corpus-census`
      reports Outbox strays split by the plan's intended read state: a small build either
      confirms the identity or kills it.
    - **Every item was filed as a draft**, so the sweep saw 6 of 40 000 in 234-367 ms.
      `Items.Add` + `Save` produces an UNSENT item and Outlook files those in Drafts whatever
      folder they were added to; the sweep covers Inbox/Sent/Deleted/Junk and not Drafts. The
      root cause was a design fault of mine: the MSGFLAG_UNSENT write lived as a rung of the
      DATE ladder, so `--allow-undated` silently disabled placement too. `CorpusPlacement` is now
      its own probed ladder, and a rung passes only when the item's `Parent` is the target folder
      AND that folder's `GetTable` returns it.
    - **`corpus-reindex` and the post-teardown count looked only at Inbox/Sent/Junk/Deleted**, so
      with all 40 000 items in Drafts the recovery path would have reported ZERO and teardown
      would have claimed a clean store. Drafts (16) and Outbox (4) are in the scan set now. Same
      lesson as the Outbox omission `ComMailbox.SweepFolderIds` already records.
    - **Nothing in the tool noticed any of the three.** A person looking at Outlook found them.
      `corpus-census` closes that: a read-only scan compared against the plan - right count,
      right folders, one copy each, nothing stranded in Drafts or the Outbox - run by every
      build on itself and settable as its exit code.

  - **The placement probe's folder-table check is fixed, 2026-08-24.** Against an empty Inbox the
    move rungs verified; with ~22,000 items present the same rungs reported the item correctly
    parented and ABSENT from the folder's table, so the build refused a placement that works.
    The refusal was right given what it observed and the observation was wrong: the filter asked
    for every corpus subject in the folder and the walk stopped at its 2,000-row cap. It now
    filters on the probe's own reserved ordinal (`CorpusPlan.DaslSubjectFragment`, bracket-free
    so a `[` can never open a DASL character class), so it selects roughly one row; and reaching
    the cap is reported as INCONCLUSIVE rather than as "not there", with a refusal that blames
    the measurement instead of the store. The date probe's exclusion half had the same defect
    and is fixed the same way.

  - **A corpus expires silently, and now it does not (2026-08-24).** Anchored on a fixed instant,
    it stops filling the narrow measurement windows within weeks, and every test asking about
    them keeps PASSING because selecting nothing is a valid answer about an empty window.
    `corpus-verify` is pure - no Outlook, no store - derives the shift already applied from the
    manifest and refuses when any window under test has emptied; the live tier runs it
    fail-closed at fixture time from a new `corpus` settings block. `corpus-reanchor --to now`
    was the repair: an ABSOLUTE target, so it is idempotent and resumable, never creating, moving
    or removing an item, guarded by EntryID allowlist AND subject tags AND the expected ordinal.
    The manifest header's anchor is deliberately not rewritten - it is half the corpus's
    identity - so the shift is derived from the item lines and the re-anchor appends a
    replacement line per item. **Superseded 2026-08-25: the repair is a REBUILD - see below.**

  - [ ] **Move `T2/CorpusFreshnessTests.cs` to T1.** It is pure - no Outlook, no COM, no settings
    file, no `Category=Live` - and belongs beside `CorpusGeneratorTests`. It sits in T2 only
    because T1 was owned by a parallel worktree while it was written. A rename, nothing else.
  - **Still to do: tear down `CP-07-CORPUS-40K` and re-run.** The 40 000 drafts are still in the
    PST; the manifest is 40 002 lines and a copy is outside the guest. Teardown deletes by
    EntryID allowlist AND tag, and now reaches Drafts.
  - **The date verdict from that run proves nothing and must not be quoted.** The probe reported
    `readBack` correct with `daslIn=False`, which reads as "the date does not drive selection" -
    but the item was in Drafts while the probe queried the Inbox's table, so "not in this folder"
    explains it equally well. The two failures were indistinguishable in the output. The probe now
    settles placement first and builds its date probe with the placement that verified, so a
    re-run isolates the question.
  - **The one thing that must be settled first, before any of it is worth doing:** whether that PST
    accepts back-dated mail at all. `MailItem.SentOn` is read-only in the object model, and an item
    created straight into a folder is UNSENT, which some stores date themselves. `corpus-probe`
    settles it empirically - it writes one throwaway item per method, re-opens it by EntryID, reads
    `ReceivedTime` back, and then asks a DASL date restriction on either side of the instant whether
    it selects the item - and `corpus-build` refuses to build an undated corpus unless
    `--allow-undated` says so in as many words. **An undated corpus would make both windows select
    the same population while looking exactly like a good corpus**, which is why the refusal is the
    default rather than a warning.
  - **A product gap fell out of it, recorded as H3 in `Docs/completeness-gaps.md` and OPEN.** The
    sweep restricts on `(datereceived >= X) OR (date >= X)`, so mail carrying neither property is
    selected by NO window - absent from the freshness tier rather than mis-dated, and absent
    however wide the window is opened - while `sweep.foldersSwept` still counts the folder and
    `freshness` still says `live`. Real users hit this with imported, copied or restored mail.
    The filter is code and is not in doubt; the supporting observation is confounded by the
    placement fault above, and the row says so.
  - **The blocker for the 180 s proposal, found while writing the plan and not yet acted on:**
    `SearchBudgetMs` is `SearchIndexTimeoutSeconds * 1000 + SweepBudgetMs`, and T1
    `BudgetCompositionTests.SearchBudget_IsComposedFromItsPartsAndFitsTheOperationDeadline` asserts
    that sum fits inside `ComOperationBudgets.OperationDeadlineMs` (120 s). (That test is now
    `...FitsTheFreshnessDeadline` and names the sweep's own class - see the RESOLVED row below.) At 180 s the sum is
    195 s and that test fails before anything reaches a mailbox. Anything above roughly 105 s moves
    the operation deadline too, and with it the child work budget and `ExhaustiveTimeBudgetMs`. Decide
    the shape of that change before measuring, so the measurement is aimed at the right question.

- [ ] **Verify the three folder-walk reporting fixes against a live profile - the COM half none of them can reach from T1.**
  G2, G3 and G4 of `Docs/completeness-gaps.md` were closed on 2026-08-18 and are pinned by T1
  `FolderWalkReportingTests` (21 tests, driving the real `MailService` through a stand-in session
  and index client; mutation-checked - removing any one of the three fix lines fails 4 of them).
  What T1 owns is everything ABOVE the COM call, which is where the whole defect lived, since all
  three drops were decisions taken there. What it cannot produce is the COM failure itself:

  - **G2 needs a store whose `DisplayName` read throws.** Never observed on any machine here; it
    was found by reading the `catch` that swallowed it. Nothing in the repo knows what actually
    provokes it - a damaged profile entry, a data file that will not open, a store mid-removal are
    guesses. Worth trying: mount a PST, then rename or delete the file underneath Outlook while a
    session holds it. What to confirm read-only: the store appears in `list_folders` under
    `(unnamed store N)` with `nameUnreadable: true`, an unscoped `search` reports
    `sweep.storesUnnamed`, hits from it carry the label as their `store`, `outlook_health` lists it
    (via the same COM store list, so it should now reach `index.perStore[].inLocalIndex: false`),
    and `search(store: "(unnamed store N)")` is refused with the placeholder message rather than
    the typo one. **Also unverified in the real world:** whether a store that will not name itself
    will still answer `GetDefaultFolder` and `GetTable`, i.e. whether the sweep of it actually
    returns mail rather than four skips. The code handles both; only a live case can say which
    happens.
  - **G3 needs 10 000 folders in one profile, or a 65-level-deep tree.** Both are constructible in
    a test PST and neither exists here. The cheap partial check is a temporary build with the cap
    lowered (it is a `public const` on `MailService`, read at the call site) against the real
    profile: confirm `truncated: true`, `walkCapReached: true`, NO `nextOffset`, and the advice
    sentence - and that a store-by-store listing then returns the tree the capped call could not.
  - **G4 needs a delegate mailbox whose folder walk hits a bound**, i.e. G3's condition inside a
    shared mailbox. The dev profile has two delegate mailboxes indexing 11 and 23 folder paths, so
    the honest statement is that this flag has never fired on real data and cannot until the walk
    cap is reachable. The lowered-cap build covers it in the same pass: a delegate folder search
    should then report `scope.folderNamesTruncated: true`, `degraded: true`, and the INCOMPLETE
    SCOPE sentence.

  Read-only throughout - `list_folders`, `search`, `outlook_health`, `list_accounts`. No mailbox
  writes are needed for any of it.

- [ ] **Verify the sweep's sort-failure detection against a live profile - the one half of H2 that T1 cannot reach.**
  H2, G5, B2 and F3 of `Docs/completeness-gaps.md` were closed on 2026-08-18 and are pinned by T1
  `SearchCoverageClaimTests` (20 tests driving the real `MailService` through a stand-in session
  and index client, mutation-checked: five separate revert-one-line runs each fail exactly the
  tests that own that line). Three of the four are settled by that, because the whole defect lived
  above the COM call. H2 is not, and the gap is narrow and specific:

  - **`SweepFolder`'s new `out bool sortApplied` is set inside the `catch` around
    `Table.Sort`, and that catch has never been observed to fire on any machine here** - the
    defect was found by reading the swallowed exception, not by hitting it. So the value the whole
    row turns on is produced by a line no test executes. What IS pinned: the flag's journey across
    the process boundary (`ComHostProtocolTests.ComSweepResult_CarriesTheUnsortedCappedFoldersAcrossTheWire`),
    its per-store filtering in `ApplySweepCounters`, the code split, and both advice sentences.
  - **What would provoke it is unknown.** `Table.Sort` needs the property present as a column, and
    the code adds `urn:schemas:httpmail:datereceived` to `Columns` before sorting, so the ordinary
    path cannot fail. Guesses worth trying, none verified: a folder whose `DefaultItemType` is mail
    but whose contents are not, a folder on a store mid-reconnect, or a search folder. A cheaper
    substitute is a temporary build that forces `sortApplied = false` and confirms the payload
    end-to-end on the real profile.
  - **What to confirm read-only**, once a folder can be made to both refuse the sort and exceed
    `SweepPerFolderCap` (200 items in the freshness window - see the `EmptyIndexSweepWindow` row in
    `Docs/magic-numbers.md` for how wide the window has to get): `sweep.itemCappedFolders` names the
    folder, `sweep.itemCappedFoldersUnsorted` names it too, `sweep.coverageGaps` carries
    `item_cap_unsorted` and NOT `item_cap`, and the advice contains the word ARBITRARY and neither
    "newest-first" nor "OLDEST".

  Also unmeasured, and cheap to settle on the dev profile while the above is open: how often the
  G5 probes now run. The trigger widened from "the merged answer was empty" to "the index tier
  returned no rows", which costs two TOP-1 statements per folder-scoped search that the index did
  not answer. It was accepted on the standing rule that completeness outranks speed, so the number
  is worth having rather than worth acting on.

- [ ] **Verify the exhaustive scan's depth guard against a live profile - the half of F4 that T1 cannot reach.**
  F4 was closed on 2026-08-18 and is pinned by T1 `ScanDepthAndSweepScopeTests` end to end from
  `ComExhaustiveResult` to the payload, plus the process-boundary round trip in
  `ComHostProtocolTests`. What no test executes is the guard itself: `if (depth > FolderWalkDepthGuard)`
  needs a folder tree more than 64 levels deep, which no CI runner and no real mailbox has. The same
  shape as H2's `sortApplied`, and the same cheap substitute applies - a temporary build with the
  guard lowered to 2 or 3, then one read-only `exhaustive: true` search of a store with any nesting
  at all, confirming `exhaustive.depthLimitReached: true`, `depth_limit` in
  `exhaustive.coverageGaps`, `freshness: "partial"`, `degraded: true`, and an advice sentence
  naming the guard's value. Worth pairing with the lowered-cap build the G3/G4 item above already
  asks for; both are read-only and neither needs a mailbox write.

- [ ] **Add the second PST to the test VM - it is what makes the count tripwire mean anything
      there.** The tripwire exempts the hub store, so a machine whose only store IS the hub gives
      it nothing to watch: it will census, report zero failures, and be structurally incapable of
      reporting anything else. Recommended layout and the reason are in
      `Docs/live-tier-on-the-vm.md` section 2.3. Add it through Outlook's own UI (File > Account
      Settings > Data Files > Add) rather than a script - creating stores is not something the
      tested helpers do, and mailbox mutation from ad-hoc shell code is the thing that once
      destroyed real mail.
      **2026-09-24:** the store is named as an `.invalid` address now (`bystander@vm.invalid` in the
      example), because a small store is found in the index only through mail addressed to it, and it
      gets the generator's bystander population (next item). Adding it by script has been permitted
      on a guest since 2026-09-15, and attaching it to the second profile while still empty since
      2026-09-24 - `Docs/live-tier-on-the-vm.md` §2.6 draws both lines.

- [ ] **Make `corpus-teardown` drain the folders it created, as it drains the items: in a PST its
      `Folder.Delete()` MOVES them into Deleted Items, so every hub rebuild leaves two more empty
      ones there.** Measured on `OutlookAI-Unindexed`, 2026-10-03 (runbook §3b item 7, §4.1d):
      after the first `Reset-HubPopulation.ps1` run the hub's Deleted Items held
      `OutlookAI-Corpus-Folder-Projects` and `-Notices`, empty; after the second, also `... (2)` of
      each - Outlook renames on collision, so they pile up - and both teardowns printed
      `folders removed 2`. Harmless to the tests (the live tier's test folders are a different
      string), but the hub grows by two folders a run and the count says something that did not
      happen. Direction (a) of §3b item 7 was chosen overnight on the maintainer's behalf, for his
      review: first a read-only corpus-tool option that resolves a manifest's recorded folder
      EntryIDs and says where each sits now - `CP-12B-POPULATIONS-V2` holds two torn-down manifests
      and their four folders to run it on - then, if a folder keeps its EntryID across the move, a
      second `Delete()` from Deleted Items by EntryID AND prefix, and a count that says which
      happened.

- [ ] **UNTESTED: what an ADVISED EVENT SINK leaves inside Outlook when its COM host is killed.
      Two of the three nominated mechanisms were measured on 2026-09-15 and both came back
      negative; the third cannot be created from PowerShell, so it is the one gap left.**

      **Measured on the test guest, 2026-09-15.** A holder process was killed with
      `TerminateProcess` - the same call `ComHostSupervisor` makes - and then a fresh client was
      timed against the same Outlook instance:

      | Holder held before being killed | Fresh client afterwards |
      | --- | --- |
      | `Application` + `NameSpace` + `Folder` | `CreateObject` 0.37 s, `GetDefaultFolder` 0.05 s |
      | the same **plus a pin `Explorer`** (never displayed, never `Close()`d) | `CreateObject` 0.37 s, `GetDefaultFolder` 0.05 s, `explorers=1` |

      **Neither reproduced anything.** A killed holder's references do not block, delay or
      otherwise inconvenience the next client, with or without a pin.

      **THE NINE-MINUTE COM BLOCK THAT STARTED THIS WAS NOT REPRODUCED AND REMAINS UNEXPLAINED.**
      It was produced by throwaway probes, not by the product, and it **must not be cited as
      evidence about the kill path** in any direction - not as a reason to change the kill, and
      not as a reason to trust it. `Docs/autonomous-session-log.md` records "COM references
      orphaned by two earlier probes" as the only surviving explanation; the trials above are the
      direct test of exactly that, and it failed, so that paragraph is superseded and the cause is
      open. Nothing about `ComHostSupervisor`'s kill was ever implicated.

      **The pin leak is real, and it is a nuisance rather than a hang.** The orphaned `Explorer`
      **persists and is visible to the next client** (`explorers=1`), and `OUTLOOK.EXE` stayed up
      throughout. So a host killed while holding a pin leaves an invisible window behind: a slow
      accumulation over many kills, and an irritation to a human who then tries to exit Outlook
      and finds it will not go. It is not a wedge and it is not why anything blocked. Note also
      that **in the first trial `OUTLOOK.EXE` survived the holder's death with NO pin at all** -
      which weakens the pin's own justification, since the pin exists to keep Outlook alive.

      **DECIDED 2026-09-15: do not chase the sink now.** It is the only one of the three where a
      pointer is handed **into** Outlook rather than held by us, which is why it is worth testing
      at all - but it cannot be created from PowerShell, so testing it means deploying something
      that can, and the new guests are about to have exactly that deployed anyway. **Fold it into
      commissioning**: the real COM host lands on each guest as a build step, and
      `McpServer/OutlookAI.McpServer.Tests/T3/ComHostSupervisionLiveTests.cs` already drives the
      timeout / kill / respawn path there.

      **The experiment, concretely, so nobody has to redesign it:**

      1. **Start from a freshly restarted guest.** Not a convenience: the two trials above only
         mean anything because no earlier probe's orphaned state was left on the machine, and the
         nine-minute block itself is now believed to be orphaned-state contamination of some kind.
         Assert `OUTLOOK.EXE` is not running before anything binds COM.
      2. **Connect the real COM host**, so a genuine advised event sink is registered - the thing
         PowerShell cannot produce.
      3. **Kill it with `TerminateProcess`** - `Process.Kill`, the supervisor's own path, not a
         graceful shutdown. A graceful exit unadvises and proves nothing.
      4. **Time a fresh client's `GetDefaultFolder` under a watchdog.** The watchdog is the point:
         the failure mode being looked for is a call that never returns, so a bare stopwatch
         cannot record it. Bound it, and record the bound as the result when it expires.
      5. Record `explorers=` and whether `OUTLOOK.EXE` survived, the same two observations the
         trials above took, so the three mechanisms stay comparable.

- [ ] **Live-only, and unguarded by any non-live test: three decisions inside the count tripwire's
      verification.** Established by construction rather than by mutation, because they sit behind
      a COM census that no CI test can execute: `CollectionFinished`'s early return on `NotLast`;
      the keep-alive release and the `_verified` latch in `Verify(final)`; and the key-based
      intersection in the confirmation census. The KEY itself is pinned in
      `T1/StoreCountTripwireTests`; the code that uses it is not. The cheap substitute is the same
      one this file already records for `sortApplied`: a temporary build that forces the branch.
      **Narrowed 2026-08-24:** the intersection moved out of `Verify` into `T2\TripwireRetryLadder`
      and is now pinned by `T1/TripwireRetryLadderTests` against a fake census source, bounds and
      reporting included. What is left live-only is the LiveRetrySource wiring around it - the
      three censuses it takes, the 30 s waits, and the collections it names as implicated.

- [ ] **Watch for the count tripwire firing on a PST's Junk folder.** The census marks self-pruning
      folders by asking the store for its default Deleted Items, Junk and sync-issue folders. A PST
      may refuse `GetDefaultFolder` for Junk, in which case a generator-made "Junk Email" folder is
      an ordinary folder and a decrease in it FAILS rather than being noted. Nothing prunes it on a
      machine with no accounts, so this should stay theoretical - but if the tripwire ever fires on
      Junk, this is why, and the fix is to mark volatility by folder NAME as well as by default-folder
      identity.

- [ ] **Two live tests still degrade silently, and were left that way deliberately.**
      `LiveDraftTests.ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain` and its `LiveSendTests` twin
      loop over `expectedStoreDisplayNames`, so on a machine with fewer stores they sweep fewer and
      still pass. **This got sharper on 2026-08-24 and is now the more urgent half of this entry:**
      both used to be `LiveTier=ProfileBound` and therefore never selected on a test machine, and
      both are now in the VM bucket (`Requires=MailAccount,MultipleStores,Transport`), so a VM run
      WILL schedule them and they will pass having swept whatever the machine happens to have.
      Changing a sweep assertion without a live run to check it was not a trade worth making
      unsupervised, so this is left for the first real VM run. The other two of the four
      (`LiveStaleIndexRowTests`, `LiveManageSignatureTests.DefaultAssignment`) now refuse on a
      Production profile.

- [ ] **Residual gaps left by the 2026-08-23 tier-3 classification pass.** `Category!=Live`
      no longer reaches a mailbox: eleven T3 tests that called `outlook_health`,
      `list_accounts` or `search` moved into `ComHostSupervisionLiveTests`,
      `OutlookAvailabilityLiveTests` and `OutlookHealthLiveToolShapeTests`
      (`Category=Live` + `Requires=OutlookInstance`), `McpStdioClient` now refuses those three
      tools unless a test declares them, and
      `T1/LiveTierInventoryTests.EveryStdioTestReachingOutlook_DeclaresIt` reads the
      declaration back out of the IL. **Confirmed by measurement 2026-08-24:** the full
      `--filter "Category!=Live"` runs 2,226 cases green on a machine with Outlook up, moving
      Outlook's own CPU by 0.19 s in 107 s and lowering its handle count - so the standing local
      `FullyQualifiedName!~Tests.T3.` half of the filter is no longer needed and CI's plain
      `Category!=Live` is safe for the whole suite. What that pass found and did NOT fix:

  - [ ] **`list_accounts` starts Outlook, and nothing in the tool layer stops it.** The
        supervisor's liveness verdict for `NotRunning` is `MayStart`, which calls
        `BeginWarmUp` and connects to `Outlook.Application` - so on a machine with Outlook
        installed but closed, a bare `list_accounts` launches it. `outlook_health` guards its
        own probe with `if (outlookRunning)`; `list_accounts`, `list_folders`, `read`,
        `search`'s sweep and every draft path do not. Correct for a shipped tool
        (S7/D17 permits the cold start); a hazard for a test tier, and it was reachable from
        the default run until this change. Decide whether the live tier should force
        `allowStartingOutlook: false` for the T3 stdio tests, which would need a server-side
        switch it does not have.
  - [ ] **Four T3 tests pass while asserting almost nothing on a machine with no Outlook**,
        each by an early `return` that is documented where it sits. They are not new and none
        is wrong, but together they are why "the CI tier is green" said less than it looked:
        `OutlookAvailabilityLiveTests.SearchAlwaysAnswers_AndSaysWhetherItIsComplete` (returns
        the moment `search` reports an error, which on an indexless machine is every run),
        `...ATransientOutlookState_AnswersFastAndCarriesRetryGuidance` (returns when Outlook
        is healthy, keeping only the timing assertion),
        `ComHostSupervisionLiveTests.NoComHostSurvivesTheServer` (returns when no COM host was
        spawned, keeping only the stdin-close assertion) and
        `OutlookHealthLiveToolShapeTests.OutlookHealth_CarriesTheFreshnessBlock_WithOrWithoutAnIndex`
        (skips the advice assertion when the index provider is unavailable). All four are now
        `Category=Live`, so the question is what the VM run should assert INSTEAD of returning.
        **The first three were decided and converted on 2026-10-03 (Q101)**: each return goes
        through `T2/LivePopulationCoverage` - a refusal on `Production`, a `PROVED NOTHING:` line on
        `Portable` - and the search test now fails on any error but an unreachable index, which it
        recognises by `outlook_health`'s own `index.provider` verdict. **The fourth is still as
        described**, and is the one left in this entry. Worth knowing before the next Production run:
        the transient-state test can only exercise its check while Outlook is starting, hung or
        unavailable, so on a workstation with a healthy Outlook it now FAILS by design; the server's
        `OUTLOOKAI_COMHOST_LIVENESS` override could make it force such a state instead, if that is
        ever preferred to the Q57 answer.
  - [ ] **The pin reads tool NAMES, not arguments.** `search`, `read`, `thread`, the draft
        tools, `move_mail` and the show-me tools all have a refusal that fires before any COM
        work, which is what the protocol-only half of T3 is built on - so they cannot be
        blanket-guarded, and a future test that calls one with arguments that DO reach Outlook
        would not be caught. A bounded exhaustive `search` is the realistic case.

- [ ] **`corpus-reindex` → `corpus-teardown` silently loses rows whenever two scan rows share an
      ordinal.** Found 2026-09-16 while establishing what the census's duplicate-ordinal fault is
      actually for. `CorpusCommands.RunReindex` writes one manifest line per scan row, taking
      `row.Ordinal` straight through
      (`McpServer/OutlookAI.RemediationTools/CorpusCommands.cs:1073-1077` since the population work of
      2026-09-24, which also writes the created folders first; the item loop is unchanged), and `CorpusManifest.Parse`
      keys items by ordinal - `manifest._items[item.Ordinal] = item;`
      (`McpServer/OutlookAI.RemediationTools/CorpusManifest.cs:226`), and `Add` does the same at
      `:248`. So the second line for an ordinal **overwrites** the first, in memory, with no
      unparseable-line record and no count anywhere that a reader could compare against the
      "Wrote N entries" the command prints.
      **What it costs.** Reindex is the RECOVERY path for a lost manifest, and a duplicated ordinal
      is one of the two things it exists to recover from: a build interrupted between the COM create
      and the manifest flush re-creates that ordinal on the next run, leaving an orphan copy that
      only a scan can find. Reindex would find both copies, write both lines, and then hand
      `corpus-teardown` a manifest naming only one of them - dropping the orphan from the very
      recovery manifest that exists to catch it. Teardown's phase 2 re-scan would in practice still
      reach it, so the item is not permanently stranded; what is lost is the EVIDENCE, at the one
      moment somebody is looking for it. Bounded by how rarely a build is interrupted, which is why
      this is recorded rather than fixed.
      **What the fix would be.** Key the manifest by EntryID rather than by ordinal (the EntryID is
      what every delete is addressed with; the ordinal is only ever a label), or - smaller - have
      `Parse` and `Add` refuse to overwrite and route the loser to `UnparseableLines`, and have
      `RunReindex` print the duplicate count beside the total it already prints.

- [ ] **Retire v3 planning ignores** — once the local v3 planning files (`v3.MD`, `Docs/v3-probes/`) are no longer needed:
  - [ ] remove the "v3 planning documents" section at the bottom of `.gitignore`
  - [ ] delete the local plan-doc backup folder (location documented in v3.MD §0.8 D16 on the machine that holds it)
  - [ ] delete this TODO entry (and this file if empty)

- [ ] **Remove the 13 stale Agent-tool worktrees and their `worktree-agent-*` branches.**
  All 13 under .claude/worktrees are clean, unlocked and 0 commits ahead of master (0 to 107
  behind), so every commit they hold is on master already. Each still has a checked-out CLAUDE.md,
  which a session started or resumed there would load. Confirm no agent is running in one (the last
  writes were 2026-09-24 to 2026-09-27), then for each: `git worktree remove .claude/worktrees/agent-<id>`
  and `git branch -d worktree-agent-<id>`; `-d` refuses a branch that is not merged, which is the check.
  Added in 2.1.277: https://code.claude.com/docs/en/changelog#2-1-277
  Extended in 2.1.281: https://code.claude.com/docs/en/changelog#2-1-281
  Remaining differences: https://github.com/anthropics/claude-code/tree/main/mods/agents-md#where-it-still-differs-from-claudemd

- [ ] **Clean up old `.work` folders (`.work/g2-cp11b`, `.work/q81-addin-registration`, `.work/testbed-livetier-payload`, `.work/worktree-archive`).**
  Their copies and clones carry CLAUDE.md/AGENTS.md files. `.claude/settings.json` keeps them out of
  Claude's context (`claudeMdExcludes`), but delete what is no longer needed.
  Added in 2.1.277: https://code.claude.com/docs/en/changelog#2-1-277
  Extended in 2.1.281: https://code.claude.com/docs/en/changelog#2-1-281
  Remaining differences: https://github.com/anthropics/claude-code/tree/main/mods/agents-md#where-it-still-differs-from-claudemd

- [ ] **Move OutlookAI from MIT to the FSL-based licence BrowserAI carries.**
  Decided by the maintainer on 2026-10-01, in his words: *"Move both OutlookAI and the new library
  to the FSL based license BrowserAI has."* The new library is the shared MCP registration library
  that is being extracted from BrowserAI. Nothing in this repository has been changed for it: this
  entry is the whole of the change so far, added from the BrowserAI session on his instruction.
  **The licence to copy** is `LICENSE` in the BrowserAI repository: the Functional Source License
  1.1 (MIT Future License), modified to a five-year term. It is a bespoke variant, and its own text
  forbids the plain `FSL-1.1-MIT` identifier. BrowserAI names it
  `LicenseRef-BrowserAI-FSL-1.1-MIT-5yr` in its SPDX headers, so OutlookAI needs a name of its own
  in the same form.
  **What the move touches, to be checked when it is done:** `LICENSE`; the licence line in
  `README.md`; any licence field in the project files and in `Installer.iss`; source-file headers,
  if this repository wants them; and `CHANGELOG.md` under Unreleased, because the licence a user
  receives changes.
  **One thing to settle first.** Every release published so far went out under MIT. Decide whether
  the change applies from the next release or from a stated date, and say so in the release notes.
