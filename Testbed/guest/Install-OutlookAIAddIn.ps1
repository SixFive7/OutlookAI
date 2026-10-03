#Requires -Version 5.1
<#
    ============================================================================================
    RUN ON OutlookAI-Unindexed 2026-10-03, FROM CP-08 (NO ADD-IN): THE TWO PHASES WORK AS DESIGNED.
    THEIR FIRST RUN FOUND A PRODUCT DEFECT: IN A NOT ELEVATED OUTLOOK THE ADD-IN'S TUNING RECONCILE
    NEVER FINISHES, SO A FRESH GUEST ENDS BROKEN, NOT ADDIN-READY. Docs/live-tier-on-the-vm.md
    section 2.3 has the record; the split itself is described under the next heading.
    ============================================================================================

    The payload was built on the host from 883ec5f by Testbed/host/Publish-AddInPayload.ps1 (guard
    3 UNCHANGED over 305 host lines). Every phase ran through Register-InteractiveTask.ps1 in
    session 1, and a read-only poller in session 0 sampled OUTLOOK.EXE every 0.5 s throughout -
    its token, command line, parent and modules.

      -Verify          NOT-INSTALLED, exit 3. v4R absent, v4 10.0.60910.
      -Execute alone   REFUSED from the default task, exit 1, naming both phases. Nothing installed.
      Install          INSTALLED-NEVER-RAN, exit 2, in 66 s: the runtime installed (v4R 10.0.60917,
                       35 s), the installer 14 s, the trust entry and the install record written.
                       No OUTLOOK.EXE at any sample, none after. -Verify: INSTALLED-NEVER-RAN, 2.
      wrong tokens     FirstRun from the default (elevated) task REFUSED, exit 1, and Install from a
                       Limited task REFUSED, exit 1 - each before it started or wrote anything.
      FirstRun         BROKEN, exit 1. Outlook started over COM in 3.5 s; "OUTLOOK.EXE pid 9064:
                       token NOT elevated", and the poller agreed (TokenElevationType Limited, Medium
                       integrity, OUTLOOK.EXE -Embedding under DcomLaunch); the add-in loaded,
                       connected, answering GetRestartNeeded(). It wrote Tuning\Initialized and
                       Enabled - and no LastReconcileUtc in 240 s. THE DEFECT: OutlookTuningService
                       .Reconcile writes D25's five Cached Mode values under HKCU\Software\Policies,
                       where the user holds ReadKey only; the first write throws, the reconcile
                       swallows it, and nothing after it runs (Tuning\Applied: 4 of 13). The single
                       elevated -Execute this replaced hid it: that Outlook COULD write there, and
                       every later start found the values in sync.
      -Verify          then said INSTALLED-NEVER-RAN - wrong: the add-in had run. Fixed; see below.
      control          those five values written from session 0, elevated, as a GPO would, and the
                       SAME FirstRun: ADDIN-READY, exit 0, in 8.4 s - LastReconcileUtc 2.1 s in, 13 of
                       13 applied, token NOT elevated. -Verify: ADDIN-READY, exit 0.
      Install again    over that valid state: INSTALLED-NEVER-RAN, exit 2 - runtime and trust entry
                       kept, record rewritten; -Verify INSTALLED-NEVER-RAN; FirstRun ADDIN-READY again.
      after it         a graceful restart, then Set-OutlookIndexingDisabled.ps1 -Verify: UNINDEXED,
                       two readings 10 minutes apart, no Outlook row - after three NOT elevated
                       Outlook starts, under CP-08's policy-only exclusion.

    FIXED FROM THAT RUN, THE SAME DAY: -Verify judged "has it run since the install" by the tuning
    state alone, so an add-in that started and never finished its tuning reconcile read as
    INSTALLED-NEVER-RAN. It now also reads the add-in's second startup marker - Mcp\LastReconcileUtc,
    written at every start, failed or not - and calls that case BROKEN, with when the add-in started
    and how far its walk got (Tuning\Applied against Tuning\Desired, printed as "tuning walk").
    FirstRun watches the same marker during its wait, so when no tuning state comes it says which
    failure it was - and its timing is no longer read after the 240 s wait. On the guest the fixed
    script then said BROKEN with the reason from -Verify, and from a FirstRun "tuning state ...
    written NEVER, in 240 s; registration reconcile ... written after 2.4 s". -SelfTest: 151
    assertions, 0 failures there (the contract section needs the repository); and 187, 0 failures,
    on the build VM under Windows PowerShell 5.1 with every other script's -SelfTest, 21 of 21
    (Testbed/host/Invoke-TestsOnBuildVm.ps1 9bfc135 -SkipSuite, run 20261003-184239-9bfc1353c5eb).
    The four Tools/Checks guards pass under both shells.

    NOT SETTLED BY IT: the defect - a product decision, open in TODO.md; the indexed guest, whose
    index an unelevated first run feeds; and -Verify -WithOutlook, whose Outlook had closed by itself
    before the attach was tried. No checkpoint was kept: the guest went back to CP-13B-LIVE-GREEN.

    ============================================================================================
    SPLIT INTO TWO PHASES 2026-10-03 (Q100, option 3).
    ============================================================================================

    -Execute used to install AND start Outlook from one elevated task, so its first-run Outlook was
    elevated - and an elevated Outlook never feeds Windows Search (Docs/live-tier-on-the-vm.md
    section 8 item 22). On OutlookAI-Indexed that broke the rule that every Outlook there starts
    unelevated. Decided by the maintainer 2026-10-03 (Q100, option 3): two steps.

      -Phase Install -Execute    ELEVATED - Register-InteractiveTask.ps1 at its default RunLevel
                                 Highest. Payload checks, the VSTO runtime, the product's installer,
                                 trust, VSTO_LOGALERTS, and a record of what it installed and when.
                                 NEVER STARTS OUTLOOK. Ends INSTALLED-NEVER-RAN, exit 2 - never
                                 ADDIN-READY, even over the state an earlier run left.
      -Phase FirstRun -Execute   NOT ELEVATED - Register-InteractiveTask.ps1 -RunLevel Limited - and it
                                 refuses an elevated token. Starts Outlook once over COM, headless,
                                 reads the started OUTLOOK.EXE's token and fails an elevated one,
                                 waits for the tuning state, and verifies everything -Verify does,
                                 the modal-dialog check included. The only route to ADDIN-READY,
                                 which means what it meant before: written by THIS start.
      -Execute with no -Phase    REFUSED, naming both phases, so an old command line cannot do half
                                 the job and look finished.

    -Verify now also reads the install record: a tuning state older than the last -Phase Install
    is INSTALLED-NEVER-RAN, not ADDIN-READY. With no record - an install made before the split - it
    reads exactly as before. The record is the one new write: install-addin-record.json beside the
    log.

    THE PROOF BEFORE THE GUEST RUN ABOVE, ALL ON THE HOST: -SelfTest, 168 assertions, 0 failures,
    under Windows PowerShell 5.1 and PowerShell 7 (115 before the split). Seven of them read this file's own syntax tree:
    the install phase reaches no Outlook start, COM attach or first-run job; the first-run phase,
    and the job it starts, reach no installer and no registry, environment or file write; each
    phase's first command refuses the wrong token; and two controls show the first two are not
    vacuous. Eleven rules were broken on purpose in scratch copies - an Outlook start reached from
    the install, directly and through a shared helper; a write in the first run and in its job;
    the token check moved down; an elevated first run let through; ADDIN-READY from the install;
    -Execute alone accepted; the install record ignored; the Limited run level dropped; an
    untrusted add-in started - and -SelfTest failed on each. Plus the four guards (Tools/Checks today),
    under both shells. That an unelevated token can read every HKLM key the first run reads (the
    crawl-scope rules, the Search policy, VSTO Runtime Setup) was measured on the host's Windows 11
    - and on the guest since: every FirstRun above read them at Limited, and an unreadable key
    throws there. Every run recorded below is of the single -Execute this replaced.

    ============================================================================================
    RUN ON OutlookAI-Unindexed 2026-09-24 (FROM CP-08): ADDIN-READY, TWICE.
    ============================================================================================

    The payload was built on the host by Testbed/host/Publish-AddInPayload.ps1 from commit fe65ced
    (376 host lines identical before and after the build), staged with the pinned
    vstor_redist.exe, and run from C:\OutlookAI-Q5 in session 1 through Register-InteractiveTask.ps1,
    with the tier profile the default, its POP3 password stored and the Q80 programmatic-access
    policy in place:

      -SelfTest   on the guest it DIED at the contract section - "Cannot bind argument to parameter
                  'Path' because it is an empty string": staged at C:\OutlookAI-Q5, two levels up is
                  '' and Join-Path refused it. Fixed (it now says SKIP, as it always meant to);
                  85 assertions, 0 failures there, the 30 contract checks skipped because the
                  suite's source is not staged yet at this step; 115 on the host.
      -Verify     before: NOT-INSTALLED, exit 3. v4R ABSENT, v4 10.0.60910 - Office LTSC 2024 does
                  NOT bring the VSTO runtime key Installer.iss looks for.
      -Execute    ADDIN-READY, exit 0. The runtime installed in 20 s (v4R 10.0.60917); the product's
                  installer in 6 s; the inclusion entry written; Outlook started over COM in 3.2 s,
                  and the add-in wrote its tuning state 3.5 s after the start (Initialized DWORD 1,
                  Enabled DWORD 1, LastReconcileUtc REG_SZ); COMAddIns Connect True, and the add-in
                  answered GetRestartNeeded() (True). No trust prompt, no other window. The index
                  exclusion UNCHANGED. The headless Outlook then closed by itself within ~45 s of
                  the last reference being released.
      restart     graceful (no Outlook running; shutdown /r /t 0), then -Execute AGAIN: ADDIN-READY;
                  "already registered: v4R 10.0.60917 - not reinstalling"; "kept the existing entry
                  36156351-... - same URL, same key"; tuning state 3 s after the start;
                  GetRestartNeeded() False this time.
      after it    Set-OutlookIndexingDisabled.ps1 -Verify: UNINDEXED (catalog reachable, no Outlook
                  row on either reading). Its first attempt said NO-INDEXER only because its first
                  reading fell 74 s after a boot, before Windows Search's delayed start.

    RUN ON OutlookAI-Indexed TOO, TWICE: 2026-09-24 (payload 98e050e, CP-14-ADDIN-READY) and
    2026-09-27 on its Q87 rebuild (payload af56efc, CP-14C-ADDIN-READY). Both times NOT-INSTALLED,
    then ADDIN-READY twice across a graceful restart; v4R absent, then 10.0.60917; the trust entry
    kept the second time; the index exclusion state UNCHANGED. Tuning state 8.3 s and 3.6 s after the
    start the first time, 3.5 s both times on the rebuild. The headless Outlook the script leaves
    running had closed by itself within 11 s of the first -Execute both times (the restart found it
    gone), and within 20 s of the second on the rebuild (the index -Verify found it gone).
    Set-OutlookIndexingDisabled.ps1 -Verify after the second -Execute: INDEXED, 20,048 and then
    20,059 rows, the same on both readings.
      On that guest every other Outlook start is unelevated (an elevated Outlook does not feed
    Windows Search). This one is not: -Execute runs from a RunLevel Highest task - the VSTO runtime
    needs it - and starts its first-run Outlook over COM from there. The index was intact after
    it, as above; the elevation of that Outlook itself was not read.

    NOT EXERCISED ON A GUEST: -Verify -WithOutlook (the COM-started Outlook had already closed by
    itself). Worth knowing before relying on it: it attaches with GetActiveObject, and an Outlook
    started BY COM (-Embedding) is not in the Running Object Table - measured the same day on CP-05's
    Outlook, MK_E_UNAVAILABLE - so against that kind of Outlook it would report an error, not a
    verdict.

    ============================================================================================
    WRITTEN 2026-09-24. WHAT HAD RUN BEFORE THE GUEST RUN WAS ON THE HOST, AND IT WAS THIS:
    ============================================================================================

      * -SelfTest: 115 assertions, 0 failures, under Windows PowerShell 5.1 and PowerShell 7 - 30
        of them read the source files this script mirrors. Six of its rules were broken on purpose
        in a scratch copy (the DWORD-only bool, the URL slashes, the key comparison, freshness, a
        contract string, the never-ran verdict) and each was caught.
      * Its pure functions, loaded without the main body, against the REAL payload
        Testbed/host/Publish-AddInPayload.ps1 built from fb19ccf: the manifest is well-formed; the
        key read out of the built OutlookAI.vsto equals the one in OutlookAI.cer, the one in the
        application manifest and the one the payload manifest names; and the zip-comment reader
        returned fb19ccf from a real `git archive` zip.

    Every mechanism below was checked against something real before it was written down - the
    VSTO runtime's own IL, the installer's own source, trust entries the runtime itself wrote on
    the maintainer's workstation - and each is labelled with how it is known. None of that is a
    guest run. The things only a guest can settle are listed at the end of -SelfTest's output -
    and the guest run above settled all but the last of them.

.SYNOPSIS
    Puts the OutlookAI add-in - built from a named commit by Testbed/host/Publish-AddInPayload.ps1 -
    onto a testbed guest and trusts it without a prompt (-Phase Install, elevated); then starts
    Outlook once, NOT elevated, so the add-in writes its tuning state, and PROVES the exact registry
    state the live tests read (-Phase FirstRun).

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    Both phases run in the interactive session, through Testbed/guest/Register-InteractiveTask.ps1 -
    never over PowerShell Direct, which lands in session 0 where Outlook cannot finish starting -
    and AT DIFFERENT RUN LEVELS:

        .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase Install -Execute"
        .\Register-InteractiveTask.ps1 -RunLevel Limited -Script "& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase FirstRun -Execute"

    Install needs elevation for the machine-wide VSTO runtime; FirstRun must not have it, because
    the Outlook it starts inherits its token and an elevated Outlook never feeds Windows Search. Each
    phase refuses the other's token. On the unindexed guest run FirstRun only once the index
    exclusion is in place (Testbed/README.md section 1, steps 5b, 7b and 7c): a NON-elevated Outlook
    adds itself to the index within a minute of starting.

    -Verify and -SelfTest run anywhere. NEVER on the maintainer's workstation: the guard refuses.

    WHY THIS EXISTS. Two live tests read state only the add-in writes, the first time it runs:

      T3/Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning
          declares Requires=AddInRegistry; asserts outlook_health's tuning.managed is true.
      T2/LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail
          declares only SearchIndex and MultipleStores - it UNDER-DECLARES AddInRegistry - and
          asserts Tuning.Managed, Tuning.Enabled and a non-null Tuning.LastReconcileUtc.

    (T2/LiveUiSearchBackendTests.FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates also
    declares AddInRegistry, but writes the value it tests itself; it does not need the add-in.)

    The script-built guests have no add-in, so both of those FAIL there - not skip. And BOTH GUESTS
    need it: the Phase-7 test declares nothing that keeps it off the unindexed guest, and it asserts
    index.wSearchStartMode == "automatic", which Set-OutlookIndexingDisabled.ps1 deliberately keeps
    true there.

    WHAT THE TESTS READ, EXACTLY. McpServer/OutlookAI.Core/Services/HealthReporting.cs,
    ReadTuningState, over HKCU\Software\OutlookAI\Tuning, value names from
    Services/AddInServerContract.cs:

      Initialized       must be a REG_DWORD, nonzero -> tuning.managed = true
      Enabled           must be a REG_DWORD, nonzero -> tuning.enabled = true
      LastReconcileUtc  must be a REG_SZ             -> tuning.lastReconcileUtc not null

    TYPE MATTERS, not just value: HealthReporting's AsBool accepts a boxed int and nothing else, so
    a REG_QWORD 1 or a REG_SZ "1" reads as not-managed. -Verify mirrors that rule exactly, and
    -SelfTest checks the mirror against the source it mirrors.

    WHAT THE TWO PHASES DO, IN ORDER, and every step is idempotent. -Phase Install is steps 1 to
    6 and 6a; -Phase FirstRun is steps 1, 2, 6b, 7 and 8:

      1. REFUSES unless: this is a guest (two-axis guard), 64-bit, in an INTERACTIVE session, and
         OUTLOOK.EXE is not running - Install would replace files under a running add-in, and
         FirstRun's proof needs the start to be this script's. And the RUN LEVEL: Install refuses a
         token that is not elevated, FirstRun one that is.
      2. CHECKS THE PAYLOAD. addin-payload.json from the host build; the installer's SHA-256 must
         match it before anything runs. (FirstRun reads the same manifest to tie the installed
         build to its commit, as the verify always did.)
      3. THE VSTO RUNTIME, from STAGED media - never a download - pinned by the SHA-256 and length
         every release pins (until 2026-10-03 the release workflow also compared it against
         Microsoft's own download on every release, which is why this hash, unlike the SDK's, has
         a default). Skipped when
         `VSTO Runtime Setup\v4R` already reports this version or newer.
         WHY THIS SCRIPT INSTALLS IT AND THE INSTALLER DOES NOT: Installer.iss runs its
         prerequisite step only `if not WizardSilent`, and a silent install is the only kind an
         unattended guest can run. [READ in Installer.iss.]
      4. THE PRODUCT'S OWN INSTALLER, silently: /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
         /NOCLOSEAPPLICATIONS. It is per-user (PrivilegesRequired=lowest) into
         %LOCALAPPDATA%\OutlookAI\Setup, registers the add-in under
         HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI as `file:///{app}\OutlookAI.vsto|vstolocal`,
         exempts it from Outlook's slow-add-in disabling, and imports its signing certificate into
         the user's TrustedPublisher store - exactly what it does for a user.
      5. TRUST, WRITTEN RATHER THAN CLICKED. See TRUST below.
      6. VSTO_LOGALERTS=1 for this user, so a load failure writes <app>\OutlookAI.vsto.log instead
         of vanishing. [MS-DOC: "Debug Office projects".] Errors are NOT shown in a dialog - that
         is VSTO_SUPPRESSDISPLAYALERTS=0, which this never sets.
     6a. RECORDS THE INSTALL - commit, version, the moment the installer finished, the install
         directory - in install-addin-record.json beside the log (the old record is removed before
         the installers run, so a failed install leaves none). Then the verify, which can only say
         INSTALLED-NEVER-RAN or BROKEN: this phase never starts Outlook, so it never says ADDIN-READY.
     6b. FIRSTRUN'S PREFLIGHT, before any Outlook start: the add-in must be installed, registered,
         trusted, present on disk as the payload's build, with the runtime in place and not in
         Outlook's hard-disable list. Any of those wrong, and a start would put the trust prompt on
         the console or prove nothing - so it stops, BROKEN or NOT-INSTALLED, Outlook untouched.
      7. STARTS OUTLOOK ONCE, over COM, headless, in a CHILD job under a deadline - so a hang is
         reported rather than inherited, the shape the measured cold-start probe in
         Testbed/MEDIA.md used. The child inherits the phase's Limited token, and so does the Outlook
         COM starts for it - which the job proves by reading the started OUTLOOK.EXE's token while it
         still holds Outlook, the way Start-OutlookUnelevated.ps1 does; an elevated one is BROKEN.
         It waits for LastReconcileUtc to be written AFTER the start, asks
         Outlook whether the add-in is connected, and calls into the add-in itself
         (COMAddIn.Object.GetRestartNeeded, AddInAutomation.cs) - which only a loaded add-in can
         answer. It releases every reference and NEVER quits or kills Outlook (mailbox-safety rule
         7). Outlook may stay up headless or close by itself once the last reference goes; both
         are graceful and the script says which happened.
      8. VERIFIES, as -Verify does, plus FRESHNESS: the state must have been written by THIS start;
         and no window that appeared during the run may still be on screen.

    TRUST, AND WHY IT IS WRITTEN. A VSTO add-in loads silently only if its manifest's signer is a
    trusted publisher CHAINING TO A TRUSTED ROOT, or an inclusion-list entry vouches for it.
    Otherwise the ClickOnce trust prompt appears - and on an unattended guest a prompt is a hang,
    which is what the hand-built guest's CP-05-ADDIN-TRUSTED checkpoint recorded somebody clicking
    through. The installer's TrustedPublisher import alone does not retire that prompt for a
    self-signed certificate (Installer.iss says so, and why it will not touch the Root store). So
    this script writes the entry the prompt itself would have written:

      HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid>
          Url        REG_SZ  file:///C:/Users/.../OutlookAI/Setup/OutlookAI.vsto
          PublicKey  REG_SZ  <RSAKeyValue><Modulus>...</Modulus><Exponent>...</Exponent></RSAKeyValue>

      [MS-DOC] "If the end user grants trust to the solution, an inclusion list entry is created
               that contains a URL and a public key" - Grant trust to Office solutions.
      [MS-SUPPORT] the key path and both value names - Microsoft Japan Office support blog on
               UserInclusionList, which also records that the public API for it is not usable
               from .NET 4 (confirmed: AddInSecurityEntry is internal in the v10 runtime).
      [READ] the v10 runtime's own IL: it stores Url as Uri.ToString(), finds entries by comparing
               Url AS A URI, and compares keys by FromXmlString + ExportCspBlob - so the key must be
               the manifest's signing key, and the XML shape is free.
      [MEASURED] the maintainer's workstation carries an entry of exactly this shape for this very
               installer's install path: forward slashes, `file:///C:/...`, the key as a bare
               <RSAKeyValue>. Who wrote it is not recorded; the trust prompt is the only writer
               this repository knows of for an installed path.

    The key is read out of the INSTALLED deployment manifest - the thing the runtime compares - and
    must agree with the installed OutlookAI.cer and with the payload manifest. Only entries for this
    one URL are ever replaced.

    WHAT IT NEVER DOES. It never kills or quits Outlook, never creates, moves, edits or deletes a
    mail item, never touches a store, a profile or MAPI, and never writes the add-in's own
    Software\OutlookAI keys - the tests read what the ADD-IN wrote, or nothing. It does not touch
    the Windows Search policy or the crawl scope; it READS both before and after, and reports any
    change (see INTERACTION). -Phase Install never starts Outlook; -Phase FirstRun installs nothing
    and writes nothing but its log - its one effect is the Outlook start it is there to prove.

    INTERACTION WITH THE INDEX EXCLUSION AND THE CORPORA. The add-in's tuning service
    (Services/OutlookTuningService.cs) writes only under HKCU: Outlook's Search key (four search-box
    preferences), the Cached Mode user and POLICY keys (Exchange-only sync settings), and the PST
    key (a larger PST/OST size cap). Set-OutlookIndexingDisabled.ps1 writes only HKLM: the Windows
    Search PreventIndexingOutlook policy and the crawl-scope rule. The two sets are DISJOINT, none
    of the four search values decides whether Outlook's stores are indexed, and nothing the add-in
    does touches an item or a store. [READ, both sources.] Each phase snapshots the exclusion state
    before and after its work - FirstRun's around its Outlook start, the one that matters - and says
    if anything moved; if it did, re-run Set-OutlookIndexingDisabled.ps1 -Verify before trusting
    the guest as unindexed. The Outlook FirstRun starts is NOT elevated, so unlike the elevated one
    this script used to start it is an Outlook the indexer serves: on the unindexed guest it must
    start only once the exclusion is in place.

    FOUR VERDICTS, and only one exits 0:

      ADDIN-READY           installed, registered (LoadBehavior 3), trusted, runtime present, and
                            the state the tests read is valid - and, under -Phase FirstRun, written
                            by THIS start. Only FirstRun and -Verify can say it. exit 0.
      INSTALLED-NEVER-RAN   everything in place, but the installed build has not written its state
                            yet: no state at all, or only one older than the last -Phase Install -
                            and no sign the build has started since either (Mcp\LastReconcileUtc,
                            which it writes at every start; one that started and never finished its
                            tuning reconcile is BROKEN). What -Phase Install ends with. A real state,
                            not a fault: run -Phase FirstRun. exit 2.
      NOT-INSTALLED         no add-in here. exit 3.
      BROKEN                something that should hold does not; every reason is printed. exit 1.

    THE GUARD, TWO AXES, same rule and shape as Testbed/guest/Install-DotnetSdk.ps1: the logged-on
    user must be one of -ExpectedUser AND the computer name must start with
    -ExpectedComputerNamePrefix. It runs first in every path except -SelfTest.

.PARAMETER PayloadRoot
    Where AddIn.zip was expanded: the installer and addin-payload.json.

.PARAMETER VstoRuntimePath
    The staged VSTO runtime redistributable. Never downloaded.

.PARAMETER VstoRuntimeSha256
    Its SHA-256. Defaulted to the release's pin - see step 3.

.PARAMETER SuiteSourceRoot
    Where Testbed/host/Publish-LiveTierPayload.ps1's source was expanded. Used to compare the add-in's
    contract files against the suite's; absent is a note, not a failure.

.PARAMETER SuiteSourceZip
    That payload's Source.zip, whose zip comment `git archive` sets to the commit it was built from.

.PARAMETER OfficeVersion
    The Office major whose Outlook hive to read. Detected the way Services/OfficeVersions.cs does
    when omitted.

.PARAMETER Phase
    Install or FirstRun - which half -Execute does, and without -Execute, which half's plan a dry
    run prints. -Execute without it is refused.

.PARAMETER InstallRecordPath
    Where -Phase Install records what it installed and when, and where -Verify reads it from.

.PARAMETER Execute
    Do the -Phase named. Session 1 only: Install elevated, FirstRun NOT elevated.

.PARAMETER Verify
    Read and report. Writes nothing but its log. Takes no -Phase.

.PARAMETER WithOutlook
    With -Verify: also ask a RUNNING Outlook in this session, over COM, whether the add-in is
    connected. Never starts Outlook. COM does not attach across integrity levels, so run it at the
    level that Outlook runs at - -RunLevel Limited for one FirstRun or Start-OutlookUnelevated.ps1
    started.

.PARAMETER SelfTest
    The pure decisions against synthetic inputs, the phase rules read from this script's own
    syntax tree, and the contract this script mirrors, read from the repository source when it is
    reachable. No registry, no COM, no guest needed.

.EXAMPLE
    .\Install-OutlookAIAddIn.ps1 -SelfTest
    .\Install-OutlookAIAddIn.ps1
    .\Install-OutlookAIAddIn.ps1 -Phase FirstRun
    .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase Install -Execute"
    .\Register-InteractiveTask.ps1 -RunLevel Limited -Script "& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase FirstRun -Execute"
    .\Install-OutlookAIAddIn.ps1 -Verify
#>
[CmdletBinding()]
param(
    [string]   $PayloadRoot                = 'C:\OutlookAI-Q5\addin',
    [string]   $VstoRuntimePath            = 'C:\OutlookAI-Q5\media\vstor_redist.exe',
    [string]   $VstoRuntimeSha256          = 'CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18',
    [string]   $SuiteSourceRoot            = 'C:\OutlookAI-Q5\src',
    [string]   $SuiteSourceZip             = 'C:\OutlookAI-Q5\Source.zip',
    [string]   $OfficeVersion,
    [string[]] $ExpectedUser               = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [int]      $InstallTimeoutMinutes      = 15,
    [int]      $FirstRunTimeoutSeconds     = 240,
    [ValidateSet('Install', 'FirstRun')]
    [string]   $Phase,
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $WithOutlook,
    [switch]   $SelfTest,
    [string]   $LogPath                    = 'C:\OutlookAI-Q5\install-addin.log',
    [string]   $InstallRecordPath          = 'C:\OutlookAI-Q5\install-addin-record.json'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Everything this script reads or writes, named in one place, so the blast radius is readable
# without reading the code. WRITES are only the three marked - all three by -Phase Install, none
# by -Phase FirstRun or -Verify - and everything else is read. (Plus the log, which is not state.)
# ---------------------------------------------------------------------------------------------
$AddinName               = 'OutlookAI'
$AddinRegistrationKey    = 'Software\Microsoft\Office\Outlook\Addins\OutlookAI'   # the installer writes
$AppKey                  = 'Software\OutlookAI'                                   # InstallDir: the installer
$InstallDirValue         = 'InstallDir'
$InclusionKey            = 'Software\Microsoft\VSTO\Security\Inclusion'           # WRITES: one entry, this add-in's URL only
$EnvironmentKey          = 'Environment'                                          # WRITES: VSTO_LOGALERTS=1
# $InstallRecordPath, a parameter                                                 # WRITES: what was installed, and when
$LogAlertsName           = 'VSTO_LOGALERTS'
$PhaseInstall            = 'Install'                                              # elevated; never starts Outlook
$PhaseFirstRun           = 'FirstRun'                                             # NOT elevated; the one Outlook start
$TuningKey               = 'Software\OutlookAI\Tuning'                            # the ADD-IN writes; this reads
$TuningDesiredKey        = 'Software\OutlookAI\Tuning\Desired'                    # the ADD-IN writes; this reads
$TuningAppliedKey        = 'Software\OutlookAI\Tuning\Applied'                    # the ADD-IN writes; this reads
$McpKey                  = 'Software\OutlookAI\Mcp'                               # the ADD-IN writes; this reads
$McpReconcileValue       = 'LastReconcileUtc'                                     # under $McpKey: written at EVERY start
$VstoRuntimeKey          = 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R'            # HKLM, 32-bit view, as Installer.iss reads it
$VstoRuntimeKeyV4        = 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4'
$VstoRuntimeBytes        = 41828424
$VstoRuntimeVersion      = '10.0.60917'                                           # what that file installs
$SearchPolicyKey         = 'SOFTWARE\Policies\Microsoft\Windows\Windows Search'  # HKLM; Set-OutlookIndexingDisabled.ps1's
$CrawlRulesKey           = 'SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\WorkingSetRules'
$SupportedOffice         = @('16.0', '17.0', '15.0')                              # Services/OfficeVersions.cs, same order
$ManifestFileName        = 'addin-payload.json'
$StatusAwaitingChoice    = 'awaiting_choice'                                      # McpRegistrationService
$StatusNoClaude          = 'claude_code_not_installed'

# What the live tests read, and which test fails when it is wrong. Printed by -Verify, pinned by
# -SelfTest against HealthReporting.cs and the two tests.
$TestReads = @(
    @{ Name = 'Initialized'; Kind = 'DWord'; Means = 'tuning.managed';
       Fails = 'T3 Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning and T2 LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail' }
    @{ Name = 'Enabled'; Kind = 'DWord'; Means = 'tuning.enabled';
       Fails = 'T2 LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail' }
    @{ Name = 'LastReconcileUtc'; Kind = 'String'; Means = 'tuning.lastReconcileUtc';
       Fails = 'T2 LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail' }
)

$VerdictReady        = 'ADDIN-READY'
$VerdictNeverRan     = 'INSTALLED-NEVER-RAN'
$VerdictNotInstalled = 'NOT-INSTALLED'
$VerdictBroken       = 'BROKEN'

if ($SelfTest) { $LogPath = $null }

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    if ($LogPath) {
        try {
            $dir = Split-Path -Parent $LogPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }
        catch {
            # A log that cannot be written must not stop the work it was recording.
        }
    }
}

# =============================================================================================
# PURE DECISIONS. Each decides from its arguments alone, so -SelfTest can drive it anywhere.
# =============================================================================================

function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

# The installer's switches. /VERYSILENT and /SUPPRESSMSGBOXES: nothing on screen - a dialog on an
# unattended guest is a hang. /NOCLOSEAPPLICATIONS: never let Restart Manager close Outlook - this
# refuses to run while Outlook is up anyway. /NORESTART: a reboot is a human's decision.
function Get-InnoArguments {
    param([Parameter(Mandatory = $true)] [string] $SetupLog)
    return @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/NOCLOSEAPPLICATIONS', '/SP-', ('/LOG="' + $SetupLog + '"'))
}

# -1, 0 or 1. A missing left side is older than anything.
function Compare-DottedVersion {
    param([string] $Left, [string] $Right)
    if (-not $Left) { return -1 }
    $l = @($Left.Trim().Split('.') | ForEach-Object { [int]($_ -replace '[^\d]', '') })
    $r = @($Right.Trim().Split('.') | ForEach-Object { [int]($_ -replace '[^\d]', '') })
    $n = [Math]::Max($l.Count, $r.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $a = 0; if ($i -lt $l.Count) { $a = $l[$i] }
        $b = 0; if ($i -lt $r.Count) { $b = $r[$i] }
        if ($a -lt $b) { return -1 }
        if ($a -gt $b) { return 1 }
    }
    return 0
}

# The URL the VSTO runtime keys an inclusion entry by, from the registered Manifest value: the
# `|vstolocal` suffix off, backslashes to forward slashes, file:/// in front - which is the
# Uri.ToString() form the runtime itself writes. [MEASURED: the maintainer's workstation holds
# `file:///C:/Users/.../OutlookAI/Setup/OutlookAI.vsto` for this installer's
# `file:///{app}\OutlookAI.vsto|vstolocal`.] Spaces stay literal, as the runtime writes them.
function ConvertTo-InclusionUrl {
    param([string] $ManifestValue)
    if (-not $ManifestValue) { return $null }
    $v = $ManifestValue.Trim()
    $bar = $v.IndexOf('|')
    if ($bar -ge 0) { $v = $v.Substring(0, $bar) }
    if ($v.StartsWith('file:///', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(8) }
    elseif ($v.StartsWith('file://', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(7) }
    $v = $v.Replace('\', '/')
    return 'file:///' + $v
}

# The runtime compares entries AS URIs, and on a DOS path that comparison ignores case.
function Test-SameInclusionUrl {
    param([string] $A, [string] $B)
    $x = ConvertTo-InclusionUrl $A
    $y = ConvertTo-InclusionUrl $B
    if (-not $x -or -not $y) { return $false }
    return [string]::Equals($x, $y, [System.StringComparison]::OrdinalIgnoreCase)
}

function ConvertTo-RsaKeyValueXml {
    param([Parameter(Mandatory = $true)] [string] $ModulusBase64, [Parameter(Mandatory = $true)] [string] $ExponentBase64)
    $m = ($ModulusBase64 -replace '\s', '')
    $e = ($ExponentBase64 -replace '\s', '')
    return "<RSAKeyValue><Modulus>$m</Modulus><Exponent>$e</Exponent></RSAKeyValue>"
}

# Same rule as Testbed/host/Publish-AddInPayload.ps1: the manifest's ds:RSAKeyValue, every
# occurrence the same key, as the <RSAKeyValue> string the trust store holds.
function Get-ManifestSigningKeyXml {
    param([Parameter(Mandatory = $true)] [string] $ManifestXml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($ManifestXml)
    $nodes = $doc.SelectNodes("//*[local-name()='RSAKeyValue']")
    if ($null -eq $nodes -or $nodes.Count -eq 0) { throw 'The manifest carries no RSAKeyValue - it is not signed.' }
    $keys = @()
    foreach ($n in $nodes) {
        $mod = $n.SelectSingleNode("*[local-name()='Modulus']")
        $exp = $n.SelectSingleNode("*[local-name()='Exponent']")
        if ($null -eq $mod -or $null -eq $exp) { throw 'An RSAKeyValue in the manifest has no Modulus or no Exponent.' }
        $keys += (ConvertTo-RsaKeyValueXml -ModulusBase64 $mod.InnerText -ExponentBase64 $exp.InnerText)
    }
    $distinct = @($keys | Sort-Object -Unique)
    if ($distinct.Count -ne 1) { throw "The manifest is signed with $($distinct.Count) different keys; expected one." }
    return $distinct[0]
}

# A key as two hex strings with leading zero bytes dropped - the same two numbers the runtime's
# ExportCspBlob comparison sees, without asking a crypto provider for them.
function ConvertFrom-RsaKeyXml {
    param([string] $Xml)
    if (-not $Xml) { throw 'empty key' }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $mod = $doc.SelectSingleNode("//*[local-name()='Modulus']")
    $exp = $doc.SelectSingleNode("//*[local-name()='Exponent']")
    if ($null -eq $mod -or $null -eq $exp) { throw 'not an RSAKeyValue: Modulus or Exponent missing' }
    $out = @{}
    foreach ($pair in @(@('Modulus', $mod.InnerText), @('Exponent', $exp.InnerText))) {
        $bytes = [Convert]::FromBase64String(($pair[1] -replace '\s', ''))
        $start = 0
        while ($start -lt ($bytes.Length - 1) -and $bytes[$start] -eq 0) { $start++ }
        $hex = New-Object System.Text.StringBuilder
        for ($i = $start; $i -lt $bytes.Length; $i++) { [void]$hex.Append($bytes[$i].ToString('X2')) }
        $out[$pair[0]] = $hex.ToString()
    }
    return $out
}

function Test-SameRsaKey {
    param([string] $A, [string] $B)
    try {
        $x = ConvertFrom-RsaKeyXml $A
        $y = ConvertFrom-RsaKeyXml $B
        return ($x.Modulus -eq $y.Modulus -and $x.Exponent -eq $y.Exponent)
    }
    catch {
        return $false
    }
}

# ---- The mirror of HealthReporting.ReadTuningState. A value is @{ Kind = <RegistryValueKind>;
# Data = <object> }; a missing value is simply not in the table; a missing KEY is $null.
function ConvertTo-HealthBool {
    param($Entry)
    # HealthReporting.AsBool: `if (value is int number) return number != 0; return null;` - so ONLY
    # a REG_DWORD counts. A REG_QWORD reads as Int64 and a REG_SZ as a string: both are null.
    if ($null -eq $Entry -or [string]$Entry.Kind -ne 'DWord') { return $null }
    return ([int]$Entry.Data -ne 0)
}

function ConvertTo-HealthString {
    param($Entry)
    # `readValue(...) as string`: REG_SZ, and REG_EXPAND_SZ (which GetValue expands to a string).
    if ($null -eq $Entry) { return $null }
    $k = [string]$Entry.Kind
    if ($k -ne 'String' -and $k -ne 'ExpandString') { return $null }
    return [string]$Entry.Data
}

function Get-TuningView {
    param([hashtable] $Values)
    $view = [ordered]@{
        KeyPresent = ($null -ne $Values); Managed = $false; Enabled = $null; SearchEnabled = $null
        CachingEnabled = $null; OstEnabled = $null; RestartNeeded = $null; PolicyConflicts = $null; LastReconcileUtc = $null
        NeedsAdministrator = $null
    }
    if ($null -eq $Values) { return [pscustomobject]$view }
    if ((ConvertTo-HealthBool $Values['Initialized']) -ne $true) { return [pscustomobject]$view }
    $view.Managed = $true
    $view.Enabled = ConvertTo-HealthBool $Values['Enabled']
    $view.SearchEnabled = ConvertTo-HealthBool $Values['SearchEnabled']
    $view.CachingEnabled = ConvertTo-HealthBool $Values['CachingEnabled']
    $view.OstEnabled = ConvertTo-HealthBool $Values['OstEnabled']
    $view.RestartNeeded = ConvertTo-HealthBool $Values['RestartNeeded']
    $conflicts = ConvertTo-HealthString $Values['PolicyConflicts']
    if (-not [string]::IsNullOrWhiteSpace($conflicts)) { $view.PolicyConflicts = $conflicts }
    # Q128 (2026-10-03): the values the reconcile was refused for lack of rights, and skipped.
    $needsAdministrator = ConvertTo-HealthString $Values['NeedsAdministrator']
    if (-not [string]::IsNullOrWhiteSpace($needsAdministrator)) { $view.NeedsAdministrator = $needsAdministrator }
    $view.LastReconcileUtc = ConvertTo-HealthString $Values['LastReconcileUtc']
    return [pscustomobject]$view
}

# Since Q128 (2026-10-03) the add-in's reconcile skips a value Windows will not let it write - D25's
# five Cached Mode values under HKCU\Software\Policies, in a NOT elevated Outlook - and lists it in
# Tuning\NeedsAdministrator, which outlook_health reports. That is the designed state of a fresh
# guest, not a fault, so it is a NOTE: neither live test reads it, and ADDIN-READY does not wait on it.
function Get-NeedsAdministratorNote {
    param($View)
    if ($null -eq $View -or -not $View.NeedsAdministrator) { return $null }
    $ids = @(([string]$View.NeedsAdministrator).Split(';') | Where-Object { $_ })
    return ("the add-in skipped $($ids.Count) value(s) it may not write without an administrator, and finished its reconcile (Q128): " +
            ($ids -join ', ') + '. By design on a NOT elevated Outlook; OutlookAI Settings applies them through UAC, and ' +
            'OutlookAI.PolicyWriter.exe beside the add-in applies them when run elevated.')
}

# What the two tests would fail on, in their own terms.
function Get-TestReadProblems {
    param($View)
    $problems = @()
    if (-not $View.KeyPresent) {
        $problems += "HKCU\$TuningKey does not exist: the add-in has never run its tuning service here. tuning.managed = false fails $($TestReads[0].Fails)."
        return $problems
    }
    if ($View.Managed -ne $true) {
        $problems += "Initialized is not a nonzero REG_DWORD, so tuning.managed = false. Fails $($TestReads[0].Fails)."
        return $problems
    }
    if ($View.Enabled -ne $true) {
        $problems += "Enabled is not a nonzero REG_DWORD, so tuning.enabled is not true. Fails $($TestReads[1].Fails)."
    }
    if ($null -eq $View.LastReconcileUtc) {
        $problems += "LastReconcileUtc is not a REG_SZ, so tuning.lastReconcileUtc is null. Fails $($TestReads[2].Fails)."
    }
    return $problems
}

# OutlookTuningService writes LastReconcileUtc as DateTime.UtcNow.ToString("o"). Fresh means
# written at or after the moment this run started Outlook - proof THIS start ran the add-in.
function Test-ReconcileFresh {
    param([string] $LastReconcileUtc, [DateTime] $StartedUtc)
    if (-not $LastReconcileUtc) { return [pscustomobject]@{ Fresh = $false; Reason = 'no LastReconcileUtc at all' } }
    $parsed = [DateTime]::MinValue
    $ok = [DateTime]::TryParse($LastReconcileUtc, [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)
    if (-not $ok) { return [pscustomobject]@{ Fresh = $false; Reason = "LastReconcileUtc '$LastReconcileUtc' does not parse as a round-trip timestamp" } }
    if ($parsed.Kind -ne [DateTimeKind]::Utc) { $parsed = $parsed.ToUniversalTime() }
    # One second of slack: both readings come from this machine's clock, and "o" keeps 7 digits.
    if ($parsed -ge $StartedUtc.AddSeconds(-1)) { return [pscustomobject]@{ Fresh = $true; Reason = '' } }
    return [pscustomobject]@{ Fresh = $false; Reason = ("LastReconcileUtc {0:o} is older than this run's Outlook start {1:o}: the add-in did not run during it" -f $parsed, $StartedUtc) }
}

function Get-RegistrationProblems {
    param([string] $Manifest, $LoadBehavior, [string] $InstallDir)
    $problems = @()
    if (-not $Manifest) { return @("HKCU\$AddinRegistrationKey has no Manifest value: Outlook has nothing to load.") }
    if (-not $Manifest.EndsWith('|vstolocal', [System.StringComparison]::OrdinalIgnoreCase)) {
        $problems += "Manifest '$Manifest' is not a |vstolocal registration, which is the only kind Installer.iss writes."
    }
    if ($InstallDir) {
        $expected = ConvertTo-InclusionUrl ((Join-Path $InstallDir 'OutlookAI.vsto'))
        if (-not (Test-SameInclusionUrl $Manifest $expected)) {
            $problems += "Manifest points at '$Manifest', not at the installed copy under $InstallDir - something else re-registered the add-in."
        }
    }
    if ($null -eq $LoadBehavior) { $problems += 'LoadBehavior is missing.' }
    elseif ([int]$LoadBehavior -eq 2) { $problems += 'LoadBehavior is 2: Outlook TRIED to load the add-in and it FAILED. Read <app>\OutlookAI.vsto.log (VSTO_LOGALERTS) for why.' }
    elseif ([int]$LoadBehavior -eq 0) { $problems += 'LoadBehavior is 0: the add-in is registered but switched off.' }
    elseif ([int]$LoadBehavior -ne 3) { $problems += "LoadBehavior is $LoadBehavior, not 3 (load at startup)." }
    return $problems
}

# Outlook's hard-disable list holds UTF-16 blobs naming the add-in's path or name.
function Test-DisabledItemsMention {
    param([object[]] $Blobs)
    foreach ($b in @($Blobs)) {
        if ($null -eq $b) { continue }
        $text = [System.Text.Encoding]::Unicode.GetString([byte[]]$b)
        if ($text.ToLowerInvariant().Contains('outlookai')) { return $true }
    }
    return $false
}

# `git archive --format=zip <commit>` stores the commit id as the ZIP COMMENT, which lives at the
# end of the file in the End Of Central Directory record (signature 50 4B 05 06; the comment length
# is at offset 20 and the comment starts at offset 22).
function Get-ZipComment {
    param([byte[]] $Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 22) { return $null }
    for ($i = $Bytes.Length - 22; $i -ge [Math]::Max(0, $Bytes.Length - 22 - 65535); $i--) {
        if ($Bytes[$i] -eq 0x50 -and $Bytes[$i + 1] -eq 0x4B -and $Bytes[$i + 2] -eq 0x05 -and $Bytes[$i + 3] -eq 0x06) {
            $len = [int]$Bytes[$i + 20] + 256 * [int]$Bytes[$i + 21]
            if ($i + 22 + $len -gt $Bytes.Length) { return $null }
            return [System.Text.Encoding]::ASCII.GetString($Bytes, $i + 22, $len)
        }
    }
    return $null
}

function Test-Sha256Text { param([string] $Value) return [bool]($Value -match '^[0-9A-F]{64}$') }

function Test-PayloadManifestShape {
    param($Manifest)
    $problems = @()
    if ($null -eq $Manifest) { return @('no manifest at all') }
    if (-not ([string]$Manifest.commit -match '^[0-9a-f]{40}$')) { $problems += 'commit: not a 40-character lower-case hex id' }
    if (-not ([string]$Manifest.version -match '^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}$')) { $problems += 'version: not a four-part version' }
    if ($null -eq $Manifest.installer -or -not $Manifest.installer.file) { $problems += 'installer.file: missing' }
    elseif (-not (Test-Sha256Text ([string]$Manifest.installer.sha256))) { $problems += 'installer.sha256: not an upper-case SHA-256' }
    foreach ($f in @('OutlookAI.dll', 'OutlookAI.vsto', 'OutlookAI.dll.manifest')) {
        if ($null -eq $Manifest.addin -or -not (Test-Sha256Text ([string]$Manifest.addin.$f))) { $problems += "addin.${f}: not an upper-case SHA-256" }
    }
    if ($null -eq $Manifest.signing -or -not ([string]$Manifest.signing.publicKeyXml).StartsWith('<RSAKeyValue><Modulus>')) {
        $problems += 'signing.publicKeyXml: not an <RSAKeyValue> string'
    }
    return $problems
}

# The exclusion state, before and after this run's Outlook start.
function Compare-InteractionFacts {
    param($Before, $After)
    $changes = @()
    if ([string]$Before.PreventIndexingOutlook -ne [string]$After.PreventIndexingOutlook) {
        $changes += "PreventIndexingOutlook went from '$($Before.PreventIndexingOutlook)' to '$($After.PreventIndexingOutlook)'"
    }
    $b = @($Before.MapiRules | Sort-Object)
    $a = @($After.MapiRules | Sort-Object)
    if (($b -join ';') -ne ($a -join ';')) {
        $changes += "the mapi crawl-scope rules went from [$($b -join '; ')] to [$($a -join '; ')]"
    }
    if ([string]$Before.SearchPolicyValue -ne [string]$After.SearchPolicyValue) {
        $changes += "the Outlook Search POLICY value DisableServerAssistedSearch went from '$($Before.SearchPolicyValue)' to '$($After.SearchPolicyValue)'"
    }
    return $changes
}

function Get-McpStatusProblem {
    param([string] $Status)
    if ($Status -eq $StatusAwaitingChoice) {
        return "The add-in has a Claude Code registration QUESTION pending ($StatusAwaitingChoice). It surfaces as a MODAL dialog the next time an Outlook window is visible - in the middle of an InteractiveDesktop test. On a guest it means Claude Code is installed here; decide it once in OutlookAI Settings."
    }
    return $null
}

# Services/OfficeVersions.cs: the first of 16.0, 17.0, 15.0 whose Outlook key is a REAL hive -
# at least one value, or a subkey other than the Resiliency shell our own installer creates under
# every major.
function Select-OfficeVersion {
    param([hashtable] $Hives)
    foreach ($v in $SupportedOffice) {
        $h = $Hives[$v]
        if ($null -eq $h) { continue }
        if (@($h.Values).Count -gt 0) { return $v }
        foreach ($s in @($h.SubKeys)) {
            if (-not [string]::Equals($s, 'Resiliency', [System.StringComparison]::OrdinalIgnoreCase)) { return $v }
        }
    }
    return $null
}

# $Facts.RanSinceInstall is $null when there is no install record to compare with (an install made
# before the phases existed, or -Phase FirstRun, whose freshness test is stricter); $false when the
# tuning state is older than the last -Phase Install. $Facts.ForbidReady is -Phase Install's: it
# never starts Outlook, so whatever state it finds was written by something else.
function Get-AddInVerdict {
    param($Facts)
    $problems = @()
    $notes = @()
    if (-not $Facts.Installed) {
        return [pscustomobject]@{ Verdict = $VerdictNotInstalled; ExitCode = 3; Problems = @(); Notes = @("No InstallDir under HKCU\Software\OutlookAI and no registration: the add-in is not installed. Run -Phase $PhaseInstall -Execute (elevated), then -Phase $PhaseFirstRun -Execute (NOT elevated).") }
    }
    $problems += @($Facts.RegistrationProblems)
    $problems += @($Facts.FileProblems)
    $problems += @($Facts.TrustProblems)
    $problems += @($Facts.RuntimeProblems)
    $problems += @($Facts.ContractProblems)
    if ($Facts.DisabledItemHit) { $problems += "Outlook's Resiliency\DisabledItems names the add-in: Outlook hard-disabled it after a crash." }
    if ($Facts.McpProblem) { $problems += $Facts.McpProblem }
    if ($Facts.InstallRecordProblem) { $problems += $Facts.InstallRecordProblem }
    $problems += @($Facts.ComProblems)
    $notes += @($Facts.Notes)

    $readProblems = @(Get-TestReadProblems $Facts.Tuning)
    # Started since the install and never finished a tuning reconcile is a build that RAN and failed.
    $unfinished = Get-UnfinishedReconcileProblem $Facts
    if ($unfinished) { $problems += $unfinished }
    $notSinceInstall = ($Facts.RanSinceInstall -eq $false) -and -not $unfinished
    $neverRan = (-not $Facts.Tuning.KeyPresent) -or $notSinceInstall
    if (-not $neverRan) { $problems += $readProblems }
    if ($Facts.RequireFresh -and -not $Facts.Fresh) { $problems += $Facts.FreshReason }

    $problems = @($problems | Where-Object { $_ })
    if ($problems.Count -gt 0) {
        return [pscustomobject]@{ Verdict = $VerdictBroken; ExitCode = 1; Problems = $problems; Notes = $notes }
    }
    if ($neverRan -or $Facts.ForbidReady) {
        $why = @()
        if ($notSinceInstall -and $Facts.RanSinceInstallReason) { $why += $Facts.RanSinceInstallReason }
        if ($Facts.ForbidReady) { $why += "-Phase $PhaseInstall never starts Outlook, so it never says $VerdictReady - whatever state it finds, it did not write. -Phase $PhaseFirstRun is what proves this build runs." }
        return [pscustomobject]@{ Verdict = $VerdictNeverRan; ExitCode = 2; Problems = @(); Notes = @($notes + $why + $readProblems) }
    }
    return [pscustomobject]@{ Verdict = $VerdictReady; ExitCode = 0; Problems = @(); Notes = $notes }
}

# ---- The two phases (Q100, decided 2026-10-03). Which one runs, at which run level, and whether
# it may start Outlook - each a pure decision so -SelfTest pins it.

# The switches to a mode. -Execute without -Phase is REFUSED rather than defaulted: before the
# split it did both halves, and a command line written then must not now do one and stop.
function Resolve-RunMode {
    param([bool] $Execute, [bool] $Verify, [string] $Phase)
    $refusal = $null
    if ($Execute -and $Verify) { $refusal = 'REFUSING: -Execute and -Verify together. -Verify reads; -Execute does one phase. Pick one.' }
    elseif ($Verify -and $Phase) { $refusal = 'REFUSING: -Verify takes no -Phase - it reads what is here, whichever phase put it there.' }
    elseif ($Execute -and -not $Phase) {
        $refusal = "REFUSING: -Execute needs -Phase since 2026-10-03 (Q100). It used to install AND start Outlook from one " +
                   "elevated task, and an elevated Outlook never feeds Windows Search. Run the two halves, in this order:`n" +
                   "    $(Get-PhaseCommand -Phase $PhaseInstall)`n    $(Get-PhaseCommand -Phase $PhaseFirstRun)"
    }
    if ($refusal) { return [pscustomobject]@{ Mode = $null; Phase = $null; Refusal = $refusal } }
    if ($Execute) { return [pscustomobject]@{ Mode = $Phase; Phase = $Phase; Refusal = $null } }
    if ($Verify) { return [pscustomobject]@{ Mode = 'Verify'; Phase = $null; Refusal = $null } }
    return [pscustomobject]@{ Mode = 'DryRun'; Phase = $Phase; Refusal = $null }
}

# The one way each phase is meant to be started, as every refusal and plan prints it.
function Get-PhaseCommand {
    param([string] $Phase, [string] $ScriptPath = 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1')
    $level = ''
    if ($Phase -eq $PhaseFirstRun) { $level = '-RunLevel Limited ' }
    return ".\Register-InteractiveTask.ps1 $level-Script ""& '$ScriptPath' -Phase $Phase -Execute"""
}

# Install must be elevated - the VSTO runtime is a machine-wide install. FirstRun must NOT be: the
# Outlook it starts over COM inherits its token, and an elevated Outlook never feeds Windows Search
# (Docs/live-tier-on-the-vm.md section 8 item 22). Returns the refusal, or $null.
function Get-ElevationRefusal {
    param([string] $Phase, [bool] $Elevated)
    if ($Phase -eq $PhaseInstall -and -not $Elevated) {
        return "REFUSING TO RUN -Phase ${PhaseInstall}: this session is NOT elevated, and the VSTO runtime is a machine-wide install. Run it through Register-InteractiveTask.ps1 at its default RunLevel Highest:`n    $(Get-PhaseCommand -Phase $PhaseInstall)"
    }
    if ($Phase -eq $PhaseFirstRun -and $Elevated) {
        return "REFUSING TO RUN -Phase ${PhaseFirstRun}: this session is ELEVATED, and the Outlook it would start over COM would be too - an elevated Outlook never feeds Windows Search (Docs/live-tier-on-the-vm.md section 8 item 22), and COM does not attach across integrity levels, so the live tier could not use it either. Run it through Register-InteractiveTask.ps1 -RunLevel Limited:`n    $(Get-PhaseCommand -Phase $PhaseFirstRun)"
    }
    return $null
}

# The record -Phase Install leaves: commit, version, when the installer finished, where it put the
# add-in. Written as JSON with the time as a round-trip ("o") UTC string.
function ConvertTo-InstallRecordJson {
    param([string] $Commit, [string] $Version, [DateTime] $InstalledUtc, [string] $InstallDir)
    $o = [ordered]@{ commit = $Commit; version = $Version; installedUtc = $InstalledUtc.ToUniversalTime().ToString('o'); installDir = $InstallDir; writtenBy = "Install-OutlookAIAddIn.ps1 -Phase $PhaseInstall" }
    return ($o | ConvertTo-Json)
}

# Read back. PowerShell 7's ConvertFrom-Json turns an ISO timestamp into a [DateTime] by itself and
# 5.1's leaves it a string, so both are accepted. Returns @{ Record; Problem } - one of them $null.
function ConvertFrom-InstallRecord {
    param([string] $Json)
    $o = $null
    $problem = $null
    $when = [DateTime]::MinValue
    try { $o = $Json | ConvertFrom-Json } catch { $problem = 'it is not JSON' }
    if (-not $problem -and $null -eq $o) { $problem = 'it is empty' }
    if (-not $problem -and -not ([string]$o.commit -match '^[0-9a-f]{40}$')) { $problem = 'commit: not a 40-character lower-case hex id' }
    if (-not $problem) {
        $raw = $o.installedUtc
        if ($raw -is [DateTime]) {
            $when = $raw
            if ($when.Kind -eq [DateTimeKind]::Unspecified) { $problem = 'installedUtc: a time with no zone' }
            elseif ($when.Kind -eq [DateTimeKind]::Local) { $when = $when.ToUniversalTime() }
        }
        else {
            $ok = [DateTime]::TryParse([string]$raw, [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$when)
            if (-not $ok) { $problem = "installedUtc: '$raw' is not a round-trip timestamp" }
            elseif ($when.Kind -eq [DateTimeKind]::Unspecified) { $problem = "installedUtc: '$raw' carries no zone" }
            elseif ($when.Kind -ne [DateTimeKind]::Utc) { $when = $when.ToUniversalTime() }
        }
    }
    if ($problem) { return [pscustomobject]@{ Record = $null; Problem = $problem } }
    return [pscustomobject]@{ Problem = $null; Record = [pscustomobject]@{ Commit = [string]$o.commit; Version = [string]$o.version; InstalledUtc = $when; InstallDir = [string]$o.installDir } }
}

# The add-in's tuning reconcile walks its catalog in order, records each value it applies under
# Tuning\Applied, and writes LastReconcileUtc LAST. Until Q128 (2026-10-03) a write that threw ENDED
# the walk: the reconcile caught the exception, logged it to the debugger and wrote nothing more, so
# the Desired values with no Applied record said where a walk stopped - in Desired's own order, which
# is the catalog's (the first run writes every Desired value, in catalog order). Since Q128
# (Services\TuningReconciler.cs) a value Windows refuses for lack of rights is SKIPPED and listed in
# Tuning\NeedsAdministrator, and the walk goes on: Desired minus Applied is then those values, not a
# stopping point.
function Get-UnappliedTuning {
    param([string[]] $Desired, [string[]] $Applied)
    return @($Desired | Where-Object { $_ -and ($Applied -notcontains $_) })
}

# Has the build -Phase Install put here STARTED since, while its tuning reconcile has not finished
# since? The add-in has two startup markers: Tuning\LastReconcileUtc, the tuning reconcile's last
# write, and Mcp\LastReconcileUtc, which its registration reconcile writes at every start, failed or
# not (McpRegistrationService.Reconcile). The second newer than the install and the first not means
# the build RAN and its tuning reconcile stopped part-way - BROKEN, never INSTALLED-NEVER-RAN, which
# is what -Verify said when it read the tuning marker alone. MEASURED 2026-10-03 on
# OutlookAI-Unindexed (the first guest run of the two phases): a NOT elevated Outlook loaded the
# add-in, connected and answering, and its reconcile stopped at the first value under
# HKCU\Software\Policies - a key only an elevated token may write. Returns the problem, or $null.
function Get-UnfinishedReconcileProblem {
    param($Facts)
    if ($Facts.RanSinceInstall -ne $false -or $Facts.StartedSinceInstall -ne $true) { return $null }
    $tuningMarker = 'is absent'
    if ($Facts.Tuning -and $Facts.Tuning.LastReconcileUtc) { $tuningMarker = "is '$($Facts.Tuning.LastReconcileUtc)', from before the install" }
    $where = ''
    $unapplied = @($Facts.TuningUnapplied | Where-Object { $_ })
    if ($Facts.TuningDesiredCount -gt 0 -and $unapplied.Count -gt 0) {
        $where = " Tuning\Applied records $($Facts.TuningDesiredCount - $unapplied.Count) of the $($Facts.TuningDesiredCount) Desired values; the walk stopped at or before '$($unapplied[0])'."
    }
    return ("the add-in HAS run since -Phase $PhaseInstall installed it - its registration reconcile wrote HKCU\$McpKey\$McpReconcileValue at $($Facts.StartedSinceInstallAt) - " +
            "but its tuning reconcile has not finished since: HKCU\$TuningKey\LastReconcileUtc $tuningMarker.$where " +
            "A build from before Q128 (2026-10-03) stopped at the first write that threw - measured at a value under HKCU\Software\Policies, which only an elevated token may write (Docs/live-tier-on-the-vm.md section 2.3); " +
            "since Q128 the reconcile skips a refused value and writes LastReconcileUtc whatever it met, so in a build that has it this means the reconcile never reached its end at all.")
}

# Has the add-in written its state since the last -Phase Install? The same comparison as freshness,
# against the install's time instead of an Outlook start's.
function Test-RanSinceInstall {
    param([string] $LastReconcileUtc, [DateTime] $InstalledUtc)
    $f = Test-ReconcileFresh $LastReconcileUtc $InstalledUtc
    if ($f.Fresh) { return [pscustomobject]@{ Ran = $true; Reason = '' } }
    $why = "the add-in has not run since -Phase $PhaseInstall installed it at {0:o}" -f $InstalledUtc
    if ($LastReconcileUtc) { $why += " - LastReconcileUtc '$LastReconcileUtc' is from before that" }
    return [pscustomobject]@{ Ran = $false; Reason = ($why + ". Run -Phase $PhaseFirstRun -Execute.") }
}

# FirstRun's preflight: what must hold BEFORE an Outlook start, because with it wrong the start would
# put the trust prompt on the console - a hang on an unattended guest - or load nothing and prove
# nothing. Returns the verdict to stop with, or $null to go. A missing or stale tuning state is NOT a
# reason to stop: writing it is what the start is for.
function Get-FirstRunPreflight {
    param($Facts)
    if (-not $Facts.Installed) {
        return [pscustomobject]@{ Verdict = $VerdictNotInstalled; ExitCode = 3; Problems = @(); Notes = @("Nothing to start: the add-in is not installed. Run -Phase $PhaseInstall -Execute first:`n    $(Get-PhaseCommand -Phase $PhaseInstall)") }
    }
    $blockers = @()
    $blockers += @($Facts.RegistrationProblems)
    $blockers += @($Facts.FileProblems)
    $blockers += @($Facts.TrustProblems)
    $blockers += @($Facts.RuntimeProblems)
    if ($Facts.DisabledItemHit) { $blockers += "Outlook's Resiliency\DisabledItems names the add-in: Outlook would not load it." }
    $blockers = @($blockers | Where-Object { $_ })
    if ($blockers.Count -eq 0) { return $null }
    return [pscustomobject]@{ Verdict = $VerdictBroken; ExitCode = 1; Problems = $blockers; Notes = @("Outlook was NOT started. Each of the above is something -Phase $PhaseInstall puts right; run it again, then this:`n    $(Get-PhaseCommand -Phase $PhaseInstall)") }
}

# =============================================================================================
# THE CONTRACT THIS SCRIPT MIRRORS, read from the repository when it is reachable. Every string
# below is copied from a source file; -SelfTest fails the day one of them stops being there.
# =============================================================================================
function Get-ContractChecks {
    return @(
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningKeyPath = @"Software\OutlookAI\Tuning";'; Why = "the tests read HKCU\$TuningKey" }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningInitializedValueName = "Initialized";'; Why = 'Initialized -> tuning.managed' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningEnabledValueName = "Enabled";'; Why = 'Enabled -> tuning.enabled' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningLastReconcileUtcValueName = "LastReconcileUtc";'; Why = 'LastReconcileUtc -> tuning.lastReconcileUtc' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningPolicyConflictsValueName = "PolicyConflicts";'; Why = 'PolicyConflicts is reported' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningNeedsAdministratorValueName = "NeedsAdministrator";'; Why = 'NeedsAdministrator is reported (Q128)' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'string? needsAdministrator = readValue(Contract.TuningNeedsAdministratorValueName) as string;'; Why = 'and read as a string, as here' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string McpKeyPath = @"Software\OutlookAI\Mcp";'; Why = 'the registration status is read here' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string McpStatusValueName = "Status";'; Why = 'awaiting_choice is read from Status' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string McpLastReconcileUtcValueName = "LastReconcileUtc";'; Why = 'the add-in''s second startup marker, Mcp\LastReconcileUtc' }
        @{ File = 'Services\McpRegistrationService.cs'; Needle = 'snap.LastReconcileUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);'; Why = 'written at every reconcile, failed or not, as a round-trip UTC time' }
        @{ File = 'ThisAddIn.cs'; Needle = 'try { McpRegistrationService.Reconcile(); }'; Why = 'and that reconcile runs at every startup' }
        @{ File = 'Services\TuningReconciler.cs'; Needle = 'internal const string DesiredKeyPath = TuningKeyPath + @"\Desired";'; Why = 'the tuning walk is read from Tuning\Desired' }
        @{ File = 'Services\TuningReconciler.cs'; Needle = 'internal const string AppliedKeyPath = TuningKeyPath + @"\Applied";'; Why = 'and Tuning\Applied' }
        @{ File = 'Services\TuningReconciler.cs'; Needle = 'result.NeedsAdministrator.Add(entry.Id);'; Why = 'a value Windows refuses for lack of rights is skipped and listed, and the walk goes on (Q128)' }
        @{ File = 'Services\TuningReconciler.cs'; Needle = 'WriteBookkeeping(store, result, isStartup, utcNow);'; Why = 'and the bookkeeping, LastReconcileUtc last, is written whatever the walk met' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'if (value is int number)'; Why = 'only a REG_DWORD counts as a bool' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'if (AsBool(readValue(Contract.TuningInitializedValueName)) != true)'; Why = 'managed rests on Initialized alone' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'LastReconcileUtc = readValue(Contract.TuningLastReconcileUtcValueName) as string,'; Why = 'LastReconcileUtc must be a string' }
        @{ File = 'Services\TuningReconciler.cs'; Needle = 'WriteDword(TuningKeyPath, AddInServerContract.TuningInitializedValueName, 1);'; Why = 'the add-in writes Initialized as a DWORD' }
        @{ File = 'Services\TuningReconciler.cs'; Needle = 'string stamp = DateTime.SpecifyKind(utcNow, DateTimeKind.Utc).ToString("o", CultureInfo.InvariantCulture);'; Why = 'freshness parses a round-trip UTC timestamp' }
        @{ File = 'Services\OutlookTuningService.cs'; Needle = 'TuningReconciler.Reconcile(Store, Catalog, isStartup, DateTime.UtcNow);'; Why = 'stamped with the time of THIS reconcile' }
        @{ File = 'Services\OfficeVersions.cs'; Needle = 'internal static readonly string[] Supported = { "16.0", "17.0", "15.0" };'; Why = 'the Office majors, in detection order' }
        @{ File = 'Services\OfficeVersions.cs'; Needle = 'internal const string InstallerFootprintSubKeyName = "Resiliency";'; Why = 'a hive holding only Resiliency is not a real Outlook' }
        @{ File = 'Services\McpRegistrationService.cs'; Needle = 'internal const string StatusAwaitingChoice = "awaiting_choice";'; Why = 'a pending registration question' }
        @{ File = 'Installer.iss'; Needle = 'Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\OutlookAI.vsto|vstolocal"'; Why = 'the registration this verifies' }
        @{ File = 'Installer.iss'; Needle = 'ValueName: "LoadBehavior"; ValueData: "3"'; Why = 'LoadBehavior 3' }
        @{ File = 'Installer.iss'; Needle = 'Root: HKCU; Subkey: "Software\OutlookAI"; ValueType: string; ValueName: "InstallDir"; ValueData: "{app}"'; Why = 'InstallDir locates {app}' }
        @{ File = 'Installer.iss'; Needle = 'DefaultDirName={localappdata}\{#MyAppName}\Setup'; Why = 'per-user install location' }
        @{ File = 'Installer.iss'; Needle = 'PrivilegesRequired=lowest'; Why = 'per-user, so every key is this user''s' }
        @{ File = 'Installer.iss'; Needle = 'if IsNetFramework48Installed and (not IsVstoInstalled) then'; Why = 'the runtime step' }
        @{ File = 'Installer.iss'; Needle = "RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Version', version)"; Why = 'the runtime is detected by the 32-bit v4R key, as here' }
        @{ File = 'Installer.iss'; Needle = 'certutil'', ExpandConstant(''-f -user -addstore TrustedPublisher'; Why = 'the installer imports the certificate for this user' }
        @{ File = 'Installer.iss'; Needle = 'Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"'; Why = 'the slow-add-in exemption' }
        @{ File = 'AddInAutomation.cs'; Needle = 'bool GetRestartNeeded();'; Why = 'the call INTO the add-in the first run makes' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T3\Phase7LiveMcpToolShapeTests.cs'; Needle = 'Assert.True(report.GetProperty("tuning").GetProperty("managed").GetBoolean());'; Why = 'what the Phase-7 test asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveHealthTests.cs'; Needle = 'Assert.True(report.Tuning.Managed);'; Why = 'what LiveHealthTests asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveHealthTests.cs'; Needle = 'Assert.True(report.Tuning.Enabled);'; Why = 'what LiveHealthTests asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveHealthTests.cs'; Needle = 'Assert.NotNull(report.Tuning.LastReconcileUtc);'; Why = 'what LiveHealthTests asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveUiSearchBackendTests.cs'; Needle = 'policyValue.HasValue ? Array.Empty<string>() : new[] { HealthReporting.OutlookSearchUserKeyPath },'; Why = 'a POLICY value leaves that test nothing to flip: PROVED NOTHING on this guest, a failure on a Production profile' }
    )
}

# The payload manifest's own copy of the contract files' hashes lets the guest compare the add-in
# it installed with the suite it is about to run - the drift that would matter.
$ContractFiles = @('Services/AddInServerContract.cs', 'Services/OfficeVersions.cs')

# =============================================================================================
# I/O
# =============================================================================================

function Read-RegistryValues {
    param([Microsoft.Win32.RegistryKey] $BaseKey, [string] $Path)
    $key = $BaseKey.OpenSubKey($Path, $false)
    if ($null -eq $key) { return $null }
    try {
        $table = @{}
        foreach ($name in $key.GetValueNames()) {
            $table[$name] = @{ Kind = [string]$key.GetValueKind($name); Data = $key.GetValue($name) }
        }
        return $table
    }
    finally { $key.Close() }
}

function Get-HkcuValues { param([string] $Path) return (Read-RegistryValues -BaseKey ([Microsoft.Win32.Registry]::CurrentUser) -Path $Path) }

# Value NAMES in the key's own order, which a hashtable would lose - Get-UnappliedTuning needs it.
function Get-HkcuValueNames {
    param([string] $Path)
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($Path, $false)
    if ($null -eq $key) { return @() }
    try { return @($key.GetValueNames()) } finally { $key.Close() }
}

function Get-SubKeyNames {
    param([Microsoft.Win32.RegistryKey] $BaseKey, [string] $Path)
    $key = $BaseKey.OpenSubKey($Path, $false)
    if ($null -eq $key) { return @() }
    try { return @($key.GetSubKeyNames()) } finally { $key.Close() }
}

function Resolve-OfficeVersion {
    if ($OfficeVersion) { return $OfficeVersion }
    $hives = @{}
    foreach ($v in $SupportedOffice) {
        $path = "Software\Microsoft\Office\$v\Outlook"
        $values = Get-HkcuValues $path
        if ($null -eq $values) { continue }
        $hives[$v] = @{ Values = @($values.Keys); SubKeys = @(Get-SubKeyNames -BaseKey ([Microsoft.Win32.Registry]::CurrentUser) -Path $path) }
    }
    $found = Select-OfficeVersion -Hives $hives
    if (-not $found) { throw 'No Outlook hive under HKCU\Software\Microsoft\Office\{16.0,17.0,15.0}\Outlook. Is Office installed, and has Outlook started once for this user?' }
    return $found
}

function Get-VstoRuntimeFacts {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
    try {
        $v4r = Read-RegistryValues -BaseKey $base -Path $VstoRuntimeKey
        $v4 = Read-RegistryValues -BaseKey $base -Path $VstoRuntimeKeyV4
        $r = [ordered]@{ V4R = $null; V4 = $null }
        if ($v4r -and $v4r['Version']) { $r.V4R = [string]$v4r['Version'].Data }
        if ($v4 -and $v4['Version']) { $r.V4 = [string]$v4['Version'].Data }
        return [pscustomobject]$r
    }
    finally { $base.Close() }
}

function Get-InstallDir {
    $v = Get-HkcuValues $AppKey
    if ($v -and $v[$InstallDirValue]) { return [string]$v[$InstallDirValue].Data }
    return $null
}

function Get-InclusionEntries {
    $entries = @()
    $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($InclusionKey, $false)
    if ($null -eq $root) { return $entries }
    try {
        foreach ($name in $root.GetSubKeyNames()) {
            $k = $root.OpenSubKey($name, $false)
            if ($null -eq $k) { continue }
            try { $entries += [pscustomobject]@{ Name = $name; Url = [string]$k.GetValue('Url'); PublicKey = [string]$k.GetValue('PublicKey') } }
            finally { $k.Close() }
        }
    }
    finally { $root.Close() }
    return $entries
}

function Get-CertificateKeyXml {
    param([Parameter(Mandatory = $true)] [string] $Path)
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Path)
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cert)
    $p = $rsa.ExportParameters($false)
    return (ConvertTo-RsaKeyValueXml -ModulusBase64 ([Convert]::ToBase64String($p.Modulus)) -ExponentBase64 ([Convert]::ToBase64String($p.Exponent)))
}

# The runtime's own parse: AddInSecurityEntry's constructor calls FromXmlString and refuses the
# entry when it throws. Guest-only - it asks a crypto provider for an ephemeral container.
function Test-KeyXmlParsesLikeTheRuntime {
    param([string] $Xml)
    $csp = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    try {
        $csp.PersistKeyInCsp = $false
        $csp.FromXmlString($Xml)
        $null = $csp.ExportCspBlob($false)
        return $true
    }
    catch { return $false }
    finally { $csp.Dispose() }
}

function Get-InteractionFacts {
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $policy = Read-RegistryValues -BaseKey $hklm -Path $SearchPolicyKey
        $prevent = $null
        if ($policy -and $policy['PreventIndexingOutlook']) { $prevent = $policy['PreventIndexingOutlook'].Data }
        $rules = @()
        foreach ($name in (Get-SubKeyNames -BaseKey $hklm -Path $CrawlRulesKey)) {
            $v = Read-RegistryValues -BaseKey $hklm -Path "$CrawlRulesKey\$name"
            if ($v -and $v['URL'] -and ([string]$v['URL'].Data).StartsWith('mapi', [System.StringComparison]::OrdinalIgnoreCase)) {
                $inc = ''
                if ($v['Include']) { $inc = [string]$v['Include'].Data }
                $rules += ([string]$v['URL'].Data + ' include=' + $inc)
            }
        }
    }
    finally { $hklm.Close() }
    $searchPolicy = $null
    if ($script:ResolvedOffice) {
        $sp = Get-HkcuValues "Software\Policies\Microsoft\Office\$script:ResolvedOffice\Outlook\Search"
        if ($sp -and $sp['DisableServerAssistedSearch']) { $searchPolicy = $sp['DisableServerAssistedSearch'].Data }
    }
    return [pscustomobject]@{ PreventIndexingOutlook = $prevent; MapiRules = $rules; SearchPolicyValue = $searchPolicy }
}

function Get-SessionWindows {
    $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    return @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sid -and $_.MainWindowTitle } |
            ForEach-Object { "[$($_.ProcessName)] '$($_.MainWindowTitle)'" })
}

# -Phase Install's record, read back. Absent is not a problem - every install made before the
# phases existed has none - but an unreadable one is.
function Read-InstallRecord {
    if (-not (Test-Path -LiteralPath $InstallRecordPath)) { return [pscustomobject]@{ Record = $null; Problem = $null } }
    $r = ConvertFrom-InstallRecord ([System.IO.File]::ReadAllText($InstallRecordPath))
    if ($r.Problem) { $r.Problem = "the install record $InstallRecordPath is unreadable ($($r.Problem)). Run -Phase $PhaseInstall -Execute again, which rewrites it." }
    return $r
}

# Whether THIS process's token is elevated - the same reading Register-InteractiveTask.ps1 makes for
# its -RunLevel Limited check.
function Test-TokenElevated {
    $p = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
    return [bool]$p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Start a process and wait for it with a deadline. .Handle is read before waiting because
# Start-Process -PassThru otherwise loses the exit code (MEASURED, Install-DotnetSdk.ps1). NEVER
# -Verb RunAs: this session is already elevated, and UAC elevation takes the foreground.
function Invoke-Installer {
    param([string] $FilePath, [string[]] $Arguments, [int] $TimeoutMinutes)
    $p = Start-Process -FilePath $FilePath -ArgumentList ($Arguments -join ' ') -PassThru -NoNewWindow
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = $null; TimedOut = $true }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false }
}

# =============================================================================================
# GUARDS
# =============================================================================================
function Assert-TestbedGuestLocal {
    if (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $ExpectedUser -Prefix $ExpectedComputerNamePrefix) { return }
    throw @"
REFUSING TO RUN.

  logged on as : '$env:USERNAME'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$env:COMPUTERNAME'      (must start with: '$ExpectedComputerNamePrefix')

This script installs the OutlookAI add-in, rewrites a VSTO trust entry and starts Outlook. On the
maintainer's workstation that would replace the add-in his Outlook really runs, and nothing would
say so until it misbehaved.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2), and
Testbed/host/New-AnswerFile.ps1 derives their computer name by replacing 'OutlookAI-' with 'OAI-',
so a guest built from the answer file matches both axes. If you named a guest something else, say so:

    -ExpectedUser <username> -ExpectedComputerNamePrefix <prefix>

Do not 'fix' this by widening either default. The defaults are the guard.
"@
}

function Assert-PhaseElevation {
    param([string] $ForPhase)
    $refusal = Get-ElevationRefusal -Phase $ForPhase -Elevated (Test-TokenElevated)
    if ($refusal) { throw $refusal }
}

function Assert-Bitness {
    if (-not [Environment]::Is64BitProcess) {
        throw 'REFUSING TO RUN: this is a 32-bit PowerShell. Outlook on the guests is x64, and a 32-bit client would start a COM surrogate rather than talk to it. Use C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe.'
    }
}

function Assert-InteractiveSession {
    param([string] $ForPhase)
    $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    if ($sid -ne 0) { return }
    $why = 'FirstRun starts Outlook, and Outlook cannot finish starting in session 0 - PowerShell Direct lands there, and the call would hang instead of failing.'
    if ($ForPhase -eq $PhaseInstall) {
        $why = 'Install runs where it was measured: the interactive session, through the task. Over PowerShell Direct its installers have never run, and the first-run half needs the task anyway.'
    }
    throw @"
REFUSING TO RUN -Phase $ForPhase IN SESSION 0.

$why Run it through the interactive session:

    $(Get-PhaseCommand -Phase $ForPhase -ScriptPath $PSCommandPath)

-Verify, without -WithOutlook, is fine here.
"@
}

function Assert-OutlookClosed {
    param([string] $ForPhase)
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return }
    $why = 'The add-in loads when Outlook STARTS, and the proof this phase gives - the add-in wrote its state AFTER a start this script made, in an Outlook whose token it read - needs that start to be this script''s.'
    if ($ForPhase -eq $PhaseInstall) { $why = 'The installer would be replacing files under a running add-in.' }
    throw @"
REFUSING TO RUN -Phase ${ForPhase}: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')).

$why

Restart the guest with Testbed/host/Restart-Guest.ps1 - the proven way to a clean Outlook here - and
run this again. DO NOT taskkill OUTLOOK.EXE: mailbox-safety rule 7 forbids it outright.
"@
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()
    # -ceq and String.Contains only. Never -like: Testbed/README.md section 4b.
    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:Checks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = ($Expected -join ' | ') }
        if ($Actual -is [System.Array]) { $a = ($Actual -join ' | ') }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else {
            $script:Failures += "$What : expected [$e], got [$a]"
            Write-Host "  FAIL $What - expected [$e], got [$a]"
        }
    }
    function V([string] $kind, $data) { return @{ Kind = $kind; Data = $data } }

    Write-Host '== the guard =='
    Test-Case 'vmadmin on OAI-INDEXED passes' $true (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'OAI-INDEXED' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'the maintainer workstation is refused' $false (Test-GuestIdentity -UserName 'jori' -ComputerName 'PC657' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'vmadmin on a non-guest name is refused' $false (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'PC657' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'the right name as another user is refused' $false (Test-GuestIdentity -UserName 'jori' -ComputerName 'OAI-INDEXED' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'an empty prefix is refused, never "matches everything"' $false (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'OAI-X' -Users @('vmadmin') -Prefix '')

    Write-Host ''
    Write-Host '== the installer switches =='
    $inno = Get-InnoArguments -SetupLog 'C:\OutlookAI-Q5\s.log'
    Test-Case 'are exactly these' '/VERYSILENT | /SUPPRESSMSGBOXES | /NORESTART | /NOCLOSEAPPLICATIONS | /SP- | /LOG="C:\OutlookAI-Q5\s.log"' $inno
    Test-Case 'never merely /SILENT, which shows progress' $false ($inno -contains '/SILENT')

    Write-Host ''
    Write-Host '== the VSTO runtime version =='
    Test-Case 'the pinned version is current' 0 (Compare-DottedVersion '10.0.60917' $VstoRuntimeVersion)
    Test-Case 'what Office itself may bring is older' -1 (Compare-DottedVersion '10.0.60910' $VstoRuntimeVersion)
    Test-Case 'a later one is newer' 1 (Compare-DottedVersion '10.0.60918' $VstoRuntimeVersion)
    Test-Case 'a two-part pad compares' 0 (Compare-DottedVersion '10.0.60917.0' $VstoRuntimeVersion)
    Test-Case 'absent is older than anything' -1 (Compare-DottedVersion '' $VstoRuntimeVersion)

    Write-Host ''
    Write-Host '== the trust entry URL =='
    Test-Case 'from the installer''s registration, as the runtime wrote it on the host' 'file:///C:/Users/vmadmin/AppData/Local/OutlookAI/Setup/OutlookAI.vsto' (ConvertTo-InclusionUrl 'file:///C:\Users\vmadmin\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal')
    Test-Case 'from a plain path' 'file:///C:/a/OutlookAI.vsto' (ConvertTo-InclusionUrl 'C:\a\OutlookAI.vsto')
    Test-Case 'spaces stay literal, as the runtime writes them' 'file:///C:/a b/x.vsto' (ConvertTo-InclusionUrl 'file:///C:\a b\x.vsto|vstolocal')
    Test-Case 'an already-normal URL is unchanged' 'file:///C:/a/x.vsto' (ConvertTo-InclusionUrl 'file:///C:/a/x.vsto')
    Test-Case 'the same URL in another case is the same entry' $true (Test-SameInclusionUrl 'file:///c:/USERS/x.vsto' 'file:///C:/Users/x.vsto')
    Test-Case 'a different folder is a different entry' $false (Test-SameInclusionUrl 'file:///C:/a/x.vsto' 'file:///C:/b/x.vsto')
    Test-Case 'nothing is never a match' $false (Test-SameInclusionUrl '' 'file:///C:/a/x.vsto')

    Write-Host ''
    Write-Host '== the signing key =='
    $k1 = '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>'
    $manifest = '<asmv1:assembly xmlns:asmv1="urn:schemas-microsoft-com:asm.v1"><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><KeyInfo><KeyValue><RSAKeyValue><Modulus>AAEC
AwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue></KeyValue></KeyInfo></Signature></asmv1:assembly>'
    Test-Case 'read out of a signed manifest, namespace and line break ignored' $k1 (Get-ManifestSigningKeyXml -ManifestXml $manifest)
    Test-Case 'the same key with a namespace and whitespace is the same key' $true (Test-SameRsaKey $k1 '<RSAKeyValue xmlns="http://www.w3.org/2000/09/xmldsig#"> <Modulus>AAEC AwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')
    Test-Case 'a leading zero byte does not make a different key' $true (Test-SameRsaKey '<RSAKeyValue><Modulus>AQIDBA==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>' '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')
    Test-Case 'a different modulus is a different key' $false (Test-SameRsaKey $k1 '<RSAKeyValue><Modulus>BBBBBBB=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')
    Test-Case 'a different exponent is a different key' $false (Test-SameRsaKey $k1 '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>Aw==</Exponent></RSAKeyValue>')
    Test-Case 'garbage is never the same key' $false (Test-SameRsaKey $k1 'not xml')
    $threw = $false; try { $null = Get-ManifestSigningKeyXml '<a/>' } catch { $threw = $true }
    Test-Case 'an unsigned manifest is refused' $true $threw

    Write-Host ''
    Write-Host '== what the tests read: the mirror of HealthReporting.ReadTuningState =='
    $good = @{ Initialized = (V 'DWord' 1); Enabled = (V 'DWord' 1); SearchEnabled = (V 'DWord' 1); CachingEnabled = (V 'DWord' 1)
               OstEnabled = (V 'DWord' 1); RestartNeeded = (V 'DWord' 1); PolicyConflicts = (V 'String' ''); LastReconcileUtc = (V 'String' '2026-09-24T12:00:00.0000000Z') }
    $v = Get-TuningView $good
    Test-Case 'a first-run state is managed' $true $v.Managed
    Test-Case 'and enabled' $true $v.Enabled
    Test-Case 'and carries its reconcile time' '2026-09-24T12:00:00.0000000Z' $v.LastReconcileUtc
    Test-Case 'an empty PolicyConflicts reads as none' $null $v.PolicyConflicts
    Test-Case 'and the tests have nothing to fail on' 0 (Get-TestReadProblems $v).Count
    $v = Get-TuningView $null
    Test-Case 'no key: not managed' $false $v.Managed
    Test-Case 'no key: named as never having run' $true ((Get-TestReadProblems $v) -join ' ').Contains('has never run its tuning service')
    $bad = $good.Clone(); $bad['Initialized'] = (V 'QWord' 1)
    Test-Case 'Initialized as a REG_QWORD 1 is NOT managed - AsBool takes an int only' $false (Get-TuningView $bad).Managed
    $bad = $good.Clone(); $bad['Initialized'] = (V 'String' '1')
    Test-Case 'Initialized as the string "1" is NOT managed' $false (Get-TuningView $bad).Managed
    $bad = $good.Clone(); $bad['Initialized'] = (V 'DWord' 0)
    Test-Case 'Initialized 0 is not managed' $false (Get-TuningView $bad).Managed
    Test-Case 'and the problem names both tests' $true ((Get-TestReadProblems (Get-TuningView $bad)) -join ' ').Contains('Phase7LiveMcpToolShapeTests')
    $bad = $good.Clone(); $bad.Remove('Initialized')
    Test-Case 'no Initialized at all is not managed' $false (Get-TuningView $bad).Managed
    $bad = $good.Clone(); $bad['Enabled'] = (V 'DWord' 0)
    Test-Case 'Enabled 0 fails LiveHealthTests only' $true ((Get-TestReadProblems (Get-TuningView $bad)) -join ' ').Contains('tuning.enabled is not true')
    $bad = $good.Clone(); $bad['Enabled'] = (V 'QWord' 1)
    Test-Case 'Enabled as a REG_QWORD is not true' $null (Get-TuningView $bad).Enabled
    $bad = $good.Clone(); $bad.Remove('LastReconcileUtc')
    Test-Case 'no LastReconcileUtc is a null the test fails on' $true ((Get-TestReadProblems (Get-TuningView $bad)) -join ' ').Contains('tuning.lastReconcileUtc is null')
    $bad = $good.Clone(); $bad['LastReconcileUtc'] = (V 'MultiString' @('2026-09-24T12:00:00Z'))
    Test-Case 'LastReconcileUtc as a REG_MULTI_SZ is null - "as string" fails on string[]' $null (Get-TuningView $bad).LastReconcileUtc
    $bad = $good.Clone(); $bad['LastReconcileUtc'] = (V 'ExpandString' '2026-09-24T12:00:00Z')
    Test-Case 'LastReconcileUtc as a REG_EXPAND_SZ is still a string' '2026-09-24T12:00:00Z' (Get-TuningView $bad).LastReconcileUtc
    $bad = $good.Clone(); $bad['PolicyConflicts'] = (V 'String' 'caching.policy.SyncWindowSetting')
    Test-Case 'a real PolicyConflicts is reported' 'caching.policy.SyncWindowSetting' (Get-TuningView $bad).PolicyConflicts
    Test-Case 'no NeedsAdministrator reads as none' $null (Get-TuningView $good).NeedsAdministrator
    $five = 'caching.policy.SyncWindowSetting;caching.policy.SyncWindowSettingDays;caching.policy.DownloadSharedFolders;caching.policy.CacheOthersMail;caching.policy.DisableSyncSliderForSharedMailbox'
    $na = $good.Clone(); $na['NeedsAdministrator'] = (V 'String' $five)
    Test-Case 'NeedsAdministrator is reported as written (Q128)' $five (Get-TuningView $na).NeedsAdministrator
    $na['NeedsAdministrator'] = (V 'String' '')
    Test-Case 'an empty NeedsAdministrator reads as none' $null (Get-TuningView $na).NeedsAdministrator
    $na['NeedsAdministrator'] = (V 'String' $five)
    Test-Case 'and it is not a problem either test fails on' 0 @(Get-TestReadProblems (Get-TuningView $na)).Count
    Test-Case 'it becomes a note naming all five' $true ([string](Get-NeedsAdministratorNote (Get-TuningView $na))).Contains('5 value(s)')
    Test-Case 'no list, no note' $null (Get-NeedsAdministratorNote (Get-TuningView $good))

    Write-Host ''
    Write-Host '== freshness =='
    $t0 = New-Object DateTime(2026, 9, 24, 12, 0, 0, [DateTimeKind]::Utc)
    Test-Case 'written after the start is fresh' $true (Test-ReconcileFresh '2026-09-24T12:00:05.1234567Z' $t0).Fresh
    Test-Case 'written within the second before is fresh' $true (Test-ReconcileFresh '2026-09-24T11:59:59.5000000Z' $t0).Fresh
    Test-Case 'written an hour before is not' $false (Test-ReconcileFresh '2026-09-24T11:00:00.0000000Z' $t0).Fresh
    Test-Case 'and says the add-in did not run during this start' $true (Test-ReconcileFresh '2026-09-24T11:00:00.0000000Z' $t0).Reason.Contains('did not run during it')
    Test-Case 'an offset timestamp is compared in UTC' $true (Test-ReconcileFresh '2026-09-24T14:00:05.0000000+02:00' $t0).Fresh
    Test-Case 'garbage does not parse' $false (Test-ReconcileFresh 'yesterday' $t0).Fresh
    Test-Case 'absent is not fresh' $false (Test-ReconcileFresh '' $t0).Fresh

    Write-Host ''
    Write-Host '== the registration =='
    $dir = 'C:\Users\vmadmin\AppData\Local\OutlookAI\Setup'
    $reg = 'file:///C:\Users\vmadmin\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal'
    Test-Case 'the installer''s own registration is clean' 0 (Get-RegistrationProblems -Manifest $reg -LoadBehavior 3 -InstallDir $dir).Count
    Test-Case 'LoadBehavior 2 is a failed load, and says where the reason is' $true ((Get-RegistrationProblems -Manifest $reg -LoadBehavior 2 -InstallDir $dir) -join ' ').Contains('OutlookAI.vsto.log')
    Test-Case 'LoadBehavior 0 is switched off' $true ((Get-RegistrationProblems -Manifest $reg -LoadBehavior 0 -InstallDir $dir) -join ' ').Contains('switched off')
    Test-Case 'a build output registered instead is caught' $true ((Get-RegistrationProblems -Manifest 'file:///C:/Source/OutlookAI/bin/Release/OutlookAI.vsto|vstolocal' -LoadBehavior 3 -InstallDir $dir) -join ' ').Contains('not at the installed copy')
    Test-Case 'a ClickOnce-style registration is caught' $true ((Get-RegistrationProblems -Manifest 'file:///C:\x\OutlookAI.vsto' -LoadBehavior 3 -InstallDir '') -join ' ').Contains('|vstolocal')
    Test-Case 'no registration at all' $true ((Get-RegistrationProblems -Manifest '' -LoadBehavior $null -InstallDir $dir) -join ' ').Contains('nothing to load')

    Write-Host ''
    Write-Host '== Outlook''s hard-disable list =='
    $blob = [System.Text.Encoding]::Unicode.GetBytes('file:///C:/Users/vmadmin/AppData/Local/OutlookAI/Setup/OutlookAI.vsto|vstolocal')
    Test-Case 'a blob naming the add-in is found' $true (Test-DisabledItemsMention @(, $blob))
    Test-Case 'another add-in is not' $false (Test-DisabledItemsMention @(, ([System.Text.Encoding]::Unicode.GetBytes('SomeOtherAddin.dll'))))
    Test-Case 'nothing is nothing' $false (Test-DisabledItemsMention @())

    Write-Host ''
    Write-Host '== the suite''s commit, from its zip comment =='
    $comment = 'fb19ccf494fe25460b985ed0f4cf0cc1435e3e40'
    $eocd = New-Object byte[] 22
    $eocd[0] = 0x50; $eocd[1] = 0x4B; $eocd[2] = 0x05; $eocd[3] = 0x06; $eocd[20] = [byte]$comment.Length; $eocd[21] = 0
    $zip = [byte[]](@(1, 2, 3, 4) + $eocd + [System.Text.Encoding]::ASCII.GetBytes($comment))
    Test-Case 'the comment is the commit' $comment (Get-ZipComment $zip)
    $noComment = [byte[]](@(9, 9) + $eocd)
    $noComment[2 + 20] = 0
    Test-Case 'a zip with no comment gives an empty one' '' (Get-ZipComment $noComment)
    Test-Case 'not a zip gives nothing' $null (Get-ZipComment ([byte[]](1..40)))

    Write-Host ''
    Write-Host '== the payload manifest =='
    $okManifest = [pscustomobject]@{
        commit = ('a' * 40); version = '99.99.99.0'
        installer = [pscustomobject]@{ file = 'OutlookAI-v99.99.99.0.exe'; sha256 = ('A' * 64) }
        addin = [pscustomobject]@{ 'OutlookAI.dll' = ('B' * 64); 'OutlookAI.vsto' = ('C' * 64); 'OutlookAI.dll.manifest' = ('D' * 64) }
        signing = [pscustomobject]@{ publicKeyXml = $k1 }
    }
    Test-Case 'a complete manifest is accepted' 0 (Test-PayloadManifestShape $okManifest).Count
    $bad = $okManifest.PSObject.Copy(); $bad.version = '3.1.0'
    Test-Case 'a three-part version is refused' $true ((Test-PayloadManifestShape $bad) -join ' ').Contains('version:')
    $bad = $okManifest.PSObject.Copy(); $bad.installer = [pscustomobject]@{ file = 'x.exe'; sha256 = 'abc' }
    Test-Case 'a short hash is refused' $true ((Test-PayloadManifestShape $bad) -join ' ').Contains('installer.sha256')

    Write-Host ''
    Write-Host '== the index exclusion, before and after =='
    $before = [pscustomobject]@{ PreventIndexingOutlook = 1; MapiRules = @('mapi16://{S-1-5-21-1}/ include=0'); SearchPolicyValue = $null }
    Test-Case 'unchanged is unchanged' 0 (Compare-InteractionFacts $before $before).Count
    $after = [pscustomobject]@{ PreventIndexingOutlook = 1; MapiRules = @('mapi16://{S-1-5-21-1}/ include=1'); SearchPolicyValue = $null }
    Test-Case 'a rule flipping back to included is caught' $true ((Compare-InteractionFacts $before $after) -join ' ').Contains('crawl-scope rules')
    $after = [pscustomobject]@{ PreventIndexingOutlook = $null; MapiRules = @('mapi16://{S-1-5-21-1}/ include=0'); SearchPolicyValue = $null }
    Test-Case 'the policy going away is caught' $true ((Compare-InteractionFacts $before $after) -join ' ').Contains('PreventIndexingOutlook')
    $after = [pscustomobject]@{ PreventIndexingOutlook = 1; MapiRules = @('mapi16://{S-1-5-21-1}/ include=0'); SearchPolicyValue = 1 }
    Test-Case 'a Search POLICY value appearing is caught' $true ((Compare-InteractionFacts $before $after) -join ' ').Contains('POLICY')

    Write-Host ''
    Write-Host '== the Office hive, as Services/OfficeVersions.cs picks it =='
    Test-Case 'a real 16.0 hive wins' '16.0' (Select-OfficeVersion @{ '16.0' = @{ Values = @('x'); SubKeys = @('Profiles') }; '15.0' = @{ Values = @(); SubKeys = @('Resiliency') } })
    Test-Case 'the installer''s Resiliency shell alone is not a hive' $null (Select-OfficeVersion @{ '15.0' = @{ Values = @(); SubKeys = @('Resiliency') }; '17.0' = @{ Values = @(); SubKeys = @('Resiliency') } })
    Test-Case '17.0 is tried before 15.0' '17.0' (Select-OfficeVersion @{ '17.0' = @{ Values = @(); SubKeys = @('Search') }; '15.0' = @{ Values = @('a'); SubKeys = @() } })

    Write-Host ''
    Write-Host '== the registration question =='
    Test-Case 'a pending question is a problem' $true ([bool](Get-McpStatusProblem $StatusAwaitingChoice))
    Test-Case 'no Claude Code on the guest is not' $null (Get-McpStatusProblem $StatusNoClaude)

    Write-Host ''
    Write-Host '== the verdict =='
    $installedFacts = @{ Installed = $true; RegistrationProblems = @(); FileProblems = @(); TrustProblems = @(); RuntimeProblems = @(); ContractProblems = @()
                         DisabledItemHit = $false; McpProblem = $null; ComProblems = @(); Notes = @(); Tuning = (Get-TuningView $good); RequireFresh = $false; Fresh = $false; FreshReason = ''
                         InstallRecordProblem = $null; RanSinceInstall = $null; RanSinceInstallReason = ''; ForbidReady = $false }
    $r = Get-AddInVerdict ([pscustomobject]$installedFacts)
    Test-Case 'everything in place is READY' $VerdictReady $r.Verdict
    Test-Case 'and exits 0' 0 $r.ExitCode
    $f = $installedFacts.Clone(); $f.Tuning = (Get-TuningView $null)
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case 'installed but never ran is its own verdict' $VerdictNeverRan $r.Verdict
    Test-Case 'exiting 2' 2 $r.ExitCode
    $f = $installedFacts.Clone(); $f.Tuning = (Get-TuningView $null); $f.TrustProblems = @('no trust entry')
    Test-Case 'never ran AND untrusted is BROKEN - the first load would prompt' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.RequireFresh = $true; $f.Fresh = $false; $f.FreshReason = 'stale'
    Test-Case 'under -Phase FirstRun a stale state is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.DisabledItemHit = $true
    Test-Case 'a hard-disabled add-in is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $bad = $good.Clone(); $bad['Enabled'] = (V 'DWord' 0); $f.Tuning = (Get-TuningView $bad)
    Test-Case 'a state the tests would fail on is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.Installed = $false
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case 'not installed is its own verdict' $VerdictNotInstalled $r.Verdict
    Test-Case 'exiting 3' 3 $r.ExitCode
    Test-Case 'and naming both phases to run' $true ((($r.Notes -join ' ').Contains("-Phase $PhaseInstall")) -and (($r.Notes -join ' ').Contains("-Phase $PhaseFirstRun")))
    $f = $installedFacts.Clone(); $f.RanSinceInstall = $false; $f.RanSinceInstallReason = 'not since the install'
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case 'a valid state OLDER than the last -Phase Install is NEVER-RAN, not READY' $VerdictNeverRan $r.Verdict
    Test-Case 'exiting 2, and saying why' '2|True' ('{0}|{1}' -f $r.ExitCode, ($r.Notes -join ' ').Contains('not since the install'))
    $f = $installedFacts.Clone(); $f.RanSinceInstall = $true
    Test-Case 'the same state written after it is READY' $VerdictReady (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.RanSinceInstall = $null
    Test-Case 'no install record - an install from before the split - reads as before: READY' $VerdictReady (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.ForbidReady = $true
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case '-Phase Install never says READY, even over a valid state an earlier run left' $VerdictNeverRan $r.Verdict
    Test-Case 'and exits 2, not 0' 2 $r.ExitCode
    Test-Case 'and says the first run is what proves it' $true (($r.Notes -join ' ').Contains("-Phase $PhaseFirstRun is what proves"))
    $f = $installedFacts.Clone(); $f.ForbidReady = $true; $f.TrustProblems = @('no trust entry')
    Test-Case '-Phase Install with anything wrong is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.InstallRecordProblem = 'the install record is unreadable'
    Test-Case 'an unreadable install record is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    # MEASURED 2026-10-03, the first guest run of the two phases: the add-in started in a NOT elevated
    # Outlook - its registration reconcile wrote Mcp\LastReconcileUtc - and its tuning reconcile
    # stopped part-way, leaving Initialized and no LastReconcileUtc. -Verify said NEVER-RAN.
    $noReconcile = $good.Clone(); $noReconcile.Remove('LastReconcileUtc')
    $f = $installedFacts.Clone(); $f.Tuning = (Get-TuningView $noReconcile); $f.RanSinceInstall = $false; $f.RanSinceInstallReason = 'not since the install'
    $f.StartedSinceInstall = $true; $f.StartedSinceInstallAt = '2026-10-03T16:01:28.0000000Z'; $f.TuningDesiredCount = 13
    $f.TuningUnapplied = @('caching.policy.SyncWindowSetting', 'caching.policy.SyncWindowSettingDays')
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case 'started since the install, tuning reconcile never finished: BROKEN, not NEVER-RAN' $VerdictBroken $r.Verdict
    Test-Case 'and exits 1' 1 $r.ExitCode
    $x = ($r.Problems -join ' ')
    Test-Case 'saying it HAS run, when, and where its walk stopped' $true ($x.Contains('HAS run since') -and $x.Contains('2026-10-03T16:01:28') -and $x.Contains("records 11 of the 13 Desired values; the walk stopped at or before 'caching.policy.SyncWindowSetting'"))
    Test-Case 'and that the test-read problem is reported too' $true $x.Contains('tuning.lastReconcileUtc is null')
    $f.Tuning = (Get-TuningView $good)
    Test-Case 'the same with an OLDER valid state - one an earlier build left - is BROKEN too' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    Test-Case 'and names that older time' $true ((Get-AddInVerdict ([pscustomobject]$f)).Problems -join ' ').Contains("is '2026-09-24T12:00:00.0000000Z', from before the install")
    $f.StartedSinceInstall = $false
    Test-Case 'neither marker since the install is still NEVER-RAN' $VerdictNeverRan (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f.StartedSinceInstall = $true; $f.RanSinceInstall = $true
    Test-Case 'both markers since the install is READY' $VerdictReady (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f.RanSinceInstall = $null
    Test-Case 'no install record ignores the second marker, as before the split' $VerdictReady (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    # Q128: a fresh guest's designed state - five values skipped for an administrator, the reconcile
    # finished - is ADDIN-READY. The values are a note, never a problem.
    $f = $installedFacts.Clone(); $f.Tuning = (Get-TuningView $na); $f.TuningDesiredCount = 13
    $f.TuningUnapplied = @(($five -split ';')); $f.StartedSinceInstall = $true; $f.RanSinceInstall = $true
    Test-Case 'five values needing an administrator, reconcile finished: READY' $VerdictReady (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.ForbidReady = $true; $f.RanSinceInstall = $false; $f.StartedSinceInstall = $false
    Test-Case '-Phase Install over a state that ran, both markers now older than its record: NEVER-RAN' $VerdictNeverRan (Get-AddInVerdict ([pscustomobject]$f)).Verdict

    Write-Host ''
    Write-Host '== how far the add-in''s tuning walk got =='
    Test-Case 'Desired minus Applied, in Desired''s order' 'c.p.a | c.p.b | o.c' (Get-UnappliedTuning -Desired @('s.a', 's.b', 'c.p.a', 'c.p.b', 'o.c') -Applied @('s.b', 's.a'))
    Test-Case 'everything applied is nothing' '' (Get-UnappliedTuning -Desired @('s.a', 's.b') -Applied @('s.a', 's.b', 'x'))
    Test-Case 'no Desired values is nothing' '' (Get-UnappliedTuning -Desired @() -Applied @('s.a'))

    Write-Host ''
    Write-Host '== the two phases: which one runs, and how each is started =='
    $m = Resolve-RunMode -Execute $true -Verify $false -Phase ''
    Test-Case '-Execute with no -Phase is refused - it used to do both halves' 'True|' ('{0}|{1}' -f [bool]$m.Refusal, $m.Mode)
    $refusal = [string]$m.Refusal
    $i1 = $refusal.IndexOf("-Phase $PhaseInstall -Execute"); $i2 = $refusal.IndexOf("-Phase $PhaseFirstRun -Execute")
    Test-Case 'and the refusal gives both commands, the install first' $true ($i1 -ge 0 -and $i2 -gt $i1)
    Test-Case '-Phase Install -Execute is the install' $PhaseInstall (Resolve-RunMode -Execute $true -Verify $false -Phase $PhaseInstall).Mode
    Test-Case '-Phase FirstRun -Execute is the first run' $PhaseFirstRun (Resolve-RunMode -Execute $true -Verify $false -Phase $PhaseFirstRun).Mode
    Test-Case '-Verify is a verify' 'Verify' (Resolve-RunMode -Execute $false -Verify $true -Phase '').Mode
    Test-Case '-Verify with a -Phase is refused' $true ([bool](Resolve-RunMode -Execute $false -Verify $true -Phase $PhaseInstall).Refusal)
    Test-Case '-Execute -Verify is refused' $true ([bool](Resolve-RunMode -Execute $true -Verify $true -Phase $PhaseInstall).Refusal)
    $m = Resolve-RunMode -Execute $false -Verify $false -Phase ''
    Test-Case 'no switch is a dry run of both' 'DryRun|' ('{0}|{1}' -f $m.Mode, $m.Phase)
    $m = Resolve-RunMode -Execute $false -Verify $false -Phase $PhaseFirstRun
    Test-Case '-Phase FirstRun alone is a dry run of it' "DryRun|$PhaseFirstRun" ('{0}|{1}' -f $m.Mode, $m.Phase)
    Test-Case 'the install is started at the task''s default run level, Highest' ".\Register-InteractiveTask.ps1 -Script ""& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase $PhaseInstall -Execute""" (Get-PhaseCommand -Phase $PhaseInstall)
    Test-Case 'the first run at -RunLevel Limited' ".\Register-InteractiveTask.ps1 -RunLevel Limited -Script ""& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase $PhaseFirstRun -Execute""" (Get-PhaseCommand -Phase $PhaseFirstRun)

    Write-Host ''
    Write-Host '== the run level each phase demands =='
    Test-Case 'Install, elevated: runs' $null (Get-ElevationRefusal -Phase $PhaseInstall -Elevated $true)
    Test-Case 'Install, NOT elevated: refused - the VSTO runtime is a machine-wide install' $true ([string](Get-ElevationRefusal -Phase $PhaseInstall -Elevated $false)).Contains('machine-wide')
    Test-Case 'FirstRun, NOT elevated: runs' $null (Get-ElevationRefusal -Phase $PhaseFirstRun -Elevated $false)
    $x = [string](Get-ElevationRefusal -Phase $PhaseFirstRun -Elevated $true)
    Test-Case 'FirstRun, elevated: refused - an elevated Outlook never feeds Windows Search' $true $x.Contains('never feeds Windows Search')
    Test-Case 'and the refusal says how to start it instead' $true $x.Contains('-RunLevel Limited -Script')

    Write-Host ''
    Write-Host '== the install record =='
    $t1 = (New-Object DateTime(2026, 10, 3, 1, 2, 3, [DateTimeKind]::Utc)).AddTicks(1234567)
    $rr = ConvertFrom-InstallRecord (ConvertTo-InstallRecordJson -Commit ('a' * 40) -Version '99.99.99.0' -InstalledUtc $t1 -InstallDir 'C:\x')
    Test-Case 'a record reads back' $null $rr.Problem
    Test-Case 'with its commit' ('a' * 40) $rr.Record.Commit
    Test-Case 'and its time to the tick, in UTC - under either shell''s ConvertFrom-Json' ('{0}|Utc' -f $t1.Ticks) ('{0}|{1}' -f $rr.Record.InstalledUtc.Ticks, $rr.Record.InstalledUtc.Kind)
    $rr = ConvertFrom-InstallRecord ('{"commit":"' + ('a' * 40) + '","installedUtc":"2026-10-03T03:02:03.0000000+02:00"}')
    Test-Case 'an offset time is read as the same instant in UTC' ('{0}|Utc' -f $t1.AddTicks(-1234567).Ticks) ('{0}|{1}' -f $rr.Record.InstalledUtc.Ticks, $rr.Record.InstalledUtc.Kind)
    Test-Case 'a short commit is refused' $true ([string](ConvertFrom-InstallRecord '{"commit":"abc","installedUtc":"2026-10-03T01:02:03.0000000Z"}').Problem).Contains('commit')
    Test-Case 'a time with no zone is refused' $true ([string](ConvertFrom-InstallRecord ('{"commit":"' + ('a' * 40) + '","installedUtc":"2026-10-03T01:02:03"}')).Problem).Contains('zone')
    Test-Case 'not JSON is refused' $true ([bool](ConvertFrom-InstallRecord 'not json').Problem)

    Write-Host ''
    Write-Host '== has the installed build run since -Phase Install? =='
    $ti = New-Object DateTime(2026, 10, 3, 12, 0, 0, [DateTimeKind]::Utc)
    Test-Case 'reconciled after the install: it has' $true (Test-RanSinceInstall '2026-10-03T12:00:05.0000000Z' $ti).Ran
    $x = Test-RanSinceInstall '2026-10-02T12:00:00.0000000Z' $ti
    Test-Case 'reconciled the day before: it has not' $false $x.Ran
    Test-Case 'and the reason says so, and what to run' $true ($x.Reason.Contains('has not run since') -and $x.Reason.Contains("-Phase $PhaseFirstRun -Execute"))
    Test-Case 'no state at all: it has not' $false (Test-RanSinceInstall '' $ti).Ran

    Write-Host ''
    Write-Host '== the first run''s preflight: what stops it before Outlook starts =='
    $pf = $installedFacts.Clone(); $pf.Tuning = (Get-TuningView $null)
    Test-Case 'installed and trusted, never ran: go' $null (Get-FirstRunPreflight ([pscustomobject]$pf))
    $pf = $installedFacts.Clone(); $bad = $good.Clone(); $bad['Enabled'] = (V 'DWord' 0); $pf.Tuning = (Get-TuningView $bad)
    Test-Case 'a state the tests would fail on does not stop it - rewriting that is the point' $null (Get-FirstRunPreflight ([pscustomobject]$pf))
    $pf = $installedFacts.Clone(); $pf.ContractProblems = @('contract drift'); $pf.McpProblem = 'a pending question'
    Test-Case 'nor do contract drift and a pending question - the verdict reports both' $null (Get-FirstRunPreflight ([pscustomobject]$pf))
    $pf = $installedFacts.Clone(); $pf.TrustProblems = @('no trust entry')
    $x = Get-FirstRunPreflight ([pscustomobject]$pf)
    Test-Case 'untrusted stops it - the start would put the trust prompt on the console' $VerdictBroken $x.Verdict
    Test-Case 'and says Outlook was not started' $true (($x.Notes -join ' ').Contains('Outlook was NOT started'))
    $pf = $installedFacts.Clone(); $pf.DisabledItemHit = $true
    Test-Case 'hard-disabled stops it' $VerdictBroken (Get-FirstRunPreflight ([pscustomobject]$pf)).Verdict
    $pf = $installedFacts.Clone(); $pf.FileProblems = @('another build is installed')
    Test-Case 'another build installed stops it' $VerdictBroken (Get-FirstRunPreflight ([pscustomobject]$pf)).Verdict
    $pf = $installedFacts.Clone(); $pf.RuntimeProblems = @('no runtime')
    Test-Case 'no runtime stops it' $VerdictBroken (Get-FirstRunPreflight ([pscustomobject]$pf)).Verdict
    $pf = $installedFacts.Clone(); $pf.Installed = $false
    $x = Get-FirstRunPreflight ([pscustomobject]$pf)
    Test-Case 'not installed stops it, exiting 3' "$VerdictNotInstalled|3" ('{0}|{1}' -f $x.Verdict, $x.ExitCode)

    Write-Host ''
    Write-Host '== the two phases, read from this script''s own syntax tree =='
    # The decision itself (Q100): Install never starts Outlook, FirstRun installs and writes nothing,
    # and each refuses the wrong token first. Read from the code, so a later edit that moves an
    # Outlook start back into the install fails here rather than on a guest's index.
    $selfAst = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$null)
    $defs = @{}
    foreach ($fd in $selfAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defs[$fd.Name] = $fd }
    function Get-PhaseReach([string] $from) {
        # Every function of this file a phase can reach, itself included.
        $seen = @{}
        $todo = New-Object System.Collections.Queue
        $todo.Enqueue($from)
        while ($todo.Count -gt 0) {
            $name = [string]$todo.Dequeue()
            if ($seen.ContainsKey($name) -or -not $defs.ContainsKey($name)) { continue }
            $seen[$name] = $true
            foreach ($c in $defs[$name].Body.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                $cn = $c.GetCommandName()
                if ($cn -and $defs.ContainsKey($cn)) { $todo.Enqueue($cn) }
            }
        }
        return @($seen.Keys)
    }
    function Get-OutlookStartsIn($Ast, [string] $label) {
        # What starts Outlook or attaches to one: a job, a COM object, GetActiveObject, the two jobs.
        $hits = @()
        foreach ($x in $Ast.FindAll({ param($y) $true }, $true)) {
            if ($x -is [System.Management.Automation.Language.CommandAst]) {
                $cn = $x.GetCommandName()
                if ($cn -eq 'Start-Job') { $hits += ('{0}: Start-Job' -f $label) }
                elseif ($cn -eq 'New-Object' -and $x.Extent.Text.Contains('-ComObject')) { $hits += ('{0}: New-Object -ComObject' -f $label) }
            }
            elseif ($x -is [System.Management.Automation.Language.VariableExpressionAst] -and @('FirstRunJob', 'AttachJob') -contains $x.VariablePath.UserPath) {
                $hits += ('{0}: ${1}' -f $label, $x.VariablePath.UserPath)
            }
            elseif ($x -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $x.Member.Extent.Text -eq 'GetActiveObject') {
                $hits += ('{0}: GetActiveObject' -f $label)
            }
        }
        return $hits
    }
    function Get-WritesIn($Ast, [string] $label) {
        # Installers, registry and environment writes, and files other than the log (Say's).
        $hits = @()
        $members = @('SetValue', 'CreateSubKey', 'DeleteSubKey', 'DeleteSubKeyTree', 'DeleteValue', 'SetEnvironmentVariable', 'WriteAllText')
        $commands = @('Start-Process', 'Invoke-Installer', 'Set-Item', 'Remove-Item', 'Set-Content', 'New-ItemProperty', 'Set-ItemProperty', 'Remove-ItemProperty')
        foreach ($x in $Ast.FindAll({ param($y) $true }, $true)) {
            if ($x -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $members -contains $x.Member.Extent.Text) {
                $hits += ('{0}: .{1}()' -f $label, $x.Member.Extent.Text)
            }
            elseif ($x -is [System.Management.Automation.Language.CommandAst] -and $commands -contains $x.GetCommandName()) {
                $hits += ('{0}: {1}' -f $label, $x.GetCommandName())
            }
        }
        return $hits
    }
    function Get-FirstCommandIn([string] $fn) {
        $all = @($defs[$fn].Body.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true) | Sort-Object { $_.Extent.StartOffset })
        if ($all.Count -eq 0) { return '' }
        return $all[0].Extent.Text
    }
    $installHits = @(); $installWrites = @()
    foreach ($n in (Get-PhaseReach 'Invoke-InstallPhase')) { $installHits += Get-OutlookStartsIn $defs[$n].Body $n; $installWrites += Get-WritesIn $defs[$n].Body $n }
    $firstHits = @(); $firstWrites = @()
    foreach ($n in (Get-PhaseReach 'Invoke-FirstRunPhase')) { $firstHits += Get-OutlookStartsIn $defs[$n].Body $n; $firstWrites += Get-WritesIn $defs[$n].Body $n }
    $jobAssign = @($selfAst.FindAll({ param($x) $x -is [System.Management.Automation.Language.AssignmentStatementAst] -and $x.Left.Extent.Text -eq '$FirstRunJob' }, $true))
    Test-Case 'the install phase reaches no Outlook start, no COM attach and no first-run job' '' ($installHits -join '; ')
    Test-Case 'the first-run phase does reach one - so the line above is not vacuous' $true ($firstHits.Count -gt 0)
    Test-Case 'the first-run phase reaches no installer and no registry, environment or file write' '' ($firstWrites -join '; ')
    Test-Case 'nor does the job it starts' '1|' ('{0}|{1}' -f $jobAssign.Count, ((@($jobAssign | ForEach-Object { Get-WritesIn $_.Right '$FirstRunJob' })) -join '; '))
    Test-Case 'the install phase does reach them - so the line above is not vacuous' $true ($installWrites.Count -gt 0)
    Test-Case 'the install phase refuses a token that is not elevated before anything else' 'Assert-PhaseElevation -ForPhase $PhaseInstall' (Get-FirstCommandIn 'Invoke-InstallPhase')
    Test-Case 'the first-run phase refuses an elevated one before anything else' 'Assert-PhaseElevation -ForPhase $PhaseFirstRun' (Get-FirstCommandIn 'Invoke-FirstRunPhase')

    Write-Host ''
    Write-Host '== the contract this script mirrors, read from the repository =='
    # Two levels up from Testbed\guest is the repository. Staged on its own - C:\OutlookAI-Q5\ on a
    # guest, before the suite's source is there - two levels up is '' and Join-Path refused it:
    # "Cannot bind argument to parameter 'Path' because it is an empty string", which ended the
    # FIRST guest -SelfTest (2026-09-24, OutlookAI-Unindexed) before this section could say SKIP.
    $repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    if ($repo -and (Test-Path -LiteralPath (Join-Path $repo 'Services\AddInServerContract.cs'))) {
        foreach ($c in (Get-ContractChecks)) {
            $path = Join-Path $repo $c.File
            $text = ''
            if (Test-Path -LiteralPath $path) { $text = [System.IO.File]::ReadAllText($path) }
            Test-Case "$($c.File): $($c.Why)" $true ($text.Contains($c.Needle))
        }
    }
    else {
        Write-Host "  SKIP no repository two levels above $PSScriptRoot - the contract cannot be read from here. Run -SelfTest from the staged suite source (C:\OutlookAI-Q5\src\Testbed\guest) to include it."
    }

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a guest can settle these, and nothing above stands in for them.'
    Write-Host 'Settled on OutlookAI-Unindexed, 2026-10-03, by the two phases (the banner):'
    Write-Host '  * vstor_redist.exe /q /norestart installs silently from an elevated interactive task, the'
    Write-Host '    installer /VERYSILENT installs for vmadmin, and the written inclusion entry retires the'
    Write-Host '    trust prompt for the build''s key - and none of it starts Outlook'
    Write-Host '  * -Phase Install ends INSTALLED-NEVER-RAN on a fresh guest and over an install that has run'
    Write-Host '  * -Phase FirstRun from a RunLevel Limited task starts an Outlook whose token reads NOT'
    Write-Host '    elevated, and the add-in loads in it, connected and answering'
    Write-Host '  * the unindexed guest is still UNINDEXED after it'
    Write-Host 'NOT SETTLED:'
    Write-Host '  * ADDIN-READY on a fresh guest. There the add-in''s tuning reconcile cannot finish NOT'
    Write-Host '    elevated - it writes under HKCU\Software\Policies - so FirstRun ends BROKEN: a product'
    Write-Host '    defect, open in TODO.md. ADDIN-READY came only with those values set first, as a GPO would'
    Write-Host '  * that the indexed guest stays INDEXED, its index taking the first run''s Outlook'
    Write-Host '  * that the two live tests then pass on BOTH guests, which is the claim all of this is for'

    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# FACTS - read, never written.
# =============================================================================================
function Get-AddInFacts {
    param([bool] $RequireFresh, [DateTime] $StartedUtc, $Payload, $InstallRecord, [string] $InstallRecordProblem)
    $facts = [ordered]@{
        Installed = $false; RegistrationProblems = @(); FileProblems = @(); TrustProblems = @(); RuntimeProblems = @()
        ContractProblems = @(); DisabledItemHit = $false; McpProblem = $null; ComProblems = @(); Notes = @()
        Tuning = $null; RequireFresh = $RequireFresh; Fresh = $false; FreshReason = ''
        InstallRecordProblem = $InstallRecordProblem; RanSinceInstall = $null; RanSinceInstallReason = ''; ForbidReady = $false
        StartedSinceInstall = $null; StartedSinceInstallAt = ''; TuningDesiredCount = 0; TuningUnapplied = @()
    }

    $installDir = Get-InstallDir
    $reg = Get-HkcuValues $AddinRegistrationKey
    $manifestValue = $null; $loadBehavior = $null
    if ($reg) {
        if ($reg['Manifest']) { $manifestValue = [string]$reg['Manifest'].Data }
        if ($reg['LoadBehavior']) { $loadBehavior = $reg['LoadBehavior'].Data }
    }
    $facts.Installed = [bool]($installDir -or $reg)
    Say "  InstallDir    : $installDir"
    Say "  registration  : Manifest=$manifestValue  LoadBehavior=$loadBehavior"
    $facts.RegistrationProblems = @(Get-RegistrationProblems -Manifest $manifestValue -LoadBehavior $loadBehavior -InstallDir $installDir)

    # Files, and which build they are.
    $vsto = $null
    if ($installDir) {
        foreach ($f in @('OutlookAI.vsto', 'OutlookAI.dll.manifest', 'OutlookAI.dll', 'OutlookAI.cer')) {
            if (-not (Test-Path -LiteralPath (Join-Path $installDir $f))) { $facts.FileProblems += "{app}\$f is missing ($installDir)." }
        }
        $vsto = Join-Path $installDir 'OutlookAI.vsto'
        if ($Payload -and (Test-Path -LiteralPath (Join-Path $installDir 'OutlookAI.dll'))) {
            $dllHash = (Get-FileHash -LiteralPath (Join-Path $installDir 'OutlookAI.dll') -Algorithm SHA256).Hash
            if ($dllHash -ne [string]$Payload.addin.'OutlookAI.dll') {
                $facts.FileProblems += "the installed OutlookAI.dll ($dllHash) is not the payload's ($($Payload.addin.'OutlookAI.dll')): this guest runs a different build than the one staged."
            }
            else { Say "  build         : OutlookAI.dll is the payload's, commit $($Payload.commit)" }
        }
        $logFile = Join-Path $installDir 'OutlookAI.vsto.log'
        if (Test-Path -LiteralPath $logFile) {
            $facts.Notes += "the VSTO runtime logged errors to $logFile (VSTO_LOGALERTS). Its tail:"
            foreach ($l in (Get-Content -LiteralPath $logFile -Tail 12)) { $facts.Notes += "    $l" }
        }
    }

    # The VSTO runtime.
    $rt = Get-VstoRuntimeFacts
    Say "  VSTO runtime  : v4R=$($rt.V4R)  v4=$($rt.V4)"
    if ((Compare-DottedVersion $rt.V4R $VstoRuntimeVersion) -lt 0) {
        $facts.RuntimeProblems += "VSTO Runtime Setup\v4R reports '$($rt.V4R)', below the pinned $VstoRuntimeVersion. -Phase $PhaseInstall -Execute installs it from $VstoRuntimePath."
    }
    $userEnv = Get-HkcuValues $EnvironmentKey
    if (-not ($userEnv -and $userEnv[$LogAlertsName] -and [string]$userEnv[$LogAlertsName].Data -eq '1')) {
        $facts.Notes += "$LogAlertsName is not 1 for this user, so a failed load would leave no OutlookAI.vsto.log. -Phase $PhaseInstall sets it."
    }

    # Trust.
    if ($manifestValue -and $vsto -and (Test-Path -LiteralPath $vsto)) {
        $url = ConvertTo-InclusionUrl $manifestValue
        $key = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText($vsto))
        # Not $matches: that is PowerShell's automatic -match variable, and any -match in scope rewrites it.
        $urlEntries = @(Get-InclusionEntries | Where-Object { Test-SameInclusionUrl $_.Url $url })
        $good = @($urlEntries | Where-Object { Test-SameRsaKey $_.PublicKey $key })
        Say "  trust         : $($urlEntries.Count) inclusion entr(y/ies) for $url, $($good.Count) with this build's key"
        if ($good.Count -eq 0) {
            $facts.TrustProblems += "no VSTO inclusion entry vouches for $url with the key the installed manifest is signed by: the first load would put the ClickOnce trust prompt on the console, which on this guest is a hang."
        }
        if (-not (Test-KeyXmlParsesLikeTheRuntime $key)) { $facts.TrustProblems += 'the manifest key does not parse with FromXmlString, which is how the runtime reads an entry.' }
        $cer = Join-Path $installDir 'OutlookAI.cer'
        if (Test-Path -LiteralPath $cer) {
            if (-not (Test-SameRsaKey (Get-CertificateKeyXml -Path $cer) $key)) { $facts.TrustProblems += '{app}\OutlookAI.cer is not the certificate the manifest is signed with.' }
        }
        if ($Payload -and -not (Test-SameRsaKey ([string]$Payload.signing.publicKeyXml) $key)) {
            $facts.TrustProblems += 'the installed manifest is not signed by the key the payload manifest names.'
        }
    }

    # Outlook's own disable machinery.
    $office = $script:ResolvedOffice
    $disabled = Get-HkcuValues "Software\Microsoft\Office\$office\Outlook\Resiliency\DisabledItems"
    if ($disabled) {
        $blobs = @(); foreach ($k in $disabled.Keys) { if ($disabled[$k].Kind -eq 'Binary') { $blobs += , ([byte[]]$disabled[$k].Data) } }
        $facts.DisabledItemHit = Test-DisabledItemsMention $blobs
    }
    $dnd = Get-HkcuValues "Software\Microsoft\Office\$office\Outlook\Resiliency\DoNotDisableAddinList"
    if (-not ($dnd -and $dnd[$AddinName])) {
        $facts.Notes += "no DoNotDisableAddinList entry for $AddinName under Office $office - Outlook may disable it after a slow start (the installer normally writes one)."
    }

    # The registration question.
    $mcp = Get-HkcuValues $McpKey
    $mcpStatus = $null
    if ($mcp -and $mcp['Status']) { $mcpStatus = [string]$mcp['Status'].Data }
    $mcpReconciled = ''
    if ($mcp -and $mcp[$McpReconcileValue]) { $mcpReconciled = [string]$mcp[$McpReconcileValue].Data }
    Say "  registration question status: $mcpStatus  (its reconcile, at every start: $mcpReconciled)"
    $facts.McpProblem = Get-McpStatusProblem $mcpStatus

    # The Search POLICY value that leaves LiveUiSearchBackendTests nothing to flip.
    $sp = Get-HkcuValues "Software\Policies\Microsoft\Office\$office\Outlook\Search"
    if ($sp -and $sp['DisableServerAssistedSearch']) {
        $facts.Notes += "a POLICY DisableServerAssistedSearch exists: T2 LiveUiSearchBackendTests cannot flip both states and will print PROVED NOTHING on this guest (it fails on a Production profile). Nothing in the add-in writes it; find what did."
    }

    # The state the tests read.
    $facts.Tuning = Get-TuningView (Get-HkcuValues $TuningKey)
    # And how far the add-in's last tuning walk got - which says where one that never finished stopped.
    $desiredNames = @(Get-HkcuValueNames $TuningDesiredKey)
    $facts.TuningDesiredCount = $desiredNames.Count
    $facts.TuningUnapplied = @(Get-UnappliedTuning -Desired $desiredNames -Applied @(Get-HkcuValueNames $TuningAppliedKey))
    if ($desiredNames.Count -gt 0) {
        $walk = "  tuning walk   : Tuning\Applied records {0} of the {1} Tuning\Desired values" -f ($desiredNames.Count - $facts.TuningUnapplied.Count), $desiredNames.Count
        if ($facts.TuningUnapplied.Count -gt 0) { $walk += '; never applied: ' + ($facts.TuningUnapplied -join ', ') }
        Say $walk
    }
    $needsAdministratorNote = Get-NeedsAdministratorNote $facts.Tuning
    if ($needsAdministratorNote) {
        Say "  needs an administrator: $(($facts.Tuning.NeedsAdministrator -split ';') -join ', ')"
        $facts.Notes += $needsAdministratorNote
    }
    if ($RequireFresh) {
        $fresh = Test-ReconcileFresh ([string]$facts.Tuning.LastReconcileUtc) $StartedUtc
        $facts.Fresh = $fresh.Fresh
        $facts.FreshReason = $fresh.Reason
    }
    elseif ($InstallRecord) {
        # Not under FirstRun: written by THIS start already implies written after the install.
        $since = Test-RanSinceInstall -LastReconcileUtc ([string]$facts.Tuning.LastReconcileUtc) -InstalledUtc $InstallRecord.InstalledUtc
        $facts.RanSinceInstall = $since.Ran
        $facts.RanSinceInstallReason = $since.Reason
        # The add-in's other startup marker: did this build START since, whatever its tuning did?
        $facts.StartedSinceInstall = (Test-ReconcileFresh $mcpReconciled $InstallRecord.InstalledUtc).Fresh
        $facts.StartedSinceInstallAt = $mcpReconciled
        $ranText = 'has run since'
        if (-not $since.Ran) { $ranText = 'has NOT run since' }
        if (-not $since.Ran -and $facts.StartedSinceInstall) { $ranText = 'has STARTED since, and its tuning reconcile has NOT finished since' }
        Say ("  install record: commit {0}, installed {1:o} - the add-in {2}" -f $InstallRecord.Commit, $InstallRecord.InstalledUtc, $ranText)
        if ($Payload -and $InstallRecord.Commit -ne [string]$Payload.commit) {
            $facts.Notes += "the install record names commit $($InstallRecord.Commit), and the payload staged now is $($Payload.commit): -Phase $PhaseInstall has not installed this payload yet."
        }
    }

    # The add-in's contract against the suite's.
    if ($Payload -and $Payload.contract) {
        foreach ($rel in $ContractFiles) {
            $suiteFile = Join-Path $SuiteSourceRoot ($rel.Replace('/', '\'))
            $want = [string]$Payload.contract.$rel
            if (-not (Test-Path -LiteralPath $suiteFile)) { $facts.Notes += "no suite source at $suiteFile, so the add-in's $rel was not compared with the suite's."; continue }
            $have = (Get-FileHash -LiteralPath $suiteFile -Algorithm SHA256).Hash
            if ($have -ne $want) {
                $facts.ContractProblems += "$rel differs between the add-in (built from $($Payload.commit)) and the suite at ${SuiteSourceRoot}: the tests may read names this add-in does not write. Build both from one commit."
            }
        }
    }
    elseif ($Payload) {
        $facts.Notes += 'the payload manifest carries no contract hashes, so the add-in was not compared with the suite.'
    }
    if ($Payload -and (Test-Path -LiteralPath $SuiteSourceZip)) {
        $suiteCommit = Get-ZipComment ([System.IO.File]::ReadAllBytes($SuiteSourceZip))
        if ($suiteCommit -and $suiteCommit -ne [string]$Payload.commit) {
            $facts.Notes += "the add-in was built from $($Payload.commit) and the suite from $suiteCommit. Fine while the contract files agree (checked above); one commit for both is the tidy state."
        }
        elseif ($suiteCommit) { Say "  suite         : built from the same commit, $suiteCommit" }
    }

    return [pscustomobject]$facts
}

function Write-TestReadBlock {
    param($View)
    Say ''
    Say "== What the live tests read - HKCU\$TuningKey, through HealthReporting.ReadTuningState =="
    $raw = Get-HkcuValues $TuningKey
    foreach ($t in $TestReads) {
        $shown = '(absent)'
        if ($raw -and $raw[$t.Name]) { $shown = "$($raw[$t.Name].Kind) $($raw[$t.Name].Data)" }
        Say ("  {0,-17} {1,-38} -> {2}" -f $t.Name, $shown, $t.Means)
    }
    Say ("  => managed={0} enabled={1} lastReconcileUtc={2}" -f $View.Managed, $View.Enabled, $View.LastReconcileUtc)
    Say ("     searchEnabled={0} cachingEnabled={1} ostEnabled={2} restartNeeded={3} policyConflicts={4}" -f
        $View.SearchEnabled, $View.CachingEnabled, $View.OstEnabled, $View.RestartNeeded, $View.PolicyConflicts)
    Say ("     needsAdministrator={0}" -f $View.NeedsAdministrator)
}

function Write-Verdict {
    param($Result)
    Say ''
    Say "VERDICT: $($Result.Verdict)"
    foreach ($p in $Result.Problems) { Say "  - $p" }
    foreach ($n in $Result.Notes) { Say "  note: $n" }
    Say ''
    Say 'A REGISTRY VALUE IS NOT A PASSING TEST. The verdict above says the state the tests read is in'
    Say 'place. The claim that matters - both tests pass on this guest - is made by running them.'
}

function Write-PhasePlan {
    param([string] $ForPhase)
    if (-not $ForPhase -or $ForPhase -eq $PhaseInstall) {
        Say "DRY RUN. Nothing is written. -Phase $PhaseInstall -Execute would, in session 1, ELEVATED, with Outlook closed:"
        Say "  1. check   $PayloadRoot\$ManifestFileName and the installer's SHA-256 against it"
        Say "  2. install the VSTO runtime from $VstoRuntimePath (SHA-256 $VstoRuntimeSha256), unless v4R >= $VstoRuntimeVersion"
        Say ("  3. run     the installer " + ((Get-InnoArguments -SetupLog 'C:\OutlookAI-Q5\install-addin-setup.log') -join ' '))
        Say "  4. write   ONE entry under HKCU\$InclusionKey for the installed manifest's URL and signing key"
        Say "  5. write   HKCU\$EnvironmentKey $LogAlertsName=1"
        Say "  6. write   $InstallRecordPath - what it installed, and when"
        Say "  7. verify  and stop at ${VerdictNeverRan}: it NEVER starts Outlook"
        Say "  through:   $(Get-PhaseCommand -Phase $PhaseInstall)"
        Say ''
    }
    if (-not $ForPhase -or $ForPhase -eq $PhaseFirstRun) {
        Say "DRY RUN. Nothing is written. -Phase $PhaseFirstRun -Execute would, in session 1, NOT elevated, with Outlook closed:"
        Say '  1. check   the add-in is installed, registered, trusted and the payload''s build - or stop, Outlook untouched'
        Say '  2. start   Outlook once over COM, headless, in a watchdogged child job; read its token; never quit or kill it'
        Say "  3. verify  the state the tests read under HKCU\$TuningKey, written by THAT start, and no dialog left on screen"
        Say "  through:   $(Get-PhaseCommand -Phase $PhaseFirstRun)"
        Say '  On the unindexed guest only once its index exclusion is in place (Testbed/README.md section 1, step 7b).'
        Say ''
    }
    Say 'Run -Verify to report on what is here now.'
}

# -Phase Install: elevated, and it never starts Outlook. Ends INSTALLED-NEVER-RAN or BROKEN.
function Invoke-InstallPhase {
    param($Payload)
    Assert-PhaseElevation -ForPhase $PhaseInstall
    Assert-Bitness
    Assert-InteractiveSession -ForPhase $PhaseInstall
    Assert-OutlookClosed -ForPhase $PhaseInstall

    $installer = Join-Path $PayloadRoot ([string]$Payload.installer.file)
    if (-not (Test-Path -LiteralPath $installer)) { throw "REFUSING: the installer the manifest names is not at $installer." }
    $installerHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
    if ($installerHash -ne [string]$Payload.installer.sha256) {
        throw "REFUSING: $installer is not the file the manifest pins (expected $($Payload.installer.sha256), got $installerHash)."
    }
    Say "  installer $installer - SHA-256 matches the manifest"

    $interactionBefore = Get-InteractionFacts
    Say ("  index exclusion before: PreventIndexingOutlook={0} mapi rules=[{1}]" -f $interactionBefore.PreventIndexingOutlook, ($interactionBefore.MapiRules -join '; '))

    # A record is written only once an install has completed; the old one goes first, so a run that
    # fails part-way leaves none rather than one describing the build it replaced.
    if (Test-Path -LiteralPath $InstallRecordPath) {
        Remove-Item -LiteralPath $InstallRecordPath -Force
        Say "  removed the previous install record $InstallRecordPath"
    }

    # ---- 3. The VSTO runtime -------------------------------------------------------------------
    Say ''
    Say '== The VSTO runtime =='
    $rt = Get-VstoRuntimeFacts
    if ((Compare-DottedVersion $rt.V4R $VstoRuntimeVersion) -ge 0) {
        Say "  already registered: v4R $($rt.V4R) - not reinstalling"
    }
    else {
        Say "  registered now: v4R '$($rt.V4R)', v4 '$($rt.V4)' - installing $VstoRuntimeVersion"
        if (-not (Test-Path -LiteralPath $VstoRuntimePath)) {
            throw "REFUSING: no VSTO runtime at $VstoRuntimePath. It is STAGED MEDIA (Testbed/MEDIA.md) - never downloaded here; the guest has no network. Copy it in with Testbed/host/Copy-ToGuest.ps1."
        }
        $rtHash = (Get-FileHash -LiteralPath $VstoRuntimePath -Algorithm SHA256).Hash
        $rtLength = (Get-Item -LiteralPath $VstoRuntimePath).Length
        if ($rtHash -ne $VstoRuntimeSha256.ToUpperInvariant() -or $rtLength -ne $VstoRuntimeBytes) {
            throw "REFUSING: $VstoRuntimePath is not the pinned redistributable ($rtHash, $rtLength bytes; expected $VstoRuntimeSha256, $VstoRuntimeBytes bytes)."
        }
        # Installer.iss's own switches for it: /q /norestart.
        $run = Invoke-Installer -FilePath $VstoRuntimePath -Arguments @('/q', '/norestart') -TimeoutMinutes $InstallTimeoutMinutes
        if ($run.TimedOut) { throw "The VSTO runtime installer did not finish within $InstallTimeoutMinutes minutes. Its logs are in $env:TEMP." }
        if ($null -eq $run.ExitCode) { throw 'The VSTO runtime installer exit code came back EMPTY - this script failed to read it, which is not the same as the install failing. Check v4R by hand.' }
        if ($run.ExitCode -ne 0 -and $run.ExitCode -ne 3010) { throw "The VSTO runtime installer exited $($run.ExitCode). Its logs are in $env:TEMP." }
        if ($run.ExitCode -eq 3010) { Say '  exit 3010: installed, and Windows wants a reboot. Not a failure; reboot before taking a checkpoint.' }
        $rt = Get-VstoRuntimeFacts
        if ((Compare-DottedVersion $rt.V4R $VstoRuntimeVersion) -lt 0) { throw "The VSTO runtime installer exited $($run.ExitCode) but v4R still reports '$($rt.V4R)'." }
        Say "  installed: v4R $($rt.V4R)"
    }

    # ---- 4. The product's installer ------------------------------------------------------------
    Say ''
    Say '== The OutlookAI installer, silent =='
    $setupLog = Join-Path (Split-Path -Parent $LogPath) 'install-addin-setup.log'
    $run = Invoke-Installer -FilePath $installer -Arguments (Get-InnoArguments -SetupLog $setupLog) -TimeoutMinutes $InstallTimeoutMinutes
    if ($run.TimedOut -or $null -eq $run.ExitCode -or $run.ExitCode -ne 0) {
        if (Test-Path -LiteralPath $setupLog) { foreach ($l in (Get-Content -LiteralPath $setupLog -Tail 25)) { Say "      | $l" } }
        throw "The installer exited '$($run.ExitCode)' (timed out: $($run.TimedOut)). Its log: $setupLog"
    }
    # From here the payload's build is what Outlook would load: any run after this moment is a run of it.
    $installedUtc = [DateTime]::UtcNow
    $installDir = Get-InstallDir
    if (-not $installDir) { throw "The installer exited 0 but wrote no HKCU\$AppKey\$InstallDirValue. Its log: $setupLog" }
    Say "  installed into $installDir (log: $setupLog)"

    # ---- 5. Trust ------------------------------------------------------------------------------
    Say ''
    Say '== Trust: the inclusion entry the trust prompt would have written =='
    $regValues = Get-HkcuValues $AddinRegistrationKey
    if (-not ($regValues -and $regValues['Manifest'])) { throw "The installer exited 0 but wrote no HKCU\$AddinRegistrationKey\Manifest." }
    $url = ConvertTo-InclusionUrl ([string]$regValues['Manifest'].Data)
    $key = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText((Join-Path $installDir 'OutlookAI.vsto')))
    if (-not (Test-SameRsaKey $key ([string]$Payload.signing.publicKeyXml))) { throw 'REFUSING TO TRUST: the installed manifest is not signed by the key the payload manifest names.' }
    if (-not (Test-SameRsaKey $key (Get-CertificateKeyXml -Path (Join-Path $installDir 'OutlookAI.cer')))) { throw 'REFUSING TO TRUST: the installed manifest is not signed by the installed OutlookAI.cer.' }
    if (-not (Test-KeyXmlParsesLikeTheRuntime $key)) { throw 'REFUSING TO TRUST: the key does not parse with FromXmlString, so the runtime would reject the entry.' }
    $root = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($InclusionKey)
    try {
        $kept = $false
        foreach ($name in $root.GetSubKeyNames()) {
            $k = $root.OpenSubKey($name, $false)
            $entryUrl = $null; $entryKey = $null
            if ($null -ne $k) { try { $entryUrl = [string]$k.GetValue('Url'); $entryKey = [string]$k.GetValue('PublicKey') } finally { $k.Close() } }
            if (-not (Test-SameInclusionUrl $entryUrl $url)) { continue }
            if ((Test-SameRsaKey $entryKey $key) -and -not $kept) { $kept = $true; Say "  kept the existing entry $name - same URL, same key"; continue }
            $root.DeleteSubKeyTree($name)
            Say "  removed entry $name - same URL, a previous build's key"
        }
        if (-not $kept) {
            $name = [guid]::NewGuid().ToString()
            $entry = $root.CreateSubKey($name)
            try {
                $entry.SetValue('Url', $url, [Microsoft.Win32.RegistryValueKind]::String)
                $entry.SetValue('PublicKey', $key, [Microsoft.Win32.RegistryValueKind]::String)
            }
            finally { $entry.Close() }
            Say "  wrote HKCU\$InclusionKey\$name  Url=$url"
        }
    }
    finally { $root.Close() }

    # ---- 6. Make a failed load explain itself --------------------------------------------------
    [Environment]::SetEnvironmentVariable($LogAlertsName, '1', 'User')
    Set-Item -Path "Env:\$LogAlertsName" -Value '1'
    Say "  set $LogAlertsName=1 for this user"

    # ---- 6a. The record ------------------------------------------------------------------------
    $recordJson = ConvertTo-InstallRecordJson -Commit ([string]$Payload.commit) -Version ([string]$Payload.version) -InstalledUtc $installedUtc -InstallDir $installDir
    $recordDir = Split-Path -Parent $InstallRecordPath
    if ($recordDir -and -not (Test-Path -LiteralPath $recordDir)) { New-Item -ItemType Directory -Path $recordDir -Force | Out-Null }
    [System.IO.File]::WriteAllText($InstallRecordPath, $recordJson, (New-Object System.Text.UTF8Encoding($false)))
    $readBack = ConvertFrom-InstallRecord ([System.IO.File]::ReadAllText($InstallRecordPath))
    if ($readBack.Problem) { throw "The install record just written to $InstallRecordPath does not read back: $($readBack.Problem)." }
    Say ("  recorded in {0}: commit {1}, installed {2:o}" -f $InstallRecordPath, $readBack.Record.Commit, $readBack.Record.InstalledUtc)

    $interactionAfter = Get-InteractionFacts
    $changes = @(Compare-InteractionFacts $interactionBefore $interactionAfter)

    Say ''
    Say '== Verify =='
    $facts = Get-AddInFacts -RequireFresh $false -StartedUtc ([DateTime]::MinValue) -Payload $Payload -InstallRecord $readBack.Record
    $facts.ForbidReady = $true
    foreach ($c in $changes) {
        $facts.Notes += "INDEX EXCLUSION STATE CHANGED during the install: $c. Nothing this phase runs writes there, and no Outlook started. Re-run Set-OutlookIndexingDisabled.ps1 -Verify and find what did."
    }
    if ($changes.Count -eq 0) { Say '  index exclusion state: UNCHANGED by the install (policy value and mapi crawl rules identical before and after)' }
    Write-TestReadBlock $facts.Tuning
    $result = Get-AddInVerdict $facts
    Write-Verdict $result
    if ($result.ExitCode -eq 2) {
        Say ''
        Say 'NEXT - NOT elevated, Outlook closed, and on the unindexed guest only once its index exclusion is in'
        Say 'place (Testbed/README.md section 1, step 7b):'
        Say "    $(Get-PhaseCommand -Phase $PhaseFirstRun)"
    }
    exit $result.ExitCode
}

# -Phase FirstRun: NOT elevated; installs nothing, writes nothing, starts Outlook once and proves it.
function Invoke-FirstRunPhase {
    param($Payload)
    Assert-PhaseElevation -ForPhase $PhaseFirstRun
    Assert-Bitness
    Assert-InteractiveSession -ForPhase $PhaseFirstRun
    Assert-OutlookClosed -ForPhase $PhaseFirstRun

    # ---- 6b. Preflight: is a start safe, and worth making? -------------------------------------
    Say ''
    Say '== Before the start: installed, registered, trusted, and the payload''s build? =='
    $pre = Get-AddInFacts -RequireFresh $false -StartedUtc ([DateTime]::MinValue) -Payload $Payload
    $stop = Get-FirstRunPreflight $pre
    if ($null -ne $stop) {
        Write-Verdict $stop
        exit $stop.ExitCode
    }
    Say '  yes - starting Outlook'

    $interactionBefore = Get-InteractionFacts
    Say ("  index exclusion before: PreventIndexingOutlook={0} mapi rules=[{1}]" -f $interactionBefore.PreventIndexingOutlook, ($interactionBefore.MapiRules -join '; '))

    # ---- 7. The first run ----------------------------------------------------------------------
    Say ''
    Say '== First run: Outlook started once, NOT elevated, headless, in a watchdogged child job =='
    $windowsBefore = Get-SessionWindows
    $startedUtc = [DateTime]::UtcNow
    $job = Start-Job -ScriptBlock $FirstRunJob -ArgumentList $startedUtc.Ticks, $FirstRunTimeoutSeconds, $TuningKey, $McpKey, $AddinName
    $done = Wait-Job -Job $job -Timeout ($FirstRunTimeoutSeconds + 120)
    $first = $null
    if ($done) {
        $first = Receive-Job -Job $job
        Remove-Job -Job $job -Force
    }
    else {
        Say "  *** NOTHING BACK WITHIN $($FirstRunTimeoutSeconds + 120) s - Outlook is most likely showing a MODAL DIALOG. On screen now:"
        $onScreen = @(Get-SessionWindows)
        foreach ($w in $onScreen) { Say "      $w" }
        if (($onScreen -join ' ').Contains('Customization Installer')) { Say '  That is the VSTO TRUST PROMPT (or its install-error box): the inclusion entry did not match what the runtime looked for.' }
        Stop-Job -Job $job; Remove-Job -Job $job -Force
        Say '  The job was stopped. Outlook was NOT touched - answer the dialog on the console, or restart the guest. Never taskkill it.'
    }

    $comProblems = @()
    if ($null -eq $first) { $comProblems += 'the first run did not finish; see above.' }
    else {
        $tuningAfter = "NEVER, in $FirstRunTimeoutSeconds s"
        if ($null -ne $first.TuningAfterSeconds) { $tuningAfter = "after $($first.TuningAfterSeconds) s" }
        $mcpAfter = "NEVER, in $($FirstRunTimeoutSeconds + 30) s"
        if ($null -ne $first.McpAfterSeconds) { $mcpAfter = "after $($first.McpAfterSeconds) s" }
        Say "  COM start $($first.ComSeconds) s; MAPI $($first.MapiInit)"
        Say "  tuning state (Tuning\LastReconcileUtc) written $tuningAfter; registration reconcile (Mcp\$McpReconcileValue) written $mcpAfter"
        Say "  COMAddIns('$AddinName').Connect = $($first.Connect); the add-in answered GetRestartNeeded() = $($first.RestartNeeded)"
        if ($first.Error) { $comProblems += "the first run reported: $($first.Error)" }
        if ($null -eq $first.TuningAfterSeconds) {
            $noTuning = "the add-in did not write HKCU\$TuningKey\LastReconcileUtc - its tuning reconcile's LAST write - within $FirstRunTimeoutSeconds s of Outlook starting."
            if ($null -ne $first.McpAfterSeconds) {
                $noTuning += " It DID start: its registration reconcile wrote HKCU\$McpKey\$McpReconcileValue $($first.McpAfterSeconds) s in. So its tuning reconcile started and did not reach its end - the 'tuning walk' line above says how far it got. A build from before Q128 (2026-10-03) stopped at the first write that threw, measured at a value under HKCU\Software\Policies, which only an elevated token may write (Docs/live-tier-on-the-vm.md section 2.3); since Q128 the reconcile skips such a value and always writes LastReconcileUtc."
            }
            $comProblems += $noTuning
        }
        if ($first.Connect -ne $true) { $comProblems += "Outlook reports the add-in as NOT connected." }
        if (-not $first.AutomationAnswered) { $comProblems += 'the add-in did not answer a call into it (COMAddIn.Object.GetRestartNeeded), so it is not demonstrably running.' }
        # The point of this phase: the Outlook it started is NOT elevated. Read, not assumed.
        if ($first.TokenError) { $comProblems += "the started OUTLOOK.EXE's token could not be read ($($first.TokenError)), so this start cannot be shown NOT elevated." }
        $levels = @($first.OutlookElevation | Where-Object { $_ })
        if ($levels.Count -eq 0 -and $null -ne $first.ComSeconds -and -not $first.TokenError) { $comProblems += 'no OUTLOOK.EXE was found to read the token of while the job held Outlook, so this start cannot be shown NOT elevated.' }
        foreach ($l in $levels) {
            $parts = ([string]$l).Split('=')
            if ($parts[1] -eq '0') { Say "  OUTLOOK.EXE pid $($parts[0]): token NOT elevated" }
            elseif ($parts[1] -eq '1') { $comProblems += "OUTLOOK.EXE pid $($parts[0]) is ELEVATED: an elevated Outlook never feeds Windows Search, which is what this phase exists to avoid. Is UAC off (EnableLUA), or was an elevated Outlook already starting?" }
            else { $comProblems += "the token of OUTLOOK.EXE pid $($parts[0]) could not be read, so this start cannot be shown NOT elevated." }
        }
    }
    $still = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($still.Count -gt 0) { Say "  Outlook is still running (pid $(($still | ForEach-Object { $_.Id }) -join ', ')), headless. Left as it is." }
    else { Say '  Outlook closed by itself once the last reference was released - gracefully; nothing quit or killed it.' }
    # A window that appeared during the run and is still up is a dialog somebody has to answer -
    # and on an unattended guest, the next step's hang.
    $newWindows = @(Get-SessionWindows | Where-Object { $windowsBefore -notcontains $_ })
    foreach ($w in $newWindows) { $comProblems += "a window appeared during the first run and is still on screen: $w - answer it on the console before the next step." }

    $interactionAfter = Get-InteractionFacts
    $changes = @(Compare-InteractionFacts $interactionBefore $interactionAfter)

    Say ''
    Say '== Verify =='
    $facts = Get-AddInFacts -RequireFresh $true -StartedUtc $startedUtc -Payload $Payload
    $facts.ComProblems = $comProblems
    foreach ($c in $changes) {
        $facts.Notes += "INDEX EXCLUSION STATE CHANGED during this run: $c. The add-in writes nothing there, so this is Outlook's own start. Re-run Set-OutlookIndexingDisabled.ps1 -Verify before trusting this guest as unindexed."
    }
    if ($changes.Count -eq 0) { Say '  index exclusion state: UNCHANGED by this run (policy value and mapi crawl rules identical before and after)' }
    Write-TestReadBlock $facts.Tuning
    $result = Get-AddInVerdict $facts
    Write-Verdict $result
    exit $result.ExitCode
}

# =============================================================================================
# MAIN
# =============================================================================================
Assert-TestbedGuestLocal

$mode = Resolve-RunMode -Execute $Execute.IsPresent -Verify $Verify.IsPresent -Phase $Phase
if ($mode.Refusal) { throw $mode.Refusal }
if ($mode.Mode -eq 'DryRun') {
    Write-PhasePlan -ForPhase $mode.Phase
    exit 0
}

# Defined AFTER the guard on purpose: this job opens an Outlook COM session, and check 9 of
# check-testbed-references.ps1 requires every write to come after the guest guard in file
# order. It is only ever started by -Phase FirstRun, so nothing about behaviour moved.
# The FIRST RUN, in a child job so a modal dialog becomes a reported timeout instead of this
# script's hang. Everything the job holds is released in its finally; it never quits Outlook.
# The job inherits FirstRun's Limited token, and COM starts Outlook with it - which the job then
# proves, reading the started OUTLOOK.EXE's token while it still holds a reference, so Outlook
# cannot have closed by itself first. (The reading is Start-OutlookUnelevated.ps1's, restated.)
$FirstRunJob = {
    param([long] $StartedTicks, [int] $TimeoutSeconds, [string] $TuningKey, [string] $McpKey, [string] $AddinName)
    $ErrorActionPreference = 'Stop'
    $started = New-Object DateTime($StartedTicks, [DateTimeKind]::Utc)
    $r = [ordered]@{ ComSeconds = $null; MapiInit = $null; TuningAfterSeconds = $null; McpAfterSeconds = $null; Connect = $null; RestartNeeded = $null; AutomationAnswered = $false; Error = $null; OutlookElevation = @(); TokenError = $null }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $app = $null; $ns = $null; $addins = $null; $addin = $null; $obj = $null
    function Get-Fresh([string] $key) {
        $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key, $false)
        if ($null -eq $k) { return $false }
        try {
            $s = $k.GetValue('LastReconcileUtc') -as [string]
            if (-not $s) { return $false }
            $t = [DateTime]::MinValue
            if (-not [DateTime]::TryParse($s, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$t)) { return $false }
            return ($t.ToUniversalTime() -ge $started.AddSeconds(-1))
        }
        finally { $k.Close() }
    }
    try {
        $app = New-Object -ComObject Outlook.Application
        $r.ComSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
        try { $ns = $app.GetNamespace('MAPI'); $null = $ns.GetDefaultFolder(6); $r.MapiInit = 'ok' }
        catch { $r.MapiInit = $_.Exception.Message }
        # The tuning state is what the tests read, so it gets the whole deadline. The registration
        # reconcile runs on a worker thread and is only reported, so it gets 30 s more at most. It
        # is watched DURING the tuning wait too: it is the add-in's other startup marker, so it
        # says whether a tuning state that never comes is an add-in that never started or one that
        # started and stopped part-way - and its time is only meaningful if it is read as it lands.
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while ([DateTime]::UtcNow -lt $deadline -and $null -eq $r.TuningAfterSeconds) {
            if ($null -eq $r.McpAfterSeconds -and (Get-Fresh $McpKey)) { $r.McpAfterSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1) }
            if (Get-Fresh $TuningKey) { $r.TuningAfterSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1) }
            else { Start-Sleep -Seconds 2 }
        }
        $mcpDeadline = [DateTime]::UtcNow.AddSeconds(30)
        while ([DateTime]::UtcNow -lt $mcpDeadline -and $null -eq $r.McpAfterSeconds) {
            if (Get-Fresh $McpKey) { $r.McpAfterSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1) }
            else { Start-Sleep -Seconds 2 }
        }
        try {
            $addins = $app.COMAddIns
            $addin = $addins.Item($AddinName)
            $r.Connect = [bool]$addin.Connect
            $obj = $addin.Object
            if ($null -ne $obj) { $r.RestartNeeded = [bool]$obj.GetRestartNeeded(); $r.AutomationAnswered = $true }
        }
        catch { $r.Error = 'COMAddIns: ' + $_.Exception.Message }
        try {
            if (-not ('OaiFirstRunToken' -as [type])) {
                Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class OaiFirstRunToken {
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint a, bool i, int p);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr p, uint a, out IntPtr t);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr t, int c, out int i, int l, out int r);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    public static int Elevation(int pid) { IntPtr p = OpenProcess(0x1000, false, pid); if (p == IntPtr.Zero) return -1; try { IntPtr t; if (!OpenProcessToken(p, 8, out t)) return -1; try { int e, r; if (!GetTokenInformation(t, 20, out e, 4, out r)) return -1; return e != 0 ? 1 : 0; } finally { CloseHandle(t); } } finally { CloseHandle(p); } }
}
'@
            }
            $r.OutlookElevation = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | ForEach-Object { '{0}={1}' -f $_.Id, [OaiFirstRunToken]::Elevation($_.Id) })
        }
        catch { $r.TokenError = $_.Exception.Message }
    }
    catch { $r.Error = $_.Exception.GetType().Name + ': ' + $_.Exception.Message }
    finally {
        foreach ($o in @($obj, $addin, $addins, $ns, $app)) {
            if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    [pscustomobject]$r
}

# -Verify -WithOutlook: attach to an Outlook ALREADY RUNNING in this session and ask the same
# two questions. Never starts one: GetActiveObject finds a running instance or throws.
$AttachJob = {
    param([string] $AddinName)
    $r = [ordered]@{ Connect = $null; RestartNeeded = $null; AutomationAnswered = $false; Error = $null }
    $app = $null; $addins = $null; $addin = $null; $obj = $null
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $addins = $app.COMAddIns
        $addin = $addins.Item($AddinName)
        $r.Connect = [bool]$addin.Connect
        $obj = $addin.Object
        if ($null -ne $obj) { $r.RestartNeeded = [bool]$obj.GetRestartNeeded(); $r.AutomationAnswered = $true }
    }
    catch { $r.Error = $_.Exception.GetType().Name + ': ' + $_.Exception.Message }
    finally {
        foreach ($o in @($obj, $addin, $addins, $app)) {
            if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    [pscustomobject]$r
}

$script:ResolvedOffice = Resolve-OfficeVersion
Say "Office hive: $script:ResolvedOffice"

$payload = $null
$payloadPath = Join-Path $PayloadRoot $ManifestFileName
if (Test-Path -LiteralPath $payloadPath) {
    $payload = Get-Content -LiteralPath $payloadPath -Raw | ConvertFrom-Json
    $shape = @(Test-PayloadManifestShape $payload)
    if ($shape.Count -gt 0) { throw ("The payload manifest $payloadPath is malformed: " + ($shape -join '; ')) }
    Say "Payload: commit $($payload.commit), version $($payload.version), built $($payload.builtUtc)"
}
elseif ($Execute) {
    throw "REFUSING: no $payloadPath. Build it on the host with Testbed/host/Publish-AddInPayload.ps1, copy AddIn.zip in with Testbed/host/Copy-ToGuest.ps1, and expand it to $PayloadRoot."
}
else {
    Say "No payload manifest at $payloadPath - the installed build cannot be tied to a commit."
}

# The two phases. Each refuses the other's run level, and each ends the script with its verdict.
if ($mode.Mode -eq $PhaseInstall) { Invoke-InstallPhase -Payload $payload }
if ($mode.Mode -eq $PhaseFirstRun) { Invoke-FirstRunPhase -Payload $payload }

# -Verify
Say ''
Say '== Verify =='
$installRecord = Read-InstallRecord
$facts = Get-AddInFacts -RequireFresh $false -StartedUtc ([DateTime]::MinValue) -Payload $payload -InstallRecord $installRecord.Record -InstallRecordProblem $installRecord.Problem
if (-not $installRecord.Record -and -not $installRecord.Problem) {
    $facts.Notes += "no install record at $InstallRecordPath - an install made before -Phase existed, or none at all - so whether the installed build has run since it was installed is not asked; the state is read as before the split."
}
$ix = Get-InteractionFacts
Say ("  index exclusion now: PreventIndexingOutlook={0} mapi rules=[{1}]  (read only; the add-in writes neither)" -f $ix.PreventIndexingOutlook, ($ix.MapiRules -join '; '))
if ($WithOutlook) {
    Assert-Bitness
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq [System.Diagnostics.Process]::GetCurrentProcess().SessionId })
    if ($running.Count -eq 0) {
        $facts.Notes += '-WithOutlook: no Outlook is running in this session, and a verify never starts one. The COM half was not checked.'
    }
    else {
        $job = Start-Job -ScriptBlock $AttachJob -ArgumentList $AddinName
        if (Wait-Job -Job $job -Timeout 90) {
            $att = Receive-Job -Job $job
            Say "  running Outlook: Connect=$($att.Connect), the add-in answered GetRestartNeeded() = $($att.RestartNeeded)"
            if ($att.Error) { $facts.ComProblems += "-WithOutlook: $($att.Error)" }
            elseif ($att.Connect -ne $true -or -not $att.AutomationAnswered) { $facts.ComProblems += 'the running Outlook does not have the add-in connected and answering.' }
        }
        else {
            Stop-Job -Job $job
            $facts.ComProblems += '-WithOutlook: the running Outlook did not answer within 90 s - it may be showing a dialog.'
        }
        Remove-Job -Job $job -Force
    }
}
Write-TestReadBlock $facts.Tuning
$result = Get-AddInVerdict $facts
Write-Verdict $result
exit $result.ExitCode
