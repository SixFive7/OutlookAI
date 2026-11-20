#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the live tier on a test guest end to end in ONE call, and writes summary.txt and
    summary.json: lease, checkpoint, staging, hub rebuild, the suite at RunLevel Limited with the
    per-run opt-in, results, resting checkpoint, save, release.

.DESCRIPTION
    The guest counterpart of Invoke-TestsOnBuildVm.ps1, written 2026-10-03 after the first guest
    runs were driven by hand from scratch scripts, hundreds of calls a run. One run is:

      0. GATE    - waits, holding nothing, while the testbed's pause switch is on or the host's
                   processors were 70 % busy or more over the last 30 seconds (Q195 (e),
                   Testbed/host/TestbedLoad.ps1): the build below is the run's first load on this
                   host. It says what it waits for, and since when, once a minute.
      1. BUILD   - the commit, never the working tree: a detached `git worktree` of it, from which
                   Publish-GuestPayload.ps1 builds the server and the tools and
                   Publish-LiveTierPayload.ps1 takes the suite's source (`git archive`) and the
                   offline feed; New-LiveTestSettings.ps1 renders the guest's settings from that
                   commit's Testbed/testbed.json. Uncommitted changes are not in the run, and it says
                   how many it left out. Nothing touches the guest yet.
      2. LEASE   - waits until nobody else holds a live lease on the guest (Set-TestbedLease.ps1),
                   NEVER taking one over, then takes it, and renews it before every long step. The
                   guests are shared between agents, and the idle-saver saves any running guest that
                   holds no live lease. It waits asking Hyper-V nothing (2026-11-18): a ticket in the
                   testbed's queue, first come first served, the lease files to tell a free guest,
                   and a sleep until a lease, a ticket or the gate's switch changes or 30 s pass
                   (TestbedPool.ps1). Every turn of that wait looks at the gate again before it
                   tries a guest, and the lease is taken only under the gate's claim, the VMs other
                   runs took in the last two minutes counted at their size (TestbedLoad.ps1).
      3. STAGE   - restores the start checkpoint, copies the payloads, the settings and the commit's
                   Testbed/guest scripts in, swaps the server, tools, source and feed, and proves the
                   suite TEST-READY (Install-DotnetSdk.ps1 -Execute, then -Verify).
      4. PREPARE - FROM A FROZEN CHECKPOINT (the default since Q130 (a)): Set-GuestClockFrozen.ps1
                   -Verify - the run stops unless it says FROZEN - then Outlook started NOT elevated on
                   the tier profile (Start-OutlookUnelevated.ps1), the state step 9a hands the suite
                   over in; no restart and no hub rebuild, which would leave the frozen instant the
                   guest's data was built for.
                   From any other checkpoint (time sync on): Restart-Guest.ps1 -Execute (graceful, or
                   it refuses), then Reset-HubPopulation.ps1 -Execute (Testbed/README.md step 9a) at
                   RunLevel Limited: every Outlook start NOT elevated (D77).
      5. RUN     - the suite through Register-InteractiveTask.ps1 -RunLevel Limited, with
                   OUTLOOKAI_LIVE_OPT_IN set to the guest's computer name INSIDE the task's script,
                   for this run only, and the guest's derived filter (T2/LiveRunFilters.cs).
      6. FETCH   - the TRX file, the console capture and VSTest's logs, the guest's Application log
                   entries for an OUTLOOK.EXE crash during the run, and every crash dump Windows Error
                   Reporting wrote (dumps\<guest>\) - armed in STAGE by guest\Set-OutlookCrashDumps.ps1,
                   full dumps of OUTLOOK.EXE and WINWORD.EXE, since 2026-10-04 - and, since 2026-11-18,
                   WER's own report of each such crash: its event 1001 and the files of its report
                   folder (wer\<guest>\), which say which way a crash went that left no dump.
      7. REST    - optionally a checkpoint of a GREEN run (-GreenCheckpoint), then the resting
                   checkpoint restored, the settings staged on it, the guest SAVED and the lease
                   released - on every path, failures included, once the guest was touched.
      8. REPORT  - summary.txt and summary.json in <RepoPath>\.work\guest-live-runs\<run id>\.

    THE EXCHANGE VM (2026-10-08). OutlookAI-Exchange runs here too, with its own STAGE, PREPARE and
    REST and two things a real mailbox adds. ITS STATE IS CARRIED FORWARD: no checkpoint is restored
    before or after a run unless -Checkpoint names one, because its mailboxes live on the Exchange
    server, which no checkpoint restores - a restored .ost is older than the server, and Outlook refuses
    it at its next start ("Outlook is using an old copy of your Outlook data file", measured on the
    second write run). So STAGE resumes the VM as the last run left it, PREPARE starts Outlook NOT
    elevated if it is not running - after Defender's signatures are brought up to date, so that it
    starts on current ones - and runs Invoke-ExchangeSignIn.ps1 -Mode Preflight: Defender current, so
    Outlook's Object Model Guard does not stall the run, and every recorded account still SIGNED-IN; REST
    saves it as it is. With -Checkpoint, the checkpoint is restored, the guest restarted gracefully and
    its .ost files set aside (guest\Reset-OutlookOfflineStore.ps1) before Outlook starts and downloads
    the mailboxes afresh. STAGE leaves a running Outlook running (the swap replaces only the suite and
    its server, which Outlook does not use). Its filter is LiveRunFilters.ExchangeTestAccount, the write run's
    own, again since 2026-11-18: the VM's never-activated Office had passed its grace and Outlook created no
    item, until decision 206 (a) re-armed the grace once with Office's own OSPPREARM.EXE (Docs/live-tier-on-the-vm.md
    section 4.4a). Past the grace - next on 2026-12-18 - every write test fails at its first item again; -FilterSuffix
    '&Writes=Nothing' narrows a run to LiveRunFilters.ExchangeTestAccountWritesNothing until it is re-armed. And every
    run gets a RUN MARKER - 'oair' and
    the UTC start to the second - recorded in the host's ledger, C:\OutlookAI-Testbed\exchange-runs\
    ledger.jsonl, BEFORE the suite starts, and handed to the run in OUTLOOKAI_LIVE_RUN_MARKER with the
    markers of earlier runs that never reported a clean sweep in OUTLOOKAI_LIVE_PRIOR_RUN_MARKERS: every
    item the run makes carries its marker, and its sweep takes back its own items and those earlier
    runs' by marker - never by a subject pattern (T2/LiveRunMarker.cs, Docs/live-tier-on-the-vm.md
    section 4.4). The ledger lives on the host because a checkpoint restore wipes the guest and leaves
    the Exchange server as it was.

    WHAT IT ENFORCES, so a caller cannot forget it (AGENTS.md "Mailbox Safety", "Test VMs stay
    saved unless in use"):
      * Only the two test guests and the Exchange VM. The build VM, the retired OutlookAI-TestVM and
        anything else are refused, and nothing runs on this workstation: the suite runs inside the guest.
      * The opt-in exists only inside the guest task of this run; it is never set here.
      * A frozen guest's clock stays where its checkpoint put it (Q130 (a), AGENTS.md): a run from a
        checkpoint whose time sync is off never restarts the guest or rebuilds its hub, the clock is
        verified before the suite, and such a run cannot make its green checkpoint the resting one.
      * Outlook is never killed. The guest restarts only through Restart-Guest.ps1, which quits it
        gracefully or refuses; otherwise only checkpoints are restored.
      * The credential only through Get-GuestCredential.ps1, and it is never printed.
      * The suite's own safety proofs are VERDICTS. A sweep that reports a tagged artifact left, a
        count-tripwire census that reports a failure, or an OUTLOOK.EXE crash fails the run whatever
        its tests did; so does `dotnet test` exiting non-zero over a TRX file in which every test
        passed (a fixture or the census failed). On a full run - no -Filter or -FilterSuffix - a
        sweep or census that never reported fails it too: zero artifacts was not proven. A narrowed
        run that selects neither says NOT PROVEN in its summary.
      * The summary names tests, stores and counts, never a subject or a body.
      * The VM rule: a live lease while the guest is in use; saved and released at the end.

    A SIGNATURE SETTING BEFORE OUTLOOK STARTS (-SignatureSettings, 2026-11-18, Q200 (a) and Q201). An
    administrator's signature policy, a default deployed in Outlook's general mail settings and an
    account's own defaults are read by Outlook when it starts, and the suite runs NOT elevated, where
    HKCU\Software\Policies cannot be written. So with -SignatureSettings GroupPolicy, Preference,
    PreferenceNoOwnNew or NoSetting (the control), PREPARE runs guest\Set-SignatureSettings.ps1 -Execute
    elevated in session 0
    between the frozen-clock guard and Outlook's start: it makes the setting, its test signatures and the
    hub account's own defaults, and says SET or the run stops. The run is narrowed to the test that
    measures what Outlook itself then inserts (T2/LiveSignatureSettingsTests.SetBeforeOutlookStarted_
    WhatOutlookInserts...), told the scenario in OUTLOOKAI_LIVE_SIGNATURE_SETTINGS. With -SignatureDialog
    as well, the guest task reads Outlook's own Signatures dialog before the suite instead -
    guest\Measure-SignatureDialog.ps1, at Outlook's integrity level, closed as Cancel does - and the run
    is narrowed to the dialog's test (..._TheSignaturesDialog...), told the reading in
    OUTLOOKAI_LIVE_SIGNATURE_DIALOG. The two never share a run: opening the dialog under a setting
    rewrites the accounts' own defaults (measured 2026-11-18), which the insert test must not meet. Only
    from a frozen start, never on the Exchange VM and never with -GreenCheckpoint or -Filter: the
    checkpoint restored at REST is what takes the setting away.

    EXIT CODES: 0 PASS; 1 FAIL - a test, a safety verdict or Outlook failed, or the filter selected
    nothing; 2 BUILD - the commit did not build, stage or become TEST-READY; 3 INFRA - not tested
    (lease, VM, guest calls, time); 4 REFUSED - the request itself.

    Run it in the background or with a timeout of an hour or more: a full run of the indexed guest
    takes about 40 minutes. A caller stopped part-way leaves the guest leased until the lease
    expires; the next run restores its checkpoint first anyway.

.PARAMETER VMName
    The guest: OutlookAI-Indexed or OutlookAI-Unindexed, or one of their clones (TestbedPool.ps1), or
    the Exchange VM, OutlookAI-Exchange, which is never cloned.
    This or -Kind, never both and never neither: no default picks a guest.

.PARAMETER Kind
    Indexed or Unindexed: whichever guest of that kind is free - the source first, then its clones,
    a clone only while the host has its memory and the reserve to spare (TestbedPool.ps1) - taken
    once the payloads are built, when the run needs it. From that guest's frozen checkpoint, resting
    on it again: -Checkpoint, -RestingCheckpoint, -GreenCheckpoint and -RestOnGreen name one
    machine's checkpoints and need -VMName.

.PARAMETER Ref
    The commit or branch to test. Default HEAD of -RepoPath.

.PARAMETER RepoPath
    The checkout to take the commit from. Default: the checkout this script is in.

.PARAMETER Filter
    Replaces the derived filter. It must start with Category=Live. Prefer -FilterSuffix.

.PARAMETER FilterSuffix
    Narrows the derived filter, e.g. '&FullyQualifiedName~LiveDisconnectRecoveryTests'. Must start
    with '&'.

.PARAMETER Checkpoint
    The checkpoint the run starts from. Default the guest's resting checkpoint.

.PARAMETER RestingCheckpoint
    The checkpoint the guest is left on. Default the guest's resting checkpoint.

.PARAMETER GreenCheckpoint
    When the run is GREEN, a checkpoint of that name is taken of the guest as the run left it, before
    the guest is put to rest. Not taken when E: has under 60 GB free.

.PARAMETER RestOnGreen
    With -GreenCheckpoint: when it was taken, rest the guest on it instead of -RestingCheckpoint.

.PARAMETER SkipHubReset
    Skips step 9a. For a narrowed run only: the frontier test fails on a hub nobody rebuilt.

.PARAMETER WithAddIn
    Installs the tested commit's add-in on the guest before Outlook starts, over the one its checkpoint
    holds - for the tests that drive the add-in itself (T2 LivePaneTests, Q168 (b)); every other test
    reads only the add-in's registry state, which any build writes the same. The add-in is built here by
    Publish-AddInPayload.ps1, with its Q81 guards, installed as the guest's default task, elevated, and
    started once NOT elevated, which must end in ADDIN-READY with Outlook closed again. A frozen-clock
    guest only. Adds about three minutes.

.PARAMETER VstoRuntimePath
    With -WithAddIn: the staged vstor_redist.exe (Testbed/MEDIA.md). Default: .work\media\vstor_redist.exe
    of the main checkout, else its Redist\vstor_redist.exe.

.PARAMETER ProveCrashDumps
    After arming the crash dumps, proves them once in the run's own conditions - session 1, NOT
    elevated, as Outlook runs: guest\Set-OutlookCrashDumps.ps1 -Prove crashes a throwaway probe
    program (never Outlook) and waits for its full dump. Recorded in the summary; a failed proof does
    not stop the run. Once per guest is enough - runbook section 4.6.

.PARAMETER HuntFirstChance
    THE HUNT for the crash of OUTLOOK.EXE in mso40uiwin32client.dll that left no dump twice on 2026-11-18
    (runbook section 4.15; the coordinator's decision of that day). A debugger - guest\Watch-OutlookFirstChance.ps1,
    THIS checkout's copy - is attached to every OUTLOOK.EXE of the suite's session and writes one full dump at the
    first first-chance access violation in that module, and one at any second-chance exception; and the suite repeats
    the tests the two crashes struck in, each leg in a dotnet test of its own, until -HuntRounds rounds, -HuntMinutes
    minutes or a dump. Off by default: no other run ever has a debugger attached, and this one is refused on the
    Exchange VM, with a filter, or with a signature step. The dumps land in dumps\<guest>\ beside the WER evidence,
    the debugger's log in results\<guest>\first-chance.log.

.PARAMETER HuntRounds
    With -HuntFirstChance: at most this many rounds, every leg once each (default 20).

.PARAMETER HuntMinutes
    With -HuntFirstChance: no round starts after this many minutes of the suite (default 60).

.PARAMETER HuntVariant
    With -HuntFirstChance: what each round runs (the coordinator's direction of 2026-11-18, option (b); runbook
    section 4.15). Legs, the default, runs the tests the crashes struck in, as before. LegsOtherFolder runs them too,
    each round after a step that moves the main window to the Calendar. Every other value runs one leg a
    round, T2.LiveHuntDerivedDraftTests, which makes derived drafts in a loop for -HuntLegMinutes or until Outlook ends,
    one attempt line a draft: ProductLoop, the control - the product's reply_draft and forward_draft in the windowed
    Outlook; OtherFolder - the same with the main window on the Calendar; NoWindow - the same in an Outlook COM started,
    with no window; NoFormPage - the product's direct route made by the object model without the "Message" form page;
    BareCreators - Reply() or Forward() and the result released unsaved.

.PARAMETER HuntLegMinutes
    With a -HuntVariant other than Legs: the minutes of one leg, one dotnet test of the loop (default 10, at most
    -HuntMinutes).

.PARAMETER HuntShowEach
    With a -HuntVariant other than Legs: every draft of the loop comes after a compose window shown and closed - a
    reply or forward the object model made, opened in an inspector and promoted the product's own way (parked,
    activated, hidden), closed discarding - as the typed-text tests read every draft before the product makes the
    next. The loops without it, which make no window, never crashed (2026-11-18); every hunted crash struck after such
    windows.

.PARAMETER HuntTypeEach
    -HuntShowEach with the window displayed on screen and typed into at Outlook's insertion point - Display and Word's
    Selection.TypeText, as the typed-text tests do - before it is closed discarding. The promoted windows of
    -HuntShowEach never crashed (2026-11-18); the typed-text tests, which display and type, did.

.PARAMETER HuntPromoteAfter
    With a -HuntVariant whose loop makes the product's drafts: each draft, as soon as it is made, is re-opened by its
    EntryID, promoted the product's own way (parked, activated, hidden) and closed discarding - what the typed-text
    tests do to read every draft of the product's. With -HuntTypeEach it is their sequence in a loop.

.PARAMETER HuntReleaseAfterClose
    With -HuntTypeEach: the typed window's Word document is kept until the window is closed and released after it, as
    the typed-text tests release theirs and as the product's own compose does. Without it the loop releases the
    document first and closes the window after, as the product's promotion read does.

.PARAMETER HuntAblate
    With -HuntVariant Legs or LegsOtherFolder: the typed-text tests leave one of their steps out (the coordinator's
    direction of 2026-11-18, after the variants) - new: no new mail beside the reply and the forward; layout: Word's
    paragraphs never walked, in the typed window or in the product's drafts re-opened after; signature: the product's
    second draft of each kind without the test signature, which is never made; source: the tests of a leg share one
    source mail. T2.HuntPlan.ParseAblation reads it.

.PARAMETER HuntTraceClr
    With -HuntFirstChance: the .NET runtime's own exception events are traced on the guest for the suite - an ETW
    session for Microsoft-Windows-DotNETRuntime's Exception keyword, started elevated before the suite, and at its end
    the runtime's rundown, which names every method each .NET process has compiled - and tracerpt's XML of both is
    fetched into results\<guest>\clr-events.xml. Windows' own logman and tracerpt; nothing is installed.

.PARAMETER HuntWithoutAddIn
    With -HuntFirstChance: OutlookAI's own add-in, which every Outlook guest's frozen checkpoint holds installed and
    loading (LoadBehavior 3), is turned off for the run - LoadBehavior 0, written before Outlook starts, so no Outlook
    of the run loads it; the restore at REST brings the checkpoint's value back. The add-in runs inside Outlook: it
    hooks every inspector and puts its task pane on each one that is activated. A frozen-clock Outlook guest only, and
    refused with -WithAddIn. Each round of the suite says whether OUTLOOK.EXE has the .NET runtime the add-in needs.

.PARAMETER SignatureSettings
    GroupPolicy (Q200 (a)), Preference (Q201), PreferenceNoOwnNew (Q201, the hub's own default for
    replies only) or NoSetting (the control - the hub's own defaults and no setting): makes that
    signature setting on the guest before Outlook starts, through
    guest\Set-SignatureSettings.ps1 run elevated, reads Outlook's Signatures dialog before the suite, and
    narrows the run to the tests that measure what Outlook then does. A frozen start of an Outlook guest
    only; refused on the Exchange VM, with -GreenCheckpoint and with -Filter.

.PARAMETER SignatureDialog
    With -SignatureSettings: read Outlook's own Signatures dialog before the suite, and run only the
    dialog's test (Q156 (d), 2026-11-18). Without it the dialog stays unopened and the run measures what
    Outlook inserts.

.PARAMETER SignatureValueKind
    With -SignatureSettings: how its two registry values are written - ExpandString (REG_EXPAND_SZ, the
    default; what Outlook was reported to honour) or String (REG_SZ).

.PARAMETER LeaseWaitMinutes
    How long the run waits at the gate before its build, and again how long it waits for a guest -
    for leases others hold and for the gate - before it gives up as INFRA, not tested. Default 240.

.PARAMETER SelfTest
    Checks the pure parts - the guest table, the filter, the lease decision, the TRX and console
    readers, the verdict table, summary.json's shape - and the constants against the files they
    copy. No Hyper-V, no git, no guest. Runs on the build VM with every other -SelfTest.

.EXAMPLE
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Unindexed
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Indexed d4e31fe -GreenCheckpoint CP-19C-LIVE-GREEN
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Indexed -FilterSuffix '&FullyQualifiedName~LiveShowMeTests' -SkipHubReset
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -SelfTest
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')] [string] $VMName,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('Indexed', 'Unindexed')] [string] $Kind,
    [Parameter(ParameterSetName = 'Run', Position = 0)] [string] $Ref = 'HEAD',
    [Parameter(ParameterSetName = 'Run')] [string] $RepoPath,
    [Parameter(ParameterSetName = 'Run')] [string] $Filter = '',
    [Parameter(ParameterSetName = 'Run')] [string] $FilterSuffix = '',
    [Parameter(ParameterSetName = 'Run')] [string] $Checkpoint = '',
    [Parameter(ParameterSetName = 'Run')] [string] $RestingCheckpoint = '',
    [Parameter(ParameterSetName = 'Run')] [string] $GreenCheckpoint = '',
    [Parameter(ParameterSetName = 'Run')] [switch] $RestOnGreen,
    [Parameter(ParameterSetName = 'Run')] [switch] $SkipHubReset,
    [Parameter(ParameterSetName = 'Run')] [switch] $ProveCrashDumps,
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntFirstChance,
    [Parameter(ParameterSetName = 'Run')] [int] $HuntRounds = 20,
    [Parameter(ParameterSetName = 'Run')] [int] $HuntMinutes = 60,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('Legs', 'LegsOtherFolder', 'ProductLoop', 'OtherFolder', 'NoWindow', 'NoFormPage', 'BareCreators', 'PersonTyping')] [string] $HuntVariant = 'Legs',
    [Parameter(ParameterSetName = 'Run')] [int] $HuntLegMinutes = 10,
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntWithoutAddIn,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('', 'new', 'layout', 'signature', 'source')] [string] $HuntAblate = '',
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntTraceClr,
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntPromoteAfter,
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntReleaseAfterClose,
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntShowEach,
    [Parameter(ParameterSetName = 'Run')] [switch] $HuntTypeEach,
    [Parameter(ParameterSetName = 'Run')] [switch] $WithAddIn,
    [Parameter(ParameterSetName = 'Run')] [string] $VstoRuntimePath,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('GroupPolicy', 'Preference', 'PreferenceNoOwnNew', 'NoSetting')] [string] $SignatureSettings,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('ExpandString', 'String')] [string] $SignatureValueKind = 'ExpandString',
    [Parameter(ParameterSetName = 'Run')] [switch] $SignatureDialog,
    [Parameter(ParameterSetName = 'Run')] [string] $ResultsRoot,
    [Parameter(ParameterSetName = 'Run')] [string] $CredentialRepoRoot,
    [Parameter(ParameterSetName = 'Run')] [string] $SdkInstallerPath,
    [Parameter(ParameterSetName = 'Run')] [int] $LeaseWaitMinutes = 240,
    [Parameter(ParameterSetName = 'Run')] [int] $RunTimeoutMinutes = 120,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# ---------------------------------------------------------------------------------------------
# The test guests and their layout, as constants. -SelfTest holds them to the files they copy.
# ---------------------------------------------------------------------------------------------
$Guests = [ordered]@{
    # The resting checkpoints are the FROZEN ones (Q130 (a), decided 2026-10-03): testbed.json's
    # frozenClocks records them, and -SelfTest holds these two names to that record.
    'OutlookAI-Indexed'   = [ordered]@{ ComputerName = 'OAI-INDEXED'; Indexed = $true; Checkpoint = 'CP-24C-FROZEN-CLOCK'; Exchange = $false }
    'OutlookAI-Unindexed' = [ordered]@{ ComputerName = 'OAI-UNINDEXED'; Indexed = $false; Checkpoint = 'CP-18B-FROZEN-CLOCK'; Exchange = $false }
    # The Exchange VM (2026-10-08): never frozen, never cloned; its resting checkpoint is the newest
    # testbed.json's exchangeVm.checkpoints records, and -SelfTest holds the two equal.
    'OutlookAI-Exchange'  = [ordered]@{ ComputerName = 'OAI-EXCHANGE'; Indexed = $true; Checkpoint = 'CP-10-SDK11'; Exchange = $true }
}
# AND THEIR CLONES (2026-11-17): TestbedPool.ps1's indexed and unindexed pools. A clone is its
# source copied at the source's frozen checkpoint - New-TestbedClone.ps1 - and rests on a checkpoint
# of the same name, with the source's computer name, stores and settings, so its facts are its
# source's: the same opt-in, the same filter, the same rendered settings, the same frozen record.
. (Join-Path $PSScriptRoot 'TestbedPool.ps1')
# THE GATE (Q195 (e)): the pause switch and the load limit, before the build and before a guest is taken.
. (Join-Path $PSScriptRoot 'TestbedLoad.ps1')
# Hyper-V's Time Synchronization integration service, by component id (Set-GuestClockFrozen.ps1).
$TimeSyncComponentId = '2497F4DE-E9FA-4204-80E4-4B75C46419C0'
# The profile the suite runs on, which a frozen start opens NOT elevated before the suite, as step 9a
# leaves it (Reset-HubPopulation.ps1's default; -SelfTest holds the two equal).
$TierProfileName = 'OutlookAI-Tier'
$GuestRoot = 'C:\OutlookAI-Q5'
$GuestPayloadDir = 'C:\OutlookAI-Q5\g4-payload'
$GuestRunsRoot = 'C:\OutlookAI-Q5\live-runs'
$GuestSettingsPath = 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json'
# Where the guest's Windows Error Reporting writes a FULL dump of a crashing OUTLOOK.EXE (2026-10-04):
# armed at every run by guest\Set-OutlookCrashDumps.ps1 - THIS checkout's copy, whatever commit is
# tested, because the frozen checkpoints predate it - and fetched into dumps\ before the guest rests.
$GuestDumpDir = 'C:\OutlookAI-Q5\crash-dumps'
$CrashDumpScript = 'Set-OutlookCrashDumps.ps1'
# THE HUNT (-HuntFirstChance; decided 2026-11-18, runbook section 4.15): the debugger, THIS checkout's copy, the module
# whose first first-chance access violation it dumps, and the legs of each round - the tests the two crashes of that
# day struck in, as suffixes of the guest's derived filter. Each leg runs in a dotnet test of its own: the Q149 gate
# runs the typed-text tests only when their collection is selected alone. Stream U's crash struck in the reply with a
# signature (its run's TRX: the crash at 17:07:41.9, inside that test's 17:07:35-17:08:00), the forward beside it in
# the same collection failed after it; K2's in the typed-text test's reply.
$FirstChanceScript = 'Watch-OutlookFirstChance.ps1'
$HuntModule = 'mso40uiwin32client.dll'
$HuntLegs = [ordered]@{
    typed  = '&FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveQ149TypedTextTests'
    drafts = '&(FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveSignatureTests.ReplyDraft_WithSignature_InsertsAboveQuote_ThenUpdateSwapsItForAnother_ThenNoneRemovesIt|FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveDraftOptionsTests.ForwardDraft_CcBccAppend_AndSubjectOverrideKeepsTheForwardedContent)'
}
# The hunt's variants (-HuntVariant; the coordinator's direction of 2026-11-18, option (b)): every one but Legs runs one
# leg a round, the loop of T2.LiveHuntDerivedDraftTests, told its variant and its minutes by these two variables - which
# no other run sets, so the loop makes nothing anywhere else (T2.HuntPlan; T1.HuntPlanTests holds both sides to one list).
$HuntVariantVariable = 'OUTLOOKAI_HUNT_VARIANT'
$HuntLegMinutesVariable = 'OUTLOOKAI_HUNT_LEG_MINUTES'
$HuntShowEachVariable = 'OUTLOOKAI_HUNT_SHOW_EACH'
$HuntAblateVariable = 'OUTLOOKAI_HUNT_ABLATE'
$HuntPromoteAfterVariable = 'OUTLOOKAI_HUNT_PROMOTE_AFTER'
$HuntReleaseAfterCloseVariable = 'OUTLOOKAI_HUNT_RELEASE_AFTER_CLOSE'
# -HuntTraceClr: the ETW session's name on the guest (its rundown session's is the same with Rundown after it).
$ClrTraceSession = 'OutlookAIHuntClr'
$HuntLoopLeg = '&FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveHuntDerivedDraftTests'
# -HuntWithoutAddIn: the key Install-OutlookAIAddIn.ps1 registers the add-in under, for the user Outlook runs as.
$AddInRegistryKey = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookAI'
# The -SignatureSettings step (Q200 (a), Q201): the guest script that makes the setting before Outlook starts, the
# variable that tells the suite which it made, and the one test such a run is narrowed to.
$SignatureSettingsScript = 'Set-SignatureSettings.ps1'
$SignatureSettingsVariable = 'OUTLOOKAI_LIVE_SIGNATURE_SETTINGS'
$SignatureSettingsTest = 'LiveSignatureSettingsTests.SetBeforeOutlookStarted'
# The two tests such a run is narrowed to, one a run: what Outlook inserts, or - with -SignatureDialog - its dialog.
$SignatureInsertsTest = "$($SignatureSettingsTest)_WhatOutlookInserts"
$SignatureDialogTest = "$($SignatureSettingsTest)_TheSignaturesDialog"
# Its reading of Outlook's own Signatures dialog, in the guest task before the suite (Q156, 2026-11-18).
$SignatureDialogScript = 'Measure-SignatureDialog.ps1'
$SignatureDialogVariable = 'OUTLOOKAI_LIVE_SIGNATURE_DIALOG'
$SignatureDialogReading = 'signature-dialog.json'
# The SDK a run's commit installs when its guest lacks it, and the SDK its verification looks for:
# Install-DotnetSdk.ps1 -Execute and -Verify, the commit's own copy, told all three below, so that
# copy's defaults do not decide. The .NET 11 SDK since the code moved to .NET 11 (Q172,
# 2026-11-18): every checkpoint this restores holds 11.0.100 - the frozen ones beside 10.0.401,
# CP-10-SDK11 beside 10.0.404 - so a run installs nothing. A branch still on .NET 10 keeps the
# runner that names 10.0.401 and the frozen checkpoints before these (Testbed/MEDIA.md). Hashed on
# the host before every run.
$SdkInstallerName = 'dotnet-sdk-11.0.100-win-x64.exe'
$SdkSha512 = '9777f9c9b8ea136a3aa5284f6bd33b0f1b3133fbc96d05d5c725683ccc101bc2fb2830c4182e199ec236c61a5e703a77b05770dcf3c881552ef77c7ada0584ff'
$SdkVersionPrefix = '11.'
$FilterBase = 'Category=Live&Requires!=DelegateStore&Requires!=CachedExchange'
$FilterUnindexedExtra = '&Requires!=SearchIndex'
# The Exchange VM with its test account as the hub: LiveRunFilters.ExchangeTestAccount, read and write. From
# 2026-11-17 to decision 206 (a) on 2026-11-18 it was narrowed to the tests that write nothing, because Office
# there was past its grace and created no item (Docs/live-tier-on-the-vm.md section 4.4a).
$FilterExchange = 'Category=Live&Requires!=LocalSignatures&Requires!=DelegateStore&Requires!=GeneratedPopulation&Requires!=IdentityAccount&Requires!=ProbePopulation&Requires!=PstStore&Requires!=ConfirmedLeftover'
# The host's record of the Exchange VM's run markers: beside the leases, machine state, not the repository.
$ExchangeLedgerRelative = 'exchange-runs\ledger.jsonl'
$LeaseReasonPrefix = 'Invoke-LiveTierOnGuest'
$MinFreeGbForCheckpoint = 60
$ExitCodes = [ordered]@{ PASS = 0; FAIL = 1; BUILD = 2; INFRA = 3; REFUSED = 4 }

# Every program call goes through Invoke-NativeCommand, so that Windows PowerShell 5.1 cannot end
# the script on a program's first stderr line (Q78). It is one of the helpers the scripts share,
# dot-sourced from Testbed/guest/ScriptHelpers.ps1 (Q154, Q166 (b), decision 215 (a)), which says why;
# so are this runner's Say, Stop-Refused and Format-Seconds, its run id, its TRX reader, its lease
# decision and the guest task's exit code, which the build-VM runner and the trust-removal probe use too.
. (Join-Path $PSScriptRoot '..\guest\ScriptHelpers.ps1')

# =============================================================================================
# PURE DECISIONS. No Hyper-V, no git, no guest. -SelfTest pins them.
# =============================================================================================

# From when the guest's Application log is searched for an OUTLOOK.EXE crash: a minute before the suite
# started BY THE GUEST'S CLOCK, which is what its events carry - the host's only when the guest's could not
# be read. On a frozen guest the two differ by hours or days (Q130 (a)).
function Get-CrashEventsSince {
    param([string] $GuestRunStartUtc, [DateTime] $HostRunStartUtc)
    $start = [DateTime]::SpecifyKind($HostRunStartUtc, [DateTimeKind]::Utc)
    $parsed = [DateTime]::MinValue
    if ($GuestRunStartUtc -and [DateTime]::TryParse($GuestRunStartUtc, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$parsed)) {
        $start = [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
    }
    return $start.AddMinutes(-1).ToString('o')
}

# The guest's dump listing - "DUMP <bytes> <file name>" lines among whatever else the call printed - as
# objects; Program is the name WER puts first ("OUTLOOK.EXE.4242.dmp" is OUTLOOK.EXE's).
function ConvertFrom-DumpListing {
    param([string] $Text)
    foreach ($m in [regex]::Matches([string]$Text, '(?m)^DUMP (\d+) (\S.*?\.dmp)\s*$')) {
        $name = $m.Groups[2].Value
        $program = $name
        $p = [regex]::Match($name, '^(.+?)\.\d+\.dmp$')
        if ($p.Success) { $program = $p.Groups[1].Value }
        [pscustomobject]@{ Name = $name; Bytes = [long]$m.Groups[1].Value; Program = $program }
    }
}

# The first look at a dump, out of host\Read-CrashDump.ps1's report: the exception, where it struck, and the
# top frames of the faulting thread - what summary.txt shows; the whole report sits beside the dump.
function Get-DumpHeadline {
    param([string] $Report, [int] $Frames = 8)
    $out = New-Object System.Collections.Generic.List[string]
    $lines = @(([string]$Report) -split "`r?`n")
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].StartsWith('EXCEPTION  ')) {
            $out.Add($lines[$i].Trim())
            for ($j = $i + 1; $j -lt $lines.Count -and $lines[$j].StartsWith('           '); $j++) { $out.Add($lines[$j].Trim()) }
        }
        if ($lines[$i].StartsWith('STACK      ')) {
            for ($j = $i + 1; $j -lt $lines.Count -and $j -le $i + $Frames -and $lines[$j] -match '^  \d\d  '; $j++) { $out.Add($lines[$j].Trim()) }
        }
    }
    return $out.ToArray()
}

# The run's OUTLOOK.EXE crash count: the Application log's events, or the dumps WER wrote, whichever
# says more - a dump is a crash even when its event was not found, and -1 (unreadable) stays -1 only
# when no dump says otherwise.
function Merge-CrashCount {
    param([int] $Events, [object[]] $Dumps)
    $outlook = @($Dumps | Where-Object { $_ -and ([string]$_.Program) -ieq 'OUTLOOK.EXE' }).Count
    if ($Events -lt 0 -and $outlook -eq 0) { return -1 }
    return [Math]::Max($Events, $outlook)
}

# Why a hunting run (-HuntFirstChance) cannot be had as asked - '' when it can. The hunt runs its own tests, on an
# Outlook test guest only - never the Exchange VM, whose Outlook holds a real mailbox that a full dump would hold too -
# and its rounds and minutes are bounded, inside the run's own time limit. Pure.
function Get-HuntRefusal {
    param([bool] $Hunt, [bool] $Exchange, [string] $Filter, [string] $FilterSuffix, [string] $SignatureSettings, [bool] $SignatureDialog, [int] $Rounds, [int] $Minutes, [int] $RunTimeoutMinutes, [string] $Variant = 'Legs', [int] $LegMinutes = 10, [bool] $WithoutAddIn = $false, [bool] $WithAddIn = $false, [bool] $ShowEach = $false, [bool] $TypeEach = $false, [string] $Ablate = '', [bool] $TraceClr = $false, [bool] $PromoteAfter = $false, [bool] $ReleaseAfterClose = $false)
    if (-not $Hunt) {
        if ($Variant -and $Variant -ne 'Legs') { return "-HuntVariant $Variant is a hunting run's: give -HuntFirstChance with it." }
        if ($WithoutAddIn) { return "-HuntWithoutAddIn is a hunting run's: give -HuntFirstChance with it." }
        if ($ShowEach) { return "-HuntShowEach is a hunting run's: give -HuntFirstChance and a -HuntVariant with it." }
        if ($Ablate) { return "-HuntAblate is a hunting run's: give -HuntFirstChance with it." }
        if ($TraceClr) { return "-HuntTraceClr is a hunting run's: give -HuntFirstChance with it." }
        if ($PromoteAfter) { return "-HuntPromoteAfter is a hunting run's: give -HuntFirstChance and a -HuntVariant with it." }
        if ($ReleaseAfterClose) { return "-HuntReleaseAfterClose is a hunting run's: give -HuntFirstChance, a -HuntVariant and -HuntShowEach -HuntTypeEach with it." }
        return ''
    }
    if ($PromoteAfter -and $Variant -notin @('ProductLoop', 'OtherFolder', 'NoWindow', 'PersonTyping')) { return "-HuntPromoteAfter re-opens the product's drafts: give a -HuntVariant whose loop makes them, not $Variant." }
    if ($Ablate -and (Test-HuntLoopVariant $Variant)) { return "-HuntAblate leaves a step out of the typed-text tests: give -HuntVariant Legs or LegsOtherFolder, not $Variant." }
    if ($TypeEach -and -not $ShowEach) { return '-HuntTypeEach types into the window -HuntShowEach shows: give both.' }
    if ($ReleaseAfterClose -and -not $TypeEach) { return '-HuntReleaseAfterClose releases the Word document of the window -HuntTypeEach types into: give -HuntShowEach -HuntTypeEach with it.' }
    if ($ShowEach -and -not (Test-HuntLoopVariant $Variant)) { return '-HuntShowEach shows a window before each draft of a loop: give a -HuntVariant other than Legs and LegsOtherFolder, whose typed-text tests show their own.' }
    if ($WithoutAddIn -and $WithAddIn) { return '-HuntWithoutAddIn turns the add-in off and -WithAddIn installs it: give one.' }
    if ($Exchange) { return '-HuntFirstChance attaches a debugger to Outlook and writes its whole memory to a dump: never on the Exchange VM, whose Outlook holds a real mailbox.' }
    if ($Filter -or $FilterSuffix) { return '-HuntFirstChance runs its own tests, the ones the crashes of 2026-11-18 struck in: give no -Filter or -FilterSuffix.' }
    if ($SignatureSettings -or $SignatureDialog) { return '-HuntFirstChance and a signature step together are two runs: give one.' }
    if ($Rounds -lt 1 -or $Rounds -gt 100) { return "-HuntRounds must be 1 to 100, not $Rounds." }
    $most = $RunTimeoutMinutes - 15
    if ($Minutes -lt 1 -or $Minutes -gt $most) { return "-HuntMinutes must be 1 to $most - inside the run's time limit of $RunTimeoutMinutes minutes, with room for the build and the stop - not $Minutes." }
    if ((Test-HuntLoopVariant $Variant) -and ($LegMinutes -lt 1 -or $LegMinutes -gt $Minutes)) { return "-HuntLegMinutes must be 1 to -HuntMinutes ($Minutes), not $LegMinutes." }
    return ''
}

# The guest task that turns OutlookAI's add-in off for a -HuntWithoutAddIn run: LoadBehavior 0 under the key the
# installer registers it under, read back - "ADDIN-OFF 0 (was 3)" - or "ADDIN-OFF none" when there is no such key,
# which the run takes for a failure: every Outlook guest's checkpoint holds the add-in installed. Pure.
function Get-AddInOffScript {
    param([string] $Key)
    return "`$k = '$Key'; if (-not (Test-Path -LiteralPath `$k)) { 'ADDIN-OFF none: no key' } else { `$was = (Get-ItemProperty -LiteralPath `$k).LoadBehavior; Set-ItemProperty -LiteralPath `$k -Name LoadBehavior -Value 0 -Type DWord; 'ADDIN-OFF {0} (was {1})' -f (Get-ItemProperty -LiteralPath `$k).LoadBehavior, `$was }"
}

# -HuntTraceClr's guest tasks, elevated in session 0. The start: an ETW session for Microsoft-Windows-DotNETRuntime's
# Exception keyword (0x8000) at Informational, writing clr.etl into the run's folder - "CLR-TRACE-START 0" when logman
# started it. The stop: the runtime's rundown (Loader, JIT and EndRundown at Verbose: every compiled method, to name the
# frames) for ten seconds, both sessions stopped, and tracerpt's XML of both - "CLR-TRACE-STOP <exit> <bytes>". Pure.
function Get-ClrTraceStartScript {
    param([string] $GuestRunDir, [string] $Session)
    return "New-Item -ItemType Directory -Force -Path '$GuestRunDir' | Out-Null; & logman stop $Session -ets *> `$null; & logman start $Session -p 'Microsoft-Windows-DotNETRuntime' 0x8000 0x4 -o '$GuestRunDir\clr.etl' -ets; `"CLR-TRACE-START `$LASTEXITCODE`""
}

function Get-ClrTraceStopScript {
    param([string] $GuestRunDir, [string] $Session)
    return "& logman start ${Session}Rundown -p 'Microsoft-Windows-DotNETRuntimeRundown' 0x118 0x5 -o '$GuestRunDir\clr-rundown.etl' -ets; Start-Sleep -Seconds 10; & logman stop ${Session}Rundown -ets; & logman stop $Session -ets; & tracerpt '$GuestRunDir\clr.etl' '$GuestRunDir\clr-rundown.etl' -o '$GuestRunDir\clr-events.xml' -of XML -y *> `$null; `$x = `$LASTEXITCODE; `"CLR-TRACE-STOP `$x `$((Get-Item -LiteralPath '$GuestRunDir\clr-events.xml' -ErrorAction SilentlyContinue).Length)`""
}

# The legs of each round of a hunting run, by -HuntVariant: Legs, the tests the crashes struck in ($HuntLegs); any other
# variant, the one loop leg. Each a suffix of the guest's derived filter. Pure.
function Get-HuntLegs {
    param([string] $Variant)
    if (-not $Variant -or $Variant -eq 'Legs') { return $HuntLegs }
    if ($Variant -eq 'LegsOtherFolder') { return [ordered]@{ calendar = $HuntLoopLeg; typed = $HuntLegs['typed']; drafts = $HuntLegs['drafts'] } }
    return [ordered]@{ loop = $HuntLoopLeg }
}

# The variant the hunting test is told for a runner's -HuntVariant: none for Legs, whose legs are not the test's;
# CalendarOnly for LegsOtherFolder, the step before its legs (T2.HuntPlan.CalendarStep); the loop's own name else. Pure.
function Get-HuntTestVariant {
    param([string] $Variant)
    switch ($Variant) {
        'Legs' { return '' }
        'LegsOtherFolder' { return 'CalendarOnly' }
        default { return $Variant }
    }
}

# Whether a runner's -HuntVariant runs the loop: every one but the two that run the tests the crashes struck in. Pure.
function Test-HuntLoopVariant {
    param([string] $Variant)
    return [bool]($Variant -and $Variant -ne 'Legs' -and $Variant -ne 'LegsOtherFolder')
}

# The suite's script for a hunting run. The debugger is started through the shell (Start-Process), so it inherits
# nothing of the suite's task - its output pipe least of all (2026-11-18) - and is stopped through its stop file, after
# which it detaches. Each round runs every leg in a dotnet test of its own, the first one building; each leg's TRX
# file is named for its round and leg. The rounds end at -HuntRounds, at -HuntMinutes, or at the first dump the
# debugger wrote. Pure.
function Get-HuntSuiteScript {
    param([string] $OptIn, [string] $GuestRunDir, [string] $GuestRoot, [System.Collections.IDictionary] $Filters, [int] $Rounds, [int] $Minutes, [string] $Module, [string] $DumpDir, [string] $ScriptName, [string] $Variant = 'Legs', [int] $LegMinutes = 10, [bool] $ShowEach = $false, [bool] $TypeEach = $false, [string] $Ablate = '', [bool] $PromoteAfter = $false, [bool] $ReleaseAfterClose = $false)
    $legLines = (@($Filters.Keys) | ForEach-Object { "`$legs += ,@('$_', '$($Filters[$_])')" }) -join "`r`n"
    $variantLines = ''
    $testVariant = Get-HuntTestVariant $Variant
    if ($testVariant) { $variantLines = "`$env:$HuntVariantVariable = '$testVariant'`r`n`$env:$HuntLegMinutesVariable = '$LegMinutes'`r`n`$env:$HuntShowEachVariable = '$(if ($TypeEach) { 'typed' } else { [int]$ShowEach })'`r`n" }
    if ($Ablate) { $variantLines += "`$env:$HuntAblateVariable = '$Ablate'`r`n" }
    if ($PromoteAfter) { $variantLines += "`$env:$HuntPromoteAfterVariable = '1'`r`n" }
    if ($ReleaseAfterClose) { $variantLines += "`$env:$HuntReleaseAfterCloseVariable = '1'`r`n" }
    return @"
$variantLines`$env:OUTLOOKAI_LIVE_OPT_IN = '$OptIn'
`$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'; `$env:DOTNET_NOLOGO = '1'; `$env:MSBUILDDISABLENODEREUSE = '1'
New-Item -ItemType Directory -Force -Path '$GuestRunDir', '$GuestRunDir\diag' | Out-Null
`$c2r = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
foreach (`$n in 'VersionToReport', 'ClientVersionToReport', 'ProductReleaseIds', 'Platform', 'AudienceData', 'UpdateChannel', 'CDNBaseUrl', 'UpdatesEnabled') { Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value ('HUNT-OFFICE {0} = {1}' -f `$n, `$(if (`$null -ne `$c2r) { `$c2r.`$n } else { '(no Click-to-Run configuration)' })) }
`$m = Get-ChildItem -LiteralPath 'C:\Program Files\Microsoft Office\root' -Recurse -Filter '$Module' -ErrorAction SilentlyContinue | Select-Object -First 1
if (`$null -ne `$m) { Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value ('HUNT-OFFICE module {0}, {1} bytes, written {2:o}, SHA-256 {3}' -f `$m.FullName, `$m.Length, `$m.LastWriteTimeUtc, (Get-FileHash -LiteralPath `$m.FullName -Algorithm SHA256).Hash) }
else { Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value 'HUNT-OFFICE module $Module not found under Office''s root' }
Set-Location '$GuestRoot\src'
`$stop = '$GuestRunDir\first-chance.stop'
`$watch = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', '$GuestRoot\$ScriptName', '-DumpFolder', '$DumpDir', '-StopFile', `$stop, '-Log', '$GuestRunDir\first-chance.log', '-Module', '$Module', '-MaxMinutes', '$($Minutes + 30)') -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 5
`$legs = @()
$legLines
`$clock = [System.Diagnostics.Stopwatch]::StartNew()
`$exit = 0
`$built = `$false
`$done = 0
while (`$done -lt $Rounds -and `$clock.Elapsed.TotalMinutes -lt $Minutes) {
    `$round = `$done + 1
    foreach (`$p in @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) { `$clr = 'unread'; try { `$clr = [string](@(`$p.Modules | Where-Object { `$_.ModuleName -eq 'clr.dll' }).Count -gt 0) } catch { }; Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value ('HUNT-OUTLOOK round {0} pid {1} clr {2}' -f `$round, `$p.Id, `$clr) }
    foreach (`$leg in `$legs) {
        `$testArgs = @('test', 'McpServer\OutlookAI.McpServer.Tests\OutlookAI.McpServer.Tests.csproj', '-c', 'Release', '--filter', `$leg[1], '--logger', ('trx;LogFileName=live-r{0:D2}-{1}.trx' -f `$round, `$leg[0]), '--logger', 'console;verbosity=detailed', '--results-directory', '$GuestRunDir', '--diag', ('$GuestRunDir\diag\vstest-r{0:D2}-{1}.log' -f `$round, `$leg[0]))
        if (`$built) { `$testArgs += '--no-build' }
        & dotnet @testArgs *>> '$GuestRunDir\console.txt'
        if (`$LASTEXITCODE -ne 0) { `$exit = `$LASTEXITCODE }
        `$built = `$true
    }
    `$done = `$round
    Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value "HUNT-ROUND `$done"
    if (@(Get-ChildItem -LiteralPath '$DumpDir' -Filter '*chance.dmp' -ErrorAction SilentlyContinue).Count -gt 0) { Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value 'HUNT-DUMP the debugger wrote a dump - the hunt stops'; break }
}
New-Item -ItemType File -Force -Path `$stop | Out-Null
if (`$null -ne `$watch) { `$null = `$watch.WaitForExit(120000) }
Add-Content -LiteralPath '$GuestRunDir\console.txt' -Value "HUNT-ROUNDS `$done"
exit `$exit
"@
}

# Every TRX file of a hunting run as one result: the counts summed, each failed or skipped test named with the file it
# came from ("[live-r03-typed] <test>"), the files' own counters summed when every file has them. Pure.
function Merge-TrxResults {
    param([object[]] $Files)    # each: Label, Result (ConvertFrom-TrxText's)
    $counts = [ordered]@{ total = 0; passed = 0; failed = 0; skipped = 0; other = 0 }
    $declared = [ordered]@{ total = 0; passed = 0; failed = 0; notExecuted = 0 }
    $allDeclared = @($Files).Count -gt 0
    $failed = @()
    $skipped = @()
    foreach ($f in @($Files)) {
        foreach ($k in @($counts.Keys)) { $counts[$k] += [int]$f.Result.Counts[$k] }
        if ($null -eq $f.Result.Declared) { $allDeclared = $false }
        else { foreach ($k in @($declared.Keys)) { $declared[$k] += [int]$f.Result.Declared[$k] } }
        foreach ($x in @($f.Result.Failed)) { $failed += [pscustomobject]@{ name = "[$($f.Label)] $($x.name)"; message = $x.message } }
        foreach ($s in @($f.Result.Skipped)) { $skipped += "[$($f.Label)] $s" }
    }
    if (-not $allDeclared) { $declared = $null }
    return [pscustomobject]@{ Counts = $counts; Declared = $declared; Failed = $failed; Skipped = $skipped }
}

# Office's build as a hunting run's guest read it - the HUNT-OFFICE lines of its console: the Click-to-Run values
# by name, and the hunted module's file under 'module'; null when there are none. Recorded since 2026-11-18, when the
# question became whether a later build of the channel fixes the crash (runbook section 4.15). Pure.
function Read-HuntOffice {
    param([string] $Text)
    $office = [ordered]@{}
    foreach ($m in [regex]::Matches([string]$Text, '(?m)^HUNT-OFFICE (\w+) = (.*?)\s*$')) { $office[$m.Groups[1].Value] = $m.Groups[2].Value }
    $module = [regex]::Match([string]$Text, '(?m)^HUNT-OFFICE module (.*?)\s*$')
    if ($module.Success) { $office['module'] = $module.Groups[1].Value }
    if ($office.Count -eq 0) { return $null }
    return $office
}

# The files of WER's own reports of an OUTLOOK.EXE crash that the guest copied beside the dumps -
# "WERFILE <bytes> <file name>" lines among whatever else the call printed - as objects; each name is the
# report folder's, a dot, and the file's. Kept since 2026-11-18: two crashes in mso40uiwin32client.dll that
# day left an Application Error event and no LocalDumps dump, and only WER's report of a crash says which
# way it went (runbook section 4.15).
function ConvertFrom-WerListing {
    param([string] $Text)
    foreach ($m in [regex]::Matches([string]$Text, '(?m)^WERFILE (\d+) (\S.*?)\s*$')) {
        [pscustomobject]@{ Name = $m.Groups[2].Value; Bytes = [long]$m.Groups[1].Value }
    }
}

# What PREPARE does, by whether the start checkpoint is frozen (its time sync off, Q130 (a)). A frozen
# start is verified, never restarted or rebuilt: either would leave the instant its data was built for.
function Get-PrepareSteps {
    param([bool] $StartIsFrozen, [bool] $SkipHubReset, [string] $SignatureSettings = '')
    if ($StartIsFrozen -and $SignatureSettings) { return @('frozen-clock guard', "signature setting $SignatureSettings, before Outlook starts", 'Outlook NOT elevated') }
    if ($StartIsFrozen) { return @('frozen-clock guard', 'Outlook NOT elevated') }
    if ($SkipHubReset) { return @('restart') }
    return @('restart', 'hub rebuild')
}

# The guest task's lines that read Outlook's own Signatures dialog after a -SignatureSettings step, before the
# suite - what the dialog then lets a user change or delete (Q156, decided 2026-11-18) - and that tell the suite
# where the reading is; '' without the step. The probe never stops the task: a missing reading fails the test that
# needs it, and the probe's log says why.
function Get-SignatureDialogLines {
    param([string] $SignatureSettings, [string] $GuestRoot, [string] $GuestRunDir, [bool] $Dialog = $true)
    if (-not $SignatureSettings -or -not $Dialog) { return '' }
    $reading = "$GuestRunDir\$SignatureDialogReading"
    $log = "$GuestRunDir\signature-dialog.log"
    return "try { & '$GuestRoot\$SignatureDialogScript' -OutPath '$reading' *> '$log' } catch { [System.IO.File]::AppendAllText('$log', 'SIGNATURE-DIALOG-THREW: ' + `$_.Exception.Message) }`r`n" +
        "`$env:$SignatureDialogVariable = '$reading'`r`n"
}

# Why a run may not take the -SignatureSettings step, or '' when it may (or when it was not asked for). The step
# writes settings Outlook reads when it starts, so it runs only where Outlook starts once, after it - a frozen start
# (Q130 (a)) - and only where the checkpoint restored at REST takes the setting away again.
function Get-SignatureSettingsRefusal {
    param([string] $Scenario, [bool] $Exchange, [bool] $StartIsFrozen, [string] $GreenCheckpoint, [string] $Override, [bool] $Dialog = $false)
    if (-not $Scenario) {
        if ($Dialog) { return '-SignatureDialog without -SignatureSettings: the dialog is read under a setting the step makes.' }
        return ''
    }
    if ($Exchange) { return '-SignatureSettings on the Exchange VM: its state is carried from run to run, so no checkpoint restore would take the setting away again.' }
    if ($GreenCheckpoint) { return "-SignatureSettings with -GreenCheckpoint: that checkpoint would keep the setting, and a run resting on it would start under it." }
    if ($Override) { return '-SignatureSettings with -Filter: the step''s run is narrowed to its own test; -FilterSuffix narrows it further.' }
    if (-not $StartIsFrozen) { return '-SignatureSettings from a checkpoint that is not frozen: such a start restarts the guest and rebuilds its hub, starting Outlook more than once. The step runs from a frozen checkpoint, where Outlook starts once, after it.' }
    return ''
}

# The guest's facts, or $null for anything that is not a test guest - one of the two Outlook guests
# or a clone of one in their pools (TestbedPool.ps1) - compared whole and ignoring case, as Hyper-V
# compares names. A clone's facts are its source's under its own name, and Source names the guest
# whose records - liveTestSettings, frozenClocks - it is rendered and verified from.
function Get-GuestFacts {
    param([string] $Name)
    # The Exchange VM is no pool's and is never cloned (TestbedPool.ps1): its facts are its own.
    if ([string]::Equals($Name, 'OutlookAI-Exchange', [System.StringComparison]::OrdinalIgnoreCase)) {
        $exchange = $Guests['OutlookAI-Exchange']
        return [pscustomobject]@{ Name = 'OutlookAI-Exchange'; Source = 'OutlookAI-Exchange'; Kind = 'Exchange'; ComputerName = $exchange.ComputerName; Indexed = $exchange.Indexed; Checkpoint = $exchange.Checkpoint; Exchange = $true }
    }
    $pool = Get-TestbedPoolOf $Name
    if ($pool -ne 'indexed' -and $pool -ne 'unindexed') { return $null }
    $source = $TestbedPools[$pool][0]
    if (-not $Guests.Contains($source)) { return $null }
    $own = @($TestbedPools[$pool] | Where-Object { [string]::Equals($_, $Name, [System.StringComparison]::OrdinalIgnoreCase) })[0]
    $facts = $Guests[$source]
    $guestKind = 'Unindexed'
    if ($facts.Indexed) { $guestKind = 'Indexed' }
    return [pscustomobject]@{ Name = $own; Source = $source; Kind = $guestKind; ComputerName = $facts.ComputerName; Indexed = $facts.Indexed; Checkpoint = $facts.Checkpoint; Exchange = $false }
}

# The guests of a kind, in the order -Kind tries them: the source first, then its clones.
function Get-KindMembers {
    param([string] $Kind)
    if ($Kind -eq 'Indexed') { return @($TestbedPools['indexed']) }
    if ($Kind -eq 'Unindexed') { return @($TestbedPools['unindexed']) }
    throw "'$Kind' is not a kind of guest: Indexed or Unindexed."
}

# The run's filter: the guest's derived one (LiveRunFilters.Guest / GuestUnindexed) narrowed by
# -FilterSuffix, or -Filter whole. Returns the refusal instead when the request is malformed. The
# filter is written into the guest task's script inside double quotes, so a character that would
# end or expand that string is refused rather than escaped.
function Get-LiveFilter {
    param([bool] $Indexed, [string] $Override, [string] $Suffix, [bool] $Exchange = $false, [string] $SignatureSettings = '', [bool] $SignatureDialog = $false)
    foreach ($given in @($Override, $Suffix)) {
        if ($given -match '["`$\r\n]') {
            return [pscustomobject]@{ Filter = $null; Narrowed = $false; Refusal = "a filter may not hold a double quote, a backtick, a dollar sign or a line break: '$given'" }
        }
    }
    if ($Override) {
        if ($Suffix) { return [pscustomobject]@{ Filter = $null; Narrowed = $false; Refusal = '-Filter and -FilterSuffix together are ambiguous; give one.' } }
        if (-not $Override.StartsWith('Category=Live', [System.StringComparison]::Ordinal)) {
            return [pscustomobject]@{ Filter = $null; Narrowed = $false; Refusal = "-Filter must start with Category=Live, so it can never select a test that is not a live one: '$Override'" }
        }
        return [pscustomobject]@{ Filter = $Override; Narrowed = $true; Refusal = $null }
    }
    if ($Suffix -and -not $Suffix.StartsWith('&', [System.StringComparison]::Ordinal)) {
        return [pscustomobject]@{ Filter = $null; Narrowed = $false; Refusal = "-FilterSuffix must start with '&' - it NARROWS the derived filter: '$Suffix'" }
    }
    $filter = $FilterBase
    if (-not $Indexed) { $filter += $FilterUnindexedExtra }
    if ($Exchange) { $filter = $FilterExchange }
    # A run that made a signature setting before Outlook started runs the one test that measures it - what Outlook
    # inserts, or with -SignatureDialog its dialog: every other signature test would meet a setting it was not written for.
    $step = ''
    if ($SignatureSettings) { $step = "&FullyQualifiedName~$(if ($SignatureDialog) { $SignatureDialogTest } else { $SignatureInsertsTest })" }
    return [pscustomobject]@{ Filter = $filter + $Suffix + $step; Narrowed = ([bool]$Suffix -or [bool]$step); Refusal = $null }
}

# What a run directory's name says between its start and its commit (New-RunId): which kind of guest
# ran it, a clone its source's kind. A name taken in the same second gets -2, -3 ... (New-RunDirectory).
function Get-RunLabel {
    param([string] $GuestName)
    if ($GuestName -match 'Exchange$') { return 'exchange' }
    if ($GuestName -match 'Unindexed\d*$') { return 'unindexed' }
    return 'indexed'
}

# How the Exchange VM's run begins: carried forward from its last run, or - only when a checkpoint is
# named - restored, restarted and its data files set aside.
function Get-ExchangeStartPlan {
    param([bool] $CheckpointNamed)
    if ($CheckpointNamed) { return @('restore the named checkpoint', 'restart', 'set the .ost files aside', 'rebuild the search index', 'Outlook NOT elevated', 'wait for one index root per indexed store', 'preflight') }
    return @('resume as the last run left it', 'Outlook NOT elevated if it is not running', 'preflight')
}

# What a start of Outlook by guest\Start-OutlookUnelevated.ps1 reports about Office's activation wizard -
# the modal nag a guest past Office's activation grace raises at every start with a window (Testbed/
# MEDIA.md, "The licence clocks"). The starter closes it and prints 'ACTIVATION-WIZARD-DISMISSALS <n>';
# an older starter prints nothing, which reads as not reported, never as zero.
function Get-OutlookStartFacts {
    param([string] $Output, [string] $Profile)
    $m = [regex]::Match([string]$Output, 'ACTIVATION-WIZARD-DISMISSALS (\d+)')
    $n = $null
    if ($m.Success) { $n = [int]$m.Groups[1].Value }
    return [ordered]@{ startedByRun = $true; profile = $Profile; activationWizardDismissed = $n }
}

# The summary's line for it.
function Format-OutlookStartLine {
    param($Facts)
    if ($null -eq $Facts) { return 'outlook    not started by this run (it was running already)' }
    $wizard = "Office's activation wizard dismissed $($Facts['activationWizardDismissed']) time(s)"
    if ($null -eq $Facts['activationWizardDismissed']) { $wizard = "Office's activation wizard: NOT REPORTED by the guest's starter" }
    return "outlook    started by this run, NOT elevated, on profile '$($Facts['profile'])'; $wizard"
}

# The Exchange VM's run marker: 'oair' and the run's UTC start to the second - lower-case letters and
# digits, 18 of them, the shape T2/LiveRunMarker.cs accepts. One run at a time holds the guest's lease, so
# two runs never share a second.
function New-RunMarker {
    param([DateTime] $WhenUtc)
    return 'oair' + $WhenUtc.ToUniversalTime().ToString('yyyyMMddHHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
}

# One ledger line: a run started, or finished and whether its sweep reported the hub clean.
function ConvertTo-LedgerLine {
    param([string] $Marker, [string] $RunId, [string] $Event, [object] $Clean, [string[]] $Priors = @(), [DateTime] $WhenUtc)
    $entry = [ordered]@{ marker = $Marker; runId = $RunId; event = $Event; utc = $WhenUtc.ToUniversalTime().ToString('o'); clean = $Clean; priors = @($Priors) }
    return ($entry | ConvertTo-Json -Compress -Depth 4)
}

# The markers the next run must sweep: every recorded run that never finished clean, newest first, at
# most -Max of them - and none a clean finish has since covered (a clean run swept its priors too).
function Get-PriorRunMarkers {
    param([string[]] $LedgerLines, [string] $Current, [int] $Max = 25)
    $started = New-Object System.Collections.Generic.List[string]
    $cleared = New-Object System.Collections.Generic.HashSet[string]
    foreach ($line in @($LedgerLines)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $e = $null
        # A ledger first written by Windows PowerShell 5.1 starts with a byte-order mark.
        try { $e = $line.TrimStart([char]0xFEFF) | ConvertFrom-Json } catch { continue }
        $m = [string]$e.marker
        if ($m -cnotmatch '^[a-z0-9]{12,40}$') { continue }
        if ([string]$e.event -eq 'started' -and -not $started.Contains($m)) { $started.Add($m) }
        if ([string]$e.event -eq 'finished' -and $e.clean -eq $true) {
            [void]$cleared.Add($m)
            foreach ($p in @($e.priors)) { if ($p) { [void]$cleared.Add([string]$p) } }
        }
    }
    $open = @($started | Where-Object { $_ -cne $Current -and -not $cleared.Contains($_) })
    [array]::Reverse($open)
    return @($open | Select-Object -First $Max)
}

# What the suite's console says about its own safety proofs. xUnit writes each output line twice -
# once under its "[xUnit.net 00:00:01.23]" prefix, once plain - so lines are compared without it:
#   ProvedNothing - the distinct "PROVED NOTHING:" lines (a check with nothing to check here);
#   SweepRan      - whether an artifact sweep ran ("artifact sweep: N store(s) ...",
#                   ArtifactSweepPlan.Describe);
#   ArtifactsLeft - the most any store's final count said ("sweep[<store>]: taggedArtifacts=N",
#                   ArtifactSweepPlan.Announce), or $null when no store reported;
#   Tripwire      - the failures the last "[tripwire] post-run census in ..." line reported, or $null
#                   when the census never reported.
function Read-ConsoleFacts {
    param([string] $Text)
    $proved = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $sweepRan = $false
    $left = $null
    $tripwire = $null
    foreach ($raw in ([string]$Text -split "`r?`n")) {
        $line = ($raw -replace '^\s*\[xUnit\.net [^\]]*\]\s*', '').Trim()
        if ($line.StartsWith('PROVED NOTHING:', [System.StringComparison]::Ordinal) -and $seen.Add($line)) { $proved.Add($line) }
        if ($line.StartsWith('artifact sweep: ', [System.StringComparison]::Ordinal)) { $sweepRan = $true }
        $m = [regex]::Match($line, '^sweep\[[^\]]*\]: taggedArtifacts=(\d+)')
        if ($m.Success) {
            $n = [int]$m.Groups[1].Value
            if ($null -eq $left -or $n -gt $left) { $left = $n }
        }
        $t = [regex]::Match($line, '^\[tripwire\] post-run census in .*?; (\d+) failure\(s\)')
        if ($t.Success) { $tripwire = [int]$t.Groups[1].Value }
    }
    return [pscustomobject]@{ ProvedNothing = $proved.ToArray(); SweepRan = $sweepRan; ArtifactsLeft = $left; Tripwire = $tripwire }
}

# The verdict, from what came back. A stage that broke is BUILD; a run that never reached the suite
# is INFRA; once the suite ran, the TRX file and the safety proofs decide - and every proof can only
# FAIL a run, never pass one.
function Get-LiveVerdict {
    param(
        [string] $Stage,        # 'BUILD' / 'INFRA' when the run stopped before the suite finished, else ''
        [string] $StageWhy,
        [object] $Trx,          # ConvertFrom-TrxText's result, or $null
        [object] $Console,      # Read-ConsoleFacts's result, or $null
        [int]    $Crashes,      # OUTLOOK.EXE crash events during the run; -1 when they could not be read
        [object] $SuiteExit,    # dotnet test's exit code, or $null when unknown
        [bool]   $Narrowed      # -Filter or -FilterSuffix was given
    )
    if ($Stage -eq 'BUILD') { return [pscustomobject]@{ Verdict = 'BUILD'; Why = $StageWhy } }
    if ($Stage -eq 'INFRA') {
        if ($null -ne $Trx -and ($Trx.Counts['failed'] + $Trx.Counts['other']) -gt 0) {
            return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$($Trx.Counts['failed'] + $Trx.Counts['other']) test(s) did not pass, by the TRX file - and the run did not finish: $StageWhy" }
        }
        return [pscustomobject]@{ Verdict = 'INFRA'; Why = $StageWhy }
    }
    if ($null -eq $Trx) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the suite ran and no TRX file came back' } }
    if ($Trx.Counts['total'] -eq 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the filter selected no test - nothing was tested' } }
    $why = @()
    $notPassed = $Trx.Counts['failed'] + $Trx.Counts['other']
    if ($notPassed -gt 0) { $why += "$notPassed test(s) did not pass" }
    if ($Crashes -gt 0) { $why += "OUTLOOK.EXE crashed $Crashes time(s) during the run" }
    if ($null -ne $Console -and $null -ne $Console.ArtifactsLeft -and $Console.ArtifactsLeft -gt 0) { $why += "an artifact sweep reported $($Console.ArtifactsLeft) tagged artifact(s) left in a store" }
    if ($null -ne $Console -and $null -ne $Console.Tripwire -and $Console.Tripwire -gt 0) { $why += "the count tripwire's post-run census reported $($Console.Tripwire) failure(s)" }
    if (-not $Narrowed) {
        if ($null -eq $Console -or -not $Console.SweepRan) { $why += 'no artifact sweep ran - zero tagged artifacts was not proven' }
        if ($null -eq $Console -or $null -eq $Console.Tripwire) { $why += "the count tripwire's post-run census never reported" }
    }
    if ($notPassed -eq 0 -and $null -ne $SuiteExit -and [int]$SuiteExit -ne 0) { $why += "dotnet test exited $SuiteExit although every test in the TRX file passed - a fixture, the census or the run itself failed (console.txt)" }
    if ($why.Count -gt 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = ($why -join '; ') } }
    return [pscustomobject]@{ Verdict = 'PASS'; Why = '' }
}

# summary.json. EVERY LIST IS A JSON ARRAY AT EVERY LENGTH (Invoke-TestsOnBuildVm.ps1 says why).
function New-LiveRunSummary {
    param(
        [string] $RunId,
        [object] $Verdict,
        [System.Collections.IDictionary] $Revision,
        [string] $Guest,
        [string] $Filter,
        [bool]   $Narrowed,
        [object] $Trx,
        [object] $Console,
        [int]    $Crashes,
        [object] $SuiteExit,
        [System.Collections.IDictionary] $Checkpoints,
        [System.Collections.IDictionary] $Timings,
        [string] $Results,
        [System.Collections.IDictionary] $CrashDumps
    )
    $suite = $null; $failed = @(); $skipped = @()
    if ($null -ne $Trx) { $suite = $Trx.Counts; $failed = @($Trx.Failed); $skipped = @($Trx.Skipped) }
    $proved = @(); $sweepRan = $false; $left = $null; $tripwire = $null
    if ($null -ne $Console) { $proved = @($Console.ProvedNothing); $sweepRan = $Console.SweepRan; $left = $Console.ArtifactsLeft; $tripwire = $Console.Tripwire }
    $crashValue = $null
    if ($Crashes -ge 0) { $crashValue = $Crashes }
    return [ordered]@{
        runId          = $RunId
        verdict        = $Verdict.Verdict
        why            = $Verdict.Why
        exitCode       = $ExitCodes[$Verdict.Verdict]
        guest          = $Guest
        revision       = $Revision
        filter         = $Filter
        narrowed       = $Narrowed
        suite          = $suite
        suiteExitCode  = $SuiteExit
        failed         = $failed
        skipped        = $skipped
        provedNothing  = [ordered]@{ count = $proved.Count; lines = $proved }
        artifacts      = [ordered]@{ sweepRan = $sweepRan; leftInAStore = $left; proven = ($sweepRan -and $null -ne $left -and $left -eq 0) }
        tripwire       = [ordered]@{ reported = ($null -ne $tripwire); failures = $tripwire }
        outlookCrashes = $crashValue
        crashDumps     = $CrashDumps
        checkpoints    = $Checkpoints
        timings        = $Timings
        results        = $Results
    }
}

# THE GUEST'S CLOCK AROUND EVERY GUEST CALL (stream K, 2026-11-18). In two Q149 measurement runs on a
# frozen-clock guest, the guest's wall clock jumped about two hours forward during the suite, and a wait
# of 9000 s or 7200 s ran out after 35 to 40 minutes of real time. The cause is not established. So every
# guest call reads the guest's UTC time and time zone as its first and its last act, and the run sets
# them beside this host's clock: a line a call in guest-clock.log, and a warning in run.log and
# summary.txt when the two clocks ran more than this many seconds apart over a call, or the zone changed.
$GuestClockToleranceSeconds = 60

# The guest's two readings in a guest call's output - the lines GUEST-CLOCK-BEFORE and GUEST-CLOCK-AFTER,
# each the UTC time in round-trip form and the zone's id: Before and After, each with Utc and Zone, or
# $null for a line the output lacks. Pure.
function Read-GuestClock {
    param([string] $Text)
    $before = $null
    $after = $null
    foreach ($m in [regex]::Matches([string]$Text, '(?m)^GUEST-CLOCK-(BEFORE|AFTER) (\S+) (.+?)\s*$')) {
        $utc = [DateTime]::MinValue
        if (-not [DateTime]::TryParse($m.Groups[2].Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref] $utc)) { continue }
        $reading = [pscustomobject]@{ Utc = $utc.ToUniversalTime(); Zone = $m.Groups[3].Value }
        if ($m.Groups[1].Value -ceq 'BEFORE') { $before = $reading } else { $after = $reading }
    }
    return [pscustomobject]@{ Before = $before; After = $after }
}

# What one guest call did to the guest's clock beside this host's: how long each clock ran over the call,
# how far apart they ran, the guest's lead on this host at each end, the zone before and after, and
# Moved - more than $ToleranceSeconds apart, or another zone after than before. This host's readings,
# UTC, are taken with the session already open, just before and just after the call, so the guest's two
# fall between them. $null when the output lacks either of the guest's. Pure.
function Get-GuestClockShift {
    param([object] $Readings, [DateTime] $HostBefore, [DateTime] $HostAfter, [int] $ToleranceSeconds = $GuestClockToleranceSeconds)
    if ($null -eq $Readings -or $null -eq $Readings.Before -or $null -eq $Readings.After) { return $null }
    $guestSeconds = ($Readings.After.Utc - $Readings.Before.Utc).TotalSeconds
    $hostSeconds = ($HostAfter - $HostBefore).TotalSeconds
    $apart = $guestSeconds - $hostSeconds
    return [pscustomobject]@{
        GuestSeconds = $guestSeconds
        HostSeconds  = $hostSeconds
        ApartSeconds = $apart
        LeadBefore   = ($Readings.Before.Utc - $HostBefore).TotalSeconds
        LeadAfter    = ($Readings.After.Utc - $HostAfter).TotalSeconds
        ZoneBefore   = [string]$Readings.Before.Zone
        ZoneAfter    = [string]$Readings.After.Zone
        Moved        = (([Math]::Abs($apart) -gt $ToleranceSeconds) -or ([string]$Readings.Before.Zone -cne [string]$Readings.After.Zone))
    }
}

# Seconds as the clock lines print them: one decimal, a point, whatever this host's culture. Pure.
function Format-ClockSeconds {
    param([double] $Seconds, [switch] $Signed)
    $format = '0.0'
    if ($Signed) { $format = '+0.0;-0.0;0.0' }
    return $Seconds.ToString($format, [System.Globalization.CultureInfo]::InvariantCulture)
}

# A move as run.log and summary.txt say it. Pure.
function Format-GuestClockMove {
    param([string] $Call, [object] $Shift)
    $zone = "its zone '$($Shift.ZoneBefore)' throughout"
    if ($Shift.ZoneBefore -cne $Shift.ZoneAfter) { $zone = "its zone '$($Shift.ZoneBefore)' before and '$($Shift.ZoneAfter)' after" }
    return ("during '{0}' the guest's clock ran {1} s and this host's {2} s, {3} s apart; {4}" -f $Call, (Format-ClockSeconds $Shift.GuestSeconds), (Format-ClockSeconds $Shift.HostSeconds), (Format-ClockSeconds $Shift.ApartSeconds -Signed), $zone)
}

# One guest call's line in guest-clock.log: when it began on this host, the call, the guest's two
# readings and how far the two clocks ran apart. Pure.
function Format-GuestClockLine {
    param([string] $Call, [DateTime] $HostBefore, [object] $Readings, [object] $Shift)
    $head = '{0}  {1}' -f $HostBefore.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [System.Globalization.CultureInfo]::InvariantCulture), $Call
    if ($null -eq $Shift) { return "$head  the guest's clock was not read: the call's output holds no GUEST-CLOCK line, or only one" }
    $line = "{0}  guest {1} '{2}' to {3} '{4}', {5} s ahead of this host at the start; the guest's clock ran {6} s, this host's {7} s, {8} s apart" -f $head, $Readings.Before.Utc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [System.Globalization.CultureInfo]::InvariantCulture), $Shift.ZoneBefore, $Readings.After.Utc.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [System.Globalization.CultureInfo]::InvariantCulture), $Shift.ZoneAfter, (Format-ClockSeconds $Shift.LeadBefore -Signed), (Format-ClockSeconds $Shift.GuestSeconds), (Format-ClockSeconds $Shift.HostSeconds), (Format-ClockSeconds $Shift.ApartSeconds -Signed)
    if ($Shift.Moved) { $line += ' - MOVED' }
    return $line
}

# summary.txt's clock line, from every guest call's record (Call, Shift) in call order. Pure.
function Format-GuestClockSummary {
    param([object[]] $Records, [int] $ToleranceSeconds = $GuestClockToleranceSeconds)
    $all = @(@($Records) | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) { return 'clock      no guest call was made' }
    $read = @($all | Where-Object { $null -ne $_.Shift })
    $moved = @($read | Where-Object { $_.Shift.Moved })
    if ($moved.Count -gt 0) {
        return ("clock      MOVED in {0} of {1} guest calls: {2} (guest-clock.log)" -f $moved.Count, $all.Count, ((@($moved | ForEach-Object { Format-GuestClockMove $_.Call $_.Shift })) -join '; '))
    }
    if ($read.Count -eq 0) { return ("clock      the guest's clock was read around none of the {0} guest calls (guest-clock.log)" -f $all.Count) }
    $zones = @($read | ForEach-Object { $_.Shift.ZoneAfter } | Select-Object -Unique)
    return ("clock      the guest's clock kept to this host's, within {0} s, over every guest call it was read around: {1} of {2}, in '{3}' (guest-clock.log)" -f $ToleranceSeconds, $read.Count, $all.Count, ($zones -join "', '"))
}

# summary.json's guestClock: how many calls, how many read, the tolerance, and every move - an array at
# every length (the summary's rule). Pure.
function New-GuestClockSummaryJson {
    param([object[]] $Records, [int] $ToleranceSeconds = $GuestClockToleranceSeconds)
    $all = @(@($Records) | Where-Object { $null -ne $_ })
    $moved = @($all | Where-Object { $null -ne $_.Shift -and $_.Shift.Moved } | ForEach-Object {
            [ordered]@{ call = [string]$_.Call; guestSeconds = [Math]::Round($_.Shift.GuestSeconds, 1); hostSeconds = [Math]::Round($_.Shift.HostSeconds, 1); apartSeconds = [Math]::Round($_.Shift.ApartSeconds, 1); zoneBefore = $_.Shift.ZoneBefore; zoneAfter = $_.Shift.ZoneAfter }
        })
    return [ordered]@{
        calls            = $all.Count
        read             = @($all | Where-Object { $null -ne $_.Shift }).Count
        toleranceSeconds = $ToleranceSeconds
        moved            = $moved
        log              = 'guest-clock.log'
    }
}

function Invoke-SelfTest {
    $script:stChecks = 0
    $script:stFailures = @()
    function Check([string] $What, $Expected, $Actual) {
        $script:stChecks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = $Expected -join '|' }
        if ($Actual -is [System.Array]) { $a = $Actual -join '|' }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else { $script:stFailures += "$What : expected [$e], got [$a]"; Write-Host "  FAIL $What - expected [$e], got [$a]" }
    }
    $repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    Write-Host "Invoke-LiveTierOnGuest.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No Hyper-V, no git, no guest, no credential."

    Write-Host '== the guests =='
    $i = Get-GuestFacts 'OutlookAI-Indexed'
    Check 'the indexed guest is a test guest, resting on its frozen checkpoint' 'OAI-INDEXED|True|CP-24C-FROZEN-CLOCK' @($i.ComputerName, $i.Indexed, $i.Checkpoint)
    $u = Get-GuestFacts 'outlookai-unindexed'
    Check 'the unindexed guest is one, in any case' 'OutlookAI-Unindexed|OAI-UNINDEXED|False' @($u.Name, $u.ComputerName, $u.Indexed)
    Check 'the unindexed guest rests on its frozen checkpoint' 'CP-18B-FROZEN-CLOCK' $u.Checkpoint
    Write-Host '== their clones (TestbedPool.ps1, 2026-11-17) =='
    Check 'every source of the two guest pools is a guest here, and nothing else is but the Exchange VM' ((@($TestbedPools['indexed'][0], $TestbedPools['unindexed'][0], 'OutlookAI-Exchange') | Sort-Object) -join '|') ((@($Guests.Keys) | Sort-Object) -join '|')
    foreach ($kindName in @('Indexed', 'Unindexed')) {
        $members = @(Get-KindMembers $kindName)
        $source = Get-GuestFacts $members[0]
        Check "-Kind $kindName tries the source first" $source.Name $members[0]
        foreach ($clone in @($members | Select-Object -Skip 1)) {
            $f = Get-GuestFacts $clone
            Check "$clone is a test guest: its source's facts under its own name" "$clone|$($source.Name)|$kindName|$($source.ComputerName)|$($source.Indexed)|$($source.Checkpoint)" ("{0}|{1}|{2}|{3}|{4}|{5}" -f $f.Name, $f.Source, $f.Kind, $f.ComputerName, $f.Indexed, $f.Checkpoint)
        }
    }
    $c = Get-GuestFacts 'outlookai-indexed3'
    Check 'a clone named in another case is the clone, spelled as the pool spells it' 'OutlookAI-Indexed3|OutlookAI-Indexed' "$($c.Name)|$($c.Source)"
    Check 'a clone the pool does not list is refused' $true ($null -eq (Get-GuestFacts 'OutlookAI-Indexed9'))
    Check 'a build clone is refused' $true ($null -eq (Get-GuestFacts 'OutlookAI-Build2'))
    $xe = Get-GuestFacts 'OutlookAI-Exchange'
    Check 'the Exchange VM is a guest of its own, no pool''s and never cloned' 'OutlookAI-Exchange|OutlookAI-Exchange|Exchange|True' "$($xe.Name)|$($xe.Source)|$($xe.Kind)|$($xe.Exchange)"
    $threw = $false; try { [void](Get-KindMembers 'Exchange') } catch { $threw = $true }
    Check 'there is no third kind' $true $threw
    $runner = [System.IO.File]::ReadAllText($PSCommandPath)
    Check 'a clone''s settings are rendered from its source''s section' $true $runner.Contains("'New-LiveTestSettings.ps1') -VMName `$facts.Source")
    Check 'a lease is taken in one step with reading it' $true ($runner.Contains('$claim = Enter-TestbedLeaseClaim -VMName $Name') -and $runner.Contains('$holder = Request-GuestLease $guest.Name 60'))
    Write-Host '== a frozen start (Q130 (a)) =='
    Check 'a frozen start is verified and handed Outlook as 9a would, not restarted or rebuilt' 'frozen-clock guard|Outlook NOT elevated' (Get-PrepareSteps -StartIsFrozen $true -SkipHubReset $false)
    Check 'and -SkipHubReset changes nothing there' 'frozen-clock guard|Outlook NOT elevated' (Get-PrepareSteps -StartIsFrozen $true -SkipHubReset $true)
    $hubReset = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\guest\Reset-HubPopulation.ps1'))
    Check 'the profile a frozen start opens is the one step 9a opens' $true $hubReset.Contains("[string] `$TierProfileName = '$TierProfileName'")
    Check 'an unfrozen start is restarted and its hub rebuilt' 'restart|hub rebuild' (Get-PrepareSteps -StartIsFrozen $false -SkipHubReset $false)
    Check 'an unfrozen start with -SkipHubReset is only restarted' 'restart' (Get-PrepareSteps -StartIsFrozen $false -SkipHubReset $true)
    Write-Host '== a signature setting before Outlook starts (Q200 (a), Q201) =='
    Check 'a frozen start makes the setting between the clock guard and Outlook''s start' 'frozen-clock guard|signature setting GroupPolicy, before Outlook starts|Outlook NOT elevated' (Get-PrepareSteps -StartIsFrozen $true -SkipHubReset $false -SignatureSettings 'GroupPolicy')
    Check 'no step asked for is no refusal' '' (Get-SignatureSettingsRefusal '' $true $false 'G' 'F')
    Check 'a frozen start of an Outlook guest may take it' '' (Get-SignatureSettingsRefusal 'Preference' $false $true '' '')
    Check 'refused on the Exchange VM' $true ((Get-SignatureSettingsRefusal 'GroupPolicy' $true $true '' '') -like '*Exchange VM*')
    Check 'refused with a green checkpoint, which would keep it' $true ((Get-SignatureSettingsRefusal 'GroupPolicy' $false $true 'CP-GREEN' '') -like '*-GreenCheckpoint*')
    Check 'refused with -Filter' $true ((Get-SignatureSettingsRefusal 'GroupPolicy' $false $true '' 'Category=Live') -like '*-Filter*')
    Check 'refused from a start that is not frozen' $true ((Get-SignatureSettingsRefusal 'GroupPolicy' $false $false '' '') -like '*not frozen*')
    $f = Get-LiveFilter $false '' '' $false 'GroupPolicy'
    Check 'the run is narrowed to the test that measures what Outlook inserts' "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange&Requires!=SearchIndex&FullyQualifiedName~$SignatureInsertsTest|True" @($f.Filter, $f.Narrowed)
    Check 'and with -SignatureDialog to the dialog''s test' $true ((Get-LiveFilter $false '' '' $false 'GroupPolicy' $true).Filter.EndsWith("&FullyQualifiedName~$SignatureDialogTest"))
    Check 'and a suffix narrows it further' $true ((Get-LiveFilter $true '' '&Requires=Transport' $false 'Preference').Filter.EndsWith("&Requires=Transport&FullyQualifiedName~$SignatureInsertsTest"))
    Check '-SignatureDialog without -SignatureSettings is refused' $true ((Get-SignatureSettingsRefusal '' $false $true '' '' $true) -like '*-SignatureDialog without*')
    $settingsTests = [System.IO.File]::ReadAllText((Join-Path $repo 'McpServer\OutlookAI.McpServer.Tests\T2\LiveSignatureSettingsTests.cs'))
    Check 'the tests the filters name exist' $true ($settingsTests.Contains('public void ' + ($SignatureInsertsTest.Split('.')[1])) -and $settingsTests.Contains('public void ' + ($SignatureDialogTest.Split('.')[1])))
    $stepCs = [System.IO.File]::ReadAllText((Join-Path $repo 'McpServer\OutlookAI.McpServer.Tests\T2\LiveSignatureSettingsStep.cs'))
    Check 'the suite reads the scenario from the variable this sets' $true ($stepCs.Contains("`"$SignatureSettingsVariable`"") -and $runner.Contains("`$SignatureSettingsVariable = '$SignatureSettingsVariable'") -and $runner.Contains('`$env:$SignatureSettingsVariable = '))
    Check 'and knows the scenarios this takes' $true $stepCs.Contains('Scenarios { get; } = new[] { GroupPolicy, Preference, PreferenceNoOwnNew, NoSetting }')
    $settingsScript = [System.IO.File]::ReadAllText((Join-Path $repo "Testbed\guest\$SignatureSettingsScript"))
    Check 'the guest script takes the same scenarios' $true ($settingsScript.Contains("[ValidateSet('GroupPolicy', 'Preference', 'PreferenceNoOwnNew', 'NoSetting')] [string] `$Scenario") -and $runner.Contains("[ValidateSet('GroupPolicy', 'Preference', 'PreferenceNoOwnNew', 'NoSetting')] [string] `$SignatureSettings"))
    Check 'and the same value kinds, with the same default' $true ($settingsScript.Contains("[ValidateSet('ExpandString', 'String')] [string] `$ValueKind = 'ExpandString'") -and $runner.Contains("[ValidateSet('ExpandString', 'String')] [string] `$SignatureValueKind = 'ExpandString'") -and $runner.Contains("-Scenario `$SignatureSettings -ValueKind `$SignatureValueKind -Execute"))
    # The LAST of each: the run's own code comes after this self-test, which names both.
    $stepAt = $runner.LastIndexOf("Invoke-Guest 'signature-settings'")
    $frozenStartAt = $runner.LastIndexOf("-Profile '`$TierProfileName'; ")
    Check 'the setting is made before the frozen start''s Outlook starts' $true ($stepAt -gt 0 -and $frozenStartAt -gt $stepAt)
    Check 'no step, no dialog reading' '' (Get-SignatureDialogLines '' 'C:\G' 'C:\R')
    Check 'a step without -SignatureDialog leaves the dialog unopened' '' (Get-SignatureDialogLines 'GroupPolicy' 'C:\G' 'C:\R' $false)
    $dialogLines = Get-SignatureDialogLines 'GroupPolicy' 'C:\G' 'C:\R'
    Check 'a step reads the dialog into the run''s folder' $true $dialogLines.Contains("& 'C:\G\$SignatureDialogScript' -OutPath 'C:\R\$SignatureDialogReading' *> 'C:\R\signature-dialog.log'")
    Check 'and tells the suite where' $true $dialogLines.Contains("`$env:$SignatureDialogVariable = 'C:\R\$SignatureDialogReading'")
    Check 'the suite reads the dialog from the variable this sets' $true $stepCs.Contains("`"$SignatureDialogVariable`"")
    Check 'the probe is a guest script' $true (Test-Path -LiteralPath (Join-Path $repo "Testbed\guest\$SignatureDialogScript"))
    Check 'the dialog is read before the suite starts' $true ($runner.LastIndexOf('$($dialogLines)Set-Location') -gt 0 -and $runner.LastIndexOf('$($dialogLines)Set-Location') -lt $runner.LastIndexOf('& dotnet test McpServer'))
    Check 'and its reading is fetched with the results' $true $runner.Contains("-Include '*.trx', 'console.txt', 'signature-dialog.*'")
    $hostStart = [DateTime]::SpecifyKind([DateTime]::new(2026, 10, 3, 21, 0, 0), [DateTimeKind]::Utc)
    Check 'crash events are searched from the GUEST''s clock - a frozen guest''s is hours behind the host''s' '2026-10-03T14:50:00.0000000Z' (Get-CrashEventsSince -GuestRunStartUtc '2026-10-03T14:51:00.0000000Z' -HostRunStartUtc $hostStart)
    Check 'and from the host''s only when the guest''s could not be read' '2026-10-03T20:59:00.0000000Z' (Get-CrashEventsSince -GuestRunStartUtc '' -HostRunStartUtc $hostStart)
    $own = [System.IO.File]::ReadAllText($PSCommandPath)
    Check 'the crash search takes the guest''s clock read before the suite' $true $own.Contains("`$since = Get-CrashEventsSince -GuestRunStartUtc `$guestRunStartUtc")
    Check 'the frozen branch calls the guard with -Verify' $true $own.Contains("'Set-GuestClockFrozen.ps1') -VMName `$facts.Name -Verify")
    $x = Get-GuestFacts 'OutlookAI-Exchange'
    Check 'the Exchange VM is one, never frozen, resting on its newest checkpoint' 'OAI-EXCHANGE|True|True|CP-10-SDK11' @($x.ComputerName, $x.Indexed, $x.Exchange, $x.Checkpoint)
    Check 'the PST guests are not' 'False|False' @((Get-GuestFacts 'OutlookAI-Indexed').Exchange, (Get-GuestFacts 'OutlookAI-Unindexed').Exchange)
    Check 'the build VM is refused' $true ($null -eq (Get-GuestFacts 'OutlookAI-Build'))
    Check 'the retired test VM is refused' $true ($null -eq (Get-GuestFacts 'OutlookAI-TestVM'))
    Check 'a padded name is refused' $true ($null -eq (Get-GuestFacts ' OutlookAI-Indexed'))
    Check 'a prefix is refused' $true ($null -eq (Get-GuestFacts 'OutlookAI-Index'))
    Check 'nothing is refused' $true ($null -eq (Get-GuestFacts ''))

    Write-Host '== the filter =='
    $f = Get-LiveFilter $true '' ''
    Check 'the indexed guest runs LiveRunFilters.Guest, not narrowed' 'Category=Live&Requires!=DelegateStore&Requires!=CachedExchange|False' @($f.Filter, $f.Narrowed)
    Check 'the unindexed guest also leaves SearchIndex out' 'Category=Live&Requires!=DelegateStore&Requires!=CachedExchange&Requires!=SearchIndex' (Get-LiveFilter $false '' '').Filter
    $f = Get-LiveFilter $true '' '&FullyQualifiedName~X'
    Check 'a suffix narrows it' 'Category=Live&Requires!=DelegateStore&Requires!=CachedExchange&FullyQualifiedName~X|True' @($f.Filter, $f.Narrowed)
    $f = Get-LiveFilter $true '' '' $true
    Check 'the Exchange VM runs LiveRunFilters.ExchangeTestAccount' $FilterExchange $f.Filter
    Check 'read and write again since decision 206 (a)' $false $FilterExchange.Contains('Writes=')
    Check 'and that run is not narrowed' $false $f.Narrowed
    $f = Get-LiveFilter $true '' '&Writes=Nothing' $true
    Check 'past Office''s grace a suffix narrows it to LiveRunFilters.ExchangeTestAccountWritesNothing' ($FilterExchange + '&Writes=Nothing|True') @($f.Filter, $f.Narrowed)
    Check 'the Exchange VM carries its state forward by default' 'resume as the last run left it|Outlook NOT elevated if it is not running|preflight' (Get-ExchangeStartPlan $false)
    Check 'and a named checkpoint is restored and its data files set aside' $true ((Get-ExchangeStartPlan $true) -contains 'set the .ost files aside')
    Check 'with the search index rebuilt after them, and its roots waited for after Outlook starts' 'set the .ost files aside|rebuild the search index|Outlook NOT elevated|wait for one index root per indexed store' ((@(Get-ExchangeStartPlan $true) | Select-Object -Skip 2 -First 4) -join '|')
    Check 'and the runner asks the reset script for both' $true ($runner.Contains("Reset-OutlookOfflineStore.ps1' -Execute -ResetSearchIndex") -and $runner.Contains("Reset-OutlookOfflineStore.ps1' -WaitForIndexRoots"))
    $os = Get-OutlookStartFacts "  ACTIVATION-WIZARD-DISMISSED 8 s: class NUIDialog, by its Close button`r`nOK: OUTLOOK.EXE pid 9548`r`nACTIVATION-WIZARD-DISMISSALS 1`r`nSTART-OUTLOOK-EXIT 0" 'Outlook'
    Check 'a start that closed the activation wizard records one dismissal' 1 $os['activationWizardDismissed']
    Check 'and the summary line says so' "outlook    started by this run, NOT elevated, on profile 'Outlook'; Office's activation wizard dismissed 1 time(s)" (Format-OutlookStartLine $os)
    Check 'a starter that reported nothing reads as not reported, never as zero' $true ((Format-OutlookStartLine (Get-OutlookStartFacts 'OK: OUTLOOK.EXE pid 1' 'X')) -match 'NOT REPORTED')
    Check 'no start at all reads as not started' 'outlook    not started by this run (it was running already)' (Format-OutlookStartLine $null)
    Check 'and a suffix narrows it' ($FilterExchange + '&FullyQualifiedName~Y') (Get-LiveFilter $true '' '&FullyQualifiedName~Y' $true).Filter
    Check 'a suffix that does not start with & is refused' $true ($null -ne (Get-LiveFilter $true '' 'FullyQualifiedName~X').Refusal)
    Check 'a filter that is not Category=Live is refused' $true ($null -ne (Get-LiveFilter $true 'FullyQualifiedName~X' '').Refusal)
    $f = Get-LiveFilter $false 'Category=Live&Requires=SearchIndex' ''
    Check 'a whole filter is taken as given, and narrows' 'Category=Live&Requires=SearchIndex|True' @($f.Filter, $f.Narrowed)
    Check 'filter and suffix together are refused' $true ($null -ne (Get-LiveFilter $true 'Category=Live' '&X').Refusal)
    Check 'a double quote is refused' $true ($null -ne (Get-LiveFilter $true '' '&FullyQualifiedName~"x').Refusal)
    Check 'a dollar sign is refused' $true ($null -ne (Get-LiveFilter $true '' '&FullyQualifiedName~$env:x').Refusal)
    Check 'a backtick is refused' $true ($null -ne (Get-LiveFilter $true 'Category=Live`n' '').Refusal)

    Write-Host '== the lease =='
    $now = 1800000000L
    $mine = "$LeaseReasonPrefix r1"
    Check 'a run''s lease reason is this runner''s name and the run id' 'Invoke-LiveTierOnGuest r1' $mine
    Check 'no lease is free' 'free' (Get-LeaseDecision $null $mine $now)
    Check 'an expired lease is free' 'free' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now - 1; reason = 'someone' }) $mine $now)
    Check 'an unreadable expiry is free' 'free' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = 'soon'; reason = 'someone' }) $mine $now)
    Check 'a live lease of this run is ours' 'ours' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now + 60; reason = 'Invoke-LiveTierOnGuest r1' }) $mine $now)
    Check 'a live lease of another run is theirs' 'theirs' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now + 60; reason = 'Invoke-LiveTierOnGuest r2' }) $mine $now)
    Check 'a live lease of another agent is theirs' 'theirs' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now + 60; reason = 'agent accec8f1: folder-path tests' }) $mine $now)
    Check 'every lease decision of the run asks with the run''s own reason' 4 ([regex]::Matches($runner, 'Get-LeaseDecision [^\r\n]*"\$LeaseReasonPrefix \$runId"')).Count

    Write-Host '== the gate (Q195 (e), TestbedLoad.ps1) and the queue (2026-11-18, TestbedPool.ps1): before the build, and on every turn of the lease wait, asking Hyper-V nothing =='
    # Looked for in the run's code only - after this function, whose own lines name the same text.
    $selfTestAst = @([System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null).FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SelfTest' }, $false))
    $runCode = ''
    if ($selfTestAst.Count -eq 1) { $runCode = $runner.Substring($selfTestAst[0].Extent.EndOffset) }
    Check 'the run''s code is found after this self-test' $true ($runCode.Length -gt 1000)
    $at = @{}
    # The run's code only: what follows the -SelfTest dispatch, below this function.
    $mainCode = $own.Substring($own.LastIndexOf('if ($SelfTest) { exit (Invoke-SelfTest) }'))
    $claimAt = $mainCode.IndexOf('$made = New-RunDirectory -Parent $ResultsRoot -RunId (New-RunId $started $sha (Get-RunLabel $facts.Name))')
    Check 'the run claims its directory with New-RunDirectory (ScriptHelpers.ps1), from the run id New-RunId makes, and goes by the name it gives' $true ($claimAt -gt 0 -and $mainCode.IndexOf('$runId = $made.RunId') -gt $claimAt)
    Check 'before its log, its build tree or its folder on the guest is named after it, and never with New-Item -Force' $true ($claimAt -gt 0 -and $claimAt -lt $mainCode.IndexOf("`$script:LogPath = Join-Path `$runDir 'run.log'") -and $claimAt -lt $mainCode.IndexOf("`$tree = Join-Path `$runDir 'tree'") -and $claimAt -lt $mainCode.IndexOf('$guestRunDir = Join-Path $GuestRunsRoot $runId') -and -not $mainCode.Contains('New-Item -ItemType Directory -Force -Path $runDir'))
    foreach ($mark in @('$gate = Wait-TestbedGate -Run $runId', "& (Join-Path `$PSScriptRoot 'Publish-GuestPayload.ps1')", '$look = Test-TestbedGate', 'if ($gated) { $why += $look.Why }', '$ticket = Join-TestbedQueue -Run $runId', '$ahead = Get-TestbedQueueTurn', '$held = Get-TestbedLease -VMName $guest.Name', '$refusals[$guest.Name] = Get-GuestRunRefusal $guest', '$startupBytes[$guest.Name] = [long](Get-VM -Name $guest.Name).MemoryStartup', '$start = Enter-TestbedStart -Look $look -Run $runId', 'if (-not $start.Go) { $gated = $true', '$holder = Request-GuestLease $guest.Name 60', 'Register-TestbedStart -Start $start', 'finally { Exit-TestbedStart $start }', '[void](Wait-TestbedChange)', 'finally { Exit-TestbedQueue $ticket; if ($null -ne $waitingSince) { Remove-TestbedWaiting } }', '# ---- STAGE.')) { $at[$mark] = $runCode.IndexOf($mark) }
    Check 'the gate is passed before the build - the run''s first load on this host' $true ($at['$gate = Wait-TestbedGate -Run $runId'] -gt 0 -and $at['$gate = Wait-TestbedGate -Run $runId'] -lt $at["& (Join-Path `$PSScriptRoot 'Publish-GuestPayload.ps1')"])
    Check 'every turn of the lease wait looks at it again before it tries a guest' $true ($at['$look = Test-TestbedGate'] -gt $at["& (Join-Path `$PSScriptRoot 'Publish-GuestPayload.ps1')"] -and $at['if ($gated) { $why += $look.Why }'] -gt $at['$look = Test-TestbedGate'] -and $at['if ($gated) { $why += $look.Why }'] -lt $at['$holder = Request-GuestLease $guest.Name 60'])
    Check 'the lease is taken only once the start is decided under the gate''s claim' $true ($at['$start = Enter-TestbedStart -Look $look -Run $runId'] -gt $at['$look = Test-TestbedGate'] -and $at['$start = Enter-TestbedStart -Look $look -Run $runId'] -lt $at['$holder = Request-GuestLease $guest.Name 60'] -and $at['if (-not $start.Go) { $gated = $true'] -lt $at['$holder = Request-GuestLease $guest.Name 60'])
    Check 'its start is written once the lease is its, and the claim let go on every path' $true ($at['Register-TestbedStart -Start $start'] -gt $at['$holder = Request-GuestLease $guest.Name 60'] -and $at['finally { Exit-TestbedStart $start }'] -gt $at['Register-TestbedStart -Start $start'])
    Check 'and the guest is touched only after all of it' $true ($at['finally { Exit-TestbedQueue $ticket; if ($null -ne $waitingSince) { Remove-TestbedWaiting } }'] -gt $at['finally { Exit-TestbedStart $start }'] -and $at['# ---- STAGE.'] -gt $at['finally { Exit-TestbedQueue $ticket; if ($null -ne $waitingSince) { Remove-TestbedWaiting } }'])
    Check 'the gate comes before the turn''s place in the queue, so a turn it holds back keeps the ticket and tries no guest' $true ($at['$look = Test-TestbedGate'] -lt $at['$ahead = Get-TestbedQueueTurn'] -and $at['$ahead = Get-TestbedQueueTurn'] -lt $at['$held = Get-TestbedLease -VMName $guest.Name'])
    Check 'a run waiting for a guest holds a ticket and reads its turn on every pass' $true ($at['$ticket = Join-TestbedQueue -Run $runId'] -gt 0 -and $at['$ahead = Get-TestbedQueueTurn'] -gt $at['$ticket = Join-TestbedQueue -Run $runId'])
    Check 'whether a guest is free is its lease file, read before Hyper-V is asked anything about it' $true ($at['$held = Get-TestbedLease -VMName $guest.Name'] -gt $at['$ahead = Get-TestbedQueueTurn'] -and $at['$held = Get-TestbedLease -VMName $guest.Name'] -lt $at['$refusals[$guest.Name] = Get-GuestRunRefusal $guest'] -and $at['$refusals[$guest.Name] = Get-GuestRunRefusal $guest'] -lt $at['$startupBytes[$guest.Name] = [long](Get-VM -Name $guest.Name).MemoryStartup'])
    Check 'and what does not change in a run - its checkpoints, its startup memory - is asked once a run' $true ($runCode.Contains('if (-not $refusals.ContainsKey($guest.Name))') -and $runCode.Contains('if (-not $startupBytes.ContainsKey($guest.Name))'))
    Check 'between passes it sleeps on the testbed''s files, and leaves the queue on every path' $true ($at['[void](Wait-TestbedChange)'] -gt $at['$holder = Request-GuestLease $guest.Name 60'] -and $at['finally { Exit-TestbedQueue $ticket; if ($null -ne $waitingSince) { Remove-TestbedWaiting } }'] -gt $at['[void](Wait-TestbedChange)'])
    # TestbedPool.ps1's functions, which the run asks Hyper-V through, by name.
    $poolAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'TestbedPool.ps1'), [ref] $null, [ref] $null)
    $poolBody = @{}
    foreach ($fn in $poolAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) { $poolBody[$fn.Name] = $fn.Body.Extent.Text }
    $poolWait = [string]$poolBody['Wait-TestbedHeartbeat']
    $poolRestore = [string]$poolBody['Restore-TestbedCheckpoint']
    Check 'the heartbeat is one narrow query on the guest''s id, every 2 s' $true ($runCode.Contains('Wait-TestbedHeartbeat -Id (Get-TestbedVmId $facts.Name)') -and $poolRestore.Contains('Wait-TestbedHeartbeat -Id (Get-TestbedVmId $Name)') -and $poolWait.Contains('Test-TestbedHeartbeat $Id') -and $poolWait.Contains('Start-Sleep -Seconds $TestbedHeartbeatSeconds') -and -not $runCode.Contains("Get-VMIntegrationService -VMName `$facts.Name -Name 'Heartbeat'"))
    Check 'the run''s guest calls share one PowerShell Direct session, opened by the guest''s id' $true ($runCode.Contains('New-PSSession -VMId (Get-TestbedVmId $facts.Name)') -and -not $runCode.Contains('Invoke-Command -VMName') -and ([regex]::Matches($runCode, 'Invoke-Command -Session ')).Count -eq 2 -and $runCode.Contains('$out = Invoke-Command -Session $session -ScriptBlock'))
    Check 'and each guest call starts as a session of its own did: no exit code from the call before, the folder it started in' $true ($runCode.Contains('$global:LASTEXITCODE = $null') -and $runCode.Contains('Set-Location -LiteralPath $global:OutlookAIGuestStart') -and $runCode.IndexOf('$global:LASTEXITCODE = $null') -lt $runCode.IndexOf('"RUN-REGISTER-EXIT: $LASTEXITCODE"'))
    $restoreAt = $runCode.IndexOf('function Restore-Guest')
    $restoreText = ''
    if ($restoreAt -ge 0) { $restoreText = $runCode.Substring($restoreAt, [Math]::Min(700, $runCode.Length - $restoreAt)) }
    Check 'a running guest is saved before a checkpoint is restored over it, and its session closed' $true ($restoreText.IndexOf('Close-GuestSession') -gt 0 -and $restoreText.IndexOf('Restore-TestbedCheckpoint -Name $facts.Name -Checkpoint $Name -Start') -gt $restoreText.IndexOf('Close-GuestSession') -and $poolRestore.IndexOf('Save-VM -Name $Name') -gt 0 -and $poolRestore.IndexOf('Save-VM -Name $Name') -lt $poolRestore.IndexOf('Restore-VMSnapshot'))
    Check 'and the session is closed before every restart, which ends it' $true (([regex]::Matches($runCode, 'Close-GuestSession\s+# ps51-native-stderr-ok:[^\n]*\n\s*&\s+\(Join-Path \$PSScriptRoot ''Restart-Guest\.ps1''\)')).Count -eq 2)
    Check 'a guest''s state is read with the narrow query, never Get-VM''s' $true ($runCode.Contains('(Get-TestbedVmState $facts.Name)') -and ([string]$poolBody['Get-TestbedVmState']).Contains('Get-TestbedVmStateById') -and -not $runCode.Contains('(Get-VM -Name $facts.Name).State'))

    Write-Host '== the run id =='
    Check 'when, guest, commit' '20261003-174500-indexed-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' (Get-RunLabel 'OutlookAI-Indexed'))
    Check 'the unindexed guest' '20261003-174500-unindexed-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' (Get-RunLabel 'OutlookAI-Unindexed'))
    Check 'the Exchange VM' '20261003-174500-exchange-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' (Get-RunLabel 'OutlookAI-Exchange'))
    Check 'and its clone' '20261003-174500-unindexed-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' (Get-RunLabel 'OutlookAI-Unindexed3'))
    Check 'and the indexed guest''s clone' '20261003-174500-indexed-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' (Get-RunLabel 'OutlookAI-Indexed3'))
    $threw = $false; try { New-RunId (Get-Date) 'HEAD' (Get-RunLabel 'OutlookAI-Indexed') | Out-Null } catch { $threw = $true }

    Write-Host '== the run marker and the ledger (the Exchange VM) =='
    $when = [DateTime]::SpecifyKind([DateTime]::new(2026, 10, 8, 15, 30, 5), [DateTimeKind]::Utc)
    Check 'the marker is oair and the UTC start to the second' 'oair20261008153005' (New-RunMarker $when)
    Check 'and has the shape LiveRunMarker accepts' $true ((New-RunMarker $when) -cmatch '^[a-z0-9]{12,40}$')
    $ledger = @(
        (ConvertTo-LedgerLine 'oair20261008100000' 'r1' 'started' $null @() $when),
        (ConvertTo-LedgerLine 'oair20261008100000' 'r1' 'finished' $false @() $when),
        (ConvertTo-LedgerLine 'oair20261008110000' 'r2' 'started' $null @('oair20261008100000') $when),
        (ConvertTo-LedgerLine 'oair20261008120000' 'r3' 'started' $null @() $when),
        'not json',
        (ConvertTo-LedgerLine 'BAD' 'r4' 'started' $null @() $when)
    )
    Check 'every run that never finished clean is a prior, newest first' 'oair20261008120000|oair20261008110000|oair20261008100000' (Get-PriorRunMarkers $ledger 'oair20261008130000')
    Check 'the current run is never its own prior' 'oair20261008110000|oair20261008100000' (Get-PriorRunMarkers $ledger 'oair20261008120000')
    $cleanLine = ConvertTo-LedgerLine 'oair20261008120000' 'r3' 'finished' $true @('oair20261008100000', 'oair20261008110000') $when
    Check 'a clean finish clears itself and the priors it swept' 0 @(Get-PriorRunMarkers ($ledger + $cleanLine) 'oair20261008130000').Count
    Check 'at most -Max priors' 1 @(Get-PriorRunMarkers $ledger 'x' -Max 1).Count
    Check 'an empty ledger has none' 0 @(Get-PriorRunMarkers @() 'x').Count
    Check 'a ref that is not a commit id is refused' $true $threw

    Write-Host '== the TRX reader =='
    $ns = 'xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010"'
    $trx = ConvertFrom-TrxText ("<TestRun $ns><Results>" +
        '<UnitTestResult testName="A.Pass" outcome="Passed" />' +
        '<UnitTestResult testName="A.Fail" outcome="Failed"><Output><ErrorInfo><Message>boom' + "`n" + 'second</Message></ErrorInfo></Output></UnitTestResult>' +
        '<UnitTestResult testName="A.Skip" outcome="NotExecuted" /></Results></TestRun>')
    Check 'counts' '3|1|1|1|0' @($trx.Counts['total'], $trx.Counts['passed'], $trx.Counts['failed'], $trx.Counts['skipped'], $trx.Counts['other'])
    Check 'a failure keeps its first message line' 'A.Fail: boom' ("$($trx.Failed[0].name): $($trx.Failed[0].message)")
    $green = ConvertFrom-TrxText ("<TestRun $ns><Results><UnitTestResult testName=`"A.Pass`" outcome=`"Passed`" /></Results></TestRun>")
    $empty = ConvertFrom-TrxText ("<TestRun $ns><Results></Results></TestRun>")

    Write-Host '== the console reader =='
    $clean = Read-ConsoleFacts ((@(
        '[xUnit.net 00:00:50.26]         post-suite: 0 tagged artifacts (incl. Archive), 0 test folders',
        '[xUnit.net 00:06:14.41]         PROVED NOTHING: the retry-guidance check iterated nothing - here.',
        '  PROVED NOTHING: the retry-guidance check iterated nothing - here.',
        'PROVED NOTHING: step 3b iterated nothing - here.',
        "[xUnit.net 00:08:10.68]             artifact sweep: 4 store(s) - 2 swept, 2 counted and left alone ('bystander@vm.invalid', 'Corpus A')",
        '[xUnit.net 00:08:10.68]             sweep[tier@vm.invalid]: taggedArtifacts=0',
        '  sweep[Corpus A]: taggedArtifacts=0 - COUNTED, NOT SWEPT (declared BYSTANDER)',
        "[tripwire] post-run census of 'tier@vm.invalid': 18 mail folder(s), 18 folder(s) measured, 0 folder(s), 0 item(s) identified, 244 ms.",
        '[tripwire] post-run census in 1331 ms (identified 11 folder(s)/308 item(s)); 0 failure(s), 1 note(s); retry: none needed.'
    )) -join "`r`n")
    Check 'PROVED NOTHING lines are counted once each' 2 $clean.ProvedNothing.Count
    Check 'a sweep ran, and every store ended at zero' 'True|0' @($clean.SweepRan, $clean.ArtifactsLeft)
    Check 'the census reported no failure' 0 $clean.Tripwire
    $dirty = Read-ConsoleFacts ("artifact sweep: 2 store(s) - 2 swept, 0 counted and left alone`nsweep[a@b]: taggedArtifacts=0`nsweep[c@d]: taggedArtifacts=2`n[tripwire] post-run census in 9 ms (identified 0 folder(s)/0 item(s)); 3 failure(s), 0 note(s); retry: none needed.")
    Check 'the worst store counts' 2 $dirty.ArtifactsLeft
    Check 'census failures are read' 3 $dirty.Tripwire
    $silent = Read-ConsoleFacts 'nothing to see'
    Check 'no sweep is unknown, not zero' 'False|True' @($silent.SweepRan, ($null -eq $silent.ArtifactsLeft))
    Check 'no census is unknown, not zero' $true ($null -eq $silent.Tripwire)

    Write-Host '== the verdict =='
    Check 'green tests and clean proofs are PASS' 'PASS' (Get-LiveVerdict '' '' $green $clean 0 0 $false).Verdict
    Check 'a failed test is FAIL' 'FAIL' (Get-LiveVerdict '' '' $trx $clean 0 1 $false).Verdict
    Check 'an Outlook crash fails a green run' 'FAIL' (Get-LiveVerdict '' '' $green $clean 1 0 $false).Verdict
    Check 'an artifact left fails a green run' 'FAIL' (Get-LiveVerdict '' '' $green $dirty 0 0 $false).Verdict
    Check 'a full run with no sweep is FAIL - not proven' 'FAIL' (Get-LiveVerdict '' '' $green $silent 0 0 $false).Verdict
    Check 'a narrowed run with no sweep can PASS' 'PASS' (Get-LiveVerdict '' '' $green $silent 0 0 $true).Verdict
    Check 'a non-zero dotnet exit over a green TRX is FAIL' 'FAIL' (Get-LiveVerdict '' '' $green $clean 0 1 $false).Verdict
    Check 'an unknown dotnet exit decides nothing' 'PASS' (Get-LiveVerdict '' '' $green $clean 0 $null $false).Verdict
    Check 'a filter that selected nothing is FAIL' 'FAIL' (Get-LiveVerdict '' '' $empty $clean 0 0 $true).Verdict
    Check 'no TRX after the suite ran is FAIL' 'FAIL' (Get-LiveVerdict '' '' $null $clean 0 0 $false).Verdict
    Check 'unreadable crash events do not fail a green run' 'PASS' (Get-LiveVerdict '' '' $green $clean -1 0 $false).Verdict
    Check 'a stage that broke is BUILD' 'BUILD' (Get-LiveVerdict 'BUILD' 'the payload did not build' $null $null -1 $null $false).Verdict
    Check 'a run that never reached the suite is INFRA' 'INFRA' (Get-LiveVerdict 'INFRA' 'the lease stayed taken' $null $null -1 $null $false).Verdict
    Check 'an unfinished run still fails on what failed' 'FAIL' (Get-LiveVerdict 'INFRA' 'the guest went away' $trx $null -1 $null $false).Verdict
    Check 'exit codes' '0|1|2|3|4' @($ExitCodes['PASS'], $ExitCodes['FAIL'], $ExitCodes['BUILD'], $ExitCodes['INFRA'], $ExitCodes['REFUSED'])

    Write-Host '== summary.json =='
    $one = New-LiveRunSummary -RunId 'r' -Verdict (Get-LiveVerdict '' '' $trx $clean 0 1 $false) -Revision ([ordered]@{ sha = 'x' }) -Guest 'OutlookAI-Indexed' -Filter 'f' -Narrowed $false -Trx $trx -Console $clean -Crashes 0 -SuiteExit 1 -Checkpoints ([ordered]@{}) -Timings ([ordered]@{}) -Results 'd'
    $json = $one | ConvertTo-Json -Depth 8
    $back = $json | ConvertFrom-Json
    Check 'one failed test is still an array' $true ($json -match '"failed":\s*\[')
    Check 'one skipped test is still an array' $true ($json -match '"skipped":\s*\[')
    Check 'the PROVED NOTHING lines are an array' $true ($json -match '"lines":\s*\[')
    Check 'the verdict, exit code and proofs round-trip' 'FAIL|1|True|True' @($back.verdict, $back.exitCode, $back.artifacts.proven, $back.tripwire.reported)
    Check 'durations read as minutes and seconds' '59 s|1 m 00 s|40 m 59 s' @((Format-Seconds 59.4), (Format-Seconds 60), (Format-Seconds 2459.9))
    $withDump = New-LiveRunSummary -RunId 'r' -Verdict (Get-LiveVerdict '' '' $green $clean 1 0 $false) -Revision ([ordered]@{ sha = 'x' }) -Guest 'OutlookAI-Indexed' -Filter 'f' -Narrowed $true -Trx $green -Console $clean -Crashes 1 -SuiteExit 0 -Checkpoints ([ordered]@{}) -Timings ([ordered]@{}) -Results 'd' -CrashDumps ([ordered]@{ armed = $true; files = @([ordered]@{ name = 'OUTLOOK.EXE.4242.dmp'; fetched = $true }) })
    $dumpJson = $withDump | ConvertTo-Json -Depth 8
    Check 'one crash dump is still an array' $true ($dumpJson -match '"files":\s*\[')
    Check 'and the dump state round-trips' 'True|OUTLOOK.EXE.4242.dmp' @(($dumpJson | ConvertFrom-Json).crashDumps.armed, @(($dumpJson | ConvertFrom-Json).crashDumps.files)[0].name)

    Write-Host '== crash dumps =='
    $d = @(ConvertFrom-DumpListing "WERFAULT-LEFT 0`r`nDUMP 1288490188 OUTLOOK.EXE.4242.dmp`r`nDUMP 52428800 WINWORD.EXE.17.dmp`r`nnoise DUMP 1 x.dmp`r`n")
    Check 'the guest''s listing is read line by line, the program from the name' 'OUTLOOK.EXE.4242.dmp|1288490188|OUTLOOK.EXE|WINWORD.EXE' @($d[0].Name, $d[0].Bytes, $d[0].Program, $d[1].Program)
    Check 'only whole DUMP lines count' 2 $d.Count
    Check 'no listing is no dump' 0 @(ConvertFrom-DumpListing 'WERFAULT-LEFT 0').Count
    Check 'an Outlook dump is a crash even when its event was not found' 1 (Merge-CrashCount 0 $d)
    Check 'the events count when they say more' 3 (Merge-CrashCount 3 $d)
    Check 'unreadable events with an Outlook dump are that dump' 1 (Merge-CrashCount -1 $d)
    Check 'unreadable events and no dump stay unreadable' -1 (Merge-CrashCount -1 @())
    Check 'a Word dump alone is no Outlook crash' 0 (Merge-CrashCount 0 @($d[1]))
    $w = @(ConvertFrom-WerListing "WERFAULT-LEFT 0`r`nDUMP 9 OUTLOOK.EXE.1.dmp`r`nWERFILE 31548 AppCrash_OUTLOOK.EXE_ab12_cd34_5e6f.Report.wer  `r`nWER-REPORTS 1`r`nnoise WERFILE 1 x.wer`r`n")
    Check 'WER''s report files are read line by line, name and size, and nothing else is' '1|AppCrash_OUTLOOK.EXE_ab12_cd34_5e6f.Report.wer|31548' @($w.Count, $w[0].Name, $w[0].Bytes)
    Check 'a WER file is no dump, and a dump no WER file' '1|1' @(@(ConvertFrom-DumpListing "DUMP 9 OUTLOOK.EXE.1.dmp`r`nWERFILE 3 a.Report.wer").Count, @(ConvertFrom-WerListing "DUMP 9 OUTLOOK.EXE.1.dmp`r`nWERFILE 3 a.Report.wer").Count)
    $saverAt = $runCode.IndexOf('function Save-CrashDumps {')
    $werReadAt = $runCode.IndexOf('$script:werFound = @(ConvertFrom-WerListing $o)', [Math]::Max(0, $saverAt))
    $noDumpAt = $runCode.IndexOf('if ($script:dumpsFound.Count -eq 0) { return }', [Math]::Max(0, $saverAt))
    Check 'WER''s reports of an Outlook crash are kept even when no dump was written' $true ($saverAt -gt 0 -and $saverAt -lt $werReadAt -and $werReadAt -lt $noDumpAt)
    $stores = @('ProgramData\Microsoft\Windows\WER\ReportQueue', 'ProgramData\Microsoft\Windows\WER\ReportArchive', 'AppData\Local\Microsoft\Windows\WER\ReportQueue', 'AppData\Local\Microsoft\Windows\WER\ReportArchive')
    Check 'and every WER store is searched, the machine''s and each user''s' 4 @($stores | Where-Object { $at = $runCode.IndexOf($_, [Math]::Max(0, $saverAt)); $at -gt $saverAt -and $at -lt $werReadAt }).Count
    $dumpScript = [System.IO.File]::ReadAllText((Join-Path $repo "Testbed\guest\$CrashDumpScript"))
    Check 'the guest script''s default folder is the one fetched' $true $dumpScript.Contains("[string]   `$DumpFolder = '$GuestDumpDir'")
    Check 'every run arms the dumps, with this checkout''s script' $true ($own.Contains("& '`$GuestRoot\`$CrashDumpScript' -Execute -DumpFolder '`$GuestDumpDir'") -and $own.Contains('Join-Path $PSScriptRoot "..\guest\$CrashDumpScript"'))
    $sample = "CRASH DUMP x
EXCEPTION  access violation (0xC0000005), flags 0x0, on thread 4242
           at 0x00007FF8A1B2C3D4  wwlib.dll+0x7BA1A (fn wwlib.dll+0x7B900 +0x11A)
           READING 0x0000000000000018 - a null (or near-null) pointer

REGISTERS  x
  rax 0x0

STACK      the faulting thread
  00  wwlib.dll+0x7BA1A (fn wwlib.dll+0x7B900 +0x11A)
  01  combase.dll!CoTaskMemFree+0x10
  02  ntdll.dll!RtlUserThreadStart+0x2C
  (3 frame(s); end of stack)
"
    $hl = @(Get-DumpHeadline $sample 2)
    Check 'the headline: the exception, where, how - and the top frames, no more than asked' 'access violation|wwlib.dll+0x7BA1A|READING|00  wwlib.dll+0x7BA1A|01  combase.dll!CoTaskMemFree+0x10|5' @($hl[0].Substring(11, 16), $hl[1].Substring(23, 17), $hl[2].Substring(0, 7), $hl[3].Substring(0, 21), $hl[4], $hl.Count)
    Check 'a report with no exception has no headline' 0 @(Get-DumpHeadline 'DUMP x').Count
    # Searched in the run's code alone: in the whole file the first of each text is in this check itself,
    # and there they stood in the right order whatever the run did - a check that could not fail
    # (found 2026-11-18).
    $fetchAt = $runCode.IndexOf('if ($suiteStarted) { Save-CrashDumps }')
    Check 'the dumps are fetched before the guest rests - the restore wipes them' $true ($fetchAt -gt 0 -and $fetchAt -lt $runCode.IndexOf('Restore-Guest $restOn'))

    Write-Host '== the hunt (-HuntFirstChance, 2026-11-18) =='
    $paramAst = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$null).ParamBlock.Parameters
    $huntSwitch = @($paramAst | Where-Object { $_.Name.VariablePath.UserPath -eq 'HuntFirstChance' })
    Check 'the hunt is a switch, off unless given - no other run has a debugger attached' 'switch|True' "$($huntSwitch[0].StaticType.Name.ToLowerInvariant() -replace 'switchparameter','switch')|$($null -eq $huntSwitch[0].DefaultValue)"
    Check 'its module is mso40uiwin32client.dll, the one both crashes of 2026-11-18 named' 'mso40uiwin32client.dll' $HuntModule
    $catcher = [System.IO.File]::ReadAllText((Join-Path $repo "Testbed\guest\$FirstChanceScript"))
    Check 'and the debugger''s own default is the same module' $true $catcher.Contains("[string]   `$Module = '$HuntModule'")
    Check 'it hunts the tests the crashes struck in: the typed-text tests, then the reply with a signature and the forward' 'typed|drafts' (@($HuntLegs.Keys) -join '|')
    Check 'every leg narrows the derived filter' $true (@($HuntLegs.Values | Where-Object { -not $_.StartsWith('&') }).Count -eq 0)
    Check 'a run without the switch is no hunt' '' (Get-HuntRefusal $false $true 'Category=Live' '&x' 'GroupPolicy' $true 0 0 120)
    Check 'never on the Exchange VM' $true ((Get-HuntRefusal $true $true '' '' '' $false 20 60 120) -match 'Exchange VM')
    Check 'never with a filter of its own' $true ((Get-HuntRefusal $true $false '' '&FullyQualifiedName~X' '' $false 20 60 120) -match 'own tests')
    Check 'never with a signature step' $true ((Get-HuntRefusal $true $false '' '' 'GroupPolicy' $false 20 60 120) -match 'two runs')
    Check 'its rounds and minutes are bounded, inside the run''s time limit' '1|1|1|' "$([int][bool](Get-HuntRefusal $true $false '' '' '' $false 0 60 120))|$([int][bool](Get-HuntRefusal $true $false '' '' '' $false 20 106 120))|$([int][bool](Get-HuntRefusal $true $false '' '' '' $false 20 0 120))|$(Get-HuntRefusal $true $false '' '' '' $false 20 60 120)"
    $hunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ typed = 'F&A'; drafts = 'F&(B|C)' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript
    $err = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($hunt, [ref]$null, [ref]$err)
    Check 'the hunt''s suite script parses' 0 @($err).Count
    Check 'it starts the debugger through the shell, hidden, with the module, its stop file and its log' $true ($hunt.Contains("Start-Process -FilePath 'powershell.exe'") -and $hunt.Contains('-WindowStyle Hidden') -and $hunt.Contains("'-Module', 'mso40uiwin32client.dll'") -and $hunt.Contains("`$stop = 'C:\r\first-chance.stop'") -and $hunt.Contains("'-Log', 'C:\r\first-chance.log'"))
    Check 'every leg runs in a dotnet test of its own, its TRX file named for the round and leg' $true ($hunt.Contains("`$legs += ,@('typed', 'F&A')") -and $hunt.Contains("`$legs += ,@('drafts', 'F&(B|C)')") -and $hunt.Contains("'trx;LogFileName=live-r{0:D2}-{1}.trx' -f `$round, `$leg[0]"))
    Check 'the rounds stop at the cap, the minutes or the first dump, and the debugger is stopped after' $true ($hunt.Contains('while ($done -lt 20 -and $clock.Elapsed.TotalMinutes -lt 60)') -and $hunt.Contains("Get-ChildItem -LiteralPath 'C:\d' -Filter '*chance.dmp'") -and $hunt.IndexOf('New-Item -ItemType File -Force -Path $stop') -gt $hunt.IndexOf('while ('))
    $ordinaryAt = $runCode.IndexOf('$liveScript = @"')
    $ordinaryEnd = $runCode.IndexOf('"@', [Math]::Max(0, $ordinaryAt))
    Check 'an ordinary run''s suite script starts no debugger' $true ($ordinaryAt -gt 0 -and $ordinaryEnd -gt $ordinaryAt -and -not $runCode.Substring($ordinaryAt, $ordinaryEnd - $ordinaryAt).Contains('FirstChance'))
    Check 'and the debugger reaches the guest only in a hunting run' $true ($runCode.Contains('if ($HuntFirstChance) { $guestScripts += [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\guest\$FirstChanceScript")) }') -and $runCode.Contains('$_.Name -ne $FirstChanceScript'))
    $r1 = [pscustomobject]@{ Counts = [ordered]@{ total = 2; passed = 1; failed = 1; skipped = 0; other = 0 }; Declared = [ordered]@{ total = 2; passed = 1; failed = 1; notExecuted = 0 }; Failed = @([pscustomobject]@{ name = 'T.A'; message = 'boom' }); Skipped = @() }
    $r2 = [pscustomobject]@{ Counts = [ordered]@{ total = 3; passed = 2; failed = 0; skipped = 1; other = 0 }; Declared = $null; Failed = @(); Skipped = @('T.C') }
    $m = Merge-TrxResults @([pscustomobject]@{ Label = 'live-r01-typed'; Result = $r1 }, [pscustomobject]@{ Label = 'live-r01-drafts'; Result = $r2 })
    Check 'a hunting run''s TRX files are one result: counts summed, each test named with its file' '5|3|1|1|[live-r01-typed] T.A|[live-r01-drafts] T.C|True' "$($m.Counts['total'])|$($m.Counts['passed'])|$($m.Counts['failed'])|$($m.Counts['skipped'])|$($m.Failed[0].name)|$($m.Skipped[0])|$($null -eq $m.Declared)"

    Write-Host '== the hunt''s variants (-HuntVariant, 2026-11-18, option (b)) =='
    $variantParam = @($paramAst | Where-Object { $_.Name.VariablePath.UserPath -eq 'HuntVariant' })
    Check 'the variant is Legs unless given - the tests the crashes struck in, as before' 'Legs' ([string]$variantParam[0].DefaultValue.Value)
    Check 'Legs runs those legs, and every other variant the one loop leg' 'typed|drafts|loop' "$(@((Get-HuntLegs 'Legs').Keys) -join '|')|$(@((Get-HuntLegs 'NoWindow').Keys) -join '|')"
    Check 'LegsOtherFolder runs the calendar step before the tests the crashes struck in' 'calendar|typed|drafts' (@((Get-HuntLegs 'LegsOtherFolder').Keys) -join '|')
    Check 'and tells the hunting test the step, CalendarOnly; Legs tells it nothing; a loop its own name' 'CalendarOnly||NoWindow' "$(Get-HuntTestVariant 'LegsOtherFolder')|$(Get-HuntTestVariant 'Legs')|$(Get-HuntTestVariant 'NoWindow')"
    $folderHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ calendar = 'F&L'; typed = 'F&A'; drafts = 'F&(B|C)' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'LegsOtherFolder' -LegMinutes 1
    Check 'whose suite script names the step to the test and runs the step first each round' $true ($folderHunt.StartsWith("`$env:OUTLOOKAI_HUNT_VARIANT = 'CalendarOnly'`r`n") -and $folderHunt.IndexOf("`$legs += ,@('calendar'") -lt $folderHunt.IndexOf("`$legs += ,@('typed'"))
    Check 'the loop leg is the hunting test''s class' '&FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveHuntDerivedDraftTests' ((Get-HuntLegs 'BareCreators')['loop'])
    Check 'a variant without the hunt is refused' $true ((Get-HuntRefusal $false $false '' '' '' $false 20 60 120 'NoWindow' 10) -match 'give -HuntFirstChance')
    Check 'a leg is 1 to -HuntMinutes minutes' '1|1|' "$([int][bool](Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 0))|$([int][bool](Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 61))|$(Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 60)"
    $loopHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ loop = 'F&L' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'NoFormPage' -LegMinutes 12
    $err = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($loopHunt, [ref]$null, [ref]$err)
    Check 'a variant''s suite script parses' 0 @($err).Count
    Check 'it names the variant and its leg minutes to the test before anything runs' $true ($loopHunt.StartsWith("`$env:OUTLOOKAI_HUNT_VARIANT = 'NoFormPage'`r`n`$env:OUTLOOKAI_HUNT_LEG_MINUTES = '12'`r`n`$env:OUTLOOKAI_HUNT_SHOW_EACH = '0'`r`n") -and $loopHunt.Contains("`$legs += ,@('loop', 'F&L')"))
    $shownHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ loop = 'F&L' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'ProductLoop' -LegMinutes 10 -ShowEach $true
    Check 'the hunt''s minutes are counted on a stopwatch, not the guest''s clock, which can jump (2026-11-18)' $true ($hunt.Contains('$clock = [System.Diagnostics.Stopwatch]::StartNew()') -and -not $hunt.Contains('(Get-Date)'))
    $typedHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ loop = 'F&L' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'ProductLoop' -LegMinutes 10 -ShowEach $true -TypeEach $true
    Check '-HuntTypeEach tells it to type into the window it shows, and needs -HuntShowEach' 'typed|both' "$(if ($typedHunt.Contains("`$env:OUTLOOKAI_HUNT_SHOW_EACH = 'typed'`r`n")) { 'typed' })|$(if ((Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 10 $false $false $false $true) -match 'give both') { 'both' })"
    $ablateHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ typed = 'F&A'; drafts = 'F&(B|C)' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'Legs' -LegMinutes 10 -Ablate 'layout'
    Check '-HuntAblate names the step the typed-text tests leave out, on the legs only' 'layout|loop|hunt' "$(if ($ablateHunt.StartsWith("`$env:OUTLOOKAI_HUNT_ABLATE = 'layout'`r`n")) { 'layout' })|$(if ((Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 10 $false $false $false $false 'layout') -match 'Legs or LegsOtherFolder') { 'loop' })|$(if ((Get-HuntRefusal $false $false '' '' '' $false 20 60 120 'Legs' 10 $false $false $false $false 'layout') -match 'give -HuntFirstChance') { 'hunt' })"
    $promoteHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ loop = 'F&L' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'ProductLoop' -LegMinutes 10 -ShowEach $true -TypeEach $true -PromoteAfter $true
    Check '-HuntPromoteAfter tells the loop to re-open and promote the product''s drafts, on a loop that makes them' 'promote|bare' "$(if ($promoteHunt.Contains("`$env:OUTLOOKAI_HUNT_PROMOTE_AFTER = '1'`r`n")) { 'promote' })|$(if ((Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'BareCreators' 10 $false $false $false $false '' $false $true) -match 'makes them') { 'bare' })"
    $lateHunt = Get-HuntSuiteScript -OptIn 'OAI-INDEXED' -GuestRunDir 'C:\r' -GuestRoot 'C:\g' -Filters ([ordered]@{ loop = 'F&L' }) -Rounds 20 -Minutes 60 -Module 'mso40uiwin32client.dll' -DumpDir 'C:\d' -ScriptName $FirstChanceScript -Variant 'ProductLoop' -LegMinutes 10 -ShowEach $true -TypeEach $true -ReleaseAfterClose $true
    Check '-HuntReleaseAfterClose tells the loop to release the typed window''s document after its close, and needs -HuntTypeEach' 'late|typed|hunt' "$(if ($lateHunt.Contains("`$env:OUTLOOKAI_HUNT_RELEASE_AFTER_CLOSE = '1'`r`n")) { 'late' })|$(if ((Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 10 $false $false $true $false '' $false $false $true) -match 'give -HuntShowEach -HuntTypeEach') { 'typed' })|$(if ((Get-HuntRefusal $false $false '' '' '' $false 20 60 120 'Legs' 10 $false $false $false $false '' $false $false $true) -match 'a hunting run') { 'hunt' })"
    $traceStart = Get-ClrTraceStartScript 'C:\r' 'S'
    $traceStop = Get-ClrTraceStopScript 'C:\r' 'S'
    $err1 = $null; $err2 = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($traceStart, [ref]$null, [ref]$err1)
    [void][System.Management.Automation.Language.Parser]::ParseInput($traceStop, [ref]$null, [ref]$err2)
    Check '-HuntTraceClr''s guest tasks parse, trace the runtime''s Exception keyword and its rundown, and make tracerpt''s XML' '0|0|True|True|True' "$(@($err1).Count)|$(@($err2).Count)|$($traceStart.Contains("logman start S -p 'Microsoft-Windows-DotNETRuntime' 0x8000 0x4 -o 'C:\r\clr.etl' -ets"))|$($traceStop.Contains("logman start SRundown -p 'Microsoft-Windows-DotNETRuntimeRundown' 0x118 0x5"))|$($traceStop.Contains("tracerpt 'C:\r\clr.etl' 'C:\r\clr-rundown.etl' -o 'C:\r\clr-events.xml' -of XML -y"))"
    $traceAt = $runCode.IndexOf("Invoke-Guest 'clr-trace-start'")
    Check 'the trace starts before the suite and stops after it, and its XML is fetched' $true ($traceAt -gt 0 -and $traceAt -lt $runCode.IndexOf("Invoke-Guest 'live'") -and $runCode.IndexOf("Invoke-Guest 'clr-trace-stop'") -gt $runCode.IndexOf("Invoke-Guest 'live'") -and $runCode.Contains("'first-chance.log', 'clr-events.xml'"))
    Check 'and -HuntShowEach tells it to show a window before each draft' $true ($shownHunt.Contains("`$env:OUTLOOKAI_HUNT_SHOW_EACH = '1'`r`n"))
    Check '-HuntShowEach is a loop''s, in a hunting run' 'hunt|loop|' "$(if ((Get-HuntRefusal $false $false '' '' '' $false 20 60 120 'ProductLoop' 10 $false $false $true) -match 'give -HuntFirstChance') { 'hunt' })|$(if ((Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'Legs' 10 $false $false $true) -match 'other than Legs') { 'loop' })|$(Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'ProductLoop' 10 $false $false $true)"
    Check 'and Legs names none - its tests are not the loop''s' $false ($hunt.Contains('OUTLOOKAI_HUNT_VARIANT'))
    Check 'every hunting run records Office''s build and the hunted module''s file before the debugger starts' $true ($hunt.Contains("'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'") -and $hunt.Contains("-Filter 'mso40uiwin32client.dll'") -and $hunt.IndexOf('HUNT-OFFICE') -gt 0 -and $hunt.IndexOf('HUNT-OFFICE') -lt $hunt.IndexOf('Start-Process'))
    $office = Read-HuntOffice "x`r`nHUNT-OFFICE VersionToReport = 16.0.17932.20996`r`nHUNT-OFFICE UpdatesEnabled = False`r`nHUNT-OFFICE module C:\m.dll, 5 bytes, written 2026-09-12T00:00:00.0000000Z, SHA-256 AB`r`n"
    Check 'the summary reads Office''s build back' '16.0.17932.20996|False|C:\m.dll, 5 bytes, written 2026-09-12T00:00:00.0000000Z, SHA-256 AB' "$($office['VersionToReport'])|$($office['UpdatesEnabled'])|$($office['module'])"
    Check 'and a console with no such line is no record' $true ($null -eq (Read-HuntOffice 'nothing here'))
    $addInSwitch = @($paramAst | Where-Object { $_.Name.VariablePath.UserPath -eq 'HuntWithoutAddIn' })
    Check 'the add-in stays as the checkpoint holds it unless -HuntWithoutAddIn is given' 'switch|True' "$($addInSwitch[0].StaticType.Name.ToLowerInvariant() -replace 'switchparameter','switch')|$($null -eq $addInSwitch[0].DefaultValue)"
    Check '-HuntWithoutAddIn is a hunting run''s, and not one that installs the add-in' 'hunt|both' "$(if ((Get-HuntRefusal $false $false '' '' '' $false 20 60 120 'Legs' 10 $true $false) -match 'give -HuntFirstChance') { 'hunt' })|$(if ((Get-HuntRefusal $true $false '' '' '' $false 20 60 120 'Legs' 10 $true $true) -match 'give one') { 'both' })"
    $off = Get-AddInOffScript 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookAI'
    $err = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($off, [ref]$null, [ref]$err)
    Check 'the add-in''s step parses, and writes LoadBehavior 0 under the installer''s key and reads it back' '0|True' "$(@($err).Count)|$($off.Contains('Set-ItemProperty -LiteralPath $k -Name LoadBehavior -Value 0 -Type DWord') -and $off.Contains("'ADDIN-OFF {0} (was {1})'"))"
    Check 'the key is the one Install-OutlookAIAddIn.ps1 registers' $true ([System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\guest\Install-OutlookAIAddIn.ps1')).Contains('HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI') -and $AddInRegistryKey -eq 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookAI')
    $offAt = $runCode.IndexOf("Invoke-Guest 'addin-off'")
    Check 'the add-in goes off in the frozen start''s prepare, before the Outlook it hands the suite starts' $true ($offAt -gt 0 -and $offAt -lt $runCode.IndexOf("Invoke-Guest 'start-outlook' `"& '`$GuestRoot\Start-OutlookUnelevated.ps1' -Profile '`$TierProfileName'"))
    Check 'each round says whether OUTLOOK.EXE holds the .NET runtime the add-in needs' $true ($hunt.Contains("HUNT-OUTLOOK round {0} pid {1} clr {2}") -and $hunt.IndexOf('HUNT-OUTLOOK') -gt $hunt.IndexOf('$round = $done + 1'))

    Write-Host '== the guest''s clock around every guest call (stream K, 2026-11-18) =='
    $sample = "noise`r`nGUEST-CLOCK-BEFORE 2026-10-03T19:46:53.1234567Z W. Europe Standard Time`r`nSDK-EXECUTE-EXIT 0`r`nGUEST-CLOCK-AFTER 2026-10-03T19:47:53.6234567Z W. Europe Standard Time  `r`n"
    $r = Read-GuestClock $sample
    Check 'both readings are read: the UTC time, as UTC, and the zone, spaces and all' '2026-10-03T19:46:53.1234567Z|W. Europe Standard Time|2026-10-03T19:47:53.6234567Z|W. Europe Standard Time|Utc' @($r.Before.Utc.ToString('o'), $r.Before.Zone, $r.After.Utc.ToString('o'), $r.After.Zone, [string]$r.After.Utc.Kind)
    Check 'an output without the lines has no readings, and one line alone gives no shift' 'True|True|True' @(($null -eq (Read-GuestClock 'SDK-EXECUTE-EXIT 0').Before), ($null -eq (Read-GuestClock 'x GUEST-CLOCK-BEFORE 2026-10-03T19:46:53Z UTC').Before), ($null -eq (Get-GuestClockShift -Readings (Read-GuestClock "GUEST-CLOCK-BEFORE 2026-10-03T19:46:53Z UTC") -HostBefore ([DateTime]::UtcNow) -HostAfter ([DateTime]::UtcNow))))
    $hostBefore = [DateTime]::new(2026, 11, 18, 6, 0, 0, [DateTimeKind]::Utc)
    $steady = Get-GuestClockShift -Readings $r -HostBefore $hostBefore -HostAfter $hostBefore.AddSeconds(61)
    Check 'a call the guest''s clock kept time over: 60.5 s against 61, the guest weeks behind, not moved' '60.5|61.0|-0.5|False' @((Format-ClockSeconds $steady.GuestSeconds), (Format-ClockSeconds $steady.HostSeconds), (Format-ClockSeconds $steady.ApartSeconds -Signed), $steady.Moved)
    $jump = Get-GuestClockShift -Readings (Read-GuestClock "GUEST-CLOCK-BEFORE 2026-10-03T19:00:00Z UTC`nGUEST-CLOCK-AFTER 2026-10-03T21:35:00Z UTC") -HostBefore $hostBefore -HostAfter $hostBefore.AddMinutes(35)
    Check 'two hours forward in a call of 35 minutes is a move of 7200 s' '+7200.0|True' @((Format-ClockSeconds $jump.ApartSeconds -Signed), $jump.Moved)
    $back = Get-GuestClockShift -Readings (Read-GuestClock "GUEST-CLOCK-BEFORE 2026-10-03T19:00:00Z UTC`nGUEST-CLOCK-AFTER 2026-10-03T19:00:30Z UTC") -HostBefore $hostBefore -HostAfter $hostBefore.AddSeconds(120)
    Check 'and a clock that ran 90 s short of this host''s moved too, backwards' '-90.0|True' @((Format-ClockSeconds $back.ApartSeconds -Signed), $back.Moved)
    $edge = Get-GuestClockShift -Readings (Read-GuestClock "GUEST-CLOCK-BEFORE 2026-10-03T19:00:00Z UTC`nGUEST-CLOCK-AFTER 2026-10-03T19:02:00Z UTC") -HostBefore $hostBefore -HostAfter $hostBefore.AddSeconds(60)
    Check 'exactly the tolerance apart is still kept time' '+60.0|False' @((Format-ClockSeconds $edge.ApartSeconds -Signed), $edge.Moved)
    $zoned = Get-GuestClockShift -Readings (Read-GuestClock "GUEST-CLOCK-BEFORE 2026-10-03T19:00:00Z W. Europe Standard Time`nGUEST-CLOCK-AFTER 2026-10-03T19:00:10Z UTC") -HostBefore $hostBefore -HostAfter $hostBefore.AddSeconds(10)
    Check 'another zone after than before is a move, however well the clock kept time' 'True' $zoned.Moved
    Check 'the log line says when, which call, both readings and how far apart' "2026-11-18T06:00:00.000Z  sdk-execute  guest 2026-10-03T19:46:53.123Z 'W. Europe Standard Time' to 2026-10-03T19:47:53.623Z 'W. Europe Standard Time', -3924786.9 s ahead of this host at the start; the guest's clock ran 60.5 s, this host's 61.0 s, -0.5 s apart" (Format-GuestClockLine -Call 'sdk-execute' -HostBefore $hostBefore -Readings $r -Shift $steady)
    Check 'and that a call''s clock was not read' $true ((Format-GuestClockLine -Call 'x' -HostBefore $hostBefore -Readings (Read-GuestClock '') -Shift $null).EndsWith('the guest''s clock was not read: the call''s output holds no GUEST-CLOCK line, or only one'))
    $records = @([pscustomobject]@{ Call = 'swap'; Shift = $steady }, [pscustomobject]@{ Call = 'live'; Shift = $jump }, [pscustomobject]@{ Call = 'crash-events'; Shift = $null })
    Check 'summary.txt names every move' "clock      MOVED in 1 of 3 guest calls: during 'live' the guest's clock ran 9300.0 s and this host's 2100.0 s, +7200.0 s apart; its zone 'UTC' throughout (guest-clock.log)" (Format-GuestClockSummary $records)
    Check 'or says the clock kept time, over how many calls, in which zone' "clock      the guest's clock kept to this host's, within 60 s, over every guest call it was read around: 1 of 2, in 'W. Europe Standard Time' (guest-clock.log)" (Format-GuestClockSummary @($records[0], $records[2]))
    Check 'and a run with no guest call says so' 'clock      no guest call was made' (Format-GuestClockSummary @())
    $invokeGuestAt = $runCode.IndexOf('function Invoke-Guest(')
    $invokeGuest = ''
    if ($invokeGuestAt -ge 0) { $invokeGuest = $runCode.Substring($invokeGuestAt, [Math]::Min(3200, $runCode.Length - $invokeGuestAt)) }
    $marks = @{}
    foreach ($mark in @('$session = Get-GuestSession', '$hostBefore = [DateTime]::UtcNow', '$out = Invoke-Command -Session $session -ScriptBlock', '"GUEST-CLOCK-BEFORE $([DateTime]::UtcNow.ToString(''o'')) $([TimeZoneInfo]::Local.Id)"', 'Set-ExecutionPolicy -Scope Process', '"GUEST-CLOCK-AFTER $([DateTime]::UtcNow.ToString(''o'')) $([TimeZoneInfo]::Local.Id)"', '$hostAfter = [DateTime]::UtcNow', 'Add-GuestClockRecord $Name $out $hostBefore $hostAfter')) { $marks[$mark] = $invokeGuest.IndexOf($mark) }
    Check 'a guest call reads this host''s clock with the session open, the guest''s first and last, and this host''s again after' $true ($marks['$session = Get-GuestSession'] -ge 0 -and $marks['$hostBefore = [DateTime]::UtcNow'] -gt $marks['$session = Get-GuestSession'] -and $marks['$out = Invoke-Command -Session $session -ScriptBlock'] -gt $marks['$hostBefore = [DateTime]::UtcNow'] -and $marks['"GUEST-CLOCK-BEFORE $([DateTime]::UtcNow.ToString(''o'')) $([TimeZoneInfo]::Local.Id)"'] -gt $marks['$out = Invoke-Command -Session $session -ScriptBlock'] -and $marks['Set-ExecutionPolicy -Scope Process'] -gt $marks['"GUEST-CLOCK-BEFORE $([DateTime]::UtcNow.ToString(''o'')) $([TimeZoneInfo]::Local.Id)"'] -and $marks['"GUEST-CLOCK-AFTER $([DateTime]::UtcNow.ToString(''o'')) $([TimeZoneInfo]::Local.Id)"'] -gt $marks['Set-ExecutionPolicy -Scope Process'] -and $marks['$hostAfter = [DateTime]::UtcNow'] -gt $marks['"GUEST-CLOCK-AFTER $([DateTime]::UtcNow.ToString(''o'')) $([TimeZoneInfo]::Local.Id)"'] -and $marks['Add-GuestClockRecord $Name $out $hostBefore $hostAfter'] -gt $marks['$hostAfter = [DateTime]::UtcNow'])
    Check 'and the guest reads its zone afresh each time, as its session''s process keeps the first it read' 2 ([regex]::Matches($invokeGuest, [regex]::Escape('[TimeZoneInfo]::ClearCachedData()'))).Count
    Check 'every guest call goes through it: the only other is the clock read before the suite' $true (([regex]::Matches($runCode, '(?m)^\s*\$o = Invoke-Guest ')).Count -ge 10 -and ([regex]::Matches($runCode, 'Invoke-Command -Session ')).Count -eq 2)
    $asTheRunKeepsThem = New-Object System.Collections.Generic.List[object]
    foreach ($record in $records) { $asTheRunKeepsThem.Add($record) }
    $listJson = New-GuestClockSummaryJson $asTheRunKeepsThem.ToArray()
    $listLine = Format-GuestClockSummary $asTheRunKeepsThem.ToArray()
    Check 'the report reads the records as the run keeps them: a list, given as an array, never @() of it' 'True|3|True' @(($runCode.Contains('New-GuestClockSummaryJson $guestClockRecords.ToArray()') -and $runCode.Contains('Format-GuestClockSummary $guestClockRecords.ToArray()') -and -not $runCode.Contains('@($guestClockRecords)')), $listJson.calls, $listLine.StartsWith('clock      MOVED in 1 of 3'))
    Check 'summary.json carries the record, its moves an array at every length' $true ($runCode.Contains("`$summary['guestClock'] = New-GuestClockSummaryJson") -and ((New-GuestClockSummaryJson $records | ConvertTo-Json -Depth 5) -match '"moved":\s*\[') -and ((New-GuestClockSummaryJson @() | ConvertTo-Json -Depth 5) -match '"moved":\s*\[\s*\]'))

    Write-Host '== the constants, against the files they copy =='
    $media = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\MEDIA.md'))
    Check 'MEDIA.md names the SDK installer' $true $media.Contains($SdkInstallerName)
    Check 'MEDIA.md records its SHA-512' $true $media.ToUpperInvariant().Contains($SdkSha512.ToUpperInvariant())
    $filters = [System.IO.File]::ReadAllText((Join-Path $repo 'McpServer\OutlookAI.McpServer.Tests\T2\LiveRunFilters.cs'))
    Check 'LiveRunFilters.Guest leaves out exactly DelegateStore and CachedExchange' $true $filters.Contains('ExchangeOnlyCapabilities { get; } = new[] { DelegateStore, CachedExchange }')
    Check 'LiveRunFilters.GuestUnindexed adds SearchIndex' $true $filters.Contains('GuestUnindexed { get; } = Guest + "&Requires!=" + SearchIndex')
    foreach ($name in @('Reset-HubPopulation.ps1')) {
        $text = [System.IO.File]::ReadAllText((Join-Path $repo "Testbed\guest\$name"))
        Check "$name reads the settings where this stages them" $true $text.Contains("`$SettingsPath = '$GuestSettingsPath'")
    }
    $sdk = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\guest\Install-DotnetSdk.ps1'))
    Check 'Install-DotnetSdk.ps1 builds the source this stages' $true $sdk.Contains("'$GuestRoot\src'")
    Check 'and takes the installer and the version prefix this names it' $true ($sdk.Contains('[string]   $InstallerPath ') -and $sdk.Contains('[string]   $ExpectedSdkVersionPrefix '))
    Check 'both calls name them' 2 ([regex]::Matches($own, [regex]::Escape("& .\Install-DotnetSdk.ps1 `$sdkNamed -")).Count)
    $testbed = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\testbed.json')) | ConvertFrom-Json
    foreach ($name in @($Guests.Keys | Where-Object { -not $Guests[$_].Exchange })) {
        Check "$name rests on the frozen checkpoint testbed.json records" ([string]$testbed.frozenClocks.$name.checkpoint) $Guests[$name].Checkpoint
    }
    Check 'the Exchange VM is the one testbed.json exchangeVm names' 'OutlookAI-Exchange|OAI-EXCHANGE' @([string]$testbed.exchangeVm.vmName, [string]$testbed.exchangeVm.computerName)
    Check 'and rests on the newest checkpoint testbed.json records for it' ([string]@($testbed.exchangeVm.checkpoints)[-1]) $Guests['OutlookAI-Exchange'].Checkpoint
    Check 'LiveRunFilters.ExchangeTestAccount leaves out exactly what the Exchange VM lacks, and the one-off removal' $true $filters.Contains('new[] { LocalSignatures, DelegateStore, GeneratedPopulation, IdentityAccount, ProbePopulation, PstStore, ConfirmedLeftover }')
    Check 'and the filter here is that list, in that order' ('Category=Live' + ((@('LocalSignatures', 'DelegateStore', 'GeneratedPopulation', 'IdentityAccount', 'ProbePopulation', 'PstStore', 'ConfirmedLeftover') | ForEach-Object { '&Requires!=' + $_ }) -join '')) $FilterExchange
    $marker = [System.IO.File]::ReadAllText((Join-Path $repo 'McpServer\OutlookAI.McpServer.Tests\T2\LiveRunMarker.cs'))
    Check 'the run reads the marker from the variable this sets' $true ($marker.Contains('"OUTLOOKAI_LIVE_RUN_MARKER"') -and $marker.Contains('"OUTLOOKAI_LIVE_PRIOR_RUN_MARKERS"') -and $own.Contains("`$env:OUTLOOKAI_LIVE_RUN_MARKER = '"))
    foreach ($name in $Guests.Keys) {
        $section = $testbed.liveTestSettings.$name
        $indexed = @(@($section.indexedStoreDisplayNames) | Where-Object { $_ }).Count -gt 0
        Check "$name is indexed exactly when testbed.json indexes its stores" $Guests[$name].Indexed $indexed
    }

    Write-Host ''
    if ($script:stFailures.Count -gt 0) {
        Write-Host "SelfTest: $($script:stChecks - $script:stFailures.Count) passed, $($script:stFailures.Count) failed."
        $script:stFailures | ForEach-Object { Write-Host "  $_" }
        return 1
    }
    Write-Host "SelfTest: $($script:stChecks) passed, 0 failed."
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE RUN.
# =============================================================================================
$started = Get-Date
# Say appends to it once the run has a directory; until then a line, a refusal too, is only printed.
$script:LogPath = $null

if ($VMName -and $Kind) { Stop-Refused '-VMName and -Kind together are ambiguous: name a guest, or a kind of guest.' }
if (-not $VMName -and -not $Kind) { Stop-Refused 'name a guest with -VMName, or a kind of guest with -Kind Indexed or -Kind Unindexed: no default picks one (Testbed/README.md section 4a).' }
if ($Kind) {
    if ($Checkpoint -or $RestingCheckpoint -or $GreenCheckpoint -or $RestOnGreen) { Stop-Refused '-Checkpoint, -RestingCheckpoint, -GreenCheckpoint and -RestOnGreen name one machine''s checkpoints; with -Kind a run starts and rests on the frozen checkpoint of whichever guest it takes. Give -VMName for them.' }
    # The kind's source stands in until a guest is taken: every guest of a kind has the same facts.
    $facts = Get-GuestFacts (Get-KindMembers $Kind)[0]
}
else {
    $facts = Get-GuestFacts $VMName
    if ($null -eq $facts) { Stop-Refused "'$VMName' is not a test guest - only the Outlook guests and their clones: $((@(Get-KindMembers 'Indexed') + @(Get-KindMembers 'Unindexed')) -join ', '), and OutlookAI-Exchange. The build VM has Invoke-TestsOnBuildVm.ps1, and nothing here runs on this workstation." }
}
$filterDecision = Get-LiveFilter $facts.Indexed $Filter $FilterSuffix $facts.Exchange $SignatureSettings ([bool]$SignatureDialog)
if ($filterDecision.Refusal) { Stop-Refused $filterDecision.Refusal }
$runFilter = $filterDecision.Filter
$narrowed = $filterDecision.Narrowed
if ($RestOnGreen -and -not $GreenCheckpoint) { Stop-Refused '-RestOnGreen needs -GreenCheckpoint.' }
# The Exchange VM carries its state forward unless a checkpoint is named (see the header) - read before
# Use-Guest gives either checkpoint its default.
$exchangeCarry = [bool]$facts.Exchange -and -not $Checkpoint
$exchangeRestCarry = [bool]$facts.Exchange -and -not $RestingCheckpoint

# Why a run cannot have this guest from Hyper-V - '' when it can: the VM, the start and resting
# checkpoints, and no checkpoint already under the green name.
function Get-GuestRunRefusal([object] $Guest) {
    if (-not (Get-VM -Name $Guest.Name -ErrorAction SilentlyContinue)) { return "Hyper-V has no VM named '$($Guest.Name)' here." }
    $start = $Checkpoint
    if (-not $start) { $start = $Guest.Checkpoint }
    $rest = $RestingCheckpoint
    if (-not $rest) { $rest = $Guest.Checkpoint }
    foreach ($cp in @($start, $rest)) {
        if (-not (Get-VMSnapshot -VMName $Guest.Name -Name $cp -ErrorAction SilentlyContinue)) { return "$($Guest.Name) has no checkpoint named '$cp'." }
    }
    if ($GreenCheckpoint -and (Get-VMSnapshot -VMName $Guest.Name -Name $GreenCheckpoint -ErrorAction SilentlyContinue)) { return "$($Guest.Name) already has a checkpoint named '$GreenCheckpoint'." }
    return ''
}
# The guest this run uses from here on, its checkpoints, and whether its start is FROZEN: a
# checkpoint whose time sync is off (Q130 (a)), the setting travelling with the checkpoint.
function Use-Guest([object] $Guest) {
    $script:facts = $Guest
    if (-not $script:Checkpoint) { $script:Checkpoint = $Guest.Checkpoint }
    if (-not $script:RestingCheckpoint) { $script:RestingCheckpoint = $Guest.Checkpoint }
    $ts = @(Get-VMSnapshot -VMName $Guest.Name -Name $script:Checkpoint -ErrorAction SilentlyContinue | Get-VMIntegrationService -ErrorAction SilentlyContinue | Where-Object { ([string]$_.Id).ToUpperInvariant().EndsWith($TimeSyncComponentId) -or $_.Name -eq 'Time Synchronization' })
    $script:startIsFrozen = ($ts.Count -gt 0 -and -not $ts[0].Enabled)
}
$startIsFrozen = $false
if (-not $Kind) {
    $refusal = Get-GuestRunRefusal $facts
    if ($refusal) { Stop-Refused $refusal }
    Use-Guest $facts
    if ($exchangeCarry) { $startIsFrozen = $false }
    if ($startIsFrozen -and $RestOnGreen) { Stop-Refused "-RestOnGreen from the frozen checkpoint '$Checkpoint': a green checkpoint of this run would stand a run later than the frozen instant testbed.json records, and every later run would start there. Make a new frozen checkpoint with Set-GuestClockFrozen.ps1 and record it instead." }
}
# The add-in goes onto a frozen-clock Outlook guest only: a -Kind run's guest, taken later, is checked again once it is.
if ($HuntWithoutAddIn -and ($facts.Exchange -or (-not $Kind -and -not $startIsFrozen))) { Stop-Refused '-HuntWithoutAddIn turns the add-in off on a frozen-clock Outlook guest only: an Outlook guest or a clone, from its frozen checkpoint.' }
if ($WithAddIn -and ($facts.Exchange -or (-not $Kind -and -not $startIsFrozen))) { Stop-Refused '-WithAddIn installs the add-in on a frozen-clock Outlook guest only: an Outlook guest or a clone, from its frozen checkpoint.' }
# With -Kind the guest is taken later, and every guest of a kind starts from its frozen checkpoint; it is checked
# again once taken.
$stepRefusal = Get-SignatureSettingsRefusal $SignatureSettings ([bool]$facts.Exchange) ($startIsFrozen -or [bool]$Kind) $GreenCheckpoint $Filter ([bool]$SignatureDialog)
if ($stepRefusal) { Stop-Refused $stepRefusal }
$huntRefusal = Get-HuntRefusal ([bool]$HuntFirstChance) ([bool]$facts.Exchange) $Filter $FilterSuffix $SignatureSettings ([bool]$SignatureDialog) $HuntRounds $HuntMinutes $RunTimeoutMinutes $HuntVariant $HuntLegMinutes ([bool]$HuntWithoutAddIn) ([bool]$WithAddIn) ([bool]$HuntShowEach) ([bool]$HuntTypeEach) $HuntAblate ([bool]$HuntTraceClr) ([bool]$HuntPromoteAfter) ([bool]$HuntReleaseAfterClose)
if ($huntRefusal) { Stop-Refused $huntRefusal }
# A hunting run runs its own tests (Get-HuntSuiteScript): narrowed, and every leg's filter the guest's derived one with
# the leg's suffix.
$huntFilters = $null
if ($HuntFirstChance) {
    $narrowed = $true
    $huntFilters = [ordered]@{}
    $legs = Get-HuntLegs $HuntVariant
    foreach ($leg in $legs.Keys) { $huntFilters[$leg] = $runFilter + $legs[$leg] }
}
if (-not $RepoPath) { $RepoPath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$repo = [System.IO.Path]::GetFullPath($RepoPath)
if (-not $CredentialRepoRoot) {
    $common = (@(Invoke-NativeCommand { & git -C $repo rev-parse --path-format=absolute --git-common-dir 2>&1 }) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { Stop-Refused "$repo is not a git checkout: $common" }
    $CredentialRepoRoot = Split-Path -Parent ([System.IO.Path]::GetFullPath($common))
}
$credentialFile = Join-Path $CredentialRepoRoot 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\vm-credentials.json'
if (-not (Test-Path -LiteralPath $credentialFile)) { Stop-Refused "no guest credential at $credentialFile. It lives in the main checkout only (gitignored); pass -CredentialRepoRoot." }
if (-not $SdkInstallerPath) { $SdkInstallerPath = Join-Path $CredentialRepoRoot ".work\media\$SdkInstallerName" }
if ($WithAddIn -and -not $VstoRuntimePath) {
    $VstoRuntimePath = @((Join-Path $CredentialRepoRoot '.work\media\vstor_redist.exe'), (Join-Path $CredentialRepoRoot 'Redist\vstor_redist.exe')) |
        Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $VstoRuntimePath) { Stop-Refused "-WithAddIn needs the staged vstor_redist.exe (Testbed/MEDIA.md): not at .work\media\ or Redist\ of $CredentialRepoRoot - name it with -VstoRuntimePath" }
}
if ($WithAddIn -and -not (Test-Path -LiteralPath $VstoRuntimePath)) { Stop-Refused "no VSTO runtime at $VstoRuntimePath (Testbed/MEDIA.md); pass -VstoRuntimePath." }
if (-not (Test-Path -LiteralPath $SdkInstallerPath)) { Stop-Refused "no SDK installer at $SdkInstallerPath (Testbed/MEDIA.md); pass -SdkInstallerPath." }

$sha = (@(Invoke-NativeCommand { & git -C $repo rev-parse --verify "$Ref^{commit}" 2>&1 }) | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { Stop-Refused "git could not resolve '$Ref' to a commit in $repo - $sha" }
$subject = (@(Invoke-NativeCommand { & git -C $repo log -1 --format=%s $sha 2>&1 }) | Out-String).Trim()
$dirtyCount = 0
if ($Ref -eq 'HEAD') { $dirtyCount = @(Invoke-NativeCommand { & git -C $repo status --porcelain 2>&1 } | Where-Object { $_ }).Count }
if (-not $ResultsRoot) { $ResultsRoot = Join-Path $repo '.work\guest-live-runs' }
# The run's own directory, claimed before anything is named after the run - its build tree, its folder
# on the guest, its lease and its queue ticket: a run of the same commit on the same kind of guest
# started in the same second gets -2, -3 ... (New-RunDirectory), and the name it gives is the run's.
$made = New-RunDirectory -Parent $ResultsRoot -RunId (New-RunId $started $sha (Get-RunLabel $facts.Name))
$runId = $made.RunId
$runDir = $made.Path
New-Item -ItemType Directory -Force -Path (Join-Path $runDir 'guest') | Out-Null
$script:LogPath = Join-Path $runDir 'run.log'
$revision = [ordered]@{ ref = $Ref; sha = $sha; subject = $subject; repo = $repo; uncommittedLeftOut = $dirtyCount }

# Not $kind: PowerShell's names ignore case, and that is the -Kind parameter.
$kindText = 'unindexed'
if ($facts.Indexed) { $kindText = 'indexed' }
if ($facts.Exchange) { $kindText = 'Exchange, test account the hub' }
# The Exchange VM's run marker and the earlier runs' it must sweep, from the host's ledger.
$runMarker = $null
$priorMarkers = @()
$ledgerPath = $null
$ledgerOutcome = $null
if ($facts.Exchange) {
    . (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
    $ledgerPath = Join-Path (Split-Path -Parent (Get-TestbedLeaseDirectory)) $ExchangeLedgerRelative
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $ledgerPath) | Out-Null
    $runMarker = New-RunMarker $started
    $ledgerLines = @()
    if (Test-Path -LiteralPath $ledgerPath) { $ledgerLines = @(Get-Content -LiteralPath $ledgerPath) }
    $priorMarkers = @(Get-PriorRunMarkers $ledgerLines $runMarker)
}
Say "== Invoke-LiveTierOnGuest: run $runId =="
if ($Kind) { Say "  guest     whichever $kindText guest is free once the payloads are built: $((Get-KindMembers $Kind) -join ', ') ($($facts.ComputerName)), from and resting on its frozen checkpoint" }
elseif ($exchangeCarry) { Say "  guest     $($facts.Name) ($($facts.ComputerName), $kindText), carried forward from its last run - no checkpoint restored" }
else { Say "  guest     $($facts.Name) ($($facts.ComputerName), $kindText), from $Checkpoint, resting on $RestingCheckpoint" }
if ($startIsFrozen) { Say "  clock     FROZEN - '$Checkpoint' has time sync off: no restart and no hub rebuild; Set-GuestClockFrozen.ps1 -Verify before the suite" }
Say "  revision  $($sha.Substring(0, 12)) - $subject"
if ($dirtyCount -gt 0) { Say "  NOTE      $dirtyCount uncommitted change(s) in $repo are NOT in this run - it tests the commit. Commit first to test them." }
Say "  filter    $runFilter"
if ($facts.Exchange) { Say "  marker    $runMarker; earlier runs to sweep by marker: $(if ($priorMarkers.Count -gt 0) { $priorMarkers -join ', ' } else { 'none' }) - ledger $ledgerPath" }
Say "  results   $runDir"

$timings = [ordered]@{}
$checkpointActions = [ordered]@{ start = $Checkpoint; resting = $RestingCheckpoint; green = $null; restedOn = $null; settingsStaged = $false; saved = $false }
$crashDumpState = [ordered]@{ armed = $false; proven = $null; guestFolder = $GuestDumpDir; fetched = $false; files = @(); werReports = -1; werFiles = @() }
# The -SignatureSettings step, when asked for: the scenario, and whether the guest script said SET.
$signatureSettingsState = $null
if ($SignatureSettings) { $signatureSettingsState = [ordered]@{ scenario = $SignatureSettings; valueKind = $SignatureValueKind; set = $false; log = 'guest\signature-settings.log' } }
# Set when this run starts Outlook itself: what the start met, Office's activation wizard included.
$outlookStart = $null
$indexRoots = $null
$script:dumpsFound = @()
$script:werFound = @()
$huntRoundsRun = -1
$huntOffice = $null
$huntAddIn = 'as the checkpoint holds it'
$huntClrTrace = $null
$suiteStarted = $false
$stage = ''
$stageWhy = ''
$trx = $null
$consoleFacts = $null
$crashes = -1
$suiteExit = $null
$touched = $false
$credential = $null
$tree = Join-Path $runDir 'tree'
$guestRunDir = Join-Path $GuestRunsRoot $runId
. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')

function Use-Lease([int] $Minutes) {
    $decision = Get-LeaseDecision (Get-TestbedLease -VMName $facts.Name) "$LeaseReasonPrefix $runId" ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
    if ($decision -eq 'theirs') { throw "the lease on $($facts.Name) is held by someone else - not touching the guest" }
    & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $facts.Name -Minutes ([Math]::Min(480, [Math]::Max(30, $Minutes))) -Reason "$LeaseReasonPrefix $runId" | Out-Null
}
# Takes the lease on a guest in one step with reading it (Enter-TestbedLeaseClaim), so two runs that
# find the same guest free in the same moment cannot both take it. Returns $null when the lease is
# this run's, or the live lease somebody else holds - which is never taken over.
function Request-GuestLease([string] $Name, [int] $Minutes) {
    $claim = Enter-TestbedLeaseClaim -VMName $Name
    try {
        $lease = Get-TestbedLease -VMName $Name
        if ((Get-LeaseDecision $lease "$LeaseReasonPrefix $runId" ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())) -eq 'theirs') { return $lease }
        & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $Name -Minutes ([Math]::Min(480, [Math]::Max(30, $Minutes))) -Reason "$LeaseReasonPrefix $runId" | Out-Null
        return $null
    }
    finally { $claim.Dispose() }
}
# The guest's Hyper-V id - asked once a run - its state and its heartbeat are TestbedPool.ps1's narrow
# queries by the guest's name (stream W, 2026-11-18): Get-TestbedVmId, Get-TestbedVmState and
# Wait-TestbedHeartbeat, every 2 s, which the trust-removal probe asks too.
# ONE PowerShell Direct session for the run's own guest calls (stream W, 2026-11-18): every guest call
# opened a session of its own, by the VM's name - nine in a run of one test class that morning.
# Opened by the guest's id at the first call that needs it, and closed wherever the guest is restored
# or restarted, which ends it. The scripts this runner starts - Copy-ToGuest.ps1 and the others - still
# open their own.
$guestSession = $null
function Close-GuestSession {
    if ($null -ne $script:guestSession) { try { Remove-PSSession -Session $script:guestSession -ErrorAction SilentlyContinue } catch { } }
    $script:guestSession = $null
}
function Get-GuestSession {
    if ($null -ne $script:guestSession -and [string]$script:guestSession.State -eq 'Opened') { return $script:guestSession }
    Close-GuestSession
    $script:guestSession = New-PSSession -VMId (Get-TestbedVmId $facts.Name) -Credential $credential -ErrorAction Stop
    return $script:guestSession
}
# Every guest call's clock record - Call and Shift (Get-GuestClockShift) - in call order, for the
# summary; each also a line in guest-clock.log, and a warning in run.log when the clock moved. A record
# that cannot be written is said and passed over: it never fails the call.
$guestClockRecords = New-Object System.Collections.Generic.List[object]
function Add-GuestClockRecord([string] $Call, [string] $Output, [DateTime] $HostBefore, [DateTime] $HostAfter) {
    try {
        $readings = Read-GuestClock $Output
        $shift = Get-GuestClockShift -Readings $readings -HostBefore $HostBefore -HostAfter $HostAfter
        $guestClockRecords.Add([pscustomobject]@{ Call = $Call; Shift = $shift })
        Add-Content -LiteralPath (Join-Path $runDir 'guest-clock.log') -Value (Format-GuestClockLine -Call $Call -HostBefore $HostBefore -Readings $readings -Shift $shift) -Encoding UTF8
        if ($null -ne $shift -and $shift.Moved) { Say "GUEST CLOCK: $(Format-GuestClockMove $Call $shift) - guest-clock.log" }
    }
    catch { Say "  (the guest's clock around '$Call' could not be recorded: $($_.Exception.Message))" }
}
# A checkpoint restored over the guest by TestbedPool.ps1's Restore-TestbedCheckpoint, which saves a
# guest that runs first (2026-11-18), then the guest started and its heartbeat waited for. The restore
# ends the run's guest session, so it is closed first.
function Restore-Guest([string] $Name) {
    Close-GuestSession
    [void](Restore-TestbedCheckpoint -Name $facts.Name -Checkpoint $Name -Start)
}
# One piece of work on the guest; its whole output goes to guest\<name>.log. Session 1 through
# Register-InteractiveTask.ps1 at RunLevel Limited (where Outlook runs, NOT elevated), or session 0
# over PowerShell Direct for work that needs no desktop - both over the run's one session, each call
# starting as a session of its own did: with no exit code left by the call before it, which
# Get-TaskExit would read as this call's, and in the folder the session started in. Each also reads the
# guest's clock and zone first and last, and this host's around it (THE GUEST'S CLOCK, above).
function Invoke-Guest([string] $Name, [string] $Script, [int] $TimeoutSeconds, [switch] $Session0) {
    Use-Lease ([int][Math]::Ceiling($TimeoutSeconds / 60) + 30)
    # The session first, so that this host's two readings of its clock hold the call and nothing more.
    $session = Get-GuestSession
    $hostBefore = [DateTime]::UtcNow
    $out = Invoke-Command -Session $session -ScriptBlock {
        param($text, $timeout, $inSession0)
        $ErrorActionPreference = 'Continue'
        $global:LASTEXITCODE = $null
        if ($null -eq $global:OutlookAIGuestStart) { $global:OutlookAIGuestStart = (Get-Location).Path }
        Set-Location -LiteralPath $global:OutlookAIGuestStart
        # The guest's clock and zone, first and last (Read-GuestClock) - the zone read afresh each time,
        # as a process that has read it once keeps it.
        [TimeZoneInfo]::ClearCachedData()
        "GUEST-CLOCK-BEFORE $([DateTime]::UtcNow.ToString('o')) $([TimeZoneInfo]::Local.Id)"
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        if ($inSession0) {
            try { & ([scriptblock]::Create($text)) *>&1 | Out-String -Width 400 } catch { "GUEST-CALL-THREW: $($_.Exception.Message)" }
        }
        else {
            try { & 'C:\OutlookAI-Q5\Register-InteractiveTask.ps1' -Script $text -TimeoutSeconds $timeout -RunLevel Limited *>&1 | Out-String -Width 400 } catch { "GUEST-CALL-THREW: $($_.Exception.Message)" }
            "RUN-REGISTER-EXIT: $LASTEXITCODE"
        }
        [TimeZoneInfo]::ClearCachedData()
        "GUEST-CLOCK-AFTER $([DateTime]::UtcNow.ToString('o')) $([TimeZoneInfo]::Local.Id)"
    } -ArgumentList $Script, $TimeoutSeconds, ([bool]$Session0) 2>&1 | Out-String -Width 400
    $hostAfter = [DateTime]::UtcNow
    [System.IO.File]::WriteAllText((Join-Path $runDir "guest\$Name.log"), $out)
    Add-GuestClockRecord $Name $out $hostBefore $hostAfter
    return $out
}
# The dumps WER wrote during the run, fetched into dumps\<guest>\ BEFORE the guest is restored - the
# restore wipes them. Waits, bounded, for WerFault.exe to finish writing first. Once per run; never throws.
function Save-CrashDumps {
    if ($crashDumpState['fetched']) { return }
    $crashDumpState['fetched'] = $true
    try {
        # WER's own reports of an OUTLOOK.EXE crash written during the run - the machine's and each user's
        # queue and archive - copied beside the dumps with their report folder's name, each file listed as
        # WERFILE (ConvertFrom-WerListing). On the test guests WER keeps only Report.wer in such a folder
        # (measured 2026-11-18: telemetry is Required, so no dump goes with a report); a dump found there is
        # copied too.
        $werSince = Get-CrashEventsSince -GuestRunStartUtc $guestRunStartUtc -HostRunStartUtc $runStartUtc
        $werStage = @"
`$cut = ([datetime]'$werSince').ToUniversalTime()
`$stage = '$GuestDumpDir\wer'
New-Item -ItemType Directory -Force -Path `$stage | Out-Null
`$roots = @('C:\ProgramData\Microsoft\Windows\WER\ReportQueue', 'C:\ProgramData\Microsoft\Windows\WER\ReportArchive')
foreach (`$u in @(Get-ChildItem -LiteralPath 'C:\Users' -Directory -ErrorAction SilentlyContinue)) { `$roots += (Join-Path `$u.FullName 'AppData\Local\Microsoft\Windows\WER\ReportQueue'); `$roots += (Join-Path `$u.FullName 'AppData\Local\Microsoft\Windows\WER\ReportArchive') }
foreach (`$r in `$roots) {
    foreach (`$d in @(Get-ChildItem -LiteralPath `$r -Directory -ErrorAction SilentlyContinue | Where-Object { `$_.LastWriteTimeUtc -ge `$cut })) {
        `$report = Join-Path `$d.FullName 'Report.wer'
        if (-not (`$d.Name -match 'OUTLOOK\.EXE' -or ((Test-Path -LiteralPath `$report) -and (Select-String -LiteralPath `$report -Pattern 'OUTLOOK\.EXE' -Encoding Unicode -Quiet)))) { continue }
        foreach (`$f in @(Get-ChildItem -LiteralPath `$d.FullName -File -ErrorAction SilentlyContinue | Where-Object { @('.wer', '.mdmp', '.hdmp', '.dmp') -contains `$_.Extension.ToLowerInvariant() })) {
            `$n = `$d.Name + '.' + `$f.Name
            Copy-Item -LiteralPath `$f.FullName -Destination (Join-Path `$stage `$n) -Force
            'WERFILE ' + `$f.Length + ' ' + `$n
        }
    }
}
"@
        $list = "`$deadline = (Get-Date).AddSeconds(300); while ((Get-Date) -lt `$deadline -and @(Get-Process -Name WerFault -ErrorAction SilentlyContinue).Count -gt 0) { Start-Sleep -Seconds 3 }; `"WERFAULT-LEFT `$(@(Get-Process -Name WerFault -ErrorAction SilentlyContinue).Count)`"; Get-ChildItem -LiteralPath '$GuestDumpDir' -Filter '*.dmp' -File -ErrorAction SilentlyContinue | ForEach-Object { 'DUMP ' + `$_.Length + ' ' + `$_.Name }"
        $o = Invoke-Guest 'crash-dumps-list' ($list + "`r`n" + $werStage) 600 -Session0
        $script:dumpsFound = @(ConvertFrom-DumpListing $o)
        $script:werFound = @(ConvertFrom-WerListing $o)
        if ($script:werFound.Count -gt 0) {
            $werDest = Join-Path $runDir 'wer'
            # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
            & (Join-Path $PSScriptRoot 'Copy-FromGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -GuestPath "$GuestDumpDir\wer" -Include '*.wer', '*.mdmp', '*.hdmp', '*.dmp' -Destination $werDest -Force *>> (Join-Path $runDir 'fetch.log')
            $werFiles = @()
            foreach ($w in $script:werFound) {
                $local = Join-Path (Join-Path $werDest $facts.Name) $w.Name
                $werFiles += [ordered]@{ name = $w.Name; bytes = $w.Bytes; fetched = ((Test-Path -LiteralPath $local) -and ((Get-Item -LiteralPath $local).Length -eq $w.Bytes)); path = $local }
            }
            $crashDumpState['werFiles'] = $werFiles
            Say "WER's own reports of OUTLOOK.EXE: $($werFiles.Count) file(s) on the guest, $(@($werFiles | Where-Object { $_['fetched'] }).Count) fetched to wer\$($facts.Name)\"
        }
        if ($script:dumpsFound.Count -eq 0) { return }
        $dest = Join-Path $runDir 'dumps'
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Copy-FromGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -GuestPath $GuestDumpDir -Include '*.dmp' -Destination $dest -Force *>> (Join-Path $runDir 'fetch.log')
        $files = @()
        foreach ($d in $script:dumpsFound) {
            $local = Join-Path (Join-Path $dest $facts.Name) $d.Name
            $ok = (Test-Path -LiteralPath $local) -and ((Get-Item -LiteralPath $local).Length -eq $d.Bytes)
            $entry = [ordered]@{ name = $d.Name; program = $d.Program; bytes = $d.Bytes; fetched = $ok; path = $local; report = $null; headline = @() }
            if ($ok) {
                # The first look, with nothing but Windows (host\Read-CrashDump.ps1): beside the dump, and its
                # headline into the summary.
                try {
                    $reportPath = "$local.txt"
                    $text = (& (Join-Path $PSScriptRoot 'Read-CrashDump.ps1') -Path $local -OutFile $reportPath | Out-String)
                    $entry['report'] = $reportPath
                    $entry['headline'] = @(Get-DumpHeadline $text)
                }
                catch { Say "reading $($d.Name) failed: $($_.Exception.Message)" }
            }
            $files += $entry
        }
        $crashDumpState['files'] = $files
        Say "crash dumps: $($script:dumpsFound.Count) on the guest, $(@($files | Where-Object { $_['fetched'] }).Count) fetched to dumps\$($facts.Name)\"
    }
    catch { Say "fetching the crash dumps failed: $($_.Exception.Message)" }
}

try {
    # ---- GATE (Q195 (e), TestbedLoad.ps1): the pause switch and the host's load, before the build puts
    # this run's first load on the host. Waits holding nothing; a run the gate never lets through is INFRA.
    $t0 = Get-Date
    Say "gate: the pause switch, then the host's load over the last $TestbedLoadWindowSeconds s (Testbed/host/TestbedLoad.ps1)"
    $gate = Wait-TestbedGate -Run $runId -What "Invoke-LiveTierOnGuest run $runId, before its build" -TimeoutMinutes $LeaseWaitMinutes -Say { param($line) Say $line }
    if (-not $gate.Go) { $stage = 'INFRA'; $stageWhy = "the gate did not open in $LeaseWaitMinutes minutes - $($gate.Why)"; throw $stageWhy }
    $timings['gate'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "gate: $(Format-TestbedGateOpen -Look $gate.Look -PendingPercent $gate.Pending.Percent -PendingCount $gate.Pending.Count) ($($timings['gate']))"

    # ---- BUILD: the commit's tree, its payloads and the guest's settings. The guest is untouched.
    $t0 = Get-Date
    try {
        # No MSBuild node outlives this run's builds holding files of the tree, so it can be deleted after.
        $env:MSBUILDDISABLENODEREUSE = '1'
        Invoke-NativeCommand { & git -C $repo worktree add --detach $tree $sha 2>&1 } | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "git worktree add failed ($LASTEXITCODE)" }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Publish-GuestPayload.ps1') -RepoRoot $tree -OutDir (Join-Path $runDir 'payload') *> (Join-Path $runDir 'publish-guest.log')
        if ($LASTEXITCODE -ne 0) { throw "Publish-GuestPayload.ps1 failed ($LASTEXITCODE) - publish-guest.log" }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Publish-LiveTierPayload.ps1') -RepoRoot $tree -OutDir (Join-Path $runDir 'livetier') -Ref HEAD -SdkInstallerPath $SdkInstallerPath -ExpectedSha512 $SdkSha512 *> (Join-Path $runDir 'publish-livetier.log')
        if ($LASTEXITCODE -ne 0) { throw "Publish-LiveTierPayload.ps1 failed ($LASTEXITCODE) - publish-livetier.log" }
        if ($WithAddIn) {
            # The tested commit's add-in, built as a release is, with the Q81 guards: nothing of it reaches this host's Outlook.
            # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
            & (Join-Path $PSScriptRoot 'Publish-AddInPayload.ps1') -RepoRoot $tree -Ref HEAD -OutDir (Join-Path $runDir 'addin') -VstoRuntimePath $VstoRuntimePath *> (Join-Path $runDir 'publish-addin.log')
            if ($LASTEXITCODE -ne 0) { throw "Publish-AddInPayload.ps1 failed ($LASTEXITCODE) - publish-addin.log" }
        }
        # A clone's settings are its source's (Get-GuestFacts): its stores, corpus and sink are copies.
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'New-LiveTestSettings.ps1') -VMName $facts.Source -RepoRoot $tree -CredentialRoot $CredentialRepoRoot -OutPath (Join-Path $runDir 'live-test-settings.json') *> (Join-Path $runDir 'settings.log')
        if ($LASTEXITCODE -ne 0) { throw "New-LiveTestSettings.ps1 refused this guest's settings ($LASTEXITCODE) - settings.log" }
    }
    catch { $stage = 'BUILD'; $stageWhy = $_.Exception.Message; throw }
    $timings['build'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "built: payloads and settings ($($timings['build']))"

    # ---- LEASE: wait, never take over. With -Kind, whichever guest of the kind is free: the source
    # first, a clone only while the host has its startup memory and the reserve (TestbedPool.ps1).
    # A run that waits asks Hyper-V nothing (2026-11-18, TestbedPool.ps1's banner): it holds a ticket
    # in the testbed's queue and tries only the guests no earlier ticket wants; whether a guest is
    # free is its lease file; Hyper-V is asked about a guest only once its lease is free - and the facts
    # that do not change in a run, its checkpoints and its startup memory, once a run. Every turn looks
    # at the gate first, before its place in the queue, and a lease is taken only under the gate's
    # claim, the VMs other runs took lately counted; a run told to wait lets the claim go
    # (TestbedLoad.ps1). Between turns it sleeps until a lease, a lock's owner file, a ticket or the
    # gate's switch changes, or 30 s pass.
    $t0 = Get-Date
    $deadline = (Get-Date).AddMinutes($LeaseWaitMinutes)
    $candidates = @($facts.Name)
    if ($Kind) { $candidates = @(Get-KindMembers $Kind) }
    $lastSaid = ''
    $lastSaidAt = [DateTime]::MinValue
    $waitingSince = $null
    $takenGuest = $null
    $refusals = @{}
    $startupBytes = @{}
    $look = $null
    $leaseStart = $null
    # The gate's readings from before the build are dropped: the first look here starts after it, so
    # the build - this run's own load - is never what holds the run back.
    $TestbedLoadReadings.Clear()
    $ticket = Join-TestbedQueue -Run $runId -What "Invoke-LiveTierOnGuest run $runId, $kindText" -Candidates $candidates
    try {
        while ($null -eq $takenGuest) {
            $why = @()
            $look = Test-TestbedGate
            $gated = -not $look.Go
            if ($gated) { $why += $look.Why }
            else {
                $ahead = Get-TestbedQueueTurn -Ticket $ticket -Run $runId -Candidates $candidates
                for ($i = 0; $i -lt $candidates.Count; $i++) {
                    $guest = Get-GuestFacts $candidates[$i]
                    if ($ahead.ContainsKey($guest.Name)) { $why += "$($guest.Name) is wanted first by run $($ahead[$guest.Name]), queued before this one"; continue }
                    $held = Get-TestbedLease -VMName $guest.Name
                    if ((Get-LeaseDecision $held "$LeaseReasonPrefix $runId" ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())) -eq 'theirs') { $why += "$($guest.Name) is leased until $($held.expiresUtc) - $($held.reason)"; continue }
                    if ($Kind) {
                        if (-not $refusals.ContainsKey($guest.Name)) { $refusals[$guest.Name] = Get-GuestRunRefusal $guest }
                        if ($refusals[$guest.Name]) { $why += $refusals[$guest.Name]; continue }
                        if ($i -gt 0) {
                            if (-not $startupBytes.ContainsKey($guest.Name)) { $startupBytes[$guest.Name] = [long](Get-VM -Name $guest.Name).MemoryStartup }
                            $need = $startupBytes[$guest.Name]
                            $available = Get-HostAvailableBytes
                            if (-not (Test-HostHasRoom -NeedBytes $need -AvailableBytes $available)) { $why += ("{0} waits for memory: the host has {1:N1} GB available, under its {2:N0} GB and the {3:N0} GB kept free" -f $guest.Name, ($available / 1GB), ($need / 1GB), ($TestbedHostReserveBytes / 1GB)); continue }
                        }
                    }
                    $start = Enter-TestbedStart -Look $look -Run $runId
                    if (-not $start.Go) { $gated = $true; $why = @($start.Why); break }
                    try {
                        $holder = Request-GuestLease $guest.Name 60
                        if ($null -eq $holder) {
                            # Its virtual processors, which the gate counts for the runs after this one: one
                            # Get-VM a run, with the lease had - none while it waited.
                            $cpu = 0
                            try { $cpu = [int](Get-VM -Name $guest.Name -ErrorAction Stop).ProcessorCount } catch { }
                            Register-TestbedStart -Start $start -Run $runId -VmName $guest.Name -VmProcessors $cpu
                            $leaseStart = $start
                            $takenGuest = $guest
                        }
                    }
                    finally { Exit-TestbedStart $start }
                    if ($null -ne $takenGuest) { break }
                    $why += "$($guest.Name) is leased until $($holder.expiresUtc) - $($holder.reason)"
                }
            }
            if ($null -eq $takenGuest) {
                $text = $why -join '; '
                if ($null -eq $waitingSince) { $waitingSince = Get-Date }
                # Said when what it waits for changes - while the gate holds it back, at most once a
                # minute, since its figures move on every turn.
                if ($text -ne $lastSaid -and (-not $gated -or ((Get-Date) - $lastSaidAt).TotalSeconds -ge 60)) {
                    Say (Format-TestbedWaitLine -Since $waitingSince -Why $text -First ($lastSaidAt -eq [DateTime]::MinValue))
                    Set-TestbedWaiting -What "Invoke-LiveTierOnGuest run $runId, for a guest" -Why $text -Since $waitingSince
                    $lastSaid = $text
                    $lastSaidAt = Get-Date
                }
                if ((Get-Date) -ge $deadline) { $stage = 'INFRA'; $stageWhy = "no guest came free in $LeaseWaitMinutes minutes - $text"; throw $stageWhy }
                [void](Wait-TestbedChange)
            }
        }
    }
    finally { Exit-TestbedQueue $ticket; if ($null -ne $waitingSince) { Remove-TestbedWaiting } }
    Say "gate: $(Format-TestbedGateOpen -Look $look -PendingPercent $leaseStart.PendingPercent -PendingCount $leaseStart.PendingCount)"
    if ($Kind) {
        Use-Guest $takenGuest
        $checkpointActions['start'] = $Checkpoint
        $checkpointActions['resting'] = $RestingCheckpoint
        $stepRefusal = Get-SignatureSettingsRefusal $SignatureSettings ([bool]$facts.Exchange) $startIsFrozen $GreenCheckpoint $Filter
        if ($stepRefusal) { $stage = 'INFRA'; $stageWhy = "the guest taken cannot take the step: $stepRefusal"; throw $stageWhy }
    }
    $timings['lease'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "lease taken on $($facts.Name) ($($timings['lease']))$(if ($Kind) { " - from $Checkpoint$(if ($startIsFrozen) { ', frozen' })" })"
    $credential = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $CredentialRepoRoot -VMName $facts.Name

    # ---- STAGE.
    $t0 = Get-Date
    try {
        $touched = $true
        if ($exchangeCarry) {
            if ((Get-TestbedVmState $facts.Name) -ne 'Running') { Start-VM -Name $facts.Name }
            Wait-TestbedHeartbeat -Id (Get-TestbedVmId $facts.Name) -Minutes 5 -What $facts.Name
            Say 'resumed the Exchange VM as its last run left it (its mailboxes live on the server; no checkpoint restored)'
        }
        else {
            Restore-Guest $Checkpoint
            Say "restored $Checkpoint"
        }
        $files = @((Join-Path $runDir 'payload\McpServer.zip'), (Join-Path $runDir 'payload\Tools.zip'), (Join-Path $runDir 'livetier\Source.zip'), (Join-Path $runDir 'livetier\NuGet.zip'), (Join-Path $runDir 'live-test-settings.json'))
        if ($WithAddIn) { $files += (Join-Path $runDir 'addin\AddIn.zip') }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Copy-ToGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -Path $files -Destination $GuestPayloadDir *> (Join-Path $runDir 'copy.log')
        if ($WithAddIn) {
            # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
            & (Join-Path $PSScriptRoot 'Copy-ToGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -Path $VstoRuntimePath -Destination "$GuestRoot\media" *>> (Join-Path $runDir 'copy.log')
        }
        # The commit's guest scripts - but the crash-dump arming is THIS checkout's, whatever commit is
        # tested, so an older commit (a bisection) is dumped the same way.
        $guestScripts = @(Get-ChildItem -LiteralPath (Join-Path $tree 'Testbed\guest') -File | Where-Object { $_.Name -ne $CrashDumpScript -and $_.Name -ne $FirstChanceScript } | ForEach-Object { $_.FullName })
        $guestScripts += [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\guest\$CrashDumpScript"))
        # The hunt's debugger, THIS checkout's too - and on the guest at all only in a hunting run.
        if ($HuntFirstChance) { $guestScripts += [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\guest\$FirstChanceScript")) }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Copy-ToGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -Path $guestScripts -Destination $GuestRoot *>> (Join-Path $runDir 'copy.log')
        # What must not be running while the payload is swapped. The Exchange VM's Outlook stays up: its
        # resting checkpoint holds it signed in, and it uses nothing the swap replaces.
        $swapGuard = 'OUTLOOK,OutlookAI.RemediationTools,OutlookAI.McpServer,OutlookAI.ComHost,testhost,dotnet'
        if ($facts.Exchange) { $swapGuard = 'OutlookAI.RemediationTools,OutlookAI.McpServer,OutlookAI.ComHost,testhost,dotnet' }
        $swap = @"
`$ErrorActionPreference = 'Stop'
if (@(Get-Process -Name $swapGuard -ErrorAction SilentlyContinue).Count -gt 0) { throw 'Outlook, the tools, the server or dotnet is running - not replacing anything' }
if ('$($facts.Exchange)' -eq 'True') { Get-ChildItem -LiteralPath '$GuestRoot' -Directory | Where-Object { `$_.Name -match '\.pre-\d{8}T\d{6}Z`$' } | Remove-Item -Recurse -Force }
`$stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
foreach (`$name in 'server', 'tools', 'src', 'nuget-offline') { `$from = Join-Path '$GuestRoot' `$name; if (Test-Path -LiteralPath `$from) { Rename-Item -LiteralPath `$from -NewName "`$name.pre-`$stamp" } }
Expand-Archive -LiteralPath '$GuestPayloadDir\McpServer.zip' -DestinationPath '$GuestRoot\server' -Force
Expand-Archive -LiteralPath '$GuestPayloadDir\Tools.zip' -DestinationPath '$GuestRoot\tools' -Force
Expand-Archive -LiteralPath '$GuestPayloadDir\Source.zip' -DestinationPath '$GuestRoot\src' -Force
Expand-Archive -LiteralPath '$GuestPayloadDir\NuGet.zip' -DestinationPath '$GuestRoot\nuget-offline' -Force
if (Test-Path -LiteralPath '$GuestPayloadDir\AddIn.zip') { if (Test-Path -LiteralPath '$GuestRoot\addin') { Remove-Item -LiteralPath '$GuestRoot\addin' -Recurse -Force }; Expand-Archive -LiteralPath '$GuestPayloadDir\AddIn.zip' -DestinationPath '$GuestRoot\addin' -Force }
New-Item -ItemType Directory -Force -Path (Split-Path -Parent '$GuestSettingsPath') | Out-Null
Copy-Item -LiteralPath '$GuestPayloadDir\live-test-settings.json' -Destination '$GuestSettingsPath' -Force
"SWAP-DONE settings `$((Get-FileHash '$GuestSettingsPath').Hash)"
"@
        $o = Invoke-Guest 'swap' $swap 900 -Session0
        if ($o -notmatch 'SWAP-DONE') { throw 'the payload swap did not finish - guest\swap.log' }
        # Crash dumps: armed after the restore, on every run. A guest that cannot be armed is still
        # tested - the dump is evidence, not a verdict - and the summary says it was not armed.
        $o = Invoke-Guest 'crash-dumps-arm' "& '$GuestRoot\$CrashDumpScript' -Execute -DumpFolder '$GuestDumpDir'; `"DUMPS-EXIT `$LASTEXITCODE`"" 300 -Session0
        $crashDumpState['armed'] = ($o -match 'VERDICT: ARMED' -and $o -match 'DUMPS-EXIT 0')
        if ($crashDumpState['armed']) { Say "crash dumps armed: a crashing OUTLOOK.EXE leaves a full dump in $GuestDumpDir" }
        else { Say 'NOTE      crash dumps NOT armed - guest\crash-dumps-arm.log; the run goes on without them' }
        if ($ProveCrashDumps -and $crashDumpState['armed']) {
            $o = Invoke-Guest 'crash-dumps-prove' "& '$GuestRoot\$CrashDumpScript' -Prove -DumpFolder '$GuestDumpDir'; exit `$LASTEXITCODE" 600
            $crashDumpState['proven'] = ($o -match 'VERDICT: PROVEN')
            Say "crash dumps proven: $($crashDumpState['proven']) - guest\crash-dumps-prove.log"
        }
        $sdkNamed = "-InstallerPath '$GuestRoot\media\$SdkInstallerName' -ExpectedSdkVersionPrefix '$SdkVersionPrefix'"
        $o = Invoke-Guest 'sdk-execute' "Set-Location '$GuestRoot'; & .\Install-DotnetSdk.ps1 $sdkNamed -ExpectedSha512 '$SdkSha512' -Execute; `"SDK-EXECUTE-EXIT `$LASTEXITCODE`"" 1800 -Session0
        if ($o -notmatch 'SDK-EXECUTE-EXIT 0') { throw 'Install-DotnetSdk.ps1 -Execute failed - guest\sdk-execute.log' }
        $o = Invoke-Guest 'sdk-verify' "Set-Location '$GuestRoot'; & .\Install-DotnetSdk.ps1 $sdkNamed -Verify; `"SDK-VERIFY-EXIT `$LASTEXITCODE`"" 1800 -Session0
        if ($o -notmatch 'VERDICT: TEST-READY') { $stage = 'BUILD'; $stageWhy = 'the suite is not TEST-READY on the guest - it did not build or enumerate there (guest\sdk-verify.log)'; throw $stageWhy }
    }
    catch { if (-not $stage) { $stage = 'INFRA'; $stageWhy = "staging failed: $($_.Exception.Message)" }; throw }
    $timings['stage'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "staged: TEST-READY ($($timings['stage']))"

    # ---- PREPARE: from a frozen checkpoint the clock guard only; otherwise a graceful restart and the hub rebuild.
    $t0 = Get-Date
    $prepared = @(Get-PrepareSteps -StartIsFrozen $startIsFrozen -SkipHubReset ([bool]$SkipHubReset) -SignatureSettings $SignatureSettings)
    try {
        Use-Lease 30
        if ($WithAddIn -and ($facts.Exchange -or -not $startIsFrozen)) { throw '-WithAddIn installs the add-in on a frozen-clock Outlook guest only' }
        if ($facts.Exchange) {
            if (-not $exchangeCarry) {
                # A restored checkpoint: its .ost files are older than the server. Outlook quits, Windows
                # restarts, the files go aside, and the Outlook started below downloads the mailboxes afresh.
                Close-GuestSession
                # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
                & (Join-Path $PSScriptRoot 'Restart-Guest.ps1') -VMName $facts.Name -Execute -RepoRoot $CredentialRepoRoot -LogPath (Join-Path $runDir 'restart.log') *> (Join-Path $runDir 'restart.out.log')
                if ($LASTEXITCODE -ne 0) { throw "Restart-Guest.ps1 did not complete ($LASTEXITCODE) - restart.log" }
                # New data files are new stores to Windows Search, and the old files' roots would stay beside
                # theirs - a name with two roots is one the live tier's scope finder never guesses between
                # (T2/LiveIndexScopes.cs). So the catalog is rebuilt from nothing after the set-aside.
                $o = Invoke-Guest 'reset-ost' "& '$GuestRoot\Reset-OutlookOfflineStore.ps1' -Execute -ResetSearchIndex; `"RESET-OST-EXIT `$LASTEXITCODE`"" 300 -Session0
                if ($o -notmatch 'RESET-OST-EXIT 0' -or $o -notmatch 'SEARCH-INDEX-RESET') { throw 'Reset-OutlookOfflineStore.ps1 did not set the data files aside and rebuild the search index - guest\reset-ost.log' }
            }
            $o = Invoke-Guest 'outlook-running' "if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) { 'OUTLOOK-RUNNING yes' } else { 'OUTLOOK-RUNNING no' }" 120 -Session0
            if ($o -match 'OUTLOOK-RUNNING no') {
                # Defender first, so the Outlook about to start finds its signatures current.
                # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
                & (Join-Path $PSScriptRoot 'Invoke-ExchangeSignIn.ps1') -VMName $facts.Name -Mode Preflight -DefenderOnly -RepoRoot $CredentialRepoRoot *> (Join-Path $runDir 'preflight-defender.log')
                if ($LASTEXITCODE -ne 0) { throw "Defender's signatures could not be brought up to date before Outlook started (exit $LASTEXITCODE) - preflight-defender.log" }
                $o = Invoke-Guest 'start-outlook' "& '$GuestRoot\Start-OutlookUnelevated.ps1' -Profile 'Outlook' -MinUptimeSeconds 240; `"START-OUTLOOK-EXIT `$LASTEXITCODE`"" 900 -Session0
                if ($o -notmatch 'START-OUTLOOK-EXIT 0') { throw 'Start-OutlookUnelevated.ps1 did not start Outlook NOT elevated on the profile Outlook - guest\start-outlook.log' }
                $outlookStart = Get-OutlookStartFacts $o 'Outlook'
                Say ((Format-OutlookStartLine $outlookStart) -replace '^outlook\s+', 'Outlook ')
            }
            if (-not $exchangeCarry) {
                # Every indexed store needs exactly one root, with something below it, before an index test
                # runs. Bounded; a store not ready at the end is the tests' to report.
                $o = Invoke-Guest 'index-roots' "& '$GuestRoot\Reset-OutlookOfflineStore.ps1' -WaitForIndexRoots -TimeoutSeconds 1200" 1500
                $m = [regex]::Match($o, 'INDEX-ROOTS-(READY|TIMEOUT) (\d+)')
                $indexRoots = 'not reported (guest\index-roots.log)'
                if ($m.Success -and $m.Groups[1].Value -eq 'READY') { $indexRoots = "every indexed store has one root, with rows, $($m.Groups[2].Value) s after Outlook started on the rebuilt index" }
                elseif ($m.Success) { $indexRoots = "NOT every indexed store has one root with rows after $($m.Groups[2].Value) s - index tests may fail for it" }
                Say "index      $indexRoots"
            }
            # Defender current and every account SIGNED-IN, or no run (Testbed/README.md section 4e).
            # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
            & (Join-Path $PSScriptRoot 'Invoke-ExchangeSignIn.ps1') -VMName $facts.Name -Mode Preflight -RepoRoot $CredentialRepoRoot *> (Join-Path $runDir 'preflight.log')
            if ($LASTEXITCODE -ne 0) { throw "Invoke-ExchangeSignIn.ps1 -Mode Preflight did not pass (exit $LASTEXITCODE) - preflight.log: Defender's signatures, or an account no longer signed in (-Mode Reauth -Account <address>)" }
            $prepared = @(Get-ExchangeStartPlan (-not $exchangeCarry))
        }
        elseif ($startIsFrozen) {
            # ps51-native-stderr-ok: a PowerShell script, not a program - it starts no program at all
            & (Join-Path $PSScriptRoot 'Set-GuestClockFrozen.ps1') -VMName $facts.Name -Verify -RepoRoot $CredentialRepoRoot -LogPath (Join-Path $runDir 'frozen-clock.log') *> (Join-Path $runDir 'frozen-clock.out.log')
            if ($LASTEXITCODE -ne 0) { throw "the frozen-clock guard did not say FROZEN (exit $LASTEXITCODE) - the guest's clock left its frozen instant, or this is not its frozen checkpoint: frozen-clock.log" }
            # The tested add-in first, whose first start opens and closes Outlook; the signature setting after it, so
            # that the Outlook start just below is the first to read the setting.
            if ($WithAddIn) {
                # The tested commit's add-in over the checkpoint's: installed as the guest's default task, elevated
                # (exit 2, INSTALLED-NEVER-RAN), then started once NOT elevated (ADDIN-READY), and Outlook closed again.
                $o = Invoke-Guest 'addin-install' "& '$GuestRoot\Register-InteractiveTask.ps1' -RunLevel Highest -Script `"& '$GuestRoot\Install-OutlookAIAddIn.ps1' -Phase Install -Execute`"; `"ADDIN-INSTALL-EXIT `$LASTEXITCODE`"" 900 -Session0
                if ($o -notmatch 'ADDIN-INSTALL-EXIT 2') { throw 'the tested commit''s add-in did not install - guest\addin-install.log' }
                $o = Invoke-Guest 'addin-firstrun' "& '$GuestRoot\Install-OutlookAIAddIn.ps1' -Phase FirstRun -Execute; exit `$LASTEXITCODE" 900
                if ($o -notmatch 'ADDIN-READY') { throw 'the tested commit''s add-in did not reach ADDIN-READY - guest\addin-firstrun.log' }
                $o = Invoke-Guest 'addin-outlook-closed' "`$until = (Get-Date).AddMinutes(4); while ((Get-Process OUTLOOK -ErrorAction SilentlyContinue) -and (Get-Date) -lt `$until) { Start-Sleep -Seconds 5 }; if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { 'OUTLOOK-STILL-RUNNING' } else { 'OUTLOOK-CLOSED' }" 300 -Session0
                if ($o -notmatch 'OUTLOOK-CLOSED') { throw 'Outlook did not close after the add-in''s first start - guest\addin-outlook-closed.log' }
                # What tells the tests that drive the add-in that this run's add-in is the tested commit's: a file in the
                # suite's source folder, which the next run's swap renames away whatever checkpoint it starts from.
                $o = Invoke-Guest 'addin-marker' "Set-Content -LiteralPath '$GuestRoot\src\.tested-addin' -Value '$sha' -Encoding ascii; 'ADDIN-MARKED'" 120 -Session0
                if ($o -notmatch 'ADDIN-MARKED') { throw 'the tested add-in could not be marked for the tests - guest\addin-marker.log' }
                Say "add-in     the tested commit's, installed over the checkpoint's: ADDIN-READY"
            }
            if ($SignatureSettings) {
                # Q200 (a), Q201: the setting, its test signatures and the hub's own defaults, made elevated while
                # Outlook is closed - Outlook reads them when it starts, just below. The restore at REST undoes them.
                $o = Invoke-Guest 'signature-settings' "& '$GuestRoot\$SignatureSettingsScript' -Scenario $SignatureSettings -ValueKind $SignatureValueKind -Execute; `"SIGNATURE-SETTINGS-EXIT `$LASTEXITCODE`"" 900 -Session0
                $signatureSettingsState['set'] = ($o -match 'VERDICT: SET' -and $o -match 'SIGNATURE-SETTINGS-EXIT 0')
                if (-not $signatureSettingsState['set']) { throw "the signature setting $SignatureSettings was not made before Outlook started - guest\signature-settings.log" }
                Say "signature  $SignatureSettings made before Outlook starts, its values $($SignatureValueKind): the setting, its test signatures and the hub's own defaults (guest\signature-settings.log)"
            }
            if ($HuntWithoutAddIn) {
                # The hunt without OutlookAI's own add-in (2026-11-18): its LoadBehavior 0 for this run, in the session
                # Outlook runs in and before it starts, so no Outlook of the run loads it. The restore at REST puts the
                # checkpoint's value back.
                $o = Invoke-Guest 'addin-off' (Get-AddInOffScript $AddInRegistryKey) 120
                if ($o -notmatch 'ADDIN-OFF 0 ') { throw 'OutlookAI''s add-in could not be turned off for the hunt - guest\addin-off.log' }
                $huntAddIn = 'off: ' + ([regex]::Match($o, 'ADDIN-OFF [^\r\n]*').Value)
                Say "add-in     OutlookAI's own turned off for this run - $huntAddIn; the checkpoint's value comes back at REST"
            }
            # The state step 9a hands the suite over in, without its rebuild: Outlook running NOT elevated on
            # the tier profile, its window up. The frozen checkpoint holds Outlook closed so STAGE could swap.
            $o = Invoke-Guest 'start-outlook' "& '$GuestRoot\Start-OutlookUnelevated.ps1' -Profile '$TierProfileName'; `"START-OUTLOOK-EXIT `$LASTEXITCODE`"" 900 -Session0
            if ($o -notmatch 'START-OUTLOOK-EXIT 0') { throw 'Start-OutlookUnelevated.ps1 did not start Outlook NOT elevated on the tier profile - guest\start-outlook.log' }
            $outlookStart = Get-OutlookStartFacts $o $TierProfileName
            Say ((Format-OutlookStartLine $outlookStart) -replace '^outlook\s+', 'Outlook ')
        }
        else {
            Close-GuestSession
            # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
            & (Join-Path $PSScriptRoot 'Restart-Guest.ps1') -VMName $facts.Name -Execute -CancelLogonPrompt -RepoRoot $CredentialRepoRoot -LogPath (Join-Path $runDir 'restart.log') *> (Join-Path $runDir 'restart.out.log')
            if ($LASTEXITCODE -ne 0) { throw "Restart-Guest.ps1 did not complete ($LASTEXITCODE) - restart.log" }
            if (-not $SkipHubReset) {
                $o = Invoke-Guest 'reset-hub' "& '$GuestRoot\Reset-HubPopulation.ps1' -Execute; exit `$LASTEXITCODE" 3600
                if ((Get-TaskExit $o) -ne 0) { throw 'the hub rebuild did not succeed - guest\reset-hub.log' }
            }
        }
    }
    catch { if (-not $stage) { $stage = 'INFRA'; $stageWhy = "preparing the guest failed: $($_.Exception.Message)" }; throw }
    $timings['prepare'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "prepared: $($prepared -join ', ') ($($timings['prepare']))"

    # ---- RUN: the suite, NOT elevated, opted in for this run only.
    $t0 = Get-Date
    $markerLines = ''
    if ($SignatureSettings) { $markerLines = "`$env:$SignatureSettingsVariable = '$SignatureSettings'`r`n" }
    $dialogLines = Get-SignatureDialogLines $SignatureSettings $GuestRoot $guestRunDir ([bool]$SignatureDialog)
    if ($facts.Exchange) {
        $markerLines = "`$env:OUTLOOKAI_LIVE_RUN_MARKER = '$runMarker'`r`n`$env:OUTLOOKAI_LIVE_PRIOR_RUN_MARKERS = '$($priorMarkers -join ',')'`r`n"
        # Recorded BEFORE the suite starts: a run that dies half-way still leaves its marker for the next one.
        Add-Content -LiteralPath $ledgerPath -Value (ConvertTo-LedgerLine $runMarker $runId 'started' $null $priorMarkers ([DateTime]::UtcNow)) -Encoding UTF8
    }
    if ($HuntFirstChance) {
        $liveScript = Get-HuntSuiteScript -OptIn $facts.ComputerName -GuestRunDir $guestRunDir -GuestRoot $GuestRoot -Filters $huntFilters -Rounds $HuntRounds -Minutes $HuntMinutes -Module $HuntModule -DumpDir $GuestDumpDir -ScriptName $FirstChanceScript -Variant $HuntVariant -LegMinutes $HuntLegMinutes -ShowEach ([bool]$HuntShowEach) -TypeEach ([bool]$HuntTypeEach) -Ablate $HuntAblate -PromoteAfter ([bool]$HuntPromoteAfter) -ReleaseAfterClose ([bool]$HuntReleaseAfterClose)
        Say "hunt: a debugger on every OUTLOOK.EXE, dumping the first first-chance access violation in $HuntModule; variant $HuntVariant, at most $HuntRounds round(s) of $(@($huntFilters.Keys) -join ' and ') in $HuntMinutes minute(s)$(if (Test-HuntLoopVariant $HuntVariant) { ", $HuntLegMinutes a leg" })"
    }
    else {
    $liveScript = @"
$markerLines`$env:OUTLOOKAI_LIVE_OPT_IN = '$($facts.ComputerName)'
`$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'; `$env:DOTNET_NOLOGO = '1'; `$env:MSBUILDDISABLENODEREUSE = '1'
New-Item -ItemType Directory -Force -Path '$guestRunDir', '$guestRunDir\diag' | Out-Null
$($dialogLines)Set-Location '$GuestRoot\src'
& dotnet test McpServer\OutlookAI.McpServer.Tests\OutlookAI.McpServer.Tests.csproj -c Release --filter "$runFilter" --logger "trx;LogFileName=live.trx" --logger "console;verbosity=detailed" --results-directory '$guestRunDir' --diag '$guestRunDir\diag\vstest.log' *> '$guestRunDir\console.txt'
exit `$LASTEXITCODE
"@
    }
    if ($HuntFirstChance -and $HuntTraceClr) {
        $o = Invoke-Guest 'clr-trace-start' (Get-ClrTraceStartScript $guestRunDir $ClrTraceSession) 120 -Session0
        $huntClrTrace = $(if ($o -match 'CLR-TRACE-START 0') { 'traced' } else { 'not started - guest\clr-trace-start.log' })
        Say "hunt: the .NET runtime's exception events $huntClrTrace"
    }
    $runStartUtc = [DateTime]::UtcNow
    # The GUEST's clock at the start, for the crash events below, which carry the guest's time: on a
    # frozen guest (Q130 (a)) the host's clock is hours or days ahead of it, and filtering the guest's
    # Application log from the host's time found no crash at all (measured 2026-10-03).
    $guestRunStartUtc = ''
    try { $guestRunStartUtc = [string](Invoke-Command -Session (Get-GuestSession) -ErrorAction Stop -ScriptBlock { [DateTime]::UtcNow.ToString('o') }) }
    catch { Say "  (the guest's clock could not be read before the suite - the crash count falls back to the host's: $($_.Exception.Message))" }
    $suiteStarted = $true
    $o = Invoke-Guest 'live' $liveScript ($RunTimeoutMinutes * 60)
    $suiteExit = Get-TaskExit $o
    $suiteSeconds = ((Get-Date) - $t0).TotalSeconds
    $timings['suite'] = Format-Seconds $suiteSeconds
    Say "suite finished, exit $suiteExit ($($timings['suite']))"
    if ($HuntFirstChance -and $HuntTraceClr -and $huntClrTrace -eq 'traced') {
        $o = Invoke-Guest 'clr-trace-stop' (Get-ClrTraceStopScript $guestRunDir $ClrTraceSession) 900 -Session0
        $cm = [regex]::Match($o, 'CLR-TRACE-STOP (-?\d+) (\d*)')
        $huntClrTrace = $(if ($cm.Success -and $cm.Groups[2].Value) { "traced: clr-events.xml, $($cm.Groups[2].Value) bytes" } else { 'traced, but tracerpt made no XML - guest\clr-trace-stop.log' })
        Say "hunt: the .NET runtime's exception events $huntClrTrace"
    }

    # ---- FETCH: results, console, logs, crash events.
    $resultsDir = Join-Path $runDir 'results'
    # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
    & (Join-Path $PSScriptRoot 'Copy-FromGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -GuestPath $guestRunDir -Include '*.trx', 'console.txt', 'signature-dialog.*', 'first-chance.log', 'clr-events.xml' -Destination $resultsDir -Force *> (Join-Path $runDir 'fetch.log')
    # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
    & (Join-Path $PSScriptRoot 'Copy-FromGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -GuestPath "$guestRunDir\diag" -Include '*.log' -Destination (Join-Path $resultsDir 'diag') -Force *>> (Join-Path $runDir 'fetch.log')
    $trxFile = Get-ChildItem -LiteralPath $resultsDir -Filter '*.trx' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($HuntFirstChance) {
        $trxFiles = @(Get-ChildItem -LiteralPath $resultsDir -Filter '*.trx' -Recurse -ErrorAction SilentlyContinue | Sort-Object Name)
        if ($trxFiles.Count -gt 0) { $trx = Merge-TrxResults @($trxFiles | ForEach-Object { [pscustomobject]@{ Label = $_.BaseName; Result = (ConvertFrom-TrxText ([System.IO.File]::ReadAllText($_.FullName))) } }) }
    }
    elseif ($trxFile) { $trx = ConvertFrom-TrxText ([System.IO.File]::ReadAllText($trxFile.FullName)) }
    $consoleFile = Get-ChildItem -LiteralPath $resultsDir -Filter 'console.txt' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($consoleFile) { $consoleFacts = Read-ConsoleFacts ([System.IO.File]::ReadAllText($consoleFile.FullName)) }
    if ($HuntFirstChance -and $consoleFile) {
        $hr = @([regex]::Matches([System.IO.File]::ReadAllText($consoleFile.FullName), '(?m)^HUNT-ROUNDS (\d+)'))
        if ($hr.Count -gt 0) { $huntRoundsRun = [int]$hr[$hr.Count - 1].Groups[1].Value }
        $huntOffice = Read-HuntOffice ([System.IO.File]::ReadAllText($consoleFile.FullName))
    }
    $since = Get-CrashEventsSince -GuestRunStartUtc $guestRunStartUtc -HostRunStartUtc $runStartUtc
    # Event 1001 too, whole: WER's report of each crash - its kind and where it filed it - which is how a crash
    # that left an Application Error event and no dump can be told apart (2026-11-18, runbook section 4.15).
    # WER-REPORTS counts the reports the events name: WER logs 1001 more than once for one report.
    $o = Invoke-Guest 'crash-events' "`$s = ([datetime]'$since').ToLocalTime(); `$e = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = `$s; Id = 1000 } -ErrorAction SilentlyContinue | Where-Object { `$_.Message -match 'OUTLOOK\.EXE' }); `"CRASHES `$(`$e.Count)`"; `$e | ForEach-Object { `$_.TimeCreated.ToString('o') + ' ' + ((`$_.Message -split [char]10 | Select-Object -First 4) -join ' | ') }; `$w = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = `$s; Id = 1001 } -ErrorAction SilentlyContinue | Where-Object { `$_.Message -match 'OUTLOOK\.EXE' }); `$ids = @(`$w | ForEach-Object { if (`$_.Message -match 'Report Id: ([0-9a-fA-F-]+)') { `$Matches[1] } } | Sort-Object -Unique); `"WER-REPORTS `$(`$ids.Count)`"; `$w | ForEach-Object { 'WER-REPORT ' + `$_.TimeCreated.ToString('o') + ' ' + ((`$_.Message -split [char]10 | ForEach-Object { `$_.Trim() } | Where-Object { `$_ }) -join ' | ') }" 300 -Session0
    $cm = [regex]::Match($o, 'CRASHES (\d+)')
    if ($cm.Success) { $crashes = [int]$cm.Groups[1].Value }
    $wm = [regex]::Match($o, 'WER-REPORTS (\d+)')
    if ($wm.Success) { $crashDumpState['werReports'] = [int]$wm.Groups[1].Value }
    Save-CrashDumps
    $crashes = Merge-CrashCount $crashes $script:dumpsFound
    if ($facts.Exchange) {
        # Clean when a sweep ran and left nothing this run is answerable for - its own items and the
        # priors it was handed are then gone, and the ledger stops handing them on.
        $ledgerOutcome = ($null -ne $consoleFacts -and $consoleFacts.SweepRan -and $null -ne $consoleFacts.ArtifactsLeft -and $consoleFacts.ArtifactsLeft -eq 0)
        Add-Content -LiteralPath $ledgerPath -Value (ConvertTo-LedgerLine $runMarker $runId 'finished' $ledgerOutcome $priorMarkers ([DateTime]::UtcNow)) -Encoding UTF8
        Say "ledger: $runMarker finished, clean=$ledgerOutcome"
    }
    if ($suiteSeconds -ge ($RunTimeoutMinutes * 60 - 60)) { $stage = 'INFRA'; $stageWhy = "the suite did not finish within $RunTimeoutMinutes minutes" }

    # ---- GREEN CHECKPOINT, when asked and earned.
    if ($GreenCheckpoint -and (Get-LiveVerdict $stage $stageWhy $trx $consoleFacts $crashes $suiteExit $narrowed).Verdict -eq 'PASS') {
        $freeGb = [math]::Round((Get-PSDrive -Name E -ErrorAction SilentlyContinue).Free / 1GB, 1)
        if ($freeGb -lt $MinFreeGbForCheckpoint) { Say "NOT taking ${GreenCheckpoint}: E: has $freeGb GB free, under $MinFreeGbForCheckpoint" }
        else {
            Use-Lease 30
            Checkpoint-VM -Name $facts.Name -SnapshotName $GreenCheckpoint
            $checkpointActions['green'] = $GreenCheckpoint
            Say "checkpoint $GreenCheckpoint taken of the green run (E: had $freeGb GB free)"
        }
    }
}
catch {
    if (-not $stage) { $stage = 'INFRA'; $stageWhy = $_.Exception.Message }
    Say "STOPPED: $stageWhy"
}
finally {
    # ---- REST: the resting checkpoint, its settings, saved; the lease released. Every path, each
    # step on its own, so one that fails does not leave the next undone.
    if ($touched) {
        $mayTouch = $true
        try { Use-Lease 30 } catch { $mayTouch = $false; Say "NOT resting the guest: $($_.Exception.Message)" }
        if ($mayTouch) {
            # A run that stopped after the suite started still keeps what WER wrote.
            if ($suiteStarted) { Save-CrashDumps }
            $restOn = $RestingCheckpoint
            if ($RestOnGreen -and $checkpointActions['green']) { $restOn = $GreenCheckpoint }
            if ($exchangeRestCarry -and -not ($RestOnGreen -and $checkpointActions['green'])) {
                # The Exchange VM rests as the run left it: a restored .ost would be older than the server.
                $checkpointActions['restedOn'] = 'its own state, carried forward (no checkpoint restored)'
            }
            else {
                try {
                    Restore-Guest $restOn
                    $checkpointActions['restedOn'] = $restOn
                }
                catch { Say "RESTORING $restOn FAILED: $($_.Exception.Message)" }
            }
            $restSettings = Join-Path $runDir 'live-test-settings.json'
            if ($checkpointActions['restedOn'] -and (Test-Path -LiteralPath $restSettings)) {
                try {
                    # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
                    & (Join-Path $PSScriptRoot 'Copy-ToGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -Path $restSettings -Destination $GuestSettingsPath *>> (Join-Path $runDir 'copy.log')
                    $checkpointActions['settingsStaged'] = $true
                }
                catch { Say "staging the settings on $restOn failed: $($_.Exception.Message)" }
            }
            try {
                Save-VM -Name $facts.Name
                $checkpointActions['saved'] = ((Get-TestbedVmState $facts.Name) -eq 'Saved')
            }
            catch { Say "SAVING THE GUEST FAILED: $($_.Exception.Message) - save it by hand." }
            Say "rested on $($checkpointActions['restedOn']); settings staged=$($checkpointActions['settingsStaged']); saved=$($checkpointActions['saved'])"
        }
    }
    Close-GuestSession
    try {
        if ((Get-LeaseDecision (Get-TestbedLease -VMName $facts.Name) "$LeaseReasonPrefix $runId" ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())) -eq 'ours') {
            & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $facts.Name -Release | Out-Null
            Say 'lease released'
        }
    }
    catch { Say "releasing the lease failed: $($_.Exception.Message)" }
    if (Test-Path -LiteralPath $tree) {
        Invoke-NativeCommand { & git -C $repo worktree remove --force $tree 2>&1 } | Out-Null
        # Measured 2026-10-04: git unregisters the tree but leaves its directory - about 180 MB of build
        # output a run - on every run (and on every run of the agent before). What it left goes here.
        if (Test-Path -LiteralPath $tree) {
            Remove-Item -LiteralPath $tree -Recurse -Force -ErrorAction SilentlyContinue
            Invoke-NativeCommand { & git -C $repo worktree prune 2>&1 } | Out-Null
            if (Test-Path -LiteralPath $tree) { Say "NOTE      the commit's build tree could not be deleted: $tree" }
        }
    }
    # The payloads are the commit's, rebuilt by any later run; only what the run found is kept.
    foreach ($built in 'payload', 'livetier') {
        $path = Join-Path $runDir $built
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# ---- REPORT.
$verdict = Get-LiveVerdict $stage $stageWhy $trx $consoleFacts $crashes $suiteExit $narrowed
$timings['total'] = Format-Seconds ((Get-Date) - $started).TotalSeconds
$summary = New-LiveRunSummary -RunId $runId -Verdict $verdict -Revision $revision -Guest $facts.Name -Filter $runFilter -Narrowed $narrowed -Trx $trx -Console $consoleFacts -Crashes $crashes -SuiteExit $suiteExit -Checkpoints $checkpointActions -Timings $timings -Results $runDir -CrashDumps $crashDumpState
if ($facts.Exchange) { $summary['runMarker'] = [ordered]@{ marker = $runMarker; priors = @($priorMarkers); ledger = $ledgerPath; clean = $ledgerOutcome } }
$summary['outlookStart'] = $outlookStart
$summary['signatureSettings'] = $signatureSettingsState
$summary['indexRoots'] = $indexRoots
$huntDumps = @(@($crashDumpState['files']) | Where-Object { [string]$_['name'] -match 'chance\.dmp$' } | ForEach-Object { [string]$_['name'] })
if ($HuntFirstChance) { $summary['hunt'] = [ordered]@{ module = $HuntModule; variant = $HuntVariant; showEach = [bool]$HuntShowEach; typeEach = [bool]$HuntTypeEach; legMinutes = $(if (Test-HuntLoopVariant $HuntVariant) { $HuntLegMinutes } else { $null }); legs = $huntFilters; roundsAsked = $HuntRounds; minutes = $HuntMinutes; roundsRun = $huntRoundsRun; dumps = $huntDumps; office = $huntOffice; addIn = $huntAddIn; ablate = $(if ($HuntAblate) { $HuntAblate } else { $null }); clrTrace = $huntClrTrace; log = "results\$($facts.Name)\first-chance.log" } }
# ToArray(), never @(...): @() of a List[object] given to an [object[]] parameter fails with "Argument
# types do not match" under Windows PowerShell 5.1 and PowerShell 7.7 alike (measured 2026-11-18).
$summary['guestClock'] = New-GuestClockSummaryJson $guestClockRecords.ToArray()
[System.IO.File]::WriteAllText((Join-Path $runDir 'summary.json'), ($summary | ConvertTo-Json -Depth 8))

$lines = New-Object System.Collections.Generic.List[string]
$head = "== $runId - $($verdict.Verdict) (exit $($ExitCodes[$verdict.Verdict]))"
if ($verdict.Why) { $head += " - $($verdict.Why)" }
$lines.Add("$head ==")
$lines.Add("guest      $($facts.Name) from $Checkpoint; filter $runFilter")
if ($facts.Exchange) { $lines.Add("marker     $runMarker (ledger clean=$ledgerOutcome); earlier runs swept by marker: $(if ($priorMarkers.Count -gt 0) { $priorMarkers -join ', ' } else { 'none' })") }
$rev = "revision   $($sha.Substring(0, 12)) - $subject"
if ($dirtyCount -gt 0) { $rev += " ($dirtyCount uncommitted change(s) NOT included)" }
$lines.Add($rev)
if ($null -ne $trx) {
    $suiteLine = "suite      $($trx.Counts['total']) total: $($trx.Counts['passed']) passed, $($trx.Counts['failed']) failed, $($trx.Counts['skipped']) skipped"
    if ($trx.Counts['other'] -gt 0) { $suiteLine += ", $($trx.Counts['other']) other" }
    $lines.Add("$suiteLine; dotnet test exit $suiteExit")
}
else { $lines.Add('suite      no TRX file') }
if ($null -ne $consoleFacts) {
    $artifactText = 'NOT PROVEN - no artifact sweep ran in this selection'
    if ($consoleFacts.SweepRan) { $artifactText = "most left in a store after a sweep: $($consoleFacts.ArtifactsLeft)" }
    $censusText = 'NOT REPORTED'
    if ($null -ne $consoleFacts.Tripwire) { $censusText = "$($consoleFacts.Tripwire) failure(s)" }
    $lines.Add("safety     artifacts: $artifactText; tripwire census: $censusText; PROVED NOTHING: $(@($consoleFacts.ProvedNothing).Count)")
}
$crashText = 'could not be read'
if ($crashes -ge 0) { $crashText = [string]$crashes }
$dumpText = 'NOT armed'
if ($crashDumpState['armed']) { $dumpText = 'armed' }
$dumpList = @($crashDumpState['files'])
if ($dumpList.Count -gt 0) {
    $dumpText += ", $(@($dumpList | Where-Object { $_['fetched'] }).Count) of $($dumpList.Count) fetched: " + (($dumpList | ForEach-Object { "dumps\$($facts.Name)\$($_['name']) ($([math]::Round($_['bytes'] / 1MB)) MB)" }) -join ', ')
}
elseif ($crashDumpState['armed']) { $dumpText += ', none written' }
$werText = ''
$werList = @($crashDumpState['werFiles'])
if ([int]$crashDumpState['werReports'] -gt 0 -or $werList.Count -gt 0) {
    $werText = "; WER's own reports: $([Math]::Max(0, [int]$crashDumpState['werReports'])) (guest\crash-events.log)"
    if ($werList.Count -gt 0) { $werText += ", $(@($werList | Where-Object { $_['fetched'] }).Count) of $($werList.Count) file(s) fetched to wer\$($facts.Name)\" }
}
$lines.Add("outlook    crashes during the run: $crashText; crash dumps: $dumpText$werText")
if ($HuntFirstChance) {
    $roundsText = 'not read'
    if ($huntRoundsRun -ge 0) { $roundsText = [string]$huntRoundsRun }
    $lines.Add("hunt       variant $HuntVariant$(if ($HuntTypeEach) { ', a window shown and typed into before each draft' } elseif ($HuntShowEach) { ', a window shown before each draft' })$(if ($HuntReleaseAfterClose) { ', its Word document released after the window closed' })$(if ($HuntPromoteAfter) { ", each of the product's drafts re-opened and promoted after" })$(if ($HuntAblate) { ", the typed-text tests without their $HuntAblate step" })$(if ($huntClrTrace) { ", the .NET runtime's exceptions $huntClrTrace" }), OutlookAI's add-in $huntAddIn, $roundsText of at most $HuntRounds round(s) of $(@($huntFilters.Keys) -join ' and ') in $HuntMinutes min, a debugger on every OUTLOOK.EXE: $($huntDumps.Count) first- or second-chance dump(s) in $HuntModule's hunt (results\$($facts.Name)\first-chance.log)")
    if ($null -ne $huntOffice) { $lines.Add("office     $($huntOffice['VersionToReport']) $($huntOffice['ProductReleaseIds']) $($huntOffice['Platform']), $($huntOffice['AudienceData']), updates enabled: $($huntOffice['UpdatesEnabled']); $($huntOffice['module'])") }
}
$lines.Add((Format-OutlookStartLine $outlookStart))
if ($null -ne $signatureSettingsState) { $lines.Add("signature  $($signatureSettingsState['scenario']), values $($signatureSettingsState['valueKind']), made before Outlook started: $(if ($signatureSettingsState['set']) { 'SET' } else { 'NOT SET' }) ($($signatureSettingsState['log']))") }
if ($indexRoots) { $lines.Add("index      $indexRoots") }
$lines.Add((Format-GuestClockSummary $guestClockRecords.ToArray()))
$endLine = "guest end  rested on $($checkpointActions['restedOn']); settings staged=$($checkpointActions['settingsStaged']); saved=$($checkpointActions['saved'])"
if ($checkpointActions['green']) { $endLine += "; green checkpoint $($checkpointActions['green'])" }
$lines.Add($endLine)
$lines.Add("time       $(($timings.Keys | ForEach-Object { "$_ $($timings[$_])" }) -join ', ')")
if ($null -ne $trx -and @($trx.Failed).Count -gt 0) {
    $lines.Add(''); $lines.Add('FAILED TESTS')
    foreach ($f in $trx.Failed) { $lines.Add("  $($f.name)"); $lines.Add("      $($f.message)") }
}
if ($dumpList.Count -gt 0) {
    $lines.Add(''); $lines.Add('CRASH DUMPS - the first look (host\Read-CrashDump.ps1); the whole report is beside each dump, <dump>.txt')
    foreach ($dl in $dumpList) {
        $lines.Add("  $($dl['name'])")
        foreach ($h in @($dl['headline'])) { if ($h) { $lines.Add("      $h") } }
    }
}
if ($null -ne $consoleFacts -and @($consoleFacts.ProvedNothing).Count -gt 0) {
    $lines.Add(''); $lines.Add('PROVED NOTHING')
    foreach ($p in $consoleFacts.ProvedNothing) { $lines.Add("  $($p.Substring(0, [Math]::Min(240, $p.Length)))") }
}
$lines.Add(''); $lines.Add("results    $runDir")
$lines.Add('           summary.txt and summary.json here; results\ the TRX, console.txt and VSTest logs; guest\ every guest call; guest-clock.log the guest''s clock around each; run.log the steps')
[System.IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), $lines)
Write-Host ''
$lines | ForEach-Object { Write-Host $_ }
exit $ExitCodes[$verdict.Verdict]
