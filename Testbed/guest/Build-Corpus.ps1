<#
.SYNOPSIS
    plan -> probe -> build -> census, with the committed corpus parameters as defaults.

.DESCRIPTION
    RUN THIS ON THE GUEST, IN SESSION 1. Everything except the plan step drives Outlook over COM,
    and COM does not work in session 0. Reach session 1 with Register-InteractiveTask.ps1:

        .\Register-InteractiveTask.ps1 -Script ".\Build-Corpus.ps1 -Execute"

    Windows PowerShell 5.1 - no ternary, no `??`.

    THE PARAMETERS ARE NOT AN EXAMPLE. corpusId vm2, seed 7777, anchor 2026-08-19, count 20000,
    default shape: those four values reproduce the corpus that every published sweep and frame
    measurement in this repository is a statement about. They are pinned equal here, in
    Testbed/testbed.json and in Docs/corpus-measurement-plan.md, and .github/scripts/check-testbed-references.ps1
    fails the build when the three stop agreeing. Change them only if you mean to build a
    DIFFERENT corpus, and give it a different id when you do.

    THE PROFILE MUST HAVE NO MAIL ACCOUNTS. There is no override. A build creates unsent items in
    bulk; the first real run put 5,532 of them into the target store's Outbox, inert only because
    that profile could not send. On a profile with an account those are 5,532 real messages queued
    for delivery. So corpus work happens in the no-accounts profile and the tier runs in the other
    one, and switching between them is a restart of Outlook.

    WHAT TO CHECK IN THE OUTPUT before letting a build proceed - the script prints these and stops
    on any of them:

      * the store line and the profile line both say accepted, and `profile accounts: 0`
      * the PLACEMENT probe named a verified rung. On this corpus it was DraftsThenMoveWithSentFlag.
      * the DATE probe named a verified rung. On this corpus it was PropertyAccessorDates.

    A build that had to be talked past either probe is a build whose measurements mean something
    other than what they say. Placement is settled before dates on purpose: a date probed against
    an item filed in the wrong folder cannot tell "the date does not select" from "the item is not
    there", and that confusion is exactly what made the first run's verdict worthless.

    COST, from the 2026-08-19 build of these exact parameters: 20,000 items, 225,282,619 body
    bytes, 13m25s, 24.8 items/s. That rate is the WITH-MOVE figure - DraftsThenMoveWithSentFlag
    writes every item twice. The build is resumable and idempotent: it creates the ordinals the
    manifest lacks, so an interrupted run is finished by running it again, and a finished one
    re-run is a no-op.

    AFTERWARDS, TWO THINGS.

    1. COPY THE MANIFEST OFF THE GUEST. It is the only thing that can tear the corpus down -
       teardown deletes by EntryID allowlist AND subject tag, and there is no second route the
       mailbox-safety rules permit. Host-side: Testbed/host/Copy-FromGuest.ps1.
    2. EXPECT THE OUTBOX TO BE EMPTY. The 2026-08-19 build left 2,761 items there, which is
       EXACTLY the plan's unread count - the MSGFLAG_SUBMIT defect, since fixed. The count is
       predictable in advance, so a non-empty Outbox after a rebuild is a specific signal that
       the fix did not take, not a mystery.

.PARAMETER Store
    Store display name. `Outlook Data File` is what the measured corpus lives in; the
    three-store layout in Docs/live-tier-on-the-vm.md calls it `Corpus A`.

.PARAMETER Execute
    Without it: plan and a dry-run build, and NEITHER probe runs (both create items).

.EXAMPLE
    .\Build-Corpus.ps1
    .\Build-Corpus.ps1 -Execute
    .\Build-Corpus.ps1 -Store "Corpus B" -CorpusId vm3 -Execute
#>
[CmdletBinding()]
param(
    [string] $ToolsExe = 'C:\OutlookAI-Q5\tools\OutlookAI.RemediationTools.exe',
    [string] $Store = 'Outlook Data File',

    # --- pinned corpus parameters; see check-testbed-references.ps1 --------------------------
    [string] $CorpusId = 'vm2',
    [long]   $Seed = 7777,
    [string] $Anchor = '2026-08-19',
    [int]    $Count = 20000,
    # ----------------------------------------------------------------------------------------

    [string] $Manifest = 'C:\OutlookAI-Q5\corpus-vm2.jsonl',
    [string] $LogPath = 'C:\OutlookAI-Q5\corpus-build.log',
    [int]    $ProgressEvery = 250,
    [switch] $SkipProbe,
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ToolsExe)) {
    throw @"
Tools not found at $ToolsExe
The guest has no .NET SDK, so nothing can be built here. Publish on the host and copy in:
    pwsh -File Testbed/host/Publish-GuestPayload.ps1
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -Path .work\testbed-payload\Tools.zip -Destination C:\OutlookAI-Q5\Tools.zip
then on the guest: Expand-Archive C:\OutlookAI-Q5\Tools.zip -DestinationPath C:\OutlookAI-Q5\tools -Force
"@
}

$identity = @('--corpus-id', $CorpusId, '--seed', "$Seed", '--anchor', $Anchor, '--count', "$Count")
$target = @('--store', $Store, '--allow-store', $Store)

function Invoke-Corpus {
    param([string] $Verb, [string[]] $Arguments, [switch] $Fatal)

    $line = "$Verb $($Arguments -join ' ')"
    Write-Host ''
    Write-Host "=== $line"
    Add-Content -LiteralPath $LogPath -Value "`r`n=== $line"

    # Output is captured and echoed rather than streamed, so a child that outlives the call
    # cannot hold the pipe open. Every number these verbs print is meant to be kept anyway.
    $output = & $ToolsExe $Verb @Arguments 2>&1
    $code = $LASTEXITCODE
    $text = ($output | Out-String)
    Write-Host $text
    Add-Content -LiteralPath $LogPath -Value $text
    Add-Content -LiteralPath $LogPath -Value "exit $code"

    if ($code -ne 0 -and $Fatal) { throw "$Verb failed with exit $code. Log: $LogPath" }
    return $code
}

Remove-Item -LiteralPath $LogPath -Force -ErrorAction SilentlyContinue

# --- 0. the expectation sheet. Pure: no Outlook, runnable anywhere, including the host. -------
# Save this. Every measurement taken later is a ratio against one of these numbers, and computing
# them afterwards from the store is both slower and less trustworthy than reading them off the
# plan that produced it.
Invoke-Corpus -Verb 'corpus-plan' -Arguments $identity -Fatal | Out-Null

if (-not $Execute) {
    Write-Host ''
    Write-Host 'Dry run. The probes are NOT run - both create items - and nothing is written.'
    Invoke-Corpus -Verb 'corpus-build' -Arguments ($target + $identity + @('--manifest', $Manifest)) | Out-Null
    Write-Host ''
    Write-Host 'Re-run with -Execute to probe and build. Read the plan above first: if the 7-day and'
    Write-Host '60-day window counts are close together, change the date bands BEFORE building, not after.'
    return
}

# --- 1. placement, then dates. Cheap; creates and deletes a handful of throwaway items. -------
if (-not $SkipProbe) {
    $probe = Invoke-Corpus -Verb 'corpus-probe' -Arguments ($target + $identity + @('--execute'))
    if ($probe -ne 0) {
        throw @"
The probe did not verify a rung (exit $probe). STOP HERE.
  * placement NOT ACHIEVABLE -> the corpus would be invisible to the freshness sweep, because
    Outlook files unsent items in Drafts and the sweep does not cover Drafts. Measurements taken
    against it would be measurements of an empty store. Docs/corpus-measurement-plan.md,
    'If placement fails', is the fallback route.
  * dates NOT ACHIEVABLE -> every step that mentions a window is void. See 'If the dates do not
    stick' in the same document.
Do not pass --allow-drafts-placement or --allow-undated to get past this without reading what
each one costs; the tool prints it.
"@
    }
}

# --- 2. build. Resumable and idempotent. ------------------------------------------------------
Invoke-Corpus -Verb 'corpus-build' `
    -Arguments ($target + $identity + @('--manifest', $Manifest, '--progress-every', "$ProgressEvery", '--execute')) `
    -Fatal | Out-Null

# --- 3. what actually landed. Read-only; the build runs this on itself, and it is repeated here
#        so an operator sees it as its own line rather than buried in build output. -------------
Invoke-Corpus -Verb 'corpus-census' -Arguments ($target + $identity + @('--manifest', $Manifest)) | Out-Null

Write-Host ''
Write-Host "Log: $LogPath"
Write-Host "Manifest: $Manifest"
Write-Host 'NOW COPY THE MANIFEST OFF THE GUEST. Without it the corpus cannot be torn down:'
Write-Host '    pwsh -File Testbed/host/Copy-FromGuest.ps1'
Write-Host 'Then take a checkpoint, and let Windows Search settle before any index measurement.'
