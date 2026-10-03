#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the live tier on a test guest end to end in ONE call, and writes summary.txt and
    summary.json: lease, checkpoint, staging, hub rebuild, throwaway reset, the suite at RunLevel
    Limited with the per-run opt-in, results, resting checkpoint, save, release.

.DESCRIPTION
    The guest counterpart of Invoke-TestsOnBuildVm.ps1, written 2026-10-03 after the first guest
    runs were driven by hand from scratch scripts, hundreds of calls a run. One run is:

      1. BUILD   - the commit, never the working tree: a detached `git worktree` of it, from which
                   Publish-GuestPayload.ps1 builds the server and the tools and
                   Publish-LiveTierPayload.ps1 takes the suite's source (`git archive`) and the
                   offline feed; New-LiveTestSettings.ps1 renders the guest's settings from that
                   commit's Testbed/testbed.json. Uncommitted changes are not in the run, and it says
                   how many it left out. Nothing touches the guest yet.
      2. LEASE   - waits until nobody else holds a live lease on the guest (Set-TestbedLease.ps1),
                   NEVER taking one over, then takes it, and renews it before every long step. The
                   guests are shared between agents, and the idle-saver saves any running guest that
                   holds no live lease.
      3. STAGE   - restores the start checkpoint, copies the payloads, the settings and the commit's
                   Testbed/guest scripts in, swaps the server, tools, source and feed, and proves the
                   suite TEST-READY (Install-DotnetSdk.ps1 -Execute, then -Verify).
      4. PREPARE - Restart-Guest.ps1 -Execute (graceful, or it refuses), then Reset-HubPopulation.ps1
                   and Reset-ThrowawayStore.ps1 -Execute (Testbed/README.md steps 9a and 9a-ii), both
                   at RunLevel Limited: every Outlook start NOT elevated (D77).
      5. RUN     - the suite through Register-InteractiveTask.ps1 -RunLevel Limited, with
                   OUTLOOKAI_LIVE_OPT_IN set to the guest's computer name INSIDE the task's script,
                   for this run only, and the guest's derived filter (T2/LiveRunFilters.cs).
      6. FETCH   - the TRX file, the console capture and VSTest's logs, and the guest's Application
                   log entries for an OUTLOOK.EXE crash during the run.
      7. REST    - optionally a checkpoint of a GREEN run (-GreenCheckpoint), then the resting
                   checkpoint restored, the settings staged on it, the guest SAVED and the lease
                   released - on every path, failures included, once the guest was touched.
      8. REPORT  - summary.txt and summary.json in <RepoPath>\.work\guest-live-runs\<run id>\.

    WHAT IT ENFORCES, so a caller cannot forget it (AGENTS.md "Mailbox Safety", "Test VMs stay
    saved unless in use"):
      * Only the two test guests. The build VM, the retired OutlookAI-TestVM and anything else are
        refused, and nothing runs on this workstation: the suite runs inside the guest.
      * The opt-in exists only inside the guest task of this run; it is never set here.
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

    EXIT CODES: 0 PASS; 1 FAIL - a test, a safety verdict or Outlook failed, or the filter selected
    nothing; 2 BUILD - the commit did not build, stage or become TEST-READY; 3 INFRA - not tested
    (lease, VM, guest calls, time); 4 REFUSED - the request itself.

    Run it in the background or with a timeout of an hour or more: a full run of the indexed guest
    takes about 40 minutes. A caller stopped part-way leaves the guest leased until the lease
    expires; the next run restores its checkpoint first anyway.

.PARAMETER VMName
    OutlookAI-Indexed or OutlookAI-Unindexed. Mandatory: no default picks a guest.

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

.PARAMETER SkipThrowawayReset
    Skips step 9a-ii. For a narrowed run only.

.PARAMETER SelfTest
    Checks the pure parts - the guest table, the filter, the lease decision, the TRX and console
    readers, the verdict table, summary.json's shape - and the constants against the files they
    copy. No Hyper-V, no git, no guest. Runs on the build VM with every other -SelfTest.

.EXAMPLE
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Unindexed
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Indexed d4e31fe -GreenCheckpoint CP-18C-LIVE-GREEN
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Indexed -FilterSuffix '&FullyQualifiedName~LiveShowMeTests' -SkipHubReset -SkipThrowawayReset
    pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -SelfTest
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')] [string] $VMName,
    [Parameter(ParameterSetName = 'Run', Position = 0)] [string] $Ref = 'HEAD',
    [Parameter(ParameterSetName = 'Run')] [string] $RepoPath,
    [Parameter(ParameterSetName = 'Run')] [string] $Filter = '',
    [Parameter(ParameterSetName = 'Run')] [string] $FilterSuffix = '',
    [Parameter(ParameterSetName = 'Run')] [string] $Checkpoint = '',
    [Parameter(ParameterSetName = 'Run')] [string] $RestingCheckpoint = '',
    [Parameter(ParameterSetName = 'Run')] [string] $GreenCheckpoint = '',
    [Parameter(ParameterSetName = 'Run')] [switch] $RestOnGreen,
    [Parameter(ParameterSetName = 'Run')] [switch] $SkipHubReset,
    [Parameter(ParameterSetName = 'Run')] [switch] $SkipThrowawayReset,
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
    'OutlookAI-Indexed'   = [ordered]@{ ComputerName = 'OAI-INDEXED'; Indexed = $true; Checkpoint = 'CP-17C-CORPUS-160K' }
    'OutlookAI-Unindexed' = [ordered]@{ ComputerName = 'OAI-UNINDEXED'; Indexed = $false; Checkpoint = 'CP-12B-POPULATIONS-V2' }
}
$GuestRoot = 'C:\OutlookAI-Q5'
$GuestPayloadDir = 'C:\OutlookAI-Q5\g4-payload'
$GuestRunsRoot = 'C:\OutlookAI-Q5\live-runs'
$GuestSettingsPath = 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json'
$SdkInstallerName = 'dotnet-sdk-10.0.401-win-x64.exe'
$SdkSha512 = 'f0d8f8e7ec24efb05172a65dd80c4a9b1ef17efcebdbf0f57c15f436eea417960a7eeb6726c473a042373d6a8b94ac1adc7d680decbf7d2c45fa5c5662d62265'
$FilterBase = 'Category=Live&Requires!=DelegateStore&Requires!=CachedExchange'
$FilterUnindexedExtra = '&Requires!=SearchIndex'
$LeaseReasonPrefix = 'Invoke-LiveTierOnGuest'
$MinFreeGbForCheckpoint = 60
$ExitCodes = [ordered]@{ PASS = 0; FAIL = 1; BUILD = 2; INFRA = 3; REFUSED = 4 }

# A NATIVE PROGRAM RUNS THROUGH HERE (Q78) - see Invoke-TestsOnBuildVm.ps1, which says why:
# 'Continue' in this scope only, stderr lines as strings, the exit code left in $LASTEXITCODE.
function Invoke-NativeCommand {
    param([Parameter(Mandatory = $true)] [scriptblock] $NativeCommand)

    $ErrorActionPreference = 'Continue'
    try {
        & $NativeCommand | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { $_ }
        }
    }
    catch {
        throw
    }
}

# =============================================================================================
# PURE DECISIONS. No Hyper-V, no git, no guest. -SelfTest pins them.
# =============================================================================================

# The guest's facts, or $null for anything that is not one of the two test guests - compared whole
# and ignoring case, as Hyper-V compares names.
function Get-GuestFacts {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -ne $Name.Trim()) { return $null }
    foreach ($key in $Guests.Keys) {
        if ([string]::Equals($Name, $key, [System.StringComparison]::OrdinalIgnoreCase)) {
            $facts = $Guests[$key]
            return [pscustomobject]@{ Name = $key; ComputerName = $facts.ComputerName; Indexed = $facts.Indexed; Checkpoint = $facts.Checkpoint }
        }
    }
    return $null
}

# The run's filter: the guest's derived one (LiveRunFilters.Guest / GuestUnindexed) narrowed by
# -FilterSuffix, or -Filter whole. Returns the refusal instead when the request is malformed. The
# filter is written into the guest task's script inside double quotes, so a character that would
# end or expand that string is refused rather than escaped.
function Get-LiveFilter {
    param([bool] $Indexed, [string] $Override, [string] $Suffix)
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
    return [pscustomobject]@{ Filter = $filter + $Suffix; Narrowed = [bool]$Suffix; Refusal = $null }
}

# A lease as Set-TestbedLease.ps1 writes it, judged for THIS run: 'free' (none, expired or
# unreadable - TestbedLeasePath.ps1's rule), 'ours', or 'theirs' - which is never taken over.
function Get-LeaseDecision {
    param([object] $Lease, [string] $RunId, [long] $NowUnix)
    if ($null -eq $Lease) { return 'free' }
    $expires = 0L
    try { $expires = [long]$Lease.expiresUnix } catch { return 'free' }
    if ($expires -le $NowUnix) { return 'free' }
    if (([string]$Lease.reason) -ceq "$LeaseReasonPrefix $RunId") { return 'ours' }
    return 'theirs'
}

# A run directory's name: when it started, which guest, which commit.
function New-RunId {
    param([DateTime] $When, [string] $Sha, [string] $GuestName)
    if ($Sha -notmatch '^[0-9a-f]{12,40}$') { throw "Not a commit id: '$Sha'." }
    $short = 'indexed'
    if ($GuestName -match 'Unindexed$') { $short = 'unindexed' }
    return ('{0:yyyyMMdd-HHmmss}-{1}-{2}' -f $When, $short, $Sha.Substring(0, 12))
}

# What a TRX file says - Invoke-TestsOnBuildVm.ps1's reader: outcomes counted from the results,
# NotExecuted as xUnit's skip, a failure's first message line.
function ConvertFrom-TrxText {
    param([string] $Text)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Text)
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('t', 'http://microsoft.com/schemas/VisualStudio/TeamTest/2010')
    $results = @($doc.SelectNodes('//t:Results/t:UnitTestResult', $ns))
    $counts = [ordered]@{ total = $results.Count; passed = 0; failed = 0; skipped = 0; other = 0 }
    $failed = @()
    $skipped = @()
    foreach ($r in $results) {
        $outcome = $r.GetAttribute('outcome')
        if ($outcome -eq 'Passed') { $counts['passed']++ }
        elseif ($outcome -eq 'Failed') {
            $counts['failed']++
            $message = ''
            $node = $r.SelectSingleNode('t:Output/t:ErrorInfo/t:Message', $ns)
            if ($null -ne $node) { $message = (($node.InnerText -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1) }
            $failed += [pscustomobject]@{ name = $r.GetAttribute('testName'); message = [string]$message }
        }
        elseif ($outcome -eq 'NotExecuted') { $counts['skipped']++; $skipped += $r.GetAttribute('testName') }
        else { $counts['other']++; $failed += [pscustomobject]@{ name = $r.GetAttribute('testName'); message = "outcome $outcome" } }
    }
    return [pscustomobject]@{ Counts = $counts; Failed = $failed; Skipped = $skipped }
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
        [string] $Results
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
        checkpoints    = $Checkpoints
        timings        = $Timings
        results        = $Results
    }
}

function Format-Seconds {
    param([double] $Seconds)
    if ($Seconds -lt 60) { return ('{0:N0} s' -f $Seconds) }
    $whole = [int][Math]::Floor($Seconds)
    return ('{0} m {1:00} s' -f [int][Math]::Floor($whole / 60), ($whole % 60))
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
    Check 'the indexed guest is a test guest' 'OAI-INDEXED|True|CP-17C-CORPUS-160K' @($i.ComputerName, $i.Indexed, $i.Checkpoint)
    $u = Get-GuestFacts 'outlookai-unindexed'
    Check 'the unindexed guest is one, in any case' 'OutlookAI-Unindexed|OAI-UNINDEXED|False' @($u.Name, $u.ComputerName, $u.Indexed)
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
    Check 'no lease is free' 'free' (Get-LeaseDecision $null 'r1' $now)
    Check 'an expired lease is free' 'free' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now - 1; reason = 'someone' }) 'r1' $now)
    Check 'an unreadable expiry is free' 'free' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = 'soon'; reason = 'someone' }) 'r1' $now)
    Check 'a live lease of this run is ours' 'ours' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now + 60; reason = 'Invoke-LiveTierOnGuest r1' }) 'r1' $now)
    Check 'a live lease of another run is theirs' 'theirs' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now + 60; reason = 'Invoke-LiveTierOnGuest r2' }) 'r1' $now)
    Check 'a live lease of another agent is theirs' 'theirs' (Get-LeaseDecision ([pscustomobject]@{ expiresUnix = $now + 60; reason = 'agent accec8f1: folder-path tests' }) 'r1' $now)

    Write-Host '== the run id =='
    Check 'when, guest, commit' '20261003-174500-indexed-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' 'OutlookAI-Indexed')
    Check 'the unindexed guest' '20261003-174500-unindexed-d4e31fe12345' (New-RunId ([datetime]'2026-10-03T17:45:00') 'd4e31fe1234567890' 'OutlookAI-Unindexed')
    $threw = $false; try { New-RunId (Get-Date) 'HEAD' 'OutlookAI-Indexed' | Out-Null } catch { $threw = $true }
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

    Write-Host '== the constants, against the files they copy =='
    $media = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\MEDIA.md'))
    Check 'MEDIA.md names the SDK installer' $true $media.Contains($SdkInstallerName)
    Check 'MEDIA.md records its SHA-512' $true $media.ToUpperInvariant().Contains($SdkSha512.ToUpperInvariant())
    $filters = [System.IO.File]::ReadAllText((Join-Path $repo 'McpServer\OutlookAI.McpServer.Tests\T2\LiveRunFilters.cs'))
    Check 'LiveRunFilters.Guest leaves out exactly DelegateStore and CachedExchange' $true $filters.Contains('WorkstationOnlyCapabilities { get; } = new[] { DelegateStore, CachedExchange }')
    Check 'LiveRunFilters.GuestUnindexed adds SearchIndex' $true $filters.Contains('GuestUnindexed { get; } = Guest + "&Requires!=" + SearchIndex')
    foreach ($name in 'Reset-HubPopulation.ps1', 'Reset-ThrowawayStore.ps1') {
        $text = [System.IO.File]::ReadAllText((Join-Path $repo "Testbed\guest\$name"))
        Check "$name reads the settings where this stages them" $true $text.Contains("`$SettingsPath = '$GuestSettingsPath'")
    }
    $sdk = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\guest\Install-DotnetSdk.ps1'))
    Check 'Install-DotnetSdk.ps1 builds the source this stages' $true $sdk.Contains("'$GuestRoot\src'")
    $testbed = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\testbed.json')) | ConvertFrom-Json
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
$script:LogPath = $null
function Say([string] $Message) {
    $line = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $Message
    Write-Host $line
    if ($script:LogPath) { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 }
}
function Stop-Refused([string] $Why) {
    Write-Host "REFUSED: $Why"
    exit $ExitCodes['REFUSED']
}

$facts = Get-GuestFacts $VMName
if ($null -eq $facts) { Stop-Refused "'$VMName' is not a test guest - only OutlookAI-Indexed and OutlookAI-Unindexed. The build VM has Invoke-TestsOnBuildVm.ps1, and nothing here runs on this workstation." }
$filterDecision = Get-LiveFilter $facts.Indexed $Filter $FilterSuffix
if ($filterDecision.Refusal) { Stop-Refused $filterDecision.Refusal }
$runFilter = $filterDecision.Filter
$narrowed = $filterDecision.Narrowed
if (-not $Checkpoint) { $Checkpoint = $facts.Checkpoint }
if (-not $RestingCheckpoint) { $RestingCheckpoint = $facts.Checkpoint }
if ($RestOnGreen -and -not $GreenCheckpoint) { Stop-Refused '-RestOnGreen needs -GreenCheckpoint.' }
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
if (-not (Test-Path -LiteralPath $SdkInstallerPath)) { Stop-Refused "no SDK installer at $SdkInstallerPath (Testbed/MEDIA.md); pass -SdkInstallerPath." }
if (-not (Get-VM -Name $facts.Name -ErrorAction SilentlyContinue)) { Stop-Refused "Hyper-V has no VM named '$($facts.Name)' here." }
foreach ($cp in @($Checkpoint, $RestingCheckpoint)) {
    if (-not (Get-VMSnapshot -VMName $facts.Name -Name $cp -ErrorAction SilentlyContinue)) { Stop-Refused "$($facts.Name) has no checkpoint named '$cp'." }
}
if ($GreenCheckpoint -and (Get-VMSnapshot -VMName $facts.Name -Name $GreenCheckpoint -ErrorAction SilentlyContinue)) { Stop-Refused "$($facts.Name) already has a checkpoint named '$GreenCheckpoint'." }

$sha = (@(Invoke-NativeCommand { & git -C $repo rev-parse --verify "$Ref^{commit}" 2>&1 }) | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { Stop-Refused "git could not resolve '$Ref' to a commit in $repo - $sha" }
$subject = (@(Invoke-NativeCommand { & git -C $repo log -1 --format=%s $sha 2>&1 }) | Out-String).Trim()
$dirtyCount = 0
if ($Ref -eq 'HEAD') { $dirtyCount = @(Invoke-NativeCommand { & git -C $repo status --porcelain 2>&1 } | Where-Object { $_ }).Count }
$runId = New-RunId $started $sha $facts.Name
if (-not $ResultsRoot) { $ResultsRoot = Join-Path $repo '.work\guest-live-runs' }
$runDir = Join-Path $ResultsRoot $runId
New-Item -ItemType Directory -Force -Path $runDir, (Join-Path $runDir 'guest') | Out-Null
$script:LogPath = Join-Path $runDir 'run.log'
$revision = [ordered]@{ ref = $Ref; sha = $sha; subject = $subject; repo = $repo; uncommittedLeftOut = $dirtyCount }

$kind = 'unindexed'
if ($facts.Indexed) { $kind = 'indexed' }
Say "== Invoke-LiveTierOnGuest: run $runId =="
Say "  guest     $($facts.Name) ($($facts.ComputerName), $kind), from $Checkpoint, resting on $RestingCheckpoint"
Say "  revision  $($sha.Substring(0, 12)) - $subject"
if ($dirtyCount -gt 0) { Say "  NOTE      $dirtyCount uncommitted change(s) in $repo are NOT in this run - it tests the commit. Commit first to test them." }
Say "  filter    $runFilter"
Say "  results   $runDir"

$timings = [ordered]@{}
$checkpointActions = [ordered]@{ start = $Checkpoint; resting = $RestingCheckpoint; green = $null; restedOn = $null; settingsStaged = $false; saved = $false }
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
    $decision = Get-LeaseDecision (Get-TestbedLease -VMName $facts.Name) $runId ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
    if ($decision -eq 'theirs') { throw "the lease on $($facts.Name) is held by someone else - not touching the guest" }
    & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $facts.Name -Minutes ([Math]::Min(480, [Math]::Max(30, $Minutes))) -Reason "$LeaseReasonPrefix $runId" | Out-Null
}
function Wait-Heartbeat {
    $deadline = (Get-Date).AddMinutes(5)
    do {
        $hb = (Get-VMIntegrationService -VMName $facts.Name -Name 'Heartbeat').PrimaryStatusDescription
        if ($hb -eq 'OK') { return }
        Start-Sleep -Seconds 3
    } while ((Get-Date) -lt $deadline)
    throw "the guest's heartbeat did not come back within 5 minutes"
}
function Restore-Guest([string] $Name) {
    Restore-VMSnapshot -VMName $facts.Name -Name $Name -Confirm:$false
    if ((Get-VM -Name $facts.Name).State -ne 'Running') { Start-VM -Name $facts.Name }
    Wait-Heartbeat
}
# One piece of work on the guest; its whole output goes to guest\<name>.log. Session 1 through
# Register-InteractiveTask.ps1 at RunLevel Limited (where Outlook runs, NOT elevated), or session 0
# over PowerShell Direct for work that needs no desktop.
function Invoke-Guest([string] $Name, [string] $Script, [int] $TimeoutSeconds, [switch] $Session0) {
    Use-Lease ([int][Math]::Ceiling($TimeoutSeconds / 60) + 30)
    if ($Session0) {
        $out = Invoke-Command -VMName $facts.Name -Credential $credential -ScriptBlock {
            param($text)
            $ErrorActionPreference = 'Continue'
            Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
            try { & ([scriptblock]::Create($text)) *>&1 | Out-String -Width 400 } catch { "GUEST-CALL-THREW: $($_.Exception.Message)" }
        } -ArgumentList $Script 2>&1 | Out-String -Width 400
    }
    else {
        $out = Invoke-Command -VMName $facts.Name -Credential $credential -ScriptBlock {
            param($text, $timeout)
            $ErrorActionPreference = 'Continue'
            Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
            try { & 'C:\OutlookAI-Q5\Register-InteractiveTask.ps1' -Script $text -TimeoutSeconds $timeout -RunLevel Limited *>&1 | Out-String -Width 400 } catch { "GUEST-CALL-THREW: $($_.Exception.Message)" }
            "RUN-REGISTER-EXIT: $LASTEXITCODE"
        } -ArgumentList $Script, $TimeoutSeconds 2>&1 | Out-String -Width 400
    }
    [System.IO.File]::WriteAllText((Join-Path $runDir "guest\$Name.log"), $out)
    return $out
}
function Get-TaskExit([string] $Text) {
    $m = [regex]::Match($Text, 'RUN-REGISTER-EXIT: (-?\d+)')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return $null
}

try {
    # ---- BUILD: the commit's tree, its payloads and the guest's settings. The guest is untouched.
    $t0 = Get-Date
    try {
        Invoke-NativeCommand { & git -C $repo worktree add --detach $tree $sha 2>&1 } | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "git worktree add failed ($LASTEXITCODE)" }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Publish-GuestPayload.ps1') -RepoRoot $tree -OutDir (Join-Path $runDir 'payload') *> (Join-Path $runDir 'publish-guest.log')
        if ($LASTEXITCODE -ne 0) { throw "Publish-GuestPayload.ps1 failed ($LASTEXITCODE) - publish-guest.log" }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Publish-LiveTierPayload.ps1') -RepoRoot $tree -OutDir (Join-Path $runDir 'livetier') -Ref HEAD -SdkInstallerPath $SdkInstallerPath -ExpectedSha512 $SdkSha512 *> (Join-Path $runDir 'publish-livetier.log')
        if ($LASTEXITCODE -ne 0) { throw "Publish-LiveTierPayload.ps1 failed ($LASTEXITCODE) - publish-livetier.log" }
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'New-LiveTestSettings.ps1') -VMName $facts.Name -RepoRoot $tree -OutPath (Join-Path $runDir 'live-test-settings.json') *> (Join-Path $runDir 'settings.log')
        if ($LASTEXITCODE -ne 0) { throw "New-LiveTestSettings.ps1 refused this guest's settings ($LASTEXITCODE) - settings.log" }
    }
    catch { $stage = 'BUILD'; $stageWhy = $_.Exception.Message; throw }
    $timings['build'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "built: payloads and settings ($($timings['build']))"

    # ---- LEASE: wait, never take over.
    $t0 = Get-Date
    $deadline = (Get-Date).AddMinutes($LeaseWaitMinutes)
    $said = $false
    while ($true) {
        $lease = Get-TestbedLease -VMName $facts.Name
        if ((Get-LeaseDecision $lease $runId ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())) -ne 'theirs') { break }
        if (-not $said) { Say "waiting: $($facts.Name) is leased until $($lease.expiresUtc) - $($lease.reason)"; $said = $true }
        if ((Get-Date) -ge $deadline) { $stage = 'INFRA'; $stageWhy = "the guest stayed leased by someone else for $LeaseWaitMinutes minutes"; throw $stageWhy }
        Start-Sleep -Seconds 60
    }
    Use-Lease 60
    $timings['lease'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "lease taken ($($timings['lease']))"
    $credential = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $CredentialRepoRoot -VMName $facts.Name

    # ---- STAGE.
    $t0 = Get-Date
    try {
        $touched = $true
        Restore-Guest $Checkpoint
        Say "restored $Checkpoint"
        $files = @((Join-Path $runDir 'payload\McpServer.zip'), (Join-Path $runDir 'payload\Tools.zip'), (Join-Path $runDir 'livetier\Source.zip'), (Join-Path $runDir 'livetier\NuGet.zip'), (Join-Path $runDir 'live-test-settings.json'))
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Copy-ToGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -Path $files -Destination $GuestPayloadDir *> (Join-Path $runDir 'copy.log')
        $guestScripts = @(Get-ChildItem -LiteralPath (Join-Path $tree 'Testbed\guest') -File | ForEach-Object { $_.FullName })
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Copy-ToGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -Path $guestScripts -Destination $GuestRoot *>> (Join-Path $runDir 'copy.log')
        $swap = @"
`$ErrorActionPreference = 'Stop'
if (@(Get-Process -Name OUTLOOK,OutlookAI.RemediationTools,OutlookAI.McpServer,OutlookAI.ComHost,testhost,dotnet -ErrorAction SilentlyContinue).Count -gt 0) { throw 'Outlook, the tools, the server or dotnet is running - not replacing anything' }
`$stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
foreach (`$name in 'server', 'tools', 'src', 'nuget-offline') { `$from = Join-Path '$GuestRoot' `$name; if (Test-Path -LiteralPath `$from) { Rename-Item -LiteralPath `$from -NewName "`$name.pre-`$stamp" } }
Expand-Archive -LiteralPath '$GuestPayloadDir\McpServer.zip' -DestinationPath '$GuestRoot\server' -Force
Expand-Archive -LiteralPath '$GuestPayloadDir\Tools.zip' -DestinationPath '$GuestRoot\tools' -Force
Expand-Archive -LiteralPath '$GuestPayloadDir\Source.zip' -DestinationPath '$GuestRoot\src' -Force
Expand-Archive -LiteralPath '$GuestPayloadDir\NuGet.zip' -DestinationPath '$GuestRoot\nuget-offline' -Force
New-Item -ItemType Directory -Force -Path (Split-Path -Parent '$GuestSettingsPath') | Out-Null
Copy-Item -LiteralPath '$GuestPayloadDir\live-test-settings.json' -Destination '$GuestSettingsPath' -Force
"SWAP-DONE settings `$((Get-FileHash '$GuestSettingsPath').Hash)"
"@
        $o = Invoke-Guest 'swap' $swap 900 -Session0
        if ($o -notmatch 'SWAP-DONE') { throw 'the payload swap did not finish - guest\swap.log' }
        $o = Invoke-Guest 'sdk-execute' "Set-Location '$GuestRoot'; & .\Install-DotnetSdk.ps1 -ExpectedSha512 '$SdkSha512' -Execute; `"SDK-EXECUTE-EXIT `$LASTEXITCODE`"" 1800 -Session0
        if ($o -notmatch 'SDK-EXECUTE-EXIT 0') { throw 'Install-DotnetSdk.ps1 -Execute failed - guest\sdk-execute.log' }
        $o = Invoke-Guest 'sdk-verify' "Set-Location '$GuestRoot'; & .\Install-DotnetSdk.ps1 -Verify; `"SDK-VERIFY-EXIT `$LASTEXITCODE`"" 1800 -Session0
        if ($o -notmatch 'VERDICT: TEST-READY') { $stage = 'BUILD'; $stageWhy = 'the suite is not TEST-READY on the guest - it did not build or enumerate there (guest\sdk-verify.log)'; throw $stageWhy }
    }
    catch { if (-not $stage) { $stage = 'INFRA'; $stageWhy = "staging failed: $($_.Exception.Message)" }; throw }
    $timings['stage'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    Say "staged: TEST-READY ($($timings['stage']))"

    # ---- PREPARE: a graceful restart, the hub rebuild and the throwaway data file.
    $t0 = Get-Date
    try {
        Use-Lease 30
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Restart-Guest.ps1') -VMName $facts.Name -Execute -CancelLogonPrompt -RepoRoot $CredentialRepoRoot -LogPath (Join-Path $runDir 'restart.log') *> (Join-Path $runDir 'restart.out.log')
        if ($LASTEXITCODE -ne 0) { throw "Restart-Guest.ps1 did not complete ($LASTEXITCODE) - restart.log" }
        if (-not $SkipHubReset) {
            $o = Invoke-Guest 'reset-hub' "& '$GuestRoot\Reset-HubPopulation.ps1' -Execute; exit `$LASTEXITCODE" 3600
            if ((Get-TaskExit $o) -ne 0) { throw 'the hub rebuild did not succeed - guest\reset-hub.log' }
        }
        if (-not $SkipThrowawayReset) {
            $o = Invoke-Guest 'reset-throwaway' "& '$GuestRoot\Reset-ThrowawayStore.ps1' -Execute; exit `$LASTEXITCODE" 900
            if ((Get-TaskExit $o) -ne 0) { throw 'the throwaway data file was not recreated - guest\reset-throwaway.log' }
        }
    }
    catch { if (-not $stage) { $stage = 'INFRA'; $stageWhy = "preparing the guest failed: $($_.Exception.Message)" }; throw }
    $timings['prepare'] = Format-Seconds ((Get-Date) - $t0).TotalSeconds
    $prepared = @('restart')
    if (-not $SkipHubReset) { $prepared += 'hub rebuild' }
    if (-not $SkipThrowawayReset) { $prepared += 'throwaway reset' }
    Say "prepared: $($prepared -join ', ') ($($timings['prepare']))"

    # ---- RUN: the suite, NOT elevated, opted in for this run only.
    $t0 = Get-Date
    $liveScript = @"
`$env:OUTLOOKAI_LIVE_OPT_IN = '$($facts.ComputerName)'
`$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'; `$env:DOTNET_NOLOGO = '1'; `$env:MSBUILDDISABLENODEREUSE = '1'
New-Item -ItemType Directory -Force -Path '$guestRunDir', '$guestRunDir\diag' | Out-Null
Set-Location '$GuestRoot\src'
& dotnet test McpServer\OutlookAI.McpServer.Tests\OutlookAI.McpServer.Tests.csproj -c Release --filter "$runFilter" --logger "trx;LogFileName=live.trx" --logger "console;verbosity=detailed" --results-directory '$guestRunDir' --diag '$guestRunDir\diag\vstest.log' *> '$guestRunDir\console.txt'
exit `$LASTEXITCODE
"@
    $runStartUtc = [DateTime]::UtcNow
    $o = Invoke-Guest 'live' $liveScript ($RunTimeoutMinutes * 60)
    $suiteExit = Get-TaskExit $o
    $suiteSeconds = ((Get-Date) - $t0).TotalSeconds
    $timings['suite'] = Format-Seconds $suiteSeconds
    Say "suite finished, exit $suiteExit ($($timings['suite']))"

    # ---- FETCH: results, console, logs, crash events.
    $resultsDir = Join-Path $runDir 'results'
    # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
    & (Join-Path $PSScriptRoot 'Copy-FromGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -GuestPath $guestRunDir -Include '*.trx', 'console.txt' -Destination $resultsDir -Force *> (Join-Path $runDir 'fetch.log')
    # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
    & (Join-Path $PSScriptRoot 'Copy-FromGuest.ps1') -VMName $facts.Name -RepoRoot $CredentialRepoRoot -GuestPath "$guestRunDir\diag" -Include '*.log' -Destination (Join-Path $resultsDir 'diag') -Force *>> (Join-Path $runDir 'fetch.log')
    $trxFile = Get-ChildItem -LiteralPath $resultsDir -Filter '*.trx' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($trxFile) { $trx = ConvertFrom-TrxText ([System.IO.File]::ReadAllText($trxFile.FullName)) }
    $consoleFile = Get-ChildItem -LiteralPath $resultsDir -Filter 'console.txt' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($consoleFile) { $consoleFacts = Read-ConsoleFacts ([System.IO.File]::ReadAllText($consoleFile.FullName)) }
    $since = $runStartUtc.AddMinutes(-1).ToString('o')
    $o = Invoke-Guest 'crash-events' "`$s = ([datetime]'$since').ToLocalTime(); `$e = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = `$s; Id = 1000 } -ErrorAction SilentlyContinue | Where-Object { `$_.Message -match 'OUTLOOK\.EXE' }); `"CRASHES `$(`$e.Count)`"; `$e | ForEach-Object { `$_.TimeCreated.ToString('o') + ' ' + ((`$_.Message -split [char]10 | Select-Object -First 4) -join ' | ') }" 300 -Session0
    $cm = [regex]::Match($o, 'CRASHES (\d+)')
    if ($cm.Success) { $crashes = [int]$cm.Groups[1].Value }
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
            $restOn = $RestingCheckpoint
            if ($RestOnGreen -and $checkpointActions['green']) { $restOn = $GreenCheckpoint }
            try {
                Restore-Guest $restOn
                $checkpointActions['restedOn'] = $restOn
            }
            catch { Say "RESTORING $restOn FAILED: $($_.Exception.Message)" }
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
                $checkpointActions['saved'] = ((Get-VM -Name $facts.Name).State -eq 'Saved')
            }
            catch { Say "SAVING THE GUEST FAILED: $($_.Exception.Message) - save it by hand." }
            Say "rested on $($checkpointActions['restedOn']); settings staged=$($checkpointActions['settingsStaged']); saved=$($checkpointActions['saved'])"
        }
    }
    try {
        if ((Get-LeaseDecision (Get-TestbedLease -VMName $facts.Name) $runId ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())) -eq 'ours') {
            & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $facts.Name -Release | Out-Null
            Say 'lease released'
        }
    }
    catch { Say "releasing the lease failed: $($_.Exception.Message)" }
    if (Test-Path -LiteralPath $tree) {
        Invoke-NativeCommand { & git -C $repo worktree remove --force $tree 2>&1 } | Out-Null
    }
}

# ---- REPORT.
$verdict = Get-LiveVerdict $stage $stageWhy $trx $consoleFacts $crashes $suiteExit $narrowed
$timings['total'] = Format-Seconds ((Get-Date) - $started).TotalSeconds
$summary = New-LiveRunSummary -RunId $runId -Verdict $verdict -Revision $revision -Guest $facts.Name -Filter $runFilter -Narrowed $narrowed -Trx $trx -Console $consoleFacts -Crashes $crashes -SuiteExit $suiteExit -Checkpoints $checkpointActions -Timings $timings -Results $runDir
[System.IO.File]::WriteAllText((Join-Path $runDir 'summary.json'), ($summary | ConvertTo-Json -Depth 8))

$lines = New-Object System.Collections.Generic.List[string]
$head = "== $runId - $($verdict.Verdict) (exit $($ExitCodes[$verdict.Verdict]))"
if ($verdict.Why) { $head += " - $($verdict.Why)" }
$lines.Add("$head ==")
$lines.Add("guest      $($facts.Name) from $Checkpoint; filter $runFilter")
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
$lines.Add("outlook    crashes during the run: $crashText")
$endLine = "guest end  rested on $($checkpointActions['restedOn']); settings staged=$($checkpointActions['settingsStaged']); saved=$($checkpointActions['saved'])"
if ($checkpointActions['green']) { $endLine += "; green checkpoint $($checkpointActions['green'])" }
$lines.Add($endLine)
$lines.Add("time       $(($timings.Keys | ForEach-Object { "$_ $($timings[$_])" }) -join ', ')")
if ($null -ne $trx -and @($trx.Failed).Count -gt 0) {
    $lines.Add(''); $lines.Add('FAILED TESTS')
    foreach ($f in $trx.Failed) { $lines.Add("  $($f.name)"); $lines.Add("      $($f.message)") }
}
if ($null -ne $consoleFacts -and @($consoleFacts.ProvedNothing).Count -gt 0) {
    $lines.Add(''); $lines.Add('PROVED NOTHING')
    foreach ($p in $consoleFacts.ProvedNothing) { $lines.Add("  $($p.Substring(0, [Math]::Min(240, $p.Length)))") }
}
$lines.Add(''); $lines.Add("results    $runDir")
$lines.Add('           summary.txt and summary.json here; results\ the TRX, console.txt and VSTest logs; guest\ every guest call; run.log the steps')
[System.IO.File]::WriteAllLines((Join-Path $runDir 'summary.txt'), $lines)
Write-Host ''
$lines | ForEach-Object { Write-Host $_ }
exit $ExitCodes[$verdict.Verdict]
