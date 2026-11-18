#Requires -Version 5.1
<#
.SYNOPSIS
    Runs Testbed/guest/Measure-TrustRemoval.ps1 on an Outlook test guest in ONE call: the lease, the
    guest's frozen checkpoint, the staging, the probe in the interactive session (RunLevel Limited
    unless told otherwise), the results fetched - and then the checkpoint restored, the guest saved
    and the lease released, on every path.

.DESCRIPTION
    The probe adds a throwaway certificate to the guest user's Root store, plants VSTO trust entries
    and runs the product's installer and uninstaller (Docs/live-tier-on-the-vm.md section 4.8). All of
    that belongs to the guest's checkpoint and none of it may outlive the run, so this script owns the
    whole cycle the way the two runners own theirs:

      1. REFUSE  anything but an Outlook test guest or its clone (TestbedPool.ps1's indexed and
                 unindexed pools), and a guest somebody else holds a live lease on - never taken over.
      2. LEASE   Set-TestbedLease.ps1, taken in one step with reading it (Enter-TestbedLeaseClaim).
      3. RESTORE the guest's frozen checkpoint (testbed.json frozenClocks, its source's for a clone),
                 start it, wait for its heartbeat and for PowerShell Direct.
      4. STAGE   C:\OutlookAI-Q5\trust-removal\: the probe, Register-InteractiveTask.ps1 and
                 OutlookMapiInterop.ps1 from this checkout, and for InstallerCleanup the installer and
                 the Installer.iss it was compiled from.
      5. RUN     the probe through Register-InteractiveTask.ps1 at -RunLevel (Limited unless
                 -RunLevel says Highest), with -Execute.
      6. FETCH   its JSON, logs and certificate into .work\trust-removal-runs\<run>\ of this checkout.
      7. REST    the frozen checkpoint restored again, the guest SAVED, the lease released.

    The guest's clock is frozen (Q130 (a)); nothing here restarts it, and the restore at the end puts
    it back where the checkpoint holds it.

    EXIT CODES: 0 the probe ran and exited 0; 1 the probe ran and failed or found a FAIL;
    3 not run (the VM, the lease, PowerShell Direct, a time limit); 4 refused.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`, ASCII only.

.PARAMETER VMName
    The guest: OutlookAI-Indexed, OutlookAI-Unindexed or one of their clones. No default.

.PARAMETER Scenario
    RootStorePrompt or InstallerCleanup - Measure-TrustRemoval.ps1's.

.PARAMETER InstallerPath
    InstallerCleanup: the compiled installer on this host - the OutlookAI-v*.exe inside the AddIn.zip
    Testbed/host/Publish-AddInPayload.ps1 builds from a commit. Installer.iss is taken from the same
    commit's archive beside it when one is there, else from this checkout.

.PARAMETER HoldSeconds
    RootStorePrompt: how long the probe leaves a removal prompt unanswered.

.PARAMETER RunLevel
    Limited (the default): not elevated, the token a per-user installer and an auto-update get.
    Highest: elevated - for InstallerCleanup on a guest whose add-in was installed elevated
    (Testbed/guest/Install-OutlookAIAddIn.ps1 -Phase Install runs the installer from an elevated
    task): there, at Limited, the uninstaller ended with exit 1 after two minutes having done
    nothing, and elevated it worked (2026-11-17; most likely it asked for elevation unanswered).

.PARAMETER CredentialRepoRoot
    The checkout holding the gitignored guest credential (Get-GuestCredential.ps1). Default: the
    main checkout of this repository, found through git.

.PARAMETER Execute
    Without it the script says what it would do and exits 0, touching nothing.

.PARAMETER SelfTest
    The pure decisions against synthetic inputs. No Hyper-V, no lease, no guest.

.EXAMPLE
    pwsh -File Testbed/host/Invoke-TrustRemovalProbe.ps1 -VMName OutlookAI-Unindexed2 -Scenario RootStorePrompt -Execute
    pwsh -File Testbed/host/Invoke-TrustRemovalProbe.ps1 -VMName OutlookAI-Unindexed2 -Scenario InstallerCleanup -InstallerPath .work\addin\stage\OutlookAI-v99.99.99.0.exe -Execute
#>
[CmdletBinding()]
param(
    [string] $VMName,
    [ValidateSet('', 'RootStorePrompt', 'InstallerCleanup')] [string] $Scenario = '',
    [string] $InstallerPath,
    [int]    $HoldSeconds = 120,
    [ValidateSet('Limited', 'Highest')] [string] $RunLevel = 'Limited',
    [string] $CredentialRepoRoot,
    [switch] $Execute,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TestbedPool.ps1')
. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
# Every program call goes through Invoke-NativeCommand, so that Windows PowerShell 5.1 cannot end
# the script on a program's first stderr line (Q78). It is one of the helpers the scripts share,
# dot-sourced from Testbed/guest/ScriptHelpers.ps1 (Q154, Q166 (b), decision 215 (a)), which says why;
# so are Say, Stop-Refused, the run id, the lease decision and the guest task's exit code, which the
# live runner uses too.
. (Join-Path $PSScriptRoot '..\guest\ScriptHelpers.ps1')

$GuestDir = 'C:\OutlookAI-Q5\trust-removal'
$LeaseReasonPrefix = 'Invoke-TrustRemovalProbe'
$ExitCodes = [ordered]@{ PASS = 0; FAIL = 1; INFRA = 3; REFUSED = 4 }
$ProbeTimeoutSeconds = 1800

# =================================================================================================
# PURE DECISIONS. -SelfTest pins them.
# =================================================================================================

# Why this VM cannot be probed, or '' when it can: an Outlook guest or its clone, and nothing else.
function Get-VmRefusal {
    param([string] $Name)
    if (-not $Name) { return 'name the guest with -VMName: no default picks one (Testbed/README.md section 4a).' }
    $pool = Get-TestbedPoolOf $Name
    if ($pool -ne 'indexed' -and $pool -ne 'unindexed') {
        return "'$Name' is not an Outlook test guest - only $((@($TestbedPools['indexed']) + @($TestbedPools['unindexed'])) -join ', '). Never the build VM and never OutlookAI-Exchange, which holds a real sign-in."
    }
    return ''
}

# The checkpoint the guest rests on: its pool source's frozen checkpoint, which a clone carries under
# the same name (testbed.json vmPools).
function Get-RestingCheckpoint {
    param([string] $Name, [object] $Testbed)
    $source = Get-TestbedPoolSource $Name
    if (-not $source) { return '' }
    $record = $Testbed.frozenClocks.$source
    if ($null -eq $record -or -not $record.checkpoint) { return '' }
    return [string]$record.checkpoint
}

# The interactive task's script: the staged probe with this run's arguments. Every value is one this
# script chose, never free text, so it is written into the task's script as it is.
function Get-ProbeCommand {
    param([string] $Scenario, [int] $HoldSeconds, [bool] $WithInstaller, [string] $InstallerName)
    $command = "& '$GuestDir\Measure-TrustRemoval.ps1' -Scenario $Scenario -Execute -LogDir '$GuestDir\results'"
    if ($Scenario -eq 'RootStorePrompt') { $command += " -HoldSeconds $HoldSeconds" }
    if ($WithInstaller) { $command += " -InstallerPath '$GuestDir\$InstallerName' -InstallerScript '$GuestDir\Installer.iss'" }
    return $command
}

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $name, [bool] $ok) { if ($ok) { $script:p++; Write-Host "  PASS  $name" } else { $script:f++; Write-Host "  FAIL  $name" } }

    Check 'no VM name is refused' ((Get-VmRefusal '') -ne '')
    Check 'the build VM is refused' ((Get-VmRefusal 'OutlookAI-Build') -ne '')
    Check 'the Exchange VM is refused' ((Get-VmRefusal 'OutlookAI-Exchange') -ne '')
    Check 'an Outlook guest is accepted' ((Get-VmRefusal 'OutlookAI-Unindexed') -eq '')
    Check 'a clone is accepted' ((Get-VmRefusal 'OutlookAI-Unindexed2') -eq '')

    $testbed = [pscustomobject]@{ frozenClocks = [pscustomobject]@{ 'OutlookAI-Unindexed' = [pscustomobject]@{ checkpoint = 'CP-X-FROZEN' } } }
    Check 'a guest rests on its frozen checkpoint' ((Get-RestingCheckpoint 'OutlookAI-Unindexed' $testbed) -eq 'CP-X-FROZEN')
    Check 'a clone rests on its source''s, by the same name' ((Get-RestingCheckpoint 'OutlookAI-Unindexed2' $testbed) -eq 'CP-X-FROZEN')
    Check 'a guest with no frozen record has none' ((Get-RestingCheckpoint 'OutlookAI-Indexed' $testbed) -eq '')
    $repoJson = Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json'
    if (Test-Path -LiteralPath $repoJson) {
        $real = Get-Content -LiteralPath $repoJson -Raw | ConvertFrom-Json
        foreach ($g in @($TestbedPools['indexed']) + @($TestbedPools['unindexed'])) {
            Check "testbed.json gives $g a frozen checkpoint" ((Get-RestingCheckpoint $g $real) -like 'CP-*FROZEN*')
        }
    }

    $mine = "$LeaseReasonPrefix r"
    Check 'a run''s lease reason is this probe''s name and the run id' ($mine -ceq 'Invoke-TrustRemovalProbe r')
    Check 'no lease is free' ((Get-LeaseDecision $null $mine 100) -eq 'free')
    Check 'an expired lease is free' ((Get-LeaseDecision ([pscustomobject]@{ expiresUnix = 99; reason = 'x' }) $mine 100) -eq 'free')
    Check 'this run''s lease is ours' ((Get-LeaseDecision ([pscustomobject]@{ expiresUnix = 200; reason = $mine }) $mine 100) -eq 'ours')
    Check 'anyone else''s live lease is theirs' ((Get-LeaseDecision ([pscustomobject]@{ expiresUnix = 200; reason = 'Invoke-LiveTierOnGuest r' }) $mine 100) -eq 'theirs')
    $own = [System.IO.File]::ReadAllText($PSCommandPath)
    # The run's code only: what follows the -SelfTest dispatch, below this function.
    $runCode = $own.Substring($own.LastIndexOf('if ($SelfTest) { Invoke-SelfTest }'))
    Check 'the probe asks with its run''s own reason' ($runCode.Contains('Get-LeaseDecision (Get-TestbedLease -VMName $VMName) "$LeaseReasonPrefix $runId"'))

    $c = Get-ProbeCommand -Scenario 'RootStorePrompt' -HoldSeconds 120 -WithInstaller $false -InstallerName ''
    Check 'the probe runs with -Execute and the hold' ($c -ceq "& '$GuestDir\Measure-TrustRemoval.ps1' -Scenario RootStorePrompt -Execute -LogDir '$GuestDir\results' -HoldSeconds 120")
    $c = Get-ProbeCommand -Scenario 'InstallerCleanup' -HoldSeconds 120 -WithInstaller $true -InstallerName 'OutlookAI-v1.exe'
    Check 'the installer round names the staged installer and Installer.iss' ($c.EndsWith("-InstallerPath '$GuestDir\OutlookAI-v1.exe' -InstallerScript '$GuestDir\Installer.iss'"))
    Check 'the task exit is read back' ((Get-TaskExit "x`nRUN-REGISTER-EXIT: 1`n") -eq 1)
    Check 'and the probe''s call writes the line it is read from' ($runCode.Contains('"RUN-REGISTER-EXIT: $LASTEXITCODE"'))
    Check 'no exit line is no exit' ($null -eq (Get-TaskExit 'nothing'))

    Write-Host ''
    Write-Host "SelfTest: $script:p passed, $script:f failed"
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

$refusal = Get-VmRefusal $VMName
if ($refusal) { Stop-Refused $refusal }
if (-not $Scenario) { Stop-Refused 'say which -Scenario: RootStorePrompt or InstallerCleanup.' }
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testbed = Get-Content -LiteralPath (Join-Path $repo 'Testbed\testbed.json') -Raw | ConvertFrom-Json
$checkpoint = Get-RestingCheckpoint $VMName $testbed
if (-not $checkpoint) { Stop-Refused "testbed.json records no frozen checkpoint for $VMName's pool." }
$installerIss = $null
if ($Scenario -eq 'InstallerCleanup') {
    if (-not $InstallerPath -or -not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) { Stop-Refused "InstallerCleanup needs -InstallerPath, the compiled installer: '$InstallerPath'." }
    $InstallerPath = (Resolve-Path -LiteralPath $InstallerPath).Path
    $installerIss = Join-Path $repo 'Installer.iss'
}
if (-not $CredentialRepoRoot) {
    $common = (@(Invoke-NativeCommand { & git -C $repo rev-parse --path-format=absolute --git-common-dir 2>&1 }) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { Stop-Refused "$repo is not a git checkout: $common" }
    $CredentialRepoRoot = Split-Path -Parent ([System.IO.Path]::GetFullPath($common))
}
$sha = (@(Invoke-NativeCommand { & git -C $repo rev-parse HEAD 2>&1 }) | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { Stop-Refused "git could not read HEAD in $repo - $sha" }
$dirty = @(Invoke-NativeCommand { & git -C $repo status --porcelain -- 'Testbed/guest/Measure-TrustRemoval.ps1' 'Testbed/guest/Register-InteractiveTask.ps1' 'Testbed/guest/OutlookMapiInterop.ps1' 'Installer.iss' 2>&1 } | Where-Object { $_ })

$runId = New-RunId (Get-Date) $sha $Scenario
$runDir = Join-Path $repo ".work\trust-removal-runs\$runId"
$command = Get-ProbeCommand -Scenario $Scenario -HoldSeconds $HoldSeconds -WithInstaller ($Scenario -eq 'InstallerCleanup') -InstallerName ([System.IO.Path]::GetFileName([string]$InstallerPath))

Write-Host "== Invoke-TrustRemovalProbe: $Scenario on $VMName at RunLevel $RunLevel, from and resting on $checkpoint"
Write-Host "   revision  $sha$(if ($dirty.Count -gt 0) { " - NOTE: $($dirty.Count) of the staged files differ from it; the run stages the working tree's copies" })"
Write-Host "   probe     $command"
Write-Host "   results   $runDir"
if (-not $Execute) {
    Write-Host 'DRY RUN - nothing touched. Add -Execute to run it.'
    exit 0
}

New-Item -ItemType Directory -Force -Path $runDir | Out-Null
$script:LogPath = Join-Path $runDir 'run.log'
$verdict = 'INFRA'
$why = ''
$leased = $false
$touched = $false
$session = $null
try {
    $claim = Enter-TestbedLeaseClaim -VMName $VMName
    try {
        $decision = Get-LeaseDecision (Get-TestbedLease -VMName $VMName) "$LeaseReasonPrefix $runId" ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
        if ($decision -eq 'theirs') { throw "the lease on $VMName is held by somebody else - not touching the guest" }
        & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $VMName -Minutes 60 -Reason "$LeaseReasonPrefix $runId" | Out-Null
        $leased = $true
    }
    finally { $claim.Dispose() }
    Say "leased $VMName"

    if (-not (Get-VMSnapshot -VMName $VMName -Name $checkpoint -ErrorAction SilentlyContinue)) { throw "$VMName has no checkpoint named $checkpoint" }
    $touched = $true
    Restore-VMSnapshot -VMName $VMName -Name $checkpoint -Confirm:$false
    if ((Get-VM -Name $VMName).State -ne 'Running') { Start-VM -Name $VMName }
    $deadline = (Get-Date).AddMinutes(5)
    while ((Get-VMIntegrationService -VMName $VMName -Name 'Heartbeat').PrimaryStatusDescription -ne 'OK') {
        if ((Get-Date) -gt $deadline) { throw "the guest's heartbeat did not come back within 5 minutes" }
        Start-Sleep -Seconds 3
    }
    Say "restored $checkpoint and running"

    $credential = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $CredentialRepoRoot -VMName $VMName
    $deadline = (Get-Date).AddMinutes(3)
    while ($null -eq $session) {
        try { $session = New-PSSession -VMName $VMName -Credential $credential -ErrorAction Stop }
        catch {
            if ((Get-Date) -gt $deadline) { throw "PowerShell Direct did not answer within 3 minutes: $($_.Exception.Message)" }
            Start-Sleep -Seconds 5
        }
    }
    Invoke-Command -Session $session -ScriptBlock { param($d) New-Item -ItemType Directory -Force -Path $d, (Join-Path $d 'results') | Out-Null } -ArgumentList $GuestDir
    $stage = @(
        (Join-Path $repo 'Testbed\guest\Measure-TrustRemoval.ps1'),
        (Join-Path $repo 'Testbed\guest\Register-InteractiveTask.ps1'),
        (Join-Path $repo 'Testbed\guest\OutlookMapiInterop.ps1'))
    if ($Scenario -eq 'InstallerCleanup') { $stage += @($InstallerPath, $installerIss) }
    foreach ($f in $stage) { Copy-Item -LiteralPath $f -Destination (Join-Path $GuestDir (Split-Path -Leaf $f)) -ToSession $session -Force }
    Say "staged $($stage.Count) file(s) in $GuestDir"

    & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $VMName -Minutes 60 -Reason "$LeaseReasonPrefix $runId" | Out-Null
    Say "running the probe at RunLevel $RunLevel (up to $ProbeTimeoutSeconds s)"
    $out = Invoke-Command -Session $session -ScriptBlock {
        param($dir, $text, $timeout, $level)
        $ErrorActionPreference = 'Continue'
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        try { & (Join-Path $dir 'Register-InteractiveTask.ps1') -Script $text -TimeoutSeconds $timeout -RunLevel $level *>&1 | Out-String -Width 400 } catch { "GUEST-CALL-THREW: $($_.Exception.Message)" }
        "RUN-REGISTER-EXIT: $LASTEXITCODE"
    } -ArgumentList $GuestDir, $command, $ProbeTimeoutSeconds, $RunLevel 2>&1 | Out-String -Width 400
    [System.IO.File]::WriteAllText((Join-Path $runDir 'probe.out.txt'), $out)
    Write-Host $out
    $exit = Get-TaskExit $out

    Copy-Item -Path "$GuestDir\results\*" -Destination $runDir -FromSession $session -Recurse -Force -ErrorAction SilentlyContinue
    Say "fetched $(@(Get-ChildItem -LiteralPath $runDir -File).Count) file(s) into $runDir"
    if ($null -eq $exit) { $why = 'the probe task reported no exit code' }
    elseif ($exit -eq 0) { $verdict = 'PASS' }
    else { $verdict = 'FAIL'; $why = "the probe exited $exit" }
}
catch {
    $why = $_.Exception.Message
    Say "STOPPED: $why"
}
finally {
    if ($null -ne $session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    if ($touched) {
        try {
            Restore-VMSnapshot -VMName $VMName -Name $checkpoint -Confirm:$false
            if ((Get-VM -Name $VMName).State -eq 'Running') { Save-VM -Name $VMName }
            Say "rested: $checkpoint restored, $VMName $((Get-VM -Name $VMName).State)"
        }
        catch { Say "REST FAILED - restore $checkpoint on $VMName and save it by hand: $($_.Exception.Message)"; $verdict = 'INFRA' }
    }
    if ($leased) { & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $VMName -Release | Out-Null; Say "released the lease on $VMName" }
}
Say "VERDICT: $verdict$(if ($why) { " - $why" })"
exit $ExitCodes[$verdict]
