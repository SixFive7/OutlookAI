#Requires -Version 5.1
<#
.SYNOPSIS
    Take, renew or release a lease on a testbed VM so the idle-saver leaves it alone.

.DESCRIPTION
    THE POINT OF THIS FILE. The testbed VMs are SAVED (not paused) whenever nobody is using
    them, so the host gets its RAM and CPU back. Something has to decide "nobody is using it",
    and the obvious answer - watch the guest's CPU - is wrong here in a way that would be found
    as a mystery rather than as a bug.

    A live tier run is roughly 27 minutes of driving Outlook through COM, and Outlook spends a
    great deal of that waiting: on a store to open, on a folder to enumerate, on a save to
    commit. A guest that looks idle for two minutes in the middle of a run is entirely
    ordinary. Saving it there suspends the run mid-COM-call, and what the operator sees
    afterwards is a test that timed out for no reason on a machine that looks fine.

    So usage is DECLARED, not inferred. Anything intending to use a VM takes a lease with an
    expiry and renews it while it works. The idle-saver refuses to touch a VM holding a live
    lease. A lease whose holder died simply stops protecting the VM when it expires, which is
    the right failure - the alternative is a VM pinned awake for ever by a process that no
    longer exists.

    Leases live in a fixed machine-wide directory, not in the repository (they are machine
    state, not source) and not in a per-user profile (the saver runs as a scheduled task and
    has to read the same leases the operator writes).

    A LEASE IS NEVER TAKEN FROM A HOLDER WHO STILL HAS IT, NOR RELEASED UNDER ONE (2026-11-20).
    Until then this script wrote over whatever lease was there, and -Release deleted whatever lease
    was there: a person who leased a VM a run was using took it from the run without a word, and a
    live-tier run stopped at its next renewal, which found the lease somebody else's. Now it takes,
    or releases, a lease only when
      - there is none, or it has expired or cannot be read (TestbedLeasePath.ps1's Get-TestbedLease);
      - it is the caller's own: the same -Reason, which is how the runners tell their own lease -
        each run's reason holds its id;
      - or the process that took it has ended - a run that was stopped, a host that restarted -
        judged as the queue's tickets are (TestbedPool.ps1's Test-TestbedTicketAlive): its process
        id, started at the second the lease records, and not another process given that id since.
    The decision is the runners' own, Testbed/guest/ScriptHelpers.ps1's Get-LeaseDecision, so a run
    takes over a lease whose process has ended as this script does. Anything else is refused, and the
    refusal says who holds the lease and until when. A lease taken BY HAND - this script run from a
    prompt, not called by a script - has no process that lives as long as its holder, so only its
    expiry or a release with its -Reason ends it. A lease written before 2026-11-20 does not say which
    it is, and is judged by its process. By hand, the read and the write are one step under the
    lease's claim (TestbedLeasePath.ps1's Enter-TestbedLeaseClaim); a script holds that claim around
    its call itself, as the runners do.

.PARAMETER VMName
    MANDATORY, and always has been. The VM to lease. There is no default here for the same
    reason there is none anywhere else in Testbed/host/: FOUR MACHINES COEXIST -
    OutlookAI-Indexed, OutlookAI-Unindexed, OutlookAI-Build and OutlookAI-Exchange - and a default that silently
    picks one of four is the exact shape of mistake this testbed keeps making. A lease taken on
    the wrong machine protects nothing and reads, in the log, exactly like one that worked.

.PARAMETER Minutes
    How long the lease lasts from now. Renew before it expires. Do NOT take a very long lease
    to avoid renewing: a crashed holder then pins the VM awake for that long.

.PARAMETER Release
    Drop the lease now instead of taking one. Do this when the work finishes, so the saver can
    reclaim the resources without waiting for the expiry. Give the -Reason the lease was taken with:
    another holder's live lease is refused, as a take is.

.PARAMETER Reason
    Free text recorded in the lease, so an operator who finds a VM awake can see why. It is also
    what makes a lease its taker's own: renew and release one with the -Reason it was taken with.

.PARAMETER SelfTest
    Checks the rule above on leases in a scratch directory - never the machine's - with two helper
    processes of its own, started hidden and stopped before it ends, and that every script that
    releases a lease gives its -Reason. No Hyper-V. Runs on the build VM with every other -SelfTest.

.EXAMPLE
    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Minutes 45 -Reason 'live tier'
    # ... run the tier, renewing every few minutes with the same -Reason ...
    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Release -Reason 'live tier'
#>
[CmdletBinding(DefaultParameterSetName = 'Take')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Take')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Release')] [string] $VMName,
    [Parameter(ParameterSetName = 'Take')] [int] $Minutes = 30,
    [Parameter(Mandatory = $true, ParameterSetName = 'Release')] [switch] $Release,
    [Parameter(ParameterSetName = 'Take')]
    [Parameter(ParameterSetName = 'Release')] [string] $Reason = '',
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
# Test-TestbedTicketAlive: whether a lease's holder still runs, judged as a queue ticket's run is.
. (Join-Path $PSScriptRoot 'TestbedPool.ps1')
# Get-LeaseDecision and Get-LeaseHolderState: the runners' judgment of a lease, which this one shares.
. (Join-Path $PSScriptRoot '..\guest\ScriptHelpers.ps1')

# Who is asking: the script that called this one, or '' when a person ran it - $MyInvocation names a
# calling script only when there is one, under Windows PowerShell 5.1 and PowerShell 7 alike.
$CallerScript = ''
if ($MyInvocation.PSCommandPath) { $CallerScript = Split-Path -Leaf $MyInvocation.PSCommandPath }

# A Unix second count for a person to read, in UTC as the line a lease taken prints its expiry.
function Format-LeaseTime([long] $Unix) {
    return [DateTimeOffset]::FromUnixTimeSeconds($Unix).UtcDateTime.ToString('u')
}

# Whether a caller giving $Reason may take - with -Release, release - $Lease, the live lease
# Get-TestbedLease read or $null, and what to say: Take, and Text - the refusal when it may not, what
# it took over or released when the holder's process had ended, '' otherwise. The decision is
# Get-LeaseDecision's, the runners' own; $IsAlive answers for a holder's process, given as a ticket is
# ({ pid, processStartUnix }), so that the -SelfTest can answer for it too.
function Get-LeaseVerdict {
    param([string] $VMName, [object] $Lease, [string] $Reason, [long] $NowUnix, [string] $ComputerName, [scriptblock] $IsAlive, [switch] $Release)

    $decision = Get-LeaseDecision -Lease $Lease -Reason $Reason -NowUnix $NowUnix -IsAlive $IsAlive -ComputerName $ComputerName
    if ($decision -eq 'free' -or $decision -eq 'ours') { return [pscustomobject]@{ Take = $true; Text = '' } }

    $held = [string]$Lease.reason
    $by = "'$held'"
    if (-not $held) { $by = 'a holder who gave no reason' }
    $holderPid = 0
    [void][int]::TryParse([string]$Lease.holderPid, [ref]$holderPid)
    if ($decision -eq 'gone') {
        $verb = 'took over'
        if ($Release) { $verb = 'released' }
        return [pscustomobject]@{ Take = $true; Text = "It $verb the lease of $by, whose process $holderPid has ended." }
    }

    $taken = ''
    $takenUnix = 0L
    if ([long]::TryParse([string]$Lease.takenUnix, [ref]$takenUnix) -and $takenUnix -gt 0) { $taken = ' at ' + (Format-LeaseTime $takenUnix) }
    $expiresUnix = [long]$Lease.expiresUnix
    $left = [long][Math]::Ceiling(($expiresUnix - $NowUnix) / 60.0)
    $refusing = 'REFUSING'
    if ($Release) { $refusing = 'REFUSING to release' }
    $head = "${refusing}: $VMName is leased until $(Format-LeaseTime $expiresUnix) ($left min from now) by $by"
    $process = "process $holderPid"
    if ($Lease.holderScript) { $process = "$([string]$Lease.holderScript), process $holderPid" }
    $yours = 'if it is yours, renew it with the same -Reason'
    if ($Release) { $yours = 'if it is yours, release it with the -Reason it was taken with' }

    switch (Get-LeaseHolderState -Lease $Lease -ComputerName $ComputerName -IsAlive $IsAlive) {
        'byHand' { $text = "$head, taken by hand$taken. A lease taken by hand ends only when it expires or is released: $yours; if not, ask its holder, or wait." }
        'noProcess' { $text = "$head$taken. It names no process, so only its expiry or a release with its -Reason ends it." }
        'elsewhere' { $text = "$head, taken by $process on $([string]$Lease.holderHost)$taken, which this computer cannot ask about. Wait for it to expire or be released." }
        default { $text = "$head, taken by $process$taken, which is still running: $yours; if not, ask its holder, or wait until the lease is released, expires or its process ends." }
    }
    return [pscustomobject]@{ Take = $false; Text = $text }
}

# Takes, or renews, the lease on $VMName for $Minutes - or refuses, by throwing the verdict's text: the
# script's work, a function so that the -SelfTest can call it both as a script and as a person would.
# $CallerScript is the calling script's file name, '' by hand; by hand, under the lease's claim.
function Set-Lease {
    param([string] $VMName, [int] $Minutes, [string] $Reason, [string] $CallerScript, [int] $ClaimTimeoutSeconds = 60)

    $leasePath = Join-Path (Get-TestbedLeaseDirectory) "$VMName.lease.json"
    $byHand = -not $CallerScript
    $claim = $null
    if ($byHand) { $claim = Enter-TestbedLeaseClaim -VMName $VMName -TimeoutSeconds $ClaimTimeoutSeconds }
    try {
        $now = [DateTime]::UtcNow
        $nowUnix = [DateTimeOffset]::new($now, [TimeSpan]::Zero).ToUnixTimeSeconds()
        $verdict = Get-LeaseVerdict -VMName $VMName -Lease (Get-TestbedLease -VMName $VMName) -Reason $Reason -NowUnix $nowUnix -ComputerName $env:COMPUTERNAME -IsAlive ${function:Test-TestbedTicketAlive}
        if (-not $verdict.Take) { throw $verdict.Text }

        $expires = $now.AddMinutes($Minutes)
        [pscustomobject]@{
            vmName          = $VMName
            takenUtc        = $now.ToString('o')
            takenUnix       = $nowUnix
            expiresUtc      = $expires.ToString('o')
            # The number is what the saver compares. ConvertFrom-Json silently coerces an ISO-8601
            # STRING into a DateTime, and re-parsing that object's local rendering loses the UTC
            # marker - which made every lease read as already expired, two hours in the past.
            # A Unix second count cannot be coerced into anything but a number.
            expiresUnix     = [long]([DateTimeOffset]::new($expires, [TimeSpan]::Zero).ToUnixTimeSeconds())
            # The holder, as a queue ticket names its run: the process, and the second it started.
            holderPid       = $PID
            holderStartUnix = ([DateTimeOffset](Get-Process -Id $PID).StartTime).ToUnixTimeSeconds()
            holderHost      = $env:COMPUTERNAME
            holderScript    = $CallerScript
            byHand          = $byHand
            reason          = $Reason
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $leasePath -Encoding UTF8

        $line = "Leased $VMName until $($expires.ToString('u')) ($Minutes min)$(if ($Reason) { " - $Reason" })."
        if ($verdict.Text) { $line += " $($verdict.Text)" }
        return $line
    }
    finally { if ($null -ne $claim) { $claim.Dispose() } }
}

# Releases the lease on $VMName - or refuses, by throwing the verdict's text - by the rule a lease is
# taken by (Get-LeaseVerdict -Release): the caller's own by its -Reason, one that has expired, or one
# whose process has ended. By hand, under the lease's claim, as a lease is taken.
function Remove-Lease {
    param([string] $VMName, [string] $Reason, [string] $CallerScript, [int] $ClaimTimeoutSeconds = 60)

    $leasePath = Join-Path (Get-TestbedLeaseDirectory) "$VMName.lease.json"
    if (-not (Test-Path -LiteralPath $leasePath)) { return "No lease on $VMName to release." }
    $claim = $null
    if (-not $CallerScript) { $claim = Enter-TestbedLeaseClaim -VMName $VMName -TimeoutSeconds $ClaimTimeoutSeconds }
    try {
        $verdict = Get-LeaseVerdict -VMName $VMName -Lease (Get-TestbedLease -VMName $VMName) -Reason $Reason -NowUnix ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) -ComputerName $env:COMPUTERNAME -IsAlive ${function:Test-TestbedTicketAlive} -Release
        if (-not $verdict.Take) { throw $verdict.Text }
        if (Test-Path -LiteralPath $leasePath) { Remove-Item -LiteralPath $leasePath -Force }
        $line = "Released the lease on $VMName."
        if ($verdict.Text) { $line += " $($verdict.Text)" }
        return $line
    }
    finally { if ($null -ne $claim) { $claim.Dispose() } }
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
    # What Set-Lease or Remove-Lease threw, or '' when it did not.
    function Get-Refusal([scriptblock] $Call) {
        try { [void](& $Call); return '' } catch { return [string]$_.Exception.Message }
    }
    Write-Host "Set-TestbedLease.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Leases on a scratch directory; no Hyper-V."

    $now = 1800000000L
    $me = 'OAI-SELFTEST'
    $alive = { param($r) $true }
    $ended = { param($r) $false }
    function New-Lease([string] $Why, [object] $ByHand, [int] $ProcessId = 4242, [string] $Computer = 'OAI-SELFTEST', [string] $Script = 'Invoke-TestsOnBuildVm.ps1') {
        $l = [ordered]@{ vmName = 'OutlookAI-Build'; takenUnix = $now - 600; expiresUnix = $now + 1800; holderPid = $ProcessId; holderStartUnix = $now - 900; holderHost = $Computer; holderScript = $Script; reason = $Why }
        if ($null -ne $ByHand) { $l['byHand'] = [bool]$ByHand }
        return [pscustomobject]$l
    }

    Write-Host ''
    Write-Host '== the rule, judged =='
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease $null -Reason 'x' -NowUnix $now -ComputerName $me -IsAlive $alive
    Check 'no live lease - none, expired or unreadable: taken' 'True|' @($v.Take, $v.Text)
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason 'build-vm run A' -NowUnix $now -ComputerName $me -IsAlive $alive
    Check 'the taker''s own, by its -Reason: renewed' 'True|' @($v.Take, $v.Text)
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason 'Build-VM run A' -NowUnix $now -ComputerName $me -IsAlive $alive
    Check 'a reason compares ordinally: another case is another holder' $false $v.Take
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $alive
    Check 'a run''s lease whose process runs: refused' $false $v.Take
    Check 'the refusal says whose lease it is and until when' $true ($v.Text.StartsWith("REFUSING: OutlookAI-Build is leased until $(Format-LeaseTime ($now + 1800)) (30 min from now) by 'build-vm run A'"))
    Check 'and who took it, and that it still runs' $true ($v.Text.Contains("taken by Invoke-TestsOnBuildVm.ps1, process 4242 at $(Format-LeaseTime ($now - 600)), which is still running"))
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $ended
    Check 'a run''s lease whose process has ended: taken over' $true $v.Take
    Check 'and it says what it took over' "It took over the lease of 'build-vm run A', whose process 4242 has ended." $v.Text
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'maintenance' $true -Script '') -Reason 'another' -NowUnix $now -ComputerName $me -IsAlive $ended
    Check 'a lease taken by hand is not taken over when its process has ended' $false $v.Take
    Check 'the refusal says it was taken by hand, and how its holder renews it' $true ($v.Text.Contains("by 'maintenance', taken by hand at ") -and $v.Text.Contains('renew it with the same -Reason'))
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'maintenance' $true -Script '') -Reason 'maintenance' -NowUnix $now -ComputerName $me -IsAlive $ended
    Check 'but its holder renews it, by its -Reason' $true $v.Take
    $old = [pscustomobject]@{ vmName = 'OutlookAI-Build'; takenUtc = '2026-11-18T19:30:00.0000000Z'; expiresUtc = '2026-11-18T20:30:00.0000000Z'; expiresUnix = $now + 1800; holderPid = 4242; holderHost = $me; reason = 'build-vm run old' }
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease $old -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $alive
    Check 'a lease from before 2026-11-20 whose process runs: refused' $false $v.Take
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease $old -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $ended
    Check 'and once its process has ended, taken over' $true $v.Take
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false -Computer 'OAI-ELSEWHERE') -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $ended
    Check 'a lease another computer''s process took is not judged from here' $false $v.Take
    Check 'and the refusal says so' $true ($v.Text.Contains('process 4242 on OAI-ELSEWHERE') -and $v.Text.Contains('which this computer cannot ask about'))
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false -ProcessId 0) -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $ended
    Check 'a lease that names no process ends only when it expires or is released' $false $v.Take
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease '' $false) -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive $alive
    Check 'a holder who gave no reason is called that' $true ($v.Text.Contains('by a holder who gave no reason'))
    $asked = $null
    [void](Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason 'by hand' -NowUnix $now -ComputerName $me -IsAlive { param($r) $script:asked = $r; $true })
    Check 'a holder''s process is asked about as a queue ticket''s: its id and the second it started' "4242|$($now - 900)" @($script:asked.pid, $script:asked.processStartUnix)
    Check 'the judgment is the runners'' own: a lease whose process has ended is gone to them too' 'gone|theirs|theirs' @((Get-LeaseDecision -Lease (New-Lease 'build-vm run A' $false) -Reason 'r' -NowUnix $now -IsAlive $ended -ComputerName $me), (Get-LeaseDecision -Lease (New-Lease 'build-vm run A' $false) -Reason 'r' -NowUnix $now -IsAlive $alive -ComputerName $me), (Get-LeaseDecision -Lease (New-Lease 'maintenance' $true -Script '') -Reason 'r' -NowUnix $now -IsAlive $ended -ComputerName $me))
    Check 'and a caller that gives no -IsAlive reads every holder as alive, as before 2026-11-20' 'theirs' (Get-LeaseDecision -Lease (New-Lease 'build-vm run A' $false) -Reason 'r' -NowUnix $now -ComputerName $me)

    Write-Host ''
    Write-Host '== a release, judged =='
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason 'build-vm run A' -NowUnix $now -ComputerName $me -IsAlive $alive -Release
    Check 'the holder releases its own lease by its -Reason' 'True|' @($v.Take, $v.Text)
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason '' -NowUnix $now -ComputerName $me -IsAlive $alive -Release
    Check 'anyone else is refused while its process runs' $false $v.Take
    Check 'and told whose lease it is, until when, and how its holder releases it' $true ($v.Text.StartsWith("REFUSING to release: OutlookAI-Build is leased until $(Format-LeaseTime ($now + 1800)) (30 min from now) by 'build-vm run A', taken by Invoke-TestsOnBuildVm.ps1, process 4242") -and $v.Text.Contains('release it with the -Reason it was taken with'))
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'build-vm run A' $false) -Reason '' -NowUnix $now -ComputerName $me -IsAlive $ended -Release
    Check 'a lease whose process has ended is released by anyone, and says whose it was' "True|It released the lease of 'build-vm run A', whose process 4242 has ended." @($v.Take, $v.Text)
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease (New-Lease 'maintenance' $true -Script '') -Reason '' -NowUnix $now -ComputerName $me -IsAlive $ended -Release
    Check 'a lease taken by hand only by its -Reason' 'False|True' @($v.Take, ($v.Text.Contains("by 'maintenance', taken by hand at ") -and $v.Text.Contains('release it with the -Reason it was taken with')))
    $v = Get-LeaseVerdict -VMName 'OutlookAI-Build' -Lease $null -Reason '' -NowUnix $now -ComputerName $me -IsAlive $alive -Release
    Check 'and one that has expired by anyone' 'True|' @($v.Take, $v.Text)

    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-lease-selftest-' + [guid]::NewGuid().ToString('N'))
    $script:TestbedLeaseDirectoryOverride = $scratch
    $shellPath = (Get-Process -Id $PID).Path
    $helperArgs = @('-NoProfile', '-NonInteractive', '-Command', 'Start-Sleep -Seconds 300')
    $live = $null
    $gone = $null
    $claim = $null
    try {
        Write-Host ''
        Write-Host '== a holder''s process, alive or not (TestbedPool.ps1''s Test-TestbedTicketAlive) =='
        # Each helper's handle is read at once (Testbed/guest/Install-DotnetSdk.ps1: a Process from
        # Start-Process -PassThru whose handle was never read cannot be waited on once it has gone).
        $live = Start-Process -FilePath $shellPath -ArgumentList $helperArgs -WindowStyle Hidden -PassThru
        [void]$live.Handle
        $liveStart = ([DateTimeOffset]$live.StartTime).ToUnixTimeSeconds()
        $gone = Start-Process -FilePath $shellPath -ArgumentList $helperArgs -WindowStyle Hidden -PassThru
        [void]$gone.Handle
        $goneStart = ([DateTimeOffset]$gone.StartTime).ToUnixTimeSeconds()
        Stop-Process -Id $gone.Id -Force
        [void]$gone.WaitForExit(30000)
        Check 'a process that runs, started at the second recorded, is alive' $true (Test-TestbedTicketAlive ([pscustomobject]@{ pid = $live.Id; processStartUnix = $liveStart }))
        Check 'one whose id now names a process started at another second is not' $false (Test-TestbedTicketAlive ([pscustomobject]@{ pid = $live.Id; processStartUnix = $liveStart - 3600 }))
        Check 'a process that has ended is not' $false (Test-TestbedTicketAlive ([pscustomobject]@{ pid = $gone.Id; processStartUnix = $goneStart }))
        Check 'a record that does not say when its process started - a lease from before 2026-11-20 - lives while its id runs' $true (Test-TestbedTicketAlive ([pscustomobject]@{ pid = $live.Id }))

        Write-Host ''
        Write-Host '== the script, on a scratch directory =='
        $vm = 'OutlookAI-SelfTest'
        $path = Join-Path (Get-TestbedLeaseDirectory) "$vm.lease.json"
        Check 'the leases are the scratch directory''s, not the machine''s' $scratch (Get-TestbedLeaseDirectory)
        function Read-LeaseFile { return ([System.IO.File]::ReadAllText($path) | ConvertFrom-Json) }
        function Write-LeaseFile([object] $Lease) { [System.IO.File]::WriteAllText($path, ($Lease | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false))) }
        $ownStart = ([DateTimeOffset](Get-Process -Id $PID).StartTime).ToUnixTimeSeconds()

        $out = Set-Lease -VMName $vm -Minutes 30 -Reason 'build-vm run S' -CallerScript 'Invoke-TestsOnBuildVm.ps1'
        $l = Read-LeaseFile
        Check 'a script takes a free lease' $true ([string]$out).StartsWith("Leased $vm until ")
        Check 'and it names the script, its process and the second it started, not by hand' "Invoke-TestsOnBuildVm.ps1|$PID|$ownStart|False|build-vm run S" @([string]$l.holderScript, [string]$l.holderPid, [string]$l.holderStartUnix, [string]$l.byHand, [string]$l.reason)
        Check 'when it was taken and when it expires, as numbers' 1800 ([long]$l.expiresUnix - [long]$l.takenUnix)
        $before = [long]$l.expiresUnix
        Write-LeaseFile ([pscustomobject]@{ vmName = $vm; takenUnix = [long]$l.takenUnix; expiresUnix = $before - 600; holderPid = $PID; holderStartUnix = $ownStart; holderHost = $env:COMPUTERNAME; holderScript = 'Invoke-TestsOnBuildVm.ps1'; byHand = $false; reason = 'build-vm run S' })
        [void](Set-Lease -VMName $vm -Minutes 30 -Reason 'build-vm run S' -CallerScript 'Invoke-TestsOnBuildVm.ps1')
        Check 'the same -Reason renews it' $true ([long](Read-LeaseFile).expiresUnix -ge $before)
        $text = [System.IO.File]::ReadAllText($path)
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'maintenance' -CallerScript '' }
        Check 'a person may not take it while the run''s process runs' $true ($why.StartsWith("REFUSING: $vm is leased until ") -and $why.Contains("by 'build-vm run S', taken by Invoke-TestsOnBuildVm.ps1, process $PID"))
        Check 'and the lease is left as it was' $text ([System.IO.File]::ReadAllText($path))
        $why = Get-Refusal { Remove-Lease -VMName $vm -Reason 'maintenance' -CallerScript '' }
        Check 'nor release it' $true ($why.StartsWith("REFUSING to release: $vm is leased until ") -and $why.Contains("by 'build-vm run S', taken by Invoke-TestsOnBuildVm.ps1, process $PID"))
        Check 'which leaves it as it was' $text ([System.IO.File]::ReadAllText($path))

        Write-LeaseFile ([pscustomobject]@{ vmName = $vm; takenUnix = $now; expiresUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() + 1200; holderPid = $live.Id; holderStartUnix = $liveStart; holderHost = $env:COMPUTERNAME; holderScript = 'Invoke-LiveTierOnGuest.ps1'; byHand = $false; reason = 'live-tier run L' })
        $text = [System.IO.File]::ReadAllText($path)
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'build-vm run S' -CallerScript 'Invoke-TestsOnBuildVm.ps1' }
        Check 'another process''s live lease is refused to a script' $true ($why.Contains("by 'live-tier run L', taken by Invoke-LiveTierOnGuest.ps1, process $($live.Id)") -and $why.Contains('which is still running'))
        Check 'and left as it was' $text ([System.IO.File]::ReadAllText($path))
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'maintenance' -CallerScript '' }
        Check 'and to a person' $true ($why.Contains("by 'live-tier run L'"))
        Stop-Process -Id $live.Id -Force
        [void]$live.WaitForExit(30000)
        $out = Set-Lease -VMName $vm -Minutes 30 -Reason 'maintenance' -CallerScript ''
        $l = Read-LeaseFile
        Check 'once that process has ended, it is taken over' 'maintenance|True|' @([string]$l.reason, [string]$l.byHand, [string]$l.holderScript)
        Check 'and the taker is told whose it was' $true (([string]$out).Contains("It took over the lease of 'live-tier run L', whose process $($live.Id) has ended."))

        Write-LeaseFile ([pscustomobject]@{ vmName = $vm; takenUnix = $now; expiresUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() + 1200; holderPid = $gone.Id; holderStartUnix = $goneStart; holderHost = $env:COMPUTERNAME; holderScript = ''; byHand = $true; reason = 'maintenance' })
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'build-vm run S' -CallerScript 'Invoke-TestsOnBuildVm.ps1' }
        Check 'a lease taken by hand holds though the process that wrote it has ended' $true ($why.Contains("by 'maintenance', taken by hand"))
        $why = Get-Refusal { Remove-Lease -VMName $vm -Reason '' -CallerScript '' }
        Check 'and is not released without its -Reason' $true ($why.StartsWith("REFUSING to release: $vm is leased until ") -and $why.Contains("by 'maintenance', taken by hand"))
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'maintenance' -CallerScript '' }
        Check 'its holder renews it with the same -Reason' '' $why

        Write-LeaseFile ([pscustomobject]@{ vmName = $vm; takenUnix = $now; expiresUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - 5; holderPid = $PID; holderStartUnix = $ownStart; holderHost = $env:COMPUTERNAME; holderScript = ''; byHand = $true; reason = 'maintenance' })
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'build-vm run S' -CallerScript 'Invoke-TestsOnBuildVm.ps1' }
        Check 'an expired lease is taken over, whoever held it' '' $why

        [System.IO.File]::Delete($path)
        $claim = Enter-TestbedLeaseClaim -VMName $vm -TimeoutSeconds 5
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'maintenance' -CallerScript '' -ClaimTimeoutSeconds 1 }
        Check 'by hand, the read and the write wait for the lease''s claim' $true ($why.Contains("the lease claim on $vm stayed held"))
        $why = Get-Refusal { Set-Lease -VMName $vm -Minutes 30 -Reason 'maintenance' -CallerScript 'New-TestbedClone.ps1' }
        Check 'a script does not - its caller holds the claim around the call, as the runners do' '' $why
        $why = Get-Refusal { Remove-Lease -VMName $vm -Reason 'maintenance' -CallerScript '' -ClaimTimeoutSeconds 1 }
        Check 'a release by hand waits for the claim too' $true ($why.Contains("the lease claim on $vm stayed held"))
        $claim.Dispose()
        $claim = $null

        Write-Host ''
        Write-Host '== a release, on the scratch directory =='
        $out = Remove-Lease -VMName $vm -Reason 'maintenance' -CallerScript 'New-TestbedClone.ps1'
        Check 'the holder releases its lease with the -Reason it took it with' "Released the lease on $vm.|False" @([string]$out, (Test-Path -LiteralPath $path))
        Check 'a release of no lease says so' "No lease on $vm to release." (Remove-Lease -VMName $vm -Reason '' -CallerScript '')
        Write-LeaseFile ([pscustomobject]@{ vmName = $vm; takenUnix = $now; expiresUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() + 1200; holderPid = $gone.Id; holderStartUnix = $goneStart; holderHost = $env:COMPUTERNAME; holderScript = 'Invoke-LiveTierOnGuest.ps1'; byHand = $false; reason = 'live-tier run G' })
        $out = Remove-Lease -VMName $vm -Reason '' -CallerScript ''
        Check 'a lease whose process has ended is released by a person who gives no -Reason' "Released the lease on $vm. It released the lease of 'live-tier run G', whose process $($gone.Id) has ended.|False" @([string]$out, (Test-Path -LiteralPath $path))
        Write-LeaseFile ([pscustomobject]@{ vmName = $vm; takenUnix = $now; expiresUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - 5; holderPid = $PID; holderStartUnix = $ownStart; holderHost = $env:COMPUTERNAME; holderScript = ''; byHand = $true; reason = 'maintenance' })
        $out = Remove-Lease -VMName $vm -Reason 'build-vm run S' -CallerScript 'Invoke-TestsOnBuildVm.ps1'
        Check 'and an expired one by anyone' "Released the lease on $vm.|False" @([string]$out, (Test-Path -LiteralPath $path))

        Write-Host ''
        Write-Host '== the command and its callers =='
        $sets = @((Get-Command -Name $PSCommandPath).ParameterSets | ForEach-Object { $_.Name } | Sort-Object)
        Check 'it takes a lease, releases one, or tests itself' 'Release|SelfTest|Take' $sets
        $needVm = @((Get-Command -Name $PSCommandPath).ParameterSets | Where-Object { @($_.Parameters | Where-Object { $_.Name -eq 'VMName' -and $_.IsMandatory }).Count -gt 0 } | ForEach-Object { $_.Name } | Sort-Object)
        Check 'and -VMName is mandatory for both of the first two' 'Release|Take' $needVm
        $withReason = @((Get-Command -Name $PSCommandPath).ParameterSets | Where-Object { @($_.Parameters | Where-Object { $_.Name -eq 'Reason' }).Count -gt 0 } | ForEach-Object { $_.Name } | Sort-Object)
        Check 'both take a -Reason' 'Release|Take' $withReason
        $repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $releases = @()
        foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $repo 'Testbed\host') -Filter '*.ps1' -File)) {
            $fileAst = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
            foreach ($c in @($fileAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))) {
                if ($c.CommandElements.Count -eq 0 -or -not $c.CommandElements[0].Extent.Text.Contains('Set-TestbedLease.ps1')) { continue }
                $names = @($c.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] } | ForEach-Object { $_.ParameterName })
                if ($names -contains 'Release') { $releases += [pscustomobject]@{ Where = "$($f.Name):$($c.Extent.StartLineNumber)"; Reason = ($names -contains 'Reason') } }
            }
        }
        Check 'it finds the releases the runners, the probe and the clone maker make' $true ($releases.Count -ge 5)
        Check 'and every one gives the -Reason its lease was taken with' '' (@($releases | Where-Object { -not $_.Reason } | ForEach-Object { $_.Where }) -join ', ')
        $readme = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\README.md'))
        Check 'Testbed/README.md says a live lease is never taken or released from under its holder' $true ($readme.Contains('never takes or releases a live lease that its holder still has'))
    }
    finally {
        if ($null -ne $claim) { $claim.Dispose() }
        foreach ($p in @($live, $gone)) { if ($null -ne $p) { try { if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force } } catch { } } }
        $script:TestbedLeaseDirectoryOverride = ''
        try { Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }

    Write-Host ''
    Write-Host "$($script:stChecks) check(s), $($script:stFailures.Count) failure(s)."
    Write-Host 'NOT COVERED HERE: the idle-saver and the runners reading these leases - their own -SelfTest and real runs cover them.'
    if ($script:stFailures.Count -gt 0) { return 1 }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# The VM name is operator-supplied and becomes a filename; keep it to a safe shape rather than
# trusting it.
if ($VMName -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    throw "VM name '$VMName' is not a plain name; refusing to build a lease path from it."
}

if ($Release) {
    Write-Output (Remove-Lease -VMName $VMName -Reason $Reason -CallerScript $CallerScript)
    return
}

if ($Minutes -lt 1 -or $Minutes -gt 480) {
    throw "A lease of $Minutes minutes is outside the sane range (1-480). A very long lease pins the VM awake if its holder dies."
}

Write-Output (Set-Lease -VMName $VMName -Minutes $Minutes -Reason $Reason -CallerScript $CallerScript)
