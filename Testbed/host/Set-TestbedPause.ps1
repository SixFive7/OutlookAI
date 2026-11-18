#Requires -Version 5.1
<#
.SYNOPSIS
    The testbed's PAUSE SWITCH (Q195 (e), decided by the maintainer 2026-11-18). -On: no new test run
    starts on this host until -Off; runs already going finish. With neither: whether it is on, the
    host's load, which runs wait and which VMs run.

.DESCRIPTION
    For when the maintainer needs the PC - for gaming, for example. Turning the switch on stops
    nothing: a build run that holds its VM, a live run that holds its guest and a clone being made all
    finish as they would have. It holds back every run that has not started yet -
    Invoke-TestsOnBuildVm.ps1 before it takes a build VM, Invoke-LiveTierOnGuest.ps1 before it builds
    and again before it takes a guest, New-TestbedClone.ps1 -Execute before it takes the source. Each
    waits holding no VM's lock or lease, prints that the testbed is paused and since when, and once the
    switch is off starts as the load limit allows (TestbedLoad.ps1): one after another, not all at once.
    A run waits up to its own time limit (-QueueTimeoutMinutes, -LeaseWaitMinutes) and then gives up as
    INFRA, not tested.

    THE SWITCH IS THE MAINTAINER'S. An agent turns it on or off only when he asks for that in the
    current message (AGENTS.md, "Test VMs stay saved unless in use").

    It is a file, %SystemDrive%\OutlookAI-Testbed\gate\paused.json, machine-wide like the leases, so
    every checkout's runner reads the same switch. It records when it was turned on and the -Reason
    given, not who turned it on. A file there that cannot be read counts as on; -Off deletes it.

    No elevation is needed. The status lists the running testbed VMs when Hyper-V answers (local
    Hyper-V Administrators membership), and says so when it does not.

    EXIT CODES: 0 done; 1 the switch could not be read or written; -SelfTest 1 when a check failed.

.PARAMETER On
    Turn the switch on: no new run starts. Turned on again, it keeps the moment it was first turned on.

.PARAMETER Off
    Turn it off: the runs that wait start as the host's load allows.

.PARAMETER Reason
    With -On: a few words every waiting run prints, such as 'gaming'.

.PARAMETER SelfTest
    TestbedLoad.ps1's decisions, its files on a scratch directory, one real reading of this machine's
    processor times, the gate's wait and start, and that AGENTS.md and Testbed/README.md name this
    command. No Hyper-V. Runs on the build VM with every other -SelfTest.

.EXAMPLE
    pwsh -File Testbed/host/Set-TestbedPause.ps1 -On -Reason 'gaming'
    pwsh -File Testbed/host/Set-TestbedPause.ps1 -Off
    pwsh -File Testbed/host/Set-TestbedPause.ps1          # is it on, and what waits and runs
#>
[CmdletBinding(DefaultParameterSetName = 'Status')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'On')] [switch] $On,
    [Parameter(ParameterSetName = 'On')] [string] $Reason = '',
    [Parameter(Mandatory = $true, ParameterSetName = 'Off')] [switch] $Off,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'TestbedLoad.ps1')
. (Join-Path $PSScriptRoot 'TestbedPool.ps1')

$Command = 'pwsh -File Testbed/host/Set-TestbedPause.ps1'

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
    Write-Host "Set-TestbedPause.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No Hyper-V; the gate's files on a scratch directory."

    Write-Host ''
    Write-Host '== the decision''s numbers (Q195 (e)) =='
    Check 'a run starts only under 70 % busy' 70 $TestbedLoadLimitPercent
    Check 'measured over the last 30 s' 30 $TestbedLoadWindowSeconds
    Check 'a VM taken counts at its full size for 120 s' 120 $TestbedStartSettleSeconds
    Check 'the gate''s claim is on a name no VM has' '' (Get-TestbedPoolOf $TestbedGateClaimName)

    Write-Host ''
    Write-Host '== the busy share and the window =='
    $a = [pscustomobject]@{ When = [DateTime]::UtcNow; Idle = 1000L; Total = 2000L }
    $b = [pscustomobject]@{ When = [DateTime]::UtcNow; Idle = 1300L; Total = 3000L }
    Check 'idle 300 of 1000 is 70 % busy' 70 (Get-BusyPercent $a $b)
    Check 'no time passed is no answer' $true ($null -eq (Get-BusyPercent $a $a))
    Check 'nor is a reading missing' $true ($null -eq (Get-BusyPercent $null $b))
    Check 'it never reads above 100' 100 (Get-BusyPercent $a ([pscustomobject]@{ When = [DateTime]::UtcNow; Idle = 900L; Total = 3000L }))
    $now = [DateTime]::new(2026, 11, 18, 12, 0, 0, [DateTimeKind]::Utc)
    $r = @(65, 40, 31, 29, 5) | ForEach-Object { [pscustomobject]@{ When = $now.AddSeconds(-$_); Idle = 0L; Total = 0L; Age = $_ } }
    Check 'a look starts from the newest reading at least a window old' 31 (Select-LoadBaseline -Readings $r -Now $now -WindowSeconds 30).Age
    Check 'none that old: a fresh window is waited out' $true ($null -eq (Select-LoadBaseline -Readings @($r[3], $r[4]) -Now $now -WindowSeconds 30))
    Check 'one more than two windows old is too old - a run''s own build lies between' $true ($null -eq (Select-LoadBaseline -Readings @($r[0], $r[4]) -Now $now -WindowSeconds 30))
    Check 'no readings at all: a fresh window' $true ($null -eq (Select-LoadBaseline -Readings @() -Now $now -WindowSeconds 30))

    Write-Host ''
    Write-Host '== the VMs other runs took lately =='
    Check 'a build VM is an eighth of a 32-processor host' 12.5 (Get-VmSharePercent 4 32)
    Check 'an Outlook guest 6 of 32' 18.75 (Get-VmSharePercent 6 32)
    Check 'a VM whose size is not known counts as one processor' 3.125 (Get-VmSharePercent 0 32)
    $t = 1800000000L
    $starts = @(
        [pscustomobject]@{ unix = $t - 10; share = 12.5; vm = 'OutlookAI-Build2'; run = 'r1' },
        [pscustomobject]@{ unix = $t - 119; share = 18.75; vm = 'OutlookAI-Indexed'; run = 'r2' },
        [pscustomobject]@{ unix = $t - 120; share = 12.5; vm = 'OutlookAI-Build3'; run = 'r3' },
        [pscustomobject]@{ unix = 'soon'; share = 12.5; vm = 'x'; run = 'r4' },
        $null,
        [pscustomobject]@{ unix = $t - 5; share = 12.5; vm = 'OutlookAI-Build'; run = 'mine' }
    )
    $p = Get-PendingStarts -Starts $starts -NowUnix $t -SettleSeconds 120 -ExceptRun 'mine'
    Check 'the starts of the last 120 s count, an expired, unreadable or own one does not' '31.25|2' "$($p.Percent)|$($p.Count)"
    Check 'none is nothing' '0|0' ("{0}|{1}" -f (Get-PendingStarts -Starts @() -NowUnix $t).Percent, (Get-PendingStarts -Starts @() -NowUnix $t).Count)

    Write-Host ''
    Write-Host '== the decision =='
    Check 'under the limit, switch off: go' '' (Get-TestbedGateVerdict -Pause $null -BusyPercent 69.9 -WindowSeconds 30)
    $why = Get-TestbedGateVerdict -Pause $null -BusyPercent 70 -WindowSeconds 30
    Check 'at the limit: wait, saying the load, the window and the limit' "the host's processors were 70 % busy over the last 30 s, not under the 70 % limit" $why
    Check 'a share just under the limit never prints as the limit' '69 %' (Format-Percent 69.96)
    $why = Get-TestbedGateVerdict -Pause $null -BusyPercent 50 -WindowSeconds 30 -PendingPercent 25 -PendingCount 2
    Check 'the VMs taken lately push it over: wait, naming them' $true ($why.Contains('50 % busy') -and $why.Contains('2 VM(s)') -and $why.Contains('25 % more') -and $why.Contains('75 %, not under the 70 % limit'))
    Check 'and under it with them, go' '' (Get-TestbedGateVerdict -Pause $null -BusyPercent 50 -WindowSeconds 30 -PendingPercent 12.5 -PendingCount 1)
    Check 'a load that could not be read is no reason to wait' '' (Get-TestbedGateVerdict -Pause $null -BusyPercent $null -WindowSeconds 30 -PendingPercent 90 -PendingCount 5)
    $pauseRecord = [pscustomobject]@{ paused = $true; sinceUnix = 1795000000L; reason = 'gaming' }
    $why = Get-TestbedGateVerdict -Pause $pauseRecord -BusyPercent 1 -WindowSeconds 30
    Check 'the switch on outranks a quiet host, and says since when, why, and how to turn it off' $true ($why.StartsWith('the testbed is paused, since ') -and $why.Contains("('gaming')") -and $why.EndsWith('Set-TestbedPause.ps1 -Off'))
    $why = Get-TestbedGateVerdict -Pause ([pscustomobject]@{ unreadable = 'bad json'; path = 'X:\gate\paused.json'; sinceUnix = 0; reason = '' }) -BusyPercent 1 -WindowSeconds 30
    Check 'an unreadable switch counts as ON, naming its file' $true ($why.Contains('X:\gate\paused.json') -and $why.Contains('counts as ON'))
    $since = [DateTime]::new(2026, 11, 18, 20, 15, 3)
    Check 'the first line of a wait says what for' 'waiting: x' (Format-TestbedWaitLine -Since $since -Why 'x' -First $true -Now $since)
    Check 'every later one since when' '  still waiting, since 20:15:03 (12 min): x' (Format-TestbedWaitLine -Since $since -Why 'x' -First $false -Now $since.AddMinutes(12.5))
    $open = Format-TestbedGateOpen -Look ([pscustomobject]@{ BusyPercent = 41.7; Seconds = 30.2; Note = '' }) -PendingPercent 12.5 -PendingCount 1
    Check 'a run let through says what it went on' "open - the pause switch is off, and the host's processors were 41 % busy over the last 30 s, 12 % more for 1 VM(s) other runs took in the last 120 s, under the 70 % limit" $open

    Write-Host ''
    Write-Host '== the switch, the starts and the waiting runs, on a scratch directory =='
    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-gate-selftest-' + [guid]::NewGuid().ToString('N'))
    $script:TestbedGateDirectoryOverride = $scratch
    try {
        Check 'the gate lives where it is pointed' $scratch (Get-TestbedGateDirectory)
        Check 'no file: the switch is off' $true ($null -eq (Get-TestbedPause))
        $on = Set-TestbedPauseOn -Reason 'gaming'
        Check 'on: read back with its reason' 'True|gaming' "$($on.paused)|$($on.reason)"
        Check 'and since now' $true ([Math]::Abs([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - [long]$on.sinceUnix) -le 5)
        $sinceBefore = [long]$on.sinceUnix
        [System.IO.File]::WriteAllText((Join-Path $scratch 'paused.json'), (([ordered]@{ paused = $true; sinceUnix = $sinceBefore - 600; sinceUtc = ''; reason = 'gaming' }) | ConvertTo-Json))
        $again = Set-TestbedPauseOn
        Check 'turned on again: the first moment and the reason are kept' "$($sinceBefore - 600)|gaming" "$($again.sinceUnix)|$($again.reason)"
        Check 'off: it was on' $true (Set-TestbedPauseOff)
        Check 'and is off now' $true ($null -eq (Get-TestbedPause))
        Check 'off again: it was not on' $false (Set-TestbedPauseOff)
        [System.IO.File]::WriteAllText((Join-Path $scratch 'paused.json'), '{ "paused": tr')
        $bad = Get-TestbedPause
        Check 'a half-written switch is read as ON, with why' $true ($null -ne $bad -and [string]$bad.unreadable -ne '')
        [void](Set-TestbedPauseOff)
        Check 'and -Off deletes it' $true ($null -eq (Get-TestbedPause))

        Check 'no starts file: no starts' 0 @(Read-TestbedStarts).Count
        Write-TestbedStarts @([pscustomobject]@{ unix = 5; share = 12.5; vm = 'OutlookAI-Build'; run = 'a' })
        Check 'one start is written as an array of one and read back as one' $true (([System.IO.File]::ReadAllText((Join-Path $scratch 'starts.json'))).TrimStart().StartsWith('[') -and @(Read-TestbedStarts).Count -eq 1)
        Write-TestbedStarts @()
        # Without its white space: Windows PowerShell 5.1 writes an empty array as '[', two line breaks
        # and ']', where PowerShell 7 writes '[]' (build VM run 20261118-081540-450d0af33ff9).
        Check 'none is written as an empty array' '[]' (([System.IO.File]::ReadAllText((Join-Path $scratch 'starts.json'))) -replace '\s', '')
        [System.IO.File]::WriteAllText((Join-Path $scratch 'starts.json'), 'not json')
        Check 'an unreadable starts file is no starts' 0 @(Read-TestbedStarts).Count
        Write-TestbedStarts @()

        Set-TestbedWaiting -What 'build run r1' -Why 'the testbed is paused' -Since ([DateTime]::Now.AddMinutes(-3))
        $w = @(Get-TestbedWaiting)
        Check 'a waiting run is listed: what, why and since when' "1|build run r1|the testbed is paused|$PID" ("{0}|{1}|{2}|{3}" -f $w.Count, $w[0].what, $w[0].why, $w[0].pid)
        Remove-TestbedWaiting
        Check 'and is gone when its wait ends' 0 @(Get-TestbedWaiting).Count
        $stale = [ordered]@{ pid = $PID; processStartUnix = 1; what = 'a run that was stopped'; why = 'x'; sinceUnix = 1; updatedUnix = 1 }
        [System.IO.File]::WriteAllText((Join-Path (Join-Path $scratch 'waiting') '424242.json'), ($stale | ConvertTo-Json))
        Check 'one whose process is gone - or is another process now - is dropped, not listed' '0|False' ("{0}|{1}" -f @(Get-TestbedWaiting).Count, (Test-Path -LiteralPath (Join-Path (Join-Path $scratch 'waiting') '424242.json')))
        # A record that goes between the list and the read is passed over, not a throw (2026-11-18).
        Set-TestbedWaiting -What 'a self-test run' -Why 'the switch' -Since (Get-Date)
        $waitingDir = Join-Path (Get-TestbedGateDirectory) 'waiting'
        $listed = @(Get-TestbedStateFiles $waitingDir)
        Remove-TestbedWaiting
        $threw = ''
        $read = @()
        try { $read = @(Read-TestbedLiveRecords -Paths $listed) } catch { $threw = $_.Exception.Message }
        Check 'a waiting record deleted between the list and the read is passed over, and nothing throws' '1|0|' ("{0}|{1}|{2}" -f $listed.Count, $read.Count, $threw)

        Write-Host ''
        Write-Host '== this machine''s processor times, read for real =='
        $reading = Get-HostProcessorTimes
        Check 'GetSystemTimes answers, idle within total' $true ($null -ne $reading -and $reading.Total -gt 0 -and $reading.Idle -ge 0 -and $reading.Idle -le $reading.Total)
        if ($null -eq $reading) { Write-Host "       ($TestbedLoadReadError)" }
        $m = Measure-HostBusy -WindowSeconds 1
        Check 'a first look waits out its window and reads 0 to 100 %' $true ($null -ne $m -and $m.Seconds -ge 0.9 -and $m.Percent -ge 0 -and $m.Percent -le 100)
        $t0 = Get-Date
        $m2 = Measure-HostBusy -WindowSeconds 1
        Check 'the next look at once, from the readings kept' $true ($null -ne $m2 -and ((Get-Date) - $t0).TotalSeconds -lt 0.9 -and $m2.Seconds -ge 1)

        Write-Host ''
        Write-Host '== the gate, end to end on the scratch directory =='
        $look = Test-TestbedGate -LimitPercent 101 -WindowSeconds 1
        Check 'switch off, any load under 101 %: go' 'True|' "$($look.Go)|$($look.Why)"
        $look = Test-TestbedGate -LimitPercent 0 -WindowSeconds 1
        Check 'under a 0 % limit: wait, for the load' $true ((-not $look.Go) -and $look.Why.Contains('not under the 0 % limit'))
        [void](Set-TestbedPauseOn -Reason 'self-test')
        $t0 = Get-Date
        $look = Test-TestbedGate -LimitPercent 101 -WindowSeconds 1
        Check 'switch on: wait, for the switch, without measuring' $true ((-not $look.Go) -and $look.Paused -and $look.Why.Contains("('self-test')") -and ((Get-Date) - $t0).TotalSeconds -lt 0.9)
        $said = New-Object System.Collections.Generic.List[string]
        $waited = Wait-TestbedGate -Run 'r-wait' -What 'a self-test wait' -TimeoutMinutes 0 -Say { param($line) $said.Add($line) } -LimitPercent 101 -WindowSeconds 1 -PollSeconds 1
        Check 'a wait that runs out says no, and why' $true ((-not $waited.Go) -and $waited.Why.Contains('paused'))
        Check 'having said once what it waited for' $true ($said.Count -eq 1 -and $said[0].StartsWith('waiting: the testbed is paused'))
        Check 'and no longer listed as waiting' 0 @(Get-TestbedWaiting).Count
        [void](Set-TestbedPauseOff)
        $waited = Wait-TestbedGate -Run 'r-wait' -What 'a self-test wait' -TimeoutMinutes 0 -Say { param($line) $said.Add($line) } -LimitPercent 101 -WindowSeconds 1 -PollSeconds 1
        Check 'switch off: the wait ends at once' 'True|0' ("{0}|{1}" -f $waited.Go, [int]$waited.WaitedSeconds)
        $gateText = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'TestbedLoad.ps1'))
        $waitAt = $gateText.IndexOf('function Wait-TestbedGate')
        $waitBody = ''
        if ($waitAt -ge 0) { $waitBody = $gateText.Substring($waitAt) }
        Check 'between looks a waiting run sleeps on the testbed''s files, 30 s at most - never a poll of its own' $true ($waitBody.Contains('[void](Wait-TestbedChange -Seconds $PollSeconds)') -and $waitBody.Contains('[int] $PollSeconds = $TestbedWaitSeconds') -and -not $waitBody.Contains('Start-Sleep') -and $TestbedWaitSeconds -eq 30)

        # The start, under the gate's claim: TestbedLeasePath.ps1's, in the machine's leases directory -
        # on the build VM, which every run restores.
        $quiet = [pscustomobject]@{ Go = $true; BusyPercent = 60.0; Seconds = 30.0; Note = '' }
        $first = Enter-TestbedStart -Look $quiet -Run 'r-first'
        try {
            Check 'a quiet host: the run may start, and holds the claim' 'True|True' ("{0}|{1}" -f $first.Go, ($null -ne $first.Claim))
            Register-TestbedStart -Start $first -Run 'r-first' -VmName 'OutlookAI-Build2' -VmProcessors 4
        }
        finally { Exit-TestbedStart $first }
        Check 'its start is written, at the VM''s share of this machine' ('OutlookAI-Build2|' + [Math]::Round((Get-VmSharePercent 4 ([Environment]::ProcessorCount)), 2)) "$(@(Read-TestbedStarts)[0].vm)|$(@(Read-TestbedStarts)[0].share)"
        Check 'and the claim is let go' $true ($null -eq $first.Claim)
        $share = [Math]::Round((Get-VmSharePercent 4 ([Environment]::ProcessorCount)), 2)
        $busyLook = [pscustomobject]@{ Go = $true; BusyPercent = (70.5 - $share); Seconds = 30.0; Note = '' }
        $second = Enter-TestbedStart -Look $busyLook -Run 'r-second'
        Check 'the next run counts it, and is told to wait - holding no claim' $true ((-not $second.Go) -and $null -eq $second.Claim -and $second.Why.Contains('1 VM(s) other runs took'))
        $own = Enter-TestbedStart -Look $busyLook -Run 'r-first'
        try { Check 'a run does not count its own start' 'True|0' ("{0}|{1}" -f $own.Go, $own.PendingCount) }
        finally { Exit-TestbedStart $own }
        [void](Set-TestbedPauseOn)
        $paused = Enter-TestbedStart -Look $quiet -Run 'r-third'
        Check 'the switch turned on since the look still stops the start' $true ((-not $paused.Go) -and $paused.Why.Contains('paused'))
        [void](Set-TestbedPauseOff)
    }
    finally {
        $script:TestbedGateDirectoryOverride = ''
        Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ''
    Write-Host '== the documents name this command =='
    $repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $agents = [System.IO.File]::ReadAllText((Join-Path $repo 'AGENTS.md'))
    $sectionAt = $agents.IndexOf('## Test VMs stay saved unless in use')
    $section = ''
    if ($sectionAt -ge 0) { $section = $agents.Substring($sectionAt); $next = $section.IndexOf("`n## ", 5); if ($next -gt 0) { $section = $section.Substring(0, $next) } }
    Check 'AGENTS.md''s "Test VMs stay saved unless in use" gives the switch, both ways' $true ($section.Contains("$Command -On") -and $section.Contains("$Command -Off"))
    Check 'and the load limit' $true ($section.Contains('70 %'))
    $readme = [System.IO.File]::ReadAllText((Join-Path $repo 'Testbed\README.md'))
    Check 'Testbed/README.md gives it too' $true ($readme.Contains("$Command -On") -and $readme.Contains("$Command -Off"))

    Write-Host ''
    Write-Host "$($script:stChecks) check(s), $($script:stFailures.Count) failure(s)."
    Write-Host 'NOT COVERED HERE: the runners waiting at the gate, which needs Hyper-V - their own -SelfTest pins where they look, and a real run covers the rest.'
    if ($script:stFailures.Count -gt 0) { return 1 }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE SWITCH, AND WHAT IT HOLDS.
# =============================================================================================

function Format-Unix([object] $Unix, [string] $Format = 'yyyy-MM-dd HH:mm') {
    try { if ([long]$Unix -gt 0) { return [DateTimeOffset]::FromUnixTimeSeconds([long]$Unix).LocalDateTime.ToString($Format) } } catch { }
    return 'an unknown moment'
}

function Write-Waiting {
    $waiting = @(Get-TestbedWaiting)
    if ($waiting.Count -eq 0) { Write-Host 'Waiting       none'; return }
    Write-Host "Waiting       $($waiting.Count) run(s), the longest-waiting first:"
    foreach ($w in $waiting) {
        $minutes = [int][Math]::Floor(([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - [long]$w.sinceUnix) / 60)
        Write-Host ("              since {0} ({1} min), pid {2}: {3}" -f (Format-Unix $w.sinceUnix 'HH:mm:ss'), $minutes, $w.pid, $w.what)
        Write-Host "                for $($w.why)"
    }
}

function Write-Running {
    $vms = @()
    try { $vms = @(Get-VM -ErrorAction Stop | Where-Object { ([string]$_.State) -eq 'Running' }) }
    catch { Write-Host "Running       not known - Hyper-V did not answer: $($_.Exception.Message)"; return }
    $testbed = @(Get-TestbedPoolMembers) + @('OutlookAI-Exchange')
    $ours = @($vms | Where-Object { $testbed -contains $_.Name })
    if ($ours.Count -eq 0) { Write-Host 'Running       no testbed VM'; return }
    Write-Host "Running       $($ours.Count) testbed VM(s), each to finish its run:"
    foreach ($vm in $ours) {
        $lease = Get-TestbedLease -VMName $vm.Name
        $leaseText = 'no live lease - the idle-saver saves it'
        if ($null -ne $lease) { $leaseText = "leased until $(Format-Unix $lease.expiresUnix 'HH:mm') - $($lease.reason)" }
        Write-Host "              $($vm.Name), $leaseText"
    }
}

function Write-Recent {
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $recent = @(@(Read-TestbedStarts) | Where-Object { try { ($now - [long]$_.unix) -lt $TestbedStartSettleSeconds } catch { $false } })
    if ($recent.Count -eq 0) { return }
    foreach ($s in $recent) {
        Write-Host ("Taken lately  {0} at {1}, counted as {2} more until {3} - {4}" -f $s.vm, (Format-Unix $s.unix 'HH:mm:ss'), (Format-Percent ([double]$s.share)), (Format-Unix ([long]$s.unix + $TestbedStartSettleSeconds) 'HH:mm:ss'), $s.run)
    }
}

$limitText = "a run starts while the host's processors are under $(Format-Percent $TestbedLoadLimitPercent) busy over the last $TestbedLoadWindowSeconds s (Testbed/host/TestbedLoad.ps1)"
try {
    if ($On) {
        $record = Set-TestbedPauseOn -Reason $Reason
        $why = ''
        if ([string]$record.reason) { $why = " ('$($record.reason)')" }
        Write-Host "Pause switch  ON since $(Format-Unix $record.sinceUnix)$why. No new test run starts - build runs, live runs, clones -"
        Write-Host "              until: $Command -Off"
        Write-Host '              Runs already going finish: nothing is stopped, saved or killed.'
        Write-Running
        Write-Waiting
        exit 0
    }
    if ($Off) {
        $was = Get-TestbedPause
        $wasOn = Set-TestbedPauseOff
        if ($wasOn) {
            $sinceText = 'an unknown moment'
            if ($null -ne $was -and -not $was.unreadable) { $sinceText = Format-Unix $was.sinceUnix }
            Write-Host "Pause switch  OFF (it was on since $sinceText). Waiting runs start one after another as the load allows:"
        }
        else { Write-Host 'Pause switch  OFF - it was not on. Waiting runs start one after another as the load allows:' }
        Write-Host "              $limitText."
        Write-Waiting
        exit 0
    }
    # The status: the switch, the load now, the VMs taken lately, the runs that wait, the VMs that run.
    $pause = Get-TestbedPause
    if ($null -eq $pause) { Write-Host "Pause switch  OFF. To hold every new test run back: $Command -On" }
    else { Write-Host "Pause switch  ON - $(Format-TestbedPause $pause)" }
    $m = Measure-HostBusy -WindowSeconds 5
    if ($null -eq $m) { Write-Host "Host load     not known: $TestbedLoadReadError" }
    else { Write-Host ("Host load     {0} busy over the last {1:N0} s; {2}" -f (Format-Percent $m.Percent), $m.Seconds, $limitText) }
    Write-Recent
    Write-Waiting
    Write-Running
    exit 0
}
catch {
    Write-Host "FAILED: $($_.Exception.Message)"
    exit 1
}
