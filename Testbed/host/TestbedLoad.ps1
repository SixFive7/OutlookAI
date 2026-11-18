#Requires -Version 5.1
<#
.SYNOPSIS
    The testbed's LOAD LIMIT and PAUSE SWITCH (Q195 (e), decided by the maintainer 2026-11-18): whether
    a test run may start now. Dot-sourced by both runners, New-TestbedClone.ps1 and
    Set-TestbedPause.ps1, so that they cannot disagree.

.DESCRIPTION
    WHY. Every runner started its VM the moment one was free. With every pool member busy, the host's
    32 logical processors read 82 to 100 % busy (Docs/live-tier-on-the-vm.md section 4.7): the
    maintainer's own work slowed down, and every run took about twice as long.

    THE LOAD LIMIT. A run starts only while the host's processors are less than 70 % busy over the
    last 30 seconds. Otherwise it waits, says what for and since when, and holds no VM's lock or lease
    while it waits. A run starts:
      - in Invoke-TestsOnBuildVm.ps1, when it takes a build VM - looked at on every turn of its queue;
      - in Invoke-LiveTierOnGuest.ps1, before it builds the commit's payloads on this host, and again
        when it takes a guest - on every turn of its lease wait;
      - in New-TestbedClone.ps1 -Execute, before it takes the source to export it.
    A VM another run took in the last 120 seconds counts at its full size - its virtual processors'
    share of the host's - on top of what was measured, because that long passes before a new run's load
    shows in a 30-second measurement. Without it, every run that waited would start in the second the
    load dipped, and the host would be back at 100 %. Taking a VM - reading those starts, deciding, and
    writing its own - is one step, under TestbedLeasePath.ps1's Enter-TestbedLeaseClaim on a name no VM
    has, held for that moment only.

    THE PAUSE SWITCH. `Set-TestbedPause.ps1 -On`: no new run starts until `Set-TestbedPause.ps1 -Off`.
    Runs already going finish - nothing is stopped, saved or killed. The janitor and the idle-saver go
    on as before: they put VMs away and start nothing.

    WHAT IS MEASURED. The share of time the host's logical processors were not idle, from kernel32's
    GetSystemTimes. It read 62.0 % where \Processor(_Total)\% Processor Time read 62.1 % over one
    11-second window on 2026-11-18, and it needs no counter name in the machine's language. On a Hyper-V
    host it includes the VMs only under the hypervisor's ROOT scheduler, which this host runs (the System
    log's Hyper-V-Hypervisor event 2, "scheduler type is 0x4", client Windows' default): their virtual
    processors run as the host's own threads. Measured the same day over one 10-second window: 83.3 %
    here, and 85.3 % by the hypervisor's own logical-processor counter, of which the guests' virtual
    processors were 29.2 points and the host's 56.0 - the rest is the hypervisor itself. Under the core
    scheduler the VMs would not show here at all. Task Manager's figure is % Processor Utility, which
    weighs frequency boost in and read 81 % where this read 62 %: the limit is on time busy.

    WHY 30 SECONDS. At about 60 % busy, sampled every second for four minutes on 2026-11-18, 10-second
    averages swung between 40 and 80 % and were at or over 70 % in 7 looks of 37; 30-second ones stayed
    between 47 and 73 %. A run's first look takes those 30 seconds. A run that waits looks again on every
    turn, at the last 30 seconds of its own readings, without waiting them out again.

    FAILURE DIRECTIONS, each chosen so that the wrong answer is visible:
      - A pause file that is there and cannot be read counts as ON, and the run names the file: a switch
        half-written must not let runs in. `-Off` deletes it.
      - Processor times that cannot be read let the run start, as every run started before this
        existed, and the run says so: a broken measurement must not stop every agent's testing.
      - A record of recent starts that cannot be read counts as none.
      - The claim on the starts not had within 10 s reads as "another run is taking a VM just now": the
        run lets go of what it holds and waits one more turn.

    WHERE ITS STATE LIVES. Machine-wide, beside the leases, because every checkout's runner reads it:
    %SystemDrive%\OutlookAI-Testbed\gate\ - paused.json while the switch is on, starts.json with the VMs
    taken in the last 120 s, and waiting\<pid>.json for every run that waits, which Set-TestbedPause.ps1
    lists. Nothing in it is secret; the switch records when it was turned on and the reason given, not
    who turned it on.

    NOTHING RUNS IN THE BACKGROUND for any of it: no task, no service, no sampling process. A run
    measures in its own process, while it waits.
#>

. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
# The pools and the waits (TestbedPool.ps1): a run that waits at the gate sleeps as a queued run does,
# on the testbed's files (Wait-TestbedChange), never on a timer of its own.
. (Join-Path $PSScriptRoot 'TestbedPool.ps1')

# The decision's numbers. Set-TestbedPause.ps1 -SelfTest pins them.
$TestbedLoadLimitPercent = 70       # a run starts only while the host is less busy than this
$TestbedLoadWindowSeconds = 30      # over the last this many seconds
$TestbedStartSettleSeconds = 120    # a VM taken counts at its full size for this long
# The claim that makes reading the recent starts and writing one a single step: TestbedLeasePath.ps1's,
# on a name no VM has, so that it is no lease of anything.
$TestbedGateClaimName = 'OutlookAI-TestbedGate'
# Set-TestbedPause.ps1 -SelfTest points the gate at a scratch directory; empty, it is the machine's.
$TestbedGateDirectoryOverride = ''
# This process's readings of the host's processor times, oldest first, two windows at most.
$TestbedLoadReadings = New-Object System.Collections.Generic.List[object]
$TestbedLoadReadError = ''

# Where the gate keeps its state: gate\ beside the leases' directory.
function Get-TestbedGateDirectory {
    $dir = [string]$TestbedGateDirectoryOverride
    if (-not $dir) { $dir = Join-Path (Split-Path -Parent (Get-TestbedLeaseDirectory)) 'gate' }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    return $dir
}

# =============================================================================================
# PURE. Set-TestbedPause.ps1 -SelfTest pins every one.
# =============================================================================================

# The busy share between two readings of the processor times, in percent, from 0 to 100 - or $null
# when no processor time passed between them.
function Get-BusyPercent {
    param([object] $From, [object] $To)
    if ($null -eq $From -or $null -eq $To) { return $null }
    $total = [double]([long]$To.Total - [long]$From.Total)
    if ($total -le 0) { return $null }
    $busy = 100.0 * (1.0 - ([double]([long]$To.Idle - [long]$From.Idle) / $total))
    return [Math]::Max(0.0, [Math]::Min(100.0, $busy))
}

# The reading a look at the last $WindowSeconds starts from: the newest one at least that old. $null
# when there is none - or when that one is more than twice that old, so that a look never reaches back
# past a gap such as a run's own build, and a fresh window is waited out instead.
function Select-LoadBaseline {
    param([object[]] $Readings, [DateTime] $Now, [int] $WindowSeconds)
    $base = $null
    foreach ($r in @($Readings)) {
        if ($null -eq $r) { continue }
        if (($Now - [DateTime]$r.When).TotalSeconds -lt $WindowSeconds) { continue }
        if ($null -eq $base -or [DateTime]$r.When -gt [DateTime]$base.When) { $base = $r }
    }
    if ($null -ne $base -and ($Now - [DateTime]$base.When).TotalSeconds -gt 2 * $WindowSeconds) { return $null }
    return $base
}

# What the VMs other runs took in the last $SettleSeconds add to a measurement: their shares summed, and
# how many. A start that cannot be read, has expired, or is this run's own does not count.
function Get-PendingStarts {
    param([object[]] $Starts, [long] $NowUnix, [int] $SettleSeconds = $TestbedStartSettleSeconds, [string] $ExceptRun = '')
    $percent = 0.0
    $count = 0
    foreach ($s in @($Starts)) {
        if ($null -eq $s) { continue }
        $at = 0L
        $share = 0.0
        try { $at = [long]$s.unix; $share = [double]$s.share } catch { continue }
        $age = $NowUnix - $at
        if ($at -le 0 -or $share -le 0 -or $age -ge $SettleSeconds -or $age -lt -$SettleSeconds) { continue }
        if ($ExceptRun -and [string]$s.run -ceq $ExceptRun) { continue }
        $percent += $share
        $count++
    }
    return [pscustomobject]@{ Percent = $percent; Count = $count }
}

# A VM's share of the host, in percent: its virtual processors over the host's logical ones.
function Get-VmSharePercent {
    param([int] $VmProcessors, [int] $HostProcessors)
    if ($HostProcessors -lt 1) { return 0.0 }
    if ($VmProcessors -lt 1) { $VmProcessors = 1 }
    return [Math]::Min(100.0, 100.0 * $VmProcessors / $HostProcessors)
}

# A percentage as these lines print it: whole, rounded DOWN, so that 69.6 never reads as the 70 it is under.
function Format-Percent {
    param([double] $Value)
    return ('{0} %' -f [int][Math]::Floor($Value))
}

# The pause switch as a waiting run says it.
function Format-TestbedPause {
    param([object] $Pause)
    if ($null -eq $Pause) { return '' }
    if ($Pause.unreadable) { return "the pause switch's file $($Pause.path) cannot be read ($($Pause.unreadable)) - it counts as ON; Set-TestbedPause.ps1 -Off deletes it" }
    $since = 'a moment its file does not say'
    try { if ([long]$Pause.sinceUnix -gt 0) { $since = [DateTimeOffset]::FromUnixTimeSeconds([long]$Pause.sinceUnix).LocalDateTime.ToString('yyyy-MM-dd HH:mm') } } catch { }
    $text = "the testbed is paused, since $since"
    if ([string]$Pause.reason) { $text += " ('$($Pause.reason)')" }
    return "$text - no run starts until Set-TestbedPause.ps1 -Off"
}

# THE DECISION: '' when a run may start, or what it waits for, as one sentence. The switch first; then
# the load measured plus the VMs other runs took in the last two minutes, against the limit. A load
# that could not be measured is no reason to wait (the banner).
function Get-TestbedGateVerdict {
    param(
        [object] $Pause,
        [object] $BusyPercent,
        [int]    $WindowSeconds = $TestbedLoadWindowSeconds,
        [double] $PendingPercent = 0,
        [int]    $PendingCount = 0,
        [int]    $LimitPercent = $TestbedLoadLimitPercent,
        [int]    $SettleSeconds = $TestbedStartSettleSeconds
    )
    if ($null -ne $Pause) { return (Format-TestbedPause $Pause) }
    if ($null -eq $BusyPercent) { return '' }
    $busy = [double]$BusyPercent
    if ($busy + $PendingPercent -lt $LimitPercent) { return '' }
    if ($PendingCount -gt 0) {
        return ("the host's processors were {0} busy over the last {1} s, and {2} VM(s) other runs took in the last {3} s count {4} more until their load shows: {5}, not under the {6} limit" -f (Format-Percent $busy), $WindowSeconds, $PendingCount, $SettleSeconds, (Format-Percent $PendingPercent), (Format-Percent ($busy + $PendingPercent)), (Format-Percent $LimitPercent))
    }
    return ("the host's processors were {0} busy over the last {1} s, not under the {2} limit" -f (Format-Percent $busy), $WindowSeconds, (Format-Percent $LimitPercent))
}

# What a run says when the gate lets it start: the switch, the load it went on and the starts it counted.
function Format-TestbedGateOpen {
    param([object] $Look, [double] $PendingPercent = 0, [int] $PendingCount = 0, [int] $LimitPercent = $TestbedLoadLimitPercent, [int] $SettleSeconds = $TestbedStartSettleSeconds)
    if ($null -eq $Look -or $null -eq $Look.BusyPercent) {
        $note = "the host's processor times were not read"
        if ($null -ne $Look -and $Look.Note) { $note = $Look.Note }
        return "open - the pause switch is off; $note"
    }
    $text = "open - the pause switch is off, and the host's processors were {0} busy over the last {1:N0} s" -f (Format-Percent $Look.BusyPercent), $Look.Seconds
    if ($PendingCount -gt 0) { $text += (", {0} more for {1} VM(s) other runs took in the last {2} s" -f (Format-Percent $PendingPercent), $PendingCount, $SettleSeconds) }
    return ($text + (', under the {0} limit' -f (Format-Percent $LimitPercent)))
}

# A line about a wait: the first says what it waits for, and its time is when the wait began; every
# later one says since when, too.
function Format-TestbedWaitLine {
    param([DateTime] $Since, [string] $Why, [bool] $First, [DateTime] $Now = (Get-Date))
    if ($First) { return "waiting: $Why" }
    return ("  still waiting, since {0:HH:mm:ss} ({1} min): {2}" -f $Since, [int][Math]::Floor(($Now - $Since).TotalMinutes), $Why)
}

# =============================================================================================
# THE SWITCH, THE STARTS AND THE WAITING RUNS: files in the gate's directory.
# =============================================================================================

# The pause switch: $null while it is off, its record while it is on - and a record with .unreadable
# when the file is there and cannot be read, which counts as on (the banner).
function Get-TestbedPause {
    $path = Join-Path (Get-TestbedGateDirectory) 'paused.json'
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $record = [System.IO.File]::ReadAllText($path) | ConvertFrom-Json
        if ($null -eq $record) { throw 'the file is empty' }
        return $record
    }
    catch {
        # Deleted between the look and the read: the switch was turned off.
        if (-not (Test-Path -LiteralPath $path)) { return $null }
        return [pscustomobject]@{ unreadable = $_.Exception.Message; path = $path; sinceUnix = 0; reason = '' }
    }
}

# Turns the switch on. Turned on again, it keeps the moment it was first turned on, and the reason it
# had unless a new one is given. Returns the record as the runs will read it.
function Set-TestbedPauseOn {
    param([string] $Reason = '')
    $path = Join-Path (Get-TestbedGateDirectory) 'paused.json'
    $sinceUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $old = Get-TestbedPause
    if ($null -ne $old -and -not $old.unreadable) {
        try { if ([long]$old.sinceUnix -gt 0) { $sinceUnix = [long]$old.sinceUnix } } catch { }
        if (-not $Reason) { $Reason = [string]$old.reason }
    }
    $record = [ordered]@{
        paused    = $true
        sinceUnix = $sinceUnix
        sinceUtc  = [DateTimeOffset]::FromUnixTimeSeconds($sinceUnix).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')
        reason    = $Reason
    }
    [System.IO.File]::WriteAllText($path, ($record | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    return (Get-TestbedPause)
}

# Turns it off: $true when it was on.
function Set-TestbedPauseOff {
    $path = Join-Path (Get-TestbedGateDirectory) 'paused.json'
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    Remove-Item -LiteralPath $path -Force
    return $true
}

# The VMs taken lately, as starts.json holds them; none when it is missing or cannot be read.
function Read-TestbedStarts {
    $path = Join-Path (Get-TestbedGateDirectory) 'starts.json'
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $parsed = $null
    try { $parsed = [System.IO.File]::ReadAllText($path) | ConvertFrom-Json }
    catch { return @() }
    return @(@($parsed) | ForEach-Object { $_ } | Where-Object { $null -ne $_ })
}

# Writes the starts as one JSON array at every length: written through the pipeline, a lone start would
# be an object, and none would be nothing at all.
function Write-TestbedStarts {
    param([object[]] $Starts)
    $path = Join-Path (Get-TestbedGateDirectory) 'starts.json'
    $json = ConvertTo-Json -InputObject @(@($Starts) | Where-Object { $null -ne $_ }) -Depth 3
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

# A run that waits says so in waiting\<pid>.json - what it is, what it waits for, since when - each time
# it says so in its own output; the file goes when the wait ends. Never fails the run.
function Set-TestbedWaiting {
    param([string] $What, [string] $Why, [DateTime] $Since)
    try {
        $dir = Join-Path (Get-TestbedGateDirectory) 'waiting'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $record = [ordered]@{
            pid              = $PID
            processStartUnix = ([DateTimeOffset](Get-Process -Id $PID).StartTime).ToUnixTimeSeconds()
            what             = $What
            why              = $Why
            sinceUnix        = ([DateTimeOffset]$Since).ToUnixTimeSeconds()
            updatedUnix      = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        }
        [System.IO.File]::WriteAllText((Join-Path $dir "$PID.json"), ($record | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    }
    catch { }
}

function Remove-TestbedWaiting {
    try {
        $path = Join-Path (Join-Path (Get-TestbedGateDirectory) 'waiting') "$PID.json"
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
    catch { }
}

# The runs that wait now, the longest-waiting first. A record whose process has ended - a run stopped
# while it waited - is deleted, not listed.
function Get-TestbedWaiting {
    $dir = Join-Path (Get-TestbedGateDirectory) 'waiting'
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $alive = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $r = $null
        try { $r = [System.IO.File]::ReadAllText($f.FullName) | ConvertFrom-Json } catch { continue }
        if ($null -eq $r) { continue }
        $same = $false
        $p = $null
        try { $p = Get-Process -Id ([int]$r.pid) -ErrorAction SilentlyContinue } catch { }
        if ($null -ne $p) {
            # The same process, not another that has its id now.
            try { $same = (([DateTimeOffset]$p.StartTime).ToUnixTimeSeconds() -eq [long]$r.processStartUnix) } catch { $same = $true }
        }
        if (-not $same) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue; continue }
        $alive += $r
    }
    return @($alive | Sort-Object { [long]$_.sinceUnix })
}

# =============================================================================================
# THE MEASUREMENT.
# =============================================================================================

# One reading of the host's processor times: every logical processor's idle time and total time
# (kernel, which includes idle, plus user) since boot, in 100 ns units, from kernel32's GetSystemTimes -
# compiled once per process by Add-Type, which on Windows PowerShell 5.1 is the .NET Framework's own
# csc.exe (nothing installed: AGENTS.md, Dependencies). C# 5, which that compiler speaks. $null, with
# $TestbedLoadReadError set, when it cannot be read.
function Get-HostProcessorTimes {
    try {
        if (-not ('OutlookAI.Testbed.HostProcessorTimes' -as [type])) {
            # A compile that failed once fails again: not tried on every turn of a wait.
            if ($TestbedLoadReadError) { return $null }
            Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace OutlookAI.Testbed
{
    public static class HostProcessorTimes
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetSystemTimes(out long idleTime, out long kernelTime, out long userTime);

        // { idle, kernel + user }: every logical processor's, together, since boot, in 100 ns units.
        public static long[] Read()
        {
            long idle, kernel, user;
            if (!GetSystemTimes(out idle, out kernel, out user)) { throw new Win32Exception(Marshal.GetLastWin32Error()); }
            return new long[] { idle, kernel + user };
        }
    }
}
'@
        }
        $t = [OutlookAI.Testbed.HostProcessorTimes]::Read()
        return [pscustomobject]@{ When = [DateTime]::UtcNow; Idle = [long]$t[0]; Total = [long]$t[1] }
    }
    catch {
        $script:TestbedLoadReadError = $_.Exception.Message
        return $null
    }
}

# Takes a reading and keeps it, forgetting those more than two windows old: no look reaches back that far.
function Add-HostReading {
    param([int] $WindowSeconds = $TestbedLoadWindowSeconds)
    $now = Get-HostProcessorTimes
    if ($null -eq $now) { return $null }
    for ($i = $TestbedLoadReadings.Count - 1; $i -ge 0; $i--) {
        if (($now.When - [DateTime]$TestbedLoadReadings[$i].When).TotalSeconds -gt 2 * $WindowSeconds) { $TestbedLoadReadings.RemoveAt($i) }
    }
    $TestbedLoadReadings.Add($now)
    return $now
}

# The host's busy share over the last $WindowSeconds, from this process's readings: one taken now, and
# the newest at least a window old as where the look starts. With none that old - a run's first look,
# or its first after a gap such as its own build - the rest of a window is waited out first. $null when
# the processor times cannot be read; else Percent, and Seconds, the span looked at.
function Measure-HostBusy {
    param([int] $WindowSeconds = $TestbedLoadWindowSeconds)
    $earlier = $TestbedLoadReadings.ToArray()
    $last = Add-HostReading -WindowSeconds $WindowSeconds
    if ($null -eq $last) { return $null }
    $base = Select-LoadBaseline -Readings $earlier -Now $last.When -WindowSeconds $WindowSeconds
    if ($null -eq $base) {
        # The oldest reading kept starts the look, waited on until it is a window old.
        $base = $TestbedLoadReadings[0]
        $wait = $WindowSeconds - ($last.When - [DateTime]$base.When).TotalSeconds
        # 20 ms over, so that the clock's granularity cannot leave the reading a hair short of a window
        # old, and the next look starts from it.
        if ($wait -gt 0) { Start-Sleep -Milliseconds ([int][Math]::Ceiling($wait * 1000) + 20) }
        $last = Add-HostReading -WindowSeconds $WindowSeconds
        if ($null -eq $last) { return $null }
    }
    $percent = Get-BusyPercent $base $last
    if ($null -eq $percent) { return $null }
    return [pscustomobject]@{ Percent = $percent; Seconds = ($last.When - [DateTime]$base.When).TotalSeconds }
}

# =============================================================================================
# THE GATE.
# =============================================================================================

# One look, on a turn of a run's wait: the switch, then the host's load over the last 30 s. Holds
# nothing: the caller goes on to its VMs only when it says Go. Go, Why, Paused, BusyPercent, Seconds, Note.
function Test-TestbedGate {
    param([int] $LimitPercent = $TestbedLoadLimitPercent, [int] $WindowSeconds = $TestbedLoadWindowSeconds)
    $pause = Get-TestbedPause
    if ($null -ne $pause) {
        # Readings go on while the switch is on, so the last 30 s are known the moment it is turned off.
        [void](Add-HostReading -WindowSeconds $WindowSeconds)
        return [pscustomobject]@{ Go = $false; Why = (Format-TestbedPause $pause); Paused = $true; BusyPercent = $null; Seconds = 0; Note = '' }
    }
    $m = Measure-HostBusy -WindowSeconds $WindowSeconds
    if ($null -eq $m) {
        return [pscustomobject]@{ Go = $true; Why = ''; Paused = $false; BusyPercent = $null; Seconds = 0; Note = "the host's processor times could not be read ($TestbedLoadReadError) - the load limit is not applied to this run" }
    }
    $why = Get-TestbedGateVerdict -Pause $null -BusyPercent $m.Percent -WindowSeconds $WindowSeconds -LimitPercent $LimitPercent
    return [pscustomobject]@{ Go = (-not $why); Why = $why; Paused = $false; BusyPercent = $m.Percent; Seconds = $m.Seconds; Note = '' }
}

# THE STEP THAT STARTS A RUN, taken just before it takes its VM's lease - in the build runner with that
# VM's lock already held, for this moment only. Under the gate's claim: the switch read again, and the
# VMs other runs took in the last 120 s added to the load $Look measured. When the run may start, the
# claim stays held in .Claim, so that no other run decides before this one's start is written
# (Register-TestbedStart); Exit-TestbedStart lets it go, on every path. It never waits on the load or
# the switch: it answers, and a run told to wait lets go of what it holds and waits holding nothing.
function Enter-TestbedStart {
    param(
        [object] $Look,
        [string] $Run,
        [int] $LimitPercent = $TestbedLoadLimitPercent,
        [int] $WindowSeconds = $TestbedLoadWindowSeconds,
        [int] $SettleSeconds = $TestbedStartSettleSeconds
    )
    $claim = $null
    try { $claim = Enter-TestbedLeaseClaim -VMName $TestbedGateClaimName -TimeoutSeconds 10 }
    catch { return [pscustomobject]@{ Go = $false; Why = "another run is taking a VM just now ($($_.Exception.Message))"; Claim = $null; PendingPercent = 0.0; PendingCount = 0 } }
    $keep = $false
    try {
        $pending = Get-PendingStarts -Starts @(Read-TestbedStarts) -NowUnix ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) -SettleSeconds $SettleSeconds -ExceptRun $Run
        $busy = $null
        if ($null -ne $Look) { $busy = $Look.BusyPercent }
        $why = Get-TestbedGateVerdict -Pause (Get-TestbedPause) -BusyPercent $busy -WindowSeconds $WindowSeconds -PendingPercent $pending.Percent -PendingCount $pending.Count -LimitPercent $LimitPercent -SettleSeconds $SettleSeconds
        $keep = -not $why
        $held = $null
        if ($keep) { $held = $claim }
        return [pscustomobject]@{ Go = $keep; Why = $why; Claim = $held; PendingPercent = $pending.Percent; PendingCount = $pending.Count }
    }
    finally { if (-not $keep) { $claim.Dispose() } }
}

# Writes the start of the run whose $Start holds the claim: the VM it took and that VM's share of the
# host, which every other run counts for the next 120 s. Older starts are dropped. A start not written
# leaves only the next run's look without it, so this never fails the run.
function Register-TestbedStart {
    param([object] $Start, [string] $Run, [string] $VmName, [int] $VmProcessors)
    if ($null -eq $Start -or $null -eq $Start.Claim) { return }
    try {
        $nowUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        if ($VmProcessors -lt 1) { $VmProcessors = 4 }   # not known: the smallest testbed VM's
        $kept = @(@(Read-TestbedStarts) | Where-Object { $age = 0L; try { $age = $nowUnix - [long]$_.unix } catch { $age = [long]::MaxValue }; $age -ge 0 -and $age -lt $TestbedStartSettleSeconds })
        $kept += [pscustomobject]@{ unix = $nowUnix; share = [Math]::Round((Get-VmSharePercent $VmProcessors ([Environment]::ProcessorCount)), 2); processors = $VmProcessors; vm = $VmName; run = $Run }
        Write-TestbedStarts $kept
    }
    catch { }
}

function Exit-TestbedStart {
    param([object] $Start)
    if ($null -ne $Start -and $null -ne $Start.Claim) { $Start.Claim.Dispose(); $Start.Claim = $null }
}

# For a caller that waits before it takes anything - a live run before its build, New-TestbedClone.ps1
# before its export: looks until the gate is open, with the VMs other runs took lately counted (read,
# not claimed: such a caller writes no start), or until $TimeoutMinutes have passed. Between looks it
# sleeps until a file of the testbed's state changes - the switch turned off, a start, a lease, a lock's
# owner file, a ticket - or $PollSeconds pass, 30 by default (TestbedPool.ps1's Wait-TestbedChange):
# never a poll of its own, which every 5 s until 2026-11-18 it was. Says what it waits for when it begins to wait and once a minute after, through $Say, and is
# listed in waiting\ meanwhile. Go, Why, WaitedSeconds, and the Look and Pending it went on.
function Wait-TestbedGate {
    param(
        [string] $Run,
        [string] $What,
        [int] $TimeoutMinutes,
        [scriptblock] $Say,
        [int] $LimitPercent = $TestbedLoadLimitPercent,
        [int] $WindowSeconds = $TestbedLoadWindowSeconds,
        [int] $PollSeconds = $TestbedWaitSeconds
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $since = $null
    $lastSaid = [DateTime]::MinValue
    try {
        while ($true) {
            $look = Test-TestbedGate -LimitPercent $LimitPercent -WindowSeconds $WindowSeconds
            $why = $look.Why
            $pending = [pscustomobject]@{ Percent = 0.0; Count = 0 }
            if ($look.Go) {
                $pending = Get-PendingStarts -Starts @(Read-TestbedStarts) -NowUnix ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) -ExceptRun $Run
                $why = Get-TestbedGateVerdict -Pause $null -BusyPercent $look.BusyPercent -WindowSeconds $WindowSeconds -PendingPercent $pending.Percent -PendingCount $pending.Count -LimitPercent $LimitPercent
                if (-not $why) {
                    $waited = 0.0
                    if ($null -ne $since) { $waited = ((Get-Date) - $since).TotalSeconds }
                    return [pscustomobject]@{ Go = $true; Why = ''; WaitedSeconds = $waited; Look = $look; Pending = $pending }
                }
            }
            if ($null -eq $since) { $since = Get-Date }
            if (((Get-Date) - $lastSaid).TotalSeconds -ge 60) {
                if ($null -ne $Say) { & $Say (Format-TestbedWaitLine -Since $since -Why $why -First ($lastSaid -eq [DateTime]::MinValue)) }
                Set-TestbedWaiting -What $What -Why $why -Since $since
                $lastSaid = Get-Date
            }
            if ((Get-Date) -ge $deadline) { return [pscustomobject]@{ Go = $false; Why = $why; WaitedSeconds = ((Get-Date) - $since).TotalSeconds; Look = $look; Pending = $pending } }
            [void](Wait-TestbedChange -Seconds $PollSeconds)
        }
    }
    finally { if ($null -ne $since) { Remove-TestbedWaiting } }
}
