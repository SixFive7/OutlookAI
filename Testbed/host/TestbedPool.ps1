#Requires -Version 5.1
<#
.SYNOPSIS
    The testbed's VM POOLS: which VMs are copies of which, and whether the host has the memory to
    start one more. Dot-sourced by the two runners, the clone maker and the idle-saver, so that
    they cannot disagree.

.DESCRIPTION
    A POOL IS A SOURCE VM AND ITS CLONES (2026-11-17). The maintainer, that day: "try to
    parallelize as much as possible. Even create extra virtual machines if you think that can speed
    things up." Eight agents were queueing on one build VM and two Outlook guests. A clone is a
    whole copy of its source made by Testbed/host/New-TestbedClone.ps1 from the checkpoint the
    source's runner restores: the same disk, the same memory and clock, the same computer name,
    account, stores and settings, and that checkpoint under the same name. So any member of a pool
    can take any run meant for the pool: a run that names a pool takes whichever member is free,
    and a run that names a VM takes that VM.

    THE MEMBERS ARE LISTED HERE, ONCE. Testbed/testbed.json's vmPools block records each clone -
    which checkpoint it was made from, when, and the run that proved it - and
    Invoke-TestsOnBuildVm.ps1 -SelfTest fails while the two disagree. A clone joins its pool here
    only after a run on it passed: a run that names it with -VMName proves it first.

    NEVER OutlookAI-Exchange. It holds a real Microsoft 365 sign-in and the device state that goes
    with it (Q108, Q111): a copy would be a second device signed in to the maintainer's tenant, with
    the same identity. It is in no pool, and New-TestbedClone.ps1 refuses it by name.

    ONE MORE VM ONLY WHEN THE HOST HAS THE MEMORY FOR IT, AND SOME TO SPARE. The maintainer works
    on this PC. A pool's first member is taken whenever it is free, as the one VM always was. Any
    other member is taken only when the host's available memory covers that VM's startup memory
    plus $TestbedHostReserveBytes, 16 GB - read from kernel32's GlobalMemoryStatusEx, which is what
    Windows calls available, free and standby together, and costs no WMI call (2026-11-18: 28.0 GB
    there where Win32_OperatingSystem.FreePhysicalMemory, read until then, said 28.1 GB - and that
    read took about 2 s, stream W measured). A run that finds no member it may take waits, as
    runs waited for the one VM before.

    WAITING ASKS HYPER-V NOTHING (2026-11-18, the maintainer's PC lagging). Every queued run asked
    Hyper-V about every member every 5 s - Get-VM, which reads each VM's summary, settings and their
    associations through WMI, served inside vmms.exe and the WMI service at normal priority whatever
    the caller's: a WMI-Activity trace at 04:09 that night counted 32 operations a second, 91 % of
    them those, all from waiting runners - about 0.2 to 0.6 of a logical processor in vmms and the
    WMI service, stream W measured. (WmiPrvSE's 3.4 to 3.8 processors that night were mostly other
    programs listing processes; the runners' share of them was their memory read, about 0.3 of a
    processor for each waiting run - Docs/live-tier-on-the-vm.md section 4.10.) So a run that waits now:
      - joins a QUEUE, a ticket file in %SystemDrive%\OutlookAI-Testbed\queue\ naming the VMs it
        can use (Join-TestbedQueue). It tries a VM only while no run that queued before it wants
        that VM (Get-TestbedQueueTurn): one run at a time tries each VM, first come first served,
        and every other run waits on the queue's files only. A run of a checkout from before the
        queue takes no ticket and takes a free VM as it always did;
      - asks whether a VM is free of FILES only - a build VM's lock, a lease, the queue - and the
        host's memory of GlobalMemoryStatusEx; what Hyper-V must tell it - whether its VMs exist,
        their ids, their startup memory - it asks once, before it starts waiting;
      - sleeps until a file of the testbed's state changes - a lock's owner file, a lease, a ticket -
        or 30 s pass (Wait-TestbedChange): no poll in between, and none more often than that;
      - and a run that runs reads its VM's state and heartbeat with one narrow CIM query each, by
        the VM's id, on a session kept for the process (Get-TestbedVms, Get-TestbedVmStateById,
        Test-TestbedHeartbeat) - never Get-VM or Get-VMIntegrationService, which cost 4 and 23 WMI
        operations a call (stream W, 2026-11-18). The live runner and the trust-removal probe ask
        them by the guest's name (Get-TestbedVmId, Get-TestbedVmState, Wait-TestbedHeartbeat), and
        restore a checkpoint over a guest the one way, saving a running guest first
        (Restore-TestbedCheckpoint).
#>

# The pools, the source first: the member every run tried first before there were clones.
$TestbedPools = [ordered]@{
    build     = @('OutlookAI-Build', 'OutlookAI-Build2', 'OutlookAI-Build3')
    indexed   = @('OutlookAI-Indexed', 'OutlookAI-Indexed2')
    unindexed = @('OutlookAI-Unindexed', 'OutlookAI-Unindexed2')
}

# What a pool member must leave free for the maintainer: see the banner.
$TestbedHostReserveBytes = 16GB

# The pool a VM is a member of - 'build', 'indexed' or 'unindexed' - or '' for any other name.
# Compared whole and ignoring case, as Hyper-V compares names; a padded name is no member.
function Get-TestbedPoolOf {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -ne $Name.Trim()) { return '' }
    foreach ($pool in $TestbedPools.Keys) {
        foreach ($member in $TestbedPools[$pool]) {
            if ([string]::Equals($Name, $member, [System.StringComparison]::OrdinalIgnoreCase)) { return $pool }
        }
    }
    return ''
}

# The member of its pool a VM is a copy of: the pool's first member. '' for a name in no pool.
function Get-TestbedPoolSource {
    param([string] $Name)
    $pool = Get-TestbedPoolOf $Name
    if (-not $pool) { return '' }
    return $TestbedPools[$pool][0]
}

# Every VM of every pool, sources first - what the idle-saver watches beside the Exchange VM.
function Get-TestbedPoolMembers {
    $all = @()
    foreach ($pool in $TestbedPools.Keys) { $all += @($TestbedPools[$pool]) }
    return $all
}

# Whether a pool may start one more VM that needs $NeedBytes, with $AvailableBytes available. Pure.
function Test-HostHasRoom {
    param([long] $NeedBytes, [long] $AvailableBytes, [long] $ReserveBytes = $TestbedHostReserveBytes)
    return ($AvailableBytes -ge ($NeedBytes + $ReserveBytes))
}

# The host's available memory, in bytes, or -1 when it cannot be read - which Test-HostHasRoom
# reads as no room: an extra VM is never started on a guess. kernel32's GlobalMemoryStatusEx, no WMI
# (the banner), compiled once a process by Add-Type - on Windows PowerShell 5.1 the .NET Framework's
# own csc.exe, so nothing is installed for it (AGENTS.md, Dependencies). C# 5.
function Get-HostAvailableBytes {
    try {
        if (-not ('OutlookAI.Testbed.HostMemory' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace OutlookAI.Testbed
{
    public static class HostMemory
    {
        [StructLayout(LayoutKind.Sequential)]
        private class MemoryStatusEx
        {
            public uint Length = 64;
            public uint MemoryLoad;
            public ulong TotalPhys;
            public ulong AvailPhys;
            public ulong TotalPageFile;
            public ulong AvailPageFile;
            public ulong TotalVirtual;
            public ulong AvailVirtual;
            public ulong AvailExtendedVirtual;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx([In, Out] MemoryStatusEx status);

        // Available physical memory - free and standby - in bytes; -1 when Windows does not say.
        public static long AvailableBytes()
        {
            MemoryStatusEx status = new MemoryStatusEx();
            if (!GlobalMemoryStatusEx(status)) { return -1; }
            return (long)status.AvailPhys;
        }
    }
}
'@
        }
        return [long][OutlookAI.Testbed.HostMemory]::AvailableBytes()
    }
    catch { return -1 }
}

# Where the build runner keeps its locks, owner files, runs.log and janitor.log: machine-wide,
# beside the leases (TestbedLeasePath.ps1), so that every checkout's runner, its janitors and
# New-TestbedClone.ps1 see the same files.
function Get-BuildRunnerDirectory {
    $dir = Join-Path $env:SystemDrive 'OutlookAI-Testbed\build-vm'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    return $dir
}

# The build runner's lock and owner file for a build VM, by name, in the runner's directory
# (Get-BuildRunnerDirectory). OutlookAI-Build keeps the names every runner before
# the pool used, runner.lock and runner.owner.json, so that a run from an older checkout - each
# branch runs its own copy of the runner - and a run of this one still exclude each other on it.
function Get-BuildRunnerLockNames {
    param([string] $VmName)
    if ([string]::Equals($VmName, $TestbedPools['build'][0], [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{ Lock = 'runner.lock'; Owner = 'runner.owner.json' }
    }
    return [pscustomobject]@{ Lock = "runner.$VmName.lock"; Owner = "runner.$VmName.owner.json" }
}

# =============================================================================================
# WAITING ASKS HYPER-V NOTHING (2026-11-18) - the banner. The build runner's -SelfTest pins the pure
# parts and the queue's files.
# =============================================================================================

# How long a waiting run sleeps at most before it looks again, and how often a running one asks its
# guest how far it got: the maintainer's ceilings on polling, 2026-11-18. A heartbeat is one narrow
# query (Test-TestbedHeartbeat), asked every 2 s while a VM comes up, as stream W measured cheap.
$TestbedWaitSeconds = 30
$TestbedGuestPollSeconds = 15
$TestbedHeartbeatSeconds = 2
# Set by a -SelfTest to a scratch directory; empty, the queue is the machine's.
$TestbedQueueDirectoryOverride = ''

# A VM's state from Hyper-V's Msvm_ComputerSystem.EnabledState, in Get-VM's words: Running, Off,
# Saved and Paused are settled, and anything else is on its way to one of them. Both numberings
# Hyper-V has used are read.
function ConvertFrom-TestbedEnabledState {
    param([int] $EnabledState)
    switch ($EnabledState) {
        2 { return 'Running' }
        3 { return 'Off' }
        6 { return 'Saved' }
        9 { return 'Paused' }
        32768 { return 'Paused' }
        32769 { return 'Saved' }
        10 { return 'Starting' }
        32770 { return 'Starting' }
        4 { return 'Stopping' }
        32774 { return 'Stopping' }
        32771 { return 'Snapshotting' }
        32773 { return 'Saving' }
        32776 { return 'Pausing' }
        32777 { return 'Resuming' }
    }
    return "EnabledState $EnabledState"
}

# Which of $Candidates a queued run may NOT try yet, each with the run ahead of it that wants it: a
# run tries a VM only while no live ticket before its own names that VM. $Tickets are the live
# tickets with Name - the file name, which sorts them in queue order - Run and Candidates. Pure.
function Select-TestbedQueueTurn {
    param([object[]] $Tickets, [string] $Mine, [string[]] $Candidates)
    $ahead = @{}
    foreach ($t in @(@($Tickets) | Where-Object { $null -ne $_ } | Sort-Object { [string]$_.Name })) {
        if ([string]$t.Name -eq $Mine) { break }
        foreach ($wanted in @($t.Candidates)) {
            foreach ($candidate in @($Candidates)) {
                if ([string]::Equals([string]$wanted, [string]$candidate, [System.StringComparison]::OrdinalIgnoreCase) -and -not $ahead.ContainsKey($candidate)) { $ahead[$candidate] = [string]$t.Run }
            }
        }
    }
    return $ahead
}

# Where waiting runs queue: machine-wide, beside the build runner's directory and the leases.
function Get-TestbedQueueDirectory {
    $dir = [string]$TestbedQueueDirectoryOverride
    if (-not $dir) { $dir = Join-Path (Split-Path -Parent (Get-BuildRunnerDirectory)) 'queue' }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    return $dir
}

# A run joins the queue for the VMs it can use: a ticket named for when it joined, so that the names
# sort in queue order, holding what it is and which VMs it wants. -Ahead puts it before every run
# already waiting: New-TestbedClone.ps1, whose export holds its source for half a minute and ends the
# queue it jumps. Returns the ticket: Path, Name and the Json it was written with.
function Join-TestbedQueue {
    param([string] $Run, [string] $What, [string[]] $Candidates, [switch] $Ahead)
    $ticks = [DateTime]::UtcNow.Ticks
    if ($Ahead) { $ticks = 0L }
    $name = '{0:D19}-{1}.json' -f $ticks, $PID
    $record = [ordered]@{
        pid              = $PID
        processStartUnix = ([DateTimeOffset](Get-Process -Id $PID).StartTime).ToUnixTimeSeconds()
        run              = $Run
        what             = $What
        candidates       = @($Candidates)
        joinedUnix       = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    }
    $json = ConvertTo-Json -InputObject $record -Depth 3
    $path = Join-Path (Get-TestbedQueueDirectory) $name
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
    return [pscustomobject]@{ Path = $path; Name = $name; Json = $json }
}

# THE TESTBED'S SHARED FILES - the queue's tickets, the gate's waiting records and its switch - are
# written and deleted by several runs at once: eight runs were queued at 09:58 on 2026-11-18. A file
# listed a moment ago can be gone when it is read, and two runs can find the same ended run's ticket and
# delete it in the same instant. So these files are listed by name alone (Get-TestbedStateFiles), a
# file that cannot be read counts as gone (Read-TestbedLiveRecords), and a delete goes through
# Remove-TestbedStateFile - never Remove-Item, which throws for a file that goes between its own look
# and its delete, "An object at the specified path ... does not exist" or, under Windows PowerShell
# 5.1, "Access is denied" for a file already pending delete, whatever -ErrorAction says. That ended
# stream M's build run 20261118-092047 before it started; on a scratch queue whose tickets came and
# went, 183 of 600 turns threw under PowerShell 7.7 and 150 of 600 under 5.1.

# The .json files of one of the testbed's state directories: full paths, in ordinal order of their
# names - for the queue, its order. Names only, so nothing is asked of a file that has gone since. None
# when the directory cannot be listed.
function Get-TestbedStateFiles {
    param([string] $Directory)
    [string[]] $paths = @()
    try { $paths = [System.IO.Directory]::GetFiles($Directory, '*.json') } catch { return @() }
    [Array]::Sort($paths, [System.StringComparer]::OrdinalIgnoreCase)
    return $paths
}

# Deletes one of those files: $true when this call deleted it, $false when it had gone already or
# could not be deleted - another run deleting it at the same instant. Never throws.
function Remove-TestbedStateFile {
    param([string] $Path)
    try {
        if (-not [System.IO.File]::Exists($Path)) { return $false }
        [System.IO.File]::Delete($Path)
        return $true
    }
    catch { return $false }
}

# Leaves the queue: the ticket goes. Never throws.
function Exit-TestbedQueue {
    param([object] $Ticket)
    try { if ($null -ne $Ticket) { [void](Remove-TestbedStateFile ([string]$Ticket.Path)) } } catch { }
}

# Whether a record's run still lives - a ticket's, or a waiting record's: its process, started when
# the record says - not another that took its id since.
function Test-TestbedTicketAlive {
    param([object] $Record)
    $p = $null
    try { $p = Get-Process -Id ([int]$Record.pid) -ErrorAction SilentlyContinue } catch { }
    if ($null -eq $p) { return $false }
    try { return (([DateTimeOffset]$p.StartTime).ToUnixTimeSeconds() -eq [long]$Record.processStartUnix) } catch { return $true }
}

# The records of $Paths - tickets, waiting records - in their order, up to the one named $Until: each
# read, and Name, Path and Record given back for every one whose run still lives. One that cannot be
# read is passed over as gone - another run may have deleted it since the list was made, or still be
# writing it - and one whose run has ended is deleted. Never throws for a file that goes meanwhile.
function Read-TestbedLiveRecords {
    param([string[]] $Paths, [string] $Until = '')
    $live = @()
    foreach ($path in @($Paths)) {
        $name = [System.IO.Path]::GetFileName($path)
        if ($Until -and $name -eq $Until) { break }
        $record = $null
        try { $record = [System.IO.File]::ReadAllText($path) | ConvertFrom-Json } catch { continue }
        if ($null -eq $record) { continue }
        if (-not (Test-TestbedTicketAlive $record)) { [void](Remove-TestbedStateFile $path); continue }
        $live += [pscustomobject]@{ Name = $name; Path = $path; Record = $record }
    }
    return $live
}

# This run's turn in the queue: the tickets ahead of its own are read, those of runs that have ended -
# stopped while they waited - deleted, and Select-TestbedQueueTurn says which of $Candidates it may
# not try yet. A ticket of its own that has gone is written again under its name, so the run keeps
# its place. Files only.
function Get-TestbedQueueTurn {
    param([object] $Ticket, [string] $Run, [string[]] $Candidates)
    $dir = Get-TestbedQueueDirectory
    if (-not (Test-Path -LiteralPath $Ticket.Path)) { [System.IO.File]::WriteAllText($Ticket.Path, $Ticket.Json, (New-Object System.Text.UTF8Encoding($false))) }
    $live = @(Read-TestbedLiveRecords -Paths @(Get-TestbedStateFiles $dir) -Until $Ticket.Name | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Run = [string]$_.Record.run; Candidates = @($_.Record.candidates) } })
    $live += [pscustomobject]@{ Name = $Ticket.Name; Run = $Run; Candidates = @($Candidates) }
    return (Select-TestbedQueueTurn -Tickets $live -Mine $Ticket.Name -Candidates $Candidates)
}

# Sleeps until a .json file under the testbed's state directory changes - a build VM's owner file
# as a run lets its lock go, a lease, a ticket - or $Seconds pass. A FileSystemWatcher: no poll at
# all in between. After a change it gives the writer a second to finish - a run removes its owner
# file just before it lets go of its lock. $true when something changed.
function Wait-TestbedChange {
    param([int] $Seconds = $TestbedWaitSeconds)
    $root = Split-Path -Parent (Get-BuildRunnerDirectory)
    $watcher = New-Object System.IO.FileSystemWatcher($root, '*.json')
    try {
        $watcher.IncludeSubdirectories = $true
        $watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite
        $result = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::All, [int]($Seconds * 1000))
    }
    finally { $watcher.Dispose() }
    if ($result.TimedOut) { return $false }
    Start-Sleep -Seconds 1
    return $true
}

# THE NARROW QUERIES (stream W, 2026-11-18): one DCOM CIM session for the process - a new session a
# call costs three WMI operations where a kept one costs one - and every query naming the properties
# it reads: Get-VM starts with an enumeration of Msvm_SummaryInformation that cost vmms 810 ms of processor when
# asked for everything, and Get-VMIntegrationService costs 23 operations a call.
function Get-TestbedCimSession {
    if ($null -eq $script:TestbedCimSession) { $script:TestbedCimSession = New-CimSession -SessionOption (New-CimSessionOption -Protocol Dcom) }
    return $script:TestbedCimSession
}

# VMs by name, in ONE query for all of them: each one's Hyper-V id - Msvm_ComputerSystem.Name, which
# the other narrow queries ask by - and its state. Asked once a run, before it waits: whether its VMs
# exist. A name that is no plain VM name is refused before anything is asked. One row per VM found -
# a name Hyper-V has twice gives two - with Name, Id and State.
function Get-TestbedVms {
    param([string[]] $Name)
    $names = @($Name | Where-Object { $_ })
    foreach ($n in $names) { if ($n -notmatch '^[A-Za-z0-9._-]{1,64}$') { throw "'$n' is not a plain VM name; refusing to ask Hyper-V about it." } }
    if ($names.Count -eq 0) { return @() }
    $where = (@($names | ForEach-Object { "ElementName='$_'" }) -join ' OR ')
    $rows = @(Get-CimInstance -CimSession (Get-TestbedCimSession) -Namespace 'root\virtualization\v2' -Query "SELECT Name, ElementName, EnabledState FROM Msvm_ComputerSystem WHERE $where" -ErrorAction Stop)
    return @($rows | ForEach-Object { [pscustomobject]@{ Name = [string]$_.ElementName; Id = [string]$_.Name; State = (ConvertFrom-TestbedEnabledState ([int]$_.EnabledState)) } })
}

# A VM's Hyper-V id as these queries take it: a GUID and nothing else, so it can go into WQL as it is.
function Assert-TestbedVmId {
    param([string] $Id)
    if ($Id -notmatch '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$') { throw "'$Id' is not a VM id; refusing to ask Hyper-V about it." }
}

# A VM's state by its id - Running, Saved, Off, Paused, or on its way to one: Msvm_ComputerSystem's
# EnabledState, one keyed query. '' when Hyper-V has no VM of that id.
function Get-TestbedVmStateById {
    param([string] $Id)
    Assert-TestbedVmId $Id
    $rows = @(Get-CimInstance -CimSession (Get-TestbedCimSession) -Namespace 'root\virtualization\v2' -Query "SELECT EnabledState FROM Msvm_ComputerSystem WHERE Name='$Id'" -ErrorAction Stop)
    if ($rows.Count -ne 1) { return '' }
    return (ConvertFrom-TestbedEnabledState ([int]$rows[0].EnabledState))
}

# Whether a VM's heartbeat reads OK, by its id: Msvm_HeartbeatComponent's OperationalStatus, whose
# first value is 2 for OK - one narrow query, where Get-VMIntegrationService costs 23 operations.
function Test-TestbedHeartbeat {
    param([string] $Id)
    Assert-TestbedVmId $Id
    $rows = @(Get-CimInstance -CimSession (Get-TestbedCimSession) -Namespace 'root\virtualization\v2' -Query "SELECT OperationalStatus FROM Msvm_HeartbeatComponent WHERE SystemName='$Id'" -ErrorAction Stop)
    if ($rows.Count -lt 1) { return $false }
    $status = @($rows[0].OperationalStatus)
    return ($status.Count -gt 0 -and [int]$status[0] -eq 2)
}

# A VM's Hyper-V id by its name: Get-TestbedVms the first time a script asks, and the answer kept for
# the rest of it - an id does not change while the VM exists. Throws unless exactly one VM has the
# name. The build-VM runner keeps its own, Get-BuildVmId: it asks only through its pool's allowlist,
# and is handed the id as it takes a VM.
function Get-TestbedVmId {
    param([string] $Name)
    if ($null -eq $script:TestbedVmIds) { $script:TestbedVmIds = @{} }
    if (-not $script:TestbedVmIds.ContainsKey($Name)) {
        $rows = @(Get-TestbedVms -Name @($Name))
        if ($rows.Count -ne 1) { throw "expected one VM named '$Name' on this host and found $($rows.Count)" }
        $script:TestbedVmIds[$Name] = [string]$rows[0].Id
    }
    return $script:TestbedVmIds[$Name]
}

# A VM's state by its name - Running, Saved, Off, Paused, or on its way - by one keyed query on its id,
# never Get-VM. Throws when Hyper-V no longer has the VM.
function Get-TestbedVmState {
    param([string] $Name)
    $state = Get-TestbedVmStateById (Get-TestbedVmId $Name)
    if (-not $state) { throw "Hyper-V no longer has '$Name' on this host" }
    return [string]$state
}

# Waits until a VM's heartbeat reads OK, asking Test-TestbedHeartbeat every $TestbedHeartbeatSeconds,
# and throws once -Minutes have passed without it. Get-VMIntegrationService, asked every 1 to 3 s until
# 2026-11-18, cost 23 WMI operations a look (stream W).
function Wait-TestbedHeartbeat {
    param([string] $Id, [int] $Minutes, [string] $What)
    $deadline = (Get-Date).AddMinutes($Minutes)
    while (-not (Test-TestbedHeartbeat $Id)) {
        if ((Get-Date) -ge $deadline) { throw "$What's heartbeat did not read OK within $Minutes minute(s)" }
        Start-Sleep -Seconds $TestbedHeartbeatSeconds
    }
}

# A checkpoint restored over a test guest, by the guest's name. A guest that runs - or is paused - is
# SAVED first (2026-11-18): a restore over a running guest left it out of PowerShell Direct's reach for
# six minutes in stream R1's runs, where saved first and then restored it answered in seconds. -Start
# starts a guest the restore left stopped or saved, and waits up to 5 minutes for its heartbeat.
# Returns the state the restore left the guest in, before any start.
function Restore-TestbedCheckpoint {
    param([string] $Name, [string] $Checkpoint, [switch] $Start)
    $state = Get-TestbedVmState $Name
    if ($state -eq 'Running' -or $state -eq 'Paused') { Save-VM -Name $Name }
    Restore-VMSnapshot -VMName $Name -Name $Checkpoint -Confirm:$false
    $after = Get-TestbedVmState $Name
    if ($Start) {
        if ($after -ne 'Running') { Start-VM -Name $Name }
        Wait-TestbedHeartbeat -Id (Get-TestbedVmId $Name) -Minutes 5 -What $Name
    }
    return $after
}
