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
    plus $TestbedHostReserveBytes, 16 GB - read from Win32_OperatingSystem.FreePhysicalMemory, which
    is what Windows calls available, free and standby together (measured on this host 2026-11-17:
    63.9 GB there, 64.2 GB from the Memory\Available MBytes counter). A run that finds no member it
    may take waits, as runs waited for the one VM before.

    WAITING ASKS HYPER-V NOTHING (2026-11-18, the maintainer's PC lagging). Every queued run asked
    Hyper-V for every member every 5 s - Get-VM, which reads each VM's summary, settings and their
    associations through WMI - and WmiPrvSE served it at normal priority: a WMI-Activity trace at
    04:09 that night counted 32 operations a second, 91 % of them those, all from waiting runners,
    and WmiPrvSE used about 5 of the host's 32 logical processors. So a run that waits now:
      - joins a QUEUE, a ticket file in %SystemDrive%\OutlookAI-Testbed\queue\ naming the VMs it
        can use (Join-TestbedQueue). It tries a VM only while no run that queued before it wants
        that VM (Get-TestbedQueueTurn): one run at a time tries each VM, first come first served,
        and every other run waits on the queue's files only. A run of a checkout from before the
        queue takes no ticket and takes a free VM as it always did;
      - asks whether a VM is free of FILES only - a build VM's lock, a lease, the queue - and asks
        Hyper-V only once one is had, for the facts it then needs;
      - sleeps until a file of the testbed's state changes - a lock's owner file, a lease, a ticket -
        or 30 s pass (Wait-TestbedChange): no poll in between, and none more often than that;
      - reads a VM's state, where it must, with one narrow CIM query on a session kept for the
        process (Get-TestbedVmState): Msvm_ComputerSystem's EnabledState, for every VM asked about
        at once - never Get-VM.
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
# reads as no room: an extra VM is never started on a guess.
function Get-HostAvailableBytes {
    try { return ([long](Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).FreePhysicalMemory) * 1024 }
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
# guest how far it got or its VM for a heartbeat: the maintainer's ceilings on polling, 2026-11-18.
$TestbedWaitSeconds = 30
$TestbedGuestPollSeconds = 15
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

# Leaves the queue: the ticket goes. Never throws.
function Exit-TestbedQueue {
    param([object] $Ticket)
    try { if ($null -ne $Ticket -and (Test-Path -LiteralPath $Ticket.Path)) { Remove-Item -LiteralPath $Ticket.Path -Force } } catch { }
}

# Whether a ticket's run still lives: its process, started when the ticket says - not another that
# took its id since.
function Test-TestbedTicketAlive {
    param([object] $Record)
    $p = $null
    try { $p = Get-Process -Id ([int]$Record.pid) -ErrorAction SilentlyContinue } catch { }
    if ($null -eq $p) { return $false }
    try { return (([DateTimeOffset]$p.StartTime).ToUnixTimeSeconds() -eq [long]$Record.processStartUnix) } catch { return $true }
}

# This run's turn in the queue: the tickets ahead of its own are read, those of runs that have ended -
# stopped while they waited - deleted, and Select-TestbedQueueTurn says which of $Candidates it may
# not try yet. A ticket of its own that has gone is written again under its name, so the run keeps
# its place. Files only.
function Get-TestbedQueueTurn {
    param([object] $Ticket, [string] $Run, [string[]] $Candidates)
    $dir = Get-TestbedQueueDirectory
    if (-not (Test-Path -LiteralPath $Ticket.Path)) { [System.IO.File]::WriteAllText($Ticket.Path, $Ticket.Json, (New-Object System.Text.UTF8Encoding($false))) }
    $live = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        if ($f.Name -eq $Ticket.Name) { break }
        $t = $null
        try { $t = [System.IO.File]::ReadAllText($f.FullName) | ConvertFrom-Json } catch { continue }
        if ($null -eq $t) { continue }
        if (-not (Test-TestbedTicketAlive $t)) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue; continue }
        $live += [pscustomobject]@{ Name = $f.Name; Run = [string]$t.run; Candidates = @($t.candidates) }
    }
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

# The state of VMs by name - Running, Saved, Off, Paused, or on its way to one - from ONE narrow CIM
# query for all of them, on a session kept for the process: Msvm_ComputerSystem's EnabledState and
# nothing else. Get-VM, which also reads each VM's summary, settings and associations, is never the
# way to ask how a VM is. A name that is no plain VM name is refused before anything is asked. One
# row per VM found - a name Hyper-V has twice gives two - with Name and State.
function Get-TestbedVmState {
    param([string[]] $Name)
    $names = @($Name | Where-Object { $_ })
    foreach ($n in $names) { if ($n -notmatch '^[A-Za-z0-9._-]{1,64}$') { throw "'$n' is not a plain VM name; refusing to ask Hyper-V about it." } }
    if ($names.Count -eq 0) { return @() }
    if ($null -eq $script:TestbedCimSession) { $script:TestbedCimSession = New-CimSession -SessionOption (New-CimSessionOption -Protocol Dcom) }
    $filter = (@($names | ForEach-Object { "ElementName='$_'" }) -join ' OR ')
    $rows = @(Get-CimInstance -CimSession $script:TestbedCimSession -Namespace 'root\virtualization\v2' -ClassName 'Msvm_ComputerSystem' -Filter $filter -Property ElementName, EnabledState -ErrorAction Stop)
    return @($rows | ForEach-Object { [pscustomobject]@{ Name = [string]$_.ElementName; State = (ConvertFrom-TestbedEnabledState ([int]$_.EnabledState)) } })
}
