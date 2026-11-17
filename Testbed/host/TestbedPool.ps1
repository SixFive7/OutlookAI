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
