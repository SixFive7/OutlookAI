#Requires -Version 5.1
<#
.SYNOPSIS
    Makes CLONES of a testbed VM for its pool (Testbed/host/TestbedPool.ps1): whole copies of the
    VM at the checkpoint its runner restores, under new names, so that runs stop queueing on one
    machine. A dry run unless -Execute.

.DESCRIPTION
    RUN ON THE HOST. Windows PowerShell 5.1 or PowerShell 7, as a member of the local Hyper-V
    Administrators group - no elevation. It needs no guest credential: nothing here logs on to a
    guest. A run that names a clone proves it afterwards.

    WHY (2026-11-17). The maintainer: "try to parallelize as much as possible. Even create extra
    virtual machines if you think that can speed things up." Eight agents queued on OutlookAI-Build
    and the two Outlook guests. A clone is the cheapest correct second machine: no install, no
    build-out, no new checkpoint to prove from nothing - the source's own runner checkpoint,
    copied.

    WHAT A CLONE IS. The source's runner checkpoint - OutlookAI-Build's runnerBaseCheckpoint, an
    Outlook guest's frozen checkpoint (Testbed/testbed.json, buildVm and frozenClocks) - written out
    ONCE with Export-VMSnapshot, which exports a checkpoint as a VM of its own: its disk chain merged
    into one VHDX, and the memory the checkpoint was taken with. Then, for each clone, that export
    imported as a COPY with a NEW ID (Import-VM -Copy -GenerateNewId) into folders of the clone's
    own, renamed in the compatibility report BEFORE it is registered - so the host never holds two
    VMs of one name - and a checkpoint of it, still saved, under the SAME NAME as the source's.
    Every runner restores that name on whichever pool member it takes, and a clone resumes exactly
    where the source does: the same disk, memory and clock. A frozen guest's clone is frozen at the
    same instant, its time synchronisation off with it. The export is deleted at the end.

    WHAT A CLONE KEEPS OF ITS SOURCE, deliberately: the computer name (OAI-BUILD, OAI-INDEXED,
    OAI-UNINDEXED), the machine SID, the account and its password, the stores, the mail sink on
    loopback and the vTPM with its key protector, which unwraps on this host only. None of it meets
    another machine: the build VMs and the Outlook guests have no network - their adapters are
    disconnected, and a clone's is disconnected if it is not - and PowerShell Direct finds a VM by
    its Hyper-V name, never by its computer name. A new computer name would need a restart, which a
    frozen guest may never have (it brings the clock back at the real date), and a new checkpoint
    to prove; it would buy nothing. So the runners hold a clone to its source's computer name, and
    a live run's opt-in is that name.

    LICENSING, checked 2026-11-17 against Testbed/MEDIA.md. Windows is the consumer Pro image,
    unactivated, as on every testbed VM: no expiry - it watermarks and nags, and stops nothing.
    Office on the Outlook guests is ProPlus 2024 volume, a KMS client that has never reached a KMS
    host (no network): a frozen guest's grace stands where its frozen instant put it - Office was
    installed 2026-09-15, so about twelve of its 30 days remain, for ever - and MEDIA.md measured
    every COM read and a cold start working past grace anyway. The build VM has no Office. A clone
    adds no component and no media, so the Dependencies rule in AGENTS.md is untouched; it is one
    more unactivated Windows (and Office) like the ones the testbed already runs.

    NEVER OutlookAI-Exchange. It holds a real Microsoft 365 sign-in and the device state that goes
    with it (Q108, Q111); a copy would be a second device signed in to the maintainer's tenant.
    Refused by name, before anything else.

    ONLY NAMES TestbedPool.ps1 ALREADY LISTS in the source's pool, and never the source itself: a
    clone no runner and no idle-saver knows is a VM nothing uses and nothing saves. List it there,
    make it, prove it with a run that names it, then record it in testbed.json's vmPools; the clone
    joins the pool on master when that branch lands.

    BUDGETS. Disk: a clone holds what its source's checkpoint chain holds, merged - at most the
    chain's size (2026-11-17: Build 49 GB, Indexed 109 GB, Unindexed 93 GB), and far less: a build
    clone measured 27.5 GB at rest and the indexed one 40 GB - plus the checkpoint's
    saved memory; the export holds as much once more until it is deleted at the end. Refused unless
    the drive has that for every clone and the export, plus 60 GB. Memory: none while a clone rests
    saved. The runners start one only when the host has its startup memory plus 16 GB to spare.

    THE GATE FIRST (Q195 (e), decided by the maintainer 2026-11-18; Testbed/host/TestbedLoad.ps1). A
    clone's export and imports load the host like a run, so -Execute starts them only when a run could
    start: the testbed's pause switch off (Set-TestbedPause.ps1) and the host's processors under 70 %
    busy over the last 30 seconds, the VMs runs took in the last two minutes counted. It waits for that
    before it takes anything, says what for and since when once a minute, and gives up after
    -LeaseWaitMinutes with nothing changed.

    WHAT IT HOLDS WHILE IT EXPORTS - once, however many clones. The source's lease - it waits for a
    live one somebody else holds and never takes one over - and, for the build VM, the runner's lock
    too, so that no run restores the checkpoint while it is read. It waits for both at the head of
    the testbed's queue (TestbedPool.ps1, Join-TestbedQueue -Ahead), because the minutes they are held
    end the queue it jumps, and without asking Hyper-V or polling: it sleeps until a lock's owner
    file, a lease or a ticket changes, or 30 s pass - where until 2026-11-18 it asked for the lock
    every quarter second. Both are let go as soon as the export is written: the imports touch only
    the clones' own folders.

    VM SETTINGS. Each clone gets the testbed's: AutomaticStartAction Nothing, AutomaticStopAction
    Save, automatic checkpoints off, Standard checkpoints (New-TestbedVm.ps1 sets the same), and
    rests SAVED on its checkpoint. The idle-saver watches every pool member (TestbedPool.ps1).

    EXIT CODES: 0 done, or a dry run that would proceed; 1 a step failed (it says what is left on
    the host); 4 refused.

.PARAMETER Source
    The pool's source VM: OutlookAI-Build, OutlookAI-Indexed or OutlookAI-Unindexed.

.PARAMETER Name
    The clones' VM names, each one TestbedPool.ps1 lists in the source's pool, such as
    OutlookAI-Build2. Several, comma-separated: `pwsh -File` hands 'a','b' over as one string.

.PARAMETER Execute
    Make the clones. Without it, everything is checked and nothing is written.

.PARAMETER ExportRoot
    Where the export is staged and deleted again. Default: an Export folder beside the host's
    default VM folder (E:\Hyper-V\Export here).

.PARAMETER LeaseWaitMinutes
    How long to wait for the gate (THE GATE FIRST), and then for the source's lease or the build
    runner's lock. Default 120.

.PARAMETER SelfTest
    The refusals, the paths, the checkpoint each source is cloned from, read from the committed
    testbed.json, and this file's syntax tree. No Hyper-V. Runs on the build VM with every other
    -SelfTest.

.EXAMPLE
    pwsh -File Testbed/host/New-TestbedClone.ps1 -Source OutlookAI-Build -Name OutlookAI-Build2,OutlookAI-Build3
    pwsh -File Testbed/host/New-TestbedClone.ps1 -Source OutlookAI-Build -Name OutlookAI-Build2,OutlookAI-Build3 -Execute
    pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -VMName OutlookAI-Build2    # then prove each
#>
[CmdletBinding(DefaultParameterSetName = 'Clone')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Clone')] [string] $Source,
    [Parameter(Mandatory = $true, ParameterSetName = 'Clone')] [string[]] $Name,
    [Parameter(ParameterSetName = 'Clone')] [switch] $Execute,
    [Parameter(ParameterSetName = 'Clone')] [string] $ExportRoot,
    [Parameter(ParameterSetName = 'Clone')] [int] $LeaseWaitMinutes = 120,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'TestbedPool.ps1')
# THE GATE (Q195 (e)): the pause switch and the load limit, before the source is taken.
. (Join-Path $PSScriptRoot 'TestbedLoad.ps1')

$ExchangeVmName = 'OutlookAI-Exchange'
$TimeSyncComponentId = '2497F4DE-E9FA-4204-80E4-4B75C46419C0'
$LeaseReasonPrefix = 'New-TestbedClone'
$DiskMarginBytes = 60GB

# =============================================================================================
# PURE. -SelfTest pins every one.
# =============================================================================================

# The names asked for, one each: split on commas, trimmed, empties dropped.
function Split-NameList {
    param([string[]] $Values)
    return @($Values | ForEach-Object { ([string]$_) -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# Why these clones may not be made - or '' when they may. $ExistingNames: the VMs on the host.
function Get-CloneRefusal {
    param([string] $Source, [string[]] $Names, [string[]] $ExistingNames)
    foreach ($n in @(@($Source) + @($Names))) {
        if ([string]::Equals([string]$n, $ExchangeVmName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return "$ExchangeVmName is never cloned: it holds a real Microsoft 365 sign-in and its device state (Q108, Q111), and a copy would be a second device signed in to the maintainer's tenant."
        }
    }
    $pool = Get-TestbedPoolOf $Source
    if (-not $pool -or -not [string]::Equals($TestbedPools[$pool][0], $Source, [System.StringComparison]::Ordinal)) {
        return "'$Source' is not the source of a pool - TestbedPool.ps1's sources are $((@($TestbedPools.Keys | ForEach-Object { $TestbedPools[$_][0] })) -join ', '), each spelled as listed."
    }
    if (@($Names).Count -lt 1) { return 'no clone was named.' }
    $members = @($TestbedPools[$pool] | Select-Object -Skip 1)
    $seen = @()
    foreach ($n in @($Names)) {
        if ([string]::Equals([string]$n, $Source, [System.StringComparison]::OrdinalIgnoreCase)) { return "'$n' is the source itself." }
        if ($members -cnotcontains $n) {
            return "'$n' is not a clone TestbedPool.ps1 lists in the $pool pool ($(if ($members.Count -gt 0) { $members -join ', ' } else { 'none listed' })). List it there first: a clone no runner and no idle-saver knows is a VM nothing uses and nothing saves."
        }
        if ($seen -ccontains $n) { return "'$n' is named twice." }
        $seen += $n
        foreach ($existing in @($ExistingNames)) {
            if ([string]::Equals([string]$existing, $n, [System.StringComparison]::OrdinalIgnoreCase)) {
                return "a VM named '$n' exists already. This script makes clones; it never replaces one."
            }
        }
    }
    return ''
}

# Where a clone's files go: folders of its own under the host's default roots, so no file name of
# the source's - its VHDX names, its checkpoint ids - can meet one of the clone's.
function Get-ClonePaths {
    param([string] $Name, [string] $VmRoot, [string] $VhdRoot)
    return [pscustomobject]@{ VmDir = (Join-Path $VmRoot $Name); VhdDir = (Join-Path $VhdRoot $Name) }
}

# Where the one export of a source is staged.
function Get-ExportDir {
    param([string] $Source, [string] $ExportRoot)
    return (Join-Path $ExportRoot "$Source-for-clones")
}

# The checkpoint a source is cloned from: the one its runner restores, as testbed.json records it.
function Get-SourceCheckpointName {
    param($Testbed, [string] $Source)
    $pool = Get-TestbedPoolOf $Source
    if ($pool -eq 'build') {
        $cp = [string]$Testbed.buildVm.runnerBaseCheckpoint
        if (-not $cp) { throw 'testbed.json has no buildVm.runnerBaseCheckpoint.' }
        return $cp
    }
    if ($pool -eq 'indexed' -or $pool -eq 'unindexed') {
        $record = $null
        if ($null -ne $Testbed.frozenClocks) { $record = $Testbed.frozenClocks.$Source }
        if ($null -eq $record -or -not [string]$record.checkpoint) { throw "testbed.json's frozenClocks has no checkpoint for '$Source'." }
        return [string]$record.checkpoint
    }
    throw "'$Source' is in no pool."
}

function Invoke-SelfTest {
    $script:stChecks = 0
    $script:stFailures = @()
    function Check([string] $What, $Expected, $Actual) {
        $script:stChecks++
        if ([string]$Expected -ceq [string]$Actual) { Write-Host "  OK   $What" }
        else { $script:stFailures += $What; Write-Host "  FAIL $What - expected [$Expected], got [$Actual]" }
    }
    Write-Host "New-TestbedClone.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No Hyper-V."

    Write-Host '== what may be cloned =='
    Check 'the Exchange VM never, naming why' $true ((Get-CloneRefusal 'OutlookAI-Exchange' @('OutlookAI-Exchange2') @()).Contains('Microsoft 365'))
    Check 'nor in any case' $true ((Get-CloneRefusal 'outlookai-exchange' @('OutlookAI-Build2') @()).Contains('never cloned'))
    Check 'nor under its name' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build2', 'OutlookAI-Exchange') @()).Contains('never cloned'))
    Check 'a pool source and clones it lists may' '' (Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build2', 'OutlookAI-Build3') @('OutlookAI-Build', 'OutlookAI-Indexed'))
    Check 'and an Outlook guest''s clone' '' (Get-CloneRefusal 'OutlookAI-Indexed' @('OutlookAI-Indexed2') @())
    Check 'the retired guest is no source' $true ((Get-CloneRefusal 'OutlookAI-TestVM' @('OutlookAI-TestVM2') @()).Contains('not the source of a pool'))
    Check 'a clone is no source either' $true ((Get-CloneRefusal 'OutlookAI-Build2' @('OutlookAI-Build3') @()).Contains('not the source of a pool'))
    Check 'nor is a source spelled otherwise' $true ((Get-CloneRefusal 'outlookai-build' @('OutlookAI-Build2') @()).Contains('not the source of a pool'))
    Check 'the source itself is no clone' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build') @()).Contains('the source itself'))
    Check 'a name the pool does not list is refused, saying where to list it' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build9') @()).Contains('List it there first'))
    Check 'and another pool''s clone' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Indexed2') @()).Contains('not a clone TestbedPool.ps1 lists'))
    Check 'and a padded name' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build2 ') @()) -ne '')
    Check 'and a name given twice' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build2', 'OutlookAI-Build2') @()).Contains('named twice'))
    Check 'and no name at all' $true ((Get-CloneRefusal 'OutlookAI-Build' @() @()).Contains('no clone'))
    Check 'a VM that exists is never replaced' $true ((Get-CloneRefusal 'OutlookAI-Build' @('OutlookAI-Build2') @('outlookai-build2')).Contains('never replaces'))
    Check 'names arriving as one comma-joined string are split' 'OutlookAI-Build2|OutlookAI-Build3' ((Split-NameList @('OutlookAI-Build2, OutlookAI-Build3')) -join '|')

    Write-Host '== where it goes =='
    $p = Get-ClonePaths -Name 'OutlookAI-Build2' -VmRoot 'E:\Hyper-V\VirtualMachines' -VhdRoot 'E:\Hyper-V\VirtualHardDisks'
    Check 'folders of its own, never the source''s' 'E:\Hyper-V\VirtualMachines\OutlookAI-Build2|E:\Hyper-V\VirtualHardDisks\OutlookAI-Build2' ("$($p.VmDir)|$($p.VhdDir)")
    Check 'and one export of the source for them all' 'E:\Hyper-V\Export\OutlookAI-Build-for-clones' (Get-ExportDir -Source 'OutlookAI-Build' -ExportRoot 'E:\Hyper-V\Export')

    Write-Host '== the checkpoint each source is cloned from (the committed testbed.json) =='
    $testbed = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json')) | ConvertFrom-Json
    Check 'the build VM: the base its runner restores' ([string]$testbed.buildVm.runnerBaseCheckpoint) (Get-SourceCheckpointName $testbed 'OutlookAI-Build')
    Check 'the indexed guest: its frozen checkpoint' ([string]$testbed.frozenClocks.'OutlookAI-Indexed'.checkpoint) (Get-SourceCheckpointName $testbed 'OutlookAI-Indexed')
    Check 'the unindexed guest: its frozen checkpoint' ([string]$testbed.frozenClocks.'OutlookAI-Unindexed'.checkpoint) (Get-SourceCheckpointName $testbed 'OutlookAI-Unindexed')
    $threw = ''
    try { [void](Get-SourceCheckpointName $testbed 'OutlookAI-Exchange') } catch { $threw = $_.Exception.Message }
    Check 'the Exchange VM has none' $true ($threw -ne '')

    Write-Host '== this file =='
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $commands = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))
    $names = @($commands | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
    foreach ($forbidden in @('Remove-VM', 'Stop-VM', 'Restart-VM', 'Start-VM', 'Restore-VMSnapshot', 'Remove-VMSnapshot', 'Remove-VMCheckpoint', 'Remove-VMHardDiskDrive', 'Set-VMProcessor', 'Set-VMMemory', 'Stop-Process')) {
        Check "it never calls $forbidden" $false ($names -contains $forbidden)
    }
    $removes = @($commands | Where-Object { $_.GetCommandName() -eq 'Remove-Item' } | ForEach-Object { $_.Extent.Text })
    Check 'it deletes nothing but its own export' '' (@($removes | Where-Object { $_ -notmatch '^Remove-Item -LiteralPath \$exportDir ' }) -join ' ; ')
    $renames = @($commands | Where-Object { $_.GetCommandName() -eq 'Rename-VM' } | ForEach-Object { $_.Extent.Text })
    Check 'it renames only the VM in an import''s report, before it is registered' '1|1' ("{0}|{1}" -f @($renames | Where-Object { $_ -match '^Rename-VM -VM \$report\.VM ' }).Count, $renames.Count)
    # The gate (Q195 (e)) and the waits (2026-11-18), looked for in the code after this function, whose own
    # lines name the same text.
    $selfTestAst = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SelfTest' }, $false))
    $runCode = ''
    if ($selfTestAst.Count -eq 1) { $runCode = [System.IO.File]::ReadAllText($PSCommandPath).Substring($selfTestAst[0].Extent.EndOffset) }
    $gateAt = $runCode.IndexOf('$gate = Wait-TestbedGate -Run $leaseReason')
    Check 'the gate is waited for before the source is taken - the runner''s lock and the lease' $true ($gateAt -gt 0 -and $gateAt -lt $runCode.IndexOf('[System.IO.File]::Open($lockPath') -and $gateAt -lt $runCode.IndexOf('Enter-TestbedLeaseClaim -VMName $Source'))
    Check 'and one that does not open ends the run with nothing changed' $true ($runCode.IndexOf('if (-not $gate.Go) {') -gt $gateAt -and $runCode.IndexOf('if (-not $gate.Go) {') -lt $runCode.IndexOf('Export-VMSnapshot'))
    $queuedAt = $runCode.IndexOf('$ticket = Join-TestbedQueue -Run $leaseReason')
    Check 'it waits for its source at the head of the queue, from before it asks for the lock or the lease' $true ($queuedAt -gt 0 -and $runCode.Contains('-Candidates @($Source) -Ahead') -and $queuedAt -lt $runCode.IndexOf('[System.IO.File]::Open($lockPath') -and $queuedAt -lt $runCode.IndexOf('Enter-TestbedLeaseClaim -VMName $Source'))
    Check 'it sleeps on the testbed''s files, never polls, and leaves the queue on every path' $true (([regex]::Matches($runCode, [regex]::Escape('[void](Wait-TestbedChange)'))).Count -eq 2 -and -not $runCode.Contains('Start-Sleep') -and $runCode.Contains('Exit-TestbedQueue $ticket'))

    Write-Host ''
    Write-Host "$($script:stChecks) check(s), $($script:stFailures.Count) failure(s)."
    Write-Host 'NOT COVERED HERE, because it needs Hyper-V: the export, the imports, the settings and the checkpoints. Clones made with -Execute, and the runs that prove them, cover them.'
    if ($script:stFailures.Count -gt 0) { return 1 }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE CLONES.
# =============================================================================================
. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
function Stop-Refused([string] $Why) { Say "REFUSED: $Why"; exit 4 }

function Get-TimeSyncEnabled($Services) {
    $ts = @($Services | Where-Object { ([string]$_.Id).ToUpperInvariant().EndsWith($TimeSyncComponentId) -or $_.Name -eq 'Time Synchronization' })
    if ($ts.Count -lt 1) { return $null }
    return [bool]$ts[0].Enabled
}

$names = @(Split-NameList $Name)
$existing = @(Get-VM -ErrorAction Stop | ForEach-Object { $_.Name })
$refusal = Get-CloneRefusal -Source $Source -Names $names -ExistingNames $existing
if ($refusal) { Stop-Refused $refusal }
$pool = Get-TestbedPoolOf $Source
$sourceVm = @(Get-VM -Name $Source -ErrorAction SilentlyContinue)
if ($sourceVm.Count -ne 1) { Stop-Refused "expected one VM named '$Source' on this host, found $($sourceVm.Count)." }
$sourceVm = $sourceVm[0]

$testbed = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json')) | ConvertFrom-Json
$checkpointName = Get-SourceCheckpointName $testbed $Source
$checkpoint = @(Get-VMSnapshot -VM $sourceVm | Where-Object { $_.Name -ceq $checkpointName })
if ($checkpoint.Count -ne 1) { Stop-Refused "'$Source' has $($checkpoint.Count) checkpoint(s) named '$checkpointName', not one." }
$checkpoint = $checkpoint[0]
# Only a checkpoint taken running carries memory - and with it a frozen guest's clock. Hyper-V says so
# on the checkpoint itself: a running one reads Saved (Set-GuestClockFrozen.ps1 reads it the same way).
if ([string]$checkpoint.State -ne 'Saved') { Stop-Refused "'$checkpointName' holds no running guest (its state is $($checkpoint.State)); a clone of it would cold-boot, and a frozen guest's would come back at the real date." }
$sourceTimeSync = Get-TimeSyncEnabled (@($checkpoint | Get-VMIntegrationService))

$vmHost = Get-VMHost
if (-not $ExportRoot) { $ExportRoot = Join-Path (Split-Path -Parent ([string]$vmHost.VirtualMachinePath)) 'Export' }
$exportDir = Get-ExportDir -Source $Source -ExportRoot $ExportRoot
$plans = @($names | ForEach-Object { [pscustomobject]@{ Name = $_; Paths = (Get-ClonePaths -Name $_ -VmRoot ([string]$vmHost.VirtualMachinePath) -VhdRoot ([string]$vmHost.VirtualHardDiskPath)) } })
foreach ($dir in @(@($exportDir) + @($plans | ForEach-Object { $_.Paths.VmDir; $_.Paths.VhdDir }))) {
    if (Test-Path -LiteralPath $dir) { Stop-Refused "'$dir' exists already. Clones get folders of their own; remove what is there by hand once you know what it is." }
}

# What it will take: the checkpoint's disk chain - a merged VHDX is at most that - and its saved memory,
# once for the export and once for every clone.
$chainBytes = 0L
foreach ($disk in @(Get-VMHardDiskDrive -VMSnapshot $checkpoint)) {
    $path = [string]$disk.Path
    while ($path) {
        $vhd = Get-VHD -Path $path
        $chainBytes += [long]$vhd.FileSize
        $path = [string]$vhd.ParentPath
    }
}
$eachBytes = $chainBytes + [long]$checkpoint.MemoryStartup
$needBytes = $eachBytes * ($plans.Count + 1) + $DiskMarginBytes
$drive = New-Object System.IO.DriveInfo([System.IO.Path]::GetPathRoot([string]$vmHost.VirtualHardDiskPath))
$freeBytes = [long]$drive.AvailableFreeSpace
Say "== New-TestbedClone: $Source -> $($names -join ', ') ($(if ($Execute) { 'EXECUTE' } else { 'DRY RUN' })) =="
Say ("  from       '{0}' of {1}, taken {2:yyyy-MM-dd HH:mm} (local), time sync {3}" -f $checkpointName, $Source, $checkpoint.CreationTime, $(if ($sourceTimeSync) { 'ON' } elseif ($null -eq $sourceTimeSync) { 'not found' } else { 'OFF (frozen)' }))
Say ("  disk       the chain {0:N1} GB and its memory {1:N1} GB, for the export and each of {2} clone(s), plus {3:N0} GB: {4:N1} GB; {5} has {6:N1} GB free" -f ($chainBytes / 1GB), ($checkpoint.MemoryStartup / 1GB), $plans.Count, ($DiskMarginBytes / 1GB), ($needBytes / 1GB), $drive.Name, ($freeBytes / 1GB))
Say "  export     $exportDir"
foreach ($plan in $plans) { Say "  $($plan.Name)  $($plan.Paths.VmDir) ; $($plan.Paths.VhdDir)" }
if ($freeBytes -lt $needBytes) { Stop-Refused ("{0} has {1:N1} GB free, under the {2:N1} GB this needs." -f $drive.Name, ($freeBytes / 1GB), ($needBytes / 1GB)) }
if (-not $Execute) {
    Say 'DRY RUN. Nothing was changed. With -Execute this would:'
    Say "  0. wait, holding nothing, for the testbed's gate: the pause switch off and the host's processors under $(Format-Percent $TestbedLoadLimitPercent) busy over the last $TestbedLoadWindowSeconds s (TestbedLoad.ps1)"
    Say "  1. take $Source's lease$(if ($pool -eq 'build') { ' and the build runner''s lock' }) (waiting, never taking one over), export '$checkpointName' once, let both go"
    Say '  2. for each clone: import the export as a copy with a new id, renamed before it is registered, into its folders'
    Say '  3. set AutomaticStartAction Nothing, AutomaticStopAction Save, automatic checkpoints off, Standard checkpoints'
    Say "  4. check it is saved, disconnected, and keeps the checkpoint's time sync; take '$checkpointName' of it"
    Say '  5. delete the export'
    exit 0
}

$leaseReason = "$LeaseReasonPrefix $($names -join ',')"
# ---- 0. the gate (THE GATE FIRST): waits holding nothing, before the source is taken.
Say "gate: the pause switch, then the host's load over the last $TestbedLoadWindowSeconds s (Testbed/host/TestbedLoad.ps1)"
$gate = Wait-TestbedGate -Run $leaseReason -What "New-TestbedClone.ps1 -Source $Source -Name $($names -join ',')" -TimeoutMinutes $LeaseWaitMinutes -Say { param($line) Say $line }
if (-not $gate.Go) { Say "FAILED: the gate did not open in $LeaseWaitMinutes minutes - $($gate.Why). Nothing was changed."; exit 1 }
Say "gate: $(Format-TestbedGateOpen -Look $gate.Look -PendingPercent $gate.Pending.Percent -PendingCount $gate.Pending.Count)"
$lock = $null
$leaseTaken = $false
$exportDone = $false
$ticket = $null
try {
    # ---- 1. hold the source, export once, let it go - waiting at the head of the testbed's queue
    $deadline = (Get-Date).AddMinutes($LeaseWaitMinutes)
    $ticket = Join-TestbedQueue -Run $leaseReason -What "New-TestbedClone.ps1 -Source $Source -Name $($names -join ',')" -Candidates @($Source) -Ahead
    if ($pool -eq 'build') {
        $lockPath = Join-Path (Get-BuildRunnerDirectory) (Get-BuildRunnerLockNames $Source).Lock
        $said = $false
        while ($null -eq $lock) {
            try { $lock = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None) }
            catch [System.IO.IOException] {
                if (-not $said) { Say "waiting: a run holds $Source's runner lock"; $said = $true }
                if ((Get-Date) -ge $deadline) { throw "$Source's runner lock stayed held for $LeaseWaitMinutes minutes" }
                # Asked again as soon as a lock's owner file goes - a run lets go of its lock right
                # after it - or after 30 s: no poll in between.
                [void](Wait-TestbedChange)
            }
        }
    }
    $said = $false
    while (-not $leaseTaken) {
        $lease = $null
        $claim = Enter-TestbedLeaseClaim -VMName $Source
        try {
            $lease = Get-TestbedLease -VMName $Source
            # A build run's own lease outlives its lock by moments, and the lock is ours now.
            $runs = ($pool -eq 'build' -and $null -ne $lease -and ([string]$lease.reason).StartsWith('Invoke-TestsOnBuildVm'))
            if ($null -eq $lease -or $runs -or ([string]$lease.reason) -ceq $leaseReason) {
                & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $Source -Minutes 90 -Reason $leaseReason | Out-Null
                $leaseTaken = $true
            }
        }
        finally { $claim.Dispose() }
        if (-not $leaseTaken) {
            if (-not $said) { Say "waiting: $Source is leased until $($lease.expiresUtc) - $($lease.reason)"; $said = $true }
            if ((Get-Date) -ge $deadline) { throw "$Source stayed leased for $LeaseWaitMinutes minutes" }
            [void](Wait-TestbedChange)
        }
    }
    Say "holding $Source$(if ($null -ne $lock) { ' (lease and runner lock)' } else { ' (lease)' }); exporting '$checkpointName'"
    $t0 = Get-Date
    New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
    Export-VMSnapshot -VMSnapshot $checkpoint -Path $exportDir
    $exportDone = $true
    Say ("  exported in {0:N0} s" -f ((Get-Date) - $t0).TotalSeconds)
}
catch {
    Say "FAILED: $($_.Exception.Message)"
    if (Test-Path -LiteralPath $exportDir) { Say "  a partial export is at $exportDir - delete it by hand" }
    $exportDone = $false
}
finally {
    if ($leaseTaken) {
        $l = Get-TestbedLease -VMName $Source
        if ($null -ne $l -and ([string]$l.reason) -ceq $leaseReason) { & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $Source -Release | Out-Null }
    }
    if ($null -ne $lock) { $lock.Dispose() }
    Exit-TestbedQueue $ticket
    Say "let $Source go"
}
if (-not $exportDone) { exit 1 }

$vmcx = @(Get-ChildItem -LiteralPath $exportDir -Recurse -Filter '*.vmcx' -File | Where-Object { $_.Directory.Name -eq 'Virtual Machines' })
if ($vmcx.Count -ne 1) { Say "FAILED: the export holds $($vmcx.Count) VM configuration(s), not one - it is at $exportDir"; exit 1 }
foreach ($plan in $plans) {
    $cloneName = $plan.Name
    try {
        # ---- 2. import as a copy with a new id, renamed before it is registered
        $t0 = Get-Date
        $report = Compare-VM -Path $vmcx[0].FullName -Copy -GenerateNewId -VirtualMachinePath $plan.Paths.VmDir -SnapshotFilePath $plan.Paths.VmDir -SmartPagingFilePath $plan.Paths.VmDir -VhdDestinationPath $plan.Paths.VhdDir
        $problems = @($report.Incompatibilities | ForEach-Object { "$($_.MessageId): $($_.Message)" })
        if ($problems.Count -gt 0) { throw ("the import reports $($problems.Count) incompatibilit(y/ies): " + ($problems -join ' | ')) }
        Rename-VM -VM $report.VM -NewName $cloneName
        $imported = Import-VM -CompatibilityReport $report
        Say ("{0}: imported in {1:N0} s, id {2}" -f $imported.Name, ((Get-Date) - $t0).TotalSeconds, $imported.Id)
        $clone = @(Get-VM -Name $cloneName)
        if ($clone.Count -ne 1) { throw "Hyper-V lists $($clone.Count) VMs named '$cloneName' after the import" }
        $clone = $clone[0]
        if ($clone.Id -eq $sourceVm.Id) { throw 'the clone has the source''s id' }
        if (@(Get-VM -Name $Source).Count -ne 1) { throw "Hyper-V no longer lists exactly one '$Source'" }

        # ---- 3. the testbed's settings
        Set-VM -VM $clone -AutomaticStartAction Nothing -AutomaticStopAction Save -AutomaticCheckpointsEnabled $false -CheckpointType Standard
        foreach ($nic in @(Get-VMNetworkAdapter -VM $clone | Where-Object { $_.SwitchName })) {
            Disconnect-VMNetworkAdapter -VMNetworkAdapter $nic
            Say "  disconnected its adapter '$($nic.Name)' from '$($nic.SwitchName)'"
        }

        # ---- 4. checked and checkpointed
        $clone = Get-VM -Name $cloneName
        if ([string]$clone.State -ne 'Saved') { throw "it is $($clone.State), not Saved - the export did not carry the checkpoint's memory, and a start would cold-boot it" }
        $cloneTimeSync = Get-TimeSyncEnabled (@(Get-VMIntegrationService -VM $clone))
        if ($cloneTimeSync -ne $sourceTimeSync) { throw "its time sync reads $cloneTimeSync, the checkpoint's $sourceTimeSync" }
        if (@(Get-VMNetworkAdapter -VM $clone | Where-Object { $_.SwitchName }).Count -gt 0) { throw 'an adapter of it is still connected' }
        Checkpoint-VM -VM $clone -SnapshotName $checkpointName
        $taken = @(Get-VMSnapshot -VM $clone | Where-Object { $_.Name -ceq $checkpointName })
        if ($taken.Count -ne 1) { throw "it has $($taken.Count) checkpoint(s) named '$checkpointName' after Checkpoint-VM" }
        $inCheckpoint = Get-TimeSyncEnabled (@($taken[0] | Get-VMIntegrationService))
        if ($inCheckpoint -ne $sourceTimeSync) { throw "time sync in its '$checkpointName' reads $inCheckpoint, the source's $sourceTimeSync" }
        $clone = Get-VM -Name $cloneName
        $disks = @(Get-VMHardDiskDrive -VM $clone | ForEach-Object { $_.Path })
        Say ("  {0}, {1} CPU, {2:N0} GB {3} memory, AutomaticStart {4}, AutomaticStop {5}, automatic checkpoints {6}, time sync {7}" -f $clone.State, $clone.ProcessorCount, ($clone.MemoryStartup / 1GB), $(if ($clone.DynamicMemoryEnabled) { 'dynamic' } else { 'static' }), $clone.AutomaticStartAction, $clone.AutomaticStopAction, $clone.AutomaticCheckpointsEnabled, $(if ($cloneTimeSync) { 'ON' } else { 'OFF' }))
        Say "  disk $($disks -join ', ')"
        Say "  checkpoint '$checkpointName' taken $($taken[0].CreationTime.ToString('yyyy-MM-dd HH:mm:ss')) (local)"
    }
    catch {
        Say "FAILED on ${cloneName}: $($_.Exception.Message)"
        Say "  the export is still at $exportDir - delete it by hand once the clones are sorted out"
        $now = @(Get-VM -Name $cloneName -ErrorAction SilentlyContinue)
        if ($now.Count -gt 0) { Say "  a VM named '$cloneName' is registered ($($now[0].State)); its files are under $($plan.Paths.VmDir) and $($plan.Paths.VhdDir). Remove it by hand (Remove-VM, then the two folders) before running this again." }
        exit 1
    }
}

# ---- 5. the export is gone
Remove-Item -LiteralPath $exportDir -Recurse -Force
Say ''
Say "== DONE: $($names -join ', ') rest saved on '$checkpointName', copies of $Source's; the export deleted =="
Say 'Next: prove each with a run that names it, then record it in Testbed/testbed.json, vmPools:'
foreach ($plan in $plans) {
    if ($pool -eq 'build') { Say "  pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -VMName $($plan.Name)" }
    else { Say "  pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName $($plan.Name)" }
}
foreach ($plan in $plans) {
    Say ("  {{ `"vmName`": `"{0}`", `"fromCheckpoint`": `"{1}`", `"checkpoint`": `"{1}`", `"madeUtc`": `"{2}`", `"provenBy`": `"<the run id>`" }}" -f $plan.Name, $checkpointName, [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))
}
exit 0
