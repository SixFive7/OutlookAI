#Requires -Version 5.1
<#
.SYNOPSIS
    Save every testbed VM that nobody is using, so the host gets its RAM and CPU back.

.DESCRIPTION
    SAVED, NOT PAUSED. `Suspend-VM` freezes a VM but keeps its memory resident - the host gets
    its CPU back and none of its RAM. `Save-VM` writes the guest's memory to disk and releases
    the RAM entirely, and resuming is still far faster than a boot because the guest never shut
    down: Outlook is still running, the profile is still open, the index service is still
    warm. That combination - resources back, fast restart - is the whole requirement, and it is
    the reason this script saves rather than pauses. The cost is disk: a saved VM's memory file
    is roughly its assigned RAM.

    WHAT COUNTS AS "IN USE" IS DECLARED, NOT MEASURED. See Set-TestbedLease.ps1 for why guest
    CPU is the wrong signal: a live tier run spends much of its 27 minutes waiting on Outlook,
    so an idle-looking guest mid-run is ordinary, and saving it there suspends a COM call and
    surfaces later as a test that timed out on a machine that looks fine.

    A VM is saved only when ALL of these hold:
      - it is one of the testbed VMs (this never touches a VM it was not told about),
      - it is Running,
      - it holds no live lease,
      - it has been up longer than -MinimumUptimeMinutes, so a VM that has just been started
        for work that has not taken its lease yet is not immediately put back to sleep.

.PARAMETER VMName
    The testbed VMs. Defaults to the names the testbed uses. Anything not named here is ignored
    entirely - this script must never be able to save a VM that is not part of the testbed.

    THIS ONE KEEPS ITS DEFAULT, AND IT IS THE ONLY ONE IN Testbed/host/ THAT DOES. Every other
    script there now requires -VMName / -Name, because THREE MACHINES COEXIST during the
    changeover (OutlookAI-Indexed, OutlookAI-Unindexed and the outgoing OutlookAI-TestVM) and a
    default that silently picks ONE OF THREE is the exact shape of mistake this testbed keeps
    making. Two things make this parameter different:

      - It is an ALLOWLIST, not a target. It does not pick a machine to act on; it bounds the
        set this script is permitted to touch at all. Naming all three is the whole intent -
        "save any testbed VM nobody is using" - not an unstated guess at one of them.
      - Its failure direction is the safe one. A wrong or stale entry makes this script do LESS:
        it skips a VM and the host keeps holding RAM, which is visible in Get-VM and fixed by
        saving by hand. A wrong default in Copy-ToGuest or New-TestbedVm acts on a machine the
        operator was not looking at, which is not visible at all.

    Making it mandatory would also break the scheduled task outright. Register-IdleSaveTask.ps1
    invokes this file with -NonInteractive and no arguments, so a mandatory parameter cannot
    prompt - it throws, every fifteen minutes, for ever, and the symptom is a host that never
    reclaims its RAM with nothing anywhere saying why. That is strictly worse than the mistake
    the mandatory rule exists to prevent.

    KEEP THIS LIST IN STEP WITH THE GUESTS THAT EXIST. Drop OutlookAI-TestVM from it once the
    old guest is gone, and add any further guest the day it is built: a testbed VM missing from
    this list is simply never saved.

.PARAMETER MinimumUptimeMinutes
    Grace period after a VM starts, before it becomes eligible to be saved.

.PARAMETER WhatIf
    Report what would be saved and change nothing.

.EXAMPLE
    Testbed/host/Invoke-TestbedIdleSave.ps1 -WhatIf
    Testbed/host/Register-IdleSaveTask.ps1        # to run it on a schedule
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # An allowlist, not a target - see .PARAMETER VMName. This is the one script in Testbed/host/
    # that keeps a default, because naming every known guest IS the intent here, and a name that
    # is wrong or stale makes it skip a VM rather than act on the wrong one.
    [string[]] $VMName = @('OutlookAI-Indexed', 'OutlookAI-Unindexed', 'OutlookAI-TestVM'),
    [int] $MinimumUptimeMinutes = 10
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')

$logDir = Join-Path $env:SystemDrive 'OutlookAI-Testbed'
$log = Join-Path $logDir 'idle-save.log'
function Write-Line {
    param([string] $Text)
    $line = "{0}  {1}" -f ([DateTime]::UtcNow.ToString('u')), $Text
    Write-Output $line
    try { Add-Content -LiteralPath $log -Value $line -Encoding UTF8 } catch { }
}

# Prove Hyper-V is reachable BEFORE the loop. Without this the per-VM
# '-ErrorAction SilentlyContinue' swallows a permissions failure exactly as it swallows a VM
# that lives on another host, and a run without Hyper-V access would skip every VM, print
# nothing, exit 0, and look like a machine where nothing was ever idle. Measured: that is
# precisely what an unelevated dry run did - the cause being the account's Hyper-V group
# membership, not the elevation, which this task deliberately does not have.
try {
    Get-VM -ErrorAction Stop | Out-Null
} catch {
    Write-Line "CANNOT QUERY HYPER-V - nothing was checked: $($_.Exception.Message)"
    Write-Line "Save-VM needs local 'Hyper-V Administrators' membership, NOT elevation - and as"
    Write-Line "registered this runs as the invoking user at ordinary privilege, by design."
    Write-Line "Group membership is fixed when a logon session is created, so a membership added"
    Write-Line "since the last logon is not in this token yet. Log on again, then re-run."
    exit 1
}

foreach ($name in $VMName) {
    $vm = Get-VM -Name $name -ErrorAction SilentlyContinue
    if (-not $vm) { continue }                       # genuinely not on this host

    if ($vm.State -ne 'Running') {
        Write-Verbose "$name is $($vm.State); nothing to do."
        continue
    }

    $lease = Get-TestbedLease -VMName $name
    if ($lease) {
        Write-Verbose "$name is leased until $($lease.expiresUtc) ($($lease.reason)); leaving it alone."
        continue
    }

    if ($vm.Uptime.TotalMinutes -lt $MinimumUptimeMinutes) {
        # A VM that has just come up may belong to work that has not taken its lease yet.
        # Saving it here would fight whoever started it.
        Write-Verbose "$name has been up $([math]::Round($vm.Uptime.TotalMinutes,1)) min, under the $MinimumUptimeMinutes min grace; leaving it alone."
        continue
    }

    if ($PSCmdlet.ShouldProcess($name, 'Save-VM')) {
        try {
            Save-VM -Name $name -ErrorAction Stop
            Write-Line "saved $name (unleased, up $([math]::Round($vm.Uptime.TotalHours,1))h, $([math]::Round($vm.MemoryAssigned/1GB,1)) GB reclaimed)"
        } catch {
            Write-Line "FAILED to save ${name}: $($_.Exception.Message)"
        }
    } else {
        Write-Line "WOULD save $name (unleased, up $([math]::Round($vm.Uptime.TotalHours,1))h)"
    }
}
