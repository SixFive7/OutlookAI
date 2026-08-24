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
# that lives on another host, and an unelevated run of this task would skip every VM, print
# nothing, exit 0, and look like a machine where nothing was ever idle. Measured: that is
# precisely what an unelevated dry run did.
try {
    Get-VM -ErrorAction Stop | Out-Null
} catch {
    Write-Line "CANNOT QUERY HYPER-V - nothing was checked: $($_.Exception.Message)"
    Write-Line "This task must run elevated; as registered it runs as SYSTEM, which is."
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
