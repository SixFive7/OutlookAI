#Requires -Version 5.1
<#
.SYNOPSIS
    Register the scheduled task that saves idle testbed VMs. Needs Hyper-V group membership, not elevation.

.DESCRIPTION
    The task runs as the invoking user at ordinary privilege - NOT as SYSTEM and NOT elevated.
    It relies on the account being a member of the local Hyper-V Administrators group, which is
    what lets Save-VM work without administrator rights. Register it once that membership is in
    the account's token (group membership is fixed when a logon session is created, so it takes
    effect at the next logon).

    Deliberately not elevated: a saver that only ever calls Save-VM does not need administrator
    rights, and a scheduled task holding them is a standing capability that outlives the reason
    it was created.

    Every fifteen minutes is deliberate. More often buys nothing - a VM that has been idle for
    an hour is not more idle at 5-minute granularity - and it costs a wake on a machine that
    might be asleep. Less often leaves several gigabytes parked for longer than necessary.

    THIS TASK ONLY EVER SAVES. It never starts, never stops, never checkpoints and never
    deletes. If it goes wrong the worst case is a VM saved while someone was using it, which is
    visible immediately and fixed by resuming.

.PARAMETER IntervalMinutes
    How often to check.

.PARAMETER Unregister
    Remove the task instead of creating it.

.EXAMPLE
    Testbed/host/Register-IdleSaveTask.ps1
    Testbed/host/Register-IdleSaveTask.ps1 -Unregister
#>
[CmdletBinding()]
param(
    [int] $IntervalMinutes = 15,
    [switch] $Unregister
)

$ErrorActionPreference = 'Stop'
$taskName = 'OutlookAI-TestbedIdleSave'

# Fail early and clearly if Hyper-V is not reachable as this account - otherwise the task
# registers happily and then does nothing every fifteen minutes for ever.
try {
    Get-VM -ErrorAction Stop | Out-Null
} catch {
    throw "Cannot query Hyper-V as $env:USERNAME. Add the account to the local 'Hyper-V Administrators' group and log on again, then re-run this. (Group membership is fixed when a logon session is created.)"
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Output "Removed $taskName (if it existed)."
    return
}

$script = Join-Path $PSScriptRoot 'Invoke-TestbedIdleSave.ps1'
if (-not (Test-Path -LiteralPath $script)) {
    throw "Cannot find the saver beside this script at $script."
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $script)

# The invoking user, ordinary privilege. Hyper-V group membership is what makes Save-VM work;
# no elevation is involved, and the task has no rights beyond what the account already has.
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

$trigger = New-ScheduledTaskTrigger -Once -At ([DateTime]::Now.AddMinutes(1)) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -DontStopOnIdleEnd

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal `
    -Trigger $trigger -Settings $settings -Description 'Saves idle OutlookAI testbed VMs so the host reclaims RAM and CPU. Never starts or deletes a VM.' | Out-Null

Write-Output "Registered ${taskName}: every $IntervalMinutes minutes, as $env:USERNAME at ordinary privilege."
Write-Output "It saves a testbed VM only when it is Running, unleased, and up past the grace period."
