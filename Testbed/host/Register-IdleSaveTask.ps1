#Requires -Version 5.1
<#
.SYNOPSIS
    Register the scheduled task that saves idle testbed VMs. Must be run elevated.

.DESCRIPTION
    Hyper-V cmdlets need administrator rights, so the task runs as SYSTEM. That is also why it
    runs in session 0 with no desktop: it can never draw a window or take focus, which matters
    on a machine somebody games on.

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

$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "This must run elevated: registering a SYSTEM task and driving Hyper-V both need administrator rights."
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

# SYSTEM, session 0: no desktop, so nothing can be drawn and nothing can steal focus.
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

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

Write-Output "Registered ${taskName}: every $IntervalMinutes minutes, as SYSTEM, session 0."
Write-Output "It saves a testbed VM only when it is Running, unleased, and up past the grace period."
