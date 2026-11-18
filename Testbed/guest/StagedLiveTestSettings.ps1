#Requires -Version 5.1
<#
.SYNOPSIS
    Dot-sourced by the guest scripts that need a mailbox the run's live-test settings name: the staged
    settings file's path, and the one way to read a field of it. Defines functions; runs nothing.

.DESCRIPTION
    No real mailbox's address is written in this repository (Q189 (b), decided by the maintainer
    2026-11-17). On a test guest the live runner stages the gitignored live-test settings beside the
    suite - Testbed/host/Invoke-LiveTierOnGuest.ps1 copies the file it rendered - so a guest script that
    needs the hub, a bystander or the indexed stores reads them from that file rather than naming them.

    One copy for every such script (AGENTS.md, Shared code, Q154): Reset-OutlookOfflineStore.ps1 and the
    three Exchange VM probes, Measure-ExchangeIndexIds.ps1, Measure-OfflineWindow.ps1 and
    Measure-SignatureRoaming.ps1. Reset-OutlookOfflineStore.ps1's -SelfTest holds these functions; the
    live runner copies every file of Testbed/guest to the guest, this one included.
#>

# Where the live runner stages the settings on a test guest.
$StagedSettingsPath = 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json'

# Every non-empty string of one field of the settings' JSON text - the field's value, or each entry of a
# list. Pure.
function Get-SettingsStrings {
    param([string] $Json, [string] $Field)
    $settings = $Json | ConvertFrom-Json
    if (-not ($settings.PSObject.Properties.Name -contains $Field)) { return @() }
    return @($settings.$Field | Where-Object { $_ -is [string] -and $_ })
}

# The first of them - testHubStoreDisplayName, or the first bystander - or $null. Pure.
function Get-SettingsValue {
    param([string] $Json, [string] $Field)
    $values = @(Get-SettingsStrings $Json $Field)
    if ($values.Count -eq 0) { return $null }
    return [string]$values[0]
}

# The same, read from the staged file at $Path; $null when there is no file.
function Get-StagedSettingsValue {
    param([string] $Path, [string] $Field)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }
    return (Get-SettingsValue (Get-Content -LiteralPath $Path -Raw) $Field)
}
