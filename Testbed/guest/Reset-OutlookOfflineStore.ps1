#Requires -Version 5.1
<#
.SYNOPSIS
    Sets aside the Exchange VM's Outlook data files (.ost) so that Outlook downloads its cached
    mailboxes afresh - the repair for "Outlook is using an old copy of your Outlook data file (.ost)",
    which every restore of an Exchange VM checkpoint brings with it.

.DESCRIPTION
    RUN ON THE GUEST, with Outlook CLOSED (Windows PowerShell 5.1), as the autologon user whose
    profile holds the data files - over PowerShell Direct, after Testbed/host/Restart-Guest.ps1 has
    quit Outlook and restarted Windows. It refuses while OUTLOOK.EXE runs.

    WHY IT EXISTS (measured 2026-10-08 on OutlookAI-Exchange). A checkpoint restores the guest's disk,
    the .ost files included, and nothing on the Exchange server: the mailboxes stay as the last run
    left them. Once a write run had changed the test account on the server, the next run's restore
    put back an .ost older than the server's state, and at the next start of Outlook - the
    disconnect-recovery test restarts it - Outlook refused the store: "Outlook is using an old copy of
    your Outlook data file (.ost). Exit Outlook, delete the file, and restart Outlook." The count
    tripwire's post-run census of the hub then failed with that message. This does what the message
    says, the documented way: the files are RENAMED aside (<name>.ost.stale-<UTC stamp>), never
    deleted, and Outlook makes new ones at its next start and downloads both mailboxes into them -
    small mailboxes, a minute or two. Nothing in either mailbox is touched: an .ost is a local copy of
    the server's mailbox, and the server's copy is the one that counts.

    The live runner no longer restores the Exchange VM's checkpoint before or after a run (its state is
    carried forward, so its .ost files stay current); this is for the time a checkpoint IS restored.

    Writes only <data file>.stale-<stamp> renames in %LOCALAPPDATA%\Microsoft\Outlook. -Execute to
    rename; without it, says what it would rename.

.PARAMETER Execute
    Rename the files. Without it, a dry run.

.PARAMETER Folder
    Where the data files are. Default %LOCALAPPDATA%\Microsoft\Outlook of the user running it.

.PARAMETER SelfTest
    The pure part: which files a folder listing makes it rename. No guest, no Outlook.
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')] [switch] $Execute,
    [Parameter(ParameterSetName = 'Run')] [string] $Folder,
    [Parameter(ParameterSetName = 'Run')] [string[]] $ExpectedUser = @('vmadmin'),
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)
$ErrorActionPreference = 'Stop'

# The files to set aside: every *.ost in the folder (an Exchange account's cached mailbox), and nothing
# else - not a .pst, not a .nst, not an .ost already set aside. Pure.
function Get-StoresToSetAside {
    param([string[]] $Names)
    return @($Names | Where-Object { $_ -and $_ -match '(?i)\.ost$' } | Sort-Object)
}

function Get-SetAsideName {
    param([string] $Name, [DateTime] $WhenUtc)
    return $Name + '.stale-' + $WhenUtc.ToUniversalTime().ToString('yyyyMMddTHHmmssZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

if ($SelfTest) {
    Write-Host "Reset-OutlookOfflineStore.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No guest, no Outlook."
    $script:fails = 0
    function Check([string] $Name, $Expected, $Actual) {
        if ("$Expected" -ceq "$Actual") { Write-Host "  PASS  $Name" } else { $script:fails++; Write-Host "  FAIL  $Name - expected '$Expected', got '$Actual'" }
    }
    $names = @('mailbox@example.invalid.ost', 'b@example.com.ost', 'archive.pst', 'old.ost.stale-20261008T120000Z', 'x.nst', '', 'UPPER.OST')
    Check 'only the .ost files are set aside, sorted' 'b@example.com.ost|mailbox@example.invalid.ost|UPPER.OST' ((Get-StoresToSetAside $names) -join '|')
    Check 'none in an empty folder' 0 @(Get-StoresToSetAside @()).Count
    Check 'the aside name carries the UTC stamp' 'a.ost.stale-20261008T153000Z' (Get-SetAsideName 'a.ost' ([DateTime]::SpecifyKind([DateTime]::new(2026, 10, 8, 15, 30, 0), [DateTimeKind]::Utc)))
    Write-Host ''
    if ($script:fails -gt 0) { Write-Host "$script:fails FAILED"; exit 1 }
    Write-Host 'All passed.'
    exit 0
}

# THE GUARD, FIRST: this renames an Outlook profile's data files, so only on a testbed guest.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Host 'REFUSING: OUTLOOK.EXE is running - it holds its data files open. Quit it with Testbed/host/Restart-Guest.ps1 first.'
    exit 4
}
if (-not $Folder) { $Folder = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook' }
if (-not (Test-Path -LiteralPath $Folder)) { Write-Host "NOTHING: no folder $Folder"; exit 0 }

$names = @(Get-ChildItem -LiteralPath $Folder -File | ForEach-Object { $_.Name })
$targets = @(Get-StoresToSetAside $names)
if ($targets.Count -eq 0) { Write-Host "NOTHING: no .ost in $Folder"; exit 0 }
$now = [DateTime]::UtcNow
foreach ($name in $targets) {
    $aside = Get-SetAsideName $name $now
    $size = (Get-Item -LiteralPath (Join-Path $Folder $name)).Length
    if ($Execute) {
        Rename-Item -LiteralPath (Join-Path $Folder $name) -NewName $aside
        Write-Host ("SET ASIDE {0} ({1:N0} KB) -> {2}" -f $name, ($size / 1KB), $aside)
    }
    else {
        Write-Host ("WOULD SET ASIDE {0} ({1:N0} KB) -> {2}" -f $name, ($size / 1KB), $aside)
    }
}
if ($Execute) { Write-Host "DONE: $($targets.Count) data file(s) set aside; Outlook makes new ones and downloads the mailboxes at its next start." }
else { Write-Host 'DRY RUN: -Execute to rename.' }
exit 0
