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

    AND THE INDEX AFTERWARDS (2026-11-17). New data files are new stores to Windows Search, each with a
    new index root, '<store>(<hash>)', and the set-aside files' roots stay in the catalog beside them.
    On 2026-11-17, after the set-asides of that day, the live tier's scope finder - which never guesses
    between two roots of one name (McpServer/OutlookAI.McpServer.Tests/T2/LiveIndexScopes.cs) - found
    three scopes for the real mailbox and none for the test account, and six index tests failed on it.
    So -ResetSearchIndex, with -Execute, has Windows Search rebuild its catalog from nothing after the
    set-aside, the documented way (SetupCompletedSuccessfully = 0, the service restarted), and
    -WaitForIndexRoots waits, once Outlook runs again, until every store the staged live-test settings
    index has exactly one root with something indexed below it. The wait only reads the index, as the
    user whose mailboxes it holds - in session 1, through Register-InteractiveTask.ps1.

.PARAMETER Execute
    Rename the files. Without it, a dry run.

.PARAMETER Folder
    Where the data files are. Default %LOCALAPPDATA%\Microsoft\Outlook of the user running it.

.PARAMETER ResetSearchIndex
    With -Execute: after the set-aside, have Windows Search rebuild its catalog from nothing - stop the
    WSearch service, set HKLM\SOFTWARE\Microsoft\Windows Search\SetupCompletedSuccessfully to 0, start it -
    so that no root of a set-aside file is left in it. Needs the elevated session PowerShell Direct gives.

.PARAMETER WaitForIndexRoots
    Wait, at most -TimeoutSeconds, until every store the staged live-test settings name in
    indexedStoreDisplayNames has exactly one index root, with something indexed below it. Prints
    'INDEX-ROOTS-READY <seconds>' or 'INDEX-ROOTS-TIMEOUT <seconds>: <how many stores are not ready>',
    and exits 0 either way: a store not ready is the tests' to report, not a reason to run none.

.PARAMETER SettingsPath
    The staged live-test settings (gitignored), whose indexedStoreDisplayNames -WaitForIndexRoots waits for.

.PARAMETER TimeoutSeconds
    How long -WaitForIndexRoots waits. Default 1200.

.PARAMETER SelfTest
    The pure part: which files a folder listing makes it rename, and which stores a list of index roots
    leaves without a new one. No guest, no Outlook.
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')] [switch] $Execute,
    [Parameter(ParameterSetName = 'Run')] [string] $Folder,
    [Parameter(ParameterSetName = 'Run')] [switch] $ResetSearchIndex,
    [Parameter(Mandatory = $true, ParameterSetName = 'WaitIndex')] [switch] $WaitForIndexRoots,
    [Parameter(ParameterSetName = 'WaitIndex')] [string] $SettingsPath,
    [Parameter(ParameterSetName = 'WaitIndex')] [ValidateRange(0, 7200)] [int] $TimeoutSeconds = 1200,
    [Parameter(ParameterSetName = 'Run')]
    [Parameter(ParameterSetName = 'WaitIndex')] [string[]] $ExpectedUser = @('vmadmin'),
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

# The index root an item's URL sits under - the first path segment after the user's mapi16 scope,
# '<store>(<hash>)' - or $null for a URL outside the scope. Pure.
function Get-IndexRootName {
    param([string] $Url, [string] $Scope)
    if (-not $Url -or -not $Scope -or -not $Url.StartsWith($Scope, [StringComparison]::OrdinalIgnoreCase)) { return $null }
    $root = ($Url.Substring($Scope.Length) -split '/')[0]
    if (-not $root) { return $null }
    return $root
}

# The stores that are not ready for an index test: not exactly one root of their name, '<store>(...)'
# compared ignoring case, or nothing indexed below the one. $RowsByRoot maps each root to its rows. Pure.
function Get-StoresNotReady {
    param([string[]] $Stores, [hashtable] $RowsByRoot)
    $notReady = New-Object System.Collections.Generic.List[string]
    foreach ($store in @($Stores | Where-Object { $_ })) {
        $own = @($RowsByRoot.Keys | Where-Object { ([string]$_).StartsWith($store + '(', [StringComparison]::OrdinalIgnoreCase) })
        if ($own.Count -ne 1 -or [int]$RowsByRoot[$own[0]] -lt 1) { $notReady.Add($store) }
    }
    return $notReady.ToArray()
}

# Each index root of this user's mapi16 scope, with the rows below it, read from Windows Search. Reads only.
function Read-IndexRoots {
    $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $scope = 'mapi16://{' + $sid + '}/'
    $rows = @{}
    $conn = New-Object -ComObject ADODB.Connection
    $conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
    try {
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.Open("SELECT System.ItemUrl FROM SystemIndex WHERE SCOPE='" + $scope + "'", $conn)
        while (-not $rs.EOF) {
            $root = Get-IndexRootName ([string]$rs.Fields.Item('System.ItemUrl').Value) $scope
            if ($root) { $rows[$root] = 1 + [int]$rows[$root] }
            $rs.MoveNext()
        }
        $rs.Close()
    }
    finally { $conn.Close() }
    return $rows
}

# The staged live-test settings and how a field of them is read (Q189 (b)): one copy for every guest
# script that needs them (AGENTS.md, Shared code).
. "$PSScriptRoot\StagedLiveTestSettings.ps1"
if (-not $SettingsPath) { $SettingsPath = $StagedSettingsPath }

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
    $scope = 'mapi16://{S-1-5-21-1}/'
    Check 'an item URL gives its root' 'hub@example.invalid($9b7b502)' (Get-IndexRootName 'mapi16://{S-1-5-21-1}/hub@example.invalid($9b7b502)/0/Inbox/at=1' $scope)
    Check 'a URL outside the scope gives none' '' ([string](Get-IndexRootName 'file:///C:/x.txt' $scope))
    $stores = @('hub@example.invalid', 'real@example.invalid')
    Check 'a store with no root is not ready' 'hub@example.invalid' ((Get-StoresNotReady $stores @{ 'real@example.invalid($2)' = 40 }) -join '|')
    Check 'nor one with two roots, whichever is fuller' 'real@example.invalid' ((Get-StoresNotReady $stores @{ 'hub@example.invalid($1)' = 19; 'real@example.invalid($2)' = 40; 'REAL@example.invalid($3)' = 2 }) -join '|')
    Check 'nor one whose one root holds nothing' 'hub@example.invalid' ((Get-StoresNotReady $stores @{ 'hub@example.invalid($1)' = 0; 'real@example.invalid($2)' = 40 }) -join '|')
    Check 'one root each, with rows, is ready' 0 @(Get-StoresNotReady $stores @{ 'Hub@Example.invalid($1)' = 19; 'real@example.invalid($2)' = 40 }).Count
    Check 'a root of a longer name is not the store''s' 'hub@example.invalid' ((Get-StoresNotReady @('hub@example.invalid') @{ 'hub@example.invalid.test($1)' = 5 }) -join '|')
    Check 'the indexed stores are read from the staged settings' 'hub@example.invalid|real@example.invalid' ((Get-SettingsStrings '{ "indexedStoreDisplayNames": [ "hub@example.invalid", "", "real@example.invalid" ] }' 'indexedStoreDisplayNames') -join '|')
    Check 'an absent list is empty' 0 @(Get-SettingsStrings '{ "probeTerm": "" }' 'indexedStoreDisplayNames').Count
    # StagedLiveTestSettings.ps1, which the Exchange VM's probes share: a field's first string, from text or a file.
    $staged = '{ "testHubStoreDisplayName": "hub@example.invalid", "bystanderStoreDisplayNames": [ "", "real@example.invalid" ], "probeTerm": "" }'
    Check 'the hub is the one the staged settings name' 'hub@example.invalid' (Get-SettingsValue $staged 'testHubStoreDisplayName')
    Check 'a list gives its first non-empty name' 'real@example.invalid' (Get-SettingsValue $staged 'bystanderStoreDisplayNames')
    Check 'an empty field gives nothing' $true ($null -eq (Get-SettingsValue $staged 'probeTerm'))
    Check 'an absent field gives nothing' $true ($null -eq (Get-SettingsValue $staged 'mailSink'))
    Check 'no staged settings give nothing' $true ($null -eq (Get-StagedSettingsValue (Join-Path ([System.IO.Path]::GetTempPath()) ('no-such-settings-' + [guid]::NewGuid().ToString('N') + '.json')) 'testHubStoreDisplayName'))
    Check 'the staged file is where the live runner puts it' 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json' $StagedSettingsPath
    Write-Host ''
    if ($script:fails -gt 0) { Write-Host "$script:fails FAILED"; exit 1 }
    Write-Host 'All passed.'
    exit 0
}

# THE GUARD, FIRST: this renames an Outlook profile's data files, so only on a testbed guest.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

if ($WaitForIndexRoots) {
    $stores = @()
    if (Test-Path -LiteralPath $SettingsPath) { $stores = @(Get-SettingsStrings (Get-Content -LiteralPath $SettingsPath -Raw) 'indexedStoreDisplayNames') }
    if ($stores.Count -eq 0) { Write-Host "INDEX-ROOTS-READY 0: no staged settings at $SettingsPath name an indexed store"; exit 0 }
    $t0 = Get-Date
    $lastError = ''
    while ($true) {
        # Right after a restart Windows Search refuses the query for a while (0x80041820, measured
        # 2026-11-17 20 s after a boot): a refused query reads as nothing indexed yet.
        $rows = @{}
        try { $rows = Read-IndexRoots; $lastError = '' } catch { $lastError = " (the last query failed: $($_.Exception.Message))" }
        $notReady = @(Get-StoresNotReady $stores $rows)
        $waited = [int]((Get-Date) - $t0).TotalSeconds
        if ($notReady.Count -eq 0) { Write-Host "INDEX-ROOTS-READY $waited"; exit 0 }
        if ($waited -ge $TimeoutSeconds) { Write-Host "INDEX-ROOTS-TIMEOUT ${waited}: $($notReady.Count) of $($stores.Count) indexed store(s) have not exactly one root with something below it$lastError"; exit 0 }
        Start-Sleep -Seconds 20
    }
}

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
if ($Execute -and $ResetSearchIndex) {
    # The set-aside files' roots would stay in the catalog beside the new ones (see the header).
    Stop-Service -Name WSearch -Force
    Set-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows Search' -Name 'SetupCompletedSuccessfully' -Value 0 -Type DWord
    Start-Service -Name WSearch
    Write-Host 'SEARCH-INDEX-RESET: Windows Search rebuilds its catalog from nothing'
}
if ($Execute) { Write-Host "DONE: $($targets.Count) data file(s) set aside; Outlook makes new ones and downloads the mailboxes at its next start." }
else { Write-Host 'DRY RUN: -Execute to rename.' }
exit 0
