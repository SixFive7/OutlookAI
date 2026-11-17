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

    AND THE INDEX AFTERWARDS (2026-11-17). New data files are new stores to Windows Search: each gets a
    new index root, '<store>(<hash>)', and the set-aside files' roots stay beside them. On 2026-11-17,
    8 minutes after Outlook started on new files, the index held three roots for the real mailbox and
    none for the test account, and six index tests failed for it. So the runner records the roots
    before it sets the files aside (-RecordIndexRoots) and, once Outlook runs again, waits until every
    store the staged live-test settings index has a root that was not there before (-WaitForIndexRoots).
    Both only read the index, as the user whose mailboxes it holds - in session 1, through
    Register-InteractiveTask.ps1; they write nothing but -RootsFile.

.PARAMETER Execute
    Rename the files. Without it, a dry run.

.PARAMETER Folder
    Where the data files are. Default %LOCALAPPDATA%\Microsoft\Outlook of the user running it.

.PARAMETER RecordIndexRoots
    Write the index roots of this user's mapi16 scope, one per line, to -RootsFile, and print
    'INDEX-ROOTS-RECORDED <n>'. Before the files are set aside.

.PARAMETER WaitForIndexRoots
    Wait, at most -TimeoutSeconds, until every store the staged live-test settings name in
    indexedStoreDisplayNames has an index root not in -RootsFile. Prints 'INDEX-ROOTS-READY <seconds>' or
    'INDEX-ROOTS-TIMEOUT <seconds>: <how many stores still have none>', and exits 0 either way: a store
    still unindexed is the tests' to report, not a reason to run none.

.PARAMETER RootsFile
    Where -RecordIndexRoots writes and -WaitForIndexRoots reads. Default C:\OutlookAI-Q5\index-roots-before.txt.

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
    [Parameter(Mandatory = $true, ParameterSetName = 'RecordIndex')] [switch] $RecordIndexRoots,
    [Parameter(Mandatory = $true, ParameterSetName = 'WaitIndex')] [switch] $WaitForIndexRoots,
    [Parameter(ParameterSetName = 'RecordIndex')]
    [Parameter(ParameterSetName = 'WaitIndex')] [string] $RootsFile = 'C:\OutlookAI-Q5\index-roots-before.txt',
    [Parameter(ParameterSetName = 'WaitIndex')] [string] $SettingsPath = 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json',
    [Parameter(ParameterSetName = 'WaitIndex')] [ValidateRange(0, 7200)] [int] $TimeoutSeconds = 1200,
    [Parameter(ParameterSetName = 'Run')]
    [Parameter(ParameterSetName = 'RecordIndex')]
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

# The stores that have no index root yet that was not there before: a root is a store's when it is
# '<store>(...)', compared ignoring case. Pure.
function Get-StoresWithoutNewRoot {
    param([string[]] $Stores, [string[]] $Before, [string[]] $Now)
    $old = @{}
    foreach ($r in @($Before | Where-Object { $_ })) { $old[$r.ToLowerInvariant()] = $true }
    $fresh = @($Now | Where-Object { $_ -and -not $old.ContainsKey($_.ToLowerInvariant()) })
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($store in @($Stores | Where-Object { $_ })) {
        $own = @($fresh | Where-Object { $_.StartsWith($store + '(', [StringComparison]::OrdinalIgnoreCase) })
        if ($own.Count -eq 0) { $missing.Add($store) }
    }
    return $missing.ToArray()
}

# Every non-empty string of one list field of the live-test settings' JSON text. Pure.
function Get-SettingsList {
    param([string] $Json, [string] $Field)
    $settings = $Json | ConvertFrom-Json
    if (-not ($settings.PSObject.Properties.Name -contains $Field)) { return @() }
    return @($settings.$Field | Where-Object { $_ -is [string] -and $_ })
}

# The distinct index roots of this user's mapi16 scope, read from Windows Search. Reads only.
function Read-IndexRoots {
    $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $scope = 'mapi16://{' + $sid + '}/'
    $roots = @{}
    $conn = New-Object -ComObject ADODB.Connection
    $conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
    try {
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.Open("SELECT System.ItemUrl FROM SystemIndex WHERE SCOPE='" + $scope + "'", $conn)
        while (-not $rs.EOF) {
            $root = Get-IndexRootName ([string]$rs.Fields.Item('System.ItemUrl').Value) $scope
            if ($root) { $roots[$root] = $true }
            $rs.MoveNext()
        }
        $rs.Close()
    }
    finally { $conn.Close() }
    return @($roots.Keys | Sort-Object)
}

# Read-IndexRoots, again every 10 s for at most $Seconds while Windows Search refuses the query: right
# after a restart it answers 0x80041820 for a while (measured 2026-11-17, 20 s after the boot).
function Read-IndexRootsPatiently {
    param([int] $Seconds)
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ($true) {
        try { return @(Read-IndexRoots) }
        catch {
            if ((Get-Date) -ge $deadline) { throw }
            Start-Sleep -Seconds 10
        }
    }
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
    $scope = 'mapi16://{S-1-5-21-1}/'
    Check 'an item URL gives its root' 'hub@example.invalid($9b7b502)' (Get-IndexRootName 'mapi16://{S-1-5-21-1}/hub@example.invalid($9b7b502)/0/Inbox/at=1' $scope)
    Check 'a URL outside the scope gives none' '' ([string](Get-IndexRootName 'file:///C:/x.txt' $scope))
    $stores = @('hub@example.invalid', 'real@example.invalid')
    $before = @('real@example.invalid($65e0d53e)')
    Check 'a store with no root at all is missing' 'hub@example.invalid|real@example.invalid' ((Get-StoresWithoutNewRoot $stores $before $before) -join '|')
    Check 'an old root does not count' 'hub@example.invalid|real@example.invalid' ((Get-StoresWithoutNewRoot $stores $before @('REAL@example.invalid($65e0d53e)')) -join '|')
    Check 'a new root does, ignoring case' 'real@example.invalid' ((Get-StoresWithoutNewRoot $stores $before @('real@example.invalid($65e0d53e)', 'Hub@Example.invalid($1a2b3c4d)')) -join '|')
    Check 'every store with a new root leaves none' 0 @(Get-StoresWithoutNewRoot $stores $before @('hub@example.invalid($1)', 'real@example.invalid($2)')).Count
    Check 'a root of a longer name is not the store''s' 'hub@example.invalid' ((Get-StoresWithoutNewRoot @('hub@example.invalid') @() @('hub@example.invalid.test($1)')) -join '|')
    Check 'the indexed stores are read from the staged settings' 'hub@example.invalid|real@example.invalid' ((Get-SettingsList '{ "indexedStoreDisplayNames": [ "hub@example.invalid", "", "real@example.invalid" ] }' 'indexedStoreDisplayNames') -join '|')
    Check 'an absent list is empty' 0 @(Get-SettingsList '{ "probeTerm": "" }' 'indexedStoreDisplayNames').Count
    Write-Host ''
    if ($script:fails -gt 0) { Write-Host "$script:fails FAILED"; exit 1 }
    Write-Host 'All passed.'
    exit 0
}

# THE GUARD, FIRST: this renames an Outlook profile's data files, so only on a testbed guest.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

if ($RecordIndexRoots) {
    $roots = @(Read-IndexRootsPatiently 300)
    [System.IO.File]::WriteAllLines($RootsFile, [string[]]$roots)
    Write-Host "INDEX-ROOTS-RECORDED $($roots.Count)"
    exit 0
}
if ($WaitForIndexRoots) {
    $stores = @()
    if (Test-Path -LiteralPath $SettingsPath) { $stores = @(Get-SettingsList (Get-Content -LiteralPath $SettingsPath -Raw) 'indexedStoreDisplayNames') }
    if ($stores.Count -eq 0) { Write-Host "INDEX-ROOTS-READY 0: no staged settings at $SettingsPath name an indexed store"; exit 0 }
    $before = @()
    if (Test-Path -LiteralPath $RootsFile) { $before = @(Get-Content -LiteralPath $RootsFile | Where-Object { $_ }) }
    $t0 = Get-Date
    $lastError = ''
    while ($true) {
        $now = @()
        try { $now = @(Read-IndexRoots); $lastError = '' } catch { $lastError = " (the last query failed: $($_.Exception.Message))" }
        $missing = @(Get-StoresWithoutNewRoot $stores $before $now)
        $waited = [int]((Get-Date) - $t0).TotalSeconds
        if ($missing.Count -eq 0) { Write-Host "INDEX-ROOTS-READY $waited"; exit 0 }
        if ($waited -ge $TimeoutSeconds) { Write-Host "INDEX-ROOTS-TIMEOUT ${waited}: $($missing.Count) of $($stores.Count) indexed store(s) still have no new index root$lastError"; exit 0 }
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
if ($Execute) { Write-Host "DONE: $($targets.Count) data file(s) set aside; Outlook makes new ones and downloads the mailboxes at its next start." }
else { Write-Host 'DRY RUN: -Execute to rename.' }
exit 0
