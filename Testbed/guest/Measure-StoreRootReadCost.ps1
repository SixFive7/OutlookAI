#Requires -Version 5.1
<#
.SYNOPSIS
    READ-ONLY measurement: what reading each store's top folder costs, and whether two stores of the
    profile carry the same StoreID or the same provider UID. Written for the store-ID follow-up of
    Q131-Q136 (option (b): a 28-character S store ID for every store, an uncached Exchange store too).

.DESCRIPTION
    RUN ON A GUEST, in session 1 through Testbed/guest/Register-InteractiveTask.ps1 -RunLevel Limited,
    with Outlook ALREADY RUNNING: it attaches to that Outlook and never starts one. On the Exchange VM
    a start past Office's activation grace raises the activation wizard, and a start is not what is
    being measured. Windows PowerShell 5.1.

    WHAT IT MEASURES. A store's ID is S and the provider UID at bytes 4..19 of its top folder's entry
    id - except for an Exchange store Outlook does not cache, whose top folder is a server round trip:
    OutlookComSession.GetStoreDetails skips that read, and the store gets T and its whole StoreID
    instead, an ID that is long and changes when the store's caching mode changes. Option (b) reads the
    top folder once per store-list refresh. So for every store this times, -Repeat + 1 times, what
    GetStoreDetails would pay: Stores.Item(i), the reads it always makes (DisplayName, StoreID,
    ExchangeStoreType, IsCachedExchange), Store.GetRootFolder() and the top folder's EntryID - each pass
    on a store object fetched afresh, as every refresh does. The first pass is reported apart from the
    rest. Outlook may already hold a store's top folder open (a windowed Outlook's folder pane opens
    every store's), so a first pass can be warm too: NamespaceState says which Outlook it was.

    And it compares every store with every other by its StoreID and by the provider UID its top folder
    carries, because Exchange automaps a mailbox the profile's account has Full Access to as a second,
    online entry beside the mailbox's own cached store (Docs/live-tier-on-the-vm.md section 4.4) - one
    mailbox in two entries, which no ID made from the mailbox's identity can tell apart.

    PRINTED: per store, its position in Namespace.Stores, Exchange store type, cached flag, the
    timings, and which other store (by position) shares its StoreID or its provider UID. Never a name,
    an address, a StoreID or a UID: those are compared through a short SHA-256 prefix, which says
    "equal" or "not equal" and nothing else.

    WHAT MAKES IT SAFE: GetRootFolder, EntryID and property reads, nothing else - no Save, Delete,
    Move, Add, Send or Display. Keep it that way if you extend it. Guarded like every guest script
    (Assert-TestbedGuest, so Testbed/guest/OutlookMapiInterop.ps1 must sit beside it) before it writes
    its result file or binds Outlook.

.PARAMETER Repeat
    Passes after the first, per store. Default 5.

.PARAMETER OutPath
    Where the result is written as well as printed. Default C:\OutlookAI-StreamK\measure\store-root-read-cost.txt.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.PARAMETER SelfTest
    The pure parts - the provider UID's place in an entry id, the hash prefix, the pairing of equal
    values and the summary. No Outlook, any machine.
#>
[CmdletBinding()]
param(
    [ValidateRange(0, 50)] [int] $Repeat = 5,
    [string] $OutPath = 'C:\OutlookAI-StreamK\measure\store-root-read-cost.txt',
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# =============================================================================================
# PURE PARTS
# =============================================================================================

# Bytes 4..19 of an entry id, as upper-case hex - the provider UID every entry id of a store carries
# (OutlookIds.ProviderUidHexOf). $null for anything shorter than 20 bytes, not hex, or an all-zero UID.
function Get-ProviderUidHex([string] $EntryIdHex) {
    if ([string]::IsNullOrEmpty($EntryIdHex) -or $EntryIdHex.Length -lt 40 -or ($EntryIdHex.Length % 2) -ne 0) { return $null }
    if ($EntryIdHex -notmatch '^[0-9A-Fa-f]+$') { return $null }
    $uid = $EntryIdHex.Substring(8, 32).ToUpperInvariant()
    if ($uid -eq ('0' * 32)) { return $null }
    return $uid
}

# The first 12 hex digits of the SHA-256 of a value's upper-case spelling: enough to say equal or not,
# not enough to say what the value is. $null for no value.
function Get-ShortHash([string] $Value) {
    if ([string]::IsNullOrEmpty($Value)) { return $null }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value.ToUpperInvariant()))
    }
    finally {
        $sha.Dispose()
    }
    return (($bytes[0..5] | ForEach-Object { $_.ToString('x2') }) -join '')
}

# For each position, the other positions whose value is the same (case-insensitive); '-' for none and
# for a position with no value. Values: an array indexed like the stores, $null where nothing read.
function Get-SharedWith([object[]] $Values) {
    $result = @()
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $mine = [string]$Values[$i]
        if ([string]::IsNullOrEmpty($mine)) { $result += '-'; continue }
        $others = @()
        for ($j = 0; $j -lt $Values.Count; $j++) {
            if ($j -ne $i -and -not [string]::IsNullOrEmpty([string]$Values[$j]) -and [string]::Equals($mine, [string]$Values[$j], [StringComparison]::OrdinalIgnoreCase)) {
                $others += ($j + 1)
            }
        }
        if ($others.Count -eq 0) { $result += '-' } else { $result += ($others -join ',') }
    }
    return , $result
}

# min/median/max of a list of milliseconds, one decimal, as 'min/median/max'; 'n/a' for none. A $null
# in the list - a pass that failed before it was timed - is left out, never counted as 0.
function Format-Spread([object[]] $Ms) {
    $list = @($Ms | Where-Object { $null -ne $_ } | ForEach-Object { [double]$_ })
    if ($list.Count -eq 0) { return 'n/a' }
    $sorted = @($list | Sort-Object)
    $mid = $sorted[[int][Math]::Floor(($sorted.Count - 1) / 2)]
    $f = { param($v) ([double]$v).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture) }
    return ((& $f $sorted[0]) + '/' + (& $f $mid) + '/' + (& $f $sorted[$sorted.Count - 1]))
}

# The one-line verdict on the identities: which pairs of stores share a provider UID - and of those,
# which share their StoreID too.
function Get-IdentityVerdict([object[]] $UidShared, [object[]] $StoreIdShared) {
    $pairs = @()
    for ($i = 0; $i -lt $UidShared.Count; $i++) {
        if ([string]$UidShared[$i] -eq '-') { continue }
        foreach ($other in ([string]$UidShared[$i] -split ',')) {
            $j = [int]$other
            if ($j -le ($i + 1)) { continue }
            $sameStoreId = ([string]$StoreIdShared[$i] -split ',') -contains [string]$j
            $pairs += ('' + ($i + 1) + '~' + $j + $(if ($sameStoreId) { '(same StoreID)' } else { '(StoreIDs differ)' }))
        }
    }
    if ($pairs.Count -eq 0) { return 'NO-SHARED-UID' }
    return 'SHARED-UID ' + ($pairs -join ' ')
}

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $Name, [bool] $Ok) { if ($Ok) { $script:p++; Write-Host "  PASS  $Name" } else { $script:f++; Write-Host "  FAIL  $Name" } }

    $exchange = '00000000' + ('A1' * 16) + '0100' + ('22' * 16) + '000000001234' + '0000'
    Check 'the provider UID is bytes 4..19' ((Get-ProviderUidHex $exchange) -eq ('A1' * 16))
    Check 'lower-case hex gives the same UID' ((Get-ProviderUidHex $exchange.ToLowerInvariant()) -eq ('A1' * 16))
    Check 'an id shorter than 20 bytes has none' ($null -eq (Get-ProviderUidHex ('00' * 19)))
    Check 'an all-zero UID is none' ($null -eq (Get-ProviderUidHex ('00' * 24)))
    Check 'a value that is not hex has none' ($null -eq (Get-ProviderUidHex ('ZZ' * 24)))
    Check 'a short hash is 12 hex digits' ((Get-ShortHash 'abc') -match '^[0-9a-f]{12}$')
    Check 'the hash ignores case' ((Get-ShortHash 'abc') -eq (Get-ShortHash 'ABC'))
    Check 'different values hash apart' ((Get-ShortHash 'abc') -ne (Get-ShortHash 'abd'))
    Check 'no value has no hash' ($null -eq (Get-ShortHash ''))
    $shared = Get-SharedWith @('X', 'y', 'x', $null, 'Z')
    Check 'equal values name each other' (($shared[0] -eq '3') -and ($shared[2] -eq '1'))
    Check 'a value of its own shares with none' (($shared[1] -eq '-') -and ($shared[4] -eq '-'))
    Check 'no value shares with none' ($shared[3] -eq '-')
    Check 'three equal values name both others' ((Get-SharedWith @('a', 'a', 'a'))[1] -eq '1,3')
    Check 'the spread is min/median/max' ((Format-Spread @(5, 1, 3)) -eq '1.0/3.0/5.0')
    Check 'an empty spread says so' ((Format-Spread @()) -eq 'n/a')
    Check 'a pass that was never timed is left out, not counted as 0' ((Format-Spread @($null, 4, 2)) -eq '2.0/2.0/4.0')
    Check 'no shared UID is said so' ((Get-IdentityVerdict @('-', '-') @('-', '-')) -eq 'NO-SHARED-UID')
    Check 'a shared UID with one StoreID is named once' ((Get-IdentityVerdict @('2', '1') @('2', '1')) -eq 'SHARED-UID 1~2(same StoreID)')
    Check 'a shared UID with two StoreIDs says they differ' ((Get-IdentityVerdict @('-', '3', '2') @('-', '-', '-')) -eq 'SHARED-UID 2~3(StoreIDs differ)')

    Write-Host ''
    Write-Host ("Measure-StoreRootReadCost.ps1 -SelfTest: {0} passed, {1} failed." -f $script:p, $script:f)
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

# =============================================================================================
# THE MEASUREMENT - on a guest only
# =============================================================================================

. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

$lines = New-Object System.Collections.Generic.List[string]
function Out-Line([string] $Text) { $lines.Add($Text); Write-Output $Text }
function Get-HResult($ErrorRecord) {
    $e = $ErrorRecord.Exception
    while ($e.InnerException -and -not ($e -is [Runtime.InteropServices.COMException])) { $e = $e.InnerException }
    return ('0x{0:X8}' -f $e.HResult) + ' ' + $e.GetType().Name
}
function Release($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

if (-not (Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)) {
    Out-Line 'OUTLOOK-NOT-RUNNING: this probe attaches to a running Outlook and never starts one. Nothing measured.'
    exit 3
}

try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
}
catch {
    Out-Line ('ATTACH-FAILED ' + (Get-HResult $_) + ': Outlook runs, but not at this integrity level or not in this session. Nothing measured.')
    exit 3
}

$ns = $outlook.GetNamespace('MAPI')
$explorers = -1
try { $explorers = [int]$outlook.Explorers.Count } catch { }
$mode = 'unreadable'
try { $mode = [string][int]$ns.ExchangeConnectionMode } catch { }
$offline = 'unreadable'
try { $offline = [string][bool]$ns.Offline } catch { }
Out-Line ('RUN ' + (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') + ' computer=' + $env:COMPUTERNAME + ' outlook=' + [string]$outlook.Version + ' repeat=' + $Repeat)
Out-Line ('NamespaceState explorers=' + $explorers + ' exchangeConnectionMode=' + $mode + ' offline=' + $offline)

$stores = $ns.Stores
$count = [int]$stores.Count
$storeIds = New-Object object[] $count
$uids = New-Object object[] $count
$rows = New-Object object[] $count

for ($i = 1; $i -le $count; $i++) {
    $first = $null
    $later = New-Object System.Collections.Generic.List[double]
    $propsMs = New-Object System.Collections.Generic.List[double]
    $type = $null; $cached = $null; $uid = $null; $storeIdHex = $null; $errorText = $null
    for ($pass = 0; $pass -le $Repeat; $pass++) {
        $store = $null; $root = $null
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $store = $stores.Item($i)
            [void][string]$store.DisplayName
            $storeIdHex = [string]$store.StoreID
            try { $type = [int]$store.ExchangeStoreType } catch { }
            try { $cached = [bool]$store.IsCachedExchange } catch { }
            $afterProps = $sw.Elapsed.TotalMilliseconds
            $root = $store.GetRootFolder()
            $entryId = [string]$root.EntryID
            $total = $sw.Elapsed.TotalMilliseconds
            $propsMs.Add($afterProps)
            if ($pass -eq 0) { $first = $total - $afterProps } else { $later.Add($total - $afterProps) }
            $uid = Get-ProviderUidHex $entryId
        }
        catch {
            $failedAt = $sw.Elapsed.TotalMilliseconds
            $errorText = (Get-HResult $_) + ' after ' + $failedAt.ToString('0.0', [Globalization.CultureInfo]::InvariantCulture) + ' ms'
            if ($pass -eq 0) { $first = $failedAt }
        }
        finally {
            Release $root
            Release $store
        }
    }

    $storeIds[$i - 1] = $storeIdHex
    $uids[$i - 1] = $uid
    $rows[$i - 1] = [pscustomobject]@{
        Index   = $i
        Type    = $type
        Cached  = $cached
        First   = $first
        Later   = $later.ToArray()
        Props   = $propsMs.ToArray()
        Error   = $errorText
    }
}

$storeIdShared = Get-SharedWith $storeIds
$uidShared = Get-SharedWith $uids
foreach ($row in $rows) {
    $k = $row.Index - 1
    $firstText = 'n/a'
    if ($null -ne $row.First) { $firstText = ([double]$row.First).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture) }
    $errorPart = ''
    if ($row.Error) { $errorPart = ' error=' + $row.Error }
    $text = 'STORE ' + $row.Index + ' exchangeStoreType=' + $row.Type + ' cached=' + $row.Cached +
        ' storeIdHash=' + (Get-ShortHash ([string]$storeIds[$k])) + ' sameStoreIdAs=' + $storeIdShared[$k] +
        ' rootUidHash=' + (Get-ShortHash ([string]$uids[$k])) + ' sameRootUidAs=' + $uidShared[$k] +
        ' rootFirstMs=' + $firstText + ' rootLaterMs=' + (Format-Spread $row.Later) +
        ' alwaysReadPropsMs=' + (Format-Spread $row.Props) + $errorPart
    Out-Line $text
}

$uncached = @($rows | Where-Object { $null -ne $_.Type -and $_.Type -ne 3 -and $_.Cached -ne $true })
$answer = 'ANSWER uncachedExchangeStores=' + $uncached.Count +
    ' rootFirstMs=' + (Format-Spread @($uncached | ForEach-Object { $_.First })) +
    ' rootLaterMs=' + (Format-Spread @($uncached | ForEach-Object { $_.Later })) +
    ' identity=' + (Get-IdentityVerdict $uidShared $storeIdShared)
Out-Line $answer

$dir = Split-Path -Parent $OutPath
if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
Set-Content -LiteralPath $OutPath -Value $lines -Encoding UTF8
Release $stores
Release $ns
Release $outlook
exit 0
