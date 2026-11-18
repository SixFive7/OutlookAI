#Requires -Version 5.1
<#
.SYNOPSIS
    READ-ONLY measurement for the folder-ID follow-ups of Q131-Q136: what reading each store's top folder
    costs, whether two stores of the profile are one mailbox, which store entry a folder ID opens in -
    and an ID no folder has - and whether the subject-and-time probe that opens an index hit finds the
    test account's newest indexed items.

.DESCRIPTION
    RUN ON A GUEST, in session 1 through Testbed/guest/Register-InteractiveTask.ps1 -RunLevel Limited,
    with Outlook ALREADY RUNNING: it attaches to that Outlook and never starts one. On the Exchange VM
    a start past Office's activation grace raises the activation wizard, and a start is not what is
    being measured. Windows PowerShell 5.1.

    -Section Roots (written for option (b) of the uncached Exchange store's ID). A store's ID is S and the
    provider UID at bytes 4..19 of its top folder's entry id - except an Exchange store Outlook does not
    cache, whose top folder OutlookComSession.GetStoreDetails did not read (a server round trip), so it got
    T and its whole StoreID. For every store this times, -Repeat + 1 times, what GetStoreDetails would pay:
    Stores.Item(i), the reads it always makes, Store.GetRootFolder() and the top folder's EntryID - each
    pass on a store object fetched afresh. The first pass is reported apart; a windowed Outlook's folder
    pane holds every store's top folder open, so NamespaceState says which Outlook it was. And it compares
    every store with every other by StoreID and by top-folder provider UID: Exchange automaps a mailbox the
    profile's account has Full Access to as a second, online entry beside the mailbox's own cached store
    (Docs/live-tier-on-the-vm.md section 4.4a) - one mailbox in two entries.

    -Section FolderIds (the twin entries and the ID no folder has). For every store whose top folder has a
    46-byte Exchange entry id: that id, and the same id with its six counter bytes set to FF - the ID
    T2/LiveIds.UnknownFolderIdIn builds, which no folder has - are opened with GetFolderFromID, once with
    the id alone and once with each store's StoreID beside it. Each line says whether it opened, in WHICH
    store entry (its position, type and cached flag), and whether the opened folder's Name, DefaultItemType
    and item count read - a folder that opens and then answers nothing is not one.

    -Section IndexHits (the thread test that could not open its one index member). The test account's
    newest index rows, as LiveMailServiceTests.Thread_IndexPath_AndComFallback picks them - its store,
    newest received first, no attachment rows - and for each: its kind, its folder path, and what the
    subject-and-time probe the index-hit locator runs finds there (OutlookComSession.TryResolveByPath):
    how many items carry the subject, and how far their ReceivedTime, CreationTime and SentOn are from the
    row's received time. The hub comes from the staged live-test settings (StagedLiveTestSettings.ps1).

    PRINTED: positions, Exchange store types, cached flags, folder paths of the test account's mailbox,
    item classes, counts, seconds and HRESULTs. Never a store's name, an address, a StoreID, a UID, an
    entry id or a subject: those are compared through a short SHA-256 prefix, which says equal or not and
    nothing else.

    WHAT MAKES IT SAFE: GetFolderFromID, GetRootFolder, Restrict, EntryID and property reads, and a SELECT
    against the search index - no Save, Delete, Move, Add, Send or Display. Keep it that way if you extend
    it; -SelfTest checks the script for a member that changes anything. Guarded like every guest script
    (Assert-TestbedGuest, so Testbed/guest/OutlookMapiInterop.ps1 must sit beside it) before it writes its
    result file or binds Outlook.

.PARAMETER Section
    Roots, FolderIds, IndexHits, or All (the default).

.PARAMETER Repeat
    Passes after the first, per store, for -Section Roots. Default 5.

.PARAMETER Hub
    The test account's store, for -Section IndexHits. Default: the hub the staged live-test settings name.

.PARAMETER OutPath
    Where the result is written as well as printed. Default C:\OutlookAI-StreamK\measure\folder-id-resolution.txt.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.PARAMETER SelfTest
    The pure parts - the provider UID's place in an entry id, the hash prefix, the pairing of equal
    values, the unknown ID, the summary - and that nothing in the script changes anything. No Outlook.
#>
[CmdletBinding()]
param(
    [ValidateSet('All', 'Roots', 'FolderIds', 'IndexHits')] [string] $Section = 'All',
    [ValidateRange(0, 50)] [int] $Repeat = 5,
    [string] $Hub,
    [string] $OutPath = 'C:\OutlookAI-StreamK\measure\folder-id-resolution.txt',
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

# The ID T2/LiveIds.UnknownFolderIdIn builds from a 46-byte Exchange folder entry id: its six counter
# bytes (38..43) set to FF - an id no folder has. $null for any other length.
function Get-UnknownExchangeFolderId([string] $EntryIdHex) {
    if ([string]::IsNullOrEmpty($EntryIdHex) -or $EntryIdHex.Length -ne 92 -or $EntryIdHex -notmatch '^[0-9A-Fa-f]+$') { return $null }
    return $EntryIdHex.Substring(0, 76).ToUpperInvariant() + 'FFFFFFFFFFFF' + $EntryIdHex.Substring(88).ToUpperInvariant()
}

# The folder names an index item URL files a row under - the segments between the store-type segment
# (/0/ or /1/) and the item's own last segment, percent-decoded for the five characters the index
# encodes (%25 %2F %5C %2A %3F). @() when the URL has no store-type segment. Pure.
function Get-UrlFolderSegments([string] $Url, [string] $RootPrefix) {
    if ([string]::IsNullOrEmpty($Url) -or [string]::IsNullOrEmpty($RootPrefix) -or -not $Url.StartsWith($RootPrefix, [StringComparison]::OrdinalIgnoreCase)) { return @() }
    $rest = $Url.Substring($RootPrefix.Length).TrimStart('/')
    $parts = @($rest -split '/')
    if ($parts.Count -lt 2 -or ($parts[0] -ne '0' -and $parts[0] -ne '1')) { return @() }
    $folders = @()
    for ($i = 1; $i -lt $parts.Count - 1; $i++) {
        $s = [string]$parts[$i]
        $s = $s.Replace('%2F', '/').Replace('%2f', '/').Replace('%5C', '\').Replace('%5c', '\').Replace('%2A', '*').Replace('%2a', '*').Replace('%3F', '?').Replace('%3f', '?').Replace('%25', '%')
        $folders += $s
    }
    return , $folders
}

# Seconds from the index's received time (UTC) to an Outlook time (local wall clock), one decimal; 'n/a'
# when either is missing. Pure but for the machine's own time zone, which is the conversion's point.
function Format-SecondsFrom($IndexUtc, $OutlookLocal) {
    if ($null -eq $IndexUtc -or $null -eq $OutlookLocal) { return 'n/a' }
    $local = [DateTime]$OutlookLocal
    if ($local.Year -ge 4500) { return 'none' }
    $utc = [DateTime]::SpecifyKind($local, [DateTimeKind]::Local).ToUniversalTime()
    return (($utc - [DateTime]$IndexUtc).TotalSeconds).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)
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
    Check 'a short hash is 12 hex digits (Get-ShortHash, shared from ScriptHelpers.ps1)' ((Get-ShortHash 'abc' -IgnoreCase) -match '^[0-9a-f]{12}$')
    Check 'the hash ignores case' ((Get-ShortHash 'abc' -IgnoreCase) -eq (Get-ShortHash 'ABC' -IgnoreCase))
    Check 'different values hash apart' ((Get-ShortHash 'abc' -IgnoreCase) -ne (Get-ShortHash 'abd' -IgnoreCase))
    Check 'it is the first 12 hex digits of the SHA-256 of the upper-case spelling, as this script''s own copy made it' ((Get-ShortHash 'abc' -IgnoreCase) -eq 'b5d4045c3f46')
    Check 'no value has no hash' ($null -eq (Get-ShortHash '' -IgnoreCase))
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
    $unknown = Get-UnknownExchangeFolderId $exchange
    Check 'the unknown ID keeps the flags, UID, type and database GUID' ($unknown.Substring(0, 76) -eq $exchange.Substring(0, 76).ToUpperInvariant())
    Check 'and sets the six counter bytes to FF' ($unknown.Substring(76, 12) -eq 'FFFFFFFFFFFF')
    Check 'and keeps the pad' ($unknown.Substring(88) -eq '0000')
    Check 'a PST-sized id has no unknown twin here' ($null -eq (Get-UnknownExchangeFolderId ('00' * 24)))
    $root = 'mapi16://{S-1-5-21-1}/someone($abc)'
    $segs = Get-UrlFolderSegments ($root + '/0/Inbox/Sub%2Fname/at=x') $root
    Check 'a URL gives its folder names, the item segment left out' (($segs.Count -eq 2) -and ($segs[0] -eq 'Inbox') -and ($segs[1] -eq 'Sub/name'))
    Check 'a URL of another root gives none' ((Get-UrlFolderSegments ('mapi16://{S-1-5-21-1}/other($1)/0/Inbox/x') $root).Count -eq 0)
    Check 'an encoded percent is decoded last, once' ((Get-UrlFolderSegments ($root + '/0/50%252A off/x') $root)[0] -eq '50%2A off')
    Check 'a missing time is n/a' ((Format-SecondsFrom $null (Get-Date)) -eq 'n/a')
    Check 'Outlook''s no-date (4501) is none' ((Format-SecondsFrom ([DateTime]::UtcNow) ([DateTime]'4501-01-01')) -eq 'none')

    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    # The index's ADODB recordset and connection are closed, and nothing else is: $rs and $conn are the only Close targets allowed.
    $members = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true) | Where-Object { -not ($_.Member.Extent.Text -eq 'Close' -and @('$rs', '$conn') -contains $_.Expression.Extent.Text) } | ForEach-Object { $_.Member.Extent.Text.Trim("'", '"') })
    $bad = @($members | Where-Object { $_ -match '^(Delete|Save|Move|Send|SetProperty|SetProperties|DeleteProperties|Remove|Copy|Close|Quit|Display|SaveAs|ExecuteMso|MoveTo|CopyTo)$' })
    Check ('no late-bound member that changes anything: ' + ($bad -join ', ')) ($bad.Count -eq 0)
    $commands = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -match '^(Set-(?!Content$)|New-Item|Remove-|Rename-|Move-|Add-Content|Out-File)' })
    # Its own result file is the one thing written: Set-Content once, and New-Item for its directory.
    $writers = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -in @('Set-Content', 'New-Item') })
    Check ('no cmdlet that writes, but the result file: ' + ($commands -join ', ')) ($commands.Count -eq 0 -or (@($commands | Where-Object { $_ -ne 'New-Item' }).Count -eq 0))
    Check 'the result file is written in one place' ($writers.Count -eq 2)

    Write-Host ''
    Write-Host ("Measure-FolderIdResolution.ps1 -SelfTest: {0} passed, {1} failed." -f $script:p, $script:f)
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

# Get-ShortHash, by which it compares and prints ids and names, is one of the helpers the scripts share,
# dot-sourced from ScriptHelpers.ps1 beside this script before the self-test that holds it (Q154, S7).
. (Join-Path $PSScriptRoot 'ScriptHelpers.ps1')

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
    return ('0x{0:X8}' -f $e.HResult)
}
function Release($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}
function Ms([double] $Value) { return $Value.ToString('0.0', [Globalization.CultureInfo]::InvariantCulture) }

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
Out-Line ('RUN ' + (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') + ' computer=' + $env:COMPUTERNAME + ' outlook=' + [string]$outlook.Version + ' section=' + $Section + ' repeat=' + $Repeat)
Out-Line ('NamespaceState explorers=' + $explorers + ' exchangeConnectionMode=' + $mode + ' offline=' + $offline)

# Every store once: what tells two entries apart, and the top folder's id.
$stores = $ns.Stores
$count = [int]$stores.Count
$entries = @()
for ($i = 1; $i -le $count; $i++) {
    $store = $null; $root = $null
    try {
        $store = $stores.Item($i)
        $type = $null; $cached = $null
        try { $type = [int]$store.ExchangeStoreType } catch { }
        try { $cached = [bool]$store.IsCachedExchange } catch { }
        $rootId = $null
        try { $root = $store.GetRootFolder(); $rootId = [string]$root.EntryID } catch { }
        $entries += [pscustomobject]@{ Index = $i; Type = $type; Cached = $cached; Name = [string]$store.DisplayName; StoreId = [string]$store.StoreID; RootId = $rootId }
    }
    finally {
        Release $root
        Release $store
    }
}

# Which entry a store object is: by StoreID, type, cached flag AND name - two entries of one mailbox share the StoreID.
function Get-EntryPosition($StoreObject) {
    $sid = $null; $t = $null; $c = $null; $n = $null
    try { $sid = [string]$StoreObject.StoreID } catch { }
    try { $t = [int]$StoreObject.ExchangeStoreType } catch { }
    try { $c = [bool]$StoreObject.IsCachedExchange } catch { }
    try { $n = [string]$StoreObject.DisplayName } catch { }
    $hit = @($entries | Where-Object { $_.StoreId -eq $sid -and $_.Type -eq $t -and $_.Cached -eq $c -and $_.Name -eq $n })
    if ($hit.Count -eq 1) { return [string]$hit[0].Index }
    return '?(' + $hit.Count + ' match)'
}

if ($Section -eq 'All' -or $Section -eq 'Roots') {
    Out-Line '== Roots: the top-folder read GetStoreDetails would pay, and the stores'' identities =='
    $rows = @()
    for ($i = 1; $i -le $count; $i++) {
        $first = $null
        $later = New-Object System.Collections.Generic.List[double]
        $propsMs = New-Object System.Collections.Generic.List[double]
        $errorText = $null
        for ($pass = 0; $pass -le $Repeat; $pass++) {
            $store = $null; $root = $null
            $sw = [Diagnostics.Stopwatch]::StartNew()
            try {
                $store = $stores.Item($i)
                [void][string]$store.DisplayName
                [void][string]$store.StoreID
                try { [void][int]$store.ExchangeStoreType } catch { }
                try { [void][bool]$store.IsCachedExchange } catch { }
                $afterProps = $sw.Elapsed.TotalMilliseconds
                $root = $store.GetRootFolder()
                [void][string]$root.EntryID
                $total = $sw.Elapsed.TotalMilliseconds
                $propsMs.Add($afterProps)
                if ($pass -eq 0) { $first = $total - $afterProps } else { $later.Add($total - $afterProps) }
            }
            catch {
                $failedAt = $sw.Elapsed.TotalMilliseconds
                $errorText = (Get-HResult $_) + ' after ' + (Ms $failedAt) + ' ms'
                if ($pass -eq 0) { $first = $failedAt }
            }
            finally {
                Release $root
                Release $store
            }
        }
        $rows += [pscustomobject]@{ Index = $i; First = $first; Later = $later.ToArray(); Props = $propsMs.ToArray(); Error = $errorText }
    }

    $storeIdShared = Get-SharedWith @($entries | ForEach-Object { $_.StoreId })
    $uidShared = Get-SharedWith @($entries | ForEach-Object { Get-ProviderUidHex $_.RootId })
    foreach ($row in $rows) {
        $k = $row.Index - 1
        $e = $entries[$k]
        $firstText = 'n/a'
        if ($null -ne $row.First) { $firstText = Ms ([double]$row.First) }
        $errorPart = ''
        if ($row.Error) { $errorPart = ' error=' + $row.Error }
        $text = 'STORE ' + $row.Index + ' exchangeStoreType=' + $e.Type + ' cached=' + $e.Cached +
            ' storeIdHash=' + (Get-ShortHash $e.StoreId -IgnoreCase) + ' sameStoreIdAs=' + $storeIdShared[$k] +
            ' rootUidHash=' + (Get-ShortHash (Get-ProviderUidHex $e.RootId) -IgnoreCase) + ' sameRootUidAs=' + $uidShared[$k] +
            ' rootFirstMs=' + $firstText + ' rootLaterMs=' + (Format-Spread $row.Later) +
            ' alwaysReadPropsMs=' + (Format-Spread $row.Props) + $errorPart
        Out-Line $text
    }

    $uncached = @($rows | Where-Object { $null -ne $entries[$_.Index - 1].Type -and $entries[$_.Index - 1].Type -ne 3 -and $entries[$_.Index - 1].Cached -ne $true })
    $answer = 'ANSWER-ROOTS uncachedExchangeStores=' + $uncached.Count +
        ' rootFirstMs=' + (Format-Spread @($uncached | ForEach-Object { $_.First })) +
        ' rootLaterMs=' + (Format-Spread @($uncached | ForEach-Object { $_.Later })) +
        ' identity=' + (Get-IdentityVerdict $uidShared $storeIdShared)
    Out-Line $answer
}

# One GetFolderFromID attempt, described: where it opened, and whether the folder answers.
function Get-OpenLine([string] $Label, [string] $FolderIdHex, [string] $StoreIdHex) {
    $folder = $null; $owner = $null; $items = $null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($StoreIdHex) { $folder = $ns.GetFolderFromID($FolderIdHex, $StoreIdHex) } else { $folder = $ns.GetFolderFromID($FolderIdHex) }
        $opened = Ms $sw.Elapsed.TotalMilliseconds
        $where = '?'
        try { $owner = $folder.Store; $where = Get-EntryPosition $owner } catch { $where = 'storeUnreadable ' + (Get-HResult $_) }
        $sameId = $false
        $openedId = $null
        try { $openedId = [string]$folder.EntryID; $sameId = $openedId.ToUpperInvariant() -eq $FolderIdHex.ToUpperInvariant() } catch { }
        # When it opened ANOTHER folder than the one asked for: is that a store's top folder?
        $isTop = @($entries | Where-Object { $_.RootId -and $openedId -and $_.RootId.ToUpperInvariant() -eq $openedId.ToUpperInvariant() }).Count -gt 0
        $nameRead = 'ok'
        try { $nameRead = 'ok (hash ' + (Get-ShortHash ([string]$folder.Name) -IgnoreCase) + ')' } catch { $nameRead = 'FAILS ' + (Get-HResult $_) }
        $typeRead = 'ok'
        try { $typeRead = 'ok (' + [int]$folder.DefaultItemType + ')' } catch { $typeRead = 'FAILS ' + (Get-HResult $_) }
        $countRead = 'ok'
        try { $items = $folder.Items; $countRead = 'ok (' + [int]$items.Count + ')' } catch { $countRead = 'FAILS ' + (Get-HResult $_) }
        # What tells a folder Outlook hands out for an id no folder has from a real one: its parent, its
        # path, and the properties a real folder carries.
        $parentText = 'none'
        $parent = $null
        try {
            $parent = $folder.Parent
            $parentId = $null
            try { $parentId = [string]$parent.EntryID } catch { }
            $parentIsTop = @($entries | Where-Object { $_.RootId -and $parentId -and $_.RootId.ToUpperInvariant() -eq $parentId.ToUpperInvariant() }).Count -gt 0
            $parentText = $(if ($parentId) { 'a folder (a store''s top folder ' + $parentIsTop + ')' } else { 'no entry id (the namespace)' })
        }
        catch { $parentText = 'FAILS ' + (Get-HResult $_) }
        finally { Release $parent }
        $pathText = 'unreadable'
        try { $pathText = 'segments ' + @(([string]$folder.FolderPath).TrimStart('\') -split '\\').Count } catch { $pathText = 'FAILS ' + (Get-HResult $_) }
        $props = @()
        foreach ($tag in @(@('parentEntryId', '0x0E090102'), @('sourceKey', '0x65E00102'), @('recordKey', '0x0FF90102'), @('folderType', '0x36010003'), @('subfolders', '0x360A000B'), @('contentCount', '0x36020003'))) {
            try { [void]$folder.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/' + $tag[1]); $props += ($tag[0] + ' ok') } catch { $props += ($tag[0] + ' ' + (Get-HResult $_)) }
        }
        return ($Label + ' OPENED in ' + $opened + ' ms in entry ' + $where + '; same entry id ' + $sameId + '; a store''s top folder ' + $isTop + '; Name ' + $nameRead + '; DefaultItemType ' + $typeRead + '; Items.Count ' + $countRead + '; Parent ' + $parentText + '; FolderPath ' + $pathText + '; ' + ($props -join ', '))
    }
    catch {
        return ($Label + ' REFUSED ' + (Get-HResult $_) + ' after ' + (Ms $sw.Elapsed.TotalMilliseconds) + ' ms')
    }
    finally {
        Release $items
        Release $owner
        Release $folder
    }
}

if ($Section -eq 'All' -or $Section -eq 'FolderIds') {
    Out-Line '== FolderIds: which entry GetFolderFromID opens an Exchange folder id in, and what it does with an id no folder has =='
    foreach ($e in $entries) {
        $unknown = Get-UnknownExchangeFolderId $e.RootId
        if (-not $unknown) { continue }
        foreach ($case in @(@('top folder', $e.RootId), @('unknown id', $unknown))) {
            Out-Line (Get-OpenLine ('STORE ' + $e.Index + ' ' + $case[0] + ' / id alone:') $case[1] $null)
            foreach ($other in $entries) {
                Out-Line (Get-OpenLine ('STORE ' + $e.Index + ' ' + $case[0] + ' / with store ' + $other.Index + '''s StoreID:') $case[1] $other.StoreId)
            }
        }
    }
}

if ($Section -eq 'All' -or $Section -eq 'IndexHits') {
    Out-Line '== IndexHits: the test account''s newest index rows, and what the subject-and-time probe finds for each =='
    . "$PSScriptRoot\StagedLiveTestSettings.ps1"
    if (-not $Hub) { $Hub = Get-StagedSettingsValue $StagedSettingsPath 'testHubStoreDisplayName' }
    $hubEntries = @($entries | Where-Object { $_.Name -ieq $Hub })
    if (-not $Hub -or $hubEntries.Count -eq 0) {
        Out-Line 'NO-HUB: no -Hub, and the staged live-test settings name no store of this profile. Section skipped.'
    }
    else {
        $hubEntry = $hubEntries[0]
        Out-Line ('hub is store ' + $hubEntry.Index + ' (' + $hubEntries.Count + ' store(s) of that name)')
        $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $mapiRoot = 'mapi16://{' + $sid + '}/'
        $conn = New-Object -ComObject ADODB.Connection
        $conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
        try {
            # The hub's roots: one per data file the index holds under its name.
            $rs = New-Object -ComObject ADODB.Recordset
            $rs.Open("SELECT System.ItemUrl FROM SystemIndex WHERE DIRECTORY='" + $mapiRoot + "'", $conn)
            $roots = @()
            while (-not $rs.EOF) {
                $u = [string]$rs.Fields.Item('System.ItemUrl').Value
                $seg = $u.Substring([Math]::Min($u.Length, $mapiRoot.Length)).TrimEnd('/')
                if ($seg.StartsWith($Hub + '($', [StringComparison]::OrdinalIgnoreCase)) { $roots += ($mapiRoot + $seg) }
                $rs.MoveNext()
            }
            $rs.Close()
            Out-Line ('index roots under the hub''s name: ' + $roots.Count)
            foreach ($r in $roots) {
                $rs = New-Object -ComObject ADODB.Recordset
                $rs.Open("SELECT TOP 40 System.ItemUrl, System.Subject, System.Message.DateReceived, System.Kind, System.ItemType, System.Message.ConversationID FROM SystemIndex WHERE SCOPE='" + $r + "' ORDER BY System.Message.DateReceived DESC", $conn)
                $n = 0
                while (-not $rs.EOF -and $n -lt 20) {
                    # As the thread test picks its seed: a conversation id and a subject.
                    if ([string]::IsNullOrEmpty([string]$rs.Fields.Item('System.Message.ConversationID').Value)) { $rs.MoveNext(); continue }
                    $n++
                    $url = [string]$rs.Fields.Item('System.ItemUrl').Value
                    $subject = [string]$rs.Fields.Item('System.Subject').Value
                    $received = $rs.Fields.Item('System.Message.DateReceived').Value
                    $kind = @($rs.Fields.Item('System.Kind').Value) -join '|'
                    $itemType = [string]$rs.Fields.Item('System.ItemType').Value
                    $isAttachment = $url.IndexOf('/at=', [StringComparison]::OrdinalIgnoreCase) -ge 0
                    $folderNames = Get-UrlFolderSegments $url $r
                    $line = 'ROW ' + $n + ' kind=' + $kind + ' itemType=' + $itemType + ' attachmentRow=' + $isAttachment + ' folder=' + ($folderNames -join '/') + ' subjectHash=' + (Get-ShortHash $subject -IgnoreCase) + ' subjectEmpty=' + [string]::IsNullOrEmpty($subject)
                    if (-not $isAttachment -and $subject) {
                        # The locator's walk (Folders[name]) and its subject probe (Items.Restrict on the subject).
                        $store = $null; $cur = $null; $all = $null; $restricted = $null
                        try {
                            $store = $stores.Item($hubEntry.Index)
                            $cur = $store.GetRootFolder()
                            foreach ($name in $folderNames) { $next = $cur.Folders.Item($name); Release $cur; $cur = $next }
                            $quoted = $subject.Replace("'", "''")
                            $all = $cur.Items
                            $restricted = $all.Restrict('@SQL="urn:schemas:httpmail:subject" = ''' + $quoted + '''')
                            $mc = [int]$restricted.Count
                            $line += ' subjectMatches=' + $mc
                            for ($m = 1; $m -le [Math]::Min($mc, 3); $m++) {
                                $it = $null
                                try {
                                    $it = $restricted.Item($m)
                                    $cls = [string]$it.MessageClass
                                    $rt = $null; $ct = $null; $so = $null
                                    try { $rt = $it.ReceivedTime } catch { }
                                    try { $ct = $it.CreationTime } catch { }
                                    try { $so = $it.SentOn } catch { }
                                    $line += ' [' + $cls + ' receivedS=' + (Format-SecondsFrom $received $rt) + ' createdS=' + (Format-SecondsFrom $received $ct) + ' sentS=' + (Format-SecondsFrom $received $so) + ']'
                                }
                                finally { Release $it }
                            }
                        }
                        catch { $line += ' walk-or-probe FAILED ' + (Get-HResult $_) }
                        finally {
                            Release $restricted
                            Release $all
                            Release $cur
                            Release $store
                        }
                    }
                    Out-Line $line
                    $rs.MoveNext()
                }
                $rs.Close()
                Out-Line ('rows under root ' + (Get-ShortHash $r -IgnoreCase) + ': ' + $n)
            }
        }
        finally {
            $conn.Close()
        }
    }
}

$dir = Split-Path -Parent $OutPath
if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
Set-Content -LiteralPath $OutPath -Value $lines -Encoding UTF8
Release $stores
Release $ns
Release $outlook
exit 0
