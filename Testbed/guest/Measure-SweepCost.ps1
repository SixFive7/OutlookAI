<#
    ============================================================================================
    RECONSTRUCTION. THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    It replaces `Docs/v3-probes/soakfix13-probe-sweep-cost.ps1`, which step 2 of
    Docs/corpus-measurement-plan.md requires and which does not exist in this repository: the
    v3-probes directory is gitignored, so the probe lived on one machine and is gone. That
    document says it is "reconstructible from the description above plus one detail that is not
    optional", and this is that reconstruction, written from the shipped sweep's own source
    (OutlookComSession.SweepFolder) rather than from memory of the original.

    It was written by an agent that was forbidden to touch Outlook or a mailbox, so nothing here
    has been run against a store. Read it before you trust a number out of it, and once it HAS
    run, replace this banner with what it actually did.

    WHAT MAKES IT SAFE TO RUN ANYWAY: it is read-only by construction. GetTable, Columns.Add,
    Sort, GetNextRow, GetItemFromID and property reads. No Save, no Delete, no Move, no Add, no
    Send. If you extend it, keep that true - mailbox mutation from ad-hoc shell code is the thing
    that once destroyed real mail on this project.

    READ-ONLY IS NOT THE SAME AS SAFE ON THE WRONG MACHINE, SO SINCE 2026-09-24 IT IS GUARDED.
    Binding Outlook on the maintainer's workstation opens the real profile, delegate mailboxes and
    all, whether or not anything is then written. It now dot-sources OutlookMapiInterop.ps1 and
    calls Assert-TestbedGuest - the guard every writing script here uses - before anything else:
    before -OutFile is deleted and rewritten, and before COM. STAGE OutlookMapiInterop.ps1 BESIDE
    IT. Proven on the maintainer's workstation the same day, under Windows PowerShell 5.1: with
    and without -OpenItems it stopped at "REFUSING TO RUN. This session is logged on as ..." with
    zero calls reaching a tripwire that stood in for every write command and New-Object, and no
    -OutFile written. That proves the refusal only; on a guest it is still never executed.

    WHY IT EXISTS AT ALL. The server reports one clock for the whole sweep (sweep.elapsedMs) and
    no per-folder or per-item timing. The sweep budget is per-item cost x items x folders x
    stores and nothing else, so the per-item cost is the single most useful number in the
    measurement plan - and it is also the fallback route for the whole document if placement
    cannot be made to work, because it measures the cost model's coefficients directly rather
    than measuring the shipped sweep.

    Run it with -OpenItems both off and on. The difference, divided by the row count, is the
    per-item cost that the 19 ms-per-folder + 15 ms-per-item model claims.

    THREE DETAILS THAT ARE NOT OPTIONAL, each learned the expensive way:

    1. THE DATE LITERAL IS YEAR-FIRST, 'yyyy-MM-dd HH:mm:ss'. Outlook parses a DASL date literal
       in the MACHINE locale. An invariant US MM/dd/yyyy literal on a day-first box transposes
       day and month for roughly 40% of dates and answers about a different window - silently,
       in both directions. An ISO literal with a 'T' separator is worse: it does not throw, it
       returns the WHOLE FOLDER.
    2. SORT BY THE EXPLICIT NAME, NOT THE NAMESPACE. Table.Sort accepts "explicit string names
       only; cannot reference properties by their namespaces". A live probe over five stores
       found Sort("ReceivedTime") applied 5 of 5 and Sort("urn:schemas:httpmail:datereceived")
       refused 5 of 5. The shipped sweep passed the namespace form for the life of the feature,
       so its 200-item cap always cut an arbitrary slice. A probe that reproduces that bug
       measures the wrong thing.
    3. POWERSHELL CANNOT DRIVE THESE COM OBJECTS BY LATE BINDING. Table.GetRows, Table.Sort and
       CSearchManager all fail in a way that reads like the API refusing. Every call on a Table
       here therefore goes through InvokeMember. If you see "method not found" on a member that
       plainly exists, this is why.

    Sent Items sorts by SentOn first, not ReceivedTime: mail a person sent was never received, so
    an item admitted by the submit-time clause with no delivery time sorts OLDEST under
    ReceivedTime and is the first thing the cap drops - the opposite of what a freshness tier is
    for. The shipped sweep makes the same distinction.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $Store,
    [int]    $WindowDays = 7,
    [int]    $Cap = 200,
    [switch] $OpenItems,
    [int]    $Repeat = 3,
    [string] $OutFile = 'C:\OutlookAI-Q5\sweep-cost.txt',
    # The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.
    [string[]] $ExpectedUser = @('vmadmin')
)

$ErrorActionPreference = 'Stop'

# THE GUARD, FIRST - before -OutFile is touched and before COM. Read-only is not enough on the
# wrong machine: binding Outlook on the workstation opens the real profile. See the banner.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

# Folder ids the shipped sweep covers. Drafts is deliberately absent: the sweep does not cover
# it, which is why a corpus accidentally filed as drafts measured as an empty store.
# StandIn is the folder the corpus builder makes at the store root when the store HAS no such
# default folder (CorpusFolderIds.StandInName) - Corpus A on OutlookAI-Indexed has three (D101).
$folderKinds = @(
    @{ Id = 6;  Name = 'Inbox';         StandIn = 'OutlookAI-Corpus-Folder-6';    Sort = @('ReceivedTime', 'urn:schemas:httpmail:datereceived') }
    @{ Id = 5;  Name = 'Sent Items';    StandIn = 'OutlookAI-Corpus-Folder-5';    Sort = @('SentOn', 'ReceivedTime', 'urn:schemas:httpmail:date', 'urn:schemas:httpmail:datereceived') }
    @{ Id = 3;  Name = 'Deleted Items'; StandIn = 'OutlookAI-Corpus-Folder-3';    Sort = @('ReceivedTime', 'urn:schemas:httpmail:datereceived') }
    @{ Id = 23; Name = 'Junk Email';    StandIn = 'OutlookAI-Corpus-Folder-Junk'; Sort = @('ReceivedTime', 'urn:schemas:httpmail:datereceived') }
)

# NEVER THE CREATING LOOKUP (added 2026-10-03, D101 follow-up). Store.GetDefaultFolder on a PST
# that lacks the folder CREATES it - measured on a POP3 PST for Junk Email and Archive - or hands
# back the store's nameless non-IPM root for an Inbox it does not have (OAI-UNINDEXED, 2026-09-24).
# So a "read-only" measurement that called it could add a folder to a store the count tripwire
# watches, and would time the wrong folder. This resolves a default folder the way the shipped
# sweep does (SpecialFolders.Resolve, Q84): on a store that is not Exchange, Inbox, Sent Items and
# Deleted Items only when the store's PR_VALID_FOLDER_MASK says it has them, Junk Email only from
# its designation on the Inbox (PR_ADDITIONAL_REN_ENTRYIDS, index 4), opened by entry id - and
# nothing at all when a designation will not read, which is reported, never guessed.
$PrValidFolderMask = 'http://schemas.microsoft.com/mapi/proptag/0x35DF0003'
$PrAdditionalRenEntryIds = 'http://schemas.microsoft.com/mapi/proptag/0x36D81102'
$ValidFolderBits = @{ 6 = 0x02; 3 = 0x08; 5 = 0x10 }
$MapiNotFound = -2147221233   # 0x8004010F, MAPI_E_NOT_FOUND: the property is not there

# Every one of these returns its value with the unary comma - "return , $x" - so PowerShell hands the
# caller the COM object itself. A plain "return $x" ENUMERATES anything enumerable on the way out: the
# Table's Columns collection reached the caller as an object[] of columns, and the first InvokeMember on
# it - Columns.Count - failed with "Method 'System.Object[].Count' not found". Measured on
# OutlookAI-Indexed, 2026-10-03, the first time this script ever ran; every Columns.Add before it had
# failed the same way inside its try, silently, so no table would have been sorted either.
function Invoke-Com {
    param($Target, [string] $Name, [System.Reflection.BindingFlags] $Flags, [object[]] $Arguments = @())
    return , $Target.GetType().InvokeMember($Name, $Flags, $null, $Target, $Arguments)
}
function Get-ComProperty { param($Target, [string] $Name, [object[]] $Arguments = @())
    return , (Invoke-Com -Target $Target -Name $Name -Flags ([System.Reflection.BindingFlags]::GetProperty) -Arguments $Arguments)
}
function Invoke-ComMethod { param($Target, [string] $Name, [object[]] $Arguments = @())
    return , (Invoke-Com -Target $Target -Name $Name -Flags ([System.Reflection.BindingFlags]::InvokeMethod) -Arguments $Arguments)
}

function Write-Line { param([string] $Text)
    Write-Host $Text
    Add-Content -LiteralPath $OutFile -Value $Text
}

<# PropertyAccessor.GetProperty through InvokeMember (detail 3). Throws what Outlook throws. #>
function Get-MapiProperty { param($Target, [string] $Schema)
    $accessor = Get-ComProperty -Target $Target -Name 'PropertyAccessor'
    try { return Invoke-ComMethod -Target $accessor -Name 'GetProperty' -Arguments @($Schema) }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($accessor) }
}

<# Whether an exception thrown through InvokeMember is MAPI_E_NOT_FOUND - "not there", not "would not read". #>
function Test-MapiNotFound { param($ErrorRecord)
    $e = $ErrorRecord.Exception
    while ($null -ne $e) { if ($e.HResult -eq $MapiNotFound) { return $true }; $e = $e.InnerException }
    return $false
}

<#
    One default folder, WITHOUT ever creating it (see the comment at $PrValidFolderMask). Returns
    @{ Folder; State = 'resolved' | 'absent' | 'unreadable'; Why }. Folder is set only when resolved.
#>
function Resolve-DefaultFolderReadOnly {
    param($TargetStore, $Namespace, [string] $StoreEntryId, [int] $Id)
    $exchangeType = $null
    try { $exchangeType = [int](Get-ComProperty -Target $TargetStore -Name 'ExchangeStoreType') } catch { }
    if ($null -ne $exchangeType -and $exchangeType -ne 3) {
        # An Exchange mailbox: its default folders are the server's, and the shipped sweep asks for them.
        try { return @{ Folder = $TargetStore.GetDefaultFolder($Id); State = 'resolved'; Why = 'Exchange server default folder' } }
        catch { return @{ Folder = $null; State = 'unreadable'; Why = "GetDefaultFolder failed: $($_.Exception.Message)" } }
    }

    if ($Id -eq 23) {
        $inbox = Resolve-DefaultFolderReadOnly -TargetStore $TargetStore -Namespace $Namespace -StoreEntryId $StoreEntryId -Id 6
        if ($inbox.State -ne 'resolved') {
            return @{ Folder = $null; State = $inbox.State; Why = "Junk Email is designated on the Inbox, which is $($inbox.State)" }
        }
        $ids = $null
        try { $ids = Get-MapiProperty -Target $inbox.Folder -Schema $PrAdditionalRenEntryIds }
        catch {
            if (Test-MapiNotFound $_) { return @{ Folder = $null; State = 'absent'; Why = 'the Inbox designates no Junk Email' } }
            return @{ Folder = $null; State = 'unreadable'; Why = "the Inbox's designations would not read: $($_.Exception.Message)" }
        }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($inbox.Folder) }
        $entries = @($ids)
        if ($entries.Count -le 4 -or $null -eq $entries[4] -or @($entries[4]).Count -eq 0) {
            return @{ Folder = $null; State = 'absent'; Why = 'the Inbox designates no Junk Email' }
        }
        $hex = ([BitConverter]::ToString([byte[]]$entries[4])) -replace '-', ''
        try { return @{ Folder = $Namespace.GetFolderFromID($hex, $StoreEntryId); State = 'resolved'; Why = 'designated on the Inbox' } }
        catch { return @{ Folder = $null; State = 'unreadable'; Why = "its designated entry id would not open: $($_.Exception.Message)" } }
    }

    $mask = $null
    try { $mask = [int](Get-MapiProperty -Target $TargetStore -Schema $PrValidFolderMask) }
    catch { return @{ Folder = $null; State = 'unreadable'; Why = "the store's PR_VALID_FOLDER_MASK would not read: $($_.Exception.Message)" } }
    if (($mask -band $ValidFolderBits[$Id]) -eq 0) {
        return @{ Folder = $null; State = 'absent'; Why = ('PR_VALID_FOLDER_MASK 0x{0:X2} has no bit for it' -f $mask) }
    }
    # The bit is set, so the folder is there and GetDefaultFolder returns it rather than making one.
    try { return @{ Folder = $TargetStore.GetDefaultFolder($Id); State = 'resolved'; Why = ('PR_VALID_FOLDER_MASK 0x{0:X2}' -f $mask) } }
    catch { return @{ Folder = $null; State = 'unreadable'; Why = "GetDefaultFolder failed: $($_.Exception.Message)" } }
}

<# The corpus builder's stand-in for a default folder: a folder of that name at the store's root, or $null. #>
function Find-StandInFolder { param($TargetStore, [string] $Name)
    $root = $TargetStore.GetRootFolder()
    $children = $root.Folders
    try {
        foreach ($child in $children) {
            if ([string]::Equals([string]$child.Name, $Name, [System.StringComparison]::Ordinal)) { return $child }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($child)
        }
        return $null
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($children)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($root)
    }
}

Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue

$sinceUtc = (Get-Date).ToUniversalTime().AddDays(-$WindowDays)
# DaslDateLiteral.FormatUtc. Detail 1 above; do not "simplify" this to an ISO 'T' form.
$literal = $sinceUtc.ToString('yyyy-MM-dd HH:mm:ss')
$filter = "@SQL=(""urn:schemas:httpmail:datereceived"" >= '$literal') OR (""urn:schemas:httpmail:date"" >= '$literal')"

Write-Line "store        : $Store"
Write-Line "window       : $WindowDays day(s), since $literal UTC"
Write-Line "cap          : $Cap rows per folder"
Write-Line "openItems    : $($OpenItems.IsPresent)"
Write-Line "filter       : $filter"
Write-Line ''

$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('MAPI')

$target = $null
foreach ($s in $ns.Stores) {
    if ($s.DisplayName -eq $Store) { $target = $s; break }
}
if ($null -eq $target) {
    $names = ($ns.Stores | ForEach-Object { $_.DisplayName }) -join ', '
    throw "No store named '$Store'. Stores on this profile: $names"
}
$storeId = $target.StoreID

# What gets timed: each default folder the shipped sweep would walk, resolved without creating it,
# and - when the store has one - the corpus builder's stand-in for it, LABELLED as such: the shipped
# sweep never walks a stand-in, so its rows say what a folder of that size costs, not what the
# product's sweep of this store costs. Every skip is said, with its reason.
$targets = New-Object System.Collections.Generic.List[object]
foreach ($kind in $folderKinds) {
    $resolved = Resolve-DefaultFolderReadOnly -TargetStore $target -Namespace $ns -StoreEntryId $storeId -Id $kind.Id
    if ($resolved.State -eq 'resolved') {
        $targets.Add(@{ Label = $kind.Name; Folder = $resolved.Folder; Sort = $kind.Sort })
        Write-Line ("{0,-22} {1}" -f $kind.Name, "the store's own default folder ($($resolved.Why)) - the shipped sweep walks it")
    }
    else {
        Write-Line ("{0,-22} {1}" -f $kind.Name, "$($resolved.State.ToUpperInvariant()) ($($resolved.Why)) - the shipped sweep skips it, and so does this")
    }
    $standIn = Find-StandInFolder -TargetStore $target -Name $kind.StandIn
    if ($null -ne $standIn) {
        $targets.Add(@{ Label = "$($kind.Name) stand-in"; Folder = $standIn; Sort = $kind.Sort })
        Write-Line ("{0,-22} {1}" -f "$($kind.Name) stand-in", "$($kind.StandIn) at the store root - the corpus builder's, NOT walked by the shipped sweep")
    }
}
Write-Line ''

Write-Line ("{0,-22} {1,5} {2,8} {3,10} {4,8} {5,10}" -f 'folder', 'pass', 'rows', 'ms', 'sorted', 'ms/row')

foreach ($entry in $targets) {
    $kind = @{ Name = $entry.Label; Sort = $entry.Sort }
    $folder = $entry.Folder

    for ($pass = 1; $pass -le $Repeat; $pass++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $rows = 0
        $sorted = $false
        $unreadable = 0

        $table = $folder.GetTable($filter)

        # Detail 2: pair each Columns.Add with a Sort under the SAME spelling, explicit names
        # first, and take the first spelling that both goes on and orders.
        foreach ($property in $kind.Sort) {
            $columns = Get-ComProperty -Target $table -Name 'Columns'
            try { Invoke-ComMethod -Target $columns -Name 'Add' -Arguments @($property) | Out-Null }
            catch { continue }
            try {
                Invoke-ComMethod -Target $table -Name 'Sort' -Arguments @($property, $true) | Out-Null
                $sorted = $true
                break
            }
            catch { }
        }

        # Column indices are 1-based on Columns and 0-based in GetValues(), which is the offset
        # the shipped FindTableColumn applies.
        $entryIdIndex = -1
        $columns = Get-ComProperty -Target $table -Name 'Columns'
        $columnCount = [int](Get-ComProperty -Target $columns -Name 'Count')
        for ($i = 1; $i -le $columnCount; $i++) {
            $column = Get-ComProperty -Target $columns -Name 'Item' -Arguments @($i)
            if ((Get-ComProperty -Target $column -Name 'Name') -ieq 'EntryID') { $entryIdIndex = $i - 1; break }
        }
        if ($entryIdIndex -lt 0) { throw "The table for $($kind.Name) carries no EntryID column." }

        while ((-not [bool](Get-ComProperty -Target $table -Name 'EndOfTable')) -and $rows -lt $Cap) {
            $row = Invoke-ComMethod -Target $table -Name 'GetNextRow'
            $values = Invoke-ComMethod -Target $row -Name 'GetValues'
            $entryId = $values[$entryIdIndex]
            if ([string]::IsNullOrEmpty($entryId)) { $unreadable++; continue }
            $rows++

            if ($OpenItems) {
                # This half is what the real sweep pays on top of the table walk: one
                # GetItemFromID per row, then property reads off the item.
                $item = $ns.GetItemFromID($entryId, $storeId)
                $null = $item.Subject
                $null = $item.ReceivedTime
                $null = $item.SenderName
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
            }
        }

        $sw.Stop()
        $perRow = 0
        if ($rows -gt 0) { $perRow = [math]::Round($sw.Elapsed.TotalMilliseconds / $rows, 2) }
        Write-Line ("{0,-22} {1,5} {2,8} {3,10} {4,8} {5,10}" -f $kind.Name, $pass, $rows, $sw.ElapsedMilliseconds, $sorted, $perRow)
        if ($unreadable -gt 0) { Write-Line ("{0,-22} {1}" -f '', "$unreadable row(s) named no item") }
    }
}

Write-Line ''
Write-Line 'Run again with the opposite -OpenItems. The difference divided by the row count is'
Write-Line 'the per-item cost; the remainder over the folder count is the per-folder fixed cost.'
Write-Line "Output: $OutFile"
