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
    [string] $OutFile = 'C:\OutlookAI-Q5\sweep-cost.txt'
)

$ErrorActionPreference = 'Stop'

# Folder ids the shipped sweep covers. Drafts is deliberately absent: the sweep does not cover
# it, which is why a corpus accidentally filed as drafts measured as an empty store.
$folderKinds = @(
    @{ Id = 6;  Name = 'Inbox';         Sort = @('ReceivedTime', 'urn:schemas:httpmail:datereceived') }
    @{ Id = 5;  Name = 'Sent Items';    Sort = @('SentOn', 'ReceivedTime', 'urn:schemas:httpmail:date', 'urn:schemas:httpmail:datereceived') }
    @{ Id = 3;  Name = 'Deleted Items'; Sort = @('ReceivedTime', 'urn:schemas:httpmail:datereceived') }
    @{ Id = 23; Name = 'Junk Email';    Sort = @('ReceivedTime', 'urn:schemas:httpmail:datereceived') }
)

function Invoke-Com {
    param($Target, [string] $Name, [System.Reflection.BindingFlags] $Flags, [object[]] $Arguments = @())
    return $Target.GetType().InvokeMember($Name, $Flags, $null, $Target, $Arguments)
}
function Get-ComProperty { param($Target, [string] $Name, [object[]] $Arguments = @())
    return Invoke-Com -Target $Target -Name $Name -Flags ([System.Reflection.BindingFlags]::GetProperty) -Arguments $Arguments
}
function Invoke-ComMethod { param($Target, [string] $Name, [object[]] $Arguments = @())
    return Invoke-Com -Target $Target -Name $Name -Flags ([System.Reflection.BindingFlags]::InvokeMethod) -Arguments $Arguments
}

function Write-Line { param([string] $Text)
    Write-Host $Text
    Add-Content -LiteralPath $OutFile -Value $Text
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

Write-Line ("{0,-16} {1,5} {2,8} {3,10} {4,8} {5,10}" -f 'folder', 'pass', 'rows', 'ms', 'sorted', 'ms/row')

foreach ($kind in $folderKinds) {
    $folder = $null
    try { $folder = $target.GetDefaultFolder($kind.Id) } catch { }
    if ($null -eq $folder) {
        Write-Line ("{0,-16} {1}" -f $kind.Name, 'no such default folder in this store - skipped')
        continue
    }

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
        Write-Line ("{0,-16} {1,5} {2,8} {3,10} {4,8} {5,10}" -f $kind.Name, $pass, $rows, $sw.ElapsedMilliseconds, $sorted, $perRow)
        if ($unreadable -gt 0) { Write-Line ("{0,-16} {1}" -f '', "$unreadable row(s) named no item") }
    }
}

Write-Line ''
Write-Line 'Run again with the opposite -OpenItems. The difference divided by the row count is'
Write-Line 'the per-item cost; the remainder over the folder count is the per-folder fixed cost.'
Write-Line "Output: $OutFile"
