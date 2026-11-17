#Requires -Version 5.1
<#
.SYNOPSIS
    READ-ONLY measurement for Q132: does the route that turns a local mail file's folder id into its
    place in the Windows Search index - the folder's own index row, found by System.ProviderItemID - work
    for the test account's CACHED EXCHANGE store too?

.DESCRIPTION
    RUN ON THE GUEST OutlookAI-Exchange, in session 1 through Testbed/guest/Register-InteractiveTask.ps1
    -RunLevel Limited, with Outlook running. Windows PowerShell 5.1. It reads the test account's store
    through the object model and the index through Windows Search's OLE DB provider; it writes nothing
    but its own result file.

    THE ROUTE, as measured on the PST guests (Docs/live-tier-on-the-vm.md, Q115): no index row holds a
    folder's whole id, but every row holds its OWN node id in System.ProviderItemID ('N' and ten decimal
    digits), and a PST folder id's last four bytes ARE that node id - so a PST folder id finds its index
    row with no Outlook call. An Exchange folder id is 46 bytes ([MS-OXCDATA] 2.2.4.1: flags, the
    provider's UID, a folder type, the mailbox database's GUID and a six-byte global counter), and an
    .ost's node ids are the data file's own. So for every top-level folder of the test account's store
    it reads the folder's own index row and compares its System.ProviderItemID with what the folder id
    and the folder's own properties offer:
      id-tail      the id's last four bytes as a node id, as for a PST;
      counter      the global counter, and its low four bytes, as numbers;
      record-key   the folder's PR_RECORD_KEY read as a node id (Outlook must be asked for it).
    VERDICT: WORKS when the id alone yields every row's System.ProviderItemID; THROUGH-OUTLOOK when
    only a property Outlook has to be asked for does; DOES-NOT-WORK otherwise. Each folder's line says
    which candidate matched.

    Printed: folder names of the test account's default folders, ids and property values of folders -
    no item, no subject, nothing of any other store. The index's roots for other stores are counted,
    never listed.

    Written 2026-11-17 for the maintainer's queue (Q132), test account only.

.PARAMETER Hub
    The test account's store. Default telefonie-ai-test@xxlnet.nl.

.PARAMETER OutPath
    Where the result is written. Default C:\OutlookAI-Q5\measure\exchange-index-ids.txt.

.PARAMETER SelfTest
    The pure parts - the id layout, the node-id spelling, the verdict. No Outlook, any machine.
#>
[CmdletBinding()]
param(
    [string] $Hub = 'telefonie-ai-test@xxlnet.nl',
    [string] $OutPath = 'C:\OutlookAI-Q5\measure\exchange-index-ids.txt',
    [string[]] $ExpectedUser = @('vmadmin'),
    [string] $ExpectedComputerNamePrefix = 'OAI-EXCHANGE',
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# =============================================================================================
# PURE PARTS
# =============================================================================================

# A 46-byte Exchange folder entry id ([MS-OXCDATA] 2.2.4.1), or $null for any other length.
function ConvertFrom-ExchangeFolderId([string] $Hex) {
    if ([string]::IsNullOrEmpty($Hex) -or $Hex.Length -ne 92 -or $Hex -notmatch '^[0-9A-Fa-f]+$') { return $null }
    $b = New-Object byte[] 46
    for ($i = 0; $i -lt 46; $i++) { $b[$i] = [Convert]::ToByte($Hex.Substring($i * 2, 2), 16) }
    [uint64]$counter = 0
    for ($i = 38; $i -lt 44; $i++) { $counter = [uint64]($counter * 256 + [uint64]$b[$i]) }   # big-endian, six bytes
    return [pscustomobject]@{
        FlagsHex        = $Hex.Substring(0, 8)
        ProviderUidHex  = $Hex.Substring(8, 32)
        FolderType      = [int]$b[20] + 256 * [int]$b[21]
        DatabaseGuidHex = $Hex.Substring(44, 32)
        GlobalCounter   = $counter
        PadHex          = $Hex.Substring(88, 4)
        TailNode        = [BitConverter]::ToUInt32($b, 42)
    }
}

# System.ProviderItemID's spelling of a node id: 'N' and ten decimal digits (measured on the PST guests).
function ConvertTo-ProviderItemId([uint64] $Node) { return 'N' + $Node.ToString('D10', [Globalization.CultureInfo]::InvariantCulture) }

# The candidates one folder offers, by name.
function Get-Candidates([string] $EntryIdHex, [string] $RecordKeyHex) {
    $c = [ordered]@{}
    $id = ConvertFrom-ExchangeFolderId $EntryIdHex
    if ($id) {
        $c['id-tail'] = ConvertTo-ProviderItemId $id.TailNode
        $c['counter'] = ConvertTo-ProviderItemId $id.GlobalCounter
        $c['counter-low32'] = ConvertTo-ProviderItemId ($id.GlobalCounter -band 0xFFFFFFFF)
    }
    elseif ($EntryIdHex -and $EntryIdHex.Length -ge 8) {
        $tail = [Convert]::ToUInt32((($EntryIdHex.Substring($EntryIdHex.Length - 8) -split '(..)' | Where-Object { $_ })[3..0] -join ''), 16)
        $c['id-tail'] = ConvertTo-ProviderItemId $tail
    }
    if ($RecordKeyHex -and $RecordKeyHex.Length -ge 8) {
        $first4 = $RecordKeyHex.Substring(0, 8)
        $le = [Convert]::ToUInt32((($first4 -split '(..)' | Where-Object { $_ })[3..0] -join ''), 16)
        $c['record-key'] = ConvertTo-ProviderItemId $le
    }
    return $c
}

# Which candidate names produce the row's System.ProviderItemID.
function Get-Matches($Candidates, [string] $ProviderItemId) {
    return @($Candidates.Keys | Where-Object { $ProviderItemId -and $Candidates[$_] -eq $ProviderItemId })
}

# The verdict over every folder compared: one string per folder, the names of its matches joined by commas.
function Get-RouteVerdict([string[]] $PerFolderMatches) {
    $rows = @($PerFolderMatches)
    if ($rows.Count -eq 0) { return 'NOT-MEASURED' }
    $fromId = @('id-tail', 'counter', 'counter-low32')
    foreach ($name in $fromId) {
        if (@($rows | Where-Object { ([string]$_ -split ',') -contains $name }).Count -eq $rows.Count) { return "WORKS ($name)" }
    }
    if (@($rows | Where-Object { ([string]$_ -split ',') -contains 'record-key' }).Count -eq $rows.Count) { return 'THROUGH-OUTLOOK (record-key)' }
    return 'DOES-NOT-WORK'
}

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $Name, [bool] $Ok) { if ($Ok) { $script:p++; Write-Host "  PASS  $Name" } else { $script:f++; Write-Host "  FAIL  $Name" } }
    # flags 00000000, provider UID 16 bytes, type 0100, database GUID 16 bytes, counter 00000000ABCD, pad 0000
    $hex = '00000000' + ('11' * 16) + '0100' + ('22' * 16) + '00000000ABCD' + '0000'
    $id = ConvertFrom-ExchangeFolderId $hex
    Check 'a 46-byte id parses' ($null -ne $id)
    Check 'its folder type is read little-endian' ($id.FolderType -eq 1)
    Check 'its global counter is read big-endian, six bytes' ($id.GlobalCounter -eq 0xABCD)
    Check 'its database GUID is bytes 22 to 37' ($id.DatabaseGuidHex -eq ('22' * 16))
    Check 'a 24-byte PST id is not an Exchange folder id' ($null -eq (ConvertFrom-ExchangeFolderId ('00' * 24)))
    Check 'a node id is N and ten digits' ((ConvertTo-ProviderItemId 32898) -eq 'N0000032898')
    $c = Get-Candidates $hex '82800000AABB'
    Check 'the counter candidate is the counter''s node spelling' ($c['counter'] -eq 'N0000043981')
    Check 'the record key reads little-endian as a node id' ($c['record-key'] -eq ('N' + (0x00008082).ToString('D10')))
    $pst = Get-Candidates ('00' * 20 + '82800000') ''
    Check 'a PST-shaped id''s tail is its node id' ($pst['id-tail'] -eq ('N' + (0x00008082).ToString('D10')))
    Check 'a match names the candidate' ((@(Get-Matches $c 'N0000043981')) -contains 'counter')
    Check 'no row id, no match' ((@(Get-Matches $c '')).Count -eq 0)
    Check 'every folder by its id: WORKS' ((Get-RouteVerdict @('counter', 'counter,record-key')) -eq 'WORKS (counter)')
    Check 'every folder by its record key only: THROUGH-OUTLOOK' ((Get-RouteVerdict @('record-key', 'record-key')) -eq 'THROUGH-OUTLOOK (record-key)')
    Check 'one folder by nothing: DOES-NOT-WORK' ((Get-RouteVerdict @('record-key', '')) -eq 'DOES-NOT-WORK')
    Check 'nothing compared: NOT-MEASURED' ((Get-RouteVerdict @()) -eq 'NOT-MEASURED')
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $members = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true) | ForEach-Object { $_.Member.Extent.Text.Trim("'", '"') })
    $bad = @($members | Where-Object { $_ -match '^(Delete|Save|Move|Send|SetProperty|SetProperties|DeleteProperties|Remove|Copy|Close|Quit|Display|SaveAs|ExecuteMso)$' })
    Check ('no late-bound member that changes anything: ' + ($bad -join ', ')) ($bad.Count -eq 0)
    $commands = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -match '^(Set-|New-Item|Remove-|Rename-|Move-|Add-Content|Out-File)' })
    Check ('no cmdlet that writes: ' + ($commands -join ', ')) ($commands.Count -eq 0)
    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:p, $script:f)
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

function Assert-TestbedGuestLocal {
    $machineOk = $env:COMPUTERNAME -and $ExpectedComputerNamePrefix -and $env:COMPUTERNAME.StartsWith($ExpectedComputerNamePrefix, [StringComparison]::OrdinalIgnoreCase)
    if (-not ($ExpectedUser -contains $env:USERNAME) -or -not $machineOk) {
        throw "REFUSING TO RUN: '$env:USERNAME' on '$env:COMPUTERNAME' is not the Exchange test VM (allowed users: $($ExpectedUser -join ', '); computer prefix '$ExpectedComputerNamePrefix')."
    }
}
Assert-TestbedGuestLocal

$lines = New-Object System.Collections.Generic.List[string]
function Say([string] $Text) { $lines.Add($Text); Write-Host $Text }
Say ("MEASURE-EXCHANGE-INDEX-IDS on {0} at {1}" -f $env:COMPUTERNAME, (Get-Date).ToString('o'))

function Get-BinaryHex($Accessor, [string] $Tag) {
    try { return [string]$Accessor.BinaryToString($Accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/' + $Tag)) } catch { return $null }
}

try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $store = $null
    foreach ($s in $app.Session.Stores) { if ([string]$s.DisplayName -ieq $Hub) { $store = $s } }
    if (-not $store) { throw "no store named '$Hub'" }
    Say ("store '{0}': ExchangeStoreType {1}, cached {2}, data file {3}" -f $Hub, [int]$store.ExchangeStoreType, $store.IsCachedExchange, $(if ($store.FilePath) { Split-Path -Leaf $store.FilePath } else { '-' }))
    $top = $store.GetRootFolder()
    $folders = @()
    foreach ($f in $top.Folders) {
        $folders += [pscustomobject]@{ Name = [string]$f.Name; EntryId = [string]$f.EntryID; RecordKey = (Get-BinaryHex $f.PropertyAccessor '0x0FF90102'); SourceKey = (Get-BinaryHex $f.PropertyAccessor '0x65E00102'); Items = [int]$f.Items.Count }
    }
    Say ("top-level folders: {0}" -f $folders.Count)

    # The index: every row under this user's mapi16 scope, grouped by root; only this store's roots are read further.
    $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $scope = 'mapi16://{' + $sid + '}/'
    $conn = New-Object -ComObject ADODB.Connection
    $conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
    $rs = New-Object -ComObject ADODB.Recordset
    $rs.Open("SELECT System.ItemUrl, System.ProviderItemID, System.ItemType, System.ItemNameDisplay FROM SystemIndex WHERE SCOPE='" + $scope + "'", $conn)
    $rows = @(); $otherRoots = @{}
    while (-not $rs.EOF) {
        $url = [string]$rs.Fields.Item('System.ItemUrl').Value
        $rest = $url.Substring([Math]::Min($url.Length, $scope.Length))
        $root = ($rest -split '/')[0]
        if ($root.StartsWith($Hub + '(', [StringComparison]::OrdinalIgnoreCase)) {
            $rows += [pscustomobject]@{ Root = $root; Path = $rest.Substring($root.Length); ProviderItemId = [string]$rs.Fields.Item('System.ProviderItemID').Value; Type = [string]$rs.Fields.Item('System.ItemType').Value; Name = [string]$rs.Fields.Item('System.ItemNameDisplay').Value }
        }
        else { $otherRoots[$root] = 1 }
        $rs.MoveNext()
    }
    $rs.Close(); $conn.Close()
    $roots = @($rows | Group-Object Root)
    Say ("index: {0} root(s) for this store, {1} row(s); {2} root(s) of other stores, not read" -f $roots.Count, $rows.Count, $otherRoots.Count)
    foreach ($g in $roots) {
        $types = ($g.Group | Group-Object Type | ForEach-Object { "'$($_.Name)' $($_.Count)" }) -join ', '
        Say ("  root {0}: {1} row(s) - by System.ItemType: {2}" -f $g.Name, $g.Count, $types)
    }

    foreach ($g in $roots) {
        Say ("== root {0} ==" -f $g.Name)
        $perFolder = @()
        foreach ($folder in $folders) {
            $own = @($g.Group | Where-Object { $_.Name -eq $folder.Name -and ($_.Path -match '^/0/[^/]+/?$') -and $_.Type -match '(?i)folder' })
            if ($own.Count -eq 0) { $own = @($g.Group | Where-Object { $_.Name -eq $folder.Name -and ($_.Path -match '^/0/[^/]+/?$') }) }
            $candidates = Get-Candidates $folder.EntryId $folder.RecordKey
            $layout = ConvertFrom-ExchangeFolderId $folder.EntryId
            $idText = if ($layout) { "46-byte id, type $($layout.FolderType), counter $($layout.GlobalCounter)" } else { "$($folder.EntryId.Length / 2)-byte id" }
            if ($own.Count -eq 0) { Say ("  '{0}' ({1} items, {2}): no own row at /0/<name>" -f $folder.Name, $folder.Items, $idText); continue }
            $rowId = $own[0].ProviderItemId
            $matched = @(Get-Matches $candidates $rowId)
            $perFolder += ($matched -join ',')
            Say ("  '{0}' ({1} items, {2}): row ProviderItemID {3}; record key {4}; candidates {5}; MATCHED {6}" -f $folder.Name, $folder.Items, $idText, $(if ($rowId) { $rowId } else { '(none)' }),
                $(if ($folder.RecordKey) { $folder.RecordKey } else { '-' }),
                (($candidates.Keys | ForEach-Object { "$_=$($candidates[$_])" }) -join ' '), $(if ($matched.Count) { $matched -join ',' } else { 'nothing' }))
        }
        Say ("VERDICT root {0}: {1} - over {2} folder(s) whose own row was found" -f $g.Name, (Get-RouteVerdict $perFolder), $perFolder.Count)
    }
}
catch { Say ('FAILED: ' + $_.Exception.Message) }
finally {
    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    [System.IO.File]::WriteAllLines($OutPath, $lines)
}
