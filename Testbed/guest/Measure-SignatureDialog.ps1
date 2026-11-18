#Requires -Version 5.1
<#
.SYNOPSIS
    READ-ONLY: what Outlook's own Signatures dialog lets a user do under the signature setting the live
    runner's -SignatureSettings step made (Q156, Q200, Q201; decided 2026-11-18): whether its default
    choice - the New messages and Replies/forwards boxes - can be changed, and whether its Delete
    button is open for each test signature.

.DESCRIPTION
    RUN ON AN OUTLOOK TEST GUEST, in session 1 at Outlook's own integrity level, with Outlook running
    and showing its main window. Testbed/host/Invoke-LiveTierOnGuest.ps1 -SignatureSettings runs it
    there, in the guest task that then runs the suite: after the step made the setting and Outlook
    started. Windows PowerShell 5.1.

    It opens the dialog by Outlook's own command - ExecuteMso('SignaturesStationeryDialog'), from a
    helper process (OutlookDialogUia.ps1) - reads it through UI Automation, selects each test signature
    (OutlookAI-McpTest-) in its list in turn to read whether Delete is enabled for it, and closes the
    dialog as Cancel does - Escape, posted to it (OutlookDialogUia.ps1). It presses nothing but Cancel,
    changes no box and types nothing, so the dialog keeps nothing; and it reads the dialog only once
    the dialog has settled - its signature list enabled or holding signatures - or after -SettleSeconds,
    saying which. The signature folder's files, every profile's per-account signature defaults and the
    two MailSettings keys are read before and after; a change is reported and fails the reading. (On
    2026-11-18 the first readings closed the dialog with WM_CLOSE, and afterwards the hub account's own
    defaults were no longer the ones the step made, in three runs of three - read by the live test that
    followed, which these readings now do themselves.)

    The reading is a JSON file, -OutPath, which the live test
    LiveSignatureSettingsTests.SetBeforeOutlookStarted_TheSignaturesDialog_... reads through
    OUTLOOKAI_LIVE_SIGNATURE_DIALOG. It holds OutlookAI to the dialog: a default the dialog will not
    let a user change, manage_signature refuses to change; a signature the dialog will not let a user
    delete, manage_signature refuses to delete.

    EXIT CODES: 0 read and closed, the folder unchanged; 1 not read, not closed, or the folder changed;
    4 refused - not a test guest, or no Outlook in this session. The reading is written in every case.

.PARAMETER OutPath
    Where the reading is written. Default C:\OutlookAI-Q5\measure\signature-dialog.json.

.PARAMETER DialogWaitSeconds
    How long the dialog may take to open.

.PARAMETER SettleSeconds
    How long the opened dialog may take to settle - its signature list enabled or holding a signature -
    before it is read as it is.

.PARAMETER SelfTest
    Pure: the choices made from a reading, and the read-only proof over this file's and
    OutlookDialogUia.ps1's syntax trees. No Outlook, any machine.
#>
[CmdletBinding()]
param(
    [string]   $OutPath = 'C:\OutlookAI-Q5\measure\signature-dialog.json',
    [int]      $DialogWaitSeconds = 30,
    [int]      $SettleSeconds = 20,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'
# Outlook's dialogs through UI Automation, and the signature folder: one copy, shared with Measure-SignatureRoaming.ps1 (Q154).
. (Join-Path $PSScriptRoot 'OutlookDialogUia.ps1')
# The guest guard every testbed script calls before it touches Outlook (Assert-TestbedGuest).
. (Join-Path $PSScriptRoot 'OutlookMapiInterop.ps1')

$TestPrefix = 'OutlookAI-McpTest-'
$DialogName = 'Signatures and Stationery'

# The dialog's controls this reads, by name: an SDM dialog (bosa_sdm_Mso96) names a box after its label.
$AccountBoxPattern = '(?i)^\s*e-?mail account'
$NewMessagesPattern = '(?i)^\s*new messages'
$RepliesPattern = '(?i)^\s*replies'
$DeleteButtonPattern = '(?i)^\s*delete\s*$'
$NewButtonPattern = '(?i)^\s*new\s*$'
$SignatureListPattern = '(?i)^\s*select signature to edit'

# =============================================================================================
# PURE PARTS. -SelfTest pins them.
# =============================================================================================

# The control a reading means by a name: of the records whose name matches, a ComboBox first - a label of
# the same name is read as Text - else the first that is not Text; $null for none. A record: Type, Name,
# Enabled, Value.
function Select-DialogControl([object[]] $Controls, [string] $NamePattern, [string] $PreferredType = 'ComboBox') {
    $named = @($Controls | Where-Object { $_ -and $_.Name -and ([string]$_.Name) -match $NamePattern })
    $preferred = @($named | Where-Object { $_.Type -eq $PreferredType })
    if ($preferred.Count -gt 0) { return $preferred[0] }
    $other = @($named | Where-Object { $_.Type -ne 'Text' })
    if ($other.Count -gt 0) { return $other[0] }
    return $null
}

# What changed in the signature folder between two readings (relative path -> SHA-256): added, removed or
# changed files, sorted. Empty when nothing did.
function Get-FolderChanges([hashtable] $Before, [hashtable] $After) {
    $changes = New-Object System.Collections.Generic.List[string]
    foreach ($k in @($Before.Keys | Sort-Object)) {
        if (-not $After.ContainsKey($k)) { $changes.Add("removed $k") }
        elseif ([string]$After[$k] -cne [string]$Before[$k]) { $changes.Add("changed $k") }
    }
    foreach ($k in @($After.Keys | Sort-Object)) {
        if (-not $Before.ContainsKey($k)) { $changes.Add("added $k") }
    }
    return @($changes)
}

# The exit code a reading earns: 4 refused; 1 when the dialog was not read, not closed, the signature folder
# changed, or a step failed; 0 otherwise. Changed signature defaults are read and reported, not failed: under an
# administrator's setting Outlook's dialog rewrites them as it opens (2026-11-18), with nothing pressed.
function Get-DialogExitCode($Reading) {
    if ($Reading.refused) { return 4 }
    if (-not $Reading.dialogFound -or -not $Reading.closed) { return 1 }
    if ($Reading.folderUnchanged -ne $true) { return 1 }
    if (@($Reading.errors).Count -gt 0) { return 1 }
    return 0
}

# A registry value as text that compares and reads: a string as it is, bytes - Outlook's per-account
# defaults are UTF-16 with a NUL - as 'bytes:' and their text, anything else as its string. The names it
# reads on a test guest are test signatures' and the guest's own.
function ConvertTo-ValueText($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [byte[]]) { return 'bytes:' + [System.Text.Encoding]::Unicode.GetString($Value).TrimEnd([char]0) }
    return [string]$Value
}

# What changed between two readings of the signature defaults, value by value: 'changed <name>: <before>
# -> <after>', 'added <name>: <after>', 'removed <name>: <before>', sorted. Empty when nothing did.
function Get-DefaultsChanges([hashtable] $Before, [hashtable] $After) {
    $changes = New-Object System.Collections.Generic.List[string]
    foreach ($k in @($Before.Keys | Sort-Object)) {
        if (-not $After.ContainsKey($k)) { $changes.Add("removed ${k}: '$($Before[$k])'") }
        elseif ([string]$After[$k] -cne [string]$Before[$k]) { $changes.Add("changed ${k}: '$($Before[$k])' -> '$($After[$k])'") }
    }
    foreach ($k in @($After.Keys | Sort-Object)) {
        if (-not $Before.ContainsKey($k)) { $changes.Add("added ${k}: '$($After[$k])'") }
    }
    return @($changes)
}

# What may never happen in this file or OutlookDialogUia.ps1, read off their syntax trees so this text
# cannot count as what it forbids: a cmdlet that writes, a late-bound member that changes a control or an
# item, a press of anything but Cancel, a selection outside the signature list. Returns the violations.
function Get-ReadOnlyViolations([string[]] $Paths) {
    $violations = New-Object System.Collections.Generic.List[string]
    foreach ($path in $Paths) {
        $leaf = Split-Path -Leaf $path
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref] $null, [ref] $null)
        $enclosing = {
            param($node)
            $p = $node.Parent
            while ($p -and -not ($p -is [System.Management.Automation.Language.FunctionDefinitionAst])) { $p = $p.Parent }
            if ($p) { return $p.Name }
            return ''
        }
        foreach ($c in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
            $name = $c.GetCommandName()
            if (-not $name) { continue }
            if ($name -match '^(Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|Remove-Item|Rename-Item|Move-Item|Copy-Item|Clear-ItemProperty|Set-Content|Add-Content|Out-File|New-Item)$') {
                $violations.Add("${leaf}: $name")
            }
            if ($name -eq 'Invoke-Button' -and (& $enclosing $c) -ne 'Close-WithCancel') {
                $violations.Add("${leaf}: Invoke-Button outside Close-WithCancel")
            }
        }
        foreach ($m in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
            $member = $m.Member.Extent.Text.Trim("'", '"')
            $where = & $enclosing $m
            if ($member -match '^(Delete|Save|SaveAs|Move|Send|SetProperty|SetProperties|DeleteProperties|Remove|Copy|Close|Quit|Display|SetValue|Toggle|Expand|Collapse|AddToSelection|RemoveFromSelection|SetFocus|Execute|ExecuteMso)$') {
                $violations.Add("${leaf}: .$member()")
            }
            if ($member -eq 'Invoke' -and $where -ne 'Invoke-Button') { $violations.Add("${leaf}: .Invoke() outside Invoke-Button") }
            if ($member -eq 'Select' -and $where -ne 'Select-SignatureInList') { $violations.Add("${leaf}: .Select() outside Select-SignatureInList") }
        }
    }
    return @($violations)
}

function Invoke-SelfTest {
    $script:stChecks = 0
    $script:stFailures = New-Object System.Collections.Generic.List[string]
    function Check([string] $Name, $Expected, $Actual) {
        $script:stChecks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = $Expected -join '|' }
        if ($Actual -is [System.Array]) { $a = $Actual -join '|' }
        if ($e -cne $a) { $script:stFailures.Add("$Name - expected '$e', got '$a'") }
    }
    Write-Host "Measure-SignatureDialog.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Pure: no Outlook, no file."

    $controls = @(
        [pscustomobject]@{ Type = 'Text'; Name = 'New messages:'; Enabled = $true; Value = $null },
        [pscustomobject]@{ Type = 'ComboBox'; Name = 'New messages:'; Enabled = $false; Value = 'OutlookAI-McpTest-Q200-Forced-New' },
        [pscustomobject]@{ Type = 'ComboBox'; Name = 'Replies/forwards:'; Enabled = $false; Value = '(none)' },
        [pscustomobject]@{ Type = 'Button'; Name = 'Delete'; Enabled = $true; Value = $null },
        [pscustomobject]@{ Type = 'ComboBox'; Name = 'E-mail account:'; Enabled = $true; Value = 'hub@vm.invalid' }
    )
    Check 'a box is read, not its label of the same name' 'ComboBox|False' @((Select-DialogControl $controls $NewMessagesPattern).Type, (Select-DialogControl $controls $NewMessagesPattern).Enabled)
    Check 'Replies/forwards is found by its label' 'Replies/forwards:' (Select-DialogControl $controls $RepliesPattern).Name
    Check 'so is the account box' 'hub@vm.invalid' (Select-DialogControl $controls $AccountBoxPattern).Value
    Check 'the Delete button is a button' 'Button' (Select-DialogControl $controls $DeleteButtonPattern 'Button').Type
    Check 'New is not New messages' $true ($null -eq (Select-DialogControl $controls $NewButtonPattern 'Button'))
    Check 'a label alone is no box' $true ($null -eq (Select-DialogControl @($controls[0]) $NewMessagesPattern))
    Check 'nothing named is nothing' $true ($null -eq (Select-DialogControl $controls '(?i)^stationery'))

    Check 'an unchanged folder has no changes' 0 @(Get-FolderChanges @{ 'a.htm' = '1'; 'a_files\i.png' = '2' } @{ 'a.htm' = '1'; 'a_files\i.png' = '2' }).Count
    Check 'every change is named, sorted' 'removed a.htm|changed b.txt|added c.rtf' (Get-FolderChanges @{ 'a.htm' = '1'; 'b.txt' = '2' } @{ 'b.txt' = '3'; 'c.rtf' = '4' })

    $good = [pscustomobject]@{ refused = $null; dialogFound = $true; closed = $true; folderUnchanged = $true; defaultsUnchanged = $true; errors = @() }
    Check 'a reading that is complete is exit 0' 0 (Get-DialogExitCode $good)
    Check 'a refusal is exit 4' 4 (Get-DialogExitCode ([pscustomobject]@{ refused = 'no Outlook'; dialogFound = $false; closed = $false; folderUnchanged = $null; defaultsUnchanged = $null; errors = @() }))
    Check 'a dialog not found is exit 1' 1 (Get-DialogExitCode ([pscustomobject]@{ refused = $null; dialogFound = $false; closed = $false; folderUnchanged = $true; defaultsUnchanged = $true; errors = @() }))
    Check 'a dialog left open is exit 1' 1 (Get-DialogExitCode ([pscustomobject]@{ refused = $null; dialogFound = $true; closed = $false; folderUnchanged = $true; defaultsUnchanged = $true; errors = @() }))
    Check 'a changed folder is exit 1' 1 (Get-DialogExitCode ([pscustomobject]@{ refused = $null; dialogFound = $true; closed = $true; folderUnchanged = $false; defaultsUnchanged = $true; errors = @() }))
    Check 'changed defaults are reported, not failed - the dialog rewrites them under a setting' 0 (Get-DialogExitCode ([pscustomobject]@{ refused = $null; dialogFound = $true; closed = $true; folderUnchanged = $true; defaultsUnchanged = $false; errors = @() }))
    Check 'the signature list is found by its name' 'List' (Select-DialogControl @([pscustomobject]@{ Type = 'Text'; Name = 'Select signature to edit'; Enabled = $true; Value = $null }, [pscustomobject]@{ Type = 'List'; Name = 'Select signature to edit'; Enabled = $false; Value = $null }) $SignatureListPattern 'List').Type
    Check 'a step that failed is exit 1' 1 (Get-DialogExitCode ([pscustomobject]@{ refused = $null; dialogFound = $true; closed = $true; folderUnchanged = $true; defaultsUnchanged = $true; errors = @('x') }))
    Check 'a default written as UTF-16 bytes reads as its text' 'bytes:A' (ConvertTo-ValueText ([byte[]](0x41, 0, 0, 0)))
    Check 'a string default as it is, none as none' 'Own|True' @((ConvertTo-ValueText 'Own'), ($null -eq (ConvertTo-ValueText $null)))
    Check 'a changed default says what it was and became' "changed P\1\New Signature: 'bytes:A' -> 'bytes:B'|added P\2\New Signature: 'X'" (Get-DefaultsChanges @{ 'P\1\New Signature' = 'bytes:A' } @{ 'P\1\New Signature' = 'bytes:B'; 'P\2\New Signature' = 'X' })

    $shared = Join-Path $PSScriptRoot 'OutlookDialogUia.ps1'
    $violations = @(Get-ReadOnlyViolations @($PSCommandPath, $shared))
    Check ('read-only, and nothing pressed but Cancel: ' + ($violations -join '; ')) 0 $violations.Count
    $source = [System.IO.File]::ReadAllText($PSCommandPath)
    Check 'the dialog is closed as Cancel does' $true $source.Contains('Close-WithCancel $dialog')
    Check 'and opened only by Outlook''s own command' $true $source.Contains('Start-SignaturesDialog')

    if ($script:stFailures.Count -gt 0) {
        foreach ($f in $script:stFailures) { Write-Host "FAIL: $f" }
        Write-Host "SelfTest: $($script:stChecks - $script:stFailures.Count) passed, $($script:stFailures.Count) failed."
        return 1
    }
    Write-Host "SelfTest: $($script:stChecks) passed, 0 failed."
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE READING
# =============================================================================================

function ConvertTo-ControlRecord($Element) {
    $type = $null; $name = $null; $enabled = $null; $value = $null
    try { $type = $Element.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '' } catch { }
    try { $name = [string]$Element.Current.Name } catch { }
    try { $enabled = [bool]$Element.Current.IsEnabled } catch { }
    $pattern = $null
    try {
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref] $pattern)) { $value = [string]$pattern.Current.Value }
    }
    catch { }
    if ($null -eq $value) {
        try {
            if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern, [ref] $pattern)) {
                $selected = @($pattern.Current.GetSelection())
                if ($selected.Count -gt 0) { $value = [string]$selected[0].Current.Name }
            }
        }
        catch { }
    }
    return [pscustomobject]@{ Type = $type; Name = $name; Enabled = $enabled; Value = $value; Element = $Element }
}

# A record as the reading writes it: no element.
function ConvertTo-Plain($Record) {
    if ($null -eq $Record) { return $null }
    return [ordered]@{ type = $Record.Type; name = $Record.Name; enabled = $Record.Enabled; value = $Record.Value }
}

# Selects one signature in the dialog's list, as a click on its name does: what the dialog shows, not what
# it keeps. Returns whether it reads as selected afterwards.
function Select-SignatureInList($Item) {
    $pattern = $Item.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $pattern.Select()
    Start-Sleep -Seconds 1
    return [bool]$Item.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Current.IsSelected
}

# Every profile's per-account signature defaults ('New Signature', 'Reply-Forward Signature' under each
# profile's 9375CFF0413111d3B88A00104B2A6676 entries) and the MailSettings keys' NewSignature and
# ReplySignature, as name -> text. Read-only.
function Get-SignatureDefaultValues {
    $values = @{}
    $profiles = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'
    foreach ($profile in @(Get-ChildItem -LiteralPath $profiles -ErrorAction SilentlyContinue)) {
        $accounts = Join-Path $profile.PSPath '9375CFF0413111d3B88A00104B2A6676'
        foreach ($entry in @(Get-ChildItem -LiteralPath $accounts -ErrorAction SilentlyContinue)) {
            foreach ($name in 'New Signature', 'Reply-Forward Signature') {
                $text = ConvertTo-ValueText ($entry.GetValue($name, $null))
                if ($null -ne $text) { $values["$($profile.PSChildName)\$($entry.PSChildName)\$name"] = $text }
            }
        }
    }
    foreach ($key in 'HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings', 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings') {
        $item = Get-Item -LiteralPath $key -ErrorAction SilentlyContinue
        if (-not $item) { continue }
        foreach ($name in 'NewSignature', 'ReplySignature') {
            $text = ConvertTo-ValueText ($item.GetValue($name, $null))
            if ($null -ne $text) { $values["$key\$name"] = $text }
        }
    }
    return $values
}

# Waits until the dialog has settled - its signature list enabled, or holding a signature - for at most
# $Seconds; returns the seconds it took, or $null when it did not.
function Wait-DialogSettled($Dialog, [int] $Seconds) {
    $started = Get-Date
    while (((Get-Date) - $started).TotalSeconds -lt $Seconds) {
        $list = Find-Named $Dialog ([System.Windows.Automation.ControlType]::List) 'Select signature to edit'
        if ($list) {
            $enabled = $false
            try { $enabled = [bool]$list.Current.IsEnabled } catch { }
            $items = $null
            try { $items = $list.FindFirst([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition) } catch { }
            if ($enabled -or $items) { return [Math]::Round(((Get-Date) - $started).TotalSeconds, 1) }
        }
        Start-Sleep -Milliseconds 500
    }
    return $null
}

function Get-FolderHashes {
    $signatures = Get-SignatureFiles
    $hashes = @{}
    foreach ($file in $signatures.Files) {
        $relative = $file.FullName.Substring($signatures.Folder.Length).TrimStart('\')
        $hashes[$relative] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
    }
    return $hashes
}

function Write-Reading($Reading) {
    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    [System.IO.File]::WriteAllText($OutPath, ($Reading | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
}

$reading = [ordered]@{
    probe              = 'Measure-SignatureDialog'
    measuredAt         = (Get-Date).ToString('o')
    computer           = $env:COMPUTERNAME
    refused            = $null
    dialogFound        = $false
    dialogClass        = $null
    closed             = $false
    account            = $null
    newMessages        = $null
    repliesForwards    = $null
    delete             = $null
    signatures         = @()
    otherSignatureItems = 0
    signatureList      = $null
    newButton          = $null
    itemsExposed       = $null
    settledAfterSeconds = $null
    folderUnchanged    = $null
    folderChanges      = @()
    defaultsUnchanged  = $null
    defaultsChanges    = @()
    controls           = @()
    errors             = @()
}
$errors = New-Object System.Collections.Generic.List[string]

# The guest guard, before anything touches Outlook. Its refusal is written as the reading's, so the live test can say why.
try { Assert-TestbedGuest -ExpectedUser $ExpectedUser }
catch { $reading['refused'] = "'$env:USERNAME' is not a test guest's user ($($ExpectedUser -join ', '))" }
$machineOk = $env:COMPUTERNAME -and $env:COMPUTERNAME.StartsWith($ExpectedComputerNamePrefix, [StringComparison]::OrdinalIgnoreCase)
if (-not $reading['refused'] -and -not $machineOk) {
    $reading['refused'] = "'$env:COMPUTERNAME' is not a test guest (its name must start with '$ExpectedComputerNamePrefix')"
}
if (-not $reading['refused']) {
    # An OUTLOOK.EXE in this session, by its process: GetActiveObject cannot tell, as an Outlook that has not lost
    # the focus is not in the Running Object Table (OutlookDialogUia.ps1).
    $session = (Get-Process -Id $PID).SessionId
    if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session }).Count -eq 0) {
        $reading['refused'] = "no OUTLOOK.EXE runs in this session ($session) - the runner starts it before the suite"
    }
}
if ($reading['refused']) {
    Write-Host "REFUSED: $($reading['refused'])"
    Write-Reading $reading
    exit 4
}

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$before = Get-FolderHashes
$defaultsBefore = Get-SignatureDefaultValues
$process = Start-SignaturesDialog
$dialog = Find-TopWindow $DialogName $DialogWaitSeconds
if ($dialog) {
    $reading['dialogFound'] = $true
    try {
        $reading['dialogClass'] = [string]$dialog.Current.ClassName
        $reading['settledAfterSeconds'] = Wait-DialogSettled $dialog $SettleSeconds
        $records = @(Get-UiaTree $dialog | ForEach-Object { ConvertTo-ControlRecord $_ })
        $named = @($records | Where-Object { $_.Name })
        $reading['controls'] = @($named | ForEach-Object { ConvertTo-Plain $_ })
        $reading['account'] = ConvertTo-Plain (Select-DialogControl $records $AccountBoxPattern)
        $reading['newMessages'] = ConvertTo-Plain (Select-DialogControl $records $NewMessagesPattern)
        $reading['repliesForwards'] = ConvertTo-Plain (Select-DialogControl $records $RepliesPattern)
        $delete = Select-DialogControl $records $DeleteButtonPattern 'Button'
        $reading['delete'] = ConvertTo-Plain $delete
        # The signature list and its buttons: a list a user cannot select in, and a New button that is shut, leave no
        # signature to delete, rename or make (the control of 2026-11-18 showed both open).
        $list = Select-DialogControl $records $SignatureListPattern 'List'
        $reading['signatureList'] = ConvertTo-Plain $list
        $reading['newButton'] = ConvertTo-Plain (Select-DialogControl $records $NewButtonPattern 'Button')
        $items = @($records | Where-Object { $_.Type -eq 'ListItem' -and $_.Name })
        # The list's items are not in the control view of this SDM dialog (the control of 2026-11-18): then no single
        # signature is selected, and the list's own state is the reading.
        $reading['itemsExposed'] = ($items.Count -gt 0)
        $testItems = @($items | Where-Object { ([string]$_.Name).StartsWith($TestPrefix, [StringComparison]::Ordinal) })
        $reading['otherSignatureItems'] = $items.Count - $testItems.Count
        $perSignature = New-Object System.Collections.Generic.List[object]
        foreach ($item in $testItems) {
            $entry = [ordered]@{ name = [string]$item.Name; selected = $false; deleteEnabled = $null; error = $null }
            try {
                $entry['selected'] = Select-SignatureInList $item.Element
                if ($delete) { $entry['deleteEnabled'] = [bool]$delete.Element.Current.IsEnabled }
            }
            catch { $entry['error'] = $_.Exception.Message }
            $perSignature.Add($entry)
        }
        # ToArray, not @(): Windows PowerShell 5.1 refuses @() of an empty List[object] as a dictionary's value
        # ("Argument types do not match", the first readings of 2026-11-18).
        $reading['signatures'] = $perSignature.ToArray()
        if (-not $reading['newMessages']) { $errors.Add('no New messages box was found') }
        if (-not $reading['repliesForwards']) { $errors.Add('no Replies/forwards box was found') }
        if (-not $delete) { $errors.Add('no Delete button was found') }
        if (-not $list) { $errors.Add('no signature list was found') }
    }
    catch { $errors.Add('reading the dialog failed: ' + $_.Exception.Message) }
    finally { $reading['closed'] = [bool](Close-WithCancel $dialog $DialogName) }
}
else { $errors.Add("'$DialogName' did not open within $DialogWaitSeconds s (helper pid $($process.Id))") }
if (-not $process.WaitForExit(15000)) { $errors.Add("the helper (pid $($process.Id)) is still waiting - the dialog may still be open") }

$after = Get-FolderHashes
$changes = @(Get-FolderChanges $before $after)
$reading['folderUnchanged'] = ($changes.Count -eq 0)
$reading['folderChanges'] = $changes
$defaultChanges = @(Get-DefaultsChanges $defaultsBefore (Get-SignatureDefaultValues))
$reading['defaultsUnchanged'] = ($defaultChanges.Count -eq 0)
$reading['defaultsChanges'] = $defaultChanges
$reading['errors'] = $errors.ToArray()
Write-Reading $reading

Write-Host ("SIGNATURE-DIALOG found={0} class={1} settled after {2} s closed={3} folderUnchanged={4} defaultsUnchanged={5}" -f $reading['dialogFound'], $reading['dialogClass'], $reading['settledAfterSeconds'], $reading['closed'], $reading['folderUnchanged'], $reading['defaultsUnchanged'])
foreach ($c in $defaultChanges) { Write-Host "  DEFAULTS CHANGED: $c" }
foreach ($box in @('account', 'newMessages', 'repliesForwards', 'delete', 'signatureList', 'newButton')) {
    $r = $reading[$box]
    if ($r) { Write-Host ("  {0}: '{1}' {2} enabled={3} value='{4}'" -f $box, $r.name, $r.type, $r.enabled, $r.value) } else { Write-Host "  ${box}: not found" }
}
foreach ($s in $reading['signatures']) { Write-Host ("  signature '{0}' selected={1} deleteEnabled={2}{3}" -f $s.name, $s.selected, $s.deleteEnabled, $(if ($s.error) { ' error=' + $s.error } else { '' })) }
foreach ($e in $errors) { Write-Host "  ERROR: $e" }
exit (Get-DialogExitCode ([pscustomobject]$reading))
