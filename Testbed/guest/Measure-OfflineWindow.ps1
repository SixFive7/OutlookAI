#Requires -Version 5.1
<#
.SYNOPSIS
    Q145 on the Exchange VM: where the test account's "Download email for the past" setting lives in the
    Outlook profile, which of Microsoft's eleven windows Office LTSC 2024 offers, and whether Outlook
    honours a value written into the profile while it is closed.

.DESCRIPTION
    RUN ON THE GUEST OutlookAI-Exchange, in session 1 through Testbed/guest/Register-InteractiveTask.ps1
    -RunLevel Limited (Windows PowerShell 5.1). THE TEST ACCOUNT'S SETTING ONLY, AND ONLY ON THIS VM: it
    changes the offline window of telefonie-ai-test@xxlnet.nl in the VM's own Outlook profile, which the
    maintainer allowed for this measurement on the condition that the VM's checkpoint is restored
    afterwards (2026-11-17). No mailbox item is read or written; telefonie@xxlnet.nl's setting is never
    changed - the account row is chosen by the hub's address, and the dialog's own header is checked
    against it before the slider is touched.

      -Step Export -Label <name>   Every value of every section of the profile, as hex, into
                                   measure\q145\<name>.json on the guest. Prints a count.
      -Step Diff -From <a> -To <b> Which values differ between two exports: section, value name and
                                   property tag, the old and new value - in full when 8 bytes or fewer.
      -Step Read                   Opens Account Settings > the test account > Change, reads the slider's
                                   position and label, and closes the dialog unchanged (WM_CLOSE, the
                                   dialog's Cancel) and Account Settings with its Close button.
      -Step Set -Position <1..11>  The same, but sets the slider to that position through UI Automation,
                                   reads the label Outlook shows for it, and confirms with Next and the
                                   finish page's own button. Every page's texts are printed.
      -Step WriteMonths -Months <n> With Outlook CLOSED: writes <n> into the test account's section(s)
                                   that hold property 0x6649 (PT_LONG, the value name 00036649) - only
                                   in a section that also carries the test account's address - and
                                   reads it back.
      -Step WritePair -Months <m> -Days <d> [-WhileRunning]
                                   Writes the pair - 0x6649 months and 0x665A days, as Outlook itself
                                   always does - into the same section(s). With Outlook closed, or,
                                   with -WhileRunning, only while it RUNS: the add-in's own case, since
                                   it lives inside Outlook (asked 2026-11-17).

    Microsoft documents eleven windows, always as a months/days pair (.work's q145-admin report, open
    question 2): 3 days, 1 week, 2 weeks, 1/3/6 months, 1/2/3/5 years and All. The slider was read on
    2026-11-17 as positions 1 to 11 with 7 = "1 year".

.PARAMETER SelfTest
    The pure parts: the property-tag reading of a profile value name, the diff, the window table, the
    guard on which section a write may touch. No Outlook, any machine.
#>
[CmdletBinding()]
param(
    [ValidateSet('Export', 'Diff', 'Read', 'Set', 'WriteMonths', 'WritePair')] [string] $Step,
    [string] $Label,
    [string] $From,
    [string] $To,
    [int]    $Position,
    [int]    $Months = -1,
    [int]    $Days = -1,
    [switch] $WhileRunning,
    [string] $Hub = 'telefonie-ai-test@xxlnet.nl',
    [string] $ProfileName = 'Outlook',
    [string] $Root = 'C:\OutlookAI-Q5\measure\q145',
    [int]    $DialogWaitSeconds = 25,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string] $ExpectedComputerNamePrefix = 'OAI-EXCHANGE',
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# The eleven windows, as Microsoft writes them (KB 3115009, KB 2733062, the ADMX), by slider position as
# read on this VM: label, months, days.
$Windows = @(
    @{ Position = 1;  Label = '3 days';   Months = 0;  Days = 3 },
    @{ Position = 2;  Label = '1 week';   Months = 0;  Days = 7 },
    @{ Position = 3;  Label = '2 weeks';  Months = 0;  Days = 14 },
    @{ Position = 4;  Label = '1 month';  Months = 1;  Days = 0 },
    @{ Position = 5;  Label = '3 months'; Months = 3;  Days = 0 },
    @{ Position = 6;  Label = '6 months'; Months = 6;  Days = 0 },
    @{ Position = 7;  Label = '1 year';   Months = 12; Days = 0 },
    @{ Position = 8;  Label = '2 years';  Months = 24; Days = 0 },
    @{ Position = 9;  Label = '3 years';  Months = 36; Days = 0 },
    @{ Position = 10; Label = '5 years';  Months = 60; Days = 0 },
    @{ Position = 11; Label = 'All';      Months = 0;  Days = 0 }
)

# =============================================================================================
# PURE PARTS
# =============================================================================================

# A profile value's name is its property tag, type first: '00036649' is PT_LONG (0x0003) of property 0x6649.
function ConvertFrom-ProfileValueName([string] $Name) {
    if ($Name -notmatch '^[0-9A-Fa-f]{8}$') { return $null }
    $type = [Convert]::ToInt32($Name.Substring(0, 4), 16)
    $id = [Convert]::ToInt32($Name.Substring(4, 4), 16)
    $types = @{ 0x0003 = 'PT_LONG'; 0x000B = 'PT_BOOLEAN'; 0x001E = 'PT_STRING8'; 0x001F = 'PT_UNICODE'; 0x0102 = 'PT_BINARY'; 0x0040 = 'PT_SYSTIME'; 0x0048 = 'PT_CLSID'; 0x1102 = 'PT_MV_BINARY'; 0x101F = 'PT_MV_UNICODE' }
    $typeName = $types[$type]
    if (-not $typeName) { $typeName = 'type 0x' + $type.ToString('X4') }
    return [pscustomobject]@{ Id = $id; IdHex = '0x' + $id.ToString('X4'); Type = $type; TypeName = $typeName }
}

# A value as it may be printed: a 4-byte one as the number it is, a short one as hex, a long one by length.
function Format-ProfileValue([string] $Hex) {
    if ($null -eq $Hex) { return '(absent)' }
    if ($Hex.StartsWith('STR:')) { return ('a string, sha256 {0}' -f (Get-HexHash $Hex)) }
    if ($Hex.Length -eq 8) {
        $b = for ($i = 0; $i -lt 8; $i += 2) { [Convert]::ToByte($Hex.Substring($i, 2), 16) }
        return ('{0} (0x{1})' -f [BitConverter]::ToInt32([byte[]]$b, 0), $Hex)
    }
    if ($Hex.Length -le 16) { return '0x' + $Hex }
    return ('{0} bytes, sha256 {1}' -f ($Hex.Length / 2), (Get-HexHash $Hex))
}

function Get-HexHash([string] $Hex) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $h = $sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($Hex)) } finally { $sha.Dispose() }
    return (($h[0..3] | ForEach-Object { $_.ToString('x2') }) -join '')
}

# The differences between two exports - each a hashtable "<section>|<name>" -> hex.
function Compare-ProfileExport($Before, $After) {
    $keys = @($Before.Keys) + @($After.Keys) | Sort-Object -Unique
    $out = @()
    foreach ($k in $keys) {
        $a = $Before[$k]; $b = $After[$k]
        if ($a -ne $b) { $out += [pscustomobject]@{ Key = $k; Before = $a; After = $b } }
    }
    return $out
}

# A 4-byte little-endian hex for a value.
function ConvertTo-LongHex([int] $Value) { return (([BitConverter]::GetBytes([int]$Value) | ForEach-Object { $_.ToString('X2') }) -join '') }

# The UTF-16 hex Outlook stores an address as, with its terminating NUL.
function ConvertTo-Utf16Hex([string] $Text) { return (([System.Text.Encoding]::Unicode.GetBytes($Text + [char]0) | ForEach-Object { $_.ToString('X2') }) -join '') }

# May a write touch this section? Only one that holds 0x6649 AND carries the hub's address in one of its values.
function Test-SectionIsTheHubs($SectionValues, [string] $HubAddress) {
    if (-not $SectionValues.ContainsKey('00036649')) { return $false }
    $needle = ConvertTo-Utf16Hex $HubAddress
    foreach ($v in $SectionValues.Values) { if ($v -and $v.ToUpperInvariant().Contains($needle.ToUpperInvariant())) { return $true } }
    return $false
}

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $Name, [bool] $Ok) { if ($Ok) { $script:p++; Write-Host "  PASS  $Name" } else { $script:f++; Write-Host "  FAIL  $Name" } }
    $n = ConvertFrom-ProfileValueName '00036649'
    Check '00036649 is PT_LONG 0x6649' ($n.IdHex -eq '0x6649' -and $n.TypeName -eq 'PT_LONG')
    Check '001f6641 is PT_UNICODE 0x6641' ((ConvertFrom-ProfileValueName '001f6641').TypeName -eq 'PT_UNICODE')
    Check 'a name that is no tag is not read as one' ($null -eq (ConvertFrom-ProfileValueName 'Account Name'))
    Check 'a 4-byte value prints as its number' ((Format-ProfileValue '0C000000') -eq '12 (0x0C000000)')
    Check 'and 36 as 36' ((Format-ProfileValue (ConvertTo-LongHex 36)) -eq '36 (0x24000000)')
    Check 'a long value prints by length and hash only' ((Format-ProfileValue ('AB' * 40)) -match '^40 bytes, sha256 [0-9a-f]{8}$')
    Check 'a string value prints by hash only' ((Format-ProfileValue 'STR:telefonie') -match '^a string, sha256 [0-9a-f]{8}$')
    $d = @(Compare-ProfileExport @{ 'S|00036649' = '0C000000'; 'S|x' = 'AA' } @{ 'S|00036649' = '24000000'; 'S|x' = 'AA'; 'S|y' = '01' })
    Check 'the diff finds a changed value and a new one, and nothing else' ($d.Count -eq 2 -and ($d | Where-Object { $_.Key -eq 'S|00036649' }).After -eq '24000000')
    Check 'eleven windows' ($Windows.Count -eq 11)
    Check 'position 7 is one year, 12 months' (($Windows | Where-Object { $_.Position -eq 7 }).Months -eq 12)
    Check 'positions 9 and 10 are 36 and 60 months' (($Windows | Where-Object { $_.Position -eq 9 }).Months -eq 36 -and ($Windows | Where-Object { $_.Position -eq 10 }).Months -eq 60)
    Check 'the day windows carry months 0' (@($Windows | Where-Object { $_.Days -gt 0 -and $_.Months -ne 0 }).Count -eq 0)
    $hubHex = ConvertTo-Utf16Hex 'hub@x.nl'
    Check 'a section with 0x6649 and the hub''s address may be written' (Test-SectionIsTheHubs @{ '00036649' = '0C000000'; '001f6641' = $hubHex } 'hub@x.nl')
    Check 'one with 0x6649 and another address may not' (-not (Test-SectionIsTheHubs @{ '00036649' = '0C000000'; '001f6641' = (ConvertTo-Utf16Hex 'other@x.nl') } 'hub@x.nl'))
    Check 'one with the hub''s address and no 0x6649 may not' (-not (Test-SectionIsTheHubs @{ '001f6641' = $hubHex } 'hub@x.nl'))
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $writes = @($ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] -and $x.GetCommandName() -match '^(Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|Remove-Item)$' }, $true))
    Check 'the one registry write is Set-ItemProperty, once' ($writes.Count -eq 1 -and $writes[0].GetCommandName() -eq 'Set-ItemProperty')
    $fns = @{}
    foreach ($x in $ast.FindAll({ param($y) $y -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $fns[$x.Name] = $x.Extent.Text }
    Check 'it sits in Write-HubValue' ($fns['Write-HubValue'] -and $fns['Write-HubValue'].Contains('Set-ItemProperty'))
    $callers = @($ast.FindAll({ param($y) $y -is [System.Management.Automation.Language.CommandAst] -and $y.GetCommandName() -eq 'Write-HubValue' }, $true) | ForEach-Object { $q = $_.Parent; while ($q -and -not ($q -is [System.Management.Automation.Language.FunctionDefinitionAst])) { $q = $q.Parent }; if ($q) { $q.Name } else { '<script>' } } | Sort-Object -Unique)
    Check ('whose only callers are the two write steps: ' + ($callers -join ', ')) ((@($callers | Sort-Object) -join ',') -eq 'Invoke-WriteMonths,Invoke-WritePair')
    Check 'each takes its sections from Get-HubSections, which keeps only the hub''s' ($fns['Invoke-WriteMonths'].Contains('Get-HubSections') -and $fns['Invoke-WritePair'].Contains('Get-HubSections') -and $fns['Get-HubSections'].Contains('Test-SectionIsTheHubs'))
    Check 'and each checks whether Outlook runs before it writes' ($fns['Invoke-WriteMonths'].Contains('Get-Process -Name OUTLOOK') -and $fns['Invoke-WritePair'].Contains('Get-Process -Name OUTLOOK'))
    $strings = @($ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $x.Value -match '^(Remove|Repair\.\.\.|New\.\.\.|Set as Default|More Settings)$' }, $true))
    Check 'no button but Change..., Next, the finish page''s and Close is ever looked for' ($strings.Count -eq 0)
    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:p, $script:f)
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }
if (-not $Step) { throw '-Step is mandatory.' }

function Assert-TestbedGuestLocal {
    $machineOk = $env:COMPUTERNAME -and $ExpectedComputerNamePrefix -and $env:COMPUTERNAME.StartsWith($ExpectedComputerNamePrefix, [StringComparison]::OrdinalIgnoreCase)
    if (-not ($ExpectedUser -contains $env:USERNAME) -or -not $machineOk) {
        throw "REFUSING TO RUN: '$env:USERNAME' on '$env:COMPUTERNAME' is not the Exchange test VM (allowed users: $($ExpectedUser -join ', '); computer prefix '$ExpectedComputerNamePrefix')."
    }
}
Assert-TestbedGuestLocal

$lines = New-Object System.Collections.Generic.List[string]
function Say([string] $Text) { $lines.Add($Text); Write-Host $Text }
if (-not (Test-Path -LiteralPath $Root)) { [void][System.IO.Directory]::CreateDirectory($Root) }
$resultPath = Join-Path $Root ("result-{0}-{1:yyyyMMdd-HHmmss}.txt" -f $Step, (Get-Date))
Say ("MEASURE-OFFLINE-WINDOW {0} on {1} at {2}" -f $Step, $env:COMPUTERNAME, (Get-Date).ToString('o'))

$ProfileKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$ProfileName"

function Read-Profile {
    $map = @{}
    $sections = @{}
    foreach ($sec in @(Get-ChildItem -LiteralPath $ProfileKey -ErrorAction Stop)) {
        $values = @{}
        $key = Get-Item -LiteralPath $sec.PSPath
        foreach ($name in $key.GetValueNames()) {
            $v = $key.GetValue($name, $null, 'DoNotExpandEnvironmentNames')
            $hex = $null
            if ($v -is [byte[]]) { $hex = (($v | ForEach-Object { $_.ToString('X2') }) -join '') }
            elseif ($v -is [int]) { $hex = ConvertTo-LongHex $v }
            elseif ($null -ne $v) { $hex = 'STR:' + [string]$v }
            $values[$name] = $hex
            $map[$sec.PSChildName + '|' + $name] = $hex
        }
        $sections[$sec.PSChildName] = $values
    }
    return [pscustomobject]@{ Map = $map; Sections = $sections }
}

function Invoke-Export {
    if (-not $Label -or $Label -notmatch '^[A-Za-z0-9_-]+$') { throw '-Label is mandatory: letters, digits, - and _.' }
    $p = Read-Profile
    $path = Join-Path $Root "$Label.json"
    $json = ($p.Map.GetEnumerator() | Sort-Object Key | ForEach-Object { [ordered]@{ k = $_.Key; v = $_.Value } }) | ConvertTo-Json -Depth 3 -Compress
    [System.IO.File]::WriteAllText($path, $json)
    $holders = @($p.Sections.Keys | Where-Object { $p.Sections[$_].ContainsKey('00036649') })
    Say ("EXPORTED {0}: {1} values in {2} sections; sections holding 00036649: {3}" -f $Label, $p.Map.Count, $p.Sections.Count, $(if ($holders.Count) { ($holders | ForEach-Object { "$_ = " + (Format-ProfileValue $p.Sections[$_]['00036649']) + $(if (Test-SectionIsTheHubs $p.Sections[$_] $Hub) { ' (the hub''s)' } else { '' }) }) -join '; ' } else { 'none' }))
}

function Read-Export([string] $Name) {
    $path = Join-Path $Root "$Name.json"
    $map = @{}
    # Assigned first: Windows PowerShell 5.1's ConvertFrom-Json hands a JSON array down the pipeline as ONE object.
    $data = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    foreach ($e in $data) { $map[[string]$e.k] = [string]$e.v }
    return $map
}

function Invoke-Diff {
    $d = @(Compare-ProfileExport (Read-Export $From) (Read-Export $To))
    Say ("DIFF {0} -> {1}: {2} value(s) changed" -f $From, $To, $d.Count)
    foreach ($x in $d) {
        $parts = $x.Key.Split('|'); $tag = ConvertFrom-ProfileValueName $parts[1]
        $what = if ($tag) { "$($tag.IdHex) $($tag.TypeName)" } else { 'not a tag' }
        Say ("  section {0} value {1} ({2}): {3} -> {4}" -f $parts[0], $parts[1], $what, (Format-ProfileValue $x.Before), (Format-ProfileValue $x.After))
    }
}

# ---------------------------------------------------------------------------------------------
# The dialogs: Account Settings (#32770, Win32 list and toolbar, read through MSAA, clicked where MSAA
# says they are) and the test account's "Change Account Settings" (a NetUI dialog, through UI Automation).
# ---------------------------------------------------------------------------------------------
function Initialize-Ui {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, Accessibility
    if (-not ('OaiQ145' -as [type])) {
        Add-Type -ReferencedAssemblies Accessibility -TypeDefinition @'
using System; using System.Runtime.InteropServices; using Accessibility;
public static class OaiQ145 {
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object acc);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    static IAccessible Client(IntPtr hwnd) { Guid iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71"); object o; return AccessibleObjectFromWindow(hwnd, 0xFFFFFFFC, ref iid, out o) == 0 ? (IAccessible)o : null; }
    public static string[] Names(IntPtr hwnd) {
        IAccessible a = Client(hwnd); int n = a.accChildCount; string[] r = new string[n];
        for (int i = 1; i <= n; i++) { try { r[i - 1] = a.get_accName(i); } catch { r[i - 1] = null; } }
        return r;
    }
    // Clicks the centre of child <child> of the window's client object, where MSAA says it is on screen.
    public static bool ClickChild(IntPtr hwnd, int child) {
        IAccessible a = Client(hwnd); int x, y, w, h;
        a.accLocation(out x, out y, out w, out h, child);
        if (w <= 0 || h <= 0) return false;
        SetCursorPos(x + w / 2, y + h / 2);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
        return true;
    }
}
'@
    }
}

function Find-Window([string] $Name, [string] $Class, [int] $Seconds) {
    $ae = [System.Windows.Automation.AutomationElement]
    $c = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($ae::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($ae::ClassNameProperty, $Class)))
    $until = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $until) {
        $w = $ae::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Children, $c)
        if (-not $w) { $w = $ae::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c) }
        if ($w) { return $w }
        Start-Sleep -Milliseconds 500
    }
    return $null
}

function Find-In($Parent, $Type, [string] $Name) {
    $ae = [System.Windows.Automation.AutomationElement]
    $conds = @((New-Object System.Windows.Automation.PropertyCondition($ae::ControlTypeProperty, $Type)))
    if ($Name) { $conds += (New-Object System.Windows.Automation.PropertyCondition($ae::NameProperty, $Name)) }
    $c = if ($conds.Count -gt 1) { New-Object System.Windows.Automation.AndCondition($conds) } else { $conds[0] }
    return $Parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}

function Get-ById($Parent, [string] $Id) {
    $ae = [System.Windows.Automation.AutomationElement]
    return $Parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($ae::AutomationIdProperty, $Id)))
}

function Get-Texts($Element) {
    $ae = [System.Windows.Automation.AutomationElement]
    return @($Element.FindAll([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($ae::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text))) | ForEach-Object { [string]$_.Current.Name } | Where-Object { $_ })
}

function Close-Window($Element, [string] $What) { Close-Handle ([IntPtr]$Element.Current.NativeWindowHandle) $What }

function Close-Handle([IntPtr] $h, [string] $What) {
    [void][OaiQ145]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    Start-Sleep -Seconds 2
    Say ("  {0} closed by WM_CLOSE (its Cancel): {1}" -f $What, (-not [OaiQ145]::IsWindow($h)))
}

# Opens Account Settings and the test account's Change dialog; returns both, or throws.
function Open-ChangeDialog {
    Initialize-Ui
    foreach ($stale in @((Find-Window 'Change Account Settings' 'NUIDialog' 1), (Find-Window 'Account Settings' '#32770' 1))) { if ($stale) { Close-Window $stale 'a dialog left open by an earlier run' } }
    $helper = "`$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application'); try { `$app.ActiveExplorer().CommandBars.ExecuteMso('AccountSettings') } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject(`$app) }"
    $process = Start-Process powershell.exe -ArgumentList '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', ([Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($helper))) -WindowStyle Hidden -PassThru
    $acct = Find-Window 'Account Settings' '#32770' $DialogWaitSeconds
    if (-not $acct) { throw "Account Settings did not open within $DialogWaitSeconds s" }
    [void][OaiQ145]::SetForegroundWindow([IntPtr]$acct.Current.NativeWindowHandle)
    $list = [IntPtr](Get-ById $acct '204').Current.NativeWindowHandle
    $names = @([OaiQ145]::Names($list))
    $row = [Array]::FindIndex([string[]]$names, [Predicate[string]] { param($n) $n -and $n.Trim().Equals($Hub, [StringComparison]::OrdinalIgnoreCase) }) + 1
    if ($row -lt 1) { throw "the test account has no row in Account Settings (rows: $($names.Count))" }
    [void][OaiQ145]::ClickChild($list, $row)
    Start-Sleep -Seconds 1
    $delivers = ''
    $p = $null; $edit = Get-ById $acct '1114'
    if ($edit -and $edit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref] $p)) { $delivers = [string]$p.Current.Value }
    if (-not $delivers.StartsWith($Hub + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "the selected row is not the test account's (it delivers to '$delivers') - nothing is changed" }
    Say "  Account Settings: row $row selected, which delivers to the test account"
    $toolbar = [IntPtr](Get-ById $acct '436').Current.NativeWindowHandle
    $buttons = @([OaiQ145]::Names($toolbar))
    $change = [Array]::FindIndex([string[]]$buttons, [Predicate[string]] { param($n) $n -and $n -match '^Change' }) + 1
    if ($change -lt 1) { throw 'no Change... button on the Account Settings toolbar' }
    [void][OaiQ145]::ClickChild($toolbar, $change)
    $dialog = Find-Window 'Change Account Settings' 'NUIDialog' $DialogWaitSeconds
    if (-not $dialog) { throw "the Change dialog did not open within $DialogWaitSeconds s" }
    $texts = @(Get-Texts $dialog)
    if (-not ($texts -contains $Hub)) { Close-Window $dialog 'the Change dialog'; throw "the Change dialog is not the test account's (its texts: $($texts -join ' | '))" }
    # The handles are read once, here: a closed dialog's element has none to give.
    return [pscustomobject]@{ Account = $acct; Dialog = $dialog; Helper = $process; AccountHandle = [IntPtr]$acct.Current.NativeWindowHandle; DialogHandle = [IntPtr]$dialog.Current.NativeWindowHandle }
}

function Close-AccountSettings($Open) {
    $close = $null
    try { $close = Get-ById $Open.Account '1113' } catch { }
    if ($close) { [void][OaiQ145]::PostMessage([IntPtr]$close.Current.NativeWindowHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) }
    else { [void][OaiQ145]::PostMessage($Open.AccountHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) }
    Start-Sleep -Seconds 2
    Say ("  Account Settings closed with its Close button: {0}; helper ended {1}" -f (-not [OaiQ145]::IsWindow($Open.AccountHandle)), $Open.Helper.WaitForExit(15000))
}

function Read-Slider($Dialog) {
    $slider = Find-In $Dialog ([System.Windows.Automation.ControlType]::Slider) 'Download email for the past'
    if (-not $slider) { throw 'no "Download email for the past" slider on the Change dialog' }
    $range = $slider.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern)
    $label = @(Get-Texts $Dialog | Where-Object { $_ -match '^(\d+ (day|days|week|weeks|month|months|year|years)|All)$' }) | Select-Object -First 1
    return [pscustomobject]@{ Slider = $slider; Range = $range; Value = [int]$range.Current.Value; Min = [int]$range.Current.Minimum; Max = [int]$range.Current.Maximum; Label = $label }
}

function Invoke-Read {
    $open = Open-ChangeDialog
    try {
        $s = Read-Slider $open.Dialog
        Say ("READ: slider position {0} of {1}..{2}, label '{3}'" -f $s.Value, $s.Min, $s.Max, $s.Label)
    }
    finally {
        if ([OaiQ145]::IsWindow($open.DialogHandle)) { Close-Handle $open.DialogHandle 'the Change dialog' }
        Close-AccountSettings $open
    }
}

function Invoke-Set {
    if ($Position -lt 1 -or $Position -gt 11) { throw '-Position must be 1 to 11.' }
    $open = Open-ChangeDialog
    $confirmed = $false
    try {
        $s = Read-Slider $open.Dialog
        Say ("  before: position {0} of {1}..{2}, label '{3}'" -f $s.Value, $s.Min, $s.Max, $s.Label)
        $s.Range.SetValue($Position)
        Start-Sleep -Seconds 1
        $after = Read-Slider $open.Dialog
        Say ("SET: position {0} -> label '{1}' (Microsoft's window for it: '{2}')" -f $after.Value, $after.Label, ($Windows | Where-Object { $_.Position -eq $Position }).Label)
        $next = Find-In $open.Dialog ([System.Windows.Automation.ControlType]::Button) 'Next'
        $next.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        # The pages after Next: printed, and finished with their own Done / Finish / OK.
        for ($page = 1; $page -le 4; $page++) {
            Start-Sleep -Seconds 4
            if (-not [OaiQ145]::IsWindow($open.DialogHandle)) { $confirmed = $true; break }
            Say ("  page {0}: {1}" -f $page, ((Get-Texts $open.Dialog) -join ' | '))
            $done = $null
            foreach ($n in 'Done', 'Finish', 'OK', 'Close') { if (-not $done) { $done = Find-In $open.Dialog ([System.Windows.Automation.ControlType]::Button) $n } }
            if ($done) { Say ("  pressing '{0}'" -f $done.Current.Name); $done.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
        }
        $box = Find-Window 'Microsoft Outlook' '#32770' 3
        if ($box) { Say ('  message box: ' + ((Get-Texts $box) -join ' | ')); $ok = Find-In $box ([System.Windows.Automation.ControlType]::Button) 'OK'; if ($ok) { [void][OaiQ145]::PostMessage([IntPtr]$ok.Current.NativeWindowHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) } }
    }
    finally {
        if (-not $confirmed -and [OaiQ145]::IsWindow($open.DialogHandle)) { Close-Handle $open.DialogHandle 'the Change dialog, NOT confirmed' }
        Close-AccountSettings $open
    }
    Say ("RESULT: position {0} {1}" -f $Position, $(if ($confirmed) { 'confirmed - the Change dialog closed after its last page' } else { 'NOT confirmed' }))
}

# THE ONE REGISTRY WRITE of this script: a 4-byte value into one section the caller has proven to be the
# test account's (Test-SectionIsTheHubs), read back at once.
function Write-HubValue([string] $Section, [string] $Name, [int] $Value, $Before) {
    Set-ItemProperty -LiteralPath (Join-Path $ProfileKey $Section) -Name $Name -Value ([byte[]][BitConverter]::GetBytes($Value)) -Type Binary
    $after = (Get-Item -LiteralPath (Join-Path $ProfileKey $Section)).GetValue($Name)
    Say ("WROTE section {0} value {1}: {2} -> {3}" -f $Section, $Name, (Format-ProfileValue $Before), (Format-ProfileValue ((@($after) | ForEach-Object { $_.ToString('X2') }) -join '')))
}

# The sections a write may touch: those that hold 0x6649 and carry the test account's address.
function Get-HubSections($Profile) {
    $targets = @($Profile.Sections.Keys | Where-Object { Test-SectionIsTheHubs $Profile.Sections[$_] $Hub })
    if ($targets.Count -eq 0) { throw "no section of profile '$ProfileName' holds 00036649 and the test account's address - nothing written" }
    return $targets
}

function Invoke-WriteMonths {
    if ($Months -lt 0) { throw '-Months is mandatory.' }
    if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) { throw 'Outlook is running - this write is measured with Outlook CLOSED (Testbed/host/Restart-Guest.ps1 closes it).' }
    $p = Read-Profile
    foreach ($sec in (Get-HubSections $p)) { Write-HubValue $sec '00036649' $Months $p.Sections[$sec]['00036649'] }
}

function Invoke-WritePair {
    if ($Months -lt 0 -or $Days -lt 0) { throw '-Months and -Days are both mandatory: Outlook writes them as a pair.' }
    $running = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0
    if ($WhileRunning -and -not $running) { throw '-WhileRunning measures a write into a RUNNING Outlook, and Outlook is not running.' }
    if (-not $WhileRunning -and $running) { throw 'Outlook is running - give -WhileRunning to write while it runs, or close it first.' }
    $p = Read-Profile
    Say ("  Outlook running: {0}" -f $running)
    foreach ($sec in (Get-HubSections $p)) {
        Write-HubValue $sec '00036649' $Months $p.Sections[$sec]['00036649']
        Write-HubValue $sec '0003665a' $Days $p.Sections[$sec]['0003665a']
    }
}

try {
    switch ($Step) {
        'Export' { Invoke-Export }
        'Diff' { Invoke-Diff }
        'Read' { Invoke-Read }
        'Set' { Invoke-Set }
        'WriteMonths' { Invoke-WriteMonths }
        'WritePair' { Invoke-WritePair }
    }
}
catch { Say ('FAILED: ' + $_.Exception.Message); Say ('  at ' + (($_.ScriptStackTrace -split "`n") -join ' <- ')) }
finally { [System.IO.File]::WriteAllLines($resultPath, $lines); [System.IO.File]::WriteAllLines((Join-Path $Root 'last.txt'), $lines) }
