#Requires -Version 5.1
<#
.SYNOPSIS
    READ-ONLY measurement of how Outlook keeps signatures on the Exchange VM (Q147, measurements R0 to
    R2 of .work's q147-signatures report, section 7): which Office this is, whether it roams signatures,
    and whether the test account's offline store holds the hidden folder roaming signatures live in.

.DESCRIPTION
    RUN ON THE GUEST OutlookAI-Exchange, in session 1 through Testbed/guest/Register-InteractiveTask.ps1
    -RunLevel Limited, with Outlook running on the profile 'Outlook'. Windows PowerShell 5.1. It writes
    nothing to Outlook, the registry or a mailbox: it reads, and for R1 it opens two of Outlook's own
    dialogs and closes each with its Cancel button.

      -Step R0  Office's build, channel and product, and its licence state; every signature setting of
                section 2 of that report - the per-account defaults ("New Signature", "Reply-Forward
                Signature") in every profile, MailSettings, the policies, the roaming switches, the
                folder name - and the signature folder's files (name, size, SHA-256), with any name that
                carries the bystander's address hashed and any "<name> (<address>)" file flagged: the
                local copy classic Outlook keeps of a roaming signature.
      -Step R1  Does this Outlook roam signatures? Outlook's Options dialog is opened with
                ExecuteMso('ApplicationOptionsDialog') from a hidden helper process, its Mail page's
                "Signatures..." opened, and the names on that dialog read through UI Automation - an
                account selector or "Signatures on this device" is the roaming dialog. Cancel. Then the
                General page is read for "Store my Outlook settings in the cloud". Cancel. Never OK,
                never a checkbox, never a field.
      -Step R2  The test account's store only, through the object model: its non-IPM root, whether
                ApplicationDataRoot\49499048-0129-47f5-b95e-f9d315b861a6 is in the offline copy, how many
                subfolders it has, and each one's hidden items' message classes and which of a fixed set
                of property tags they carry - presence only, no value.

    EVERYTHING it prints is content-free by construction: names of controls and settings, counts,
    sizes, hashes, message classes and property tags. A signature name or a subfolder name is printed
    only when it is the test prefix's (OutlookAI-McpTest-) or GUID-shaped; any other is an 8-hex hash.

    Written 2026-11-17 for the maintainer's queue (Q147 R0-R2), at his direction that every measurement
    on this VM goes through a committed, self-tested probe, test account only.

.PARAMETER Step
    R0, R1 or R2. One per run.

.PARAMETER Hub
    The test account's store - the only store R2 reads. Default telefonie-ai-test@xxlnet.nl.

.PARAMETER Bystander
    The address whose appearance in a file or control name is hashed before it is printed.

.PARAMETER OutPath
    Where the result is written (the task's own output is short). Default C:\OutlookAI-Q5\measure\signatures-<step>.txt.

.PARAMETER SelfTest
    The pure parts and the read-only proof over this file's own syntax tree. No Outlook, any machine.
#>
[CmdletBinding()]
param(
    [ValidateSet('R0', 'R1', 'R2')] [string] $Step,
    [string] $Hub = 'telefonie-ai-test@xxlnet.nl',
    [string] $Bystander = 'telefonie@xxlnet.nl',
    [string] $OutPath,
    [int]    $DialogWaitSeconds = 20,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string] $ExpectedComputerNamePrefix = 'OAI-EXCHANGE',
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
$SignaturePrefix = 'OutlookAI-McpTest-'
$RoamingFolderGuid = '49499048-0129-47f5-b95e-f9d315b861a6'

# The property tags R2 asks each hidden item about - presence only. The four PR_ROAMING_* tags are the
# documented roaming-settings properties ([MS-OXOCFG]); the rest frame what kind of item it is.
$ProbeTags = [ordered]@{
    'PR_MESSAGE_CLASS'         = '0x001A001F'
    'PR_SUBJECT'               = '0x0037001F'
    'PR_ROAMING_DATATYPES'     = '0x7C060003'
    'PR_ROAMING_DICTIONARY'    = '0x7C070102'
    'PR_ROAMING_XMLSTREAM'     = '0x7C080102'
    'PR_ROAMING_BINARYSTREAM'  = '0x7C090102'
    'PR_CREATION_TIME'         = '0x30070040'
    'PR_LAST_MODIFICATION_TIME' = '0x30080040'
    'PR_MESSAGE_SIZE'          = '0x0E080003'
    'PR_HASATTACH'             = '0x0E1B000B'
}

# =============================================================================================
# PURE PARTS
# =============================================================================================

function Get-ShortHash([string] $Text) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $h = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes([string]$Text)) } finally { $sha.Dispose() }
    return (($h[0..3] | ForEach-Object { $_.ToString('x2') }) -join '')
}

# A name as it may be printed: the test prefix's and GUID-shaped ones as they are, any other hashed.
function Get-PrintableName([string] $Name, [string] $BystanderAddress) {
    if ([string]::IsNullOrEmpty($Name)) { return '(empty)' }
    if ($BystanderAddress -and $Name.IndexOf($BystanderAddress, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return 'name#' + (Get-ShortHash $Name) + ' (carries the bystander''s address)' }
    if ($Name.StartsWith($SignaturePrefix, [StringComparison]::Ordinal)) { return $Name }
    if ($Name -match '^\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?$') { return $Name }
    return 'name#' + (Get-ShortHash $Name)
}

# The local copy classic Outlook keeps of a roaming signature is named '<name> (<address>)' (report
# section 3). The address part, or null.
function Get-RoamingCopyAddress([string] $BaseName) {
    $m = [regex]::Match([string]$BaseName, '\(([^()\s]+@[^()\s]+)\)\s*$')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

# A REG_BINARY signature name as Outlook writes it - UTF-16, NUL-terminated - or a REG_SZ.
function ConvertFrom-SignatureValue($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [byte[]]) { return ([System.Text.Encoding]::Unicode.GetString([byte[]]$Value)).TrimEnd([char]0) }
    return ([string]$Value).TrimEnd([char]0)
}

# R1's verdict from the names on the Signatures dialog and the General page.
function Get-RoamingVerdict([string[]] $SignatureDialogNames, [string[]] $GeneralPageNames, [string[]] $RoamingCopies) {
    $names = @($SignatureDialogNames | Where-Object { $_ })
    $onDevice = @($names | Where-Object { $_ -match '(?i)signatures on this device' }).Count -gt 0
    $selector = @($names | Where-Object { $_ -match '(?i)^(select account|account|edit signatures for)\b' }).Count -gt 0
    $cloudSetting = @($GeneralPageNames | Where-Object { $_ -match '(?i)store my outlook settings in the cloud' }).Count -gt 0
    $copies = @($RoamingCopies | Where-Object { $_ }).Count
    $roams = $onDevice -or $selector -or $copies -gt 0
    $why = @()
    if ($onDevice) { $why += '"Signatures on this device" on the Signatures dialog' }
    if ($selector) { $why += 'an account selector on the Signatures dialog' }
    if ($copies -gt 0) { $why += "$copies '<name> (<address>)' signature file(s)" }
    if (-not $roams) { $why += 'no account selector, no "Signatures on this device", no "(address)" file' }
    return [pscustomobject]@{ Roams = $roams; OnDevice = $onDevice; AccountSelector = $selector; CloudSettingOffered = $cloudSetting; RoamingCopies = $copies; Why = ($why -join '; ') }
}

# =============================================================================================
# SELF-TEST
# =============================================================================================

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $Name, [bool] $Ok) { if ($Ok) { $script:p++; Write-Host "  PASS  $Name" } else { $script:f++; Write-Host "  FAIL  $Name" } }

    Check 'a test-prefixed signature name prints as it is' ((Get-PrintableName 'OutlookAI-McpTest-Roam1117' 'b@x.nl') -eq 'OutlookAI-McpTest-Roam1117')
    Check 'a GUID-shaped folder name prints as it is' ((Get-PrintableName '49499048-0129-47f5-b95e-f9d315b861a6' 'b@x.nl') -eq '49499048-0129-47f5-b95e-f9d315b861a6')
    Check 'any other name is hashed' ((Get-PrintableName 'Jori work' 'b@x.nl') -match '^name#[0-9a-f]{8}$')
    Check 'a name carrying the bystander''s address is hashed and says so' ((Get-PrintableName 'Sig (b@x.nl)' 'b@x.nl') -match '^name#[0-9a-f]{8} \(carries the bystander')
    Check 'even a test-prefixed one, when it carries the bystander''s address' ((Get-PrintableName 'OutlookAI-McpTest-X (b@x.nl)' 'b@x.nl') -match '^name#')
    Check 'the roaming copy''s address is read off "<name> (<address>)"' ((Get-RoamingCopyAddress 'Work (a@b.nl)') -eq 'a@b.nl')
    Check 'a plain name has none' ($null -eq (Get-RoamingCopyAddress 'Work'))
    Check 'nor a bracketed word that is no address' ($null -eq (Get-RoamingCopyAddress 'Work (old)'))
    $utf16 = [System.Text.Encoding]::Unicode.GetBytes('Work' + [char]0)
    Check 'a REG_BINARY UTF-16 name decodes without its NUL' ((ConvertFrom-SignatureValue $utf16) -eq 'Work')
    Check 'a REG_SZ name passes through' ((ConvertFrom-SignatureValue 'Work') -eq 'Work')
    $v = Get-RoamingVerdict @('Signatures and Stationery', 'E-mail Signature', 'Select signature to edit', 'Cancel') @('General') @()
    Check 'the classic dialog with no copies reads as NOT roaming' (-not $v.Roams)
    $v = Get-RoamingVerdict @('Signatures and Stationery', 'Signatures on this device') @() @()
    Check '"Signatures on this device" reads as roaming' ($v.Roams -and $v.OnDevice)
    $v = Get-RoamingVerdict @('Select account') @() @()
    Check 'an account selector reads as roaming' ($v.Roams -and $v.AccountSelector)
    $v = Get-RoamingVerdict @() @() @('a@b.nl')
    Check 'a "(address)" file alone reads as roaming' ($v.Roams -and $v.RoamingCopies -eq 1)
    $v = Get-RoamingVerdict @() @('Store my Outlook settings in the cloud') @()
    Check 'the cloud-settings checkbox is reported, apart from the verdict' ($v.CloudSettingOffered -and -not $v.Roams)

    # Read-only, read off this file's own syntax tree so this text cannot count as what it forbids.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $commands = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
    $writers = @($commands | Where-Object { $_ -match '^(Set-ItemProperty|New-ItemProperty|Remove-ItemProperty|Remove-Item|Rename-Item|Move-Item|Clear-ItemProperty|Set-Content|Add-Content|Out-File)$' })
    # The one file it writes is its own result, through [IO.File]::WriteAllLines in Write-Result.
    Check ('no registry, file or item cmdlet that writes: ' + ($writers -join ', ')) ($writers.Count -eq 0)
    $members = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true) | ForEach-Object { $_.Member.Extent.Text.Trim("'", '"') })
    $mutating = @($members | Where-Object { $_ -match '^(Delete|Save|Move|Send|SetProperty|SetProperties|DeleteProperties|Add|Remove|Copy|Close|Quit|Display|SaveAs|SetValue|Toggle|Select)$' })
    $allowed = @('Add', 'Select')   # Columns.Add on a table, List.Add; SelectionItemPattern.Select on Options' own navigation list
    $bad = @($mutating | Where-Object { $allowed -notcontains $_ })
    Check ('no late-bound member that changes an item, a store or a setting: ' + ($bad -join ', ')) ($bad.Count -eq 0)
    $invokes = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^(OK|Apply|Save|Delete|New)$' }, $true))
    Check 'no button named OK, Apply, Save, Delete or New is ever looked for' ($invokes.Count -eq 0)
    Check 'the Signatures dialog and Options are closed with Cancel' ((Get-Content -LiteralPath $PSCommandPath -Raw) -match "Close-WithCancel")

    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:p, $script:f)
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }
if (-not $Step) { throw '-Step is mandatory: R0, R1 or R2.' }

# The guest guard: the autologon account on the Exchange VM, both.
function Assert-TestbedGuestLocal {
    $machineOk = $env:COMPUTERNAME -and $ExpectedComputerNamePrefix -and $env:COMPUTERNAME.StartsWith($ExpectedComputerNamePrefix, [StringComparison]::OrdinalIgnoreCase)
    if (-not ($ExpectedUser -contains $env:USERNAME) -or -not $machineOk) {
        throw "REFUSING TO RUN: '$env:USERNAME' on '$env:COMPUTERNAME' is not the Exchange test VM (allowed users: $($ExpectedUser -join ', '); computer prefix '$ExpectedComputerNamePrefix')."
    }
}
Assert-TestbedGuestLocal

if (-not $OutPath) { $OutPath = "C:\OutlookAI-Q5\measure\signatures-$Step.txt" }
$lines = New-Object System.Collections.Generic.List[string]
function Say([string] $Text) { $lines.Add($Text); Write-Host $Text }
function Write-Result {
    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    [System.IO.File]::WriteAllLines($OutPath, $lines)
}
Say ("MEASURE-SIGNATURE-ROAMING {0} on {1} at {2}" -f $Step, $env:COMPUTERNAME, (Get-Date).ToString('o'))

function Get-Outlook {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
    catch { throw 'Outlook is not running in this session - start it first (Testbed/guest/Start-OutlookUnelevated.ps1 -Profile Outlook).' }
}

# ---------------------------------------------------------------------------------------------
# R0
# ---------------------------------------------------------------------------------------------
function Get-SignatureFiles {
    $folderName = 'Signatures'
    try { $v = (Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Office\16.0\Common\General' -Name 'Signatures' -ErrorAction Stop).Signatures; if ($v) { $folderName = $v } } catch { }
    $folder = Join-Path $env:APPDATA ('Microsoft\' + $folderName)
    $files = @()
    if (Test-Path -LiteralPath $folder) { $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue) }
    return [pscustomobject]@{ Folder = $folder; FolderName = $folderName; Files = $files }
}

function Invoke-R0 {
    Say '== R0: Office =='
    $c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path -LiteralPath $c2r) {
        $cfg = Get-ItemProperty -LiteralPath $c2r
        foreach ($n in 'VersionToReport', 'ClientVersionToReport', 'ProductReleaseIds', 'Platform', 'UpdateChannel', 'CDNBaseUrl', 'AudienceId', 'AudienceData', 'UpdatesEnabled') {
            $val = $cfg.$n
            if ($null -ne $val) { Say ("  {0} = {1}" -f $n, $val) }
        }
    }
    else { Say '  no ClickToRun configuration key' }
    $exe = 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'
    if (Test-Path -LiteralPath $exe) { Say ('  OUTLOOK.EXE file version ' + (Get-Item -LiteralPath $exe).VersionInfo.FileVersion) }
    try {
        $products = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey IS NOT NULL" -ErrorAction Stop)
        foreach ($pr in $products) {
            Say ("  licence: {0}; LicenseStatus {1}; GracePeriodRemaining {2} min; reason 0x{3:X8}; channel {4}" -f $pr.Name, $pr.LicenseStatus, $pr.GracePeriodRemaining, [uint32]$pr.LicenseStatusReason, $pr.ProductKeyChannel)
        }
        if ($products.Count -eq 0) { Say '  licence: no Office product with a key in SoftwareLicensingProduct' }
    }
    catch { Say ('  licence: could not be read - ' + $_.Exception.Message) }

    Say '== R0: signature settings (report section 2) =='
    $profiles = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'
    foreach ($prof in @(Get-ChildItem -LiteralPath $profiles -ErrorAction SilentlyContinue)) {
        $accountsKey = Join-Path $prof.PSPath '9375CFF0413111d3B88A00104B2A6676'
        foreach ($acc in @(Get-ChildItem -LiteralPath $accountsKey -ErrorAction SilentlyContinue)) {
            $props = Get-ItemProperty -LiteralPath $acc.PSPath
            $label = ConvertFrom-SignatureValue $props.'Account Name'
            if (-not $label) { $label = ConvertFrom-SignatureValue $props.'Display Name' }
            $new = ConvertFrom-SignatureValue $props.'New Signature'
            $reply = ConvertFrom-SignatureValue $props.'Reply-Forward Signature'
            $shownLabel = $label
            if ($label -and $Bystander -and $label.IndexOf($Bystander, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $shownLabel = '<bystander>' }
            Say ("  profile '{0}' account {1} '{2}': New Signature {3}; Reply-Forward Signature {4}" -f $prof.PSChildName, $acc.PSChildName, $shownLabel,
                $(if ($null -ne $new) { "'" + (Get-PrintableName $new $Bystander) + "'" } else { '(absent)' }),
                $(if ($null -ne $reply) { "'" + (Get-PrintableName $reply $Bystander) + "'" } else { '(absent)' }))
        }
    }
    $values = @(
        @('HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings', 'NewSignature'),
        @('HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings', 'ReplySignature'),
        @('HKCU:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings', 'NewSignature'),
        @('HKCU:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings', 'ReplySignature'),
        @('HKCU:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings', 'DisableSignatures'),
        @('HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup', 'DisableRoamingSignatures'),
        @('HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup', 'DisableRoamingSignaturesTemporaryToggle'),
        @('HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Setup', 'DisableRoamingSignatures'),
        @('HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Setup', 'DisableRoamingSignaturesTemporaryToggle'),
        @('HKCU:\Software\Microsoft\Office\16.0\Common\General', 'Signatures')
    )
    foreach ($pair in $values) {
        $val = $null
        try { $val = (Get-ItemProperty -LiteralPath $pair[0] -Name $pair[1] -ErrorAction Stop).($pair[1]) } catch { }
        $shown = '(absent)'
        if ($null -ne $val) { $shown = "'" + (Get-PrintableName (ConvertFrom-SignatureValue $val) $Bystander) + "'"; if ($val -is [int]) { $shown = [string]$val } }
        Say ("  {0}\{1} = {2}" -f ($pair[0] -replace '^HKCU:\\', 'HKCU\'), $pair[1], $shown)
    }
    $cloud = 'HKCU:\Software\Policies\Microsoft\Cloud\Office\16.0'
    Say ("  Cloud Policy key {0}: {1}" -f ($cloud -replace '^HKCU:\\', 'HKCU\'), $(if (Test-Path -LiteralPath $cloud) { 'PRESENT' } else { 'absent' }))

    Say '== R0: the signature folder =='
    $sig = Get-SignatureFiles
    Say ("  folder %APPDATA%\Microsoft\{0}: {1}" -f $sig.FolderName, $(if (Test-Path -LiteralPath $sig.Folder) { "$($sig.Files.Count) file(s)" } else { 'does not exist' }))
    foreach ($file in $sig.Files) {
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        $copy = Get-RoamingCopyAddress ([IO.Path]::GetFileNameWithoutExtension($file.Name))
        $note = ''
        if ($copy) { $note = ' ROAMING-COPY name, address ' + $(if ($Bystander -and $copy -ieq $Bystander) { '<bystander>' } else { $copy }) }
        Say ("  {0} {1} bytes sha256 {2}{3}" -f (Get-PrintableName $file.Name $Bystander), $file.Length, $hash.Substring(0, 16), $note)
    }
}

# ---------------------------------------------------------------------------------------------
# R1 - UI Automation over Outlook's own dialogs, closed with Cancel.
# ---------------------------------------------------------------------------------------------
function Get-UiaNames($Element, [int] $Max = 1500) {
    $names = New-Object System.Collections.Generic.List[string]
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $stack = New-Object System.Collections.Stack
    $stack.Push($Element)
    while ($stack.Count -gt 0 -and $names.Count -lt $Max) {
        $e = $stack.Pop()
        try {
            $type = $e.Current.ControlType.ProgrammaticName -replace '^ControlType\.', ''
            $name = [string]$e.Current.Name
            if ($type -ne 'Edit' -and $type -ne 'Document' -and $name) {
                $shown = $name
                if ($Bystander -and $name.IndexOf($Bystander, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $shown = $name.Replace($Bystander, '<bystander>') }
                $names.Add("$type`t$shown")
            }
        }
        catch { }
        $child = $null
        try { $child = $walker.GetFirstChild($e) } catch { }
        $kids = @()
        while ($child) { $kids += $child; try { $child = $walker.GetNextSibling($child) } catch { $child = $null } }
        for ($i = $kids.Count - 1; $i -ge 0; $i--) { $stack.Push($kids[$i]) }
    }
    return $names
}

function Find-TopWindow([string] $Name, [int] $Seconds) {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)
    $until = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $until) {
        $w = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
        if ($w) { return $w }
        # An owned dialog can sit one level down, under Outlook's window.
        $w = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
        if ($w -and $w.Current.ControlType -eq [System.Windows.Automation.ControlType]::Window) { return $w }
        Start-Sleep -Milliseconds 500
    }
    return $null
}

function Find-Named($Parent, [System.Windows.Automation.ControlType] $Type, [string] $Name) {
    $c = New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $Type)),
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)))
    return $Parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
}

function Invoke-Button($Button) {
    $Button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
}

# Closes a dialog as Cancel does - the one way this probe ever closes anything: its own Cancel button, or,
# where UI Automation shows none (the Signatures dialog, class bosa_sdm_Mso96, measured 2026-11-17),
# WM_CLOSE to that one window, which a dialog treats as Cancel.
function Close-WithCancel($Window, [string] $What) {
    $cancel = Find-Named $Window ([System.Windows.Automation.ControlType]::Button) 'Cancel'
    if ($cancel) { Invoke-Button $cancel; Start-Sleep -Seconds 2; return $true }
    if (-not ('OaiProbeWin' -as [type])) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class OaiProbeWin { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l); [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h); }'
    }
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    [void][OaiProbeWin]::PostMessage($handle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    Start-Sleep -Seconds 3
    $closed = -not [OaiProbeWin]::IsWindow($handle)
    Say ("  {0}: no Cancel button through UI Automation; WM_CLOSE to the dialog's own window - closed {1}" -f $What, $closed)
    return $closed
}

function Select-NavigationPage($Options, [string] $Page) {
    $item = Find-Named $Options ([System.Windows.Automation.ControlType]::ListItem) $Page
    if (-not $item) { return $false }
    $item.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
    Start-Sleep -Seconds 2
    return $true
}

# A helper process of its own for each command that opens a modal dialog: the command returns only when
# the dialog closes, so the probe must be free to read and Cancel it meanwhile.
function Start-Helper([string] $Script) {
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script))
    return Start-Process -FilePath powershell.exe -ArgumentList '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', $encoded -WindowStyle Hidden -PassThru
}

# The helper that opens the Signatures dialog by its own command. GetLabelMso('SignaturesStationeryDialog')
# reads 'Signatures...' and enabled from the explorer, ApplicationOptionsDialog disabled (measured 2026-11-17).
$SignaturesHelper = @'
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
try { $app.ActiveExplorer().CommandBars.ExecuteMso('SignaturesStationeryDialog') } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
'@

# The helper that opens Options through the File tab: the tab is invoked, then the backstage's Options.
$OptionsHelper = @'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$ae = [System.Windows.Automation.AutomationElement]
$main = $ae::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Children, (New-Object System.Windows.Automation.PropertyCondition($ae::ClassNameProperty, 'rctrl_renwnd32')))
$file = $main.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($ae::NameProperty, 'File Tab')))
if (-not $file) { $file = $main.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($ae::NameProperty, 'File'))) }
$file.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
$until = (Get-Date).AddSeconds(20); $opt = $null
while (-not $opt -and (Get-Date) -lt $until) { Start-Sleep -Milliseconds 500; $opt = $main.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($ae::NameProperty, 'Options'))) }
$p = $null
if ($opt.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref] $p)) { $p.Invoke() }
elseif ($opt.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref] $p)) { $p.Select() }
'@

function Invoke-R1 {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $null = Get-Outlook
    $signatureNames = @(); $generalNames = @()

    Say '== R1: the Signatures dialog, opened through ExecuteMso(''SignaturesStationeryDialog'') =='
    $process = Start-Helper $SignaturesHelper
    $dialog = Find-TopWindow 'Signatures and Stationery' $DialogWaitSeconds
    if ($dialog) {
        try {
            $signatureNames = @(Get-UiaNames $dialog)
            Say ("  'Signatures and Stationery' (class {0}): {1} named controls" -f $dialog.Current.ClassName, $signatureNames.Count)
            foreach ($n in $signatureNames) { Say ("    {0}" -f $n) }
        }
        finally { [void](Close-WithCancel $dialog 'Signatures and Stationery') }
    }
    else { Say "  NOT-FOUND: 'Signatures and Stationery' did not open within $DialogWaitSeconds s (helper pid $($process.Id))" }
    if (-not $process.WaitForExit(15000)) { Say "  NOTE: the helper (pid $($process.Id)) is still waiting - the dialog may still be open" }

    # Options may be switched off in this Outlook altogether: measured 2026-11-17, ApplicationOptionsDialog
    # reads disabled and the File tab shows Options greyed. Read that, and the privacy policies the testbed
    # sets (connected experiences off), which bear on any cloud feature.
    $app = Get-Outlook
    $optionsEnabled = $null
    try { $optionsEnabled = [bool]$app.ActiveExplorer().CommandBars.GetEnabledMso('ApplicationOptionsDialog') } catch { }
    Say ("  File > Options enabled: {0}" -f $(if ($null -eq $optionsEnabled) { 'could not be read' } else { $optionsEnabled }))
    $privacy = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'
    foreach ($n in 'disconnectedstate', 'usercontentdisabled', 'downloadcontentdisabled', 'controllerconnectedservicesenabled') {
        $val = $null
        try { $val = (Get-ItemProperty -LiteralPath $privacy -Name $n -ErrorAction Stop).$n } catch { }
        Say ("  privacy policy {0} = {1}" -f $n, $(if ($null -eq $val) { '(absent)' } else { $val }))
    }
    Say '== R1: Outlook''s Options, General page, through the File tab =='
    if ($optionsEnabled -eq $false) {
        Say '  NOT READ: Options is disabled in this Outlook, so "Store my Outlook settings in the cloud" cannot be looked for there'
        $copies = @((Get-SignatureFiles).Files | ForEach-Object { Get-RoamingCopyAddress ([IO.Path]::GetFileNameWithoutExtension($_.Name)) } | Where-Object { $_ })
        $v = Get-RoamingVerdict $signatureNames @() $copies
        Say ("VERDICT R1: roams={0} - {1}; 'Store my Outlook settings in the cloud' not readable (Options disabled)" -f $v.Roams, $v.Why)
        return
    }
    $process = Start-Helper $OptionsHelper
    $options = Find-TopWindow 'Outlook Options' ($DialogWaitSeconds + 20)
    if ($options) {
        try {
            Say ("  'Outlook Options' open (class {0})" -f $options.Current.ClassName)
            if (Select-NavigationPage $options 'General') {
                $generalNames = @(Get-UiaNames $options)
                $cloud = @($generalNames | Where-Object { $_ -match '(?i)settings in the cloud|roam' })
                Say ("  General page: {0} named controls; about cloud settings: {1}" -f $generalNames.Count, $(if ($cloud.Count) { $cloud -join ' | ' } else { 'none' }))
            }
            else { Say "  NOT-FOUND: no 'General' page in Options' navigation" }
        }
        finally { [void](Close-WithCancel $options 'Outlook Options') }
    }
    else { Say "  NOT-FOUND: no 'Outlook Options' window (helper pid $($process.Id))" }
    if (-not $process.WaitForExit(15000)) { Say "  NOTE: the helper (pid $($process.Id)) is still waiting" }
    # Back out of the File tab's backstage, by its own Back button.
    try {
        $main = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst([System.Windows.Automation.TreeScope]::Children, (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ClassNameProperty, 'rctrl_renwnd32')))
        $back = Find-Named $main ([System.Windows.Automation.ControlType]::Button) 'Back'
        if ($back) { Invoke-Button $back; Say '  left the File tab by its Back button' }
    }
    catch { Say ('  NOTE: the File tab could not be left: ' + $_.Exception.Message) }

    $copies = @((Get-SignatureFiles).Files | ForEach-Object { Get-RoamingCopyAddress ([IO.Path]::GetFileNameWithoutExtension($_.Name)) } | Where-Object { $_ })
    $v = Get-RoamingVerdict $signatureNames $generalNames $copies
    Say ("VERDICT R1: roams={0} - {1}; 'Store my Outlook settings in the cloud' offered={2}" -f $v.Roams, $v.Why, $v.CloudSettingOffered)
}

# ---------------------------------------------------------------------------------------------
# R2 - the hidden folder, in the test account's store only.
# ---------------------------------------------------------------------------------------------
function Get-StoreRecordKey($Store) {
    try { return [string]$Store.PropertyAccessor.BinaryToString($Store.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x0FFA0102')) } catch { return $null }
}

# One entry of the test mailbox, walked read-only: its non-IPM root, ApplicationDataRoot, the roaming folder.
# The test account's folder names are its own test data, and are printed as they are.
function Read-RoamingFolder($Ns, $Store, [string] $Label) {
    Say ("== R2: {0} - entry '{1}' (ExchangeStoreType {2}, cached {3}) ==" -f $Label, [string]$Store.DisplayName, [int]$Store.ExchangeStoreType, $Store.IsCachedExchange)
    $ipm = $Store.GetRootFolder()
    $parentHex = $null
    try { $parentHex = $ipm.PropertyAccessor.BinaryToString($ipm.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E090102')) } catch { }
    if (-not $parentHex) { return 'the top folder''s parent id could not be read - the non-IPM root is out of reach' }
    $root = $null
    try { $root = $Ns.GetFolderFromID($parentHex, $Store.StoreID) } catch { return ('the non-IPM root did not open: ' + $_.Exception.Message) }
    $top = @(); foreach ($f in $root.Folders) { $top += [string]$f.Name }
    Say ("  non-IPM root '{0}': {1} folder(s): {2}" -f [string]$root.Name, $top.Count, ($top -join ', '))
    # Where else a mailbox keeps settings: the hidden (associated) items of the non-IPM root and of the top of
    # the information store - message classes and which probed tags they carry, no value.
    Say ("  hidden items of the non-IPM root: {0}" -f (Get-HiddenCount $root))
    Write-HiddenItems $root '    '
    Say ("  hidden items of the top of the information store: {0}" -f (Get-HiddenCount $ipm))
    Write-HiddenItems $ipm '    '
    $appData = $null
    foreach ($f in $root.Folders) { if ([string]$f.Name -eq 'ApplicationDataRoot') { $appData = $f } }
    if (-not $appData) { return 'ApplicationDataRoot is NOT there' }
    $children = @(); foreach ($f in $appData.Folders) { $children += [string]$f.Name }
    Say ("  ApplicationDataRoot: {0} subfolder(s): {1}" -f $children.Count, ($children -join ', '))
    $roam = $null
    foreach ($f in $appData.Folders) { if ([string]$f.Name -ieq $RoamingFolderGuid) { $roam = $f } }
    if (-not $roam) { return "ApplicationDataRoot is there, but $RoamingFolderGuid is NOT" }
    $sigFolders = @($roam.Folders)
    Say ("  {0}: {1} subfolder(s), {2} hidden item(s) of its own" -f $RoamingFolderGuid, $sigFolders.Count, (Get-HiddenCount $roam))
    Write-HiddenItems $roam '    (the folder itself) '
    foreach ($sf in $sigFolders) {
        Say ("    subfolder {0}: container class '{1}', {2} hidden item(s), {3} item(s)" -f (Get-PrintableName ([string]$sf.Name) $Bystander), $(try { [string]$sf.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3613001F') } catch { '-' }), (Get-HiddenCount $sf), [int]$sf.Items.Count)
        Write-HiddenItems $sf '      '
    }
    return ("{0} is there, with {1} subfolder(s)" -f $RoamingFolderGuid, $sigFolders.Count)
}

function Invoke-R2 {
    $app = Get-Outlook
    $ns = $app.Session
    $store = $null
    foreach ($s in $ns.Stores) { if ([string]$s.DisplayName -ieq $Hub) { $store = $s } }
    if (-not $store) { Say "  NOT-FOUND: no store named '$Hub'"; return }
    $offline = Read-RoamingFolder $ns $store 'the offline copy (the test account''s own cached store)'
    Say ('VERDICT R2 offline copy: ' + $offline)
    # The same mailbox, read on the server: the online entry Exchange automapped (2026-11-17), found by its
    # record key - read-only, as everything here. It tells "not synced into the .ost" from "not in the mailbox".
    $key = Get-StoreRecordKey $store
    $online = $null
    foreach ($s in $ns.Stores) { if ([int]$s.ExchangeStoreType -eq 1 -and $key -and (Get-StoreRecordKey $s) -eq $key) { $online = $s } }
    if ($online) { Say ('VERDICT R2 on the server (the online entry of the same mailbox): ' + (Read-RoamingFolder $ns $online 'the server, through the online entry')) }
    else { Say 'VERDICT R2 on the server: not read - no online entry of the same mailbox in this profile' }
}

function Get-HiddenCount($Folder) {
    try { $t = $Folder.GetTable('', 1); return [int]$t.GetRowCount() } catch { return -1 }
}

function Write-HiddenItems($Folder, [string] $Indent) {
    $table = $null
    try { $table = $Folder.GetTable('', 1) } catch { Say ($Indent + 'hidden items could not be listed: ' + $_.Exception.Message); return }
    $table.Columns.RemoveAll()
    # A column the table will not take is noted and left out, so one refusal does not cost the rest.
    $columns = New-Object System.Collections.Generic.List[string]
    $refused = @()
    foreach ($name in $ProbeTags.Keys) {
        try { [void]$table.Columns.Add('http://schemas.microsoft.com/mapi/proptag/' + $ProbeTags[$name]); $columns.Add($name) }
        catch { $refused += $name }
    }
    if ($refused.Count) { Say ($Indent + 'columns the table would not take: ' + ($refused -join ', ')) }
    $n = 0
    while (-not $table.EndOfTable -and $n -lt 50) {
        $row = $table.GetNextRow()
        $values = $row.GetValues()
        $present = @(); $class = ''
        for ($i = 0; $i -lt $columns.Count; $i++) {
            $val = $values[$i]
            if ($columns[$i] -eq 'PR_MESSAGE_CLASS') { $class = [string]$val }
            elseif ($null -ne $val -and [string]$val -ne '') { $present += $columns[$i] }
        }
        Say ("{0}hidden item {1}: class '{2}'; carries {3}" -f $Indent, ($n + 1), $class, $(if ($present.Count) { $present -join ', ' } else { 'none of the probed tags' }))
        $n++
    }
}

try {
    switch ($Step) {
        'R0' { Invoke-R0 }
        'R1' { Invoke-R1 }
        'R2' { Invoke-R2 }
    }
}
catch { Say ('FAILED: ' + $_.Exception.Message) }
finally { Write-Result }
