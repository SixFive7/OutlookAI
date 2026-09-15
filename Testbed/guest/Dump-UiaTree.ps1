#Requires -Version 5.1
<#
    ============================================================================================
    DRAFT. THIS SCRIPT HAS NEVER BEEN EXECUTED AGAINST OUTLOOK.
    ============================================================================================

    It was written by an agent forbidden to launch Outlook or attach UI Automation to any Office
    process, on a workstation holding real mail. Its Add-Type lines and the types it uses were
    checked for resolution on a stock Windows 11 with no SDK; the tree walk itself has been run
    against nothing. Verified by PARSING only. Once it HAS run on a guest, replace this banner
    with what it actually did.

.SYNOPSIS
    Dumps the UI Automation tree of every visible standard dialog. Read-only. Invokes nothing.

.DESCRIPTION
    RUN ON THE GUEST, IN THE INTERACTIVE SESSION. Windows PowerShell 5.1 - no ternary, no `??`.

    WHAT THIS DECIDES. Driving Outlook's account wizard by UI Automation is the third of three
    routes to the tier profile's POP3 account (.work/pop3-account-routes.md, route C). The whole
    route rests on one question nobody has published an answer to: does the classic Add Account
    wizard expose usable, locale-invariant AutomationIds, or is it an owner-drawn DirectUI
    surface with nothing but localised Names to key on? This script answers it in two minutes,
    and the answer is a go/no-go:

        AutomationId is a non-empty number AND FrameworkId is Win32
            -> route C is viable, and THIS DUMP IS THE SPEC. Key on AutomationId + ClassName.

        AutomationId is empty OR FrameworkId is DirectUI
            -> route C is dead for a PowerShell 5.1 managed client. Keying on Name would mean
               keying on en-GB strings on a guest whose display language is en-GB and whose
               formats are nl-NL, and the managed API cannot reach LegacyIAccessible at all
               (System.Windows.Automation.LegacyIAccessiblePattern does not exist - the pattern
               is COM-only, and IUIAutomation's vtable is invisible to PowerShell's late binder).

    HOW TO USE IT.

      1. Run Testbed/guest/Set-AccountWizardClassic.ps1 -Execute first. Without it, modern
         builds open the simplified one-box wizard, which is the surface this route cannot use.
      2. Open the target dialog BY HAND. The most automatable surface is the Control Panel Mail
         applet, which needs no Outlook process and opens no store:

             C:\Windows\System32\control.exe "C:\Program Files\Microsoft Office\root\Office16\MLCFG32.CPL"

         Use SysWOW64\control.exe and the (x86) path for 32-bit Office. Get the bitness wrong
         and you get no window and no error. The guests are 64-bit (Testbed/MEDIA.md).
      3. Click through to the page you care about - "POP and IMAP Account Settings" is the one
         that matters - and leave it open.
      4. Run this script and read tree.txt.

    WHY IT IS SAFE TO RUN. It is read-only by construction: FindAll, TreeWalker, and property
    reads. No InvokePattern, no ValuePattern.SetValue, no TogglePattern, no SendKeys, no window
    is created, moved or closed, and no Outlook process is started. It does not touch a mail
    store, a profile, or the registry. If you extend it, keep that true.

    APARTMENT STATE. Run it STA, which is powershell.exe's default - no switch needed. UI
    Automation's MTA guidance applies to threads that register EVENT HANDLERS; this script
    registers none. Do not "fix" a problem here by switching to -MTA.

    BITNESS. Irrelevant. UI Automation handles 32/64-bit interop itself, so a 64-bit PowerShell
    reads a 32-bit Office's tree.

    ONE PERFORMANCE RULE, AND IT IS A CORRECTNESS RULE TOO. The search from RootElement is
    TreeScope::Children, never Descendants: a descendants search from the desktop root walks
    every element of every window on the machine and Microsoft's own guidance warns it can
    overflow the stack. Descendant walking happens only inside a dialog we have already found.

.PARAMETER ClassName
    Which top-level windows to dump. #32770 is the standard Win32 dialog class and is what a
    classic property-sheet wizard uses. Pass '*' to dump every visible top-level window, which
    is what you want if #32770 finds nothing - that result is itself the answer.

.PARAMETER ProcessName
    Optional filter, without the .exe. Use it to cut the noise once you know which process owns
    the dialog (`rundll32` for the Mail applet, `OUTLOOK` for an in-Outlook wizard).

.PARAMETER OutFile
    Where the dump goes. It also goes to stdout.

.PARAMETER MaxDepth
    Guard against a pathological tree. A classic wizard page is about six levels deep.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Dump-UiaTree.ps1

.EXAMPLE
    .\Dump-UiaTree.ps1 -ClassName '*' -ProcessName rundll32 -OutFile C:\OutlookAI-Tier\mail-applet.txt
#>
[CmdletBinding()]
param(
    [string] $ClassName   = '#32770',
    [string] $ProcessName,
    [string] $OutFile     = 'C:\OutlookAI-Tier\uia-tree.txt',
    [int]    $MaxDepth    = 12
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$AE   = [System.Windows.Automation.AutomationElement]
$Walk = [System.Windows.Automation.TreeWalker]::ControlViewWalker

$script:Out = @()

# Verdict tallies, filled during the walk. Declared here rather than inside the walker because a
# recursive function would otherwise reset them on every call.
$script:TallyWithId    = 0
$script:TallyWithoutId = 0
$script:TallyDirectUi  = 0

function Emit {
    param([string] $Text)
    $script:Out += $Text
    Write-Host $Text
}

function Get-PatternNames {
    param([System.Windows.Automation.AutomationElement] $Element)
    $names = @()
    try {
        foreach ($p in $Element.GetSupportedPatterns()) { $names += $p.ProgrammaticName }
    }
    catch {
        return '<unreadable>'
    }
    if ($names.Count -eq 0) { return '<none>' }
    return ($names -join ',')
}

function Write-UiaTree {
    param(
        [System.Windows.Automation.AutomationElement] $Element,
        [int] $Depth = 0
    )

    if ($Depth -gt $MaxDepth) {
        Emit ((' ' * ($Depth * 2)) + '... depth limit reached, stopping this branch')
        return
    }

    $c = $null
    try { $c = $Element.Current }
    catch {
        Emit ((' ' * ($Depth * 2)) + '<element vanished while being read>')
        return
    }

    # TALLY WHILE WALKING, so the script can reach a VERDICT rather than leaving a human to read
    # a two-hundred-line tree and squint. A dump nobody can interpret is a dump nobody acts on,
    # and whoever picks this up may not have read the route notes.
    # Only INTERACTIVE controls count: a pane or a static text with no AutomationId says nothing
    # about whether the wizard can be driven, and counting them would drown the signal.
    if ($c.FrameworkId -eq 'DirectUI') { $script:TallyDirectUi++ }
    if ($c.ControlType.ProgrammaticName -match 'Edit|Button|CheckBox|ComboBox|RadioButton|List') {
        if ($c.AutomationId) { $script:TallyWithId++ } else { $script:TallyWithoutId++ }
    }

    # AutomationId FIRST, because it is the only property worth keying on and the whole point of
    # the dump is to find out whether it is there. FrameworkId second: Win32 means a classic
    # dialog, DirectUI means route C is dead.
    Emit ('{0}AutomationId="{1}" Framework="{2}" Class="{3}" Type={4} Enabled={5} Offscreen={6} Name="{7}" Patterns={8}' -f `
        (' ' * ($Depth * 2)),
        $c.AutomationId,
        $c.FrameworkId,
        $c.ClassName,
        $c.ControlType.ProgrammaticName,
        $c.IsEnabled,
        $c.IsOffscreen,
        $c.Name,
        (Get-PatternNames -Element $Element))

    $child = $null
    try { $child = $Walk.GetFirstChild($Element) } catch { $child = $null }
    while ($null -ne $child) {
        Write-UiaTree -Element $child -Depth ($Depth + 1)
        try { $child = $Walk.GetNextSibling($child) } catch { $child = $null }
    }
}

Emit "Dump-UiaTree.ps1 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Emit ("apartment={0}  process bitness={1}-bit" -f `
    [System.Threading.Thread]::CurrentThread.GetApartmentState(),
    $(if ([Environment]::Is64BitProcess) { 64 } else { 32 }))
Emit ''

$visible = New-Object System.Windows.Automation.PropertyCondition($AE::IsOffscreenProperty, $false)

if ($ClassName -eq '*') {
    $condition = $visible
}
else {
    $byClass = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, $ClassName)
    # AndCondition's constructor takes `params Condition[]`. Hand it ONE argument that is already
    # a typed array: New-Object's argument splatting and a params array disagree often enough
    # that the explicit form is the one worth writing down.
    $both = [System.Windows.Automation.Condition[]] @($byClass, $visible)
    $condition = New-Object System.Windows.Automation.AndCondition($both)
}

# Children of the desktop root ONLY. Never Descendants from here.
$windows = $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)

$matched = 0
foreach ($w in $windows) {
    $cur = $null
    try { $cur = $w.Current } catch { continue }

    if ($ProcessName) {
        $proc = $null
        try { $proc = Get-Process -Id $cur.ProcessId -ErrorAction Stop } catch { $proc = $null }
        if ($null -eq $proc) { continue }
        if ($proc.ProcessName -ne $ProcessName) { continue }
    }

    $matched++
    Emit ''
    Emit ('=== pid={0} hwnd=0x{1:X} class="{2}" name="{3}" ===' -f `
        $cur.ProcessId, $cur.NativeWindowHandle, $cur.ClassName, $cur.Name)
    Write-UiaTree -Element $w
}

Emit ''
if ($matched -eq 0) {
    # This is a RESULT, not an error to swallow. A wizard page that is open on screen and
    # produces no match here is either not the class asked for or not in the tree at all, and
    # those are different answers - re-run with -ClassName '*' to tell them apart.
    Emit ("NOTHING MATCHED. ClassName='{0}'{1}." -f $ClassName, $(if ($ProcessName) { ", ProcessName='$ProcessName'" } else { '' }))
    Emit 'If the dialog IS on screen, re-run with -ClassName ''*'' and no -ProcessName.'
    Emit 'If THAT also finds nothing, the surface is not in the UI Automation tree and route C is dead.'
}
else {
    Emit ("{0} window(s) dumped." -f $matched)
    Emit ''
    Emit '=============================== VERDICT ==============================='
    Emit ("  interactive controls WITH an AutomationId : {0}" -f $script:TallyWithId)
    Emit ("  interactive controls WITHOUT one          : {0}" -f $script:TallyWithoutId)
    Emit ("  elements reporting FrameworkId=DirectUI   : {0}" -f $script:TallyDirectUi)
    Emit ''
    if ($script:TallyWithId -gt 0 -and $script:TallyDirectUi -eq 0 -and $script:TallyWithoutId -eq 0) {
        Emit '  VIABLE - a classic Win32 property sheet. Key on AutomationId + ClassName, and'
        Emit '  THIS DUMP IS THE SPEC: the ids above are the only thing a driver was missing.'
    }
    elseif ($script:TallyWithId -gt 0) {
        Emit '  MIXED - some interactive controls are addressable and some are not. Usable ONLY if'
        Emit '  the specific controls the wizard sequence needs are in the addressable set. Check'
        Emit '  them one at a time against the page-by-page sequence; do not assume from this total.'
    }
    else {
        Emit '  DEAD for a PowerShell 5.1 managed client. Nothing stable to key on: Name is en-GB'
        Emit '  text on this guest, and LegacyIAccessible is unreachable from the managed API.'
        Emit '  Do not build a driver on this - fall back to the PRF route, or to one manual'
        Emit '  account creation captured in a checkpoint.'
    }
    Emit '======================================================================='
}

$dir = Split-Path -Parent $OutFile
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
Set-Content -LiteralPath $OutFile -Value $script:Out -Encoding UTF8
Write-Host ''
Write-Host "Written: $OutFile"

if ($matched -eq 0) { exit 1 }
exit 0
