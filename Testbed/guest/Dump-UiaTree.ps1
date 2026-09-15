<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED ON A GUEST.
    ============================================================================================

    Its body was verified by parsing, and the UIAutomation assembly loads it depends on were
    measured on a stock Windows 11 box with no SDK. Nothing here has been run against Outlook.

    WHAT IT DECIDES, IN ONE RUN. Whether Outlook's account wizard can be driven programmatically
    from PowerShell at all. That question has exactly two outcomes and this script tells them
    apart:

        AutomationId is a non-empty number AND FrameworkId is "Win32"
            -> the wizard is a classic Win32 property sheet. It can be driven deterministically,
               keyed on locale-invariant numeric control ids, and THIS DUMP IS THE SPEC.

        AutomationId is empty OR FrameworkId is "DirectUI"
            -> Office's own owner-drawn chrome. There are no stable ids, so a driver would have
               to key on localised Name strings - on a guest whose display language is en-GB -
               and the managed API cannot reach LegacyIAccessible at all (measured: the type
               does not exist in UIAutomationClient, and the COM fallback is unreachable from
               PowerShell because IUIAutomation is not IDispatch-derived). Treat as DEAD.

    IT IS READ-ONLY. It walks the tree and prints. It invokes nothing, sets no value, toggles
    nothing, opens no mail store, and starts no Outlook process. Safe to run at any point.

    HOW TO USE IT.
      1. Set the knob that restores the CLASSIC wizard, because the modern one is the bad case:
             reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\setup" `
                 /v DisableOffice365SimplifiedAccountCreation /t REG_DWORD /d 1 /f
      2. Open the Mail applet with the control.exe whose BITNESS MATCHES OFFICE - a mismatch
         silently opens nothing at all:
             64-bit Office: C:\Windows\System32\control.exe  "...\root\Office16\MLCFG32.CPL"
             32-bit Office: C:\Windows\SysWOW64\control.exe  "...\root\Office16\MLCFG32.CPL"
      3. Click through by hand to the POP and IMAP Account Settings page.
      4. Run this, in an INTERACTIVE session, and keep the output.

    WHY A HUMAN CLICKS FOR THE DUMP. This is the one step that cannot be automated before the
    dump exists, because the dump is what tells you how to automate it.

    STA IS REQUIRED and is PowerShell 5.1's default, so pass no switch. Do NOT use -MTA: MTA is
    only wanted for UIA event handlers, and this script subscribes to none.

.PARAMETER ClassName
    Window class to enumerate. #32770 is the standard Win32 dialog class, which is what a classic
    property-sheet wizard is built from.

.PARAMETER ProcessName
    Optionally narrow to one process, e.g. rundll32 for the Mail applet, or outlook.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Dump-UiaTree.ps1 > C:\uia.txt
#>

[CmdletBinding()]
param(
    [string] $ClassName = '#32770',
    [string] $ProcessName
)

$ErrorActionPreference = 'Stop'

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'This must run STA. PowerShell 5.1 is STA by default - you have passed -MTA, or are in a runspace that is not.'
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$AE   = [System.Windows.Automation.AutomationElement]
$Walk = [System.Windows.Automation.TreeWalker]::ControlViewWalker

# Verdict counters, so the script ANSWERS the question rather than leaving a human to squint at
# a tree. A dump nobody can interpret is a dump nobody acts on.
$script:win32Ids = 0
$script:emptyIds = 0
$script:directUi = 0

function Write-UiaTree {
    param([System.Windows.Automation.AutomationElement] $Element, [int] $Depth = 0)

    $c = $Element.Current
    $patterns = ($Element.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName }) -join ','

    if ($c.FrameworkId -eq 'DirectUI') { $script:directUi++ }
    if ($c.ControlType.ProgrammaticName -match 'Edit|Button|CheckBox|ComboBox|RadioButton') {
        if ($c.AutomationId) { $script:win32Ids++ } else { $script:emptyIds++ }
    }

    '{0}[{1}] AutomationId="{2}" Class="{3}" Framework="{4}" Enabled={5} Name="{6}" Patterns={7}' -f `
        (' ' * ($Depth * 2)), $c.ControlType.ProgrammaticName, $c.AutomationId, $c.ClassName,
        $c.FrameworkId, $c.IsEnabled, $c.Name, $patterns

    # ControlViewWalker, not RawViewWalker: the raw view carries MSAA noise that Microsoft's own
    # guidance criticises, and it makes the output unreadable.
    $child = $Walk.GetFirstChild($Element)
    while ($null -ne $child) {
        Write-UiaTree -Element $child -Depth ($Depth + 1)
        $child = $Walk.GetNextSibling($child)
    }
}

# Children of RootElement ONLY. Microsoft: a Descendants search from the root "may iterate
# through hundreds or even thousands of elements, possibly resulting in a stack overflow".
$conditions = @(
    (New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, $ClassName)),
    (New-Object System.Windows.Automation.PropertyCondition($AE::IsOffscreenProperty, $false))
)
$cond = New-Object System.Windows.Automation.AndCondition($conditions)
$dialogs = @($AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond))

if ($ProcessName) {
    $wanted = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $dialogs = @($dialogs | Where-Object { $wanted -contains $_.Current.ProcessId })
}

"Found $($dialogs.Count) visible '$ClassName' window(s)$(if ($ProcessName) { " in process '$ProcessName'" })."
if ($dialogs.Count -eq 0) {
    'NOTHING TO DUMP. Either the wizard is not open, or it is not this window class - try'
    '-ClassName NUIDialog (Office chrome) or drop -ProcessName. A mismatched control.exe bitness'
    'opens no window at all and looks exactly like this.'
    exit 2
}

foreach ($d in $dialogs) {
    ''
    '=== pid={0} hwnd=0x{1:X} ===' -f $d.Current.ProcessId, $d.Current.NativeWindowHandle
    Write-UiaTree -Element $d
}

''
'=============================== VERDICT ==============================='
"  interactive controls WITH an AutomationId : $script:win32Ids"
"  interactive controls WITHOUT one          : $script:emptyIds"
"  elements reporting FrameworkId=DirectUI   : $script:directUi"
if ($script:win32Ids -gt 0 -and $script:directUi -eq 0 -and $script:emptyIds -eq 0) {
    '  VIABLE - classic Win32. Key on AutomationId; this dump is the spec.'
} elseif ($script:win32Ids -gt 0) {
    '  MIXED - some controls are addressable and some are not. Usable only if the ones you need'
    '  are in the addressable set; check each control in the page-by-page sequence individually.'
} else {
    '  DEAD for a PowerShell 5.1 managed client - no stable ids to key on. Do not build on this.'
}
'======================================================================='
