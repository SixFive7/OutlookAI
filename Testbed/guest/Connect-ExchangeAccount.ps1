#Requires -Version 5.1
<#
.SYNOPSIS
    The guest half of the Exchange VM's sign-in: finds each page of Outlook's account setup and
    of the Microsoft 365 sign-in, puts the keyboard focus where the next keystrokes belong, and
    says so - so that the HOST can type the secrets through the VM's synthetic keyboard. It never
    receives, holds, reads or writes a password or a code.

.DESCRIPTION
    RUN ON THE GUEST, IN THE INTERACTIVE SESSION (Windows PowerShell 5.1), through
    Register-InteractiveTask.ps1 -RunLevel Limited - the level the Outlook it drives runs at.
    Driven step by step by Testbed/host/Invoke-ExchangeSignIn.ps1; not meant to be run by hand,
    though -Step Inspect is safe to.

    WHY THE SECRETS NEVER PASS THROUGH HERE (decided on the maintainer's behalf, 2026-10-03, under
    Q108's "decrypt them only in memory, on the host, at the moment of sign-in; pass them into the
    VM only for that sign-in; on the VM the only copy is Outlook's own token cache"). The host types
    them through Hyper-V's synthetic keyboard (Msvm_Keyboard.TypeText), which delivers keystrokes to
    whatever has the keyboard focus on the guest's console. So no guest process but the sign-in page
    itself ever holds them, and nothing in the guest - not this script, not a scheduled task's
    definition, not a job directory, not a pipe - carries them. The price is that the keystrokes go
    wherever the focus is, so this script's whole job is to make the focus certain:

      1. find the page's window and the one field the next keystrokes are for;
      2. bring that window to the foreground and focus the field;
      3. read the focus back - the field, in that window, and for a password page a field that
         reports IsPassword - and only then write READY;
      4. keep watching the focus until the host says it has typed (the .go file), and write
         LOST-FOCUS the moment it moves, so a host that has not typed yet does not type at all.
    The host types only on READY and stops on anything else.

    STEPS (-Step):
      Inspect        read-only: every top-level window, and the UI Automation tree of each one
                     that is not an Outlook explorer or inspector (those can show mail). Nothing is
                     invoked, focused or typed. Writes <StatusDir>\inspect.txt.
      Address        Outlook's "Email Account Setup" dialog: focus the email address box.
      Password       the Microsoft sign-in page: a password field (IsPassword) in focus.
      Code           the verification-code page: the code field in focus.
      Choose         invoke ONE named control on the sign-in page or an Outlook setup dialog
                     (-ControlName): a "Next", "Use a verification code", "No, sign in to this
                     app only", "Done". Invoked through its InvokePattern - no keystroke.
    Each writes <StatusDir>\<Step>.status.txt: a STATE line (READY, DONE, NOT-FOUND, LOST-FOCUS,
    TIMEOUT, REFUSED) and, below it, what has the focus - control type, AutomationId, class and the
    window's title. Never a field's value.

    CONTENT. Outlook's main window is never dumped (class rctrl_renwnd32): once the mailbox syncs it
    shows real mail. Every other line this script writes is a control's type, AutomationId, class,
    and Name - and on the sign-in pages a Name can be the account's address, which is not a secret.

.PARAMETER Step
    Inspect, Address, Password, Code or Choose.

.PARAMETER ControlName
    With -Step Choose: the Name of the control to invoke, matched exactly, case-insensitively.

.PARAMETER WaitSeconds
    How long to wait for the page to appear.

.PARAMETER HoldSeconds
    After READY: how long to keep the focus watched for the host's .go file.

.PARAMETER StatusDir
    Where the status files go. The host reads them over PowerShell Direct.
#>
[CmdletBinding()]
param(
    [ValidateSet('Inspect', 'Address', 'Password', 'Code', 'Choose')] [string] $Step = 'Inspect',
    [string] $ControlName,
    [int] $WaitSeconds = 90,
    [int] $HoldSeconds = 60,
    [string] $StatusDir = 'C:\OutlookAI-Q5\exchange\signin',
    [string[]] $ExpectedUser = @('vmadmin')
)

$ErrorActionPreference = 'Stop'

# THE GUARD, FIRST: this drives a mail account's setup, so only on a testbed guest.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
if (-not ('OutlookAI.SignInWindow' -as [type])) {
    # MSAA (oleacc.dll and the Framework's Accessibility assembly, both part of Windows) for Outlook's
    # NetUI dialogs: measured 2026-10-03, their windowless controls - the "Done" button, the "Set up
    # Outlook Mobile" check box - are absent from the managed UI Automation tree, which shows the
    # NetUIHWND pane and nothing under it, but answer through IAccessible.
    Add-Type -ReferencedAssemblies 'Accessibility' -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;
namespace OutlookAI {
    public static class SignInWindow {
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object accessible);
        [DllImport("oleacc.dll")] private static extern int AccessibleChildren(IAccessible container, int start, int count, [Out] object[] children, out int obtained);

        public sealed class Node {
            public IAccessible Owner; public int ChildId; public string Name; public int Role; public int State; public int Depth;
        }

        public static List<Node> Walk(IntPtr hwnd, int maxNodes) {
            var nodes = new List<Node>();
            Guid iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
            object root;
            if (AccessibleObjectFromWindow(hwnd, 0xFFFFFFFC, ref iid, out root) != 0 || root == null) { return nodes; }
            Visit((IAccessible)root, 0, nodes, maxNodes);
            return nodes;
        }

        private static void Visit(IAccessible acc, int depth, List<Node> nodes, int maxNodes) {
            if (depth > 30 || nodes.Count >= maxNodes) { return; }
            int count = 0;
            try { count = acc.accChildCount; } catch { return; }
            if (count <= 0) { return; }
            object[] children = new object[count];
            int got;
            if (AccessibleChildren(acc, 0, count, children, out got) != 0) { return; }
            for (int i = 0; i < got && nodes.Count < maxNodes; i++) {
                IAccessible child = children[i] as IAccessible;
                if (child != null) {
                    nodes.Add(Describe(child, 0, depth));
                    Visit(child, depth + 1, nodes, maxNodes);
                } else if (children[i] is int) {
                    nodes.Add(Describe(acc, (int)children[i], depth));
                }
            }
        }

        private static Node Describe(IAccessible acc, int childId, int depth) {
            var n = new Node { Owner = acc, ChildId = childId, Depth = depth, Name = "", Role = 0, State = 0 };
            try { n.Name = acc.get_accName(childId) ?? ""; } catch { }
            try { object r = acc.get_accRole(childId); if (r is int) { n.Role = (int)r; } } catch { }
            try { object s = acc.get_accState(childId); if (s is int) { n.State = (int)s; } } catch { }
            return n;
        }

        public static int StateOf(Node n) {
            try { object s = n.Owner.get_accState(n.ChildId); if (s is int) { return (int)s; } } catch { }
            return -1;
        }

        public static void DoDefault(Node n) { n.Owner.accDoDefaultAction(n.ChildId); }
    }
}
'@
}

# MSAA roles and states this script reads (oleacc.h).
$MsaaRolePushButton = 0x2B
$MsaaRoleCheckButton = 0x2C
$MsaaStateUnavailable = 0x1
$MsaaStateChecked = 0x10
$MsaaStateInvisible = 0x8000

# The NetUI host windows of one top-level Outlook dialog: the dialog itself and every NetUIHWND under it.
function Get-MsaaNodes($Window) {
    $handles = @([IntPtr]$Window.Handle)
    foreach ($e in (Get-Descendants $Window.Element 500)) {
        try { if ($e.Current.ClassName -eq 'NetUIHWND' -and $e.Current.NativeWindowHandle -ne 0) { $handles += [IntPtr]$e.Current.NativeWindowHandle } } catch { }
    }
    $nodes = @()
    foreach ($h in $handles) { $nodes += [OutlookAI.SignInWindow]::Walk($h, 800) }
    return $nodes
}

$A = [System.Windows.Automation.AutomationElement]
$Raw = [System.Windows.Automation.TreeWalker]::RawViewWalker
if (-not (Test-Path -LiteralPath $StatusDir)) { New-Item -ItemType Directory -Force -Path $StatusDir | Out-Null }
$statusPath = Join-Path $StatusDir ($Step + '.status.txt')
$goPath = Join-Path $StatusDir ($Step + '.go')
Remove-Item -LiteralPath $statusPath, $goPath -Force -ErrorAction SilentlyContinue

function Short([string] $s) {
    if ($null -eq $s) { return '' }
    $t = $s -replace '[\r\n]+', ' '
    if ($t.Length -gt 90) { return $t.Substring(0, 90) + '...' }
    return $t
}

function Write-Status([string] $State, [string[]] $Detail) {
    $lines = @(('STATE ' + $State), ('AT ' + (Get-Date).ToString('o')))
    if ($Detail) { $lines += $Detail }
    [IO.File]::WriteAllLines($statusPath, $lines)
    Write-Host ($lines -join [Environment]::NewLine)
}

function Get-ProcessName([int] $ProcessId) {
    try { return (Get-Process -Id $ProcessId -ErrorAction Stop).ProcessName } catch { return '?' }
}

# Top-level windows, with their process names. Outlook's explorer and inspector windows are listed
# but flagged, so nothing below ever walks into them.
function Get-TopWindows {
    $list = @()
    $w = $Raw.GetFirstChild($A::RootElement)
    while ($null -ne $w) {
        try {
            $c = $w.Current
            $list += [pscustomobject]@{
                Element = $w; Name = $c.Name; Class = $c.ClassName; ProcessId = $c.ProcessId
                Process = (Get-ProcessName $c.ProcessId); Handle = $c.NativeWindowHandle; Offscreen = $c.IsOffscreen
                MailWindow = ($c.ClassName -eq 'rctrl_renwnd32')
            }
        }
        catch { }
        $w = $Raw.GetNextSibling($w)
    }
    return $list
}

# Every element under one window in the RAW view: NetUI and web content put controls under panes
# the control view leaves out.
function Get-Descendants($Element, [int] $Max = 3000) {
    $found = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push($Element)
    while ($stack.Count -gt 0 -and $found.Count -lt $Max) {
        $e = $stack.Pop()
        $child = $Raw.GetFirstChild($e)
        while ($null -ne $child) {
            $found.Add($child)
            $stack.Push($child)
            $child = $Raw.GetNextSibling($child)
        }
    }
    return $found
}

function Describe-Focus {
    $f = $null
    try { $f = $A::FocusedElement } catch { }
    if ($null -eq $f) { return [pscustomobject]@{ Element = $null; Text = 'FOCUS (none)'; TopHandle = 0; IsPassword = $false } }
    $c = $f.Current
    $top = [OutlookAI.SignInWindow]::GetAncestor([IntPtr]$c.NativeWindowHandle, 2)
    if ($c.NativeWindowHandle -eq 0) {
        $p = $f
        while ($true) {
            $parent = $Raw.GetParent($p)
            if ($null -eq $parent -or [System.Windows.Automation.Automation]::Compare($parent, $A::RootElement)) { break }
            $p = $parent
        }
        $top = [IntPtr]$p.Current.NativeWindowHandle
    }
    $fg = [OutlookAI.SignInWindow]::GetForegroundWindow()
    $text = ('FOCUS {0} aid="{1}" class="{2}" password={3} process={4} topHandle={5} foreground={6}' -f
        $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), $c.AutomationId, $c.ClassName, $c.IsPassword,
        (Get-ProcessName $c.ProcessId), $top, $fg)
    return [pscustomobject]@{ Element = $f; Text = $text; TopHandle = [int64]$top; Foreground = [int64]$fg; IsPassword = [bool]$c.IsPassword; Aid = $c.AutomationId; Class = $c.ClassName }
}

function Bring-ToFront($Window) {
    $h = [IntPtr]$Window.Handle
    if ([OutlookAI.SignInWindow]::IsIconic($h)) { [void][OutlookAI.SignInWindow]::ShowWindow($h, 9) }
    [void][OutlookAI.SignInWindow]::SetForegroundWindow($h)
    Start-Sleep -Milliseconds 400
}

# The windows each page can be in. The sign-in page is hosted by Windows' web account manager
# (WAM): a window of the AAD broker plugin or of the application frame host that hosts it. Matched
# by process and title rather than by one class name, because the host differs between builds.
function Find-PageWindow([string] $Page) {
    $tops = @(Get-TopWindows | Where-Object { -not $_.MailWindow -and -not $_.Offscreen })
    switch ($Page) {
        'Address' { return @($tops | Where-Object { $_.Process -eq 'OUTLOOK' -and $_.Class -eq 'NUIDialog' }) }
        default {
            # Measured 2026-10-03 on OutlookAI-Exchange: the sign-in page is a WebView titled "Sign in
            # to your account" inside an ApplicationFrameWindow ("Workplace or school account"), and
            # UI Automation reports that frame under explorer's process id - so the frame CLASS is
            # what identifies it. Outlook's own setup dialogs are NUIDialog windows of OUTLOOK.
            return @($tops | Where-Object {
                    $_.Class -eq 'ApplicationFrameWindow' -or
                    $_.Process -match '^(Microsoft\.AAD\.BrokerPlugin|ApplicationFrameHost)$' -or
                    ($_.Process -eq 'OUTLOOK' -and $_.Class -eq 'NUIDialog')
                })
        }
    }
}

function Find-Field($Window, [string] $Page) {
    $all = Get-Descendants $Window.Element
    foreach ($e in $all) {
        try { $c = $e.Current } catch { continue }
        switch ($Page) {
            'Address' { if ($c.ClassName -eq 'RICHEDIT60W') { return $e } }
            'Password' { if ($c.IsPassword -and $c.IsEnabled) { return $e } }
            'Code' {
                if (-not $c.IsPassword -and $c.IsEnabled -and $c.ControlType -eq [System.Windows.Automation.ControlType]::Edit -and
                    ($c.AutomationId -match 'OTC|otc|code' -or $c.Name -match '(?i)code')) { return $e }
            }
        }
    }
    return $null
}

# ------------------------------------------------------------------------------------------------
if ($Step -eq 'Inspect') {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('INSPECT at ' + (Get-Date).ToString('o') + ' session ' + [Diagnostics.Process]::GetCurrentProcess().SessionId)
    $tops = Get-TopWindows
    foreach ($t in $tops) {
        $lines.Add(('TOP "{0}" class={1} process={2} handle={3} offscreen={4}{5}' -f (Short $t.Name), $t.Class, $t.Process, $t.Handle, $t.Offscreen, $(if ($t.MailWindow) { ' (Outlook window - not walked)' } else { '' })))
    }
    foreach ($t in @($tops | Where-Object { -not $_.MailWindow -and -not $_.Offscreen -and $_.Class -notmatch '^(Shell_TrayWnd|Shell_SecondaryTrayWnd|Progman|ConsoleWindowClass)$' })) {
        $lines.Add('')
        $lines.Add('=== ' + (Short $t.Name) + ' / ' + $t.Class + ' / ' + $t.Process)
        foreach ($e in (Get-Descendants $t.Element 1500)) {
            try {
                $c = $e.Current
                if ($c.IsOffscreen -and -not $c.HasKeyboardFocus) { continue }
                $pats = @($e.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName -replace 'PatternIdentifiers.Pattern', '' })
                $lines.Add(('  {0} name="{1}" aid="{2}" class="{3}" enabled={4}{5}{6} [{7}]' -f $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), (Short $c.Name), $c.AutomationId, $c.ClassName, $c.IsEnabled, $(if ($c.IsPassword) { ' PASSWORD' } else { '' }), $(if ($c.HasKeyboardFocus) { ' FOCUSED' } else { '' }), ($pats -join ',')))
            }
            catch { }
        }
        if ($t.Class -eq 'NUIDialog') {
            foreach ($n in (Get-MsaaNodes $t)) {
                if (-not $n.Name -and $n.Role -ne $MsaaRolePushButton -and $n.Role -ne $MsaaRoleCheckButton) { continue }
                $lines.Add(('  MSAA{0} role=0x{1:X2} state=0x{2:X} name="{3}"' -f ('  ' * $n.Depth), $n.Role, $n.State, (Short $n.Name)))
            }
        }
    }
    $lines.Add('')
    $lines.Add((Describe-Focus).Text)
    $out = Join-Path $StatusDir 'inspect.txt'
    [IO.File]::WriteAllLines($out, $lines)
    Write-Status 'DONE' @("inspect: $($lines.Count) line(s) -> $out")
    exit 0
}

# ------------------------------------------------------------------------------------------------
if ($Step -eq 'Choose') {
    if (-not $ControlName) { Write-Status 'REFUSED' @('-Step Choose needs -ControlName'); exit 4 }
    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    while ((Get-Date) -lt $deadline) {
        foreach ($w in (Find-PageWindow 'Any')) {
            # NetUI publishes its controls only while its window is active (see the focus steps).
            Bring-ToFront $w
            foreach ($e in (Get-Descendants $w.Element)) {
                try { $c = $e.Current } catch { continue }
                if (-not $c.IsEnabled) { continue }
                if (-not [string]::Equals($c.Name, $ControlName, [StringComparison]::OrdinalIgnoreCase)) { continue }

                # A check box is CLEARED, never ticked: the one this flow meets is Outlook's "Set up
                # Outlook Mobile on my phone, too", which would otherwise send the account's owner off
                # to a phone set-up page. Clearing one that is already clear does nothing.
                $toggle = $null
                try { $toggle = $e.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern) } catch { }
                if ($null -ne $toggle) {
                    $before = $toggle.Current.ToggleState
                    if ($before -eq [System.Windows.Automation.ToggleState]::On) { $toggle.Toggle() }
                    Start-Sleep -Milliseconds 500
                    $after = $toggle.Current.ToggleState
                    if ($after -ne [System.Windows.Automation.ToggleState]::Off) {
                        Write-Status 'UNEXPECTED' @("check box '$($c.Name)' reads $after after clearing")
                        exit 2
                    }
                    Write-Status 'DONE' @("CLEARED check box '$($c.Name)' (was $before) in '$(Short $w.Name)' ($($w.Process))")
                    exit 0
                }

                $invoke = $null
                try { $invoke = $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern) } catch { }
                if ($null -eq $invoke) { continue }
                # Described BEFORE the invoke: the page it belongs to is usually gone afterwards, and
                # then every property read of it returns null (measured: the first run of this step
                # invoked its button and then died describing it).
                $what = ('INVOKED {0} "{1}" aid="{2}" in "{3}" ({4})' -f $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), $c.Name, $c.AutomationId, (Short $w.Name), $w.Process)
                Bring-ToFront $w
                $invoke.Invoke()
                Start-Sleep -Seconds 2
                Write-Status 'DONE' @($what)
                exit 0
            }

            # The NetUI fallback: the same name, through MSAA.
            if ($w.Class -ne 'NUIDialog') { continue }
            foreach ($n in (Get-MsaaNodes $w)) {
                if (-not [string]::Equals($n.Name, $ControlName, [StringComparison]::OrdinalIgnoreCase)) { continue }
                if (($n.State -band ($MsaaStateUnavailable -bor $MsaaStateInvisible)) -ne 0) { continue }
                if ($n.Role -eq $MsaaRoleCheckButton) {
                    $wasChecked = (($n.State -band $MsaaStateChecked) -ne 0)
                    if ($wasChecked) { [OutlookAI.SignInWindow]::DoDefault($n) }
                    Start-Sleep -Milliseconds 700
                    $now = [OutlookAI.SignInWindow]::StateOf($n)
                    if ($now -lt 0 -or ($now -band $MsaaStateChecked) -ne 0) {
                        Write-Status 'UNEXPECTED' @("check box '$($n.Name)' still reads checked (MSAA state $now) after clearing")
                        exit 2
                    }
                    Write-Status 'DONE' @("CLEARED check box '$($n.Name)' through MSAA (was checked: $wasChecked) in '$(Short $w.Name)'")
                    exit 0
                }
                if ($n.Role -eq $MsaaRolePushButton) {
                    $what = "INVOKED button '$($n.Name)' through MSAA in '$(Short $w.Name)' ($($w.Process))"
                    [OutlookAI.SignInWindow]::DoDefault($n)
                    Start-Sleep -Seconds 2
                    Write-Status 'DONE' @($what)
                    exit 0
                }
            }
        }
        Start-Sleep -Seconds 2
    }
    Write-Status 'NOT-FOUND' @("no enabled, invokable control named '$ControlName' on any sign-in or setup window within $WaitSeconds s", (Describe-Focus).Text)
    exit 3
}

# ------------------------------------------------------------------------------------------------
# Address, Password, Code: focus, verify, hold.
$deadline = (Get-Date).AddSeconds($WaitSeconds)
$window = $null
$field = $null
while ((Get-Date) -lt $deadline -and $null -eq $field) {
    foreach ($w in (Find-PageWindow $Step)) {
        # Outlook's NetUI dialogs publish their controls to UI Automation only while they are the
        # active window (measured: the address box was absent from the tree of the same dialog left
        # in the background), so each candidate is brought to the front before it is searched.
        Bring-ToFront $w
        $f = Find-Field $w $Step
        if ($null -ne $f) { $window = $w; $field = $f; break }
    }
    if ($null -eq $field) { Start-Sleep -Seconds 2 }
}
if ($null -eq $field) {
    Write-Status 'NOT-FOUND' @("no $Step field on any sign-in or setup window within $WaitSeconds s", (Describe-Focus).Text)
    exit 3
}

Bring-ToFront $window
try { $field.SetFocus() } catch { }
Start-Sleep -Milliseconds 500

function Test-FocusOnField {
    $focus = Describe-Focus
    $ok = $false
    if ($null -ne $focus.Element) {
        $same = $false
        try { $same = [System.Windows.Automation.Automation]::Compare($focus.Element, $field) } catch { }
        $inWindow = ($focus.Foreground -eq [int64]$window.Handle) -or ($focus.TopHandle -eq [int64]$window.Handle)
        switch ($Step) {
            'Password' { $ok = $focus.IsPassword -and ($same -or $inWindow) }
            default { $ok = $same -or ($inWindow -and $focus.Class -eq $field.Current.ClassName -and $focus.Aid -eq $field.Current.AutomationId) }
        }
    }
    return [pscustomobject]@{ Ok = $ok; Focus = $focus }
}

$check = Test-FocusOnField
if (-not $check.Ok) {
    Write-Status 'LOST-FOCUS' @(('window "{0}" ({1}) handle={2}' -f (Short $window.Name), $window.Process, $window.Handle), 'the field could not be given the focus', $check.Focus.Text)
    exit 2
}
Write-Status 'READY' @(('window "{0}" ({1}) handle={2}' -f (Short $window.Name), $window.Process, $window.Handle), $check.Focus.Text)

# Hold: watch the focus until the host writes the .go file (it has typed), or time out.
$holdUntil = (Get-Date).AddSeconds($HoldSeconds)
while ((Get-Date) -lt $holdUntil) {
    if (Test-Path -LiteralPath $goPath) {
        Write-Status 'DONE' @('the host typed while the focus was held', (Describe-Focus).Text)
        exit 0
    }
    $check = Test-FocusOnField
    if (-not $check.Ok) {
        # A page that has moved on because the host already typed and pressed Enter also reads as
        # lost focus here; the host decides from the order of its own .go and this line.
        Write-Status 'LOST-FOCUS' @('the focus moved while held', $check.Focus.Text)
        exit 2
    }
    Start-Sleep -Milliseconds 250
}
Write-Status 'TIMEOUT' @("the host did not type within $HoldSeconds s", (Describe-Focus).Text)
exit 1
