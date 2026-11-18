#Requires -Version 5.1
<#
.SYNOPSIS
    UI Automation over Outlook's own dialogs, for the guest probes that read them - one copy (Q154).
    Dot-sourced; runs nothing when it is.

.DESCRIPTION
    Testbed/guest/Measure-SignatureRoaming.ps1 (R1, on the Exchange VM) and
    Testbed/guest/Measure-SignatureDialog.ps1 (under the live runner's -SignatureSettings step, on the
    Outlook guests) both open Outlook's Signatures and Stationery dialog and read it. This is the code
    they share, moved here from the first on 2026-11-18:

      Start-SignaturesDialog  opens the dialog by Outlook's own command,
                              ExecuteMso('SignaturesStationeryDialog') on the active Explorer, from a
                              helper process of its own: the command returns only when the dialog
                              closes, so the caller must be free to read and close it meanwhile.
      Start-Helper            that helper process, hidden, for any such command.
      Find-TopWindow          a window by its name, top-level or owned.
      Find-Named              a control by its type and name.
      Get-UiaTree             every element under one, in the order a reader meets them.
      Close-WithCancel        closes a dialog as Cancel does - the one way a probe closes anything.
      Invoke-Button           presses a button: called by Close-WithCancel, for Cancel, and by
                              nothing else (Measure-SignatureDialog.ps1 -SelfTest proves it).
      Get-SignatureFiles      the signature folder Outlook records, and its files.

    Windows PowerShell 5.1. The UI Automation assemblies are the caller's to load (Add-Type
    -AssemblyName UIAutomationClient, UIAutomationTypes) before it calls anything here but
    Start-Helper and Get-SignatureFiles; a type named in a parameter is only resolved when the function
    runs, so a self-test may dot-source this file on a machine that never loads them. ASCII only.
#>

# A helper process of its own for each command that opens a modal dialog: the command returns only when
# the dialog closes, so the probe must be free to read and Cancel it meanwhile.
function Start-Helper([string] $Script) {
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script))
    return Start-Process -FilePath powershell.exe -ArgumentList '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', $encoded -WindowStyle Hidden -PassThru
}

# The command that opens the Signatures dialog. GetLabelMso('SignaturesStationeryDialog') reads
# 'Signatures...' and enabled from the explorer (measured 2026-11-17 on the Exchange VM).
$OutlookSignaturesDialogHelper = @'
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
try { $app.ActiveExplorer().CommandBars.ExecuteMso('SignaturesStationeryDialog') } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
'@

# Opens Outlook's Signatures and Stationery dialog by its own command; returns the helper process.
function Start-SignaturesDialog {
    return Start-Helper $OutlookSignaturesDialogHelper
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

# Every element under $Element, itself first, depth first in the order a reader meets them, through the
# control view - at most $Max of them.
function Get-UiaTree($Element, [int] $Max = 3000) {
    $found = New-Object System.Collections.Generic.List[object]
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $stack = New-Object System.Collections.Stack
    $stack.Push($Element)
    while ($stack.Count -gt 0 -and $found.Count -lt $Max) {
        $e = $stack.Pop()
        $found.Add($e)
        $child = $null
        try { $child = $walker.GetFirstChild($e) } catch { }
        $kids = @()
        while ($child) { $kids += $child; try { $child = $walker.GetNextSibling($child) } catch { $child = $null } }
        for ($i = $kids.Count - 1; $i -ge 0; $i--) { $stack.Push($kids[$i]) }
    }
    return $found
}

function Invoke-Button($Button) {
    $Button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
}

# Closes a dialog as Cancel does - the one way a probe ever closes anything: its own Cancel button, or,
# where UI Automation shows none (the Signatures dialog, class bosa_sdm_Mso96, measured 2026-11-17),
# WM_CLOSE to that one window, which a dialog treats as Cancel. Says how in the caller's report when it
# keeps one (Write-ReportLine and $ReportLines, ScriptHelpers.ps1), else on the host; returns whether the
# window is gone.
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
    $message = "  {0}: no Cancel button through UI Automation; WM_CLOSE to the dialog's own window - closed {1}" -f $What, $closed
    if ((Get-Command -Name Write-ReportLine -CommandType Function -ErrorAction SilentlyContinue) -and (Get-Variable -Name ReportLines -ErrorAction SilentlyContinue)) { Write-ReportLine $message } else { Write-Host $message }
    return $closed
}

# The signature folder Outlook records (Common\General\Signatures, relative to %APPDATA%\Microsoft), and
# every file in it.
function Get-SignatureFiles {
    $folderName = 'Signatures'
    try { $v = (Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Office\16.0\Common\General' -Name 'Signatures' -ErrorAction Stop).Signatures; if ($v) { $folderName = $v } } catch { }
    $folder = Join-Path $env:APPDATA ('Microsoft\' + $folderName)
    $files = @()
    if (Test-Path -LiteralPath $folder) { $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue) }
    return [pscustomobject]@{ Folder = $folder; FolderName = $folderName; Files = $files }
}
