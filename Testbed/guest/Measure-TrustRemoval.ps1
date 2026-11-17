#Requires -Version 5.1
<#
.SYNOPSIS
    Measures how Windows and the product's installer remove trust in the leaked signing key (Q193):
    whether removing a certificate from the user's Root store puts up a prompt, and whether the
    installer removes VSTO's trust entries for the leaked key, and only those.

.DESCRIPTION
    RUN ON AN OUTLOOK TEST GUEST, in the interactive session, NOT elevated - the token a per-user
    installer gets - through guest/Register-InteractiveTask.ps1 at -RunLevel Limited, staged beside
    it and beside guest/OutlookMapiInterop.ps1:
        .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 1800 -Script "& 'C:\OutlookAI-Q5\trust-removal\Measure-TrustRemoval.ps1' -Scenario RootStorePrompt -Execute"
    Then restore the guest's checkpoint: everything either scenario changes is the guest's own, and
    the checkpoint takes it all back. Docs/live-tier-on-the-vm.md section 4.8 is the procedure and
    the record.

    WHY (Q193 (b), the maintainer's, 2026-11-17). The certificate that signed v1.0.0.1 to v2.3.1.111
    leaked with its private key. Installer.iss removes it by thumbprint from every trust store,
    the user's Root store last, only when it is there, and without waiting - because Windows guards
    that store, and ADDING a certificate to it is known to put up a confirmation (Installer.iss,
    InstallCertificate). Whether REMOVING one does was not documented and not measured.

    SCENARIOS (-Scenario):

      RootStorePrompt   A throwaway self-signed certificate, made here with no store involved
                        (CertificateRequest), put into this user's Root store and taken out again
                        with certutil, started exactly as Inno Setup's Exec starts it for
                        Installer.iss: the 32-bit certutil a 32-bit setup reaches, STARTUPINFO
                        with STARTF_USESHOWWINDOW and SW_HIDE, a console of its own, no handles
                        inherited, CREATE_DEFAULT_ERROR_MODE. Each case watches every top-level
                        window on the desktop - hidden ones too - for one that was not there
                        before, records it (owner, class, title, visible, foreground, every child
                        control's text, and the UI Automation tree when it has one), answers it
                        through UI Automation, or with the dialog's own WM_COMMAND when UI
                        Automation cannot reach it, and records certutil's exit code and whether
                        the certificate is still in the store. The cases:
                          store-absent      certutil -user -store Root <thumbprint>, not there
                          add-no, add-yes   certutil -user -addstore Root <file>, answered No, Yes
                          store-present     the installer's own check, there
                          delete-no         certutil -user -delstore Root <thumbprint>, as
                                            Installer.iss runs it, answered No
                          delete-unanswered the same, left for -HoldSeconds, then answered No
                          delete-visible    the same, started with SW_SHOWNORMAL, answered No
                          delete-yes        the same as delete-no, answered Yes
                          registry-delete   the certificate's own key under
                                            HKCU\Software\Microsoft\SystemCertificates\Root\
                                            Certificates deleted directly - the route with no
                                            certutil at all
                          registry-write-after-confirmed-delete,
                          registry-write-after-registry-delete
                                            that key written back, as any program running as the
                                            user can, after each kind of removal
                        After every case: whether Windows still lists the certificate, whether it
                        builds a trusted chain for it, and whether ProtectedRoots - the list of
                        roots the user confirmed, which only the system writes - still names it.
                        A prompt is watched for -VisibleWithinSeconds after it appears, because a
                        dialog exists a moment before it is shown. Then nothing of the certificate
                        is left in the store.

      InstallerCleanup  The product's installer, -InstallerPath, run /VERYSILENT the way the
                        add-in's updater runs it, and then its uninstaller /VERYSILENT, each with
                        VSTO trust entries planted under
                        HKCU\Software\Microsoft\VSTO\Security\Inclusion first: three for the
                        leaked key - its public key spelled three ways the VSTO runtime reads as
                        the same key - at three URLs, one for another key, one with no public key
                        at all. Afterwards the three leaked-key entries must be gone and every
                        other entry, the add-in's own trust entry among them, still there. The
                        leaked key's modulus is read out of -InstallerScript, the Installer.iss
                        the installer was compiled from, so the probe and the installer cannot
                        disagree about it. A public key holds no name: nothing planted names the
                        maintainer, his computer or a user.

    WHAT IT NEVER DOES. It touches no mailbox, no store, no profile and no Outlook. It writes the
    current user's Root store only with the certificate it made, the Inclusion key only with the
    entries it planted, and the registry key of the Root store only for its own certificate. It
    never adds anything to Disallowed and never writes a machine store. The only processes it
    stops are certutil processes it started itself that are still waiting when a case ends.

    OUTPUT. Readable lines on stdout (out.txt of the interactive task) and one JSON file per run in
    -LogDir: trust-removal-<scenario>-<stamp>.json. It is about a test guest, holds no secret and no
    personal data, and stays in .work/ on the host when it is fetched back.

    THE GUARD. vmadmin (Assert-TestbedGuest, guest/OutlookMapiInterop.ps1) AND a computer name
    starting OAI- that is neither the build VM's nor the Exchange VM's, before anything but
    -SelfTest runs; and -Execute, without which a scenario only says what it would do.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`, ASCII only. The C# below is
    C# 5, which is what Windows PowerShell's Add-Type compiles.

.PARAMETER Scenario
    RootStorePrompt or InstallerCleanup.

.PARAMETER Execute
    Without it a scenario prints what it would do and exits 0.

.PARAMETER LogDir
    Where the JSON goes. Created when missing.

.PARAMETER HoldSeconds
    RootStorePrompt: how long delete-unanswered leaves the prompt unanswered.

.PARAMETER VisibleWithinSeconds
    RootStorePrompt: how long a prompt, once it exists, is watched for being shown on the screen.

.PARAMETER InstallerPath
    InstallerCleanup: the compiled OutlookAI installer to run.

.PARAMETER InstallerScript
    InstallerCleanup: the Installer.iss it was compiled from, for the leaked key's modulus.

.PARAMETER SelfTest
    The pure decisions against synthetic inputs, and this file's agreement with Installer.iss when
    the repository is around it. No certificate, no registry, no window, no process. Safe anywhere.

.EXAMPLE
    .\Measure-TrustRemoval.ps1 -SelfTest
    .\Measure-TrustRemoval.ps1 -Scenario RootStorePrompt -Execute
    .\Measure-TrustRemoval.ps1 -Scenario InstallerCleanup -InstallerPath C:\OutlookAI-Q5\trust-removal\OutlookAI-v99.99.99.0.exe -InstallerScript C:\OutlookAI-Q5\trust-removal\Installer.iss -Execute
#>
[CmdletBinding()]
param(
    [ValidateSet('', 'RootStorePrompt', 'InstallerCleanup')] [string] $Scenario = '',
    [switch]   $Execute,
    [string]   $LogDir = 'C:\OutlookAI-Q5\trust-removal\results',
    [int]      $HoldSeconds = 120,
    [int]      $VisibleWithinSeconds = 5,
    [string]   $InstallerPath,
    [string]   $InstallerScript,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'

# The 32-bit certutil: Inno Setup's setup and uninstaller are 32-bit processes, and Exec starts
# 'certutil' with the file system redirector on, so System32 is SysWOW64 for them.
$Certutil32 = Join-Path $env:SystemRoot 'SysWOW64\certutil.exe'
$WorkingDirectory = Join-Path $env:SystemRoot 'System32'
$InclusionKey = 'Software\Microsoft\VSTO\Security\Inclusion'
$RootCertificatesKey = 'Software\Microsoft\SystemCertificates\Root\Certificates'
$StillActive = 259

$script:NativeSource = @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

public static class OaiTrustNative
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct STARTUPINFO
    {
        public int cb; public string lpReserved; public string lpDesktop; public string lpTitle;
        public int dwX; public int dwY; public int dwXSize; public int dwYSize;
        public int dwXCountChars; public int dwYCountChars; public int dwFillAttribute;
        public int dwFlags; public short wShowWindow; public short cbReserved2;
        public IntPtr lpReserved2; public IntPtr hStdInput; public IntPtr hStdOutput; public IntPtr hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct PROCESS_INFORMATION { public IntPtr hProcess; public IntPtr hThread; public int dwProcessId; public int dwThreadId; }

    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool CreateProcessW(string app, StringBuilder cmd, IntPtr pa, IntPtr ta, bool inherit, int flags, IntPtr env, string dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
    [DllImport("kernel32.dll", SetLastError = true)] static extern uint WaitForSingleObject(IntPtr h, uint ms);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool GetExitCodeProcess(IntPtr h, out uint code);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool TerminateProcess(IntPtr h, uint code);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool CloseHandle(IntPtr h);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr hwnd, out int pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr hwnd, StringBuilder sb, int max);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hwnd, uint cmd);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT r);
    [DllImport("user32.dll")] static extern bool PostMessageW(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW")]
    static extern IntPtr SendLength(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)]
    static extern IntPtr SendText(IntPtr hwnd, uint msg, IntPtr wParam, StringBuilder lParam, uint flags, uint timeout, out IntPtr result);

    const int STARTF_USESHOWWINDOW = 0x1;
    const int CREATE_NEW_CONSOLE = 0x10;
    const int CREATE_DEFAULT_ERROR_MODE = 0x04000000;
    const uint WM_GETTEXT = 0x000D;
    const uint WM_GETTEXTLENGTH = 0x000E;
    const uint WM_COMMAND = 0x0111;
    const uint SMTO_ABORTIFHUNG = 0x2;
    const uint GW_OWNER = 4;
    const uint DM_GETDEFID = 0x0400;
    const long DC_HASDEFID = 0x534B;

    public sealed class Launched { public int Pid; public IntPtr Handle; }

    public sealed class Win
    {
        public long Hwnd; public int Pid; public string Class; public string Text; public bool Visible;
        public long Owner; public int ControlId; public int Left; public int Top; public int Right; public int Bottom;
    }

    // As Inno Setup's Exec starts a program: STARTF_USESHOWWINDOW with the show command, no handle
    // inherited, CREATE_DEFAULT_ERROR_MODE - and a console of the child's own, because setup is a
    // GUI process with no console for a console program to share.
    public static Launched Launch(string commandLine, string workingDirectory, int showCommand)
    {
        STARTUPINFO si = new STARTUPINFO();
        si.cb = Marshal.SizeOf(typeof(STARTUPINFO));
        si.dwFlags = STARTF_USESHOWWINDOW;
        si.wShowWindow = (short)showCommand;
        PROCESS_INFORMATION pi;
        if (!CreateProcessW(null, new StringBuilder(commandLine), IntPtr.Zero, IntPtr.Zero, false, CREATE_DEFAULT_ERROR_MODE | CREATE_NEW_CONSOLE, IntPtr.Zero, workingDirectory, ref si, out pi))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        CloseHandle(pi.hThread);
        Launched l = new Launched();
        l.Pid = pi.dwProcessId;
        l.Handle = pi.hProcess;
        return l;
    }

    public static bool HasExited(Launched l, int waitMilliseconds) { return WaitForSingleObject(l.Handle, (uint)waitMilliseconds) == 0; }

    public static long ExitCode(Launched l) { uint c; if (!GetExitCodeProcess(l.Handle, out c)) return -1; return c; }

    public static void Terminate(Launched l) { TerminateProcess(l.Handle, 1); }

    public static void Release(Launched l) { if (l.Handle != IntPtr.Zero) { CloseHandle(l.Handle); l.Handle = IntPtr.Zero; } }

    public static long Foreground() { return GetForegroundWindow().ToInt64(); }

    public static bool Exists(long hwnd) { return IsWindow(new IntPtr(hwnd)); }

    public static bool Visible(long hwnd) { return IsWindowVisible(new IntPtr(hwnd)); }

    // The dialog's default button - the one Enter presses: DM_GETDEFID gives DC_HASDEFID in the high
    // word and the id in the low word. 0 when it has none or does not answer.
    public static int DefaultButtonId(long dialog)
    {
        IntPtr result;
        if (SendLength(new IntPtr(dialog), DM_GETDEFID, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out result) == IntPtr.Zero) return 0;
        long v = result.ToInt64();
        if (((v >> 16) & 0xFFFF) != DC_HASDEFID) return 0;
        return (int)(v & 0xFFFF);
    }

    static string TextOf(IntPtr h)
    {
        IntPtr length;
        if (SendLength(h, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, 2000, out length) == IntPtr.Zero) return "";
        int n = length.ToInt32();
        if (n <= 0) return "";
        StringBuilder sb = new StringBuilder(n + 1);
        IntPtr copied;
        SendText(h, WM_GETTEXT, new IntPtr(n + 1), sb, SMTO_ABORTIFHUNG, 2000, out copied);
        return sb.ToString();
    }

    static Win Describe(IntPtr h)
    {
        Win w = new Win();
        w.Hwnd = h.ToInt64();
        int pid;
        GetWindowThreadProcessId(h, out pid);
        w.Pid = pid;
        StringBuilder cls = new StringBuilder(256);
        GetClassNameW(h, cls, cls.Capacity);
        w.Class = cls.ToString();
        w.Text = TextOf(h);
        w.Visible = IsWindowVisible(h);
        w.Owner = GetWindow(h, GW_OWNER).ToInt64();
        w.ControlId = GetDlgCtrlID(h);
        RECT r;
        if (GetWindowRect(h, out r)) { w.Left = r.Left; w.Top = r.Top; w.Right = r.Right; w.Bottom = r.Bottom; }
        return w;
    }

    // Every top-level window on this desktop, hidden ones included.
    public static List<Win> TopLevel()
    {
        List<Win> list = new List<Win>();
        EnumWindows(delegate (IntPtr h, IntPtr p) { list.Add(Describe(h)); return true; }, IntPtr.Zero);
        return list;
    }

    public static List<Win> Children(long hwnd)
    {
        List<Win> list = new List<Win>();
        EnumChildWindows(new IntPtr(hwnd), delegate (IntPtr h, IntPtr p) { list.Add(Describe(h)); return true; }, IntPtr.Zero);
        return list;
    }

    // What a click on a dialog's button sends: WM_COMMAND, the button's id with BN_CLICKED (0) in
    // the high word, the button's window as lParam. Works on a dialog nobody can see.
    public static bool Press(long dialog, int controlId, long button)
    {
        return PostMessageW(new IntPtr(dialog), WM_COMMAND, new IntPtr(controlId & 0xFFFF), new IntPtr(button));
    }
}
'@

function Add-NativeTypes {
    if (-not ('OaiTrustNative' -as [type])) { Add-Type -TypeDefinition $script:NativeSource -Language CSharp }
}

# =================================================================================================
# PURE DECISIONS. -SelfTest pins every one of them.
# =================================================================================================

# certutil's exit code as Windows reports it: an HRESULT, so 0x8007xxxx for a Win32 error.
function Format-ExitCode {
    param([long] $Code)
    if ($Code -lt 0) { return 'unknown' }
    if ($Code -eq $StillActive) { return 'still running' }
    return ('0x{0:X8}' -f [long]$Code)
}

# The command line Installer.iss's DeleteCompromisedCertificate builds for one store, with
# 'certutil' resolved the way a 32-bit setup resolves it.
function Get-CertutilCommandLine {
    param([Parameter(Mandatory = $true)] [ValidateSet('Store', 'Delete', 'Add')] [string] $Operation, [Parameter(Mandatory = $true)] [string] $Argument, [string] $Certutil = $Certutil32)
    switch ($Operation) {
        'Store' { $verb = '-store' }
        'Delete' { $verb = '-delstore' }
        'Add' { $verb = '-addstore'; $Argument = '"' + $Argument + '"' }
    }
    return '"' + $Certutil + '" -user ' + $verb + ' Root ' + $Argument
}

# The windows that were not there before, other than this process's own.
function Select-NewWindows {
    param([object[]] $Now, [hashtable] $Before, [int] $OwnPid)
    return @($Now | Where-Object { -not $Before.ContainsKey([string]$_.Hwnd) -and $_.Pid -ne $OwnPid })
}

# The prompt among the new windows: a dialog box (#32770). A console window, an IME window or a
# tooltip is not one.
function Select-Prompt {
    param([object[]] $NewWindows)
    $dialogs = @($NewWindows | Where-Object { $_.Class -eq '#32770' })
    if ($dialogs.Count -eq 0) { return $null }
    $visible = @($dialogs | Where-Object { $_.Visible })
    if ($visible.Count -gt 0) { return $visible[0] }
    return $dialogs[0]
}

# The button that gives an answer: by its text (with or without the & of its access key), else by
# the standard dialog id - IDYES 6, IDNO 7, IDOK 1, IDCANCEL 2.
function Select-AnswerButton {
    param([object[]] $Children, [ValidateSet('Yes', 'No')] [string] $Answer)
    $buttons = @($Children | Where-Object { $_.Class -eq 'Button' })
    foreach ($b in $buttons) {
        if ((([string]$b.Text) -replace '&', '').Trim() -ieq $Answer) { return $b }
    }
    $ids = @{ Yes = @(6, 1); No = @(7, 2) }
    foreach ($id in $ids[$Answer]) {
        foreach ($b in $buttons) { if ($b.ControlId -eq $id) { return $b } }
    }
    return $null
}

# A run's verdict from its cases: whether removing a certificate from the user's Root store puts up
# a prompt.
function Get-RootPromptVerdict {
    param([object[]] $Cases)
    $deletes = @($Cases | Where-Object { $_.Name -like 'delete-*' })
    if ($deletes.Count -eq 0) { return 'INCONCLUSIVE: no removal case ran' }
    $prompted = @($deletes | Where-Object { @($_.Prompts).Count -gt 0 })
    $silentGone = @($deletes | Where-Object { @($_.Prompts).Count -eq 0 -and $_.ExitCode -eq 0 -and -not $_.PresentAfter })
    $blocked = @($deletes | Where-Object { @($_.Prompts).Count -eq 0 -and $_.ExitCode -eq $StillActive })
    if ($prompted.Count -eq $deletes.Count) { return 'REMOVAL-PROMPTS' }
    if ($silentGone.Count -eq $deletes.Count) { return 'REMOVAL-SILENT' }
    if ($blocked.Count -gt 0) { return 'REMOVAL-BLOCKS-WITH-NO-WINDOW' }
    return 'INCONCLUSIVE: the removal cases disagree'
}

# The leaked key's modulus from Installer.iss: the one #define CompromisedKeyModulus.
function Get-CompromisedModulus {
    param([Parameter(Mandatory = $true)] [string] $IssText)
    $m = [regex]::Matches($IssText, '(?m)^#define CompromisedKeyModulus "([A-Za-z0-9+/=]+)"\r?$')
    if ($m.Count -ne 1) { throw "Installer.iss defines CompromisedKeyModulus $($m.Count) times; expected once." }
    $value = $m[0].Groups[1].Value
    $bytes = [Convert]::FromBase64String($value)
    if ([Convert]::ToBase64String($bytes) -cne $value) { throw 'CompromisedKeyModulus is not canonical base64.' }
    if ($bytes.Length -lt 64 -or $bytes[0] -lt 0x80) { throw "CompromisedKeyModulus is not an RSA modulus: $($bytes.Length) bytes, first byte $($bytes[0])." }
    return $value
}

# The public key XML of an entry the probe plants, spelled the ways the VSTO runtime reads as one
# key (its own comparison - FromXmlString, then ExportCspBlob - measured equal for all three).
function New-KeyXml {
    param([Parameter(Mandatory = $true)] [string] $ModulusBase64, [string] $ExponentBase64 = 'AQAB', [ValidateSet('Canonical', 'Wrapped', 'Prefixed')] [string] $Spelling = 'Canonical')
    switch ($Spelling) {
        'Canonical' { return '<RSAKeyValue><Modulus>' + $ModulusBase64 + '</Modulus><Exponent>' + $ExponentBase64 + '</Exponent></RSAKeyValue>' }
        'Wrapped' {
            $lines = @()
            for ($i = 0; $i -lt $ModulusBase64.Length; $i += 64) { $lines += $ModulusBase64.Substring($i, [Math]::Min(64, $ModulusBase64.Length - $i)) }
            return "<RSAKeyValue xmlns=`"http://www.w3.org/2000/09/xmldsig#`">`r`n  <Modulus>`r`n    " + ($lines -join "`r`n    ") + "`r`n  </Modulus>`r`n  <Exponent>" + $ExponentBase64 + "</Exponent>`r`n</RSAKeyValue>"
        }
        'Prefixed' { return '<ds:RSAKeyValue xmlns:ds="http://www.w3.org/2000/09/xmldsig#"><ds:Exponent>' + $ExponentBase64 + '</ds:Exponent><ds:Modulus>' + $ModulusBase64 + '</ds:Modulus></ds:RSAKeyValue>' }
    }
}

# Which planted entries a run expects gone and which kept, against what it found afterwards.
function Get-CleanupVerdict {
    param([object[]] $Planted, [string[]] $NamesAfter, [string[]] $KeptBefore)
    $problems = @()
    foreach ($p in $Planted) {
        $present = $NamesAfter -contains $p.Name
        if ($p.ExpectGone -and $present) { $problems += "$($p.Label) ($($p.Name)) is still there" }
        if (-not $p.ExpectGone -and -not $present) { $problems += "$($p.Label) ($($p.Name)) was removed" }
    }
    foreach ($k in $KeptBefore) {
        if ($NamesAfter -notcontains $k) { $problems += "the entry that was there before, $k, was removed" }
    }
    if ($problems.Count -eq 0) { return [pscustomobject]@{ Verdict = 'PASS'; Problems = @() } }
    return [pscustomobject]@{ Verdict = 'FAIL'; Problems = $problems }
}

# =================================================================================================
# SELF-TEST
# =================================================================================================

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $name, [bool] $ok) { if ($ok) { $script:p++; Write-Host "  PASS  $name" } else { $script:f++; Write-Host "  FAIL  $name" } }

    $compiled = $true
    try { Add-NativeTypes } catch { $compiled = $false; Write-Host "  compile error: $($_.Exception.Message)" }
    Check 'the native helper compiles under this PowerShell' $compiled

    Check 'exit 0 reads as 0x00000000' ((Format-ExitCode 0) -eq '0x00000000')
    Check 'ERROR_CANCELLED as an HRESULT reads as 0x800704C7' ((Format-ExitCode 2147943623) -eq '0x800704C7')
    Check 'STILL_ACTIVE reads as still running' ((Format-ExitCode 259) -eq 'still running')

    $t = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'
    Check 'the delete is the installer''s -user -delstore Root <thumbprint>' ((Get-CertutilCommandLine -Operation Delete -Argument $t -Certutil 'C:\W\SysWOW64\certutil.exe') -ceq ('"C:\W\SysWOW64\certutil.exe" -user -delstore Root ' + $t))
    Check 'the check is the installer''s -user -store Root <thumbprint>' ((Get-CertutilCommandLine -Operation Store -Argument $t -Certutil 'c') -ceq ('"c" -user -store Root ' + $t))
    Check 'the add quotes the file' ((Get-CertutilCommandLine -Operation Add -Argument 'C:\a b\x.cer' -Certutil 'c') -ceq '"c" -user -addstore Root "C:\a b\x.cer"')
    Check 'certutil is the 32-bit one a 32-bit setup reaches' ($Certutil32 -like '*\SysWOW64\certutil.exe')

    # Bound to Installer.iss when the repository is around this file (the build VM's archive is).
    $iss = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'Installer.iss'
    if (Test-Path -LiteralPath $iss) {
        $issText = [System.IO.File]::ReadAllText($iss)
        Check 'Installer.iss still deletes with Scope + -delstore + Store + thumbprint' ($issText.Contains("Scope + '-delstore ' + Store + ' {#CompromisedCertThumbprint}'"))
        Check 'Installer.iss still checks with Scope + -store + Store + thumbprint' ($issText.Contains("Scope + '-store ' + Store + ' {#CompromisedCertThumbprint}'"))
        Check 'Installer.iss still gives the user scope as -user' ($issText.Contains("DeleteCompromisedCertificateFromStores('-user ', '{#CompromisedCertUserStores}');"))
        Check 'Installer.iss still starts certutil hidden' ($issText -match "Exec\('certutil', Scope \+ '-delstore ' \+ Store \+ ' \{#CompromisedCertThumbprint\}', '', SW_HIDE")
        Check 'Installer.iss still names this Inclusion key' ($issText.Contains("'" + $InclusionKey + "'"))
        $modulus = $null
        try { $modulus = Get-CompromisedModulus $issText } catch { Write-Host "  $($_.Exception.Message)" }
        Check 'Installer.iss defines the leaked key''s modulus once, canonical, 128 bytes' ($null -ne $modulus -and ([Convert]::FromBase64String($modulus)).Length -eq 128)
    }
    else {
        Write-Host "  (no Installer.iss beside the repository at $iss - the agreement checks did not run)"
    }

    Check 'a modulus defined twice is refused' ($null -eq $(try { Get-CompromisedModulus ("#define CompromisedKeyModulus `"AQAB`"`n#define CompromisedKeyModulus `"AQAB`"") } catch { $null }))
    $fake = [Convert]::ToBase64String([byte[]](@(0xC0) + @(1..127 | ForEach-Object { [byte]($_ % 256) })))
    Check 'a modulus defined once is read' ((Get-CompromisedModulus ("x`n#define CompromisedKeyModulus `"$fake`"`ny")) -ceq $fake)
    $padded = [Convert]::ToBase64String([byte[]](@(0) + @(0xC0) + @(1..127 | ForEach-Object { [byte]($_ % 256) })))
    Check 'a modulus with a leading zero byte is refused' ($null -eq $(try { Get-CompromisedModulus "#define CompromisedKeyModulus `"$padded`"" } catch { $null }))

    foreach ($s in @('Canonical', 'Wrapped', 'Prefixed')) {
        $xml = New-KeyXml -ModulusBase64 $fake -Spelling $s
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($xml)
        $mod = $doc.SelectSingleNode("//*[local-name()='Modulus']").InnerText -replace '\s', ''
        Check "the $s spelling is XML that holds the modulus" ($mod -ceq $fake)
        Check "the $s spelling holds the modulus once its white space is gone" (($xml -replace '\s', '').Contains($fake))
    }
    Check 'the canonical spelling is what the runtime writes' ((New-KeyXml -ModulusBase64 'AA==' ) -ceq '<RSAKeyValue><Modulus>AA==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')

    $before = @{ '1' = $true; '2' = $true }
    $now = @(
        [pscustomobject]@{ Hwnd = 1; Pid = 10; Class = '#32770'; Visible = $true },
        [pscustomobject]@{ Hwnd = 3; Pid = 99; Class = 'ConsoleWindowClass'; Visible = $false },
        [pscustomobject]@{ Hwnd = 4; Pid = 99; Class = '#32770'; Visible = $false },
        [pscustomobject]@{ Hwnd = 5; Pid = 7; Class = '#32770'; Visible = $true })
    $new = @(Select-NewWindows -Now $now -Before $before -OwnPid 7)
    Check 'only windows that are new and not this process''s are new' ((($new | ForEach-Object { $_.Hwnd }) -join ',') -eq '3,4')
    Check 'the prompt is the dialog, not the console window' ((Select-Prompt $new).Hwnd -eq 4)
    Check 'a visible dialog wins over a hidden one' ((Select-Prompt @($now[2], [pscustomobject]@{ Hwnd = 8; Pid = 1; Class = '#32770'; Visible = $true })).Hwnd -eq 8)
    Check 'no dialog is no prompt' ($null -eq (Select-Prompt @($now[1])))

    $kids = @(
        [pscustomobject]@{ Hwnd = 11; Class = 'Static'; Text = 'Do you want to?'; ControlId = 65535 },
        [pscustomobject]@{ Hwnd = 12; Class = 'Button'; Text = '&Yes'; ControlId = 6 },
        [pscustomobject]@{ Hwnd = 13; Class = 'Button'; Text = '&No'; ControlId = 7 })
    Check 'Yes is found by its text' ((Select-AnswerButton $kids 'Yes').Hwnd -eq 12)
    Check 'No is found by its text' ((Select-AnswerButton $kids 'No').Hwnd -eq 13)
    $untitled = @([pscustomobject]@{ Hwnd = 21; Class = 'Button'; Text = 'Ja'; ControlId = 6 }, [pscustomobject]@{ Hwnd = 22; Class = 'Button'; Text = 'Nee'; ControlId = 7 })
    Check 'in another language Yes is IDYES' ((Select-AnswerButton $untitled 'Yes').Hwnd -eq 21)
    Check 'in another language No is IDNO' ((Select-AnswerButton $untitled 'No').Hwnd -eq 22)
    Check 'a static text is never a button' ($null -eq (Select-AnswerButton @($kids[0]) 'Yes'))

    $prompt = [pscustomobject]@{ Hwnd = 1 }
    Check 'every removal prompting is REMOVAL-PROMPTS' ((Get-RootPromptVerdict @([pscustomobject]@{ Name = 'add-yes'; Prompts = @($prompt) }, [pscustomobject]@{ Name = 'delete-no'; Prompts = @($prompt) }, [pscustomobject]@{ Name = 'delete-yes'; Prompts = @($prompt) })) -eq 'REMOVAL-PROMPTS')
    Check 'every removal silent and done is REMOVAL-SILENT' ((Get-RootPromptVerdict @([pscustomobject]@{ Name = 'delete-no'; Prompts = @(); ExitCode = 0; PresentAfter = $false })) -eq 'REMOVAL-SILENT')
    Check 'a removal still running with no window says so' ((Get-RootPromptVerdict @([pscustomobject]@{ Name = 'delete-no'; Prompts = @(); ExitCode = 259; PresentAfter = $true })) -eq 'REMOVAL-BLOCKS-WITH-NO-WINDOW')
    Check 'no removal case is inconclusive' ((Get-RootPromptVerdict @([pscustomobject]@{ Name = 'add-yes'; Prompts = @() })) -like 'INCONCLUSIVE*')

    $planted = @(
        [pscustomobject]@{ Label = 'leaked'; Name = 'a'; ExpectGone = $true },
        [pscustomobject]@{ Label = 'other'; Name = 'b'; ExpectGone = $false })
    Check 'the leaked entry gone and the rest kept passes' ((Get-CleanupVerdict -Planted $planted -NamesAfter @('b', 'own') -KeptBefore @('own')).Verdict -eq 'PASS')
    Check 'a leaked entry left fails' ((Get-CleanupVerdict -Planted $planted -NamesAfter @('a', 'b', 'own') -KeptBefore @('own')).Verdict -eq 'FAIL')
    Check 'another key''s entry removed fails' ((Get-CleanupVerdict -Planted $planted -NamesAfter @('own') -KeptBefore @('own')).Verdict -eq 'FAIL')
    Check 'an entry that was there before, removed, fails' ((Get-CleanupVerdict -Planted $planted -NamesAfter @('b') -KeptBefore @('own')).Verdict -eq 'FAIL')

    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $stops = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Stop-Process' }, $true))
    Check 'nothing here calls Stop-Process' ($stops.Count -eq 0)
    $text = [System.IO.File]::ReadAllText($PSCommandPath)
    Check 'nothing here adds to Disallowed' (-not ($text -match ('add' + 'store\s+Disallowed')))
    Check 'nothing here reads or writes a machine certificate store' (-not $text.Contains('Cert:\Local' + 'Machine'))
    Check 'every certutil command names the user scope' (@('Store', 'Delete', 'Add' | Where-Object { -not (Get-CertutilCommandLine -Operation $_ -Argument 'x' -Certutil 'c').Contains(' -user ') }).Count -eq 0)

    Write-Host ''
    Write-Host "SelfTest: $script:p passed, $script:f failed"
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

if (-not $Scenario) { throw 'Say which -Scenario: RootStorePrompt or InstallerCleanup. -SelfTest runs the self-test.' }

# THE GUARD, before anything that touches a certificate, the registry, a window or a process.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser
if ($env:COMPUTERNAME -notlike 'OAI-*' -or $env:COMPUTERNAME -eq 'OAI-BUILD' -or $env:COMPUTERNAME -eq 'OAI-EXCHANGE') {
    throw "REFUSING TO RUN on '$env:COMPUTERNAME': this probe changes the user's Root store and VSTO trust entries, and runs only on an Outlook test guest (OAI-*, not the build VM, not the Exchange VM)."
}

$self = [System.Diagnostics.Process]::GetCurrentProcess()
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$elevated = (New-Object System.Security.Principal.WindowsPrincipal($identity)).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
$os = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$facts = [ordered]@{
    scenario    = $Scenario
    computer    = $env:COMPUTERNAME
    session     = $self.SessionId
    elevated    = $elevated
    windows     = "$($os.ProductName) $($os.DisplayVersion) build $($os.CurrentBuild).$($os.UBR)"
    startedUtc  = [DateTime]::UtcNow.ToString('o')
    guestClock  = 'the guest''s own; on a frozen guest it stands at its checkpoint''s instant'
}
Write-Host "== Measure-TrustRemoval: $Scenario on $($facts.computer), session $($facts.session), elevated $elevated, $($facts.windows)"

if (-not $Execute) {
    Write-Host 'DRY RUN - nothing was changed. Add -Execute to run the scenario.'
    exit 0
}
if ($self.SessionId -eq 0) { throw 'REFUSING: this is session 0, which has no desktop for a prompt to appear on. Run it through Register-InteractiveTask.ps1.' }
if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$jsonPath = Join-Path $LogDir ("trust-removal-{0}-{1}.json" -f $Scenario, $stamp)
Add-NativeTypes
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Get-WindowSet {
    $set = @{}
    foreach ($w in [OaiTrustNative]::TopLevel()) { $set[[string]$w.Hwnd] = $true }
    return $set
}

function Get-ProcessName([int] $ProcessId) {
    try { return (Get-Process -Id $ProcessId -ErrorAction Stop).ProcessName } catch { return '(gone)' }
}

# The UI Automation tree of a window: every element's type and name. Empty when UI Automation
# cannot reach it.
function Get-UiaReading([long] $Hwnd) {
    try {
        $el = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Hwnd)
        $all = $el.FindAll([System.Windows.Automation.TreeScope]::Subtree, [System.Windows.Automation.Condition]::TrueCondition)
        $out = @()
        foreach ($e in $all) {
            $patterns = @($e.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName -replace 'Identifiers\.Pattern$', '' })
            $out += [pscustomobject]@{ type = $e.Current.ControlType.ProgrammaticName; name = $e.Current.Name; automationId = $e.Current.AutomationId; offscreen = $e.Current.IsOffscreen; patterns = ($patterns -join ',') }
        }
        return $out
    }
    catch { return @([pscustomobject]@{ error = $_.Exception.Message }) }
}

function Get-PromptRecord([object] $Window, [long] $ForegroundBefore, [double] $AppearedSeconds) {
    $children = @([OaiTrustNative]::Children($Window.Hwnd))
    return [pscustomobject]@{
        hwnd            = $Window.Hwnd
        pid             = $Window.Pid
        process         = (Get-ProcessName $Window.Pid)
        class           = $Window.Class
        title           = $Window.Text
        visible         = $Window.Visible
        rect            = "$($Window.Left),$($Window.Top),$($Window.Right),$($Window.Bottom)"
        foreground      = ([OaiTrustNative]::Foreground() -eq $Window.Hwnd)
        foregroundBefore = $ForegroundBefore
        owner           = $Window.Owner
        appearedSeconds = [Math]::Round($AppearedSeconds, 2)
        controls        = @($children | ForEach-Object { [pscustomobject]@{ class = $_.Class; id = $_.ControlId; text = $_.Text; visible = $_.Visible } })
        uia             = @(Get-UiaReading $Window.Hwnd)
        answeredWith    = $null
        answeredBy      = $null
        children        = $children
    }
}

# Answers a prompt: UI Automation's Invoke on the button's element, as a person's click, when it
# offers one; otherwise the dialog's own WM_COMMAND for that button, which is what a click sends and
# which reaches a dialog nobody can see. Says which it used, and why UI Automation did not.
function Send-Answer([object] $Prompt, [string] $Answer) {
    $button = Select-AnswerButton -Children $Prompt.children -Answer $Answer
    $uiaWhy = 'no element for the button'
    try {
        $el = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Prompt.hwnd)
        foreach ($b in $el.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
            $match = ($b.Current.Name.Replace('&', '').Trim() -ieq $Answer)
            if (-not $match -and $null -ne $button) { $match = ($b.Current.NativeWindowHandle -eq $button.Hwnd) }
            if (-not $match) { continue }
            $invoke = $null
            if ($b.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref] $invoke)) {
                $invoke.Invoke()
                return 'UI Automation Invoke on the ' + $b.Current.ControlType.ProgrammaticName + ' "' + $b.Current.Name + '"'
            }
            $uiaWhy = 'its element, a ' + $b.Current.ControlType.ProgrammaticName + ', offers no InvokePattern'
            break
        }
    }
    catch { $uiaWhy = 'UI Automation threw: ' + $_.Exception.Message }
    if ($null -eq $button) { return 'no button for ' + $Answer + ' - left unanswered (' + $uiaWhy + ')' }
    if ([OaiTrustNative]::Press($Prompt.hwnd, $button.ControlId, $button.Hwnd)) { return "WM_COMMAND id $($button.ControlId) (`"$($button.Text)`"), because $uiaWhy" }
    return 'WM_COMMAND could not be posted'
}

# Whether Windows builds a trusted chain for the certificate as it stands: what a root in the store
# is for. No revocation check - the throwaway has nowhere to check.
function Test-TrustedAsRoot([string] $CerPath) {
    $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CerPath
    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    try {
        $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
        $ok = $chain.Build($c)
        $status = @($chain.ChainStatus | ForEach-Object { [string]$_.Status }) -join ','
        return [pscustomobject]@{ trusted = $ok; status = $status }
    }
    finally { $chain.Reset(); $c.Reset() }
}

# Whether Windows' list of the roots a user confirmed - the ProtectedRoots value, which only the
# system writes - still names the certificate: its SHA-1 among the value's bytes.
function Test-InProtectedRoots([string] $Thumbprint) {
    try {
        $blob = [byte[]](Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\SystemCertificates\Root\ProtectedRoots' -Name Certificates -ErrorAction Stop).Certificates
        $hash = New-Object byte[] 20
        for ($i = 0; $i -lt 20; $i++) { $hash[$i] = [Convert]::ToByte($Thumbprint.Substring($i * 2, 2), 16) }
        $hex = [BitConverter]::ToString($blob)
        return [string]($hex.Contains([BitConverter]::ToString($hash)))
    }
    catch { return 'unreadable: ' + $_.Exception.Message }
}

function Test-InUserRoot([string] $Thumbprint) {
    return [pscustomobject]@{
        store    = (Test-Path -LiteralPath ("Cert:\CurrentUser\Root\" + $Thumbprint))
        registry = (Test-Path -LiteralPath ("HKCU:\" + $RootCertificatesKey + '\' + $Thumbprint))
    }
}

# One certutil run, started as Inno's Exec starts it, every prompt it raises recorded and answered.
function Invoke-CertutilCase {
    param(
        [string] $Name, [string] $CommandLine, [int] $Show, [ValidateSet('Yes', 'No', 'None')] [string] $Answer,
        [string] $Thumbprint, [int] $Hold = 0, [int] $AppearSeconds = 30, [int] $ExitSeconds = 60
    )
    Write-Host ''
    Write-Host "-- case $Name : $CommandLine (show $Show, answer $Answer$(if ($Hold -gt 0) { ", held $Hold s first" }))"
    $before = Get-WindowSet
    $fg = [OaiTrustNative]::Foreground()
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $proc = [OaiTrustNative]::Launch($CommandLine, $WorkingDirectory, $Show)
    $prompts = @()
    $heldFor = $null
    $stillThereAfterHold = $null
    $exited = $false
    try {
        $deadline = $AppearSeconds
        while ($clock.Elapsed.TotalSeconds -lt $deadline) {
            if ([OaiTrustNative]::HasExited($proc, 0)) { $exited = $true; break }
            $new = @(Select-NewWindows -Now ([OaiTrustNative]::TopLevel()) -Before $before -OwnPid $self.Id)
            $window = Select-Prompt $new
            if ($null -eq $window) { Start-Sleep -Milliseconds 200; continue }
            $record = Get-PromptRecord -Window $window -ForegroundBefore $fg -AppearedSeconds $clock.Elapsed.TotalSeconds
            # A dialog exists a moment before it is shown, so "not visible" when it is first seen says
            # nothing yet: watch it for VisibleWithinSeconds, as a person at the screen would.
            $shownAt = $null
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            while ($watch.Elapsed.TotalSeconds -lt $VisibleWithinSeconds -and [OaiTrustNative]::Exists($window.Hwnd)) {
                if ([OaiTrustNative]::Visible($window.Hwnd)) { $shownAt = [Math]::Round($clock.Elapsed.TotalSeconds, 2); break }
                Start-Sleep -Milliseconds 100
            }
            $record | Add-Member -NotePropertyName shownAtSeconds -NotePropertyValue $shownAt
            $record | Add-Member -NotePropertyName foregroundOnceShown -NotePropertyValue ([OaiTrustNative]::Foreground() -eq $window.Hwnd)
            $record | Add-Member -NotePropertyName defaultButtonId -NotePropertyValue ([OaiTrustNative]::DefaultButtonId($window.Hwnd))
            $record | Add-Member -NotePropertyName uiaOnceShown -NotePropertyValue @(Get-UiaReading $window.Hwnd)
            Write-Host ("   PROMPT after {0} s: '{1}' class {2}, pid {3} ({4}); visible when first seen {5}, shown at {6} s, foreground once shown {7}, default button {8}" -f $record.appearedSeconds, $record.title, $record.class, $record.pid, $record.process, $record.visible, $(if ($null -eq $shownAt) { "NEVER within $VisibleWithinSeconds" } else { $shownAt }), $record.foregroundOnceShown, $record.defaultButtonId)
            foreach ($c in $record.controls) { if ($c.text) { Write-Host ("     [{0} {1}] {2}" -f $c.class, $c.id, ($c.text -replace "`r?`n", ' | ')) } }
            foreach ($u in @($record.uiaOnceShown | Where-Object { $_.automationId -eq '6' -or $_.automationId -eq '7' })) { Write-Host ("     UIA: {0} '{1}' id {2}, patterns {3}" -f $u.type, $u.name, $u.automationId, $u.patterns) }
            if ($Hold -gt 0 -and $null -eq $heldFor) {
                $h = [System.Diagnostics.Stopwatch]::StartNew()
                while ($h.Elapsed.TotalSeconds -lt $Hold) {
                    if ([OaiTrustNative]::HasExited($proc, 1000)) { break }
                    if (-not [OaiTrustNative]::Exists($window.Hwnd)) { break }
                }
                $heldFor = [Math]::Round($h.Elapsed.TotalSeconds, 1)
                $stillThereAfterHold = [pscustomobject]@{ prompt = [OaiTrustNative]::Exists($window.Hwnd); visible = [OaiTrustNative]::Visible($window.Hwnd); certutilRunning = -not [OaiTrustNative]::HasExited($proc, 0) }
                Write-Host ("   after {0} s unanswered: prompt still there {1}, visible {2}, certutil still running {3}" -f $heldFor, $stillThereAfterHold.prompt, $stillThereAfterHold.visible, $stillThereAfterHold.certutilRunning)
            }
            $how = 'not answered (expected no prompt)'
            if ($Answer -ne 'None') { $how = Send-Answer -Prompt $record -Answer $Answer }
            $record.answeredWith = $Answer
            $record.answeredBy = $how
            Write-Host "   answered $Answer by $how"
            $record.PSObject.Properties.Remove('children')
            $prompts += $record
            $before[[string]$window.Hwnd] = $true
            # The next prompt, if this one is followed by another, gets as long again.
            $deadline = $clock.Elapsed.TotalSeconds + $AppearSeconds
            if ($Answer -eq 'None') { break }
            for ($i = 0; $i -lt 25; $i++) { if (-not [OaiTrustNative]::Exists($window.Hwnd)) { break }; Start-Sleep -Milliseconds 200 }
        }
        if (-not $exited) { $exited = [OaiTrustNative]::HasExited($proc, $ExitSeconds * 1000) }
        $code = [OaiTrustNative]::ExitCode($proc)
        $terminated = $false
        if (-not $exited) {
            [OaiTrustNative]::Terminate($proc)
            $terminated = $true
            $null = [OaiTrustNative]::HasExited($proc, 5000)
        }
    }
    finally { [OaiTrustNative]::Release($proc) }
    $state = Test-InUserRoot $Thumbprint
    $trust = Test-TrustedAsRoot $script:ProbeCer
    $protected = Test-InProtectedRoots $Thumbprint
    Write-Host ("   certutil exit {0} after {1} s{2}; in the user's Root store now: {3} (its registry key: {4}); trusted as a root: {5}{6}; in ProtectedRoots: {7}" -f (Format-ExitCode $code), [Math]::Round($clock.Elapsed.TotalSeconds, 1), $(if ($terminated) { ', NOT ENDED - terminated by the probe' } else { '' }), $state.store, $state.registry, $trust.trusted, $(if ($trust.status) { " ($($trust.status))" } else { '' }), $protected)
    return [pscustomobject]@{
        Name = $Name; CommandLine = $CommandLine; Show = $Show; Answer = $Answer
        Prompts = $prompts; HeldSeconds = $heldFor; AfterHold = $stillThereAfterHold
        ExitCode = $code; ExitCodeText = (Format-ExitCode $code); EndedOnItsOwn = (-not $terminated)
        Seconds = [Math]::Round($clock.Elapsed.TotalSeconds, 2); PresentAfter = $state.store; RegistryKeyAfter = $state.registry
        TrustedAfter = $trust.trusted; ChainStatusAfter = $trust.status; InProtectedRootsAfter = $protected
    }
}

# The certificate's own key in the user's registry Root store, every value of it, to write back later.
function Save-RootKey([string] $Thumbprint) {
    $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($RootCertificatesKey + '\' + $Thumbprint, $false)
    if ($null -eq $k) { return @() }
    try {
        $values = @()
        foreach ($n in $k.GetValueNames()) { $values += [pscustomobject]@{ Name = $n; Kind = $k.GetValueKind($n); Data = $k.GetValue($n, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) } }
        return $values
    }
    finally { $k.Close() }
}

# Writes that key back with no certutil and no crypt32 - what any program running as the user can
# do - and says whether Windows then lists and trusts the certificate again. ProtectedRoots, the list
# of roots the user confirmed, is the difference this case looks for.
function Invoke-RootKeyWriteCase([string] $Name, [string] $Thumbprint, [object[]] $Saved) {
    Write-Host ''
    Write-Host "-- case $Name : the certificate's key written back under HKCU:\$RootCertificatesKey, no certutil"
    $before = Get-WindowSet
    $err = $null
    try {
        $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($RootCertificatesKey + '\' + $Thumbprint)
        try { foreach ($v in $Saved) { $k.SetValue($v.Name, $v.Data, $v.Kind) } } finally { $k.Close() }
    }
    catch { $err = $_.Exception.Message }
    Start-Sleep -Seconds 3
    $prompt = Select-Prompt @(Select-NewWindows -Now ([OaiTrustNative]::TopLevel()) -Before $before -OwnPid $self.Id)
    $state = Test-InUserRoot $Thumbprint
    $trust = Test-TrustedAsRoot $script:ProbeCer
    $protected = Test-InProtectedRoots $Thumbprint
    Write-Host ("   {0} ({1} value(s)); a prompt within 3 s: {2}; in the user's Root store now: {3} (its registry key: {4}); trusted as a root: {5}{6}; in ProtectedRoots: {7}" -f $(if ($err) { "FAILED: $err" } else { 'written' }), @($Saved).Count, ($null -ne $prompt), $state.store, $state.registry, $trust.trusted, $(if ($trust.status) { " ($($trust.status))" } else { '' }), $protected)
    if ($prompt) { $null = Send-Answer -Prompt (Get-PromptRecord -Window $prompt -ForegroundBefore 0 -AppearedSeconds 0) -Answer No }
    return [pscustomobject]@{
        Name = $Name; CommandLine = "the key written back under HKCU:\$RootCertificatesKey\<thumbprint>"; Show = $null; Answer = 'None'
        Prompts = @(); HeldSeconds = $null; AfterHold = $null
        ExitCode = $(if ($err) { 1 } else { 0 }); ExitCodeText = $(if ($err) { $err } else { 'written' }); EndedOnItsOwn = $true; Seconds = 3
        PresentAfter = $state.store; RegistryKeyAfter = $state.registry
        TrustedAfter = $trust.trusted; ChainStatusAfter = $trust.status; InProtectedRootsAfter = $protected
    }
}

function Invoke-RootStorePrompt {
    $result = [ordered]@{ facts = $facts; certificate = $null; cases = @(); verdict = $null }
    $subject = 'CN=OutlookAI trust-removal probe ' + $stamp
    $rsa = New-Object System.Security.Cryptography.RSACng(2048)
    try {
        $request = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest($subject, $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $made = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(30))
    }
    finally { $rsa.Dispose() }
    $cerPath = Join-Path $LogDir ("probe-{0}.cer" -f $stamp)
    [System.IO.File]::WriteAllBytes($cerPath, $made.RawData)
    $thumb = $made.Thumbprint
    $made.Dispose()
    $script:ProbeCer = $cerPath
    $result.certificate = [ordered]@{ subject = $subject; thumbprint = $thumb; file = $cerPath; privateKey = 'never stored: made in memory, only the public certificate written' }
    Write-Host "certificate: $subject, $thumb (public part only, $cerPath)"
    if ((Test-InUserRoot $thumb).store) { throw "The throwaway certificate $thumb is already in the user's Root store - a fresh one cannot be." }

    $hide = 0; $normal = 1
    $add = Get-CertutilCommandLine -Operation Add -Argument $cerPath
    $check = Get-CertutilCommandLine -Operation Store -Argument $thumb
    $delete = Get-CertutilCommandLine -Operation Delete -Argument $thumb
    $cases = @()
    $cases += Invoke-CertutilCase -Name 'store-absent' -CommandLine $check -Show $hide -Answer None -Thumbprint $thumb -AppearSeconds 10
    $cases += Invoke-CertutilCase -Name 'add-no' -CommandLine $add -Show $hide -Answer No -Thumbprint $thumb
    $cases += Invoke-CertutilCase -Name 'add-yes' -CommandLine $add -Show $hide -Answer Yes -Thumbprint $thumb
    $savedKey = @(Save-RootKey $thumb)
    $cases += Invoke-CertutilCase -Name 'store-present' -CommandLine $check -Show $hide -Answer None -Thumbprint $thumb -AppearSeconds 10
    foreach ($spec in @(
            @{ Name = 'delete-no'; Show = $hide; Answer = 'No'; Hold = 0 },
            @{ Name = 'delete-unanswered'; Show = $hide; Answer = 'No'; Hold = $HoldSeconds },
            @{ Name = 'delete-visible'; Show = $normal; Answer = 'No'; Hold = 0 },
            @{ Name = 'delete-yes'; Show = $hide; Answer = 'Yes'; Hold = 0 })) {
        if (-not (Test-InUserRoot $thumb).store) {
            $cases += Invoke-CertutilCase -Name ('re-add before ' + $spec.Name) -CommandLine $add -Show $hide -Answer Yes -Thumbprint $thumb
        }
        $cases += Invoke-CertutilCase -Name $spec.Name -CommandLine $delete -Show $spec.Show -Answer $spec.Answer -Thumbprint $thumb -Hold $spec.Hold
    }

    # After a removal the user confirmed: does writing the key back make it a root again?
    if ($savedKey.Count -gt 0 -and -not (Test-InUserRoot $thumb).registry) {
        $cases += Invoke-RootKeyWriteCase -Name 'registry-write-after-confirmed-delete' -Thumbprint $thumb -Saved $savedKey
        Remove-Item -LiteralPath ("HKCU:\" + $RootCertificatesKey + '\' + $thumb) -Recurse -Force -ErrorAction SilentlyContinue
    }

    # The route with no certutil: the certificate's own key in the registry store under HKCU.
    if (-not (Test-InUserRoot $thumb).store) {
        $cases += Invoke-CertutilCase -Name 're-add before registry-delete' -CommandLine $add -Show $hide -Answer Yes -Thumbprint $thumb
    }
    Write-Host ''
    Write-Host "-- case registry-delete : Remove-Item HKCU:\$RootCertificatesKey\$thumb"
    $before = Get-WindowSet
    $regClock = [System.Diagnostics.Stopwatch]::StartNew()
    $regError = $null
    try { Remove-Item -LiteralPath ("HKCU:\" + $RootCertificatesKey + '\' + $thumb) -Recurse -Force -ErrorAction Stop } catch { $regError = $_.Exception.Message }
    $regSeconds = [Math]::Round($regClock.Elapsed.TotalSeconds, 2)
    Start-Sleep -Seconds 5
    $regPrompt = Select-Prompt @(Select-NewWindows -Now ([OaiTrustNative]::TopLevel()) -Before $before -OwnPid $self.Id)
    $regState = Test-InUserRoot $thumb
    $regTrust = Test-TrustedAsRoot $cerPath
    $regProtected = Test-InProtectedRoots $thumb
    Write-Host ("   {0} in {1} s; a prompt within 5 s: {2}; in the user's Root store now: {3} (its registry key: {4}); trusted as a root: {5}{6}; in ProtectedRoots: {7}" -f $(if ($regError) { "FAILED: $regError" } else { 'deleted' }), $regSeconds, ($null -ne $regPrompt), $regState.store, $regState.registry, $regTrust.trusted, $(if ($regTrust.status) { " ($($regTrust.status))" } else { '' }), $regProtected)
    $cases += [pscustomobject]@{
        Name = 'registry-delete'; CommandLine = "Remove-Item HKCU:\$RootCertificatesKey\<thumbprint>"; Show = $null; Answer = 'None'
        Prompts = @($(if ($regPrompt) { Get-PromptRecord -Window $regPrompt -ForegroundBefore 0 -AppearedSeconds 0 })); HeldSeconds = $null; AfterHold = $null
        ExitCode = $(if ($regError) { 1 } else { 0 }); ExitCodeText = $(if ($regError) { $regError } else { 'deleted' }); EndedOnItsOwn = $true
        Seconds = $regSeconds; PresentAfter = $regState.store; RegistryKeyAfter = $regState.registry
        TrustedAfter = $regTrust.trusted; ChainStatusAfter = $regTrust.status; InProtectedRootsAfter = $regProtected
    }
    if ($regPrompt) { $null = Send-Answer -Prompt (Get-PromptRecord -Window $regPrompt -ForegroundBefore 0 -AppearedSeconds 0) -Answer No }
    # What the installer's own check, certutil -user -store Root, finds after that.
    $cases += Invoke-CertutilCase -Name 'store-after-registry-delete' -CommandLine $check -Show $hide -Answer None -Thumbprint $thumb -AppearSeconds 10
    # And after the registry route, which leaves ProtectedRoots as it was: the same question.
    if ($savedKey.Count -gt 0 -and -not (Test-InUserRoot $thumb).registry) {
        $cases += Invoke-RootKeyWriteCase -Name 'registry-write-after-registry-delete' -Thumbprint $thumb -Saved $savedKey
    }

    if ((Test-InUserRoot $thumb).store) {
        $cases += Invoke-CertutilCase -Name 'cleanup' -CommandLine $delete -Show $normal -Answer Yes -Thumbprint $thumb
    }
    Remove-Item -LiteralPath ("HKCU:\" + $RootCertificatesKey + '\' + $thumb) -Recurse -Force -ErrorAction SilentlyContinue
    $result.cases = $cases
    $result.verdict = Get-RootPromptVerdict $cases
    $removalPrompts = @($cases | Where-Object { $_.Name -like 'delete-*' } | ForEach-Object { @($_.Prompts) })
    $result['removalPromptsNeverShown'] = @($removalPrompts | Where-Object { $null -eq $_.shownAtSeconds }).Count
    $result['removalPrompts'] = $removalPrompts.Count
    $result['leftInStore'] = (Test-InUserRoot $thumb).store
    return $result
}

function Get-InclusionNames {
    $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($InclusionKey, $false)
    if ($null -eq $k) { return @() }
    try { return @($k.GetSubKeyNames()) } finally { $k.Close() }
}

function Add-InclusionEntry([string] $Url, [string] $PublicKey) {
    $name = [guid]::NewGuid().ToString()
    $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($InclusionKey + '\' + $name)
    try {
        if ($null -ne $Url) { $k.SetValue('Url', $Url, [Microsoft.Win32.RegistryValueKind]::String) }
        if ($null -ne $PublicKey) { $k.SetValue('PublicKey', $PublicKey, [Microsoft.Win32.RegistryValueKind]::String) }
    }
    finally { $k.Close() }
    return $name
}

function Add-PlantedEntries([string] $Modulus, [string] $Round) {
    $profileUrl = 'file:///' + ($env:LOCALAPPDATA -replace '\\', '/')
    $other = [System.Security.Cryptography.RSA]::Create(1024)
    try { $otherModulus = [Convert]::ToBase64String($other.ExportParameters($false).Modulus) } finally { $other.Dispose() }
    $plan = @(
        @{ Label = "$Round leaked key, canonical, at the install folder"; Url = "$profileUrl/OutlookAI/Setup/OutlookAI.vsto"; Key = (New-KeyXml -ModulusBase64 $Modulus); Gone = $true },
        @{ Label = "$Round leaked key, wrapped, at an old unzipped copy"; Url = ($profileUrl -replace '/AppData/Local$', '') + '/Downloads/OutlookAI-v2.3.1.111/OutlookAI.vsto'; Key = (New-KeyXml -ModulusBase64 $Modulus -Spelling Wrapped); Gone = $true },
        @{ Label = "$Round leaked key, prefixed, elsewhere"; Url = 'file:///C:/OutlookAI-old/OutlookAI.vsto'; Key = (New-KeyXml -ModulusBase64 $Modulus -Spelling Prefixed); Gone = $true },
        @{ Label = "$Round another key, at the old unzipped copy"; Url = ($profileUrl -replace '/AppData/Local$', '') + '/Downloads/OutlookAI-v2.3.1.111/OutlookAI.vsto'; Key = (New-KeyXml -ModulusBase64 $otherModulus); Gone = $false },
        @{ Label = "$Round no public key at all"; Url = 'file:///C:/OutlookAI-old/Broken.vsto'; Key = $null; Gone = $false })
    $planted = @()
    foreach ($e in $plan) {
        $name = Add-InclusionEntry -Url $e.Url -PublicKey $e.Key
        $planted += [pscustomobject]@{ Label = $e.Label; Name = $name; ExpectGone = $e.Gone; Url = $e.Url }
        Write-Host ("   planted {0}: {1}" -f $name, $e.Label)
    }
    return $planted
}

function Wait-ProcessGone([string[]] $Names, [int] $Seconds) {
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    while ($clock.Elapsed.TotalSeconds -lt $Seconds) {
        $left = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $p = $_.ProcessName; @($Names | Where-Object { $p -like $_ }).Count -gt 0 })
        if ($left.Count -eq 0) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# The installer, or the uninstaller it registered, /VERYSILENT - as the add-in's updater runs the
# installer - with a log.
function Start-InstallerRound([string] $Round, [string] $Log) {
    $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', ('/LOG="' + $Log + '"'))
    if ($Round -eq 'install') { return (Start-Process -FilePath $InstallerPath -ArgumentList $arguments -PassThru -WindowStyle Hidden) }
    $u = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{78AF2871-0CEB-4451-B80D-455552E37C91}_is1' -ErrorAction Stop
    return (Start-Process -FilePath ([string]$u.UninstallString).Trim('"') -ArgumentList $arguments -PassThru -WindowStyle Hidden)
}

function Invoke-InstallerCleanup {
    if (-not $InstallerPath -or -not (Test-Path -LiteralPath $InstallerPath)) { throw "-InstallerPath must name the compiled installer on this guest: '$InstallerPath'." }
    if (-not $InstallerScript -or -not (Test-Path -LiteralPath $InstallerScript)) { throw "-InstallerScript must name the Installer.iss it was compiled from: '$InstallerScript'." }
    if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) { throw 'REFUSING: Outlook is running. The installer would ask it to close; this probe never closes Outlook.' }
    $modulus = Get-CompromisedModulus ([System.IO.File]::ReadAllText($InstallerScript))
    $result = [ordered]@{ facts = $facts; installer = $InstallerPath; rounds = @(); verdict = $null }

    $allPass = $true
    foreach ($round in @('install', 'uninstall')) {
        Write-Host ''
        Write-Host "-- round $round"
        $existing = @(Get-InclusionNames)
        Write-Host "   entries there before planting: $($existing.Count)"
        $planted = @(Add-PlantedEntries -Modulus $modulus -Round $round)
        $log = Join-Path $LogDir ("{0}-{1}.log" -f $round, $stamp)
        $clock = [System.Diagnostics.Stopwatch]::StartNew()
        $p = Start-InstallerRound -Round $round -Log $log
        $ended = $p.WaitForExit(600000)
        # An Inno uninstaller hands over to a copy of itself in %TEMP% and ends; wait for that too.
        $settled = Wait-ProcessGone -Names @('_iu*', 'unins*', 'OutlookAI-v*') -Seconds 300
        $code = $null
        if ($ended) { $code = $p.ExitCode }
        $after = @(Get-InclusionNames)
        $verdict = Get-CleanupVerdict -Planted $planted -NamesAfter $after -KeptBefore $existing
        if ($verdict.Verdict -ne 'PASS') { $allPass = $false }
        $removedLines = @()
        if (Test-Path -LiteralPath $log) { $removedLines = @(Get-Content -LiteralPath $log | Where-Object { $_ -match 'VSTO trust entry|compromised' }) }
        Write-Host ("   {0} exit {1} after {2} s (settled {3}); entries after: {4}; verdict {5}" -f $round, $code, [Math]::Round($clock.Elapsed.TotalSeconds, 1), $settled, $after.Count, $verdict.Verdict)
        foreach ($pr in $verdict.Problems) { Write-Host "   PROBLEM: $pr" }
        foreach ($l in $removedLines) { Write-Host "   log: $l" }
        $result.rounds += [pscustomobject]@{
            name = $round; exitCode = $code; ended = $ended; settled = $settled; seconds = [Math]::Round($clock.Elapsed.TotalSeconds, 1)
            existingBefore = $existing; planted = $planted; namesAfter = $after; verdict = $verdict.Verdict; problems = $verdict.Problems; logLines = $removedLines; log = $log
        }
    }
    $result.verdict = $(if ($allPass) { 'PASS' } else { 'FAIL' })
    return $result
}

$outcome = $null
switch ($Scenario) {
    'RootStorePrompt' { $outcome = Invoke-RootStorePrompt }
    'InstallerCleanup' { $outcome = Invoke-InstallerCleanup }
}
$outcome['finishedUtc'] = [DateTime]::UtcNow.ToString('o')
[System.IO.File]::WriteAllText($jsonPath, ($outcome | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($false)))
Write-Host ''
if ($Scenario -eq 'RootStorePrompt') { Write-Host ("removal prompts: {0}, of which never shown within {1} s: {2}" -f $outcome['removalPrompts'], $VisibleWithinSeconds, $outcome['removalPromptsNeverShown']) }
Write-Host "VERDICT: $($outcome.verdict)"
Write-Host "JSON: $jsonPath"
if ($outcome.verdict -like 'INCONCLUSIVE*' -or $outcome.verdict -eq 'FAIL') { exit 1 }
exit 0
