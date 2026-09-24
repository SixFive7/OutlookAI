#Requires -Version 5.1
<#
.SYNOPSIS
    Restarts a testbed guest WITHOUT FORCING ANYTHING: Outlook is quit gracefully first, exactly
    as mailbox-safety rule 7 describes, then Windows restarts unforced, and the script waits until
    the guest is usable again - heartbeat OK and the console session Active.

.DESCRIPTION
    RUN ON THE HOST. Windows PowerShell 5.1 or PowerShell 7. Everything inside the guest runs
    over PowerShell Direct, which needs no network.

    WHY THIS EXISTS. The guest restarts on record before 2026-09-24 - the earlier rounds' scratch
    scripts under .work/ - were `shutdown /r` with a timeout: /t 5 in most, 10, 15 or 20 in others,
    /t 0 in a few. Microsoft documents, on the shutdown command's own page, that "if the timeout
    period is greater than 0, the /f parameter is implied" - and /f is "forces running applications
    to close without warning users". So every one of those with a timeout FORCE-CLOSED whatever
    Outlook was still running, PSTs open, which is as close to `taskkill OUTLOOK.EXE` as makes no
    difference. Mailbox-safety rule 7
    forbids the kill; a forced close skips the same shutdown path the kill skips. Nothing broke
    that anyone noticed - which is not the same as nothing having broken.
    https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/shutdown

    WHAT IT DOES, in order. Each phase must pass before the next one runs, and nothing is ever
    forced - a phase that cannot complete gracefully STOPS the script and says why.

      1. PREFLIGHT (session 0, read-only). The VM exists and is Running, its heartbeat is OK, the
         guest's current boot time (so a restart can be PROVED rather than inferred from a
         heartbeat blip), whether OUTLOOK.EXE is running, at which integrity level, and whether a
         known COM client of Outlook is running. The project README says why that last one
         refuses: Application.Quit() with agent sessions attached makes Outlook "tear down and
         park indefinitely, waiting for the attached clients".

      2. QUIT OUTLOOK GRACEFULLY, in session 1 - only if it is running. PowerShell Direct lands in
         session 0, where Outlook's object is not reachable, so this runs as a one-shot
         scheduled task in the interactive session, at the SAME integrity level as the running
         Outlook (an elevated caller cannot see a non-elevated Outlook in the Running Object
         Table, and the reverse - the "elevation mismatch breaks COM attach" rule). The task:
           * REFUSES if the add-in installer's OutlookAISetup mutex is held (v3 S7);
           * REFUSES if Outlook shows ANY visible dialog - a modal wait nobody on an unattended
             guest answers, and Quit() behind one parks. It names the dialog. The one exception,
             only with -CancelLogonPrompt: the "Internet Email - <account>" logon prompt the tier
             profile's POP3 account raises at every start (no stored password, no sink), which
             is cancelled - that sends nothing and stores nothing. A security prompt, such as the
             Object Model Guard's, is never clicked from here;
           * ATTACHES to the running instance - through the Running Object Table first
             (Marshal.GetActiveObject), which cannot start anything. Office registers there only
             once its window has LOST focus (Microsoft KB 238610), and on an unattended guest
             nothing may ever take focus from it: measured 2026-09-24, an Outlook nine minutes up
             answered MK_E_UNAVAILABLE for a full minute. So after 20 s it asks COM for the
             RUNNING instance's class object (Activator.CreateInstance on the ProgID - how the
             product's own server attaches). Outlook is single-instance, so with OUTLOOK.EXE
             running that binds to it and cannot start a second one; the process is checked
             immediately before, and an Outlook that has gone is "nothing to quit";
           * REFUSES if ANY Inspector is open. Rule 7 names unsent compose windows; this is
             stricter on purpose, because an Inspector with unsaved edits makes Quit() either
             discard them or raise a save prompt, and a prompt on an unattended guest is a hang;
           * REFUSES if the default store's Outbox cannot be read, or ANY store's Outbox holds
             an item (rule 7: "the Outbox is empty");
           * RELEASES every COM reference it took - folders, item collections, stores, the
             NameSpace - and runs the finalizers, BEFORE calling Quit() (rule 7: "quitting
             while refs are held zombifies the process"). Only the Application reference is
             still held when Quit() is called, because it is the object Quit() is called on;
             it is released immediately after, and the task process then exits, which drops
             anything PowerShell's COM adapter might still have held.
         Then the host WAITS for OUTLOOK.EXE to leave the process list, up to
         -OutlookExitTimeoutSeconds. IT NEVER KILLS IT. If Outlook is still there at the
         deadline the script stops, names the candidate processes still holding it, and does
         not restart the guest - a restart would force exactly what this script exists to avoid.

      3. RESTART, UNFORCED: `shutdown.exe /r /t 0 /d p:4:1 /c <reason>`. A timeout of 0 is the
         only one that does not imply /f, and /f is not passed. If an application in the guest
         refuses to close, Windows holds the restart for it - and this script reports "the
         restart did not happen" rather than forcing it.

      4. WAIT UNTIL USABLE. The heartbeat integration service leaves OK and comes back to OK;
         PowerShell Direct answers; the guest reports a LastBootUpTime LATER than the one read
         in phase 1 (the proof that a restart happened, rather than a heartbeat blip); and
         `query session` shows the console session Active with the autologon user in it. That
         last one is what everything in session 1 depends on - Register-InteractiveTask.ps1 has
         nowhere to land until it is true.

    EXIT CODES. 0 restarted and usable. 2 refused before anything changed (Outlook state - an
    Inspector, a dialog, an Outbox item, the installer mutex - a COM client attached, the VM not
    ready). 3 Outlook was asked to quit and did not exit - the guest was NOT restarted. 4 the
    restart was requested and did not happen. 5 the guest restarted but did not come back to a
    usable state within the deadline. 1 anything unexpected.

    WHAT IT NEVER DOES. It never kills a process, never passes /f, never uses Stop-VM,
    Restart-VM or a Hyper-V reset (each is a power operation, not a graceful OS restart), never
    creates an Outlook COM object, never touches a mail item, and never touches the host's own
    Outlook: it runs no Outlook, COM or MAPI code on the host at all.

    HOLD A LEASE WHILE YOU USE IT. A restart takes minutes, and the idle-saver saves a testbed VM
    that holds no live lease (Testbed/README.md section 5b). This script warns when there is no
    lease; it does not take one, because the lease belongs to the work around the restart.

.PARAMETER VMName
    MANDATORY. The guest to restart. There is no default, for the reason there is none anywhere
    in Testbed/host/: three machines coexist and a default that silently picks one of them is the
    exact shape of mistake this testbed keeps making.

.PARAMETER Execute
    Do it. Without it the script runs the read-only preflight, prints what it would do, and
    changes nothing - the convention every writing script in Testbed/ follows.

.PARAMETER Reason
    Free text passed to shutdown.exe /c and printed in the log.

.PARAMETER OutlookExitTimeoutSeconds
    How long to wait for OUTLOOK.EXE to exit after Quit(). Outlook flushes its PSTs on the way
    out; the corpus store is ~400 MB.

.PARAMETER CancelLogonPrompt
    Press CANCEL on Outlook's "Internet Email - <account>" logon prompts before quitting - every one:
    Outlook raises one per POP3 account, the next only after the previous is answered (measured:
    'tier', then 'identity' three seconds later), so it cancels in rounds until a round finds none.
    Without it such a prompt - like any other visible Outlook dialog - is a refusal. The tier
    profile's POP3 accounts raise them at every start on the guests (no stored password, no sink), so
    a restart with that profile running needs this. Cancelling sends nothing and stores nothing. Nothing else
    is ever clicked: a security prompt is a refusal whatever this says.

.PARAMETER RestartTimeoutMinutes
    How long each wait in phase 4 may take.

.PARAMETER RepoRoot
    Repository root, for the credential loader.

.PARAMETER LogPath
    Optional transcript file on the host.

.PARAMETER SelfTest
    Pure. Checks the parsers and the decision table this script's phases rest on - the
    `query session` reader, the quit task's exit-code meanings, the COM-client list and the
    integrity-level to run-level mapping - against synthetic inputs. Needs no VM and no
    Hyper-V; ignores every other switch. Exit 0 only if every case passes.

.EXAMPLE
    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Minutes 45 -Reason 'restart'
    Testbed/host/Restart-Guest.ps1 -VMName OutlookAI-Indexed
    Testbed/host/Restart-Guest.ps1 -VMName OutlookAI-Indexed -Execute -Reason 'after scope change'
    Testbed/host/Restart-Guest.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $VMName,
    [switch] $Execute,
    [string] $Reason = 'OutlookAI testbed: graceful restart (Testbed/host/Restart-Guest.ps1)',
    [int]    $OutlookExitTimeoutSeconds = 600,
    [int]    $RestartTimeoutMinutes = 15,
    [string] $RepoRoot,
    [string] $LogPath,
    [switch] $CancelLogonPrompt,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# Resolved here rather than as a parameter default: Windows PowerShell 5.1 run with -File leaves
# $PSScriptRoot empty while it evaluates parameter defaults (measured 2026-09-24).
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }

# Processes that hold Outlook's object model open while they run. Quit() with any of them
# attached parks Outlook indefinitely (README.md, "Outlook Lifetime and the Tray Icon"), so their
# presence is a refusal, not a warning. testhost / vstest.console are the live suite.
$script:ComClientProcessNames = @('OutlookAI.McpServer', 'OutlookAI.RemediationTools', 'testhost', 'testhost.x86', 'vstest.console')

# The quit task's exit codes, in one table so the host and the guest cannot disagree.
$script:QuitCodes = @{
    0  = 'QUIT-CALLED'
    10 = 'ATTACH-FAILED'
    11 = 'REFUSED-INSPECTOR-OPEN'
    12 = 'REFUSED-OUTBOX-NOT-EMPTY'
    13 = 'REFUSED-OUTBOX-UNREADABLE'
    14 = 'REFUSED-INSTALLER-MUTEX'
    15 = 'NOT-RUNNING'
    16 = 'REFUSED-DIALOG-OPEN'
}

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    if ($LogPath) {
        try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
    }
}

# `query session` output -> the console row, or $null. Pure, so -SelfTest can pin it.
# Active console with a user looks like: " console           vmadmin                   1  Active"
# (a leading '>' marks the caller's own session). Before autologon completes the console row
# has no user and a state of Conn - which is NOT usable, and must not match.
function Get-ConsoleSessionState {
    param([string] $QueryText, [string] $ExpectedUser)
    foreach ($line in ($QueryText -split "`r?`n")) {
        $m = [regex]::Match($line, '^\s*>?console\s+(\S+)\s+(\d+)\s+(\S+)')
        if (-not $m.Success) { continue }
        $user = $m.Groups[1].Value
        $state = $m.Groups[3].Value
        $usable = ($state -eq 'Active')
        if ($ExpectedUser -and $user -ne $ExpectedUser) { $usable = $false }
        return [pscustomobject]@{ User = $user; SessionId = [int]$m.Groups[2].Value; State = $state; Usable = $usable }
    }
    return $null
}

# The token elevation of the running Outlook decides the run level of the quit task: the Running
# Object Table does not show a non-elevated server to an elevated caller or the reverse.
# 1 = elevated -> Highest; 0 = not elevated -> Limited; anything else -> try both, Highest first.
function Get-QuitRunLevels {
    param([int] $Elevation)
    if ($Elevation -eq 1) { return @('Highest') }
    if ($Elevation -eq 0) { return @('Limited') }
    return @('Highest', 'Limited')
}

function Get-QuitMeaning {
    param([int] $Code)
    if ($script:QuitCodes.ContainsKey($Code)) { return $script:QuitCodes[$Code] }
    return "UNEXPECTED($Code)"
}

# ---------------------------------------------------------------------------------------------
# The session-1 quit script. Sent as text, written into a job directory in the guest, run by a
# one-shot interactive scheduled task. Windows PowerShell 5.1. Exit codes: $script:QuitCodes.
# ---------------------------------------------------------------------------------------------
$script:QuitScript = @'
$ErrorActionPreference = 'Stop'
$Marshal = [System.Runtime.InteropServices.Marshal]
function Out-Line([string] $m) { Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
function Release($o) { if ($null -ne $o) { try { [void]$Marshal::ReleaseComObject($o) } catch { } } }

if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -eq 0) { Out-Line 'OUTLOOK.EXE is not running - nothing to quit'; exit 15 }

# A visible DIALOG is a modal wait nobody on an unattended guest will answer, and Quit() behind one
# parks. So every visible dialog Outlook owns is named and refused - with one exception, and only
# when the caller asked for it: the "Internet Email - <account>" logon prompt, which the tier
# profile's POP3 account raises at every start (no stored password, no sink - runbook 1.4).
# Cancelling it sends nothing and stores nothing. Security prompts are never answered from here.
if (-not ('OaiRestartDialogs' -as [type])) {
    Add-Type -TypeDefinition @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class OaiRestartDialogs {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint a, bool i, int p);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr p, uint a, out IntPtr t);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr t, int c, out int i, int l, out int r);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    // 1 elevated, 0 not, -1 unreadable (TokenElevation = 20).
    public static int Elevation(int pid) { IntPtr p = OpenProcess(0x1000, false, pid); if (p == IntPtr.Zero) return -1; try { IntPtr t; if (!OpenProcessToken(p, 8, out t)) return -1; try { int e, r; if (!GetTokenInformation(t, 20, out e, 4, out r)) return -1; return e != 0 ? 1 : 0; } finally { CloseHandle(t); } } finally { CloseHandle(p); } }
    // Visible dialogs (#32770) of the process, as their titles. With cancelLogon, presses Cancel on
    // each "Internet Email - " logon prompt and reports it as "CANCELLED: <title>".
    public static List<string> Dialogs(uint pid, bool cancelLogon) {
        List<string> found = new List<string>();
        EnumWindows(delegate (IntPtr h, IntPtr l) {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != pid || !IsWindowVisible(h)) return true;
            StringBuilder c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() != "#32770") return true;
            StringBuilder t = new StringBuilder(512); GetWindowText(h, t, 512);
            string title = t.ToString();
            if (cancelLogon && title.StartsWith("Internet Email - ")) {
                IntPtr cancel = IntPtr.Zero;
                EnumChildWindows(h, delegate (IntPtr ch, IntPtr l2) {
                    StringBuilder cc = new StringBuilder(64); GetClassName(ch, cc, 64);
                    StringBuilder ct = new StringBuilder(64); GetWindowText(ch, ct, 64);
                    if (cc.ToString() == "Button" && ct.ToString().Replace("&", "") == "Cancel") { cancel = ch; return false; }
                    return true;
                }, IntPtr.Zero);
                if (cancel != IntPtr.Zero) { SendMessage(cancel, 0x00F5, IntPtr.Zero, IntPtr.Zero); found.Add("CANCELLED: " + title); return true; }
            }
            found.Add(title);
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@
}
$outlookPid = [uint32](@(Get-Process -Name OUTLOOK)[0].Id)
# One logon prompt per POP3 account, and Outlook raises the next only once the previous one is
# answered - measured 2026-09-24: cancelling 'Internet Email - tier' was followed within three
# seconds by 'Internet Email - identity'. So cancelling repeats, with a settle between rounds,
# until a round finds no logon prompt; a bounded number of rounds, then whatever is left refuses.
$dialogs = @([OaiRestartDialogs]::Dialogs($outlookPid, $false))
for ($round = 1; [bool]$CancelLogonPrompt -and $round -le 8; $round++) {
    $dialogs = @([OaiRestartDialogs]::Dialogs($outlookPid, $true))
    $cancelled = @($dialogs | Where-Object { $_.StartsWith('CANCELLED: ') })
    foreach ($d in $cancelled) { Out-Line ("pressed Cancel on the logon prompt '" + $d.Substring(11) + "' (-CancelLogonPrompt)") }
    Start-Sleep -Seconds 4
    $dialogs = @([OaiRestartDialogs]::Dialogs($outlookPid, $false))
    if (@($dialogs | Where-Object { $_.StartsWith('Internet Email - ') }).Count -eq 0) { break }
}
if ($dialogs.Count -gt 0) {
    Out-Line ("REFUSING: Outlook has a dialog open: '" + ($dialogs -join "', '") + "'. Quit() behind a modal dialog parks.")
    Out-Line 'Answer it in the guest. An "Internet Email - <account>" logon prompt can be cancelled with -CancelLogonPrompt;'
    Out-Line 'anything else - a security prompt above all - is for a person to read, not for this script to click.'
    exit 16
}

foreach ($name in @('OutlookAISetup', 'Global\OutlookAISetup')) {
    $m = $null
    $held = $false
    # TryOpenExisting throws only when the mutex EXISTS and cannot be opened - which is held.
    try { $held = [System.Threading.Mutex]::TryOpenExisting($name, [ref] $m) } catch { $held = $true }
    if ($m) { $m.Dispose() }
    if ($held) { Out-Line "REFUSING: the add-in installer mutex '$name' is held - an add-in update is running (v3 S7)"; exit 14 }
}

# ATTACHING. The Running Object Table first - it cannot start anything. But Office registers its
# Application object there only once its window has LOST focus (Microsoft KB 238610), and on an
# unattended guest nothing may ever take focus from it - measured 2026-09-24: an Outlook started by
# Start-OutlookUnelevated.ps1 answered GetActiveObject with MK_E_UNAVAILABLE for a full minute,
# nine minutes after its window came up. So after 20 s the running instance is asked for through
# COM's class object instead - Activator.CreateInstance on the ProgID, which is exactly how the
# product's own server attaches (OutlookComSession.Connect). Outlook is a single-instance server: with
# OUTLOOK.EXE running, that call binds to the instance that is running and cannot start a second one;
# the process is checked immediately before it, and a vanished Outlook is "nothing to quit".
$app = $null
$attachUntil = (Get-Date).AddSeconds(20)
while ($null -eq $app) {
    try { $app = $Marshal::GetActiveObject('Outlook.Application') }
    catch {
        if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -eq 0) { Out-Line 'OUTLOOK.EXE exited while waiting to attach - nothing to quit'; exit 15 }
        if ((Get-Date) -ge $attachUntil) { break }
        Start-Sleep -Seconds 3
    }
}
if ($null -eq $app) {
    Out-Line 'not in the Running Object Table after 20 s (KB 238610) - attaching through the class object of the RUNNING instance'
    # Only at Outlook's OWN integrity level: across levels COM would not reach the running
    # instance's class object - it would launch a second OUTLOOK.EXE at this task's level.
    $mine = 0
    if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { $mine = 1 }
    $theirs = [OaiRestartDialogs]::Elevation([int]$outlookPid)
    if ($theirs -ne $mine) { Out-Line "Outlook's elevation reads $theirs and this task's is $mine - the class object would START a second Outlook here, so no fallback"; exit 10 }
    if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -eq 0) { Out-Line 'OUTLOOK.EXE is not running - nothing to quit'; exit 15 }
    try { $app = [Activator]::CreateInstance([Type]::GetTypeFromProgID('Outlook.Application', $true)) }
    catch { Out-Line ("could not attach to the running Outlook: " + $_.Exception.Message); exit 10 }
}
Out-Line ("attached: Outlook " + $app.Version)

$code = 0
$ns = $null
try {
    $inspectors = $app.Inspectors
    $inspectorCount = $inspectors.Count
    Release $inspectors; $inspectors = $null
    Out-Line "open inspectors: $inspectorCount"
    if ($inspectorCount -gt 0) {
        Out-Line 'REFUSING: an Inspector is open. It may be an unsent compose window (rule 7), and any Inspector with'
        Out-Line 'unsaved edits turns Quit() into a discard or a save prompt. Close it in the guest, then run this again.'
        $code = 11
    }

    if ($code -eq 0) {
        $ns = $app.GetNamespace('MAPI')
        # The DEFAULT store's Outbox must be readable: fail closed.
        try {
            $outbox = $ns.GetDefaultFolder(4)
            $items = $outbox.Items
            $n = $items.Count
            Release $items; Release $outbox; $items = $null; $outbox = $null
            Out-Line "default store Outbox: $n item(s)"
            if ($n -gt 0) { $code = 12 }
        }
        catch { Out-Line ("REFUSING: the default store's Outbox could not be read: " + $_.Exception.Message); $code = 13 }
    }

    if ($code -eq 0) {
        # Every other store's Outbox, where it has one. A data file that is nobody's delivery store
        # may have no Outbox at all; that is reported and is not a refusal.
        $stores = $ns.Stores
        $storeCount = $stores.Count
        for ($i = 1; $i -le $storeCount; $i++) {
            $store = $stores.Item($i)
            $label = $store.DisplayName
            try {
                $outbox = $store.GetDefaultFolder(4)
                $items = $outbox.Items
                $n = $items.Count
                Release $items; Release $outbox; $items = $null; $outbox = $null
                Out-Line "store '$label' Outbox: $n item(s)"
                if ($n -gt 0) { $code = 12 }
            }
            catch { Out-Line "store '$label' has no readable Outbox ($($_.Exception.Message.Trim())) - nothing can be queued in it" }
            Release $store; $store = $null
        }
        Release $stores; $stores = $null
        if ($code -eq 12) { Out-Line 'REFUSING: an Outbox holds an item (rule 7: Quit only with the Outbox empty). Nothing was quit.' }
    }
}
catch {
    Out-Line ("unexpected error while checking: " + $_.Exception.Message)
    $code = 1
}
finally {
    Release $ns; $ns = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($code -ne 0) { Release $app; $app = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); exit $code }

# Every reference but the Application itself is released and finalised. Quit, then drop it.
Out-Line 'all references released except Application; calling Quit()'
$app.Quit()
Release $app; $app = $null
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
Out-Line 'Quit() returned; Application released'
exit 0
'@

# ---------------------------------------------------------------------------------------------
# -SelfTest: pure.
# ---------------------------------------------------------------------------------------------
# The commands a graceful restart must never issue, and the arguments shutdown.exe must never get.
# Checked on the PARSED script (command names from the AST), not on its text, so the prose above
# that names them in order to forbid them does not count as a use.
function Get-ForbiddenCommandUse {
    param([Parameter(Mandatory = $true)] [string] $ScriptText)
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref] $null, [ref] $errors)
    $found = @()
    $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    foreach ($c in $commands) {
        $name = $c.GetCommandName()
        if (-not $name) { continue }
        $leaf = [IO.Path]::GetFileNameWithoutExtension($name).ToLowerInvariant()
        if (@('stop-vm', 'restart-vm', 'stop-process', 'taskkill', 'kill', 'spps') -contains $leaf) { $found += "$name (line $($c.Extent.StartLineNumber))" }
        if ($leaf -eq 'shutdown') {
            # Microsoft: "if the timeout period is greater than 0, the /f parameter is implied".
            $words = @($c.CommandElements | ForEach-Object { $_.Extent.Text.Trim('"', "'").ToLowerInvariant() })
            for ($i = 1; $i -lt $words.Count; $i++) {
                if ($words[$i] -eq '/f' -or $words[$i] -eq '-f') { $found += "shutdown with $($words[$i]) (line $($c.Extent.StartLineNumber))" }
                if (($words[$i] -eq '/t' -or $words[$i] -eq '-t') -and ($i + 1 -lt $words.Count) -and ($words[$i + 1] -ne '0')) {
                    $found += "shutdown /t $($words[$i + 1]) implies /f (line $($c.Extent.StartLineNumber))"
                }
            }
        }
    }
    return @($found)
}

function Invoke-SelfTest {
    function Check([string] $name, $got, $want) {
        if ("$got" -eq "$want") { $script:stPass++; Write-Host "  PASS  $name" }
        else { $script:stFail++; Write-Host "  FAIL  $name - got '$got', want '$want'" }
    }
    $script:stPass = 0
    $script:stFail = 0

    $active = @"
 SESSIONNAME               USERNAME                 ID  STATE   TYPE        DEVICE
>services                                            0  Disc
 console                   vmadmin                   1  Active
 31c5ce94259d4006a9e4                            65536  Listen
"@
    $s = Get-ConsoleSessionState -QueryText $active -ExpectedUser 'vmadmin'
    Check 'query session: console Active with the autologon user is usable' $s.Usable $true
    Check 'query session: session id read' $s.SessionId 1
    $conn = @"
 SESSIONNAME               USERNAME                 ID  STATE   TYPE        DEVICE
>services                                            0  Disc
 console                                             1  Conn
"@
    $c = Get-ConsoleSessionState -QueryText $conn -ExpectedUser 'vmadmin'
    Check 'query session: console with no user (autologon not done) is NOT usable' ([bool]($c -and $c.Usable)) $false
    $wrongUser = Get-ConsoleSessionState -QueryText $active -ExpectedUser 'someoneelse'
    Check 'query session: console Active as another user is NOT usable' $wrongUser.Usable $false
    $mine = Get-ConsoleSessionState -QueryText ($active -replace ' console', '>console') -ExpectedUser 'vmadmin'
    Check 'query session: the caller''s own console row (leading >) is read' $mine.Usable $true
    Check 'query session: no console row at all -> null' ($null -eq (Get-ConsoleSessionState -QueryText ">services  0  Disc" -ExpectedUser 'vmadmin')) $true

    Check 'run level: elevated Outlook -> Highest only' ((Get-QuitRunLevels -Elevation 1) -join ',') 'Highest'
    Check 'run level: non-elevated Outlook -> Limited only' ((Get-QuitRunLevels -Elevation 0) -join ',') 'Limited'
    Check 'run level: unknown elevation -> Highest then Limited' ((Get-QuitRunLevels -Elevation -1) -join ',') 'Highest,Limited'

    Check 'quit code 0 means Quit() was called' (Get-QuitMeaning 0) 'QUIT-CALLED'
    Check 'quit code 11 is a refusal for an open Inspector' (Get-QuitMeaning 11) 'REFUSED-INSPECTOR-OPEN'
    Check 'quit code 12 is a refusal for a non-empty Outbox' (Get-QuitMeaning 12) 'REFUSED-OUTBOX-NOT-EMPTY'
    Check 'quit code 13 is a refusal for an unreadable default Outbox' (Get-QuitMeaning 13) 'REFUSED-OUTBOX-UNREADABLE'
    Check 'quit code 16 is a refusal for an open dialog' (Get-QuitMeaning 16) 'REFUSED-DIALOG-OPEN'
    Check 'an unknown quit code is reported as unexpected' (Get-QuitMeaning 99) 'UNEXPECTED(99)'

    foreach ($n in @('OutlookAI.McpServer', 'OutlookAI.RemediationTools', 'testhost')) {
        Check "COM client list names $n" ($script:ComClientProcessNames -contains $n) $true
    }

    # The quit script must parse, must attach rather than create, must release before Quit(),
    # and must contain no kill. Checked on its text, because that text is what runs in the guest.
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($script:QuitScript, [ref] $null, [ref] $errors)
    Check 'the quit script parses' (@($errors).Count) 0
    Check 'the quit script attaches with GetActiveObject' ($script:QuitScript -match 'GetActiveObject\(''Outlook\.Application''\)') $true
    $rotAt = $script:QuitScript.IndexOf("GetActiveObject('Outlook.Application')")
    $classAt = $script:QuitScript.IndexOf('[Activator]::CreateInstance(')
    Check 'the quit script tries the Running Object Table before the class object' (($rotAt -gt 0) -and ($classAt -gt $rotAt)) $true
    $quitLines = $script:QuitScript -split "`n"
    $classLine = [Array]::FindIndex($quitLines, [Predicate[string]] { param($l) $l -match '\[Activator\]::CreateInstance\(' })
    Check 'the class-object attach is preceded by a check that OUTLOOK.EXE is running' (($classLine -gt 0) -and ($quitLines[$classLine - 1] -match 'Get-Process -Name OUTLOOK')) $true
    $guardAt = $script:QuitScript.IndexOf('if ($theirs -ne $mine)')
    Check 'the class-object attach happens only when Outlook''s elevation equals the task''s' (($guardAt -gt 0) -and ($guardAt -lt $classAt)) $true
    Check 'the quit script never uses New-Object -ComObject' ($script:QuitScript -match 'New-Object\s+-ComObject') $false
    $releaseAt = $script:QuitScript.IndexOf('Release $ns; $ns = $null')
    $quitAt = $script:QuitScript.IndexOf('$app.Quit()')
    Check 'the quit script releases the NameSpace before Quit()' (($releaseAt -gt 0) -and ($releaseAt -lt $quitAt)) $true
    Check 'the quit script issues no kill command' ((Get-ForbiddenCommandUse -ScriptText $script:QuitScript) -join '; ') ''
    Check 'the quit script calls no .Kill() method' ($script:QuitScript -match '\.Kill\(') $false
    Check 'this file issues no forced shutdown, Stop-VM, Restart-VM, Stop-Process or taskkill' ((Get-ForbiddenCommandUse -ScriptText ([IO.File]::ReadAllText($PSCommandPath))) -join '; ') ''

    # Every command this file calls must exist - a function defined here, or a command this
    # PowerShell knows (Hyper-V's cmdlets are allowed by name: the self-test must also run on a
    # machine without that module). A call to a function that is not there is otherwise found only
    # by running the one path that makes it.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $defined = @{}
    foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fd.Name] = $true }
    # ...and the functions of the file it dot-sources.
    $lease = Join-Path $PSScriptRoot 'TestbedLeasePath.ps1'
    if (Test-Path -LiteralPath $lease) { foreach ($fd in ([System.Management.Automation.Language.Parser]::ParseFile($lease, [ref] $null, [ref] $null)).FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fd.Name] = $true } }
    $external = @('Get-VM', 'Get-VMIntegrationService')
    $unknown = @()
    foreach ($cmd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $name = $cmd.GetCommandName()
        if (-not $name -or $defined.ContainsKey($name) -or $external -contains $name) { continue }
        if (Get-Command -Name $name -ErrorAction SilentlyContinue) { continue }
        $unknown += "$name (line $($cmd.Extent.StartLineNumber))"
    }
    Check 'every command this file calls exists' (($unknown | Select-Object -Unique) -join ', ') ''
    # The detector itself, so a clean result above means something.
    Check 'detector: a shutdown with /f is caught' (@(Get-ForbiddenCommandUse -ScriptText 'shutdown.exe /r /f /t 0').Count -gt 0) $true
    Check 'detector: a shutdown with /t 5 (implies /f) is caught' (@(Get-ForbiddenCommandUse -ScriptText 'shutdown /r /t 5').Count -gt 0) $true
    Check 'detector: Stop-Process is caught' (@(Get-ForbiddenCommandUse -ScriptText 'Stop-Process -Name OUTLOOK').Count -gt 0) $true
    Check 'detector: taskkill is caught' (@(Get-ForbiddenCommandUse -ScriptText 'taskkill /im outlook.exe').Count -gt 0) $true
    Check 'detector: shutdown /r /t 0 without /f passes' (@(Get-ForbiddenCommandUse -ScriptText '& shutdown.exe /r /t 0 /d p:4:1 /c $why').Count) 0

    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:stPass, $script:stFail)
    if ($script:stFail -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

if (-not $VMName) { throw '-VMName is mandatory (Testbed/README.md section 4a): name the guest you mean.' }
if ($VMName -notmatch '^[A-Za-z0-9._-]{1,64}$') { throw "VM name '$VMName' is not a plain name." }
if ($OutlookExitTimeoutSeconds -lt 30) { throw '-OutlookExitTimeoutSeconds below 30 cannot cover a PST flush.' }

# ---------------------------------------------------------------------------------------------
# Plumbing
# ---------------------------------------------------------------------------------------------
$cred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$expectedUser = $cred.UserName
if ($expectedUser -match '\\') { $expectedUser = $expectedUser.Split('\')[-1] }

function Get-Heartbeat {
    $hb = Get-VMIntegrationService -VMName $VMName -Name 'Heartbeat' -ErrorAction SilentlyContinue
    if (-not $hb) { return 'absent' }
    return [string]$hb.PrimaryStatusDescription
}

function Invoke-InGuest([scriptblock] $Block, [object[]] $ArgumentList = @()) {
    Invoke-Command -VMName $VMName -Credential $cred -ScriptBlock $Block -ArgumentList $ArgumentList -ErrorAction Stop
}

function Read-GuestState {
    Invoke-InGuest -Block {
        param($clientNames)
        $os = Get-CimInstance Win32_OperatingSystem
        $outlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        $elevation = -2
        if ($outlook.Count -gt 0) {
            if (-not ('OaiRestartTokenProbe' -as [type])) {
                Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OaiRestartTokenProbe {
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr token, int infoClass, out int info, int length, out int returned);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    // 1 elevated, 0 not elevated, -1 could not be read. TokenElevation = 20.
    public static int Elevation(int pid) {
        IntPtr process = OpenProcess(0x1000, false, pid);
        if (process == IntPtr.Zero) return -1;
        try {
            IntPtr token;
            if (!OpenProcessToken(process, 0x0008, out token)) return -1;
            try {
                int elevated; int returned;
                if (!GetTokenInformation(token, 20, out elevated, 4, out returned)) return -1;
                return elevated != 0 ? 1 : 0;
            } finally { CloseHandle(token); }
        } finally { CloseHandle(process); }
    }
}
'@
            }
            $elevation = [OaiRestartTokenProbe]::Elevation($outlook[0].Id)
        }
        $clients = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $clientNames -contains $_.ProcessName } | ForEach-Object { "$($_.ProcessName)($($_.Id))" })
        [pscustomobject]@{
            LastBoot       = $os.LastBootUpTime.ToUniversalTime().ToString('o')
            OutlookPids    = @($outlook | ForEach-Object { $_.Id })
            OutlookSession = @($outlook | ForEach-Object { $_.SessionId })
            Elevation      = $elevation
            Clients        = $clients
            Sessions       = ((query session 2>&1) -join "`n")
        }
    } -ArgumentList @(, $script:ComClientProcessNames)
}

# Runs the quit script as a one-shot scheduled task in the interactive session at the given run
# level, and returns its exit code and output. The task is removed afterwards.
function Invoke-QuitTask([string] $RunLevel) {
    Invoke-InGuest -Block {
        param($scriptText, $runLevel, $timeoutSeconds)
        $ErrorActionPreference = 'Stop'
        $jobDir = Join-Path 'C:\OutlookAI-Q5\jobs' ('graceful-quit-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path $jobDir | Out-Null
        $work = Join-Path $jobDir 'quit.ps1'
        $cmd = Join-Path $jobDir 'cmd.ps1'
        $out = Join-Path $jobDir 'out.txt'
        $exitFile = Join-Path $jobDir 'exit.txt'
        Set-Content -LiteralPath $work -Value $scriptText -Encoding UTF8
        # exit.txt is written LAST, so its existence means finished (Register-InteractiveTask.ps1's contract).
        $wrapper = @"
`$code = 1
try {
    & '$work' *>&1 | Out-File -FilePath '$out' -Encoding utf8 -Append
    `$code = `$LASTEXITCODE
}
catch { "WRAPPER CAUGHT: `$(`$_.Exception.Message)" | Out-File -FilePath '$out' -Encoding utf8 -Append }
finally { "`$code" | Out-File -FilePath '$exitFile' -Encoding utf8 }
"@
        Set-Content -LiteralPath $cmd -Value $wrapper -Encoding UTF8
        $taskName = 'OutlookAI-GracefulQuit'
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false }
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$cmd`""
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $runLevel
        $settings = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::FromSeconds($timeoutSeconds + 120)) -MultipleInstances IgnoreNew
        Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings | Out-Null
        Start-ScheduledTask -TaskName $taskName
        $deadline = (Get-Date).AddSeconds($timeoutSeconds)
        while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $exitFile)) { Start-Sleep -Seconds 2 }
        $code = -1
        if (Test-Path -LiteralPath $exitFile) { $code = [int]((Get-Content -LiteralPath $exitFile -Raw).Trim()) }
        $text = ''
        if (Test-Path -LiteralPath $out) { $text = (Get-Content -LiteralPath $out) -join "`n" }
        # Removed only once it has finished. A task still running at the deadline is left for its
        # own ExecutionTimeLimit to end, and reported - never stopped from here.
        if ($code -ne -1) { try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false } catch { } }
        else { $text += "`n(no exit.txt after $timeoutSeconds s; task '$taskName' left to its own time limit)" }
        [pscustomobject]@{ Code = $code; Output = $text; JobDir = $jobDir }
    } -ArgumentList @(('$CancelLogonPrompt = $' + ([bool]$CancelLogonPrompt).ToString().ToLowerInvariant() + "`n" + $script:QuitScript), $RunLevel, 300)
}

# ---------------------------------------------------------------------------------------------
# Phase 1: preflight
# ---------------------------------------------------------------------------------------------
$t0 = Get-Date
Say "== Restart-Guest: $VMName ($(if ($Execute) { 'EXECUTE' } else { 'DRY RUN' })) =="

$vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
if (-not $vm) { Say "REFUSING: no VM named '$VMName' on this host."; exit 2 }
if ($vm.State -ne 'Running') { Say "REFUSING: '$VMName' is $($vm.State), not Running. Start or resume it first."; exit 2 }
$hb = Get-Heartbeat
if ($hb -ne 'OK') { Say "REFUSING: the heartbeat reads '$hb', not OK - the guest is not in a state to be restarted cleanly."; exit 2 }

try {
    . (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
    if (-not (Get-TestbedLease -VMName $VMName)) { Say "WARNING: no live lease on $VMName. The idle-saver may save it mid-restart (Testbed/README.md section 5b)." }
}
catch { Say "WARNING: could not read the lease directory: $($_.Exception.Message)" }

$before = Read-GuestState
Say "  guest boot time (UTC): $($before.LastBoot)"
$console = Get-ConsoleSessionState -QueryText $before.Sessions -ExpectedUser $expectedUser
if ($console) { Say "  console session: user='$($console.User)' id=$($console.SessionId) state=$($console.State)" } else { Say '  console session: none' }
$outlookRunning = (@($before.OutlookPids).Count -gt 0)
if ($outlookRunning) {
    $elevText = 'unknown'
    if ($before.Elevation -eq 1) { $elevText = 'ELEVATED' } elseif ($before.Elevation -eq 0) { $elevText = 'not elevated' }
    Say "  OUTLOOK.EXE running: pid $(@($before.OutlookPids) -join ',') in session $(@($before.OutlookSession) -join ','), $elevText"
}
else { Say '  OUTLOOK.EXE: not running' }

if (@($before.Clients).Count -gt 0) {
    Say "REFUSING: Outlook COM clients are running in the guest: $(@($before.Clients) -join ', ')."
    Say '  Application.Quit() with clients attached parks Outlook indefinitely (README.md, Outlook Lifetime).'
    Say '  Stop them first - end the agent session or the test run - then run this again. Nothing was changed.'
    exit 2
}

if (-not $Execute) {
    Say ''
    Say 'DRY RUN. Nothing was changed. With -Execute this would:'
    if ($outlookRunning) {
        Say "  1. quit Outlook in session 1 at run level $((Get-QuitRunLevels -Elevation $before.Elevation) -join ' then ') - refusing on an open dialog,"
        Say '     an open Inspector, a non-empty Outbox or the installer mutex; release every COM reference; Quit();'
        Say "     wait up to $OutlookExitTimeoutSeconds s for OUTLOOK.EXE to exit, and never kill it"
    }
    else { Say '  1. skip the Outlook quit - it is not running' }
    Say '  2. shutdown.exe /r /t 0 - no /f, and a timeout of 0 is the only one that does not imply it'
    Say '  3. wait for the heartbeat to drop and return, a later boot time, and the console session Active'
    exit 0
}

# ---------------------------------------------------------------------------------------------
# Phase 2: quit Outlook gracefully
# ---------------------------------------------------------------------------------------------
if ($outlookRunning) {
    $quit = $null
    foreach ($level in (Get-QuitRunLevels -Elevation $before.Elevation)) {
        Say "== quitting Outlook: session-1 task at run level $level =="
        $quit = Invoke-QuitTask -RunLevel $level
        foreach ($l in ($quit.Output -split "`n")) { if ($l.Trim()) { Say "  | $($l.TrimEnd())" } }
        Say "  quit task exit $($quit.Code) = $(Get-QuitMeaning $quit.Code)"
        if ($quit.Code -ne 10) { break }
    }

    if ($quit.Code -eq 15) { Say '  Outlook exited on its own before the quit task attached.' }
    elseif ($quit.Code -ne 0) {
        Say "STOPPING: Outlook was NOT quit ($(Get-QuitMeaning $quit.Code)). The guest was NOT restarted. Nothing was forced."
        Say "  job directory in the guest: $($quit.JobDir)"
        if ($quit.Code -eq 10 -or $quit.Code -eq -1) { exit 1 }
        exit 2
    }

    Say "== waiting up to $OutlookExitTimeoutSeconds s for OUTLOOK.EXE to exit (never killed) =="
    $deadline = (Get-Date).AddSeconds($OutlookExitTimeoutSeconds)
    $gone = $false
    $tq = Get-Date
    while ((Get-Date) -lt $deadline) {
        $left = Invoke-InGuest -Block { @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count }
        if ($left -eq 0) { $gone = $true; break }
        Start-Sleep -Seconds 3
    }
    if (-not $gone) {
        $after = Read-GuestState
        Say "STOPPING: OUTLOOK.EXE is still running $OutlookExitTimeoutSeconds s after Quit() (pid $(@($after.OutlookPids) -join ','))."
        Say '  It is not killed and the guest is not restarted. The documented cause is a COM client still attached;'
        Say "  known clients now running: $(if (@($after.Clients).Count) { @($after.Clients) -join ', ' } else { 'none of the known names' })."
        Say '  Find and end the client, and Outlook finishes exiting by itself (README.md, Outlook Lifetime).'
        exit 3
    }
    Say ("  OUTLOOK.EXE exited {0:N0} s after Quit()" -f ((Get-Date) - $tq).TotalSeconds)
}

# ---------------------------------------------------------------------------------------------
# Phase 3: restart, unforced
# ---------------------------------------------------------------------------------------------
Say '== restarting Windows: shutdown.exe /r /t 0 (no /f) =='
$reasonText = $Reason
if ($reasonText.Length -gt 500) { $reasonText = $reasonText.Substring(0, 500) }
try {
    $sd = Invoke-InGuest -Block {
        param($why)
        if (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) { return 'OUTLOOK-RUNNING' }
        $o = & shutdown.exe /r /t 0 /d p:4:1 /c $why 2>&1
        return "exit=$LASTEXITCODE $o"
    } -ArgumentList @($reasonText)
    Say "  shutdown.exe: $sd"
    if ($sd -eq 'OUTLOOK-RUNNING') { Say 'STOPPING: Outlook started again before the restart. Nothing was forced; run this again.'; exit 2 }
    if ($sd -notmatch '^exit=0') { Say 'STOPPING: shutdown.exe did not accept the restart.'; exit 4 }
}
catch {
    # The session can be torn down by the restart itself before the call returns.
    Say "  (the PowerShell Direct call ended with: $($_.Exception.Message.Trim()) - expected when the restart wins the race)"
}
$tr = Get-Date

# ---------------------------------------------------------------------------------------------
# Phase 4: wait until the guest is usable
# ---------------------------------------------------------------------------------------------
$limit = [TimeSpan]::FromMinutes($RestartTimeoutMinutes)
# The heartbeat alone is not enough to see a restart: measured 2026-09-24, a guest went down and
# was back within the polling interval and the heartbeat never read anything but OK, so this loop
# sat out its full deadline in front of a guest that had long since restarted. So every ~10 s it
# also asks the guest for its boot time, and a later one ends the wait just as a dropped heartbeat
# does.
$dropped = $false
$seenRestart = $false
$lastBootCheck = Get-Date
while (((Get-Date) - $tr) -lt $limit) {
    $hb = Get-Heartbeat
    if ($hb -ne 'OK') { $dropped = $true; break }
    if (((Get-Date) - $lastBootCheck).TotalSeconds -ge 10) {
        $lastBootCheck = Get-Date
        try {
            $now = Invoke-InGuest -Block { (Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().ToString('o') }
            if (([DateTime]::Parse($now) - [DateTime]::Parse($before.LastBoot)).TotalSeconds -gt 10) { $seenRestart = $true; $dropped = $true; break }
        }
        catch { }   # a guest going down may not answer; the next check, or the heartbeat, will say
    }
    Start-Sleep -Seconds 2
}
if ($seenRestart) { Say ("  the guest reports a new boot time {0:N0} s after the request (the heartbeat never read anything but OK)" -f ((Get-Date) - $tr).TotalSeconds) }
elseif ($dropped) { Say ("  heartbeat left OK after {0:N0} s ('{1}')" -f ((Get-Date) - $tr).TotalSeconds, $hb) }
else { Say "  heartbeat never left OK in $RestartTimeoutMinutes min - checking the boot time before concluding anything" }

$tb = Get-Date
$back = $false
while (((Get-Date) - $tb) -lt $limit) {
    if ((Get-Heartbeat) -eq 'OK') { $back = $true; break }
    Start-Sleep -Seconds 3
}
if (-not $back) { Say "STOPPING: the heartbeat did not return to OK within $RestartTimeoutMinutes min."; exit 5 }
Say ("  heartbeat OK again {0:N0} s after the restart request" -f ((Get-Date) - $tr).TotalSeconds)

$tc = Get-Date
$usable = $false
$restarted = $false
$last = $null
while (((Get-Date) - $tc) -lt $limit) {
    try {
        $state = Invoke-InGuest -Block {
            $os = Get-CimInstance Win32_OperatingSystem
            [pscustomobject]@{ LastBoot = $os.LastBootUpTime.ToUniversalTime().ToString('o'); Sessions = ((query session 2>&1) -join "`n"); Outlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count }
        }
        $last = $state
        # LastBootUpTime is derived (now minus uptime), and two reads of an unrestarted guest were
        # measured 6 s apart on 2026-09-24; a real restart moves it by the whole previous uptime
        # plus the shutdown and the boot. So "later" means later by more than 10 s.
        $restarted = (([DateTime]::Parse($state.LastBoot) - [DateTime]::Parse($before.LastBoot)).TotalSeconds -gt 10)
        if (-not $restarted -and -not $dropped) {
            Say 'STOPPING: the guest did NOT restart - its boot time is unchanged and its heartbeat never dropped.'
            Say '  Something in the guest refused to close and this script does not force it. Look for a window'
            Say '  holding the restart in session 1, close it, and run this again.'
            exit 4
        }
        $cs = Get-ConsoleSessionState -QueryText $state.Sessions -ExpectedUser $expectedUser
        if ($restarted -and $cs -and $cs.Usable) { $usable = $true; break }
    }
    catch { }
    Start-Sleep -Seconds 5
}
if (-not $usable) {
    Say "STOPPING: the guest did not reach a usable state within $RestartTimeoutMinutes min (restarted=$restarted)."
    if ($last) { Say "  last query session: $($last.Sessions -replace "`n", ' | ')" }
    exit 5
}

Say ("== DONE: $VMName restarted, boot time {0} (was {1}), console Active as '{2}', {3:N0} s in total ==" -f $last.LastBoot, $before.LastBoot, $expectedUser, ((Get-Date) - $t0).TotalSeconds)
if ($last.Outlook -gt 0) { Say '  NOTE: OUTLOOK.EXE is already running again after the restart - something in the guest starts it at logon.' }
exit 0
