#Requires -Version 5.1
<#
.SYNOPSIS
    Starts Outlook in the guest's interactive session WITHOUT elevation - the only Outlook that
    feeds the Windows Search index - on a named profile, and returns once its main window is up.

.DESCRIPTION
    RUN ON THE GUEST, from the elevated PowerShell Direct session (session 0) as the autologon
    account. Windows PowerShell 5.1. Starts Outlook and nothing else: it reads no mail item, holds
    no COM reference, and never quits or kills Outlook - Testbed/host/Restart-Guest.ps1 closes it.

    WHY IT EXISTS. Measured on OutlookAI-Indexed, 2026-09-24 (Docs/live-tier-on-the-vm.md section
    8 item 22): an ELEVATED Outlook does not use Windows Search at all - it never adds itself to the
    crawl scope, never pushes a store's items to the indexer, loads no Search proxy, and reports
    Store.IsInstantSearchEnabled = False. The same Outlook started NOT elevated registers within
    seconds and pushes every item of the stores its profile opens. Testbed/guest/
    Register-InteractiveTask.ps1 - the testbed's route into session 1 - registered every task at
    RunLevel Highest, so everything it started was elevated; that is why the indexed guest sat with
    an empty index for nine days. (It takes -RunLevel Limited since 2026-09-24, and the live tier
    runs at it; Highest is still its default.) This script is the other route: a one-shot
    interactive task at RunLevel LIMITED, the filtered token a user double-clicking Outlook gets.

    USE. Step 8c of Testbed/README.md section 1: with the crawl scope already set
    (Set-OutlookIndexingDisabled.ps1 -Enable -Execute), start Outlook here once per profile and let
    Set-OutlookIndexingDisabled.ps1 -Verify -WaitMinutes say when the crawl has finished.

    WHAT IT CHECKS. Refuses if Outlook is already running - a running Outlook's integrity level is
    whatever started it, and quitting it is Restart-Guest.ps1's job, not this script's. Refuses a
    profile that does not exist under HKCU\...\Outlook\Profiles (Outlook would otherwise raise its
    profile picker on an unattended desktop). Afterwards it reads the new OUTLOOK.EXE's token and
    FAILS if it came up elevated after all, because then nothing this script exists for happened.

    THE GUARD. The autologon account and the OAI- computer-name prefix, as in every guest script
    that changes Outlook's state.

.PARAMETER Profile
    The Outlook profile to open, exactly as named under HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles.

.PARAMETER MaxMinutes
    How long to wait for Outlook's main window.

.PARAMETER MinUptimeSeconds
    Start Outlook only once the guest has been up this long; 0, the default, starts it at once. The
    Exchange VM's runs pass 240: there an Outlook started within about two minutes of a boot raised the
    Object Model Guard prompt on the first out-of-process address read although Security Center read
    Defender up to date, and one started later did not (measured 2026-11-17; the maintainer keeps the
    Defender preflight as the only remedy, no auto-approve policy - Q139 (a)).

    OFFICE'S ACTIVATION WIZARD. On a guest past Office's 30-day activation grace (Testbed/MEDIA.md, "The
    licence clocks") every start with a window raises the modal "Microsoft Office Activation Wizard",
    and the Explorer window does not open until it is closed. The starter closes it with its own Close
    button - WM_CLOSE to that one window if the button cannot be pressed - never anything else on it,
    and says so: an ACTIVATION-WIZARD-DISMISSED line per dismissal, and ACTIVATION-WIZARD-DISMISSALS
    <n> at the end, which Testbed/host/Invoke-LiveTierOnGuest.ps1 records in its summary. A wizard that
    will not close fails the start.

.PARAMETER SelfTest
    Pure: checks the task definition this script builds (Interactive, RunLevel LIMITED, the profile
    passed through quoting intact) and the refusal texts. No task, no process, any machine.

.EXAMPLE
    .\Start-OutlookUnelevated.ps1 -Profile CorpusProfile
    .\Start-OutlookUnelevated.ps1 -Profile OutlookAI-Tier
    .\Start-OutlookUnelevated.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string]   $Profile,
    [int]      $MaxMinutes = 10,
    [int]      $MinUptimeSeconds = 0,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [string]   $OutlookExe = 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
    [string]   $JobRoot = 'C:\OutlookAI-Q5\jobs',
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'
$TaskName = 'OutlookAI-Unelevated'

# Office's nag once its activation grace has run out (Testbed/MEDIA.md, "The licence clocks"): a modal
# window of OUTLOOK.EXE with exactly this title, which keeps the Explorer window from opening until it
# is closed. Measured on OutlookAI-Exchange, 2026-11-17 - 45 days after Office was installed there,
# 15 past its 30-day grace.
$ActivationWizardTitle = 'Microsoft Office Activation Wizard'

# The script the task runs, in session 1, at RunLevel Limited: start Outlook on the profile, then
# wait until its Explorer window (class rctrl_renwnd32) is visible. Office's activation wizard, if
# it comes up, is closed by its own Close button - WM_CLOSE to that window alone if the button cannot
# be pressed - and every dismissal is a line of the result. Writes result.txt last.
function New-StarterScript([string] $Exe, [string] $ProfileName, [int] $Minutes, [string] $ResultPath) {
    $p = $ProfileName.Replace("'", "''")
    $e = $Exe.Replace("'", "''")
    $r = $ResultPath.Replace("'", "''")
    $w = $ActivationWizardTitle.Replace("'", "''")
    return @"
`$ErrorActionPreference = 'Stop'
`$lines = New-Object System.Collections.Generic.List[string]
try {
    if (-not ('OaiUnelevatedWin' -as [type])) {
        Add-Type -TypeDefinition @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class OaiUnelevatedWin {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    public static string Explorer(uint pid) {
        string found = null;
        EnumWindows(delegate (IntPtr h, IntPtr l) {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != pid || !IsWindowVisible(h)) return true;
            StringBuilder c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() != "rctrl_renwnd32") return true;
            StringBuilder t = new StringBuilder(512); GetWindowText(h, t, 512);
            found = t.ToString(); return false;
        }, IntPtr.Zero);
        return found;
    }
    public static IntPtr Titled(uint pid, string title) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate (IntPtr h, IntPtr l) {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != pid || !IsWindowVisible(h)) return true;
            StringBuilder t = new StringBuilder(512); GetWindowText(h, t, 512);
            if (t.ToString() != title) return true;
            found = h; return false;
        }, IntPtr.Zero);
        return found;
    }
    public static string ClassOf(IntPtr h) { StringBuilder c = new StringBuilder(256); GetClassName(h, c, 256); return c.ToString(); }
    public static bool Shown(IntPtr h) { return IsWindow(h) && IsWindowVisible(h); }
    public static bool PostClose(IntPtr h) { return PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); }
}
'@
    }
    function Close-ActivationWizard([IntPtr] `$h, [datetime] `$t0) {
        `$cls = [OaiUnelevatedWin]::ClassOf(`$h)
        `$how = `$null
        try {
            Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
            `$ae = [System.Windows.Automation.AutomationElement]
            `$button = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition(`$ae::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)), (New-Object System.Windows.Automation.PropertyCondition(`$ae::NameProperty, 'Close')))
            `$close = `$ae::FromHandle(`$h).FindFirst([System.Windows.Automation.TreeScope]::Descendants, `$button)
            if (`$close) { `$close.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); `$how = 'its Close button' }
        }
        catch { `$how = `$null }
        `$until = (Get-Date).AddSeconds(10)
        while (`$how -and [OaiUnelevatedWin]::Shown(`$h) -and (Get-Date) -lt `$until) { Start-Sleep -Milliseconds 500 }
        if ([OaiUnelevatedWin]::Shown(`$h)) {
            [void][OaiUnelevatedWin]::PostClose(`$h)
            `$how = 'WM_CLOSE to that window' + `$(if (`$how) { ' (its Close button did not close it)' } else { ' (no Close button could be pressed)' })
            `$until = (Get-Date).AddSeconds(10)
            while ([OaiUnelevatedWin]::Shown(`$h) -and (Get-Date) -lt `$until) { Start-Sleep -Milliseconds 500 }
        }
        if ([OaiUnelevatedWin]::Shown(`$h)) { return ('ACTIVATION-WIZARD-STUCK {0:N0} s: class {1} - neither its Close button nor WM_CLOSE closed it' -f ((Get-Date) - `$t0).TotalSeconds, `$cls) }
        return ('ACTIVATION-WIZARD-DISMISSED {0:N0} s: class {1}, by {2}' -f ((Get-Date) - `$t0).TotalSeconds, `$cls, `$how)
    }
    `$t0 = Get-Date
    Start-Process -FilePath '$e' -ArgumentList @('/profile', '"$p"') | Out-Null
    `$lines.Add(('started OUTLOOK.EXE /profile "{0}" at {1:HH:mm:ss}' -f '$p', `$t0))
    `$window = `$null
    `$wizards = 0
    while (((Get-Date) - `$t0).TotalMinutes -lt $Minutes) {
        `$o = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
        if (`$o) {
            `$h = [OaiUnelevatedWin]::Titled([uint32]`$o.Id, '$w')
            if (`$h -ne [IntPtr]::Zero -and `$wizards -lt 5) { `$wizards++; `$lines.Add((Close-ActivationWizard `$h `$t0)) }
            `$window = [OaiUnelevatedWin]::Explorer([uint32]`$o.Id); if (`$window) { break }
        }
        Start-Sleep -Seconds 3
    }
    if (`$window) {
        `$lines.Add(('EXPLORER-UP {0:N0} s: {1}' -f ((Get-Date) - `$t0).TotalSeconds, `$window))
        # A wizard raised just after the Explorer window is closed the same way.
        Start-Sleep -Seconds 5
        `$h = [OaiUnelevatedWin]::Titled([uint32]`$o.Id, '$w')
        if (`$h -ne [IntPtr]::Zero -and `$wizards -lt 5) { `$wizards++; `$lines.Add((Close-ActivationWizard `$h `$t0)) }
    }
    else { `$lines.Add('NO-EXPLORER within $Minutes min') }
}
catch { `$lines.Add('STARTER FAILED: ' + `$_.Exception.Message) }
finally { Set-Content -LiteralPath '$r' -Value `$lines -Encoding UTF8 }
"@
}

# How long to wait before starting Outlook so that the guest has been up -MinUptimeSeconds. Measured on
# OutlookAI-Exchange, 2026-11-17: an Outlook started 38 s and 79 s after a boot put the Object Model Guard's
# Allow/Deny prompt in front of the first out-of-process address read, with Security Center reading Defender
# up to date (productState 0x061100) before the start; one started 2 min and 8 min after a boot did not. Pure.
function Get-UptimeWait([double] $UptimeSeconds, [int] $MinUptimeSeconds) {
    if ($MinUptimeSeconds -le 0 -or $UptimeSeconds -ge $MinUptimeSeconds) { return 0 }
    return [int][Math]::Ceiling($MinUptimeSeconds - $UptimeSeconds)
}

# What the starter's result says about Office's activation wizard: how many times it was dismissed,
# and whether one stayed up. Pure, so -SelfTest pins it.
function Get-ActivationWizardTally([string[]] $Lines) {
    $all = @($Lines | Where-Object { $null -ne $_ })
    return [pscustomobject]@{
        Dismissed = @($all | Where-Object { $_ -match '^ACTIVATION-WIZARD-DISMISSED ' }).Count
        Stuck     = @($all | Where-Object { $_ -match '^ACTIVATION-WIZARD-STUCK ' }).Count
    }
}

function Get-TaskDefinitionText {
    # What Register-ScheduledTask is given, as text, so -SelfTest can pin it without registering.
    return "LogonType=Interactive RunLevel=Limited Hidden=True Action=powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File <job>\start.ps1"
}

function Invoke-SelfTest {
    function Check([string] $name, [bool] $ok) { if ($ok) { $script:p++; Write-Host "  PASS  $name" } else { $script:f++; Write-Host "  FAIL  $name" } }
    $script:p = 0; $script:f = 0
    $text = New-StarterScript -Exe 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE' -ProfileName "Corp'o Profile" -Minutes 7 -ResultPath 'C:\x\result.txt'
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($text, [ref] $null, [ref] $errors)
    Check 'the starter script parses' (@($errors).Count -eq 0)
    Check 'the profile name reaches Start-Process quoted, apostrophe escaped' ($text -match [regex]::Escape("'""Corp''o Profile""'"))
    Check 'the starter waits for the Explorer window class' ($text -match 'rctrl_renwnd32')
    Check 'the starter writes its result last, in a finally' ($text -match 'finally \{ Set-Content -LiteralPath')
    Check 'the starter never quits or kills Outlook' (-not ($text -match '(?i)(\.Quit\(|Stop-Process|taskkill|\.Kill\()'))
    Check 'the starter makes no COM call' (-not ($text -match '(?i)(GetActiveObject|New-Object\s+-ComObject)'))
    Check 'the starter looks for Office''s activation wizard by its exact title' ($text -match [regex]::Escape("Titled([uint32]`$o.Id, 'Microsoft Office Activation Wizard')") -and $text -match [regex]::Escape('if (t.ToString() != title) return true;'))
    Check 'it presses the wizard''s Close button - never Change Product Key or a link' ($text -match [regex]::Escape("NameProperty, 'Close')") -and -not ($text -match '(?i)(Change Product Key|Learn more|Hyperlink)'))
    Check 'WM_CLOSE goes only to the wizard''s own window, and only when the button did not close it' (@([regex]::Matches($text, 'PostClose\(')).Count -eq 2 -and $text -match [regex]::Escape('if ([OaiUnelevatedWin]::Shown($h)) {') -and $text -match [regex]::Escape('[void][OaiUnelevatedWin]::PostClose($h)'))
    Check 'every dismissal, and a wizard that stayed up, is a line of the result' ($text -match 'ACTIVATION-WIZARD-DISMISSED' -and $text -match 'ACTIVATION-WIZARD-STUCK')
    Check 'a wizard raised again is dismissed at most five times' (@([regex]::Matches($text, [regex]::Escape('$wizards -lt 5'))).Count -eq 2)
    $tally = Get-ActivationWizardTally @('started OUTLOOK.EXE', 'ACTIVATION-WIZARD-DISMISSED 270 s: class NUIDialog, by its Close button', 'EXPLORER-UP 277 s: Inbox')
    Check 'the tally counts one dismissal and nothing stuck' ($tally.Dismissed -eq 1 -and $tally.Stuck -eq 0)
    $tally = Get-ActivationWizardTally @('ACTIVATION-WIZARD-STUCK 40 s: class X - neither', 'NO-EXPLORER within 10 min')
    Check 'and a wizard that stayed up as stuck' ($tally.Dismissed -eq 0 -and $tally.Stuck -eq 1)
    Check 'and no wizard at all as nothing' ((Get-ActivationWizardTally @('EXPLORER-UP 3 s: Inbox')).Dismissed -eq 0 -and (Get-ActivationWizardTally @()).Stuck -eq 0)
    Check 'a guest up 79 s waits 161 s more for a 240 s minimum' ((Get-UptimeWait 79 240) -eq 161)
    Check 'a guest up longer than the minimum does not wait' ((Get-UptimeWait 480 240) -eq 0)
    Check 'no minimum, no wait - the default, for every guest but the Exchange VM' ((Get-UptimeWait 5 0) -eq 0)
    Check 'the task definition is Interactive and LIMITED' ((Get-TaskDefinitionText) -match 'LogonType=Interactive RunLevel=Limited')
    # Read off the PARSED file, so the text of this check cannot count as a use of what it forbids.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $levels = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-ScheduledTaskPrincipal' }, $true) | ForEach-Object {
            $els = @($_.CommandElements | ForEach-Object { $_.Extent.Text })
            $i = [Array]::IndexOf($els, '-RunLevel'); if ($i -ge 0 -and $i + 1 -lt $els.Count) { $els[$i + 1] } else { '<none>' } })
    Check 'this file builds exactly one task principal' ($levels.Count -eq 1)
    Check 'and it is RunLevel Limited - never Highest' ((@($levels) -join ',') -eq 'Limited')
    $ast2 = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $defined = @{}
    foreach ($fd in $ast2.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fd.Name] = $true }
    $unknown = @()
    foreach ($cmd in $ast2.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $name = $cmd.GetCommandName()
        if (-not $name -or $defined.ContainsKey($name)) { continue }
        if (Get-Command -Name $name -ErrorAction SilentlyContinue) { continue }
        $unknown += "$name (line $($cmd.Extent.StartLineNumber))"
    }
    Check ('every command this file calls exists' + $(if ($unknown.Count) { ': MISSING ' + ($unknown -join ', ') } else { '' })) ($unknown.Count -eq 0)    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:p, $script:f)
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

if (-not $Profile) { throw '-Profile is mandatory: name the Outlook profile to open.' }

# travels-alone: it needs nothing staged beside it, wherever a run or a step starts it from.
# The guest guard: Test-GuestIdentity and Assert-TestbedGuestIdentity, copies of ScriptHelpers.ps1's that
# Tools/Checks/check-powershell-51.ps1 check 4 holds identical, then what this script would have done elsewhere.
function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

function Assert-TestbedGuestIdentity {
    param([string] $Why)
    $users = @($script:ExpectedUser)
    $prefix = [string]$script:ExpectedComputerNamePrefix
    if (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $users -Prefix $prefix) { return }
    throw ("REFUSING TO RUN.`n`n  logged on as : '{0}'          (allowed: {1})`n  computer name: '{2}'      (must start with: '{3}')`n`n{4}" -f $env:USERNAME, ($users -join ', '), $env:COMPUTERNAME, $prefix, $Why)
}

$GuestGuardWhy = @"
This script registers a scheduled task that starts Outlook on this machine's desktop: that belongs on a
test guest, never on a machine with real mail on it.
"@
Assert-TestbedGuestIdentity -Why $GuestGuardWhy

$running = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Write-Host "REFUSING: OUTLOOK.EXE is already running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')). Its integrity level is whatever started it."
    Write-Host 'Close it with Testbed/host/Restart-Guest.ps1 -VMName <guest> -Execute, then run this again.'
    exit 2
}
$profileKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile"
if (-not (Test-Path -LiteralPath $profileKey)) {
    Write-Host "REFUSING: there is no Outlook profile named '$Profile' ($profileKey). Outlook would raise its profile picker on an unattended desktop."
    exit 2
}
if (-not (Test-Path -LiteralPath $OutlookExe)) { throw "Outlook is not at $OutlookExe." }

$uptime = ((Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime).TotalSeconds
$wait = Get-UptimeWait $uptime $MinUptimeSeconds
if ($wait -gt 0) {
    Write-Host ("waiting {0} s before starting Outlook: the guest has been up {1:N0} s, and an Outlook started within about two minutes of a boot met the Object Model Guard prompt (measured 2026-11-17)" -f $wait, $uptime)
    Start-Sleep -Seconds $wait
}

$jobDir = Join-Path $JobRoot ('unelevated-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $jobDir | Out-Null
$starter = Join-Path $jobDir 'start.ps1'
$result = Join-Path $jobDir 'result.txt'
Set-Content -LiteralPath $starter -Value (New-StarterScript -Exe $OutlookExe -ProfileName $Profile -Minutes $MaxMinutes -ResultPath $result) -Encoding UTF8

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false }
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$starter`""
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::FromMinutes($MaxMinutes + 5)) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $TaskName
Write-Host "task $TaskName (Interactive, RunLevel Limited) started; job $jobDir"

$deadline = (Get-Date).AddMinutes($MaxMinutes + 2)
while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $result)) { Start-Sleep -Seconds 2 }
if (-not (Test-Path -LiteralPath $result)) { Write-Host "FAILED: the starter wrote nothing in $MaxMinutes min (job kept at $jobDir)."; exit 1 }
$lines = @(Get-Content -LiteralPath $result)
foreach ($l in $lines) { Write-Host "  $l" }
try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false } catch { }

# The one thing this script exists for: an Outlook that is NOT elevated. Read its token.
if (-not ('OaiUnelevatedToken' -as [type])) {
    Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class OaiUnelevatedToken {
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint a, bool i, int p);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr p, uint a, out IntPtr t);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr t, int c, out int i, int l, out int r);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    public static int Elevation(int pid) { IntPtr p = OpenProcess(0x1000, false, pid); if (p == IntPtr.Zero) return -1; try { IntPtr t; if (!OpenProcessToken(p, 8, out t)) return -1; try { int e, r; if (!GetTokenInformation(t, 20, out e, 4, out r)) return -1; return e != 0 ? 1 : 0; } finally { CloseHandle(t); } } finally { CloseHandle(p); } }
}
'@
}
$outlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($outlook.Count -eq 0) { Write-Host 'FAILED: OUTLOOK.EXE is not running after the start.'; exit 1 }
$elevation = [OaiUnelevatedToken]::Elevation($outlook[0].Id)
if ($elevation -ne 0) {
    Write-Host "FAILED: OUTLOOK.EXE pid $($outlook[0].Id) reads elevation=$elevation. An elevated Outlook does not feed the index; this start achieved nothing."
    exit 1
}
$wizard = Get-ActivationWizardTally $lines
if ($wizard.Stuck -gt 0) { Write-Host "FAILED: Office's activation wizard ('$ActivationWizardTitle') stayed up - see the ACTIVATION-WIZARD-STUCK line above. Outlook (pid $($outlook[0].Id)) is left running."; exit 1 }
if (-not ($lines -match '^EXPLORER-UP')) { Write-Host "FAILED: Outlook (pid $($outlook[0].Id), not elevated) showed no main window in $MaxMinutes min."; exit 1 }
Write-Host "OK: OUTLOOK.EXE pid $($outlook[0].Id) in session $($outlook[0].SessionId), NOT elevated, profile '$Profile'. Left running - close it with Testbed/host/Restart-Guest.ps1."
# The runner reads this line into its summary: a run that started Outlook says whether Office's nag was in the way.
Write-Host "ACTIVATION-WIZARD-DISMISSALS $($wizard.Dismissed)"
exit 0
