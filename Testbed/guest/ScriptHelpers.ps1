#Requires -Version 5.1
<#
.SYNOPSIS
    The helpers this repository's scripts share, in one copy (Q154, Q166 (b), decision 215 (a)).
    Not run on its own: dot-sourced by every script that uses one, before it does.

.DESCRIPTION
    WHY ONE FILE, AND WHY HERE. The scripts used to restate these helpers word for word, as many as
    seventeen copies of one function. Q166 (b) puts them in one file that the scripts which travel
    together dot-source, and Tools/Checks/check-powershell-51.ps1 holds the copies a script that has
    to reach a machine on its own keeps - marked "# travels-alone: <how>" - identical to this file.
    Decision 215 (a) put the file in Testbed/guest/, so that every route that carries the guest
    scripts to a test machine - the whole folder, flat, by Testbed/host/Invoke-LiveTierOnGuest.ps1,
    or by hand as Testbed/README.md step 4a stages them - carries it unchanged:

        Testbed/guest/*.ps1    . (Join-Path $PSScriptRoot 'ScriptHelpers.ps1')
        Testbed/host/*.ps1     . (Join-Path $PSScriptRoot '..\guest\ScriptHelpers.ps1')
        Tools/*.ps1            . (Join-Path $PSScriptRoot '..\Testbed\guest\ScriptHelpers.ps1')
        Tools/Checks/*.ps1     . (Join-Path $PSScriptRoot '..\..\Testbed\guest\ScriptHelpers.ps1')

    WHY A DOT-SOURCED .ps1 AND NOT A MODULE. Under Windows PowerShell 5.1 a .psm1 does not pass the
    caller's error settings through: a script block a module's function runs belongs to the caller's
    session state, so the 'Continue' that Invoke-NativeCommand sets in its own scope never reaches
    the caller's script block, and 5.1 ends the caller on the first line a program writes to stderr.
    Measured 2026-11-18 under Windows PowerShell 5.1 and PowerShell 7.7, a script with 'Stop' running
    cmd.exe through Invoke-NativeCommand: from a module, 5.1 died with NativeCommandError; dot-sourced,
    both shells came back with exit 3 and both lines of output. PowerShell 7 does not show the
    difference.

    NOTHING HERE RUNS WHEN THE FILE IS DOT-SOURCED but the function definitions, and the file takes
    no parameters: a param() block would bind in the caller's scope and overwrite the caller's own
    -SelfTest switch. ASCII only, so Windows PowerShell 5.1 reads it the same without a byte order
    mark.
#>

# A NATIVE PROGRAM RUNS THROUGH HERE (Q78). Under $ErrorActionPreference = 'Stop', Windows
# PowerShell 5.1 turns the first line a native program writes to a redirected stderr - 2>$null, 2>&1
# and *> alike - into a terminating NativeCommandError, so one warning or progress line ends the
# script before its exit code can be read; and a script under Testbed/ or Tools/ runs behind a
# redirection or in a remoting host as often as not, where even an unredirected call dies that way.
# PowerShell 7 does not. Measured on the host 2026-09-24 and 2026-09-27. So, here and only here:
#   * 'Continue' holds in THIS function's scope. The caller's 'Stop' is never changed, so there is
#     nothing to restore and nothing else is relaxed.
#   * The try is load-bearing. Without one, 'Continue' also demotes a terminating error inside the
#     block - the program not being found at all - to a printed message, and the caller goes on to
#     read a stale $LASTEXITCODE. Inside a try it stops the caller exactly as it always did.
#     Measured in both shells.
#   * Stderr lines come back as plain strings in both shells, never as ErrorRecords: 5.1 renders
#     those with a position block around every line, and an empty one as an exception type name.
#   * The exit code is left in $LASTEXITCODE, and the caller checks it.
# Tools/Checks/check-powershell-51.ps1 fails the build on a native call under Testbed/ or Tools/, or
# a redirected one anywhere, that does not go through a function like this one.
function Invoke-NativeCommand {
    param([Parameter(Mandatory = $true)] [scriptblock] $NativeCommand)

    $ErrorActionPreference = 'Continue'
    try {
        & $NativeCommand | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { $_ }
        }
    }
    catch {
        throw
    }
}

# A duration as a person reads it: "45 s" under a minute, "4 m 05 s" from one minute on. Whole
# seconds, rounded down at every step, so the seconds never read 60: the build-VM runner's copy
# rounded them instead and printed 119.6 s as "1 m 60 s".
function Format-Seconds {
    param([double] $Seconds)
    $whole = [int][Math]::Floor($Seconds)
    if ($whole -lt 60) { return ('{0} s' -f $whole) }
    return ('{0} m {1:00} s' -f [int][Math]::Floor($whole / 60), ($whole % 60))
}

# A long-running program with its output going STRAIGHT TO FILES - never through this process's pipe -
# and a deadline. A program that leaves children behind can hold a pipe open after it has finished,
# and a declared timeout then never fires. .Handle is read before waiting because Start-Process
# -PassThru otherwise loses the exit code (measured, Install-DotnetSdk.ps1). The arguments are joined
# with spaces and quoted by nobody: quote one that needs it with Format-Argument. The result says the
# exit code (-1 on a timeout, when the program is killed), whether it timed out, and where its two
# logs are: <LogStem>.out.txt and <LogStem>.err.txt.
function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)] [string] $FilePath,
        [Parameter(Mandatory = $true)] [string[]] $ArgumentList,
        [Parameter(Mandatory = $true)] [string] $LogStem,
        [int] $TimeoutMinutes = 20,
        [string] $WorkingDirectory
    )
    $out = "$LogStem.out.txt"
    $err = "$LogStem.err.txt"
    foreach ($f in @($out, $err)) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } }
    $startArgs = @{
        FilePath               = $FilePath
        ArgumentList           = ($ArgumentList -join ' ')
        RedirectStandardOutput = $out
        RedirectStandardError  = $err
        NoNewWindow            = $true
        PassThru               = $true
    }
    if ($WorkingDirectory) { $startArgs['WorkingDirectory'] = $WorkingDirectory }
    $p = Start-Process @startArgs
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = -1; TimedOut = $true; Out = $out; Err = $err }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false; Out = $out; Err = $err }
}

# The last lines of a log Invoke-Logged wrote, indented under the step that failed.
function Show-LogTail([string] $Path, [int] $Lines = 30) {
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        foreach ($l in (Get-Content -LiteralPath $Path -Tail $Lines)) { Write-Host "      | $l" }
    }
}

# One argument for Start-Process's joined command line: quoted when it holds white space.
function Format-Argument([string] $Value) {
    if ($Value -match '\s') { return '"' + $Value + '"' }
    return $Value
}

# Whether this process runs elevated - as a member of Administrators with its full token.
function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Whether ANOTHER process's token is elevated, by its id: 1 elevated, 0 not, -1 when it cannot be read - no
# such process, or no right to open it or its token. TokenElevation (20) of GetTokenInformation, through
# OpenProcess with PROCESS_QUERY_LIMITED_INFORMATION (0x1000) and OpenProcessToken with TOKEN_QUERY (8); C# 5,
# which Windows PowerShell 5.1's Add-Type compiles. Four scripts each compiled a class of their own for it
# until 2026-11-20 (Q154, the survey's P10). A job, or a session on another machine, that cannot dot-source
# this file is handed the function's own text and runs it: [scriptblock]::Create($text), then -ProcessId.
function Get-ProcessElevation {
    param([Parameter(Mandatory = $true)] [int] $ProcessId)
    if (-not ('OaiProcessToken' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OaiProcessToken {
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr token, int infoClass, out int info, int length, out int returned);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
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
    return [OaiProcessToken]::Elevation($ProcessId)
}

# THE ELEVATION GUARD of a script whose work needs an administrator (Q154, P05): it returns in an
# elevated session and otherwise throws a refusal that says so, then -Why - what the script writes, and
# what an unelevated run would leave half done.
function Assert-Elevated([string] $Why) {
    if (Test-IsElevated) { return }
    throw "REFUSING TO RUN: this session is not elevated.`n`n$Why"
}

# THE BITNESS GUARD of a script that needs a 64-bit PowerShell - for x64 Outlook, or for the Windows
# Search provider, which only an x64 host loads: it returns in a 64-bit process and otherwise throws a
# refusal that says so, then -Why, then where the 64-bit Windows PowerShell is.
function Assert-Bitness([string] $Why) {
    if ([Environment]::Is64BitProcess) { return }
    throw "REFUSING TO RUN: this is a 32-bit PowerShell.`n`n$Why`n`nStart Windows PowerShell from C:\Windows\System32\WindowsPowerShell\v1.0."
}

# THE OUTLOOK GUARD of a script that must not run while Outlook does: it returns when no OUTLOOK.EXE runs
# and otherwise throws a refusal that names each one's pid, then -Why - what a running Outlook would undo
# or break. It never closes Outlook itself: mailbox-safety rule 7 forbids taskkill on OUTLOOK.EXE outright.
function Assert-OutlookNotRunning([string] $Why) {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return }
    throw ("REFUSING TO RUN: OUTLOOK.EXE is running (pid {0}).`n`n{1}`n`nQuit Outlook gracefully and run this again. DO NOT taskkill it: mailbox-safety rule 7 forbids that outright." -f (($running | ForEach-Object { $_.Id }) -join ', '), $Why)
}

# A line of progress: the time, then the message, on the console - and, where the calling script keeps
# a log at $LogPath, the same line appended there, its folder made first. A log that cannot be written
# does not stop the work it records. $LogPath is the calling script's own - its parameter, or a variable
# it sets outside its functions - read as $script:LogPath. Read plainly, it was also found in a script
# that ran the caller with &: Publish-LiveTierPayload.ps1, which keeps no log, wrote its lines into the
# live runner's run.log too (measured 2026-11-18, in a live run and in both shells).
function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    $log = $script:LogPath
    if ($log) {
        try {
            $dir = Split-Path -Parent $log
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Add-Content -LiteralPath $log -Value $line -Encoding UTF8
        }
        catch {
            # A log that cannot be written must not stop the work it was recording.
        }
    }
}

# A line of a report the calling script writes out whole at its end: printed as it is, with no time,
# and kept in $ReportLines - the list the script makes before its first line.
function Write-ReportLine([string] $Text = '') {
    Write-Host $Text
    [void] $ReportLines.Add($Text)
}

# A line printed as it is and appended to the calling script's own log, $script:LogPath, with no time
# before it - for Build-Corpus.ps1 and Reset-HubPopulation.ps1, whose logs read as plain reports. A log
# that cannot be written is passed over.
function Write-LogLine([string] $Text) {
    Write-Host $Text
    $log = $script:LogPath
    if ($log) { Add-Content -LiteralPath $log -Value $Text -ErrorAction SilentlyContinue }
}

# A request refused before anything was touched: said as a line of progress, and the script ends with
# 4, the exit code every runner here gives a refusal.
function Stop-Refused([string] $Why) {
    Say "REFUSED: $Why"
    exit 4
}

# A run directory's name: when the run started, what it ran where the runner names that - a guest's
# kind, a probe's scenario - and the first 12 characters of the commit. Sortable and readable in a
# directory listing, but not unique: two runs of one commit started in the same second get the same
# name, which New-RunDirectory then tells apart. Anything but a commit id is refused.
function New-RunId {
    param([DateTime] $When, [string] $Sha, [string] $Label = '')
    if ($Sha -notmatch '^[0-9a-f]{12,40}$') { throw "Not a commit id: '$Sha'." }
    if ($Label) { return ('{0:yyyyMMdd-HHmmss}-{1}-{2}' -f $When, $Label, $Sha.Substring(0, 12)) }
    return ('{0:yyyyMMdd-HHmmss}-{1}' -f $When, $Sha.Substring(0, 12))
}

# A run's own directory under -Parent, made so that no other run can have it, before anything is named
# after the run. The directory is made under a name nobody else uses and then renamed to -RunId; a
# rename fails when the name is taken, which New-Item -Force never does, and a taken name gets -2, then
# -3, and so on. Two runs of one commit started in the same second shared one directory until
# 2026-11-18: stream K2's two indexed runs at 12:14:18 wrote one log, one's git worktree add failed, and
# the other's build failed when the first deleted the tree. Returns RunId - the name the run goes by
# from here on - and Path, the directory's full path. -Parent is made when it is missing. -MaxTries 1
# claims the one name or refuses: Switch-AddInBuild.ps1's dev build folder, which is never reused.
function New-RunDirectory {
    param([Parameter(Mandatory = $true)] [string] $Parent, [Parameter(Mandatory = $true)] [string] $RunId, [int] $MaxTries = 99)
    if ($RunId -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { throw "Not a run id: '$RunId'." }
    if (-not (Test-Path -LiteralPath $Parent)) { New-Item -ItemType Directory -Force -Path $Parent | Out-Null }
    $parentFull = (Resolve-Path -LiteralPath $Parent).ProviderPath
    $claim = Join-Path $parentFull ('.claim-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($claim)
    try {
        for ($n = 1; $n -le $MaxTries; $n++) {
            $id = $RunId
            if ($n -gt 1) { $id = '{0}-{1}' -f $RunId, $n }
            $path = Join-Path $parentFull $id
            try {
                [System.IO.Directory]::Move($claim, $path)
                return [pscustomobject]@{ RunId = $id; Path = $path }
            }
            catch {
                # Taken: the name is there. Anything else - the parent gone, no access - is no reason to
                # try the next name.
                if (-not (Test-Path -LiteralPath $path)) { throw }
            }
        }
        if ($MaxTries -le 1) { throw "REFUSING: $RunId is already taken under $parentFull." }
        throw "REFUSING: $RunId and -2 to -$MaxTries are all taken under $parentFull."
    }
    finally {
        if (Test-Path -LiteralPath $claim) { Remove-Item -LiteralPath $claim -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# What a TRX file says. Outcomes are counted from the results themselves, and the file's own
# counters are kept beside them so a disagreement can be seen; NotExecuted is how xUnit's skips
# arrive. Failed tests carry the first line of their message.
function ConvertFrom-TrxText {
    param([string] $Text)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Text)
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('t', 'http://microsoft.com/schemas/VisualStudio/TeamTest/2010')
    $results = @($doc.SelectNodes('//t:Results/t:UnitTestResult', $ns))
    $counts = [ordered]@{ total = $results.Count; passed = 0; failed = 0; skipped = 0; other = 0 }
    $failed = @()
    $skipped = @()
    foreach ($r in $results) {
        $outcome = $r.GetAttribute('outcome')
        if ($outcome -eq 'Passed') { $counts['passed']++ }
        elseif ($outcome -eq 'Failed') {
            $counts['failed']++
            $message = ''
            $node = $r.SelectSingleNode('t:Output/t:ErrorInfo/t:Message', $ns)
            if ($null -ne $node) { $message = (($node.InnerText -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1) }
            $failed += [pscustomobject]@{ name = $r.GetAttribute('testName'); message = [string]$message }
        }
        elseif ($outcome -eq 'NotExecuted') { $counts['skipped']++; $skipped += $r.GetAttribute('testName') }
        else { $counts['other']++; $failed += [pscustomobject]@{ name = $r.GetAttribute('testName'); message = "outcome $outcome" } }
    }
    $counters = $doc.SelectSingleNode('//t:ResultSummary/t:Counters', $ns)
    $declared = $null
    if ($null -ne $counters) {
        $declared = [ordered]@{
            total       = [int]$counters.GetAttribute('total')
            passed      = [int]$counters.GetAttribute('passed')
            failed      = [int]$counters.GetAttribute('failed')
            notExecuted = [int]$counters.GetAttribute('notExecuted')
        }
    }
    return [pscustomobject]@{ Counts = $counts; Declared = $declared; Failed = $failed; Skipped = $skipped }
}

# A test VM's lease, as Set-TestbedLease.ps1 writes it, judged for one run or caller: 'free' when there
# is none, it has expired or its expiry cannot be read (Testbed/host/TestbedLeasePath.ps1's rule); 'ours'
# when its reason is exactly the caller's own; 'gone' when the process that took it has ended
# (Get-LeaseHolderState), which may be taken over as a free one is (2026-11-20) - so a host that
# restarts under a run no longer holds the runs after it off until the dead run's lease expires; and
# 'theirs' otherwise, which is never taken over. $IsAlive answers for a holder's process, given as a
# queue ticket is: the runners and Set-TestbedLease.ps1 pass Testbed/host/TestbedPool.ps1's
# Test-TestbedTicketAlive, which asks for its id, started at the second the lease records. A caller
# that gives none reads every holder as alive, as every caller did before 2026-11-20.
function Get-LeaseDecision {
    param([object] $Lease, [string] $Reason, [long] $NowUnix, [scriptblock] $IsAlive = $null, [string] $ComputerName = $env:COMPUTERNAME)
    if ($null -eq $Lease) { return 'free' }
    $expires = 0L
    try { $expires = [long]$Lease.expiresUnix } catch { return 'free' }
    if ($expires -le $NowUnix) { return 'free' }
    if (([string]$Lease.reason) -ceq $Reason) { return 'ours' }
    if ((Get-LeaseHolderState -Lease $Lease -ComputerName $ComputerName -IsAlive $IsAlive) -eq 'gone') { return 'gone' }
    return 'theirs'
}

# Who holds a live lease, for Get-LeaseDecision and for Set-TestbedLease.ps1's words: 'byHand' - taken
# from a prompt, not by a script, so no process lives as long as its holder and only its expiry or a
# release with its reason ends it; 'noProcess' - it names none; 'elsewhere' - another computer's
# process, which this one cannot ask about; 'alive' - its process runs, by $IsAlive, or there is no
# $IsAlive to ask; 'gone' - that process has ended. A lease written before 2026-11-20 says neither how
# it was taken nor when its process started, and is judged by its process id alone.
function Get-LeaseHolderState {
    param([object] $Lease, [string] $ComputerName, [scriptblock] $IsAlive)
    if ($Lease.byHand -eq $true) { return 'byHand' }
    $holderPid = 0
    [void][int]::TryParse([string]$Lease.holderPid, [ref]$holderPid)
    if ($holderPid -le 0) { return 'noProcess' }
    $holderHost = [string]$Lease.holderHost
    if ($holderHost -and -not [string]::Equals($holderHost, $ComputerName, [System.StringComparison]::OrdinalIgnoreCase)) { return 'elsewhere' }
    if ($null -eq $IsAlive) { return 'alive' }
    if (& $IsAlive ([pscustomobject]@{ pid = $holderPid; processStartUnix = $Lease.holderStartUnix })) { return 'alive' }
    return 'gone'
}

# The exit code of work a runner ran on a guest through Register-InteractiveTask.ps1, from the line
# the runner's own call writes after it: 'RUN-REGISTER-EXIT: <code>'. $null when there is no such line.
function Get-TaskExit([string] $Text) {
    $m = [regex]::Match($Text, 'RUN-REGISTER-EXIT: (-?\d+)')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return $null
}

# What a VM's heartbeat reads, in Get-VMIntegrationService's own words - OK, No Contact, Lost
# Communication - or 'absent' when the VM has no heartbeat service: for Restart-Guest.ps1 and
# Set-GuestClockFrozen.ps1, which say what they read. A wait for OK that says nothing more asks
# Testbed/host/TestbedPool.ps1's Test-TestbedHeartbeat, one narrow query to this one's 23 WMI operations.
function Get-VmHeartbeatStatus([string] $VMName) {
    $hb = Get-VMIntegrationService -VMName $VMName -Name 'Heartbeat' -ErrorAction SilentlyContinue
    if (-not $hb) { return 'absent' }
    return [string]$hb.PrimaryStatusDescription
}

# Whether a session is a test guest's, by the two marks a guest built from the answer file carries: the
# account logged on is one of -Users, and the computer name starts with -Prefix - Testbed/host/
# New-AnswerFile.ps1 names a guest 'OAI-...'. An empty prefix matches nothing, never everything.
function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

# THE GUEST GUARD of a script that must touch no machine but a test guest: it returns on one - the
# calling script's own $ExpectedUser and $ExpectedComputerNamePrefix, read as $script: variables - and
# otherwise throws a refusal that says who is logged on where and then -Why, what the script would
# have done to the machine. Tools/Checks/check-testbed-references.ps1 check 9 holds every guest script
# that writes to calling a guard before its first write.
function Assert-TestbedGuestIdentity {
    param([string] $Why)
    $users = @($script:ExpectedUser)
    $prefix = [string]$script:ExpectedComputerNamePrefix
    if (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $users -Prefix $prefix) { return }
    throw ("REFUSING TO RUN.`n`n  logged on as : '{0}'          (allowed: {1})`n  computer name: '{2}'      (must start with: '{3}')`n`n{4}" -f $env:USERNAME, ($users -join ', '), $env:COMPUTERNAME, $prefix, $Why)
}

# ---- Attaching to Outlook (Q154, S7).

# Whether an OUTLOOK.EXE runs in this process's own session: what an attach to a running Outlook asks before it
# may use the class factory, and what a probe asks before it opens anything, since the Running Object Table cannot
# tell - an Outlook that has not yet lost the focus is not in it.
function Test-OutlookInThisSession {
    $session = (Get-Process -Id $PID).SessionId
    return (@(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session }).Count -gt 0)
}

# Outlook's Application object, from the Outlook already running in this process's own session - never a new one.
# Through the Running Object Table first. Office registers there only once its window has first lost the focus
# (MK_E_UNAVAILABLE before that: the Outlook guests, 2026-11-18), and PowerShell 7 has no GetActiveObject at all, so
# then through the class factory - Outlook runs one instance, and its class hands back the running one - but only once
# an OUTLOOK.EXE is seen running in this session (Test-OutlookInThisSession), so the factory starts nothing. Returns
# App ($null when it did not attach), Via - 'RunningObjectTable' or 'ClassFactory', for a log line to say - Running,
# whether an OUTLOOK.EXE runs in this session, and Why, what stopped it, or $null. Reset-HubPopulation.ps1,
# OutlookDialogUia.ps1's dialog helper, Measure-FolderIdResolution.ps1 and Measure-SignatureRoaming.ps1 each attached
# their own way until 2026-11-18, and the probe could not attach where the first two could.
function Connect-RunningOutlook {
    if (-not (Test-OutlookInThisSession)) {
        return [pscustomobject]@{ App = $null; Via = $null; Running = $false; Why = "no OUTLOOK.EXE runs in this session ($((Get-Process -Id $PID).SessionId)), and none is started" }
    }
    $rot = ''
    try { return [pscustomobject]@{ App = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application'); Via = 'RunningObjectTable'; Running = $true; Why = $null } }
    catch { $rot = $_.Exception.Message }
    try { return [pscustomobject]@{ App = (New-Object -ComObject Outlook.Application); Via = 'ClassFactory'; Running = $true; Why = $null } }
    catch { return [pscustomobject]@{ App = $null; Via = $null; Running = $true; Why = "neither the Running Object Table ($rot) nor the class factory ($($_.Exception.Message)) gave the running Outlook" } }
}

# ---- Core's window calls, for the scripts that look at windows (Q171 (c), the survey's P10).

# The text of OutlookAI.Core's NativeWindows.cs - EnumWindows, GetWindowThreadProcessId, IsWindowVisible,
# GetClassNameW and the rest, C# 5 with its usings inside its namespace - for a script to compile after a C#
# source of its own in one Add-Type, in place of declaring those calls itself. Staged beside the guest scripts
# (Invoke-LiveTierOnGuest.ps1, Invoke-TrustRemovalProbe.ps1 and Testbed/README.md step 4a stage it with them);
# in a checkout, where Core keeps it. check-pinned-constants.ps1 #16 holds every route to staging it. A process
# compiles it once: a second Add-Type of another source that carries it fails with TYPE_ALREADY_EXISTS (measured
# 2026-11-20 under 5.1 and 7), so each script compiles it guarded by its own type, and none runs in a process
# another of them compiled it in.
function Get-NativeWindowsSource {
    foreach ($candidate in @((Join-Path $PSScriptRoot 'NativeWindows.cs'), (Join-Path $PSScriptRoot '..\..\McpServer\OutlookAI.Core\AddIn\NativeWindows.cs'))) {
        if (Test-Path -LiteralPath $candidate) { return [System.IO.File]::ReadAllText($candidate) }
    }
    throw "NativeWindows.cs is neither beside $PSScriptRoot nor in this checkout's McpServer\OutlookAI.Core\AddIn: stage it with the guest scripts, as Invoke-LiveTierOnGuest.ps1 does."
}

# ---- The measurements' helpers (Q154, S7).

# The first -Digits hex digits - 12 unless a caller says otherwise - of the SHA-256 of a value's UTF-8, or
# of its upper-case spelling with -IgnoreCase: enough for a measurement to say two values are equal or
# not, and not enough to say what a value is. $null for no value. Measure-FolderIdResolution.ps1 compares
# ids and names whatever their case; Measure-SignatureRoaming.ps1 prints a name as 8 digits, as its
# earlier records do. Each held a copy until 2026-11-18.
function Get-ShortHash {
    param([string] $Value, [ValidateRange(1, 64)] [int] $Digits = 12, [switch] $IgnoreCase)
    if ([string]::IsNullOrEmpty($Value)) { return $null }
    if ($IgnoreCase) { $Value = $Value.ToUpperInvariant() }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Value)) } finally { $sha.Dispose() }
    return (-join @($bytes | ForEach-Object { $_.ToString('x2') })).Substring(0, $Digits)
}

# ---- The self-tests' helpers (Q154, S7a: P04).

# A value as a self-test prints and compares it: '<null>' for $null, an array as [a, b] item by item, and
# anything else as its string - so $null, '' and an empty array never read alike. Add-OutlookPstStore.ps1,
# New-OutlookProfile.ps1, Set-DefaultOutlookProfile.ps1 and New-LiveTestSettings.ps1 each held a copy.
function ConvertTo-SelfTestText {
    param($Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Array]) {
        $parts = @()
        foreach ($item in $Value) { $parts += (ConvertTo-SelfTestText $item) }
        return '[' + ($parts -join ', ') + ']'
    }
    return [string] $Value
}

# ---- The Outlook profile scripts' helpers (Q154, S7c step 3, and S7a for the hive choice and the
# ImportPRF clearance): New-OutlookProfile.ps1, Set-DefaultOutlookProfile.ps1, New-TierProfile.ps1 and
# Build-Corpus.ps1 each held a copy. The constants they read are the calling script's own, as $script:
# variables.

# Whether an ...\Office\<major>\Outlook key with these values and subkeys is a REAL Outlook hive, or just
# the shell this product's installer leaves behind - whose one subkey the calling script names in
# $script:InstallerFootprintSubKeyName. Mirrors OfficeVersions.IsOutlookHive exactly, case-insensitive
# footprint comparison included.
function Test-IsOutlookHive {
    param([string[]] $ValueNames, [string[]] $SubKeyNames)

    if ($null -ne $ValueNames -and $ValueNames.Count -gt 0) { return $true }
    if ($null -eq $SubKeyNames) { return $false }

    foreach ($subKey in $SubKeyNames) {
        if ($subKey -ne $script:InstallerFootprintSubKeyName) { return $true }
    }
    return $false
}

# The line(s) a profile script prints under the chosen hive when other Outlook-shaped keys exist; the
# calling script's $script:SupportedOfficeVersions tells a supported major from one never chosen. Pure.
# Measured on OAI-UNINDEXED: HKCU\Software\Microsoft\Office\8.0\Outlook holds one value (First-Run) and
# nothing else, and an old note offered "-OfficeVersion 8.0" as the fix. An unsupported major is named,
# never offered.
function Format-OtherHiveNote {
    param([string] $ChosenVersion, [object[]] $RealHives)

    $supported = @()
    $unsupported = @()
    if ($null -ne $RealHives) {
        foreach ($hive in $RealHives) {
            if ($null -eq $hive -or $hive.Version -eq $ChosenVersion) { continue }
            if ($script:SupportedOfficeVersions -contains $hive.Version) { $supported += $hive.Version }
            else { $unsupported += $hive.Version }
        }
    }

    $lines = @()
    if ($supported.Count -gt 0) {
        $lines += "other supported Outlook hive(s) here: $($supported -join ', ') - use -OfficeVersion if this one is wrong"
    }
    if ($unsupported.Count -gt 0) {
        $lines += "also an Outlook-shaped key under unsupported major(s) $($unsupported -join ', ') - never chosen, and not a fix for anything"
    }
    return , $lines
}

# Every ...\Office\<major>\Outlook key on this machine - under the calling script's
# $script:OfficeRootKeyPath - as a profile script's candidates for its hive. Reads only: no value is
# written and no key is created.
function Get-OutlookHiveCandidate {
    if (-not (Test-Path -LiteralPath $script:OfficeRootKeyPath)) { return , @() }

    $candidates = @()
    foreach ($version in (Get-ChildItem -LiteralPath $script:OfficeRootKeyPath -ErrorAction SilentlyContinue)) {
        if ($version.PSChildName -notmatch '^\d+\.\d+$') { continue }

        $outlookPath = "$script:OfficeRootKeyPath\$($version.PSChildName)\Outlook"
        if (-not (Test-Path -LiteralPath $outlookPath)) { continue }

        $key = Get-Item -LiteralPath $outlookPath
        $candidates += [pscustomobject]@{
            Version     = $version.PSChildName
            Path        = $outlookPath
            ValueNames  = @($key.GetValueNames())
            SubKeyNames = @($key.GetSubKeyNames())
        }
    }
    return , $candidates
}

# The hive a profile script works in, picked from candidates shaped
# @{ Version = '16.0'; Path = '...'; ValueNames = @(); SubKeyNames = @() } - Get-OutlookHiveCandidate's.
# Pure. Returns Chosen (a candidate, or $null), RealHives (every candidate passing the shape rule, in the
# order given) and Problem (the refusal text, when Chosen is $null); Build-Corpus.ps1 reads Chosen alone.
# -RequestedVersion overrides the probe order and the supported list both: an operator who names a
# version has said something no script can know. It does NOT override the shape rule, because writing
# ImportPRF or DefaultProfile into a key Outlook never reads is a silent no-op - the operator starts
# Outlook, sees nothing change, and has nothing at all to read to find out why.
function Select-OutlookHive {
    param([object[]] $Candidates, [string] $RequestedVersion)

    $real = @()
    $present = @()
    if ($null -ne $Candidates) {
        foreach ($candidate in $Candidates) {
            $present += $candidate.Version
            if (Test-IsOutlookHive -ValueNames $candidate.ValueNames -SubKeyNames $candidate.SubKeyNames) {
                $real += $candidate
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
        $named = $null
        if ($null -ne $Candidates) {
            foreach ($candidate in $Candidates) {
                if ($candidate.Version -eq $RequestedVersion) { $named = $candidate }
            }
        }

        if ($null -eq $named) {
            $presentText = '<none>'
            if ($present.Count -gt 0) { $presentText = ($present -join ', ') }
            return [pscustomobject]@{
                Chosen    = $null
                RealHives = $real
                Problem   = "REFUSING: -OfficeVersion $RequestedVersion was asked for, and $script:OfficeRootKeyPath\$RequestedVersion\Outlook does not exist. Version keys present: $presentText. Drop -OfficeVersion and let the detection choose."
            }
        }

        if (-not (Test-IsOutlookHive -ValueNames $named.ValueNames -SubKeyNames $named.SubKeyNames)) {
            return [pscustomobject]@{
                Chosen    = $null
                RealHives = $real
                Problem   = "REFUSING: -OfficeVersion $RequestedVersion was asked for, and $($named.Path) holds nothing but a '$script:InstallerFootprintSubKeyName' subkey - which this product's own installer writes under every supported major, on every install. Outlook does not read that key, so anything this script writes into it would be a silent no-op. Name the major Outlook really uses, or drop -OfficeVersion."
            }
        }

        return [pscustomobject]@{ Chosen = $named; RealHives = $real; Problem = $null }
    }

    foreach ($version in $script:SupportedOfficeVersions) {
        foreach ($candidate in $real) {
            if ($candidate.Version -eq $version) {
                return [pscustomobject]@{ Chosen = $candidate; RealHives = $real; Problem = $null }
            }
        }
    }

    # Nothing supported qualified. If some OTHER major did, name it - choosing it is an operator
    # decision (-OfficeVersion), not one this script may take on itself.
    $unsupported = @()
    foreach ($candidate in $real) { $unsupported += $candidate.Version }

    $problem = @"
REFUSING: found no supported Outlook hive with any content under $script:OfficeRootKeyPath.

Probed in this order: $($script:SupportedOfficeVersions -join ', ') - the order the shipped product
probes in (OfficeVersions.Supported).

Every one of them was either absent or held nothing but a '$script:InstallerFootprintSubKeyName'
subkey, which this product writes under 15.0, 16.0 and 17.0 alike on every install - so a
present-but-empty key proves nothing. Either classic Outlook is not installed, or it has never
been started.
"@

    if ($unsupported.Count -gt 0) {
        $problem += @"


There IS a real Outlook hive at: $($unsupported -join ', '). That major is not one this product
supports, so nothing here will choose it for you. If it is the one Outlook really uses, say so:
    -OfficeVersion $($unsupported[0])
"@
    }

    return [pscustomobject]@{ Chosen = $null; RealHives = $real; Problem = $problem }
}

# The profile names: the subkey names under <hive>\Profiles. A subkey IS a profile.
function Get-OutlookProfileName {
    param([string] $ProfilesKeyPath)

    if (-not (Test-Path -LiteralPath $ProfilesKeyPath)) { return , @() }
    $names = @(Get-ChildItem -LiteralPath $ProfilesKeyPath -ErrorAction SilentlyContinue |
            ForEach-Object { $_.PSChildName })
    return , $names
}

# The ProfileName= a .prf carries, read from its text. Pure. Only the [General] section counts - it is
# where Outlook reads the name - and the first ProfileName= there wins. $null when there is none: a file
# that names no profile is not one the calling script wrote, and nothing is decided on its behalf.
function Get-PrfProfileName {
    param([string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $section = ''
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.Trim()
        if ($line.StartsWith(';')) { continue }
        if ($line.StartsWith('[') -and $line.EndsWith(']')) {
            $section = $line.Substring(1, $line.Length - 2).Trim()
            continue
        }
        if ($section -eq 'General' -and $line.StartsWith('ProfileName=', [System.StringComparison]::OrdinalIgnoreCase)) {
            $value = $line.Substring('ProfileName='.Length).Trim()
            if ($value.Length -gt 0) { return $value }
            return $null
        }
    }
    return $null
}

# What -Verify does about ImportPRF, for New-OutlookProfile.ps1 and New-TierProfile.ps1, whose .prf files
# both carry OverwriteProfile=Yes. Pure: the machine half gathers five facts and acts on the answer, so
# both scripts' -SelfTest walk every branch.
#
# WHY -VERIFY CLEARS IT, AND WHY IT CANNOT CLEAR IT ANY EARLIER. Outlook reads ImportPRF at startup, so
# an ImportPRF left behind is a rebuild waiting for a start - and a rebuild is what would detach a store
# attached or filled since, a corpus store among them. Removing it closes that. But removing it BEFORE
# Outlook has read it cancels the import itself, silently: the profile is never built.
#
# WHAT TELLS THE TWO APART, and it is two measured facts, not one. (1) Outlook REMOVES ImportPRF itself
# within 5 s of the start that imports it (OAI-UNINDEXED, 2026-09-24, sampled every 5 s) - so a value
# still set means no start has acted on it. (2) Outlook ignores ImportPRF while First-Run or FirstRun
# exists (documented), -Execute deletes both, and Outlook writes First-Run back only once a start has
# COMPLETED its first run: during the tier profile's first session, yes; during a start that stopped on
# the account wizard, not in 300 s and not at shutdown. So "still set, First-Run absent" is an import
# still to come - leave it - and "still set, First-Run present" is a value Outlook will not act on until
# something removes First-Run again - a latent rebuild, and the only state this ever removes.
#
# Decision is one of:
#   NotSet       ImportPRF is absent. Nothing can re-import.
#   Clear        it names THIS profile's .prf, the profile exists, and First-Run is back: the import
#                has had its start, and all that is left is a re-import. REMOVE IT.
#   KeepPending  it names this profile's .prf and neither First-Run nor FirstRun is there: Outlook has
#                not started since -Execute, so the import has not happened and removing the value
#                would cancel it. -Verify ran too early.
#   KeepFailed   it names this profile's .prf, Outlook HAS started since (First-Run is back), and the
#                profile is not there: the import did not take. Left as evidence - Outlook ignores it
#                while First-Run exists.
#   KeepOther    it names another profile's .prf, or a file whose profile name cannot be read: ANOTHER
#                import may be pending, and removing the value would cancel that one. Left alone.
#                Build-Corpus.ps1 refuses while ImportPRF is set, so a corpus is not built underneath
#                it either way.
function Resolve-ImportPrfClearance {
    param(
        [string] $ImportPrfValue,
        [string] $PrfProfileName,
        [Parameter(Mandatory = $true)] [string] $ProfileName,
        [bool] $ProfileExists,
        [bool] $FirstRunPresent
    )

    if ([string]::IsNullOrEmpty($ImportPrfValue)) {
        return [pscustomobject]@{
            Decision = 'NotSet'
            Message  = 'ImportPRF is not set, so no Outlook start can re-import. (Outlook removes it itself once it has imported the file - measured on this build.)'
        }
    }

    if ([string]::IsNullOrEmpty($PrfProfileName) -or $PrfProfileName -ne $ProfileName) {
        $whose = "a file whose profile name could not be read"
        if (-not [string]::IsNullOrEmpty($PrfProfileName)) { $whose = "the .prf for profile '$PrfProfileName'" }
        return [pscustomobject]@{
            Decision = 'KeepOther'
            Message  = "ImportPRF is set to '$ImportPrfValue' - $whose, not '$ProfileName'. Left alone: another import may still be pending, and removing the value would cancel it without a word. Build-Corpus.ps1 refuses to build while ImportPRF is set. To remove it anyway: .\New-OutlookProfile.ps1 -ClearImportPrf -Execute"
        }
    }

    if (-not $FirstRunPresent) {
        return [pscustomobject]@{
            Decision = 'KeepPending'
            Message  = "ImportPRF still names this profile's .prf and neither First-Run nor FirstRun is back, so Outlook has NOT started since -Execute - the import has not happened yet. Removing the value now would cancel it, so it is left in place. Start Outlook once, let it settle, and run -Verify again."
        }
    }

    if (-not $ProfileExists) {
        return [pscustomobject]@{
            Decision = 'KeepFailed'
            Message  = "ImportPRF still names this profile's .prf, Outlook HAS started since -Execute (First-Run is back), and there is no profile named '$ProfileName': the import did not take. Left in place as evidence; while First-Run exists Outlook does not act on it."
        }
    }

    return [pscustomobject]@{
        Decision = 'Clear'
        Message  = "ImportPRF still names this profile's .prf after the import has run. Removing it: left in place, it is a rebuild waiting for the next start that finds First-Run gone - and a rebuild detaches whatever was attached or filled since."
    }
}

# Whether one entry of a profile's account-manager key is a MAIL account. Pure; from a measurement on
# OAI-UNINDEXED, 2026-09-24: Outlook lists every MAPI service there, wrapping data files and address books
# under one CLSID - the account-less CorpusProfile holds {ED475414-...} 'Outlook Data File' (MSUPST MS) and
# {ED475414-...} 'Outlook Address Book' (CONTAB) with COM Accounts.Count 0, and the tier profile's POP3
# account is {ED475411-...}. Only a wrapper around a data file or an address book is excluded; anything
# else - an Exchange wrapper (MSEMS), a service nobody named, an unreadable clsid - counts, so an unknown
# shape refuses rather than slipping through. The service names are the ones section 6 of this
# project's .prf files maps.
function Test-IsMailAccountEntry {
    param([string] $Clsid, [string] $ServiceName)

    if ($Clsid -eq '{ED475414-B0D6-11D2-8C3B-00104B2A6676}' -and @('MSUPST MS', 'MSPST MS', 'CONTAB', 'EMABLT', 'MSPST AB') -contains $ServiceName) { return $false }
    return $true
}

# ---------------------------------------------------------------------------------------------
# THE ADD-IN'S VSTO TRUST AND ITS PAYLOAD (Q154, the survey's P08): what Testbed/guest/Install-OutlookAIAddIn.ps1,
# Tools/Switch-AddInBuild.ps1 and Testbed/host/Publish-AddInPayload.ps1 each restated, two or three copies
# apiece until 2026-11-20 - how the VSTO runtime keys and compares a trust entry, the key a signed manifest
# carries, the trust entries themselves, Outlook's hard-disable list, and the payload manifest's shape.
# ---------------------------------------------------------------------------------------------

# The URL the VSTO runtime keys an inclusion entry by, from the registered Manifest value: the
# `|vstolocal` suffix off, backslashes to forward slashes, file:/// in front - which is the
# Uri.ToString() form the runtime itself writes. [MEASURED: the maintainer's workstation holds
# `file:///C:/Users/.../OutlookAI/Setup/OutlookAI.vsto` for this installer's
# `file:///{app}\OutlookAI.vsto|vstolocal`.] Spaces stay literal, as the runtime writes them.
function ConvertTo-InclusionUrl {
    param([string] $ManifestValue)
    if (-not $ManifestValue) { return $null }
    $v = $ManifestValue.Trim()
    $bar = $v.IndexOf('|')
    if ($bar -ge 0) { $v = $v.Substring(0, $bar) }
    if ($v.StartsWith('file:///', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(8) }
    elseif ($v.StartsWith('file://', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(7) }
    $v = $v.Replace('\', '/')
    return 'file:///' + $v
}

# The runtime compares entries AS URIs, and on a DOS path that comparison ignores case.
function Test-SameInclusionUrl {
    param([string] $A, [string] $B)
    $x = ConvertTo-InclusionUrl $A
    $y = ConvertTo-InclusionUrl $B
    if (-not $x -or -not $y) { return $false }
    return [string]::Equals($x, $y, [System.StringComparison]::OrdinalIgnoreCase)
}

# A key as the <RSAKeyValue> string a trust entry's PublicKey holds, from its two base64 numbers.
function ConvertTo-RsaKeyValueXml {
    param([Parameter(Mandatory = $true)] [string] $ModulusBase64, [Parameter(Mandatory = $true)] [string] $ExponentBase64)
    $m = ($ModulusBase64 -replace '\s', '')
    $e = ($ExponentBase64 -replace '\s', '')
    return "<RSAKeyValue><Modulus>$m</Modulus><Exponent>$e</Exponent></RSAKeyValue>"
}

# The key a signed manifest (.vsto or .dll.manifest) is signed with: its ds:RSAKeyValue, every
# occurrence the same key, as the <RSAKeyValue> string the trust store holds.
function Get-ManifestSigningKeyXml {
    param([Parameter(Mandatory = $true)] [string] $ManifestXml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($ManifestXml)
    $nodes = $doc.SelectNodes("//*[local-name()='RSAKeyValue']")
    if ($null -eq $nodes -or $nodes.Count -eq 0) { throw 'The manifest carries no RSAKeyValue - it is not signed.' }
    $keys = @()
    foreach ($n in $nodes) {
        $mod = $n.SelectSingleNode("*[local-name()='Modulus']")
        $exp = $n.SelectSingleNode("*[local-name()='Exponent']")
        if ($null -eq $mod -or $null -eq $exp) { throw 'An RSAKeyValue in the manifest has no Modulus or no Exponent.' }
        $keys += (ConvertTo-RsaKeyValueXml -ModulusBase64 $mod.InnerText -ExponentBase64 $exp.InnerText)
    }
    $distinct = @($keys | Sort-Object -Unique)
    if ($distinct.Count -ne 1) { throw "The manifest is signed with $($distinct.Count) different keys; expected one." }
    return $distinct[0]
}

# A key as two hex strings with leading zero bytes dropped - the same two numbers the runtime's
# ExportCspBlob comparison sees, without asking a crypto provider for them.
function ConvertFrom-RsaKeyXml {
    param([string] $Xml)
    if (-not $Xml) { throw 'empty key' }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $mod = $doc.SelectSingleNode("//*[local-name()='Modulus']")
    $exp = $doc.SelectSingleNode("//*[local-name()='Exponent']")
    if ($null -eq $mod -or $null -eq $exp) { throw 'not an RSAKeyValue: Modulus or Exponent missing' }
    $out = @{}
    foreach ($pair in @(@('Modulus', $mod.InnerText), @('Exponent', $exp.InnerText))) {
        $bytes = [Convert]::FromBase64String(($pair[1] -replace '\s', ''))
        $start = 0
        while ($start -lt ($bytes.Length - 1) -and $bytes[$start] -eq 0) { $start++ }
        $hex = New-Object System.Text.StringBuilder
        for ($i = $start; $i -lt $bytes.Length; $i++) { [void]$hex.Append($bytes[$i].ToString('X2')) }
        $out[$pair[0]] = $hex.ToString()
    }
    return $out
}

# Whether two <RSAKeyValue> strings are the same key; a string that is not one is no key at all.
function Test-SameRsaKey {
    param([string] $A, [string] $B)
    try {
        $x = ConvertFrom-RsaKeyXml $A
        $y = ConvertFrom-RsaKeyXml $B
        return ($x.Modulus -eq $y.Modulus -and $x.Exponent -eq $y.Exponent)
    }
    catch {
        return $false
    }
}

# Every VSTO inclusion entry under HKCU\<InclusionKey> - its subkey name, Url and PublicKey -
# read-only; none when the key is not there.
function Get-InclusionEntries {
    param([string] $InclusionKey)
    $entries = @()
    $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($InclusionKey, $false)
    if ($null -eq $root) { return $entries }
    try {
        foreach ($name in $root.GetSubKeyNames()) {
            $k = $root.OpenSubKey($name, $false)
            if ($null -eq $k) { continue }
            try { $entries += [pscustomobject]@{ Name = $name; Url = [string]$k.GetValue('Url'); PublicKey = [string]$k.GetValue('PublicKey') } }
            finally { $k.Close() }
        }
    }
    finally { $root.Close() }
    return $entries
}

# Outlook's hard-disable list holds UTF-16 blobs naming the add-in's path or name.
function Test-DisabledItemsMention {
    param([object[]] $Blobs)
    foreach ($b in @($Blobs)) {
        if ($null -eq $b) { continue }
        $text = [System.Text.Encoding]::Unicode.GetString([byte[]]$b)
        if ($text.ToLowerInvariant().Contains('outlookai')) { return $true }
    }
    return $false
}

# The payload manifest's text forms: an upper-case SHA-256, a full lower-case commit id, and a
# four-part version with every part 0 to 65535.
function Test-Sha256Text { param([string] $Value) return [bool]($Value -match '^[0-9A-F]{64}$') }
function Test-CommitText { param([string] $Value) return [bool]($Value -match '^[0-9a-f]{40}$') }
function Test-FourPartVersion {
    param([string] $Value)
    if (-not $Value) { return $false }
    if ($Value -notmatch '^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}$') { return $false }
    foreach ($part in $Value.Split('.')) { if ([int]$part -gt 65535) { return $false } }
    return $true
}

# What is wrong with an add-in payload manifest (addin-payload.json), as problems; none when it is whole.
# Its writer, Publish-AddInPayload.ps1, also names the pinned VSTO redistributable's SHA-256 and the
# contract files, and then the manifest must carry exactly that runtime and a hash for each. The guest's
# installer names neither: it compares each contract file with the suite's itself, and reports a drift
# as a finding rather than as a manifest it cannot read.
function Test-PayloadManifestShape {
    param($Manifest, [string] $VstoRuntimeSha256, [string[]] $ContractFiles)
    $problems = @()
    if ($null -eq $Manifest) { return @('no manifest at all') }
    if (-not (Test-CommitText ([string]$Manifest.commit))) { $problems += 'commit: not a 40-character lower-case hex id' }
    if (-not (Test-FourPartVersion ([string]$Manifest.version))) { $problems += 'version: not a four-part version' }
    if ($null -eq $Manifest.installer -or -not $Manifest.installer.file) { $problems += 'installer.file: missing' }
    elseif (-not (Test-Sha256Text ([string]$Manifest.installer.sha256))) { $problems += 'installer.sha256: not an upper-case SHA-256' }
    foreach ($f in @('OutlookAI.dll', 'OutlookAI.vsto', 'OutlookAI.dll.manifest')) {
        if ($null -eq $Manifest.addin -or -not (Test-Sha256Text ([string]$Manifest.addin.$f))) { $problems += "addin.${f}: not an upper-case SHA-256" }
    }
    if ($null -eq $Manifest.signing -or -not ([string]$Manifest.signing.publicKeyXml).StartsWith('<RSAKeyValue><Modulus>')) {
        $problems += 'signing.publicKeyXml: not an <RSAKeyValue> string'
    }
    if ($VstoRuntimeSha256 -and ($null -eq $Manifest.vstoRuntime -or ([string]$Manifest.vstoRuntime.sha256) -ne $VstoRuntimeSha256)) {
        $problems += 'vstoRuntime.sha256: not the pinned redistributable'
    }
    foreach ($rel in @($ContractFiles | Where-Object { $_ })) {
        if ($null -eq $Manifest.contract -or -not (Test-Sha256Text ([string]$Manifest.contract.$rel))) { $problems += "contract.${rel}: not an upper-case SHA-256" }
    }
    return $problems
}
