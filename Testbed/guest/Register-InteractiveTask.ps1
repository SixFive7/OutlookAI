<#
.SYNOPSIS
    Runs a script in the guest's INTERACTIVE session and brings its output back. The one route
    to anything that touches Outlook.

.DESCRIPTION
    RUN THIS ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.

    THE PROBLEM IT SOLVES. PowerShell Direct - `Invoke-Command -VMName`, which is how the host
    reaches this guest without a network - lands in SESSION 0. Outlook can never finish starting
    there. Every corpus verb, the MCP server, the COM host and the whole live suite therefore
    cannot be launched directly from the host, and the failure is not a clean error: Outlook
    half-starts and the call hangs until something times out.

    THE ROUTE. A scheduled task whose principal is registered with LogonType=Interactive runs in
    the logged-on user's session - session 1 on this guest, which stays alive because autologon
    is enabled. So: write the work into a job directory, start a task that runs it, and poll for
    a sentinel file.

    WHY A FILE AND NOT A RETURN VALUE. A scheduled task's stdout does not come back to whatever
    started it. There is no pipe. This is the same file-based shape a host-side runner would use, and for the
    same reason; a script that expects to read the output of a task it started will read nothing
    and report success.

    THE CONTRACT, so a caller can drive it without reading the implementation:

        <JobRoot>\<jobId>\cmd.ps1     the wrapper the task runs. It runs -ScriptPath's script, or work.ps1.
        <JobRoot>\<jobId>\work.ps1    an inline -Script, written as a script of its own (since 2026-09-27).
        <JobRoot>\<jobId>\out.txt     everything the work wrote, stdout and stderr merged.
        <JobRoot>\<jobId>\exit.txt    the exit code. ITS EXISTENCE IS THE COMPLETION SIGNAL.

    exit.txt is written last, in a `finally`, so it exists even when the work throws. A job that
    has out.txt and no exit.txt is still running or died without unwinding - those two are
    genuinely different and this script says which by whether the deadline was reached.

    WHAT exit.txt HOLDS, for -ScriptPath and -Script alike: the work's own `exit N`; 1 when it ended
    on a terminating error - a throw, an -ErrorAction Stop, a script that does not parse; otherwise
    the exit code of the last native program it ran; otherwise 0. (And 4 for a Limited job that
    refused - below.) UNTIL 2026-09-27 AN INLINE -Script's `exit N` WROTE 0 (Q91): the body was
    pasted into cmd.ps1, where `exit` ended the wrapper itself, and the `finally` recorded the 0 the
    wrapper had started from - a failure that read as success. A body that did not parse made
    cmd.ps1 unparseable, so no exit.txt came at all and the caller waited out -TimeoutSeconds. Both
    measured on the host, both shells; the body now runs as work.ps1, the way a -ScriptPath script
    always has, which ends both.

    NO WINDOW, EVER. `-WindowStyle Hidden` is passed to powershell.exe, where it hides that
    process's own console; the task settings ask for hidden as well. Do not "fix" a problem here
    with `Start-Process -Verb RunAs`: UAC elevation takes the foreground even with a hidden
    window, which was measured on this project and is why that verb is banned outright.

    RESIDUAL RISK, stated rather than hidden: a task in an interactive session is not the same as
    a windowless service, and a child process that creates its own window will draw one on the
    guest's console. Nothing in the testbed should - the MCP server and the tools are console
    apps started with CreateNoWindow - but if you see a flash on the guest, this is where to look.

    THE RUN LEVEL, SINCE 2026-09-24: -RunLevel Highest (the default) or Limited. Until then every
    job ran elevated, and that was the whole reason OutlookAI-Indexed was never indexed: an
    ELEVATED Outlook does not use Windows Search at all - no crawl-scope rule, no item pushed,
    Store.IsInstantSearchEnabled False (Docs/live-tier-on-the-vm.md section 8 item 22). Highest
    stays the default because every caller written before this parameter expects it, and the
    installers need it: Install-OutlookAIAddIn.ps1 -Phase Install refuses outright when not elevated
    (it installs the VSTO runtime machine-wide), Install-DotnetSdk.ps1 installs under Program Files
    and writes the machine environment. A default of Limited would have broken those, loudly,
    mid-build. (Install-OutlookAIAddIn.ps1 -Phase FirstRun is the reverse, since 2026-10-03: it starts
    Outlook, so it refuses an ELEVATED token and runs here at -RunLevel Limited.)
    Limited is the filtered token a user double-clicking a program gets - NOT elevated, medium
    integrity - and it is what the LIVE TIER runs at (Testbed/README.md section 4c), for three
    reasons: the index tests need an index that moves while they run, which only a non-elevated
    Outlook gives; the product itself runs non-elevated by design (ComGateway.cs, S8); and COM does
    not attach across integrity levels, so the suite and the Outlook it drives must share one.
    Two checks come with it. A Limited job whose token is elevated anyway (UAC off, say) REFUSES:
    none of the work runs, out.txt says why, and the exit code is 4. And before any job starts,
    an OUTLOOK.EXE already running at the OTHER level is named in a WARNING - not a refusal,
    because not every job touches Outlook - with the fix: Testbed/host/Restart-Guest.ps1.

    THE GUEST GUARD, SINCE 2026-09-24. This registers a scheduled task that runs whatever script
    text it is handed, at the requested run level, in the logged-on user's session - on the maintainer's workstation
    that is a task running as the maintainer. It now dot-sources OutlookMapiInterop.ps1 and calls
    Assert-TestbedGuest before anything else: before the job directory is created, before cmd.ps1
    is written, and before any task is unregistered, registered or started. For the account it is
    always run as - vmadmin, over PowerShell Direct - the guard returns without a word and every
    line after it is the same as before. STAGE OutlookMapiInterop.ps1 BESIDE IT, as
    Testbed/README.md section 1 step 4a already does; copied alone it now fails loudly on the
    dot-source. The guard checks who REGISTERS the task, here; -UserId, who the task runs AS, is
    unchanged, and the work handed over carries its own guard when it is a guest script. Proven
    on the maintainer's workstation the same day, under Windows PowerShell 5.1, with -Script and
    with -ScriptPath: "REFUSING TO RUN. This session is logged on as ...", zero calls reaching a
    tripwire that stood in for every write command, no job directory and no task. The vmadmin
    half RAN on OAI-UNINDEXED the same day, over PowerShell Direct: the guard returned without a
    word and three jobs ran - one at the default level (High Mandatory Level) and two at
    -RunLevel Limited (Medium Mandatory Level); the Limited job's own cmd.ps1, run again from the
    elevated session, refused with exit 4 and ran none of its work.

.PARAMETER ScriptPath
    A .ps1 on the guest to run in session 1.

.PARAMETER Script
    Inline script text, as an alternative to -ScriptPath.

.PARAMETER TaskName
    Scheduled task name. One task is reused across jobs; it is registered on first use.

.PARAMETER UserId
    The interactive account. Defaults to the current user, which is right when this is invoked
    over PowerShell Direct as the autologon account.

.PARAMETER TimeoutSeconds
    How long to wait for exit.txt. Generous by default: a corpus build is ~13 minutes and a
    600 s exhaustive scan is a legal outcome, not a hang.

.PARAMETER ExpectedUser
    The account the guest guard accepts - who may REGISTER the task, not who it runs as (that is
    -UserId). The default is the guard; see OutlookMapiInterop.ps1.

.PARAMETER RunLevel
    Highest (the default): the work runs ELEVATED, as every job did before this parameter existed -
    what installers need. Limited: the work runs NOT elevated, at medium integrity - what the live
    tier needs, and what anything that should feed Windows Search needs. With Limited the work
    refuses to run (exit 4) if its token is elevated anyway. See the banner for why.

.EXAMPLE
    .\Register-InteractiveTask.ps1 -ScriptPath C:\OutlookAI-Q5\guest-measure.ps1
    .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 7200 -Script "<the live-tier lines of Testbed/README.md section 4c>"
    .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\tools\OutlookAI.RemediationTools.exe' corpus-census --store 'Outlook Data File' --allow-store 'Outlook Data File' --corpus-id vm2 --seed 7777 --anchor 2026-08-19 --count 20000"
#>
[CmdletBinding(DefaultParameterSetName = 'Path')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Path')] [string] $ScriptPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Inline')] [string] $Script,
    [string] $TaskName = 'OutlookAI-Interactive',
    [string] $JobRoot = 'C:\OutlookAI-Q5\jobs',
    [string] $UserId,
    [int]    $TimeoutSeconds = 2400,
    [string[]] $ExpectedUser = @('vmadmin'),
    [ValidateSet('Highest', 'Limited')] [string] $RunLevel = 'Highest',
    [switch] $KeepJob
)

$ErrorActionPreference = 'Stop'

# THE GUARD, FIRST - before the job directory, cmd.ps1 and the scheduled task. For vmadmin on a
# guest it returns without a word and nothing below it changes. See the banner.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

if (-not $UserId) { $UserId = "$env:USERDOMAIN\$env:USERNAME" }

# OUTLOOK AND THE WORK MUST SHARE A LEVEL. COM does not attach across integrity levels - the
# product's own rule (ComGateway.cs, S8) - so a job at one level that drives an Outlook already
# running at the other fails on every COM call. Said up front, as a WARNING and not a refusal,
# because not every job touches Outlook. The token is read the way Start-OutlookUnelevated.ps1
# reads it, restated here as this repository restates its shared helpers.
$outlookNow = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
if ($outlookNow.Count -gt 0) {
    if (-not ('OaiInteractiveTaskToken' -as [type])) {
        Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class OaiInteractiveTaskToken {
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint a, bool i, int p);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr p, uint a, out IntPtr t);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr t, int c, out int i, int l, out int r);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    public static int Elevation(int pid) { IntPtr p = OpenProcess(0x1000, false, pid); if (p == IntPtr.Zero) return -1; try { IntPtr t; if (!OpenProcessToken(p, 8, out t)) return -1; try { int e, r; if (!GetTokenInformation(t, 20, out e, 4, out r)) return -1; return e != 0 ? 1 : 0; } finally { CloseHandle(t); } } finally { CloseHandle(p); } }
}
'@
    }
    $workElevated = ($RunLevel -eq 'Highest')
    foreach ($o in $outlookNow) {
        $e = [OaiInteractiveTaskToken]::Elevation($o.Id)
        if ($e -lt 0) {
            Write-Warning "OUTLOOK.EXE pid $($o.Id) is running and its token could not be read, so whether it is at this job's level (RunLevel $RunLevel) is unknown."
        }
        elseif (($e -eq 1) -ne $workElevated) {
            $outlookLevel = 'NOT elevated'
            if ($e -eq 1) { $outlookLevel = 'ELEVATED' }
            Write-Warning ("OUTLOOK.EXE pid {0} (session {1}) is running {2}, and this job runs at RunLevel {3}. COM does not attach across that difference: anything this job does to Outlook will fail. Close Outlook first - Testbed/host/Restart-Guest.ps1 -VMName <guest> -Execute - and the work starts its own at its own level." -f $o.Id, $o.SessionId, $outlookLevel, $RunLevel)
        }
    }
}

# THE JOB'S SCRIPTS: cmd.ps1 always, and work.ps1 for an inline -Script. A function so that the
# host can exercise exactly this code - lifted out by its syntax tree, and what it writes run under
# both PowerShells - without running this script, whose guard refuses there (and must). Defined
# after the guard, called once below. Returns the job's four paths.
function Write-InteractiveTaskJob {
    param(
        [Parameter(Mandatory = $true)] [string] $JobDir,
        [Parameter(Mandatory = $true)] [string] $RunLevel,
        [string] $ScriptPath,
        [string] $Script
    )

    $paths = [pscustomobject]@{
        Cmd  = (Join-Path $JobDir 'cmd.ps1')
        Work = (Join-Path $JobDir 'work.ps1')
        Out  = (Join-Path $JobDir 'out.txt')
        Exit = (Join-Path $JobDir 'exit.txt')
    }
    $outPath = $paths.Out
    $exitPath = $paths.Exit

    if ($ScriptPath) {
        if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "Not found on the guest: $ScriptPath" }
        $body = "& '" + ($ScriptPath -replace "'", "''") + "'"
    }
    else {
        # AN INLINE -Script RUNS AS ITS OWN SCRIPT, work.ps1 (Q91, decided 2026-09-27) - the way a
        # -ScriptPath script always has. Pasted into the wrapper, as it was, its `exit N` ended
        # cmd.ps1 itself and exit.txt said 0 (the banner). As work.ps1, `exit N` ends work.ps1 alone
        # and reaches $LASTEXITCODE; a throw still reaches the wrapper's catch; the last native
        # program's exit code still reaches $LASTEXITCODE; and a body that does not parse fails at
        # once, with its parse error, instead of leaving no exit.txt. Written with a byte order
        # mark, so Windows PowerShell 5.1 reads any non-ASCII in it as UTF-8, whichever PowerShell
        # wrote it.
        [System.IO.File]::WriteAllText($paths.Work, $Script, (New-Object System.Text.UTF8Encoding($true)))
        $body = "& '" + ($paths.Work -replace "'", "''") + "'"
    }

    # The level check, Limited only: the token is read BEFORE the work, and an elevated one runs
    # none of it - an elevated job asked to be non-elevated is the exact failure -RunLevel exists to
    # end (the banner), and it would otherwise look like a normal run. With Highest the check is
    # empty, $code is still 0 when the gate below is reached, and the work runs as it always did.
    $levelCheck = ''
    if ($RunLevel -eq 'Limited') {
        $levelCheck = @"
    `$levelPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (`$levelPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        'REFUSED: this job was registered at RunLevel Limited and its token is ELEVATED anyway, so none of the work ran. Is UAC off (EnableLUA)? An elevated Outlook never feeds Windows Search, and COM will not attach across integrity levels.' | Out-File -FilePath '$outPath' -Encoding utf8 -Append
        `$code = 4
    }
    else {
        'run level: Limited - this token is NOT elevated (read before the work ran)' | Out-File -FilePath '$outPath' -Encoding utf8 -Append
    }
"@
    }

    # The wrapper, not the work. It exists to guarantee three things the work cannot guarantee
    # about itself: everything it writes reaches out.txt, an exit code is recorded even when it
    # throws, and exit.txt is written LAST so its existence really does mean "finished".
    $wrapper = @"
`$ErrorActionPreference = 'Continue'
`$code = 0
try {
$levelCheck
    if (`$code -eq 0) {
    & {
$body
    } *>&1 | Out-File -FilePath '$outPath' -Encoding utf8 -Append
    if (`$LASTEXITCODE -ne `$null) { `$code = `$LASTEXITCODE }
    }
}
catch {
    `$code = 1
    "WRAPPER CAUGHT: `$(`$_.Exception.GetType().Name): `$(`$_.Exception.Message)" | Out-File -FilePath '$outPath' -Encoding utf8 -Append
}
finally {
    "`$code" | Out-File -FilePath '$exitPath' -Encoding utf8
}
"@
    Set-Content -LiteralPath $paths.Cmd -Value $wrapper -Encoding UTF8
    return $paths
}

$jobId = [guid]::NewGuid().ToString('N')
$jobDir = Join-Path $JobRoot $jobId
New-Item -ItemType Directory -Force -Path $jobDir | Out-Null

if ($PSCmdlet.ParameterSetName -eq 'Path') { $job = Write-InteractiveTaskJob -JobDir $jobDir -RunLevel $RunLevel -ScriptPath $ScriptPath }
else { $job = Write-InteractiveTaskJob -JobDir $jobDir -RunLevel $RunLevel -Script $Script }
$cmdPath = $job.Cmd
$outPath = $job.Out
$exitPath = $job.Exit

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false }

# -WindowStyle Hidden here is an argument to powershell.exe, which honours it for its own
# console. That is NOT the same as putting it inside a -ArgumentList handed to Start-Process,
# where it is silently a no-op - a distinction this project has been bitten by.
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$cmdPath`""

# Interactive is the entire point: it puts the process in the logged-on session, where Outlook
# can actually finish starting. The run level is -RunLevel's: Highest by default, which prompts
# for nothing because a task's elevation is granted at registration; Limited for the live tier
# and anything else that must not be elevated (the banner says why).
$principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel $RunLevel

$settings = New-ScheduledTaskSettingsSet -Hidden `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::FromSeconds([Math]::Max($TimeoutSeconds, 60) + 300)) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings | Out-Null

Write-Host "job $jobId -> $jobDir"
if ($RunLevel -ne 'Highest') { Write-Host "run level: $RunLevel - the work runs NOT elevated, and refuses (exit 4) if its token is elevated anyway" }
Start-ScheduledTask -TaskName $TaskName

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $exitPath) { break }
    Start-Sleep -Seconds 2
}

if (-not (Test-Path -LiteralPath $exitPath)) {
    $state = 'unknown'
    $info = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($info) { $state = $info.State }
    Write-Warning "No exit.txt after $TimeoutSeconds s. Task state: $state. Job kept at $jobDir."
    if (Test-Path -LiteralPath $outPath) {
        Write-Host '--- partial output ---'
        Get-Content -LiteralPath $outPath
    }
    else {
        Write-Warning 'out.txt does not exist either, so the work never started writing. Check that the'
        Write-Warning 'guest has an interactive session at all: Get-Process explorer, or query user.'
    }
    exit 3
}

$code = (Get-Content -LiteralPath $exitPath -Raw).Trim()
if (Test-Path -LiteralPath $outPath) { Get-Content -LiteralPath $outPath }

Write-Host ''
Write-Host "exit $code   (job $jobId)"

if (-not $KeepJob -and $code -eq '0') {
    Remove-Item -LiteralPath $jobDir -Recurse -Force -ErrorAction SilentlyContinue
}
else {
    Write-Host "job directory kept at $jobDir"
}

exit [int]$code
