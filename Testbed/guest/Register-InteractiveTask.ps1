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

        <JobRoot>\<jobId>\cmd.ps1     the work. Written by this script from -ScriptPath or -Script.
        <JobRoot>\<jobId>\out.txt     everything the work wrote, stdout and stderr merged.
        <JobRoot>\<jobId>\exit.txt    the exit code. ITS EXISTENCE IS THE COMPLETION SIGNAL.

    exit.txt is written last, in a `finally`, so it exists even when the work throws. A job that
    has out.txt and no exit.txt is still running or died without unwinding - those two are
    genuinely different and this script says which by whether the deadline was reached.

    NO WINDOW, EVER. `-WindowStyle Hidden` is passed to powershell.exe, where it hides that
    process's own console; the task settings ask for hidden as well. Do not "fix" a problem here
    with `Start-Process -Verb RunAs`: UAC elevation takes the foreground even with a hidden
    window, which was measured on this project and is why that verb is banned outright.

    RESIDUAL RISK, stated rather than hidden: a task in an interactive session is not the same as
    a windowless service, and a child process that creates its own window will draw one on the
    guest's console. Nothing in the testbed should - the MCP server and the tools are console
    apps started with CreateNoWindow - but if you see a flash on the guest, this is where to look.

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

.EXAMPLE
    .\Register-InteractiveTask.ps1 -ScriptPath C:\OutlookAI-Q5\guest-measure.ps1
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
    [switch] $KeepJob
)

$ErrorActionPreference = 'Stop'

if (-not $UserId) { $UserId = "$env:USERDOMAIN\$env:USERNAME" }

$jobId = [guid]::NewGuid().ToString('N')
$jobDir = Join-Path $JobRoot $jobId
New-Item -ItemType Directory -Force -Path $jobDir | Out-Null

$cmdPath = Join-Path $jobDir 'cmd.ps1'
$outPath = Join-Path $jobDir 'out.txt'
$exitPath = Join-Path $jobDir 'exit.txt'

if ($PSCmdlet.ParameterSetName -eq 'Path') {
    if (-not (Test-Path -LiteralPath $ScriptPath)) { throw "Not found on the guest: $ScriptPath" }
    $body = "& '" + ($ScriptPath -replace "'", "''") + "'"
}
else {
    $body = $Script
}

# The wrapper, not the work. It exists to guarantee three things the work cannot guarantee about
# itself: everything it writes reaches out.txt, an exit code is recorded even when it throws, and
# exit.txt is written LAST so its existence really does mean "finished".
$wrapper = @"
`$ErrorActionPreference = 'Continue'
`$code = 0
try {
    & {
$body
    } *>&1 | Out-File -FilePath '$outPath' -Encoding utf8 -Append
    if (`$LASTEXITCODE -ne `$null) { `$code = `$LASTEXITCODE }
}
catch {
    `$code = 1
    "WRAPPER CAUGHT: `$(`$_.Exception.GetType().Name): `$(`$_.Exception.Message)" | Out-File -FilePath '$outPath' -Encoding utf8 -Append
}
finally {
    "`$code" | Out-File -FilePath '$exitPath' -Encoding utf8
}
"@
Set-Content -LiteralPath $cmdPath -Value $wrapper -Encoding UTF8

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false }

# -WindowStyle Hidden here is an argument to powershell.exe, which honours it for its own
# console. That is NOT the same as putting it inside a -ArgumentList handed to Start-Process,
# where it is silently a no-op - a distinction this project has been bitten by.
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$cmdPath`""

# Interactive is the entire point: it puts the process in the logged-on session, where Outlook
# can actually finish starting. RunLevel Highest because Outlook COM and the add-in registry
# reads want it; it prompts for nothing, because a task's elevation is granted at registration.
$principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet -Hidden `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::FromSeconds([Math]::Max($TimeoutSeconds, 60) + 300)) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings | Out-Null

Write-Host "job $jobId -> $jobDir"
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
