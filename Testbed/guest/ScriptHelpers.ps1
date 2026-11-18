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
