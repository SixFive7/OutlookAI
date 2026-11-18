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
# kind, a probe's scenario - and the first 12 characters of the commit. Sortable, unique per second and
# commit, and readable in a directory listing. Anything but a commit id is refused.
function New-RunId {
    param([DateTime] $When, [string] $Sha, [string] $Label = '')
    if ($Sha -notmatch '^[0-9a-f]{12,40}$') { throw "Not a commit id: '$Sha'." }
    if ($Label) { return ('{0:yyyyMMdd-HHmmss}-{1}-{2}' -f $When, $Label, $Sha.Substring(0, 12)) }
    return ('{0:yyyyMMdd-HHmmss}-{1}' -f $When, $Sha.Substring(0, 12))
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

# A test VM's lease, as Set-TestbedLease.ps1 writes it, judged for one run: 'free' when there is none,
# it has expired or its expiry cannot be read (Testbed/host/TestbedLeasePath.ps1's rule); 'ours' when
# its reason is exactly the run's own; 'theirs' otherwise, which a run never takes over.
function Get-LeaseDecision {
    param([object] $Lease, [string] $Reason, [long] $NowUnix)
    if ($null -eq $Lease) { return 'free' }
    $expires = 0L
    try { $expires = [long]$Lease.expiresUnix } catch { return 'free' }
    if ($expires -le $NowUnix) { return 'free' }
    if (([string]$Lease.reason) -ceq $Reason) { return 'ours' }
    return 'theirs'
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

# ---- The Outlook profile scripts' helpers (Q154, S7c step 3): New-OutlookProfile.ps1,
# Set-DefaultOutlookProfile.ps1, New-TierProfile.ps1 and Build-Corpus.ps1 each held a copy. The constants
# the first three read are the calling script's own, as $script: variables.

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
