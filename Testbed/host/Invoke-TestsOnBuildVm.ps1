#Requires -Version 5.1
<#
.SYNOPSIS
    Tests one revision on the build VM, OutlookAI-Build: the non-live suite and every script
    self-test, on a machine restored to the same checkpoint before every run, and brings the
    results back - a summary, the TRX file and every log under .work\. Its exit code is the
    verdict. THE route by which non-live tests and self-tests run (Q94, Q102, 2026-10-03).

.DESCRIPTION
    RUN ON THE HOST. Windows PowerShell 5.1 or PowerShell 7. Needs local Hyper-V Administrators
    membership, not elevation, and the guest credential (below). Never runs a test on this machine:
    it archives a commit here and everything else happens on OutlookAI-Build.

    WHY. The maintainer's workstation runs no test any more except the Exchange-only read-only
    live tests (Q94), and the two Outlook test VMs are often busy for hours, so the non-live suite
    and the script self-tests run on a small VM of their own (Q102 (b)). Agents call this for every
    change, so it is built to answer in minutes and to answer the same way twice.

    WHAT ONE RUN DOES
      1. Resolves -Ref in -RepoPath to ONE COMMIT and archives it with `git archive` - a commit,
         never a working tree, so every result names a revision anybody can check out. A HEAD with
         uncommitted changes is tested as committed, and the run says how many changes it left out.
      2. Takes the runner lock (queues behind any run already going - see CONCURRENT CALLERS).
      3. Takes a lease on the VM (Testbed/README.md section 5b), so the idle-saver keeps away.
      4. Restores the base checkpoint and resumes the VM from it, waits for its heartbeat and
         PowerShell Direct, proves it is OAI-BUILD logged on as the autologon account, and waits
         for its clock to agree with this host's (a resumed guest starts at its checkpoint's time).
      5. Copies in the archive, Testbed/guest/Invoke-BuildVmRun.ps1 and a request file, starts the
         guest script detached, and follows its log until it writes done.txt. That script restores
         from the VM's offline feed, builds, runs `dotnet test --filter Category!=Live` (ANDed in
         code with -Filter) into a TRX file, and runs every -SelfTest it finds under Testbed\ and
         Tools\.
      6. If the revision needs a NuGet package the VM's feed lacks, stages that revision's packages
         here with Testbed/host/Publish-LiveTierPayload.ps1, adds them to the VM's feed and runs
         step 5 again - once.
      7. Fetches the results - the guest's zip of them, or, when it left none, its log, run.json
         and TRX file one at a time, each read with sharing - restores the base checkpoint again,
         which leaves the VM SAVED, holding no RAM, releases the lease and the lock, and writes
         summary.txt and summary.json.

    FRESH STATE: A CHECKPOINT PER RUN, NOT A WORK DIRECTORY RESET - decided here, and why. Every
    run starts from the base checkpoint, restored by Hyper-V, and the run's changes are thrown away
    the same way afterwards. Resetting a directory instead would be no faster - resuming a saved
    VM is the same few seconds either way - and would leave everything a test run writes OUTSIDE
    that directory to the next run: the NuGet and temp folders, HKCU, a test host that did not
    exit, a half-finished background task. A run would then depend on the runs before it, which is
    exactly what an agent calling this for every change cannot see. With a checkpoint, two runs of
    one commit start from byte-identical machines; the only inputs are the archive and the request.
    Restoring at the END as well is what returns the RAM: the base checkpoint was taken running, so
    restoring it leaves the VM in its saved state - "started on demand, saved when idle" without a
    separate save. Reverting discards a running VM's state the way Hyper-V's own Apply does; it is
    not a restart and nothing in that state is kept, so there is no shutdown to make graceful - the
    machine holds no Office, no PST and nothing a run must keep. Restarts of this VM, which only a
    rebuild needs, go through Testbed/host/Restart-Guest.ps1 like any guest's.

    CONCURRENT CALLERS ARE SERIALISED. One VM, one run at a time: a lock file opened exclusively
    under %SystemDrive%\OutlookAI-Testbed\build-vm\ (machine-wide, beside the leases). A second
    caller waits in the queue, printing who holds the VM every minute, for up to
    -QueueTimeoutMinutes. The lock is an open file handle, so a caller that dies - killed, crashed,
    its terminal closed - releases it with its process, and the next run restores the checkpoint
    whatever state it finds.

    A CALLER THAT DIES DOES NOT LEAVE THE VM HOLDING ITS RAM. Measured 2026-10-03: a run stopped
    part-way left the VM running, 6 GB held, for three hours, until the next run came - nothing on
    this host saves it, the idle-saver's task not being registered here, and an agent's tool call is
    exactly the kind of caller that gets stopped on a time limit. So every run, once it holds the
    lock, starts a JANITOR: this same file with -JanitorForPid, launched through WMI's
    Win32_Process.Create so that it is no child of the caller and outlives whatever stops it - a
    killed process tree, a closed terminal - and hidden. It waits for the run's process to end. A
    run that finished wrote its end line to runs.log, and the janitor just exits; one that did not
    is put right: the janitor takes the lock (behind any run that took over first, which restores
    the checkpoint itself), restores the base checkpoint, saves the VM and releases the dead run's
    lease, and says so in janitor.log beside the lock. It never touches the VM without the lock.
    A LEASE someone else holds on the VM (Set-TestbedLease.ps1,
    for maintenance by hand) also holds runs back until it is released or expires. Running several
    VMs in parallel was not chosen: the suite runs its collections one at a time and finishes in
    minutes, and a second VM would double what the host keeps on disk for a queue that is rarely
    long.

    CREDENTIALS only through Testbed/host/Get-GuestCredential.ps1, from the main checkout - found
    with `git rev-parse --git-common-dir`, so this runs from any worktree - or -CredentialRepoRoot.
    Never printed, never written.

    ONLY OutlookAI-Build. The VM name is a constant, not a parameter - Testbed/README.md section 4a
    forbids a default that picks one of several machines, and there is nothing here to pick: this
    restores checkpoints, and on any other VM that would destroy its state. Every Hyper-V call goes
    through the VM object Get-BuildVm returns after checking that name; the computer name inside
    must read OAI-BUILD before anything is copied in. -SelfTest pins both.

    EXIT CODES
        0  PASS      the revision built, every selected test and every self-test passed
        1  FAIL      a test or a self-test failed, the suite hung or crashed, or the filter
                     selected no test at all
        2  BUILD     the revision did not restore or build
        3  INFRA     the revision was NOT tested - no result came back: the VM, the lock, the copy
                     or a time limit failed. An error on this side AFTER the guest's results came
                     back does not make a run INFRA: the results decide, and the error is a note
                     beside them. Nor does a run that stopped part-way once its TRX file showed a
                     failing test: that is FAIL (2026-10-03 - Get-HostVerdict says why)
        4  REFUSED   bad arguments, a ref that is not a commit, no credential

.PARAMETER Ref
    What to test: a commit, a branch, a tag, HEAD - anything `git rev-parse` resolves to a commit
    in -RepoPath. Default HEAD. Positional.

.PARAMETER RepoPath
    The repository or worktree -Ref is resolved in, and whose .work\ receives the results. Default:
    the checkout this script is in - so an agent runs the copy in its own worktree and tests its
    own HEAD with no arguments at all.

.PARAMETER Filter
    A narrower dotnet test filter, ANDed with Category!=Live in the guest script's code - it can
    narrow the run, never widen it into the live tier. Example: 'FullyQualifiedName~T1.SweepSortWiringTests'.

.PARAMETER SkipSuite
    Run only the self-tests.

.PARAMETER SkipSelfTests
    Run only the suite.

.PARAMETER SelfTestInclude
    Run only the self-tests of scripts matching these wildcards on repository-relative paths, for
    example 'Testbed/host/Restart-Guest.ps1' or 'Testbed/guest/*'. Several, comma-separated in one
    string: `pwsh -File` hands 'a','b' over as the single string "a,b" (measured 2026-10-03 - the
    run matched nothing and, rightly, failed as "nothing was tested"), so every value is split on
    commas here.

.PARAMETER ResultsRoot
    Where run directories go. Default <RepoPath>\.work\build-vm-runs, gitignored.

.PARAMETER CredentialRepoRoot
    The checkout holding McpServer\OutlookAI.McpServer.Tests\live-fixtures\vm-credentials.json.
    Default: the main checkout of -RepoPath's repository.

.PARAMETER QueueTimeoutMinutes
    How long to wait for a run already holding the VM. Default 120.

.PARAMETER RunTimeoutMinutes
    How long the guest half may take before the run is abandoned as INFRA. Default 60.

.PARAMETER KeepVmRunning
    Leave the VM running after the run, with its state, for inspection over PowerShell Direct.
    The next run restores the checkpoint regardless; release nothing by hand.

.PARAMETER JanitorForPid
    Internal: what a run starts its janitor with (see A CALLER THAT DIES). The process id of the
    run to watch. Not for use by hand.

.PARAMETER JanitorRunId
    Internal: the id of the run the janitor watches.

.PARAMETER SelfTest
    The run-id format, the TRX reader, the verdict table, summary.json's shape, the VM-name
    allowlist, that every Hyper-V call targets the build VM's variables and none stops, removes or
    checkpoints a VM, that this file, the guest script and Testbed/testbed.json agree on the VM's
    names and paths, and - on scratch files under %TEMP% that it locks itself and removes after -
    that a held guest log reads as Busy and never throws. No Hyper-V, no git, no VM, no
    credential: it runs on the build VM's own self-test pass.

.EXAMPLE
    # From an agent's worktree: test its HEAD, the whole non-live suite and every self-test.
    pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1

.EXAMPLE
    pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -Ref origin/master
    pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -Ref 1a2b3c4 -Filter 'FullyQualifiedName~T1.Sweep' -SkipSelfTests
    pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -SkipSuite -SelfTestInclude 'Testbed/host/*'

.EXAMPLE
    pwsh -File Testbed/host/Invoke-TestsOnBuildVm.ps1 -SelfTest
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run', Position = 0)] [string] $Ref = 'HEAD',
    [Parameter(ParameterSetName = 'Run')] [string]   $RepoPath,
    [Parameter(ParameterSetName = 'Run')] [string]   $Filter = '',
    [Parameter(ParameterSetName = 'Run')] [switch]   $SkipSuite,
    [Parameter(ParameterSetName = 'Run')] [switch]   $SkipSelfTests,
    [Parameter(ParameterSetName = 'Run')] [string[]] $SelfTestInclude = @(),
    [Parameter(ParameterSetName = 'Run')] [string]   $ResultsRoot,
    [Parameter(ParameterSetName = 'Run')] [string]   $CredentialRepoRoot,
    [Parameter(ParameterSetName = 'Run')] [int]      $QueueTimeoutMinutes = 120,
    [Parameter(ParameterSetName = 'Run')] [int]      $RunTimeoutMinutes = 60,
    [Parameter(ParameterSetName = 'Run')] [switch]   $KeepVmRunning,
    [Parameter(Mandatory = $true, ParameterSetName = 'Janitor')] [int]    $JanitorForPid,
    [Parameter(Mandatory = $true, ParameterSetName = 'Janitor')] [string] $JanitorRunId,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# ---------------------------------------------------------------------------------------------
# The build VM, as constants. -SelfTest holds them equal to Testbed/testbed.json's buildVm block
# and to the guest script's defaults.
# ---------------------------------------------------------------------------------------------
$BuildVmName = 'OutlookAI-Build'
$BuildVmComputerName = 'OAI-BUILD'
$BuildVmUser = 'vmadmin'
$BaseCheckpointName = 'CP-02-SDK-TEST-READY'
$GuestRunRoot = 'C:\OutlookAI-Q5\run'
$GuestFeedRoot = 'C:\OutlookAI-Q5\nuget-offline'
$GuestScriptName = 'Invoke-BuildVmRun.ps1'
$LeaseReasonPrefix = 'Invoke-TestsOnBuildVm'
$MaxClockSkewSeconds = 2

# Exit codes - see the banner.
$ExitCodes = [ordered]@{ PASS = 0; FAIL = 1; BUILD = 2; INFRA = 3; REFUSED = 4 }

# What the host asks the guest every few seconds while a run goes: the new whole lines of the guest
# script's log, whether that script is alive, and whether it has finished. Sent over PowerShell
# Direct as a script block, and run locally by -SelfTest against a file it locks itself.
# A MOMENT WHEN THE LOG CANNOT BE OPENED IS 'BUSY', NEVER AN ERROR. On 2026-10-03 the host's read
# hit the guest script's own append - Windows PowerShell 5.1's Add-Content opens a file so that
# nobody else may read it while it writes (measured) - and the exception ended a run whose suite
# had already finished as INFRA. The guest now appends with sharing, and this still never throws
# on a busy log: the next poll reads what this one could not. FileShare.ReadWrite on this side so
# the guest's writer is never shut out either.
$PollGuestBlock = {
    param($log, $from, $procId, $done)
    $text = ''
    $next = $from
    $busy = $false
    if (Test-Path -LiteralPath $log) {
        try {
            $fs = [System.IO.File]::Open($log, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                if ($fs.Length -gt $from) {
                    [void]$fs.Seek($from, [System.IO.SeekOrigin]::Begin)
                    $buffer = New-Object byte[] ([int]($fs.Length - $from))
                    $read = $fs.Read($buffer, 0, $buffer.Length)
                    $end = -1
                    if ($read -gt 0) { $end = [Array]::LastIndexOf($buffer, [byte]10, $read - 1) }
                    if ($end -ge 0) {
                        $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $end + 1)
                        $next = $from + $end + 1
                    }
                }
            }
            finally { $fs.Dispose() }
        }
        catch [System.IO.IOException] { $busy = $true }
    }
    $isDone = Test-Path -LiteralPath $done
    $code = $null
    if ($isDone) {
        try { $code = ([System.IO.File]::ReadAllText($done)).Trim() }
        catch [System.IO.IOException] { $isDone = $false; $busy = $true }
        # There but empty: the guest's Set-Content between creating the file and writing to it. Not
        # finished until the exit code can be read - a '20' read as '' would skip the package retry.
        if ($isDone -and -not $code) { $isDone = $false; $busy = $true }
    }
    [pscustomobject]@{
        Text  = $text.TrimStart([char]0xFEFF)
        Next  = $next
        Alive = ($null -ne (Get-Process -Id $procId -ErrorAction SilentlyContinue))
        Done  = $isDone
        Code  = $code
        Busy  = $busy
    }
}

# Only whether the guest script lives and has finished: what Wait-GuestDone needs, and nothing that
# a held log could stand in the way of.
$GuestStateBlock = {
    param($procId, $done)
    [pscustomobject]@{
        Alive = ($null -ne (Get-Process -Id $procId -ErrorAction SilentlyContinue))
        Done  = (Test-Path -LiteralPath $done)
    }
}

# One of the guest's result files as it stands, for a run that left no results.zip - its log, its
# run.json, the suite's TRX file - read with sharing and retried, for the same reason as above, and
# $null rather than an error when it is not there or cannot be read at all.
$ReadGuestLogBlock = {
    param($path)
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                return $reader.ReadToEnd()
            }
            finally { $fs.Dispose() }
        }
        catch [System.IO.IOException] { Start-Sleep -Milliseconds (200 * $attempt) }
    }
    return $null
}

# A NATIVE PROGRAM RUNS THROUGH HERE (Q78). Under $ErrorActionPreference = 'Stop', Windows
# PowerShell 5.1 turns the first line a native program writes to stderr into a terminating
# NativeCommandError when anything redirects it - and a script under Testbed/ runs behind a
# redirection or in a remoting host as often as not, so every program call here goes through this:
#   * 'Continue' holds in THIS function's scope only; the caller's 'Stop' is never changed.
#   * The try is load-bearing: without it 'Continue' demotes "the program was not found" to a
#     printed message and the caller reads a stale $LASTEXITCODE.
#   * Stderr lines come back as plain strings in both shells, never as ErrorRecords.
#   * The exit code is left in $LASTEXITCODE, and the caller checks it.
# Restated in each script that needs it, as this repository restates its shared rules.
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

# =============================================================================================
# PURE DECISIONS. No Hyper-V, no git, no file but the ones -SelfTest names. -SelfTest pins them.
# =============================================================================================

# The only VM this script will act on. Compared whole and ordinal-ignore-case, as Hyper-V compares
# names; a prefix, a suffix or another testbed guest is refused.
function Test-IsBuildVmName {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    return [string]::Equals($Name.Trim(), $BuildVmName, [System.StringComparison]::OrdinalIgnoreCase) -and ($Name -eq $Name.Trim())
}

# A run directory's name: when it started, and which commit. Sortable, unique per second and
# commit, and readable in a directory listing.
function New-RunId {
    param([DateTime] $When, [string] $Sha)
    if ($Sha -notmatch '^[0-9a-f]{12,40}$') { throw "Not a commit id: '$Sha'." }
    return ('{0:yyyyMMdd-HHmmss}-{1}' -f $When, $Sha.Substring(0, 12))
}

# Whether a string is a run id New-RunId could have made - what the janitor is handed on its
# command line, and what it then searches runs.log for, so nothing else is accepted.
function Test-IsRunId {
    param([string] $Text)
    return ([string]$Text) -cmatch '^\d{8}-\d{6}-[0-9a-f]{12}$'
}

# Whether runs.log says a run finished: its own end line, written after the lock is released and
# the summary is written - the one thing a janitor needs to know to stand down.
function Test-RunEnded {
    param([string] $History, [string] $RunId)
    foreach ($line in ([string]$History -split "`r?`n")) {
        if ($line -cmatch ('^\S+ \S+  end    ' + [regex]::Escape($RunId) + '  ')) { return $true }
    }
    return $false
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

# The host's verdict, from what came back. WHEREVER THE GUEST'S run.json AND THE TRX FILE CAME
# BACK, THEY DECIDE: an error on this side - a log the host could not read for a moment, a poll that
# failed, a fetch that had to retry - does not turn a run that was tested into "not tested". On
# 2026-10-03 a run whose suite had finished - 238 tests, one failing - came back INFRA because the
# host's read of the guest's log hit the guest script's own write, and a caller reading the exit
# code took a finished, failing run for an untested one. INFRA only when no verdict came back at
# all; the error is then the reason, and otherwise a note beside the verdict. The TRX file decides
# what the guest cannot see: a run whose filter selected nothing passes in the guest and fails
# here, because "no test failed" is not "the tests passed". And a TRX file that came back from a
# run that did NOT finish still counts for what it shows failing: a test that failed failed,
# whatever happened after it. It cannot show a pass - the rest of that run was never seen.
function Get-HostVerdict {
    param(
        [string] $GuestVerdict,     # the guest script's verdict from run.json, or '' when none came back
        [string] $InfraError,       # set when something on the host side went wrong during the run
        [object] $Trx,              # ConvertFrom-TrxText's result, or $null
        [bool]   $SuiteRequested,
        [int]    $SelfTestsRun,
        [int]    $SelfTestsFailed
    )
    if (-not $GuestVerdict) {
        $why = 'the guest script left no verdict - no run.json came back'
        if ($InfraError) { $why = $InfraError }
        if ($SuiteRequested -and $null -ne $Trx -and ($Trx.Counts['failed'] + $Trx.Counts['other']) -gt 0) {
            return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$($Trx.Counts['failed'] + $Trx.Counts['other']) test(s) did not pass, by the suite's TRX file - the run did not finish: $why" }
        }
        return [pscustomobject]@{ Verdict = 'INFRA'; Why = $why }
    }
    switch ($GuestVerdict) {
        'REFUSED' { return [pscustomobject]@{ Verdict = 'REFUSED'; Why = 'the guest script refused the request' } }
        'PACKAGES-MISSING' {
            $why = 'the offline feed still lacks packages after staging them'
            if ($InfraError) { $why = "the packages this revision needs could not be staged: $InfraError" }
            return [pscustomobject]@{ Verdict = 'INFRA'; Why = $why }
        }
        'RESTORE-FAILED' { return [pscustomobject]@{ Verdict = 'BUILD'; Why = 'the restore failed' } }
        'BUILD-FAILED' { return [pscustomobject]@{ Verdict = 'BUILD'; Why = 'the build failed' } }
        'SUITE-BROKEN' { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the suite hung, crashed or left no TRX file' } }
        'FAIL' { }
        'PASS' { }
        default { return [pscustomobject]@{ Verdict = 'INFRA'; Why = "the guest script left a verdict this runner does not know ('$GuestVerdict')" } }
    }
    if ($SuiteRequested) {
        if ($null -eq $Trx) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the suite ran and no TRX file came back' } }
        if ($Trx.Counts['total'] -eq 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the filter selected no test - nothing was tested' } }
        if ($Trx.Counts['failed'] -gt 0 -or $Trx.Counts['other'] -gt 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$($Trx.Counts['failed'] + $Trx.Counts['other']) test(s) did not pass" } }
    }
    if ($SelfTestsFailed -gt 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$SelfTestsFailed self-test(s) failed" } }
    if ($GuestVerdict -eq 'FAIL') { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the guest script reported a failure' } }
    if (-not $SuiteRequested -and $SelfTestsRun -eq 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'no self-test was selected - nothing was tested' } }
    return [pscustomobject]@{ Verdict = 'PASS'; Why = '' }
}

# -SelfTestInclude's patterns, one each: every value split on commas, trimmed, empties dropped.
# `pwsh -File` hands a list written 'a','b' over as the one string "a,b" (the parameter's help).
function Split-IncludeList {
    param([string[]] $Values)
    return @($Values | ForEach-Object { ([string]$_) -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# What summary.json says, for a caller that reads it rather than summary.txt. EVERY LIST IS A JSON
# ARRAY AT EVERY LENGTH. Built inside $( ), PowerShell unrolled them: the 2026-10-03 run's one
# failed test came out as an object and its no skipped tests as null, so a caller reading .failed[]
# - or counting it under Windows PowerShell 5.1, which gives a lone object no Count - broke on the
# commonest failing run there is. The verdict fields are Get-HostVerdict's; the figures are the
# TRX file's and run.json's, the two things that verdict is decided from.
function New-RunSummary {
    param(
        [string]   $RunId,
        [object]   $Verdict,        # Get-HostVerdict's result
        [System.Collections.IDictionary] $Revision,
        [string]   $Filter,         # '' when the suite was skipped
        [object]   $Trx,            # ConvertFrom-TrxText's result, or $null
        [object[]] $SelfTests,      # run.json's selfTests
        [object[]] $SelfTestSkips,  # run.json's selfTestSkips
        [string]   $GuestVerdict,
        [string]   $HostError,
        [object[]] $BuildErrors,
        [bool]     $PackagesStaged,
        [System.Collections.IDictionary] $Timings,
        [System.Collections.IDictionary] $GuestPhases,
        [string]   $Results
    )
    $failed = @()
    $skipped = @()
    $suite = $null
    if ($null -ne $Trx) { $failed = @($Trx.Failed); $skipped = @($Trx.Skipped); $suite = $Trx.Counts }
    $ran = @($SelfTests | Where-Object { $null -ne $_ })
    $failures = @($ran | Where-Object { -not $_.passed } | ForEach-Object { [string]$_.path })
    $filterValue = $null
    if ($Filter) { $filterValue = $Filter }
    return [ordered]@{
        runId          = $RunId
        verdict        = $Verdict.Verdict
        why            = $Verdict.Why
        exitCode       = $ExitCodes[$Verdict.Verdict]
        revision       = $Revision
        filter         = $filterValue
        suite          = $suite
        failed         = $failed
        skipped        = $skipped
        selfTests      = [ordered]@{ run = $ran.Count; failed = $failures.Count; skippedByReason = @($SelfTestSkips | Where-Object { $null -ne $_ }).Count; failures = $failures }
        guestVerdict   = $GuestVerdict
        hostError      = $HostError
        buildErrors    = @($BuildErrors | Where-Object { $null -ne $_ })
        packagesStaged = $PackagesStaged
        timings        = $Timings
        guestPhases    = $GuestPhases
        results        = $Results
    }
}

function Format-Seconds {
    param([double] $Seconds)
    if ($Seconds -lt 60) { return ('{0:N0} s' -f $Seconds) }
    return ('{0} m {1:00} s' -f [int][Math]::Floor($Seconds / 60), [int]($Seconds % 60))
}

function Invoke-SelfTest {
    $script:stChecks = 0
    $script:stFailures = @()
    function Check([string] $What, $Expected, $Actual) {
        $script:stChecks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = $Expected -join ' | ' }
        if ($Actual -is [System.Array]) { $a = $Actual -join ' | ' }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else { $script:stFailures += "$What : expected [$e], got [$a]"; Write-Host "  FAIL $What - expected [$e], got [$a]" }
    }

    Write-Host "Invoke-TestsOnBuildVm.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No Hyper-V, no git, no VM, no credential."
    Write-Host ''
    Write-Host '== the only VM it acts on =='
    Check 'OutlookAI-Build is the build VM' $true (Test-IsBuildVmName 'OutlookAI-Build')
    Check 'in any case, as Hyper-V compares names' $true (Test-IsBuildVmName 'outlookai-build')
    Check 'OutlookAI-Indexed is not' $false (Test-IsBuildVmName 'OutlookAI-Indexed')
    Check 'OutlookAI-Unindexed is not' $false (Test-IsBuildVmName 'OutlookAI-Unindexed')
    Check 'OutlookAI-TestVM is not' $false (Test-IsBuildVmName 'OutlookAI-TestVM')
    Check 'a name that merely starts the same way is not' $false (Test-IsBuildVmName 'OutlookAI-Build2')
    Check 'nor one padded with a space' $false (Test-IsBuildVmName 'OutlookAI-Build ')
    Check 'nor an empty one' $false (Test-IsBuildVmName '')

    Write-Host ''
    Write-Host '== run ids =='
    Check 'a run id is the time and twelve hex of the commit' '20261003-041500-ea40cc831e68' (New-RunId -When (New-Object DateTime(2026, 10, 3, 4, 15, 0)) -Sha 'ea40cc831e68ae3e4f09602309bbc98569394979')
    $threw = $false
    try { [void](New-RunId -When (Get-Date) -Sha 'HEAD') } catch { $threw = $true }
    Check 'a ref that is not a commit id is refused' $true $threw
    Check 'a run id New-RunId makes is one the janitor accepts' $true (Test-IsRunId (New-RunId -When (Get-Date) -Sha 'ea40cc831e68ae3e4f09602309bbc98569394979'))
    Check 'the janitor accepts nothing else' $false (Test-IsRunId '20261003-041500-ea40cc831e68; Remove-Item')
    Check 'not an upper-case one either' $false (Test-IsRunId '20261003-041500-EA40CC831E68')
    $history = @(
        '2026-10-03 04:19:17  start  20261003-041915-ff55a603e9c2  HEAD  C:\x'
        '2026-10-03 04:23:15  end    20261003-041915-ff55a603e9c2  PASS  3 m 59 s'
        '2026-10-03 04:34:54  start  20261003-043453-01058213466b  HEAD  C:\x'
    ) -join "`r`n"
    Check 'a run with its end line has ended' $true (Test-RunEnded -History $history -RunId '20261003-041915-ff55a603e9c2')
    Check 'a run with only its start line has not' $false (Test-RunEnded -History $history -RunId '20261003-043453-01058213466b')
    Check 'a run the log never mentions has not' $false (Test-RunEnded -History $history -RunId '20261003-050000-0123456789ab')

    Write-Host ''
    Write-Host '== the TRX reader =='
    $trxText = @'
<?xml version="1.0" encoding="utf-8"?>
<TestRun id="1" name="x" xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
  <Results>
    <UnitTestResult testName="T1.A.Passes" outcome="Passed" />
    <UnitTestResult testName="T1.A.Fails" outcome="Failed"><Output><ErrorInfo><Message>Assert.Equal() Failure
Expected: 1</Message></ErrorInfo></Output></UnitTestResult>
    <UnitTestResult testName="T1.A.Skips" outcome="NotExecuted" />
    <UnitTestResult testName="T1.A.Times" outcome="Timeout" />
  </Results>
  <ResultSummary outcome="Failed"><Counters total="4" executed="3" passed="1" failed="1" notExecuted="1" /></ResultSummary>
</TestRun>
'@
    $trx = ConvertFrom-TrxText $trxText
    Check 'every result is counted' 4 $trx.Counts['total']
    Check 'passed' 1 $trx.Counts['passed']
    Check 'failed' 1 $trx.Counts['failed']
    Check 'a NotExecuted result is a skip' 1 $trx.Counts['skipped']
    Check 'any other outcome is counted, not lost' 1 $trx.Counts['other']
    Check 'a failure carries the first line of its message' 'Assert.Equal() Failure' (@($trx.Failed | Where-Object { $_.name -eq 'T1.A.Fails' })[0].message)
    Check 'an outcome that is neither pass nor fail is listed with the failures' 'outcome Timeout' (@($trx.Failed | Where-Object { $_.name -eq 'T1.A.Times' })[0].message)
    Check 'the file''s own counters are kept beside the count' 4 $trx.Declared['total']
    Check 'the skipped test is named' 'T1.A.Skips' (@($trx.Skipped) -join ',')

    Write-Host ''
    Write-Host '== the verdict =='
    $green = ConvertFrom-TrxText ($trxText -replace 'outcome="Failed"', 'outcome="Passed"' -replace 'outcome="Timeout"', 'outcome="Passed"')
    $empty = ConvertFrom-TrxText '<TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010"><Results /></TestRun>'
    Check 'a green suite and green self-tests is PASS' 'PASS' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError '' -Trx $green -SuiteRequested $true -SelfTestsRun 5 -SelfTestsFailed 0).Verdict
    Check 'a failed test is FAIL even if the guest said PASS' 'FAIL' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError '' -Trx $trx -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'a filter that selected nothing is FAIL, not PASS' 'FAIL' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError '' -Trx $empty -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'a requested suite with no TRX is FAIL' 'FAIL' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'a failed self-test is FAIL' 'FAIL' (Get-HostVerdict -GuestVerdict 'FAIL' -InfraError '' -Trx $green -SuiteRequested $true -SelfTestsRun 5 -SelfTestsFailed 1).Verdict
    Check 'self-tests only, none selected, is FAIL' 'FAIL' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError '' -Trx $null -SuiteRequested $false -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'self-tests only, all green, is PASS' 'PASS' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError '' -Trx $null -SuiteRequested $false -SelfTestsRun 3 -SelfTestsFailed 0).Verdict
    Check 'a build failure is BUILD' 'BUILD' (Get-HostVerdict -GuestVerdict 'BUILD-FAILED' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'a restore failure is BUILD' 'BUILD' (Get-HostVerdict -GuestVerdict 'RESTORE-FAILED' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'a hung or crashed suite is FAIL' 'FAIL' (Get-HostVerdict -GuestVerdict 'SUITE-BROKEN' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'packages still missing after staging is INFRA - the revision was not tested' 'INFRA' (Get-HostVerdict -GuestVerdict 'PACKAGES-MISSING' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'no guest verdict at all is INFRA' 'INFRA' (Get-HostVerdict -GuestVerdict '' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'a host-side error with no result back is INFRA' 'INFRA' (Get-HostVerdict -GuestVerdict '' -InfraError 'the VM did not resume' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'and names the error' 'the VM did not resume' (Get-HostVerdict -GuestVerdict '' -InfraError 'the VM did not resume' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Why
    $lockedLog = 'Exception calling "Open" with "4" argument(s): "The process cannot access the file ''C:\OutlookAI-Q5\run\results\run.log'' because it is being used by another process."'
    Check 'the 2026-10-03 case: a host-side error AFTER the results came back does not hide a failure - FAIL, not INFRA' 'FAIL' (Get-HostVerdict -GuestVerdict 'FAIL' -InfraError $lockedLog -Trx $trx -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'and a run whose results passed is PASS' 'PASS' (Get-HostVerdict -GuestVerdict 'PASS' -InfraError $lockedLog -Trx $green -SuiteRequested $true -SelfTestsRun 5 -SelfTestsFailed 0).Verdict
    Check 'and a build failure is still BUILD' 'BUILD' (Get-HostVerdict -GuestVerdict 'BUILD-FAILED' -InfraError $lockedLog -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'packages that could not be staged are INFRA, naming why' $true ((Get-HostVerdict -GuestVerdict 'PACKAGES-MISSING' -InfraError 'nuget.org did not answer' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Why.Contains('nuget.org did not answer'))
    Check 'a guest refusal is REFUSED' 'REFUSED' (Get-HostVerdict -GuestVerdict 'REFUSED' -InfraError '' -Trx $null -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    $unfinished = 'the guest run did not finish within 60 minute(s)'
    Check 'a run that did not finish, whose TRX came back showing a failure, is FAIL - a test that failed failed' 'FAIL' (Get-HostVerdict -GuestVerdict '' -InfraError $unfinished -Trx $trx -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'and says that the run did not finish, and why' $true ((Get-HostVerdict -GuestVerdict '' -InfraError $unfinished -Trx $trx -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Why.Contains("did not finish: $unfinished"))
    Check 'but one whose TRX came back green is INFRA - the rest of that run was never seen' 'INFRA' (Get-HostVerdict -GuestVerdict '' -InfraError $unfinished -Trx $green -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'and one whose guest died with no word at all, TRX green, is INFRA too' 'INFRA' (Get-HostVerdict -GuestVerdict '' -InfraError '' -Trx $green -SuiteRequested $true -SelfTestsRun 0 -SelfTestsFailed 0).Verdict
    Check 'PASS alone exits 0' 'PASS' (@($ExitCodes.Keys | Where-Object { $ExitCodes[$_] -eq 0 }) -join ',')
    Check 'durations read as minutes and seconds' '4 m 05 s' (Format-Seconds 245)
    Check 'an include list arriving as one comma-joined string is split' 'Testbed/host/a.ps1 | Testbed/guest/*' (Split-IncludeList @('Testbed/host/a.ps1,Testbed/guest/*'))
    Check 'a real list is kept, spaces and empties dropped' 'x | y | z' (Split-IncludeList @('x', ' y ,', 'z'))
    Check 'no list stays no list' 0 @(Split-IncludeList @()).Count

    Write-Host ''
    Write-Host '== summary.json: every list is a JSON array at every length (2026-10-03: one failure came out as an object) =='
    $oneFailure = ConvertFrom-TrxText ($trxText -replace 'outcome="Timeout"', 'outcome="Passed"' -replace 'outcome="NotExecuted"', 'outcome="Passed"')
    $twoFailures = ConvertFrom-TrxText ($trxText -replace 'outcome="Timeout"', 'outcome="Failed"')
    function Get-SummaryJson($Trx, [object[]] $SelfTests, [object[]] $BuildErrors) {
        $v = Get-HostVerdict -GuestVerdict 'FAIL' -InfraError '' -Trx $Trx -SuiteRequested ($null -ne $Trx) -SelfTestsRun @($SelfTests).Count -SelfTestsFailed 0
        $s = New-RunSummary -RunId '20261003-091919-d7e58af90183' -Verdict $v -Revision ([ordered]@{ ref = 'HEAD' }) -Filter 'Category!=Live' -Trx $Trx -SelfTests $SelfTests -SelfTestSkips @() -GuestVerdict 'FAIL' -HostError '' -BuildErrors $BuildErrors -PackagesStaged $false -Timings ([ordered]@{ total = 1.0 }) -GuestPhases ([ordered]@{}) -Results 'C:\x'
        return ($s | ConvertTo-Json -Depth 6)
    }
    $j = Get-SummaryJson $oneFailure @() @()
    Check 'one failed test is an array of one, not an object' $true ($j -match '"failed":\s*\[\s*\{')
    Check 'and reads back as an array' $true ((($j | ConvertFrom-Json).failed) -is [System.Array])
    Check 'no skipped test is an empty array, not null' $true ($j -match '"skipped":\s*\[\s*\]')
    Check 'no failed self-test is an empty array' $true ($j -match '"failures":\s*\[\s*\]')
    Check 'no build error is an empty array' $true ($j -match '"buildErrors":\s*\[\s*\]')
    Check 'its exit code is its verdict''s' 1 (($j | ConvertFrom-Json).exitCode)
    $j = Get-SummaryJson $twoFailures @([pscustomobject]@{ path = 'Testbed/x.ps1'; passed = $false }) @('a.cs(1,1): error CS1002: ; expected')
    Check 'two failed tests are an array of two' 2 @(($j | ConvertFrom-Json).failed).Count
    Check 'one skipped test is an array of one' $true ($j -match '"skipped":\s*\[\s*"T1\.A\.Skips"\s*\]')
    Check 'one failed self-test is an array of one' $true ($j -match '"failures":\s*\[\s*"Testbed/x\.ps1"\s*\]')
    Check 'one build error is an array of one' $true ($j -match '"buildErrors":\s*\[\s*"a\.cs')
    $j = Get-SummaryJson $null @() @()
    Check 'no TRX at all: no suite, and the test lists are still empty arrays' $true (($j -match '"suite":\s*null') -and ($j -match '"failed":\s*\[\s*\]') -and ($j -match '"skipped":\s*\[\s*\]'))

    Write-Host ''
    Write-Host '== the guest log is read with sharing, and a locked one is Busy, never an error =='
    # Real files, in a scratch directory under %TEMP% that is removed afterwards: this is the one
    # place the self-test touches a disk, because what is being pinned is a file-sharing behaviour.
    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-runner-selftest-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $scratch | Out-Null
    try {
        $log = Join-Path $scratch 'run.log'
        $done = Join-Path $scratch 'done.txt'
        [System.IO.File]::WriteAllText($log, "first`r`nsecond`r`npartial", (New-Object System.Text.UTF8Encoding($true)))
        $p = & $PollGuestBlock $log 0 $PID $done
        Check 'whole lines are read, the byte order mark dropped' "first`r`nsecond`r`n" $p.Text
        Check 'the half-written last line is left for the next poll (3 + 7 + 8 bytes read)' 18 $p.Next
        Check 'a live script is Alive, not Done, not Busy' 'True|False|False' ('{0}|{1}|{2}' -f $p.Alive, $p.Done, $p.Busy)
        $p = & $PollGuestBlock $log 18 $PID $done
        Check 'nothing new until that line ends, and the offset stays' '18|' ('{0}|{1}' -f $p.Next, $p.Text)
        # The 2026-10-03 failure, reproduced: the log held so that nobody else may read it - what
        # Windows PowerShell 5.1's Add-Content does while it appends (measured the same day).
        $holder = [System.IO.File]::Open($log, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            $threw = ''
            $p = $null
            try { $p = & $PollGuestBlock $log 0 $PID $done } catch { $threw = $_.Exception.Message }
            Check 'a log nobody else may open does not throw' '' $threw
            Check 'it reads as Busy, with nothing read and the offset kept' 'True|0|' ('{0}|{1}|{2}' -f $p.Busy, $p.Next, $p.Text)
            Check 'and the run is not taken for finished' $false $p.Done
        }
        finally { $holder.Dispose() }
        # What the guest's Say does since 2026-10-03: append through FileShare.ReadWrite.
        $holder = [System.IO.File]::Open($log, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        try {
            $p = & $PollGuestBlock $log 0 $PID $done
            Check 'beside the guest''s own writer, which shares, the log reads - not Busy' 'False|18' ('{0}|{1}' -f $p.Busy, $p.Next)
            Check 'and so does the fallback read' $true (([string](& $ReadGuestLogBlock $log)).Contains('second'))
        }
        finally { $holder.Dispose() }
        [System.IO.File]::WriteAllText($done, '')
        $p = & $PollGuestBlock $log 0 $PID $done
        Check 'a done.txt not yet written to is not the end: Busy, and read again' 'False|True' ('{0}|{1}' -f $p.Done, $p.Busy)
        [System.IO.File]::WriteAllText($done, '1')
        $p = & $PollGuestBlock $log 0 $PID $done
        Check 'done.txt ends the wait, and its exit code is read' 'True|1' ('{0}|{1}' -f $p.Done, $p.Code)
        $s = & $GuestStateBlock $PID $done
        Check 'the state block Wait-GuestDone asks: alive, and done' 'True|True' ('{0}|{1}' -f $s.Alive, $s.Done)
        Check 'and no done.txt is not done' $false (& $GuestStateBlock $PID (Join-Path $scratch 'absent.txt')).Done
        $holder = [System.IO.File]::Open($log, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try { Check 'the fallback read of a log works beside another reader' $true (([string](& $ReadGuestLogBlock $log)).Contains('second')) }
        finally { $holder.Dispose() }
        Check 'the fallback read of a log that is not there is null, not an error' $true ($null -eq (& $ReadGuestLogBlock (Join-Path $scratch 'absent.log')))
    }
    finally { Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue }

    Write-Host ''
    Write-Host '== every Hyper-V call names the build VM, and none stops, removes or checkpoints one =='
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $forbidden = @('Stop-VM', 'Restart-VM', 'Remove-VM', 'Remove-VMSnapshot', 'Remove-VMCheckpoint', 'Checkpoint-VM', 'Suspend-VM', 'Set-VM', 'Set-VMProcessor', 'Set-VMMemory', 'Export-VM', 'Import-VM', 'Rename-VM', 'Stop-Process')
    $vmCommands = @('Get-VM', 'Start-VM', 'Restore-VMSnapshot', 'Get-VMSnapshot', 'Get-VMIntegrationService', 'New-PSSession', 'Save-VM')
    $used = @()
    $loose = @()
    foreach ($c in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $name = $c.GetCommandName()
        if (-not $name) { continue }
        if ($forbidden -contains $name) { $used += "$name (line $($c.Extent.StartLineNumber))" }
        if ($vmCommands -contains $name) {
            $elements = @($c.CommandElements)
            for ($i = 1; $i -lt $elements.Count; $i++) {
                if ($elements[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and @('Name', 'VMName', 'VMId', 'Id') -contains $elements[$i].ParameterName) {
                    $value = ''
                    if ($i + 1 -lt $elements.Count) { $value = $elements[$i + 1].Extent.Text }
                    $ok = ($value -eq '$BuildVmName') -or ($name -eq 'Get-VMSnapshot' -and $elements[$i].ParameterName -eq 'Name' -and $value -eq '$BaseCheckpointName') -or ($name -eq 'Get-VMIntegrationService' -and $elements[$i].ParameterName -eq 'Name')
                    if (-not $ok) { $loose += "$name -$($elements[$i].ParameterName) $value (line $($c.Extent.StartLineNumber))" }
                }
            }
        }
    }
    Check 'no command that stops, removes, reconfigures or checkpoints a VM, and no Stop-Process' '' ($used -join '; ')
    Check 'every -Name / -VMName handed to a VM command is $BuildVmName (or the base checkpoint''s name)' '' ($loose -join '; ')
    $vmLookups = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-VM' }, $true))
    Check 'the VM is looked up in exactly one place, Get-BuildVm' 1 $vmLookups.Count

    Write-Host ''
    Write-Host '== this file, the guest script and testbed.json agree =='
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $testbedPath = Join-Path $repoRoot 'Testbed\testbed.json'
    $guestPath = Join-Path $repoRoot (Join-Path 'Testbed\guest' $GuestScriptName)
    $buildVm = $null
    if (Test-Path -LiteralPath $testbedPath) { $buildVm = ([System.IO.File]::ReadAllText($testbedPath) | ConvertFrom-Json).buildVm }
    Check 'testbed.json has a buildVm block' $true ($null -ne $buildVm)
    if ($null -ne $buildVm) {
        Check 'testbed.json names the same VM' $BuildVmName $buildVm.vmName
        Check 'and the same computer name' $BuildVmComputerName $buildVm.computerName
        Check 'and the same base checkpoint' $BaseCheckpointName $buildVm.runnerBaseCheckpoint
        Check 'and the base checkpoint is one of its checkpoints' $true (@($buildVm.checkpoints) -contains $BaseCheckpointName)
        Check 'and the same run directory' $GuestRunRoot $buildVm.guest.runRoot
        Check 'and the same offline feed' $GuestFeedRoot $buildVm.guest.offlineFeed
    }
    $guestText = ''
    if (Test-Path -LiteralPath $guestPath) { $guestText = [System.IO.File]::ReadAllText($guestPath) }
    Check 'the guest script exists beside this one' $true ($guestText.Length -gt 0)
    Check 'its run directory is this one' $true ($guestText.Contains("[string]   `$RunRoot = '$GuestRunRoot'"))
    Check 'its offline feed is this one' $true ($guestText.Contains("[string]   `$OfflineFeed = '$GuestFeedRoot'"))
    Check 'its guard names this computer' $true ($guestText.Contains("[string]   `$ExpectedComputerName = '$BuildVmComputerName'"))
    Check 'and this account' $true ($guestText.Contains("[string[]] `$ExpectedUser = @('$BuildVmUser')"))

    Write-Host ''
    Write-Host "$($script:stChecks) check(s), $($script:stFailures.Count) failure(s)."
    Write-Host 'NOT COVERED HERE, because it needs Hyper-V and the VM: the lock, the lease, the restore, the resume,'
    Write-Host 'PowerShell Direct, the copy, the detached guest run, the package staging and the fetch. A real run covers them.'
    if ($script:stFailures.Count -gt 0) { return 1 }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# SHARED BY A RUN AND ITS JANITOR: the lock, the history, and the VM.
# =============================================================================================
. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
$lockDir = Join-Path $env:SystemDrive 'OutlookAI-Testbed\build-vm'
New-Item -ItemType Directory -Force -Path $lockDir | Out-Null
$lockPath = Join-Path $lockDir 'runner.lock'
$ownerPath = Join-Path $lockDir 'runner.owner.json'
$historyPath = Join-Path $lockDir 'runs.log'
$vm = $null
$checkpoint = $null

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

function Enter-RunnerLock([int] $TimeoutMinutes) {
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $lastSaid = [DateTime]::MinValue
    $queued = $false
    while ($true) {
        try {
            return [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            if ((Get-Date) -ge $deadline) { return $null }
            if (((Get-Date) - $lastSaid).TotalSeconds -ge 60) {
                $who = 'another run'
                try {
                    $o = [System.IO.File]::ReadAllText($ownerPath) | ConvertFrom-Json
                    $who = "run $($o.runId) (pid $($o.pid), $($o.ref) in $($o.repo), started $($o.startedLocal))"
                }
                catch { }
                if (-not $queued) { Say "queued: the build VM is busy with $who" } else { Say "  still waiting for $who" }
                $queued = $true
                $lastSaid = Get-Date
            }
            Start-Sleep -Seconds 5
        }
    }
}

function Write-History([string] $Line) {
    try { Add-Content -LiteralPath $historyPath -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $Line) -Encoding UTF8 } catch { }
}

function Get-BuildVm {
    $found = @(Get-VM -Name $BuildVmName -ErrorAction SilentlyContinue)
    if ($found.Count -ne 1) { throw "expected one VM named '$BuildVmName' on this host and found $($found.Count). Testbed/README.md section 1c builds it." }
    if (-not (Test-IsBuildVmName $found[0].Name)) { throw "Hyper-V returned '$($found[0].Name)' for '$BuildVmName'. Refusing to act on it." }
    return $found[0]
}

function Get-BaseCheckpoint {
    $found = @(Get-VMSnapshot -VM $vm -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $BaseCheckpointName })
    if ($found.Count -ne 1) { throw "'$BuildVmName' has $($found.Count) checkpoint(s) named '$BaseCheckpointName', not one. Testbed/README.md section 1c takes it." }
    return $found[0]
}

# The VM's state, read afresh each time: a VirtualMachine object keeps the state it was read with
# and has no Refresh(), so a loop over one object would wait for ever on a stale value.
function Wait-VmSettled([int] $Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        $state = [string](Get-BuildVm).State
        if (@('Running', 'Off', 'Saved', 'Paused') -contains $state) { return $state }
        Start-Sleep -Seconds 2
    }
    return [string](Get-BuildVm).State
}

function Restore-Base {
    $state = Wait-VmSettled 120
    if (@('Running', 'Off', 'Saved', 'Paused') -notcontains $state) { throw "the VM stayed '$state' for two minutes; not restoring a checkpoint over a VM in transition." }
    Restore-VMSnapshot -VMSnapshot $checkpoint -Confirm:$false
    return (Wait-VmSettled 60)
}

# Restores the base checkpoint and leaves the VM saved: restored over a running VM, the base
# resumes at once, so it is saved - the BASE, not the run - which is what gives the host its RAM back.
function Restore-BaseAndSave {
    $final = Restore-Base
    if ($final -eq 'Running') {
        Save-VM -VM $vm
        $final = Wait-VmSettled 120
    }
    return $final
}

# =============================================================================================
# THE JANITOR - A CALLER THAT DIES in the banner. Started hidden by a run, through WMI, so that it
# is no child of the caller; it never touches the VM without the lock.
# =============================================================================================
if ($PSCmdlet.ParameterSetName -eq 'Janitor') {
    $janitorLog = Join-Path $lockDir 'janitor.log'
    function Write-JanitorLine([string] $Text) {
        try { Add-Content -LiteralPath $janitorLog -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}  {2}" -f (Get-Date), $JanitorRunId, $Text) -Encoding UTF8 } catch { }
    }
    if (-not (Test-IsRunId $JanitorRunId)) { exit $ExitCodes['REFUSED'] }

    # The run's process, pinned by its start time as well, so a process id reused after it ends
    # is not mistaken for it.
    $watched = Get-Process -Id $JanitorForPid -ErrorAction SilentlyContinue
    $watchedStart = $null
    if ($null -ne $watched) { $watchedStart = $watched.StartTime }
    $giveUp = (Get-Date).AddHours(6)
    while ($null -ne $watched) {
        if ((Get-Date) -ge $giveUp) { Write-JanitorLine 'the run is still going after six hours; leaving it'; exit 0 }
        Start-Sleep -Seconds 10
        $watched = Get-Process -Id $JanitorForPid -ErrorAction SilentlyContinue
        if ($null -ne $watched -and $watched.StartTime -ne $watchedStart) { $watched = $null }
    }

    $history = ''
    try { $history = [System.IO.File]::ReadAllText($historyPath) } catch { }
    if (Test-RunEnded -History $history -RunId $JanitorRunId) { exit 0 }

    Write-JanitorLine 'the run ended without finishing - its caller was stopped - so the janitor puts the VM back'
    $janitorLock = Enter-RunnerLock -TimeoutMinutes 10
    if ($null -eq $janitorLock) { Write-JanitorLine 'another run held the VM for ten minutes; it restores the checkpoint itself'; exit 0 }
    try {
        [ordered]@{ runId = "janitor for $JanitorRunId"; pid = $PID; ref = '-'; sha = '-'; repo = '-'; startedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') } |
            ConvertTo-Json | Set-Content -LiteralPath $ownerPath -Encoding UTF8
        $vm = Get-BuildVm
        $checkpoint = Get-BaseCheckpoint
        $state = Wait-VmSettled 120
        if ($state -eq 'Saved' -or $state -eq 'Off') { Write-JanitorLine "the VM is already $state - nothing to put back" }
        else {
            $final = Restore-BaseAndSave
            Write-JanitorLine "restored '$BaseCheckpointName' over the dead run's VM ($state); it is now $final"
        }
        $lease = Get-TestbedLease -VMName $BuildVmName
        if ($null -ne $lease -and ([string]$lease.reason) -eq "$LeaseReasonPrefix $JanitorRunId") {
            & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $BuildVmName -Release | Out-Null
            Write-JanitorLine 'released the dead run''s lease'
        }
        Write-History "janitor $JanitorRunId  its caller stopped part-way; the janitor put the VM back"
    }
    catch { Write-JanitorLine "FAILED: $($_.Exception.Message)" }
    finally {
        try { Remove-Item -LiteralPath $ownerPath -Force -ErrorAction SilentlyContinue } catch { }
        $janitorLock.Dispose()
    }
    exit 0
}

# =============================================================================================
# A RUN.
# =============================================================================================
$tStart = Get-Date
$timings = [ordered]@{}
function Stop-Refused([string] $Why) {
    Say "REFUSED: $Why"
    exit $ExitCodes['REFUSED']
}
function Measure-Phase([string] $Name, [DateTime] $Since) { $timings[$Name] = [math]::Round(((Get-Date) - $Since).TotalSeconds, 1) }

if ($QueueTimeoutMinutes -lt 1 -or $RunTimeoutMinutes -lt 5) { Stop-Refused '-QueueTimeoutMinutes must be at least 1 and -RunTimeoutMinutes at least 5.' }
if ($SkipSuite -and $SkipSelfTests) { Stop-Refused '-SkipSuite and -SkipSelfTests together leave nothing to run.' }

# ---- the repository and the revision ---------------------------------------------------------
if (-not $RepoPath) { $RepoPath = $PSScriptRoot }
$top = (@(Invoke-NativeCommand { & git -C $RepoPath rev-parse --show-toplevel 2>&1 }) | Out-String).Trim()
if ($LASTEXITCODE -ne 0) { Stop-Refused "'$RepoPath' is not inside a git repository: $top" }
$repo = [System.IO.Path]::GetFullPath($top)

$sha = (@(Invoke-NativeCommand { & git -C $repo rev-parse --verify --end-of-options ($Ref + '^{commit}') 2>&1 }) | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { Stop-Refused "'$Ref' does not name a commit in $repo - $sha" }
$subject = (@(Invoke-NativeCommand { & git -C $repo log -1 --format=%s $sha 2>&1 }) | Out-String).Trim()
$headSha = (@(Invoke-NativeCommand { & git -C $repo rev-parse --verify HEAD 2>&1 }) | Out-String).Trim()
$dirtyCount = 0
if ($headSha -eq $sha) {
    $dirtyCount = @(Invoke-NativeCommand { & git -C $repo status --porcelain 2>&1 } | Where-Object { $_ }).Count
}

if (-not $CredentialRepoRoot) {
    $common = (@(Invoke-NativeCommand { & git -C $repo rev-parse --path-format=absolute --git-common-dir 2>&1 }) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { Stop-Refused "git could not name the common directory of $repo - $common" }
    $CredentialRepoRoot = Split-Path -Parent ([System.IO.Path]::GetFullPath($common))
}
$credentialFile = Join-Path $CredentialRepoRoot 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\vm-credentials.json'
if (-not (Test-Path -LiteralPath $credentialFile)) { Stop-Refused "no guest credential at $credentialFile. It lives in the main checkout only (gitignored); pass -CredentialRepoRoot if that is somewhere else." }

if (-not $ResultsRoot) { $ResultsRoot = Join-Path $repo '.work\build-vm-runs' }
$runId = New-RunId -When $tStart -Sha $sha
$runDir = Join-Path $ResultsRoot $runId
New-Item -ItemType Directory -Force -Path $runDir | Out-Null
$archive = Join-Path $runDir 'Source.zip'

Say "== Invoke-TestsOnBuildVm: run $runId =="
Say "  revision  $($sha.Substring(0, 12)) - $subject"
Say "  from      $Ref in $repo"
if ($dirtyCount -gt 0) { Say "  NOTE      $dirtyCount uncommitted change(s) in $repo are NOT in this run - it tests the commit, as archived. Commit first to test them." }
$requestFilter = $Filter.Trim()
$filterShown = 'Category!=Live'
if ($requestFilter) { $filterShown = "Category!=Live&($requestFilter)" }
if ($SkipSuite) { Say '  suite     skipped (-SkipSuite)' } else { Say "  suite     dotnet test --filter $filterShown" }
$SelfTestInclude = @(Split-IncludeList $SelfTestInclude)
if ($SkipSelfTests) { Say '  selftests skipped (-SkipSelfTests)' }
elseif (@($SelfTestInclude).Count -gt 0) { Say "  selftests only $(@($SelfTestInclude) -join ', ')" }
else { Say '  selftests every script under Testbed\ and Tools\ that has -SelfTest' }
Say "  results   $runDir"

$t0 = Get-Date
Invoke-NativeCommand { & git -C $repo archive --format=zip -o $archive $sha 2>&1 } | Out-Null
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $archive)) { Stop-Refused "git archive of $sha failed (exit $LASTEXITCODE)." }
Measure-Phase 'archive' $t0

$request = [ordered]@{
    filter          = $requestFilter
    configuration   = 'Release'
    skipSuite       = [bool]$SkipSuite
    skipSelfTests   = [bool]$SkipSelfTests
    selfTestInclude = @($SelfTestInclude)
    revision        = [ordered]@{ ref = $Ref; sha = $sha; subject = $subject; repo = $repo; uncommittedLeftOut = $dirtyCount; runId = $runId }
}
$requestPath = Join-Path $runDir 'request.json'
($request | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $requestPath -Encoding UTF8

# ---- the lock ---------------------------------------------------------------------------------
$t0 = Get-Date
$lock = Enter-RunnerLock -TimeoutMinutes $QueueTimeoutMinutes
Measure-Phase 'queue' $t0
if ($null -eq $lock) {
    Say "INFRA: the build VM stayed busy for $QueueTimeoutMinutes minute(s); this run did not start. Nothing on the VM was touched."
    exit $ExitCodes['INFRA']
}
[ordered]@{ runId = $runId; pid = $PID; ref = $Ref; sha = $sha; repo = $repo; startedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') } |
    ConvertTo-Json | Set-Content -LiteralPath $ownerPath -Encoding UTF8
if ($timings['queue'] -ge 5) { Say "the build VM is ours after $(Format-Seconds $timings['queue']) in the queue" }
Write-History "start  $runId  $Ref  $repo"

# The janitor, for the case this process is stopped part-way (the banner). Through WMI, so it is
# no child of this process and survives whatever stops it; hidden, so no console appears. A run
# that cannot start one still runs, and says so.
if (-not $KeepVmRunning) {
    try {
        $shellExe = (Get-Process -Id $PID).Path
        $janitorCommand = '"{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}" -JanitorForPid {2} -JanitorRunId {3}' -f $shellExe, $PSCommandPath, $PID, $runId
        $hidden = New-CimInstance -CimClass (Get-CimClass -ClassName Win32_ProcessStartup) -Property @{ ShowWindow = [uint16]0 } -ClientOnly
        $created = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $janitorCommand; ProcessStartupInformation = $hidden }
        if ($created.ReturnValue -ne 0) { Say "janitor: Win32_Process.Create returned $($created.ReturnValue) - this run has none, and if it is stopped part-way the VM stays running until the next run" }
        else { Write-Verbose "janitor pid $($created.ProcessId)" }
    }
    catch { Say "janitor: not started ($($_.Exception.Message)) - if this run is stopped part-way the VM stays running until the next run" }
}

# ---- everything that touches the VM, with the lock held --------------------------------------
$infraError = ''
$guestVerdict = ''
$session = $null
$leaseTaken = $false
$packagesStaged = $false
$script:guestPid = $null
$script:guestTimedOut = $false
$guestStdout = Join-Path $GuestRunRoot 'guest.out.txt'
$guestStderr = Join-Path $GuestRunRoot 'guest.err.txt'

function Invoke-InGuest([scriptblock] $Block, [object[]] $ArgumentList = @()) {
    Invoke-Command -Session $session -ScriptBlock $Block -ArgumentList $ArgumentList -ErrorAction Stop
}

# Starts the guest script detached and follows its log until done.txt appears, the process ends,
# or the time limit passes. Returns the guest's exit code as written in done.txt, or $null.
function Invoke-GuestRun([int] $Attempt) {
    Say "guest: starting $GuestScriptName (attempt $Attempt)"
    $guestPid = Invoke-InGuest -ArgumentList @((Join-Path $GuestRunRoot $GuestScriptName), (Join-Path $GuestRunRoot 'request.json'), $guestStdout, $guestStderr, $GuestRunRoot) -Block {
        param($scriptPath, $requestPath, $outPath, $errPath, $root)
        # Everything an earlier attempt left, gone BEFORE the new one starts: the log is followed
        # from its first byte, and on the first package retry (2026-10-03) the host read the old
        # attempt's log back in the seconds before the new process replaced it.
        foreach ($f in @($outPath, $errPath, (Join-Path $root 'done.txt'), (Join-Path $root 'results.zip'), (Join-Path $root 'results'))) {
            if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Recurse -Force }
        }
        $p = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, '-RequestPath', $requestPath) -RedirectStandardOutput $outPath -RedirectStandardError $errPath -NoNewWindow -PassThru
        $p.Id
    }
    $script:guestPid = $guestPid
    $logPath = Join-Path $GuestRunRoot 'results\run.log'
    $donePath = Join-Path $GuestRunRoot 'done.txt'
    $offset = 0
    $deadline = (Get-Date).AddMinutes($RunTimeoutMinutes)
    # A poll that fails - PowerShell Direct answering late, the guest briefly out of reach - is
    # retried with a growing pause, and only twelve in a row (about three minutes) end the run.
    # A BUSY log is not a failure at all: $PollGuestBlock answers it, and the next poll reads on -
    # said once when it starts and once when it ends, so a stall in the log has a visible reason.
    # Only for a script that has died does a busy log stop counting, after ten polls (about half a
    # minute): there is nothing more to wait for, and the results are fetched as they are.
    $failedPolls = 0
    $busyPolls = 0
    while ($true) {
        if ((Get-Date) -ge $deadline) { $script:guestTimedOut = $true; throw "the guest run did not finish within $RunTimeoutMinutes minute(s)" }
        $poll = $null
        try { $poll = Invoke-InGuest -ArgumentList @($logPath, $offset, $guestPid, $donePath) -Block $PollGuestBlock }
        catch {
            $failedPolls++
            if ($failedPolls -ge 12) { throw "the guest did not answer $failedPolls polls in a row: $($_.Exception.Message)" }
            Start-Sleep -Seconds ([Math]::Min(20, 2 * $failedPolls))
            continue
        }
        $failedPolls = 0
        foreach ($line in ($poll.Text -split "`r?`n")) { if ($line.Trim()) { Write-Host "  vm| $line" } }
        $offset = $poll.Next
        if ($poll.Busy) {
            if ($busyPolls -eq 0) { Say 'guest: its log is held by another process just now - not an error; the next poll reads on' }
            $busyPolls++
        }
        elseif ($busyPolls -gt 0) {
            Say "guest: its log reads again, after $busyPolls busy poll(s)"
            $busyPolls = 0
        }
        if ($poll.Done) { return $poll.Code }
        if (-not $poll.Alive -and (-not $poll.Busy -or $busyPolls -ge 10)) {
            # Ended without done.txt - or wrote it a moment ago: one more look, then $null.
            Start-Sleep -Seconds 2
            $late = $null
            try { $late = Invoke-InGuest -ArgumentList @($logPath, $offset, $guestPid, $donePath) -Block $PollGuestBlock } catch { }
            if ($null -ne $late) {
                foreach ($line in ($late.Text -split "`r?`n")) { if ($line.Trim()) { Write-Host "  vm| $line" } }
                if ($late.Done) { return $late.Code }
            }
            return $null
        }
        Start-Sleep -Seconds 3
    }
}

# After a run ends for a reason on THIS side, with the guest script perhaps still going: wait for
# its done.txt while its process lives, a few minutes at most, so its results can still be fetched
# and decide the verdict (Get-HostVerdict). Not after the run's own time limit, and not over a
# PowerShell Direct session that has broken - nothing would come back through it.
function Wait-GuestDone([int] $Minutes) {
    if ($null -eq $session -or $null -eq $script:guestPid -or $script:guestTimedOut) { return }
    $donePath = Join-Path $GuestRunRoot 'done.txt'
    $until = (Get-Date).AddMinutes($Minutes)
    $said = $false
    while ((Get-Date) -lt $until) {
        $state = $null
        try { $state = Invoke-InGuest -ArgumentList @($script:guestPid, $donePath) -Block $GuestStateBlock }
        catch { if ([string]$session.State -ne 'Opened') { return } }
        if ($null -ne $state) {
            if ($state.Done -or -not $state.Alive) { return }
            if (-not $said) { Say "guest: still running - waiting up to $Minutes minute(s) for it to finish, so its results can decide"; $said = $true }
        }
        Start-Sleep -Seconds 5
    }
}

try {
    # The lease: an operator's own lease on the VM holds runs back; ours keeps the idle-saver off.
    $t0 = Get-Date
    $holdDeadline = (Get-Date).AddMinutes($QueueTimeoutMinutes)
    $saidHold = $false
    while ($true) {
        $lease = Get-TestbedLease -VMName $BuildVmName
        if ($null -eq $lease -or ([string]$lease.reason).StartsWith($LeaseReasonPrefix)) { break }
        if (-not $saidHold) { Say "queued: $BuildVmName is leased by hand until $($lease.expiresUtc) - '$($lease.reason)'. Waiting for it to be released."; $saidHold = $true }
        if ((Get-Date) -ge $holdDeadline) {
            $timings['queue'] = $timings['queue'] + [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
            throw "$BuildVmName stayed leased by hand ('$($lease.reason)') for $QueueTimeoutMinutes minute(s)"
        }
        Start-Sleep -Seconds 10
    }
    $timings['queue'] = $timings['queue'] + [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
    $leaseOutput = & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $BuildVmName -Minutes ([Math]::Min(480, $RunTimeoutMinutes + 20)) -Reason "$LeaseReasonPrefix $runId"
    $leaseTaken = $true
    Write-Verbose ([string]$leaseOutput)

    # The VM, from its base checkpoint.
    $t0 = Get-Date
    $vm = Get-BuildVm
    $checkpoint = Get-BaseCheckpoint
    $stateBefore = [string]$vm.State
    $stateAfter = Restore-Base
    Say "VM: restored '$BaseCheckpointName' ($stateBefore -> $stateAfter)"
    # A checkpoint applied to a RUNNING VM resumes it at once (measured on the first run,
    # 2026-10-03: Running -> Running); applied to a saved or stopped one it leaves it saved.
    if ($stateAfter -ne 'Running') { Start-VM -VM $vm }
    $hbDeadline = (Get-Date).AddMinutes(3)
    while ($true) {
        $hb = Get-VMIntegrationService -VM $vm -Name 'Heartbeat'
        if ([string]$hb.PrimaryStatusDescription -eq 'OK') { break }
        if ((Get-Date) -ge $hbDeadline) { throw "the VM's heartbeat did not read OK within 3 minutes of Start-VM (it reads '$($hb.PrimaryStatusDescription)')" }
        Start-Sleep -Seconds 1
    }
    Measure-Phase 'resume' $t0

    $t0 = Get-Date
    $credential = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $CredentialRepoRoot -VMName $BuildVmName
    $psDeadline = (Get-Date).AddMinutes(3)
    $lastError = ''
    while ($null -eq $session) {
        try { $session = New-PSSession -VMName $BuildVmName -Credential $credential -ErrorAction Stop }
        catch {
            $lastError = $_.Exception.Message
            if ((Get-Date) -ge $psDeadline) { throw "PowerShell Direct did not answer within 3 minutes: $lastError" }
            Start-Sleep -Seconds 2
        }
    }
    $credential = $null
    $who = Invoke-InGuest -Block { [pscustomobject]@{ Computer = $env:COMPUTERNAME; User = $env:USERNAME } }
    if (-not [string]::Equals([string]$who.Computer, $BuildVmComputerName, [System.StringComparison]::OrdinalIgnoreCase) -or [string]$who.User -ne $BuildVmUser) {
        throw "the VM answers as '$($who.User)' on '$($who.Computer)', not '$BuildVmUser' on '$BuildVmComputerName'. Refusing to copy anything into it."
    }
    # A resumed guest starts at its checkpoint's time and Hyper-V's time sync moves it to now within
    # seconds; a build or a test that started before that would see the clock jump under it.
    $skew = $null
    $clockDeadline = (Get-Date).AddMinutes(2)
    while ($true) {
        $before = [DateTime]::UtcNow
        $guestNow = Invoke-InGuest -Block { [DateTime]::UtcNow }
        $after = [DateTime]::UtcNow
        $hostMid = $before.AddTicks([long](($after - $before).Ticks / 2))
        $skew = [math]::Round(([DateTime]$guestNow - $hostMid).TotalSeconds, 1)
        if ([math]::Abs($skew) -le $MaxClockSkewSeconds) { break }
        if ((Get-Date) -ge $clockDeadline) { throw "the VM's clock is still $skew s from this host's two minutes after it resumed - Hyper-V time sync did not catch it up" }
        Start-Sleep -Seconds 2
    }
    Measure-Phase 'connect' $t0
    Say "VM: $($who.Computer) as $($who.User), clock within $skew s of this host, $(Format-Seconds ($timings['resume'] + $timings['connect'])) after the restore"

    # Stage the run.
    $t0 = Get-Date
    Invoke-InGuest -ArgumentList @($GuestRunRoot) -Block {
        param($root)
        if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
        New-Item -ItemType Directory -Force -Path $root | Out-Null
    }
    Copy-Item -LiteralPath $archive -Destination (Join-Path $GuestRunRoot 'Source.zip') -ToSession $session
    Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'guest' $GuestScriptName)) -Destination (Join-Path $GuestRunRoot $GuestScriptName) -ToSession $session
    Copy-Item -LiteralPath $requestPath -Destination (Join-Path $GuestRunRoot 'request.json') -ToSession $session
    Measure-Phase 'stage' $t0

    # Run - and once more if the feed lacked a package this revision needs.
    $t0 = Get-Date
    $code = Invoke-GuestRun -Attempt 1
    if ($code -eq '20') {
        Measure-Phase 'guest-attempt-1' $t0
        $tp = Get-Date
        Say "packages: the VM's offline feed lacks package(s) this revision needs - staging its restore closure on this host"
        $feedDir = Join-Path $runDir 'feed'
        New-Item -ItemType Directory -Force -Path $feedDir | Out-Null
        $publishLog = Join-Path $feedDir 'publish.log'
        # Its whole transcript goes to a file - it ends with copy-in instructions meant for a person
        # staging a guest by hand - and only the lines about the packages are shown here.
        # ps51-native-stderr-ok: a PowerShell script, not a program - every program it starts goes through its own Invoke-NativeCommand, under 'Continue' inside a try
        & (Join-Path $PSScriptRoot 'Publish-LiveTierPayload.ps1') -RepoRoot $repo -Ref $sha -OutDir $feedDir *> $publishLog
        foreach ($line in @(Get-Content -LiteralPath $publishLog | Where-Object { $_ -match 'package\(s\)|NuGet\.zip  |The feed restores|FAIL' })) { Write-Host "  host| $($line.Trim())" }
        $nugetZip = Join-Path $feedDir 'NuGet.zip'
        if (-not (Test-Path -LiteralPath $nugetZip)) { throw "Publish-LiveTierPayload.ps1 left no $nugetZip - its transcript is $publishLog" }
        Copy-Item -LiteralPath $nugetZip -Destination (Join-Path $GuestRunRoot 'NuGet.zip') -ToSession $session
        $added = Invoke-InGuest -ArgumentList @((Join-Path $GuestRunRoot 'NuGet.zip'), $GuestFeedRoot) -Block {
            param($zip, $feed)
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $n = 0
            $z = [System.IO.Compression.ZipFile]::OpenRead($zip)
            try {
                foreach ($e in $z.Entries) {
                    if (-not $e.Name) { continue }
                    $target = Join-Path $feed $e.Name
                    if (-not (Test-Path -LiteralPath $target)) { [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, $target); $n++ }
                }
            }
            finally { $z.Dispose() }
            $n
        }
        foreach ($bulky in @('source', 'package-cache', 'nuget-offline')) {
            $p = Join-Path $feedDir $bulky
            if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
        }
        $packagesStaged = $true
        Measure-Phase 'packages' $tp
        Say "packages: $added package(s) added to the VM's feed for this run only - the base checkpoint keeps its own feed"
        $t0 = Get-Date
        $code = Invoke-GuestRun -Attempt 2
    }
    Measure-Phase 'guest' $t0
    if ($null -eq $code) { Say 'guest: the script ended without writing done.txt' }
}
catch {
    $infraError = $_.Exception.Message
    Say "INFRA: $infraError"
}
finally {
    # Fetch whatever the guest left, even after a failure: a partial log is how a failure is read.
    # And when the failure was on THIS side, the guest may well finish anyway - it did on
    # 2026-10-03 - so it is given a few minutes to, and its results then decide the verdict.
    $t0 = Get-Date
    $vmDir = Join-Path $runDir 'vm'
    if ($null -ne $session) {
        if ($infraError) { Wait-GuestDone -Minutes 5 }
        try {
            New-Item -ItemType Directory -Force -Path $vmDir | Out-Null
            $present = Invoke-InGuest -ArgumentList @($GuestRunRoot) -Block {
                param($root)
                @('results.zip', 'guest.out.txt', 'guest.err.txt', 'done.txt') | Where-Object { Test-Path -LiteralPath (Join-Path $root $_) }
            }
            foreach ($name in @($present)) {
                # Retried: a file the guest still holds for a moment is not a missing result.
                for ($attempt = 1; $true; $attempt++) {
                    try {
                        Copy-Item -FromSession $session -LiteralPath (Join-Path $GuestRunRoot $name) -Destination (Join-Path $vmDir $name)
                        break
                    }
                    catch {
                        if ($attempt -ge 5) { throw }
                        Start-Sleep -Seconds (2 * $attempt)
                    }
                }
            }
            $zip = Join-Path $vmDir 'results.zip'
            if (Test-Path -LiteralPath $zip) {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $vmDir)
                Remove-Item -LiteralPath $zip -Force
            }
            else {
                # No zip: the guest did not reach its end, or could not zip what it had. Bring back
                # what it left one file at a time, each read with sharing - the guest may still be
                # writing - so that a run.json and a TRX file that exist still decide the verdict.
                $fetched = @()
                foreach ($rel in @('run.log', 'run.json', 'trx\suite.trx', 'restore.out.txt', 'build.out.txt')) {
                    $text = Invoke-InGuest -ArgumentList @((Join-Path $GuestRunRoot (Join-Path 'results' $rel))) -Block $ReadGuestLogBlock
                    if ($null -eq $text) { continue }
                    $target = Join-Path $vmDir $rel
                    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
                    [System.IO.File]::WriteAllText($target, [string]$text, (New-Object System.Text.UTF8Encoding($false)))
                    $fetched += $rel
                }
                $fetchedText = 'nothing was there'
                if ($fetched.Count -gt 0) { $fetchedText = $fetched -join ', ' }
                Say "fetch: no results.zip came back - read one by one instead: $fetchedText"
            }
        }
        catch { Say "fetch: $($_.Exception.Message)"; if (-not $infraError) { $infraError = "fetching the results failed: $($_.Exception.Message)" } }
        try { Remove-PSSession -Session $session } catch { }
    }
    Measure-Phase 'fetch' $t0

    # Put the VM back to its base checkpoint, which leaves it saved and holding no RAM.
    $t0 = Get-Date
    if ($null -ne $vm -and $null -ne $checkpoint) {
        if ($KeepVmRunning) { Say "VM: left running as the run left it (-KeepVmRunning). The next run restores '$BaseCheckpointName' anyway." }
        else {
            try {
                $final = Restore-BaseAndSave
                if ($final -eq 'Saved' -or $final -eq 'Off') { Say "VM: restored '$BaseCheckpointName' again and $($final.ToLowerInvariant()) - no RAM held until the next run" }
                else { Say "VM: restored '$BaseCheckpointName' again, but it is $final rather than saved" }
            }
            catch { Say "VM: could not restore '$BaseCheckpointName' after the run: $($_.Exception.Message). The next run restores it before it starts." }
        }
    }
    Measure-Phase 'revert' $t0

    if ($leaseTaken) {
        try {
            $current = Get-TestbedLease -VMName $BuildVmName
            if ($null -ne $current -and ([string]$current.reason) -eq "$LeaseReasonPrefix $runId") { & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $BuildVmName -Release | Out-Null }
        }
        catch { Say "lease: could not release it ($($_.Exception.Message)); it expires by itself" }
    }
    try { Remove-Item -LiteralPath $ownerPath -Force -ErrorAction SilentlyContinue } catch { }
    $lock.Dispose()
}

# ---- the verdict, from what came back ----------------------------------------------------------
$vmDir = Join-Path $runDir 'vm'
$runJsonPath = Join-Path $vmDir 'run.json'
$guestRecord = $null
if (Test-Path -LiteralPath $runJsonPath) {
    try { $guestRecord = [System.IO.File]::ReadAllText($runJsonPath) | ConvertFrom-Json; $guestVerdict = [string]$guestRecord.verdict }
    catch { if (-not $infraError) { $infraError = "run.json came back unreadable: $($_.Exception.Message)" } }
}
$trxPath = Join-Path $vmDir 'trx\suite.trx'
$trx = $null
if (Test-Path -LiteralPath $trxPath) {
    try { $trx = ConvertFrom-TrxText ([System.IO.File]::ReadAllText($trxPath)) }
    catch { if (-not $infraError) { $infraError = "the TRX file came back unreadable: $($_.Exception.Message)" } }
}
$selfTests = @()
$selfTestSkips = @()
if ($null -ne $guestRecord) { $selfTests = @($guestRecord.selfTests); $selfTestSkips = @($guestRecord.selfTestSkips) }
$selfTestsFailed = @($selfTests | Where-Object { -not $_.passed }).Count
$hostVerdict = Get-HostVerdict -GuestVerdict $guestVerdict -InfraError $infraError -Trx $trx -SuiteRequested (-not $SkipSuite) -SelfTestsRun $selfTests.Count -SelfTestsFailed $selfTestsFailed
$timings['total'] = [math]::Round(((Get-Date) - $tStart).TotalSeconds, 1)
$exitCode = $ExitCodes[$hostVerdict.Verdict]

# The guest's own phase timings, beside the host's.
$guestPhases = [ordered]@{}
if ($null -ne $guestRecord -and $null -ne $guestRecord.phases) {
    foreach ($p in $guestRecord.phases.PSObject.Properties) { $guestPhases[$p.Name] = $p.Value.seconds }
}

$lines = New-Object System.Collections.Generic.List[string]
$headline = "== $runId - $($hostVerdict.Verdict) (exit $exitCode)"
if ($hostVerdict.Why) { $headline += " - $($hostVerdict.Why)" }
$lines.Add($headline + ' ==')
$lines.Add("revision   $($sha.Substring(0, 12)) - $subject")
$lines.Add("           $Ref in $repo$(if ($dirtyCount -gt 0) { " ($dirtyCount uncommitted change(s) NOT included)" })")
if ($infraError -and $guestVerdict -and $hostVerdict.Verdict -ne 'INFRA') {
    $lines.Add("host note  something failed on this side - $infraError - and the guest's results came back anyway; they decided the verdict")
}
if ($SkipSuite) { $lines.Add('suite      skipped') }
elseif ($null -ne $trx) {
    $c = $trx.Counts
    $lines.Add(("suite      {0} total: {1} passed, {2} failed, {3} skipped{4} - filter {5}" -f $c['total'], $c['passed'], $c['failed'], $c['skipped'], $(if ($c['other'] -gt 0) { ", $($c['other']) other" } else { '' }), $filterShown))
}
else { $lines.Add("suite      no TRX file - guest verdict '$guestVerdict'") }
if ($SkipSelfTests) { $lines.Add('self-tests skipped') }
elseif ($null -eq $guestRecord) { $lines.Add('self-tests not known - no run.json came back') }
else { $lines.Add(("self-tests {0} run: {1} passed, {2} failed; {3} skipped by reason" -f $selfTests.Count, ($selfTests.Count - $selfTestsFailed), $selfTestsFailed, $selfTestSkips.Count)) }
if ($null -ne $guestRecord -and $null -ne $guestRecord.machine) {
    $m = $guestRecord.machine
    $lines.Add(("machine    {0} ({1}), {2} / {3}, {4}, {5} CPU, {6} GB, session {7}, Outlook.Application registered: {8}" -f $m.computerName, $m.os, $m.culture, $m.uiCulture, $m.timeZone, $m.logicalProcessors, $m.memoryGB, $m.sessionId, $m.outlookInstalled))
}
$timeParts = @()
foreach ($k in @('queue', 'archive', 'resume', 'connect', 'stage', 'packages', 'guest', 'fetch', 'revert')) {
    if ($timings.Contains($k)) { $timeParts += "$k $(Format-Seconds $timings[$k])" }
}
$lines.Add("time       $(Format-Seconds $timings['total']) in all: $($timeParts -join ', ')")
if ($guestPhases.Count -gt 0) {
    $lines.Add("           in the VM: $((@($guestPhases.Keys | ForEach-Object { "$_ $(Format-Seconds $guestPhases[$_])" })) -join ', ')")
}
if ($packagesStaged) { $lines.Add('packages   staged for this revision on the host and added to the VM''s feed for this run (the base checkpoint''s feed predates them)') }
# A revision that does not build: its compiler or NuGet errors, once each, so the summary alone says why.
$buildErrors = @()
if ($guestVerdict -eq 'BUILD-FAILED' -or $guestVerdict -eq 'RESTORE-FAILED') {
    $phaseLog = Join-Path $vmDir 'build.out.txt'
    if ($guestVerdict -eq 'RESTORE-FAILED') { $phaseLog = Join-Path $vmDir 'restore.out.txt' }
    if (Test-Path -LiteralPath $phaseLog) {
        $buildErrors = @(Get-Content -LiteralPath $phaseLog | Where-Object { $_ -match ':\s+error\s+[A-Z]+\d+' } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
    }
    if ($buildErrors.Count -gt 0) {
        $lines.Add('')
        $lines.Add($(if ($guestVerdict -eq 'RESTORE-FAILED') { 'RESTORE ERRORS' } else { 'BUILD ERRORS' }))
        foreach ($e in @($buildErrors | Select-Object -First 30)) { $lines.Add("  $e") }
        if ($buildErrors.Count -gt 30) { $lines.Add("  ... and $($buildErrors.Count - 30) more - see $phaseLog") }
    }
}
if ($null -ne $trx -and @($trx.Failed).Count -gt 0) {
    $lines.Add('')
    $lines.Add('FAILED TESTS')
    foreach ($f in @($trx.Failed | Select-Object -First 50)) { $lines.Add("  $($f.name)"); if ($f.message) { $lines.Add("      $($f.message)") } }
    if (@($trx.Failed).Count -gt 50) { $lines.Add("  ... and $(@($trx.Failed).Count - 50) more - see vm\trx\suite.trx") }
}
if ($null -ne $trx -and @($trx.Skipped).Count -gt 0) {
    $lines.Add('')
    $lines.Add('SKIPPED TESTS')
    foreach ($s in @($trx.Skipped | Select-Object -First 20)) { $lines.Add("  $s") }
    if (@($trx.Skipped).Count -gt 20) { $lines.Add("  ... and $(@($trx.Skipped).Count - 20) more") }
}
if ($selfTestsFailed -gt 0) {
    $lines.Add('')
    $lines.Add('FAILED SELF-TESTS')
    foreach ($s in @($selfTests | Where-Object { -not $_.passed })) { $lines.Add("  $($s.path) - exit $($s.exitCode)$(if ($s.timedOut) { ', timed out' }) - $($s.lastLine)"); $lines.Add("      log: vm\$($s.log.Replace('/', '\'))") }
}
if ($selfTestSkips.Count -gt 0) {
    $lines.Add('')
    $lines.Add('SELF-TESTS NOT RUN HERE, BY REASON')
    foreach ($s in $selfTestSkips) { $lines.Add("  $($s.path) - $($s.reason)") }
}
if ($null -ne $guestRecord -and @($guestRecord.notes).Count -gt 0) {
    $lines.Add('')
    $lines.Add('NOTES FROM THE VM')
    foreach ($n in @($guestRecord.notes)) { $lines.Add("  $n") }
}
$lines.Add('')
$lines.Add("results    $runDir")
$lines.Add('           summary.txt and summary.json here; vm\run.log, vm\trx\suite.trx and every build, test and self-test log under vm\')

Set-Content -LiteralPath (Join-Path $runDir 'summary.txt') -Value $lines -Encoding UTF8
$summaryFilter = $filterShown
if ($SkipSuite) { $summaryFilter = '' }
$summary = New-RunSummary -RunId $runId -Verdict $hostVerdict -Revision ([ordered]@{ ref = $Ref; sha = $sha; subject = $subject; repo = $repo; uncommittedLeftOut = $dirtyCount }) -Filter $summaryFilter -Trx $trx -SelfTests $selfTests -SelfTestSkips $selfTestSkips -GuestVerdict $guestVerdict -HostError $infraError -BuildErrors $buildErrors -PackagesStaged $packagesStaged -Timings $timings -GuestPhases $guestPhases -Results $runDir
($summary | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $runDir 'summary.json') -Encoding UTF8
Write-History ("end    $runId  $($hostVerdict.Verdict)  $(Format-Seconds $timings['total'])")

Write-Host ''
foreach ($l in $lines) { Write-Host $l }
exit $exitCode
