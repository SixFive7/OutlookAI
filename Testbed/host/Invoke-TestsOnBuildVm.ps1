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
      2. Takes a build VM of the pool that is free - its lock, queueing while every one is busy
         (see CONCURRENT CALLERS) - or the one -VMName names.
      3. Takes a lease on that VM (Testbed/README.md section 5b), so the idle-saver keeps away.
      4. Restores the base checkpoint and resumes the VM from it, waits for its heartbeat and
         PowerShell Direct, proves it is OAI-BUILD logged on as the autologon account in the UTC
         time zone (Q126 (a): the suite runs in a zone other than the workstation's), and waits
         for its clock to agree with this host's (a resumed guest starts at its checkpoint's time).
      5. Copies in the archive, Testbed/guest/Invoke-BuildVmRun.ps1 and a request file, starts the
         guest script detached, and follows its log until it writes done.txt. That script restores
         from the VM's offline feed, builds, runs `dotnet test --filter Category!=Live` (ANDed in
         code with -Filter) into a TRX file, then THE ZONE PASS (Q144): it switches the VM to W. Europe
         Standard Time ($ZonePassTimeZone), runs the tests marked `TimeZone=Machine` - those that can
         read the machine's zone, which on a UTC machine is UTC itself - again into a second TRX
         file, and switches it back; then it runs every -SelfTest it finds under Testbed\ and Tools\.
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

    CONCURRENT CALLERS SHARE THE BUILD POOL (2026-11-17). One run per VM at a time, and as many VMs
    as the pool has: OutlookAI-Build and its clones (Testbed/host/TestbedPool.ps1), whole copies of
    it made by Testbed/host/New-TestbedClone.ps1 - the same base checkpoint by the same name, the
    same computer name and account - so any member gives a commit the same run. Each VM has a lock
    file opened exclusively under %SystemDrive%\OutlookAI-Testbed\build-vm\ (machine-wide, beside
    the leases); OutlookAI-Build's keeps the name every runner before the pool used, runner.lock,
    so a run from an older checkout and a run from this one never share it. A run takes the first
    member that is free (Enter-PoolVm) - a member other than the first only while the host has its
    memory and a reserve to spare - and waits in the queue while none is, printing what holds each
    every minute, for up to -QueueTimeoutMinutes. The lock is an open file handle, so a caller that
    dies - killed, crashed, its terminal closed - releases it with its process, and the next run
    on that VM restores the checkpoint whatever state it finds.

    A RUN THAT WAITS ASKS HYPER-V NOTHING (2026-11-18). Queued runs asked Get-VM for every member
    every 5 s, and on the night eighteen runners were alive the WMI service answering them used about
    5 of the host's 32 logical processors and the maintainer's PC lagged. Now a waiting run holds a
    ticket in the testbed's queue (TestbedPool.ps1) and tries only the members no earlier ticket
    wants, first come first served; it learns whether one is free from its lock and lease files
    alone, and asks Hyper-V once one is - for its startup memory, any member but the first. Between
    turns it sleeps until a lock's owner file, a lease or a ticket changes, or 30 s pass. Which
    members are on this host is asked once, and a VM's state - while a restore or a save moves it -
    with one narrow CIM query, never Get-VM. While a run executes it asks its guest how far it got
    every 15 s and its VM for a heartbeat every 15 s after the first two looks.

    A CALLER THAT DIES DOES NOT LEAVE THE VM HOLDING ITS RAM. Measured 2026-10-03: a run stopped
    part-way left the VM running, 6 GB held, for three hours, until the next run came - nothing on
    this host saves it, the idle-saver's task not being registered here, and an agent's tool call is
    exactly the kind of caller that gets stopped on a time limit. So every run, once it holds the
    lock, starts a JANITOR: this same file with -JanitorForPid, launched through WMI's
    Win32_Process.Create so that it is no child of the caller and outlives whatever stops it - a
    killed process tree, a closed terminal - and hidden. It waits for the run's process to end, on
    the process's own handle, with no poll. A
    run that finished wrote its end line to runs.log, and the janitor just exits; one that did not
    is put right: the janitor takes the lock (behind any run that took over first, which restores
    the checkpoint itself), restores the base checkpoint, saves the VM and releases the dead run's
    lease, and says so in janitor.log beside the lock. It never touches the VM without the lock.
    A LEASE someone else holds on a VM (Set-TestbedLease.ps1,
    for maintenance by hand) holds runs off that VM until it is released or expires; they take
    another member. Running several VMs in parallel was not chosen on 2026-10-03, when a queue was
    rarely long; with eight agents queueing it was (the maintainer, 2026-11-17: "try to parallelize
    as much as possible"), at 27.5 GB of disk a clone (measured) and, only while it runs, 6 GB of memory.

    THE PRIVATE LIST OF PERSONAL DATA (Q196 (c), 2026-11-18). Tools/Checks/check-personal-data.ps1
    knows the identifiers it guards only from .work\personal-data\identifiers.json - nothing about
    them is committed - so each run that runs self-tests copies that file into the VM: from this
    checkout's .work, else from the main checkout's. It goes straight from there to the VM, beside
    the run and outside its results, so it is never fetched back, and the revert after the run takes
    it with everything else. A run that finds no list says so in its log, its summary's headline and
    summary.json: the known identifiers were NOT checked.

    CREDENTIALS only through Testbed/host/Get-GuestCredential.ps1, from the main checkout - found
    with `git rev-parse --git-common-dir`, so this runs from any worktree - or -CredentialRepoRoot.
    Never printed, never written.

    ONLY THE BUILD POOL. Testbed/README.md section 4a forbids a default that picks one of several
    DIFFERENT machines; the members of a pool are one machine, copied, and this restores
    checkpoints, which on any other VM would destroy its state. So the pool is a list, not a
    parameter, -VMName may only pin one of its members, Find-PoolVm refuses any other name before
    Hyper-V is asked, every Hyper-V call goes through the VM object Get-BuildVm returns for the
    member the run holds, and the computer name inside must read OAI-BUILD before anything is
    copied in. -SelfTest pins all of it.

    EXIT CODES
        0  PASS      the revision built, every selected test and every self-test passed
        1  FAIL      a test or a self-test failed - in the suite or in the zone pass - the suite or
                     the zone pass hung or crashed, the zone pass could not switch the VM's zone or
                     put it back, or the filter selected no test at all. A zone pass that selects no
                     test is no failure: a revision from before Q144 marks none, and a narrow
                     -Filter may select none of the marked ones
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

.PARAMETER VMName
    Pin the run to one build VM of the pool (Testbed/host/TestbedPool.ps1), waiting for that one
    while it is busy - to prove a new clone before it joins the pool, or to look at one machine.
    Default: whichever member is free. Any name outside the build pool is refused.

.PARAMETER JanitorForPid
    Internal: what a run starts its janitor with (see A CALLER THAT DIES). The process id of the
    run to watch. Not for use by hand.

.PARAMETER JanitorRunId
    Internal: the id of the run the janitor watches.

.PARAMETER JanitorVmName
    Internal: the pool member the run the janitor watches took.

.PARAMETER SelfTest
    The run-id format, the TRX reader, the verdict table with the zone pass, summary.json's
    shape, the second zone, the VM-name
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
    [Parameter(ParameterSetName = 'Run')] [string]   $VMName,
    [Parameter(Mandatory = $true, ParameterSetName = 'Janitor')] [int]    $JanitorForPid,
    [Parameter(Mandatory = $true, ParameterSetName = 'Janitor')] [string] $JanitorRunId,
    [Parameter(Mandatory = $true, ParameterSetName = 'Janitor')] [string] $JanitorVmName,
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
# THE BUILD POOL (2026-11-17): TestbedPool.ps1's - OutlookAI-Build and its clones, each a whole copy
# of it at the base checkpoint, its computer name and account included - so every member answers as
# OAI-BUILD and restores the same base by name. A run takes whichever member is free (Enter-PoolVm);
# -VMName pins one. $BuildVmName is the member THIS run or janitor holds, set by Use-PoolVm once it
# is chosen, and every Hyper-V call below names it. -SelfTest holds the pool to testbed.json.
. (Join-Path $PSScriptRoot 'TestbedPool.ps1')
$BuildVmPool = @($TestbedPools['build'])
$BuildVmName = ''
$BuildVmComputerName = 'OAI-BUILD'
$BuildVmUser = 'vmadmin'
$BaseCheckpointName = 'CP-03-SDK-TEST-READY-UTC'
# The guest's time zone: UTC, decided by the maintainer 2026-10-03 (Q126 (a)), so the suite runs in a
# zone other than the workstation's W. Europe - GitHub's UTC runner, which caught Q95's zone bug, is
# gone. The base checkpoint holds it, and a run refuses a guest that reads anything else: a base
# quietly back in W. Europe would pass exactly the tests this zone is here to fail.
$BuildVmTimeZone = 'UTC'
$GuestRunRoot = 'C:\OutlookAI-Q5\run'
$GuestFeedRoot = 'C:\OutlookAI-Q5\nuget-offline'
$GuestScriptName = 'Invoke-BuildVmRun.ps1'
# The private list of personal data (Q196 (c)): where the guard reads it in a checkout, and where a run
# stages it in the VM - beside the run, outside its results. Invoke-BuildVmRun.ps1 moves it into the
# source tree's .work for the self-tests; -SelfTest holds the two scripts to the same name.
$PrivateListRelative = '.work\personal-data\identifiers.json'
$GuestPrivateListRelative = 'personal-data\identifiers.json'
$LeaseReasonPrefix = 'Invoke-TestsOnBuildVm'
$MaxClockSkewSeconds = 2

# Exit codes - see the banner.
$ExitCodes = [ordered]@{ PASS = 0; FAIL = 1; BUILD = 2; INFRA = 3; REFUSED = 4 }

# The zone pass's zone (Q144, decided by the maintainer 2026-10-11): the build VM stays in UTC, and the
# tests that can read the machine's zone run a second time in the zone of the maintainer's workstation
# and the two Outlook guests, which is not UTC and keeps daylight saving. Every run asks for it; the
# guest switches to it and back inside the run, on the base checkpoint.
$ZonePassTimeZone = 'W. Europe Standard Time'

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

# Every program call goes through Invoke-NativeCommand, so that Windows PowerShell 5.1 cannot end
# the script on a program's first stderr line (Q78). It is one of the helpers the scripts share,
# dot-sourced from Testbed/guest/ScriptHelpers.ps1 (Q166 (b), decision 215 (a)), which says why; so are
# Say, Stop-Refused, Format-Seconds, the run id and the TRX reader, which the live runner uses too.
. (Join-Path $PSScriptRoot '..\guest\ScriptHelpers.ps1')

# =============================================================================================
# PURE DECISIONS. No Hyper-V, no git, no file but the ones -SelfTest names. -SelfTest pins them.
# =============================================================================================

# The only VMs this script will act on: the build pool's. Compared whole and ordinal-ignore-case, as
# Hyper-V compares names; a prefix, a suffix, a clone the pool does not list or another testbed
# guest is refused.
function Test-IsBuildVmName {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -ne $Name.Trim()) { return $false }
    foreach ($member in $BuildVmPool) {
        if ([string]::Equals($Name, $member, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

# The pool members a run may take, in the order it tries them: the whole pool, or the one -VMName
# pins, spelled as the pool spells it. A name that is no build VM is refused.
function Get-PoolCandidates {
    param([string] $Pinned)
    if (-not $Pinned) { return @($BuildVmPool) }
    if (-not (Test-IsBuildVmName $Pinned)) { throw "'$Pinned' is not a build VM of the pool ($($BuildVmPool -join ', '))." }
    return @($BuildVmPool | Where-Object { [string]::Equals($_, $Pinned, [System.StringComparison]::OrdinalIgnoreCase) })
}

# Why the resumed guest is refused before anything is copied in - or '' when it is OAI-BUILD, logged
# on as the autologon account, in $BuildVmTimeZone. The zone is compared ordinal: Windows' ids are
# exact strings, and 'UTC' is not 'GMT Standard Time', which keeps British summer time.
function Get-GuestRefusal {
    param([string] $Computer, [string] $User, [string] $TimeZone)
    if (-not [string]::Equals($Computer, $BuildVmComputerName, [System.StringComparison]::OrdinalIgnoreCase) -or $User -ne $BuildVmUser) {
        return "the VM answers as '$User' on '$Computer', not '$BuildVmUser' on '$BuildVmComputerName'. Refusing to copy anything into it."
    }
    if (-not [string]::Equals($TimeZone, $BuildVmTimeZone, [System.StringComparison]::Ordinal)) {
        return "the VM's time zone is '$TimeZone', not '$BuildVmTimeZone' (Q126 (a)): '$BaseCheckpointName' is not the base Testbed/README.md section 1c takes. Refusing to run the suite in the wrong zone."
    }
    return ''
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
        [int]    $SelfTestsFailed,
        [bool]   $ZonePassRequested = $false,   # the run asked for the zone pass (Q144)
        [object] $ZonePass = $null,             # run.json's zonePass record, or $null
        [object] $ZoneTrx = $null               # ConvertFrom-TrxText's result for the zone pass, or $null
    )
    if (-not $GuestVerdict) {
        $why = 'the guest script left no verdict - no run.json came back'
        if ($InfraError) { $why = $InfraError }
        if ($SuiteRequested -and $null -ne $Trx -and ($Trx.Counts['failed'] + $Trx.Counts['other']) -gt 0) {
            return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$($Trx.Counts['failed'] + $Trx.Counts['other']) test(s) did not pass, by the suite's TRX file - the run did not finish: $why" }
        }
        if ($SuiteRequested -and $null -ne $ZoneTrx -and ($ZoneTrx.Counts['failed'] + $ZoneTrx.Counts['other']) -gt 0) {
            return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$($ZoneTrx.Counts['failed'] + $ZoneTrx.Counts['other']) test(s) did not pass in the zone pass, by its TRX file - the run did not finish: $why" }
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
        'SUITE-BROKEN' {
            if ($null -ne $ZonePass -and [string]$ZonePass.state -eq 'broken') { return [pscustomobject]@{ Verdict = 'FAIL'; Why = "the zone pass in $($ZonePass.timeZone) did not complete: $($ZonePass.why)" } }
            return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the suite hung, crashed or left no TRX file' }
        }
        'FAIL' { }
        'PASS' { }
        default { return [pscustomobject]@{ Verdict = 'INFRA'; Why = "the guest script left a verdict this runner does not know ('$GuestVerdict')" } }
    }
    if ($SuiteRequested) {
        if ($null -eq $Trx) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the suite ran and no TRX file came back' } }
        if ($Trx.Counts['total'] -eq 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the filter selected no test - nothing was tested' } }
        $zoneWhy = ''
        if ($ZonePassRequested) { $zoneWhy = Get-ZonePassFailure -ZonePass $ZonePass -ZoneTrx $ZoneTrx }
        if ($Trx.Counts['failed'] -gt 0 -or $Trx.Counts['other'] -gt 0) {
            $why = "$($Trx.Counts['failed'] + $Trx.Counts['other']) test(s) did not pass"
            if ($zoneWhy) { $why += "; $zoneWhy" }
            return [pscustomobject]@{ Verdict = 'FAIL'; Why = $why }
        }
        if ($zoneWhy) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = $zoneWhy } }
    }
    if ($SelfTestsFailed -gt 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = "$SelfTestsFailed self-test(s) failed" } }
    if ($GuestVerdict -eq 'FAIL') { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'the guest script reported a failure' } }
    if (-not $SuiteRequested -and $SelfTestsRun -eq 0) { return [pscustomobject]@{ Verdict = 'FAIL'; Why = 'no self-test was selected - nothing was tested' } }
    return [pscustomobject]@{ Verdict = 'PASS'; Why = '' }
}

# Why the zone pass fails a run whose suite ran, or '' when it does not (Q144). A pass that selected
# no test does not: a revision from before Q144 marks none, and a narrow -Filter may select none of
# the marked tests - T1/TimeZoneMarkerTests is what fails a test that should have been marked.
function Get-ZonePassFailure {
    param([object] $ZonePass, [object] $ZoneTrx)
    if ($null -eq $ZonePass) { return 'the zone pass left no record in run.json' }
    $zone = [string]$ZonePass.timeZone
    switch ([string]$ZonePass.state) {
        'broken' { return "the zone pass in $zone did not complete: $($ZonePass.why)" }
        'none-selected' { return '' }
        'skipped' { return "the zone pass in $zone did not run: $($ZonePass.why)" }
    }
    if ($null -eq $ZoneTrx) { return "the zone pass in $zone ran and no TRX file came back" }
    $bad = $ZoneTrx.Counts['failed'] + $ZoneTrx.Counts['other']
    if ($bad -gt 0) { return "$bad test(s) did not pass in $zone" }
    if ([string]$ZonePass.state -ne 'passed') { return "the zone pass in $zone ended '$($ZonePass.state)'" }
    return ''
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
# The private list a run stages: this checkout's, else the main checkout's; $null when neither has one.
function Find-PrivateList([string] $Checkout, [string] $MainCheckout) {
    foreach ($root in @($Checkout, $MainCheckout)) {
        if (-not $root) { continue }
        $candidate = Join-Path $root $PrivateListRelative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

# What a run says about the personal-data check: State for summary.json, Line for summary.txt, and
# Headline for the first line when the known identifiers were NOT checked - never a quiet pass.
function Get-PersonalDataState([bool] $SelfTestsSkipped, [string] $StagedFrom, [object] $GuestPlaced, [string[]] $LookedIn) {
    if ($SelfTestsSkipped) { return [pscustomobject]@{ State = 'self-tests skipped'; Line = $null; Headline = $null } }
    if (-not $StagedFrom) {
        return [pscustomobject]@{
            State    = 'not checked: no private list'
            Line     = "personal   NOT CHECKED: no private list at $($LookedIn -join ' or ') - check-personal-data's self-test held the tree to the generic shapes only, and the known identifiers were not checked"
            Headline = 'personal-data identifiers NOT checked (no private list)'
        }
    }
    if ([string]$GuestPlaced -ne 'placed') {
        return [pscustomobject]@{
            State    = 'not checked: staged, not placed'
            Line     = "personal   NOT CHECKED: the private list was staged but the VM did not put it in the source tree ('$GuestPlaced') - the known identifiers were not checked"
            Headline = 'personal-data identifiers NOT checked (the list did not reach the tree)'
        }
    }
    return [pscustomobject]@{
        State    = 'checked'
        Line     = 'personal   the known identifiers were checked against the private list, staged for this run only - never fetched back, and reverted with the VM'
        Headline = $null
    }
}

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
        [string]   $Results,
        [string]   $Vm,             # the pool member the run took; '' when it took none
        [bool]     $ZonePassRequested = $false,
        [object]   $ZonePass = $null,      # run.json's zonePass record, or $null
        [object]   $ZoneTrx = $null,       # ConvertFrom-TrxText's result for the zone pass, or $null
        [string]   $PersonalData = ''      # Get-PersonalDataState's State; '' when not known
    )
    $failed = @()
    $skipped = @()
    $suite = $null
    if ($null -ne $Trx) { $failed = @($Trx.Failed); $skipped = @($Trx.Skipped); $suite = $Trx.Counts }
    # The zone pass (Q144): null when the run did not ask for one, and its lists arrays at every length.
    $zonePassSummary = $null
    if ($ZonePassRequested) {
        $zoneFailed = @()
        $zoneSkipped = @()
        $zoneCounts = $null
        if ($null -ne $ZoneTrx) { $zoneFailed = @($ZoneTrx.Failed); $zoneSkipped = @($ZoneTrx.Skipped); $zoneCounts = $ZoneTrx.Counts }
        $zonePassSummary = [ordered]@{ timeZone = $ZonePassTimeZone; state = 'no record'; why = 'the zone pass left no record in run.json'; filter = $null; readBack = $null; restoredTo = $null; suite = $zoneCounts; failed = $zoneFailed; skipped = $zoneSkipped }
        if ($null -ne $ZonePass) {
            foreach ($k in @('timeZone', 'state', 'why', 'filter', 'readBack', 'restoredTo')) {
                if ($null -ne $ZonePass.$k) { $zonePassSummary[$k] = [string]$ZonePass.$k }
            }
        }
    }
    $ran = @($SelfTests | Where-Object { $null -ne $_ })
    $failures = @($ran | Where-Object { -not $_.passed } | ForEach-Object { [string]$_.path })
    $filterValue = $null
    if ($Filter) { $filterValue = $Filter }
    return [ordered]@{
        runId          = $RunId
        verdict        = $Verdict.Verdict
        why            = $Verdict.Why
        exitCode       = $ExitCodes[$Verdict.Verdict]
        vm             = $(if ($Vm) { $Vm } else { $null })
        revision       = $Revision
        filter         = $filterValue
        suite          = $suite
        failed         = $failed
        skipped        = $skipped
        zonePass       = $zonePassSummary
        selfTests      = [ordered]@{ run = $ran.Count; failed = $failures.Count; skippedByReason = @($SelfTestSkips | Where-Object { $null -ne $_ }).Count; failures = $failures }
        guestVerdict   = $GuestVerdict
        hostError      = $HostError
        buildErrors    = @($BuildErrors | Where-Object { $null -ne $_ })
        packagesStaged = $PackagesStaged
        timings        = $Timings
        guestPhases    = $GuestPhases
        results        = $Results
        personalData   = $(if ($PersonalData) { $PersonalData } else { $null })
    }
}

# summary.txt's lines for the zone pass (Q144): what it ran and how it came out, and the zone a new
# process read in it and after it - the evidence that it ran where it says, and that the self-tests
# after it ran in the VM's own zone again.
function Format-ZonePassLines {
    param([bool] $Requested, [object] $ZonePass, [object] $ZoneTrx)
    if (-not $Requested) { return @('zone pass  skipped, with the suite') }
    if ($null -eq $ZonePass) { return @("zone pass  $ZonePassTimeZone - not known: no zone pass came back in run.json") }
    $head = "zone pass  $($ZonePass.timeZone)"
    $state = [string]$ZonePass.state
    $out = @()
    if ($null -ne $ZoneTrx -and $ZoneTrx.Counts['total'] -gt 0) {
        $c = $ZoneTrx.Counts
        $other = ''
        if ($c['other'] -gt 0) { $other = ", $($c['other']) other" }
        $out += ('{0}: {1} total: {2} passed, {3} failed, {4} skipped{5} - filter {6}' -f $head, $c['total'], $c['passed'], $c['failed'], $c['skipped'], $other, $ZonePass.filter)
    }
    elseif ($state -eq 'none-selected' -or $null -ne $ZoneTrx) { $out += "${head}: no test this run selects is marked - filter $($ZonePass.filter)" }
    else {
        $why = [string]$ZonePass.why
        if (-not $why) { $why = 'its TRX file did not come back' }
        $out += "${head}: $state - $why"
    }
    if ($ZonePass.readBack -or $ZonePass.restoredTo) { $out += "           a new process read '$($ZonePass.readBack)' in it and '$($ZonePass.restoredTo)' after it" }
    return $out
}

# A list of failed tests for summary.txt, at most 50 named: the suite's, and the zone pass's.
function Add-FailedTests {
    param([System.Collections.Generic.List[string]] $Lines, [string] $Title, [object[]] $Failed, [string] $TrxFile)
    $all = @($Failed | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) { return }
    $Lines.Add('')
    $Lines.Add($Title)
    foreach ($f in @($all | Select-Object -First 50)) { $Lines.Add("  $($f.name)"); if ($f.message) { $Lines.Add("      $($f.message)") } }
    if ($all.Count -gt 50) { $Lines.Add("  ... and $($all.Count - 50) more - see $TrxFile") }
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
    Write-Host '== the only VMs it acts on: the build pool =='
    Check 'OutlookAI-Build is a build VM' $true (Test-IsBuildVmName 'OutlookAI-Build')
    Check 'in any case, as Hyper-V compares names' $true (Test-IsBuildVmName 'outlookai-build')
    Check 'the pool starts with OutlookAI-Build, the one VM every run took before the pool' 'OutlookAI-Build' $BuildVmPool[0]
    Check 'and every clone it lists is a build VM' 0 @($BuildVmPool | Where-Object { -not (Test-IsBuildVmName $_) }).Count
    Check 'OutlookAI-Indexed is not' $false (Test-IsBuildVmName 'OutlookAI-Indexed')
    Check 'OutlookAI-Unindexed is not' $false (Test-IsBuildVmName 'OutlookAI-Unindexed')
    Check 'nor an Outlook guest''s clone' $false (Test-IsBuildVmName 'OutlookAI-Indexed2')
    Check 'nor the Exchange VM' $false (Test-IsBuildVmName 'OutlookAI-Exchange')
    Check 'a clone the pool does not list is not' $false (Test-IsBuildVmName 'OutlookAI-Build9')
    Check 'nor a name that merely starts the same way' $false (Test-IsBuildVmName 'OutlookAI-Builder')
    Check 'nor one padded with a space' $false (Test-IsBuildVmName 'OutlookAI-Build ')
    Check 'nor an empty one' $false (Test-IsBuildVmName '')
    Check 'a run takes from the whole pool, in its order' ($BuildVmPool -join '|') ((Get-PoolCandidates -Pinned '') -join '|')
    Check '-VMName pins one member, spelled as the pool spells it' 'OutlookAI-Build2' ((Get-PoolCandidates -Pinned 'outlookai-build2') -join '|')
    $threw = ''
    try { [void](Get-PoolCandidates -Pinned 'OutlookAI-Indexed') } catch { $threw = $_.Exception.Message }
    Check 'and refuses any other VM' $true ($threw.Contains('not a build VM'))
    $lockNames = Get-BuildRunnerLockNames 'OutlookAI-Build'
    Check 'OutlookAI-Build keeps the lock every runner took before the pool, so older checkouts'' runs and these exclude each other' 'runner.lock|runner.owner.json' "$($lockNames.Lock)|$($lockNames.Owner)"
    $lockNames = Get-BuildRunnerLockNames 'OutlookAI-Build2'
    Check 'a clone has a lock and an owner file of its own' 'runner.OutlookAI-Build2.lock|runner.OutlookAI-Build2.owner.json' "$($lockNames.Lock)|$($lockNames.Owner)"
    Check 'another member starts only with its memory and the reserve to spare' 'True|False|False' ('{0}|{1}|{2}' -f (Test-HostHasRoom -NeedBytes 6GB -AvailableBytes 22GB -ReserveBytes 16GB), (Test-HostHasRoom -NeedBytes 6GB -AvailableBytes 21GB -ReserveBytes 16GB), (Test-HostHasRoom -NeedBytes 6GB -AvailableBytes -1 -ReserveBytes 16GB))
    Check 'the reserve kept for the maintainer is 16 GB' 16 ($TestbedHostReserveBytes / 1GB)
    Check 'OAI-BUILD as vmadmin in UTC is the guest a run expects' '' (Get-GuestRefusal -Computer 'OAI-BUILD' -User 'vmadmin' -TimeZone 'UTC')
    Check 'another computer is refused' $true ((Get-GuestRefusal -Computer 'OAI-INDEXED' -User 'vmadmin' -TimeZone 'UTC') -like "*not 'vmadmin' on 'OAI-BUILD'*")
    Check 'another account is refused' $true ((Get-GuestRefusal -Computer 'OAI-BUILD' -User 'tier' -TimeZone 'UTC') -like "*not 'vmadmin' on 'OAI-BUILD'*")
    Check 'a guest still in W. Europe - the base before Q126 (a) - is refused, naming both zones' $true ((Get-GuestRefusal -Computer 'OAI-BUILD' -User 'vmadmin' -TimeZone 'W. Europe Standard Time') -like "*'W. Europe Standard Time', not 'UTC'*")
    Check 'and so is GMT Standard Time, which is not UTC: it keeps British summer time' $true ((Get-GuestRefusal -Computer 'OAI-BUILD' -User 'vmadmin' -TimeZone 'GMT Standard Time') -ne '')
    Check 'and a zone the guest did not report' $true ((Get-GuestRefusal -Computer 'OAI-BUILD' -User 'vmadmin' -TimeZone '') -ne '')

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
    Check 'a duration never reads 60 seconds' '59 s|1 m 59 s' ((Format-Seconds 59.6), (Format-Seconds 119.6) -join '|')
    Check 'an include list arriving as one comma-joined string is split' 'Testbed/host/a.ps1 | Testbed/guest/*' (Split-IncludeList @('Testbed/host/a.ps1,Testbed/guest/*'))
    Check 'a real list is kept, spaces and empties dropped' 'x | y | z' (Split-IncludeList @('x', ' y ,', 'z'))
    Check 'no list stays no list' 0 @(Split-IncludeList @()).Count

    Write-Host ''
    Write-Host '== summary.json: every list is a JSON array at every length (2026-10-03: one failure came out as an object) =='
    $oneFailure = ConvertFrom-TrxText ($trxText -replace 'outcome="Timeout"', 'outcome="Passed"' -replace 'outcome="NotExecuted"', 'outcome="Passed"')
    $twoFailures = ConvertFrom-TrxText ($trxText -replace 'outcome="Timeout"', 'outcome="Failed"')
    function Get-SummaryJson($Trx, [object[]] $SelfTests, [object[]] $BuildErrors) {
        $v = Get-HostVerdict -GuestVerdict 'FAIL' -InfraError '' -Trx $Trx -SuiteRequested ($null -ne $Trx) -SelfTestsRun @($SelfTests).Count -SelfTestsFailed 0
        $s = New-RunSummary -RunId '20261003-091919-d7e58af90183' -Verdict $v -Revision ([ordered]@{ ref = 'HEAD' }) -Filter 'Category!=Live' -Trx $Trx -SelfTests $SelfTests -SelfTestSkips @() -GuestVerdict 'FAIL' -HostError '' -BuildErrors $BuildErrors -PackagesStaged $false -Timings ([ordered]@{ total = 1.0 }) -GuestPhases ([ordered]@{}) -Results 'C:\x' -Vm 'OutlookAI-Build2'
        return ($s | ConvertTo-Json -Depth 6)
    }
    $j = Get-SummaryJson $oneFailure @() @()
    Check 'it names the pool member the run took' 'OutlookAI-Build2' (($j | ConvertFrom-Json).vm)
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
    Write-Host '== the zone pass (Q144): the marked tests again in a second zone, in the verdict and both summaries =='
    Check 'the second zone is not the VM''s own' $true ($ZonePassTimeZone -ne $BuildVmTimeZone)
    $zoneKnown = $true
    try { [void][System.TimeZoneInfo]::FindSystemTimeZoneById($ZonePassTimeZone) } catch { $zoneKnown = $false }
    Check 'and Windows knows it, so the guest can switch to it' $true $zoneKnown
    function ZoneRecord([string] $State, [string] $Why = '') {
        return [pscustomobject]@{ timeZone = $ZonePassTimeZone; state = $State; why = $Why; filter = 'Category!=Live&TimeZone=Machine'; readBack = $ZonePassTimeZone; restoredTo = $BuildVmTimeZone }
    }
    function ZoneVerdict([string] $GuestVerdict = 'PASS', [object] $Trx = $green, [object] $ZonePass = (ZoneRecord 'passed'), [object] $ZoneTrx = $green, [bool] $Requested = $true, [string] $InfraError = '') {
        return (Get-HostVerdict -GuestVerdict $GuestVerdict -InfraError $InfraError -Trx $Trx -SuiteRequested $true -SelfTestsRun 5 -SelfTestsFailed 0 -ZonePassRequested $Requested -ZonePass $ZonePass -ZoneTrx $ZoneTrx)
    }
    $zoneOneFailure = ConvertFrom-TrxText ($trxText -replace 'outcome="Timeout"', 'outcome="Passed"')
    Check 'a green suite and a green zone pass is PASS' 'PASS' (ZoneVerdict).Verdict
    Check 'a test that fails only in the zone pass is FAIL' 'FAIL' (ZoneVerdict -GuestVerdict 'FAIL' -ZonePass (ZoneRecord 'failed') -ZoneTrx $zoneOneFailure).Verdict
    Check 'and the why names the zone it failed in' $true ((ZoneVerdict -GuestVerdict 'FAIL' -ZonePass (ZoneRecord 'failed') -ZoneTrx $zoneOneFailure).Why.Contains("1 test(s) did not pass in $ZonePassTimeZone"))
    Check 'failures in both passes are both named' '2 test(s) did not pass; 1 test(s) did not pass in W. Europe Standard Time' (ZoneVerdict -GuestVerdict 'FAIL' -Trx $trx -ZonePass (ZoneRecord 'failed') -ZoneTrx $zoneOneFailure).Why
    Check 'a zone pass that selected no test is no failure - a revision before Q144 marks none' 'PASS' (ZoneVerdict -ZonePass (ZoneRecord 'none-selected' 'no test this run selects is marked TimeZone=Machine') -ZoneTrx $null).Verdict
    Check 'nor is one whose TRX file holds no test' 'PASS' (ZoneVerdict -ZoneTrx $empty).Verdict
    Check 'a zone pass that could not switch the zone is FAIL, saying why' 'FAIL|True' ('{0}|{1}' -f (ZoneVerdict -GuestVerdict 'SUITE-BROKEN' -ZonePass (ZoneRecord 'broken' 'the VM did not reach the zone: Access is denied') -ZoneTrx $null).Verdict, (ZoneVerdict -GuestVerdict 'SUITE-BROKEN' -ZonePass (ZoneRecord 'broken' 'the VM did not reach the zone: Access is denied') -ZoneTrx $null).Why.Contains('did not complete: the VM did not reach the zone'))
    Check 'and a broken suite with a zone pass that never ran is still the suite''s fault' 'the suite hung, crashed or left no TRX file' (ZoneVerdict -GuestVerdict 'SUITE-BROKEN' -Trx $null -ZonePass (ZoneRecord 'skipped' 'the suite did not run to its end') -ZoneTrx $null).Why
    Check 'a requested zone pass that left no record is FAIL' 'FAIL' (ZoneVerdict -ZonePass $null -ZoneTrx $null).Verdict
    Check 'one that ran and brought back no TRX file is FAIL' 'FAIL' (ZoneVerdict -ZoneTrx $null).Verdict
    Check 'one that did not run although the suite did is FAIL' 'FAIL' (ZoneVerdict -ZonePass (ZoneRecord 'skipped' 'the suite did not run to its end') -ZoneTrx $null).Verdict
    Check 'a run that asked for none is judged without one' 'PASS' (ZoneVerdict -Requested $false -ZonePass $null -ZoneTrx $null).Verdict
    Check 'a run that did not finish, whose zone TRX came back failing, is FAIL - a test that failed failed' 'FAIL' (ZoneVerdict -GuestVerdict '' -InfraError $unfinished -ZonePass $null -ZoneTrx $zoneOneFailure).Verdict
    $z = New-RunSummary -RunId '20261117-120000-0123456789ab' -Verdict (ZoneVerdict -GuestVerdict 'FAIL' -ZonePass (ZoneRecord 'failed') -ZoneTrx $zoneOneFailure) -Revision ([ordered]@{ ref = 'HEAD' }) -Filter 'Category!=Live' -Trx $green -SelfTests @() -SelfTestSkips @() -GuestVerdict 'FAIL' -HostError '' -BuildErrors @() -PackagesStaged $false -Timings ([ordered]@{ total = 1.0 }) -GuestPhases ([ordered]@{}) -Results 'C:\x' -ZonePassRequested $true -ZonePass (ZoneRecord 'failed') -ZoneTrx $zoneOneFailure
    $j = $z | ConvertTo-Json -Depth 6
    Check 'summary.json: the zone pass''s one failed test is an array of one' $true ($j -match '"zonePass":\s*\{[\s\S]*"failed":\s*\[\s*\{')
    Check 'and reads back as one, with its zone, state and counts' "1|$ZonePassTimeZone|failed|4" ('{0}|{1}|{2}|{3}' -f @(($j | ConvertFrom-Json).zonePass.failed).Count, ($j | ConvertFrom-Json).zonePass.timeZone, ($j | ConvertFrom-Json).zonePass.state, ($j | ConvertFrom-Json).zonePass.suite.total)
    Check 'and its one skipped test an array of one' 1 @(($j | ConvertFrom-Json).zonePass.skipped).Count
    $j = (New-RunSummary -RunId '20261117-120000-0123456789ab' -Verdict (ZoneVerdict) -Revision ([ordered]@{ ref = 'HEAD' }) -Filter '' -Trx $null -SelfTests @() -SelfTestSkips @() -GuestVerdict 'PASS' -HostError '' -BuildErrors @() -PackagesStaged $false -Timings ([ordered]@{ total = 1.0 }) -GuestPhases ([ordered]@{}) -Results 'C:\x') | ConvertTo-Json -Depth 6
    Check 'a run that asked for no zone pass has none in summary.json' $true ($j -match '"zonePass":\s*null')
    $lines = @(Format-ZonePassLines -Requested $true -ZonePass (ZoneRecord 'failed') -ZoneTrx $zoneOneFailure)
    Check 'summary.txt: the zone pass''s counts and filter' "zone pass  ${ZonePassTimeZone}: 4 total: 2 passed, 1 failed, 1 skipped - filter Category!=Live&TimeZone=Machine" $lines[0]
    Check 'and the zones a new process read in it and after it' "           a new process read '$ZonePassTimeZone' in it and '$BuildVmTimeZone' after it" $lines[1]
    Check 'a pass that selected nothing says so' "zone pass  ${ZonePassTimeZone}: no test this run selects is marked - filter Category!=Live&TimeZone=Machine" (@(Format-ZonePassLines -Requested $true -ZonePass (ZoneRecord 'none-selected') -ZoneTrx $null)[0])
    Check 'a broken one says why' "zone pass  ${ZonePassTimeZone}: broken - the VM did not reach the zone: x" (@(Format-ZonePassLines -Requested $true -ZonePass (ZoneRecord 'broken' 'the VM did not reach the zone: x') -ZoneTrx $null)[0])
    Check 'a run without the suite has none' 'zone pass  skipped, with the suite' (@(Format-ZonePassLines -Requested $false -ZonePass $null -ZoneTrx $null)[0])
    $list = New-Object System.Collections.Generic.List[string]
    Add-FailedTests -Lines $list -Title 'FAILED IN THE ZONE PASS' -Failed @($zoneOneFailure.Failed) -TrxFile 'vm\trx\zone.trx'
    Check 'summary.txt lists the zone pass''s failed tests under their own title' 'FAILED IN THE ZONE PASS|  T1.A.Fails|      Assert.Equal() Failure' (@($list | Where-Object { $_ }) -join '|')
    $list = New-Object System.Collections.Generic.List[string]
    Add-FailedTests -Lines $list -Title 'FAILED TESTS' -Failed @() -TrxFile 'vm\trx\suite.trx'
    Check 'and adds nothing when none failed' 0 $list.Count

    Write-Host ''
    Write-Host '== the private list of personal data (Q196 (c)): staged when found, and a run without it says so =='
    $pdRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-buildvm-pd-' + [guid]::NewGuid().ToString('N'))
    try {
        $checkoutDir = Join-Path $pdRoot 'worktree'
        $mainDir = Join-Path $pdRoot 'main'
        New-Item -ItemType Directory -Force -Path (Join-Path $mainDir '.work\personal-data') | Out-Null
        New-Item -ItemType Directory -Force -Path $checkoutDir | Out-Null
        Check 'neither checkout has one: none' $true ($null -eq (Find-PrivateList $checkoutDir $mainDir))
        [System.IO.File]::WriteAllText((Join-Path $mainDir $PrivateListRelative), '{}')
        Check 'an agent''s worktree finds the main checkout''s' (Join-Path $mainDir $PrivateListRelative) (Find-PrivateList $checkoutDir $mainDir)
        New-Item -ItemType Directory -Force -Path (Join-Path $checkoutDir '.work\personal-data') | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $checkoutDir $PrivateListRelative), '{}')
        Check 'and its own first, when it has one' (Join-Path $checkoutDir $PrivateListRelative) (Find-PrivateList $checkoutDir $mainDir)
    }
    finally { if (Test-Path -LiteralPath $pdRoot) { Remove-Item -LiteralPath $pdRoot -Recurse -Force } }
    $none = Get-PersonalDataState $false '' $null @('C:\a\.work\personal-data\identifiers.json')
    Check 'no list: not checked, in summary.json' 'not checked: no private list' $none.State
    Check 'and on the summary''s first line' $true ([string]$none.Headline).Contains('NOT checked')
    Check 'and in its own line, naming where it looked' $true ($none.Line.StartsWith('personal   NOT CHECKED') -and $none.Line.Contains('C:\a\.work'))
    Check 'staged but not placed: not checked either' 'not checked: staged, not placed' (Get-PersonalDataState $false 'C:\a\x.json' 'absent' @()).State
    $ok = Get-PersonalDataState $false 'C:\a\x.json' 'placed' @()
    Check 'staged and placed: checked, nothing on the first line' 'checked |' "$($ok.State) |$($ok.Headline)"
    Check 'self-tests skipped: nothing to say' 'self-tests skipped' (Get-PersonalDataState $true '' $null @()).State
    Check 'the guest stages it where this stages it' $true ([System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "..\guest\$GuestScriptName")).Contains("'$GuestPrivateListRelative'"))
    $j = New-RunSummary -RunId 'x' -Verdict ([pscustomobject]@{ Verdict = 'PASS'; Why = '' }) -Revision ([ordered]@{}) -Filter '' -Trx $null -SelfTests @() -SelfTestSkips @() -GuestVerdict 'PASS' -HostError '' -BuildErrors @() -PackagesStaged $false -Timings ([ordered]@{}) -GuestPhases ([ordered]@{}) -Results 'C:\x' -Vm 'OutlookAI-Build' -PersonalData $none.State
    Check 'summary.json carries the state' 'not checked: no private list' $j.personalData

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
                    # The one lookup by another variable: Get-VM -Name $Name inside Find-PoolVm, which
                    # refuses any name that is not a build VM of the pool before Hyper-V is asked.
                    $enclosing = $c.Parent
                    while ($null -ne $enclosing -and $enclosing -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $enclosing = $enclosing.Parent }
                    $inFindPoolVm = ($null -ne $enclosing -and $enclosing.Name -eq 'Find-PoolVm')
                    $ok = ($value -eq '$BuildVmName') -or ($name -eq 'Get-VMSnapshot' -and $elements[$i].ParameterName -eq 'Name' -and $value -eq '$BaseCheckpointName') -or ($name -eq 'Get-VMIntegrationService' -and $elements[$i].ParameterName -eq 'Name') -or ($name -eq 'Get-VM' -and $value -eq '$Name' -and $inFindPoolVm)
                    if (-not $ok) { $loose += "$name -$($elements[$i].ParameterName) $value (line $($c.Extent.StartLineNumber))" }
                }
            }
        }
    }
    Check 'no command that stops, removes, reconfigures or checkpoints a VM, and no Stop-Process' '' ($used -join '; ')
    Check 'every -Name / -VMName handed to a VM command is $BuildVmName (or the base checkpoint''s name, or the pool lookup''s)' '' ($loose -join '; ')
    $vmLookups = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-VM' }, $true))
    Check 'a VM is looked up in exactly one place, Find-PoolVm' 1 $vmLookups.Count
    $findPoolVm = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Find-PoolVm' }, $true))
    Check 'and it refuses a name that is no build VM before it asks Hyper-V' $true ($findPoolVm.Count -eq 1 -and $findPoolVm[0].Body.Extent.Text.IndexOf('Test-IsBuildVmName') -ge 0 -and $findPoolVm[0].Body.Extent.Text.IndexOf('Test-IsBuildVmName') -lt $findPoolVm[0].Body.Extent.Text.IndexOf('Get-VM '))
    # The run's code only: this self-test asks the narrow queries too, to see them refuse.
    function Get-RunCalls([string] $Command) {
        return @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $Command }, $true) | Where-Object {
                $up = $_.Parent
                while ($null -ne $up -and -not ($up -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $up.Name -eq 'Invoke-SelfTest')) { $up = $up.Parent }
                $null -eq $up
            })
    }
    $getVms = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PoolVms' }, $true))
    Check 'members are found by name in exactly one place, Get-PoolVms, which refuses a name that is no build VM before it asks' $true ((Get-RunCalls 'Get-TestbedVms').Count -eq 1 -and $getVms.Count -eq 1 -and $getVms[0].Body.Extent.Text.IndexOf('Test-IsBuildVmName') -ge 0 -and $getVms[0].Body.Extent.Text.IndexOf('Test-IsBuildVmName') -lt $getVms[0].Body.Extent.Text.IndexOf('Get-TestbedVms'))
    Check 'and the state and the heartbeat by the held member''s id, one place each' '1|1' ("{0}|{1}" -f (Get-RunCalls 'Get-TestbedVmStateById').Count, (Get-RunCalls 'Wait-TestbedHeartbeat').Count)
    Check 'Get-VMIntegrationService is asked nothing any more' 0 (Get-RunCalls 'Get-VMIntegrationService').Count

    Write-Host ''
    Write-Host '== a run that waits asks Hyper-V nothing (2026-11-18, TestbedPool.ps1) =='
    Check 'a waiting run looks again at most every 30 s; a running one asks its guest every 15 s and a heartbeat every 2 s' '30|15|2' "$TestbedWaitSeconds|$TestbedGuestPollSeconds|$TestbedHeartbeatSeconds"
    $enterPoolVm = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Enter-PoolVm' }, $true))
    $body = ''
    if ($enterPoolVm.Count -eq 1) { $body = $enterPoolVm[0].Body.Extent.Text }
    $at = @{}
    foreach ($mark in @('Get-PoolVms $Candidates', 'Find-PoolVm $Candidates[$i]', '$ticket = Join-TestbedQueue', 'while ($true)', '$ahead = Get-TestbedQueueTurn', 'Get-TestbedLease -VMName $name', '$available = Get-HostAvailableBytes', '[System.IO.File]::Open($script:lockPath', '[void](Wait-TestbedChange)', 'finally { Exit-TestbedQueue $ticket }')) { $at[$mark] = $body.IndexOf($mark) }
    Check 'which members are on this host, their ids and startup memory are asked once, before the queue is joined' $true ($at['Get-PoolVms $Candidates'] -ge 0 -and $at['Find-PoolVm $Candidates[$i]'] -gt $at['Get-PoolVms $Candidates'] -and $at['$ticket = Join-TestbedQueue'] -gt $at['Find-PoolVm $Candidates[$i]'] -and $at['while ($true)'] -gt $at['$ticket = Join-TestbedQueue'])
    Check 'and never in a turn: no Hyper-V call after the turns begin' $true ($body.LastIndexOf('Get-PoolVms') -lt $at['while ($true)'] -and $body.LastIndexOf('Find-PoolVm') -lt $at['while ($true)'] -and $body.IndexOf('Get-VM') -lt 0)
    Check 'each turn reads its place in the queue, then a member''s lease, the host''s memory and the member''s lock' $true ($at['$ahead = Get-TestbedQueueTurn'] -gt $at['while ($true)'] -and $at['Get-TestbedLease -VMName $name'] -gt $at['$ahead = Get-TestbedQueueTurn'] -and $at['$available = Get-HostAvailableBytes'] -gt $at['Get-TestbedLease -VMName $name'] -and $at['[System.IO.File]::Open($script:lockPath'] -gt $at['$available = Get-HostAvailableBytes'])
    Check 'between turns it sleeps on the testbed''s files, never on a timer of its own, and leaves the queue on every path' $true ($at['[void](Wait-TestbedChange)'] -gt $at['[System.IO.File]::Open($script:lockPath'] -and $body.IndexOf('Start-Sleep') -lt 0 -and $at['finally { Exit-TestbedQueue $ticket }'] -gt $at['[void](Wait-TestbedChange)'])
    $memory = Get-HostAvailableBytes
    $poolText = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'TestbedPool.ps1'))
    $memAt = $poolText.IndexOf('function Get-HostAvailableBytes')
    $memBody = ''
    if ($memAt -ge 0) { $memBody = $poolText.Substring($memAt, [Math]::Min(2500, $poolText.Length - $memAt)) }
    Check 'the host''s memory is read without WMI, from GlobalMemoryStatusEx, and answers' $true ($memory -gt 0 -and $memBody.Contains('GlobalMemoryStatusEx') -and -not $memBody.Contains('Get-CimInstance'))
    $runText = [System.IO.File]::ReadAllText($PSCommandPath)
    $guestRun = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-GuestRun' }, $true))
    $guestDone = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Wait-GuestDone' }, $true))
    Check 'a running run asks its guest every 15 s, and so does the wait for a guest that outlives its run' $true ($guestRun.Count -eq 1 -and $guestRun[0].Body.Extent.Text.Contains('Start-Sleep -Seconds $TestbedGuestPollSeconds') -and $guestDone.Count -eq 1 -and $guestDone[0].Body.Extent.Text.Contains('Start-Sleep -Seconds $TestbedGuestPollSeconds') -and -not $guestRun[0].Body.Extent.Text.Contains('Start-Sleep -Seconds 3'))
    $waitAt = $poolText.IndexOf('function Wait-TestbedHeartbeat')
    $waitBody = ''
    if ($waitAt -ge 0) { $waitBody = $poolText.Substring($waitAt, [Math]::Min(600, $poolText.Length - $waitAt)) }
    Check 'the heartbeat is one narrow query every 2 s' $true ($waitBody.Contains('while (-not (Test-TestbedHeartbeat $Id))') -and $waitBody.Contains('Start-Sleep -Seconds $TestbedHeartbeatSeconds'))
    $restoreBase = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Restore-Base' }, $true))
    $restoreText = ''
    if ($restoreBase.Count -eq 1) { $restoreText = $restoreBase[0].Body.Extent.Text }
    Check 'a running VM is saved before the checkpoint is restored over it' $true ($restoreText.IndexOf("if (`$state -eq 'Running' -or `$state -eq 'Paused')") -ge 0 -and $restoreText.IndexOf('Save-VM -VM $vm') -gt 0 -and $restoreText.IndexOf('Save-VM -VM $vm') -lt $restoreText.IndexOf('Restore-VMSnapshot'))
    $runnerLock = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Enter-RunnerLock' }, $true))
    Check 'the janitor waits on its run''s process handle, and for a lock on the testbed''s files' $true ($runText.Contains('$ended = $watched.WaitForExit(60000)') -and $runnerLock.Count -eq 1 -and $runnerLock[0].Body.Extent.Text.Contains('[void](Wait-TestbedChange)') -and -not $runnerLock[0].Body.Extent.Text.Contains('Start-Sleep'))
    Check 'a VM''s state is Hyper-V''s EnabledState in Get-VM''s words: running, off, saved, paused, and the rest on the way' 'Running|Off|Saved|Paused|Saved|Paused|Starting|Saving|EnabledState 1' ((@(2, 3, 6, 9, 32769, 32768, 10, 32773, 1) | ForEach-Object { ConvertFrom-TestbedEnabledState $_ }) -join '|')
    $threw = ''
    try { [void](Get-TestbedVms -Name @("OutlookAI-Build' OR ElementName LIKE '%")) } catch { $threw = $_.Exception.Message }
    Check 'the narrow queries refuse a name that is not plain, before they ask' $true ($threw.Contains('not a plain VM name'))
    $threw = ''
    try { [void](Get-TestbedVmStateById "x' OR Name LIKE '%") } catch { $threw = $_.Exception.Message }
    Check 'and an id that is no GUID' $true ($threw.Contains('not a VM id'))
    Check 'and ask nothing for no name' 0 @(Get-TestbedVms -Name @()).Count
    $t = { param($name, $run, [string[]] $wants) [pscustomobject]@{ Name = $name; Run = $run; Candidates = $wants } }
    $pool = @('OutlookAI-Build', 'OutlookAI-Build2', 'OutlookAI-Build3')
    $turn = Select-TestbedQueueTurn -Tickets @((& $t '01-1.json' 'early' $pool), (& $t '02-2.json' 'mine' $pool)) -Mine '02-2.json' -Candidates $pool
    Check 'a run behind one that wants the whole pool may try no member yet' 'early|early|early' (@($pool | ForEach-Object { $turn[$_] }) -join '|')
    $turn = Select-TestbedQueueTurn -Tickets @((& $t '01-1.json' 'pinned' @('OutlookAI-Build2')), (& $t '02-2.json' 'mine' $pool)) -Mine '02-2.json' -Candidates $pool
    Check 'behind a run pinned to one member, it may try the others - no member waits on another''s queue' 'OutlookAI-Build2' (@($turn.Keys) -join '|')
    $turn = Select-TestbedQueueTurn -Tickets @((& $t '02-2.json' 'mine' $pool), (& $t '03-3.json' 'later' $pool)) -Mine '02-2.json' -Candidates $pool
    Check 'a run with nobody before it may try every member, whoever queued after it' 0 $turn.Count
    $turn = Select-TestbedQueueTurn -Tickets @((& $t '01-1.json' 'early' @('outlookai-build3')), (& $t '02-2.json' 'mine' $pool)) -Mine '02-2.json' -Candidates $pool
    Check 'names are compared as Hyper-V compares them' 'OutlookAI-Build3' (@($turn.Keys) -join '|')
    $turn = Select-TestbedQueueTurn -Tickets @((& $t '02-2.json' 'mine' $pool), (& $t '01-1.json' 'early' $pool)) -Mine '02-2.json' -Candidates $pool
    Check 'queue order is the tickets'' names, whatever order they are read in' 3 $turn.Count

    Write-Host ''
    Write-Host '== the queue''s files, on a scratch directory =='
    $queueScratch = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-queue-selftest-' + [guid]::NewGuid().ToString('N'))
    $script:TestbedQueueDirectoryOverride = $queueScratch
    try {
        $first = Join-TestbedQueue -Run 'r-first' -What 'self-test, first' -Candidates $pool
        Start-Sleep -Milliseconds 20
        $second = Join-TestbedQueue -Run 'r-second' -What 'self-test, second' -Candidates @('OutlookAI-Build2')
        Check 'a ticket is named for when it joined, so the names sort in queue order' $true ([string]::CompareOrdinal($first.Name, $second.Name) -lt 0 -and $first.Name.EndsWith("-$PID.json"))
        Check 'and holds its run and the members it wants, as an array at every length' 'r-second|True' ("{0}|{1}" -f ([System.IO.File]::ReadAllText($second.Path) | ConvertFrom-Json).run, ([System.IO.File]::ReadAllText($second.Path) -match '"candidates":\s*\['))
        $turn = Get-TestbedQueueTurn -Ticket $second -Run 'r-second' -Candidates @('OutlookAI-Build2')
        Check 'a live ticket before its own holds back the member both want' 'r-first' $turn['OutlookAI-Build2']
        Check 'the first ticket has nobody before it' 0 (Get-TestbedQueueTurn -Ticket $first -Run 'r-first' -Candidates $pool).Count
        $dead = Join-Path $queueScratch ('{0:D19}-4242.json' -f 1L)
        [System.IO.File]::WriteAllText($dead, (ConvertTo-Json -InputObject ([ordered]@{ pid = $PID; processStartUnix = 1; run = 'r-stopped'; what = 'x'; candidates = @($pool); joinedUnix = 1 })))
        $turn = Get-TestbedQueueTurn -Ticket $first -Run 'r-first' -Candidates $pool
        Check 'the ticket of a run that has ended holds nothing back, and is deleted' '0|False' ("{0}|{1}" -f $turn.Count, (Test-Path -LiteralPath $dead))
        $ahead = Join-TestbedQueue -Run 'r-clone' -What 'self-test, ahead' -Candidates @('OutlookAI-Build') -Ahead
        Check 'a ticket that goes ahead sorts before every other' $true ([string]::CompareOrdinal($ahead.Name, $first.Name) -lt 0)
        Check 'and holds back the member it wants from the first run' 'r-clone' (Get-TestbedQueueTurn -Ticket $first -Run 'r-first' -Candidates $pool)['OutlookAI-Build']
        Exit-TestbedQueue $ahead
        Remove-Item -LiteralPath $first.Path -Force
        [void](Get-TestbedQueueTurn -Ticket $first -Run 'r-first' -Candidates $pool)
        Check 'a run whose ticket has gone writes it again under its name, and keeps its place' $true (Test-Path -LiteralPath $first.Path)
        Exit-TestbedQueue $first
        Exit-TestbedQueue $second
        Check 'leaving the queue deletes the ticket' 0 @(Get-ChildItem -LiteralPath $queueScratch -Filter '*.json').Count
    }
    finally {
        $script:TestbedQueueDirectoryOverride = ''
        Remove-Item -LiteralPath $queueScratch -Recurse -Force -ErrorAction SilentlyContinue
    }
    $t0 = Get-Date
    $changed = Wait-TestbedChange -Seconds 1
    Check 'with nothing changing, the wait ends when its time is up' 'False|True' ("{0}|{1}" -f $changed, (((Get-Date) - $t0).TotalSeconds -lt 5))
    $poke = Join-Path (Split-Path -Parent (Get-BuildRunnerDirectory)) ('selftest-poke-' + [guid]::NewGuid().ToString('N') + '.json')
    $job = Start-Job -ArgumentList $poke -ScriptBlock { param($p) Start-Sleep -Seconds 2; [System.IO.File]::WriteAllText($p, '{}') }
    try {
        $t0 = Get-Date
        $changed = Wait-TestbedChange -Seconds 60
        Check 'a file of the testbed''s state written meanwhile ends it at once, not at its time limit' 'True|True' ("{0}|{1}" -f $changed, (((Get-Date) - $t0).TotalSeconds -lt 45))
    }
    finally {
        Wait-Job $job -Timeout 60 | Out-Null
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $poke -Force -ErrorAction SilentlyContinue
    }

    Write-Host ''
    Write-Host '== this file, the guest script and testbed.json agree =='
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $testbedPath = Join-Path $repoRoot 'Testbed\testbed.json'
    $guestPath = Join-Path $repoRoot (Join-Path 'Testbed\guest' $GuestScriptName)
    $buildVm = $null
    $testbedJson = $null
    if (Test-Path -LiteralPath $testbedPath) { $testbedJson = [System.IO.File]::ReadAllText($testbedPath) | ConvertFrom-Json; $buildVm = $testbedJson.buildVm }
    Check 'testbed.json has a buildVm block' $true ($null -ne $buildVm)
    if ($null -ne $buildVm) {
        Check 'testbed.json names the same VM as the pool''s source' $BuildVmPool[0] $buildVm.vmName
        Check 'and the same computer name' $BuildVmComputerName $buildVm.computerName
        Check 'and the same base checkpoint' $BaseCheckpointName $buildVm.runnerBaseCheckpoint
        Check 'and the base checkpoint is one of its checkpoints' $true (@($buildVm.checkpoints) -contains $BaseCheckpointName)
        Check 'and the same time zone (Q126 (a))' $BuildVmTimeZone $buildVm.timeZone
        Check 'and the same run directory' $GuestRunRoot $buildVm.guest.runRoot
        Check 'and the same offline feed' $GuestFeedRoot $buildVm.guest.offlineFeed
    }
    # TestbedPool.ps1 lists the pools; testbed.json's vmPools records each clone. They must agree, and
    # each clone must have been made from the checkpoint its source's runner restores now.
    $vmPools = $null
    if ($null -ne $testbedJson) { $vmPools = $testbedJson.vmPools }
    Check 'testbed.json has a vmPools block' $true ($null -ne $vmPools)
    if ($null -ne $vmPools) {
        foreach ($pool in $TestbedPools.Keys) {
            $record = $vmPools.$pool
            $clones = @()
            $recorded = @()
            if ($null -ne $record) { $clones = @(@($record.clones) | Where-Object { $null -ne $_ }); $recorded = @([string]$record.source) + @($clones | ForEach-Object { [string]$_.vmName }) }
            Check "the $pool pool is the one testbed.json records, its source first" (@($TestbedPools[$pool]) -join '|') ($recorded -join '|')
            $from = ''
            if ($pool -eq 'build' -and $null -ne $buildVm) { $from = [string]$buildVm.runnerBaseCheckpoint }
            elseif ($null -ne $testbedJson.frozenClocks -and $null -ne $record) { $from = [string]$testbedJson.frozenClocks.([string]$record.source).checkpoint }
            foreach ($clone in $clones) {
                Check "$($clone.vmName) was made from the checkpoint its source's runner restores ($from), and rests on that name" "$from|$from" "$($clone.fromCheckpoint)|$($clone.checkpoint)"
            }
        }
    }
    $guestText = ''
    if (Test-Path -LiteralPath $guestPath) { $guestText = [System.IO.File]::ReadAllText($guestPath) }
    Check 'the guest script exists beside this one' $true ($guestText.Length -gt 0)
    Check 'its run directory is this one' $true ($guestText.Contains("[string]   `$RunRoot = '$GuestRunRoot'"))
    Check 'its offline feed is this one' $true ($guestText.Contains("[string]   `$OfflineFeed = '$GuestFeedRoot'"))
    Check 'its guard names this computer' $true ($guestText.Contains("[string]   `$ExpectedComputerName = '$BuildVmComputerName'"))
    Check 'and this account' $true ($guestText.Contains("[string[]] `$ExpectedUser = @('$BuildVmUser')"))
    Check 'it reads the zone pass''s zone from the request this file writes' $true ($guestText.Contains('zonePassTimeZone') -and $guestText.Contains('Set-TimeZone -Id $zone'))

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
$lockDir = Get-BuildRunnerDirectory
$lockPath = ''
$ownerPath = ''
$historyPath = Join-Path $lockDir 'runs.log'
$vm = $null
$checkpoint = $null

# The pool member this run or janitor holds from here on: its name, and its lock and owner file.
function Use-PoolVm([string] $Name) {
    $script:BuildVmName = $Name
    $names = Get-BuildRunnerLockNames $Name
    $script:lockPath = Join-Path $lockDir $names.Lock
    $script:ownerPath = Join-Path $lockDir $names.Owner
}

# Who holds a member's lock, from its owner file - what a queued run says while it waits.
function Get-LockHolder([string] $OwnerFile) {
    try {
        $o = [System.IO.File]::ReadAllText($OwnerFile) | ConvertFrom-Json
        return "run $($o.runId) (pid $($o.pid), $($o.ref) in $($o.repo), started $($o.startedLocal))"
    }
    catch { return 'another run' }
}

# A RUN TAKES WHICHEVER BUILD VM IS FREE (2026-11-17), AND WAITS FOR ONE ASKING HYPER-V NOTHING
# (2026-11-18, TestbedPool.ps1's banner). Which members are on this host, their ids and the startup
# memory of each but the first are asked once, before the wait. Then the run joins the queue, and each
# turn tries - in pool order - the members no run that queued before it wants. A member is passed over
# while somebody holds a lease on it by hand (Set-TestbedLease.ps1, for maintenance), while - any
# member but the first - the host lacks its startup memory and the reserve TestbedPool.ps1 keeps for
# the maintainer, or while a run holds its lock: files and GlobalMemoryStatusEx, no Hyper-V and no
# WMI in a turn at all. Once the lock is ours the lease is read again and taken in one step
# (Enter-TestbedLeaseClaim), so a lease taken by hand in between still holds the VM back. Between turns
# the run sleeps until a lock's owner file, a lease or a ticket changes, or 30 s pass. What holds each
# member back is said once a minute. Returns the lock; $null when the time ran out, with the reasons
# in $script:queueWhy; throws when no candidate is on this host at all.
function Enter-PoolVm([string[]] $Candidates, [int] $TimeoutMinutes) {
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $lastSaid = [DateTime]::MinValue
    $queued = $false
    # What Hyper-V must say, said once, before the run waits: which members are on this host and their
    # ids - one narrow query - and the startup memory of each but the first, for the reserve.
    $ids = @{}
    foreach ($row in @(Get-PoolVms $Candidates)) { if (-not $ids.ContainsKey($row.Name)) { $ids[$row.Name] = $row.Id } }
    $here = @($Candidates | Where-Object { $ids.ContainsKey($_) })
    if ($here.Count -eq 0) { throw "no build VM of the pool is on this host ($($Candidates -join ', ')); Testbed/README.md section 1c builds OutlookAI-Build." }
    $startupBytes = @{}
    for ($i = 1; $i -lt $Candidates.Count; $i++) {
        if ($here -notcontains $Candidates[$i]) { continue }
        $found = @(Find-PoolVm $Candidates[$i])
        if ($found.Count -eq 1) { $startupBytes[$Candidates[$i]] = [long]$found[0].MemoryStartup }
    }
    $ticket = Join-TestbedQueue -Run $runId -What "Invoke-TestsOnBuildVm run $runId, $Ref in $repo" -Candidates $here
    try {
        while ($true) {
            $why = New-Object System.Collections.Generic.List[string]
            $ahead = Get-TestbedQueueTurn -Ticket $ticket -Run $runId -Candidates $here
            for ($i = 0; $i -lt $Candidates.Count; $i++) {
                $name = $Candidates[$i]
                if ($here -notcontains $name) { $why.Add("$name is not on this host"); continue }
                if ($ahead.ContainsKey($name)) { $why.Add("$name is wanted first by run $($ahead[$name]), queued before this one"); continue }
                $lease = Get-TestbedLease -VMName $name
                if ($null -ne $lease -and -not ([string]$lease.reason).StartsWith($LeaseReasonPrefix)) { $why.Add("$name is leased by hand until $($lease.expiresUtc) - '$($lease.reason)'"); continue }
                if ($i -gt 0) {
                    # The reserve: the startup memory read before the wait, the host's memory now - no WMI.
                    if (-not $startupBytes.ContainsKey($name)) { $why.Add("$name is passed over: Hyper-V did not say its startup memory before the wait"); continue }
                    $need = [long]$startupBytes[$name]
                    $available = Get-HostAvailableBytes
                    if (-not (Test-HostHasRoom -NeedBytes $need -AvailableBytes $available)) {
                        $why.Add(("{0} waits for memory: the host has {1:N1} GB available, under its {2:N0} GB and the {3:N0} GB kept free" -f $name, ($available / 1GB), ($need / 1GB), ($TestbedHostReserveBytes / 1GB)))
                        continue
                    }
                }
                Use-PoolVm $name
                $held = $null
                try { $held = [System.IO.File]::Open($script:lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None) }
                catch [System.IO.IOException] { $why.Add("$name is busy with $(Get-LockHolder $script:ownerPath)"); continue }
                $script:BuildVmId = $ids[$name]
                $claim = Enter-TestbedLeaseClaim -VMName $name
                $byHand = $null
                try {
                    $lease = Get-TestbedLease -VMName $name
                    if ($null -ne $lease -and -not ([string]$lease.reason).StartsWith($LeaseReasonPrefix)) { $byHand = $lease }
                    else {
                        $leaseOutput = & (Join-Path $PSScriptRoot 'Set-TestbedLease.ps1') -VMName $name -Minutes ([Math]::Min(480, $RunTimeoutMinutes + 20)) -Reason "$LeaseReasonPrefix $runId"
                        $script:leaseTaken = $true
                        Write-Verbose ([string]$leaseOutput)
                    }
                }
                finally { $claim.Dispose() }
                if ($null -ne $byHand) {
                    $held.Dispose()
                    $why.Add("$name is leased by hand until $($byHand.expiresUtc) - '$($byHand.reason)'")
                    continue
                }
                return $held
            }
            $script:queueWhy = $why -join '; '
            if ((Get-Date) -ge $deadline) { return $null }
            if (((Get-Date) - $lastSaid).TotalSeconds -ge 60) {
                if (-not $queued) { Say "queued: $($script:queueWhy)" } else { Say "  still waiting: $($script:queueWhy)" }
                $queued = $true
                $lastSaid = Get-Date
            }
            [void](Wait-TestbedChange)
        }
    }
    finally { Exit-TestbedQueue $ticket }
}

# The janitor's wait for its VM's lock: no poll - it sleeps until a lock's owner file changes, or 30 s pass.
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
            [void](Wait-TestbedChange)
        }
    }
}

# One line of runs.log, which every run and janitor on the pool appends to. Appended with sharing and
# retried, never Add-Content: with several runs ending at once, Windows PowerShell 5.1's Add-Content
# - which shuts every other writer out while it writes - would lose an end line, and a run's janitor
# reads a missing end line as a run that died part-way.
function Write-History([string] $Line) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $Line) + [Environment]::NewLine)
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        try {
            $fs = [System.IO.File]::Open($historyPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            try { $fs.Write($bytes, 0, $bytes.Length) } finally { $fs.Dispose() }
            return
        }
        catch { Start-Sleep -Milliseconds (50 * $attempt) }
    }
}

# The one place a VM is looked up: a member of the build pool, by name, refused before Hyper-V is
# asked when it is no build VM, and checked again on what Hyper-V returns.
function Find-PoolVm([string] $Name) {
    if (-not (Test-IsBuildVmName $Name)) { throw "'$Name' is not a build VM of the pool ($($BuildVmPool -join ', ')). Refusing to look it up." }
    return @(Get-VM -Name $Name -ErrorAction SilentlyContinue | Where-Object { Test-IsBuildVmName $_.Name })
}

function Get-BuildVm {
    $found = @(Find-PoolVm $BuildVmName)
    if ($found.Count -ne 1) { throw "expected one VM named '$BuildVmName' on this host and found $($found.Count). Testbed/README.md section 1c builds OutlookAI-Build; New-TestbedClone.ps1 makes its clones." }
    return $found[0]
}

# The one place members are found by name in the narrow way: each refused before Hyper-V is asked when
# it is no build VM, then one query for all of them (TestbedPool.ps1's Get-TestbedVms) - name, id and
# state - and only rows that are build VMs back.
function Get-PoolVms([string[]] $Names) {
    foreach ($n in @($Names)) { if (-not (Test-IsBuildVmName $n)) { throw "'$n' is not a build VM of the pool ($($BuildVmPool -join ', ')). Refusing to ask about it." } }
    return @(Get-TestbedVms -Name $Names | Where-Object { Test-IsBuildVmName $_.Name })
}

# The Hyper-V id of the member this run or janitor holds, asked once (Enter-PoolVm sets it as it
# takes the member; the janitor asks here). Not TestbedPool.ps1's Get-TestbedVmId, which the guest
# runners ask: this one asks only through the pool's allowlist, Get-PoolVms.
function Get-BuildVmId {
    if ($script:BuildVmId) { return $script:BuildVmId }
    $rows = @(Get-PoolVms @($BuildVmName))
    if ($rows.Count -ne 1) { throw "expected one VM named '$BuildVmName' on this host and found $($rows.Count). Testbed/README.md section 1c builds OutlookAI-Build; New-TestbedClone.ps1 makes its clones." }
    $script:BuildVmId = [string]$rows[0].Id
    return $script:BuildVmId
}

# The state of the member this run or janitor holds - Running, Saved, Off, Paused, or on its way -
# by its id, one keyed query.
function Get-BuildVmState {
    $state = Get-TestbedVmStateById (Get-BuildVmId)
    if (-not $state) { throw "Hyper-V no longer has '$BuildVmName' (id $(Get-BuildVmId)) on this host." }
    return [string]$state
}

function Get-BaseCheckpoint {
    $found = @(Get-VMSnapshot -VM $vm -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $BaseCheckpointName })
    if ($found.Count -ne 1) { throw "'$BuildVmName' has $($found.Count) checkpoint(s) named '$BaseCheckpointName', not one. Testbed/README.md section 1c takes it." }
    return $found[0]
}

# The VM's state, read afresh each time: a VirtualMachine object keeps the state it was read with
# and has no Refresh(), so a loop over one object would wait for ever on a stale value. Read with the
# narrow query, every 2 s while a restore or a save moves it - seconds, and a run waits on it.
function Wait-VmSettled([int] $Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        $state = Get-BuildVmState
        if (@('Running', 'Off', 'Saved', 'Paused') -contains $state) { return $state }
        Start-Sleep -Seconds 2
    }
    return (Get-BuildVmState)
}

# A VM that runs - or is paused - is SAVED before the checkpoint is restored over it (2026-11-18): a
# restore over a running guest left it out of PowerShell Direct's reach for six minutes in stream R1's
# runs, where saved first and then restored it answered in seconds. A restore over a saved VM leaves
# it saved, and Start-VM resumes it.
function Restore-Base {
    $state = Wait-VmSettled 120
    if (@('Running', 'Off', 'Saved', 'Paused') -notcontains $state) { throw "the VM stayed '$state' for two minutes; not restoring a checkpoint over a VM in transition." }
    if ($state -eq 'Running' -or $state -eq 'Paused') {
        Save-VM -VM $vm
        $state = Wait-VmSettled 120
        if ($state -ne 'Saved' -and $state -ne 'Off') { throw "the VM is '$state' after Save-VM; not restoring a checkpoint over it." }
    }
    Restore-VMSnapshot -VMSnapshot $checkpoint -Confirm:$false
    return (Wait-VmSettled 60)
}

# Restores the base checkpoint and leaves the VM saved: Restore-Base saves a running VM before it
# restores, so the base comes back saved; a base that resumes all the same is saved here - the BASE,
# not the run - which is what gives the host its RAM back.
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
    # The pool member the run took - and only a member: the janitor restores checkpoints.
    if (-not (Test-IsBuildVmName $JanitorVmName)) { Write-JanitorLine "refused: '$JanitorVmName' is not a build VM of the pool"; exit $ExitCodes['REFUSED'] }
    Use-PoolVm (@(Get-PoolCandidates -Pinned $JanitorVmName)[0])

    # The run's process: a process id reused after it ends is not mistaken for it, because the wait
    # below holds the process's own handle.
    $watched = Get-Process -Id $JanitorForPid -ErrorAction SilentlyContinue
    $giveUp = (Get-Date).AddHours(6)
    # Waited on through the process itself - its handle, held from here, cannot come to mean another
    # process - so no poll at all, where it looked every 10 s until 2026-11-18.
    while ($null -ne $watched) {
        if ((Get-Date) -ge $giveUp) { Write-JanitorLine 'the run is still going after six hours; leaving it'; exit 0 }
        $ended = $true
        try { $ended = $watched.WaitForExit(60000) } catch { }
        if ($ended) { $watched = $null }
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
            Write-JanitorLine "restored '$BaseCheckpointName' over the dead run's VM, $BuildVmName ($state); it is now $final"
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
function Measure-Phase([string] $Name, [DateTime] $Since) { $timings[$Name] = [math]::Round(((Get-Date) - $Since).TotalSeconds, 1) }

if ($QueueTimeoutMinutes -lt 1 -or $RunTimeoutMinutes -lt 5) { Stop-Refused '-QueueTimeoutMinutes must be at least 1 and -RunTimeoutMinutes at least 5.' }
if ($SkipSuite -and $SkipSelfTests) { Stop-Refused '-SkipSuite and -SkipSelfTests together leave nothing to run.' }
$candidates = @()
try { $candidates = @(Get-PoolCandidates -Pinned $VMName) } catch { Stop-Refused $_.Exception.Message }

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

# The private list of personal data, for check-personal-data's self-test (Q196 (c)). Only its path is
# kept here: the file goes from where it lies straight to the VM, and nothing on this host copies it.
$privateListLookedIn = @((Join-Path $repo $PrivateListRelative), (Join-Path $CredentialRepoRoot $PrivateListRelative)) | Select-Object -Unique
$privateList = $null
if (-not $SkipSelfTests) { $privateList = Find-PrivateList $repo $CredentialRepoRoot }

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
if (-not $SkipSuite) { Say "  zone pass the tests that can read the machine's zone, again in $ZonePassTimeZone (Q144)" }
$SelfTestInclude = @(Split-IncludeList $SelfTestInclude)
if ($SkipSelfTests) { Say '  selftests skipped (-SkipSelfTests)' }
elseif (@($SelfTestInclude).Count -gt 0) { Say "  selftests only $(@($SelfTestInclude) -join ', ')" }
else { Say '  selftests every script under Testbed\ and Tools\ that has -SelfTest' }
if (-not $SkipSelfTests) {
    if ($privateList) { Say "  personal  the private list is staged for check-personal-data's self-test (from $(Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $privateList))))" }
    else { Say "  personal  NO private list at $($privateListLookedIn -join ' or ') - the known identifiers will NOT be checked" }
}
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
    zonePassTimeZone = $ZonePassTimeZone
    revision        = [ordered]@{ ref = $Ref; sha = $sha; subject = $subject; repo = $repo; uncommittedLeftOut = $dirtyCount; runId = $runId }
}
$requestPath = Join-Path $runDir 'request.json'
($request | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $requestPath -Encoding UTF8

# ---- a build VM of the pool: its lock, then its lease (Enter-PoolVm) ---------------------------
$t0 = Get-Date
$leaseTaken = $false
$lock = $null
try { $lock = Enter-PoolVm -Candidates $candidates -TimeoutMinutes $QueueTimeoutMinutes }
catch {
    Say "INFRA: $($_.Exception.Message) This run did not start; nothing was touched."
    exit $ExitCodes['INFRA']
}
Measure-Phase 'queue' $t0
if ($null -eq $lock) {
    Say "INFRA: no build VM came free in $QueueTimeoutMinutes minute(s) - $($script:queueWhy). This run did not start; nothing on a VM was touched."
    exit $ExitCodes['INFRA']
}
[ordered]@{ runId = $runId; pid = $PID; ref = $Ref; sha = $sha; repo = $repo; vm = $BuildVmName; startedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') } |
    ConvertTo-Json | Set-Content -LiteralPath $ownerPath -Encoding UTF8
$queueText = ''
if ($timings['queue'] -ge 5) { $queueText = " after $(Format-Seconds $timings['queue']) in the queue" }
Say "VM: $BuildVmName is ours$queueText"
Write-History "start  $runId  $Ref  $repo  $BuildVmName"

# The janitor, for the case this process is stopped part-way (the banner). Through WMI, so it is
# no child of this process and survives whatever stops it; hidden, so no console appears. A run
# that cannot start one still runs, and says so.
if (-not $KeepVmRunning) {
    try {
        $shellExe = (Get-Process -Id $PID).Path
        $janitorCommand = '"{0}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}" -JanitorForPid {2} -JanitorRunId {3} -JanitorVmName {4}' -f $shellExe, $PSCommandPath, $PID, $runId, $BuildVmName
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
    # Only for a script that has died does a busy log stop counting, after about half a minute of
    # polls: there is nothing more to wait for, and the results are fetched as they are.
    # The guest is asked every $TestbedGuestPollSeconds (15 s, TestbedPool.ps1) - every 3 s until
    # 2026-11-18 - so its log arrives in batches of that length.
    $failedPolls = 0
    $busyPolls = 0
    $busyLimit = [Math]::Max(2, [int][Math]::Ceiling(30 / $TestbedGuestPollSeconds))
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
        if (-not $poll.Alive -and (-not $poll.Busy -or $busyPolls -ge $busyLimit)) {
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
        Start-Sleep -Seconds $TestbedGuestPollSeconds
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
        Start-Sleep -Seconds $TestbedGuestPollSeconds
    }
}

try {
    # The lease is ours already (Enter-PoolVm): an operator's own lease on a VM holds runs off it, and
    # ours keeps the idle-saver off this one. The VM, from its base checkpoint.
    $t0 = Get-Date
    $vm = Get-BuildVm
    $checkpoint = Get-BaseCheckpoint
    $stateBefore = [string]$vm.State
    $stateAfter = Restore-Base
    Say "VM: restored '$BaseCheckpointName' ($stateBefore -> $stateAfter)"
    # A checkpoint applied to a RUNNING VM resumes it at once (measured on the first run,
    # 2026-10-03: Running -> Running); applied to a saved or stopped one it leaves it saved.
    if ($stateAfter -ne 'Running') { Start-VM -VM $vm }
    # The heartbeat, every 2 s, by one narrow query on the VM's id (TestbedPool.ps1's
    # Wait-TestbedHeartbeat) - Get-VMIntegrationService, every second until 2026-11-18, cost 23 WMI
    # operations a look.
    Wait-TestbedHeartbeat -Id (Get-BuildVmId) -Minutes 3 -What 'the VM'
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
    $who = Invoke-InGuest -Block { [pscustomobject]@{ Computer = $env:COMPUTERNAME; User = $env:USERNAME; TimeZone = (Get-TimeZone).Id } }
    $refusal = Get-GuestRefusal -Computer ([string]$who.Computer) -User ([string]$who.User) -TimeZone ([string]$who.TimeZone)
    if ($refusal) { throw $refusal }
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
    Say "VM: $($who.Computer) as $($who.User), time zone $($who.TimeZone), clock within $skew s of this host, $(Format-Seconds ($timings['resume'] + $timings['connect'])) after the restore"

    # Stage the run.
    $t0 = Get-Date
    Invoke-InGuest -ArgumentList @($GuestRunRoot) -Block {
        param($root)
        if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
        New-Item -ItemType Directory -Force -Path $root | Out-Null
        New-Item -ItemType Directory -Force -Path (Join-Path $root 'personal-data') | Out-Null
    }
    Copy-Item -LiteralPath $archive -Destination (Join-Path $GuestRunRoot 'Source.zip') -ToSession $session
    Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) (Join-Path 'guest' $GuestScriptName)) -Destination (Join-Path $GuestRunRoot $GuestScriptName) -ToSession $session
    Copy-Item -LiteralPath $requestPath -Destination (Join-Path $GuestRunRoot 'request.json') -ToSession $session
    if ($privateList) { Copy-Item -LiteralPath $privateList -Destination (Join-Path $GuestRunRoot $GuestPrivateListRelative) -ToSession $session }
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
                foreach ($rel in @('run.log', 'run.json', 'trx\suite.trx', 'trx\zone.trx', 'restore.out.txt', 'build.out.txt')) {
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
function Read-TrxFile([string] $Path, [string] $What) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (ConvertFrom-TrxText ([System.IO.File]::ReadAllText($Path))) }
    catch {
        if (-not $script:infraError) { $script:infraError = "$What came back unreadable: $($_.Exception.Message)" }
        return $null
    }
}
$trx = Read-TrxFile (Join-Path $vmDir 'trx\suite.trx') 'the TRX file'
# The zone pass (Q144): its TRX file and its record in run.json.
$zoneTrx = Read-TrxFile (Join-Path $vmDir 'trx\zone.trx') "the zone pass's TRX file"
$zonePassRequested = -not $SkipSuite
$zonePass = $null
$selfTests = @()
$selfTestSkips = @()
if ($null -ne $guestRecord) { $selfTests = @($guestRecord.selfTests); $selfTestSkips = @($guestRecord.selfTestSkips); $zonePass = $guestRecord.zonePass }
$selfTestsFailed = @($selfTests | Where-Object { -not $_.passed }).Count
$hostVerdict = Get-HostVerdict -GuestVerdict $guestVerdict -InfraError $infraError -Trx $trx -SuiteRequested (-not $SkipSuite) -SelfTestsRun $selfTests.Count -SelfTestsFailed $selfTestsFailed -ZonePassRequested $zonePassRequested -ZonePass $zonePass -ZoneTrx $zoneTrx
$timings['total'] = [math]::Round(((Get-Date) - $tStart).TotalSeconds, 1)
$exitCode = $ExitCodes[$hostVerdict.Verdict]

# The guest's own phase timings, beside the host's.
$guestPhases = [ordered]@{}
if ($null -ne $guestRecord -and $null -ne $guestRecord.phases) {
    foreach ($p in $guestRecord.phases.PSObject.Properties) { $guestPhases[$p.Name] = $p.Value.seconds }
}

$lines = New-Object System.Collections.Generic.List[string]
$guestPlaced = $null
if ($null -ne $guestRecord) { $guestPlaced = $guestRecord.privateList }
$personalData = Get-PersonalDataState ([bool]$SkipSelfTests) ([string]$privateList) $guestPlaced @($privateListLookedIn)
$headline = "== $runId - $($hostVerdict.Verdict) (exit $exitCode)"
if ($hostVerdict.Why) { $headline += " - $($hostVerdict.Why)" }
if ($personalData.Headline) { $headline += " - $($personalData.Headline)" }
$lines.Add($headline + ' ==')
$lines.Add("revision   $($sha.Substring(0, 12)) - $subject")
$lines.Add("           $Ref in $repo$(if ($dirtyCount -gt 0) { " ($dirtyCount uncommitted change(s) NOT included)" })")
$lines.Add("vm         $BuildVmName, of the build pool's $(@($BuildVmPool).Count): $($BuildVmPool -join ', ')")
if ($infraError -and $guestVerdict -and $hostVerdict.Verdict -ne 'INFRA') {
    $lines.Add("host note  something failed on this side - $infraError - and the guest's results came back anyway; they decided the verdict")
}
if ($SkipSuite) { $lines.Add('suite      skipped') }
elseif ($null -ne $trx) {
    $c = $trx.Counts
    $lines.Add(("suite      {0} total: {1} passed, {2} failed, {3} skipped{4} - filter {5}" -f $c['total'], $c['passed'], $c['failed'], $c['skipped'], $(if ($c['other'] -gt 0) { ", $($c['other']) other" } else { '' }), $filterShown))
}
else { $lines.Add("suite      no TRX file - guest verdict '$guestVerdict'") }
foreach ($l in @(Format-ZonePassLines -Requested $zonePassRequested -ZonePass $zonePass -ZoneTrx $zoneTrx)) { $lines.Add($l) }
if ($SkipSelfTests) { $lines.Add('self-tests skipped') }
elseif ($null -eq $guestRecord) { $lines.Add('self-tests not known - no run.json came back') }
else { $lines.Add(("self-tests {0} run: {1} passed, {2} failed; {3} skipped by reason" -f $selfTests.Count, ($selfTests.Count - $selfTestsFailed), $selfTestsFailed, $selfTestSkips.Count)) }
if ($personalData.Line) { $lines.Add($personalData.Line) }
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
if ($null -ne $trx) { Add-FailedTests -Lines $lines -Title 'FAILED TESTS' -Failed @($trx.Failed) -TrxFile 'vm\trx\suite.trx' }
if ($null -ne $zoneTrx) { Add-FailedTests -Lines $lines -Title "FAILED IN THE ZONE PASS, IN $($ZonePassTimeZone.ToUpperInvariant())" -Failed @($zoneTrx.Failed) -TrxFile 'vm\trx\zone.trx' }
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
$lines.Add('           summary.txt and summary.json here; vm\run.log, vm\trx\suite.trx, vm\trx\zone.trx and every build, test and self-test log under vm\')

Set-Content -LiteralPath (Join-Path $runDir 'summary.txt') -Value $lines -Encoding UTF8
$summaryFilter = $filterShown
if ($SkipSuite) { $summaryFilter = '' }
$summary = New-RunSummary -RunId $runId -Verdict $hostVerdict -Revision ([ordered]@{ ref = $Ref; sha = $sha; subject = $subject; repo = $repo; uncommittedLeftOut = $dirtyCount }) -Filter $summaryFilter -Trx $trx -SelfTests $selfTests -SelfTestSkips $selfTestSkips -GuestVerdict $guestVerdict -HostError $infraError -BuildErrors $buildErrors -PackagesStaged $packagesStaged -Timings $timings -GuestPhases $guestPhases -Results $runDir -Vm $BuildVmName -ZonePassRequested $zonePassRequested -ZonePass $zonePass -ZoneTrx $zoneTrx -PersonalData $personalData.State
($summary | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $runDir 'summary.json') -Encoding UTF8
Write-History ("end    $runId  $($hostVerdict.Verdict)  $(Format-Seconds $timings['total'])  $BuildVmName")

Write-Host ''
foreach ($l in $lines) { Write-Host $l }
exit $exitCode
