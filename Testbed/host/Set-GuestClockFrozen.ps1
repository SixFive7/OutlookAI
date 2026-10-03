#Requires -Version 5.1
<#
.SYNOPSIS
    The frozen guest clock of the two Outlook test guests (Q130 (a)): -Verify checks it before a live
    run's suite starts; -Execute makes a guest's frozen checkpoint from a running one.

.DESCRIPTION
    RUN ON THE HOST. Windows PowerShell 5.1 or PowerShell 7. Everything inside the guest runs over
    PowerShell Direct, which needs no network.

    WHY A FROZEN CLOCK. Decided by the maintainer 2026-10-03 (Q130 (a), Docs/overnight-review-2026-10-03.md):
    the live tier's test data has an age, and three checks depend on it - the frontier test needs the
    hub's newest item younger than the guest's UTC offset less five minutes (T2/LiveHubPopulationFreshness),
    the unindexed guest's searches reach only the last 7 days (MailService.EmptyIndexSweepWindow), and
    guest one's tier refuses once a declared corpus window selects none of Corpus A's items
    (T2/LiveCorpusFreshness, from 2026-11-01 23:59:16Z on a real clock). So on OutlookAI-Indexed and
    OutlookAI-Unindexed, Hyper-V time synchronisation is OFF and every run restores a checkpoint taken
    with it off, just after the data's anchor: every run then starts at the same instant, and the data is
    exactly as old on every run as it was on the first. Testbed/testbed.json, frozenClocks, records each
    guest's checkpoint and instant; Docs/live-tier-on-the-vm.md section 4.4 is the record of how they
    were made and what was measured.

    WHAT HOLDS IT, measured on OutlookAI-Unindexed 2026-10-03 (runbook section 4.1f) and again here
    (section 4.4): with time sync off nothing else sets the guest's clock (the Windows Time service is
    not running); a saved guest's clock stops; a checkpoint taken with time sync off restores to the same
    instant every time, and the setting travels with the checkpoint. WHAT BREAKS IT: a cold boot or an OS
    restart after the restore - the guest comes back at the host's time plus the offset it last WROTE to
    its clock, not at the frozen instant. So nothing in a run may restart the guest:
    Testbed/host/Restart-Guest.ps1 refuses a guest whose time sync is off unless it is told -Refreeze,
    which carries the guest's own time across the restart, and a run never passes that.

    NEVER the build VM, OutlookAI-Build: its runner requires the host's clock within 2 s
    (Invoke-TestsOnBuildVm.ps1). NEVER OutlookAI-Exchange: Microsoft 365 sign-in needs real time. Both
    are refused by name.

    -Verify (THE DEFAULT, READ-ONLY): the guard a live run calls after its restore and staging, right
    before the suite. It reads this guest's record in testbed.json and checks, changing nothing:
      * time synchronisation is OFF - on means the guest was not restored from its frozen checkpoint
        (the setting travels with the checkpoint), or somebody turned it back on;
      * the guest was restored from the recorded frozen checkpoint (Get-VM's ParentCheckpointName);
      * the guest's clock, read over PowerShell Direct, is no more than -ToleranceSeconds BEFORE the
        frozen instant and no more than the record's suiteStartWithinMinutes after it. Before: someone
        set the clock by hand, or the checkpoint is not the recorded one. Long after: the guest
        restarted or cold-booted after the restore, time sync was on for a while, or the run spent too
        long before its suite - every date-window check is proven (T1/FrozenGuestClockTests) only up to
        checksHoldForMinutes after the frozen instant, and the suite needs the difference.
    It also prints the guest's time zone and UTC offset - the frontier margin is that offset less five
    minutes, and the frozen date lies in summer time for good.
      EXIT CODES: 0 FROZEN. 1 NOT-FROZEN or CLOCK-MOVED - do not run the suite; restore the frozen
      checkpoint and stage again. 3 CANNOT-TELL - the VM is not running, or the guest did not answer.
      4 REFUSED - no record for this VM, or a VM that is never frozen.

    -Execute -FromCheckpoint <running checkpoint> -NewCheckpoint <name>: makes a frozen checkpoint.
    Without -Execute it prints what it would do and changes nothing. In order:
      1. refuses unless the VM is one of the two Outlook guests, -FromCheckpoint exists and was taken
         RUNNING (only a saved memory image carries a clock; a cold boot comes up at host time), and no
         checkpoint is named -NewCheckpoint - it never deletes or overwrites a checkpoint;
      2. saves the VM if it is running (a checkpoint applied to a RUNNING VM resumes it at once, before
         time sync could be turned off - runbook section 4.3), restores -FromCheckpoint, which leaves
         it Saved, and turns time synchronisation OFF while it is saved;
      3. starts it - it resumes at -FromCheckpoint's own instant, with nothing to set it to the host's
         time - and reads the guest's clock three times, 10 s apart: the offset from the host must hold
         to 2 s (nothing re-syncs it), and time sync must still read off;
      4. takes -NewCheckpoint (a standard checkpoint of the running guest) and reads the clock right
         before and right after it: the frozen instant is the midpoint;
      5. proves it: saves, restores -NewCheckpoint, checks time sync is off in it, starts it, and reads
         the clock again - within 10 s of the frozen instant - then saves, restores it once more and
         leaves the VM SAVED on it, holding no RAM, its next start at the frozen instant;
      6. prints the frozenClocks record for testbed.json, the instant as text AND as a Unix second
         count: the count is what scripts read, because ConvertFrom-Json turns an ISO-8601 string into
         a DateTime whose zone it then loses (Testbed/README.md section 5b).
    It takes no lease and releases none: the lease belongs to the work around it, and it warns when
    there is none. It never starts or quits Outlook, never touches a mail item, and never restarts or
    cold-boots the guest.
      EXIT CODES: 0 made and proven. 1 a check failed part-way - the VM is left as the message says.
      4 refused before anything changed.

.PARAMETER VMName
    MANDATORY. OutlookAI-Indexed or OutlookAI-Unindexed - no default picks a guest (Testbed/README.md
    section 4a).

.PARAMETER Verify
    The default. Read-only.

.PARAMETER Execute
    With -FromCheckpoint and -NewCheckpoint: make the frozen checkpoint.

.PARAMETER FromCheckpoint
    The running checkpoint whose instant becomes the frozen one - it should be taken just after the
    guest's data anchor (the hub rebuild).

.PARAMETER NewCheckpoint
    The name of the frozen checkpoint. Must not exist yet.

.PARAMETER ToleranceSeconds
    How far BEFORE the frozen instant -Verify accepts the guest's clock. A restore comes up 1.7-1.8 s
    before the instant a checkpoint was taken (runbook section 4.1f). Default 120.

.PARAMETER RepoRoot
    The repository root for the credential loader (Get-GuestCredential.ps1). testbed.json is always
    read from the repository this script is in.

.PARAMETER LogPath
    Optional transcript file on the host.

.PARAMETER SelfTest
    Pure: the record reader, the verdict table, the refusals and this file's own syntax tree, against
    synthetic inputs and the committed testbed.json. No Hyper-V, no guest. Runs on the build VM with
    every other -SelfTest.

.EXAMPLE
    pwsh -File Testbed/host/Set-GuestClockFrozen.ps1 -VMName OutlookAI-Indexed -Verify
    pwsh -File Testbed/host/Set-GuestClockFrozen.ps1 -VMName OutlookAI-Unindexed -FromCheckpoint CP-13B-LIVE-GREEN -NewCheckpoint CP-14B-FROZEN-CLOCK
    pwsh -File Testbed/host/Set-GuestClockFrozen.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $VMName,
    [switch] $Verify,
    [switch] $Execute,
    [string] $FromCheckpoint,
    [string] $NewCheckpoint,
    [int]    $ToleranceSeconds = 120,
    [string] $RepoRoot,
    [string] $LogPath,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# Resolved here rather than as a parameter default: Windows PowerShell 5.1 run with -File leaves
# $PSScriptRoot empty while it evaluates parameter defaults (Tools/Checks/check-powershell-51.ps1, rule 1).
$OwnRepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $RepoRoot) { $RepoRoot = $OwnRepoRoot }

# Hyper-V's Time Synchronization integration service, by its component id: the display name is
# localised, the id is not. Read off this host 2026-10-03 (Get-VMIntegrationService ... | Select Id).
$script:TimeSyncComponentId = '2497F4DE-E9FA-4204-80E4-4B75C46419C0'

# The only machines this script freezes, and the two it refuses with the reason.
$script:OutlookGuests = @('OutlookAI-Indexed', 'OutlookAI-Unindexed')
$script:NeverFrozen = @{
    'OutlookAI-Build'    = 'the build VM is never frozen: its runner requires the host''s clock within 2 s (Testbed/host/Invoke-TestsOnBuildVm.ps1)'
    'OutlookAI-Exchange' = 'the Exchange VM is never frozen: Microsoft 365 sign-in needs real time'
}

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    if ($LogPath) {
        try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
    }
}

# ---------------------------------------------------------------------------------------------
# Pure parts - -SelfTest pins them
# ---------------------------------------------------------------------------------------------

# Why a VM may not be frozen, or $null when it may.
function Get-NeverFrozenReason([string] $Name) {
    if ($script:NeverFrozen.ContainsKey($Name)) { return $script:NeverFrozen[$Name] }
    if ($script:OutlookGuests -notcontains $Name) { return "'$Name' is not one of the two Outlook test guests ($($script:OutlookGuests -join ', '))" }
    return $null
}

function ConvertFrom-UnixSeconds([long] $Seconds) {
    return [DateTime]::SpecifyKind([DateTime]::new(1970, 1, 1, 0, 0, 0).AddSeconds($Seconds), [DateTimeKind]::Utc)
}

function Format-Utc([DateTime] $Utc) {
    return [DateTime]::SpecifyKind($Utc, [DateTimeKind]::Utc).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-Span([TimeSpan] $Span) {
    $sign = ''
    if ($Span -lt [TimeSpan]::Zero) { $sign = '-' }
    $d = $Span.Duration()
    if ($d.TotalDays -ge 2) { return ('{0}{1:N1} days' -f $sign, $d.TotalDays) }
    if ($d.TotalMinutes -ge 2) { return ('{0}{1:N1} min' -f $sign, $d.TotalMinutes) }
    return ('{0}{1:N1} s' -f $sign, $d.TotalSeconds)
}

# A guest's frozenClocks record out of a parsed testbed.json, as the fields the checks need - or a
# throw naming what is missing. The instant is read from frozenUnix only: a number survives
# ConvertFrom-Json in both shells, an ISO-8601 string does not (Testbed/README.md section 5b).
function Get-FrozenClockRecord($Testbed, [string] $Name) {
    $section = $null
    if ($Testbed.PSObject.Properties.Name -contains 'frozenClocks') { $section = $Testbed.frozenClocks }
    if ($null -eq $section) { throw 'Testbed/testbed.json has no frozenClocks section.' }
    if ($section.PSObject.Properties.Name -notcontains $Name) {
        throw "Testbed/testbed.json's frozenClocks section has no record for '$Name' - its clock is not frozen, so there is nothing to verify."
    }
    $r = $section.$Name
    foreach ($field in @('checkpoint', 'parentCheckpoint', 'frozenUnix', 'suiteStartWithinMinutes', 'checksHoldForMinutes')) {
        if ($r.PSObject.Properties.Name -notcontains $field -or $null -eq $r.$field -or "$($r.$field)" -eq '') {
            throw "frozenClocks.$Name has no '$field'."
        }
    }
    $start = [int]$r.suiteStartWithinMinutes
    $hold = [int]$r.checksHoldForMinutes
    if ($start -le 0 -or $hold -le $start) {
        throw "frozenClocks.${Name}: suiteStartWithinMinutes ($start) must be above zero and below checksHoldForMinutes ($hold) - the difference is what the suite itself may take."
    }
    return [pscustomobject]@{
        Name                    = $Name
        Checkpoint              = [string]$r.checkpoint
        ParentCheckpoint        = [string]$r.parentCheckpoint
        FrozenUtc               = (ConvertFrom-UnixSeconds ([long]$r.frozenUnix))
        SuiteStartWithinMinutes = $start
        ChecksHoldForMinutes    = $hold
    }
}

# The guard's verdict. Pure: what the host read in, a verdict and its reasons out.
function Get-FrozenClockVerdict {
    param(
        [Parameter(Mandatory = $true)] $Record,
        [Parameter(Mandatory = $true)] [bool] $TimeSyncEnabled,
        [string] $ParentCheckpointName,
        [Parameter(Mandatory = $true)] [DateTime] $GuestUtc,
        [int] $Tolerance = 120
    )
    $reasons = @()
    $verdict = 'FROZEN'
    $elapsed = [DateTime]::SpecifyKind($GuestUtc, [DateTimeKind]::Utc) - $Record.FrozenUtc
    if ($TimeSyncEnabled) {
        $verdict = 'NOT-FROZEN'
        $reasons += "time synchronisation is ON: this guest was not restored from its frozen checkpoint '$($Record.Checkpoint)' (the setting travels with the checkpoint), or somebody turned it back on. Never re-enable it on an Outlook guest (AGENTS.md)."
    }
    if ($ParentCheckpointName -ne $Record.Checkpoint) {
        $verdict = 'NOT-FROZEN'
        $reasons += "the guest was restored from '$ParentCheckpointName', not from its frozen checkpoint '$($Record.Checkpoint)'. A new frozen checkpoint is used only once testbed.json records it."
    }
    if ($elapsed.TotalSeconds -lt -$Tolerance) {
        if ($verdict -eq 'FROZEN') { $verdict = 'CLOCK-MOVED' }
        $reasons += "the guest's clock is $(Format-Span $elapsed.Duration()) BEFORE the frozen instant $(Format-Utc $Record.FrozenUtc): somebody set it by hand, or this is not the recorded checkpoint."
    }
    elseif ($elapsed.TotalMinutes -gt $Record.SuiteStartWithinMinutes) {
        if ($verdict -eq 'FROZEN') { $verdict = 'CLOCK-MOVED' }
        if ($elapsed.TotalDays -ge 1) {
            $reasons += "the guest's clock is $(Format-Span $elapsed) past the frozen instant $(Format-Utc $Record.FrozenUtc) - not a slow run: the guest restarted or cold-booted after the restore (it comes back at the host's time plus the offset it last wrote), or time sync was on. Restore '$($Record.Checkpoint)' and stage again, and never restart a frozen guest in a run."
        }
        else {
            $reasons += "the guest's clock is $(Format-Span $elapsed) past the frozen instant $(Format-Utc $Record.FrozenUtc), more than the $($Record.SuiteStartWithinMinutes) min a run may spend before its suite: the date checks are proven only to $($Record.ChecksHoldForMinutes) min after it, and the suite needs the rest. Restore '$($Record.Checkpoint)' and stage again, faster."
        }
    }
    return [pscustomobject]@{ Verdict = $verdict; Reasons = $reasons; Elapsed = $elapsed }
}

# The time sync integration service of a VM, found by component id.
function Get-TimeSyncService([string] $Name) {
    $all = @(Get-VMIntegrationService -VMName $Name -ErrorAction Stop)
    $byId = @($all | Where-Object { ([string]$_.Id).ToUpperInvariant().EndsWith($script:TimeSyncComponentId) })
    if ($byId.Count -eq 1) { return $byId[0] }
    $byName = @($all | Where-Object { $_.Name -eq 'Time Synchronization' })
    if ($byName.Count -eq 1) { return $byName[0] }
    throw "'$Name' has no Time Synchronization integration service this script can find (component id $($script:TimeSyncComponentId))."
}

# ---------------------------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------------------------
function Invoke-SelfTest {
    function Check([string] $name, $got, $want) {
        if ("$got" -ceq "$want") { $script:stPass++; Write-Host "  PASS  $name" }
        else { $script:stFail++; Write-Host "  FAIL  $name - got '$got', want '$want'" }
    }
    $script:stPass = 0
    $script:stFail = 0

    Write-Host '== which machines may be frozen =='
    Check 'the indexed guest may' (Get-NeverFrozenReason 'OutlookAI-Indexed') ''
    Check 'the unindexed guest may' (Get-NeverFrozenReason 'OutlookAI-Unindexed') ''
    Check 'the build VM never: its runner needs the host clock' ([string](Get-NeverFrozenReason 'OutlookAI-Build')).Contains('within 2 s') $true
    Check 'the Exchange VM never: Microsoft 365 sign-in needs real time' ([string](Get-NeverFrozenReason 'OutlookAI-Exchange')).Contains('real time') $true
    Check 'the retired guest is not an Outlook test guest' ([string](Get-NeverFrozenReason 'OutlookAI-TestVM')).Contains('not one of the two') $true

    Write-Host '== instants =='
    Check 'Unix seconds read as UTC' (Format-Utc (ConvertFrom-UnixSeconds 1791049320)) '2026-10-03T17:42:00Z'
    Check 'and the kind is Utc' ((ConvertFrom-UnixSeconds 0).Kind) 'Utc'

    Write-Host '== the record =='
    $tb = [pscustomobject]@{ frozenClocks = [pscustomobject]@{
            'OutlookAI-Indexed' = [pscustomobject]@{ checkpoint = 'CP-X'; parentCheckpoint = 'CP-W'; frozenUtc = '2026-10-03T17:42:00Z'; frozenUnix = 1791049320; suiteStartWithinMinutes = 45; checksHoldForMinutes = 90 }
        }
    }
    $rec = Get-FrozenClockRecord $tb 'OutlookAI-Indexed'
    Check 'the record is read' ('{0}|{1}|{2}|{3}' -f $rec.Checkpoint, (Format-Utc $rec.FrozenUtc), $rec.SuiteStartWithinMinutes, $rec.ChecksHoldForMinutes) 'CP-X|2026-10-03T17:42:00Z|45|90'
    $threw = ''
    try { [void](Get-FrozenClockRecord $tb 'OutlookAI-Unindexed') } catch { $threw = $_.Exception.Message }
    Check 'a guest with no record is refused, naming it' $threw.Contains("no record for 'OutlookAI-Unindexed'") $true
    $threw = ''
    try { [void](Get-FrozenClockRecord ([pscustomobject]@{ corpus = 1 }) 'OutlookAI-Indexed') } catch { $threw = $_.Exception.Message }
    Check 'a testbed.json with no frozenClocks is refused' $threw.Contains('no frozenClocks section') $true
    $bad = [pscustomobject]@{ frozenClocks = [pscustomobject]@{ 'OutlookAI-Indexed' = [pscustomobject]@{ checkpoint = 'CP-X'; parentCheckpoint = 'CP-W'; frozenUnix = 1791049320; suiteStartWithinMinutes = 90; checksHoldForMinutes = 90 } } }
    $threw = ''
    try { [void](Get-FrozenClockRecord $bad 'OutlookAI-Indexed') } catch { $threw = $_.Exception.Message }
    Check 'a suite allowance of nothing is refused' $threw.Contains('below checksHoldForMinutes') $true
    $noUnix = [pscustomobject]@{ frozenClocks = [pscustomobject]@{ 'OutlookAI-Indexed' = [pscustomobject]@{ checkpoint = 'CP-X'; parentCheckpoint = 'CP-W'; frozenUtc = '2026-10-03T17:42:00Z'; suiteStartWithinMinutes = 45; checksHoldForMinutes = 90 } } }
    $threw = ''
    try { [void](Get-FrozenClockRecord $noUnix 'OutlookAI-Indexed') } catch { $threw = $_.Exception.Message }
    Check 'the text instant alone is not read - the Unix count is required' $threw.Contains("no 'frozenUnix'") $true

    Write-Host '== the verdict =='
    $f = $rec.FrozenUtc
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-X' -GuestUtc $f.AddSeconds(-2)
    Check 'restored, 2 s before the instant (as measured): FROZEN' $v.Verdict 'FROZEN'
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-X' -GuestUtc $f.AddMinutes(20)
    Check '20 min in, staging done: FROZEN' $v.Verdict 'FROZEN'
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-X' -GuestUtc $f.AddMinutes(45)
    Check 'at exactly the allowance: FROZEN' $v.Verdict 'FROZEN'
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-X' -GuestUtc $f.AddMinutes(46)
    Check 'one minute past it: CLOCK-MOVED' $v.Verdict 'CLOCK-MOVED'
    Check 'and it says the run was too slow' ([string]$v.Reasons[0]).Contains('a run may spend before its suite') $true
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-X' -GuestUtc $f.AddDays(40)
    Check 'forty days past it - a restart after the restore: CLOCK-MOVED' $v.Verdict 'CLOCK-MOVED'
    Check 'and it names the restart, not a slow run' ([string]$v.Reasons[0]).Contains('restarted or cold-booted') $true
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-X' -GuestUtc $f.AddMinutes(-5)
    Check 'five minutes before the instant: CLOCK-MOVED' $v.Verdict 'CLOCK-MOVED'
    Check 'and it says BEFORE' ([string]$v.Reasons[0]).Contains('BEFORE the frozen instant') $true
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $true -ParentCheckpointName 'CP-X' -GuestUtc $f
    Check 'time sync on: NOT-FROZEN, whatever the clock reads' $v.Verdict 'NOT-FROZEN'
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $false -ParentCheckpointName 'CP-W' -GuestUtc $f
    Check 'restored from another checkpoint: NOT-FROZEN' $v.Verdict 'NOT-FROZEN'
    Check 'and it names both' (([string]$v.Reasons[0]).Contains("'CP-W'") -and ([string]$v.Reasons[0]).Contains("'CP-X'")) $true
    $v = Get-FrozenClockVerdict -Record $rec -TimeSyncEnabled $true -ParentCheckpointName 'CP-W' -GuestUtc $f.AddDays(40)
    Check 'every fault is reported, not only the first' @($v.Reasons).Count 3

    Write-Host '== the committed records (Testbed/testbed.json) =='
    $tbPath = Join-Path $OwnRepoRoot 'Testbed\testbed.json'
    if (Test-Path -LiteralPath $tbPath) {
        $committed = Get-Content -LiteralPath $tbPath -Raw | ConvertFrom-Json
        foreach ($g in $script:OutlookGuests) {
            $r = Get-FrozenClockRecord $committed $g
            Check "$g has a frozen checkpoint that is not its parent" (($r.Checkpoint -ne '') -and ($r.Checkpoint -ne $r.ParentCheckpoint)) $true
            Check "$g's instant is in 2026, before summer time ends on 2026-10-25" (($r.FrozenUtc.Year -eq 2026) -and ($r.FrozenUtc -lt [DateTime]::SpecifyKind([DateTime]::new(2026, 10, 25, 1, 0, 0), [DateTimeKind]::Utc))) $true
        }
        foreach ($never in $script:NeverFrozen.Keys) {
            Check "$never has no frozen record" ($committed.frozenClocks.PSObject.Properties.Name -contains $never) $false
        }
    }
    else { Write-Host "  SKIP  the committed records - $tbPath is not here" }

    Write-Host '== this file =='
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $names = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
    foreach ($forbidden in @('Stop-VM', 'Restart-VM', 'Remove-VMSnapshot', 'Remove-VMCheckpoint', 'Enable-VMIntegrationService', 'Stop-Process', 'Set-Date')) {
        Check "it never calls $forbidden" ($names -contains $forbidden) $false
    }
    Check 'it turns time sync off and never on' ($names -contains 'Disable-VMIntegrationService') $true

    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:stPass, $script:stFail)
    if ($script:stFail -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

# ---------------------------------------------------------------------------------------------
# Arguments
# ---------------------------------------------------------------------------------------------
if (-not $VMName) { Say 'REFUSING: -VMName is mandatory (Testbed/README.md section 4a): name the guest you mean.'; exit 4 }
if ($VMName -notmatch '^[A-Za-z0-9._-]{1,64}$') { Say "REFUSING: VM name '$VMName' is not a plain name."; exit 4 }
$never = Get-NeverFrozenReason $VMName
if ($never) { Say "REFUSING: $never."; exit 4 }
$making = ($FromCheckpoint -or $NewCheckpoint)
if ($making -and $Verify) { Say 'REFUSING: -Verify reads; -FromCheckpoint and -NewCheckpoint make a checkpoint. Pick one.'; exit 4 }
if ($making -and (-not $FromCheckpoint -or -not $NewCheckpoint)) { Say 'REFUSING: making a frozen checkpoint needs both -FromCheckpoint and -NewCheckpoint.'; exit 4 }
if ($Execute -and -not $making) { Say 'REFUSING: -Execute makes a frozen checkpoint and needs -FromCheckpoint and -NewCheckpoint; -Verify needs neither.'; exit 4 }
if ($making -and ($NewCheckpoint -notmatch '^[A-Za-z0-9._-]{1,80}$')) { Say "REFUSING: '$NewCheckpoint' is not a plain checkpoint name."; exit 4 }

# ---------------------------------------------------------------------------------------------
# Plumbing
# ---------------------------------------------------------------------------------------------
function Get-VmState { return [string](Get-VM -Name $VMName -ErrorAction Stop).State }

function Get-Heartbeat {
    $hb = Get-VMIntegrationService -VMName $VMName -Name 'Heartbeat' -ErrorAction SilentlyContinue
    if (-not $hb) { return 'absent' }
    return [string]$hb.PrimaryStatusDescription
}

# The guest's clock against the host's, the host read on both sides of the call.
function Read-GuestClock {
    $h0 = [DateTime]::UtcNow
    $g = Invoke-Command -VMName $VMName -Credential $script:Cred -ErrorAction Stop -ScriptBlock {
        $tz = Get-TimeZone
        [pscustomobject]@{ Ticks = [DateTime]::UtcNow.Ticks; Zone = $tz.Id; OffsetMinutes = [int]$tz.GetUtcOffset([DateTime]::UtcNow).TotalMinutes }
    }
    $h1 = [DateTime]::UtcNow
    $host_ = $h0.AddTicks([long](($h1 - $h0).Ticks / 2))
    $guest = [DateTime]::new([long]$g.Ticks, [DateTimeKind]::Utc)
    return [pscustomobject]@{ GuestUtc = $guest; HostUtc = $host_; Skew = ($guest - $host_); Zone = [string]$g.Zone; OffsetMinutes = [int]$g.OffsetMinutes }
}

function Wait-GuestAnswers([int] $Minutes) {
    $deadline = (Get-Date).AddMinutes($Minutes)
    while ((Get-Date) -lt $deadline) {
        if ((Get-Heartbeat) -eq 'OK') {
            try { return (Read-GuestClock) } catch { }
        }
        Start-Sleep -Seconds 2
    }
    throw "the guest did not answer over PowerShell Direct within $Minutes min"
}

function Wait-VmState([string] $Want, [int] $Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        if ((Get-VmState) -eq $Want) { return $true }
        Start-Sleep -Seconds 1
    }
    return ((Get-VmState) -eq $Want)
}

$vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
if (-not $vm) { Say "REFUSING: no VM named '$VMName' on this host."; exit 4 }
try {
    . (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
    if (-not (Get-TestbedLease -VMName $VMName)) { Say "WARNING: no live lease on $VMName. The idle-saver may save it under you (Testbed/README.md section 5b)." }
}
catch { Say "WARNING: could not read the lease directory: $($_.Exception.Message)" }

# ---------------------------------------------------------------------------------------------
# -Verify: the pre-suite guard
# ---------------------------------------------------------------------------------------------
if (-not $making) {
    Say "== Set-GuestClockFrozen -Verify: $VMName =="
    try {
        $testbed = Get-Content -LiteralPath (Join-Path $OwnRepoRoot 'Testbed\testbed.json') -Raw | ConvertFrom-Json
        $record = Get-FrozenClockRecord $testbed $VMName
    }
    catch { Say "REFUSED: $($_.Exception.Message)"; exit 4 }
    Say "  record     frozen at $(Format-Utc $record.FrozenUtc) in '$($record.Checkpoint)' (parent '$($record.ParentCheckpoint)'); the suite starts within $($record.SuiteStartWithinMinutes) min, the date checks hold to $($record.ChecksHoldForMinutes) min"
    $state = Get-VmState
    if ($state -ne 'Running') { Say "CANNOT-TELL: '$VMName' is $state. The guard reads a running guest - restore '$($record.Checkpoint)' and start it."; exit 3 }
    $timeSync = Get-TimeSyncService $VMName
    $parent = [string](Get-VM -Name $VMName).ParentCheckpointName
    try {
        $script:Cred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
        $clock = Read-GuestClock
    }
    catch { Say "CANNOT-TELL: the guest did not answer over PowerShell Direct: $($_.Exception.Message)"; exit 3 }
    Say "  time sync  $(if ($timeSync.Enabled) { 'ON' } else { 'off' })"
    Say "  restored   from '$parent'"
    Say "  guest      $(Format-Utc $clock.GuestUtc) UTC ($(Format-Span ($clock.GuestUtc - $record.FrozenUtc)) after the frozen instant; $(Format-Span $clock.Skew) from the host's clock)"
    Say "  zone       $($clock.Zone), UTC offset $($clock.OffsetMinutes) min - the frontier margin is that less 5 min"
    $v = Get-FrozenClockVerdict -Record $record -TimeSyncEnabled ([bool]$timeSync.Enabled) -ParentCheckpointName $parent -GuestUtc $clock.GuestUtc -Tolerance $ToleranceSeconds
    foreach ($r in $v.Reasons) { Say "  - $r" }
    Say "== $($v.Verdict) =="
    if ($v.Verdict -eq 'FROZEN') { exit 0 }
    exit 1
}

# ---------------------------------------------------------------------------------------------
# -FromCheckpoint -NewCheckpoint: make a frozen checkpoint
# ---------------------------------------------------------------------------------------------
Say "== Set-GuestClockFrozen: $VMName, '$FromCheckpoint' -> '$NewCheckpoint' ($(if ($Execute) { 'EXECUTE' } else { 'DRY RUN' })) =="
$from = Get-VMSnapshot -VMName $VMName -Name $FromCheckpoint -ErrorAction SilentlyContinue
if (-not $from) { Say "REFUSING: '$VMName' has no checkpoint named '$FromCheckpoint'."; exit 4 }
if (@($from).Count -ne 1) { Say "REFUSING: '$VMName' has $(@($from).Count) checkpoints named '$FromCheckpoint'."; exit 4 }
if (Get-VMSnapshot -VMName $VMName -Name $NewCheckpoint -ErrorAction SilentlyContinue) { Say "REFUSING: '$VMName' already has a checkpoint named '$NewCheckpoint'. This script never overwrites or deletes one."; exit 4 }
# A checkpoint of a running VM carries its memory - and with it the clock. Hyper-V says so on the
# checkpoint itself: its State is Saved (read on this host 2026-10-03: CP-18C-ALL-KINDS, Standard,
# Saved); a checkpoint of a VM that was off reads Off.
if ([string]$from.State -ne 'Saved') { Say "REFUSING: '$FromCheckpoint' holds no running guest (its state is $($from.State), not Saved). Only a running checkpoint carries a clock; a cold boot comes up at the host's time."; exit 4 }
Say "  from       '$FromCheckpoint', taken $($from.CreationTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')) - the instant its guest resumes at"
if (-not $Execute) {
    Say ''
    Say 'DRY RUN. Nothing was changed. With -Execute this would:'
    Say '  1. save the VM if it runs, restore the checkpoint (it is then Saved), turn time sync OFF'
    Say '  2. start it, and read its clock three times 10 s apart: the offset from the host must hold'
    Say "  3. take '$NewCheckpoint' and read the clock on both sides of it - the frozen instant"
    Say '  4. save, restore it, check time sync is off in it, start it, read the clock (within 10 s)'
    Say '  5. save, restore it again and leave the VM SAVED on it; print the testbed.json record'
    exit 0
}

$script:Cred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
try {
    # 1. Saved, restored, time sync off while saved.
    if ((Get-VmState) -eq 'Running') {
        Say '== saving the VM before the restore (a checkpoint applied to a running VM resumes it at once) =='
        Save-VM -Name $VMName
        if (-not (Wait-VmState 'Saved' 300)) { throw "the VM did not reach Saved (it is $(Get-VmState))" }
    }
    Say "== restoring '$FromCheckpoint' =="
    Restore-VMSnapshot -VMName $VMName -Name $FromCheckpoint -Confirm:$false
    $afterRestore = Get-VmState
    Say "  the VM is $afterRestore"
    if ($afterRestore -ne 'Saved') { throw "after the restore the VM is $afterRestore, not Saved - it would not resume at the checkpoint's instant" }
    Say '== turning time synchronisation OFF while it is saved =='
    $ts = Get-TimeSyncService $VMName
    Disable-VMIntegrationService -VMName $VMName -Name $ts.Name
    $ts = Get-TimeSyncService $VMName
    if ($ts.Enabled) { throw 'time synchronisation still reads ON after Disable-VMIntegrationService' }
    Say '  time sync off'

    # 2. Started; the clock must stay where it was.
    Say '== starting it: it resumes at the checkpoint''s instant =='
    Start-VM -Name $VMName
    $c1 = Wait-GuestAnswers 5
    Say "  guest $(Format-Utc $c1.GuestUtc) UTC, $(Format-Span $c1.Skew) from the host; the checkpoint was taken $(Format-Span ($c1.GuestUtc - $from.CreationTime.ToUniversalTime())) from that"
    Start-Sleep -Seconds 10
    $c2 = Read-GuestClock
    Start-Sleep -Seconds 10
    $c3 = Read-GuestClock
    $drift = [Math]::Max([Math]::Abs(($c2.Skew - $c1.Skew).TotalSeconds), [Math]::Abs(($c3.Skew - $c1.Skew).TotalSeconds))
    Say ("  offset from the host {0} / {1} / {2} over 20 s - it moved {3:N1} s at most" -f (Format-Span $c1.Skew), (Format-Span $c2.Skew), (Format-Span $c3.Skew), $drift)
    if ($drift -gt 2) { throw "the guest's offset from the host moved $drift s in 20 s - something is setting its clock" }
    if ((Get-TimeSyncService $VMName).Enabled) { throw 'time synchronisation reads ON again after the start' }
    Say "  zone $($c3.Zone), UTC offset $($c3.OffsetMinutes) min"

    # 3. The frozen checkpoint, bracketed by two clock readings.
    Say "== taking '$NewCheckpoint' =="
    $before = Read-GuestClock
    Checkpoint-VM -Name $VMName -SnapshotName $NewCheckpoint
    $after = Read-GuestClock
    $frozen = $before.GuestUtc.AddTicks([long](($after.GuestUtc - $before.GuestUtc).Ticks / 2))
    $frozen = [DateTime]::new($frozen.Ticks - ($frozen.Ticks % [TimeSpan]::TicksPerSecond), [DateTimeKind]::Utc)
    Say "  the guest read $(Format-Utc $before.GuestUtc) before it and $(Format-Utc $after.GuestUtc) after it: frozen at $(Format-Utc $frozen)"
    if (-not (Get-VMSnapshot -VMName $VMName -Name $NewCheckpoint -ErrorAction SilentlyContinue)) {
        # Checkpoint-VM has returned before the checkpoint listed, once (New-TestbedVm.ps1) - wait for it.
        Start-Sleep -Seconds 10
        if (-not (Get-VMSnapshot -VMName $VMName -Name $NewCheckpoint -ErrorAction SilentlyContinue)) { throw "Checkpoint-VM returned and no checkpoint '$NewCheckpoint' is listed" }
    }

    # 4. Proved: restored, still off, the same instant.
    Say "== proving it: restore '$NewCheckpoint', start, read the clock =="
    Save-VM -Name $VMName
    if (-not (Wait-VmState 'Saved' 300)) { throw "the VM did not reach Saved (it is $(Get-VmState))" }
    $inCheckpoint = @(Get-VMSnapshot -VMName $VMName -Name $NewCheckpoint | Get-VMIntegrationService | Where-Object { ([string]$_.Id).ToUpperInvariant().EndsWith($script:TimeSyncComponentId) -or $_.Name -eq 'Time Synchronization' })
    if ($inCheckpoint.Count -lt 1 -or $inCheckpoint[0].Enabled) { throw "'$NewCheckpoint' itself does not record time synchronisation off" }
    Restore-VMSnapshot -VMName $VMName -Name $NewCheckpoint -Confirm:$false
    if ((Get-TimeSyncService $VMName).Enabled) { throw "time synchronisation reads ON in '$NewCheckpoint' - the setting did not travel with the checkpoint" }
    Start-VM -Name $VMName
    $p = Wait-GuestAnswers 5
    $off = $p.GuestUtc - $frozen
    Say "  the guest came back at $(Format-Utc $p.GuestUtc) UTC - $(Format-Span $off) from the frozen instant; time sync off"
    if ([Math]::Abs($off.TotalSeconds) -gt 10) { throw "the restored guest is $(Format-Span $off) from the frozen instant, not within 10 s" }

    # 5. Rested: saved on the frozen checkpoint, its next start at the frozen instant.
    Save-VM -Name $VMName
    if (-not (Wait-VmState 'Saved' 300)) { throw "the VM did not reach Saved (it is $(Get-VmState))" }
    Restore-VMSnapshot -VMName $VMName -Name $NewCheckpoint -Confirm:$false
    Say "  restored '$NewCheckpoint' again; the VM is $(Get-VmState)"
}
catch {
    Say "STOPPED: $($_.Exception.Message)"
    Say "  The VM is $(Get-VmState) on '$([string](Get-VM -Name $VMName).ParentCheckpointName)'. No checkpoint was deleted."
    exit 1
}

$unix = [long](($frozen - [DateTime]::SpecifyKind([DateTime]::new(1970, 1, 1), [DateTimeKind]::Utc)).TotalSeconds)
Say ''
Say "== DONE: '$NewCheckpoint' is frozen at $(Format-Utc $frozen) ($unix) =="
Say 'Record it in Testbed/testbed.json, frozenClocks:'
Say "  `"$VMName`": { `"checkpoint`": `"$NewCheckpoint`", `"parentCheckpoint`": `"$FromCheckpoint`", `"frozenUtc`": `"$(Format-Utc $frozen)`", `"frozenUnix`": $unix, ... }"
exit 0
