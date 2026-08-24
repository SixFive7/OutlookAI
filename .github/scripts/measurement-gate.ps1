#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-release gate: compares this run's performance and coverage numbers against THIS
    MACHINE'S OWN history, and fails when anything moved - in either direction.

.DESCRIPTION
    Every number this project derives a constant from was measured once, on one machine, under
    conditions nobody wrote down. That has already gone wrong here: `SweepBudgetMs` was set to
    180 s from a measurement taken while `Table.Sort` was silently failing, so the number
    described broken behaviour and nothing recorded that. This script is the mechanism that
    stops it recurring - not by re-deriving the constants, but by refusing to let a number move
    quietly between releases.

    THREE THINGS IT DOES, AND WHY EACH IS THE WAY IT IS.

    1. MOVEMENT IN EITHER DIRECTION IS A SIGNAL. A sweep that got 40% faster is not good news
       by default; the most common cause in this codebase is that something stopped being done
       (a sort that started failing, a folder set that shrank, a filter that stopped matching).
       So the default verdict class is symmetric and a FASTER number fails exactly as a slower
       one does. The report says which direction it moved and does not congratulate.

    2. THE BIAS IS TO FAIL. The maintainer's words: "fail aggressively over leaking slow
       performance degradations". A borderline call fails. A metric that has history and is
       missing from this run fails. A cold start does not silently pass - it exits 2 and has to
       be accepted by hand.

    3. THE NUMBERS NEVER REACH THE REPOSITORY. They are statistics about one person's machine,
       meaningful only against older values from that same machine, and the maintainer does not
       want them published. So the store lives under %LOCALAPPDATA%, this script REFUSES to
       write anywhere inside a git working tree, and it REFUSES to run its comparing modes at
       all under CI (where stdout is a public build log). `.github/scripts/check-measurement-privacy.ps1`
       is the CI-side half that proves nothing measurement-shaped ever got committed.

    PROVENANCE IS PART OF THE RECORD, NOT A NICETY. Every run stores the commit, the branch,
    whether the tree was dirty, the machine, the OS, the SDK, and the CONDITIONS: which profile
    (production / vm), whether the stores were indexed, which corpus, what the store set looked
    like. Comparison is scoped by profile and corpus, because a VM number and a production
    number are not the same measurement, and a change of the `indexed` condition fails outright
    rather than being compared - an indexed sweep and an unindexed sweep run different code.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER StoreRoot
    Baseline store. Defaults to %LOCALAPPDATA%\OutlookAI\Measurements. Refused if it resolves
    inside any git working tree.

.PARAMETER Run
    A measurement-run JSON to ingest. See Docs/measurement-gate.md for the shape, or run
    -Template to print one.

.PARAMETER Collect
    Also collect the metrics that need no mailbox: the pinned-constant invariant counts, the
    budget constants read out of the sources, and (with -TestLog) the suite numbers.

.PARAMETER TestLog
    Output of `dotnet test` (or a .trx) to read the suite numbers from.

.PARAMETER Require
    Present (default) compares whatever the run carries, and fails on any metric that has
    history but is missing now. All demands every catalogued metric - this is the pre-release
    invocation, and the one that stops a partial run passing as a full one.

.PARAMETER Tolerance
    Default symmetric tolerance for `both`-class metrics, as a fraction. Default 0.10.

.EXAMPLE
    pwsh -File .github/scripts/measurement-gate.ps1 -Collect -TestLog .work/test.log -ProfileKind production
.EXAMPLE
    pwsh -File .github/scripts/measurement-gate.ps1 -Run .work/live-run.json -Collect -TestLog .work/test.log -Require All
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string] $StoreRoot,

    # --- ingest -------------------------------------------------------------------------
    [string] $Run,
    [switch] $Collect,
    [string] $TestLog,
    [string] $Label,

    # --- conditions ---------------------------------------------------------------------
    [ValidateSet('production', 'vm', 'unknown')]
    [string] $ProfileKind,
    [ValidateSet('indexed', 'unindexed', 'mixed', 'unknown')]
    [string] $Indexed,
    [string] $CorpusId,
    [string] $CorpusAnchor,
    [string] $StoreSet,
    [string] $Notes,

    # --- policy -------------------------------------------------------------------------
    [double] $Tolerance = 0.10,
    [ValidateRange(1, 99)]
    [int] $BaselineRuns = 5,
    [ValidateSet('Present', 'All')]
    [string] $Require = 'Present',
    [switch] $AcceptNewBaseline,
    [switch] $AllowUnknownMetrics,
    [switch] $StrictConditions,
    [switch] $DryRun,

    # --- other modes --------------------------------------------------------------------
    [switch] $Show,
    [int] $Last = 10,
    [switch] $Annotate,
    [switch] $ResetBaseline,
    [string] $Metric,
    [string] $Reason,
    [switch] $SelfTest,
    [switch] $ListMetrics,
    [switch] $Template,
    [switch] $ReleaseNoteSummary,
    [switch] $AllowCi
)

$ErrorActionPreference = 'Stop'

# EVERY number in this script is formatted and parsed in the invariant culture, deliberately.
# This machine is Dutch-locale, where "180.000" means one hundred and eighty thousand and
# "10,0%" means ten percent - so a report read by anyone else, or by an agent, says something
# other than what was measured, and [double]"1.5" parses as fifteen. This repository has already
# lost a measurement to exactly this class of bug: a DASL date literal in MM/dd/yyyy form was
# parsed in the machine locale and silently selected the wrong rows in both directions, which is
# what made a seven-day sweep report four folders swept and nothing found.
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::InvariantCulture

# PowerShell 7.4+ turns a non-zero exit from a native command into a terminating error when
# $ErrorActionPreference is Stop. This script deliberately RUNS things that are allowed to fail
# (check-pinned-constants.ps1, git in a directory that is not a repo) and reads their output, so
# that behaviour is switched off here and every such call checks $LASTEXITCODE for itself.
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# The marker every stored record carries. Written as two halves ON PURPOSE: assembled it is
# "OUTLOOKAI" + "-MEASUREMENT-RECORD-V1", and because that string never appears contiguously in
# any repository file, check-measurement-privacy.ps1 can grep the tracked tree for it and know
# that ANY hit is a real measurement record - whatever the file is called and wherever it
# landed. Spelling it out here as one literal would give that check a permanent false positive
# and force an allowlist, and an allowlist is a hole.
$script:RecordMarker = 'OUTLOOKAI' + '-MEASUREMENT-RECORD-V1'
$script:SchemaVersion = 1

# =====================================================================================
# The catalogue.
#
# Repo-side policy, and deliberately carries NO measured values - only ids, units, verdict
# classes and tolerances. That is what makes it safe to commit while the numbers are not.
#
# Direction classes, each of which exists because the others get it wrong somewhere:
#   both       - symmetric. Beyond tolerance in EITHER direction fails. The default, and the
#                one that catches "it got faster because it stopped doing something".
#   coverage   - a DECREASE of any size fails; an increase is reported and passes. Test counts,
#                folders scanned, samples taken: less coverage is never acceptable, more is
#                never the alarm.
#   noIncrease - an INCREASE of any size fails. Counters of things that went wrong.
#   mustBeZero - any non-zero fails, with or without a baseline. Claims the codebase makes
#                about itself that a run can check.
#   pinned     - any change at all fails. For values read out of the SOURCES: a constant is not
#                a noisy measurement, it either moved or it did not, and a budget constant
#                moving between releases is precisely the "course correction" this gate is for.
#                Move one legitimately with -Annotate.
#
# Source: 'machine' = measured here, privacy-sensitive.  'repo' = read out of the sources,
# not sensitive on its own - but it is stored in the same file, so the whole store is local.
# =====================================================================================

function New-Metric {
    param(
        [Parameter(Mandatory)][string] $Id,
        [Parameter(Mandatory)][string] $Unit,
        [Parameter(Mandatory)][string] $Group,
        [Parameter(Mandatory)][string] $Description,
        [string] $Direction = 'both',
        [object] $Tolerance = $null,
        [double] $MinAbsolute = 0,
        [string] $Source = 'machine',
        [hashtable] $Collector = $null
    )
    return [pscustomobject]@{
        Id          = $Id
        Unit        = $Unit
        Group       = $Group
        Description = $Description
        Direction   = $Direction
        Tolerance   = $Tolerance
        MinAbsolute = $MinAbsolute
        Source      = $Source
        Collector   = $Collector
    }
}

function New-SourceConstant {
    param(
        [Parameter(Mandatory)][string] $Id,
        [Parameter(Mandatory)][string] $File,
        [Parameter(Mandatory)][string] $Pattern,
        [Parameter(Mandatory)][string] $Description,
        [string] $Unit = 'ms',
        [double] $Scale = 1
    )
    return New-Metric -Id $Id -Unit $Unit -Group 'source constants' -Description $Description `
        -Direction 'pinned' -Source 'repo' `
        -Collector @{ Kind = 'SourceConstant'; File = $File; Pattern = $Pattern; Scale = $Scale }
}

$MailService = 'McpServer/OutlookAI.Core/Services/MailService.cs'
$Budgets = 'McpServer/OutlookAI.Core/Com/ComOperationBudgets.cs'
$Session = 'McpServer/OutlookAI.Core/Com/OutlookComSession.cs'
$Policy = 'McpServer/OutlookAI.ComHost/Supervision/ComHostPolicy.cs'
$Supervisor = 'McpServer/OutlookAI.ComHost/Supervision/ComHostSupervisor.cs'
$ProtocolFile = 'McpServer/OutlookAI.ComHost/Protocol/ComHostProtocol.cs'
$StaRunner = 'McpServer/OutlookAI.Core/Com/PumpedStaRunner.cs'
$CensusPlan = 'McpServer/OutlookAI.McpServer.Tests/T2/CensusIdentityPlan.cs'
$InboxArrival = 'McpServer/OutlookAI.McpServer.Tests/T2/LiveInboxArrival.cs'

$script:Catalogue = @(
    # ------------------------------------------------------------------ freshness sweep ----
    # Docs/vm-coverage-analysis.md 6.4(2): "whole-store 7-day sweep elapsed and per-store
    # breakdown". 36.6 s on the five-store profile is what OperationDeadlineMs is calibrated
    # against, so this is the single most load-bearing number in the set.
    (New-Metric -Id 'sweep.wholeStore7Day.elapsedMs' -Unit 'ms' -Group 'freshness sweep' `
        -Description 'search payload sweep.elapsedMs for a whole-profile 7-day sweep, COM host already warm (a cold host adds the 90 s connect floor and is not comparable).'),
    (New-Metric -Id 'sweep.wholeStore7Day.itemsSeen' -Unit 'items' -Group 'freshness sweep' `
        -Description 'sweep.itemsSeen for the same sweep. The denominator every elapsed figure above is a rate over; a drop here with elapsed flat means the sweep got slower per item.'),
    (New-Metric -Id 'sweep.perStore.elapsedMs.max' -Unit 'ms' -Group 'freshness sweep' `
        -Description 'slowest single store in the per-store breakdown. The number that decides whether one bad store can spend the whole budget.'),
    (New-Metric -Id 'sweep.perStore.elapsedMs.total' -Unit 'ms' -Group 'freshness sweep' `
        -Description 'sum over stores. Compared against sweep.wholeStore7Day.elapsedMs it says how much of the sweep is not per-store work.'),
    (New-Metric -Id 'sweep.foldersSwept' -Unit 'folders' -Group 'freshness sweep' -Direction 'coverage' `
        -Description 'sweep.foldersSwept. Fewer folders swept is lost coverage however fast the sweep got.'),
    (New-Metric -Id 'sweep.sortRefusedFolders' -Unit 'folders' -Group 'freshness sweep' -Direction 'mustBeZero' `
        -Description 'sweep.sortRefusedFolders. bea7fc9 made the claim "the received-date sort applies" checkable and nothing has checked it on a real profile since. Non-zero means capped folders kept an ARBITRARY slice of the window, which is the exact shape that produced a wrong 180 s budget.'),
    (New-Metric -Id 'sweep.itemCappedFolders' -Unit 'folders' -Group 'freshness sweep' `
        -Description 'count of folders truncated by SweepPerFolderCap. Documented as "never fires in steady state" on a real profile; movement off zero is that claim breaking.'),
    (New-Metric -Id 'sweep.itemsBodyCapped' -Unit 'items' -Group 'freshness sweep' `
        -Description 'items whose body was cut at SweepBodyCharsCap. Moves with the correspondents, not with the code, which is why it is reported rather than pinned.'),

    # ------------------------------------------------------------------ exhaustive scan ----
    # 6.4(2): "Inbox-only and Inbox-with-subfolders exhaustive scan elapsed; a 60-day
    # whole-store scan's foldersScanned and elapsedMs".
    (New-Metric -Id 'scan.inboxOnly.elapsedMs' -Unit 'ms' -Group 'exhaustive scan' `
        -Description 'exhaustive.elapsedMs, Inbox only, no subfolders.'),
    (New-Metric -Id 'scan.inboxWithSubfolders.elapsedMs' -Unit 'ms' -Group 'exhaustive scan' `
        -Description 'exhaustive.elapsedMs, Inbox with subfolders. 66.5 s here is the other half of the OperationDeadlineMs derivation.'),
    (New-Metric -Id 'scan.wholeStore60Day.elapsedMs' -Unit 'ms' -Group 'exhaustive scan' `
        -Description 'exhaustive.elapsedMs for a 60-day whole-store scan. Use a term that matches nothing, so the scan runs to the end of its budget instead of stopping at SearchTopCap.'),
    (New-Metric -Id 'scan.wholeStore60Day.foldersScanned' -Unit 'folders' -Group 'exhaustive scan' -Direction 'coverage' `
        -Description 'exhaustive.foldersScanned for that scan. 3 of 32 in 105 s is why ExhaustiveScanDeadlineMs is 615 s; this is the number that says whether it is enough now.'),
    (New-Metric -Id 'scan.wholeStore60Day.itemsPerSecond' -Unit 'items/s' -Group 'exhaustive scan' `
        -Description 'throughput for that scan. Step 5 of corpus-measurement-plan.md, which has never been run on either machine - so ExhaustiveTimeBudgetMs has no throughput measurement behind it at all. Expect NO BASELINE until somebody runs it.'),

    # ------------------------------------------------------------------ tripwire census ----
    # 6.4(2): "the census elapsed per store and folders fell back to counting".
    (New-Metric -Id 'census.elapsedMs.total' -Unit 'ms' -Group 'tripwire census' `
        -Description 'whole-profile baseline census wall clock. The 2026-08-20 table-read rewrite put this at 16.9 s for 5 stores / 159 folders / 2,044 items; before it, one store alone exceeded the 3-minute STA budget.'),
    (New-Metric -Id 'census.elapsedMs.maxStore' -Unit 'ms' -Group 'tripwire census' `
        -Description 'slowest single store. This is what the 3-minute STA join in LiveOutlookTestMailer actually bounds, and the number a live run refuses to start on.'),
    (New-Metric -Id 'census.storesWalked' -Unit 'stores' -Group 'tripwire census' -Direction 'coverage' `
        -Description 'stores the census reached. Recorded as a metric and not only as a condition, so a shrinking profile shows up in the diff instead of silently making every elapsed figure look better.'),
    (New-Metric -Id 'census.foldersWalked' -Unit 'folders' -Group 'tripwire census' -Direction 'coverage' `
        -Description 'folders the census reached.'),
    (New-Metric -Id 'census.itemsWalked' -Unit 'items' -Group 'tripwire census' -Direction 'coverage' `
        -Description 'items the census identified. The denominator of census.elapsedMs.total.'),
    (New-Metric -Id 'census.foldersDegradedToCount' -Unit 'folders' -Group 'tripwire census' -Direction 'noIncrease' `
        -Description 'folders the plan chose to walk and could not, so identity was lost for them. A table missing its columns on every folder would disable half the tripwire, and it must not do that quietly.'),

    # ------------------------------------------------------------------ search index -------
    # 6.4(2): "index frontier age sampled N times". StaleIndexNoticeMinutes is the p90 of this
    # over 177 probes; the whole staleness ladder is a statistic about one live mailbox.
    (New-Metric -Id 'index.frontierAgeMinutes.median' -Unit 'min' -Group 'search index' -Tolerance 0.50 `
        -Description 'median index frontier age over the sampled probes. Measured ~6 min. Tolerance is deliberately wide: this is a race between an indexer and arriving mail, not a code path.'),
    (New-Metric -Id 'index.frontierAgeMinutes.p90' -Unit 'min' -Group 'search index' -Tolerance 0.50 `
        -Description 'p90 of the same samples. StaleIndexNoticeMinutes (30) IS this number; if it has moved, the constant is stale.'),
    (New-Metric -Id 'index.frontierSampleCount' -Unit 'samples' -Group 'search index' -Direction 'coverage' `
        -Description 'how many probes the two figures above are computed from. A p90 over 5 samples is not a p90; fewer samples than last time is a weaker statistic and fails.'),
    (New-Metric -Id 'search.indexQueryMs.median' -Unit 'ms' -Group 'search index' `
        -Description 'median wall clock of the index query alone. Measured healthy at 60-550 ms, which is the whole justification for SearchIndexTimeoutSeconds = 60.'),
    (New-Metric -Id 'storeIndexProbe.delegateMissMs' -Unit 'ms' -Group 'search index' -MinAbsolute 5 `
        -Description 'per-store index probe, delegate-subtree miss. Measured 9-10 ms, which is why StoreIndexProbeBudgetMs is 1,500. Absolute floor of 5 ms so single-millisecond jitter cannot fail it.'),
    (New-Metric -Id 'storeIndexProbe.discoveryMissMs' -Unit 'ms' -Group 'search index' -MinAbsolute 5 `
        -Description 'per-store index probe, @-discovery miss. Measured 27-30 ms.'),

    # ------------------------------------------------------------------ COM host -----------
    (New-Metric -Id 'comHost.largestFrameBytes' -Unit 'bytes' -Group 'COM host' `
        -Description 'outlook_health comHost.largestFrameBytes high-water. THE cautionary tale: 432 KB read as 152x headroom was really a number bounded by a timeout, and the VM corpus later produced 10,734,599 bytes. Take it after the sweep, never before.'),
    (New-Metric -Id 'connect.attachAndHealthMs' -Unit 'ms' -Group 'COM host' `
        -Description 'attach to a running Outlook plus one health probe. Measured 1.0 s; ConnectDeadlineMs is 180 s against it.'),
    (New-Metric -Id 'connect.coldSearchMs' -Unit 'ms' -Group 'COM host' `
        -Description 'first search after a cold COM host start. Measured 6.2 s. Not comparable to any warm figure and must never be recorded as one.'),

    # ------------------------------------------------------------------ write paths --------
    (New-Metric -Id 'move.batch50.elapsedMs' -Unit 'ms' -Group 'write paths' `
        -Description 'the hub-only 50-item move/archive batch (6.4(3)). A PST move is a local file operation and an Exchange move is a server round trip; MoveBatchBudgetMs (240 s) has never been measured against the second.'),
    (New-Metric -Id 'transport.inboxArrivalSeconds' -Unit 's' -Group 'write paths' -Tolerance 0.50 `
        -Description 'send-to-self round trip. LiveInboxArrival.DeadlineSeconds is 180 because a real round trip once exceeded 120 s and failed a 17-minute run. Wide tolerance: this is the mail system, not the code.'),

    # ------------------------------------------------------------------ suite --------------
    (New-Metric -Id 'suite.testsPassed' -Unit 'tests' -Group 'suite' -Direction 'coverage' -Source 'repo' `
        -Description 'passing tests under the standing verification filter. Baseline 1,936.' `
        -Collector @{ Kind = 'TestLog'; Field = 'Passed' }),
    (New-Metric -Id 'suite.testsTotal' -Unit 'tests' -Group 'suite' -Direction 'coverage' -Source 'repo' `
        -Description 'total discovered under the same filter.' `
        -Collector @{ Kind = 'TestLog'; Field = 'Total' }),
    (New-Metric -Id 'suite.testsFailed' -Unit 'tests' -Group 'suite' -Direction 'mustBeZero' -Source 'repo' `
        -Description 'failures. Any is a failure, baseline or not.' `
        -Collector @{ Kind = 'TestLog'; Field = 'Failed' }),
    (New-Metric -Id 'suite.testsSkipped' -Unit 'tests' -Group 'suite' -Direction 'noIncrease' -Source 'repo' `
        -Description 'skips. A test that starts skipping is coverage lost without the count dropping, which is the quiet version of the same failure.' `
        -Collector @{ Kind = 'TestLog'; Field = 'Skipped' }),
    (New-Metric -Id 'suite.durationMs' -Unit 'ms' -Group 'suite' -Tolerance 0.35 -Source 'machine' `
        -Description 'suite wall clock. Tolerance 0.35 rather than the default 0.10 because this shares a machine with whatever else is running; tighter produced false failures in trial.' `
        -Collector @{ Kind = 'TestLog'; Field = 'DurationMs' }),

    # ------------------------------------------------------------------ invariants ---------
    (New-Metric -Id 'invariants.pinnedConstantChecks' -Unit 'checks' -Group 'invariants' -Direction 'coverage' -Source 'repo' `
        -Description 'cross-file invariants check-pinned-constants.ps1 asserted. Baseline 11. A check deleted is a check that stops proving anything, and nothing else notices.' `
        -Collector @{ Kind = 'PinnedConstants'; Field = 'Checks' }),
    (New-Metric -Id 'invariants.comHostThrownTypes' -Unit 'types' -Group 'invariants' -Direction 'coverage' -Source 'repo' `
        -Description 'exception types raised behind the IOutlookSession contract, all of which must be modelled by ComHostErrorMapper.' `
        -Collector @{ Kind = 'PinnedConstants'; Field = 'ThrownTypes' }),
    (New-Metric -Id 'invariants.comHostFilesScanned' -Unit 'files' -Group 'invariants' -Direction 'coverage' -Source 'repo' `
        -Description 'COM-host source files that scan reached. A drop means the scan stopped seeing a directory, which turns the check into one that always passes.' `
        -Collector @{ Kind = 'PinnedConstants'; Field = 'FilesScanned' }),

    # ------------------------------------------------------------------ source constants ---
    # Every one of these is a value SOMEBODY DERIVED FROM A MEASUREMENT. Pinned, so a change
    # fails until it is annotated with the measurement that justified it - which is exactly the
    # record that did not exist when 180 s was chosen.
    (New-SourceConstant -Id 'source.OperationDeadlineMs' -File $Budgets -Pattern 'OperationDeadlineMs\s*=\s*([0-9_]+)' `
        -Description 'shared COM operation deadline; 4.5x the slowest healthy operation measured.'),
    (New-SourceConstant -Id 'source.ConnectDeadlineMs' -File $Budgets -Pattern 'ConnectDeadlineMs\s*=\s*([0-9_]+)' `
        -Description 'COM session establishment, cold Outlook start included.'),
    (New-SourceConstant -Id 'source.HealthProbeDeadlineMs' -File $Budgets -Pattern 'HealthProbeDeadlineMs\s*=\s*([0-9_]+)' `
        -Description 'health probe. The instrument, not the work - deliberately short.'),
    (New-SourceConstant -Id 'source.HandshakeBudgetMs' -File $Budgets -Pattern 'HandshakeBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'COM host pipe handshake, both ends.'),
    (New-SourceConstant -Id 'source.ResultReturnHeadroomMs' -File $Budgets -Pattern 'ResultReturnHeadroomMs\s*=\s*([0-9_]+)' `
        -Description 'reserved for handing the result back rather than for doing work.'),
    (New-SourceConstant -Id 'source.ExhaustiveScanDeadlineMs' -File $Budgets -Pattern 'ExhaustiveScanDeadlineMs\s*=\s*([0-9_]+)' `
        -Description 'exhaustive scan hard deadline, its own class.'),
    (New-SourceConstant -Id 'source.SweepBudgetMs' -File $Budgets -Pattern 'FreshnessSweepBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'freshness sweep budget. The one that was derived from a measurement taken while the sort was failing; 600 s since 2026-08-24 is a CEILING awaiting that re-measurement, and it moved to ComOperationBudgets with the sweep''s own deadline class.'),
    (New-SourceConstant -Id 'source.FreshnessSweepDeadlineMs' -File $Budgets -Pattern 'FreshnessSweepDeadlineMs\s*=\s*([0-9_]+)' `
        -Description 'freshness class hard deadline - the threshold the sweep budget is judged against. Derived as SearchBudgetMs + ResultReturnHeadroomMs; narrowing the sweep budget must move it too.'),
    (New-SourceConstant -Id 'source.SweepPerFolderCap' -File $MailService -Pattern 'SweepPerFolderCap\s*=\s*([0-9_]+)' -Unit 'items' `
        -Description 'items per folder the sweep will open.'),
    (New-SourceConstant -Id 'source.ScopedSweepTimeBudgetMs' -File $Session -Pattern 'ScopedSweepTimeBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'subtree walk budget for a scoped sweep.'),
    (New-SourceConstant -Id 'source.SweepBodyCharsCap' -File $Session -Pattern 'SweepBodyCharsCap\s*=\s*([0-9_]+)' -Unit 'chars' `
        -Description 'per-body character cut in the sweep.'),
    (New-SourceConstant -Id 'source.SweepBodyBytesBudgetMiB' -File $Session -Pattern 'SweepBodyBytesBudget\s*=\s*([0-9_]+)L?\s*\*\s*1024\s*\*\s*1024' -Unit 'MiB' `
        -Description 'accumulated body bytes one sweep may return. Load-bearing since the 10,734,599-byte corpus high-water.'),
    (New-SourceConstant -Id 'source.SearchIndexTimeoutSeconds' -File $MailService -Pattern 'SearchIndexTimeoutSeconds\s*=\s*([0-9_]+)' -Unit 's' `
        -Description 'index half of the search budget.'),
    (New-SourceConstant -Id 'source.HealthIndexTimeoutSeconds' -File $MailService -Pattern 'HealthIndexTimeoutSeconds\s*=\s*([0-9_]+)' -Unit 's' `
        -Description 'index query timeout inside outlook_health.'),
    (New-SourceConstant -Id 'source.HealthPerStoreIndexBudgetMs' -File $MailService -Pattern 'HealthPerStoreIndexBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'per-store index probe budget inside outlook_health.'),
    (New-SourceConstant -Id 'source.StoreIndexProbeBudgetMs' -File $MailService -Pattern 'StoreIndexProbeBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'per-store index probe budget on the search path.'),
    (New-SourceConstant -Id 'source.StaleIndexNoticeMinutes' -File $MailService -Pattern 'StaleIndexNoticeMinutes\s*=\s*([0-9_.]+)' -Unit 'min' `
        -Description 'p90 of the dev profile index frontier age over 177 probes. Compare against index.frontierAgeMinutes.p90.'),
    (New-SourceConstant -Id 'source.VeryStaleAdviceMinutes' -File $MailService -Pattern 'VeryStaleAdviceMinutes\s*=\s*([0-9_.]+)' -Unit 'min' `
        -Description 'upper rung of the staleness ladder.'),
    (New-SourceConstant -Id 'source.MoveBatchBudgetMs' -File $MailService -Pattern 'MoveBatchBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'whole move/archive batch. Compare against move.batch50.elapsedMs.'),
    (New-SourceConstant -Id 'source.MinimumItemBudgetMs' -File $MailService -Pattern 'MinimumItemBudgetMs\s*=\s*([0-9_]+)' `
        -Description 'floor per item inside a batch move.'),
    (New-SourceConstant -Id 'source.SearchTopCap' -File $MailService -Pattern 'SearchTopCap\s*=\s*([0-9_]+)' -Unit 'items' `
        -Description 'largest result set that crosses the pipe; ResultReturnHeadroomMs is sized against it.'),
    (New-SourceConstant -Id 'source.MaxFrameBytesMiB' -File $ProtocolFile -Pattern 'MaxFrameBytes\s*=\s*([0-9_]+)\s*\*\s*1024\s*\*\s*1024' -Unit 'MiB' `
        -Description 'protocol frame ceiling. comHost.largestFrameBytes is measured against it.'),
    (New-SourceConstant -Id 'source.UnresponsiveTimeoutThreshold' -File $Policy -Pattern 'UnresponsiveTimeoutThreshold\s*=\s*([0-9_]+)' -Unit 'timeouts' `
        -Description 'consecutive timeouts before the breaker opens.'),
    (New-SourceConstant -Id 'source.UnresponsiveCooldownMs' -File $Policy -Pattern 'UnresponsiveCooldownMilliseconds\s*=\s*([0-9_]+)' `
        -Description 'how long the breaker stays open.'),
    (New-SourceConstant -Id 'source.StartFailureBackoffThreshold' -File $Policy -Pattern 'StartFailureBackoffThreshold\s*=\s*([0-9_]+)' -Unit 'failures' `
        -Description 'start failures before backing off.'),
    (New-SourceConstant -Id 'source.StartBackoffMs' -File $Policy -Pattern 'StartBackoffMilliseconds\s*=\s*([0-9_]+)' `
        -Description 'backoff after repeated start failures.'),
    (New-SourceConstant -Id 'source.AutostartCooldownMs' -File $Policy -Pattern 'AutostartCooldownMilliseconds\s*=\s*([0-9_]+)' `
        -Description 'cooldown between autostart attempts.'),
    (New-SourceConstant -Id 'source.MinimumDispatchDeadlineMs' -File $Policy -Pattern 'MinimumDispatchDeadlineMilliseconds\s*=\s*([0-9_]+)' `
        -Description 'floor under any dispatch deadline.'),
    (New-SourceConstant -Id 'source.CleanExitGraceMs' -File $Supervisor -Pattern 'CleanExitGraceMilliseconds\s*=\s*([0-9_]+)' `
        -Description 'grace given to a COM host asked to exit.'),
    (New-SourceConstant -Id 'source.StaRetryAfterMs' -File $StaRunner -Pattern 'RetryAfterMs\s*=\s*([0-9_]+)' `
        -Description 'SERVERCALL_RETRYLATER retry interval.'),
    (New-SourceConstant -Id 'source.StaGiveUpAfterMs' -File $StaRunner -Pattern 'GiveUpAfterMs\s*=\s*([0-9_]+)' `
        -Description 'how long the STA runner retries a busy Outlook before giving up.'),
    (New-SourceConstant -Id 'source.CensusPerFolderLimit' -File $CensusPlan -Pattern 'DefaultPerFolderLimit\s*=\s*([0-9_]+)' -Unit 'items' `
        -Description 'largest folder the tripwire baseline will identify item by item.'),
    (New-SourceConstant -Id 'source.CensusPerStoreItemBudget' -File $CensusPlan -Pattern 'DefaultPerStoreItemBudget\s*=\s*([0-9_]+)' -Unit 'items' `
        -Description 'identity budget per store. Bounds the whole profile at stores x this.'),
    (New-SourceConstant -Id 'source.CensusRepeatGrowthHeadroom' -File $CensusPlan -Pattern 'RepeatGrowthHeadroom\s*=\s*([0-9_]+)' -Unit 'x' `
        -Description 'growth a folder may show between the two censuses and still be walked.'),
    (New-SourceConstant -Id 'source.LiveInboxArrivalDeadlineSeconds' -File $InboxArrival -Pattern 'DeadlineSeconds\s*=\s*([0-9_]+)' -Unit 's' `
        -Description 'live transport arrival deadline. Compare against transport.inboxArrivalSeconds.'),
    (New-SourceConstant -Id 'source.RecipientResnapshotDelayMs' -File $Session -Pattern 'RecipientResnapshotDelayMs\s*=\s*([0-9_]+)' `
        -Description 'wait before re-reading resolved recipients.'),
    (New-SourceConstant -Id 'source.ExplorerFolderSettleDelayMs' -File $Session -Pattern 'ExplorerFolderSettleDelayMs\s*=\s*([0-9_]+)' `
        -Description 'wait for an Explorer to settle on a folder change.')
)

$script:CatalogueById = @{}
foreach ($m in $script:Catalogue) { $script:CatalogueById[$m.Id] = $m }

# =====================================================================================
# Paths, and the guards that keep the store out of the repository.
# =====================================================================================

function Resolve-FullPath([string] $path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }
    return [System.IO.Path]::GetFullPath($path.TrimEnd('\', '/'))
}

function Test-PathUnder([string] $child, [string] $parent) {
    if (-not $child -or -not $parent) { return $false }
    $c = (Resolve-FullPath $child) + [System.IO.Path]::DirectorySeparatorChar
    $p = (Resolve-FullPath $parent) + [System.IO.Path]::DirectorySeparatorChar
    return $c.StartsWith($p, [System.StringComparison]::OrdinalIgnoreCase)
}

# Walks up from a directory looking for .git - a directory (ordinary clone) or a FILE (a
# worktree, which is what this repo uses for agent branches, and which a naive Test-Path -Type
# Container check would walk straight past).
function Find-GitWorkTree([string] $start) {
    $dir = Resolve-FullPath $start
    while ($dir) {
        $dotGit = Join-Path $dir '.git'
        if (Test-Path -LiteralPath $dotGit) { return $dir }
        $parent = Split-Path -Parent $dir
        if (-not $parent -or $parent -eq $dir) { return $null }
        $dir = $parent
    }
    return $null
}

function Assert-StoreOutsideRepository([string] $store, [string] $repo) {
    $full = Resolve-FullPath $store

    if (Test-PathUnder $full $repo) {
        throw "REFUSING TO WRITE: the baseline store resolves to '$full', which is inside the repository at '$repo'. These numbers are statistics about one machine and must never become repository content. Point -StoreRoot somewhere outside the repo, or leave it at the default under %LOCALAPPDATA%."
    }

    # Broader than the repo check on purpose: a store under ANY working tree is one `git add`
    # from being published, including a repo this script has never heard of.
    $tree = Find-GitWorkTree $full
    if ($tree) {
        throw "REFUSING TO WRITE: the baseline store resolves to '$full', which is inside the git working tree at '$tree'. The store must live outside every working tree - the default is %LOCALAPPDATA%\OutlookAI\Measurements."
    }

    return $full
}

function Get-DefaultStoreRoot {
    $localAppData = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = [System.Environment]::GetFolderPath('LocalApplicationData')
    }
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw "Cannot locate %LOCALAPPDATA%, so there is nowhere machine-local to put the baseline. Pass -StoreRoot explicitly (it must be outside every git working tree)."
    }
    return (Join-Path (Join-Path $localAppData 'OutlookAI') 'Measurements')
}

# Single funnel for every write. Nothing in this script writes a file except through here, so
# the "outside the repository" guarantee is one function rather than a habit.
function Write-StoreFile {
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)][AllowEmptyString()][string] $Content,
        [switch] $Append
    )
    $full = Resolve-FullPath $Path
    if (-not (Test-PathUnder $full $script:StoreRootResolved) -and $full -ne $script:StoreRootResolved) {
        throw "INTERNAL GUARD: refused a write to '$full', which is not under the validated store root '$($script:StoreRootResolved)'."
    }
    $dir = Split-Path -Parent $full
    if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    if ($Append) {
        Add-Content -LiteralPath $full -Value $Content -Encoding UTF8
    } else {
        Set-Content -LiteralPath $full -Value $Content -Encoding UTF8
    }
}

function Initialize-Store([string] $root) {
    if (-not (Test-Path -LiteralPath $root)) {
        $null = New-Item -ItemType Directory -Path $root -Force
    }
    $readme = Join-Path $root 'README.txt'
    if (-not (Test-Path -LiteralPath $readme)) {
        $text = @"
OutlookAI measurement baselines - MACHINE LOCAL, DO NOT COPY OUT OF THIS MACHINE.

What this is
  An append-only history of measurement runs taken before releases of OutlookAI, written by
  .github/scripts/measurement-gate.ps1 in the OutlookAI repository.

Why it is here and not in the repository
  These numbers are meaningful only relative to older numbers from THIS machine. They are not
  representative of anyone else's system, and they are statistics about this machine that the
  maintainer does not want published. The repository is public. So: they are never committed,
  they never appear in release notes, and the gate refuses to write inside any git working
  tree or to run its comparing modes under CI.

Files
  history.jsonl      one JSON record per run, appended, never rewritten.
  annotations.jsonl  one record per deliberate baseline move, with the reason. Appended.
  reports\           the human-readable verdict for each run.

Reading it
  Get-Content history.jsonl | ForEach-Object { `$_ | ConvertFrom-Json }
  or:  pwsh -File .github\scripts\measurement-gate.ps1 -Show

Deleting it
  Safe. The next run reports "no baseline - nothing to compare" and refuses to pass until the
  new baseline is accepted by hand with -AcceptNewBaseline.
"@
        Write-StoreFile -Path $readme -Content $text
    }
}

# =====================================================================================
# Store I/O.
# =====================================================================================

function Get-HistoryPath { return (Join-Path $script:StoreRootResolved 'history.jsonl') }
function Get-AnnotationsPath { return (Join-Path $script:StoreRootResolved 'annotations.jsonl') }

function Read-JsonLines([string] $path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $out = New-Object System.Collections.ArrayList
    $lineNo = 0
    foreach ($line in (Get-Content -LiteralPath $path -Encoding UTF8)) {
        $lineNo++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $null = $out.Add(($line | ConvertFrom-Json))
        } catch {
            # Never skip silently: a record we cannot parse is a record we cannot compare
            # against, and pretending it is not there is how a baseline quietly narrows.
            throw "Corrupt record at line $lineNo of '$path': $($_.Exception.Message). Fix or remove the line; the file is append-only plain JSON Lines and every line stands alone."
        }
    }
    return @($out.ToArray())
}

# =====================================================================================
# Collectors. All of them offline: no Outlook, no mailbox, no network.
# =====================================================================================

function Read-RepoText([string] $relative) {
    $full = Join-Path $RepoRoot $relative
    if (-not (Test-Path -LiteralPath $full)) { return $null }
    return (Get-Content -LiteralPath $full -Raw)
}

function Invoke-SourceConstantCollector([pscustomobject] $metric) {
    $text = Read-RepoText $metric.Collector.File
    if ($null -eq $text) {
        return @{ Ok = $false; Detail = "$($metric.Collector.File) does not exist - it moved, and this collector now proves nothing." }
    }
    $found = [regex]::Match($text, $metric.Collector.Pattern)
    if (-not $found.Success) {
        # Same discipline as check-pinned-constants.ps1: a collector whose regex stopped
        # matching is a collector that has switched itself off, and that is worse than absent.
        return @{ Ok = $false; Detail = "pattern did not match in $($metric.Collector.File): $($metric.Collector.Pattern)" }
    }
    $raw = $found.Groups[1].Value -replace '_', ''
    return @{ Ok = $true; Value = ([double]$raw * $metric.Collector.Scale) }
}

# Reads the shape both `dotnet test` and a .trx use. Asserts it found what it was looking for
# rather than defaulting to zero, because "0 failed" from a log that was never parsed is the
# most reassuring wrong answer available.
function Read-TestLog([string] $path) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "-TestLog '$path' does not exist."
    }
    $text = Get-Content -LiteralPath $path -Raw

    $summary = [regex]::Match($text,
        'Failed:\s*(?<failed>\d+),\s*Passed:\s*(?<passed>\d+),\s*Skipped:\s*(?<skipped>\d+),\s*Total:\s*(?<total>\d+),\s*Duration:\s*(?<dur>[0-9.]+)\s*(?<unit>ms|s|m)\b')
    if (-not $summary.Success) {
        throw "Could not find a test summary line in '$path'. Expected the `"Failed: N, Passed: N, Skipped: N, Total: N, Duration: X`" shape that 'dotnet test' prints. Parsing nothing and reporting zeroes would be worse than failing here."
    }

    $duration = [double]$summary.Groups['dur'].Value
    switch ($summary.Groups['unit'].Value) {
        'ms' { $durationMs = $duration }
        's' { $durationMs = $duration * 1000 }
        'm' { $durationMs = $duration * 60000 }
        default { $durationMs = $duration * 1000 }
    }

    return @{
        Failed     = [double]$summary.Groups['failed'].Value
        Passed     = [double]$summary.Groups['passed'].Value
        Skipped    = [double]$summary.Groups['skipped'].Value
        Total      = [double]$summary.Groups['total'].Value
        DurationMs = $durationMs
    }
}

function Invoke-PinnedConstantsCollector {
    $script = Join-Path $RepoRoot '.github/scripts/check-pinned-constants.ps1'
    if (-not (Test-Path -LiteralPath $script)) {
        return @{ Ok = $false; Detail = 'check-pinned-constants.ps1 is missing.' }
    }

    $output = ''
    $exit = -1
    try {
        $output = & pwsh -NoProfile -File $script 2>&1 | Out-String
        $exit = $LASTEXITCODE
    } catch {
        return @{ Ok = $false; Detail = "could not run check-pinned-constants.ps1: $($_.Exception.Message)" }
    }

    $checks = [regex]::Match($output, 'All (\d+) cross-file invariants hold')
    if (-not $checks.Success) {
        $checks = [regex]::Match($output, '(\d+) of (\d+) cross-file invariants failed')
        if ($checks.Success) {
            return @{ Ok = $false; Detail = "check-pinned-constants.ps1 FAILED ($($checks.Groups[1].Value) of $($checks.Groups[2].Value) invariants). Fix that before measuring anything - a drifted pin means the constants under measurement are not the ones that ship." }
        }
        return @{ Ok = $false; Detail = "could not read a check count out of check-pinned-constants.ps1 (exit $exit). Its output shape changed and this collector now proves nothing." }
    }

    $result = @{ Ok = $true; Checks = [double]$checks.Groups[1].Value }

    $comHost = [regex]::Match($output, 'COM host failure types - (\d+) raised, all modelled \((\d+) files scanned\)')
    if ($comHost.Success) {
        $result.ThrownTypes = [double]$comHost.Groups[1].Value
        $result.FilesScanned = [double]$comHost.Groups[2].Value
    }
    return $result
}

# =====================================================================================
# Building a run record.
# =====================================================================================

function Get-GitFact([string] $gitArgs) {
    try {
        $argv = @('-C', $RepoRoot) + @($gitArgs -split ' ')
        $out = & git @argv 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        return ($out | Out-String).Trim()
    } catch {
        return $null
    }
}

function New-RunRecord {
    $now = [DateTime]::UtcNow

    $dirty = $null
    $status = Get-GitFact 'status --porcelain'
    if ($null -ne $status) { $dirty = -not [string]::IsNullOrWhiteSpace($status) }

    $sdk = $null
    try { $sdk = (& dotnet --version 2>$null | Out-String).Trim() } catch { $sdk = $null }

    return [pscustomobject]@{
        marker       = $script:RecordMarker
        schema       = $script:SchemaVersion
        runId        = ($now.ToString('yyyyMMdd-HHmmss') + '-' + ([guid]::NewGuid().ToString('N').Substring(0, 6)))
        recordedAtUtc = $now.ToString('o')
        label        = $Label
        provenance   = [pscustomobject]@{
            machine       = $env:COMPUTERNAME
            user          = $env:USERNAME
            os            = [System.Environment]::OSVersion.VersionString
            powershell    = $PSVersionTable.PSVersion.ToString()
            dotnetSdk     = $sdk
            repoRoot      = $RepoRoot
            gitCommit     = (Get-GitFact 'rev-parse HEAD')
            gitShort      = (Get-GitFact 'rev-parse --short HEAD')
            gitBranch     = (Get-GitFact 'rev-parse --abbrev-ref HEAD')
            gitDirty      = $dirty
            gateVersion   = $script:SchemaVersion
        }
        conditions   = [pscustomobject]@{
            profile       = $script:RunConditions.profile
            indexed       = $script:RunConditions.indexed
            corpusId      = $script:RunConditions.corpusId
            # The corpus ANCHOR, not the build date. corpus-build is deterministic from
            # (corpusId, seed, anchor), and every date band in it is relative to the anchor -
            # so an anchor plus the run date is the only thing that says whether the 1/7/30/60
            # day windows still select anything. A corpus that has aged past its own windows
            # measures an EMPTY selection and reports a healthy, stable, very fast number.
            corpusAnchor  = $script:RunConditions.corpusAnchor
            corpusAgeDays = $script:CorpusAgeDays
            storeSet      = $script:RunConditions.storeSet
            notes         = $script:RunConditions.notes
        }
        metrics      = @()
    }
}

function Get-ScopeKey([object] $conditions) {
    $profile = 'unknown'
    $corpus = ''
    if ($conditions) {
        if ($conditions.PSObject.Properties['profile'] -and $conditions.profile) { $profile = [string]$conditions.profile }
        if ($conditions.PSObject.Properties['corpusId'] -and $conditions.corpusId) { $corpus = [string]$conditions.corpusId }
    }
    return "$profile|$corpus"
}

# =====================================================================================
# Comparison.
# =====================================================================================

function Get-Median([double[]] $values) {
    $sorted = @($values | Sort-Object)
    $n = $sorted.Count
    if ($n -eq 0) { return $null }
    if ($n % 2 -eq 1) { return $sorted[[int](($n - 1) / 2)] }
    return (($sorted[$n / 2 - 1] + $sorted[$n / 2]) / 2)
}

function Get-AnnotationCutoff([object[]] $annotations, [string] $metricId) {
    $cut = $null
    foreach ($a in $annotations) {
        $covers = ($a.metric -eq '*') -or ($a.metric -eq $metricId)
        if (-not $covers) { continue }
        $at = [DateTime]::Parse($a.recordedAtUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
        if ($null -eq $cut -or $at -gt $cut) { $cut = $at }
    }
    return $cut
}

function Get-BaselineFor {
    param(
        [object[]] $History,
        [string] $MetricId,
        [string] $ScopeKey,
        [object[]] $Annotations
    )
    $cutoff = Get-AnnotationCutoff $Annotations $MetricId
    $values = New-Object System.Collections.ArrayList
    $samples = New-Object System.Collections.ArrayList

    # Newest first, take at most $BaselineRuns. The median of a handful resists one noisy run
    # without letting an old value anchor the comparison forever. The individual samples are
    # kept as well as the median: the reader needs the spread, and a median alone hides both
    # "these three runs are wildly apart" and "these three runs are byte-identical", which are
    # the two shapes a tolerance is blind to.
    for ($i = $History.Count - 1; $i -ge 0; $i--) {
        $rec = $History[$i]
        if ((Get-ScopeKey $rec.conditions) -ne $ScopeKey) { continue }
        if ($cutoff) {
            $at = [DateTime]::Parse($rec.recordedAtUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            if ($at -le $cutoff) { continue }
        }
        foreach ($m in @($rec.metrics)) {
            if ($m.id -eq $MetricId) {
                $null = $values.Add([double]$m.value)
                $sampleConditions = $rec.conditions
                if ($m.PSObject.Properties['conditions'] -and $m.conditions) {
                    # A per-metric conditions block overrides the run's, because some numbers
                    # are taken under conditions of their own (a frame high-water read after
                    # the sweep is not the same reading as one taken before it).
                    $sampleConditions = $m.conditions
                }
                $null = $samples.Add([pscustomobject]@{
                        runId         = $rec.runId
                        recordedAtUtc = $rec.recordedAtUtc
                        value         = [double]$m.value
                        gitShort      = $rec.provenance.gitShort
                        gitDirty      = $rec.provenance.gitDirty
                        conditions    = $sampleConditions
                    })
                break
            }
        }
        if ($values.Count -ge $BaselineRuns) { break }
    }

    if ($values.Count -eq 0) {
        return [pscustomobject]@{ HasBaseline = $false; Value = $null; Runs = 0; Cutoff = $cutoff; Values = @(); Samples = @() }
    }
    return [pscustomobject]@{
        HasBaseline = $true
        Value       = (Get-Median ([double[]]$values.ToArray()))
        Runs        = $values.Count
        Cutoff      = $cutoff
        Values      = @($values.ToArray())
        Samples     = @($samples.ToArray())
    }
}

function Get-EffectiveTolerance([pscustomobject] $metric) {
    if ($null -ne $metric.Tolerance) { return [double]$metric.Tolerance }
    return $Tolerance
}

function Compare-Metric {
    param(
        [pscustomobject] $Metric,
        [double] $Now,
        [pscustomobject] $Baseline
    )

    $tol = Get-EffectiveTolerance $Metric

    if ($Metric.Direction -eq 'mustBeZero') {
        if ($Now -ne 0) {
            return @{ Verdict = 'FAIL'; Detail = "must be zero, is $(Format-Value $Now $Metric.Unit). This is a claim the codebase makes about itself; a non-zero value means the claim is false right now." }
        }
        return @{ Verdict = 'OK'; Detail = 'zero, as claimed' }
    }

    if (-not $Baseline.HasBaseline) {
        return @{ Verdict = 'NEW'; Detail = 'no baseline - nothing to compare' }
    }

    $base = [double]$Baseline.Value
    $delta = $Now - $base

    if ($Metric.Direction -eq 'pinned') {
        if ($delta -ne 0) {
            return @{ Verdict = 'FAIL'; Detail = "constant moved $(Format-Value $base $Metric.Unit) -> $(Format-Value $Now $Metric.Unit). A derived constant is not a noisy measurement. If a measurement justified this, record it with -Annotate -Metric $($Metric.Id) -Reason '...'." }
        }
        return @{ Verdict = 'OK'; Detail = 'unchanged' }
    }

    if ($Metric.Direction -eq 'coverage') {
        if ($delta -lt 0) {
            return @{ Verdict = 'FAIL'; Detail = "coverage fell $(Format-Value $base $Metric.Unit) -> $(Format-Value $Now $Metric.Unit). Any decrease fails; less coverage is never the acceptable direction." }
        }
        if ($delta -gt 0) {
            return @{ Verdict = 'OK'; Detail = "grew by $(Format-Value $delta $Metric.Unit)" }
        }
        return @{ Verdict = 'OK'; Detail = 'unchanged' }
    }

    if ($Metric.Direction -eq 'noIncrease') {
        if ($delta -gt 0) {
            return @{ Verdict = 'FAIL'; Detail = "rose $(Format-Value $base $Metric.Unit) -> $(Format-Value $Now $Metric.Unit). Any increase fails - this counts something going wrong." }
        }
        if ($delta -lt 0) {
            return @{ Verdict = 'OK'; Detail = "fell by $(Format-Value ([math]::Abs($delta)) $Metric.Unit)" }
        }
        return @{ Verdict = 'OK'; Detail = 'unchanged' }
    }

    # 'both' - symmetric, and the direction is named in the message because a number that got
    # FASTER is the one a reader is most likely to wave through.
    if ($base -eq 0) {
        if ($Now -ne 0) {
            return @{ Verdict = 'FAIL'; Detail = "moved off zero to $(Format-Value $Now $Metric.Unit). The baseline was zero, so there is no percentage to compare - any movement is the whole signal." }
        }
        return @{ Verdict = 'OK'; Detail = 'zero, as before' }
    }

    $relative = $delta / $base
    $absDelta = [math]::Abs($delta)

    if ($Metric.MinAbsolute -gt 0 -and $absDelta -lt $Metric.MinAbsolute) {
        return @{ Verdict = 'OK'; Detail = ('{0:+0.0%;-0.0%;0.0%}' -f $relative) + ", inside the $(Format-Value $Metric.MinAbsolute $Metric.Unit) absolute floor" }
    }

    if ([math]::Abs($relative) -gt $tol) {
        $direction = if ($relative -gt 0) { 'larger/slower' } else { 'smaller/faster' }
        $extra = ''
        if ($relative -lt 0) {
            $extra = ' A number that got smaller is not good news by default - the usual cause here is that something stopped being done (a sort that started failing, a folder set that shrank, a filter that stopped matching). Prove which before accepting it.'
        }
        return @{
            Verdict = 'FAIL'
            Detail  = ("$direction by " + ('{0:0.0%}' -f [math]::Abs($relative)) + " (tolerance " + ('{0:0.0%}' -f $tol) + "): $(Format-Value $base $Metric.Unit) -> $(Format-Value $Now $Metric.Unit).$extra")
        }
    }

    return @{ Verdict = 'OK'; Detail = ('{0:+0.0%;-0.0%;0.0%}' -f $relative) + " of $(Format-Value $base $Metric.Unit)" }
}

function Format-Value([double] $v, [string] $unit) {
    if ([math]::Abs($v - [math]::Round($v)) -lt 0.0000001) {
        return ('{0:N0} {1}' -f $v, $unit)
    }
    return ('{0:N3} {1}' -f $v, $unit)
}

# =====================================================================================
# Modes.
# =====================================================================================

function Show-Template {
    $template = @'
{
  "label": "pre-release 2.2.0",
  "conditions": {
    "profile": "production",
    "indexed": "indexed",
    "corpusId": null,
    "corpusAnchor": null,
    "storeSet": "5 stores / 159 folders / 2044 items",
    "notes": "warm COM host; first sweep after start discarded; Outlook otherwise idle"
  },
  "metrics": {
    "sweep.wholeStore7Day.elapsedMs": 0,
    "sweep.wholeStore7Day.itemsSeen": 0,
    "sweep.perStore.elapsedMs.max": 0,
    "sweep.perStore.elapsedMs.total": 0,
    "sweep.foldersSwept": 0,
    "sweep.sortRefusedFolders": 0,
    "sweep.itemCappedFolders": 0,
    "sweep.itemsBodyCapped": 0,
    "scan.inboxOnly.elapsedMs": 0,
    "scan.inboxWithSubfolders.elapsedMs": 0,
    "scan.wholeStore60Day.elapsedMs": 0,
    "scan.wholeStore60Day.foldersScanned": 0,
    "scan.wholeStore60Day.itemsPerSecond": 0,
    "census.elapsedMs.total": 0,
    "census.elapsedMs.maxStore": 0,
    "census.storesWalked": 0,
    "census.foldersWalked": 0,
    "census.itemsWalked": 0,
    "census.foldersDegradedToCount": 0,
    "index.frontierAgeMinutes.median": 0,
    "index.frontierAgeMinutes.p90": 0,
    "index.frontierSampleCount": 0,
    "search.indexQueryMs.median": 0,
    "storeIndexProbe.delegateMissMs": 0,
    "storeIndexProbe.discoveryMissMs": 0,
    "comHost.largestFrameBytes": { "value": 0, "conditions": { "notes": "read AFTER the sweep, never before" } },
    "connect.attachAndHealthMs": 0,
    "connect.coldSearchMs": 0,
    "move.batch50.elapsedMs": 0,
    "transport.inboxArrivalSeconds": 0
  }
}
'@
    Write-Host $template
}

function Invoke-SelfTest {
    # CI-safe on purpose: touches no store, prints no measured value, and is what
    # check-measurement-privacy.ps1 calls to prove the catalogue and the document still
    # describe the same set.
    $problems = New-Object System.Collections.ArrayList
    $checked = 0

    $seen = @{}
    foreach ($m in $script:Catalogue) {
        if ($seen.ContainsKey($m.Id)) { $null = $problems.Add("duplicate metric id '$($m.Id)'") }
        $seen[$m.Id] = $true
        if ($m.Direction -notin @('both', 'coverage', 'noIncrease', 'mustBeZero', 'pinned')) {
            $null = $problems.Add("metric '$($m.Id)' has unknown direction '$($m.Direction)'")
        }
        if ($null -ne $m.Tolerance -and ([double]$m.Tolerance -le 0 -or [double]$m.Tolerance -ge 3)) {
            $null = $problems.Add("metric '$($m.Id)' has an implausible tolerance $($m.Tolerance)")
        }
    }
    $checked++

    # Every source-constant collector must still match. A collector whose regex stopped
    # matching is a gate entry that has quietly switched itself off.
    foreach ($m in $script:Catalogue) {
        if ($m.Collector -and $m.Collector.Kind -eq 'SourceConstant') {
            $r = Invoke-SourceConstantCollector $m
            if (-not $r.Ok) { $null = $problems.Add("collector for '$($m.Id)': $($r.Detail)") }
        }
    }
    $checked++

    # Catalogue <-> document, both directions.
    $docPath = Join-Path $RepoRoot 'Docs/measurement-gate.md'
    if (-not (Test-Path -LiteralPath $docPath)) {
        $null = $problems.Add("Docs/measurement-gate.md is missing - the gate would be undocumented.")
    } else {
        $doc = Get-Content -LiteralPath $docPath -Raw
        foreach ($m in $script:Catalogue) {
            if ($doc -cnotmatch ('`' + [regex]::Escape($m.Id) + '`')) {
                $null = $problems.Add("metric '$($m.Id)' is gated but not documented in Docs/measurement-gate.md")
            }
        }
        $documented = @([regex]::Matches($doc, '`((?:sweep|scan|census|index|search|storeIndexProbe|comHost|connect|move|transport|suite|invariants|source)\.[A-Za-z0-9_.]+)`') |
            ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        if ($documented.Count -eq 0) {
            $null = $problems.Add("found no metric ids in Docs/measurement-gate.md - the document changed shape and this check no longer proves anything.")
        }
        foreach ($id in $documented) {
            if (-not $script:CatalogueById.ContainsKey($id)) {
                $null = $problems.Add("Docs/measurement-gate.md documents '$id', which the catalogue does not gate")
            }
        }
    }
    $checked++

    Write-Host "Measurement gate self-test: $($script:Catalogue.Count) metrics, $checked structural checks."
    if ($problems.Count -gt 0) {
        foreach ($p in $problems) { Write-Host "::error::MEASUREMENT GATE SELF-TEST - $p" }
        Write-Host "$($problems.Count) self-test problem(s). The gate is describing something other than what it gates."
        return 1
    }
    Write-Host "  OK   catalogue is internally consistent, every source collector still matches, and the catalogue and Docs/measurement-gate.md describe the same $($script:Catalogue.Count) metrics."
    return 0
}

function Show-History([int] $count) {
    $history = Read-JsonLines (Get-HistoryPath)
    if ($history.Count -eq 0) {
        Write-Host "No history yet at $(Get-HistoryPath)."
        return 0
    }
    Write-Host "$($history.Count) runs in $(Get-HistoryPath). Showing the last $([math]::Min($count, $history.Count))."
    Write-Host ''
    $start = [math]::Max(0, $history.Count - $count)
    for ($i = $start; $i -lt $history.Count; $i++) {
        $r = $history[$i]
        $dirty = if ($r.provenance.gitDirty) { ' +dirty' } else { '' }
        # ConvertFrom-Json turns the ISO timestamp back into a DateTime, and letting it render
        # itself gives MM/dd/yyyy - which reads as a different day to half the world. Always
        # round-trip it.
        $when = $r.recordedAtUtc
        if ($when -is [DateTime]) { $when = $when.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
        Write-Host ("{0}  {1}  {2}{3}  [{4}]  {5} metrics  {6}" -f `
                $r.runId, $when, $r.provenance.gitShort, $dirty,
            (Get-ScopeKey $r.conditions), @($r.metrics).Count, $r.label)
        Write-Host ("      conditions: indexed={0} storeSet={1} {2}" -f $r.conditions.indexed, $r.conditions.storeSet, $r.conditions.notes)
    }
    return 0
}

function Add-Annotation([string] $metricId, [string] $why) {
    if ([string]::IsNullOrWhiteSpace($why)) {
        Write-Host "::error::-Reason is required. An unexplained baseline reset is indistinguishable from hiding a regression, which is the one thing this gate exists to prevent. Say what moved and what measurement justifies it."
        return 3
    }
    if ($metricId -ne '*' -and -not $script:CatalogueById.ContainsKey($metricId)) {
        Write-Host "::error::'$metricId' is not a catalogued metric. Use -ResetBaseline for all of them, or check the id with -SelfTest."
        return 3
    }

    $history = Read-JsonLines (Get-HistoryPath)
    $lastRun = if ($history.Count -gt 0) { $history[$history.Count - 1].runId } else { $null }

    $record = [pscustomobject]@{
        marker        = $script:RecordMarker
        schema        = $script:SchemaVersion
        kind          = 'annotation'
        recordedAtUtc = [DateTime]::UtcNow.ToString('o')
        metric        = $metricId
        reason        = $why
        afterRunId    = $lastRun
        author        = $env:USERNAME
        gitCommit     = (Get-GitFact 'rev-parse HEAD')
    }
    Write-StoreFile -Path (Get-AnnotationsPath) -Content ($record | ConvertTo-Json -Depth 6 -Compress) -Append

    if ($metricId -eq '*') {
        Write-Host "Baseline reset for ALL metrics. Runs recorded before now are no longer compared against."
    } else {
        Write-Host "Baseline reset for '$metricId'. Runs recorded before now are no longer compared against for that metric."
    }
    Write-Host "  reason: $why"
    Write-Host "  recorded in $(Get-AnnotationsPath)"
    return 0
}

# =====================================================================================
# Main.
# =====================================================================================

$RepoRoot = Resolve-FullPath $RepoRoot

# ---- modes that never touch the store, and are therefore CI-safe -----------------------
if ($Template) { Show-Template; exit 0 }
if ($SelfTest) { exit (Invoke-SelfTest) }
if ($ListMetrics) {
    # Markdown rows for Docs/measurement-gate.md. The document has to name every metric (the
    # self-test insists on it in both directions), so adding one means editing the document -
    # this is how you get the row without retyping it and introducing the drift by hand.
    $lastGroup = ''
    foreach ($m in ($script:Catalogue | Sort-Object Group, Id)) {
        if ($m.Group -ne $lastGroup) {
            if ($lastGroup -ne '') { Write-Host '' }
            Write-Host ("#### {0} ({1})" -f $m.Group, @($script:Catalogue | Where-Object { $_.Group -eq $m.Group }).Count)
            Write-Host ''
            Write-Host '| Metric | Unit | Class | Tol. | What it is, and what it is evidence for |'
            Write-Host '| --- | --- | --- | --- | --- |'
            $lastGroup = $m.Group
        }
        $tol = if ($null -ne $m.Tolerance) { '{0:0.##}' -f [double]$m.Tolerance } else { 'default' }
        if ($m.Direction -in @('pinned', 'mustBeZero', 'coverage', 'noIncrease')) { $tol = 'n/a' }
        $desc = $m.Description.Replace('|', '\|')
        Write-Host ('| `{0}` | {1} | {2} | {3} | {4} |' -f $m.Id, $m.Unit, $m.Direction, $tol, $desc)
    }
    exit 0
}

# ---- everything below reads or writes measurements -------------------------------------
# The gate prints real measured values. Under CI stdout is a build log, and on this repo that
# log is public - so the comparing modes refuse to run there at all rather than relying on
# nobody having wired them up. -AllowCi exists for a private runner and says so out loud.
if (-not $AllowCi -and ($env:CI -or $env:GITHUB_ACTIONS -or $env:TF_BUILD)) {
    Write-Host "::error::measurement-gate.ps1 refuses to run under CI. It prints measurements taken on the maintainer's machine, and a CI log is not a machine-local place to put them. Run it locally before a release. (-AllowCi overrides, for a private runner.)"
    exit 3
}

try {
    if ([string]::IsNullOrWhiteSpace($StoreRoot)) { $StoreRoot = Get-DefaultStoreRoot }
    $script:StoreRootResolved = Assert-StoreOutsideRepository $StoreRoot $RepoRoot
} catch {
    Write-Host "::error::$($_.Exception.Message)"
    exit 3
}
Initialize-Store $script:StoreRootResolved

if ($Show) { exit (Show-History $Last) }
if ($ResetBaseline) { exit (Add-Annotation '*' $Reason) }
if ($Annotate) {
    if ([string]::IsNullOrWhiteSpace($Metric)) { Write-Host "::error::-Annotate needs -Metric <id> (or use -ResetBaseline for all of them)."; exit 3 }
    exit (Add-Annotation $Metric $Reason)
}

if (-not $Run -and -not $Collect) {
    Write-Host "Nothing to do: pass -Run <file> and/or -Collect. See Docs/measurement-gate.md, or -Template for an input skeleton."
    exit 3
}

# ---- assemble this run -----------------------------------------------------------------
$script:RunConditions = @{ profile = 'unknown'; indexed = 'unknown'; corpusId = $null; corpusAnchor = $null; storeSet = $null; notes = $null }
$script:CorpusAgeDays = $null
$incoming = @{}          # id -> @{ Value; Conditions }

if ($Run) {
    if (-not (Test-Path -LiteralPath $Run)) { Write-Host "::error::-Run '$Run' does not exist."; exit 3 }
    $runJson = Get-Content -LiteralPath $Run -Raw | ConvertFrom-Json

    if ($runJson.PSObject.Properties['label'] -and $runJson.label -and -not $Label) { $Label = [string]$runJson.label }
    if ($runJson.PSObject.Properties['conditions'] -and $runJson.conditions) {
        foreach ($p in $runJson.conditions.PSObject.Properties) {
            if ($script:RunConditions.ContainsKey($p.Name) -and $null -ne $p.Value -and "$($p.Value)" -ne '') {
                $script:RunConditions[$p.Name] = $p.Value
            }
        }
    }
    if (-not $runJson.PSObject.Properties['metrics'] -or -not $runJson.metrics) {
        Write-Host "::error::'$Run' carries no `"metrics`" object."; exit 3
    }
    foreach ($p in $runJson.metrics.PSObject.Properties) {
        $v = $p.Value
        if ($v -is [System.Management.Automation.PSCustomObject] -and $v.PSObject.Properties['value']) {
            $incoming[$p.Name] = @{ Value = [double]$v.value; Conditions = $(if ($v.PSObject.Properties['conditions']) { $v.conditions } else { $null }) }
        } else {
            $incoming[$p.Name] = @{ Value = [double]$v; Conditions = $null }
        }
    }
}

# Command-line conditions win over the file, so a run can be re-labelled without editing it.
if ($ProfileKind) { $script:RunConditions.profile = $ProfileKind }
if ($Indexed) { $script:RunConditions.indexed = $Indexed }
if ($PSBoundParameters.ContainsKey('CorpusId')) { $script:RunConditions.corpusId = $CorpusId }
if ($CorpusAnchor) { $script:RunConditions.corpusAnchor = $CorpusAnchor }
if ($StoreSet) { $script:RunConditions.storeSet = $StoreSet }
if ($Notes) { $script:RunConditions.notes = $Notes }

# How far past its own anchor the corpus has drifted, recorded on every run. corpus-build is
# deterministic from (corpusId, seed, anchor) and every date band in it is relative to the
# ANCHOR, not to the clock - which is the point of forbidding a defaulted anchor. So the anchor
# plus the run date is the only thing that says whether the 1/7/30/60-day windows still select
# anything at all, and it is what tells "this got faster" apart from "this measured a corpus
# that no longer selects anything".
if ($script:RunConditions.corpusAnchor) {
    try {
        $anchorDate = [DateTime]::Parse(
            [string]$script:RunConditions.corpusAnchor,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        $script:CorpusAgeDays = [math]::Round(([DateTime]::UtcNow - $anchorDate).TotalDays, 1)
    } catch {
        Write-Host "::error::-CorpusAnchor '$($script:RunConditions.corpusAnchor)' is not a date. It must be the anchor corpus-build was given (yyyy-MM-dd), not the date the corpus was built - the bands are relative to the anchor."
        exit 3
    }
}

$collectorProblems = New-Object System.Collections.ArrayList

if ($Collect) {
    foreach ($m in $script:Catalogue) {
        if (-not $m.Collector) { continue }
        if ($m.Collector.Kind -ne 'SourceConstant') { continue }
        $r = Invoke-SourceConstantCollector $m
        if ($r.Ok) {
            $incoming[$m.Id] = @{ Value = $r.Value; Conditions = $null }
        } else {
            $null = $collectorProblems.Add("$($m.Id): $($r.Detail)")
        }
    }

    $pinned = Invoke-PinnedConstantsCollector
    if ($pinned.Ok) {
        $incoming['invariants.pinnedConstantChecks'] = @{ Value = $pinned.Checks; Conditions = $null }
        if ($pinned.ContainsKey('ThrownTypes')) {
            $incoming['invariants.comHostThrownTypes'] = @{ Value = $pinned.ThrownTypes; Conditions = $null }
            $incoming['invariants.comHostFilesScanned'] = @{ Value = $pinned.FilesScanned; Conditions = $null }
        }
    } else {
        $null = $collectorProblems.Add("invariants.*: $($pinned.Detail)")
    }

    if ($TestLog) {
        $suite = Read-TestLog $TestLog
        $incoming['suite.testsPassed'] = @{ Value = $suite.Passed; Conditions = $null }
        $incoming['suite.testsTotal'] = @{ Value = $suite.Total; Conditions = $null }
        $incoming['suite.testsFailed'] = @{ Value = $suite.Failed; Conditions = $null }
        $incoming['suite.testsSkipped'] = @{ Value = $suite.Skipped; Conditions = $null }
        $incoming['suite.durationMs'] = @{ Value = $suite.DurationMs; Conditions = $null }
    } else {
        $null = $collectorProblems.Add("suite.*: no -TestLog given, so the suite numbers were not collected. Run the standing verification command with its output redirected to a file and pass it here.")
    }
}

# ---- unknown ids ------------------------------------------------------------------------
$unknown = @($incoming.Keys | Where-Object { -not $script:CatalogueById.ContainsKey($_) } | Sort-Object)
if ($unknown.Count -gt 0 -and -not $AllowUnknownMetrics) {
    Write-Host "::error::the run carries metric ids the catalogue does not know: $($unknown -join ', '). A typo'd id would otherwise create a fresh metric that reports `"no baseline`" forever and never fails. Fix the id, add it to the catalogue, or pass -AllowUnknownMetrics."
    exit 3
}

# ---- record ------------------------------------------------------------------------------
$record = New-RunRecord
$metricRecords = New-Object System.Collections.ArrayList
foreach ($id in ($incoming.Keys | Sort-Object)) {
    $cat = $script:CatalogueById[$id]
    $null = $metricRecords.Add([pscustomobject]@{
            id         = $id
            value      = $incoming[$id].Value
            unit       = $(if ($cat) { $cat.Unit } else { '' })
            source     = $(if ($cat) { $cat.Source } else { 'unknown' })
            conditions = $incoming[$id].Conditions
        })
}
$record.metrics = @($metricRecords.ToArray())

$scopeKey = Get-ScopeKey $record.conditions
$history = Read-JsonLines (Get-HistoryPath)
$annotations = Read-JsonLines (Get-AnnotationsPath)

# ---- compare ------------------------------------------------------------------------------
$rows = New-Object System.Collections.ArrayList
$failures = New-Object System.Collections.ArrayList
$coldStarts = New-Object System.Collections.ArrayList
$notices = New-Object System.Collections.ArrayList

# Windows the corpus's date bands have to still cover for a measurement over them to mean
# anything. corpus-build anchors every band on --anchor, so once the run date is more than a
# window past the anchor, that window selects zero items - and a sweep over zero items is fast,
# stable and completely healthy-looking. That is the stable-but-wrong shape, and it is why this
# is a hard failure rather than a notice.
$corpusWindowMetrics = @{
    'sweep.wholeStore7Day.elapsedMs'      = 7
    'sweep.wholeStore7Day.itemsSeen'      = 7
    'sweep.perStore.elapsedMs.max'        = 7
    'sweep.perStore.elapsedMs.total'      = 7
    'sweep.foldersSwept'                  = 7
    'sweep.itemCappedFolders'             = 7
    'sweep.itemsBodyCapped'               = 7
    'scan.wholeStore60Day.elapsedMs'      = 60
    'scan.wholeStore60Day.foldersScanned' = 60
    'scan.wholeStore60Day.itemsPerSecond' = 60
}

foreach ($m in $script:Catalogue) {
    $present = $incoming.ContainsKey($m.Id)
    $baseline = Get-BaselineFor -History $history -MetricId $m.Id -ScopeKey $scopeKey -Annotations $annotations

    if (-not $present) {
        if ($baseline.HasBaseline) {
            # The important one. A run that quietly stops carrying a metric is a gate that
            # quietly stopped covering it, and it would otherwise pass.
            $null = $failures.Add("$($m.Id): MISSING from this run but present in $($baseline.Runs) earlier run(s) in this scope. Coverage cannot be dropped silently - collect it, or retire it from the catalogue on purpose.")
            $null = $rows.Add([pscustomobject]@{ Id = $m.Id; Metric = $m; Base = $baseline.Value; Now = $null; Verdict = 'MISSING'; Detail = 'had history, not collected now'; Baseline = $baseline })
        } elseif ($Require -eq 'All') {
            $null = $failures.Add("$($m.Id): NOT COLLECTED, and -Require All was asked for. This is the pre-release invocation; a partial run must not pass as a full one.")
            $null = $rows.Add([pscustomobject]@{ Id = $m.Id; Metric = $m; Base = $null; Now = $null; Verdict = 'ABSENT'; Detail = 'never collected'; Baseline = $baseline })
        } else {
            $null = $rows.Add([pscustomobject]@{ Id = $m.Id; Metric = $m; Base = $null; Now = $null; Verdict = '-'; Detail = 'not collected in this run'; Baseline = $baseline })
        }
        continue
    }

    $now = [double]$incoming[$m.Id].Value

    # Before comparing: is this number a measurement of anything at all?
    if ($null -ne $script:CorpusAgeDays -and $corpusWindowMetrics.ContainsKey($m.Id)) {
        $window = $corpusWindowMetrics[$m.Id]
        if ($script:CorpusAgeDays -gt $window) {
            $null = $failures.Add("$($m.Id): the corpus anchor is $($script:CorpusAgeDays) days old and this metric is measured over a $window-day window, so the window selects NOTHING. The number is a measurement of an empty selection - fast, stable and meaningless. Re-anchor and rebuild the corpus, or drop this metric from the run.")
        }
    }

    $result = Compare-Metric -Metric $m -Now $now -Baseline $baseline

    switch ($result.Verdict) {
        'FAIL' { $null = $failures.Add("$($m.Id): $($result.Detail)") }
        'NEW' { $null = $coldStarts.Add($m.Id) }
    }
    $null = $rows.Add([pscustomobject]@{
            Id       = $m.Id
            Metric   = $m
            Base     = $baseline.Value
            Now      = $now
            Verdict  = $result.Verdict
            Detail   = $result.Detail
            Baseline = $baseline
        })
}

# ---- suspicious stability -------------------------------------------------------------------
# Tolerances cannot see this shape and never will: a number bounded by another defect does not
# drift, so it passes forever. The 432 KB frame high-water was exactly that - bounded by a
# timeout, read as 152x headroom, stable across every run that measured it. What CAN be
# surfaced mechanically is the fact of the stability, so a reader is pointed at it.
# Restricted to CONTINUOUS metrics - wall clock, byte counts, rates. A count of folders or of
# tests can legitimately be identical run after run and flagging those would bury the signal in
# noise; a millisecond figure that repeats exactly is telling you something.
$continuousUnits = @('ms', 's', 'min', 'bytes', 'items/s')
foreach ($row in $rows) {
    if ($row.Verdict -notin @('OK', 'FAIL')) { continue }
    if ($row.Metric.Direction -ne 'both') { continue }
    if ($row.Metric.Unit -notin $continuousUnits) { continue }
    $b = $row.Baseline
    if ($b.Runs -lt 3) { continue }
    $values = @($b.Values)
    if ($values.Count -lt 3) { continue }
    $distinct = @($values | Sort-Object -Unique)
    if ($distinct.Count -eq 1 -and [double]$distinct[0] -eq [double]$row.Now) {
        $null = $notices.Add("SUSPICIOUSLY STABLE: $($row.Id) has been byte-identical across the last $($b.Runs) runs and again now. Wall-clock measurements do not repeat exactly. Either it is not being measured at all (a cached value, a window that selects nothing, a code path that no longer runs), or it is bounded by something other than the work - which is the shape a tolerance check passes forever.")
    }
}

# ---- condition drift -----------------------------------------------------------------------
$previousInScope = $null
for ($i = $history.Count - 1; $i -ge 0; $i--) {
    if ((Get-ScopeKey $history[$i].conditions) -eq $scopeKey) { $previousInScope = $history[$i]; break }
}
if ($previousInScope) {
    # 'indexed' is decisive rather than advisory: an indexed store takes the index path and an
    # unindexed one takes the seven-day sweep fallback. Comparing those two is comparing
    # different code, and it is exactly the mistake that makes a measurement describe something
    # other than what the reader thinks.
    if ($previousInScope.conditions.indexed -ne $record.conditions.indexed) {
        $null = $failures.Add("conditions: 'indexed' changed '$($previousInScope.conditions.indexed)' -> '$($record.conditions.indexed)'. An indexed store and an unindexed one run different code, so these two runs are not comparable at all. Re-take the run under the previous condition, or reset the baseline with -ResetBaseline and say why.")
    }
    if ($previousInScope.conditions.storeSet -ne $record.conditions.storeSet) {
        $line = "conditions: 'storeSet' changed '$($previousInScope.conditions.storeSet)' -> '$($record.conditions.storeSet)'. Every elapsed figure is a rate over this; read the table with that in mind."
        if ($StrictConditions) { $null = $failures.Add($line) } else { $null = $notices.Add($line) }
    }
    $prevAnchor = $null
    if ($previousInScope.conditions.PSObject.Properties['corpusAnchor']) { $prevAnchor = $previousInScope.conditions.corpusAnchor }
    if ($prevAnchor -ne $record.conditions.corpusAnchor) {
        $null = $notices.Add("conditions: the corpus anchor changed '$prevAnchor' -> '$($record.conditions.corpusAnchor)'. A rebuilt corpus is a different population; the earlier numbers describe the old one.")
    }
}
if ($record.conditions.profile -eq 'unknown') {
    $null = $notices.Add("conditions: 'profile' is 'unknown', so this run is compared only against other 'unknown' runs. Pass -ProfileKind production or -ProfileKind vm.")
}
if ($record.conditions.indexed -eq 'unknown') {
    $null = $notices.Add("conditions: 'indexed' is 'unknown'. Read index.perStore[] out of outlook_health and say which it was - an indexed store and an unindexed one measure different code, and once one run records 'unknown' the next real value looks like a condition change.")
}
if ($record.provenance.gitDirty) {
    $null = $notices.Add("provenance: the working tree was DIRTY, so the commit recorded beside these numbers does not fully describe the code they were taken from.")
}
if ($null -ne $script:CorpusAgeDays) {
    $null = $notices.Add("conditions: the corpus anchor is $($script:CorpusAgeDays) days old. Every date band in a built corpus is relative to that anchor.")
}
foreach ($p in $collectorProblems) { $null = $notices.Add("collector: $p") }

# ---- report ----------------------------------------------------------------------------------
# Everything printed is also kept, so the run's own verdict outlives the terminal scrollback.
$script:ReportLines = New-Object System.Collections.ArrayList
function Say([string] $text = '') {
    Write-Host $text
    $null = $script:ReportLines.Add($text)
}

$verdictOrder = @{ 'FAIL' = 0; 'MISSING' = 1; 'ABSENT' = 2; 'NEW' = 3; 'OK' = 4; '-' = 5 }

if ($ReleaseNoteSummary) {
    # A summary carrying no measurements at all, for the one place the numbers may not go.
    $status = if ($failures.Count -gt 0) { 'FAIL' } elseif ($coldStarts.Count -gt 0) { 'NO BASELINE' } else { 'PASS' }
    $compared = @($rows | Where-Object { $_.Verdict -in @('OK', 'FAIL') }).Count
    Say "Measurement gate: $status - $compared of $($script:Catalogue.Count) metrics compared against this machine's history, $($failures.Count) problem(s). Values are machine-local and deliberately not published; read the full table on the console or in the run report."
} else {
    Say ''
    Say '=================================================================================================='
    Say ' OutlookAI measurement gate'
    Say ' MACHINE-LOCAL NUMBERS. Do not commit them, do not paste them into release notes, do not send them'
    Say ' anywhere. They are statistics about one machine and mean nothing off it.'
    Say '=================================================================================================='
    Say (" store      : {0}" -f (Get-HistoryPath))
    Say (" run        : {0}   {1}" -f $record.runId, $record.recordedAtUtc)
    Say (" commit     : {0} ({1}){2}" -f $record.provenance.gitShort, $record.provenance.gitBranch, $(if ($record.provenance.gitDirty) { ' DIRTY' } else { '' }))
    Say (" conditions : profile={0} indexed={1} corpus={2} anchor={3}{4}" -f `
            $record.conditions.profile, $record.conditions.indexed,
        $(if ($record.conditions.corpusId) { $record.conditions.corpusId } else { '<none>' }),
        $(if ($record.conditions.corpusAnchor) { $record.conditions.corpusAnchor } else { '<none>' }),
        $(if ($null -ne $script:CorpusAgeDays) { " (aged $($script:CorpusAgeDays) d)" } else { '' }))
    Say ("              storeSet={0}" -f $(if ($record.conditions.storeSet) { $record.conditions.storeSet } else { '<unrecorded>' }))
    Say ("              notes={0}" -f $(if ($record.conditions.notes) { $record.conditions.notes } else { '<none>' }))
    Say (" scope      : {0}   ({1} run(s) already in this scope)" -f $scopeKey, @($history | Where-Object { (Get-ScopeKey $_.conditions) -eq $scopeKey }).Count)
    Say (" tolerance  : +/-{0:P1} by default, per-metric overrides in the catalogue; baseline = median of the last {1} run(s)" -f $Tolerance, $BaselineRuns)
    Say (" require    : {0}" -f $Require)
    Say ''

    # The FULL table, every catalogued metric, pass or fail. This is the guarantee that the raw
    # numbers are in front of whoever triggered the release even if nobody runs the reader.
    Say ('{0,-42} {1,13} {2,13} {3,9} {4,5}  {5}' -f 'METRIC', 'BASELINE', 'NOW', 'DELTA', 'RUNS', 'VERDICT')
    Say ('-' * 98)

    foreach ($row in ($rows | Sort-Object @{ Expression = { $verdictOrder[$_.Verdict] } }, Id)) {
        $b = if ($null -eq $row.Base) { '-' } else { '{0:N0}' -f $row.Base }
        $n = if ($null -eq $row.Now) { '-' } else { '{0:N0}' -f $row.Now }
        $d = '-'
        if ($null -ne $row.Base -and $null -ne $row.Now) {
            if ([double]$row.Base -eq 0) {
                $d = if ([double]$row.Now -eq 0) { '0' } else { 'off 0' }
            } else {
                $d = '{0:+0.0%;-0.0%;0.0%}' -f (([double]$row.Now - [double]$row.Base) / [double]$row.Base)
            }
        }
        $runs = if ($row.Baseline.Runs -gt 0) { '{0}' -f $row.Baseline.Runs } else { '-' }
        Say ('{0,-42} {1,13} {2,13} {3,9} {4,5}  {5}' -f $row.Id, $b, $n, $d, $runs, $row.Verdict)
        if ($row.Verdict -notin @('OK', '-')) {
            Say ('{0}{1}' -f (' ' * 4), $row.Detail)
        }
    }

    $notCollected = @($rows | Where-Object { $_.Verdict -eq '-' })
    if ($notCollected.Count -gt 0) {
        Say ''
        Say (" {0} metric(s) above read '-': never collected on this machine, and not collected now." -f $notCollected.Count)
        Say "   Not failures under -Require Present. They ARE gaps. Use -Require All for a release run."
    }

    if ($notices.Count -gt 0) {
        Say ''
        Say ' NOTICES - not failures, and not safe to skip either. This is what a tolerance cannot see.'
        foreach ($n in $notices) { Say "   - $n" }
    }
}

# ---- artifacts -------------------------------------------------------------------------------
# Two files, both under the store root, both machine-local:
#   reports\<runId>.txt              the transcript above, so the verdict outlives the scrollback
#   reports\<runId>.comparison.json  the same comparison in a shape an agent can be handed
$comparisonPath = $null
$reportPath = $null
if (-not $DryRun) {
    $reportsDir = Join-Path $script:StoreRootResolved 'reports'
    $reportPath = Join-Path $reportsDir "$($record.runId).txt"
    Write-StoreFile -Path $reportPath -Content (($script:ReportLines -join [Environment]::NewLine))

    $comparisonRows = New-Object System.Collections.ArrayList
    foreach ($row in ($rows | Sort-Object Id)) {
        $rel = $null
        $abs = $null
        if ($null -ne $row.Base -and $null -ne $row.Now) {
            $abs = [double]$row.Now - [double]$row.Base
            if ([double]$row.Base -ne 0) { $rel = $abs / [double]$row.Base }
        }
        $null = $comparisonRows.Add([pscustomobject]@{
                id              = $row.Id
                group           = $row.Metric.Group
                unit            = $row.Metric.Unit
                description     = $row.Metric.Description
                direction       = $row.Metric.Direction
                tolerance       = (Get-EffectiveTolerance $row.Metric)
                source          = $row.Metric.Source
                currentValue    = $row.Now
                baselineValue   = $row.Base
                baselineRuns    = $row.Baseline.Runs
                baselineSamples = @($row.Baseline.Samples)
                absoluteDelta   = $abs
                relativeDelta   = $rel
                verdict         = $row.Verdict
                detail          = $row.Detail
            })
    }

    $comparison = [pscustomobject]@{
        marker           = $script:RecordMarker
        schema           = $script:SchemaVersion
        kind             = 'comparison'
        runId            = $record.runId
        recordedAtUtc    = $record.recordedAtUtc
        provenance       = $record.provenance
        conditions       = $record.conditions
        scopeKey         = $scopeKey
        defaultTolerance = $Tolerance
        baselineRuns     = $BaselineRuns
        require          = $Require
        verdictCounts    = [pscustomobject]@{
            fail    = $failures.Count
            new     = $coldStarts.Count
            ok      = @($rows | Where-Object { $_.Verdict -eq 'OK' }).Count
            missing = @($rows | Where-Object { $_.Verdict -eq 'MISSING' }).Count
            absent  = @($rows | Where-Object { $_.Verdict -eq 'ABSENT' }).Count
        }
        failures         = @($failures)
        notices          = @($notices)
        metrics          = @($comparisonRows.ToArray())
        readerQuestion   = 'Is there cause for alarm or a course correction here? Movement in either direction counts - a number that got faster usually means something stopped being done. Look past the verdicts for values that are suspiciously STABLE, for stated conditions that do not match what the metric claims to measure, and for a number whose bound is plausibly something other than the work itself. Each metric carries baselineSamples: the individual earlier runs behind the median, each with its own value, commit and conditions.'
    }
    $comparisonPath = Join-Path $reportsDir "$($record.runId).comparison.json"
    Write-StoreFile -Path $comparisonPath -Content ($comparison | ConvertTo-Json -Depth 10)
}

# ---- append -----------------------------------------------------------------------------------
if (-not $DryRun) {
    Write-StoreFile -Path (Get-HistoryPath) -Content ($record | ConvertTo-Json -Depth 8 -Compress) -Append
}

# ---- verdict ------------------------------------------------------------------------------------
Write-Host ''
$exitCode = 0

if ($failures.Count -gt 0) {
    # -ReleaseNoteSummary withholds the detail rather than the verdict, because the detail IS
    # measurements and that mode exists for the one place they may not go.
    Write-Host " GATE FAILED - $($failures.Count) problem(s)$(if ($ReleaseNoteSummary) { '. Re-run without -ReleaseNoteSummary to see which.' } else { ':' })"
    if (-not $ReleaseNoteSummary) {
        foreach ($f in $failures) { Write-Host "   * $f" }
        Write-Host ''
        Write-Host " If one of these is a change you MEANT to make, do not widen the tolerance. Record why:"
        Write-Host "   pwsh -File .github/scripts/measurement-gate.ps1 -Annotate -Metric <id> -Reason 'what moved and what measurement justifies it'"
    }
    $exitCode = 1
} elseif ($coldStarts.Count -gt 0 -and -not $AcceptNewBaseline) {
    Write-Host " NO BASELINE - nothing to compare for $($coldStarts.Count) metric(s)."
    if (-not $ReleaseNoteSummary) { Write-Host "   $($coldStarts -join ', ')" }
    Write-Host ''
    Write-Host " The run has been recorded, so the next one will have something to compare against."
    Write-Host " This is not a pass. Re-run with -AcceptNewBaseline once you have read the values above"
    Write-Host " and believe they describe healthy behaviour - a first measurement can just as easily be"
    Write-Host " a first measurement of something already broken."
    $exitCode = 2
} else {
    if ($coldStarts.Count -gt 0) {
        Write-Host " New baseline accepted for $($coldStarts.Count) metric(s): $($coldStarts -join ', ')"
    }
    Write-Host " GATE PASSED - $(@($rows | Where-Object { $_.Verdict -eq 'OK' }).Count) metric(s) within tolerance."
}

if ($comparisonPath) {
    Write-Host ''
    Write-Host " THE GATE IS ONLY HALF THE CHECK. Tolerances catch drift; they structurally cannot catch a"
    Write-Host " number that never moves BECAUSE something else is broken. Hand the comparison to a reader:"
    Write-Host ''
    Write-Host "   $comparisonPath"
    Write-Host ''
    Write-Host ' Ask it: "Read this measurement comparison. Is there cause for alarm or a course correction?"'
    Write-Host ' The file is machine-local and stays that way - do not attach it to anything public.'
    Write-Host " Transcript of the table above: $reportPath"
}

if ($DryRun) { Write-Host ' (dry run: nothing was appended to the history, and no report was written)' }
exit $exitCode
