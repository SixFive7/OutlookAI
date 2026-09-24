#Requires -Version 5.1
<#
    ============================================================================================
    WRITTEN 2026-09-24 AND NEVER RUN ON A GUEST. WHAT HAS RUN IS -SelfTest, ON THE HOST, UNDER
    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 - AND THE GUEST GUARD'S REFUSAL ON THE HOST.
    ============================================================================================

    Nothing below the guard has executed anywhere: not the profile switches, not the Outlook
    starts, not the one Quit, not a single corpus verb against a store. Every step it drives is a
    step that has run by hand on a guest - Set-DefaultOutlookProfile.ps1 (measured end to end,
    2026-09-24), OUTLOOK.EXE started in session 1, corpus-teardown and corpus-build, and a
    graceful Quit attached through the Running Object Table (the build-out's own scratch script,
    OAI-UNINDEXED 2026-09-24: Outlook left in about 2 s, each time) - except corpus-indexed, written
    the same day, and this script's own Quit. And the population it rebuilds is v2, whose probes and
    build were changed after that guest's build (placement off the default store, visible folders,
    the owner as a resolved recipient): the first run of this script is also their first run.
    Replace this banner with what it did once it has rebuilt a real hub and a live run has read the
    result.

.SYNOPSIS
    Rebuilds the hub's generated population against NOW, before a live run. Guest only, in session
    1. The step a run procedure calls FIRST - decided 2026-09-24 (question D, option (a)).

.DESCRIPTION
    RUN THIS ON THE GUEST, IN SESSION 1, RIGHT AFTER A GUEST RESTART, BEFORE EVERY LIVE RUN:

        .\Register-InteractiveTask.ps1 -TimeoutSeconds 3600 -Script "& 'C:\OutlookAI-Q5\Reset-HubPopulation.ps1' -Execute"

    WHY IT EXISTS. LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier asserts that the
    index frontier is not in the future. A product that read the index's local time as UTC would
    push the frontier forward by the guest's UTC offset - into the future, and so caught - ONLY
    while the real frontier is younger than that offset less the test's five minutes of
    tolerance: 55 minutes on a W. Europe guest in winter, 115 in summer. The hub population's
    newest item is one minute older than the anchor it is built against, so a hub built weeks ago
    makes the test pass whatever the product does. The maintainer decided the hub is rebuilt
    before every run, and that the rebuild is a SCRIPT STEP a run procedure calls, not a paragraph
    somebody has to remember - the pattern chosen for ImportPRF. The other half of that pattern is
    the refusal: the frontier test reads the manifest this script writes
    (hubPopulationManifestPath in the live-test settings) and FAILS on a hub nobody rebuilt,
    naming this script (T2/LiveHubPopulationFreshness.cs).

    EVERYTHING IT NEEDS COMES FROM THE GUEST'S OWN LIVE-TEST SETTINGS. The hub is
    testHubStoreDisplayName; the manifest is hubPopulationManifestPath, whose file name,
    corpus-<id>.jsonl, is the population's id; whether to wait for the index is whether
    indexedStoreDisplayNames names the hub. The seed comes from the manifest's own header line,
    and so does the anchor the old population is torn down by - teardown refuses any other,
    because the anchor is part of the shape key. Nothing is typed on a normal run but -Execute.

    WHAT IT DOES, IN ORDER, with -Execute:

      0. The guest guard (Assert-TestbedGuest), before anything touches the machine.
      1. Reads the settings and the manifest's header, and refuses on anything that does not
         agree: a Production profile, no hubPopulationManifestPath, a manifest of another
         population or another store, a manifest that is not a hub population's.
      2. Refuses unless OUTLOOK.EXE is NOT running - it never closes an Outlook it did not start
         - and unless no Setup key holds ImportPRF (a pending import rebuilds a profile at the
         next start, and this script starts Outlook twice).
      3. Makes the ACCOUNT-LESS profile the default (Set-DefaultOutlookProfile.ps1): the corpus
         tool refuses any profile that holds a mail account, and the tier profile holds the
         dummy account the hub delivers for.
      4. Starts Outlook and waits for it to be warm (180 s by default - the measured bar;
         Build-Corpus.ps1 says why a cold one fails).
      5. corpus-teardown --execute with the manifest's own anchor, and requires "0 remaining".
         Two keys, EntryID allowlist and ordinal tag, as for every deletion in this repository.
      6. Moves the torn-down manifest into hub-history\ beside it - named <id>.<anchor>.jsonl, so
         Testbed/host/Copy-FromGuest.ps1's corpus-*.jsonl never mistakes it for a live one.
      7. corpus-build --execute against NOW, to the second (UTC). The build probes, builds, and
         runs its own census and read-back, and fails on any difference; this script requires
         exit 0, then reads the new manifest's header back and requires the anchor it asked for.
      8. Quits the Outlook it started in step 4 - the only Outlook this script ever quits, under
         mailbox-safety rule 7 (below).
      9. Makes the TIER profile the default again, starts Outlook on it NOT ELEVATED - through
         Start-OutlookUnelevated.ps1, because an elevated Outlook never feeds the Windows Search
         index (measured on OutlookAI-Indexed, 2026-09-24) and the live run attaches at the user's
         own integrity level - and on the indexed guest runs corpus-indexed until the index holds
         every item of the new population (-IndexWaitSeconds, 900 by default). A run started on a
         hub the indexer has half taken in measures half a hub.
     10. Leaves Outlook running on the tier profile, warm, and prints how long the frontier test
         can still catch a local-time misreading - the margin left for the run to reach it - with
         this guest's opt-in value and filter for the run (Testbed/README.md section 4c).

    STAGE BESIDE IT in C:\OutlookAI-Q5: OutlookMapiInterop.ps1 (the guard), Set-DefaultOutlookProfile.ps1
    and Start-OutlookUnelevated.ps1, which it calls, and Register-InteractiveTask.ps1, which runs it.

    Without -Execute: steps 0 to 2 as a report, the population's plan sheet (corpus-plan, pure),
    and the exact corpus-tool command lines -Execute would run. Nothing written, no Outlook.

    THE ONE QUIT, AND RULE 7. A profile switch needs Outlook closed (Set-DefaultOutlookProfile.ps1
    refuses otherwise), and step 8 closes the Outlook step 4 started. Rule 7 allows a graceful
    Application.Quit() only when no compose window is open and the Outbox is empty, with every COM
    reference released first. So it binds, reads the bound profile's name, the open inspectors and
    every store's Outbox - an Outbox proven present by the store's PR_VALID_FOLDER_MASK before it
    is opened, never by the lookup that creates one - releases all of that, and quits only when the
    profile is the account-less one, OUTLOOK.EXE is the very process it started, nothing is open,
    and every Outbox is provably empty. Otherwise it refuses and leaves Outlook running. It then
    releases its own reference and waits for the process to leave. If it does not within
    -QuitTimeoutSeconds, it stops: NEVER taskkill OUTLOOK.EXE - restart the guest and run this
    again with -SkipRebuild, which finishes steps 9 and 10 against the hub already rebuilt.

    WHAT IT NEVER DOES. Kill a process. Quit an Outlook it did not start. Touch a store but the
    hub, or delete anything except through corpus-teardown's two keys. Run a live test. Rebuild the
    bystander or identity populations - no test reads their dates, so they are built once
    (Docs/live-tier-on-the-vm.md section 3b).

    IF IT STOPS HALF WAY, running it again is the recovery, and each stop says so:
      * stopped before the teardown finished: the manifest is where it was; the next run tears
        down again (idempotent: "already gone" is not a failure) and builds.
      * stopped after the teardown, before the build wrote its manifest: the hub is empty and the
        manifest is in hub-history\; the next run takes the seed from the newest one there, says
        so, and builds.
      * stopped during the build: the new manifest records what was built; the next run tears
        that down by ITS anchor and builds again.
      * stopped at the quit: restart the guest, then -SkipRebuild.
    Every stop but the last leaves the ACCOUNT-LESS profile the default. A live run started then
    refuses on its own (its census does not find the tier profile's stores as named), and the
    frontier test refuses a stale hub - but do not rely on that: re-run this script.

    TIMING. Two Outlook starts at 180 s each, a teardown and a build of 68 items with their probes,
    one quit, and the indexer's crawl - budget 15 to 25 minutes, hence -TimeoutSeconds 3600 on the
    interactive task. The frontier test must then run within the margin this script prints at the
    end - 55 minutes after the newest item in a W. Europe winter - or it refuses the run as STALE.

    WHAT IS TESTED WITHOUT A GUEST. -SelfTest exercises every decision here - the settings and
    manifest readers, the plan with each of its refusals, the tool's argument lists, the quit's
    preconditions, the history name, the ImportPRF check and the frontier arithmetic - against
    synthetic inputs. It reads no file, no registry and no process list, and needs no Outlook: it
    is safe anywhere, the maintainer's workstation included. The guard, the registry reads, the
    profile switches, the Outlook starts, the quit and every corpus verb are guest-only.

    Windows PowerShell 5.1: no ternary, no '??', and every native call through
    Invoke-NativeCommand - the interactive task's wrapper redirects this script's output, and
    under 5.1 a redirected native stderr line is a terminating error. Pure ASCII.

.PARAMETER Execute
    Rebuild. Without it, a report of what would run.

.PARAMETER SkipRebuild
    Finish a rebuild that stopped at the quit: switch the default back to the tier profile, start
    Outlook, wait for the index and report the margin. Needs the manifest; tears nothing down.

.PARAMETER Seed
    The population's seed, for the hub's FIRST build only - when there is no manifest and no
    torn-down one in hub-history\. Testbed/testbed.json's corpusIdConvention.populations records
    it (8181 for both guests' hubs). With a manifest present it must agree with the header or the
    run refuses.

.PARAMETER SelfTest
    Run the decision tests and exit. Touches nothing.

.PARAMETER SettingsPath
    The guest's live-test settings file - where Testbed/host/Publish-LiveTierPayload.ps1 expands
    the suite and Testbed/host/New-LiveTestSettings.ps1 prints the copy line to.

.PARAMETER ToolsExe
    The corpus tool, as Testbed/host/Publish-GuestPayload.ps1 stages it. It must be built from a
    commit that knows population v2 and corpus-indexed.

.PARAMETER OutlookExe
    OUTLOOK.EXE. Default: the Office16 Click-to-Run path, 64-bit then 32-bit.

.PARAMETER CorpusProfileName
    The account-less profile the population is built in (Testbed/README.md section 1, step 4e).

.PARAMETER TierProfileName
    The profile the live tier runs in, with the dummy account (step 4c).

.PARAMETER WarmupSeconds
    How long an Outlook this script starts must be up before anything is driven at it. 180 was
    used successfully; 75 was not always enough.

.PARAMETER QuitTimeoutSeconds
    How long to wait for OUTLOOK.EXE to leave after the graceful Quit.

.PARAMETER IndexWaitSeconds
    How long corpus-indexed may wait for the index to hold the whole population (indexed guest).

.PARAMETER LogPath
    Where this run's log goes. Rewritten each run.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.EXAMPLE
    .\Reset-HubPopulation.ps1 -SelfTest
    .\Reset-HubPopulation.ps1
    .\Register-InteractiveTask.ps1 -TimeoutSeconds 3600 -Script "& 'C:\OutlookAI-Q5\Reset-HubPopulation.ps1' -Execute"
    .\Register-InteractiveTask.ps1 -TimeoutSeconds 1800 -Script "& 'C:\OutlookAI-Q5\Reset-HubPopulation.ps1' -Execute -SkipRebuild"
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')] [switch] $Execute,
    [Parameter(ParameterSetName = 'Run')] [switch] $SkipRebuild,
    [Parameter(ParameterSetName = 'Run')] [string] $Seed,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest,
    [string] $SettingsPath = 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json',
    [string] $ToolsExe = 'C:\OutlookAI-Q5\tools\OutlookAI.RemediationTools.exe',
    [string] $OutlookExe,
    [string] $CorpusProfileName = 'CorpusProfile',
    [string] $TierProfileName = 'OutlookAI-Tier',
    [int] $WarmupSeconds = 180,
    [int] $QuitTimeoutSeconds = 120,
    [int] $IndexWaitSeconds = 900,
    [string] $LogPath = 'C:\OutlookAI-Q5\reset-hub-population.log',
    [string[]] $ExpectedUser = @('vmadmin')
)

$ErrorActionPreference = 'Stop'

$script:Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$script:UtcFormat = "yyyy-MM-dd'T'HH:mm:ss'Z'"

# A manifest is corpus-<id>.jsonl (Testbed/testbed.json, corpusIdConvention.manifestFileName).
$script:ManifestLeafPattern = '(?i)^corpus-(?<id>[A-Za-z0-9_-]+)\.jsonl$'

# Where a torn-down manifest goes. Its files are named <id>.<anchor>.jsonl - never corpus-*.jsonl,
# which Copy-FromGuest.ps1 collects as a live manifest.
$script:HistoryDirectoryName = 'hub-history'

# What a hub population's shape key carries (CorpusPopulation.ShapeKeySuffixFor): '|p:hub:v<N>|o:...'.
$script:HubShapeMarker = '|p:hub:v'

# The only manifest format there is (CorpusManifest.CurrentVersion).
$script:ManifestVersion = 1

# LiveHubPopulationFreshness.FrontierFutureTolerance - the frontier test's own five minutes.
$script:FrontierToleranceMinutes = 5

# PR_VALID_FOLDER_MASK and its Outbox bit (Microsoft's MAPIDefS.h, FOLDER_IPM_OUTBOX_VALID), as
# OutlookAI.Core's SpecialFolders reads them: an Outbox is proven present before it is opened.
$script:ValidFolderMaskSchema = 'http://schemas.microsoft.com/mapi/proptag/0x35DF0003'
$script:FolderIpmOutboxValid = 0x04
$script:OlFolderOutbox = 4

# The Office majors this product supports, in its probe order (OfficeVersions.Supported).
$script:SupportedOfficeVersions = @('16.0', '17.0', '15.0')

# =============================================================================================
# PURE DECISIONS. No file, no registry, no process, no Outlook, no output: -SelfTest decides
# everything below exactly as a run does.
# =============================================================================================

<#
    An anchor as the manifest carries it, yyyy-MM-ddTHH:mm:ssZ, or $null. Accepts that string, the
    date-only form, and the DateTime PowerShell 7's ConvertFrom-Json makes of an ISO string.
#>
function ConvertTo-UtcText {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Utc) { return $Value.ToString($script:UtcFormat, $script:Invariant) }
        if ($Value.Kind -eq [System.DateTimeKind]::Local) { return $Value.ToUniversalTime().ToString($script:UtcFormat, $script:Invariant) }
        return $null
    }
    if (-not ($Value -is [string])) { return $null }
    if ($Value -cnotmatch '^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}Z)?$') { return $null }
    $parsed = [datetime]::MinValue
    $formats = [string[]]@('yyyy-MM-dd', $script:UtcFormat)
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    if (-not [datetime]::TryParseExact($Value, $formats, $script:Invariant, $styles, [ref]$parsed)) { return $null }
    return $parsed.ToString($script:UtcFormat, $script:Invariant)
}

function ConvertFrom-UtcText {
    param([string] $Text)
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    if (-not [datetime]::TryParseExact($Text, $script:UtcFormat, $script:Invariant, $styles, [ref]$parsed)) { return $null }
    return [datetime]::SpecifyKind($parsed, [System.DateTimeKind]::Utc)
}

<# The population id a manifest path names - corpus-<id>.jsonl on an absolute path - or $null. #>
function Get-HubPopulationId {
    param([string] $ManifestPath)
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { return $null }
    if ($ManifestPath -cnotmatch '^[A-Za-z]:\\') { return $null }
    $leaf = $ManifestPath.Substring($ManifestPath.LastIndexOf('\') + 1)
    $match = [regex]::Match($leaf, $script:ManifestLeafPattern)
    if (-not $match.Success) { return $null }
    return $match.Groups['id'].Value
}

function Get-JsonField {
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

<#
    What the guest's live-test settings say about the hub - read the way the loader reads them,
    names matched ignoring case. Refuses a Production profile: that is the maintainer's machine,
    whose hub is real mail, and nothing here is for it.
#>
function Read-HubSettingsFact {
    param($Settings)
    $problems = New-Object System.Collections.Generic.List[string]
    $hub = $null
    $manifestPath = $null
    $populationId = $null
    $indexed = $false

    if ($null -eq $Settings -or -not ($Settings -is [System.Management.Automation.PSCustomObject])) {
        $problems.Add('the live-test settings are not a JSON object.')
        return [pscustomobject]@{ Hub = $null; ManifestPath = $null; PopulationId = $null; Indexed = $false; Problems = $problems.ToArray() }
    }

    $profileValue = Get-JsonField $Settings 'machineProfile'
    if (-not ($profileValue -is [string]) -or $profileValue -ne 'Portable') {
        $problems.Add("machineProfile is '$profileValue', not 'Portable'. This script rebuilds a GENERATED hub on a test guest; a Production profile is the maintainer's own machine, whose hub is real mail.")
    }

    $hubValue = Get-JsonField $Settings 'testHubStoreDisplayName'
    if ($hubValue -is [string] -and -not [string]::IsNullOrWhiteSpace($hubValue)) { $hub = $hubValue }
    else { $problems.Add('testHubStoreDisplayName is missing or blank - there is no hub to rebuild.') }

    $manifestValue = Get-JsonField $Settings 'hubPopulationManifestPath'
    if (-not ($manifestValue -is [string]) -or [string]::IsNullOrWhiteSpace($manifestValue)) {
        $problems.Add('hubPopulationManifestPath is missing. It names the hub population''s manifest, which is what this script rebuilds from and the frontier test reads; render the settings again with Testbed/host/New-LiveTestSettings.ps1, which requires it.')
    }
    else {
        $manifestPath = $manifestValue
        $populationId = Get-HubPopulationId $manifestValue
        if ($null -eq $populationId) {
            $problems.Add("hubPopulationManifestPath '$manifestValue' is not an absolute path ending corpus-<population id>.jsonl, so it names no population.")
        }
    }

    $indexedValue = Get-JsonField $Settings 'indexedStoreDisplayNames'
    if ($null -ne $hub -and $null -ne $indexedValue) {
        foreach ($store in @($indexedValue)) {
            if ($store -is [string] -and [string]::Equals($store, $hub, [System.StringComparison]::OrdinalIgnoreCase)) { $indexed = $true }
        }
    }

    return [pscustomobject]@{ Hub = $hub; ManifestPath = $manifestPath; PopulationId = $populationId; Indexed = $indexed; Problems = $problems.ToArray() }
}

<# What a manifest's header line says, with every reason it is not a hub population's. #>
function Read-HubManifestHeader {
    param([string] $HeaderLine)
    $problems = New-Object System.Collections.Generic.List[string]
    $parsed = $null
    if ([string]::IsNullOrWhiteSpace($HeaderLine)) {
        $problems.Add('it has no header line.')
    }
    else {
        try { $parsed = ConvertFrom-Json -InputObject $HeaderLine -ErrorAction Stop }
        catch { $problems.Add('its header line is not JSON.') }
    }

    $corpusId = $null
    $seedValue = $null
    $anchor = $null
    $store = $null
    if ($null -ne $parsed) {
        $version = Get-JsonField $parsed 'Version'
        if (-not (($version -is [int] -or $version -is [long]) -and [long]$version -eq $script:ManifestVersion)) {
            $problems.Add("its format version is '$version', where this script reads version $($script:ManifestVersion).")
        }
        $corpusId = Get-JsonField $parsed 'CorpusId'
        if (-not ($corpusId -is [string]) -or $corpusId.Length -eq 0) { $problems.Add('it names no corpus id.'); $corpusId = $null }
        $rawSeed = Get-JsonField $parsed 'Seed'
        if ($rawSeed -is [int] -or $rawSeed -is [long]) { $seedValue = [long]$rawSeed }
        else { $problems.Add("its seed '$rawSeed' is not an integer.") }
        # Read off the LINE, not the parsed object: PowerShell 7's ConvertFrom-Json turns an ISO
        # string into a DateTime and converts any offset away, so '...+02:00' would pass there as a
        # UTC instant and fail under 5.1. The raw text is the same in both shells.
        $rawAnchor = [regex]::Match($HeaderLine, '"AnchorUtc"\s*:\s*"(?<a>[^"]*)"')
        if ($rawAnchor.Success) { $anchor = ConvertTo-UtcText $rawAnchor.Groups['a'].Value }
        if ($null -eq $anchor) { $problems.Add('its anchor is not a UTC instant.') }
        $shapeKey = Get-JsonField $parsed 'ShapeKey'
        if (-not ($shapeKey -is [string]) -or $shapeKey.IndexOf($script:HubShapeMarker, [System.StringComparison]::Ordinal) -lt 0) {
            $problems.Add('its shape key is not a HUB population''s - another population''s manifest, or the measurement corpus''s.')
        }
        $store = Get-JsonField $parsed 'StoreDisplayName'
        if (-not ($store -is [string]) -or $store.Length -eq 0) { $problems.Add('it names no store.'); $store = $null }
    }

    return [pscustomobject]@{ CorpusId = $corpusId; Seed = $seedValue; AnchorUtc = $anchor; Store = $store; Problems = $problems.ToArray() }
}

<#
    The newest torn-down manifest of population $CorpusId in hub-history\ - what the next run
    rebuilds from when a run stopped between its teardown and its build. Entries are
    @{ Name; HeaderLine }. $null when there is none.
#>
function Select-NewestHistorySeed {
    param([string] $CorpusId, [object[]] $Entries)
    $best = $null
    foreach ($entry in @($Entries)) {
        if ($null -eq $entry -or -not ($entry.Name -is [string])) { continue }
        if (-not $entry.Name.StartsWith($CorpusId + '.', [System.StringComparison]::Ordinal)) { continue }
        $header = Read-HubManifestHeader $entry.HeaderLine
        if ($header.Problems.Count -gt 0 -or $header.CorpusId -cne $CorpusId) { continue }
        if ($null -eq $best -or [string]::CompareOrdinal($header.AnchorUtc, $best.AnchorUtc) -gt 0) {
            $best = [pscustomobject]@{ Name = $entry.Name; Seed = $header.Seed; AnchorUtc = $header.AnchorUtc }
        }
    }
    return $best
}

<# The file name a torn-down manifest is kept under: <id>.<anchor, compact>.jsonl. #>
function Format-HubHistoryName {
    param([string] $CorpusId, [string] $AnchorText)
    return $CorpusId + '.' + $AnchorText.Replace('-', '').Replace(':', '') + '.jsonl'
}

<#
    The whole plan: what to tear down and by which anchor, which seed to build with, and the
    anchor to build against - or every reason not to.
#>
function Resolve-HubResetPlan {
    param(
        $SettingsFact,
        [bool] $ManifestExists,
        $Header,
        $HistorySeed,
        [string] $SeedText,
        [bool] $SkipRebuild,
        [datetime] $NowUtc
    )

    $problems = New-Object System.Collections.Generic.List[string]
    $notes = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($SettingsFact.Problems)) { $problems.Add("settings: $p") }

    $seedGiven = $null
    if (-not [string]::IsNullOrEmpty($SeedText)) {
        if ($SeedText -cmatch '^-?\d{1,18}$') { $seedGiven = [long]::Parse($SeedText, $script:Invariant) }
        else { $problems.Add("-Seed '$SeedText' is not an integer.") }
    }

    $seedValue = $null
    $teardownAnchor = $null
    $manifestPath = $SettingsFact.ManifestPath
    $corpusId = $SettingsFact.PopulationId
    if ($problems.Count -eq 0) {
        if ($ManifestExists) {
            if ($null -eq $Header -or @($Header.Problems).Count -gt 0) {
                foreach ($p in @($Header.Problems)) { $problems.Add("the manifest at ${manifestPath}: $p") }
            }
            else {
                if ($Header.CorpusId -cne $corpusId) {
                    $problems.Add("the manifest at $manifestPath records population '$($Header.CorpusId)', but its name says '$corpusId'. A manifest is named after its population; one of the two is wrong, and teardown by the wrong one deletes nothing or the wrong thing.")
                }
                if (-not [string]::Equals($Header.Store, $SettingsFact.Hub, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $problems.Add("the manifest at $manifestPath records store '$($Header.Store)', but the settings' hub is '$($SettingsFact.Hub)'. It is not this hub's population.")
                }
                if ($null -ne $seedGiven -and $seedGiven -ne $Header.Seed) {
                    $problems.Add("-Seed $seedGiven disagrees with the manifest, which records $($Header.Seed). Leave -Seed out: the manifest's own seed is the one its population was built with.")
                }
                $seedValue = $Header.Seed
                $teardownAnchor = $Header.AnchorUtc
            }
        }
        elseif ($SkipRebuild) {
            $problems.Add("-SkipRebuild finishes a rebuild whose manifest is there, and there is none at $manifestPath. Run without -SkipRebuild.")
        }
        elseif ($null -ne $seedGiven) {
            $seedValue = $seedGiven
            $notes.Add("no manifest at $manifestPath - building population '$corpusId' with -Seed $seedGiven. If this is the hub's FIRST build, Docs/live-tier-on-the-vm.md section 3b's first-build steps come first: the hub attached to the account-less profile while it is still empty.")
        }
        elseif ($null -ne $HistorySeed) {
            $seedValue = $HistorySeed.Seed
            $notes.Add("no manifest at $manifestPath - the last rebuild stopped after its teardown and before its build wrote one. Rebuilding with seed $($HistorySeed.Seed), from the newest torn-down manifest, $($script:HistoryDirectoryName)\$($HistorySeed.Name).")
        }
        else {
            $problems.Add("no manifest at $manifestPath, and no torn-down one in $($script:HistoryDirectoryName)\ beside it: this is the hub population's FIRST build. Pass -Seed <the seed Testbed/testbed.json's corpusIdConvention.populations assigns '$corpusId'>, after Docs/live-tier-on-the-vm.md section 3b's first-build steps.")
        }
    }

    $whole = [datetime]::SpecifyKind($NowUtc, [System.DateTimeKind]::Utc)
    $whole = $whole.AddTicks(-($whole.Ticks % [TimeSpan]::TicksPerSecond))
    return [pscustomobject]@{
        Proceed        = ($problems.Count -eq 0)
        Problems       = $problems.ToArray()
        Notes          = $notes.ToArray()
        CorpusId       = $corpusId
        Store          = $SettingsFact.Hub
        ManifestPath   = $manifestPath
        Seed           = $seedValue
        TeardownAnchor = $teardownAnchor
        NewAnchor      = $whole.ToString($script:UtcFormat, $script:Invariant)
        Rebuild        = (-not $SkipRebuild)
        Indexed        = $SettingsFact.Indexed
    }
}

<#
    One corpus verb's arguments for this population. NEVER the write flag: the call site that
    means it appends it, where it can be read.
#>
function Get-HubToolArgument {
    param([string] $Verb, [string] $Store, [string] $CorpusId, [long] $SeedValue, [string] $Anchor, [string] $ManifestPath, [int] $WaitSeconds = -1)
    $arguments = @($Verb, '--population', 'hub', '--store', $Store)
    if ($Verb -ceq 'corpus-teardown' -or $Verb -ceq 'corpus-build') { $arguments += @('--allow-store', $Store) }
    $arguments += @('--corpus-id', $CorpusId, '--seed', $SeedValue.ToString($script:Invariant), '--anchor', $Anchor)
    if ($Verb -cne 'corpus-plan') { $arguments += @('--manifest', $ManifestPath) }
    if ($WaitSeconds -ge 0) { $arguments += @('--wait-seconds', $WaitSeconds.ToString($script:Invariant)) }
    return , $arguments
}

<#
    Whether the Outlook this script started may be quit - mailbox-safety rule 7. $Outboxes are
    @{ Store; Count }: Count -1 for a store with no Outbox (its mask says so), $null for one whose
    Outbox could not be proven either way.
#>
function Test-QuitPrecondition {
    param(
        [string] $ProfileName,
        [string] $ExpectedProfile,
        [int[]] $RunningPids,
        [int] $StartedPid,
        $InspectorCount,
        [object[]] $Outboxes
    )
    $problems = New-Object System.Collections.Generic.List[string]
    if (-not [string]::Equals($ProfileName, $ExpectedProfile, [System.StringComparison]::OrdinalIgnoreCase)) {
        $problems.Add("the running Outlook is on profile '$ProfileName', not '$ExpectedProfile' - it is not the Outlook this script started, and this script quits no other.")
    }
    $pids = @($RunningPids)
    if ($pids.Count -ne 1 -or $pids[0] -ne $StartedPid) {
        $problems.Add("OUTLOOK.EXE is running as pid(s) $($pids -join ', '), where this script started pid $StartedPid alone.")
    }
    if ($null -eq $InspectorCount) { $problems.Add('the number of open item windows could not be read, so no compose window can be ruled out.') }
    elseif ([int]$InspectorCount -gt 0) { $problems.Add("$InspectorCount item window(s) are open - a compose window among them would be lost or queued by a Quit.") }
    foreach ($outbox in @($Outboxes)) {
        if ($null -eq $outbox) { continue }
        if ($null -eq $outbox.Count) { $problems.Add("the Outbox of '$($outbox.Store)' could not be proven empty.") }
        elseif ([int]$outbox.Count -gt 0) { $problems.Add("the Outbox of '$($outbox.Store)' holds $($outbox.Count) item(s).") }
    }
    return , $problems.ToArray()
}

<#
    Whether any Setup key holds ImportPRF. $Readings are @{ Path; Value }: '' absent, $null
    unreadable. FAIL-CLOSED, as Build-Corpus.ps1's precondition 3: a pending import rebuilds a
    profile at the next Outlook start, and this script starts Outlook twice.
#>
function Test-ImportPrfReading {
    param([object[]] $Readings)
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($reading in @($Readings)) {
        if ($null -eq $reading) { continue }
        if ($null -eq $reading.Value) { $problems.Add("ImportPRF could not be read from $($reading.Path), so nothing proves no Outlook start is waiting to rebuild a profile.") }
        elseif ([string]$reading.Value -ne '') { $problems.Add("ImportPRF is set under $($reading.Path): '$($reading.Value)'. The next Outlook start would re-import a profile - and a re-import REBUILDS it. Remove it first: .\New-OutlookProfile.ps1 -ClearImportPrf -Execute.") }
    }
    return , $problems.ToArray()
}

<#
    How long the frontier test can still catch a local-time misreading: the hub's newest item must
    be younger than the UTC offset less the test's tolerance when the test runs.
#>
function Get-FrontierWindow {
    param([datetime] $NewestUtc, [timespan] $UtcOffset, [datetime] $NowUtc)
    $margin = $UtcOffset.Duration() - [timespan]::FromMinutes($script:FrontierToleranceMinutes)
    $deadline = $NewestUtc + $margin
    return [pscustomobject]@{
        Discriminates = ($margin -gt [timespan]::Zero)
        MarginMinutes = [int][math]::Round($margin.TotalMinutes)
        DeadlineUtc   = $deadline
        MinutesLeft   = [int][math]::Floor(($deadline - $NowUtc).TotalMinutes)
    }
}

<# The newest received instant off a plan sheet's "received range : A .. B" line, or $null. #>
function Get-NewestFromPlanText {
    param([string] $Text)
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $match = [regex]::Match($Text, 'received range\s*:\s*(?<oldest>\S+)\s*\.\.\s*(?<newest>\S+)')
    if (-not $match.Success) { return $null }
    return ConvertTo-UtcText $match.Groups['newest'].Value
}

<# What to do when Outlook is up where it must not be, or would not leave. #>
function Format-RestartAdvice {
    param([string] $Then)
    return "Restart the guest - from the host, over PowerShell Direct, which lands in session 0: Invoke-Command -VMName <this guest> -Credential <vmadmin> { shutdown /r /t 0 } (Testbed/README.md section 1, step 4d). NEVER taskkill OUTLOOK.EXE: mailbox-safety rule 7 forbids it outright. Once autologon has brought session 1 back, $Then"
}

# =============================================================================================
# SELF-TEST. Pure: no file, no registry, no process list, no tool, no Outlook. Runs anywhere.
# =============================================================================================

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    # Compared as text with -ceq, searched with String.Contains - never -like, whose brackets are
    # wildcards (Testbed/README.md section 4b).
    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:SelfTestChecks++
        $expectedText = '<null>'
        if ($null -ne $Expected) { $expectedText = (@($Expected) | ForEach-Object { "$_" }) -join '|' }
        $actualText = '<null>'
        if ($null -ne $Actual) { $actualText = (@($Actual) | ForEach-Object { "$_" }) -join '|' }
        if ($expectedText -ceq $actualText) { Write-Host ("  OK   {0}" -f $What) }
        else {
            $script:SelfTestFailures += "$What : expected $expectedText, got $actualText"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $expectedText, $actualText)
        }
    }

    function Test-Says {
        param([string] $What, [string[]] $Lines, [string] $Needle)
        $hit = $false
        foreach ($line in @($Lines)) { if ($null -ne $line -and $line.Contains($Needle)) { $hit = $true } }
        if (-not $hit) { Write-Host ("       (lines were: {0})" -f (@($Lines) -join ' | ')) }
        Test-Case $What $true $hit
    }

    function New-Settings {
        param([string] $MachineProfile = 'Portable', [string] $Hub = 'tier@vm.invalid', $Manifest = 'C:\OutlookAI-Q5\corpus-hub-indexed.jsonl', [string[]] $Indexed = @('tier@vm.invalid', 'bystander@vm.invalid', 'Corpus A'))
        $settings = [pscustomobject]@{ machineProfile = $MachineProfile; testHubStoreDisplayName = $Hub; indexedStoreDisplayNames = $Indexed }
        if ($null -ne $Manifest) { $settings | Add-Member -NotePropertyName 'hubPopulationManifestPath' -NotePropertyValue $Manifest }
        return $settings
    }

    function New-HeaderLine {
        param([int] $Version = 1, [string] $CorpusId = 'hub-indexed', [string] $SeedJson = '8181', [string] $Anchor = '2026-09-24T08:00:00Z', [string] $ShapeKey = 'v1|hub-indexed|8181|2026-09-24T08:00:00Z|s:x|p:hub:v2|o:Tier <tier@vm.invalid>', [string] $Store = 'tier@vm.invalid')
        return '{"Version":' + $Version + ',"CorpusId":"' + $CorpusId + '","Seed":' + $SeedJson + ',"AnchorUtc":"' + $Anchor + '","ShapeKey":"' + $ShapeKey + '","StoreDisplayName":"' + $Store + '","StoreFilePath":"C:\\OutlookAI-Tier\\Outlook.pst","DateWriteMethod":"PropertyAccessorDates","PlacementMethod":"DraftsThenMoveWithSentFlag"}'
    }

    $now = [datetime]::SpecifyKind([datetime]'2026-09-24T09:15:42.678', [System.DateTimeKind]::Utc)

    Write-Host 'Reset-HubPopulation self-test. No file, no registry, no process list, no tool, no Outlook.'
    Write-Host ''
    Write-Host '== anchors =='
    Test-Case 'a manifest anchor reads back as itself' '2026-09-24T08:00:00Z' (ConvertTo-UtcText '2026-09-24T08:00:00Z')
    Test-Case 'a date-only anchor is midnight UTC' '2026-08-19T00:00:00Z' (ConvertTo-UtcText '2026-08-19')
    Test-Case 'a UTC DateTime - PowerShell 7''s reading of the same text - comes back exactly' '2026-09-24T08:00:00Z' (ConvertTo-UtcText ([datetime]::SpecifyKind([datetime]'2026-09-24T08:00:00', [System.DateTimeKind]::Utc)))
    Test-Case 'a local DateTime is converted, not reinterpreted' ([datetime]::SpecifyKind([datetime]'2026-09-24T08:00:00', [System.DateTimeKind]::Utc).ToLocalTime().ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")) (ConvertTo-UtcText ([datetime]::SpecifyKind([datetime]'2026-09-24T08:00:00', [System.DateTimeKind]::Utc).ToLocalTime()))
    Test-Case 'an offset is not a UTC instant' '<null>' (ConvertTo-UtcText '2026-09-24T08:00:00+02:00')
    Test-Case 'nor is a number' '<null>' (ConvertTo-UtcText 42)

    Write-Host ''
    Write-Host '== the population id, off the manifest''s name =='
    Test-Case 'corpus-<id>.jsonl names the id' 'hub-indexed' (Get-HubPopulationId 'C:\OutlookAI-Q5\corpus-hub-indexed.jsonl')
    Test-Case 'in any case' 'hub-indexed' (Get-HubPopulationId 'C:\OutlookAI-Q5\CORPUS-hub-indexed.JSONL')
    Test-Case 'a relative path names nothing' '<null>' (Get-HubPopulationId 'corpus-hub-indexed.jsonl')
    Test-Case 'nor another file name' '<null>' (Get-HubPopulationId 'C:\OutlookAI-Q5\hub-indexed.jsonl')
    Test-Case 'nor an id with a space' '<null>' (Get-HubPopulationId 'C:\OutlookAI-Q5\corpus-hub indexed.jsonl')

    Write-Host ''
    Write-Host '== the settings =='
    $fact = Read-HubSettingsFact (New-Settings)
    Test-Case 'a test guest''s settings read cleanly' 0 $fact.Problems.Count
    Test-Case 'the hub, the manifest and the id' 'tier@vm.invalid|C:\OutlookAI-Q5\corpus-hub-indexed.jsonl|hub-indexed' @($fact.Hub, $fact.ManifestPath, $fact.PopulationId)
    Test-Case 'the hub in the indexed list means: wait for the index' $true $fact.Indexed
    Test-Case 'an empty indexed list - the unindexed guest - means: do not' $false (Read-HubSettingsFact (New-Settings -Indexed @())).Indexed
    Test-Case 'and so does an indexed list without the hub' $false (Read-HubSettingsFact (New-Settings -Indexed @('Corpus A'))).Indexed
    Test-Says 'a Production profile is refused' (Read-HubSettingsFact (New-Settings -MachineProfile 'Production')).Problems 'whose hub is real mail'
    Test-Says 'settings without the manifest are refused, naming the renderer' (Read-HubSettingsFact (New-Settings -Manifest $null)).Problems 'New-LiveTestSettings.ps1'
    Test-Says 'a manifest path naming no population is refused' (Read-HubSettingsFact (New-Settings -Manifest 'C:\OutlookAI-Q5\hub.jsonl')).Problems 'names no population'
    Test-Says 'no hub is refused' (Read-HubSettingsFact (New-Settings -Hub ' ')).Problems 'no hub to rebuild'
    Test-Says 'a document that is not an object is refused' (Read-HubSettingsFact 'text').Problems 'not a JSON object'

    Write-Host ''
    Write-Host '== the manifest header =='
    $header = Read-HubManifestHeader (New-HeaderLine)
    Test-Case 'a hub population''s header reads cleanly' 0 $header.Problems.Count
    Test-Case 'id, seed, anchor and store' 'hub-indexed|8181|2026-09-24T08:00:00Z|tier@vm.invalid' @($header.CorpusId, $header.Seed, $header.AnchorUtc, $header.Store)
    Test-Says 'the bystander''s manifest is not a hub''s' (Read-HubManifestHeader (New-HeaderLine -ShapeKey 'v1|b|1|x|p:bystander:v2|o:B <b@vm.invalid>')).Problems 'not a HUB population'
    Test-Says 'nor is the measurement corpus''s' (Read-HubManifestHeader (New-HeaderLine -ShapeKey 'v1|vm2|7777|2026-08-19T00:00:00Z|s:short:45:200:1200')).Problems 'not a HUB population'
    Test-Says 'another format version is refused' (Read-HubManifestHeader (New-HeaderLine -Version 2)).Problems 'format version'
    Test-Says 'a seed that is not an integer is refused' (Read-HubManifestHeader (New-HeaderLine -SeedJson '"8181"')).Problems 'is not an integer'
    Test-Says 'an anchor that is not UTC is refused' (Read-HubManifestHeader (New-HeaderLine -Anchor '2026-09-24T10:00:00+02:00')).Problems 'not a UTC instant'
    Test-Says 'a line that is not JSON is refused' (Read-HubManifestHeader '{ torn').Problems 'not JSON'
    Test-Says 'an empty file is refused' (Read-HubManifestHeader '').Problems 'no header line'

    Write-Host ''
    Write-Host '== the plan =='
    $plan = Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header $header -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now
    Test-Case 'a manifest in place: proceed' $true $plan.Proceed
    Test-Case 'tear down by the anchor the manifest records' '2026-09-24T08:00:00Z' $plan.TeardownAnchor
    Test-Case 'with the manifest''s own seed' 8181 $plan.Seed
    Test-Case 'and build against now, to the second' '2026-09-24T09:15:42Z' $plan.NewAnchor
    Test-Case 'into the settings'' hub' 'tier@vm.invalid|hub-indexed' @($plan.Store, $plan.CorpusId)
    Test-Case '-Seed agreeing with the manifest is accepted' $true (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header $header -HistorySeed $null -SeedText '8181' -SkipRebuild $false -NowUtc $now).Proceed
    Test-Says '-Seed disagreeing with it is refused' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header $header -HistorySeed $null -SeedText '7' -SkipRebuild $false -NowUtc $now).Problems 'disagrees with the manifest'
    Test-Says '-Seed that is not a number is refused' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header $header -HistorySeed $null -SeedText 'x1' -SkipRebuild $false -NowUtc $now).Problems 'is not an integer'
    Test-Says 'a manifest of another population is refused' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header (Read-HubManifestHeader (New-HeaderLine -CorpusId 'hub-unindexed')) -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now).Problems "records population 'hub-unindexed'"
    Test-Says 'a manifest of another store is refused' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header (Read-HubManifestHeader (New-HeaderLine -Store 'other@vm.invalid')) -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now).Problems "records store 'other@vm.invalid'"
    Test-Case 'the store is matched ignoring case, as the tool matches it' $true (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header (Read-HubManifestHeader (New-HeaderLine -Store 'TIER@vm.invalid')) -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now).Proceed
    Test-Says 'a broken manifest is refused with its reason' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header (Read-HubManifestHeader '{ torn') -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now).Problems 'not JSON'
    $first = Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $false -Header $null -HistorySeed $null -SeedText '8181' -SkipRebuild $false -NowUtc $now
    Test-Case 'no manifest and -Seed: a first build' 'True|8181|<null>' @($first.Proceed, $first.Seed, $(if ($null -eq $first.TeardownAnchor) { '<null>' } else { $first.TeardownAnchor }))
    Test-Says 'and it says the first-build steps come first' $first.Notes 'section 3b'
    $resume = Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $false -Header $null -HistorySeed ([pscustomobject]@{ Name = 'hub-indexed.20260923T070000Z.jsonl'; Seed = [long]8181; AnchorUtc = '2026-09-23T07:00:00Z' }) -SeedText '' -SkipRebuild $false -NowUtc $now
    Test-Case 'no manifest, a torn-down one in history: rebuild with its seed' 'True|8181' @($resume.Proceed, $resume.Seed)
    Test-Says 'and say where the seed came from' $resume.Notes 'hub-indexed.20260923T070000Z.jsonl'
    Test-Says 'no manifest, no history, no -Seed: refused as a first build' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $false -Header $null -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now).Problems 'FIRST build'
    $finish = Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $true -Header $header -HistorySeed $null -SeedText '' -SkipRebuild $true -NowUtc $now
    Test-Case '-SkipRebuild with a manifest: finish, rebuild nothing' 'True|False' @($finish.Proceed, $finish.Rebuild)
    Test-Says '-SkipRebuild with no manifest is refused' (Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $false -Header $null -HistorySeed $null -SeedText '8181' -SkipRebuild $true -NowUtc $now).Problems '-SkipRebuild finishes'
    Test-Says 'settings problems stop the plan, labelled' (Resolve-HubResetPlan -SettingsFact (Read-HubSettingsFact (New-Settings -MachineProfile 'Production')) -ManifestExists $true -Header $header -HistorySeed $null -SeedText '' -SkipRebuild $false -NowUtc $now).Problems 'settings: machineProfile'

    Write-Host ''
    Write-Host '== history =='
    $entries = @(
        [pscustomobject]@{ Name = 'hub-indexed.20260922T070000Z.jsonl'; HeaderLine = (New-HeaderLine -Anchor '2026-09-22T07:00:00Z' -SeedJson '8181') }
        [pscustomobject]@{ Name = 'hub-indexed.20260923T070000Z.jsonl'; HeaderLine = (New-HeaderLine -Anchor '2026-09-23T07:00:00Z' -SeedJson '8181') }
        [pscustomobject]@{ Name = 'hub-unindexed.20260924T070000Z.jsonl'; HeaderLine = (New-HeaderLine -CorpusId 'hub-unindexed' -Anchor '2026-09-24T07:00:00Z' -SeedJson '9') }
        [pscustomobject]@{ Name = 'hub-indexed.20260924T070000Z.jsonl'; HeaderLine = '{ torn' }
    )
    Test-Case 'the newest readable one of THIS population is chosen' 'hub-indexed.20260923T070000Z.jsonl|8181' @((Select-NewestHistorySeed -CorpusId 'hub-indexed' -Entries $entries).Name, (Select-NewestHistorySeed -CorpusId 'hub-indexed' -Entries $entries).Seed)
    Test-Case 'none of it: nothing' '<null>' (Select-NewestHistorySeed -CorpusId 'hub-other' -Entries $entries)
    $historyName = Format-HubHistoryName -CorpusId 'hub-indexed' -AnchorText '2026-09-24T08:00:00Z'
    Test-Case 'a torn-down manifest is kept as <id>.<anchor>.jsonl' 'hub-indexed.20260924T080000Z.jsonl' $historyName
    Test-Case 'which Copy-FromGuest.ps1 never takes for a live manifest' $false ($historyName.StartsWith('corpus-', [System.StringComparison]::OrdinalIgnoreCase))

    Write-Host ''
    Write-Host '== the corpus tool''s arguments =='
    Test-Case 'teardown' 'corpus-teardown|--population|hub|--store|tier@vm.invalid|--allow-store|tier@vm.invalid|--corpus-id|hub-indexed|--seed|8181|--anchor|2026-09-24T08:00:00Z|--manifest|C:\m\corpus-hub-indexed.jsonl' (Get-HubToolArgument -Verb 'corpus-teardown' -Store 'tier@vm.invalid' -CorpusId 'hub-indexed' -SeedValue 8181 -Anchor '2026-09-24T08:00:00Z' -ManifestPath 'C:\m\corpus-hub-indexed.jsonl')
    Test-Case 'build' 'corpus-build|--population|hub|--store|tier@vm.invalid|--allow-store|tier@vm.invalid|--corpus-id|hub-indexed|--seed|8181|--anchor|2026-09-24T09:15:42Z|--manifest|C:\m\corpus-hub-indexed.jsonl' (Get-HubToolArgument -Verb 'corpus-build' -Store 'tier@vm.invalid' -CorpusId 'hub-indexed' -SeedValue 8181 -Anchor '2026-09-24T09:15:42Z' -ManifestPath 'C:\m\corpus-hub-indexed.jsonl')
    Test-Case 'the plan sheet - pure, so no allowlist and no manifest' 'corpus-plan|--population|hub|--store|tier@vm.invalid|--corpus-id|hub-indexed|--seed|8181|--anchor|2026-09-24T09:15:42Z' (Get-HubToolArgument -Verb 'corpus-plan' -Store 'tier@vm.invalid' -CorpusId 'hub-indexed' -SeedValue 8181 -Anchor '2026-09-24T09:15:42Z' -ManifestPath 'C:\m\corpus-hub-indexed.jsonl')
    Test-Case 'the index wait - read-only, so no allowlist' 'corpus-indexed|--population|hub|--store|tier@vm.invalid|--corpus-id|hub-indexed|--seed|8181|--anchor|2026-09-24T09:15:42Z|--manifest|C:\m\corpus-hub-indexed.jsonl|--wait-seconds|900' (Get-HubToolArgument -Verb 'corpus-indexed' -Store 'tier@vm.invalid' -CorpusId 'hub-indexed' -SeedValue 8181 -Anchor '2026-09-24T09:15:42Z' -ManifestPath 'C:\m\corpus-hub-indexed.jsonl' -WaitSeconds 900)
    # Assigned first: the function returns its list as ONE object, and piping a variable is what
    # enumerates it.
    $spaced = Get-HubToolArgument -Verb 'corpus-plan' -Store 'Tier Hub' -CorpusId 'h' -SeedValue 1 -Anchor 'a' -ManifestPath 'm'
    Test-Case 'a store name with a space is one argument' 3 @($spaced | Where-Object { $_ -ceq 'Tier Hub' -or $_ -ceq '--store' -or $_ -ceq 'h' }).Count
    Test-Case 'and the list is exactly as long as its parts' 11 @($spaced).Count

    Write-Host ''
    Write-Host '== the one Quit - mailbox-safety rule 7 =='
    $clean = @([pscustomobject]@{ Store = 'tier@vm.invalid'; Count = 0 }, [pscustomobject]@{ Store = 'bystander@vm.invalid'; Count = -1 })
    Test-Case 'our own account-less Outlook, nothing open, Outboxes empty or absent: quit' 0 (Test-QuitPrecondition -ProfileName 'CorpusProfile' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242) -StartedPid 4242 -InspectorCount 0 -Outboxes $clean).Count
    Test-Says 'the tier profile''s Outlook is never quit' (Test-QuitPrecondition -ProfileName 'OutlookAI-Tier' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242) -StartedPid 4242 -InspectorCount 0 -Outboxes $clean) 'quits no other'
    Test-Says 'nor an Outlook that is not the process it started' (Test-QuitPrecondition -ProfileName 'CorpusProfile' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242, 5151) -StartedPid 4242 -InspectorCount 0 -Outboxes $clean) 'pid(s) 4242, 5151'
    Test-Says 'an open item window refuses' (Test-QuitPrecondition -ProfileName 'CorpusProfile' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242) -StartedPid 4242 -InspectorCount 1 -Outboxes $clean) 'item window(s) are open'
    Test-Says 'an unreadable window count refuses' (Test-QuitPrecondition -ProfileName 'CorpusProfile' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242) -StartedPid 4242 -InspectorCount $null -Outboxes $clean) 'could not be read'
    Test-Says 'an Outbox holding anything refuses' (Test-QuitPrecondition -ProfileName 'CorpusProfile' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242) -StartedPid 4242 -InspectorCount 0 -Outboxes @([pscustomobject]@{ Store = 'tier@vm.invalid'; Count = 2 })) "Outbox of 'tier@vm.invalid' holds 2"
    Test-Says 'an Outbox that cannot be proven empty refuses' (Test-QuitPrecondition -ProfileName 'CorpusProfile' -ExpectedProfile 'CorpusProfile' -RunningPids @(4242) -StartedPid 4242 -InspectorCount 0 -Outboxes @([pscustomobject]@{ Store = 'Corpus A'; Count = $null })) 'could not be proven empty'

    Write-Host ''
    Write-Host '== ImportPRF =='
    Test-Case 'absent everywhere: fine' 0 (Test-ImportPrfReading @([pscustomobject]@{ Path = 'HKCU:\x\16.0\Outlook\Setup'; Value = '' })).Count
    Test-Says 'set anywhere: refused, with the remedy' (Test-ImportPrfReading @([pscustomobject]@{ Path = 'HKCU:\x\16.0\Outlook\Setup'; Value = '' }, [pscustomobject]@{ Path = 'HKCU:\x\17.0\Outlook\Setup'; Value = 'C:\p.prf' })) '-ClearImportPrf -Execute'
    Test-Says 'unreadable: refused - fail-closed' (Test-ImportPrfReading @([pscustomobject]@{ Path = 'HKCU:\x\16.0\Outlook\Setup'; Value = $null })) 'could not be read'

    Write-Host ''
    Write-Host '== the frontier margin =='
    $newest = [datetime]::SpecifyKind([datetime]'2026-09-24T09:14:42', [System.DateTimeKind]::Utc)
    $winter = Get-FrontierWindow -NewestUtc $newest -UtcOffset ([timespan]::FromHours(1)) -NowUtc ([datetime]::SpecifyKind([datetime]'2026-09-24T09:35:00', [System.DateTimeKind]::Utc))
    Test-Case 'W. Europe in winter: 55 minutes from the newest item, 34 of them left' 'True|55|2026-09-24T10:09:42Z|34' @($winter.Discriminates, $winter.MarginMinutes, $winter.DeadlineUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'"), $winter.MinutesLeft)
    Test-Case 'in summer: 115' 115 (Get-FrontierWindow -NewestUtc $newest -UtcOffset ([timespan]::FromHours(2)) -NowUtc $newest).MarginMinutes
    Test-Case 'west of UTC the size of the offset counts, not its sign' 295 (Get-FrontierWindow -NewestUtc $newest -UtcOffset ([timespan]::FromHours(-5)) -NowUtc $newest).MarginMinutes
    Test-Case 'on UTC it cannot tell at all' $false (Get-FrontierWindow -NewestUtc $newest -UtcOffset ([timespan]::Zero) -NowUtc $newest).Discriminates
    Test-Case 'the newest item comes off the plan sheet' '2026-09-24T09:14:42Z' (Get-NewestFromPlanText "  items                 : 68`r`n  received range        : 2024-10-02T11:00:00Z .. 2026-09-24T09:14:42Z`r`n")
    Test-Case 'and a sheet without one gives nothing' '<null>' (Get-NewestFromPlanText 'items : 68')

    Write-Host ''
    Write-Host '== what it says when Outlook is in the way =='
    $advice = Format-RestartAdvice -Then 'run this again.'
    Test-Says 'restart the guest, from session 0' @($advice) 'shutdown /r'
    Test-Says 'and never kill Outlook' @($advice) 'NEVER taskkill OUTLOOK.EXE'

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. These need a guest, and nothing on this machine stands in for them:'
    Write-Host '  * the guest guard (Assert-TestbedGuest), which runs only once this self-test has exited'
    Write-Host '  * reading the settings file, the manifest, hub-history\ and the Setup keys'
    Write-Host '  * the two profile switches (Set-DefaultOutlookProfile.ps1) and the two Outlook starts'
    Write-Host '  * the one graceful Quit - never yet run by any script here - and OUTLOOK.EXE leaving after it'
    Write-Host '  * corpus-teardown, corpus-build and corpus-indexed against a real hub, and the index taking it in'

    if ($script:SelfTestFailures.Count -gt 0) {
        Write-Host ''
        foreach ($failure in $script:SelfTestFailures) { Write-Host "  $failure" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# EVERYTHING BELOW TOUCHES THE MACHINE. Guest only.
# =============================================================================================

# THE GUARD, FIRST - before the log, the settings, the registry, the process list and the tool.
# For vmadmin on a guest it returns without a word.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

function Write-Line {
    param([string] $Text)
    Write-Host $Text
    Add-Content -LiteralPath $LogPath -Value $Text -ErrorAction SilentlyContinue
}

# A NATIVE PROGRAM WHOSE STDERR IS REDIRECTED RUNS THROUGH HERE (Q78). Under
# $ErrorActionPreference = 'Stop', Windows PowerShell 5.1 turns the first line a native program
# writes to a redirected stderr - 2>$null, 2>&1 and *> alike - into a terminating
# NativeCommandError, so one warning or progress line ends the script before its exit code can be
# read. PowerShell 7 does not. Measured on the host 2026-09-24. So, here and only here:
#   * 'Continue' holds in THIS function's scope. The caller's 'Stop' is never changed, so there is
#     nothing to restore and nothing else is relaxed.
#   * The try is load-bearing. Without one, 'Continue' also demotes a terminating error inside the
#     block - the program not being found at all - to a printed message, and the caller goes on to
#     read a stale $LASTEXITCODE. Inside a try it stops the caller exactly as it always did.
#     Measured in both shells.
#   * Stderr lines come back as plain strings in both shells, never as ErrorRecords: 5.1 renders
#     those with a position block around every line, and an empty one as an exception type name.
#   * The exit code is left in $LASTEXITCODE, and the caller checks it.
# Restated in each script that needs it, as this repository restates its shared rules.
# .github/scripts/check-powershell-51.ps1 fails the build on a redirected native call that does
# not go through a function like this one.
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

<# Runs one corpus verb, logs and echoes everything it says, and returns its exit code and text. #>
function Invoke-HubTool {
    param([string[]] $Arguments, [switch] $Write)
    $toolArguments = @($Arguments)
    if ($Write) { $toolArguments += '--execute' }
    Write-Line ''
    Write-Line "=== $($toolArguments -join ' ')"
    # Captured and echoed rather than streamed, so a child that outlives the call cannot hold the
    # pipe open; every number these verbs print is meant to be kept anyway.
    $output = Invoke-NativeCommand { & $ToolsExe @toolArguments 2>&1 }
    $code = $LASTEXITCODE
    $text = ($output | Out-String)
    Write-Line $text.TrimEnd()
    Write-Line "exit $code"
    return [pscustomobject]@{ Code = $code; Text = $text }
}

function Get-OutlookProcessId {
    return , @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Find-OutlookExe {
    if (-not [string]::IsNullOrWhiteSpace($OutlookExe)) {
        if (-not (Test-Path -LiteralPath $OutlookExe)) { throw "-OutlookExe '$OutlookExe' does not exist." }
        return $OutlookExe
    }
    foreach ($candidate in @(
            'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
            'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE')) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    throw 'OUTLOOK.EXE is in neither Office16 Click-to-Run location. Pass -OutlookExe.'
}

<# ImportPRF in every supported major's Setup key: '' absent, $null unreadable. Reads only. #>
function Read-ImportPrfReading {
    $readings = @()
    foreach ($version in $script:SupportedOfficeVersions) {
        $setup = "HKCU:\Software\Microsoft\Office\$version\Outlook\Setup"
        $value = ''
        try {
            if (Test-Path -LiteralPath $setup) {
                $raw = (Get-Item -LiteralPath $setup -ErrorAction Stop).GetValue('ImportPRF', $null)
                if ($null -ne $raw) { $value = [string]$raw }
            }
        }
        catch { $value = $null }
        $readings += [pscustomobject]@{ Path = $setup; Value = $value }
    }
    return , $readings
}

function Read-FirstLine {
    param([string] $Path)
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if (-not [string]::IsNullOrWhiteSpace($line)) { return $line.Trim() }
    }
    return ''
}

function Get-HistoryEntry {
    param([string] $Directory)
    $entries = @()
    if (-not [System.IO.Directory]::Exists($Directory)) { return , $entries }
    foreach ($file in [System.IO.Directory]::GetFiles($Directory, '*.jsonl')) {
        $headerLine = ''
        try { $headerLine = Read-FirstLine $file } catch { $headerLine = '' }
        $entries += [pscustomobject]@{ Name = [System.IO.Path]::GetFileName($file); HeaderLine = $headerLine }
    }
    return , $entries
}

function Set-HubDefaultProfile {
    param([string] $Name)
    Write-Line ''
    Write-Line "=== default profile -> $Name"
    & (Join-Path $PSScriptRoot 'Set-DefaultOutlookProfile.ps1') -Name $Name -Execute -ExpectedUser $ExpectedUser | ForEach-Object { Write-Line ([string]$_) }
}

<#
    Starts Outlook on the tier profile NOT ELEVATED, through Start-OutlookUnelevated.ps1 (staged beside
    this script), and returns @{ Id; Clock } for it. Why not Start-Process: this script runs in
    Register-InteractiveTask.ps1's task, at RunLevel Highest, and an ELEVATED Outlook never feeds the
    Windows Search index (measured on OutlookAI-Indexed, 2026-09-24; runbook section 8 item 22) - so on
    the indexed guest corpus-indexed would wait for a crawl that never comes - and the live run, which
    attaches at the user's own integrity level, could not attach to it either. The rebuild's OWN
    Outlook (step 4) stays elevated, on purpose: the corpus tool runs in this task too, and attaches
    only to an Outlook at its own level.
#>
function Start-HubOutlookUnelevated {
    param([string] $ProfileName, [string] $What)
    Write-Line ''
    Write-Line "=== starting Outlook for $What, NOT elevated: Start-OutlookUnelevated.ps1 -Profile $ProfileName"
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    & (Join-Path $PSScriptRoot 'Start-OutlookUnelevated.ps1') -Profile $ProfileName -ExpectedUser $ExpectedUser | ForEach-Object { Write-Line ([string]$_) }
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        throw "Start-OutlookUnelevated.ps1 exited $code - its output above says why. The hub IS rebuilt; the default profile is '$ProfileName'. Start Outlook NOT elevated by hand (that script), then re-run this with -Execute -SkipRebuild."
    }
    $pids = Get-OutlookProcessId
    if ($pids.Count -ne 1) {
        throw "Start-OutlookUnelevated.ps1 returned, and OUTLOOK.EXE is running as $($pids.Count) process(es) ($($pids -join ', ')). The hub IS rebuilt. " + (Format-RestartAdvice -Then 'run this script with -Execute -SkipRebuild.')
    }
    return [pscustomobject]@{ Id = $pids[0]; Clock = $clock }
}

<# Starts Outlook on the default profile and waits until it has been up $Seconds, alive throughout. #>
function Start-HubOutlook {
    param([string] $What, [int] $Seconds)
    $exe = Find-OutlookExe
    Write-Line ''
    Write-Line "=== starting Outlook for $What ($exe); waiting ${Seconds}s for it to settle"
    $process = Start-Process -FilePath $exe -PassThru
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    Wait-HubOutlook -ProcessId $process.Id -Clock $clock -Seconds $Seconds -What $What
    return [pscustomobject]@{ Id = $process.Id; Clock = $clock }
}

function Wait-HubOutlook {
    param([int] $ProcessId, [System.Diagnostics.Stopwatch] $Clock, [int] $Seconds, [string] $What)
    while ($Clock.Elapsed.TotalSeconds -lt $Seconds) {
        Start-Sleep -Seconds 5
        if (@(Get-Process -Id $ProcessId -ErrorAction SilentlyContinue).Count -eq 0) {
            throw "OUTLOOK.EXE (pid $ProcessId), started for $What, left after $([int]$Clock.Elapsed.TotalSeconds)s. A profile that would not open - a dialog, a missing store - is the usual cause; start it by hand in session 1 and look."
        }
    }
    $responding = $null
    try { $responding = (Get-Process -Id $ProcessId -ErrorAction Stop).Responding } catch { $responding = $null }
    if ($false -eq $responding) {
        Write-Line "  note: OUTLOOK.EXE is not responding to window messages after $([int]$Clock.Elapsed.TotalSeconds)s - what a modal dialog looks like from outside. Not refused on; a busy Outlook looks the same for a moment."
    }
}

<#
    Waits up to $Seconds for OUTLOOK.EXE to be the one process $StartedPid alone, and returns the pids
    running when it stops waiting. The corpus tool attaches to Outlook over COM, and attaching to an
    Outlook started as a PROGRAM launches a transient 'OUTLOOK.EXE -Embedding' that hands off and
    exits within seconds (measured on OAI-UNINDEXED, 2026-09-24) - one may still be leaving when the
    build returns, and the quit's precondition would otherwise read it as a second Outlook.
#>
function Wait-OnlyStartedOutlook {
    param([int] $StartedPid, [int] $Seconds)
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $pids = Get-OutlookProcessId
        if (($pids.Count -eq 1 -and $pids[0] -eq $StartedPid) -or $clock.Elapsed.TotalSeconds -ge $Seconds) { return , $pids }
        Start-Sleep -Seconds 2
    }
}

<#
    Quits the Outlook THIS SCRIPT STARTED, gracefully, under mailbox-safety rule 7 - or refuses and
    leaves it running. Every reference but the Application is released BEFORE the Quit; the
    Application's own right after it.

    It attaches through the Running Object Table first, as the build-out's quits of 2026-09-24 did
    (each left in about 2 s): an Outlook started as a program is registered there, and attaching that
    way starts nothing. Only when the ROT has no Outlook - or the shell has no GetActiveObject, which
    PowerShell 7's .NET does not - does it attach through the class factory. The process list is read
    BEFORE either attach, so the pid check compares what was running, not what the attach started.
#>
function Close-HubOutlook {
    param([int] $StartedPid, [string] $ExpectedProfile, [int] $TimeoutSeconds)
    Write-Line ''
    Write-Line "=== quitting the Outlook this script started (pid $StartedPid), under mailbox-safety rule 7"

    $app = $null
    $ns = $null
    $inspectors = $null
    $stores = $null
    $held = New-Object System.Collections.ArrayList
    $profileName = $null
    $inspectorCount = $null
    $outboxes = @()
    $runningBefore = Wait-OnlyStartedOutlook -StartedPid $StartedPid -Seconds 30
    try {
        try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            Write-Line '  attached through the Running Object Table'
        }
        catch {
            $app = New-Object -ComObject Outlook.Application
            Write-Line '  not in the Running Object Table - attached through the class factory'
        }
        $ns = $app.GetNamespace('MAPI')
        try { $profileName = [string]$ns.CurrentProfileName } catch { $profileName = $null }
        try { $inspectors = $app.Inspectors; $inspectorCount = [int]$inspectors.Count } catch { $inspectorCount = $null }
        $stores = $ns.Stores
        $storeCount = [int]$stores.Count
        for ($i = 1; $i -le $storeCount; $i++) {
            $store = $stores.Item($i)
            [void]$held.Add($store)
            $name = '(unnamed)'
            try { $name = [string]$store.DisplayName } catch { $name = '(unnamed)' }
            $count = $null
            try {
                # The Outbox is proven present by the store's own mask before it is opened: on a PST
                # the plain lookup would CREATE one the store lacks (Q84).
                $accessor = $store.PropertyAccessor
                [void]$held.Add($accessor)
                $mask = [int]$accessor.GetProperty($script:ValidFolderMaskSchema)
                if (($mask -band $script:FolderIpmOutboxValid) -eq 0) { $count = -1 }
                else {
                    $outbox = $store.GetDefaultFolder($script:OlFolderOutbox)
                    [void]$held.Add($outbox)
                    $items = $outbox.Items
                    [void]$held.Add($items)
                    $count = [int]$items.Count
                }
            }
            catch { $count = $null }
            $outboxes += [pscustomobject]@{ Store = $name; Count = $count }
        }
    }
    finally {
        for ($i = $held.Count - 1; $i -ge 0; $i--) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($held[$i]) }
        foreach ($reference in @($inspectors, $stores, $ns)) {
            if ($null -ne $reference) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($reference) }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    foreach ($outbox in $outboxes) {
        $shown = 'no Outbox'
        if ($null -eq $outbox.Count) { $shown = 'Outbox UNREADABLE' } elseif ($outbox.Count -ge 0) { $shown = "Outbox $($outbox.Count)" }
        Write-Line ("  {0,-40} {1}" -f $outbox.Store, $shown)
    }
    Write-Line "  profile '$profileName', item windows open: $inspectorCount"

    $problems = Test-QuitPrecondition -ProfileName $profileName -ExpectedProfile $ExpectedProfile `
        -RunningPids $runningBefore -StartedPid $StartedPid -InspectorCount $inspectorCount -Outboxes $outboxes
    if ($problems.Count -gt 0) {
        if ($null -ne $app) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
        [GC]::Collect()
        throw ("REFUSING to quit Outlook - mailbox-safety rule 7:`n  - " + ($problems -join "`n  - ") +
            "`nOutlook is left running. The hub IS rebuilt. " + (Format-RestartAdvice -Then 'run this script with -Execute -SkipRebuild to switch back to the tier profile and wait for the index.'))
    }

    try { $app.Quit() }
    finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        $app = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
    }

    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    while ($clock.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ((Get-OutlookProcessId).Count -eq 0) {
            Write-Line "  OUTLOOK.EXE left $([int]$clock.Elapsed.TotalSeconds)s after the Quit."
            return
        }
        Start-Sleep -Seconds 2
    }
    throw ("OUTLOOK.EXE is still running ${TimeoutSeconds}s after a graceful Quit. The hub IS rebuilt. The likeliest " +
        "cause is a modal dialog: one swallows Application.Quit() - measured on OAI-UNINDEXED 2026-09-24 with an Object " +
        "Model Guard prompt, where Quit() returned and Outlook was still up four minutes later. " +
        (Format-RestartAdvice -Then 'run this script with -Execute -SkipRebuild to switch back to the tier profile and wait for the index.'))
}

# ---------------------------------------------------------------------------------------------
# 1. What the settings, the manifest and the history say.
# ---------------------------------------------------------------------------------------------
Remove-Item -LiteralPath $LogPath -Force -ErrorAction SilentlyContinue
$mode = 'dry run'
if ($Execute -and $SkipRebuild) { $mode = '-Execute -SkipRebuild' } elseif ($Execute) { $mode = '-Execute' }
Write-Line "Reset-HubPopulation, $mode, $([datetime]::UtcNow.ToString($script:UtcFormat, $script:Invariant)) on $env:COMPUTERNAME"

if (-not (Test-Path -LiteralPath $SettingsPath)) {
    throw "REFUSING: no live-test settings at $SettingsPath. Render them on the host with Testbed/host/New-LiveTestSettings.ps1 -VMName <this guest> and copy them in with the line it prints."
}
$settings = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($SettingsPath)) -ErrorAction Stop
$fact = Read-HubSettingsFact $settings

$manifestExists = $false
$header = $null
$historyDirectory = $null
$historySeed = $null
if ($null -ne $fact.ManifestPath) {
    $manifestExists = [System.IO.File]::Exists($fact.ManifestPath)
    $historyDirectory = Join-Path ([System.IO.Path]::GetDirectoryName($fact.ManifestPath)) $script:HistoryDirectoryName
    if ($manifestExists) { $header = Read-HubManifestHeader (Read-FirstLine $fact.ManifestPath) }
    elseif ($null -ne $fact.PopulationId) { $historySeed = Select-NewestHistorySeed -CorpusId $fact.PopulationId -Entries (Get-HistoryEntry $historyDirectory) }
}

$plan = Resolve-HubResetPlan -SettingsFact $fact -ManifestExists $manifestExists -Header $header -HistorySeed $historySeed `
    -SeedText $Seed -SkipRebuild ([bool]$SkipRebuild) -NowUtc ([datetime]::UtcNow)

Write-Line ''
Write-Line "  settings    $SettingsPath"
Write-Line "  hub         $($plan.Store)"
Write-Line "  population  $($plan.CorpusId), manifest $($plan.ManifestPath) ($(if ($manifestExists) { 'present' } else { 'ABSENT' }))"
if ($null -ne $header -and $header.Problems.Count -eq 0) { Write-Line "  built       against $($header.AnchorUtc), seed $($header.Seed)" }
Write-Line "  index       $(if ($plan.Indexed) { 'the hub is indexed here - the run waits for the index to take the new population' } else { 'the hub is not indexed here - no index wait' })"
foreach ($note in $plan.Notes) { Write-Line "  note: $note" }

# ---------------------------------------------------------------------------------------------
# 2. The machine's state: Outlook closed, nothing waiting to rebuild a profile, the tool there.
# ---------------------------------------------------------------------------------------------
$refusals = New-Object System.Collections.Generic.List[string]
foreach ($problem in $plan.Problems) { $refusals.Add($problem) }
$running = Get-OutlookProcessId
if ($running.Count -gt 0) {
    $refusals.Add("OUTLOOK.EXE is running (pid $($running -join ', ')). This script switches the default profile, which needs Outlook closed, and it closes no Outlook it did not start. " + (Format-RestartAdvice -Then 'run this again.'))
}
foreach ($problem in (Test-ImportPrfReading (Read-ImportPrfReading))) { $refusals.Add($problem) }
if (-not (Test-Path -LiteralPath $ToolsExe)) {
    $refusals.Add("the corpus tool is not at $ToolsExe. Publish it on the host (Testbed/host/Publish-GuestPayload.ps1) from a commit that knows population v2 and corpus-indexed, and copy it in.")
}

if ($refusals.Count -gt 0) {
    $text = "REFUSING to rebuild the hub population - $($refusals.Count) reason(s):`n  - " + ($refusals -join "`n  - ") + "`nNothing was changed."
    Write-Line ''
    Write-Line $text
    throw $text
}

if (-not $Execute) {
    [void](Invoke-HubTool -Arguments (Get-HubToolArgument -Verb 'corpus-plan' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor $plan.NewAnchor -ManifestPath $plan.ManifestPath))
    Write-Line ''
    Write-Line 'Dry run. Nothing written, no Outlook started. With -Execute this would, in order:'
    if ($plan.Rebuild) {
        Write-Line "  - make '$CorpusProfileName' the default profile, start Outlook, wait ${WarmupSeconds}s"
        if ($null -ne $plan.TeardownAnchor) {
            Write-Line ("  - " + ((Get-HubToolArgument -Verb 'corpus-teardown' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor $plan.TeardownAnchor -ManifestPath $plan.ManifestPath) -join ' ') + ' --execute')
            Write-Line "  - move the manifest to $historyDirectory\$(Format-HubHistoryName -CorpusId $plan.CorpusId -AnchorText $plan.TeardownAnchor)"
        }
        Write-Line ("  - " + ((Get-HubToolArgument -Verb 'corpus-build' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor '<now>' -ManifestPath $plan.ManifestPath) -join ' ') + ' --execute')
        Write-Line '  - quit that Outlook gracefully (mailbox-safety rule 7)'
    }
    Write-Line "  - make '$TierProfileName' the default profile, start Outlook NOT elevated (Start-OutlookUnelevated.ps1)"
    if ($plan.Indexed) { Write-Line ("  - " + ((Get-HubToolArgument -Verb 'corpus-indexed' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor '<the anchor built>' -ManifestPath $plan.ManifestPath -WaitSeconds $IndexWaitSeconds) -join ' ')) }
    Write-Line '  - leave Outlook running, warm, and report the frontier test''s margin'
    return
}

# ---------------------------------------------------------------------------------------------
# 3-8. The rebuild, in the account-less profile.
# ---------------------------------------------------------------------------------------------
$anchorBuilt = $null
$newestText = $null
if ($null -ne $header -and $header.Problems.Count -eq 0) { $anchorBuilt = $header.AnchorUtc }

if ($plan.Rebuild) {
    Set-HubDefaultProfile -Name $CorpusProfileName
    $corpusOutlook = Start-HubOutlook -What "the rebuild, on '$CorpusProfileName'" -Seconds $WarmupSeconds

    if ($null -ne $plan.TeardownAnchor) {
        $teardown = Invoke-HubTool -Write -Arguments (Get-HubToolArgument -Verb 'corpus-teardown' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor $plan.TeardownAnchor -ManifestPath $plan.ManifestPath)
        if ($teardown.Code -ne 0) {
            throw "corpus-teardown failed (exit $($teardown.Code)); the manifest is where it was, and Outlook is running on '$CorpusProfileName'. Read the output above. If it says the manifest records a different shape, the population was built by an older generator: corpus-reindex it into a NEW manifest path, inspect that, and tear down with it - never by hand. Then restart the guest and run this again. Log: $LogPath"
        }

        [void][System.IO.Directory]::CreateDirectory($historyDirectory)
        $historyPath = Join-Path $historyDirectory (Format-HubHistoryName -CorpusId $plan.CorpusId -AnchorText $plan.TeardownAnchor)
        $suffix = 1
        while ([System.IO.File]::Exists($historyPath)) {
            $historyPath = Join-Path $historyDirectory ((Format-HubHistoryName -CorpusId $plan.CorpusId -AnchorText $plan.TeardownAnchor) + '.' + $suffix)
            $suffix++
        }
        [System.IO.File]::Move($plan.ManifestPath, $historyPath)
        Write-Line "  torn-down manifest kept as $historyPath"
    }

    # The anchor is taken NOW - after the warm-up and the teardown - so the newest item is as young
    # as it can be when the run starts.
    $anchorBuilt = [datetime]::UtcNow.ToString($script:UtcFormat, $script:Invariant)
    $build = Invoke-HubTool -Write -Arguments (Get-HubToolArgument -Verb 'corpus-build' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor $anchorBuilt -ManifestPath $plan.ManifestPath)
    if ($build.Code -ne 0) {
        throw "corpus-build failed (exit $($build.Code)); Outlook is running on '$CorpusProfileName'. Read the output above: a probe that refused wrote nothing, and a build that failed after writing left a manifest the next run tears down. Restart the guest and run this again. Log: $LogPath"
    }

    $rebuilt = Read-HubManifestHeader (Read-FirstLine $plan.ManifestPath)
    if ($rebuilt.Problems.Count -gt 0 -or $rebuilt.AnchorUtc -cne $anchorBuilt -or $rebuilt.CorpusId -cne $plan.CorpusId) {
        throw "The new manifest at $($plan.ManifestPath) does not read back as the population just built (anchor '$($rebuilt.AnchorUtc)', wanted '$anchorBuilt'; $($rebuilt.Problems -join ' ')). Do not start a run on it."
    }
    $newestText = Get-NewestFromPlanText $build.Text

    Close-HubOutlook -StartedPid $corpusOutlook.Id -ExpectedProfile $CorpusProfileName -TimeoutSeconds $QuitTimeoutSeconds
}

# ---------------------------------------------------------------------------------------------
# 9-10. Back to the tier profile, and the index.
# ---------------------------------------------------------------------------------------------
Set-HubDefaultProfile -Name $TierProfileName
$tierOutlook = Start-HubOutlookUnelevated -ProfileName $TierProfileName -What "the live run, on '$TierProfileName'"

if ($plan.Indexed) {
    $indexed = Invoke-HubTool -Arguments (Get-HubToolArgument -Verb 'corpus-indexed' -Store $plan.Store -CorpusId $plan.CorpusId -SeedValue $plan.Seed -Anchor $anchorBuilt -ManifestPath $plan.ManifestPath -WaitSeconds $IndexWaitSeconds)
    if ($indexed.Code -ne 0) {
        throw "The index did not take the rebuilt hub in within ${IndexWaitSeconds}s. A run now would measure part of it, and the frontier test would fail on a frontier older than the population. Outlook is running on '$TierProfileName'; if the indexer is slow, run corpus-indexed again with a longer --wait-seconds, and if it never advances, check the guest is in scope (Set-OutlookIndexingDisabled.ps1 -Verify). Log: $LogPath"
    }
}
Wait-HubOutlook -ProcessId $tierOutlook.Id -Clock $tierOutlook.Clock -Seconds $WarmupSeconds -What "the live run, on '$TierProfileName'"

$newest = $null
if ($null -ne $newestText) { $newest = ConvertFrom-UtcText $newestText }
$about = ''
if ($null -eq $newest) {
    $newest = (ConvertFrom-UtcText $anchorBuilt).AddMinutes(-1)
    $about = 'about '
}
$offset = [System.TimeZoneInfo]::Local.GetUtcOffset([datetime]::UtcNow)
$window = Get-FrontierWindow -NewestUtc $newest -UtcOffset $offset -NowUtc ([datetime]::UtcNow)

Write-Line ''
Write-Line "Hub population '$($plan.CorpusId)' in '$($plan.Store)' is anchored $anchorBuilt; its newest item is ${about}$($newest.ToString($script:UtcFormat, $script:Invariant))."
if (-not $window.Discriminates) {
    Write-Line 'This guest is on UTC, where the frontier test cannot tell a local-time frontier from a UTC one at all; it refuses here whatever the hub holds.'
}
elseif ($window.MinutesLeft -le 0) {
    Write-Line "The frontier test's margin ($($window.MarginMinutes) min) is ALREADY SPENT: it would refuse the run as STALE. Run this script again and start the run straight after it."
}
else {
    Write-Line "The frontier test can catch a local-time misreading until $($window.DeadlineUtc.ToString($script:UtcFormat, $script:Invariant)) - $($window.MinutesLeft) min from now ($($window.MarginMinutes) min on this guest's UTC offset of $offset). START THE RUN NOW, by Testbed/README.md section 4c, with:"
    $filter = 'Category=Live&Requires!=DelegateStore'
    if (-not $plan.Indexed) { $filter = 'Category=Live&Requires!=DelegateStore&Requires!=SearchIndex' }
    Write-Line "  the opt-in   `$env:OUTLOOKAI_LIVE_OPT_IN = '$env:COMPUTERNAME'"
    Write-Line "  the filter   --filter `"$filter`""
}
Write-Line "Outlook is running on '$TierProfileName', NOT elevated, and is left running. Log: $LogPath"
