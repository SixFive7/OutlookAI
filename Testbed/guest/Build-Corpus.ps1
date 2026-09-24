<#
.SYNOPSIS
    preflight -> plan -> probe -> build -> census, with the committed corpus parameters as
    defaults.

.DESCRIPTION
    RUN THIS ON THE GUEST, IN SESSION 1. Everything except the plan step drives Outlook over COM,
    and COM does not work in session 0. Reach session 1 with Register-InteractiveTask.ps1:

        .\Register-InteractiveTask.ps1 -Script ".\Build-Corpus.ps1 -Execute"

    Windows PowerShell 5.1 - no ternary, no `??`.

    ============================================================================================
    THE PREFLIGHT'S FIRST GUEST RUN, OAI-UNINDEXED 2026-09-24, FROM CP-05
    ============================================================================================

      both failing     default OutlookAI-Tier, Outlook closed, -Execute: REFUSED with BOTH
                       reasons, and wrote nothing - the manifest's SHA-256 and line count, the
                       corpus .pst's size AND last-write time, DefaultProfile, PickLogonProfile and
                       ImportPRF all read the same before and after. Only this run's own log was
                       rewritten, which the refusal used to deny; it now says so.
      the good case    default CorpusProfile, Outlook up 210 s: REFUSED - "it has 2 account entries
                       under ...\CorpusProfile\9375CFF0413111d3B88A00104B2A6676". That profile has
                       no mail account (COM: Accounts.Count 0, "profile accounts: 0"); the two
                       entries are its data file and its address book, which Outlook lists in the
                       same key. The preflight would have refused every correct build. Fixed:
                       Test-IsMailAccountEntry, below, and the measured shapes are in -SelfTest.
      after the fix    fresh CP-05, default CorpusProfile, Outlook up 208 s: the preflight PASSED -
                       "mail accounts : 0", "ImportPRF : <not set>", then "OK - the default profile
                       has no accounts, Outlook is up and warm, and ImportPRF is not set." - and
                       the whole pipeline ran against this guest's corpus (vm-unindexed, seed 7777,
                       anchor 2026-09-16, 20,000 items, store 'Outlook Data File', its existing
                       manifest): both probes VERIFIED (DraftsThenMoveWithSentFlag,
                       PropertyAccessorDates), "Build finished: created 0, already present 20,000,
                       failed 0", and the census "20,000 item(s) found for 20,000 planned ...
                       Every ordinal exists exactly once, in the folder the plan names." Exit 0,
                       1,490 s wall clock, the manifest's SHA-256 unchanged.
      ImportPRF set    on top of both failures above: REFUSED with THREE reasons; with
                       -SkipPreflight it still REFUSED, on ImportPRF alone; nothing written either
                       time. New-OutlookProfile.ps1 -ClearImportPrf -Execute then removed the
                       value, read back absent.

    ============================================================================================
    THREE PRECONDITIONS. THE FIRST TWO COST FIVE FAILED BUILDS TO ESTABLISH (2026-09-16).
    ============================================================================================

    A corpus build succeeds only if BOTH of the first two hold before it starts. Neither is a
    nicety and neither auto-repairs, so this script CHECKS both and REFUSES with the remedy rather
    than letting the tool discover them one expensive attempt at a time. The third is different in
    kind - it protects a corpus rather than predicting a build - and is described where it is
    listed.

    1. THE DEFAULT OUTLOOK PROFILE MUST BE THE ACCOUNT-LESS ONE. The corpus tool LOGS ON with the
       DEFAULT profile - it does not attach to whatever Outlook happens to be running, and there
       is no flag that points it at a profile. With the wrong default, the store guard refuses
       with "REFUSING to build a corpus in store '(unnamed store)' ... profile accounts: 1".

    2. OUTLOOK MUST ALREADY BE RUNNING, AND WARM. The tool COM-activates Outlook.Application
       (Type.GetTypeFromProgID + Activator.CreateInstance). With nothing running, that starts a
       COLD Outlook inside the tool's own STA thread, which blocks past the three-minute bound on
       ReadStoreFacts / ReadProfileFacts and comes back as
       "FATAL: TimeoutException: Corpus STA operation timed out".

    THE ONLY BUILD THAT EVER SUCCEEDED SATISFIED BOTH BY ACCIDENT, because an earlier task had
    left an Outlook running on the corpus profile. That is exactly why it looked reproducible and
    was not - and it is why this check exists rather than a paragraph in a runbook.

    3. NOTHING MAY BE WAITING TO REBUILD THE PROFILE - ImportPRF must not be set. Added 2026-09-24
       (Q66), and unlike the two above it protects the CORPUS rather than predicting the build.
       The profile scripts build profiles by pointing ImportPRF at a .prf carrying
       OverwriteProfile=Yes, and a re-import rebuilds the profile from the file - a store attached
       or filled since stops being part of it. Outlook DOES remove the value itself when it
       imports (measured on OAI-UNINDEXED 2026-09-24: gone within 5 s of the start that imported
       it), so on the normal path this never fires; it is here for the abnormal one, because a
       corpus is thirteen minutes and twenty thousand items to lose. It is FAIL-CLOSED - an
       unreadable value refuses - because nothing downstream ever re-checks it, and -SkipPreflight
       does NOT skip it. The remedy is printed: the -Verify of the script that wrote the value
       removes it once its import has run, and New-OutlookProfile.ps1 -ClearImportPrf -Execute
       removes it unconditionally.

    THE PREFLIGHT DOES NOT FIX EITHER OF THEM, DELIBERATELY. A script that silently rewrites the
    DefaultProfile value changes the machine out from under the next step - the tier tests need
    the TIER profile as the default - and one that starts Outlook for you hides the cold-start
    cost it is meant to be avoiding. It refuses, says which precondition failed, and prints the
    exact commands that repair it.

    WHAT THE PREFLIGHT READS. The registry (which profile is default, which profiles exist, how
    many account entries the default profile has) and the process list (is OUTLOOK.EXE up, and for
    how long). It makes NO COM call of its own: a preflight that could itself hang, or could
    itself start Outlook, would be the very fault it is checking for. The authoritative account
    count is still the tool's own - it prints "profile accounts: N" from COM when it vets the
    store - and this is only the early, cheap version of the same question.

    THE ACCOUNT COUNT IS THE MAIL ACCOUNTS under the profile's account-manager key
    (9375CFF0413111d3B88A00104B2A6676) - NOT its subkey count, which is what it was until the
    first guest run of this preflight, 2026-09-24, refused the real corpus profile with "2 account
    entries". Outlook lists every MAPI service of a profile there once it has opened it: the
    account-less CorpusProfile, whose COM Accounts.Count is 0, holds its data file and its address
    book as two {ED475414-...} wrappers, and the tier profile holds the same two plus its POP3
    account ({ED475411-...}). The old "zero subkeys is the account-less shape" was measured on a
    profile Outlook had never opened - which also holds the key empty. Test-IsMailAccountEntry
    now excludes only a wrapper around a data file or an address book, and counts everything else,
    so an unknown shape refuses rather than slips through. If the read fails for any reason the
    preflight says so and does NOT refuse on it - the tool's COM count is the one that decides.

    THE PARAMETERS ARE NOT AN EXAMPLE. corpusId vm2, seed 7777, anchor 2026-08-19, count 20000,
    default shape: those four values reproduce the corpus that every published sweep and frame
    measurement in this repository is a statement about. They are pinned equal here, in
    Testbed/testbed.json and in Docs/corpus-measurement-plan.md, and .github/scripts/check-testbed-references.ps1
    fails the build when the three stop agreeing. Change them only if you mean to build a
    DIFFERENT corpus, and give it a different id when you do.

    ONE CORPUS ID PER GUEST, AND THE MANIFEST IS NAMED AFTER IT. The ids for the two guests being
    built are vm-indexed and vm-unindexed, and a manifest is corpus-<corpusId>.jsonl. That is not
    tidiness: Testbed/host/Copy-FromGuest.ps1 pulls every guest's MANIFEST into one shared
    directory - deliberately, so that a reused id still collides where a human can see it - so two
    guests sharing an id means the second pull replaces the first's manifest, and a manifest is
    the only thing corpus-teardown can remove a corpus with. Change -CorpusId and -Manifest
    together, always. Testbed/README.md section 3 carries the convention and why these names.

    THE PROFILE MUST HAVE NO MAIL ACCOUNTS. There is no override. A build creates unsent items in
    bulk; the first real run put 5,532 of them into the target store's Outbox, inert only because
    that profile could not send. On a profile with an account those are 5,532 real messages queued
    for delivery. So corpus work happens in the no-accounts profile and the tier runs in the other
    one, and switching between them is a restart of Outlook.

    WHAT TO CHECK IN THE OUTPUT before letting a build proceed - the script prints these and stops
    on any of them:

      * the store line and the profile line both say accepted, and `profile accounts: 0`
      * the PLACEMENT probe named a verified rung. On this corpus it was DraftsThenMoveWithSentFlag.
      * the DATE probe named a verified rung. On this corpus it was PropertyAccessorDates.

    A build that had to be talked past either probe is a build whose measurements mean something
    other than what they say. Placement is settled before dates on purpose: a date probed against
    an item filed in the wrong folder cannot tell "the date does not select" from "the item is not
    there", and that confusion is exactly what made the first run's verdict worthless.

    COST, from the 2026-08-19 build of these exact parameters: 20,000 items, 225,282,619 body
    bytes, 13m25s, 24.8 items/s. That rate is the WITH-MOVE figure - DraftsThenMoveWithSentFlag
    writes every item twice. The build is resumable and idempotent: it creates the ordinals the
    manifest lacks, so an interrupted run is finished by running it again, and a finished one
    re-run is a no-op.

    AFTERWARDS, TWO THINGS.

    1. COPY THE MANIFEST OFF THE GUEST. It is the only thing that can tear the corpus down -
       teardown deletes by EntryID allowlist AND subject tag, and there is no second route the
       mailbox-safety rules permit. Host-side: Testbed/host/Copy-FromGuest.ps1.
    2. EXPECT THE OUTBOX TO BE EMPTY. The 2026-08-19 build left 2,761 items there, which is
       EXACTLY the plan's unread count - the MSGFLAG_SUBMIT defect, since fixed. The count is
       predictable in advance, so a non-empty Outbox after a rebuild is a specific signal that
       the fix did not take, not a mystery.

    WHAT IS TESTED WITHOUT A GUEST. -SelfTest exercises every preflight decision and the whole
    refusal text against synthetic facts. It reads no registry, lists no processes, runs no tool
    and needs no Outlook, so it is safe anywhere - the maintainer's workstation included.

.PARAMETER Store
    Store display name. `Outlook Data File` is what the measured corpus lives in; the
    three-store layout in Docs/live-tier-on-the-vm.md calls it `Corpus A`.

.PARAMETER WarmupSeconds
    How long OUTLOOK.EXE must have been up before this script will drive the tool at it. 180 was
    used successfully; 75 was NOT always enough. Lower it only with a measurement in hand.

.PARAMETER OfficeVersion
    Force the Office major whose hive the preflight reads, e.g. '16.0'. Only for a machine the
    detection reads wrongly.

.PARAMETER SkipPreflight
    Run without checking preconditions 1 and 2 (the default profile, a warm Outlook). This does NOT
    weaken any safety guard - the tool's own store guard still refuses a profile that can send, and
    that guard is what protects the mailbox - and it does NOT skip precondition 3, the ImportPRF
    check, which is the only thing that stands between a corpus and a profile rebuild. Use it when
    the preflight is wrong about this machine, and say so in the run log.

.PARAMETER SelfTest
    Run the preflight decision tests and exit. Touches nothing.

.PARAMETER Execute
    Without it: plan and a dry-run build, and NEITHER probe runs (both create items).

.EXAMPLE
    .\Build-Corpus.ps1 -SelfTest
    .\Build-Corpus.ps1
    .\Build-Corpus.ps1 -Execute
    .\Build-Corpus.ps1 -Store "Corpus B" -CorpusId vm-unindexed -Manifest C:\OutlookAI-Q5\corpus-vm-unindexed.jsonl -Execute
#>
[CmdletBinding()]
param(
    [string] $ToolsExe = 'C:\OutlookAI-Q5\tools\OutlookAI.RemediationTools.exe',
    [string] $Store = 'Outlook Data File',

    # --- pinned corpus parameters; see check-testbed-references.ps1 --------------------------
    [string] $CorpusId = 'vm2',
    [long]   $Seed = 7777,
    [string] $Anchor = '2026-08-19',
    [int]    $Count = 20000,
    # ----------------------------------------------------------------------------------------

    [string] $Manifest = 'C:\OutlookAI-Q5\corpus-vm2.jsonl',
    [string] $LogPath = 'C:\OutlookAI-Q5\corpus-build.log',
    [int]    $ProgressEvery = 250,
    [int]    $WarmupSeconds = 180,
    [string] $OfficeVersion,
    [switch] $SkipProbe,
    [switch] $SkipPreflight,
    [switch] $SelfTest,
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

# The Office majors this product supports, IN PROBE ORDER, mirrored from OfficeVersions.Supported
# exactly as Set-DefaultOutlookProfile.ps1 mirrors it. The two copies exist because there is no
# shared testbed library to hold one, and adding a file under Testbed/ means adding a row to
# Testbed/README.md section 5 - the inventory check fails otherwise.
$script:SupportedOfficeVersions = @('16.0', '17.0', '15.0')
$script:InstallerFootprintSubKeyName = 'Resiliency'
$script:OfficeRootKeyPath = 'HKCU:\Software\Microsoft\Office'

# The per-account container inside a profile. An Outlook-internal GUID, stable across every
# Outlook version this product supports: OutlookProfileRegistry.AccountsSubKeyName.
$script:AccountsSubKeyName = '9375CFF0413111d3B88A00104B2A6676'

# NOT EVERY ENTRY IN THAT CONTAINER IS A MAIL ACCOUNT - measured on OAI-UNINDEXED 2026-09-24, and
# the preflight's first guest run refused the real corpus profile over it. Outlook lists EVERY
# MAPI service of the profile there, wrapping the non-account ones under one CLSID: CorpusProfile,
# whose COM Accounts.Count is 0, holds two entries - clsid {ED475414-...} 'Outlook Address Book'
# (Service Name CONTAB) and {ED475414-...} 'Outlook Data File' (MSUPST MS) - and the tier
# profile holds the same two plus its POP3 account, clsid {ED475411-...}. A wrapper around a data
# file or an address book is not an account; anything else is counted, including a wrapper around
# a service not named here (an Exchange mailbox is a wrapper around MSEMS), so an unknown shape
# refuses rather than slipping through. The service names are the ones section 6 of this
# project's .prf files maps.
$script:ServiceWrapperClsid = '{ED475414-B0D6-11D2-8C3B-00104B2A6676}'
$script:NonMailServiceNames = @('MSUPST MS', 'MSPST MS', 'CONTAB', 'EMABLT', 'MSPST AB')

# =============================================================================================
# PURE PREFLIGHT DECISIONS. No registry, no processes, no tool, no Outlook, no output.
# =============================================================================================

<#
    Mirrors OfficeVersions.IsOutlookHive: at least one VALUE, or at least one subkey that is not
    'Resiliency' - which this product's own installer writes under every supported major, so a
    bare key-exists probe answers 16.0 on a machine running Outlook 2013.
#>
function Test-IsOutlookHive {
    param([string[]] $ValueNames, [string[]] $SubKeyNames)

    if ($null -ne $ValueNames -and $ValueNames.Count -gt 0) { return $true }
    if ($null -eq $SubKeyNames) { return $false }

    foreach ($subKey in $SubKeyNames) {
        if ($subKey -ne $script:InstallerFootprintSubKeyName) { return $true }
    }
    return $false
}

<#
    The first candidate that is a real hive, in the product's probe order - or the one named by
    -OfficeVersion, if it is a real hive. $null when nothing qualifies.

    Candidates are shaped @{ Version = '16.0'; Path = '...'; ValueNames = @(); SubKeyNames = @() }.
#>
function Select-OutlookHive {
    param([object[]] $Candidates, [string] $RequestedVersion)

    if ($null -eq $Candidates) { return $null }

    if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
        foreach ($candidate in $Candidates) {
            if ($candidate.Version -eq $RequestedVersion -and
                (Test-IsOutlookHive -ValueNames $candidate.ValueNames -SubKeyNames $candidate.SubKeyNames)) {
                return $candidate
            }
        }
        return $null
    }

    foreach ($version in $script:SupportedOfficeVersions) {
        foreach ($candidate in $Candidates) {
            if ($candidate.Version -eq $version -and
                (Test-IsOutlookHive -ValueNames $candidate.ValueNames -SubKeyNames $candidate.SubKeyNames)) {
                return $candidate
            }
        }
    }
    return $null
}

<#
    Whether one entry of the account-manager container is a MAIL account. Pure.

    $false only for a service wrapper ($script:ServiceWrapperClsid) around a data file or an
    address book. Everything else is an account - including a wrapper around a service this list
    does not name, and an entry whose clsid cannot be read - because the cost of calling a real
    account "not one" is a build that queues mail, and the cost of the opposite is a refusal that
    names the entry.
#>
function Test-IsMailAccountEntry {
    param([string] $Clsid, [string] $ServiceName)

    if ($Clsid -eq $script:ServiceWrapperClsid -and $script:NonMailServiceNames -contains $ServiceName) { return $false }
    return $true
}

<#
    Precondition 3, and the one that protects the CORPUS rather than predicting the build.

    Returns the refusal text, or $null. Pure, and deliberately a function of its own: it is the one
    check -SkipPreflight does NOT skip, so the script calls it on that path too.

    THE HAZARD. The profile scripts build profiles by pointing ImportPRF at a .prf that carries
    OverwriteProfile=Yes. Outlook acts on ImportPRF at a start where First-Run is absent - and a
    profile it re-imports is REBUILT from the file, so a store attached or filled since is no
    longer part of it. The corpus is exactly such a store: twenty thousand items, about thirteen
    minutes to rebuild. Building one while ImportPRF is set is building it under a value that can
    detach it.

    WHY IT IS FAIL-CLOSED WHEN THE OTHER TWO ARE NOT. Preconditions 1 and 2 predict whether the
    build will SUCCEED, and the tool's own guards decide that authoritatively, so an unreadable
    fact there is a note. Nothing downstream ever looks at ImportPRF - the tool vets accounts and
    stores, not the Setup key - so if this cannot prove the value absent, nothing will, and it
    refuses.

    First-Run's state changes the WORDING, not the verdict. Present, Outlook ignores ImportPRF for
    now and the hazard is latent; absent, the very next start acts on it. Either way the remedy is
    the same and it is one command.
#>
function Test-CorpusImportPrfGuard {
    param([psobject] $Facts)

    if ($null -eq $Facts.ImportPrfValue) {
        $where = $Facts.SetupKeyPath
        if ([string]::IsNullOrEmpty($where)) { $where = "the Setup key (no Outlook hive was found to hold one)" }
        return "ImportPRF could not be read from $where, so nothing proves that no Outlook start is waiting to rebuild the profile this corpus goes into. Refused on, unlike an unreadable account count or uptime: nothing downstream re-checks ImportPRF. If the hive detection is wrong for this machine, pass -OfficeVersion."
    }

    if ($Facts.ImportPrfValue.Length -eq 0) { return $null }

    $state = "First-Run is ABSENT, so the very next Outlook start acts on it"
    if ($Facts.FirstRunPresent -eq $true) {
        $state = "First-Run is present, so Outlook ignores it for now - latent, not gone: whatever next removes First-Run arms it"
    }
    return "ImportPRF is set under $($Facts.SetupKeyPath): '$($Facts.ImportPrfValue)'. $state. The profile scripts' .prf files carry OverwriteProfile=Yes, and a re-import REBUILDS the profile - a store attached or filled since stops being part of it, and this build is about to fill one with twenty thousand items. Remove it first: the -Verify of the script that wrote it does so once the import has run (.\New-OutlookProfile.ps1 -Name <profile> -Verify, or .\New-TierProfile.ps1 -Verify), and .\New-OutlookProfile.ps1 -ClearImportPrf -Execute removes it unconditionally."
}

<#
    All three preconditions, judged against facts somebody else gathered.

    Returns Ok, Problems (each one a refusal reason, all of them collected rather than the first -
    an operator who has to fix them one run at a time is exactly what this replaces) and Notes
    (things worth saying that do not justify refusing).

    WHAT IS A PROBLEM AND WHAT IS A NOTE. For preconditions 1 and 2, a problem is a fact that
    PROVES the build will fail, and a note is a fact this preflight could not read, or one that is
    merely unusual. Unreadable is never a refusal for those two: the tool's own guards are the
    ones that protect the mailbox, and a preflight that blocked a good build over a registry read
    it could not make would cost more than it saves. Precondition 3 is the exception and says why
    - see Test-CorpusImportPrfGuard.
#>
function Test-CorpusPrecondition {
    param([psobject] $Facts)

    $problems = @()
    $notes = @()

    # ---- precondition 1: the default profile is the account-less one -------------------------
    if ([string]::IsNullOrEmpty($Facts.HivePath)) {
        $problems += "No Outlook hive found under $script:OfficeRootKeyPath (probed $($script:SupportedOfficeVersions -join ', '), and a key holding nothing but '$script:InstallerFootprintSubKeyName' does not count). Without it there is no way to read which profile is default. If Outlook is installed under another major, pass -OfficeVersion."
    }
    elseif ([string]::IsNullOrWhiteSpace($Facts.DefaultProfile)) {
        $problems += "There is no DefaultProfile value under $($Facts.HivePath). The corpus tool LOGS ON with the default profile, so with nothing named there it gets whatever MAPI decides - which is not something to find out during a 20,000-item build."
    }
    elseif ($Facts.ProfileNames -notcontains $Facts.DefaultProfile) {
        $names = '<none>'
        if ($null -ne $Facts.ProfileNames -and @($Facts.ProfileNames).Count -gt 0) { $names = (@($Facts.ProfileNames) -join ', ') }
        $problems += "DefaultProfile names '$($Facts.DefaultProfile)', and no profile by that name exists. This machine has: $names. Outlook has nothing to log on to."
    }
    elseif ($null -eq $Facts.AccountEntryCount) {
        $notes += "The account count for '$($Facts.DefaultProfile)' could not be read from $($Facts.AccountsKeyPath). NOT refused on: the tool's own COM count is the one that decides, and it prints it as 'profile accounts: N' when it vets the store. Read that line before you trust this run."
    }
    elseif ($Facts.AccountEntryCount -gt 0) {
        $plural = 'accounts'
        if ($Facts.AccountEntryCount -eq 1) { $plural = 'account' }
        $named = ''
        if ($null -ne $Facts.AccountEntryNames -and @($Facts.AccountEntryNames).Count -gt 0) {
            $named = " (" + ((@($Facts.AccountEntryNames) | ForEach-Object { "'$_'" }) -join ', ') + ")"
        }
        $problems += "The default profile is '$($Facts.DefaultProfile)', and it has $($Facts.AccountEntryCount) mail $plural$named under $($Facts.AccountsKeyPath) - data files and address books listed there are not counted. The corpus tool refuses ANY profile holding a mail account, with no override: a build creates unsent items in bulk, and a real one put 5,532 of them into the target store's Outbox - inert only because that profile could not send. Make the ACCOUNT-LESS profile the default."
    }

    # ---- precondition 2: Outlook is already running, and warm --------------------------------
    if ($Facts.OutlookProcessCount -lt 1) {
        $problems += "OUTLOOK.EXE is not running. The corpus tool COM-activates Outlook.Application, so with nothing running it starts a COLD Outlook inside its own STA thread - which blocks past the three-minute bound and returns 'FATAL: TimeoutException: Corpus STA operation timed out'."
    }
    else {
        if ($Facts.OutlookProcessCount -gt 1) {
            $notes += "$($Facts.OutlookProcessCount) OUTLOOK.EXE processes are running. COM will hand the tool one of them and nothing here says which. Not refused on, but if this build behaves oddly, that is the first thing to look at."
        }

        if ($null -eq $Facts.OutlookUptimeSeconds) {
            $notes += 'How long OUTLOOK.EXE has been up could not be read, so its warmth is unproven. NOT refused on - the cost of being wrong is one timed-out attempt, not a damaged mailbox.'
        }
        elseif ($Facts.OutlookUptimeSeconds -lt $Facts.WarmupSeconds) {
            $remaining = [int][math]::Ceiling($Facts.WarmupSeconds - $Facts.OutlookUptimeSeconds)
            $problems += "OUTLOOK.EXE has only been up $([int]$Facts.OutlookUptimeSeconds)s; this run wants $($Facts.WarmupSeconds)s. Wait about ${remaining}s more and run this again. 180s was used successfully and 75s was NOT always enough, which is why the bar is where it is; -WarmupSeconds moves it if you have a measurement that says otherwise."
        }

        if ($false -eq $Facts.OutlookResponding) {
            $notes += 'OUTLOOK.EXE is running but not responding to window messages. That is what a modal dialog looks like from outside - a profile prompt, for instance, which no COM call can answer. Not refused on, because a busy Outlook looks the same for a moment.'
        }

        if (-not [string]::IsNullOrWhiteSpace($Facts.OutlookCommandLine)) {
            $notes += "Running as: $($Facts.OutlookCommandLine)"
        }
    }

    # ---- precondition 3: nothing is waiting to rebuild the profile ---------------------------
    $importProblem = Test-CorpusImportPrfGuard -Facts $Facts
    if ($null -ne $importProblem) { $problems += $importProblem }

    return [pscustomobject]@{
        Ok       = ($problems.Count -eq 0)
        Problems = $problems
        Notes    = $notes
    }
}

<#
    The refusal, in full, with the remedy. Separate from the judging above so the text can be
    asserted without inventing a machine state to produce it.
#>
function Format-CorpusPreflightRefusal {
    param([string[]] $Problems)

    $text = @"
REFUSING TO BUILD. The corpus preconditions are not met on this machine.


"@

    $index = 1
    foreach ($problem in $Problems) {
        $text += "  $index. $problem`r`n`r`n"
        $index++
    }

    $text += @"
THE KNOWN-GOOD RECIPE, in this order. Steps 1 to 4 are none of them optional; step 0 applies only
when ImportPRF is one of the reasons above.

  0. Remove ImportPRF. Once the import it belongs to has run - Outlook started once since the
     profile script's -Execute - that script's own -Verify removes it:
         .\New-OutlookProfile.ps1 -Name <the profile> -Verify      (or .\New-TierProfile.ps1 -Verify)
     and this removes it unconditionally, whichever script wrote it:
         .\New-OutlookProfile.ps1 -ClearImportPrf -Execute

  1. Make the account-less profile the default:
         .\Set-DefaultOutlookProfile.ps1 -Name <the account-less profile> -Execute
     It refuses a profile that does not exist, and it refuses while Outlook is running - which is
     why it comes before the restart rather than after it.

  2. RESTART THE GUEST. That is the proven way to get a clean Outlook on the new default profile.
     NEVER taskkill OUTLOOK.EXE: mailbox-safety rule 7 forbids it outright, and a forced kill can
     leave the profile mid-write.

  3. Start Outlook and let it SETTLE. 180 seconds was used successfully; 75 seconds was not always
     enough. The tool COM-activates Outlook.Application, so a cold one starts inside its STA
     thread and blocks past the three-minute bound.

  4. Run this script again. It re-checks all three preconditions, and the tool then confirms the
     account count over COM and prints it as "profile accounts: 0" when it vets the store. That
     printed line is the authoritative one; this preflight is only the cheap early version of it.

Nothing in Outlook, the store, the manifest or the registry has been created, written or deleted -
only this run's own log, which every run starts afresh. -SkipPreflight skips preconditions 1 and 2
and NOT the ImportPRF check, so it weakens no safety guard: the tool's own store guard still
refuses a profile that can send, and the ImportPRF check still refuses a corpus that a re-import
could detach. It does hand back the five failed builds this preflight was written to prevent.
"@

    return $text
}

# =============================================================================================
# SELF-TEST. Pure: no registry, no processes, no tool, no Outlook. Runs anywhere.
# =============================================================================================

function New-PreflightFact {
    param(
        [string] $HivePath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook',
        [string] $DefaultProfile = 'CorpusProfile',
        [string[]] $ProfileNames = @('Outlook', 'CorpusProfile', 'TierProfile'),
        $AccountEntryCount = 0,
        [string[]] $AccountEntryNames = @(),
        [string] $AccountsKeyPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\CorpusProfile\9375CFF0413111d3B88A00104B2A6676',
        [int] $OutlookProcessCount = 1,
        $OutlookUptimeSeconds = 900,
        $OutlookResponding = $true,
        [string] $OutlookCommandLine = '',
        [int] $WarmupSeconds = 180,
        # '' = absent, $null = could not be read. Untyped on purpose: a [string] parameter would
        # turn $null into '' and the unreadable case could not be expressed at all.
        $ImportPrfValue = '',
        $FirstRunPresent = $true,
        [string] $SetupKeyPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup'
    )

    return [pscustomobject]@{
        HivePath             = $HivePath
        DefaultProfile       = $DefaultProfile
        ProfileNames         = $ProfileNames
        AccountEntryCount    = $AccountEntryCount
        AccountEntryNames    = $AccountEntryNames
        AccountsKeyPath      = $AccountsKeyPath
        OutlookProcessCount  = $OutlookProcessCount
        OutlookUptimeSeconds = $OutlookUptimeSeconds
        OutlookResponding    = $OutlookResponding
        OutlookCommandLine   = $OutlookCommandLine
        WarmupSeconds        = $WarmupSeconds
        ImportPrfValue       = $ImportPrfValue
        FirstRunPresent      = $FirstRunPresent
        SetupKeyPath         = $SetupKeyPath
    }
}

function New-HiveFact {
    param([string] $Version, [string[]] $ValueNames = @(), [string[]] $SubKeyNames = @())

    return [pscustomobject]@{
        Version     = $Version
        Path        = "$script:OfficeRootKeyPath\$Version\Outlook"
        ValueNames  = $ValueNames
        SubKeyNames = $SubKeyNames
    }
}

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    function Test-Case {
        param([string] $What, $Expected, $Actual)

        $script:SelfTestChecks++
        $expectedText = "$Expected"
        if ($null -eq $Expected) { $expectedText = '<null>' }
        $actualText = "$Actual"
        if ($null -eq $Actual) { $actualText = '<null>' }

        if ($expectedText -ceq $actualText) {
            Write-Host ("  OK   {0}" -f $What)
        }
        else {
            $script:SelfTestFailures += "$What : expected $expectedText, got $actualText"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $expectedText, $actualText)
        }
    }

    Write-Host 'Build-Corpus preflight self-test. No registry, no processes, no tool, no Outlook.'
    Write-Host ''
    Write-Host '== the hive shape rule (OfficeVersions.IsOutlookHive) =='

    Test-Case 'the installer footprint alone is not a hive' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('Resiliency'))
    Test-Case 'nor is it in another casing' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('RESILIENCY'))
    Test-Case 'an empty key is not a hive' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @())
    Test-Case 'nulls are not a hive, and do not throw' $false (Test-IsOutlookHive -ValueNames $null -SubKeyNames $null)
    Test-Case 'one value is enough' $true (Test-IsOutlookHive -ValueNames @('DefaultProfile') -SubKeyNames @('Resiliency'))
    Test-Case 'one non-footprint subkey is enough' $true (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('Profiles', 'Resiliency'))

    Write-Host ''
    Write-Host '== which hive gets read =='

    $real16 = New-HiveFact -Version '16.0' -ValueNames @('DefaultProfile') -SubKeyNames @('Profiles', 'Resiliency')
    $real17 = New-HiveFact -Version '17.0' -ValueNames @('DefaultProfile') -SubKeyNames @('Profiles')
    $real15 = New-HiveFact -Version '15.0' -ValueNames @() -SubKeyNames @('Profiles')
    $shell16 = New-HiveFact -Version '16.0' -SubKeyNames @('Resiliency')
    $shell17 = New-HiveFact -Version '17.0' -SubKeyNames @('Resiliency')

    Test-Case '16.0 wins over a real 17.0 (probe order, not highest number)' '16.0' (Select-OutlookHive -Candidates @($real17, $real16)).Version
    Test-Case '17.0 wins when 16.0 is only the footprint' '17.0' (Select-OutlookHive -Candidates @($shell16, $real17)).Version
    Test-Case '15.0 is read when it is the only real one' '15.0' (Select-OutlookHive -Candidates @($shell16, $shell17, $real15)).Version
    Test-Case 'all-footprint finds nothing' '<null>' (Select-OutlookHive -Candidates @($shell16, $shell17))
    Test-Case 'no candidates at all finds nothing' '<null>' (Select-OutlookHive -Candidates @())
    Test-Case 'a null candidate list finds nothing, and does not throw' '<null>' (Select-OutlookHive -Candidates $null)
    Test-Case '-OfficeVersion overrides the probe order' '15.0' (Select-OutlookHive -Candidates @($real16, $real15) -RequestedVersion '15.0').Version
    Test-Case '-OfficeVersion does NOT override the shape rule' '<null>' (Select-OutlookHive -Candidates @($real16, $shell17) -RequestedVersion '17.0')

    Write-Host ''
    Write-Host '== both preconditions met =='

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact)
    Test-Case 'a good machine passes' $true $verdict.Ok
    Test-Case 'with nothing to report' 0 $verdict.Problems.Count
    Test-Case 'and nothing to note' 0 $verdict.Notes.Count

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookUptimeSeconds 180)
    Test-Case 'uptime exactly at the bar is warm enough' $true $verdict.Ok

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -DefaultProfile 'corpusprofile')
    Test-Case 'a default differing only in case is not a refusal' $true $verdict.Ok

    Write-Host ''
    Write-Host '== precondition 1: the default profile =='

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -HivePath '')
    Test-Case 'no Outlook hive REFUSES' $false $verdict.Ok
    Test-Case 'and names the probe order' $true ($verdict.Problems[0] -like '*16.0, 17.0, 15.0*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -DefaultProfile '')
    Test-Case 'no DefaultProfile value REFUSES' $false $verdict.Ok
    Test-Case 'and says the tool logs on with it' $true ($verdict.Problems[0] -like '*LOGS ON with the default profile*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -DefaultProfile 'Vanished')
    Test-Case 'a default naming no profile REFUSES' $false $verdict.Ok
    Test-Case 'and lists the profiles that do exist' $true ($verdict.Problems[0] -like '*Outlook, CorpusProfile, TierProfile*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -ProfileNames @())
    Test-Case 'a machine with no profiles at all REFUSES' $false $verdict.Ok
    Test-Case 'and says so rather than listing nothing' $true ($verdict.Problems[0] -like '*This machine has: <none>*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -ProfileNames $null)
    Test-Case 'an unreadable Profiles key is that same refusal, not a crash' $false $verdict.Ok

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -AccountEntryCount 1)
    Test-Case 'one account on the default profile REFUSES' $false $verdict.Ok
    Test-Case 'and says "1 mail account", singular' $true ($verdict.Problems[0] -like '*1 mail account under*')
    Test-Case 'and says why there is no override' $true ($verdict.Problems[0] -like '*5,532*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -AccountEntryCount 1 -AccountEntryNames @('OutlookAI tier sink'))
    Test-Case 'and names the account when it can' $true ($verdict.Problems[0].Contains("1 mail account ('OutlookAI tier sink') under"))

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -AccountEntryCount 3)
    Test-Case 'three accounts pluralise' $true ($verdict.Problems[0] -like '*3 mail accounts under*')

    Write-Host ''
    Write-Host '== which account-manager entries are MAIL accounts (measured shapes, 2026-09-24) =='

    $wrapper = '{ED475414-B0D6-11D2-8C3B-00104B2A6676}'
    Test-Case "CorpusProfile's 'Outlook Data File' (MSUPST MS wrapper) is NOT an account" $false (Test-IsMailAccountEntry -Clsid $wrapper -ServiceName 'MSUPST MS')
    Test-Case "CorpusProfile's 'Outlook Address Book' (CONTAB wrapper) is NOT an account" $false (Test-IsMailAccountEntry -Clsid $wrapper -ServiceName 'CONTAB')
    Test-Case "the tier profile's POP3 account ({ED475411-...}) IS one" $true (Test-IsMailAccountEntry -Clsid '{ED475411-B0D6-11D2-8C3B-00104B2A6676}' -ServiceName '')
    Test-Case 'an ANSI PST wrapper is not one either' $false (Test-IsMailAccountEntry -Clsid $wrapper -ServiceName 'MSPST MS')
    Test-Case 'an Exchange mailbox (a wrapper around MSEMS) IS one' $true (Test-IsMailAccountEntry -Clsid $wrapper -ServiceName 'MSEMS')
    Test-Case 'a wrapper around a service nobody named IS one - fail closed' $true (Test-IsMailAccountEntry -Clsid $wrapper -ServiceName 'SOMETHINGNEW')
    Test-Case 'an entry with no readable clsid IS one - fail closed' $true (Test-IsMailAccountEntry -Clsid '' -ServiceName 'MSUPST MS')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -AccountEntryCount $null)
    Test-Case 'an unreadable account count does NOT refuse' $true $verdict.Ok
    Test-Case 'but it is noted' 1 $verdict.Notes.Count
    Test-Case 'and points at the COM count instead' $true ($verdict.Notes[0] -like "*profile accounts: N*")

    Write-Host ''
    Write-Host '== precondition 2: Outlook running and warm =='

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookProcessCount 0)
    Test-Case 'Outlook not running REFUSES' $false $verdict.Ok
    Test-Case 'and names the failure it prevents' $true ($verdict.Problems[0] -like '*Corpus STA operation timed out*')
    Test-Case 'and nothing is said about warmth when it is not running' 0 $verdict.Notes.Count

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookUptimeSeconds 40)
    Test-Case 'a cold Outlook REFUSES' $false $verdict.Ok
    Test-Case 'and says how much longer to wait' $true ($verdict.Problems[0] -like '*Wait about 140s more*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookUptimeSeconds 40 -WarmupSeconds 30)
    Test-Case '-WarmupSeconds lowers the bar' $true $verdict.Ok

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookUptimeSeconds $null)
    Test-Case 'an unreadable uptime does NOT refuse' $true $verdict.Ok
    Test-Case 'but it is noted' $true ($verdict.Notes[0] -like '*warmth is unproven*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookProcessCount 2)
    Test-Case 'two Outlooks do NOT refuse' $true $verdict.Ok
    Test-Case 'but are noted' $true ($verdict.Notes[0] -like '*2 OUTLOOK.EXE processes*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookResponding $false)
    Test-Case 'a non-responding Outlook does NOT refuse' $true $verdict.Ok
    Test-Case 'but is noted as what a modal dialog looks like' $true ($verdict.Notes[0] -like '*modal dialog*')

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -OutlookCommandLine '"C:\OUTLOOK.EXE" /PIM CorpusProfile')
    Test-Case 'the command line is reported' $true ($verdict.Notes[0] -like '*/PIM CorpusProfile*')

    Write-Host ''
    Write-Host '== precondition 3: nothing is waiting to rebuild the profile (ImportPRF) =='

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -ImportPrfValue '')
    Test-Case 'ImportPRF absent passes' $true $verdict.Ok

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -ImportPrfValue 'C:\OutlookAI-Profiles\CorpusProfile.prf' -FirstRunPresent $false)
    Test-Case 'ImportPRF set REFUSES' $false $verdict.Ok
    Test-Case 'as exactly one reason' 1 $verdict.Problems.Count
    Test-Case 'and names the value it found' $true ($verdict.Problems[0].Contains("'C:\OutlookAI-Profiles\CorpusProfile.prf'"))
    Test-Case 'and says the next start acts on it when First-Run is absent' $true ($verdict.Problems[0].Contains('the very next Outlook start acts on it'))
    Test-Case 'and says what a re-import does to a store' $true ($verdict.Problems[0].Contains('REBUILDS the profile'))
    Test-Case 'and gives the unconditional remedy' $true ($verdict.Problems[0].Contains('-ClearImportPrf -Execute'))

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -ImportPrfValue 'C:\OutlookAI-Profiles\CorpusProfile.prf' -FirstRunPresent $true)
    Test-Case 'with First-Run present it STILL refuses - latent is not gone' $false $verdict.Ok
    Test-Case 'and says it is latent' $true ($verdict.Problems[0].Contains('latent, not gone'))

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -ImportPrfValue $null)
    Test-Case 'an UNREADABLE ImportPRF refuses - fail-closed, unlike the other two' $false $verdict.Ok
    Test-Case 'and says why this one is different' $true ($verdict.Problems[0].Contains('nothing downstream re-checks ImportPRF'))

    $guard = Test-CorpusImportPrfGuard -Facts (New-PreflightFact -ImportPrfValue 'C:\x.prf' -OutlookProcessCount 0 -AccountEntryCount 5)
    Test-Case 'the guard alone - the -SkipPreflight path - refuses on its own' $true ($null -ne $guard)
    $guard = Test-CorpusImportPrfGuard -Facts (New-PreflightFact -ImportPrfValue '' -OutlookProcessCount 0 -AccountEntryCount 5)
    Test-Case 'and does NOT judge preconditions 1 and 2, which -SkipPreflight skips' '<null>' $guard

    Write-Host ''
    Write-Host '== all failing at once, and the refusal text =='

    $verdict = Test-CorpusPrecondition -Facts (New-PreflightFact -AccountEntryCount 1 -OutlookProcessCount 0)
    Test-Case 'both preconditions failing gives BOTH reasons' 2 $verdict.Problems.Count

    $verdict3 = Test-CorpusPrecondition -Facts (New-PreflightFact -AccountEntryCount 1 -OutlookProcessCount 0 -ImportPrfValue 'C:\x.prf')
    Test-Case 'all three failing gives THREE reasons' 3 $verdict3.Problems.Count

    $refusal = Format-CorpusPreflightRefusal -Problems $verdict.Problems
    Test-Case 'the refusal opens by refusing' $true ($refusal -like 'REFUSING TO BUILD.*')
    Test-Case 'it numbers the reasons' $true ($refusal -like '*  1. *' -and $refusal -like '*  2. *')
    Test-Case 'it names the profile script' $true ($refusal -like '*Set-DefaultOutlookProfile.ps1 -Name*')
    Test-Case 'it says restart the guest' $true ($refusal -like '*RESTART THE GUEST*')
    Test-Case 'it forbids taskkill' $true ($refusal -like '*NEVER taskkill OUTLOOK.EXE*')
    Test-Case 'it gives the settle time and the one that failed' $true ($refusal -like '*180 seconds was used successfully; 75 seconds was not always*')
    Test-Case 'it says what was NOT written, and the one thing that was' $true ($refusal.Contains('Nothing in Outlook, the store, the manifest or the registry has been created, written or deleted') -and $refusal.Contains("only this run's own log"))
    Test-Case 'it names the escape hatch' $true ($refusal -like '*-SkipPreflight*')
    Test-Case 'and says the escape hatch does not skip the ImportPRF check' $true ($refusal.Contains('NOT the ImportPRF check'))
    Test-Case 'it carries the ImportPRF step' $true ($refusal.Contains('-ClearImportPrf -Execute'))

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. These need a guest:'
    Write-Host '  * reading the hive, DefaultProfile, the Profiles subkeys and the account subkeys'
    Write-Host '  * reading ImportPRF and First-Run/FirstRun out of the Setup key'
    Write-Host '  * reading OUTLOOK.EXE out of the process list, with its start time and command line'
    Write-Host '  * every corpus verb - plan, probe, build, census - and the tool itself'

    if ($script:SelfTestFailures.Count -gt 0) {
        Write-Host ''
        foreach ($failure in $script:SelfTestFailures) { Write-Host "  $failure" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# EVERYTHING BELOW TOUCHES THE MACHINE.
# =============================================================================================

if (-not (Test-Path -LiteralPath $ToolsExe)) {
    throw @"
Tools not found at $ToolsExe
The guest has no .NET SDK, so nothing can be built here. Publish on the host and copy in:
    pwsh -File Testbed/host/Publish-GuestPayload.ps1
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <the guest> -Path .work\testbed-payload\Tools.zip -Destination C:\OutlookAI-Q5\Tools.zip
then on the guest: Expand-Archive C:\OutlookAI-Q5\Tools.zip -DestinationPath C:\OutlookAI-Q5\tools -Force
-VMName is mandatory: three guests coexist during the changeover and nothing guesses which.
"@
}

$identity = @('--corpus-id', $CorpusId, '--seed', "$Seed", '--anchor', $Anchor, '--count', "$Count")
$target = @('--store', $Store, '--allow-store', $Store)

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

function Invoke-Corpus {
    param([string] $Verb, [string[]] $Arguments, [switch] $Fatal)

    $line = "$Verb $($Arguments -join ' ')"
    Write-Host ''
    Write-Host "=== $line"
    Add-Content -LiteralPath $LogPath -Value "`r`n=== $line"

    # Output is captured and echoed rather than streamed, so a child that outlives the call
    # cannot hold the pipe open. Every number these verbs print is meant to be kept anyway.
    # Through Invoke-NativeCommand because this runs on the guests' Windows PowerShell 5.1: until
    # it did, one line on the tool's stderr - an unhandled exception's report, say - surfaced as a
    # NativeCommandError here, and the tool's own exit code and message were never read.
    $output = Invoke-NativeCommand { & $ToolsExe $Verb @Arguments 2>&1 }
    $code = $LASTEXITCODE
    $text = ($output | Out-String)
    Write-Host $text
    Add-Content -LiteralPath $LogPath -Value $text
    Add-Content -LiteralPath $LogPath -Value "exit $code"

    if ($code -ne 0 -and $Fatal) { throw "$Verb failed with exit $code. Log: $LogPath" }
    return $code
}

function Write-Preflight {
    param([string] $Text)

    Write-Host $Text
    Add-Content -LiteralPath $LogPath -Value $Text -ErrorAction SilentlyContinue
}

<#
    Gathers what Test-CorpusPrecondition judges. READ-ONLY, and it makes no COM call: a preflight
    that could itself start Outlook, or itself hang, would be the very fault it checks for.

    Anything unreadable comes back as $null rather than as a guess, and the judging above knows
    the difference between "no accounts" and "no accounts I could count".
#>
function Get-CorpusPreflightFact {
    param([string] $RequestedOfficeVersion, [int] $WarmupSecondsWanted)

    $candidates = @()
    if (Test-Path -LiteralPath $script:OfficeRootKeyPath) {
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
    }

    $hive = Select-OutlookHive -Candidates $candidates -RequestedVersion $RequestedOfficeVersion

    $hivePath = ''
    $defaultProfile = ''
    $profileNames = @()
    $accountsKeyPath = ''
    $accountEntryCount = $null
    $accountEntryNames = @()
    $setupKeyPath = ''
    $importPrfValue = $null
    $firstRunPresent = $null

    if ($null -ne $hive) {
        $hivePath = $hive.Path
        $defaultProfile = (Get-ItemProperty -Path $hivePath -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile

        # ImportPRF and First-Run. An absent Setup key is an ANSWER (nothing set), not a failure to
        # read one; only an exception leaves $importPrfValue at $null, which precondition 3 refuses.
        $setupKeyPath = Join-Path $hivePath 'Setup'
        try {
            $importPrfValue = ''
            $firstRunPresent = $false
            if (Test-Path -LiteralPath $setupKeyPath) {
                $setupKey = Get-Item -LiteralPath $setupKeyPath -ErrorAction Stop
                $raw = $setupKey.GetValue('ImportPRF', $null)
                if ($null -ne $raw) { $importPrfValue = [string] $raw }
                $firstRunPresent = ($null -ne $setupKey.GetValue('First-Run', $null)) -or ($null -ne $setupKey.GetValue('FirstRun', $null))
            }
        }
        catch {
            $importPrfValue = $null
            $firstRunPresent = $null
        }

        $profilesKeyPath = Join-Path $hivePath 'Profiles'
        if (Test-Path -LiteralPath $profilesKeyPath) {
            $profileNames = @(Get-ChildItem -LiteralPath $profilesKeyPath -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.PSChildName })
        }

        if (-not [string]::IsNullOrWhiteSpace($defaultProfile)) {
            $accountsKeyPath = Join-Path (Join-Path $profilesKeyPath $defaultProfile) $script:AccountsSubKeyName
            try {
                if (Test-Path -LiteralPath $accountsKeyPath) {
                    # MAIL accounts, not entries: see $script:ServiceWrapperClsid for why the two differ.
                    $accountEntryCount = 0
                    foreach ($entry in @(Get-ChildItem -LiteralPath $accountsKeyPath -ErrorAction Stop)) {
                        $clsid = [string] $entry.GetValue('clsid', '')
                        $serviceName = [string] $entry.GetValue('Service Name', '')
                        if (Test-IsMailAccountEntry -Clsid $clsid -ServiceName $serviceName) {
                            $accountEntryCount++
                            $accountEntryNames += [string] $entry.GetValue('Account Name', $entry.PSChildName)
                        }
                    }
                }
                else {
                    # The key is absent, which is not a failure to read: a profile that never had
                    # an account never had the container either. Measured 2026-09-15 on this
                    # project's guest - New-TierProfile.ps1 -Verify found NOTHING under it on a
                    # profile that had a PST and no internet account.
                    $accountEntryCount = 0
                }
            }
            catch {
                $accountEntryCount = $null
            }
        }
    }

    $processes = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    $uptime = $null
    $responding = $null
    $commandLine = ''

    if ($processes.Count -gt 0) {
        $oldest = $processes[0]
        foreach ($process in $processes) {
            try {
                if ($process.StartTime -lt $oldest.StartTime) { $oldest = $process }
            }
            catch {
                # A start time this session cannot read tells us nothing about which is oldest.
            }
        }

        try { $uptime = ((Get-Date) - $oldest.StartTime).TotalSeconds } catch { $uptime = $null }
        try { $responding = $oldest.Responding } catch { $responding = $null }
        try {
            $commandLine = (Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $($oldest.Id)" -ErrorAction Stop).CommandLine
        }
        catch {
            $commandLine = ''
        }
    }

    return [pscustomobject]@{
        HivePath             = $hivePath
        DefaultProfile       = $defaultProfile
        ProfileNames         = $profileNames
        AccountEntryCount    = $accountEntryCount
        AccountEntryNames    = $accountEntryNames
        AccountsKeyPath      = $accountsKeyPath
        OutlookProcessCount  = $processes.Count
        OutlookUptimeSeconds = $uptime
        OutlookResponding    = $responding
        OutlookCommandLine   = $commandLine
        WarmupSeconds        = $WarmupSecondsWanted
        ImportPrfValue       = $importPrfValue
        FirstRunPresent      = $firstRunPresent
        SetupKeyPath         = $setupKeyPath
    }
}

Remove-Item -LiteralPath $LogPath -Force -ErrorAction SilentlyContinue

# --- 0. the expectation sheet. Pure: no Outlook, runnable anywhere, including the host. -------
# Save this. Every measurement taken later is a ratio against one of these numbers, and computing
# them afterwards from the store is both slower and less trustworthy than reading them off the
# plan that produced it.
#
# It runs BEFORE the preflight on purpose: it is the one step that touches nothing, it is what the
# header tells an operator to read before building, and a machine that fails the preflight should
# still hand back the expectation sheet rather than nothing at all.
Invoke-Corpus -Verb 'corpus-plan' -Arguments $identity -Fatal | Out-Null

# --- 0b. the three preconditions. Everything after this line drives Outlook over COM, including
#         the dry run - corpus-build vets the store whether or not --execute is passed. ---------
$facts = Get-CorpusPreflightFact -RequestedOfficeVersion $OfficeVersion -WarmupSecondsWanted $WarmupSeconds
$importText = '(unreadable)'
if ($null -ne $facts.ImportPrfValue) {
    $importText = '<not set>'
    if ($facts.ImportPrfValue.Length -gt 0) { $importText = "'$($facts.ImportPrfValue)'" }
    $importText += "   First-Run/FirstRun present: $($facts.FirstRunPresent)"
}

if ($SkipPreflight) {
    # Preconditions 1 and 2 predict whether the build SUCCEEDS, and -SkipPreflight exists for a
    # machine they read wrongly. Precondition 3 protects the corpus, and nothing downstream
    # re-checks it - so it is not skipped. See Test-CorpusImportPrfGuard.
    Write-Preflight ''
    Write-Preflight '=== preflight SKIPPED (-SkipPreflight) for the default profile and warmth. The tool''s own'
    Write-Preflight '    store guard still applies, and the ImportPRF check below is NOT skipped.'
    Write-Preflight "  ImportPRF       : $importText"
    $guard = Test-CorpusImportPrfGuard -Facts $facts
    if ($null -ne $guard) {
        $refusal = Format-CorpusPreflightRefusal -Problems @($guard)
        Add-Content -LiteralPath $LogPath -Value $refusal -ErrorAction SilentlyContinue
        throw $refusal
    }
    Write-Preflight '  OK - nothing is waiting to rebuild the profile.'
}
else {
    Write-Preflight ''
    Write-Preflight '=== preflight: default profile, a warm Outlook, and nothing waiting to rebuild the profile'

    $uptimeText = '(unreadable)'
    if ($null -ne $facts.OutlookUptimeSeconds) { $uptimeText = "$([int]$facts.OutlookUptimeSeconds)s" }
    $accountsText = '(unreadable)'
    if ($null -ne $facts.AccountEntryCount) { $accountsText = "$($facts.AccountEntryCount)" }

    Write-Preflight "  hive            : $($facts.HivePath)"
    Write-Preflight "  DefaultProfile  : $($facts.DefaultProfile)"
    Write-Preflight "  mail accounts   : $accountsText   (registry, data-file and address-book entries excluded; the tool's COM count is the authoritative one)"
    Write-Preflight "  OUTLOOK.EXE     : $($facts.OutlookProcessCount) running, up $uptimeText, wanted $($facts.WarmupSeconds)s"
    Write-Preflight "  ImportPRF       : $importText"

    $verdict = Test-CorpusPrecondition -Facts $facts
    foreach ($note in $verdict.Notes) { Write-Preflight "  note: $note" }

    if (-not $verdict.Ok) {
        $refusal = Format-CorpusPreflightRefusal -Problems $verdict.Problems
        Add-Content -LiteralPath $LogPath -Value $refusal -ErrorAction SilentlyContinue
        throw $refusal
    }

    Write-Preflight '  OK - the default profile has no accounts, Outlook is up and warm, and ImportPRF is not set.'
}

if (-not $Execute) {
    Write-Host ''
    Write-Host 'Dry run. The probes are NOT run - both create items - and nothing is written.'
    Invoke-Corpus -Verb 'corpus-build' -Arguments ($target + $identity + @('--manifest', $Manifest)) | Out-Null
    Write-Host ''
    Write-Host 'Re-run with -Execute to probe and build. Read the plan above first: if the 7-day and'
    Write-Host '60-day window counts are close together, change the date bands BEFORE building, not after.'
    return
}

# --- 1. placement, then dates. Cheap; creates and deletes a handful of throwaway items. -------
if (-not $SkipProbe) {
    $probe = Invoke-Corpus -Verb 'corpus-probe' -Arguments ($target + $identity + @('--execute'))
    if ($probe -ne 0) {
        throw @"
The probe did not verify a rung (exit $probe). STOP HERE.
  * placement NOT ACHIEVABLE -> the corpus would be invisible to the freshness sweep, because
    Outlook files unsent items in Drafts and the sweep does not cover Drafts. Measurements taken
    against it would be measurements of an empty store. Docs/corpus-measurement-plan.md,
    'If placement fails', is the fallback route.
  * dates NOT ACHIEVABLE -> every step that mentions a window is void. See 'If the dates do not
    stick' in the same document.
Do not pass --allow-drafts-placement or --allow-undated to get past this without reading what
each one costs; the tool prints it.
"@
    }
}

# --- 2. build. Resumable and idempotent. ------------------------------------------------------
Invoke-Corpus -Verb 'corpus-build' `
    -Arguments ($target + $identity + @('--manifest', $Manifest, '--progress-every', "$ProgressEvery", '--execute')) `
    -Fatal | Out-Null

# --- 3. what actually landed. Read-only; the build runs this on itself, and it is repeated here
#        so an operator sees it as its own line rather than buried in build output. -------------
Invoke-Corpus -Verb 'corpus-census' -Arguments ($target + $identity + @('--manifest', $Manifest)) | Out-Null

Write-Host ''
Write-Host "Log: $LogPath"
Write-Host "Manifest: $Manifest"
Write-Host 'NOW COPY THE MANIFEST OFF THE GUEST. Without it the corpus cannot be torn down:'
Write-Host '    pwsh -File Testbed/host/Copy-FromGuest.ps1 -VMName <this guest>'
Write-Host '    (-VMName is mandatory - three guests coexist and nothing guesses which.)'
Write-Host 'Then take a checkpoint, and let Windows Search settle before any index measurement.'
