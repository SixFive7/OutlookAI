#Requires -Version 5.1
<#
    ============================================================================================
    RAN ON OutlookAI-Unindexed with -Execute before each of the first live runs there (2026-10-03,
    Docs/live-tier-on-the-vm.md section 4.1e): every time "verify : 1 store(s) named
    'throwaway@vm.invalid' ... Drafts designation NotFound, top-level folders [Deleted Items]" and
    READY - so PowerShell does surface MAPI_E_NOT_FOUND as the HResult this script reads. Once more
    after the last run, without a restore: it detached that run's store, attached a fresh one and
    deleted the old file, which a freshly started Outlook did not hold. -SelfTest covers the decisions.
    ============================================================================================

.SYNOPSIS
    Recreates the THROWAWAY data file - a small data file with NO Drafts folder - in the tier
    profile, before every live run. Guest only, in session 1. Q96 (iv), decided 2026-10-03.

.DESCRIPTION
    RUN THIS ON THE GUEST, IN SESSION 1, STRAIGHT AFTER THE HUB REBUILD AND BEFORE THE RUN:

        .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 900 -Script "& 'C:\OutlookAI-Q5\Reset-ThrowawayStore.ps1' -Execute"

    WHY IT EXISTS. Since Q85 a draft tool that files a draft in a mailbox with no Drafts folder lets
    Outlook create one and REPORTS it in createdFolders. The only store the live tier may write to
    was the hub, which has a Drafts folder, so on a guest the "created" branch never ran - every
    proof of it was against fakes. T2/LiveCreatedFolderTests proves it live: it replies to a post in
    THIS data file, whose Drafts folder the product must create and report, then discards the reply.
    The data file has to start every run without a Drafts folder, so it is recreated every run -
    a script step, the pattern the hub rebuild set (Reset-HubPopulation.ps1), not a paragraph
    somebody has to remember.

    WHAT IT IS. A data file attached with NameSpace.AddStoreEx, which holds Deleted Items and
    nothing else (Docs/live-tier-on-the-vm.md section 1.3), renamed to the guest's
    throwawayStoreDisplayName - read from the guest's own live-test settings, where the write
    allowlist grants it draft and delete and nothing else, and no census watches it. Its file is
    <Directory>\throwaway-<yyyyMMddTHHmmssZ>.pst: a NEW name every run, because Outlook can hold a
    detached data file open until it restarts, and a fixed name could not be recreated then.

    WHAT IT DOES, with -Execute:

      0. The guest guard (Assert-TestbedGuest), before anything touches the machine.
      1. Reads the settings and refuses on anything that does not agree: a profile that is not
         Portable (the maintainer's workstation is read-only for live tests, Q74), no
         throwawayStoreDisplayName, or one the live tier itself refuses - the hub, or a store in any
         other list. Refuses to run elevated: COM does not attach across integrity levels, and the
         tier profile's Outlook runs NOT elevated after the hub rebuild.
      2. Binds the running Outlook and requires it to be on the tier profile.
      3. Detaches every throwaway store of an earlier run - and ONLY those: a store whose file is
         in <Directory> and named throwaway-*.pst. A store carrying the throwaway's NAME whose file is
         anywhere else is not one this script made, and the run REFUSES rather than touch it.
      4. Attaches a fresh data file and renames it, through Add-OutlookPstStore.ps1 - the attach
         route measured on both guests - so its name is exactly the settings' one.
      5. Proves the result without asking Outlook for any folder, which could create one: exactly
         one store under that name, in the new file; no Drafts designation on the store
         (PR_IPM_DRAFTS_ENTRYID, MAPI_E_NOT_FOUND); no top-level folder called Drafts.
      6. Deletes the throwaway-*.pst files in <Directory> that no store holds any more. One that
         Outlook still has open is left, said so, and tried again next run.

    With -Remove -Execute: steps 0 to 3, then 6 - the store is gone after a run. Without -Execute:
    steps 0 to 2 and a report of what would happen; nothing detached, attached or deleted.

    WHAT IT NEVER DOES. Ask Outlook for a default folder (GetDefaultFolder creates a missing one -
    the very thing the proof depends on not having happened). Create, change or delete a mail item
    - mailbox-safety rule 1: the live test and its helpers do that. Touch a store it did not make.
    Quit, kill or start Outlook.

    WHAT IS TESTED WITHOUT A GUEST. -SelfTest exercises every decision here - the settings rules,
    the store and file selection, the file name, the verification - against synthetic inputs, and
    reads this script's own source for the calls it must never make. It reads no settings file,
    binds no Outlook and touches no file, and it runs on the build VM with every other self-test
    (Testbed/host/Invoke-TestsOnBuildVm.ps1, AGENTS.md). The guard, the COM session, the attach and
    the deletes are guest-only.

    STAGE BESIDE IT in C:\OutlookAI-Q5: OutlookMapiInterop.ps1 (the guard and the COM session),
    Add-OutlookPstStore.ps1 (the attach), and Register-InteractiveTask.ps1, which runs it.

    Windows PowerShell 5.1: no ternary, no '??'. Pure ASCII. It runs no native program.

.PARAMETER Execute
    Recreate (or, with -Remove, remove). Without it, a report of what would happen.

.PARAMETER Remove
    Detach the throwaway data file and delete its files - for leaving a guest without one.

.PARAMETER SelfTest
    Run the decision tests and exit. Touches nothing.

.PARAMETER SettingsPath
    The guest's live-test settings file, where Testbed/host/New-LiveTestSettings.ps1 prints the copy line to.

.PARAMETER Directory
    Where the throwaway data files live. Created if missing. Nothing outside it is ever deleted.

.PARAMETER TierProfileName
    The profile the live tier runs in, and so the one the throwaway data file is attached to.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.EXAMPLE
    .\Reset-ThrowawayStore.ps1 -SelfTest
    .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 900 -Script "& 'C:\OutlookAI-Q5\Reset-ThrowawayStore.ps1' -Execute"
    .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 900 -Script "& 'C:\OutlookAI-Q5\Reset-ThrowawayStore.ps1' -Remove -Execute"
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')] [switch] $Execute,
    [Parameter(ParameterSetName = 'Run')] [switch] $Remove,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest,
    [string] $SettingsPath = 'C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json',
    [string] $Directory = 'C:\OutlookAI-Tier\Throwaway',
    [string] $TierProfileName = 'OutlookAI-Tier',
    [string[]] $ExpectedUser = @('vmadmin')
)

$ErrorActionPreference = 'Stop'

$script:Invariant = [System.Globalization.CultureInfo]::InvariantCulture

# A throwaway data file's name: throwaway-<UTC stamp>.pst. Matched case-insensitively on the LEAF.
$script:FileLeafPattern = '(?i)^throwaway-\d{8}T\d{6}Z\.pst$'

# PR_IPM_DRAFTS_ENTRYID on the store object - where OutlookAI.Core's SpecialFolders.Resolve reads a
# data file's Drafts designation when it has no Inbox - and MAPI_E_NOT_FOUND, the answer for "none".
$script:DraftsEntryIdSchema = 'http://schemas.microsoft.com/mapi/proptag/0x36D70102'
$script:MapiENotFound = -2147221233

# The lists the live-test settings keep the throwaway out of - the loader refuses it in any of them.
$script:OtherLists = @('expectedStoreDisplayNames', 'indexedStoreDisplayNames', 'expectedDelegateStoreDisplayNames', 'bystanderStoreDisplayNames')

# =============================================================================================
# PURE DECISIONS. No file, no registry, no process, no Outlook, no output: -SelfTest decides
# everything below exactly as a run does.
# =============================================================================================

function Get-JsonField {
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Test-SameName {
    param([string] $Left, [string] $Right)
    return [string]::Equals($Left, $Right, [System.StringComparison]::OrdinalIgnoreCase)
}

<#
    What the guest's live-test settings say about the throwaway data file, read the way the loader
    reads them, with every reason the run must not go ahead.
#>
function Read-ThrowawaySettingsFact {
    param($Settings)
    $problems = New-Object System.Collections.Generic.List[string]
    $name = $null

    if ($null -eq $Settings -or -not ($Settings -is [System.Management.Automation.PSCustomObject])) {
        $problems.Add('the live-test settings are not a JSON object.')
        return [pscustomobject]@{ Name = $null; Problems = $problems.ToArray() }
    }

    $profileValue = Get-JsonField $Settings 'machineProfile'
    if (-not ($profileValue -is [string]) -or $profileValue -cne 'Portable') {
        $problems.Add("machineProfile is '$profileValue', not 'Portable'. The throwaway data file is a test guest's; a Production profile is the maintainer's workstation, which is read-only for live tests (Q74).")
    }

    $hub = Get-JsonField $Settings 'testHubStoreDisplayName'
    $value = Get-JsonField $Settings 'throwawayStoreDisplayName'
    if (-not ($value -is [string]) -or [string]::IsNullOrWhiteSpace($value)) {
        $problems.Add('throwawayStoreDisplayName is missing or blank, so there is no throwaway data file to recreate. Render the settings again with Testbed/host/New-LiveTestSettings.ps1, which requires it.')
        return [pscustomobject]@{ Name = $null; Problems = $problems.ToArray() }
    }

    if ($value -cne $value.Trim() -or $value.IndexOfAny([char[]]@('\', '/')) -ge 0) {
        $problems.Add("throwawayStoreDisplayName '$value' has leading or trailing whitespace or a slash; Outlook would never show a store under that name.")
    }

    if ($hub -is [string] -and (Test-SameName $value $hub)) {
        $problems.Add("throwawayStoreDisplayName '$value' is the hub. The throwaway data file exists because it has NO Drafts folder; the hub has one.")
    }

    foreach ($list in $script:OtherLists) {
        foreach ($store in @(Get-JsonField $Settings $list)) {
            if ($store -is [string] -and (Test-SameName $store $value)) {
                $problems.Add("throwawayStoreDisplayName '$value' is also in $list. The live tier refuses that; the throwaway data file is in no other list.")
            }
        }
    }

    if ($problems.Count -eq 0) { $name = $value }
    return [pscustomobject]@{ Name = $name; Problems = $problems.ToArray() }
}

<# True when a path names a throwaway data file this script makes: throwaway-<stamp>.pst, directly in $Directory. #>
function Test-ThrowawayFile {
    param([string] $Path, [string] $Directory)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $parent = [System.IO.Path]::GetDirectoryName($Path)
    $leaf = [System.IO.Path]::GetFileName($Path)
    return ($null -ne $parent) -and (Test-SameName $parent.TrimEnd('\') $Directory.TrimEnd('\')) -and ($leaf -match $script:FileLeafPattern)
}

<# The file a run attaches: a new name every run. #>
function New-ThrowawayFileName {
    param([datetime] $UtcNow)
    return 'throwaway-' + $UtcNow.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", $script:Invariant) + '.pst'
}

<#
    Which stores to detach, from rows of { DisplayName; FilePath }: every store whose file is one of
    this script's throwaway files - whatever it is called now - and nothing else. A store carrying the
    throwaway's NAME whose file is anything else is a problem: this script did not make it, and it
    is not touched.
#>
function Select-ThrowawayStore {
    param([object[]] $Stores, [string] $Name, [string] $Directory)
    $detach = New-Object System.Collections.Generic.List[object]
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($row in @($Stores)) {
        if ($null -eq $row) { continue }
        if (Test-ThrowawayFile -Path $row.FilePath -Directory $Directory) { $detach.Add($row); continue }
        if (Test-SameName $row.DisplayName $Name) {
            $problems.Add("A store named '$($row.DisplayName)' is attached from '$($row.FilePath)', which is not a throwaway-*.pst in $Directory. This script did not make it and will not touch it; find out what it is before running again.")
        }
    }
    return [pscustomobject]@{ Detach = $detach.ToArray(); Problems = $problems.ToArray() }
}

<# Which files to delete: throwaway-*.pst directly in $Directory that no attached store holds. #>
function Select-StaleThrowawayFile {
    param([string[]] $Files, [string[]] $AttachedPaths, [string] $Directory)
    $stale = New-Object System.Collections.Generic.List[string]
    foreach ($file in @($Files)) {
        if (-not (Test-ThrowawayFile -Path $file -Directory $Directory)) { continue }
        $held = $false
        foreach ($attached in @($AttachedPaths)) { if (Test-SameName $attached $file) { $held = $true } }
        if (-not $held) { $stale.Add($file) }
    }
    return , $stale.ToArray()
}

<#
    Whether the recreated store is what the proof needs, from what was read back WITHOUT asking
    Outlook for any folder: { Matching (stores under the name); FilePath; ExpectedPath;
    DraftsDesignation ('NotFound', 'Found' or 'Failed'); TopLevelNames }.
#>
function Get-VerificationProblem {
    param($Facts)
    $problems = New-Object System.Collections.Generic.List[string]
    if ($Facts.Matching -ne 1) {
        $problems.Add("$($Facts.Matching) store(s) carry the throwaway's name, where exactly one must.")
    }
    elseif (-not (Test-SameName $Facts.FilePath $Facts.ExpectedPath)) {
        $problems.Add("The store under the throwaway's name is attached from '$($Facts.FilePath)', not the new file '$($Facts.ExpectedPath)'.")
    }

    if ($Facts.DraftsDesignation -ceq 'Found') {
        $problems.Add('The new data file already carries a Drafts designation (PR_IPM_DRAFTS_ENTRYID), so the created-folder proof would find a Drafts folder and prove nothing.')
    }
    elseif ($Facts.DraftsDesignation -cne 'NotFound') {
        $problems.Add('Whether the new data file carries a Drafts designation could not be read; the proof would refuse it the same way. Run again.')
    }

    foreach ($folder in @($Facts.TopLevelNames)) {
        if ($folder -is [string] -and (Test-SameName $folder 'Drafts')) {
            $problems.Add("The new data file already has a top-level folder called '$folder'.")
        }
    }
    return , $problems.ToArray()
}

<# The designation read's answer, from what the read threw (or $null when it returned a value). #>
function ConvertTo-DesignationState {
    param($Failure)
    if ($null -eq $Failure) { return 'Found' }
    $current = $Failure
    while ($null -ne $current) {
        if ($current.HResult -eq $script:MapiENotFound) { return 'NotFound' }
        $current = $current.InnerException
    }
    return 'Failed'
}

# =============================================================================================
# SELF-TEST
# =============================================================================================

function ConvertTo-SelfTestText {
    param($Value)
    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Array]) { return '[' + ((@($Value) | ForEach-Object { ConvertTo-SelfTestText $_ }) -join ', ') + ']' }
    return [string] $Value
}

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    # Compared with -ceq and searched with String.Contains. Never -like (Testbed/README.md section 4b).
    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:SelfTestChecks++
        $expectedText = ConvertTo-SelfTestText $Expected
        $actualText = ConvertTo-SelfTestText $Actual
        if ($expectedText -ceq $actualText) { Write-Host ("  OK   {0}" -f $What) }
        else {
            $script:SelfTestFailures += "$What : expected $expectedText, got $actualText"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $expectedText, $actualText)
        }
    }

    function Test-HasProblem {
        param([string] $What, [string[]] $Problems, [string] $Needle)
        $hit = $false
        foreach ($problem in @($Problems)) { if ($null -ne $problem -and $problem.Contains($Needle)) { $hit = $true } }
        if (-not $hit) { Write-Host ("       (problems were: {0})" -f (@($Problems) -join ' | ')) }
        Test-Case $What $true $hit
    }

    function New-Settings {
        param([string] $MachineProfile = 'Portable', $Throwaway = 'throwaway@vm.invalid')
        $json = '{ "machineProfile": "' + $MachineProfile + '", "testHubStoreDisplayName": "tier@vm.invalid", ' +
            '"expectedStoreDisplayNames": ["tier@vm.invalid", "bystander@vm.invalid", "identity@vm.invalid"], ' +
            '"indexedStoreDisplayNames": [], "expectedDelegateStoreDisplayNames": [], ' +
            '"bystanderStoreDisplayNames": ["bystander@vm.invalid"] }'
        $settings = $json | ConvertFrom-Json
        if ($null -ne $Throwaway) { $settings | Add-Member -NotePropertyName 'throwawayStoreDisplayName' -NotePropertyValue $Throwaway }
        return $settings
    }

    $dir = 'C:\OutlookAI-Tier\Throwaway'
    Write-Host 'Reset-ThrowawayStore self-test. No settings file, no Outlook, no file touched.'

    Write-Host ''
    Write-Host '== the settings =='
    $fact = Read-ThrowawaySettingsFact (New-Settings)
    Test-Case 'a guest''s settings name the throwaway' 'throwaway@vm.invalid' $fact.Name
    Test-Case 'with no problem' 0 $fact.Problems.Count
    Test-HasProblem 'a Production profile is refused' (Read-ThrowawaySettingsFact (New-Settings -MachineProfile 'Production')).Problems 'read-only for live tests'
    Test-HasProblem 'a profile spelt in the wrong case is refused' (Read-ThrowawaySettingsFact (New-Settings -MachineProfile 'portable')).Problems "not 'Portable'"
    Test-HasProblem 'no throwaway is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway $null)).Problems 'throwawayStoreDisplayName is missing or blank'
    Test-HasProblem 'a blank throwaway is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway ' ')).Problems 'missing or blank'
    Test-HasProblem 'a non-string throwaway is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 7)).Problems 'missing or blank'
    Test-HasProblem 'the hub as the throwaway is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 'TIER@vm.invalid')).Problems 'is the hub'
    Test-HasProblem 'a watched store as the throwaway is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 'identity@vm.invalid')).Problems 'is also in expectedStoreDisplayNames'
    Test-HasProblem 'a bystander as the throwaway is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 'bystander@vm.invalid')).Problems 'is also in bystanderStoreDisplayNames'
    Test-HasProblem 'a name with a slash is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 'a/b')).Problems 'a slash'
    Test-HasProblem 'a name with trailing space is refused' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 'throwaway@vm.invalid ')).Problems 'whitespace'
    Test-Case 'and a refused name is never handed on' '<null>' (Read-ThrowawaySettingsFact (New-Settings -Throwaway 'TIER@vm.invalid')).Name
    Test-HasProblem 'something that is not an object is refused' (Read-ThrowawaySettingsFact 'text').Problems 'not a JSON object'

    Write-Host ''
    Write-Host '== the file name =='
    $stamp = New-ThrowawayFileName -UtcNow ([datetime]::SpecifyKind([datetime]'2026-10-03 05:07:09', [System.DateTimeKind]::Utc))
    Test-Case 'a run''s file is named by its UTC second' 'throwaway-20261003T050709Z.pst' $stamp
    Test-Case 'which is a throwaway file in the directory' $true (Test-ThrowawayFile -Path (Join-Path $dir $stamp) -Directory $dir)
    Test-Case 'in any case' $true (Test-ThrowawayFile -Path 'c:\outlookai-tier\throwaway\THROWAWAY-20261003T050709Z.PST' -Directory $dir)
    Test-Case 'but not in another directory' $false (Test-ThrowawayFile -Path ('C:\OutlookAI-Tier\' + $stamp) -Directory $dir)
    Test-Case 'nor in a directory below it' $false (Test-ThrowawayFile -Path (Join-Path $dir ('sub\' + $stamp)) -Directory $dir)
    Test-Case 'nor the tier''s own data file' $false (Test-ThrowawayFile -Path 'C:\OutlookAI-Tier\Outlook.pst' -Directory $dir)
    Test-Case 'nor a file with another name in the directory' $false (Test-ThrowawayFile -Path (Join-Path $dir 'bystander.pst') -Directory $dir)
    Test-Case 'nor a near miss' $false (Test-ThrowawayFile -Path (Join-Path $dir 'throwaway-20261003T050709Z.pst.bak') -Directory $dir)
    Test-Case 'nor nothing' $false (Test-ThrowawayFile -Path $null -Directory $dir)

    Write-Host ''
    Write-Host '== which stores are detached =='
    $old = Join-Path $dir 'throwaway-20261002T010203Z.pst'
    $rows = @(
        [pscustomobject]@{ DisplayName = 'tier@vm.invalid'; FilePath = 'C:\OutlookAI-Tier\Outlook.pst' },
        [pscustomobject]@{ DisplayName = 'throwaway@vm.invalid'; FilePath = $old },
        [pscustomobject]@{ DisplayName = 'Outlook Data File'; FilePath = (Join-Path $dir 'throwaway-20261001T010203Z.pst') },
        [pscustomobject]@{ DisplayName = 'bystander@vm.invalid'; FilePath = 'C:\OutlookAI-Q5\pst\bystander.pst' },
        [pscustomobject]@{ DisplayName = 'Exchange'; FilePath = $null }
    )
    $selection = Select-ThrowawayStore -Stores $rows -Name 'throwaway@vm.invalid' -Directory $dir
    Test-Case 'the earlier throwaway is detached, and one this script made under another name too' (($old, (Join-Path $dir 'throwaway-20261001T010203Z.pst')) -join '|') ((@($selection.Detach) | ForEach-Object { $_.FilePath }) -join '|')
    Test-Case 'and nothing else is touched' 0 $selection.Problems.Count
    $foreign = @($rows) + [pscustomobject]@{ DisplayName = 'THROWAWAY@vm.invalid'; FilePath = 'D:\elsewhere\real.pst' }
    $refused = Select-ThrowawayStore -Stores $foreign -Name 'throwaway@vm.invalid' -Directory $dir
    Test-HasProblem 'a store with the name but not this script''s file is refused, never detached' $refused.Problems 'will not touch it'
    Test-Case 'and is not on the detach list' $false (@($refused.Detach | ForEach-Object { $_.FilePath }) -contains 'D:\elsewhere\real.pst')
    Test-Case 'an empty profile detaches nothing' 0 @((Select-ThrowawayStore -Stores @() -Name 'throwaway@vm.invalid' -Directory $dir).Detach).Count

    Write-Host ''
    Write-Host '== which files are deleted =='
    $files = @($old, (Join-Path $dir 'throwaway-20261001T010203Z.pst'), (Join-Path $dir 'notes.txt'), 'C:\OutlookAI-Tier\Outlook.pst')
    $stale = Select-StaleThrowawayFile -Files $files -AttachedPaths @((Join-Path $dir 'throwaway-20261001T010203Z.pst')) -Directory $dir
    Test-Case 'only throwaway files no store holds' $old ($stale -join '|')
    Test-Case 'nothing when every one is held' 0 (Select-StaleThrowawayFile -Files @($old) -AttachedPaths @($old.ToUpperInvariant()) -Directory $dir).Count
    Test-Case 'never a file outside the directory' 0 (Select-StaleThrowawayFile -Files @('C:\OutlookAI-Tier\throwaway-20261001T010203Z.pst') -AttachedPaths @() -Directory $dir).Count

    Write-Host ''
    Write-Host '== the verification =='
    $good = [pscustomobject]@{ Matching = 1; FilePath = $old; ExpectedPath = $old; DraftsDesignation = 'NotFound'; TopLevelNames = @('Deleted Items', 'Search Folders') }
    Test-Case 'a fresh data file passes' 0 (Get-VerificationProblem $good).Count
    $twice = [pscustomobject]@{ Matching = 2; FilePath = $old; ExpectedPath = $old; DraftsDesignation = 'NotFound'; TopLevelNames = @() }
    Test-HasProblem 'two stores under the name are refused' (Get-VerificationProblem $twice) 'where exactly one must'
    $wrong = [pscustomobject]@{ Matching = 1; FilePath = 'C:\x.pst'; ExpectedPath = $old; DraftsDesignation = 'NotFound'; TopLevelNames = @() }
    Test-HasProblem 'the name on another file is refused' (Get-VerificationProblem $wrong) 'not the new file'
    $designated = [pscustomobject]@{ Matching = 1; FilePath = $old; ExpectedPath = $old; DraftsDesignation = 'Found'; TopLevelNames = @() }
    Test-HasProblem 'a Drafts designation is refused' (Get-VerificationProblem $designated) 'already carries a Drafts designation'
    $unread = [pscustomobject]@{ Matching = 1; FilePath = $old; ExpectedPath = $old; DraftsDesignation = 'Failed'; TopLevelNames = @() }
    Test-HasProblem 'an unreadable designation is refused, never read as absent' (Get-VerificationProblem $unread) 'could not be read'
    $folder = [pscustomobject]@{ Matching = 1; FilePath = $old; ExpectedPath = $old; DraftsDesignation = 'NotFound'; TopLevelNames = @('Deleted Items', 'DRAFTS') }
    Test-HasProblem 'a top-level Drafts folder is refused' (Get-VerificationProblem $folder) "folder called 'DRAFTS'"

    $notFound = New-Object System.Runtime.InteropServices.COMException('The property cannot be found.', $script:MapiENotFound)
    Test-Case 'MAPI_E_NOT_FOUND reads as no designation' 'NotFound' (ConvertTo-DesignationState $notFound)
    Test-Case 'and so does one wrapped by the binder' 'NotFound' (ConvertTo-DesignationState (New-Object System.Management.Automation.MethodInvocationException('wrapped', $notFound)))
    Test-Case 'any other failure is a failure' 'Failed' (ConvertTo-DesignationState (New-Object System.Runtime.InteropServices.COMException('Failed.', -2147467259)))
    Test-Case 'a value that came back is a designation' 'Found' (ConvertTo-DesignationState $null)

    Write-Host ''
    Write-Host '== what this file never calls =='
    $sourcePath = $PSCommandPath
    if ([string]::IsNullOrEmpty($sourcePath)) { $sourcePath = $MyInvocation.ScriptName }
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$null, [ref]$null)
    $members = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true) | ForEach-Object { $_.Member.Extent.Text.Trim('''', '"') })
    $commands = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
    foreach ($forbidden in @('GetDefaultFolder', 'GetDefaultFolderReportingCreation', 'CreateItem', 'Delete', 'PermanentlyDelete', 'Move', 'Save', 'Send', 'Quit', 'Kill')) {
        Test-Case "no .$forbidden() call - no folder asked for, no item touched, no Outlook ended" $false ($members -contains $forbidden)
    }
    # .Add() is the one name an item collection and a .NET list share, so it is judged by what it is
    # called on: the three lists the decisions above fill, and nothing else - never Items, Folders or Stores.
    $adds = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -eq 'Add' }, $true) | ForEach-Object { $_.Expression.Extent.Text })
    Test-Case '.Add() is called only on the decisions'' own lists' '' ((@($adds | Where-Object { @('$detach', '$problems', '$stale') -notcontains $_ }) | Sort-Object -Unique) -join ', ')
    foreach ($forbidden in @('Stop-Process', 'taskkill', 'taskkill.exe', 'Start-Process')) {
        Test-Case "no $forbidden" $false ($commands -contains $forbidden)
    }
    $removeStoreCalls = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -eq 'RemoveStore' }, $true))
    $removeItemCalls = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Remove-Item' }, $true))
    Test-Case 'RemoveStore is called in one place' 1 $removeStoreCalls.Count
    Test-Case 'and only on a store the selection chose' $true ($removeStoreCalls.Count -eq 1 -and $removeStoreCalls[0].Extent.Text.Contains('$row.Store'))
    Test-Case 'Remove-Item is called in one place' 1 $removeItemCalls.Count
    Test-Case 'and only on a file the selection chose, literally' $true ($removeItemCalls.Count -eq 1 -and $removeItemCalls[0].Extent.Text.Contains('-LiteralPath $file'))

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host 'NOT exercised here - guest-only: the guard, the elevation check, the COM session, RemoveStore,'
    Write-Host 'Add-OutlookPstStore.ps1''s attach and rename, the designation read and the file deletes.'
    if ($script:SelfTestFailures.Count -gt 0) {
        foreach ($failure in $script:SelfTestFailures) { Write-Host "  $failure" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# EVERYTHING BELOW TOUCHES THE MACHINE. Guest only, session 1 only.
# =============================================================================================

. "$PSScriptRoot\OutlookMapiInterop.ps1"

Assert-TestbedGuest -ExpectedUser $ExpectedUser

# The tier profile's Outlook runs NOT elevated after the hub rebuild, and COM does not attach across
# integrity levels: an elevated run would reach a different Outlook, or start one.
$principal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'REFUSING: this runs ELEVATED. Run it at Outlook''s integrity level - Register-InteractiveTask.ps1 -RunLevel Limited - straight after Reset-HubPopulation.ps1, which leaves the tier profile''s Outlook running NOT elevated.'
}

# Binding Outlook over COM STARTS one when none is up, on whatever profile is the default. This
# script never starts Outlook: the tier profile's is left running by the hub rebuild, so none
# running means that has not happened, and the order of the steps is wrong.
if (@(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue).Count -eq 0) {
    throw 'REFUSING: Outlook is not running, and binding it would start one. Run Reset-HubPopulation.ps1 first - it leaves the tier profile''s Outlook running, NOT elevated - then this, then the tier.'
}

if (-not (Test-Path -LiteralPath $SettingsPath)) {
    throw "REFUSING: no live-test settings at $SettingsPath. Render them on the host with Testbed/host/New-LiveTestSettings.ps1 -VMName <this guest> and copy them in with the line it prints."
}
$fact = Read-ThrowawaySettingsFact (Get-Content -LiteralPath $SettingsPath -Raw | ConvertFrom-Json)
if ($fact.Problems.Count -gt 0) {
    throw ("REFUSING: the live-test settings at $SettingsPath do not describe a throwaway data file this script may recreate:`n  - " + ($fact.Problems -join "`n  - "))
}
$name = $fact.Name
$Directory = [System.IO.Path]::GetFullPath($Directory)
$mode = 'recreate'
if ($Remove) { $mode = 'remove' }
Write-Host "throwaway data file : $name"
Write-Host "directory           : $Directory"
Write-Host "profile             : $TierProfileName"
Write-Host "mode                : $mode$(if (-not $Execute) { ' (dry run - nothing changes)' })"
Write-Host ''

# Steps 2 and 3: bind, check the profile, pick and detach the earlier throwaway stores.
$script:Attached = @()
Invoke-WithOutlookSession -Body {
    param($ns)
    if ($ns.CurrentProfileName -ne $TierProfileName) {
        throw "REFUSING: Outlook is on profile '$($ns.CurrentProfileName)', not '$TierProfileName'. Run this straight after Reset-HubPopulation.ps1, which leaves the tier profile's Outlook running."
    }

    $rows = @()
    for ($index = 1; $index -le $ns.Stores.Count; $index++) {
        $store = $ns.Stores.Item($index)
        $filePath = $null
        try { $filePath = $store.FilePath } catch { $filePath = $null }
        $rows += [pscustomobject]@{ DisplayName = $store.DisplayName; FilePath = $filePath; Store = $store }
    }

    $selection = Select-ThrowawayStore -Stores $rows -Name $name -Directory $Directory
    if ($selection.Problems.Count -gt 0) { throw ("REFUSING:`n  - " + ($selection.Problems -join "`n  - ")) }
    foreach ($row in @($selection.Detach)) {
        Write-Host ("detach              : '{0}' <- {1}" -f $row.DisplayName, $row.FilePath)
        if ($Execute) { $ns.RemoveStore($row.Store.GetRootFolder()) }
    }
    if (@($selection.Detach).Count -eq 0) { Write-Host 'detach              : none - no earlier throwaway data file is attached' }

    $script:Attached = @($rows | Where-Object { $null -ne $_.FilePath -and @($selection.Detach) -notcontains $_ } | ForEach-Object { [string]$_.FilePath })
}

# Step 4: a fresh data file, under the settings' name.
$newPath = $null
if (-not $Remove) {
    if (-not (Test-Path -LiteralPath $Directory)) {
        Write-Host "create directory    : $Directory"
        if ($Execute) { [void](New-Item -ItemType Directory -Path $Directory) }
    }
    $newPath = Join-Path $Directory (New-ThrowawayFileName -UtcNow ([datetime]::UtcNow))
    Write-Host "attach              : $newPath, renamed '$name'"
    if ($Execute) {
        & (Join-Path $PSScriptRoot 'Add-OutlookPstStore.ps1') -ProfileName $TierProfileName -DisplayName $name -Path $newPath -ExpectedUser $ExpectedUser -Execute
        $script:Attached += $newPath
    }
}

# Step 5: prove it, without asking Outlook for a folder.
if ($Execute -and -not $Remove) {
    $script:Facts = $null
    Invoke-WithOutlookSession -Body {
        param($ns)
        $matching = @()
        for ($index = 1; $index -le $ns.Stores.Count; $index++) {
            $store = $ns.Stores.Item($index)
            if (Test-SameName $store.DisplayName $name) { $matching += $store }
        }
        $filePath = $null
        $designation = 'Failed'
        $topLevel = @()
        if ($matching.Count -eq 1) {
            $store = $matching[0]
            try { $filePath = $store.FilePath } catch { $filePath = $null }
            $failure = $null
            try { [void]$store.PropertyAccessor.GetProperty($script:DraftsEntryIdSchema) } catch { $failure = $_.Exception }
            $designation = ConvertTo-DesignationState $failure
            $root = $store.GetRootFolder()
            for ($index = 1; $index -le $root.Folders.Count; $index++) { $topLevel += [string]$root.Folders.Item($index).Name }
        }
        $script:Facts = [pscustomobject]@{ Matching = $matching.Count; FilePath = $filePath; ExpectedPath = $newPath; DraftsDesignation = $designation; TopLevelNames = $topLevel }
    }
    Write-Host ("verify              : {0} store(s) named '{1}', file {2}, Drafts designation {3}, top-level folders [{4}]" -f $script:Facts.Matching, $name, $script:Facts.FilePath, $script:Facts.DraftsDesignation, ($script:Facts.TopLevelNames -join ', '))
    $problems = Get-VerificationProblem $script:Facts
    if ($problems.Count -gt 0) { throw ("The recreated throwaway data file is not what the created-folder proof needs:`n  - " + ($problems -join "`n  - ")) }
}

# Step 6: the files no store holds any more.
if (Test-Path -LiteralPath $Directory) {
    $present = @(Get-ChildItem -LiteralPath $Directory -File -Filter 'throwaway-*.pst' | ForEach-Object { $_.FullName })
    foreach ($file in (Select-StaleThrowawayFile -Files $present -AttachedPaths $script:Attached -Directory $Directory)) {
        Write-Host "delete              : $file"
        if ($Execute) {
            try { Remove-Item -LiteralPath $file -Force }
            catch { Write-Host "  left in place - Outlook still holds it, most likely until it restarts; the next run tries again ($($_.Exception.GetType().Name))" }
        }
    }
}

Write-Host ''
if (-not $Execute) { Write-Host 'Dry run. Nothing detached, attached or deleted. Re-run with -Execute.' }
elseif ($Remove) { Write-Host "Removed: no throwaway data file is attached to '$TierProfileName'." }
else { Write-Host "READY: '$name' is attached to '$TierProfileName' from $newPath, with no Drafts folder. Run the tier next (Testbed/README.md section 4c)." }
