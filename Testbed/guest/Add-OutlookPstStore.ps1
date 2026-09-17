#Requires -Version 5.1
<#
    ============================================================================================
    THE MAPI ROUTE IS DEAD ON THIS BUILD. THIS SCRIPT USES AddStoreEx PLUS A ROOT-FOLDER RENAME.
    ============================================================================================

    WHAT THIS SCRIPT USED TO BE, AND WHY IT NEVER WORKED. Until now it added the PST through
    Extended MAPI - `IMsgServiceAdmin::CreateMsgService('MSUPST MS')` followed by
    `ConfigureMsgService` carrying `PR_DISPLAY_NAME` - because that route names the store AT
    CREATION TIME, and the object model was believed unable to name a store at all. It carried a
    banner saying it had never been executed. The interop it stood on HAS now been executed, by a
    sibling script, on a guest, 2026-09-16, and its very first call failed:

        Exception calling "MAPIAdminProfiles" with "2" argument(s): "Unable to cast COM object of
        type 'System.__ComObject' to interface type 'OutlookAI.Testbed.IProfAdmin' ... No such
        interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE))."

    on guest OAI-UNINDEXED, user vmadmin, 64-bit elevated Windows PowerShell 5.1, Office LTSC
    2024 (ProPlus2024Volume / PerpetualVL2024) build 16.0.17932.20884. All three of this script's
    MAPI call sites sat behind that one, so all three were dead.

    ============================================================================================
    WHAT IT DOES NOW, AND WHY THE OBJECTION TO THIS ROUTE NO LONGER HOLDS
    ============================================================================================

    Two calls, and the second is the one that used to be impossible:

      1. `NameSpace.AddStoreEx(<path>, olStoreUnicode)` attaches the PST, creating the file when
         it is not there. MS-documented. It cannot name the store - it has never been able to.
      2. The store's ROOT FOLDER is renamed, and `Store.DisplayName` FOLLOWS.

    STEP 2 IS WHY THIS SCRIPT EXISTS IN THIS SHAPE, and it is measured rather than hoped for.
    `Store.DisplayName` is read-only; `PropertyAccessor.SetProperty` on `0x3001001F` at store
    level is reported blocked; and whether a root-folder rename carries through to
    `Store.DisplayName` is a question FOUR research passes could not settle in any documentation,
    Microsoft's or the community's. Testbed/guest/Rename-OutlookStore.ps1 settled it empirically
    on 2026-09-15, on Office LTSC 2024 build 16.0.17932:

        set root.Name        was 'Outlook Data File', now 'tier@vm.invalid'
        Store.DisplayName    'tier@vm.invalid'
        account afterwards   SmtpAddress='tier@vm.invalid' DeliveryStore='tier@vm.invalid'

    So the rename carries, the `@` survives, and it does not break an account's delivery-store
    binding. The old banner's central claim - "AddStoreEx ... CANNOT NAME IT" - is still true and
    no longer matters, because naming it afterwards works.

    ============================================================================================
    -NameProbe IS GONE, BECAUSE THE QUESTION IT EXISTED TO ASK IS ANSWERED
    ============================================================================================

    The probe created a throwaway profile, added a PST called `test@vm.invalid`, reported
    ACCEPTED / REJECTED / TRANSFORMED, and deleted the profile again. It was pointed at
    Testbed/README.md section 6 item 10 and Docs/live-tier-on-the-vm.md section 8 item 2: DOES
    OUTLOOK ACCEPT '@' IN A STORE DISPLAY NAME? It gates a whole draft family, because several
    tests hand a store's display name straight to NewDraft as an address.

    THE ANSWER IS YES, and it is measured twice, independently, on this exact Office build:

      * through a .prf - `Store.DisplayName` read back OVER COM as `tier@vm.invalid`, the exact
        string the file's `[Service1] Name=` carried (2026-09-15, both guests);
      * through the root-folder rename above, which is this script's own route.

    So the probe would now be asking a question it already has two answers to, in a throwaway
    profile - and DELETING A PROFILE HAS NO FREE ROUTE LEFT. `IProfAdmin::DeleteProfile` died
    with the rest of the interface; removing the profile's registry key is reverse engineering
    with no published recipe, leaves the PST behind and can leave `DefaultProfile` naming nothing.
    A probe that cannot clean up after itself is worse than no probe. It is removed rather than
    left half-working, and the capability it needed is recorded as HAVING NO ROUTE in
    OutlookMapiInterop.ps1's banner rather than quietly dropped.

    ============================================================================================
    THE THREE THINGS THAT ARE DIFFERENT ABOUT RUNNING THIS, AND THEY ALL BITE
    ============================================================================================

    1. OUTLOOK MUST BE RUNNING - the exact opposite of what this script used to demand. The
       object model is the route now, so there has to be a live Outlook to drive. It binds one,
       starting one if none is up, and it NEVER quits it: mailbox-safety rule 7 forbids taskkill
       outright and wants COM references released before any quit. Leaving it headless is also
       what Build-Corpus.ps1 wants next - that script refuses unless Outlook is already warm.

    2. IT OPERATES ON THE DEFAULT PROFILE, and refuses if that is not the one you named. MAPI is
       initialised with `GetDefaultFolder`, not `NameSpace.Logon`, because Logon is documented to
       be able to raise the profile picker even when a default is set - and a dialog on an
       unattended guest is a hang, not a prompt. Rename-OutlookStore.ps1 does the same and says
       so. Point the default where you mean first: `Set-DefaultOutlookProfile.ps1`.

    3. AddStoreEx HAS BEEN SEEN TO SPIN ON THIS PROJECT'S GUEST, once, and it was never explained.
       Docs/autonomous-session-log.md records it beside one other unexplained COM block, and is
       careful to say it was SPINNING rather than blocked. Nothing here can time it out - a COM
       call on this thread owns the thread. If it does not return: DO NOT taskkill OUTLOOK.EXE
       (rule 7), wait, and if it never returns, revert the checkpoint. The script prints a line
       immediately before the call so that a transcript shows exactly where it stopped.

    IT IS A FOLDER RENAME, NOT ITEM MUTATION. No item is created, deleted, moved or modified.
    Rename-OutlookStore.ps1 makes the same argument in the same words and is the precedent: the
    narrowed rule forbids item mutation from a script, and renaming a store's root node on a guest
    whose identity the script verifies is not that.

    IDEMPOTENCY, AND IT IS STRONGER THAN THE OLD SCRIPT'S. Matching is on `Store.FilePath`,
    resolved and compared case-insensitively - not on display name, which is the thing under test
    and is user-editable besides. So: the same file already attached under the right name is a
    no-op; already attached under a WRONG name is a rename with no second attach; not attached is
    attach-then-rename. The old script matched on display name only and could therefore attach the
    same .pst twice under two names without noticing.

    WHAT IS TESTED WITHOUT A GUEST. -SelfTest exercises every decision - the display-name rules
    and the four-way add/rename/no-op/refuse verdict - against synthetic store tables. It makes no
    COM call, reads no registry and starts nothing, so it is safe on any machine including the
    maintainer's workstation. The COM half is guest-only and the self-test prints that list.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1) - COM does not work in session 0.
    Windows PowerShell 5.1 - no ternary, no `??`.

.PARAMETER ProfileName
    The profile the store goes into. It must be the DEFAULT profile; the script refuses otherwise
    rather than silently working on whichever profile Outlook actually opened. Named -ProfileName
    rather than -Profile because $Profile is a PowerShell automatic variable and a parameter of
    that name shadows it for the whole script.

.PARAMETER DisplayName
    What Outlook must show. Exact: the script fails if what comes back differs by so much as a
    character, compared case-sensitively.

.PARAMETER Path
    Where the .pst is, or goes. AddStoreEx creates the file; the DIRECTORY must already exist.

.PARAMETER ListOnly
    Report the current profile's stores and account count over COM, and change nothing.

.PARAMETER SelfTest
    Run the decision tests and exit. Touches nothing: no COM, no registry, no processes.

.PARAMETER Execute
    Without it, nothing is attached and nothing is renamed.

.EXAMPLE
    .\Add-OutlookPstStore.ps1 -SelfTest
    .\Add-OutlookPstStore.ps1 -ListOnly
    .\Add-OutlookPstStore.ps1 -ProfileName CorpusProfile -DisplayName 'test@vm.invalid' -Path C:\OutlookAI-Q5\pst\hub.pst -Execute
#>
[CmdletBinding(DefaultParameterSetName = 'Add')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Add')] [string] $ProfileName,
    [Parameter(Mandatory = $true, ParameterSetName = 'Add')] [string] $DisplayName,
    [Parameter(Mandatory = $true, ParameterSetName = 'Add')] [string] $Path,
    [Parameter(Mandatory = $true, ParameterSetName = 'List')] [switch] $ListOnly,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

# olStoreUnicode. MS-documented, and the only sensible choice here: the corpus store is ~400 MB
# and Build-Corpus writes 20,000 items, which is far past where the ANSI format's 2 GB ceiling is
# a good idea at all.
$script:OlStoreUnicode = 2

# How long to let the rename settle before re-reading the store. Rename-OutlookStore.ps1 uses the
# same two seconds, and that is the run this route's evidence comes from.
$script:RenameSettleSeconds = 2

# =============================================================================================
# PURE DECISIONS. No COM, no registry, no files, no output. Everything this script decides is
# decided here, so -SelfTest can decide it too without a guest.
# =============================================================================================

<#
    Whether a string is usable as a store display name. Returns the refusal text, or $null.

    A STORE'S DISPLAY NAME IS A FOLDER NAME - that is the whole basis of this script's route - so
    the rules are folder-name rules, and each refusal below is a failure that would otherwise be
    SILENT rather than loud: Outlook would accept the rename and hand back something else, and the
    tests would then miss the store on a machine that looked correctly built.
#>
function Test-StoreDisplayName {
    param([string] $DisplayName)

    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        return 'REFUSING: the display name is empty. It is what the tests look the store up by; there is no useful default for it.'
    }

    if ($DisplayName -ne $DisplayName.Trim()) {
        # Outlook trims a folder name, so this would come back different from what was asked for
        # and the verify below would fail AFTER the store had been attached. Refuse first.
        return "REFUSING: the display name '$DisplayName' has leading or trailing whitespace. Outlook trims a folder name, so the store would come back named '$($DisplayName.Trim())' and every exact-match lookup would miss it. Pass the trimmed name if that is what you meant."
    }

    if ($DisplayName -match '[\\/]') {
        return "REFUSING: '$DisplayName' contains a slash. '/' is the one character Outlook is consistently reported to reject in a folder name, and a store's display name IS its root folder's name."
    }

    return $null
}

<#
    The whole add-or-rename decision, made against the stores Outlook actually reports.

    $ExistingStores is whatever the caller read: objects with DisplayName and FilePath. FilePath
    may be $null - Outlook does not give one for every store kind - and a null one simply never
    matches, which is correct: this script only ever touches a store it can identify by file.

    Decision is one of:
      InvalidName    the display name cannot be used             <- refuses before anything runs
      AlreadyCorrect the file is attached under that exact name  <- no-op
      RenameOnly     the file is attached under another name     <- rename, no second attach
      NameTaken      a DIFFERENT store already has that name     <- refuses; see why below
      AddThenRename  the file is not attached                    <- attach, then rename

    NameTaken IS THE INTERESTING REFUSAL. Two stores sharing a display name is a profile the live
    suite cannot resolve - `expectedStoreDisplayNames` censuses BY NAME, and
    `bystanderStoreDisplayNames` refuses writes BY NAME - so a duplicate does not fail loudly, it
    makes one of the two stores invisible to every lookup. Refusing here is much cheaper than
    finding it during a tier run.
#>
function Resolve-StoreAddition {
    param(
        [Parameter(Mandatory = $true)] [string] $RequestedPath,
        [string] $RequestedDisplayName,
        [object[]] $ExistingStores
    )

    $nameProblem = Test-StoreDisplayName -DisplayName $RequestedDisplayName
    if ($null -ne $nameProblem) {
        return [pscustomobject]@{ Decision = 'InvalidName'; Message = $nameProblem }
    }

    $stores = @()
    if ($null -ne $ExistingStores) { $stores = @($ExistingStores) }

    $byPath = $null
    foreach ($store in $stores) {
        if ($null -ne $store.FilePath -and $store.FilePath -ieq $RequestedPath) { $byPath = $store }
    }

    # A name collision is only a collision with a store that is not the one being worked on.
    $collision = $null
    foreach ($store in $stores) {
        if ($store.DisplayName -ceq $RequestedDisplayName) {
            $isTarget = ($null -ne $byPath -and $null -ne $store.FilePath -and $store.FilePath -ieq $RequestedPath)
            if (-not $isTarget) { $collision = $store }
        }
    }

    if ($null -ne $collision) {
        $where = '<no file path>'
        if ($null -ne $collision.FilePath) { $where = $collision.FilePath }
        return [pscustomobject]@{
            Decision = 'NameTaken'
            Message  = "REFUSING: a different store is already called '$RequestedDisplayName' - it is $where. Two stores sharing a display name is a profile the live suite cannot resolve: expectedStoreDisplayNames censuses by name and bystanderStoreDisplayNames refuses writes by name, so one of the two would simply be invisible to every lookup. Rename or detach that one first."
        }
    }

    if ($null -eq $byPath) {
        return [pscustomobject]@{
            Decision = 'AddThenRename'
            Message  = "$RequestedPath is not in this profile. AddStoreEx attaches it (creating the file if it is not there), then its root folder is renamed to '$RequestedDisplayName'."
        }
    }

    if ($byPath.DisplayName -ceq $RequestedDisplayName) {
        return [pscustomobject]@{
            Decision = 'AlreadyCorrect'
            Message  = "$RequestedPath is already in this profile as '$RequestedDisplayName', byte for byte - nothing to do."
        }
    }

    return [pscustomobject]@{
        Decision = 'RenameOnly'
        Message  = "$RequestedPath is already in this profile, as '$($byPath.DisplayName)'. Only the rename to '$RequestedDisplayName' is needed - the store is NOT attached a second time."
    }
}

<#
    The store listing, as printed. Pure, so the markers and the warnings under them are exercised
    without an Outlook.
#>
function Format-StoreListing {
    param([object[]] $Stores, [string] $ProfileName, $AccountCount)

    $list = @()
    if ($null -ne $Stores) { $list = @($Stores) }

    $lines = @("profile '$ProfileName' holds $($list.Count) store(s):")
    if ($list.Count -eq 0) { $lines += '  <none>' }

    foreach ($store in $list) {
        $where = '<no file path - not a PST>'
        if ($null -ne $store.FilePath) { $where = $store.FilePath }
        $lines += ("  '{0}'  <-  {1}" -f $store.DisplayName, $where)
    }

    $seen = @()
    $duplicates = @()
    foreach ($store in $list) {
        if ($seen -contains $store.DisplayName) {
            if ($duplicates -notcontains $store.DisplayName) { $duplicates += $store.DisplayName }
        }
        $seen += $store.DisplayName
    }
    if ($duplicates.Count -gt 0) {
        $lines += ''
        $lines += "WARNING: $($duplicates.Count) display name(s) appear more than once: $($duplicates -join ', ')."
        $lines += '         The live suite looks stores up BY NAME, so one of each pair is invisible to it.'
    }

    if ($null -ne $AccountCount) {
        $lines += ''
        $lines += "Accounts.Count : $AccountCount"
        if ([int] $AccountCount -ne 0) {
            $lines += '         NOT the account-less shape. corpus-build refuses any profile with an account,'
            $lines += '         with no override - so this profile cannot be the corpus one.'
        }
    }

    return , $lines
}

# =============================================================================================
# SELF-TEST. Pure: no COM, no registry, no files, no guest.
# =============================================================================================

function ConvertTo-SelfTestText {
    param($Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Array]) {
        $parts = @()
        foreach ($item in $Value) { $parts += (ConvertTo-SelfTestText $item) }
        return '[' + ($parts -join ', ') + ']'
    }
    return [string] $Value
}

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    function Test-Case {
        param([string] $What, $Expected, $Actual)

        $script:SelfTestChecks++
        $expectedText = ConvertTo-SelfTestText $Expected
        $actualText = ConvertTo-SelfTestText $Actual
        if ($expectedText -ceq $actualText) {
            Write-Host ("  OK   {0}" -f $What)
        }
        else {
            $script:SelfTestFailures += "$What : expected $expectedText, got $actualText"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $expectedText, $actualText)
        }
    }

    function New-StoreRow {
        param([string] $DisplayName, $FilePath)
        return [pscustomobject]@{ DisplayName = $DisplayName; FilePath = $FilePath }
    }

    Write-Host 'Add-OutlookPstStore self-test. No COM call is made, nothing is read and nothing is written.'
    Write-Host ''
    Write-Host '== what may be a store display name =='

    Test-Case 'an ordinary name is fine' '<null>' (Test-StoreDisplayName -DisplayName 'Corpus A')
    Test-Case 'an @ is fine - measured twice on this build' '<null>' (Test-StoreDisplayName -DisplayName 'test@vm.invalid')
    Test-Case 'a dotted address is fine' '<null>' (Test-StoreDisplayName -DisplayName 'identity@vm.invalid')
    Test-Case 'an empty name refuses' $true ((Test-StoreDisplayName -DisplayName '') -like '*is empty*')
    Test-Case 'a whitespace-only name refuses' $true ((Test-StoreDisplayName -DisplayName '   ') -like '*is empty*')
    Test-Case 'a null name refuses rather than throwing' $true ((Test-StoreDisplayName -DisplayName $null) -like '*is empty*')
    Test-Case 'a trailing space refuses' $true ((Test-StoreDisplayName -DisplayName 'Corpus A ') -like '*trailing whitespace*')
    Test-Case 'and says what it would have become' $true ((Test-StoreDisplayName -DisplayName 'Corpus A ') -like "*named 'Corpus A'*")
    Test-Case 'a leading space refuses too' $true ((Test-StoreDisplayName -DisplayName ' Corpus A') -like '*trailing whitespace*')
    Test-Case 'a forward slash refuses' $true ((Test-StoreDisplayName -DisplayName 'Corpus/A') -like '*contains a slash*')
    Test-Case 'a backslash refuses' $true ((Test-StoreDisplayName -DisplayName 'Corpus\A') -like '*contains a slash*')

    Write-Host ''
    Write-Host '== the add-or-rename decision =='

    $profileStores = @(
        (New-StoreRow 'Corpus A' 'C:\pst\corpus-a.pst'),
        (New-StoreRow 'Outlook Data File' 'C:\pst\hub.pst'),
        (New-StoreRow 'Public Folders' $null))

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\new.pst' -RequestedDisplayName 'Bystander' -ExistingStores $profileStores
    Test-Case 'a new file is attached and then renamed' 'AddThenRename' $decision.Decision
    Test-Case 'and the message says both halves happen' $true ($decision.Message -like '*AddStoreEx attaches it*')

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\corpus-a.pst' -RequestedDisplayName 'Corpus A' -ExistingStores $profileStores
    Test-Case 'the same file under the same name is a no-op' 'AlreadyCorrect' $decision.Decision

    $decision = Resolve-StoreAddition -RequestedPath 'C:\PST\CORPUS-A.PST' -RequestedDisplayName 'Corpus A' -ExistingStores $profileStores
    Test-Case 'the path match is case-insensitive, as Outlook normalises it' 'AlreadyCorrect' $decision.Decision

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\hub.pst' -RequestedDisplayName 'test@vm.invalid' -ExistingStores $profileStores
    Test-Case 'an attached store under the wrong name is renamed only' 'RenameOnly' $decision.Decision
    Test-Case 'and it says it is not attached twice' $true ($decision.Message -like '*NOT attached a second time*')
    Test-Case 'and names what it is called now' $true ($decision.Message -like "*as 'Outlook Data File'*")

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\corpus-a.pst' -RequestedDisplayName 'corpus a' -ExistingStores $profileStores
    Test-Case 'a name differing only in case is a rename, not a match' 'RenameOnly' $decision.Decision

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\new.pst' -RequestedDisplayName 'Corpus A' -ExistingStores $profileStores
    Test-Case 'a name another store already has REFUSES' 'NameTaken' $decision.Decision
    Test-Case 'and says which store holds it' $true ($decision.Message -like '*C:\pst\corpus-a.pst*')
    Test-Case 'and why a duplicate is not merely untidy' $true ($decision.Message -like '*invisible to every lookup*')

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\hub.pst' -RequestedDisplayName 'Corpus A' -ExistingStores $profileStores
    Test-Case 'renaming ONTO another store''s name also refuses' 'NameTaken' $decision.Decision

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\new.pst' -RequestedDisplayName 'Public Folders' -ExistingStores $profileStores
    Test-Case 'a name held by a store with no file path still refuses' 'NameTaken' $decision.Decision
    Test-Case 'and says it could not name the file' $true ($decision.Message -like '*<no file path>*')

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\new.pst' -RequestedDisplayName ' ' -ExistingStores $profileStores
    Test-Case 'a bad name refuses before anything else is considered' 'InvalidName' $decision.Decision

    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\new.pst' -RequestedDisplayName 'Bystander' -ExistingStores @()
    Test-Case 'an empty profile just attaches' 'AddThenRename' $decision.Decision
    $decision = Resolve-StoreAddition -RequestedPath 'C:\pst\new.pst' -RequestedDisplayName 'Bystander' -ExistingStores $null
    Test-Case 'a null store table is the same, and does not throw' 'AddThenRename' $decision.Decision

    Write-Host ''
    Write-Host '== the listing, as printed =='

    $listing = Format-StoreListing -Stores $profileStores -ProfileName 'CorpusProfile' -AccountCount 0
    Test-Case 'the header counts them and names the profile' "profile 'CorpusProfile' holds 3 store(s):" $listing[0]
    Test-Case 'a store is printed with its file' "  'Corpus A'  <-  C:\pst\corpus-a.pst" $listing[1]
    Test-Case 'a store with no file says so rather than printing blank' "  'Public Folders'  <-  <no file path - not a PST>" $listing[3]
    Test-Case 'zero accounts is reported' $true (($listing -join "`n") -like '*Accounts.Count : 0*')
    Test-Case 'and raises no corpus warning' $false (($listing -join "`n") -like '*corpus-build refuses*')

    $listing = Format-StoreListing -Stores $profileStores -ProfileName 'TierProfile' -AccountCount 1
    Test-Case 'an account on the profile is called out' $true (($listing -join "`n") -like '*corpus-build refuses*')

    $listing = Format-StoreListing -Stores $profileStores -ProfileName 'CorpusProfile' -AccountCount $null
    Test-Case 'an unread account count prints nothing about accounts' $false (($listing -join "`n") -like '*Accounts.Count*')

    $duplicated = @((New-StoreRow 'Corpus A' 'C:\pst\a.pst'), (New-StoreRow 'Corpus A' 'C:\pst\b.pst'))
    $listing = Format-StoreListing -Stores $duplicated -ProfileName 'CorpusProfile' -AccountCount 0
    Test-Case 'two stores with one name WARNS' $true (($listing -join "`n") -like '*appear more than once: Corpus A*')
    Test-Case 'and says what it costs' $true (($listing -join "`n") -like '*invisible to it*')

    $listing = Format-StoreListing -Stores @() -ProfileName 'Empty' -AccountCount 0
    Test-Case 'no stores prints <none>' '  <none>' $listing[1]
    $listing = Format-StoreListing -Stores $null -ProfileName 'Empty' -AccountCount $null
    Test-Case 'a null store list does not throw' '  <none>' $listing[1]

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. These need a guest with Outlook, and nothing here can stand in for them:'
    Write-Host '  * binding Outlook at all, and which profile GetDefaultFolder lands on'
    Write-Host '  * NameSpace.AddStoreEx - including whether it spins, which it has done once'
    Write-Host '  * whether Store.DisplayName follows the root-folder rename (measured 2026-09-15,'
    Write-Host '    by Rename-OutlookStore.ps1, and NOT re-measured by this self-test)'
    Write-Host '  * the FilePath Outlook reports back, and how it normalises the one it was given'
    Write-Host '  * the guard (Assert-TestbedGuest) and the PST directory check (Resolve-PstPath)'

    if ($script:SelfTestFailures.Count -gt 0) {
        Write-Host ''
        foreach ($failure in $script:SelfTestFailures) { Write-Host "  $failure" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# EVERYTHING BELOW TOUCHES THE MACHINE. Guest only, session 1 only.
# =============================================================================================

# Dot-sourced for Assert-TestbedGuest, Resolve-PstPath and Invoke-WithOutlookSession. There is no
# MAPI left in that file to call - see its banner. Assert-OutlookNotRunning is deliberately NOT
# used here: this script drives the object model and therefore REQUIRES a live Outlook.
. "$PSScriptRoot\OutlookMapiInterop.ps1"

Assert-TestbedGuest -ExpectedUser $ExpectedUser

<#
    Every store Outlook reports, as { DisplayName; FilePath }. FilePath is $null for a store that
    does not have one, and reading it is wrapped because not every store kind answers.
#>
function Get-StoreRow {
    param([Parameter(Mandatory = $true)] $NameSpaceObject)

    $rows = @()
    for ($index = 1; $index -le $NameSpaceObject.Stores.Count; $index++) {
        $store = $NameSpaceObject.Stores.Item($index)
        $filePath = $null
        try { $filePath = $store.FilePath } catch { $filePath = $null }
        $rows += [pscustomobject]@{ DisplayName = $store.DisplayName; FilePath = $filePath }
    }
    return , $rows
}

<#
    The live Store object whose FilePath matches, re-fetched from the collection every time.

    RE-FETCHING IS THE POINT, not an inefficiency. Rename-OutlookStore.ps1 makes the same argument
    and it is what its measurement rests on: reusing a Store reference after a rename reads a
    cached RCW, which can report the value that was just written without it having reached the
    store - and Store.DisplayName and the root folder's Name are different properties on different
    objects, which is the entire question this route turns on.
#>
function Get-StoreByPath {
    param([Parameter(Mandatory = $true)] $NameSpaceObject, [Parameter(Mandatory = $true)] [string] $FilePath)

    for ($index = 1; $index -le $NameSpaceObject.Stores.Count; $index++) {
        $store = $NameSpaceObject.Stores.Item($index)
        $candidate = $null
        try { $candidate = $store.FilePath } catch { $candidate = $null }
        if ($null -ne $candidate -and $candidate -ieq $FilePath) { return $store }
    }
    return $null
}

# ---------------------------------------------------------------------------------------------
# LIST. Read-only, and it still binds Outlook - there is no other way to ask.
# ---------------------------------------------------------------------------------------------
if ($ListOnly) {
    Write-Host 'Reading the DEFAULT profile over COM. This binds Outlook, starting one if none is up,'
    Write-Host 'and never quits it. Nothing is written.'
    Write-Host ''

    Invoke-WithOutlookSession -Body {
        param($ns)

        $accountCount = $null
        try { $accountCount = $ns.Accounts.Count } catch { $accountCount = $null }

        foreach ($line in (Format-StoreListing -Stores (Get-StoreRow -NameSpaceObject $ns) -ProfileName $ns.CurrentProfileName -AccountCount $accountCount)) {
            Write-Host $line
        }
    }

    Write-Host ''
    Write-Host 'Read-only. Nothing changed.'
    return
}

# ---------------------------------------------------------------------------------------------
# ADD.
# ---------------------------------------------------------------------------------------------
$nameProblem = Test-StoreDisplayName -DisplayName $DisplayName
if ($null -ne $nameProblem) { throw $nameProblem }

# Resolve-PstPath normalises the spelling ONCE and checks the DIRECTORY exists - AddStoreEx
# creates the .pst, not the folder around it. Every comparison below is against this exact string.
$resolved = Resolve-PstPath -Path $Path

Write-Host "profile      : $ProfileName   (must be the DEFAULT profile - see below)"
Write-Host "display name : $DisplayName"
Write-Host "pst          : $resolved"
Write-Host ''

$script:AppliedDecision = $null
$script:AppliedMessage = $null
$script:FinalName = $null

Invoke-WithOutlookSession -Body {
    param($ns)

    $sessionProfile = $ns.CurrentProfileName
    Write-Host "Outlook is on profile '$sessionProfile'."
    if ($sessionProfile -ne $ProfileName) {
        throw @"
REFUSING: Outlook is logged on to '$sessionProfile' and you asked for '$ProfileName'.

This script drives the object model, and the object model gives you the profile Outlook opened -
which is the DEFAULT one. There is no flag that points it elsewhere: NameSpace.Logon is documented
to be able to raise the profile picker even when a default is set, and a dialog on an unattended
guest is a hang rather than a prompt, so this project does not use it.

Make '$ProfileName' the default, close Outlook, and run this again:

    .\Set-DefaultOutlookProfile.ps1 -Name $ProfileName -Execute
"@
    }

    $before = Get-StoreRow -NameSpaceObject $ns
    foreach ($line in (Format-StoreListing -Stores $before -ProfileName $sessionProfile -AccountCount $null)) {
        Write-Host $line
    }
    Write-Host ''

    $decision = Resolve-StoreAddition -RequestedPath $resolved -RequestedDisplayName $DisplayName -ExistingStores $before
    Write-Host "decision : $($decision.Decision)"
    Write-Host "  $($decision.Message)"
    $script:AppliedDecision = $decision.Decision
    $script:AppliedMessage = $decision.Message

    # The refusals fire BEFORE the -Execute gate on purpose: a dry run that says "this name is
    # already taken by another store" has told you the useful thing, and holding that back until
    # the real run wastes the dry run entirely.
    if ($decision.Decision -eq 'InvalidName' -or $decision.Decision -eq 'NameTaken') {
        throw $decision.Message
    }

    if (-not $Execute) {
        Write-Host ''
        Write-Host 'Dry run. Nothing attached, nothing renamed. Re-run with -Execute.'
        return
    }

    if ($decision.Decision -eq 'AlreadyCorrect') {
        $script:FinalName = $DisplayName
        return
    }

    if ($decision.Decision -eq 'AddThenRename') {
        Write-Host ''
        Write-Host "  calling NameSpace.AddStoreEx('$resolved', $script:OlStoreUnicode) ..."
        Write-Host '  IF THIS LINE IS THE LAST THING IN THE TRANSCRIPT, AddStoreEx IS SPINNING. It has done'
        Write-Host '  that once on this project, unexplained (Docs/autonomous-session-log.md). Nothing here'
        Write-Host '  can time out a COM call on its own thread. Wait. DO NOT taskkill OUTLOOK.EXE -'
        Write-Host '  mailbox-safety rule 7 forbids it outright. If it never returns, revert the checkpoint.'
        $ns.AddStoreEx($resolved, $script:OlStoreUnicode)
        Write-Host '  AddStoreEx returned.'
    }

    $store = Get-StoreByPath -NameSpaceObject $ns -FilePath $resolved
    if ($null -eq $store) {
        $lines = @()
        foreach ($row in (Get-StoreRow -NameSpaceObject $ns)) { $lines += ("  '{0}'  <-  {1}" -f $row.DisplayName, $row.FilePath) }
        throw "AddStoreEx returned and no store in this profile has FilePath '$resolved'. What Outlook does report:`n$($lines -join "`n")"
    }

    $root = $store.GetRootFolder()
    Write-Host ("  before : Store.DisplayName='{0}'  root.Name='{1}'" -f $store.DisplayName, $root.Name)

    $root.Name = $DisplayName
    Start-Sleep -Seconds $script:RenameSettleSeconds

    # RE-FETCH rather than re-read $store: see Get-StoreByPath's comment. The whole route rests on
    # Store.DisplayName following a property set on a different object, so reading it back off a
    # cached wrapper would prove nothing.
    $after = Get-StoreByPath -NameSpaceObject $ns -FilePath $resolved
    if ($null -eq $after) { throw "The store vanished from the profile after the rename." }

    $script:FinalName = $after.DisplayName
    Write-Host ("  after  : Store.DisplayName='{0}'  root.Name='{1}'" -f $after.DisplayName, $after.GetRootFolder().Name)

    if ($after.DisplayName -cne $DisplayName) {
        throw @"
Store.DisplayName reads '$($after.DisplayName)', not '$DisplayName'.

The root folder was renamed and the store did not follow. That is the outcome four research passes
could not rule out and one guest run (Rename-OutlookStore.ps1, 2026-09-15, build 16.0.17932) then
did - so if you are seeing it, something about this store or this build is different from that one,
and it matters: every test that looks a store up by display name will miss this one, on a machine
that otherwise looks correctly built. Do not build anything further on this profile until it is
understood.
"@
    }
}

Write-Host ''
if (-not $Execute) {
    return
}

if ($script:AppliedDecision -eq 'AlreadyCorrect') {
    Write-Host $script:AppliedMessage
}
else {
    Write-Host "Done. Outlook reports this store as '$script:FinalName' at $resolved."
}

Write-Host ''
Write-Host 'Remember the two DECLARATIONS this store may need in the live-test settings file:'
Write-Host '  expectedStoreDisplayNames   - censuses it. Every store needs this.'
Write-Host '  bystanderStoreDisplayNames  - refuses every write to it. Corpus and bystander stores'
Write-Host '                                need it; the HUB and the IDENTITY store must NOT have it.'
Write-Host 'Naming a store in exactly one of the two REFUSES the tier - see Testbed/README.md section 3b.'
Write-Host ''
Write-Host 'Outlook was left running, deliberately: mailbox-safety rule 7 prefers it headless to any'
Write-Host 'kill, and Build-Corpus.ps1 refuses to run unless Outlook is already up and warm.'
