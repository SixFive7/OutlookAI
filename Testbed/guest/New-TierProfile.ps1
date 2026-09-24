#Requires -Version 5.1
<#
    ============================================================================================
    2026-09-24 (LATER): THE SCRIPTS ALONE NOW REBUILD THE TIER PROFILE - AND THE ORDER MATTERS
    ============================================================================================

    WHAT WAS MISSING. The route that works needs two things this directory did not provide: the
    forcepst template, which was not the default (tier-profile.prf was, and it leaves the account
    unbound), and ForcePSTPath, which NOTHING under Testbed/ set. Both guests had it because two
    hand-run scratch scripts wrote it in session 1 - .work\run-forcepst.ps1 (OutlookAI-Indexed)
    and .work\guest2-chain.ps1 (OutlookAI-Unindexed), both 2026-09-15 - as

        New-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -Force
        New-ItemProperty -Path <that key> -Name 'ForcePSTPath' -PropertyType ExpandString -Value 'C:\OutlookAI-Tier' -Force

    and then ran this script with -TemplatePath pointed at a copy of the forcepst file.

    THAT FIRST LINE WAS NOT HARMLESS. New-Item -Force on an EXISTING registry key deletes every
    value and subkey under it - measured 2026-09-24 on a scratch key in the guest: a value and a
    subkey before, nothing after. It ran straight after Set-OfficeFirstRunSuppressed.ps1 -Execute
    had written three values INSIDE the Outlook key, and CP-05 shows exactly those three missing
    (Options\General does not exist, so HideNewOutlookToggle and DoNewOutlookAutoMigration are
    gone, and Preferences\UseNewOutlook is absent) while every value that script writes OUTSIDE
    the Outlook key is present, the policy form of the new-Outlook block included. The mechanism
    and the damage are measured; that this line caused it is inferred - it is the only thing in
    either chain that could have. OutlookAI-Indexed ran the same line and was not checked.

    WHAT CHANGED. The default template is tier-profile-forcepst.prf. -Execute writes ForcePSTPath
    (REG_EXPAND_SZ = -WorkDir) with New-ItemProperty on the existing key - no New-Item on it - and
    reads it back, kind included, before it sets ImportPRF. A template that names a PST service or
    DefaultStore is refused before anything is written (Test-TierTemplateShape), and so is a POLICY
    ForcePSTPath naming another directory (Resolve-ForcePstPathPlan). -Verify reads ForcePSTPath
    back, reads the delivery store's file out of the account's own EntryID, and requires it to be
    the ONE PST the tier profile references, under ForcePSTPath (Resolve-TierPstLayout).

    PROVEN FROM CP-02, IN BOTH ORDERS, 2026-09-24, OAI-UNINDEXED, Office 16.0.17932.20996:

      TIER FIRST - the order both guests were really built in. First-run suppression, -Execute,
      then the machine's FIRST Outlook start: ImportPRF gone (absent at the first look, 90 s in;
      sampled every 5 s in the other order, it was gone by the first sample), OutlookAI-Tier
      created and made default, C:\OutlookAI-Tier\Outlook.pst minted 7 s in, the
      POP3 password dialog up - and the account BOUND in that same start (COM: DeliveryStore
      'Outlook Data File' <- C:\OutlookAI-Tier\Outlook.pst, Drafts resolving). Rename-OutlookStore,
      then -Verify: all checks passed. The corpus profile by /PIM afterwards needed nothing extra.

      CORPUS FIRST - the order Testbed/README.md section 1 used to give. The import then happens
      at Outlook's SECOND start, which is the one that shows Office's one-time 'Check out our new
      look' dialog (NUIDialog), and the account came up UNBOUND: no 'Delivery Store EntryID',
      COM DeliveryStore NULL, and -Verify FAILED on it, as it should. It bound at the NEXT start
      with nothing else changed; whether the dialog is the cause is inferred from one run of each
      order, not proven. That order also leaves a stray, empty, DEFAULT profile called 'Outlook',
      made by the /PIM start on a machine with no profile at all. So section 1 now puts the tier
      first.

    ============================================================================================
    RUN ON BOTH GUESTS, 2026-09-15/16, AND IT WORKS. NOT A DRAFT.
    ============================================================================================

    This banner replaces the "never been executed" one, as that banner asked. It built the tier
    profile - one Unicode PST plus a POP3 account - on `OutlookAI-Indexed` (after two PRF
    variants) and on `OutlookAI-Unindexed` (first attempt, from the committed scripts, untouched
    by hand). That is the step three research passes had concluded could not be done for free.

    THE BUG THE FIRST RUN FOUND, because it is the reason to trust the script now and the reason
    not to trust a verifier generally. `-Verify` reported the PRF route DEAD five times over -
    "no profile named", "no accounts", "127.0.0.1 appears in no account value" - while its own
    raw dump, printed directly underneath, contradicted every line. It was reading the LEGACY
    Windows Messaging Subsystem hive; profiles moved to the Office hive at Outlook 2013. A
    verifier reading the wrong place does not fail loudly: it reports a working route as dead,
    and here that would have meant abandoning the only free path to a POP3 account. Fixed, and
    the fix is why the script now prefers the modern hive and falls back to the legacy one.

    WHAT MAKES THE PRF WORK, since it is not obvious and was expensive to find: the file must
    name NO PST service and NO `DefaultStore`, and `ForcePSTPath` decides where Outlook mints
    the store. A store Outlook MINTS is a store Outlook BINDS, and binding is the one step a
    text file cannot perform - name the store in the file and `Account.DeliveryStore` comes back
    NULL, after which `NewDraft` fails.

    ============================================================================================
    2026-09-24: IMPORTPRF IS MEASURED, AND -VERIFY NOW REMOVES IT; AND THIS SCRIPT HAS A GUARD
    ============================================================================================

    WHETHER OUTLOOK CLEARS ImportPRF WAS THE OPEN QUESTION, and the tier profile this script built
    on OAI-UNINDEXED answered it before anyone asked: -Execute set ImportPRF at 2026-09-15
    19:54:37 (its log says so), Outlook imported at 19:54:43, and ImportPRF was ABSENT the next
    time anybody looked - with First-Run present, and the Setup key last written at 19:57:59,
    three minutes after that start. Nothing in this repository or the scratch that drove that guest
    ever removes the value, so Outlook did. A deliberate measurement on 2026-09-24 (a New-OutlookProfile
    .ps1 import, sampled every 5 s) saw it go within 5 s of the start that imported it. CLEARED.

    -Verify NOW REMOVES IT ANYWAY (the maintainer's call, Q66 option 2: clear it regardless of
    the answer, and do not make the protection depend on anyone remembering a flag). It removes
    only a value naming THIS profile's .prf, and only once First-Run is back - evidence Outlook has
    run since -Execute deleted it - because removing it any earlier cancels the import itself. On
    the measured path it finds nothing to remove and says so. Resolve-ImportPrfClearance decides;
    -SelfTest walks every branch; the reasoning is in New-OutlookProfile.ps1 beside its twin.

    THE GUARD. This was the one script in this directory that writes the Outlook Setup key and
    never asked which machine it was on - and -Execute makes the next Outlook start build, and
    DEFAULT to, a profile. It now dot-sources OutlookMapiInterop.ps1 and calls
    Assert-TestbedGuest before anything touches the machine. STAGE OutlookMapiInterop.ps1 BESIDE
    IT: a from-scratch chain that copies only this script and its template now fails loudly on
    the dot-source, which is the intended failure. Proven both ways the same day: as vmadmin on
    the guest it passes; the dry run on the maintainer's workstation stops at "REFUSING TO RUN.
    This session is logged on as ..." and creates nothing.

    -VERIFY, RUN AGAINST THE WORKING TIER PROFILE (CP-05), FAILED IT FOUR WAYS THAT WERE ALL WRONG,
    and all four are fixed. (1) "the account manager holds at least one account" passed on ANY
    entry, and that key also lists the profile's data file and address book - it now counts MAIL
    accounts (Test-IsMailAccountEntry): 1, 'OutlookAI tier sink'. (2) "the account has a delivery
    store" looked only for a value named 00180102, and this build names it 'Delivery Store
    EntryID' - the dump directly underneath showed it bound to C:\OutlookAI-Tier\Outlook.pst.
    (3) "a POP3 port property exists" FAILED on no port value, and the working account stores none
    for either default port (110, 25): it now warns for a default port and fails only for a
    non-default one that did not land. (4) "the named PST exists" demanded tier.pst, which the
    ForcePSTPath route - the one that works - never creates: it is now not applicable when the
    imported .prf names no PST service. (Later the same day (4) and the stray-PST scan beside it
    were replaced by the PST-layout check described at the top of this banner.)

.SYNOPSIS
    Creates the testbed TIER profile - one POP3 account on a loopback mail sink, whose delivery
    store Outlook mints itself under ForcePSTPath - by importing a .prf file, and then PROVES
    whether Outlook honoured it.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    NEVER run this on the maintainer's workstation: it writes the Outlook Setup key and
    ForcePSTPath, and causes the next Outlook start to build - and default to - a mail profile.

    WHY THIS EXISTS. The tier profile needs a POP3/SMTP account pointed at a loopback sink. The
    Outlook object model is read-only for accounts, Extended MAPI can no longer create POP3
    services, and the one off-the-shelf component that can is excluded by the repository's
    Dependencies rule. A .prf file plus the ImportPRF registry value is the remaining free,
    documented, non-interactive route. The full evidence - with a source for every key - is in
    Docs/research/pop3-account-routes.md section A.

    IT WAS AN EXPERIMENT, AND THE THREE QUESTIONS IT TURNED ON ARE ANSWERED:

      1. Does Outlook 16.x process a PRF's INTERNET ACCOUNT sections (3, 5 and 7) at all? YES,
         measured 2026-09-15: the imported profile carries a genuine POP3 account
         (CLSID_OlkPOP3Account) with the sink's host, user and address. -Verify still fails
         loudly if no mail account appears.
      2. Does `[General] DefaultStore=Service1` bind the POP3 account's delivery store? NO,
         measured 2026-09-15 on OutlookAI-Indexed with tier-profile.prf: it binds the PROFILE's
         default store and leaves PROP_ACCT_DELIVERY_STORE (PT_BINARY - no .prf can carry it)
         unset, so DeliveryStore is NULL. The route that binds is the opposite: name NO PST
         service and NO DefaultStore, set ForcePSTPath, and let Outlook MINT the store - a store
         Outlook mints is a store Outlook binds. That is tier-profile-forcepst.prf, the default,
         and this script refuses a template of the other shape (Test-TierTemplateShape).
      3. Is the import genuinely silent? YES for the profile - but ONLY while `FirstRun` and
         `First-Run` are absent from the Setup key, which this script deletes, and Office's own
         first-run dialogs are a separate matter: Set-OfficeFirstRunSuppressed.ps1 runs first.

    IT DOES NOT START OUTLOOK, AND THAT IS DELIBERATE. The import happens at the next Outlook
    start, and a script that starts Outlook then owns Outlook's lifetime - which this project
    handles carefully and never from ad-hoc code (never taskkill; release COM references BEFORE
    any quit). So the flow is four steps, and the second is yours:

        .\New-TierProfile.ps1                 # dry run: prints the plan and changes nothing
        .\New-TierProfile.ps1 -Execute        # writes ForcePSTPath, the .prf and the Setup values
        <start Outlook once, in session 1, and let it settle - it imports and mints the store>
        .\Rename-OutlookStore.ps1 -StoreFilePath C:\OutlookAI-Tier\Outlook.pst -DisplayName tier@vm.invalid -Execute
        .\New-TierProfile.ps1 -Verify         # reads the profile hive and asserts; writes its log, and
                                              # removes ImportPRF if it is still set once Outlook has run

    STARTING STATE IT EXPECTS. A guest checkpoint where Office is installed and the tier profile
    does not exist yet. -Execute asserts that state rather than assuming it: it refuses if
    OUTLOOK.EXE is running, if the Office Setup key is missing, or if a profile of the target name
    already exists without -Force. Re-running -Execute against the same checkpoint is safe and
    converges; re-running it against a machine that has already imported once is what -Force is
    for, and the .prf's `OverwriteProfile=Yes` + `BackupProfile=No` are what should stop that
    creating a "Backup Of <name>" profile. Should. Which spelling 16.x honours is untested - if a
    backup profile appears, try BackupProfile=False in the template.

    WHAT IT NEVER DOES: create a COM object, open a store, read or write a mail item, or touch
    any profile other than the one named. Its registry writes are named before they are made, and
    there are exactly these: -Execute sets ForcePSTPath (REG_EXPAND_SZ, under HKCU\...\Outlook)
    and ImportPRF, and deletes First-Run and FirstRun (under HKCU\...\Outlook\Setup, creating that
    key if it is missing); -Verify may remove ImportPRF again, and only in the one case
    Resolve-ImportPrfClearance calls Clear. ForcePSTPath is left in place afterwards - it is
    per-user, so any PST Outlook later mints by default (a /PIM profile's store, for one) lands in
    -WorkDir too. That is how OAI-UNINDEXED's corpus store came to be
    C:\OutlookAI-Tier\Outlook Data File - CorpusProfile.pst: its /PIM profile was made after the tier.

.PARAMETER ProfileName
    The MAPI profile to create. Also the name asserted to be the ONLY profile of that name
    afterwards.

.PARAMETER StoreDisplayName
    The name the hub store must end up with. The .prf names NO store - Outlook mints it as
    'Outlook Data File' - so this is only printed, in the Rename-OutlookStore.ps1 line the next
    steps give. The live suite treats the hub's display name as an SMTP address, so this and
    -EmailAddress are the same string.

.PARAMETER WorkDir
    Where the rendered .prf goes AND what ForcePSTPath is set to, so it is where Outlook mints the
    delivery store (as Outlook.pst, measured). Deliberately space-free: the ImportPRF registry
    value takes a raw unquoted path.

.PARAMETER TemplatePath
    The .prf template. Defaults to tier-profile-forcepst.prf beside this script - the one shape
    that produces a bound account. A template that names a PST service or DefaultStore is REFUSED
    before anything is written; tier-profile.prf beside it is that shape, retired, kept as evidence.

.PARAMETER SinkHost
    The loopback sink's address. Used for both POP3 and SMTP.

.PARAMETER Execute
    Actually write. Without it nothing is created or modified and the plan is printed instead.

.PARAMETER Verify
    Read the profile hive and assert the import worked. Touches no profile, no store and no mail
    item. It writes its own log at -LogPath and, with no -Execute needed, removes ImportPRF when it
    still names the tier .prf after Outlook has run (see the 2026-09-24 banner section). Safe to run
    repeatedly.

.PARAMETER SelfTest
    Run the decision tests and exit: ImportPRF clearance, the template shape, the ForcePSTPath
    plan, the delivery-store path and the PST layout. Touches nothing - no registry, no guard, no
    process; the only files it opens are the two committed templates beside it, read-only.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.PARAMETER Force
    Allow -Execute to proceed when a profile of the target name already exists.

.EXAMPLE
    .\New-TierProfile.ps1
    .\New-TierProfile.ps1 -Execute
    .\New-TierProfile.ps1 -Verify
#>
[CmdletBinding()]
param(
    [string] $ProfileName      = 'OutlookAI-Tier',
    [string] $StoreDisplayName = 'tier@vm.invalid',
    [string] $EmailAddress     = 'tier@vm.invalid',
    [string] $AccountName      = 'OutlookAI tier sink',
    [string] $DisplayName      = 'OutlookAI Tier',
    [string] $Pop3User         = 'tier',
    [string] $WorkDir          = 'C:\OutlookAI-Tier',
    [string] $SinkHost         = '127.0.0.1',
    [int]    $Pop3Port         = 110,
    [int]    $SmtpPort         = 25,
    [string] $OfficeVersion    = '16.0',
    [string] $TemplatePath,
    [string] $LogPath          = 'C:\OutlookAI-Tier\new-tier-profile.log',
    [string[]] $ExpectedUser   = @('vmadmin'),
    [switch] $Execute,
    [switch] $Verify,
    [switch] $Force,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# =============================================================================================
# PURE DECISIONS - the ImportPRF clearance -Verify performs. No registry, no files, no output, so
# -SelfTest can walk every branch on any machine. COPIED from New-OutlookProfile.ps1 rather than
# shared, like the hive rule in the sibling scripts: this file dot-sources nothing before its
# guard, and the two copies are asserted by the same cases.
# =============================================================================================

<#
    The ProfileName= a .prf carries, from its [General] section. $null when there is none.
#>
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

<#
    What -Verify does about ImportPRF. The reasoning is in New-OutlookProfile.ps1, beside the
    function of the same name; in one line: remove it only when it names THIS profile's .prf and
    Outlook has demonstrably started since -Execute (First-Run is back), because removing it any
    earlier cancels the import itself.

    Decision: NotSet | Clear | KeepPending | KeepFailed | KeepOther.
#>
function Resolve-ImportPrfClearance {
    param(
        [string] $ImportPrfValue,
        [string] $PrfProfileName,
        [Parameter(Mandatory = $true)] [string] $ProfileName,
        [bool] $ProfileExists,
        [bool] $FirstRunPresent
    )

    if ([string]::IsNullOrEmpty($ImportPrfValue)) {
        return [pscustomobject]@{ Decision = 'NotSet'; Message = 'ImportPRF is not set, so no Outlook start can re-import. (Outlook removes it itself once it has imported the file - measured on this build.)' }
    }
    if ([string]::IsNullOrEmpty($PrfProfileName) -or $PrfProfileName -ne $ProfileName) {
        $whose = 'a file whose profile name could not be read'
        if (-not [string]::IsNullOrEmpty($PrfProfileName)) { $whose = "the .prf for profile '$PrfProfileName'" }
        return [pscustomobject]@{ Decision = 'KeepOther'; Message = "ImportPRF is set to '$ImportPrfValue' - $whose, not '$ProfileName'. Left alone: another import may still be pending, and removing the value would cancel it without a word. Build-Corpus.ps1 refuses to build while ImportPRF is set. To remove it anyway: .\New-OutlookProfile.ps1 -ClearImportPrf -Execute" }
    }
    if (-not $FirstRunPresent) {
        return [pscustomobject]@{ Decision = 'KeepPending'; Message = "ImportPRF still names the tier .prf and neither First-Run nor FirstRun is back, so Outlook has NOT started since -Execute - the import has not happened yet. Removing the value now would cancel it, so it is left in place. Start Outlook once, let it settle, and run -Verify again." }
    }
    if (-not $ProfileExists) {
        return [pscustomobject]@{ Decision = 'KeepFailed'; Message = "ImportPRF still names the tier .prf, Outlook HAS started since -Execute (First-Run is back), and there is no profile named '$ProfileName': the import did not take. Left in place as evidence; while First-Run exists Outlook does not act on it." }
    }
    return [pscustomobject]@{ Decision = 'Clear'; Message = 'ImportPRF still names the tier .prf after the import has run. Removing it: left in place, it is a rebuild waiting for the next start that finds First-Run gone.' }
}

<#
    Whether one entry of the account-manager key is a MAIL account. The same rule, from the same
    measurement, as New-OutlookProfile.ps1 and Build-Corpus.ps1: only a {ED475414-...} wrapper
    around a data file or an address book is excluded.
#>
function Test-IsMailAccountEntry {
    param([string] $Clsid, [string] $ServiceName)

    if ($Clsid -eq '{ED475414-B0D6-11D2-8C3B-00104B2A6676}' -and @('MSUPST MS', 'MSPST MS', 'CONTAB', 'EMABLT', 'MSPST AB') -contains $ServiceName) { return $false }
    return $true
}

<#
    What is wrong with a tier template, as a list of reasons; empty means usable. Pure.

    THE RULE IS THE MEASUREMENT OF 2026-09-15. A .prf that names a PST service and points
    [General] DefaultStore at it imports perfectly and leaves the POP3 account's delivery store
    UNBOUND - DeliveryStore NULL, and NewDraft fails - because PROP_ACCT_DELIVERY_STORE is a binary
    EntryID and a text file cannot carry one. That is tier-profile.prf, and it was measured on
    OutlookAI-Indexed. The route that works names NO PST service and NO DefaultStore, and lets
    Outlook MINT the delivery store under ForcePSTPath: a store Outlook mints is a store it binds.

    Both halves are refused, not just the measured pair. A PST service WITHOUT DefaultStore has
    not been measured either way, and the working route has neither - so a template carrying one
    is a template nobody has shown to work, and the build refuses it rather than finding out at the
    next step. A template with no internet account at all is not a tier template.
#>
function Test-TierTemplateShape {
    param([string] $Text)

    $problems = @()
    if ([string]::IsNullOrEmpty($Text)) { return , @('the template is empty') }

    $section = ''
    $accounts = 0
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.Trim()
        if ($line.Length -eq 0 -or $line.StartsWith(';')) { continue }
        if ($line.StartsWith('[') -and $line.EndsWith(']')) {
            $section = $line.Substring(1, $line.Length - 2).Trim()
            continue
        }
        $eq = $line.IndexOf('=')
        if ($eq -lt 1) { continue }
        $key = $line.Substring(0, $eq).Trim()
        $value = $line.Substring($eq + 1).Trim()
        if ($section -eq 'General' -and $key -eq 'DefaultStore') {
            $problems += "[General] carries DefaultStore=$value. DefaultStore binds the PROFILE's default store, not the POP3 account's delivery store - measured 2026-09-15: DeliveryStore came back NULL."
        }
        if ($section -eq 'Service List' -and @('Personal Folders', 'Unicode Personal Folders') -contains $value) {
            $problems += "[Service List] names a PST service ($key=$value). The working route names none and lets Outlook mint the delivery store under ForcePSTPath; a named PST is a store Outlook does not bind to the account."
        }
        if ($section -eq 'Internet Account List' -and $key -like 'Account*' -and $value.Length -gt 0) { $accounts++ }
    }
    if ($accounts -eq 0) { $problems += 'the template has no [Internet Account List] entry, so it would create no POP3 account - it is not a tier template.' }
    return , $problems
}

<#
    What -Execute does about ForcePSTPath, the value that decides where Outlook mints the tier's
    delivery store. Pure: the caller reads the registry and hands the facts in.

    Decision: AlreadySet | Set | Replace | RefusePolicy | RefuseKind.
      RefusePolicy  a POLICY ForcePSTPath (HKCU or HKLM ...\Policies\...) names another directory.
                    Policy outranks the user value, so the store would be minted there and the
                    rename step would look in the wrong place.
      RefuseKind    the value exists as something other than REG_EXPAND_SZ or REG_SZ.
#>
function Resolve-ForcePstPathPlan {
    param(
        [string] $Current,
        [string] $CurrentKind,
        [Parameter(Mandatory = $true)] [string] $Wanted,
        [string[]] $PolicyValues
    )

    foreach ($policy in @($PolicyValues)) {
        if (-not [string]::IsNullOrEmpty($policy) -and $policy.TrimEnd('\') -ne $Wanted.TrimEnd('\')) {
            return [pscustomobject]@{ Decision = 'RefusePolicy'; Message = "A POLICY value ForcePSTPath = '$policy' is set, and policy outranks the user value this script writes: Outlook would mint the tier store there, not in '$Wanted', and every later step looks for it in the wrong place. Remove the policy or pass -WorkDir '$policy'." }
        }
    }
    if ([string]::IsNullOrEmpty($Current)) {
        return [pscustomobject]@{ Decision = 'Set'; Message = "ForcePSTPath is not set; it will be written as REG_EXPAND_SZ '$Wanted'." }
    }
    if (@('ExpandString', 'String') -notcontains $CurrentKind) {
        return [pscustomobject]@{ Decision = 'RefuseKind'; Message = "ForcePSTPath exists as $CurrentKind, not a string. Nothing here knows what wrote it; remove it by hand and re-run." }
    }
    if ($Current.TrimEnd('\') -eq $Wanted.TrimEnd('\') -and $CurrentKind -eq 'ExpandString') {
        return [pscustomobject]@{ Decision = 'AlreadySet'; Message = "ForcePSTPath is already REG_EXPAND_SZ '$Current'." }
    }
    return [pscustomobject]@{ Decision = 'Replace'; Message = "ForcePSTPath is $CurrentKind '$Current'; it will be rewritten as REG_EXPAND_SZ '$Wanted'. It is per-user, so every PST Outlook mints by default from now on lands there too." }
}

<#
    The .pst path inside a store EntryID, read out of the raw bytes of the account's 'Delivery
    Store EntryID' value. A PST store's EntryID carries its file path as text - which is how the
    2026-09-24 dump showed the working account bound to C:\OutlookAI-Tier\Outlook.pst - in the ANSI
    or the UTF-16 form depending on the provider, so both decodings are searched. $null when
    neither holds a path.
#>
function Get-PstPathFromEntryId {
    param([byte[]] $Bytes)

    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return $null }
    foreach ($text in @([System.Text.Encoding]::Unicode.GetString($Bytes), [System.Text.Encoding]::ASCII.GetString($Bytes))) {
        $m = [regex]::Match($text, '[A-Za-z]:\\[^\x00-\x1f"<>*?|]*?\.pst', 'IgnoreCase')
        if ($m.Success) { return $m.Value }
    }
    return $null
}

<#
    Whether the tier profile's PST layout is the one the route produces: ONE .pst in the profile,
    it is the account's delivery store, and it sits under ForcePSTPath. Pure; the caller reads the
    profile's PST paths (Get-ProfilePstPath) and the delivery store's path out of its EntryID.

    Decision: Pass | Fail | Unknown.

    WHY THE PROFILE AND NOT THE DIRECTORY. Until 2026-09-24 this was a scan of Documents\Outlook
    Files that failed on ANY .pst there - and a corpus profile built FIRST keeps its /PIM store
    there, legitimately - while never looking in ForcePSTPath, where the import actually mints. A
    directory holds other profiles' files: on OAI-UNINDEXED (CP-05, and the tier-first rehearsal)
    C:\OutlookAI-Tier also holds the corpus profile's store, minted there later because
    ForcePSTPath outlives the import. What the tier profile itself references is the only list
    that is about the tier profile.
#>
function Resolve-TierPstLayout {
    param([string[]] $ProfilePstPaths, [string] $DeliveryStorePath, [Parameter(Mandatory = $true)] [string] $WorkDir)

    $paths = @($ProfilePstPaths | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object -Unique)
    if ([string]::IsNullOrEmpty($DeliveryStorePath)) {
        return [pscustomobject]@{ Decision = 'Unknown'; Message = "the delivery store's EntryID carries no readable .pst path, so which file it is cannot be shown here. The profile references: $(if ($paths.Count) { $paths -join ', ' } else { '<none read>' })." }
    }
    if ((Split-Path -Parent $DeliveryStorePath).TrimEnd('\') -ine $WorkDir.TrimEnd('\')) {
        return [pscustomobject]@{ Decision = 'Fail'; Message = "the account delivers to $DeliveryStorePath, which is not under ForcePSTPath ($WorkDir): ForcePSTPath did not decide where the store was minted." }
    }
    if ($paths.Count -eq 0) {
        return [pscustomobject]@{ Decision = 'Unknown'; Message = "the delivery store is $DeliveryStorePath, under ForcePSTPath; no PST path value was read out of the profile itself, so whether it holds a second PST is not shown here." }
    }
    $others = @($paths | Where-Object { $_ -ine $DeliveryStorePath })
    if ($others.Count -gt 0) {
        return [pscustomobject]@{ Decision = 'Fail'; Message = "the profile references another PST besides the delivery store ${DeliveryStorePath}: $($others -join ', '). The route produces exactly one - a second means the import made a store the account is not bound to." }
    }
    if (@($paths | Where-Object { $_ -ieq $DeliveryStorePath }).Count -eq 0) {
        return [pscustomobject]@{ Decision = 'Fail'; Message = "the account delivers to $DeliveryStorePath and the profile references no such PST." }
    }
    return [pscustomobject]@{ Decision = 'Pass'; Message = "one PST in the profile, $DeliveryStorePath, and it is the account's delivery store" }
}

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:SelfTestChecks++
        $e = "$Expected"; if ($null -eq $Expected) { $e = '<null>' }
        $a = "$Actual"; if ($null -eq $Actual) { $a = '<null>' }
        if ($e -ceq $a) { Write-Host ("  OK   {0}" -f $What) }
        else {
            $script:SelfTestFailures += "$What : expected $e, got $a"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $e, $a)
        }
    }

    Write-Host 'New-TierProfile self-test. Nothing is written and nothing is started; the only thing read is the two committed templates beside this script.'
    Write-Host ''
    Write-Host '== which profile a .prf names =='
    Test-Case 'the [General] ProfileName is read' 'OutlookAI-Tier' (Get-PrfProfileName -Text "; c`r`n[General]`r`nCustom=1`r`nProfileName=OutlookAI-Tier`r`nDefaultProfile=Yes`r`n")
    Test-Case 'a template token is returned as written, never guessed at' '{{PROFILE_NAME}}' (Get-PrfProfileName -Text "[General]`r`nProfileName={{PROFILE_NAME}}`r`n")
    Test-Case 'outside [General] it does not count' '<null>' (Get-PrfProfileName -Text "[Account1]`r`nProfileName=x`r`n")
    Test-Case 'nothing in, nothing out, no throw' '<null>' (Get-PrfProfileName -Text $null)

    Write-Host ''
    Write-Host '== what -Verify does about ImportPRF =='
    $prf = 'C:\OutlookAI-Tier\tier-profile.prf'
    Test-Case 'absent: NotSet' 'NotSet' (Resolve-ImportPrfClearance -ImportPrfValue '' -PrfProfileName $null -ProfileName 'OutlookAI-Tier' -ProfileExists $true -FirstRunPresent $true).Decision
    Test-Case 'ours, imported, Outlook ran since: Clear' 'Clear' (Resolve-ImportPrfClearance -ImportPrfValue $prf -PrfProfileName 'OutlookAI-Tier' -ProfileName 'OutlookAI-Tier' -ProfileExists $true -FirstRunPresent $true).Decision
    Test-Case 'ours, Outlook NOT started since -Execute: KeepPending, even with the profile there (-Force rebuild)' 'KeepPending' (Resolve-ImportPrfClearance -ImportPrfValue $prf -PrfProfileName 'OutlookAI-Tier' -ProfileName 'OutlookAI-Tier' -ProfileExists $true -FirstRunPresent $false).Decision
    Test-Case 'ours, Outlook ran, no profile: KeepFailed' 'KeepFailed' (Resolve-ImportPrfClearance -ImportPrfValue $prf -PrfProfileName 'OutlookAI-Tier' -ProfileName 'OutlookAI-Tier' -ProfileExists $false -FirstRunPresent $true).Decision
    Test-Case "the corpus profile's .prf is never the tier script's to cancel" 'KeepOther' (Resolve-ImportPrfClearance -ImportPrfValue 'C:\OutlookAI-Profiles\CorpusProfile.prf' -PrfProfileName 'CorpusProfile' -ProfileName 'OutlookAI-Tier' -ProfileExists $true -FirstRunPresent $true).Decision
    Test-Case 'nor is a file whose profile cannot be read' 'KeepOther' (Resolve-ImportPrfClearance -ImportPrfValue 'C:\gone.prf' -PrfProfileName $null -ProfileName 'OutlookAI-Tier' -ProfileExists $true -FirstRunPresent $true).Decision

    Write-Host ''
    Write-Host '== which account-manager entries are MAIL accounts (the tier profile, as measured) =='
    Test-Case "its POP3 account 'OutlookAI tier sink' ({ED475411-...}) is one" $true (Test-IsMailAccountEntry -Clsid '{ED475411-B0D6-11D2-8C3B-00104B2A6676}' -ServiceName '')
    Test-Case "its 'Outlook Address Book' wrapper (CONTAB) is not" $false (Test-IsMailAccountEntry -Clsid '{ED475414-B0D6-11D2-8C3B-00104B2A6676}' -ServiceName 'CONTAB')
    Test-Case "its 'Outlook Data File' wrapper (MSUPST MS) is not" $false (Test-IsMailAccountEntry -Clsid '{ED475414-B0D6-11D2-8C3B-00104B2A6676}' -ServiceName 'MSUPST MS')

    Write-Host ''
    Write-Host '== which templates the build accepts (the measured route: no PST service, no DefaultStore) =='
    $good = "[General]`r`nProfileName=x`r`n[Service List]`r`nService1=Outlook Address Book`r`n[Internet Account List]`r`nAccount1=I_Mail`r`n"
    Test-Case 'the forcepst shape is accepted' 0 (Test-TierTemplateShape -Text $good).Count
    $withDefaultStore = $good.Replace("ProfileName=x`r`n", "ProfileName=x`r`nDefaultStore=Service1`r`n")
    Test-Case 'DefaultStore is refused (measured: DeliveryStore NULL)' 1 (Test-TierTemplateShape -Text $withDefaultStore).Count
    $withPst = $good.Replace('Service1=Outlook Address Book', "Service1=Unicode Personal Folders`r`nService2=Outlook Address Book")
    Test-Case 'a named PST service is refused' 1 (Test-TierTemplateShape -Text $withPst).Count
    Test-Case 'the ANSI PST service is refused too' 1 (Test-TierTemplateShape -Text ($good.Replace('Service1=Outlook Address Book', 'Service1=Personal Folders'))).Count
    Test-Case 'a commented-out DefaultStore does not count' 0 (Test-TierTemplateShape -Text ($good.Replace("ProfileName=x`r`n", "ProfileName=x`r`n;DefaultStore=Service1`r`n"))).Count
    Test-Case 'no internet account is not a tier template' 1 (Test-TierTemplateShape -Text ($good.Replace('Account1=I_Mail', ''))).Count
    Test-Case 'an empty template is refused, and does not throw' 1 (Test-TierTemplateShape -Text '').Count
    # The two COMMITTED templates, read from beside this script - the only files -SelfTest opens,
    # and only to read. This is what pins the default to the working route.
    foreach ($pair in @(@('tier-profile-forcepst.prf', 0), @('tier-profile.prf', 2))) {
        $path = Join-Path $PSScriptRoot $pair[0]
        if (Test-Path -LiteralPath $path) {
            Test-Case ("committed {0}: {1} problem(s)" -f $pair[0], $pair[1]) $pair[1] (Test-TierTemplateShape -Text ([System.IO.File]::ReadAllText($path))).Count
        }
        else { Write-Host ("  --   committed {0} is not beside this script; not checked" -f $pair[0]) }
    }

    Write-Host ''
    Write-Host '== what -Execute does about ForcePSTPath =='
    Test-Case 'absent: Set' 'Set' (Resolve-ForcePstPathPlan -Current $null -CurrentKind $null -Wanted 'C:\OutlookAI-Tier').Decision
    Test-Case 'already the expandable string: AlreadySet' 'AlreadySet' (Resolve-ForcePstPathPlan -Current 'C:\OutlookAI-Tier' -CurrentKind 'ExpandString' -Wanted 'C:\OutlookAI-Tier').Decision
    Test-Case 'a trailing backslash is the same directory' 'AlreadySet' (Resolve-ForcePstPathPlan -Current 'C:\OutlookAI-Tier\' -CurrentKind 'ExpandString' -Wanted 'C:\OutlookAI-Tier').Decision
    Test-Case 'right directory, plain REG_SZ: Replace (the kind is part of the value)' 'Replace' (Resolve-ForcePstPathPlan -Current 'C:\OutlookAI-Tier' -CurrentKind 'String' -Wanted 'C:\OutlookAI-Tier').Decision
    Test-Case 'another directory: Replace' 'Replace' (Resolve-ForcePstPathPlan -Current 'D:\Elsewhere' -CurrentKind 'ExpandString' -Wanted 'C:\OutlookAI-Tier').Decision
    Test-Case 'a DWORD there: RefuseKind' 'RefuseKind' (Resolve-ForcePstPathPlan -Current '1' -CurrentKind 'DWord' -Wanted 'C:\OutlookAI-Tier').Decision
    Test-Case 'a policy naming another directory: RefusePolicy' 'RefusePolicy' (Resolve-ForcePstPathPlan -Current $null -CurrentKind $null -Wanted 'C:\OutlookAI-Tier' -PolicyValues @('D:\Policy')).Decision
    Test-Case 'a policy naming the same directory is no obstacle' 'Set' (Resolve-ForcePstPathPlan -Current $null -CurrentKind $null -Wanted 'C:\OutlookAI-Tier' -PolicyValues @('C:\OutlookAI-Tier\')).Decision
    Test-Case 'empty policy slots are ignored' 'Set' (Resolve-ForcePstPathPlan -Current '' -CurrentKind $null -Wanted 'C:\OutlookAI-Tier' -PolicyValues @($null, '')).Decision

    Write-Host ''
    Write-Host '== the delivery store path, out of its EntryID bytes =='
    $prefix = [byte[]](0, 0, 0, 0, 0x38, 0xA1, 0xBB, 0x10, 5, 0xE5, 0x10, 0x1A, 0xA1, 0xBB, 8, 0, 0x2B, 0x2A, 0x56, 0xC2, 0, 0, 0x6D, 0x73, 0x70, 0x73, 0x74, 0x2E, 0x64, 0x6C, 0x6C, 0)
    $unicodeId = $prefix + [System.Text.Encoding]::Unicode.GetBytes('C:\OutlookAI-Tier\Outlook.pst') + [byte[]](0, 0)
    Test-Case 'a UTF-16 path is found' 'C:\OutlookAI-Tier\Outlook.pst' (Get-PstPathFromEntryId -Bytes $unicodeId)
    $ansiId = $prefix + [System.Text.Encoding]::ASCII.GetBytes('C:\OutlookAI-Tier\Outlook Data File - CorpusProfile.pst') + [byte[]](0)
    Test-Case 'an ANSI path with spaces is found' 'C:\OutlookAI-Tier\Outlook Data File - CorpusProfile.pst' (Get-PstPathFromEntryId -Bytes $ansiId)
    Test-Case 'no path, no answer' '<null>' (Get-PstPathFromEntryId -Bytes $prefix)
    Test-Case 'nothing in, nothing out, no throw' '<null>' (Get-PstPathFromEntryId -Bytes $null)

    Write-Host ''
    Write-Host "== the tier profile's PST layout: one PST, the delivery store, under ForcePSTPath =="
    $wd = 'C:\OutlookAI-Tier'
    $minted = 'C:\OutlookAI-Tier\Outlook.pst'
    Test-Case 'one PST, and it is the delivery store: Pass' 'Pass' (Resolve-TierPstLayout -ProfilePstPaths @($minted) -DeliveryStorePath $minted -WorkDir $wd).Decision
    Test-Case 'the comparison is case-insensitive' 'Pass' (Resolve-TierPstLayout -ProfilePstPaths @('c:\outlookai-tier\OUTLOOK.PST') -DeliveryStorePath $minted -WorkDir 'C:\OutlookAI-Tier\').Decision
    Test-Case 'one PST named in two profile sections is still one PST' 'Pass' (Resolve-TierPstLayout -ProfilePstPaths @($minted, $minted) -DeliveryStorePath $minted -WorkDir $wd).Decision
    Test-Case 'a second PST in the TIER profile: Fail' 'Fail' (Resolve-TierPstLayout -ProfilePstPaths @($minted, 'C:\OutlookAI-Tier\tier.pst') -DeliveryStorePath $minted -WorkDir $wd).Decision
    Test-Case 'a delivery store outside ForcePSTPath: Fail' 'Fail' (Resolve-TierPstLayout -ProfilePstPaths @('C:\Users\vmadmin\Documents\Outlook Files\Outlook.pst') -DeliveryStorePath 'C:\Users\vmadmin\Documents\Outlook Files\Outlook.pst' -WorkDir $wd).Decision
    Test-Case 'a delivery store the profile does not reference: Fail' 'Fail' (Resolve-TierPstLayout -ProfilePstPaths @('C:\OutlookAI-Tier\x.pst') -DeliveryStorePath $minted -WorkDir $wd).Decision
    Test-Case 'no readable delivery path: Unknown, never Pass' 'Unknown' (Resolve-TierPstLayout -ProfilePstPaths @($minted) -DeliveryStorePath $null -WorkDir $wd).Decision
    Test-Case 'no PST path read out of the profile: Unknown, never Pass' 'Unknown' (Resolve-TierPstLayout -ProfilePstPaths @() -DeliveryStorePath $minted -WorkDir $wd).Decision

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Guest-only: the guard, every registry read and write (ImportPRF, First-Run,'
    Write-Host 'ForcePSTPath and its read-back), the .prf render and read-back, whether Outlook honours the file'
    Write-Host 'and mints the store under ForcePSTPath, and the ImportPRF removal -Verify makes.'
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

# THE GUARD, added 2026-09-24. This script writes the Outlook Setup key and makes the next Outlook
# start build (and, with DefaultProfile=Yes in the template, DEFAULT to) a profile - and until
# now it was the one script in this directory that writes the Outlook Setup key and never asked
# which machine it was on. Asserted before ANYTHING, including the dry run: on the maintainer's
# workstation even -Verify reads a real profile hive and would print its account values into a
# log.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

# ---------------------------------------------------------------------------------------------
# Paths and constants. Every registry path this script touches is named here and nowhere else,
# so a reader can see the whole blast radius in one place.
# ---------------------------------------------------------------------------------------------

$setupKey    = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Setup"
$outlookKey  = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook"
# PROFILES MOVED AT OUTLOOK 2013. This read the legacy Windows Messaging Subsystem path and
# therefore found nothing on a 16.x guest - reporting "no profile named X" and "the account
# manager holds no accounts" while the profile and a fully populated POP3 account sat in the
# Office hive a few keys away. Measured 2026-09-15: five checks failed and the script's own raw
# dump, printed directly underneath them, contradicted every one.
#
# This is the exact trap Docs/research/pop3-account-routes.md warns about - "older tooling that
# hardcodes the WMS path will silently look in the wrong place on 16.x" - and it is worth more
# than a one-line fix, because a verifier that reads the wrong hive does not fail loudly: it
# reports a working route as dead, which is the most expensive wrong answer available here.
#
# Both paths are checked, newest first, so this keeps working on an older Outlook.
$profilesKeyModern = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Profiles"
$profilesKeyLegacy = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
$profilesKey = $profilesKeyLegacy
if (Test-Path $profilesKeyModern) { $profilesKey = $profilesKeyModern }

# The account-manager subkey under a profile. Stable since Outlook 2002.
$acctMgrSubkey = '9375CFF0413111d3B88A00104B2A6676'

# PROP_ACCT_DELIVERY_STORE, PT_BINARY, tag 0x00180102 - the property a .prf cannot express.
$deliveryStoreValueName = '00180102'

# FORCEPSTPATH - where Outlook mints a PST nobody named, and therefore where the tier's delivery
# store lands. REG_EXPAND_SZ under the Outlook key (not under Setup). Written by -Execute since
# 2026-09-24; before that it was set by hand-run scratch scripts on both guests and nothing under
# Testbed/ set it, so the working route could not be rebuilt from this directory. The policy
# locations are READ ONLY, to refuse a build a policy would redirect.
$forcePstName = 'ForcePSTPath'
$forcePstPolicyKeys = @(
    "HKCU:\Software\Policies\Microsoft\Office\$OfficeVersion\Outlook",
    "HKLM:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Outlook"
)

$prfPath = Join-Path $WorkDir 'tier-profile.prf'

# THE DEFAULT IS THE ROUTE THAT WORKS, since 2026-09-24. It used to be tier-profile.prf, which
# names a PST service and DefaultStore and leaves the account's delivery store unbound (measured
# 2026-09-15); every working tier profile was built by passing the forcepst file explicitly.
if (-not $TemplatePath) {
    $TemplatePath = Join-Path $PSScriptRoot 'tier-profile-forcepst.prf'
}

<#
    ForcePSTPath as it stands: the user value's data (unexpanded) and kind, and every policy
    value. Reads only.
#>
function Get-ForcePstPathState {
    $current = $null
    $kind = $null
    $key = Get-Item -LiteralPath $outlookKey -ErrorAction SilentlyContinue
    if ($null -ne $key -and @($key.GetValueNames()) -contains $forcePstName) {
        $current = [string] $key.GetValue($forcePstName, $null, 'DoNotExpandEnvironmentNames')
        $kind = [string] $key.GetValueKind($forcePstName)
    }
    $policies = @()
    foreach ($p in $forcePstPolicyKeys) {
        $pk = Get-Item -LiteralPath $p -ErrorAction SilentlyContinue
        if ($null -ne $pk -and @($pk.GetValueNames()) -contains $forcePstName) {
            $policies += [string] $pk.GetValue($forcePstName, $null, 'DoNotExpandEnvironmentNames')
        }
    }
    return [pscustomobject]@{ Current = $current; Kind = $kind; Policies = $policies }
}

$script:Failures = @()
$script:Lines    = @()

function Say {
    param([string] $Text)
    $script:Lines += $Text
    Write-Host $Text
}

function Pass {
    param([string] $What, [string] $Detail)
    Say ("  OK   {0}{1}" -f $What, $(if ($Detail) { " - $Detail" } else { '' }))
}

function Fail {
    param([string] $What, [string] $Why)
    $script:Failures += "$What : $Why"
    Say ("  FAIL {0} - {1}" -f $What, $Why)
}

function Save-Log {
    # A dry run creates NOTHING, including its own log directory. A "dry run" that leaves a
    # directory behind is a dry run whose promise is already false.
    if (-not ($Execute -or $Verify)) { return }

    $dir = Split-Path -Parent $LogPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Set-Content -LiteralPath $LogPath -Value $script:Lines -Encoding UTF8
    Write-Host ''
    Write-Host "Log: $LogPath"
}

# ---------------------------------------------------------------------------------------------
# Reading the profile hive. Read-only, and it never opens a store: these are registry values,
# not MAPI calls. Account property values are REG_BINARY holding either 8-bit or UTF-16LE text,
# so a search decodes both and reports which matched.
# ---------------------------------------------------------------------------------------------

function Get-ProfileNames {
    if (-not (Test-Path -LiteralPath $profilesKey)) {
        return @()
    }
    return @(Get-ChildItem -LiteralPath $profilesKey -ErrorAction SilentlyContinue |
        ForEach-Object { $_.PSChildName })
}

function ConvertTo-ReadableText {
    param($Data)
    if ($null -eq $Data) { return '' }
    if ($Data -is [string]) { return $Data }

    # Only REG_BINARY gets the two-encoding treatment. A DWORD is a number and rendering it as
    # text would invent a string that is not in the registry - which is exactly the kind of
    # made-up evidence the verify step must not produce.
    if (-not ($Data -is [byte[]] -or $Data -is [System.Array])) {
        return [string] $Data
    }

    $bytes = New-Object 'System.Collections.Generic.List[byte]'
    foreach ($b in $Data) {
        $n = 0
        if (-not [int]::TryParse([string] $b, [ref] $n)) { return [string] $Data }
        if ($n -lt 0 -or $n -gt 255) { return [string] $Data }
        $bytes.Add([byte] $n)
    }
    if ($bytes.Count -eq 0) { return '' }

    $arr     = $bytes.ToArray()
    $ascii   = [System.Text.Encoding]::ASCII.GetString($arr)
    $unicode = [System.Text.Encoding]::Unicode.GetString($arr)
    return ($ascii + '|' + $unicode)
}

function Get-AccountValueDump {
    param([string] $Profile)

    $result = New-Object psobject -Property @{
        AccountKeys        = @()
        MailAccountNames   = @()
        Text               = ''
        HasDeliveryStore   = $false
        DeliveryStoreBytes = 0
        DeliveryStorePath  = $null
    }

    $mgrPath = Join-Path (Join-Path $profilesKey $Profile) $acctMgrSubkey
    if (-not (Test-Path -LiteralPath $mgrPath)) {
        return $result
    }

    # MAIL accounts among the direct entries - Outlook lists the profile's data files and address
    # book here too (Test-IsMailAccountEntry), so "any subkey at all" passed on a profile whose
    # account had never landed. Measured 2026-09-24: this profile holds its POP3 account
    # ({ED475411-...}) plus two such wrappers.
    foreach ($entry in @(Get-ChildItem -LiteralPath $mgrPath -ErrorAction SilentlyContinue)) {
        if (Test-IsMailAccountEntry -Clsid ([string] $entry.GetValue('clsid', '')) -ServiceName ([string] $entry.GetValue('Service Name', ''))) {
            $result.MailAccountNames += [string] $entry.GetValue('Account Name', $entry.PSChildName)
        }
    }

    $texts = @()
    $keys  = @()
    foreach ($child in (Get-ChildItem -LiteralPath $mgrPath -Recurse -ErrorAction SilentlyContinue)) {
        $keys += $child.Name
        $props = $null
        try { $props = Get-ItemProperty -LiteralPath $child.PSPath -ErrorAction Stop } catch { $props = $null }
        if ($null -eq $props) { continue }
        foreach ($p in $props.PSObject.Properties) {
            if ($p.Name -like 'PS*') { continue }
            $texts += ("{0}={1}" -f $p.Name, (ConvertTo-ReadableText $p.Value))
            # TWO SPELLINGS OF ONE PROPERTY. The account manager on this build names its values by
            # FRIENDLY NAME - 'Delivery Store EntryID', 'POP3 Server', 'POP3 User' - not by the hex
            # property tag this was written against, so the 00180102 test alone FAILED the working
            # tier profile on 2026-09-24 while its own dump, below it, showed the delivery store
            # bound to C:\OutlookAI-Tier\Outlook.pst. Either spelling counts.
            if ($p.Name -eq $deliveryStoreValueName -or $p.Name -eq 'Delivery Store EntryID') {
                $result.HasDeliveryStore = $true
                if ($p.Value -is [array]) {
                    $result.DeliveryStoreBytes = $p.Value.Count
                    $result.DeliveryStorePath = Get-PstPathFromEntryId -Bytes ([byte[]] $p.Value)
                }
            }
        }
    }

    $result.AccountKeys = $keys
    $result.Text        = ($texts -join "`n")
    return $result
}

<#
    Every .pst path the profile's own sections name: PR_PST_PATH, stored under a profile key as a
    binary value named by its property tag - '001f6700' (UTF-16) or '001e6700' (ANSI). Reads only.
    An empty list means none was READ, which Resolve-TierPstLayout reports as unknown, never as a
    pass.
#>
function Get-ProfilePstPath {
    param([string] $Profile)

    $root = Join-Path $profilesKey $Profile
    if (-not (Test-Path -LiteralPath $root)) { return , @() }
    $paths = @()
    foreach ($key in @(@(Get-Item -LiteralPath $root) + @(Get-ChildItem -LiteralPath $root -Recurse -ErrorAction SilentlyContinue))) {
        foreach ($pair in @(@('001f6700', [System.Text.Encoding]::Unicode), @('001e6700', [System.Text.Encoding]::ASCII))) {
            $value = $key.GetValue($pair[0], $null)
            if ($value -is [byte[]] -and $value.Length -gt 0) {
                $text = $pair[1].GetString($value).TrimEnd([char]0)
                if ($text) { $paths += $text }
            }
            elseif ($value -is [string] -and $value) { $paths += $value.TrimEnd([char]0) }
        }
    }
    return , @($paths | Sort-Object -Unique)
}

# ---------------------------------------------------------------------------------------------
# VERIFY. Reads only. Every check either passes with evidence or fails with a reason - there is
# no branch that prints success without having looked.
# ---------------------------------------------------------------------------------------------

function Invoke-Verify {
    Say '== Verify: did Outlook honour the .prf? =='
    Say ''

    $profiles = Get-ProfileNames
    Say ("Profiles present: " + $(if ($profiles.Count -gt 0) { $profiles -join ', ' } else { '<none>' }))
    Say ''

    if ($profiles -contains $ProfileName) {
        Pass 'the tier profile exists' $ProfileName
    }
    else {
        Fail 'the tier profile exists' "No profile named '$ProfileName'. Outlook did not process the .prf, or it has not been started since -Execute ran. Check that ImportPRF is still set and that FirstRun/First-Run are absent under $setupKey."
    }

    $backups = @($profiles | Where-Object { $_ -like 'Backup Of*' })
    if ($backups.Count -eq 0) {
        Pass 'no backup profile was created' 'BackupProfile=No held'
    }
    else {
        Fail 'no backup profile was created' ("Found: " + ($backups -join ', ') + ". BackupProfile=No was not honoured - try BackupProfile=False in the template, and delete these before re-running. A repeat import that keeps making these makes the guest non-reproducible.")
    }

    $dump = Get-AccountValueDump -Profile $ProfileName
    if ($dump.MailAccountNames.Count -gt 0) {
        Pass 'the account manager holds at least one account' ("{0} mail account(s): {1} ({2} entries in all - the rest are data-file and address-book wrappers)" -f $dump.MailAccountNames.Count, (($dump.MailAccountNames | ForEach-Object { "'$_'" }) -join ', '), $dump.AccountKeys.Count)
    }
    else {
        Fail 'the account manager holds at least one account' "Nothing under $acctMgrSubkey. THIS IS THE ANSWER TO THE OPEN QUESTION: Outlook 16.x did not process the .prf's internet-account sections (3, 5, 7). The PST half may still have worked - check the profile above. Route A is dead on this build; see Docs/research/pop3-account-routes.md section A.8 item 1."
    }

    if ($dump.Text -match [regex]::Escape($SinkHost)) {
        Pass 'the POP3 host reached the profile' $SinkHost
    }
    else {
        Fail 'the POP3 host reached the profile' "'$SinkHost' appears in no account value. The account subkeys exist but the section-5 values did not land - which points at a section-7 mapping problem rather than at the import as a whole."
    }

    # 0x0104 is PROP_ACCT_POP3_PORT. The VALUE is in the dump at the bottom, because a port stored
    # as a DWORD is not text and pretending to match it as text would be a check that passes for
    # the wrong reason. MEASURED 2026-09-24 on the working tier profile: the account holds NO port
    # value under either spelling - nor an SMTP one - with both ports at their defaults (110, 25),
    # while its host, user and delivery store all landed. So an absent port is the shape of a
    # default port, and it is a WARN, not a FAIL; a port the template set away from its default
    # and that did not land would be the thing to chase, and the warning says so.
    if ($dump.Text -match '(?m)^([0-9a-fA-F]{4}0104|POP3 Port)=') {
        Pass 'a POP3 port property exists' 'read its value in the dump below'
    }
    elseif ($Pop3Port -eq 110) {
        Say "  WARN a POP3 port property exists - none stored, and -Pop3Port is the default 110: measured on the working tier profile, Outlook stores no port (POP3 or SMTP) at its default. Only a NON-default port that is missing here would mean section 5's port did not land."
    }
    else {
        Fail 'a POP3 port property exists' "No port value under the account, and -Pop3Port is $Pop3Port, not the default 110. Section 5's port did not land - read the dump before concluding more than that."
    }

    # FORCEPSTPATH, as -Execute wrote it. The value outlives the import - it is per-user and
    # decides where EVERY PST Outlook mints by default lands from then on - so it is reported here,
    # where a rebuilder reads.
    $forceState = Get-ForcePstPathState
    if ($forceState.Current -and $forceState.Current.TrimEnd('\') -eq $WorkDir.TrimEnd('\') -and $forceState.Kind -eq 'ExpandString') {
        Pass 'ForcePSTPath names the work directory' ("REG_EXPAND_SZ '{0}' under {1}" -f $forceState.Current, $outlookKey)
    }
    else {
        Fail 'ForcePSTPath names the work directory' ("It reads '{0}' ({1}); -Execute writes REG_EXPAND_SZ '{2}'. Outlook mints the tier's delivery store wherever this points, so the rename step and this check would look in the wrong place." -f $forceState.Current, $forceState.Kind, $WorkDir)
    }
    foreach ($policy in @($forceState.Policies)) { Say "  WARN a POLICY ForcePSTPath is set: '$policy' - policy outranks the value above." }

    if ($dump.HasDeliveryStore) {
        Pass 'the account has a delivery store' ("{0} byte(s)" -f $dump.DeliveryStoreBytes)
    }
    else {
        Fail 'the account has a delivery store' "No '$deliveryStoreValueName' / 'Delivery Store EntryID' value, so the account is UNBOUND and NewDraft will fail. That is what a template naming a PST service and DefaultStore produces (measured 2026-09-15); this script refuses such a template, so check which one was imported: $prfPath."
    }

    # WHERE THE DELIVERY STORE IS, and whether it is the profile's only PST - read out of the
    # account's own EntryID and the profile's own PST path values. Replaces two checks that could
    # not pass on the working build: "the named PST exists" demanded tier.pst, which the
    # ForcePSTPath route never creates, and "Outlook minted no PST of its own" failed on any .pst in
    # Documents\Outlook Files - exactly where a corpus profile made FIRST keeps its /PIM store -
    # while never looking in ForcePSTPath, where the import mints. Resolve-TierPstLayout decides.
    $profilePsts = Get-ProfilePstPath -Profile $ProfileName
    Say ("  delivery store file (from its EntryID): {0}" -f $(if ($dump.DeliveryStorePath) { $dump.DeliveryStorePath } else { '<no readable path>' }))
    Say ("  PST files the tier profile references : {0}" -f $(if ($profilePsts.Count) { $profilePsts -join ', ' } else { '<none read>' }))
    $layout = Resolve-TierPstLayout -ProfilePstPaths $profilePsts -DeliveryStorePath $dump.DeliveryStorePath -WorkDir $WorkDir
    switch ($layout.Decision) {
        'Pass' {
            if (Test-Path -LiteralPath $dump.DeliveryStorePath) { Pass 'the one PST is the store Outlook minted under ForcePSTPath' $layout.Message }
            else { Fail 'the one PST is the store Outlook minted under ForcePSTPath' "$($layout.Message) - and there is no such file on disk. Outlook creates a minted store at the first start that opens the profile; has it opened the tier profile yet?" }
        }
        'Fail' { Fail 'the one PST is the store Outlook minted under ForcePSTPath' $layout.Message }
        default { Say ("  WARN the one PST is the store Outlook minted under ForcePSTPath - not shown: {0}" -f $layout.Message) }
    }

    # -- ImportPRF: the one write -Verify makes, and it needs no -Execute on purpose -------------
    # A protection that waits for somebody to remember a flag is the gap this closes. It removes
    # only a value naming THIS profile's .prf, and only once First-Run is back - proof Outlook has
    # started since -Execute, which deleted it - because removing it any earlier cancels the import.
    $setupItem = Get-Item -LiteralPath $setupKey -ErrorAction SilentlyContinue
    $importNow = $null
    $firstRunPresent = $false
    if ($null -ne $setupItem) {
        $importNow = $setupItem.GetValue('ImportPRF', $null)
        $firstRunPresent = ($null -ne $setupItem.GetValue('First-Run', $null)) -or ($null -ne $setupItem.GetValue('FirstRun', $null))
    }
    $prfProfileName = $null
    if (-not [string]::IsNullOrEmpty($importNow)) {
        try { $prfProfileName = Get-PrfProfileName -Text ([System.IO.File]::ReadAllText($importNow)) } catch { $prfProfileName = $null }
    }
    $clearance = Resolve-ImportPrfClearance -ImportPrfValue $importNow -PrfProfileName $prfProfileName `
        -ProfileName $ProfileName -ProfileExists ($profiles -contains $ProfileName) -FirstRunPresent $firstRunPresent
    Say ("  ImportPRF now: {0}; First-Run back: {1}; decision: {2}" -f $(if ([string]::IsNullOrEmpty($importNow)) { '<not set>' } else { "'$importNow'" }), $firstRunPresent, $clearance.Decision)

    switch ($clearance.Decision) {
        'NotSet' { Pass 'ImportPRF is no longer set' 'no Outlook start can re-import the tier profile' }
        'Clear' {
            Remove-ItemProperty -LiteralPath $setupKey -Name 'ImportPRF' -Force
            $readBack = (Get-ItemProperty -LiteralPath $setupKey -Name 'ImportPRF' -ErrorAction SilentlyContinue).ImportPRF
            if ([string]::IsNullOrEmpty($readBack)) {
                Pass 'ImportPRF is no longer set' 'it was still set after the import had run, so -Verify removed it and read the removal back'
            }
            else {
                Fail 'ImportPRF is no longer set' "It was removed and still reads '$readBack'. Do not start Outlook until that is resolved: the template carries OverwriteProfile=Yes, so a start that honours it rebuilds the tier profile."
            }
        }
        'KeepPending' { Fail 'ImportPRF is no longer set' $clearance.Message }
        default { Say ("  WARN ImportPRF is no longer set - {0}" -f $clearance.Message) }
    }

    Say ''
    Say '-- account value dump (this is the evidence; read it before trusting any line above) --'
    if ($dump.Text) { Say $dump.Text } else { Say '<empty>' }
}

# ---------------------------------------------------------------------------------------------
# PLAN and EXECUTE.
# ---------------------------------------------------------------------------------------------

function Get-RenderedPrf {
    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        throw "PRF template not found: $TemplatePath. It ships beside this script as tier-profile-forcepst.prf - stage it beside the script, or pass -TemplatePath."
    }
    $text = Get-Content -LiteralPath $TemplatePath -Raw

    # REFUSED BEFORE ANYTHING IS WRITTEN: a template that names the store leaves the account
    # unbound, and nothing downstream says so until NewDraft fails. See Test-TierTemplateShape.
    $shapeProblems = Test-TierTemplateShape -Text $text
    if ($shapeProblems.Count -gt 0) {
        throw ("REFUSING the template $TemplatePath - it is not the shape that produces a bound POP3 account:`n  - " +
            ($shapeProblems -join "`n  - ") +
            "`nUse the default, tier-profile-forcepst.prf. (tier-profile.prf beside it is the RETIRED shape, kept as evidence.)")
    }

    $map = @{
        '{{PROFILE_NAME}}'      = $ProfileName
        '{{ACCOUNT_NAME}}'      = $AccountName
        '{{POP3_HOST}}'         = $SinkHost
        '{{SMTP_HOST}}'         = $SinkHost
        '{{POP3_USER}}'         = $Pop3User
        '{{EMAIL_ADDRESS}}'     = $EmailAddress
        '{{DISPLAY_NAME}}'      = $DisplayName
        '{{POP3_PORT}}'         = [string] $Pop3Port
        '{{SMTP_PORT}}'         = [string] $SmtpPort
    }
    foreach ($k in $map.Keys) {
        $text = $text.Replace($k, $map[$k])
    }

    $left = [regex]::Matches($text, '\{\{[A-Z0-9_]+\}\}')
    if ($left.Count -gt 0) {
        $names = @($left | ForEach-Object { $_.Value } | Sort-Object -Unique)
        throw ("The template still holds unsubstituted tokens after rendering: " + ($names -join ', ') +
            ". A token Outlook cannot parse is a value written silently wrong, so this refuses rather than shipping it.")
    }

    # ASCII, CRLF, no BOM. Section A.6 item 4: every published sample is plain 8-bit text, and a
    # UTF-8 BOM would prefix the first line.
    $normalised = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    return $normalised
}

function Show-NextSteps {
    $minted = Join-Path $WorkDir 'Outlook.pst'
    Say 'NEXT - none of this is done by this script:'
    Say '  1. Start Outlook once, IN SESSION 1 (Register-InteractiveTask.ps1), and let it finish starting.'
    Say '     It imports the .prf, creates the profile, and mints the delivery store under ForcePSTPath.'
    Say '     Expect the POP3 password dialog on screen; it does not block COM.'
    Say "  2. .\Rename-OutlookStore.ps1 -StoreFilePath '$minted' -DisplayName '$StoreDisplayName' -Execute"
    Say "     (Outlook names the store it mints 'Outlook Data File'; the hub has to be named as the address.)"
    Say '  3. .\New-TierProfile.ps1 -Verify'
    Say 'This script does not start Outlook: importing is a startup-time action, and owning'
    Say "Outlook's lifetime from a setup script is how a guest ends up with a zombie OUTLOOK.EXE."
}

function Show-Plan {
    Say '== Plan (nothing below has been done) =='
    Say ''
    Say "  profile name        : $ProfileName"
    Say "  email address       : $EmailAddress"
    Say "  hub store name      : $StoreDisplayName  (set AFTER the import, by Rename-OutlookStore.ps1)"
    Say "  POP3                : $SinkHost`:$Pop3Port  (no SSL, no SPA, no stored password)"
    Say "  SMTP                : $SinkHost`:$SmtpPort  (no SSL, no auth)"
    Say "  PRF template        : $TemplatePath"
    Say "  PRF written to      : $prfPath"
    Say "  delivery store      : minted by Outlook under ForcePSTPath, expected $(Join-Path $WorkDir 'Outlook.pst')"
    Say ''
    Say '  Registry it would write:'
    Say "    $outlookKey"
    Say "      ForcePSTPath REG_EXPAND_SZ = $WorkDir   (per-user: every PST Outlook mints by default lands here afterwards)"
    Say "    $setupKey"
    Say "      ImportPRF    REG_SZ        = $prfPath"
    Say '      First-Run    DELETED       (ImportPRF is ignored while it exists)'
    Say '      FirstRun     DELETED       (same)'
    Say ''
    Say '  It writes nothing else, and it does NOT start Outlook.'
    Say ''
    Show-NextSteps
}

function Invoke-Execute {
    Say '== Execute =='
    Say ''

    # -- preflight: assert the starting state rather than assuming it -------------------------

    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw ("OUTLOOK.EXE is running (pid " + (($running | ForEach-Object { $_.Id }) -join ', ') +
            "). A .prf is read at startup, so importing into a running Outlook does nothing visible now " +
            "and something surprising later. Close Outlook and re-run. This script will not kill it.")
    }

    if (-not (Test-Path -LiteralPath $outlookKey)) {
        throw ("Not found: $outlookKey. Office $OfficeVersion does not look installed for this user. " +
            "Pass -OfficeVersion if the guest has a different major version.")
    }

    $existing = Get-ProfileNames
    if ($existing -contains $ProfileName -and -not $Force) {
        throw ("A profile named '$ProfileName' already exists. This script expects a checkpoint where it does not. " +
            "Re-run with -Force to import over it - the template's OverwriteProfile=Yes + BackupProfile=No are " +
            "what should make that idempotent, and -Verify will tell you whether they did.")
    }
    Say ("Profiles before: " + $(if ($existing.Count -gt 0) { $existing -join ', ' } else { '<none>' }))

    # ForcePSTPath is decided BEFORE anything is written: a policy that would send the mint
    # elsewhere is a refusal, not a surprise at the rename step.
    $forceState = Get-ForcePstPathState
    $forcePlan = Resolve-ForcePstPathPlan -Current $forceState.Current -CurrentKind $forceState.Kind -Wanted $WorkDir -PolicyValues $forceState.Policies
    if ($forcePlan.Decision -like 'Refuse*') { throw $forcePlan.Message }
    Say "ForcePSTPath: $($forcePlan.Message)"

    # -- render and write the .prf ------------------------------------------------------------

    # Renders AND refuses a template that names the store (Test-TierTemplateShape) - before the
    # first write below.
    $prfText = Get-RenderedPrf

    # Section A.6 item 1, and ForcePSTPath's own precondition: the directory Outlook mints into
    # must already exist - Outlook creates the .pst, not the folder.
    if (-not (Test-Path -LiteralPath $WorkDir)) {
        New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
        Say "Created $WorkDir"
    }
    else {
        Say "$WorkDir already exists"
        $already = @(Get-ChildItem -LiteralPath $WorkDir -Filter '*.pst' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
        if ($already.Count -gt 0) {
            Say ("  WARN it already holds: {0}. If one is called Outlook.pst, the store Outlook mints now may get another name - read the path -Verify reports before renaming." -f ($already -join ', '))
        }
    }

    [System.IO.File]::WriteAllText($prfPath, $prfText, (New-Object System.Text.ASCIIEncoding))
    Say "Wrote $prfPath ($($prfText.Length) chars, ASCII, CRLF)"

    $readBack = [System.IO.File]::ReadAllText($prfPath)
    if ($readBack -ne $prfText) {
        throw "The .prf read back differently from what was written. Refusing to continue."
    }
    if ($readBack -notmatch '(?m)^ProfileName=') {
        throw "The written .prf has no ProfileName= line. The template is not what this script expects."
    }
    Pass 'the .prf reads back byte-identical' $prfPath

    # -- registry -----------------------------------------------------------------------------

    # ForcePSTPath FIRST, then ImportPRF: the value has to be in place before the start that
    # imports, because that start is the one that mints. The Outlook key exists - checked above -
    # so no New-Item here: New-Item -Force on an EXISTING registry key is not a no-op, and the
    # hand-run scripts that set this value before 2026-09-24 did exactly that to the Outlook key.
    if ($forcePlan.Decision -ne 'AlreadySet') {
        New-ItemProperty -LiteralPath $outlookKey -Name $forcePstName -Value $WorkDir -PropertyType ExpandString -Force | Out-Null
    }
    $forceBack = Get-ForcePstPathState
    if ($forceBack.Current -eq $WorkDir -and $forceBack.Kind -eq 'ExpandString') {
        Pass 'ForcePSTPath reads back' ("REG_EXPAND_SZ '{0}' under {1}" -f $forceBack.Current, $outlookKey)
    }
    else {
        throw ("ForcePSTPath was written as REG_EXPAND_SZ '$WorkDir' and reads back as {0} '{1}'. Do not start Outlook: the store would be minted somewhere the next steps do not look." -f $forceBack.Kind, $forceBack.Current)
    }

    if (-not (Test-Path -LiteralPath $setupKey)) {
        New-Item -Path $setupKey -Force | Out-Null
        Say "Created $setupKey"
    }

    New-ItemProperty -LiteralPath $setupKey -Name 'ImportPRF' -Value $prfPath -PropertyType String -Force | Out-Null
    Say "Set ImportPRF = $prfPath"

    foreach ($name in @('First-Run', 'FirstRun')) {
        $present = $null
        try { $present = Get-ItemProperty -LiteralPath $setupKey -Name $name -ErrorAction Stop } catch { $present = $null }
        if ($null -ne $present) {
            Remove-ItemProperty -LiteralPath $setupKey -Name $name -Force
            Say "Deleted $name"
        }
        else {
            Say "$name already absent"
        }
    }

    # -- verify what we just wrote, rather than assuming the writes took ----------------------

    $back = Get-ItemProperty -LiteralPath $setupKey
    if ($back.ImportPRF -eq $prfPath) {
        Pass 'ImportPRF reads back' $prfPath
    }
    else {
        Fail 'ImportPRF reads back' ("Expected '$prfPath', found '" + $back.ImportPRF + "'.")
    }

    foreach ($name in @('First-Run', 'FirstRun')) {
        $still = $back.PSObject.Properties | Where-Object { $_.Name -eq $name }
        if ($null -eq $still) {
            Pass "$name is absent" 'ImportPRF will be honoured at the next Outlook start'
        }
        else {
            Fail "$name is absent" "It is still present, and ImportPRF is ignored while it exists - with no diagnostic anywhere."
        }
    }

    Say ''
    Show-NextSteps
}

# ---------------------------------------------------------------------------------------------
# Entry point.
# ---------------------------------------------------------------------------------------------

Say "New-TierProfile.ps1 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Say ''

if ($Verify -and $Execute) {
    throw 'Pass -Execute or -Verify, not both: verifying in the same run as writing would assert against a profile Outlook has not read yet, and would pass for the wrong reason.'
}

if ($Verify) {
    Invoke-Verify
}
elseif ($Execute) {
    Invoke-Execute
}
else {
    Show-Plan
    Say ''
    Say 'Dry run. Nothing written. Re-run with -Execute.'
}

Say ''
if ($script:Failures.Count -gt 0) {
    Say ("{0} check(s) FAILED:" -f $script:Failures.Count)
    foreach ($f in $script:Failures) { Say "  - $f" }
    Save-Log
    exit 1
}

if ($Verify -or $Execute) {
    Say 'All checks passed.'
}
Save-Log
exit 0
