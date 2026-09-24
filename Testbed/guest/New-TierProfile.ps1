#Requires -Version 5.1
<#
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
    imported .prf names no PST service.

.SYNOPSIS
    Creates the testbed TIER profile - one Unicode PST plus one POP3 account on a loopback mail
    sink - by importing a .prf file, and then PROVES whether Outlook honoured it.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    NEVER run this on the maintainer's workstation: it writes the Outlook Setup key and causes
    the next Outlook start to rebuild the mail profile.

    WHY THIS EXISTS. The tier profile needs a POP3/SMTP account pointed at a loopback sink. The
    Outlook object model is read-only for accounts, Extended MAPI can no longer create POP3
    services, and the one off-the-shelf component that can is excluded by the repository's
    Dependencies rule. A .prf file plus the ImportPRF registry value is the remaining free,
    documented, non-interactive route. The full evidence - with a source for every key - is in
    .work/pop3-account-routes.md section A.

    THIS IS AN EXPERIMENT, AND IT IS WRITTEN TO REPORT A FAILURE RATHER THAN HIDE ONE. The route
    turns on three things nobody has verified on a 16.x build, and each has its own assertion:

      1. Does Outlook 16.x process a PRF's INTERNET ACCOUNT sections (3, 5 and 7) at all? Every
         literal POP3 .prf Microsoft ever published is 2000-2007 era, and the Office Customization
         Tool's own page says a legacy .prf imports "provided that the profile defines only MAPI
         services" - which a POP3 account is not. -Verify fails loudly if no account subkey
         appears under the new profile.
      2. Does `[General] DefaultStore=Service1` bind the POP3 account's delivery store? It cannot
         be expressed as a property: PROP_ACCT_DELIVERY_STORE is PT_BINARY and a PRF carries only
         PT_UNICODE / PT_LONG / PT_BOOLEAN. DefaultStore names a SERVICE and leaves the processor
         to resolve it, which is why the gap is not automatically fatal - but Outlook 2010+ mints
         its OWN pst per POP3 account by default. -Verify reports whether 00180102 exists and
         whether a stray .pst appeared under Documents\Outlook Files.
      3. Is the import genuinely silent? ImportPRF is documented as applying at the next Outlook
         start with no prompt, but ONLY while `FirstRun` and `First-Run` are absent from the
         Setup key. This script deletes both, and says so.

    IT DOES NOT START OUTLOOK, AND THAT IS DELIBERATE. The import happens at the next Outlook
    start, and a script that starts Outlook then owns Outlook's lifetime - which this project
    handles carefully and never from ad-hoc code (never taskkill; release COM references BEFORE
    any quit). So the flow is three steps, and the middle one is yours:

        .\New-TierProfile.ps1                 # dry run: prints the plan and changes nothing
        .\New-TierProfile.ps1 -Execute        # writes the .prf and the registry values
        <start Outlook once, let it settle, close it>
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
    any profile other than the one named. Its registry writes are all under HKCU\...\Outlook\Setup
    and each is named before it is made: -Execute sets ImportPRF and deletes First-Run and FirstRun
    (creating the Setup key if it is missing); -Verify may remove ImportPRF again, and only in the
    one case Resolve-ImportPrfClearance calls Clear.

.PARAMETER ProfileName
    The MAPI profile to create. Also the name asserted to be the ONLY profile of that name
    afterwards.

.PARAMETER StoreDisplayName
    The PST service's display name. The live suite treats a hub store's display name as an SMTP
    address, so this and -EmailAddress are usually the same string; see
    Docs/live-tier-on-the-vm.md section 2.6 and its open question about `@` in a store name.

.PARAMETER WorkDir
    Where the rendered .prf and the .pst go. Deliberately space-free: the ImportPRF registry
    value takes a raw unquoted path.

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
    Run the ImportPRF decision tests and exit. Touches nothing - no registry, no files, no guard.

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

    Write-Host 'New-TierProfile self-test. Nothing is read, nothing is written, nothing is started.'
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
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Guest-only: the guard, every registry read and write, the .prf render and'
    Write-Host 'read-back, whether Outlook honours the file, and the ImportPRF removal -Verify makes.'
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

$prfPath = Join-Path $WorkDir 'tier-profile.prf'
$pstPath = Join-Path $WorkDir 'tier.pst'

if (-not $TemplatePath) {
    $TemplatePath = Join-Path $PSScriptRoot 'tier-profile.prf'
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
                if ($p.Value -is [array]) { $result.DeliveryStoreBytes = $p.Value.Count }
            }
        }
    }

    $result.AccountKeys = $keys
    $result.Text        = ($texts -join "`n")
    return $result
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
        Fail 'the account manager holds at least one account' "Nothing under $acctMgrSubkey. THIS IS THE ANSWER TO THE OPEN QUESTION: Outlook 16.x did not process the .prf's internet-account sections (3, 5, 7). The PST half may still have worked - check the profile above. Route A is dead on this build; see .work/pop3-account-routes.md section A.8 item 1."
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

    if ($dump.HasDeliveryStore) {
        Pass 'the account has a delivery store' ("{0} byte(s) - read the log to see WHICH store" -f $dump.DeliveryStoreBytes)
    }
    else {
        Fail 'the account has a delivery store' "No $deliveryStoreValueName value. [General] DefaultStore=Service1 did not bind a per-account delivery store. The fallback is section A.5 of .work/pop3-account-routes.md: drop DefaultStore, set ForcePSTPath, and discover the minted .pst by account instead of by path."
    }

    # WHICH ROUTE WAS IMPORTED decides whether there is a named PST to find at all. The route that
    # WORKS on both guests - tier-profile-forcepst.prf - names NO PST service and lets Outlook mint
    # the store under ForcePSTPath, so tier.pst never exists there and this check used to FAIL on
    # the one working build (found 2026-09-24, OAI-UNINDEXED: the store is C:\OutlookAI-Tier\
    # Outlook.pst). A PathToPersonalFolders= line that is a path, not a section-6 PT_ mapping, is
    # what a named-PST .prf carries; the rendered file is read rather than the template guessed at.
    $namesPst = $false
    if (Test-Path -LiteralPath $prfPath) {
        $namesPst = @(Get-Content -LiteralPath $prfPath | Where-Object { $_ -match '^\s*PathToPersonalFolders=(?!PT_)\S' }).Count -gt 0
    }
    if (-not $namesPst -and (Test-Path -LiteralPath $prfPath)) {
        Say "  --   the named PST exists - not applicable: $prfPath names no PST service (the ForcePSTPath route), so Outlook minted the store and the delivery-store check above is the one that says the PST half worked."
    }
    elseif (Test-Path -LiteralPath $pstPath) {
        Pass 'the named PST exists' $pstPath
    }
    else {
        Fail 'the named PST exists' "$pstPath was not created. Outlook creates the file lazily on first use, so this can be a false alarm on a profile that has never been opened - but with the delivery-store check above it is the pair that says whether the PST half worked."
    }

    $strayDir = Join-Path $env:USERPROFILE 'Documents\Outlook Files'
    $stray = @()
    if (Test-Path -LiteralPath $strayDir) {
        $stray = @(Get-ChildItem -LiteralPath $strayDir -Filter '*.pst' -File -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Name })
    }
    if ($stray.Count -eq 0) {
        Pass 'Outlook minted no PST of its own' $strayDir
    }
    else {
        Fail 'Outlook minted no PST of its own' ("Found in $strayDir : " + ($stray -join ', ') + ". This is the documented Outlook 2010+ behaviour - a POP3 account gets its own data file unless told otherwise - and it means DefaultStore did not take. Not necessarily fatal: see the ForcePSTPath fallback in section A.5.")
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
        throw "PRF template not found: $TemplatePath. It ships beside this script as tier-profile.prf; pass -TemplatePath if it lives elsewhere."
    }
    $text = Get-Content -LiteralPath $TemplatePath -Raw

    $map = @{
        '{{PROFILE_NAME}}'      = $ProfileName
        '{{STORE_DISPLAY_NAME}}' = $StoreDisplayName
        '{{PST_PATH}}'          = $pstPath
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

function Show-Plan {
    Say '== Plan (nothing below has been done) =='
    Say ''
    Say "  profile name        : $ProfileName"
    Say "  store display name  : $StoreDisplayName"
    Say "  email address       : $EmailAddress"
    Say "  POP3                : $SinkHost`:$Pop3Port  (no SSL, no SPA, no stored password)"
    Say "  SMTP                : $SinkHost`:$SmtpPort  (no SSL, no auth)"
    Say "  PRF template        : $TemplatePath"
    Say "  PRF written to      : $prfPath"
    Say "  PST named as        : $pstPath"
    Say ''
    Say '  Registry it would write:'
    Say "    $setupKey"
    Say "      ImportPRF   REG_SZ   = $prfPath"
    Say '      First-Run   DELETED  (ImportPRF is ignored while it exists)'
    Say '      FirstRun    DELETED  (same)'
    Say ''
    Say '  It writes nothing else, and it does NOT start Outlook.'
    Say ''
    Say '  Then, by hand: start Outlook once, let it settle, close it. Then:'
    Say '    .\New-TierProfile.ps1 -Verify'
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

    # -- render and write the .prf ------------------------------------------------------------

    $prfText = Get-RenderedPrf

    if (-not (Test-Path -LiteralPath $WorkDir)) {
        New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
        Say "Created $WorkDir"
    }
    else {
        Say "$WorkDir already exists"
    }

    # Section A.6 item 1: the directories in the path to the personal folders must already exist.
    $pstDir = Split-Path -Parent $pstPath
    if (-not (Test-Path -LiteralPath $pstDir)) {
        New-Item -ItemType Directory -Path $pstDir -Force | Out-Null
        Say "Created $pstDir"
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
    Say 'NEXT, BY HAND: start Outlook once, let it finish starting, then close it.'
    Say 'Then run:  .\New-TierProfile.ps1 -Verify'
    Say 'This script does not start Outlook: importing is a startup-time action, and owning'
    Say "Outlook's lifetime from a setup script is how a guest ends up with a zombie OUTLOOK.EXE."
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
