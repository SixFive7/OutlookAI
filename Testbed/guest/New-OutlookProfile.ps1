#Requires -Version 5.1
<#
    ============================================================================================
    THE MAPI ROUTE IS DEAD ON THIS BUILD. THIS SCRIPT IMPORTS A .prf INSTEAD.
    ============================================================================================

    WHAT THIS SCRIPT USED TO BE, AND WHY IT NEVER WORKED. Until now it created the profile with
    `IProfAdmin::CreateProfile` and added each PST with `IMsgServiceAdmin::CreateMsgService` +
    `ConfigureMsgService`, through Testbed/guest/OutlookMapiInterop.ps1. It carried a banner
    saying it had never been executed. The interop HAS now been executed - by a sibling script,
    on a guest, on 2026-09-16 - and its very first call failed:

        Exception calling "MAPIAdminProfiles" with "2" argument(s): "Unable to cast COM object of
        type 'System.__ComObject' to interface type 'OutlookAI.Testbed.IProfAdmin' ... No such
        interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE))."

    on guest OAI-UNINDEXED, user vmadmin, 64-bit elevated Windows PowerShell 5.1, Office LTSC
    2024 (ProPlus2024Volume / PerpetualVL2024) build 16.0.17932.20884. `MAPIInitialize` and
    `MAPIAdminProfiles` both SUCCEEDED; only the QueryInterface for `IID_IProfAdmin` failed, and
    the cause is not established. So this script's OLD route was not merely unrun: both of its
    calls sat behind a call that is measured broken, and a from-scratch guest build would have
    died on the first one. That is what this rewrite removes.

    ============================================================================================
    WHAT IT DOES NOW, AND THE HONEST EVIDENCE FOR EVERY PART OF IT
    ============================================================================================

    It writes an Outlook Profile (.prf) file and points `ImportPRF` at it. Outlook reads that at
    its NEXT START and builds the profile - and, in the same pass, every PST this script was
    asked for, each under EXACTLY the display name asked for.

    THE ROUTE IS NOT A GUESS. It is the same mechanism `New-TierProfile.ps1` uses, which is
    MEASURED WORKING ON BOTH GUESTS (2026-09-15 and 2026-09-16, Office LTSC 2024 16.0.17932) -
    first attempt on the second guest, from the committed scripts, untouched by hand. The parts
    this script relies on were each measured on that run:

      * the profile appeared, created by Outlook at startup from `ImportPRF`;
      * the PST THE FILE NAMED was the one Outlook used - `C:\OutlookAI-Tier\tier.pst` - and
        Outlook minted no data file of its own;
      * `Store.DisplayName` read back OVER COM as `tier@vm.invalid`, the exact string the .prf's
        `[Service1] Name=` carried, `@` and all;
      * no first-run dialog, after Set-OfficeFirstRunSuppressed.ps1.

    THAT LAST ONE ALSO CLOSES A QUESTION THIS SCRIPT USED TO POINT AT. Testbed/README.md section
    6 item 10 and Docs/live-tier-on-the-vm.md section 8 item 2 asked whether Outlook accepts `@`
    in a store display name, and `Add-OutlookPstStore.ps1 -NameProbe` existed to settle it. It is
    settled, twice and independently: through the .prf above, and through the root-folder rename
    in Rename-OutlookStore.ps1 ('Outlook Data File' -> 'tier@vm.invalid', measured 2026-09-15).
    The probe is gone; see that script's banner.

    WHAT IS STILL INFERENCE, STATED SO NOBODY READS THIS AS ALL-MEASURED. The .prf that was
    measured carried ONE PST service and ONE POP3 account. This script emits a .prf with N PST
    services and NO account sections at all, which is a strict SUBSET of the measured file plus a
    repetition of its one service block. Two things about it are therefore INFERRED rather than
    measured, and `-Verify` asserts both rather than assuming them:

      1. that a .prf whose `[Internet Account List]` is empty produces a profile with ZERO mail
         accounts. Corroboration, and it is decent: on that same guest a profile that had a PST
         store and no internet account held NOTHING under the account-manager key
         (9375CFF0413111d3B88A00104B2A6676), which is what Build-Corpus.ps1's own preflight reads
         and what its comment calls "the account-less shape, and it is not a guess".
      2. that `[Service List]` may name `Unicode Personal Folders` more than once. `UniqueService=No`
         - which the measured template carries verbatim - is the key whose documented job is to
         permit exactly that. It has not been exercised with two PSTs on this build.

      IF EITHER TURNS OUT FALSE, THE FALLBACK IS NAMED RATHER THAN LEFT TO BE INVENTED. For (1):
      `outlook.exe /PIM <name>` is MEASURED on this Office build to create an account-less
      profile (`accounts=0`) - it just cannot name a store. For (2): run this script once per
      store is not possible (a second import rebuilds the profile), so use one .prf for the first
      store and `Add-OutlookPstStore.ps1` for the rest, which is the AddStoreEx-plus-rename route
      and needs no .prf at all.

    WHY AN ACCOUNT-LESS PROFILE IS MANDATORY AND NOT A CONVENIENCE. Docs/live-tier-on-the-vm.md
    section 1.2: corpus-build refuses any profile that has a mail account, with no override flag,
    because a build creates unsent items in bulk and the first real run left 5,532 of them in an
    Outbox - inert only because that profile could not send. The predicate is
    `Accounts.Count == 0` read over COM (CorpusSafety.EvaluateProfile over
    ComCorpusMailbox.ReadProfileFacts), fail-closed, so an unreadable answer refuses too.

    WHAT IT STILL CANNOT MAKE: a mail account. Nothing free can, and that has not changed - see
    THE WALL below. The tier profile's POP3 account has its own script (New-TierProfile.ps1) and
    its own measured recipe; this script is the CORPUS half.

    ============================================================================================
    THE SHAPE OF A RUN, AND WHY IT IS THREE STEPS RATHER THAN ONE
    ============================================================================================

        .\New-OutlookProfile.ps1 -Preflight                      # read-only; run this first
        .\New-OutlookProfile.ps1 -Name CorpusProfile -Store ... -Execute
        <start Outlook once, let it settle, close it>            # YOURS. See below.
        .\New-OutlookProfile.ps1 -Name CorpusProfile -Store ... -Verify
        .\New-OutlookProfile.ps1 -ClearImportPrf -Execute        # see 'THE RE-IMPORT TRAP'

    THE MIDDLE STEP IS DELIBERATELY NOT AUTOMATED, and New-TierProfile.ps1 says the same thing
    for the same reason: importing is a startup-time action, and a script that starts Outlook then
    owns Outlook's lifetime - which this project handles carefully and never from ad-hoc code
    (never taskkill; release COM references BEFORE any quit). A setup script that starts Outlook
    is how a guest ends up with a zombie OUTLOOK.EXE.

    THE RE-IMPORT TRAP, and it is the one thing in here that could cost a corpus. `ImportPRF` is
    read at every Outlook start, and this .prf carries `OverwriteProfile=Yes`. WHETHER OUTLOOK
    CLEARS THE VALUE AFTER PROCESSING IT HAS NOT BEEN ESTABLISHED ON THIS BUILD - New-TierProfile
    .ps1 never checked, and neither has anything else. If it does NOT clear it, then every later
    Outlook start re-imports and REBUILDS the profile, which would detach a corpus store that had
    since been filled. `-Verify` reads the value back and says loudly if it is still there;
    `-ClearImportPrf -Execute` removes it. Do that before building a corpus, not after.

    ============================================================================================
    THE WALL, stated here so nobody goes looking for the missing half
    ============================================================================================

    There is NO free programmatic route to creating a POP3 account with a bound delivery store,
    and the evidence runs three ways:

      * the object model has no Accounts.Add, and Account.DeliveryStore is read-only - so the OM
        cannot set 'deliver new messages to', which section 2.8 calls the failure to bet on;
      * MAPI cannot, because POP3 stopped being a MAPI message service: account administration
        moved behind the undocumented IOlkAccountManager;
      * the registry has no published working recipe on 16.x, and the stored password blobs are
        DPAPI-sealed per Windows user per machine, so they cannot be authored offline.

    A .prf CAN create the account and cannot BIND it: PROP_ACCT_DELIVERY_STORE (00180102) is a
    binary EntryID and a text file cannot carry one. The way past that was found and is measured -
    name NO PST service in the .prf and let Outlook MINT the account's store, because a store
    Outlook mints is a store Outlook binds - and it lives in New-TierProfile.ps1 with
    tier-profile-forcepst.prf. It is the TIER profile's recipe and it is the opposite of this
    script's: this one names its stores precisely because it has no account to bind them to.

    ============================================================================================

    OUTLOOK MUST NOT BE RUNNING for -Execute. A .prf is read at startup, so importing into a
    running Outlook does nothing visible now and something surprising later. The script refuses
    and does NOT kill it: mailbox safety rule 7 forbids taskkill on OUTLOOK.EXE outright.

    WHICH OFFICE HIVE. Detected, never hardcoded, by the same rule as
    Set-DefaultOutlookProfile.ps1 and the shipped OfficeVersions.IsOutlookHive: a real Outlook key
    has at least one VALUE, or at least one subkey that is not 'Resiliency'. That rule exists
    because Installer.iss writes a resiliency exemption under 15.0, 16.0 AND 17.0 on every
    install, so on any machine this product has touched a bare key-exists probe answers 16.0 just
    as confidently on an Outlook 2013 machine. The PROBE ORDER is OfficeVersions.Supported's -
    16.0, then 17.0, then 15.0 - and NOT "highest number first". The rule is copied rather than
    shared on purpose: OutlookMapiInterop.ps1 is dot-sourced AFTER these functions are defined, so
    a function of the same name living there would silently replace this one.

    WHAT IS TESTED WITHOUT A GUEST. -SelfTest exercises every decision in this script - the hive
    shape rule and probe order, the -Store argument grammar with every one of its refusals, the
    whole .prf that gets written, and the verify verdict with each of its failures - against
    synthetic inputs. It reads no registry, starts no process, writes no file and makes no COM
    call, so it is safe on any machine including the maintainer's workstation. What it cannot
    cover is guest-only, and it prints that list so nobody mistakes a green run for full coverage.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER Name
    Profile name to create. Also the name -Verify asserts against.

.PARAMETER Store
    Repeatable. 'DisplayName=C:\path\to.pst'. The display name is what Outlook shows and what the
    tests match on; the path is where the file goes and MUST be absolute. Omit entirely for a
    store-less profile - which is legal and is almost certainly not what you want; see -Verify.

.PARAMETER MakeDefault
    Write DefaultProfile=Yes into the .prf, so the imported profile becomes the default. It does
    NOT suppress the profile prompt - that is PickLogonProfile, a separate setting, and
    Set-DefaultOutlookProfile.ps1 is what does it. A prompting profile cannot be driven over COM.

.PARAMETER PrfPath
    Where the rendered .prf goes. Defaults under C:\OutlookAI-Profiles. It must contain NO SPACE:
    the ImportPRF registry value takes a raw unquoted path.

.PARAMETER Preflight
    Report what the import route needs, and change nothing: the Office hive, the Setup key, the
    profiles already present, the current ImportPRF value, and whether Outlook is running. It
    makes no COM call and no MAPI call - there is no MAPI left in this project to call.

.PARAMETER Verify
    After Outlook has been started once: read the profile hive back and assert that the import
    did what the .prf said. Registry and filesystem only; add -WithOutlook for the COM half.

.PARAMETER WithOutlook
    With -Verify: also bind Outlook and read Session.Stores and Accounts.Count. That is the only
    check that answers 'what will the tests see'. It requires the profile to be the DEFAULT one.

.PARAMETER ClearImportPrf
    Remove the ImportPRF value so the next Outlook start does not re-import. Needs -Execute.

.PARAMETER SelfTest
    Run the decision tests and exit. Touches nothing: no registry, no files, no processes.

.PARAMETER OfficeVersion
    Force the Office major whose hive is read and written, e.g. '16.0'.

.PARAMETER Execute
    Without it, nothing is written.

.EXAMPLE
    .\New-OutlookProfile.ps1 -SelfTest
    .\New-OutlookProfile.ps1 -Preflight
    .\New-OutlookProfile.ps1 -Name CorpusProfile -Store 'Corpus A=C:\OutlookAI-Q5\pst\corpus-a.pst' -MakeDefault -Execute
    .\New-OutlookProfile.ps1 -Name CorpusProfile -Store 'Corpus A=C:\OutlookAI-Q5\pst\corpus-a.pst' -Verify -WithOutlook
    .\New-OutlookProfile.ps1 -ClearImportPrf -Execute
#>
[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Build')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Verify')]
    [string] $Name,

    [Parameter(ParameterSetName = 'Build')]
    [Parameter(ParameterSetName = 'Verify')]
    [string[]] $Store = @(),

    [Parameter(ParameterSetName = 'Build')] [switch] $MakeDefault,
    [Parameter(ParameterSetName = 'Build')] [string] $PrfPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Verify')] [switch] $Verify,
    [Parameter(ParameterSetName = 'Verify')] [switch] $WithOutlook,

    [Parameter(Mandatory = $true, ParameterSetName = 'Preflight')] [switch] $Preflight,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest,
    [Parameter(Mandatory = $true, ParameterSetName = 'ClearImportPrf')] [switch] $ClearImportPrf,

    [string] $OfficeVersion,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

# The Office majors this product supports, IN PROBE ORDER, mirrored from OfficeVersions.Supported.
$script:SupportedOfficeVersions = @('16.0', '17.0', '15.0')

# The subkey this product's own installer creates under every supported major, which is why a
# bare key-exists probe proves nothing. OfficeVersions.InstallerFootprintSubKeyName.
$script:InstallerFootprintSubKeyName = 'Resiliency'

$script:OfficeRootKeyPath = 'HKCU:\Software\Microsoft\Office'

# The account-manager subkey under a profile. Stable since Outlook 2002, and the same key
# Build-Corpus.ps1's preflight counts and SignatureCatalog.ReadProfileAccountValueSets walks.
$script:AccountManagerSubKeyName = '9375CFF0413111d3B88A00104B2A6676'

# The section-6 block name a Unicode PST service maps to, and the address book service. Both are
# spelled exactly as tier-profile.prf spells them, which is the file Outlook 16.x measurably
# processed on 2026-09-15.
$script:PstServiceBlockName = 'Unicode Personal Folders'
$script:AddressBookBlockName = 'Outlook Address Book'

# =============================================================================================
# PURE DECISIONS. No registry, no files, no processes, no COM, no output. Everything this script
# decides is decided here, so -SelfTest can decide it too without a guest.
# =============================================================================================

<#
    Whether an ...\Office\<major>\Outlook key with these values and subkeys is a REAL Outlook
    hive, or just the shell this product's installer leaves behind. Mirrors
    OfficeVersions.IsOutlookHive exactly, case-insensitive footprint comparison included.
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
    Picks the hive to work in, from candidates shaped
    @{ Version = '16.0'; Path = '...'; ValueNames = @(); SubKeyNames = @() }.

    Returns Chosen (a candidate, or $null), RealHives (every candidate passing the shape rule) and
    Problem (the refusal text, when Chosen is $null).

    -RequestedVersion overrides the probe order and the supported list both, and does NOT override
    the shape rule: writing ImportPRF into a key Outlook never reads is a silent no-op, and a
    silent no-op here means an operator starts Outlook, sees no new profile, and has nothing at
    all to read to find out why.
#>
function Select-OutlookHive {
    param([object[]] $Candidates, [string] $RequestedVersion)

    $real = @()
    $present = @()
    if ($null -ne $Candidates) {
        foreach ($candidate in $Candidates) {
            $present += $candidate.Version
            if (Test-IsOutlookHive -ValueNames $candidate.ValueNames -SubKeyNames $candidate.SubKeyNames) {
                $real += $candidate
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
        $named = $null
        if ($null -ne $Candidates) {
            foreach ($candidate in $Candidates) {
                if ($candidate.Version -eq $RequestedVersion) { $named = $candidate }
            }
        }

        if ($null -eq $named) {
            $presentText = '<none>'
            if ($present.Count -gt 0) { $presentText = ($present -join ', ') }
            return [pscustomobject]@{
                Chosen    = $null
                RealHives = $real
                Problem   = "REFUSING: -OfficeVersion $RequestedVersion was asked for, and $script:OfficeRootKeyPath\$RequestedVersion\Outlook does not exist. Version keys present: $presentText. Drop -OfficeVersion and let the detection choose."
            }
        }

        if (-not (Test-IsOutlookHive -ValueNames $named.ValueNames -SubKeyNames $named.SubKeyNames)) {
            return [pscustomobject]@{
                Chosen    = $null
                RealHives = $real
                Problem   = "REFUSING: -OfficeVersion $RequestedVersion was asked for, and $($named.Path) holds nothing but a '$script:InstallerFootprintSubKeyName' subkey - which this product's own installer writes under every supported major, on every install. Outlook does not read that key, so an ImportPRF written into it would be a silent no-op. Name the major Outlook really uses, or drop -OfficeVersion."
            }
        }

        return [pscustomobject]@{ Chosen = $named; RealHives = $real; Problem = $null }
    }

    foreach ($version in $script:SupportedOfficeVersions) {
        foreach ($candidate in $real) {
            if ($candidate.Version -eq $version) {
                return [pscustomobject]@{ Chosen = $candidate; RealHives = $real; Problem = $null }
            }
        }
    }

    $unsupported = @()
    foreach ($candidate in $real) { $unsupported += $candidate.Version }

    $problem = @"
REFUSING: found no supported Outlook hive with any content under $script:OfficeRootKeyPath.

Probed in this order: $($script:SupportedOfficeVersions -join ', ') - the order the shipped product
probes in (OfficeVersions.Supported).

Every one of them was either absent or held nothing but a '$script:InstallerFootprintSubKeyName'
subkey, which this product writes under 15.0, 16.0 and 17.0 alike on every install - so a
present-but-empty key proves nothing. Either classic Outlook is not installed, or it has never
been started.
"@

    if ($unsupported.Count -gt 0) {
        $problem += @"


There IS a real Outlook hive at: $($unsupported -join ', '). That major is not one this product
supports, so nothing here will choose it for you. If it is the one Outlook really uses, say so:
    -OfficeVersion $($unsupported[0])
"@
    }

    return [pscustomobject]@{ Chosen = $null; RealHives = $real; Problem = $problem }
}

<#
    Whether a string can survive being written into the .prf. Returns the refusal text, or $null.

    THE .prf IS WRITTEN AS ASCII, and that is not an arbitrary choice: every published sample is
    plain 8-bit text, the measured-working tier-profile.prf is written the same way, and a UTF-8
    BOM would be read by Outlook as part of the first section header. So a character outside
    printable ASCII does not fail - it is written as '?' and the value silently becomes something
    nobody asked for. A control character is worse: a newline inside a value ends the line and
    turns the rest into a key Outlook does not recognise, which it ignores without a word.
#>
function Test-PrfValueSafe {
    param([string] $Value, [string] $What)

    if ($null -eq $Value) { return $null }
    foreach ($character in $Value.ToCharArray()) {
        if ([int] $character -gt 126 -or [int] $character -lt 32) {
            return "REFUSING: the $What contains a character outside printable ASCII. The .prf is written as ASCII with CRLF line endings - which is what every published sample is, and what the measured-working tier-profile.prf is - so that character would be written as '?' or would end the line, and the value would silently become something nobody asked for."
        }
    }
    return $null
}

<#
    Parses -Store into ordered { DisplayName; Path } pairs, or refuses.

    Every refusal here is a failure that would otherwise be SILENT. A .prf is an INI file Outlook
    parses with no diagnostics available to us: a value it dislikes is a property written nowhere,
    and the first sign of it is a store that is not there or is there under a name nothing can
    find. So the argument grammar is checked hard, before anything is written.

    Returns Stores (the parsed pairs) and Problem ($null when every spec was accepted).
#>
function Resolve-StoreSpec {
    param([string[]] $Spec)

    $stores = @()
    $specs = @()
    if ($null -ne $Spec) { $specs = @($Spec) }

    foreach ($one in $specs) {
        $text = ''
        if ($null -ne $one) { $text = [string] $one }

        $split = $text.IndexOf('=')
        if ($split -lt 1 -or $split -eq ($text.Length - 1)) {
            return [pscustomobject]@{
                Stores  = @()
                Problem = "REFUSING: -Store wants 'DisplayName=C:\path\to.pst'; got '$text'. The display name comes first because that is the half the tests match on."
            }
        }

        $displayName = $text.Substring(0, $split).Trim()
        $path = $text.Substring($split + 1).Trim()

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: '$text' has an empty display name. The display name is what the tests look the store up by; there is no useful default for it." }
        }

        if ($displayName -match '[\\/]') {
            # Rename-OutlookStore.ps1 refuses the same character for the same reason: a store's
            # display name is a folder name, and '/' is the one character Outlook is consistently
            # reported to reject in one.
            return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: the display name '$displayName' contains a slash. A store's display name is its root folder's name, and '/' is the one character Outlook is consistently reported to reject in a folder name." }
        }

        foreach ($pair in @(@{ Text = $displayName; What = "display name '$displayName'" }, @{ Text = $path; What = "path '$path'" })) {
            $unsafe = Test-PrfValueSafe -Value $pair.Text -What $pair.What
            if ($null -ne $unsafe) { return [pscustomobject]@{ Stores = @(); Problem = $unsafe } }
        }

        if (-not $path.EndsWith('.pst', [System.StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: '$path' does not end in .pst. PathToPersonalFolders names the data file itself, not the directory it goes in." }
        }

        $full = $null
        try { $full = [System.IO.Path]::GetFullPath($path) } catch { $full = $null }
        if ($null -eq $full) {
            return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: '$path' is not a usable file path." }
        }
        if (-not [System.IO.Path]::IsPathRooted($path)) {
            # The .prf's path is resolved by OUTLOOK, at startup, in whatever directory Outlook
            # happens to be started from - never in this script's. A relative path here is a data
            # file that lands somewhere nobody predicted.
            return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: '$path' is a relative path. Outlook resolves PathToPersonalFolders at startup in ITS working directory, not this script's, so a relative path puts the store somewhere nobody chose. Give an absolute path - '$full' is what it would be from here." }
        }

        foreach ($already in $stores) {
            if ($already.DisplayName -eq $displayName) {
                return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: '$displayName' was asked for twice. Two stores sharing a display name is exactly the state the tests cannot resolve - expectedStoreDisplayNames censuses BY NAME - so it is refused here rather than discovered later." }
            }
            if ($already.Path -ieq $full) {
                return [pscustomobject]@{ Stores = @(); Problem = "REFUSING: '$full' was asked for twice. The same .pst cannot be two stores in one profile." }
            }
        }

        $stores += [pscustomobject]@{ DisplayName = $displayName; Path = $full }
    }

    return [pscustomobject]@{ Stores = $stores; Problem = $null }
}

<#
    Sections 6 and 7 of the .prf: the property mappings. SHIPPED VERBATIM from the OCT-generated
    output that tier-profile.prf carries, which is the file Outlook 16.x measurably processed.

    Microsoft: "You typically do not modify existing entries in Sections 6 and 7." The reason is
    sharper than a style note - a property with no mapping in section 6 or 7 is written NOWHERE,
    silently. So the blocks below are kept whole, including the ones this profile never uses.
#>
function Get-PrfMappingSection {
    return @'
;***************************************************************
; Section 6 - Mapping for profile properties.  DO NOT MODIFY.
;***************************************************************

[Microsoft Exchange Server]
ServiceName=MSEMS
MDBGUID=5494A1C0297F101BA58708002B2A2517
MailboxName=PT_STRING8,0x6607
HomeServer=PT_STRING8,0x6608
OfflineAddressBookPath=PT_STRING8,0x660E
OfflineFolderPath=PT_STRING8,0x6610

[Exchange Global Section]
SectionGUID=13dbb0c8aa05101a9bb000aa002fc45a
MailboxName=PT_STRING8,0x6607
HomeServer=PT_STRING8,0x6608
RPCoverHTTPflags=PT_LONG,0x6623
RPCProxyServer=PT_UNICODE,0x6622
RPCProxyPrincipalName=PT_UNICODE,0x6625
RPCProxyAuthScheme=PT_LONG,0x6627
CachedExchangeConfigFlags=PT_LONG,0x6629

[Personal Folders]
ServiceName=MSPST MS
Name=PT_STRING8,0x3001
PathToPersonalFolders=PT_STRING8,0x6700
RememberPassword=PT_BOOLEAN,0x6701
EncryptionType=PT_LONG,0x6702
Password=PT_STRING8,0x6703

[Unicode Personal Folders]
ServiceName=MSUPST MS
Name=PT_UNICODE,0x3001
PathToPersonalFolders=PT_STRING8,0x6700
RememberPassword=PT_BOOLEAN,0x6701
EncryptionType=PT_LONG,0x6702
Password=PT_STRING8,0x6703

[Outlook Address Book]
ServiceName=CONTAB

[LDAP Directory]
ServiceName=EMABLT
ServerName=PT_STRING8,0x6600
UserName=PT_STRING8,0x6602
UseSSL=PT_BOOLEAN,0x6613
UseSPA=PT_BOOLEAN,0x6615
DisableVLV=PT_LONG,0x6616
DisplayName=PT_STRING8,0x3001
ConnectionPort=PT_STRING8,0x6601
SearchTimeout=PT_STRING8,0x6607
MaxEntriesReturned=PT_STRING8,0x6608
SearchBase=PT_STRING8,0x6603

[Microsoft Outlook Client]
SectionGUID=0a0d020000000000c000000000000046
FormDirectoryPage=PT_STRING8,0x0270
WebServicesLocation=PT_STRING8,0x0271
ComposeWithWebServices=PT_BOOLEAN,0x0272
PromptWhenUsingWebServices=PT_BOOLEAN,0x0273
OpenWithWebServices=PT_BOOLEAN,0x0274
CachedExchangeMode=PT_LONG,0x041f
CachedExchangeSlowDetect=PT_BOOLEAN,0x0420

[Personal Address Book]
ServiceName=MSPST AB
NameOfPAB=PT_STRING8,0x001e3001
Path=PT_STRING8,0x001e6600
ShowNamesBy=PT_LONG,0x00036601

; ************************************************************************
; Section 7 - Mapping for internet account properties.  DO NOT MODIFY.
; ************************************************************************

[I_Mail]
AccountType=POP3
;--- POP3 Account Settings ---
AccountName=PT_UNICODE,0x0002
DisplayName=PT_UNICODE,0x000B
EmailAddress=PT_UNICODE,0x000C
;--- POP3 Account Settings ---
POP3Server=PT_UNICODE,0x0100
POP3UserName=PT_UNICODE,0x0101
POP3UseSPA=PT_LONG,0x0108
Organization=PT_UNICODE,0x0107
ReplyEmailAddress=PT_UNICODE,0x0103
POP3Port=PT_LONG,0x0104
POP3UseSSL=PT_LONG,0x0105
; --- SMTP Account Settings ---
SMTPServer=PT_UNICODE,0x0200
SMTPUseAuth=PT_LONG,0x0203
SMTPAuthMethod=PT_LONG,0x0208
SMTPUserName=PT_UNICODE,0x0204
SMTPUseSPA=PT_LONG,0x0207
ConnectionType=PT_LONG,0x000F
ConnectionOID=PT_UNICODE,0x0010
SMTPPort=PT_LONG,0x0201
SMTPUseSSL=PT_LONG,0x0202
ServerTimeOut=PT_LONG,0x0209
LeaveOnServer=PT_LONG,0x1000

[IMAP_I_Mail]
AccountType=IMAP
;--- IMAP Account Settings ---
AccountName=PT_UNICODE,0x0002
DisplayName=PT_UNICODE,0x000B
EmailAddress=PT_UNICODE,0x000C
;--- IMAP Account Settings ---
IMAPServer=PT_UNICODE,0x0100
IMAPUserName=PT_UNICODE,0x0101
IMAPUseSPA=PT_LONG,0x0108
Organization=PT_UNICODE,0x0107
ReplyEmailAddress=PT_UNICODE,0x0103
IMAPPort=PT_LONG,0x0104
IMAPUseSSL=PT_LONG,0x0105
; --- SMTP Account Settings ---
SMTPServer=PT_UNICODE,0x0200
SMTPUseAuth=PT_LONG,0x0203
SMTPAuthMethod=PT_LONG,0x0208
SMTPUserName=PT_UNICODE,0x0204
SMTPUseSPA=PT_LONG,0x0207
ConnectionType=PT_LONG,0x000F
ConnectionOID=PT_UNICODE,0x0010
SMTPPort=PT_LONG,0x0201
SMTPUseSSL=PT_LONG,0x0202
ServerTimeOut=PT_LONG,0x0209
CheckNewImap=PT_LONG,0x1100
RootFolder=PT_UNICODE,0x1101
'@
}

<#
    Builds the whole .prf, as one string with CRLF line endings.

    IT IS GENERATED RATHER THAN TOKEN-SUBSTITUTED INTO A COMMITTED TEMPLATE, and that is a
    deliberate difference from New-TierProfile.ps1. The number of [ServiceN] blocks varies with
    -Store, and a template cannot carry a variable number of sections; the choice is between
    generating the file and post-processing a template, and generating it is the one a pure
    function can test every branch of.

    THE THREE THINGS THAT MAKE IT WORK, each copied from the measured file rather than reasoned
    out:
      * [ServiceN] Name=<display name> maps through section 6's [Unicode Personal Folders] to
        PT_UNICODE,0x3001 - PR_DISPLAY_NAME_W. That is the property the object model refuses to
        let anything set afterwards, and setting it here is the entire reason this route was
        chosen over Namespace.AddStoreEx.
      * UniqueService=No is what permits the same service block to appear more than once.
      * EncryptionType=0x80000000 is verbatim from tier-profile.prf's [Service1]. It is not
        reasoned about here; it is what was on the guest when a store came out correctly named.

    DefaultStore names a SERVICE, not a path, and is written only when there IS a store. On the
    measured run it is what stopped Outlook minting a data file of its own.
#>
function New-ProfilePrfText {
    param(
        [Parameter(Mandatory = $true)] [string] $ProfileName,
        [object[]] $Stores,
        [bool] $MakeDefault = $false
    )

    $list = @()
    if ($null -ne $Stores) { $list = @($Stores) }

    $defaultProfile = 'No'
    if ($MakeDefault) { $defaultProfile = 'Yes' }

    $lines = @()
    $lines += ';============================================================================================='
    $lines += '; GENERATED by Testbed/guest/New-OutlookProfile.ps1. Do not hand-edit: the next -Execute'
    $lines += '; overwrites it, and a hand edit here is a change nothing in the repository records.'
    $lines += ';'
    $lines += '; It creates ONE profile and its PST stores, each under an exact display name, and NO mail'
    $lines += '; account: the corpus generator refuses any profile that has one, with no override flag.'
    $lines += ';'
    $lines += '; Sections 6 and 7 are shipped VERBATIM from OCT-generated output, as tier-profile.prf'
    $lines += '; carries them - the file Outlook 16.x measurably processed on 2026-09-15. A property with'
    $lines += '; no mapping there is written NOWHERE, silently, so nothing in them is tidied.'
    $lines += ';============================================================================================='
    $lines += ''
    $lines += '; **************************************************************'
    $lines += '; Section 1 - Profile Defaults'
    $lines += '; **************************************************************'
    $lines += ''
    $lines += '[General]'
    $lines += 'Custom=1'
    $lines += "ProfileName=$ProfileName"
    $lines += "DefaultProfile=$defaultProfile"
    $lines += 'OverwriteProfile=Yes'
    $lines += 'BackupProfile=No'
    $lines += 'ModifyDefaultProfileIfPresent=FALSE'
    if ($list.Count -gt 0) { $lines += 'DefaultStore=Service1' }
    $lines += ''
    $lines += '; **************************************************************'
    $lines += '; Section 2 - Services in Profile'
    $lines += '; **************************************************************'
    $lines += ''
    $lines += '[Service List]'

    $index = 0
    foreach ($store in $list) {
        $index++
        $lines += "Service$index=$script:PstServiceBlockName"
    }
    $index++
    $addressBookIndex = $index
    $lines += "Service$addressBookIndex=$script:AddressBookBlockName"

    $lines += ''
    $lines += ';***************************************************************'
    $lines += '; Section 3 - List of internet accounts'
    $lines += ';***************************************************************'
    $lines += ''
    $lines += '; DELIBERATELY EMPTY. No AccountN= line means no account, and no account is the whole'
    $lines += '; point of this profile - corpus-build refuses any profile whose Accounts.Count is not'
    $lines += '; zero, and there is no flag that talks it past that.'
    $lines += '[Internet Account List]'
    $lines += ''
    $lines += ';***************************************************************'
    $lines += '; Section 4 - Default values for each service.'
    $lines += ';***************************************************************'

    $index = 0
    foreach ($store in $list) {
        $index++
        $lines += ''
        $lines += "[Service$index]"
        $lines += 'UniqueService=No'
        $lines += "Name=$($store.DisplayName)"
        $lines += "PathToPersonalFolders=$($store.Path)"
        $lines += 'EncryptionType=0x80000000'
    }
    $lines += ''
    $lines += "[Service$addressBookIndex]"
    $lines += ''
    $lines += ';***************************************************************'
    $lines += '; Section 5 - Values for each internet account.'
    $lines += ';***************************************************************'
    $lines += ''
    $lines += '; Nothing here, by design. See section 3.'
    $lines += ''
    $lines += (Get-PrfMappingSection)

    # ASCII, CRLF, no BOM. Every published sample is plain 8-bit text and a UTF-8 BOM would prefix
    # the first line, which Outlook would read as part of the first section header.
    $text = ($lines -join "`n")
    $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    return $text
}

<#
    The verify verdict, decided against what was read back rather than against what was asked for.

    Returns Checks - each { Level = 'pass'|'fail'|'warn'; What; Detail } - plus the two counts.
    Every branch is here so -SelfTest can walk all of them; the machine half only gathers inputs
    and prints the result.

    $AccountSubKeyCount is $null when the read FAILED, which is deliberately not the same as zero:
    a preflight that refuses a legitimate build on a registry read it could not make would be
    worse than no preflight, so an unreadable count is a warning and the COM count is what decides.
#>
function Test-ProfileOutcome {
    param(
        [Parameter(Mandatory = $true)] [string] $ProfileName,
        [string[]] $ProfileNames,
        $AccountSubKeyCount,
        [object[]] $Stores,
        [string[]] $PstPathsPresent,
        [string[]] $StrayPstNames,
        [string] $ImportPrfValue
    )

    $names = @()
    if ($null -ne $ProfileNames) { $names = @($ProfileNames) }
    $list = @()
    if ($null -ne $Stores) { $list = @($Stores) }
    $present = @()
    if ($null -ne $PstPathsPresent) { $present = @($PstPathsPresent) }
    $stray = @()
    if ($null -ne $StrayPstNames) { $stray = @($StrayPstNames) }

    $checks = @()

    if ($names -contains $ProfileName) {
        $checks += [pscustomobject]@{ Level = 'pass'; What = 'the profile exists'; Detail = $ProfileName }
    }
    else {
        $presentText = '<none>'
        if ($names.Count -gt 0) { $presentText = ($names -join ', ') }
        $checks += [pscustomobject]@{
            Level  = 'fail'
            What   = 'the profile exists'
            Detail = "No profile named '$ProfileName'. Profiles present: $presentText. Either Outlook has not been started since -Execute ran, or it did not process the .prf - check that ImportPRF is still set and that FirstRun/First-Run are absent under the Setup key."
        }
    }

    $backups = @($names | Where-Object { $_ -like 'Backup Of*' })
    if ($backups.Count -eq 0) {
        $checks += [pscustomobject]@{ Level = 'pass'; What = 'no backup profile was created'; Detail = 'BackupProfile=No held' }
    }
    else {
        $checks += [pscustomobject]@{
            Level  = 'fail'
            What   = 'no backup profile was created'
            Detail = "Found: $($backups -join ', '). BackupProfile=No was not honoured - Microsoft spells it False once, in passing, and the community spells it No; try the other spelling. Delete these before re-running: a repeat import that keeps making them makes the guest non-reproducible."
        }
    }

    if ($null -eq $AccountSubKeyCount) {
        $checks += [pscustomobject]@{
            Level  = 'warn'
            What   = 'the profile has no mail accounts'
            Detail = "Could not read the account-manager key ($script:AccountManagerSubKeyName) under this profile, so this says nothing either way. The count that DECIDES is the tool's own COM read - it prints 'profile accounts: N' when it vets the store - and -WithOutlook makes the same read."
        }
    }
    elseif ([int] $AccountSubKeyCount -eq 0) {
        $checks += [pscustomobject]@{ Level = 'pass'; What = 'the profile has no mail accounts'; Detail = "0 subkeys under $script:AccountManagerSubKeyName" }
    }
    else {
        $checks += [pscustomobject]@{
            Level  = 'fail'
            What   = 'the profile has no mail accounts'
            Detail = "$AccountSubKeyCount subkey(s) under $script:AccountManagerSubKeyName. corpus-build will REFUSE this profile and there is no override. The .prf named no account, so something else put one here - or the empty [Internet Account List] is not enough on this build, which is the inference this script's banner flags. The measured fallback is 'outlook.exe /PIM <name>', which produces accounts=0 on this Office build."
        }
    }

    if ($list.Count -eq 0) {
        $checks += [pscustomobject]@{
            Level  = 'warn'
            What   = 'the profile has at least one store'
            Detail = 'No -Store was named, so nothing about the stores was checked. If the profile really was built store-less that is legal and almost never what you want - Outlook is liable to mint a data file of its own at first use, under a name nobody chose. If it was not, you left -Store off this command: pass the SAME -Store arguments to -Verify as to -Execute, or this check is vacuous.'
        }
    }
    foreach ($store in $list) {
        if ($present -contains $store.Path) {
            $checks += [pscustomobject]@{ Level = 'pass'; What = "the PST for '$($store.DisplayName)' exists"; Detail = $store.Path }
        }
        else {
            $checks += [pscustomobject]@{
                Level  = 'fail'
                What   = "the PST for '$($store.DisplayName)' exists"
                Detail = "$($store.Path) is not there. Outlook creates the file when the service is configured, so an absent one means the [ServiceN] block did not take. Read the .prf that was written before assuming the route is broken."
            }
        }
    }

    if ($stray.Count -eq 0) {
        $checks += [pscustomobject]@{ Level = 'pass'; What = 'Outlook minted no PST of its own'; Detail = 'Documents\Outlook Files is empty of .pst' }
    }
    else {
        $checks += [pscustomobject]@{
            Level  = 'fail'
            What   = 'Outlook minted no PST of its own'
            Detail = "Found in Documents\Outlook Files: $($stray -join ', '). DefaultStore=Service1 did not take. A store nobody named is a store the tests cannot census, and expectedStoreDisplayNames will refuse the tier over it."
        }
    }

    if ([string]::IsNullOrEmpty($ImportPrfValue)) {
        $checks += [pscustomobject]@{ Level = 'pass'; What = 'ImportPRF is no longer set'; Detail = 'the next Outlook start will not re-import' }
    }
    else {
        $checks += [pscustomobject]@{
            Level  = 'warn'
            What   = 'ImportPRF is no longer set'
            Detail = "It still reads '$ImportPrfValue'. Whether Outlook clears it after processing has NOT been established on this build. If it does not, every later Outlook start re-imports - and this .prf carries OverwriteProfile=Yes, so it would REBUILD the profile and detach a store that had since been filled. Clear it: New-OutlookProfile.ps1 -ClearImportPrf -Execute"
        }
    }

    $failures = @($checks | Where-Object { $_.Level -eq 'fail' })
    $warnings = @($checks | Where-Object { $_.Level -eq 'warn' })
    return [pscustomobject]@{
        Checks       = $checks
        FailureCount = $failures.Count
        WarningCount = $warnings.Count
    }
}

# =============================================================================================
# SELF-TEST. Pure: reads no registry, writes no file, starts no process, needs no guest.
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

    function New-HiveCandidate {
        param([string] $Version, [string[]] $ValueNames = @(), [string[]] $SubKeyNames = @())

        return [pscustomobject]@{
            Version     = $Version
            Path        = "$script:OfficeRootKeyPath\$Version\Outlook"
            ValueNames  = $ValueNames
            SubKeyNames = $SubKeyNames
        }
    }

    function Get-CheckLevel {
        param($Outcome, [string] $What)

        foreach ($check in $Outcome.Checks) {
            if ($check.What -eq $What) { return $check.Level }
        }
        return '<no such check>'
    }

    function Get-CheckDetail {
        param($Outcome, [string] $What)

        foreach ($check in $Outcome.Checks) {
            if ($check.What -eq $What) { return $check.Detail }
        }
        return '<no such check>'
    }

    Write-Host 'New-OutlookProfile self-test. Nothing is read, nothing is written, nothing is started.'
    Write-Host ''
    Write-Host '== the hive shape rule (OfficeVersions.IsOutlookHive) =='

    Test-Case 'the installer footprint alone is not a hive' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('Resiliency'))
    Test-Case 'an empty key is not a hive' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @())
    Test-Case 'nulls are not a hive, and do not throw' $false (Test-IsOutlookHive -ValueNames $null -SubKeyNames $null)
    Test-Case 'one value is enough' $true (Test-IsOutlookHive -ValueNames @('DefaultProfile') -SubKeyNames @('Resiliency'))
    Test-Case 'one non-footprint subkey is enough' $true (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('Profiles', 'Resiliency'))

    Write-Host ''
    Write-Host '== which hive gets chosen =='

    $real16 = New-HiveCandidate -Version '16.0' -ValueNames @('DefaultProfile') -SubKeyNames @('Profiles', 'Resiliency')
    $real17 = New-HiveCandidate -Version '17.0' -ValueNames @('DefaultProfile') -SubKeyNames @('Profiles')
    $real15 = New-HiveCandidate -Version '15.0' -ValueNames @() -SubKeyNames @('Profiles')
    $real14 = New-HiveCandidate -Version '14.0' -ValueNames @('DefaultProfile') -SubKeyNames @('Profiles')
    $shell15 = New-HiveCandidate -Version '15.0' -SubKeyNames @('Resiliency')
    $shell16 = New-HiveCandidate -Version '16.0' -SubKeyNames @('Resiliency')
    $shell17 = New-HiveCandidate -Version '17.0' -SubKeyNames @('Resiliency')

    $pick = Select-OutlookHive -Candidates @($shell15, $real16, $shell17)
    Test-Case '16.0 wins over the installer footprint keys' '16.0' $pick.Chosen.Version
    Test-Case 'and only the real one is counted as real' 1 $pick.RealHives.Count

    $pick = Select-OutlookHive -Candidates @($real16, $real17)
    Test-Case '16.0 wins over a real 17.0 (probe order, not highest number)' '16.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell16, $real17)
    Test-Case '17.0 wins when 16.0 is only the footprint' '17.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell16, $shell17, $real15)
    Test-Case '15.0 is chosen when it is the only real one' '15.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell15, $shell16, $shell17)
    Test-Case 'all-footprint refuses' '<null>' $pick.Chosen
    Test-Case 'and names the probe order it used' $true ($pick.Problem -like '*16.0, 17.0, 15.0*')

    $pick = Select-OutlookHive -Candidates $null
    Test-Case 'a null candidate list refuses, and does not throw' '<null>' $pick.Chosen

    $pick = Select-OutlookHive -Candidates @($shell16, $real14)
    Test-Case 'an unsupported major is never chosen for you' '<null>' $pick.Chosen
    Test-Case 'but it is named, with the flag that would use it' $true ($pick.Problem -like '*-OfficeVersion 14.0*')

    $pick = Select-OutlookHive -Candidates @($real16, $real15) -RequestedVersion '15.0'
    Test-Case '-OfficeVersion overrides the probe order' '15.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($real16) -RequestedVersion '15.0'
    Test-Case '-OfficeVersion naming an absent key refuses' '<null>' $pick.Chosen
    Test-Case 'and lists the version keys that are present' $true ($pick.Problem -like '*Version keys present: 16.0*')

    $pick = Select-OutlookHive -Candidates @($real16, $shell15) -RequestedVersion '15.0'
    Test-Case '-OfficeVersion does NOT override the shape rule' '<null>' $pick.Chosen
    Test-Case 'and says why that key is not real' $true ($pick.Problem -like '*silent no-op*')

    Write-Host ''
    Write-Host '== what may be written into a .prf at all =='

    Test-Case 'an ordinary value is fine' '<null>' (Test-PrfValueSafe -Value 'CorpusProfile' -What 'x')
    Test-Case 'an @ is fine' '<null>' (Test-PrfValueSafe -Value 'test@vm.invalid' -What 'x')
    Test-Case 'a path with a space is fine here - it is the PRF PATH that may not have one' '<null>' (Test-PrfValueSafe -Value 'C:\Some Folder\a.pst' -What 'x')
    Test-Case 'a null is fine and does not throw' '<null>' (Test-PrfValueSafe -Value $null -What 'x')
    Test-Case 'a non-ASCII character refuses' $true ((Test-PrfValueSafe -Value ([string]([char]0x00E9)) -What 'profile name') -like '*outside printable ASCII*')
    Test-Case 'and names what it was checking' $true ((Test-PrfValueSafe -Value ([string]([char]0x00E9)) -What 'profile name') -like '*the profile name contains*')
    Test-Case 'an embedded newline refuses - it would end the line' $true ((Test-PrfValueSafe -Value "a`nb" -What 'x') -like '*outside printable ASCII*')
    Test-Case 'a tab refuses too' $true ((Test-PrfValueSafe -Value "a`tb" -What 'x') -like '*outside printable ASCII*')

    Write-Host ''
    Write-Host '== the -Store grammar, and every one of its refusals =='

    $parsed = Resolve-StoreSpec -Spec @('Corpus A=C:\OutlookAI-Q5\pst\corpus-a.pst')
    Test-Case 'a well-formed spec parses' '<null>' $parsed.Problem
    Test-Case 'and keeps the display name exactly' 'Corpus A' $parsed.Stores[0].DisplayName
    Test-Case 'and the full path' 'C:\OutlookAI-Q5\pst\corpus-a.pst' $parsed.Stores[0].Path

    $parsed = Resolve-StoreSpec -Spec @('  Corpus A  =  C:\p\a.pst  ')
    Test-Case 'surrounding whitespace is trimmed off both halves' 'Corpus A' $parsed.Stores[0].DisplayName
    Test-Case 'on the path too' 'C:\p\a.pst' $parsed.Stores[0].Path

    $parsed = Resolve-StoreSpec -Spec @('test@vm.invalid=C:\p\hub.pst')
    Test-Case 'an @ in a display name is ACCEPTED - measured, twice' '<null>' $parsed.Problem
    Test-Case 'and survives intact' 'test@vm.invalid' $parsed.Stores[0].DisplayName

    $parsed = Resolve-StoreSpec -Spec @('C:\p\a.pst')
    Test-Case 'a spec with no = refuses' $true ($parsed.Problem -like "*wants 'DisplayName=*")

    $parsed = Resolve-StoreSpec -Spec @('=C:\p\a.pst')
    Test-Case 'an = in position 0 refuses' $true ($parsed.Problem -like "*wants 'DisplayName=*")

    $parsed = Resolve-StoreSpec -Spec @('Corpus A=')
    Test-Case 'an empty path refuses' $true ($parsed.Problem -like "*wants 'DisplayName=*")

    $parsed = Resolve-StoreSpec -Spec @('   =C:\p\a.pst')
    Test-Case 'a whitespace-only display name refuses' $true ($parsed.Problem -like '*empty display name*')

    $parsed = Resolve-StoreSpec -Spec @('Corpus/A=C:\p\a.pst')
    Test-Case 'a slash in the display name refuses' $true ($parsed.Problem -like '*contains a slash*')

    $parsed = Resolve-StoreSpec -Spec @('Corpus\A=C:\p\a.pst')
    Test-Case 'a backslash in the display name refuses too' $true ($parsed.Problem -like '*contains a slash*')

    $parsed = Resolve-StoreSpec -Spec @([string]([char]0x00E9) + "quipe=C:\p\a.pst")
    Test-Case 'a non-ASCII display name refuses' $true ($parsed.Problem -like '*outside printable ASCII*')

    $parsed = Resolve-StoreSpec -Spec @('Corpus A=C:\p\corpus-a')
    Test-Case 'a path that is not a .pst refuses' $true ($parsed.Problem -like '*does not end in .pst*')

    $parsed = Resolve-StoreSpec -Spec @('Corpus A=pst\corpus-a.pst')
    Test-Case 'a relative path refuses' $true ($parsed.Problem -like '*relative path*')
    Test-Case 'and says whose working directory decides' $true ($parsed.Problem -like '*ITS working directory*')

    $parsed = Resolve-StoreSpec -Spec @('Corpus A=C:\p\a.pst', 'Corpus A=C:\p\b.pst')
    Test-Case 'a repeated display name refuses' $true ($parsed.Problem -like '*asked for twice*')

    $parsed = Resolve-StoreSpec -Spec @('Corpus A=C:\p\a.pst', 'Corpus B=C:\P\A.PST')
    Test-Case 'the same file twice refuses, case-insensitively' $true ($parsed.Problem -like '*cannot be two stores*')

    $parsed = Resolve-StoreSpec -Spec @()
    Test-Case 'no -Store at all is legal and yields nothing' 0 $parsed.Stores.Count
    $parsed = Resolve-StoreSpec -Spec $null
    Test-Case 'a null -Store is the same, and does not throw' 0 $parsed.Stores.Count

    Write-Host ''
    Write-Host '== the .prf that gets written =='

    # EVERY ASSERTION BELOW USES String.Contains, NOT -like, AND THAT IS NOT A STYLE CHOICE.
    # -like treats '[...]' as a CHARACTER CLASS, so `$prf -like '*[Account1]*'` asks "does this
    # text contain any one of A,c,o,u,n,t,1" - which is true of almost any text, and the assertion
    # then passes for the wrong reason. Written with -like first, four of these lied. It is the
    # same wildcard trap CLAUDE.md's mailbox-safety rule 2 is about, and every section header in a
    # .prf is in square brackets, so it would have hit every interesting assertion here.
    # Contains is ordinal and case-SENSITIVE, which is what a .prf needs anyway.
    $oneStore = @([pscustomobject]@{ DisplayName = 'Corpus A'; Path = 'C:\OutlookAI-Q5\pst\corpus-a.pst' })
    $prf = New-ProfilePrfText -ProfileName 'CorpusProfile' -Stores $oneStore -MakeDefault $true

    Test-Case 'the profile name is written' $true ($prf.Contains("`r`nProfileName=CorpusProfile`r`n"))
    Test-Case '-MakeDefault writes DefaultProfile=Yes' $true ($prf.Contains("`r`nDefaultProfile=Yes`r`n"))
    Test-Case 'a repeat import overwrites rather than backing up' $true ($prf.Contains("`r`nOverwriteProfile=Yes`r`n"))
    Test-Case 'and makes no Backup Of profile' $true ($prf.Contains("`r`nBackupProfile=No`r`n"))
    Test-Case 'the one store becomes the profile default store' $true ($prf.Contains("`r`nDefaultStore=Service1`r`n"))
    Test-Case 'Service1 is the Unicode PST service' $true ($prf.Contains("`r`nService1=Unicode Personal Folders`r`n"))
    Test-Case 'the address book follows it' $true ($prf.Contains("`r`nService2=Outlook Address Book`r`n"))
    Test-Case 'the display name lands on the service block' $true ($prf.Contains("`r`n[Service1]`r`nUniqueService=No`r`nName=Corpus A`r`n"))
    Test-Case 'and so does the path' $true ($prf.Contains("`r`nPathToPersonalFolders=C:\OutlookAI-Q5\pst\corpus-a.pst`r`n"))
    Test-Case 'the encryption value is the one from the measured file' $true ($prf.Contains('EncryptionType=0x80000000'))
    Test-Case 'UniqueService=No is what lets the block repeat' $true ($prf.Contains('UniqueService=No'))
    Test-Case 'the internet account list is present' $true ($prf.Contains("`r`n[Internet Account List]`r`n"))
    Test-Case 'and holds no account' $false ($prf.Contains("`r`nAccount1="))
    Test-Case 'so there is no [Account1] block either' $false ($prf.Contains('[Account1]'))
    Test-Case 'section 6 maps Name to PR_DISPLAY_NAME_W' $true ($prf.Contains('Name=PT_UNICODE,0x3001'))
    Test-Case 'section 6 names the Unicode PST provider' $true ($prf.Contains('ServiceName=MSUPST MS'))
    Test-Case 'section 7 is shipped whole even though nothing uses it' $true ($prf.Contains('[IMAP_I_Mail]'))
    Test-Case 'every line ends CRLF' $false ($prf -match "[^`r]`n")
    Test-Case 'and there is no BOM or leading blank' ';' $prf.Substring(0, 1)

    $prf = New-ProfilePrfText -ProfileName 'CorpusProfile' -Stores $oneStore -MakeDefault $false
    Test-Case 'without -MakeDefault it is DefaultProfile=No' $true ($prf.Contains("`r`nDefaultProfile=No`r`n"))

    $threeStores = @(
        [pscustomobject]@{ DisplayName = 'Corpus A'; Path = 'C:\p\a.pst' },
        [pscustomobject]@{ DisplayName = 'test@vm.invalid'; Path = 'C:\p\hub.pst' },
        [pscustomobject]@{ DisplayName = 'Bystander'; Path = 'C:\p\by.pst' })
    $prf = New-ProfilePrfText -ProfileName 'CorpusProfile' -Stores $threeStores

    Test-Case 'three stores give three PST services' $true ($prf.Contains("`r`nService3=Unicode Personal Folders`r`n"))
    Test-Case 'and the address book is numbered after them' $true ($prf.Contains("`r`nService4=Outlook Address Book`r`n"))
    Test-Case 'there is no Service5' $false ($prf.Contains('Service5='))
    Test-Case 'the second store keeps its @' $true ($prf.Contains("`r`nName=test@vm.invalid`r`n"))
    Test-Case 'each store gets its own block' $true ($prf.Contains("`r`n[Service3]`r`nUniqueService=No`r`nName=Bystander`r`n"))
    Test-Case 'the address book block is present and empty' $true ($prf.Contains("`r`n[Service4]`r`n`r`n"))

    $prf = New-ProfilePrfText -ProfileName 'BareProfile' -Stores @()
    Test-Case 'no stores means no DefaultStore line' $false ($prf.Contains('DefaultStore='))
    Test-Case 'and the address book is Service1' $true ($prf.Contains("`r`nService1=Outlook Address Book`r`n"))
    Test-Case 'and no PST service at all' $false ($prf.Contains('=Unicode Personal Folders'))

    Write-Host ''
    Write-Host '== the verify verdict =='

    $good = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('Outlook', 'CorpusProfile') -AccountSubKeyCount 0 -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'a clean import has no failures' 0 $good.FailureCount
    Test-Case 'and no warnings either' 0 $good.WarningCount
    Test-Case 'the profile check passes' 'pass' (Get-CheckLevel $good 'the profile exists')
    Test-Case 'the account check passes' 'pass' (Get-CheckLevel $good 'the profile has no mail accounts')
    Test-Case "the store's PST check passes" 'pass' (Get-CheckLevel $good "the PST for 'Corpus A' exists")

    $bad = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('Outlook') -AccountSubKeyCount 0 -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'a missing profile FAILS' 'fail' (Get-CheckLevel $bad 'the profile exists')
    Test-Case 'and names what is there instead' $true ((Get-CheckDetail $bad 'the profile exists') -like '*Profiles present: Outlook*')

    $bad = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('CorpusProfile', 'Backup Of CorpusProfile') -AccountSubKeyCount 0 -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'a Backup Of profile FAILS' 'fail' (Get-CheckLevel $bad 'no backup profile was created')
    Test-Case 'and names the other spelling to try' $true ((Get-CheckDetail $bad 'no backup profile was created') -like '*try the other spelling*')

    $bad = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('CorpusProfile') -AccountSubKeyCount 1 -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'an account on the corpus profile FAILS' 'fail' (Get-CheckLevel $bad 'the profile has no mail accounts')
    Test-Case 'and names the measured fallback' $true ((Get-CheckDetail $bad 'the profile has no mail accounts') -like '*/PIM*')

    $unknown = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('CorpusProfile') -AccountSubKeyCount $null -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'an unreadable account count WARNS rather than failing' 'warn' (Get-CheckLevel $unknown 'the profile has no mail accounts')
    Test-Case 'and no failure is recorded for it' 0 $unknown.FailureCount

    $bad = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('CorpusProfile') -AccountSubKeyCount 0 -Stores $oneStore `
        -PstPathsPresent @() -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'a PST that was never created FAILS' 'fail' (Get-CheckLevel $bad "the PST for 'Corpus A' exists")

    $bad = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('CorpusProfile') -AccountSubKeyCount 0 -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @('corpus-a(1).pst') -ImportPrfValue ''
    Test-Case 'a store Outlook minted itself FAILS' 'fail' (Get-CheckLevel $bad 'Outlook minted no PST of its own')
    Test-Case 'and says what it costs' $true ((Get-CheckDetail $bad 'Outlook minted no PST of its own') -like '*expectedStoreDisplayNames*')

    $warned = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames @('CorpusProfile') -AccountSubKeyCount 0 -Stores $oneStore `
        -PstPathsPresent @('C:\OutlookAI-Q5\pst\corpus-a.pst') -StrayPstNames @() -ImportPrfValue 'C:\OutlookAI-Profiles\CorpusProfile.prf'
    Test-Case 'a lingering ImportPRF WARNS' 'warn' (Get-CheckLevel $warned 'ImportPRF is no longer set')
    Test-Case 'and says it would rebuild the profile' $true ((Get-CheckDetail $warned 'ImportPRF is no longer set') -like '*REBUILD the profile*')
    Test-Case 'and gives the command that clears it' $true ((Get-CheckDetail $warned 'ImportPRF is no longer set') -like '*-ClearImportPrf -Execute*')
    Test-Case 'but it is not a failure' 0 $warned.FailureCount

    $bare = Test-ProfileOutcome -ProfileName 'BareProfile' `
        -ProfileNames @('BareProfile') -AccountSubKeyCount 0 -Stores @() `
        -PstPathsPresent @() -StrayPstNames @() -ImportPrfValue ''
    Test-Case 'a store-less profile WARNS about having nowhere to deliver' 'warn' (Get-CheckLevel $bare 'the profile has at least one store')
    Test-Case 'and is not otherwise a failure' 0 $bare.FailureCount

    $nulls = Test-ProfileOutcome -ProfileName 'CorpusProfile' `
        -ProfileNames $null -AccountSubKeyCount 0 -Stores $null `
        -PstPathsPresent $null -StrayPstNames $null -ImportPrfValue $null
    Test-Case 'nulls throughout do not throw' 'fail' (Get-CheckLevel $nulls 'the profile exists')

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. These need a guest, and nothing on this machine can stand in for them:'
    Write-Host '  * reading the Office version keys and the shape of each Outlook subkey out of HKCU'
    Write-Host '  * the ImportPRF write, the FirstRun/First-Run deletes, and their read-backs'
    Write-Host '  * writing the .prf and reading it back byte-identical'
    Write-Host '  * WHETHER OUTLOOK ACCEPTS A .prf WITH AN EMPTY [Internet Account List] and produces'
    Write-Host '    a profile with zero accounts - the central inference of this rewrite'
    Write-Host '  * WHETHER [Service List] MAY NAME Unicode Personal Folders MORE THAN ONCE'
    Write-Host '  * whether Outlook clears ImportPRF after processing it'
    Write-Host '  * the two guards (Assert-TestbedGuest, Assert-OutlookNotRunning)'
    Write-Host '  * every COM read in -WithOutlook, and what Outlook actually shows'

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

# Dot-sourced for Assert-TestbedGuest, Assert-OutlookNotRunning, Resolve-PstPath and
# Invoke-WithOutlookSession. There is no MAPI left in that file to call - see its banner.
. "$PSScriptRoot\OutlookMapiInterop.ps1"

# Asserted before ANYTHING, including the read-only paths: these scripts create and reconfigure
# Outlook profiles, and on the maintainer's workstation that is a real profile with real delegate
# mailboxes. Refusing early costs nothing and is the whole point of the guard.
Assert-TestbedGuest -ExpectedUser $ExpectedUser

<#
    Every ...\Office\<major>\Outlook key on this machine, as Select-OutlookHive's candidates.
    Reads only: no value is written and no key is created.
#>
function Get-OutlookHiveCandidate {
    if (-not (Test-Path -LiteralPath $script:OfficeRootKeyPath)) { return , @() }

    $candidates = @()
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
    return , $candidates
}

<#
    The profile names: the subkey names under <hive>\Profiles. A subkey IS a profile.
#>
function Get-OutlookProfileName {
    param([string] $ProfilesKeyPath)

    if (-not (Test-Path -LiteralPath $ProfilesKeyPath)) { return , @() }
    $names = @(Get-ChildItem -LiteralPath $ProfilesKeyPath -ErrorAction SilentlyContinue |
            ForEach-Object { $_.PSChildName })
    return , $names
}

<#
    How many accounts a profile's account-manager key holds. $null when the key or the profile is
    not there to read - which is NOT the same as zero, and Test-ProfileOutcome treats it differently.
#>
function Get-ProfileAccountCount {
    param([string] $ProfilesKeyPath, [string] $ProfileName)

    $managerPath = "$ProfilesKeyPath\$ProfileName\$script:AccountManagerSubKeyName"
    if (-not (Test-Path -LiteralPath "$ProfilesKeyPath\$ProfileName")) { return $null }

    # The profile exists and the manager key does not: that IS zero accounts, and it is the shape
    # measured on this project's guest for a profile with a PST store and no internet account.
    if (-not (Test-Path -LiteralPath $managerPath)) { return 0 }

    $children = $null
    try { $children = @(Get-ChildItem -LiteralPath $managerPath -ErrorAction Stop) }
    catch { return $null }
    return $children.Count
}

<#
    Any .pst Outlook minted for itself, which is the documented Outlook 2010+ behaviour this .prf
    is trying to avoid. Names only - the directory is the user's own.
#>
function Get-StrayPstName {
    $strayDir = Join-Path $env:USERPROFILE 'Documents\Outlook Files'
    if (-not (Test-Path -LiteralPath $strayDir)) { return , @() }
    $names = @(Get-ChildItem -LiteralPath $strayDir -Filter '*.pst' -File -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Name })
    return , $names
}

$selection = Select-OutlookHive -Candidates (Get-OutlookHiveCandidate) -RequestedVersion $OfficeVersion
if ($null -eq $selection.Chosen) { throw $selection.Problem }

$root = $selection.Chosen
$profilesKeyPath = "$($root.Path)\Profiles"
$setupKeyPath = "$($root.Path)\Setup"

Write-Host "outlook hive : $($root.Path)   (Office $($root.Version); $($root.ValueNames.Count) value(s), $($root.SubKeyNames.Count) subkey(s))"
if ($selection.RealHives.Count -gt 1) {
    $others = @()
    foreach ($hive in $selection.RealHives) {
        if ($hive.Version -ne $root.Version) { $others += $hive.Version }
    }
    Write-Host "               other real Outlook hive(s) here: $($others -join ', ') - use -OfficeVersion if this one is wrong"
}

$currentImportPrf = (Get-ItemProperty -Path $setupKeyPath -Name 'ImportPRF' -ErrorAction SilentlyContinue).ImportPRF
$existingProfiles = Get-OutlookProfileName -ProfilesKeyPath $profilesKeyPath

# ---------------------------------------------------------------------------------------------
# PREFLIGHT. Read-only. It makes no COM call and no MAPI call on purpose: a preflight that could
# itself start Outlook, or itself hang, would be the fault it is checking for.
# ---------------------------------------------------------------------------------------------
if ($Preflight) {
    Write-Host ''
    Write-Host 'PREFLIGHT - what the .prf import route needs. Nothing is written.'
    Write-Host ''
    Write-Host ("  {0,-22} {1}" -f 'Setup key', $setupKeyPath)
    Write-Host ("  {0,-22} {1}" -f '  exists', (Test-Path -LiteralPath $setupKeyPath))
    Write-Host ("  {0,-22} {1}" -f 'Profiles key', $profilesKeyPath)

    $profilesText = '<none>'
    if ($existingProfiles.Count -gt 0) { $profilesText = ($existingProfiles -join ', ') }
    Write-Host ("  {0,-22} {1}" -f '  profiles', $profilesText)

    $defaultProfile = (Get-ItemProperty -Path $root.Path -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
    Write-Host ("  {0,-22} {1}" -f '  DefaultProfile', $defaultProfile)

    $importText = '<not set>'
    if (-not [string]::IsNullOrEmpty($currentImportPrf)) { $importText = $currentImportPrf }
    Write-Host ("  {0,-22} {1}" -f 'ImportPRF', $importText)

    foreach ($valueName in @('First-Run', 'FirstRun')) {
        $value = (Get-ItemProperty -Path $setupKeyPath -Name $valueName -ErrorAction SilentlyContinue).$valueName
        $state = 'absent - good, ImportPRF will be honoured'
        if ($null -ne $value) { $state = 'PRESENT - ImportPRF is ignored while it exists, with no diagnostic anywhere' }
        Write-Host ("  {0,-22} {1}" -f "  $valueName", $state)
    }

    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    $runningText = 'no - good, -Execute needs it closed'
    if ($running.Count -gt 0) { $runningText = "YES (pid $(($running | ForEach-Object { $_.Id }) -join ', ')) - -Execute will refuse" }
    Write-Host ("  {0,-22} {1}" -f 'OUTLOOK.EXE running', $runningText)

    Write-Host ''
    Write-Host 'No MAPI call was made, here or anywhere else in this project: IProfAdmin is measured'
    Write-Host 'broken on this Office build and the interop that wrapped it has been removed.'
    Write-Host 'Read-only. Nothing changed.'
    return
}

# ---------------------------------------------------------------------------------------------
# CLEAR THE IMPORT. One value, named before it is written, and read back.
# ---------------------------------------------------------------------------------------------
if ($ClearImportPrf) {
    Write-Host ''
    if ([string]::IsNullOrEmpty($currentImportPrf)) {
        Write-Host "ImportPRF is not set under $setupKeyPath. Nothing to clear."
        return
    }

    Write-Host "would remove : $setupKeyPath\ImportPRF   (currently '$currentImportPrf')"
    if (-not $Execute) {
        Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
        return
    }

    Remove-ItemProperty -LiteralPath $setupKeyPath -Name 'ImportPRF' -Force
    $readBack = (Get-ItemProperty -Path $setupKeyPath -Name 'ImportPRF' -ErrorAction SilentlyContinue).ImportPRF
    if (-not [string]::IsNullOrEmpty($readBack)) {
        throw "ImportPRF was removed and still reads '$readBack'. Do not start Outlook until that is resolved - a lingering ImportPRF re-imports the .prf at every start, and this one carries OverwriteProfile=Yes."
    }
    Write-Host 'ImportPRF removed. The next Outlook start will not re-import.'
    return
}

# ---------------------------------------------------------------------------------------------
# Parse -Store before anything else, so a typo fails at argument time rather than half way
# through a registry write.
# ---------------------------------------------------------------------------------------------
$parsedStores = Resolve-StoreSpec -Spec $Store
if ($null -ne $parsedStores.Problem) { throw $parsedStores.Problem }
$requested = $parsedStores.Stores

# ---------------------------------------------------------------------------------------------
# VERIFY. Registry and filesystem; -WithOutlook adds the only reading that answers what the tests
# will see.
# ---------------------------------------------------------------------------------------------
if ($Verify) {
    Write-Host ''
    Write-Host "VERIFY - did Outlook honour the .prf for '$Name'?"
    Write-Host ''

    $present = @()
    foreach ($store in $requested) {
        if (Test-Path -LiteralPath $store.Path) { $present += $store.Path }
    }

    $outcome = Test-ProfileOutcome -ProfileName $Name -ProfileNames $existingProfiles `
        -AccountSubKeyCount (Get-ProfileAccountCount -ProfilesKeyPath $profilesKeyPath -ProfileName $Name) `
        -Stores $requested -PstPathsPresent $present -StrayPstNames (Get-StrayPstName) `
        -ImportPrfValue $currentImportPrf

    foreach ($check in $outcome.Checks) {
        $marker = 'OK  '
        if ($check.Level -eq 'fail') { $marker = 'FAIL' }
        elseif ($check.Level -eq 'warn') { $marker = 'WARN' }
        Write-Host ("  {0} {1} - {2}" -f $marker, $check.What, $check.Detail)
    }

    if ($WithOutlook) {
        Write-Host ''
        Write-Host 'Reading it back over COM. This binds Outlook - starting one if none is up - and'
        Write-Host 'operates on the DEFAULT profile, because NameSpace.Logon can raise the profile'
        Write-Host 'picker even when a default is set, and a dialog on an unattended guest is a hang.'

        Invoke-WithOutlookSession -Body {
            param($ns)

            $sessionProfile = $ns.CurrentProfileName
            Write-Host "  CurrentProfileName : $sessionProfile"
            if ($sessionProfile -ne $Name) {
                Write-Host "  SKIPPED: Outlook is on '$sessionProfile', not '$Name'. This reads the default profile and"
                Write-Host "           nothing else. Make '$Name' the default first:"
                Write-Host "             .\Set-DefaultOutlookProfile.ps1 -Name $Name -Execute"
                Write-Host '           then close Outlook and run this again.'
                return
            }

            $seen = @()
            foreach ($storeItem in $ns.Stores) {
                $filePath = $null
                try { $filePath = $storeItem.FilePath } catch { $filePath = $null }
                $seen += [pscustomobject]@{ DisplayName = $storeItem.DisplayName; FilePath = $filePath }
                Write-Host ("  store              : '{0}'  <-  {1}" -f $storeItem.DisplayName, $filePath)
            }

            $accountCount = $ns.Accounts.Count
            Write-Host "  Accounts.Count     : $accountCount"

            $missing = @()
            foreach ($store in $requested) {
                $found = $false
                foreach ($actual in $seen) {
                    if ($actual.DisplayName -ceq $store.DisplayName) { $found = $true }
                }
                if (-not $found) { $missing += $store.DisplayName }
            }
            if ($missing.Count -gt 0) {
                Write-Host ''
                Write-Host "  FAIL Outlook does not show $($missing.Count) store(s) under the exact name asked for: $($missing -join ', ')"
                Write-Host '       A name that came back DIFFERENT rather than absent is the dangerous case: the store'
                Write-Host '       exists and every test that looks it up by display name misses it. Rename-OutlookStore.ps1'
                Write-Host '       fixes that in place - renaming the root folder carries through to Store.DisplayName,'
                Write-Host '       measured on this build.'
            }
            else {
                Write-Host '  OK   every store is there under the exact name asked for.'
            }

            if ($accountCount -ne 0) {
                Write-Host ''
                Write-Host "  FAIL Accounts.Count is $accountCount and this profile was built account-less."
                Write-Host '       corpus-build will refuse it - CorpusSafety.EvaluateProfile over'
                Write-Host '       ComCorpusMailbox.ReadProfileFacts, fail-closed, no override. Find out what the'
                Write-Host '       account is before doing anything else.'
            }
            else {
                Write-Host '  OK   Accounts.Count is 0 - corpus-build will accept this profile.'
            }
        }
    }

    Write-Host ''
    Write-Host "$($outcome.FailureCount) failure(s), $($outcome.WarningCount) warning(s) from the registry checks."
    if ($outcome.FailureCount -gt 0) { exit 1 }
    return
}

# ---------------------------------------------------------------------------------------------
# BUILD.
# ---------------------------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($Name)) {
    throw 'REFUSING: -Name was empty or whitespace. A profile name is a registry key name; there is no such key and there never will be.'
}
$unsafeName = Test-PrfValueSafe -Value $Name -What "profile name '$Name'"
if ($null -ne $unsafeName) { throw $unsafeName }

if ([string]::IsNullOrWhiteSpace($PrfPath)) {
    $safeName = ($Name -replace '[^A-Za-z0-9._-]', '-')
    $PrfPath = "C:\OutlookAI-Profiles\$safeName.prf"
}
$PrfPath = [System.IO.Path]::GetFullPath($PrfPath)
if ($PrfPath -match '\s') {
    throw "REFUSING: -PrfPath '$PrfPath' contains a space. The ImportPRF registry value takes a RAW UNQUOTED path, so a space in it truncates the value where nothing reports an error - Outlook simply finds no file and creates no profile. Choose a space-free path."
}

$prfText = New-ProfilePrfText -ProfileName $Name -Stores $requested -MakeDefault ([bool] $MakeDefault.IsPresent)

Write-Host ''
Write-Host "profile        : $Name"
Write-Host "make default   : $($MakeDefault.IsPresent)"
Write-Host "stores         : $($requested.Count)"
foreach ($item in $requested) {
    Write-Host ("  {0,-28} {1}" -f $item.DisplayName, $item.Path)
}
Write-Host ''
Write-Host 'It would write:'
Write-Host "  $PrfPath                    ($($prfText.Length) chars, ASCII, CRLF)"
Write-Host "  $setupKeyPath"
Write-Host "    ImportPRF   REG_SZ   = $PrfPath"
Write-Host '    First-Run   DELETED  (ImportPRF is ignored while it exists)'
Write-Host '    FirstRun    DELETED  (same)'
Write-Host ''
Write-Host 'It writes nothing else, and it does NOT start Outlook.'

if ($existingProfiles -contains $Name) {
    Write-Host ''
    Write-Host "NOTE: a profile named '$Name' already exists. The .prf carries OverwriteProfile=Yes, so the"
    Write-Host '      import REPLACES it - any store that was added to it since, and any mail in a store that'
    Write-Host '      is not re-named here, stops being part of it. That is not a delete: the .pst files stay'
    Write-Host '      on disk. Deleting a profile has no scripted route at all; see OutlookMapiInterop.ps1.'
}

if (-not $Execute) {
    Write-Host ''
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    Write-Host 'Run -Preflight first if you have not: it is the cheapest way to see whether this guest is'
    Write-Host 'in the state the import needs.'
    return
}

Assert-OutlookNotRunning

if (-not (Test-Path -LiteralPath $setupKeyPath)) {
    throw "Not found: $setupKeyPath. ImportPRF lives under it, and Outlook $($root.Version) does not look installed for this user despite the hive above passing the shape rule. Do not continue without understanding why."
}

# The DIRECTORY each .pst goes in must already exist - Outlook creates the file, not the folder -
# and Resolve-PstPath is what says so, in the one place every path in this project goes through.
foreach ($item in $requested) {
    [void](Resolve-PstPath -Path $item.Path)
}

$prfDirectory = [System.IO.Path]::GetDirectoryName($PrfPath)
if (-not (Test-Path -LiteralPath $prfDirectory)) {
    New-Item -ItemType Directory -Path $prfDirectory -Force | Out-Null
    Write-Host "Created $prfDirectory"
}

[System.IO.File]::WriteAllText($PrfPath, $prfText, (New-Object System.Text.ASCIIEncoding))
$readBack = [System.IO.File]::ReadAllText($PrfPath)
if ($readBack -cne $prfText) {
    throw "The .prf read back differently from what was written. Refusing to continue - a .prf Outlook parses differently from what this script composed is a profile built to a specification nobody has seen."
}
Write-Host "Wrote $PrfPath and read it back byte-identical."

New-ItemProperty -LiteralPath $setupKeyPath -Name 'ImportPRF' -Value $PrfPath -PropertyType String -Force | Out-Null
$importBack = (Get-ItemProperty -Path $setupKeyPath -Name 'ImportPRF' -ErrorAction SilentlyContinue).ImportPRF
if ($importBack -cne $PrfPath) {
    throw "ImportPRF was set to '$PrfPath' and reads back as '$importBack'. Do not continue."
}
Write-Host "Set ImportPRF = $importBack (and read it back)."

foreach ($valueName in @('First-Run', 'FirstRun')) {
    $value = (Get-ItemProperty -Path $setupKeyPath -Name $valueName -ErrorAction SilentlyContinue).$valueName
    if ($null -ne $value) {
        Remove-ItemProperty -LiteralPath $setupKeyPath -Name $valueName -Force
        Write-Host "Deleted $valueName"
    }
    else {
        Write-Host "$valueName already absent"
    }

    $still = (Get-ItemProperty -Path $setupKeyPath -Name $valueName -ErrorAction SilentlyContinue).$valueName
    if ($null -ne $still) {
        throw "$valueName is still present under $setupKeyPath. ImportPRF is ignored while it exists, with no diagnostic anywhere, so this refuses rather than letting you start Outlook and wonder."
    }
}

Write-Host ''
Write-Host 'Done. NEXT, AND IT IS YOURS: start Outlook once, let it finish starting, then close it.'
Write-Host 'This script does not start Outlook - importing is a startup-time action, and a setup script'
Write-Host 'that owns Outlook''s lifetime is how a guest ends up with a zombie OUTLOOK.EXE.'
Write-Host ''
Write-Host 'Then, in order:'
Write-Host ("  .\New-OutlookProfile.ps1 -Name {0}{1} -Verify" -f $Name, $(if ($requested.Count -gt 0) { ' -Store ...' } else { '' }))
Write-Host '  .\New-OutlookProfile.ps1 -ClearImportPrf -Execute   - so the next start does not re-import'
Write-Host ("  .\Set-DefaultOutlookProfile.ps1 -Name {0} -Execute   - and switch the profile prompt OFF" -f $Name)
Write-Host '  .\Add-OutlookPstStore.ps1                          - any further stores, without a re-import'
Write-Host ''
Write-Host 'The registry and the .prf are all this script can prove. Whether Outlook AGREES is proven by'
Write-Host 'starting it and reading the profile back - which is what -Verify -WithOutlook does, and what'
Write-Host 'the corpus tool prints as "profile accounts: N" when it vets a store.'
