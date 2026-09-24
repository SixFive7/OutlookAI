<#
    ============================================================================================
    RUN ON OAI-UNINDEXED 2026-09-24 IN THIS REGISTRY FORM, FROM CP-05, AND IT WORKS.
    ============================================================================================

    The banner below asked for a guest run to replace its promises with results. Every step here
    ran from a fresh restore of CP-05-CORPUS-B-CLEAN-UNINDEXED, in session 1, on Office LTSC 2024
    16.0.17932:

      -ListOnly              hive 16.0 (5 values, 15 subkeys), DefaultProfile 'CorpusProfile',
                             PickLogonProfile absent, profiles CorpusProfile and OutlookAI-Tier.
      -Name NoSuchProfile    REFUSED: "no Outlook profile named 'NoSuchProfile' exists". And it
        -Execute             wrote nothing, shown rather than asserted: DefaultProfile,
                             PickLogonProfile, the Outlook key's value list AND the key's own
                             last-write time (2026-09-16 02:11:42.939) read identical before and
                             after.
      -Name OutlookAI-Tier   with the checkpoint's Outlook still up: REFUSED by
        -Execute             Assert-OutlookNotRunning ("OUTLOOK.EXE is running (pid 9356)"), key
                             untouched. After a guest restart: DefaultProfile = 'OutlookAI-Tier'
                             and PickLogonProfile = 0 (REG_DWORD), both read back. Run a second
                             time it said AlreadyDefault and rewrote only the prompt setting.
      Outlook's own answer   after another restart, Outlook came up on the tier profile with no
                             profile prompt, and COM from session 1 read CurrentProfileName
                             'OutlookAI-Tier', Accounts.Count 1 (AccountType 2 = POP3,
                             DeliveryStore 'tier@vm.invalid'), and one store: 'tier@vm.invalid'
                             <- C:\OutlookAI-Tier\Outlook.pst.

    ONE THING THE RUN FOUND IN THIS SCRIPT, AND FIXED. The listing printed "other real Outlook
    hive(s) here: 8.0 - use -OfficeVersion if this one is wrong". On this guest
    HKCU\...\Office\8.0\Outlook holds exactly one value, First-Run, and nothing else; the shape
    rule rightly calls that a hive, and the old line then offered Office 97's hive as the fix.
    Format-OtherHiveNote names an unsupported major without offering it, and -SelfTest pins that.

    TWO THINGS IT FOUND ABOUT THE TIER PROFILE, not about this script - recorded here because this
    is the step that makes the tier profile the one Outlook opens:
      * every start of the tier profile raises the POP3 logon dialog ('Internet Email - tier',
        "Enter your user name and password for the following server"): its account points at
        127.0.0.1:110 and stores no password. It did NOT block the COM read above, which ran with
        it on screen - but it is a modal dialog on an unattended guest, on every start.
      * reading Account.SmtpAddress over COM raised Outlook's object-model guard - "A program is
        trying to access email address information stored in Outlook", Allow / Deny - and THAT
        blocks: the reading call waited on it for minutes and never returned. The same read did
        not prompt on 2026-09-15. Defender's signatures on this guest are 372 days old (last
        updated 2025-09-17; the guest has no network); an out-of-date antivirus is the documented
        trigger for that guard, but it was not proven to be the trigger here.

    ============================================================================================
    THE MAPI ROUTE IS DEAD ON THIS BUILD. THIS SCRIPT WRITES THE REGISTRY INSTEAD.
    ============================================================================================

    WHAT HAPPENED, 2026-09-16. The first execution of this script in its life - until then it
    carried a banner saying it had been verified by PARSING only - died before it wrote anything:

        Exception calling "MAPIAdminProfiles" with "2" argument(s): "Unable to cast COM object of
        type 'System.__ComObject' to interface type 'OutlookAI.Testbed.IProfAdmin'. This operation
        failed because the QueryInterface call on the COM component for the interface with IID
        '{00020379-0000-0000-C000-000000000046}' failed due to the following error: No such
        interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE))."

    WHERE. Guest OAI-UNINDEXED, user vmadmin, 64-bit elevated Windows PowerShell 5.1, Office LTSC
    2024 (ProPlus2024Volume / PerpetualVL2024), build 16.0.17932.20884.

    WHAT THAT MEANS AND WHAT IT DOES NOT. MAPIInitialize succeeded and MAPIAdminProfiles returned
    an object. What failed is the QueryInterface for IID_IProfAdmin on the object it returned - so
    this is not "MAPI is missing", and it is not a mis-declared vtable either: a wrong vtable
    produces nonsense answers or an access violation that takes the process down with no error
    text, never a clean E_NOINTERFACE from the cast itself. THE CAUSE WAS NOT ESTABLISHED, and
    deliberately so: exactly one setting in this project ever needed IProfAdmin, and that setting
    has a route which is measured working.

    WHAT REPLACED IT. HKCU\Software\Microsoft\Office\<major>\Outlook\DefaultProfile - a REG_SZ
    holding the profile name. Measured working on that same guest on that same day: the value was
    written directly, and the corpus tool then logged on to the profile the value names.

    THERE IS NO MAPI FALLBACK LEFT HERE, DELIBERATELY. A second route that nobody exercises is
    precisely the bug this file was just bitten by, twice over - the script had never run, and the
    interop it called had never run either. One route, exercised on every run, or none.

    WHAT LEAVING MAPI COSTS. The API was the contract and the registry is only the storage, so
    this script is now pinned to a layout that has already moved once (the Windows Messaging
    Subsystem hive -> the Office hive, at Outlook 2013). Two things hold that cost down: the hive
    is DETECTED rather than hardcoded (below), and every write is read back and re-checked against
    the profile list, so a layout that moves again fails loudly on the next run instead of
    reporting success into a value that nothing reads.

    HOW A PROFILE IS PROVEN TO EXIST - the other half of what MAPI used to do. The subkeys of
    ...\Outlook\Profiles ARE the profiles; a subkey name is a profile name. This script REFUSES to
    write a name that is not one of them. That refusal is the entire point of the rewrite: the
    write itself always succeeds, so a blind one leaves Outlook pointing at a profile that was
    never created, and the failure then surfaces much later as something that looks unrelated - a
    tool that binds the wrong store, or a profile prompt with nobody there to answer it.

    WHAT IT DOES. Switches which Outlook profile is the default, and switches OFF the profile
    prompt. Those are TWO settings and both have to be right.

    WHY IT MATTERS HERE. Docs/live-tier-on-the-vm.md section 1.2: corpus work happens in a profile
    with no accounts and the tier runs in a profile that has the dummy account, and switching
    between them is a restart of Outlook. It recurs - every corpus rebuild is another switch. And
    section 2.5 is blunt about the second setting: A PROMPTING PROFILE CANNOT BE DRIVEN OVER COM.
    Get the default right and leave the prompt on, and every COM call sits waiting on a dialog box
    nobody is there to answer.

    HOW THE PROMPT IS SUPPRESSED, and the honesty about it. PickLogonProfile, a DWORD under the
    Outlook root: 0 means 'always use this profile', 1 means 'prompt for a profile to be used'.
    That value is COMMUNITY-REPORTED, not documented by Microsoft. It is used here anyway for two
    reasons: it is read back and verified immediately, and its failure direction is benign - get
    it wrong and Outlook prompts, which is loud rather than silent, and which the next COM call
    reports as a hang rather than as a wrong answer.

    WHICH OFFICE HIVE. Detected, never hardcoded. The rule is this project's own, from
    OfficeVersions.IsOutlookHive: a real Outlook key has at least one VALUE, or at least one
    subkey that is not 'Resiliency'. That rule exists because Installer.iss writes a resiliency
    exemption under 15.0, 16.0 AND 17.0 on every install, so on any machine this product has
    touched, all three keys exist and a bare key-exists probe answers 16.0 just as confidently on
    an Outlook 2013 machine. The PROBE ORDER is OfficeVersions.Supported's - 16.0, then 17.0, then
    15.0 - and NOT "highest number first": a machine with two Office majors must be read the way
    the shipped product reads it, or this script and the add-in disagree about which profile set
    is real. -OfficeVersion overrides the detection, and still applies the shape rule.

    OUTLOOK MUST NOT BE RUNNING. The script refuses when it is, and does NOT kill it: mailbox
    safety rule 7 forbids taskkill on OUTLOOK.EXE outright. A running Outlook also writes its own
    view of the profile back when it closes, so killing it would trade a clean refusal for a
    silent revert an hour later.

    IDEMPOTENT. Setting the default to the profile that is already default is a no-op that still
    verifies, and the prompt setting is written on every -Execute run regardless. Run it twice and
    the second run reports 'already'.

    WHAT IS TESTED WITHOUT A GUEST. -SelfTest exercises every decision in this script - the hive
    shape rule, the hive probe order, the -OfficeVersion override, the listing, and the profile
    resolution with every one of its refusals - against synthetic inputs. It reads no registry,
    starts no process and
    is safe to run on any machine, the maintainer's workstation included. What it cannot cover is
    the registry reads and writes themselves; those are guest-only, and the self-test prints the
    list of them so nobody mistakes a green run for full coverage.

    OutlookMapiInterop.ps1 IS STILL DOT-SOURCED, for exactly two guard functions -
    Assert-TestbedGuest and Assert-OutlookNotRunning. NO MAPI CALL IS MADE ANYWHERE IN THIS FILE.
    Those two are shared rather than copied on purpose: Assert-TestbedGuest is what stands between
    this script and a real mailbox, and a copied guard is a guard that can drift.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER Name
    The profile to make default. It MUST already exist - see the refusal above.

.PARAMETER ListOnly
    Report the profiles, which one is default and what the prompt setting is. Changes nothing.

.PARAMETER SelfTest
    Run the decision tests and exit. Touches nothing: no registry, no processes, no Outlook.

.PARAMETER LeavePromptAlone
    Set the default but do not touch PickLogonProfile. For the rare case where you want the prompt
    - which on this machine is never, so it is a switch rather than the default.

.PARAMETER OfficeVersion
    Force the Office major whose hive is read and written, e.g. '16.0'. Only for a machine the
    detection reads wrongly; the named key must still be a real Outlook hive.

.PARAMETER Execute
    Without it, nothing is written.

.EXAMPLE
    .\Set-DefaultOutlookProfile.ps1 -ListOnly
    .\Set-DefaultOutlookProfile.ps1 -SelfTest
    .\Set-DefaultOutlookProfile.ps1 -Name OutlookAICorpus -Execute
#>
[CmdletBinding(DefaultParameterSetName = 'Set')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Set')] [string] $Name,
    [Parameter(Mandatory = $true, ParameterSetName = 'List')] [switch] $ListOnly,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest,
    [Parameter(ParameterSetName = 'Set')] [switch] $LeavePromptAlone,
    [string] $OfficeVersion,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

# The Office majors this product supports, IN PROBE ORDER, mirrored from OfficeVersions.Supported.
# "Newest first" is what the comment there says; 16.0-then-17.0-then-15.0 is what the array there
# does, and the array is what ships.
$script:SupportedOfficeVersions = @('16.0', '17.0', '15.0')

# The subkey this product's own installer creates under every supported major, which is why a
# bare key-exists probe proves nothing. OfficeVersions.InstallerFootprintSubKeyName.
$script:InstallerFootprintSubKeyName = 'Resiliency'

$script:OfficeRootKeyPath = 'HKCU:\Software\Microsoft\Office'

# =============================================================================================
# PURE DECISIONS. No registry, no processes, no Outlook, no output. Everything this script
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

    Returns Chosen (a candidate, or $null), RealHives (every candidate passing the shape rule, in
    the order given) and Problem (the refusal text, when Chosen is $null).

    -RequestedVersion overrides the probe order and the supported list both: an operator who names
    a version has said something this script cannot know. It does NOT override the shape rule,
    because writing DefaultProfile into a key Outlook never reads is a silent no-op - which is the
    failure mode this whole rewrite exists to avoid.
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
                Problem   = "REFUSING: -OfficeVersion $RequestedVersion was asked for, and $($named.Path) holds nothing but a '$script:InstallerFootprintSubKeyName' subkey - which this product's own installer writes under every supported major, on every install. Outlook does not read that key, so a DefaultProfile written into it would be a silent no-op. Name the major Outlook really uses, or drop -OfficeVersion."
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

    # Nothing supported qualified. If some OTHER major did, name it - choosing it is an operator
    # decision (-OfficeVersion), not one this script may take on itself.
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
    The line(s) printed under the chosen hive when other Outlook-shaped keys exist. Pure.

    MEASURED 2026-09-24 on OAI-UNINDEXED: HKCU\Software\Microsoft\Office\8.0\Outlook holds exactly
    one value, First-Run, and nothing else. The shape rule rightly calls any key with a value a
    hive, so the first guest run printed "other real Outlook hive(s) here: 8.0 - use -OfficeVersion
    if this one is wrong" - an invitation to point Outlook at Office 97's hive. An unsupported
    major is still NAMED, because hiding it would also hide a genuine second Outlook, but it is no
    longer offered as the fix.
#>
function Format-OtherHiveNote {
    param([string] $ChosenVersion, [object[]] $RealHives)

    $supported = @()
    $unsupported = @()
    if ($null -ne $RealHives) {
        foreach ($hive in $RealHives) {
            if ($null -eq $hive -or $hive.Version -eq $ChosenVersion) { continue }
            if ($script:SupportedOfficeVersions -contains $hive.Version) { $supported += $hive.Version }
            else { $unsupported += $hive.Version }
        }
    }

    $lines = @()
    if ($supported.Count -gt 0) {
        $lines += "other supported Outlook hive(s) here: $($supported -join ', ') - use -OfficeVersion if this one is wrong"
    }
    if ($unsupported.Count -gt 0) {
        $lines += "also an Outlook-shaped key under unsupported major(s) $($unsupported -join ', ') - never chosen, and not a fix for anything"
    }
    return , $lines
}

<#
    How to render PickLogonProfile for a human. $null means the value is absent, which is Outlook's
    own default and means 'do not prompt'.
#>
function Format-PromptSetting {
    param($Value)

    if ($null -eq $Value) { return 'absent (Outlook default: always use the default profile)' }
    if ($Value -eq 0) { return '0 - always use this profile' }
    if ($Value -eq 1) { return '1 - PROMPT FOR A PROFILE (COM cannot be driven)' }
    return "$Value - unrecognised"
}

<#
    The profile listing, as printed. Pure, so the marker and the warning underneath it are
    exercised without a registry.

    THE WARNING IS THE INTERESTING LINE. A DefaultProfile naming a profile that is not there is
    the exact state this script refuses to create, and a machine can arrive in it by other means -
    a profile deleted through Mail in Control Panel, or an older tool that wrote the value blind.
    Saying so on every run costs one comparison.
#>
function Format-ProfileListing {
    param([string[]] $ProfileNames, [string] $CurrentDefault, [string] $ProfilesKeyPath)

    $names = @()
    if ($null -ne $ProfileNames) { $names = @($ProfileNames) }

    $lines = @("profiles ($($names.Count)), read from ${ProfilesKeyPath}:")
    if ($names.Count -eq 0) { $lines += '  <none>' }

    foreach ($profileName in $names) {
        $marker = ''
        # Case-insensitive, like the registry itself: a default that differs only in case DOES
        # find the profile, and marking it is not the same as calling the value correct.
        # Resolve-DefaultProfileChange is what decides whether it gets rewritten.
        if ($profileName -eq $CurrentDefault) { $marker = '   <- DefaultProfile names this one' }
        $lines += ("  {0}{1}" -f $profileName, $marker)
    }

    if (-not [string]::IsNullOrEmpty($CurrentDefault) -and ($names -notcontains $CurrentDefault)) {
        $lines += ''
        $lines += "WARNING: DefaultProfile is '$CurrentDefault' and no profile by that name exists. Outlook has"
        $lines += '         nothing to log on to. That is exactly the state this script refuses to create.'
    }

    return , $lines
}

<#
    The whole default-profile decision, made against the profile names read out of the registry.

    Decision is one of:
      InvalidName    the name is empty or whitespace
      NoProfiles     this machine has no Outlook profile at all
      NoSuchProfile  the name is not one of them              <- the refusal this rewrite is for
      AlreadyDefault the registry already names it, byte for byte
      Change         write it

    ResolvedName is THE PROFILE'S OWN SPELLING, not the caller's. Registry key names match
    case-insensitively, so -Name corpusprofile finds the key 'CorpusProfile' - and 'CorpusProfile'
    is what gets written, because the value is read by something else entirely and there is no
    reason to hand it a spelling that appears nowhere on the machine.
#>
function Resolve-DefaultProfileChange {
    param(
        [string] $RequestedName,
        [string[]] $ProfileNames,
        [string] $CurrentDefault,
        [string] $ProfilesKeyPath
    )

    $names = @()
    if ($null -ne $ProfileNames) { $names = @($ProfileNames) }

    $currentText = '<not set>'
    if (-not [string]::IsNullOrEmpty($CurrentDefault)) { $currentText = "'$CurrentDefault'" }

    if ([string]::IsNullOrWhiteSpace($RequestedName)) {
        return [pscustomobject]@{
            Decision     = 'InvalidName'
            ResolvedName = $null
            Message      = 'REFUSING: -Name was empty or whitespace. A profile name is a registry key name; there is no such key and there never will be.'
        }
    }

    if ($names.Count -eq 0) {
        return [pscustomobject]@{
            Decision     = 'NoProfiles'
            ResolvedName = $null
            Message      = @"
REFUSING: this machine has no Outlook profile at all.

Read from $ProfilesKeyPath
That key has no subkeys - and a subkey of it IS a profile, so there is nothing here to make
default. DefaultProfile would be written, Outlook would find no such profile at its next logon,
and the failure would surface later as something that looks unrelated.

Create a profile first, then run this again.
"@
        }
    }

    $exact = $null
    $folded = $null
    foreach ($candidate in $names) {
        if ($candidate -ceq $RequestedName) { $exact = $candidate }
        elseif ($candidate -eq $RequestedName) { $folded = $candidate }
    }

    $resolved = $exact
    if ($null -eq $resolved) { $resolved = $folded }

    if ($null -eq $resolved) {
        return [pscustomobject]@{
            Decision     = 'NoSuchProfile'
            ResolvedName = $null
            Message      = @"
REFUSING: no Outlook profile named '$RequestedName' exists on this machine.

Read from $ProfilesKeyPath, this machine has: $($names -join ', ')
The current DefaultProfile value is $currentText.

The write itself would have succeeded - which is exactly why this check exists. A DefaultProfile
naming a profile nobody created leaves Outlook with nothing to log on to at its next start, and
the damage then shows up much later looking like a different fault entirely: a tool that binds the
wrong store, or a profile prompt with nobody there to answer it.

Create it first (Testbed/README.md section 5 lists the profile scripts), or pass one of the names
above. Casing does not matter - the name written is always the profile's own spelling.
"@
        }
    }

    if ($CurrentDefault -ceq $resolved) {
        return [pscustomobject]@{
            Decision     = 'AlreadyDefault'
            ResolvedName = $resolved
            Message      = "'$resolved' is already the DefaultProfile value, byte for byte - no change needed."
        }
    }

    $note = ''
    if ($null -eq $exact) {
        $note = " (you asked for '$RequestedName'; the profile's own spelling is '$resolved', and that is what gets written)"
    }
    elseif ($CurrentDefault -eq $resolved) {
        $note = " (the current value differs from the profile's own spelling only in case; it is being rewritten to match)"
    }

    return [pscustomobject]@{
        Decision     = 'Change'
        ResolvedName = $resolved
        Message      = "DefaultProfile $currentText -> '$resolved'$note"
    }
}

# =============================================================================================
# SELF-TEST. Pure: reads no registry, starts no process, needs no guest. Runs anywhere.
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

    Write-Host 'Set-DefaultOutlookProfile self-test. Nothing is read and nothing is written.'
    Write-Host ''
    Write-Host '== the hive shape rule (OfficeVersions.IsOutlookHive) =='

    Test-Case 'the installer footprint alone is not a hive' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('Resiliency'))
    Test-Case 'nor is it in another casing' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('RESILIENCY'))
    Test-Case 'an empty key is not a hive' $false (Test-IsOutlookHive -ValueNames @() -SubKeyNames @())
    Test-Case 'nulls are not a hive, and do not throw' $false (Test-IsOutlookHive -ValueNames $null -SubKeyNames $null)
    Test-Case 'one value is enough' $true (Test-IsOutlookHive -ValueNames @('DefaultProfile') -SubKeyNames @('Resiliency'))
    Test-Case 'one non-footprint subkey is enough' $true (Test-IsOutlookHive -ValueNames @() -SubKeyNames @('Profiles', 'Resiliency'))
    Test-Case 'a used hive is a hive' $true (Test-IsOutlookHive -ValueNames @('DefaultProfile', 'PickLogonProfile') -SubKeyNames @('Profiles', 'Options', 'Setup', 'Resiliency'))

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

    # The order is OfficeVersions.Supported's, NOT "highest number first". A sort-descending probe
    # would answer 17.0 here and disagree with the shipped add-in on the same machine.
    $pick = Select-OutlookHive -Candidates @($real16, $real17)
    Test-Case '16.0 wins over a real 17.0 (probe order, not highest number)' '16.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell16, $real17)
    Test-Case '17.0 wins when 16.0 is only the footprint' '17.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell16, $shell17, $real15)
    Test-Case '15.0 is chosen when it is the only real one' '15.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell15, $shell16, $shell17)
    Test-Case 'all-footprint refuses' '<null>' $pick.Chosen
    Test-Case 'and says what that means' $true ($pick.Problem -like '*classic Outlook is not installed*')
    Test-Case 'and names the probe order it used' $true ($pick.Problem -like '*16.0, 17.0, 15.0*')

    $pick = Select-OutlookHive -Candidates @()
    Test-Case 'no version keys at all refuses' '<null>' $pick.Chosen

    $pick = Select-OutlookHive -Candidates $null
    Test-Case 'a null candidate list refuses, and does not throw' '<null>' $pick.Chosen

    $pick = Select-OutlookHive -Candidates @($shell16, $real14)
    Test-Case 'an unsupported major is never chosen for you' '<null>' $pick.Chosen
    Test-Case 'but it is named, with the flag that would use it' $true ($pick.Problem -like '*-OfficeVersion 14.0*')

    $pick = Select-OutlookHive -Candidates @($real16, $real15) -RequestedVersion '15.0'
    Test-Case '-OfficeVersion overrides the probe order' '15.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($shell16, $real14) -RequestedVersion '14.0'
    Test-Case '-OfficeVersion reaches an unsupported major' '14.0' $pick.Chosen.Version

    $pick = Select-OutlookHive -Candidates @($real16) -RequestedVersion '15.0'
    Test-Case '-OfficeVersion naming an absent key refuses' '<null>' $pick.Chosen
    Test-Case 'and lists the version keys that are present' $true ($pick.Problem -like '*Version keys present: 16.0*')

    $pick = Select-OutlookHive -Candidates @($real16, $shell15) -RequestedVersion '15.0'
    Test-Case '-OfficeVersion does NOT override the shape rule' '<null>' $pick.Chosen
    Test-Case 'and says why that key is not real' $true ($pick.Problem -like '*silent no-op*')

    Write-Host ''
    Write-Host '== the note under the chosen hive =='

    # The measured shape: 8.0\Outlook holding one value, First-Run, beside a real 16.0.
    $stray8 = New-HiveCandidate -Version '8.0' -ValueNames @('First-Run')
    $pick = Select-OutlookHive -Candidates @($real16, $stray8)
    Test-Case 'the 8.0 First-Run shell does not displace 16.0' '16.0' $pick.Chosen.Version
    $note = Format-OtherHiveNote -ChosenVersion $pick.Chosen.Version -RealHives $pick.RealHives
    Test-Case 'it is still named - one line' 1 $note.Count
    Test-Case 'but as unsupported, never as the fix' $true ($note[0].Contains('unsupported major(s) 8.0 - never chosen'))
    Test-Case 'and -OfficeVersion is NOT offered for it' $false ($note[0].Contains('-OfficeVersion'))
    $note = Format-OtherHiveNote -ChosenVersion '16.0' -RealHives @($real16, $real15)
    Test-Case 'a second SUPPORTED hive still gets the -OfficeVersion hint' $true ($note[0].Contains('other supported Outlook hive(s) here: 15.0 - use -OfficeVersion'))
    $note = Format-OtherHiveNote -ChosenVersion '16.0' -RealHives @($real16)
    Test-Case 'the chosen hive alone prints nothing' 0 $note.Count
    $note = Format-OtherHiveNote -ChosenVersion '16.0' -RealHives $null
    Test-Case 'a null list prints nothing, and does not throw' 0 $note.Count

    Write-Host ''
    Write-Host '== the prompt setting, as printed =='

    Test-Case 'absent reads as Outlook default' 'absent (Outlook default: always use the default profile)' (Format-PromptSetting $null)
    Test-Case '0 reads as no prompt' '0 - always use this profile' (Format-PromptSetting 0)
    Test-Case '1 reads as the COM-breaking one' '1 - PROMPT FOR A PROFILE (COM cannot be driven)' (Format-PromptSetting 1)
    Test-Case 'anything else is reported, never guessed' '7 - unrecognised' (Format-PromptSetting 7)

    Write-Host ''
    Write-Host '== the profile listing, as printed =='

    $profiles = @('Outlook', 'CorpusProfile', 'TierProfile')
    $profilesKey = "$script:OfficeRootKeyPath\16.0\Outlook\Profiles"

    $listing = Format-ProfileListing -ProfileNames $profiles -CurrentDefault 'CorpusProfile' -ProfilesKeyPath $profilesKey
    Test-Case 'header, three profiles, no warning' 4 $listing.Count
    Test-Case 'the header counts them and names the key' "profiles (3), read from ${profilesKey}:" $listing[0]
    Test-Case 'the default one is marked' '  CorpusProfile   <- DefaultProfile names this one' $listing[2]
    Test-Case 'the others are not' '  TierProfile' $listing[3]

    $listing = Format-ProfileListing -ProfileNames $profiles -CurrentDefault 'corpusprofile' -ProfilesKeyPath $profilesKey
    Test-Case 'a case-different default still marks its profile' '  CorpusProfile   <- DefaultProfile names this one' $listing[2]
    Test-Case 'and raises no missing-profile warning' 4 $listing.Count

    $listing = Format-ProfileListing -ProfileNames $profiles -CurrentDefault 'Vanished' -ProfilesKeyPath $profilesKey
    Test-Case 'a default naming nothing WARNS' $true (($listing -join "`n") -like "*WARNING: DefaultProfile is 'Vanished'*")
    Test-Case 'and says what that state is' $true (($listing -join "`n") -like '*nothing to log on to*')

    $listing = Format-ProfileListing -ProfileNames @() -CurrentDefault '' -ProfilesKeyPath $profilesKey
    Test-Case 'no profiles prints <none>' '  <none>' $listing[1]
    Test-Case 'and an unset default raises no warning' 2 $listing.Count

    $listing = Format-ProfileListing -ProfileNames $null -CurrentDefault $null -ProfilesKeyPath $profilesKey
    Test-Case 'nulls print the empty listing, and do not throw' 2 $listing.Count

    Write-Host ''
    Write-Host '== the default-profile decision =='

    $decision = Resolve-DefaultProfileChange -RequestedName 'CorpusProfile' -ProfileNames $profiles -CurrentDefault 'Outlook' -ProfilesKeyPath $profilesKey
    Test-Case 'an existing profile is a Change' 'Change' $decision.Decision
    Test-Case 'and resolves to itself' 'CorpusProfile' $decision.ResolvedName

    $decision = Resolve-DefaultProfileChange -RequestedName 'CorpusProfile' -ProfileNames $profiles -CurrentDefault 'CorpusProfile' -ProfilesKeyPath $profilesKey
    Test-Case 'the profile that is already default is a no-op' 'AlreadyDefault' $decision.Decision

    $decision = Resolve-DefaultProfileChange -RequestedName 'corpusprofile' -ProfileNames $profiles -CurrentDefault 'Outlook' -ProfilesKeyPath $profilesKey
    Test-Case 'a case-different name still finds the profile' 'Change' $decision.Decision
    Test-Case "and is written in the profile's own spelling" 'CorpusProfile' $decision.ResolvedName
    Test-Case 'and says so out loud' $true ($decision.Message -like "*the profile's own spelling is 'CorpusProfile'*")

    $decision = Resolve-DefaultProfileChange -RequestedName 'CorpusProfile' -ProfileNames $profiles -CurrentDefault 'corpusprofile' -ProfilesKeyPath $profilesKey
    Test-Case 'a value differing only in case is rewritten, not called equal' 'Change' $decision.Decision
    Test-Case 'and says why' $true ($decision.Message -like '*only in case*')

    $decision = Resolve-DefaultProfileChange -RequestedName 'NoSuchThing' -ProfileNames $profiles -CurrentDefault 'Outlook' -ProfilesKeyPath $profilesKey
    Test-Case 'a profile that does not exist REFUSES' 'NoSuchProfile' $decision.Decision
    Test-Case 'and lists what is there' $true ($decision.Message -like '*Outlook, CorpusProfile, TierProfile*')
    Test-Case 'and names the key it read them from' $true ($decision.Message -like "*$profilesKey*")

    $decision = Resolve-DefaultProfileChange -RequestedName 'Anything' -ProfileNames @() -CurrentDefault '' -ProfilesKeyPath $profilesKey
    Test-Case 'a machine with no profiles REFUSES' 'NoProfiles' $decision.Decision

    $decision = Resolve-DefaultProfileChange -RequestedName 'Anything' -ProfileNames $null -CurrentDefault $null -ProfilesKeyPath $profilesKey
    Test-Case 'a null profile list is that same refusal, not a crash' 'NoProfiles' $decision.Decision

    $decision = Resolve-DefaultProfileChange -RequestedName '   ' -ProfileNames $profiles -CurrentDefault 'Outlook' -ProfilesKeyPath $profilesKey
    Test-Case 'a whitespace name REFUSES' 'InvalidName' $decision.Decision

    $decision = Resolve-DefaultProfileChange -RequestedName 'CorpusProfile ' -ProfileNames $profiles -CurrentDefault 'Outlook' -ProfilesKeyPath $profilesKey
    Test-Case 'a trailing space is not silently trimmed into a match' 'NoSuchProfile' $decision.Decision

    $decision = Resolve-DefaultProfileChange -RequestedName 'CorpusProfile' -ProfileNames $profiles -CurrentDefault '' -ProfilesKeyPath $profilesKey
    Test-Case 'no DefaultProfile value yet is still a Change' 'Change' $decision.Decision
    Test-Case 'and the empty value is reported honestly' $true ($decision.Message -like '*<not set>*')

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. These need a guest, and nothing on this machine can stand in for them:'
    Write-Host '  * reading the Office version keys and the shape of each Outlook subkey out of HKCU'
    Write-Host '  * reading DefaultProfile, PickLogonProfile and the Profiles subkeys'
    Write-Host '  * the two guards (Assert-TestbedGuest, Assert-OutlookNotRunning)'
    Write-Host '  * the DefaultProfile and PickLogonProfile writes, and their read-backs'
    Write-Host '  * whether Outlook honours the value at its next logon'

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

# Dot-sourced for Assert-TestbedGuest and Assert-OutlookNotRunning ONLY. No MAPI call is made
# anywhere in this script - see the banner at the top.
. "$PSScriptRoot\OutlookMapiInterop.ps1"

# Asserted before ANYTHING, including the read-only paths: this script reads the Outlook profile
# registry, and on the maintainer's workstation that is a real profile with real delegate
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

$selection = Select-OutlookHive -Candidates (Get-OutlookHiveCandidate) -RequestedVersion $OfficeVersion
if ($null -eq $selection.Chosen) { throw $selection.Problem }

$root = $selection.Chosen
$profilesKeyPath = "$($root.Path)\Profiles"

Write-Host "outlook hive : $($root.Path)   (Office $($root.Version); $($root.ValueNames.Count) value(s), $($root.SubKeyNames.Count) subkey(s))"
foreach ($line in (Format-OtherHiveNote -ChosenVersion $root.Version -RealHives $selection.RealHives)) {
    Write-Host "               $line"
}

$currentDefault = (Get-ItemProperty -Path $root.Path -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
$currentPrompt = (Get-ItemProperty -Path $root.Path -Name 'PickLogonProfile' -ErrorAction SilentlyContinue).PickLogonProfile

Write-Host "DefaultProfile   : $currentDefault"
Write-Host "PickLogonProfile : $(Format-PromptSetting $currentPrompt)"
Write-Host ''

# The pre-2013 hive is REPORTED and never written. Profiles moved out of it at Outlook 2013, and a
# script that reads the wrong one does not fail loudly - it reports a working setup as broken,
# which is what happened to New-TierProfile.ps1 on 2026-09-15.
$legacyProfilesKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
if (Test-Path -LiteralPath $legacyProfilesKey) {
    $legacyDefault = (Get-ItemProperty -Path $legacyProfilesKey -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
    Write-Host "note: the pre-2013 profile hive also exists here, with DefaultProfile = '$legacyDefault'. It is"
    Write-Host '      NOT written by this script, and Outlook 2013 and later do not read it.'
    Write-Host ''
}

$profileNames = Get-OutlookProfileName -ProfilesKeyPath $profilesKeyPath
foreach ($line in (Format-ProfileListing -ProfileNames $profileNames -CurrentDefault $currentDefault -ProfilesKeyPath $profilesKeyPath)) {
    Write-Host $line
}

if ($ListOnly) {
    Write-Host ''
    Write-Host 'Read-only. Nothing changed.'
    return
}

$decision = Resolve-DefaultProfileChange -RequestedName $Name -ProfileNames $profileNames `
    -CurrentDefault $currentDefault -ProfilesKeyPath $profilesKeyPath

# The refusals fire BEFORE the -Execute gate on purpose: a dry run that says "would set the default
# to a profile that does not exist" has told you the useful thing, and holding that back until the
# real run wastes the dry run entirely.
if ($decision.Decision -eq 'InvalidName' -or $decision.Decision -eq 'NoProfiles' -or $decision.Decision -eq 'NoSuchProfile') {
    throw $decision.Message
}

Write-Host ''
Write-Host "would set default to : $($decision.ResolvedName)   [$($decision.Decision)]"
Write-Host "  $($decision.Message)"
if (-not $LeavePromptAlone) { Write-Host 'would set PickLogonProfile to 0 (always use this profile)' }
Write-Host ''

if (-not $Execute) {
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    return
}

Assert-OutlookNotRunning

if ($decision.Decision -eq 'AlreadyDefault') {
    Write-Host $decision.Message
}
else {
    Set-ItemProperty -Path $root.Path -Name 'DefaultProfile' -Value $decision.ResolvedName -Type String

    # READ IT BACK, ORDINALLY. A registry write cannot fail visibly - that is the whole hazard of
    # this route - so the only evidence anything happened is the value itself, compared
    # case-sensitively against the profile's own spelling.
    $readBack = (Get-ItemProperty -Path $root.Path -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
    if ($readBack -cne $decision.ResolvedName) {
        throw "DefaultProfile was set to '$($decision.ResolvedName)' and reads back as '$readBack'. Do not continue - something else is writing this key."
    }
    Write-Host "DefaultProfile = '$readBack' (written to $($root.Path), and read back)."
}

# And the profile is STILL there afterwards. Cheap, and it catches the one ordering that would
# otherwise slip through: a profile removed between the listing above and the write.
$profilesAfter = Get-OutlookProfileName -ProfilesKeyPath $profilesKeyPath
if ($profilesAfter -notcontains $decision.ResolvedName) {
    throw "DefaultProfile now names '$($decision.ResolvedName)' and $profilesKeyPath no longer has a subkey by that name. Do not start Outlook until that is resolved."
}

if (-not $LeavePromptAlone) {
    Set-ItemProperty -Path $root.Path -Name 'PickLogonProfile' -Value 0 -Type DWord
    $readBack = (Get-ItemProperty -Path $root.Path -Name 'PickLogonProfile' -ErrorAction SilentlyContinue).PickLogonProfile
    if ($readBack -ne 0) {
        throw "PickLogonProfile was set to 0 and reads back as '$readBack'. Outlook will prompt for a profile, and a prompting profile cannot be driven over COM."
    }
    Write-Host 'PickLogonProfile = 0 - Outlook will always use the default profile, no prompt.'
}

Write-Host ''
Write-Host 'Done. Outlook picks this up at its next start; a running Outlook would not have seen it,'
Write-Host 'which is why this script refuses to run while one is up.'
Write-Host ''
Write-Host 'The registry is all this script can prove. Whether Outlook AGREES is proven by starting it'
Write-Host 'and reading the profile back over COM - which is what Build-Corpus.ps1 preflights, and what'
Write-Host 'the corpus tool prints as "profile accounts: N" when it vets a store.'
