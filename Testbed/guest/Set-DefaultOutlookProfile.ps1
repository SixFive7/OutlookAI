<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it - the machine it was written on is the maintainer's
    own workstation, and this script reads and writes the Outlook profile registry. Verified by
    PARSING only. Replace this banner with what it actually did once it has run on a guest.

    WHAT IT DOES. Switches which Outlook profile is the default, and switches OFF the profile
    prompt. Those are TWO settings and both have to be right.

    WHY IT MATTERS HERE. Docs/live-tier-on-the-vm.md section 1.2: corpus work happens in a profile
    with no accounts and the tier runs in a profile that has the dummy account, and switching
    between them is a restart of Outlook. It recurs - every corpus rebuild is another switch. And
    section 2.5 is blunt about the second setting: A PROMPTING PROFILE CANNOT BE DRIVEN OVER COM.
    Get the default right and leave the prompt on, and every COM call sits waiting on a dialog box
    nobody is there to answer.

    Testbed/README.md section 6 item 5 is the gap this closes: "Which Outlook profile is default,
    and how the switch is automated ... the exact value and whether anything automates it is
    unrecorded."

    HOW THE DEFAULT IS SET. IProfAdmin::SetDefaultProfile - the documented, supported call. Not a
    registry poke: the registry is the storage, the API is the contract, and the API is what keeps
    working when the storage moves (which it already did once, from the Windows Messaging
    Subsystem hive to the Office one).

    HOW THE PROMPT IS SUPPRESSED, and the honesty about it. PickLogonProfile, a DWORD under the
    Outlook root: 0 means 'always use this profile', 1 means 'prompt for a profile to be used'.
    That value is COMMUNITY-REPORTED, not documented by Microsoft. It is used here anyway for two
    reasons: it is read back and verified immediately, and its failure direction is benign - get
    it wrong and Outlook prompts, which is loud rather than silent, and which the next COM call
    reports as a hang rather than as a wrong answer.

    WHICH OFFICE HIVE. Detected, never hardcoded. The rule is this project's own, from
    OfficeVersions.IsOutlookHive and pinned by .github/scripts/check-pinned-constants.ps1: a real
    Outlook key has at least one VALUE, or at least one subkey that is not 'Resiliency'. That rule
    exists because Installer.iss writes a resiliency exemption under 15.0, 16.0 AND 17.0 on every
    install, so on any machine this product has touched, all three keys exist and a bare
    key-exists probe answers 16.0 just as confidently on an Outlook 2013 machine.

    OUTLOOK MUST NOT BE RUNNING. The script refuses when it is, and does NOT kill it: mailbox
    safety rule 7 forbids taskkill on OUTLOOK.EXE outright. A running Outlook also writes its own
    view of the profile back when it closes, so killing it would trade a clean refusal for a
    silent revert an hour later.

    IDEMPOTENT. Setting the default to the profile that is already default is a no-op that still
    verifies. Run it twice and the second run reports 'already'.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER Name
    The profile to make default.

.PARAMETER ListOnly
    Report the profiles, which one is default and what the prompt setting is. Changes nothing.

.PARAMETER LeavePromptAlone
    Set the default but do not touch PickLogonProfile. For the rare case where you want the prompt
    - which on this machine is never, so it is a switch rather than the default.

.PARAMETER Execute
    Without it, nothing is written.

.EXAMPLE
    .\Set-DefaultOutlookProfile.ps1 -ListOnly
    .\Set-DefaultOutlookProfile.ps1 -Name OutlookAICorpus -Execute
#>
[CmdletBinding(DefaultParameterSetName = 'Set')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Set')] [string] $Name,
    [Parameter(Mandatory = $true, ParameterSetName = 'List')] [switch] $ListOnly,
    [Parameter(ParameterSetName = 'Set')] [switch] $LeavePromptAlone,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\OutlookMapiInterop.ps1"

# Asserted before ANYTHING, including the read-only paths: this script reads the Outlook profile
# registry, and on the maintainer's workstation that is a real profile with real delegate
# mailboxes. Refusing early costs nothing and is the whole point of the guard.
Assert-TestbedGuest -ExpectedUser $ExpectedUser

<#
    Finds the Outlook hive the way this product finds it: by SHAPE, not by trying 16.0 first.
    Mirrors OfficeVersions.IsOutlookHive - at least one value, or a subkey other than Resiliency.
#>
function Get-OutlookRootKey {
    $office = 'HKCU:\Software\Microsoft\Office'
    if (-not (Test-Path $office)) {
        throw "No $office key. Is Office installed on this machine at all?"
    }

    $candidates = @(Get-ChildItem -Path $office -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
            Sort-Object { [double] $_.PSChildName } -Descending)

    foreach ($candidate in $candidates) {
        $outlook = Join-Path $candidate.PSPath 'Outlook'
        if (-not (Test-Path $outlook)) { continue }

        $key = Get-Item -LiteralPath $outlook
        $valueNames = @($key.GetValueNames())
        $subKeyNames = @($key.GetSubKeyNames())
        $realSubKeys = @($subKeyNames | Where-Object { $_ -ne 'Resiliency' })

        if ($valueNames.Count -gt 0 -or $realSubKeys.Count -gt 0) {
            return [pscustomobject]@{
                Path    = $outlook
                Version = $candidate.PSChildName
                Values  = $valueNames.Count
                SubKeys = $subKeyNames.Count
            }
        }
    }

    throw @"
Found no Outlook hive with any content under $office.

Every version key present held nothing but a 'Resiliency' subkey, which this product writes on
every install under 15.0, 16.0 and 17.0 alike - so an empty-but-present key proves nothing. Either
classic Outlook is not installed, or it has never been started.
"@
}

$root = Get-OutlookRootKey
Write-Host "outlook hive : $($root.Path)   (Office $($root.Version); $($root.Values) value(s), $($root.SubKeys) subkey(s))"

$currentDefault = (Get-ItemProperty -Path $root.Path -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
$currentPrompt = (Get-ItemProperty -Path $root.Path -Name 'PickLogonProfile' -ErrorAction SilentlyContinue).PickLogonProfile

$promptText = 'absent (Outlook default: always use the default profile)'
if ($null -ne $currentPrompt) {
    if ($currentPrompt -eq 0) { $promptText = '0 - always use this profile' }
    elseif ($currentPrompt -eq 1) { $promptText = '1 - PROMPT FOR A PROFILE (COM cannot be driven)' }
    else { $promptText = "$currentPrompt - unrecognised" }
}

Write-Host "DefaultProfile   : $currentDefault"
Write-Host "PickLogonProfile : $promptText"
Write-Host ''

Invoke-WithProfAdmin -Body {
    param($admin)
    $profiles = Get-MapiProfile -ProfAdmin $admin
    Write-Host "profiles ($($profiles.Count)):"
    foreach ($p in $profiles) {
        $marker = ''
        if ($p.IsDefault) { $marker = '   <- default, per MAPI' }
        Write-Host ("  {0}{1}" -f $p.Name, $marker)
    }
}

if ($ListOnly) {
    Write-Host ''
    Write-Host 'Read-only. Nothing changed.'
    return
}

Write-Host ''
Write-Host "would set default to : $Name"
if (-not $LeavePromptAlone) { Write-Host 'would set PickLogonProfile to 0 (always use this profile)' }
Write-Host ''

if (-not $Execute) {
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    return
}

Assert-OutlookNotRunning

Invoke-WithProfAdmin -Body {
    param($admin)

    $profiles = Get-MapiProfile -ProfAdmin $admin
    $exists = $false
    $alreadyDefault = $false
    foreach ($p in $profiles) {
        if ($p.Name -eq $Name) {
            $exists = $true
            if ($p.IsDefault) { $alreadyDefault = $true }
        }
    }
    if (-not $exists) {
        $names = @()
        foreach ($p in $profiles) { $names += $p.Name }
        throw "No profile named '$Name'. This machine has: $($names -join ', '). Create it with New-OutlookProfile.ps1 first - switching to a profile that does not exist would leave Outlook with no profile at all."
    }

    if ($alreadyDefault) {
        Write-Host "'$Name' is already the default, per the profile table - no change made."
    }
    else {
        Assert-MapiOk -HResult $admin.SetDefaultProfile($Name, 0) -What "IProfAdmin::SetDefaultProfile('$Name')"

        # VERIFY THROUGH THE TABLE. An HRESULT from this interface is not evidence: DeleteProfile
        # on the same interface returns S_OK without deleting when the profile is in use.
        $after = Get-MapiProfile -ProfAdmin $admin
        $ok = $false
        foreach ($p in $after) {
            if ($p.Name -eq $Name -and $p.IsDefault) { $ok = $true }
        }
        if (-not $ok) {
            throw "SetDefaultProfile returned success and the profile table still does not report '$Name' as default. Do not continue - the call and the table disagree, and the table is the one that is true."
        }
        Write-Host "Default profile is now '$Name' (confirmed in the profile table)."
    }
}

# The registry half is a second, independent reading of the same fact. If MAPI and the registry
# disagree, something is writing this behind our back and that is worth stopping for.
$registryDefault = (Get-ItemProperty -Path $root.Path -Name 'DefaultProfile' -ErrorAction SilentlyContinue).DefaultProfile
if ($registryDefault -ne $Name) {
    throw "MAPI reports '$Name' as default and $($root.Path)\DefaultProfile says '$registryDefault'. Those cannot both be right. Do not continue."
}
Write-Host "Registry agrees: DefaultProfile = $registryDefault"

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
