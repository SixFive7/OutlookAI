#Requires -Version 5.1
<#
.SYNOPSIS
    Finishes the guest's configuration at first logon: the things an answer file cannot say.

.DESCRIPTION
    RUN ON THE GUEST, AT FIRST LOGON, BY THE ANSWER FILE. Windows PowerShell 5.1 - no ternary,
    no `??`, no `-p` on mkdir. Testbed/host/New-AnswerFile.ps1 copies this onto the answer-file
    volume beside autounattend.xml, and the FirstLogonCommands entry in
    Testbed/guest/autounattend.template.xml finds it on whatever drive letter that volume got
    and runs it. Running it again by hand later is safe: every step is idempotent.

    WHY IT EXISTS AT ALL. The answer file already sets SystemLocale, UserLocale, UILanguage and
    the time zone. Three things it cannot set, and one it should not:

    1. THE PREFERRED LANGUAGE LIST, because en-NL is a TRANSIENT language. Windows hands
       English (Netherlands) an LCID out of the 0x2000 transient block at runtime; on the host
       it currently sits at 2000, but that is an allocation, not an identity. An answer file has
       nowhere to put a language it cannot name, so the list is set here with
       Set-WinUserLanguageList instead, and the input method tips are rewritten to force KLID
       00020409 (United States-International) on both entries - which is what the host has for
       both of its languages.

    2. THE HOME LOCATION. GeoId 176, Netherlands.

    3. POWER. A guest that sleeps in the middle of a twelve-minute corpus build or a
       twenty-seven-minute live tier run does not fail; it hangs, and it hangs in a way that
       reads as a COM call that never returned. Fast startup goes too: hiberboot leaves a guest
       that "shut down" holding a stale kernel session, which is the wrong thing entirely when
       the point of the machine is that a rebuild is reproducible.

    4. WHAT IT DELIBERATELY DOES NOT DO: Set-WinUILanguageOverride. The host has no override.
       Its display language is en-GB because en-NL has no MUI and Windows falls back, and
       reproducing that fallback rather than pinning past it is the whole point. Pinning en-GB
       would produce the same string from Get-UICulture by a different mechanism, and the first
       time the fallback mattered the guest would not show it.

    EVERY STEP IS INDEPENDENT AND LOGGED. One failure does not stop the rest, because a guest
    that is 90% configured and says so is far more useful than one that stopped at step two
    with no record. The log ends with a readback of every setting, so the operator can diff the
    guest against the host table in Testbed/MEDIA.md instead of trusting this script.

    ELEVATION: FirstLogonCommands run in the context of the auto-logon account, which the answer
    file puts in Administrators. Set-WinSystemLocale and the powercfg calls need that. If this
    is re-run by hand from an unelevated shell those steps will fail and say so in the log.

    A REBOOT IS REQUIRED before the system locale and the language list are fully in effect.
    This script does not reboot: the caller is mid-OOBE and Windows is about to do it anyway.

    THE GUEST GUARD, SINCE 2026-09-24 - and the one thing here that DOES stop everything. Until
    then this rewrote the language list, locales, home location, power policy and screen saver of
    whatever machine ran it. It now refuses unless BOTH hold: the session is logged on as
    -ExpectedUser (vmadmin) AND the computer name starts with -ExpectedComputerNamePrefix (OAI-).
    Those defaults are the answer file's own convention, not a guess: autounattend.template.xml
    creates the account from the guest credential, whose account is vmadmin (Testbed/README.md
    section 2), and Testbed/host/New-AnswerFile.ps1 names the computer by replacing 'OutlookAI-'
    with 'OAI-' in the VM name. And they are measured at the very moment this runs: the first
    logon of OutlookAI-Unindexed's unattended install, 2026-09-15, logged
    "Complete-FirstLogon.ps1 on OAI-UNINDEXED as vmadmin" - this script's own header line, built
    from the two variables the guard reads.

    RESTATED, NOT DOT-SOURCED. Every other writing script here dot-sources Assert-TestbedGuest from
    OutlookMapiInterop.ps1, but this one travels ALONE: New-AnswerFile.ps1 puts it on the answer
    volume beside autounattend.xml and nothing else, and FirstLogonCommands runs it from there. So
    the guard is written out below, the same two-axis shape as Set-OutlookIndexingDisabled.ps1's.

    IT RUNS BEFORE THE LOG IS OPENED, so a refusal writes nothing at all - and on a guest, a
    first-logon.log that is ABSENT is how a refusal shows. (If the script is not found, the
    answer file's own command writes a log saying so, so "no log" means found and refused - or,
    less likely, died before its first line.) Run it by hand to read which. A guest built
    outside the convention - New-AnswerFile.ps1 -ComputerName without the OAI- prefix, a VM name
    that does not start 'OutlookAI-', or a credential for another account - refuses at first
    logon by design; pass -ExpectedUser and -ExpectedComputerNamePrefix and run it by hand.
    Proven on the maintainer's workstation the same day, under Windows PowerShell 5.1, with no
    arguments (as FirstLogonCommands runs it) and with -SkipPower: "REFUSING TO RUN.", zero calls
    reaching a tripwire that stood in for every write command, the native powercfg.exe and
    New-Object, and every setting it names read back unchanged. Passing, and then running to
    DONE, at a real first logon is not yet proven: that takes a fresh unattended install from an
    answer volume built after this change.

.PARAMETER LogPath
    Where the transcript of what was set goes. Under C:\Windows\Setup by default, beside the
    other setup artefacts, so it survives and is easy to find.

.PARAMETER LanguageList
    The preferred languages, most preferred first. en-NL then nl-NL matches the host.

.PARAMETER KeyboardLayout
    The KLID forced onto every language in the list. 00020409 is United States-International.

.PARAMETER GeoId
    Home location. 176 is the Netherlands.

.PARAMETER SystemLocaleName
    The non-Unicode (ANSI) system locale. en-US on the host.

.PARAMETER UserLocaleName
    The formats culture - dates, numbers, currency, first day of week. nl-NL on the host.

.PARAMETER SkipPower
    Leave the power policy and fast startup alone. For re-running the locale half only.

.PARAMETER ExpectedUser
    Accounts this script is allowed to run as. The answer file's account is vmadmin. The default
    IS the guard; do not widen it.

.PARAMETER ExpectedComputerNamePrefix
    Computer-name prefix this script is allowed to run on. Testbed/host/New-AnswerFile.ps1
    derives a guest's name by replacing 'OutlookAI-' with 'OAI-'. Pass your own if you named a
    guest something else; do not widen it to an empty string - that refuses, never matches all.

.EXAMPLE
    .\Complete-FirstLogon.ps1
    .\Complete-FirstLogon.ps1 -SkipPower -LogPath C:\Temp\relocale.log
#>
[CmdletBinding()]
param(
    [string]   $LogPath = 'C:\Windows\Setup\first-logon.log',
    [string[]] $LanguageList = @('en-NL', 'nl-NL'),
    [string]   $KeyboardLayout = '00020409',
    [int]      $GeoId = 176,
    [string]   $SystemLocaleName = 'en-US',
    [string]   $UserLocaleName = 'nl-NL',
    [switch]   $SkipPower,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-'
)

$ErrorActionPreference = 'Continue'

# ---------------------------------------------------------------------------------------------
# THE GUARD. FIRST, before the log is opened and before any step - and unlike every step below,
# a refusal stops everything. Restated rather than dot-sourced: the answer volume carries this
# file alone. See THE GUEST GUARD in the banner.
# ---------------------------------------------------------------------------------------------
function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

function Assert-TestbedGuestLocal {
    if (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $ExpectedUser -Prefix $ExpectedComputerNamePrefix) { return }
    throw @"
REFUSING TO RUN.

  logged on as : '$env:USERNAME'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$env:COMPUTERNAME'      (must start with: '$ExpectedComputerNamePrefix')

This script rewrites the language list, the system and user locales, the home location, the
power policy, fast startup and the screen saver of the machine it runs on. On the maintainer's
workstation that is a working machine's regional and power settings, changed without a word.

It is meant to run once, unattended, at a test guest's first logon, from the answer volume
Testbed/host/New-AnswerFile.ps1 builds. That answer file creates the account from the guest
credential - 'vmadmin' (Testbed/README.md section 2) - and names the computer by replacing
'OutlookAI-' with 'OAI-' in the VM name, so a guest built from it matches both axes. If you built
a guest outside that convention, say so and run this by hand:

    -ExpectedUser <username> -ExpectedComputerNamePrefix <prefix>

Nothing has been written - not even the log, which is opened only once this check has passed.
Do not 'fix' this by widening either default. The defaults are the guard.
"@
}

Assert-TestbedGuestLocal

$script:Failed = 0
$script:Ran = 0

$logDir = Split-Path -Parent $LogPath
if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
}

function Write-Line([string] $text) {
    $stamped = ('{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $text)
    Write-Output $stamped
    try { Add-Content -LiteralPath $LogPath -Value $stamped -Encoding UTF8 } catch { }
}

function Invoke-Step([string] $what, [scriptblock] $body) {
    $script:Ran++
    try {
        & $body
        Write-Line ("  OK      {0}" -f $what)
    }
    catch {
        $script:Failed++
        Write-Line ("  FAILED  {0} :: {1}" -f $what, $_.Exception.Message)
    }
}

Write-Line '================================================================'
Write-Line ("Complete-FirstLogon.ps1 on {0} as {1}" -f $env:COMPUTERNAME, $env:USERNAME)
Write-Line '================================================================'

# ---------------------------------------------------------------------------------------------
# Languages, keyboards, locale.
# ---------------------------------------------------------------------------------------------

Invoke-Step ("preferred language list = " + ($LanguageList -join ', ') + " (keyboard $KeyboardLayout on every entry)") {
    if ($LanguageList.Count -lt 1) { throw 'No languages given.' }

    # PASS 1 - REGISTER, so Windows allocates an LCID for any TRANSIENT language.
    # New-WinUserLanguageList only CONSTRUCTS an object. A transient language such as en-NL has
    # no LCID until the list has actually been SET, and with no LCID it has no input method tip
    # either. Measured on a fresh guest 2026-09-15: en-NL came back with an EMPTY
    # InputMethodTips collection, and the previous single-pass version threw right here with
    # "Windows gave 'en-NL' no input method tip to rewrite" - correctly reporting a failure, but
    # for a reason that reads as Windows misbehaving rather than as the list not existing yet.
    $seed = New-WinUserLanguageList -Language $LanguageList[0]
    for ($i = 1; $i -lt $LanguageList.Count; $i++) { $seed.Add($LanguageList[$i]) }
    Set-WinUserLanguageList -LanguageList $seed -Force

    # PASS 2 - FORCE THE LAYOUT, KEEP THE LCID. Each tip reads '<lcid-hex>:<klid>'. The LCID half
    # is whatever Windows just allocated, which for a transient language is an allocation and not
    # an identity, so only the half after the colon is rewritten.
    $list = Get-WinUserLanguageList
    foreach ($entry in $list) {
        $wanted = New-Object System.Collections.Generic.List[string]
        foreach ($tip in @($entry.InputMethodTips)) {
            $forced = [regex]::Replace($tip, ':[0-9A-Fa-f]+$', (':' + $KeyboardLayout))
            if (-not $wanted.Contains($forced)) { $wanted.Add($forced) }
        }
        if ($wanted.Count -eq 0) {
            # Still nothing after registering it. The transient block begins at 0x2000 and this
            # guest has one transient language, so 2000 is the allocation - but that is a
            # FALLBACK, not knowledge, which is why the verification below is not optional.
            $wanted.Add('2000:' + $KeyboardLayout)
        }
        $entry.InputMethodTips.Clear()
        foreach ($tip in $wanted) { $entry.InputMethodTips.Add($tip) }
    }
    Set-WinUserLanguageList -LanguageList $list -Force

    # VERIFY, because both passes above can succeed and still leave the wrong thing set.
    $final = Get-WinUserLanguageList
    $tags = @($final | ForEach-Object { $_.LanguageTag })
    foreach ($requested in $LanguageList) {
        if ($tags -notcontains $requested) {
            throw ("'{0}' is not in the language list afterwards; got: {1}" -f $requested, ($tags -join ', '))
        }
    }
    foreach ($entry in $final) {
        foreach ($tip in @($entry.InputMethodTips)) {
            if ($tip -notmatch (':' + [regex]::Escape($KeyboardLayout) + '$')) {
                throw ("'{0}' kept input method tip '{1}', which is not keyboard {2}." -f $entry.LanguageTag, $tip, $KeyboardLayout)
            }
        }
    }
}

Invoke-Step "home location GeoId $GeoId" {
    Set-WinHomeLocation -GeoId $GeoId
}

Invoke-Step "system locale (non-Unicode) = $SystemLocaleName" {
    Set-WinSystemLocale -SystemLocale $SystemLocaleName
}

Invoke-Step "user locale / formats = $UserLocaleName" {
    Set-Culture -CultureInfo $UserLocaleName
}

# ---------------------------------------------------------------------------------------------
# Power. A testbed that sleeps mid-run is a mystery failure.
# ---------------------------------------------------------------------------------------------

if ($SkipPower) {
    Write-Line '  SKIPPED power policy and fast startup (-SkipPower)'
}
else {
    Invoke-Step 'fast startup off (hibernation disabled, HiberbootEnabled = 0)' {
        # powercfg /hibernate off removes hiberfil.sys and takes fast startup with it. The
        # registry value is set as well rather than instead: they are separate switches and a
        # machine can have hibernation off and hiberboot still nominally enabled.
        $out = & powercfg.exe /hibernate off 2>&1
        if ($LASTEXITCODE -ne 0) { throw ("powercfg /hibernate off exited $LASTEXITCODE : " + ($out -join ' ')) }
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' `
            -Name 'HiberbootEnabled' -PropertyType DWord -Value 0 -Force | Out-Null
    }

    # Every timeout to zero means never, on mains and on battery alike. A VM has no battery, but
    # Hyper-V can present one to a guest and a policy that only covers AC is a trap that fires
    # once, months later.
    $timeouts = @(
        'standby-timeout-ac', 'standby-timeout-dc',
        'monitor-timeout-ac', 'monitor-timeout-dc',
        'disk-timeout-ac', 'disk-timeout-dc',
        'hibernate-timeout-ac', 'hibernate-timeout-dc'
    )
    foreach ($t in $timeouts) {
        Invoke-Step "$t = 0 (never)" {
            $out = & powercfg.exe /change $t 0 2>&1
            if ($LASTEXITCODE -ne 0) { throw ("powercfg /change $t 0 exited $LASTEXITCODE : " + ($out -join ' ')) }
        }
    }

    Invoke-Step 'screen saver off, so session 1 is never locked out from under a running test' {
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'ScreenSaveActive' -Value '0' -Force
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'ScreenSaverIsSecure' -Value '0' -Force
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'ScreenSaveTimeOut' -Value '0' -Force
    }
}

# ---------------------------------------------------------------------------------------------
# Read it all back. This is the half an operator actually uses: diff it against the host table
# in Testbed/MEDIA.md. Some of it only settles after the reboot that OOBE is about to do, which
# is why the readback says what it is and does not pretend to be a verdict.
# ---------------------------------------------------------------------------------------------

Write-Line ''
Write-Line 'READBACK (some values only settle after the next reboot):'
foreach ($probe in @(
        @{ What = 'display language (Get-UICulture)'; Body = { (Get-UICulture).Name } }
        @{ What = 'user locale (Get-Culture)'; Body = { (Get-Culture).Name } }
        @{ What = 'short date pattern'; Body = { (Get-Culture).DateTimeFormat.ShortDatePattern } }
        @{ What = 'number format 4000.5'; Body = { (4000.5).ToString('N2', (Get-Culture)) } }
        @{ What = 'system locale (Get-WinSystemLocale)'; Body = { (Get-WinSystemLocale).Name } }
        @{ What = 'home location (Get-WinHomeLocation)'; Body = { (Get-WinHomeLocation).GeoId } }
        @{ What = 'time zone (Get-TimeZone)'; Body = { (Get-TimeZone).Id } }
        @{ What = 'language list'; Body = {
                ((Get-WinUserLanguageList) | ForEach-Object {
                    $_.LanguageTag + ' [' + (($_.InputMethodTips) -join ',') + ']'
                }) -join ' ; '
            }
        }
        @{ What = 'Windows'; Body = { (Get-CimInstance Win32_OperatingSystem).Caption + ' ' + [Environment]::OSVersion.Version } }
    )) {
    try {
        Write-Line ("  {0,-38} {1}" -f $probe.What, (& $probe.Body))
    }
    catch {
        Write-Line ("  {0,-38} <could not read: {1}>" -f $probe.What, $_.Exception.Message)
    }
}

Write-Line ''
if ($script:Failed -gt 0) {
    Write-Line ("DONE WITH {0} FAILURE(S) out of {1} step(s). The guest is NOT fully configured; read the FAILED lines above." -f $script:Failed, $script:Ran)
}
else {
    Write-Line ("DONE. All {0} step(s) succeeded. Log: {1}" -f $script:Ran, $LogPath)
}
