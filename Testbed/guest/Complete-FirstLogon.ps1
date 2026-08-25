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
    [switch]   $SkipPower
)

$ErrorActionPreference = 'Continue'

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

    $list = New-WinUserLanguageList -Language $LanguageList[0]
    for ($i = 1; $i -lt $LanguageList.Count; $i++) {
        $list.Add($LanguageList[$i])
    }

    # FORCE THE LAYOUT, KEEP THE LCID. Each tip reads '<lcid-hex>:<klid>' and the LCID half is
    # whatever Windows assigned - for a transient language such as en-NL that is an allocation
    # made at runtime and there is no constant to write here. So rewrite only the half after the
    # colon and leave the half in front of it exactly as Windows produced it.
    foreach ($entry in $list) {
        $existing = @($entry.InputMethodTips)
        $wanted = New-Object System.Collections.Generic.List[string]
        foreach ($tip in $existing) {
            $forced = [regex]::Replace($tip, ':[0-9A-Fa-f]+$', (':' + $KeyboardLayout))
            if (-not $wanted.Contains($forced)) { $wanted.Add($forced) }
        }
        if ($wanted.Count -eq 0) {
            throw ("Windows gave '{0}' no input method tip to rewrite." -f $entry.LanguageTag)
        }
        $entry.InputMethodTips.Clear()
        foreach ($tip in $wanted) { $entry.InputMethodTips.Add($tip) }
    }

    Set-WinUserLanguageList -LanguageList $list -Force
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
