#Requires -Version 5.1
<#
    ============================================================================================
    THE REGISTRY HALF HAS RUN ON A GUEST (2026-09-24). THE EFFECT IS STILL UNVERIFIED.
    ============================================================================================

    Run on OAI-UNINDEXED from a fresh CP-02 (Office LTSC 2024, 16.0.17932.20996), as vmadmin in
    session 1: the guard passed; -Execute wrote all 18 values, creating three of the four keys,
    and read every one back; a second -Execute wrote them again and read them back, creating
    nothing; -Revert -Execute removed all 18 and left the keys in place, as documented - the
    three it had created are left empty, and Outlook's own Setup key keeps its own values. Each
    run exited 0.

    WHAT IS STILL NOT KNOWN is the only thing this script is for: whether Outlook then opens the
    CLASSIC account wizard on this Office 2024 build. Nobody has opened the Mail applet afterwards
    and looked, and a value reading back proves the write, not the effect. The original
    banner's doubt stands until that is done (Dump-UiaTree.ps1 on whichever wizard appears).
    It was first written by an agent forbidden to read or write any Outlook registry key, on a
    workstation holding real mail, and verified by parsing only.

.SYNOPSIS
    Restores Outlook's CLASSIC account-setup wizard and stops AutoDiscover reaching the network.
    A prerequisite for the UI Automation route; harmless and useful on its own.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`.
    NEVER run this on the maintainer's workstation: it changes how Outlook adds accounts.

    THAT SENTENCE WAS THE ONLY THING STOPPING IT UNTIL 2026-09-24. -Execute writes eighteen HKCU
    values - nine names, each under both the user key and the Policies key - with no check of which
    machine it was on, and -Revert removes them just as blindly. It now
    dot-sources OutlookMapiInterop.ps1 and calls Assert-TestbedGuest, the guard every other
    writing script here uses, before anything else runs - the dry run included, so the plan is
    never printed on a machine it would refuse to change. STAGE OutlookMapiInterop.ps1 BESIDE IT.
    Proven on the maintainer's workstation the same day with -Execute: "REFUSING TO RUN. This
    session is logged on as ...", and all four keys it names read back unchanged - every value,
    and each key's own last-write time.

    TWO THINGS, AND THEY ARE INDEPENDENT.

    1. DisableOffice365SimplifiedAccountCreation = 1 (REG_DWORD).
       Since Click-to-Run 16.0.6769.2015, Outlook's Add Account flow is a one-box "enter your
       email address" wizard that is drawn by Office rather than by Windows. That surface is a
       poor automation target: Office chrome is DirectUI, which typically exposes no
       AutomationId, so a script would have to key on localised control names - on a guest whose
       display language is en-GB - and the managed UI Automation API cannot reach
       LegacyIAccessible at all. This value restores the classic Win32 property-sheet wizard,
       which is the surface route C can actually drive.

       Microsoft documents it in KB3189194 at both of these, and this script writes both because
       the two reports of it failing both concern precedence:
           HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\setup
           HKCU\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\setup

       THE RISK, STATED RATHER THAN BURIED. An unresolved 2025 Microsoft Q&A thread reports that
       this stopped working in Office 2024, and Microsoft never answered it. The testbed guests
       ARE Office 2024 - ProPlus2024Volume / PerpetualVL2024, build 16.0.17932.20884 per
       Testbed/MEDIA.md. So this may simply not work here, and the way you find out is by opening
       the Add Account dialog afterwards and seeing which wizard appears. This script cannot
       check that for you: the value reading back correctly proves the write, not the effect.

    2. The AutoDiscover Exclude* values (REG_DWORD, all 1).
       The guest's network is Internal, Private or disconnected. A wizard that tries to reach
       autodiscover.<domain>, an SCP lookup in a non-existent AD, or an SRV record on a resolver
       that is not there does not fail fast - it hangs, and a UI Automation script waiting for
       the next page then times out for a reason that has nothing to do with UI Automation.
       These turn the lookups off. On the classic path they are belt and braces, because
       "Manual setup or additional server types" bypasses AutoDiscover anyway; on any other path
       they are the difference between a ten-second wizard and a multi-minute one.

       NOT WRITTEN, DELIBERATELY: ExcludeSrvLookup. Microsoft's own troubleshooting page says it
       "doesn't exist in Outlook code" - it is a widely-copied myth. Only ExcludeSrvRecord is
       read. Writing the myth would leave a value that looks like configuration and is not.

    IDEMPOTENT BY CONSTRUCTION. Every write is a set-to-a-constant, and every one is read back
    afterwards. Running it twice changes nothing the second time and says so.

    IT WRITES NOTHING ELSE. No profile, no account, no store, no mail item. -Revert removes
    exactly what -Execute wrote and nothing else.

.PARAMETER OfficeVersion
    Office major version. 16.0 covers Outlook 2016 through 2024 and Microsoft 365.

.PARAMETER Execute
    Actually write. Without it the plan is printed and nothing changes.

.PARAMETER Revert
    Remove the values this script writes, leaving the keys themselves alone.

.PARAMETER SkipAutoDiscover
    Write only the wizard value, not the AutoDiscover ones. For a guest that genuinely has a
    reachable mail server and wants AutoDiscover left working.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.EXAMPLE
    .\Set-AccountWizardClassic.ps1
    .\Set-AccountWizardClassic.ps1 -Execute
    .\Set-AccountWizardClassic.ps1 -Revert -Execute
#>
[CmdletBinding()]
param(
    [string] $OfficeVersion = '16.0',
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute,
    [switch] $Revert,
    [switch] $SkipAutoDiscover
)

$ErrorActionPreference = 'Stop'

# THE GUARD, FIRST - before the plan is even printed. See the DESCRIPTION.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

$setupUser   = "HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Outlook\setup"
$setupPolicy = "HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Outlook\setup"
$autoUser    = "HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\Outlook\AutoDiscover"
$autoPolicy  = "HKCU:\SOFTWARE\Policies\Microsoft\Office\$OfficeVersion\Outlook\AutoDiscover"

$wizardValue = 'DisableOffice365SimplifiedAccountCreation'

# Every one of these is documented by Microsoft as read by Outlook. ExcludeSrvLookup is NOT in
# this list on purpose - see the header.
$autoDiscoverValues = @(
    'PreferLocalXML',
    'ExcludeHttpRedirect',
    'ExcludeHttpsAutoDiscoverDomain',
    'ExcludeHttpsRootDomain',
    'ExcludeScpLookup',
    'ExcludeSrvRecord',
    'ExcludeLastKnownGoodURL',
    'ExcludeExplicitO365Endpoint'
)

$script:Failures = @()

function Pass {
    param([string] $What, [string] $Detail)
    if ($Detail) { Write-Host "  OK   $What - $Detail" } else { Write-Host "  OK   $What" }
}

function Fail {
    param([string] $What, [string] $Why)
    $script:Failures += "$What : $Why"
    Write-Host "  FAIL $What - $Why"
}

function Set-Dword {
    param([string] $Key, [string] $Name, [int] $Value)

    if (-not (Test-Path -LiteralPath $Key)) {
        New-Item -Path $Key -Force | Out-Null
        Write-Host "  created key $Key"
    }
    New-ItemProperty -LiteralPath $Key -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null

    $back = $null
    try { $back = (Get-ItemProperty -LiteralPath $Key -Name $Name -ErrorAction Stop).$Name } catch { $back = $null }
    if ($back -eq $Value) {
        Pass "$Name" "$Key = $Value"
    }
    else {
        Fail "$Name" "wrote $Value to $Key but read back '$back'"
    }
}

function Remove-Value {
    param([string] $Key, [string] $Name)

    if (-not (Test-Path -LiteralPath $Key)) {
        Pass "$Name removed" "$Key does not exist"
        return
    }
    $present = $null
    try { $present = Get-ItemProperty -LiteralPath $Key -Name $Name -ErrorAction Stop } catch { $present = $null }
    if ($null -eq $present) {
        Pass "$Name removed" "already absent from $Key"
        return
    }
    Remove-ItemProperty -LiteralPath $Key -Name $Name -Force

    $still = $null
    try { $still = Get-ItemProperty -LiteralPath $Key -Name $Name -ErrorAction Stop } catch { $still = $null }
    if ($null -eq $still) {
        Pass "$Name removed" $Key
    }
    else {
        Fail "$Name removed" "it is still present in $Key"
    }
}

# ---------------------------------------------------------------------------------------------

Write-Host "Set-AccountWizardClassic.ps1 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ''

$targets = @()
$targets += New-Object psobject -Property @{ Key = $setupUser;   Name = $wizardValue }
$targets += New-Object psobject -Property @{ Key = $setupPolicy; Name = $wizardValue }
if (-not $SkipAutoDiscover) {
    foreach ($v in $autoDiscoverValues) {
        $targets += New-Object psobject -Property @{ Key = $autoUser;   Name = $v }
        $targets += New-Object psobject -Property @{ Key = $autoPolicy; Name = $v }
    }
}

if (-not $Execute) {
    if ($Revert) { Write-Host 'Plan: REVERT. It would remove these values (the keys are left alone):' }
    else         { Write-Host 'Plan: it would set each of these to REG_DWORD 1:' }
    Write-Host ''
    foreach ($t in $targets) {
        Write-Host ("  {0}`n      {1}" -f $t.Key, $t.Name)
    }
    Write-Host ''
    Write-Host 'It writes nothing else - no profile, no account, no store, no mail item.'
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    exit 0
}

if ($Revert) {
    Write-Host 'Reverting.'
    Write-Host ''
    foreach ($t in $targets) { Remove-Value -Key $t.Key -Name $t.Name }
}
else {
    Write-Host 'Writing.'
    Write-Host ''
    foreach ($t in $targets) { Set-Dword -Key $t.Key -Name $t.Name -Value 1 }
}

Write-Host ''
if ($script:Failures.Count -gt 0) {
    Write-Host ("{0} value(s) FAILED:" -f $script:Failures.Count)
    foreach ($f in $script:Failures) { Write-Host "  - $f" }
    exit 1
}

Write-Host 'All values read back as written.'
if (-not $Revert) {
    Write-Host ''
    Write-Host 'THAT PROVES THE WRITE, NOT THE EFFECT. Whether Outlook now opens the CLASSIC wizard'
    Write-Host 'is a separate question, and on Office 2024 it is genuinely in doubt. Open the Mail'
    Write-Host 'applet and look:'
    Write-Host ''
    Write-Host '  C:\Windows\System32\control.exe "C:\Program Files\Microsoft Office\root\Office16\MLCFG32.CPL"'
    Write-Host ''
    Write-Host 'A page offering "Manual setup or additional server types" is the classic wizard and'
    Write-Host 'route C can proceed. A single box asking only for an email address is the simplified'
    Write-Host 'one, and it cannot. Then run Testbed/guest/Dump-UiaTree.ps1 on whichever you get.'
}
exit 0
