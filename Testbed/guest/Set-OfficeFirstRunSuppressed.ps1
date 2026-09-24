<#
    ============================================================================================
    Suppresses the Office first-run dialogs, so a scripted guest build is not stopped by one.
    ============================================================================================

    RUN, AND WHAT IT DOES NOT COVER (2026-09-24, OAI-UNINDEXED, from a fresh CP-02, twice). As
    vmadmin in session 1 the guard passed and -Execute wrote all 13 values, creating six of the
    seven keys they live in, then -Verify read every one back. At Outlook's FIRST start after it,
    with every visible top-level window enumerated, the only dialog was the POP3 password prompt:
    the privacy consent screen this was written for did not appear. But the SECOND start
    of Outlook on the machine - whichever profile it opens - shows Office's one-time "Check out
    our new look" dialog (window class NUIDialog), which nothing here suppresses, and no registry
    value recording it was found. COM worked with it on screen. It matters anyway: when that second
    start was the tier profile's import, the POP3 account came up unbound until the next start -
    see New-TierProfile.ps1, and Testbed/README.md section 1 for the order that avoids it.

    It was not new on the guests: .work\run-forcepst.ps1 and .work\guest2-chain.ps1 ran it with
    -Execute on 2026-09-15, and then ran `New-Item -Force` on the Outlook key, which deleted the
    three values it had just written there - CP-05 lacks exactly those three.

    RUN ON THE GUEST, in the interactive session, as the account Outlook will run as. Windows
    PowerShell 5.1 - no ternary, no `??`.

    WHY THIS EXISTS. ImportPRF suppresses Outlook's PROFILE wizard and nothing else. Measured on
    a guest on 2026-09-15: a profile imported perfectly from a .prf, and Outlook then came up
    showing "Your privacy matters" - Office's own first-run consent dialog. A dialog on an
    unattended guest is not a prompt, it is a hang: the object model cannot answer one, and every
    COM call behind it waits forever.

    WHAT IS ACTUALLY WELL-ATTESTED, and what is not. The three first-run values below are
    long-standing and appear in Microsoft's own ADMX tooling and in every deployment guide. The
    PRIVACY values are policy keys that Microsoft documents for controlling connected
    experiences, and using them to suppress the first-run consent screen is INFERRED rather than
    documented. That is why this script ends by asking whether the dialog actually went away
    instead of reporting success from having written registry values - writing a value proves
    nothing about whether it was read.

    NOT A POLICY HACK ON A REAL MACHINE. Everything here is HKCU on a disposable testbed guest
    whose only purpose is running Outlook headlessly. Do not run it on a workstation; the privacy
    settings in particular are a deliberate choice this project has made for a machine with no
    user and no network, and are not a recommendation for anyone else.

    "DO NOT RUN IT ON A WORKSTATION" WAS A SENTENCE, NOT A CHECK, UNTIL 2026-09-24. -Execute wrote
    thirteen values - five of them Office privacy POLICY keys - on whatever machine it was started
    on. It now dot-sources OutlookMapiInterop.ps1 and calls Assert-TestbedGuest, the guard every
    other writing script here uses, before anything else runs - the dry run and -Verify included.
    STAGE OutlookMapiInterop.ps1 BESIDE IT. Proven on the maintainer's workstation the same day
    with -Execute: "REFUSING TO RUN. This session is logged on as ...", and all seven keys it names
    read back unchanged - every value, and each key's own last-write time.

.PARAMETER OfficeVersion
    Office hive version. 16.0 covers Outlook 2016 through 2024 and Microsoft 365.

.PARAMETER Execute
    Write. Without it, prints what it would set and changes nothing.

.PARAMETER Verify
    Read every value back and report. Safe at any time; writes nothing.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.EXAMPLE
    .\Set-OfficeFirstRunSuppressed.ps1
    .\Set-OfficeFirstRunSuppressed.ps1 -Execute
    .\Set-OfficeFirstRunSuppressed.ps1 -Verify
#>
[CmdletBinding()]
param(
    [string] $OfficeVersion = '16.0',
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute,
    [switch] $Verify
)

$ErrorActionPreference = 'Stop'

# THE GUARD, FIRST - before the plan is printed and before -Verify reads a value. See the banner.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

function Say($m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# Every value this script touches, named in one place so the blast radius is readable.
# 'Why' is carried with each so a reader does not have to go looking.
$settings = @(
    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Common\General"
       Name = 'ShownFirstRunOptin'; Type = 'DWord'; Value = 1
       Why  = 'The classic first-run opt-in marker. Present means "already shown".' }

    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\FirstRun"
       Name = 'BootedRTM'; Type = 'DWord'; Value = 1
       Why  = 'Marks the first-run boot as already done.' }

    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\FirstRun"
       Name = 'disablemovie'; Type = 'DWord'; Value = 1
       Why  = 'Suppresses the first-run welcome animation.' }

    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Registration"
       Name = 'AcceptAllEulas'; Type = 'DWord'; Value = 1
       Why  = 'Accepts the licence terms so no EULA page appears.' }

    # ---- privacy / connected experiences. INFERRED for this purpose - see the header. ----
    @{ Key = "HKCU:\Software\Policies\Microsoft\office\$OfficeVersion\common\privacy"
       Name = 'disconnectedstate'; Type = 'DWord'; Value = 2
       Why  = 'Connected experiences off. 2 = disconnected. INFERRED as the first-run-consent suppressor.' }

    @{ Key = "HKCU:\Software\Policies\Microsoft\office\$OfficeVersion\common\privacy"
       Name = 'usercontentdisabled'; Type = 'DWord'; Value = 2
       Why  = 'Experiences analysing user content off.' }

    @{ Key = "HKCU:\Software\Policies\Microsoft\office\$OfficeVersion\common\privacy"
       Name = 'downloadcontentdisabled'; Type = 'DWord'; Value = 2
       Why  = 'Experiences downloading online content off.' }

    @{ Key = "HKCU:\Software\Policies\Microsoft\office\$OfficeVersion\common\privacy"
       Name = 'controllerconnectedservicesenabled'; Type = 'DWord'; Value = 2
       Why  = 'Optional connected experiences off.' }

    @{ Key = "HKCU:\Software\Policies\Microsoft\office\$OfficeVersion\common\privacy"
       Name = 'sendtelemetry'; Type = 'DWord'; Value = 3
       Why  = 'Diagnostic data level. 3 = neither required nor optional data sent.' }

    # ---- the NEW Outlook, which has no COM object model and would end the project ----
    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Options\General"
       Name = 'HideNewOutlookToggle'; Type = 'DWord'; Value = 1
       Why  = 'Hides the "Try the new Outlook" switch. Documented by Microsoft.' }

    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Options\General"
       Name = 'DoNewOutlookAutoMigration'; Type = 'DWord'; Value = 0
       Why  = 'Blocks automatic migration to the new Outlook.' }

    @{ Key = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Preferences"
       Name = 'UseNewOutlook'; Type = 'DWord'; Value = 0
       Why  = 'Keeps classic Outlook as the client. The new one has no COM object model, so this is load-bearing, not cosmetic.' }

    @{ Key = "HKCU:\Software\Policies\Microsoft\office\$OfficeVersion\outlook\preferences"
       Name = 'NewOutlookMigrationUserSetting'; Type = 'DWord'; Value = 0
       Why  = 'Policy form of the same block. Documented by Microsoft.' }
)

if (-not ($Execute -or $Verify)) {
    Say 'DRY RUN. Would set:'
    foreach ($s in $settings) { "    {0}\{1} = {2}   # {3}" -f $s.Key, $s.Name, $s.Value, $s.Why }
    Say 'Re-run with -Execute to write, or -Verify to read back.'
    return
}

if ($Execute) {
    Say '== Execute =='
    foreach ($s in $settings) {
        if (-not (Test-Path $s.Key)) { New-Item -Path $s.Key -Force | Out-Null; Say "  created $($s.Key)" }
        New-ItemProperty -Path $s.Key -Name $s.Name -PropertyType $s.Type -Value $s.Value -Force | Out-Null
        Say "  set $($s.Name) = $($s.Value)"
    }
}

Say '== Verify =='
$bad = 0
foreach ($s in $settings) {
    $actual = $null
    if (Test-Path $s.Key) {
        $p = Get-ItemProperty -Path $s.Key -ErrorAction SilentlyContinue
        if ($p) { $actual = $p.($s.Name) }
    }
    if ($actual -eq $s.Value) { Say ("  OK   {0} = {1}" -f $s.Name, $actual) }
    else { $bad++; Say ("  FAIL {0} = {1}, wanted {2}" -f $s.Name, $actual, $s.Value) }
}

Say ''
if ($bad -eq 0) {
    Say 'Every value is set.'
} else {
    Say "$bad value(s) are not set."
}

# THE ONLY CHECK THAT MATTERS, and it is not the one above. Registry values prove that something
# was written, not that Outlook read it. The question is whether a dialog still appears, and the
# only way to know is to start Outlook and look.
Say ''
Say 'REGISTRY VALUES ARE NOT THE ANSWER. Start Outlook and confirm no dialog appears:'
Say '    $p = Start-Process OUTLOOK.EXE -PassThru; Start-Sleep 60'
Say '    Get-Process | Where-Object { $_.MainWindowHandle -ne 0 } | Select ProcessName, MainWindowTitle'
Say 'A window titled anything other than an Outlook folder view is a dialog, and a dialog on an'
Say 'unattended guest is a hang rather than a prompt.'

if ($bad -gt 0) { exit 1 }
exit 0
