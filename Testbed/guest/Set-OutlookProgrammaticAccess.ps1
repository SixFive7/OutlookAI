#Requires -Version 5.1
<#
    ============================================================================================
    RUN ON OutlookAI-Unindexed 2026-09-24 (FROM CP-06) AND IT WORKS - PROVEN AGAINST A CONTROL.
    ============================================================================================

    Written the same day for Q80. Office LTSC 2024 16.0.17932; Defender signatures 372 days old
    (WSC productState 0x061110, 'signatures out of date'); the tier profile running (one POP3
    account, 'OutlookAI tier sink'); Outlook started as a program in session 1, with its POP3 logon
    dialog on screen the whole time:

      -SelfTest     43 checks, 0 failures - on the guest's Windows PowerShell 5.1, and on the host
                    under 5.1 and 7.
      THE CONTROL   -Verify with the nine values ABSENT: the guard prompt ("A program is trying to
                    access email address information stored in Outlook") appeared 0.7 s after the
                    Account.SmtpAddress read; -Verify answered Deny, the prompt closed and the read
                    returned '' - VERDICT: PROMPTED, exit 1. So the proof below is not vacuous.
      -Execute      Outlook closed; the key did not exist; all nine written and read back REG_DWORD.
      -Verify       after a fresh Outlook start: SmtpAddress='tier@vm.invalid' in 16 ms, no prompt
                    - VERDICT: NO-PROMPT, exit 0.
      ON THE HOST   the dry run, -Execute, -Verify and -Execute -Revert each stopped at "REFUSING TO
                    RUN. This session is logged on as 'jori'", and the host's own
                    ...\Office\16.0\Outlook\Security policy key read identical before and after.

    ONE THING THE FIRST GUEST RUN FOUND, AND FIXED. Attaching from a second process to an Outlook
    that was started as a program (not by COM) makes COM launch a TRANSIENT 'OUTLOOK.EXE -Embedding'
    that hands the call to the running instance and exits within seconds - seen twice (pids 524 and
    9436). The first -Verify called that "the attach started another Outlook" and answered
    NOT-VERIFIABLE; it now waits up to 30 s and refuses only a second Outlook that STAYS.

    NOT EXERCISED ON A GUEST: -Execute -Revert.

.SYNOPSIS
    Auto-approves Outlook's programmatic-access prompts (the Object Model Guard) on an OFFLINE TEST
    GUEST, through the documented Outlook security Group Policy - and then PROVES it by reading
    Account.SmtpAddress over COM, inside a time limit, with no prompt.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`. -Verify runs in the
    INTERACTIVE session, through Testbed/guest/Register-InteractiveTask.ps1. NEVER on the
    maintainer's workstation: the guard below refuses there, and on that machine the Object Model
    Guard is doing its job.

    WHY THIS EXISTS - Q80, DECIDED BY THE MAINTAINER 2026-09-24: auto-approve Outlook's
    programmatic-access prompt on the offline test guests ONLY. Outlook's Object Model Guard lets
    an out-of-process COM caller read protected members (Account.SmtpAddress, MailItem.Body,
    Recipients, PropertyAccessor, ...) without a prompt only while Windows Security Center reports
    the antivirus "Good" [MS-DOC 1]. The guests have no network, so Defender's signatures are out
    of date (372 days old on 2026-09-24: WSC productState 0x061110, 'signatures out of date'), and
    every out-of-process caller - the MCP server the live tier drives, every corpus verb, every
    testbed probe - gets "A program is trying to access email address information stored in
    Outlook", Allow / Deny: a MODAL dialog that blocks the calling COM request until a human
    answers. On an unattended guest that is a hang (Docs/live-tier-on-the-vm.md section 8 item 23).
    CP-05 of OutlookAI-Unindexed carries one in its saved memory.

    THE MAINTAINER'S WORKSTATION HAS NO SUCH PROMPT, because its antivirus is current, and under
    that condition Microsoft documents that cross-process callers run WITHOUT security warnings
    [MS-DOC 1]. Approving every guarded category here is what makes a guest behave like that
    machine - it does not make the guest more permissive than the machine the tier is compared with.

    ============================================================================================
    WHAT IT WRITES, AND WHERE EVERY VALUE COMES FROM
    ============================================================================================

    Key HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security, all REG_DWORD:

      AdminSecurityMode                    3   "Use the security policy from the GPO settings"
                                               [MS-DOC 2]. Microsoft's current baseline reference
                                               says the Outlook Security Mode policy "must be
                                               enabled and the Outlook Security Policy dropdown set
                                               to 'Use Outlook Security Group Policy'" for ANY of
                                               the dependent Outlook security policies to apply
                                               [MS-DOC 3]. Without it, the eight values below are
                                               read by nothing.
      PromptOOMAddressInformationAccess    2   reading address information - Account.SmtpAddress,
                                               the member measured to prompt on these guests
      PromptOOMAddressBookAccess           2   accessing an address book
      PromptOOMSend                        2   sending items - the product's own two-step `send`
      PromptOOMMeetingTaskRequestResponse  2   responding to meeting and task requests
      PromptOOMSaveAs                      2   the Save As command
      PromptOOMFormulaAccess               2   UserProperty.Formula
      PromptOOMAddressUserPropertyFind     2   address information through UserProperties.Find
      PromptOOMCustomAction                2   Outlook object model custom actions

    For every PromptOOM* value Microsoft documents exactly three data: 0 = Automatically deny,
    1 = Prompt user (the default), 2 = Automatically approve [MS-DOC 2]. KB 926512 is written for
    Outlook 2007 under ...\12.0\...; the same value names are what the current policy templates
    configure under the version key - 16.0 is Outlook 2016 through LTSC 2024, which these guests
    run (16.0.17932). Every category Microsoft documents for the object model is set, because the
    point is that NO prompt can appear, as on a machine whose antivirus is current. The three
    Simple MAPI prompts are deliberately NOT set: KB 926512 says those settings "were not added to
    the product", and nothing in this repository uses Simple MAPI.

    WHICH ROUTE, AND WHY THIS ONE. Microsoft documents two:

      A. THIS ONE - the per-user Group Policy above [MS-DOC 2, 3].
      B. The machine-wide Trust Center setting: HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security
         ObjectModelGuard = 2, "Never warn me about suspicious activity (not recommended)"
         [MS-DOC 4], which per [MS-DOC 1] means "the Object Model Guard will be disabled".

    A was chosen, for four reasons:
      1. It APPROVES each guarded category by name; B switches the guard off altogether. A
         is the smaller statement, and it is the one this decision asked for.
      2. It is a policy: Outlook reads it from HKCU\Software\Policies, which Office's Click-to-Run
         registry virtualisation does not redirect. B lives under HKLM\SOFTWARE\Microsoft\Office,
         which on a Click-to-Run install (the guests' Office is one) has a virtualised twin under
         ...\ClickToRun\REGISTRY\MACHINE, and Microsoft's page for B names only the plain and
         Wow6432Node paths. Which copy a C2R Outlook reads for B was NOT tested here.
      3. A policy outranks the Trust Center: with it set, the Programmatic Access page cannot
         override it from the UI. B IS the Trust Center setting.
      4. Nothing is lost by its being per-user: each guest has exactly one Windows account,
         vmadmin, the autologon account Outlook runs as (Testbed/README.md section 2).

    ============================================================================================
    WHAT -Verify PROVES, AND WHAT IT DOES NOT
    ============================================================================================

    NOT a read-back of the registry - that proves a value was written, not that Outlook read it.
    It attaches to the RUNNING Outlook from a child job and reads Account.SmtpAddress for every
    account, while the parent watches this session's windows for a guard prompt:

      NO-PROMPT       every account's SmtpAddress came back non-empty within -TimeoutSeconds and
                      no guard prompt appeared. exit 0.
      PROMPTED        a guard prompt appeared. -Verify ANSWERS IT 'DENY' - the answer that grants
                      nothing - so it never leaves a hang behind, and FAILS. exit 1.
      TIMED-OUT       no prompt seen and no answer within -TimeoutSeconds. The child job is stopped
                      (it is this script's own process; Outlook is never touched). exit 1.
      NOT-VERIFIABLE  session 0, Outlook not running, the attach left a SECOND Outlook running, or
                      the running profile has no account to read. exit 2. (A TRANSIENT
                      'OUTLOOK.EXE -Embedding' that COM launches, which hands the request to the
                      running Outlook and exits within seconds, is expected when Outlook was started
                      as a program rather than by COM - measured - and is reported, not refused.)

    The registry values are printed beside the verdict as context, never as the verdict.

    It does not start or stop Outlook, ever: start it on the profile you mean (a profile with a
    mail account - the tier profile) first. It reads no mail item and writes nothing.

    ============================================================================================
    SOURCES
    ============================================================================================

      [MS-DOC 1] "Security Behavior of the Outlook Object Model", Microsoft Learn (Office VBA),
                 learn.microsoft.com/en-us/office/vba/outlook/how-to/security/security-behavior-of-the-outlook-object-model
                 - cross-process callers run without warnings when WSC reports the antivirus
                 "Good"; Group Policy and the Trust Center override that default.
      [MS-DOC 2] "Information about e-mail security settings" (KB 926512), Microsoft Learn,
                 learn.microsoft.com/en-us/microsoft-365-apps/outlook/email-security/email-security-settings
                 - AdminSecurityMode 0..3 and every PromptOOM* value with 0/1/2.
      [MS-DOC 3] "List of settings for the Microsoft 365 Apps for Enterprise security baseline in
                 Intune", Microsoft Learn, section "Security > Security Form Settings",
                 learn.microsoft.com/en-us/intune/device-security/security-baselines/ref-v2-office-settings
                 - Outlook Security Mode must be 'Use Outlook Security Group Policy' for the
                 dependent policies to apply.
      [MS-DOC 4] "Program is trying to send an e-mail message on your behalf" (KB 3189806),
                 Microsoft Learn,
                 learn.microsoft.com/en-us/troubleshoot/outlook/security/a-program-is-trying-to-send-an-email-message-on-your-behalf
                 - the machine-wide ObjectModelGuard alternative, not used.

    THE GUARD. Two axes, both required: logged on as vmadmin (the shared Assert-TestbedGuest in
    OutlookMapiInterop.ps1 - STAGE IT BESIDE THIS SCRIPT) AND a computer name starting 'OAI-'
    (Testbed/host/New-AnswerFile.ps1 names every guest that way). It runs before anything else
    except -SelfTest, which reads nothing and writes nothing.

.PARAMETER Execute
    Write the nine values. Refused while OUTLOOK.EXE runs: Outlook reads its security policy when
    it starts, so the next start has to be one that happens after the write. With -Revert,
    removes exactly those nine values instead.

.PARAMETER Revert
    With -Execute: delete the nine values this script writes, and nothing else - not the key, not
    any other value in it.

.PARAMETER Verify
    The COM proof above. Session 1, Outlook running on a profile with a mail account.

.PARAMETER TimeoutSeconds
    How long -Verify waits for the reads. Loopback COM against a local PST answers in well under a
    second; the default is generous so that a cold Outlook is not reported as a failure.

.PARAMETER SelfTest
    The pure decisions against synthetic inputs, and this file's own syntax tree. No registry, no
    COM, no guard. Runs anywhere, the host included.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.PARAMETER ExpectedComputerNamePrefix
    The computer-name prefix the second axis of the guard accepts.

.PARAMETER LogPath
    This script's own log.

.EXAMPLE
    .\Set-OutlookProgrammaticAccess.ps1 -SelfTest
    .\Set-OutlookProgrammaticAccess.ps1                  # the plan; reads and writes nothing
    .\Set-OutlookProgrammaticAccess.ps1 -Execute         # Outlook closed
    .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\Set-OutlookProgrammaticAccess.ps1' -Verify"
#>
[CmdletBinding()]
param(
    [switch]   $Execute,
    [switch]   $Revert,
    [switch]   $Verify,
    [switch]   $SelfTest,
    [int]      $TimeoutSeconds             = 90,
    [string]   $OfficeVersion              = '16.0',
    [string[]] $ExpectedUser               = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [string]   $LogPath                    = 'C:\OutlookAI-Q5\set-outlook-programmatic-access.log'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Everything this script writes, in one place. HKCU only, one key, nine REG_DWORD values.
# ---------------------------------------------------------------------------------------------
$PolicySubKey = "Software\Policies\Microsoft\Office\$OfficeVersion\Outlook\Security"

$PolicyValues = @(
    @{ Name = 'AdminSecurityMode';                   Data = 3; Why = 'Use the security policy from the GPO settings - required for every value below to be read [MS-DOC 2, 3]' }
    @{ Name = 'PromptOOMAddressInformationAccess';   Data = 2; Why = 'reading address information (Account.SmtpAddress, Recipient.Address, ...) - automatically approve' }
    @{ Name = 'PromptOOMAddressBookAccess';          Data = 2; Why = 'accessing an address book - automatically approve' }
    @{ Name = 'PromptOOMSend';                       Data = 2; Why = 'sending items - automatically approve' }
    @{ Name = 'PromptOOMMeetingTaskRequestResponse'; Data = 2; Why = 'responding to meeting and task requests - automatically approve' }
    @{ Name = 'PromptOOMSaveAs';                     Data = 2; Why = 'the Save As command - automatically approve' }
    @{ Name = 'PromptOOMFormulaAccess';              Data = 2; Why = 'UserProperty.Formula - automatically approve' }
    @{ Name = 'PromptOOMAddressUserPropertyFind';    Data = 2; Why = 'address information through UserProperties.Find - automatically approve' }
    @{ Name = 'PromptOOMCustomAction';               Data = 2; Why = 'object model custom actions - automatically approve' }
)

# The documented data for a PromptOOM* value [MS-DOC 2]. Anything else is not a setting.
$PromptMeaning = @{ 0 = 'Automatically deny'; 1 = 'Prompt user (the default)'; 2 = 'Automatically approve' }

$VerdictOk        = 'NO-PROMPT'
$VerdictPrompted  = 'PROMPTED'
$VerdictTimedOut  = 'TIMED-OUT'
$VerdictNotHere   = 'NOT-VERIFIABLE'

# =============================================================================================
# PURE DECISIONS. Each decides from its arguments alone, so -SelfTest can drive it anywhere.
# =============================================================================================

# Problems with what was read back, one string each. $Observed maps a value name to
# @{ Kind = <RegistryValueKind name, or $null when absent>; Data = <the data> }.
function Get-PolicyProblems {
    param([Parameter(Mandatory = $true)] [object[]] $Wanted, [Parameter(Mandatory = $true)] [hashtable] $Observed)
    $problems = @()
    foreach ($w in $Wanted) {
        $o = $Observed[$w.Name]
        if ($null -eq $o -or $null -eq $o.Kind) { $problems += "$($w.Name) is absent (wanted REG_DWORD $($w.Data))"; continue }
        if ([string]$o.Kind -cne 'DWord') { $problems += "$($w.Name) is a $($o.Kind), not a REG_DWORD"; continue }
        if ([int]$o.Data -ne [int]$w.Data) { $problems += "$($w.Name) = $($o.Data), wanted $($w.Data)" }
    }
    return ,$problems
}

# The verdict of -Verify, from what was observed. Order matters: a prompt is the finding even if
# the reads later returned (Deny makes them fail, or a pending prompt was answered elsewhere).
function Get-VerifyVerdict {
    param(
        [int]    $SessionId,
        [bool]   $OutlookRunning,
        [bool]   $AttachStartedAnother,
        [bool]   $PromptSeen,
        [bool]   $JobFinished,
        [int]    $AccountCount,
        [int]    $NonEmptySmtpCount,
        [string] $JobError
    )
    if ($SessionId -eq 0)         { return $VerdictNotHere }
    if (-not $OutlookRunning)     { return $VerdictNotHere }
    if ($PromptSeen)              { return $VerdictPrompted }
    if ($AttachStartedAnother)    { return $VerdictNotHere }
    if (-not $JobFinished)        { return $VerdictTimedOut }
    if ($JobError)                { return $VerdictNotHere }
    if ($AccountCount -lt 1)      { return $VerdictNotHere }
    if ($NonEmptySmtpCount -ne $AccountCount) { return $VerdictTimedOut }
    return $VerdictOk
}

# The OUTLOOK.EXE pids present after the reads that were not there before. MEASURED 2026-09-24 on
# OutlookAI-Unindexed: attaching from a second process to an Outlook that was started as a program
# (not by COM) makes COM launch a TRANSIENT 'OUTLOOK.EXE -Embedding', which hands the request to the
# running instance and exits within seconds - so this is judged only after that has had time to go.
function Get-LastingNewPids {
    param([int[]] $Before, [int[]] $After)
    return ,@($After | Where-Object { $Before -notcontains $_ } | Sort-Object)
}

# Is this the text of an Object Model Guard prompt? Its message always opens "A program is trying
# to ..." (access email address information / send an e-mail message on your behalf / ...).
# Ordinal StartsWith, never -like: a bracket in a -like pattern is a character class.
function Test-IsGuardPromptText {
    param([string] $Text)
    if (-not $Text) { return $false }
    return $Text.TrimStart().StartsWith('A program is trying to', [System.StringComparison]::Ordinal)
}

function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

# WSC productState, decoded the community-documented way - context only, never a verdict. The low
# byte is the signature state: 0x00 up to date, 0x10 out of date.
function Format-WscState {
    param([int] $ProductState)
    $sig = 'unknown'
    if (($ProductState -band 0xFF) -eq 0x00) { $sig = 'signatures up to date' }
    if (($ProductState -band 0xFF) -eq 0x10) { $sig = 'signatures OUT OF DATE' }
    return ('0x{0:X6} ({1})' -f $ProductState, $sig)
}

# =============================================================================================
# SELF-TEST. Pure: no registry, no COM, no guard, no process.
# =============================================================================================
function Invoke-SelfTest {
    $script:StChecks = 0
    $script:StFailures = @()
    function Check([string] $What, $Expected, $Actual) {
        $script:StChecks++
        $e = [string]$Expected; $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = $Expected -join ' | ' }
        if ($Actual -is [System.Array]) { $a = $Actual -join ' | ' }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else { $script:StFailures += "$What : expected [$e], got [$a]"; Write-Host "  FAIL $What - expected [$e], got [$a]" }
    }

    Write-Host '== the values are exactly the documented set =='
    Check 'nine values' 9 $PolicyValues.Count
    Check 'AdminSecurityMode is 3 (GPO settings)' 3 (($PolicyValues | Where-Object { $_.Name -ceq 'AdminSecurityMode' }).Data)
    $prompts = @($PolicyValues | Where-Object { $_.Name.StartsWith('PromptOOM', [System.StringComparison]::Ordinal) })
    Check 'eight PromptOOM* values' 8 $prompts.Count
    Check 'every PromptOOM* value is 2 = Automatically approve' 'Automatically approve' (@($prompts | ForEach-Object { $PromptMeaning[[int]$_.Data] } | Sort-Object -Unique) -join ',')
    Check 'no value name repeats' 9 (@($PolicyValues | ForEach-Object { $_.Name.ToLowerInvariant() } | Sort-Object -Unique).Count)
    Check 'no Simple MAPI value (KB 926512: not in the product)' 0 (@($PolicyValues | Where-Object { $_.Name.Contains('SimpleMAPI') }).Count)
    Check 'the key is the 16.0 per-user POLICY key' 'Software\Policies\Microsoft\Office\16.0\Outlook\Security' $PolicySubKey
    Check 'PromptOOMAddressInformationAccess is set (the measured trigger)' 1 (@($PolicyValues | Where-Object { $_.Name -ceq 'PromptOOMAddressInformationAccess' }).Count)
    Check 'PromptOOMSend is set (the product''s own send)' 1 (@($PolicyValues | Where-Object { $_.Name -ceq 'PromptOOMSend' }).Count)

    Write-Host '== reading the values back =='
    $good = @{}
    foreach ($v in $PolicyValues) { $good[$v.Name] = @{ Kind = 'DWord'; Data = $v.Data } }
    Check 'all nine present and right: no problem' 0 (Get-PolicyProblems -Wanted $PolicyValues -Observed $good).Count
    $missing = $good.Clone(); $missing.Remove('AdminSecurityMode')
    Check 'AdminSecurityMode absent is named' 'AdminSecurityMode is absent (wanted REG_DWORD 3)' (Get-PolicyProblems -Wanted $PolicyValues -Observed $missing)
    $wrongKind = $good.Clone(); $wrongKind['PromptOOMSend'] = @{ Kind = 'String'; Data = '2' }
    Check 'a REG_SZ "2" is not a DWORD' 'PromptOOMSend is a String, not a REG_DWORD' (Get-PolicyProblems -Wanted $PolicyValues -Observed $wrongKind)
    $wrongData = $good.Clone(); $wrongData['PromptOOMSaveAs'] = @{ Kind = 'DWord'; Data = 1 }
    Check 'Prompt user (1) is a problem' 'PromptOOMSaveAs = 1, wanted 2' (Get-PolicyProblems -Wanted $PolicyValues -Observed $wrongData)
    $absentKind = $good.Clone(); $absentKind['PromptOOMCustomAction'] = @{ Kind = $null; Data = $null }
    Check 'a null kind reads as absent' 'PromptOOMCustomAction is absent (wanted REG_DWORD 2)' (Get-PolicyProblems -Wanted $PolicyValues -Observed $absentKind)

    Write-Host '== the verdict =='
    Check 'all read, no prompt' $VerdictOk (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $false -JobFinished $true -AccountCount 2 -NonEmptySmtpCount 2 -JobError '')
    Check 'a prompt wins over reads that returned' $VerdictPrompted (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $true -JobFinished $true -AccountCount 1 -NonEmptySmtpCount 1 -JobError '')
    Check 'a prompt while still blocked' $VerdictPrompted (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $true -JobFinished $false -AccountCount 0 -NonEmptySmtpCount 0 -JobError '')
    Check 'no answer and no prompt' $VerdictTimedOut (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $false -JobFinished $false -AccountCount 0 -NonEmptySmtpCount 0 -JobError '')
    Check 'an empty SmtpAddress is not a pass' $VerdictTimedOut (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $false -JobFinished $true -AccountCount 2 -NonEmptySmtpCount 1 -JobError '')
    Check 'session 0 cannot verify' $VerdictNotHere (Get-VerifyVerdict -SessionId 0 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $false -JobFinished $true -AccountCount 1 -NonEmptySmtpCount 1 -JobError '')
    Check 'Outlook not running cannot verify' $VerdictNotHere (Get-VerifyVerdict -SessionId 1 -OutlookRunning $false -AttachStartedAnother $false -PromptSeen $false -JobFinished $false -AccountCount 0 -NonEmptySmtpCount 0 -JobError '')
    Check 'an attach that left a second Outlook running cannot verify' $VerdictNotHere (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $true -PromptSeen $false -JobFinished $true -AccountCount 1 -NonEmptySmtpCount 1 -JobError '')
    Check 'a prompt is still the finding when a second Outlook stayed' $VerdictPrompted (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $true -PromptSeen $true -JobFinished $true -AccountCount 1 -NonEmptySmtpCount 1 -JobError '')
    Check 'a transient Outlook that went away again is not a second one' @() (Get-LastingNewPids -Before @(1468) -After @(1468))
    Check 'a pid that stayed is named' @('524') (Get-LastingNewPids -Before @(1468) -After @(524, 1468))
    Check 'a profile with no account cannot verify' $VerdictNotHere (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $false -JobFinished $true -AccountCount 0 -NonEmptySmtpCount 0 -JobError '')
    Check 'a COM error cannot verify' $VerdictNotHere (Get-VerifyVerdict -SessionId 1 -OutlookRunning $true -AttachStartedAnother $false -PromptSeen $false -JobFinished $true -AccountCount 1 -NonEmptySmtpCount 0 -JobError 'boom')

    Write-Host '== recognising a guard prompt =='
    Check 'the address prompt' $true (Test-IsGuardPromptText 'A program is trying to access email address information stored in Outlook. If this is unexpected, click Deny')
    Check 'the send prompt' $true (Test-IsGuardPromptText 'A program is trying to send an e-mail message on your behalf.')
    Check 'the POP3 logon dialog is not one' $false (Test-IsGuardPromptText 'Enter your user name and password for the following server.')
    Check 'empty text is not one' $false (Test-IsGuardPromptText '')
    Check 'a bracketed text is not matched as a wildcard' $false (Test-IsGuardPromptText '[A program is trying to]')

    Write-Host '== the guard =='
    Check 'vmadmin on OAI-UNINDEXED passes' $true (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'OAI-UNINDEXED' -Users @('vmadmin') -Prefix 'OAI-')
    Check 'the maintainer workstation is refused' $false (Test-GuestIdentity -UserName 'jori' -ComputerName 'PC657' -Users @('vmadmin') -Prefix 'OAI-')
    Check 'vmadmin on a machine not named OAI-* is refused' $false (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'PC657' -Users @('vmadmin') -Prefix 'OAI-')
    Check 'an empty prefix refuses everything' $false (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'OAI-UNINDEXED' -Users @('vmadmin') -Prefix '')

    Write-Host '== WSC context decoding =='
    Check 'the guests state (0x061110)' '0x061110 (signatures OUT OF DATE)' (Format-WscState 397584)
    Check 'a current Defender (0x061100)' '0x061100 (signatures up to date)' (Format-WscState 397568)

    Write-Host '== this file =='
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$parseErrors)
    Check 'it parses' 0 @($parseErrors).Count
    $newItems = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-Item' }, $true))
    # New-Item -Force on an EXISTING registry key deletes every value under it - that is how CP-05
    # lost three first-run values (New-TierProfile.ps1 banner). This file creates the key with
    # RegistryKey.CreateSubKey, which opens an existing key and never empties it.
    Check 'no New-Item anywhere (CreateSubKey never empties a key)' 0 $newItems.Count
    $bytes = [System.IO.File]::ReadAllBytes($PSCommandPath)
    Check 'pure ASCII (Windows PowerShell 5.1 reads a BOM-less file as ANSI)' 0 (@($bytes | Where-Object { $_ -gt 127 }).Count)
    $kills = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and @('Stop-Process', 'taskkill', 'taskkill.exe') -contains $n.GetCommandName() }, $true))
    Check 'nothing here can kill a process (Outlook least of all)' 0 $kills.Count
    $quits = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -eq 'Quit' }, $true))
    Check 'nothing here quits Outlook' 0 $quits.Count

    Write-Host ''
    Write-Host "$($script:StChecks) checks, $($script:StFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE - only a guest can settle these (see the banner for what a guest settled):'
    Write-Host '  * that Outlook honours these values - -Verify on a guest is the proof, and a control run'
    Write-Host '    with the values absent is what shows the proof is not vacuous;'
    Write-Host '  * whether a RUNNING Outlook picks up a change - untested by design: -Execute refuses while'
    Write-Host '    Outlook runs, so the next start always comes after the write;'
    Write-Host '  * that the child job reached the Outlook under test - -Verify compares the OUTLOOK.EXE pid'
    Write-Host '    set before and 30 s after, and calls a second Outlook that stayed NOT-VERIFIABLE;'
    Write-Host '  * that Deny through WM_COMMAND dismisses a real guard prompt, and -Execute -Revert.'
    if ($script:StFailures.Count -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

# =============================================================================================
# THE GUARD, FIRST - before the plan is printed and before -Verify reads anything.
# =============================================================================================
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser
if (-not (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $ExpectedUser -Prefix $ExpectedComputerNamePrefix)) {
    throw "REFUSING TO RUN: the computer name '$env:COMPUTERNAME' does not start with '$ExpectedComputerNamePrefix'. The testbed guests are named OAI-* by Testbed/host/New-AnswerFile.ps1; this script disables Outlook's programmatic-access prompts and must never run anywhere else. If you named a guest differently, pass -ExpectedComputerNamePrefix. Do not widen the default."
}

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    if ($LogPath) {
        try {
            $dir = Split-Path -Parent $LogPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }
        catch { }
    }
}

# =============================================================================================
# READING THE MACHINE
# =============================================================================================
function Read-PolicyValues {
    $observed = @{}
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($PolicySubKey, $false)
    try {
        foreach ($v in $PolicyValues) {
            if ($null -eq $key) { $observed[$v.Name] = @{ Kind = $null; Data = $null }; continue }
            $names = @($key.GetValueNames())
            if ($names -notcontains $v.Name) { $observed[$v.Name] = @{ Kind = $null; Data = $null }; continue }
            $observed[$v.Name] = @{ Kind = [string]$key.GetValueKind($v.Name); Data = $key.GetValue($v.Name) }
        }
    }
    finally { if ($null -ne $key) { $key.Close() } }
    return $observed
}

function Show-PolicyState {
    param([hashtable] $Observed)
    foreach ($v in $PolicyValues) {
        $o = $Observed[$v.Name]
        if ($null -eq $o.Kind) { Say ("    {0,-36} absent" -f $v.Name) }
        else { Say ("    {0,-36} {1} {2}" -f $v.Name, $o.Kind, $o.Data) }
    }
}

function Get-WscSummary {
    try {
        $av = @(Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName 'AntivirusProduct' -ErrorAction Stop)
        if ($av.Count -eq 0) { return 'no antivirus product registered with Windows Security Center' }
        return (@($av | ForEach-Object { "$($_.displayName) " + (Format-WscState ([int]$_.productState)) }) -join '; ')
    }
    catch { return "Windows Security Center could not be read: $($_.Exception.Message)" }
}

# Visible #32770 dialogs owned by OUTLOOK.EXE in THIS session that are guard prompts: handle, the
# Deny button's handle and control id, and the message text.
function Find-GuardPrompt {
    if (-not ('OutlookAITestbed.GuardWin' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
namespace OutlookAITestbed {
  public class GuardDialog { public IntPtr Dialog; public IntPtr Deny; public int DenyId; public uint Pid; public string Text; }
  public static class GuardWin {
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr w, StringBuilder l, uint flags, uint timeout, out IntPtr result);
    static string Text(IntPtr h) { var sb = new StringBuilder(2048); IntPtr r; SendMessageTimeout(h, 0x000D, (IntPtr)2048, sb, 0x0002, 2000, out r); return sb.ToString(); }
    static string Cls(IntPtr h) { var c = new StringBuilder(256); GetClassName(h, c, 256); return c.ToString(); }
    public static List<GuardDialog> Dialogs(uint[] pids) {
      var found = new List<GuardDialog>();
      var wanted = new HashSet<uint>(pids);
      EnumWindows(delegate (IntPtr h, IntPtr l) {
        if (!IsWindowVisible(h) || Cls(h) != "#32770") return true;
        uint pid; GetWindowThreadProcessId(h, out pid);
        if (!wanted.Contains(pid)) return true;
        var d = new GuardDialog { Dialog = h, Pid = pid, Text = "" };
        EnumChildWindows(h, delegate (IntPtr k, IntPtr l2) {
          string c = Cls(k); string t = Text(k);
          if (c == "Button" && t.Replace("&", "") == "Deny") { d.Deny = k; d.DenyId = GetDlgCtrlID(k); }
          if (c == "Static" && t.Length > d.Text.Length) d.Text = t;
          return true;
        }, IntPtr.Zero);
        found.Add(d);
        return true;
      }, IntPtr.Zero);
      return found;
    }
  }
}
'@
    }
    $pids = [uint32[]]@(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | ForEach-Object { [uint32]$_.Id })
    if ($pids.Count -eq 0) { return @() }
    return @([OutlookAITestbed.GuardWin]::Dialogs($pids) | Where-Object { Test-IsGuardPromptText $_.Text })
}

# Deny grants nothing. The dialog ignores a BM_CLICK while it is not the active window (measured on
# CP-05's orphaned prompt, 2026-09-24), so this sends the WM_COMMAND its own Deny button would send
# - measured to close the prompt twice the same day, on that orphaned prompt and in this script's
# own control run.
function Deny-GuardPrompt {
    param($Prompt)
    if ($Prompt.Deny -eq [IntPtr]::Zero) { return $false }
    [void][OutlookAITestbed.GuardWin]::PostMessage($Prompt.Dialog, 0x0111, [IntPtr]$Prompt.DenyId, $Prompt.Deny)
    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt 20 -and [OutlookAITestbed.GuardWin]::IsWindow($Prompt.Dialog)) { Start-Sleep -Milliseconds 250 }
    return (-not [OutlookAITestbed.GuardWin]::IsWindow($Prompt.Dialog))
}

# =============================================================================================
# MODES
# =============================================================================================
function Show-Plan {
    Say 'PLAN. Dry run: nothing has been read or written.'
    Say "  key: HKCU\$PolicySubKey   (REG_DWORD values; the key is created only if it is missing, and never emptied)"
    foreach ($v in $PolicyValues) { Say ("    {0,-36} = {1}   {2}" -f $v.Name, $v.Data, $v.Why) }
    Say '  -Execute writes them (Outlook must be closed); -Verify proves them over COM in session 1.'
}

function Invoke-Execute {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw "REFUSING: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')). Outlook reads its security policy when it starts, so the start that follows this write must come after it. Quit Outlook gracefully - never taskkill it (mailbox-safety rule 7) - and run this again."
    }
    Say ("== Execute{0} ==" -f $(if ($Revert) { ' -Revert' } else { '' }))
    Say "  before:"
    Show-PolicyState -Observed (Read-PolicyValues)

    if ($Revert) {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($PolicySubKey, $true)
        if ($null -eq $key) { Say '  the key does not exist; nothing to remove.' }
        else {
            try {
                foreach ($v in $PolicyValues) {
                    if (@($key.GetValueNames()) -contains $v.Name) { $key.DeleteValue($v.Name); Say "  removed $($v.Name)" }
                }
            }
            finally { $key.Close() }
        }
        $after = Read-PolicyValues
        $left = @($PolicyValues | Where-Object { $null -ne $after[$_.Name].Kind })
        Say "  after:"
        Show-PolicyState -Observed $after
        if ($left.Count -gt 0) { Say "FAILED: $($left.Count) value(s) are still present."; exit 1 }
        Say 'REVERTED. Outlook falls back to its default Object Model Guard behaviour at its next start.'
        exit 0
    }

    # CreateSubKey opens the key when it exists and creates it when it does not. It never empties
    # an existing key - unlike New-Item -Force, which is how CP-05 lost three values.
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($PolicySubKey)
    try {
        foreach ($v in $PolicyValues) {
            $key.SetValue($v.Name, [int]$v.Data, [Microsoft.Win32.RegistryValueKind]::DWord)
            Say ("  set {0} = {1}" -f $v.Name, $v.Data)
        }
    }
    finally { $key.Close() }

    $after = Read-PolicyValues
    Say "  read back:"
    Show-PolicyState -Observed $after
    $problems = Get-PolicyProblems -Wanted $PolicyValues -Observed $after
    if ($problems.Count -gt 0) {
        foreach ($p in $problems) { Say "  FAIL $p" }
        exit 1
    }
    Say 'WRITTEN AND READ BACK. That proves the values exist, not that Outlook honours them. Start Outlook'
    Say 'on a profile with a mail account, in session 1, and run -Verify: that is the proof.'
    exit 0
}

function Invoke-Verify {
    $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    Say '== Verify =='
    Say "  antivirus (Windows Security Center, context): $(Get-WscSummary)"
    $observed = Read-PolicyValues
    Say '  policy values (context, not the verdict):'
    Show-PolicyState -Observed $observed
    $problems = Get-PolicyProblems -Wanted $PolicyValues -Observed $observed
    foreach ($p in $problems) { Say "  NOTE $p" }

    $pidsBefore = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Sort-Object)
    $running = ($pidsBefore.Count -gt 0)
    if ($sid -eq 0 -or -not $running) {
        $v = Get-VerifyVerdict -SessionId $sid -OutlookRunning $running -AttachStartedAnother $false -PromptSeen $false -JobFinished $false -AccountCount 0 -NonEmptySmtpCount 0 -JobError ''
        if ($sid -eq 0) { Say '  this is session 0: Outlook cannot be driven here. Run -Verify through Register-InteractiveTask.ps1.' }
        if (-not $running) { Say '  OUTLOOK.EXE is not running. This script never starts it: start it on a profile with a mail account first.' }
        Say "VERDICT: $v"
        exit 2
    }
    $openBefore = @(Find-GuardPrompt)
    if ($openBefore.Count -gt 0) {
        Say "  a guard prompt was ALREADY open before this run: '$($openBefore[0].Text)'"
        Say "VERDICT: $VerdictPrompted"
        Say '  It was left alone: this run did not raise it, so its caller may still be waiting on it.'
        exit 1
    }

    # The COM reads run in a child job, so a call blocked on a prompt blocks the child and not this
    # process, which keeps watching the session's windows.
    $reader = {
        $ErrorActionPreference = 'Stop'
        $r = [ordered]@{ PidsAfterAttach = @(); Profile = ''; Accounts = @(); Error = '' }
        $app = $null; $ns = $null; $accounts = $null; $held = @()
        try {
            $app = New-Object -ComObject Outlook.Application
            $r.PidsAfterAttach = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Sort-Object)
            $ns = $app.GetNamespace('MAPI')
            $r.Profile = [string]$ns.CurrentProfileName
            $accounts = $ns.Accounts
            for ($i = 1; $i -le $accounts.Count; $i++) {
                $a = $accounts.Item($i); $held += $a
                $t0 = [DateTime]::UtcNow
                $smtp = [string]$a.SmtpAddress      # THE protected read - Object Model Guard, address information
                $ms = [int]([DateTime]::UtcNow - $t0).TotalMilliseconds
                $r.Accounts += [pscustomobject]@{ DisplayName = [string]$a.DisplayName; SmtpAddress = $smtp; Ms = $ms }
            }
        }
        catch { $r.Error = $_.Exception.Message }
        finally {
            foreach ($x in (@($held) + @($accounts, $ns, $app))) { if ($null -ne $x) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) } catch { } } }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
        [pscustomobject]$r
    }

    Say ("  reading Account.SmtpAddress over COM from a child job; deadline {0} s; Outlook pid(s) {1}" -f $TimeoutSeconds, ($pidsBefore -join ','))
    $started = Get-Date
    $job = Start-Job -ScriptBlock $reader
    $prompt = $null
    $deadline = $started.AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (@('Completed', 'Failed', 'Stopped') -contains [string]$job.State) { break }
        $seen = @(Find-GuardPrompt)
        if ($seen.Count -gt 0) { $prompt = $seen[0]; break }
        Start-Sleep -Milliseconds 250
    }
    if ($null -ne $prompt) {
        Say ("  GUARD PROMPT after {0:N1} s: '{1}'" -f ((Get-Date) - $started).TotalSeconds, $prompt.Text)
        if (Deny-GuardPrompt -Prompt $prompt) { Say '  answered Deny (grants nothing); the prompt is gone.' }
        else { Say '  COULD NOT answer it - the prompt is STILL OPEN on the guest desktop.' }
        [void](Wait-Job -Job $job -Timeout 30)
    }
    $finished = (@('Completed', 'Failed', 'Stopped') -contains [string]$job.State)
    $result = $null
    if ($finished) { $result = Receive-Job -Job $job -ErrorAction SilentlyContinue | Select-Object -Last 1 }
    else { Stop-Job -Job $job }
    Remove-Job -Job $job -Force
    $elapsed = ((Get-Date) - $started).TotalSeconds

    # A transient OUTLOOK.EXE from the attach goes away within seconds (see Get-LastingNewPids).
    # One that is still there after this wait is a second Outlook, and nothing it answered counts.
    $lasting = @()
    $settle = (Get-Date).AddSeconds(30)
    do {
        $now = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Sort-Object)
        $lasting = Get-LastingNewPids -Before $pidsBefore -After $now
        if ($lasting.Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $settle)
    $another = ($lasting.Count -gt 0)
    if ($null -ne $result) {
        $transient = Get-LastingNewPids -Before $pidsBefore -After @($result.PidsAfterAttach)
        if ($transient.Count -gt 0 -and -not $another) { Say ("  the attach launched a transient OUTLOOK.EXE (pid {0}) that handed off to pid {1} and exited" -f ($transient -join ','), ($pidsBefore -join ',')) }
    }
    if ($another) { Say ("  a SECOND OUTLOOK.EXE is still running after 30 s (pid {0}) - the reads may not have come from the Outlook under test" -f ($lasting -join ',')) }

    $accountCount = 0; $nonEmpty = 0; $jobError = ''
    if ($null -ne $result) {
        $jobError = [string]$result.Error
        Say ("  profile '{0}', {1} account(s), {2:N1} s in all" -f $result.Profile, @($result.Accounts).Count, $elapsed)
        foreach ($a in @($result.Accounts)) {
            Say ("    account '{0}'  SmtpAddress='{1}'  ({2} ms)" -f $a.DisplayName, $a.SmtpAddress, $a.Ms)
            $accountCount++
            if ($a.SmtpAddress) { $nonEmpty++ }
        }
        if ($jobError) { Say "  the reader reported: $jobError" }
    }
    elseif (-not $finished) { Say ("  no answer within {0} s and no prompt seen; the child job was stopped (Outlook was not touched)." -f $TimeoutSeconds) }

    $verdict = Get-VerifyVerdict -SessionId $sid -OutlookRunning $true -AttachStartedAnother $another -PromptSeen ($null -ne $prompt) -JobFinished $finished -AccountCount $accountCount -NonEmptySmtpCount $nonEmpty -JobError $jobError
    Say "VERDICT: $verdict"
    if ($verdict -eq $VerdictOk) { exit 0 }
    if ($verdict -eq $VerdictNotHere) { exit 2 }
    exit 1
}

if ($Revert -and -not $Execute) { throw '-Revert only means something with -Execute.' }
if ($Execute -and $Verify) { throw 'Run -Execute and -Verify separately: Outlook has to start between them.' }
if ($Execute) { Invoke-Execute }
if ($Verify) { Invoke-Verify }
Show-Plan
