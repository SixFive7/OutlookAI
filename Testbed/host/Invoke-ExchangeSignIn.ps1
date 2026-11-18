#Requires -Version 5.1
<#
.SYNOPSIS
    Signs the Exchange test VM's Outlook in to a Microsoft 365 mailbox - a new profile's first
    account, a second account added to that profile, or a re-authentication - typing the password
    and, where Microsoft asks and a TOTP secret exists, a fresh code into the guest through its
    synthetic keyboard, so that no guest process or file ever holds either.

.DESCRIPTION
    RUN ON THE HOST, as the Windows user who stored the credential (DPAPI, CurrentUser), with a lease
    on the VM (Testbed/host/Set-TestbedLease.ps1). Windows PowerShell 5.1 or PowerShell 7.

    WHAT IT DOES, step by step, each one a call into the guest's Testbed/guest/Connect-ExchangeAccount.ps1
    in session 1 at -RunLevel Limited (the level the Outlook it drives runs at):

      -Mode NewProfile (Outlook's first start, no profile yet; -StartOutlook starts it, NOT elevated):
        1. Address   the "Email Account Setup" box gets the focus; the HOST types the address, Enter.
        2. Password  the Microsoft sign-in page's password field (IsPassword) gets the focus and is held;
                     the host types the password, then Enter.
        3. Code      only if the sign-in asks: the code field gets the focus; the host makes a TOTP code
                     with at least 12 s left on it and types it, then Enter.
        4. Choose "No, this app only" when Windows offers to sign in to all apps and register the
           device with the organisation - NEVER "Yes": that would add a device to the maintainer's
           tenant. Measured 2026-10-03: dsregcmd then reads AzureAdJoined NO, WorkplaceJoined NO.
        5. Clear Outlook's "Set up Outlook Mobile on my phone, too", then "Done".
        6. Verify    over COM, in session 1: an Exchange account with this address, cached.
      -Mode AddAccount -Account <address> (a SECOND mailbox in the existing profile, 2026-10-08): step 0
        asks the running Outlook for its own "Add Account..." command (ExecuteMso('AddAccount')),
        which opens the same "Email Account Setup" dialog as its first start; then steps 1 to 6 as
        for a new profile, with the account's own credential file.
      -Mode Reauth (an existing profile whose token has lapsed - after a checkpoint restore, say): the
        same steps from 2, for whatever sign-in page Outlook has raised.

    STEP 3 READS THE PAGE BEFORE IT TYPES ANYTHING (2026-10-08). After the password, the guest's
    -Step Page says which page Microsoft shows next. A code page is answered only for an account whose
    credential file holds a TOTP secret. ANY OTHER SECOND-FACTOR PAGE - "More information required"
    (MFA registration), "Approve sign in request", "Verify your identity" - STOPS the run with exit 5
    and the page's own words printed, so that the maintainer can do what only he can: nothing is
    clicked or skipped on his behalf, "Skip for now" included. A code page for an account WITHOUT a
    TOTP secret stops the same way.

    THE SECRETS. Testbed/host/Get-ExchangeCredential.ps1 decrypts them in memory at the moment each is
    typed; the password leaves it as a SecureString and becomes a string only for the one call that
    types it; the TOTP secret never leaves it - only a code does. Neither is printed, logged or passed
    to the guest any other way than as keystrokes: not in a script block, not as an argument, not in a
    file. Hyper-V's Msvm_Keyboard.TypeText delivers them to whatever has the keyboard focus on the
    guest's console - which is why the guest step must report READY, with the focus verified on the
    right field of the right window and HELD, before a single key is typed, and why a LOST-FOCUS, a
    NOT-FOUND or a TIMEOUT stops the run with nothing typed.

    MEASURED 2026-10-03 on OutlookAI-Exchange (Testbed/README.md section 4e): the whole of steps 1-5 in
    about three minutes, the password accepted without an MFA prompt that time, the account added in
    cached mode.

    WHY ONE PowerShell Direct SESSION. A second PowerShell Direct connection opened while one is being
    set up failed with "The credential is invalid" - and each such failure is a failed logon on the
    guest, whose lockout threshold is 10 in 10 minutes (measured: five failures from two concurrent
    attempts). So every guest call here goes through one session, sequentially, and the step that has
    to stay alive while the host types runs as a background job INSIDE the guest.

.PARAMETER VMName
    MANDATORY. Must be the VM testbed.json's exchangeVm block names - the script refuses any other:
    typing a mailbox password into a VM nobody meant is the mistake it exists to make impossible.

.PARAMETER Account
    The mailbox to sign in: the role testbed.json's exchangeVm.accounts gives it - hub or bystander - or
    its address. The script refuses any other. The record names no address (Q189 (b)): each account's is
    the "account" field of its own gitignored credential file, which Get-ExchangeCredential.ps1
    -AddressOf reads at run time. Mandatory with -Mode AddAccount; elsewhere the default is the VM's
    first account, the bystander, and -Mode Preflight checks EVERY listed account.

.PARAMETER Mode
    NewProfile, AddAccount, Reauth or Preflight. Preflight types nothing: before every live run it brings Defender's
    signatures up to date and waits for Security Center to say so - while they are stale, Outlook's
    Object Model Guard puts an Allow/Deny prompt in front of every out-of-process address read, and a
    run hangs on it (measured 2026-10-03: six stacked prompts, after a checkpoint restore brought back
    year-old signatures) - and then reads whether the account is still signed in. Exit 0 when both
    hold; when only the sign-in does not, run -Mode Reauth - unless it says OLD-OST: an account
    Outlook holds disconnected (ExchangeConnectionMode 400) since it refused the account's .ost as an
    old copy of the mailbox - Office's alert, OAlerts event 300, since the running Outlook started -
    has no lapsed sign-in, Outlook raises no sign-in page for it, and Reauth cannot fix it. Then the
    .ost files go aside: the live runner's -Checkpoint path, or guest/Reset-OutlookOfflineStore.ps1.
    Measured 2026-11-18: the VM restored to a checkpoint outside the runner, its .ost files left in
    place, held the bystander at 400 from Outlook's next start, and a Reauth found no page.

    EXIT CODES: 0 signed in (Preflight: everything holds); 1 failed; 4 refused; 5 STOPPED FOR THE
    MAINTAINER - Microsoft asked for a second factor, or for one to be registered, that only he can
    give (the page's words are printed).

.PARAMETER Mode (OwaSignIn, OwaSignOut)
    OUTLOOK ON THE WEB, for measuring roaming signatures (Q161 (c), decided by the maintainer 2026-11-17).
    OwaSignIn opens Outlook on the web in an InPrivate Edge window on the guest, types the address, the
    password and - if asked - a TOTP code exactly as the other modes do, answers "Stay signed in?" with
    No, and waits for the mail page. OwaSignOut signs out through Outlook on the web's own sign-out and
    closes every InPrivate window. THE TEST ACCOUNT ONLY: both refuse any account testbed.json does not
    record with the role "hub" - never the VM's real mailbox. The VM is restored to its checkpoint
    afterwards by whoever ran it (AGENTS.md, Q161 (c)) - and its .ost files set aside, as after every
    restore outside the runner: a restored .ost is older than the server, and Outlook refuses it.

.PARAMETER DefenderOnly
    With -Mode Preflight: bring Defender's signatures up to date and stop there - for the moment
    before Outlook starts, so that the Outlook which then starts finds them current (an Outlook started
    on stale signatures prompted on its first address read on 2026-10-08, although they were current a
    minute later).

.PARAMETER StartOutlook
    With -Mode NewProfile: start OUTLOOK.EXE in session 1, NOT elevated, before step 1.

.PARAMETER RepoRoot
    The main checkout: the guest credential and the Exchange credential are gitignored and live there.

.PARAMETER SelfTest
    The pure parts - reading a step's status file, the decisions taken from it, the VM-name guard.
    No VM, no credential.
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')] [string] $VMName,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('NewProfile', 'AddAccount', 'Reauth', 'Preflight', 'OwaSignIn', 'OwaSignOut')] [string] $Mode = 'NewProfile',
    [Parameter(ParameterSetName = 'Run')] [string] $Account,
    [Parameter(ParameterSetName = 'Run')] [switch] $StartOutlook,
    [Parameter(ParameterSetName = 'Run')] [switch] $DefenderOnly,
    [Parameter(ParameterSetName = 'Run')] [string] $RepoRoot,
    [Parameter(ParameterSetName = 'Run')] [int] $PageWaitSeconds = 90,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)
if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) }
$ErrorActionPreference = 'Stop'

# =============================================================================================
# PURE DECISIONS
# =============================================================================================

# A status file as Connect-ExchangeAccount.ps1 writes it: 'STATE <word>', 'AT <time>', detail lines.
function ConvertFrom-StepStatus {
    param([string[]] $Lines)
    $all = @($Lines | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0 -or $all[0] -notmatch '^STATE (\S+)$') {
        return [pscustomobject]@{ State = 'NONE'; Detail = @() }
    }
    $state = $Matches[1]
    $detail = @()
    if ($all.Count -gt 2) { $detail = @($all[2..($all.Count - 1)]) }
    return [pscustomobject]@{ State = $state; Detail = $detail }
}

# Whether the host may type now: only on READY, and only when the focus line the guest wrote says
# the focus is where this step needs it - a password field for the password.
function Test-MayType {
    param($Status, [string] $Step)
    if ($null -eq $Status -or $Status.State -ne 'READY') { return $false }
    $focus = @($Status.Detail | Where-Object { $_ -like 'FOCUS *' })
    if ($focus.Count -ne 1) { return $false }
    if ($Step -eq 'Password') { return ($focus[0].Contains(' password=True ')) }
    return ($focus[0].Contains(' password=False '))
}

# The VM this may type into: exactly the one testbed.json's exchangeVm names.
function Get-ExchangeVmProblem {
    param([string] $VMName, $Testbed)
    $record = $null
    if ($null -ne $Testbed -and $Testbed.PSObject.Properties.Name -contains 'exchangeVm') { $record = $Testbed.exchangeVm }
    if ($null -eq $record -or -not ($record.vmName -is [string]) -or -not $record.vmName) {
        return 'REFUSING: Testbed/testbed.json records no exchangeVm, so no VM is the Exchange test VM.'
    }
    if ($VMName -cne $record.vmName) {
        return "REFUSING: '$VMName' is not the Exchange test VM ('$($record.vmName)' in Testbed/testbed.json). This script types a mailbox password into the VM it is given."
    }
    return $null
}

# The accounts this may type a password for: exactly the ones testbed.json's exchangeVm.accounts lists.
# Q189 (b): the record names each by its role and its credential file, never by its address - the
# address is that gitignored file's "account" field, which $AddressOf reads at run time
# (Get-ExchangeCredential.ps1 -AddressOf). In the record's order: the first is the default.
function Get-ExchangeAccounts {
    param($Testbed, [scriptblock] $AddressOf)
    if ($null -eq $Testbed -or -not ($Testbed.PSObject.Properties.Name -contains 'exchangeVm')) { return @() }
    $vm = $Testbed.exchangeVm
    if ($null -eq $vm -or -not ($vm.PSObject.Properties.Name -contains 'accounts')) { return @() }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($a in @($vm.accounts)) {
        if ($null -eq $a -or -not ($a.credentialFile -is [string]) -or -not $a.credentialFile) { continue }
        $address = [string](& $AddressOf ([string]$a.credentialFile))
        $list.Add([pscustomobject]@{ Role = [string]$a.role; CredentialFile = [string]$a.credentialFile; Address = $address })
    }
    return $list.ToArray()
}
# -Account as a role or an address, compared ignoring case (an address is case-insensitive). Returns the
# address as its credential file spells it, and its role - or the refusal, which names roles, never addresses.
function Resolve-ExchangeAccount {
    param([string] $Account, [object[]] $Known)
    $Known = @($Known | Where-Object { $null -ne $_ })
    if ($Known.Count -eq 0) { return [pscustomobject]@{ Account = $null; Role = $null; Problem = 'REFUSING: Testbed/testbed.json records no exchangeVm.accounts, so no mailbox is this VM''s to sign in.' } }
    if ([string]::IsNullOrWhiteSpace($Account)) { return [pscustomobject]@{ Account = $Known[0].Address; Role = $Known[0].Role; Problem = $null } }
    $given = $Account.Trim()
    foreach ($k in $Known) {
        if ([string]::Equals([string]$k.Role, $given, [StringComparison]::OrdinalIgnoreCase) -or [string]::Equals([string]$k.Address, $given, [StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{ Account = $k.Address; Role = $k.Role; Problem = $null }
        }
    }
    return [pscustomobject]@{ Account = $null; Role = $null; Problem = "REFUSING: '$Account' is neither the role nor the address of an account testbed.json's exchangeVm.accounts records (roles: $(@($Known | ForEach-Object { $_.Role }) -join ', ')). This script types a mailbox password; it types one only for a mailbox the record names." }
}

# Q161 (c): Outlook on the web signs in the TEST account only - the one testbed.json records as the hub.
function Get-OwaAccountProblem {
    param([string] $Account, $Resolved)
    if ([string]::IsNullOrWhiteSpace($Account)) { return 'REFUSING: Outlook on the web needs -Account hub (or the test account''s address), and only the test account may be given (Q161 (c)).' }
    if ($null -ne $Resolved -and $null -eq $Resolved.Problem -and ([string]$Resolved.Role) -ceq 'hub') { return $null }
    return "REFUSING: Outlook on the web signs in the test account only (Q161 (c)); '$Account' is not the account testbed.json records as the hub."
}

# What a -Step Page status says: the page kind and its texts.
function Get-PageKind {
    param($Status)
    if ($null -eq $Status -or $Status.State -ne 'DONE') { return 'UNREAD' }
    $line = @($Status.Detail | Where-Object { $_ -like 'PAGE *' } | Select-Object -First 1)
    if ($line.Count -eq 0) { return 'UNREAD' }
    return ($line[0].Substring(5)).Trim()
}

# What step 3 does with the page it read: TYPE a code, CONTINUE, STOP-MAINTAINER (exit 5) or FAIL.
function Get-AfterPasswordAction {
    param([string] $Kind, [bool] $HasTotp)
    switch ($Kind) {
        'CODE' { if ($HasTotp) { return 'TYPE' } else { return 'STOP-MAINTAINER' } }
        'MFA' { return 'STOP-MAINTAINER' }
        'PASSWORD' { return 'FAIL' }
        'ERROR' { return 'FAIL' }
        default { return 'CONTINUE' }
    }
}

# Whether an Office alert is the one Outlook raises when it refuses a data file as older than the mailbox on
# the server: OAlerts event 300, recorded on OutlookAI-Exchange on 2026-11-18 as "Outlook is using an old copy
# of your Outlook data file (.ost). Exit Outlook, delete the file, and restart Outlook." English Office only: a
# translated alert does not match, and the advice falls back to the sign-in's. Pure.
function Test-OldCopyAlert {
    param([string] $Text)
    return ([string]$Text) -match '(?i)\bold copy of your Outlook data file\b'
}

# The ExchangeConnectionMode a Verify status names - 'account <a>: type 0, ExchangeConnectionMode 400, ...' -
# or $null when it names none. Pure.
function Get-VerifyConnectionMode {
    param([string[]] $Detail)
    foreach ($line in @($Detail)) {
        $m = [regex]::Match([string]$line, 'ExchangeConnectionMode (\d+)')
        if ($m.Success) { return [int]$m.Groups[1].Value }
    }
    return $null
}

# What preflight 2 says of an account that is not signed in. OLD-OST when Outlook holds it disconnected
# (ExchangeConnectionMode 400) and an old-copy alert (Test-OldCopyAlert) came since the running Outlook started:
# that is no lapsed sign-in, Outlook raises no sign-in page for it, and -Mode Reauth cannot fix it (measured on
# OutlookAI-Exchange on 2026-11-18). Otherwise the sign-in's advice. Names the alert's time, never its text.
# Pure.
function Get-NotSignedInAdvice {
    param([string] $Role, [string[]] $Detail, [object[]] $Alerts, [string] $Checkpoint)
    $old = @(@($Alerts) | Where-Object { $null -ne $_ -and (Test-OldCopyAlert ([string]$_.Text)) })
    if ((Get-VerifyConnectionMode $Detail) -eq 400 -and $old.Count -gt 0) {
        if (-not $Checkpoint) { $Checkpoint = '<its newest checkpoint>' }
        return "OLD-OST: Outlook refused a data file (.ost) as an old copy of its mailbox at $($old[0].Utc) " `
            + "(OAlerts event 300, since this Outlook started), and holds the $Role disconnected (ExchangeConnectionMode 400). " `
            + 'No sign-in has lapsed, and -Mode Reauth cannot fix it: Outlook raises no sign-in for it. ' `
            + "Run the live tier with -Checkpoint $Checkpoint (Testbed/host/Invoke-LiveTierOnGuest.ps1), which sets both .ost files " `
            + 'aside before Outlook starts, or set them aside with guest/Reset-OutlookOfflineStore.ps1 -Execute -ResetSearchIndex ' `
            + 'once host/Restart-Guest.ps1 has quit Outlook.'
    }
    return "not signed in - run this script with -Mode Reauth -Account $Role"
}

if ($SelfTest) {
    Write-Host "Invoke-ExchangeSignIn.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No VM, no credential."
    $script:fails = 0
    function Check([string] $Name, $Expected, $Actual) {
        if ("$Expected" -ceq "$Actual") { Write-Host "  PASS  $Name" } else { $script:fails++; Write-Host "  FAIL  $Name - expected '$Expected', got '$Actual'" }
    }
    $ready = ConvertFrom-StepStatus @('STATE READY', 'AT 2026-10-03T19:54:03', 'window "" (explorer) handle=1', 'FOCUS Edit aid="i0118" class="" password=True process=Microsoft.AAD.BrokerPlugin topHandle=1 foreground=1')
    Check 'a status file is read' 'READY' $ready.State
    Check 'its detail lines are kept' 2 $ready.Detail.Count
    Check 'a password page held on a password field may be typed into' $true (Test-MayType $ready 'Password')
    Check 'but not as a code page' $false (Test-MayType $ready 'Code')
    $plain = ConvertFrom-StepStatus @('STATE READY', 'AT x', 'window "Email Account Setup" (OUTLOOK) handle=2', 'FOCUS Pane aid="1" class="RICHEDIT60W" password=False process=OUTLOOK topHandle=2 foreground=2')
    Check 'an address box may take the address' $true (Test-MayType $plain 'Address')
    Check 'and never the password' $false (Test-MayType $plain 'Password')
    Check 'LOST-FOCUS types nothing' $false (Test-MayType (ConvertFrom-StepStatus @('STATE LOST-FOCUS', 'AT x', 'FOCUS Edit aid="i0118" class="" password=True')) 'Password')
    Check 'NOT-FOUND types nothing' $false (Test-MayType (ConvertFrom-StepStatus @('STATE NOT-FOUND', 'AT x')) 'Code')
    Check 'a READY with no focus line types nothing' $false (Test-MayType (ConvertFrom-StepStatus @('STATE READY', 'AT x')) 'Address')
    Check 'an empty file is no state' 'NONE' (ConvertFrom-StepStatus @()).State
    Check 'a garbled first line is no state' 'NONE' (ConvertFrom-StepStatus @('READY')).State
    $record = [pscustomobject]@{ exchangeVm = [pscustomobject]@{ vmName = 'OutlookAI-Exchange' } }
    Check 'the recorded VM is accepted' '' ([string](Get-ExchangeVmProblem 'OutlookAI-Exchange' $record))
    Check 'another VM is refused' $true ((Get-ExchangeVmProblem 'OutlookAI-Indexed' $record) -like 'REFUSING*')
    Check 'a case variant is refused' $true ((Get-ExchangeVmProblem 'outlookai-exchange' $record) -like 'REFUSING*')
    Check 'no record refuses' $true ((Get-ExchangeVmProblem 'OutlookAI-Exchange' ([pscustomobject]@{ vmName = 'x' })) -like 'REFUSING*')
    $real = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json') -Raw | ConvertFrom-Json
    Check 'and testbed.json records the Exchange VM this script is for' '' ([string](Get-ExchangeVmProblem 'OutlookAI-Exchange' $real))
    $fakeAddressOf = { param([string] $File) switch ($File) { 'exchange-credentials.json' { 'first@example.invalid' } 'exchange-credentials-two.json' { 'Second@example.invalid' } default { throw "no $File" } } }
    $two = [pscustomobject]@{ exchangeVm = [pscustomobject]@{ vmName = 'OutlookAI-Exchange'; accounts = @(
                [pscustomobject]@{ address = 'bystander@exchange-vm.invalid'; credentialFile = 'exchange-credentials.json'; role = 'bystander' },
                [pscustomobject]@{ address = 'hub@exchange-vm.invalid'; credentialFile = 'exchange-credentials-two.json'; role = 'hub' }) } }
    $known = @(Get-ExchangeAccounts $two $fakeAddressOf)
    Check 'each account''s address is read from its own credential file' 'first@example.invalid|Second@example.invalid' (@($known | ForEach-Object { $_.Address }) -join '|')
    Check 'no -Account means the first recorded account' 'first@example.invalid' (Resolve-ExchangeAccount '' $known).Account
    Check 'a recorded account is accepted by address, in its file''s spelling' 'Second@example.invalid' (Resolve-ExchangeAccount 'second@EXAMPLE.invalid' $known).Account
    Check 'and by its role' 'Second@example.invalid|hub' ((Resolve-ExchangeAccount 'HUB' $known) | ForEach-Object { "$($_.Account)|$($_.Role)" })
    Check 'an unrecorded account is refused' $true ((Resolve-ExchangeAccount 'third@example.invalid' $known).Problem -like 'REFUSING*')
    Check 'and the refusal lists roles, never a recorded address' $false ((Resolve-ExchangeAccount 'third@example.invalid' $known).Problem -like '*first@*')
    Check 'no accounts on record refuses' $true ((Resolve-ExchangeAccount 'first@example.invalid' @(Get-ExchangeAccounts $record $fakeAddressOf)).Problem -like 'REFUSING*')
    $realAccounts = @($real.exchangeVm.accounts)
    Check 'testbed.json records the bystander first and the hub' 'bystander|hub' (@($realAccounts | ForEach-Object { [string]$_.role }) -join '|')
    Check 'and names each by a placeholder, never by an address (Q189 (b))' 'bystander@exchange-vm.invalid|hub@exchange-vm.invalid' (@($realAccounts | ForEach-Object { [string]$_.address }) -join '|')
    Check 'each with a credential file -AddressOf will read' 2 @($realAccounts | Where-Object { [string]$_.credentialFile -cmatch '^exchange-credentials(-[A-Za-z0-9-]+)?\.json$' }).Count
    Check 'a page status is read' 'MFA' (Get-PageKind (ConvertFrom-StepStatus @('STATE DONE', 'AT x', 'PAGE MFA', 'window "x"', 'TEXT More information required')))
    Check 'an unread page' 'UNREAD' (Get-PageKind (ConvertFrom-StepStatus @('STATE TIMEOUT', 'AT x')))
    Check 'an MFA page stops for the maintainer' 'STOP-MAINTAINER' (Get-AfterPasswordAction 'MFA' $true)
    Check 'a code page with a TOTP secret is typed' 'TYPE' (Get-AfterPasswordAction 'CODE' $true)
    Check 'a code page without one stops for the maintainer' 'STOP-MAINTAINER' (Get-AfterPasswordAction 'CODE' $false)
    Check 'the password page again fails' 'FAIL' (Get-AfterPasswordAction 'PASSWORD' $false)
    Check 'an error page fails' 'FAIL' (Get-AfterPasswordAction 'ERROR' $true)
    Check 'the device offer continues' 'CONTINUE' (Get-AfterPasswordAction 'DEVICE' $false)
    Check 'no page continues' 'CONTINUE' (Get-AfterPasswordAction 'NONE' $false)
    Check 'Outlook on the web signs in the hub' $true ($null -eq (Get-OwaAccountProblem 'hub' (Resolve-ExchangeAccount 'hub' $known)))
    Check 'given by its address too' $true ($null -eq (Get-OwaAccountProblem 'second@example.invalid' (Resolve-ExchangeAccount 'second@example.invalid' $known)))
    Check 'and never the bystander' $true ([string](Get-OwaAccountProblem 'first@example.invalid' (Resolve-ExchangeAccount 'first@example.invalid' $known))).StartsWith('REFUSING')
    Check 'nor by its role' $true ([string](Get-OwaAccountProblem 'bystander' (Resolve-ExchangeAccount 'bystander' $known))).StartsWith('REFUSING')
    Check 'nor an account given by nobody' $true ([string](Get-OwaAccountProblem '' (Resolve-ExchangeAccount '' $known))).StartsWith('REFUSING')
    # The two OAlerts entries OutlookAI-Exchange recorded at Outlook's start on 2026-11-18, word for word.
    $oldCopy = [pscustomobject]@{ Utc = '2026-11-18T07:41:14Z'; Text = 'Microsoft Outlook Outlook is using an old copy of your Outlook data file (.ost). Exit Outlook, delete the file, and restart Outlook. A new file will be automatically created the next time you initiate a send/receive. P1: 301737 P2: 16.0.17932.21000 P3: 4q0i4P4:' }
    $compositor = [pscustomobject]@{ Utc = '2026-11-18T07:41:08Z'; Text = 'Compositor Type: 1OUTLOOKP1: %3P2: %4P3: %5P4: %6' }
    Check 'the old-copy alert is told by its words' $true (Test-OldCopyAlert $oldCopy.Text)
    Check 'and another alert is not' $false (Test-OldCopyAlert $compositor.Text)
    $disconnected = @('account bystander@example.invalid: type 0, ExchangeConnectionMode 400, delivery store "bystander@example.invalid" cached=True, profile "Outlook"')
    Check 'a Verify status names its connection mode' 400 (Get-VerifyConnectionMode $disconnected)
    Check 'and one that names none, nothing' $true ($null -eq (Get-VerifyConnectionMode @('no Outlook to attach to')))
    $advice = Get-NotSignedInAdvice 'bystander' $disconnected @($compositor, $oldCopy) 'CP-10-SDK11'
    Check 'disconnected after an old-copy alert is OLD-OST, at the alert''s time' $true ($advice.StartsWith('OLD-OST: ') -and $advice.Contains(' at 2026-11-18T07:41:14Z '))
    Check 'and names the set-aside, not Reauth' $true ($advice.Contains('-Checkpoint CP-10-SDK11') -and $advice.Contains('Reset-OutlookOfflineStore.ps1') -and -not $advice.Contains('run this script with -Mode Reauth'))
    Check 'and no address' $false ($advice.Contains('@'))
    Check 'disconnected with no such alert is the sign-in''s' 'not signed in - run this script with -Mode Reauth -Account bystander' (Get-NotSignedInAdvice 'bystander' $disconnected @($compositor) 'CP-10-SDK11')
    Check 'an old-copy alert with the account in another mode is the sign-in''s' $true ((Get-NotSignedInAdvice 'hub' @('account hub@example.invalid: type 0, ExchangeConnectionMode 300') @($oldCopy) 'CP-10-SDK11').StartsWith('not signed in'))
    Check 'and an account Outlook could not be asked about too' $true ((Get-NotSignedInAdvice 'hub' @('no Outlook to attach to') @($oldCopy) '').StartsWith('not signed in'))
    Write-Host ''
    if ($script:fails -gt 0) { Write-Host "$script:fails FAILED"; exit 1 }
    Write-Host 'All passed.'
    exit 0
}

# =============================================================================================
# THE RUN
# =============================================================================================
# The record this script belongs to - its own checkout's, not -RepoRoot's, which is only where the
# gitignored credentials live (from a worktree, the main checkout).
$testbed = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json') -Raw | ConvertFrom-Json
$problem = Get-ExchangeVmProblem $VMName $testbed
if ($null -ne $problem) { Write-Host $problem; exit 4 }

# Say is one of the helpers the scripts share, dot-sourced from Testbed/guest/ScriptHelpers.ps1 (Q154,
# Q166 (b), decision 215 (a)), which says why.
. (Join-Path $PSScriptRoot '..\guest\ScriptHelpers.ps1')

if ($Mode -eq 'AddAccount' -and [string]::IsNullOrWhiteSpace($Account)) { Write-Host 'REFUSING: -Mode AddAccount needs -Account <address>: which mailbox to add is never a default.'; exit 4 }
$owa = ($Mode -eq 'OwaSignIn' -or $Mode -eq 'OwaSignOut')
# Q189 (b): every recorded account's address, read from its credential file in the main checkout.
$addressOf = { param([string] $File) & (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot -AddressOf $File }
try { $known = @(Get-ExchangeAccounts $testbed $addressOf) }
catch { Write-Host "REFUSING: an account's address could not be read from its credential file - $($_.Exception.Message)"; exit 4 }
$resolved = Resolve-ExchangeAccount $Account $known
if ($owa) { $owaProblem = Get-OwaAccountProblem $Account $resolved; if ($null -ne $owaProblem) { Write-Host $owaProblem; exit 4 } }
# Every guest step of the Outlook-on-the-web modes reads the InPrivate Edge window instead of Outlook's.
$GuestFlags = ''
if ($owa) { $GuestFlags = ' -Browser' }
if ($null -ne $resolved.Problem) { Write-Host $resolved.Problem; exit 4 }
$guestCred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$credential = & (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot -Account $resolved.Account
$account = $credential.Account
$hasTotp = [bool]$credential.HasTotp
$credential = $null

$guest = @(Get-CimInstance -Namespace 'root\virtualization\v2' -ClassName 'Msvm_ComputerSystem' -Filter "ElementName='$VMName'")
if ($guest.Count -ne 1) { Write-Host "REFUSING: expected one Msvm_ComputerSystem named '$VMName', found $($guest.Count)."; exit 4 }
$keyboard = @($guest[0] | Get-CimAssociatedInstance -ResultClassName 'Msvm_Keyboard')
if ($keyboard.Count -lt 1) { Write-Host "REFUSING: '$VMName' has no synthetic keyboard."; exit 4 }
$keyboard = $keyboard[0]

function Send-Text([string] $Text) {
    $r = Invoke-CimMethod -InputObject $keyboard -MethodName 'TypeText' -Arguments @{ asciiText = $Text }
    if ($r.ReturnValue -ne 0) { throw "the synthetic keyboard refused the text (TypeText returned $($r.ReturnValue))" }
}
function Send-Enter {
    $r = Invoke-CimMethod -InputObject $keyboard -MethodName 'TypeKey' -Arguments @{ keyCode = [uint32]0x0D }
    if ($r.ReturnValue -ne 0) { throw "the synthetic keyboard refused Enter (TypeKey returned $($r.ReturnValue))" }
}

$session = New-PSSession -VMName $VMName -Credential $guestCred

# A step that only reads or invokes: run it in session 1 and wait for it.
function Invoke-GuestStep([string] $Arguments, [int] $Timeout = 300) {
    $lines = Invoke-Command -Session $session -ScriptBlock {
        param($a, $t)
        Set-ExecutionPolicy -Scope Process Bypass -Force
        Set-Location C:\OutlookAI-Q5
        # ps51-native-stderr-ok: a PowerShell script on the guest, not a program, run inside an Invoke-Command block where the host's 'Stop' does not apply - the remote session's default is 'Continue'
        & .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds $t -Script "& 'C:\OutlookAI-Q5\Connect-ExchangeAccount.ps1' $a" *>&1 | Out-Null
        $step = ($a -split ' ')[1]
        Get-Content -LiteralPath "C:\OutlookAI-Q5\exchange\signin\$step.status.txt" -ErrorAction SilentlyContinue
    } -ArgumentList ($Arguments + $GuestFlags), $Timeout
    return (ConvertFrom-StepStatus @($lines))
}

# A step the host types into: the guest holds the focus in a background job while the host types.
function Invoke-TypedStep([string] $Step, [scriptblock] $GetText, [int] $WaitSeconds) {
    Invoke-Command -Session $session -ScriptBlock {
        param($s, $w, $g)
        Remove-Item "C:\OutlookAI-Q5\exchange\signin\$s.status.txt", "C:\OutlookAI-Q5\exchange\signin\$s.go" -Force -ErrorAction SilentlyContinue
        $global:typedStepJob = Start-Job -ScriptBlock {
            param($s, $w, $g)
            Set-ExecutionPolicy -Scope Process Bypass -Force
            Set-Location C:\OutlookAI-Q5
            # ps51-native-stderr-ok: a PowerShell script on the guest, not a program, run inside an Invoke-Command block where the host's 'Stop' does not apply - the remote session's default is 'Continue'
            & .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds ($w + 200) -Script "& 'C:\OutlookAI-Q5\Connect-ExchangeAccount.ps1' -Step $s -WaitSeconds $w -HoldSeconds 60$g" *>&1 | Out-Null
        } -ArgumentList $s, $w, $g
    } -ArgumentList $Step, $WaitSeconds, $GuestFlags

    $status = $null
    $deadline = (Get-Date).AddSeconds($WaitSeconds + 90)
    while ((Get-Date) -lt $deadline) {
        $lines = Invoke-Command -Session $session -ScriptBlock { param($s) $p = "C:\OutlookAI-Q5\exchange\signin\$s.status.txt"; if (Test-Path $p) { Get-Content $p } } -ArgumentList $Step
        $status = ConvertFrom-StepStatus @($lines)
        if ($status.State -ne 'NONE') { break }
        Start-Sleep -Milliseconds 500
    }

    $typed = $false
    if (Test-MayType $status $Step) {
        $text = & $GetText
        try { Send-Text $text } finally { $text = $null }
        Invoke-Command -Session $session -ScriptBlock { param($s) Set-Content -Path "C:\OutlookAI-Q5\exchange\signin\$s.go" -Value 'typed' } -ArgumentList $Step
        Start-Sleep -Seconds 2
        $after = ConvertFrom-StepStatus @(Invoke-Command -Session $session -ScriptBlock { param($s) Get-Content "C:\OutlookAI-Q5\exchange\signin\$s.status.txt" } -ArgumentList $Step)
        if ($after.State -eq 'DONE') { Send-Enter; $typed = $true }
        else { Say "  the focus did not hold while typing ($($after.State)) - Enter NOT pressed" }
    }
    Invoke-Command -Session $session -ScriptBlock { if ($global:typedStepJob) { $null = Wait-Job $global:typedStepJob -Timeout 240; Remove-Job $global:typedStepJob -Force } } | Out-Null
    return [pscustomobject]@{ Status = $status; Typed = $typed }
}

$exit = 1
try {
    # Every guest step is a scheduled task whose powershell.exe is started -WindowStyle Hidden - which
    # Windows Terminal, the default console host on this Windows 11 build, does not honour: measured
    # 2026-10-03, the task's console came up as a visible Windows Terminal window, and a visible window
    # can take the keyboard focus the host is about to type into. The documented per-user setting
    # (Settings > Privacy & security > For developers > Terminal: Windows Console Host) puts consoles
    # back in conhost, which does honour it. Set before anything else, every run; it is idempotent.
    Invoke-Command -Session $session -ScriptBlock {
        $key = 'HKCU:\Console\%%Startup'
        if (-not (Test-Path -LiteralPath $key)) { New-Item -Path $key -Force | Out-Null }
        Set-ItemProperty -LiteralPath $key -Name DelegationConsole -Value '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
        Set-ItemProperty -LiteralPath $key -Name DelegationTerminal -Value '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
    }

    if ($Mode -eq 'Preflight') {
        Say 'preflight 1: Defender signatures, so that Outlook''s Object Model Guard does not prompt'
        $av = Invoke-Command -Session $session -ScriptBlock {
            $updated = $true
            try { Update-MpSignature -ErrorAction Stop } catch { $updated = $false }
            $state = 0x10
            $deadline = (Get-Date).AddMinutes(3)
            do {
                $product = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | Where-Object { $_.displayName -eq 'Windows Defender' } | Select-Object -First 1
                if ($null -ne $product) { $state = [int]$product.productState }
                if (($state -band 0xFF) -eq 0) { break }
                Start-Sleep -Seconds 10
            } while ((Get-Date) -lt $deadline)
            # Defender's WMI provider can be briefly unavailable right after an update ("Provider load
            # failure", measured): the age is a detail, the Security Center state above is the answer.
            $age = -1
            try { $age = (Get-MpComputerStatus -ErrorAction Stop).AntivirusSignatureAge } catch { $age = -1 }
            [pscustomobject]@{ Updated = $updated; State = $state; Age = $age }
        }
        if (($av.State -band 0xFF) -ne 0) {
            Say ('  Security Center still reports the definitions out of date (0x{0:X6}, update ran: {1}) - a live run would meet the guard''s prompts' -f $av.State, $av.Updated)
            exit 1
        }
        Say ('  up to date (productState 0x{0:X6}); signatures {1} day(s) old' -f $av.State, $av.Age)
        if ($DefenderOnly) { exit 0 }
        $allSignedIn = $true
        $alerts = $null
        foreach ($recordedAccount in $known) {
            $recorded = $recordedAccount.Address
            Say "preflight 2: is the $($recordedAccount.Role) ($recorded) still signed in?"
            $v = Invoke-GuestStep "-Step Verify -Account '$recorded' -WaitSeconds 60"
            Say "  $($v.State) $($v.Detail -join ' | ')"
            if ($v.State -ne 'SIGNED-IN') {
                # Office's alerts since the running Outlook started, read once (session 0, read only): an old-copy
                # alert there means no sign-in will fix the account (Get-NotSignedInAdvice).
                if ($null -eq $alerts) {
                    $alerts = @(Invoke-Command -Session $session -ScriptBlock {
                            $outlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Sort-Object StartTime)
                            if ($outlook.Count -gt 0) {
                                Get-WinEvent -FilterHashtable @{ LogName = 'OAlerts'; Id = 300; StartTime = $outlook[0].StartTime } -ErrorAction SilentlyContinue | ForEach-Object {
                                    [pscustomobject]@{ Utc = $_.TimeCreated.ToUniversalTime().ToString('s') + 'Z'; Text = [string]$_.Message }
                                }
                            }
                        })
                }
                Say ('  ' + (Get-NotSignedInAdvice $recordedAccount.Role @($v.Detail) $alerts ([string]@($testbed.exchangeVm.checkpoints)[-1])))
                $allSignedIn = $false
            }
        }
        if (-not $allSignedIn) { exit 1 }
        exit 0
    }

    if ($Mode -eq 'OwaSignOut') {
        Say 'Outlook on the web: signing out and closing every InPrivate Edge window'
        $o = Invoke-GuestStep "-Step SignOutBrowser -WaitSeconds 60"
        Say "  $($o.State) $($o.Detail -join ' | ')"
        if ($o.State -eq 'DONE') { $exit = 0 }
        exit $exit
    }
    if ($Mode -eq 'OwaSignIn') {
        Say "step 0: Outlook on the web for $account in an InPrivate Edge window (Q161 (c))"
        $o = Invoke-GuestStep "-Step OpenBrowser -WaitSeconds 60"
        Say "  $($o.State) $($o.Detail -join ' | ')"
        if ($o.State -ne 'DONE') { exit 1 }
        Start-Sleep -Seconds 8
        Say "step 1: the address ($account)"
        $r = Invoke-TypedStep -Step 'Address' -GetText { $account } -WaitSeconds $PageWaitSeconds
        if (-not $r.Typed) { Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"; exit 1 }
        Start-Sleep -Seconds 3
    }

    if ($Mode -eq 'AddAccount') {
        Say "step 0: Outlook's own Add Account dialog (ExecuteMso AddAccount)"
        $o = Invoke-GuestStep "-Step OpenAddAccount -WaitSeconds 60"
        Say "  $($o.State) $($o.Detail -join ' | ')"
        if ($o.State -ne 'DONE') { exit 1 }
        Start-Sleep -Seconds 2
    }

    if ($Mode -eq 'NewProfile' -or $Mode -eq 'AddAccount') {
        if ($StartOutlook -and $Mode -eq 'NewProfile') {
            Say 'starting Outlook in session 1, NOT elevated (no profile yet: it opens its account setup)'
            Invoke-Command -Session $session -ScriptBlock {
                Set-ExecutionPolicy -Scope Process Bypass -Force
                Set-Location C:\OutlookAI-Q5
                # ps51-native-stderr-ok: a PowerShell script on the guest, not a program, run inside an Invoke-Command block where the host's 'Stop' does not apply - the remote session's default is 'Continue'
                & .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 120 -Script "Start-Process 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'" *>&1 | Out-Null
            }
            Start-Sleep -Seconds 20
        }
        Say "step 1: the address ($account)"
        $r = Invoke-TypedStep -Step 'Address' -GetText { $account } -WaitSeconds $PageWaitSeconds
        if (-not $r.Typed) { Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"; exit 1 }
    }

    Say 'step 2: the password (typed from the host; never shown)'
    $r = Invoke-TypedStep -Step 'Password' -WaitSeconds $PageWaitSeconds -GetText {
        $secure = (& (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot -Account $account).Password
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    if (-not $r.Typed) {
        Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"
        $page = Invoke-GuestStep "-Step Page -WaitSeconds 5"
        Say "  the page showing instead: $($page.Detail -join ' | ')"
        if ((Get-PageKind $page) -eq 'MFA') { Say '  STOPPED FOR THE MAINTAINER: Microsoft asks for a second factor before the password page (above).'; exit 5 }
        exit 1
    }

    Say 'step 3: what the sign-in asks for next (read, nothing typed)'
    $page = Invoke-GuestStep "-Step Page -WaitSeconds 20"
    $kind = Get-PageKind $page
    Say "  $kind - $($page.Detail -join ' | ')"
    switch (Get-AfterPasswordAction $kind $hasTotp) {
        'TYPE' {
            $r = Invoke-TypedStep -Step 'Code' -WaitSeconds 25 -GetText {
                (& (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot -Account $account -TotpCode -MinimumValiditySeconds 12).Code
            }
            if ($r.Typed) { Say '  a code was asked for and typed' }
            else { Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"; exit 1 }
        }
        'STOP-MAINTAINER' {
            if ($kind -eq 'CODE') { Say "  STOPPED FOR THE MAINTAINER: Microsoft asks $account for a verification code, and its credential file holds no TOTP secret - nothing was typed." }
            else { Say "  STOPPED FOR THE MAINTAINER: Microsoft asks $account for a second factor or its registration (the page's words above) - nothing was clicked, 'Skip for now' included." }
            exit 5
        }
        'FAIL' { Say "  stopped: the sign-in did not move past the password ($kind)"; exit 1 }
        default { Say '  no code asked for' }
    }

    if ($Mode -eq 'OwaSignIn') {
        Say 'step 4: "Stay signed in?" - No'
        $c = Invoke-GuestStep "-Step Choose -ControlName 'No' -WaitSeconds 40"
        Say "  $($c.State) $($c.Detail -join ' | ')"
        $v = Invoke-GuestStep "-Step Verify -WaitSeconds 120"
        Say "step 5: $($v.State) $($v.Detail -join ' | ')"
        if ($v.State -eq 'SIGNED-IN') { $exit = 0 }
        exit $exit
    }

    Say "step 4: decline signing in to all apps and registering the device"
    $c = Invoke-GuestStep "-Step Choose -ControlName 'No, this app only' -WaitSeconds 40"
    Say "  $($c.State) $($c.Detail -join ' | ')"

    if ($Mode -eq 'NewProfile' -or $Mode -eq 'AddAccount') {
        # Outlook now finds and configures the mailbox, a page or two of its own setup dialog. Read
        # until it says added or failed, each page it shows printed once (2026-10-08: the second
        # account's first attempt ended back on the address page, and nothing said why).
        Say "step 4b: Outlook's setup dialog, until it says the account is added (read, nothing clicked)"
        $outcome = 'SETUP-BUSY'
        $seenPages = @{}
        $setupDeadline = (Get-Date).AddSeconds(240)
        while ((Get-Date) -lt $setupDeadline) {
            $page = Invoke-GuestStep "-Step Page -WaitSeconds 10"
            $outcome = Get-PageKind $page
            $text = $page.Detail -join ' | '
            if (-not $seenPages.ContainsKey($text)) { $seenPages[$text] = $true; Say "  $outcome - $text" }
            if (@('SETUP-ADDED', 'SETUP-FAILED', 'SETUP-START', 'MFA', 'CODE', 'PASSWORD', 'ERROR') -contains $outcome) { break }
            Start-Sleep -Seconds 3
        }
        if ($outcome -eq 'MFA') { Say '  STOPPED FOR THE MAINTAINER: Microsoft asks for a second factor or its registration (above) - nothing was clicked.'; exit 5 }
        if ($outcome -ne 'SETUP-ADDED') { Say "  stopped: Outlook's setup did not report the account added ($outcome)"; exit 1 }

        Say "step 5: clear 'Set up Outlook Mobile on my phone, too', then Done"
        $c = Invoke-GuestStep "-Step Choose -ControlName 'Set up Outlook Mobile on my phone, too' -WaitSeconds 30"
        Say "  $($c.State) $($c.Detail -join ' | ')"
        # The box is offered with a profile's FIRST account; a second one's success page may not carry
        # it, and a box that is not there is a box that will not send anybody to a phone set-up page.
        if ($c.State -ne 'DONE' -and -not ($Mode -eq 'AddAccount' -and $c.State -eq 'NOT-FOUND')) { exit 1 }
        $c = Invoke-GuestStep "-Step Choose -ControlName 'Done' -WaitSeconds 30"
        Say "  $($c.State) $($c.Detail -join ' | ')"
        if ($c.State -ne 'DONE') { exit 1 }
        Start-Sleep -Seconds 20
    }

    Say 'step 6: verify over COM'
    $v = Invoke-GuestStep "-Step Verify -Account '$account' -WaitSeconds 120"
    Say "  $($v.State) $($v.Detail -join ' | ')"
    if ($v.State -eq 'SIGNED-IN') { $exit = 0 }
}
finally {
    Remove-PSSession $session
}
exit $exit
