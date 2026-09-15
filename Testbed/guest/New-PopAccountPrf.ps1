<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED, AND IT IS A SPIKE RATHER THAN A ROUTE.
    ============================================================================================

    Written by an agent forbidden to run it. Verified by PARSING only.

    READ THIS BEFORE SPENDING TIME ON IT. This script is NOT expected to work. It exists because
    'we tried the one remaining free candidate and here is exactly how far it got' is worth more
    than another round of searching, and because the measurement it takes afterwards is useful
    whatever the answer.

    THE PROBLEM. Docs/live-tier-on-the-vm.md sections 2.8 and 2.8b need TWO POP3 accounts - the
    dummy account and the identity account - each delivering into its OWN PST. There is no free
    programmatic route to creating either one, and the evidence runs three ways:

      * THE OBJECT MODEL CANNOT. Namespace.Accounts has no Add, and Account.DeliveryStore is
        read-only - so even given an account, the OM cannot set 'deliver new messages to', which
        section 2.8 calls "the failure to bet on".
      * MAPI CANNOT. POP3 stopped being a MAPI message service; account administration moved
        behind IOlkAccountManager, which is undocumented and whose create method is one of
        fourteen placeholder vtable slots. There is no service name to pass to CreateMsgService.
      * THE REGISTRY HAS NO RECIPE. The account subkeys under 9375CFF0413111d3B88A00104B2A6676
        are readable - this product reads them, in SignatureCatalog - but nobody has published a
        working synthesis on 16.x, and the stored password blobs are DPAPI-sealed per Windows user
        per machine, so they cannot be authored offline or copied between guests.

    WHICH LEAVES .prf, AND WHY IT IS STILL PROBABLY NOT THE ANSWER. Two independent objections:

      1. STALENESS. Microsoft's PRF documentation is the Office Resource Kit's and stops at Office
         2013. No source at all discusses PRF POP3 sections on 16.x. That is not evidence it
         broke; it is a complete absence of evidence that it works.
      2. A STRUCTURAL ONE, WHICH IS WORSE. The delivery-store binding is PROP_ACCT_DELIVERY_STORE,
         a PT_BINARY EntryID. A .prf is an INI file of text assignments, and a store's EntryID
         does not exist until that store exists in that profile on that machine. So a .prf
         structurally cannot carry the binding, and even a PRF that works leaves both accounts
         delivering to the profile default - which is the exact misconfiguration section 2.8 says
         looks like a sink that is not working, and costs a 180-second timeout pointing nowhere
         useful.

    SO WHAT IS THIS FOR. Two things, and the second is the real one:

      * it generates the .prf, so the spike costs minutes rather than an afternoon;
      * IT MEASURES THE OUTCOME. -Verify reads Accounts.Count and, for each account,
        Account.DeliveryStore.DisplayName over COM. Account.DeliveryStore is read-only but it IS
        readable, so this settles objection 2 empirically rather than by argument. If it turns out
        the binding does land, that is a genuinely valuable surprise and the note in
        .work/profile-automation-research.md section 4 should be corrected.

    THE PROPERTY NAMES BELOW ARE THE LEAST CERTAIN THING IN THIS WHOLE SCRIPT SET. They come from
    Office Resource Kit material for an Office version three releases behind this one. If the
    import does nothing, suspect the spelling before you suspect the mechanism, and check it
    against a .prf that Outlook itself produced rather than against more searching.

    NO PASSWORD IS WRITTEN, deliberately and for two reasons. The sink accepts anything so there
    is nothing worth storing, and the stored form is DPAPI-sealed per user per machine so a
    literal in a file could not become it anyway. Type it once in Outlook if it is ever asked for.
    Testbed/README.md section 4 also forbids a credential anywhere under Testbed/, and
    .github/scripts/check-testbed-references.ps1 check 6 fails the build over one.

    WHAT IT EXPECTS TO START FROM. The profile already exists and already carries the PST that is
    meant to be the delivery store - New-OutlookProfile.ps1 and Add-OutlookPstStore.ps1 first. The
    PRF route cannot create the store and bind to it in one pass, because of objection 2.

    IF THIS SPIKE FAILS - which is the expected outcome - the build keeps ONE GUI pass: add the
    two accounts in Outlook's Account Settings wizard and set each one's delivery store. Then TAKE
    A CHECKPOINT. That turns the only unscripted step from a per-rebuild cost into a
    per-guest-lifetime one, which is the whole consolation prize here.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER ProfileName
    The profile the account should land in. Must already exist.

.PARAMETER EmailAddress
    The account's address, e.g. test@vm.invalid. Under RFC 2606 .invalid is guaranteed
    unresolvable, so a misconfiguration cannot leak mail anywhere.

.PARAMETER DeliveryStoreDisplayName
    The store this account should deliver into. Written into the .prf for completeness and
    CHECKED afterwards by -Verify, which is the part that actually tells you something.

.PARAMETER PrfPath
    Where to write the .prf.

.PARAMETER Import
    Also run outlook.exe /importprf. THIS STARTS OUTLOOK, and on a guest that must never draw a
    window that is a real consideration - see the warning it prints.

.PARAMETER Verify
    After importing, read back over COM what accounts exist and where each one delivers.

.PARAMETER Execute
    Without it, nothing is written.

.EXAMPLE
    .\New-PopAccountPrf.ps1 -ProfileName OutlookAITest -EmailAddress test@vm.invalid -DeliveryStoreDisplayName 'test@vm.invalid' -Execute
    .\New-PopAccountPrf.ps1 -ProfileName OutlookAITest -EmailAddress test@vm.invalid -DeliveryStoreDisplayName 'test@vm.invalid' -Execute -Import -Verify
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $ProfileName,
    [Parameter(Mandatory = $true)] [string] $EmailAddress,
    [Parameter(Mandatory = $true)] [string] $DeliveryStoreDisplayName,
    [string] $DisplayName = 'OutlookAI Testbed',
    [string] $PopHost = '127.0.0.1',
    [int]    $PopPort = 110,
    [string] $SmtpHost = '127.0.0.1',
    [int]    $SmtpPort = 25,
    [string] $UserName,
    [string] $PrfPath = 'C:\OutlookAI-Q5\prf\account.prf',
    [switch] $Import,
    [switch] $Verify,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\OutlookMapiInterop.ps1"

if (-not $UserName) { $UserName = $EmailAddress }

Write-Host 'PRF SPIKE. This is not expected to work - read the header before reading the output.'
Write-Host ''
Write-Host "profile        : $ProfileName"
Write-Host "address        : $EmailAddress"
Write-Host "delivery store : $DeliveryStoreDisplayName   (cannot be expressed in a .prf - see header)"
Write-Host "pop            : ${PopHost}:$PopPort"
Write-Host "smtp           : ${SmtpHost}:$SmtpPort"
Write-Host "prf            : $PrfPath"
Write-Host ''

# The template. Every line carrying a name I could not verify against a current source is marked,
# because a reader needs to know which half to suspect when nothing happens.
$prf = @"
; Generated by Testbed/guest/New-PopAccountPrf.ps1. A SPIKE, not a supported route.
;
; The section layout is Office Resource Kit shape. The PROPERTY NAMES in the account section are
; the least certain thing here: they are documented for Office 2003-2013 and this machine is
; Office LTSC 2024 (16.0.17932). If the import does nothing, suspect these spellings first.
;
; PROP_ACCT_DELIVERY_STORE is deliberately ABSENT. It is a PT_BINARY EntryID and an INI file
; cannot carry one; see the script header. That is why the delivery store must be set by hand
; afterwards, and why -Verify reads it back rather than assuming.
;
; No credential appears in this file, by design. The sink accepts anything, and the stored form is
; DPAPI-sealed per user per machine so a literal here could never become it.

[General]
Custom=1
ProfileName=$ProfileName
DefaultProfile=No
OverwriteProfile=No
ModifyDefaultProfileIfPresent=true
BackupProfile=false

[Service List]
ServiceEGS0=Internet E-mail

[Internet Account List]
Account1=Internet E-mail

[Account1]
; --- unverified spellings from here down ---------------------------------------------------
UNIQUE_SERVICE=No
ServiceName=INETMAIL
AccountName=$EmailAddress
PROP_ACCT_NAME=$EmailAddress
PROP_ACCT_USER_DISPLAY_NAME=$DisplayName
PROP_ACCT_USER_EMAIL_ADDR=$EmailAddress
PROP_ACCT_POP3_SERVER=$PopHost
PROP_ACCT_POP3_PORT=$PopPort
PROP_ACCT_POP3_USER_NAME=$UserName
PROP_ACCT_POP3_USE_SPA=0
PROP_ACCT_POP3_LEAVE_ON_SERVER=0
PROP_ACCT_SMTP_SERVER=$SmtpHost
PROP_ACCT_SMTP_PORT=$SmtpPort
PROP_ACCT_SMTP_USE_AUTH=0
"@

if (-not $Execute) {
    Write-Host 'Dry run. This is what would be written:'
    Write-Host ''
    Write-Host $prf
    Write-Host ''
    Write-Host 'Nothing written. Re-run with -Execute.'
    return
}

Assert-TestbedGuest -ExpectedUser $ExpectedUser

$prfDirectory = Split-Path -Parent $PrfPath
if (-not (Test-Path -LiteralPath $prfDirectory)) {
    New-Item -ItemType Directory -Force -Path $prfDirectory | Out-Null
}

# ANSI, not UTF-8. A .prf is read as an INI file by code that predates anybody caring, and a BOM
# at the top of one is a documented way to make the first section header invisible.
[System.IO.File]::WriteAllText($PrfPath, $prf, [System.Text.Encoding]::Default)

if (-not (Test-Path -LiteralPath $PrfPath)) {
    throw "Wrote $PrfPath and it is not there."
}
Write-Host "Wrote $PrfPath ($((Get-Item -LiteralPath $PrfPath).Length) bytes, ANSI, no BOM)."

if ($Import) {
    Assert-OutlookNotRunning

    $outlookExe = $null
    foreach ($candidate in @(
            'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
            'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE')) {
        if (Test-Path -LiteralPath $candidate) { $outlookExe = $candidate; break }
    }
    if (-not $outlookExe) {
        throw 'Could not find OUTLOOK.EXE in either Office16 location. Import by hand: outlook.exe /importprf "' + $PrfPath + '"'
    }

    Write-Host ''
    Write-Warning 'STARTING OUTLOOK. /importprf is reported to bring up the full UI rather than running'
    Write-Warning 'headless. On a guest that must never draw a window this is worth watching for - if a'
    Write-Warning 'window appears, that alone is a reason to abandon this route.'
    Write-Host ''

    # Output goes nowhere useful from a GUI process, and holding the pipe on something that spawns
    # its own children is how a call never sees EOF. Start it and wait on the process, not a pipe.
    $process = Start-Process -FilePath $outlookExe -ArgumentList @('/importprf', "`"$PrfPath`"") -PassThru
    Write-Host "started $($process.Id); giving it 120 s to settle"
    $deadline = (Get-Date).AddSeconds(120)
    while ((Get-Date) -lt $deadline -and -not $process.HasExited) {
        Start-Sleep -Seconds 5
    }
    if ($process.HasExited) { Write-Host "Outlook exited with $($process.ExitCode)." }
    else { Write-Host 'Outlook is still running. That is normal for /importprf; it does not exit by itself.' }
}

if ($Verify) {
    Write-Host ''
    Write-Host 'Reading back what actually exists. This is the half of the spike that is worth having.'

    $outlook = $null
    $ns = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace('MAPI')
        $ns.Logon($ProfileName, '', $false, $false)

        $count = $ns.Accounts.Count
        Write-Host "  Accounts.Count : $count"

        if ($count -eq 0) {
            Write-Host ''
            Write-Host 'THE SPIKE FAILED, which is the expected outcome. The PRF import created no account.'
            Write-Host 'Add the account in Outlook Account Settings by hand, set its delivery store, and TAKE'
            Write-Host 'A CHECKPOINT - that makes it a once-per-guest cost instead of a once-per-rebuild one.'
            exit 1
        }

        $bindingCorrect = $false
        for ($i = 1; $i -le $count; $i++) {
            $account = $ns.Accounts.Item($i)
            $smtp = $null
            $delivery = $null
            try { $smtp = $account.SmtpAddress } catch { $smtp = '(unreadable)' }
            try { $delivery = $account.DeliveryStore.DisplayName } catch { $delivery = '(unreadable)' }
            Write-Host ("  account {0} : {1}   delivers to: {2}" -f $i, $smtp, $delivery)
            if ($smtp -eq $EmailAddress -and $delivery -eq $DeliveryStoreDisplayName) { $bindingCorrect = $true }
        }

        Write-Host ''
        if ($bindingCorrect) {
            Write-Host 'SURPRISE, AND A VALUABLE ONE: the account exists AND delivers into the intended store.'
            Write-Host 'That contradicts objection 2 in this header - a .prf was not thought able to carry the'
            Write-Host 'delivery-store binding. Record what actually happened and correct section 4 of'
            Write-Host '.work/profile-automation-research.md; it changes the recommendation.'
        }
        else {
            Write-Host 'PARTIAL: an account exists but it does not deliver into the intended store. This is'
            Write-Host 'the predicted outcome of objection 2. Fix the delivery store by hand in Account'
            Write-Host 'Settings - it is the setting section 2.8 calls the failure to bet on, and a wrong one'
            Write-Host 'shows up only as a 180-second arrival timeout pointing nowhere useful.'
            exit 1
        }
    }
    finally {
        if ($null -ne $ns) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns) }
        if ($null -ne $outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
    }
}
