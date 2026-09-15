<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it - the machine it was written on is the maintainer's
    own workstation, carrying a real Outlook profile and real delegate mailboxes. It was verified
    by PARSING and by nothing else. Once it has run on a guest, replace this banner with what it
    actually did.

    WHAT IT MAKES. An Outlook profile, from nothing, with no GUI. Two shapes:

      * ACCOUNT-LESS - the corpus profile. Docs/live-tier-on-the-vm.md section 1.2: corpus-build
        refuses any profile that has a mail account, with no override flag, because a build
        creates unsent items in bulk and the first real run left 5,532 of them in an Outbox -
        inert only because that profile could not send. This shape is MANDATORY, not a
        convenience.
      * WITH PST STORES - pass -Store once per store. Each store gets an EXACT display name; see
        Add-OutlookPstStore.ps1 for why that is hard and why it is done through MAPI.

    WHAT IT CANNOT MAKE: a mail account. Nothing free can - see 'THE WALL' below. This script
    builds the profile and its stores; the POP3 accounts are a separate, manual step.

    WHAT IT EXPECTS TO START FROM. A guest with Office installed and Outlook NOT running, logged
    on as vmadmin. It does not care whether other profiles exist. It is IDEMPOTENT on the profile:
    an existing profile of the same name is left alone and reported, never recreated, because
    recreating one silently discards whatever stores it already carried.

    THE WALL, stated here so nobody goes looking for the missing half. There is NO free
    programmatic route to creating a POP3 account, and the evidence runs three ways:

      * the object model has no Accounts.Add, and Account.DeliveryStore is read-only - so the OM
        cannot set 'deliver new messages to', which section 2.8 calls the failure to bet on;
      * MAPI cannot, because POP3 stopped being a MAPI message service: account administration
        moved behind the undocumented IOlkAccountManager;
      * the registry has no published working recipe on 16.x, and the stored password blobs are
        DPAPI-sealed per Windows user per machine, so they cannot be authored offline.

    The one free candidate is a .prf file, and New-PopAccountPrf.ps1 is that spike - read its
    header before spending time on it. So the honest build sequence is: this script, then
    Add-OutlookPstStore.ps1, then Set-DefaultOutlookProfile.ps1, then ONE GUI pass to add the two
    POP3 accounts and point each one's delivery store, then Set-AccountSignature.ps1.

    TAKE A CHECKPOINT IMMEDIATELY AFTER THAT GUI PASS. It is the only unscripted step left, so a
    checkpoint there turns it from a per-rebuild cost into a per-guest-lifetime one.

    HOW IT VERIFIES ITSELF, because a script that half-works costs a checkpoint revert:

      1. the profile appears in IProfAdmin::GetProfileTable - NOT merely 'CreateProfile returned
         S_OK'. The HRESULT is not evidence: DeleteProfile returns S_OK without deleting when the
         profile is in use, and that is the same interface.
      2. every store asked for appears in the message service table under EXACTLY the display
         name asked for. A near-miss fails the script.
      3. -VerifyWithOutlook additionally starts Outlook on the profile and reads Session.Stores
         and Session.Accounts.Count over COM. That is the only check that answers 'what will the
         tests see', and for an account-less profile it is the check that matters most: the
         generator's predicate is Accounts.Count == 0 read over COM (ComCorpusMailbox
         .ReadProfileFacts), fail-closed, so an unreadable answer refuses too.

    A NOTE ON THE ACCOUNT-LESS PROFILE THAT WILL OTHERWISE LOOK LIKE A BUG. MAPI injects the
    Contact Address Book provider (CONTAB) into a bare profile the first time it is opened, so a
    profile created with zero services does not stay at zero SERVICES. That is expected and should
    be harmless: CONTAB is an address book provider, and the OM's Accounts collection holds
    Account objects, which address book providers are not. That last step is an INFERENCE from the
    OM's type surface, not a measurement - which is exactly why -VerifyWithOutlook exists and why
    you should use it on the corpus profile at least once.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER Name
    Profile name to create.

.PARAMETER Store
    Repeatable. 'DisplayName=C:\path\to.pst'. The display name is what Outlook shows and what the
    tests match on; the path is where the file goes. Omit entirely for an account-less,
    store-less profile.

.PARAMETER MakeDefault
    Also make this the default profile. Equivalent to running Set-DefaultOutlookProfile.ps1
    afterwards, and it does the same prompt-suppression.

.PARAMETER VerifyWithOutlook
    After building, start Outlook on this profile over COM and read back what it actually shows.
    Slower, needs session 1, and is the only check that proves anything about what the tests see.

.PARAMETER Preflight
    MAPIInitialize, report, MAPIUninitialize. Writes nothing, touches no profile. RUN THIS FIRST
    on a checkpoint you are willing to lose: if the interop is going to crash the process, this is
    where it does it, with nothing at stake.

.PARAMETER Execute
    Without it, nothing is written and the script prints what it would do.

.EXAMPLE
    .\New-OutlookProfile.ps1 -Preflight
    .\New-OutlookProfile.ps1 -Name OutlookAICorpus -Store 'Corpus A=C:\OutlookAI-Q5\pst\corpus-a.pst' -Execute -VerifyWithOutlook
    .\New-OutlookProfile.ps1 -Name OutlookAITest -Execute
#>
[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Build')] [string] $Name,
    [Parameter(ParameterSetName = 'Build')] [string[]] $Store = @(),
    [Parameter(ParameterSetName = 'Build')] [switch] $MakeDefault,
    [Parameter(ParameterSetName = 'Build')] [switch] $VerifyWithOutlook,
    [Parameter(Mandatory = $true, ParameterSetName = 'Preflight')] [switch] $Preflight,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\OutlookMapiInterop.ps1"

# ---------------------------------------------------------------------------------------------
# Preflight. Nothing here writes, so it runs without -Execute on purpose.
# ---------------------------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'Preflight') {
    Write-Host 'PREFLIGHT - MAPIInitialize, read the profile table, MAPIUninitialize. Writes nothing.'
    Write-Host ''

    $clients = 'HKLM:\SOFTWARE\Clients\Mail\Microsoft Outlook'
    if (Test-Path $clients) {
        $values = Get-ItemProperty -Path $clients
        foreach ($valueName in @('DLLPath', 'DLLPathEx')) {
            $value = $values.$valueName
            $exists = $false
            if ($value) { $exists = Test-Path -LiteralPath $value }
            Write-Host ("  {0,-10} = {1}  [exists: {2}]" -f $valueName, $value, $exists)
        }
    }
    else {
        Write-Warning "$clients does not exist. Extended MAPI will not resolve; is classic Outlook installed?"
    }
    Write-Host ("  {0,-10} = {1}" -f 'PS 64-bit', [Environment]::Is64BitProcess)
    Write-Host ("  {0,-10} = {1}" -f 'Apartment', [Threading.Thread]::CurrentThread.GetApartmentState())
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        Write-Warning 'This thread is not STA. MAPI wants STA; powershell.exe -MTA gives RPC_E_CHANGED_MODE.'
    }
    Write-Host ''

    Invoke-WithProfAdmin -Body {
        param($admin)
        $profiles = Get-MapiProfile -ProfAdmin $admin
        Write-Host "MAPI is alive. $($profiles.Count) profile(s):"
        foreach ($p in $profiles) {
            $marker = ''
            if ($p.IsDefault) { $marker = '   <- default' }
            Write-Host ("  {0}{1}" -f $p.Name, $marker)
        }
    }

    Write-Host ''
    Write-Host 'Preflight passed. The interop loads, MAPI initialises, and the profile table reads.'
    return
}

# ---------------------------------------------------------------------------------------------
# Parse -Store into ordered pairs before anything is touched, so a typo fails at argument time.
# ---------------------------------------------------------------------------------------------
$requested = @()
foreach ($spec in $Store) {
    $split = $spec.IndexOf('=')
    if ($split -lt 1 -or $split -eq ($spec.Length - 1)) {
        throw "-Store wants 'DisplayName=C:\path\to.pst'; got '$spec'. The display name comes first because that is the half the tests match on."
    }
    $displayName = $spec.Substring(0, $split).Trim()
    $path = $spec.Substring($split + 1).Trim()
    $requested += [pscustomobject]@{
        DisplayName = $displayName
        Path        = (Resolve-PstPath -Path $path)
    }
}

Write-Host "profile        : $Name"
Write-Host "stores         : $($requested.Count)"
foreach ($item in $requested) {
    Write-Host ("  {0,-28} {1}" -f $item.DisplayName, $item.Path)
}
Write-Host "make default   : $($MakeDefault.IsPresent)"
Write-Host ''

if (-not $Execute) {
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    Write-Host 'Run -Preflight first if you have not: it is the cheapest way to find out whether'
    Write-Host 'the interop loads at all on this guest.'
    return
}

Assert-TestbedGuest -ExpectedUser $ExpectedUser
Assert-OutlookNotRunning

# ---------------------------------------------------------------------------------------------
# Build.
# ---------------------------------------------------------------------------------------------
Invoke-WithProfAdmin -Body {
    param($admin)

    $existing = Get-MapiProfile -ProfAdmin $admin
    $already = $false
    foreach ($p in $existing) {
        if ($p.Name -eq $Name) { $already = $true }
    }

    if ($already) {
        # Idempotent, and deliberately NOT 'delete and recreate': recreating discards whatever
        # stores the profile already carried, which on this machine could be a 400 MB corpus.
        Write-Host "Profile '$Name' already exists - left exactly as it is."
        Write-Host 'Stores are added by Add-OutlookPstStore.ps1, which is idempotent per store.'
    }
    else {
        # ulFlags 0, NOT MAPI_DEFAULT_SERVICES: that flag is what would add the default service
        # set, and an account-less profile is the entire point of the corpus half.
        $hr = $admin.CreateProfile($Name, $null, [IntPtr]::Zero, 0)
        Assert-MapiOk -HResult $hr -What "IProfAdmin::CreateProfile('$Name')"

        # VERIFY THROUGH THE TABLE, NOT THE HRESULT.
        $after = Get-MapiProfile -ProfAdmin $admin
        $found = $false
        foreach ($p in $after) {
            if ($p.Name -eq $Name) { $found = $true }
        }
        if (-not $found) {
            throw "CreateProfile returned success and '$Name' is NOT in the profile table. Do not continue: the HRESULT and the table disagree, and the table is the one that is true."
        }
        Write-Host "Created profile '$Name' (confirmed in the profile table)."
    }

    if ($requested.Count -gt 0) {
        $serviceAdmin = $null
        $hr = $admin.AdminServices($Name, $null, [IntPtr]::Zero, 0, [ref] $serviceAdmin)
        Assert-MapiOk -HResult $hr -What "IProfAdmin::AdminServices('$Name')"
        try {
            foreach ($item in $requested) {
                $services = Get-MapiService -ServiceAdmin $serviceAdmin
                $present = $false
                foreach ($service in $services) {
                    if ($service.DisplayName -eq $item.DisplayName) { $present = $true }
                }
                if ($present) {
                    Write-Host "  store '$($item.DisplayName)' is already in this profile - skipped."
                    continue
                }

                Write-Host "  adding '$($item.DisplayName)' -> $($item.Path)"
                [void](Add-MapiPstService -ServiceAdmin $serviceAdmin -PstPath $item.Path -DisplayName $item.DisplayName)
            }

            # Verify every requested store, by exact name, in one pass at the end.
            $final = Get-MapiService -ServiceAdmin $serviceAdmin
            $names = @()
            foreach ($service in $final) { $names += $service.DisplayName }
            $missing = @()
            foreach ($item in $requested) {
                if ($names -notcontains $item.DisplayName) { $missing += $item.DisplayName }
            }
            if ($missing.Count -gt 0) {
                throw @"
$($missing.Count) store(s) are NOT in the service table under the name asked for: $($missing -join ', ')
The table holds: $($names -join ', ')

A name that came back DIFFERENT rather than absent is the interesting case and the dangerous one:
it means the provider accepted the store and renamed it, and every test that looks the store up by
display name will miss it. Docs/live-tier-on-the-vm.md section 8 item 2 is that question.
"@
            }
            Write-Host "All $($requested.Count) store(s) present under the exact names asked for."
        }
        finally {
            if ($null -ne $serviceAdmin) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($serviceAdmin) }
        }
    }

    if ($MakeDefault) {
        $hr = $admin.SetDefaultProfile($Name, 0)
        Assert-MapiOk -HResult $hr -What "IProfAdmin::SetDefaultProfile('$Name')"
        $check = Get-MapiProfile -ProfAdmin $admin
        $isDefault = $false
        foreach ($p in $check) {
            if ($p.Name -eq $Name -and $p.IsDefault) { $isDefault = $true }
        }
        if (-not $isDefault) {
            throw "SetDefaultProfile returned success and the profile table does not report '$Name' as default. Do not continue."
        }
        Write-Host "'$Name' is now the default profile (confirmed in the profile table)."
        Write-Host 'Run Set-DefaultOutlookProfile.ps1 to also switch OFF the profile prompt, which is a'
        Write-Host 'separate setting - a prompting profile cannot be driven over COM.'
    }
}

# ---------------------------------------------------------------------------------------------
# The only verification that answers 'what will the tests see'.
# ---------------------------------------------------------------------------------------------
if ($VerifyWithOutlook) {
    Write-Host ''
    Write-Host 'Verifying over COM. This starts Outlook and needs session 1.'

    $outlook = $null
    $ns = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace('MAPI')
        $ns.Logon($Name, '', $false, $false)

        $seen = @()
        foreach ($s in $ns.Stores) { $seen += $s.DisplayName }
        Write-Host "  stores Outlook reports : $($seen -join ' | ')"

        $accountCount = $ns.Accounts.Count
        Write-Host "  Accounts.Count         : $accountCount"

        $missing = @()
        foreach ($item in $requested) {
            if ($seen -notcontains $item.DisplayName) { $missing += $item.DisplayName }
        }
        if ($missing.Count -gt 0) {
            throw "Outlook does not show $($missing.Count) of the stores under the expected name: $($missing -join ', '). MAPI and Outlook disagree, and Outlook is what the tests read."
        }

        if ($requested.Count -eq 0 -and $accountCount -ne 0) {
            throw @"
This profile was built account-less and Outlook reports Accounts.Count = $accountCount.

corpus-build will refuse it. The predicate is exactly this read - CorpusSafety.EvaluateProfile
over ComCorpusMailbox.ReadProfileFacts - and it is fail-closed, so there is no way to talk it
past this. Find out what the account is before doing anything else.
"@
        }
        Write-Host '  verified.'
    }
    finally {
        if ($null -ne $ns) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns) }
        if ($null -ne $outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
        # Deliberately NOT quitting Outlook: mailbox-safety rule 7 wants COM references released
        # before any Quit, and leaving it headless is the documented preference.
    }
}

Write-Host ''
Write-Host 'Done. Next steps, in order:'
Write-Host '  Add-OutlookPstStore.ps1      - any further stores, and the -NameProbe that settles the @ question'
Write-Host '  Set-DefaultOutlookProfile.ps1 - switch default and suppress the profile prompt'
Write-Host '  THE GUI PASS                 - two POP3 accounts and their delivery stores. No free route exists.'
Write-Host '  Set-AccountSignature.ps1     - the identity signature, AFTER the accounts exist'
