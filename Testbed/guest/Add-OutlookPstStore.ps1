<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it - the machine it was written on is the maintainer's
    own workstation, with a real Outlook profile on it. Verified by PARSING only. Replace this
    banner with what it actually did once it has run on a guest.

    WHAT IT DOES. Adds a PST to an existing Outlook profile with an EXACT display name, and
    settles the one open question that gates a whole family of tests.

    WHY THIS IS HARDER THAN IT LOOKS, and why it is not Namespace.AddStoreEx. AddStoreEx is two
    lines and adds a PST perfectly well - but it CANNOT NAME IT. Store.DisplayName is read-only in
    the object model; PropertyAccessor.SetProperty on 0x3001001F at store level is reported
    blocked; and whether renaming the store's root folder makes DisplayName follow is something
    nobody could establish either way.

    So the store is added through Extended MAPI instead, where the PST provider's own configure
    call takes PR_DISPLAY_NAME and applies it AT CREATION TIME - which sidesteps the rename
    question entirely rather than answering it. The corroboration that this is real behaviour and
    not wishful reading: the provider defines PST_CONFIG_PRESERVE_DISPLAY_NAME, a flag whose only
    job is to SUPPRESS that behaviour. A switch to turn something off is evidence it happens.

    THE QUESTION THIS SCRIPT EXISTS TO SETTLE: -NameProbe.

    Docs/live-tier-on-the-vm.md section 8 item 2 and Testbed/README.md section 6 item 10 both
    carry it, and both say it costs five minutes and gates the whole draft family: DOES OUTLOOK
    ACCEPT '@' IN A STORE DISPLAY NAME? The hub store must be called literally 'test@vm.invalid'
    and the identity store 'identity@vm.invalid', because several tests hand the store's display
    name straight to NewDraft as an address.

    Nothing in the documentation or in any community source settles it - no stated character
    restriction anywhere, which is weak evidence and should not be built on. -NameProbe settles it
    empirically, in a THROWAWAY PROFILE it creates and deletes, and reports one of three outcomes:

      ACCEPTED    the name comes back exactly. The gate opens; nothing else changes.
      REJECTED    the call fails. The gate closes: the hub and identity stores need a different
                  naming scheme, which is a test-side change, not a machine one.
      TRANSFORMED the name comes back DIFFERENT. This is the outcome worth having a probe for:
                  it would otherwise be found much later, as stores the tests cannot find by name,
                  on a machine that looked like it had been built correctly.

    The probe checks BOTH layers, and they are not the same question. MAPI reporting the name back
    proves the property was stored. Only Outlook reporting it proves the tests will see it,
    because Store.DisplayName is what they read.

    WHAT IT EXPECTS TO START FROM. A guest, logged on as vmadmin, Outlook not running, and - for
    the non-probe mode - a profile that already exists (New-OutlookProfile.ps1 makes one).

    IDEMPOTENCY, and where it is weaker than it looks. Adding a store that is already in the
    profile UNDER THAT DISPLAY NAME is a no-op. The match is on display name, because that is what
    the message service table exposes and what this script controls. It is NOT a match on file
    path: reading a service's PR_PST_PATH back needs OpenProfileSection, which this interop does
    not implement. So adding the same FILE twice under two different display names would not be
    caught here - it would be caught by -VerifyWithOutlook, which matches on Store.FilePath
    resolved and compared case-insensitively.

    ONE TRAP, from the same source as the rest. A PST path spelled with different casing in two
    profiles is reported to give MAPI_E_FAILONEPROVIDER. Every path here goes through
    Resolve-PstPath once and the resulting string is what is used everywhere.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER ProfileName
    The profile to add the store to. Must already exist. Named -ProfileName rather than the more
    obvious -Profile because $Profile is a PowerShell automatic variable and a parameter of that
    name shadows it for the whole script.

.PARAMETER DisplayName
    What Outlook must show. Exact: the script fails if what comes back differs by so much as a
    character.

.PARAMETER Path
    Where the .pst goes. The provider creates the file; the DIRECTORY must already exist.

.PARAMETER NameProbe
    Settle the '@' question instead of building anything. Creates a throwaway profile and a
    throwaway PST, reports what happened to the name, and removes both.

.PARAMETER ProbeName
    The display name the probe asks for. Defaults to the name the hub store actually needs.

.PARAMETER VerifyWithOutlook
    Read the result back over COM as well as over MAPI. Slower, needs session 1, and is the only
    half that proves what the tests will see.

.PARAMETER Execute
    Without it, nothing is written.

.EXAMPLE
    .\Add-OutlookPstStore.ps1 -NameProbe -Execute -VerifyWithOutlook
    .\Add-OutlookPstStore.ps1 -ProfileName OutlookAITest -DisplayName 'test@vm.invalid' -Path C:\OutlookAI-Q5\pst\hub.pst -Execute -VerifyWithOutlook
#>
[CmdletBinding(DefaultParameterSetName = 'Add')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Add')] [string] $ProfileName,
    [Parameter(Mandatory = $true, ParameterSetName = 'Add')] [string] $DisplayName,
    [Parameter(Mandatory = $true, ParameterSetName = 'Add')] [string] $Path,
    [Parameter(Mandatory = $true, ParameterSetName = 'Probe')] [switch] $NameProbe,
    [Parameter(ParameterSetName = 'Probe')] [string] $ProbeName = 'test@vm.invalid',
    [Parameter(ParameterSetName = 'Probe')] [string] $ProbeDirectory = 'C:\OutlookAI-Q5\name-probe',
    [switch] $VerifyWithOutlook,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\OutlookMapiInterop.ps1"

# Reads what Outlook itself reports for a store, which is the only reading the tests care about.
function Get-OutlookStoreView {
    param([Parameter(Mandatory = $true)] [string] $ProfileName)

    $outlook = $null
    $ns = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $ns = $outlook.GetNamespace('MAPI')
        $ns.Logon($ProfileName, '', $false, $false)

        $view = @()
        foreach ($s in $ns.Stores) {
            $filePath = $null
            try { $filePath = $s.FilePath } catch { $filePath = $null }
            $view += [pscustomobject]@{ DisplayName = $s.DisplayName; FilePath = $filePath }
        }
        return , $view
    }
    finally {
        if ($null -ne $ns) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns) }
        if ($null -ne $outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
    }
}

# =============================================================================================
# PROBE MODE
# =============================================================================================
if ($PSCmdlet.ParameterSetName -eq 'Probe') {

    $probeProfile = 'OutlookAI-NameProbe'
    $probePst = Join-Path $ProbeDirectory 'name-probe.pst'

    Write-Host 'NAME PROBE - does Outlook accept this as a store display name?'
    Write-Host ''
    Write-Host "  name asked for : $ProbeName"
    Write-Host "  profile        : $probeProfile   (created and deleted by this script)"
    Write-Host "  pst            : $probePst       (created and deleted by this script)"
    Write-Host ''
    Write-Host 'It touches no existing profile and no existing store. It answers'
    Write-Host 'Docs/live-tier-on-the-vm.md section 8 item 2.'
    Write-Host ''

    if (-not $Execute) {
        Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
        return
    }

    Assert-TestbedGuest -ExpectedUser $ExpectedUser
    Assert-OutlookNotRunning

    if (Test-Path -LiteralPath $probePst) {
        throw "A file is already at $probePst. This script only ever removes a PST it created in this run, so it will not touch that one. Move it aside and try again."
    }
    if (-not (Test-Path -LiteralPath $ProbeDirectory)) {
        New-Item -ItemType Directory -Force -Path $ProbeDirectory | Out-Null
    }
    $probePst = Resolve-PstPath -Path $probePst

    $script:verdict = 'UNKNOWN'
    $script:mapiSaw = $null
    $script:outlookSaw = $null
    $script:failure = $null

    Invoke-WithProfAdmin -Body {
        param($admin)

        $existing = Get-MapiProfile -ProfAdmin $admin
        foreach ($p in $existing) {
            if ($p.Name -eq $probeProfile) {
                throw "A profile named '$probeProfile' already exists. A previous probe did not clean up. Inspect it and remove it before running again - this script will not delete a profile it did not just create."
            }
        }

        Assert-MapiOk -HResult $admin.CreateProfile($probeProfile, $null, [IntPtr]::Zero, 0) -What 'CreateProfile (probe)'

        $serviceAdmin = $null
        try {
            Assert-MapiOk -HResult $admin.AdminServices($probeProfile, $null, [IntPtr]::Zero, 0, [ref] $serviceAdmin) -What 'AdminServices (probe)'

            try {
                [void](Add-MapiPstService -ServiceAdmin $serviceAdmin -PstPath $probePst -DisplayName $ProbeName)
            }
            catch {
                $script:failure = $_.Exception.Message
                $script:verdict = 'REJECTED'
            }

            if ($script:verdict -ne 'REJECTED') {
                $services = Get-MapiService -ServiceAdmin $serviceAdmin
                foreach ($service in $services) {
                    if ($service.ServiceName -eq 'MSUPST MS') { $script:mapiSaw = $service.DisplayName }
                }
            }
        }
        finally {
            if ($null -ne $serviceAdmin) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($serviceAdmin) }
        }
    }

    if ($script:verdict -ne 'REJECTED' -and $VerifyWithOutlook) {
        $view = Get-OutlookStoreView -ProfileName $probeProfile
        foreach ($store in $view) {
            if ($null -ne $store.FilePath -and $store.FilePath -ieq $probePst) { $script:outlookSaw = $store.DisplayName }
        }
        if ($null -eq $script:outlookSaw) {
            # Fall back to whichever store is not the profile's own default-ish one; report it all
            # rather than guess, because guessing here is how a TRANSFORMED result gets missed.
            Write-Warning 'Could not match the probe PST by FilePath. Everything Outlook reported:'
            foreach ($store in $view) { Write-Warning ("  {0}  <-  {1}" -f $store.DisplayName, $store.FilePath) }
        }
    }

    if ($script:verdict -ne 'REJECTED') {
        $observed = $script:mapiSaw
        if ($VerifyWithOutlook -and $null -ne $script:outlookSaw) { $observed = $script:outlookSaw }
        if ($observed -ceq $ProbeName) { $script:verdict = 'ACCEPTED' }
        elseif ($null -eq $observed) { $script:verdict = 'UNREADABLE' }
        else { $script:verdict = 'TRANSFORMED' }
    }

    # --- teardown. Only ever removes what this run created. ------------------------------------
    $script:cleanupProblems = @()
    try {
        Invoke-WithProfAdmin -Body {
            param($admin)
            [void]$admin.DeleteProfile($probeProfile, 0)

            # DeleteProfile returns S_OK without deleting when the profile is in use, so the
            # HRESULT is ignored on purpose and the table is what is believed.
            $after = Get-MapiProfile -ProfAdmin $admin
            foreach ($p in $after) {
                if ($p.Name -eq $probeProfile) {
                    $script:cleanupProblems += "The probe profile '$probeProfile' still exists after DeleteProfile. Remove it by hand (Mail control panel) before probing again."
                }
            }
        }
    }
    catch {
        $script:cleanupProblems += "Removing the probe profile threw: $($_.Exception.Message)"
    }

    if (Test-Path -LiteralPath $probePst) {
        try { Remove-Item -LiteralPath $probePst -Force }
        catch { $script:cleanupProblems += "Could not delete the probe PST at $probePst : $($_.Exception.Message)" }
    }

    Write-Host ''
    Write-Host '--------------------------------------------------------------------'
    Write-Host ("  asked for      : {0}" -f $ProbeName)
    Write-Host ("  MAPI reported  : {0}" -f $script:mapiSaw)
    if ($VerifyWithOutlook) { Write-Host ("  Outlook showed : {0}" -f $script:outlookSaw) }
    if ($null -ne $script:failure) { Write-Host ("  failure        : {0}" -f $script:failure) }
    Write-Host ("  VERDICT        : {0}" -f $script:verdict)
    Write-Host '--------------------------------------------------------------------'
    Write-Host ''

    switch ($script:verdict) {
        'ACCEPTED' {
            Write-Host "Outlook accepts '@' in a store display name. The hub and identity stores can be"
            Write-Host 'named after their SMTP addresses as the tests require. Record this against'
            Write-Host 'Docs/live-tier-on-the-vm.md section 8 item 2 and close it.'
        }
        'REJECTED' {
            Write-Host "Outlook REFUSED the name. The hub and identity stores cannot be named after their"
            Write-Host 'addresses, which means testHubStoreDisplayName can no longer double as an SMTP'
            Write-Host 'address - a test-side change, not a machine one. Docs/live-tier-on-the-vm.md'
            Write-Host 'section 9 records that doubling as a known limit; this is the day it bites.'
        }
        'TRANSFORMED' {
            Write-Host 'THE DANGEROUS OUTCOME. The store exists under a name nobody asked for. Every test'
            Write-Host 'that looks a store up by display name would miss it, on a machine that otherwise'
            Write-Host 'looks correctly built. Do not build the real stores until this is understood.'
        }
        default {
            Write-Host 'The probe could not read the name back at all, so it proved nothing. That is not'
            Write-Host 'a pass - treat the question as still open.'
        }
    }

    if ($script:cleanupProblems.Count -gt 0) {
        Write-Host ''
        Write-Warning 'CLEANUP DID NOT COMPLETE:'
        foreach ($problem in $script:cleanupProblems) { Write-Warning "  $problem" }
        exit 2
    }

    Write-Host ''
    Write-Host 'Probe profile and probe PST removed.'
    if ($script:verdict -ne 'ACCEPTED') { exit 1 }
    return
}

# =============================================================================================
# ADD MODE
# =============================================================================================
$resolved = Resolve-PstPath -Path $Path

Write-Host "profile      : $ProfileName"
Write-Host "display name : $DisplayName"
Write-Host "pst          : $resolved"
Write-Host ''

if ($DisplayName -match '@' ) {
    Write-Host 'This name contains an @. If -NameProbe has not been run on this guest yet, run it'
    Write-Host 'first: it costs one command and it is what tells you whether this can work at all.'
    Write-Host ''
}

if (-not $Execute) {
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    return
}

Assert-TestbedGuest -ExpectedUser $ExpectedUser
Assert-OutlookNotRunning

Invoke-WithProfAdmin -Body {
    param($admin)

    $profiles = Get-MapiProfile -ProfAdmin $admin
    $exists = $false
    foreach ($p in $profiles) {
        if ($p.Name -eq $ProfileName) { $exists = $true }
    }
    if (-not $exists) {
        $names = @()
        foreach ($p in $profiles) { $names += $p.Name }
        throw "No profile named '$ProfileName'. Profiles on this machine: $($names -join ', '). Create it with New-OutlookProfile.ps1 first."
    }

    $serviceAdmin = $null
    Assert-MapiOk -HResult $admin.AdminServices($ProfileName, $null, [IntPtr]::Zero, 0, [ref] $serviceAdmin) -What "AdminServices('$ProfileName')"
    try {
        $services = Get-MapiService -ServiceAdmin $serviceAdmin
        foreach ($service in $services) {
            if ($service.DisplayName -eq $DisplayName) {
                Write-Host "'$DisplayName' is already a store in this profile - nothing to do."
                return
            }
        }

        [void](Add-MapiPstService -ServiceAdmin $serviceAdmin -PstPath $resolved -DisplayName $DisplayName)

        $after = Get-MapiService -ServiceAdmin $serviceAdmin
        $names = @()
        foreach ($service in $after) { $names += $service.DisplayName }
        if ($names -notcontains $DisplayName) {
            throw @"
The store was added and the service table does NOT hold '$DisplayName'.
It holds: $($names -join ', ')

If one of those is a near-miss of the name asked for, the provider renamed the store - run
-NameProbe to characterise it properly before building anything else on this profile.
"@
        }
        Write-Host "MAPI reports the store under the exact name asked for."
    }
    finally {
        if ($null -ne $serviceAdmin) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($serviceAdmin) }
    }
}

if ($VerifyWithOutlook) {
    Write-Host ''
    Write-Host 'Verifying over COM. This starts Outlook and needs session 1.'
    $view = Get-OutlookStoreView -ProfileName $ProfileName

    $byPath = $null
    foreach ($store in $view) {
        if ($null -ne $store.FilePath -and $store.FilePath -ieq $resolved) { $byPath = $store }
    }

    if ($null -eq $byPath) {
        $lines = @()
        foreach ($store in $view) { $lines += ("  {0}  <-  {1}" -f $store.DisplayName, $store.FilePath) }
        throw "Outlook does not report any store whose FilePath is $resolved. What it does report:`n$($lines -join "`n")"
    }
    if ($byPath.DisplayName -cne $DisplayName) {
        throw "Outlook shows this store as '$($byPath.DisplayName)', not '$DisplayName'. MAPI and Outlook disagree, and Outlook is what the tests read. Run -NameProbe."
    }
    Write-Host "  Outlook shows '$($byPath.DisplayName)' at $($byPath.FilePath) - verified."
}

Write-Host ''
Write-Host 'Done. Remember the two DECLARATIONS this store may need in the settings file:'
Write-Host '  expectedStoreDisplayNames   - censuses it. Every store needs this.'
Write-Host '  bystanderStoreDisplayNames  - refuses every write to it. Corpus and bystander stores'
Write-Host '                                need it; the HUB and the IDENTITY store must NOT have it.'
Write-Host 'Naming a store in exactly one of the two REFUSES the tier - see Testbed/README.md section 3b.'
