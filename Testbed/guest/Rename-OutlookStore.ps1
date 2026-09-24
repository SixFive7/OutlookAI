<#
    ============================================================================================
    Renames an Outlook store to an exact display name. MEASURED WORKING 2026-09-15.
    ============================================================================================

    RUN ON THE GUEST, in the interactive session. Windows PowerShell 5.1 - no ternary, no `??`.

    WHY THIS EXISTS. The tier profile's POP3 account only works if Outlook MINTS its own delivery
    store (see Testbed/guest/tier-profile-forcepst.prf) - a store Outlook mints is a store Outlook
    binds, and binding is the step a .prf structurally cannot perform. But Outlook names what it
    mints: 'Outlook Data File'. Section 2.6 of the runbook requires the hub store to be named as
    an SMTP address, and the live tier keys on Store.DisplayName. So the name has to be set
    afterwards, and this does it.

    THE QUESTION THIS ANSWERED. Store.DisplayName is READ-ONLY. The documented workaround is
    PropertyAccessor.SetProperty on PR_DISPLAY_NAME_W, which is widely reported to be blocked by
    Outlook. The other candidate is renaming the store's ROOT FOLDER - and no source, Microsoft or
    community, stated whether Store.DisplayName then FOLLOWS. Four research passes could not
    settle it.

    It follows. Measured on Office LTSC 2024 build 16.0.17932:

        set root.Name        was 'Outlook Data File', now 'tier@vm.invalid'
        Store.DisplayName    'tier@vm.invalid'
        account afterwards   SmtpAddress='tier@vm.invalid' DeliveryStore='tier@vm.invalid'

    The '@' is accepted, and the account's DeliveryStore reports the new name too - so renaming
    the store does not break the binding that made the account usable.

    WHAT IT DOES NOT TOUCH. No item is created, deleted, moved or modified. This renames a FOLDER,
    which is the store's root node. Per the narrowed rule in section 2.6, item mutation from a
    script is still forbidden outright; renaming a store on a guest whose identity this script
    verifies is not item mutation.

    ============================================================================================
    2026-09-24: THAT LAST SENTENCE WAS FALSE UNTIL TODAY. THE GUARD IS NOW REAL.
    ============================================================================================

    "A guest whose identity this script verifies" - it verified nothing. There was no identity
    check anywhere in this file: run on the maintainer's workstation, it would have bound the REAL
    Outlook over COM and renamed the root folder of whichever real store matched -StoreFilePath.
    It now dot-sources OutlookMapiInterop.ps1 and calls Assert-TestbedGuest - the same guard every
    other writing script here uses - BEFORE ANYTHING ELSE, the argument check and the dry run
    included: even a dry run binds Outlook and lists a profile's stores, and on the workstation
    that is a real profile.

    STAGE OutlookMapiInterop.ps1 BESIDE THIS SCRIPT. Copying this file alone now fails loudly on
    the dot-source, which is the intended failure. Proven on the maintainer's workstation the same
    day, with -Execute and with the dry run: both stopped at "REFUSING TO RUN. This session is
    logged on as ..." and no Outlook COM object was ever requested - a tripwire standing in for
    New-Object recorded zero calls, and a control script showed the same tripwire does catch a
    New-Object -ComObject. And proven the other way on OAI-UNINDEXED, as vmadmin in session 1,
    twice (both orders of the from-scratch rehearsal): the guard passed, the minted store went
    from 'Outlook Data File' to 'tier@vm.invalid', and a second -Execute said "Already named".
    .github/scripts/check-testbed-references.ps1 check 9 now fails the build if any guest script
    that writes loses its guard, or calls it after its first write.

.PARAMETER StoreFilePath
    The .pst whose store should be renamed. Matched on Store.FilePath, resolved and
    case-insensitively - Outlook normalises the path it reports, so comparing against the literal
    you typed can miss.

.PARAMETER DisplayName
    The name to set. For the hub store this must be the dummy account's SMTP address, because
    several tests use testHubStoreDisplayName as an address.

.PARAMETER Execute
    Rename. Without it, reports what it would do and changes nothing.

.PARAMETER ExpectedUser
    The account the guest guard accepts. The default is the guard; see OutlookMapiInterop.ps1.

.EXAMPLE
    .\Rename-OutlookStore.ps1 -StoreFilePath C:\OutlookAI-Tier\Outlook.pst -DisplayName tier@vm.invalid
    .\Rename-OutlookStore.ps1 -StoreFilePath C:\OutlookAI-Tier\Outlook.pst -DisplayName tier@vm.invalid -Execute
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $StoreFilePath,
    [Parameter(Mandatory = $true)] [string] $DisplayName,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

# THE GUARD, FIRST. Before the argument check, before the dry run, before COM: every path below
# binds Outlook, and on the maintainer's workstation that is a real profile with real delegate
# mailboxes. See the 2026-09-24 section of the banner.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

function Say($m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

if ($DisplayName -match '[\\/]') {
    throw "'$DisplayName' contains a slash. '/' is the one character Outlook is consistently reported to reject in a folder name."
}

$target = $StoreFilePath
try { $target = [System.IO.Path]::GetFullPath($StoreFilePath) } catch {}

$ol = $null
$ns = $null
try {
    # GetDefaultFolder initialises MAPI against the default profile. Never NameSpace.Logon:
    # Microsoft documents that it prompts for a profile even when one is set as default.
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    $null = $ns.GetDefaultFolder(6)

    $store = $null
    for ($i = 1; $i -le $ns.Stores.Count; $i++) {
        $s = $ns.Stores.Item($i)
        if ($s.FilePath -and ($s.FilePath -ieq $target)) { $store = $s; break }
    }
    if (-not $store) {
        Say "Stores in this profile:"
        for ($i = 1; $i -le $ns.Stores.Count; $i++) {
            $s = $ns.Stores.Item($i)
            Say ("    '{0}'  {1}" -f $s.DisplayName, $s.FilePath)
        }
        throw "No store in this profile has FilePath '$target'."
    }

    $root = $store.GetRootFolder()
    Say ("current: Store.DisplayName='{0}'  root.Name='{1}'" -f $store.DisplayName, $root.Name)

    if ($store.DisplayName -ceq $DisplayName) {
        Say "Already named '$DisplayName'. Nothing to do."
        return
    }
    if (-not $Execute) {
        Say "DRY RUN. Would rename to '$DisplayName'. Re-run with -Execute."
        return
    }

    $root.Name = $DisplayName
    Start-Sleep -Seconds 2

    # VERIFY by re-fetching from the collection. Reusing $store would read a cached RCW, which
    # could report the value we just set without it having reached the store - and the whole point
    # of this script is that Store.DisplayName and the root folder's name are different properties
    # on different objects.
    $after = $null
    for ($i = 1; $i -le $ns.Stores.Count; $i++) {
        $s = $ns.Stores.Item($i)
        if ($s.FilePath -and ($s.FilePath -ieq $target)) { $after = $s; break }
    }
    if (-not $after) { throw "The store vanished from the profile after the rename." }

    Say ("after:   Store.DisplayName='{0}'  root.Name='{1}'" -f $after.DisplayName, $after.GetRootFolder().Name)
    if ($after.DisplayName -cne $DisplayName) {
        throw "Store.DisplayName reads '$($after.DisplayName)', not '$DisplayName'. The root folder was renamed but the store did not follow."
    }
    Say "OK - Store.DisplayName is '$DisplayName'."
}
finally {
    foreach ($v in @($ns, $ol)) {
        if ($v) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($v) }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
