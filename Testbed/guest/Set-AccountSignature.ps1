<#
    ============================================================================================
    RUN THREE TIMES: WRONG ENTRY WITH THE OLD SERVER (2026-09-24), THE ACCOUNT WITH THE FIXED ONE
    (2026-09-27, on both guests). READ WHAT "Verified" DOES AND DOES NOT MEAN BEFORE TRUSTING IT.
    ============================================================================================

      OutlookAI-Unindexed, 2026-09-27, from CP-10B-IDENTITY-REAL-INBOX (the identity store minted,
        Q87 (a)) with a server published from af56efc: Outlook started NOT elevated
        (Start-OutlookUnelevated.ps1), this script in a RunLevel Limited task. "Verified", and
        right: every account entry read '' before; afterwards 'New Signature' = 'Identity' on
        00000004 (the POP3 account, clsid {ED475411-...}, Email identity@vm.invalid) and nothing on
        00000005 (the data-file entry named identity@vm.invalid) or anywhere else; Identity.htm, .rtf
        and .txt written; a separate read-only list_signatures call read back identity@vm.invalid
        newMessage 'Identity' and tier@vm.invalid with none. Checkpoint CP-11B-SIGNATURE-SDK-SUITE
        (Docs/live-tier-on-the-vm.md section 4.1c).

      OutlookAI-Unindexed, 2026-09-24, with a server built before the manage_signature fix:
        printed "Verified" and was WRONG. The server wrote 'New Signature' onto the identity PST's
        DATA-FILE entry (profile subkey 00000005, whose 'Account Name' is the store's name
        identity@vm.invalid), not onto the POP3 account (00000004), and list_signatures read it
        back from the same wrong place (Docs/live-tier-on-the-vm.md section 2.8b).
      OutlookAI-Indexed, 2026-09-27, with a server published from 4e23866, which carries the fix:
        "Verified", and right. 'New Signature' = 'Identity' landed on 00000004 (clsid
        {ED475411-...}, Email identity@vm.invalid), nothing on 00000005; Identity.htm, .rtf and
        .txt were written; a separate list_signatures call afterwards read back account
        identity@vm.invalid, newMessage 'Identity'. Both subkeys were read straight from the
        registry before and after (Docs/live-tier-on-the-vm.md section 4.2).

    WHAT "Verified" MEANS: the product's own read-back agrees with the product's own write - which
    the first run shows can be wrong together. Read the account's subkey yourself. Whether Outlook
    then INSERTS the signature into a new mail from that account is the live tier's question
    (LiveDraftOptionsTests), not this script's, and nothing here creates an item to find out.

    RUN IT AT OUTLOOK'S INTEGRITY LEVEL. The MCP server attaches to the running Outlook over COM
    (list_accounts). On the indexed guest Outlook runs NOT elevated (Start-OutlookUnelevated.ps1),
    so this ran in a RunLevel Limited task in session 1: an elevated caller cannot reach a
    non-elevated Outlook (Testbed/host/Restart-Guest.ps1). Register-InteractiveTask.ps1 -RunLevel
    Limited is that route now; its default (Highest) is for an Outlook the testbed started elevated.

    WHAT IT DOES. Gives the identity account the signature that Docs/live-tier-on-the-vm.md
    section 2.8b requires: a signature configured for NEW MAIL on a specific account, so
    LiveDraftOptionsTests can assert SignatureInjected and read the injected HTML back.

    WHAT IT DELIBERATELY DOES NOT DO: implement any of that. This product already ships it, tested:

      * SignatureManager writes the file set - .htm + .txt + .rtf - in the signature folder Outlook
        records (Common\General\Signatures, %APPDATA%\Microsoft\Signatures when none is), deriving
        whichever renditions a CREATE was not given. All three are written because each mail format
        reads only its own and silently omits the signature when it is missing.
      * ProfileSignatureDefaultsStore writes the per-account binding as REG_SZ into 'New Signature'
        and 'Reply-Forward Signature' under the account's subkey of
        9375CFF0413111d3B88A00104B2A6676. Its writes are narrow on purpose: only those two value
        names, only on subkeys carrying an SMTP-shaped 'Account Name', never creating a subkey.
      * Both are reachable through the shipped MCP tool manage_signature, whose default_for
        argument takes { profile, account, scope } with scope 'new' | 'reply' | 'both' - the profile
        always explicit since Q155 (2026-11-17), so -OutlookProfile names it.

    So this script drives the shipped tool over stdio. THAT IS THE POINT, not a shortcut: mailbox
    safety rule 1 says creating or editing anything happens only through the project's tested
    helper code or the shipped MCP tools, never improvised shell code. A hand-rolled registry poke
    here would be exactly the thing that rule exists to forbid. The JSON-RPC shape is the one
    already proven in Testbed/guest/Invoke-GuestMeasure.ps1.

    ORDERING - THIS WILL BITE. The binding is written onto an ACCOUNT's registry subkey, so THE
    ACCOUNT MUST ALREADY EXIST. That means after the GUI pass that adds the POP3 accounts, not
    before. There is no free programmatic route to creating the account (New-PopAccountPrf.ps1's
    header sets out why, three ways), so the order is forced:

        New-OutlookProfile.ps1 -> Add-OutlookPstStore.ps1 -> Set-DefaultOutlookProfile.ps1
        -> THE GUI PASS (two POP3 accounts + delivery stores) -> checkpoint -> THIS

    Run it earlier and manage_signature writes the files and finds no account to bind them to.
    This script checks for the account FIRST and refuses, rather than half-succeeding.

    DO NOT PREFIX THIS SIGNATURE 'OutlookAI-McpTest-'. Section 2.8b is explicit: the identity
    signature is ORDINARY USER DATA on this machine, not one of the signatures the suite creates
    and deletes. The suite's SHA-256 signature-directory snapshot requires the user's real
    signatures to come back bit-identical, which they will, because nothing in the suite writes to
    them - but only as long as this one is not named as if it were a test artefact. The script
    refuses that prefix.

    ROAMING, so nobody goes looking: on Microsoft 365 Apps 2303+ roaming signatures can overrule
    local files unless DisableRoamingSignatures=1. Whether Office LTSC 2024 roams is not measured
    (Q147 R1), so the product treats it as if it may - which matters only for an Exchange account,
    whose defaults it then leaves to Outlook. The identity account is POP3: nothing roams.

    IDEMPOTENT, with one wrinkle. Re-running with the same name and body is a no-op in effect, but
    manage_signature's 'update' action rewrites the plain-text rendition and backs the previous set
    up under %LOCALAPPDATA%\OutlookAI\signature-backups first. So a second run is safe and leaves
    one more backup directory behind. The script chooses create-or-update by asking list_signatures
    first, and sets the default with the default_for of that same call.

    RUN IT IN SESSION 1, AT OUTLOOK'S INTEGRITY LEVEL (above): Register-InteractiveTask.ps1, with
    -RunLevel Limited for a non-elevated Outlook. Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER Account
    The account's SMTP address, exactly as list_accounts reports it.

.PARAMETER SignatureName
    What the signature is called. May not begin with OutlookAI-McpTest-.

.PARAMETER BodyText
    Plain-text body. A new signature's HTML and RTF renditions are derived from it when -BodyHtml and
    -BodyRtf are omitted.

.PARAMETER BodyHtml
    HTML body. Optional when the signature is new; derived from -BodyText when omitted.

.PARAMETER BodyRtf
    RTF body. Optional when the signature is new; derived from -BodyText when omitted. An update sends
    every rendition the signature has (Q205 (d)), and one this script made has all three: so a run on a
    signature that already exists takes -BodyHtml and -BodyRtf beside -BodyText, or the server refuses
    it, naming the ones left out, and nothing is written.

.PARAMETER Scope
    'new' | 'reply' | 'both'. Section 2.8b needs at least 'new'.

.PARAMETER OutlookProfile
    The Outlook profile the account belongs to - the tier profile on the guests. Explicit, as the
    product requires (Q155).

.PARAMETER ServerExe
    The MCP server, staged by Testbed/host/Publish-GuestPayload.ps1.

.PARAMETER Execute
    Without it, nothing is written.

.EXAMPLE
    .\Set-AccountSignature.ps1 -Account identity@vm.invalid -SignatureName 'Identity' -BodyText 'OutlookAI testbed identity account.' -Execute
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $Account,
    [string] $SignatureName = 'Identity',
    [string] $BodyText = 'OutlookAI testbed identity account.',
    [string] $BodyHtml,
    [string] $BodyRtf,
    [ValidateSet('new', 'reply', 'both')] [string] $Scope = 'new',
    [string] $OutlookProfile = 'OutlookAI-Tier',
    [string] $ServerExe = 'C:\OutlookAI-Q5\server\OutlookAI.McpServer.exe',
    [int]    $ReplyTimeoutSeconds = 300,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\OutlookMapiInterop.ps1"
. "$PSScriptRoot\McpStdioClient.ps1"

if ($SignatureName -like 'OutlookAI-McpTest-*') {
    throw @"
REFUSING: '$SignatureName' uses the prefix the live suite reserves for its own signatures.

Docs/live-tier-on-the-vm.md section 2.8b wants the identity signature to be ORDINARY USER DATA.
The suite creates and deletes OutlookAI-McpTest-* signatures and proves, with a SHA-256 snapshot,
that it left everything else bit-identical. Naming this one as a test artefact puts it inside the
set the suite is entitled to delete, which is the opposite of what section 2.8b asks for.
"@
}

Write-Host "account        : $Account"
Write-Host "signature      : $SignatureName"
Write-Host "scope          : $Scope"
Write-Host "profile        : $OutlookProfile"
Write-Host "server         : $ServerExe"
Write-Host ''

if (-not $Execute) {
    Write-Host 'Dry run. Nothing written. Re-run with -Execute.'
    Write-Host 'Remember the ordering: the ACCOUNT must already exist, which means after the GUI pass.'
    return
}

Assert-TestbedGuest -ExpectedUser $ExpectedUser

if (-not (Test-Path -LiteralPath $ServerExe)) {
    throw @"
The MCP server is not at $ServerExe

The guest has no .NET SDK, so nothing can be built here. Publish on the host and copy it in:
    pwsh -File Testbed/host/Publish-GuestPayload.ps1
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <the guest> -Path .work\testbed-payload\McpServer.zip -Destination C:\OutlookAI-Q5\McpServer.zip
then on the guest: Expand-Archive C:\OutlookAI-Q5\McpServer.zip -DestinationPath C:\OutlookAI-Q5\server -Force
-VMName is mandatory: several testbed VMs coexist and nothing guesses which.
"@
}

# --- the stdio client: McpStdioClient.ps1, the one copy of the shape Invoke-GuestMeasure.ps1 proved (Q154). -----
$session = $null

try {
    $session = Start-McpStdioServer -ServerExe $ServerExe -ClientName 'set-signature' -TimeoutSeconds $ReplyTimeoutSeconds

    # --- 1. the account must exist BEFORE anything is written. --------------------------------
    $accounts = Invoke-McpTool -Session $session -Tool 'list_accounts' -Id 10

    # The delivery store is printed beside each account because section 2.8b needs the identity
    # account to deliver into its OWN PST, and section 2.8 calls a wrong delivery store "the
    # failure to bet on" - it shows up only as a 180-second arrival timeout pointing nowhere. It
    # is free to read here, and this is the last step that looks at the accounts at all.
    $addresses = @()
    foreach ($a in $accounts.accounts) {
        $addresses += $a.smtpAddress
        Write-Host ("  {0,-28} delivers to: {1}" -f $a.smtpAddress, $a.deliveryStore)
    }

    if ($addresses -notcontains $Account) {
        throw @"
No account with the address '$Account' on this profile.

The signature binding is written onto an ACCOUNT's registry subkey, so the account has to exist
first. There is no free programmatic route to creating one - Testbed/guest/New-PopAccountPrf.ps1's
header sets out why, three independent ways - so add it in Outlook's Account Settings wizard,
point its delivery store at the right PST, take a checkpoint, and run this again.

Accounts this profile does have: $($addresses -join ', ')
"@
    }

    # --- 2. create or update, decided by asking rather than by guessing. ----------------------
    $signatures = Invoke-McpTool -Session $session -Tool 'list_signatures' -Id 11

    $existing = @()
    foreach ($s in $signatures.signatures) { $existing += $s.name }

    $action = 'create'
    foreach ($name in $existing) {
        if ($name -ieq $SignatureName) { $action = 'update' }
    }
    Write-Host "signature '$SignatureName' -> $action"

    $callArgs = @{
        action      = $action
        name        = $SignatureName
        body_text   = $BodyText
        default_for = @{ profile = $OutlookProfile; account = $Account; scope = $Scope }
    }
    if ($BodyHtml) { $callArgs['body_html'] = $BodyHtml }
    if ($BodyRtf) { $callArgs['body_rtf'] = $BodyRtf }

    # A refusal or an error throws, carrying what the server said (McpStdioClient.ps1, Get-McpToolResult).
    [void](Invoke-McpTool -Session $session -Tool 'manage_signature' -Arguments $callArgs -Id 12)

    # --- 3. verify by reading it back, not by trusting the call. ------------------------------
    $after = Invoke-McpTool -Session $session -Tool 'list_signatures' -Id 13

    $present = $false
    foreach ($s in $after.signatures) {
        if ($s.name -ieq $SignatureName) { $present = $true }
    }
    if (-not $present) {
        throw "manage_signature reported success and list_signatures does not show '$SignatureName'. Do not continue."
    }

    # The binding is the half that actually matters: the file set existing proves nothing about
    # whether the account will use it. list_signatures reports the per-account assignments it can
    # read out of the profile registry, and MISSING means unknown rather than absent - so an
    # unreadable answer is reported as unproven rather than quietly passed.
    # list_signatures returns every profile under 'profiles', each { name, accounts[] }, and each
    # account { account, newMessage, replyForward, ... } - read in the profile the default was set in.
    $bound = $null
    foreach ($p in $after.profiles) {
        if ($p.name -ieq $OutlookProfile) {
            foreach ($assignment in $p.accounts) {
                if ($assignment.account -ieq $Account) { $bound = $assignment }
            }
        }
    }

    if ($null -eq $bound) {
        Write-Warning "The signature exists, and no per-account assignment could be read for '$Account'."
        Write-Warning 'That is UNKNOWN, not "absent" - the reader reports missing state as unknown by design.'
        Write-Warning 'Check it in Outlook (File > Options > Mail > Signatures) before trusting this machine:'
        Write-Warning 'LiveDraftOptionsTests asserts SignatureInjected, and it will simply fail if the binding'
        Write-Warning 'is not really there.'
        exit 2
    }

    Write-Host ''
    Write-Host ("  account       : {0}" -f $bound.account)
    Write-Host ("  new message   : {0}" -f $bound.newMessage)
    Write-Host ("  reply/forward : {0}" -f $bound.replyForward)

    $newOk = ($Scope -eq 'reply') -or ($bound.newMessage -ieq $SignatureName)
    $replyOk = ($Scope -eq 'new') -or ($bound.replyForward -ieq $SignatureName)
    if (-not $newOk -or -not $replyOk) {
        throw "The assignment does not match what was asked for (scope '$Scope'). Do not continue - section 2.8b needs the NEW-mail binding specifically."
    }

    Write-Host ''
    Write-Host "Verified. Outlook picks the assignment up at its next start."
}
finally {
    # The server exits when its stdin closes. It is given a moment rather than killed: it owns a COM
    # host that is attached to Outlook.
    [void](Stop-McpStdioServer -Session $session)
}
