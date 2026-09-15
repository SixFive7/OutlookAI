<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it. Verified by PARSING only. Replace this banner with
    what it actually did once it has run on a guest.

    WHAT IT DOES. Gives the identity account the signature that Docs/live-tier-on-the-vm.md
    section 2.8b requires: a signature configured for NEW MAIL on a specific account, so
    LiveDraftOptionsTests can assert SignatureInjected and read the injected HTML back.

    WHAT IT DELIBERATELY DOES NOT DO: implement any of that. This product already ships it, tested:

      * SignatureManager writes the file set - .htm + .txt + .rtf - under
        %APPDATA%\Microsoft\Signatures, deriving whichever renditions it was not given. All three
        are always written because each mail format reads only its own and silently omits the
        signature when it is missing.
      * ProfileSignatureDefaultsStore writes the per-account binding as REG_SZ into 'New Signature'
        and 'Reply-Forward Signature' under the account's subkey of
        9375CFF0413111d3B88A00104B2A6676. Its writes are narrow on purpose: only those two value
        names, only on subkeys carrying an SMTP-shaped 'Account Name', never creating a subkey.
      * Both are reachable through the shipped MCP tool manage_signature, whose set_default_for
        argument takes { account, scope } with scope 'new' | 'reply' | 'both'.

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

    ONE THING ALREADY SETTLED, so nobody goes looking: on Microsoft 365 Apps 2303+ roaming
    signatures can overrule local files unless DisableRoamingSignatures=1. The guests are Office
    LTSC 2024, where LOCAL FILES ARE AUTHORITATIVE - that is recorded in SignatureManager's own
    documentation. No roaming setting is needed here.

    IDEMPOTENT, with one wrinkle. Re-running with the same name and body is a no-op in effect, but
    manage_signature's 'update' action rewrites the files and backs the previous set up under
    %LOCALAPPDATA%\OutlookAI\signature-backups first. So a second run is safe and leaves one more
    backup directory behind. The script chooses create-or-update by asking list_signatures first.

    RUN IT IN SESSION 1 (Register-InteractiveTask.ps1). Windows PowerShell 5.1 - no ternary,
    no `??`.

.PARAMETER Account
    The account's SMTP address, exactly as list_accounts reports it.

.PARAMETER SignatureName
    What the signature is called. May not begin with OutlookAI-McpTest-.

.PARAMETER BodyText
    Plain-text body. The HTML rendition is derived from it when -BodyHtml is omitted.

.PARAMETER BodyHtml
    HTML body. Optional; derived from -BodyText when omitted.

.PARAMETER Scope
    'new' | 'reply' | 'both'. Section 2.8b needs at least 'new'.

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
    [ValidateSet('new', 'reply', 'both')] [string] $Scope = 'new',
    [string] $ServerExe = 'C:\OutlookAI-Q5\server\OutlookAI.McpServer.exe',
    [int]    $ReplyTimeoutSeconds = 300,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\OutlookMapiInterop.ps1"

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
-VMName is mandatory: three guests coexist during the changeover and nothing guesses which.
"@
}

# --- the stdio client. Same shape as Invoke-GuestMeasure.ps1, which is where it was proven. -----
$psi = New-Object Diagnostics.ProcessStartInfo
$psi.FileName = $ServerExe
$psi.WorkingDirectory = Split-Path $ServerExe -Parent
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)

$proc = [Diagnostics.Process]::Start($psi)
$stdin = New-Object IO.StreamWriter($proc.StandardInput.BaseStream, (New-Object Text.UTF8Encoding($false)))
$stdin.AutoFlush = $true

function Send-Rpc {
    param([string] $Method, $Params, $Id, [switch] $Notification)
    $msg = [ordered]@{ jsonrpc = '2.0'; method = $Method }
    if (-not $Notification) { $msg['id'] = $Id }
    if ($null -ne $Params) { $msg['params'] = $Params }
    $stdin.WriteLine(($msg | ConvertTo-Json -Depth 20 -Compress))
}

function Read-Reply {
    param($ForId)
    $deadline = (Get-Date).AddSeconds($ReplyTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($proc.HasExited -and $proc.StandardOutput.EndOfStream) { throw "server exited $($proc.ExitCode)" }
        $remaining = [int]([Math]::Max(1000, ($deadline - (Get-Date)).TotalMilliseconds))
        $task = $proc.StandardOutput.ReadLineAsync()
        if (-not $task.Wait($remaining)) { throw "read timed out for id $ForId" }
        $line = $task.Result
        if ($null -eq $line) { throw "stdout closed before id $ForId" }
        if ($line.Trim() -eq '') { continue }
        $o = $null
        try { $o = $line | ConvertFrom-Json } catch { continue }
        if (($o.PSObject.Properties.Name -contains 'id') -and ("$($o.id)" -eq "$ForId")) { return $o }
    }
    throw "no reply to id $ForId"
}

function Get-ResultJson {
    param($Reply)
    return $Reply.result.content[0].text | ConvertFrom-Json
}

try {
    Send-Rpc -Method 'initialize' -Id 1 -Params @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = 'set-signature'; version = '1.0' } }
    Read-Reply -ForId 1 | Out-Null
    Send-Rpc -Method 'notifications/initialized' -Notification

    # --- 1. the account must exist BEFORE anything is written. --------------------------------
    Send-Rpc -Method 'tools/call' -Id 10 -Params @{ name = 'list_accounts'; arguments = @{} }
    $accounts = Get-ResultJson (Read-Reply -ForId 10)

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
    Send-Rpc -Method 'tools/call' -Id 11 -Params @{ name = 'list_signatures'; arguments = @{} }
    $signatures = Get-ResultJson (Read-Reply -ForId 11)

    $existing = @()
    foreach ($s in $signatures.signatures) { $existing += $s.name }

    $action = 'create'
    foreach ($name in $existing) {
        if ($name -ieq $SignatureName) { $action = 'update' }
    }
    Write-Host "signature '$SignatureName' -> $action"

    $callArgs = @{
        action          = $action
        name            = $SignatureName
        body_text       = $BodyText
        set_default_for = @{ account = $Account; scope = $Scope }
    }
    if ($BodyHtml) { $callArgs['body_html'] = $BodyHtml }

    Send-Rpc -Method 'tools/call' -Id 12 -Params @{ name = 'manage_signature'; arguments = $callArgs }
    $written = Read-Reply -ForId 12

    if ($written.PSObject.Properties.Name -contains 'error') {
        throw "manage_signature failed: $($written.error | ConvertTo-Json -Depth 10 -Compress)"
    }

    # --- 3. verify by reading it back, not by trusting the call. ------------------------------
    Send-Rpc -Method 'tools/call' -Id 13 -Params @{ name = 'list_signatures'; arguments = @{} }
    $after = Get-ResultJson (Read-Reply -ForId 13)

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
    # list_signatures returns { signatures[], accounts[], note } - the assignments live under
    # 'accounts', each { account, newMessage, replyForward }. When nothing could be read the whole
    # list is omitted and 'note' says why, which is the unknown case handled below.
    $bound = $null
    foreach ($assignment in $after.accounts) {
        if ($assignment.account -ieq $Account) { $bound = $assignment }
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
    if ($null -ne $stdin) { $stdin.Dispose() }
    if ($null -ne $proc -and -not $proc.HasExited) {
        # The server exits when its stdin closes. Give it a moment rather than killing it: it owns
        # a COM host that is attached to Outlook.
        [void]$proc.WaitForExit(30000)
    }
}
