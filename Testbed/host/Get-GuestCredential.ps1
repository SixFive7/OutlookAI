#Requires -Version 5.1
<#
.SYNOPSIS
    Loads the guest credential from the gitignored live-fixtures directory. Never prints it.

.DESCRIPTION
    THIS FILE CONTAINS NO CREDENTIAL AND MUST NEVER CONTAIN ONE. This repository is public, and a
    guest password has already been published from it once - it is in git history and had to be
    rotated. What this script does is name the one place a credential is allowed to live and the
    shape it takes there, so that no script anywhere else has an excuse to hard-code one.

    The file:

        McpServer/OutlookAI.McpServer.Tests/live-fixtures/vm-credentials.json

    gitignored by the `McpServer/**/live-fixtures/` rule. `.github/scripts/check-testbed-references.ps1`
    asserts that rule still covers it, because an ignore rule that gets deleted is silent until
    the day something lands.

    Its shape:

        {
          "vmName":     "",
          "username":   "<the guest account>",
          "password":   "<the guest account's password>",
          "rotatedUtc": "2026-08-24T00:00:00Z",
          "rotatedWhy": "<why, so the next rotation has context>",
          "note":       "<anything a reader needs>"
        }

    `vmName` IS EMPTY, AND THAT IS THE CURRENT SETTING, NOT A GAP. The field pins a credential
    to one guest, and the match below only fires when it has a value - so an empty string means
    "usable for any guest". Three machines coexist during the changeover (OutlookAI-Indexed,
    OutlookAI-Unindexed and the outgoing OutlookAI-TestVM) and one account serves all of them,
    so a pin naming any single one of the three refuses the other two. Fill it in only when the
    credential really is for one machine alone.

    Create it by hand when you create the guest account. Set the password to NEVER EXPIRE: a
    maximum password age silently breaks the tier weeks later and recreates the whole problem.

    If you ever rotate it AFTER the dummy mail account exists, do not fall back to an admin
    reset. An admin reset destroys that account's DPAPI master key, and Outlook's saved account
    password goes with it. That fallback was free in August 2026 only because the profile had no
    mail accounts yet.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER VMName
    MANDATORY. Which guest the credential is being fetched for. When the file's `vmName` has a
    value it must match, so a credential pinned to one guest cannot be handed to another by
    accident.

    It is mandatory rather than optional because THREE MACHINES COEXIST during the changeover -
    OutlookAI-Indexed, OutlookAI-Unindexed and the outgoing OutlookAI-TestVM - and a caller that
    does not say which one it means is not making a safe request, it is making an unanswerable
    one. There is deliberately no default anywhere in Testbed/host/: a default that silently
    picks one of three is the exact shape of mistake this testbed keeps making.

.OUTPUTS
    [PSCredential]

.EXAMPLE
    $cred = & Testbed/host/Get-GuestCredential.ps1 -VMName OutlookAI-Indexed
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [Parameter(Mandatory = $true)] [string] $VMName
)

$ErrorActionPreference = 'Stop'

$path = Join-Path $RepoRoot 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\vm-credentials.json'
if (-not (Test-Path -LiteralPath $path)) {
    throw @"
No guest credential at:
    $path
That directory is gitignored and machine-local, so a fresh clone never has it. Create it with
the shape documented in the header of this script. Never put it anywhere else.
"@
}

$json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json

foreach ($field in @('username', 'password')) {
    if (-not $json.PSObject.Properties.Name.Contains($field) -or [string]::IsNullOrWhiteSpace($json.$field)) {
        throw "vm-credentials.json is missing '$field'. See the header of $PSCommandPath."
    }
}

# An EMPTY `vmName` means "usable for any guest", and that is load-bearing rather than lax: the
# `-and $json.vmName` term is what makes one credential file serve all three coexisting guests.
# The refusal still fires the moment the field is given a value, so pinning remains available to
# anyone who genuinely has one credential per machine.
if ($VMName -and $json.PSObject.Properties.Name.Contains('vmName') -and $json.vmName -and $json.vmName -ne $VMName) {
    throw @"
vm-credentials.json holds a credential pinned to '$($json.vmName)', not '$VMName'. Refusing to use it.
Three guests coexist during the changeover and one account serves all of them, so the intended
setting for that field is an empty string. Clear it, or create a per-guest credential file.
"@
}

# The username is safe to show and is worth showing: "wrong account" is otherwise indistinguishable
# from "wrong password". The password is never written to a stream, a log or a variable that
# outlives this call.
Write-Verbose "Guest credential loaded for user '$($json.username)'."

New-Object System.Management.Automation.PSCredential(
    $json.username,
    (ConvertTo-SecureString $json.password -AsPlainText -Force))
