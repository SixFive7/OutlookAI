#Requires -Version 5.1
<#
.SYNOPSIS
    Take, renew or release a lease on a testbed VM so the idle-saver leaves it alone.

.DESCRIPTION
    THE POINT OF THIS FILE. The testbed VMs are SAVED (not paused) whenever nobody is using
    them, so the host gets its RAM and CPU back. Something has to decide "nobody is using it",
    and the obvious answer - watch the guest's CPU - is wrong here in a way that would be found
    as a mystery rather than as a bug.

    A live tier run is roughly 27 minutes of driving Outlook through COM, and Outlook spends a
    great deal of that waiting: on a store to open, on a folder to enumerate, on a save to
    commit. A guest that looks idle for two minutes in the middle of a run is entirely
    ordinary. Saving it there suspends the run mid-COM-call, and what the operator sees
    afterwards is a test that timed out for no reason on a machine that looks fine.

    So usage is DECLARED, not inferred. Anything intending to use a VM takes a lease with an
    expiry and renews it while it works. The idle-saver refuses to touch a VM holding a live
    lease. A lease whose holder died simply stops protecting the VM when it expires, which is
    the right failure - the alternative is a VM pinned awake for ever by a process that no
    longer exists.

    Leases live in a fixed machine-wide directory, not in the repository (they are machine
    state, not source) and not in a per-user profile (the saver runs as a scheduled task and
    has to read the same leases the operator writes).

.PARAMETER VMName
    The VM to lease.

.PARAMETER Minutes
    How long the lease lasts from now. Renew before it expires. Do NOT take a very long lease
    to avoid renewing: a crashed holder then pins the VM awake for that long.

.PARAMETER Release
    Drop the lease now instead of taking one. Do this when the work finishes, so the saver can
    reclaim the resources without waiting for the expiry.

.PARAMETER Reason
    Free text recorded in the lease, so an operator who finds a VM awake can see why.

.EXAMPLE
    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Minutes 45 -Reason 'live tier'
    # ... run the tier, renewing every few minutes ...
    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Release
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $VMName,
    [int] $Minutes = 30,
    [switch] $Release,
    [string] $Reason = ''
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'TestbedLeasePath.ps1')
$leaseDir = Get-TestbedLeaseDirectory

# The VM name is operator-supplied and becomes a filename; keep it to a safe shape rather than
# trusting it.
if ($VMName -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    throw "VM name '$VMName' is not a plain name; refusing to build a lease path from it."
}
$leasePath = Join-Path $leaseDir "$VMName.lease.json"

if ($Release) {
    if (Test-Path -LiteralPath $leasePath) {
        Remove-Item -LiteralPath $leasePath -Force
        Write-Output "Released the lease on $VMName."
    } else {
        Write-Output "No lease on $VMName to release."
    }
    return
}

if ($Minutes -lt 1 -or $Minutes -gt 480) {
    throw "A lease of $Minutes minutes is outside the sane range (1-480). A very long lease pins the VM awake if its holder dies."
}

$expires = [DateTime]::UtcNow.AddMinutes($Minutes)
[pscustomobject]@{
    vmName     = $VMName
    takenUtc   = [DateTime]::UtcNow.ToString('o')
    expiresUtc  = $expires.ToString('o')
    # The number is what the saver compares. ConvertFrom-Json silently coerces an ISO-8601
    # STRING into a DateTime, and re-parsing that object's local rendering loses the UTC
    # marker - which made every lease read as already expired, two hours in the past.
    # A Unix second count cannot be coerced into anything but a number.
    expiresUnix = [long]([DateTimeOffset]::new($expires, [TimeSpan]::Zero).ToUnixTimeSeconds())
    holderPid  = $PID
    holderHost = $env:COMPUTERNAME
    reason     = $Reason
} | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $leasePath -Encoding UTF8

Write-Output "Leased $VMName until $($expires.ToString('u')) ($Minutes min)$(if ($Reason) { " - $Reason" })."
