#Requires -Version 5.1
<#
.SYNOPSIS
    Where testbed VM leases live. Dot-sourced by the lease script and the idle-saver so the two
    cannot disagree about it.

.DESCRIPTION
    A fixed machine-wide directory, deliberately, and not either of the two obvious
    alternatives:

      - NOT in the repository. Leases are machine state with a holder PID in them; committing
        one would be meaningless on any other machine and the repo is public.
      - NOT in a per-user profile. The saver runs as a scheduled task and has to read the same
        leases an operator writes from an interactive shell; a per-user path would give them
        two different views and the saver would suspend a VM someone was using.

    Nothing here is secret: a lease says a VM is in use, by which process, until when, and why.
#>

function Get-TestbedLeaseDirectory {
    [CmdletBinding()]
    param()

    $dir = Join-Path $env:SystemDrive 'OutlookAI-Testbed\leases'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return $dir
}

function Get-TestbedLease {
    <#
    .SYNOPSIS
        The live lease for a VM, or $null. An expired or unreadable lease counts as absent.
    .DESCRIPTION
        Unreadable counts as absent on purpose. The alternative - treating a corrupt lease as
        live - would let one truncated write pin a VM awake indefinitely, and the cost of
        getting it wrong this way is only that a VM gets saved while someone is using it, which
        is visible and recoverable. The cost the other way is invisible.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $VMName)

    $path = Join-Path (Get-TestbedLeaseDirectory) "$VMName.lease.json"
    if (-not (Test-Path -LiteralPath $path)) { return $null }

    try {
        $lease = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        # Compare on the epoch number, never on expiresUtc. ConvertFrom-Json turns an
        # ISO-8601 string into a DateTime of Kind Unspecified; re-parsing that object's
        # rendering drops the UTC marker, and ToUniversalTime() then treats a UTC instant as
        # local and shifts it by the offset. Measured on this machine: every lease read as
        # expired two hours before it was written.
        $expiresUnix = [long]$lease.expiresUnix
    } catch {
        Write-Verbose "Lease for $VMName is unreadable ($($_.Exception.Message)); treating as absent."
        return $null
    }

    if ($expiresUnix -le [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) { return $null }
    return $lease
}
