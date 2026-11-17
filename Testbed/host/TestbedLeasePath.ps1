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

function Enter-TestbedLeaseClaim {
    <#
    .SYNOPSIS
        An exclusive hold on a VM's lease while a caller decides whether to take it. Dispose it as
        soon as the lease is written.
    .DESCRIPTION
        Reading a lease and writing one are two steps, and Set-TestbedLease.ps1 writes over
        whatever is there. Two callers that find the same VM free in the same moment would both
        take it, and the one that wrote first would find out only when its next renewal read the
        other's lease - after it had restored a checkpoint under the other's run. With pools
        (TestbedPool.ps1, 2026-11-17) runs that ask for a kind of guest look for the same free VM,
        so the read and the write are made one step here: a file beside the lease, opened with no
        sharing, held from the read to the write. A person taking a lease by hand reads it first
        and does not need this.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [string] $VMName, [int] $TimeoutSeconds = 60)

    if ($VMName -notmatch '^[A-Za-z0-9._-]{1,64}$') { throw "VM name '$VMName' is not a plain name; refusing to build a path from it." }
    $path = Join-Path (Get-TestbedLeaseDirectory) "$VMName.claim"
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ($true) {
        try { return [System.IO.File]::Open($path, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None) }
        catch [System.IO.IOException] {
            if ((Get-Date) -ge $deadline) { throw "the lease claim on $VMName stayed held for $TimeoutSeconds s" }
            Start-Sleep -Milliseconds 200
        }
    }
}
