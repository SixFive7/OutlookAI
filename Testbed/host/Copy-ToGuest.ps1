#Requires -Version 5.1
<#
.SYNOPSIS
    Copies a file into the guest over PowerShell Direct.

.DESCRIPTION
    PowerShell Direct needs no network on the guest, which is the point: the testbed's guest is
    meant to have no route to the outside once the toolchain is in, and a file transfer that
    needed one would undo that.

    This lands in SESSION 0. That is fine for what this does - copying bytes touches no COM - but
    it is why nothing here starts Outlook or runs a corpus verb. Those go through
    Testbed/guest/Register-InteractiveTask.ps1.

    `Copy-VMFile` is deliberately not used. It needs the Guest Service Interface integration
    service enabled, it is host-to-guest only, and it gives no error worth reading when the
    service is off. A PSSession over VMBus does both directions with one mechanism.

.PARAMETER VMName
    Guest name. `OutlookAI-TestVM` by convention.

.PARAMETER Path
    Host file to copy. Repeatable.

.PARAMETER Destination
    Guest path. A directory when several files are given; a file path when one is.

.EXAMPLE
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -Path .work/testbed-payload/McpServer.zip -Destination C:\OutlookAI-Q5\McpServer.zip
#>
[CmdletBinding()]
param(
    [string] $VMName = 'OutlookAI-TestVM',
    [Parameter(Mandatory = $true)] [string[]] $Path,
    [Parameter(Mandatory = $true)] [string] $Destination,
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = 'Stop'

$resolved = @()
foreach ($p in $Path) {
    $item = Resolve-Path -LiteralPath $p -ErrorAction SilentlyContinue
    if (-not $item) { throw "Not found on the host: $p" }
    $resolved += $item.Path
}

$cred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$session = New-PSSession -VMName $VMName -Credential $cred
try {
    $destDir = Split-Path -Parent $Destination
    if ($resolved.Count -gt 1) { $destDir = $Destination }
    if ($destDir) {
        Invoke-Command -Session $session -ScriptBlock {
            param($d)
            if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
        } -ArgumentList $destDir
    }

    foreach ($p in $resolved) {
        $target = $Destination
        if ($resolved.Count -gt 1) { $target = Join-Path $Destination (Split-Path -Leaf $p) }
        Write-Host "  -> $target"
        Copy-Item -LiteralPath $p -Destination $target -ToSession $session -Force
    }
}
finally {
    Remove-PSSession $session
}

Write-Host "Copied $($resolved.Count) file(s) into $VMName."
