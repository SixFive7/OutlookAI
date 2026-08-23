#Requires -Version 5.1
<#
.SYNOPSIS
    Gets results, logs and the corpus manifest OUT of the guest.

.DESCRIPTION
    How results leave the guest was one of the unrecorded items in the runbook, and its absence
    cost real work: on 2026-08-24 the corpus manifest was found on the guest at a path nobody had
    written down, having been assumed lost. Without the manifest the corpus cannot be torn down
    at all - `corpus-teardown` deletes only what the manifest records, by EntryID allowlist AND
    subject tag, and there is no second route that the mailbox-safety rules permit.

    So the default set this script pulls is exactly the set whose loss hurts:

        corpus-*.jsonl      the manifest. THE ONLY THING THAT CAN REMOVE THE CORPUS.
        measure.jsonl       the measurement transcript
        *.log               build and run logs

    WHERE THEY LAND, and why not in the repository: the gitignored live-fixtures directory.
    A manifest is megabytes of EntryIDs describing one machine's mailbox state, and measurement
    output is statistics about one machine that this repository refuses to carry
    (.github/scripts/check-measurement-privacy.ps1 fails a build over it). The PARAMETERS that
    reproduce the corpus are committed instead, in Testbed/testbed.json - four values, and the
    manifest is regenerable from a build.

    Pull the manifest after every corpus build and after every re-anchor. A re-anchor appends a
    replacement line per item, so an old copy is not equivalent.

.PARAMETER VMName
    Guest name.

.PARAMETER GuestPath
    Directory on the guest to collect from. Default C:\OutlookAI-Q5, which is where the guest's
    tooling lives.

.PARAMETER Include
    Filename patterns to pull.

.PARAMETER Destination
    Host directory. Defaults to the gitignored live-fixtures/vm-corpus.

.EXAMPLE
    pwsh -File Testbed/host/Copy-FromGuest.ps1
    pwsh -File Testbed/host/Copy-FromGuest.ps1 -Include *.jsonl -Destination C:\somewhere\else
#>
[CmdletBinding()]
param(
    [string]   $VMName = 'OutlookAI-TestVM',
    [string]   $GuestPath = 'C:\OutlookAI-Q5',
    [string[]] $Include = @('corpus-*.jsonl', 'measure.jsonl', '*.log'),
    [string]   $Destination,
    [string]   $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = 'Stop'

if (-not $Destination) {
    $Destination = Join-Path $RepoRoot 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\vm-corpus'
}
New-Item -ItemType Directory -Force -Path $Destination | Out-Null

$cred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$session = New-PSSession -VMName $VMName -Credential $cred
try {
    $files = Invoke-Command -Session $session -ScriptBlock {
        param($root, $patterns)
        if (-not (Test-Path -LiteralPath $root)) { return @() }
        Get-ChildItem -LiteralPath $root -File |
            Where-Object { $n = $_.Name; ($patterns | Where-Object { $n -like $_ }).Count -gt 0 } |
            ForEach-Object { [pscustomobject]@{ FullName = $_.FullName; Name = $_.Name; Length = $_.Length } }
    } -ArgumentList $GuestPath, $Include

    if (-not $files -or $files.Count -eq 0) {
        Write-Warning "Nothing matched $($Include -join ', ') under $GuestPath on $VMName."
        Write-Warning 'If you expected a manifest here, look wider before concluding it is lost:'
        Write-Warning '  Get-ChildItem C:\ -Recurse -Filter corpus-*.jsonl -ErrorAction SilentlyContinue'
        return
    }

    foreach ($f in $files) {
        $target = Join-Path $Destination $f.Name
        Write-Host ("  {0,-28} {1,10:N0} bytes" -f $f.Name, $f.Length)
        Copy-Item -LiteralPath $f.FullName -Destination $target -FromSession $session -Force
    }
}
finally {
    Remove-PSSession $session
}

Write-Host ''
Write-Host "Collected into $Destination (gitignored)."
Write-Host 'The manifest is the only thing that can tear the corpus down. Keep a copy off this machine too.'
