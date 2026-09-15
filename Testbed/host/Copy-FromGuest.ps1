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

    THE DESTINATION IS SHARED, AND THAT IS SAFE ONLY BECAUSE THE NAMES DIFFER. A manifest is named
    corpus-<corpusId>.jsonl after the CORPUS, not after the guest, so two guests land on top of
    each other the moment they share a corpus id. They must not: each guest gets its own id -
    vm-indexed and vm-unindexed for the two being built - because the two corpora genuinely are
    different populations on different machines with different index state, and they already have
    to be told apart in testbed.json, in each machine's settings file and in every measurement.
    See Testbed/README.md section 3.

    The guard below is the backstop for the day somebody forgets: an existing manifest is never
    silently replaced with different content.

.PARAMETER VMName
    MANDATORY. Which guest to pull from. There is no default: THREE MACHINES COEXIST during the
    changeover - OutlookAI-Indexed, OutlookAI-Unindexed and the outgoing OutlookAI-TestVM - and a
    default that silently picks one of three is the exact shape of mistake this testbed keeps
    making. Here it would be pulling one guest's manifest and landing it on top of another's.

.PARAMETER GuestPath
    Directory on the guest to collect from. Default C:\OutlookAI-Q5, which is where the guest's
    tooling lives.

.PARAMETER Include
    Filename patterns to pull.

.PARAMETER Destination
    Host directory. Defaults to the gitignored live-fixtures/vm-corpus, and it is SHARED by every
    guest - see the note above on why that is safe and what makes it unsafe.

.PARAMETER Force
    Overwrite a manifest that is already here and differs. Without it the script refuses, because
    the file it would replace may be the only allowlist that can remove a corpus from a real
    store. Pass it when you know the incoming copy is the newer one for the SAME corpus - after a
    rebuild, or after a re-anchor appended its replacement lines.

.EXAMPLE
    pwsh -File Testbed/host/Copy-FromGuest.ps1 -VMName OutlookAI-Indexed
    pwsh -File Testbed/host/Copy-FromGuest.ps1 -VMName OutlookAI-Unindexed
    pwsh -File Testbed/host/Copy-FromGuest.ps1 -VMName OutlookAI-Indexed -Include *.jsonl -Destination C:\somewhere\else
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $VMName,
    [string]   $GuestPath = 'C:\OutlookAI-Q5',
    [string[]] $Include = @('corpus-*.jsonl', 'measure.jsonl', '*.log'),
    [string]   $Destination,
    [switch]   $Force,
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
            ForEach-Object {
                # Hashed on the guest, and only for manifests: it is the one file here whose
                # replacement is unrecoverable, and comparing hashes is what lets the host refuse
                # BEFORE the copy rather than after it.
                $sha = $null
                if ($_.Name -like 'corpus-*.jsonl') {
                    $sha = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                }
                [pscustomobject]@{ FullName = $_.FullName; Name = $_.Name; Length = $_.Length; Sha256 = $sha }
            }
    } -ArgumentList $GuestPath, $Include

    if (-not $files -or $files.Count -eq 0) {
        Write-Warning "Nothing matched $($Include -join ', ') under $GuestPath on $VMName."
        Write-Warning 'If you expected a manifest here, look wider before concluding it is lost:'
        Write-Warning '  Get-ChildItem C:\ -Recurse -Filter corpus-*.jsonl -ErrorAction SilentlyContinue'
        return
    }

    foreach ($f in $files) {
        $target = Join-Path $Destination $f.Name
        if ($f.Sha256 -and -not $Force -and (Test-Path -LiteralPath $target) -and
            (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $f.Sha256) {
            throw @"
REFUSED: $target already exists here and holds DIFFERENT content from the copy on $VMName.

That file is a corpus manifest - the EntryID allowlist 'corpus-teardown' requires, and the file
'corpus-verify' reads to decide whether a corpus is still measurable. Teardown deletes only what a
manifest records, by EntryID AND ordinal subject tag, both required; there is no second route the
mailbox-safety rules permit. Replace the wrong one and the corpus it described becomes items in a
real store that nothing is entitled to remove, and the loss is silent - a manifest that is 2.9 MB
of plausible EntryIDs for the wrong machine looks exactly like the right one.

Stopping costs you one command. Overwriting costs a store nobody can clean up.

Two manifests with the same name means two corpora share a corpus id. They should not: give each
guest its own id (vm-indexed, vm-unindexed - Testbed/README.md section 3), rebuild with it, and
the names stop colliding. If this really is a newer copy of the SAME corpus - after a rebuild, or
a re-anchor - move the old one aside and re-run, or pass -Force.
"@
        }

        Write-Host ("  {0,-28} {1,10:N0} bytes" -f $f.Name, $f.Length)
        Copy-Item -LiteralPath $f.FullName -Destination $target -FromSession $session -Force
    }
}
finally {
    Remove-PSSession $session
}

Write-Host ''
Write-Host "Collected into $Destination (gitignored), which every guest shares."
Write-Host 'The manifest is the only thing that can tear the corpus down. Keep a copy off this machine too.'
Write-Host 'Only manifests are name-distinct per guest, and only because each guest has its own corpus id.'
Write-Host 'measure.jsonl and the logs are NOT: pulling a second guest overwrites the first guest''s copies.'
