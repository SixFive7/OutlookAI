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

    TWO GUESTS, AND THE PULL IS SPLIT SO THEY CANNOT LAND ON EACH OTHER:

        <Destination>\corpus-<corpusId>.jsonl       manifests - the SHARED root, named by CORPUS
        <Destination>\<VMName>\measure.jsonl        everything else - a directory per MACHINE
        <Destination>\<VMName>\*.log

    THE SPLIT IS DELIBERATE AND IT IS NOT AN INCONSISTENCY. Read this before tidying it, because
    it looks like two answers to one question and it is not:

      * MANIFESTS WERE NOT FIXED WITH A DIRECTORY. They were fixed by giving the two corpora
        different ids - vm-indexed and vm-unindexed - because what was wrong there was IDENTITY:
        two genuinely different populations, different machines and different index state, were
        sharing one name. A directory would have let them keep that one name and hidden a
        modelling error behind a path. The id is what appears in every subject, in the teardown
        match and in every measurement that quotes one, so the id is what had to differ.
        Testbed/README.md section 3.

      * measure.jsonl AND THE LOGS ARE THE OPPOSITE CASE. Their identity is already correct -
        there is one thing called "the measurement transcript" and this really is it. What differs
        between two copies is not WHAT they are but WHICH MACHINE PRODUCED THEM, and a directory
        is exactly what expresses that. Stamping the machine into the file name instead would be
        inventing an identity they do not have.

      * So: rename what is genuinely two different things; put in separate directories what is one
        kind of thing arriving from two different places.

    WHY THE MANIFEST STAYS IN THE SHARED ROOT while the rest of the pull moved down a level.
    Because moving it would DISARM THE GUARD BELOW. That refusal can only fire when two different
    contents arrive at one path, and a per-guest directory guarantees they never do - so two
    guests mistakenly given the same corpus id would quietly stop colliding, and the modelling
    error would go back to being silent, which is the state the id decision was made to end. The
    shared root is what keeps a reused id detectable on the host, at the moment a human is
    standing in front of it. The cost is that one pull lands in two places; that is the price of a
    backstop that can still fire.

    AN EXISTING FLAT DESTINATION FROM AN EARLIER PULL IS LEFT ALONE, AND NAMED OUT LOUD. Before
    this change everything landed flat, so a destination used before today may hold a measure.jsonl
    or a .log that this script would now write one directory down. Those files are not moved and
    not deleted: NOBODY CAN TELL WHICH GUEST WROTE THEM - that is the defect this change exists to
    fix - and filing them under a guess would turn a merely stale file into a false attribution in
    exactly the indexed-versus-unindexed comparison the two guests exist to produce. The script
    warns about them by name on every pull until they are gone. Move them yourself if you know
    where they came from; delete them if you do not.

    THE CONTENT HASH IS TAKEN FOR EVERY FILE, not only for manifests. The guest hashes before
    anything is copied, so a refusal happens INSTEAD OF the overwrite rather than after it. For a
    manifest the refusal is about a corpus in a real store that nothing would then be entitled to
    remove; for a transcript or a log it is about a measurement nobody can re-take once its
    machine has moved on. -Force means "yes, this is the newer copy of the same thing".

.PARAMETER VMName
    MANDATORY. Which guest to pull from. There is no default: THREE MACHINES COEXIST during the
    changeover - OutlookAI-Indexed, OutlookAI-Unindexed and the outgoing OutlookAI-TestVM - and a
    default that silently picks one of three is the exact shape of mistake this testbed keeps
    making. Here it would be pulling one guest's manifest and landing it on top of another's.

    It also NAMES THE SUBDIRECTORY the results land in, so it must be a single path segment. A
    Hyper-V VM name is free text and may legally contain a separator; one that does is rejected
    rather than mangled, because a mangled name silently files results under the wrong machine.

.PARAMETER GuestPath
    Directory on the guest to collect from. Default C:\OutlookAI-Q5, which is where the guest's
    tooling lives.

.PARAMETER Include
    Filename patterns to pull. Anything named corpus-*.jsonl is treated as a manifest wherever it
    came from - the classification is by FILE NAME, not by which pattern matched it, so narrowing
    -Include cannot accidentally reclassify one.

.PARAMETER Destination
    Host directory. Defaults to the gitignored live-fixtures/vm-corpus. Its ROOT is shared by
    every guest and holds manifests only; each guest's results go in <Destination>\<VMName>. See
    the note above on why those two halves are settled differently.

.PARAMETER Force
    Overwrite a file that is already here and differs. Without it the script refuses. For a
    manifest that refusal protects the only allowlist that can remove a corpus from a real store;
    for a transcript or a log it protects a measurement that may not be repeatable. Pass it when
    you know the incoming copy is the newer one for the SAME thing - after a rebuild, after a
    re-anchor appended its replacement lines, or after re-measuring the same guest.

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

# A manifest is recognised by its NAME, here and nowhere else, so the two places that care - which
# directory it lands in, and which refusal message it gets - cannot drift apart.
$ManifestPattern = 'corpus-*.jsonl'

# The guest name becomes a directory name. Hyper-V does not constrain VM names to path segments,
# and a name carrying a separator would either escape the destination or be silently rewritten -
# both of which file one guest's results somewhere the operator is not looking.
if ($VMName -ne $VMName.Trim() -or $VMName -match '[\\/:\*\?"<>\|]' -or $VMName -eq '.' -or $VMName -eq '..') {
    throw "REFUSED: -VMName '$VMName' is not usable as a directory name. It names the subdirectory this guest's results land in, and a name that is not a single plain path segment would file them somewhere other than where you would look for them. Rename the VM, or pass -Destination and pull one guest at a time."
}

if (-not $Destination) {
    $Destination = Join-Path $RepoRoot 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\vm-corpus'
}
New-Item -ItemType Directory -Force -Path $Destination | Out-Null
$GuestDestination = Join-Path $Destination $VMName

# Files a PREVIOUS version of this script left flat. Checked before the session is opened, so the
# notice appears even when the pull below goes on to refuse something. Manifests are excluded
# because the root is where they still belong.
$legacyFlat = @(Get-ChildItem -LiteralPath $Destination -File -ErrorAction SilentlyContinue |
        Where-Object {
            $n = $_.Name
            ($n -notlike $ManifestPattern) -and (($Include | Where-Object { $n -like $_ }).Count -gt 0)
        })
if ($legacyFlat.Count -gt 0) {
    Write-Warning "$($legacyFlat.Count) file(s) sit directly in $Destination that a pull would now write into a per-guest subdirectory:"
    foreach ($l in $legacyFlat) { Write-Warning "    $($l.Name)" }
    Write-Warning 'They are NOT touched: which guest produced them is not recorded anywhere, and that is the defect'
    Write-Warning 'this layout exists to fix - filing them under a guess would make a stale file into a wrong'
    Write-Warning 'attribution in the indexed-versus-unindexed comparison. Move them into the right guest'
    Write-Warning 'directory if you know which one it was, or delete them if you do not. This repeats every pull.'
    Write-Host ''
}

$cred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$session = New-PSSession -VMName $VMName -Credential $cred
try {
    $files = Invoke-Command -Session $session -ScriptBlock {
        param($root, $patterns)
        if (-not (Test-Path -LiteralPath $root)) { return @() }
        Get-ChildItem -LiteralPath $root -File |
            Where-Object { $n = $_.Name; ($patterns | Where-Object { $n -like $_ }).Count -gt 0 } |
            ForEach-Object {
                # Hashed on the guest, and now for EVERY file rather than only manifests:
                # comparing hashes is what lets the host refuse BEFORE the copy rather than after
                # it, and a measurement transcript whose machine has moved on is as unrepeatable
                # as a manifest is unrecoverable. A file that cannot be read - a log still held
                # open by a running build - yields a null hash rather than failing the pull; the
                # host then refuses only if something is actually at risk of being replaced.
                $sha = $null
                try { $sha = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256 -ErrorAction Stop).Hash }
                catch { $sha = $null }
                [pscustomobject]@{ FullName = $_.FullName; Name = $_.Name; Length = $_.Length; Sha256 = $sha }
            }
    } -ArgumentList $GuestPath, $Include

    if (-not $files -or $files.Count -eq 0) {
        Write-Warning "Nothing matched $($Include -join ', ') under $GuestPath on $VMName."
        Write-Warning 'If you expected a manifest here, look wider before concluding it is lost:'
        Write-Warning '  Get-ChildItem C:\ -Recurse -Filter corpus-*.jsonl -ErrorAction SilentlyContinue'
        return
    }

    # Created only when there is something to put in it, so a pull that finds nothing but a
    # manifest does not leave an empty directory implying results that were never taken.
    $pulledResults = @($files | Where-Object { $_.Name -notlike $ManifestPattern }).Count -gt 0
    if ($pulledResults) {
        New-Item -ItemType Directory -Force -Path $GuestDestination | Out-Null
    }

    foreach ($f in $files) {
        $isManifest = $f.Name -like $ManifestPattern
        if ($isManifest) { $targetDir = $Destination } else { $targetDir = $GuestDestination }
        $target = Join-Path $targetDir $f.Name

        if (-not $Force -and (Test-Path -LiteralPath $target)) {
            if (-not $f.Sha256) {
                throw @"
REFUSED: $target already exists here, and $($f.Name) could not be hashed on $VMName - so there is
no way to tell whether the incoming copy is the same file, a newer one, or another machine's.

A file open for writing is the usual cause: a build or a measurement still running on the guest.
Let it finish and pull again. If you already know the incoming copy is the one you want, -Force.
"@
            }

            $existing = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
            if ($existing -ne $f.Sha256 -and $isManifest) {
                throw @"
REFUSED: $target already exists here and holds DIFFERENT content from the copy on $VMName.

That file is a corpus manifest - the EntryID allowlist 'corpus-teardown' requires, and the file
'corpus-verify' reads to decide whether a corpus is still measurable. Teardown deletes only what a
manifest records, by EntryID AND ordinal subject tag, both required; there is no second route the
mailbox-safety rules permit. Replace the wrong one and the corpus it described becomes items in a
real store that nothing is entitled to remove, and the loss is silent - a manifest that is 2.9 MB
of plausible EntryIDs for the wrong machine looks exactly like the right one.

Stopping costs you one command. Overwriting costs a store nobody can clean up.

Manifests are the one part of the pull that still lands in the SHARED root, and this refusal is
why: two manifests with the same name means two corpora share a corpus id, and a per-guest
directory would have hidden that instead of surfacing it. They should not share one: give each
guest its own id (vm-indexed, vm-unindexed - Testbed/README.md section 3), rebuild with it, and
the names stop colliding. If this really is a newer copy of the SAME corpus - after a rebuild, or
a re-anchor - move the old one aside and re-run, or pass -Force.
"@
            }
            elseif ($existing -ne $f.Sha256) {
                throw @"
REFUSED: $target already exists here and holds DIFFERENT content from the copy on $VMName.

This directory is $VMName's alone, so the file underneath is an EARLIER PULL OF THIS SAME GUEST -
not another machine's, which is what the per-guest directory rules out. Almost always that means
you re-ran the measurement or the build and this is the newer transcript.

Nothing in a real store depends on this file the way it depends on a manifest. What it does hold
is a measurement of one machine at one moment, and once that machine has been rebuilt, re-indexed
or re-anchored the old numbers cannot be taken again. Comparing the indexed and the unindexed
guest is the entire reason there are two of them.

So: -Force if the incoming copy is the one you want, or move the old one aside first if you want
to keep both.
"@
            }
        }

        $shown = $f.Name
        if (-not $isManifest) { $shown = Join-Path $VMName $f.Name }
        Write-Host ("  {0,-44} {1,10:N0} bytes" -f $shown, $f.Length)
        Copy-Item -LiteralPath $f.FullName -Destination $target -FromSession $session -Force
    }
}
finally {
    Remove-PSSession $session
}

Write-Host ''
Write-Host "Manifests      -> $Destination"
# Only claimed when something actually landed there: the directory is not created otherwise, and a
# banner naming a path that does not exist is the sort of thing a reader later remembers as fact.
if ($pulledResults) { Write-Host "Results, logs  -> $GuestDestination" }
else { Write-Host "Results, logs  -> nothing matched this time; $GuestDestination was not created." }
Write-Host 'Both are gitignored. The manifest is the only thing that can tear the corpus down - keep a copy off this machine too.'
Write-Host 'The root stays shared for manifests ON PURPOSE: it is what makes a reused corpus id collide, and collide visibly.'
Write-Host 'Everything else is per guest, because what differs about a transcript is the machine that produced it, not what it is.'
