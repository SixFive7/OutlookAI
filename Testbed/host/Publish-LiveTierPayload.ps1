#Requires -Version 5.1
<#
    ============================================================================================
    RUN 2026-09-17, FIRST TIME, AND IT WORKED. 21 SECONDS, EXIT 0.
    ============================================================================================

    What it produced, on the maintainer's workstation against HEAD: 5 projects restored,
    **54 packages / 75.7 MB** - exactly the figure predicted by reading the existing
    project.assets.json files without running a restore - `Source.zip` 2.6 MB, `NuGet.zip`
    74.5 MB, and the feed check passing on all five projects with every other source cleared.

    That last step is the one that earns its keep: it proves the offline feed is self-sufficient
    HERE, in seconds, rather than after a ~283 MB copy into a guest. The payload was then copied
    into OutlookAI-Indexed and `guest/Install-DotnetSdk.ps1` reached VERDICT: TEST-READY on it -
    2,698 tests discovered, 17 executed, 0 failed, entirely offline.

    THE THING IT CANNOT STAGE, and it says so at the end of every run: the gitignored per-guest
    `live-test-settings.json`. It names real stores, so it is per-machine and never committed.
    Without it the live tier has no write allowlist, and that is a REFUSAL rather than a pass.

    Written by an agent that was not allowed to run it: it restores NuGet packages, which is a
    download, and downloads were out of scope for the session that produced it. Verified by
    PARSING only - the same check .github/scripts/check-testbed-references.ps1 applies to every
    script under Testbed/. No archive was produced, no package was fetched and no hash was
    computed to write it.

    Replace this banner with what it actually did once it has run, and record the two sizes it
    prints in Testbed/MEDIA.md - the entry there says they are estimates.

.SYNOPSIS
    Stages, on the HOST, the three things a guest needs before `dotnet test` can run: the SDK
    installer's identity, the source of the suite, and every NuGet package it uses, as an
    offline folder feed.

.DESCRIPTION
    THE GUEST HAS NO NETWORK, AND THAT IS ON PURPOSE. Testbed/MEDIA.md's whole model is that
    anything a guest needs is staged on the host beforehand and copied in - that is how Windows
    and Office get there, and Testbed/guest/Install-MailSink.ps1 says the same of the mail sink.
    A restore that reaches nuget.org from a guest is a restore that fails, so the packages are
    media like everything else.

    THREE ARTEFACTS, and the second and third are the ones people forget:

      Source.zip   -> C:\OutlookAI-Q5\src              the repository at a named commit. The
                                                       suite is BUILT on the guest; there is no
                                                       prebuilt test assembly anywhere.
      NuGet.zip    -> C:\OutlookAI-Q5\nuget-offline    every .nupkg the suite's restore graph
                                                       names, as a folder feed.
      the SDK      -> C:\OutlookAI-Q5\media\           NOT copied or produced here. This script
                                                       only PROVES which file you staged, by
                                                       printing its hash. See below.

    WHY THE SDK INSTALLER IS NOT AN OUTPUT OF THIS SCRIPT. Testbed/MEDIA.md draws a hard line
    between a PRECONDITION and an ARTEFACT: media is staged by a human, declared in that file,
    and never fetched by a script. The Windows ISO and the Office Deployment Tool both sit on
    that side of the line and so does this. What a script CAN usefully do is answer "is the file
    I staged the file the documentation means", which is a hash - so this one computes it and
    prints it in the form Testbed/guest/Install-DotnetSdk.ps1 wants it.

    WHY THE SOURCE COMES FROM `git archive` AND NOT THE WORKING TREE. A payload built from
    uncommitted edits is a payload nobody can reproduce, and every measurement taken with it is a
    statement about a tree that exists on one machine. `git archive` also excludes bin/ and obj/
    for free, which matters: a stale obj/ carrying the host's absolute package paths is exactly
    the kind of thing that makes a guest build fail in a way that reads as an SDK fault.

    AND WHY THE RESTORE RUNS AGAINST THE ARCHIVE RATHER THAN THE CHECKOUT. A `dotnet restore
    --packages <elsewhere>` rewrites every project's obj/project.assets.json to point at that
    elsewhere. Run against the maintainer's checkout, that silently breaks their next build until
    they restore again. Run against the expanded archive, it touches nothing anyone is using -
    and it proves the archived source restores, which is the thing the guest is about to do.

    THE FEED IS VERIFIED OFFLINE BEFORE IT SHIPS. Collecting nupkgs and hoping is how you find
    out on the guest, 250 MB of copy-in later, that one transitive package was resolved from a
    machine-wide fallback folder and never landed in the feed. So the last step restores the
    suite AGAIN, through a generated NuGet.config that clears every source and every fallback
    folder and names only the feed. If that succeeds, the feed is self-sufficient; if it fails,
    it names the missing package while the fix still costs a minute.

    WHAT IS DELIBERATELY NOT IN Source.zip, AND IT IS NOT AN OVERSIGHT.
    McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json is gitignored, so
    `git archive` cannot carry it and must not: it names real stores and this repository is
    public. Each guest needs its OWN, describing that guest's mailbox.
    Testbed/live-test-settings.example.json is the committed shape. A tier run without it is a
    tier run with no idea which store it may write to, which is the state the StoreWriteAllowlist
    exists to refuse.

    THIS SCRIPT TAKES NO -VMName, AND THAT IS CORRECT - the same rule
    Testbed/host/Publish-GuestPayload.ps1 states. It only builds on the host; naming a guest
    happens in the next command, Testbed/host/Copy-ToGuest.ps1, where -VMName is mandatory
    because several machines coexist and a default that silently picks one is the shape of
    mistake this testbed keeps making.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER OutDir
    Where the artefacts land. Default: .work/testbed-livetier-payload under the repository root,
    which is gitignored.

.PARAMETER SdkInstallerPath
    The staged .NET SDK installer to identify. Not copied, not produced, never downloaded.
    Default: the file Testbed/MEDIA.md declares.

.PARAMETER ExpectedSha512
    When given, the staged installer's SHA-512 must match or the script stops. Without it the
    hash is printed and nothing is asserted - which is the honest state before anyone has
    recorded one in Testbed/MEDIA.md.

.PARAMETER Ref
    The commit to archive. HEAD by default.

.PARAMETER SkipSource
    Leave Source.zip alone and reuse what a previous run staged.

.PARAMETER SkipPackages
    Leave NuGet.zip alone. Useful when only the source changed - but note the two can drift: a
    new PackageReference needs a new feed.

.PARAMETER SkipFeedVerify
    Skip the offline re-restore. Only do this if it is failing for a reason you have already
    understood; it is the check that stops a broken feed reaching a guest.

.EXAMPLE
    pwsh -File Testbed/host/Publish-LiveTierPayload.ps1
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName OutlookAI-Indexed -Path .work/testbed-livetier-payload/Source.zip -Destination C:\OutlookAI-Q5\Source.zip
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string] $OutDir,
    [string] $SdkInstallerPath,
    [string] $ExpectedSha512,
    [string] $Ref = 'HEAD',
    [switch] $SkipSource,
    [switch] $SkipPackages,
    [switch] $SkipFeedVerify
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

if (-not $OutDir) { $OutDir = Join-Path $RepoRoot '.work\testbed-livetier-payload' }
if (-not $SdkInstallerPath) { $SdkInstallerPath = Join-Path $RepoRoot '.work\media\dotnet-sdk-10.0.401-win-x64.exe' }

# Where each artefact goes on the guest. Repeated in the instructions this script prints, and
# kept here so there is one place to change it.
$GuestSourceRoot = 'C:\OutlookAI-Q5\src'
$GuestFeedRoot   = 'C:\OutlookAI-Q5\nuget-offline'
$GuestMediaRoot  = 'C:\OutlookAI-Q5\media'

$sourceZip = Join-Path $OutDir 'Source.zip'
$nugetZip  = Join-Path $OutDir 'NuGet.zip'
$sourceDir = Join-Path $OutDir 'source'
$feedDir   = Join-Path $OutDir 'nuget-offline'
$cacheDir  = Join-Path $OutDir 'package-cache'
$checkDir  = Join-Path $OutDir 'feed-check'

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# A restore or a build that leaves idle MSBuild workers behind is a nuisance on a host and a
# checkpoint hazard on a guest. Off everywhere, for one reason.
$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_NOLOGO = '1'

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
function SizeMb([string] $path) { return [math]::Round((Get-Item -LiteralPath $path).Length / 1MB, 1) }

# ---------------------------------------------------------------------------------------------
# 1. The SDK installer. Identified, never produced.
# ---------------------------------------------------------------------------------------------
Say '== The SDK installer =='
if (-not (Test-Path -LiteralPath $SdkInstallerPath)) {
    Say "  NOT STAGED: $SdkInstallerPath"
    Say '  This is a PRECONDITION, not an artefact - nothing here produces it and nothing here'
    Say '  downloads it. Get the win-x64 .NET SDK installer from Microsoft on this host, put it'
    Say "  there, and re-run. Testbed/MEDIA.md declares which version and where it comes from."
    Say '  Everything below still runs; only the hash is missing.'
}
else {
    $sdkHash = (Get-FileHash -LiteralPath $SdkInstallerPath -Algorithm SHA512).Hash
    Say "  file    $SdkInstallerPath"
    Say "  size    $(SizeMb $SdkInstallerPath) MB"
    Say "  sha512  $sdkHash"
    if ($ExpectedSha512) {
        $wanted = $ExpectedSha512.Trim().Replace('-', '').ToUpperInvariant()
        if ($sdkHash -ne $wanted) {
            throw @"
The staged installer is not the expected file.

  expected  $wanted
  actual    $sdkHash
  file      $SdkInstallerPath

A mismatch means "check which version you staged", not necessarily "the file is bad".
"@
        }
        Say '  MATCHES the expected hash.'
    }
    else {
        Say '  UNPINNED. Compare that value against the SHA-512 Microsoft publishes for this'
        Say '  installer, record it in Testbed/MEDIA.md, and pass it to the guest script:'
        Say "      .\Install-DotnetSdk.ps1 -ExpectedSha512 $sdkHash -Execute"
        Say '  A hash taken from the file you are about to trust proves only that the file has'
        Say '  not changed since you looked. The comparison against Microsoft is the check.'
    }
}

# ---------------------------------------------------------------------------------------------
# 2. Source.zip, from a commit rather than from the working tree.
# ---------------------------------------------------------------------------------------------
Say ''
Say '== Source =='
if (-not $SkipSource) {
    if (Test-Path -LiteralPath $sourceZip) { Remove-Item -LiteralPath $sourceZip -Force }
    $resolved = (@(& git -C $RepoRoot rev-parse --verify $Ref 2>&1) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { throw "git could not resolve '$Ref' in $RepoRoot - $resolved" }

    $dirty = @(& git -C $RepoRoot status --porcelain)
    if ($dirty.Count -gt 0) {
        Say "  NOTE the working tree has $($dirty.Count) uncommitted change(s). They are NOT in this"
        Say '       archive - it is built from a commit on purpose. Commit first if the guest is'
        Say '       meant to see them.'
    }

    Say "  archiving $Ref ($resolved)"
    & git -C $RepoRoot archive --format=zip -o $sourceZip $Ref
    if ($LASTEXITCODE -ne 0) { throw "git archive failed (exit $LASTEXITCODE)." }
    Say "  Source.zip  $(SizeMb $sourceZip) MB  -> expand on the guest into $GuestSourceRoot"
}
else {
    if (-not (Test-Path -LiteralPath $sourceZip)) { throw "-SkipSource, but $sourceZip does not exist." }
    Say "  reusing $sourceZip  ($(SizeMb $sourceZip) MB)"
}

# The restore below runs against the ARCHIVE, not the checkout - see the banner.
if (Test-Path -LiteralPath $sourceDir) { Remove-Item -LiteralPath $sourceDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $sourceDir | Out-Null
Expand-Archive -LiteralPath $sourceZip -DestinationPath $sourceDir -Force

$projects = @(Get-ChildItem -Path (Join-Path $sourceDir 'McpServer') -Filter '*.csproj' -Recurse -File |
        Sort-Object FullName)
if ($projects.Count -eq 0) {
    throw "No .csproj under $sourceDir\McpServer. Either the archive is wrong or the projects moved, and either way nothing below would prove anything."
}
Say "  $($projects.Count) project(s) in the archive:"
foreach ($p in $projects) { Say ("    " + $p.Name) }

# The add-in at the repository root is MSBuild-only (VSTO) and `dotnet restore` cannot read it.
# Restricting to McpServer/ is what CLAUDE.md says about building this repository, restated as
# code so it cannot be forgotten here.

# ---------------------------------------------------------------------------------------------
# 3. NuGet.zip - the offline folder feed.
# ---------------------------------------------------------------------------------------------
Say ''
Say '== Packages =='
if (-not $SkipPackages) {
    foreach ($d in @($cacheDir, $feedDir)) {
        if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force }
        New-Item -ItemType Directory -Force -Path $d | Out-Null
    }

    foreach ($p in $projects) {
        Say "  restoring $($p.Name)"
        $log = Join-Path $OutDir ("restore-" + [IO.Path]::GetFileNameWithoutExtension($p.Name) + ".log")
        # Output to a file, never streamed: a restore that spawns its own children can hold the
        # pipe open long after it has finished.
        & dotnet restore $p.FullName --packages $cacheDir *> $log
        if ($LASTEXITCODE -ne 0) {
            Write-Host (Get-Content -LiteralPath $log -Tail 40 | Out-String)
            throw "dotnet restore failed for $($p.Name) (exit $LASTEXITCODE). Full log: $log"
        }
    }

    # Flatten to .nupkg files. NuGet reads a folder source in either layout, and flat is both
    # smaller than the extracted cache and obvious to eyeball.
    $nupkgs = @(Get-ChildItem -Path $cacheDir -Filter '*.nupkg' -Recurse -File)
    if ($nupkgs.Count -eq 0) { throw "The restore produced no .nupkg under $cacheDir, so there is nothing to ship." }
    foreach ($n in $nupkgs) {
        Copy-Item -LiteralPath $n.FullName -Destination (Join-Path $feedDir $n.Name) -Force
    }
    $feedBytes = ($nupkgs | Measure-Object -Property Length -Sum).Sum
    Say "  $($nupkgs.Count) package(s), $([math]::Round($feedBytes / 1MB, 1)) MB"

    if (Test-Path -LiteralPath $nugetZip) { Remove-Item -LiteralPath $nugetZip -Force }
    Compress-Archive -Path (Join-Path $feedDir '*') -DestinationPath $nugetZip
    Say "  NuGet.zip  $(SizeMb $nugetZip) MB  -> expand on the guest into $GuestFeedRoot"
}
else {
    if (-not (Test-Path -LiteralPath $nugetZip)) { throw "-SkipPackages, but $nugetZip does not exist." }
    Say "  reusing $nugetZip  ($(SizeMb $nugetZip) MB)"
}

# ---------------------------------------------------------------------------------------------
# 4. Prove the feed is self-sufficient, here, where a fix is cheap.
# ---------------------------------------------------------------------------------------------
Say ''
Say '== Feed check =='
if ($SkipFeedVerify) {
    Say '  SKIPPED by -SkipFeedVerify. The guest is now the place this gets discovered.'
}
elseif (-not (Test-Path -LiteralPath $feedDir)) {
    Say "  SKIPPED: $feedDir does not exist (this run reused a zip it did not build)."
    Say '  Re-run without -SkipPackages to check the feed.'
}
else {
    if (Test-Path -LiteralPath $checkDir) { Remove-Item -LiteralPath $checkDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $checkDir | Out-Null

    # The same config shape Install-DotnetSdk.ps1 writes on the guest: every source cleared,
    # every fallback folder cleared, the feed named once. Passing it with --configfile is what
    # makes this an OFFLINE test on a machine that has a perfectly good network.
    $checkConfig = Join-Path $checkDir 'NuGet.config'
    Set-Content -LiteralPath $checkConfig -Encoding UTF8 -Value @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="testbed-offline" value="$feedDir" />
  </packageSources>
  <fallbackPackageFolders>
    <clear />
  </fallbackPackageFolders>
</configuration>
"@

    $checkPackages = Join-Path $checkDir 'packages'
    $ok = $true
    foreach ($p in $projects) {
        $log = Join-Path $OutDir ("feedcheck-" + [IO.Path]::GetFileNameWithoutExtension($p.Name) + ".log")
        & dotnet restore $p.FullName --configfile $checkConfig --packages $checkPackages *> $log
        if ($LASTEXITCODE -ne 0) {
            $ok = $false
            Write-Host (Get-Content -LiteralPath $log -Tail 30 | Out-String)
            Say "  FAIL $($p.Name) cannot restore from the feed alone. Log: $log"
        }
        else {
            Say "  OK   $($p.Name)"
        }
    }
    Remove-Item -LiteralPath $checkDir -Recurse -Force -ErrorAction SilentlyContinue
    if (-not $ok) {
        throw 'The feed is not self-sufficient. A guest would fail the same way, 300 MB of copy-in later. The log above names the package.'
    }
    Say '  The feed restores the whole suite with every other source cleared.'
}

# ---------------------------------------------------------------------------------------------
Say ''
Say "Payload in $OutDir"
Say ''
Say 'Copy in - one machine at a time, and say which (Testbed/README.md section 4a):'
Say "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $sourceZip -Destination C:\OutlookAI-Q5\Source.zip"
Say "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $nugetZip  -Destination C:\OutlookAI-Q5\NuGet.zip"
Say "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $SdkInstallerPath -Destination $GuestMediaRoot\$(Split-Path -Leaf $SdkInstallerPath)"
Say ''
Say 'Then on the guest, in an ELEVATED shell (PowerShell Direct is fine - no COM in any of this):'
Say "  Expand-Archive C:\OutlookAI-Q5\Source.zip -DestinationPath $GuestSourceRoot -Force"
Say "  Expand-Archive C:\OutlookAI-Q5\NuGet.zip  -DestinationPath $GuestFeedRoot   -Force"
Say '  .\Install-DotnetSdk.ps1 -ExpectedSha512 <the sha512 printed above> -Execute'
Say ''
Say 'STILL MISSING ON THE GUEST, and no script can supply it: the gitignored live-test settings at'
Say '  McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json'
Say 'It names real stores, so it is per-guest and never committed. Testbed/live-test-settings.example.json'
Say 'is the shape. Without it the live tier has no write allowlist, which is a refusal, not a pass.'
Say ''
Say 'And once the tier itself runs: everything that TOUCHES OUTLOOK goes through'
Say 'Testbed/guest/Register-InteractiveTask.ps1. Outlook cannot finish starting in session 0, so a'
Say '`dotnet test --filter Category=Live` driven straight over PowerShell Direct hangs.'
