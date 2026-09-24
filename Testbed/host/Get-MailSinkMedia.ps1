#Requires -Version 5.1
<#
    ============================================================================================
    RUN 2026-09-24 ON THE HOST, UNDER WINDOWS POWERSHELL 5.1, AND IT WORKED.
    ============================================================================================

    What ran, and what it showed: -SelfTest (every assertion passing under 5.1 and 7); the dry run
    against a copy staged by hand, which verified it; and -Execute into a scratch destination,
    which downloaded the checksums file and the zip, verified the zip three ways, and only then
    renamed it into place - "STAGED AND VERIFIED". The hashes are in Testbed/MEDIA.md, "The mail
    sink". It never touches a guest, so nothing about a guest is claimed here.

.SYNOPSIS
    Stages the testbed's mail sink - the pinned Inbucket release - on the HOST, and proves it is
    the file its maintainers published before anything copies it to a guest.

.DESCRIPTION
    RUN ON THE HOST. It downloads, so it runs where there is a network; the guests have none, and
    Testbed/guest/Install-MailSink.ps1 never downloads anything. Windows PowerShell 5.1 or 7.

    WHAT IT CHECKS, AND AGAINST WHAT. Three values have to agree before it says STAGED:

      1. the SHA-256 of the zip it holds;
      2. the line for that zip in the release's own checksums file - inbucket_<v>_checksums.txt,
         which the maintainers' release pipeline (goreleaser) attaches to every release. That is
         the "hash published by its maintainers", and it is fetched fresh, beside the zip;
      3. the hash this repository RECORDED when the release was chosen - in Testbed/MEDIA.md and
         in -ExpectedSha256's lookup below, which must agree with each other as well.

    (1) against (2) catches a corrupt or substituted download. (2) against (3) catches a release
    that changed after it was recorded - a re-tagged or re-uploaded asset - which is the case a
    maintainer-published checksum alone cannot catch, because it would change with the asset.
    An unsigned checksums file served from the same page as the zip proves provenance only as far
    as that page does; this file says so rather than implying more, exactly as Testbed/MEDIA.md
    says of Microsoft's .NET SDK hash.

    WHERE IT WRITES. Only into -Destination, by default the repository's gitignored .work\media -
    beside the Windows ISO and the .NET SDK, for the reason Testbed/MEDIA.md gives about
    volatile directories. It refuses a destination inside the repository that is not under .work,
    because an 11 MB binary in a tracked directory is one commit away from being in history for
    good. A download lands under a temporary name and is renamed only once it has verified, so a
    failed or interrupted run never leaves a file that looks staged.

    WHAT IT NEVER DOES. It never runs the sink, never unpacks it, never touches a guest and never
    touches this machine's mail. Copying into a guest is Testbed/host/Copy-ToGuest.ps1, whose
    lines it prints.

.PARAMETER Version
    The Inbucket release. Changing it needs -ExpectedSha256 too, unless this file records a hash
    for that version: a new version is a new decision, and Testbed/MEDIA.md is where it is made.

.PARAMETER ExpectedSha256
    The recorded hash for -Version. Defaults to the one this file records for it, which must also
    be the one Testbed/MEDIA.md records.

.PARAMETER Destination
    Where the zip and its checksums file are staged. Default: <repository>\.work\media.

.PARAMETER Execute
    Download whatever is missing, then verify. Without it nothing is downloaded or written: what
    is already staged is verified, and what would be fetched is printed.

.PARAMETER Force
    With -Execute, download again even when a staged copy exists.

.PARAMETER SelfTest
    Exercise the pure decisions - checksum-file parsing, the three-way verdict, the destination
    rule, the URLs - against synthetic inputs. No network, no file.

.EXAMPLE
    pwsh -File Testbed/host/Get-MailSinkMedia.ps1
    pwsh -File Testbed/host/Get-MailSinkMedia.ps1 -Execute
    powershell.exe -NoProfile -File Testbed\host\Get-MailSinkMedia.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $Version = '3.1.1',
    [string] $ExpectedSha256,
    [string] $Destination,
    [switch] $Execute,
    [switch] $Force,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# The recorded choice. One entry per release this repository has decided to use; the value must
# equal the one in Testbed/MEDIA.md, and the dry run fails when it does not.
$RecordedSha256 = @{
    '3.1.1' = '232fb49c92f88505be1feceb4be90b70ca59bb7853216dee7c8b2814c85235d0'
}

$Repository = 'inbucket/inbucket'

# Derived in the body, never in param(): Windows PowerShell 5.1 leaves $PSScriptRoot empty in
# parameter defaults when a script is started with -File.
$RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# =============================================================================================
# PURE DECISIONS - driven by -SelfTest.
# =============================================================================================

function Get-AssetName {
    param([string] $ForVersion)
    return "inbucket_${ForVersion}_windows_amd64.zip"
}

function Get-ChecksumsName {
    param([string] $ForVersion)
    return "inbucket_${ForVersion}_checksums.txt"
}

function Get-ReleaseAssetUrl {
    param([string] $ForVersion, [string] $Name)
    return "https://github.com/$Repository/releases/download/v$ForVersion/$Name"
}

function ConvertTo-NormalHash {
    param([string] $Hash)
    if ($null -eq $Hash) { return '' }
    return $Hash.Trim().Replace('-', '').ToLowerInvariant()
}

<#
    The published hash for one file, from a sha256sum-style list: "<hash>  <name>" per line, with
    the name optionally marked '*' for binary mode. Matched on the exact file name - never a
    prefix, never a pattern - and refused when the list names the file twice.
#>
function Get-PublishedHash {
    param([string[]] $Lines, [string] $Name)
    $found = @()
    foreach ($line in @($Lines)) {
        $m = [regex]::Match([string] $line, '^\s*([0-9A-Fa-f]{64})\s+\*?(\S.*?)\s*$')
        if ($m.Success -and ($m.Groups[2].Value -ceq $Name)) { $found += (ConvertTo-NormalHash $m.Groups[1].Value) }
    }
    if ($found.Count -gt 1) { throw "The checksums file names $Name $($found.Count) times. Refusing to pick one." }
    if ($found.Count -eq 0) { return $null }
    return $found[0]
}

# The three-way verdict, as a list of problems. Empty means staged and verified.
function Get-StagingProblems {
    param([string] $Actual, [string] $Published, [string] $Recorded)
    $problems = @()
    $a = ConvertTo-NormalHash $Actual
    $p = ConvertTo-NormalHash $Published
    $r = ConvertTo-NormalHash $Recorded
    if ($r -notmatch '^[0-9a-f]{64}$') { $problems += "the recorded hash '$Recorded' is not a SHA-256" }
    if (-not $p) { $problems += 'the maintainers'' checksums file does not name this zip' }
    elseif ($a -cne $p) { $problems += "the zip does not match the maintainers' published hash: file $a, published $p - a corrupt or substituted download" }
    if ($p -and $r -and ($p -cne $r)) { $problems += "the maintainers' published hash $p is not the one this repository recorded ($r) - the release changed after it was chosen, or the record is wrong" }
    return $problems
}

# A destination inside the repository must be under .work, so the binary cannot be committed.
function Get-DestinationProblem {
    param([string] $Root, [string] $Path)
    if (-not $Path) { return 'no destination' }
    $full = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\')
    $work = $rootFull + '\.work'
    $insideRepo = [string]::Equals($full, $rootFull, [System.StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith($rootFull + '\', [System.StringComparison]::OrdinalIgnoreCase)
    if (-not $insideRepo) { return $null }
    $underWork = [string]::Equals($full, $work, [System.StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith($work + '\', [System.StringComparison]::OrdinalIgnoreCase)
    if ($underWork) { return $null }
    return "$full is inside the repository but not under .work - a staged binary there is one commit away from being in history for good"
}

# =============================================================================================
# SELF-TEST.
# =============================================================================================

function Invoke-SelfTest {
    $script:StChecks = 0
    $script:StFailures = @()
    function Test-Case([string] $What, $Expected, $Actual) {
        $script:StChecks++
        $e = if ($null -eq $Expected) { '<null>' } else { (@($Expected) -join '|') }
        $a = if ($null -eq $Actual) { '<null>' } else { (@($Actual) -join '|') }
        if ($e -cne $a) { $script:StFailures += "$What : expected $e, got $a" }
    }

    $h = '232fb49c92f88505be1feceb4be90b70ca59bb7853216dee7c8b2814c85235d0'
    $other = 'f5b7942557a81e07c1d69a17c2f258d0676898d6ed177eaabbdda42ca8b6bd7b'
    $list = @(
        "0374ef7bd3c593a80e55aeb0503412083e6adfd0549efb607c6417b513405053  inbucket_3.1.1_darwin_amd64.tar.gz"
        "$h  inbucket_3.1.1_windows_amd64.zip"
        "d6d3018a522e71c3231766adb2668ae15a5708f29259755d2f0b7b805a6ed27d  inbucket_3.1.1_windows_arm64.zip"
    )

    Test-Case 'the asset name'                       'inbucket_3.1.1_windows_amd64.zip' (Get-AssetName '3.1.1')
    Test-Case 'the checksums name'                   'inbucket_3.1.1_checksums.txt'     (Get-ChecksumsName '3.1.1')
    Test-Case 'the download URL'                     'https://github.com/inbucket/inbucket/releases/download/v3.1.1/inbucket_3.1.1_windows_amd64.zip' (Get-ReleaseAssetUrl '3.1.1' 'inbucket_3.1.1_windows_amd64.zip')
    Test-Case 'the published hash is found'          $h       (Get-PublishedHash -Lines $list -Name 'inbucket_3.1.1_windows_amd64.zip')
    Test-Case 'a binary-mode * is understood'        $h       (Get-PublishedHash -Lines @("$h *inbucket_3.1.1_windows_amd64.zip") -Name 'inbucket_3.1.1_windows_amd64.zip')
    Test-Case 'upper-case hex is normalised'         $h       (Get-PublishedHash -Lines @(($h.ToUpperInvariant()) + '  x.zip') -Name 'x.zip')
    Test-Case 'the arm64 zip is not taken for amd64' '<null>' (Get-PublishedHash -Lines @($list[2]) -Name 'inbucket_3.1.1_windows_amd64.zip')
    Test-Case 'a prefix is not a match'              '<null>' (Get-PublishedHash -Lines @("$h  inbucket_3.1.1_windows_amd64.zip.sig") -Name 'inbucket_3.1.1_windows_amd64.zip')
    Test-Case 'a name is matched case-sensitively'   '<null>' (Get-PublishedHash -Lines @("$h  INBUCKET_3.1.1_windows_amd64.zip") -Name 'inbucket_3.1.1_windows_amd64.zip')
    $threw = $false
    try { [void](Get-PublishedHash -Lines @("$h  a.zip", "$other  a.zip") -Name 'a.zip') } catch { $threw = $true }
    Test-Case 'a file named twice is refused'        $true    $threw

    Test-Case 'three agreeing hashes are staged'     0        @(Get-StagingProblems -Actual $h -Published $h -Recorded $h.ToUpperInvariant()).Count
    Test-Case 'a corrupt download is caught'         $true    (((Get-StagingProblems -Actual $other -Published $h -Recorded $h) -join ' ').Contains('corrupt or substituted'))
    Test-Case 'a re-uploaded release is caught'      $true    (((Get-StagingProblems -Actual $other -Published $other -Recorded $h) -join ' ').Contains('changed after it was chosen'))
    Test-Case 'a missing published line is caught'   $true    (((Get-StagingProblems -Actual $h -Published $null -Recorded $h) -join ' ').Contains('does not name this zip'))
    Test-Case 'a malformed record is caught'         $true    (((Get-StagingProblems -Actual $h -Published $h -Recorded 'abc') -join ' ').Contains('not a SHA-256'))

    $root = 'C:\repo'
    Test-Case '.work\media is allowed'               '<null>' (Get-DestinationProblem -Root $root -Path 'C:\repo\.work\media')
    Test-Case '.work itself is allowed'              '<null>' (Get-DestinationProblem -Root $root -Path 'C:\repo\.work')
    Test-Case 'outside the repository is allowed'    '<null>' (Get-DestinationProblem -Root $root -Path 'D:\media')
    Test-Case 'a tracked directory is refused'       $true    ([bool] (Get-DestinationProblem -Root $root -Path 'C:\repo\Testbed\media'))
    Test-Case 'the repository root is refused'       $true    ([bool] (Get-DestinationProblem -Root $root -Path 'C:\repo'))
    Test-Case 'a look-alike sibling is not inside'   '<null>' (Get-DestinationProblem -Root $root -Path 'C:\repo-other\media')
    Test-Case '.. is resolved before judging'        $true    ([bool] (Get-DestinationProblem -Root $root -Path 'C:\repo\.work\..\Testbed'))

    Test-Case 'the recorded 3.1.1 hash is well formed' $true  ($RecordedSha256['3.1.1'] -match '^[0-9a-f]{64}$')

    Write-Host ("SELF-TEST: {0} assertion(s), {1} failure(s)." -f $script:StChecks, $script:StFailures.Count)
    foreach ($f in $script:StFailures) { Write-Host "  FAIL $f" }
    return $script:StFailures.Count
}

if ($SelfTest) {
    $failed = [int] (@(Invoke-SelfTest)[-1])
    if ($failed -gt 0) { exit 1 }
    exit 0
}

# =============================================================================================
# STAGE.
# =============================================================================================

$problems = @()
function Fail([string] $What) { $script:problems += $What; Write-Host "  FAIL $What" }

if (-not $Destination) { $Destination = Join-Path $RepoRoot '.work\media' }
$destinationProblem = Get-DestinationProblem -Root $RepoRoot -Path $Destination
if ($destinationProblem) { throw "REFUSING: $destinationProblem." }

if (-not $ExpectedSha256) {
    if (-not $RecordedSha256.ContainsKey($Version)) {
        throw "REFUSING: this file records no hash for Inbucket $Version. A new version is a new decision: check the maintainers' checksums file, record the hash in Testbed/MEDIA.md, and pass it with -ExpectedSha256."
    }
    $ExpectedSha256 = $RecordedSha256[$Version]
}

$assetName = Get-AssetName $Version
$checksumsName = Get-ChecksumsName $Version
$assetPath = Join-Path $Destination $assetName
$checksumsPath = Join-Path $Destination $checksumsName

Write-Host "Inbucket $Version for Windows x64 - the testbed's loopback mail sink (Testbed/MEDIA.md)."
Write-Host "  destination : $Destination"
Write-Host "  zip         : $(Get-ReleaseAssetUrl $Version $assetName)"
Write-Host "  checksums   : $(Get-ReleaseAssetUrl $Version $checksumsName)"
Write-Host "  recorded    : $(ConvertTo-NormalHash $ExpectedSha256)"
Write-Host ''

# The record has two homes; they must say the same thing.
$mediaDoc = Join-Path $RepoRoot 'Testbed\MEDIA.md'
if ($RecordedSha256.ContainsKey($Version) -and (Test-Path -LiteralPath $mediaDoc)) {
    $recorded = $RecordedSha256[$Version]
    if (-not ([System.IO.File]::ReadAllText($mediaDoc)).Contains($recorded)) {
        Fail "Testbed/MEDIA.md does not contain the hash this file records for $Version ($recorded). The record has two homes and they disagree; fix whichever is wrong."
    }
    else { Write-Host "  OK   Testbed/MEDIA.md records the same hash" }
}

$missing = @()
foreach ($p in @($assetPath, $checksumsPath)) { if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { $missing += $p } }

if ($Execute) {
    if (-not (Test-Path -LiteralPath $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
    # GitHub serves TLS 1.2 or later; Windows PowerShell 5.1 may not offer it by default.
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    foreach ($name in @($checksumsName, $assetName)) {
        $target = Join-Path $Destination $name
        if ((Test-Path -LiteralPath $target) -and -not $Force) { Write-Host "  present     $name"; continue }
        $partial = $target + '.partial'
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force }
        Write-Host "  downloading $name"
        Invoke-WebRequest -Uri (Get-ReleaseAssetUrl $Version $name) -OutFile $partial -UseBasicParsing
        if ($name -eq $checksumsName) { Move-Item -LiteralPath $partial -Destination $target -Force; continue }

        # The zip is only given its real name once it verifies.
        $published = Get-PublishedHash -Lines ([System.IO.File]::ReadAllLines((Join-Path $Destination $checksumsName))) -Name $assetName
        $actual = (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash
        $early = @(Get-StagingProblems -Actual $actual -Published $published -Recorded $ExpectedSha256)
        if ($early.Count -gt 0) {
            Remove-Item -LiteralPath $partial -Force
            foreach ($e in $early) { Fail $e }
            throw 'REFUSING: the download did not verify, and has been deleted rather than left looking staged.'
        }
        Move-Item -LiteralPath $partial -Destination $target -Force
    }
    $missing = @()
}
elseif ($missing.Count -gt 0) {
    Write-Host 'Dry run. Not staged yet:'
    foreach ($m in $missing) { Write-Host "  missing     $m" }
    Write-Host ''
    Write-Host 'Re-run with -Execute to download and verify them.'
    if ($problems.Count -gt 0) { exit 1 }
    exit 0
}

$published = Get-PublishedHash -Lines ([System.IO.File]::ReadAllLines($checksumsPath)) -Name $assetName
$actual = (Get-FileHash -LiteralPath $assetPath -Algorithm SHA256).Hash
$size = (Get-Item -LiteralPath $assetPath).Length
# Invariant, not -f: -f formats in the current culture, and this host renders 11444052 as
# "11.444.052" - a number that reads differently here than in Testbed/MEDIA.md.
Write-Host ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '  file        {0}  ({1:N0} bytes)', (ConvertTo-NormalHash $actual), $size))
Write-Host ("  published   {0}" -f $(if ($published) { $published } else { '<not named in the checksums file>' }))
Write-Host ("  recorded    {0}" -f (ConvertTo-NormalHash $ExpectedSha256))
foreach ($p in @(Get-StagingProblems -Actual $actual -Published $published -Recorded $ExpectedSha256)) { Fail $p }

Write-Host ''
if ($problems.Count -gt 0) {
    Write-Host "NOT STAGED: $($problems.Count) problem(s) above. Nothing here should be copied to a guest."
    exit 1
}

Write-Host 'STAGED AND VERIFIED: the file, the maintainers'' checksums file and this repository''s record agree.'
Write-Host ''
Write-Host 'Next, for each guest (Testbed/README.md section 1, step 7):'
Write-Host "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $assetPath -Destination C:\OutlookAI-Q5\media\$assetName"
Write-Host "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $(Join-Path $RepoRoot 'Testbed\guest\Install-MailSink.ps1') -Destination C:\OutlookAI-Q5\Install-MailSink.ps1"
Write-Host 'then on the guest, elevated (PowerShell Direct is fine: nothing here touches COM or Outlook):'
Write-Host "  C:\OutlookAI-Q5\Install-MailSink.ps1 -ExpectedSha256 $(ConvertTo-NormalHash $ExpectedSha256) -Execute"
exit 0
