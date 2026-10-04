#Requires -Version 5.1
<#
    ============================================================================================
    WRITTEN AND FIRST RUN 2026-10-04 (Q125), ON THE MAINTAINER'S WORKSTATION, UNDER POWERSHELL 7.
    ============================================================================================

    The pinned archive came down at 698,202,451 bytes with the pinned SHA-256, and Windows' tar
    unpacked it in 37 seconds; `codeql version` said 2.27.1. Over 04327b7 the database took 1.2
    minutes (432 of 432 C# files extracted) and the suite 1.2 minutes - 55 security queries and 8
    summary ones - with NO result. The verdict was then proven on a commit made for it and never put
    on a branch, which added a DES cipher in ECB mode: exit 1, one untriaged cs/ecb-encryption at
    the file and line; with that result's rule, path and line hash in the accepted list, exit 0 and
    the result shown as accepted, and an entry that matched nothing reported as stale. The other
    threat models were measured over 0ccf5df's sources (-ThreatModel has the numbers), each on a
    fresh database or with --rerun: a database analysed once hands its stored results back to a
    second `database analyze`, and the first try at "local" said no result for exactly that reason.

    What build-mode none costs, measured in the same run: the extractor restores each project with
    `dotnet restore` (the add-in's two old-style projects included - a restore runs no build target,
    and Tools/Switch-AddInBuild.ps1 -Status afterwards showed Outlook still on the installed release
    with its one trust entry), and it resolves the references it can: 10 stayed unresolved, the
    Visual Studio Tools for Office runtime (Microsoft.Office.Tools, 35 types) among them, so a flow
    that passes only through those types is one it cannot follow.

.SYNOPSIS
    Runs CodeQL's C# security queries over one commit of this repository, on this workstation, with
    the CodeQL bundle pinned below - downloaded once and held to the SHA-256 GitHub publishes for it.

.DESCRIPTION
    WHY THIS EXISTS. The GitHub CI that ran CodeQL on every push was removed on 2026-10-03 ("No CI
    pipelines on github"), and CodeQL's scan went with it. Decided by the maintainer the same day
    (Q125): the .NET SDK's security analysers run in every build (NetSecurityAnalyzers.targets,
    McpServer\Directory.Build.props), and CodeQL itself may still be run - locally, as an exception
    to the Dependencies rule. This is that run. Tools/Publish-Release.ps1 runs it before every
    release and refuses to release on any finding nobody has triaged.

    THE EXCEPTION IT RUNS UNDER (AGENTS.md, Dependencies, Q125), all of it enforced here or by
    where this lives:
      * the CodeQL CLI bundle and nothing else, at the version pinned below, held to the SHA-256
        GitHub publishes beside it (codeql-bundle-win64.tar.gz.checksum.txt on the release) BEFORE
        anything is unpacked;
      * on this workstation only, unpacked under the gitignored .work\ - never installed, never on
        PATH, never committed, never in the installer, never copied to anyone;
      * on this repository only, and only while it is an Open Source Codebase hosted on GitHub.com:
        step 1 refuses unless LICENSE is the MIT licence and origin is a github.com repository.

    THE LICENCE: the GitHub CodeQL Terms and Conditions, read 2026-10-04 from
    github/codeql-cli-binaries LICENSE.md. Free, per user, with no account and no key. For an "Open
    Source Codebase" - "released under an OSI-approved License", which MIT is - it permits
    "analysis on the Open Source Codebase" and, when that codebase "is hosted and maintained on
    GitHub.com", generating "CodeQL databases for or during automated analysis, CI, or CD" - which
    covers the release gate. It forbids redistributing the software ("share, publish, distribute or
    lend"), any use on a codebase that is not open source (a private repository included), and
    working around its technical limits. The queries themselves are github/codeql, MIT. Were this
    repository made private, the licence would no longer cover the run, and the release gate would
    have to go with it.

    WHERE IT RUNS, AND WHY THERE. Here, on the workstation: it has the internet the first run needs
    to fetch the bundle (700 MB) and build-mode none needs to restore the server's NuGet packages,
    builds are allowed here, and nothing in this run builds or registers the add-in. The build VM
    is offline and holds one run at a time for every agent's suite; the bundle would have to be
    staged there as media, and the run would hold the VM for as long as the analysis takes.

    WHAT ONE RUN DOES
      1. THE LICENCE CONDITION: LICENSE begins "MIT License" and origin is on github.com - or
         nothing runs.
      2. THE BUNDLE: <BundleRoot>\v<version>\codeql\codeql.exe, reused when its verified.json
         records the pinned hash and `codeql version` says the pinned version. Otherwise the
         archive is downloaded with Windows' own curl.exe, its length and SHA-256 compared with the
         pin, unpacked with Windows' own tar.exe, checked again, and deleted (the record stays).
         -FetchOnly stops here.
      3. THE SOURCE: `git archive` of -Ref, unpacked into the run folder - exactly the committed
         tree: no bin\ or obj\, no .work\, no other agent's worktree.
      4. THE DATABASE: `codeql database create --language=csharp --build-mode=none`. Nothing is
         built - so no MSBuild runs the add-in's project, and the Q81 rules are not in play. The
         extractor restores the server's NuGet packages (from nuget.org) to resolve its references.
      5. THE QUERIES: the C# security-extended suite - every security query CodeQL ships for C#,
         the default code-scanning set plus the lower-precision ones - into results.sarif, under
         CodeQL's default threat model (untrusted data comes from the network), as GitHub's code
         scanning ran it. -ThreatModel adds others; the release adds none - see the parameter.
      6. THE VERDICT: every result, by rule, file and line, against Tools\codeql-accepted.json -
         the findings already triaged as false positives, each with its reason. Any result not on
         that list fails the run. An entry that no longer matches anything is reported, so the
         list does not outlive its findings.

    WHAT IT LEAVES: <OutDir>\summary.txt, summary.json, results.sarif and logs\. The source copy and
    the database are deleted at the end unless -KeepDatabase.

    EXIT CODES: 0 no untriaged finding; 1 at least one; 2 the database or the analysis failed;
    3 not run - the bundle could not be fetched, verified or unpacked; 4 refused.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`, ASCII only.

.PARAMETER Ref
    The commit to analyse. HEAD by default. Uncommitted edits are never in the run.

.PARAMETER RepoRoot
    Repository root. Defaults to the folder above this script's.

.PARAMETER OutDir
    Where the run lands. Default .work\codeql\runs\<time>-<commit> under the repository root.
    Refused inside the git working tree anywhere but under .work\.

.PARAMETER BundleRoot
    Where the bundle is unpacked and kept. Default .work\codeql\bundle under the repository root. An
    agent in a worktree can point it at another checkout's copy instead of fetching 700 MB again.

.PARAMETER FetchOnly
    Fetch, verify and unpack the bundle (step 2), and stop.

.PARAMETER KeepDatabase
    Keep the source copy and the CodeQL database in the run folder, for querying afterwards.

.PARAMETER ThreatModel
    Threat models to add to CodeQL's default, each passed as --threat-model: "local" (standard
    input, files, the command line, the environment and the registry count as untrusted), or a part
    of it, and "!name" to take one away again. None by default, and none in a release. Measured over
    0ccf5df (2026-10-04), which has no result under the default: "local" gives 174 - 169
    cs/path-injection, 4 cs/command-line-injection, 1 cs/sql-injection - nearly all from the user's
    own profile folders and TEMP; "local", "!environment" gives 6, all reading HKCU or HKLM or a
    tool's command line, one of them the elevated helper's path (TODO.md); "stdin" alone gives none.
    Whether the release should run one of those is an open question in TODO.md.

.PARAMETER TimeoutMinutes
    The longest the download, the database or the analysis may each take. Default 60.

.PARAMETER SelfTest
    Run the pure decisions against synthetic inputs and exit. No network, no git, no CodeQL - it
    runs on the build VM like every other self-test.

.EXAMPLE
    pwsh -File Tools/Invoke-CodeQL.ps1
    pwsh -File Tools/Invoke-CodeQL.ps1 -Ref 1a2b3c4 -KeepDatabase
    pwsh -File Tools/Invoke-CodeQL.ps1 -FetchOnly
    pwsh -File Tools/Invoke-CodeQL.ps1 -ThreatModel local,!environment
    pwsh -File Tools/Invoke-CodeQL.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $Ref = 'HEAD',
    [string] $RepoRoot,
    [string] $OutDir,
    [string] $BundleRoot,
    [switch] $FetchOnly,
    [switch] $KeepDatabase,
    [string[]] $ThreatModel = @(),
    [int]    $TimeoutMinutes = 60,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}
# The running edition's OWN Security and Utility modules, by path - Testbed/host/OwnEditionModules.ps1
# says why (Windows PowerShell 5.1 started from PowerShell 7 resolves both to 7's copies and loses
# Get-FileHash). Restated inline, as Tools/Publish-Release.ps1 does, because this lives outside Testbed/.
foreach ($ownEditionModuleName in @('Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility')) {
    Import-Module (Join-Path $PSHOME "Modules\$ownEditionModuleName\$ownEditionModuleName.psd1")
}

# Defaults that need $PSScriptRoot are set HERE, not in param(): Windows PowerShell 5.1 leaves it
# empty while param() defaults are evaluated (Q78).
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
if (-not $BundleRoot) { $BundleRoot = Join-Path $RepoRoot '.work\codeql\bundle' }

# ---------------------------------------------------------------------------------------------
# THE PIN. Taken 2026-10-04 from the release codeql-bundle-v2.27.1 of github/codeql-action - the
# CodeQL CLI 2.27.1 with its query packs, the latest bundle that day. The hash is the one its
# maintainers publish beside the archive, in codeql-bundle-win64.tar.gz.checksum.txt; GitHub's own
# asset digest for the archive said the same. A new version is a deliberate edit of all four lines,
# from that file of the new release - never a hash computed from whatever was downloaded.
# ---------------------------------------------------------------------------------------------
$CodeQLVersion = '2.27.1'
$BundleUrl = 'https://github.com/github/codeql-action/releases/download/codeql-bundle-v2.27.1/codeql-bundle-win64.tar.gz'
$BundleSha256 = '721209b52b74d0cf742c701eeb452a70f570f236c75dda6087bd4ad2a6f97202'
$BundleBytes = [long]698202451

# The queries: every security query CodeQL ships for C#. The bundle carries the pack; nothing is
# downloaded for it.
$QuerySuite = 'codeql/csharp-queries:codeql-suites/csharp-security-extended.qls'


# The triaged false positives, relative to the repository root.
$AcceptedFile = 'Tools\codeql-accepted.json'

# Windows' own tools, by path: nothing on PATH gets to stand in for them.
$CurlExe = Join-Path $env:SystemRoot 'System32\curl.exe'
$TarExe = Join-Path $env:SystemRoot 'System32\tar.exe'

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# Every short native call goes through this: under Windows PowerShell 5.1 with 'Stop', the first
# line a program writes to stderr ends the script when anything redirects it (Tools/Checks/
# check-powershell-51.ps1, check 3). 'Continue' holds only in here; the try keeps "program not
# found" terminating; callers judge the call by $LASTEXITCODE. Restated, as in the other scripts.
function Invoke-NativeCommand {
    param([Parameter(Mandatory = $true)] [scriptblock] $NativeCommand)

    $ErrorActionPreference = 'Continue'
    try {
        & $NativeCommand | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { $_ }
        }
    }
    catch {
        throw
    }
}

# A long-running program with its output going STRAIGHT TO FILES, never through this process's
# pipe, and a deadline: a program that leaves children behind can hold a pipe open after it has
# finished, and a declared timeout then never fires. Restated from Tools/Publish-Release.ps1.
function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)] [string] $FilePath,
        [Parameter(Mandatory = $true)] [string[]] $ArgumentList,
        [Parameter(Mandatory = $true)] [string] $LogStem,
        [int] $TimeoutMinutes = 20,
        [string] $WorkingDirectory
    )
    $out = "$LogStem.out.txt"
    $err = "$LogStem.err.txt"
    foreach ($f in @($out, $err)) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } }
    $startArgs = @{
        FilePath               = $FilePath
        ArgumentList           = ($ArgumentList -join ' ')
        RedirectStandardOutput = $out
        RedirectStandardError  = $err
        NoNewWindow            = $true
        PassThru               = $true
    }
    if ($WorkingDirectory) { $startArgs['WorkingDirectory'] = $WorkingDirectory }
    $p = Start-Process @startArgs
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = -1; TimedOut = $true; Out = $out; Err = $err }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false; Out = $out; Err = $err }
}

function Show-LogTail([string] $Path, [int] $Lines = 20) {
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        foreach ($l in (Get-Content -LiteralPath $Path -Tail $Lines)) { Write-Host "      | $l" }
    }
}

# One argument for Start-Process's joined command line.
function Format-Argument([string] $Value) {
    if ($Value -match '\s') { return '"' + $Value + '"' }
    return $Value
}

# =============================================================================================
# PURE DECISIONS - everything -SelfTest checks
# =============================================================================================

# Step 1: the licence's condition, from the two things that state it. $null when it holds.
function Get-LicenceProblem {
    param([string] $LicenseText, [string] $OriginUrl)
    if ([string]::IsNullOrWhiteSpace($LicenseText)) {
        return 'there is no LICENSE file, so this is not an Open Source Codebase in the CodeQL licence''s terms, and the licence does not cover the run'
    }
    if (-not $LicenseText.TrimStart().StartsWith('MIT License', [System.StringComparison]::Ordinal)) {
        return 'LICENSE is no longer the MIT licence. The CodeQL licence covers only a codebase under an OSI-approved licence: check the new one is, then change this check'
    }
    if ($OriginUrl -notmatch '^(https://github\.com/|git@github\.com:|ssh://git@github\.com/)') {
        return "origin is '$OriginUrl', not a github.com repository. The CodeQL licence lets automated analysis run only for an open-source codebase hosted and maintained on GitHub.com"
    }
    return $null
}

# Step 2: the downloaded archive against the pin. $null when it is the pinned file.
function Get-ArchiveProblem {
    param([long] $Bytes, [string] $Sha256, [long] $ExpectedBytes, [string] $ExpectedSha256)
    if ($Bytes -ne $ExpectedBytes) { return "the archive is $Bytes bytes, not the pinned $ExpectedBytes" }
    if (-not [string]::Equals($Sha256, $ExpectedSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        return "the archive's SHA-256 is $Sha256, not the pinned $ExpectedSha256 that GitHub publishes for it"
    }
    return $null
}

# Step 2: whether an unpacked bundle can be reused - its record names the pin, and the CLI in it
# reports the pinned version. 'reuse', or why not.
function Get-BundleDecision {
    param([string] $RecordJson, [string] $ReportedVersion, [string] $ExpectedSha256, [string] $ExpectedVersion)
    if ([string]::IsNullOrWhiteSpace($RecordJson)) { return 'no record of a verified bundle' }
    try { $record = $RecordJson | ConvertFrom-Json } catch { return 'its record does not parse' }
    if (-not [string]::Equals([string]$record.sha256, $ExpectedSha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        return "its record names $($record.sha256), not the pinned archive"
    }
    if ([string]$ReportedVersion -cne $ExpectedVersion) { return "codeql version reports '$ReportedVersion', not $ExpectedVersion" }
    return 'reuse'
}

# Step 3: an output folder inside the git working tree is refused anywhere but under .work\ - the
# rule every build script here keeps, so a run can never leave files git would see.
function Test-IsUnderWorkTreeButNotWork {
    param([string] $Path, [string] $WorkTreeRoot)
    if (-not $WorkTreeRoot) { return $false }
    $p = [System.IO.Path]::GetFullPath($Path).TrimEnd('\') + '\'
    $root = [System.IO.Path]::GetFullPath($WorkTreeRoot).TrimEnd('\') + '\'
    if (-not $p.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    return -not $p.StartsWith($root + '.work\', [System.StringComparison]::OrdinalIgnoreCase)
}

# Step 6: one object per result in a CodeQL SARIF file. Path is repository-relative with forward
# slashes; LineHash is CodeQL's primaryLocationLineHash, which follows the line's text rather than
# its number, so an unrelated edit above a finding does not orphan its triage.
function Get-SarifFindings {
    param([string] $SarifText)
    $sarif = $SarifText | ConvertFrom-Json
    $findings = @()
    foreach ($run in @($sarif.runs)) {
        $levels = @{}
        $components = @()
        if ($run.tool.driver) { $components += $run.tool.driver }
        if ($run.tool.extensions) { $components += @($run.tool.extensions) }
        foreach ($c in $components) {
            foreach ($rule in @($c.rules)) {
                if ($null -eq $rule -or -not $rule.id) { continue }
                $sev = ''
                if ($rule.properties -and $rule.properties.'security-severity') { $sev = [string]$rule.properties.'security-severity' }
                $lvl = ''
                if ($rule.defaultConfiguration -and $rule.defaultConfiguration.level) { $lvl = [string]$rule.defaultConfiguration.level }
                $levels[[string]$rule.id] = [pscustomobject]@{ Level = $lvl; SecuritySeverity = $sev }
            }
        }
        foreach ($r in @($run.results)) {
            if ($null -eq $r) { continue }
            $ruleId = [string]$r.ruleId
            if (-not $ruleId -and $r.rule) { $ruleId = [string]$r.rule.id }
            $path = ''
            $line = 0
            $loc = @($r.locations) | Select-Object -First 1
            if ($loc -and $loc.physicalLocation) {
                $path = ([string]$loc.physicalLocation.artifactLocation.uri).Replace('\', '/')
                if ($loc.physicalLocation.region -and $loc.physicalLocation.region.startLine) { $line = [int]$loc.physicalLocation.region.startLine }
            }
            $hash = ''
            if ($r.partialFingerprints -and $r.partialFingerprints.primaryLocationLineHash) { $hash = [string]$r.partialFingerprints.primaryLocationLineHash }
            $level = [string]$r.level
            $sev = ''
            if ($levels.ContainsKey($ruleId)) {
                if (-not $level) { $level = $levels[$ruleId].Level }
                $sev = $levels[$ruleId].SecuritySeverity
            }
            if (-not $level) { $level = 'warning' }
            $findings += [pscustomobject]@{
                Rule = $ruleId; Level = $level; SecuritySeverity = $sev; Path = $path; Line = $line
                LineHash = $hash; Message = ([string]$r.message.text -replace '\s+', ' ').Trim()
            }
        }
    }
    return $findings
}

# Step 6: the accepted list - every entry names a rule, a path, a line hash and a reason, or the
# list is refused whole: a triage without its reason is exactly what the list exists to prevent.
function Read-AcceptedFindings {
    param([string] $Json)
    $result = [pscustomobject]@{ Entries = @(); Problem = $null }
    if ([string]::IsNullOrWhiteSpace($Json)) { $result.Problem = 'the accepted-findings file is empty'; return $result }
    try { $doc = $Json | ConvertFrom-Json } catch { $result.Problem = "the accepted-findings file does not parse: $($_.Exception.Message)"; return $result }
    $entries = @()
    $n = 0
    foreach ($e in @($doc.accepted)) {
        if ($null -eq $e) { continue }
        $n++
        foreach ($field in @('rule', 'path', 'lineHash', 'reason')) {
            if ([string]::IsNullOrWhiteSpace([string]$e.$field)) {
                $result.Problem = "accepted entry $n has no '$field' - every triaged finding names its rule, its file, its line hash and why it is safe"
                return $result
            }
        }
        $entries += [pscustomobject]@{ Rule = [string]$e.rule; Path = ([string]$e.path).Replace('\', '/'); LineHash = [string]$e.lineHash; Reason = [string]$e.reason }
    }
    $result.Entries = $entries
    return $result
}

# Step 6: the verdict. A finding is accepted when an entry names its rule, path and line hash
# exactly; anything else is untriaged. An entry that matches nothing is stale.
function Get-FindingsVerdict {
    param([object[]] $Findings, [object[]] $Accepted)
    $untriaged = @()
    $acceptedHits = @()
    $used = @{}
    foreach ($f in @($Findings | Where-Object { $null -ne $_ })) {
        $match = $null
        for ($i = 0; $i -lt @($Accepted).Count; $i++) {
            $a = @($Accepted)[$i]
            if ($null -eq $a) { continue }
            if ($a.Rule -ceq $f.Rule -and [string]::Equals($a.Path, $f.Path, [System.StringComparison]::OrdinalIgnoreCase) -and $a.LineHash -ceq $f.LineHash) { $match = $i; break }
        }
        if ($null -eq $match) { $untriaged += $f } else { $acceptedHits += $f; $used[$match] = $true }
    }
    $stale = @()
    for ($i = 0; $i -lt @($Accepted).Count; $i++) {
        if (-not $used.ContainsKey($i) -and $null -ne @($Accepted)[$i]) { $stale += @($Accepted)[$i] }
    }
    return [pscustomobject]@{ Untriaged = $untriaged; Accepted = $acceptedHits; Stale = $stale }
}

# Step 5: -ThreatModel's values, one each - every value split on commas, because `-File` hands
# "local,!environment" over as ONE string - trimmed, blanks dropped, order kept.
function Split-ThreatModel {
    param([string[]] $Models)
    $out = @()
    foreach ($m in @($Models)) {
        foreach ($part in ([string]$m).Split(',')) {
            if (-not [string]::IsNullOrWhiteSpace($part)) { $out += $part.Trim() }
        }
    }
    return $out
}

# As CodeQL's arguments: one --threat-model per value, in order. Nothing at all for none - CodeQL's
# default threat model, the one the release runs.
function Get-ThreatModelArguments {
    param([string[]] $Models)
    $out = @()
    foreach ($m in @(Split-ThreatModel $Models)) {
        $out += '--threat-model'
        $out += $m
    }
    return $out
}

# A value Start-Process's joined command line would split or misquote is refused, not passed on.
function Get-ThreatModelProblem {
    param([string[]] $Models)
    foreach ($m in @(Split-ThreatModel $Models)) {
        if ($m -notmatch '^!?[A-Za-z0-9_-]+$') { return "'$m' is not a threat model name (letters, digits, '-', '_', an optional leading '!')" }
    }
    return $null
}

function Format-ThreatModel {
    param([string[]] $Models)
    $named = @(Split-ThreatModel $Models)
    if ($named.Count -eq 0) { return 'default' }
    return 'default + ' + ($named -join ', ')
}

# One line per finding, for the summary and the console.
function Format-Finding {
    param($Finding)
    $sev = ''
    if ($Finding.SecuritySeverity) { $sev = " security-severity $($Finding.SecuritySeverity)" }
    return "$($Finding.Rule) [$($Finding.Level)$sev] $($Finding.Path):$($Finding.Line) - $($Finding.Message)"
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()

    # Compared with -ceq. Never -like: a bracket in a -like pattern is a character class.
    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:Checks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = ($Expected -join ' | ') }
        if ($Actual -is [System.Array]) { $a = ($Actual -join ' | ') }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else {
            $script:Failures += "$What : expected [$e], got [$a]"
            Write-Host "  FAIL $What - expected [$e], got [$a]"
        }
    }

    Write-Host "Tools/Invoke-CodeQL.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)). No network, no git, no CodeQL."
    Write-Host ''
    Write-Host '== the pin =='
    Test-Case 'the version is major.minor.patch' $true ($CodeQLVersion -match '^\d+\.\d+\.\d+$')
    Test-Case 'the URL is the bundle release of that version' $true ($BundleUrl -ceq "https://github.com/github/codeql-action/releases/download/codeql-bundle-v$CodeQLVersion/codeql-bundle-win64.tar.gz")
    Test-Case 'the hash is 64 lower-case hex characters' $true ($BundleSha256 -cmatch '^[0-9a-f]{64}$')
    Test-Case 'the length is pinned too' $true ($BundleBytes -gt 100MB)
    Test-Case 'the queries are the C# security suite, from the bundle''s own pack' 'codeql/csharp-queries:codeql-suites/csharp-security-extended.qls' $QuerySuite
    Test-Case 'no -ThreatModel is CodeQL''s default and adds no argument' '' ((Get-ThreatModelArguments @()) -join ' ')
    Test-Case 'each -ThreatModel value becomes one --threat-model, in order' '--threat-model local --threat-model !environment' ((Get-ThreatModelArguments @('local', '!environment')) -join ' ')
    Test-Case 'blank values are dropped' '--threat-model stdin' ((Get-ThreatModelArguments @('', ' stdin ')) -join ' ')
    Test-Case 'a comma list - what -File hands over - is split' '--threat-model local --threat-model !environment' ((Get-ThreatModelArguments @('local,!environment')) -join ' ')
    Test-Case 'a value that could split the command line is refused' $true ([string](Get-ThreatModelProblem @('local env'))).Contains('local env')
    Test-Case 'the summary names the default' 'default' (Format-ThreatModel @())
    Test-Case 'and what was added to it' 'default + local, !environment' (Format-ThreatModel @('local', '!environment'))
    Test-Case 'curl.exe and tar.exe are Windows'' own, by path' $true ($CurlExe.EndsWith('\System32\curl.exe') -and $TarExe.EndsWith('\System32\tar.exe'))

    Write-Host ''
    Write-Host '== step 1: the licence condition =='
    $mit = "MIT License`n`nCopyright (c) 2026 Jori Huisman`n"
    Test-Case 'MIT on github.com over https holds' $true ($null -eq (Get-LicenceProblem $mit 'https://github.com/SixFive7/OutlookAI.git'))
    Test-Case 'MIT on github.com over ssh holds' $true ($null -eq (Get-LicenceProblem $mit 'git@github.com:SixFive7/OutlookAI.git'))
    Test-Case 'no LICENSE is refused' $true ([string](Get-LicenceProblem '' 'https://github.com/SixFive7/OutlookAI.git')).Contains('no LICENSE')
    Test-Case 'another licence is refused until someone checks it' $true ([string](Get-LicenceProblem 'All rights reserved.' 'https://github.com/SixFive7/OutlookAI.git')).Contains('no longer the MIT licence')
    Test-Case 'a remote off github.com is refused' $true ([string](Get-LicenceProblem $mit 'https://gitlab.com/x/y.git')).Contains('not a github.com repository')
    Test-Case 'a look-alike host is refused' $true ([string](Get-LicenceProblem $mit 'https://github.com.evil.example/x/y.git')).Contains('not a github.com repository')

    Write-Host ''
    Write-Host '== step 2: the archive and the unpacked bundle =='
    Test-Case 'the pinned file passes' $true ($null -eq (Get-ArchiveProblem -Bytes $BundleBytes -Sha256 $BundleSha256.ToUpperInvariant() -ExpectedBytes $BundleBytes -ExpectedSha256 $BundleSha256))
    Test-Case 'a short download is refused before hashing matters' $true ([string](Get-ArchiveProblem -Bytes 5 -Sha256 $BundleSha256 -ExpectedBytes $BundleBytes -ExpectedSha256 $BundleSha256)).Contains('not the pinned')
    Test-Case 'a different hash is refused' $true ([string](Get-ArchiveProblem -Bytes $BundleBytes -Sha256 ('0' * 64) -ExpectedBytes $BundleBytes -ExpectedSha256 $BundleSha256)).Contains('GitHub publishes')
    $record = '{"version":"' + $CodeQLVersion + '","sha256":"' + $BundleSha256 + '"}'
    Test-Case 'a recorded, matching bundle is reused' 'reuse' (Get-BundleDecision -RecordJson $record -ReportedVersion $CodeQLVersion -ExpectedSha256 $BundleSha256 -ExpectedVersion $CodeQLVersion)
    Test-Case 'no record is not reused' 'no record of a verified bundle' (Get-BundleDecision -RecordJson '' -ReportedVersion $CodeQLVersion -ExpectedSha256 $BundleSha256 -ExpectedVersion $CodeQLVersion)
    Test-Case 'a record of another archive is not reused' $true ([string](Get-BundleDecision -RecordJson '{"sha256":"abc"}' -ReportedVersion $CodeQLVersion -ExpectedSha256 $BundleSha256 -ExpectedVersion $CodeQLVersion)).StartsWith('its record names abc')
    Test-Case 'a CLI reporting another version is not reused' $true ([string](Get-BundleDecision -RecordJson $record -ReportedVersion '2.0.0' -ExpectedSha256 $BundleSha256 -ExpectedVersion $CodeQLVersion)).Contains("not $CodeQLVersion")
    Test-Case 'a record that does not parse is not reused' 'its record does not parse' (Get-BundleDecision -RecordJson '{nope' -ReportedVersion $CodeQLVersion -ExpectedSha256 $BundleSha256 -ExpectedVersion $CodeQLVersion)

    Write-Host ''
    Write-Host '== step 3: where a run may land =='
    Test-Case 'under .work is accepted' $false (Test-IsUnderWorkTreeButNotWork 'C:\r\.work\codeql\runs\x' 'C:\r')
    Test-Case 'elsewhere in the working tree is refused' $true (Test-IsUnderWorkTreeButNotWork 'C:\r\Tools\out' 'C:\r')
    Test-Case 'a folder merely named like .work is refused' $true (Test-IsUnderWorkTreeButNotWork 'C:\r\.workx\out' 'C:\r')
    Test-Case 'outside the working tree is accepted' $false (Test-IsUnderWorkTreeButNotWork 'D:\scratch\codeql' 'C:\r')

    Write-Host ''
    Write-Host '== step 6: findings, triage and the verdict =='
    $sarif = @'
{"version":"2.1.0","runs":[{"tool":{"driver":{"name":"CodeQL","rules":[]},"extensions":[{"name":"codeql/csharp-queries","rules":[
 {"id":"cs/sql-injection","defaultConfiguration":{"level":"error"},"properties":{"security-severity":"8.8"}},
 {"id":"cs/path-injection","defaultConfiguration":{"level":"error"},"properties":{"security-severity":"7.5"}}]}]},
 "results":[
 {"ruleId":"cs/sql-injection","message":{"text":"This query depends on\n a user-provided value."},"locations":[{"physicalLocation":{"artifactLocation":{"uri":"McpServer/A.cs","uriBaseId":"%SRCROOT%"},"region":{"startLine":12}}}],"partialFingerprints":{"primaryLocationLineHash":"aaaa:1"}},
 {"ruleId":"cs/path-injection","level":"warning","message":{"text":"Path from input."},"locations":[{"physicalLocation":{"artifactLocation":{"uri":"Services/B.cs"},"region":{"startLine":40}}}],"partialFingerprints":{"primaryLocationLineHash":"bbbb:1"}}
 ]}]}
'@
    $findings = Get-SarifFindings $sarif
    Test-Case 'two results are two findings' 2 @($findings).Count
    Test-Case 'the rule is read' 'cs/sql-injection' $findings[0].Rule
    Test-Case 'the path is repository-relative with forward slashes' 'McpServer/A.cs' $findings[0].Path
    Test-Case 'the line is read' 12 $findings[0].Line
    Test-Case 'the line hash is read' 'aaaa:1' $findings[0].LineHash
    Test-Case 'a level the result omits comes from its rule' 'error' $findings[0].Level
    Test-Case 'a level the result states wins' 'warning' $findings[1].Level
    Test-Case 'the security severity comes from the rule' '8.8' $findings[0].SecuritySeverity
    Test-Case 'the message is one line' 'This query depends on a user-provided value.' $findings[0].Message
    Test-Case 'a SARIF with no results is no findings' 0 @(Get-SarifFindings '{"runs":[{"tool":{"driver":{"name":"CodeQL"}},"results":[]}]}').Count

    $acc = Read-AcceptedFindings '{"accepted":[{"rule":"cs/path-injection","path":"Services\\B.cs","lineHash":"bbbb:1","reason":"the path is a constant"},{"rule":"cs/x","path":"Gone.cs","lineHash":"cccc:1","reason":"fixed since"}]}'
    Test-Case 'a complete accepted list is read' $true ($null -eq $acc.Problem -and @($acc.Entries).Count -eq 2)
    Test-Case 'an entry''s path is normalised to forward slashes' 'Services/B.cs' $acc.Entries[0].Path
    Test-Case 'an entry without a reason refuses the whole list' $true ([string](Read-AcceptedFindings '{"accepted":[{"rule":"r","path":"p","lineHash":"h"}]}').Problem).Contains("no 'reason'")
    Test-Case 'an empty list is a list' $true ($null -eq (Read-AcceptedFindings '{"accepted":[]}').Problem)
    Test-Case 'a file that does not parse is refused' $true ([string](Read-AcceptedFindings '{').Problem).Contains('does not parse')
    $v = Get-FindingsVerdict -Findings $findings -Accepted $acc.Entries
    Test-Case 'the untriaged finding is untriaged' 'cs/sql-injection' (@($v.Untriaged | ForEach-Object { $_.Rule }) -join ',')
    Test-Case 'the triaged one is accepted' 'cs/path-injection' (@($v.Accepted | ForEach-Object { $_.Rule }) -join ',')
    Test-Case 'an entry matching nothing is stale' 'cs/x' (@($v.Stale | ForEach-Object { $_.Rule }) -join ',')
    $moved = @([pscustomobject]@{ Rule = 'cs/path-injection'; Path = 'Services/B.cs'; LineHash = 'dddd:1'; Line = 40; Level = 'error'; SecuritySeverity = ''; Message = 'm' })
    Test-Case 'the same rule and file on a changed line is untriaged again' 1 @((Get-FindingsVerdict -Findings $moved -Accepted $acc.Entries).Untriaged).Count
    Test-Case 'no findings and no entries is clean' 0 @((Get-FindingsVerdict -Findings @() -Accepted @()).Untriaged).Count
    Test-Case 'a finding prints as one line' 'cs/sql-injection [error security-severity 8.8] McpServer/A.cs:12 - This query depends on a user-provided value.' (Format-Finding $findings[0])

    Write-Host ''
    Write-Host '== the repository: what a run reads exists =='
    $acceptedPath = Join-Path $RepoRoot $AcceptedFile
    Test-Case "$($AcceptedFile.Replace('\', '/')) exists" $true (Test-Path -LiteralPath $acceptedPath)
    if (Test-Path -LiteralPath $acceptedPath) {
        $real = Read-AcceptedFindings ([System.IO.File]::ReadAllText($acceptedPath))
        Test-Case 'and every entry in it names its rule, file, line hash and reason' '' ([string]$real.Problem)
    }
    Test-Case 'LICENSE meets step 1 against this repository''s origin' $true ($null -eq (Get-LicenceProblem ([System.IO.File]::ReadAllText((Join-Path $RepoRoot 'LICENSE'))) 'https://github.com/SixFive7/OutlookAI.git'))

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a run on the workstation settles these:'
    Write-Host '  * that the pinned archive downloads, matches and unpacks, and the CLI in it reports the pinned version'
    Write-Host '  * that the database builds and the suite runs with this bundle'
    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

# =============================================================================================
# I/O
# =============================================================================================
function Invoke-Git {
    param([Parameter(Mandatory = $true)] [string[]] $Arguments)
    $out = @(Invoke-NativeCommand { & git -C $RepoRoot @Arguments 2>&1 })
    $code = $LASTEXITCODE
    if ($code -ne 0) { throw "git $($Arguments -join ' ') failed (exit $code): $(($out | Out-String).Trim())" }
    return (($out | Out-String).Trim())
}

function Get-CodeQLReportedVersion([string] $CodeQLExe) {
    if (-not (Test-Path -LiteralPath $CodeQLExe)) { return '' }
    $out = @(Invoke-NativeCommand { & $CodeQLExe version --format=terse 2>&1 })
    if ($LASTEXITCODE -ne 0) { return '' }
    return (($out | Out-String).Trim())
}

# Step 2 in full. Returns codeql.exe's path, or throws with the reason - the caller maps that to exit 3.
function Get-VerifiedBundle {
    param([string] $Root, [string] $LogDir)
    $versionDir = Join-Path $Root "v$CodeQLVersion"
    $codeqlExe = Join-Path $versionDir 'codeql\codeql.exe'
    $recordPath = Join-Path $versionDir 'verified.json'
    $recordJson = ''
    if (Test-Path -LiteralPath $recordPath) { $recordJson = [System.IO.File]::ReadAllText($recordPath) }
    $decision = Get-BundleDecision -RecordJson $recordJson -ReportedVersion (Get-CodeQLReportedVersion $codeqlExe) -ExpectedSha256 $BundleSha256 -ExpectedVersion $CodeQLVersion
    if ($decision -eq 'reuse') {
        Say "  reusing    $codeqlExe - CodeQL $CodeQLVersion, verified $(($recordJson | ConvertFrom-Json).verified)"
        return $codeqlExe
    }
    Say "  bundle     not ready here: $decision"
    New-Item -ItemType Directory -Force -Path $Root | Out-Null
    $archive = Join-Path $Root "codeql-bundle-v$CodeQLVersion-win64.tar.gz"
    $partial = "$archive.partial"
    $problem = 'not checked'
    if (Test-Path -LiteralPath $archive) {
        # A download a run left behind, its unpacking unfinished: checked again, not fetched again.
        Say "  found      $archive - checking it against the pin"
        $problem = Get-ArchiveProblem -Bytes (Get-Item -LiteralPath $archive).Length -Sha256 (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ExpectedBytes $BundleBytes -ExpectedSha256 $BundleSha256
        if ($problem) { Say "  discarded  it: $problem"; Remove-Item -LiteralPath $archive -Force }
    }
    if ($problem) {
        Say "  fetching   $BundleUrl"
        if (-not (Test-Path -LiteralPath $CurlExe)) { throw "Windows' curl.exe is not at $CurlExe." }
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force }
        $r = Invoke-Logged -FilePath $CurlExe -ArgumentList @('--fail', '--location', '--silent', '--show-error', '--retry', '3', '--output', (Format-Argument $partial), $BundleUrl) -LogStem (Join-Path $LogDir 'bundle-download') -TimeoutMinutes $TimeoutMinutes
        if ($r.TimedOut -or $r.ExitCode -ne 0) {
            Show-LogTail $r.Err 10
            throw "the download failed (curl exit $($r.ExitCode), timed out: $($r.TimedOut))."
        }
        $problem = Get-ArchiveProblem -Bytes (Get-Item -LiteralPath $partial).Length -Sha256 (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash -ExpectedBytes $BundleBytes -ExpectedSha256 $BundleSha256
        if ($problem) {
            Remove-Item -LiteralPath $partial -Force
            throw "REFUSING the download: $problem. Nothing was unpacked."
        }
        Move-Item -LiteralPath $partial -Destination $archive -Force
    }
    Say "  verified   $archive - $BundleBytes bytes, SHA-256 $BundleSha256, as pinned"
    if (-not (Test-Path -LiteralPath $TarExe)) { throw "Windows' tar.exe is not at $TarExe." }
    $staging = "$versionDir.partial"
    foreach ($d in @($staging, $versionDir)) { if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force } }
    New-Item -ItemType Directory -Force -Path $staging | Out-Null
    $r = Invoke-Logged -FilePath $TarExe -ArgumentList @('-xzf', (Format-Argument $archive), '-C', (Format-Argument $staging)) -LogStem (Join-Path $LogDir 'bundle-unpack') -TimeoutMinutes $TimeoutMinutes
    if ($r.TimedOut -or $r.ExitCode -ne 0) {
        Show-LogTail $r.Err 10
        throw "unpacking failed (tar exit $($r.ExitCode), timed out: $($r.TimedOut)). The verified archive stays at $archive for the next run."
    }
    Move-Item -LiteralPath $staging -Destination $versionDir
    $reported = Get-CodeQLReportedVersion $codeqlExe
    if ($reported -cne $CodeQLVersion) { throw "the unpacked CLI reports '$reported', not $CodeQLVersion." }
    $rec = [ordered]@{ version = $CodeQLVersion; url = $BundleUrl; bytes = $BundleBytes; sha256 = $BundleSha256; verified = (Get-Date).ToString('o') }
    Set-Content -LiteralPath $recordPath -Value ($rec | ConvertTo-Json) -Encoding ASCII
    Remove-Item -LiteralPath $archive -Force
    Say "  unpacked   $codeqlExe - reports $reported; the archive is deleted, $recordPath keeps the record"
    return $codeqlExe
}

# =============================================================================================
# MAIN
# =============================================================================================
if ($SelfTest) { exit (Invoke-SelfTest) }

$started = Get-Date
$RepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)
$BundleRoot = [System.IO.Path]::GetFullPath($BundleRoot)
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
Say '== Tools/Invoke-CodeQL.ps1 =='
Say "  repository $RepoRoot"

# ---- 1. The licence condition --------------------------------------------------------------
$licensePath = Join-Path $RepoRoot 'LICENSE'
$licenseText = ''
if (Test-Path -LiteralPath $licensePath) { $licenseText = [System.IO.File]::ReadAllText($licensePath) }
$origin = ''
try { $origin = Invoke-Git -Arguments @('remote', 'get-url', 'origin') } catch { $origin = '' }
$licenceProblem = Get-LicenceProblem -LicenseText $licenseText -OriginUrl $origin
if ($licenceProblem) { Say "REFUSED: $licenceProblem."; exit 4 }
$threatModelProblem = Get-ThreatModelProblem $ThreatModel
if ($threatModelProblem) { Say "REFUSED: -ThreatModel $threatModelProblem."; exit 4 }
Say "  licence    MIT, origin $origin - an open-source codebase on GitHub.com, as the CodeQL licence requires"

# ---- The run folder ------------------------------------------------------------------------
$sha = ''
if (-not $FetchOnly) {
    try { $sha = Invoke-Git -Arguments @('rev-parse', '--verify', "$Ref^{commit}") } catch { Say "REFUSED: '$Ref' is not a commit here."; exit 4 }
}
$workTree = ''
try { $workTree = (Invoke-Git -Arguments @('rev-parse', '--show-toplevel')).Replace('/', '\') } catch { $workTree = '' }
if (-not $OutDir) {
    $stem = (Get-Date).ToString('yyyyMMdd-HHmmss')
    if ($sha) { $stem += '-' + $sha.Substring(0, 12) } else { $stem += '-fetch' }
    $OutDir = Join-Path $RepoRoot ".work\codeql\runs\$stem"
}
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
if (Test-IsUnderWorkTreeButNotWork -Path $OutDir -WorkTreeRoot $workTree) { Say "REFUSED: -OutDir $OutDir is inside the working tree $workTree but not under .work\."; exit 4 }
if (Test-IsUnderWorkTreeButNotWork -Path $BundleRoot -WorkTreeRoot $workTree) { Say "REFUSED: -BundleRoot $BundleRoot is inside the working tree $workTree but not under .work\."; exit 4 }
$logDir = Join-Path $OutDir 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
Say "  output     $OutDir"

# ---- 2. The bundle -------------------------------------------------------------------------
Say ''
Say "== 2. The CodeQL bundle, $CodeQLVersion =="
try { $codeql = Get-VerifiedBundle -Root $BundleRoot -LogDir $logDir }
catch { Say "NOT RUN: the bundle could not be made ready - $($_.Exception.Message)"; exit 3 }
if ($FetchOnly) { Say ''; Say "FETCHED. $codeql"; exit 0 }

# ---- 3. The source -------------------------------------------------------------------------
Say ''
Say "== 3. The source: git archive of $sha =="
$subject = Invoke-Git -Arguments @('log', '-1', '--format=%s', $sha)
Say "  commit     $sha - $subject"
$srcDir = Join-Path $OutDir 'src'
$dbDir = Join-Path $OutDir 'db'
$tarFile = Join-Path $OutDir 'source.tar'
New-Item -ItemType Directory -Force -Path $srcDir | Out-Null
$null = Invoke-Git -Arguments @('archive', '--format=tar', '-o', $tarFile, $sha)
$r = Invoke-Logged -FilePath $TarExe -ArgumentList @('-xf', (Format-Argument $tarFile), '-C', (Format-Argument $srcDir)) -LogStem (Join-Path $logDir 'source-unpack') -TimeoutMinutes 10
if ($r.TimedOut -or $r.ExitCode -ne 0) { Show-LogTail $r.Err 10; Say "FAILED: unpacking the source archive (tar exit $($r.ExitCode))."; exit 2 }
Remove-Item -LiteralPath $tarFile -Force
$csCount = @(Get-ChildItem -LiteralPath $srcDir -Recurse -Filter '*.cs' -File).Count
Say "  $csCount C# file(s) in the committed tree"

# ---- 4. The database -----------------------------------------------------------------------
Say ''
Say '== 4. The database: C#, build-mode none - nothing is built =='
$createArgs = @('database', 'create', (Format-Argument $dbDir), '--language=csharp', '--build-mode=none',
    "--source-root=$(Format-Argument $srcDir)", '--threads=0', '--overwrite')
$t0 = Get-Date
$r = Invoke-Logged -FilePath $codeql -ArgumentList $createArgs -LogStem (Join-Path $logDir 'database-create') -TimeoutMinutes $TimeoutMinutes -WorkingDirectory $srcDir
if ($r.TimedOut -or $r.ExitCode -ne 0) {
    Show-LogTail $r.Out 15
    Show-LogTail $r.Err 15
    Say "FAILED: codeql database create (exit $($r.ExitCode), timed out: $($r.TimedOut)). Logs: $logDir"
    exit 2
}
$createMinutes = [math]::Round(((Get-Date) - $t0).TotalMinutes, 1)
Say "  built in $createMinutes minute(s)"

# ---- 5. The queries ------------------------------------------------------------------------
Say ''
Say "== 5. The queries: $QuerySuite, threat model $(Format-ThreatModel $ThreatModel) =="
$sarifPath = Join-Path $OutDir 'results.sarif'
$analyzeArgs = @('database', 'analyze', (Format-Argument $dbDir), $QuerySuite) + @(Get-ThreatModelArguments $ThreatModel) +
    @('--format=sarif-latest', "--output=$(Format-Argument $sarifPath)", '--threads=0', '--no-download')
$t0 = Get-Date
$r = Invoke-Logged -FilePath $codeql -ArgumentList $analyzeArgs -LogStem (Join-Path $logDir 'database-analyze') -TimeoutMinutes $TimeoutMinutes
if ($r.TimedOut -or $r.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $sarifPath)) {
    Show-LogTail $r.Out 15
    Show-LogTail $r.Err 15
    Say "FAILED: codeql database analyze (exit $($r.ExitCode), timed out: $($r.TimedOut)). Logs: $logDir"
    exit 2
}
$analyzeMinutes = [math]::Round(((Get-Date) - $t0).TotalMinutes, 1)
Say "  analysed in $analyzeMinutes minute(s)"

# ---- 6. The verdict ------------------------------------------------------------------------
Say ''
Say '== 6. The verdict =='
$findings = Get-SarifFindings ([System.IO.File]::ReadAllText($sarifPath))
$acceptedText = ''
$acceptedPath = Join-Path $RepoRoot $AcceptedFile
if (Test-Path -LiteralPath $acceptedPath) { $acceptedText = [System.IO.File]::ReadAllText($acceptedPath) }
$accepted = Read-AcceptedFindings $acceptedText
if ($accepted.Problem) { Say "REFUSED: $($AcceptedFile.Replace('\', '/')) - $($accepted.Problem)."; exit 4 }
$verdict = Get-FindingsVerdict -Findings $findings -Accepted $accepted.Entries

$lines = @()
$lines += "CodeQL $CodeQLVersion, $QuerySuite, threat model $(Format-ThreatModel $ThreatModel)"
$lines += "commit $sha - $subject"
$lines += "$csCount C# file(s); database $createMinutes min, analysis $analyzeMinutes min"
$lines += "$(@($findings).Count) result(s): $(@($verdict.Untriaged).Count) untriaged, $(@($verdict.Accepted).Count) accepted as triaged, $(@($verdict.Stale).Count) stale accepted entr(y/ies)"
foreach ($f in @($verdict.Untriaged)) { $lines += "  UNTRIAGED $(Format-Finding $f)" }
foreach ($f in @($verdict.Accepted)) { $lines += "  accepted  $(Format-Finding $f)" }
foreach ($s in @($verdict.Stale)) { $lines += "  STALE     $($s.Rule) $($s.Path) $($s.LineHash) - no result matches it any more; delete it from $($AcceptedFile.Replace('\', '/'))" }
Set-Content -LiteralPath (Join-Path $OutDir 'summary.txt') -Value $lines -Encoding ASCII
$summary = [ordered]@{
    schema = 1; producedBy = 'Tools/Invoke-CodeQL.ps1'; codeql = $CodeQLVersion; suite = $QuerySuite; commit = $sha
    started = $started.ToString('o'); finished = (Get-Date).ToString('o'); csFiles = $csCount
    databaseMinutes = $createMinutes; analysisMinutes = $analyzeMinutes
    results = @($findings).Count; untriaged = @($verdict.Untriaged); accepted = @($verdict.Accepted); stale = @($verdict.Stale)
}
Set-Content -LiteralPath (Join-Path $OutDir 'summary.json') -Value ($summary | ConvertTo-Json -Depth 5) -Encoding ASCII
foreach ($l in $lines) { Say "  $l" }

if (-not $KeepDatabase) {
    foreach ($d in @($srcDir, $dbDir)) { if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force } }
}
Say ''
if (@($verdict.Untriaged).Count -gt 0) {
    Say "FINDINGS: $(@($verdict.Untriaged).Count) result(s) nobody has triaged. Fix each one, or - for a false positive - add it to $($AcceptedFile.Replace('\', '/')) with its reason. $OutDir\summary.txt"
    exit 1
}
Say "CLEAN: no untriaged result in $(@($findings).Count). $OutDir\summary.txt"
exit 0
