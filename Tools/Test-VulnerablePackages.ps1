#Requires -Version 5.1
<#
    ============================================================================================
    WRITTEN AND FIRST RUN 2026-10-04 (Q125), ON THE MAINTAINER'S WORKSTATION, .NET SDK 10.0.401.
    ============================================================================================

    The six projects under McpServer\: clean, in 26 seconds. The same script pointed at a scratch
    repository whose one project took Newtonsoft.Json 12.0.1: exit 1, naming GHSA-5crp-9r3c-p9vr
    (High) - and `dotnet list package` itself exited 0 on that, which is why only the JSON counts.
    A project whose only source could not be reached: `dotnet list` exit 1 and plain-text errors,
    no JSON - read here as not checked.

.SYNOPSIS
    Fails when a NuGet package the MCP server's projects use - directly or through another package -
    has a known vulnerability: `dotnet list package --vulnerable --include-transitive` for every
    project under McpServer\, read as JSON.

.DESCRIPTION
    WHY THIS EXISTS. Until 2026-10-03 GitHub's dependency review flagged a pull request that brought
    in a known-vulnerable package. It went with the CI ("No CI pipelines on github"), and the
    maintainer's answer to Q125 leaned towards this check as a release gate in its place.
    Tools/Publish-Release.ps1 runs it before every release and refuses to release on anything but
    exit 0.

    WHAT IT CHECKS. Every *.csproj under McpServer\ - the server, the COM host and Core that ship,
    and the test and tool projects that run on the test VMs against real mailboxes, as the review
    it replaces read every manifest. The add-in and its helper are not in it because they have no
    package to check: both are old-style projects whose references are the .NET Framework's and
    Office's own. `dotnet list package` restores a project before it lists it, so nothing needs
    building first; the advisories come from nuget.org, which serves the GitHub Advisory Database.

    A RUN THAT COULD NOT ASK IS NOT A PASS. The command exits 0 when it finds a vulnerable package -
    the JSON is the only verdict - and it says nothing at all when no source it asked publishes
    advisories. So a non-zero exit, output that is not its JSON, a project missing from it, or a
    source list without nuget.org is "not checked", never "clean".

    EXIT CODES: 0 no known vulnerability; 1 at least one vulnerable package; 3 not checked - the
    network, the restore, or no nuget.org; 4 refused.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`, ASCII only.

.PARAMETER RepoRoot
    Repository root. Defaults to the folder above this script's.

.PARAMETER OutDir
    Where the per-project output lands. Default .work\vulnerable-packages\<time> under the
    repository root.

.PARAMETER SelfTest
    Run the pure decisions against synthetic inputs and exit. No network, no dotnet.

.EXAMPLE
    pwsh -File Tools/Test-VulnerablePackages.ps1
    pwsh -File Tools/Test-VulnerablePackages.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $RepoRoot,
    [string] $OutDir,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# Defaults that need $PSScriptRoot are set HERE, not in param(): Windows PowerShell 5.1 leaves it
# empty while param() defaults are evaluated (Q78).
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }

# The one source here that publishes advisories. Without it in the list, silence means nothing.
$AdvisorySource = 'https://api.nuget.org/v3/index.json'
$ProjectRoot = 'McpServer'
$ListTimeoutMinutes = 10

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# A program with its output going STRAIGHT TO FILES and a deadline. Restated from
# Tools/Publish-Release.ps1, which says why.
function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)] [string] $FilePath,
        [Parameter(Mandatory = $true)] [string[]] $ArgumentList,
        [Parameter(Mandatory = $true)] [string] $LogStem,
        [int] $TimeoutMinutes = 20
    )
    $out = "$LogStem.out.txt"
    $err = "$LogStem.err.txt"
    foreach ($f in @($out, $err)) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } }
    $p = Start-Process -FilePath $FilePath -ArgumentList ($ArgumentList -join ' ') -RedirectStandardOutput $out -RedirectStandardError $err -NoNewWindow -PassThru
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = -1; TimedOut = $true; Out = $out; Err = $err }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false; Out = $out; Err = $err }
}

# =============================================================================================
# PURE DECISIONS - everything -SelfTest checks
# =============================================================================================

# One project's `dotnet list package --vulnerable --include-transitive --format json`, read.
# Problem is set when the output cannot be trusted as an answer; Findings lists every vulnerable
# package, top-level and transitive, once per framework and advisory.
function Read-VulnerabilityReport {
    param([string] $Json, [string] $ExpectedSource)
    $result = [pscustomobject]@{ Findings = @(); Problem = $null; Projects = 0 }
    if ([string]::IsNullOrWhiteSpace($Json)) { $result.Problem = 'no output at all'; return $result }
    try { $doc = $Json | ConvertFrom-Json } catch { $result.Problem = 'the output is not the JSON report'; return $result }
    if ($null -eq $doc.projects) { $result.Problem = 'the report names no project'; return $result }
    $sources = @($doc.sources | ForEach-Object { ([string]$_).TrimEnd('/') })
    if ($sources -notcontains $ExpectedSource.TrimEnd('/')) {
        $result.Problem = "nuget.org ($ExpectedSource) was not among the sources asked ($($sources -join ', ')), and it is the one that publishes advisories - silence from the others proves nothing"
        return $result
    }
    $findings = @()
    foreach ($project in @($doc.projects)) {
        if ($null -eq $project) { continue }
        $result.Projects++
        foreach ($fw in @($project.frameworks)) {
            if ($null -eq $fw) { continue }
            foreach ($kind in @('topLevelPackages', 'transitivePackages')) {
                foreach ($pkg in @($fw.$kind)) {
                    if ($null -eq $pkg) { continue }
                    foreach ($v in @($pkg.vulnerabilities)) {
                        if ($null -eq $v) { continue }
                        $findings += [pscustomobject]@{
                            Project = [System.IO.Path]::GetFileNameWithoutExtension([string]$project.path)
                            Framework = [string]$fw.framework; Package = [string]$pkg.id; Version = [string]$pkg.resolvedVersion
                            Transitive = ($kind -eq 'transitivePackages'); Severity = [string]$v.severity; Advisory = [string]$v.advisoryurl
                        }
                    }
                }
            }
        }
    }
    if ($result.Projects -eq 0) { $result.Problem = 'the report names no project'; return $result }
    $result.Findings = $findings
    return $result
}

function Format-VulnerablePackage {
    param($Finding)
    $how = 'direct'
    if ($Finding.Transitive) { $how = 'transitive' }
    return "$($Finding.Package) $($Finding.Version) ($how, $($Finding.Project) $($Finding.Framework)) - $($Finding.Severity): $($Finding.Advisory)"
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

    Write-Host "Tools/Test-VulnerablePackages.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)). No network, no dotnet."
    Write-Host ''
    Write-Host '== the report, read =='
    # The shape SDK 10.0.401 printed on 2026-10-04 for a scratch project on Newtonsoft.Json 12.0.1,
    # with a made-up transitive package added in the same shape.
    $vulnerable = @'
{"version":1,"parameters":"--vulnerable --include-transitive","sources":["https://api.nuget.org/v3/index.json","C:/Program Files (x86)/Microsoft SDKs/NuGetPackages/"],
 "projects":[{"path":"C:/r/McpServer/A/A.csproj","frameworks":[{"framework":"net10.0",
  "topLevelPackages":[{"id":"Newtonsoft.Json","requestedVersion":"12.0.1","resolvedVersion":"12.0.1","vulnerabilities":[{"severity":"High","advisoryurl":"https://github.com/advisories/GHSA-5crp-9r3c-p9vr"}]}],
  "transitivePackages":[{"id":"Example.Transitive","resolvedVersion":"1.0.0","vulnerabilities":[{"severity":"High","advisoryurl":"https://example.invalid/advisory-1"},{"severity":"Moderate","advisoryurl":"https://example.invalid/advisory-2"}]}]}]}]}
'@
    $r = Read-VulnerabilityReport -Json $vulnerable -ExpectedSource $AdvisorySource
    Test-Case 'a report with vulnerable packages is an answer' '' ([string]$r.Problem)
    Test-Case 'every advisory of every package is a finding' 3 @($r.Findings).Count
    Test-Case 'a direct package is named as such' 'Newtonsoft.Json 12.0.1 (direct, A net10.0) - High: https://github.com/advisories/GHSA-5crp-9r3c-p9vr' (Format-VulnerablePackage $r.Findings[0])
    Test-Case 'a transitive one too' $true ($r.Findings[1].Transitive -and $r.Findings[1].Package -ceq 'Example.Transitive')
    $clean = '{"version":1,"parameters":"--vulnerable --include-transitive","sources":["https://api.nuget.org/v3/index.json"],"projects":[{"path":"C:/r/McpServer/A/A.csproj"}]}'
    $r = Read-VulnerabilityReport -Json $clean -ExpectedSource $AdvisorySource
    Test-Case 'a project with no vulnerable package is clean' $true ($null -eq $r.Problem -and @($r.Findings).Count -eq 0 -and $r.Projects -eq 1)
    Test-Case 'a source with a trailing slash still counts' $true ($null -eq (Read-VulnerabilityReport -Json ($clean.Replace('index.json"', 'index.json/"')) -ExpectedSource $AdvisorySource).Problem)

    Write-Host ''
    Write-Host '== a run that could not ask is not a pass =='
    Test-Case 'no output is not checked' 'no output at all' (Read-VulnerabilityReport -Json '' -ExpectedSource $AdvisorySource).Problem
    Test-Case 'the error text an unreachable source prints is not checked' 'the output is not the JSON report' (Read-VulnerabilityReport -Json "error: Unable to load the service index for source https://nuget.invalid/v3/index.json." -ExpectedSource $AdvisorySource).Problem
    Test-Case 'a report naming no project is not checked' 'the report names no project' (Read-VulnerabilityReport -Json '{"sources":["https://api.nuget.org/v3/index.json"],"projects":[]}' -ExpectedSource $AdvisorySource).Problem
    $offline = '{"version":1,"sources":["C:/OutlookAI-Q5/nuget-offline"],"projects":[{"path":"C:/r/McpServer/A/A.csproj"}]}'
    Test-Case 'silence from a feed that publishes no advisories is not checked' $true ([string](Read-VulnerabilityReport -Json $offline -ExpectedSource $AdvisorySource).Problem).Contains('proves nothing')

    Write-Host ''
    Write-Host '== the repository =='
    $projects = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot $ProjectRoot) -Filter '*.csproj' -Recurse -Depth 1 -File)
    Test-Case "there are projects under $ProjectRoot to check" $true ($projects.Count -gt 0)

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE: that nuget.org answers and the projects restore - only a run on the workstation settles that.'
    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

# =============================================================================================
# MAIN
# =============================================================================================
if ($SelfTest) { exit (Invoke-SelfTest) }

$RepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)
if (-not $OutDir) { $OutDir = Join-Path $RepoRoot ('.work\vulnerable-packages\' + (Get-Date).ToString('yyyyMMdd-HHmmss')) }
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$dotnet = @(Get-Command dotnet.exe -CommandType Application -ErrorAction SilentlyContinue)
if ($dotnet.Count -eq 0) { Say 'REFUSED: dotnet.exe is not on PATH - the .NET SDK is a precondition (AGENTS.md).'; exit 4 }
$projects = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot $ProjectRoot) -Filter '*.csproj' -Recurse -Depth 1 -File | Sort-Object FullName)
if ($projects.Count -eq 0) { Say "REFUSED: no project under $ProjectRoot - a check of nothing proves nothing."; exit 4 }
Say "== Tools/Test-VulnerablePackages.ps1: $($projects.Count) project(s), advisories from nuget.org =="

$all = @()
$notChecked = @()
foreach ($proj in $projects) {
    $name = $proj.BaseName
    $r = Invoke-Logged -FilePath $dotnet[0].Source -ArgumentList @('list', ('"' + $proj.FullName + '"'), 'package', '--vulnerable', '--include-transitive', '--format', 'json') -LogStem (Join-Path $OutDir $name) -TimeoutMinutes $ListTimeoutMinutes
    $text = ''
    if (Test-Path -LiteralPath $r.Out) { $text = [System.IO.File]::ReadAllText($r.Out) }
    if ($r.TimedOut -or $r.ExitCode -ne 0) {
        $why = "dotnet list exit $($r.ExitCode), timed out: $($r.TimedOut)"
        $first = @(Get-Content -LiteralPath $r.Out, $r.Err -ErrorAction SilentlyContinue | Where-Object { $_ -match '\S' } | Select-Object -First 2)
        if ($first.Count -gt 0) { $why += ' - ' + ($first -join ' ') }
        $notChecked += "${name}: $why"
        Say "  NOT CHECKED  $name - $why"
        continue
    }
    $report = Read-VulnerabilityReport -Json $text -ExpectedSource $AdvisorySource
    if ($report.Problem) {
        $notChecked += "${name}: $($report.Problem)"
        Say "  NOT CHECKED  $name - $($report.Problem)"
        continue
    }
    $all += @($report.Findings)
    if (@($report.Findings).Count -eq 0) { Say "  clean        $name" }
    foreach ($f in @($report.Findings)) { Say "  VULNERABLE   $(Format-VulnerablePackage $f)" }
}

$lines = @("$($projects.Count) project(s); $(@($all).Count) vulnerable package advisor(y/ies); $($notChecked.Count) project(s) not checked")
foreach ($f in @($all)) { $lines += "  VULNERABLE  $(Format-VulnerablePackage $f)" }
foreach ($n in $notChecked) { $lines += "  NOT CHECKED $n" }
Set-Content -LiteralPath (Join-Path $OutDir 'summary.txt') -Value $lines -Encoding ASCII
Say ''
if (@($all).Count -gt 0) {
    Say "VULNERABLE: $(@($all).Count) known advisor(y/ies) against the packages above. Update the package - or the one that brings it in - and run again. $OutDir\summary.txt"
    exit 1
}
if ($notChecked.Count -gt 0) {
    Say "NOT CHECKED: $($notChecked.Count) project(s) could not be asked, and silence is not a pass. $OutDir\summary.txt"
    exit 3
}
Say "CLEAN: no known vulnerability in any package of $($projects.Count) project(s). $OutDir\summary.txt"
exit 0
