#Requires -Version 5.1
<#
.SYNOPSIS
    ONE RUN ON THE BUILD VM: expands a staged revision, restores it from the offline feed, builds
    it, runs the non-live suite into a TRX file and every script self-test, and leaves all of it in
    one results directory for the host to fetch. Runs on OAI-BUILD and nowhere else.

.DESCRIPTION
    RUN ON OAI-BUILD ONLY, and normally only by Testbed/host/Invoke-TestsOnBuildVm.ps1, which copies
    this file in for every run together with a request file, starts it, follows its log and fetches
    what it leaves. Windows PowerShell 5.1: no ternary, no ??, ASCII only.

    WHY THIS EXISTS (Q94, Q102, 2026-10-03). The maintainer's workstation runs no test any more
    except the Exchange-only read-only live tests. The non-live suite and the script self-tests run
    here instead: on a VM with no Office, no mailbox, no network and no state carried from one run
    to the next, because the host restores the same checkpoint before every run and again after it.
    Everything this script writes therefore disappears with the run, and it has no -Execute: a run
    IS the action, on a machine that exists for nothing else.

    WHAT ONE RUN DOES, in order. Every program it starts writes to a file, never to a pipe, and has
    a time limit; a program that outlives its limit is stopped with its whole process tree, so the
    run always ends and always leaves a run.json saying how far it got.

      1. EXPAND <RunRoot>\Source.zip - a `git archive` of one commit, made on the host - into
         <RunRoot>\src. Never the working tree of anything: the host refuses to test what is not a
         commit, so a result always names a revision somebody can check out.
      2. RESTORE it from the offline folder feed only, through a NuGet configuration that clears
         every other source and fallback folder. The VM has no network, and a restore that tried
         nuget.org would fail as a timeout naming the network rather than the package. A package
         the feed lacks is reported by id as PACKAGES-MISSING (exit 20): the host then stages that
         revision's packages from its own restore and runs this again.
      3. BUILD the test project and everything it references, Release, no restore, no MSBuild node
         reuse - a VM that is about to be reverted gains nothing from idle workers.
      4. TEST: `dotnet test --no-build` with the filter `Category!=Live`, ANDed in CODE with whatever
         narrower filter the request carries - a request cannot remove it - into results\trx\suite.trx,
         with a hang timeout so one wedged test ends the run with its name instead of a silence.
      5. SELF-TESTS: every script under Testbed\ and Tools\ that declares a [switch] $SelfTest is run
         with -SelfTest alone, in Windows PowerShell 5.1, one process each, output to a file, time
         limited. Found by the syntax tree, not by a list, so a new self-test is run the day it lands;
         the only ones skipped are named in $SelfTestExclusions below, each with its reason.
      6. RECORD run.json - the verdict, every phase's exit code and duration, the machine's facts -
         and zip the results directory. done.txt, written last, is the completion signal.

    THE VERDICTS, and the exit code each carries:
        PASS              0   built, every selected test passed, every self-test passed
        FAIL              1   built, and a test or a self-test failed
        PACKAGES-MISSING 20   the offline feed lacks a package this revision needs (listed in run.json)
        RESTORE-FAILED   21   the restore failed for any other reason
        BUILD-FAILED     30   the revision does not build
        SUITE-BROKEN     40   dotnet test timed out, crashed, or left no TRX file
        REFUSED           2   the guard, or a request this script will not run
    The HOST decides the final verdict from the TRX file, not from these exit codes alone: a run
    whose filter selected nothing passes here and fails there.

    THE GUARD, AND WHY IT IS STRICTER THAN THE OTHER GUEST SCRIPTS'. Assert-TestbedGuestLocal below
    requires the autologon account AND the computer name OAI-BUILD exactly - not merely an OAI-
    prefix. This builds and executes test code; on the maintainer's workstation that is precisely
    what Q94 moved off it, and on the two Outlook guests it is not what they are for. Both axes
    must hold, and the only way past either is to name the value you mean.

    WHAT IT DOES NOT DO. It never starts Outlook, creates a COM object, touches a profile, a store
    or a mail item, or sets the live tier's opt-in; the VM has no Office for any of that to reach,
    and Category!=Live is not removable. It never writes outside <RunRoot> and the offline feed,
    and the host reverts the VM afterwards either way.

.PARAMETER RequestPath
    The request the host wrote: a JSON object with any of filter, configuration, skipSuite,
    skipSelfTests, selfTestInclude (wildcards on repository-relative paths such as
    'Testbed/host/*'), restoreTimeoutMinutes, buildTimeoutMinutes, testTimeoutMinutes,
    hangTimeoutMinutes and selfTestTimeoutSeconds, plus a revision object recorded as given.
    A file rather than parameters because a test filter is made of &, |, ( and ", and a file is
    the one route that carries it unchanged. Defaults to <RunRoot>\request.json.

.PARAMETER RunRoot
    The run's directory on the VM. Everything this script writes is under it, except the
    packages a restore extracts into the user's NuGet folder.

.PARAMETER OfflineFeed
    The folder feed staged on the VM by Testbed/host/Publish-LiveTierPayload.ps1.

.PARAMETER ExpectedUser
    Accounts this script may run as. The default is the guard; do not widen it.

.PARAMETER ExpectedComputerName
    The one computer this script may run on. The default is the guard; do not widen it.

.PARAMETER SelfTest
    Pure. Checks the filter composition, the request reader, the missing-package parser, the
    self-test discovery and the verdict table against synthetic inputs. Runs anywhere - it is
    ahead of the guard - and reads nothing but this file.

.EXAMPLE
    # On OAI-BUILD, as the host runner does it:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\OutlookAI-Q5\run\Invoke-BuildVmRun.ps1 -RequestPath C:\OutlookAI-Q5\run\request.json

.EXAMPLE
    powershell.exe -NoProfile -File Testbed\guest\Invoke-BuildVmRun.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string]   $RequestPath,
    [string]   $RunRoot = 'C:\OutlookAI-Q5\run',
    [string]   $OfflineFeed = 'C:\OutlookAI-Q5\nuget-offline',
    [string]   $TestProject = 'McpServer\OutlookAI.McpServer.Tests\OutlookAI.McpServer.Tests.csproj',
    [string]   $DotnetRoot = 'C:\Program Files\dotnet',
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerName = 'OAI-BUILD',
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# The fixed parts of a run, in one place.
# ---------------------------------------------------------------------------------------------

# The filter half no request can remove. A live test needs a mailbox; this VM has none, and the
# live tier's own opt-in is never set here either - three locks, of which this is the first.
$LiveExclusion = 'Category!=Live'

# Self-tests that cannot run on this VM, by repository-relative path, each with its reason. Every
# other script that declares -SelfTest is run. An entry whose script no longer declares -SelfTest,
# or no longer exists, is reported, so the list cannot outlive its reasons silently.
$SelfTestExclusions = @(
)

# The verdicts and their exit codes - see the banner.
$Verdicts = [ordered]@{
    'PASS'             = 0
    'FAIL'             = 1
    'REFUSED'          = 2
    'PACKAGES-MISSING' = 20
    'RESTORE-FAILED'   = 21
    'BUILD-FAILED'     = 30
    'SUITE-BROKEN'     = 40
}

# ---------------------------------------------------------------------------------------------
# PURE DECISIONS. No file, no process, no machine state - so -SelfTest can pin every one of them.
# ---------------------------------------------------------------------------------------------

# The filter dotnet test is given: Category!=Live, and the request's own filter in parentheses so
# its | cannot escape the AND. An empty request filter is the whole non-live suite.
function Get-SuiteFilter {
    param([string] $RequestFilter)
    if ([string]::IsNullOrWhiteSpace($RequestFilter)) { return $LiveExclusion }
    return ($LiveExclusion + '&(' + $RequestFilter.Trim() + ')')
}

# A filter this script will not hand to dotnet test, or $null. Not a safety guard - the AND above is
# that - but a filter that cannot be passed through intact would test something nobody asked for.
function Get-FilterRefusal {
    param([string] $RequestFilter)
    if ([string]::IsNullOrWhiteSpace($RequestFilter)) { return $null }
    if ($RequestFilter.Contains('"')) { return 'the filter contains a double quote, which cannot reach dotnet test intact through a command line' }
    $depth = 0
    foreach ($ch in $RequestFilter.ToCharArray()) {
        if ($ch -eq '(') { $depth++ }
        elseif ($ch -eq ')') { $depth--; if ($depth -lt 0) { return 'the filter closes a parenthesis it never opened, which would let part of it escape the Category!=Live AND' } }
    }
    if ($depth -ne 0) { return 'the filter leaves a parenthesis open' }
    return $null
}

# The request, with every value it may carry and its default. Unknown keys are reported, not
# ignored silently: a misspelt key would otherwise run the whole suite while the caller believes
# it asked for less.
function Read-RunRequest {
    param([string] $Json)
    $defaults = [ordered]@{
        filter                 = ''
        configuration          = 'Release'
        skipSuite              = $false
        skipSelfTests          = $false
        selfTestInclude        = @()
        restoreTimeoutMinutes  = 10
        buildTimeoutMinutes    = 20
        testTimeoutMinutes     = 40
        hangTimeoutMinutes     = 15
        selfTestTimeoutSeconds = 300
        revision               = $null
    }
    $request = [ordered]@{}
    foreach ($k in $defaults.Keys) { $request[$k] = $defaults[$k] }
    $unknown = @()
    if (-not [string]::IsNullOrWhiteSpace($Json)) {
        $parsed = $Json | ConvertFrom-Json
        foreach ($p in $parsed.PSObject.Properties) {
            if ($request.Contains($p.Name)) { $request[$p.Name] = $p.Value }
            else { $unknown += $p.Name }
        }
    }
    if (@('Release', 'Debug') -cnotcontains [string]$request['configuration']) { throw "configuration '$($request['configuration'])' is not Release or Debug." }
    foreach ($k in @('restoreTimeoutMinutes', 'buildTimeoutMinutes', 'testTimeoutMinutes', 'hangTimeoutMinutes', 'selfTestTimeoutSeconds')) {
        $v = 0
        if (-not [int]::TryParse([string]$request[$k], [ref] $v) -or $v -lt 1) { throw "$k must be a positive whole number; got '$($request[$k])'." }
        $request[$k] = $v
    }
    $request['skipSuite'] = [bool]$request['skipSuite']
    $request['skipSelfTests'] = [bool]$request['skipSelfTests']
    $request['selfTestInclude'] = @($request['selfTestInclude'] | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { ([string]$_).Replace('\', '/') })
    $request['filter'] = [string]$request['filter']
    return [pscustomobject]@{ Request = $request; Unknown = $unknown }
}

# Package ids a restore log says the feed lacks: NU1101 (no such id), NU1102 (no such version),
# NU1103 (only a prerelease). Each names the package the same way.
function Get-MissingPackages {
    param([string] $RestoreLog)
    $ids = @()
    foreach ($m in [regex]::Matches([string]$RestoreLog, 'NU110[123]:\s+Unable to find package\s+([A-Za-z0-9_.\-]+?)(?:\.\s|\s+with version|\s*$)')) {
        $ids += $m.Groups[1].Value
    }
    return @($ids | Sort-Object -Unique)
}

# Whether a script's own param() block declares a [switch] $SelfTest. Read from the syntax tree, so
# a comment or a string mentioning -SelfTest does not count, and a script that does not parse is
# reported rather than run.
function Test-DeclaresSelfTest {
    param([string] $ScriptText)
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref] $null, [ref] $errors)
    if (@($errors).Count -gt 0) { return 'unparseable' }
    if ($null -eq $ast.ParamBlock) { return 'no' }
    foreach ($p in $ast.ParamBlock.Parameters) {
        if ($p.Name.VariablePath.UserPath -ne 'SelfTest') { continue }
        foreach ($a in $p.Attributes) {
            if ($a -is [System.Management.Automation.Language.TypeConstraintAst] -and $a.TypeName.Name -match '^(?:System\.Management\.Automation\.)?SwitchParameter$|^switch$') { return 'yes' }
        }
    }
    return 'no'
}

# Whether a repository-relative path is selected by the request's include list. Wildcards compare
# ordinally on forward-slash paths; an empty list selects everything. -like is right HERE because
# these are path patterns the caller typed as patterns - none of them holds a bracket - and the
# README's warning about -like is about literal bracketed text.
function Test-SelfTestSelected {
    param([string] $RelativePath, [string[]] $Include)
    $includeList = @($Include | Where-Object { $_ })
    if ($includeList.Count -eq 0) { return $true }
    foreach ($pattern in $includeList) {
        if ($RelativePath -like $pattern) { return $true }
    }
    return $false
}

# The run's verdict from what each phase reported. Pure: the phases decide, in the order they ran.
function Get-RunVerdict {
    param(
        [string] $Restore,      # 'ok' | 'missing' | 'failed' | 'skipped'
        [string] $Build,        # 'ok' | 'failed' | 'skipped'
        [string] $Suite,        # 'passed' | 'failed' | 'broken' | 'skipped'
        [int]    $SelfTestsFailed
    )
    if ($Restore -eq 'missing') { return 'PACKAGES-MISSING' }
    if ($Restore -eq 'failed') { return 'RESTORE-FAILED' }
    if ($Build -eq 'failed') { return 'BUILD-FAILED' }
    if ($Suite -eq 'broken') { return 'SUITE-BROKEN' }
    if ($Suite -eq 'failed' -or $SelfTestsFailed -gt 0) { return 'FAIL' }
    return 'PASS'
}

function Invoke-SelfTest {
    $script:stChecks = 0
    $script:stFailures = @()
    function Check([string] $What, $Expected, $Actual) {
        $script:stChecks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = $Expected -join ' | ' }
        if ($Actual -is [System.Array]) { $a = $Actual -join ' | ' }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else { $script:stFailures += "$What : expected [$e], got [$a]"; Write-Host "  FAIL $What - expected [$e], got [$a]" }
    }

    Write-Host "Invoke-BuildVmRun.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Pure: no file, no process, no machine state."
    Write-Host ''
    Write-Host '== the filter =='
    Check 'no request filter is the whole non-live suite' 'Category!=Live' (Get-SuiteFilter '')
    Check 'whitespace is no filter' 'Category!=Live' (Get-SuiteFilter '   ')
    Check 'a request filter is ANDed in parentheses' 'Category!=Live&(FullyQualifiedName~T1.Foo)' (Get-SuiteFilter 'FullyQualifiedName~T1.Foo')
    Check 'an OR in the request stays inside the parentheses' 'Category!=Live&(FullyQualifiedName~A|FullyQualifiedName~B)' (Get-SuiteFilter 'FullyQualifiedName~A|FullyQualifiedName~B')
    Check 'asking for the live tier still carries the exclusion, so it selects nothing' 'Category!=Live&(Category=Live)' (Get-SuiteFilter 'Category=Live')
    Check 'a balanced filter is accepted' '' ([string](Get-FilterRefusal '(A|B)&C'))
    Check 'a stray closing parenthesis is refused - it would let the rest escape the AND' $true ([bool](Get-FilterRefusal 'A)|(Category=Live'))
    Check 'an unclosed parenthesis is refused' $true ([bool](Get-FilterRefusal '(A|B'))
    Check 'a double quote is refused' $true ([bool](Get-FilterRefusal 'Name="x"'))
    Check 'no filter is not refused' '' ([string](Get-FilterRefusal ''))

    Write-Host ''
    Write-Host '== the request =='
    $r = Read-RunRequest ''
    Check 'an empty request gets every default: Release' 'Release' $r.Request['configuration']
    Check 'and the 40-minute suite limit' 40 $r.Request['testTimeoutMinutes']
    Check 'and no unknown key' 0 @($r.Unknown).Count
    $r = Read-RunRequest '{"filter":"FullyQualifiedName~X","skipSelfTests":true,"selfTestInclude":["Testbed\\host\\*"],"revision":{"sha":"abc"},"filtr":"oops"}'
    Check 'the filter is read' 'FullyQualifiedName~X' $r.Request['filter']
    Check 'a boolean is read' $true $r.Request['skipSelfTests']
    Check 'an include pattern is turned to forward slashes' 'Testbed/host/*' (@($r.Request['selfTestInclude'])[0])
    Check 'a misspelt key is reported, not ignored' 'filtr' (@($r.Unknown) -join ',')
    $threw = $false
    try { [void](Read-RunRequest '{"configuration":"Retail"}') } catch { $threw = $true }
    Check 'an unknown configuration is refused' $true $threw
    $threw = $false
    try { [void](Read-RunRequest '{"testTimeoutMinutes":0}') } catch { $threw = $true }
    Check 'a zero time limit is refused' $true $threw

    Write-Host ''
    Write-Host '== missing packages =='
    $log = @(
        '  Determining projects to restore...'
        'C:\OutlookAI-Q5\run\src\McpServer\X\X.csproj : error NU1101: Unable to find package Contoso.Widgets. No packages exist with this id in source(s): testbed-offline'
        'C:\OutlookAI-Q5\run\src\McpServer\X\X.csproj : error NU1102: Unable to find package xunit with version (>= 2.9.4)'
        'C:\OutlookAI-Q5\run\src\McpServer\Y\Y.csproj : error NU1101: Unable to find package Contoso.Widgets. No packages exist with this id in source(s): testbed-offline'
        'C:\OutlookAI-Q5\run\src\McpServer\Y\Y.csproj : error NU1103: Unable to find package Some.Pre with version (>= 1.0.0)'
    ) -join "`r`n"
    Check 'NU1101, NU1102 and NU1103 each name their package, once' 'Contoso.Widgets | Some.Pre | xunit' (Get-MissingPackages $log)
    Check 'a restore that failed for another reason names none' 0 @(Get-MissingPackages 'error NU1301: Unable to load the service index for source https://api.nuget.org/v3/index.json.').Count
    Check 'a clean restore names none' 0 @(Get-MissingPackages '  Restored C:\x\x.csproj (in 1,2 sec).').Count

    Write-Host ''
    Write-Host '== which scripts have a self-test =='
    Check 'a [switch] $SelfTest parameter is found' 'yes' (Test-DeclaresSelfTest "[CmdletBinding()]`nparam([string] `$VMName, [switch] `$SelfTest)`n'body'")
    Check 'one in a parameter set is found' 'yes' (Test-DeclaresSelfTest "param([Parameter(Mandatory = `$true, ParameterSetName = 'SelfTest')] [switch] `$SelfTest)")
    Check 'a mention in a comment is not' 'no' (Test-DeclaresSelfTest "# run me with -SelfTest`nparam([string] `$Name)")
    Check 'a [string] $SelfTest is not a switch' 'no' (Test-DeclaresSelfTest 'param([string] $SelfTest)')
    Check 'a script with no param() block has none' 'no' (Test-DeclaresSelfTest "'just a body'")
    Check 'a script that does not parse is reported as such' 'unparseable' (Test-DeclaresSelfTest 'param([switch] $SelfTest')
    Check 'no include list selects every script' $true (Test-SelfTestSelected 'Testbed/host/Restart-Guest.ps1' @())
    Check 'an include list selects its matches' $true (Test-SelfTestSelected 'Testbed/host/Restart-Guest.ps1' @('Testbed/host/*'))
    Check 'and only them' $false (Test-SelfTestSelected 'Tools/Switch-AddInBuild.ps1' @('Testbed/host/*'))
    $thisText = [System.IO.File]::ReadAllText($PSCommandPath)
    Check 'this script declares its own -SelfTest, so a run tests this file too' 'yes' (Test-DeclaresSelfTest $thisText)

    Write-Host ''
    Write-Host '== the verdict =='
    Check 'everything ok is PASS' 'PASS' (Get-RunVerdict -Restore ok -Build ok -Suite passed -SelfTestsFailed 0)
    Check 'a skipped suite with passing self-tests is PASS' 'PASS' (Get-RunVerdict -Restore skipped -Build skipped -Suite skipped -SelfTestsFailed 0)
    Check 'a failing test is FAIL' 'FAIL' (Get-RunVerdict -Restore ok -Build ok -Suite failed -SelfTestsFailed 0)
    Check 'a failing self-test is FAIL' 'FAIL' (Get-RunVerdict -Restore ok -Build ok -Suite passed -SelfTestsFailed 1)
    Check 'a missing package outranks everything after it' 'PACKAGES-MISSING' (Get-RunVerdict -Restore missing -Build skipped -Suite skipped -SelfTestsFailed 3)
    Check 'a build failure outranks a failing self-test' 'BUILD-FAILED' (Get-RunVerdict -Restore ok -Build failed -Suite skipped -SelfTestsFailed 1)
    Check 'a broken suite is SUITE-BROKEN, not FAIL' 'SUITE-BROKEN' (Get-RunVerdict -Restore ok -Build ok -Suite broken -SelfTestsFailed 0)
    Check 'every verdict has an exit code, and PASS alone has 0' 'PASS' (@($Verdicts.Keys | Where-Object { $Verdicts[$_] -eq 0 }) -join ',')

    Write-Host ''
    Write-Host '== the guard is a real one =='
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $guard = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Assert-TestbedGuestLocal' }, $true))
    Check 'Assert-TestbedGuestLocal is defined once' 1 $guard.Count
    $guardText = ''
    if ($guard.Count -eq 1) { $guardText = $guard[0].Body.Extent.Text }
    Check 'it reads the account' $true $guardText.Contains('$env:USERNAME')
    Check 'it reads the computer name' $true $guardText.Contains('$env:COMPUTERNAME')
    Check 'it throws' $true ($guardText -match '\bthrow\b')
    Check 'it compares the whole computer name, not a prefix' $true ($guardText.Contains('[string]::Equals(') -and -not $guardText.Contains('StartsWith('))

    Write-Host ''
    Write-Host "$($script:stChecks) check(s), $($script:stFailures.Count) failure(s)."
    if ($script:stFailures.Count -gt 0) { return 1 }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# ---------------------------------------------------------------------------------------------
# THE GUARD. First, before any directory, file or process. See the banner.
# ---------------------------------------------------------------------------------------------
function Assert-TestbedGuestLocal {
    $who = $env:USERNAME
    $userOk = $false
    foreach ($candidate in $ExpectedUser) {
        if ($who -eq $candidate) { $userOk = $true }
    }
    $machine = $env:COMPUTERNAME
    $machineOk = [string]::Equals($machine, $ExpectedComputerName, [System.StringComparison]::OrdinalIgnoreCase)
    if ($userOk -and $machineOk) { return }

    throw @"
REFUSING TO RUN.

  logged on as : '$who'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$machine'      (must be exactly: '$ExpectedComputerName')

This builds a revision and executes its test code. Q94 moved that OFF the maintainer's workstation,
and the two Outlook guests are not for it either: it runs on the build VM, OutlookAI-Build, whose
computer name is OAI-BUILD, and nowhere else. Testbed/host/Invoke-TestsOnBuildVm.ps1 is how to run
it. Do not 'fix' this by widening either default; the defaults are the guard.
"@
}

Assert-TestbedGuestLocal

# ---------------------------------------------------------------------------------------------
# Everything below runs on OAI-BUILD.
# ---------------------------------------------------------------------------------------------
if (-not $PSBoundParameters.ContainsKey('RequestPath')) { $RequestPath = Join-Path $RunRoot 'request.json' }
$ResultsDir = Join-Path $RunRoot 'results'
$SourceZip = Join-Path $RunRoot 'Source.zip'
$SourceRoot = Join-Path $RunRoot 'src'
$OfflineConfig = Join-Path $RunRoot 'NuGet.offline.config'
$RunLog = Join-Path $ResultsDir 'run.log'
$RunJson = Join-Path $ResultsDir 'run.json'
$ResultsZip = Join-Path $RunRoot 'results.zip'
$DoneFile = Join-Path $RunRoot 'done.txt'

foreach ($stale in @($ResultsZip, $DoneFile)) {
    if (Test-Path -LiteralPath $stale) { Remove-Item -LiteralPath $stale -Force }
}
if (Test-Path -LiteralPath $ResultsDir) { Remove-Item -LiteralPath $ResultsDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $ResultsDir | Out-Null

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    # Retried, because the log is read while it is written - the host follows it. On the first
    # run (2026-10-03) the second line, a second after the file was created, never reached it and
    # survived only in the process's stdout; why the append failed was not established.
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try { Add-Content -LiteralPath $RunLog -Value $line -Encoding UTF8 -ErrorAction Stop; break }
        catch { Start-Sleep -Milliseconds 50 }
    }
}

# Every process this script starts: a file for stdout, a file for stderr, a time limit, and the
# whole tree stopped at the limit. Start-Process -PassThru drops the native handle once the process
# exits unless .Handle is read first, after which ExitCode reads $null - measured on OAI-INDEXED
# 2026-09-17 by guest/Install-DotnetSdk.ps1, which this restates.
function Stop-ProcessTree([int] $RootId) {
    $all = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue)
    $ids = New-Object 'System.Collections.Generic.List[int]'
    $ids.Add($RootId)
    for ($i = 0; $i -lt $ids.Count; $i++) {
        foreach ($p in $all) { if ([int]$p.ParentProcessId -eq $ids[$i] -and -not $ids.Contains([int]$p.ProcessId)) { $ids.Add([int]$p.ProcessId) } }
    }
    for ($i = $ids.Count - 1; $i -ge 0; $i--) {
        try { Stop-Process -Id $ids[$i] -Force -ErrorAction Stop } catch { }
    }
    return $ids.Count
}

function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)] [string]   $FilePath,
        [Parameter(Mandatory = $true)] [string[]] $Arguments,
        [Parameter(Mandatory = $true)] [string]   $LogName,
        [Parameter(Mandatory = $true)] [int]      $TimeoutSeconds,
        [string] $WorkingDirectory
    )
    $out = Join-Path $ResultsDir "$LogName.out.txt"
    $err = Join-Path $ResultsDir "$LogName.err.txt"
    $parent = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }

    # Start-Process joins the list with spaces and quotes nothing, so an argument holding a space
    # or one of the filter's operators is quoted here - the filter's & and | are only special to
    # cmd.exe, which is not in this path, but its spaces are.
    $quoted = @()
    foreach ($a in $Arguments) {
        if ($a -match '[\s&|()]') { $quoted += ('"' + $a + '"') } else { $quoted += $a }
    }
    $startArgs = @{
        FilePath               = $FilePath
        ArgumentList           = $quoted
        RedirectStandardOutput = $out
        RedirectStandardError  = $err
        NoNewWindow            = $true
        PassThru               = $true
    }
    if ($WorkingDirectory) { $startArgs['WorkingDirectory'] = $WorkingDirectory }
    $t0 = Get-Date
    $process = Start-Process @startArgs
    $null = $process.Handle
    $exited = $process.WaitForExit($TimeoutSeconds * 1000)
    if ($exited) { $process.WaitForExit() }
    $stopped = 0
    if (-not $exited) { $stopped = Stop-ProcessTree -RootId $process.Id }
    $code = $null
    if ($exited) { $code = $process.ExitCode }
    return [pscustomobject]@{
        ExitCode = $code
        TimedOut = (-not $exited)
        Stopped  = $stopped
        Seconds  = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
        Out      = $out
        Err      = $err
    }
}

function Get-LogText($result) {
    $text = ''
    foreach ($f in @($result.Out, $result.Err)) {
        if (Test-Path -LiteralPath $f) {
            $chunk = [System.IO.File]::ReadAllText($f)
            if ($chunk) { $text += $chunk + "`n" }
        }
    }
    return $text
}

function Show-Tail([string] $Text, [int] $Lines = 30) {
    $all = @($Text -split "`r?`n" | Where-Object { $_.Trim() })
    $start = [Math]::Max(0, $all.Count - $Lines)
    for ($i = $start; $i -lt $all.Count; $i++) { Say ("    | " + $all[$i]) }
}

# What the run reports, built up phase by phase and written once at the end.
$record = [ordered]@{
    schema        = 1
    script        = 'Testbed/guest/Invoke-BuildVmRun.ps1'
    startedUtc    = [DateTime]::UtcNow.ToString('o')
    finishedUtc   = $null
    verdict       = $null
    exitCode      = $null
    revision      = $null
    request       = $null
    machine       = $null
    phases        = [ordered]@{}
    suite         = $null
    selfTests     = @()
    selfTestSkips = @()
    missingPackages = @()
    notes         = @()
}

function Complete-Run([string] $Verdict) {
    if ($null -ne $script:memoryJob) {
        try { Stop-Job -Job $script:memoryJob; Remove-Job -Job $script:memoryJob -Force } catch { }
        try {
            $rows = @(Import-Csv -LiteralPath $MemoryLog)
            if ($rows.Count -gt 0) {
                $record['memory'] = [ordered]@{
                    samples         = $rows.Count
                    minFreeMB       = ($rows | ForEach-Object { [int]$_.freeMB } | Measure-Object -Minimum).Minimum
                    maxCommittedMB  = ($rows | ForEach-Object { [int]$_.committedMB } | Measure-Object -Maximum).Maximum
                    physicalMB      = [int]($record['machine'].memoryGB * 1024)
                }
                Say ("memory: lowest free {0} MB, highest committed {1} MB, over {2} sample(s)" -f $record['memory'].minFreeMB, $record['memory'].maxCommittedMB, $rows.Count)
            }
        }
        catch { Say "memory: could not summarise $MemoryLog - $($_.Exception.Message)" }
    }
    $record['verdict'] = $Verdict
    $record['exitCode'] = $Verdicts[$Verdict]
    $record['finishedUtc'] = [DateTime]::UtcNow.ToString('o')
    ($record | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $RunJson -Encoding UTF8
    Say "VERDICT: $Verdict (exit $($Verdicts[$Verdict]))"
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        if (Test-Path -LiteralPath $ResultsZip) { Remove-Item -LiteralPath $ResultsZip -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($ResultsDir, $ResultsZip)
    }
    catch { Say "could not zip the results: $($_.Exception.Message)" }
    # Written LAST: its existence is what the host waits for.
    Set-Content -LiteralPath $DoneFile -Value ([string]$Verdicts[$Verdict]) -Encoding ASCII
    exit $Verdicts[$Verdict]
}

# ---------------------------------------------------------------------------------------------
# The request.
# ---------------------------------------------------------------------------------------------
Say "Invoke-BuildVmRun on $env:COMPUTERNAME as $env:USERNAME, PowerShell $($PSVersionTable.PSVersion)"
$requestText = ''
if (Test-Path -LiteralPath $RequestPath) { $requestText = [System.IO.File]::ReadAllText($RequestPath) }
else { Say "no request at $RequestPath - every default applies" }
try { $read = Read-RunRequest $requestText }
catch { Say "REFUSED: the request cannot be read: $($_.Exception.Message)"; Complete-Run 'REFUSED' }
$request = $read.Request
foreach ($u in @($read.Unknown)) { Say "REFUSED: the request carries '$u', which this script does not know. A misspelt key would otherwise be ignored and the run would test something other than what was asked."; }
if (@($read.Unknown).Count -gt 0) { Complete-Run 'REFUSED' }
$refusal = Get-FilterRefusal $request['filter']
if ($refusal) { Say "REFUSED: $refusal."; Complete-Run 'REFUSED' }
$record['revision'] = $request['revision']
$request.Remove('revision')
$record['request'] = $request

# ---------------------------------------------------------------------------------------------
# The machine, for the record - and so a result that differs from the workstation's can be
# explained by something written down rather than remembered.
# ---------------------------------------------------------------------------------------------
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$cs = Get-CimInstance -ClassName Win32_ComputerSystem
$dotnetExe = Join-Path $DotnetRoot 'dotnet.exe'
$record['machine'] = [ordered]@{
    computerName      = $env:COMPUTERNAME
    user              = $env:USERNAME
    os                = "$($os.Caption) $($os.Version) build $($os.BuildNumber)"
    culture           = (Get-Culture).Name
    uiCulture         = (Get-UICulture).Name
    timeZone          = (Get-TimeZone).Id
    logicalProcessors = [int]$cs.NumberOfLogicalProcessors
    memoryGB          = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
    powershell        = $PSVersionTable.PSVersion.ToString()
    sessionId         = (Get-Process -Id $PID).SessionId
    outlookInstalled  = (Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application')
}
Say ("machine: {0}; {1} / {2}; {3}; {4} CPU, {5} GB; session {6}; Outlook.Application registered: {7}" -f $record['machine'].os, $record['machine'].culture, $record['machine'].uiCulture, $record['machine'].timeZone, $record['machine'].logicalProcessors, $record['machine'].memoryGB, $record['machine'].sessionId, $record['machine'].outlookInstalled)

# Memory, sampled every 5 s for the whole run into results\memory.csv: free physical and committed
# megabytes. It is what the VM's RAM was sized from, and what says when it stops being enough.
$MemoryLog = Join-Path $ResultsDir 'memory.csv'
Set-Content -LiteralPath $MemoryLog -Value 'time,freeMB,committedMB' -Encoding ASCII
$memoryJob = Start-Job -ArgumentList $MemoryLog -ScriptBlock {
    param($path)
    while ($true) {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        Add-Content -LiteralPath $path -Encoding ASCII -Value ('{0:HH:mm:ss},{1},{2}' -f (Get-Date), [int]($os.FreePhysicalMemory / 1024), [int](($os.TotalVirtualMemorySize - $os.FreeVirtualMemory) / 1024))
        Start-Sleep -Seconds 5
    }
}

$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_NOLOGO = '1'
# The live tier's per-run opt-in is a computer name set for one run; it must never be set here.
if ($env:OUTLOOKAI_LIVE_OPT_IN) { Say 'OUTLOOKAI_LIVE_OPT_IN was set in this environment; it is cleared for this run.'; $env:OUTLOOKAI_LIVE_OPT_IN = $null }

# ---------------------------------------------------------------------------------------------
# 1. Expand.
# ---------------------------------------------------------------------------------------------
$t0 = Get-Date
if (-not (Test-Path -LiteralPath $SourceZip)) { Say "REFUSED: no source archive at $SourceZip."; Complete-Run 'REFUSED' }
if (Test-Path -LiteralPath $SourceRoot) { Remove-Item -LiteralPath $SourceRoot -Recurse -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($SourceZip, $SourceRoot)
$fileCount = @(Get-ChildItem -LiteralPath $SourceRoot -Recurse -File).Count
$record['phases']['expand'] = [ordered]@{ seconds = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1); files = $fileCount }
Say "expanded $fileCount file(s) into $SourceRoot in $($record['phases']['expand'].seconds) s"

$restoreState = 'skipped'
$buildState = 'skipped'
$suiteState = 'skipped'

# ---------------------------------------------------------------------------------------------
# 2-4. Restore, build, test.
# ---------------------------------------------------------------------------------------------
if (-not $request['skipSuite']) {
    $projectFull = Join-Path $SourceRoot $TestProject
    if (-not (Test-Path -LiteralPath $projectFull)) { Say "REFUSED: the revision has no $TestProject."; Complete-Run 'REFUSED' }
    if (-not (Test-Path -LiteralPath $dotnetExe)) { Say "SUITE-BROKEN: no dotnet at $dotnetExe - this VM's base checkpoint has no SDK."; $record['notes'] += "no dotnet at $dotnetExe"; Complete-Run 'SUITE-BROKEN' }

    # The offline configuration, outside the source tree so the revision is built exactly as
    # archived. <clear/> is the load-bearing line: see guest/Install-DotnetSdk.ps1.
    $feedForConfig = [System.Security.SecurityElement]::Escape($OfflineFeed)
    Set-Content -LiteralPath $OfflineConfig -Encoding UTF8 -Value @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="testbed-offline" value="$feedForConfig" />
  </packageSources>
  <fallbackPackageFolders>
    <clear />
  </fallbackPackageFolders>
</configuration>
"@

    Say "restore: $TestProject, offline feed $OfflineFeed only"
    $restore = Invoke-Logged -FilePath $dotnetExe -Arguments @('restore', $projectFull, '--configfile', $OfflineConfig) -LogName 'restore' -TimeoutSeconds ($request['restoreTimeoutMinutes'] * 60) -WorkingDirectory $SourceRoot
    $restoreText = Get-LogText $restore
    $record['phases']['restore'] = [ordered]@{ seconds = $restore.Seconds; exitCode = $restore.ExitCode; timedOut = $restore.TimedOut }
    if ($restore.ExitCode -eq 0 -and -not $restore.TimedOut) {
        $restoreState = 'ok'
        Say "restore: ok in $($restore.Seconds) s"
    }
    else {
        $missing = @(Get-MissingPackages $restoreText)
        if ($missing.Count -gt 0) {
            $restoreState = 'missing'
            $record['missingPackages'] = $missing
            Say "restore: the offline feed lacks $($missing.Count) package(s): $($missing -join ', ')"
        }
        else {
            $restoreState = 'failed'
            Say "restore: FAILED (exit $($restore.ExitCode), timed out: $($restore.TimedOut))"
            Show-Tail $restoreText
        }
    }

    if ($restoreState -eq 'ok') {
        Say "build: $($request['configuration']), no restore"
        $build = Invoke-Logged -FilePath $dotnetExe -Arguments @('build', $projectFull, '-c', $request['configuration'], '--no-restore', '-nodeReuse:false') -LogName 'build' -TimeoutSeconds ($request['buildTimeoutMinutes'] * 60) -WorkingDirectory $SourceRoot
        $record['phases']['build'] = [ordered]@{ seconds = $build.Seconds; exitCode = $build.ExitCode; timedOut = $build.TimedOut }
        if ($build.ExitCode -eq 0 -and -not $build.TimedOut) {
            $buildState = 'ok'
            Say "build: ok in $($build.Seconds) s"
        }
        else {
            $buildState = 'failed'
            Say "build: FAILED (exit $($build.ExitCode), timed out: $($build.TimedOut)) in $($build.Seconds) s"
            Show-Tail (Get-LogText $build) 40
        }
    }

    if ($buildState -eq 'ok') {
        $filter = Get-SuiteFilter $request['filter']
        $trxDir = Join-Path $ResultsDir 'trx'
        New-Item -ItemType Directory -Force -Path $trxDir | Out-Null
        Say "test: --filter $filter"
        $testArgs = @('test', $projectFull, '-c', $request['configuration'], '--no-build', '--no-restore',
            '--filter', $filter,
            '--logger', 'trx;LogFileName=suite.trx',
            '--results-directory', $trxDir,
            '--blame-hang-timeout', ("{0}m" -f $request['hangTimeoutMinutes']),
            '--blame-hang-dump-type', 'none')
        $test = Invoke-Logged -FilePath $dotnetExe -Arguments $testArgs -LogName 'test' -TimeoutSeconds ($request['testTimeoutMinutes'] * 60) -WorkingDirectory $SourceRoot
        $trx = Join-Path $trxDir 'suite.trx'
        $testText = Get-LogText $test
        $summaryLine = ''
        foreach ($line in ($testText -split "`r?`n")) { if ($line -match '(Passed|Failed)!\s+-\s+Failed:') { $summaryLine = $line.Trim() } }
        $record['phases']['test'] = [ordered]@{ seconds = $test.Seconds; exitCode = $test.ExitCode; timedOut = $test.TimedOut }
        $record['suite'] = [ordered]@{ filter = $filter; trx = 'trx/suite.trx'; trxPresent = (Test-Path -LiteralPath $trx); summaryLine = $summaryLine }
        if ($test.TimedOut) {
            $suiteState = 'broken'
            Say "test: STOPPED after $($request['testTimeoutMinutes']) min ($($test.Stopped) process(es) ended)"
            Show-Tail $testText 40
        }
        elseif (-not (Test-Path -LiteralPath $trx)) {
            $suiteState = 'broken'
            Say "test: no TRX file (exit $($test.ExitCode)) - the test host did not finish its run"
            Show-Tail $testText 40
        }
        elseif ($test.ExitCode -eq 0) {
            $suiteState = 'passed'
            Say "test: exit 0 in $($test.Seconds) s - $summaryLine"
        }
        else {
            $suiteState = 'failed'
            Say "test: exit $($test.ExitCode) in $($test.Seconds) s - $summaryLine"
            Show-Tail $testText 40
        }
    }
}
else {
    Say 'suite: skipped by the request'
}

# ---------------------------------------------------------------------------------------------
# 5. Self-tests. Run whatever the build did: they do not depend on it.
# ---------------------------------------------------------------------------------------------
$selfTestsFailed = 0
if (-not $request['skipSelfTests'] -and $restoreState -ne 'missing') {
    $t0 = Get-Date
    # An EMPTY .git directory at the root of the expanded archive, made HERE - after the build and
    # the suite, never before them. The archive carries no .git and the VM has no git (D5 of
    # Docs/overnight-review-2026-10-03.md); Testbed/host/New-LiveTestSettings.ps1 -SelfTest asks
    # only whether a .git exists above a path, to tell a working tree from the rest of the disk,
    # and two of its checks fail without one. Made before the build it could be read by the
    # SDK's source-control discovery as a repository that is not one.
    $gitMarker = Join-Path $SourceRoot '.git'
    if (-not (Test-Path -LiteralPath $gitMarker)) {
        New-Item -ItemType Directory -Force -Path $gitMarker | Out-Null
        $record['notes'] += 'an empty .git directory was made at the root of the expanded archive before the self-tests (New-LiveTestSettings.ps1 -SelfTest tells a working tree by it)'
    }
    $candidates = @()
    foreach ($root in @('Testbed', 'Tools')) {
        $dir = Join-Path $SourceRoot $root
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -Recurse -File -Filter '*.ps1' | Sort-Object FullName)) {
            $relative = $f.FullName.Substring($SourceRoot.Length + 1).Replace('\', '/')
            $declares = Test-DeclaresSelfTest ([System.IO.File]::ReadAllText($f.FullName))
            if ($declares -eq 'unparseable') { $record['notes'] += "$relative does not parse under Windows PowerShell $($PSVersionTable.PSVersion); its self-test, if any, could not be found"; $selfTestsFailed++; continue }
            if ($declares -eq 'yes') { $candidates += [pscustomobject]@{ Relative = $relative; Full = $f.FullName } }
        }
    }
    $excludedPaths = @($SelfTestExclusions | ForEach-Object { $_.Path })
    foreach ($e in $SelfTestExclusions) {
        if (@($candidates | Where-Object { $_.Relative -eq $e.Path }).Count -eq 0) {
            $record['notes'] += "the self-test exclusion for $($e.Path) names a script that no longer declares -SelfTest; delete the entry"
        }
    }
    $selected = @($candidates | Where-Object { Test-SelfTestSelected $_.Relative $request['selfTestInclude'] })
    Say ("self-tests: {0} script(s) declare -SelfTest; {1} selected, {2} excluded by reason" -f $candidates.Count, $selected.Count, @($selected | Where-Object { $excludedPaths -contains $_.Relative }).Count)
    $n = 0
    foreach ($c in $selected) {
        $n++
        $exclusion = @($SelfTestExclusions | Where-Object { $_.Path -eq $c.Relative })
        if ($exclusion.Count -gt 0) {
            $record['selfTestSkips'] += [ordered]@{ path = $c.Relative; reason = $exclusion[0].Why }
            Say ("  [{0}/{1}] SKIP {2} - {3}" -f $n, $selected.Count, $c.Relative, $exclusion[0].Why)
            continue
        }
        $logName = 'selftests\' + ($c.Relative -replace '[\\/:]', '_')
        $r = Invoke-Logged -FilePath (Join-Path $PSHOME 'powershell.exe') -Arguments @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $c.Full, '-SelfTest') -LogName $logName -TimeoutSeconds $request['selfTestTimeoutSeconds'] -WorkingDirectory $SourceRoot
        $ok = ($r.ExitCode -eq 0 -and -not $r.TimedOut)
        if (-not $ok) { $selfTestsFailed++ }
        $lastLine = ''
        $text = Get-LogText $r
        $tail = @($text -split "`r?`n" | Where-Object { $_.Trim() })
        if ($tail.Count -gt 0) { $lastLine = $tail[$tail.Count - 1].Trim() }
        $record['selfTests'] += [ordered]@{
            path     = $c.Relative
            exitCode = $r.ExitCode
            timedOut = $r.TimedOut
            seconds  = $r.Seconds
            passed   = $ok
            lastLine = $lastLine
            log      = ($logName.Replace('\', '/') + '.out.txt')
        }
        $state = 'ok  '
        if (-not $ok) { $state = 'FAIL' }
        Say ("  [{0}/{1}] {2} {3} ({4} s) - {5}" -f $n, $selected.Count, $state, $c.Relative, $r.Seconds, $lastLine)
    }
    $record['phases']['selfTests'] = [ordered]@{ seconds = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1); run = @($record['selfTests']).Count; failed = $selfTestsFailed }
}
elseif ($request['skipSelfTests']) {
    Say 'self-tests: skipped by the request'
}

# ---------------------------------------------------------------------------------------------
# 6. The verdict.
# ---------------------------------------------------------------------------------------------
Complete-Run (Get-RunVerdict -Restore $restoreState -Build $buildState -Suite $suiteState -SelfTestsFailed $selfTestsFailed)
