#Requires -Version 5.1
<#
.SYNOPSIS
    Installs, shows or removes the git hooks that keep personal data out of every commit and push
    (Q190 (b)): pre-commit, commit-msg and pre-push, each running Tools/Checks/check-personal-data.ps1.

.DESCRIPTION
    DECIDED BY THE MAINTAINER 2026-11-17 (Q190 (b), QUESTIONS.md). Most of the personal data an audit
    found in this public repository came in through agents' commits, commit messages included. The
    hooks check each commit before it is made and each push before it leaves, on the machine where the
    private list of the real values is: the maintainer's.

    WHERE THEY GO. Into git's common directory, <repository>\.git\hooks, which every worktree of the
    repository uses: one install covers the maintainer's checkout and every agent's worktree, whatever
    branch it has out. Each hook is a three-line sh script - Git for Windows runs hooks in its own sh -
    that hands over to Windows PowerShell 5.1, which every Windows has, running a COPY of the check
    and of its committed baseline that this script puts in <repository>\.git\hooks\outlookai-personal-data\.
    A copy, so that a worktree checked out at an older commit, which has no check at all, is checked
    all the same. Run this again after either file changes; without -Execute it says whether the
    installed copy is current. The private list stays where it is, in the main checkout's
    .work\personal-data\, where the check finds it from any worktree.

    WHAT A HOOK DOES. pre-commit fails a commit that adds an occurrence of personal data to a file;
    commit-msg fails a message that holds any; pre-push fails a push of a commit that did either.
    Each names the file, line and rule, never the value. `--no-verify` skips them, and AGENTS.md
    (Personal data) forbids it.

    WHAT IT REFUSES. A hook of the same name that this script did not write: it would be replaced, so
    it is named and left alone. And core.hooksPath set: git would then run no hook from .git\hooks.

    EXIT CODES: 0 done, or nothing to do; 4 refused; 3 not a git repository, or git failed.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no ??, ASCII only.

.PARAMETER RepoRoot
    A checkout of the repository: the main one or any worktree. Defaults to the folder above this
    script's.

.PARAMETER Execute
    Install or update the hooks. Without it, nothing is written.

.PARAMETER Uninstall
    With -Execute, remove the hooks this script wrote and their folder, and nothing else.

.PARAMETER SelfTest
    Install, update, refuse and uninstall in a temporary folder that stands in for .git\hooks. No git.

.EXAMPLE
    pwsh -File Tools/Install-GitHooks.ps1                 # what is installed, and what -Execute would do
    pwsh -File Tools/Install-GitHooks.ps1 -Execute
    pwsh -File Tools/Install-GitHooks.ps1 -Uninstall -Execute
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')] [string] $RepoRoot,
    [Parameter(ParameterSetName = 'Run')] [switch] $Execute,
    [Parameter(ParameterSetName = 'Run')] [switch] $Uninstall,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)
# Defaults that need $PSScriptRoot are set in the body (Tools/Checks/check-powershell-51.ps1, 1).
if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = Split-Path -Parent $PSScriptRoot }

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$HookNames = @('pre-commit', 'commit-msg', 'pre-push')
# The line every hook this script writes carries: how it tells its own from somebody else's.
$Marker = '# outlookai-personal-data: written by Tools/Install-GitHooks.ps1 (Q190 (b)); run it again to change this hook.'
$CopyFolder = 'outlookai-personal-data'
$CheckFiles = @('Tools\Checks\check-personal-data.ps1', 'Tools\Checks\personal-data.json')

# Every program call goes through Invoke-NativeCommand, so that Windows PowerShell 5.1 cannot end
# the script on a program's first stderr line (Q78). It is one of the helpers the scripts share,
# dot-sourced from Testbed/guest/ScriptHelpers.ps1 (Q166 (b), decision 215 (a)), which says why.
. (Join-Path $PSScriptRoot '..\Testbed\guest\ScriptHelpers.ps1')

# ---------------------------------------------------------------------------------------------
# Pure: the hook text, and what an install or an uninstall does to a hooks folder.
# ---------------------------------------------------------------------------------------------

# A hook: sh, LF line ends, ASCII. "$(dirname "$0")" is the hooks folder whichever worktree runs it.
function Get-HookText([string] $Name) {
    $lines = @(
        '#!/bin/sh',
        $Marker,
        "exec powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"`$(dirname `"`$0`")/$CopyFolder/check-personal-data.ps1`" -Hook $Name `"`$@`""
    )
    return (($lines -join "`n") + "`n")
}

function Test-OurHook([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    return ([System.IO.File]::ReadAllText($Path).Contains($Marker))
}

function Test-SameBytes([string] $A, [string] $B) {
    if (-not (Test-Path -LiteralPath $A -PathType Leaf) -or -not (Test-Path -LiteralPath $B -PathType Leaf)) { return $false }
    $x = [System.IO.File]::ReadAllBytes($A)
    $y = [System.IO.File]::ReadAllBytes($B)
    if ($x.Length -ne $y.Length) { return $false }
    for ($i = 0; $i -lt $x.Length; $i++) { if ($x[$i] -ne $y[$i]) { return $false } }
    return $true
}

# What -Execute would do: Actions (what to write), Current (what already holds), Refusals.
function Get-InstallPlan([string] $HooksDir, [string] $SourceRoot, [bool] $HooksPathSet) {
    $actions = @()
    $current = @()
    $refusals = @()
    if ($HooksPathSet) { $refusals += 'core.hooksPath is set, so git runs no hook from .git\hooks; unset it, or install these hooks there yourself.' }
    foreach ($name in $HookNames) {
        $path = Join-Path $HooksDir $name
        if ((Test-Path -LiteralPath $path) -and -not (Test-OurHook $path)) {
            $refusals += "$path exists and this script did not write it; it is left alone. Remove or rename it, or merge it with this one by hand."
            continue
        }
        if ((Test-Path -LiteralPath $path) -and ([System.IO.File]::ReadAllText($path) -ceq (Get-HookText $name))) { $current += "hook $name" }
        else { $actions += [pscustomobject]@{ Kind = 'hook'; Name = $name; Path = $path } }
    }
    foreach ($rel in $CheckFiles) {
        $source = Join-Path $SourceRoot $rel
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { $refusals += "$source does not exist: nothing to install."; continue }
        $target = Join-Path (Join-Path $HooksDir $CopyFolder) (Split-Path -Leaf $rel)
        if (Test-SameBytes $source $target) { $current += "copy of $(Split-Path -Leaf $rel)" }
        else { $actions += [pscustomobject]@{ Kind = 'copy'; Name = (Split-Path -Leaf $rel); Path = $target; Source = $source } }
    }
    return [pscustomobject]@{ Actions = @($actions); Current = @($current); Refusals = @($refusals) }
}

function Invoke-InstallPlan($Plan) {
    foreach ($a in $Plan.Actions) {
        $dir = Split-Path -Parent $a.Path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        if ($a.Kind -eq 'hook') {
            [System.IO.File]::WriteAllText($a.Path, (Get-HookText $a.Name), (New-Object System.Text.UTF8Encoding($false)))
        }
        else {
            [System.IO.File]::Copy($a.Source, $a.Path, $true)
        }
    }
}

# What -Uninstall -Execute removes: this script's hooks and its folder. Somebody else's hook stays.
function Get-UninstallPlan([string] $HooksDir) {
    $remove = @()
    $kept = @()
    foreach ($name in $HookNames) {
        $path = Join-Path $HooksDir $name
        if (-not (Test-Path -LiteralPath $path)) { continue }
        if (Test-OurHook $path) { $remove += $path } else { $kept += $path }
    }
    $folder = Join-Path $HooksDir $CopyFolder
    if (Test-Path -LiteralPath $folder) { $remove += $folder }
    return [pscustomobject]@{ Remove = @($remove); Kept = @($kept) }
}

function Invoke-UninstallPlan($Plan) {
    foreach ($p in $Plan.Remove) { Remove-Item -LiteralPath $p -Recurse -Force }
}

# ---------------------------------------------------------------------------------------------
# SELF-TEST
# ---------------------------------------------------------------------------------------------
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()
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

    Write-Host "Tools/Install-GitHooks.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    Write-Host ''
    Write-Host '== the hooks themselves =='
    foreach ($name in $HookNames) {
        $text = Get-HookText $name
        Test-Case "$name starts with a shebang git's sh runs" $true $text.StartsWith("#!/bin/sh`n")
        Test-Case "$name carries the marker" $true $text.Contains($Marker)
        Test-Case "$name has no carriage return, which sh would read as part of a word" $false $text.Contains("`r")
        Test-Case "$name is ASCII" $true (-not ($text -match '[^\x00-\x7F]'))
        Test-Case "$name runs the installed copy of the check as that hook, passing git's arguments" $true $text.Contains("/$CopyFolder/check-personal-data.ps1`" -Hook $name `"`$@`"")
        Test-Case "$name runs Windows PowerShell with no profile and no prompt" $true $text.Contains('exec powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ')
    }

    $root = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-githooks-selftest-' + [guid]::NewGuid().ToString('N'))
    $hooks = Join-Path $root 'hooks'
    $source = Join-Path $root 'repo'
    try {
        New-Item -ItemType Directory -Force -Path $hooks | Out-Null
        New-Item -ItemType Directory -Force -Path (Join-Path $source 'Tools\Checks') | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $source 'Tools\Checks\check-personal-data.ps1'), "'check v1'")
        [System.IO.File]::WriteAllText((Join-Path $source 'Tools\Checks\personal-data.json'), '{}')
        [System.IO.File]::WriteAllText((Join-Path $hooks 'pre-commit.sample'), '#!/bin/sh')

        Write-Host ''
        Write-Host '== install, again, and after the check changes =='
        $plan = Get-InstallPlan $hooks $source $false
        Test-Case 'a fresh install writes three hooks and copies two files' '3 2' "$(@($plan.Actions | Where-Object { $_.Kind -eq 'hook' }).Count) $(@($plan.Actions | Where-Object { $_.Kind -eq 'copy' }).Count)"
        Invoke-InstallPlan $plan
        Test-Case 'the hooks are there' 3 @($HookNames | Where-Object { Test-OurHook (Join-Path $hooks $_) }).Count
        Test-Case 'and the copies' $true ((Test-SameBytes (Join-Path $source 'Tools\Checks\check-personal-data.ps1') (Join-Path $hooks "$CopyFolder\check-personal-data.ps1")) -and (Test-SameBytes (Join-Path $source 'Tools\Checks\personal-data.json') (Join-Path $hooks "$CopyFolder\personal-data.json")))
        $again = Get-InstallPlan $hooks $source $false
        Test-Case 'a second install has nothing to do' '0 5 0' "$($again.Actions.Count) $($again.Current.Count) $($again.Refusals.Count)"
        [System.IO.File]::WriteAllText((Join-Path $source 'Tools\Checks\check-personal-data.ps1'), "'check v2'")
        $changed = Get-InstallPlan $hooks $source $false
        Test-Case 'a changed check is copied again, and only it' 'copy check-personal-data.ps1' (@($changed.Actions | ForEach-Object { "$($_.Kind) $($_.Name)" }) -join ' ')
        Invoke-InstallPlan $changed

        Write-Host ''
        Write-Host '== what it refuses =='
        Test-Case 'core.hooksPath set is refused' 1 @((Get-InstallPlan $hooks $source $true).Refusals).Count
        [System.IO.File]::WriteAllText((Join-Path $hooks 'commit-msg'), "#!/bin/sh`necho somebody else's hook`n")
        $foreign = Get-InstallPlan $hooks $source $false
        Test-Case 'somebody else''s hook is named and left alone' $true (@($foreign.Refusals).Count -eq 1 -and $foreign.Refusals[0].Contains('commit-msg'))
        Test-Case 'and not in the actions' 0 @($foreign.Actions | Where-Object { $_.Name -eq 'commit-msg' }).Count

        Write-Host ''
        Write-Host '== uninstall =='
        $un = Get-UninstallPlan $hooks
        Test-Case 'uninstall removes our two hooks and the copies, keeps the other' '3 1' "$($un.Remove.Count) $($un.Kept.Count)"
        Invoke-UninstallPlan $un
        Test-Case 'what is left: the other hook and the sample' 'commit-msg pre-commit.sample' (@(Get-ChildItem -LiteralPath $hooks | ForEach-Object { $_.Name } | Sort-Object) -join ' ')
    }
    finally {
        if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
    }

    Write-Host ''
    Write-Host '== the files it installs exist in this commit =='
    $repo = Split-Path -Parent $PSScriptRoot
    foreach ($rel in $CheckFiles) { Test-Case "$($rel.Replace('\', '/')) exists" $true (Test-Path -LiteralPath (Join-Path $repo $rel) -PathType Leaf) }

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host 'NOT COVERED HERE: git running the hooks - the build VM has no git. Checked by hand when installed.'
    if ($script:Failures.Count -gt 0) { return 1 }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# ---------------------------------------------------------------------------------------------
# Run.
# ---------------------------------------------------------------------------------------------
Push-Location $RepoRoot
try {
    $common = @(Invoke-NativeCommand { & git rev-parse --path-format=absolute --git-common-dir 2>&1 })
    if ($LASTEXITCODE -ne 0) { Write-Host "NOT A GIT REPOSITORY: $RepoRoot ($($common -join ' '))"; exit 3 }
    $hooksPath = @(Invoke-NativeCommand { & git config --get core.hooksPath 2>&1 })
    $hooksPathSet = ($LASTEXITCODE -eq 0 -and ($hooksPath -join '').Trim().Length -gt 0)
}
finally {
    Pop-Location
}
$hooksDir = Join-Path ([string]$common[0]).Trim() 'hooks'
Write-Host "Hooks folder: $hooksDir (git's common directory: every worktree of this repository uses it)"

if ($Uninstall) {
    $un = Get-UninstallPlan $hooksDir
    foreach ($p in $un.Remove) { Write-Host "  remove $p" }
    foreach ($p in $un.Kept) { Write-Host "  keep   $p (not written by this script)" }
    if ($un.Remove.Count -eq 0) { Write-Host 'Nothing of ours is installed.'; exit 0 }
    if (-not $Execute) { Write-Host 'Dry run: nothing removed. -Uninstall -Execute removes the above.'; exit 0 }
    Invoke-UninstallPlan $un
    Write-Host 'Removed.'
    exit 0
}

$plan = Get-InstallPlan $hooksDir $RepoRoot $hooksPathSet
foreach ($c in $plan.Current) { Write-Host "  current $c" }
foreach ($a in $plan.Actions) { Write-Host "  write   $($a.Path)" }
foreach ($r in $plan.Refusals) { Write-Host "  REFUSED $r" }
if ($plan.Refusals.Count -gt 0) { exit 4 }
if ($plan.Actions.Count -eq 0) { Write-Host 'Installed and current: nothing to do.'; exit 0 }
if (-not $Execute) { Write-Host 'Dry run: nothing written. -Execute writes the above.'; exit 0 }
Invoke-InstallPlan $plan
Write-Host "Installed: $($HookNames -join ', '). The private list is read from .work\personal-data\identifiers.json of the main checkout."
exit 0
