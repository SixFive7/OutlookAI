#Requires -Version 5.1
<#
.SYNOPSIS
    Fails when a PowerShell script in this repository would break under Windows PowerShell 5.1 -
    the PowerShell that ships with Windows, and the only one the test guests have.

.DESCRIPTION
    DECIDED 2026-09-24 (Q78): every script here runs on the Windows PowerShell 5.1 that ships with
    Windows as well as on PowerShell 7, so that nothing has to be installed on a guest. Until then
    CI and the documentation ran everything under pwsh 7, and three different 5.1-only failures
    reached the tree unnoticed - each one passes every pwsh run and fails the moment 5.1 runs it.
    This script fails the build on all three, statically, in every .ps1, .psm1 and .psd1 that git
    tracks, plus any untracked one git does not ignore, so a new script is checked in the commit
    that adds it.

    1. $PSScriptRoot, $PSCommandPath OR $MyInvocation IN THE SCRIPT'S param() BLOCK. Windows
       PowerShell 5.1 leaves all three EMPTY while an advanced script's param() defaults are
       evaluated - under -File, and when it is dot-sourced - so a default built from one throws
       before the script runs, or quietly points somewhere else. Measured 2026-09-24:
       check-pinned-constants.ps1 died under powershell.exe -File on
       ParameterArgumentValidationErrorEmptyStringNotAllowed. A plain script happens to get them,
       but one [CmdletBinding()] or [Parameter()] makes it advanced, so the pattern is refused
       outright. The fix: leave the parameter bare and set the default in the body, behind
       $PSBoundParameters.ContainsKey, as every guard here does.

    2. NON-ASCII BYTES IN A FILE WITH NO BYTE ORDER MARK. Windows PowerShell 5.1 reads a script
       that has no BOM in the machine's ANSI code page, so UTF-8 text in it arrives mangled -
       harmless in a comment, a wrong value in a string, and a parse error when a mangled byte
       happens to be one of the quotes or dashes PowerShell treats as syntax. PowerShell 7 assumes
       UTF-8 and never shows the problem. Keep scripts ASCII, or save them as UTF-8 WITH a BOM.

    3. A NATIVE PROGRAM'S STDERR REDIRECTED UNDER 'Stop'. With $ErrorActionPreference = 'Stop',
       Windows PowerShell 5.1 turns the first line a native program writes to a redirected stderr
       - 2>$null, 2>&1 and *> alike - into a terminating NativeCommandError, so one warning or
       progress line ends the script before its exit code can be read. oscdimg died that way on its
       own "0% complete". PowerShell 7 does not. Measured 2026-09-24.
       The redirection does not have to sit on the program itself: one on a function, a script
       block or a script that runs it reaches it just the same (measured), so this check also
       follows a redirected call into the functions of the same file and the script blocks it runs.
       In a file that sets 'Stop', such a call is accepted only when the preference is not 'Stop'
       where it runs - it sits in a script block handed to a function of the same file that sets
       $ErrorActionPreference to 'Continue' (the Invoke-NativeCommand restated in the scripts that
       need one is the model), or an assignment of 'Continue' precedes it in its own function or
       script block - AND a try encloses it. The try is not decoration: under 'Continue' without
       one, "the program was not found at all" is demoted to a printed message and the caller goes
       on to read a stale $LASTEXITCODE. Measured in both shells. Anything else needs the marker
       on its line or the line above, with the reason:
           # ps51-native-stderr-ok: <why this call cannot die under 5.1>

    AND A SCRIPT THAT DOES NOT PARSE FAILS. Under 5.1 that is what PowerShell 7-only syntax - a
    ternary, ??, && or || between pipelines - looks like, which is one reason CI runs this script
    under both shells.

    WHAT IT CANNOT SEE, stated rather than implied. A redirection applied by whatever RUNS a
    script reaches every native call inside it, redirected or not, and so does a host that
    captures stderr without being asked. Both happen to the guest scripts: the wrapper
    Testbed/guest/Register-InteractiveTask.ps1 puts around its work ends in *>&1, and a
    remoting host captures on its own. Measured 2026-09-24 under Windows PowerShell 5.1 with
    'Stop': an UNREDIRECTED native call died inside that wrapper's shape, and inside a Start-Job
    job, whose host is the ServerRemoteHost that PowerShell Direct also uses (PowerShell Direct
    itself was not measured). That is a property of how a script is run, not of its text, so no
    static check of the script can hold it: a guest script that sets 'Stop' is safe there only if
    every native call it makes goes through Invoke-NativeCommand. Nor does this check read a
    script that is dot-sourced into one that sets 'Stop' as if it set 'Stop' itself.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.EXAMPLE
    pwsh -File .github/scripts/check-powershell-51.ps1
    powershell -NoProfile -File .github/scripts/check-powershell-51.ps1
#>
[CmdletBinding()]
param(
    [string] $RepoRoot
)
# Defaults that need $PSScriptRoot are set HERE, not in param() - check 1 below says why.
if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) }

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$script:Failures = @()
$script:Checks = 0

function Fail([string] $invariant, [string] $detail) {
    $script:Failures += "$invariant`n    $detail"
}
function Pass([string] $invariant, [string] $detail) {
    Write-Host "  OK   $invariant - $detail"
}

# A NATIVE PROGRAM WHOSE STDERR IS REDIRECTED RUNS THROUGH HERE (Q78). Under
# $ErrorActionPreference = 'Stop', Windows PowerShell 5.1 turns the first line a native program
# writes to a redirected stderr - 2>$null, 2>&1 and *> alike - into a terminating
# NativeCommandError, so one warning or progress line ends the script before its exit code can be
# read. PowerShell 7 does not. Measured on the host 2026-09-24. So, here and only here:
#   * 'Continue' holds in THIS function's scope. The caller's 'Stop' is never changed, so there is
#     nothing to restore and nothing else is relaxed.
#   * The try is load-bearing. Without one, 'Continue' also demotes a terminating error inside the
#     block - the program not being found at all - to a printed message, and the caller goes on to
#     read a stale $LASTEXITCODE. Inside a try it stops the caller exactly as it always did.
#     Measured in both shells.
#   * Stderr lines come back as plain strings in both shells, never as ErrorRecords: 5.1 renders
#     those with a position block around every line, and an empty one as an exception type name.
#   * The exit code is left in $LASTEXITCODE, and the caller checks it.
# Restated in each script that needs it, as this repository restates its shared rules.
# .github/scripts/check-powershell-51.ps1 fails the build on a redirected native call that does
# not go through a function like this one.
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

# ---------------------------------------------------------------------------------------------
# The syntax tree, in terms this script needs.
# ---------------------------------------------------------------------------------------------
$script:AllowMarker = '#\s*ps51-native-stderr-ok:\s*\S.{9,}'

# Aliases PowerShell resolves before it looks for any program, identical in 5.1 and 7 (read from
# both shells 2026-09-24). Deliberately not curl, wget or sc: those are cmdlets in 5.1 and
# programs in 7.
$script:BuiltInAliases = @('%', '?', 'foreach', 'where', 'select', 'sort', 'echo', 'measure',
    'group', 'tee', 'write', 'cat', 'type', 'ls', 'dir', 'diff')

function Get-EnclosingScriptBlock($node) {
    $p = $node.Parent
    while ($null -ne $p -and -not ($p -is [System.Management.Automation.Language.ScriptBlockAst])) { $p = $p.Parent }
    return $p
}

function Test-InsideTryBody($node) {
    $child = $node
    $p = $node.Parent
    while ($null -ne $p) {
        if ($p -is [System.Management.Automation.Language.TryStatementAst] -and [object]::ReferenceEquals($p.Body, $child)) { return $true }
        $child = $p
        $p = $p.Parent
    }
    return $false
}

function Get-ErrorRedirection($command) {
    foreach ($r in $command.Redirections) {
        if ($r.FromStream -eq [System.Management.Automation.Language.RedirectionStream]::Error -or
            $r.FromStream -eq [System.Management.Automation.Language.RedirectionStream]::All) { return $r }
    }
    return $null
}

# 'native' for a program or anything this file cannot see into (a script path, a variable, an
# expression); 'function' for a function defined in this file; 'scriptblock' for & { ... };
# 'powershell' for a cmdlet or module function, recognised by its Verb-Noun name or as one of the
# aliases above.
function Get-CommandKind($command, $context) {
    if ($command.CommandElements[0] -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) { return 'scriptblock' }
    $name = $command.GetCommandName()
    if (-not $name) { return 'native' }
    if ($context.Functions.ContainsKey($name.ToLowerInvariant())) { return 'function' }
    if ($script:BuiltInAliases -contains $name) { return 'powershell' }
    if ($name -match '^[A-Za-z]+-[A-Za-z0-9]+$') { return 'powershell' }
    return 'native'
}

# A function of this file that sets 'Continue' in its own scope: what a script block handed to it
# runs under. HasTry says whether it also holds the try that keeps a missing program loud.
function Get-RelaxingFunction([string] $name, $context) {
    $key = $name.ToLowerInvariant()
    if (-not $context.Functions.ContainsKey($key)) { return $null }
    foreach ($f in $context.Functions[$key]) {
        $relaxes = @($context.Assignments | Where-Object { $_.Relaxes -and [object]::ReferenceEquals($_.Scope, $f.Body) }).Count -gt 0
        if ($relaxes) {
            $hasTry = $null -ne $f.Body.Find({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] }, $true)
            return [pscustomobject]@{ Name = $f.Name; HasTry = $hasTry }
        }
    }
    return $null
}

# What the preference is where $command runs, as far as the text can say. Walks outward: in each
# function or script block, the nearest assignment BEFORE the call decides; a script block handed
# to a relaxing function of this file runs under that function's 'Continue'. A function boundary
# ends the walk - a function runs whenever it is called, which in a file that sets 'Stop' is after
# it did - and so does the top of the file. $null means 'Stop' governs the call.
function Get-Relaxation($command, $context) {
    $node = $command
    while ($null -ne $node) {
        if ($node -is [System.Management.Automation.Language.ScriptBlockAst]) {
            $before = @($context.Assignments | Where-Object {
                    [object]::ReferenceEquals($_.Scope, $node) -and $_.Ast.Extent.EndOffset -le $command.Extent.StartOffset
                } | Sort-Object { $_.Ast.Extent.StartOffset })
            if ($before.Count -gt 0) {
                $last = $before[$before.Count - 1]
                if ($last.Relaxes) {
                    return [pscustomobject]@{
                        Kind = "an assignment of 'Continue' in its own scope"
                        By   = "the assignment at line $($last.Ast.Extent.StartLineNumber)"
                        Loud = (Test-InsideTryBody $command)
                    }
                }
                return $null
            }
            $holder = $node.Parent
            if ($holder -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -and
                $holder.Parent -is [System.Management.Automation.Language.CommandAst] -and
                -not [object]::ReferenceEquals($holder.Parent.CommandElements[0], $holder)) {
                $caller = $holder.Parent
                $callee = $caller.GetCommandName()
                if ($callee) {
                    $relaxing = Get-RelaxingFunction $callee $context
                    if ($null -ne $relaxing) {
                        return [pscustomobject]@{
                            Kind = $relaxing.Name
                            By   = "$($relaxing.Name), line $($caller.Extent.StartLineNumber)"
                            Loud = ($relaxing.HasTry -or (Test-InsideTryBody $caller))
                        }
                    }
                }
            }
            if ($node.Parent -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $null }
        }
        $node = $node.Parent
    }
    return $null
}

# The native calls a redirection on $command reaches without being written on them: inside the
# script block it runs, the script blocks handed to it, and the functions of this file it calls,
# transitively. Calls that carry their own redirection are judged on their own and skipped here.
function Get-ReachedNativeCall($command, $context, $visited) {
    $reached = New-Object System.Collections.ArrayList
    $bodies = New-Object System.Collections.ArrayList
    $kind = Get-CommandKind $command $context
    if ($kind -eq 'scriptblock') { $null = $bodies.Add($command.CommandElements[0].ScriptBlock) }
    if ($kind -eq 'function') {
        foreach ($f in $context.Functions[$command.GetCommandName().ToLowerInvariant()]) {
            if ($visited.Add($f)) { $null = $bodies.Add($f.Body) }
        }
    }
    for ($i = 1; $i -lt $command.CommandElements.Count; $i++) {
        $element = $command.CommandElements[$i]
        if ($element -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) { $null = $bodies.Add($element.ScriptBlock) }
    }
    foreach ($body in $bodies) {
        foreach ($inner in $body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
            $innerKind = Get-CommandKind $inner $context
            if ($innerKind -eq 'native') {
                if ($null -ne (Get-ErrorRedirection $inner)) { continue }
                if ($null -eq (Get-Relaxation $inner $context)) { $null = $reached.Add($inner) }
            }
            elseif ($innerKind -eq 'function') {
                foreach ($r in (Get-ReachedNativeCall $inner $context $visited)) { $null = $reached.Add($r) }
            }
        }
    }
    return , $reached
}

function Format-Command($command) {
    $text = ($command.Extent.Text -replace '\s+', ' ').Trim()
    if ($text.Length -gt 110) { $text = $text.Substring(0, 107) + '...' }
    return $text
}

function Test-HasMarker($command, [string[]] $lines) {
    $line = $command.Extent.StartLineNumber
    foreach ($n in @($line, ($line - 1))) {
        if ($n -ge 1 -and $n -le $lines.Count -and $lines[$n - 1] -match $script:AllowMarker) { return $true }
    }
    return $false
}

Write-Host "Checking that every PowerShell script under $RepoRoot runs on Windows PowerShell 5.1"
Write-Host "(this pass: $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion))"
Write-Host ''

# ---------------------------------------------------------------------------------------------
# The file list. Tracked, plus untracked-and-not-ignored: a script added in the same commit as
# the mistake should be caught at the moment it is still cheap.
# ---------------------------------------------------------------------------------------------
Push-Location $RepoRoot
try {
    $listed = @(Invoke-NativeCommand { & git ls-files --cached --others --exclude-standard -- '*.ps1' '*.psm1' '*.psd1' 2>$null })
    $listExit = $LASTEXITCODE
}
finally {
    Pop-Location
}
$scripts = @($listed | Sort-Object -Unique | Where-Object { Test-Path -LiteralPath (Join-Path $RepoRoot $_) -PathType Leaf })
if ($listExit -ne 0 -or $scripts.Count -eq 0) {
    Fail 'script list' "git ls-files returned no PowerShell scripts under $RepoRoot (exit $listExit). Every check below would pass on an empty list, so this fails instead."
}

$parseProblems = @()
$paramProblems = @()
$encodingProblems = @()
$stderrProblems = @()
$parsed = 0
$paramBlocks = 0
$withBom = 0
$stopFiles = 0
$redirectedCalls = 0
$acceptedBy = @{}
$marked = 0

foreach ($relative in $scripts) {
    $full = Join-Path $RepoRoot $relative

    # --- 2. the bytes, before anything reads them as text ----------------------------------
    $bytes = [System.IO.File]::ReadAllBytes($full)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) -or
        ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)))
    if ($hasBom) {
        $withBom++
    }
    else {
        $line = 1
        $column = 1
        $first = $null
        $count = 0
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            $b = $bytes[$i]
            if ($b -gt 0x7F) {
                $count++
                if ($null -eq $first) { $first = 'line {0}, column {1}, byte 0x{2:X2}' -f $line, $column, $b }
            }
            if ($b -eq 0x0A) { $line++; $column = 1 } else { $column++ }
        }
        if ($count -gt 0) {
            $encodingProblems += "$relative - $count non-ASCII byte(s) and no BOM; the first at $first"
        }
    }

    # --- 0. it parses, under THIS PowerShell -----------------------------------------------
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($full, [ref] $tokens, [ref] $errors)
    if ($errors.Count -gt 0) {
        $shown = @($errors | Select-Object -First 3 | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Message)" })
        $parseProblems += "$relative - $($errors.Count) parse error(s): $($shown -join ' | ')"
        continue
    }
    $parsed++
    $lines = @(Get-Content -LiteralPath $full)

    # --- 1. the param() block ----------------------------------------------------------------
    if ($null -ne $ast.ParamBlock) {
        $paramBlocks++
        $uses = $ast.ParamBlock.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $n.VariablePath.UserPath -match '^(?:(?:global|script|local|private):)?(?:PSScriptRoot|PSCommandPath|MyInvocation)$'
            }, $true)
        foreach ($u in $uses) {
            $paramProblems += "${relative}:$($u.Extent.StartLineNumber) - $($u.Extent.Text) in the script's param() block"
        }
    }

    # --- 3. redirected stderr under 'Stop' ---------------------------------------------------
    $functions = @{}
    foreach ($f in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        $key = $f.Name.ToLowerInvariant()
        if (-not $functions.ContainsKey($key)) { $functions[$key] = New-Object System.Collections.ArrayList }
        $null = $functions[$key].Add($f)
    }
    $assignments = @(
        foreach ($a in $ast.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                    $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                    $n.Left.VariablePath.UserPath -match '^(?:(?:global|script|local|private):)?ErrorActionPreference$'
                }, $true)) {
            $value = ''
            if ($a.Right.Extent.Text -match '^\s*[''"]?(Stop|Continue|SilentlyContinue|Ignore|Inquire)[''"]?\s*$') { $value = $Matches[1] }
            elseif ($a.Right.Extent.Text -match '::\s*(Stop|Continue|SilentlyContinue|Ignore|Inquire)\s*$') { $value = $Matches[1] }
            # global: and private: do not reach the script block a helper runs, so neither relaxes.
            $scoped = $a.Left.VariablePath.UserPath -match '^(?:global|private):'
            [pscustomobject]@{
                Ast     = $a
                Scope   = (Get-EnclosingScriptBlock $a)
                IsStop  = ($value -eq 'Stop')
                Relaxes = (($value -eq 'Continue' -or $value -eq 'SilentlyContinue') -and -not $scoped)
            }
        })
    if (@($assignments | Where-Object { $_.IsStop }).Count -eq 0) { continue }
    $stopFiles++
    $context = [pscustomobject]@{ Functions = $functions; Assignments = $assignments }

    foreach ($command in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        if ($null -eq (Get-ErrorRedirection $command)) { continue }
        $kind = Get-CommandKind $command $context
        $where = "${relative}:$($command.Extent.StartLineNumber)"

        if ($kind -eq 'native') {
            $redirectedCalls++
            if (Test-HasMarker $command $lines) { $marked++; continue }
            $relaxation = Get-Relaxation $command $context
            if ($null -eq $relaxation) {
                $stderrProblems += "$where - $(Format-Command $command)`n        redirects a native program's stderr under 'Stop'. Run it through Invoke-NativeCommand { ... } - see the one in Testbed/host/Publish-GuestPayload.ps1 - or mark it with the reason it cannot die."
            }
            elseif (-not $relaxation.Loud) {
                $stderrProblems += "$where - $(Format-Command $command)`n        runs under 'Continue' ($($relaxation.By)) with no try around it, so a missing program would print a message and the caller would read a stale `$LASTEXITCODE. Put the try in."
            }
            else {
                $acceptedBy[$relaxation.Kind] = 1 + [int] $acceptedBy[$relaxation.Kind]
            }
            continue
        }

        # A redirected function, script block or cmdlet: what native calls does it reach?
        if (Test-HasMarker $command $lines) { $marked++; continue }
        $visited = New-Object 'System.Collections.Generic.HashSet[object]'
        foreach ($inner in (Get-ReachedNativeCall $command $context $visited)) {
            $redirectedCalls++
            $stderrProblems += "$where - $(Format-Command $command)`n        redirects stderr around a native call it runs at line $($inner.Extent.StartLineNumber), $(Format-Command $inner), and the redirection reaches it. Run that call through Invoke-NativeCommand, or drop the redirection."
        }
    }
}

# ---------------------------------------------------------------------------------------------
# Verdicts. Parse failures first: a script that does not parse was not checked for the rest.
# ---------------------------------------------------------------------------------------------
$script:Checks++
if ($parseProblems.Count -gt 0) {
    Fail "every script parses under PowerShell $($PSVersionTable.PSVersion)" (($parseProblems -join "`n    ") +
        "`n    A script this PowerShell cannot parse cannot run on it either. Under 5.1 this is what PowerShell 7-only syntax looks like: a ternary, ??, && or || between pipelines.")
}
else {
    Pass "every script parses under PowerShell $($PSVersionTable.PSVersion)" "$parsed script(s)"
}

$script:Checks++
if ($paramProblems.Count -gt 0) {
    Fail 'no $PSScriptRoot, $PSCommandPath or $MyInvocation in a param() block' (($paramProblems -join "`n    ") +
        "`n    Windows PowerShell 5.1 leaves these EMPTY while an advanced script's param() defaults are evaluated, under -File and when dot-sourced. Leave the parameter bare and set the default in the body, behind `$PSBoundParameters.ContainsKey('<name>').")
}
else {
    Pass 'no $PSScriptRoot, $PSCommandPath or $MyInvocation in a param() block' "$paramBlocks param() block(s) in $parsed script(s)"
}

$script:Checks++
if ($encodingProblems.Count -gt 0) {
    Fail 'no non-ASCII byte in a script without a BOM' (($encodingProblems -join "`n    ") +
        "`n    Windows PowerShell 5.1 reads a BOM-less script in the ANSI code page, so these bytes do not mean there what they mean in an editor. Replace them with ASCII, or save the file as UTF-8 WITH a BOM.")
}
else {
    Pass 'no non-ASCII byte in a script without a BOM' "$($scripts.Count) script(s) read, $withBom with a BOM"
}

$script:Checks++
if ($stderrProblems.Count -gt 0) {
    Fail "no native program's stderr redirected under 'Stop'" (($stderrProblems -join "`n    ") +
        "`n    Windows PowerShell 5.1 turns the first line such a program writes to stderr into a terminating NativeCommandError when `$ErrorActionPreference is 'Stop'.")
}
else {
    $how = @($acceptedBy.Keys | Sort-Object | ForEach-Object { "$($acceptedBy[$_]) through $_" })
    if ($marked -gt 0) { $how += "$marked marked" }
    $howText = ''
    if ($how.Count -gt 0) { $howText = ': ' + ($how -join ', ') }
    Pass "no native program's stderr redirected under 'Stop'" "$redirectedCalls redirected native call(s) in $stopFiles script(s) that set 'Stop'$howText"
}

Write-Host ''
if ($script:Failures.Count -gt 0) {
    foreach ($f in $script:Failures) { Write-Host "::error::POWERSHELL 5.1 - $f" }
    Write-Host "$($script:Failures.Count) of $($script:Checks) PowerShell 5.1 checks failed."
    exit 1
}

Write-Host "All $($script:Checks) PowerShell 5.1 checks hold across $($scripts.Count) script(s)."
exit 0
