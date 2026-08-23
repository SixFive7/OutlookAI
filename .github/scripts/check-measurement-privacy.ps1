#Requires -Version 5.1
<#
.SYNOPSIS
    Fails when measurement data has reached the repository, or when the measurement gate has
    stopped describing what it gates.

.DESCRIPTION
    The numbers `measurement-gate.ps1` records are statistics about one person's machine. They
    are meaningful only against older values from that same machine, they are not representative
    of anyone else's system, and this repository is public. So the rule is absolute: they live
    under %LOCALAPPDATA%, they never get committed, and they never appear in release notes.

    `measurement-gate.ps1` enforces its half structurally - it refuses to write inside any git
    working tree and refuses to run its comparing modes under CI at all. This script is the
    other half: the one that runs WHERE THE COMMIT LANDS and can say that nothing measurement-
    shaped is tracked. It needs no measurements of its own, touches no store and prints no
    values, so it is safe in a public build log.

    THREE CHECKS.

    1. NOTHING MEASUREMENT-SHAPED IS TRACKED. Every record the gate writes carries a marker
       string, and that string is assembled from two halves in both scripts SO THAT IT NEVER
       APPEARS CONTIGUOUSLY IN ANY REPOSITORY FILE. That is what makes this check exact rather
       than heuristic: any tracked file containing the assembled marker is a real measurement
       record, whatever it was named and wherever it landed, and there is no allowlist to poke
       a hole in. Path shapes are checked too, for a record that was truncated or reformatted
       past the marker.

    2. THE .gitignore STILL COVERS THE SHAPES. An ignore rule that gets deleted is silent until
       the day something lands.

    3. THE GATE STILL DESCRIBES ITSELF. Delegated to `measurement-gate.ps1 -SelfTest`, which
       checks that its catalogue is internally consistent, that every collector still matches
       its source, and that the catalogue and `Docs/measurement-gate.md` describe the same set
       of metrics. A metric added to the gate and left out of the document is a gate nobody can
       read; a metric documented but not gated is a promise the gate does not keep.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER IncludeWorkingTree
    Also scan untracked, non-ignored files. Off by default (CI checks what is committed); worth
    passing locally before a commit, which is the moment the mistake is still cheap.

.EXAMPLE
    pwsh -File .github/scripts/check-measurement-privacy.ps1
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [switch] $IncludeWorkingTree
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# Assembled, never written out whole. See the header: this is what makes check 1 exact.
$marker = 'OUTLOOKAI' + '-MEASUREMENT-RECORD-V1'

$failures = @()
$checks = 0

function Fail([string] $invariant, [string] $detail) {
    $script:failures += "$invariant`n    $detail"
}
function Pass([string] $invariant, [string] $detail) {
    Write-Host "  OK   $invariant - $detail"
}

Write-Host "Checking that no measurement data has reached $RepoRoot"
Write-Host ''

# ---------------------------------------------------------------------------------------------
# 1. Nothing measurement-shaped is tracked.
# ---------------------------------------------------------------------------------------------
$checks++

$files = @()
try {
    Push-Location $RepoRoot
    $files = @(& git ls-files 2>$null)
    if ($LASTEXITCODE -ne 0) { $files = @() }
    if ($IncludeWorkingTree) {
        $extra = @(& git ls-files --others --exclude-standard 2>$null)
        if ($LASTEXITCODE -eq 0) { $files += $extra }
    }
} finally {
    Pop-Location
}

if ($files.Count -eq 0) {
    Fail 'measurement data is not tracked' `
        "git ls-files returned nothing under $RepoRoot. This check cannot prove anything about a file list it could not read, and reporting a clean result from an empty list is exactly the failure mode it exists to avoid."
} else {
    $offenders = @()

    # Path shapes: a store copied into the tree, or a run/comparison file saved next to the code.
    $pathPatterns = @(
        '(^|/)Measurements/',
        '(^|/)history\.jsonl$',
        '(^|/)annotations\.jsonl$',
        '\.comparison\.json$',
        'measurement-history',
        'measurement-run',
        '\.measurements?\.json$'
    )

    foreach ($file in $files) {
        foreach ($pattern in $pathPatterns) {
            if ($file -match $pattern) {
                $offenders += "$file (path looks like a measurement store: /$pattern/)"
                break
            }
        }
    }

    # Content: the decisive check. Text files only - a binary is not a JSON Lines record, and
    # reading every asset in the tree would make this slow for nothing.
    $skipExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.ico', '.exe', '.dll', '.pdb', '.zip', '.snk', '.pfx', '.msi', '.cab')
    $scanned = 0
    foreach ($file in $files) {
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
        if ($skipExtensions -contains $ext) { continue }
        $full = Join-Path $RepoRoot $file
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { continue }
        $scanned++
        $text = Get-Content -LiteralPath $full -Raw -ErrorAction SilentlyContinue
        if ($null -eq $text) { continue }
        if ($text.Contains($marker)) {
            $offenders += "$file (contains a measurement record marker)"
        }
    }

    if ($scanned -eq 0) {
        Fail 'measurement data is not tracked' `
            'scanned no files at all, which cannot be right - this check has switched itself off.'
    } elseif ($offenders.Count -gt 0) {
        Fail 'measurement data is not tracked' `
            ("these tracked files carry measurement data:`n      " + ($offenders -join "`n      ") +
            "`n    Remove them from the index and from history. They are statistics about one machine, this repository is public, and the maintainer's decision is that they are never published. See Docs/measurement-gate.md.")
    } else {
        Pass 'measurement data is not tracked' "$scanned text files scanned, $($files.Count) tracked paths checked"
    }
}

# ---------------------------------------------------------------------------------------------
# 2. The ignore rules that would catch a store landing inside the tree are still there.
# ---------------------------------------------------------------------------------------------
$checks++
$gitignorePath = Join-Path $RepoRoot '.gitignore'
if (-not (Test-Path -LiteralPath $gitignorePath)) {
    Fail 'measurement ignore rules' '.gitignore does not exist.'
} else {
    $gitignore = Get-Content -LiteralPath $gitignorePath -Raw
    $required = @('Measurements/', 'history.jsonl', 'annotations.jsonl', '*.comparison.json')
    $missing = @($required | Where-Object { $gitignore -notmatch [regex]::Escape($_) })
    if ($missing.Count -gt 0) {
        Fail 'measurement ignore rules' `
            ".gitignore no longer covers $($missing -join ', '). Check 1 catches a file that is already tracked; these rules are what stop one becoming tracked in the first place, and a deleted ignore rule is silent until the day something lands."
    } else {
        Pass 'measurement ignore rules' "$($required.Count) shapes still ignored"
    }
}

# ---------------------------------------------------------------------------------------------
# 3. The gate still describes what it gates.
# ---------------------------------------------------------------------------------------------
$checks++
$gate = Join-Path $RepoRoot '.github/scripts/measurement-gate.ps1'
if (-not (Test-Path -LiteralPath $gate)) {
    Fail 'measurement gate self-test' 'measurement-gate.ps1 is missing - the checks above now protect nothing.'
} else {
    $output = & pwsh -NoProfile -File $gate -SelfTest 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        Fail 'measurement gate self-test' ("measurement-gate.ps1 -SelfTest failed:`n      " + ($output.Trim() -replace "`r?`n", "`n      "))
    } else {
        $summary = [regex]::Match($output, 'Measurement gate self-test: (\d+) metrics')
        if (-not $summary.Success) {
            Fail 'measurement gate self-test' 'could not read a metric count out of -SelfTest. Its output shape changed and this check no longer proves anything.'
        } else {
            Pass 'measurement gate self-test' "$($summary.Groups[1].Value) metrics, catalogue and Docs/measurement-gate.md agree"
        }
    }
}

Write-Host ''
if ($failures.Count -gt 0) {
    foreach ($f in $failures) { Write-Host "::error::MEASUREMENT PRIVACY - $f" }
    Write-Host "$($failures.Count) of $checks measurement-privacy checks failed."
    exit 1
}

Write-Host "All $checks measurement-privacy checks hold. No measurement data is in this repository."
exit 0
