#Requires -Version 5.1
<#
.SYNOPSIS
    Fails when two files that must agree about a value have stopped agreeing.

.DESCRIPTION
    Some values in this repository genuinely exist twice, because the two sides are written in
    languages that cannot see each other: C# and MSBuild XML, C# and Inno Setup's Pascal, C#
    and a GitHub Actions PowerShell step. A comment saying "keep these in step" is not a
    mechanism - the audit that produced Docs/magic-numbers.md found one such comment that had
    already become false. This script is the mechanism.

    It is deliberately text-based. It cannot compile the add-in (net48/VSTO) and it cannot run
    Inno Setup, so it reads the sources and compares what it finds. Every check therefore also
    asserts that it FOUND both sides: a regex that silently stops matching would otherwise turn
    into a check that always passes, which is worse than no check at all.

    Run it from anywhere:
        pwsh -File .github/scripts/check-pinned-constants.ps1

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER ExpectedSigningThumbprint
    Optional. When given (the release workflow passes the certificate it has just imported),
    both pinned copies of the thumbprint must also match this.
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string] $ExpectedSigningThumbprint
)

$ErrorActionPreference = 'Stop'

$script:Failures = @()
$script:Checks = 0

function Fail([string] $invariant, [string] $detail) {
    $script:Failures += "$invariant`n    $detail"
}

function Pass([string] $invariant, [string] $detail) {
    Write-Host "  OK   $invariant - $detail"
}

function Read-Source([string] $relativePath) {
    $full = Join-Path $RepoRoot $relativePath
    if (-not (Test-Path $full)) {
        Fail "file present" "$relativePath does not exist (did it move? this check now proves nothing)"
        return $null
    }
    return Get-Content -LiteralPath $full -Raw
}

# One capture group, or a failure naming the file. A check whose regex stopped matching is a
# check that has quietly switched itself off, so this never returns "nothing found" silently.
function Get-Pinned([string] $relativePath, [string] $pattern, [string] $what) {
    $text = Read-Source $relativePath
    if ($null -eq $text) { return $null }
    $m = [regex]::Match($text, $pattern)
    if (-not $m.Success) {
        Fail $what "could not find it in $relativePath - the file changed shape and this check no longer proves anything. Pattern: $pattern"
        return $null
    }
    return $m.Groups[1].Value
}

function Get-PinnedAll([string] $relativePath, [string] $pattern, [string] $what) {
    $text = Read-Source $relativePath
    if ($null -eq $text) { return $null }
    # Not $Matches: that is a PowerShell automatic variable and writing to it is asking for
    # a surprise in whatever runs next.
    $found = [regex]::Matches($text, $pattern)
    if ($found.Count -eq 0) {
        Fail $what "found no matches in $relativePath - the file changed shape and this check no longer proves anything. Pattern: $pattern"
        return $null
    }
    return @($found | ForEach-Object { $_.Groups[1].Value })
}

Write-Host "Checking cross-file invariants under $RepoRoot"
Write-Host ""

# ---------------------------------------------------------------------------------------------
# 1. Installer signing certificate.
#    UpdateService.ExpectedCertThumbprint == csproj ManifestCertificateThumbprint.
#    The csproj half fails loudly at build time on a rotation; the C# half fails CLOSED and
#    SILENTLY - every future installer is refused as "not signed by the expected OutlookAI
#    certificate" and auto-update dies across the installed base.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$csThumb = Get-Pinned 'Services/UpdateService.cs' `
    'ExpectedCertThumbprint\s*=\s*"([0-9A-Fa-f]{40})"' 'signing thumbprint (UpdateService.cs)'
$projThumb = Get-Pinned 'OutlookAI.csproj' `
    '<ManifestCertificateThumbprint>\s*([0-9A-Fa-f]{40})\s*</ManifestCertificateThumbprint>' 'signing thumbprint (OutlookAI.csproj)'
if ($csThumb -and $projThumb) {
    if ($csThumb -ine $projThumb) {
        Fail "signing thumbprint pin" "UpdateService.ExpectedCertThumbprint is $csThumb but OutlookAI.csproj ManifestCertificateThumbprint is $projThumb. Rotating the signing certificate must change BOTH - the updater's copy fails closed and silently, so every installed copy would stop auto-updating with no error anyone sees."
    } else {
        Pass "signing thumbprint pin" $csThumb
    }
}
if ($ExpectedSigningThumbprint) {
    $script:Checks++
    if ($csThumb -and ($csThumb -ine $ExpectedSigningThumbprint)) {
        Fail "signing thumbprint matches the certificate in use" "the certificate being signed with is $ExpectedSigningThumbprint, but the shipped updater pins $csThumb. Every installer produced from this build would be rejected by every installed copy."
    } elseif ($csThumb) {
        Pass "signing thumbprint matches the certificate in use" $ExpectedSigningThumbprint
    }
}

# ---------------------------------------------------------------------------------------------
# 2. Installer mutex.
#    ThisAddIn.InstallerMutexName == Installer.iss SetupMutex. Rename one side and the add-in
#    initialises during a silent auto-update, with the installer tearing its processes down
#    mid-flight - no error, just the failure the mutex exists to prevent.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$csMutex = Get-Pinned 'ThisAddIn.cs' `
    'InstallerMutexName\s*=\s*"([^"]+)"' 'installer mutex (ThisAddIn.cs)'
$issMutex = Get-Pinned 'Installer.iss' `
    '(?m)^\s*SetupMutex\s*=\s*(\S+)\s*$' 'installer mutex (Installer.iss)'
if ($csMutex -and $issMutex) {
    if ($csMutex -cne $issMutex) {
        Fail "installer mutex name" "ThisAddIn.InstallerMutexName is '$csMutex' but Installer.iss SetupMutex is '$issMutex'. The add-in would no longer detect a running installer."
    } else {
        Pass "installer mutex name" $csMutex
    }
}

# ---------------------------------------------------------------------------------------------
# 3. Auto-updater download cap.
#    UpdateService.MaxDownloadBytes == the release workflow's installer-size gate. Shipping an
#    asset over the cap silently stops auto-update everywhere; the gate is what turns that into
#    a failed release instead.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$capMb = Get-Pinned 'Services/UpdateService.cs' `
    'MaxDownloadBytes\s*=\s*(\d+)L?\s*\*\s*1024\s*\*\s*1024' 'download cap (UpdateService.cs)'
$gateMb = Get-Pinned '.github/workflows/release.yml' `
    '\$exe\.Length\s+-gt\s+(\d+)MB' 'download cap (release.yml gate)'
if ($capMb -and $gateMb) {
    if ([int]$capMb -ne [int]$gateMb) {
        Fail "installer size cap" "UpdateService.MaxDownloadBytes is ${capMb} MB but release.yml refuses installers over ${gateMb} MB. The release gate must refuse exactly what the updater refuses."
    } else {
        Pass "installer size cap" "$capMb MB"
    }
}

# ---------------------------------------------------------------------------------------------
# 4. Office versions.
#    OfficeVersions.Supported == the Office majors Installer.iss writes resiliency exemptions
#    for. A version the installer exempts but the add-in never probes (or the reverse) is a
#    machine where half the product's Office integration silently does nothing.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$supportedRaw = Get-Pinned 'Services/OfficeVersions.cs' `
    'Supported\s*=\s*\{([^}]*)\}' 'Office versions (OfficeVersions.cs)'
$issVersions = Get-PinnedAll 'Installer.iss' `
    'Software\\Microsoft\\Office\\([0-9]+\.[0-9]+)\\Outlook\\Resiliency\\DoNotDisableAddinList' 'Office versions (Installer.iss)'
if ($supportedRaw -and $issVersions) {
    $csVersions = @([regex]::Matches($supportedRaw, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    $left = ($csVersions | Sort-Object) -join ','
    $right = (@($issVersions) | Sort-Object -Unique) -join ','
    if ($left -ne $right) {
        Fail "supported Office versions" "OfficeVersions.Supported is {$left} but Installer.iss writes resiliency exemptions for {$right}. Every Office major the add-in supports needs the exemption, and exempting one the add-in never looks at is a claim the product does not honour."
    } else {
        Pass "supported Office versions" $left
    }
}

# ---------------------------------------------------------------------------------------------
# 5. .NET runtime download page.
#    McpRegistrationService.DotnetRuntimeDownloadUrl == Installer.iss NetRuntime10ManualUrl.
#    Both send a user with no runtime to the same page, one from setup and one from Settings.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$csUrl = Get-Pinned 'Services/McpRegistrationService.cs' `
    'DotnetRuntimeDownloadUrl\s*=\s*"([^"]+)"' 'runtime download URL (McpRegistrationService.cs)'
$issUrl = Get-Pinned 'Installer.iss' `
    '(?m)^\s*#define\s+NetRuntime10ManualUrl\s+"([^"]+)"' 'runtime download URL (Installer.iss)'
if ($csUrl -and $issUrl) {
    if ($csUrl -cne $issUrl) {
        Fail "runtime download URL" "McpRegistrationService.DotnetRuntimeDownloadUrl is '$csUrl' but Installer.iss NetRuntime10ManualUrl is '$issUrl'. One of the two places that tell a user where to get the runtime is now pointing somewhere else."
    } else {
        Pass "runtime download URL" $csUrl
    }
}

Write-Host ""
if ($script:Failures.Count -gt 0) {
    foreach ($f in $script:Failures) {
        Write-Host "::error::PINNED CONSTANT DRIFT - $f"
    }
    Write-Error "$($script:Failures.Count) of $($script:Checks) cross-file invariants failed. See Docs/magic-numbers.md for what each one protects."
    exit 1
}

Write-Host "All $($script:Checks) cross-file invariants hold."
exit 0
