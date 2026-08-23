#Requires -Version 5.1
<#
.SYNOPSIS
    Publishes the MCP server and the remediation tools on the HOST and zips them for the guest.

.DESCRIPTION
    THE GUEST HAS A .NET RUNTIME AND NO SDK. Nothing can be compiled there. That single fact is
    why this script exists and why the guest's working directory holds `McpServer.zip` and
    `Tools.zip` rather than a clone.

    Two payloads, because they are used at different times by different things:

      McpServer.zip -> C:\OutlookAI-Q5\server\   the MCP server plus its COM host child. The
                                                 measurement driver talks raw stdio to it.
      Tools.zip     -> C:\OutlookAI-Q5\tools\    OutlookAI.RemediationTools, which is every
                                                 corpus-* verb.

    Framework-dependent, not self-contained: the guest already has the matching runtime
    (10.0.10 as of 2026-08-24) and a self-contained publish would triple the copy for nothing.
    If you change the guest's runtime, check it still satisfies net10.0-windows before assuming
    this still works.

    x64 is not optional. Both projects set PlatformTarget x64, and the index tier reads the
    Search.CollatorDSO OLE DB provider, which has no 32-bit story in this arrangement.

    WHAT THIS DOES NOT SOLVE. Tier-3 tests find the server through
    AssemblyMetadata("McpServerExePath"), baked in at test-build time and pointing inside the
    repository's bin tree. This script stages a DIFFERENT path. The guest cannot run `dotnet
    test` at all today, so the question has not had to be answered - see Testbed/README.md
    question 12.

.PARAMETER OutDir
    Where the two zips land. Default: .work/testbed-payload under the repository root, which is
    gitignored.

.PARAMETER Configuration
    Release by default. A Debug payload is fine for diagnosis but do not measure with one.

.EXAMPLE
    pwsh -File Testbed/host/Publish-GuestPayload.ps1
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName OutlookAI-TestVM -Path .work/testbed-payload/McpServer.zip -Destination C:\OutlookAI-Q5\McpServer.zip
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string] $OutDir,
    [ValidateSet('Release', 'Debug')] [string] $Configuration = 'Release',
    [switch] $SkipBuild
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

if (-not $OutDir) { $OutDir = Join-Path $RepoRoot '.work\testbed-payload' }

$payloads = @(
    @{ Zip = 'McpServer.zip'; Project = 'McpServer\OutlookAI.McpServer\OutlookAI.McpServer.csproj';           Guest = 'C:\OutlookAI-Q5\server' }
    @{ Zip = 'Tools.zip';     Project = 'McpServer\OutlookAI.RemediationTools\OutlookAI.RemediationTools.csproj'; Guest = 'C:\OutlookAI-Q5\tools'  }
)

$staging = Join-Path $OutDir 'staging'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

foreach ($p in $payloads) {
    $proj = Join-Path $RepoRoot $p.Project
    if (-not (Test-Path -LiteralPath $proj)) { throw "Project not found: $proj" }

    $dest = Join-Path $staging ([IO.Path]::GetFileNameWithoutExtension($p.Zip))
    $zip = Join-Path $OutDir $p.Zip

    if (-not $SkipBuild) {
        if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Recurse -Force }
        Write-Host "Publishing $($p.Project) -> $dest"
        # Output is captured rather than streamed: a publish that spawns its own children can
        # otherwise hold the pipe open long after it has finished.
        $log = Join-Path $OutDir ("publish-" + [IO.Path]::GetFileNameWithoutExtension($p.Zip) + ".log")
        & dotnet publish $proj -c $Configuration -f net10.0-windows --self-contained false -o $dest *> $log
        if ($LASTEXITCODE -ne 0) {
            Write-Host (Get-Content -LiteralPath $log -Tail 40 | Out-String)
            throw "dotnet publish failed for $($p.Project) (exit $LASTEXITCODE). Full log: $log"
        }
    }

    if (-not (Test-Path -LiteralPath $dest)) { throw "Nothing staged at $dest (did you pass -SkipBuild without a previous run?)" }

    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    Compress-Archive -Path (Join-Path $dest '*') -DestinationPath $zip
    $size = [math]::Round((Get-Item -LiteralPath $zip).Length / 1MB, 1)
    Write-Host "  $($p.Zip)  $size MB  -> expand on the guest into $($p.Guest)"
}

Write-Host ''
Write-Host "Payloads in $OutDir"
Write-Host 'On the guest, in an ELEVATED shell (PowerShell Direct is fine - no COM here):'
Write-Host '  Expand-Archive C:\OutlookAI-Q5\McpServer.zip -DestinationPath C:\OutlookAI-Q5\server -Force'
Write-Host '  Expand-Archive C:\OutlookAI-Q5\Tools.zip     -DestinationPath C:\OutlookAI-Q5\tools  -Force'
Write-Host 'Anything that then TOUCHES OUTLOOK must go through Testbed/guest/Register-InteractiveTask.ps1.'
