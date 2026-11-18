#Requires -Version 5.1
<#
.SYNOPSIS
    Fails when two files that must agree about a value have stopped agreeing.

.DESCRIPTION
    Some values in this repository genuinely exist twice, because the two sides are written in
    languages that cannot see each other: C# and MSBuild XML, C# and Inno Setup's Pascal, C#
    and the release script's PowerShell. A comment saying "keep these in step" is not a
    mechanism - the audit that produced Docs/magic-numbers.md found one such comment that had
    already become false. This script is the mechanism.

    Two checks (#12, #13) are the opposite shape and belong here for the same reason: they hold a
    value that USED to exist twice down to one copy. A mirror closed by sharing a file stays
    closed only while nobody types the literal again, and re-typing it compiles - so "this exists
    exactly once, here" is as much a cross-file invariant as "these two agree". #14 is that shape
    for a build rule: the guard in OutlookAI.csproj that keeps a command-line build from
    registering the add-in, which no compilation can see and whose deletion breaks nothing. #15 is
    a boundary of the same kind: the one part of OutlookAI.Core the add-in may use, and what that
    part may not do inside Outlook's process (Q164).

    It is deliberately text-based. It cannot compile the add-in (net48/VSTO) and it cannot run
    Inno Setup, so it reads the sources and compares what it finds. Every check therefore also
    asserts that it FOUND both sides: a regex that silently stops matching would otherwise turn
    into a check that always passes, which is worse than no check at all.

    Run it from anywhere:
        pwsh -File Tools/Checks/check-pinned-constants.ps1

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER ExpectedSigningThumbprint
    Optional. When given (Tools/Publish-Release.ps1 passes the certificate it is about to sign
    the release with), both pinned copies of the thumbprint must also match this.
#>
[CmdletBinding()]
param(
    [string] $RepoRoot,
    [string] $ExpectedSigningThumbprint
)
# Defaults that need $PSScriptRoot are set HERE, not in param(). Windows PowerShell 5.1 leaves
# $PSScriptRoot empty while param() defaults are evaluated under -File, so a default built from
# it threw before the script ran. Measured 2026-09-24; this repository targets 5.1 (Q78).
if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) }

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
#    UpdateService.MaxDownloadBytes == the release script's installer-size gate. Shipping an
#    asset over the cap silently stops auto-update everywhere; the gate is what turns that into
#    a failed release instead.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$capMb = Get-Pinned 'Services/UpdateService.cs' `
    'MaxDownloadBytes\s*=\s*(\d+)L?\s*\*\s*1024\s*\*\s*1024' 'download cap (UpdateService.cs)'
$gateMb = Get-Pinned 'Tools/Publish-Release.ps1' `
    '\$InstallerCapMB\s*=\s*(\d+)' 'download cap (Publish-Release.ps1 gate)'
if ($capMb -and $gateMb) {
    if ([int]$capMb -ne [int]$gateMb) {
        Fail "installer size cap" "UpdateService.MaxDownloadBytes is ${capMb} MB but Tools/Publish-Release.ps1 refuses installers over ${gateMb} MB. The release gate must refuse exactly what the updater refuses."
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
#    McpRegistrationService.DotnetRuntimeDownloadUrl == Installer.iss ServerRuntimeManualUrl.
#    Both send a user whose server has no runtime to the same page, one from setup and one from
#    Settings, when setup could not put the private copy in place (Q172).
# ---------------------------------------------------------------------------------------------
$script:Checks++
$csUrl = Get-Pinned 'Services/McpRegistrationService.cs' `
    'DotnetRuntimeDownloadUrl\s*=\s*"([^"]+)"' 'runtime download URL (McpRegistrationService.cs)'
$issUrl = Get-Pinned 'Installer.iss' `
    '(?m)^\s*#define\s+ServerRuntimeManualUrl\s+"([^"]+)"' 'runtime download URL (Installer.iss)'
if ($csUrl -and $issUrl) {
    if ($csUrl -cne $issUrl) {
        Fail "runtime download URL" "McpRegistrationService.DotnetRuntimeDownloadUrl is '$csUrl' but Installer.iss ServerRuntimeManualUrl is '$issUrl'. One of the two places that tell a user where to get the runtime is now pointing somewhere else."
    } else {
        Pass "runtime download URL" $csUrl
    }
}

# ---------------------------------------------------------------------------------------------
# 6. outlook_health's registration.status vocabulary.
#    HealthReporting.Registration* == the values both READMEs publish for that field. The C#
#    side is already one definition, and the add-in never sees these strings - it is the DOCS
#    that are the second copy, and Markdown cannot read a C# constant. This is the field an
#    agent is told to read to find out whether the server it is talking to is the registered
#    one, so a renamed status leaves the published vocabulary describing a value the tool no
#    longer emits, with nothing failing anywhere.
#
#    Two directions, deliberately split between the two files: the root README carries a clean
#    "/"-separated list, so that one is compared as a SET (an ADDED status that nobody
#    documented fails here); McpServer/README.md spells each value out with its own explanation
#    in a table cell, so that one is checked for CONTAINMENT (a RENAMED status fails there).
# ---------------------------------------------------------------------------------------------
$script:Checks++
$codeStatuses = Get-PinnedAll 'McpServer/OutlookAI.Core/Services/HealthReporting.cs' `
    '(?m)^\s*public const string Registration[A-Za-z]+\s*=\s*"([^"]+)"\s*;' 'registration statuses (HealthReporting.cs)'
$rootList = Get-Pinned 'README.md' `
    '`status`\s*\(([^)]*)\)' 'registration statuses (README.md)'
if ($codeStatuses -and $rootList) {
    $rootStatuses = @([regex]::Matches($rootList, '`([^`]+)`') | ForEach-Object { $_.Groups[1].Value })
    $left = (@($codeStatuses) | Sort-Object) -join ','
    $right = (@($rootStatuses) | Sort-Object) -join ','
    if ($left -cne $right) {
        Fail "registration status vocabulary (README.md)" "HealthReporting emits {$left} but README.md documents {$right} for outlook_health's registration.status. An agent is told to read that field to learn whether this server is the registered one."
    } else {
        Pass "registration status vocabulary (README.md)" $left
    }
}
$script:Checks++
$serverRow = Get-Pinned 'McpServer/README.md' `
    '(?m)^\|\s*`status`\s*\|(.+)$' 'registration statuses (McpServer/README.md)'
if ($codeStatuses -and $serverRow) {
    $missing = @(@($codeStatuses) | Where-Object { $serverRow -cnotmatch ('`' + [regex]::Escape($_) + '`') })
    if ($missing.Count -gt 0) {
        Fail "registration status vocabulary (McpServer/README.md)" "the registration.status row does not document $($missing -join ', ') - HealthReporting emits $((@($codeStatuses) | Sort-Object) -join ','). The field's own reference table has stopped describing what the tool returns."
    } else {
        Pass "registration status vocabulary (McpServer/README.md)" "all $(@($codeStatuses).Count) documented"
    }
}

# ---------------------------------------------------------------------------------------------
# 7. The ADD-IN's registration status vocabulary.
#    McpRegistrationService.Status* == the list McpServer/README.md publishes for
#    registration.addInStatus. A DIFFERENT vocabulary from #6, for a different field, and the one
#    with weaker protection of the two: the server surfaces these strings VERBATIM (it never
#    compares them against anything), so no compilation anywhere can notice a rename - the add-in
#    keeps writing, the server keeps reporting, and only the published meaning is wrong. That is
#    also why the docs are the second copy again: Markdown cannot read a C# constant.
#
#    Compared as a SET, in both directions, because both directions are real failures here: a
#    RENAMED code leaves the README describing a value the add-in no longer writes, and an ADDED
#    one (awaiting_choice was missing from this list until it was added by hand) leaves an agent
#    reading a status the reference does not explain.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$addInStatuses = Get-PinnedAll 'Services/McpRegistrationService.cs' `
    '(?m)^\s*internal const string Status[A-Za-z]+\s*=\s*"([^"]+)"\s*;' 'add-in status codes (McpRegistrationService.cs)'
$addInDocList = Get-Pinned 'McpServer/README.md' `
    'Status codes:\s*([^.]*)\.' 'add-in status codes (McpServer/README.md)'
if ($addInStatuses -and $addInDocList) {
    $docStatuses = @([regex]::Matches($addInDocList, '`([^`]+)`') | ForEach-Object { $_.Groups[1].Value })
    $left = (@($addInStatuses) | Sort-Object) -join ','
    $right = (@($docStatuses) | Sort-Object) -join ','
    if ($left -cne $right) {
        Fail "add-in status vocabulary (McpServer/README.md)" "McpRegistrationService writes {$left} but McpServer/README.md documents {$right} for outlook_health's registration.addInStatus. The server passes this string through untouched, so nothing else in the product can notice the difference."
    } else {
        Pass "add-in status vocabulary (McpServer/README.md)" $left
    }
}

# ---------------------------------------------------------------------------------------------
# 8. The mail server's private .NET runtime (Q172, decided by the maintainer 2026-11-18).
#    Four files name the same runtime, and no compiler sees more than one of them:
#      * OutlookAI.Core.AddIn's ServerDotnetRuntime.Version - the version setup's helper installs, and
#        the one the add-in looks for - must have the major of the target framework of both
#        shipped executables, OutlookAI.McpServer and OutlookAI.ComHost: their runtimeconfig asks
#        for that major and no other, so a private copy of another major is one they never use;
#      * McpServer/Directory.Build.props' AppHostRelativeDotNet must be ..\..\ and
#        ServerDotnetRuntime.FolderName - the folder setup fills, seen from {app}\McpServer - and
#        AppHostDotNetSearch must look there first;
#      * Installer.iss's [UninstallDelete] must delete that folder, in {localappdata}\OutlookAI;
#      * and the manual download page both #5's copies name must be that major's.
#    Each mismatch is silent where it happens: a server that falls back to a machine-wide runtime,
#    or to none, a runtime the uninstaller leaves behind, a page for the wrong .NET.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$runtimeVersion = Get-Pinned 'McpServer/OutlookAI.Core/AddIn/ServerDotnetRuntime.cs' `
    'public const string Version = "([^"]+)";' 'private runtime version (ServerDotnetRuntime.cs)'
$runtimeFolder = Get-Pinned 'McpServer/OutlookAI.Core/AddIn/ServerDotnetRuntime.cs' `
    'public const string FolderName = "([^"]+)";' 'private runtime folder (ServerDotnetRuntime.cs)'
$serverTfm = Get-Pinned 'McpServer/OutlookAI.McpServer/OutlookAI.McpServer.csproj' `
    '<TargetFramework>net(\d+)\.0-windows</TargetFramework>' 'target framework (OutlookAI.McpServer.csproj)'
$comHostTfm = Get-Pinned 'McpServer/OutlookAI.ComHost/OutlookAI.ComHost.csproj' `
    '<TargetFramework>net(\d+)\.0-windows</TargetFramework>' 'target framework (OutlookAI.ComHost.csproj)'
$relativeDotnet = Get-Pinned 'McpServer/Directory.Build.props' `
    '<AppHostRelativeDotNet>([^<]+)</AppHostRelativeDotNet>' 'apphost runtime folder (McpServer/Directory.Build.props)'
$dotnetSearch = Get-Pinned 'McpServer/Directory.Build.props' `
    '<AppHostDotNetSearch>([^<]+)</AppHostDotNetSearch>' 'apphost runtime search (McpServer/Directory.Build.props)'
$uninstallText = Read-Source 'Installer.iss'
if ($runtimeVersion -and $runtimeFolder -and $serverTfm -and $comHostTfm -and $relativeDotnet -and $dotnetSearch -and $null -ne $uninstallText -and $csUrl -and $issUrl) {
    $major = $runtimeVersion.Split('.')[0]
    $problems = @()
    if ($serverTfm -ne $major -or $comHostTfm -ne $major) {
        $problems += "ServerDotnetRuntime.Version is $runtimeVersion, but OutlookAI.McpServer targets net$serverTfm.0 and OutlookAI.ComHost net$comHostTfm.0: their runtimeconfig asks for their own major, and a private runtime of another is one they never start on."
    }
    if ($relativeDotnet -cne ('..\..\' + $runtimeFolder)) {
        $problems += "AppHostRelativeDotNet is '$relativeDotnet', not '..\..\$runtimeFolder': from {app}\McpServer the apphost would look somewhere setup never puts the runtime."
    }
    if (-not ($dotnetSearch -split ';')[0].Trim().Equals('AppRelative', [System.StringComparison]::Ordinal)) {
        $problems += "AppHostDotNetSearch is '$dotnetSearch': the apphost must look in the private runtime folder first, or a machine-wide .NET shadows the copy setup installed."
    }
    $uninstallLine = '(?m)^Type:\s*filesandordirs;\s*Name:\s*"\{localappdata\}\\OutlookAI\\' + [regex]::Escape($runtimeFolder) + '"\s*$'
    if ($uninstallText -notmatch '(?m)^\[UninstallDelete\]' -or $uninstallText -notmatch $uninstallLine) {
        $problems += "Installer.iss's [UninstallDelete] does not delete {localappdata}\OutlookAI\${runtimeFolder}: an uninstall would leave the private runtime behind."
    }
    $page = 'https://dotnet.microsoft.com/download/dotnet/' + $major + '.0'
    if ($csUrl -cne $page) {
        $problems += "the manual download page is '$csUrl', not '$page', the page of .NET $major."
    }
    if ($problems.Count -gt 0) {
        Fail "private .NET runtime" ($problems -join ' ')
    } else {
        Pass "private .NET runtime" "$runtimeVersion in {localappdata}\OutlookAI\$runtimeFolder, net$major.0, $dotnetSearch"
    }
}

# ---------------------------------------------------------------------------------------------
# 10. Deliberate failures raised inside the COM host.
#     Every exception type thrown by the code BEHIND the IOutlookSession contract must be a
#     `case nameof(...)` in ComHostErrorMapper's switch.
#
#     The COM host runs in a CHILD PROCESS, so an exception cannot simply propagate: the child
#     renders it as four fields on the wire and the parent rebuilds the closest equivalent from
#     that switch. A type the switch does not name is not lost, but it is DEMOTED - it arrives
#     as ComHostRemoteException, and every layer that branches on exception type stops
#     recognising it. OutlookTools.GuardAsync picks the advice an agent reads that way, and
#     ComGateway keys its disconnect-and-rebuild on COMException; on 2026-08-18 both of those
#     were found to have been unreachable in this process since the split, because everything
#     arrived wrapped and nothing said so.
#
#     The switch is therefore a SECOND COPY of a list that exists nowhere else - it is spread
#     across every `throw new` in the COM layer, and no compilation can relate the two. Adding
#     `throw new SomeException(...)` down there compiles, ships, and quietly demotes itself.
#     That is exactly the shape this script exists for.
#
#     Scope is the session side, not the transport: Core/Com and Core/Text are what runs behind
#     the contract, and ComHost/Host is the dispatch that calls it. ComHost/Protocol is left out
#     deliberately, and the reason has narrowed since 2026-08-18. A DESYNC still takes the host
#     down and never becomes a ComHostError. A response too large to frame no longer does: the
#     serve loop answers with one, stamped ComHostResponseTooLargeException. But it does that by
#     WRITING the type name onto the wire error, not by throwing - the failure happens while
#     encoding the reply, past the point where a throw could still produce one - so there is no
#     `throw new` for this scan to find, and including the directory would only make it demand a
#     mapper case for the desync exception, which genuinely never travels.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$comHostSourceDirs = @(
    'McpServer/OutlookAI.Core/Com',
    'McpServer/OutlookAI.Core/Text',
    'McpServer/OutlookAI.ComHost/Host')

$thrownTypes = New-Object System.Collections.Generic.HashSet[string]
$comHostFilesScanned = 0
foreach ($dir in $comHostSourceDirs) {
    $fullDir = Join-Path $RepoRoot $dir
    if (-not (Test-Path $fullDir)) {
        Fail "COM host failure types" "$dir does not exist (did it move? this check now proves nothing)"
        continue
    }

    foreach ($file in @(Get-ChildItem -LiteralPath $fullDir -Filter '*.cs' -File)) {
        $comHostFilesScanned++
        $source = Get-Content -LiteralPath $file.FullName -Raw
        foreach ($m in [regex]::Matches($source, 'throw\s+new\s+([A-Za-z_][A-Za-z0-9_.]*)\s*\(')) {
            # Namespace-qualified throws are written both ways in these files, so compare on
            # the leaf name - which is also what nameof() produces on the other side.
            $null = $thrownTypes.Add(($m.Groups[1].Value -split '\.')[-1])
        }
    }
}

$mapperCases = Get-PinnedAll 'McpServer/OutlookAI.ComHost/Supervision/ComHostErrorMapper.cs' `
    'case nameof\(([A-Za-z_][A-Za-z0-9_]*)\)\s*:' 'COM host error mapper cases'

if ($comHostFilesScanned -eq 0) {
    Fail "COM host failure types" "no .cs files found under $($comHostSourceDirs -join ', ') - the layout changed and this check no longer proves anything."
} elseif ($thrownTypes.Count -eq 0) {
    Fail "COM host failure types" "found no 'throw new' anywhere in $comHostFilesScanned COM-host source files, which cannot be right - the pattern stopped matching and this check has switched itself off."
} elseif ($mapperCases) {
    $unmapped = @($thrownTypes | Where-Object { $mapperCases -cnotcontains $_ } | Sort-Object)
    if ($unmapped.Count -gt 0) {
        Fail "COM host failure types" "the COM host raises $($unmapped -join ', ') but ComHostErrorMapper has no case for $(if ($unmapped.Count -eq 1) { 'it' } else { 'them' }). Raised in the child, $(if ($unmapped.Count -eq 1) { 'it arrives' } else { 'they arrive' }) as ComHostRemoteException instead, so the tool layer cannot choose advice by type and ComGateway cannot recognise a disconnect - the message survives, the meaning does not. Add the case to ComHostErrorMapper.ToException, or raise a type it already models."
    } else {
        Pass "COM host failure types" "$($thrownTypes.Count) raised, all modelled ($comHostFilesScanned files scanned)"
    }
}

# ---------------------------------------------------------------------------------------------
# 11. Live-tier capability vocabulary.
#     T1/LiveTierInventoryTests holds the ONE vocabulary a live test may use to say what it needs
#     of a machine, and Docs/live-tier-on-the-vm.md is where a human reads it. Since the LiveTier
#     axis was deleted, that list decides everything: which bucket a test is in is a question
#     asked of Requires, and the runbook's own filter expressions are written against these
#     names. Add a capability to the C# and not the runbook and the runbook silently
#     under-reports what the VM has to provide - which is the document somebody builds the VM
#     from. C# and Markdown cannot see each other, so this is the mechanism.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$inventorySource = Read-Source 'McpServer/OutlookAI.McpServer.Tests/T1/LiveTierInventoryTests.cs'
$runbook = Read-Source 'Docs/live-tier-on-the-vm.md'
if ($inventorySource -and $runbook) {
    # The WHOLE vocabulary, not just the production-only subset: with one axis left, a value the
    # runbook never mentions is a value nobody building the VM is told about, whether or not it
    # is the one that pins a test to the dev machine.
    $block = [regex]::Match($inventorySource,
        'string\[\]\s+AllCapabilities\s*=\s*\{(?<body>[^}]*)\}')
    if (-not $block.Success) {
        Fail "live-tier capability vocabulary" "could not find AllCapabilities in LiveTierInventoryTests.cs - the file changed shape and this check no longer proves anything."
    } else {
        # Entries are either a quoted literal or the name of a `const string` in the same file
        # (the two that other code references by name). Resolve both, or a renamed constant would
        # quietly drop out of the checked set.
        # Comments first, then split: the explanatory comments between entries contain commas of
        # their own, and splitting before stripping them tears entries in half.
        $body = $block.Groups['body'].Value -replace '(?m)//[^\r\n]*', ''
        $capabilities = @(
            foreach ($entry in ($body -split ',')) {
                $entry = $entry.Trim()
                if (-not $entry) { continue }
                $literal = [regex]::Match($entry, '^"([A-Za-z]+)"$')
                if ($literal.Success) { $literal.Groups[1].Value; continue }
                if ($entry -notmatch '^[A-Za-z_]\w*$') { continue }
                $const = [regex]::Match($inventorySource, "const\s+string\s+$([regex]::Escape($entry))\s*=\s*""([A-Za-z]+)""")
                if ($const.Success) { $const.Groups[1].Value }
                else { "<unresolved:$entry>" }
            })
        $unresolved = @($capabilities | Where-Object { $_ -like '<unresolved:*' })
        if ($capabilities.Count -eq 0) {
            Fail "live-tier capability vocabulary" "AllCapabilities parsed as empty, which cannot be right - the pattern stopped matching and this check has switched itself off."
        } elseif ($unresolved.Count -gt 0) {
            Fail "live-tier capability vocabulary" "AllCapabilities names $($unresolved -join ', ') but no matching 'const string' exists in LiveTierInventoryTests.cs, so this check cannot tell what those capabilities are called and would pass over them in silence."
        } else {
            $undocumented = @($capabilities | Where-Object { $runbook -cnotmatch [regex]::Escape("``$_``") } | Sort-Object)
            if ($undocumented.Count -gt 0) {
                Fail "live-tier capability vocabulary" "LiveTierInventoryTests allows $($undocumented -join ', ') but Docs/live-tier-on-the-vm.md never mentions $(if ($undocumented.Count -eq 1) { 'it' } else { 'them' }). The runbook is what somebody reads to decide whether a live test can move to a test machine, and what they build that machine from; a capability missing from it is a requirement nobody is told about."
            } else {
                Pass "live-tier capability vocabulary" "$($capabilities.Count) capabilities, all documented in the runbook"
            }
        }
    }
}

# ---------------------------------------------------------------------------------------------
# 11b. The live tier's Writes trait (Q74, 2026-10-03).
#     Beside the capability vocabulary of #11, one trait says what a live test DOES to the machine:
#     Writes=Nothing, declared per method, the claim the maintainer's read-only workstation run rests on
#     (Testbed/README.md section 4d). Its key and its one value live in T2/LiveRunFilters.cs, which
#     derives the run filters from them, and Docs/live-tier-on-the-vm.md is where a human reads what
#     it means. C# and Markdown cannot see each other, so a renamed key or value would leave the
#     runbook describing a trait the suite no longer carries - and a filter clause on a trait key no
#     test carries is one VSTest does not evaluate at all. Same shape as #11, one trait wide.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$filtersSource = Read-Source 'McpServer/OutlookAI.McpServer.Tests/T2/LiveRunFilters.cs'
if ($filtersSource -and $runbook) {
    $writesKey = [regex]::Match($filtersSource, 'const\s+string\s+WritesTrait\s*=\s*"([A-Za-z]+)"')
    $writesValue = [regex]::Match($filtersSource, 'const\s+string\s+WritesNothing\s*=\s*"([A-Za-z]+)"')
    if (-not $writesKey.Success -or -not $writesValue.Success) {
        Fail "live-tier Writes trait" "could not find WritesTrait and WritesNothing in T2/LiveRunFilters.cs - the file changed shape and this check no longer proves anything."
    } else {
        $writesToken = '`' + $writesKey.Groups[1].Value + '=' + $writesValue.Groups[1].Value + '`'
        if ($runbook -cnotmatch [regex]::Escape($writesToken)) {
            Fail "live-tier Writes trait" "T2/LiveRunFilters.cs declares the trait $writesToken but Docs/live-tier-on-the-vm.md never mentions it. It is what the maintainer's read-only workstation run selects on; a trait the runbook does not describe is a claim nobody reading it can check."
        } else {
            Pass "live-tier Writes trait" "$writesToken documented in the runbook"
        }
    }
}

# ---------------------------------------------------------------------------------------------
# 12. The Outlook Search value name, and the fact that it is spelled ONCE.
#     AddInServerContract.DisableServerAssistedSearchValueName is the single definition of a
#     value name that used to be typed out six times (the add-in's tuning catalog, twice in
#     HealthReporting, three times in the T2 live test). Both trees compile that file, so a
#     RENAME of the constant cannot drift - it breaks both compilations. Two things still can,
#     and neither is visible to any compiler:
#
#     (a) SOMEBODY TYPES THE LITERAL AGAIN. Re-introducing "DisableServerAssistedSearch" as a
#         string somewhere is a second copy of a closed mirror, and the drift it re-opens is
#         silent in the worst direction: the add-in goes on writing the value, the server reads a
#         name nobody wrote and reports uiSearchBackend as "server-assisted" on a tuned machine
#         for ever, which is what show_search_results uses to warn an agent that the user's
#         Outlook search box is looking somewhere else.
#     (b) THE ONE COPY IS MISTYPED. It is OUTLOOK'S name, not ours - it cannot be renamed, only
#         got wrong - and a wrong name is agreed on by both sides, so nothing in the product
#         disagrees with anything. The only other place it is written down is McpServer/README.md,
#         which publishes it to users as the value they may toggle; Markdown cannot read a C#
#         constant, so this is the mechanism.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$searchValueName = Get-Pinned 'Services/AddInServerContract.cs' `
    'DisableServerAssistedSearchValueName\s*=\s*"([^"]+)"' 'Outlook Search value name (AddInServerContract.cs)'
if ($searchValueName) {
    # Every .cs file in the tree, minus build output and dot-directories (which is also what
    # keeps a local git worktree under .claude/ from being scanned as if it were source).
    $literal = '"' + $searchValueName + '"'
    $copies = @(
        Get-ChildItem -LiteralPath $RepoRoot -Filter '*.cs' -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object {
                $rel = $_.FullName.Substring($RepoRoot.Length).TrimStart('\', '/')
                -not ($rel -split '[\\/]' | Where-Object { $_ -like '.*' -or $_ -eq 'bin' -or $_ -eq 'obj' })
            } |
            Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -cmatch [regex]::Escape($literal) } |
            ForEach-Object { $_.FullName.Substring($RepoRoot.Length).TrimStart('\', '/').Replace('\', '/') })

    if ($copies.Count -eq 0) {
        Fail "Outlook Search value name is defined once" "no .cs file contains the literal $literal, not even Services/AddInServerContract.cs where it is declared - this check has stopped finding what it counts and proves nothing."
    } elseif ($copies.Count -gt 1 -or $copies[0] -ne 'Services/AddInServerContract.cs') {
        Fail "Outlook Search value name is defined once" "$literal is written out in $($copies.Count) file$(if ($copies.Count -ne 1) { 's' }) - $($copies -join ', ') - but it may exist only in Services/AddInServerContract.cs. Everything else reads AddInServerContract.DisableServerAssistedSearchValueName (the server re-exports it as HealthReporting.DisableServerAssistedSearchValueName). A second copy re-opens a mirror that fails silently: the add-in keeps writing the value, the server reads a name nobody wrote and reports uiSearchBackend as 'server-assisted' on a tuned machine for ever."
    } else {
        Pass "Outlook Search value name is defined once" "$searchValueName (Services/AddInServerContract.cs)"
    }

    $serverReadme = Read-Source 'McpServer/README.md'
    if ($serverReadme) {
        $script:Checks++
        if ($serverReadme -cnotmatch [regex]::Escape($searchValueName)) {
            Fail "Outlook Search value name is documented" "AddInServerContract spells the Outlook Search value '$searchValueName' but McpServer/README.md never mentions it. That README is where the value the user may toggle is published, and it is the only copy outside the C# - a typo in the constant is agreed on by both halves of the product and disagrees with nothing except the documentation."
        } else {
            Pass "Outlook Search value name is documented" "$searchValueName (McpServer/README.md)"
        }
    }
}

# ---------------------------------------------------------------------------------------------
# 13. The Outlook Search KEY PATH is built by the shared helper on BOTH sides.
#     The add-in writes ...\Office\<major>\Outlook\Search and the server reads it. Sharing
#     Services/OfficeVersions.cs settled the <major> (check #4's neighbour); it did NOT settle
#     the construction, and until 2026-08-24 the add-in concatenated the whole path by hand while
#     the server went through OfficeVersions.OutlookKeyPath. One address, two spellings, across a
#     boundary no compiler crosses - and the add-in half is unreachable from every test in this
#     repository, because it is net48/VSTO and builds only under MSBuild.
#
#     So the invariant is textual and it is the only thing that can hold it: neither file may
#     concatenate an Office hive root of its own, and both must call the shared builder. A prose
#     mention of the path is not a build (the diagnostic sentence in HealthReporting contains
#     one) - the pattern matches a hive-root literal being CONCATENATED, which is what a
#     hand-built path looks like and what a sentence never does.
#
#     A deliberate rename of the shared member means updating the member name below; that is the
#     check demanding to be told, not a false alarm.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$sharedSearchBuilder = 'OfficeVersions.OutlookSearchKeyPath'
$handBuiltHiveRoot = '@"Software\\(?:Policies\\Microsoft|Microsoft)\\Office\\"\s*\+'
$searchPathSides = @{
    'Services/OutlookTuningService.cs'                       = 'the add-in, which WRITES the value'
    'McpServer/OutlookAI.Core/Services/HealthReporting.cs'   = 'the server, which READS it back as uiSearchBackend'
}
$searchPathProblems = @()
foreach ($side in @($searchPathSides.Keys | Sort-Object)) {
    $text = Read-Source $side
    if ($null -eq $text) { continue }
    if ($text -cnotmatch [regex]::Escape($sharedSearchBuilder + '(')) {
        $searchPathProblems += "$side ($($searchPathSides[$side])) no longer calls $sharedSearchBuilder"
    }
    $handBuilt = [regex]::Matches($text, $handBuiltHiveRoot)
    if ($handBuilt.Count -gt 0) {
        $searchPathProblems += "$side ($($searchPathSides[$side])) concatenates an Office hive root of its own, $($handBuilt.Count) time$(if ($handBuilt.Count -ne 1) { 's' })"
    }
}
if ($searchPathProblems.Count -gt 0) {
    Fail "Outlook Search key path is built by one expression" "$($searchPathProblems -join '; '). Both sides must build that key through $sharedSearchBuilder (and any Policies-hive path through OfficeVersions.PolicyOutlookKeyPath), because the add-in's half is net48/VSTO and no test in this repository can reach it - a mistyped path there writes into a key Outlook never reads while everything still compiles and outlook_health still answers."
} else {
    Pass "Outlook Search key path is built by one expression" "$sharedSearchBuilder, both sides"
}

# ---------------------------------------------------------------------------------------------
# 14. A build of the add-in outside Visual Studio does not register it (Q81, 2026-09-27).
#     The VSTO build targets write the BUILD MACHINE's registry on every build: RegisterOfficeAddin
#     runs SetOffice2007AddInRegistration, which points HKCU\...\Outlook\Addins\OutlookAI - the key
#     the installer writes - at the build output; SetInclusionListEntry adds a VSTO trust entry for
#     it; and Clean runs UnregisterOfficeAddin and RemoveOfficeAddInSecurity. Until OutlookAI.csproj
#     took that away, every plain msbuild of the add-in - in any checkout, any agent worktree -
#     silently repointed the maintainer's own Outlook at whatever it had just built (AGENTS.md,
#     "The add-in on the maintainer's workstation").
#
#     This is the opposite shape again (like #12 and #13): a guard that exists ONCE, and that no
#     compilation, test or build can see. Deleting it breaks nothing - the build simply starts
#     registering again. So its shape is pinned here, as the csproj's own comment explains it and
#     as Tools/Switch-AddInBuild.ps1 -SelfTest already reads it locally:
#       * each of the two writers has a stand-in: ONE UsingTask, conditioned exactly
#         '$(BuildingInsideVisualStudio)' != 'true', declared BEFORE the VSTO targets import (the
#         first UsingTask for a name wins), an inline RoslynCodeTaskFactory task that only logs its
#         NOT-REGISTERED-OUTSIDE-VS line - and NOT Override="true": Publish-AddInPayload.ps1 and
#         Switch-AddInBuild.ps1 bring Override stand-ins of their own, which win over these only
#         while these are not;
#       * RegisterFormRegions has NO stand-in, on purpose - a form region added later would be
#         dropped from the manifest of every command-line build, the release build's included;
#       * AFTER the import, under the same condition, PrepareForRun and VSTOClean lose the targets
#         that call the writers, through $([MSBuild]::Unescape(...)) - a bare .Replace() returns the
#         list escaped, and MSBuild then looks for ONE target named after all of them (MSB4057).
# ---------------------------------------------------------------------------------------------
$script:Checks++
$q81 = 'the add-in build does not register itself outside Visual Studio (Q81)'
$projText = Read-Source 'OutlookAI.csproj'
if ($null -ne $projText) {
    $q81Problems = @()
    $outsideVs = "'`$(BuildingInsideVisualStudio)' != 'true'"
    $proj = New-Object System.Xml.XmlDocument
    $parsed = $true
    try { $proj.LoadXml($projText) } catch { $parsed = $false; $q81Problems += "OutlookAI.csproj does not parse as XML ($($_.Exception.Message)), so nothing here can be checked." }
    if ($parsed) {
        $ns = New-Object System.Xml.XmlNamespaceManager($proj.NameTable)
        $ns.AddNamespace('m', 'http://schemas.microsoft.com/developer/msbuild/2003')
        $top = @($proj.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
        $vstoImport = @($top | Where-Object {
                $_.LocalName -eq 'Import' -and
                $_.GetAttribute('Project').EndsWith('\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets', [System.StringComparison]::OrdinalIgnoreCase)
            })
        if ($vstoImport.Count -ne 1) {
            $q81Problems += "OutlookAI.csproj imports OfficeTools\Microsoft.VisualStudio.Tools.Office.targets $($vstoImport.Count) time(s), not once - the file changed shape and this check no longer proves anything."
        }
        else {
            $importAt = [array]::IndexOf($top, $vstoImport[0])
            foreach ($task in @('SetOffice2007AddInRegistration', 'SetInclusionListEntry')) {
                $standIns = @($top | Where-Object { $_.LocalName -eq 'UsingTask' -and $_.GetAttribute('TaskName') -ceq $task })
                if ($standIns.Count -ne 1) {
                    $q81Problems += "${task}: $($standIns.Count) UsingTask(s) for it, not one - its stand-in is what keeps a command-line build from writing the registry."
                    continue
                }
                $ut = $standIns[0]
                if ($ut.GetAttribute('Condition') -cne $outsideVs) { $q81Problems += "${task}: its Condition is `"$($ut.GetAttribute('Condition'))`", not `"$outsideVs`"." }
                if ($ut.HasAttribute('Override')) { $q81Problems += "${task}: it carries Override=`"$($ut.GetAttribute('Override'))`" - the testbed's and Switch-AddInBuild.ps1's Override stand-ins win over this one only while it has none." }
                if ([array]::IndexOf($top, $ut) -gt $importAt) { $q81Problems += "${task}: declared AFTER the VSTO targets import, so the real task is the first UsingTask for the name and the stand-in never runs." }
                if ($ut.GetAttribute('TaskFactory') -cne 'RoslynCodeTaskFactory') { $q81Problems += "${task}: TaskFactory `"$($ut.GetAttribute('TaskFactory'))`", not RoslynCodeTaskFactory - it is no longer the inline stand-in." }
                $code = $ut.SelectSingleNode('m:Task/m:Code', $ns)
                if ($null -eq $code) { $q81Problems += "${task}: no inline Code - it is no longer the stand-in." }
                else {
                    if (-not $code.InnerText.Contains('NOT-REGISTERED-OUTSIDE-VS')) { $q81Problems += "${task}: its code no longer logs NOT-REGISTERED-OUTSIDE-VS." }
                    foreach ($writer in @('Registry', 'File.', 'Process.')) {
                        if ($code.InnerText.Contains($writer)) { $q81Problems += "${task}: its code names '$writer' - the stand-in may only log." }
                    }
                }
            }
            if (@($top | Where-Object { $_.LocalName -eq 'UsingTask' -and $_.GetAttribute('TaskName') -ceq 'RegisterFormRegions' }).Count -gt 0) {
                $q81Problems += "RegisterFormRegions has a stand-in. It is left alone on purpose: the add-in has no form region, so the real task writes nothing, and a stand-in would silently drop any future one from the manifest of every command-line build, the release build's included."
            }
            $chains = [ordered]@{
                'PrepareForRunDependsOn' = "`$([MSBuild]::Unescape(`$(PrepareForRunDependsOn.Replace('RegisterOfficeAddin', ''))))"
                'VSTOCleanDependsOn'     = "`$([MSBuild]::Unescape(`$(VSTOCleanDependsOn.Replace('RemoveOfficeAddInSecurity', '').Replace('UnregisterOfficeAddin', ''))))"
            }
            $groupsAfter = @($top | Where-Object { $_.LocalName -eq 'PropertyGroup' -and $_.GetAttribute('Condition') -ceq $outsideVs -and [array]::IndexOf($top, $_) -gt $importAt })
            foreach ($property in $chains.Keys) {
                $found = @($groupsAfter | ForEach-Object { $_.SelectNodes("m:$property", $ns) } | Where-Object { $_.InnerText -ceq $chains[$property] })
                if ($found.Count -ne 1) {
                    $q81Problems += "${property}: $($found.Count) of <$property>$($chains[$property])</$property> in a PropertyGroup conditioned `"$outsideVs`" after the VSTO import, not one - without it the targets that call the registry writers are scheduled again outside Visual Studio."
                }
            }
        }
    }
    if ($q81Problems.Count -gt 0) {
        Fail $q81 "$($q81Problems -join ' | ') The guard is the only thing between a plain msbuild of the add-in and the maintainer's own Outlook loading that build; see the comment above its UsingTasks in OutlookAI.csproj."
    } else {
        Pass $q81 "two stand-ins ahead of the VSTO import, RegisterFormRegions untouched, both chains stripped after it"
    }
}

# ---------------------------------------------------------------------------------------------
# 15. The add-in reaches OutlookAI.Core through ONE namespace, and nothing in that namespace runs
#     Outlook's COM session, a COM apartment thread or the search index (Q164 (a), decided by the
#     maintainer 2026-11-18).
#     Since Q167 (a) the add-in references Core's net48 build (OutlookAI.csproj). Core also holds
#     what must never run inside Outlook's own process - the COM session that attaches to Outlook
#     and registers a message filter, the STA runner, the index search over OLE DB, the MAPI
#     probes - where a hang or a crash takes Outlook down with it. So the add-in may use one part of
#     Core, the namespace OutlookAI.Core.AddIn in McpServer/OutlookAI.Core/AddIn/, and that part
#     uses nothing else of Core. No compiler holds either half - the rest of Core is public, because
#     the server uses it - so the rule is text, as #12 to #14 are:
#       (a) every reference to OutlookAI.Core in what the add-in compiles (OutlookAI.csproj's Compile
#           items) names OutlookAI.Core.AddIn: a using directive, an alias, a qualified or global::
#           name, and a type or assembly name in a string, which is how reflection would get round
#           the rule. And no Core.X written relative to the add-in's own OutlookAI namespace;
#       (b) every file of Core - linked ones included - that declares the namespace lies in that
#           folder, and every file in the folder declares it or a namespace inside it; and no file
#           of Core declares OutlookAI.Core itself, whose types a file in the folder could name
#           without writing a namespace at all;
#       (c) the folder's files name no other namespace of Core, whole or relative to their own (Com.X,
#           Services.X), and none of the calls that attach to Outlook, start a COM apartment thread,
#           register a COM message filter, open a MAPI session or open the index.
#     Comments are taken out first: they are not code. Strings stay in for (a) and for (c)'s calls -
#     "Outlook.Application" and "Search.CollatorDSO" live in strings - and come out for the names
#     only code can carry, Core.X and (c)'s relative namespaces.
# ---------------------------------------------------------------------------------------------
$script:Checks++
$q164 = 'the add-in reaches OutlookAI.Core only through OutlookAI.Core.AddIn, which stays out of Outlook''s COM session, COM threads and the index (Q164)'
$facadeNamespace = 'OutlookAI.Core.AddIn'
$coreDirRel = 'McpServer/OutlookAI.Core'
$facadeDirRel = 'McpServer/OutlookAI.Core/AddIn'
# One regex over a C# text for its comments and its string and character literals; the leftmost match
# wins, so a // inside a string stays in the string. Two of .NET's own replacements, with no script
# code per match, which kept this check from taking minutes on a busy machine: every comment becomes a
# space, and every literal stays, followed by a space, or becomes "" ('' for a character). A space
# rather than nothing, so that two tokens a comment separated stay apart.
$csCommentsAndLiterals = New-Object System.Text.RegularExpressions.Regex('//[^\n]*|/\*[\s\S]*?\*/|(?<lit>(?<q>")"{2,}[\s\S]*?"{3,}|@\$?(?<q>")(?:[^"]|"")*"|\$@(?<q>")(?:[^"]|"")*"|\$?(?<q>")(?:[^"\\\n]|\\.)*"|(?<q>'')(?:[^''\\\n]|\\.)+'')')
function Get-CSharpCode([string] $Text, [switch] $KeepStrings) {
    if ($KeepStrings) { return $csCommentsAndLiterals.Replace($Text, '${lit} ') }
    return $csCommentsAndLiterals.Replace($Text, '${q}${q} ')
}
# Every match of a pattern as "file: what it found", the rest of its line included.
function Find-CodeMatches([string] $Code, [string] $Pattern, [string] $Rel, [switch] $IgnoreCase) {
    $options = [System.Text.RegularExpressions.RegexOptions]::None
    if ($IgnoreCase) { $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
    $out = @()
    foreach ($m in [regex]::Matches($Code, $Pattern, $options)) {
        $end = $Code.IndexOf([char]10, $m.Index)
        if ($end -lt 0) { $end = $Code.Length }
        $snippet = $Code.Substring($m.Index, [Math]::Min($end - $m.Index, 60))
        $out += ('{0}: {1}' -f $Rel, ($snippet.Trim() -replace '\s+', ' '))
    }
    return $out
}
# The namespaces a text declares, block-scoped or file-scoped.
function Get-DeclaredNamespaces([string] $Text) {
    return @([regex]::Matches($Text, '(?m)^\s*namespace\s+([A-Za-z_][\w.]*)\s*[;{]') | ForEach-Object { $_.Groups[1].Value })
}
# OutlookAI.Core named anywhere, unless .AddIn follows it.
$namesCoreButNotFacade = '(?<![\w.])OutlookAI\s*\.\s*Core\b(?!\s*\.\s*AddIn\b)'
# Core.X written relative to an enclosing OutlookAI namespace.
$namesCoreRelatively = '(?<![\w.])Core\s*\.\s*(?!AddIn\b)[A-Za-z_]\w*'
# What a file in the facade must not call or name, and what each one does.
$inProcessForbidden = [ordered]@{
    'GetActiveObject'         = 'attaches to a running Outlook'
    'GetTypeFromProgID'       = 'creates or attaches to a COM server by name'
    'GetTypeFromCLSID'        = 'creates or attaches to a COM server by class id'
    'Outlook\.Application'    = 'names Outlook''s own COM class'
    'CoRegisterMessageFilter' = 'registers a COM message filter'
    'IMessageFilter'          = 'is a COM message filter'
    'ApartmentState'          = 'starts a COM apartment thread'
    'CoInitialize'            = 'joins a thread to a COM apartment'
    'OleDb'                   = 'is the search index''s OLE DB provider'
    'CollatorDSO'             = 'is the Windows Search provider'
    'SystemIndex'             = 'is the Windows Search catalog'
    'MAPIInitialize'          = 'starts Extended MAPI'
    'MAPILogon'               = 'opens a MAPI session'
    'OutlookComSession'       = 'is Core''s COM session'
    'ComStaRunner'            = 'is Core''s STA runner'
}
$q164Problems = @()
$addinProjText = Read-Source 'OutlookAI.csproj'
$coreProjText = Read-Source "$coreDirRel/OutlookAI.Core.csproj"
if ($null -ne $addinProjText -and $null -ne $coreProjText) {
    $addinProj = New-Object System.Xml.XmlDocument
    $addinProj.LoadXml($addinProjText)
    $msb = New-Object System.Xml.XmlNamespaceManager($addinProj.NameTable)
    $msb.AddNamespace('m', 'http://schemas.microsoft.com/developer/msbuild/2003')
    $addinFiles = @($addinProj.SelectNodes('//m:Compile', $msb) | ForEach-Object { $_.GetAttribute('Include').Replace('\', '/') })

    # Core's files: every .cs under its folder but its build output, and every file it links.
    # The full path, so that a relative -RepoRoot cuts the same relative names as the default.
    $rootLength = (Resolve-Path -LiteralPath $RepoRoot).ProviderPath.TrimEnd('\', '/').Length + 1
    $coreDir = Join-Path $RepoRoot $coreDirRel
    $coreFiles = @(Get-ChildItem -LiteralPath $coreDir -Filter '*.cs' -File -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { $_.FullName.Substring($rootLength).Replace('\', '/') } |
            Where-Object { -not (($_.Substring($coreDirRel.Length + 1) -split '/') | Where-Object { $_ -eq 'bin' -or $_ -eq 'obj' }) })
    $coreProj = New-Object System.Xml.XmlDocument
    $coreProj.LoadXml($coreProjText)
    foreach ($c in @($coreProj.SelectNodes('//Compile[@Include]'))) {
        $linked = [System.IO.Path]::GetFullPath((Join-Path $coreDir $c.GetAttribute('Include')))
        $coreFiles += $linked.Substring($rootLength).Replace('\', '/')
    }

    # The scans must have found what they read, or this check proves nothing.
    foreach ($known in @('ThisAddIn.cs', 'TaskPane/AITaskPane.cs')) {
        if ($addinFiles -notcontains $known) { $q164Problems += "OutlookAI.csproj compiles no $known - the add-in's file list changed shape and part (a) no longer proves anything." }
    }
    if ($coreFiles -notcontains "$coreDirRel/Com/OutlookComSession.cs") { $q164Problems += "no $coreDirRel/Com/OutlookComSession.cs among Core's files - the scan of Core changed shape and parts (b) and (c) no longer prove anything." }

    # (a) The add-in's own sources.
    foreach ($rel in $addinFiles) {
        $text = Read-Source $rel
        if ($null -eq $text) { continue }
        $hits = @(Find-CodeMatches -Code (Get-CSharpCode $text -KeepStrings) -Pattern $namesCoreButNotFacade -Rel $rel) +
            @(Find-CodeMatches -Code (Get-CSharpCode $text) -Pattern $namesCoreRelatively -Rel $rel)
        foreach ($h in $hits) { if ($h) { $q164Problems += "(a) the add-in names Core outside $facadeNamespace - $h" } }
    }

    # (b) Where the namespace is declared, and that the root namespace is not.
    $siblings = @()
    $facadeFiles = @()
    foreach ($rel in $coreFiles) {
        $text = Read-Source $rel
        if ($null -eq $text) { continue }
        $inFacadeDir = $rel.StartsWith($facadeDirRel + '/', [System.StringComparison]::OrdinalIgnoreCase)
        if ($inFacadeDir) { $facadeFiles += $rel }
        # Read from the raw text, which is cheap; a file in the folder, and any other whose declarations
        # would break the rule, is read again without its comments and literals, so a namespace in a
        # comment is not a declaration.
        $declared = @(Get-DeclaredNamespaces $text)
        $suspect = @($declared | Where-Object { $_ -ceq 'OutlookAI.Core' -or $_ -ceq $facadeNamespace -or $_.StartsWith($facadeNamespace + '.', [System.StringComparison]::Ordinal) })
        if ($inFacadeDir -or $suspect.Count -gt 0) { $declared = @(Get-DeclaredNamespaces (Get-CSharpCode $text)) }
        foreach ($n in $declared) {
            $inFacade = ($n -ceq $facadeNamespace -or $n.StartsWith($facadeNamespace + '.', [System.StringComparison]::Ordinal))
            if ($inFacade -and -not $inFacadeDir) { $q164Problems += "(b) $rel declares $n outside $facadeDirRel/ - the add-in could use it, and part (c) would never read it." }
            if ($inFacadeDir -and -not $inFacade) { $q164Problems += "(b) $rel lies in $facadeDirRel/ but declares $n, not $facadeNamespace." }
            if ($n -ceq 'OutlookAI.Core') { $q164Problems += "(b) $rel declares the root namespace OutlookAI.Core, whose types a file in $facadeDirRel/ could use by name alone." }
            if (-not $inFacade -and $n.StartsWith('OutlookAI.Core.', [System.StringComparison]::Ordinal)) { $siblings += $n.Split('.')[2] }
        }
        if ($inFacadeDir -and $declared.Count -eq 0) { $q164Problems += "(b) $rel lies in $facadeDirRel/ but declares no namespace." }
    }
    $siblings = @($siblings | Sort-Object -Unique)
    foreach ($known in @('Com', 'IndexSearch')) {
        if ($siblings -notcontains $known) { $q164Problems += "no file of Core declares OutlookAI.Core.$known - the namespace scan changed shape and part (c) no longer proves anything." }
    }

    # (c) What the facade's files may not name or call.
    $namesSiblingRelatively = '(?<![\w.])(?:' + (($siblings | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\s*\.\s*[A-Za-z_]\w*'
    foreach ($rel in $facadeFiles) {
        $text = Read-Source $rel
        if ($null -eq $text) { continue }
        $withStrings = Get-CSharpCode $text -KeepStrings
        $codeOnly = Get-CSharpCode $text
        $hits = @(Find-CodeMatches -Code $withStrings -Pattern $namesCoreButNotFacade -Rel $rel) +
            @(Find-CodeMatches -Code $codeOnly -Pattern $namesCoreRelatively -Rel $rel) +
            @(Find-CodeMatches -Code $codeOnly -Pattern $namesSiblingRelatively -Rel $rel)
        foreach ($h in $hits) { if ($h) { $q164Problems += "(c) the facade names Core outside itself - $h" } }
        foreach ($token in $inProcessForbidden.Keys) {
            foreach ($h in @(Find-CodeMatches -Code $withStrings -Pattern ('\b' + $token) -Rel $rel -IgnoreCase)) {
                if ($h) { $q164Problems += "(c) $h $($inProcessForbidden[$token]), which must never run inside Outlook's process" }
            }
        }
    }

    if ($q164Problems.Count -gt 0) {
        Fail $q164 "$(@($q164Problems | Sort-Object -Unique) -join ' | ') The add-in runs inside Outlook's process, where a hang or a crash in Core's COM, MAPI or index code takes Outlook down; see the comment above Core's ProjectReference in OutlookAI.csproj."
    } elseif ($facadeFiles.Count -eq 0) {
        Pass $q164 "the add-in's $($addinFiles.Count) files name no part of Core, and $facadeDirRel/ holds no file yet"
    } else {
        Pass $q164 "the add-in's $($addinFiles.Count) files name Core only as $facadeNamespace; its $($facadeFiles.Count) file(s) name no other part of Core ($($siblings -join ', ')) and call nothing that attaches to Outlook, starts a COM thread or opens the index"
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
