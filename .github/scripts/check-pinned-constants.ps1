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
# 8. The .NET 10 runtime probe.
#    McpRegistrationService.IsDotnetRuntime10Installed's version prefix ==
#    Installer.iss IsNetRuntime10Installed's. Both walk the shared-framework directory and
#    accept a Microsoft.NETCore.App folder whose name starts with the same three characters;
#    the server's roll-forward is Minor, so EXACTLY 10.x is the answer both must give - a
#    ">= 10" reading would report a satisfied prerequisite on a machine where the server
#    cannot start.
#
#    They are a knowing mirror in two languages, and Pascal cannot read a C# constant, so this
#    is the only thing that can relate them. The failure is quiet in both directions: setup
#    skipping a runtime install the machine needed, or Settings reporting a missing runtime
#    that is right there - each one a status message rather than a crash, which is exactly why
#    a drifted pair could sit unnoticed.
#
#    The Pascal side is checked twice, because it spells the prefix as a LENGTH plus a literal
#    and the two can disagree with each other: Copy(name, 1, 3) = '10.' is only meaningful
#    while the length matches the literal it is compared against. Copy(name, 1, 2) = '10.' can
#    never be true, and Copy(name, 1, 4) = '10.' is false for every 10.x folder there is - both
#    of which would silently turn the installer's probe into "no runtime, ever".
# ---------------------------------------------------------------------------------------------
$script:Checks++
$csRuntimePrefix = Get-Pinned 'Services/McpRegistrationService.cs' `
    'name\.StartsWith\("([^"]+)",\s*StringComparison\.Ordinal\)' '.NET runtime version prefix (McpRegistrationService.cs)'
$issRuntimePrefix = Get-Pinned 'Installer.iss' `
    "Copy\(FindRec\.Name,\s*1,\s*\d+\)\s*=\s*'([^']*)'" '.NET runtime version prefix (Installer.iss)'
$issRuntimePrefixLength = Get-Pinned 'Installer.iss' `
    "Copy\(FindRec\.Name,\s*1,\s*(\d+)\)\s*=\s*'[^']*'" '.NET runtime prefix length (Installer.iss)'
if ($csRuntimePrefix -and $issRuntimePrefix -and $issRuntimePrefixLength) {
    if ($csRuntimePrefix -cne $issRuntimePrefix) {
        Fail ".NET runtime version prefix" "McpRegistrationService accepts a shared-framework folder starting '$csRuntimePrefix' but Installer.iss accepts one starting '$issRuntimePrefix'. Setup and the add-in would disagree about whether the runtime the mail server needs is installed - one of them silently, because both report the answer rather than failing on it."
    } elseif ([int]$issRuntimePrefixLength -ne $issRuntimePrefix.Length) {
        Fail ".NET runtime version prefix" "Installer.iss compares Copy(FindRec.Name, 1, $issRuntimePrefixLength) against '$issRuntimePrefix', which is $($issRuntimePrefix.Length) characters. Those cannot both be right: the comparison is either always false or matching a prefix nobody intended, so setup would decide the runtime is missing on every machine."
    } else {
        Pass ".NET runtime version prefix" "$csRuntimePrefix (Copy length $issRuntimePrefixLength)"
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
