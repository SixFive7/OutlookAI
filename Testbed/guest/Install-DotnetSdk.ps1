#Requires -Version 5.1
<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it: the machine it was written on is the maintainer's
    workstation, which already has the SDK this installs, and no guest was reachable from where
    it was written. Verified by PARSING only - the same check
    .github/scripts/check-testbed-references.ps1 applies to every script under Testbed/. Nothing
    below has run anywhere: no installer was downloaded, none was executed, no SDK was installed
    and no `dotnet test` was driven to write it.

    That class of script has just cost this project five failed builds, so the logic below is
    deliberately dull. Every mechanism it uses is one this testbed has already proved somewhere
    else - staged, hash-pinned media (guest/Install-MailSink.ps1), a quiet installer with no UI
    (the Office Deployment Tool, Testbed/MEDIA.md), the two-axis guest guard
    (guest/Set-OutlookIndexingDisabled.ps1), and a -Verify that asks the real thing rather than
    reading back what it wrote. Nothing here is clever, and that is the point.

    Replace this banner with what it actually did once it has run on a guest, and say which of
    the four verdicts it printed.

.SYNOPSIS
    Installs the .NET SDK on a testbed guest from staged media, and then PROVES `dotnet test`
    can run the suite rather than reporting that an installer exited zero.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    NEVER run this on the maintainer's workstation. It refuses there; see THE GUARD below.

    WHY THIS EXISTS. The live tier is an xUnit suite invoked by `dotnet test`, and
    Testbed/README.md question 13 records that nothing has ever run it on a guest, "which the
    guest cannot run, having no SDK". That single missing precondition is the blocker under
    roughly twenty TODO items that need a live Outlook profile. The guests are built from the
    answer file and get Windows and Office and nothing else -
    Testbed/host/Publish-GuestPayload.ps1 publishes `--self-contained` precisely because the
    first corpus build on a fresh guest died with "You must install .NET to run this
    application".

    WHY AN SDK ON THE GUEST, AND NOT A RUNNER BUILT ON THE HOST. The live tier's safety
    machinery - the per-store count tripwire, the StoreWriteAllowlist, the signature-directory
    snapshot, the zero-artifact sweep - lives in xUnit FIXTURES, not in the tools. Running the
    tier by any route that is not `dotnet test` would run it with its guards absent, on a real
    mailbox. That is worse than not running it at all, and it is why the answer is an SDK rather
    than a cleverer launcher.

    WHAT AN SDK ALONE DOES NOT BUY, WHICH IS THE PART THAT WILL BITE FIRST. `dotnet test` needs
    three things on the guest and the SDK is only one of them:

      1. THE SDK.               This script, from staged media.
      2. THE SOURCE.            The suite is built from source; there is no prebuilt test
                                assembly. Testbed/host/Publish-LiveTierPayload.ps1 stages it.
      3. THE NuGet PACKAGES.    54 packages, ~76 MB, and THE GUEST HAS NO NETWORK. A restore
                                that reaches nuget.org is a restore that fails. The same host
                                script stages them as a folder feed, and -Execute writes the
                                NuGet.config that makes it the ONLY source.

    A machine with 1 but not 2 and 3 gets the SDK-ONLY verdict below, which is a real state and
    is reported as one rather than as a success.

    WHY THE INSTALLER AND NOT THE ZIP. Because it is the mechanism this machine already uses:
    staged media plus a quiet, UI-less installer is exactly how Office gets onto a guest
    (`setup.exe /configure Testbed.xml`, with `Display Level="None"`). The .exe also puts dotnet
    on the MACHINE PATH, which matters more here than it looks - the live tier runs through
    guest/Register-InteractiveTask.ps1, and a scheduled task gets a fresh environment block
    built from the registry, so a machine-wide PATH is what makes `dotnet` resolvable in session
    1 without a logon. An extracted zip would need DOTNET_ROOT and a PATH edit of our own, which
    is more moving parts for no gain.

    WHICH VERSION, AND WHY IT IS PINNED. .NET 10, x64. Every project in McpServer/ targets
    `net10.0-windows`; OutlookAI.Core additionally targets `net48`, which needs no separate
    install because it builds against the `Microsoft.NETFramework.ReferenceAssemblies` package
    (in the staged feed). `.github/workflows/mcpserver.yml` asks setup-dotnet for `10.0.x` and
    there is NO global.json anywhere in the repository, so nothing pins a feature band and any
    .NET 10 SDK would compile. The default below pins the version the HOST runs anyway, because
    the host is what publishes the payload the guest measures with, and one toolchain across
    both machines is one fewer difference to suspect when a guest behaves unlike the host - the
    same reasoning Testbed/MEDIA.md applies to locale and to the Office build gap.

    x64 IS NOT OPTIONAL. Both the server and the test project set PlatformTarget x64, and the
    index tier reads the Search.CollatorDSO OLE DB provider, which has no 32-bit story in this
    arrangement.

    WHAT -Verify ACTUALLY PROVES, AND WHY IT IS NOT `dotnet --info`. `dotnet --info` proves an
    installer exited zero. The thing being enabled is `dotnet test`, so -Verify drives it, in
    four rungs that get progressively closer to the real question:

      1. THE SDK ANSWERS.      `--list-sdks` reports a 10.x SDK, and `--info` reports a win-x64
                               RID. Anything else is the wrong installer.
      2. IT BUILDS OFFLINE.    A throwaway `net10.0-windows` console project is created,
                               restored, built and RUN. It has no PackageReference, so it needs
                               no feed at all - which separates "the SDK is broken" from "the
                               offline feed is missing", two failures that otherwise look
                               identical on a machine with no network.
      3. THE SUITE ENUMERATES. `dotnet test --list-tests` against the real test project. This
                               restores from the staged feed, builds every referenced project,
                               launches the test host and asks xUnit what it has. If that
                               prints tests, the whole chain works.
      4. A TEST ACTUALLY RUNS. One named T1 class is executed. Rung 3 proves discovery; only
                               execution proves the test host can load and run the assembly.

    MAILBOX SAFETY - READ THIS BEFORE CHANGING THE FILTER. Rung 4 runs tests on a machine that
    may have a real Outlook profile on it, so the selection is guarded twice. The filter is
    ALWAYS `Category!=Live` ANDed with a named class, the script REFUSES if -SmokeTestFilter is
    empty or itself mentions Live, and the default class - T1/CorpusTagSeparationTests - is pure
    in-memory string work with no COM, no store and no filesystem. Rung 3 is discovery only and
    executes nothing; checked on 2026-09-17, the suite contains NO MemberData and NO ClassData
    anywhere, only compile-time InlineData, so xUnit's discovery-time pre-enumeration of theory
    data cannot reach a data provider that touches Outlook. Nothing in this script starts
    Outlook, creates a COM object, touches MAPI, reads or writes a profile, or goes near a mail
    item.

    AND IT ANSWERS README QUESTION 12 BY CONSTRUCTION, which is worth knowing before anyone
    tries to solve that separately. The T3 harness finds the server through
    `AssemblyMetadata("McpServerExePath")`, an absolute path baked in AT BUILD TIME pointing
    into the repository's own bin tree; question 12 is open only because the guest could not
    build, so nothing ever put a binary there. A guest that builds the suite bakes a path into
    ITS OWN tree, where its own build has just put the exe. The staged
    `C:\OutlookAI-Q5\server\` payload stays what it is - the measurement driver's server - and
    the two stop needing to be the same path.

    IDEMPOTENT. Run it twice and the second run reports the SDK as already present, skips the
    installer and still runs the whole verification. There is no -Uninstall: removing an SDK is
    a Windows uninstall, the guest is disposable, and a checkpoint is the supported way back.

    STARTING STATE IT EXPECTS.
      * A guest checkpoint where Windows and Office are installed and you are logged on as
        -ExpectedUser.
      * An ELEVATED session. A machine-wide install and a machine-wide environment variable are
        both HKLM; the script asserts elevation rather than failing halfway with an
        access-denied nobody can interpret.
      * The SDK installer STAGED at -InstallerPath. It is not downloaded here, it is not in this
        repository, and the guest has no network. See Testbed/MEDIA.md.
      * Its published SHA-512 in -ExpectedSha512. There is no default and there must not be one
        - see THE HASH IS MANDATORY below.
      * For rungs 3 and 4: the source at -SourceRoot and the package feed at -OfflineFeed, both
        from Testbed/host/Publish-LiveTierPayload.ps1. Without them the run stops at SDK-ONLY
        and says so.
      * ~2.5 GB free on C:. The SDK is ~900 MB installed, the extracted package cache ~300 MB,
        and a Release build of the suite several hundred more.

    THE HASH IS MANDATORY, AND THAT IS DELIBERATE. -Execute refuses without -ExpectedSha512.
    guest/Install-MailSink.ps1 carries a default hash taken from a published manifest; this
    script carries none, because the agent that wrote it could not stage or download the file
    and a hash nobody has compared against anything is worse than an empty field - it reads as
    verified. Microsoft publishes the SHA-512 for every installer on the .NET download page and
    in its release metadata; Testbed/host/Publish-LiveTierPayload.ps1 prints the hash of the
    file you staged so the two can be compared in one line. Record it in Testbed/MEDIA.md once,
    and this parameter stops being work.

    THE GUARD, AND IT HAS TWO AXES ON PURPOSE. Same rule and same fail-closed shape as
    Assert-TestbedGuest in Testbed/guest/OutlookMapiInterop.ps1 - restated here rather than
    dot-sourced, for the reason Install-MailSink.ps1 gives: that file compiles Extended MAPI
    interop with Add-Type at dot-source time, and a script that has nothing to do with MAPI must
    not be able to die with a MAPI compile error. The second axis is the computer name, because
    what this script does to a machine is install software system-wide and edit a machine
    environment variable, and a workstation is exactly where that must never happen by accident.
    Both axes must pass, and the only way past either is to name the value you mean.

.PARAMETER InstallerPath
    The staged SDK installer on the guest. Never downloaded here.

.PARAMETER ExpectedSha512
    Microsoft's published SHA-512 for that file. REQUIRED for -Execute; there is no default and
    there must not be one. See THE HASH IS MANDATORY above.

.PARAMETER ExpectedSdkVersionPrefix
    What `dotnet --list-sdks` must report for the verification to accept the install. A prefix,
    so '10.' accepts any .NET 10 SDK and '10.0.401' pins one exactly.

.PARAMETER DotnetRoot
    Where the installer puts dotnet. Only used as the fallback when `dotnet` is not yet on this
    session's PATH - a machine PATH edit is not visible to a shell that was already running.

.PARAMETER SourceRoot
    The staged repository source on the guest. Rungs 3 and 4 need it.

.PARAMETER OfflineFeed
    The staged folder feed of .nupkg files. Rungs 3 and 4 need it, because the guest has no
    network and every package the suite uses has to come from somewhere.

.PARAMETER TestProject
    The test project, relative to -SourceRoot.

.PARAMETER SmokeTestFilter
    The class rung 4 executes. ALWAYS ANDed with Category!=Live; see MAILBOX SAFETY above.

.PARAMETER Configuration
    Release by default, matching CI and matching what a measurement run should use.

.PARAMETER ProbeDir
    Where rung 2's throwaway console project is written. Removed afterwards unless -KeepProbe.

.PARAMETER ExpectedUser
    Accounts this script is allowed to run as. The guests autologon as vmadmin. The default IS
    the guard; do not widen it.

.PARAMETER ExpectedComputerNamePrefix
    Computer-name prefix this script is allowed to run on. Testbed/host/New-AnswerFile.ps1
    derives a guest's name by replacing 'OutlookAI-' with 'OAI-'.

.PARAMETER Execute
    Install. Without it, prints what it would do and changes nothing.

.PARAMETER Verify
    Run the four rungs and report a verdict. Writes nothing but its own log, the throwaway probe
    project (which it removes) and the test project's own build output - a build is not
    optional, since a build happening at all is part of what is being proven.

.PARAMETER KeepProbe
    Leave rung 2's throwaway project behind for inspection.

.PARAMETER LogPath
    Transcript file. Everything printed also lands here.

.EXAMPLE
    .\Install-DotnetSdk.ps1
    .\Install-DotnetSdk.ps1 -ExpectedSha512 <hash> -Execute
    .\Install-DotnetSdk.ps1 -Verify
#>
[CmdletBinding()]
param(
    [string]   $InstallerPath             = 'C:\OutlookAI-Q5\media\dotnet-sdk-10.0.401-win-x64.exe',
    [string]   $ExpectedSha512,
    [string]   $ExpectedSdkVersionPrefix  = '10.',
    [string]   $DotnetRoot                = 'C:\Program Files\dotnet',
    [string]   $SourceRoot                = 'C:\OutlookAI-Q5\src',
    [string]   $OfflineFeed               = 'C:\OutlookAI-Q5\nuget-offline',
    [string]   $TestProject               = 'McpServer\OutlookAI.McpServer.Tests\OutlookAI.McpServer.Tests.csproj',
    [string]   $SmokeTestFilter           = 'CorpusTagSeparationTests',
    [ValidateSet('Release', 'Debug')] [string] $Configuration = 'Release',
    [string]   $ProbeDir                  = 'C:\OutlookAI-Q5\sdk-probe',
    [string[]] $ExpectedUser              = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [int]      $InstallTimeoutMinutes     = 30,
    [int]      $BuildTimeoutMinutes       = 30,
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $KeepProbe,
    [string]   $LogPath                   = 'C:\OutlookAI-Q5\install-dotnet-sdk.log'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Everything this script writes or runs, named in one place, so the blast radius is readable
# without reading the code.
# ---------------------------------------------------------------------------------------------

# The installer's own switches. /quiet is the Office-install equivalent of Display Level="None":
# an installer UI on an unattended guest is a hang, not a prompt. /norestart because a reboot
# mid-build is decided by a human, not by a bundle.
$InstallerArguments = @('/install', '/quiet', '/norestart')

# WiX bundle exit codes. 3010 is a SUCCESS that wants a reboot; treating it as a failure would
# fail a perfectly good install, and treating it as an ordinary success would hide the reboot.
$ExitOk = 0
$ExitRebootRequired = 3010

# Machine-wide environment. Both are documented .NET CLI variables and both matter on a machine
# with no route to the outside: the first stops the CLI trying to post telemetry it can never
# deliver, the second keeps the first-run banner out of captured output that gets parsed.
$MachineEnvironment = @(
    @{ Name = 'DOTNET_CLI_TELEMETRY_OPTOUT'; Value = '1'; Why = 'No telemetry. The guest has no network; an attempt is a delay, not a send.' }
    @{ Name = 'DOTNET_NOLOGO';               Value = '1'; Why = 'No first-run banner in output this script parses.' }
)

# The NuGet configuration -Execute writes at the source root. <clear/> is the load-bearing line:
# without it, an inherited nuget.org source turns every restore on an offline machine into a
# timeout, and the error names a network fault rather than a missing package.
$NuGetConfigName = 'NuGet.config'

# The verdicts. Four, because "installed" and "the tier can run" are different states and so are
# the two ways of not getting there.
$VerdictSdkAbsent = 'SDK-ABSENT'
$VerdictSdkOnly   = 'SDK-ONLY'
$VerdictTestReady = 'TEST-READY'
$VerdictBroken    = 'BROKEN'

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    if ($LogPath) {
        try {
            $dir = Split-Path -Parent $LogPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }
        catch {
            # A log that cannot be written must not stop the work it was recording.
        }
    }
}

# ---------------------------------------------------------------------------------------------
# THE GUARD. FIRST CALL IN EVERY PATH, -Verify included - a verification builds a Release tree
# inside the source root and executes test code, and neither belongs on a workstation either.
# ---------------------------------------------------------------------------------------------
function Assert-TestbedGuestLocal {
    $who = $env:USERNAME
    $userOk = $false
    foreach ($candidate in $ExpectedUser) {
        if ($who -eq $candidate) { $userOk = $true }
    }

    $machine = $env:COMPUTERNAME
    $machineOk = $false
    if ($ExpectedComputerNamePrefix -and $machine -and
        $machine.StartsWith($ExpectedComputerNamePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        $machineOk = $true
    }

    if ($userOk -and $machineOk) { return }

    throw @"
REFUSING TO RUN.

  logged on as : '$who'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$machine'      (must start with: '$ExpectedComputerNamePrefix')

This script installs a .NET SDK system-wide, edits machine environment variables, and then
builds and executes test code out of a staged source tree. None of that belongs on a machine
with real mail on it, and an SDK arriving on a workstation by accident is the kind of change
nobody attributes correctly weeks later.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2), and
Testbed/host/New-AnswerFile.ps1 derives their computer name by replacing 'OutlookAI-' with
'OAI-', so a guest built from the answer file matches both axes. If you named a guest something
else, say so:

    -ExpectedUser <username> -ExpectedComputerNamePrefix <prefix>

Do not 'fix' this by widening either default. The defaults are the guard.
"@
}

function Assert-Elevated {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw @"
REFUSING TO RUN: this session is not elevated.

A machine-wide SDK install writes under Program Files and the machine environment block. An
unelevated run would fail somewhere in the middle, having installed some of it, which is the one
outcome worse than not running at all.
"@
    }
}

# The filter guard. Rung 4 executes tests on a machine that may carry a real Outlook profile,
# so the Category!=Live half is added HERE and cannot be removed by a parameter.
function Get-SmokeFilter {
    if ([string]::IsNullOrWhiteSpace($SmokeTestFilter)) {
        throw 'REFUSING TO RUN: -SmokeTestFilter is empty. An empty filter would run the whole suite, live tier included, against whatever mailbox this guest is pointed at.'
    }
    if ($SmokeTestFilter -match '(?i)live') {
        throw "REFUSING TO RUN: -SmokeTestFilter is '$SmokeTestFilter', which names Live. Rung 4 is a smoke test, not a tier run; the live tier runs through the runbook, with its fixtures, never from here."
    }
    return "Category!=Live&FullyQualifiedName~$SmokeTestFilter"
}

# ---------------------------------------------------------------------------------------------
# Running dotnet. Output goes to a FILE and is read back, never streamed through the call: a
# build that spawns its own children can hold the pipe open long after it has finished, and the
# declared timeout then never fires. MSBuild node reuse is off for the same family of reason -
# a guest that gets checkpointed should not carry idle MSBuild workers into the checkpoint.
# ---------------------------------------------------------------------------------------------
function Invoke-Dotnet {
    param(
        [Parameter(Mandatory = $true)] [string]   $DotnetPath,
        [Parameter(Mandatory = $true)] [string[]] $Arguments,
        [Parameter(Mandatory = $true)] [string]   $LogName,
        [int] $TimeoutMinutes = 30,
        [string] $WorkingDirectory
    )

    $logDir = Split-Path -Parent $LogPath
    if (-not $logDir) { $logDir = '.' }
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

    $outFile = Join-Path $logDir ("dotnet-$LogName.out.txt")
    $errFile = Join-Path $logDir ("dotnet-$LogName.err.txt")
    foreach ($f in @($outFile, $errFile)) {
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
    }

    $env:MSBUILDDISABLENODEREUSE = '1'
    $env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
    $env:DOTNET_NOLOGO = '1'

    $startArgs = @{
        FilePath               = $DotnetPath
        ArgumentList           = $Arguments
        RedirectStandardOutput = $outFile
        RedirectStandardError  = $errFile
        NoNewWindow            = $true
        PassThru               = $true
    }
    if ($WorkingDirectory) { $startArgs['WorkingDirectory'] = $WorkingDirectory }

    Say ("    dotnet " + ($Arguments -join ' '))
    $process = Start-Process @startArgs
    $exited = $process.WaitForExit($TimeoutMinutes * 60 * 1000)

    if (-not $exited) {
        try { $process.Kill() } catch { }
        return [pscustomobject]@{
            ExitCode = -1
            TimedOut = $true
            Output   = (Get-DotnetLogText $outFile $errFile)
            OutFile  = $outFile
        }
    }

    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        TimedOut = $false
        Output   = (Get-DotnetLogText $outFile $errFile)
        OutFile  = $outFile
    }
}

function Get-DotnetLogText([string] $outFile, [string] $errFile) {
    $text = ''
    foreach ($f in @($outFile, $errFile)) {
        if (Test-Path -LiteralPath $f) {
            $chunk = Get-Content -LiteralPath $f -Raw -ErrorAction SilentlyContinue
            if ($chunk) { $text = $text + $chunk + "`n" }
        }
    }
    return $text
}

function Show-Tail([string] $text, [int] $lines = 25) {
    if (-not $text) { return }
    $all = $text -split "`r?`n"
    $start = [Math]::Max(0, $all.Count - $lines)
    for ($i = $start; $i -lt $all.Count; $i++) {
        if ($all[$i]) { Say ("      | " + $all[$i]) }
    }
}

# Resolve dotnet HONESTLY: PATH first, because that is what every other caller on this guest
# will get, and only then the install directory. Which one answered is reported, because "only
# the absolute path works" is a real and recoverable state - a machine PATH edit is invisible to
# a shell that was already running - and it is not the same as a healthy machine.
function Resolve-Dotnet {
    $onPath = Get-Command dotnet -CommandType Application -ErrorAction SilentlyContinue
    if ($onPath) {
        return [pscustomobject]@{ Path = $onPath.Source; From = 'PATH' }
    }
    $absolute = Join-Path $DotnetRoot 'dotnet.exe'
    if (Test-Path -LiteralPath $absolute) {
        return [pscustomobject]@{ Path = $absolute; From = 'DotnetRoot' }
    }
    return $null
}

function Update-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    $parts = @()
    foreach ($p in @($machine, $user)) {
        if ($p) { $parts += ($p -split ';' | Where-Object { $_ }) }
    }
    $env:Path = ($parts -join ';')
}

# ---------------------------------------------------------------------------------------------
# Dry run.
# ---------------------------------------------------------------------------------------------
if (-not ($Execute -or $Verify)) {
    Assert-TestbedGuestLocal
    Say 'DRY RUN. Would:'
    Say "    check   $InstallerPath against a SHA-512 supplied with -ExpectedSha512"
    Say ("    run     " + (Split-Path -Leaf $InstallerPath) + ' ' + ($InstallerArguments -join ' '))
    foreach ($e in $MachineEnvironment) { Say ("    set     {0}={1}   # {2}" -f $e.Name, $e.Value, $e.Why) }
    Say "    write   $(Join-Path $SourceRoot $NuGetConfigName)  pointing at $OfflineFeed and clearing every other source"
    Say '    verify  four rungs: the SDK answers, it builds offline, the suite enumerates, one T1 class runs'
    Say ''
    Say 'Re-run with -ExpectedSha512 <hash> -Execute to install, or -Verify to test what is already here.'
    return
}

Assert-TestbedGuestLocal

# ---------------------------------------------------------------------------------------------
# Execute.
# ---------------------------------------------------------------------------------------------
if ($Execute) {
    Assert-Elevated
    Say '== Execute =='

    $existing = Resolve-Dotnet
    $alreadyInstalled = $false
    if ($existing) {
        $probe = Invoke-Dotnet -DotnetPath $existing.Path -Arguments @('--list-sdks') -LogName 'list-sdks-pre' -TimeoutMinutes 5
        if ($probe.ExitCode -eq 0 -and $probe.Output -match ('(?m)^' + [regex]::Escape($ExpectedSdkVersionPrefix))) {
            $alreadyInstalled = $true
            Say "  already an SDK matching '$ExpectedSdkVersionPrefix' at $($existing.Path) - skipping the installer"
        }
    }

    if (-not $alreadyInstalled) {
        if (-not (Test-Path -LiteralPath $InstallerPath)) {
            throw @"
REFUSING TO RUN: no installer at $InstallerPath.

It is STAGED MEDIA, not an artefact - nothing in this repository produces it, the guest has no
network, and it must never be downloaded here. Testbed/MEDIA.md declares it; stage it on the
host, copy it in with Testbed/host/Copy-ToGuest.ps1, and pass -InstallerPath if you put it
somewhere else.
"@
        }

        if ([string]::IsNullOrWhiteSpace($ExpectedSha512)) {
            throw @"
REFUSING TO RUN: -ExpectedSha512 was not supplied.

This script is about to run an installer with SYSTEM rights on a machine that is rebuilt from
committed scripts, and there is deliberately no default hash to fall back on - see THE HASH IS
MANDATORY in the banner. Microsoft publishes a SHA-512 for every .NET installer on its download
page and in its release metadata; Testbed/host/Publish-LiveTierPayload.ps1 prints the hash of
the file you staged so the two can be compared in one line. Record it in Testbed/MEDIA.md.
"@
        }

        Say "  hashing $InstallerPath"
        $actual = (Get-FileHash -LiteralPath $InstallerPath -Algorithm SHA512).Hash
        $wanted = $ExpectedSha512.Trim().Replace('-', '').ToUpperInvariant()
        if ($actual -ne $wanted) {
            throw @"
REFUSING TO RUN: the staged installer is not the expected file.

  expected  $wanted
  actual    $actual
  file      $InstallerPath

A mismatch means "check the hash", not necessarily "the file is bad" - a newer patch release has
a different hash and a different name. Confirm which version you staged, take its published
SHA-512, and update Testbed/MEDIA.md so the next rebuilder is not doing this again.
"@
        }
        Say "  hash matches: $actual"

        $installLog = Join-Path (Split-Path -Parent $LogPath) 'dotnet-sdk-install.log'
        $arguments = $InstallerArguments + @('/log', $installLog)
        Say ("  running " + (Split-Path -Leaf $InstallerPath) + ' ' + ($arguments -join ' '))
        Say "  installer log: $installLog"

        # NEVER -Verb RunAs. UAC elevation takes the foreground even with a hidden window, which
        # this project has measured and banned outright (guest/Register-InteractiveTask.ps1).
        # The session is already elevated, asserted above, so the bundle needs no elevation of
        # its own.
        $installer = Start-Process -FilePath $InstallerPath -ArgumentList $arguments -PassThru -NoNewWindow
        $exited = $installer.WaitForExit($InstallTimeoutMinutes * 60 * 1000)
        if (-not $exited) {
            try { $installer.Kill() } catch { }
            throw "The installer did not finish within $InstallTimeoutMinutes minutes. Its own log is at $installLog; read that before re-running."
        }

        $code = $installer.ExitCode
        if ($code -eq $ExitOk) {
            Say '  installer exited 0'
        }
        elseif ($code -eq $ExitRebootRequired) {
            Say "  installer exited $ExitRebootRequired - SUCCESS, but Windows wants a reboot."
            Say '  THIS IS NOT A FAILURE. Finish the verification, then reboot the guest before'
            Say '  taking a checkpoint, so the checkpoint is of a settled machine.'
        }
        else {
            throw "The installer exited $code. Its own log is at $installLog; the failing line is usually near the end of it."
        }

        # -Wait on a bundle can return before the work is on disk. Poll for the artefact rather
        # than trusting the exit code on its own.
        $deadline = (Get-Date).AddMinutes(5)
        $dotnetExe = Join-Path $DotnetRoot 'dotnet.exe'
        while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $dotnetExe)) {
            Start-Sleep -Seconds 5
        }
        if (-not (Test-Path -LiteralPath $dotnetExe)) {
            throw "The installer reported $code but $dotnetExe does not exist. Read $installLog."
        }
        Say "  $dotnetExe is present"
    }

    foreach ($e in $MachineEnvironment) {
        [Environment]::SetEnvironmentVariable($e.Name, $e.Value, 'Machine')
        Set-Item -Path ("Env:\" + $e.Name) -Value $e.Value
        Say ("  set {0}={1}" -f $e.Name, $e.Value)
    }

    # The machine PATH the installer just wrote is not in THIS session's environment. Rebuild it
    # so the verification below resolves dotnet the same way a fresh session will.
    Update-SessionPath

    # The offline NuGet configuration. Written at the source root so it covers every project
    # under it, and only there - nothing machine-wide, so a guest that later gets a network is
    # not left with a crippled NuGet.
    if (Test-Path -LiteralPath $SourceRoot) {
        $configPath = Join-Path $SourceRoot $NuGetConfigName
        $feedForConfig = $OfflineFeed
        $configXml = @"
<?xml version="1.0" encoding="utf-8"?>
<!--
  WRITTEN BY Testbed/guest/Install-DotnetSdk.ps1. Not part of the repository - the host-side
  checkout has no NuGet.config and must not acquire one.

  THE GUEST HAS NO NETWORK. <clear/> is the load-bearing element: without it an inherited
  nuget.org source turns every restore into a timeout whose error names a network fault rather
  than a missing package, which is the wrong thing to go looking for. With it, a package that
  was not staged fails immediately and says which one.
-->
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
        Set-Content -LiteralPath $configPath -Value $configXml -Encoding UTF8
        Say "  wrote $configPath -> $feedForConfig"
    }
    else {
        Say "  NOTE $SourceRoot does not exist, so no NuGet.config was written."
        Say '       The SDK is installed; the suite cannot be built until the source and the feed'
        Say '       are staged by Testbed/host/Publish-LiveTierPayload.ps1. Re-run -Execute after'
        Say '       copying them in, or write the config by hand.'
    }
}

# ---------------------------------------------------------------------------------------------
# Verify. Four rungs, one verdict.
# ---------------------------------------------------------------------------------------------
Say ''
Say '== Verify =='

$verdict = $VerdictBroken
$notes = @()

$dotnet = Resolve-Dotnet
if (-not $dotnet) {
    Say "  FAIL no dotnet on PATH and none at $(Join-Path $DotnetRoot 'dotnet.exe')"
    $verdict = $VerdictSdkAbsent
    $notes += 'Nothing is installed. Run with -ExpectedSha512 <hash> -Execute.'
}
else {
    Say "  dotnet: $($dotnet.Path)   (found via $($dotnet.From))"
    if ($dotnet.From -eq 'DotnetRoot') {
        Say '  NOTE dotnet is NOT on this session PATH. That is expected immediately after an'
        Say '       install - a machine PATH edit is invisible to a shell that was already'
        Say '       running - and it clears on the next logon or the next scheduled task, which'
        Say '       gets a fresh environment block. If it persists into a NEW session, the'
        Say '       installer did not edit the machine PATH and nothing else here will notice.'
        $notes += 'dotnet resolved only by absolute path; re-run -Verify from a new session to confirm PATH.'
    }

    # --- rung 1: the SDK answers -------------------------------------------------------------
    $rung1 = $false
    $listSdks = Invoke-Dotnet -DotnetPath $dotnet.Path -Arguments @('--list-sdks') -LogName 'list-sdks' -TimeoutMinutes 5
    if ($listSdks.ExitCode -ne 0) {
        Say "  FAIL rung 1: dotnet --list-sdks exited $($listSdks.ExitCode)"
        Show-Tail $listSdks.Output
    }
    elseif ($listSdks.Output -notmatch ('(?m)^' + [regex]::Escape($ExpectedSdkVersionPrefix))) {
        Say "  FAIL rung 1: no SDK matching '$ExpectedSdkVersionPrefix'. Installed:"
        Show-Tail $listSdks.Output
        $notes += "The wrong SDK is installed. The suite targets net10.0-windows and will not build without a .NET 10 SDK."
    }
    else {
        $info = Invoke-Dotnet -DotnetPath $dotnet.Path -Arguments @('--info') -LogName 'info' -TimeoutMinutes 5
        $ridOk = ($info.ExitCode -eq 0 -and $info.Output -match '(?i)win-x64')
        if (-not $ridOk) {
            Say '  FAIL rung 1: --info does not report a win-x64 RID.'
            Show-Tail $info.Output
            $notes += 'x64 is not optional: both projects set PlatformTarget x64 and the index tier needs an x64 host for Search.CollatorDSO.'
        }
        else {
            foreach ($line in ($listSdks.Output -split "`r?`n")) {
                if ($line.Trim()) { Say ("  rung 1 sdk: " + $line.Trim()) }
            }
            Say '  OK   rung 1: a matching x64 SDK answers.'
            $rung1 = $true
        }
    }

    # --- rung 2: it builds and runs offline, with no feed at all ------------------------------
    $rung2 = $false
    if ($rung1) {
        if (Test-Path -LiteralPath $ProbeDir) { Remove-Item -LiteralPath $ProbeDir -Recurse -Force }
        New-Item -ItemType Directory -Path $ProbeDir -Force | Out-Null

        $sentinel = 'OUTLOOKAI-SDK-PROBE-OK'
        $probeProj = Join-Path $ProbeDir 'SdkProbe.csproj'
        $probeMain = Join-Path $ProbeDir 'Program.cs'

        # No PackageReference, deliberately. Everything this needs ships inside the SDK, so it
        # proves the toolchain WITHOUT the staged feed - which is what separates a broken SDK
        # from a missing feed on a machine that cannot reach nuget.org either way.
        Set-Content -LiteralPath $probeProj -Encoding UTF8 -Value @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0-windows</TargetFramework>
    <PlatformTarget>x64</PlatformTarget>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
</Project>
"@
        Set-Content -LiteralPath $probeMain -Encoding UTF8 -Value @"
System.Console.WriteLine("$sentinel " + System.Environment.Version + " " + (System.Environment.Is64BitProcess ? "x64" : "x86"));
"@

        $run = Invoke-Dotnet -DotnetPath $dotnet.Path -Arguments @('run', '--project', $probeProj, '-c', $Configuration) -LogName 'probe-run' -TimeoutMinutes 15 -WorkingDirectory $ProbeDir
        if ($run.ExitCode -eq 0 -and $run.Output -match [regex]::Escape($sentinel)) {
            $sentinelLine = ''
            foreach ($line in ($run.Output -split "`r?`n")) {
                if ($line -match [regex]::Escape($sentinel)) {
                    $sentinelLine = $line.Trim()
                    Say ("  rung 2 says: " + $sentinelLine)
                }
            }
            # Read the bitness off the sentinel line only. Build output mentions x86 for its own
            # reasons, and matching the whole transcript would fail a healthy machine.
            if ($sentinelLine -match 'x86') {
                Say '  FAIL rung 2: the probe ran as a 32-bit process.'
                $notes += 'A 32-bit host cannot open Search.CollatorDSO; the index tier would report a missing index on a healthy machine.'
            }
            else {
                Say '  OK   rung 2: the SDK restores, builds and runs net10.0-windows offline.'
                $rung2 = $true
            }
        }
        else {
            Say "  FAIL rung 2: the throwaway console project did not build and run (exit $($run.ExitCode), timed out: $($run.TimedOut))."
            Show-Tail $run.Output
            $notes += 'The SDK itself is not working. Nothing about the staged feed or the source can explain this one.'
        }

        if (-not $KeepProbe) {
            Remove-Item -LiteralPath $ProbeDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        else {
            Say "  probe kept at $ProbeDir"
        }
    }

    # --- rungs 3 and 4: the real suite --------------------------------------------------------
    $projectFull = Join-Path $SourceRoot $TestProject
    $configFull = Join-Path $SourceRoot $NuGetConfigName
    $missing = @()
    if (-not (Test-Path -LiteralPath $SourceRoot))   { $missing += "source at $SourceRoot" }
    elseif (-not (Test-Path -LiteralPath $projectFull)) { $missing += "test project at $projectFull" }
    if (-not (Test-Path -LiteralPath $OfflineFeed))  { $missing += "package feed at $OfflineFeed" }
    if (-not (Test-Path -LiteralPath $configFull))   { $missing += "$NuGetConfigName at $configFull (written by -Execute)" }

    if ($rung2 -and $missing.Count -eq 0) {
        $listTests = Invoke-Dotnet -DotnetPath $dotnet.Path `
            -Arguments @('test', $projectFull, '-c', $Configuration, '--list-tests') `
            -LogName 'list-tests' -TimeoutMinutes $BuildTimeoutMinutes -WorkingDirectory $SourceRoot

        # Discovery only. It executes nothing - and the suite has no MemberData or ClassData
        # anywhere, so there is no discovery-time data provider that could reach Outlook.
        $discovered = 0
        foreach ($line in ($listTests.Output -split "`r?`n")) {
            if ($line -match '^\s{2,}\S+\.\S+') { $discovered++ }
        }

        if ($listTests.ExitCode -ne 0 -or $discovered -le 0) {
            Say "  FAIL rung 3: dotnet test --list-tests exited $($listTests.ExitCode) and enumerated $discovered test(s)."
            Show-Tail $listTests.Output 40
            if ($listTests.Output -match '(?i)unable to load the service index|NU1301|no such host is known|nuget\.org') {
                $notes += 'The restore tried to reach the network. Check that NuGet.config really cleared the inherited sources and that every package the suite needs is in the staged feed.'
            }
            $verdict = $VerdictBroken
        }
        else {
            Say "  OK   rung 3: the suite enumerates - $discovered test(s) discovered."
            Say "       full listing: $($listTests.OutFile)"

            $filter = Get-SmokeFilter
            Say "  rung 4 filter: $filter"
            $smoke = Invoke-Dotnet -DotnetPath $dotnet.Path `
                -Arguments @('test', $projectFull, '-c', $Configuration, '--no-build', '--filter', $filter) `
                -LogName 'smoke' -TimeoutMinutes $BuildTimeoutMinutes -WorkingDirectory $SourceRoot

            $passed = 0
            $failed = 0
            $m = [regex]::Match($smoke.Output, '(?i)Failed:\s*(\d+),\s*Passed:\s*(\d+)')
            if ($m.Success) {
                $failed = [int]$m.Groups[1].Value
                $passed = [int]$m.Groups[2].Value
            }

            if ($smoke.ExitCode -eq 0 -and $passed -gt 0 -and $failed -eq 0) {
                Say "  OK   rung 4: $passed test(s) executed and passed, 0 failed."
                $verdict = $VerdictTestReady
            }
            else {
                Say "  FAIL rung 4: exit $($smoke.ExitCode), passed $passed, failed $failed."
                Show-Tail $smoke.Output 40
                $verdict = $VerdictBroken
                $notes += "Discovery worked and execution did not. That is a test-host problem, not an SDK one - read $($smoke.OutFile)."
            }
        }
    }
    elseif ($rung2) {
        Say '  SKIP rungs 3 and 4: the suite is not staged on this guest.'
        foreach ($m in $missing) { Say "       missing: $m" }
        $verdict = $VerdictSdkOnly
        $notes += 'Run Testbed/host/Publish-LiveTierPayload.ps1 on the host, copy Source.zip and NuGet.zip in with Testbed/host/Copy-ToGuest.ps1, expand them, then re-run -Execute (to write NuGet.config) and -Verify.'
    }
    elseif ($rung1) {
        $verdict = $VerdictBroken
    }
}

# ---------------------------------------------------------------------------------------------
Say ''
Say "VERDICT: $verdict"
switch ($verdict) {
    $VerdictTestReady {
        Say '  `dotnet test` builds, discovers and executes the suite on this guest.'
        Say '  The live tier itself still runs through the runbook and through'
        Say '  guest/Register-InteractiveTask.ps1: Outlook cannot start in session 0, so a tier'
        Say '  run driven over PowerShell Direct hangs rather than failing.'
    }
    $VerdictSdkOnly {
        Say '  The SDK is installed and working. The suite is not on this guest yet, so the thing'
        Say '  this script exists to enable has NOT been proven. That is a missing precondition,'
        Say '  not a fault.'
    }
    $VerdictSdkAbsent {
        Say '  No SDK. Nothing was proven and nothing was broken.'
    }
    $VerdictBroken {
        Say '  Something that should work does not. The rung that failed says which, and the'
        Say "  captured output is beside $LogPath."
    }
}
foreach ($n in $notes) { Say "  - $n" }

Say ''
Say 'REGISTRY AND EXIT CODES ARE NOT THE ANSWER HERE EITHER. The only claim this script makes is'
Say 'about the four rungs it actually ran. It says nothing about whether the live tier passes,'
Say 'whether the tier profile is correct, or whether the mailbox guards are satisfied - those are'
Say 'the runbook, and they are proven by running it.'

if ($verdict -eq $VerdictTestReady) { exit 0 }
if ($verdict -eq $VerdictSdkOnly)   { exit 2 }
if ($verdict -eq $VerdictSdkAbsent) { exit 3 }
exit 1
