#Requires -Version 5.1
<#
    ============================================================================================
    RUN 2026-09-24 ON THE MAINTAINER'S WORKSTATION, THREE TIMES. THE HOST WAS LEFT UNCHANGED.
    ITS PAYLOAD HAS SINCE BEEN INSTALLED ON BOTH GUESTS, AND PRINTED ADDIN-READY THERE.
    ============================================================================================

    On a guest (Docs/live-tier-on-the-vm.md): OutlookAI-Unindexed, 2026-09-24, built from fe65ced -
    Testbed/guest/Install-OutlookAIAddIn.ps1 printed NOT-INSTALLED, then ADDIN-READY twice (section
    4.1 step 4, CP-09-ADDIN-READY); OutlookAI-Indexed from 98e050e the same day (section 4.2 step 4,
    CP-14-ADDIN-READY) and again from af56efc on 2026-09-27 (section 4.2b step 4.4,
    CP-14C-ADDIN-READY), the same verdicts each time.

    What ran, against HEAD fb19ccf, with VS 2026 Community 18.10 (MSBuild 18.10.1, Office
    workload) and Inno Setup 6, while the maintainer's own Outlook was running with its own add-in:

      * RUN 1 REFUSED, TWICE, AND BOTH REFUSALS WERE RIGHT. The target list had been passed as
        `CopyFilesToOutputDirectory%3BVisualStudioForApplicationsBuild`, and MSBuild looked for ONE
        target of that name (MSB4057) - a list takes a real, quoted semicolon; see
        Format-MSBuildProperty. And guard 3 found the throwaway certificate's thumbprint in
        Cert:\CurrentUser\CA: New-SelfSignedCertificate puts a copy of what it makes there as well
        as in My. That copy was removed by hand and Remove-ThrowawayCertificate now sweeps every
        store. Nothing else had been written.
      * RUNS 2 AND 3: guard 1 proven by MSBuild's own evaluation before anything ran; the build
        took 4 seconds with 0 warnings and 0 errors; SignFile signed both manifests with the
        throwaway key; both stand-ins logged and RegisterOfficeAddin never appeared in the
        detailed log; the throwaway certificate was removed from My and CA; and guard 3 compared
        376 host lines - the Outlook add-in registration, every VSTO trust entry, the form-region
        keys, Software\OutlookAI and six certificate stores - IDENTICAL before and after. An
        independent snapshot taken outside the script agreed after all three runs. The installer
        compiled at 41.1 MB; AddIn.zip is 40.6 MB.
      * -SelfTest: 96 assertions, 0 failures, under Windows PowerShell 5.1 and PowerShell 7. Four
        of its rules were broken on purpose in a scratch copy - the registration target left on
        the chain, Override dropped, the %3B list spelling, the forward-slash footprint - and each
        was caught.
      * 2026-09-27, THE LAUNCH MODE Q81 FOUND IN Tools/Switch-AddInBuild.ps1. Windows PowerShell
        5.1 started by Start-Process from PowerShell 7 ran this script for 2 seconds and stopped
        on Get-FileHash, before building anything (OwnEditionModules.ps1 says why). Fixed the
        way that script was, then run end to end from 4dca745 both ways - 5.1 by Start-Process
        from 7 (21 s) and 5.1 started directly from a 7 prompt (18 s): exit 0 both times, guard 3
        UNCHANGED over 358 lines, and a reg export of the add-in registration and of the VSTO
        trust list byte-identical before and after each run. The AddIn.zip the framework's zip
        now writes expands under 5.1's Expand-Archive to the stage's two files, hash for hash.
        -SelfTest 96/0 under PowerShell 7.6.6 and under 5.1 in both launch modes.
      * 2026-09-27, later: the import moved into Testbed/host/OwnEditionModules.ps1, which every host
        script that needs either module now dot-sources.
      * 2026-09-27, later still: an -OutDir holding ',' or ';' is refused before anything runs
        (Get-MSBuildPathProblem says how each breaks MSBuild). Refused in under a second under 5.1
        from 7 and under 7, for both characters, with no directory and no certificate made; then,
        with both changes above, run end to end from 5e1e19e under 5.1 from 7 with the default
        -OutDir: exit 0, guard 3 UNCHANGED over 358 lines, the reg export byte-identical.
        -SelfTest 100/0 under 7 and under 5.1 in both launch modes.

.SYNOPSIS
    Builds the OutlookAI Outlook add-in on the HOST from a NAMED COMMIT, packages it with the
    product's own installer, and stages it for a test guest - WITHOUT registering, trusting or
    loading anything on the host.

.DESCRIPTION
    WHY THIS EXISTS. Two live tests read state that only the add-in writes, the first time it
    runs inside Outlook: T3/Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning
    (declares Requires=AddInRegistry) and T2/LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail
    (needs it too and does not declare it). Both read HKCU\Software\OutlookAI\Tuning through
    HealthReporting.ReadTuningState. The script-built guests have no add-in, so on a guest both
    fail - not skip. This is the host half of fixing that; the guest half is
    Testbed/guest/Install-OutlookAIAddIn.ps1.

    WHY A BUILD AND NOT A RELEASED INSTALLER. The tests should exercise the code being tested. The
    add-in and the mail server agree about those registry values through ONE file compiled into
    both (Services/AddInServerContract.cs), and the suite on the guest is built from a named commit
    by Testbed/host/Publish-LiveTierPayload.ps1. A released installer would pair that suite with
    whatever add-in shipped last, and a contract change made since would go untested - or pass for
    the wrong reason. So the add-in comes from a commit too, the same way the suite does: a
    `git archive` of it, never the working tree. The manifest this script writes records which
    commit, and the guest script compares it with the suite's.

    WHAT IT PRODUCES, in -OutDir:

      AddIn.zip -> C:\OutlookAI-Q5\addin\   the add-in's own installer (Installer.iss, compiled
                                            here) plus addin-payload.json, the manifest that
                                            pins every hash the guest checks.

    and it IDENTIFIES, never produces, one piece of media:

      vstor_redist.exe -> C:\OutlookAI-Q5\media\   the VSTO runtime redistributable, pinned by the
                                                   same SHA-256 and length every release
                                                   pins. It is staged by a human
                                                   (Testbed/MEDIA.md). The installer compiles it
                                                   in, as every release does, but a SILENT install
                                                   skips it - so the guest script installs it
                                                   itself, from this copy.

    ============================================================================================
    THE HAZARD THIS SCRIPT IS BUILT AROUND: A PLAIN BUILD OF THE ADD-IN REGISTERS IT ON THE HOST.
    ============================================================================================

    Read in the VSTO build targets (MSBuild\Microsoft\VisualStudio\v18.0\OfficeTools\
    Microsoft.VisualStudio.Tools.Office.targets) on 2026-09-24. Every build of a VSTO add-in
    project runs, as part of an ordinary Build:

      RegisterOfficeAddin     -> SetOffice2007AddInRegistration writes
                                 HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI with
                                 Manifest = <this build's output>\OutlookAI.vsto|vstolocal.
                                 THAT IS THE SAME KEY THE INSTALLER WRITES. On the maintainer's
                                 workstation it silently repoints the REAL Outlook at the build
                                 output, from its next start.
      SetInclusionListEntry   -> a VSTO trust entry under HKCU\Software\Microsoft\VSTO\Security\
                                 Inclusion for the build output, signed by the build's key.
      RegisterFormRegions     -> Outlook form-region registrations (none exist in this add-in, so
                                 the real task writes nothing - but it is a writer).

    The measured consequence: the maintainer's workstation carries VSTO trust entries for add-in
    builds under this repository's own agent worktrees, left by builds nobody meant to register.
    None of that is acceptable here: the hard rule is that the add-in is NEVER installed,
    registered or run on the host. So the build runs with THREE independent guards, and the third
    one is the one that counts:

      1. RegisterOfficeAddin IS NOT SCHEDULED. PrepareForRunDependsOn is passed as a GLOBAL
         property, which MSBuild does not let a project file override, set to exactly the chain
         the VSTO targets would have built (CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild)
         minus the registration target.
      2. THE THREE WRITING TASKS ARE REPLACED BY NO-OPS. A generated targets file, imported through
         the documented CustomBeforeMicrosoftCommonTargets hook, declares
         SetOffice2007AddInRegistration, SetInclusionListEntry and RegisterFormRegions as inline
         tasks with Override="true" (MSBuild 17.2+; documented: "the first UsingTask ... is used;
         to override tasks you must define a new UsingTask before the existing one, or specify
         Override="true""). Each logs a TESTBED-NO-HOST-WRITE line naming what it did NOT write,
         and the build log is checked for them afterwards - a stand-in that never ran is a real
         task that did.
      3. THE HOST IS SNAPSHOTTED BEFORE AND AFTER, AND THE SCRIPT FAILS ON ANY TRACE. The keys
         those tasks write, the add-in's own key, and six CurrentUser certificate stores. Any
         value naming this build's directory, the throwaway signing key or its thumbprint is a
         failure, and so is ANY change to the OutlookAI add-in registration - which, if it ever
         happens, is put back to its snapshotted value before the script stops.

    Nothing Microsoft ships is edited, copied or patched: the VSTO targets run unmodified. The
    manifests they produce do not depend on the three tasks (the form-region list they would have
    fed the add-in manifest is empty either way, and the script refuses a source tree that
    declares a form region, because then it would not be).

    SIGNING WITHOUT THE RELEASE KEY. VSTO requires signed manifests. A release is signed with the
    maintainer's own key (THE RELEASE BUILD, below); a testbed build signs with a THROWAWAY
    self-signed certificate made for this one build in Cert:\CurrentUser\My, and deletes it, private
    key included, before the script ends. The guest trusts that build's PUBLIC key, which the
    manifest carries; nobody can sign anything else with it afterwards, because the private half no
    longer exists.

    VERSION. 99.99.99.0 by default - what every developer build of this repository carries
    (Properties/AssemblyInfo.cs, McpServer/Directory.Build.props). It is higher than any release,
    so the add-in's own updater can never decide a release is newer and replace the build under
    test. The commit is the identity; the manifest records it.

    WHAT THE INSTALLER DOES NOT CARRY: the MCP server. The release puts it in {app}\McpServer; the
    live tier never uses that copy - it builds and spawns its own from the suite source - and the
    add-in only needs it to register itself with Claude Code, which a guest does not have.

    THE RELEASE BUILD (-ReleaseSigningThumbprint). There is no CI any more: on 2026-10-03 the
    maintainer decided that only his own machine builds and tests ("No CI pipelines on github"),
    and Tools/Publish-Release.ps1 builds every release THROUGH THIS SCRIPT, so the three guards
    above stand between a release build and his Outlook as well. The release build differs in
    exactly what the release workflow did differently, and in nothing else:
      * the manifests are signed with THAT certificate - the release key, already in
        Cert:\CurrentUser\My with its private key - and no certificate is created or deleted.
        Guard 3 then counts the key's thumbprint and public key as a trace only on lines the build
        ADDED, because the key is on the host before the build and must still be there after it;
        the build directory still counts wherever it appears;
      * Properties/AssemblyInfo.cs is stamped with -Version, and the flattened OutlookAI.dll must
        carry it - the file the registration loads;
      * the MCP server is published into publish\McpServer, framework-dependent, with the same
        version, so the installer carries it, and both of its executables must carry the version;
      * it stops at the UNSIGNED installer: no guest manifest, no AddIn.zip. Signing the installer
        and publishing it are Tools/Publish-Release.ps1's.

    -CompareInstalledTargets RUNS ONE READ-ONLY COMPARISON AND NOTHING ELSE: guard 2's stand-ins
    against the VSTO targets of every Visual Studio on this machine that has the Office workload -
    the comparison the self-test makes, and skips where there is no Visual Studio, as on the build
    VM. A Visual Studio update can change those targets, and a stand-in that no longer covers a
    parameter or a writer reopens Q81's hole without a word. So Tools/Publish-Release.ps1 runs this
    before every release, on the machine that builds it, and refuses to release unless it passes
    (D7 (c), decided by the maintainer 2026-10-03). In this mode a machine with no such Visual
    Studio FAILS rather than skips: a comparison that did not run proves nothing.

    THIS SCRIPT TAKES NO -VMName, AND THAT IS CORRECT - the same rule
    Testbed/host/Publish-GuestPayload.ps1 states. It only builds on the host; naming a guest
    happens in the next command, Testbed/host/Copy-ToGuest.ps1, where -VMName is mandatory.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER OutDir
    Where everything lands. Default: .work/testbed-addin-payload under the repository root, which
    is gitignored. Refused anywhere inside a git working tree except under .work/, and refused when
    the full path contains ',' or ';', which MSBuild's command line splits on (see
    Get-MSBuildPathProblem) - before anything is built.

.PARAMETER Ref
    The commit to build. HEAD by default. Uncommitted edits are NOT in the build, on purpose.

.PARAMETER VstoRuntimePath
    The staged VSTO runtime redistributable. Default: .work/media/vstor_redist.exe, the place
    Testbed/MEDIA.md declares. Identified by hash, copied into the build tree for the installer
    compile, never downloaded.

.PARAMETER MSBuildPath
    MSBuild.exe to use. Default: found with vswhere, from a Visual Studio that has the Office
    workload - the only kind that carries the VSTO build targets.

.PARAMETER IsccPath
    Inno Setup's ISCC.exe. Default: PATH, then the per-user and per-machine Inno Setup 6 folders.

.PARAMETER Version
    Four-part version stamped on the ClickOnce manifests and the installer. See VERSION above.

.PARAMETER SelfTest
    Run the pure decisions against synthetic inputs and exit. No build, no git, no certificate, no
    registry read. Safe anywhere.

.PARAMETER ReleaseSigningThumbprint
    FOR Tools/Publish-Release.ps1. Build the RELEASE (see THE RELEASE BUILD above): sign with this
    certificate from Cert:\CurrentUser\My, stamp -Version into the add-in and the MCP server, put
    the server in the installer, and stop at the unsigned installer. -Version must then be the
    release's own, never 99.99.99.0.

.PARAMETER CompareInstalledTargets
    Compare guard 2's stand-ins with the installed VSTO targets and exit: 0 when every Visual Studio
    with the Office workload here passes, 1 when one does not or when there is none. Reads one XML
    file per Visual Studio and writes nothing.

.EXAMPLE
    pwsh -File Testbed/host/Publish-AddInPayload.ps1 -SelfTest
    pwsh -File Testbed/host/Publish-AddInPayload.ps1 -CompareInstalledTargets
    pwsh -File Testbed/host/Publish-AddInPayload.ps1
    pwsh -File Testbed/host/Publish-AddInPayload.ps1 -Ref 1a2b3c4 -VstoRuntimePath D:\media\vstor_redist.exe
#>
[CmdletBinding()]
param(
    [string] $RepoRoot,
    [string] $OutDir,
    [string] $Ref = 'HEAD',
    [string] $VstoRuntimePath,
    [string] $MSBuildPath,
    [string] $IsccPath,
    [string] $Version = '99.99.99.0',
    [int]    $BuildTimeoutMinutes = 20,
    [string] $ReleaseSigningThumbprint,
    [switch] $CompareInstalledTargets,
    [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# The running edition's OWN Security and Utility modules, imported by path - the fix
# Tools/Switch-AddInBuild.ps1 made for Q81, shared by the host scripts through the file below, whose
# header has the measurement. Windows PowerShell 5.1 started by Start-Process from PowerShell 7
# resolves both modules to 7's copies and can use neither: this script stopped on Get-FileHash,
# before it built anything, and the Cert: drive and Get-AuthenticodeSignature were missing too.
. (Join-Path $PSScriptRoot 'OwnEditionModules.ps1')

# Defaulted HERE rather than in param(): Windows PowerShell 5.1 leaves $PSScriptRoot EMPTY inside a
# param() default (the reason Testbed/host/New-LiveTestSettings.ps1 does the same).
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }

# ---------------------------------------------------------------------------------------------
# Everything this script names, in one place.
# ---------------------------------------------------------------------------------------------

# The VSTO runtime redistributable, pinned as every release has pinned it - the same file every
# release compiles into its installer, and Tools/Publish-Release.ps1 builds through this script.
# Microsoft-signed, version 10.0.60917.00.
$VstoRuntimeSha256 = 'CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18'
$VstoRuntimeBytes  = 41828424
$VstoRuntimeSource = 'https://download.microsoft.com/download/5/d/2/5d24f8f8-efbb-4b63-aa33-3785e3104713/vstor_redist.exe'

# The chain the VSTO targets would have given PrepareForRun, minus RegisterOfficeAddin. The
# common targets' own default is CopyFilesToOutputDirectory; the VSTO targets append
# VisualStudioForApplicationsBuild and then, for an add-in, RegisterOfficeAddin.
$PrepareForRunChain = @('CopyFilesToOutputDirectory', 'VisualStudioForApplicationsBuild')

# The three VSTO build tasks that write the build machine's registry, and every parameter the
# targets pass each of them. An inline stand-in must declare all of them or MSBuild refuses the
# call (MSB4064) - which is loud, and the self-test checks the list against the installed targets.
$StandInTasks = [ordered]@{
    'SetOffice2007AddInRegistration' = @{
        Parameters = @('Url', 'AddInName', 'OfficeApplication', 'FriendlyName', 'Description', 'LoadBehavior', 'Unregister', 'SolutionID', 'IsDocument')
        Output     = $null
        Writes     = 'HKCU\Software\Microsoft\Office\<app>\Addins\<name> - the add-in registration Outlook loads from'
        Scheduled  = $false
    }
    'SetInclusionListEntry' = @{
        Parameters = @('DeploymentManifestFullPath', 'CertificateThumbprint', 'Unregister')
        Output     = $null
        Writes     = 'HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid> - a VSTO trust entry for the build output'
        Scheduled  = $true
    }
    'RegisterFormRegions' = @{
        Parameters = @('AddInName', 'AssemblyName', 'OfficeApplication', 'Unregister')
        Output     = 'FormRegionNamesAndMessageClasses'
        Writes     = 'HKCU\Software\Microsoft\Office\Outlook\FormRegions\<message class> - Outlook form-region registrations'
        Scheduled  = $true
    }
}
$StandInSentinel = 'TESTBED-NO-HOST-WRITE'

# What every developer build carries (Properties/AssemblyInfo.cs, McpServer/Directory.Build.props),
# and therefore what a release build may never carry.
$DeveloperVersion = '99.99.99.0'

# The two executables the MCP server's publish must produce, each stamped with the release version.
$ServerExecutables = @('OutlookAI.McpServer.exe', 'OutlookAI.ComHost.exe')

# The host state guard 3 compares. Keys under HKCU, and CurrentUser certificate stores.
$WatchedRegistry = @(
    'Software\Microsoft\Office\Outlook\Addins',
    'Software\Microsoft\VSTO',
    'Software\Microsoft\Office\Outlook\FormRegions',
    'Software\Microsoft\Office\15.0\Outlook\FormRegions',
    'Software\Microsoft\Office\16.0\Outlook\FormRegions',
    'Software\Microsoft\Office\17.0\Outlook\FormRegions',
    'Software\OutlookAI'
)
$WatchedCertStores = @('My', 'TrustedPublisher', 'Root', 'CA', 'TrustedPeople', 'Disallowed')

# The one registry key whose change is always a failure, never a warning: it is what makes the
# maintainer's own Outlook load an add-in, and the build is the only thing here that could touch it.
$AddinRegistrationKey = 'Software\Microsoft\Office\Outlook\Addins\OutlookAI'

# The files through which the add-in and the mail server agree (both are compiled into both). The
# manifest carries their hashes; Testbed/guest/Install-OutlookAIAddIn.ps1 holds the same list.
$ContractFiles = @('Services/AddInServerContract.cs', 'Services/OfficeVersions.cs')

# What the guest script expects, repeated in the instructions this prints.
$GuestAddInRoot = 'C:\OutlookAI-Q5\addin'
$GuestMediaRoot = 'C:\OutlookAI-Q5\media'
$ManifestFileName = 'addin-payload.json'

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# Every native call whose stderr is redirected goes through this. Under Windows PowerShell 5.1,
# with $ErrorActionPreference = 'Stop', the FIRST line a native program writes to stderr is a
# terminating NativeCommandError whether the redirection is 2>$null, 2>&1 or *> - measured
# 2026-09-24; PowerShell 7 does not do it. 'Continue' holds only inside this function, so the
# caller's 'Stop' is untouched; callers still judge the call by $LASTEXITCODE. The same helper is
# restated in the other host scripts, in this repository's "restated locally" style.
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

# =============================================================================================
# PURE DECISIONS. Everything below this line up to the I/O section decides from its arguments
# alone, so -SelfTest can drive it on any machine.
# =============================================================================================

function Test-IsTestbedGuestIdentity {
    param([string] $ComputerName, [string] $UserName)
    # This script builds on the HOST. A guest has no Visual Studio, and the guest identity is the
    # one Testbed/host/New-AnswerFile.ps1 hands out: OAI-* machines, logged on as vmadmin.
    if ($ComputerName -and $ComputerName.StartsWith('OAI-', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($UserName -and $UserName -eq 'vmadmin') { return $true }
    return $false
}

function Test-FourPartVersion {
    param([string] $Value)
    if (-not $Value) { return $false }
    if ($Value -notmatch '^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}$') { return $false }
    foreach ($part in $Value.Split('.')) { if ([int]$part -gt 65535) { return $false } }
    return $true
}

function Get-PrepareForRunOverride {
    return ($PrepareForRunChain -join ';')
}

# The generated targets file. Only the task NAMES and parameters come from outside; the body of
# every stand-in is a single log line, so there is nothing in it that can write anything.
function Get-NoHostWriteTargets {
    $blocks = @()
    foreach ($name in $StandInTasks.Keys) {
        $spec = $StandInTasks[$name]
        $params = @()
        foreach ($p in $spec.Parameters) {
            $params += "      <$p ParameterType=`"System.String`" />"
        }
        if ($spec.Output) {
            $params += "      <$($spec.Output) ParameterType=`"Microsoft.Build.Framework.ITaskItem[]`" Output=`"true`" />"
        }
        $what = $spec.Writes.Replace('<', '&lt;').Replace('>', '&gt;')
        $blocks += @"
  <!-- The real $name writes: $what -->
  <UsingTask TaskName="$name" TaskFactory="RoslynCodeTaskFactory" AssemblyFile="`$(MSBuildToolsPath)\Microsoft.Build.Tasks.Core.dll" Override="true">
    <ParameterGroup>
$($params -join "`r`n")
    </ParameterGroup>
    <Task>
      <Code Type="Fragment" Language="cs"><![CDATA[
Log.LogMessage(MessageImportance.High, "$StandInSentinel`: $name did not run on the build host; nothing was written.");
]]></Code>
    </Task>
  </UsingTask>
"@
    }

    return @"
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <!--
    WRITTEN BY Testbed/host/Publish-AddInPayload.ps1 FOR ONE BUILD. NOT PART OF THE REPOSITORY.

    A plain build of a VSTO add-in REGISTERS it on the machine that builds it, and on the
    maintainer's workstation that repoints the real Outlook at a build output. These stand-ins
    replace the three VSTO build tasks that write the build machine's registry with tasks that
    only say they did not run. Override="true" makes them win whatever the import order
    (MSBuild 17.2+). The VSTO targets themselves run unmodified.

    This file is imported IN PLACE OF the machine's own CustomBeforeMicrosoftCommonTargets hook,
    so it imports that hook itself, from the same expression Microsoft.Common.CurrentVersion.targets
    uses to find it - a machine that has one keeps it.
  -->
  <Import Project="`$(MSBuildExtensionsPath)\v`$(MSBuildToolsVersion)\Custom.Before.Microsoft.Common.targets" Condition="Exists('`$(MSBuildExtensionsPath)\v`$(MSBuildToolsVersion)\Custom.Before.Microsoft.Common.targets')" />
$($blocks -join "`r`n")
</Project>
"@
}

# One argument string for MSBuild.exe. Values with spaces are quoted. A semicolon inside a value
# needs one of TWO spellings, and they are not interchangeable:
#   * a SCALAR (DefineConstants) takes %3B, the release workflow's spelling - the value is unescaped when a
#     task reads it;
#   * a LIST (PrepareForRunDependsOn) needs a REAL semicolon inside double quotes, because a target
#     list is split on semicolons BEFORE it is unescaped. MEASURED 2026-09-24, first run of this
#     script: %3B there made MSBuild look for ONE target called
#     "CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild" and stop with MSB4057.
function Format-MSBuildProperty {
    param([string] $Name, [string] $Value, [switch] $List)
    if ($List) { return "/p:$Name=`"$Value`"" }
    $v = $Value.Replace(';', '%3B')
    if ($v -match '\s') { return "/p:$Name=`"$v`"" }
    return "/p:$Name=$v"
}

function Get-MSBuildArgumentList {
    param(
        [Parameter(Mandatory = $true)] [string] $ProjectPath,
        [Parameter(Mandatory = $true)] [string] $Thumbprint,
        [Parameter(Mandatory = $true)] [string] $StandInTargets,
        [Parameter(Mandatory = $true)] [string] $ApplicationVersion,
        [Parameter(Mandatory = $true)] [string] $FileLog
    )

    $projectArg = $ProjectPath
    if ($projectArg -match '\s') { $projectArg = '"' + $projectArg + '"' }

    $list = @(
        $projectArg,
        '/t:Publish',
        # Mirrors the release workflow's "Build and Publish" step, property for property - and
        # since 2026-10-03 it IS the release's build (-ReleaseSigningThumbprint).
        (Format-MSBuildProperty 'Configuration' 'Release'),
        (Format-MSBuildProperty 'ApplicationVersion' $ApplicationVersion),
        (Format-MSBuildProperty 'PublishDir' 'publish\'),
        (Format-MSBuildProperty 'BootstrapperEnabled' 'true'),
        (Format-MSBuildProperty 'IsWebBootstrapper' 'false'),
        (Format-MSBuildProperty 'DefineConstants' 'VSTO40;TRACE'),
        # The throwaway key - or, for a release, the release key (-ReleaseSigningThumbprint).
        (Format-MSBuildProperty 'ManifestCertificateThumbprint' $Thumbprint),
        # Guard 1 and guard 2.
        (Format-MSBuildProperty 'PrepareForRunDependsOn' (Get-PrepareForRunOverride) -List),
        (Format-MSBuildProperty 'CustomBeforeMicrosoftCommonTargets' $StandInTargets),
        # No compiler server and no reused nodes: nothing may outlive the build and hold a handle.
        (Format-MSBuildProperty 'UseSharedCompilation' 'false'),
        '/nodeReuse:false',
        '/m:1',
        '/nologo',
        '/v:minimal'
    )

    # DETAILED, because the proof reads which TARGETS ran, and only detailed names every one.
    $flp = "LogFile=$FileLog;Verbosity=detailed;Encoding=UTF-8"
    if ($flp -match '\s') { $flp = '"' + $flp + '"' }
    $list += "/flp:$flp"
    return $list
}

# Guard 1's proof, BEFORE anything runs: MSBuild's own evaluation of the project with the same
# global properties (-getProperty evaluates and executes no target). What it prints is what
# PrepareForRun will depend on.
function Test-EvaluatedPrepareForRun {
    param([string] $EvaluatedValue)
    $v = ''
    if ($EvaluatedValue) { $v = ($EvaluatedValue -replace '\s', '') }
    $parts = @($v.Split(';') | Where-Object { $_ })
    $problems = @()
    if ($parts -contains 'RegisterOfficeAddin') { $problems += 'RegisterOfficeAddin is still on PrepareForRun - guard 1 did not take.' }
    if (-not ($parts -contains 'VisualStudioForApplicationsBuild')) { $problems += 'VisualStudioForApplicationsBuild is missing - the build would produce no signed manifests.' }
    if (-not ($parts -contains 'CopyFilesToOutputDirectory')) { $problems += 'CopyFilesToOutputDirectory is missing - the build would produce no output.' }
    return $problems
}

# Guard 1 and guard 2, proven from the build's own DETAILED log. VisualStudioForApplicationsBuild
# must have finished (it makes and signs the manifests); RegisterOfficeAddin must not appear at
# all; each stand-in that is scheduled must have said so; and the registration stand-in must NOT
# have run, because guard 1 removed its target - if it did, guard 1 failed and only guard 2
# caught it, which is worth knowing.
function Test-StandInSentinels {
    param([string] $LogText)
    $problems = @()
    $notes = @()
    if (-not $LogText) { $LogText = '' }
    $lower = $LogText.ToLowerInvariant()
    if (-not $lower.Contains('done building target "visualstudioforapplicationsbuild"')) {
        $problems += 'VisualStudioForApplicationsBuild never finished, so no signed manifests were made - or the log is not detailed enough to say.'
    }
    if ($lower.Contains('target "registerofficeaddin"')) {
        $problems += 'the RegisterOfficeAddin TARGET appears in the build log: guard 1 did not hold.'
    }
    foreach ($name in $StandInTasks.Keys) {
        $line = "$StandInSentinel`: $name did not run"
        $seen = $LogText.Contains($line)
        if ($StandInTasks[$name].Scheduled -and -not $seen) {
            $problems += "the $name stand-in never logged. Either the real task ran - and may have written the host's registry - or the build stopped before it; guard 3 decides which."
        }
        if (-not $StandInTasks[$name].Scheduled -and $seen) {
            $notes += "the $name stand-in RAN, so RegisterOfficeAddin was still scheduled: guard 1 did not hold and guard 2 caught it."
        }
    }
    return [pscustomobject]@{ Ok = ($problems.Count -eq 0); Problems = $problems; Notes = $notes }
}

# A form region would make RegisterFormRegions' output non-empty, and then the stand-in would
# change the add-in manifest. This add-in has none; refuse loudly the day one is added.
function Find-FormRegionDeclarations {
    param([string[]] $SourceTexts)
    $hits = @()
    foreach ($text in $SourceTexts) {
        if (-not $text) { continue }
        foreach ($needle in @('[FormRegionMessageClass', '[FormRegionName', 'FormRegionFactory', 'IFormRegionFactory')) {
            if ($text.Contains($needle)) { $hits += $needle }
        }
    }
    return @($hits | Sort-Object -Unique)
}

function ConvertTo-RsaKeyValueXml {
    param([Parameter(Mandatory = $true)] [string] $ModulusBase64, [Parameter(Mandatory = $true)] [string] $ExponentBase64)
    $m = ($ModulusBase64 -replace '\s', '')
    $e = ($ExponentBase64 -replace '\s', '')
    return "<RSAKeyValue><Modulus>$m</Modulus><Exponent>$e</Exponent></RSAKeyValue>"
}

# The signing key a deployment manifest carries, from ds:KeyInfo/ds:KeyValue/ds:RSAKeyValue, as
# the <RSAKeyValue> string the VSTO trust store holds. Every occurrence must be the same key.
function Get-ManifestSigningKeyXml {
    param([Parameter(Mandatory = $true)] [string] $ManifestXml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($ManifestXml)
    $nodes = $doc.SelectNodes("//*[local-name()='RSAKeyValue']")
    if ($null -eq $nodes -or $nodes.Count -eq 0) { throw 'The manifest carries no RSAKeyValue - it is not signed.' }
    $keys = @()
    foreach ($n in $nodes) {
        $mod = $n.SelectSingleNode("*[local-name()='Modulus']")
        $exp = $n.SelectSingleNode("*[local-name()='Exponent']")
        if ($null -eq $mod -or $null -eq $exp) { throw 'An RSAKeyValue in the manifest has no Modulus or no Exponent.' }
        $keys += (ConvertTo-RsaKeyValueXml -ModulusBase64 $mod.InnerText -ExponentBase64 $exp.InnerText)
    }
    $distinct = @($keys | Sort-Object -Unique)
    if ($distinct.Count -ne 1) { throw "The manifest is signed with $($distinct.Count) different keys; expected one." }
    return $distinct[0]
}

# Guard 3. A snapshot is a sorted list of lines; comparing two is set arithmetic.
function Compare-HostSnapshot {
    param([string[]] $Before, [string[]] $After)
    $b = @{}
    foreach ($l in @($Before)) { if ($null -ne $l) { $b[$l] = $true } }
    $a = @{}
    foreach ($l in @($After)) { if ($null -ne $l) { $a[$l] = $true } }
    $added = @()
    foreach ($l in $a.Keys) { if (-not $b.ContainsKey($l)) { $added += $l } }
    $removed = @()
    foreach ($l in $b.Keys) { if (-not $a.ContainsKey($l)) { $removed += $l } }
    return [pscustomobject]@{
        Added   = @($added | Sort-Object)
        Removed = @($removed | Sort-Object)
        Same    = ($added.Count -eq 0 -and $removed.Count -eq 0)
    }
}

# A trace of THIS build anywhere on the host: its directory (either slash direction, any case),
# the throwaway certificate's thumbprint, or its public modulus.
function Find-BuildFootprint {
    param([string[]] $Lines, [string] $BuildDirectory, [string] $Thumbprint, [string] $ModulusBase64)
    $needles = @()
    if ($BuildDirectory) {
        $d = $BuildDirectory.TrimEnd('\', '/')
        $needles += $d.ToLowerInvariant()
        $needles += $d.Replace('\', '/').ToLowerInvariant()
    }
    if ($Thumbprint) { $needles += $Thumbprint.ToLowerInvariant() }
    if ($ModulusBase64) {
        # A prefix is enough to be unambiguous and survives any re-wrapping of the value.
        $m = ($ModulusBase64 -replace '\s', '')
        if ($m.Length -gt 48) { $m = $m.Substring(0, 48) }
        $needles += $m.ToLowerInvariant()
    }
    $hits = @()
    foreach ($line in @($Lines)) {
        if ($null -eq $line) { continue }
        $low = $line.ToLowerInvariant()
        foreach ($n in $needles) {
            if ($n -and $low.Contains($n)) { $hits += $line; break }
        }
    }
    return @($hits)
}

# Guard 3's verdict on traces. A THROWAWAY key did not exist before the build, so any line naming it
# afterwards is a trace. The RELEASE key did, and must still be there afterwards - in My, and in
# TrustedPublisher and CA where the installed release put its public half - so for it only lines the
# build ADDED count. The build directory counts wherever it appears, in both modes.
function Find-HostTraces {
    param([string[]] $Before, [string[]] $After, [string] $BuildDirectory, [string] $Thumbprint, [string] $ModulusBase64, [switch] $KeyExistedBefore)
    if (-not $KeyExistedBefore) {
        return @(Find-BuildFootprint -Lines $After -BuildDirectory $BuildDirectory -Thumbprint $Thumbprint -ModulusBase64 $ModulusBase64)
    }
    $added = @((Compare-HostSnapshot -Before $Before -After $After).Added)
    $hits = @(Find-BuildFootprint -Lines $After -BuildDirectory $BuildDirectory) +
        @(Find-BuildFootprint -Lines $added -Thumbprint $Thumbprint -ModulusBase64 $ModulusBase64)
    return @($hits | Where-Object { $null -ne $_ } | Sort-Object -Unique)
}

# The release workflow's "Stamp assembly version", kept by the release build: every AssemblyVersion and
# AssemblyFileVersion attribute gets the release version. Both must be there - a source with only
# one would ship a DLL whose two versions disagree - so a text missing either returns $null.
function Set-AssemblyInfoVersion {
    param([string] $Text, [string] $Version)
    $assembly = [regex]'AssemblyVersion\("[^"]*"\)'
    $file = [regex]'AssemblyFileVersion\("[^"]*"\)'
    if ($assembly.Matches($Text).Count -eq 0 -or $file.Matches($Text).Count -eq 0) { return $null }
    $Text = $assembly.Replace($Text, ('AssemblyVersion("' + $Version + '")'))
    return $file.Replace($Text, ('AssemblyFileVersion("' + $Version + '")'))
}

# The release workflow's "Publish MCP server": framework-dependent on purpose (self-contained would add about
# 70 MB and pass the updater's 50 MB download cap), win-x64, and -p:Version so the server carries
# the add-in's version - one app, one version. No build servers: nothing may outlive the publish.
function Get-ServerPublishArgumentList {
    param([string] $ProjectPath, [string] $Version, [string] $OutputDir)
    $project = $ProjectPath
    if ($project -match '\s') { $project = '"' + $project + '"' }
    $output = $OutputDir
    if ($output -match '\s') { $output = '"' + $output + '"' }
    return @('publish', $project, '-c', 'Release', '-r', 'win-x64', '--self-contained', 'false',
        "-p:Version=$Version", '-o', $output, '--disable-build-servers', '--nologo')
}

function Test-ThumbprintText { param([string] $Value) return [bool]($Value -match '^[0-9A-Fa-f]{40}$') }

# Lines of a snapshot that belong to one key, so the registration key can be compared on its own.
function Select-KeyLines {
    param([string[]] $Lines, [string] $KeyPath)
    $prefix = ('HKCU\' + $KeyPath + '\').ToLowerInvariant()
    $exact = ('HKCU\' + $KeyPath + '|').ToLowerInvariant()
    $out = @()
    foreach ($l in @($Lines)) {
        if ($null -eq $l) { continue }
        $low = $l.ToLowerInvariant()
        if ($low.StartsWith($prefix) -or $low.StartsWith($exact)) { $out += $l }
    }
    return @($out)
}

function Test-Sha256Text { param([string] $Value) return [bool]($Value -match '^[0-9A-F]{64}$') }
function Test-CommitText { param([string] $Value) return [bool]($Value -match '^[0-9a-f]{40}$') }

function Test-PayloadManifestShape {
    param($Manifest)
    $problems = @()
    if ($null -eq $Manifest) { return @('no manifest at all') }
    if (-not (Test-CommitText ([string]$Manifest.commit))) { $problems += 'commit: not a 40-character lower-case hex id' }
    if (-not (Test-FourPartVersion ([string]$Manifest.version))) { $problems += 'version: not a four-part version' }
    if ($null -eq $Manifest.installer -or -not $Manifest.installer.file) { $problems += 'installer.file: missing' }
    elseif (-not (Test-Sha256Text ([string]$Manifest.installer.sha256))) { $problems += 'installer.sha256: not an upper-case SHA-256' }
    foreach ($f in @('OutlookAI.dll', 'OutlookAI.vsto', 'OutlookAI.dll.manifest')) {
        if ($null -eq $Manifest.addin -or -not (Test-Sha256Text ([string]$Manifest.addin.$f))) { $problems += "addin.${f}: not an upper-case SHA-256" }
    }
    if ($null -eq $Manifest.signing -or -not ([string]$Manifest.signing.publicKeyXml).StartsWith('<RSAKeyValue><Modulus>')) {
        $problems += 'signing.publicKeyXml: not an <RSAKeyValue> string'
    }
    if ($null -eq $Manifest.vstoRuntime -or ([string]$Manifest.vstoRuntime.sha256) -ne $VstoRuntimeSha256) {
        $problems += 'vstoRuntime.sha256: not the pinned redistributable'
    }
    foreach ($rel in $ContractFiles) {
        if ($null -eq $Manifest.contract -or -not (Test-Sha256Text ([string]$Manifest.contract.$rel))) { $problems += "contract.${rel}: not an upper-case SHA-256" }
    }
    return $problems
}

function Test-IsUnderGitWorkTreeButNotWork {
    param([string] $Path, [string] $WorkTreeRoot)
    if (-not $WorkTreeRoot) { return $false }
    $p = $Path.TrimEnd('\') + '\'
    $root = $WorkTreeRoot.TrimEnd('\') + '\'
    if (-not $p.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    $work = $root + '.work\'
    return (-not $p.StartsWith($work, [System.StringComparison]::OrdinalIgnoreCase))
}

# MSBuild's command line cuts a /p: switch into separate properties at every ',' and ';' outside
# quotes - its reference: "use a semicolon or a comma to separate multiple properties". Measured
# 2026-09-27 on MSBuild 18.10, on a one-target probe project:
#   * ','  /p:X=C:\a,b\t.targets stops on MSB1006, "Property is not valid". The first run to meet it
#          was an -OutDir ending "-A,B": its stand-in targets path reached MSBuild as a property
#          called "B\NoHostWrite.targets", and guard 1's evaluation refused.
#   * ';'  Format-MSBuildProperty writes it as %3B, which MSBuild reads back as ';' - but the /flp:
#          switch has no such escape: with a ';' in its LogFile path the detailed log went to a file
#          cut off at the ';', with no error, and guards 1 and 2 are proven FROM that log.
# Every path this script hands MSBuild - the project, the stand-in targets, the log - lies under
# -OutDir, so a ',' or ';' there is refused up front, before anything is archived, snapshotted,
# signed or built, rather than every argument having to get an escaping rule right. $null when the
# path is fine.
function Get-MSBuildPathProblem {
    param([string] $Path)
    $separators = @(@(',', ';') | Where-Object { $Path.Contains($_) })
    if ($separators.Count -eq 0) { return $null }
    $named = ($separators | ForEach-Object { "'$_'" }) -join ' and '
    return "-OutDir '$Path' contains $named, which MSBuild's command line reads as a separator - a ',' stops the build on MSB1006 ('Property is not valid'), a ';' sends the detailed log this script proves its guards from somewhere else. Every path this build hands MSBuild lies under -OutDir; choose one with no ',' or ';' in it"
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()

    # Compared with -ceq and searched with String.Contains. Never -like: Testbed/README.md
    # section 4b records the false greens a bracket in a -like pattern produced.
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

    Write-Host '== host identity guard =='
    Test-Case 'the maintainer workstation is not a guest' $false (Test-IsTestbedGuestIdentity -ComputerName 'PC657' -UserName 'jori')
    Test-Case 'an OAI- machine is a guest' $true (Test-IsTestbedGuestIdentity -ComputerName 'OAI-INDEXED' -UserName 'someone')
    Test-Case 'the prefix is case-insensitive' $true (Test-IsTestbedGuestIdentity -ComputerName 'oai-unindexed' -UserName 'x')
    Test-Case 'vmadmin anywhere is a guest' $true (Test-IsTestbedGuestIdentity -ComputerName 'PC657' -UserName 'vmadmin')

    Write-Host ''
    Write-Host '== version =='
    Test-Case 'the developer version is accepted' $true (Test-FourPartVersion '99.99.99.0')
    Test-Case 'three parts are refused' $false (Test-FourPartVersion '3.1.0')
    Test-Case 'a part over 65535 is refused' $false (Test-FourPartVersion '1.2.3.70000')
    Test-Case 'text is refused' $false (Test-FourPartVersion '1.2.3.x')

    Write-Host ''
    Write-Host '== guard 1: RegisterOfficeAddin is not scheduled =='
    $chain = Get-PrepareForRunOverride
    Test-Case 'the chain is exactly the VSTO one minus registration' 'CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild' $chain
    Test-Case 'the chain does not name RegisterOfficeAddin' $false ($chain.Contains('RegisterOfficeAddin'))

    Write-Host ''
    Write-Host '== guard 2: the stand-in targets =='
    $targets = Get-NoHostWriteTargets
    $doc = New-Object System.Xml.XmlDocument
    $parsed = $true
    try { $doc.LoadXml($targets) } catch { $parsed = $false }
    Test-Case 'the generated targets file is well-formed XML' $true $parsed
    if ($parsed) {
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('m', 'http://schemas.microsoft.com/developer/msbuild/2003')
        $tasks = $doc.SelectNodes('//m:UsingTask', $ns)
        Test-Case 'it declares exactly three tasks' 3 $tasks.Count
        foreach ($t in $tasks) {
            $name = $t.GetAttribute('TaskName')
            Test-Case "$name overrides (Override=true)" 'true' $t.GetAttribute('Override')
            Test-Case "$name is an inline task" 'RoslynCodeTaskFactory' $t.GetAttribute('TaskFactory')
            $declared = @($t.SelectNodes('m:ParameterGroup/*', $ns) | ForEach-Object { $_.LocalName })
            foreach ($p in $StandInTasks[$name].Parameters) {
                Test-Case "$name declares $p" $true ($declared -contains $p)
            }
            $code = $t.SelectSingleNode('m:Task/m:Code', $ns).InnerText
            Test-Case "$name only logs" $true ($code.Contains('Log.LogMessage') -and -not $code.Contains('Registry') -and -not $code.Contains('File.'))
            Test-Case "$name logs the sentinel" $true ($code.Contains("$StandInSentinel`: $name did not run"))
        }
        Test-Case 'the output of RegisterFormRegions is declared as an output' 'true' ($doc.SelectSingleNode("//m:UsingTask[@TaskName='RegisterFormRegions']/m:ParameterGroup/m:FormRegionNamesAndMessageClasses", $ns).GetAttribute('Output'))
        $import = $doc.SelectSingleNode('//m:Import', $ns)
        Test-Case 'the machine hook it replaces is chained, from MSBuild''s own expression' '$(MSBuildExtensionsPath)\v$(MSBuildToolsVersion)\Custom.Before.Microsoft.Common.targets' $import.GetAttribute('Project')
        Test-Case 'and only when it exists' $true ($import.GetAttribute('Condition').StartsWith('Exists('))
    }

    # Against the REAL VSTO targets, when this machine has them: every attribute the targets pass
    # each writing task must be declared by its stand-in, or MSBuild refuses the call. The same
    # comparison -CompareInstalledTargets makes - which FAILS where this skips (D7 (c)).
    $installedTargets = @(Find-InstalledVstoTargets)
    if ($installedTargets.Count -gt 0) {
        foreach ($r in @(Compare-StandInsWithTargets -TargetsText (Get-Content -LiteralPath $installedTargets[0] -Raw))) {
            Test-Case $r.What $r.Expected $r.Actual
        }
        Write-Host "       (read: $($installedTargets[0]))"
    }
    else {
        Write-Host '  SKIP the installed-targets comparison: no Visual Studio with the Office workload here.'
    }

    Write-Host ''
    Write-Host '== D7 (c): the comparison itself, on synthetic targets - runs everywhere =='
    $synthetic = @'
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup><PrepareForRunDependsOn>CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild;RegisterOfficeAddin;</PrepareForRunDependsOn></PropertyGroup>
  <Target Name="RegisterOfficeAddin"><SetOffice2007AddInRegistration Url="u" AddInName="n" Condition="c" /></Target>
  <Target Name="T"><SetInclusionListEntry DeploymentManifestFullPath="d" CertificateThumbprint="t" /><RegisterFormRegions AddInName="n" AssemblyName="a" OfficeApplication="o" /></Target>
</Project>
'@
    $r = @(Compare-StandInsWithTargets -TargetsText $synthetic)
    Test-Case 'targets the stand-ins cover pass every assertion' 0 @($r | Where-Object { [string]$_.Expected -cne [string]$_.Actual }).Count
    Test-Case 'one assertion pair per stand-in, and two about the chain' 8 $r.Count
    $r = @(Compare-StandInsWithTargets -TargetsText $synthetic.Replace('CertificateThumbprint="t"', 'CertificateThumbprint="t" NewWriterSwitch="x"'))
    Test-Case 'a parameter the stand-in lacks is named' 'NewWriterSwitch' (@($r | Where-Object { $_.What -eq 'SetInclusionListEntry - the installed targets pass nothing the stand-in lacks' })[0].Actual)
    $r = @(Compare-StandInsWithTargets -TargetsText $synthetic.Replace('RegisterOfficeAddin;', ''))
    Test-Case 'a chain that lost RegisterOfficeAddin; fails' 'False' ([string](@($r | Where-Object { $_.What -eq 'the installed targets still put RegisterOfficeAddin on PrepareForRun' })[0].Actual))
    $r = @(Compare-StandInsWithTargets -TargetsText $synthetic.Replace('<RegisterFormRegions AddInName="n" AssemblyName="a" OfficeApplication="o" />', ''))
    Test-Case 'a writer the targets stopped calling fails' 'False' ([string](@($r | Where-Object { $_.What -eq 'RegisterFormRegions - the installed targets do call it' })[0].Actual))

    Write-Host ''
    Write-Host '== the MSBuild command line =='
    $args1 = Get-MSBuildArgumentList -ProjectPath 'C:\b\source\OutlookAI.csproj' -Thumbprint 'ABCDEF' -StandInTargets 'C:\b\NoHostWrite.targets' -ApplicationVersion '99.99.99.0' -FileLog 'C:\b\msbuild.log'
    $joined = $args1 -join ' '
    Test-Case 'publishes, like every release' $true ($args1 -contains '/t:Publish')
    Test-Case 'semicolons in DefineConstants are escaped' $true ($args1 -contains '/p:DefineConstants=VSTO40%3BTRACE')
    Test-Case 'the chain is ONE quoted global property with a REAL semicolon' $true ($args1 -contains '/p:PrepareForRunDependsOn="CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild"')
    Test-Case 'and never the %3B spelling, which names one bogus target' $false ($joined.Contains('CopyFilesToOutputDirectory%3B'))
    Test-Case 'the detailed log is asked for' $true ($joined.Contains('Verbosity=detailed'))
    Test-Case 'the stand-ins are imported' $true ($args1 -contains '/p:CustomBeforeMicrosoftCommonTargets=C:\b\NoHostWrite.targets')
    Test-Case 'the throwaway key signs' $true ($args1 -contains '/p:ManifestCertificateThumbprint=ABCDEF')
    Test-Case 'no compiler server' $true ($args1 -contains '/p:UseSharedCompilation=false')
    Test-Case 'no reused nodes' $true ($args1 -contains '/nodeReuse:false')
    Test-Case 'nothing registers by name' $false ($joined.Contains('RegisterOfficeAddin'))
    $args2 = Get-MSBuildArgumentList -ProjectPath 'C:\a b\OutlookAI.csproj' -Thumbprint 'AB' -StandInTargets 'C:\a b\t.targets' -ApplicationVersion '1.2.3.4' -FileLog 'C:\a b\l.log'
    Test-Case 'a project path with a space is quoted' '"C:\a b\OutlookAI.csproj"' $args2[0]
    Test-Case 'a property value with a space is quoted' $true ($args2 -contains '/p:CustomBeforeMicrosoftCommonTargets="C:\a b\t.targets"')
    Test-Case 'a file log path with a space is quoted whole' $true ($args2 -contains '/flp:"LogFile=C:\a b\l.log;Verbosity=detailed;Encoding=UTF-8"')

    Write-Host ''
    Write-Host '== guard 1 proof: MSBuild''s own evaluation =='
    Test-Case 'the evaluated chain is accepted' 0 (Test-EvaluatedPrepareForRun "`r`n      CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild`r`n").Count
    Test-Case 'a chain still carrying RegisterOfficeAddin is refused' $true ((Test-EvaluatedPrepareForRun 'CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild;RegisterOfficeAddin;') -join ' ').Contains('guard 1 did not take')
    Test-Case 'an empty evaluation is refused' 2 (Test-EvaluatedPrepareForRun '').Count

    Write-Host ''
    Write-Host '== guards 1 and 2, proven from the build log =='
    $vsfab = 'Done building target "VisualStudioForApplicationsBuild" in project "OutlookAI.csproj".'
    $good = "a`r`n$StandInSentinel`: SetInclusionListEntry did not run on the build host; nothing was written.`r`n$StandInSentinel`: RegisterFormRegions did not run on the build host; nothing was written.`r`n$vsfab"
    $v = Test-StandInSentinels -LogText $good
    Test-Case 'both scheduled stand-ins logged and the manifests were made: accepted' $true $v.Ok
    Test-Case 'and nothing to note' 0 $v.Notes.Count
    $v = Test-StandInSentinels -LogText $vsfab
    Test-Case 'a log with no stand-in lines is refused' $false $v.Ok
    Test-Case 'naming both scheduled tasks' 2 $v.Problems.Count
    $v = Test-StandInSentinels -LogText ($good.Replace($vsfab, ''))
    Test-Case 'a build that never made the manifests is refused' $true (($v.Problems -join ' ').Contains('never finished'))
    $v = Test-StandInSentinels -LogText ($good + "`r`nTarget ""RegisterOfficeAddin"" in file ""x.targets"":")
    Test-Case 'the registration TARGET appearing is refused' $true (($v.Problems -join ' ').Contains('guard 1 did not hold'))
    $v = Test-StandInSentinels -LogText ($good + "`r`n$StandInSentinel`: SetOffice2007AddInRegistration did not run on the build host; nothing was written.")
    Test-Case 'the registration stand-in running is accepted' $true $v.Ok
    Test-Case 'but noted as guard 1 failing' 1 $v.Notes.Count

    Write-Host ''
    Write-Host '== form regions =='
    Test-Case 'the designer boilerplate is not a form region' 0 (Find-FormRegionDeclarations -SourceTexts @('internal sealed partial class ThisFormRegionCollection : Microsoft.Office.Tools.Outlook.FormRegionCollectionBase')).Count
    Test-Case 'a declared form region is found' '[FormRegionMessageClass' ((Find-FormRegionDeclarations -SourceTexts @('[FormRegionMessageClass("IPM.Note")]')) -join ',')

    Write-Host ''
    Write-Host '== the signing key, as the VSTO trust store holds it =='
    $manifest = '<asmv1:assembly xmlns:asmv1="urn:schemas-microsoft-com:asm.v1"><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><KeyInfo><KeyValue><RSAKeyValue><Modulus>AAEC
AwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue></KeyValue></KeyInfo></Signature><x><RSAKeyValue xmlns="http://www.w3.org/2000/09/xmldsig#"><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue></x></asmv1:assembly>'
    Test-Case 'namespace and line breaks do not matter' '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>' (Get-ManifestSigningKeyXml -ManifestXml $manifest)
    $twoKeys = $manifest.Replace('<x><RSAKeyValue xmlns="http://www.w3.org/2000/09/xmldsig#"><Modulus>AAECAwQ=', '<x><RSAKeyValue xmlns="http://www.w3.org/2000/09/xmldsig#"><Modulus>BBBBBBB=')
    $threw = $false
    try { $null = Get-ManifestSigningKeyXml -ManifestXml $twoKeys } catch { $threw = $true }
    Test-Case 'a manifest signed by two keys is refused' $true $threw
    $threw = $false
    try { $null = Get-ManifestSigningKeyXml -ManifestXml '<a/>' } catch { $threw = $true }
    Test-Case 'an unsigned manifest is refused' $true $threw

    Write-Host ''
    Write-Host '== guard 3: the host snapshot =='
    $before = @('HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI|Manifest=String:file:///C:/Source/SixFive7/OutlookAI/bin/Release/OutlookAI.vsto|vstolocal', 'CERT\My\AAAA')
    $same = Compare-HostSnapshot -Before $before -After $before
    Test-Case 'an unchanged host is the same' $true $same.Same
    $after = @('HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI|Manifest=String:file:///C:/w/.work/testbed-addin-payload/source/bin/Release/OutlookAI.vsto|vstolocal', 'CERT\My\AAAA')
    $diff = Compare-HostSnapshot -Before $before -After $after
    Test-Case 'a repointed registration is a difference' $false $diff.Same
    Test-Case 'with one line added' 1 $diff.Added.Count
    Test-Case 'and one removed' 1 $diff.Removed.Count
    $reg = Select-KeyLines -Lines $after -KeyPath 'Software\Microsoft\Office\Outlook\Addins\OutlookAI'
    Test-Case 'the registration key is selected on its own' 1 $reg.Count
    Test-Case 'a sibling add-in is not selected' 0 (Select-KeyLines -Lines @('HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAIX|Manifest=String:x') -KeyPath 'Software\Microsoft\Office\Outlook\Addins\OutlookAI').Count
    $fp = Find-BuildFootprint -Lines $after -BuildDirectory 'C:\w\.work\testbed-addin-payload' -Thumbprint 'FFFF' -ModulusBase64 ''
    Test-Case 'the build directory is found in forward-slash form' 1 $fp.Count
    $fp = Find-BuildFootprint -Lines @('CERT\TrustedPublisher\ABCDEF0123') -BuildDirectory 'C:\nowhere' -Thumbprint 'abcdef0123' -ModulusBase64 ''
    Test-Case 'the throwaway thumbprint is found, any case' 1 $fp.Count
    $fp = Find-BuildFootprint -Lines @('HKCU\Software\Microsoft\VSTO\Security\Inclusion\{g}|PublicKey=String:<RSAKeyValue><Modulus>xPrBMfdqBmy/OOrjKDK+dHjJIpIm85CujHMgp6EKe9PSDSKL</Modulus>') -BuildDirectory '' -Thumbprint '' -ModulusBase64 'xPrBMfdqBmy/OOrjKDK+dHjJIpIm85CujHMgp6EKe9PSDSKLzZrFlPsG'
    Test-Case 'the throwaway public key is found in a trust entry' 1 $fp.Count
    Test-Case 'the maintainer own entries are not a footprint' 0 (Find-BuildFootprint -Lines $before -BuildDirectory 'C:\w\.work\testbed-addin-payload' -Thumbprint 'FFFF' -ModulusBase64 'ZZZZ').Count

    Write-Host ''
    Write-Host '== the payload manifest =='
    $okManifest = [pscustomobject]@{
        commit      = ('a' * 40)
        version     = '99.99.99.0'
        installer   = [pscustomobject]@{ file = 'OutlookAI-v99.99.99.0.exe'; sha256 = ('A' * 64) }
        addin       = [pscustomobject]@{ 'OutlookAI.dll' = ('B' * 64); 'OutlookAI.vsto' = ('C' * 64); 'OutlookAI.dll.manifest' = ('D' * 64) }
        signing     = [pscustomobject]@{ publicKeyXml = '<RSAKeyValue><Modulus>AA==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>' }
        vstoRuntime = [pscustomobject]@{ sha256 = $VstoRuntimeSha256 }
        contract    = [pscustomobject]@{ 'Services/AddInServerContract.cs' = ('E' * 64); 'Services/OfficeVersions.cs' = ('F' * 64) }
    }
    Test-Case 'a complete manifest has no problems' 0 (Test-PayloadManifestShape $okManifest).Count
    $bad = $okManifest.PSObject.Copy()
    $bad.commit = 'HEAD'
    Test-Case 'a symbolic ref is not a commit' $true ((Test-PayloadManifestShape $bad) -join ' ').Contains('commit:')
    $bad = $okManifest.PSObject.Copy()
    $bad.contract = [pscustomobject]@{ 'Services/AddInServerContract.cs' = ('E' * 64) }
    Test-Case 'a manifest without every contract hash is refused' $true ((Test-PayloadManifestShape $bad) -join ' ').Contains('contract.Services/OfficeVersions.cs')

    Write-Host ''
    Write-Host '== the release build (-ReleaseSigningThumbprint) =='
    $info = "using System.Reflection;`r`n[assembly: AssemblyVersion(""99.99.99.0"")]`r`n[assembly: AssemblyFileVersion(""99.99.99.0"")]`r`n"
    $stampedInfo = Set-AssemblyInfoVersion -Text $info -Version '3.1.1.900'
    Test-Case 'AssemblyVersion is stamped' $true ($stampedInfo.Contains('AssemblyVersion("3.1.1.900")'))
    Test-Case 'AssemblyFileVersion is stamped' $true ($stampedInfo.Contains('AssemblyFileVersion("3.1.1.900")'))
    Test-Case 'and nothing else changes' $info.Replace('99.99.99.0', '3.1.1.900') $stampedInfo
    Test-Case 'a source without AssemblyFileVersion is refused' $true ($null -eq (Set-AssemblyInfoVersion -Text 'AssemblyVersion("1.0.0.0")' -Version '3.1.1.900'))
    Test-Case 'a source without AssemblyVersion is refused' $true ($null -eq (Set-AssemblyInfoVersion -Text 'AssemblyFileVersion("1.0.0.0")' -Version '3.1.1.900'))
    Test-Case 'a 40-character thumbprint is accepted' $true (Test-ThumbprintText '2578F7B869383572E751DD6B61B5374C55C6E995')
    Test-Case 'a short one is refused' $false (Test-ThumbprintText '2578F7B869383572E751DD6B61B5374C55C6E99')
    Test-Case 'a non-hex one is refused' $false (Test-ThumbprintText '2578F7B869383572E751DD6B61B5374C55C6E99Z')
    $releaseKey = @('CERT\My\2578F7B869383572E751DD6B61B5374C55C6E995|CN=OutlookAI', 'CERT\TrustedPublisher\2578F7B869383572E751DD6B61B5374C55C6E995|CN=OutlookAI')
    Test-Case 'the release key, there before and after, is no trace' 0 (Find-HostTraces -Before $releaseKey -After $releaseKey -BuildDirectory 'C:\w\.work\release' -Thumbprint '2578F7B869383572E751DD6B61B5374C55C6E995' -KeyExistedBefore).Count
    Test-Case 'but the same lines ARE a trace of a throwaway key' 2 (Find-HostTraces -Before $releaseKey -After $releaseKey -BuildDirectory 'C:\w\.work\release' -Thumbprint '2578F7B869383572E751DD6B61B5374C55C6E995').Count
    $newTrust = 'HKCU\Software\Microsoft\VSTO\Security\Inclusion\{g}|Thumbprint=String:2578f7b869383572e751dd6b61b5374c55c6e995'
    Test-Case 'a line the build added naming the release key is a trace' 1 (Find-HostTraces -Before $releaseKey -After ($releaseKey + $newTrust) -BuildDirectory 'C:\w\.work\release' -Thumbprint '2578F7B869383572E751DD6B61B5374C55C6E995' -KeyExistedBefore).Count
    $oldDirTrace = 'HKCU\Software\Microsoft\VSTO\Security\Inclusion\{h}|Url=String:file:///C:/w/.work/release/source/publish/OutlookAI.vsto'
    Test-Case 'the build directory is a trace even on a line that was already there' 1 (Find-HostTraces -Before @($oldDirTrace) -After @($oldDirTrace) -BuildDirectory 'C:\w\.work\release' -Thumbprint 'FFFF' -KeyExistedBefore).Count
    $publishArgs = Get-ServerPublishArgumentList -ProjectPath 'C:\b\source\McpServer\OutlookAI.McpServer\OutlookAI.McpServer.csproj' -Version '3.1.1.900' -OutputDir 'C:\b\source\publish\McpServer'
    Test-Case 'the server is published, Release, win-x64' 'publish|-c Release|-r win-x64' ('{0}|{1} {2}|{3} {4}' -f $publishArgs[0], $publishArgs[2], $publishArgs[3], $publishArgs[4], $publishArgs[5])
    Test-Case 'framework-dependent - self-contained would pass the 50 MB cap' $true (($publishArgs -join ' ').Contains('--self-contained false'))
    Test-Case 'with the release version' $true ($publishArgs -contains '-p:Version=3.1.1.900')
    Test-Case 'into publish\McpServer, where Installer.iss picks it up' $true ($publishArgs -contains 'C:\b\source\publish\McpServer')
    Test-Case 'with no build server left behind' $true ($publishArgs -contains '--disable-build-servers')
    $spaced = Get-ServerPublishArgumentList -ProjectPath 'C:\a b\s.csproj' -Version '1.2.3.4' -OutputDir 'C:\a b\out'
    Test-Case 'paths with a space are quoted' '"C:\a b\s.csproj"|"C:\a b\out"' ($spaced[1] + '|' + $spaced[10])
    Test-Case 'the two executables the release must stamp' 'OutlookAI.McpServer.exe,OutlookAI.ComHost.exe' ($ServerExecutables -join ',')

    Write-Host ''
    Write-Host '== the output directory =='
    Test-Case 'under .work is allowed' $false (Test-IsUnderGitWorkTreeButNotWork -Path 'C:\r\.work\testbed-addin-payload' -WorkTreeRoot 'C:\r')
    Test-Case 'elsewhere in the working tree is refused' $true (Test-IsUnderGitWorkTreeButNotWork -Path 'C:\r\Testbed\out' -WorkTreeRoot 'C:\r')
    Test-Case 'outside the working tree is allowed' $false (Test-IsUnderGitWorkTreeButNotWork -Path 'D:\payloads' -WorkTreeRoot 'C:\r')
    Test-Case 'an -OutDir with a comma is refused - MSBuild would split it' $true ([string](Get-MSBuildPathProblem 'C:\w\.work\testbed-addin-payload-A,B')).Contains("contains ','")
    Test-Case 'and one with a semicolon' $true ([string](Get-MSBuildPathProblem 'C:\w\.work\a;b')).Contains("contains ';'")
    Test-Case 'a path with both is refused naming both' $true ([string](Get-MSBuildPathProblem 'C:\a,b;c')).Contains("',' and ';'")
    Test-Case 'a space is accepted - Format-MSBuildProperty quotes it' '' ([string](Get-MSBuildPathProblem 'C:\a b\.work\testbed-addin-payload'))

    Write-Host '== the toolchain refusal belongs to the VSTO check (this file''s own syntax tree) =='
    # edc9dfd inserted the release block between `if ($vstoTargets.Count -gt 0)` and its else, so the
    # else - the "no VSTO targets" refusal - became the release check's and refused EVERY testbed build.
    $ownAst = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$null)
    $refusalIfs = @($ownAst.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.IfStatementAst] -and $null -ne $n.ElseClause -and
                $n.ElseClause.Extent.Text.Contains('vswhere found no VSTO build targets')
            }, $true))
    Test-Case 'one if statement carries the no-VSTO-targets refusal in its else' 1 $refusalIfs.Count
    $refusalCondition = ''
    if ($refusalIfs.Count -eq 1) { $refusalCondition = $refusalIfs[0].Clauses[0].Item1.Extent.Text }
    Test-Case 'and its condition is the VSTO targets count, not the release switch' '$vstoTargets.Count -gt 0' $refusalCondition

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a real build, and then a guest, can settle these:'
    Write-Host '  * that MSBuild on this machine honours Override="true" and the stand-ins actually run'
    Write-Host '  * that the host snapshot really is unchanged after a build'
    Write-Host '  * that SignFile accepts the throwaway certificate, and that it is gone afterwards'
    Write-Host '  * that the release key signs, and that the MCP server publishes with the release version'
    Write-Host '    - which Tools/Publish-Release.ps1''s dry run shows'
    Write-Host '  * that the installer this produces installs silently and the add-in LOADS on a guest'
    Write-Host '    - which is Testbed/guest/Install-OutlookAIAddIn.ps1''s job, not this script''s'

    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

# =============================================================================================
# I/O
# =============================================================================================

# D7 (c): guard 2's stand-ins against the text of one VSTO targets file. Each result is one
# assertion - What, Expected, Actual. Pure, so the self-test feeds it synthetic targets anywhere and
# -CompareInstalledTargets the real ones.
function Compare-StandInsWithTargets {
    param([Parameter(Mandatory = $true)] [string] $TargetsText)
    $vsto = [xml]$TargetsText
    $results = @()
    foreach ($name in $StandInTasks.Keys) {
        $used = @()
        foreach ($node in $vsto.GetElementsByTagName($name)) {
            foreach ($attr in $node.Attributes) { if ($attr.Name -ne 'Condition') { $used += $attr.Name } }
        }
        $used = @($used | Sort-Object -Unique)
        $missing = @($used | Where-Object { $StandInTasks[$name].Parameters -notcontains $_ })
        $results += [pscustomobject]@{ What = "$name - the installed targets pass nothing the stand-in lacks"; Expected = ''; Actual = ($missing -join ',') }
        $results += [pscustomobject]@{ What = "$name - the installed targets do call it"; Expected = $true; Actual = ($used.Count -gt 0) }
    }
    $results += [pscustomobject]@{ What = 'the installed targets still put RegisterOfficeAddin on PrepareForRun'; Expected = $true; Actual = $TargetsText.Contains('RegisterOfficeAddin;') }
    $results += [pscustomobject]@{ What = 'and still run VisualStudioForApplicationsBuild from it'; Expected = $true; Actual = $TargetsText.Contains('VisualStudioForApplicationsBuild;') }
    return $results
}

# -CompareInstalledTargets: the comparison against EVERY installed VSTO targets file, as a verdict.
# 0 when each passes; 1 when one fails or when there is none - a comparison that did not run proves
# nothing, and the release needs it to have run (D7 (c)).
function Invoke-InstalledTargetsComparison {
    Write-Host "Testbed/host/Publish-AddInPayload.ps1 -CompareInstalledTargets under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    $paths = @(Find-InstalledVstoTargets)
    if ($paths.Count -eq 0) {
        Write-Host '  FAIL no Visual Studio with the Office workload on this machine, so the stand-ins were compared with nothing.'
        return 1
    }
    $checks = 0
    $failures = 0
    foreach ($p in $paths) {
        Write-Host "  == $p"
        foreach ($r in @(Compare-StandInsWithTargets -TargetsText (Get-Content -LiteralPath $p -Raw))) {
            $checks++
            if ([string]$r.Expected -ceq [string]$r.Actual) { Write-Host "  OK   $($r.What)" }
            else {
                $failures++
                Write-Host "  FAIL $($r.What) - expected [$($r.Expected)], got [$($r.Actual)]"
            }
        }
    }
    Write-Host "$checks assertion(s) across $($paths.Count) targets file(s), $failures failure(s)."
    if ($failures -gt 0) { return 1 }
    return 0
}

function Find-InstalledVstoTargets {
    $found = @()
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere)) { return $found }
    $roots = @(Invoke-NativeCommand { & $vswhere -all -products * -requires Microsoft.VisualStudio.Workload.Office -property installationPath 2>$null })
    foreach ($root in $roots) {
        if (-not $root) { continue }
        $dir = Join-Path $root 'MSBuild\Microsoft\VisualStudio'
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $found += @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue |
                ForEach-Object { Join-Path $_.FullName 'OfficeTools\Microsoft.VisualStudio.Tools.Office.targets' } |
                Where-Object { Test-Path -LiteralPath $_ })
    }
    return $found
}

function Resolve-MSBuild {
    if ($MSBuildPath) {
        if (-not (Test-Path -LiteralPath $MSBuildPath)) { throw "-MSBuildPath does not exist: $MSBuildPath" }
        return $MSBuildPath
    }
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere)) {
        throw @"
REFUSING TO BUILD: vswhere.exe is not on this machine, so there is no Visual Studio to build with.

The add-in is a VSTO project and only MSBuild with Visual Studio's Office/SharePoint development
workload can build it - the toolchain every release is built with.
Testbed/MEDIA.md says what to install. Or pass -MSBuildPath.
"@
    }
    $roots = @(Invoke-NativeCommand { & $vswhere -latest -products * -requires Microsoft.VisualStudio.Workload.Office -property installationPath 2>$null })
    if ($roots.Count -eq 0 -or -not $roots[0]) {
        throw @"
REFUSING TO BUILD: no Visual Studio on this machine has the Office/SharePoint development workload
(Microsoft.VisualStudio.Workload.Office). Without it there are no VSTO build targets, and the add-in
cannot be built. Add the workload in the Visual Studio Installer, or pass -MSBuildPath.
"@
    }
    $candidate = Join-Path $roots[0] 'MSBuild\Current\Bin\amd64\MSBuild.exe'
    if (-not (Test-Path -LiteralPath $candidate)) { $candidate = Join-Path $roots[0] 'MSBuild\Current\Bin\MSBuild.exe' }
    if (-not (Test-Path -LiteralPath $candidate)) { throw "Visual Studio at $($roots[0]) has no MSBuild.exe under MSBuild\Current\Bin." }
    return $candidate
}

function Resolve-Iscc {
    if ($IsccPath) {
        if (-not (Test-Path -LiteralPath $IsccPath)) { throw "-IsccPath does not exist: $IsccPath" }
        return $IsccPath
    }
    $onPath = Get-Command iscc.exe -CommandType Application -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    foreach ($c in @(
            (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
            (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe'))) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    throw @"
REFUSING TO BUILD: Inno Setup 6's ISCC.exe was not found (PATH, %LOCALAPPDATA%\Programs\Inno Setup 6,
Program Files). It compiles Installer.iss - the product's own installer, the one
every release is built with. Testbed/MEDIA.md says what to install.
Or pass -IsccPath.
"@
}

function Resolve-Dotnet {
    $onPath = @(Get-Command dotnet.exe -CommandType Application -ErrorAction SilentlyContinue)
    if ($onPath.Count -gt 0) { return $onPath[0].Source }
    $c = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    if (Test-Path -LiteralPath $c) { return $c }
    throw 'REFUSING TO BUILD THE RELEASE: dotnet.exe was not found (PATH, Program Files\dotnet). The release publishes the MCP server with the .NET 10 SDK.'
}

# Start a tool with its output going STRAIGHT TO FILES - never through this process's pipe - and
# a deadline. A build that spawns children can otherwise hold a pipe open after it has finished,
# and the declared timeout then never fires. .Handle is read before waiting because
# Start-Process -PassThru otherwise loses the exit code (measured, Install-DotnetSdk.ps1).
function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)] [string] $FilePath,
        [Parameter(Mandatory = $true)] [string[]] $ArgumentList,
        [Parameter(Mandatory = $true)] [string] $LogStem,
        [int] $TimeoutMinutes = 20,
        [string] $WorkingDirectory
    )
    $out = "$LogStem.out.txt"
    $err = "$LogStem.err.txt"
    foreach ($f in @($out, $err)) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } }
    $startArgs = @{
        FilePath               = $FilePath
        ArgumentList           = ($ArgumentList -join ' ')
        RedirectStandardOutput = $out
        RedirectStandardError  = $err
        NoNewWindow            = $true
        PassThru               = $true
    }
    if ($WorkingDirectory) { $startArgs['WorkingDirectory'] = $WorkingDirectory }
    $p = Start-Process @startArgs
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = -1; TimedOut = $true; Out = $out; Err = $err }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false; Out = $out; Err = $err }
}

function Show-LogTail([string] $Path, [int] $Lines = 30) {
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        foreach ($l in (Get-Content -LiteralPath $Path -Tail $Lines)) { Write-Host "      | $l" }
    }
}

function Format-RegistryData($Value, [Microsoft.Win32.RegistryValueKind] $Kind) {
    if ($null -eq $Value) { return '<null>' }
    switch ($Kind) {
        'Binary' { return (($Value | ForEach-Object { '{0:X2}' -f $_ }) -join '') }
        'MultiString' { return ($Value -join '|') }
        default { return [string]$Value }
    }
}

# Read-only. Every value under every watched key, one line each.
function Get-RegistryLines {
    param([string[]] $Keys)
    $lines = @()
    foreach ($rel in $Keys) {
        $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($rel, $false)
        if ($null -eq $root) { continue }
        $stack = New-Object System.Collections.Stack
        $stack.Push(@($rel, $root))
        while ($stack.Count -gt 0) {
            $pair = $stack.Pop()
            $path = $pair[0]
            $key = $pair[1]
            try {
                $lines += ('HKCU\' + $path + '|<key>')
                foreach ($name in $key.GetValueNames()) {
                    $kind = $key.GetValueKind($name)
                    $data = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    $lines += ('HKCU\' + $path + '|' + $name + '=' + $kind + ':' + (Format-RegistryData $data $kind))
                }
                foreach ($sub in $key.GetSubKeyNames()) {
                    $child = $key.OpenSubKey($sub, $false)
                    if ($null -ne $child) { $stack.Push(@(($path + '\' + $sub), $child)) }
                }
            }
            finally {
                $key.Close()
            }
        }
    }
    return @($lines | Sort-Object)
}

function Get-CertLines {
    param([string[]] $Stores)
    $lines = @()
    foreach ($s in $Stores) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            foreach ($c in $store.Certificates) { $lines += ('CERT\' + $s + '\' + $c.Thumbprint) }
        }
        catch {
            $lines += ('CERT\' + $s + '|unreadable:' + $_.Exception.Message)
        }
        finally {
            $store.Close()
        }
    }
    return @($lines | Sort-Object)
}

function Get-HostSnapshot {
    return @((Get-RegistryLines -Keys $WatchedRegistry) + (Get-CertLines -Stores $WatchedCertStores))
}

# The throwaway certificate, out of EVERY CurrentUser store, private key included.
#
# NOT ONLY Cert:\CurrentUser\My. MEASURED 2026-09-24, first run of this script (pwsh 7.6.6,
# Windows 11 10.0.26200): New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My ALSO
# puts a copy of the certificate into Cert:\CurrentUser\CA, Intermediate Certification
# Authorities. Removing it from My left that copy behind, and guard 3 caught it as a trace of
# the build. Only a certificate carrying BOTH this build's thumbprint and its subject is removed,
# so nothing of the maintainer's can match.
function Remove-ThrowawayCertificate {
    param([Parameter(Mandatory = $true)] [string] $Thumbprint)
    $my = "Cert:\CurrentUser\My\$Thumbprint"
    if (Test-Path -LiteralPath $my) {
        Remove-Item -LiteralPath $my -DeleteKey -ErrorAction SilentlyContinue
    }
    foreach ($s in $WatchedCertStores) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $mine = @($store.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint -and $_.Subject -eq 'CN=OutlookAI Testbed' })
            foreach ($c in $mine) {
                $store.Remove($c)
                Say "  removed the copy New-SelfSignedCertificate left in Cert:\CurrentUser\$s"
            }
        }
        catch {
            Say "  !! could not open Cert:\CurrentUser\$s to clean it: $($_.Exception.Message)"
        }
        finally {
            $store.Close()
        }
    }
    $left = @()
    foreach ($s in $WatchedCertStores) {
        if (Test-Path -LiteralPath "Cert:\CurrentUser\$s\$Thumbprint") { $left += $s }
    }
    if ($left.Count -gt 0) {
        Say "  !! the throwaway certificate is STILL in: $($left -join ', ') - guard 3 will refuse. Remove it by thumbprint $Thumbprint."
    }
    else {
        Say '  throwaway certificate deleted from every CurrentUser store, private key included'
    }
}

# Only ever called when guard 3 has already found the registration changed: puts back exactly
# what the snapshot recorded, and says so. Reversing our own side effect is strictly safer than
# leaving the maintainer's Outlook pointed at a build.
function Restore-AddinRegistration {
    param([string[]] $BeforeLines)
    # The key's OWN values only - not a subkey's, which Select-KeyLines also returns.
    $own = 'HKCU\' + $AddinRegistrationKey + '|'
    $mine = @($BeforeLines | Where-Object { $_ -and $_.StartsWith($own, [System.StringComparison]::OrdinalIgnoreCase) -and -not $_.EndsWith('|<key>') })
    $existedBefore = @(Select-KeyLines -Lines $BeforeLines -KeyPath $AddinRegistrationKey).Count -gt 0
    if (-not $existedBefore) {
        Say "  RESTORE: HKCU\$AddinRegistrationKey did not exist before the build - deleting it"
        [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($AddinRegistrationKey, $false)
        return
    }
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($AddinRegistrationKey)
    try {
        $wanted = @{}
        foreach ($line in $mine) {
            $afterPipe = $line.Substring($line.IndexOf('|') + 1)
            $name = $afterPipe.Substring(0, $afterPipe.IndexOf('='))
            $rest = $afterPipe.Substring($afterPipe.IndexOf('=') + 1)
            $kind = $rest.Substring(0, $rest.IndexOf(':'))
            $data = $rest.Substring($rest.IndexOf(':') + 1)
            $wanted[$name] = $true
            if ($kind -eq 'DWord') { $key.SetValue($name, [int]$data, [Microsoft.Win32.RegistryValueKind]::DWord) }
            else { $key.SetValue($name, $data, [Microsoft.Win32.RegistryValueKind]::String) }
            Say "  RESTORE: $name = $data"
        }
        foreach ($name in $key.GetValueNames()) {
            if (-not $wanted.ContainsKey($name)) { $key.DeleteValue($name, $false); Say "  RESTORE: removed $name, which the build added" }
        }
    }
    finally { $key.Close() }
}

# =============================================================================================
# MAIN
# =============================================================================================
if ($SelfTest -and ($CompareInstalledTargets -or $ReleaseSigningThumbprint)) { throw '-SelfTest runs alone.' }
if ($CompareInstalledTargets -and $ReleaseSigningThumbprint) { throw '-CompareInstalledTargets runs alone.' }
if ($SelfTest) { exit (Invoke-SelfTest) }
if ($CompareInstalledTargets) { exit (Invoke-InstalledTargetsComparison) }

if (Test-IsTestbedGuestIdentity -ComputerName $env:COMPUTERNAME -UserName $env:USERNAME) {
    throw "REFUSING TO RUN on '$env:COMPUTERNAME' as '$env:USERNAME': this is a testbed guest's identity. This script builds on the HOST; the guest half is Testbed/guest/Install-OutlookAIAddIn.ps1."
}
if (-not (Test-FourPartVersion $Version)) { throw "-Version '$Version' is not a four-part version (each part 0-65535)." }

$Release = [bool]$ReleaseSigningThumbprint
$releaseCert = $null
if ($Release) {
    if (-not (Test-ThumbprintText $ReleaseSigningThumbprint)) { throw "-ReleaseSigningThumbprint '$ReleaseSigningThumbprint' is not a 40-character certificate thumbprint." }
    if ($Version -eq $DeveloperVersion) { throw "REFUSING: a release build carries its own -Version. $DeveloperVersion is what every developer build carries, and no updater would ever replace it." }
    $releaseCert = Get-Item -LiteralPath "Cert:\CurrentUser\My\$ReleaseSigningThumbprint" -ErrorAction SilentlyContinue
    if (-not $releaseCert) { throw "REFUSING: the release certificate $ReleaseSigningThumbprint is not in Cert:\CurrentUser\My." }
    if (-not $releaseCert.HasPrivateKey) { throw "REFUSING: Cert:\CurrentUser\My\$ReleaseSigningThumbprint has no private key on this machine, so it cannot sign." }
}

if (-not $OutDir) { $OutDir = Join-Path $RepoRoot '.work\testbed-addin-payload' }
if (-not $VstoRuntimePath) { $VstoRuntimePath = Join-Path $RepoRoot '.work\media\vstor_redist.exe' }
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
$outDirProblem = Get-MSBuildPathProblem $OutDir
if ($outDirProblem) { throw "REFUSING: $outDirProblem. Nothing was created, signed or built." }

$workTree = (@(Invoke-NativeCommand { & git -C $RepoRoot rev-parse --show-toplevel 2>$null }) | Out-String).Trim().Replace('/', '\')
if (Test-IsUnderGitWorkTreeButNotWork -Path $OutDir -WorkTreeRoot $workTree) {
    throw "REFUSING: -OutDir $OutDir is inside the git working tree $workTree but not under .work\. A build tree there is one 'git add .' away from being committed."
}

$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

$sourceZip = Join-Path $OutDir 'Source.zip'
$sourceDir = Join-Path $OutDir 'source'
$standInPath = Join-Path $OutDir 'NoHostWrite.targets'
$msbuildLog = Join-Path $OutDir 'msbuild.log'
$installerDir = Join-Path $OutDir 'installer'
$stageDir = Join-Path $OutDir 'stage'
$addinZip = Join-Path $OutDir 'AddIn.zip'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# ---------------------------------------------------------------------------------------------
Say '== Toolchain =='
$msbuild = Resolve-MSBuild
$iscc = Resolve-Iscc
$vstoTargets = @(Find-InstalledVstoTargets)
$msbuildVersion = (@(Invoke-NativeCommand { & $msbuild -version -nologo 2>$null }) | Select-Object -Last 1)
Say "  MSBuild  $msbuild ($msbuildVersion)"
Say "  ISCC     $iscc"
if ($vstoTargets.Count -gt 0) { Say "  VSTO     $($vstoTargets[0])" }
else { throw 'REFUSING TO BUILD: vswhere found no VSTO build targets (OfficeTools\Microsoft.VisualStudio.Tools.Office.targets) in any Visual Studio with the Office workload.' }
# That refusal belongs to the VSTO targets and to nothing else. From edc9dfd until 2026-10-03 the
# release's dotnet lines sat between that if and its else, which bound the else to `if ($Release)`:
# every testbed build stopped here, "found no VSTO build targets", on a machine that has them, and a
# release build on a machine without them was never refused.
# -SelfTest pins the shape from this file's own syntax tree.
$dotnet = $null
if ($Release) {
    $dotnet = Resolve-Dotnet
    Say "  dotnet   $dotnet"
}

# ---------------------------------------------------------------------------------------------
Say ''
Say '== Media: the VSTO runtime =='
if (-not (Test-Path -LiteralPath $VstoRuntimePath)) {
    throw @"
NOT STAGED: $VstoRuntimePath

This is a PRECONDITION, not an artefact - nothing here produces it and nothing here downloads it.
It is the Visual Studio 2010 Tools for Office runtime redistributable, the same file every
release compiles into its installer. Get it from
    $VstoRuntimeSource
put it there (or pass -VstoRuntimePath), and re-run. Testbed/MEDIA.md declares it.
"@
}
$vstoHash = (Get-FileHash -LiteralPath $VstoRuntimePath -Algorithm SHA256).Hash
$vstoLength = (Get-Item -LiteralPath $VstoRuntimePath).Length
if ($vstoHash -ne $VstoRuntimeSha256 -or $vstoLength -ne $VstoRuntimeBytes) {
    throw @"
The staged VSTO runtime is not the pinned file.

  expected  $VstoRuntimeSha256, $VstoRuntimeBytes bytes
  actual    $vstoHash, $vstoLength bytes
  file      $VstoRuntimePath

The pin is every release's: this script builds them (Tools/Publish-Release.ps1).
"@
}
$vstoSignature = Get-AuthenticodeSignature -LiteralPath $VstoRuntimePath
Say "  $VstoRuntimePath"
Say "  sha256 $vstoHash, $vstoLength bytes - MATCHES the pin"
Say "  signed $($vstoSignature.Status) by $($vstoSignature.SignerCertificate.Subject)"

# ---------------------------------------------------------------------------------------------
Say ''
Say '== Source, from a commit =='
$commit = (@(Invoke-NativeCommand { & git -C $RepoRoot rev-parse --verify "$Ref^{commit}" 2>&1 }) | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or -not (Test-CommitText $commit)) { throw "git could not resolve '$Ref' to a commit in $RepoRoot - $commit" }
$dirty = @(Invoke-NativeCommand { & git -C $RepoRoot status --porcelain })
if ($dirty.Count -gt 0) {
    Say "  NOTE the working tree has $($dirty.Count) uncommitted change(s). They are NOT in this build - it is"
    Say '       built from a commit on purpose. Commit first if the guest is meant to see them.'
}
foreach ($p in @($sourceZip, $sourceDir, $installerDir, $stageDir)) {
    if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
}
Invoke-NativeCommand { & git -C $RepoRoot archive --format=zip -o $sourceZip $commit }
if ($LASTEXITCODE -ne 0) { throw "git archive failed (exit $LASTEXITCODE)." }
# The framework's own zip rather than Expand-Archive and Compress-Archive, whose module a mixed
# PSModulePath (see OwnEditionModules.ps1, dot-sourced at the top) hands to the other edition -
# measured: 7's Archive 1.2.6 running inside 5.1. As Tools/Switch-AddInBuild.ps1 does.
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZip, $sourceDir)
Say "  $Ref = $commit"
if ($Release) {
    # The release workflow's "Stamp assembly version" - in the build tree only; the commit keeps 99.99.99.0.
    $assemblyInfo = Join-Path $sourceDir 'Properties\AssemblyInfo.cs'
    $infoBytes = [System.IO.File]::ReadAllBytes($assemblyInfo)
    $infoHasBom = ($infoBytes.Length -ge 3 -and $infoBytes[0] -eq 0xEF -and $infoBytes[1] -eq 0xBB -and $infoBytes[2] -eq 0xBF)
    $stampedInfo = Set-AssemblyInfoVersion -Text ([System.IO.File]::ReadAllText($assemblyInfo)) -Version $Version
    if ($null -eq $stampedInfo) { throw "Properties/AssemblyInfo.cs at $commit has no AssemblyVersion or no AssemblyFileVersion to stamp." }
    [System.IO.File]::WriteAllText($assemblyInfo, $stampedInfo, (New-Object System.Text.UTF8Encoding($infoHasBom)))
    Say "  Properties/AssemblyInfo.cs stamped $Version in the build tree"
}

$project = Join-Path $sourceDir 'OutlookAI.csproj'
if (-not (Test-Path -LiteralPath $project)) { throw "The archive has no OutlookAI.csproj at its root - $sourceDir" }
$addinSources = @(Get-ChildItem -LiteralPath $sourceDir -Filter '*.cs' -File | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) +
    @(foreach ($d in @('Services', 'TaskPane')) { Get-ChildItem -LiteralPath (Join-Path $sourceDir $d) -Filter '*.cs' -File -ErrorAction SilentlyContinue | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw } })
$regions = Find-FormRegionDeclarations -SourceTexts $addinSources
if ($regions.Count -gt 0) {
    throw "REFUSING TO BUILD: this commit declares an Outlook form region ($($regions -join ', ')). The RegisterFormRegions stand-in would then drop it from the add-in manifest, and the guest would test a different add-in than the one released. Teach this script about form regions first."
}

Copy-Item -LiteralPath $VstoRuntimePath -Destination (New-Item -ItemType Directory -Force -Path (Join-Path $sourceDir 'Redist')).FullName -Force

# ---------------------------------------------------------------------------------------------
Say ''
Say '== Guard 3, before: snapshotting the host =='
$before = Get-HostSnapshot
Say "  $(@($before).Count) line(s) across $($WatchedRegistry.Count) HKCU key(s) and $($WatchedCertStores.Count) certificate store(s)"

Set-Content -LiteralPath $standInPath -Value (Get-NoHostWriteTargets) -Encoding UTF8

$cert = $null
$thumb = $null
$built = $false
$buildProblems = @()
$signingKeyXml = $null
try {
    Say ''
    if ($Release) {
        Say '== The release signing certificate - already here; nothing is created and nothing deleted =='
        $cert = $releaseCert
        $thumb = [string]$cert.Thumbprint
        Say "  Cert:\CurrentUser\My\$thumb  $($cert.Subject), valid until $($cert.NotAfter.ToString('yyyy-MM-dd'))"
    }
    else {
        Say '== A throwaway signing certificate, for this build only =='
        $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=OutlookAI Testbed' `
            -FriendlyName "OutlookAI testbed build $commit - throwaway, deleted after signing" `
            -CertStoreLocation 'Cert:\CurrentUser\My' -NotAfter (Get-Date).AddYears(10)
        $thumb = [string]$cert.Thumbprint
        Say "  Cert:\CurrentUser\My\$thumb"
    }

    $msbuildArgs = Get-MSBuildArgumentList -ProjectPath $project -Thumbprint $thumb -StandInTargets $standInPath `
        -ApplicationVersion $Version -FileLog $msbuildLog

    Say ''
    Say '== Guard 1, proven before anything runs: MSBuild evaluates the chain =='
    # The same arguments minus the target, the logger and the verbosity, plus -getProperty:
    # evaluation only, no target executes.
    $evalArgs = @($msbuildArgs | Where-Object { $_ -ne '/t:Publish' -and -not $_.StartsWith('/flp:') -and $_ -ne '/v:minimal' }) + '-getProperty:PrepareForRunDependsOn'
    $eval = Invoke-Logged -FilePath $msbuild -ArgumentList $evalArgs -LogStem (Join-Path $OutDir 'msbuild-evaluate') -TimeoutMinutes 5 -WorkingDirectory $sourceDir
    $evaluated = ''
    if (Test-Path -LiteralPath $eval.Out) { $evaluated = (Get-Content -LiteralPath $eval.Out -Raw) }
    $chainProblems = @()
    if ($eval.ExitCode -ne 0) { $chainProblems += "MSBuild could not evaluate the project (exit $($eval.ExitCode)); see $($eval.Out)" }
    else { $chainProblems += @(Test-EvaluatedPrepareForRun $evaluated) }
    if ($chainProblems.Count -gt 0) {
        Show-LogTail $eval.Out 20
        throw ('REFUSING TO BUILD: ' + ($chainProblems -join ' '))
    }
    Say "  PrepareForRunDependsOn = $(($evaluated -replace '\s', ''))"

    Say ''
    Say '== Build (the release''s "Build and Publish", with guards 1 and 2) =='
    Say "  msbuild $($msbuildArgs -join ' ')"
    $build = Invoke-Logged -FilePath $msbuild -ArgumentList $msbuildArgs -LogStem (Join-Path $OutDir 'msbuild-console') -TimeoutMinutes $BuildTimeoutMinutes -WorkingDirectory $sourceDir
    if ($build.TimedOut -or $build.ExitCode -ne 0) {
        Show-LogTail $build.Out 40
        $buildProblems += "MSBuild exited $($build.ExitCode) (timed out: $($build.TimedOut)). Full log: $msbuildLog"
    }
    else {
        $logText = ''
        if (Test-Path -LiteralPath $msbuildLog) { $logText = Get-Content -LiteralPath $msbuildLog -Raw }
        $sentinels = Test-StandInSentinels -LogText $logText
        foreach ($n in $sentinels.Notes) { Say "  NOTE $n" }
        if (-not $sentinels.Ok) { $buildProblems += $sentinels.Problems }
        else { Say '  both scheduled stand-ins logged; the registration target was not scheduled' }

        # The release workflow's "Flatten VSTO payload next to the manifest (|vstolocal)", line for line.
        $publish = Join-Path $sourceDir 'publish'
        $appFiles = Join-Path $publish 'Application Files'
        $verDirs = @(Get-ChildItem -LiteralPath $appFiles -Directory -ErrorAction SilentlyContinue)
        if ($verDirs.Count -ne 1) { $buildProblems += "Expected exactly one versioned folder under $appFiles, found $($verDirs.Count)." }
        else {
            Get-ChildItem -LiteralPath $verDirs[0].FullName -File | Where-Object { $_.Name -ne 'OutlookAI.vsto' } | ForEach-Object {
                Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $publish ($_.Name -replace '\.deploy$', '')) -Force
            }
            foreach ($required in @('OutlookAI.vsto', 'OutlookAI.dll.manifest', 'OutlookAI.dll')) {
                if (-not (Test-Path -LiteralPath (Join-Path $publish $required))) { $buildProblems += "Missing after flatten: publish\$required" }
            }
            # The release workflow's check: the assembly next to the manifest is the one the registration
            # loads, so any version but the release's is a dead add-in.
            if ($Release -and (Test-Path -LiteralPath (Join-Path $publish 'OutlookAI.dll'))) {
                $dllVersion = (Get-Item -LiteralPath (Join-Path $publish 'OutlookAI.dll')).VersionInfo.FileVersion
                if ($dllVersion -ne $Version) { $buildProblems += "The flattened OutlookAI.dll is $dllVersion, not $Version." }
                else { Say "  publish\OutlookAI.dll carries $dllVersion" }
            }
            # The PUBLIC half only, DER - what the release workflow's Export-Certificate -Type CERT wrote. The
            # installer imports it into the guest user's TrustedPublisher store, as it does for users.
            [System.IO.File]::WriteAllBytes((Join-Path $publish 'OutlookAI.cer'),
                $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
            $signingKeyXml = Get-ManifestSigningKeyXml -ManifestXml (Get-Content -LiteralPath (Join-Path $publish 'OutlookAI.vsto') -Raw)
            $built = ($buildProblems.Count -eq 0)
        }
    }
}
finally {
    # Never the release key: it was here before the build and stays.
    if ($thumb -and -not $Release) { Remove-ThrowawayCertificate -Thumbprint $thumb }
}

# ---------------------------------------------------------------------------------------------
Say ''
Say '== Guard 3, after: nothing of this build may remain on the host =='
$after = Get-HostSnapshot
$modulus = $null
if ($signingKeyXml) { $modulus = [regex]::Match($signingKeyXml, '<Modulus>([^<]+)</Modulus>').Groups[1].Value }
$footprints = @(Find-HostTraces -Before $before -After $after -BuildDirectory $OutDir -Thumbprint $thumb -ModulusBase64 $modulus -KeyExistedBefore:$Release)
$regBefore = @(Select-KeyLines -Lines $before -KeyPath $AddinRegistrationKey)
$regAfter = @(Select-KeyLines -Lines $after -KeyPath $AddinRegistrationKey)
$regDiff = Compare-HostSnapshot -Before $regBefore -After $regAfter
$allDiff = Compare-HostSnapshot -Before $before -After $after
$hostFailures = @()
if ($footprints.Count -gt 0) {
    $hostFailures += 'traces of this build are on the host:'
    $hostFailures += @($footprints | ForEach-Object { '    ' + $_ })
}
if (-not $regDiff.Same) {
    $hostFailures += "the OutlookAI add-in registration CHANGED during the build (HKCU\$AddinRegistrationKey):"
    $hostFailures += @($regDiff.Removed | ForEach-Object { '    was: ' + $_ })
    $hostFailures += @($regDiff.Added | ForEach-Object { '    now: ' + $_ })
    Restore-AddinRegistration -BeforeLines $before
}
if ($hostFailures.Count -gt 0) {
    foreach ($l in $hostFailures) { Say "  FAIL $l" }
    throw 'REFUSING TO STAGE: the host was touched by the build (above). Nothing was packaged.'
}
if (-not $allDiff.Same) {
    Say '  WARN something else under the watched keys changed while the build ran. Nothing of this build is in'
    Say '       it, so the likeliest writer is the maintainer''s own Outlook; read it anyway:'
    foreach ($l in $allDiff.Removed) { Say "         - $l" }
    foreach ($l in $allDiff.Added) { Say "         + $l" }
}
else {
    Say "  UNCHANGED: $(@($after).Count) line(s), identical before and after"
}
if (-not $built) {
    foreach ($p in $buildProblems) { Say "  FAIL $p" }
    throw 'The build did not produce a usable add-in (above). The host is untouched; nothing was packaged.'
}

# ---------------------------------------------------------------------------------------------
if ($Release) {
    Say ''
    Say '== MCP server (the release''s "Publish MCP server", into the same tree) =='
    $serverOut = Join-Path $sourceDir 'publish\McpServer'
    $serverProject = Join-Path $sourceDir 'McpServer\OutlookAI.McpServer\OutlookAI.McpServer.csproj'
    $publishArgs = Get-ServerPublishArgumentList -ProjectPath $serverProject -Version $Version -OutputDir $serverOut
    Say "  dotnet $($publishArgs -join ' ')"
    $serverPublish = Invoke-Logged -FilePath $dotnet -ArgumentList $publishArgs -LogStem (Join-Path $OutDir 'dotnet-publish') -TimeoutMinutes $BuildTimeoutMinutes -WorkingDirectory $sourceDir
    if ($serverPublish.TimedOut -or $serverPublish.ExitCode -ne 0) {
        Show-LogTail $serverPublish.Out 40
        throw "dotnet publish of the MCP server exited $($serverPublish.ExitCode) (timed out: $($serverPublish.TimedOut))."
    }
    # Both executables: the COM host rides in by ProjectReference, and nothing else would notice
    # one that went missing or unstamped - the installer's mail tools would fail on their first call.
    foreach ($name in $ServerExecutables) {
        $exe = Join-Path $serverOut $name
        if (-not (Test-Path -LiteralPath $exe)) { throw "$name is missing from the MCP server publish output." }
        $stamped = (Get-Item -LiteralPath $exe).VersionInfo.FileVersion
        if ($stamped -ne $Version) { throw "$name is stamped $stamped, not $Version." }
        Say "  $name $stamped"
    }
    $payloadMb = [math]::Round((Get-ChildItem -LiteralPath (Join-Path $sourceDir 'publish') -Recurse -File | Measure-Object Length -Sum).Sum / 1MB, 2)
    Say "  total uncompressed payload $payloadMb MB"
}

# ---------------------------------------------------------------------------------------------
Say ''
Say '== Installer (the release''s "Create installer", from the same tree) =='
New-Item -ItemType Directory -Force -Path $installerDir | Out-Null
$isccArgs = @("/DMyAppVersion=$Version", ('"/O' + $installerDir + '"'), ('"' + (Join-Path $sourceDir 'Installer.iss') + '"'))
$compile = Invoke-Logged -FilePath $iscc -ArgumentList $isccArgs -LogStem (Join-Path $OutDir 'iscc') -TimeoutMinutes 10 -WorkingDirectory $sourceDir
$installer = Join-Path $installerDir "OutlookAI-v$Version.exe"
if ($compile.TimedOut -or $compile.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $installer)) {
    Show-LogTail $compile.Out 40
    throw "ISCC exited $($compile.ExitCode) and did not produce $installer."
}
Say "  $installer  $([math]::Round((Get-Item -LiteralPath $installer).Length / 1MB, 1)) MB"

if ($Release) {
    Say ''
    Say "RELEASE BUILD DONE, UNSIGNED: $installer"
    Say "  sha256 $((Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash)"
    Say '  Signing and publishing it are Tools/Publish-Release.ps1''s. No guest manifest and no AddIn.zip were made.'
    exit 0
}

# ---------------------------------------------------------------------------------------------
Say ''
Say '== Manifest and package =='
$publish = Join-Path $sourceDir 'publish'
$manifest = [ordered]@{
    schema      = 1
    producedBy  = 'Testbed/host/Publish-AddInPayload.ps1'
    commit      = $commit
    ref         = $Ref
    version     = $Version
    builtUtc    = [DateTime]::UtcNow.ToString('o')
    installer   = [ordered]@{
        file   = [System.IO.Path]::GetFileName($installer)
        sha256 = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
        bytes  = (Get-Item -LiteralPath $installer).Length
    }
    addin       = [ordered]@{
        'OutlookAI.dll'          = (Get-FileHash -LiteralPath (Join-Path $publish 'OutlookAI.dll') -Algorithm SHA256).Hash
        'OutlookAI.vsto'         = (Get-FileHash -LiteralPath (Join-Path $publish 'OutlookAI.vsto') -Algorithm SHA256).Hash
        'OutlookAI.dll.manifest' = (Get-FileHash -LiteralPath (Join-Path $publish 'OutlookAI.dll.manifest') -Algorithm SHA256).Hash
    }
    signing     = [ordered]@{
        subject          = 'CN=OutlookAI Testbed'
        thumbprint       = $thumb
        publicKeyXml     = $signingKeyXml
        privateKeyExists = $false
    }
    vstoRuntime = [ordered]@{
        file   = 'vstor_redist.exe'
        sha256 = $VstoRuntimeSha256
        bytes  = $VstoRuntimeBytes
    }
    # The files the add-in and the mail server agree through, as this commit has them. The guest
    # compares them with the suite it will run: equal means the tests read the names this add-in
    # writes, whatever commits the two were built from.
    contract    = [ordered]@{}
    toolchain   = [ordered]@{
        msbuild = [string]$msbuildVersion
        iscc    = (Split-Path -Leaf (Split-Path -Parent $iscc))
    }
    hostUntouched = [ordered]@{
        linesCompared = @($after).Count
        footprints    = 0
        registrationUnchanged = $true
    }
}
foreach ($rel in $ContractFiles) {
    $manifest.contract[$rel] = (Get-FileHash -LiteralPath (Join-Path $sourceDir ($rel.Replace('/', '\'))) -Algorithm SHA256).Hash
}
$shape = Test-PayloadManifestShape ([pscustomobject]($manifest | ConvertTo-Json -Depth 6 | ConvertFrom-Json))
if ($shape.Count -gt 0) { throw ('The manifest this script just built is malformed: ' + ($shape -join '; ')) }

New-Item -ItemType Directory -Force -Path $stageDir | Out-Null
Copy-Item -LiteralPath $installer -Destination $stageDir -Force
Set-Content -LiteralPath (Join-Path $stageDir $ManifestFileName) -Value ($manifest | ConvertTo-Json -Depth 6) -Encoding UTF8
if (Test-Path -LiteralPath $addinZip) { Remove-Item -LiteralPath $addinZip -Force }
# The stage folder's CONTENTS at the root of the zip, as Compress-Archive of stage\* made them.
[System.IO.Compression.ZipFile]::CreateFromDirectory($stageDir, $addinZip)
Say "  AddIn.zip  $([math]::Round((Get-Item -LiteralPath $addinZip).Length / 1MB, 1)) MB  (commit $commit)"

# ---------------------------------------------------------------------------------------------
Say ''
Say 'Copy in - one machine at a time, and say which (Testbed/README.md section 4a):'
Say "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $addinZip -Destination C:\OutlookAI-Q5\AddIn.zip"
Say "  pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path $VstoRuntimePath -Destination $GuestMediaRoot\vstor_redist.exe"
Say ''
Say 'Then on the guest - PowerShell Direct is fine for the unpacking, NOT for the two phases, which run in'
Say 'session 1: the install elevated (the task''s default), the first run NOT elevated (Testbed/README.md section 1,'
Say 'steps 5b and 7c - on the unindexed guest, the first run only once its index exclusion is in place):'
Say "  Expand-Archive C:\OutlookAI-Q5\AddIn.zip -DestinationPath $GuestAddInRoot -Force"
Say "  .\Register-InteractiveTask.ps1 -Script ""& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase Install -Execute"""
Say "  .\Register-InteractiveTask.ps1 -RunLevel Limited -Script ""& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase FirstRun -Execute"""
Say ''
Say 'BOTH GUESTS need it: Phase7LiveMcpToolShapeTests.Health_OverStdio_... declares only Requires=AddInRegistry,'
Say 'so it runs - and fails without the add-in - on the unindexed guest as well as the indexed one.'
