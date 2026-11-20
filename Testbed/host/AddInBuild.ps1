#Requires -Version 5.1
<#
.SYNOPSIS
    The add-in build's three guards and the helpers around them, in one copy for the two scripts that
    may build the add-in (Q81 rule 2): Testbed/host/Publish-AddInPayload.ps1 and Tools/Switch-AddInBuild.ps1.
    Not run on its own: dot-sourced by both, after Testbed/guest/ScriptHelpers.ps1, whose
    Invoke-NativeCommand and Say it uses.

.DESCRIPTION
    WHY ONE FILE (Q154, the survey's P07). The two scripts restated 21 functions, 6 of them word for
    word and 15 drifted apart: the registering behaviour of Visual Studio was pinned off in one build's
    MSBuild command line and not in the other's, and one script's comparison of its stand-ins with the
    installed VSTO targets (D7 (c)) checked three things fewer than the other's. Each now builds the same
    way, compares the same way and proves the same way, from here.

        Testbed/host/*.ps1     . (Join-Path $PSScriptRoot 'AddInBuild.ps1')
        Tools/*.ps1            . (Join-Path $PSScriptRoot '..\Testbed\host\AddInBuild.ps1')

    WHAT IT DEFINES WHEN DOT-SOURCED, besides its functions: the four values the guards rest on, which
    the two scripts declared alike - $PrepareForRunChain (guard 1), $StandInTasks (guard 2), and
    $WatchedRegistry and $WatchedCertStores (guard 3). WHAT EACH SCRIPT KEEPS: its stand-ins' sentinel,
    $StandInSentinel, which tells its build log's lines from the other's; its throwaway certificate's
    subject; its -MSBuildPath, which Resolve-MSBuild reads; and the names it hands the functions here -
    its own path, for the stand-in file it writes and for -CompareInstalledTargets, and the parameter
    under which every path it hands MSBuild lies.

    It builds nothing and writes nothing when dot-sourced. Tools/Checks/check-powershell-51.ps1 check 4
    holds every function in it to one copy. ASCII only, so Windows PowerShell 5.1 reads it the same
    without a byte order mark.
#>

# ---------------------------------------------------------------------------------------------
# What the guards rest on.
# ---------------------------------------------------------------------------------------------

# The chain the VSTO targets would have given PrepareForRun, minus RegisterOfficeAddin. The
# common targets' own default is CopyFilesToOutputDirectory; the VSTO targets append
# VisualStudioForApplicationsBuild and then, for an add-in, RegisterOfficeAddin.
$PrepareForRunChain = @('CopyFilesToOutputDirectory', 'VisualStudioForApplicationsBuild')

# The three VSTO build tasks that write the build machine's registry, and every parameter the
# targets pass each of them. An inline stand-in must declare all of them or MSBuild refuses the
# call (MSB4064) - which is loud, and the self-tests check the list against the installed targets.
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

# ---------------------------------------------------------------------------------------------
# Guard 1: RegisterOfficeAddin is not scheduled.
# ---------------------------------------------------------------------------------------------

function Get-PrepareForRunOverride {
    return ($PrepareForRunChain -join ';')
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

# ---------------------------------------------------------------------------------------------
# Guard 2: the three writing tasks replaced by stand-ins that only say they did not run.
# ---------------------------------------------------------------------------------------------

# The generated targets file. Only the task NAMES and parameters come from outside; the body of
# every stand-in is a single log line, so there is nothing in it that can write anything. -WrittenBy
# is the script that writes it, named in its header; the stand-ins log the caller's $StandInSentinel.
function Get-NoHostWriteTargets {
    param([Parameter(Mandatory = $true)] [string] $WrittenBy)
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
    WRITTEN BY $WrittenBy FOR ONE BUILD. NOT PART OF THE REPOSITORY.

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

# ---------------------------------------------------------------------------------------------
# The MSBuild command line.
# ---------------------------------------------------------------------------------------------

# One argument string for MSBuild.exe. Values with spaces are quoted. A semicolon inside a value
# needs one of TWO spellings, and they are not interchangeable:
#   * a SCALAR (DefineConstants) takes %3B, the release workflow's spelling - the value is unescaped when a
#     task reads it;
#   * a LIST (PrepareForRunDependsOn) needs a REAL semicolon inside double quotes, because a target
#     list is split on semicolons BEFORE it is unescaped. MEASURED 2026-09-24, first run of
#     Testbed/host/Publish-AddInPayload.ps1: %3B there made MSBuild look for ONE target called
#     "CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild" and stop with MSB4057.
function Format-MSBuildProperty {
    param([string] $Name, [string] $Value, [switch] $List)
    if ($List) { return "/p:$Name=`"$Value`"" }
    $v = $Value.Replace(';', '%3B')
    if ($v -match '\s') { return "/p:$Name=`"$v`"" }
    return "/p:$Name=$v"
}

# The release workflow's "Build and Publish", property for property - since 2026-10-03 the release's
# own build is Publish-AddInPayload.ps1's (-ReleaseSigningThumbprint) - with guards 1 and 2. Both
# builds pin BuildingInsideVisualStudio to false: inside Visual Studio OutlookAI.csproj lets the VSTO
# tasks register the build (Q81), so a value from the environment could otherwise turn that back on.
# Switch-AddInBuild.ps1's build pinned it; Publish-AddInPayload.ps1's did not, until 2026-11-20.
function Get-MSBuildArgumentList {
    param(
        [Parameter(Mandatory = $true)] [string] $ProjectPath,
        [Parameter(Mandatory = $true)] [string] $Thumbprint,
        [Parameter(Mandatory = $true)] [string] $StandInTargets,
        [Parameter(Mandatory = $true)] [string] $ApplicationVersion,
        [Parameter(Mandatory = $true)] [string] $FileLog,
        # The release build's -Version, for the assemblies MSBuild stamps from a property - Core's.
        # The add-in's own comes from Properties/AssemblyInfo.cs, which the release stamps itself.
        [string] $AssemblyVersion
    )

    $projectArg = $ProjectPath
    if ($projectArg -match '\s') { $projectArg = '"' + $projectArg + '"' }

    $list = @(
        $projectArg,
        '/t:Publish',
        # OutlookAI.Core's packages (Q167 (a)): restored in an evaluation of its own before the
        # build's - MSBuild's documented switch, not a Restore target in the same evaluation.
        '/restore',
        (Format-MSBuildProperty 'Configuration' 'Release'),
        (Format-MSBuildProperty 'ApplicationVersion' $ApplicationVersion),
        (Format-MSBuildProperty 'PublishDir' 'publish\'),
        (Format-MSBuildProperty 'BootstrapperEnabled' 'true'),
        (Format-MSBuildProperty 'IsWebBootstrapper' 'false'),
        (Format-MSBuildProperty 'DefineConstants' 'VSTO40;TRACE'),
        # The throwaway key - or, for a release, the release key (-ReleaseSigningThumbprint).
        (Format-MSBuildProperty 'ManifestCertificateThumbprint' $Thumbprint),
        # Guard 1 and guard 2, and never the IDE's registering behaviour, whatever the environment says.
        (Format-MSBuildProperty 'PrepareForRunDependsOn' (Get-PrepareForRunOverride) -List),
        (Format-MSBuildProperty 'CustomBeforeMicrosoftCommonTargets' $StandInTargets),
        (Format-MSBuildProperty 'BuildingInsideVisualStudio' 'false'),
        # No compiler server and no reused nodes: nothing may outlive the build and hold a handle.
        (Format-MSBuildProperty 'UseSharedCompilation' 'false'),
        '/nodeReuse:false',
        '/m:1',
        '/nologo',
        '/v:minimal'
    )

    if ($AssemblyVersion) { $list += (Format-MSBuildProperty 'Version' $AssemblyVersion) }

    # DETAILED, because the proof reads which TARGETS ran, and only detailed names every one.
    $flp = "LogFile=$FileLog;Verbosity=detailed;Encoding=UTF-8"
    if ($flp -match '\s') { $flp = '"' + $flp + '"' }
    $list += "/flp:$flp"
    return $list
}

# Guard 1's evaluation: the build's own arguments - so the same global properties - minus everything
# that would make it more than an evaluation: the target, the restore, the logger and the verbosity.
# -getProperty with no target evaluates the project and runs nothing; /restore would run the Restore
# target first.
function Get-MSBuildEvaluationArgumentList {
    param([Parameter(Mandatory = $true)] [string[]] $BuildArguments)
    $kept = @($BuildArguments | Where-Object {
            -not $_.StartsWith('/t:') -and $_ -ne '/restore' -and -not $_.StartsWith('/flp:') -and $_ -ne '/v:minimal'
        })
    return @($kept + '-getProperty:PrepareForRunDependsOn')
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
#   Tools/Switch-AddInBuild.ps1, measured the same day with -BuildOnly under 5.1 started from 7 and
#   under 7, met both: a ',' stopped guard 1's evaluation on MSB1006 after the throwaway certificate
#   was made and the source archived, and with a ';' the build succeeded but its detailed log went to
#   the cut-off file, where all three log checks failed the good build. The dry run printed both paths
#   as where -Execute would build, and exited 0.
# Every path a build script hands MSBuild - the project, the stand-in targets, the log - lies under
# one of its parameters, -ParameterName, so a ',' or ';' there is refused up front, before anything
# is archived, snapshotted, signed or built, rather than every argument having to get an escaping
# rule right. -Where says where that parameter points by default, when it has one. $null when the
# path is fine.
function Get-MSBuildPathProblem {
    param([string] $Path, [string] $ParameterName = '-OutDir', [string] $Where = '')
    $separators = @(@(',', ';') | Where-Object { $Path.Contains($_) })
    if ($separators.Count -eq 0) { return $null }
    $named = ($separators | ForEach-Object { "'$_'" }) -join ' and '
    return "$ParameterName '$Path' contains $named, which MSBuild's command line reads as a separator - a ',' stops the build on MSB1006 ('Property is not valid'), a ';' sends the detailed log this script proves its guards from somewhere else. Every path this build hands MSBuild lies under $ParameterName$Where; choose one with no ',' or ';' in it"
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

# The MSBuild of a Visual Studio with the Office/SharePoint development workload - or the calling
# script's -MSBuildPath, when it names one.
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

# ---------------------------------------------------------------------------------------------
# Guard 3: the host, snapshotted before and after.
# ---------------------------------------------------------------------------------------------

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

# The throwaway certificate, out of EVERY CurrentUser store, private key included.
#
# NOT ONLY Cert:\CurrentUser\My. MEASURED 2026-09-24, first run of Testbed/host/Publish-AddInPayload.ps1
# (pwsh 7.6.6, Windows 11 10.0.26200): New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My
# ALSO puts a copy of the certificate into Cert:\CurrentUser\CA, Intermediate Certification Authorities.
# Removing it from My left that copy behind, and guard 3 caught it as a trace of the build. Only a
# certificate carrying BOTH this build's thumbprint and the -Subject the calling script makes its
# throwaway certificates with is removed, so nothing of the maintainer's can match.
function Remove-ThrowawayCertificate {
    param([Parameter(Mandatory = $true)] [string] $Thumbprint, [Parameter(Mandatory = $true)] [string] $Subject)
    $my = "Cert:\CurrentUser\My\$Thumbprint"
    if (Test-Path -LiteralPath $my) {
        Remove-Item -LiteralPath $my -DeleteKey -ErrorAction SilentlyContinue
    }
    foreach ($s in $WatchedCertStores) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $mine = @($store.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint -and $_.Subject -eq $Subject })
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

# ---------------------------------------------------------------------------------------------
# D7 (c): the stand-ins against Visual Studio's own VSTO targets.
# ---------------------------------------------------------------------------------------------

# Guard 2's stand-ins against the text of one VSTO targets file. Each result is one assertion -
# What, Expected, Actual. Pure, so a self-test feeds it synthetic targets anywhere and
# -CompareInstalledTargets the real ones: every parameter the targets pass a writing task is one its
# stand-in declares, the targets still call each writing task, and they still put RegisterOfficeAddin
# and VisualStudioForApplicationsBuild on PrepareForRun - what guard 1 assumes.
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

# -CompareInstalledTargets: the comparison against EVERY installed VSTO targets file, as a verdict.
# 0 when each passes; 1 when one fails or when there is none - a comparison that did not run proves
# nothing, and the release needs it to have run (D7 (c)). -ScriptName is the calling script's, for the
# first line.
function Invoke-InstalledTargetsComparison {
    param([Parameter(Mandatory = $true)] [string] $ScriptName)
    Write-Host "$ScriptName -CompareInstalledTargets under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
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
