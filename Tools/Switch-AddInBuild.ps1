#Requires -Version 5.1
<#
    ============================================================================================
    WRITTEN AND RUN 2026-09-27 ON THE MAINTAINER'S WORKSTATION, FOR Q81. -Execute HAS NOT RUN.
    ============================================================================================

    Every run below was wrapped in a `reg export` of the add-in registration and of the VSTO trust
    list (plus VSTO\SolutionMetadata, every Outlook add-in, FormRegions, Software\OutlookAI, the
    slow-add-in exemption and six certificate stores) before and after; every run left them
    byte-identical.

      * -SelfTest: 121 assertions, 0 failures, under Windows PowerShell 5.1.26100 and PowerShell
        7.6.6. Every registry write this script can make ran against a scratch key under
        HKCU\Software\OutlookAI-SelfTest\<run id>, which it then deleted; the real registration,
        trust, uninstall and exemption keys (46 lines) were identical before and after.
      * -Status, and the dry runs of -Restore (with and without -RemoveBuildTrust) and of the
        default mode. They only read. -Status found Outlook registered at the main checkout's
        bin\Release, a build of 2026-08-24, not at the installed 3.0.1.321.
      * -BuildOnly from 4e23866, a commit from BEFORE the project's own Q81 guard - so this
        script's guards alone stood between a registering build and the host: built in 3 s, both
        scheduled stand-ins logged, the throwaway certificate removed from My and from CA, 358
        host lines identical. The first two attempts refused before building, both correctly and
        both now fixed: 5.1 started by Start-Process from PowerShell 7 had no Cert: drive, and a
        plain Import-Module of the Security module failed there too (see the Import-Module below).
      * -BuildOnly from 9f24e5c, which has the guard: this script's Override stand-ins took over
        from the project's own ("Created an override using task"), with no conflict; 358 lines
        identical.

    NOT RUN: -Execute, in either direction - the maintainer decides when his Outlook changes - so
    nothing here has yet been loaded by Outlook. Replace this banner with what the first -Execute
    did, and whether Outlook then loaded the build with no trust prompt.

.SYNOPSIS
    Puts a development build of the OutlookAI add-in on THIS machine's Outlook when the maintainer
    asks for one, takes it off again, and says which build Outlook will load.

.DESCRIPTION
    WHY THIS EXISTS. Decided by the maintainer 2026-09-27 (Q81), in his words: "I want to be able
    to ask you to put the dev build on my machine if I want to test it for a release. I do not want
    it on my machine whilst agents are still actively building it." Until then any plain build of
    the add-in registered ITSELF with the Outlook on the machine that built it (see the comment in
    OutlookAI.csproj), so the maintainer's Outlook loaded whatever build folder had been built last.
    This script is the one sanctioned way to change what his Outlook loads - CLAUDE.md, "The add-in
    on the maintainer's workstation".

    FOUR MODES. Dry run by default: nothing changes without -Execute.

      (default) -Commit <rev>   Put a dev build on. HEAD when -Commit is not given.
          1. Builds the add-in AT THAT COMMIT - a `git archive`, never the working tree - with the
             build Testbed/host/Publish-AddInPayload.ps1 measured: release.yml's Publish,
             signed with a throwaway certificate whose private key is deleted before the script
             goes on, the VSTO tasks that write the registry stood in, and the host's watched
             registry keys and certificate stores snapshotted before and after. Any trace of the
             build on the host is a failure, and a changed add-in registration is put back first.
             These guards do not rely on the commit's own OutlookAI.csproj, so a commit from
             before Q81 builds just as safely.
          2. Copies the flattened publish output into a NEW folder,
             %LOCALAPPDATA%\OutlookAI\DevBuilds\<yyyyMMdd-HHmmss>-<commit>, which no build ever
             writes into.
          3. Records the registration it is about to replace in dev-build.json in that folder.
          4. Trusts it: one VSTO inclusion entry, the one the trust prompt itself would write -
             HKCU\Software\Microsoft\VSTO\Security\Inclusion\<new guid> with Url = the manifest's
             file:/// URL and PublicKey = the build's signing key as <RSAKeyValue>. How the VSTO
             runtime forms, finds and compares these entries was read in its own IL; the entry is
             written exactly as Testbed/guest/Install-OutlookAIAddIn.ps1 writes it, which a guest
             run proved loads with no prompt.
          5. Registers it the way the installer registers a release: every value Installer.iss
             writes under HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI (Manifest =
             file:///<folder>\OutlookAI.vsto|vstolocal, LoadBehavior 3, FriendlyName,
             Description), and the Resiliency\DoNotDisableAddinList values. They are READ OUT OF
             Installer.iss at run time rather than restated here, so the two cannot drift.
      -Restore      Puts the installed release back: the same Installer.iss values with {app} =
                    the InstallLocation of the release's uninstall entry, whatever is registered
                    now and whether or not this script registered it. Then removes the trust
                    entries of every folder under DevBuilds. -RemoveBuildTrust also removes the
                    trust entries of every OTHER OutlookAI.vsto - build folders such as
                    bin\Release and bin\Debug - keeping only the installed release's.
      -Status       Read-only: which build Outlook will load at its next start, its version and
                    path, whether it will load without a trust prompt, the installed release,
                    the dev builds and the OutlookAI trust entries.
      -BuildOnly    Step 1 alone, into .work. Registers nothing, copies nothing to DevBuilds, and
                    needs no -Execute: the build's own snapshot proves the host unchanged. For
                    anyone who only needs to know that a commit builds.
      -SelfTest     Every decision against synthetic inputs, and every registry write against a
                    scratch key it then deletes.

    WHAT IT NEVER DOES.
      * Never runs elevated: an elevated process can carry another account's HKCU.
      * Never starts, quits or kills Outlook. Outlook reads its add-in list when it starts, so a
        change here takes effect at the next Outlook restart - and the script says so.
      * Never touches the MCP server, the installed release's folder, or the server executable
        that is deliberately disabled there (TODO.md, "Restore the installed MCP server"). A dev
        build carries no server, and this script refuses a build output that does.
      * Never writes HKCU\Software\OutlookAI. The installer's InstallDir value there is how the
        add-in finds the installed server, so it keeps naming the installed release. The add-in
        looks for a server under InstallDir and then up to four folders above itself; neither
        finds one for a dev build today, so its Claude Code reconcile stands down exactly as the
        installed release's does while the server is disabled.
      * Never adds a certificate to a trust store. The installer imports the release certificate
        into TrustedPublisher; a dev build is signed by a throwaway key and is trusted through its
        own inclusion entry, which names one URL and one key.

    VERSION 99.99.99.0, what every developer build of this repository carries. It is higher than
    any release, so the add-in's own updater never decides a release is newer and replaces the
    build under test. The commit is the identity; dev-build.json records it.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`, ASCII only.

.PARAMETER Commit
    The commit to build. HEAD by default. Uncommitted edits are never in the build.

.PARAMETER Restore
    Put the installed release back (see above).

.PARAMETER Status
    Report what Outlook will load at its next start. Read-only.

.PARAMETER BuildOnly
    Build and prove the host unchanged; register nothing.

.PARAMETER SelfTest
    Run the self-test and exit.

.PARAMETER Execute
    Without it the default mode and -Restore are dry runs that only read and report.

.PARAMETER RemoveBuildTrust
    With -Restore: also remove the trust entries of OutlookAI build folders (every OutlookAI.vsto
    other than the installed release's and the dev builds', which go anyway).

.PARAMETER RepoRoot
    Repository root. Defaults to the folder above this script's.

.PARAMETER DevBuildsRoot
    Where dev builds are copied. Default %LOCALAPPDATA%\OutlookAI\DevBuilds.

.PARAMETER WorkRoot
    Where builds happen. Default .work\switch-addin-build under the repository root (gitignored).

.PARAMETER MSBuildPath
    MSBuild.exe. Default: found with vswhere, from a Visual Studio with the Office workload.

.EXAMPLE
    powershell -NoProfile -File Tools\Switch-AddInBuild.ps1 -Status
    powershell -NoProfile -File Tools\Switch-AddInBuild.ps1 -Commit 1a2b3c4             # dry run
    powershell -NoProfile -File Tools\Switch-AddInBuild.ps1 -Commit 1a2b3c4 -Execute
    powershell -NoProfile -File Tools\Switch-AddInBuild.ps1 -Restore -Execute
    powershell -NoProfile -File Tools\Switch-AddInBuild.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $Commit,
    [switch] $Restore,
    [switch] $Status,
    [switch] $BuildOnly,
    [switch] $SelfTest,
    [switch] $Execute,
    [switch] $RemoveBuildTrust,
    [string] $RepoRoot,
    [string] $DevBuildsRoot,
    [string] $WorkRoot,
    [string] $MSBuildPath,
    [int]    $BuildTimeoutMinutes = 20
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}
# The Cert: drive belongs to this module, and nothing here loads it on its own. Measured 2026-09-27,
# this script's first -BuildOnly: Windows PowerShell 5.1 started by Start-Process from PowerShell 7
# inherits 7's PSModulePath, so autoloading picked 7's copy of the module, which 5.1 cannot load -
# New-SelfSignedCertificate then failed on "A drive with the name 'Cert' does not exist", and a
# plain Import-Module on "the member AuditToString is already present". So the running edition's
# OWN copy is imported, by path.
Import-Module (Join-Path $PSHOME 'Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1')

# Defaults that need $PSScriptRoot or the environment are set HERE, not in param(): Windows
# PowerShell 5.1 leaves $PSScriptRoot empty while param() defaults are evaluated (Q78).
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
if (-not $DevBuildsRoot) { $DevBuildsRoot = Join-Path $env:LOCALAPPDATA 'OutlookAI\DevBuilds' }
if (-not $WorkRoot) { $WorkRoot = Join-Path $RepoRoot '.work\switch-addin-build' }
$CommitGiven = [bool]$Commit
if (-not $Commit) { $Commit = 'HEAD' }

# ---------------------------------------------------------------------------------------------
# Everything this script names, in one place.
# ---------------------------------------------------------------------------------------------
$AddinKeyRel      = 'Software\Microsoft\Office\Outlook\Addins\OutlookAI'
$InclusionKeyRel  = 'Software\Microsoft\VSTO\Security\Inclusion'
$UninstallRootRel = 'Software\Microsoft\Windows\CurrentVersion\Uninstall'
$SelfTestRootRel  = 'Software\OutlookAI-SelfTest'
$ManifestFileName = 'OutlookAI.vsto'
$RecordFileName   = 'dev-build.json'
$DevVersion       = '99.99.99.0'
$ThrowawaySubject = 'CN=OutlookAI DevBuild'
$StandInSentinel  = 'DEVBUILD-NO-HOST-WRITE'

# The chain the VSTO targets give PrepareForRun, minus RegisterOfficeAddin - guard 1, as in
# Testbed/host/Publish-AddInPayload.ps1.
$PrepareForRunChain = @('CopyFilesToOutputDirectory', 'VisualStudioForApplicationsBuild')

# The three VSTO build tasks that write the build machine's registry, and every parameter the
# targets pass each of them - guard 2. An inline stand-in must declare all of them or MSBuild
# refuses the call (MSB4064); -SelfTest checks the list against the installed targets.
$StandInTasks = [ordered]@{
    'SetOffice2007AddInRegistration' = @{
        Parameters = @('Url', 'AddInName', 'OfficeApplication', 'FriendlyName', 'Description', 'LoadBehavior', 'Unregister', 'SolutionID', 'IsDocument')
        Output     = $null
        Scheduled  = $false
    }
    'SetInclusionListEntry' = @{
        Parameters = @('DeploymentManifestFullPath', 'CertificateThumbprint', 'Unregister')
        Output     = $null
        Scheduled  = $true
    }
    'RegisterFormRegions' = @{
        Parameters = @('AddInName', 'AssemblyName', 'OfficeApplication', 'Unregister')
        Output     = 'FormRegionNamesAndMessageClasses'
        Scheduled  = $true
    }
}

# Guard 3's view of the host: HKCU keys and CurrentUser certificate stores.
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

$script:SelfTestActive = $false

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# A native program whose stderr is redirected runs through here. Under Windows PowerShell 5.1,
# with 'Stop', the first line a native program writes to a redirected stderr is a terminating
# NativeCommandError; 'Continue' holds only inside this function, the try keeps "program not
# found" terminating, and callers judge the call by $LASTEXITCODE. Restated from the repository's
# other scripts (.github/scripts/check-powershell-51.ps1 explains the rule).
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
# THE REGISTRY, THROUGH ONE DOOR. Every path is relative to HKCU. The self-test points a layout at
# a scratch key, and while it runs Assert-WritableKey refuses any write outside that key.
# =============================================================================================

function New-Layout {
    param([string] $ScratchRoot)
    $prefix = ''
    if ($ScratchRoot) { $prefix = $ScratchRoot.TrimEnd('\') + '\' }
    return [pscustomobject]@{
        Prefix        = $prefix
        AddinKey      = $prefix + $AddinKeyRel
        InclusionKey  = $prefix + $InclusionKeyRel
        UninstallRoot = $prefix + $UninstallRootRel
    }
}

function Assert-WritableKey {
    param([string] $RelPath)
    if (-not $RelPath) { throw 'Refusing a registry write with no key path.' }
    if ($script:SelfTestActive -and -not $RelPath.StartsWith($SelfTestRootRel + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "SELF-TEST SAFETY: refused a write to HKCU\$RelPath, which is outside HKCU\$SelfTestRootRel."
    }
}

# $null for a missing key; otherwise name -> Kind/Data, in the key's own order.
function Get-KeyValues {
    param([string] $RelPath)
    $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($RelPath, $false)
    if ($null -eq $k) { return $null }
    try {
        $out = [ordered]@{}
        foreach ($n in $k.GetValueNames()) {
            $out[$n] = [pscustomobject]@{
                Kind = [string]$k.GetValueKind($n)
                Data = $k.GetValue($n, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            }
        }
        return $out
    }
    finally { $k.Close() }
}

function Set-KeyValue {
    param([string] $RelPath, [string] $Name, [string] $Kind, $Data)
    Assert-WritableKey $RelPath
    $k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($RelPath)
    try {
        if ($Kind -eq 'DWord') { $k.SetValue($Name, [int]$Data, [Microsoft.Win32.RegistryValueKind]::DWord) }
        elseif ($Kind -eq 'String') { $k.SetValue($Name, [string]$Data, [Microsoft.Win32.RegistryValueKind]::String) }
        else { throw "Set-KeyValue: '$Kind' is not a kind this script writes." }
    }
    finally { $k.Close() }
}

function Remove-KeyValue {
    param([string] $RelPath, [string] $Name)
    Assert-WritableKey $RelPath
    $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($RelPath, $true)
    if ($null -eq $k) { return }
    try { $k.DeleteValue($Name, $false) } finally { $k.Close() }
}

function Remove-KeyTree {
    param([string] $RelPath)
    Assert-WritableKey $RelPath
    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($RelPath, $false)
}

function Format-RegValue {
    param($Value)
    if ($null -eq $Value) { return '<absent>' }
    return ('{0} ({1})' -f $Value.Data, $Value.Kind)
}

# =============================================================================================
# WHAT THE INSTALLER WRITES, READ OUT OF Installer.iss. Pure.
# =============================================================================================

# The [Registry] entries for the add-in key and for Resiliency\DoNotDisableAddinList, and the
# AppId that names the uninstall entry. Every other [Registry] line is SKIPPED ON PURPOSE and
# reported as such: today that is HKCU\Software\OutlookAI\InstallDir, which must keep naming the
# installed release (see WHAT IT NEVER DOES). A line this parser cannot read is a refusal, never
# a silent skip - a spec that quietly lost a value would register a different add-in.
function Get-InstallerSpec {
    param([string] $IssText)
    if (-not $IssText) { throw 'Installer.iss is empty or missing.' }
    $appId = [regex]::Match($IssText, '(?m)^\s*AppId\s*=\s*\{\{([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\}')
    if (-not $appId.Success) { throw 'Installer.iss has no AppId={{<guid>} line, so the uninstall entry cannot be found.' }
    $guid = '{' + $appId.Groups[1].Value.ToUpperInvariant() + '}'
    $section = ''
    $addin = @()
    $doNotDisable = @()
    $skipped = @()
    foreach ($raw in ($IssText -split "`r?`n")) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith(';')) { continue }
        $header = [regex]::Match($line, '^\[(\w+)\]$')
        if ($header.Success) { $section = $header.Groups[1].Value; continue }
        if ($section -ne 'Registry') { continue }
        $m = [regex]::Match($line, '^Root:\s*(\w+);\s*Subkey:\s*"([^"]*)";\s*ValueType:\s*(\w+);\s*ValueName:\s*"([^"]*)";\s*ValueData:\s*"([^"]*)"')
        if (-not $m.Success) { throw "Installer.iss [Registry] has a line this script cannot read, so it will not guess what the installer writes: $line" }
        $root = $m.Groups[1].Value
        $subkey = $m.Groups[2].Value
        $type = $m.Groups[3].Value.ToLowerInvariant()
        $name = $m.Groups[4].Value
        $data = $m.Groups[5].Value
        if ($root -ne 'HKCU') { throw "Installer.iss writes under $root ($subkey\$name); this script only knows HKCU." }
        $kind = $null
        if ($type -eq 'string') { $kind = 'String' }
        elseif ($type -eq 'dword') { $kind = 'DWord' }
        else { throw "Installer.iss writes a '$type' value ($subkey\$name); this script only knows string and dword." }
        $dnd = [regex]::Match($subkey, '^Software\\Microsoft\\Office\\(\d+\.\d+)\\Outlook\\Resiliency\\DoNotDisableAddinList$')
        if ([string]::Equals($subkey, $AddinKeyRel, [System.StringComparison]::OrdinalIgnoreCase)) {
            $addin += [pscustomobject]@{ Name = $name; Kind = $kind; Template = $data }
        }
        elseif ($dnd.Success) {
            $doNotDisable += [pscustomobject]@{ Version = $dnd.Groups[1].Value; Subkey = $subkey; Name = $name; Kind = $kind; Data = $data }
        }
        else {
            $skipped += [pscustomobject]@{ Subkey = $subkey; Name = $name }
        }
    }
    $manifest = @($addin | Where-Object { $_.Name -eq 'Manifest' })
    if ($manifest.Count -ne 1) { throw "Installer.iss: expected one Manifest value under $AddinKeyRel, found $($manifest.Count)." }
    if (-not $manifest[0].Template.Contains('{app}') -or -not $manifest[0].Template.EndsWith('|vstolocal')) {
        throw "Installer.iss: the Manifest value '$($manifest[0].Template)' is not the {app}...|vstolocal shape this script registers."
    }
    $load = @($addin | Where-Object { $_.Name -eq 'LoadBehavior' })
    if ($load.Count -ne 1 -or $load[0].Kind -ne 'DWord') { throw "Installer.iss: expected one DWORD LoadBehavior under $AddinKeyRel." }
    if ($doNotDisable.Count -eq 0) { throw 'Installer.iss: no Resiliency\DoNotDisableAddinList value - the file changed shape.' }
    return [pscustomobject]@{
        AppId            = $guid
        UninstallKeyName = $guid + '_is1'
        AddinValues      = $addin
        DoNotDisable     = $doNotDisable
        Skipped          = $skipped
    }
}

function Expand-SpecData {
    param([string] $Template, [string] $AppDir)
    $v = $Template.Replace('{app}', $AppDir.TrimEnd('\'))
    if ([regex]::IsMatch($v, '\{[A-Za-z]+\}')) { throw "Installer.iss value '$Template' uses an Inno Setup constant other than {app}." }
    return $v
}

# name -> Kind/Data: exactly what the installer would write for a payload in $AppDir.
function Get-TargetRegistration {
    param($Spec, [string] $AppDir)
    $t = [ordered]@{}
    foreach ($v in $Spec.AddinValues) {
        $data = Expand-SpecData -Template $v.Template -AppDir $AppDir
        if ($v.Kind -eq 'DWord') { $data = [int]$data }
        $t[$v.Name] = [pscustomobject]@{ Kind = $v.Kind; Data = $data }
    }
    return $t
}

function Get-RegistrationChanges {
    param($Current, $Target)
    $values = @()
    foreach ($name in $Target.Keys) {
        $want = $Target[$name]
        $have = $null
        if ($null -ne $Current -and $Current.Contains($name)) { $have = $Current[$name] }
        $same = ($null -ne $have -and $have.Kind -eq $want.Kind -and ([string]$have.Data) -ceq ([string]$want.Data))
        $values += [pscustomobject]@{ Name = $name; Kind = $want.Kind; Data = $want.Data; Before = (Format-RegValue $have); After = (Format-RegValue $want); Same = $same }
    }
    $extra = @()
    if ($null -ne $Current) {
        foreach ($name in $Current.Keys) {
            if (-not $Target.Contains($name)) { $extra += [pscustomobject]@{ Name = $name; Value = (Format-RegValue $Current[$name]) } }
        }
    }
    return [pscustomobject]@{ KeyExisted = ($null -ne $Current); Values = $values; Extra = $extra; Pending = @($values | Where-Object { -not $_.Same }).Count }
}

function Get-DoNotDisableChanges {
    param($Layout, $Spec)
    $out = @()
    foreach ($e in $Spec.DoNotDisable) {
        $rel = $Layout.Prefix + $e.Subkey
        $vals = Get-KeyValues $rel
        $have = $null
        if ($null -ne $vals -and $vals.Contains($e.Name)) { $have = $vals[$e.Name] }
        $want = $e.Data
        if ($e.Kind -eq 'DWord') { $want = [int]$e.Data }
        $same = ($null -ne $have -and $have.Kind -eq $e.Kind -and ([string]$have.Data) -eq ([string]$want))
        $out += [pscustomobject]@{ Key = $rel; Version = $e.Version; Name = $e.Name; Kind = $e.Kind; Data = $want; Before = (Format-RegValue $have); Same = $same }
    }
    return $out
}

# =============================================================================================
# PATHS, URLS AND KEYS. Pure. The URL and key rules are restated from
# Testbed/guest/Install-OutlookAIAddIn.ps1, which read them in the VSTO runtime's IL and proved
# them on a guest: an entry's Url is Uri.ToString() of the manifest location - forward slashes,
# file:///, spaces literal - entries are found by comparing Url AS A URI (case-insensitive on a
# DOS path), and keys are compared as numbers, so the XML shape is free.
# =============================================================================================

function ConvertFrom-ManifestValue {
    param([string] $Value)
    if (-not $Value) { return $null }
    $v = $Value.Trim()
    $bar = $v.IndexOf('|')
    if ($bar -ge 0) { $v = $v.Substring(0, $bar) }
    if ($v.StartsWith('file:///', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(8) }
    elseif ($v.StartsWith('file://', [System.StringComparison]::OrdinalIgnoreCase)) { $v = '\\' + $v.Substring(7) }
    $v = $v.Replace('/', '\')
    if ([regex]::IsMatch($v, '%[0-9A-Fa-f]{2}')) { $v = [System.Uri]::UnescapeDataString($v) }
    return $v
}

function ConvertTo-InclusionUrl {
    param([string] $ManifestValue)
    if (-not $ManifestValue) { return $null }
    $v = $ManifestValue.Trim()
    $bar = $v.IndexOf('|')
    if ($bar -ge 0) { $v = $v.Substring(0, $bar) }
    if ($v.StartsWith('file:///', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(8) }
    elseif ($v.StartsWith('file://', [System.StringComparison]::OrdinalIgnoreCase)) { $v = $v.Substring(7) }
    $v = $v.Replace('\', '/')
    return 'file:///' + $v
}

function Test-SameInclusionUrl {
    param([string] $A, [string] $B)
    $x = ConvertTo-InclusionUrl $A
    $y = ConvertTo-InclusionUrl $B
    if (-not $x -or -not $y) { return $false }
    return [string]::Equals($x, $y, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-NormalPath {
    param([string] $Path)
    if (-not $Path) { return '' }
    try { return [System.IO.Path]::GetFullPath($Path).TrimEnd('\') } catch { return $Path.TrimEnd('\') }
}

function Test-SamePath {
    param([string] $A, [string] $B)
    if (-not $A -or -not $B) { return $false }
    return [string]::Equals((Get-NormalPath $A), (Get-NormalPath $B), [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-PathUnder {
    param([string] $Path, [string] $Root)
    if (-not $Path -or -not $Root) { return $false }
    $p = (Get-NormalPath $Path) + '\'
    $r = (Get-NormalPath $Root) + '\'
    return $p.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)
}

function ConvertTo-RsaKeyValueXml {
    param([Parameter(Mandatory = $true)] [string] $ModulusBase64, [Parameter(Mandatory = $true)] [string] $ExponentBase64)
    $m = ($ModulusBase64 -replace '\s', '')
    $e = ($ExponentBase64 -replace '\s', '')
    return "<RSAKeyValue><Modulus>$m</Modulus><Exponent>$e</Exponent></RSAKeyValue>"
}

# The deployment manifest's signing key, from ds:RSAKeyValue, as the <RSAKeyValue> string the trust
# store holds. Every occurrence must be the same key.
function Get-ManifestSigningKeyXml {
    param([Parameter(Mandatory = $true)] [string] $ManifestXml)
    $doc = New-Object System.Xml.XmlDocument
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

# A key as two hex numbers with leading zero bytes dropped: what the runtime's ExportCspBlob
# comparison sees, without asking a crypto provider for anything.
function ConvertFrom-RsaKeyXml {
    param([string] $Xml)
    if (-not $Xml) { throw 'empty key' }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $mod = $doc.SelectSingleNode("//*[local-name()='Modulus']")
    $exp = $doc.SelectSingleNode("//*[local-name()='Exponent']")
    if ($null -eq $mod -or $null -eq $exp) { throw 'not an RSAKeyValue: Modulus or Exponent missing' }
    $out = @{}
    foreach ($pair in @(@('Modulus', $mod.InnerText), @('Exponent', $exp.InnerText))) {
        $bytes = [Convert]::FromBase64String(($pair[1] -replace '\s', ''))
        $start = 0
        while ($start -lt ($bytes.Length - 1) -and $bytes[$start] -eq 0) { $start++ }
        $hex = New-Object System.Text.StringBuilder
        for ($i = $start; $i -lt $bytes.Length; $i++) { [void]$hex.Append($bytes[$i].ToString('X2')) }
        $out[$pair[0]] = $hex.ToString()
    }
    return $out
}

function Test-SameRsaKey {
    param([string] $A, [string] $B)
    try {
        $x = ConvertFrom-RsaKeyXml $A
        $y = ConvertFrom-RsaKeyXml $B
        return ($x.Modulus -eq $y.Modulus -and $x.Exponent -eq $y.Exponent)
    }
    catch {
        return $false
    }
}

function Get-TextSha256 {
    param([string] $Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes([string]$Text)))).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Outlook's hard-disable list holds UTF-16 blobs naming the add-in (restated from the guest script).
function Test-DisabledItemsMention {
    param([object[]] $Blobs)
    foreach ($b in @($Blobs)) {
        if ($null -eq $b) { continue }
        $text = [System.Text.Encoding]::Unicode.GetString([byte[]]$b)
        if ($text.ToLowerInvariant().Contains('outlookai')) { return $true }
    }
    return $false
}

function Get-LoadBehaviorMeaning {
    param($Value)
    if ($null -eq $Value) { return 'missing - Outlook will not load it' }
    switch ([int]$Value) {
        3 { return 'load at startup' }
        2 { return 'NOT LOADED - Outlook sets 2 when a load fails' }
        0 { return 'NOT LOADED - switched off' }
        1 { return 'loaded now, not at startup' }
        9 { return 'load on demand' }
        16 { return 'load once at the next start, then on demand' }
    }
    return 'unusual value'
}

function New-DevBuildFolderName {
    param([string] $CommitSha, [DateTime] $When)
    return ('{0:yyyyMMdd-HHmmss}-{1}' -f $When, $CommitSha.Substring(0, 12))
}

# =============================================================================================
# READING THE MACHINE. Read-only.
# =============================================================================================

function Get-InclusionEntries {
    param([string] $InclusionKey)
    $entries = @()
    $root = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($InclusionKey, $false)
    if ($null -eq $root) { return $entries }
    try {
        foreach ($name in $root.GetSubKeyNames()) {
            $k = $root.OpenSubKey($name, $false)
            if ($null -eq $k) { continue }
            try { $entries += [pscustomobject]@{ Name = $name; Url = [string]$k.GetValue('Url'); PublicKey = [string]$k.GetValue('PublicKey') } }
            finally { $k.Close() }
        }
    }
    finally { $root.Close() }
    return $entries
}

function Get-InstalledRelease {
    param($Layout, $Spec)
    $rel = $Layout.UninstallRoot + '\' + $Spec.UninstallKeyName
    $r = [pscustomobject]@{
        Found = $false; KeyPath = "HKCU\$rel"; InstallLocation = $null; AppDir = $null
        DisplayVersion = $null; Vsto = $null; VstoExists = $false; DllVersion = $null
    }
    $vals = Get-KeyValues $rel
    if ($null -eq $vals) { return $r }
    $r.Found = $true
    if ($vals.Contains('InstallLocation')) { $r.InstallLocation = [string]$vals['InstallLocation'].Data }
    if ($vals.Contains('DisplayVersion')) { $r.DisplayVersion = [string]$vals['DisplayVersion'].Data }
    if ($r.InstallLocation) {
        $r.AppDir = $r.InstallLocation.TrimEnd('\')
        $r.Vsto = Join-Path $r.AppDir $ManifestFileName
        $r.VstoExists = Test-Path -LiteralPath $r.Vsto -PathType Leaf
        $dll = Join-Path $r.AppDir 'OutlookAI.dll'
        if (Test-Path -LiteralPath $dll -PathType Leaf) { $r.DllVersion = (Get-Item -LiteralPath $dll).VersionInfo.FileVersion }
    }
    return $r
}

# installed | dev-build | build-folder | none
function Get-BuildClass {
    param([string] $VstoPath, [string] $InstalledVsto, [string] $DevRoot)
    if (-not $VstoPath) { return 'none' }
    if ($InstalledVsto -and (Test-SamePath $VstoPath $InstalledVsto)) { return 'installed' }
    if (Test-PathUnder $VstoPath $DevRoot) { return 'dev-build' }
    return 'build-folder'
}

function Read-DevRecord {
    param([string] $Folder)
    $f = Join-Path $Folder $RecordFileName
    if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $f -Raw | ConvertFrom-Json) } catch { return $null }
}

function Write-DevRecord {
    param([string] $Path, $Record)
    $json = $Record | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

# Whether a manifest will load without the trust prompt: an inclusion entry for its URL AND its key.
function Get-TrustState {
    param([string] $VstoPath, $Entries)
    if (-not $VstoPath -or -not (Test-Path -LiteralPath $VstoPath -PathType Leaf)) { return [pscustomobject]@{ Trusted = $false; Entry = $null; Text = 'no manifest to trust' } }
    $key = $null
    try { $key = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText($VstoPath)) } catch { return [pscustomobject]@{ Trusted = $false; Entry = $null; Text = "the manifest's signing key cannot be read: $($_.Exception.Message)" } }
    $sameUrl = @($Entries | Where-Object { Test-SameInclusionUrl $_.Url $VstoPath })
    foreach ($e in $sameUrl) {
        if (Test-SameRsaKey $e.PublicKey $key) { return [pscustomobject]@{ Trusted = $true; Entry = $e.Name; Text = "loads without a trust prompt - inclusion entry $($e.Name), same URL and same signing key" } }
    }
    if ($sameUrl.Count -gt 0) { return [pscustomobject]@{ Trusted = $false; Entry = $null; Text = 'an inclusion entry names this URL but a DIFFERENT key - Office will show the trust prompt' } }
    return [pscustomobject]@{ Trusted = $false; Entry = $null; Text = "no inclusion entry - Office will show the 'Publisher cannot be verified' prompt at the next start" }
}

function Get-StatusReport {
    param($Layout, $Spec, [string] $DevRoot, [switch] $IncludeProcesses)
    $installed = Get-InstalledRelease $Layout $Spec
    $entries = @(Get-InclusionEntries $Layout.InclusionKey)
    $reg = Get-KeyValues $Layout.AddinKey
    $manifestValue = $null
    $load = $null
    if ($null -ne $reg -and $reg.Contains('Manifest')) { $manifestValue = [string]$reg['Manifest'].Data }
    if ($null -ne $reg -and $reg.Contains('LoadBehavior')) { $load = $reg['LoadBehavior'].Data }
    $vsto = ConvertFrom-ManifestValue $manifestValue
    $class = Get-BuildClass -VstoPath $vsto -InstalledVsto $installed.Vsto -DevRoot $DevRoot
    $vstoExists = $false
    $dllVersion = $null
    $dllWritten = $null
    $record = $null
    if ($vsto) {
        $vstoExists = Test-Path -LiteralPath $vsto -PathType Leaf
        $dll = Join-Path (Split-Path -Parent $vsto) 'OutlookAI.dll'
        if (Test-Path -LiteralPath $dll -PathType Leaf) {
            $item = Get-Item -LiteralPath $dll
            $dllVersion = $item.VersionInfo.FileVersion
            $dllWritten = $item.LastWriteTime
        }
        if ($class -eq 'dev-build') { $record = Read-DevRecord (Split-Path -Parent $vsto) }
    }
    $trust = Get-TrustState -VstoPath $vsto -Entries $entries
    $installedTrust = $null
    if ($installed.VstoExists) { $installedTrust = Get-TrustState -VstoPath $installed.Vsto -Entries $entries }

    $trustRows = @()
    $otherAddins = 0
    foreach ($e in $entries) {
        $p = ConvertFrom-ManifestValue $e.Url
        if (-not $p -or -not ([System.IO.Path]::GetFileName($p)).Equals($ManifestFileName, [System.StringComparison]::OrdinalIgnoreCase)) { $otherAddins++; continue }
        $trustRows += [pscustomobject]@{ Name = $e.Name; Url = $e.Url; Class = (Get-BuildClass -VstoPath $p -InstalledVsto $installed.Vsto -DevRoot $DevRoot) }
    }

    $devBuilds = @()
    if ($DevRoot -and (Test-Path -LiteralPath $DevRoot -PathType Container)) {
        foreach ($d in @(Get-ChildItem -LiteralPath $DevRoot -Directory | Sort-Object Name)) {
            $rec = Read-DevRecord $d.FullName
            $commit = $null
            if ($null -ne $rec) { $commit = [string]$rec.commit }
            $devBuilds += [pscustomobject]@{ Name = $d.Name; Folder = $d.FullName; Commit = $commit; Registered = ($vsto -and (Test-PathUnder $vsto $d.FullName)) }
        }
    }

    $dnd = @(Get-DoNotDisableChanges $Layout $Spec)
    $disabledHits = @()
    foreach ($e in $Spec.DoNotDisable) {
        $vals = Get-KeyValues ($Layout.Prefix + "Software\Microsoft\Office\$($e.Version)\Outlook\Resiliency\DisabledItems")
        if ($null -eq $vals) { continue }
        $blobs = @()
        foreach ($n in $vals.Keys) { if ($vals[$n].Data -is [byte[]]) { $blobs += , $vals[$n].Data } }
        if (Test-DisabledItemsMention $blobs) { $disabledHits += $e.Version }
    }

    $outlook = @()
    if ($IncludeProcesses) { $outlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }

    return [pscustomobject]@{
        RegistrationKey = "HKCU\$($Layout.AddinKey)"; Registration = $reg; ManifestValue = $manifestValue; LoadBehavior = $load
        Vsto = $vsto; VstoExists = $vstoExists; DllVersion = $dllVersion; DllWritten = $dllWritten; Class = $class; Record = $record
        Trust = $trust; Installed = $installed; InstalledTrust = $installedTrust; TrustRows = $trustRows; OtherAddinEntries = $otherAddins
        DevRoot = $DevRoot; DevBuilds = $devBuilds; DoNotDisable = $dnd; DisabledHits = $disabledHits; OutlookPids = $outlook
    }
}

function Get-StatusVerdict {
    param($Report)
    if ($Report.Class -eq 'none') { return 'NOTHING REGISTERED: Outlook will not load OutlookAI. -Restore puts the installed release back.' }
    if (-not $Report.VstoExists) { return "BROKEN: the registered manifest does not exist, so Outlook will fail to load the add-in. -Restore puts the installed release back." }
    if ($Report.DisabledHits.Count -gt 0) { return 'HARD-DISABLED: Outlook''s Resiliency\DisabledItems names the add-in; re-enable it in File > Options > Add-ins > Disabled Items.' }
    if ($null -eq $Report.LoadBehavior -or [int]$Report.LoadBehavior -ne 3) { return "NOT LOADING: LoadBehavior is not 3. -Restore (or -Execute of a dev build) writes 3." }
    $prompt = ''
    if (-not $Report.Trust.Trusted) { $prompt = ' Office will show its trust prompt for it at the next start.' }
    if ($Report.Class -eq 'installed') { return "THE INSTALLED RELEASE $($Report.Installed.DisplayVersion).$prompt" }
    if ($Report.Class -eq 'dev-build') { return "A DEV BUILD this script put there. -Restore puts the installed release back.$prompt" }
    return "A BUILD FOLDER, neither the installed release nor a dev-build copy - builds write into it. -Restore puts the installed release back.$prompt"
}

function Write-StatusReport {
    param($Report)
    Say "== What Outlook will load at its next start =="
    Say "  registration  $($Report.RegistrationKey)"
    if ($null -eq $Report.Registration) { Say '    (the key does not exist)' }
    else {
        foreach ($n in $Report.Registration.Keys) { Say ("    {0,-13} {1}" -f $n, (Format-RegValue $Report.Registration[$n])) }
    }
    Say "    LoadBehavior  means: $(Get-LoadBehaviorMeaning $Report.LoadBehavior)"
    if ($Report.Vsto) {
        Say "  loads         $($Report.Vsto)"
        if ($Report.DllVersion) { Say ("    version     OutlookAI.dll {0}, written {1:yyyy-MM-dd HH:mm}" -f $Report.DllVersion, $Report.DllWritten) }
        elseif ($Report.VstoExists) { Say '    version     (no OutlookAI.dll beside the manifest)' }
        else { Say '    version     THE MANIFEST DOES NOT EXIST' }
        $what = 'A BUILD FOLDER - not the installed release and not a dev-build copy; builds write into it'
        if ($Report.Class -eq 'installed') { $what = 'the installed release' }
        if ($Report.Class -eq 'dev-build') { $what = 'a dev build under ' + $Report.DevRoot }
        Say "    this is     $what"
        if ($null -ne $Report.Record) { Say "    built from  $($Report.Record.commit) - $($Report.Record.subject)" }
        Say "    trust       $($Report.Trust.Text)"
    }
    Say ''
    Say '== The installed release =='
    if (-not $Report.Installed.Found) { Say "  NONE: no uninstall entry at $($Report.Installed.KeyPath)" }
    else {
        Say "  $($Report.Installed.DisplayVersion) at $($Report.Installed.AppDir)  ($($Report.Installed.KeyPath))"
        if ($Report.Installed.VstoExists) {
            Say "    payload     $ManifestFileName present, OutlookAI.dll $($Report.Installed.DllVersion)"
            Say "    trust       $($Report.InstalledTrust.Text)"
        }
        else { Say "    payload     MISSING: $($Report.Installed.Vsto)" }
    }
    Say ''
    Say "== Dev builds in $($Report.DevRoot) =="
    if ($Report.DevBuilds.Count -eq 0) { Say '  none' }
    foreach ($d in $Report.DevBuilds) {
        $mark = ''
        if ($d.Registered) { $mark = '  <- registered' }
        Say ("  {0}  commit {1}{2}" -f $d.Name, $d.Commit, $mark)
    }
    Say ''
    Say "== OutlookAI trust entries (HKCU\$InclusionKeyRel) =="
    if ($Report.TrustRows.Count -eq 0) { Say '  none' }
    foreach ($t in $Report.TrustRows) { Say ("  {0}  {1,-12} {2}" -f $t.Name, $t.Class, $t.Url) }
    if ($Report.OtherAddinEntries -gt 0) { Say "  ($($Report.OtherAddinEntries) entry(s) for other add-ins, not shown and never touched)" }
    Say ''
    Say '== Slow-add-in exemption and hard-disable list =='
    foreach ($d in $Report.DoNotDisable) {
        $state = 'present'
        if (-not $d.Same) { $state = "NOT AS THE INSTALLER WRITES IT: $($d.Before)" }
        Say ("  {0}  {1}" -f $d.Version, $state)
    }
    if ($Report.DisabledHits.Count -gt 0) { Say "  DisabledItems NAMES OutlookAI under $($Report.DisabledHits -join ', ')" }
    else { Say '  DisabledItems does not name OutlookAI' }
    Say ''
    if ($Report.OutlookPids.Count -gt 0) {
        Say "Outlook is running (pid $($Report.OutlookPids -join ', ')). It chose its add-ins when it started; a change takes effect at its next start."
    }
    else { Say 'Outlook is not running. It reads the registration when it next starts.' }
    Say "VERDICT: $(Get-StatusVerdict $Report)"
}

# =============================================================================================
# WRITING THE MACHINE. Only ever called with -Execute, or by -SelfTest against a scratch layout.
# =============================================================================================

# The entry the trust prompt writes. Any other entry for the same URL goes first - what the VSTO
# runtime's own UserInclusionList.Add does.
function Add-InclusionEntry {
    param([string] $InclusionKey, [string] $Url, [string] $PublicKey)
    Assert-WritableKey $InclusionKey
    $removed = @()
    foreach ($e in @(Get-InclusionEntries $InclusionKey)) {
        if (Test-SameInclusionUrl $e.Url $Url) { Remove-KeyTree ($InclusionKey + '\' + $e.Name); $removed += $e.Name }
    }
    $name = [guid]::NewGuid().ToString()
    Set-KeyValue ($InclusionKey + '\' + $name) 'Url' 'String' $Url
    Set-KeyValue ($InclusionKey + '\' + $name) 'PublicKey' 'String' $PublicKey
    return [pscustomobject]@{ Name = $name; Removed = $removed }
}

function Set-RegistrationValues {
    param([string] $AddinKey, $Changes)
    foreach ($c in $Changes.Values) { if (-not $c.Same) { Set-KeyValue $AddinKey $c.Name $c.Kind $c.Data } }
}

function Set-DoNotDisableValues {
    param($DoNotDisable)
    foreach ($d in $DoNotDisable) { if (-not $d.Same) { Set-KeyValue $d.Key $d.Name $d.Kind $d.Data } }
}

function ConvertTo-RecordValues {
    param($Values)
    if ($null -eq $Values) { return $null }
    $o = [ordered]@{}
    foreach ($n in $Values.Keys) { $o[$n] = [ordered]@{ kind = $Values[$n].Kind; data = $Values[$n].Data } }
    return $o
}

function Assert-RegistrationIs {
    param([string] $AddinKey, $Target)
    $verify = Get-RegistrationChanges (Get-KeyValues $AddinKey) $Target
    if ($verify.Pending -gt 0) {
        $bad = @($verify.Values | Where-Object { -not $_.Same } | ForEach-Object { "$($_.Name): $($_.Before), expected $($_.After)" })
        throw "READ-BACK MISMATCH under HKCU\$AddinKey - $($bad -join '; ')"
    }
}

# Copy a flattened publish folder into DevBuilds, record what it replaces, trust it, register it.
function Install-DevBuild {
    param($Layout, $Spec, [string] $PublishDir, [string] $SigningKeyXml, [string] $DevRoot, [string] $FolderName, $Provenance, [string] $ExpectedDllVersion)
    $folder = Join-Path $DevRoot $FolderName
    if (Test-Path -LiteralPath $folder) { throw "REFUSING: $folder already exists; a dev build always gets a new folder." }
    # Everything that can refuse, before anything is written.
    foreach ($req in @($ManifestFileName, 'OutlookAI.dll.manifest', 'OutlookAI.dll')) {
        if (-not (Test-Path -LiteralPath (Join-Path $PublishDir $req) -PathType Leaf)) { throw "REFUSING: the build output has no $req ($PublishDir)." }
    }
    if (Test-Path -LiteralPath (Join-Path $PublishDir 'McpServer')) { throw 'REFUSING: the build output carries a McpServer folder. A dev build never brings the MCP server (TODO.md).' }
    $key = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText((Join-Path $PublishDir $ManifestFileName)))
    if (-not (Test-SameRsaKey $key $SigningKeyXml)) { throw 'REFUSING: the build output is not signed by the key the build signed with.' }
    if ($ExpectedDllVersion) {
        $v = (Get-Item -LiteralPath (Join-Path $PublishDir 'OutlookAI.dll')).VersionInfo.FileVersion
        if ($v -ne $ExpectedDllVersion) { throw "REFUSING: the built OutlookAI.dll is $v, expected $ExpectedDllVersion." }
    }

    $vsto = Join-Path $folder $ManifestFileName
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    try {
        Copy-Item -Path (Join-Path $PublishDir '*') -Destination $folder -Recurse -Force
        $copied = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText($vsto))
        if (-not (Test-SameRsaKey $copied $key)) { throw 'REFUSING: the copied manifest is not the one that was built.' }
    }
    catch {
        # Nothing is registered yet, so a half-copied folder is simply removed.
        Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }

    $installedRelease = Get-InstalledRelease $Layout $Spec
    $current = Get-KeyValues $Layout.AddinKey
    $currentVsto = $null
    if ($null -ne $current -and $current.Contains('Manifest')) { $currentVsto = ConvertFrom-ManifestValue ([string]$current['Manifest'].Data) }
    $target = Get-TargetRegistration $Spec $folder
    $changes = Get-RegistrationChanges $current $target
    $dnd = @(Get-DoNotDisableChanges $Layout $Spec)
    $url = ConvertTo-InclusionUrl $vsto
    $recordPath = Join-Path $folder $RecordFileName
    $record = [ordered]@{
        schema      = 1
        producedBy  = 'Tools/Switch-AddInBuild.ps1'
        status      = 'copied'
        commit      = $Provenance.Commit
        ref         = $Provenance.Ref
        subject     = $Provenance.Subject
        version     = $DevVersion
        builtUtc    = $Provenance.BuiltUtc
        folder      = $folder
        replaced    = [ordered]@{
            key       = "HKCU\$($Layout.AddinKey)"
            existed   = $changes.KeyExisted
            values    = (ConvertTo-RecordValues $current)
            was       = (Get-BuildClass -VstoPath $currentVsto -InstalledVsto $installedRelease.Vsto -DevRoot $DevRoot)
        }
        registered  = (ConvertTo-RecordValues $target)
        doNotDisable = @($dnd | ForEach-Object { [ordered]@{ key = "HKCU\$($_.Key)"; name = $_.Name; before = $_.Before; written = (-not $_.Same) } })
        trust       = [ordered]@{ url = $url; publicKeyXml = $key; publicKeySha256 = (Get-TextSha256 $key); entry = $null }
        hostUntouchedByBuild = $Provenance.HostProof
    }
    Write-DevRecord $recordPath $record

    # Trust first, registration second: if this stops between the two, a trust entry with nothing
    # registered is harmless; a registration with no trust would put up a prompt.
    $trust = Add-InclusionEntry $Layout.InclusionKey $url $key
    $record.trust.entry = $trust.Name
    Write-DevRecord $recordPath $record
    Set-RegistrationValues $Layout.AddinKey $changes
    Set-DoNotDisableValues $dnd
    Assert-RegistrationIs $Layout.AddinKey $target
    $record.status = 'registered'
    $record['registeredUtc'] = [DateTime]::UtcNow.ToString('o')
    Write-DevRecord $recordPath $record
    return [pscustomobject]@{ Folder = $folder; Vsto = $vsto; Record = $recordPath; TrustEntry = $trust.Name; TrustRemoved = $trust.Removed; Changes = $changes; DoNotDisable = $dnd }
}

# The -Restore plan, and with -Apply the writes. Works from the uninstall entry alone.
function Invoke-Restore {
    param($Layout, $Spec, [string] $DevRoot, [switch] $RemoveBuildTrustEntries, [switch] $Apply)
    $installed = Get-InstalledRelease $Layout $Spec
    if (-not $installed.Found) { throw "REFUSING: there is no installed release to go back to - no uninstall entry at $($installed.KeyPath). Install a release first." }
    if (-not $installed.AppDir) { throw "REFUSING: the uninstall entry $($installed.KeyPath) has no InstallLocation." }
    if (-not $installed.VstoExists) { throw "REFUSING: the installed release's manifest $($installed.Vsto) does not exist; registering it would leave Outlook nothing to load." }

    $current = Get-KeyValues $Layout.AddinKey
    $target = Get-TargetRegistration $Spec $installed.AppDir
    $changes = Get-RegistrationChanges $current $target
    $dnd = @(Get-DoNotDisableChanges $Layout $Spec)
    $currentVsto = $null
    if ($null -ne $current -and $current.Contains('Manifest')) { $currentVsto = ConvertFrom-ManifestValue ([string]$current['Manifest'].Data) }
    $currentClass = Get-BuildClass -VstoPath $currentVsto -InstalledVsto $installed.Vsto -DevRoot $DevRoot

    $remove = @()
    $keep = @()
    foreach ($e in @(Get-InclusionEntries $Layout.InclusionKey)) {
        $p = ConvertFrom-ManifestValue $e.Url
        if (-not $p -or -not ([System.IO.Path]::GetFileName($p)).Equals($ManifestFileName, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $class = Get-BuildClass -VstoPath $p -InstalledVsto $installed.Vsto -DevRoot $DevRoot
        if ($class -eq 'installed') { $keep += [pscustomobject]@{ Name = $e.Name; Url = $e.Url; Why = 'the installed release' }; continue }
        if ($class -eq 'dev-build') { $remove += [pscustomobject]@{ Name = $e.Name; Url = $e.Url; Why = 'a dev build' }; continue }
        if ($RemoveBuildTrustEntries) { $remove += [pscustomobject]@{ Name = $e.Name; Url = $e.Url; Why = 'a build folder (-RemoveBuildTrust)' } }
        else { $keep += [pscustomobject]@{ Name = $e.Name; Url = $e.Url; Why = 'a build folder - kept; -RemoveBuildTrust removes it' } }
    }
    $installedTrust = Get-TrustState -VstoPath $installed.Vsto -Entries @(Get-InclusionEntries $Layout.InclusionKey)

    $plan = [pscustomobject]@{
        Installed = $installed; Changes = $changes; DoNotDisable = $dnd; Remove = $remove; Keep = $keep
        CurrentVsto = $currentVsto; CurrentClass = $currentClass; InstalledTrust = $installedTrust; Target = $target; Applied = $false; Audit = $null
    }
    if (-not $Apply) { return $plan }

    # Registration first, trust second: if this stops between the two, a dev build that is no
    # longer registered but still trusted is harmless.
    Set-RegistrationValues $Layout.AddinKey $changes
    foreach ($r in $remove) { Remove-KeyTree ($Layout.InclusionKey + '\' + $r.Name) }
    Set-DoNotDisableValues $dnd
    Assert-RegistrationIs $Layout.AddinKey $target
    $plan.Applied = $true

    # An audit line of what was replaced: beside the dev build when it was one, and always as a
    # restore-<time>.json in DevBuilds. Best effort - the registry is already right.
    try {
        New-Item -ItemType Directory -Force -Path $DevRoot | Out-Null
        $audit = Join-Path $DevRoot ('restore-{0:yyyyMMdd-HHmmss}.json' -f (Get-Date))
        Write-DevRecord $audit ([ordered]@{
                schema = 1; producedBy = 'Tools/Switch-AddInBuild.ps1 -Restore'; restoredUtc = [DateTime]::UtcNow.ToString('o')
                replaced = (ConvertTo-RecordValues $current); replacedWas = $currentClass
                restoredTo = (ConvertTo-RecordValues $target); installedVersion = $installed.DisplayVersion
                trustRemoved = @($remove | ForEach-Object { [ordered]@{ entry = $_.Name; url = $_.Url; why = $_.Why } })
            })
        $plan.Audit = $audit
        if ($currentClass -eq 'dev-build') {
            $rec = Join-Path (Split-Path -Parent $currentVsto) $RecordFileName
            if (Test-Path -LiteralPath $rec) {
                $o = Get-Content -LiteralPath $rec -Raw | ConvertFrom-Json
                $o | Add-Member -NotePropertyName restoredUtc -NotePropertyValue ([DateTime]::UtcNow.ToString('o')) -Force
                $o | Add-Member -NotePropertyName status -NotePropertyValue 'restored' -Force
                Write-DevRecord $rec $o
            }
        }
    }
    catch { Say "  NOTE the registry is restored, but the audit record could not be written: $($_.Exception.Message)" }
    return $plan
}

function Write-RestorePlan {
    param($Plan)
    $i = $Plan.Installed
    Say "  installed release  $($i.DisplayVersion) at $($i.AppDir)  ($($i.KeyPath))"
    Say "  registered now     $($Plan.CurrentVsto)  [$($Plan.CurrentClass)]"
    Say "  HKCU\$AddinKeyRel"
    foreach ($c in $Plan.Changes.Values) {
        if ($c.Same) { Say ("    {0,-13} {1}  (unchanged)" -f $c.Name, $c.After) }
        else { Say ("    {0,-13} {1}`n    {2,-13} -> {3}" -f $c.Name, $c.Before, '', $c.After) }
    }
    foreach ($x in $Plan.Changes.Extra) { Say ("    {0,-13} {1}  (not the installer's; left as it is)" -f $x.Name, $x.Value) }
    foreach ($d in $Plan.DoNotDisable) {
        if ($d.Same) { Say "  DoNotDisableAddinList $($d.Version)  unchanged" }
        else { Say "  DoNotDisableAddinList $($d.Version)  $($d.Before) -> $($d.Data) ($($d.Kind))" }
    }
    Say "  trust entries (HKCU\$InclusionKeyRel), OutlookAI's only:"
    if ($Plan.Remove.Count -eq 0 -and $Plan.Keep.Count -eq 0) { Say '    none' }
    foreach ($r in $Plan.Remove) { Say "    remove  $($r.Name)  $($r.Url)  - $($r.Why)" }
    foreach ($k in $Plan.Keep) { Say "    keep    $($k.Name)  $($k.Url)  - $($k.Why)" }
    Say "  the installed release, once registered: $($Plan.InstalledTrust.Text)"
}

# =============================================================================================
# THE BUILD - Testbed/host/Publish-AddInPayload.ps1's mechanism, restated. Guard 1: RegisterOfficeAddin
# off PrepareForRun by a GLOBAL property. Guard 2: the three registry-writing tasks replaced by
# Override="true" stand-ins that only log. Guard 3: the host snapshotted before and after, any
# trace a failure, a changed registration put back. None of them depends on the commit's own
# OutlookAI.csproj.
# =============================================================================================

function Get-PrepareForRunOverride { return ($PrepareForRunChain -join ';') }

function Get-NoHostWriteTargets {
    $blocks = @()
    foreach ($name in $StandInTasks.Keys) {
        $spec = $StandInTasks[$name]
        $params = @()
        foreach ($p in $spec.Parameters) { $params += "      <$p ParameterType=`"System.String`" />" }
        if ($spec.Output) { $params += "      <$($spec.Output) ParameterType=`"Microsoft.Build.Framework.ITaskItem[]`" Output=`"true`" />" }
        $blocks += @"
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
    WRITTEN BY Tools/Switch-AddInBuild.ps1 FOR ONE BUILD. NOT PART OF THE REPOSITORY.
    Stand-ins for the three VSTO build tasks that write the build machine's registry; they only
    say they did not run. Imported in place of the machine's own CustomBeforeMicrosoftCommonTargets
    hook, so it imports that hook itself - a machine that has one keeps it.
  -->
  <Import Project="`$(MSBuildExtensionsPath)\v`$(MSBuildToolsVersion)\Custom.Before.Microsoft.Common.targets" Condition="Exists('`$(MSBuildExtensionsPath)\v`$(MSBuildToolsVersion)\Custom.Before.Microsoft.Common.targets')" />
$($blocks -join "`r`n")
</Project>
"@
}

# A scalar takes %3B for a semicolon; a target LIST needs a real semicolon inside double quotes,
# because a list is split before it is unescaped (MSB4057 otherwise - measured by the testbed script).
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
        [Parameter(Mandatory = $true)] [string] $FileLog
    )
    $projectArg = $ProjectPath
    if ($projectArg -match '\s') { $projectArg = '"' + $projectArg + '"' }
    $list = @(
        $projectArg,
        '/t:Publish',
        # .github/workflows/release.yml's "Build and Publish", property for property.
        (Format-MSBuildProperty 'Configuration' 'Release'),
        (Format-MSBuildProperty 'ApplicationVersion' $DevVersion),
        (Format-MSBuildProperty 'PublishDir' 'publish\'),
        (Format-MSBuildProperty 'BootstrapperEnabled' 'true'),
        (Format-MSBuildProperty 'IsWebBootstrapper' 'false'),
        (Format-MSBuildProperty 'DefineConstants' 'VSTO40;TRACE'),
        # The throwaway key instead of the release secret.
        (Format-MSBuildProperty 'ManifestCertificateThumbprint' $Thumbprint),
        # Guards 1 and 2, and never the IDE's registering behaviour, whatever the environment says.
        (Format-MSBuildProperty 'PrepareForRunDependsOn' (Get-PrepareForRunOverride) -List),
        (Format-MSBuildProperty 'CustomBeforeMicrosoftCommonTargets' $StandInTargets),
        (Format-MSBuildProperty 'BuildingInsideVisualStudio' 'false'),
        # Nothing may outlive the build and hold a handle.
        (Format-MSBuildProperty 'UseSharedCompilation' 'false'),
        '/nodeReuse:false',
        '/m:1',
        '/nologo',
        '/v:minimal'
    )
    $flp = "LogFile=$FileLog;Verbosity=detailed;Encoding=UTF-8"
    if ($flp -match '\s') { $flp = '"' + $flp + '"' }
    $list += "/flp:$flp"
    return $list
}

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

function Test-StandInSentinels {
    param([string] $LogText)
    $problems = @()
    $notes = @()
    if (-not $LogText) { $LogText = '' }
    $lower = $LogText.ToLowerInvariant()
    if (-not $lower.Contains('done building target "visualstudioforapplicationsbuild"')) {
        $problems += 'VisualStudioForApplicationsBuild never finished, so no signed manifests were made - or the log is not detailed enough to say.'
    }
    if ($lower.Contains('target "registerofficeaddin"')) { $problems += 'the RegisterOfficeAddin TARGET appears in the build log: guard 1 did not hold.' }
    foreach ($name in $StandInTasks.Keys) {
        $seen = $LogText.Contains("$StandInSentinel`: $name did not run")
        if ($StandInTasks[$name].Scheduled -and -not $seen) { $problems += "the $name stand-in never logged: either the real task ran, or the build stopped before it; guard 3 decides which." }
        if (-not $StandInTasks[$name].Scheduled -and $seen) { $notes += "the $name stand-in RAN, so RegisterOfficeAddin was still scheduled: guard 1 did not hold and guard 2 caught it." }
    }
    return [pscustomobject]@{ Ok = ($problems.Count -eq 0); Problems = $problems; Notes = $notes }
}

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

function Format-RegistryData($Value, [Microsoft.Win32.RegistryValueKind] $Kind) {
    if ($null -eq $Value) { return '<null>' }
    if ($Kind -eq [Microsoft.Win32.RegistryValueKind]::Binary) { return (($Value | ForEach-Object { '{0:X2}' -f $_ }) -join '') }
    if ($Kind -eq [Microsoft.Win32.RegistryValueKind]::MultiString) { return ($Value -join '|') }
    return [string]$Value
}

# Read-only. Every value under every given key, one line each.
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
            finally { $key.Close() }
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
        catch { $lines += ('CERT\' + $s + '|unreadable:' + $_.Exception.Message) }
        finally { $store.Close() }
    }
    return @($lines | Sort-Object)
}

function Get-HostSnapshot { return @((Get-RegistryLines -Keys $WatchedRegistry) + (Get-CertLines -Stores $WatchedCertStores)) }

function Compare-Lines {
    param([string[]] $Before, [string[]] $After)
    $b = @{}
    foreach ($l in @($Before)) { if ($null -ne $l) { $b[$l] = $true } }
    $a = @{}
    foreach ($l in @($After)) { if ($null -ne $l) { $a[$l] = $true } }
    $added = @(); foreach ($l in $a.Keys) { if (-not $b.ContainsKey($l)) { $added += $l } }
    $removed = @(); foreach ($l in $b.Keys) { if (-not $a.ContainsKey($l)) { $removed += $l } }
    return [pscustomobject]@{ Added = @($added | Sort-Object); Removed = @($removed | Sort-Object); Same = ($added.Count -eq 0 -and $removed.Count -eq 0) }
}

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
        $m = ($ModulusBase64 -replace '\s', '')
        if ($m.Length -gt 48) { $m = $m.Substring(0, 48) }
        $needles += $m.ToLowerInvariant()
    }
    $hits = @()
    foreach ($line in @($Lines)) {
        if ($null -eq $line) { continue }
        $low = $line.ToLowerInvariant()
        foreach ($n in $needles) { if ($n -and $low.Contains($n)) { $hits += $line; break } }
    }
    return @($hits)
}

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

# Only ever called when guard 3 found the registration changed: puts back exactly the key's own
# values the snapshot recorded (a key that did not exist is deleted), and says so.
function Restore-KeyFromLines {
    param([string[]] $BeforeLines, [string] $KeyPath)
    $own = 'HKCU\' + $KeyPath + '|'
    $mine = @($BeforeLines | Where-Object { $_ -and $_.StartsWith($own, [System.StringComparison]::OrdinalIgnoreCase) -and -not $_.EndsWith('|<key>') })
    $existed = @(Select-KeyLines -Lines $BeforeLines -KeyPath $KeyPath).Count -gt 0
    if (-not $existed) {
        Say "  RESTORE: HKCU\$KeyPath did not exist before the build - deleting it"
        Remove-KeyTree $KeyPath
        return
    }
    $wanted = @{}
    foreach ($line in $mine) {
        $afterPipe = $line.Substring($line.IndexOf('|') + 1)
        $name = $afterPipe.Substring(0, $afterPipe.IndexOf('='))
        $rest = $afterPipe.Substring($afterPipe.IndexOf('=') + 1)
        $kind = $rest.Substring(0, $rest.IndexOf(':'))
        $data = $rest.Substring($rest.IndexOf(':') + 1)
        $wanted[$name] = $true
        if ($kind -eq 'DWord') { Set-KeyValue $KeyPath $name 'DWord' ([int]$data) } else { Set-KeyValue $KeyPath $name 'String' $data }
        Say "  RESTORE: $name = $data"
    }
    $now = Get-KeyValues $KeyPath
    if ($null -ne $now) {
        foreach ($name in @($now.Keys)) {
            if (-not $wanted.ContainsKey($name)) { Remove-KeyValue $KeyPath $name; Say "  RESTORE: removed $name, which the build added" }
        }
    }
}

# Only ever called when guard 3 found an inclusion entry naming this build's own directory.
function Remove-BuildTrustFootprints {
    param([string] $InclusionKey, [string] $BuildDirectory)
    $removed = @()
    foreach ($e in @(Get-InclusionEntries $InclusionKey)) {
        $p = ConvertFrom-ManifestValue $e.Url
        if ($p -and (Test-PathUnder $p $BuildDirectory)) {
            Remove-KeyTree ($InclusionKey + '\' + $e.Name)
            $removed += $e.Name
            Say "  RESTORE: removed trust entry $($e.Name), which named this build ($($e.Url))"
        }
    }
    return $removed
}

function Remove-ThrowawayCertificate {
    param([Parameter(Mandatory = $true)] [string] $Thumbprint)
    $my = "Cert:\CurrentUser\My\$Thumbprint"
    if (Test-Path -LiteralPath $my) { Remove-Item -LiteralPath $my -DeleteKey -ErrorAction SilentlyContinue }
    # NOT ONLY My: New-SelfSignedCertificate also puts a copy into CurrentUser\CA (measured by
    # Testbed/host/Publish-AddInPayload.ps1). Only a certificate with BOTH this build's thumbprint
    # and this script's subject is removed, so nothing of the maintainer's can match.
    foreach ($s in $WatchedCertStores) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($s, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            foreach ($c in @($store.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint -and $_.Subject -eq $ThrowawaySubject })) {
                $store.Remove($c)
                Say "  removed the copy New-SelfSignedCertificate left in Cert:\CurrentUser\$s"
            }
        }
        catch { Say "  !! could not open Cert:\CurrentUser\$s to clean it: $($_.Exception.Message)" }
        finally { $store.Close() }
    }
    $left = @($WatchedCertStores | Where-Object { Test-Path -LiteralPath "Cert:\CurrentUser\$_\$Thumbprint" })
    if ($left.Count -gt 0) { Say "  !! the throwaway certificate is STILL in: $($left -join ', ') - guard 3 will refuse. Remove it by thumbprint $Thumbprint." }
    else { Say '  throwaway certificate deleted from every CurrentUser store, private key included' }
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
    if (-not (Test-Path -LiteralPath $vswhere)) { throw 'REFUSING TO BUILD: vswhere.exe is not on this machine, so there is no Visual Studio to build with. Or pass -MSBuildPath.' }
    $roots = @(Invoke-NativeCommand { & $vswhere -latest -products * -requires Microsoft.VisualStudio.Workload.Office -property installationPath 2>$null })
    if ($roots.Count -eq 0 -or -not $roots[0]) { throw 'REFUSING TO BUILD: no Visual Studio here has the Office/SharePoint development workload, so there are no VSTO build targets. Or pass -MSBuildPath.' }
    $candidate = Join-Path $roots[0] 'MSBuild\Current\Bin\amd64\MSBuild.exe'
    if (-not (Test-Path -LiteralPath $candidate)) { $candidate = Join-Path $roots[0] 'MSBuild\Current\Bin\MSBuild.exe' }
    if (-not (Test-Path -LiteralPath $candidate)) { throw "Visual Studio at $($roots[0]) has no MSBuild.exe under MSBuild\Current\Bin." }
    return $candidate
}

# Output straight to files, never through this process's pipe, under a deadline: a build that
# spawns children can hold a pipe open after it has finished, and the timeout would never fire.
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
        FilePath = $FilePath; ArgumentList = ($ArgumentList -join ' ')
        RedirectStandardOutput = $out; RedirectStandardError = $err; NoNewWindow = $true; PassThru = $true
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

function Test-IsUnderGitWorkTreeButNotWork {
    param([string] $Path, [string] $WorkTreeRoot)
    if (-not $WorkTreeRoot) { return $false }
    $p = $Path.TrimEnd('\') + '\'
    $root = $WorkTreeRoot.TrimEnd('\') + '\'
    if (-not $p.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    return (-not $p.StartsWith($root + '.work\', [System.StringComparison]::OrdinalIgnoreCase))
}

function Resolve-Commit {
    param([string] $Ref)
    $sha = (@(Invoke-NativeCommand { & git -C $RepoRoot rev-parse --verify "$Ref^{commit}" 2>&1 }) | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') { throw "git could not resolve '$Ref' to a commit in $RepoRoot - $sha" }
    $subject = (@(Invoke-NativeCommand { & git -C $RepoRoot log -1 --format=%s $sha 2>$null }) | Out-String).Trim()
    return [pscustomobject]@{ Sha = $sha; Subject = $subject; Ref = $Ref }
}

# Build one commit with all three guards into $BuildDir. Returns the flattened publish folder
# and the signing key; throws on anything else.
function Invoke-GuardedBuild {
    param($CommitInfo, [string] $BuildDir)
    $msbuild = Resolve-MSBuild
    $vstoTargets = @(Find-InstalledVstoTargets)
    if ($vstoTargets.Count -eq 0) { throw 'REFUSING TO BUILD: no VSTO build targets (OfficeTools\Microsoft.VisualStudio.Tools.Office.targets) in any Visual Studio with the Office workload.' }
    $msbuildVersion = (@(Invoke-NativeCommand { & $msbuild -version -nologo 2>$null }) | Select-Object -Last 1)
    Say "  MSBuild  $msbuild ($msbuildVersion)"
    Say "  VSTO     $($vstoTargets[0])"
    if (Test-Path -LiteralPath $BuildDir) { throw "REFUSING: $BuildDir already exists." }
    New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
    $sourceZip = Join-Path $BuildDir 'source.zip'
    $sourceDir = Join-Path $BuildDir 'source'
    $standIn = Join-Path $BuildDir 'NoHostWrite.targets'
    $log = Join-Path $BuildDir 'msbuild.log'

    Invoke-NativeCommand { & git -C $RepoRoot archive --format=zip -o $sourceZip $CommitInfo.Sha }
    if ($LASTEXITCODE -ne 0) { throw "git archive failed (exit $LASTEXITCODE)." }
    # The framework's own unzip rather than Expand-Archive, whose module is exactly the kind a
    # mixed PSModulePath (see the Import-Module at the top) can hand to the wrong edition.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZip, $sourceDir)
    $project = Join-Path $sourceDir 'OutlookAI.csproj'
    if (-not (Test-Path -LiteralPath $project)) { throw "The archive has no OutlookAI.csproj at its root - $sourceDir" }
    $sources = @(Get-ChildItem -LiteralPath $sourceDir -Filter '*.cs' -File | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) +
        @(foreach ($d in @('Services', 'TaskPane')) { Get-ChildItem -LiteralPath (Join-Path $sourceDir $d) -Filter '*.cs' -File -ErrorAction SilentlyContinue | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw } })
    $regions = Find-FormRegionDeclarations -SourceTexts $sources
    if ($regions.Count -gt 0) { throw "REFUSING TO BUILD: this commit declares an Outlook form region ($($regions -join ', ')); the RegisterFormRegions stand-in would drop it from the manifest." }
    Set-Content -LiteralPath $standIn -Value (Get-NoHostWriteTargets) -Encoding UTF8
    Say "  source   $($CommitInfo.Sha) - $($CommitInfo.Subject)"

    $before = Get-HostSnapshot
    Say "  guard 3  $(@($before).Count) host line(s) snapshotted before the build"
    $thumb = $null
    $signingKeyXml = $null
    $problems = @()
    $publish = Join-Path $sourceDir 'publish'
    try {
        $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $ThrowawaySubject `
            -FriendlyName "OutlookAI dev build $($CommitInfo.Sha) - throwaway, deleted after signing" `
            -CertStoreLocation 'Cert:\CurrentUser\My' -NotAfter (Get-Date).AddYears(10)
        $thumb = [string]$cert.Thumbprint
        Say "  signing  throwaway Cert:\CurrentUser\My\$thumb"
        $msbuildArgs = Get-MSBuildArgumentList -ProjectPath $project -Thumbprint $thumb -StandInTargets $standIn -FileLog $log

        $evalArgs = @($msbuildArgs | Where-Object { $_ -ne '/t:Publish' -and -not $_.StartsWith('/flp:') -and $_ -ne '/v:minimal' }) + '-getProperty:PrepareForRunDependsOn'
        $eval = Invoke-Logged -FilePath $msbuild -ArgumentList $evalArgs -LogStem (Join-Path $BuildDir 'msbuild-evaluate') -TimeoutMinutes 5 -WorkingDirectory $sourceDir
        $evaluated = ''
        if (Test-Path -LiteralPath $eval.Out) { $evaluated = Get-Content -LiteralPath $eval.Out -Raw }
        $chain = @()
        if ($eval.ExitCode -ne 0) { $chain += "MSBuild could not evaluate the project (exit $($eval.ExitCode)); see $($eval.Out)" }
        else { $chain += @(Test-EvaluatedPrepareForRun $evaluated) }
        if ($chain.Count -gt 0) { Show-LogTail $eval.Out 20; throw ('REFUSING TO BUILD: ' + ($chain -join ' ')) }
        Say "  guard 1  PrepareForRunDependsOn = $(($evaluated -replace '\s', ''))"

        $build = Invoke-Logged -FilePath $msbuild -ArgumentList $msbuildArgs -LogStem (Join-Path $BuildDir 'msbuild-console') -TimeoutMinutes $BuildTimeoutMinutes -WorkingDirectory $sourceDir
        if ($build.TimedOut -or $build.ExitCode -ne 0) {
            Show-LogTail $build.Out 40
            $problems += "MSBuild exited $($build.ExitCode) (timed out: $($build.TimedOut)). Full log: $log"
        }
        else {
            $logText = ''
            if (Test-Path -LiteralPath $log) { $logText = Get-Content -LiteralPath $log -Raw }
            $sentinels = Test-StandInSentinels -LogText $logText
            foreach ($n in $sentinels.Notes) { Say "  NOTE $n" }
            if (-not $sentinels.Ok) { $problems += $sentinels.Problems }
            else { Say '  guard 2  both scheduled stand-ins logged; the registration target was not scheduled' }
            # release.yml's "Flatten VSTO payload next to the manifest (|vstolocal)".
            $appFiles = Join-Path $publish 'Application Files'
            $verDirs = @(Get-ChildItem -LiteralPath $appFiles -Directory -ErrorAction SilentlyContinue)
            if ($verDirs.Count -ne 1) { $problems += "Expected exactly one versioned folder under $appFiles, found $($verDirs.Count)." }
            else {
                Get-ChildItem -LiteralPath $verDirs[0].FullName -File | Where-Object { $_.Name -ne $ManifestFileName } | ForEach-Object {
                    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $publish ($_.Name -replace '\.deploy$', '')) -Force
                }
                foreach ($required in @($ManifestFileName, 'OutlookAI.dll.manifest', 'OutlookAI.dll')) {
                    if (-not (Test-Path -LiteralPath (Join-Path $publish $required))) { $problems += "Missing after flatten: publish\$required" }
                }
                if ($problems.Count -eq 0) { $signingKeyXml = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText((Join-Path $publish $ManifestFileName))) }
            }
        }
    }
    catch {
        # Recorded, not rethrown: guard 3 below must look at the host whatever stopped the build.
        $problems += "the build stopped: $($_.Exception.Message)"
    }
    finally {
        if ($thumb) { Remove-ThrowawayCertificate -Thumbprint $thumb }
    }

    $after = Get-HostSnapshot
    $modulus = $null
    if ($signingKeyXml) { $modulus = [regex]::Match($signingKeyXml, '<Modulus>([^<]+)</Modulus>').Groups[1].Value }
    $footprints = @(Find-BuildFootprint -Lines $after -BuildDirectory $BuildDir -Thumbprint $thumb -ModulusBase64 $modulus)
    $regDiff = Compare-Lines -Before @(Select-KeyLines -Lines $before -KeyPath $AddinKeyRel) -After @(Select-KeyLines -Lines $after -KeyPath $AddinKeyRel)
    $allDiff = Compare-Lines -Before $before -After $after
    $hostFailures = @()
    if (-not $regDiff.Same) {
        $hostFailures += "the OutlookAI add-in registration CHANGED during the build (HKCU\$AddinKeyRel):"
        $hostFailures += @($regDiff.Removed | ForEach-Object { '    was: ' + $_ })
        $hostFailures += @($regDiff.Added | ForEach-Object { '    now: ' + $_ })
        Restore-KeyFromLines -BeforeLines $before -KeyPath $AddinKeyRel
    }
    if ($footprints.Count -gt 0) {
        $hostFailures += 'traces of this build are on the host:'
        $hostFailures += @($footprints | ForEach-Object { '    ' + $_ })
        $null = Remove-BuildTrustFootprints -InclusionKey $InclusionKeyRel -BuildDirectory $BuildDir
    }
    if ($hostFailures.Count -gt 0) {
        foreach ($l in $hostFailures) { Say "  FAIL $l" }
        throw 'REFUSING: the build touched the host (above). What this script could put back, it has; nothing was registered.'
    }
    if (-not $allDiff.Same) {
        Say '  WARN something else under the watched keys changed while the build ran. Nothing of this build is in it,'
        Say '       so the likeliest writer is Outlook itself; read it anyway:'
        foreach ($l in $allDiff.Removed) { Say "         - $l" }
        foreach ($l in $allDiff.Added) { Say "         + $l" }
    }
    else { Say "  guard 3  UNCHANGED: $(@($after).Count) host line(s) identical before and after" }
    if ($problems.Count -gt 0) {
        foreach ($p in $problems) { Say "  FAIL $p" }
        throw 'The build did not produce a usable add-in (above). The host is untouched; nothing was registered.'
    }
    return [pscustomobject]@{
        PublishDir = $publish; SigningKeyXml = $signingKeyXml; MSBuild = $msbuild; MSBuildVersion = [string]$msbuildVersion
        HostProof = [ordered]@{ linesCompared = @($after).Count; footprints = 0; registrationUnchanged = $true; identical = $allDiff.Same }
    }
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()

    # Compared with -ceq; searched with String.Contains. Never -like: a bracket in a -like pattern
    # is a character class (Testbed/README.md section 4b; CLAUDE.md mailbox rule 2).
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
    function Test-Throws {
        param([string] $What, [scriptblock] $Block, [string] $Needle)
        $script:Checks++
        $msg = $null
        try { & $Block | Out-Null } catch { $msg = $_.Exception.Message }
        if ($null -ne $msg -and (-not $Needle -or $msg.Contains($Needle))) { Write-Host "  OK   $What" }
        else {
            $script:Failures += "$What : expected a refusal containing [$Needle], got [$msg]"
            Write-Host "  FAIL $What - expected a refusal containing [$Needle], got [$msg]"
        }
    }

    Write-Host "Tools/Switch-AddInBuild.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    Write-Host ''
    Write-Host '== Installer.iss, the real one =='
    $issPath = Join-Path $RepoRoot 'Installer.iss'
    $spec = Get-InstallerSpec ([System.IO.File]::ReadAllText($issPath))
    Test-Case 'the uninstall entry is the AppId plus _is1' '{78AF2871-0CEB-4451-B80D-455552E37C91}_is1' $spec.UninstallKeyName
    Test-Case 'the add-in key carries four values' 'Manifest,LoadBehavior,FriendlyName,Description' (($spec.AddinValues | ForEach-Object { $_.Name }) -join ',')
    Test-Case 'Manifest is the installer''s |vstolocal template' 'file:///{app}\OutlookAI.vsto|vstolocal' (@($spec.AddinValues | Where-Object { $_.Name -eq 'Manifest' })[0].Template)
    Test-Case 'LoadBehavior is a DWORD 3' 'DWord:3' ((@($spec.AddinValues | Where-Object { $_.Name -eq 'LoadBehavior' }) | ForEach-Object { "$($_.Kind):$($_.Template)" }) -join '')
    Test-Case 'the slow-add-in exemption covers 16.0, 15.0 and 17.0' '15.0,16.0,17.0' (($spec.DoNotDisable | ForEach-Object { $_.Version } | Sort-Object) -join ',')
    Test-Case 'InstallDir is skipped on purpose - it must keep naming the installed release' 'Software\OutlookAI|InstallDir' (($spec.Skipped | ForEach-Object { "$($_.Subkey)|$($_.Name)" }) -join ',')
    $target = Get-TargetRegistration $spec 'C:\Users\u\AppData\Local\OutlookAI\Setup'
    Test-Case 'expanded for the installed release, exactly as Inno writes it' 'file:///C:\Users\u\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal' $target['Manifest'].Data
    Test-Case 'a trailing backslash on InstallLocation does not double up' 'file:///C:\x\OutlookAI.vsto|vstolocal' (Get-TargetRegistration $spec 'C:\x\')['Manifest'].Data
    Test-Case 'LoadBehavior expands to the integer 3' 'DWord 3' ("{0} {1}" -f $target['LoadBehavior'].Kind, $target['LoadBehavior'].Data)

    Write-Host ''
    Write-Host '== Installer.iss, shapes that must be refused =='
    $good = "AppId={{11111111-2222-3333-4444-555555555555}`r`n[Registry]`r`nRoot: HKCU; Subkey: ""Software\Microsoft\Office\Outlook\Addins\OutlookAI""; ValueType: string; ValueName: ""Manifest""; ValueData: ""file:///{app}\OutlookAI.vsto|vstolocal""; Flags: uninsdeletekey`r`nRoot: HKCU; Subkey: ""Software\Microsoft\Office\Outlook\Addins\OutlookAI""; ValueType: dword; ValueName: ""LoadBehavior""; ValueData: ""3""`r`nRoot: HKCU; Subkey: ""Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList""; ValueType: dword; ValueName: ""OutlookAI""; ValueData: ""1""`r`n[Code]`r`nRoot: nonsense that is not registry"
    Test-Case 'a minimal file parses' '{11111111-2222-3333-4444-555555555555}_is1' (Get-InstallerSpec $good).UninstallKeyName
    Test-Throws 'no AppId is refused' { Get-InstallerSpec ($good.Replace('AppId=', 'Id=')) } 'AppId'
    Test-Throws 'a [Registry] line it cannot read is refused, not skipped' { Get-InstallerSpec ($good.Replace('[Code]', 'Root: HKCU; Subkey: "x"; broken')) } 'cannot read'
    Test-Throws 'a registration that is not |vstolocal is refused' { Get-InstallerSpec ($good.Replace('|vstolocal', '')) } 'vstolocal'
    Test-Throws 'an HKLM line is refused' { Get-InstallerSpec ($good.Replace('Root: HKCU; Subkey: "Software\Microsoft\Office\16.0', 'Root: HKLM; Subkey: "Software\Microsoft\Office\16.0')) } 'HKLM'
    Test-Throws 'a constant other than {app} is refused' { Expand-SpecData -Template '{localappdata}\x' -AppDir 'C:\a' } '{app}'

    Write-Host ''
    Write-Host '== manifest values, trust URLs and keys (the guest script''s measured rules) =='
    Test-Case 'the installer spelling reads back as a path' 'C:\Users\u\AppData\Local\OutlookAI\Setup\OutlookAI.vsto' (ConvertFrom-ManifestValue 'file:///C:\Users\u\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal')
    Test-Case 'the build task''s spelling reads back as the same path' 'C:\Source\SixFive7\OutlookAI\bin\Release\OutlookAI.vsto' (ConvertFrom-ManifestValue 'file:///C:/Source/SixFive7/OutlookAI/bin/Release/OutlookAI.vsto|vstolocal')
    Test-Case 'an escaped space is unescaped' 'C:\a b\OutlookAI.vsto' (ConvertFrom-ManifestValue 'file:///C:/a%20b/OutlookAI.vsto|vstolocal')
    Test-Case 'the trust URL of the installer spelling is what the runtime wrote on this workstation' 'file:///C:/Users/u/AppData/Local/OutlookAI/Setup/OutlookAI.vsto' (ConvertTo-InclusionUrl 'file:///C:\Users\u\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal')
    Test-Case 'the trust URL of a plain path' 'file:///C:/a b/OutlookAI.vsto' (ConvertTo-InclusionUrl 'C:\a b\OutlookAI.vsto')
    Test-Case 'URLs compare case-insensitively, as the runtime does on a DOS path' $true (Test-SameInclusionUrl 'file:///c:/USERS/x.vsto' 'file:///C:/Users/x.vsto')
    Test-Case 'a different folder is a different URL' $false (Test-SameInclusionUrl 'file:///C:/a/x.vsto' 'file:///C:/b/x.vsto')
    $k1 = '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>'
    $k2 = '<RSAKeyValue><Modulus>BQYHCAk=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>'
    Test-Case 'a leading zero byte is the same key' $true (Test-SameRsaKey '<RSAKeyValue><Modulus>AQIDBA==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>' $k1)
    Test-Case 'another modulus is another key' $false (Test-SameRsaKey $k1 $k2)
    $fakeManifest = { param($key) "<asmv1:assembly xmlns:asmv1=""urn:schemas-microsoft-com:asm.v1""><Signature xmlns=""http://www.w3.org/2000/09/xmldsig#""><KeyInfo><KeyValue>$key</KeyValue></KeyInfo></Signature></asmv1:assembly>" }
    Test-Case 'the signing key is read out of a manifest' $k1 (Get-ManifestSigningKeyXml -ManifestXml (& $fakeManifest $k1))
    Test-Throws 'an unsigned manifest is refused' { Get-ManifestSigningKeyXml -ManifestXml '<a/>' } 'not signed'
    Test-Case 'a folder name carries the time and twelve hex of the commit' '20260927-153000-0123456789ab' (New-DevBuildFolderName -CommitSha '0123456789abcdef0123456789abcdef01234567' -When (New-Object DateTime(2026, 9, 27, 15, 30, 0)))
    Test-Case 'LoadBehavior 2 is read as a failed load' $true ((Get-LoadBehaviorMeaning 2).Contains('fails'))
    Test-Case 'a hard-disable blob naming the add-in is found' $true (Test-DisabledItemsMention @(, ([System.Text.Encoding]::Unicode.GetBytes('file:///C:/x/OutlookAI.vsto|vstolocal'))))

    Write-Host ''
    Write-Host '== classification =='
    Test-Case 'the installed manifest is the installed release' 'installed' (Get-BuildClass -VstoPath 'C:\U\Setup\OutlookAI.vsto' -InstalledVsto 'c:\u\setup\OutlookAI.vsto' -DevRoot 'C:\U\DevBuilds')
    Test-Case 'a folder under DevBuilds is a dev build' 'dev-build' (Get-BuildClass -VstoPath 'C:\U\DevBuilds\x\OutlookAI.vsto' -InstalledVsto 'C:\U\Setup\OutlookAI.vsto' -DevRoot 'C:\U\DevBuilds')
    Test-Case 'bin\Release is a build folder' 'build-folder' (Get-BuildClass -VstoPath 'C:\Source\OutlookAI\bin\Release\OutlookAI.vsto' -InstalledVsto 'C:\U\Setup\OutlookAI.vsto' -DevRoot 'C:\U\DevBuilds')
    Test-Case 'a sibling that merely starts with the DevBuilds name is not under it' 'build-folder' (Get-BuildClass -VstoPath 'C:\U\DevBuildsX\OutlookAI.vsto' -InstalledVsto '' -DevRoot 'C:\U\DevBuilds')
    Test-Case 'nothing registered is none' 'none' (Get-BuildClass -VstoPath $null -InstalledVsto 'C:\U\Setup\OutlookAI.vsto' -DevRoot 'C:\U\DevBuilds')

    Write-Host ''
    Write-Host '== the build guards (Testbed/host/Publish-AddInPayload.ps1''s, restated) =='
    Test-Case 'guard 1: the chain is the VSTO one minus registration' 'CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild' (Get-PrepareForRunOverride)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml((Get-NoHostWriteTargets))
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('m', 'http://schemas.microsoft.com/developer/msbuild/2003')
    $tasks = @($doc.SelectNodes('//m:UsingTask', $ns))
    Test-Case 'guard 2: three stand-ins' 3 $tasks.Count
    foreach ($t in $tasks) {
        $name = $t.GetAttribute('TaskName')
        Test-Case "$name overrides" 'true' $t.GetAttribute('Override')
        $declared = @($t.SelectNodes('m:ParameterGroup/*', $ns) | ForEach-Object { $_.LocalName })
        $missing = @($StandInTasks[$name].Parameters | Where-Object { $declared -notcontains $_ })
        Test-Case "$name declares every parameter" '' ($missing -join ',')
        $code = $t.SelectSingleNode('m:Task/m:Code', $ns).InnerText
        Test-Case "$name only logs its sentinel" $true ($code.Contains("$StandInSentinel`: $name did not run") -and -not $code.Contains('Registry') -and -not $code.Contains('File.'))
    }
    $installedTargets = @(Find-InstalledVstoTargets)
    if ($installedTargets.Count -gt 0) {
        $vsto = [xml](Get-Content -LiteralPath $installedTargets[0] -Raw)
        foreach ($name in $StandInTasks.Keys) {
            $used = @()
            foreach ($node in $vsto.GetElementsByTagName($name)) { foreach ($attr in $node.Attributes) { if ($attr.Name -ne 'Condition') { $used += $attr.Name } } }
            $missing = @($used | Sort-Object -Unique | Where-Object { $StandInTasks[$name].Parameters -notcontains $_ })
            Test-Case "$name - the installed targets pass nothing the stand-ins lack" '' ($missing -join ',')
        }
        Write-Host "       (read: $($installedTargets[0]))"
    }
    else { Write-Host '  SKIP the installed-targets comparison: no Visual Studio with the Office workload here.' }
    $a1 = Get-MSBuildArgumentList -ProjectPath 'C:\b\source\OutlookAI.csproj' -Thumbprint 'ABCDEF' -StandInTargets 'C:\b\NoHostWrite.targets' -FileLog 'C:\b\msbuild.log'
    Test-Case 'the chain is one quoted global property with a real semicolon' $true ($a1 -contains '/p:PrepareForRunDependsOn="CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild"')
    Test-Case 'DefineConstants semicolons are escaped' $true ($a1 -contains '/p:DefineConstants=VSTO40%3BTRACE')
    Test-Case 'the IDE''s registering behaviour is pinned off' $true ($a1 -contains '/p:BuildingInsideVisualStudio=false')
    Test-Case 'the dev version is stamped' $true ($a1 -contains "/p:ApplicationVersion=$DevVersion")
    Test-Case 'nothing names the registration target' $false (($a1 -join ' ').Contains('RegisterOfficeAddin'))
    Test-Case 'an evaluated chain without registration is accepted' 0 @(Test-EvaluatedPrepareForRun "CopyFilesToOutputDirectory;`r`n VisualStudioForApplicationsBuild;").Count
    Test-Case 'a chain still carrying RegisterOfficeAddin is refused' $true ((Test-EvaluatedPrepareForRun 'CopyFilesToOutputDirectory;VisualStudioForApplicationsBuild;RegisterOfficeAddin') -join ' ').Contains('did not take')
    $vsfab = 'Done building target "VisualStudioForApplicationsBuild" in project "OutlookAI.csproj".'
    $goodLog = "$StandInSentinel`: SetInclusionListEntry did not run on the build host; nothing was written.`r`n$StandInSentinel`: RegisterFormRegions did not run on the build host; nothing was written.`r`n$vsfab"
    Test-Case 'a log with both scheduled stand-ins is accepted' $true (Test-StandInSentinels $goodLog).Ok
    Test-Case 'a log without them is refused' $false (Test-StandInSentinels $vsfab).Ok
    Test-Case 'the registration target in the log is refused' $false (Test-StandInSentinels ($goodLog + "`r`nTarget ""RegisterOfficeAddin"" in file")).Ok
    Test-Case 'a declared form region is found' '[FormRegionMessageClass' ((Find-FormRegionDeclarations @('[FormRegionMessageClass("IPM.Note")]')) -join ',')
    Test-Case 'the build directory is a footprint in forward-slash form' 1 @(Find-BuildFootprint -Lines @('HKCU\Software\Microsoft\VSTO\Security\Inclusion\{g}|Url=String:file:///C:/w/.work/switch-addin-build/x/source/bin/Release/OutlookAI.vsto') -BuildDirectory 'C:\w\.work\switch-addin-build\x' -Thumbprint '' -ModulusBase64 '').Count
    Test-Case 'a WorkRoot elsewhere in the working tree is refused' $true (Test-IsUnderGitWorkTreeButNotWork -Path 'C:\r\Tools\out' -WorkTreeRoot 'C:\r')
    Test-Case 'a WorkRoot under .work is allowed' $false (Test-IsUnderGitWorkTreeButNotWork -Path 'C:\r\.work\switch-addin-build' -WorkTreeRoot 'C:\r')

    Write-Host ''
    Write-Host '== OutlookAI.csproj keeps its own Q81 guard =='
    $projText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot 'OutlookAI.csproj'))
    $proj = New-Object System.Xml.XmlDocument
    $proj.LoadXml($projText)
    $pns = New-Object System.Xml.XmlNamespaceManager($proj.NameTable)
    $pns.AddNamespace('m', 'http://schemas.microsoft.com/developer/msbuild/2003')
    $vstoImport = $proj.SelectSingleNode("//m:Import[contains(@Project,'Microsoft.VisualStudio.Tools.Office.targets')]", $pns)
    Test-Case 'the project imports the VSTO targets' $true ($null -ne $vstoImport)
    foreach ($name in @('SetOffice2007AddInRegistration', 'SetInclusionListEntry')) {
        $ut = $proj.SelectSingleNode("//m:UsingTask[@TaskName='$name']", $pns)
        Test-Case "the project stands in $name" $true ($null -ne $ut)
        if ($null -eq $ut -or $null -eq $vstoImport) { continue }
        Test-Case "$name - only outside Visual Studio" "'`$(BuildingInsideVisualStudio)' != 'true'" $ut.GetAttribute('Condition')
        Test-Case "$name - NOT an override, so this script's and the testbed's stand-ins still win" '' $ut.GetAttribute('Override')
        $precedes = $false
        $n = $ut
        while ($null -ne $n) { if ($n -eq $vstoImport) { $precedes = $true; break }; $n = $n.NextSibling }
        Test-Case "$name - declared BEFORE the VSTO import, so the first-declared rule makes it win" $true $precedes
        $declared = @($ut.SelectNodes('m:ParameterGroup/*', $pns) | ForEach-Object { $_.LocalName })
        $missing = @($StandInTasks[$name].Parameters | Where-Object { $declared -notcontains $_ })
        Test-Case "$name - declares every parameter the targets pass" '' ($missing -join ',')
    }
    Test-Case 'the project takes RegisterOfficeAddin off PrepareForRun, unescaped' $true ($projText.Contains("<PrepareForRunDependsOn>`$([MSBuild]::Unescape(`$(PrepareForRunDependsOn.Replace('RegisterOfficeAddin', ''))))</PrepareForRunDependsOn>"))
    Test-Case 'and the clean-time unregistration off VSTOClean' $true ($projText.Contains("UnregisterOfficeAddin', ''))))</VSTOCleanDependsOn>"))

    Write-Host ''
    Write-Host '== every registry write, against a scratch key =='
    $realKeys = @($AddinKeyRel, $InclusionKeyRel, ($UninstallRootRel + '\' + $spec.UninstallKeyName)) + @($spec.DoNotDisable | ForEach-Object { $_.Subkey })
    $realBefore = Get-RegistryLines -Keys $realKeys
    $runId = [guid]::NewGuid().ToString('N')
    $scratchKey = "$SelfTestRootRel\$runId"
    $scratchDir = Join-Path ($WorkRoot + '-selftest') $runId
    $script:SelfTestActive = $true
    try {
        Test-Throws 'while it runs, a write to the REAL registration is refused before it happens' { Set-KeyValue $AddinKeyRel 'Manifest' 'String' 'x' } 'SELF-TEST SAFETY'
        Test-Throws 'and so is one to the real trust store' { Remove-KeyTree $InclusionKeyRel } 'SELF-TEST SAFETY'
        $L = New-Layout -ScratchRoot $scratchKey
        $setup = Join-Path $scratchDir 'Setup'
        $publish = Join-Path $scratchDir 'publish'
        $devRoot = Join-Path $scratchDir 'DevBuilds'
        foreach ($d in @($setup, $publish)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
        [System.IO.File]::WriteAllText((Join-Path $setup $ManifestFileName), (& $fakeManifest $k1))
        [System.IO.File]::WriteAllText((Join-Path $setup 'OutlookAI.dll'), 'not a real assembly')
        [System.IO.File]::WriteAllText((Join-Path $publish $ManifestFileName), (& $fakeManifest $k2))
        [System.IO.File]::WriteAllText((Join-Path $publish 'OutlookAI.dll.manifest'), '<x/>')
        [System.IO.File]::WriteAllText((Join-Path $publish 'OutlookAI.dll'), 'not a real assembly')
        New-Item -ItemType Directory -Force -Path (Join-Path $publish 'Application Files\OutlookAI_99_99_99_0') | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $publish 'Application Files\OutlookAI_99_99_99_0\OutlookAI.dll.deploy'), 'x')

        # Today's workstation, in miniature: registered at a build folder, the installed release
        # and two build folders trusted, another add-in's entry, one exemption missing.
        $uninstall = $L.UninstallRoot + '\' + $spec.UninstallKeyName
        Set-KeyValue $uninstall 'InstallLocation' 'String' ($setup + '\')
        Set-KeyValue $uninstall 'DisplayVersion' 'String' '3.0.1.321'
        $buildUrl = 'file:///C:/Source/SixFive7/OutlookAI/bin/Release/OutlookAI.vsto'
        Set-KeyValue $L.AddinKey 'Manifest' 'String' ($buildUrl + '|vstolocal')
        Set-KeyValue $L.AddinKey 'LoadBehavior' 'DWord' 3
        Set-KeyValue $L.AddinKey 'FriendlyName' 'String' 'OutlookAI'
        Set-KeyValue $L.AddinKey 'Description' 'String' 'OutlookAI'
        $eSetup = (Add-InclusionEntry $L.InclusionKey (ConvertTo-InclusionUrl (Join-Path $setup $ManifestFileName)) $k1).Name
        $first = (Add-InclusionEntry $L.InclusionKey $buildUrl $k2).Name
        $second = Add-InclusionEntry $L.InclusionKey ($buildUrl.Replace('/C:/', '/c:/')) $k1
        $eRelease = $second.Name
        $eDebug = (Add-InclusionEntry $L.InclusionKey 'file:///C:/Source/SixFive7/OutlookAI/bin/Debug/OutlookAI.vsto' $k1).Name
        $eOther = (Add-InclusionEntry $L.InclusionKey 'file:///C:/Users/u/source/repos/Other/bin/Debug/Other.vsto' $k2).Name
        foreach ($d in $spec.DoNotDisable) { if ($d.Version -ne '15.0') { Set-KeyValue ($L.Prefix + $d.Subkey) $d.Name 'DWord' 1 } }
        Test-Case 'a second entry for the same URL (another case) replaces the first, as the runtime does' "$first|1" ("{0}|{1}" -f ($second.Removed -join ','), @(Get-InclusionEntries $L.InclusionKey | Where-Object { Test-SameInclusionUrl $_.Url $buildUrl }).Count)

        $s0 = Get-StatusReport -Layout $L -Spec $spec -DevRoot $devRoot
        Test-Case 'status: registered at a build folder' 'build-folder' $s0.Class
        Test-Case 'status: the installed release is found from its uninstall entry' '3.0.1.321' $s0.Installed.DisplayVersion
        Test-Case 'status: the installed release loads without a prompt' $true $s0.InstalledTrust.Trusted
        Test-Case 'status: another add-in''s entry is counted, not listed' 1 $s0.OtherAddinEntries

        # Guard 3's repair path: a build that repointed the registration and trusted itself.
        $beforeBuild = Get-RegistryLines -Keys @($L.AddinKey)
        Set-KeyValue $L.AddinKey 'Manifest' 'String' 'file:///C:/w/.work/switch-addin-build/x/source/bin/Release/OutlookAI.vsto|vstolocal'
        Set-KeyValue $L.AddinKey 'Extra' 'String' 'added by a build'
        $null = Add-InclusionEntry $L.InclusionKey 'file:///C:/w/.work/switch-addin-build/x/source/bin/Release/OutlookAI.vsto' $k2
        Restore-KeyFromLines -BeforeLines $beforeBuild -KeyPath $L.AddinKey
        Test-Case 'guard 3 repair: the registration is put back value for value' (($beforeBuild) -join ';') ((Get-RegistryLines -Keys @($L.AddinKey)) -join ';')
        $gone = @(Remove-BuildTrustFootprints -InclusionKey $L.InclusionKey -BuildDirectory 'C:\w\.work\switch-addin-build\x')
        Test-Case 'guard 3 repair: the trust entry naming the build directory is removed, and only it' '1 4' ("{0} {1}" -f $gone.Count, @(Get-InclusionEntries $L.InclusionKey).Count)

        # Install a dev build from a (fake) publish folder.
        $prov = [pscustomobject]@{ Commit = ('a' * 40); Ref = 'HEAD'; Subject = 'self-test'; BuiltUtc = [DateTime]::UtcNow.ToString('o'); HostProof = [ordered]@{ linesCompared = 0 } }
        $inst = Install-DevBuild -Layout $L -Spec $spec -PublishDir $publish -SigningKeyXml $k2 -DevRoot $devRoot -FolderName 'selftest-build' -Provenance $prov -ExpectedDllVersion ''
        $devVsto = Join-Path (Join-Path $devRoot 'selftest-build') $ManifestFileName
        $reg = Get-KeyValues $L.AddinKey
        Test-Case 'install: Manifest names the DevBuilds copy, in the installer''s spelling' ('file:///' + $devVsto + '|vstolocal') $reg['Manifest'].Data
        Test-Case 'install: LoadBehavior is a DWORD 3' 'DWord 3' ("{0} {1}" -f $reg['LoadBehavior'].Kind, $reg['LoadBehavior'].Data)
        Test-Case 'install: FriendlyName and Description are the installer''s' 'OutlookAI|OutlookAI' ("{0}|{1}" -f $reg['FriendlyName'].Data, $reg['Description'].Data)
        Test-Case 'install: the whole publish tree was copied' $true (Test-Path -LiteralPath (Join-Path (Join-Path $devRoot 'selftest-build') 'Application Files\OutlookAI_99_99_99_0\OutlookAI.dll.deploy'))
        $entry = @(Get-InclusionEntries $L.InclusionKey | Where-Object { $_.Name -eq $inst.TrustEntry })
        Test-Case 'install: one trust entry for the copy''s URL' (ConvertTo-InclusionUrl $devVsto) ($entry | ForEach-Object { $_.Url })
        Test-Case 'install: with the build''s key' $true (Test-SameRsaKey ($entry | ForEach-Object { $_.PublicKey }) $k2)
        $dnd = @(Get-DoNotDisableChanges $L $spec)
        Test-Case 'install: the missing exemption was written, the others left' 0 @($dnd | Where-Object { -not $_.Same }).Count
        Test-Case 'install: HKCU\Software\OutlookAI was never written' $false (Test-Path -LiteralPath ("Registry::HKEY_CURRENT_USER\" + $L.Prefix + 'Software\OutlookAI'))
        $rec = Get-Content -LiteralPath $inst.Record -Raw | ConvertFrom-Json
        Test-Case 'record: status registered' 'registered' $rec.status
        Test-Case 'record: it names the registration it replaced' ($buildUrl + '|vstolocal') $rec.replaced.values.Manifest.data
        Test-Case 'record: and what that was' 'build-folder' $rec.replaced.was
        Test-Case 'record: and the trust entry it wrote' $inst.TrustEntry $rec.trust.entry
        $s1 = Get-StatusReport -Layout $L -Spec $spec -DevRoot $devRoot
        Test-Case 'status: now a dev build' 'dev-build' $s1.Class
        Test-Case 'status: which loads without a prompt' $true $s1.Trust.Trusted
        Test-Case 'status: and names its commit' ('a' * 40) $s1.Record.commit
        Test-Throws 'install: the same folder is never reused' { Install-DevBuild -Layout $L -Spec $spec -PublishDir $publish -SigningKeyXml $k2 -DevRoot $devRoot -FolderName 'selftest-build' -Provenance $prov -ExpectedDllVersion '' } 'already exists'
        Test-Throws 'install: a build signed by another key is refused' { Install-DevBuild -Layout $L -Spec $spec -PublishDir $publish -SigningKeyXml $k1 -DevRoot $devRoot -FolderName 'selftest-other' -Provenance $prov -ExpectedDllVersion '' } 'not signed by'
        New-Item -ItemType Directory -Force -Path (Join-Path $publish 'McpServer') | Out-Null
        Test-Throws 'install: a build that carries the MCP server is refused' { Install-DevBuild -Layout $L -Spec $spec -PublishDir $publish -SigningKeyXml $k2 -DevRoot $devRoot -FolderName 'selftest-mcp' -Provenance $prov -ExpectedDllVersion '' } 'McpServer'
        Remove-Item -LiteralPath (Join-Path $publish 'McpServer') -Recurse -Force

        # Restore: dry run first, which must write nothing.
        $scratchBefore = Get-RegistryLines -Keys @($scratchKey)
        $plan = Invoke-Restore -Layout $L -Spec $spec -DevRoot $devRoot
        Test-Case 'restore dry run: writes nothing' (($scratchBefore) -join ';') ((Get-RegistryLines -Keys @($scratchKey)) -join ';')
        Test-Case 'restore dry run: one registration value to change' 'Manifest' (($plan.Changes.Values | Where-Object { -not $_.Same } | ForEach-Object { $_.Name }) -join ',')
        Test-Case 'restore dry run: the dev build''s trust entry goes' $inst.TrustEntry (($plan.Remove | ForEach-Object { $_.Name }) -join ',')
        Test-Case 'restore dry run: the build folders'' entries stay without -RemoveBuildTrust' 3 $plan.Keep.Count
        $done = Invoke-Restore -Layout $L -Spec $spec -DevRoot $devRoot -Apply
        $reg = Get-KeyValues $L.AddinKey
        Test-Case 'restore: Manifest is the installed release, exactly as Inno writes it' ('file:///' + $setup + '\OutlookAI.vsto|vstolocal') $reg['Manifest'].Data
        Test-Case 'restore: the dev build''s trust entry is gone' 0 @(Get-InclusionEntries $L.InclusionKey | Where-Object { $_.Name -eq $inst.TrustEntry }).Count
        Test-Case 'restore: the installed release''s, the build folders'' and the other add-in''s entries stay' 4 @(Get-InclusionEntries $L.InclusionKey).Count
        Test-Case 'restore: the audit record was written' $true ([bool]$done.Audit -and (Test-Path -LiteralPath $done.Audit))
        Test-Case 'restore: the dev build''s own record says restored' 'restored' (Get-Content -LiteralPath $inst.Record -Raw | ConvertFrom-Json).status
        $s2 = Get-StatusReport -Layout $L -Spec $spec -DevRoot $devRoot
        Test-Case 'status: the installed release again' 'installed' $s2.Class
        Test-Case 'status: loading without a prompt' $true $s2.Trust.Trusted
        $again = Invoke-Restore -Layout $L -Spec $spec -DevRoot $devRoot -Apply
        Test-Case 'restore again: nothing left to change' 0 $again.Changes.Pending
        $withBuild = Invoke-Restore -Layout $L -Spec $spec -DevRoot $devRoot -RemoveBuildTrustEntries -Apply
        Test-Case '-RemoveBuildTrust: removes bin\Release and bin\Debug' (@($eRelease, $eDebug) | Sort-Object) (@($withBuild.Remove | ForEach-Object { $_.Name }) | Sort-Object)
        $left = @(Get-InclusionEntries $L.InclusionKey | ForEach-Object { $_.Name } | Sort-Object)
        Test-Case '-RemoveBuildTrust: keeps the installed release''s and never touches another add-in''s' (@($eSetup, $eOther) | Sort-Object) $left

        Remove-KeyTree $uninstall
        Test-Throws 'restore refuses with no uninstall entry' { Invoke-Restore -Layout $L -Spec $spec -DevRoot $devRoot } 'no installed release'
        Set-KeyValue $uninstall 'InstallLocation' 'String' ($scratchDir + '\NotThere\')
        Test-Throws 'restore refuses when the installed manifest is missing' { Invoke-Restore -Layout $L -Spec $spec -DevRoot $devRoot } 'does not exist'
    }
    finally {
        $script:SelfTestActive = $false
        [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($scratchKey, $false)
        $parent = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($SelfTestRootRel, $false)
        if ($null -ne $parent) {
            $empty = ($parent.SubKeyCount -eq 0 -and $parent.ValueCount -eq 0)
            $parent.Close()
            if ($empty) { [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($SelfTestRootRel, $false) }
        }
        if (Test-Path -LiteralPath $scratchDir) { Remove-Item -LiteralPath $scratchDir -Recurse -Force }
        $selfRoot = $WorkRoot + '-selftest'
        if ((Test-Path -LiteralPath $selfRoot) -and @(Get-ChildItem -LiteralPath $selfRoot -Force).Count -eq 0) { Remove-Item -LiteralPath $selfRoot -Force }
    }
    Test-Case 'the scratch key is gone' $false (Test-Path -LiteralPath ("Registry::HKEY_CURRENT_USER\$scratchKey"))
    Test-Case 'the scratch folder is gone' $false (Test-Path -LiteralPath $scratchDir)
    $realAfter = Get-RegistryLines -Keys $realKeys
    $realDiff = Compare-Lines -Before $realBefore -After $realAfter
    Test-Case "the REAL registration, trust, uninstall and exemption keys are identical ($(@($realBefore).Count) lines)" $true $realDiff.Same

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a real run can settle these:'
    Write-Host '  * that a build succeeds and leaves the host unchanged     (-BuildOnly settles it)'
    Write-Host '  * that Outlook loads the registered dev build with no trust prompt, at its next start'
    Write-Host '  * that Outlook loads the installed release again after -Restore'
    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# =============================================================================================
# MAIN
# =============================================================================================
$modes = @()
if ($Restore) { $modes += '-Restore' }
if ($Status) { $modes += '-Status' }
if ($BuildOnly) { $modes += '-BuildOnly' }
if ($SelfTest) { $modes += '-SelfTest' }
if ($modes.Count -gt 1) { throw "Pick one of $($modes -join ', ')." }
if ($CommitGiven -and ($Restore -or $Status -or $SelfTest)) { throw '-Commit belongs to putting a dev build on (the default mode) and to -BuildOnly.' }
if ($Execute -and ($Status -or $SelfTest -or $BuildOnly)) { throw '-Execute applies only to putting a dev build on and to -Restore. -Status and -SelfTest never write, and -BuildOnly never registers anything.' }
if ($RemoveBuildTrust -and -not $Restore) { throw '-RemoveBuildTrust applies only to -Restore.' }

if ($SelfTest) { exit (Invoke-SelfTest) }

if (Test-IsElevated) {
    throw 'REFUSING TO RUN ELEVATED. An elevated process can carry another account''s HKCU, and this script only ever changes the Outlook of the user running it. Run it from an ordinary, unelevated shell.'
}
if (Test-PathUnder $DevBuildsRoot $RepoRoot) { throw "REFUSING: -DevBuildsRoot $DevBuildsRoot is inside the repository; dev builds live outside every working tree." }
if ((Test-PathUnder $WorkRoot $DevBuildsRoot) -or (Test-PathUnder $DevBuildsRoot $WorkRoot)) { throw 'REFUSING: -WorkRoot and -DevBuildsRoot overlap. No build may ever write into DevBuilds.' }

$spec = Get-InstallerSpec ([System.IO.File]::ReadAllText((Join-Path $RepoRoot 'Installer.iss')))
$layout = New-Layout ''
$outlookPids = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$restartNote = 'Outlook is not running; it picks this up when it next starts.'
if ($outlookPids.Count -gt 0) { $restartNote = "Outlook is running (pid $($outlookPids -join ', ')) and keeps what it loaded until it restarts. This script never starts, quits or kills Outlook: the change takes effect at the next Outlook restart." }

if ($Status) {
    Write-StatusReport (Get-StatusReport -Layout $layout -Spec $spec -DevRoot $DevBuildsRoot -IncludeProcesses)
    exit 0
}

if ($Restore) {
    if (-not $Execute) {
        Say 'DRY RUN of -Restore. Nothing is written. With -Execute this would put the installed release back:'
        Write-RestorePlan (Invoke-Restore -Layout $layout -Spec $spec -DevRoot $DevBuildsRoot -RemoveBuildTrustEntries:$RemoveBuildTrust)
        Say ''
        Say $restartNote
        exit 0
    }
    Say '== -Restore -Execute: putting the installed release back =='
    $plan = Invoke-Restore -Layout $layout -Spec $spec -DevRoot $DevBuildsRoot -RemoveBuildTrustEntries:$RemoveBuildTrust -Apply
    Write-RestorePlan $plan
    if ($plan.Audit) { Say "  audit record  $($plan.Audit)" }
    Say ''
    Say "DONE. Outlook will load the installed release $($plan.Installed.DisplayVersion) at its next start."
    Say $restartNote
    exit 0
}

# ---- Putting a dev build on, or -BuildOnly -----------------------------------------------------
$commitInfo = Resolve-Commit $Commit
$dirty = @(Invoke-NativeCommand { & git -C $RepoRoot status --porcelain })
if ($dirty.Count -gt 0) { Say "NOTE the working tree has $($dirty.Count) uncommitted change(s). They are NOT in the build - it is built from $($commitInfo.Sha) on purpose." }
$workTree = (@(Invoke-NativeCommand { & git -C $RepoRoot rev-parse --show-toplevel 2>$null }) | Out-String).Trim().Replace('/', '\')
if (Test-IsUnderGitWorkTreeButNotWork -Path $WorkRoot -WorkTreeRoot $workTree) { throw "REFUSING: -WorkRoot $WorkRoot is inside the git working tree $workTree but not under .work\." }
$folderName = New-DevBuildFolderName -CommitSha $commitInfo.Sha -When (Get-Date)
$buildDir = Join-Path $WorkRoot $folderName

if (-not $BuildOnly -and -not $Execute) {
    $current = Get-KeyValues $layout.AddinKey
    $devFolder = Join-Path $DevBuildsRoot $folderName
    $changes = Get-RegistrationChanges $current (Get-TargetRegistration $spec $devFolder)
    Say "DRY RUN. Nothing is built, copied, trusted or registered. With -Execute this would:"
    Say "  1. build     $($commitInfo.Sha) - $($commitInfo.Subject)"
    Say "               in $buildDir, from a git archive, with all three host guards and a throwaway signing key"
    Say "  2. copy      the flattened publish output to $devFolder"
    Say "  3. record    the registration it replaces in $(Join-Path $devFolder $RecordFileName)"
    Say "  4. trust     HKCU\$InclusionKeyRel\<new guid>  Url = $(ConvertTo-InclusionUrl (Join-Path $devFolder $ManifestFileName))"
    Say '               PublicKey = the build''s signing key'
    Say "  5. register  HKCU\$AddinKeyRel"
    foreach ($c in $changes.Values) {
        if ($c.Same) { Say ("               {0,-13} {1}  (unchanged)" -f $c.Name, $c.After) }
        else { Say ("               {0,-13} {1}`n               {2,-13} -> {3}" -f $c.Name, $c.Before, '', $c.After) }
    }
    foreach ($d in @(Get-DoNotDisableChanges $layout $spec)) {
        if ($d.Same) { Say "               DoNotDisableAddinList $($d.Version) unchanged" } else { Say "               DoNotDisableAddinList $($d.Version) $($d.Before) -> $($d.Data)" }
    }
    Say ''
    Say "Toolchain: $(Resolve-MSBuild)"
    Say $restartNote
    Say 'Taking it off again later: -Restore -Execute.'
    exit 0
}

Say "== Building $($commitInfo.Sha) with the host guards =="
$built = Invoke-GuardedBuild -CommitInfo $commitInfo -BuildDir $buildDir
if ($BuildOnly) {
    Say ''
    Say "BUILT. $($built.PublishDir)"
    Say 'Nothing was registered or copied, and the host is as it was. Delete the folder when you are done with it.'
    exit 0
}

Say ''
Say "== Putting $folderName on this machine's Outlook =="
$prov = [pscustomobject]@{ Commit = $commitInfo.Sha; Ref = $commitInfo.Ref; Subject = $commitInfo.Subject; BuiltUtc = [DateTime]::UtcNow.ToString('o'); HostProof = $built.HostProof }
$installed = Install-DevBuild -Layout $layout -Spec $spec -PublishDir $built.PublishDir -SigningKeyXml $built.SigningKeyXml -DevRoot $DevBuildsRoot -FolderName $folderName -Provenance $prov -ExpectedDllVersion $DevVersion
Say "  copied      $($installed.Folder)"
Say "  record      $($installed.Record)"
Say "  trust       HKCU\$InclusionKeyRel\$($installed.TrustEntry)"
foreach ($c in $installed.Changes.Values) {
    if ($c.Same) { Say ("  {0,-11} {1}  (unchanged)" -f $c.Name, $c.After) } else { Say ("  {0,-11} {1} -> {2}" -f $c.Name, $c.Before, $c.After) }
}
Say ''
Say "DONE. Outlook will load the dev build of $($commitInfo.Sha) at its next start."
Say $restartNote
Say 'Taking it off again: -Restore -Execute.'
exit 0
