#Requires -Version 5.1
<#
.SYNOPSIS
    Fills the unattended-install answer file with the real guest credential and packages it as a
    small ISO for Hyper-V. Writes into gitignored scratch only.

.DESCRIPTION
    THE TEMPLATE IN THE REPOSITORY HAS NO PASSWORD IN IT AND MUST NEVER HAVE ONE. This
    repository is public and a guest password has already been published from it once; it
    survives in git history and had to be rotated. So the shape is: a committed template
    (Testbed/guest/autounattend.template.xml) carrying placeholder tokens, this script to fill
    them, and an output path that is gitignored scratch. The filled file never goes back into
    Testbed/, and this script refuses to write it there even if asked.

    The credential comes from exactly one place, the one Testbed/host/Get-GuestCredential.ps1
    already documents: McpServer/OutlookAI.McpServer.Tests/live-fixtures/vm-credentials.json,
    gitignored. This script calls that loader rather than reading the file itself, so there is
    one definition of where a credential may live and no second mechanism to keep in step.

    WHY AN ISO AND NOT A VHD. Windows Setup looks for autounattend.xml at the root of removable
    media before it looks anywhere a person would have to configure, so any volume that presents
    as a removable drive is enough. Three reasons the removable one is an ISO on a second DVD
    drive rather than a small VHDX:

      * A VHDX attached to the VM is a FIXED disk, not removable media, so finding the answer
        file on it depends on behaviour Microsoft does not document. A DVD is the documented
        case.
      * A second disk makes "disk 0" ambiguous at exactly the moment the answer file is wiping
        disk 0 and creating partitions on it. Getting that wrong destroys the wrong volume, and
        the failure is silent until much later. A DVD drive is not a disk and cannot be
        confused with one.
      * Writing a file into a VHDX means creating, mounting, partitioning, formatting and
        dismounting it - which needs elevation and the Hyper-V or storage modules on the host.
        An ISO is built from a directory by a plain user with no mounting at all.

    HOW THE ISO IS BUILT, AND WHAT HAPPENS WHEN THE TOOL IS MISSING. Two builders, tried in this
    order, and the choice is reported rather than assumed:

      1. oscdimg.exe, from the Windows ADK's Deployment Tools. Microsoft's own image builder and
         the one their documentation assumes. Found on PATH, via the Windows Kits registry root,
         or in the usual install location.
      2. IMAPI2FS (the MsftFileSystemImage COM object), which ships with Windows itself and
         needs nothing installed. This is the fallback that keeps a fresh machine able to build
         the answer volume from this repository alone.

    If neither is usable the script THROWS AND NAMES WHAT IS MISSING. It does not write half an
    ISO: the artefact is verified after it is built (non-empty, and carrying the ISO 9660
    signature at the standard offset) and deleted if that check fails, because a broken answer
    volume presents as an install that stops on the language page for no visible reason.

    WHAT COMES OUT, ALL OF IT IN SCRATCH AND ALL OF IT HOLDING A CREDENTIAL:

        <OutDir>\answer-volume\autounattend.xml        the filled answer file
        <OutDir>\answer-volume\Complete-FirstLogon.ps1 the first-logon fix-ups, verbatim
        <OutDir>\<VMName>-unattend.iso                 the volume to attach to the VM

    The password is in the ISO in clear text, because the answer file declares PlainText. That
    is deliberate and it is the only honest option here: the alternative Windows offers is a
    base64 encoding with a documented salt suffix, which is obfuscation rather than encryption
    and would be untested code standing between a rebuild and a working guest. Treat the ISO as
    a secret, keep it in scratch, and delete it once the guest is built. The script says so on
    the way out.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script. Used to find the credential and
    to enforce where output may go.

.PARAMETER VMName
    Which guest this answer file is for. Passed to Get-GuestCredential.ps1, which refuses a
    credential recorded for a different VM, and used to name the ISO.

    TWO GUESTS, ONE CREDENTIAL FILE. That refusal is per-name, so building answer volumes for
    both guests from a single vm-credentials.json needs its `vmName` left empty or absent - the
    loader only enforces the match when the field has a value. The alternative is one credential
    file per guest, which means one of them is not at the documented path. Empty `vmName` is the
    intended route; the field exists to stop a credential being handed to a machine it was not
    meant for, and with one shared account there is no such machine.

.PARAMETER ComputerName
    The Windows computer name to give the guest. Defaults to the VM name with its 'OutlookAI-'
    prefix shortened to 'OAI-', which is what keeps it inside the 15 characters NetBIOS allows.
    Refuses rather than truncates if that is still too long.

.PARAMETER ImageName
    Which edition to install out of the multi-edition ISO. 'Windows 11 Pro' matches the host.

.PARAMETER TemplatePath
    The committed template. Defaults to Testbed/guest/autounattend.template.xml beside this
    script's checkout.

.PARAMETER FirstLogonScriptPath
    The first-logon script to put on the volume. Defaults to
    Testbed/guest/Complete-FirstLogon.ps1.

.PARAMETER OutDir
    Where the filled answer file and the ISO go. Defaults to .work/testbed-answer/<VMName>,
    which is gitignored. Anything inside the repository that is not under .work is refused.

.PARAMETER VolumeLabel
    ISO volume label. Uppercase, no spaces, 32 characters or fewer.

.PARAMETER IsoBuilder
    Auto (default), Oscdimg or Imapi. Force one when you want to know which produced an image.

.PARAMETER OscdimgPath
    An explicit oscdimg.exe, when it is somewhere this script would not look.

.PARAMETER KeepStaging
    Leave the staging directory in place. It holds the same credential the ISO does.

.EXAMPLE
    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName OutlookAI-TestVM

.EXAMPLE
    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName OutlookAI-Indexed -ComputerName OAI-INDEXED
#>
[CmdletBinding()]
param(
    [string] $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string] $VMName = 'OutlookAI-TestVM',
    [string] $ComputerName,
    [string] $ImageName = 'Windows 11 Pro',
    [string] $TemplatePath,
    [string] $FirstLogonScriptPath,
    [string] $OutDir,
    [string] $VolumeLabel = 'UNATTEND',
    [ValidateSet('Auto', 'Oscdimg', 'Imapi')] [string] $IsoBuilder = 'Auto',
    [string] $OscdimgPath,
    [switch] $KeepStaging
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$guestDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'guest'
if (-not $TemplatePath) { $TemplatePath = Join-Path $guestDir 'autounattend.template.xml' }
if (-not $FirstLogonScriptPath) { $FirstLogonScriptPath = Join-Path $guestDir 'Complete-FirstLogon.ps1' }
if (-not $OutDir) { $OutDir = Join-Path $RepoRoot (Join-Path '.work\testbed-answer' $VMName) }

if (-not $ComputerName) {
    # NetBIOS caps a computer name at 15 characters and Setup fails the specialize pass on a
    # longer one rather than truncating for you. Every testbed VM name starts 'OutlookAI-',
    # which alone is two thirds of the budget, so abbreviate that prefix rather than chop the
    # end off - 'OutlookAI-TestVM' truncated is 'OutlookAI-TestV', which reads like a typo
    # forever afterwards.
    $ComputerName = $VMName -replace '^OutlookAI-', 'OAI-'
    if ($ComputerName.Length -gt 15) {
        throw "Cannot derive a NetBIOS computer name from VM name '$VMName' - '$ComputerName' is $($ComputerName.Length) characters and the limit is 15. Pass -ComputerName explicitly."
    }
}
if ($ComputerName -notmatch '^[A-Za-z0-9-]{1,15}$') {
    throw "Computer name '$ComputerName' is not a legal NetBIOS name (letters, digits and hyphens, 15 characters or fewer)."
}
if ($VolumeLabel -notmatch '^[A-Z0-9_-]{1,32}$') {
    throw "Volume label '$VolumeLabel' is not usable. Use up to 32 uppercase letters, digits, underscores or hyphens."
}

foreach ($required in @($TemplatePath, $FirstLogonScriptPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Not found: $required"
    }
}

# ---------------------------------------------------------------------------------------------
# WHERE OUTPUT MAY GO. This guard is the whole reason the template is safe to commit, so it
# refuses rather than warns.
# ---------------------------------------------------------------------------------------------
# Resolve WITHOUT creating anything. Checking a path after making the directory would already
# have put a directory named after the answer file inside Testbed/ before refusing to fill it.
$outFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir).TrimEnd('\')
$repoFull = (Resolve-Path -LiteralPath $RepoRoot).ProviderPath.TrimEnd('\')
$scratchFull = Join-Path $repoFull '.work'

if (($outFull -split '\\') -contains 'Testbed') {
    throw "Refusing to write a filled answer file under Testbed/: $outFull. Testbed/ is committed and this repository is public."
}
if ($outFull.StartsWith($repoFull, [StringComparison]::OrdinalIgnoreCase) -and
    -not $outFull.StartsWith($scratchFull, [StringComparison]::OrdinalIgnoreCase)) {
    throw @"
Refusing to write inside the repository outside gitignored scratch.
    asked for : $outFull
    allowed   : $scratchFull\... (gitignored), or any path outside $repoFull
The output holds the guest password in clear text.
"@
}

New-Item -ItemType Directory -Force -Path $outFull | Out-Null

# ---------------------------------------------------------------------------------------------
# The credential, from the one place one may live.
# ---------------------------------------------------------------------------------------------
$credential = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$guestUser = $credential.UserName
$guestSecret = $credential.GetNetworkCredential().Password
if ([string]::IsNullOrWhiteSpace($guestSecret)) {
    throw 'The guest credential has an empty secret. Windows will not create a local account without one.'
}
if ($guestUser -notmatch '^[A-Za-z0-9._-]{1,20}$') {
    throw "Guest user name '$guestUser' is not a plain local account name. Refusing to build an answer file from it."
}
foreach ($ch in @('<', '>', '&')) {
    if ($guestSecret.Contains($ch)) {
        throw "The guest password contains '$ch', which would have to be XML-escaped and is not. Rotate it to a password without < > or &, or teach this script to escape."
    }
}

# ---------------------------------------------------------------------------------------------
# Fill the template.
# ---------------------------------------------------------------------------------------------
$templateText = Get-Content -LiteralPath $TemplatePath -Raw

$substitutions = [ordered]@{
    'GUEST_USERNAME' = $guestUser
    'GUEST_PASSWORD' = $guestSecret
    'COMPUTER_NAME'  = $ComputerName
    'IMAGE_NAME'     = $ImageName
}

$filled = $templateText
foreach ($name in $substitutions.Keys) {
    $token = '{{' + $name + '}}'
    if (-not $filled.Contains($token)) {
        throw "The template does not contain the token $token. It has changed shape, and filling it now would produce an answer file that is quietly wrong. Template: $TemplatePath"
    }
    $filled = $filled.Replace($token, $substitutions[$name])
}

$leftover = @([regex]::Matches($filled, '\{\{[A-Za-z0-9_]+\}\}') |
        ForEach-Object { $_.Value } | Sort-Object -Unique)
if ($leftover.Count -gt 0) {
    throw ('The template still has unfilled tokens after substitution: ' + ($leftover -join ', ') +
        ". Teach this script what they mean, or take them out of $TemplatePath.")
}

# Parse what is about to be written, not what was read: an answer file Setup cannot parse is
# indistinguishable, from the VM console, from one it never found.
$doc = New-Object System.Xml.XmlDocument
try { $doc.LoadXml($filled) }
catch { throw "The filled answer file is not well-formed XML: $($_.Exception.Message)" }

$passwordValues = @($doc.SelectNodes('//*') | Where-Object {
        $_.LocalName -eq 'Value' -and $_.ParentNode -and
        ($_.ParentNode.LocalName -eq 'Password' -or $_.ParentNode.LocalName -eq 'AdministratorPassword')
    })
if ($passwordValues.Count -lt 2) {
    throw "The filled answer file has $($passwordValues.Count) password value(s); expected the local account's and the autologon's. Template: $TemplatePath"
}
foreach ($node in $passwordValues) {
    if ([string]::IsNullOrWhiteSpace($node.InnerText)) {
        throw 'The filled answer file has an empty password value.'
    }
}

# ---------------------------------------------------------------------------------------------
# Stage the volume contents.
# ---------------------------------------------------------------------------------------------
$staging = Join-Path $outFull 'answer-volume'
if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force }
New-Item -ItemType Directory -Force -Path $staging | Out-Null

# UTF-8 WITH a byte order mark, which is what Windows System Image Manager itself writes and
# what Setup is known to accept.
$answerPath = Join-Path $staging 'autounattend.xml'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($answerPath, $filled, $utf8Bom)

Copy-Item -LiteralPath $FirstLogonScriptPath -Destination (Join-Path $staging 'Complete-FirstLogon.ps1') -Force

# ---------------------------------------------------------------------------------------------
# Build the ISO.
# ---------------------------------------------------------------------------------------------
function Find-Oscdimg {
    param([string] $Explicit)

    if ($Explicit) {
        if (Test-Path -LiteralPath $Explicit) { return (Resolve-Path -LiteralPath $Explicit).ProviderPath }
        throw "-OscdimgPath was given but there is nothing at: $Explicit"
    }

    $onPath = Get-Command 'oscdimg.exe' -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }

    $kitRoots = New-Object System.Collections.Generic.List[string]
    foreach ($key in @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots',
                       'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots')) {
        try {
            $value = (Get-ItemProperty -Path $key -Name 'KitsRoot10' -ErrorAction Stop).KitsRoot10
            if ($value) { $kitRoots.Add($value) }
        }
        catch { }
    }
    foreach ($guess in @("${env:ProgramFiles(x86)}\Windows Kits\10", "$env:ProgramFiles\Windows Kits\10")) {
        if ($guess) { $kitRoots.Add($guess) }
    }

    foreach ($root in ($kitRoots | Select-Object -Unique)) {
        foreach ($arch in @('amd64', 'arm64', 'x86')) {
            $candidate = Join-Path $root "Assessment and Deployment Kit\Deployment Tools\$arch\Oscdimg\oscdimg.exe"
            if (Test-Path -LiteralPath $candidate) { return (Resolve-Path -LiteralPath $candidate).ProviderPath }
        }
    }
    return $null
}

function New-IsoWithOscdimg {
    param([string] $Exe, [string] $SourceDir, [string] $IsoPath, [string] $Label)

    # -u1 writes both an ISO 9660 file system with 8.3 names and a UDF one with the real names.
    # Long names matter: 'autounattend.xml' does not survive 8.3, and a Setup that cannot see
    # that name behaves exactly as though no answer file were attached.
    $logPath = [System.IO.Path]::ChangeExtension($IsoPath, '.oscdimg.log')
    & $Exe '-u1' '-udfver102' ('-l' + $Label) $SourceDir $IsoPath *> $logPath
    if ($LASTEXITCODE -ne 0) {
        throw "oscdimg exited $LASTEXITCODE. Its output is in $logPath."
    }
}

function New-IsoWithImapi {
    param([string] $SourceDir, [string] $IsoPath, [string] $Label)

    if (-not ('OutlookAI.Testbed.IsoStreamWriter' -as [type])) {
        # Deliberately not /unsafe: the block count comes back through an unmanaged int that a
        # pinned pointer would normally supply, and Marshal does the same job without needing a
        # compiler option that Windows PowerShell and PowerShell 7 spell differently.
        Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OutlookAI.Testbed
{
    public static class IsoStreamWriter
    {
        public static void Write(string path, object imageStream, int blockSize, int totalBlocks)
        {
            IStream source = imageStream as IStream;
            if (source == null) { throw new ArgumentException("The image result is not an IStream.", "imageStream"); }
            IntPtr read = Marshal.AllocCoTaskMem(sizeof(int));
            try
            {
                byte[] buffer = new byte[blockSize];
                using (FileStream target = File.Open(path, FileMode.Create, FileAccess.Write))
                {
                    while (totalBlocks-- > 0)
                    {
                        source.Read(buffer, blockSize, read);
                        int got = Marshal.ReadInt32(read);
                        if (got <= 0) { break; }
                        target.Write(buffer, 0, got);
                    }
                    target.Flush();
                }
            }
            finally { Marshal.FreeCoTaskMem(read); }
        }
    }
}
'@
    }

    $image = New-Object -ComObject 'IMAPI2FS.MsftFileSystemImage'
    try {
        # 1 = ISO 9660, 2 = Joliet, 4 = UDF. All three, for the same long-name reason as -u1.
        $image.FileSystemsToCreate = 7
        $image.VolumeName = $Label
        $image.Root.AddTree($SourceDir, $false)
        $result = $image.CreateResultImage()
        [OutlookAI.Testbed.IsoStreamWriter]::Write($IsoPath, $result.ImageStream, $result.BlockSize, $result.TotalBlocks)
    }
    finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($image)
    }
}

$isoPath = Join-Path $outFull ($VMName + '-unattend.iso')
if (Test-Path -LiteralPath $isoPath) { Remove-Item -LiteralPath $isoPath -Force }

$oscdimg = $null
if ($IsoBuilder -ne 'Imapi') { $oscdimg = Find-Oscdimg -Explicit $OscdimgPath }

$builderUsed = $null
if ($IsoBuilder -eq 'Oscdimg') {
    if (-not $oscdimg) {
        throw @"
-IsoBuilder Oscdimg was asked for and oscdimg.exe was not found.
It is part of the Deployment Tools feature of the Windows Assessment and Deployment Kit (ADK),
normally at:
    <ProgramFiles(x86)>\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe
Install the ADK's Deployment Tools, pass -OscdimgPath, or drop the -IsoBuilder switch to use the
IMAPI2FS builder that ships with Windows.
"@
    }
    New-IsoWithOscdimg -Exe $oscdimg -SourceDir $staging -IsoPath $isoPath -Label $VolumeLabel
    $builderUsed = "oscdimg ($oscdimg)"
}
elseif ($IsoBuilder -eq 'Imapi') {
    New-IsoWithImapi -SourceDir $staging -IsoPath $isoPath -Label $VolumeLabel
    $builderUsed = 'IMAPI2FS (MsftFileSystemImage)'
}
else {
    if ($oscdimg) {
        New-IsoWithOscdimg -Exe $oscdimg -SourceDir $staging -IsoPath $isoPath -Label $VolumeLabel
        $builderUsed = "oscdimg ($oscdimg)"
    }
    else {
        try {
            New-IsoWithImapi -SourceDir $staging -IsoPath $isoPath -Label $VolumeLabel
            $builderUsed = 'IMAPI2FS (MsftFileSystemImage) - oscdimg.exe was not found'
        }
        catch {
            throw @"
No usable ISO builder on this machine, so no answer volume was written.

  oscdimg.exe : NOT FOUND. It is the Deployment Tools feature of the Windows Assessment and
                Deployment Kit (ADK), normally at
                <ProgramFiles(x86)>\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe
                Pass -OscdimgPath if it is somewhere else.
  IMAPI2FS    : FAILED. $($_.Exception.Message)

Install the ADK Deployment Tools and re-run. Nothing partial has been left behind.
"@
        }
    }
}

# ---------------------------------------------------------------------------------------------
# Verify the artefact rather than assume it. A broken answer volume presents on the VM console
# as an ordinary interactive Setup, with nothing to say why.
# ---------------------------------------------------------------------------------------------
$bad = $null
if (-not (Test-Path -LiteralPath $isoPath)) {
    $bad = 'the builder reported success and produced no file'
}
else {
    $iso = Get-Item -LiteralPath $isoPath
    if ($iso.Length -lt 32KB) {
        $bad = "the image is only $($iso.Length) bytes"
    }
    else {
        # The ISO 9660 primary volume descriptor sits at sector 16 and starts with a type byte
        # then 'CD001'.
        $stream = [System.IO.File]::OpenRead($isoPath)
        try {
            $buffer = New-Object byte[] 5
            [void]$stream.Seek(0x8001, [System.IO.SeekOrigin]::Begin)
            [void]$stream.Read($buffer, 0, 5)
            if ([System.Text.Encoding]::ASCII.GetString($buffer) -ne 'CD001') {
                $bad = 'the image has no ISO 9660 signature at sector 16'
            }
        }
        finally { $stream.Dispose() }
    }
}
if ($bad) {
    if (Test-Path -LiteralPath $isoPath) { Remove-Item -LiteralPath $isoPath -Force }
    throw "The answer volume was not written correctly - $bad. Builder: $builderUsed. Nothing usable has been left behind."
}

if (-not $KeepStaging) {
    Remove-Item -LiteralPath $staging -Recurse -Force
}

# ---------------------------------------------------------------------------------------------
# Report. The user name is worth printing - "wrong account" is otherwise indistinguishable from
# "wrong password" - and the password is never written to a stream, a log or the console.
# ---------------------------------------------------------------------------------------------
Write-Host ''
Write-Host "Answer volume for '$VMName':"
Write-Host ("  iso            {0}" -f $isoPath)
Write-Host ("  size           {0:N0} bytes" -f (Get-Item -LiteralPath $isoPath).Length)
Write-Host ("  label          {0}" -f $VolumeLabel)
Write-Host ("  built with     {0}" -f $builderUsed)
Write-Host ("  account        {0} (local administrator, autologon)" -f $guestUser)
Write-Host ("  computer name  {0}" -f $ComputerName)
Write-Host ("  edition        {0}" -f $ImageName)
Write-Host ("  template       {0}" -f $TemplatePath)
if ($KeepStaging) { Write-Host ("  staging kept   {0}" -f $staging) }
Write-Host ''
Write-Host 'THIS ISO CONTAINS THE GUEST PASSWORD IN CLEAR TEXT. It is in gitignored scratch;'
Write-Host 'do not copy it anywhere tracked, and delete it once the guest is built.'
Write-Host ''
Write-Host 'Next:'
Write-Host ("  pwsh -File Testbed/host/New-TestbedVm.ps1 -Name {0} -IsoPath <windows.iso> -AnswerIsoPath {1} -Execute -Start" -f $VMName, $isoPath)
