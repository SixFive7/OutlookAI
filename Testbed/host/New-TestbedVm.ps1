#Requires -Version 5.1
<#
.SYNOPSIS
    Creates the Hyper-V guest the live tier runs on, and records the spec it chose.

.DESCRIPTION
    THE SPECIFICATION BELOW IS CHOSEN BY THIS SCRIPT. IT IS NOT A RECORD OF THE ORIGINAL VM.

    That distinction is the whole point of this header. The generation, firmware, vCPU count,
    RAM, disk size and checkpoint type of the machine every measurement in this repository was
    taken on were never written down, and there is no way to recover them from the repository.
    A script that quietly defaulted to plausible values would read, a year from now, as a record
    of what was there - which is worse than an admitted gap, because it would be believed.

    So: these are defaults picked to work, with a reason each. Override any of them. Whatever you
    end up with, the script writes to -SpecOut, and THAT file is the record. Keep it somewhere
    that is not this machine.

    WHY EACH DEFAULT

    * Generation 2. UEFI, no legacy emulation, and it is what a current Windows 11 ISO expects.
      Generation 1 would need a different disk and boot arrangement throughout.
    * Secure Boot on, vTPM on. Windows 11 setup refuses without them unless you fight it, and
      fighting it is a source of differences from an ordinary desktop that nobody wants in a
      test machine.
    * 4 vCPU / 8 GB. The corpus build is COM round trips - one item at a time, latency-bound,
      not parallel - so more cores buy nothing. 8 GB is Outlook plus the indexer plus headroom.
    * 128 GB dynamic disk. A 20,000-item corpus is ~225 MB of body text and the PST lands near
      400 MB; Windows plus Office plus a couple of checkpoints is what actually fills this.
    * STANDARD checkpoints, deliberately. Production checkpoints use VSS, which quiesces the
      guest - with Outlook mid-run and a PST open that is a different and less predictable
      thing to restore than a plain saved state.
    * NO network switch by default. The mail sink binds to loopback and nothing here needs the
      internet after the toolchain is in. A test VM running a mail server with a route to the
      outside is an open relay waiting to happen. Pass -SwitchName for the install phase, then
      disconnect.

    THIS SCRIPT HAS NOT BEEN RUN. It was written from Hyper-V's documented cmdlet surface by an
    agent that was forbidden to touch Hyper-V. Read it before you run it, and run it with
    -Execute only once you have.

.PARAMETER Name
    VM name. `OutlookAI-TestVM` by convention - Docs/live-tier-on-the-vm.md and Testbed/testbed.json
    both use it.

.PARAMETER IsoPath
    Windows 11 installation ISO. You supply this; see Testbed/README.md section 6.

.PARAMETER SpecOut
    Where to write the record of what was created. Defaults beside the VHD.

.PARAMETER Execute
    Without it, the script prints the plan and creates nothing.

.EXAMPLE
    pwsh -File Testbed/host/New-TestbedVm.ps1 -IsoPath D:\iso\Win11.iso
    pwsh -File Testbed/host/New-TestbedVm.ps1 -IsoPath D:\iso\Win11.iso -SwitchName Default -Execute
#>
[CmdletBinding()]
param(
    [string] $Name = 'OutlookAI-TestVM',
    [Parameter(Mandatory = $true)] [string] $IsoPath,
    [string] $VhdPath,
    [int]    $Generation = 2,
    [int]    $ProcessorCount = 4,
    [int64]  $MemoryStartupBytes = 8GB,
    [int64]  $VhdSizeBytes = 128GB,
    [ValidateSet('Standard', 'Production')] [string] $CheckpointType = 'Standard',
    [string] $SwitchName,
    [string] $SpecOut,
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'

if (-not $VhdPath) {
    $root = (Get-VMHost).VirtualHardDiskPath
    if (-not $root) { $root = 'C:\Hyper-V\Virtual Hard Disks' }
    $VhdPath = Join-Path $root ("$Name.vhdx")
}
if (-not $SpecOut) { $SpecOut = [IO.Path]::ChangeExtension($VhdPath, '.spec.json') }

$spec = [ordered]@{
    createdUtc         = (Get-Date).ToUniversalTime().ToString('o')
    createdBy          = 'Testbed/host/New-TestbedVm.ps1'
    warning            = 'These values were CHOSEN when this VM was built. They are not a record of any earlier VM.'
    name               = $Name
    generation         = $Generation
    secureBoot         = ($Generation -eq 2)
    vtpm               = ($Generation -eq 2)
    processorCount     = $ProcessorCount
    memoryStartupBytes = $MemoryStartupBytes
    vhdPath            = $VhdPath
    vhdSizeBytes       = $VhdSizeBytes
    checkpointType     = $CheckpointType
    switchName         = $SwitchName
    isoPath            = $IsoPath
    hostOs             = [Environment]::OSVersion.VersionString
}

Write-Host "Plan for '$Name':"
$spec.GetEnumerator() | ForEach-Object { Write-Host ("  {0,-19}{1}" -f $_.Key, $_.Value) }
Write-Host ''

if (-not $Execute) {
    Write-Host 'Dry run. Nothing created. Re-run with -Execute.'
    Write-Host "Spec would be written to $SpecOut"
    return
}

if (-not (Test-Path -LiteralPath $IsoPath)) { throw "ISO not found: $IsoPath" }
if (Get-VM -Name $Name -ErrorAction SilentlyContinue) { throw "A VM named '$Name' already exists. Refusing to touch it." }
if (Test-Path -LiteralPath $VhdPath) { throw "A VHD already exists at '$VhdPath'. Refusing to overwrite it." }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $VhdPath) | Out-Null

$newVmArgs = @{
    Name               = $Name
    Generation         = $Generation
    MemoryStartupBytes = $MemoryStartupBytes
    NewVHDPath         = $VhdPath
    NewVHDSizeBytes    = $VhdSizeBytes
}
if ($SwitchName) { $newVmArgs['SwitchName'] = $SwitchName }
New-VM @newVmArgs | Out-Null

Set-VM -Name $Name -ProcessorCount $ProcessorCount -CheckpointType $CheckpointType `
    -AutomaticCheckpointsEnabled $false -AutomaticStopAction ShutDown

if ($Generation -eq 2) {
    # Order matters: the key protector has to exist before the vTPM can be enabled.
    Set-VMKeyProtector -VMName $Name -NewLocalKeyProtector
    Enable-VMTPM -VMName $Name
    Set-VMFirmware -VMName $Name -EnableSecureBoot On -SecureBootTemplate 'MicrosoftWindows'
}

Add-VMDvdDrive -VMName $Name -Path $IsoPath
$dvd = Get-VMDvdDrive -VMName $Name
Set-VMFirmware -VMName $Name -FirstBootDevice $dvd

if (-not $SwitchName) {
    Get-VMNetworkAdapter -VMName $Name | Disconnect-VMNetworkAdapter
    Write-Host 'Network adapter left DISCONNECTED. Connect it for the install phase only.'
}

$spec | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $SpecOut -Encoding UTF8

Write-Host ''
Write-Host "Created '$Name'. Spec recorded at $SpecOut - keep a copy off this machine."
Write-Host 'Next, by hand: install Windows, then set autologon, no lock screen, no sleep.'
Write-Host 'Then take checkpoint CP-01-WIN-CLEAN. See Docs/live-tier-on-the-vm.md section 2.1.'
