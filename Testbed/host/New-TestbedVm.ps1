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

    THE INSTALL IS UNATTENDED. Pass -AnswerIsoPath and the guest installs Windows, creates its
    local administrator, sets the locale to match the maintainer's machine and logs itself in,
    with nobody watching. Build that ISO first:

        pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName <name>

    Two DVD drives are attached: the Windows ISO, which is the first boot device, and the answer
    volume. Windows Setup finds autounattend.xml at the root of any removable volume, which is
    exactly what the second drive is. Nothing here changes the firmware requirements: Secure Boot
    and the vTPM stay ON, because a Generation 2 VM meets Windows 11's requirements natively and
    a guest built by disabling those checks is not the machine the userbase runs.

    WITHOUT -AnswerIsoPath the VM is still created, but the install is an ordinary interactive
    one and the script says so rather than letting you discover it at the language page.

    THE ONE KEYSTROKE NOBODY CAN AVOID. Microsoft's retail ISO boots through a loader that
    prints "Press any key to boot from CD or DVD" and gives up after a few seconds. That prompt
    comes from the boot image inside the ISO, not from Hyper-V, so no VM setting turns it off.
    -Start therefore types a key at the guest's synthetic keyboard through the Msvm_Keyboard WMI
    class for the first few seconds - no window, no focus change, nothing on the host's screen.
    If your host cannot reach that class, start the VM yourself and press a key in the console
    once; everything after that is unattended either way. The deterministic alternative, if this
    ever needs to be hands-off on a machine where the WMI route is blocked, is to rebuild the
    Windows ISO with oscdimg using efisys_noprompt.bin as the EFI boot image - correct, but it
    rewrites 8 GB to save one keystroke.

    THIS SCRIPT HAS NOT BEEN RUN. It was written from Hyper-V's documented cmdlet surface by an
    agent that was forbidden to touch Hyper-V. Read it before you run it, and run it with
    -Execute only once you have.

.PARAMETER Name
    VM name. `OutlookAI-TestVM` by convention - Docs/live-tier-on-the-vm.md and Testbed/testbed.json
    both use it.

.PARAMETER IsoPath
    Windows 11 installation ISO. You supply this; see Testbed/README.md section 6.

.PARAMETER AnswerIsoPath
    The unattended-install answer volume from Testbed/host/New-AnswerFile.ps1. Attached as a
    second DVD drive. Omit it and the install is interactive.

.PARAMETER SpecOut
    Where to write the record of what was created. Defaults beside the VHD.

.PARAMETER Execute
    Without it, the script prints the plan and creates nothing.

.PARAMETER Start
    Start the VM once it is created, and type at its keyboard for the first few seconds so the
    ISO's "press any key" prompt does not time out. Ignored without -Execute.

.PARAMETER NoBootKeystroke
    Start the VM but do not type anything. Use it if the Msvm_Keyboard route is blocked on this
    host and you would rather press the key in the console yourself.

.EXAMPLE
    pwsh -File Testbed/host/New-TestbedVm.ps1 -IsoPath D:\iso\Win11.iso

.EXAMPLE
    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName OutlookAI-Indexed
    pwsh -File Testbed/host/New-TestbedVm.ps1 -Name OutlookAI-Indexed -IsoPath .work\media\Win11_25H2_EnglishInternational_x64_v2.iso -AnswerIsoPath .work\testbed-answer\OutlookAI-Indexed\OutlookAI-Indexed-unattend.iso -Execute -Start
#>
[CmdletBinding()]
param(
    [string] $Name = 'OutlookAI-TestVM',
    [Parameter(Mandatory = $true)] [string] $IsoPath,
    [string] $AnswerIsoPath,
    [string] $VhdPath,
    [int]    $Generation = 2,
    [int]    $ProcessorCount = 4,
    [int64]  $MemoryStartupBytes = 8GB,
    [int64]  $VhdSizeBytes = 128GB,
    [ValidateSet('Standard', 'Production')] [string] $CheckpointType = 'Standard',
    [string] $SwitchName,
    [string] $SpecOut,
    [switch] $Execute,
    [switch] $Start,
    [switch] $NoBootKeystroke,
    [int]    $BootKeystrokeSeconds = 20
)

$ErrorActionPreference = 'Stop'

if ($Name -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    throw "VM name '$Name' is not a plain name; refusing to build WMI filters and file paths from it."
}

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
    answerIsoPath      = $AnswerIsoPath
    unattended         = [bool] $AnswerIsoPath
    startedByScript    = [bool] $Start
    hostOs             = [Environment]::OSVersion.VersionString
}

Write-Host "Plan for '$Name':"
$spec.GetEnumerator() | ForEach-Object { Write-Host ("  {0,-19}{1}" -f $_.Key, $_.Value) }
Write-Host ''

if (-not $AnswerIsoPath) {
    Write-Warning @'
No -AnswerIsoPath, so this install will be INTERACTIVE: somebody has to sit through the language
page, the edition picker, the disk layout, OOBE and the Microsoft-account nag, and then set the
locale by hand to match the host. Build the answer volume first and the whole thing is hands-off:
    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName <name>
'@
}

if (-not $Execute) {
    Write-Host 'Dry run. Nothing created. Re-run with -Execute.'
    Write-Host "Spec would be written to $SpecOut"
    return
}

if (-not (Test-Path -LiteralPath $IsoPath)) { throw "ISO not found: $IsoPath" }
if ($AnswerIsoPath -and -not (Test-Path -LiteralPath $AnswerIsoPath)) {
    throw "Answer volume not found: $AnswerIsoPath. Build it with Testbed/host/New-AnswerFile.ps1."
}
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

# The Windows ISO first, so it is DVD drive zero, then the answer volume beside it. Setup scans
# every removable volume for autounattend.xml at its root, so the order of the two only matters
# for which one boots - and that is pinned explicitly below rather than left to the enumeration.
Add-VMDvdDrive -VMName $Name -Path $IsoPath
$windowsDvd = @(Get-VMDvdDrive -VMName $Name | Where-Object { $_.Path -eq $IsoPath })
if ($windowsDvd.Count -ne 1) {
    throw "Expected exactly one DVD drive holding '$IsoPath'; found $($windowsDvd.Count)."
}

if ($AnswerIsoPath) {
    Add-VMDvdDrive -VMName $Name -Path $AnswerIsoPath
    Write-Host "Answer volume attached as a second DVD drive: $AnswerIsoPath"
}

# FirstBootDevice takes ONE device. Resolving it after both drives exist, by matching the path,
# is what stops this handing Set-VMFirmware an array the day a second drive was added.
Set-VMFirmware -VMName $Name -FirstBootDevice $windowsDvd[0]

if (-not $SwitchName) {
    Get-VMNetworkAdapter -VMName $Name | Disconnect-VMNetworkAdapter
    Write-Host 'Network adapter left DISCONNECTED. Connect it for the install phase only.'
}

$spec | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $SpecOut -Encoding UTF8

Write-Host ''
Write-Host "Created '$Name'. Spec recorded at $SpecOut - keep a copy off this machine."

if ($Start) {
    Start-VM -Name $Name
    Write-Host "Started '$Name'."

    if ($NoBootKeystroke) {
        Write-Host 'Not typing anything. Press a key in the VM console within a few seconds or the DVD will not boot.'
    }
    else {
        # "Press any key to boot from CD or DVD" comes from the boot image inside Microsoft's
        # ISO, so there is no VM setting that removes it. Msvm_Keyboard.TypeKey posts to the
        # guest's synthetic keyboard over WMI: no console window, no focus change, nothing drawn
        # on the host. 0x20 is VK_SPACE. It is sent repeatedly because the prompt appears a few
        # seconds into firmware POST and lasts only a few seconds after that, and a single
        # well-timed keystroke is a race this script would lose.
        $typed = 0
        $lastError = $null
        try {
            # The Hyper-V host is itself an Msvm_ComputerSystem, so this is filtered by name and
            # then counted rather than trusted: typing at the wrong one is not a thing to leave
            # to an assumption.
            $guest = @(Get-CimInstance -Namespace 'root\virtualization\v2' -ClassName 'Msvm_ComputerSystem' `
                    -Filter "ElementName='$Name'" -ErrorAction Stop)
            if ($guest.Count -ne 1) { throw "Expected one Msvm_ComputerSystem named '$Name'; found $($guest.Count)." }
            $keyboard = @($guest[0] | Get-CimAssociatedInstance -ResultClassName 'Msvm_Keyboard' -ErrorAction Stop)
            if ($keyboard.Count -lt 1) { throw "'$Name' has no Msvm_Keyboard to type at." }
            $keyboard = $keyboard[0]
            $deadline = (Get-Date).AddSeconds($BootKeystrokeSeconds)
            while ((Get-Date) -lt $deadline) {
                try {
                    Invoke-CimMethod -InputObject $keyboard -MethodName 'TypeKey' `
                        -Arguments @{ keyCode = [uint32] 0x20 } -ErrorAction Stop | Out-Null
                    $typed++
                }
                catch { $lastError = $_ }
                Start-Sleep -Milliseconds 750
            }
        }
        catch { $lastError = $_ }

        if ($typed -gt 0) {
            Write-Host ("Typed {0} keystroke(s) at the guest over {1}s to get past the boot prompt." -f $typed, $BootKeystrokeSeconds)
        }
        else {
            $why = 'no reason given'
            if ($lastError) { $why = $lastError.Exception.Message }
            Write-Warning ("Could not type at the guest's keyboard ($why). If the VM is sitting at a boot " +
                'failure, open its console and press a key; the rest of the install is unattended regardless.')
        }
    }
}

Write-Host ''
if ($AnswerIsoPath) {
    Write-Host 'The install is unattended from here: Windows, the local administrator, autologon, the'
    Write-Host 'locale and the power policy all come from the answer volume. Watch the console if you'
    Write-Host 'like; nothing needs you. When it settles at the desktop, read the first-logon log at'
    Write-Host 'C:\Windows\Setup\first-logon.log on the guest and diff it against the host table in'
    Write-Host 'Testbed/MEDIA.md.'
}
else {
    Write-Host 'Next, by hand: install Windows, then set autologon, no lock screen, no sleep.'
}
Write-Host 'Then take checkpoint CP-01-WIN-CLEAN. See Docs/live-tier-on-the-vm.md section 2.1.'
