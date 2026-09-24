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

    THE DISCS COME OUT BEFORE THE FIRST CHECKPOINT - -CompleteInstall, added 2026-09-24. The
    answer volume carries the guest password in clear text, and a checkpoint taken while a disc
    is attached keeps a reference to that disc's file: delete the file afterwards and a restore
    of that checkpoint has nothing to attach. Both guests built on 2026-09-15 are in exactly that
    state - their CP-01-WIN-CLEAN was taken with the answer disc in (read off both guests
    2026-09-24) - so their answer ISOs have to stay until those guests are replaced. So a
    from-scratch build now ends with:

        pwsh -File Testbed/host/New-TestbedVm.ps1 -Name <name> -CompleteInstall -Execute

    which, in this order and refusing at the first thing that is not so:
      1. waits for C:\Windows\Setup\first-logon.log on the guest to end in its DONE line. The
         answer volume is needed until then, because Complete-FirstLogon.ps1 runs FROM it at
         first logon; a run that reported failures, or never found the script, keeps both discs
         for the re-run and stops here;
      2. ejects every disc - the Windows ISO and the answer volume - and reads the drives back
         empty;
      3. takes CP-01-WIN-CLEAN, and reads it back holding no disc;
      4. deletes the answer ISO - only a file named <name>-unattend.iso, the shape
         New-AnswerFile.ps1 writes, and only once no VM and no checkpoint on this host still
         references it - with the build log beside it and any staging directory that
         New-AnswerFile.ps1 -KeepStaging left, which holds the same password.
    It refuses outright on a guest whose CP-01-WIN-CLEAN already holds a disc, changing nothing,
    and it can be re-run: once CP-01-WIN-CLEAN exists without a disc it goes straight to step 4,
    which is how an ISO that something else still referenced gets deleted later. The Windows ISO
    is ejected and never deleted: it is staged media, a precondition, not an artefact. Its
    decisions - the first-logon log's verdict, which file is the answer ISO, what references it -
    are pure functions, and -SelfTest runs them anywhere, reading no VM.

    WHAT HAS RUN AND WHAT HAS NOT. The create path has: it built OutlookAI-Unindexed on
    2026-09-15, and the first time it ran it found the relative-path bug commented below.
    -CompleteInstall HAS NOT RUN. It was written from Hyper-V's documented cmdlet surface by an
    agent that was forbidden to touch Hyper-V, and only its -SelfTest has been executed. Read it
    before you run it, and run it with -Execute only once you have.

.PARAMETER Name
    MANDATORY. The VM name to create. There is no default: THREE MACHINES COEXIST during the
    changeover - OutlookAI-Indexed, OutlookAI-Unindexed and the outgoing OutlookAI-TestVM - and a
    default that silently picks one of three is the exact shape of mistake this testbed keeps
    making. This script also derives the VHD path and the spec file from the name, so a wrong
    default is a new disk in somebody else's directory, or a refusal on top of a VM that already
    exists.

    `OutlookAI-TestVM` is the OLD guest, the one being replaced; Docs/live-tier-on-the-vm.md and
    Testbed/testbed.json still name it because they describe the machine the published
    measurements were taken on. It is not a name to build under.

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

.PARAMETER CompleteInstall
    Finish a guest this script created: wait for its first logon, eject both discs, take
    CP-01-WIN-CLEAN without them, and delete the answer ISO. See THE DISCS COME OUT BEFORE THE
    FIRST CHECKPOINT above. Without -Execute it reads and reports, and changes nothing. Exit 0
    when all of it is done; exit 3 when everything but the deletion is, because something on this
    host still references the ISO - it names what, and a re-run finishes the job once that is gone.

.PARAMETER InstallTimeoutMinutes
    How long -CompleteInstall waits for the first-logon log's DONE line. The unattended build of
    2026-09-15 reached its desktop about seven minutes after -Start.

.PARAMETER RepoRoot
    Repository root, for the guest credential -CompleteInstall reads the first-logon log with.
    Defaults to two levels above this script; point it at the main checkout when running from a
    worktree, because the credential is gitignored and lives only there.

.PARAMETER SelfTest
    Run -CompleteInstall's decisions against synthetic inputs and exit. Reads no VM, no guest and
    no credential - two source files beside this one, and nothing else.

.EXAMPLE
    pwsh -File Testbed/host/New-TestbedVm.ps1 -Name OutlookAI-Unindexed -IsoPath D:\iso\Win11.iso

.EXAMPLE
    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName OutlookAI-Indexed
    pwsh -File Testbed/host/New-TestbedVm.ps1 -Name OutlookAI-Indexed -IsoPath .work\media\Win11_25H2_EnglishInternational_x64_v2.iso -AnswerIsoPath .work\testbed-answer\OutlookAI-Indexed\OutlookAI-Indexed-unattend.iso -Execute -Start
    pwsh -File Testbed/host/New-TestbedVm.ps1 -Name OutlookAI-Indexed -CompleteInstall -Execute

.EXAMPLE
    pwsh -File Testbed/host/New-TestbedVm.ps1 -SelfTest
#>
[CmdletBinding(DefaultParameterSetName = 'Create')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Create')]
    [Parameter(Mandatory = $true, ParameterSetName = 'CompleteInstall')]
    [string] $Name,
    [Parameter(Mandatory = $true, ParameterSetName = 'Create')] [string] $IsoPath,
    [Parameter(ParameterSetName = 'Create')]
    [Parameter(ParameterSetName = 'CompleteInstall')]
    [string] $AnswerIsoPath,
    [Parameter(ParameterSetName = 'Create')] [string] $VhdPath,
    [Parameter(ParameterSetName = 'Create')] [int]    $Generation = 2,
    [Parameter(ParameterSetName = 'Create')] [int]    $ProcessorCount = 4,
    [Parameter(ParameterSetName = 'Create')] [int64]  $MemoryStartupBytes = 8GB,
    [Parameter(ParameterSetName = 'Create')] [int64]  $VhdSizeBytes = 128GB,
    [Parameter(ParameterSetName = 'Create')] [ValidateSet('Standard', 'Production')] [string] $CheckpointType = 'Standard',
    [Parameter(ParameterSetName = 'Create')] [string] $SwitchName,
    [Parameter(ParameterSetName = 'Create')] [string] $SpecOut,
    [Parameter(ParameterSetName = 'Create')]
    [Parameter(ParameterSetName = 'CompleteInstall')]
    [switch] $Execute,
    [Parameter(ParameterSetName = 'Create')] [switch] $Start,
    [Parameter(ParameterSetName = 'Create')] [switch] $NoBootKeystroke,
    [Parameter(ParameterSetName = 'Create')] [int]    $BootKeystrokeSeconds = 20,
    [Parameter(Mandatory = $true, ParameterSetName = 'CompleteInstall')] [switch] $CompleteInstall,
    [Parameter(ParameterSetName = 'CompleteInstall')] [int] $InstallTimeoutMinutes = 60,
    [Parameter(ParameterSetName = 'CompleteInstall')] [string] $RepoRoot,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# -CompleteInstall's names. The first checkpoint is the one Docs/live-tier-on-the-vm.md section
# 2.11 and Testbed/testbed.json list first; the log is Complete-FirstLogon.ps1's default -LogPath,
# which -SelfTest holds equal to that script's own default.
# ---------------------------------------------------------------------------------------------
$FirstCheckpointName = 'CP-01-WIN-CLEAN'
$FirstLogonLogPath = 'C:\Windows\Setup\first-logon.log'

# =============================================================================================
# PURE DECISIONS for -CompleteInstall. No Hyper-V, no guest, no credential, no output.
# =============================================================================================

<#
    What C:\Windows\Setup\first-logon.log says about the first logon, judged on the LAST run in
    it: Complete-FirstLogon.ps1 appends, so a re-run by hand after a failure is a second header
    and a second DONE line further down, and only the last one counts.

      Absent    no log. Setup has not reached first logon yet - or Complete-FirstLogon.ps1 was
                found and REFUSED: its guest guard runs before it opens the log, so a refusal
                writes nothing (or, less likely, it died before its first line). Wait; at the
                deadline, say so.
      Running   a log with no DONE line after its last header (or no header yet - the first line
                it writes is a separator). Wait.
      Done      the last run ended "DONE. All N step(s) succeeded." The answer volume is spent.
      Failed    the last run ended "DONE WITH N FAILURE(S)". Keep the discs: the answer volume
                holds the script to re-run.
      NotFound  the answer file's own fallback line - the script was on no drive at all.

    The strings are Complete-FirstLogon.ps1's and autounattend.template.xml's own; -SelfTest
    reads both files and fails if either stops writing them.
#>
function Get-FirstLogonVerdict {
    param([string[]] $Lines)

    $all = @($Lines | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) { return [pscustomobject]@{ State = 'Absent'; Failed = @() } }

    $header = -1
    for ($i = $all.Count - 1; $i -ge 0; $i--) {
        if ($all[$i] -like '*Complete-FirstLogon.ps1 on * as *') { $header = $i; break }
    }
    if ($header -lt 0) {
        if (@($all | Where-Object { $_ -like '*Complete-FirstLogon.ps1 was not found on any drive*' }).Count -gt 0) {
            return [pscustomobject]@{ State = 'NotFound'; Failed = @() }
        }
        return [pscustomobject]@{ State = 'Running'; Failed = @() }
    }

    $run = @($all[$header..($all.Count - 1)])
    $failed = @($run | Where-Object { $_ -match '^\S.*\s{2}FAILED\s{2}' })
    for ($i = $run.Count - 1; $i -ge 0; $i--) {
        if ($run[$i] -match 'DONE\. All \d+ step\(s\) succeeded') { return [pscustomobject]@{ State = 'Done'; Failed = $failed } }
        if ($run[$i] -match 'DONE WITH \d+ FAILURE\(S\)') { return [pscustomobject]@{ State = 'Failed'; Failed = $failed } }
    }
    return [pscustomobject]@{ State = 'Running'; Failed = $failed }
}

# Whether a path names THIS guest's answer ISO: the one file shape New-AnswerFile.ps1 writes,
# '<VMName>-unattend.iso'. Nothing else is ever deleted - not the Windows ISO, not another guest's.
function Test-IsAnswerIsoFor {
    param([string] $Path, [string] $VMName)

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($VMName)) { return $false }
    return [string]::Equals([System.IO.Path]::GetFileName($Path.Trim()), "$VMName-unattend.iso", [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-NormalisedPath {
    param([string] $Path)

    try { return [System.IO.Path]::GetFullPath($Path.Trim()).TrimEnd('\') }
    catch { return $Path.Trim().TrimEnd('\') }
}

<#
    Which of the given disc references - @{ Where; Path }, one per DVD drive of a VM or of a
    checkpoint that holds a disc - point at $Path. Compared on the full path, ignoring case,
    because Hyper-V stores what it was handed resolved and Windows paths are case-insensitive.
#>
function Find-DiscReference {
    param([string] $Path, [object[]] $References)

    # Emitted one by one, NOT returned as ', $array': a caller's @() around a comma-returned
    # array counts ONE element whatever it holds (measured in both shells, 2026-09-24).
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $wanted = Get-NormalisedPath $Path
    foreach ($reference in @($References)) {
        if ($null -eq $reference -or [string]::IsNullOrWhiteSpace([string]$reference.Path)) { continue }
        if ([string]::Equals((Get-NormalisedPath ([string]$reference.Path)), $wanted, [System.StringComparison]::OrdinalIgnoreCase)) {
            $reference
        }
    }
}

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    function Test-Case {
        param([string] $What, $Expected, $Actual)

        $script:SelfTestChecks++
        $expectedText = "$Expected"
        if ($null -eq $Expected) { $expectedText = '<null>' }
        $actualText = "$Actual"
        if ($null -eq $Actual) { $actualText = '<null>' }
        if ($expectedText -ceq $actualText) {
            Write-Host ("  OK   {0}" -f $What)
        }
        else {
            $script:SelfTestFailures += "$What : expected $expectedText, got $actualText"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $expectedText, $actualText)
        }
    }

    Write-Host 'New-TestbedVm -CompleteInstall self-test. No VM, no guest, no credential.'
    Write-Host ''
    Write-Host '== the first-logon log verdict =='

    $header = '2026-09-15 19:50:06  Complete-FirstLogon.ps1 on OAI-UNINDEXED as vmadmin'
    $rule = '2026-09-15 19:50:06  ================================================================'
    $ok = '2026-09-15 19:50:09    OK      home location GeoId 176'
    $bad = '2026-09-15 19:50:09    FAILED  system locale (non-Unicode) = en-US :: Access is denied.'
    $done = '2026-09-15 19:50:10  DONE. All 14 step(s) succeeded. Log: C:\Windows\Setup\first-logon.log'
    $doneBad = '2026-09-15 19:50:10  DONE WITH 1 FAILURE(S) out of 14 step(s). The guest is NOT fully configured; read the FAILED lines above.'
    $notFound = 'Complete-FirstLogon.ps1 was not found on any drive. The guest locale and power settings are NOT the ones the testbed expects.'

    Test-Case 'no log at all is Absent' 'Absent' (Get-FirstLogonVerdict -Lines $null).State
    Test-Case 'an empty log is Absent' 'Absent' (Get-FirstLogonVerdict -Lines @()).State
    Test-Case 'only the opening separator is Running, not a verdict' 'Running' (Get-FirstLogonVerdict -Lines @($rule)).State
    Test-Case 'a header with no DONE line is Running' 'Running' (Get-FirstLogonVerdict -Lines @($rule, $header, $rule, $ok)).State
    Test-Case 'the 2026-09-15 shape, ending DONE, is Done' 'Done' (Get-FirstLogonVerdict -Lines @($rule, $header, $rule, $ok, '', $done)).State
    $verdict = Get-FirstLogonVerdict -Lines @($rule, $header, $rule, $ok, $bad, '', $doneBad)
    Test-Case 'a run that reported failures is Failed' 'Failed' $verdict.State
    Test-Case 'and carries its FAILED line' 1 @($verdict.Failed).Count
    Test-Case 'a re-run that succeeded after a failed one is Done - the last run counts' 'Done' (Get-FirstLogonVerdict -Lines @($rule, $header, $rule, $bad, $doneBad, $rule, $header, $rule, $ok, $done)).State
    Test-Case 'a re-run still going after a finished one is Running, not the old verdict' 'Running' (Get-FirstLogonVerdict -Lines @($rule, $header, $rule, $ok, $done, $rule, $header, $rule, $ok)).State
    Test-Case "the answer file's not-found line is NotFound" 'NotFound' (Get-FirstLogonVerdict -Lines @($notFound)).State
    Test-Case 'a by-hand run after a not-found is judged on that run' 'Done' (Get-FirstLogonVerdict -Lines @($notFound, $rule, $header, $rule, $ok, $done)).State

    Write-Host ''
    Write-Host '== those strings are still the ones the two files write =='

    $guestDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'guest'
    $firstLogon = Join-Path $guestDir 'Complete-FirstLogon.ps1'
    $template = Join-Path $guestDir 'autounattend.template.xml'
    $firstLogonText = ''
    $templateText = ''
    if (Test-Path -LiteralPath $firstLogon) { $firstLogonText = [System.IO.File]::ReadAllText($firstLogon) }
    if (Test-Path -LiteralPath $template) { $templateText = [System.IO.File]::ReadAllText($template) }
    Test-Case 'Complete-FirstLogon.ps1 still writes its header' $true $firstLogonText.Contains('"Complete-FirstLogon.ps1 on {0} as {1}"')
    Test-Case 'and its success line' $true $firstLogonText.Contains('"DONE. All {0} step(s) succeeded. Log: {1}"')
    Test-Case 'and its failure line' $true $firstLogonText.Contains('"DONE WITH {0} FAILURE(S) out of {1} step(s).')
    Test-Case 'and its FAILED lines' $true $firstLogonText.Contains('"  FAILED  {0} :: {1}"')
    Test-Case 'and its default -LogPath is the file this reads' $true $firstLogonText.Contains("`$LogPath = '$FirstLogonLogPath'")
    Test-Case "the answer file's fallback still writes the not-found line" $true $templateText.Contains('Complete-FirstLogon.ps1 was not found on any drive.')

    Write-Host ''
    Write-Host '== which file is the answer ISO =='

    $answer = 'C:\Source\SixFive7\OutlookAI\.work\testbed-answer\OutlookAI-Indexed\OutlookAI-Indexed-unattend.iso'
    Test-Case "New-AnswerFile.ps1's own name is the answer ISO" $true (Test-IsAnswerIsoFor -Path $answer -VMName 'OutlookAI-Indexed')
    Test-Case 'in any case' $true (Test-IsAnswerIsoFor -Path $answer.ToUpperInvariant() -VMName 'outlookai-indexed')
    Test-Case 'the Windows ISO never is' $false (Test-IsAnswerIsoFor -Path 'C:\Source\SixFive7\OutlookAI\.work\media\Win11_25H2_EnglishInternational_x64_v2.iso' -VMName 'OutlookAI-Indexed')
    Test-Case "another guest's answer ISO never is" $false (Test-IsAnswerIsoFor -Path $answer.Replace('OutlookAI-Indexed-unattend', 'OutlookAI-Unindexed-unattend') -VMName 'OutlookAI-Indexed')
    Test-Case 'a name that merely ends the same way is not' $false (Test-IsAnswerIsoFor -Path 'C:\x\Old-OutlookAI-Indexed-unattend.iso' -VMName 'OutlookAI-Indexed')
    Test-Case 'an empty path is not, and does not throw' $false (Test-IsAnswerIsoFor -Path '' -VMName 'OutlookAI-Indexed')

    Write-Host ''
    Write-Host '== what still references it =='

    $references = @(
        [pscustomobject]@{ Where = "VM 'OutlookAI-Indexed', checkpoint 'CP-01-WIN-CLEAN'"; Path = $answer.ToLowerInvariant() }
        [pscustomobject]@{ Where = "VM 'OutlookAI-Indexed', its current configuration"; Path = 'C:\Source\SixFive7\OutlookAI\.work\media\Win11_25H2_EnglishInternational_x64_v2.iso' }
        [pscustomobject]@{ Where = "VM 'OutlookAI-Old', checkpoint 'CP-05'"; Path = 'C:\Source\SixFive7\OutlookAI\.work\testbed-answer\OutlookAI-Indexed\.\OutlookAI-Indexed-unattend.iso' }
        [pscustomobject]@{ Where = 'an empty drive'; Path = '' }
        $null
    )
    $hits = @(Find-DiscReference -Path $answer -References $references)
    Test-Case 'a checkpoint holding it in another case is a reference' $true (@($hits | Where-Object { $_.Where -like '*CP-01-WIN-CLEAN*' }).Count -eq 1)
    Test-Case 'so is one spelling it with a dot segment' $true (@($hits | Where-Object { $_.Where -like '*CP-05*' }).Count -eq 1)
    Test-Case 'exactly those two - not the Windows ISO, the empty drive or the null beside them' 2 $hits.Count
    Test-Case 'nothing references it when nothing holds it' 0 @(Find-DiscReference -Path $answer -References @($references[1], $references[3])).Count
    Test-Case 'no references at all is none, and does not throw' 0 @(Find-DiscReference -Path $answer -References $null).Count

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. These need Hyper-V and a guest, and none of them has run:'
    Write-Host '  * reading the first-logon log over PowerShell Direct, and the wait'
    Write-Host '  * Set-VMDvdDrive -Path $null emptying a drive, and the read-back'
    Write-Host '  * Checkpoint-VM, and Get-VMDvdDrive -VMSnapshot reading the new checkpoint back'
    Write-Host '  * gathering every VM''s and every checkpoint''s discs, and the deletion itself'

    if ($script:SelfTestFailures.Count -gt 0) {
        Write-Host ''
        foreach ($failure in $script:SelfTestFailures) { Write-Host "  $failure" }
        return 1
    }
    return 0
}

if ($PSCmdlet.ParameterSetName -eq 'SelfTest') { exit (Invoke-SelfTest) }

if ($Name -notmatch '^[A-Za-z0-9._-]{1,64}$') {
    throw "VM name '$Name' is not a plain name; refusing to build WMI filters and file paths from it."
}

# =============================================================================================
# -CompleteInstall. Everything below this block, down to the end of the file, is the create path.
# =============================================================================================
if ($PSCmdlet.ParameterSetName -eq 'CompleteInstall') {
    if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }

    $vms = @(Get-VM -Name $Name -ErrorAction SilentlyContinue)
    if ($vms.Count -eq 0) { throw "No VM named '$Name' on this host. -CompleteInstall finishes a guest that -Execute -Start created; it creates nothing." }
    if ($vms.Count -gt 1) { throw "$($vms.Count) VMs on this host are named '$Name'. Refusing to guess which one to eject, checkpoint and clean up after." }
    $vm = $vms[0]

    $mode = 'DRY RUN - nothing is ejected, checkpointed or deleted. Re-run with -Execute.'
    if ($Execute) { $mode = 'EXECUTE' }
    Write-Host "Completing the install of '$Name' (VM state: $($vm.State)). $mode"
    Write-Host ''

    # ---- what is in the drives, and is there a first checkpoint already ----------------------
    $held = @(Get-VMDvdDrive -VM $vm | Where-Object { $_.Path })
    foreach ($drive in $held) { Write-Host ("  disc in {0}:{1}   {2}" -f $drive.ControllerNumber, $drive.ControllerLocation, $drive.Path) }
    if ($held.Count -eq 0) { Write-Host '  no disc in any DVD drive' }

    $firstCheckpoint = @(Get-VMSnapshot -VM $vm -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $FirstCheckpointName })
    $checkpointed = $false
    if ($firstCheckpoint.Count -gt 0) {
        $checkpointDiscs = @($firstCheckpoint | ForEach-Object { Get-VMDvdDrive -VMSnapshot $_ } | Where-Object { $_.Path })
        if ($checkpointDiscs.Count -gt 0) {
            throw @"
REFUSING: '$Name' already has a checkpoint '$FirstCheckpointName', and it holds a disc:
    $(($checkpointDiscs | ForEach-Object { $_.Path }) -join "`r`n    ")

It was taken before the discs came out, so it references those files, and a restore of it needs
them. This guest was built before -CompleteInstall existed - both guests built on 2026-09-15 were.
Nothing has been ejected, checkpointed or deleted. Keep the answer ISO until this guest is
replaced; a fresh build runs -CompleteInstall BEFORE its first checkpoint, and never gets here.
"@
        }
        $checkpointed = $true
        Write-Host "  '$FirstCheckpointName' already exists and holds no disc - going straight to the answer ISO."
    }

    # ---- which file is the answer ISO --------------------------------------------------------
    $answerIso = $null
    if ($AnswerIsoPath) {
        $answerIso = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AnswerIsoPath)
        if (-not (Test-IsAnswerIsoFor -Path $answerIso -VMName $Name)) {
            throw "-AnswerIsoPath '$answerIso' is not named '$Name-unattend.iso'. This deletes only the answer ISO New-AnswerFile.ps1 writes for this VM, never anything else."
        }
    }
    else {
        $attached = @($held | Where-Object { Test-IsAnswerIsoFor -Path $_.Path -VMName $Name })
        if ($attached.Count -gt 0) { $answerIso = [string]$attached[0].Path }
        else {
            $default = Join-Path $RepoRoot (Join-Path '.work\testbed-answer' (Join-Path $Name "$Name-unattend.iso"))
            if (Test-Path -LiteralPath $default) { $answerIso = $default }
        }
    }

    # ---- 1. the first logon must have consumed the answer volume -----------------------------
    if (-not $checkpointed) {
        if ($vm.State -ne 'Running') {
            throw "REFUSING: '$Name' is $($vm.State), not Running. Its first logon has to have happened, and its log is read over PowerShell Direct. Start it; if the idle-saver saved it, Testbed/host/Set-TestbedLease.ps1 keeps it awake while this runs."
        }

        $answerAttached = @($held | Where-Object { Test-IsAnswerIsoFor -Path $_.Path -VMName $Name }).Count -gt 0
        if (-not $answerAttached) {
            Write-Host '  no answer volume in either drive, so there is no first-logon run to wait for (an interactive install, or a disc already ejected)'
        }
        else {
            $credential = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $Name
            $deadline = (Get-Date).AddMinutes($InstallTimeoutMinutes)
            $lines = $null
            $unreachable = $null
            while ($true) {
                $lines = $null
                $unreachable = $null
                try {
                    $lines = @(Invoke-Command -VMId $vm.Id -Credential $credential -ErrorAction Stop -ArgumentList $FirstLogonLogPath -ScriptBlock {
                            param($path)
                            if (Test-Path -LiteralPath $path) { Get-Content -LiteralPath $path }
                        })
                }
                catch { $unreachable = $_.Exception.Message }

                $verdict = Get-FirstLogonVerdict -Lines $lines
                if (@('Done', 'Failed', 'NotFound') -contains $verdict.State) { break }
                if (-not $Execute -or (Get-Date) -ge $deadline) { break }
                Start-Sleep -Seconds 20
            }

            Write-Host ''
            Write-Host "  first-logon log ($FirstLogonLogPath): $($verdict.State)"
            foreach ($line in @($lines)) { Write-Host "    | $line" }
            if ($unreachable) { Write-Host "  the guest did not answer over PowerShell Direct: $unreachable" }

            if ($verdict.State -eq 'Failed') {
                throw "REFUSING: the first logon reported $(@($verdict.Failed).Count) failed step(s), above. Both discs stay in: the answer volume holds Complete-FirstLogon.ps1, so re-run it on the guest - elevated, from that drive - read its log, and run this again. Nothing has been ejected, checkpointed or deleted."
            }
            if ($verdict.State -eq 'NotFound') {
                throw "REFUSING: the answer file's first-logon command found Complete-FirstLogon.ps1 on no drive, so the guest's language list, locales and power policy were never set. Nothing has been ejected, checkpointed or deleted; the discs stay in for the diagnosis."
            }
            if ($verdict.State -ne 'Done') {
                $why = "No DONE line in the first-logon log after $InstallTimeoutMinutes minute(s)."
                if ($verdict.State -eq 'Absent') {
                    $why = "No first-logon log after $InstallTimeoutMinutes minute(s). Either Windows setup has not reached its first logon - look at the console - or Complete-FirstLogon.ps1 REFUSED: its guest guard runs before it opens the log, so a refusal writes no log at all (less likely, it died before its first line). Run it by hand on the guest to read which."
                }
                if ($Execute) { throw "REFUSING: $why Nothing has been ejected, checkpointed or deleted." }
                Write-Host "  -Execute would wait up to $InstallTimeoutMinutes minute(s) for the DONE line, and refuse without it."
            }
        }

        # ---- 2. eject, and 3. checkpoint --------------------------------------------------------
        if (-not $Execute) {
            Write-Host "  -Execute would eject $($held.Count) disc(s), then take '$FirstCheckpointName' with every drive empty."
        }
        else {
            foreach ($drive in $held) {
                # -Path $null empties the drive and leaves the drive itself in place.
                Set-VMDvdDrive -VMDvdDrive $drive -Path $null
            }
            $still = @(Get-VMDvdDrive -VM $vm | Where-Object { $_.Path })
            if ($still.Count -gt 0) {
                throw "Ejected, and a drive still holds a disc: $(($still | ForEach-Object { $_.Path }) -join ', '). No checkpoint has been taken and nothing deleted."
            }
            Write-Host "  ejected $($held.Count) disc(s); every DVD drive reads back empty"

            Checkpoint-VM -VM $vm -SnapshotName $FirstCheckpointName
            $firstCheckpoint = @(Get-VMSnapshot -VM $vm | Where-Object { $_.Name -eq $FirstCheckpointName })
            if ($firstCheckpoint.Count -ne 1) {
                throw "Checkpoint-VM returned, and there are $($firstCheckpoint.Count) checkpoint(s) named '$FirstCheckpointName'. The answer ISO has NOT been deleted."
            }
            $checkpointDiscs = @(Get-VMDvdDrive -VMSnapshot $firstCheckpoint[0] | Where-Object { $_.Path })
            if ($checkpointDiscs.Count -gt 0) {
                throw "'$FirstCheckpointName' was taken and reads back holding a disc: $(($checkpointDiscs | ForEach-Object { $_.Path }) -join ', '). The answer ISO has NOT been deleted."
            }
            Write-Host "  took '$FirstCheckpointName'; it reads back holding no disc"
        }
    }

    # ---- 4. delete the answer ISO, once nothing references it ---------------------------------
    Write-Host ''
    if (-not $answerIso -or -not (Test-Path -LiteralPath $answerIso)) {
        Write-Host "  no answer ISO for '$Name' to delete$(if ($answerIso) { " - $answerIso is already gone" } else { '' })."
        if ($Execute) { Write-Host ''; Write-Host "Done. Next: Testbed/README.md section 1, step 4." }
        return
    }

    $references = @()
    foreach ($other in @(Get-VM)) {
        foreach ($drive in @(Get-VMDvdDrive -VM $other | Where-Object { $_.Path })) {
            $references += [pscustomobject]@{ Where = "VM '$($other.Name)', its current configuration"; Path = [string]$drive.Path }
        }
        foreach ($checkpoint in @(Get-VMSnapshot -VM $other -ErrorAction SilentlyContinue)) {
            foreach ($drive in @(Get-VMDvdDrive -VMSnapshot $checkpoint | Where-Object { $_.Path })) {
                $references += [pscustomobject]@{ Where = "VM '$($other.Name)', checkpoint '$($checkpoint.Name)'"; Path = [string]$drive.Path }
            }
        }
    }
    $holders = @(Find-DiscReference -Path $answerIso -References $references)
    if ($holders.Count -gt 0) {
        Write-Warning ("NOT deleting the answer ISO $answerIso - it still holds the guest password, and it is still referenced by:`r`n    " +
            (($holders | ForEach-Object { $_.Where }) -join "`r`n    ") +
            "`r`nA restore of any of those needs the file. Run this again once they are gone; it goes straight to this step.")
        exit 3
    }

    $companions = @()
    $buildLog = [System.IO.Path]::ChangeExtension($answerIso, '.oscdimg.log')
    if (Test-Path -LiteralPath $buildLog) { $companions += $buildLog }
    $staging = Join-Path (Split-Path -Parent $answerIso) 'answer-volume'
    if (Test-Path -LiteralPath (Join-Path $staging 'autounattend.xml')) { $companions += $staging }

    if (-not $Execute) {
        Write-Host "  -Execute would delete the answer ISO $answerIso - nothing on this host references it -"
        foreach ($companion in $companions) { Write-Host "    and $companion" }
        return
    }

    Remove-Item -LiteralPath $answerIso -Force
    foreach ($companion in $companions) { Remove-Item -LiteralPath $companion -Recurse -Force }
    $left = @(@($answerIso) + $companions | Where-Object { Test-Path -LiteralPath $_ })
    if ($left.Count -gt 0) { throw "Deleted, and still present: $($left -join ', ')" }
    Write-Host "  deleted the answer ISO $answerIso$(if ($companions.Count -gt 0) { ', and ' + ($companions -join ', ') } else { '' }) - the guest password is no longer on this host in it"
    Write-Host ''
    Write-Host "Done. Next: Testbed/README.md section 1, step 4."
    return
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

# RESOLVE BOTH TO ABSOLUTE PATHS, and do it here rather than at each use.
# Hyper-V accepts a relative path on Add-VMDvdDrive but STORES the resolved absolute one, so the
# verification below - which reads the drive back and matches on Path - compared an absolute
# against a relative and found zero drives. It threw after the VM, its disk, Secure Boot and the
# vTPM had all been created, leaving a half-built guest with no answer volume and no boot device.
# Measured the first time this script was ever run, 2026-09-15.
$IsoPath = (Resolve-Path -LiteralPath $IsoPath).ProviderPath
if ($AnswerIsoPath) { $AnswerIsoPath = (Resolve-Path -LiteralPath $AnswerIsoPath).ProviderPath }
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
Write-Host ''
Write-Host 'Then finish it - do NOT take the first checkpoint by hand, with the discs still in:'
Write-Host ("    pwsh -File Testbed/host/New-TestbedVm.ps1 -Name {0} -CompleteInstall -Execute" -f $Name)
Write-Host 'It waits for the first-logon log, ejects the Windows ISO and the answer volume, takes'
Write-Host 'CP-01-WIN-CLEAN with no disc in it, and then deletes the answer ISO, which holds the guest'
Write-Host 'password in clear text. A checkpoint taken with a disc attached references that file for as'
Write-Host 'long as the checkpoint exists. See Docs/live-tier-on-the-vm.md section 2.1 for what follows.'
