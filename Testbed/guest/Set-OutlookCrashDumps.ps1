#Requires -Version 5.1
<#
.SYNOPSIS
    Makes a test guest keep a FULL crash dump of OUTLOOK.EXE (and WINWORD.EXE) in one known folder,
    C:\OutlookAI-Q5\crash-dumps, through Windows Error Reporting's LocalDumps key - Windows' own
    mechanism, nothing installed.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`.

    WHY (2026-10-04). OUTLOOK.EXE crashed on both Outlook guests during the live tier's compose tests
    - `wwlib.dll` once on OutlookAI-Indexed, `OLMAPI32.DLL` on OutlookAI-Unindexed - and the guests
    kept nothing but the Application log's one-line "Application Error" event, so there was nothing
    to analyse (Docs/live-tier-on-the-vm.md sections 4.2f, 4.5 and 4.6). WER's LocalDumps key makes
    WerFault.exe write the crashing process's whole memory to a folder before the process goes.

    WHAT IT WRITES, under HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps\<program>:
      DumpFolder  REG_EXPAND_SZ  the folder (-DumpFolder)
      DumpType    REG_DWORD      2 - a full dump: every committed page, the heap included
      DumpCount   REG_DWORD      10 - WER keeps the newest ten per folder and deletes older ones
    for OUTLOOK.EXE; WINWORD.EXE, in case Word ever runs as its own process (the compose editor is
    Word INSIDE OUTLOOK.EXE - wwlib.dll in Outlook's process - so an editor crash is an OUTLOOK.EXE
    dump); and OutlookAI.CrashDumpProbe.exe, the throwaway program -Prove crashes. It also makes the
    folder and grants Authenticated Users Modify on it: Outlook runs NOT elevated (Testbed/README.md
    section 4c), and a not-elevated token has the Administrators group as deny-only, so a folder only
    administrators may write would get no dump.

    A test guest is restored from a checkpoint at every run, and the frozen checkpoints were taken
    before this existed, so Testbed/host/Invoke-LiveTierOnGuest.ps1 runs -Execute and -Verify on
    every run, after the restore - no checkpoint has to change, and the frozen clocks (runbook
    section 4.5) are not touched.

    MODES:
      (none)    Dry run: what it would write, and what is there now. Changes nothing.
      -Execute  Writes the keys and the folder, then verifies. ELEVATED (session 0 over PowerShell
                Direct is).
      -Verify   Read-only: ARMED, exit 0, or NOT-ARMED, exit 1, naming each part that is missing -
                a key value, the folder, or Windows Error Reporting itself disabled (its Disabled
                value, the policy's, or the WerSvc service).
      -Prove    Proves the mechanism end to end WITHOUT touching Outlook: copies rundll32.exe to
                OutlookAI.CrashDumpProbe.exe, runs it on kernel32's DebugBreak - an unhandled
                breakpoint, which WER handles as a crash - and waits for a FULL dump of it in the
                folder: PROVEN, exit 0, or NOT-PROVEN, exit 1. Run it as Outlook runs - session 1,
                NOT elevated: Register-InteractiveTask.ps1 -RunLevel Limited - after -Execute. The
                probe's dump and copy are deleted afterwards; its Application Error event stays.
      -SelfTest Pure: the plan, the verdict and the dump-header reader. No registry, no file.

    EXIT CODES: 0 ARMED / PROVEN / the dry run; 1 NOT-ARMED / NOT-PROVEN, or REFUSING TO RUN off a test
    guest (it throws); 4 -Execute not elevated.

.EXAMPLE
    # On the guest, elevated (what the runner does at every run):
    & 'C:\OutlookAI-Q5\Set-OutlookCrashDumps.ps1' -Execute
    # Once per guest, as Outlook runs:
    .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 300 -Script "& 'C:\OutlookAI-Q5\Set-OutlookCrashDumps.ps1' -Prove"
    # Anywhere, pure:
    powershell.exe -NoProfile -File Testbed\guest\Set-OutlookCrashDumps.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $Prove,
    [switch]   $SelfTest,
    [string]   $DumpFolder = 'C:\OutlookAI-Q5\crash-dumps',
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [int]      $ProveTimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'

$LocalDumpsKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps'
$WerKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting',
    'HKCU:\Software\Microsoft\Windows\Windows Error Reporting'
)
$ProbeProgram = 'OutlookAI.CrashDumpProbe.exe'
$Programs = @('OUTLOOK.EXE', 'WINWORD.EXE', $ProbeProgram)
$FullDumpType = 2
$KeptDumps = 10
# MINIDUMP_HEADER: 'MDMP', and MiniDumpWithFullMemory in its Flags (minidumpapiset.h).
$MinidumpSignature = 0x504D444D
$MiniDumpWithFullMemory = 0x2

# =============================================================================================
# PURE DECISIONS. -SelfTest pins them.
# =============================================================================================

# The values each program's key must hold.
function Get-DumpPlan {
    param([string] $Folder)
    foreach ($program in $Programs) {
        [pscustomobject]@{ Program = $program; DumpFolder = $Folder; DumpType = $FullDumpType; DumpCount = $KeptDumps }
    }
}

# ARMED when every program's key holds the plan's values, the folder exists and WER is on; else the
# parts that are not. $Observed: Keys (program -> @{DumpFolder; DumpType; DumpCount}, a missing key
# $null), WerDisabled (the Disabled values found set, by key), WerSvcStartType, FolderExists.
function Get-ArmedVerdict {
    param([object[]] $Plan, [hashtable] $Observed)
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Plan) {
        $seen = $Observed.Keys[$p.Program]
        if ($null -eq $seen) { $missing.Add("no LocalDumps key for $($p.Program)"); continue }
        if ([string]$seen.DumpFolder -ne $p.DumpFolder) { $missing.Add("$($p.Program): DumpFolder is '$($seen.DumpFolder)', not '$($p.DumpFolder)'") }
        if ($seen.DumpType -ne $p.DumpType) { $missing.Add("$($p.Program): DumpType is '$($seen.DumpType)', not $($p.DumpType) (full)") }
        if ($seen.DumpCount -ne $p.DumpCount) { $missing.Add("$($p.Program): DumpCount is '$($seen.DumpCount)', not $($p.DumpCount)") }
    }
    foreach ($k in @($Observed.WerDisabled)) { if ($k) { $missing.Add("Windows Error Reporting is disabled: $k\Disabled is set") } }
    if ([string]$Observed.WerSvcStartType -eq 'Disabled') { $missing.Add('the WerSvc service is disabled') }
    if (-not $Observed.FolderExists) { $missing.Add("the folder $($Plan[0].DumpFolder) does not exist") }
    $verdict = 'ARMED'
    if ($missing.Count -gt 0) { $verdict = 'NOT-ARMED' }
    return [pscustomobject]@{ Verdict = $verdict; Missing = @($missing) }
}

# What a file's first 32 bytes say: a minidump at all, and a FULL one.
function Read-MinidumpHeader {
    param([byte[]] $Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 32) { return [pscustomobject]@{ IsMinidump = $false; Flags = 0; FullMemory = $false } }
    $signature = [BitConverter]::ToUInt32($Bytes, 0)
    $flags = [BitConverter]::ToUInt64($Bytes, 24)
    return [pscustomobject]@{
        IsMinidump = ($signature -eq $MinidumpSignature)
        Flags      = $flags
        FullMemory = (($signature -eq $MinidumpSignature) -and (($flags -band $MiniDumpWithFullMemory) -ne 0))
    }
}

function Invoke-SelfTest {
    $script:stFailures = New-Object System.Collections.Generic.List[string]
    $script:stChecks = 0
    function Check([string] $Name, $Expected, $Actual) {
        $script:stChecks++
        if ([string]$Expected -cne [string]$Actual) { $script:stFailures.Add("$Name - expected '$Expected', got '$Actual'") }
    }
    Write-Host "Set-OutlookCrashDumps.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Pure: no registry, no file."

    $plan = @(Get-DumpPlan 'C:\OutlookAI-Q5\crash-dumps')
    Check 'the plan covers three programs' 3 $plan.Count
    Check 'Outlook first' 'OUTLOOK.EXE' $plan[0].Program
    Check 'Word' 'WINWORD.EXE' $plan[1].Program
    Check 'the probe' 'OutlookAI.CrashDumpProbe.exe' $plan[2].Program
    Check 'a FULL dump' 2 $plan[0].DumpType
    Check 'ten kept' 10 $plan[0].DumpCount
    Check 'the folder' 'C:\OutlookAI-Q5\crash-dumps' $plan[0].DumpFolder

    $good = @{}
    foreach ($p in $plan) { $good[$p.Program] = @{ DumpFolder = $p.DumpFolder; DumpType = 2; DumpCount = 10 } }
    $v = Get-ArmedVerdict $plan @{ Keys = $good; WerDisabled = @(); WerSvcStartType = 'Manual'; FolderExists = $true }
    Check 'everything in place is ARMED' 'ARMED' $v.Verdict
    Check '... with nothing missing' 0 @($v.Missing).Count

    $noWord = @{}
    foreach ($k in $good.Keys) { $noWord[$k] = $good[$k] }
    $noWord['WINWORD.EXE'] = $null
    $v = Get-ArmedVerdict $plan @{ Keys = $noWord; WerDisabled = @(); WerSvcStartType = 'Manual'; FolderExists = $true }
    Check 'a missing key is NOT-ARMED' 'NOT-ARMED' $v.Verdict
    Check '... and names the program' 'no LocalDumps key for WINWORD.EXE' (@($v.Missing)[0])

    $mini = @{}
    foreach ($k in $good.Keys) { $mini[$k] = $good[$k] }
    $mini['OUTLOOK.EXE'] = @{ DumpFolder = 'C:\OutlookAI-Q5\crash-dumps'; DumpType = 1; DumpCount = 10 }
    $v = Get-ArmedVerdict $plan @{ Keys = $mini; WerDisabled = @(); WerSvcStartType = 'Manual'; FolderExists = $true }
    Check 'a MINI dump is NOT-ARMED' 'NOT-ARMED' $v.Verdict
    Check '... and says full' "OUTLOOK.EXE: DumpType is '1', not 2 (full)" (@($v.Missing)[0])

    $v = Get-ArmedVerdict $plan @{ Keys = $good; WerDisabled = @('HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting'); WerSvcStartType = 'Manual'; FolderExists = $true }
    Check 'WER disabled by policy is NOT-ARMED' 'NOT-ARMED' $v.Verdict
    $v = Get-ArmedVerdict $plan @{ Keys = $good; WerDisabled = @(); WerSvcStartType = 'Disabled'; FolderExists = $true }
    Check 'WerSvc disabled is NOT-ARMED' 'the WerSvc service is disabled' (@($v.Missing)[0])
    $v = Get-ArmedVerdict $plan @{ Keys = $good; WerDisabled = @(); WerSvcStartType = 'Manual'; FolderExists = $false }
    Check 'no folder is NOT-ARMED' 'NOT-ARMED' $v.Verdict

    $header = New-Object byte[] 32
    [BitConverter]::GetBytes([uint32]0x504D444D).CopyTo($header, 0)
    [BitConverter]::GetBytes([uint64]0x1826).CopyTo($header, 24)
    $h = Read-MinidumpHeader $header
    Check 'MDMP is a minidump' 'True' $h.IsMinidump
    Check "WER's full-dump flags carry full memory" 'True' $h.FullMemory
    [BitConverter]::GetBytes([uint64]0x1824).CopyTo($header, 24)
    Check 'flags without 0x2 are not full' 'False' (Read-MinidumpHeader $header).FullMemory
    [BitConverter]::GetBytes([uint32]0x5A4D).CopyTo($header, 0)
    Check 'another file is not a minidump' 'False' (Read-MinidumpHeader $header).IsMinidump
    Check 'a short file is not a minidump' 'False' (Read-MinidumpHeader (New-Object byte[] 8)).IsMinidump

    if ($script:stFailures.Count -gt 0) {
        foreach ($f in $script:stFailures) { Write-Host "FAIL: $f" }
        Write-Host "SelfTest: $($script:stChecks - $script:stFailures.Count) passed, $($script:stFailures.Count) failed."
        return 1
    }
    Write-Host "SelfTest: $($script:stChecks) passed, 0 failed."
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE GUEST.
# =============================================================================================

function Assert-TestbedGuestLocal {
    $who = $env:USERNAME
    $machine = $env:COMPUTERNAME
    $userOk = @($ExpectedUser | Where-Object { $_ -eq $who }).Count -gt 0
    $machineOk = $machine -and $machine.StartsWith($ExpectedComputerNamePrefix, [System.StringComparison]::OrdinalIgnoreCase)
    if ($userOk -and $machineOk) { return }
    throw @"
REFUSING TO RUN.

  logged on as : '$who'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$machine'      (must start with: '$ExpectedComputerNamePrefix')

This script changes Windows Error Reporting for the whole machine and makes it write the full memory
of a crashing Outlook - mail included - to a folder. That belongs on a test guest, never on a machine
with real mail on it. The testbed guests autologon as 'vmadmin' and are named 'OAI-...'
(Testbed/host/New-AnswerFile.ps1). Do not 'fix' this by widening either default.
"@
}

function Read-Observed {
    $keys = @{}
    foreach ($program in $Programs) {
        $path = Join-Path $LocalDumpsKey $program
        $keys[$program] = $null
        if (Test-Path -LiteralPath $path) {
            $item = Get-Item -LiteralPath $path
            $keys[$program] = @{
                DumpFolder = $item.GetValue('DumpFolder', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                DumpType   = $item.GetValue('DumpType', $null)
                DumpCount  = $item.GetValue('DumpCount', $null)
            }
        }
    }
    $disabled = @()
    foreach ($k in $WerKeys) {
        if (Test-Path -LiteralPath $k) {
            $value = (Get-Item -LiteralPath $k).GetValue('Disabled', $null)
            if ($null -ne $value -and [int]$value -ne 0) { $disabled += $k }
        }
    }
    $svc = Get-Service -Name WerSvc -ErrorAction SilentlyContinue
    $startType = 'absent'
    if ($svc) { $startType = [string]$svc.StartType }
    return @{ Keys = $keys; WerDisabled = $disabled; WerSvcStartType = $startType; FolderExists = (Test-Path -LiteralPath $DumpFolder -PathType Container) }
}

function Show-Verdict([object] $Verdict) {
    if ($Verdict.Verdict -eq 'ARMED') {
        Write-Host "VERDICT: ARMED - a crashing $($Programs -join ', ') leaves a full dump in $DumpFolder (at most $KeptDumps each)."
        return 0
    }
    Write-Host 'VERDICT: NOT-ARMED'
    foreach ($m in $Verdict.Missing) { Write-Host "  - $m" }
    return 1
}

Assert-TestbedGuestLocal
$plan = @(Get-DumpPlan $DumpFolder)

if ($Prove) {
    # Session 1, NOT elevated - as Outlook runs - so the proof covers the folder's permissions too.
    $verdict = Get-ArmedVerdict $plan (Read-Observed)
    if ($verdict.Verdict -ne 'ARMED') { [void](Show-Verdict $verdict); Write-Host 'VERDICT: NOT-PROVEN - not armed; run -Execute first.'; exit 1 }
    $probeDir = Join-Path $env:TEMP 'OutlookAI-CrashDumpProbe'
    New-Item -ItemType Directory -Force -Path $probeDir | Out-Null
    $probe = Join-Path $probeDir $ProbeProgram
    Copy-Item -LiteralPath (Join-Path $env:WINDIR 'System32\rundll32.exe') -Destination $probe -Force
    $before = @(Get-ChildItem -LiteralPath $DumpFolder -Filter "$ProbeProgram.*.dmp" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    $started = Get-Date
    $process = Start-Process -FilePath $probe -ArgumentList 'kernel32.dll,DebugBreak' -PassThru -WindowStyle Hidden
    $null = $process.Handle   # held now, or ExitCode reads empty after the exit (Start-Process -PassThru)
    $dump = $null
    $deadline = (Get-Date).AddSeconds($ProveTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $new = @(Get-ChildItem -LiteralPath $DumpFolder -Filter "$ProbeProgram.$($process.Id).dmp" -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.FullName })
        # Complete once WerFault has let go of it: the probe is gone and the file opens for reading alone.
        if ($new.Count -gt 0 -and $process.HasExited -and -not @(Get-Process -Name WerFault -ErrorAction SilentlyContinue).Count) {
            $dump = $new[0]
            break
        }
        Start-Sleep -Seconds 2
    }
    $exitText = 'still running'
    if ($process.HasExited) { $exitText = ('0x{0:X8}' -f $process.ExitCode) }
    if ($null -eq $dump) {
        Write-Host "VERDICT: NOT-PROVEN - no dump of $ProbeProgram (pid $($process.Id), exit $exitText) appeared in $DumpFolder within $ProveTimeoutSeconds s."
        if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
        exit 1
    }
    $stream = [System.IO.File]::OpenRead($dump.FullName)
    try { $bytes = New-Object byte[] 32; [void]$stream.Read($bytes, 0, 32) } finally { $stream.Dispose() }
    $header = Read-MinidumpHeader $bytes
    $line = "the probe (pid $($process.Id), exit $exitText) left $($dump.Name), $([math]::Round($dump.Length / 1MB, 1)) MB, flags 0x$('{0:X}' -f $header.Flags), $([math]::Round(((Get-Date) - $started).TotalSeconds, 1)) s after it started"
    Remove-Item -LiteralPath $dump.FullName -Force
    Remove-Item -LiteralPath $probeDir -Recurse -Force -ErrorAction SilentlyContinue
    if ($header.FullMemory) { Write-Host "VERDICT: PROVEN - $line; a FULL dump. Deleted again."; exit 0 }
    Write-Host "VERDICT: NOT-PROVEN - $line; NOT a full dump (MiniDumpWithFullMemory absent). Deleted again."
    exit 1
}

if (-not $Execute) {
    $observed = Read-Observed
    if (-not $Verify) {
        Write-Host 'DRY RUN - nothing changed. -Execute would write, under'
        Write-Host "  $LocalDumpsKey :"
        foreach ($p in $plan) { Write-Host ("    {0,-30} DumpFolder={1} DumpType={2} (full) DumpCount={3}" -f $p.Program, $p.DumpFolder, $p.DumpType, $p.DumpCount) }
        Write-Host "  and make $DumpFolder, Authenticated Users Modify. What is there now:"
    }
    exit (Show-Verdict (Get-ArmedVerdict $plan $observed))
}

$principal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'REFUSING TO RUN: -Execute writes HKLM and needs an elevated session (session 0 over PowerShell Direct is one).'
    exit 4
}
New-Item -ItemType Directory -Force -Path $DumpFolder | Out-Null
$acl = Get-Acl -LiteralPath $DumpFolder
$authenticatedUsers = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11')
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($authenticatedUsers, 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
$acl.AddAccessRule($rule)
Set-Acl -LiteralPath $DumpFolder -AclObject $acl
foreach ($p in $plan) {
    $path = Join-Path $LocalDumpsKey $p.Program
    if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force | Out-Null }
    New-ItemProperty -LiteralPath $path -Name DumpFolder -PropertyType ExpandString -Value $p.DumpFolder -Force | Out-Null
    New-ItemProperty -LiteralPath $path -Name DumpType -PropertyType DWord -Value $p.DumpType -Force | Out-Null
    New-ItemProperty -LiteralPath $path -Name DumpCount -PropertyType DWord -Value $p.DumpCount -Force | Out-Null
}
Write-Host "written: $($Programs.Count) LocalDumps keys, $DumpFolder (Authenticated Users Modify)"
exit (Show-Verdict (Get-ArmedVerdict $plan (Read-Observed)))
