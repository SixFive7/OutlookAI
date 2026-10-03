#Requires -Version 5.1
<#
.SYNOPSIS
    Signs the Exchange test VM's Outlook in to its Microsoft 365 mailbox - a new profile's first
    account, or a re-authentication - typing the password and a fresh TOTP code into the guest
    through its synthetic keyboard, so that no guest process or file ever holds either.

.DESCRIPTION
    RUN ON THE HOST, as the Windows user who stored the credential (DPAPI, CurrentUser), with a lease
    on the VM (Testbed/host/Set-TestbedLease.ps1). Windows PowerShell 5.1 or PowerShell 7.

    WHAT IT DOES, step by step, each one a call into the guest's Testbed/guest/Connect-ExchangeAccount.ps1
    in session 1 at -RunLevel Limited (the level the Outlook it drives runs at):

      -Mode NewProfile (Outlook's first start, no profile yet; -StartOutlook starts it, NOT elevated):
        1. Address   the "Email Account Setup" box gets the focus; the HOST types the address, Enter.
        2. Password  the Microsoft sign-in page's password field (IsPassword) gets the focus and is held;
                     the host types the password, then Enter.
        3. Code      only if the sign-in asks: the code field gets the focus; the host makes a TOTP code
                     with at least 12 s left on it and types it, then Enter.
        4. Choose "No, this app only" when Windows offers to sign in to all apps and register the
           device with the organisation - NEVER "Yes": that would add a device to the maintainer's
           tenant. Measured 2026-10-03: dsregcmd then reads AzureAdJoined NO, WorkplaceJoined NO.
        5. Clear Outlook's "Set up Outlook Mobile on my phone, too", then "Done".
        6. Verify    over COM, in session 1: an Exchange account with this address, cached.
      -Mode Reauth (an existing profile whose token has lapsed - after a checkpoint restore, say): the
        same steps from 2, for whatever sign-in page Outlook has raised.

    THE SECRETS. Testbed/host/Get-ExchangeCredential.ps1 decrypts them in memory at the moment each is
    typed; the password leaves it as a SecureString and becomes a string only for the one call that
    types it; the TOTP secret never leaves it - only a code does. Neither is printed, logged or passed
    to the guest any other way than as keystrokes: not in a script block, not as an argument, not in a
    file. Hyper-V's Msvm_Keyboard.TypeText delivers them to whatever has the keyboard focus on the
    guest's console - which is why the guest step must report READY, with the focus verified on the
    right field of the right window and HELD, before a single key is typed, and why a LOST-FOCUS, a
    NOT-FOUND or a TIMEOUT stops the run with nothing typed.

    MEASURED 2026-10-03 on OutlookAI-Exchange (Testbed/README.md section 4e): the whole of steps 1-5 in
    about three minutes, the password accepted without an MFA prompt that time, the account added in
    cached mode.

    WHY ONE PowerShell Direct SESSION. A second PowerShell Direct connection opened while one is being
    set up failed with "The credential is invalid" - and each such failure is a failed logon on the
    guest, whose lockout threshold is 10 in 10 minutes (measured: five failures from two concurrent
    attempts). So every guest call here goes through one session, sequentially, and the step that has
    to stay alive while the host types runs as a background job INSIDE the guest.

.PARAMETER VMName
    MANDATORY. Must be the VM testbed.json's exchangeVm block names - the script refuses any other:
    typing a mailbox password into a VM nobody meant is the mistake it exists to make impossible.

.PARAMETER Mode
    NewProfile, Reauth or Preflight. Preflight types nothing: before every live run it brings Defender's
    signatures up to date and waits for Security Center to say so - while they are stale, Outlook's
    Object Model Guard puts an Allow/Deny prompt in front of every out-of-process address read, and a
    run hangs on it (measured 2026-10-03: six stacked prompts, after a checkpoint restore brought back
    year-old signatures) - and then reads whether the account is still signed in. Exit 0 when both
    hold; when only the sign-in does not, run -Mode Reauth.

.PARAMETER StartOutlook
    With -Mode NewProfile: start OUTLOOK.EXE in session 1, NOT elevated, before step 1.

.PARAMETER RepoRoot
    The main checkout: the guest credential and the Exchange credential are gitignored and live there.

.PARAMETER SelfTest
    The pure parts - reading a step's status file, the decisions taken from it, the VM-name guard.
    No VM, no credential.
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')] [string] $VMName,
    [Parameter(ParameterSetName = 'Run')] [ValidateSet('NewProfile', 'Reauth', 'Preflight')] [string] $Mode = 'NewProfile',
    [Parameter(ParameterSetName = 'Run')] [switch] $StartOutlook,
    [Parameter(ParameterSetName = 'Run')] [string] $RepoRoot,
    [Parameter(ParameterSetName = 'Run')] [int] $PageWaitSeconds = 90,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)
if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) }
$ErrorActionPreference = 'Stop'

# =============================================================================================
# PURE DECISIONS
# =============================================================================================

# A status file as Connect-ExchangeAccount.ps1 writes it: 'STATE <word>', 'AT <time>', detail lines.
function ConvertFrom-StepStatus {
    param([string[]] $Lines)
    $all = @($Lines | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0 -or $all[0] -notmatch '^STATE (\S+)$') {
        return [pscustomobject]@{ State = 'NONE'; Detail = @() }
    }
    $state = $Matches[1]
    $detail = @()
    if ($all.Count -gt 2) { $detail = @($all[2..($all.Count - 1)]) }
    return [pscustomobject]@{ State = $state; Detail = $detail }
}

# Whether the host may type now: only on READY, and only when the focus line the guest wrote says
# the focus is where this step needs it - a password field for the password.
function Test-MayType {
    param($Status, [string] $Step)
    if ($null -eq $Status -or $Status.State -ne 'READY') { return $false }
    $focus = @($Status.Detail | Where-Object { $_ -like 'FOCUS *' })
    if ($focus.Count -ne 1) { return $false }
    if ($Step -eq 'Password') { return ($focus[0].Contains(' password=True ')) }
    return ($focus[0].Contains(' password=False '))
}

# The VM this may type into: exactly the one testbed.json's exchangeVm names.
function Get-ExchangeVmProblem {
    param([string] $VMName, $Testbed)
    $record = $null
    if ($null -ne $Testbed -and $Testbed.PSObject.Properties.Name -contains 'exchangeVm') { $record = $Testbed.exchangeVm }
    if ($null -eq $record -or -not ($record.vmName -is [string]) -or -not $record.vmName) {
        return 'REFUSING: Testbed/testbed.json records no exchangeVm, so no VM is the Exchange test VM.'
    }
    if ($VMName -cne $record.vmName) {
        return "REFUSING: '$VMName' is not the Exchange test VM ('$($record.vmName)' in Testbed/testbed.json). This script types a mailbox password into the VM it is given."
    }
    return $null
}

if ($SelfTest) {
    Write-Host "Invoke-ExchangeSignIn.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). No VM, no credential."
    $script:fails = 0
    function Check([string] $Name, $Expected, $Actual) {
        if ("$Expected" -ceq "$Actual") { Write-Host "  PASS  $Name" } else { $script:fails++; Write-Host "  FAIL  $Name - expected '$Expected', got '$Actual'" }
    }
    $ready = ConvertFrom-StepStatus @('STATE READY', 'AT 2026-10-03T19:54:03', 'window "" (explorer) handle=1', 'FOCUS Edit aid="i0118" class="" password=True process=Microsoft.AAD.BrokerPlugin topHandle=1 foreground=1')
    Check 'a status file is read' 'READY' $ready.State
    Check 'its detail lines are kept' 2 $ready.Detail.Count
    Check 'a password page held on a password field may be typed into' $true (Test-MayType $ready 'Password')
    Check 'but not as a code page' $false (Test-MayType $ready 'Code')
    $plain = ConvertFrom-StepStatus @('STATE READY', 'AT x', 'window "Email Account Setup" (OUTLOOK) handle=2', 'FOCUS Pane aid="1" class="RICHEDIT60W" password=False process=OUTLOOK topHandle=2 foreground=2')
    Check 'an address box may take the address' $true (Test-MayType $plain 'Address')
    Check 'and never the password' $false (Test-MayType $plain 'Password')
    Check 'LOST-FOCUS types nothing' $false (Test-MayType (ConvertFrom-StepStatus @('STATE LOST-FOCUS', 'AT x', 'FOCUS Edit aid="i0118" class="" password=True')) 'Password')
    Check 'NOT-FOUND types nothing' $false (Test-MayType (ConvertFrom-StepStatus @('STATE NOT-FOUND', 'AT x')) 'Code')
    Check 'a READY with no focus line types nothing' $false (Test-MayType (ConvertFrom-StepStatus @('STATE READY', 'AT x')) 'Address')
    Check 'an empty file is no state' 'NONE' (ConvertFrom-StepStatus @()).State
    Check 'a garbled first line is no state' 'NONE' (ConvertFrom-StepStatus @('READY')).State
    $record = [pscustomobject]@{ exchangeVm = [pscustomobject]@{ vmName = 'OutlookAI-Exchange' } }
    Check 'the recorded VM is accepted' '' ([string](Get-ExchangeVmProblem 'OutlookAI-Exchange' $record))
    Check 'another VM is refused' $true ((Get-ExchangeVmProblem 'OutlookAI-Indexed' $record) -like 'REFUSING*')
    Check 'a case variant is refused' $true ((Get-ExchangeVmProblem 'outlookai-exchange' $record) -like 'REFUSING*')
    Check 'no record refuses' $true ((Get-ExchangeVmProblem 'OutlookAI-Exchange' ([pscustomobject]@{ vmName = 'x' })) -like 'REFUSING*')
    $real = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json') -Raw | ConvertFrom-Json
    Check 'and testbed.json records the Exchange VM this script is for' '' ([string](Get-ExchangeVmProblem 'OutlookAI-Exchange' $real))
    Write-Host ''
    if ($script:fails -gt 0) { Write-Host "$script:fails FAILED"; exit 1 }
    Write-Host 'All passed.'
    exit 0
}

# =============================================================================================
# THE RUN
# =============================================================================================
# The record this script belongs to - its own checkout's, not -RepoRoot's, which is only where the
# gitignored credentials live (from a worktree, the main checkout).
$testbed = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'testbed.json') -Raw | ConvertFrom-Json
$problem = Get-ExchangeVmProblem $VMName $testbed
if ($null -ne $problem) { Write-Host $problem; exit 4 }

function Say([string] $m) { Write-Host ('[{0:HH:mm:ss}] {1}' -f (Get-Date), $m) }

$guestCred = & (Join-Path $PSScriptRoot 'Get-GuestCredential.ps1') -RepoRoot $RepoRoot -VMName $VMName
$account = (& (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot).Account

$guest = @(Get-CimInstance -Namespace 'root\virtualization\v2' -ClassName 'Msvm_ComputerSystem' -Filter "ElementName='$VMName'")
if ($guest.Count -ne 1) { Write-Host "REFUSING: expected one Msvm_ComputerSystem named '$VMName', found $($guest.Count)."; exit 4 }
$keyboard = @($guest[0] | Get-CimAssociatedInstance -ResultClassName 'Msvm_Keyboard')
if ($keyboard.Count -lt 1) { Write-Host "REFUSING: '$VMName' has no synthetic keyboard."; exit 4 }
$keyboard = $keyboard[0]

function Send-Text([string] $Text) {
    $r = Invoke-CimMethod -InputObject $keyboard -MethodName 'TypeText' -Arguments @{ asciiText = $Text }
    if ($r.ReturnValue -ne 0) { throw "the synthetic keyboard refused the text (TypeText returned $($r.ReturnValue))" }
}
function Send-Enter {
    $r = Invoke-CimMethod -InputObject $keyboard -MethodName 'TypeKey' -Arguments @{ keyCode = [uint32]0x0D }
    if ($r.ReturnValue -ne 0) { throw "the synthetic keyboard refused Enter (TypeKey returned $($r.ReturnValue))" }
}

$session = New-PSSession -VMName $VMName -Credential $guestCred

# A step that only reads or invokes: run it in session 1 and wait for it.
function Invoke-GuestStep([string] $Arguments, [int] $Timeout = 300) {
    $lines = Invoke-Command -Session $session -ScriptBlock {
        param($a, $t)
        Set-ExecutionPolicy -Scope Process Bypass -Force
        Set-Location C:\OutlookAI-Q5
        # ps51-native-stderr-ok: a PowerShell script on the guest, not a program, run inside an Invoke-Command block where the host's 'Stop' does not apply - the remote session's default is 'Continue'
        & .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds $t -Script "& 'C:\OutlookAI-Q5\Connect-ExchangeAccount.ps1' $a" *>&1 | Out-Null
        $step = ($a -split ' ')[1]
        Get-Content -LiteralPath "C:\OutlookAI-Q5\exchange\signin\$step.status.txt" -ErrorAction SilentlyContinue
    } -ArgumentList $Arguments, $Timeout
    return (ConvertFrom-StepStatus @($lines))
}

# A step the host types into: the guest holds the focus in a background job while the host types.
function Invoke-TypedStep([string] $Step, [scriptblock] $GetText, [int] $WaitSeconds) {
    Invoke-Command -Session $session -ScriptBlock {
        param($s, $w)
        Remove-Item "C:\OutlookAI-Q5\exchange\signin\$s.status.txt", "C:\OutlookAI-Q5\exchange\signin\$s.go" -Force -ErrorAction SilentlyContinue
        $global:typedStepJob = Start-Job -ScriptBlock {
            param($s, $w)
            Set-ExecutionPolicy -Scope Process Bypass -Force
            Set-Location C:\OutlookAI-Q5
            # ps51-native-stderr-ok: a PowerShell script on the guest, not a program, run inside an Invoke-Command block where the host's 'Stop' does not apply - the remote session's default is 'Continue'
            & .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds ($w + 200) -Script "& 'C:\OutlookAI-Q5\Connect-ExchangeAccount.ps1' -Step $s -WaitSeconds $w -HoldSeconds 60" *>&1 | Out-Null
        } -ArgumentList $s, $w
    } -ArgumentList $Step, $WaitSeconds

    $status = $null
    $deadline = (Get-Date).AddSeconds($WaitSeconds + 90)
    while ((Get-Date) -lt $deadline) {
        $lines = Invoke-Command -Session $session -ScriptBlock { param($s) $p = "C:\OutlookAI-Q5\exchange\signin\$s.status.txt"; if (Test-Path $p) { Get-Content $p } } -ArgumentList $Step
        $status = ConvertFrom-StepStatus @($lines)
        if ($status.State -ne 'NONE') { break }
        Start-Sleep -Milliseconds 500
    }

    $typed = $false
    if (Test-MayType $status $Step) {
        $text = & $GetText
        try { Send-Text $text } finally { $text = $null }
        Invoke-Command -Session $session -ScriptBlock { param($s) Set-Content -Path "C:\OutlookAI-Q5\exchange\signin\$s.go" -Value 'typed' } -ArgumentList $Step
        Start-Sleep -Seconds 2
        $after = ConvertFrom-StepStatus @(Invoke-Command -Session $session -ScriptBlock { param($s) Get-Content "C:\OutlookAI-Q5\exchange\signin\$s.status.txt" } -ArgumentList $Step)
        if ($after.State -eq 'DONE') { Send-Enter; $typed = $true }
        else { Say "  the focus did not hold while typing ($($after.State)) - Enter NOT pressed" }
    }
    Invoke-Command -Session $session -ScriptBlock { if ($global:typedStepJob) { $null = Wait-Job $global:typedStepJob -Timeout 240; Remove-Job $global:typedStepJob -Force } } | Out-Null
    return [pscustomobject]@{ Status = $status; Typed = $typed }
}

$exit = 1
try {
    # Every guest step is a scheduled task whose powershell.exe is started -WindowStyle Hidden - which
    # Windows Terminal, the default console host on this Windows 11 build, does not honour: measured
    # 2026-10-03, the task's console came up as a visible Windows Terminal window, and a visible window
    # can take the keyboard focus the host is about to type into. The documented per-user setting
    # (Settings > Privacy & security > For developers > Terminal: Windows Console Host) puts consoles
    # back in conhost, which does honour it. Set before anything else, every run; it is idempotent.
    Invoke-Command -Session $session -ScriptBlock {
        $key = 'HKCU:\Console\%%Startup'
        if (-not (Test-Path -LiteralPath $key)) { New-Item -Path $key -Force | Out-Null }
        Set-ItemProperty -LiteralPath $key -Name DelegationConsole -Value '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
        Set-ItemProperty -LiteralPath $key -Name DelegationTerminal -Value '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
    }

    if ($Mode -eq 'Preflight') {
        Say 'preflight 1: Defender signatures, so that Outlook''s Object Model Guard does not prompt'
        $av = Invoke-Command -Session $session -ScriptBlock {
            $updated = $true
            try { Update-MpSignature -ErrorAction Stop } catch { $updated = $false }
            $state = 0x10
            $deadline = (Get-Date).AddMinutes(3)
            do {
                $product = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | Where-Object { $_.displayName -eq 'Windows Defender' } | Select-Object -First 1
                if ($null -ne $product) { $state = [int]$product.productState }
                if (($state -band 0xFF) -eq 0) { break }
                Start-Sleep -Seconds 10
            } while ((Get-Date) -lt $deadline)
            # Defender's WMI provider can be briefly unavailable right after an update ("Provider load
            # failure", measured): the age is a detail, the Security Center state above is the answer.
            $age = -1
            try { $age = (Get-MpComputerStatus -ErrorAction Stop).AntivirusSignatureAge } catch { $age = -1 }
            [pscustomobject]@{ Updated = $updated; State = $state; Age = $age }
        }
        if (($av.State -band 0xFF) -ne 0) {
            Say ('  Security Center still reports the definitions out of date (0x{0:X6}, update ran: {1}) - a live run would meet the guard''s prompts' -f $av.State, $av.Updated)
            exit 1
        }
        Say ('  up to date (productState 0x{0:X6}); signatures {1} day(s) old' -f $av.State, $av.Age)
        Say 'preflight 2: is the account still signed in?'
        $v = Invoke-GuestStep "-Step Verify -Account '$account' -WaitSeconds 60"
        Say "  $($v.State) $($v.Detail -join ' | ')"
        if ($v.State -ne 'SIGNED-IN') { Say '  not signed in - run this script with -Mode Reauth'; exit 1 }
        exit 0
    }

    if ($Mode -eq 'NewProfile') {
        if ($StartOutlook) {
            Say 'starting Outlook in session 1, NOT elevated (no profile yet: it opens its account setup)'
            Invoke-Command -Session $session -ScriptBlock {
                Set-ExecutionPolicy -Scope Process Bypass -Force
                Set-Location C:\OutlookAI-Q5
                # ps51-native-stderr-ok: a PowerShell script on the guest, not a program, run inside an Invoke-Command block where the host's 'Stop' does not apply - the remote session's default is 'Continue'
                & .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 120 -Script "Start-Process 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'" *>&1 | Out-Null
            }
            Start-Sleep -Seconds 20
        }
        Say "step 1: the address ($account)"
        $r = Invoke-TypedStep -Step 'Address' -GetText { $account } -WaitSeconds $PageWaitSeconds
        if (-not $r.Typed) { Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"; exit 1 }
    }

    Say 'step 2: the password (typed from the host; never shown)'
    $r = Invoke-TypedStep -Step 'Password' -WaitSeconds $PageWaitSeconds -GetText {
        $secure = (& (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot).Password
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
    if (-not $r.Typed) { Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"; exit 1 }

    Say 'step 3: a verification code, if the sign-in asks for one'
    $r = Invoke-TypedStep -Step 'Code' -WaitSeconds 25 -GetText {
        (& (Join-Path $PSScriptRoot 'Get-ExchangeCredential.ps1') -RepoRoot $RepoRoot -TotpCode -MinimumValiditySeconds 12).Code
    }
    if ($r.Typed) { Say '  a code was asked for and typed' }
    elseif ($r.Status.State -eq 'NOT-FOUND') { Say '  none asked for' }
    else { Say "  stopped: $($r.Status.State) $($r.Status.Detail -join ' | ')"; exit 1 }

    Say "step 4: decline signing in to all apps and registering the device"
    $c = Invoke-GuestStep "-Step Choose -ControlName 'No, this app only' -WaitSeconds 40"
    Say "  $($c.State) $($c.Detail -join ' | ')"

    if ($Mode -eq 'NewProfile') {
        Say "step 5: clear 'Set up Outlook Mobile on my phone, too', then Done"
        $c = Invoke-GuestStep "-Step Choose -ControlName 'Set up Outlook Mobile on my phone, too' -WaitSeconds 60"
        Say "  $($c.State) $($c.Detail -join ' | ')"
        if ($c.State -ne 'DONE') { exit 1 }
        $c = Invoke-GuestStep "-Step Choose -ControlName 'Done' -WaitSeconds 30"
        Say "  $($c.State) $($c.Detail -join ' | ')"
        if ($c.State -ne 'DONE') { exit 1 }
        Start-Sleep -Seconds 20
    }

    Say 'step 6: verify over COM'
    $v = Invoke-GuestStep "-Step Verify -Account '$account' -WaitSeconds 120"
    Say "  $($v.State) $($v.Detail -join ' | ')"
    if ($v.State -eq 'SIGNED-IN') { $exit = 0 }
}
finally {
    Remove-PSSession $session
}
exit $exit
