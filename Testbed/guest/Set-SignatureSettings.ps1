#Requires -Version 5.1
<#
.SYNOPSIS
    Before Outlook starts on an Outlook test guest, makes the signature setting an administrator deploys - a Group
    Policy naming every account's default signature (Q200 (a)), or a default in Outlook's general mail settings
    (Q201) - with the test signatures it names and the hub account's own defaults beside it. The live runner's
    -SignatureSettings step; the guest's checkpoint restore undoes all of it.

.DESCRIPTION
    RUN ON THE GUEST, ELEVATED, WITH OUTLOOK NOT RUNNING: session 0 over PowerShell Direct, where
    Testbed/host/Invoke-LiveTierOnGuest.ps1 -SignatureSettings runs it - after the frozen-clock guard, before it
    starts Outlook. Windows PowerShell 5.1: no ternary, no `??`.

    WHY A STEP BEFORE OUTLOOK STARTS (decided by the maintainer 2026-11-18, Q200 (a) and Q201). Outlook reads an
    administrator's policy and its accounts' own signature defaults when it starts: values written while it ran were
    not seen (G1 and G2 of McpServer/OutlookAI.McpServer.Tests/T2/LiveSignatureSettingsTests.cs, 2026-11-17). The
    live tier runs NOT elevated, and HKCU\Software\Policies is read-only there - its ACL gives the user ReadKey. So
    the runner makes the setting here, elevated, and the checkpoint it restores at the end of the run takes every
    part of it away again. The live test LiveSignatureSettingsTests.SetBeforeOutlookStarted_... then reads what
    Outlook itself inserts in a new mail, a reply and a forward, and holds OutlookAI to the same.

    THE SCENARIOS (-Scenario), each a plan Get-SignatureSettingsPlan returns:
      GroupPolicy         Q200. HKCU\Software\Policies\Microsoft\Office\<major>\Common\MailSettings: NewSignature
                          and ReplySignature, naming two test signatures.
      Preference          Q201. HKCU\Software\Microsoft\Office\<major>\Common\MailSettings: the same two values,
                          there not a policy but a deployed setting.
      PreferenceNoOwnNew  Q201 again, with the hub's own default for replies instead of new mail (decided
                          2026-11-18): the one case the first two leave open, a deployed new-mail value for an
                          account with no new-mail default of its own.
      NoSetting           The control: the hub's own defaults for new mail and for replies, their test
                          signatures, and no setting at all - what Outlook's Signatures dialog and its compose
                          windows do without one (Q156 (d), 2026-11-18).
    The hub account gets ONE default of its own, a third test signature: for new mail under Preference, for replies
    and forwards under GroupPolicy and PreferenceNoOwnNew. So each run ranks the setting against the account's own
    default for one scope, and shows for the other what the setting does for an account that has none of its own.
    (The first policy runs, 2026-11-18, gave the hub its own default for both: Outlook put the account's own in, not
    the policy's.)

    THE VALUES' KIND (-ValueKind). ExpandString, REG_EXPAND_SZ, by default: what signature vendors tell
    administrators to write and what Outlook was reported to honour where a plain string stopped working (Microsoft
    Q&A, 2023-2024). String, REG_SZ, measures the other. On 2026-11-18 Outlook LTSC 2024 put an account's own default
    in a new mail, a reply and a forward, not the policy's, with either kind (Docs/live-tier-on-the-vm.md section 4.14).

    WHAT -Execute DOES, in this order:
      1. Refuses - nothing written - unless the session is 'vmadmin' on a computer named OAI-..., the guest is not
         the Exchange VM (OAI-EXCHANGE carries its state from run to run, so nothing would undo this), the session
         is elevated, OUTLOOK.EXE is not running, no value of the plan is set and no signature of the plan's names
         is in the signature folder: it changes nothing a person or an earlier run made.
      2. Through the shipped MCP server over stdio (McpStdioClient.ps1) - the tested route mailbox safety rule 1
         names, which Set-AccountSignature.ps1 takes too: list_signatures, for the folder, the profile and the
         account; manage_signature create for each signature of the plan, its HTML holding its own name, so a mail
         shows which went in; then manage_signature set_default for the hub account's own defaults. All before
         step 3: manage_signature refuses a default change while a policy names one.
      3. Writes the plan's two registry values.
      4. Verifies through list_signatures and the registry: VERDICT: SET, exit 0, or VERDICT: NOT-SET with what is
         not as planned, exit 1.
    -Verify does step 4 alone, read-only. Neither: a dry run - the plan, and nothing touched.
    -SelfTest: pure - the plans, the refusal, the verdict, and the stdio client's message shapes. No server, no
    registry, no file.

    EXIT CODES: 0 SET, or the dry run; 1 NOT-SET; 4 REFUSED - step 1, or the request itself.

.PARAMETER Scenario
    GroupPolicy (Q200), Preference (Q201), PreferenceNoOwnNew (Q201, the hub's own default for replies only) or
    NoSetting (the control: the hub's own defaults, no setting).

.PARAMETER Account
    The hub account's address, whose own defaults the plan sets. Default: the hub the staged live-test settings
    name (testHubStoreDisplayName - on the Outlook guests the hub account's address).

.PARAMETER OutlookProfile
    The profile the hub account is in: the tier profile the live tier runs on.

.PARAMETER ValueKind
    ExpandString (REG_EXPAND_SZ, the default) or String (REG_SZ): how the two registry values are written.

.PARAMETER OfficeVersion
    The Office major whose keys are written: 16.0, the Outlook guests' Office LTSC 2024. The script refuses when
    Outlook keeps nothing under it.

.EXAMPLE
    # On the guest, elevated, Outlook closed - what the runner does:
    & 'C:\OutlookAI-Q5\Set-SignatureSettings.ps1' -Scenario GroupPolicy -Execute
    # Anywhere, pure:
    powershell.exe -NoProfile -File Testbed\guest\Set-SignatureSettings.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [ValidateSet('GroupPolicy', 'Preference', 'PreferenceNoOwnNew', 'NoSetting')] [string] $Scenario,
    [ValidateSet('ExpandString', 'String')] [string] $ValueKind = 'ExpandString',
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $SelfTest,
    [string]   $Account,
    [string]   $OutlookProfile = 'OutlookAI-Tier',
    [string]   $OfficeVersion = '16.0',
    [string]   $SettingsPath,
    [string]   $ServerExe = 'C:\OutlookAI-Q5\server\OutlookAI.McpServer.exe',
    [int]      $ReplyTimeoutSeconds = 300,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ScriptHelpers.ps1')
. (Join-Path $PSScriptRoot 'OutlookMapiInterop.ps1')
. (Join-Path $PSScriptRoot 'McpStdioClient.ps1')
. (Join-Path $PSScriptRoot 'StagedLiveTestSettings.ps1')

$TestPrefix = 'OutlookAI-McpTest-'
$ExchangeComputerName = 'OAI-EXCHANGE'

# =============================================================================================
# PURE DECISIONS. -SelfTest pins them.
# =============================================================================================

# What a scenario sets: the registry key (under HKCU) and its two values, the hub's own defaults, and every test
# signature named, in the order they are made.
function Get-SignatureSettingsPlan {
    param([string] $Scenario, [string] $OfficeVersion, [string] $ValueKind = 'ExpandString')
    if ($ValueKind -cne 'ExpandString' -and $ValueKind -cne 'String') { throw "'$ValueKind' is no value kind: ExpandString or String." }
    if ($Scenario -eq 'GroupPolicy') {
        $settingNew = $TestPrefix + 'Q200-Forced-New'
        $settingReply = $TestPrefix + 'Q200-Forced-Reply'
        $ownNew = $null
        $ownReply = $TestPrefix + 'Q200-Own-Reply'
        $key = "Software\Policies\Microsoft\Office\$OfficeVersion\Common\MailSettings"
        $token = 'groupPolicy'
    }
    elseif ($Scenario -eq 'Preference' -or $Scenario -eq 'PreferenceNoOwnNew') {
        $settingNew = $TestPrefix + 'Q201-Deployed-New'
        $settingReply = $TestPrefix + 'Q201-Deployed-Reply'
        $ownNew = $TestPrefix + 'Q201-Own-New'
        $ownReply = $null
        if ($Scenario -eq 'PreferenceNoOwnNew') {
            $ownNew = $null
            $ownReply = $TestPrefix + 'Q201-Own-Reply'
        }
        $key = "Software\Microsoft\Office\$OfficeVersion\Common\MailSettings"
        $token = 'preference'
    }
    elseif ($Scenario -eq 'NoSetting') {
        # The control writes no value: the deployed setting's key is only read, to be sure it holds none.
        $settingNew = $null
        $settingReply = $null
        $ownNew = $TestPrefix + 'Q156-Own-New'
        $ownReply = $TestPrefix + 'Q156-Own-Reply'
        $key = "Software\Microsoft\Office\$OfficeVersion\Common\MailSettings"
        $token = $null
    }
    else { throw "'$Scenario' is no scenario: GroupPolicy, Preference, PreferenceNoOwnNew or NoSetting." }
    $signatures = @(@($settingNew, $settingReply, $ownNew, $ownReply) | Where-Object { $_ })
    return [pscustomobject]@{
        Scenario      = $Scenario
        Source        = $token
        Key           = $key
        ValueKind     = $ValueKind
        WritesSetting = ($null -ne $settingNew)
        SettingNew    = $settingNew
        SettingReply  = $settingReply
        OwnNew        = $ownNew
        OwnReply      = $ownReply
        Signatures    = $signatures
    }
}

# The HTML a test signature of the plan is made from: its own name, so the mail it goes into says which it is.
function Get-SignatureHtml {
    param([string] $Name)
    return "<p>OutlookAI signature step</p><p>$Name</p>"
}

# Why -Execute may not run here, or '' when it may. $Observed: User, Computer, Elevated, OutlookRunning,
# ValuesSet (names of the plan's registry values already there), SignaturesThere (the plan's names already in the
# folder), Profiles (the profiles list_signatures names), AccountsOfProfile (the accounts of -OutlookProfile).
function Get-StepRefusal {
    param([object] $Plan, [hashtable] $Observed, [string] $Account, [string] $Profile)
    if (@($ExpectedUser | Where-Object { $_ -eq $Observed.User }).Count -eq 0) { return "the session is '$($Observed.User)', not one of: $($ExpectedUser -join ', ')" }
    if (-not ([string]$Observed.Computer).StartsWith($ExpectedComputerNamePrefix, [System.StringComparison]::OrdinalIgnoreCase)) { return "'$($Observed.Computer)' is not a test guest (its name must start with '$ExpectedComputerNamePrefix')" }
    if ([string]::Equals([string]$Observed.Computer, $ExchangeComputerName, [System.StringComparison]::OrdinalIgnoreCase)) { return "this is the Exchange VM, whose state is carried from run to run - no checkpoint restore would take the setting away" }
    if (-not $Observed.Elevated) { return 'the session is not elevated - a policy key is written by an administrator (session 0 over PowerShell Direct is one)' }
    if ($Observed.OutlookRunning) { return 'Outlook is running - the setting is made before Outlook starts, which reads it then' }
    if (@($Observed.ValuesSet).Count -gt 0) { return "HKCU\$($Plan.Key) already holds $(@($Observed.ValuesSet) -join ' and ') - this changes nothing it did not make" }
    if (@($Observed.SignaturesThere).Count -gt 0) { return "the signature folder already holds $(@($Observed.SignaturesThere) -join ', ') - this changes nothing it did not make" }
    if (-not $Account) { return 'no account: pass -Account, or stage the live-test settings that name the hub' }
    if (@($Observed.Profiles | Where-Object { [string]::Equals([string]$_, $Profile, [System.StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) { return "there is no Outlook profile '$Profile'" }
    if (@($Observed.AccountsOfProfile | Where-Object { [string]::Equals([string]$_, $Account, [System.StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) { return "profile '$Profile' has no mail account '$Account'" }
    return ''
}

# SET when every part of the plan is in place, else NOT-SET and what is not. $Observed: Values (registry value name
# -> its data, or $null), Kinds (registry value name -> its kind, or $null), Policy (list_signatures' policy:
# newMessage, newMessageFrom, replyForward, replyForwardFrom), Own (the account's row: newMessage, replyForward),
# Signatures (the names list_signatures lists).
function Get-SettingsVerdict {
    param([object] $Plan, [hashtable] $Observed)
    $missing = New-Object System.Collections.Generic.List[string]
    $policy = $Observed.Policy
    if (-not $Plan.WritesSetting) {
        # The control: no value of either name, and nothing OutlookAI reads as a setting.
        foreach ($name in @('NewSignature', 'ReplySignature')) {
            if ($null -ne $Observed.Values[$name]) { $missing.Add("HKCU\$($Plan.Key)\$name holds '$($Observed.Values[$name])', where the control has none") }
        }
        if ($null -ne $policy -and ($policy.newMessage -or $policy.replyForward)) { $missing.Add("OutlookAI reads a setting ('$($policy.newMessage)', '$($policy.replyForward)') where the control has none") }
    }
    foreach ($pair in @(@('NewSignature', $Plan.SettingNew), @('ReplySignature', $Plan.SettingReply))) {
        if (-not $Plan.WritesSetting) { break }
        $data = $Observed.Values[$pair[0]]
        if ([string]$data -cne $pair[1]) { $missing.Add("HKCU\$($Plan.Key)\$($pair[0]) is '$data', not '$($pair[1])'") }
        $kind = [string]$Observed.Kinds[$pair[0]]
        if ($null -ne $data -and $kind -cne $Plan.ValueKind) { $missing.Add("HKCU\$($Plan.Key)\$($pair[0]) is a $kind value, not $($Plan.ValueKind)") }
    }
    if (-not $Plan.WritesSetting) { }
    elseif ($null -eq $policy) { $missing.Add('list_signatures reported no policy or deployed default') }
    else {
        if ([string]$policy.newMessage -cne $Plan.SettingNew -or [string]$policy.newMessageFrom -cne $Plan.Source) { $missing.Add("OutlookAI reads the new-mail default as '$($policy.newMessage)' from '$($policy.newMessageFrom)', not '$($Plan.SettingNew)' from '$($Plan.Source)'") }
        if ([string]$policy.replyForward -cne $Plan.SettingReply -or [string]$policy.replyForwardFrom -cne $Plan.Source) { $missing.Add("OutlookAI reads the reply default as '$($policy.replyForward)' from '$($policy.replyForwardFrom)', not '$($Plan.SettingReply)' from '$($Plan.Source)'") }
    }
    $own = $Observed.Own
    if ($null -eq $own) { $missing.Add("the account's row is not in the profile list_signatures reports") }
    else {
        if ([string]$own.newMessage -cne [string]$Plan.OwnNew) { $missing.Add("the account's own new-mail default is '$($own.newMessage)', not '$($Plan.OwnNew)'") }
        $expectedReply = [string]$Plan.OwnReply
        if ([string]$own.replyForward -cne $expectedReply) { $missing.Add("the account's own reply default is '$($own.replyForward)', not '$expectedReply'") }
    }
    foreach ($name in $Plan.Signatures) {
        if (@($Observed.Signatures | Where-Object { [string]$_ -ceq $name }).Count -eq 0) { $missing.Add("signature '$name' is not listed") }
    }
    $verdict = 'SET'
    if ($missing.Count -gt 0) { $verdict = 'NOT-SET' }
    return [pscustomobject]@{ Verdict = $verdict; Missing = @($missing) }
}

function Invoke-SelfTest {
    $script:stFailures = New-Object System.Collections.Generic.List[string]
    $script:stChecks = 0
    function Check([string] $Name, $Expected, $Actual) {
        $script:stChecks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = $Expected -join '|' }
        if ($Actual -is [System.Array]) { $a = $Actual -join '|' }
        if ($e -cne $a) { $script:stFailures.Add("$Name - expected '$e', got '$a'") }
    }
    Write-Host "Set-SignatureSettings.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Pure: no server, no registry, no file."

    $g = Get-SignatureSettingsPlan 'GroupPolicy' '16.0'
    Check 'the policy is Group Policy''s MailSettings key' 'Software\Policies\Microsoft\Office\16.0\Common\MailSettings' $g.Key
    Check 'and list_signatures names its source groupPolicy' 'groupPolicy' $g.Source
    Check 'it names a new-mail and a reply default, and the hub has its own reply default' 'OutlookAI-McpTest-Q200-Forced-New|OutlookAI-McpTest-Q200-Forced-Reply|OutlookAI-McpTest-Q200-Own-Reply' $g.Signatures
    Check 'and no new-mail default of its own' 'OutlookAI-McpTest-Q200-Own-Reply|True' @($g.OwnReply, ($null -eq $g.OwnNew))
    $p = Get-SignatureSettingsPlan 'Preference' '16.0'
    Check 'the deployed default is in Outlook''s own MailSettings, not a policy' 'Software\Microsoft\Office\16.0\Common\MailSettings' $p.Key
    Check 'and list_signatures names its source preference' 'preference' $p.Source
    Check 'the hub has its own new-mail default only' 'OutlookAI-McpTest-Q201-Own-New|True' @($p.OwnNew, ($null -eq $p.OwnReply))
    Check 'three signatures' 3 @($p.Signatures).Count
    $n = Get-SignatureSettingsPlan 'PreferenceNoOwnNew' '16.0'
    Check 'the third scenario is the deployed setting too' 'Software\Microsoft\Office\16.0\Common\MailSettings|preference' @($n.Key, $n.Source)
    Check 'and the hub has its own reply default only' 'OutlookAI-McpTest-Q201-Deployed-New|OutlookAI-McpTest-Q201-Deployed-Reply|OutlookAI-McpTest-Q201-Own-Reply|True' (@($n.Signatures) + @($null -eq $n.OwnNew))
    $c = Get-SignatureSettingsPlan 'NoSetting' '16.0'
    Check 'the control writes no setting, and gives the hub both own defaults' 'False|OutlookAI-McpTest-Q156-Own-New|OutlookAI-McpTest-Q156-Own-Reply' (@($c.WritesSetting) + @($c.Signatures))
    Check 'the other three write theirs' 'True|True|True' @($g.WritesSetting, $p.WritesSetting, $n.WritesSetting)
    foreach ($plan in @($g, $p, $n, $c)) {
        foreach ($name in $plan.Signatures) {
            Check "$name is a test signature" $true $name.StartsWith($TestPrefix, [System.StringComparison]::Ordinal)
            Check "$name is in no other name of its plan" 1 @($plan.Signatures | Where-Object { $_.Contains($name) }).Count
        }
    }
    Check 'the Office major is the one asked for' 'Software\Policies\Microsoft\Office\17.0\Common\MailSettings' (Get-SignatureSettingsPlan 'GroupPolicy' '17.0').Key
    Check 'the values are expandable strings unless asked otherwise' 'ExpandString|String' @($g.ValueKind, (Get-SignatureSettingsPlan 'Preference' '16.0' 'String').ValueKind)
    $threw = $false; try { [void](Get-SignatureSettingsPlan 'GroupPolicy' '16.0' 'Binary') } catch { $threw = $true }
    Check 'an unknown value kind is refused' $true $threw
    $threw = $false; try { [void](Get-SignatureSettingsPlan 'CloudPolicy' '16.0') } catch { $threw = $true }
    Check 'an unknown scenario is refused' $true $threw
    Check 'a signature''s HTML holds its own name' $true ((Get-SignatureHtml 'OutlookAI-McpTest-X').Contains('<p>OutlookAI-McpTest-X</p>'))

    $ok = @{ User = 'vmadmin'; Computer = 'OAI-UNINDEXED'; Elevated = $true; OutlookRunning = $false; ValuesSet = @(); SignaturesThere = @(); Profiles = @('OutlookAI-Tier', 'CorpusProfile'); AccountsOfProfile = @('tier@vm.invalid', 'identity@vm.invalid') }
    Check 'a closed Outlook on a test guest, elevated, may be set' '' (Get-StepRefusal $g $ok 'tier@vm.invalid' 'OutlookAI-Tier')
    $cases = [ordered]@{
        'another user'          = @{ User = 'someone' }
        'another machine'       = @{ Computer = 'WORKSTATION' }
        'the Exchange VM'       = @{ Computer = 'oai-exchange' }
        'not elevated'          = @{ Elevated = $false }
        'Outlook running'       = @{ OutlookRunning = $true }
        'a value already there' = @{ ValuesSet = @('NewSignature') }
        'a signature there'     = @{ SignaturesThere = @('OutlookAI-McpTest-Q200-Own-New') }
        'no such profile'       = @{ Profiles = @('Outlook') }
        'no such account'       = @{ AccountsOfProfile = @('identity@vm.invalid') }
    }
    foreach ($case in $cases.Keys) {
        $o = @{}
        foreach ($k in $ok.Keys) { $o[$k] = $ok[$k] }
        foreach ($k in $cases[$case].Keys) { $o[$k] = $cases[$case][$k] }
        Check "refused: $case" $true ((Get-StepRefusal $g $o 'tier@vm.invalid' 'OutlookAI-Tier') -ne '')
    }
    Check 'refused: no account' $true ((Get-StepRefusal $g $ok '' 'OutlookAI-Tier') -ne '')
    Check 'the profile and the account are compared ignoring case' '' (Get-StepRefusal $g $ok 'TIER@vm.invalid' 'outlookai-tier')

    $set = @{
        Values     = @{ NewSignature = $g.SettingNew; ReplySignature = $g.SettingReply }
        Kinds      = @{ NewSignature = 'ExpandString'; ReplySignature = 'ExpandString' }
        Policy     = [pscustomobject]@{ newMessage = $g.SettingNew; newMessageFrom = 'groupPolicy'; replyForward = $g.SettingReply; replyForwardFrom = 'groupPolicy' }
        Own        = [pscustomobject]@{ newMessage = $null; replyForward = $g.OwnReply }
        Signatures = @($g.Signatures) + @('Identity')
    }
    $v = Get-SettingsVerdict $g $set
    Check 'everything in place is SET' 'SET|0' @($v.Verdict, @($v.Missing).Count)
    $wrongSource = @{}
    foreach ($k in $set.Keys) { $wrongSource[$k] = $set[$k] }
    $wrongSource['Policy'] = [pscustomobject]@{ newMessage = $g.SettingNew; newMessageFrom = 'preference'; replyForward = $g.SettingReply; replyForwardFrom = 'groupPolicy' }
    Check 'a default read from the wrong source is NOT-SET' 'NOT-SET' (Get-SettingsVerdict $g $wrongSource).Verdict
    $wrongKind = @{}
    foreach ($k in $set.Keys) { $wrongKind[$k] = $set[$k] }
    $wrongKind['Kinds'] = @{ NewSignature = 'String'; ReplySignature = 'ExpandString' }
    $v = Get-SettingsVerdict $g $wrongKind
    Check 'a value of the other kind is NOT-SET, and named' $true ($v.Verdict -eq 'NOT-SET' -and (@($v.Missing)[0]).Contains('a String value, not ExpandString'))
    $noReply = @{}
    foreach ($k in $set.Keys) { $noReply[$k] = $set[$k] }
    $noReply['Values'] = @{ NewSignature = $g.SettingNew; ReplySignature = $null }
    $v = Get-SettingsVerdict $g $noReply
    Check 'a value not written is NOT-SET, and named' $true ($v.Verdict -eq 'NOT-SET' -and (@($v.Missing)[0]).Contains('ReplySignature'))
    $noSignature = @{}
    foreach ($k in $set.Keys) { $noSignature[$k] = $set[$k] }
    $noSignature['Signatures'] = @($g.SettingNew)
    Check 'a signature not listed is NOT-SET' 2 @((Get-SettingsVerdict $g $noSignature).Missing).Count
    $pSet = @{
        Values     = @{ NewSignature = $p.SettingNew; ReplySignature = $p.SettingReply }
        Kinds      = @{ NewSignature = 'ExpandString'; ReplySignature = 'ExpandString' }
        Policy     = [pscustomobject]@{ newMessage = $p.SettingNew; newMessageFrom = 'preference'; replyForward = $p.SettingReply; replyForwardFrom = 'preference' }
        Own        = [pscustomobject]@{ newMessage = $p.OwnNew; replyForward = $null }
        Signatures = @($p.Signatures)
    }
    Check 'the deployed setting with no own reply default is SET' 'SET' (Get-SettingsVerdict $p $pSet).Verdict
    $gOwnReply = @{}
    foreach ($k in $set.Keys) { $gOwnReply[$k] = $set[$k] }
    $gOwnReply['Own'] = [pscustomobject]@{ newMessage = 'Identity'; replyForward = $g.OwnReply }
    Check 'under the policy too, an own default the plan did not set is NOT-SET' 'NOT-SET' (Get-SettingsVerdict $g $gOwnReply).Verdict
    $pSet['Own'] = [pscustomobject]@{ newMessage = $p.OwnNew; replyForward = 'Identity' }
    Check 'an own reply default the plan did not set is NOT-SET' 'NOT-SET' (Get-SettingsVerdict $p $pSet).Verdict
    $nSet = @{
        Values     = @{ NewSignature = $n.SettingNew; ReplySignature = $n.SettingReply }
        Kinds      = @{ NewSignature = 'ExpandString'; ReplySignature = 'ExpandString' }
        Policy     = [pscustomobject]@{ newMessage = $n.SettingNew; newMessageFrom = 'preference'; replyForward = $n.SettingReply; replyForwardFrom = 'preference' }
        Own        = [pscustomobject]@{ newMessage = $null; replyForward = $n.OwnReply }
        Signatures = @($n.Signatures)
    }
    Check 'the deployed setting with no own new-mail default is SET' 'SET' (Get-SettingsVerdict $n $nSet).Verdict
    $nSet['Own'] = [pscustomobject]@{ newMessage = $p.OwnNew; replyForward = $n.OwnReply }
    Check 'and NOT-SET when the hub has a new-mail default of its own' 'NOT-SET' (Get-SettingsVerdict $n $nSet).Verdict
    $cSet = @{
        Values     = @{ NewSignature = $null; ReplySignature = $null }
        Kinds      = @{ NewSignature = $null; ReplySignature = $null }
        Policy     = [pscustomobject]@{ newMessage = $null; newMessageFrom = $null; replyForward = $null; replyForwardFrom = $null }
        Own        = [pscustomobject]@{ newMessage = $c.OwnNew; replyForward = $c.OwnReply }
        Signatures = @($c.Signatures)
    }
    Check 'the control with both own defaults and no setting is SET' 'SET' (Get-SettingsVerdict $c $cSet).Verdict
    $cSet['Values'] = @{ NewSignature = 'Somebody'; ReplySignature = $null }
    Check 'and NOT-SET where a setting is there after all' 'NOT-SET' (Get-SettingsVerdict $c $cSet).Verdict
    Check 'no policy at all is NOT-SET' 'NOT-SET' (Get-SettingsVerdict $g @{ Values = @{}; Kinds = @{}; Policy = $null; Own = $null; Signatures = @() }).Verdict

    # The stdio client's pure parts (McpStdioClient.ps1).
    Check 'a request carries its id and params' '{"jsonrpc":"2.0","method":"tools/call","id":7,"params":{"name":"list_signatures"}}' (ConvertTo-McpMessage -Method 'tools/call' -Params ([ordered]@{ name = 'list_signatures' }) -Id 7)
    Check 'a notification carries no id' '{"jsonrpc":"2.0","method":"notifications/initialized"}' (ConvertTo-McpMessage -Method 'notifications/initialized' -Notification)
    Check 'the reply to the request is found' 'ok' (Get-McpReplyFor '{"jsonrpc":"2.0","id":7,"result":{"v":"ok"}}' 7).result.v
    Check 'another request''s reply is skipped' $true ($null -eq (Get-McpReplyFor '{"jsonrpc":"2.0","id":8,"result":{}}' 7))
    Check 'a notification is skipped' $true ($null -eq (Get-McpReplyFor '{"jsonrpc":"2.0","method":"notifications/message","params":{}}' 7))
    Check 'a line that is not JSON is skipped' $true ($null -eq (Get-McpReplyFor 'info: starting' 7))
    $answer = Get-McpToolResult ('{"jsonrpc":"2.0","id":3,"result":{"content":[{"type":"text","text":"{\"name\":\"x\"}"}]}}' | ConvertFrom-Json) 'manage_signature'
    Check 'a tool''s answer is its text, parsed' 'x' $answer.name
    $threw = $false; try { [void](Get-McpToolResult ('{"jsonrpc":"2.0","id":3,"result":{"isError":true,"content":[{"type":"text","text":"{\"error\":{\"type\":\"InvalidArgument\"}}"}]}}' | ConvertFrom-Json) 'manage_signature') } catch { $threw = $_.Exception.Message.Contains('manage_signature refused') }
    Check 'a tool''s refusal throws, naming the tool' $true $threw
    $threw = $false; try { [void](Get-McpToolResult ('{"jsonrpc":"2.0","id":3,"error":{"code":-32602,"message":"bad"}}' | ConvertFrom-Json) 'list_signatures') } catch { $threw = $_.Exception.Message.Contains('list_signatures failed') }
    Check 'a JSON-RPC error throws' $true $threw
    Check 'an error answer to the request is its reply too' '-32601' ([string](Get-McpReplyFor '{"jsonrpc":"2.0","id":3,"error":{"code":-32601}}' 3).error.code)
    Check 'a line that carries the id but neither a result nor an error is no reply' $true ($null -eq (Get-McpReplyFor '{"jsonrpc":"2.0","id":7,"method":"sampling/createMessage","params":{}}' 7))

    # The client against a stand-in server: a Windows PowerShell that answers each request on a line of its own,
    # after a notification that carries the request's id in its params, and takes 4 s over one called 'slow'. For
    # what no pure check can show: the reads, a silence, the read a silence leaves waiting, and the end (Q154, S7).
    $standIn = Join-Path ([System.IO.Path]::GetTempPath()) ('oai-mcp-standin-' + [guid]::NewGuid().ToString('N') + '.ps1')
    $standInText = @'
$reader = [Console]::In
[Console]::Error.WriteLine('stand-in: started')
while ($null -ne ($line = $reader.ReadLine())) {
    $m = $line | ConvertFrom-Json
    if (-not ($m.PSObject.Properties.Name -contains 'id')) { continue }
    if ($m.method -eq 'slow') { Start-Sleep -Seconds 4 }
    [Console]::Out.WriteLine('{"jsonrpc":"2.0","method":"notifications/message","params":{"id":' + $m.id + '}}')
    [Console]::Out.WriteLine('{"jsonrpc":"2.0","id":' + $m.id + ',"result":{"method":"' + $m.method + '"}}')
    [Console]::Out.Flush()
}
[Console]::Error.WriteLine('stand-in: input closed')
'@
    [System.IO.File]::WriteAllText($standIn, $standInText)
    $standInShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $s = $null
    try {
        $s = Start-McpStdioServer -ServerExe $standInShell -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $standIn + '"') -TimeoutSeconds 30
        Check 'the client starts a server and initializes it: request 1, then the initialized notification' 2 $s.NextId
        Check 'a request is answered by its own reply, past a notification that carries its id' 'tools/list' ([string](Invoke-McpRequest -Session $s -Method 'tools/list').result.method)
        $slowId = Send-McpRequest -Session $s -Method 'slow'
        Check 'a request sent now is numbered next, for a reply read later' 3 $slowId
        Check 'a reply that has not come within its time is $null with -NullWhenNone' $true ($null -eq (Read-McpReply -Session $s -Id $slowId -TimeoutSeconds 1 -NullWhenNone))
        Check 'and the read still waiting is kept for the next call' $true ($null -ne $s.PendingRead)
        Check 'which takes it up: the late reply is skipped and the next request''s found' 'ping' ([string](Invoke-McpRequest -Session $s -Method 'ping').result.method)
        Check 'and the line a reply came in is kept as the server wrote it' $true ([string]$s.LastReplyLine).Contains('"method":"ping"')
        $threw = ''
        try { [void](Read-McpReply -Session $s -Id 999 -TimeoutSeconds 1) } catch { $threw = $_.Exception.Message }
        Check 'without -NullWhenNone a silence throws, naming the request' $true ($threw.Contains('request 999'))
        $stderr = Stop-McpStdioServer -Session $s -WaitMilliseconds 15000
        Check 'closing its input ends the server, and what it wrote to standard error comes back to the end' 'True|True' ('{0}|{1}' -f $s.Process.HasExited, $stderr.Contains('stand-in: input closed'))
        $s = $null
    }
    finally {
        if ($null -ne $s) { [void](Stop-McpStdioServer -Session $s -WaitMilliseconds 5000 -KillIfStillRunning) }
        Remove-Item -LiteralPath $standIn -Force -ErrorAction SilentlyContinue
    }

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

if (-not $Scenario) { Write-Host 'REFUSED: name the setting with -Scenario GroupPolicy or -Scenario Preference.'; exit 4 }
if (-not $PSBoundParameters.ContainsKey('Account')) {
    $path = $SettingsPath
    if (-not $path) { $path = $StagedSettingsPath }
    $Account = Get-StagedSettingsValue $path 'testHubStoreDisplayName'
}
$plan = Get-SignatureSettingsPlan $Scenario $OfficeVersion $ValueKind
Write-Host "scenario  : $($plan.Scenario) - HKCU\$($plan.Key): NewSignature='$($plan.SettingNew)', ReplySignature='$($plan.SettingReply)', as $($plan.ValueKind)"
Write-Host "own       : $Account in profile '$OutlookProfile': new mail '$(if ($plan.OwnNew) { $plan.OwnNew } else { '(none)' })', replies and forwards '$(if ($plan.OwnReply) { $plan.OwnReply } else { '(none)' })'"
Write-Host "signatures: $($plan.Signatures -join ', ')"

if (-not $Execute -and -not $Verify) {
    Write-Host 'DRY RUN - nothing touched. -Execute makes the setting (elevated, Outlook closed); -Verify reads it back.'
    exit 0
}

# The registry values of the plan as they stand: Values (name -> data, unexpanded, or $null) and Kinds (name -> kind,
# or $null).
function Read-PlanValues {
    $values = @{ NewSignature = $null; ReplySignature = $null }
    $kinds = @{ NewSignature = $null; ReplySignature = $null }
    $path = "HKCU:\$($plan.Key)"
    if (Test-Path -LiteralPath $path) {
        $item = Get-Item -LiteralPath $path
        foreach ($name in @('NewSignature', 'ReplySignature')) {
            $values[$name] = $item.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if ($null -ne $values[$name]) { $kinds[$name] = [string]$item.GetValueKind($name) }
        }
    }
    return @{ Values = $values; Kinds = $kinds }
}

# What list_signatures says the setting is: the policy view, the account's row in the profile, the names listed.
function Read-ListedState {
    param([hashtable] $Session)
    $listing = Invoke-McpTool -Session $Session -Tool 'list_signatures'
    $row = $null
    foreach ($p in @($listing.profiles)) {
        if ([string]::Equals([string]$p.name, $OutlookProfile, [System.StringComparison]::OrdinalIgnoreCase)) {
            foreach ($a in @($p.accounts)) {
                if ([string]::Equals([string]$a.account, $Account, [System.StringComparison]::OrdinalIgnoreCase)) { $row = $a }
            }
        }
    }
    return @{ Listing = $listing; Policy = $listing.policy; Own = $row; Signatures = @($listing.signatures | ForEach-Object { [string]$_.name }) }
}

function Show-Verdict([object] $Verdict) {
    if ($Verdict.Verdict -eq 'SET') {
        Write-Host "VERDICT: SET - $($plan.Scenario): the setting, its signatures and the account's own defaults are in place; Outlook reads them when it starts."
        return 0
    }
    Write-Host 'VERDICT: NOT-SET'
    foreach ($m in $Verdict.Missing) { Write-Host "  - $m" }
    return 1
}

# The guest guard every testbed script that writes calls first (OutlookMapiInterop.ps1); Get-StepRefusal then
# says why -Execute may not run here, with the rest of what it needs.
Assert-TestbedGuest -ExpectedUser $ExpectedUser

$session = $null
$code = 1
try {
    if (-not (Test-Path -LiteralPath $ServerExe)) { Write-Host "REFUSED: the MCP server is not at $ServerExe - the live runner stages it there."; exit 4 }
    if (-not (Test-Path -LiteralPath "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook")) { Write-Host "REFUSED: Outlook keeps nothing under Office $OfficeVersion here - name the major with -OfficeVersion."; exit 4 }
    $session = Start-McpStdioServer -ServerExe $ServerExe -ClientName 'set-signature-settings' -TimeoutSeconds $ReplyTimeoutSeconds
    $before = Read-ListedState $session

    if ($Execute) {
        $values = (Read-PlanValues).Values
        $observed = @{
            User              = $env:USERNAME
            Computer          = $env:COMPUTERNAME
            Elevated          = (Test-IsElevated)
            OutlookRunning    = (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0)
            ValuesSet         = @($values.Keys | Where-Object { $null -ne $values[$_] } | Sort-Object)
            SignaturesThere   = @($plan.Signatures | Where-Object { $before.Signatures -contains $_ })
            Profiles          = @($before.Listing.profiles | ForEach-Object { [string]$_.name })
            AccountsOfProfile = @($before.Listing.profiles | Where-Object { [string]::Equals([string]$_.name, $OutlookProfile, [System.StringComparison]::OrdinalIgnoreCase) } | ForEach-Object { @($_.accounts) } | ForEach-Object { [string]$_.account })
        }
        $refusal = Get-StepRefusal $plan $observed $Account $OutlookProfile
        if ($refusal) { Write-Host "REFUSED: $refusal. Nothing was written."; exit 4 }

        foreach ($name in $plan.Signatures) {
            [void](Invoke-McpTool -Session $session -Tool 'manage_signature' -Arguments @{ action = 'create'; name = $name; body_html = (Get-SignatureHtml $name) })
            Write-Host "made      : $name"
        }
        $defaults = @()
        if ($plan.OwnNew) { $defaults += [pscustomobject]@{ Name = $plan.OwnNew; Scope = 'new' } }
        if ($plan.OwnReply) { $defaults += [pscustomobject]@{ Name = $plan.OwnReply; Scope = 'reply' } }
        foreach ($d in $defaults) {
            [void](Invoke-McpTool -Session $session -Tool 'manage_signature' -Arguments @{ action = 'set_default'; name = $d.Name; default_for = @{ profile = $OutlookProfile; account = $Account; scope = $d.Scope } })
            Write-Host "default   : $Account, $($d.Scope): $($d.Name)"
        }
        if ($plan.WritesSetting) {
            $path = "HKCU:\$($plan.Key)"
            if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force | Out-Null }
            New-ItemProperty -LiteralPath $path -Name 'NewSignature' -PropertyType $plan.ValueKind -Value $plan.SettingNew -Force | Out-Null
            New-ItemProperty -LiteralPath $path -Name 'ReplySignature' -PropertyType $plan.ValueKind -Value $plan.SettingReply -Force | Out-Null
            Write-Host "written   : HKCU\$($plan.Key) NewSignature, ReplySignature, as $($plan.ValueKind)"
        }
        else { Write-Host 'written   : no setting - the control' }
    }

    $after = Read-ListedState $session
    $now = Read-PlanValues
    $code = Show-Verdict (Get-SettingsVerdict $plan @{ Values = $now.Values; Kinds = $now.Kinds; Policy = $after.Policy; Own = $after.Own; Signatures = $after.Signatures })
}
finally {
    [void](Stop-McpStdioServer -Session $session)
}
exit $code
