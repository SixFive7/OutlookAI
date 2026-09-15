#Requires -Version 5.1
<#
    ============================================================================================
    DRAFT. THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    It was written by an agent that was forbidden to touch Outlook, MAPI, a mail profile or the
    profile registry hive, on a workstation holding real mail and delegate mailboxes. Nothing
    here has been run anywhere. It has been verified by PARSING only - the same check
    .github/scripts/check-testbed-references.ps1 applies to every script under Testbed/. Once it
    HAS run on a guest, replace this banner with what it actually did.

.SYNOPSIS
    Creates the testbed TIER profile - one Unicode PST plus one POP3 account on a loopback mail
    sink - by importing a .prf file, and then PROVES whether Outlook honoured it.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    NEVER run this on the maintainer's workstation: it writes the Outlook Setup key and causes
    the next Outlook start to rebuild the mail profile.

    WHY THIS EXISTS. The tier profile needs a POP3/SMTP account pointed at a loopback sink. The
    Outlook object model is read-only for accounts, Extended MAPI can no longer create POP3
    services, and the one off-the-shelf component that can is excluded by the repository's
    Dependencies rule. A .prf file plus the ImportPRF registry value is the remaining free,
    documented, non-interactive route. The full evidence - with a source for every key - is in
    .work/pop3-account-routes.md section A.

    THIS IS AN EXPERIMENT, AND IT IS WRITTEN TO REPORT A FAILURE RATHER THAN HIDE ONE. The route
    turns on three things nobody has verified on a 16.x build, and each has its own assertion:

      1. Does Outlook 16.x process a PRF's INTERNET ACCOUNT sections (3, 5 and 7) at all? Every
         literal POP3 .prf Microsoft ever published is 2000-2007 era, and the Office Customization
         Tool's own page says a legacy .prf imports "provided that the profile defines only MAPI
         services" - which a POP3 account is not. -Verify fails loudly if no account subkey
         appears under the new profile.
      2. Does `[General] DefaultStore=Service1` bind the POP3 account's delivery store? It cannot
         be expressed as a property: PROP_ACCT_DELIVERY_STORE is PT_BINARY and a PRF carries only
         PT_UNICODE / PT_LONG / PT_BOOLEAN. DefaultStore names a SERVICE and leaves the processor
         to resolve it, which is why the gap is not automatically fatal - but Outlook 2010+ mints
         its OWN pst per POP3 account by default. -Verify reports whether 00180102 exists and
         whether a stray .pst appeared under Documents\Outlook Files.
      3. Is the import genuinely silent? ImportPRF is documented as applying at the next Outlook
         start with no prompt, but ONLY while `FirstRun` and `First-Run` are absent from the
         Setup key. This script deletes both, and says so.

    IT DOES NOT START OUTLOOK, AND THAT IS DELIBERATE. The import happens at the next Outlook
    start, and a script that starts Outlook then owns Outlook's lifetime - which this project
    handles carefully and never from ad-hoc code (never taskkill; release COM references BEFORE
    any quit). So the flow is three steps, and the middle one is yours:

        .\New-TierProfile.ps1                 # dry run: prints the plan and changes nothing
        .\New-TierProfile.ps1 -Execute        # writes the .prf and the registry values
        <start Outlook once, let it settle, close it>
        .\New-TierProfile.ps1 -Verify         # reads the profile hive and asserts; writes only its log

    STARTING STATE IT EXPECTS. A guest checkpoint where Office is installed and the tier profile
    does not exist yet. -Execute asserts that state rather than assuming it: it refuses if
    OUTLOOK.EXE is running, if the Office Setup key is missing, or if a profile of the target name
    already exists without -Force. Re-running -Execute against the same checkpoint is safe and
    converges; re-running it against a machine that has already imported once is what -Force is
    for, and the .prf's `OverwriteProfile=Yes` + `BackupProfile=No` are what should stop that
    creating a "Backup Of <name>" profile. Should. Which spelling 16.x honours is untested - if a
    backup profile appears, try BackupProfile=False in the template.

    WHAT IT NEVER DOES: create a COM object, open a store, read or write a mail item, or touch
    any profile other than the one named. It writes exactly four registry values, in two keys,
    both under HKCU, and it names every one of them before it writes it.

.PARAMETER ProfileName
    The MAPI profile to create. Also the name asserted to be the ONLY profile of that name
    afterwards.

.PARAMETER StoreDisplayName
    The PST service's display name. The live suite treats a hub store's display name as an SMTP
    address, so this and -EmailAddress are usually the same string; see
    Docs/live-tier-on-the-vm.md section 2.6 and its open question about `@` in a store name.

.PARAMETER WorkDir
    Where the rendered .prf and the .pst go. Deliberately space-free: the ImportPRF registry
    value takes a raw unquoted path.

.PARAMETER SinkHost
    The loopback sink's address. Used for both POP3 and SMTP.

.PARAMETER Execute
    Actually write. Without it nothing is created or modified and the plan is printed instead.

.PARAMETER Verify
    Read the profile hive and assert the import worked. Touches no profile, no store and no mail
    item; the only thing it writes is its own log at -LogPath. Safe to run repeatedly.

.PARAMETER Force
    Allow -Execute to proceed when a profile of the target name already exists.

.EXAMPLE
    .\New-TierProfile.ps1
    .\New-TierProfile.ps1 -Execute
    .\New-TierProfile.ps1 -Verify
#>
[CmdletBinding()]
param(
    [string] $ProfileName      = 'OutlookAI-Tier',
    [string] $StoreDisplayName = 'tier@vm.invalid',
    [string] $EmailAddress     = 'tier@vm.invalid',
    [string] $AccountName      = 'OutlookAI tier sink',
    [string] $DisplayName      = 'OutlookAI Tier',
    [string] $Pop3User         = 'tier',
    [string] $WorkDir          = 'C:\OutlookAI-Tier',
    [string] $SinkHost         = '127.0.0.1',
    [int]    $Pop3Port         = 110,
    [int]    $SmtpPort         = 25,
    [string] $OfficeVersion    = '16.0',
    [string] $TemplatePath,
    [string] $LogPath          = 'C:\OutlookAI-Tier\new-tier-profile.log',
    [switch] $Execute,
    [switch] $Verify,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Paths and constants. Every registry path this script touches is named here and nowhere else,
# so a reader can see the whole blast radius in one place.
# ---------------------------------------------------------------------------------------------

$setupKey    = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Setup"
$outlookKey  = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook"
$profilesKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'

# The account-manager subkey under a profile. Stable since Outlook 2002.
$acctMgrSubkey = '9375CFF0413111d3B88A00104B2A6676'

# PROP_ACCT_DELIVERY_STORE, PT_BINARY, tag 0x00180102 - the property a .prf cannot express.
$deliveryStoreValueName = '00180102'

$prfPath = Join-Path $WorkDir 'tier-profile.prf'
$pstPath = Join-Path $WorkDir 'tier.pst'

if (-not $TemplatePath) {
    $TemplatePath = Join-Path $PSScriptRoot 'tier-profile.prf'
}

$script:Failures = @()
$script:Lines    = @()

function Say {
    param([string] $Text)
    $script:Lines += $Text
    Write-Host $Text
}

function Pass {
    param([string] $What, [string] $Detail)
    Say ("  OK   {0}{1}" -f $What, $(if ($Detail) { " - $Detail" } else { '' }))
}

function Fail {
    param([string] $What, [string] $Why)
    $script:Failures += "$What : $Why"
    Say ("  FAIL {0} - {1}" -f $What, $Why)
}

function Save-Log {
    # A dry run creates NOTHING, including its own log directory. A "dry run" that leaves a
    # directory behind is a dry run whose promise is already false.
    if (-not ($Execute -or $Verify)) { return }

    $dir = Split-Path -Parent $LogPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Set-Content -LiteralPath $LogPath -Value $script:Lines -Encoding UTF8
    Write-Host ''
    Write-Host "Log: $LogPath"
}

# ---------------------------------------------------------------------------------------------
# Reading the profile hive. Read-only, and it never opens a store: these are registry values,
# not MAPI calls. Account property values are REG_BINARY holding either 8-bit or UTF-16LE text,
# so a search decodes both and reports which matched.
# ---------------------------------------------------------------------------------------------

function Get-ProfileNames {
    if (-not (Test-Path -LiteralPath $profilesKey)) {
        return @()
    }
    return @(Get-ChildItem -LiteralPath $profilesKey -ErrorAction SilentlyContinue |
        ForEach-Object { $_.PSChildName })
}

function ConvertTo-ReadableText {
    param($Data)
    if ($null -eq $Data) { return '' }
    if ($Data -is [string]) { return $Data }

    # Only REG_BINARY gets the two-encoding treatment. A DWORD is a number and rendering it as
    # text would invent a string that is not in the registry - which is exactly the kind of
    # made-up evidence the verify step must not produce.
    if (-not ($Data -is [byte[]] -or $Data -is [System.Array])) {
        return [string] $Data
    }

    $bytes = New-Object 'System.Collections.Generic.List[byte]'
    foreach ($b in $Data) {
        $n = 0
        if (-not [int]::TryParse([string] $b, [ref] $n)) { return [string] $Data }
        if ($n -lt 0 -or $n -gt 255) { return [string] $Data }
        $bytes.Add([byte] $n)
    }
    if ($bytes.Count -eq 0) { return '' }

    $arr     = $bytes.ToArray()
    $ascii   = [System.Text.Encoding]::ASCII.GetString($arr)
    $unicode = [System.Text.Encoding]::Unicode.GetString($arr)
    return ($ascii + '|' + $unicode)
}

function Get-AccountValueDump {
    param([string] $Profile)

    $result = New-Object psobject -Property @{
        AccountKeys        = @()
        Text               = ''
        HasDeliveryStore   = $false
        DeliveryStoreBytes = 0
    }

    $mgrPath = Join-Path (Join-Path $profilesKey $Profile) $acctMgrSubkey
    if (-not (Test-Path -LiteralPath $mgrPath)) {
        return $result
    }

    $texts = @()
    $keys  = @()
    foreach ($child in (Get-ChildItem -LiteralPath $mgrPath -Recurse -ErrorAction SilentlyContinue)) {
        $keys += $child.Name
        $props = $null
        try { $props = Get-ItemProperty -LiteralPath $child.PSPath -ErrorAction Stop } catch { $props = $null }
        if ($null -eq $props) { continue }
        foreach ($p in $props.PSObject.Properties) {
            if ($p.Name -like 'PS*') { continue }
            $texts += ("{0}={1}" -f $p.Name, (ConvertTo-ReadableText $p.Value))
            if ($p.Name -eq $deliveryStoreValueName) {
                $result.HasDeliveryStore = $true
                if ($p.Value -is [array]) { $result.DeliveryStoreBytes = $p.Value.Count }
            }
        }
    }

    $result.AccountKeys = $keys
    $result.Text        = ($texts -join "`n")
    return $result
}

# ---------------------------------------------------------------------------------------------
# VERIFY. Reads only. Every check either passes with evidence or fails with a reason - there is
# no branch that prints success without having looked.
# ---------------------------------------------------------------------------------------------

function Invoke-Verify {
    Say '== Verify: did Outlook honour the .prf? =='
    Say ''

    $profiles = Get-ProfileNames
    Say ("Profiles present: " + $(if ($profiles.Count -gt 0) { $profiles -join ', ' } else { '<none>' }))
    Say ''

    if ($profiles -contains $ProfileName) {
        Pass 'the tier profile exists' $ProfileName
    }
    else {
        Fail 'the tier profile exists' "No profile named '$ProfileName'. Outlook did not process the .prf, or it has not been started since -Execute ran. Check that ImportPRF is still set and that FirstRun/First-Run are absent under $setupKey."
    }

    $backups = @($profiles | Where-Object { $_ -like 'Backup Of*' })
    if ($backups.Count -eq 0) {
        Pass 'no backup profile was created' 'BackupProfile=No held'
    }
    else {
        Fail 'no backup profile was created' ("Found: " + ($backups -join ', ') + ". BackupProfile=No was not honoured - try BackupProfile=False in the template, and delete these before re-running. A repeat import that keeps making these makes the guest non-reproducible.")
    }

    $dump = Get-AccountValueDump -Profile $ProfileName
    if ($dump.AccountKeys.Count -gt 0) {
        Pass 'the account manager holds at least one account' ("{0} subkey(s)" -f $dump.AccountKeys.Count)
    }
    else {
        Fail 'the account manager holds at least one account' "Nothing under $acctMgrSubkey. THIS IS THE ANSWER TO THE OPEN QUESTION: Outlook 16.x did not process the .prf's internet-account sections (3, 5, 7). The PST half may still have worked - check the profile above. Route A is dead on this build; see .work/pop3-account-routes.md section A.8 item 1."
    }

    if ($dump.Text -match [regex]::Escape($SinkHost)) {
        Pass 'the POP3 host reached the profile' $SinkHost
    }
    else {
        Fail 'the POP3 host reached the profile' "'$SinkHost' appears in no account value. The account subkeys exist but the section-5 values did not land - which points at a section-7 mapping problem rather than at the import as a whole."
    }

    # 0x0104 is PROP_ACCT_POP3_PORT. Its presence is what says section 5 reached the account at
    # all; the VALUE is in the dump at the bottom, because a port stored as a DWORD is not text
    # and pretending to match it as text would be a check that passes for the wrong reason.
    if ($dump.Text -match '(?m)^[0-9a-fA-F]{4}0104=') {
        Pass 'a POP3 port property exists' 'tag 0x0104 - read its value in the dump below'
    }
    else {
        Fail 'a POP3 port property exists' "No value named ...0104 under the account. Section 5's port did not land. Outlook may have defaulted to 110 anyway; read the dump before concluding it is broken."
    }

    if ($dump.HasDeliveryStore) {
        Pass 'the account has a delivery store' ("{0} = {1} byte(s) - read the log to see WHICH store" -f $deliveryStoreValueName, $dump.DeliveryStoreBytes)
    }
    else {
        Fail 'the account has a delivery store' "No $deliveryStoreValueName value. [General] DefaultStore=Service1 did not bind a per-account delivery store. The fallback is section A.5 of .work/pop3-account-routes.md: drop DefaultStore, set ForcePSTPath, and discover the minted .pst by account instead of by path."
    }

    if (Test-Path -LiteralPath $pstPath) {
        Pass 'the named PST exists' $pstPath
    }
    else {
        Fail 'the named PST exists' "$pstPath was not created. Outlook creates the file lazily on first use, so this can be a false alarm on a profile that has never been opened - but with the delivery-store check above it is the pair that says whether the PST half worked."
    }

    $strayDir = Join-Path $env:USERPROFILE 'Documents\Outlook Files'
    $stray = @()
    if (Test-Path -LiteralPath $strayDir) {
        $stray = @(Get-ChildItem -LiteralPath $strayDir -Filter '*.pst' -File -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Name })
    }
    if ($stray.Count -eq 0) {
        Pass 'Outlook minted no PST of its own' $strayDir
    }
    else {
        Fail 'Outlook minted no PST of its own' ("Found in $strayDir : " + ($stray -join ', ') + ". This is the documented Outlook 2010+ behaviour - a POP3 account gets its own data file unless told otherwise - and it means DefaultStore did not take. Not necessarily fatal: see the ForcePSTPath fallback in section A.5.")
    }

    Say ''
    Say '-- account value dump (this is the evidence; read it before trusting any line above) --'
    if ($dump.Text) { Say $dump.Text } else { Say '<empty>' }
}

# ---------------------------------------------------------------------------------------------
# PLAN and EXECUTE.
# ---------------------------------------------------------------------------------------------

function Get-RenderedPrf {
    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        throw "PRF template not found: $TemplatePath. It ships beside this script as tier-profile.prf; pass -TemplatePath if it lives elsewhere."
    }
    $text = Get-Content -LiteralPath $TemplatePath -Raw

    $map = @{
        '{{PROFILE_NAME}}'      = $ProfileName
        '{{STORE_DISPLAY_NAME}}' = $StoreDisplayName
        '{{PST_PATH}}'          = $pstPath
        '{{ACCOUNT_NAME}}'      = $AccountName
        '{{POP3_HOST}}'         = $SinkHost
        '{{SMTP_HOST}}'         = $SinkHost
        '{{POP3_USER}}'         = $Pop3User
        '{{EMAIL_ADDRESS}}'     = $EmailAddress
        '{{DISPLAY_NAME}}'      = $DisplayName
        '{{POP3_PORT}}'         = [string] $Pop3Port
        '{{SMTP_PORT}}'         = [string] $SmtpPort
    }
    foreach ($k in $map.Keys) {
        $text = $text.Replace($k, $map[$k])
    }

    $left = [regex]::Matches($text, '\{\{[A-Z0-9_]+\}\}')
    if ($left.Count -gt 0) {
        $names = @($left | ForEach-Object { $_.Value } | Sort-Object -Unique)
        throw ("The template still holds unsubstituted tokens after rendering: " + ($names -join ', ') +
            ". A token Outlook cannot parse is a value written silently wrong, so this refuses rather than shipping it.")
    }

    # ASCII, CRLF, no BOM. Section A.6 item 4: every published sample is plain 8-bit text, and a
    # UTF-8 BOM would prefix the first line.
    $normalised = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    return $normalised
}

function Show-Plan {
    Say '== Plan (nothing below has been done) =='
    Say ''
    Say "  profile name        : $ProfileName"
    Say "  store display name  : $StoreDisplayName"
    Say "  email address       : $EmailAddress"
    Say "  POP3                : $SinkHost`:$Pop3Port  (no SSL, no SPA, no stored password)"
    Say "  SMTP                : $SinkHost`:$SmtpPort  (no SSL, no auth)"
    Say "  PRF template        : $TemplatePath"
    Say "  PRF written to      : $prfPath"
    Say "  PST named as        : $pstPath"
    Say ''
    Say '  Registry it would write:'
    Say "    $setupKey"
    Say "      ImportPRF   REG_SZ   = $prfPath"
    Say '      First-Run   DELETED  (ImportPRF is ignored while it exists)'
    Say '      FirstRun    DELETED  (same)'
    Say ''
    Say '  It writes nothing else, and it does NOT start Outlook.'
    Say ''
    Say '  Then, by hand: start Outlook once, let it settle, close it. Then:'
    Say '    .\New-TierProfile.ps1 -Verify'
}

function Invoke-Execute {
    Say '== Execute =='
    Say ''

    # -- preflight: assert the starting state rather than assuming it -------------------------

    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw ("OUTLOOK.EXE is running (pid " + (($running | ForEach-Object { $_.Id }) -join ', ') +
            "). A .prf is read at startup, so importing into a running Outlook does nothing visible now " +
            "and something surprising later. Close Outlook and re-run. This script will not kill it.")
    }

    if (-not (Test-Path -LiteralPath $outlookKey)) {
        throw ("Not found: $outlookKey. Office $OfficeVersion does not look installed for this user. " +
            "Pass -OfficeVersion if the guest has a different major version.")
    }

    $existing = Get-ProfileNames
    if ($existing -contains $ProfileName -and -not $Force) {
        throw ("A profile named '$ProfileName' already exists. This script expects a checkpoint where it does not. " +
            "Re-run with -Force to import over it - the template's OverwriteProfile=Yes + BackupProfile=No are " +
            "what should make that idempotent, and -Verify will tell you whether they did.")
    }
    Say ("Profiles before: " + $(if ($existing.Count -gt 0) { $existing -join ', ' } else { '<none>' }))

    # -- render and write the .prf ------------------------------------------------------------

    $prfText = Get-RenderedPrf

    if (-not (Test-Path -LiteralPath $WorkDir)) {
        New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
        Say "Created $WorkDir"
    }
    else {
        Say "$WorkDir already exists"
    }

    # Section A.6 item 1: the directories in the path to the personal folders must already exist.
    $pstDir = Split-Path -Parent $pstPath
    if (-not (Test-Path -LiteralPath $pstDir)) {
        New-Item -ItemType Directory -Path $pstDir -Force | Out-Null
        Say "Created $pstDir"
    }

    [System.IO.File]::WriteAllText($prfPath, $prfText, (New-Object System.Text.ASCIIEncoding))
    Say "Wrote $prfPath ($($prfText.Length) chars, ASCII, CRLF)"

    $readBack = [System.IO.File]::ReadAllText($prfPath)
    if ($readBack -ne $prfText) {
        throw "The .prf read back differently from what was written. Refusing to continue."
    }
    if ($readBack -notmatch '(?m)^ProfileName=') {
        throw "The written .prf has no ProfileName= line. The template is not what this script expects."
    }
    Pass 'the .prf reads back byte-identical' $prfPath

    # -- registry -----------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $setupKey)) {
        New-Item -Path $setupKey -Force | Out-Null
        Say "Created $setupKey"
    }

    New-ItemProperty -LiteralPath $setupKey -Name 'ImportPRF' -Value $prfPath -PropertyType String -Force | Out-Null
    Say "Set ImportPRF = $prfPath"

    foreach ($name in @('First-Run', 'FirstRun')) {
        $present = $null
        try { $present = Get-ItemProperty -LiteralPath $setupKey -Name $name -ErrorAction Stop } catch { $present = $null }
        if ($null -ne $present) {
            Remove-ItemProperty -LiteralPath $setupKey -Name $name -Force
            Say "Deleted $name"
        }
        else {
            Say "$name already absent"
        }
    }

    # -- verify what we just wrote, rather than assuming the writes took ----------------------

    $back = Get-ItemProperty -LiteralPath $setupKey
    if ($back.ImportPRF -eq $prfPath) {
        Pass 'ImportPRF reads back' $prfPath
    }
    else {
        Fail 'ImportPRF reads back' ("Expected '$prfPath', found '" + $back.ImportPRF + "'.")
    }

    foreach ($name in @('First-Run', 'FirstRun')) {
        $still = $back.PSObject.Properties | Where-Object { $_.Name -eq $name }
        if ($null -eq $still) {
            Pass "$name is absent" 'ImportPRF will be honoured at the next Outlook start'
        }
        else {
            Fail "$name is absent" "It is still present, and ImportPRF is ignored while it exists - with no diagnostic anywhere."
        }
    }

    Say ''
    Say 'NEXT, BY HAND: start Outlook once, let it finish starting, then close it.'
    Say 'Then run:  .\New-TierProfile.ps1 -Verify'
    Say 'This script does not start Outlook: importing is a startup-time action, and owning'
    Say "Outlook's lifetime from a setup script is how a guest ends up with a zombie OUTLOOK.EXE."
}

# ---------------------------------------------------------------------------------------------
# Entry point.
# ---------------------------------------------------------------------------------------------

Say "New-TierProfile.ps1 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Say ''

if ($Verify -and $Execute) {
    throw 'Pass -Execute or -Verify, not both: verifying in the same run as writing would assert against a profile Outlook has not read yet, and would pass for the wrong reason.'
}

if ($Verify) {
    Invoke-Verify
}
elseif ($Execute) {
    Invoke-Execute
}
else {
    Show-Plan
    Say ''
    Say 'Dry run. Nothing written. Re-run with -Execute.'
}

Say ''
if ($script:Failures.Count -gt 0) {
    Say ("{0} check(s) FAILED:" -f $script:Failures.Count)
    foreach ($f in $script:Failures) { Say "  - $f" }
    Save-Log
    exit 1
}

if ($Verify -or $Execute) {
    Say 'All checks passed.'
}
Save-Log
exit 0
