#Requires -Version 5.1
<#
    ============================================================================================
    NEVER EXECUTED ON A GUEST. WRITTEN 2026-09-24. WHAT HAS RUN IS ON THE HOST, AND IT IS THIS:
    ============================================================================================

      * -SelfTest: 115 assertions, 0 failures, under Windows PowerShell 5.1 and PowerShell 7 - 30
        of them read the source files this script mirrors. Six of its rules were broken on purpose
        in a scratch copy (the DWORD-only bool, the URL slashes, the key comparison, freshness, a
        contract string, the never-ran verdict) and each was caught.
      * Its pure functions, loaded without the main body, against the REAL payload
        Testbed/host/Publish-AddInPayload.ps1 built from fb19ccf: the manifest is well-formed; the
        key read out of the built OutlookAI.vsto equals the one in OutlookAI.cer, the one in the
        application manifest and the one the payload manifest names; and the zip-comment reader
        returned fb19ccf from a real `git archive` zip.

    Every mechanism below was checked against something real before it was written down - the
    VSTO runtime's own IL, the installer's own source, trust entries the runtime itself wrote on
    the maintainer's workstation - and each is labelled with how it is known. None of that is a
    guest run. The things only a guest can settle are listed at the end of -SelfTest's output.

    Replace this banner with what it actually did the first time it runs on a guest, and say which
    of the four verdicts it printed.

.SYNOPSIS
    Puts the OutlookAI add-in - built from a named commit by Testbed/host/Publish-AddInPayload.ps1 -
    onto a testbed guest, trusts it without a prompt, starts Outlook once so the add-in writes its
    tuning state, and then PROVES the exact registry state the live tests read.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    -Execute STARTS OUTLOOK, so it must run in the interactive session, through
    Testbed/guest/Register-InteractiveTask.ps1 - never over PowerShell Direct, which lands in
    session 0 where Outlook cannot finish starting:

        .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\src\Testbed\guest\Install-OutlookAIAddIn.ps1' -Execute"

    -Verify and -SelfTest run anywhere. NEVER on the maintainer's workstation: the guard refuses.

    WHY THIS EXISTS. Two live tests read state only the add-in writes, the first time it runs:

      T3/Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning
          declares Requires=AddInRegistry; asserts outlook_health's tuning.managed is true.
      T2/LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail
          declares only SearchIndex and MultipleStores - it UNDER-DECLARES AddInRegistry - and
          asserts Tuning.Managed, Tuning.Enabled and a non-null Tuning.LastReconcileUtc.

    (T2/LiveUiSearchBackendTests.FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates also
    declares AddInRegistry, but writes the value it tests itself; it does not need the add-in.)

    The script-built guests have no add-in, so both of those FAIL there - not skip. And BOTH GUESTS
    need it: the Phase-7 test declares nothing that keeps it off the unindexed guest, and it asserts
    index.wSearchStartMode == "automatic", which Set-OutlookIndexingDisabled.ps1 deliberately keeps
    true there.

    WHAT THE TESTS READ, EXACTLY. McpServer/OutlookAI.Core/Services/HealthReporting.cs,
    ReadTuningState, over HKCU\Software\OutlookAI\Tuning, value names from
    Services/AddInServerContract.cs:

      Initialized       must be a REG_DWORD, nonzero -> tuning.managed = true
      Enabled           must be a REG_DWORD, nonzero -> tuning.enabled = true
      LastReconcileUtc  must be a REG_SZ             -> tuning.lastReconcileUtc not null

    TYPE MATTERS, not just value: HealthReporting's AsBool accepts a boxed int and nothing else, so
    a REG_QWORD 1 or a REG_SZ "1" reads as not-managed. -Verify mirrors that rule exactly, and
    -SelfTest checks the mirror against the source it mirrors.

    WHAT -Execute DOES, IN ORDER, and every step is idempotent:

      1. REFUSES unless: this is a guest (two-axis guard), elevated, 64-bit, in an INTERACTIVE
         session, and OUTLOOK.EXE is not running - the add-in loads when Outlook STARTS, and the
         proof below needs that start to be this script's.
      2. CHECKS THE PAYLOAD. addin-payload.json from the host build; the installer's SHA-256 must
         match it before anything runs.
      3. THE VSTO RUNTIME, from STAGED media - never a download - pinned by the SHA-256 and length
         .github/workflows/release.yml pins (compared against Microsoft's own download on every
         release, which is why this hash, unlike the SDK's, has a default). Skipped when
         `VSTO Runtime Setup\v4R` already reports this version or newer.
         WHY THIS SCRIPT INSTALLS IT AND THE INSTALLER DOES NOT: Installer.iss runs its
         prerequisite step only `if not WizardSilent`, and a silent install is the only kind an
         unattended guest can run. [READ in Installer.iss.]
      4. THE PRODUCT'S OWN INSTALLER, silently: /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
         /NOCLOSEAPPLICATIONS. It is per-user (PrivilegesRequired=lowest) into
         %LOCALAPPDATA%\OutlookAI\Setup, registers the add-in under
         HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI as `file:///{app}\OutlookAI.vsto|vstolocal`,
         exempts it from Outlook's slow-add-in disabling, and imports its signing certificate into
         the user's TrustedPublisher store - exactly what it does for a user.
      5. TRUST, WRITTEN RATHER THAN CLICKED. See TRUST below.
      6. VSTO_LOGALERTS=1 for this user, so a load failure writes <app>\OutlookAI.vsto.log instead
         of vanishing. [MS-DOC: "Debug Office projects".] Errors are NOT shown in a dialog - that
         is VSTO_SUPPRESSDISPLAYALERTS=0, which this never sets.
      7. STARTS OUTLOOK ONCE, over COM, headless, in a CHILD job under a deadline - so a hang is
         reported rather than inherited, the shape the measured cold-start probe in
         Testbed/MEDIA.md used. It waits for LastReconcileUtc to be written AFTER the start, asks
         Outlook whether the add-in is connected, and calls into the add-in itself
         (COMAddIn.Object.GetRestartNeeded, AddInAutomation.cs) - which only a loaded add-in can
         answer. It releases every reference and NEVER quits or kills Outlook (mailbox-safety rule
         7). Outlook may stay up headless or close by itself once the last reference goes; both
         are graceful and the script says which happened.
      8. VERIFIES, as -Verify does, plus FRESHNESS: the state must have been written by THIS start.

    TRUST, AND WHY IT IS WRITTEN. A VSTO add-in loads silently only if its manifest's signer is a
    trusted publisher CHAINING TO A TRUSTED ROOT, or an inclusion-list entry vouches for it.
    Otherwise the ClickOnce trust prompt appears - and on an unattended guest a prompt is a hang,
    which is what the hand-built guest's CP-05-ADDIN-TRUSTED checkpoint recorded somebody clicking
    through. The installer's TrustedPublisher import alone does not retire that prompt for a
    self-signed certificate (Installer.iss says so, and why it will not touch the Root store). So
    this script writes the entry the prompt itself would have written:

      HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid>
          Url        REG_SZ  file:///C:/Users/.../OutlookAI/Setup/OutlookAI.vsto
          PublicKey  REG_SZ  <RSAKeyValue><Modulus>...</Modulus><Exponent>...</Exponent></RSAKeyValue>

      [MS-DOC] "If the end user grants trust to the solution, an inclusion list entry is created
               that contains a URL and a public key" - Grant trust to Office solutions.
      [MS-SUPPORT] the key path and both value names - Microsoft Japan Office support blog on
               UserInclusionList, which also records that the public API for it is not usable
               from .NET 4 (confirmed: AddInSecurityEntry is internal in the v10 runtime).
      [READ] the v10 runtime's own IL: it stores Url as Uri.ToString(), finds entries by comparing
               Url AS A URI, and compares keys by FromXmlString + ExportCspBlob - so the key must be
               the manifest's signing key, and the XML shape is free.
      [MEASURED] the maintainer's workstation carries an entry of exactly this shape for this very
               installer's install path: forward slashes, `file:///C:/...`, the key as a bare
               <RSAKeyValue>. Who wrote it is not recorded; the trust prompt is the only writer
               this repository knows of for an installed path.

    The key is read out of the INSTALLED deployment manifest - the thing the runtime compares - and
    must agree with the installed OutlookAI.cer and with the payload manifest. Only entries for this
    one URL are ever replaced.

    WHAT IT NEVER DOES. It never kills or quits Outlook, never creates, moves, edits or deletes a
    mail item, never touches a store, a profile or MAPI, and never writes the add-in's own
    Software\OutlookAI keys - the tests read what the ADD-IN wrote, or nothing. It does not touch
    the Windows Search policy or the crawl scope; it READS both before and after, and reports any
    change (see INTERACTION).

    INTERACTION WITH THE INDEX EXCLUSION AND THE CORPORA. The add-in's tuning service
    (Services/OutlookTuningService.cs) writes only under HKCU: Outlook's Search key (four search-box
    preferences), the Cached Mode user and POLICY keys (Exchange-only sync settings), and the PST
    key (a larger PST/OST size cap). Set-OutlookIndexingDisabled.ps1 writes only HKLM: the Windows
    Search PreventIndexingOutlook policy and the crawl-scope rule. The two sets are DISJOINT, none
    of the four search values decides whether Outlook's stores are indexed, and nothing the add-in
    does touches an item or a store. [READ, both sources.] -Execute snapshots the exclusion state
    before and after its Outlook start and says if anything moved; if it did, re-run
    Set-OutlookIndexingDisabled.ps1 -Verify before trusting the guest as unindexed.

    FOUR VERDICTS, and only one exits 0:

      ADDIN-READY           installed, registered (LoadBehavior 3), trusted, runtime present, and
                            the state the tests read is valid - and, under -Execute, written by
                            THIS start. exit 0.
      INSTALLED-NEVER-RAN   everything in place, but the add-in has not written its state yet. A
                            real state, not a fault: run -Execute. exit 2.
      NOT-INSTALLED         no add-in here. exit 3.
      BROKEN                something that should hold does not; every reason is printed. exit 1.

    THE GUARD, TWO AXES, same rule and shape as Testbed/guest/Install-DotnetSdk.ps1: the logged-on
    user must be one of -ExpectedUser AND the computer name must start with
    -ExpectedComputerNamePrefix. It runs first in every path except -SelfTest.

.PARAMETER PayloadRoot
    Where AddIn.zip was expanded: the installer and addin-payload.json.

.PARAMETER VstoRuntimePath
    The staged VSTO runtime redistributable. Never downloaded.

.PARAMETER VstoRuntimeSha256
    Its SHA-256. Defaulted to release.yml's pin - see step 3.

.PARAMETER SuiteSourceRoot
    Where Testbed/host/Publish-LiveTierPayload.ps1's source was expanded. Used to compare the add-in's
    contract files against the suite's; absent is a note, not a failure.

.PARAMETER SuiteSourceZip
    That payload's Source.zip, whose zip comment `git archive` sets to the commit it was built from.

.PARAMETER OfficeVersion
    The Office major whose Outlook hive to read. Detected the way Services/OfficeVersions.cs does
    when omitted.

.PARAMETER Execute
    Install, trust, start Outlook once, verify. Session 1 only.

.PARAMETER Verify
    Read and report. Writes nothing but its log.

.PARAMETER WithOutlook
    With -Verify: also ask a RUNNING Outlook in this session, over COM, whether the add-in is
    connected. Never starts Outlook.

.PARAMETER SelfTest
    The pure decisions against synthetic inputs, plus the contract this script mirrors, read from
    the repository source when it is reachable. No registry, no COM, no guest needed.

.EXAMPLE
    .\Install-OutlookAIAddIn.ps1 -SelfTest
    .\Install-OutlookAIAddIn.ps1
    .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\src\Testbed\guest\Install-OutlookAIAddIn.ps1' -Execute"
    .\Install-OutlookAIAddIn.ps1 -Verify
#>
[CmdletBinding()]
param(
    [string]   $PayloadRoot                = 'C:\OutlookAI-Q5\addin',
    [string]   $VstoRuntimePath            = 'C:\OutlookAI-Q5\media\vstor_redist.exe',
    [string]   $VstoRuntimeSha256          = 'CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18',
    [string]   $SuiteSourceRoot            = 'C:\OutlookAI-Q5\src',
    [string]   $SuiteSourceZip             = 'C:\OutlookAI-Q5\Source.zip',
    [string]   $OfficeVersion,
    [string[]] $ExpectedUser               = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [int]      $InstallTimeoutMinutes      = 15,
    [int]      $FirstRunTimeoutSeconds     = 240,
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $WithOutlook,
    [switch]   $SelfTest,
    [string]   $LogPath                    = 'C:\OutlookAI-Q5\install-addin.log'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Everything this script reads or writes, named in one place, so the blast radius is readable
# without reading the code. WRITES are only the three marked; everything else is read.
# ---------------------------------------------------------------------------------------------
$AddinName               = 'OutlookAI'
$AddinRegistrationKey    = 'Software\Microsoft\Office\Outlook\Addins\OutlookAI'   # the installer writes
$AppKey                  = 'Software\OutlookAI'                                   # InstallDir: the installer
$InstallDirValue         = 'InstallDir'
$InclusionKey            = 'Software\Microsoft\VSTO\Security\Inclusion'           # WRITES: one entry, this add-in's URL only
$EnvironmentKey          = 'Environment'                                          # WRITES: VSTO_LOGALERTS=1
$LogAlertsName           = 'VSTO_LOGALERTS'
$TuningKey               = 'Software\OutlookAI\Tuning'                            # the ADD-IN writes; this reads
$McpKey                  = 'Software\OutlookAI\Mcp'                               # the ADD-IN writes; this reads
$VstoRuntimeKey          = 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R'            # HKLM, 32-bit view, as Installer.iss reads it
$VstoRuntimeKeyV4        = 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4'
$VstoRuntimeBytes        = 41828424
$VstoRuntimeVersion      = '10.0.60917'                                           # what that file installs
$SearchPolicyKey         = 'SOFTWARE\Policies\Microsoft\Windows\Windows Search'  # HKLM; Set-OutlookIndexingDisabled.ps1's
$CrawlRulesKey           = 'SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\WorkingSetRules'
$SupportedOffice         = @('16.0', '17.0', '15.0')                              # Services/OfficeVersions.cs, same order
$ManifestFileName        = 'addin-payload.json'
$StatusAwaitingChoice    = 'awaiting_choice'                                      # McpRegistrationService
$StatusNoClaude          = 'claude_code_not_installed'

# What the live tests read, and which test fails when it is wrong. Printed by -Verify, pinned by
# -SelfTest against HealthReporting.cs and the two tests.
$TestReads = @(
    @{ Name = 'Initialized'; Kind = 'DWord'; Means = 'tuning.managed';
       Fails = 'T3 Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning and T2 LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail' }
    @{ Name = 'Enabled'; Kind = 'DWord'; Means = 'tuning.enabled';
       Fails = 'T2 LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail' }
    @{ Name = 'LastReconcileUtc'; Kind = 'String'; Means = 'tuning.lastReconcileUtc';
       Fails = 'T2 LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail' }
)

$VerdictReady        = 'ADDIN-READY'
$VerdictNeverRan     = 'INSTALLED-NEVER-RAN'
$VerdictNotInstalled = 'NOT-INSTALLED'
$VerdictBroken       = 'BROKEN'

if ($SelfTest) { $LogPath = $null }

function Say([string] $m) {
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m
    Write-Host $line
    if ($LogPath) {
        try {
            $dir = Split-Path -Parent $LogPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }
        catch {
            # A log that cannot be written must not stop the work it was recording.
        }
    }
}

# =============================================================================================
# PURE DECISIONS. Each decides from its arguments alone, so -SelfTest can drive it anywhere.
# =============================================================================================

function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

# The installer's switches. /VERYSILENT and /SUPPRESSMSGBOXES: nothing on screen - a dialog on an
# unattended guest is a hang. /NOCLOSEAPPLICATIONS: never let Restart Manager close Outlook - this
# refuses to run while Outlook is up anyway. /NORESTART: a reboot is a human's decision.
function Get-InnoArguments {
    param([Parameter(Mandatory = $true)] [string] $SetupLog)
    return @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/NOCLOSEAPPLICATIONS', '/SP-', ('/LOG="' + $SetupLog + '"'))
}

# -1, 0 or 1. A missing left side is older than anything.
function Compare-DottedVersion {
    param([string] $Left, [string] $Right)
    if (-not $Left) { return -1 }
    $l = @($Left.Trim().Split('.') | ForEach-Object { [int]($_ -replace '[^\d]', '') })
    $r = @($Right.Trim().Split('.') | ForEach-Object { [int]($_ -replace '[^\d]', '') })
    $n = [Math]::Max($l.Count, $r.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $a = 0; if ($i -lt $l.Count) { $a = $l[$i] }
        $b = 0; if ($i -lt $r.Count) { $b = $r[$i] }
        if ($a -lt $b) { return -1 }
        if ($a -gt $b) { return 1 }
    }
    return 0
}

# The URL the VSTO runtime keys an inclusion entry by, from the registered Manifest value: the
# `|vstolocal` suffix off, backslashes to forward slashes, file:/// in front - which is the
# Uri.ToString() form the runtime itself writes. [MEASURED: the maintainer's workstation holds
# `file:///C:/Users/.../OutlookAI/Setup/OutlookAI.vsto` for this installer's
# `file:///{app}\OutlookAI.vsto|vstolocal`.] Spaces stay literal, as the runtime writes them.
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

# The runtime compares entries AS URIs, and on a DOS path that comparison ignores case.
function Test-SameInclusionUrl {
    param([string] $A, [string] $B)
    $x = ConvertTo-InclusionUrl $A
    $y = ConvertTo-InclusionUrl $B
    if (-not $x -or -not $y) { return $false }
    return [string]::Equals($x, $y, [System.StringComparison]::OrdinalIgnoreCase)
}

function ConvertTo-RsaKeyValueXml {
    param([Parameter(Mandatory = $true)] [string] $ModulusBase64, [Parameter(Mandatory = $true)] [string] $ExponentBase64)
    $m = ($ModulusBase64 -replace '\s', '')
    $e = ($ExponentBase64 -replace '\s', '')
    return "<RSAKeyValue><Modulus>$m</Modulus><Exponent>$e</Exponent></RSAKeyValue>"
}

# Same rule as Testbed/host/Publish-AddInPayload.ps1: the manifest's ds:RSAKeyValue, every
# occurrence the same key, as the <RSAKeyValue> string the trust store holds.
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

# A key as two hex strings with leading zero bytes dropped - the same two numbers the runtime's
# ExportCspBlob comparison sees, without asking a crypto provider for them.
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

# ---- The mirror of HealthReporting.ReadTuningState. A value is @{ Kind = <RegistryValueKind>;
# Data = <object> }; a missing value is simply not in the table; a missing KEY is $null.
function ConvertTo-HealthBool {
    param($Entry)
    # HealthReporting.AsBool: `if (value is int number) return number != 0; return null;` - so ONLY
    # a REG_DWORD counts. A REG_QWORD reads as Int64 and a REG_SZ as a string: both are null.
    if ($null -eq $Entry -or [string]$Entry.Kind -ne 'DWord') { return $null }
    return ([int]$Entry.Data -ne 0)
}

function ConvertTo-HealthString {
    param($Entry)
    # `readValue(...) as string`: REG_SZ, and REG_EXPAND_SZ (which GetValue expands to a string).
    if ($null -eq $Entry) { return $null }
    $k = [string]$Entry.Kind
    if ($k -ne 'String' -and $k -ne 'ExpandString') { return $null }
    return [string]$Entry.Data
}

function Get-TuningView {
    param([hashtable] $Values)
    $view = [ordered]@{
        KeyPresent = ($null -ne $Values); Managed = $false; Enabled = $null; SearchEnabled = $null
        CachingEnabled = $null; OstEnabled = $null; RestartNeeded = $null; PolicyConflicts = $null; LastReconcileUtc = $null
    }
    if ($null -eq $Values) { return [pscustomobject]$view }
    if ((ConvertTo-HealthBool $Values['Initialized']) -ne $true) { return [pscustomobject]$view }
    $view.Managed = $true
    $view.Enabled = ConvertTo-HealthBool $Values['Enabled']
    $view.SearchEnabled = ConvertTo-HealthBool $Values['SearchEnabled']
    $view.CachingEnabled = ConvertTo-HealthBool $Values['CachingEnabled']
    $view.OstEnabled = ConvertTo-HealthBool $Values['OstEnabled']
    $view.RestartNeeded = ConvertTo-HealthBool $Values['RestartNeeded']
    $conflicts = ConvertTo-HealthString $Values['PolicyConflicts']
    if (-not [string]::IsNullOrWhiteSpace($conflicts)) { $view.PolicyConflicts = $conflicts }
    $view.LastReconcileUtc = ConvertTo-HealthString $Values['LastReconcileUtc']
    return [pscustomobject]$view
}

# What the two tests would fail on, in their own terms.
function Get-TestReadProblems {
    param($View)
    $problems = @()
    if (-not $View.KeyPresent) {
        $problems += "HKCU\$TuningKey does not exist: the add-in has never run its tuning service here. tuning.managed = false fails $($TestReads[0].Fails)."
        return $problems
    }
    if ($View.Managed -ne $true) {
        $problems += "Initialized is not a nonzero REG_DWORD, so tuning.managed = false. Fails $($TestReads[0].Fails)."
        return $problems
    }
    if ($View.Enabled -ne $true) {
        $problems += "Enabled is not a nonzero REG_DWORD, so tuning.enabled is not true. Fails $($TestReads[1].Fails)."
    }
    if ($null -eq $View.LastReconcileUtc) {
        $problems += "LastReconcileUtc is not a REG_SZ, so tuning.lastReconcileUtc is null. Fails $($TestReads[2].Fails)."
    }
    return $problems
}

# OutlookTuningService writes LastReconcileUtc as DateTime.UtcNow.ToString("o"). Fresh means
# written at or after the moment this run started Outlook - proof THIS start ran the add-in.
function Test-ReconcileFresh {
    param([string] $LastReconcileUtc, [DateTime] $StartedUtc)
    if (-not $LastReconcileUtc) { return [pscustomobject]@{ Fresh = $false; Reason = 'no LastReconcileUtc at all' } }
    $parsed = [DateTime]::MinValue
    $ok = [DateTime]::TryParse($LastReconcileUtc, [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)
    if (-not $ok) { return [pscustomobject]@{ Fresh = $false; Reason = "LastReconcileUtc '$LastReconcileUtc' does not parse as a round-trip timestamp" } }
    if ($parsed.Kind -ne [DateTimeKind]::Utc) { $parsed = $parsed.ToUniversalTime() }
    # One second of slack: both readings come from this machine's clock, and "o" keeps 7 digits.
    if ($parsed -ge $StartedUtc.AddSeconds(-1)) { return [pscustomobject]@{ Fresh = $true; Reason = '' } }
    return [pscustomobject]@{ Fresh = $false; Reason = ("LastReconcileUtc {0:o} is older than this run's Outlook start {1:o}: the add-in did not run during it" -f $parsed, $StartedUtc) }
}

function Get-RegistrationProblems {
    param([string] $Manifest, $LoadBehavior, [string] $InstallDir)
    $problems = @()
    if (-not $Manifest) { return @("HKCU\$AddinRegistrationKey has no Manifest value: Outlook has nothing to load.") }
    if (-not $Manifest.EndsWith('|vstolocal', [System.StringComparison]::OrdinalIgnoreCase)) {
        $problems += "Manifest '$Manifest' is not a |vstolocal registration, which is the only kind Installer.iss writes."
    }
    if ($InstallDir) {
        $expected = ConvertTo-InclusionUrl ((Join-Path $InstallDir 'OutlookAI.vsto'))
        if (-not (Test-SameInclusionUrl $Manifest $expected)) {
            $problems += "Manifest points at '$Manifest', not at the installed copy under $InstallDir - something else re-registered the add-in."
        }
    }
    if ($null -eq $LoadBehavior) { $problems += 'LoadBehavior is missing.' }
    elseif ([int]$LoadBehavior -eq 2) { $problems += 'LoadBehavior is 2: Outlook TRIED to load the add-in and it FAILED. Read <app>\OutlookAI.vsto.log (VSTO_LOGALERTS) for why.' }
    elseif ([int]$LoadBehavior -eq 0) { $problems += 'LoadBehavior is 0: the add-in is registered but switched off.' }
    elseif ([int]$LoadBehavior -ne 3) { $problems += "LoadBehavior is $LoadBehavior, not 3 (load at startup)." }
    return $problems
}

# Outlook's hard-disable list holds UTF-16 blobs naming the add-in's path or name.
function Test-DisabledItemsMention {
    param([object[]] $Blobs)
    foreach ($b in @($Blobs)) {
        if ($null -eq $b) { continue }
        $text = [System.Text.Encoding]::Unicode.GetString([byte[]]$b)
        if ($text.ToLowerInvariant().Contains('outlookai')) { return $true }
    }
    return $false
}

# `git archive --format=zip <commit>` stores the commit id as the ZIP COMMENT, which lives at the
# end of the file in the End Of Central Directory record (signature 50 4B 05 06; the comment length
# is at offset 20 and the comment starts at offset 22).
function Get-ZipComment {
    param([byte[]] $Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 22) { return $null }
    for ($i = $Bytes.Length - 22; $i -ge [Math]::Max(0, $Bytes.Length - 22 - 65535); $i--) {
        if ($Bytes[$i] -eq 0x50 -and $Bytes[$i + 1] -eq 0x4B -and $Bytes[$i + 2] -eq 0x05 -and $Bytes[$i + 3] -eq 0x06) {
            $len = [int]$Bytes[$i + 20] + 256 * [int]$Bytes[$i + 21]
            if ($i + 22 + $len -gt $Bytes.Length) { return $null }
            return [System.Text.Encoding]::ASCII.GetString($Bytes, $i + 22, $len)
        }
    }
    return $null
}

function Test-Sha256Text { param([string] $Value) return [bool]($Value -match '^[0-9A-F]{64}$') }

function Test-PayloadManifestShape {
    param($Manifest)
    $problems = @()
    if ($null -eq $Manifest) { return @('no manifest at all') }
    if (-not ([string]$Manifest.commit -match '^[0-9a-f]{40}$')) { $problems += 'commit: not a 40-character lower-case hex id' }
    if (-not ([string]$Manifest.version -match '^\d{1,5}\.\d{1,5}\.\d{1,5}\.\d{1,5}$')) { $problems += 'version: not a four-part version' }
    if ($null -eq $Manifest.installer -or -not $Manifest.installer.file) { $problems += 'installer.file: missing' }
    elseif (-not (Test-Sha256Text ([string]$Manifest.installer.sha256))) { $problems += 'installer.sha256: not an upper-case SHA-256' }
    foreach ($f in @('OutlookAI.dll', 'OutlookAI.vsto', 'OutlookAI.dll.manifest')) {
        if ($null -eq $Manifest.addin -or -not (Test-Sha256Text ([string]$Manifest.addin.$f))) { $problems += "addin.${f}: not an upper-case SHA-256" }
    }
    if ($null -eq $Manifest.signing -or -not ([string]$Manifest.signing.publicKeyXml).StartsWith('<RSAKeyValue><Modulus>')) {
        $problems += 'signing.publicKeyXml: not an <RSAKeyValue> string'
    }
    return $problems
}

# The exclusion state, before and after this run's Outlook start.
function Compare-InteractionFacts {
    param($Before, $After)
    $changes = @()
    if ([string]$Before.PreventIndexingOutlook -ne [string]$After.PreventIndexingOutlook) {
        $changes += "PreventIndexingOutlook went from '$($Before.PreventIndexingOutlook)' to '$($After.PreventIndexingOutlook)'"
    }
    $b = @($Before.MapiRules | Sort-Object)
    $a = @($After.MapiRules | Sort-Object)
    if (($b -join ';') -ne ($a -join ';')) {
        $changes += "the mapi crawl-scope rules went from [$($b -join '; ')] to [$($a -join '; ')]"
    }
    if ([string]$Before.SearchPolicyValue -ne [string]$After.SearchPolicyValue) {
        $changes += "the Outlook Search POLICY value DisableServerAssistedSearch went from '$($Before.SearchPolicyValue)' to '$($After.SearchPolicyValue)'"
    }
    return $changes
}

function Get-McpStatusProblem {
    param([string] $Status)
    if ($Status -eq $StatusAwaitingChoice) {
        return "The add-in has a Claude Code registration QUESTION pending ($StatusAwaitingChoice). It surfaces as a MODAL dialog the next time an Outlook window is visible - in the middle of an InteractiveDesktop test. On a guest it means Claude Code is installed here; decide it once in OutlookAI Settings."
    }
    return $null
}

# Services/OfficeVersions.cs: the first of 16.0, 17.0, 15.0 whose Outlook key is a REAL hive -
# at least one value, or a subkey other than the Resiliency shell our own installer creates under
# every major.
function Select-OfficeVersion {
    param([hashtable] $Hives)
    foreach ($v in $SupportedOffice) {
        $h = $Hives[$v]
        if ($null -eq $h) { continue }
        if (@($h.Values).Count -gt 0) { return $v }
        foreach ($s in @($h.SubKeys)) {
            if (-not [string]::Equals($s, 'Resiliency', [System.StringComparison]::OrdinalIgnoreCase)) { return $v }
        }
    }
    return $null
}

function Get-AddInVerdict {
    param($Facts)
    $problems = @()
    $notes = @()
    if (-not $Facts.Installed) {
        return [pscustomobject]@{ Verdict = $VerdictNotInstalled; ExitCode = 3; Problems = @(); Notes = @('No InstallDir under HKCU\Software\OutlookAI and no registration: the add-in is not installed. Run -Execute.') }
    }
    $problems += @($Facts.RegistrationProblems)
    $problems += @($Facts.FileProblems)
    $problems += @($Facts.TrustProblems)
    $problems += @($Facts.RuntimeProblems)
    $problems += @($Facts.ContractProblems)
    if ($Facts.DisabledItemHit) { $problems += "Outlook's Resiliency\DisabledItems names the add-in: Outlook hard-disabled it after a crash." }
    if ($Facts.McpProblem) { $problems += $Facts.McpProblem }
    $problems += @($Facts.ComProblems)
    $notes += @($Facts.Notes)

    $readProblems = @(Get-TestReadProblems $Facts.Tuning)
    $neverRan = (-not $Facts.Tuning.KeyPresent)
    if (-not $neverRan) { $problems += $readProblems }
    if ($Facts.RequireFresh -and -not $Facts.Fresh) { $problems += $Facts.FreshReason }

    $problems = @($problems | Where-Object { $_ })
    if ($problems.Count -gt 0) {
        return [pscustomobject]@{ Verdict = $VerdictBroken; ExitCode = 1; Problems = $problems; Notes = $notes }
    }
    if ($neverRan) {
        return [pscustomobject]@{ Verdict = $VerdictNeverRan; ExitCode = 2; Problems = @(); Notes = @($notes + $readProblems) }
    }
    return [pscustomobject]@{ Verdict = $VerdictReady; ExitCode = 0; Problems = @(); Notes = $notes }
}

# =============================================================================================
# THE CONTRACT THIS SCRIPT MIRRORS, read from the repository when it is reachable. Every string
# below is copied from a source file; -SelfTest fails the day one of them stops being there.
# =============================================================================================
function Get-ContractChecks {
    return @(
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningKeyPath = @"Software\OutlookAI\Tuning";'; Why = "the tests read HKCU\$TuningKey" }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningInitializedValueName = "Initialized";'; Why = 'Initialized -> tuning.managed' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningEnabledValueName = "Enabled";'; Why = 'Enabled -> tuning.enabled' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningLastReconcileUtcValueName = "LastReconcileUtc";'; Why = 'LastReconcileUtc -> tuning.lastReconcileUtc' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string TuningPolicyConflictsValueName = "PolicyConflicts";'; Why = 'PolicyConflicts is reported' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string McpKeyPath = @"Software\OutlookAI\Mcp";'; Why = 'the registration status is read here' }
        @{ File = 'Services\AddInServerContract.cs'; Needle = 'internal const string McpStatusValueName = "Status";'; Why = 'awaiting_choice is read from Status' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'if (value is int number)'; Why = 'only a REG_DWORD counts as a bool' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'if (AsBool(readValue(Contract.TuningInitializedValueName)) != true)'; Why = 'managed rests on Initialized alone' }
        @{ File = 'McpServer\OutlookAI.Core\Services\HealthReporting.cs'; Needle = 'LastReconcileUtc = readValue(Contract.TuningLastReconcileUtcValueName) as string,'; Why = 'LastReconcileUtc must be a string' }
        @{ File = 'Services\OutlookTuningService.cs'; Needle = 'WriteDword(TuningKeyPath, AddInServerContract.TuningInitializedValueName, 1);'; Why = 'the add-in writes Initialized as a DWORD' }
        @{ File = 'Services\OutlookTuningService.cs'; Needle = 'WriteString(TuningKeyPath, AddInServerContract.TuningLastReconcileUtcValueName, DateTime.UtcNow.ToString("o"));'; Why = 'freshness parses a round-trip UTC timestamp' }
        @{ File = 'Services\OfficeVersions.cs'; Needle = 'internal static readonly string[] Supported = { "16.0", "17.0", "15.0" };'; Why = 'the Office majors, in detection order' }
        @{ File = 'Services\OfficeVersions.cs'; Needle = 'internal const string InstallerFootprintSubKeyName = "Resiliency";'; Why = 'a hive holding only Resiliency is not a real Outlook' }
        @{ File = 'Services\McpRegistrationService.cs'; Needle = 'internal const string StatusAwaitingChoice = "awaiting_choice";'; Why = 'a pending registration question' }
        @{ File = 'Installer.iss'; Needle = 'Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\OutlookAI"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\OutlookAI.vsto|vstolocal"'; Why = 'the registration this verifies' }
        @{ File = 'Installer.iss'; Needle = 'ValueName: "LoadBehavior"; ValueData: "3"'; Why = 'LoadBehavior 3' }
        @{ File = 'Installer.iss'; Needle = 'Root: HKCU; Subkey: "Software\OutlookAI"; ValueType: string; ValueName: "InstallDir"; ValueData: "{app}"'; Why = 'InstallDir locates {app}' }
        @{ File = 'Installer.iss'; Needle = 'DefaultDirName={localappdata}\{#MyAppName}\Setup'; Why = 'per-user install location' }
        @{ File = 'Installer.iss'; Needle = 'PrivilegesRequired=lowest'; Why = 'per-user, so every key is this user''s' }
        @{ File = 'Installer.iss'; Needle = 'if IsNetFramework48Installed and (not IsVstoInstalled) then'; Why = 'the runtime step' }
        @{ File = 'Installer.iss'; Needle = "RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Version', version)"; Why = 'the runtime is detected by the 32-bit v4R key, as here' }
        @{ File = 'Installer.iss'; Needle = 'certutil'', ExpandConstant(''-f -user -addstore TrustedPublisher'; Why = 'the installer imports the certificate for this user' }
        @{ File = 'Installer.iss'; Needle = 'Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "OutlookAI"'; Why = 'the slow-add-in exemption' }
        @{ File = 'AddInAutomation.cs'; Needle = 'bool GetRestartNeeded();'; Why = 'the call INTO the add-in the first run makes' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T3\Phase7LiveMcpToolShapeTests.cs'; Needle = 'Assert.True(report.GetProperty("tuning").GetProperty("managed").GetBoolean());'; Why = 'what the Phase-7 test asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveHealthTests.cs'; Needle = 'Assert.True(report.Tuning.Managed);'; Why = 'what LiveHealthTests asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveHealthTests.cs'; Needle = 'Assert.True(report.Tuning.Enabled);'; Why = 'what LiveHealthTests asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveHealthTests.cs'; Needle = 'Assert.NotNull(report.Tuning.LastReconcileUtc);'; Why = 'what LiveHealthTests asserts' }
        @{ File = 'McpServer\OutlookAI.McpServer.Tests\T2\LiveUiSearchBackendTests.cs'; Needle = 'SKIP: policy-hive DisableServerAssistedSearch='; Why = 'a POLICY value makes that test return green having proved nothing' }
    )
}

# The payload manifest's own copy of the contract files' hashes lets the guest compare the add-in
# it installed with the suite it is about to run - the drift that would matter.
$ContractFiles = @('Services/AddInServerContract.cs', 'Services/OfficeVersions.cs')

# =============================================================================================
# I/O
# =============================================================================================

function Read-RegistryValues {
    param([Microsoft.Win32.RegistryKey] $BaseKey, [string] $Path)
    $key = $BaseKey.OpenSubKey($Path, $false)
    if ($null -eq $key) { return $null }
    try {
        $table = @{}
        foreach ($name in $key.GetValueNames()) {
            $table[$name] = @{ Kind = [string]$key.GetValueKind($name); Data = $key.GetValue($name) }
        }
        return $table
    }
    finally { $key.Close() }
}

function Get-HkcuValues { param([string] $Path) return (Read-RegistryValues -BaseKey ([Microsoft.Win32.Registry]::CurrentUser) -Path $Path) }

function Get-SubKeyNames {
    param([Microsoft.Win32.RegistryKey] $BaseKey, [string] $Path)
    $key = $BaseKey.OpenSubKey($Path, $false)
    if ($null -eq $key) { return @() }
    try { return @($key.GetSubKeyNames()) } finally { $key.Close() }
}

function Resolve-OfficeVersion {
    if ($OfficeVersion) { return $OfficeVersion }
    $hives = @{}
    foreach ($v in $SupportedOffice) {
        $path = "Software\Microsoft\Office\$v\Outlook"
        $values = Get-HkcuValues $path
        if ($null -eq $values) { continue }
        $hives[$v] = @{ Values = @($values.Keys); SubKeys = @(Get-SubKeyNames -BaseKey ([Microsoft.Win32.Registry]::CurrentUser) -Path $path) }
    }
    $found = Select-OfficeVersion -Hives $hives
    if (-not $found) { throw 'No Outlook hive under HKCU\Software\Microsoft\Office\{16.0,17.0,15.0}\Outlook. Is Office installed, and has Outlook started once for this user?' }
    return $found
}

function Get-VstoRuntimeFacts {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
    try {
        $v4r = Read-RegistryValues -BaseKey $base -Path $VstoRuntimeKey
        $v4 = Read-RegistryValues -BaseKey $base -Path $VstoRuntimeKeyV4
        $r = [ordered]@{ V4R = $null; V4 = $null }
        if ($v4r -and $v4r['Version']) { $r.V4R = [string]$v4r['Version'].Data }
        if ($v4 -and $v4['Version']) { $r.V4 = [string]$v4['Version'].Data }
        return [pscustomobject]$r
    }
    finally { $base.Close() }
}

function Get-InstallDir {
    $v = Get-HkcuValues $AppKey
    if ($v -and $v[$InstallDirValue]) { return [string]$v[$InstallDirValue].Data }
    return $null
}

function Get-InclusionEntries {
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

function Get-CertificateKeyXml {
    param([Parameter(Mandatory = $true)] [string] $Path)
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Path)
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cert)
    $p = $rsa.ExportParameters($false)
    return (ConvertTo-RsaKeyValueXml -ModulusBase64 ([Convert]::ToBase64String($p.Modulus)) -ExponentBase64 ([Convert]::ToBase64String($p.Exponent)))
}

# The runtime's own parse: AddInSecurityEntry's constructor calls FromXmlString and refuses the
# entry when it throws. Guest-only - it asks a crypto provider for an ephemeral container.
function Test-KeyXmlParsesLikeTheRuntime {
    param([string] $Xml)
    $csp = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    try {
        $csp.PersistKeyInCsp = $false
        $csp.FromXmlString($Xml)
        $null = $csp.ExportCspBlob($false)
        return $true
    }
    catch { return $false }
    finally { $csp.Dispose() }
}

function Get-InteractionFacts {
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $policy = Read-RegistryValues -BaseKey $hklm -Path $SearchPolicyKey
        $prevent = $null
        if ($policy -and $policy['PreventIndexingOutlook']) { $prevent = $policy['PreventIndexingOutlook'].Data }
        $rules = @()
        foreach ($name in (Get-SubKeyNames -BaseKey $hklm -Path $CrawlRulesKey)) {
            $v = Read-RegistryValues -BaseKey $hklm -Path "$CrawlRulesKey\$name"
            if ($v -and $v['URL'] -and ([string]$v['URL'].Data).StartsWith('mapi', [System.StringComparison]::OrdinalIgnoreCase)) {
                $inc = ''
                if ($v['Include']) { $inc = [string]$v['Include'].Data }
                $rules += ([string]$v['URL'].Data + ' include=' + $inc)
            }
        }
    }
    finally { $hklm.Close() }
    $searchPolicy = $null
    if ($script:ResolvedOffice) {
        $sp = Get-HkcuValues "Software\Policies\Microsoft\Office\$script:ResolvedOffice\Outlook\Search"
        if ($sp -and $sp['DisableServerAssistedSearch']) { $searchPolicy = $sp['DisableServerAssistedSearch'].Data }
    }
    return [pscustomobject]@{ PreventIndexingOutlook = $prevent; MapiRules = $rules; SearchPolicyValue = $searchPolicy }
}

function Get-SessionWindows {
    $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    return @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sid -and $_.MainWindowTitle } |
            ForEach-Object { "[$($_.ProcessName)] '$($_.MainWindowTitle)'" })
}

# Start a process and wait for it with a deadline. .Handle is read before waiting because
# Start-Process -PassThru otherwise loses the exit code (MEASURED, Install-DotnetSdk.ps1). NEVER
# -Verb RunAs: this session is already elevated, and UAC elevation takes the foreground.
function Invoke-Installer {
    param([string] $FilePath, [string[]] $Arguments, [int] $TimeoutMinutes)
    $p = Start-Process -FilePath $FilePath -ArgumentList ($Arguments -join ' ') -PassThru -NoNewWindow
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = $null; TimedOut = $true }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false }
}

# =============================================================================================
# GUARDS
# =============================================================================================
function Assert-TestbedGuestLocal {
    if (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $ExpectedUser -Prefix $ExpectedComputerNamePrefix) { return }
    throw @"
REFUSING TO RUN.

  logged on as : '$env:USERNAME'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$env:COMPUTERNAME'      (must start with: '$ExpectedComputerNamePrefix')

This script installs the OutlookAI add-in, rewrites a VSTO trust entry and starts Outlook. On the
maintainer's workstation that would replace the add-in his Outlook really runs, and nothing would
say so until it misbehaved.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2), and
Testbed/host/New-AnswerFile.ps1 derives their computer name by replacing 'OutlookAI-' with 'OAI-',
so a guest built from the answer file matches both axes. If you named a guest something else, say so:

    -ExpectedUser <username> -ExpectedComputerNamePrefix <prefix>

Do not 'fix' this by widening either default. The defaults are the guard.
"@
}

function Assert-Elevated {
    $p = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'REFUSING TO RUN: this session is not elevated. The VSTO runtime is a machine-wide install. Register-InteractiveTask.ps1 registers its task with RunLevel Highest, which is elevated.'
    }
}

function Assert-Bitness {
    if (-not [Environment]::Is64BitProcess) {
        throw 'REFUSING TO RUN: this is a 32-bit PowerShell. Outlook on the guests is x64, and a 32-bit client would start a COM surrogate rather than talk to it. Use C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe.'
    }
}

function Assert-InteractiveSession {
    $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    if ($sid -eq 0) {
        throw @"
REFUSING TO RUN -Execute IN SESSION 0.

It starts Outlook, and Outlook cannot finish starting in session 0 - PowerShell Direct lands there, and
the call would hang instead of failing. Run it through the interactive session:

    .\Register-InteractiveTask.ps1 -Script "& '$PSCommandPath' -Execute"

-Verify, without -WithOutlook, is fine here.
"@
    }
}

function Assert-OutlookClosed {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return }
    throw @"
REFUSING TO RUN: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')).

The add-in loads when Outlook STARTS, and the proof this script gives - the add-in wrote its state
AFTER a start this script made - needs that start to be this script's. The installer would also be
replacing files under a running add-in.

Restart the guest - the proven way to a clean Outlook here - and run this again. DO NOT taskkill
OUTLOOK.EXE: mailbox-safety rule 7 forbids it outright.
"@
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()
    # -ceq and String.Contains only. Never -like: Testbed/README.md section 4b.
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
    function V([string] $kind, $data) { return @{ Kind = $kind; Data = $data } }

    Write-Host '== the guard =='
    Test-Case 'vmadmin on OAI-INDEXED passes' $true (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'OAI-INDEXED' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'the maintainer workstation is refused' $false (Test-GuestIdentity -UserName 'jori' -ComputerName 'PC657' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'vmadmin on a non-guest name is refused' $false (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'PC657' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'the right name as another user is refused' $false (Test-GuestIdentity -UserName 'jori' -ComputerName 'OAI-INDEXED' -Users @('vmadmin') -Prefix 'OAI-')
    Test-Case 'an empty prefix is refused, never "matches everything"' $false (Test-GuestIdentity -UserName 'vmadmin' -ComputerName 'OAI-X' -Users @('vmadmin') -Prefix '')

    Write-Host ''
    Write-Host '== the installer switches =='
    $inno = Get-InnoArguments -SetupLog 'C:\OutlookAI-Q5\s.log'
    Test-Case 'are exactly these' '/VERYSILENT | /SUPPRESSMSGBOXES | /NORESTART | /NOCLOSEAPPLICATIONS | /SP- | /LOG="C:\OutlookAI-Q5\s.log"' $inno
    Test-Case 'never merely /SILENT, which shows progress' $false ($inno -contains '/SILENT')

    Write-Host ''
    Write-Host '== the VSTO runtime version =='
    Test-Case 'the pinned version is current' 0 (Compare-DottedVersion '10.0.60917' $VstoRuntimeVersion)
    Test-Case 'what Office itself may bring is older' -1 (Compare-DottedVersion '10.0.60910' $VstoRuntimeVersion)
    Test-Case 'a later one is newer' 1 (Compare-DottedVersion '10.0.60918' $VstoRuntimeVersion)
    Test-Case 'a two-part pad compares' 0 (Compare-DottedVersion '10.0.60917.0' $VstoRuntimeVersion)
    Test-Case 'absent is older than anything' -1 (Compare-DottedVersion '' $VstoRuntimeVersion)

    Write-Host ''
    Write-Host '== the trust entry URL =='
    Test-Case 'from the installer''s registration, as the runtime wrote it on the host' 'file:///C:/Users/vmadmin/AppData/Local/OutlookAI/Setup/OutlookAI.vsto' (ConvertTo-InclusionUrl 'file:///C:\Users\vmadmin\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal')
    Test-Case 'from a plain path' 'file:///C:/a/OutlookAI.vsto' (ConvertTo-InclusionUrl 'C:\a\OutlookAI.vsto')
    Test-Case 'spaces stay literal, as the runtime writes them' 'file:///C:/a b/x.vsto' (ConvertTo-InclusionUrl 'file:///C:\a b\x.vsto|vstolocal')
    Test-Case 'an already-normal URL is unchanged' 'file:///C:/a/x.vsto' (ConvertTo-InclusionUrl 'file:///C:/a/x.vsto')
    Test-Case 'the same URL in another case is the same entry' $true (Test-SameInclusionUrl 'file:///c:/USERS/x.vsto' 'file:///C:/Users/x.vsto')
    Test-Case 'a different folder is a different entry' $false (Test-SameInclusionUrl 'file:///C:/a/x.vsto' 'file:///C:/b/x.vsto')
    Test-Case 'nothing is never a match' $false (Test-SameInclusionUrl '' 'file:///C:/a/x.vsto')

    Write-Host ''
    Write-Host '== the signing key =='
    $k1 = '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>'
    $manifest = '<asmv1:assembly xmlns:asmv1="urn:schemas-microsoft-com:asm.v1"><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><KeyInfo><KeyValue><RSAKeyValue><Modulus>AAEC
AwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue></KeyValue></KeyInfo></Signature></asmv1:assembly>'
    Test-Case 'read out of a signed manifest, namespace and line break ignored' $k1 (Get-ManifestSigningKeyXml -ManifestXml $manifest)
    Test-Case 'the same key with a namespace and whitespace is the same key' $true (Test-SameRsaKey $k1 '<RSAKeyValue xmlns="http://www.w3.org/2000/09/xmldsig#"> <Modulus>AAEC AwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')
    Test-Case 'a leading zero byte does not make a different key' $true (Test-SameRsaKey '<RSAKeyValue><Modulus>AQIDBA==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>' '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')
    Test-Case 'a different modulus is a different key' $false (Test-SameRsaKey $k1 '<RSAKeyValue><Modulus>BBBBBBB=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>')
    Test-Case 'a different exponent is a different key' $false (Test-SameRsaKey $k1 '<RSAKeyValue><Modulus>AAECAwQ=</Modulus><Exponent>Aw==</Exponent></RSAKeyValue>')
    Test-Case 'garbage is never the same key' $false (Test-SameRsaKey $k1 'not xml')
    $threw = $false; try { $null = Get-ManifestSigningKeyXml '<a/>' } catch { $threw = $true }
    Test-Case 'an unsigned manifest is refused' $true $threw

    Write-Host ''
    Write-Host '== what the tests read: the mirror of HealthReporting.ReadTuningState =='
    $good = @{ Initialized = (V 'DWord' 1); Enabled = (V 'DWord' 1); SearchEnabled = (V 'DWord' 1); CachingEnabled = (V 'DWord' 1)
               OstEnabled = (V 'DWord' 1); RestartNeeded = (V 'DWord' 1); PolicyConflicts = (V 'String' ''); LastReconcileUtc = (V 'String' '2026-09-24T12:00:00.0000000Z') }
    $v = Get-TuningView $good
    Test-Case 'a first-run state is managed' $true $v.Managed
    Test-Case 'and enabled' $true $v.Enabled
    Test-Case 'and carries its reconcile time' '2026-09-24T12:00:00.0000000Z' $v.LastReconcileUtc
    Test-Case 'an empty PolicyConflicts reads as none' $null $v.PolicyConflicts
    Test-Case 'and the tests have nothing to fail on' 0 (Get-TestReadProblems $v).Count
    $v = Get-TuningView $null
    Test-Case 'no key: not managed' $false $v.Managed
    Test-Case 'no key: named as never having run' $true ((Get-TestReadProblems $v) -join ' ').Contains('has never run its tuning service')
    $bad = $good.Clone(); $bad['Initialized'] = (V 'QWord' 1)
    Test-Case 'Initialized as a REG_QWORD 1 is NOT managed - AsBool takes an int only' $false (Get-TuningView $bad).Managed
    $bad = $good.Clone(); $bad['Initialized'] = (V 'String' '1')
    Test-Case 'Initialized as the string "1" is NOT managed' $false (Get-TuningView $bad).Managed
    $bad = $good.Clone(); $bad['Initialized'] = (V 'DWord' 0)
    Test-Case 'Initialized 0 is not managed' $false (Get-TuningView $bad).Managed
    Test-Case 'and the problem names both tests' $true ((Get-TestReadProblems (Get-TuningView $bad)) -join ' ').Contains('Phase7LiveMcpToolShapeTests')
    $bad = $good.Clone(); $bad.Remove('Initialized')
    Test-Case 'no Initialized at all is not managed' $false (Get-TuningView $bad).Managed
    $bad = $good.Clone(); $bad['Enabled'] = (V 'DWord' 0)
    Test-Case 'Enabled 0 fails LiveHealthTests only' $true ((Get-TestReadProblems (Get-TuningView $bad)) -join ' ').Contains('tuning.enabled is not true')
    $bad = $good.Clone(); $bad['Enabled'] = (V 'QWord' 1)
    Test-Case 'Enabled as a REG_QWORD is not true' $null (Get-TuningView $bad).Enabled
    $bad = $good.Clone(); $bad.Remove('LastReconcileUtc')
    Test-Case 'no LastReconcileUtc is a null the test fails on' $true ((Get-TestReadProblems (Get-TuningView $bad)) -join ' ').Contains('tuning.lastReconcileUtc is null')
    $bad = $good.Clone(); $bad['LastReconcileUtc'] = (V 'MultiString' @('2026-09-24T12:00:00Z'))
    Test-Case 'LastReconcileUtc as a REG_MULTI_SZ is null - "as string" fails on string[]' $null (Get-TuningView $bad).LastReconcileUtc
    $bad = $good.Clone(); $bad['LastReconcileUtc'] = (V 'ExpandString' '2026-09-24T12:00:00Z')
    Test-Case 'LastReconcileUtc as a REG_EXPAND_SZ is still a string' '2026-09-24T12:00:00Z' (Get-TuningView $bad).LastReconcileUtc
    $bad = $good.Clone(); $bad['PolicyConflicts'] = (V 'String' 'caching.policy.SyncWindowSetting')
    Test-Case 'a real PolicyConflicts is reported' 'caching.policy.SyncWindowSetting' (Get-TuningView $bad).PolicyConflicts

    Write-Host ''
    Write-Host '== freshness =='
    $t0 = New-Object DateTime(2026, 9, 24, 12, 0, 0, [DateTimeKind]::Utc)
    Test-Case 'written after the start is fresh' $true (Test-ReconcileFresh '2026-09-24T12:00:05.1234567Z' $t0).Fresh
    Test-Case 'written within the second before is fresh' $true (Test-ReconcileFresh '2026-09-24T11:59:59.5000000Z' $t0).Fresh
    Test-Case 'written an hour before is not' $false (Test-ReconcileFresh '2026-09-24T11:00:00.0000000Z' $t0).Fresh
    Test-Case 'and says the add-in did not run during this start' $true (Test-ReconcileFresh '2026-09-24T11:00:00.0000000Z' $t0).Reason.Contains('did not run during it')
    Test-Case 'an offset timestamp is compared in UTC' $true (Test-ReconcileFresh '2026-09-24T14:00:05.0000000+02:00' $t0).Fresh
    Test-Case 'garbage does not parse' $false (Test-ReconcileFresh 'yesterday' $t0).Fresh
    Test-Case 'absent is not fresh' $false (Test-ReconcileFresh '' $t0).Fresh

    Write-Host ''
    Write-Host '== the registration =='
    $dir = 'C:\Users\vmadmin\AppData\Local\OutlookAI\Setup'
    $reg = 'file:///C:\Users\vmadmin\AppData\Local\OutlookAI\Setup\OutlookAI.vsto|vstolocal'
    Test-Case 'the installer''s own registration is clean' 0 (Get-RegistrationProblems -Manifest $reg -LoadBehavior 3 -InstallDir $dir).Count
    Test-Case 'LoadBehavior 2 is a failed load, and says where the reason is' $true ((Get-RegistrationProblems -Manifest $reg -LoadBehavior 2 -InstallDir $dir) -join ' ').Contains('OutlookAI.vsto.log')
    Test-Case 'LoadBehavior 0 is switched off' $true ((Get-RegistrationProblems -Manifest $reg -LoadBehavior 0 -InstallDir $dir) -join ' ').Contains('switched off')
    Test-Case 'a build output registered instead is caught' $true ((Get-RegistrationProblems -Manifest 'file:///C:/Source/OutlookAI/bin/Release/OutlookAI.vsto|vstolocal' -LoadBehavior 3 -InstallDir $dir) -join ' ').Contains('not at the installed copy')
    Test-Case 'a ClickOnce-style registration is caught' $true ((Get-RegistrationProblems -Manifest 'file:///C:\x\OutlookAI.vsto' -LoadBehavior 3 -InstallDir '') -join ' ').Contains('|vstolocal')
    Test-Case 'no registration at all' $true ((Get-RegistrationProblems -Manifest '' -LoadBehavior $null -InstallDir $dir) -join ' ').Contains('nothing to load')

    Write-Host ''
    Write-Host '== Outlook''s hard-disable list =='
    $blob = [System.Text.Encoding]::Unicode.GetBytes('file:///C:/Users/vmadmin/AppData/Local/OutlookAI/Setup/OutlookAI.vsto|vstolocal')
    Test-Case 'a blob naming the add-in is found' $true (Test-DisabledItemsMention @(, $blob))
    Test-Case 'another add-in is not' $false (Test-DisabledItemsMention @(, ([System.Text.Encoding]::Unicode.GetBytes('SomeOtherAddin.dll'))))
    Test-Case 'nothing is nothing' $false (Test-DisabledItemsMention @())

    Write-Host ''
    Write-Host '== the suite''s commit, from its zip comment =='
    $comment = 'fb19ccf494fe25460b985ed0f4cf0cc1435e3e40'
    $eocd = New-Object byte[] 22
    $eocd[0] = 0x50; $eocd[1] = 0x4B; $eocd[2] = 0x05; $eocd[3] = 0x06; $eocd[20] = [byte]$comment.Length; $eocd[21] = 0
    $zip = [byte[]](@(1, 2, 3, 4) + $eocd + [System.Text.Encoding]::ASCII.GetBytes($comment))
    Test-Case 'the comment is the commit' $comment (Get-ZipComment $zip)
    $noComment = [byte[]](@(9, 9) + $eocd)
    $noComment[2 + 20] = 0
    Test-Case 'a zip with no comment gives an empty one' '' (Get-ZipComment $noComment)
    Test-Case 'not a zip gives nothing' $null (Get-ZipComment ([byte[]](1..40)))

    Write-Host ''
    Write-Host '== the payload manifest =='
    $okManifest = [pscustomobject]@{
        commit = ('a' * 40); version = '99.99.99.0'
        installer = [pscustomobject]@{ file = 'OutlookAI-v99.99.99.0.exe'; sha256 = ('A' * 64) }
        addin = [pscustomobject]@{ 'OutlookAI.dll' = ('B' * 64); 'OutlookAI.vsto' = ('C' * 64); 'OutlookAI.dll.manifest' = ('D' * 64) }
        signing = [pscustomobject]@{ publicKeyXml = $k1 }
    }
    Test-Case 'a complete manifest is accepted' 0 (Test-PayloadManifestShape $okManifest).Count
    $bad = $okManifest.PSObject.Copy(); $bad.version = '3.1.0'
    Test-Case 'a three-part version is refused' $true ((Test-PayloadManifestShape $bad) -join ' ').Contains('version:')
    $bad = $okManifest.PSObject.Copy(); $bad.installer = [pscustomobject]@{ file = 'x.exe'; sha256 = 'abc' }
    Test-Case 'a short hash is refused' $true ((Test-PayloadManifestShape $bad) -join ' ').Contains('installer.sha256')

    Write-Host ''
    Write-Host '== the index exclusion, before and after =='
    $before = [pscustomobject]@{ PreventIndexingOutlook = 1; MapiRules = @('mapi16://{S-1-5-21-1}/ include=0'); SearchPolicyValue = $null }
    Test-Case 'unchanged is unchanged' 0 (Compare-InteractionFacts $before $before).Count
    $after = [pscustomobject]@{ PreventIndexingOutlook = 1; MapiRules = @('mapi16://{S-1-5-21-1}/ include=1'); SearchPolicyValue = $null }
    Test-Case 'a rule flipping back to included is caught' $true ((Compare-InteractionFacts $before $after) -join ' ').Contains('crawl-scope rules')
    $after = [pscustomobject]@{ PreventIndexingOutlook = $null; MapiRules = @('mapi16://{S-1-5-21-1}/ include=0'); SearchPolicyValue = $null }
    Test-Case 'the policy going away is caught' $true ((Compare-InteractionFacts $before $after) -join ' ').Contains('PreventIndexingOutlook')
    $after = [pscustomobject]@{ PreventIndexingOutlook = 1; MapiRules = @('mapi16://{S-1-5-21-1}/ include=0'); SearchPolicyValue = 1 }
    Test-Case 'a Search POLICY value appearing is caught' $true ((Compare-InteractionFacts $before $after) -join ' ').Contains('POLICY')

    Write-Host ''
    Write-Host '== the Office hive, as Services/OfficeVersions.cs picks it =='
    Test-Case 'a real 16.0 hive wins' '16.0' (Select-OfficeVersion @{ '16.0' = @{ Values = @('x'); SubKeys = @('Profiles') }; '15.0' = @{ Values = @(); SubKeys = @('Resiliency') } })
    Test-Case 'the installer''s Resiliency shell alone is not a hive' $null (Select-OfficeVersion @{ '15.0' = @{ Values = @(); SubKeys = @('Resiliency') }; '17.0' = @{ Values = @(); SubKeys = @('Resiliency') } })
    Test-Case '17.0 is tried before 15.0' '17.0' (Select-OfficeVersion @{ '17.0' = @{ Values = @(); SubKeys = @('Search') }; '15.0' = @{ Values = @('a'); SubKeys = @() } })

    Write-Host ''
    Write-Host '== the registration question =='
    Test-Case 'a pending question is a problem' $true ([bool](Get-McpStatusProblem $StatusAwaitingChoice))
    Test-Case 'no Claude Code on the guest is not' $null (Get-McpStatusProblem $StatusNoClaude)

    Write-Host ''
    Write-Host '== the verdict =='
    $installedFacts = @{ Installed = $true; RegistrationProblems = @(); FileProblems = @(); TrustProblems = @(); RuntimeProblems = @(); ContractProblems = @()
                         DisabledItemHit = $false; McpProblem = $null; ComProblems = @(); Notes = @(); Tuning = (Get-TuningView $good); RequireFresh = $false; Fresh = $false; FreshReason = '' }
    $r = Get-AddInVerdict ([pscustomobject]$installedFacts)
    Test-Case 'everything in place is READY' $VerdictReady $r.Verdict
    Test-Case 'and exits 0' 0 $r.ExitCode
    $f = $installedFacts.Clone(); $f.Tuning = (Get-TuningView $null)
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case 'installed but never ran is its own verdict' $VerdictNeverRan $r.Verdict
    Test-Case 'exiting 2' 2 $r.ExitCode
    $f = $installedFacts.Clone(); $f.Tuning = (Get-TuningView $null); $f.TrustProblems = @('no trust entry')
    Test-Case 'never ran AND untrusted is BROKEN - the first load would prompt' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.RequireFresh = $true; $f.Fresh = $false; $f.FreshReason = 'stale'
    Test-Case 'under -Execute a stale state is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.DisabledItemHit = $true
    Test-Case 'a hard-disabled add-in is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $bad = $good.Clone(); $bad['Enabled'] = (V 'DWord' 0); $f.Tuning = (Get-TuningView $bad)
    Test-Case 'a state the tests would fail on is BROKEN' $VerdictBroken (Get-AddInVerdict ([pscustomobject]$f)).Verdict
    $f = $installedFacts.Clone(); $f.Installed = $false
    $r = Get-AddInVerdict ([pscustomobject]$f)
    Test-Case 'not installed is its own verdict' $VerdictNotInstalled $r.Verdict
    Test-Case 'exiting 3' 3 $r.ExitCode

    Write-Host ''
    Write-Host '== the contract this script mirrors, read from the repository =='
    $repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    if (Test-Path -LiteralPath (Join-Path $repo 'Services\AddInServerContract.cs')) {
        foreach ($c in (Get-ContractChecks)) {
            $path = Join-Path $repo $c.File
            $text = ''
            if (Test-Path -LiteralPath $path) { $text = [System.IO.File]::ReadAllText($path) }
            Test-Case "$($c.File): $($c.Why)" $true ($text.Contains($c.Needle))
        }
    }
    else {
        Write-Host "  SKIP no repository at $repo - the contract cannot be read from here."
    }

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a guest can settle these, and nothing above stands in for them:'
    Write-Host '  * that vstor_redist.exe /q /norestart installs silently in the interactive session, and what'
    Write-Host '    Office LTSC 2024 already registers under VSTO Runtime Setup\v4R before it does'
    Write-Host '  * that the installer, run /VERYSILENT from an elevated interactive task, installs for vmadmin'
    Write-Host '  * that the written inclusion entry really retires the trust prompt for this build''s key'
    Write-Host '  * that a COM-started, headless Outlook loads the add-in and it writes its tuning state'
    Write-Host '  * whether Outlook stays up or closes by itself once the last COM reference is released'
    Write-Host '  * that the two live tests then pass on BOTH guests, which is the claim all of this is for'

    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# FACTS - read, never written.
# =============================================================================================
function Get-AddInFacts {
    param([bool] $RequireFresh, [DateTime] $StartedUtc, $Payload)
    $facts = [ordered]@{
        Installed = $false; RegistrationProblems = @(); FileProblems = @(); TrustProblems = @(); RuntimeProblems = @()
        ContractProblems = @(); DisabledItemHit = $false; McpProblem = $null; ComProblems = @(); Notes = @()
        Tuning = $null; RequireFresh = $RequireFresh; Fresh = $false; FreshReason = ''
    }

    $installDir = Get-InstallDir
    $reg = Get-HkcuValues $AddinRegistrationKey
    $manifestValue = $null; $loadBehavior = $null
    if ($reg) {
        if ($reg['Manifest']) { $manifestValue = [string]$reg['Manifest'].Data }
        if ($reg['LoadBehavior']) { $loadBehavior = $reg['LoadBehavior'].Data }
    }
    $facts.Installed = [bool]($installDir -or $reg)
    Say "  InstallDir    : $installDir"
    Say "  registration  : Manifest=$manifestValue  LoadBehavior=$loadBehavior"
    $facts.RegistrationProblems = @(Get-RegistrationProblems -Manifest $manifestValue -LoadBehavior $loadBehavior -InstallDir $installDir)

    # Files, and which build they are.
    $vsto = $null
    if ($installDir) {
        foreach ($f in @('OutlookAI.vsto', 'OutlookAI.dll.manifest', 'OutlookAI.dll', 'OutlookAI.cer')) {
            if (-not (Test-Path -LiteralPath (Join-Path $installDir $f))) { $facts.FileProblems += "{app}\$f is missing ($installDir)." }
        }
        $vsto = Join-Path $installDir 'OutlookAI.vsto'
        if ($Payload -and (Test-Path -LiteralPath (Join-Path $installDir 'OutlookAI.dll'))) {
            $dllHash = (Get-FileHash -LiteralPath (Join-Path $installDir 'OutlookAI.dll') -Algorithm SHA256).Hash
            if ($dllHash -ne [string]$Payload.addin.'OutlookAI.dll') {
                $facts.FileProblems += "the installed OutlookAI.dll ($dllHash) is not the payload's ($($Payload.addin.'OutlookAI.dll')): this guest runs a different build than the one staged."
            }
            else { Say "  build         : OutlookAI.dll is the payload's, commit $($Payload.commit)" }
        }
        $logFile = Join-Path $installDir 'OutlookAI.vsto.log'
        if (Test-Path -LiteralPath $logFile) {
            $facts.Notes += "the VSTO runtime logged errors to $logFile (VSTO_LOGALERTS). Its tail:"
            foreach ($l in (Get-Content -LiteralPath $logFile -Tail 12)) { $facts.Notes += "    $l" }
        }
    }

    # The VSTO runtime.
    $rt = Get-VstoRuntimeFacts
    Say "  VSTO runtime  : v4R=$($rt.V4R)  v4=$($rt.V4)"
    if ((Compare-DottedVersion $rt.V4R $VstoRuntimeVersion) -lt 0) {
        $facts.RuntimeProblems += "VSTO Runtime Setup\v4R reports '$($rt.V4R)', below the pinned $VstoRuntimeVersion. -Execute installs it from $VstoRuntimePath."
    }

    # Trust.
    if ($manifestValue -and $vsto -and (Test-Path -LiteralPath $vsto)) {
        $url = ConvertTo-InclusionUrl $manifestValue
        $key = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText($vsto))
        # Not $matches: that is PowerShell's automatic -match variable, and any -match in scope rewrites it.
        $urlEntries = @(Get-InclusionEntries | Where-Object { Test-SameInclusionUrl $_.Url $url })
        $good = @($urlEntries | Where-Object { Test-SameRsaKey $_.PublicKey $key })
        Say "  trust         : $($urlEntries.Count) inclusion entr(y/ies) for $url, $($good.Count) with this build's key"
        if ($good.Count -eq 0) {
            $facts.TrustProblems += "no VSTO inclusion entry vouches for $url with the key the installed manifest is signed by: the first load would put the ClickOnce trust prompt on the console, which on this guest is a hang."
        }
        if (-not (Test-KeyXmlParsesLikeTheRuntime $key)) { $facts.TrustProblems += 'the manifest key does not parse with FromXmlString, which is how the runtime reads an entry.' }
        $cer = Join-Path $installDir 'OutlookAI.cer'
        if (Test-Path -LiteralPath $cer) {
            if (-not (Test-SameRsaKey (Get-CertificateKeyXml -Path $cer) $key)) { $facts.TrustProblems += '{app}\OutlookAI.cer is not the certificate the manifest is signed with.' }
        }
        if ($Payload -and -not (Test-SameRsaKey ([string]$Payload.signing.publicKeyXml) $key)) {
            $facts.TrustProblems += 'the installed manifest is not signed by the key the payload manifest names.'
        }
    }

    # Outlook's own disable machinery.
    $office = $script:ResolvedOffice
    $disabled = Get-HkcuValues "Software\Microsoft\Office\$office\Outlook\Resiliency\DisabledItems"
    if ($disabled) {
        $blobs = @(); foreach ($k in $disabled.Keys) { if ($disabled[$k].Kind -eq 'Binary') { $blobs += , ([byte[]]$disabled[$k].Data) } }
        $facts.DisabledItemHit = Test-DisabledItemsMention $blobs
    }
    $dnd = Get-HkcuValues "Software\Microsoft\Office\$office\Outlook\Resiliency\DoNotDisableAddinList"
    if (-not ($dnd -and $dnd[$AddinName])) {
        $facts.Notes += "no DoNotDisableAddinList entry for $AddinName under Office $office - Outlook may disable it after a slow start (the installer normally writes one)."
    }

    # The registration question.
    $mcp = Get-HkcuValues $McpKey
    $mcpStatus = $null
    if ($mcp -and $mcp['Status']) { $mcpStatus = [string]$mcp['Status'].Data }
    Say "  registration question status: $mcpStatus"
    $facts.McpProblem = Get-McpStatusProblem $mcpStatus

    # The Search POLICY value that turns LiveUiSearchBackendTests into a green that proves nothing.
    $sp = Get-HkcuValues "Software\Policies\Microsoft\Office\$office\Outlook\Search"
    if ($sp -and $sp['DisableServerAssistedSearch']) {
        $facts.Notes += "a POLICY DisableServerAssistedSearch exists: T2 LiveUiSearchBackendTests will print SKIP and pass having proved nothing. Nothing in the add-in writes it; find what did."
    }

    # The state the tests read.
    $facts.Tuning = Get-TuningView (Get-HkcuValues $TuningKey)
    if ($RequireFresh) {
        $fresh = Test-ReconcileFresh ([string]$facts.Tuning.LastReconcileUtc) $StartedUtc
        $facts.Fresh = $fresh.Fresh
        $facts.FreshReason = $fresh.Reason
    }

    # The add-in's contract against the suite's.
    if ($Payload -and $Payload.contract) {
        foreach ($rel in $ContractFiles) {
            $suiteFile = Join-Path $SuiteSourceRoot ($rel.Replace('/', '\'))
            $want = [string]$Payload.contract.$rel
            if (-not (Test-Path -LiteralPath $suiteFile)) { $facts.Notes += "no suite source at $suiteFile, so the add-in's $rel was not compared with the suite's."; continue }
            $have = (Get-FileHash -LiteralPath $suiteFile -Algorithm SHA256).Hash
            if ($have -ne $want) {
                $facts.ContractProblems += "$rel differs between the add-in (built from $($Payload.commit)) and the suite at ${SuiteSourceRoot}: the tests may read names this add-in does not write. Build both from one commit."
            }
        }
    }
    elseif ($Payload) {
        $facts.Notes += 'the payload manifest carries no contract hashes, so the add-in was not compared with the suite.'
    }
    if ($Payload -and (Test-Path -LiteralPath $SuiteSourceZip)) {
        $suiteCommit = Get-ZipComment ([System.IO.File]::ReadAllBytes($SuiteSourceZip))
        if ($suiteCommit -and $suiteCommit -ne [string]$Payload.commit) {
            $facts.Notes += "the add-in was built from $($Payload.commit) and the suite from $suiteCommit. Fine while the contract files agree (checked above); one commit for both is the tidy state."
        }
        elseif ($suiteCommit) { Say "  suite         : built from the same commit, $suiteCommit" }
    }

    return [pscustomobject]$facts
}

function Write-TestReadBlock {
    param($View)
    Say ''
    Say "== What the live tests read - HKCU\$TuningKey, through HealthReporting.ReadTuningState =="
    $raw = Get-HkcuValues $TuningKey
    foreach ($t in $TestReads) {
        $shown = '(absent)'
        if ($raw -and $raw[$t.Name]) { $shown = "$($raw[$t.Name].Kind) $($raw[$t.Name].Data)" }
        Say ("  {0,-17} {1,-38} -> {2}" -f $t.Name, $shown, $t.Means)
    }
    Say ("  => managed={0} enabled={1} lastReconcileUtc={2}" -f $View.Managed, $View.Enabled, $View.LastReconcileUtc)
    Say ("     searchEnabled={0} cachingEnabled={1} ostEnabled={2} restartNeeded={3} policyConflicts={4}" -f
        $View.SearchEnabled, $View.CachingEnabled, $View.OstEnabled, $View.RestartNeeded, $View.PolicyConflicts)
}

function Write-Verdict {
    param($Result)
    Say ''
    Say "VERDICT: $($Result.Verdict)"
    foreach ($p in $Result.Problems) { Say "  - $p" }
    foreach ($n in $Result.Notes) { Say "  note: $n" }
    Say ''
    Say 'A REGISTRY VALUE IS NOT A PASSING TEST. The verdict above says the state the tests read is in'
    Say 'place. The claim that matters - both tests pass on this guest - is made by running them.'
}

# =============================================================================================
# MAIN
# =============================================================================================
if (-not ($Execute -or $Verify)) {
    Assert-TestbedGuestLocal
    Say 'DRY RUN. Nothing is written. -Execute would, in session 1, elevated, with Outlook closed:'
    Say "  1. check   $PayloadRoot\$ManifestFileName and the installer's SHA-256 against it"
    Say "  2. install the VSTO runtime from $VstoRuntimePath (SHA-256 $VstoRuntimeSha256), unless v4R >= $VstoRuntimeVersion"
    Say ("  3. run     the installer " + ((Get-InnoArguments -SetupLog 'C:\OutlookAI-Q5\install-addin-setup.log') -join ' '))
    Say "  4. write   ONE entry under HKCU\$InclusionKey for the installed manifest's URL and signing key"
    Say "  5. write   HKCU\$EnvironmentKey $LogAlertsName=1"
    Say '  6. start   Outlook once over COM, headless, in a watchdogged child job; never quit or kill it'
    Say "  7. verify  the state the tests read under HKCU\$TuningKey, written by THAT start"
    Say ''
    Say 'Run -Verify to report on what is here now, or through Register-InteractiveTask.ps1 with -Execute.'
    exit 0
}

Assert-TestbedGuestLocal

# Defined AFTER the guard on purpose: this job opens an Outlook COM session, and check 9 of
# check-testbed-references.ps1 requires every write to come after the guest guard in file
# order. It is only ever started on the -Execute path below, so nothing about behaviour moved.
# The FIRST RUN, in a child job so a modal dialog becomes a reported timeout instead of this
# script's hang. Everything the job holds is released in its finally; it never quits Outlook.
$FirstRunJob = {
    param([long] $StartedTicks, [int] $TimeoutSeconds, [string] $TuningKey, [string] $McpKey, [string] $AddinName)
    $ErrorActionPreference = 'Stop'
    $started = New-Object DateTime($StartedTicks, [DateTimeKind]::Utc)
    $r = [ordered]@{ ComSeconds = $null; MapiInit = $null; TuningAfterSeconds = $null; McpAfterSeconds = $null; Connect = $null; RestartNeeded = $null; AutomationAnswered = $false; Error = $null }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $app = $null; $ns = $null; $addins = $null; $addin = $null; $obj = $null
    function Get-Fresh([string] $key) {
        $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($key, $false)
        if ($null -eq $k) { return $false }
        try {
            $s = $k.GetValue('LastReconcileUtc') -as [string]
            if (-not $s) { return $false }
            $t = [DateTime]::MinValue
            if (-not [DateTime]::TryParse($s, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$t)) { return $false }
            return ($t.ToUniversalTime() -ge $started.AddSeconds(-1))
        }
        finally { $k.Close() }
    }
    try {
        $app = New-Object -ComObject Outlook.Application
        $r.ComSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
        try { $ns = $app.GetNamespace('MAPI'); $null = $ns.GetDefaultFolder(6); $r.MapiInit = 'ok' }
        catch { $r.MapiInit = $_.Exception.Message }
        # The tuning state is what the tests read, so it gets the whole deadline. The registration
        # reconcile runs on a worker thread and is only reported, so it gets 30 s more at most.
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while ([DateTime]::UtcNow -lt $deadline -and $null -eq $r.TuningAfterSeconds) {
            if (Get-Fresh $TuningKey) { $r.TuningAfterSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1) }
            else { Start-Sleep -Seconds 2 }
        }
        $mcpDeadline = [DateTime]::UtcNow.AddSeconds(30)
        while ([DateTime]::UtcNow -lt $mcpDeadline -and $null -eq $r.McpAfterSeconds) {
            if (Get-Fresh $McpKey) { $r.McpAfterSeconds = [Math]::Round($sw.Elapsed.TotalSeconds, 1) }
            else { Start-Sleep -Seconds 2 }
        }
        try {
            $addins = $app.COMAddIns
            $addin = $addins.Item($AddinName)
            $r.Connect = [bool]$addin.Connect
            $obj = $addin.Object
            if ($null -ne $obj) { $r.RestartNeeded = [bool]$obj.GetRestartNeeded(); $r.AutomationAnswered = $true }
        }
        catch { $r.Error = 'COMAddIns: ' + $_.Exception.Message }
    }
    catch { $r.Error = $_.Exception.GetType().Name + ': ' + $_.Exception.Message }
    finally {
        foreach ($o in @($obj, $addin, $addins, $ns, $app)) {
            if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    [pscustomobject]$r
}

# -Verify -WithOutlook: attach to an Outlook ALREADY RUNNING in this session and ask the same
# two questions. Never starts one: GetActiveObject finds a running instance or throws.
$AttachJob = {
    param([string] $AddinName)
    $r = [ordered]@{ Connect = $null; RestartNeeded = $null; AutomationAnswered = $false; Error = $null }
    $app = $null; $addins = $null; $addin = $null; $obj = $null
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $addins = $app.COMAddIns
        $addin = $addins.Item($AddinName)
        $r.Connect = [bool]$addin.Connect
        $obj = $addin.Object
        if ($null -ne $obj) { $r.RestartNeeded = [bool]$obj.GetRestartNeeded(); $r.AutomationAnswered = $true }
    }
    catch { $r.Error = $_.Exception.GetType().Name + ': ' + $_.Exception.Message }
    finally {
        foreach ($o in @($obj, $addin, $addins, $app)) {
            if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    [pscustomobject]$r
}

$script:ResolvedOffice = Resolve-OfficeVersion
Say "Office hive: $script:ResolvedOffice"

$payload = $null
$payloadPath = Join-Path $PayloadRoot $ManifestFileName
if (Test-Path -LiteralPath $payloadPath) {
    $payload = Get-Content -LiteralPath $payloadPath -Raw | ConvertFrom-Json
    $shape = @(Test-PayloadManifestShape $payload)
    if ($shape.Count -gt 0) { throw ("The payload manifest $payloadPath is malformed: " + ($shape -join '; ')) }
    Say "Payload: commit $($payload.commit), version $($payload.version), built $($payload.builtUtc)"
}
elseif ($Execute) {
    throw "REFUSING: no $payloadPath. Build it on the host with Testbed/host/Publish-AddInPayload.ps1, copy AddIn.zip in with Testbed/host/Copy-ToGuest.ps1, and expand it to $PayloadRoot."
}
else {
    Say "No payload manifest at $payloadPath - the installed build cannot be tied to a commit."
}

if ($Execute) {
    Assert-Bitness
    Assert-Elevated
    Assert-InteractiveSession
    Assert-OutlookClosed

    $installer = Join-Path $PayloadRoot ([string]$payload.installer.file)
    if (-not (Test-Path -LiteralPath $installer)) { throw "REFUSING: the installer the manifest names is not at $installer." }
    $installerHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
    if ($installerHash -ne [string]$payload.installer.sha256) {
        throw "REFUSING: $installer is not the file the manifest pins (expected $($payload.installer.sha256), got $installerHash)."
    }
    Say "  installer $installer - SHA-256 matches the manifest"

    $interactionBefore = Get-InteractionFacts
    Say ("  index exclusion before: PreventIndexingOutlook={0} mapi rules=[{1}]" -f $interactionBefore.PreventIndexingOutlook, ($interactionBefore.MapiRules -join '; '))

    # ---- 3. The VSTO runtime -------------------------------------------------------------------
    Say ''
    Say '== The VSTO runtime =='
    $rt = Get-VstoRuntimeFacts
    if ((Compare-DottedVersion $rt.V4R $VstoRuntimeVersion) -ge 0) {
        Say "  already registered: v4R $($rt.V4R) - not reinstalling"
    }
    else {
        Say "  registered now: v4R '$($rt.V4R)', v4 '$($rt.V4)' - installing $VstoRuntimeVersion"
        if (-not (Test-Path -LiteralPath $VstoRuntimePath)) {
            throw "REFUSING: no VSTO runtime at $VstoRuntimePath. It is STAGED MEDIA (Testbed/MEDIA.md) - never downloaded here; the guest has no network. Copy it in with Testbed/host/Copy-ToGuest.ps1."
        }
        $rtHash = (Get-FileHash -LiteralPath $VstoRuntimePath -Algorithm SHA256).Hash
        $rtLength = (Get-Item -LiteralPath $VstoRuntimePath).Length
        if ($rtHash -ne $VstoRuntimeSha256.ToUpperInvariant() -or $rtLength -ne $VstoRuntimeBytes) {
            throw "REFUSING: $VstoRuntimePath is not the pinned redistributable ($rtHash, $rtLength bytes; expected $VstoRuntimeSha256, $VstoRuntimeBytes bytes)."
        }
        # Installer.iss's own switches for it: /q /norestart.
        $run = Invoke-Installer -FilePath $VstoRuntimePath -Arguments @('/q', '/norestart') -TimeoutMinutes $InstallTimeoutMinutes
        if ($run.TimedOut) { throw "The VSTO runtime installer did not finish within $InstallTimeoutMinutes minutes. Its logs are in $env:TEMP." }
        if ($null -eq $run.ExitCode) { throw 'The VSTO runtime installer exit code came back EMPTY - this script failed to read it, which is not the same as the install failing. Check v4R by hand.' }
        if ($run.ExitCode -ne 0 -and $run.ExitCode -ne 3010) { throw "The VSTO runtime installer exited $($run.ExitCode). Its logs are in $env:TEMP." }
        if ($run.ExitCode -eq 3010) { Say '  exit 3010: installed, and Windows wants a reboot. Not a failure; reboot before taking a checkpoint.' }
        $rt = Get-VstoRuntimeFacts
        if ((Compare-DottedVersion $rt.V4R $VstoRuntimeVersion) -lt 0) { throw "The VSTO runtime installer exited $($run.ExitCode) but v4R still reports '$($rt.V4R)'." }
        Say "  installed: v4R $($rt.V4R)"
    }

    # ---- 4. The product's installer ------------------------------------------------------------
    Say ''
    Say '== The OutlookAI installer, silent =='
    $setupLog = Join-Path (Split-Path -Parent $LogPath) 'install-addin-setup.log'
    $run = Invoke-Installer -FilePath $installer -Arguments (Get-InnoArguments -SetupLog $setupLog) -TimeoutMinutes $InstallTimeoutMinutes
    if ($run.TimedOut -or $null -eq $run.ExitCode -or $run.ExitCode -ne 0) {
        if (Test-Path -LiteralPath $setupLog) { foreach ($l in (Get-Content -LiteralPath $setupLog -Tail 25)) { Say "      | $l" } }
        throw "The installer exited '$($run.ExitCode)' (timed out: $($run.TimedOut)). Its log: $setupLog"
    }
    $installDir = Get-InstallDir
    if (-not $installDir) { throw "The installer exited 0 but wrote no HKCU\$AppKey\$InstallDirValue. Its log: $setupLog" }
    Say "  installed into $installDir (log: $setupLog)"

    # ---- 5. Trust ------------------------------------------------------------------------------
    Say ''
    Say '== Trust: the inclusion entry the trust prompt would have written =='
    $regValues = Get-HkcuValues $AddinRegistrationKey
    if (-not ($regValues -and $regValues['Manifest'])) { throw "The installer exited 0 but wrote no HKCU\$AddinRegistrationKey\Manifest." }
    $url = ConvertTo-InclusionUrl ([string]$regValues['Manifest'].Data)
    $key = Get-ManifestSigningKeyXml -ManifestXml ([System.IO.File]::ReadAllText((Join-Path $installDir 'OutlookAI.vsto')))
    if (-not (Test-SameRsaKey $key ([string]$payload.signing.publicKeyXml))) { throw 'REFUSING TO TRUST: the installed manifest is not signed by the key the payload manifest names.' }
    if (-not (Test-SameRsaKey $key (Get-CertificateKeyXml -Path (Join-Path $installDir 'OutlookAI.cer')))) { throw 'REFUSING TO TRUST: the installed manifest is not signed by the installed OutlookAI.cer.' }
    if (-not (Test-KeyXmlParsesLikeTheRuntime $key)) { throw 'REFUSING TO TRUST: the key does not parse with FromXmlString, so the runtime would reject the entry.' }
    $root = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($InclusionKey)
    try {
        $kept = $false
        foreach ($name in $root.GetSubKeyNames()) {
            $k = $root.OpenSubKey($name, $false)
            $entryUrl = $null; $entryKey = $null
            if ($null -ne $k) { try { $entryUrl = [string]$k.GetValue('Url'); $entryKey = [string]$k.GetValue('PublicKey') } finally { $k.Close() } }
            if (-not (Test-SameInclusionUrl $entryUrl $url)) { continue }
            if ((Test-SameRsaKey $entryKey $key) -and -not $kept) { $kept = $true; Say "  kept the existing entry $name - same URL, same key"; continue }
            $root.DeleteSubKeyTree($name)
            Say "  removed entry $name - same URL, a previous build's key"
        }
        if (-not $kept) {
            $name = [guid]::NewGuid().ToString()
            $entry = $root.CreateSubKey($name)
            try {
                $entry.SetValue('Url', $url, [Microsoft.Win32.RegistryValueKind]::String)
                $entry.SetValue('PublicKey', $key, [Microsoft.Win32.RegistryValueKind]::String)
            }
            finally { $entry.Close() }
            Say "  wrote HKCU\$InclusionKey\$name  Url=$url"
        }
    }
    finally { $root.Close() }

    # ---- 6. Make a failed load explain itself --------------------------------------------------
    [Environment]::SetEnvironmentVariable($LogAlertsName, '1', 'User')
    Set-Item -Path "Env:\$LogAlertsName" -Value '1'
    Say "  set $LogAlertsName=1 for this user"

    # ---- 7. The first run ----------------------------------------------------------------------
    Say ''
    Say '== First run: Outlook started once, headless, in a watchdogged child job =='
    $windowsBefore = Get-SessionWindows
    $startedUtc = [DateTime]::UtcNow
    $job = Start-Job -ScriptBlock $FirstRunJob -ArgumentList $startedUtc.Ticks, $FirstRunTimeoutSeconds, $TuningKey, $McpKey, $AddinName
    $done = Wait-Job -Job $job -Timeout ($FirstRunTimeoutSeconds + 120)
    $first = $null
    if ($done) {
        $first = Receive-Job -Job $job
        Remove-Job -Job $job -Force
    }
    else {
        Say "  *** NOTHING BACK WITHIN $($FirstRunTimeoutSeconds + 120) s - Outlook is most likely showing a MODAL DIALOG. On screen now:"
        $onScreen = @(Get-SessionWindows)
        foreach ($w in $onScreen) { Say "      $w" }
        if (($onScreen -join ' ').Contains('Customization Installer')) { Say '  That is the VSTO TRUST PROMPT (or its install-error box): the inclusion entry did not match what the runtime looked for.' }
        Stop-Job -Job $job; Remove-Job -Job $job -Force
        Say '  The job was stopped. Outlook was NOT touched - answer the dialog on the console, or restart the guest. Never taskkill it.'
    }

    $comProblems = @()
    if ($null -eq $first) { $comProblems += 'the first run did not finish; see above.' }
    else {
        Say "  COM start $($first.ComSeconds) s; MAPI $($first.MapiInit)"
        Say "  tuning state written after $($first.TuningAfterSeconds)s; registration reconcile after $($first.McpAfterSeconds)s"
        Say "  COMAddIns('$AddinName').Connect = $($first.Connect); the add-in answered GetRestartNeeded() = $($first.RestartNeeded)"
        if ($first.Error) { $comProblems += "the first run reported: $($first.Error)" }
        if ($null -eq $first.TuningAfterSeconds) { $comProblems += "the add-in did not write HKCU\$TuningKey within $FirstRunTimeoutSeconds s of Outlook starting." }
        if ($first.Connect -ne $true) { $comProblems += "Outlook reports the add-in as NOT connected." }
        if (-not $first.AutomationAnswered) { $comProblems += 'the add-in did not answer a call into it (COMAddIn.Object.GetRestartNeeded), so it is not demonstrably running.' }
    }
    $still = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($still.Count -gt 0) { Say "  Outlook is still running (pid $(($still | ForEach-Object { $_.Id }) -join ', ')), headless. Left as it is." }
    else { Say '  Outlook closed by itself once the last reference was released - gracefully; nothing quit or killed it.' }
    # A window that appeared during the run and is still up is a dialog somebody has to answer -
    # and on an unattended guest, the next step's hang.
    $newWindows = @(Get-SessionWindows | Where-Object { $windowsBefore -notcontains $_ })
    foreach ($w in $newWindows) { $comProblems += "a window appeared during the first run and is still on screen: $w - answer it on the console before the next step." }

    $interactionAfter = Get-InteractionFacts
    $changes = @(Compare-InteractionFacts $interactionBefore $interactionAfter)

    Say ''
    Say '== Verify =='
    $facts = Get-AddInFacts -RequireFresh $true -StartedUtc $startedUtc -Payload $payload
    $facts.ComProblems = $comProblems
    foreach ($c in $changes) {
        $facts.Notes += "INDEX EXCLUSION STATE CHANGED during this run: $c. The add-in writes nothing there, so this is Outlook's own start. Re-run Set-OutlookIndexingDisabled.ps1 -Verify before trusting this guest as unindexed."
    }
    if ($changes.Count -eq 0) { Say '  index exclusion state: UNCHANGED by this run (policy value and mapi crawl rules identical before and after)' }
    Write-TestReadBlock $facts.Tuning
    $result = Get-AddInVerdict $facts
    Write-Verdict $result
    exit $result.ExitCode
}

# -Verify
Say ''
Say '== Verify =='
$facts = Get-AddInFacts -RequireFresh $false -StartedUtc ([DateTime]::MinValue) -Payload $payload
$ix = Get-InteractionFacts
Say ("  index exclusion now: PreventIndexingOutlook={0} mapi rules=[{1}]  (read only; the add-in writes neither)" -f $ix.PreventIndexingOutlook, ($ix.MapiRules -join '; '))
if ($WithOutlook) {
    Assert-Bitness
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq [System.Diagnostics.Process]::GetCurrentProcess().SessionId })
    if ($running.Count -eq 0) {
        $facts.Notes += '-WithOutlook: no Outlook is running in this session, and a verify never starts one. The COM half was not checked.'
    }
    else {
        $job = Start-Job -ScriptBlock $AttachJob -ArgumentList $AddinName
        if (Wait-Job -Job $job -Timeout 90) {
            $att = Receive-Job -Job $job
            Say "  running Outlook: Connect=$($att.Connect), the add-in answered GetRestartNeeded() = $($att.RestartNeeded)"
            if ($att.Error) { $facts.ComProblems += "-WithOutlook: $($att.Error)" }
            elseif ($att.Connect -ne $true -or -not $att.AutomationAnswered) { $facts.ComProblems += 'the running Outlook does not have the add-in connected and answering.' }
        }
        else {
            Stop-Job -Job $job
            $facts.ComProblems += '-WithOutlook: the running Outlook did not answer within 90 s - it may be showing a dialog.'
        }
        Remove-Job -Job $job -Force
    }
}
Write-TestReadBlock $facts.Tuning
$result = Get-AddInVerdict $facts
Write-Verdict $result
exit $result.ExitCode
