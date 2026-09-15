#Requires -Version 5.1
<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it: the machine it was written on is the maintainer's
    workstation, with a real Outlook profile, real mail and a real Windows Search catalog on it,
    and this script changes what Windows Search indexes. Verified by PARSING only - the same
    check .github/scripts/check-testbed-references.ps1 applies to every script under Testbed/.
    Nothing below has run anywhere; no crawl scope was altered, no service restarted and no
    catalog rebuilt to write it.

    What WAS measured, read-only, on that workstation, and is therefore fact rather than
    inference: the registry shape of the Outlook crawl-scope rule, the COM registration of the
    Search Crawl Scope Manager, and the host of that COM class. Those measurements are in
    .work/unindexed-guest.md with their evidence class. Replace this banner with what the script
    actually did once it has run on a guest, and say which of the four verdicts it printed.

.SYNOPSIS
    Takes a testbed guest's Outlook OUT of the Windows Search index - and then proves the store
    is unindexed rather than reporting that a registry value was written.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    NEVER run this on the maintainer's workstation. It refuses there; see THE GUARD below.

    WHY THIS EXISTS. Docs/live-tier-on-the-vm.md section 1.3. The testbed is two guests:
    OutlookAI-Indexed carries Corpus A and must be INDEXED, OutlookAI-Unindexed carries Corpus B
    and must be NOT INDEXED. Half the live tier wants a populated index; the other half wants a
    store with NO INDEX FRONTIER, which is what drives the seven-day fallback window and the
    sweep and frame measurements. Both guests are built from the same answer file and both come
    up indexed, so until something does this, the two halves measure the same machine and the
    degraded half is measuring nothing.

    WHAT "UNINDEXED" HAS TO MEAN HERE, AND WHAT IT MUST NOT MEAN. The product distinguishes
    "the index answered and holds nothing for this store" from "the index could not be reached",
    and they are different code paths with different payloads:

      * NO ROWS FOR THE STORE is the wanted state. The scoped frontier probe RUNS and returns
        nothing, so MailService opens EmptyIndexSweepWindow (7 days), raises
        FreshMerge.GapNoIndexFrontier, sets search.indexFrontierMissing and names the store in
        sweep.storesWithoutIndex. outlook_health reports index.perStore[] with a row for that
        store carrying inLocalIndex:false and no newestIndexedUtc.

      * INDEX UNREACHABLE is a DIFFERENT machine. outlook_health sets
        index.provider = "unavailable: <exception>" and index.perStore stays NULL - and because
        the server serialises with JsonIgnoreCondition.WhenWritingNull, the per-store block is
        ABSENT FROM THE PAYLOAD ENTIRELY. The instrument section 1.1 tells you to read is the
        one thing that state removes.

    SO THIS SCRIPT NEVER DISABLES THE WINDOWS SEARCH SERVICE, and neither should anything else
    on these guests. The reasoning is in .work/unindexed-guest.md; the short form is three
    measured facts. (1) CLSID_CSearchManager's AppID names LocalService = WSearch, so the crawl
    scope manager - the thing that both SETS and READS BACK the scope - lives inside the service
    and is unreachable when it is stopped. (2) MailService's non-exhaustive search path does not
    wrap its index calls in a catch, and no test in the suite covers a throwing index client, so
    what `search` does on a machine with no catalog is unspecified. (3)
    T3/Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning
    asserts index.wSearchStartMode == "automatic" and declares only Requires=AddInRegistry, so
    it does not filter out on an unindexed guest - it just fails.

    WHAT IT DOES INSTEAD. It takes the Outlook MAPI scope out of the crawl scope and leaves the
    indexer running, so the catalog stays alive, keeps answering, and simply holds no Outlook
    rows. Three layers, ordered by how well attested each is:

      1. POLICY (documented): HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\
         PreventIndexingOutlook = 1 (DWORD), the Group Policy "Prevent indexing Microsoft Office
         Outlook" from Search.admx. Microsoft Support prescribes this exact registry edit. It is
         the only layer Microsoft documents as a supported switch, and it is the one that is
         DURABLE BY CONSTRUCTION: the Crawl Scope Manager's documented rule precedence is GROUP
         POLICY > user rules > default rules, and RevertToDefaultScopes is documented not to
         remove policy rules - so a policy exclusion outranks any user rule Outlook re-creates
         on its next start, which is the single biggest risk to layer 2 on its own.
         TWO CAVEATS, both stated because they are real. It is class="Machine": HKLM, no
         per-user variant, so it hits EVERY Windows account on the guest - fine here, where
         there is one. And whether it REMOVES rows already indexed or merely stops adding new
         ones is not documented anywhere, which is why -RebuildCatalog exists below.
      2. THE CRAWL-SCOPE RULE (reverse-engineered): the mapi16://{SID}/ rule under the crawl
         scope manager's WorkingSetRules. MEASURED on the maintainer's workstation - the rule
         exists, is a USER rule (Default = 0), and carries Include = 1. Setting Include = 0 is
         the registry form of untickng the account in Indexing Options. Nothing Microsoft
         publishes describes these keys; they are documented by the Crawl Scope Manager COM API
         instead, which is why layer 3 exists.
      3. A SERVICE RESTART (necessary, not sufficient): the gatherer caches its scope, so a raw
         registry edit is not read until the service restarts. RESTART, never disable.

    AND THE PART THAT IS NOT A SETTING AT ALL: PURGING WHAT WAS ALREADY CRAWLED. If Outlook ran
    on this guest before this script did, the catalog already holds rows for the corpus, and
    excluding the scope afterwards does not reliably delete them. -RebuildCatalog handles that
    by throwing the whole catalog away and letting it re-crawl WITHOUT the Outlook scope; it is
    off by default because it is expensive and because -Verify tells you whether you need it.

    THE CHEAPEST ANSWER IS ORDERING, AND IT BELONGS IN THE RUNBOOK RATHER THAN IN A FLAG. Run
    this BEFORE step 8 of Testbed/README.md section 1 - before Build-Corpus.ps1 - and no Outlook
    row is ever created, so there is nothing to purge and nothing to wait for. Run it afterwards
    and you own both problems.

    WHAT -Verify ACTUALLY PROVES, AND WHY IT IS NOT A REGISTRY READ-BACK.
    Testbed/guest/Set-OfficeFirstRunSuppressed.ps1 ends by saying registry values prove that
    something was written, not that the application read them. The same is true here and worse,
    because the thing that has to honour the value is a service with its own cache. So -Verify
    asks the INDEX, not the registry, using the same OLE DB provider and the same statements the
    shipped server uses (OutlookAI.Core/IndexSearch/WsSqlBuilder.cs):

      * a CONTROL probe - "does the catalog hold any row at all" - which separates "no Outlook
        rows" from "no catalog";
      * a MAPI probe - "does the catalog hold any row under mapi16://{SID}/" - which is the
        actual question;
      * a MAIL probe - "does the catalog hold any row of System.Kind='email'" - which is the
        one the product's own frontier probe asks, and the one that decides whether a store has
        a frontier. It carries no SCOPE, deliberately: it is the SAFETY NET under the MAPI
        probe. That probe's predicate is a URL built at runtime, and a SCOPE the provider
        matches against nothing would look identical to a store with nothing in it. On a guest
        the only mail is Outlook's, so a mail probe that returns rows disproves "unindexed"
        whatever the scoped one says. UNINDEXED needs BOTH to be empty, and every probe to have
        reached the catalog at all.

    Both probes are taken TWICE, -SettleMinutes apart, because a single reading cannot tell
    "not indexed" from "not indexed yet". Four verdicts come out, and the third and fourth are
    the ones that exist to stop a believed-unindexed guest from quietly indexing:

      INDEXED        control yes, mapi yes, mail yes, stable across both readings.
      UNINDEXED      control yes, mapi no, mail no, on BOTH readings. The wanted state.
      SETTLING       the two readings disagree, or mapi rows are present while the scope says
                     excluded. Either a crawl is still running or a purge is. NOT an answer -
                     wait and re-run, and use -RebuildCatalog if it will not converge.
      NO-INDEXER     ANY probe failed to reach the catalog, or the service is not running. This
                     is the state
                     that LOOKS like success and is not: it is a machine with no index at all,
                     not a store with no index frontier. Reported as a FAILURE.

    IDEMPOTENT. Run it twice and the second run reports "already" for every value and still runs
    the full verification. -Enable is the inverse and takes the guest back to indexed, so a
    checkpoint can be re-taken either way.

    WHAT IT NEVER DOES. It never starts Outlook, never creates an Outlook COM object, never
    touches MAPI, never reads or writes an Outlook profile registry key, and never touches a
    mail item in any store. It never stops or disables the Windows Search service. Everything it
    writes is under HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search and the crawl scope
    manager's own keys, and everything it reads is a registry value, a service state, or a
    SELECT against the local catalog.

    THE GUARD, AND IT HAS TWO AXES ON PURPOSE. Same rule and same fail-closed shape as
    Assert-TestbedGuest in Testbed/guest/OutlookMapiInterop.ps1 - restated here rather than
    dot-sourced, for the reason Install-MailSink.ps1 gives: that file compiles Extended MAPI
    interop with Add-Type at dot-source time, and a script that has nothing to do with MAPI must
    not be able to die with a MAPI compile error. The second axis is the computer name, because
    the consequence here is different from a destroyed profile: excluding Outlook from the index
    on a machine with real mail on it silently breaks that person's search, and it is the kind
    of damage nobody notices for weeks. Both axes must pass, and the only way past either is to
    name the value you mean.

    STARTING STATE IT EXPECTS.
      * A guest checkpoint where Windows is installed and you are logged on as -ExpectedUser.
      * An ELEVATED session. The policy key and the crawl scope keys are HKLM, and restarting a
        service needs it; the script asserts elevation rather than failing halfway with an
        access-denied nobody can interpret.
      * A 64-BIT PowerShell. The Search.CollatorDSO provider needs an x64 host
        (Docs/live-tier-on-the-vm.md section 2.3), so a 32-bit session would report NO-INDEXER
        on a perfectly healthy machine.
      * Outlook not running is preferred but not required - nothing here touches it.

.PARAMETER ExpectedUser
    Accounts this script is allowed to run as. The guests autologon as vmadmin.

.PARAMETER ExpectedComputerNamePrefix
    Computer-name prefix this script is allowed to run on. Testbed/host/New-AnswerFile.ps1
    derives a guest's name by replacing 'OutlookAI-' with 'OAI-', so every guest built from the
    answer file matches. Pass your own if you named a guest something else; do not widen it to
    an empty string.

.PARAMETER Execute
    Write. Without it, prints what it would set and changes nothing.

.PARAMETER Enable
    The inverse: put Outlook back IN the index. Needs -Execute to write. Expect the verification
    that follows it to report SETTLING and exit non-zero - re-including a scope does not crawl it,
    it only lets the gatherer start, and 20,000 items do not appear in the time this script waits.
    Re-run -Verify later; the exit code is honest about not knowing yet rather than reporting a
    success on the strength of having written a value.

.PARAMETER Verify
    Probe the catalog and report the actual state. Safe at any time; writes nothing.

.PARAMETER RebuildCatalog
    Throw the whole Windows Search catalog away and let it re-crawl. Only meaningful with
    -Execute, and only needed when Outlook was already crawled on this guest. Expensive.

.PARAMETER SettleMinutes
    Minutes between the two verification readings. The default is deliberately long enough that a
    crawl in progress moves between them, which means -Verify BLOCKS for that long; the guest runs
    it through a scheduled task writing to a file, so that is a cost rather than a problem.
    `-SettleMinutes 0` takes one reading instead of two, and gives up the only thing that
    distinguishes "not indexed" from "not indexed yet" - so a 0 run can report UNINDEXED over a
    crawl that has not started. Use it to look, never to conclude.

.PARAMETER LogPath
    Transcript file. Everything printed also lands here.

.EXAMPLE
    .\Set-OutlookIndexingDisabled.ps1
    .\Set-OutlookIndexingDisabled.ps1 -Execute
    .\Set-OutlookIndexingDisabled.ps1 -Verify
    .\Set-OutlookIndexingDisabled.ps1 -Execute -RebuildCatalog
    .\Set-OutlookIndexingDisabled.ps1 -Enable -Execute
#>
[CmdletBinding()]
param(
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [switch]   $Execute,
    [switch]   $Enable,
    [switch]   $Verify,
    [switch]   $RebuildCatalog,
    [int]      $SettleMinutes = 10,
    [string]   $LogPath = 'C:\OutlookAI-Q5\set-outlook-indexing.log'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Everything this script names, in one place, so the blast radius is readable without reading
# the code. EVIDENCE is the honest label on each: DOC means Microsoft documents this value;
# MEASURED means it was read off a real machine on 2026-09-16; INFERRED means neither.
# ---------------------------------------------------------------------------------------------
$SearchPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search'
$SearchPolicyName = 'PreventIndexingOutlook'   # EVIDENCE: DOC (Search.admx). Absent by default.

$CsmRoot = 'HKLM:\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex'
$WorkingSetRules = Join-Path $CsmRoot 'WorkingSetRules'   # EVIDENCE: MEASURED
$SearchRoots = Join-Path $CsmRoot 'SearchRoots'           # EVIDENCE: MEASURED

$SearchSetupKey = 'HKLM:\SOFTWARE\Microsoft\Windows Search'
$SetupCompletedName = 'SetupCompletedSuccessfully'        # EVIDENCE: INFERRED as a rebuild trigger

$WSearchServiceKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\WSearch'

# The connection string the shipped server uses, character for character
# (OutlookAI.Core/IndexSearch/IndexClients.cs, OleDbIndexClient.ConnectionString). Using the
# same one is the point: a probe that connects differently proves something about a different
# client.
$CollatorConnectionString = "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

# The statements, matching OutlookAI.Core/IndexSearch/WsSqlBuilder.cs. Kept as literals rather
# than built, because the whole value of this probe is that it asks what the product asks.
$ProbeControlSql = 'SELECT TOP 1 System.ItemUrl FROM SystemIndex'
$ProbeMailSql = "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE System.Kind='email'"

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

# ---------------------------------------------------------------------------------------------
# THE GUARD. FIRST CALL IN EVERY PATH THAT WRITES - and in -Verify too, because a verification
# run opens a catalog connection and reads a mailbox-shaped URL list, and neither belongs on the
# maintainer's machine either.
#
# Two axes, both fail-closed, neither silenceable by a bare flag. The account, because the
# guests autologon as vmadmin and no other machine in this project does. The computer name,
# because the damage this particular script can do to a real machine - silently emptying its
# mail search - is not visible for weeks, and one axis is not enough for that.
# ---------------------------------------------------------------------------------------------
function Assert-TestbedGuestLocal {
    $who = $env:USERNAME
    $userOk = $false
    foreach ($candidate in $ExpectedUser) {
        if ($who -eq $candidate) { $userOk = $true }
    }

    $machine = $env:COMPUTERNAME
    $machineOk = $false
    if ($ExpectedComputerNamePrefix -and $machine -and
        $machine.StartsWith($ExpectedComputerNamePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        $machineOk = $true
    }

    if ($userOk -and $machineOk) { return }

    throw @"
REFUSING TO RUN.

  logged on as : '$who'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$machine'      (must start with: '$ExpectedComputerNamePrefix')

This script takes Outlook OUT of the Windows Search index. On a machine with real mail on it
that silently empties mail search - for every account, in Outlook's own search box as well as
through this project's tools - and nothing announces it. It is not a mistake anyone notices on
the day they make it.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2), and
Testbed/host/New-AnswerFile.ps1 derives their computer name by replacing 'OutlookAI-' with
'OAI-', so a guest built from the answer file matches both axes. If you named a guest something
else, say so:

    -ExpectedUser <username> -ExpectedComputerNamePrefix <prefix>

Do not 'fix' this by widening either default. The defaults are the guard.
"@
}

function Assert-Elevated {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw @"
REFUSING TO RUN: this session is not elevated.

Everything this script writes is HKLM - the Windows Search policy key and the crawl scope
manager's own keys - and it restarts a service. An unelevated run would fail somewhere in the
middle, having written some values and not others, which is the one outcome worse than not
running at all.
"@
    }
}

function Assert-OutlookClosed {
    <#
        REFUSES while Outlook is running, and this is the one refusal here that protects the
        GUEST rather than the measurement.

        EVIDENCE: MS-DOC, verbatim, from the MAPI dev article on wrapped PSTs and indexing:
        "The PST provider is very sensitive to the indexing state changing while the PST is
        open. If the state changes, the PST may end up kicking off an installer to repair
        Outlook."
        https://learn.microsoft.com/en-us/archive/blogs/stephen_griffin/wrapped-pst-and-indexing

        This script changes exactly that state and then restarts WSearch, so running it against
        a guest with Outlook open is the documented way to trigger an Office repair. On a
        testbed guest that is not a crash, it is worse: the repair runs unattended, the guest's
        Office state stops matching the one the build script produced, and the checkpoint the
        whole testbed depends on no longer describes the machine.

        Naming the PSTs matters. "Close Outlook" is easy to satisfy by killing it, and the
        project's safety rules forbid that outright (never taskkill OUTLOOK.EXE - a killed
        Outlook is how a PST gets left mid-write). So this says WHICH process and leaves the
        closing to the caller.
    #>
    $outlook = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($outlook.Count -eq 0) {
        return
    }

    $pids = ($outlook | ForEach-Object { $_.Id }) -join ', '
    throw @"
REFUSING TO RUN: Outlook is running on this guest (PID $pids).

Microsoft documents that the PST provider is "very sensitive to the indexing state changing
while the PST is open", and that if the state changes "the PST may end up kicking off an
installer to repair Outlook". This script changes precisely that state and then restarts
WSearch, so running it now is the documented way to start an unattended Office repair on a
guest whose Office state is supposed to match its checkpoint.

Close Outlook first - GRACEFULLY. Do not taskkill it: the project's mailbox-safety rules
forbid that, and a PST left mid-write is a worse outcome than an unindexed guest. Quit Outlook
(or restart the guest, which is the proven way to get a clean Outlook here), then run this
again.

This refusal does not apply to -Verify, which only reads.
"@
}

function Assert-Bitness {
    if (-not [Environment]::Is64BitProcess) {
        throw @"
REFUSING TO RUN: this is a 32-bit PowerShell.

The Search.CollatorDSO provider needs an x64 host (Docs/live-tier-on-the-vm.md section 2.3), so
every catalog probe below would fail and this script would report NO-INDEXER on a machine whose
index is perfectly healthy - a false negative that reads exactly like the failure it exists to
detect. Start Windows PowerShell from C:\Windows\System32\WindowsPowerShell\v1.0.
"@
    }
}

function Get-CurrentUserSid {
    return ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
}

# The scope URL section 1.1 of the runbook names, and the shape MEASURED in the crawl scope
# manager on a real machine: one URL per Windows account, SID in braces, trailing slash.
function Get-MapiScopeUrl {
    return 'mapi16://{' + (Get-CurrentUserSid) + '}/'
}

# ---------------------------------------------------------------------------------------------
# Reading the crawl scope. Read-only, and it never guesses: a rule that is not there is reported
# as absent rather than as excluded, because those are different states and only one of them
# survives an Outlook restart.
# ---------------------------------------------------------------------------------------------
function Get-MapiScopeRules {
    $wanted = Get-MapiScopeUrl
    $found = @()
    foreach ($container in @($WorkingSetRules, $SearchRoots)) {
        if (-not (Test-Path -LiteralPath $container)) { continue }
        foreach ($child in Get-ChildItem -LiteralPath $container -ErrorAction SilentlyContinue) {
            $values = Get-ItemProperty -LiteralPath $child.PSPath -ErrorAction SilentlyContinue
            if (-not $values) { continue }
            $url = $values.URL
            if (-not $url) { continue }
            if ($url -notlike 'mapi*') { continue }
            $found += [pscustomobject]@{
                Container = Split-Path -Leaf $container
                KeyPath   = $child.PSPath
                KeyName   = $child.PSChildName
                Url       = $url
                IsThisUser = ($url -eq $wanted)
                Include   = $values.Include
                Suppress  = $values.Suppress
                Default   = $values.Default
                Policy    = $values.Policy
            }
        }
    }

    return @($found)
}

function Get-PolicyValue {
    if (-not (Test-Path -LiteralPath $SearchPolicyKey)) { return $null }
    $p = Get-ItemProperty -LiteralPath $SearchPolicyKey -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    return $p.$SearchPolicyName
}

function Get-WSearchState {
    $start = $null
    if (Test-Path -LiteralPath $WSearchServiceKey) {
        $k = Get-ItemProperty -LiteralPath $WSearchServiceKey -ErrorAction SilentlyContinue
        if ($k) { $start = $k.Start }
    }

    $startMode = 'unknown'
    if ($start -eq 2) { $startMode = 'automatic' }
    elseif ($start -eq 3) { $startMode = 'manual' }
    elseif ($start -eq 4) { $startMode = 'disabled' }
    elseif ($null -ne $start) { $startMode = "other($start)" }

    $status = 'absent'
    $svc = Get-Service -Name 'WSearch' -ErrorAction SilentlyContinue
    if ($svc) { $status = [string]$svc.Status }

    $indexerRunning = $false
    if (@(Get-Process -Name 'SearchIndexer' -ErrorAction SilentlyContinue).Count -gt 0) { $indexerRunning = $true }

    return [pscustomobject]@{
        StartValue     = $start
        StartMode      = $startMode
        Status         = $status
        IndexerRunning = $indexerRunning
    }
}

# ---------------------------------------------------------------------------------------------
# The catalog probe. Late-bound ADODB over Search.CollatorDSO - the same fallback path the
# shipped server carries (OutlookAI.Core/IndexSearch/IndexClients.cs, AdodbIndexClient), chosen
# here because it is the one a Windows PowerShell 5.1 session can actually reach: ADODB is an
# automation object, and the Search Crawl Scope Manager is not (MEASURED: its interfaces carry a
# custom proxy/stub and no type library, so there is nothing to late-bind to).
#
# Returns $null when the catalog could not be reached AT ALL, which is a different answer from
# zero rows and is never conflated with it.
# ---------------------------------------------------------------------------------------------
function Invoke-IndexProbe {
    param([Parameter(Mandatory = $true)][string] $Sql)

    $connection = $null
    $recordset = $null
    try {
        $connection = New-Object -ComObject 'ADODB.Connection'
        $connection.Open($CollatorConnectionString)
        $recordset = New-Object -ComObject 'ADODB.Recordset'
        $recordset.Open($Sql, $connection)

        $rows = @()
        while (-not $recordset.EOF -and $rows.Count -lt 1) {
            $rows += [string]$recordset.Fields.Item(0).Value
            $recordset.MoveNext()
        }

        return [pscustomobject]@{ Reached = $true; RowCount = $rows.Count; Error = $null }
    }
    catch {
        return [pscustomobject]@{ Reached = $false; RowCount = 0; Error = $_.Exception.Message }
    }
    finally {
        foreach ($o in @($recordset, $connection)) {
            if ($null -eq $o) { continue }
            try { if ($o.State -ne 0) { $o.Close() } } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
        }
    }
}

# The MAPI probe is built per-reading because it carries this account's SID. SCOPE is the
# recursive store-scope predicate the product uses; the quoting matches WsSqlBuilder's.
function Get-MapiProbeSql {
    return "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + (Get-MapiScopeUrl) + "'"
}

function Read-IndexState {
    $control = Invoke-IndexProbe -Sql $ProbeControlSql
    $mapi = Invoke-IndexProbe -Sql (Get-MapiProbeSql)
    $mail = Invoke-IndexProbe -Sql $ProbeMailSql
    return [pscustomobject]@{
        TakenAt = Get-Date
        Control = $control
        Mapi    = $mapi
        Mail    = $mail
    }
}

function Show-IndexState {
    param($State, [string] $Label)
    Say ("  {0}: catalog reachable={1} anyRow={2} mapiRows={3} mailRows={4}" -f
        $Label,
        $State.Control.Reached,
        $State.Control.RowCount,
        $State.Mapi.RowCount,
        $State.Mail.RowCount)
    foreach ($pair in @(@('control', $State.Control), @('mapi', $State.Mapi), @('mail', $State.Mail))) {
        if (-not $pair[1].Reached) { Say ("    {0} probe FAILED: {1}" -f $pair[0], $pair[1].Error) }
    }
}

# ---------------------------------------------------------------------------------------------
# What gets written. Each entry says what it is and how well attested it is, because the three
# layers are NOT equally well founded and a reader has to be able to tell.
# ---------------------------------------------------------------------------------------------
function Write-PolicyValue {
    param([int] $Value)
    if (-not (Test-Path -LiteralPath $SearchPolicyKey)) {
        New-Item -Path $SearchPolicyKey -Force | Out-Null
        Say "  created $SearchPolicyKey"
    }

    $current = Get-PolicyValue
    if ($current -eq $Value) {
        Say "  already $SearchPolicyName = $Value"
        return
    }

    New-ItemProperty -Path $SearchPolicyKey -Name $SearchPolicyName -PropertyType DWord -Value $Value -Force | Out-Null
    Say "  set $SearchPolicyName = $Value"
}

function Set-MapiScopeRuleInclude {
    param([int] $Include)
    $rules = @(Get-MapiScopeRules | Where-Object { $_.IsThisUser -and $_.Container -eq 'WorkingSetRules' })
    if ($rules.Count -eq 0) {
        Say "  no mapi rule for this account under WorkingSetRules - nothing to change"
        Say "  (that is NOT the same as excluded: a rule Outlook has not created yet is a rule"
        Say "   Outlook can create on its next start, which is why the policy layer above exists)"
        return
    }

    foreach ($rule in $rules) {
        if ($rule.Include -eq $Include) {
            Say ("  already Include = {0} on {1}\{2}" -f $Include, $rule.Container, $rule.KeyName)
            continue
        }

        New-ItemProperty -Path $rule.KeyPath -Name 'Include' -PropertyType DWord -Value $Include -Force | Out-Null
        Say ("  set Include = {0} on {1}\{2}  ({3})" -f $Include, $rule.Container, $rule.KeyName, $rule.Url)
    }
}

function Restart-SearchService {
    $state = Get-WSearchState
    if ($state.StartMode -eq 'disabled') {
        throw @"
REFUSING TO CONTINUE: the Windows Search service is DISABLED on this guest.

That is not the state this testbed wants and this script will not work around it. A machine with
no indexer is not a store with no index frontier: outlook_health reports
index.provider = "unavailable: ..." and omits index.perStore entirely, which is the one
instrument Docs/live-tier-on-the-vm.md section 1.1 tells you to read. Set WSearch back to
Automatic and start it, then run this again.
"@
    }

    Say '  restarting WSearch so the gatherer re-reads its scope (restart, never disable)'
    Restart-Service -Name 'WSearch' -Force
    Say ('  WSearch is now {0}' -f (Get-Service -Name 'WSearch').Status)
}

function Invoke-CatalogRebuild {
    Say '  requesting a full catalog rebuild'
    Say '  EVIDENCE: INFERRED. Setting SetupCompletedSuccessfully = 0 and restarting the service'
    Say '  is the community recipe for the Indexing Options "Rebuild" button. Microsoft documents'
    Say '  the button, not the value. It is used here because the alternative - the documented'
    Say '  ISearchCatalogManager::Reset - is not reachable from PowerShell 5.1 without declaring'
    Say '  COM vtables by hand (MEASURED: no type library, custom proxy/stub, so no late binding).'
    New-ItemProperty -Path $SearchSetupKey -Name $SetupCompletedName -PropertyType DWord -Value 0 -Force | Out-Null
    Restart-SearchService
    Say '  the catalog is now rebuilding. It will re-crawl everything EXCEPT what the scope now'
    Say '  excludes, which is the point of doing it in this order. Expect this to take a while;'
    Say '  -Verify is what tells you it has finished, not a clock.'
}

# ---------------------------------------------------------------------------------------------
# The verdict. Four states, and the two that are NOT answers are named as loudly as the two that
# are - a guest believed unindexed while it is quietly still crawling produces measurements that
# look fine and mean nothing.
# ---------------------------------------------------------------------------------------------
function Get-Verdict {
    param($First, $Second, $Rules, $PolicyValue, $Service)

    # EVERY probe must have REACHED the catalog, not just the control one. A probe that threw
    # reports RowCount = 0, and zero rows is the evidence UNINDEXED rests on - so treating a
    # failed probe as an empty result is the one way this function could hand back a false pass.
    # It is not hypothetical: the SCOPE predicate carries a URL built at runtime, and a
    # provider that rejects the syntax would look exactly like a store with nothing in it.
    foreach ($reading in @($First, $Second)) {
        foreach ($probe in @($reading.Control, $reading.Mapi, $reading.Mail)) {
            if (-not $probe.Reached) { return 'NO-INDEXER' }
        }
    }

    if ($Service.StartMode -eq 'disabled' -or -not $Service.IndexerRunning) { return 'NO-INDEXER' }
    if ($First.Control.RowCount -eq 0 -and $Second.Control.RowCount -eq 0) {
        # The catalog answered and holds nothing at all - it is rebuilding, or it is empty.
        # Reporting UNINDEXED here would be the exact false pass this verdict exists to stop.
        return 'SETTLING'
    }

    $mapiFirst = ($First.Mapi.RowCount -gt 0)
    $mapiSecond = ($Second.Mapi.RowCount -gt 0)
    $mailFirst = ($First.Mail.RowCount -gt 0)
    $mailSecond = ($Second.Mail.RowCount -gt 0)

    if ($mapiFirst -ne $mapiSecond -or $mailFirst -ne $mailSecond) { return 'SETTLING' }

    $excluded = $false
    if ($PolicyValue -eq 1) { $excluded = $true }
    foreach ($r in $Rules) { if ($r.IsThisUser -and $r.Container -eq 'WorkingSetRules' -and $r.Include -eq 0) { $excluded = $true } }

    if (-not $mapiFirst -and -not $mailFirst) {
        if ($excluded) { return 'UNINDEXED' }
        # No Outlook rows, and nothing says Outlook is excluded. That is a guest that has not
        # been crawled YET, not one that will not be.
        return 'SETTLING'
    }

    if ($excluded) {
        # Rows present under a scope that is excluded: the exclusion landed after the crawl did,
        # and the old rows have not been purged.
        return 'SETTLING'
    }

    return 'INDEXED'
}

# =============================================================================================
# Main
# =============================================================================================
Assert-Bitness

if (-not ($Execute -or $Verify -or $Enable)) {
    Say 'DRY RUN. Nothing is written. This is what -Execute would do:'
    Say ''
    Say "  1. $SearchPolicyKey\$SearchPolicyName = 1        [DOC]       policy: prevent indexing Outlook"
    Say "  2. Include = 0 on the mapi16://{SID}/ rule under WorkingSetRules  [MEASURED]  the Indexing Options tick"
    Say '  3. restart WSearch (never disable it)                             [INFERRED]  the gatherer caches its scope'
    Say '  4. with -RebuildCatalog: SetupCompletedSuccessfully = 0 + restart [INFERRED]  purge what was already crawled'
    Say ''
    Say '  and then the same verification -Verify runs on its own.'
    Say ''
    Say 'Re-run with -Execute to write, -Verify to probe only, or -Enable -Execute to reverse it.'
    Say ''
    Say 'ORDERING BEATS EVERY FLAG HERE. Run this BEFORE Build-Corpus.ps1 (step 8 of'
    Say 'Testbed/README.md section 1) and no Outlook row is ever created, so there is nothing to'
    Say 'purge and nothing to wait for. Run it afterwards and you need -RebuildCatalog and'
    Say 'patience.'
    return
}

Assert-TestbedGuestLocal

if ($Execute -or $Enable) {
    Assert-Elevated
    Assert-OutlookClosed

    $service = Get-WSearchState
    Say ('== Windows Search: startMode={0} status={1} indexerRunning={2} ==' -f
        $service.StartMode, $service.Status, $service.IndexerRunning)

    if ($Enable) {
        Say '== Execute (-Enable): putting Outlook BACK in the index =='
        Write-PolicyValue -Value 0
        Set-MapiScopeRuleInclude -Include 1
    }
    else {
        Say '== Execute: taking Outlook OUT of the index =='
        Write-PolicyValue -Value 1
        Set-MapiScopeRuleInclude -Include 0
    }

    if ($RebuildCatalog) { Invoke-CatalogRebuild }
    else { Restart-SearchService }
}

# -------------------------------------------------------------------------------------------
# Verification. It runs after -Execute as well as on its own, because a script that reports
# what it set has reported nothing: the service caches its scope, the gatherer decides when to
# act, and only the catalog knows what is actually in it.
# -------------------------------------------------------------------------------------------
Assert-TestbedGuestLocal

Say ''
Say '== Verify =='

$service = Get-WSearchState
Say ('  WSearch: startMode={0} status={1} indexerRunning={2}' -f
    $service.StartMode, $service.Status, $service.IndexerRunning)

$policy = Get-PolicyValue
if ($null -eq $policy) { Say "  policy ${SearchPolicyName}: ABSENT" }
else { Say "  policy ${SearchPolicyName}: $policy" }

$rules = Get-MapiScopeRules
if ($rules.Count -eq 0) {
    Say '  crawl scope: NO mapi rule of any kind. Outlook has never registered a scope here, or'
    Say '               something removed it. Absent is not the same as excluded - Outlook can'
    Say '               create one on its next start unless the policy above forbids it.'
}
else {
    foreach ($r in $rules) {
        Say ('  crawl scope: {0}\{1}  include={2} suppress={3} default={4} policy={5} thisAccount={6}' -f
            $r.Container, $r.KeyName, $r.Include, $r.Suppress, $r.Default, $r.Policy, $r.IsThisUser)
    }
}

Say ('  scope URL for this account: {0}' -f (Get-MapiScopeUrl))
Say ''
Say '  Probing the CATALOG, not the registry. Two readings, so "not indexed" can be told apart'
Say ('  from "not indexed yet". Second reading in {0} minute(s).' -f $SettleMinutes)

$first = Read-IndexState
Show-IndexState -State $first -Label 'reading 1'

if ($SettleMinutes -gt 0) { Start-Sleep -Seconds ($SettleMinutes * 60) }
$second = Read-IndexState
Show-IndexState -State $second -Label 'reading 2'

$verdict = Get-Verdict -First $first -Second $second -Rules $rules -PolicyValue $policy -Service $service

Say ''
Say "  VERDICT: $verdict"
Say ''

switch ($verdict) {
    'UNINDEXED' {
        Say '  The catalog is alive and answering, and holds no Outlook row for this account.'
        Say '  That is the shape the degraded tier wants: the scoped frontier probe RUNS and'
        Say '  returns nothing, so search opens the seven-day fallback window, raises'
        Say '  no_index_frontier and names the store in sweep.storesWithoutIndex.'
        Say ''
        Say '  CONFIRM IT THROUGH THE PRODUCT, which is the check that actually counts. Run'
        Say '  outlook_health and read index.perStore[]: every store on this guest must appear'
        Say '  with inLocalIndex:false and no newestIndexedUtc, index.provider must NOT start'
        Say '  with "unavailable", and index.wSearchStartMode must still say "automatic".'
        Say '  A payload with no index.perStore block at all is the NO-INDEXER state wearing'
        Say '  this one''s clothes.'
    }
    'INDEXED' {
        Say '  Outlook IS indexed on this guest. Correct for OutlookAI-Indexed; wrong for'
        Say '  OutlookAI-Unindexed. Run with -Execute.'
    }
    'SETTLING' {
        Say '  NOT AN ANSWER. Either a crawl is still running, or an exclusion landed after one'
        Say '  did and the old rows have not been purged, or nothing has been crawled yet at all.'
        Say '  Wait and re-run -Verify. If it will not converge, the rows predate the exclusion:'
        Say '  run -Execute -RebuildCatalog, which throws the catalog away and re-crawls without'
        Say '  the Outlook scope.'
    }
    'NO-INDEXER' {
        Say '  THIS IS A FAILURE, not a quiet success, and it is the one that looks like one.'
        Say '  There is no working index on this machine at all. A machine with no indexer is'
        Say '  NOT a store with no index frontier: the product takes its "index unreachable"'
        Say '  branch instead, outlook_health omits index.perStore entirely, and'
        Say '  MailService''s non-exhaustive search path does not wrap its index calls in a'
        Say '  catch - so what search does there is unspecified and untested.'
        Say '  Set WSearch back to Automatic, start it, and run this again.'
    }
}

Say ''
Say 'THE REGISTRY IS NOT THE ANSWER, AND NEITHER IS THIS SCRIPT. The verdict above is a'
Say 'statement about the catalog at two moments. The statement the live tier rests on is'
Say 'outlook_health''s index.perStore[], read on BOTH guests and compared - one showing'
Say 'inLocalIndex:true with a frontier, the other inLocalIndex:false with none. Until that'
Say 'comparison has been made, the two guests are assumed different rather than known to be.'

if ($verdict -eq 'UNINDEXED') { exit 0 }
if ($verdict -eq 'INDEXED' -and $Enable) { exit 0 }
exit 1
