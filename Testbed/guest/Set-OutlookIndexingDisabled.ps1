#Requires -Version 5.1
<#
    ============================================================================================
    2026-09-24 (Q69), ON `OutlookAI-Indexed`: WHY NO GUEST WAS EVER INDEXED - AND THE CRAWL SCOPE
    IS NOW WRITTEN THROUGH ITS DOCUMENTED API, IN BOTH DIRECTIONS.
    ============================================================================================

    WHY NO GUEST WAS EVER INDEXED. Every Outlook the testbed started was ELEVATED, and an elevated
    Outlook does not use Windows Search at all. The session-1 route,
    Testbed/guest/Register-InteractiveTask.ps1, registers its task at RunLevel Highest; every
    Outlook started through it - UI watches, Build-Corpus.ps1, COM starts - inherited that token.
    A NON-elevated Outlook registers itself within seconds. Measured on OAI-INDEXED from the same
    clean checkpoint (CP-09: no mapi16 rule, no Outlook row), same profile, a graceful restart in
    between, only the integrity level differing:

                                          ELEVATED (RunLevel Highest)   NOT elevated (Limited)
        mapi16://{SID}/ in the crawl      no rule in 8 min, UI up       root + user INCLUDE rule
                                                                        within 8 s of the start
        catalog                           0 Outlook rows                crawl at once; 19,874
                                                                        notifications in 70 s
        Store.IsInstantSearchEnabled      False                         True
        mssprxy.dll in OUTLOOK.EXE        not loaded                    loaded
        per-store marker, HKCU\...\Outlook\Search   never written       written

    Ruled out by evidence: the Mapi16 protocol handler (registered as on the maintainer's
    workstation; it crawls the moment a rule exists), Outlook's search settings, Group Policy and
    Office policy (none), the event logs (no handler failure), the 25H2 key and "Default indexed
    paths" (Q68). Full record: Docs/live-tier-on-the-vm.md section 8 item 22.

    THE CRAWL SCOPE IS NOW WRITTEN THROUGH ITS DOCUMENTED API, IN BOTH DIRECTIONS
    (SearchCrawlScope.cs, which must be staged beside this script). -Enable -Execute adds the root
    and a user INCLUDE rule for mapi16://{SID}/ - value for value the rule a non-elevated Outlook
    writes, and the maintainer's (WorkingSetRules Include=1 Suppress=0 Default=0 Policy=0
    NoContent=0 Container=0; SearchRoots ProvidesNotifications=1 Container=0). -Execute adds a user
    EXCLUDE rule and the policy. Both ask the service afterwards (IncludedInCrawlScopeEx) and throw
    if it disagrees. Every IID and vtable slot comes from SearchAPI.h 10.0.26100.0, cited per line
    in that file and pinned twice (T1/SearchCrawlScopeInteropTests; -SelfTest).
    MEASURED ON OAI-INDEXED: from a clean CP-09 (NOT-IN-SCOPE, no rule, no row),
    -Enable -Execute wrote policy 0, restarted WSearch, and the service answered included=True
    reason=USER; the registry then held WorkingSetRules\20 and SearchRoots\3 with exactly the values
    above - the same key numbers the self-registering Outlook had been given on the same checkpoint.

    THE CRAWL, AND HOW TO KNOW IT HAS FINISHED. The rule alone crawls nothing: with the scope in, zero
    Outlook rows for 4 minutes with Outlook closed and for 6 with it running ELEVATED
    (IsInstantSearchEnabled still False). Started NOT elevated (Start-OutlookUnelevated.ps1), Outlook
    queued ~19,800 item notifications in its first minute and the 20,000-item corpus was fully
    crawled 7.6 to 9.6 minutes after its start, four runs (9.6, 8.6, 7.6, 8.8 min to the first
    reading with nothing left), ~3,500 items a minute at the peak; the tier profile's two small stores took
    under three more. "Finished" is three signals that fell due in the same minute: the catalog's
    NumberOfItemsToIndex queues all 0, GetCatalogStatus back to IDLE (it runs INCREMENTAL_CRAWL ->
    PROCESSING_NOTIFICATIONS -> IDLE), and the Outlook row count standing still (20,030 for the
    corpus: its 20,000 items plus folders). INDEXED requires all three across two readings, and
    -WaitMinutes waits for them - that is build step 8c.

    THE EXCLUSION, FROM THE INDEXED CHECKPOINT - AND THE ORDER IS THE FINDING. Five ways, each from
    CP-10-INDEXED (20,048 Outlook rows: the corpus 20,028, the identity store 4, the tier store 16),
    Outlook closed; each then watched with Outlook closed and - all but R - with a NON-elevated
    Outlook running, the one that registers itself:

      U  the user EXCLUDE rule alone (AddUserScopeRule + SaveAll). The service reports the scope
         out at once (USER), and after a minute or two THE INDEXER PURGES THE ROWS ITSELF - no
         rebuild; the deletions go through its notification queue - 0 Outlook rows 4.2 min after
         the rule.
      R  the same rule, then a WSearch restart one second later: NOTHING PURGED - 20,048 rows
         6.5 minutes on.
      S  this script as it stood until then (policy 1 + rule + restart, in one go): nothing purged
         either - 20,048 rows after 6 min closed, 5 min of the self-registering Outlook, and a
         reboot. -Execute -RebuildCatalog then took them all out at once (ISearchCatalogManager::
         Reset; 0 rows at the first reading) and the catalog re-crawled its other 431 items in
         about 2 minutes.
      P  PreventIndexingOutlook = 1 alone (+ restart): NOT a crawl-scope rule - the service still
         reports the scope IN (USER) - and it purges nothing: 20,048 rows through 6 min closed,
         5 min of Outlook, and a reboot.
      N  THIS SCRIPT NOW (-Execute: the rule, the wait, then the policy and the restart): the purge
         began 1.3 to 1.9 min after the rule (20,048 -> 16,115 -> 12,879 -> 9,375 -> 4,931 -> 888 ->
         0, half a minute apart), 0 Outlook rows 4.5 min after it, then the policy and the restart,
         and -Verify said UNINDEXED (reason USER). Still 0 after 3 min closed, 5 min of the
         self-registering Outlook, and a reboot. (The notification queue still held 1,623 entries
         when the rows reached 0; it drained after the restart - the queue survives one, the
         pending purge does not.)

    So the purge of a newly excluded scope is work the service holds and LOSES IN A RESTART, and
    does not redo afterwards: R and S still held all 20,048 rows after a reboot. Hence the order in
    -Execute, which -SelfTest pins. Whether the policy on its own would also stop the purge is not
    established; R shows the restart alone does.

    THE EXCLUSION HOLDS AGAINST OUTLOOK. In U, S and N a non-elevated Outlook - which registers an
    INCLUDED scope within seconds (the A/B above) - ran 5 minutes on top of the exclusion and left
    it excluded: no rule change, nothing queued, no row back, and its own IsInstantSearchEnabled
    read False (as it did in P, where only the policy was set). And on a guest that was NEVER
    indexed (CP-09: no rule, no row) the policy ALONE kept a non-elevated Outlook from registering:
    6 minutes, no root or rule added, no row, no per-store marker in HKCU\...\Outlook\Search,
    IsInstantSearchEnabled False, and the same after a reboot. So OutlookAI-Unindexed's recorded
    state (policy only, never indexed) holds against either kind of Outlook as it stands; what the
    policy cannot do is exclude, or purge, a scope that something else has included.
    EnumerateScopeRules stops listing the mapi16 rule once it excludes, while IncludedInCrawlScopeEx
    and WorkingSetRules\<n> Include=0 both show it - which is why -Verify judges by
    IncludedInCrawlScopeEx and prints the registry beside it.

    ============================================================================================
    2026-09-24 (Q68), ON `OutlookAI-Indexed`: THE REGISTRY HALF WAS EXERCISED, AND IT CANNOT
    WRITE. (HISTORY - its "why" is answered by the banner above.)
    ============================================================================================

    1. THE ADMINISTRATORS GROUP CANNOT WRITE THE CRAWL-SCOPE KEYS AT ALL. Their ACL, read on
       OAI-INDEXED (Windows 11 25H2, 26200.8037), identical on SystemIndex, WorkingSetRules, a
       WorkingSetRules\<n> rule, SearchRoots and DefaultRules:
           BUILTIN\Administrators   ReadKey (+CreateLink)          <- no SetValue, no CreateSubKey
           NT AUTHORITY\SYSTEM      SetValue, CreateSubKey, ReadKey
           NT SERVICE\WSearch       FullControl
           NT SERVICE\TrustedInstaller FullControl
       The exact call this script used to make on a rule - New-ItemProperty -Name Include
       -PropertyType DWord -Force - failed from an elevated administrator session with
       "System.Security.SecurityException: Requested registry access is not allowed." So the old
       registry layer could never have worked. THIS STILL HOLDS: the registry is read, never
       written; -SelfTest pins that. What changed on Q69 is that the documented writer - the Crawl
       Scope Manager API, which runs inside WSearch - is now reachable (SearchCrawlScope.cs), and
       both directions go through it.

    2. "THE INDEXED GUEST IS NOT INDEXED, AND NEVER WAS" - no mapi16 rule anywhere, "catalog
       reachable=True anyRow=1 mapiRows=0 mailRows=0", nine days after its corpus was built; not
       after 10 minutes of Outlook UI on the tier profile, 8 on the corpus profile, the
       community-reported 25H2 key, or the "Default indexed paths" policy. Every one of those
       Outlook starts was ELEVATED - see the Q69 banner, which is the explanation.

    3. So the 2026-09-16 UNINDEXED below proves that OAI-UNINDEXED holds no Outlook row, not that
       PreventIndexingOutlook excluded anything. -Verify prints that caveat under an UNINDEXED
       verdict whenever the service reports no rule of any kind for this account's scope.

    ============================================================================================
    RUN 2026-09-16 ON `OutlookAI-Unindexed`, AND IT WORKED. VERDICT: UNINDEXED.  (HISTORY)
    ============================================================================================

    What it did, on OAI-UNINDEXED, 64-bit elevated over PowerShell Direct, with Outlook closed:
    a dry run first; `-Execute -RebuildCatalog` (created the policy key, set
    `PreventIndexingOutlook = 1`, requested a catalog rebuild, restarted WSearch - never disabled
    it); `-Verify`: two readings 10 minutes apart, `catalog reachable=True anyRow=1 mapiRows=0
    mailRows=0`, `VERDICT: UNINDEXED`, indexer running, `wSearchStartMode` `automatic`. It
    reported `no mapi rule for this account under WorkingSetRules - nothing to change`: that guest
    had never had Outlook in its crawl scope either (item 2 above), so the Group Policy layer did
    all the work, against nothing. `Assert-OutlookClosed` was added after that run.

.SYNOPSIS
    Takes a testbed guest's Outlook OUT of the Windows Search index, or with -Enable puts it IN -
    through the Crawl Scope Manager API and the documented Group Policy value - and then proves
    the state by asking the CATALOG, not by reading back what was written.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir. NEEDS
    SearchCrawlScope.cs STAGED BESIDE IT (the interop; see -InteropPath).
    NEVER run this on the maintainer's workstation. It refuses there; see THE GUARD below.

    WHY THIS EXISTS. Docs/live-tier-on-the-vm.md section 1.3. The testbed is two guests:
    OutlookAI-Indexed carries Corpus A and must be INDEXED, OutlookAI-Unindexed carries Corpus B
    and must be NOT INDEXED. Half the live tier wants a populated index; the other half wants a
    store with NO INDEX FRONTIER, which is what drives the seven-day fallback window and the
    sweep and frame measurements. Both guests are built from the same answer file, so until
    something sets the index state on purpose, the two halves measure the same machine.

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
    on these guests: (1) the crawl scope manager - the thing that both SETS and READS BACK the
    scope - lives inside WSearch (CLSID_CSearchManager's AppID names LocalService = WSearch) and
    is unreachable when it is stopped; (2) MailService's non-exhaustive search path does not wrap
    its index calls in a catch, and no test covers a throwing index client; (3)
    T3/Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning
    asserts index.wSearchStartMode == "automatic" and does not filter out on an unindexed guest.

    WHAT IT WRITES, in both directions, and how well attested each part is:

      1. THE CRAWL SCOPE, THROUGH ITS API (DOC). ISearchManager -> ISearchCatalogManager ->
         ISearchCrawlScopeManager: AddUserScopeRule for mapi16://{SID}/ - INCLUDE with -Enable,
         EXCLUDE without - and SaveAll; when including, AddRoot for the same URL first if the
         service has no root for it. That is what the Indexing Options dialog calls, it runs
         inside the service (so the ReadKey-only ACL on the registry keys does not apply), and
         the shape it leaves is the shape a non-elevated Outlook leaves when it registers itself
         (measured, see the banner). The service is asked afterwards, with IncludedInCrawlScopeEx,
         and the script THROWS if it does not now say what was asked.
         EXCLUDING, THE RULE COMES FIRST AND IS LEFT ALONE UNTIL THE INDEXER HAS PURGED THE ROWS
         CRAWLED BEFORE IT (MEASURED - banner, "THE EXCLUSION"): -Execute waits for that, up to
         -PurgeMinutes, before it touches anything else.
      2. THE POLICY (DOC): HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\
         PreventIndexingOutlook - 1 to exclude, 0 with -Enable. The Group Policy "Prevent
         indexing Microsoft Office Outlook" from Search.admx; Microsoft Support prescribes this
         exact registry edit. Machine-wide (HKLM, class="Machine"), and Outlook itself reads it.
         MEASURED 2026-09-24: IT IS NOT A CRAWL-SCOPE RULE. With the policy alone the service still
         reports the scope IN (reason USER, not GROUPPOLICY) and purges nothing; what it does do is
         turn Outlook's own Store.IsInstantSearchEnabled off. So the user rule is what excludes,
         and the policy is a second layer on Outlook's side - which is why it is written alongside
         the rule, never instead of it.
      3. A SERVICE RESTART, after the policy, and never while a purge is pending: a restart one
         second after the exclude rule lost the purge - every row still there after a reboot (banner). The rule itself needs no
         restart - the service answers "excluded" at once and purges on its own. The restart is
         there so the service starts with the policy in place; whether the service acts on the
         policy at all is not established. RESTART, never disable.
      4. With -RebuildCatalog: ISearchCatalogManager::Reset (DOC) - the documented rebuild.

    WHAT -Verify ACTUALLY PROVES, AND WHY IT IS NOT A REGISTRY READ-BACK. Registry values prove
    that something was written, not that the service acts on them. So -Verify asks three things:

      * THE SERVICE'S VIEW OF THE SCOPE - IncludedInCrawlScopeEx for mapi16://{SID}/ (included,
        and why: USER, DEFAULT, GROUPPOLICY or UNKNOWNSCOPE), every mapi rule and root the
        service holds, and the catalog's own counters: its status (IDLE, crawling, ...) and
        NumberOfItemsToIndex, the "items remaining" Indexing Options shows;
      * THE CATALOG, through the same OLE DB provider and statements the shipped server uses
        (OutlookAI.Core/IndexSearch/WsSqlBuilder.cs): a CONTROL probe ("any row at all" - which
        separates "no Outlook rows" from "no catalog"), a MAPI probe (a row under
        mapi16://{SID}/), and a MAIL probe (System.Kind='email', no scope - the safety net under
        the MAPI probe, whose predicate is a URL built at runtime);
      * HOW MANY Outlook rows it holds, PER STORE - every row under the scope, counted and
        grouped by the store segment of its URL.

    All of it TWICE, -SettleMinutes apart, because one reading cannot tell "not indexed" from
    "not indexed yet". Five verdicts, and the three that are not answers are named as loudly as
    the two that are:

      INDEXED        in scope, Outlook rows on both readings, the same count on both, at least
                     -MinimumOutlookRows of them, nothing left in the catalog's queues and the
                     catalog IDLE - the three signals that fell due in the same minute when the
                     20,000-item corpus finished crawling (measured 2026-09-24).
      UNINDEXED      no Outlook row and no mail row on both readings, and something EXCLUDES
                     Outlook - the policy, or a user rule the service reports. With no rule of any
                     kind it is printed with a caveat: only the policy is excluding, and the policy
                     is not a crawl-scope rule (banner).
      SETTLING       the readings disagree, the count is still moving, the catalog still has items
                     queued, fewer rows than -MinimumOutlookRows, rows present under an exclusion
                     (the purge question - see the banner), or in scope with nothing crawled yet.
                     NOT an answer - wait and re-run.
      NOT-IN-SCOPE   no Outlook row, nothing excluding Outlook, and the service reports this
                     account's scope as outside the crawl (no rule). Waiting will not change it.
                     A FAILURE on the indexed guest: run -Enable -Execute.
      NO-INDEXER     any probe, the counters or the scope view failed to reach the service, or
                     the service is not running. The state that LOOKS like success and is not.

    IDEMPOTENT. Run it twice and the second run writes the same rule and policy again and still
    runs the full verification. -Enable is the inverse, so a checkpoint can be re-taken either way.

    WHAT IT NEVER DOES. It never starts Outlook, never creates an Outlook COM object, never
    touches MAPI, never reads or writes an Outlook profile key, and never touches a mail item. It
    never stops or disables the Windows Search service. It never WRITES a crawl-scope registry key
    (an administrator cannot, and the API is the supported writer); it reads them for comparison.
    Its only writes are the one policy value and the crawl scope rule through the API - plus,
    with -RebuildCatalog, a catalog reset through the API.

    THE GUARD, AND IT HAS TWO AXES ON PURPOSE. The account, because the guests autologon as
    vmadmin and no other machine in this project does; and the computer name, because the damage
    this script can do to a real machine - silently emptying its mail search, or indexing mail
    somebody excluded on purpose - is not visible for weeks. Both must pass; the only way past
    either is to name the value you mean. Restated here rather than dot-sourced from
    OutlookMapiInterop.ps1, so this script depends on nothing but its own interop.

    STARTING STATE IT EXPECTS.
      * A guest where Windows is installed and you are logged on as -ExpectedUser. The rule is
        per Windows account (mapi16://{SID}/, the SID of the account running this), so run it
        AS the account whose Outlook is meant - PowerShell Direct as vmadmin is that account.
      * ELEVATED for -Execute (the policy key is HKLM and a service is restarted). -Verify does
        not need it.
      * A 64-BIT PowerShell. The Search.CollatorDSO provider needs an x64 host, so a 32-bit
        session would report NO-INDEXER on a healthy machine.
      * OUTLOOK CLOSED for -Execute. Microsoft documents that the PST provider is "very sensitive
        to the indexing state changing while the PST is open"; close it GRACEFULLY
        (Testbed/host/Restart-Guest.ps1), never taskkill it. -Verify does not need it closed -
        and the crawl needs Outlook RUNNING, NOT ELEVATED (see the banner).

.PARAMETER ExpectedUser
    Accounts this script is allowed to run as. The guests autologon as vmadmin.

.PARAMETER ExpectedComputerNamePrefix
    Computer-name prefix this script is allowed to run on. Testbed/host/New-AnswerFile.ps1
    derives a guest's name by replacing 'OutlookAI-' with 'OAI-'. Do not widen it to an empty
    string.

.PARAMETER Execute
    Write. Without it, prints what it would do and changes nothing.

.PARAMETER Enable
    The inverse: put Outlook IN the index - policy 0, and an INCLUDE user rule (plus a root) for
    mapi16://{SID}/. Needs -Execute to write. The verification that follows will usually say
    SETTLING: including a scope lets the gatherer start, and 20,000 items are not crawled in the
    time this script waits - the crawl also needs Outlook running NOT elevated (banner). Re-run
    -Verify until it says INDEXED. EXIT CODES: -Enable -Execute and -Execute exit 0 when the
    verification that follows says SETTLING or the verdict they were aiming at, because the write
    itself has already been proved by asking the service (a disagreement throws); -Verify on its
    own exits 0 only on INDEXED or UNINDEXED - SETTLING is not an answer.

.PARAMETER Verify
    Probe the service and the catalog and report the actual state. Safe at any time; writes
    nothing; needs no elevation.

.PARAMETER RebuildCatalog
    With -Execute: after the rule, the policy and the restart, ISearchCatalogManager::Reset -
    throw the whole catalog away and re-crawl what the scope now includes. Excluding, it replaces
    the wait for the indexer's own purge (-PurgeMinutes). The way out when rows outlived an
    exclusion that was written in the wrong order (the banner: the pre-Q69 order left all 20,048).

.PARAMETER PurgeMinutes
    Excluding only: how long to wait, after the rule and before the policy and the restart, for
    the indexer to purge the Outlook rows it crawled before the exclusion. No row, no wait. If rows
    remain when it runs out, the script STOPS there - policy not written, service not restarted -
    and exits 1; re-run -Execute later, or -Execute -RebuildCatalog.

.PARAMETER SettleMinutes
    Minutes between the two verification readings. -SettleMinutes 0 takes one reading and gives up
    the only thing that tells "not indexed" from "not indexed yet". Use it to look, never to
    conclude.

.PARAMETER WaitMinutes
    Keep verifying, one reading every -SettleMinutes, for up to this long while the verdict is
    SETTLING - so one call waits out a crawl and says when it FINISHED (the elapsed time is
    printed). 0, the default, takes exactly two readings. The build step (Testbed/README.md
    section 1, 8c) uses it.

.PARAMETER MinimumOutlookRows
    INDEXED also requires at least this many rows under mapi16://{SID}/. The runbook's build step
    passes the corpus size, so "indexed" means "the corpus is in the index", not "one row is".

.PARAMETER InteropPath
    Where SearchCrawlScope.cs is. Defaults to beside this script.

.PARAMETER LogPath
    Transcript file. Everything printed also lands here.

.PARAMETER SelfTest
    Pure: no registry value is read, no catalog opened, no service called. Drives the verdict
    table through synthetic readings; asserts that this file contains no write to the crawl scope
    manager's registry keys; and compiles SearchCrawlScope.cs with THIS PowerShell's compiler and
    checks every interface's IID and every method's vtable slot against the header citations in
    that file (with Marshal.GetComSlotForMethodInfo where the runtime has it - Windows PowerShell
    5.1 does). Exit 0 only if every case passes.

.EXAMPLE
    .\Set-OutlookIndexingDisabled.ps1
    .\Set-OutlookIndexingDisabled.ps1 -Enable -Execute
    .\Set-OutlookIndexingDisabled.ps1 -Verify -MinimumOutlookRows 20000
    .\Set-OutlookIndexingDisabled.ps1 -Verify -SettleMinutes 1 -WaitMinutes 30 -MinimumOutlookRows 20000
    .\Set-OutlookIndexingDisabled.ps1 -Execute
    .\Set-OutlookIndexingDisabled.ps1 -Execute -RebuildCatalog
    .\Set-OutlookIndexingDisabled.ps1 -SelfTest
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
    [int]      $MinimumOutlookRows = 1,
    [int]      $WaitMinutes = 0,
    [int]      $PurgeMinutes = 15,
    [string]   $InteropPath,
    [string]   $LogPath = 'C:\OutlookAI-Q5\set-outlook-indexing.log',
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'

# Resolved here rather than as a parameter default: Windows PowerShell 5.1 run with -File leaves
# $PSScriptRoot empty while it evaluates parameter defaults.
if (-not $InteropPath) { $InteropPath = Join-Path $PSScriptRoot 'SearchCrawlScope.cs' }

# ---------------------------------------------------------------------------------------------
# Everything this script names, in one place, so the blast radius is readable without reading
# the code. EVIDENCE is the honest label on each: DOC means Microsoft documents this value;
# MEASURED means it was read off a real machine; INFERRED means neither.
# ---------------------------------------------------------------------------------------------
$SearchPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search'
$SearchPolicyName = 'PreventIndexingOutlook'   # EVIDENCE: DOC (Search.admx). Absent by default.

$CsmRoot = 'HKLM:\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex'
$WorkingSetRules = Join-Path $CsmRoot 'WorkingSetRules'   # EVIDENCE: MEASURED. Read, never written.
$SearchRoots = Join-Path $CsmRoot 'SearchRoots'           # EVIDENCE: MEASURED. Read, never written.

$WSearchServiceKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\WSearch'

# The follow flags on the rule. Outlook's OWN rule - registered by a non-elevated Outlook on
# OAI-INDEXED, 2026-09-24 - carries Suppress=0 in the registry, as does the maintainer's; the rule
# written here passes 0. (The API cannot report follow flags: Microsoft documents
# ISearchScopeRule::get_FollowFlags as "Not supported", and it returns E_NOTIMPL here.)
$MapiRuleFollowFlags = 0

# The connection string the shipped server uses, character for character
# (OutlookAI.Core/IndexSearch/IndexClients.cs, OleDbIndexClient.ConnectionString).
$CollatorConnectionString = "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

# The statements, matching OutlookAI.Core/IndexSearch/WsSqlBuilder.cs.
$ProbeControlSql = 'SELECT TOP 1 System.ItemUrl FROM SystemIndex'
$ProbeMailSql = "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE System.Kind='email'"

$ClusionReasonNames = @{ 0 = 'UNKNOWNSCOPE'; 1 = 'DEFAULT'; 2 = 'USER'; 3 = 'GROUPPOLICY' }   # SearchAPI.h:2697-2702
$CatalogStatusNames = @{ 0 = 'IDLE'; 1 = 'PAUSED'; 2 = 'RECOVERING'; 3 = 'FULL_CRAWL'; 4 = 'INCREMENTAL_CRAWL'; 5 = 'PROCESSING_NOTIFICATIONS'; 6 = 'SHUTTING_DOWN' }   # SearchAPI.h:3720-3729

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

This script changes whether Outlook is in the Windows Search index. On a machine with real mail
on it that silently empties mail search - for every account, in Outlook's own search box as well
as through this project's tools - or indexes mail somebody excluded on purpose, and nothing
announces either. It is not a mistake anyone notices on the day they make it.

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

-Execute writes an HKLM policy value and restarts a service. An unelevated run would fail
somewhere in the middle, having written some things and not others, which is the one outcome
worse than not running at all.
"@
    }
}

function Assert-OutlookClosed {
    <#
        EVIDENCE: MS-DOC, verbatim, from the MAPI dev article on wrapped PSTs and indexing:
        "The PST provider is very sensitive to the indexing state changing while the PST is
        open. If the state changes, the PST may end up kicking off an installer to repair
        Outlook." https://learn.microsoft.com/en-us/archive/blogs/stephen_griffin/wrapped-pst-and-indexing
        This script changes exactly that state and restarts WSearch, so it refuses while Outlook
        runs - and says WHICH process, leaving the closing to the caller: never taskkill it.
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
WSearch.

Close Outlook first - GRACEFULLY. Testbed/host/Restart-Guest.ps1 quits it the way mailbox-safety
rule 7 describes and restarts the guest without forcing anything. Never taskkill it.

This refusal does not apply to -Verify, which only reads.
"@
}

function Assert-Bitness {
    if (-not [Environment]::Is64BitProcess) {
        throw @"
REFUSING TO RUN: this is a 32-bit PowerShell.

The Search.CollatorDSO provider needs an x64 host (Docs/live-tier-on-the-vm.md section 2.3), so
every catalog probe below would fail and this script would report NO-INDEXER on a machine whose
index is perfectly healthy. Start Windows PowerShell from C:\Windows\System32\WindowsPowerShell\v1.0.
"@
    }
}

function Get-CurrentUserSid {
    return ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
}

# The scope URL: one per Windows account, SID in braces, trailing slash - MEASURED on the
# maintainer's workstation and written in exactly this form by Outlook itself on OAI-INDEXED.
function Get-MapiScopeUrl {
    return 'mapi16://{' + (Get-CurrentUserSid) + '}/'
}

# ---------------------------------------------------------------------------------------------
# The crawl scope, through its documented API (Testbed/guest/SearchCrawlScope.cs).
# ---------------------------------------------------------------------------------------------
function Import-CrawlScopeInterop {
    if ('OutlookAI.Testbed.SearchScope.CrawlScope' -as [type]) { return }
    if (-not (Test-Path -LiteralPath $InteropPath)) {
        throw @"
The crawl-scope interop is not here: $InteropPath

Stage Testbed/guest/SearchCrawlScope.cs BESIDE this script - every mode of this script reads or
writes the crawl scope through it. Copy both files:

    Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path Testbed/guest/Set-OutlookIndexingDisabled.ps1,Testbed/guest/SearchCrawlScope.cs -Destination C:\OutlookAI-Q5\
"@
    }
    # Windows PowerShell 5.1's compiler rejects '#nullable' even inside an inactive #if (measured
    # 2026-09-24), and the test project needs the directive; so it is stripped here and only here.
    $source = [IO.File]::ReadAllText($InteropPath) -replace '(?m)^#nullable.*$', ''
    Add-Type -TypeDefinition $source
}

# What the SERVICE says about this account's Outlook scope - policy, user and default rules all
# applied - plus every mapi rule and root it holds. Never throws: an unreachable service is a
# reading (Reached = $false), and the verdict turns it into NO-INDEXER.
function Get-ScopeView {
    $url = Get-MapiScopeUrl
    try {
        Import-CrawlScopeInterop
        $inc = [OutlookAI.Testbed.SearchScope.CrawlScope]::IncludedInCrawlScope($url)
        $rules = @([OutlookAI.Testbed.SearchScope.CrawlScope]::ReadRules() | Where-Object { $_.Url -like 'mapi*' })
        $roots = @([OutlookAI.Testbed.SearchScope.CrawlScope]::ReadRoots() | Where-Object { $_.Url -like 'mapi*' })
        $reasonName = 'reason ' + $inc[1]
        if ($ClusionReasonNames.ContainsKey([int]$inc[1])) { $reasonName = $ClusionReasonNames[[int]$inc[1]] }
        return [pscustomobject]@{ Reached = $true; Url = $url; Included = ($inc[0] -ne 0); Reason = [int]$inc[1]; ReasonName = $reasonName; Rules = $rules; Roots = $roots; Error = $null }
    }
    catch {
        return [pscustomobject]@{ Reached = $false; Url = $url; Included = $false; Reason = -1; ReasonName = 'unreadable'; Rules = @(); Roots = @(); Error = $_.Exception.Message }
    }
}

function Get-CatalogCounters {
    # Three attempts, 5 s apart: measured 2026-09-24, NumberOfItems failed with 0xC004365A for a
    # read taken seconds after a guest restart, while WSearch was still starting. A service that is
    # still not answering after that is reported as unreachable - NO-INDEXER, never a guess.
    $lastError = 'not attempted'
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Import-CrawlScopeInterop
            $c = [OutlookAI.Testbed.SearchScope.CrawlScope]::ReadCounters()
            $statusName = 'status ' + $c.Status
            if ($CatalogStatusNames.ContainsKey([int]$c.Status)) { $statusName = $CatalogStatusNames[[int]$c.Status] }
            return [pscustomobject]@{
                Reached = $true; Status = $c.Status; StatusName = $statusName; PausedReason = $c.PausedReason
                NumberOfItems = $c.NumberOfItems; Queued = ($c.IncrementalCount + $c.NotificationQueue + $c.HighPriorityQueue)
                Incremental = $c.IncrementalCount; Notifications = $c.NotificationQueue; HighPriority = $c.HighPriorityQueue
                IndexerVersion = $c.IndexerVersion; Error = $null
            }
        }
        catch {
            $lastError = $_.Exception.Message
            if ($attempt -lt 3) { Start-Sleep -Seconds 5 }
        }
    }
    return [pscustomobject]@{ Reached = $false; Status = -1; StatusName = 'unreadable'; PausedReason = -1; NumberOfItems = -1; Queued = -1; Incremental = -1; Notifications = -1; HighPriority = -1; IndexerVersion = ''; Error = $lastError }
}
# THE WRITE. One user rule for mapi16://{SID}/ through AddUserScopeRule + SaveAll, and - when
# including - the search root it needs (AddRoot with ProvidesNotifications, if the service has no
# root for the URL yet). Read back from the SERVICE afterwards; throws if the service does not now
# say what was asked. "The call returned S_OK" is not the claim this script makes.
function Set-MapiScopeRule {
    param([Parameter(Mandatory = $true)] [bool] $Include)
    Import-CrawlScopeInterop
    $url = Get-MapiScopeUrl
    $verb = 'EXCLUDE'
    if ($Include) { $verb = 'INCLUDE' }
    $before = Get-ScopeView
    Say ("  crawl scope (service) before: {0} included={1} reason={2}" -f $url, $before.Included, $before.ReasonName)
    [OutlookAI.Testbed.SearchScope.CrawlScope]::SetUserRule($url, $Include, $Include, [uint32]$MapiRuleFollowFlags)
    $rootNote = ''
    if ($Include) { $rootNote = ' + AddRoot (if missing)' }
    Say ("  AddUserScopeRule({0}, {1}, overrideChildren=TRUE, followFlags={2}){3} + SaveAll: done" -f $url, $verb, $MapiRuleFollowFlags, $rootNote)
}

function Assert-ScopeIs {
    param([Parameter(Mandatory = $true)] [bool] $Included)
    $view = Get-ScopeView
    Say ("  crawl scope (service) now:    {0} included={1} reason={2}" -f $view.Url, $view.Included, $view.ReasonName)
    if (-not $view.Reached) { throw "The crawl scope could not be read back from the service: $($view.Error)" }
    if ($view.Included -ne $Included) {
        throw ("The service does not report the state that was written: IncludedInCrawlScopeEx says included={0} reason={1}, wanted included={2}. " -f $view.Included, $view.ReasonName, $Included) +
            'The documented precedence is Group Policy > user rules > default rules: if the reason is GROUPPOLICY, the policy value is what decides.'
    }
}

# ---------------------------------------------------------------------------------------------
# Reading the crawl scope from the REGISTRY - for comparison with the service's own view only.
# Never written: an administrator cannot (banner), and the API above is the supported writer.
# ---------------------------------------------------------------------------------------------
function Get-MapiRegistryRules {
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
                Container  = Split-Path -Leaf $container
                KeyName    = $child.PSChildName
                Url        = $url
                IsThisUser = ($url -eq $wanted)
                Include    = $values.Include
                Suppress   = $values.Suppress
                Default    = $values.Default
                Policy     = $values.Policy
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
# The catalog probes. Late-bound ADODB over Search.CollatorDSO - the fallback path the shipped
# server carries (OutlookAI.Core/IndexSearch/IndexClients.cs, AdodbIndexClient). Returns
# Reached = $false when the catalog could not be reached AT ALL, which is a different answer from
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

# EVERY row under this account's Outlook scope, counted and grouped by store - the store is the
# first path segment after mapi16://{SID}/, "<StoreDisplayName>($hash)". A few seconds for tens of
# thousands of rows. Reached = $false when the catalog could not be read; never a silent zero.
function Measure-MapiRows {
    $scope = Get-MapiScopeUrl
    $connection = $null
    $recordset = $null
    $perStore = @{}
    $total = 0
    try {
        $connection = New-Object -ComObject 'ADODB.Connection'
        $connection.Open($CollatorConnectionString)
        $recordset = New-Object -ComObject 'ADODB.Recordset'
        $recordset.Open("SELECT System.ItemUrl FROM SystemIndex WHERE SCOPE='" + $scope + "'", $connection)
        while (-not $recordset.EOF) {
            $u = [string]$recordset.Fields.Item(0).Value
            $store = '<scope root>'
            if ($u.Length -gt $scope.Length) { $store = ($u.Substring($scope.Length) -split '/')[0] }
            if (-not $perStore.ContainsKey($store)) { $perStore[$store] = 0 }
            $perStore[$store]++
            $total++
            $recordset.MoveNext()
        }
        return [pscustomobject]@{ Reached = $true; Total = $total; PerStore = $perStore; Error = $null }
    }
    catch {
        return [pscustomobject]@{ Reached = $false; Total = 0; PerStore = @{}; Error = $_.Exception.Message }
    }
    finally {
        foreach ($o in @($recordset, $connection)) {
            if ($null -eq $o) { continue }
            try { if ($o.State -ne 0) { $o.Close() } } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
        }
    }
}

# The MAPI probe carries this account's SID. SCOPE is the recursive store-scope predicate the
# product uses; the quoting matches WsSqlBuilder's.
function Get-MapiProbeSql {
    return "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + (Get-MapiScopeUrl) + "'"
}

function Read-IndexState {
    return [pscustomobject]@{
        TakenAt  = Get-Date
        Control  = Invoke-IndexProbe -Sql $ProbeControlSql
        Mapi     = Invoke-IndexProbe -Sql (Get-MapiProbeSql)
        Mail     = Invoke-IndexProbe -Sql $ProbeMailSql
        Count    = Measure-MapiRows
        Counters = Get-CatalogCounters
    }
}

function Show-IndexState {
    param($State, [string] $Label)
    Say ("  {0}: catalog reachable={1} anyRow={2} mapiRows={3} mailRows={4} outlookRowsTotal={5}" -f
        $Label, $State.Control.Reached, $State.Control.RowCount, $State.Mapi.RowCount, $State.Mail.RowCount, $State.Count.Total)
    foreach ($store in ($State.Count.PerStore.Keys | Sort-Object)) {
        Say ("    store {0}: {1} row(s)" -f $store, $State.Count.PerStore[$store])
    }
    $c = $State.Counters
    Say ("    catalog: status={0} pausedReason={1} items={2} queued={3} (incremental={4} notifications={5} highPriority={6})" -f
        $c.StatusName, $c.PausedReason, $c.NumberOfItems, $c.Queued, $c.Incremental, $c.Notifications, $c.HighPriority)
    foreach ($pair in @(@('control', $State.Control), @('mapi', $State.Mapi), @('mail', $State.Mail), @('count', $State.Count), @('counters', $State.Counters))) {
        if (-not $pair[1].Reached) { Say ("    {0} probe FAILED: {1}" -f $pair[0], $pair[1].Error) }
    }
}

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

function Restart-SearchService {
    $state = Get-WSearchState
    if ($state.StartMode -eq 'disabled') {
        throw @"
REFUSING TO CONTINUE: the Windows Search service is DISABLED on this guest.

That is not the state this testbed wants and this script will not work around it. A machine with
no indexer is not a store with no index frontier: outlook_health reports
index.provider = "unavailable: ..." and omits index.perStore entirely. Set WSearch back to
Automatic and start it, then run this again.
"@
    }

    Say '  restarting WSearch so it starts with the policy in place (restart, never disable)'
    Restart-Service -Name 'WSearch' -Force
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline -and (Get-Service -Name 'WSearch').Status -ne 'Running') { Start-Sleep -Seconds 2 }
    Say ('  WSearch is now {0}' -f (Get-Service -Name 'WSearch').Status)
}

# The documented catalog rebuild: ISearchCatalogManager::Reset throws the catalog away and
# re-crawls everything the scope now includes - what Indexing Options' Rebuild button does. It
# replaced the SetupCompletedSuccessfully = 0 recipe, which Q68 measured working but which is a
# registry side door; the API is the front one.
function Invoke-CatalogReset {
    Import-CrawlScopeInterop
    Say '  ISearchCatalogManager::Reset - the documented rebuild: the catalog is discarded and re-crawled'
    [OutlookAI.Testbed.SearchScope.CrawlScope]::ResetCatalog()
    Say '  reset requested. -Verify is what says it has finished, not a clock.'
}

# THE PURGE, waited for rather than assumed: every Outlook row under this account's scope, counted
# every 30 s until there are none or -Minutes have passed. The exclusion path calls it after the rule
# and BEFORE the policy and the restart (banner, "THE EXCLUSION"). A catalog that cannot be read
# throws - "could not count" is never "zero".
function Wait-OutlookRowsPurged {
    param([Parameter(Mandatory = $true)] [int] $Minutes)
    $started = Get-Date
    $deadline = $started.AddMinutes($Minutes)
    while ($true) {
        $count = Measure-MapiRows
        if (-not $count.Reached) { throw "The catalog could not be read while waiting for the purge: $($count.Error)" }
        $counters = Get-CatalogCounters
        $elapsed = ((Get-Date) - $started).TotalMinutes
        Say ('    {0,4:N1} min: {1} Outlook row(s); catalog {2}, {3} queued' -f $elapsed, $count.Total, $counters.StatusName, $counters.Queued)
        if ($count.Total -eq 0) { return [pscustomobject]@{ Purged = $true; Remaining = 0; Minutes = $elapsed } }
        if ((Get-Date) -ge $deadline) { return [pscustomobject]@{ Purged = $false; Remaining = $count.Total; Minutes = $elapsed } }
        Start-Sleep -Seconds 30
    }
}

# ---------------------------------------------------------------------------------------------
# The verdict. Five states; the three that are not answers are named as loudly as the two that
# are - a guest believed unindexed while it quietly crawls, or believed indexed while its crawl is
# half done, produces measurements that look fine and mean nothing.
# ---------------------------------------------------------------------------------------------
function Get-Verdict {
    param($First, $Second, $Scope, $PolicyValue, $Service, [int] $MinimumRows = 1)

    # EVERY probe must have REACHED the service, not just the control one. A probe that threw
    # reports zero rows, and zero rows is the evidence UNINDEXED rests on.
    foreach ($reading in @($First, $Second)) {
        foreach ($probe in @($reading.Control, $reading.Mapi, $reading.Mail, $reading.Count, $reading.Counters)) {
            if (-not $probe.Reached) { return 'NO-INDEXER' }
        }
    }
    if (-not $Scope.Reached) { return 'NO-INDEXER' }
    if ($Service.StartMode -eq 'disabled' -or -not $Service.IndexerRunning) { return 'NO-INDEXER' }

    if ($First.Control.RowCount -eq 0 -and $Second.Control.RowCount -eq 0) {
        # The catalog answered and holds nothing at all - it is rebuilding, or empty.
        return 'SETTLING'
    }

    $mapiFirst = ($First.Mapi.RowCount -gt 0)
    $mapiSecond = ($Second.Mapi.RowCount -gt 0)
    $mailFirst = ($First.Mail.RowCount -gt 0)
    $mailSecond = ($Second.Mail.RowCount -gt 0)
    if ($mapiFirst -ne $mapiSecond -or $mailFirst -ne $mailSecond) { return 'SETTLING' }

    # Excluded: the policy says so, or the SERVICE reports the scope out because of a user or
    # policy rule. In scope: the service reports it included. Neither: no rule reaches it.
    $excluded = ($PolicyValue -eq 1) -or ((-not $Scope.Included) -and ($Scope.Reason -eq 2 -or $Scope.Reason -eq 3))
    $inScope = [bool]$Scope.Included

    if (-not $mapiFirst -and -not $mailFirst) {
        if ($excluded) { return 'UNINDEXED' }
        if (-not $inScope) { return 'NOT-IN-SCOPE' }
        return 'SETTLING'
    }

    # Rows present under an exclusion: the exclusion landed after a crawl and the rows are still
    # there. Rows present with no rule at all: the same, from a removed rule.
    if ($excluded -or -not $inScope) { return 'SETTLING' }

    # In scope with rows: indexed only when the crawl has FINISHED - the same count on both
    # readings, at least the minimum, nothing left in the catalog's queues, and the catalog IDLE.
    # MEASURED on OAI-INDEXED 2026-09-24: through a 20,000-item crawl the status went
    # INCREMENTAL_CRAWL -> PROCESSING_NOTIFICATIONS -> IDLE, and it reached IDLE in the same minute
    # the notification queue reached 0 and the row count stopped moving.
    if ($First.Count.Total -ne $Second.Count.Total) { return 'SETTLING' }
    if ($Second.Count.Total -lt $MinimumRows) { return 'SETTLING' }
    if ($Second.Counters.Queued -gt 0) { return 'SETTLING' }
    if ($Second.Counters.Status -ne 0) { return 'SETTLING' }
    return 'INDEXED'
}

# ---------------------------------------------------------------------------------------------
# -SelfTest. Pure: synthetic readings through Get-Verdict, a source check that this file writes
# no crawl-scope registry key, and a compile of the interop with this PowerShell's own compiler
# whose IIDs and vtable slots are checked against the header citations in that file.
# ---------------------------------------------------------------------------------------------
function Invoke-SelfTest {
    $script:stFail = 0
    $script:stPass = 0
    function Pass([string] $m) { $script:stPass++; Write-Host "  PASS  $m" }
    function Fail([string] $m) { $script:stFail++; Write-Host "  FAIL  $m" }
    function New-Probe([bool] $Reached, [int] $Rows) { [pscustomobject]@{ Reached = $Reached; RowCount = $Rows; Error = $null } }
    function New-Reading([int] $Control = 1, [int] $Mapi = 0, [int] $Mail = 0, [bool] $Reached = $true, [int] $Total = -1, [int] $Queued = 0, [int] $Status = 0) {
        if ($Total -lt 0) { $Total = $Mapi }
        [pscustomobject]@{
            TakenAt  = Get-Date
            Control  = (New-Probe $Reached $Control)
            Mapi     = (New-Probe $true $Mapi)
            Mail     = (New-Probe $true $Mail)
            Count    = [pscustomobject]@{ Reached = $true; Total = $Total; PerStore = @{}; Error = $null }
            Counters = [pscustomobject]@{ Reached = $true; Status = $Status; StatusName = 'synthetic'; Queued = $Queued }
        }
    }
    function New-Scope([bool] $Included, [int] $Reason, [bool] $Reached = $true) { [pscustomobject]@{ Reached = $Reached; Included = $Included; Reason = $Reason } }
    $ok = [pscustomobject]@{ StartMode = 'automatic'; Status = 'Running'; IndexerRunning = $true }
    $off = [pscustomobject]@{ StartMode = 'disabled'; Status = 'Stopped'; IndexerRunning = $false }
    $noRule = New-Scope $false 0
    $inUser = New-Scope $true 2
    $outUser = New-Scope $false 2
    $outPolicy = New-Scope $false 3

    $cases = @(
        @('no rule, no policy, no rows (OAI-INDEXED as measured 2026-09-24)', (New-Reading), (New-Reading), $noRule, $null, $ok, 1, 'NOT-IN-SCOPE'),
        @('no rule, policy=1, no rows (OAI-UNINDEXED as measured 2026-09-16)', (New-Reading), (New-Reading), $noRule, 1, $ok, 1, 'UNINDEXED'),
        @('user exclude rule, no policy, no rows', (New-Reading), (New-Reading), $outUser, $null, $ok, 1, 'UNINDEXED'),
        @('policy exclusion reported by the service, no rows', (New-Reading), (New-Reading), $outPolicy, 1, $ok, 1, 'UNINDEXED'),
        @('included, no rows yet', (New-Reading), (New-Reading), $inUser, $null, $ok, 1, 'SETTLING'),
        @('included, rows, same count, queues empty', (New-Reading 1 1 1 $true 20150), (New-Reading 1 1 1 $true 20150), $inUser, 0, $ok, 20000, 'INDEXED'),
        @('included, rows, count still moving', (New-Reading 1 1 1 $true 14186), (New-Reading 1 1 1 $true 20150), $inUser, $null, $ok, 1, 'SETTLING'),
        @('included, rows, same count, items still queued', (New-Reading 1 1 1 $true 20150), (New-Reading 1 1 1 $true 20150 12), $inUser, $null, $ok, 1, 'SETTLING'),
        @('included, rows, same count, nothing queued, catalog not IDLE yet', (New-Reading 1 1 1 $true 20030), (New-Reading 1 1 1 $true 20030 0 5), $inUser, $null, $ok, 1, 'SETTLING'),
        @('as measured 2026-09-24 when the corpus crawl finished (20,030 rows, IDLE)', (New-Reading 1 1 1 $true 20030 0 0), (New-Reading 1 1 1 $true 20030 0 0), $inUser, 0, $ok, 20000, 'INDEXED'),
        @('included, rows, stable but below the minimum', (New-Reading 1 1 1 $true 145), (New-Reading 1 1 1 $true 145), $inUser, $null, $ok, 20000, 'SETTLING'),
        @('rows present under a policy exclusion (not purged yet)', (New-Reading 1 1 1), (New-Reading 1 1 1), $outPolicy, 1, $ok, 1, 'SETTLING'),
        @('rows present under a user exclusion (not purged yet)', (New-Reading 1 1 1), (New-Reading 1 1 1), $outUser, $null, $ok, 1, 'SETTLING'),
        @('rows present, no rule at all (a removed rule)', (New-Reading 1 1 1), (New-Reading 1 1 1), $noRule, $null, $ok, 1, 'SETTLING'),
        @('readings disagree (crawl moving)', (New-Reading 1 0 0), (New-Reading 1 1 1), $inUser, $null, $ok, 1, 'SETTLING'),
        @('a probe did not reach the catalog', (New-Reading 1 0 0 $false), (New-Reading), $noRule, 1, $ok, 1, 'NO-INDEXER'),
        @('the scope view did not reach the service', (New-Reading), (New-Reading), (New-Scope $false 0 $false), 1, $ok, 1, 'NO-INDEXER'),
        @('service disabled', (New-Reading), (New-Reading), $noRule, 1, $off, 1, 'NO-INDEXER'),
        @('empty catalog (rebuilding)', (New-Reading 0), (New-Reading 0), $noRule, 1, $ok, 1, 'SETTLING')
    )
    foreach ($c in $cases) {
        $got = Get-Verdict -First $c[1] -Second $c[2] -Scope $c[3] -PolicyValue $c[4] -Service $c[5] -MinimumRows $c[6]
        if ($got -eq $c[7]) { Pass ("{0,-13} {1}" -f $got, $c[0]) }
        else { Fail ("expected {0}, got {1}: {2}" -f $c[7], $got, $c[0]) }
    }

    # The invariant the Q68 fix exists for: no write cmdlet in this file touches the crawl scope
    # manager's registry keys. Comment lines are skipped; everything else is scanned as text.
    $writers = '(New-ItemProperty|Set-ItemProperty|New-Item\b|Remove-ItemProperty|Remove-Item\b|Rename-ItemProperty)'
    $targets = '(WorkingSetRules|SearchRoots|DefaultRules|CrawlScopeManager|CsmRoot|\$rule\.KeyPath)'
    $offending = @()
    $n = 0
    foreach ($line in [IO.File]::ReadAllLines($PSCommandPath)) {
        $n++
        $t = $line.Trim()
        if ($t.StartsWith('#') -or $t.StartsWith('Say ') -or $t.StartsWith("'") -or $t.StartsWith('$targets') -or $t.StartsWith('$writers')) { continue }
        if ($t -match $writers -and $t -match $targets) { $offending += "line ${n}: $t" }
    }
    if ($offending.Count -eq 0) { Pass 'source        no write cmdlet in this file targets the crawl scope manager''s registry keys' }
    else { Fail 'source        a registry write to the crawl scope keys is back:'; foreach ($o in $offending) { Write-Host "          $o" } }

    # Every command this file calls must exist - as a function defined here, or as a command this
    # PowerShell knows. Added 2026-09-24 after -RebuildCatalog was found calling a function that no
    # longer existed: nothing short of running that exact switch would have noticed.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $defined = @{}
    foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fd.Name] = $true }
    $unknown = @()
    foreach ($cmd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $name = $cmd.GetCommandName()
        if (-not $name -or $defined.ContainsKey($name)) { continue }
        if (Get-Command -Name $name -ErrorAction SilentlyContinue) { continue }
        $unknown += "$name (line $($cmd.Extent.StartLineNumber))"
    }
    if ($unknown.Count -eq 0) { Pass ('source        every command this file calls exists ({0} functions defined here)' -f $defined.Count) }
    else { Fail ('source        calls to commands that do not exist: ' + (($unknown | Select-Object -Unique) -join ', ')) }

    # The exclusion's ORDER is a measured finding, not a style (banner, "THE EXCLUSION"): the rule,
    # then the wait for the indexer's own purge, and only then the policy and the restart. Pinned on
    # the statement block whose Say prints the exclusion's heading, in source order.
    $heading = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Say' -and $n.CommandElements.Count -eq 2 -and $n.CommandElements[1].Extent.Text -eq "'== Execute: taking Outlook OUT of the index =='" }, $true)
    $block = $heading
    while ($block -and -not ($block -is [System.Management.Automation.Language.StatementBlockAst])) { $block = $block.Parent }
    $wantedOrder = @('Set-MapiScopeRule', 'Wait-OutlookRowsPurged', 'Write-PolicyValue', 'Restart-SearchService')
    $seen = @()
    if ($block) {
        foreach ($cmd in ($block.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | Sort-Object { $_.Extent.StartOffset })) {
            $name = $cmd.GetCommandName()
            if ($wantedOrder -contains $name -and $seen -notcontains $name) { $seen += $name }
        }
    }
    if (($seen -join ' > ') -eq ($wantedOrder -join ' > ')) { Pass ('source        the exclusion runs {0}' -f ($seen -join ' > ')) }
    else { Fail ('source        the exclusion must run {0}; it runs {1}' -f ($wantedOrder -join ' > '), ($seen -join ' > ')) }

    # The verdict is judged on the service and the scope AS OF THE READING, not as of the start: the
    # reading loop must re-read both before it calls Get-Verdict (the stale-Stopped NO-INDEXER above).
    $loop = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.WhileStatementAst] -and $n.Body.Find({ param($m) $m -is [System.Management.Automation.Language.CommandAst] -and $m.GetCommandName() -eq 'Get-Verdict' }, $true) }, $true)
    $loopOrder = @()
    if ($loop) {
        foreach ($cmd in ($loop.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | Sort-Object { $_.Extent.StartOffset })) {
            $name = $cmd.GetCommandName()
            if (@('Read-IndexState', 'Get-ScopeView', 'Get-WSearchState', 'Get-Verdict') -contains $name) { $loopOrder += $name }
        }
    }
    if (($loopOrder -join ' > ') -eq 'Read-IndexState > Get-ScopeView > Get-WSearchState > Get-Verdict') { Pass 'source        each reading re-reads the scope and the service before the verdict' }
    else { Fail ('source        the reading loop must run Read-IndexState > Get-ScopeView > Get-WSearchState > Get-Verdict; it runs {0}' -f ($loopOrder -join ' > ')) }

    # The interop, compiled by THIS PowerShell's compiler. Every interface's IID must be the one its
    # declaration cites; every method's vtable slot must be the one its line cites (":<header line>
    # slot <n>"). T1/SearchCrawlScopeInteropTests pins those citations to SearchAPI.h itself.
    if (-not (Test-Path -LiteralPath $InteropPath)) {
        Fail "interop       $InteropPath is not here - stage SearchCrawlScope.cs beside this script"
    }
    else {
        try {
            Import-CrawlScopeInterop
            $lines = [IO.File]::ReadAllLines($InteropPath)
            $asm = [OutlookAI.Testbed.SearchScope.CrawlScope].Assembly
            $interfaces = @($asm.GetTypes() | Where-Object { $_.IsInterface -and $_.Namespace -eq 'OutlookAI.Testbed.SearchScope' })
            $perMethodSlot = [bool]([System.Runtime.InteropServices.Marshal].GetMethod('GetComSlotForMethodInfo'))
            $checked = 0
            $bad = @()
            # The GUIDs, pinned HERE as well as in T1/SearchCrawlScopeInteropTests - independently of
            # the interop file, so an edited IID there fails this check too. Windows SDK 10.0.26100.0.
            $pinned = @{
                'ISearchManager'           = 'AB310581-AC80-11D1-8DF3-00C04FB6EF69'   # SearchAPI.h:5428
                'ISearchCatalogManager'    = 'AB310581-AC80-11D1-8DF3-00C04FB6EF50'   # SearchAPI.h:3763
                'ISearchCrawlScopeManager' = 'AB310581-AC80-11D1-8DF3-00C04FB6EF55'   # SearchAPI.h:2720
                'ISearchRoot'              = '04C18CCF-1F57-4CBD-88CC-3900F5195CE3'   # SearchAPI.h:2018
                'IEnumSearchRoots'         = 'AB310581-AC80-11D1-8DF3-00C04FB6EF52'   # SearchAPI.h:2333
                'ISearchScopeRule'         = 'AB310581-AC80-11D1-8DF3-00C04FB6EF53'   # SearchAPI.h:2467
                'IEnumSearchScopeRules'    = 'AB310581-AC80-11D1-8DF3-00C04FB6EF54'   # SearchAPI.h:2584
            }
            if ([guid][OutlookAI.Testbed.SearchScope.SearchApiGuids]::CLSID_CSearchManager -ne [guid]'7D096C5F-AC08-4F1F-BEB7-5C22C517CE39') { $bad += 'CLSID_CSearchManager is not SearchAPI.h:6057' }
            if ([guid][OutlookAI.Testbed.SearchScope.SearchApiGuids]::CLSID_CSearchRoot -ne [guid]'30766BD2-EA1C-4F28-BF27-0B44E2F68DB7') { $bad += 'CLSID_CSearchRoot is not SearchAPI.h:6065' }
            foreach ($t in $interfaces) {
                if (-not $pinned.ContainsKey($t.Name)) { $bad += "$($t.Name): an interface this check does not pin"; continue }
                if ($t.GUID -ne [guid]$pinned[$t.Name]) { $bad += "$($t.Name): IID $($t.GUID) is not $($pinned[$t.Name]) (SearchAPI.h)" }
                $start = ($lines | Select-String ('public interface ' + $t.Name + '\s*$')).LineNumber
                $cited = @{}
                for ($i = $start; $i -lt $lines.Count; $i++) {
                    if ($lines[$i] -match '^\s*\}') { break }
                    if ($lines[$i] -match '\[PreserveSig\]\s+int\s+(\w+)\(') {
                        $name = $Matches[1]
                        if ((($lines[$i], $lines[$i + 1]) -join ' ') -match '//\s*:(\d+)\s+slot\s+(\d+)') { $cited[$name] = [int]$Matches[2] }
                    }
                }
                $methods = @($t.GetMethods() | Sort-Object MetadataToken)
                $startSlot = [System.Runtime.InteropServices.Marshal]::GetStartComSlot($t)
                if ($startSlot -ne 3) { $bad += "$($t.Name): the CLR starts it at slot $startSlot, not 3 (IUnknown)" }
                for ($k = 0; $k -lt $methods.Count; $k++) {
                    $m = $methods[$k]
                    $checked++
                    $slot = $startSlot + $k
                    if ($perMethodSlot) { $slot = [System.Runtime.InteropServices.Marshal]::GetComSlotForMethodInfo($m) }
                    if (-not $cited.ContainsKey($m.Name)) { $bad += "$($t.Name).$($m.Name): no ':<line> slot <n>' citation" }
                    elseif ($cited[$m.Name] -ne $slot) { $bad += "$($t.Name).$($m.Name): the CLR puts it in slot $slot, the file cites slot $($cited[$m.Name])" }
                }
            }
            $how = 'metadata order from GetStartComSlot'
            if ($perMethodSlot) { $how = 'Marshal.GetComSlotForMethodInfo' }
            if ($interfaces.Count -ne 7) { $bad += "expected 7 interfaces in the interop, found $($interfaces.Count)" }
            if ($bad.Count -eq 0) { Pass ("interop       compiled; {0} interfaces, {1} methods: every IID and every slot as cited ({2})" -f $interfaces.Count, $checked, $how) }
            else { Fail 'interop       the compiled vtable does not match the citations:'; foreach ($b in $bad) { Write-Host "          $b" } }
        }
        catch { Fail "interop       did not compile or could not be inspected: $($_.Exception.Message)" }
    }

    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed." -f $script:stPass, $script:stFail)
    if ($script:stFail -gt 0) { exit 1 }
    exit 0
}

# =============================================================================================
# Main
# =============================================================================================
if ($SelfTest) { Invoke-SelfTest }

Assert-Bitness

if (-not ($Execute -or $Verify -or $Enable)) {
    Say 'DRY RUN. Nothing is written. This is what -Execute would do:'
    Say ''
    Say '  1. AddUserScopeRule(mapi16://{SID}/, EXCLUDE) + SaveAll            [DOC]  the Crawl Scope Manager API'
    Say "  2. wait up to -PurgeMinutes ($PurgeMinutes) for the indexer to purge old Outlook rows   [MEASURED]  before 3 and 4"
    Say "  3. $SearchPolicyKey\$SearchPolicyName = 1             [DOC]  policy: prevent indexing Outlook"
    Say '  4. restart WSearch (never disable it)'
    Say '  5. with -RebuildCatalog: ISearchCatalogManager::Reset (instead of the wait in 2)   [DOC]'
    Say ''
    Say '  and with -Enable -Execute, the inverse: policy = 0, AddRoot (if missing) + AddUserScopeRule(INCLUDE)'
    Say '  + SaveAll, restart WSearch. Then the same verification -Verify runs on its own.'
    Say ''
    Say 'Re-run with -Execute to exclude, -Enable -Execute to include, or -Verify to probe only.'
    return
}

Assert-TestbedGuestLocal

if ($Execute -or $Enable) {
    if (-not $Execute) { throw '-Enable writes nothing without -Execute. Run -Enable -Execute, or -Verify to look.' }
    Assert-Elevated
    Assert-OutlookClosed

    $service = Get-WSearchState
    Say ('== Windows Search: startMode={0} status={1} indexerRunning={2} ==' -f
        $service.StartMode, $service.Status, $service.IndexerRunning)

    if ($Enable) {
        Say '== Execute (-Enable): putting Outlook IN the index =='
        Write-PolicyValue -Value 0
        # The policy first and a restart, then the rule - the order the measured -Enable runs used.
        # (The policy turned out not to be a crawl-scope rule at all - banner - so this restart is
        # belt and braces; nothing measured depends on it.)
        Restart-SearchService
        Set-MapiScopeRule -Include $true
        Assert-ScopeIs -Included $true
    }
    else {
        Say '== Execute: taking Outlook OUT of the index =='
        # THE ORDER IS THE FINDING (banner, "THE EXCLUSION"): the rule first, then - while nothing
        # else has touched the service - the indexer's own purge of the rows crawled before it, and
        # only then the policy and the restart. Policy + rule + restart in one go, the order this
        # script used until Q69, left every one of 20,048 rows in place through a reboot.
        Set-MapiScopeRule -Include $false
        Assert-ScopeIs -Included $false
        $before = Measure-MapiRows
        if (-not $before.Reached) { throw "The catalog could not be read to see whether Outlook rows need purging: $($before.Error)" }
        if ($before.Total -eq 0) {
            Say '  no Outlook row in the catalog: nothing to purge'
        }
        elseif ($RebuildCatalog) {
            Say ('  {0} Outlook row(s) in the catalog; -RebuildCatalog resets it below instead of waiting for the purge' -f $before.Total)
        }
        else {
            Say ('  {0} Outlook row(s) crawled before the exclusion. Waiting up to {1} min for the indexer to purge them' -f $before.Total, $PurgeMinutes)
            Say '  itself - BEFORE the policy and the service restart, which is the order that lets it.'
            $purge = Wait-OutlookRowsPurged -Minutes $PurgeMinutes
            if (-not $purge.Purged) {
                Say ''
                Say ('  STOPPING: {0} Outlook row(s) are still in the catalog after {1:N1} min. Neither the policy nor the' -f $purge.Remaining, $purge.Minutes)
                Say '  restart has been done, so nothing here stops the purge from finishing: re-run -Execute later'
                Say '  (it is idempotent), or -Execute -RebuildCatalog to discard the catalog instead.'
                exit 1
            }
            Say ('  purged by the indexer itself: 0 Outlook rows {0:N1} min after the rule' -f $purge.Minutes)
        }
        Write-PolicyValue -Value 1
        Restart-SearchService
        Assert-ScopeIs -Included $false
    }

    if ($RebuildCatalog) { Invoke-CatalogReset }
}

# -------------------------------------------------------------------------------------------
# Verification. It runs after -Execute as well as on its own, because a script that reports what
# it set has reported nothing: the service caches its scope, the gatherer decides when to act,
# and only the catalog knows what is actually in it.
# -------------------------------------------------------------------------------------------
Say ''
Say '== Verify =='

$service = Get-WSearchState
Say ('  WSearch: startMode={0} status={1} indexerRunning={2}' -f
    $service.StartMode, $service.Status, $service.IndexerRunning)

$policy = Get-PolicyValue
if ($null -eq $policy) { Say "  policy ${SearchPolicyName}: ABSENT" }
else { Say "  policy ${SearchPolicyName}: $policy" }

$scopeView = Get-ScopeView
Say ('  scope URL for this account: {0}' -f $scopeView.Url)
if ($scopeView.Reached) {
    Say ('  the service says: included={0} reason={1}' -f $scopeView.Included, $scopeView.ReasonName)
    foreach ($r in $scopeView.Rules) { Say ('    rule (service): {0} included={1} default={2}' -f $r.Url, $r.Included, $r.IsDefault) }
    foreach ($r in $scopeView.Roots) { Say ('    root (service): {0} providesNotifications={1}' -f $r.Url, $r.ProvidesNotifications) }
    if ($scopeView.Rules.Count -eq 0 -and $scopeView.Reason -eq 2 -and -not $scopeView.Included) {
        # MEASURED 2026-09-24: EnumerateScopeRules does not list a mapi16 user rule once it EXCLUDES.
        Say '    EnumerateScopeRules lists no mapi rule - it omits an excluding user rule (measured); the reason'
        Say '    USER above, and the registry below, are the evidence that the rule is there'
    }
    elseif ($scopeView.Rules.Count -eq 0) { Say '    no mapi rule of any kind in the service' }
}
else { Say ('  the service''s scope view could NOT be read: {0}' -f $scopeView.Error) }

$registryRules = Get-MapiRegistryRules
foreach ($r in $registryRules) {
    Say ('    registry (read-only): {0}\{1} include={2} suppress={3} default={4} policy={5} thisAccount={6}' -f
        $r.Container, $r.KeyName, $r.Include, $r.Suppress, $r.Default, $r.Policy, $r.IsThisUser)
}

Say ''
Say '  Probing the CATALOG. Two readings, so "not indexed" can be told apart from "not indexed yet".'
Say ('  Second reading in {0} minute(s). INDEXED needs at least {1} Outlook row(s).' -f $SettleMinutes, $MinimumOutlookRows)
if ($WaitMinutes -gt 0) {
    if ($SettleMinutes -lt 1) { throw '-WaitMinutes needs -SettleMinutes of at least 1: it keeps taking readings that far apart.' }
    Say ('  -WaitMinutes {0}: while the verdict is SETTLING, keep reading every {1} minute(s) until it is an answer.' -f $WaitMinutes, $SettleMinutes)
}

$first = Read-IndexState
Show-IndexState -State $first -Label 'reading 1'
$waitStarted = Get-Date
$waitDeadline = $waitStarted.AddMinutes($WaitMinutes)
$readingNumber = 1

while ($true) {
    if ($SettleMinutes -gt 0) { Start-Sleep -Seconds ($SettleMinutes * 60) }
    $readingNumber++
    $second = Read-IndexState
    Show-IndexState -State $second -Label "reading $readingNumber"
    # The service's view is re-read with every reading: a waiting run must not judge the catalog of
    # minute twenty against the scope of minute zero.
    $scopeView = Get-ScopeView
    # And so is the service itself. MEASURED 2026-09-24: a -Verify started seconds after a guest
    # restart read WSearch as Stopped, then reached the catalog on both readings (IDLE, 431 items) -
    # the first call through the API starts the service - and was judged NO-INDEXER on the stale state.
    $serviceNow = Get-WSearchState
    if ($serviceNow.Status -ne $service.Status -or $serviceNow.IndexerRunning -ne $service.IndexerRunning) {
        Say ('    WSearch is now: startMode={0} status={1} indexerRunning={2}' -f $serviceNow.StartMode, $serviceNow.Status, $serviceNow.IndexerRunning)
    }
    $service = $serviceNow
    $verdict = Get-Verdict -First $first -Second $second -Scope $scopeView -PolicyValue $policy -Service $service -MinimumRows $MinimumOutlookRows
    if ($verdict -ne 'SETTLING' -or $WaitMinutes -le 0 -or (Get-Date) -ge $waitDeadline) { break }
    Say ("    SETTLING at reading {0}, {1:N1} min in - reading again in {2} minute(s)" -f $readingNumber, ((Get-Date) - $waitStarted).TotalMinutes, $SettleMinutes)
    $first = $second
}
if ($WaitMinutes -gt 0) {
    Say ("  waited {0:N1} min over {1} readings; the verdict below is from the last two" -f ((Get-Date) - $waitStarted).TotalMinutes, $readingNumber)
}

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
        Say '  CONFIRM IT THROUGH THE PRODUCT: outlook_health''s index.perStore[] must list every store'
        Say '  with inLocalIndex:false and no newestIndexedUtc, index.provider must NOT start with'
        Say '  "unavailable", and index.wSearchStartMode must still say "automatic".'
        if ($scopeView.Reason -eq 0 -and @($scopeView.Rules).Count -eq 0) {
            Say ''
            Say '  CAVEAT - ONLY THE POLICY IS EXCLUDING: the service holds no rule of any kind for this'
            Say '  account''s Outlook scope. Measured 2026-09-24 (Q69): on a guest that was never indexed the'
            Say '  policy alone does keep a non-elevated Outlook from registering itself - but it is not a'
            Say '  crawl-scope rule, so it neither excludes nor purges a scope that anything else includes.'
            Say '  Run -Execute to add the user EXCLUDE rule the service itself reports.'
        }
    }
    'NOT-IN-SCOPE' {
        Say '  NOT AN ANSWER, AND WAITING WILL NOT MAKE IT ONE. The catalog is alive, holds no Outlook'
        Say '  row, nothing excludes Outlook - and the service reports this account''s Outlook scope as'
        Say '  outside the crawl. On OutlookAI-Indexed that is a FAILURE: run -Enable -Execute.'
        Say '  (Outlook adds this rule itself only when it runs NOT ELEVATED - see the banner; every'
        Say '  Outlook the testbed starts through a RunLevel Highest task is elevated.)'
    }
    'INDEXED' {
        Say '  Outlook IS indexed on this guest and the crawl has finished: the same row count on both'
        Say '  readings and nothing queued. Correct for OutlookAI-Indexed; wrong for'
        Say '  OutlookAI-Unindexed (run -Execute there).'
    }
    'SETTLING' {
        Say '  NOT AN ANSWER. A crawl is still running (the count moved or items are queued), or an'
        Say '  exclusion landed after a crawl and the old rows are still there, or the scope is included'
        Say '  and nothing has been crawled yet. The crawl needs Outlook RUNNING and NOT ELEVATED.'
        Say '  Wait and re-run -Verify. Rows under an exclusion go only if nothing restarts the service'
        Say '  while the indexer purges them (banner); rows that outlived one: -Execute -RebuildCatalog.'
    }
    'NO-INDEXER' {
        Say '  THIS IS A FAILURE, not a quiet success, and it is the one that looks like one. There is'
        Say '  no working index here: the product takes its "index unreachable" branch instead,'
        Say '  outlook_health omits index.perStore entirely, and MailService''s non-exhaustive search'
        Say '  path does not wrap its index calls in a catch. Set WSearch back to Automatic, start it,'
        Say '  and run this again.'
    }
}

Say ''
Say 'THE REGISTRY IS NOT THE ANSWER, AND NEITHER IS THIS SCRIPT. The verdict above is a statement'
Say 'about the service and the catalog at two moments. The statement the live tier rests on is'
Say 'outlook_health''s index.perStore[], read on BOTH guests and compared - one showing'
Say 'inLocalIndex:true with a frontier, the other inLocalIndex:false with none.'

if ($Enable) {
    if ($verdict -eq 'INDEXED' -or $verdict -eq 'SETTLING') { exit 0 }
    exit 1
}
if ($Execute) {
    if ($verdict -eq 'UNINDEXED' -or $verdict -eq 'SETTLING') { exit 0 }
    exit 1
}
if ($verdict -eq 'INDEXED' -or $verdict -eq 'UNINDEXED') { exit 0 }
exit 1
