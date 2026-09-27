#Requires -Version 5.1
<#
    ============================================================================================
    2026-09-27 (Q87 (a)): THE IDENTITY STORE IS MINTED, NOT CREATED. MEASURED ON
    `OutlookAI-Unindexed` FROM CP-09-ADDIN-READY, THEN SCRIPTED HERE AS -Phase Mint AND
    -Phase CaptureMint.
    ============================================================================================

    WHY. A PST that AddStoreEx creates has no Inbox (section 4.1 of the runbook, defect 4), so the
    identity account was bound to its hidden root on both guests, and CaptureStore now refuses such a
    store. The maintainer chose direction (a) of Docs/live-tier-on-the-vm.md section 3b: let Outlook
    MINT the store as the DEFAULT store of a throwaway account-less profile - `OUTLOOK.EXE /PIM <name>`,
    the mechanism that gave the hub its full folder set - then name it and attach it to the tier
    profile, where it keeps its Inbox.

    MEASURED, 2026-09-27 (every step in session 1 at RunLevel Highest, every Outlook close a graceful
    quit through Testbed/host/Restart-Guest.ps1, every read through a repository tool):
      * `OUTLOOK.EXE /PIM IdentityMint` opened "Outlook Today" with no dialog, minted
        C:\OutlookAI-Tier\Outlook Data File - IdentityMint.pst (ForcePSTPath), and left the default
        profile OutlookAI-Tier. The store was named 'Outlook Data File'.
      * In that profile the store has a designated, visible Inbox and a designated Drafts - CaptureStore
        read it as `Inbox designated=True, EntryID 24 bytes, name 'Inbox', visible=True; Drafts
        designated` - and every default folder but Junk Email, every one empty.
      * Rename-OutlookStore.ps1 there: 'Outlook Data File' -> 'identity@vm.invalid'.
      * Attached to the tier profile by Add-OutlookPstStore.ps1 (AddStoreEx opens the existing file):
        before and after, 'identity@vm.invalid'. CaptureStore FROM THE TIER PROFILE then PASSED with
        the same line, the Inbox's PST node id 0x8082; Bind re-pointed the account from the tier store;
        Verify: `the identity account delivers into 'Inbox' (visible=True, PST node id 0x8082)`, two
        accounts on two distinct stores, both SmtpAddress reads, and no logon dialog once
        New-TierProfile.ps1 -StoreSinkPassword had run.
      * The store's display name at each stage, for Q92: after the mint 'Outlook Data File', after the
        rename 'identity@vm.invalid', after the attach 'identity@vm.invalid' (Store.DisplayName over
        COM). And in the registry afterwards, the name each PROFILE holds for the file - its service
        section's PR_DISPLAY_NAME_W, the name the index was measured to use (runbook section 8 item
        22): 'identity@vm.invalid' in the tier profile AND in the mint profile.
      * The same route again from the same checkpoint, run as the phases below (pass 2): the same
        lines, mask 0xFF, 14 folders counted and 0 items at CaptureMint, and the sink logged
        `read USER identity` / `read PASS any-value` at the start after the password was stored.

    THE ROUTE, as this script runs it (Outlook closed / running as each phase says; it never starts,
    quits or kills Outlook itself - start it in session 1 through Register-InteractiveTask.ps1, at
    -RunLevel Limited on the indexed guest, and close it with Testbed/host/Restart-Guest.ps1 from the
    host). Set-OutlookProgrammaticAccess.ps1 -Execute must already have run - see the indexed guest's
    run below:

      -Phase Mint -Execute          Outlook CLOSED. Refuses a pending ImportPRF (a /PIM start would
                                    process it, unmeasured), a mint profile that already exists (/PIM
                                    would open it, not mint) and a missing ForcePSTPath. Records the
                                    PSTs already in ForcePSTPath.
      <start OUTLOOK.EXE /PIM IdentityMint, ~90 s>
      -Phase CaptureMint -Execute   Outlook RUNNING on the mint profile. Finds the one store it holds -
                                    the default, a NEW file in ForcePSTPath - requires a designated,
                                    visible Inbox and not one item in any folder, and records its path.
      Rename-OutlookStore.ps1 -StoreFilePath <minted> -DisplayName identity@vm.invalid -Execute
      <quit: Restart-Guest.ps1 -Execute>
      -Phase Import -Execute        Outlook CLOSED (skip it if the account exists). Then start Outlook.
      Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName identity@vm.invalid
                              -Path <minted> -Execute                     Outlook RUNNING (tier)
      -Phase CaptureStore -Execute  Outlook RUNNING (tier) - the question (a) turned on.
      <quit: Restart-Guest.ps1 -Execute -CancelLogonPrompt>
      -Phase Bind -Execute, then New-TierProfile.ps1 -StoreSinkPassword -Execute   Outlook CLOSED
      <start Outlook>
      -Phase Verify -TrySmtpAddress

    From here on the phases take the identity store's path from the mint record (-MintRecordPath)
    unless -PstPath names one: the file name is Outlook's choice, read off the guest, not assumed.
    A guest whose ImportPRF is already pending - CP-09-ADDIN-READY on both guests - completes that
    import first: start Outlook once on the tier profile, quit it (-CancelLogonPrompt), then Mint.

    RUN AGAIN 2026-09-27 ON `OutlookAI-Indexed`, REBUILT FROM CP-08B-RESTORED-BEFORE-IDENTITY (no
    ImportPRF pending there, so -Phase Mint ran first), to checkpoint CP-09C-IDENTITY-REAL-INBOX -
    the same lines as above: CaptureMint 'Outlook Data File' at ...\Outlook Data File -
    IdentityMint.pst, mask 0xFF, Inbox designated and visible, 14 folders, 0 items; the rename to
    'identity@vm.invalid'; the attach, before and after 'identity@vm.invalid'; CaptureStore PASSED
    from the tier profile (Inbox EntryID 24 bytes, node 0x8082); Bind re-pointed the account from
    the tier store; Verify `delivers into 'Inbox' (visible=True, PST node id 0x8082)`, two accounts
    on two distinct stores, both SmtpAddress reads. Two things that guest added:
      * EVERY OUTLOOK THERE STARTS NOT ELEVATED (an elevated one never feeds the index), so every
        phase that attaches ran through Register-InteractiveTask.ps1 -RunLevel Limited, and the
        /PIM start too: Start-OutlookUnelevated.ps1 opens only a profile that already exists, so the
        mint start was a -RunLevel Limited job running `Start-Process OUTLOOK.EXE /PIM IdentityMint`
        ('Outlook Today - Outlook' in 6 s, no dialog, the new OUTLOOK.EXE's token not elevated).
      * APPLY Set-OutlookProgrammaticAccess.ps1 (Q80) BEFORE THIS ROUTE. That guest had not had it
        yet, and CaptureStore's COM reads include members the Object Model Guard protects: about two
        minutes after a boot the guard prompt came up ("A program is trying to access email
        address information ...") and the read blocked behind it - these phases have no deadline -
        until the job's time limit ended the job. With the values written, CaptureStore passed. The
        unindexed guest never met this only because it had Q80 from CP-07.

    WHAT IT LEAVES BEHIND: the mint profile. A profile has no free delete route (runbook section 1),
    and nothing opens it again. It still names the minted file; the tier profile now does too.

    NOT DONE HERE: a guest that already has an AddStoreEx identity.pst attached as
    'identity@vm.invalid' (both guests from their CP-10 on) cannot attach the minted store under the
    same name - Add-OutlookPstStore.ps1 refuses a taken name - and nothing in this repository removes a
    store from a profile. Rebuild such a guest from its checkpoint before the identity account.

    ============================================================================================
    RUN 2026-09-24 ON `OutlookAI-Indexed`, END TO END, AND IT WORKS - WITH ONE STEP THAT IS NOT
    DOCUMENTED BY MICROSOFT. READ WHICH ONE BEFORE YOU TRUST IT.
    ============================================================================================

    WHAT IT BUILDS. The IDENTITY account of Docs/live-tier-on-the-vm.md section 2.8b: a second
    POP3 account in the tier profile, with ITS OWN delivery store named as its address
    (identity@vm.invalid), so the two identity tests assert instead of printing PROVED NOTHING.
    No GUI and no paid component - the repository's Dependencies rule.

    WHAT WAS MEASURED BEFORE THIS SCRIPT EXISTED (Q64, 2026-09-24, Office LTSC 2024 16.0.17932),
    and it decides its shape:

      * A .prf CAN create the account. `OverwriteProfile=Append` into the existing tier profile
        added it (a new POP3 subkey under 9375CFF0413111d3B88A00104B2A6676), left the tier account
        and its store untouched, and made no "Backup Of" profile.
      * A .prf CANNOT give it its own store. Outlook bound the appended account, at import time,
        to the profile's DEFAULT store - the tier store - and minted nothing. A second variant, both
        accounts in one freshly built profile, got ONE minted store (Outlook1.pst) bound to BOTH.
        "A store Outlook mints is a store Outlook binds" holds - but Outlook mints one store per
        profile, only when the profile has none, and binds every account a .prf adds to it.
      * The binding is two values in the account's registry subkey, written by Outlook itself:
            Delivery Store EntryID   REG_BINARY = the store's Store.StoreID, byte for byte
            Delivery Folder EntryID  REG_BINARY = that store's Inbox EntryID, byte for byte
        (PROP_ACCT_DELIVERY_STORE / PROP_ACCT_DELIVERY_FOLDER; on this build the values carry
        these NAMES, not the 00180102-style tags.) Measured by comparing the tier account's values
        with what COM reports for the tier store and its Inbox: identical.

    SO IT RUNS IN FOUR PHASES, because the steps alternate between "Outlook must be closed" and
    "Outlook must be running", and this script never starts, quits or kills Outlook itself:

      -Phase Import       Outlook CLOSED. Renders identity-account.prf, points ImportPRF at it,
                          deletes Setup\First-Run and \FirstRun. [documented mechanism, MEASURED]
                          Then: start Outlook once and let it import.
      (existing script)   Outlook RUNNING. Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier
                          -DisplayName identity@vm.invalid -Path C:\OutlookAI-Tier\identity.pst
                          -Execute  - AddStoreEx, then a root-folder rename. [MEASURED]
      -Phase CaptureStore Outlook RUNNING. Reads that store's StoreID and Inbox EntryID over COM
                          and writes them to a JSON file beside the PST. Reads only.
                          Then: close Outlook - a graceful Quit() in session 1, or restart the
                          guest; never taskkill it.
      -Phase Bind         Outlook CLOSED. Writes those two byte strings into the identity
                          account's two registry values. THIS IS THE UNDOCUMENTED STEP. It writes
                          only bytes Outlook produced, only into values Outlook created, and reads
                          them back; the object model offers no setter (Account.DeliveryStore is
                          read-only) and Microsoft documents only the GUI's "Change Folder".
      (existing script)   Outlook CLOSED. New-TierProfile.ps1 -StoreSinkPassword -Execute - the
                          account's POP3 password, or Outlook prompts for it. [MEASURED]
      -Phase Verify      Reads over COM what NewDraft needs, for EVERY account: DeliveryStore,
                          its Drafts folder, and that no two accounts share a store. Reads only.

    MEASURED RESULT of that sequence on OAI-INDEXED (driven from checkpoint
    CP-08B-RESTORED-BEFORE-IDENTITY): Accounts.Count = 2, Stores.Count = 2; the tier account on
    tier@vm.invalid (C:\OutlookAI-Tier\Outlook.pst), the identity account on identity@vm.invalid
    (C:\OutlookAI-Tier\identity.pst), each Drafts folder resolving, two distinct delivery stores
    for two accounts - and the same after a guest restart, i.e. Outlook neither rejected nor
    rewrote the binding across its own shutdown and a fresh start.

    RUN AGAIN 2026-09-24 ON `OutlookAI-Unindexed` (OAI-UNINDEXED), FROM CP-09-ADDIN-READY, AND IT
    WORKS THERE TOO - checkpoint CP-10-IDENTITY-ACCOUNT. Same sequence, same result: two POP3
    accounts, `OutlookAI tier sink` on tier@vm.invalid (Outlook.pst) and `OutlookAI identity sink`
    on its own identity@vm.invalid (identity.pst), Drafts resolving on both, two distinct delivery
    stores - and the same after Outlook's own graceful quit and a fresh start. Before Bind the
    identity account sat on the tier store's EntryID, exactly as Q64 predicts. Three things that
    run added, all measured on that guest:
      * THE IDENTITY ACCOUNT PROMPTS FOR ITS POP3 PASSWORD. The start that imports it raises an
        "Internet Email - identity" logon dialog (user name filled, password empty), because the
        .prf stores no password - the same finding as the tier account's (New-TierProfile.ps1's
        banner). So after Bind, still with Outlook closed, run
        `New-TierProfile.ps1 -StoreSinkPassword -Execute`: it stores the password on EVERY
        account that polls the sink, the identity account included. Proven: the next start raised
        no dialog, and the sink's debug log shows `read USER identity` then `read PASS any-value`
        and a STAT. The NEXT lines below now say so.
      * "CLOSE OUTLOOK" DOES NOT NEED A GUEST RESTART. A graceful quit in session 1 - attach,
        refuse while an Inspector is open or an Outbox holds anything, release every COM reference
        but the Application, Application.Quit(), wait for OUTLOOK.EXE to exit, never kill it - is
        enough for Bind, and took 2-3 s each time. Cancel the identity logon dialog first (post
        IDCANCEL to it: nothing typed, nothing stored); a modal Outlook dialog is known on this
        guest to make Quit() be ignored (the Object Model Guard prompt did exactly that).
      * ACCOUNT.SMTPADDRESS READS. With Set-OutlookProgrammaticAccess.ps1 applied (Q80),
        -TrySmtpAddress returned tier@vm.invalid and identity@vm.invalid with no prompt, before and
        after the restart. Without Q80 the section below still holds.

    AND ONE THING THAT RUN GOT WRONG, found afterwards (Docs/live-tier-on-the-vm.md section 4.1,
    step 6, defect 4):
      * THE "INBOX ENTRYID" CAPTURED IS THE PST'S ROOT. identity.pst is attached by AddStoreEx and
        has no Inbox; Store.GetDefaultFolder(6) on it returns the PST's non-IPM root folder (NID
        0x122, no display name - its EntryID ends 22010000). CaptureStore recorded that, Bind wrote
        it into `Delivery Folder EntryID`, and Verify could not see it, because it checked the
        delivery STORE and its Drafts. On both guests, then, POP3 mail for this account would be
        filed in a folder Outlook's folder tree does not show.
      * CAPTURESTORE AND VERIFY WERE NOT PURE READS. GetDefaultFolder(16) creates a missing Drafts
        folder on such a PST; the identity PST's Drafts very likely came from these two phases.

    WHAT CHANGED BECAUSE OF IT (2026-09-24, later; host-side, SelfTest only - NOT YET RUN ON A GUEST):
      * CaptureStore finds the Inbox only through the store's PR_VALID_FOLDER_MASK - no Inbox bit,
        no Inbox, and the lookup that would hand back the root is never made - and REFUSES unless it
        is a NAMED folder under the store's root folder whose EntryID's node id is not one of the
        PST's fixed non-Inbox folders (the root 0x122, Top of Outlook data file 0x8022, the search
        roots 0x8042 and 0x8062). Bind refuses the same on the captured bytes, and refuses a capture
        written before these checks. Verify now reads the identity account's `Delivery Folder
        EntryID` and FAILS on the root. Get-DeliveryFolderRefusal is the one rule; -SelfTest pins it.
      * No phase asks for Drafts by the creating lookup any more: its designation, PR_IPM_DRAFTS_ENTRYID,
        is read off the Inbox or the store object and opened by EntryID. Not designated is reported,
        not failed - the product's new_draft makes Drafts on first use, in a store it may write.
      * So on the guests as they stand (identity.pst attached by AddStoreEx, no Inbox) CaptureStore
        REFUSES. That is the point: HOW identity.pst gets a real, designated Inbox was an open
        question - DECIDED 2026-09-27 (Q87 (a)) and measured: the store is minted, the banner at the
        top of this file.

    WHAT IS STILL NOT KNOWN, stated where it matters:
      * Account.SmtpAddress over COM ON A GUEST WITHOUT Q80. It is on Microsoft's list of members
        protected by the Object Model Guard, and on these guests the guard prompts ("A program is
        trying to access email address information...") because Windows Security Center reports
        Defender's signatures out of date - the guests have no network. -TrySmtpAddress attempts it
        last, in its own job, and reports a block as a block. The registry's `Email` value is read
        regardless. Set-OutlookProgrammaticAccess.ps1 removes the prompt (measured, above).
      * Whether a later Outlook REPAIR or an account edit in the GUI rewrites the binding. Not tried.
      * The signature section 2.8b also wants on this account. Not done here: Set-AccountSignature.ps1.
        Run on OAI-UNINDEXED 2026-09-24 it reported "Verified" and was NOT: the shipped
        manage_signature bound 'Identity' to the identity PST's DATA-FILE entry (subkey 00000005,
        whose 'Account Name' is the store name identity@vm.invalid), not to this account (00000004,
        'Account Name' = 'OutlookAI identity sink'). Docs/live-tier-on-the-vm.md section 2.8b.

    WHAT IT NEVER DOES: start, quit or kill Outlook; create, modify, move or delete an item; write to
    any profile other than -ProfileName (-Phase CaptureMint READS the mint profile's one store: its
    folders, their item counts, its Inbox designation - nothing else), or touch any account other
    than the one whose Email is -EmailAddress. Every write is under HKCU\...\Office\<ver>\Outlook,
    plus the rendered .prf, the JSON capture and the mint record under -WorkDir.

    THE GUARD: Assert-TestbedGuest from OutlookMapiInterop.ps1 (dot-sourced; stage it beside this
    script). It refuses anywhere not logged on as vmadmin - on the maintainer's workstation this
    would be operating on a real profile carrying real delegate mailboxes.

    Windows PowerShell 5.1 - no ternary, no `??`. COM phases run in SESSION 1
    (Register-InteractiveTask.ps1); Outlook cannot start in session 0.

.PARAMETER Phase
    Mint, CaptureMint, Import, CaptureStore, Bind or Verify - see above. Without -Execute, Mint,
    CaptureMint, Import, CaptureStore and Bind print what they would do and change nothing.

.PARAMETER PstPath
    The identity store's file. Leave it out: the phases after -Phase CaptureMint take it from the
    mint record, because Outlook, not this script, names a minted file.

.PARAMETER MintProfileName
    The throwaway account-less profile `OUTLOOK.EXE /PIM` creates to mint the store in. It must not
    exist yet - /PIM opens an existing profile instead of minting.

.PARAMETER MintRecordPath
    Where -Phase Mint records the PSTs already in ForcePSTPath, and -Phase CaptureMint the minted one.

.PARAMETER SelfTest
    Pure: renders the template, and drives the account-selection, EntryID-sanity, mint and verdict
    decisions with synthetic inputs. No registry, no COM, no files written. Runs anywhere.

.PARAMETER TrySmtpAddress
    With -Phase Verify: also read Account.SmtpAddress, last, in its own job with a deadline.

.EXAMPLE
    .\Add-IdentityAccount.ps1 -SelfTest
    .\Add-IdentityAccount.ps1 -Phase Mint -Execute
    <start OUTLOOK.EXE /PIM IdentityMint in session 1>
    .\Add-IdentityAccount.ps1 -Phase CaptureMint -Execute
    .\Rename-OutlookStore.ps1 -StoreFilePath '<the minted file CaptureMint printed>' -DisplayName identity@vm.invalid -Execute
    <quit Outlook: Testbed/host/Restart-Guest.ps1 -VMName <guest> -Execute, from the host>
    .\Add-IdentityAccount.ps1 -Phase Import -Execute
    <start Outlook on the tier profile>
    .\Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName identity@vm.invalid -Path '<the minted file>' -Execute
    .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute
    <quit Outlook: Restart-Guest.ps1 -VMName <guest> -Execute -CancelLogonPrompt>
    .\Add-IdentityAccount.ps1 -Phase Bind -Execute
    .\New-TierProfile.ps1 -StoreSinkPassword -Execute
    <start Outlook on the tier profile>
    .\Add-IdentityAccount.ps1 -Phase Verify -TrySmtpAddress
#>
[CmdletBinding()]
param(
    [ValidateSet('Mint', 'CaptureMint', 'Import', 'CaptureStore', 'Bind', 'Verify')] [string] $Phase,
    [string] $ProfileName      = 'OutlookAI-Tier',
    [string] $EmailAddress     = 'identity@vm.invalid',
    [string] $StoreDisplayName = 'identity@vm.invalid',
    [string] $AccountName      = 'OutlookAI identity sink',
    [string] $DisplayName      = 'OutlookAI Identity',
    [string] $Pop3User         = 'identity',
    [string] $SinkHost         = '127.0.0.1',
    [int]    $Pop3Port         = 110,
    [int]    $SmtpPort         = 25,
    [string] $WorkDir          = 'C:\OutlookAI-Tier',
    [string] $PstPath,
    [string] $MintProfileName  = 'IdentityMint',
    [string] $MintRecordPath   = 'C:\OutlookAI-Tier\identity-mint.json',
    [string] $IdsPath          = 'C:\OutlookAI-Tier\identity-store-ids.json',
    [string] $TemplatePath,
    [string] $OfficeVersion    = '16.0',
    [string[]] $ExpectedUser   = @('vmadmin'),
    [switch] $Execute,
    [switch] $SelfTest,
    [switch] $TrySmtpAddress
)

$ErrorActionPreference = 'Stop'
if (-not $TemplatePath) { $TemplatePath = Join-Path $PSScriptRoot 'identity-account.prf' }

$OutlookKey   = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook"
$SetupKey     = "$OutlookKey\Setup"
$ProfilesKey  = "$OutlookKey\Profiles"
$AcctMgrName  = '9375CFF0413111d3B88A00104B2A6676'
$Pop3Clsid    = '{ED475411-B0D6-11D2-8C3B-00104B2A6676}'   # CLSID_OlkPOP3Account, MEASURED on both guests
$StoreValue   = 'Delivery Store EntryID'                    # MEASURED name on 16.0.17932
$FolderValue  = 'Delivery Folder EntryID'                   # MEASURED name on 16.0.17932
$RenderedPrf  = Join-Path $WorkDir 'identity-account.prf'

# PR_VALID_FOLDER_MASK and its Inbox bit (MAPIDefS.h FOLDER_IPM_INBOX_VALID), read as
# OutlookAI.Core's SpecialFolders reads them: an Inbox is proven present before it is opened.
$ValidFolderMaskSchema = 'http://schemas.microsoft.com/mapi/proptag/0x35DF0003'
$FolderIpmInboxValid   = 0x02
# PR_IPM_DRAFTS_ENTRYID - where Drafts is designated (on the Inbox, or the store object).
$DraftsEntryIdSchema   = 'http://schemas.microsoft.com/mapi/proptag/0x36D70102'

# The PST's fixed folders that are NOT an Inbox, by node id (MS-PST section 2.7.3: the root folder,
# and the three folders every PST's root holds). A delivery folder with one of these ids is the
# 2026-09-24 defect, whatever else is true of it.
$NonInboxPstNids = @{
    0x122  = 'the PST''s non-IPM ROOT folder'
    0x8022 = 'Top of Outlook data file (the IPM subtree root)'
    0x8042 = 'the search root'
    0x8062 = 'the spam search folder'
}

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# =============================================================================================
# PURE DECISIONS. No registry, no COM, no files. -SelfTest drives every one of them.
# =============================================================================================

function Get-RenderedIdentityPrf {
    param([string] $TemplateText, [hashtable] $Tokens)
    $text = $TemplateText
    foreach ($k in $Tokens.Keys) { $text = $text.Replace($k, [string]$Tokens[$k]) }
    $left = [regex]::Matches($text, '\{\{[A-Z0-9_]+\}\}')
    if ($left.Count -gt 0) {
        throw ('The template still holds unsubstituted tokens: ' + ((@($left | ForEach-Object { $_.Value }) | Sort-Object -Unique) -join ', '))
    }
    # ASCII + CRLF, no BOM - the shape every measured .prf import on these guests used.
    $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    if ($text -match '[^\x00-\x7F]') { throw 'The rendered .prf holds non-ASCII text.' }
    return $text
}

function ConvertFrom-HexString {
    param([string] $Hex)
    if (-not $Hex -or ($Hex.Length % 2) -ne 0 -or $Hex -notmatch '^[0-9A-Fa-f]+$') { throw "Not an even-length hex string: '$Hex'" }
    $bytes = New-Object byte[] ($Hex.Length / 2)
    for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [Convert]::ToByte($Hex.Substring($i * 2, 2), 16) }
    return ,$bytes
}

function ConvertTo-HexString {
    param([byte[]] $Bytes)
    if ($null -eq $Bytes) { return '' }
    return (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '')
}

# A PST store's EntryID carries its file path as UTF-16. Refuse bytes that do not name the PST we
# mean: binding an account to the wrong store is exactly the silent failure section 2.8 warns of.
function Test-StoreEntryIdNamesPath {
    param([byte[]] $EntryId, [string] $Path)
    if ($null -eq $EntryId -or $EntryId.Length -lt 24) { return $false }
    $want = $Path.ToLowerInvariant()
    foreach ($offset in @(0, 1)) {
        $text = [Text.Encoding]::Unicode.GetString($EntryId, $offset, $EntryId.Length - $offset)
        if ($text.ToLowerInvariant().Contains($want)) { return $true }
    }
    return $false
}

# Rows: objects with KeyName, Clsid, Email. Exactly one POP3 row with this Email, or a refusal.
function Select-IdentityAccountRow {
    param([object[]] $Rows, [string] $Email)
    $hits = @($Rows | Where-Object { $_.Email -is [string] -and $_.Email -ceq $Email })
    if ($hits.Count -eq 0) { return [pscustomobject]@{ Row = $null; Refusal = "No account in the profile has Email '$Email'. Run -Phase Import and start Outlook once first." } }
    if ($hits.Count -gt 1) { return [pscustomobject]@{ Row = $null; Refusal = "$($hits.Count) accounts have Email '$Email' ($(($hits | ForEach-Object { $_.KeyName }) -join ', ')). Refusing to guess which one to bind." } }
    if ($hits[0].Clsid -ne $Pop3Clsid) { return [pscustomobject]@{ Row = $null; Refusal = "The account with Email '$Email' is not a POP3 account (clsid $($hits[0].Clsid))." } }
    return [pscustomobject]@{ Row = $hits[0]; Refusal = $null }
}

# A PST folder's EntryID is 24 bytes - 4 flag bytes, the store's 16-byte provider UID, and the
# folder's 4-byte node id, little-endian (MS-PST's EntryID structure). The node id, or $null for an
# EntryID of any other shape (an Exchange folder's, say), which this check then has no opinion on.
function Get-PstFolderNid {
    param([byte[]] $EntryId)
    if ($null -eq $EntryId -or $EntryId.Length -ne 24) { return $null }
    return [long][BitConverter]::ToUInt32($EntryId, 20)
}

# Whether a folder is one a person can SEE: NAMED, and a descendant of the store's root folder
# (Store.GetRootFolder(), the top of the tree Outlook draws). Ancestors nearest first; a $null or
# non-string entry ends the chain - the parent was not a folder, or would not say. The same rule as
# the corpus tool's CorpusFolderVisibility.IsVisible.
function Test-FolderIsVisible {
    param([string] $Name, [object[]] $AncestorEntryIds, [string] $RootEntryId)
    if ([string]::IsNullOrWhiteSpace($Name) -or [string]::IsNullOrEmpty($RootEntryId)) { return $false }
    foreach ($id in @($AncestorEntryIds)) {
        if ($null -eq $id -or -not ($id -is [string]) -or $id.Length -eq 0) { return $false }
        if ([string]::Equals($id, $RootEntryId, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

# Why a folder may NOT be an account's delivery folder - or $null when it may. The one rule behind
# CaptureStore, Bind and Verify, written after both guests bound the identity account to the PST's
# ROOT (Docs/live-tier-on-the-vm.md section 4.1, step 6, defect 4). $Visible is Test-FolderIsVisible's
# answer; anything but $true refuses.
function Get-DeliveryFolderRefusal {
    param([string] $EntryIdHex, [string] $Name, $Visible)
    if ([string]::IsNullOrWhiteSpace($EntryIdHex)) {
        return 'the store has no Inbox to deliver into - its PR_VALID_FOLDER_MASK carries no FOLDER_IPM_INBOX_VALID bit, as on every PST attached with AddStoreEx'
    }
    $bytes = $null
    try { $bytes = ConvertFrom-HexString $EntryIdHex } catch { return "its EntryID '$EntryIdHex' is not hex" }
    $nid = Get-PstFolderNid $bytes
    if ($null -ne $nid -and $NonInboxPstNids.ContainsKey([int]$nid)) {
        return ("its EntryID names {0} (node id 0x{1:X}), not an Inbox - mail delivered there lands where Outlook's folder tree does not show it" -f $NonInboxPstNids[[int]$nid], $nid)
    }
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'it has no display name - no folder anyone can pick out' }
    if ($Visible -ne $true) { return 'it is not a descendant of the store''s root folder, so Outlook''s folder tree does not show it' }
    return $null
}

# Accounts: objects with Name, DeliveryPath (null when DeliveryStore is NULL), DeliveryStoreId,
# DraftsOk (Drafts designated and it opened) and DraftsAbsent (not designated at all - the product's
# new_draft makes it on first use, so that is reported, not failed). The identity account must have
# its OWN store at -PstPath, shared with no other account, and a delivery FOLDER nothing refuses.
function Get-IdentityVerdict {
    param([object[]] $Accounts, [string] $IdentityName, [string] $Path, [string] $DeliveryFolderRefusal)
    $problems = @()
    $mine = @($Accounts | Where-Object { $_.Name -eq $IdentityName })
    if ($mine.Count -ne 1) { return @("expected exactly one account named '$IdentityName', found $($mine.Count)") }
    $me = $mine[0]
    if (-not $me.DeliveryPath) { $problems += 'its DeliveryStore is NULL - NewDraft would fail with AccountHasNoDeliveryStore' }
    elseif ($me.DeliveryPath -ine $Path) { $problems += "its DeliveryStore is '$($me.DeliveryPath)', not '$Path'" }
    if (-not $me.DraftsOk -and $me.DraftsAbsent -ne $true) { $problems += 'its delivery store designates a Drafts folder that did not open' }
    if (-not [string]::IsNullOrEmpty($DeliveryFolderRefusal)) { $problems += "its delivery FOLDER is refused: $DeliveryFolderRefusal" }
    foreach ($other in @($Accounts | Where-Object { $_.Name -ne $IdentityName })) {
        if ($me.DeliveryStoreId -and $other.DeliveryStoreId -eq $me.DeliveryStoreId) { $problems += "it SHARES its delivery store with '$($other.Name)'" }
        if (-not $other.DeliveryPath) { $problems += "'$($other.Name)' has no delivery store" }
    }
    return $problems
}

# ---- The mint (Q87 (a), 2026-09-27) --------------------------------------------------------------

# Why -Phase Mint may NOT prepare a mint - or $null when it may. The /PIM start it prepares must MINT:
# a pending ImportPRF would be processed by that same start, which nothing here has measured; a
# profile that already exists is OPENED by /PIM, not minted; and Outlook mints the file into
# ForcePSTPath, which must be set and exist.
function Get-MintPreflightRefusal {
    param([string] $ImportPrf, [bool] $MintProfileExists, [string] $MintProfileName, [string] $ForcePstPath, [bool] $ForcePstPathExists)
    if (-not [string]::IsNullOrWhiteSpace($ImportPrf)) {
        return ("an ImportPRF is pending ('$ImportPrf'), and the /PIM start would process it as well - which has not been measured. " +
            'Complete it first: start Outlook once on the default profile, then quit it gracefully (Testbed/host/Restart-Guest.ps1 -Execute -CancelLogonPrompt).')
    }
    if ($MintProfileExists) {
        return ("a profile named '$MintProfileName' already exists, and OUTLOOK.EXE /PIM opens an existing profile instead of minting one. " +
            'A profile has no free delete route; pick another -MintProfileName - or, if this route minted that profile''s store already, start it with /PIM and go on to -Phase CaptureMint.')
    }
    if ([string]::IsNullOrWhiteSpace($ForcePstPath)) {
        return 'ForcePSTPath is not set, so Outlook would mint the store under Documents\Outlook Files. New-TierProfile.ps1 -Execute writes it.'
    }
    if (-not $ForcePstPathExists) { return "ForcePSTPath '$ForcePstPath' is not a directory that exists." }
    return $null
}

# Stores: objects with FilePath and IsDefault - every store of the running MINT profile. The minted
# store is the one store a /PIM profile holds, its default, and a NEW file in ForcePSTPath - not one of
# the PSTs -Phase Mint found there before the start. Anything else refuses.
function Select-MintedStore {
    param([object[]] $Stores, [string[]] $PstsBefore, [string] $ForcePstPath)
    $all = @($Stores | Where-Object { $null -ne $_ })
    if ($all.Count -ne 1) {
        return [pscustomobject]@{ Row = $null; Refusal = "the mint profile holds $($all.Count) store(s), and a profile /PIM just made holds exactly the one it minted - something else is attached, so refusing to guess which" }
    }
    $s = $all[0]
    if ($s.IsDefault -ne $true) { return [pscustomobject]@{ Row = $null; Refusal = "its one store, '$($s.FilePath)', is not the profile's default store" } }
    if ([string]::IsNullOrWhiteSpace($s.FilePath)) { return [pscustomobject]@{ Row = $null; Refusal = 'its one store has no file path - it is not a PST' } }
    $dir = [IO.Path]::GetDirectoryName($s.FilePath)
    if ([string]::IsNullOrWhiteSpace($ForcePstPath) -or -not [string]::Equals($dir.TrimEnd('\'), $ForcePstPath.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{ Row = $null; Refusal = "its store '$($s.FilePath)' is not in ForcePSTPath '$ForcePstPath', where Outlook mints" }
    }
    $name = [IO.Path]::GetFileName($s.FilePath)
    foreach ($b in @($PstsBefore)) {
        if ($null -ne $b -and [string]::Equals([string]$b, $name, [StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{ Row = $null; Refusal = "'$name' was in ForcePSTPath before the /PIM start - it is not a store that start minted" }
        }
    }
    return [pscustomobject]@{ Row = $s; Refusal = $null }
}

# Folders: objects with Path and Items (an [int], or $null when the count would not read). A store is
# attached to a second profile by script only while it holds NO item (Docs/live-tier-on-the-vm.md
# section 2.6), and a count that would not read, or a walk that stopped at its limit, proves nothing.
function Get-MintItemsRefusal {
    param([object[]] $Folders, [bool] $Truncated)
    $all = @($Folders | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) { return 'no folder of it was counted at all' }
    $unread = @($all | Where-Object { $null -eq $_.Items })
    if ($unread.Count -gt 0) { return ("the item count of {0} folder(s) would not read ({1}), so nothing proves the store is empty" -f $unread.Count, (($unread | ForEach-Object { $_.Path }) -join ', ')) }
    $held = @($all | Where-Object { $_.Items -gt 0 })
    if ($held.Count -gt 0) {
        return ("it already holds items ({0}), and a store with items is never attached to a second profile by script (Docs/live-tier-on-the-vm.md section 2.6)" -f (($held | ForEach-Object { '{0}={1}' -f $_.Path, $_.Items }) -join ', '))
    }
    if ($Truncated) { return 'the folder walk stopped at its limit before it had counted every folder' }
    return $null
}

# Where the identity store is: -PstPath when given, else the minted file -Phase CaptureMint recorded.
# There is no default path any more: the AddStoreEx identity.pst the old default named has no Inbox.
function Resolve-IdentityPstPath {
    param([string] $BoundPath, $MintRecord)
    if (-not [string]::IsNullOrWhiteSpace($BoundPath)) { return [pscustomobject]@{ Path = $BoundPath; Source = '-PstPath'; Refusal = $null } }
    if ($null -ne $MintRecord -and $null -ne $MintRecord.PSObject.Properties['MintedPath'] -and -not [string]::IsNullOrWhiteSpace([string]$MintRecord.MintedPath)) {
        return [pscustomobject]@{ Path = [string]$MintRecord.MintedPath; Source = 'the mint record'; Refusal = $null }
    }
    return [pscustomobject]@{ Path = $null; Source = $null; Refusal = 'No identity store to work on: mint it first (-Phase Mint -Execute, start OUTLOOK.EXE /PIM, -Phase CaptureMint -Execute), which records its path - or pass -PstPath.' }
}

function Invoke-SelfTest {
    $script:pass = 0; $script:fail = 0
    function Check([string] $what, [bool] $ok) {
        if ($ok) { $script:pass++; Write-Host "  PASS  $what" } else { $script:fail++; Write-Host "  FAIL  $what" }
    }
    $tokens = @{ '{{PROFILE_NAME}}' = 'OutlookAI-Tier'; '{{ACCOUNT_NAME}}' = 'OutlookAI identity sink'; '{{POP3_HOST}}' = '127.0.0.1'; '{{SMTP_HOST}}' = '127.0.0.1'
                 '{{POP3_USER}}' = 'identity'; '{{EMAIL_ADDRESS}}' = 'identity@vm.invalid'; '{{DISPLAY_NAME}}' = 'OutlookAI Identity'; '{{POP3_PORT}}' = '110'; '{{SMTP_PORT}}' = '25' }
    $template = [IO.File]::ReadAllText($TemplatePath)
    $prf = Get-RenderedIdentityPrf -TemplateText $template -Tokens $tokens
    # String.Contains, never -like: every .prf header is [bracketed], and -like reads brackets as
    # a character class - Testbed/README.md section 4b says why that once passed for the wrong reason.
    Check 'rendered .prf appends to the tier profile' ($prf.Contains("ProfileName=OutlookAI-Tier`r`n") -and $prf.Contains("OverwriteProfile=Append`r`n"))
    Check 'rendered .prf names no DefaultStore' (-not $prf.Contains('DefaultStore='))
    Check 'rendered .prf lists no service (no PST, no second address book)' ([regex]::Matches($prf, '(?m)^Service\d+=').Count -eq 0)
    Check 'rendered .prf carries exactly one internet account, POP3' ($prf.Contains("Account1=I_Mail`r`n") -and -not $prf.Contains('Account2='))
    Check 'rendered .prf carries the identity address and is not the default account' ($prf.Contains("EmailAddress=identity@vm.invalid`r`n") -and $prf.Contains("DefaultAccount=FALSE`r`n"))
    Check 'rendered .prf is CRLF throughout' (-not ($prf -replace "`r`n", '').Contains("`n"))
    $threw = $false; try { $null = Get-RenderedIdentityPrf -TemplateText 'X={{MISSING_TOKEN}}' -Tokens @{} } catch { $threw = $true }
    Check 'an unsubstituted token refuses' $threw

    $hex = '0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E00000000'
    $pathBytes = [Text.Encoding]::Unicode.GetBytes('C:\OutlookAI-Tier\identity.pst')
    $eid = [byte[]]((ConvertFrom-HexString $hex) + $pathBytes + [byte[]](0, 0))
    Check 'hex round-trips' ((ConvertTo-HexString (ConvertFrom-HexString 'A0ff01')) -eq 'A0FF01')
    $threw = $false; try { $null = ConvertFrom-HexString 'ABC' } catch { $threw = $true }
    Check 'odd-length hex refuses' $threw
    Check 'a PST EntryID naming the path is accepted' (Test-StoreEntryIdNamesPath -EntryId $eid -Path 'c:\outlookai-tier\IDENTITY.pst')
    Check 'a PST EntryID naming another path is refused' (-not (Test-StoreEntryIdNamesPath -EntryId $eid -Path 'C:\OutlookAI-Tier\Outlook.pst'))

    $rows = @([pscustomobject]@{ KeyName = '00000001'; Clsid = $Pop3Clsid; Email = 'tier@vm.invalid' },
              [pscustomobject]@{ KeyName = '00000002'; Clsid = '{ED475414-B0D6-11D2-8C3B-00104B2A6676}'; Email = $null },
              [pscustomobject]@{ KeyName = '00000004'; Clsid = $Pop3Clsid; Email = 'identity@vm.invalid' })
    Check 'the one POP3 account with the address is selected' ((Select-IdentityAccountRow -Rows $rows -Email 'identity@vm.invalid').Row.KeyName -eq '00000004')
    Check 'no such account refuses' ($null -ne (Select-IdentityAccountRow -Rows $rows -Email 'nobody@vm.invalid').Refusal)
    Check 'the match is case-sensitive, as the tests compare it' ($null -ne (Select-IdentityAccountRow -Rows $rows -Email 'IDENTITY@vm.invalid').Refusal)
    $dupes = @($rows) + @([pscustomobject]@{ KeyName = '00000006'; Clsid = $Pop3Clsid; Email = 'identity@vm.invalid' })
    Check 'two accounts with the address refuse' ($null -ne (Select-IdentityAccountRow -Rows $dupes -Email 'identity@vm.invalid').Refusal)
    $notPop = @([pscustomobject]@{ KeyName = '00000004'; Clsid = '{ED475420-B0D6-11D2-8C3B-00104B2A6676}'; Email = 'identity@vm.invalid' })
    Check 'a non-POP3 account with the address refuses' ($null -ne (Select-IdentityAccountRow -Rows $notPop -Email 'identity@vm.invalid').Refusal)

    $tier = [pscustomobject]@{ Name = 'OutlookAI tier sink'; DeliveryPath = 'C:\OutlookAI-Tier\Outlook.pst'; DeliveryStoreId = 'AA'; DraftsOk = $true }
    $good = [pscustomobject]@{ Name = 'OutlookAI identity sink'; DeliveryPath = 'C:\OutlookAI-Tier\identity.pst'; DeliveryStoreId = 'BB'; DraftsOk = $true }
    $shared = [pscustomobject]@{ Name = 'OutlookAI identity sink'; DeliveryPath = 'C:\OutlookAI-Tier\Outlook.pst'; DeliveryStoreId = 'AA'; DraftsOk = $true }
    $null1 = [pscustomobject]@{ Name = 'OutlookAI identity sink'; DeliveryPath = $null; DeliveryStoreId = $null; DraftsOk = $false }
    Check 'own store, distinct from the tier account''s: no problems' ((@(Get-IdentityVerdict -Accounts @($tier, $good) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst')).Count -eq 0)
    Check 'bound to the tier store (what the .prf alone produces): refused' ((@(Get-IdentityVerdict -Accounts @($tier, $shared) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst')).Count -gt 0)
    Check 'NULL delivery store: refused' ((@(Get-IdentityVerdict -Accounts @($tier, $null1) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst')).Count -gt 0)
    Check 'identity account missing: refused' ((@(Get-IdentityVerdict -Accounts @($tier) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst')).Count -gt 0)

    # The delivery FOLDER (Docs/live-tier-on-the-vm.md section 4.1, step 6, defect 4). Both guests'
    # identity account was bound to the PST's non-IPM ROOT: an EntryID ending 22010000, node 0x122.
    $uid = '0000000038A1BB1005E5101AA1BB08002B2A56C2'
    $rootHex = $uid + '22010000'
    Check 'a PST folder EntryID carries its node id in its last four bytes, little-endian' ((Get-PstFolderNid (ConvertFrom-HexString $rootHex)) -eq 0x122)
    Check 'an EntryID of another shape has no PST node id' ($null -eq (Get-PstFolderNid (ConvertFrom-HexString ('00' * 46))))
    $why = Get-DeliveryFolderRefusal -EntryIdHex $rootHex -Name '' -Visible $false
    Check 'the measured binding - the PST''s ROOT - is refused, by name' ($null -ne $why -and $why.Contains('non-IPM ROOT'))
    Check 'even if it were named and reported visible' ($null -ne (Get-DeliveryFolderRefusal -EntryIdHex $rootHex -Name 'Inbox' -Visible $true))
    Check 'Top of Outlook data file is not an Inbox either' ((Get-DeliveryFolderRefusal -EntryIdHex ($uid + '22800000') -Name 'identity@vm.invalid' -Visible $true).Contains('IPM subtree root'))
    Check 'nor the search root' ($null -ne (Get-DeliveryFolderRefusal -EntryIdHex ($uid + '42800000') -Name 'Search Root' -Visible $true))
    Check 'no Inbox at all is refused, and says why' ((Get-DeliveryFolderRefusal -EntryIdHex '' -Name '' -Visible $false).Contains('FOLDER_IPM_INBOX_VALID'))
    $inboxHex = $uid + 'A2800000'
    Check 'a named, visible folder of its own is accepted' ($null -eq (Get-DeliveryFolderRefusal -EntryIdHex $inboxHex -Name 'Inbox' -Visible $true))
    Check 'a nameless one is refused' ($null -ne (Get-DeliveryFolderRefusal -EntryIdHex $inboxHex -Name ' ' -Visible $true))
    Check 'an invisible one is refused' ($null -ne (Get-DeliveryFolderRefusal -EntryIdHex $inboxHex -Name 'Inbox' -Visible $false))
    Check 'a visibility nobody established is refused' ($null -ne (Get-DeliveryFolderRefusal -EntryIdHex $inboxHex -Name 'Inbox' -Visible $null))
    Check 'hex that is not hex is refused' ($null -ne (Get-DeliveryFolderRefusal -EntryIdHex 'XYZ1' -Name 'Inbox' -Visible $true))
    Check 'a folder whose parent chain reaches the root folder is visible' (Test-FolderIsVisible -Name 'Inbox' -AncestorEntryIds @('ROOT') -RootEntryId 'root')
    Check 'at any depth' (Test-FolderIsVisible -Name 'Projects' -AncestorEntryIds @('INBOX', 'ROOT') -RootEntryId 'ROOT')
    Check 'a chain that ends before the root is not' (-not (Test-FolderIsVisible -Name 'Inbox' -AncestorEntryIds @('NONIPMROOT', $null) -RootEntryId 'ROOT'))
    Check 'nor a nameless folder, nor one whose root would not read' ((-not (Test-FolderIsVisible -Name '' -AncestorEntryIds @('ROOT') -RootEntryId 'ROOT')) -and (-not (Test-FolderIsVisible -Name 'Inbox' -AncestorEntryIds @('ROOT') -RootEntryId '')))
    Check 'a delivery-folder refusal fails the verdict' ((@(Get-IdentityVerdict -Accounts @($tier, $good) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst' -DeliveryFolderRefusal $why) -join ' ').Contains('delivery FOLDER'))
    $absentDrafts = [pscustomobject]@{ Name = 'OutlookAI identity sink'; DeliveryPath = 'C:\OutlookAI-Tier\identity.pst'; DeliveryStoreId = 'BB'; DraftsOk = $false; DraftsAbsent = $true }
    Check 'Drafts not designated is not a failure - new_draft makes it' ((@(Get-IdentityVerdict -Accounts @($tier, $absentDrafts) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst')).Count -eq 0)
    $brokenDrafts = [pscustomobject]@{ Name = 'OutlookAI identity sink'; DeliveryPath = 'C:\OutlookAI-Tier\identity.pst'; DeliveryStoreId = 'BB'; DraftsOk = $false; DraftsAbsent = $false }
    Check 'a designated Drafts that will not open is' ((@(Get-IdentityVerdict -Accounts @($tier, $brokenDrafts) -IdentityName 'OutlookAI identity sink' -Path 'C:\OutlookAI-Tier\identity.pst')).Count -gt 0)

    # The mint (Q87 (a)). Every input below is the MEASURED shape of 2026-09-27 on OAI-UNINDEXED, and
    # each refusal has its control beside it.
    $minted = 'C:\OutlookAI-Tier\Outlook Data File - IdentityMint.pst'
    $before = @('Outlook Data File - CorpusProfile.pst', 'Outlook.pst')
    Check 'the mint preflight passes the measured guest: no import pending, a new profile name, ForcePSTPath there (the control)' ($null -eq (Get-MintPreflightRefusal -ImportPrf '' -MintProfileExists $false -MintProfileName 'IdentityMint' -ForcePstPath 'C:\OutlookAI-Tier' -ForcePstPathExists $true))
    $why = Get-MintPreflightRefusal -ImportPrf 'C:\OutlookAI-Tier\identity-account.prf' -MintProfileExists $false -MintProfileName 'IdentityMint' -ForcePstPath 'C:\OutlookAI-Tier' -ForcePstPathExists $true
    Check 'a pending ImportPRF - CP-09-ADDIN-READY on both guests - refuses the mint, and says how to complete it' ($null -ne $why -and $why.Contains('ImportPRF') -and $why.Contains('-CancelLogonPrompt'))
    Check 'a mint profile that already exists refuses (/PIM would open it, not mint)' ((Get-MintPreflightRefusal -ImportPrf '' -MintProfileExists $true -MintProfileName 'IdentityMint' -ForcePstPath 'C:\OutlookAI-Tier' -ForcePstPathExists $true).Contains('already exists'))
    Check 'no ForcePSTPath refuses' ($null -ne (Get-MintPreflightRefusal -ImportPrf '' -MintProfileExists $false -MintProfileName 'IdentityMint' -ForcePstPath '' -ForcePstPathExists $false))
    Check 'a ForcePSTPath that is not there refuses' ($null -ne (Get-MintPreflightRefusal -ImportPrf '' -MintProfileExists $false -MintProfileName 'IdentityMint' -ForcePstPath 'C:\Nowhere' -ForcePstPathExists $false))

    $one = @([pscustomobject]@{ FilePath = $minted; IsDefault = $true })
    Check 'the measured minted store is found: the one default store, a new file in ForcePSTPath (the control)' ((Select-MintedStore -Stores $one -PstsBefore $before -ForcePstPath 'C:\OutlookAI-Tier').Row.FilePath -eq $minted)
    Check 'and ForcePSTPath compares without its case or a trailing backslash' ($null -eq (Select-MintedStore -Stores $one -PstsBefore $before -ForcePstPath 'c:\outlookai-tier\').Refusal)
    $two = @($one) + @([pscustomobject]@{ FilePath = 'C:\OutlookAI-Tier\Outlook.pst'; IsDefault = $false })
    Check 'a mint profile holding two stores refuses' ($null -ne (Select-MintedStore -Stores $two -PstsBefore $before -ForcePstPath 'C:\OutlookAI-Tier').Refusal)
    Check 'a store that is not the default refuses' ($null -ne (Select-MintedStore -Stores @([pscustomobject]@{ FilePath = $minted; IsDefault = $false }) -PstsBefore $before -ForcePstPath 'C:\OutlookAI-Tier').Refusal)
    Check 'a store outside ForcePSTPath refuses' ($null -ne (Select-MintedStore -Stores @([pscustomobject]@{ FilePath = 'C:\Users\vmadmin\Documents\Outlook Files\Outlook.pst'; IsDefault = $true }) -PstsBefore $before -ForcePstPath 'C:\OutlookAI-Tier').Refusal)
    $old = Select-MintedStore -Stores @([pscustomobject]@{ FilePath = 'C:\OutlookAI-Tier\OUTLOOK.PST'; IsDefault = $true }) -PstsBefore $before -ForcePstPath 'C:\OutlookAI-Tier'
    Check 'a file that was there before the start - the tier store, say - refuses, whatever its case' ($null -ne $old.Refusal -and $old.Refusal.Contains('before the /PIM start'))
    Check 'no store at all refuses' ($null -ne (Select-MintedStore -Stores @() -PstsBefore $before -ForcePstPath 'C:\OutlookAI-Tier').Refusal)

    $empty = @('Deleted Items', 'Inbox', 'Outbox', 'Sent Items', 'Calendar', 'Contacts', 'Journal', 'Notes', 'Tasks', 'Drafts', 'RSS Feeds') | ForEach-Object { [pscustomobject]@{ Path = $_; Items = 0 } }
    Check 'the measured minted store - every folder empty - may be attached (the control)' ($null -eq (Get-MintItemsRefusal -Folders $empty -Truncated $false))
    $withItem = @($empty) + @([pscustomobject]@{ Path = 'Inbox/Sub'; Items = 1 })
    $why = Get-MintItemsRefusal -Folders $withItem -Truncated $false
    Check 'one item anywhere refuses, and names the folder' ($null -ne $why -and $why.Contains('Inbox/Sub=1'))
    Check 'a count that would not read refuses' ($null -ne (Get-MintItemsRefusal -Folders (@($empty) + @([pscustomobject]@{ Path = 'Calendar'; Items = $null })) -Truncated $false))
    Check 'a walk that stopped at its limit refuses' ($null -ne (Get-MintItemsRefusal -Folders $empty -Truncated $true))
    Check 'no folder counted refuses' ($null -ne (Get-MintItemsRefusal -Folders @() -Truncated $false))

    $record = [pscustomobject]@{ MintProfile = 'IdentityMint'; MintedPath = $minted }
    Check 'the identity store is the minted file the record names, when -PstPath is not given' ((Resolve-IdentityPstPath -BoundPath '' -MintRecord $record).Path -eq $minted)
    Check '-PstPath wins over the record' ((Resolve-IdentityPstPath -BoundPath 'D:\x.pst' -MintRecord $record).Path -eq 'D:\x.pst')
    Check 'a record from -Phase Mint alone - nothing minted yet - refuses' ($null -ne (Resolve-IdentityPstPath -BoundPath '' -MintRecord ([pscustomobject]@{ MintProfile = 'IdentityMint' })).Refusal)
    Check 'no record and no -PstPath refuses - there is no default path any more' ($null -ne (Resolve-IdentityPstPath -BoundPath '' -MintRecord $null).Refusal)
    # The MEASURED delivery folder of the minted store, read from the tier profile: its Inbox, PST node
    # id 0x8082 - accepted by the same rule that refuses the AddStoreEx store's root.
    Check 'the minted store''s measured Inbox (node 0x8082) is a delivery folder nothing refuses' ($null -eq (Get-DeliveryFolderRefusal -EntryIdHex '0000000006BC1DA715E59C419E92901D3513E13682800000' -Name 'Inbox' -Visible $true))

    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed. The guest halves - the import, AddStoreEx, the COM reads and whether Outlook honours the bound values - are what -Phase proves, not this." -f $script:pass, $script:fail)
    if ($script:fail -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }
if (-not $Phase) { throw 'Pass -Phase Mint|CaptureMint|Import|CaptureStore|Bind|Verify, or -SelfTest.' }

# =============================================================================================
# GUEST ONLY FROM HERE.
# =============================================================================================
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

# The mint record: what -Phase Mint found in ForcePSTPath, and what -Phase CaptureMint minted. $null
# when there is none; a record that will not read refuses rather than being taken for none.
function Read-MintRecord {
    if (-not (Test-Path -LiteralPath $MintRecordPath)) { return $null }
    try { return (Get-Content -LiteralPath $MintRecordPath -Raw | ConvertFrom-Json) }
    catch { throw "The mint record at $MintRecordPath would not read: $($_.Exception.Message)" }
}

# The phases that work on the identity store find it here: -PstPath, else the minted file.
if (@('CaptureStore', 'Bind', 'Verify') -contains $Phase) {
    $resolvedPst = Resolve-IdentityPstPath -BoundPath $PstPath -MintRecord (Read-MintRecord)
    if ($resolvedPst.Refusal) { throw "REFUSING: $($resolvedPst.Refusal)" }
    $PstPath = $resolvedPst.Path
    Say "identity store: $PstPath (from $($resolvedPst.Source))"
}

# What a store offers an account as its delivery target, read WITHOUT the lookups that create folders:
# the Inbox only when the store's PR_VALID_FOLDER_MASK proves it - the rule OutlookAI.Core's
# SpecialFolders follows - and Drafts by its designation. Asked for without that proof,
# GetDefaultFolder(6) on an AddStoreEx PST hands back its non-IPM ROOT (measured on both guests,
# 2026-09-24), and that is what CaptureStore used to record. Shared by CaptureStore and CaptureMint.
function Read-DeliveryTarget {
    param($Store)
    $accessor = $Store.PropertyAccessor
    $mask = $null
    try { $mask = [int]$accessor.GetProperty($ValidFolderMaskSchema) } catch { $mask = $null }
    if ($null -eq $mask) { throw "The store's PR_VALID_FOLDER_MASK would not read, so nothing proves whether it has an Inbox. Refusing to capture." }
    $inboxHex = ''; $inboxName = ''; $inboxVisible = $false; $draftsHex = ''
    if (($mask -band $FolderIpmInboxValid) -ne 0) {
        $inbox = $Store.GetDefaultFolder(6)
        $inboxHex = [string]$inbox.EntryID
        $inboxName = [string]$inbox.Name
        $rootId = [string]$Store.GetRootFolder().EntryID
        $ancestors = New-Object System.Collections.Generic.List[object]
        $parent = $null
        try { $parent = $inbox.Parent } catch { $parent = $null }
        for ($depth = 0; $depth -lt 32 -and $null -ne $parent; $depth++) {
            $parentId = $null
            try { $parentId = [string]$parent.EntryID } catch { $parentId = $null }
            $ancestors.Add($parentId)
            if ($null -eq $parentId -or $parentId -ieq $rootId) { break }
            try { $parent = $parent.Parent } catch { $parent = $null }
        }
        $inboxVisible = Test-FolderIsVisible -Name $inboxName -AncestorEntryIds $ancestors.ToArray() -RootEntryId $rootId
        # Drafts by its DESIGNATION, never by GetDefaultFolder(16), which creates one a PST lacks.
        try { $draftsHex = [string]$inbox.PropertyAccessor.BinaryToString($inbox.PropertyAccessor.GetProperty($DraftsEntryIdSchema)) } catch { $draftsHex = '' }
    }
    if (-not $draftsHex) {
        try { $draftsHex = [string]$accessor.BinaryToString($accessor.GetProperty($DraftsEntryIdSchema)) } catch { $draftsHex = '' }
    }
    return [pscustomobject]@{
        Mask = $mask; InboxDesignated = (($mask -band $FolderIpmInboxValid) -ne 0)
        InboxEntryID = $inboxHex; InboxName = $inboxName; InboxVisible = $inboxVisible; DraftsEntryID = $draftsHex
    }
}

# Every folder under a store's root, breadth first, with its item count - READ ONLY: Folders and
# Items.Count, no lookup that creates anything. A count or a subfolder list that will not read is
# recorded as unread ($null), which Get-MintItemsRefusal refuses; so is stopping at the limits.
function Get-FolderItemCounts {
    param($Root, [int] $MaxDepth = 8, [int] $MaxFolders = 500)
    $rows = New-Object System.Collections.Generic.List[object]
    $truncated = $false
    $rootCount = $null
    try { $rootCount = [int]$Root.Items.Count } catch { $rootCount = $null }
    $rows.Add([pscustomobject]@{ Path = '(the root folder)'; Items = $rootCount })
    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue(@($Root, '', 0))
    while ($queue.Count -gt 0 -and -not $truncated) {
        $entry = $queue.Dequeue()
        $folder = $entry[0]; $path = [string]$entry[1]; $depth = [int]$entry[2]
        $children = $null
        try { $children = $folder.Folders } catch { $children = $null }
        if ($null -eq $children) { $rows.Add([pscustomobject]@{ Path = "$path (its subfolders)"; Items = $null }); continue }
        for ($i = 1; $i -le $children.Count; $i++) {
            if ($rows.Count -ge $MaxFolders) { $truncated = $true; break }
            $child = $children.Item($i)
            $name = [string]$child.Name
            $childPath = $name
            if ($path) { $childPath = "$path/$name" }
            $count = $null
            try { $count = [int]$child.Items.Count } catch { $count = $null }
            $rows.Add([pscustomobject]@{ Path = $childPath; Items = $count })
            if ($depth + 1 -lt $MaxDepth) { $queue.Enqueue(@($child, $childPath, ($depth + 1))) }
            else {
                # At the depth limit: a folder that has (or will not say whether it has) subfolders is
                # a part of the store this walk did not count.
                $below = -1
                try { $below = [int]$child.Folders.Count } catch { $below = -1 }
                if ($below -ne 0) { $truncated = $true }
            }
        }
    }
    return [pscustomobject]@{ Folders = $rows.ToArray(); Truncated = $truncated }
}

function Get-AccountRows {
    $mgr = Join-Path (Join-Path $ProfilesKey $ProfileName) $AcctMgrName
    if (-not (Test-Path -LiteralPath $mgr)) { throw "Profile '$ProfileName' has no account manager key ($mgr). Does the profile exist?" }
    $rows = @()
    foreach ($k in Get-ChildItem -LiteralPath $mgr) {
        $rows += [pscustomobject]@{
            KeyName = $k.PSChildName; KeyPath = $k.PSPath; Clsid = $k.GetValue('clsid'); Email = $k.GetValue('Email')
            AccountName = $k.GetValue('Account Name'); StoreBytes = $k.GetValue($StoreValue); FolderBytes = $k.GetValue($FolderValue)
        }
    }
    return $rows
}

function Show-Accounts {
    foreach ($r in Get-AccountRows) {
        if ($r.Clsid -ne $Pop3Clsid) { continue }
        Say ("  account {0}: '{1}' Email='{2}'" -f $r.KeyName, $r.AccountName, $r.Email)
        Say ("      {0} = {1}" -f $StoreValue, (ConvertTo-HexString $r.StoreBytes))
        Say ("      {0} = {1}" -f $FolderValue, (ConvertTo-HexString $r.FolderBytes))
    }
}

switch ($Phase) {
    'Mint' {
        Assert-OutlookNotRunning
        $importPrf = ''
        if (Test-Path -LiteralPath $SetupKey) { $importPrf = [string](Get-ItemProperty -LiteralPath $SetupKey).ImportPRF }
        $profileExists = Test-Path -LiteralPath (Join-Path $ProfilesKey $MintProfileName)
        $force = ''
        if (Test-Path -LiteralPath $OutlookKey) { $force = [Environment]::ExpandEnvironmentVariables([string](Get-ItemProperty -LiteralPath $OutlookKey).ForcePSTPath) }
        $forceExists = (-not [string]::IsNullOrWhiteSpace($force)) -and (Test-Path -LiteralPath $force -PathType Container)
        $refusal = Get-MintPreflightRefusal -ImportPrf $importPrf -MintProfileExists $profileExists -MintProfileName $MintProfileName -ForcePstPath $force -ForcePstPathExists $forceExists
        if ($refusal) { throw "REFUSING to prepare the mint: $refusal" }
        $before = @(Get-ChildItem -LiteralPath $force -Filter '*.pst' -File | ForEach-Object { $_.Name })
        Say "mint profile '$MintProfileName' (does not exist yet); ForcePSTPath '$force' holds $($before.Count) PST(s): $($before -join ', ')"
        if (-not $Execute) { Say "Dry run. Would record that in $MintRecordPath."; return }
        $record = [pscustomobject]@{ MintProfile = $MintProfileName; ForcePSTPath = $force; PstsBefore = $before; RecordedAt = (Get-Date).ToString('o') }
        Set-Content -LiteralPath $MintRecordPath -Value ($record | ConvertTo-Json) -Encoding UTF8
        Say "wrote $MintRecordPath"
        Say "NEXT: start  OUTLOOK.EXE /PIM $MintProfileName  in session 1 (Register-InteractiveTask.ps1) and let it settle ~90 s - it opens 'Outlook Today' with no dialog (measured) - then:"
        Say '  .\Add-IdentityAccount.ps1 -Phase CaptureMint -Execute'
    }
    'CaptureMint' {
        $record = Read-MintRecord
        if ($null -eq $record) { throw "No mint record at $MintRecordPath. Run -Phase Mint -Execute first, with Outlook closed." }
        if ([string]$record.MintProfile -ne $MintProfileName) { throw "The mint record is for profile '$($record.MintProfile)', not '$MintProfileName'." }
        $script:minted = $null
        Invoke-WithOutlookSession -Body {
            param($ns)
            if ($ns.CurrentProfileName -ne $MintProfileName) { throw "Outlook is on profile '$($ns.CurrentProfileName)', not the mint profile '$MintProfileName'. Start it as OUTLOOK.EXE /PIM $MintProfileName." }
            if ($ns.Accounts.Count -ne 0) { throw "The mint profile holds $($ns.Accounts.Count) account(s); a profile /PIM made holds none." }
            $defaultId = [string]$ns.DefaultStore.StoreID
            $rows = @()
            for ($i = 1; $i -le $ns.Stores.Count; $i++) {
                $s = $ns.Stores.Item($i)
                $rows += [pscustomobject]@{ FilePath = [string]$s.FilePath; IsDefault = ([string]$s.StoreID -eq $defaultId); Store = $s }
            }
            $sel = Select-MintedStore -Stores $rows -PstsBefore @($record.PstsBefore) -ForcePstPath ([string]$record.ForcePSTPath)
            if ($sel.Refusal) { throw "REFUSING: $($sel.Refusal)." }
            $store = $sel.Row.Store
            $script:minted = [pscustomobject]@{
                MintedPath = $sel.Row.FilePath; DisplayName = [string]$store.DisplayName
                Target = (Read-DeliveryTarget $store); Counts = (Get-FolderItemCounts -Root $store.GetRootFolder())
            }
        }
        $m = $script:minted
        $items = 0
        foreach ($f in @($m.Counts.Folders)) { if ($null -ne $f.Items) { $items += $f.Items } }
        Say ("minted store '{0}' at {1} (profile '{2}'): mask 0x{3:X}; Inbox designated={4}, name '{5}', visible={6}; Drafts {7}; {8} folder(s) counted, {9} item(s)" -f $m.DisplayName, $m.MintedPath, $MintProfileName, $m.Target.Mask, $m.Target.InboxDesignated, $m.Target.InboxName, $m.Target.InboxVisible, $(if ($m.Target.DraftsEntryID) { 'designated' } else { 'NOT designated' }), @($m.Counts.Folders).Count, $items)
        $refusal = Get-DeliveryFolderRefusal -EntryIdHex $m.Target.InboxEntryID -Name $m.Target.InboxName -Visible $m.Target.InboxVisible
        if ($refusal) { throw "REFUSING: the minted store's Inbox - $refusal. The mint route does not hold here; nothing was recorded." }
        $itemsRefusal = Get-MintItemsRefusal -Folders $m.Counts.Folders -Truncated $m.Counts.Truncated
        if ($itemsRefusal) { throw "REFUSING: $itemsRefusal. Nothing was recorded." }
        if (-not $Execute) { Say "Dry run. Would record '$($m.MintedPath)' in $MintRecordPath."; return }
        foreach ($p in @(@('MintedPath', $m.MintedPath), @('DisplayNameAfterMint', $m.DisplayName), @('ValidFolderMask', $m.Target.Mask), @('InboxName', $m.Target.InboxName), @('CapturedAt', (Get-Date).ToString('o')))) {
            $record | Add-Member -NotePropertyName $p[0] -NotePropertyValue $p[1] -Force
        }
        Set-Content -LiteralPath $MintRecordPath -Value ($record | ConvertTo-Json) -Encoding UTF8
        Say "recorded in $MintRecordPath - the phases from here take the identity store's path from it"
        Say 'NEXT, with Outlook still on the mint profile - name the store:'
        Say "  .\Rename-OutlookStore.ps1 -StoreFilePath '$($m.MintedPath)' -DisplayName $StoreDisplayName -Execute"
        Say 'THEN: quit Outlook gracefully - from the host, Testbed/host/Restart-Guest.ps1 -VMName <this guest> -Execute - and, Outlook closed:'
        Say '  .\Add-IdentityAccount.ps1 -Phase Import -Execute     (skip it if the identity account exists already)'
        Say 'THEN: start Outlook on the tier profile, and with it running:'
        Say "  .\Add-OutlookPstStore.ps1 -ProfileName $ProfileName -DisplayName $StoreDisplayName -Path '$($m.MintedPath)' -Execute"
        Say '  .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute'
    }
    'Import' {
        Assert-OutlookNotRunning
        if (-not (Test-Path -LiteralPath (Join-Path $ProfilesKey $ProfileName))) { throw "REFUSING: profile '$ProfileName' does not exist. Build the tier profile first (New-TierProfile.ps1)." }
        $default = (Get-ItemProperty -LiteralPath $OutlookKey).DefaultProfile
        if ($default -ne $ProfileName) { throw "REFUSING: the default profile is '$default', not '$ProfileName'. Outlook imports at start-up into the profile it opens; make '$ProfileName' the default first (Set-DefaultOutlookProfile.ps1)." }
        $already = @(Get-AccountRows | Where-Object { $_.Email -is [string] -and $_.Email -ceq $EmailAddress })
        if ($already.Count -gt 0) { throw "REFUSING: an account with Email '$EmailAddress' already exists ($($already[0].KeyName)). An Append import is not known to be idempotent; go on to CaptureStore / Bind / Verify instead." }
        $tokens = @{ '{{PROFILE_NAME}}' = $ProfileName; '{{ACCOUNT_NAME}}' = $AccountName; '{{POP3_HOST}}' = $SinkHost; '{{SMTP_HOST}}' = $SinkHost; '{{POP3_USER}}' = $Pop3User
                     '{{EMAIL_ADDRESS}}' = $EmailAddress; '{{DISPLAY_NAME}}' = $DisplayName; '{{POP3_PORT}}' = $Pop3Port; '{{SMTP_PORT}}' = $SmtpPort }
        $prf = Get-RenderedIdentityPrf -TemplateText ([IO.File]::ReadAllText($TemplatePath)) -Tokens $tokens
        Say "profile $ProfileName (default), account '$AccountName' <$EmailAddress>, POP3 $SinkHost`:$Pop3Port, SMTP $SinkHost`:$SmtpPort"
        Say "would write $RenderedPrf, set $SetupKey\ImportPRF to it, and delete First-Run and FirstRun"
        if (-not $Execute) { Say 'Dry run. Nothing written.'; return }
        if (-not (Test-Path -LiteralPath $WorkDir)) { New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null }
        [IO.File]::WriteAllText($RenderedPrf, $prf, (New-Object Text.ASCIIEncoding))
        if ([IO.File]::ReadAllText($RenderedPrf) -ne $prf) { throw 'The .prf read back differently from what was written.' }
        Say "wrote $RenderedPrf ($($prf.Length) chars, ASCII, CRLF) and read it back byte-identical"
        if (-not (Test-Path -LiteralPath $SetupKey)) { New-Item -Path $SetupKey -Force | Out-Null }
        New-ItemProperty -LiteralPath $SetupKey -Name 'ImportPRF' -PropertyType String -Value $RenderedPrf -Force | Out-Null
        foreach ($n in @('First-Run', 'FirstRun')) {
            $sk = Get-Item -LiteralPath $SetupKey
            if (@($sk.GetValueNames()) -contains $n) { Remove-ItemProperty -LiteralPath $SetupKey -Name $n -Force; Say "deleted Setup\$n (ImportPRF is ignored while it exists)" }
            else { Say "Setup\$n already absent" }
        }
        if ((Get-ItemProperty -LiteralPath $SetupKey).ImportPRF -ne $RenderedPrf) { throw 'ImportPRF did not read back.' }
        Say 'NEXT: start Outlook once (it imports at start-up and deletes ImportPRF itself - measured).'
        $minted = Resolve-IdentityPstPath -BoundPath $PstPath -MintRecord (Read-MintRecord)
        if ($minted.Refusal) {
            Say 'The identity store is not minted yet, and a /PIM start must not find an import pending - so let this start complete the import, quit Outlook gracefully (Restart-Guest.ps1 -Execute -CancelLogonPrompt), then:'
            Say '  .\Add-IdentityAccount.ps1 -Phase Mint -Execute'
        }
        else {
            Say 'Then, with it running:'
            Say "  .\Add-OutlookPstStore.ps1 -ProfileName $ProfileName -DisplayName $StoreDisplayName -Path '$($minted.Path)' -Execute"
            Say '  .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute'
        }
    }
    'CaptureStore' {
        $script:captured = $null
        Invoke-WithOutlookSession -Body {
            param($ns)
            if ($ns.CurrentProfileName -ne $ProfileName) { throw "Outlook is on profile '$($ns.CurrentProfileName)', not '$ProfileName'." }
            $target = [IO.Path]::GetFullPath($PstPath)
            $hit = $null
            for ($i = 1; $i -le $ns.Stores.Count; $i++) { $s = $ns.Stores.Item($i); if ($s.FilePath -and ($s.FilePath -ieq $target)) { $hit = $s } }
            if (-not $hit) { throw "No store in '$ProfileName' has FilePath '$target'. Run Add-OutlookPstStore.ps1 first." }
            if ($hit.DisplayName -cne $StoreDisplayName) { throw "The store at '$target' is named '$($hit.DisplayName)', not '$StoreDisplayName'. Name it first (Add-OutlookPstStore.ps1 renames)." }

            $t = Read-DeliveryTarget $hit
            $script:captured = [pscustomobject]@{
                FilePath = $hit.FilePath; DisplayName = $hit.DisplayName; StoreID = $hit.StoreID
                InboxEntryID = $t.InboxEntryID; InboxName = $t.InboxName; InboxVisible = $t.InboxVisible
                InboxDesignated = $t.InboxDesignated; DraftsEntryID = $t.DraftsEntryID; ValidFolderMask = $t.Mask
                Profile = $ns.CurrentProfileName; CapturedAt = (Get-Date).ToString('o')
            }
        }
        $c = $script:captured
        Say ("store '{0}' at {1}: StoreID {2} bytes; mask 0x{3:X}; Inbox designated={4}, EntryID {5} bytes, name '{6}', visible={7}; Drafts {8}" -f $c.DisplayName, $c.FilePath, ($c.StoreID.Length / 2), $c.ValidFolderMask, $c.InboxDesignated, ($c.InboxEntryID.Length / 2), $c.InboxName, $c.InboxVisible, $(if ($c.DraftsEntryID) { 'designated' } else { 'NOT designated' }))
        $refusal = Get-DeliveryFolderRefusal -EntryIdHex $c.InboxEntryID -Name $c.InboxName -Visible $c.InboxVisible
        if ($refusal) {
            throw ("REFUSING to capture '$($c.DisplayName)' as the identity account's delivery target: $refusal. " +
                'Binding it anyway is exactly the 2026-09-24 defect - POP3 mail filed where nobody can see it. A store AddStoreEx creates has no ' +
                'Inbox; the identity store is MINTED instead (-Phase Mint, then -Phase CaptureMint - the banner). Nothing was written.')
        }
        if (-not (Test-StoreEntryIdNamesPath -EntryId (ConvertFrom-HexString $c.StoreID) -Path $c.FilePath)) { throw 'The StoreID does not carry the PST path; refusing to record it.' }
        if (-not $Execute) { Say "Dry run. Would write $IdsPath."; return }
        Set-Content -LiteralPath $IdsPath -Value ($c | ConvertTo-Json) -Encoding UTF8
        Say "wrote $IdsPath"
        Say 'NEXT: close Outlook the one way the runbook allows - from the host, Testbed/host/Restart-Guest.ps1 -VMName <this guest> -Execute -CancelLogonPrompt (it cancels the identity logon dialog, quits gracefully, restarts unforced; never taskkill) - then:  .\Add-IdentityAccount.ps1 -Phase Bind -Execute'
    }
    'Bind' {
        Assert-OutlookNotRunning
        if (-not (Test-Path -LiteralPath $IdsPath)) { throw "No capture at $IdsPath. Run -Phase CaptureStore -Execute first, with Outlook running." }
        $ids = Get-Content -LiteralPath $IdsPath -Raw | ConvertFrom-Json
        if ($ids.FilePath -ine [IO.Path]::GetFullPath($PstPath)) { throw "The capture is for '$($ids.FilePath)', not '$PstPath'." }
        if ($null -eq $ids.PSObject.Properties['InboxVisible']) {
            throw "REFUSING: $IdsPath was written by a CaptureStore that did not check its Inbox is a visible folder - the version that captured the PST's ROOT on both guests (2026-09-24). Run -Phase CaptureStore -Execute again."
        }
        $refusal = Get-DeliveryFolderRefusal -EntryIdHex $ids.InboxEntryID -Name $ids.InboxName -Visible $ids.InboxVisible
        if ($refusal) { throw "REFUSING to bind: the captured delivery folder - $refusal." }
        $storeBytes = ConvertFrom-HexString $ids.StoreID
        $folderBytes = ConvertFrom-HexString $ids.InboxEntryID
        if (-not (Test-StoreEntryIdNamesPath -EntryId $storeBytes -Path $ids.FilePath)) { throw 'The captured StoreID does not carry the PST path; refusing.' }
        $sel = Select-IdentityAccountRow -Rows (Get-AccountRows) -Email $EmailAddress
        if ($sel.Refusal) { throw "REFUSING: $($sel.Refusal)" }
        $row = $sel.Row
        Say "account $($row.KeyName) '$($row.AccountName)' <$EmailAddress>"
        Say ("  before: {0} = {1}" -f $StoreValue, (ConvertTo-HexString $row.StoreBytes))
        Say ("  before: {0} = {1}" -f $FolderValue, (ConvertTo-HexString $row.FolderBytes))
        if ((ConvertTo-HexString $row.StoreBytes) -eq $ids.StoreID.ToUpperInvariant() -and (ConvertTo-HexString $row.FolderBytes) -eq $ids.InboxEntryID.ToUpperInvariant()) { Say 'already bound to that store - nothing to write'; return }
        if (-not $Execute) { Say "Dry run. Would bind it to '$($ids.DisplayName)' ($($ids.FilePath))."; return }
        Set-ItemProperty -LiteralPath $row.KeyPath -Name $StoreValue -Value $storeBytes -Type Binary
        Set-ItemProperty -LiteralPath $row.KeyPath -Name $FolderValue -Value $folderBytes -Type Binary
        $back = Get-Item -LiteralPath $row.KeyPath
        if ((ConvertTo-HexString $back.GetValue($StoreValue)) -ne $ids.StoreID.ToUpperInvariant()) { throw "$StoreValue did not read back." }
        if ((ConvertTo-HexString $back.GetValue($FolderValue)) -ne $ids.InboxEntryID.ToUpperInvariant()) { throw "$FolderValue did not read back." }
        Say "  after:  bound to '$($ids.DisplayName)' - both values read back byte-identical"
        Say 'NEXT: with Outlook still closed, store the POP3 password it will otherwise prompt for:  .\New-TierProfile.ps1 -StoreSinkPassword -Execute'
        Say 'THEN: start Outlook, then  .\Add-IdentityAccount.ps1 -Phase Verify -TrySmtpAddress'
    }
    'Verify' {
        Say "registry, profile ${ProfileName}:"
        Show-Accounts
        # The identity account's delivery FOLDER, as Outlook will use it: the bytes Bind wrote.
        $identityRow = $null
        foreach ($r in @(Get-AccountRows)) { if ($r.Clsid -eq $Pop3Clsid -and $r.Email -is [string] -and $r.Email -ceq $EmailAddress) { $identityRow = $r } }
        $folderHex = ''; $folderStoreHex = ''
        if ($null -ne $identityRow) { $folderHex = ConvertTo-HexString $identityRow.FolderBytes; $folderStoreHex = ConvertTo-HexString $identityRow.StoreBytes }
        # Every COM read in a child job with a deadline, one member per line, so a blocked call is
        # reported as exactly that - the pattern this project settled on after reads that hung in a
        # scheduled task's own runspace. NO read here asks for a folder by the lookup that creates one:
        # Drafts is found by its designation, PR_IPM_DRAFTS_ENTRYID, on the Inbox the store's
        # PR_VALID_FOLDER_MASK proves - or on the store object - and opened by EntryID.
        $job = Start-Job -ArgumentList @($ProfileName, $folderHex, $folderStoreHex) -ScriptBlock {
            param($want, $folderHex, $folderStoreHex)
            $mask = 'http://schemas.microsoft.com/mapi/proptag/0x35DF0003'
            $draftsTag = 'http://schemas.microsoft.com/mapi/proptag/0x36D70102'
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace('MAPI')
            $null = $ns.GetDefaultFolder(6)
            "PROFILE|$($ns.CurrentProfileName)"
            for ($i = 1; $i -le $ns.Stores.Count; $i++) { $s = $ns.Stores.Item($i); "STORE|$($s.DisplayName)|$($s.FilePath)" }
            for ($i = 1; $i -le $ns.Accounts.Count; $i++) {
                $a = $ns.Accounts.Item($i)
                $name = $a.DisplayName
                $ds = $a.DeliveryStore
                if ($null -eq $ds) { "ACCOUNT|$name|$($a.AccountType)|||0|" }
                else {
                    $ok = 0; $dn = ''; $dh = ''
                    $m = $null
                    try { $m = [int]$ds.PropertyAccessor.GetProperty($mask) } catch { $m = $null }
                    if ($null -ne $m -and ($m -band 2) -ne 0) {
                        try { $ib = $ds.GetDefaultFolder(6); $dh = [string]$ib.PropertyAccessor.BinaryToString($ib.PropertyAccessor.GetProperty($draftsTag)) } catch { $dh = '' }
                    }
                    if (-not $dh) { try { $dh = [string]$ds.PropertyAccessor.BinaryToString($ds.PropertyAccessor.GetProperty($draftsTag)) } catch { $dh = '' } }
                    if ($dh) { try { $d = $ns.GetFolderFromID($dh, $ds.StoreID); $dn = $d.Name; $ok = 1 } catch { $dn = 'ERR ' + $_.Exception.Message } }
                    else { $dn = 'NOT DESIGNATED - new_draft makes it on first use'; $ok = 2 }
                    "ACCOUNT|$name|$($a.AccountType)|$($ds.DisplayName)|$($ds.FilePath)|$ok|$dn|$($ds.StoreID)"
                }
            }
            if ($folderHex) {
                try {
                    $df = $ns.GetFolderFromID($folderHex, $folderStoreHex)
                    $rootId = [string]$df.Store.GetRootFolder().EntryID
                    $chain = @()
                    $p = $null
                    try { $p = $df.Parent } catch { $p = $null }
                    for ($k = 0; $k -lt 32 -and $null -ne $p; $k++) {
                        $pe = ''
                        try { $pe = [string]$p.EntryID } catch { $pe = '' }
                        $chain += $pe
                        if (-not $pe -or $pe -ieq $rootId) { break }
                        try { $p = $p.Parent } catch { $p = $null }
                    }
                    "DFOLDER|$($df.Name)|$rootId|$($chain -join ',')"
                }
                catch { "DFOLDER-ERR|$($_.Exception.Message)" }
            }
            'END'
        }
        $deadline = (Get-Date).AddSeconds(180)
        while ((Get-Date) -lt $deadline -and $job.State -eq 'Running') { Start-Sleep -Seconds 2 }
        $lines = @(Receive-Job -Job $job 2>&1 | ForEach-Object { [string]$_ })
        if ($job.State -eq 'Running') { Stop-Job -Job $job; Remove-Job -Job $job -Force; throw "The COM read did not finish in 180 s. Last line: $($lines[-1])" }
        Remove-Job -Job $job -Force
        $accounts = @()
        $folderRefusal = $null
        if (-not $folderHex) { $folderRefusal = "the account has no '$FolderValue' at all" }
        foreach ($l in $lines) {
            $f = $l.Split('|')
            switch ($f[0]) {
                'PROFILE' { Say "COM: profile '$($f[1])'"; if ($f[1] -ne $ProfileName) { throw "Outlook is on '$($f[1])', not '$ProfileName'." } }
                'STORE' { Say "COM: store '$($f[1])'  $($f[2])" }
                'ACCOUNT' {
                    $acc = [pscustomobject]@{ Name = $f[1]; DeliveryPath = $(if ($f[4]) { $f[4] } else { $null }); DeliveryStoreId = $(if ($f.Count -gt 7) { $f[7] } else { $null }); DraftsOk = ($f[5] -eq '1'); DraftsAbsent = ($f[5] -eq '2') }
                    $accounts += $acc
                    $dsText = '<NULL>'; if ($f[4]) { $dsText = "'$($f[3])' ($($f[4]))" }
                    Say ("COM: account '{0}' type={1} DeliveryStore={2} Drafts={3}" -f $f[1], $f[2], $dsText, $(if ($f.Count -gt 6) { $f[6] } else { '' }))
                }
                'DFOLDER' {
                    $chain = @(); foreach ($e in @($f[3].Split(','))) { if ($e) { $chain += $e } else { $chain += $null } }
                    $visible = Test-FolderIsVisible -Name $f[1] -AncestorEntryIds $chain -RootEntryId $f[2]
                    $folderRefusal = Get-DeliveryFolderRefusal -EntryIdHex $folderHex -Name $f[1] -Visible $visible
                    Say ("COM: the identity account delivers into '{0}' (visible={1}, PST node id {2})" -f $f[1], $visible, $(if ($null -ne (Get-PstFolderNid (ConvertFrom-HexString $folderHex))) { '0x{0:X}' -f (Get-PstFolderNid (ConvertFrom-HexString $folderHex)) } else { 'n/a' }))
                }
                'DFOLDER-ERR' { $folderRefusal = "its '$FolderValue' would not open: $($f[1])" }
                'END' { }
                default { Say "COM: $l" }
            }
        }
        $distinct = @($accounts | Where-Object { $_.DeliveryStoreId } | ForEach-Object { $_.DeliveryStoreId } | Sort-Object -Unique).Count
        Say "COM: $($accounts.Count) account(s), $distinct distinct delivery store(s)"
        $problems = @(Get-IdentityVerdict -Accounts $accounts -IdentityName $AccountName -Path ([IO.Path]::GetFullPath($PstPath)) -DeliveryFolderRefusal $folderRefusal)
        if ($TrySmtpAddress) {
            $sj = Start-Job -ScriptBlock {
                $ol = New-Object -ComObject Outlook.Application
                $ns = $ol.GetNamespace('MAPI')
                for ($i = 1; $i -le $ns.Accounts.Count; $i++) { $a = $ns.Accounts.Item($i); "SMTP|$($a.DisplayName)|reading"; "SMTP|$($a.DisplayName)|$($a.SmtpAddress)" }
            }
            $d2 = (Get-Date).AddSeconds(45)
            while ((Get-Date) -lt $d2 -and $sj.State -eq 'Running') { Start-Sleep -Seconds 2 }
            $sl = @(Receive-Job -Job $sj 2>&1 | ForEach-Object { [string]$_ })
            foreach ($l in $sl) { $f = $l.Split('|'); if ($f[2] -ne 'reading') { Say "COM: account '$($f[1])' SmtpAddress='$($f[2])'" } }
            if ($sj.State -eq 'Running') {
                Say 'COM: Account.SmtpAddress BLOCKED - the Object Model Guard prompt (antivirus reported out of date). Not a binding fault; the registry Email above is what Outlook stores.'
                Stop-Job -Job $sj
            }
            Remove-Job -Job $sj -Force -ErrorAction SilentlyContinue
        }
        if ($problems.Count -gt 0) {
            foreach ($p in $problems) { Say "FAIL: the identity account - $p" }
            exit 1
        }
        Say "OK: '$AccountName' delivers to its own store '$StoreDisplayName' ($PstPath), into a named, visible folder that is not the PST's root; its Drafts is designated and opens, or is not designated yet; and no other account shares the store."
        exit 0
    }
}
