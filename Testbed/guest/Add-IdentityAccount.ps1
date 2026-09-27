#Requires -Version 5.1
<#
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
        REFUSES. That is the point: HOW identity.pst gets a real, designated Inbox is an open
        question, and its candidates are listed in Docs/live-tier-on-the-vm.md section 3b, "The
        identity store has no Inbox" - each to be measured on a guest before any is written in here.

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

    WHAT IT NEVER DOES: start, quit or kill Outlook; create, modify, move or delete an item; touch
    any profile other than -ProfileName, or any account other than the one whose Email is
    -EmailAddress. Every write is under HKCU\...\Office\<ver>\Outlook, plus the rendered .prf and
    the JSON capture under -WorkDir.

    THE GUARD: Assert-TestbedGuest from OutlookMapiInterop.ps1 (dot-sourced; stage it beside this
    script). It refuses anywhere not logged on as vmadmin - on the maintainer's workstation this
    would be operating on a real profile carrying real delegate mailboxes.

    Windows PowerShell 5.1 - no ternary, no `??`. COM phases run in SESSION 1
    (Register-InteractiveTask.ps1); Outlook cannot start in session 0.

.PARAMETER Phase
    Import, CaptureStore, Bind or Verify - see above. Without -Execute, Import, CaptureStore and
    Bind print what they would do and change nothing.

.PARAMETER SelfTest
    Pure: renders the template, and drives the account-selection, EntryID-sanity and verdict
    decisions with synthetic inputs. No registry, no COM, no files written. Runs anywhere.

.PARAMETER TrySmtpAddress
    With -Phase Verify: also read Account.SmtpAddress, last, in its own job with a deadline.

.EXAMPLE
    .\Add-IdentityAccount.ps1 -SelfTest
    .\Add-IdentityAccount.ps1 -Phase Import -Execute
    .\Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName identity@vm.invalid -Path C:\OutlookAI-Tier\identity.pst -Execute
    .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute
    .\Add-IdentityAccount.ps1 -Phase Bind -Execute
    .\New-TierProfile.ps1 -StoreSinkPassword -Execute
    .\Add-IdentityAccount.ps1 -Phase Verify -TrySmtpAddress
#>
[CmdletBinding()]
param(
    [ValidateSet('Import', 'CaptureStore', 'Bind', 'Verify')] [string] $Phase,
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
    [string] $PstPath          = 'C:\OutlookAI-Tier\identity.pst',
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

    Write-Host ''
    Write-Host ("SelfTest: {0} passed, {1} failed. The guest halves - the import, AddStoreEx, the COM reads and whether Outlook honours the bound values - are what -Phase proves, not this." -f $script:pass, $script:fail)
    if ($script:fail -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }
if (-not $Phase) { throw 'Pass -Phase Import|CaptureStore|Bind|Verify, or -SelfTest.' }

# =============================================================================================
# GUEST ONLY FROM HERE.
# =============================================================================================
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser

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
        Say 'NEXT: start Outlook once (it imports at start-up and deletes ImportPRF itself - measured), then with it running:'
        Say "  .\Add-OutlookPstStore.ps1 -ProfileName $ProfileName -DisplayName $StoreDisplayName -Path $PstPath -Execute"
        Say '  .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute'
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

            # The Inbox is PROVEN present by the store's own PR_VALID_FOLDER_MASK before it is asked
            # for - the rule OutlookAI.Core's SpecialFolders follows. Asked for without that proof,
            # GetDefaultFolder(6) on an AddStoreEx PST hands back its non-IPM ROOT (measured on both
            # guests, 2026-09-24), and that is what this phase used to record.
            $accessor = $hit.PropertyAccessor
            $mask = $null
            try { $mask = [int]$accessor.GetProperty($ValidFolderMaskSchema) } catch { $mask = $null }
            if ($null -eq $mask) { throw "The store's PR_VALID_FOLDER_MASK would not read, so nothing proves whether it has an Inbox. Refusing to capture." }
            $inboxHex = ''; $inboxName = ''; $inboxVisible = $false; $draftsHex = ''
            if (($mask -band $FolderIpmInboxValid) -ne 0) {
                $inbox = $hit.GetDefaultFolder(6)
                $inboxHex = [string]$inbox.EntryID
                $inboxName = [string]$inbox.Name
                $rootId = [string]$hit.GetRootFolder().EntryID
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

            $script:captured = [pscustomobject]@{
                FilePath = $hit.FilePath; DisplayName = $hit.DisplayName; StoreID = $hit.StoreID
                InboxEntryID = $inboxHex; InboxName = $inboxName; InboxVisible = $inboxVisible
                InboxDesignated = (($mask -band $FolderIpmInboxValid) -ne 0); DraftsEntryID = $draftsHex
                Profile = $ns.CurrentProfileName; CapturedAt = (Get-Date).ToString('o')
            }
        }
        $c = $script:captured
        Say ("store '{0}' at {1}: StoreID {2} bytes; Inbox designated={3}, EntryID {4} bytes, name '{5}', visible={6}; Drafts {7}" -f $c.DisplayName, $c.FilePath, ($c.StoreID.Length / 2), $c.InboxDesignated, ($c.InboxEntryID.Length / 2), $c.InboxName, $c.InboxVisible, $(if ($c.DraftsEntryID) { 'designated' } else { 'NOT designated' }))
        $refusal = Get-DeliveryFolderRefusal -EntryIdHex $c.InboxEntryID -Name $c.InboxName -Visible $c.InboxVisible
        if ($refusal) {
            throw ("REFUSING to capture '$($c.DisplayName)' as the identity account's delivery target: $refusal. " +
                'Binding it anyway is exactly the 2026-09-24 defect - POP3 mail filed where nobody can see it. The store needs a real, ' +
                'designated Inbox first, and HOW is an open decision: Docs/live-tier-on-the-vm.md section 3b, "The identity store has no Inbox". Nothing was written.')
        }
        if (-not (Test-StoreEntryIdNamesPath -EntryId (ConvertFrom-HexString $c.StoreID) -Path $c.FilePath)) { throw 'The StoreID does not carry the PST path; refusing to record it.' }
        if (-not $Execute) { Say "Dry run. Would write $IdsPath."; return }
        Set-Content -LiteralPath $IdsPath -Value ($c | ConvertTo-Json) -Encoding UTF8
        Say "wrote $IdsPath"
        Say 'NEXT: close Outlook - a graceful Application.Quit() in session 1 (cancel the identity logon dialog first), never taskkill it - then:  .\Add-IdentityAccount.ps1 -Phase Bind -Execute'
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
