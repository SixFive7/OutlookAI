# A free route to the tier profile's POP3 account — three routes examined

Researched 2026-09-15. **Nothing in this document was executed against Outlook, MAPI, a mail
profile, or the Windows Messaging Subsystem hive.** The research machine is the maintainer's own
workstation with real mail and delegate mailboxes; every script here is a draft for a Hyper-V
guest and carries the never-executed banner.

Evidence labels: **[MS-DOC]** Microsoft documentation · **[COMMUNITY]** third-party or forum ·
**[INFERRED]** reasoning, with the chain stated · **[REPO]** read out of this repository, with a
path and line · **[MEASURED]** run during this research on the workstation, touching nothing
Office-related.

**Answers measured later, on a guest, are added in place and marked [GUEST-MEASURED, date]**,
beside the text they answer, which is kept as written: this is a record of what was known on
2026-09-15, and the gap between that and what a guest later showed is part of the record. The
guest is Office LTSC 2024, 16.0.17932.

---

## 0. The verdict, before the detail

**Route B substantially dissolves the problem, and it should be read first.** The requirement was
never "a POP3 account with a working transport". It is two much smaller things wearing one name:

1. an `Account` object whose `SmtpAddress` matches the hub store's display name, and
2. for a minority of tests, an item that *looks* received sitting in the hub Inbox.

Requirement 1 does not care what transport the account has — **`Account.AccountType` is never read
anywhere in the product** [REPO]. Requirement 2 is already solved in this repository twice over,
by the corpus generator and by one live helper that made exactly this trade on purpose.

**Measured against the live tier's own test methods:**

| What the guest has | Reachable, of the 34 methods in the four "blocked" families |
| --- | --- |
| PST-only, **zero** accounts | **3 / 34** |
| PST-only + **any one** account of any type whose SMTP matches the hub | **26 / 34** |
| + a PST-direct seeding helper | **33 / 34** |
| full POP3-plus-sink | **34 / 34** |

The last row buys **one** test method that nothing else can cover:
`T3/Phase5LiveMcpToolShapeTests.SendTool_TwoStepFlow_RoundTrip_OverRealStdio_WithAuditLines`.

**Recommended order: B, then A, then C.** Reasoning in §4.

---

## Route B — question the requirement

**This is the highest-value finding in the investigation.**

### B.1 The runbook's "six" is a file count mistaken for a method count

`Docs/live-tier-on-the-vm.md:102` says *"Six live methods also need the mail to genuinely
arrive"*, duplicated verbatim at
`McpServer/OutlookAI.McpServer.Tests/T2/LiveMailSink.cs:45`. [REPO]

The six is the number of **files** that hold an arrival wait, not the number of methods.
`T2/LiveInboxArrival.cs:10-21` says the helper replaced *"five verbatim copies … A SIXTH copy was
missed by that consolidation"*. [REPO] Independently confirmed here: exactly six files call
`LiveInboxArrival.WaitFor` — `LiveDraftOptionsTests`, `LiveDraftTests`, `LiveHtmlDraftTests`,
`LiveMoveArchiveTests`, `LiveSignatureTests`, `LiveUpdateDiscardTests`. [MEASURED]

Three further arrival waits were never in that consolidation: `LiveFreshModeTests`' own loop,
`LiveSweepScopeTests`' own 240 s loop, and the two T3 stdio pollers.

**The real figure is 13 methods that put mail on the wire and then depend on its arrival:**

| # | Method | Send | Arrival wait |
| --- | --- | --- | --- |
| 1 | `T2/LiveDraftTests.cs:101` `DerivedDrafts_Hub_ThreadingQuotedHistoryAndPlacement` | `:110` | `:111` |
| 2 | `T2/LiveDraftOptionsTests.cs:176` `DerivedDrafts_CcBccAppend_SubjectOverrideKeepsThreading_ImportanceAndReceiptRoundTrip` | `:180` | `:182` |
| 3 | `T2/LiveDraftOptionsTests.cs:272` `ForwardDraft_CcBccAppend_AndSubjectOverrideKeepsTheForwardedContent` | `:277` | `:279` |
| 4 | `T2/LiveHtmlDraftTests.cs:145` `ReplyDraft_BodyHtml_LeavesTheQuotedOriginalIntactAndBelowTheBody` | `:149` | `:151` |
| 5 | `T2/LiveSignatureTests.cs:76` `ReplyDraft_WithSignatureOverride_InsertsAboveQuote_ThenReplaceBranchReapplies` | `:85` | `:86` |
| 6 | `T2/LiveUpdateDiscardTests.cs:221` `UpdateDraft_OnAReply_ReplacesTheBody_AndKeepsTheQuotedOriginal` | `:228` | `:231` |
| 7 | `T2/LiveUpdateDiscardTests.cs:390` `DiscardDraft_RefusesASentItem_EvenWhenTheRegistryGateIsSatisfied` | `:396` | `:400` |
| 8 | `T2/LiveMoveArchiveTests.cs:66` `MoveChain_TestFolderRoundTrip_Archive_Guards_Audit_Cleanup` | `:78` | `:80` |
| 9 | `T2/LiveSweepScopeTests.cs:90` `ControlledCorpus_CrossColumnTermsMatch_AndTheSweepFollowsTheSearchScope` | `:105` | own loop `:361` |
| 10 | `T2/LiveFreshModeTests.cs:34` `FreshSearch_FindsSelfSentMail_BeforeIndexCatchesUp_ThenCleansUp` | `:45` | own loop `:60`, 120 s |
| 11 | `T3/MoveArchiveLiveMcpToolTests.cs:41` `MoveAndArchive_GoldenShapes_OverRealStdio` | `:49` | stdio poll `:55-81` |
| 12 | `T3/Phase4LiveMcpToolShapeTests.cs:43` `DraftTools_GoldenShapes_OverRealStdio_WithAuditLines` | `:51` | stdio poll `:60-87` |
| 13 | `T3/Phase5LiveMcpToolShapeTests.cs:50` `SendTool_TwoStepFlow_RoundTrip_OverRealStdio_WithAuditLines` | product `send` `:101` | `:121`, Sent Items `:139` |

Numbers 1–12 seed via `LiveOutlookTestMailer.SendSelfMail`. **Number 13 is the only one that sends
through the product.** [REPO]

### B.2 The `Requires=Transport` trait over-declares by 12

25 methods carry `[Trait("Requires", "Transport")]` [MEASURED — 25 occurrences across 12 files].
The vocabulary defines it as *"mail that actually goes out and comes back"*
(`Docs/live-tier-on-the-vm.md` §5; `T1/LiveTierInventoryTests.cs:102-103`). [REPO]

**12 of the 25 never put mail on the wire at all:** `T2/LiveAttachmentKindRecallTests.cs:321`
(seeds via `SaveTaggedDraftWithAttachments`, a PST write), `T2/LiveDraftTests.cs:45`, `:179`,
`:241`, `:316`, `T2/LiveHtmlDraftTests.cs:329`, `T2/LiveUpdateDiscardTests.cs:145`, `:315`,
`:358`, `:427`, `:574`, `:621`. [REPO]

Corroborating this from the other direction: `T2/LiveSendTests.cs` declares
`Requires=MailAccount` on all four methods and `Requires=Transport` on **none**, and its class
comment says why — *"all WITHOUT any transport (every path here refuses BEFORE `Send()`)"*.
[MEASURED + REPO]

### B.3 What `NewDraft` actually does with the resolved `Account`

`OutlookComSession.TryCreateNewDraft`, `McpServer/OutlookAI.Core/Com/OutlookComSession.cs:3470`:

| Step | Line | `Account` member touched |
| --- | --- | --- |
| `FindAccountBySmtp(accountSmtpAddress)` | `:3517` (impl `:7496-7526`) | **`SmtpAddress` only**, `OrdinalIgnoreCase`, walking `NameSpace.Session.Accounts` by index |
| null → `AccountNotFound` | `:3520` | — |
| outcome snapshot | `:3526` | `SmtpAddress` |
| `account.DeliveryStore` | `:3531`, null → `AccountHasNoDeliveryStore` `:3538` | `DeliveryStore` |
| store identity | `:3542-3543` | `DeliveryStore.DisplayName`, `.StoreID` |
| `GetDefaultFolder(16)` | `:3545` | the delivery store's Drafts folder |
| `SetSendUsingAccount(mail, account)` | `:3563` (impl `:6134`) | the account **object**, as a putref argument |

**That is the complete list.** `Account.AccountType` is never read anywhere in the product — a
repo-wide search for `AccountType|olExchange|olPop3|olImap|olOtherAccount` across `McpServer/`,
`Services/` and the root `.cs` files returns one hit, `MailService.cs:7434`, and that is
`OlExchangeStoreType`, a **store** property used for delegate detection. **Exchange vs POP3 vs
IMAP is structurally invisible to this code path.** [REPO]

A live-verified footgun worth carrying forward: `OutlookComSession.cs:6126-6133` records that
*"`MailItem.SendUsingAccount` is a PROPERTYPUTREF property. A late-bound dynamic assignment
SILENTLY NO-OPS on this Outlook build."* [REPO]

**The minimum that satisfies `NewDraft`:** any `Account` in the profile whose `SmtpAddress`
matches case-insensitively and whose `DeliveryStore` has a Drafts default folder. The transport
underneath is never consulted, probed or asserted. **An account that already exists for a
different transport would do, unchanged.**

**Where a zero-account PST-only profile fails, exactly:**

* `new_draft` → `OutlookComSession.cs:3519-3522`, `AccountNotFound`, nothing created.
* `send` → `MailService.cs:5996-6001`, `no_sending_account`, in **step 1**, before a confirm
  token is issued — so even the transport-free negative send tests need an account.
* **Derived drafts do *not* fail.** `TryCreateDerivedDraft:3714` resolves identity by
  `FindAccountByDeliveryStore:3778` and, when nothing matches, records `accountResolved = false`
  and leaves `SendUsingAccount` for Outlook (`:3771-3773`). The draft is still created and saved.
  What fails is the **test's** `Assert.True(outcome.AccountResolved, …)` at
  `T2/LiveDraftTests.cs:58` and `:358`. [REPO]

### B.4 PST-direct seeding is already shipped, twice

**(i) The corpus generator builds 20,000 items that read as received mail, with zero transport.**
`McpServer/OutlookAI.RemediationTools/ComCorpusMailbox.cs`: `Items.Add` + `Save`, then
`ApplyMessageFlags:478` does a read-modify-write of `PR_MESSAGE_FLAGS`
(`…/mapi/proptag/0x0E070003`, `:36`), clearing `MSGFLAG_UNSENT 0x8` (`:43`) and always clearing
`MSGFLAG_SUBMIT 0x4` (`:41`, *"the bit that MEANS 'queued for delivery'"*), then `Move()` to the
target folder (`:346-350`, `:684-689`). Dates are written as `PR_MESSAGE_DELIVERY_TIME`
`0x0E060040` — *"what `MailItem.ReceivedTime` and `urn:schemas:httpmail:datereceived` read"*
(`:29-30`) — and `PR_CLIENT_SUBMIT_TIME` `0x00390040` (`:32-33`). [REPO]

`CorpusPlacement.cs:91-104` records why the ladder exists: a first build of 40,000 items with
`Items.Add`+`Save` straight into Inbox/Sent/Deleted/Junk left **all 40,000 in Drafts**, because
such an item is `MSGFLAG_UNSENT`. Clearing the flag is what makes it live where you put it. That
problem is solved, verified on this corpus, and shipped. [REPO]

**(ii) The live suite already made this exact trade once, deliberately.**
`T2/LiveOutlookTestMailer.cs:736` `SaveTaggedDraftWithAttachments`, comment at `:731-735`:
*"Deliberately a draft rather than a self-send: … transport adds a dependency on Outlook actually
flushing the Outbox (which a headless instance may not do — soak fix 15). **Drafts are indexed
exactly like received mail.**"* It already backs `T2/LiveAttachmentKindRecallTests.cs:321` and
`T2/LiveUpdateDiscardTests.cs:358` — **both of which still carry a `Requires=Transport` trait they
do not use.** [REPO]

**What the 12 `SendSelfMail` seeds actually need from the arrival:** an item in the hub **Inbox**
(`T2/LiveInboxArrival.cs:78-79` filters on `FolderKind == "inbox"` and an ordinal subject match),
with a known subject, carrying a `ConversationIndex`. `MoveChain` additionally asserts
`item1.FromFolder == "Inbox"` (`T2/LiveMoveArchiveTests.cs:97`). The corpus generator's
`DraftsThenMoveWithSentFlag` rung produces exactly that shape. [REPO]

**Two things break, and both must be named:**

1. **`ConversationIndex` is the one unverified assumption. [INFERRED]** Nothing in this repo
   records whether an `Items.Add`+`Save` item receives a store-generated
   `PR_CONVERSATION_INDEX`; `ComCorpusMailbox` never writes one, and `ConversationIndex` appears
   nowhere in `Docs/`, `TODO.md` or `QUESTIONS.md`. Outlook is believed to generate one on save,
   so `Reply()` would still build a child index — but no measurement in this repo proves it, and
   **five assertions rest on it**: `T2/LiveDraftTests.cs:365-368`,
   `T2/LiveDraftOptionsTests.cs:228`, `:246`, `:305`, `T2/LiveSignatureTests.cs:102`.
   **This is a five-minute probe and it must be run before this route is chosen.**
2. **Sender identity is lost.** `T3/Phase5LiveMcpToolShapeTests.cs:122-124` asserts the received
   copy reports the hub as sender. A locally created item has no `SenderEmailAddress`; on a real
   send it comes from the transport. [REPO + INFERRED]

**Cost:** a new mutation helper in `LiveOutlookTestMailer`, a matching `StoreWriteKind` in
`StoreWriteAllowlist` / `LiveStoreWriteGuard`, plus tests. Per `CLAUDE.md` mailbox-safety rule 1
this cannot be improvised in shell code — it must go into the tested helpers.

### B.5 What would be lost

**One capability is genuinely irreplaceable: the product's own `send` tool reaching a transport.**
`T3/Phase5LiveMcpToolShapeTests.cs:50` is the only test in all 127 live methods that calls the
product's `send` with a valid confirm token and lets Outlook submit. Skipping it loses:

* `status == "sent"`, `sent == true`, **`accountVerified == true`** (`:112-114`) — i.e. the
  `SendUsingAccount` putref-then-getter-readback verification at
  `OutlookComSession.cs:4928-4951` executing against a real send. That code aborts with
  `SendIdentityVerificationFailed` on mismatch and is the guard against sending from the wrong
  account.
* Arrival plus From-identity on the received copy (`:121-124`).
* The negative: `send` on an already-arrived copy refuses `not_an_unsent_draft` (`:133-135`).
* Sent Items filing latency (`:139`).
* Audit ordering `send_token_issued` before `send`, plus both refusal audit lines (`:143-149`).

**Covered elsewhere:** token binding, single-use, expiry, content-hash invalidation and
deleted-draft fail-closed are all in `T2/LiveSendTests.cs` (4 methods) and
`T2/LiveUpdateDiscardTests.cs:574`, none of which declares `Requires=Transport`. Decision logic is
covered in CI by `T1/SendValidationTests.cs`, `T1/SendConfirmationTests.cs` and
`T1/AtomicityClaimsTests.cs:259`. **Not covered anywhere else:** the `accountVerified` readback
and `MailItem.Send()` itself — only the maintainer's production machine would still exercise
them. **That is the honest cost, and it is one test method.** [REPO]

**Secondary losses, recoverable but weaker:**

* `T2/LiveFreshModeTests.cs:52-122` measures raw sweep arrival latency against a real delivery
  instant. Under PST-seeding the "fresh beats the index" claim becomes a claim about item
  *creation* time, not delivery time. Still a valid test of the sweep's gap window; weaker as a
  claim about mail.
* **The Outbox stops being a canary.** `T2/LiveMailSink.EnsureOutboxDrained:166` and the
  zero-artifact sweep over folder 4 (`LiveOutlookTestMailer.cs:58`) exist to catch a genuine
  send-path leak. With no transport the Outbox can never fill from a seed — and can never prove a
  leak either. **That guard goes vacuous on the VM**, and the runbook should say so rather than
  leave a guard that reads as protective.

**Removal mechanics are free and already tested.** Dropping the `mailSink` block from the settings
file is a supported state requiring no code change: `LiveMailSink.EnsureReachable` returns at
`:102-107` when `settings.MailSink == null`, `EnsureOutboxDrained` at `:170-173`, and
`T2/CorpusFreshnessTests.cs:712 LiveSettings_AbsentMailSinkMeansRealTransport` pins it.
[REPO + MEASURED — both early-return paths read directly]

**But the semantics are inverted and must be rewritten, not silently reused.** Absent currently
means *"this machine has real transport"* (`Testbed/live-test-settings.example.json`
`mailSink._note`; `Docs/live-tier-on-the-vm.md:477`). Left as is, a machine with **no** transport
would read as a machine with **perfect** transport.

### B.6 Grading

Families named at `Docs/live-tier-on-the-vm.md:96-98`, counted from source, not from the stale
inventory: draft creation 11 (`LiveDraftTests` 5 + `LiveDraftOptionsTests` 5 + `Phase4` 1),
update/discard 12, HTML draft 6, send 5 (`LiveSendTests` 4 + `Phase5` 1) — **34 total**.
[REPO; the five per-file `[Fact]` counts independently confirmed as 5/5/12/6/4 — MEASURED]

| Scenario | Reachable | Which, and why |
| --- | --- | --- |
| **(a) PST-only, zero accounts** | **3 / 34** | `T2/LiveDraftTests.cs:316` and `T2/LiveSendTests.cs:182` (both `ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain`, which only count and delete tagged artifacts), and `T2/LiveUpdateDiscardTests.cs:358` (seeds via a PST write; `DiscardDraft` refuses at `MailService.cs:5632` "before any COM work"). The other 31 hit `AccountNotFound`, `no_sending_account`, or an `AccountResolved` assertion. |
| **(b) PST-only + any one account of any type whose SMTP matches the hub** | **26 / 34** | The 8 arrival-dependent methods in these families still fail: `LiveDraftTests:101`; `LiveDraftOptionsTests:176`, `:272`; `LiveUpdateDiscardTests:221`, `:390`; `LiveHtmlDraftTests:145`; `Phase4:43`; `Phase5:50`. **Caveat:** with exactly one account the two `IdentityAccount` tests (`LiveDraftTests:241`, `LiveDraftOptionsTests:111`) run but iterate nothing and announce that they proved nothing — so **26 reachable, 24 asserting**. |
| **(c) + PST-direct seeding helper** | **33 / 34** | 7 of the 8 are seedable. |
| **(d) full POP3-plus-sink** | **34 / 34** | Same `IdentityAccount` caveat unless a second account exists. |

**The delta from (b) to (d) is 8 methods out of 34; seven are seedable without transport. One
account of any type buys 26 of 34.**

### B.7 Repo defects found along the way

* **`Docs/live-test-inventory.txt` is stale and cannot be reconciled — regenerate it.** It prints
  a `[Portable]/[ProfileBound]` column for the `LiveTier` trait, which is on `RetiredTraits`
  (`T1/LiveTierInventoryTests.cs:155`) and refused by `NoTestDeclaresARetiredTrait`; and it prints
  `req=` as a **per-class union**, the exact shape `EveryLiveTestMethod_NamesItsOwnCapabilities`
  now forbids. Consequences in the checked-in file: all 5 `LiveAttachmentKindRecallTests` methods
  show `Transport` when only `:321` declares it; all 5 `LiveDraftOptionsTests` show it when 2 do;
  all 12 `LiveUpdateDiscardTests` show it when 8 do. [REPO]
* **`Docs/vm-coverage-analysis.md:36` carries the same stale total, "`Transport` 41".** The true
  per-method figure is 25. [REPO]
* **`Testbed/README.md` open question 11** already flags that smtp4dev v3 is usually described as
  SMTP-plus-**IMAP**, and that if its POP3 side does not exist the sink section is wrong
  regardless of how the account is created. That is a second, independent risk sitting in front of
  every route that keeps the requirement. [REPO]
* **`Docs/live-tier-on-the-vm.md:212-217`** — whether Outlook accepts `@` in a store display name
  is *"untested and it gates the whole draft family"*, because `testHubStoreDisplayName` doubles
  as an SMTP address and goes straight to `FindAccountBySmtp`. This gates (b) as much as (d).
  [REPO]

---

## Route C — UIAutomation

### C.1 Verdict

**Alive, but only on one of three surfaces, and not the one first suggested.** The modern
"simplified" wizard is a bad target; **the classic Win32 property-sheet wizard reached through the
Control Panel Mail applet is a good one**. No public UIA tree dump of either wizard exists, so the
go/no-go needs one two-minute dump on the guest — `Testbed/guest/Dump-UiaTree.ps1`, written for
exactly that.

### C.2 Does a UIA tree exist?

**The classic wizard — yes, with high confidence, by this chain:**

1. UIA's core sits between provider and client, and *"Microsoft Active Accessibility servers can
   provide information to UI Automation client applications."* Any window with an `IAccessible` —
   which every standard Win32 dialog gets free from `oleacc`/comctl32 — appears in the UIA tree.
   [MS-DOC] <https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-msaa>
2. The role→control-type table on that page maps `ROLE_SYSTEM_PUSHBUTTON`→Button,
   `ROLE_SYSTEM_TEXT`→Edit, `ROLE_SYSTEM_CHECKBUTTON`→CheckBox, `ROLE_SYSTEM_PAGETAB`→TabItem — a
   property-sheet wizard is exactly those roles. [MS-DOC]
3. `AutomationId` is unsupported only for *"UI Automation elements derived from Win32 controls
   that do not have a control ID"*; the negative phrasing presupposes the positive case. [MS-DOC]
   <https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/use-the-automationid-property>
4. Confirmed empirically elsewhere: the common file dialog's filename box is
   `AutomationId = "1148"` and its Open button `AutomationId = "1"` — the raw `IDOK`/resource
   control IDs as decimal strings. [COMMUNITY]
   <https://gist.github.com/natritmeyer/966981>
5. Microsoft publishes step-by-step **screen-reader** instructions for Outlook desktop account
   setup, naming announced controls verbatim. Narrator is a pure UIA client. [MS-DOC]
   <https://support.microsoft.com/en-us/office/use-a-screen-reader-to-set-up-your-email-account-in-outlook-0f620d5b-85a3-4728-8ee2-6f3aec900a86>

**[INFERRED]** (3)+(4) ⇒ classic pages built from standard dialog controls expose numeric,
locale-invariant `AutomationId`s. (1)+(2)+(5) ⇒ the tree exists and is navigable. The classic
Outlook pages **specifically** were not verified; that is the experiment in C.7.

**The modern simplified wizard — visible, but a poor target.** It *is* announced by screen readers
(same MS-DOC as (5)), so "UIA cannot see it" is **false**. But Office chrome is `NUIDialog`
hosting `NetUIHWND`/DirectUI [COMMUNITY]; UIA's `FrameworkId` does list `"DirectUI"` as a value
[MS-DOC <https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids>];
and NVDA's maintainers concluded *"it is probably best to disable UI Automation for the MS Office
ribbons entirely"* over *"deficiencies with roles and naming"* [COMMUNITY
<https://github.com/nvaccess/nvda/issues/4207>]. **[INFERRED]** DirectUI ⇒ no `AutomationId` ⇒ you
key on localised `Name`, on an en-GB guest, in a wizard that fires AutoDiscover on an isolated
network. Three compounding hazards. Do not target it.

### C.3 The registry knob

| | |
| --- | --- |
| User key | `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\setup` |
| Policy key | `HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\setup` |
| Value name | `DisableOffice365SimplifiedAccountCreation` |
| Type | `REG_DWORD` |
| Data | `1` |

[MS-DOC] KB3189194 —
<https://support.microsoft.com/en-us/topic/how-to-disable-simplified-account-creation-in-outlook-2016-outlook-2019-and-outlook-for-office-365-662bf4f8-c357-dbc8-53b3-ff8f445e8247>.
Applies to Outlook 2016, 2019 and Microsoft 365; the wizard it suppresses arrived in Click-to-Run
`16.0.6769.2015`. **No deprecation statement exists in Microsoft documentation.**

**Two risks, both [COMMUNITY] and both unresolved:**

* A 2025-03-07 Microsoft Q&A reports *"in Office 2024, this no longer works"* and that *"the online
  documentation does not mention Office 2024"*. Microsoft support asked for diagnostics and never
  answered.
  <https://learn.microsoft.com/en-us/answers/questions/4748152/office-2024-advanced-account-creation-(disableoffi>
* A 2023 thread reports it working on one machine and not another of the same vintage, and that
  *"the Mail App in Control Panel does not communicate with Outlook when the Simplified Account
  Creation interface is active."*

**This matters here specifically.** `Testbed/MEDIA.md` records the guests as
`ProPlus2024Volume` / `PerpetualVL2024`, measured at `16.0.17932.20884` — i.e. **Office 2024, the
exact configuration the unresolved report names.** [REPO] **Treat the knob as unproven on this
guest until a checkpoint says otherwise.**

**AutoDiscover suppressors** — `HKCU\Software\Microsoft\Office\16.0\Outlook\AutoDiscover` and the
parallel `…\Policies\…` key, all `REG_DWORD` [MS-DOC
<https://learn.microsoft.com/en-AU/outlook/troubleshoot/profiles-and-accounts/unexpected-autodiscover-behavior>]:
`PreferLocalXML`, `ExcludeHttpRedirect`, `ExcludeHttpsAutoDiscoverDomain`,
`ExcludeHttpsRootDomain`, `ExcludeScpLookup`, `ExcludeSrvRecord`, `ExcludeLastKnownGoodURL`,
`ExcludeExplicitO365Endpoint`. The same page warns that `ExcludeSrvLookup` *"doesn't exist in
Outlook code"* — a widely-copied myth; only `ExcludeSrvRecord` is read. **No `DisableAutoDiscover`
and no `EnableConservativeTracing` exist in any Microsoft source** — treat both as non-existent.

### C.4 Reachability from PowerShell 5.1

```powershell
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
```

Two calls, by simple name, no `-Path`, no `-ReferencedAssemblies`; resolved from the GAC to
4.0.0.0. `WindowsBase` loads transitively — do not add it. `AutomationElement`, `TreeWalker`,
`InvokePattern`, `ValuePattern`, `WindowPattern`, `TogglePattern`, `SelectionItemPattern` and
`ExpandCollapsePattern` all resolve. [MEASURED on the workstation — assembly loading only; no
Office process was touched]

**Landmine 1 — `LegacyIAccessiblePattern` does not exist in the managed API.** [MEASURED]
`System.Windows.Automation.LegacyIAccessiblePattern` fails to resolve; `UIAutomationClient.dll`
exports no Legacy anything, and `AutomationElement` has no `LegacyIAccessible.ChildId`.
`IUIAutomationLegacyIAccessiblePattern` exists only in the **COM** client API [MS-DOC]. **If a
control turns out to be LegacyIAccessible-only — the realistic risk on the modern DirectUI wizard
— a managed-PowerShell script cannot touch it at all.**

**Landmine 2 — the COM fallback is not reachable from PowerShell.** [MEASURED]
`[Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'ff48dba4-…'))` succeeds, but the
resulting object surfaces only `System.Object`/`MarshalByRefObject` members — zero `IUIAutomation`
members, because `IUIAutomation` is not IDispatch-derived and PowerShell's late binder cannot see
its vtable. Both `CUIAutomation` and `CUIAutomation8` are registered; registration is not the
problem. **⇒ managed wrapper, or an `Add-Type` C# interop shim. There is no third option.**

**Bitness — no constraint.** *"the … run-time components automatically handle all of the issues and
complexities involved in performing interprocess communications, including the interoperability
issues involved when one process is 32-bits and the other is 64-bits."* [MS-DOC]
<https://learn.microsoft.com/en-us/windows/win32/winauto/32-bit-and-64-bit-interoperability>.
The only stated exception is in-context WinEvent hooks, which a UIA client does not use.
**64-bit PowerShell reads 32-bit Outlook's tree fine.** Bitness *does* matter for launching the
CPL (C.5).

**Apartment state — STA is the default and is correct.** *"In Windows PowerShell 3.0, single-threaded
apartment (STA) is the default."* [MS-DOC]. UIA's MTA guidance is narrower than it looks — it is
about **event handlers**: *"A UI Automation client should use the COM MTA threading model for
threads that implement event handlers."* **[INFERRED]** a poll-and-assert script that never calls
`AddAutomationEventHandler` and never targets its own process is unaffected. **Run STA.**

**Scheduled task in session 1 — yes, and it is the right call.** `Register-InteractiveTask.ps1`
already registers `-LogonType Interactive -RunLevel Highest`. Microsoft's winappCli docs split the
verbs cleanly: `invoke`, `set-value`, `get-property`, `wait-for` are *"headless/locked-session
friendly"*, while `click`, `hover`, `send-keys --via send-input` *"synthesize OS-level input, so
they need an unlocked, interactive desktop"* and on a locked workstation *"fail fast with
`no_interactive_desktop`"*. [MS-DOC]
<https://github.com/microsoft/winappCli/blob/main/docs/ui-automation.md>
**This is a second, independent reason never to use SendKeys.**

**Performance:** *"you should try to obtain only direct children of the RootElement. A search for
descendants may iterate through hundreds or even thousands of elements, possibly resulting in a
stack overflow."* [MS-DOC] ⇒ `RootElement.FindFirst(TreeScope.Children, …)` for the dialog, then
`Descendants` **scoped to that dialog**. Use `ControlViewWalker`, not `RawViewWalker`; use
**current**, not cached, property requests — stale reads are exactly what breaks a
verify-the-state-changed loop.

### C.5 Which surface

| Surface | Verdict |
| --- | --- |
| `File > Add Account` in running Outlook | ✗ Needs Outlook fully started, a profile to already exist, and navigation of the **DirectUI backstage** to reach the button. Worst of both worlds. |
| `outlook.exe` first-run on a profile-less machine | ✗ One shot per checkpoint, no retry, and on modern builds it lands in the simplified wizard. Fallback only. |
| **Control Panel Mail applet** | **✓ Recommended.** No Outlook process, no mail store opened, plain Win32 property sheets, re-runnable. |

The repo already leans this way: `Docs/live-tier-on-the-vm.md:195-196` says how profiles are
created *"is **not recorded**; the Mail control panel works and is the obvious route."* [REPO]

**Launch, bitness-sensitive** [COMMUNITY
<https://www.slipstick.com/how-to-outlook/where-is-the-mail-icon/>]:

```
C:\Windows\System32\control.exe  "C:\Program Files\Microsoft Office\root\Office16\MLCFG32.CPL"        # 64-bit Office
C:\Windows\SysWOW64\control.exe  "C:\Program Files (x86)\Microsoft Office\root\Office16\MLCFG32.CPL"  # 32-bit Office
```

**[INFERRED]** `control.exe` hosts the CPL in a `rundll32.exe` of its own bitness, and a 32-bit CPL
cannot load into 64-bit `rundll32` — which is why the Control Panel item reads "Mail (Microsoft
Outlook) (32-bit)" on a 32-bit install. **Get this wrong and you get a silent no-window failure.**
`Testbed/MEDIA.md` records the guest Office as 64-bit, so the System32 form is the one to use.

`outlook.exe /profiles` [MS-DOC] and `/manageprofiles` [COMMUNITY] both reach Show Profiles, but
both start `outlook.exe`. The CPL does not. Prefer the CPL.

**Page-by-page sequence.** `Name` values are en-GB and shown for diagnostics only; window classes
are [INFERRED] and must be confirmed by the dump.

| # | Page | Class | Action |
| --- | --- | --- | --- |
| 0 | Mail Setup – Outlook | `#32770` | Invoke **Show Profiles…** |
| 1 | Mail | `#32770` | Invoke **Add…** |
| 2 | New Profile | `#32770` | `ValuePattern.SetValue` the name Edit → **OK** |
| 3 | Add Account — Auto Account Setup | `#32770` | `SelectionItemPattern.Select()` the **Manual setup** radio → **Next >** |
| 4 | Choose Service | `#32770` | Select **POP or IMAP** → **Next >** |
| 5 | POP and IMAP Account Settings | `#32770` | SetValue name, address, incoming `127.0.0.1`, outgoing `127.0.0.1`, user name; password blank; **clear** *Remember password*; **clear** *Automatically test account settings when Next is clicked* |
| 6 | More Settings → Advanced / Outgoing Server | `#32770` | POP port `110`, SMTP port `25`, no SSL; SMTP auth **off** → **OK** |
| 7 | back on page 5 | | **Next >** |
| 8 | Test Account Settings | `#32770` | **Should never appear.** If it does, hard-fail: with no stored password it prompts, and it sends a test message into the sink. |
| 9 | Congratulations / You're all set | `#32770` | **Finish** |

**Verification that touches no mail store** [INFERRED]: back on the Mail dialog, invoke **E-mail
Accounts…** and assert the account list contains a row matching the address with a POP/SMTP type.
Pure UIA re-read — no Outlook launch, no `.pst` opened, no COM. A read-only registry confirmation
under `HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\<name>`
is available as a second signal [MS-DOC, PRF section 6] — read only; never write it.

### C.6 Localisation — what to key on

The guest's **display language** is the only thing that matters; nl-NL *formats* affect date and
number rendering, not accessible names.

**Key on:** `AutomationId` — *"the `AutomationId` of an element must be the same in any instance of
the application, regardless of the local language"*, and *"For test automation, the clients should
consider using the `AutomationId` or `RuntimeId` property"* [MS-DOC]. On a classic dialog this is
the numeric control ID as a decimal string. **Primary key.** `ClassName` — *"as assigned by the
control developer … can be used to verify that an application is working with the expected
automation element"* [MS-DOC]; `#32770`, `Edit`, `Button`, `ComboBox`, `SysListView32` are Win32
constants and never localised. **Good secondary filter.** `ControlType` — enum, locale-invariant,
good filter but never a unique key. `ProcessId` / `NativeWindowHandle` — integers, use to scope.

**Traps:** `Name` — *"should be the same as the label text on screen"* and must be *"localized to
the application UI language"* [MS-DOC]. **Never a key.** `LocalizedControlType` — localised by
definition. `LegacyIAccessible.ChildId` — unreachable from the managed API entirely [MEASURED].
Child ordinal position — brittle across builds *and* across the conditional pages of C.7; tie-break
only. **`AutomationId` across builds** — [MS-DOC], verbatim: *"`AutomationId` is not guaranteed to
be stable across different releases or builds of an application."* Mitigation: the guest is
checkpointed — capture the dump once, pin the Office build, re-dump after any Office update.
`AutomationId` uniqueness — *"unique among sibling elements, but not necessarily unique across the
entire desktop"*; always search scoped to the dialog.

**The honest answer:** classic Win32 property-sheet controls generally expose the raw control ID as
a numeric `AutomationId`, and it is reliable *for a pinned build* — but that is inferred from
Microsoft's negative phrasing plus a confirmed file-dialog example, **not** from an observation of
the Outlook wizard. **If the dump returns `AutomationId=""` on every control, Route C collapses to
`Name`-keying on an en-GB guest, which should not be shipped.**

### C.7 How it fails, and forcing loud failure

**Silent no-op — the enemy:**

* `InvokePattern.Invoke()` on a **disabled** button returns `S_OK` and does nothing.
* `FindFirst` returns `null`, PowerShell assigns `$null`, and the failure surfaces somewhere
  unrelated.
* **The big one [INFERRED]:** `ValuePattern.SetValue` on a Win32 edit maps to `WM_SETTEXT`.
  Whether that reliably raises `EN_CHANGE` to the property sheet — which is what enables
  **Next >** — is not documented. If it does not, you get fields that look correct on screen with
  a permanently greyed Next. **This is the single most likely way Route C dies at implementation
  time.** Detect it by asserting `Next.Current.IsEnabled -eq $true` before invoking.

**Loud-failure design, non-negotiable:** assert the expected window exists before every step,
matched on `ProcessId` + `ClassName` + a known `AutomationId`, and throw on absence; **never
`SendKeys`**; never sleep-and-hope — bounded waits with an explicit deadline and a `throw` on
expiry; **verify by re-reading the tree**, never by trusting a return value; enumerate unexpected
`#32770` windows in the target process, dump them and throw; cap any toggle loop at 3 iterations
so a control that ignores you cannot spin forever.

**Pages that appear only sometimes:** an "Add Account" welcome page on some builds; "Change Account
Settings"; the Test Account Settings modal (suppressed by clearing the checkbox on page 5 — do
this, it is also what stops a test message reaching the sink); encrypted-connection retry prompts;
a Congratulations page whose text and button label vary by build.

### C.8 What could not be established without a machine

1. **Whether the classic Add Account pages expose non-empty numeric `AutomationId`s.** *The* open
   question; everything in C.5–C.6 rests on it.
2. Whether the classic pages are `#32770` or an Office-custom class hosting `NetUIHWND`.
3. Whether `ValuePattern.SetValue` enables the Next button.
4. Whether `DisableOffice365SimplifiedAccountCreation` works on the guest's Office 2024 build.
5. Whether the Mail applet honours the knob, or opens the simplified wizard regardless.

**One experiment settles 1, 2, 3 and 5 at once, in two minutes:** set the knob, launch the Mail
applet with the correct-bitness `control.exe`, click by hand to the POP and IMAP Account Settings
page, and run `Testbed/guest/Dump-UiaTree.ps1`. **Decision rule: if `AutomationId` is a non-empty
number and `FrameworkId` is `Win32`, Route C is viable and the dump *is* the spec. If
`AutomationId` is empty or `FrameworkId` is `DirectUI`, Route C is dead for a PowerShell 5.1
managed client.**

---

## Route A — the `.prf` spike

### A.1 Verdict, and a claim that needed correcting

**Not dead — and the widely-repeated "PRF doesn't work on Outlook 2016" is scoped to Exchange
accounts only.** Traced to origin:

* *"The concept of Reliable Profiles was introduced in Microsoft Outlook 2016. As a result, **we
  don't recommend** using Outlook profile (`.prf`) files to create profiles in Outlook 2016."* — a
  **recommendation**, on a page entirely about Exchange/`ZeroConfigExchange`. [MS-DOC]
  <https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/profiles-and-accounts/zeroconfigexchange>
* The Outlook-team blog behind that line enumerates exactly **two** unsupported scenarios and
  **both are Exchange**: Exchange Online, because *"the .PRF file requires an Exchange server name
  to be hard coded … the server name contains a variable GUID"*; and Outlook 2016, because
  *"changes to how Exchange Autodiscover information is stored … prevents the use of .PRF files
  **to configure Exchange accounts** in Outlook 2016."* [COMMUNITY — verbatim mirror of the MS
  Outlook team blog]
  <https://thewindowsupdate.com/2019/04/08/zeroconfigexchange-automating-the-creation-of-an-outlook-profile-for-exchange-accounts/>
* A Microsoft support engineer repeats the same scoping verbatim. [COMMUNITY — Microsoft employee,
  archived on learn.microsoft.com]
  <https://learn.microsoft.com/en-us/archive/msdn-technet-forums/db5a49ad-c258-42bc-a2f6-02bb2aed425c>
* POP3 remains a first-class OCT 2016 account type, and the OCT is the PRF generator. [MS-DOC]
  <https://learn.microsoft.com/en-us/office/customization-tool/oct-2016-help-add-account-and-account-settings-dialog-box>

**The counter-evidence, and it is the single biggest risk to this route:** the OCT 2016 "Apply
PRF" page says *"If you created a PRF file for a previous version of Outlook, you can import it to
Outlook, **provided that the profile defines only MAPI services**."* POP3 accounts live in
sections 3/5/7 and are **not** MAPI services. [MS-DOC]
<https://learn.microsoft.com/en-GB/deployoffice/oct/oct-2016-help-outlook-profile> No document
resolves whether that is a real restriction on the internet-account sections or loose wording
about the Exchange/Autodiscover problem. **That ambiguity is the go/no-go.**

### A.2 The literal `.prf`

Shipped as `Testbed/guest/tier-profile.prf`. ASCII, no BOM, CRLF. Provenance per key:

> **[GUEST-MEASURED, 2026-09-15; retired 2026-09-24]** This file, as described here, imports and
> leaves the account unbound (A.8 item 2). The route that works is the same file minus its PST
> service and `DefaultStore` - `Testbed/guest/tier-profile-forcepst.prf` - and the provenance below
> applies to it key for key, less those two. `Testbed/guest/New-TierProfile.ps1` now defaults to the
> forcepst file and refuses this one.

* **Section 1** — `Custom=1`, `ProfileName`, `DefaultProfile`, `OverwriteProfile`,
  `ModifyDefaultProfileIfPresent`, `DefaultStore` are all annotated in Microsoft's first-party
  whitepaper *"Outlook Deployment Options: Customizing a PRF File"* [MS-DOC]
  <https://download.microsoft.com/download/d/9/b/d9b5e9dd-c681-4430-a858-3427808677c1/PRFwhitepaper.doc>
  (the English SKU is withdrawn; this is the ja-JP SKU, and the literal config text and registry
  paths are language-independent). `Custom=1` also in KB Q259957. `BackupProfile` is the exception
  — see A.4.
* **Section 2** — `Unicode Personal Folders` / `MSUPST MS` and `Outlook Address Book` / `CONTAB`
  [MS-DOC, whitepaper Appendix B].
* **Section 3** — `Account1=I_Mail`; `I_Mail` is the POP3 account-type token, `IMAP_I_Mail` the
  IMAP one [MS-DOC, whitepaper Appendix B and the Office 2010 PRF article].
* **Section 4** — `UniqueService`, `Name`, `PathToPersonalFolders`, `EncryptionType` [MS-DOC].
  `EncryptionType=0x80000000` is "no encryption" per KB Q259957, chosen deliberately to sidestep
  the OCT/CIW bug in which Unicode PSTs get `0x40000000` where they should get `0x50000000`
  [COMMUNITY]
  <https://www.slipstick.com/outlook/tips-for-using-outlook-prf-files-to-configure-profiles/>.
  Unicode-ness comes from the **service** (`MSUPST MS`), not from `EncryptionType`. [INFERRED]
* **Section 5 — every key is [MS-DOC].** Microsoft's own literal POP3 block, with these exact
  spellings and `POP3Port=110` / `SMTPPort=25` / `POP3UseSSL=0` / `SMTPUseSSL=0` / `POP3UseSPA=0`
  / `SMTPUseAuth=0` / `LeaveOnServer=0x0` / `ConnectionType=0` / `DefaultAccount=TRUE`:
  <https://learn.microsoft.com/en-us/previous-versions/office/office-2007-resource-kit/cc764475(v=office.12)>
  and whitepaper Appendix B.
* **Sections 6/7** shipped verbatim from OCT-generated output. *"You typically do not modify
  existing entries in Sections 6 and 7 … However, if you define new services in the .prf file, you
  must add the appropriate mappings for those services."* [MS-DOC]
  <https://learn.microsoft.com/en-us/previous-versions/office/office-2010/cc179062(v=office.14)>

**Deliberately omitted:** `SMTPSecureConnection=0` — it appears in the Office 2007 RK sample but
has **no Section-7 mapping** in any OCT-generated mapping block examined, so it would be silently
discarded [INFERRED]. Also `Service=Internet Mail` in `[Service List]`: Microsoft's Appendix B does
not have it and there is no `[Internet Mail]` Section-6 mapping anywhere, so it is inert at best
[INFERRED]. The 1997-era `[Internet E-Mail] ServiceName=IMAIL` + `LongAccountName` convention from
KB Q259957 is **obsolete** and must not be used — the modern format replaced it with the `[I_Mail]`
Section-7 block. [INFERRED from the two mapping generations]

**No password key exists, by design.** No published POP3 `[AccountN]` block — Microsoft's or
anyone's — contains a password field, and one community source states outright that passwords are
deliberately not read from PRF files [COMMUNITY]
<https://qiita.com/flutter_bird/items/353c8d770fd061dc5614>. **The "no stored password"
requirement is therefore the only thing a PRF can do — a happy accident.** (The legacy
`Password=PT_STRING8,0x6703` in `[Personal Folders]` is the *PST* password, not the account's.)

**Semantics [MS-DOC, cc179062]:** `OverwriteProfile` = `Yes` (overwrite with a new profile) /
`Append` (preserve, update changed sections only) / `No`. `ModifyDefaultProfileIfPresent=True`
makes Outlook modify the default profile *even if the names differ*. There is no explicit format
version key; `Custom=1` is the only marker. Explicit warning: `DefaultProfile=No` together with
`OverwriteProfile=Yes` *"can produce unexpected results"*.

### A.3 `ImportPRF` — genuinely non-interactive, with one precondition

From the whitepaper, verbatim (paths ASCII, prose ja-JP; translation follows):

```
HKCU\Software\Microsoft\Office\11.0\Outlook\Setup\ImportPRF=\\server1\share\Outlook.prf
HKCU\Software\Microsoft\Office\11.0\Outlook\Setup\FirstRun
```

> *"Configure the registry so that the PRF file is imported when Outlook starts. Set the
> `ImportPRF` value to the location of the PRF file… Next, reset the Outlook `FirstRun` key so
> that Outlook processes the PRF file: delete the `FirstRun` registry value … or set its value to
> `0`."* [MS-DOC]

Corroborated without naming the value: after an OCT-customised install, *"A file named Custom12.prf
is placed in the \Program Files\Microsoft Office\ folder. **The registry is updated to import the
.prf file the next time you start** Office Outlook 2007."* [MS-DOC] cc764475.

**`11.0` to `16.0`** is [INFERRED], but Microsoft's own current docs use
`HKCU\Software\Microsoft\Office\16.0\Outlook\…` as the OCT-deployed (non-policy) hive for
2016/2019/2021/2024/365 [MS-DOC]
<https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/configuration/control-pst-use>.

* **`ImportPRF` is the silent path.** *"The `/importprf` switch will also directly launch Outlook
  and will execute at each logon and could therefore possibly also modify end-user alterations to
  the mail profile"*, whereas the registry value *"can be set without needing to open Outlook"* and
  applies *"the first time that Outlook is launched"*. [COMMUNITY]
  <https://robert365.com/article/deployprf>
* **Its documented precondition:** *"For this Registry value to work, the `FirstRun` and
  `First-Run` value may **not exist** in the Setup key."* [COMMUNITY] (same source; corroborated by
  <https://pixelchef.net/post/how-to-automatically-create-outlook-user-profiles-when-a-user-opens-outlook/>,
  which ships a `.reg` deleting both with `=-` and reports them as `REG_BINARY`). Microsoft's own
  text is weaker — "delete … or set to 0". **Delete both**; that satisfies both formulations.
* **Whether the `ImportPRF` value is consumed after import could not be established from any
  source, Microsoft or community.** Every source frames `FirstRun`/`First-Run` as the one-shot
  gate, which implies `ImportPRF` persists and is simply not re-evaluated [INFERRED]. Treat it as
  persistent and delete it yourself if you want one-shot semantics.
  **[GUEST-MEASURED, 2026-09-24] It IS consumed - the inference above was wrong. Outlook removes
  `ImportPRF` within about 5 s of the start that imports the file:** sampled every 5 s through a
  plain first start, a `/PIM` first start and a tier-profile import, it was gone at the first
  sample each time, with the new profile already listed; and the tier profile imported on
  2026-09-15 had lost it within three minutes, with nothing in the repository or its scratch ever
  removing it. The advice to delete it yourself is kept anyway, as a guard rather than a need:
  both profile scripts' `-Verify` remove a lingering value that names their own `.prf` once the
  import has run, and `Testbed/guest/Build-Corpus.ps1` refuses to build while one is set.
* `/importprf` *"Starts Outlook and opens/imports the defined MAPI profile"*; `/promptimportprf` is
  *"Same as /importprf except a prompt appears and the user can cancel"* [MS-DOC]
  <https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6>.
  So the switch is also non-prompting, but it necessarily launches Outlook. The registry value does
  not.

### A.4 The idempotency trap — confirmed

* **"Backup Of \<profile name\>" is real and Microsoft documents it.** In the OCT "Modify Profile"
  scenarios — which generate `OverwriteProfile=Append` — Microsoft states three times: *"The
  original existing profile is maintained, but renamed to **Backup Of \<profile name\>**."*
  [MS-DOC] cc764475. **Note this happens with `Append`, so `Append` is not a safe harbour.**
* **Microsoft names the suppressor, spelled `False`:** *"you must create a .prf file and set the
  properties **`BackupProfile=False`** and **`UniqueService=Yes`**."* [MS-DOC] cc179062. That is
  the only Microsoft mention found, and it is not in the key reference tables.
* **Community calls it undocumented and spells it `No`**, and reports that `DefaultProfile=Yes` +
  `OverwriteProfile=No` + `ModifyDefaultProfileIfPresent=False` alone is *not* sufficient.
  [COMMUNITY] Slipstick.
* **Which spelling 16.x honours is untested.** The processor parses `Yes/No` for two keys and
  `True/False` for another, so a shared boolean parser accepting both is likely [INFERRED]. If a
  single run leaves a `Backup Of…` profile, try the other spelling before concluding the key is
  dead.

**The idempotent combination** is `OverwriteProfile=Yes` + `BackupProfile=No` + `DefaultProfile=Yes`
+ `ModifyDefaultProfileIfPresent=FALSE` — replace in place, same end state regardless of run count
[INFERRED from the MS-DOC semantics]. `OverwriteProfile=No` is trivially idempotent but makes a
second run a no-op, so PRF edits never land — wrong for a testbed you will iterate on.

**For a checkpointed guest, do not rely on PRF idempotency at all.** Revert to checkpoint, import
once, and make *"exactly one profile exists, named `OutlookAI-Tier`"* a hard assertion of the
provisioning script. `Testbed/guest/New-TierProfile.ps1` does exactly that.

### A.5 The structural gap — honest assessment

**The hard fact.** `PROP_ACCT_DELIVERY_STORE` is identifier `0x0018`, type `PT_BINARY`, tag
`0x00180102`, read/write, *"Represents the Entry ID of the default delivery store for the
account"*, settable only via `IOlkAccount::GetProp`/`SetProp`. [MS-DOC]
<https://learn.microsoft.com/en-us/office/client-developer/outlook/auxiliary/prop_acct_delivery_store>

**A PRF cannot express it.** Across four independent Section-6/7 mapping blocks (whitepaper
Appendix B, KB Q259957, and two OCT-generated files) the only property types that ever appear are
`PT_STRING8`, `PT_UNICODE`, `PT_LONG`, `PT_BOOLEAN`, `PT_DWORD`. **There is no `PT_BINARY` mapping
anywhere, and the file syntax has no hex-blob literal.** [INFERRED from exhaustive inspection of
the published mappings]

**There is no `MailDeliveryStore` / `DeliveryStore` PRF key. Plainly, none exists.** The only
store-related key in the entire format is `[General] DefaultStore`.

**But `DefaultStore` is not a property write — it is an instruction to the processor,** and that
is why the PT_BINARY gap is not automatically fatal. Microsoft's own annotation:

> *"Specifies the Exchange Server or Personal Folders as **the delivery destination for new
> mail**. The delivery destination's parameters are defined in Section 2 below."* [MS-DOC,
> whitepaper]

The processor resolves the named service to a real store and can therefore write the EntryID
itself. **Whether the 16.x processor still does that, and whether it applies per-account
(`0x0018`) or only to the profile-wide `PR_DEFAULT_STORE`, cannot be determined without a guest.**
The model changed underneath that document: it is Outlook 2002/2003-era, i.e. *before* per-account
delivery stores existed.

**If `DefaultStore` is not sufficient.** Microsoft documents the 2010+ behaviour outright: *"By
default, Outlook 2010 automatically adds a new Outlook Data File (.pst) when you add a new POP3
account. However, earlier versions of Outlook let you select an existing Personal Folders file as
the default delivery location."* The **only** remedies Microsoft documents are UI — the "Deliver
new messages to → Existing Outlook Data File" radio at account creation, and "Account Settings →
Change Folder" afterwards. No registry or PRF knob is offered. [MS-DOC]
<https://learn.microsoft.com/en-us/previous-versions/troubleshoot/outlook/new-outlook-data-file-is-created-default-adding-pop3-account>

Corroborating that the 2007-era model was profile-wide: *"if you leave the **Deliver new mail to
the following location** option set to `<default>`, both Exchange and POP3 accounts deliver mail to
the Exchange mailbox."* [MS-DOC] cc764475 — i.e. the OCT's delivery-location control, whose PRF
representation is `DefaultStore`, genuinely governed POP3 delivery in that generation.

**Where an auto-minted PST lands, and the knob that moves it.**
`HKCU\Software\Microsoft\Office\16.0\Outlook`, value **`ForcePSTPath`**, type **`REG_EXPAND_SZ`**,
data a directory; sets the default directory for newly created PSTs and accepts environment
variables. [COMMUNITY — Microsoft Q&A answer, not a doc page]
<https://learn.microsoft.com/en-us/answers/questions/5720744/is-it-possible-to-set-the-default-pst-file-locatio>
The matching Group Policy is "Default location for PST files" in `outlk16.admx`, defaulting to
`%USERPROFILE%\Documents\Outlook Files\` when unconfigured [COMMUNITY, admx.help — **the page
returned HTTP 522 during this research and this is second-hand via search results; re-verify**].
`ForcePSTPath` is **not** mentioned in Microsoft's current PST-policy doc, which covers only
`PSTDisableGrow` and `DisablePST`. [MS-DOC] control-pst-use.

**The reframing worth putting to the maintainer.** The tests need *a* deliverable store they can
locate — not specifically the PST named in the PRF. If `DefaultStore` turns out not to bind the
POP3 account, the fallback is to drop `[Service1]`/`DefaultStore` from the PRF, set `ForcePSTPath`
to a known directory, let Outlook mint its own PST there, and have the harness discover it by
account rather than by path. **That converts an "unfixable hole" into a naming convention.**

> **[GUEST-MEASURED, 2026-09-15 and 2026-09-24] `DefaultStore` did NOT bind it, and this fallback
> is the route.** With a PST service and `DefaultStore=Service1` the import worked and
> `Account.DeliveryStore` came back NULL (`Testbed/guest/tier-profile.prf`, now retired). With
> neither, and `ForcePSTPath` set, Outlook minted `C:\OutlookAI-Tier\Outlook.pst` and bound the
> account to it - a store Outlook mints is a store Outlook binds. That is
> `Testbed/guest/tier-profile-forcepst.prf`, and since 2026-09-24 `Testbed/guest/New-TierProfile.ps1`
> defaults to it and writes `ForcePSTPath` itself; before that, only hand-run scratch scripts had
> set it. Two refinements this section could not have predicted: the binding happens at the start
> that first reaches the account - on a guest where the import was not Outlook's first start, the
> account stayed unbound until the next one - and the minted store is named `Outlook Data File`,
> so it is renamed afterwards (root-folder rename; `Store.DisplayName` follows).

### A.6 Prerequisites and silent-failure traps

1. **Directories must pre-exist.** *"The directories in the path to the personal folders must
   already exist."* [MS-DOC] KB Q259957.
2. **Cannot be applied to a running Outlook.** *"If Outlook is already open, queues the profile to
   be imported on the next clean start."* [MS-DOC] The registry route has the same constraint
   implicitly — it is read at startup.
3. **Double-clicking a `.prf` does nothing** since Outlook 2007: *"The .prf file is no longer
   associated with Outlook.exe."* [MS-DOC] cc764475. Any script relying on `Start-Process file.prf`
   is broken.
4. **Encoding.** Every published sample is plain 8-bit text and no source specifies an encoding.
   Ship **ASCII, no BOM, CRLF**: a UTF-8 BOM would prefix the first line and could break
   `[General]` detection. [INFERRED]
5. **Bitness is irrelevant to the file.** The PRF is parsed by `outlook.exe`, and
   `HKCU\Software\Microsoft\Office\16.0` is not WOW64-redirected. [INFERRED]
6. **Path quoting.** The registry value takes an unquoted raw path; prefer a space-free path and
   sidestep it. UNC is supported [MS-DOC, whitepaper example] but a local path removes a failure
   mode on an isolated guest. [INFERRED]
7. **Silent-do-nothing modes.** (a) A property with no Section-6/7 mapping is written nowhere —
   the flip side of *"you must add the appropriate mappings"* [MS-DOC cc179062]. (b) Historically a
   mid-file prompt caused Outlook to *"stop processing the rest of the .prf file settings … All
   other services in the .prf file are not processed by Outlook and are **silently not added**"* —
   fixed in 2007, but it establishes partial-application-without-error as a real failure shape for
   this processor [MS-DOC cc764475]. (c) `FirstRun`/`First-Run` present means `ImportPRF` is
   ignored, with no diagnostic [COMMUNITY].
8. **`UniqueService`** — `Yes` for services of which only one may exist (Exchange); `No` for PSTs
   and for POP3/IMAP accounts. [MS-DOC samples + COMMUNITY]
9. **First-run UI that `ImportPRF` does *not* suppress** — Office's "First things first"
   licence/privacy dialogs, the Click-to-Run activation prompt, the "Let's get started" screens.
   Budget for suppressing those separately on the guest image. [INFERRED]

### A.7 Registry values a setup script writes

| Key | Value | Type | Data / action |
| --- | --- | --- | --- |
| `HKCU\Software\Microsoft\Office\16.0\Outlook\Setup` | `ImportPRF` | `REG_SZ` | `C:\OutlookAI-Tier\tier-profile.prf` |
| `HKCU\Software\Microsoft\Office\16.0\Outlook\Setup` | `First-Run` | `REG_BINARY` | **delete** (fallback: set to `00`) |
| `HKCU\Software\Microsoft\Office\16.0\Outlook\Setup` | `FirstRun` | `REG_BINARY` | **delete** (fallback: set to `00`) |
| `HKCU\Software\Microsoft\Office\16.0\Outlook` | `ForcePSTPath` | `REG_EXPAND_SZ` | `C:\OutlookAI-Tier` — optional, only for the A.5 fallback |

Key path, value name, `REG_SZ` and the PRF path are [MS-DOC] (whitepaper, at `11.0`) with `16.0`
[INFERRED]. `First-Run`/`FirstRun` deletion and their `REG_BINARY` type are [COMMUNITY];
Microsoft's weaker form is "delete `FirstRun` … or set its value to 0". `ForcePSTPath` is
[COMMUNITY].

> **[GUEST-MEASURED, 2026-09-24] `ForcePSTPath` is not optional - it is half of the route that
> works** (A.5), and `Testbed/guest/New-TierProfile.ps1 -Execute` now writes it, as
> `REG_EXPAND_SZ`, with `New-ItemProperty` on the existing key, and reads it back. Two things to
> know when writing it by any other means: do not create the key with `New-Item -Force` - on an
> existing key that deletes every value under it, measured on a scratch key, and the hand-run
> scripts that first set this value did exactly that to the Outlook key; and it is per-user and
> outlives the import, so every PST Outlook later mints by default - a `/PIM` profile's store, for
> one - lands in the same directory.

### A.8 What could not be established without a machine

1. **Whether Outlook 16.x processes sections 3/5/7 (internet accounts) at all.** Every literal POP3
   PRF Microsoft ever published is 2000/2002/2003/2007-era. Against this sits the OCT 2016 sentence
   *"provided that the profile defines only MAPI services"*. **This is the go/no-go test.**
   **[GUEST-MEASURED, 2026-09-15] YES** - a genuine POP3 account (`CLSID_OlkPOP3Account`), with
   the sink's host, user and address from section 5.
2. **Whether `DefaultStore=Service1` binds the POP3 account's `PROP_ACCT_DELIVERY_STORE`,** or
   whether 16.x mints its own PST regardless.
   **[GUEST-MEASURED, 2026-09-15] It does NOT bind it** (`DeliveryStore` NULL). With no PST service
   and no `DefaultStore`, 16.x mints its own PST under `ForcePSTPath` and binds that - see A.5.
3. **Whether `BackupProfile` is honoured on 16.x, and whether the spelling is `No` or `False`.**
4. **Whether the `ImportPRF` value is consumed after a successful import.**
   **[GUEST-MEASURED, 2026-09-24] YES - Outlook removes it within about 5 s** of the start that
   imports the file; see A.3.
5. **Whether `ImportPRF` produces *zero* UI on a never-run 16.x profile.**
6. **Whether `POP3UseSSL=0` / `SMTPUseSSL=0` actually yields plain-text SMTP to `127.0.0.1:25`.**
   Modern Outlook exposes an *encryption method* enum (None/SSL/TLS/Auto); the PRF has only a
   boolean and no Section-7 key corresponds to the enum. Outlook could land on "Auto" and attempt
   STARTTLS against a dumb sink.
7. **Whether `.invalid` or a loopback literal trips client-side validation** in the PRF processor.
8. **Whether `%VAR%` expands in a `REG_SZ` `ImportPRF`.**
9. **Whether a PRF-created POP3 account appears as an `Account` in the object model with a usable
   `DeliveryStore`** — i.e. whether the harness can even see it.
   **[GUEST-MEASURED, 2026-09-15 and 2026-09-24] YES, by the A.5 route** - `Accounts.Count` 1,
   `AccountType` 2, `DeliveryStore` the minted PST, its Drafts folder resolving - **once Outlook
   has reached the account at a start**: imported at a machine's first Outlook start it was bound
   in that start; imported at a later one it was NULL until the next start.

**The single experiment that settles 1, 2 and 6 at once** is what `New-TierProfile.ps1 -Execute`
does and then verifies: import, then dump the profile hive under
`HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\OutlookAI-Tier`
and check for (a) an account subkey at all, (b) POP3 server/port values, (c) whether `00180102`
exists and which store EntryID it points at, and (d) whether a stray `.pst` appeared under
`%USERPROFILE%\Documents\Outlook Files`. **None of it may be run on the maintainer's host.**

---

## 4. Recommendation

### Order: **B first, then A, then C.** And B is not a fallback — it is the answer.

**1 — Route B, immediately, because it is the only route that makes the problem smaller.**

It costs no guest time, no Office build risk and no wizard: it is a re-reading of what the tier
needs, and the re-reading is already done above. **One account of any type buys 26 of the 34
blocked methods; the full sink buys one method more than a PST-seeding helper does.** Even if
Route A succeeds tomorrow, B's findings must land anyway, because three artefacts in this
repository are currently wrong about this (B.7) and one guard would silently go vacuous (B.5).

Do these in order, and none of them needs a VM:

* Correct `Docs/live-tier-on-the-vm.md:102` and `T2/LiveMailSink.cs:45` — **13, not six**, with the
  file-versus-method explanation so the error cannot regrow.
* Remove `Requires=Transport` from the 12 methods in B.2 that never put mail on the wire. Check
  `.github/scripts/check-pinned-constants.ps1` first: it fails the build if a capability name stops
  appearing in the runbook, and `Transport` must still appear.
* Regenerate `Docs/live-test-inventory.txt` per-method, and fix `Transport 41` at
  `Docs/vm-coverage-analysis.md:36`.
* **Run the `ConversationIndex` probe** (B.4). It is five minutes on the guest and it decides
  whether PST-seeding reaches 33 of 34 or stalls at 26. It must go through the tested helpers or
  `RemediationTools`, **not** an ad-hoc script — mailbox-safety rule 1.
* Rewrite the meaning of an absent `mailSink` block so "no transport" and "real transport" stop
  being the same state (B.5).

**2 — Route A, as the first thing tried on the guest, because it is cheap, deterministic and
locale-proof.**

A `.prf` is a text file and a registry value. No wizard, no `AutomationId`, no en-GB strings, no
build-dependent control layout — it either works or it visibly does not. `New-TierProfile.ps1` is
written to answer the go/no-go in one checkpoint cycle and to **fail loudly rather than report a
success it did not check**. The whole route turns on A.8 item 1, and that is a ten-minute
experiment on a machine that already exists.

Two things make this the right second step rather than the first: the OCT-2016 "MAPI services
only" sentence is unresolved, and the guest is **Office 2024** (`ProPlus2024Volume`,
`16.0.17932.20884`, 64-bit, per `Testbed/MEDIA.md`), which is newer than any PRF evidence found.

**3 — Route C, only if A fails, and only after the dump says it can work.**

Route C is real but it is the most expensive and the most fragile: it depends on a registry knob
with an **unresolved community report that it stopped working in Office 2024** — which is exactly
the guest's build — and on `AutomationId`s that Microsoft explicitly declines to guarantee across
builds. `Dump-UiaTree.ps1` exists to settle that for two minutes of a human's time before anybody
writes a driver. **Do not write the driver first.** If the dump shows empty `AutomationId`s or
`FrameworkId=DirectUI`, C is dead for a PowerShell 5.1 managed client and the answer is B.

### If all three fail

Route B alone still delivers **26 of 34** with no account work at all beyond one account of any
type — and the guest can be given an account of *some* type by hand, once, and checkpointed. The
question "how do we create a POP3 account programmatically" has an answer that was never asked for:
**create it once by hand and checkpoint the VM.** The testbed's whole premise is reproducible
guests from checkpoints; a one-time manual step captured in a checkpoint is reproducible in exactly
the sense that matters. That should be weighed seriously against funding Route C.

### Open questions for the maintainer

1. **Should the `Requires=Transport` correction land now, or wait for the guest?** It is a pure
   source change that CI can verify, but it narrows a trait that a future test might want.
2. **Should the `ConversationIndex` probe be added to `RemediationTools` (reusable, tested) or to
   the live suite as a one-shot?** The first costs more and leaves a tool; the second is faster.
3. **Is a one-time hand-built account captured in a checkpoint acceptable**, or must account
   creation be scripted end-to-end? This decides whether Route C is ever worth funding.
4. **Does smtp4dev actually serve POP3?** `Testbed/README.md` open question 11 already flags this,
   and if the answer is no, Routes A and C both satisfy a requirement that cannot be met anyway.

---

## 5. Honesty notes

* **Nothing here was executed against Outlook, MAPI, a mail profile, or the profile registry
  hive.** The `[MEASURED]` items are: repository greps and file reads; `Add-Type` of
  `UIAutomationClient`/`UIAutomationTypes` and type resolution; and `[Activator]::CreateInstance`
  of the `CUIAutomation` COM class to establish that its vtable is invisible to PowerShell. No
  Office process was started or attached to, and no `AutomationElement` tree was walked.
* **All three scripts carry the never-executed banner** and have been verified by **parsing only**
  (`[System.Management.Automation.Language.Parser]::ParseFile`), which is also what
  `.github/scripts/check-testbed-references.ps1` check 5 does.
* **This document lives in `.work/`, which is gitignored**, so it is not committed. If it should
  survive, it needs a home under `Docs/`.
