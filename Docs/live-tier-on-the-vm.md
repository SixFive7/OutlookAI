# The live-tier test VM: building it from nothing, and running it

**Who this is for.** Someone rebuilding this machine after it has been deleted, corrupted or
moved to another host, with nothing but this repository and a Windows ISO. It assumes no
knowledge of how the tier grew up. It is also the reference for running the tier once the
machine exists.

**Read `CLAUDE.md`'s Mailbox Safety section first.** Nothing here overrides it. Every rule in
it applies on a test machine too, and the guards described below are what enforce it.

**Secrets are not in this repository, which is public.** Guest account passwords, the host's
scratch paths and anything else that identifies a real machine live in the maintainer's own
notes. Where this document needs one it says which secret, never the value.

---

## 1. What this machine is, and why it is shaped that way

Four shape decisions drive everything below, and three of them are counter-intuitive enough
that they get their reasons here rather than in passing.

### 1.1 Two WINDOWS ACCOUNTS, because a profile cannot split the index

Half the live tier needs an indexed store and the other half needs an unindexed one. The
obvious arrangement, two data files in one Outlook profile with one of them excluded from
indexing, **does not work**, and the reason is structural rather than a setting anyone can
find.

Windows Search does not index Outlook per data file. It indexes a MAPI scope, and the scope
the **Indexing Options dialog** manipulates is **one URL per Windows user account**, of the
form `mapi16://{SID}/`, covering that account's whole Outlook profile. The dialog offers
exactly one switch per account: the profile is indexed, or it is not. Two stores in one
profile are therefore indexed together or excluded together *through that dialog*.

> **CORRECTED 2026-09-16 - the conclusion above holds, the reason this section used to give
> did not.** It said "there is no per-store URL underneath it". **There is one**:
> `mapi16://{SID}/StoreDisplayName($Hash)/`, and Microsoft documents excluding a single store
> with it through `ISearchCrawlScopeManager::AddUserScopeRule`. Per-store URLs are also
> visible in live index rows. What is actually true is narrower and is about the **GUI**:
> Outlook's shell extension reports "Microsoft Outlook" as the friendly name for any excluded
> `mapi` URL, so the dialog cannot even display the difference between a per-account and a
> per-store exclusion - which is why it offers one switch. The layout below is unaffected, but
> a document that justifies a layout with a false fact misleads whoever next reconsiders the
> layout, and someone would have concluded a per-store split was impossible when it is merely
> not reachable from that dialog.

So the split is per **Windows account**: one account whose Outlook profile is indexed and one
whose profile is not. That is the whole reason this machine has two logons.

> **The scope SHAPE is now measured** (2026-09-16): one `mapi16://{SID}/` rule per Windows
> account, carrying an `Include` flag, under `WorkingSetRules`. What remains **underived is
> DURABILITY** - whether that exclusion survives Outlook re-creating its own user rule, which
> is why `Testbed/guest/Set-OutlookIndexingDisabled.ps1` writes the Group Policy value as well
> (documented precedence: Group Policy > user > default). Verify before building anything
> else: add a store, let the indexer settle, and read `outlook_health`'s `index.perStore[]`.
> Section 4.1 of the findings behind that script says which fields and which values, including
> the fourth verdict that looks like success and is not. Section 8 says what to do if the
> durability half turns out to be wrong.

### 1.1a SUPERSEDED 2026-09-15 - there are two GUESTS now, so one account each

**Everything in 1.1 is history.** The build is now **two virtual machines**, `OutlookAI-Indexed`
and `OutlookAI-Unindexed`, with `Corpus A` on the first and `Corpus B` on the second - see the id
table in `Testbed/README.md` section 3. That is **exactly the fallback section 8 item 1 named**:
*"the fallback is two VMs, and the store layout collapses to one corpus per machine."*

**Three consequences, and the third is the valuable one:**

1. **Each guest needs ONE Windows account**, the `vmadmin` the answer file already creates. Section
   2.4's "create two local accounts" is **not work anyone has to do**.
2. **The toolchain is installed once per guest, not twice.** 2.4's "both accounts need the
   repository, the SDK and a built server exe" was the expensive half of that step.
3. **The riskiest unverified assumption in the whole layout is no longer load-bearing.** Whether
   Windows Search can exclude one Windows account's `mapi16://{SID}/` scope while indexing
   another's was *derived from how the scope is addressed, never measured*, and the entire
   three-store design rested on it. **Two guests do not need it to be true**: index state is now a
   property of the machine, which is the one thing Indexing Options unambiguously controls.

**We did not take this fallback because the assumption failed.** It was taken for an unrelated
reason - the decision to rebuild from nothing rather than repair the old guest - and the
simplification came free with it. Worth saying plainly, because "the assumption was disproved" and
"we stopped depending on the assumption" are different facts and only the second one happened.

**Two OUTLOOK PROFILES per guest is still required** - that is 1.2, and it is a different
constraint entirely: the corpus generator refuses any profile holding a mail account, so corpus
work and tier work cannot share one profile even on a machine with a single logon.

### 1.2 Two OUTLOOK PROFILES, because the corpus generator refuses an account

`corpus-build` refuses any profile that has **a mail account at all**, with no override flag.
That refusal is deliberate and it is not a nicety: the generator creates unsent items in bulk,
and the first real run put 5,532 of them into the target store's Outbox, inert only because
that profile could not send. On a profile with an account those would have been 5,532 real
messages queued for delivery.

So corpus work happens in a profile with **no accounts**, and the tier runs in a profile that
has the dummy account. Switching between them is a restart of Outlook, and it recurs: every
corpus rebuild is another switch.

### 1.3 Three stores, because two do not compose

| Store | Indexed | Purpose |
| --- | --- | --- |
| Corpus A | **yes** | the index tier, and the shape most `Requires=SearchIndex` tests want |
| Corpus B | **no** | the degraded path: no index frontier, the seven-day fallback window, the sweep and frame measurements |
| Bystander | either | the store the count tripwire actually watches, and the absent-arrival-folders shape |

**The "no" in row two is a property of the MACHINE, not of that store** (2026-09-16). The
exclusion is per Windows account, so on the unindexed guest **the hub and the bystander are
unindexed too** - every store on it is. That is intended and costs nothing, because the tests
that need an index run on the other guest; but the table reads as though only Corpus B were
affected, and a rebuilder will otherwise expect the hub on that guest to be searchable and
treat its empty `index.perStore[]` row as a fault.

The bystander is the one people leave out, and the tripwire is useless without it. The
tripwire **exempts the hub**, because the hub is where the suite writes; a machine whose only
store is the hub gets a guard that censuses, reports zero failures, and is structurally
incapable of reporting anything else. The bystander must therefore be a store **no test ever
touches**, and it must hold a few hundred items rather than none, because an empty store
exercises the item-by-item identity path over nothing.

A corpus is the wrong shape for that job. The identity budget is 500 items per folder and
3,000 per store; a 20,000-item corpus is over both in all four populated folders, so every one
of them falls back to a bare count. A few hundred items in a small store is what the guard
wants.

**A bystander is DECLARED, not merely listed - and listing it in only one place is a REFUSAL,
not a warning (2026-08-24).** It must appear in **both**:

* `expectedStoreDisplayNames` - which is what censuses it, and what `list_accounts` exactness
  counts. Removing it from here does not make it a bystander; it makes it invisible.
* `bystanderStoreDisplayNames` - which is what the write allowlist refuses on.

Name it in one and not the other and the tier stops, deliberately. The half-declared state used
to be a warning, which meant a machine could run for months believing a store was protected
when the guard had never been told.

**Both corpus stores are declared bystanders too.** They are stores no test may write to, which
is exactly what the declaration means. Before that was true, the identity tests resolved to
`Corpus A` and drafted **into the measurement corpus**.

**`Corpus B` is deliberately NOT declared in this machine's settings file.** It lives in the
other Windows account's profile, and a declared bystander the running profile does not mount is
censused, not found, and refuses the tier. It belongs in *that* account's settings file.

### 1.4 A dummy account, and NO SINK - decided 2026-09-15

**DECISION: the testbed guests get no mail sink.** The dummy account exists and is fully built by
script - it resolves over COM with `SmtpAddress`, a bound `DeliveryStore` and a Drafts folder -
but nothing listens on `127.0.0.1`. Mail submitted there goes nowhere.

**Why, when a sink turned out to be buildable after all.** smtp4dev *does* serve POP3: the issue
that said otherwise was closed "not planned" in 2022 and POP3 shipped three years later from an
unrelated pull request. `Pop3Server.cs` is absent at tag 3.10.3 and present at 3.11.0, and 3.15.0
is the version to pin. So this is not a "cannot"; it is a "not worth it", and the arithmetic is:

* A sink buys **exactly one** test method that nothing else covers - the Phase 5 two-step `send`
  round-trip. One account of any type already reaches 26 of the 34 otherwise-blocked methods, and
  seeding items straight into the PST reaches 33.
* It costs a **third precondition**. `Testbed/MEDIA.md` names two - a Windows ISO and the Office
  Deployment Tool - and `CLAUDE.md`'s Dependencies rule forbids anything a rebuilder must obtain
  beyond those. The licence is fine (BSD-3-Clause); the precondition is the problem.
* It costs an **undocumented manual step**: smtp4dev's POP3 refuses an empty password and the
  `.prf` deliberately carries none, so somebody has to type one once per guest.
* And the implementation is eleven months old with two RFC 1939 violations found by reading its
  source - `TOP` advertised in `CAPA` with no handler, and `DELE` re-listing the mailbox per
  command so `DELE 1; DELE 2` removes the wrong message.

**WHAT THIS GIVES UP, stated plainly rather than discovered later.** The Outbox stops being a
canary. `LiveMailSink.EnsureOutboxDrained` and the zero-artifact sweep over folder 4 exist to
catch a genuine send-path leak; with no transport the Outbox can never fill from a seed, so it can
never prove a leak either. **That guard goes vacuous on the guests** - it will pass, and its
passing will mean nothing. It still means something on the maintainer's machine, which has real
transport.

**The reserve.** `Testbed/guest/Install-MailSink.ps1` is written, staged-package-only,
SHA-256-pinned, and its `-Verify` speaks SMTP and POP3 itself. It is not part of the build path.
If the Phase 5 method ever has to run on a guest, that script is the route and this decision is
the thing to revisit - not the research, which is done and is in `Docs/research/`.

### 1.4a Why the account still points at a sink that is not there

The account exists because `NewDraft` resolves an `Account` object by SMTP address and refuses
when none matches, which is what puts the entire draft, update/discard, HTML-draft and send
families out of reach of an account-less machine.

Pointing it at an unroutable server was the first plan and it is wrong. A send **queues** and
never leaves, the Outbox is in the mandatory zero-artifact sweep, and so every run that sent
anything would fail its own teardown forever on residue nothing could remove.

**"Six live methods also need the mail to genuinely arrive" - CORRECTED 2026-09-15, it is 13, and
the error was mechanical.** Six *files* hold an arrival wait; `T2/LiveInboxArrival.cs` says so in
its own header, describing the five verbatim copies it replaced "and a SIXTH copy [that] was
missed". A file count was read as a method count, and three further arrival waits were never in
that consolidation at all - `LiveFreshModeTests`' own loop, `LiveSweepScopeTests`' 240 s loop, and
the T3 stdio pollers. **13 methods put mail on the wire and all 13 depend on it arriving.**

**And the number that actually matters is 1.** Of 127 live methods, exactly one needs the wire in
a way nothing else can substitute: `T3/Phase5LiveMcpToolShapeTests.SendTool_TwoStepFlow_RoundTrip_-
OverRealStdio_WithAuditLines`, which is the only test that calls the product's own `send` tool with
a valid token and lets Outlook submit. The other 12 need an item to *appear in the Inbox with a
known subject*, which direct PST creation already produces - the corpus generator builds 20,000
such items, and `LiveOutlookTestMailer.SaveTaggedDraftWithAttachments` already made exactly this
trade for exactly this reason, noting that **drafts are indexed exactly like received mail**.

**`Requires=Transport` over-declares by 12**: 25 methods carry it, 13 use it.

**So the account is configured for a local sink that delivers back - and on the guests there is
nothing there.** It points at `127.0.0.1` on 25 and 110 because that is what the design calls for
and what the reserve installer would satisfy; per 1.4 no sink is installed, so a send queues in
the Outbox and stays there. On a guest, do not send. The 13 methods that put mail on the wire are
out of reach there by design, and 12 of them are reachable instead by seeding items straight into
the PST, which is what the corpus generator already does 20,000 times.

---

## 2. Building the machine from nothing

Do these in order. Checkpoint where the section says to; a checkpoint is much cheaper than
redoing the step above it.

### 2.1 The hypervisor and the guest

* Hyper-V guest, named `OutlookAI-TestVM` by convention. Generation, firmware, vCPU, RAM and
  disk size are **not recorded anywhere and are yours to choose**; see section 8.
* Windows 11. The edition, image and locale the guests are built to **are** recorded, in
  `Testbed/MEDIA.md` (see section 8, item 6); the exact build of whatever you install is still
  yours to record beside the VM. **An Outlook build difference is the first thing to suspect when
  a live test behaves differently here than on the maintainer's machine** - and that is not
  hypothetical. Measured 2026-09-15: the guest's Office is **3,598 builds ahead** of the
  maintainer's - 16.0.**17932**.20884 (`ProPlus2024Volume`, `PerpetualVL2024`) against
  16.0.**14334**.20848 (`ProPlusSPLA2021Volume`, `Production::LTSC2021`). That gap is deliberate
  and accepted; it is a known limit, recorded in section 9 and in `Testbed/MEDIA.md`.
* **Networking should be Internal, Private or disconnected.** The sink binds to loopback and
  nothing on this machine needs to reach the internet after the toolchain is installed. A test
  VM with a mail server on it and a route to the outside is an open relay waiting to happen.
* **Auto-logon, no lock screen, no sleep.** Outlook never finishes starting in session 0, so
  anything driving it must run in an interactive session. Reached over PowerShell Direct that
  means a scheduled task registered with `-LogonType Interactive`, never a direct remote call.
  A guest that sleeps mid-build loses a twelve-minute corpus run.
* The guest shell is **Windows PowerShell 5.1**. No `??`, no ternary, no `-p` on `mkdir`.
* Set the guest's **time zone and locale deliberately and write them down**. Outlook parses
  DASL date literals in the MACHINE locale, and a day-first literal on a Dutch-locale box
  silently returns the wrong rows. The corpus tool formats its own literals year-first for
  exactly this reason, but nothing protects a query typed by hand.

**Checkpoint `CP-01-WIN-CLEAN`.**

### 2.2 Office

* **Classic Outlook, desktop.** The "new Outlook" has no MAPI and no `Outlook.Application`, so
  the entire suite dies the moment the machine is migrated to it. Pin classic explicitly and
  suppress the migration toggle.
* Version, channel and **bitness are unrecorded**. The test host is `net10.0-windows` with
  `PlatformTarget x64`, partly because the `Search.CollatorDSO` OLE DB provider the index tier
  reads needs an x64 host. Whether Office itself must be x64 is untested; x64 is the safer
  choice and is what you should record.
* Suppress the first-run wizard and the "add an account" prompt. A profile that opens a dialog
  cannot be driven over COM, and the suite cannot answer one.
* Pin the update channel. An Office auto-update invalidates the Office checkpoint silently.

**Checkpoint `CP-04-OFFICE-GOLD`** once Outlook opens to an empty profile without prompting.

### 2.3 Toolchain, repository and add-in

> **CORRECTED 2026-09-17, and it was wrong in three ways.** This section used to say ".NET SDK
> (version unrecorded), git, and a clone of this repository", which described the hand-built
> August guest and was never true of the scripted ones. **The guests have no SDK, no git, no
> clone and no network**, and that - not anything about Outlook - is why the live tier has never
> run on one.

* **.NET SDK 10.0.401, win-x64, installed from STAGED media** by `Testbed/guest/Install-DotnetSdk.ps1`
  (`Testbed/MEDIA.md` declares the precondition). Nothing pins a feature band - there is no
  `global.json` in this repository and CI asks `setup-dotnet` for `10.0.x` - so any .NET 10 SDK
  would compile. 10.0.401 is chosen because it is what the host runs, and the host publishes the
  payload the guest measures with: one toolchain across both is one fewer difference to suspect.
  **x64 is not optional** (`PlatformTarget x64`, and `Search.CollatorDSO` has no 32-bit story here).
* **`net48` needs no separate install.** `OutlookAI.Core` carries `Microsoft.NETFramework.ReferenceAssemblies`,
  so the reference assemblies travel in the offline package feed. The tests project targets
  `net10.0-windows` only, so a `dotnet test` run never builds the `net48` target at all.
* **No git and no clone.** The source arrives as a `git archive` of a NAMED COMMIT, and every
  NuGet package in the restore closure arrives as an offline folder feed - both staged by
  `Testbed/host/Publish-LiveTierPayload.ps1`, which re-restores the whole suite against that feed
  **on the host** first, so a feed short of one transitive package fails in seconds here rather
  than after a slow copy-in there. An archive rather than the working tree, so no stale `obj/`
  travels across carrying the host's package paths.
* Build once so the server exe exists where the tier-3 tests look for it. That path is baked
  into the test assembly at build time as `AssemblyMetadata("McpServerExePath")` and points at
  `McpServer\OutlookAI.McpServer\bin\<Config>\net10.0-windows\OutlookAI.McpServer.exe`.
  **This stops being a question once the guest builds the suite itself**: a guest that builds it
  bakes a path into its own tree where its own build just put the exe. The staged
  `C:\OutlookAI-Q5\server\` payload stays what it is, and the two no longer have to be the same
  path.
* Install the add-in and let it run once. Tests carrying `Requires=AddInRegistry` read tuning
  state the add-in writes on first run; without it they have nothing to read.

**Checkpoint `CP-03-OUTLOOKAI-INSTALLED`, then `CP-05-ADDIN-TRUSTED`.**

### 2.4 The two Windows accounts - NOT NEEDED, skip to 2.5

**SUPERSEDED 2026-09-15. Do not do this step.** Each guest has one Windows account, `vmadmin`,
created by the answer file, and index state is a property of the **machine** rather than of an
account - see 1.1a. Creating a second account, and installing the repository, the SDK and a built
server exe under it, is work with nothing behind it.

The original text is kept below because it explains a design somebody may meet in the old guest,
and because if the two-guest arrangement is ever collapsed back to one machine this is what it
would have to become again.

---

Create two local accounts. Section 1.1 says why. Suggested roles, since neither is recorded:

* an **indexed** account, whose Outlook profile carries Corpus A and the dummy account, and
  where the index tier and the send path run;
* an **unindexed** account, whose Outlook profile carries Corpus B, with its `mapi16://{SID}/`
  scope excluded.

Both accounts need the repository, the SDK and a built server exe, or the tier can only run
under one of them. Whether that is a clone each or one clone with both accounts granted access
is your call; record which.

---

**WHAT ACTUALLY MAKES A GUEST UNINDEXED, now that there is one account per guest.** The
sentence above named a GUI on a machine that is driven headlessly, and that sentence was the
entire specification of half the testbed. The step is
**`Testbed/guest/Set-OutlookIndexingDisabled.ps1`**, run on the unindexed guest. It writes two
layers - the documented Group Policy value `PreventIndexingOutlook` and the `mapi16://{SID}/`
rule's `Include` flag - and **leaves the indexer running**, because stopping the Windows Search
service produces a machine with no search rather than a mailbox search has not been told about,
and the product takes a different, untested code path there. Section 8 item 21 keeps that
distinction from collapsing.

**Run it BEFORE `Build-Corpus.ps1`, and this is LOAD-BEARING rather than an optimisation.**
Exclude first and no corpus row is ever crawled, so the "indexed, but not **yet**" state cannot
arise and does not have to be waited out. `Testbed/README.md` section 1 carries it as step 7b.

> **Why it is load-bearing (2026-09-16).** Nobody has established whether setting
> `PreventIndexingOutlook = 1` **removes rows already in the catalog** or merely stops new ones
> being added - Microsoft's wording is ambiguous and no source resolves it. If it only stops
> adding, then a store crawled before the exclusion **stays searchable indefinitely**, and the
> `-Verify` probe's two readings cannot tell "these rows persist by design" from "the indexer
> has not settled yet". Excluding first makes the question not arise. If you ever have to
> exclude a guest whose corpus is already built, do not trust a settle window - rebuild the
> catalog, or rebuild the guest.

**Three durability risks the script does not defend against**, all established 2026-09-16 and
none of them a reason to avoid it - they are the reason the Group Policy layer is written *as
well as* the registry rule:

* **The registry rule can be silently clobbered.** Crawl-scope state lives in a shared memory
  view with the registry as backing store (`ISearchCrawlScopeManager2::GetVersion` "does not
  result in a cross-process call" and hands back a mapped view). So any other crawl-scope client
  calling `SaveAll()` can rewrite `WorkingSetRules` from its own in-memory copy and undo
  `Include = 0`. The supported equivalent is `AddUserScopeRule(url, fInclude: FALSE, …)` then
  `SaveAll()` - unreachable from PowerShell 5.1 without hand-declared COM vtables.
* **`RevertToDefaultScopes()` deletes the rule outright**, and there is no default to fall back
  to: the `mapi16` rule is `Default = 0` and no `DefaultRules` entry for `mapi16` exists anywhere.
* **`PreventIndexingOutlook` is machine-wide and is not indexing-only.** It is an HKLM `Machine`
  policy with no per-user variant, so it hits **every** Windows account on the guest - unlike the
  per-SID `mapi16` rule. And Outlook reads it itself and switches its own UI to built-in search,
  surfacing a banner about search performance. That does not change what this product measures
  (the server queries the catalog directly, not Outlook's UI search), but it does mean mechanism
  one is a **client-behaviour change**, not purely a scope change.

**Delegate mailboxes are governed separately** and `PreventIndexingOutlook` does not reach them:
`Search.adml` says in terms that "the 'Enable Indexing of Uncached Exchange Folders' has no effect
on delegate mailboxes. To stop indexing of online and delegate mailboxes you must disable both
policies" - the second being `PreventIndexingUncachedExchangeFolders`, same key, **inverted
polarity**. No guest has a delegate mailbox, so nothing here depends on it today; it is recorded
because the delegate stores are the ones the tier treats as read-only production data.

**The script REFUSES while Outlook is running**, and that refusal protects the guest rather than
the measurement: Microsoft documents that the PST provider is "very sensitive to the indexing
state changing while the PST is open", and that if it changes "the PST may end up kicking off an
installer to repair Outlook". Quit Outlook gracefully or restart the guest - **never** `taskkill`
it.

**Verify by asking the INDEX, not the registry.** `Set-OutlookIndexingDisabled.ps1 -Verify` runs
a control probe, a scoped-MAPI probe and a scope-free mail probe through the same
`Search.CollatorDSO` provider the product uses, **twice**, `-SettleMinutes` apart - because a
machine believed unindexed while it is quietly still indexing produces measurements that look
fine and mean nothing. It returns four verdicts, and **two of them are not answers**:
`SETTLING` and `NO-INDEXER` both mean "ask again", not "pass". Reading back the settings that
were written proves nothing; neither does a single zero. Cross-check with `outlook_health`'s
`index.perStore[]`, which is the instrument section 1.1 names.

**That verification needs an x64 host** - the `Search.CollatorDSO` provider has no 32-bit
in-process form, so a 32-bit PowerShell cannot load it. This is a precondition of every index
check on the guest, not only of the build.

### 2.5 The Outlook profiles

Two profiles are needed on the account that does corpus work:

* a **corpus profile with no mail accounts at all** (section 1.2), and
* a **tier profile** with the dummy account.

Set Outlook to "always use this profile" and switch by changing that setting, not by prompting:
a prompting profile cannot be driven over COM. The Mail control panel works and was the assumed
route.

**All of it is scripted, and all of it has now run on a guest.** The first attempt at the scripts
went through Extended MAPI's `IProfAdmin`, which is **measured broken** on Office LTSC 2024
16.0.17932.20884 (2026-09-16: `E_NOINTERFACE` on `IID_IProfAdmin`, with `MAPIInitialize` and
`MAPIAdminProfiles` both succeeding; cause not established). The rewrites then ran on
`OAI-UNINDEXED` on **2026-09-24**, each from a fresh restore of `CP-05-CORPUS-B-CLEAN-UNINDEXED`:

| Step | Route | What the guest run showed |
| --- | --- | --- |
| Switch the default, prompt off | `Testbed/guest/Set-DefaultOutlookProfile.ps1` - the HKCU `DefaultProfile` and `PickLogonProfile` values | **Works end to end.** Both values written and read back; a profile that does not exist is refused with the key's own last-write time unchanged; a running Outlook is refused. After a restart Outlook opened the named profile with no prompt, and COM agreed: `CurrentProfileName`, `Accounts.Count`, the store names. |
| An account-less profile | `outlook.exe /PIM <name>` | **The route, and measured again.** Opens the new profile with no dialog at all ("Outlook Today", on a guest whose default stayed `CorpusProfile` - `/PIM` does **not** change the default). It names its one store `Outlook Data File`; both guests' `CorpusProfile` were made this way. |
| A named PST into an existing profile | `Testbed/guest/Add-OutlookPstStore.ps1` - `NameSpace.AddStoreEx`, then a root-folder rename (section 2.6) | **Works, every path, unchanged.** `AddStoreEx` returned at once (the spin it warns about did not happen); `Store.DisplayName` followed the rename, `@` included; a re-run is a no-op, a new name renames without a second attach, a taken name or the wrong profile is refused before anything runs. |
| An account-less profile from a `.prf` | `Testbed/guest/New-OutlookProfile.ps1` | **The import works; the profile it makes does not open unattended.** Outlook consumes `ImportPRF` within 5 s, creates the profile and every named PST - two in one `.prf` - and then stops on its **"Email Account Setup"** dialog, at that start and every later one, with First-Run absent and again with it put back. COM reads "You are not connected". `-Execute` now **refuses** unless `-AcceptAccountWizard`, and names the `/PIM` route. |

So **the corpus profile is `/PIM` plus `Add-OutlookPstStore.ps1`**, then - after a guest restart,
because it refuses while Outlook runs - `Set-DefaultOutlookProfile.ps1` to make it the default. That
is also exactly how both guests' corpus profiles already exist.

Two things the runs showed about the machine rather than the scripts, recorded here because this
is where a rebuilder meets them:

* **Every start of the tier profile raises the POP3 logon dialog** ("Internet Email - tier", "Enter
  your user name and password for the following server"): the account points at 127.0.0.1:110 and
  stores no password. It did **not** block COM - the read above ran with it on screen.
* **Reading an e-mail address over COM can raise Outlook's object-model guard** - "A program is
  trying to access email address information stored in Outlook", Allow / Deny - and that prompt
  **blocks the call that raised it** until a human answers. It happened on 2026-09-24 when a probe
  read `Account.SmtpAddress`; the same read did not prompt on 2026-09-15. **`CP-05`'s saved memory
  already carries one**: the checkpoint's Outlook (running since 2026-09-16 02:11:41) shows that
  prompt the moment the checkpoint is restored, before any COM call. Defender's signatures on this
  guest are 372 days old (last updated 2025-09-17 - it has no network), and an out-of-date antivirus
  is Microsoft's documented trigger for the guard; that it is the trigger HERE was not proven.
  **Anything that reads addresses on these guests should expect it - the live tier does.**

`Testbed/README.md` section 4b is the summary; `Docs/research/profile-automation-research.md` is the
evidence with every claim labelled by source. Run `New-OutlookProfile.ps1 -Preflight` first - it
reads only, makes no COM call and no MAPI call, and needs no checkpoint.

### 2.5a `ImportPRF`: Outlook clears it itself - measured - and it is cleared anyway

**The worry.** The profile scripts build profiles by pointing `ImportPRF` (under
`HKCU\...\Outlook\Setup`) at a `.prf` that carries `OverwriteProfile=Yes`. If Outlook read that
value at every start and never removed it, every later start would rebuild the profile from the
file, and a store attached or filled since - the corpus - would stop being part of it: no `.pst`
deleted, a corpus that has apparently vanished, and a ~13-minute rebuild of 20,000 items.

**MEASURED 2026-09-24 on `OAI-UNINDEXED`: Outlook REMOVES `ImportPRF` within 5 s of the start that
imports it.** Sampled every 5 s through a plain first start and through a `/PIM` first start - gone
at the first sample both times, with the new profile already listed. And the tier profile had
already answered it after the fact: `New-TierProfile.ps1 -Execute` set the value at 2026-09-15
19:54:37 (its own log), Outlook imported at 19:54:43, and the value has been absent ever since, with
First-Run present and the `Setup` key last written at 19:57:59, three minutes after that start - while
nothing in the repository, or in the scratch that drove the guest, ever removes it. **CLEARED.** So
on the normal path no later start re-imports.

**It is cleared anyway** - the maintainer's choice: remove it regardless of the answer, and do not
make the protection depend on anyone remembering a flag. Three places, none of them a flag:

* **`-Verify` removes it.** `New-OutlookProfile.ps1 -Verify` and `New-TierProfile.ps1 -Verify` take
  out a lingering `ImportPRF` that names their own `.prf` - but only once First-Run is back, because
  removing it before Outlook has read it cancels the import itself. Run `-Verify` too early and it
  FAILS, says the import has not happened yet, and leaves the value exactly where it was (measured).
* **`Testbed/guest/Build-Corpus.ps1` refuses while it is set** - a third preflight precondition,
  fail-closed (an unreadable value refuses too, because nothing downstream ever re-checks it), and
  **not** skipped by `-SkipPreflight`. The corpus is precisely what a rebuild would detach.
* `New-OutlookProfile.ps1 -ClearImportPrf -Execute` still removes it by hand, unconditionally.

The reasoning lives beside the code that acts on it: `Resolve-ImportPrfClearance` in
`Testbed/guest/New-OutlookProfile.ps1` and `Test-CorpusImportPrfGuard` in
`Testbed/guest/Build-Corpus.ps1`, each walked by its script's `-SelfTest`.

**Turn AutoArchive OFF, on every store, in both profiles.** It is a client-side actor that
moves items out of a PST on a schedule, and to a before/after census that is **indistinguishable
from mail loss** - the count tripwire would fail a run over Outlook tidying up on its own. It is
the one such actor a test machine can realistically have, so it is worth turning off explicitly
rather than assuming the default.

### 2.6 The stores

**NARROWED 2026-09-15, deliberately and with the boundary written out rather than left to
precedent.** The old text said: *"Add every store through Outlook itself. Do not improvise a
script."* The rule's real subject is **items in a mailbox that matters** - ad-hoc shell code
deleting real mail is what it was written for - and creating an empty data file on a disposable
guest is a different act. So the line now runs between items and stores, not between GUI and
script:

* **No script may create, delete, move or modify an ITEM.** That is unchanged, absolute, and
  `CLAUDE.md` mailbox-safety rule 1 remains the authority: item mutation goes through the tested
  helpers or the shipped MCP tools, never through improvised code.
* **Creating a NEW, EMPTY store on a testbed guest is permitted from a script**, provided the
  script verifies which machine and which account it is running as before it writes anything.
  `Testbed/guest/Add-OutlookPstStore.ps1` and `New-OutlookProfile.ps1` are that, and they refuse
  unless logged on as the guest's own account.
* **Attaching an existing store that already holds items is NOT covered by this narrowing.** It
  is one step from there to a script that opens a real mailbox, which is where the original rule
  came from.

**Why write the boundary out rather than just permitting it.** A rule narrowed by precedent keeps
narrowing - the next person reasons "it's only a store" and then "it's only one item". A rule
narrowed by text stops where the text stops.

The GUI route (File > Account Settings > Data Files > Add) remains correct and is what to use if
you are building a guest by hand.

> **HOW THE TIER STORE GETS ITS NAME, measured 2026-09-15.** Outlook names the store it mints
`Outlook Data File`, and the tier profile REQUIRES Outlook to mint it (only a store Outlook mints
gets bound as the account's delivery store). So the name is set afterwards, by
`Testbed/guest/Rename-OutlookStore.ps1`, and this works:

* `Store.DisplayName` is read-only, and `PropertyAccessor.SetProperty` on `PR_DISPLAY_NAME_W` is
  widely reported blocked - but **renaming the store's ROOT FOLDER does work, and
  `Store.DisplayName` FOLLOWS.** No source could answer that; it is now measured.
* The `@` is accepted: the store reads `tier@vm.invalid` over COM.
* The account's `DeliveryStore` reports the new name too, so renaming does not break the binding.

**A TENSION WORTH NAMING RATHER THAN QUIETLY RESOLVING (2026-09-15, restated 2026-09-17).**
> `Testbed/guest/Add-OutlookPstStore.ps1` is a script that adds stores, which is the thing the
> paragraph above tells you not to write. It exists because the display name has to be exact and
> the GUI route costs a vision-model session, and it is drawn as narrowly as the objection
> deserves: it creates **new, empty** PST files and registers them in a profile, it never reads an
> item, moves one or deletes one, and it refuses to run on any machine not logged on as the guest
> account. So it is not in the class of thing that destroyed real mail - that was shell-side
> subject matching against a live mailbox.
>
> **ONE THING IN THAT LIST CHANGED AND IS WORTH SAYING OUT LOUD.** The script now names a store by
> **renaming the store's root folder**, because the MAPI route that set the name at creation time
> is measured broken (2026-09-16, `E_NOINTERFACE` on `IID_IProfAdmin`). So it does now modify a
> FOLDER - the store's root node - where the older version did not. That is deliberately on the
> permitted side of the line drawn above, and for the same reason
> `Testbed/guest/Rename-OutlookStore.ps1` gives: the rule's subject is **items**, the folder in
> question is a store this project created and has just attached, and no item is created, deleted,
> moved or modified by it. If that reading is too generous it is the maintainer's call to narrow
> it - but it should be narrowed explicitly, not by leaving this paragraph describing a script
> that no longer matches it.
>
> **Its guest half has still never been executed**, so nothing here is yet evidence about the
> script - though the mechanism under it is measured, by `Rename-OutlookStore.ps1`, on this build.
> Its decision logic is covered by `-SelfTest` (39 assertions), which is a statement about the
> decisions and not about Outlook. Whether this paragraph's rule should be relaxed for store
> creation specifically is the maintainer's call, not a script's: until it is made, the GUI route
> above remains the recorded one and the script is a draft beside it.

Naming matters more than it looks:

* **The hub store must be named exactly the dummy account's SMTP address.** Several tests use
  `testHubStoreDisplayName` as an address (`NewDraft(Hub, Hub, ...)`, `FindAccountBySmtp(Hub)`),
  so the hub PST has to be called something like `test@vm.invalid` literally. **Outlook accepts
  `@` there - ANSWERED and measured; section 8 item 2 has the evidence and the two caveats.** This
  used to be called out here as untested and gating the whole draft family; it is neither any
  more, and it is not restated here so that there is one place to correct if it ever changes.
  `.invalid` is guaranteed unresolvable by RFC 2606, so a misconfiguration cannot leak mail
  anywhere.
* **The dummy account's delivery store must be a separate throwaway PST**, not a corpus store.
  An account delivering into the corpus store can flip that store's `IsDataFileStore`, and
  `CorpusSafety` reads that property as one of four independent facts it requires before it
  will write anything. Get it wrong and the generator refuses that store permanently.
* **The bystander is DECLARED in two places and is never the hub.** It goes in
  `expectedStoreDisplayNames` (which censuses it) **and** in `bystanderStoreDisplayNames`
  (which makes the write allowlist refuse it). Naming it in only one of the two **refuses the
  tier** - see section 1.3. The same double declaration applies to `Corpus A`.

Populate the bystander with a few hundred ordinary items. The corpus generator cannot honestly
do this: it tags everything it creates, and the bystander's whole job is to be untouched.

### 2.7 The mail sink

**The sink is a third-party component and is deliberately not in this repository.** A loopback
SMTP-plus-POP3 server is a few hundred lines of RFC 1939 whose failure modes - dot-stuffing a
body line that begins with a period, UIDL identities that move when the store is recreated,
`STAT` octet counts - all produce INTERMITTENT wrong answers against Outlook, which is the
fussiest POP3 client there is. This suite exists to eliminate intermittent artifacts; writing a
new source of them to serve it is the wrong trade, and a maintained component already does the
job.

**Use smtp4dev** (`rnwood/smtp4dev`, BSD-3-Clause, actively maintained). It is the only
candidate that is simultaneously deliver-back, maintained, a native Windows service, and
catch-all by default. That last point matters here specifically: the hub is named after its own
fabricated address, and smtp4dev's auto-created mailbox accepts `Recipients="*"`, so there is
nothing to provision per address and nothing to re-provision after a rebuild.

Install and configure:

```
winget install RnwoodLtd.smtp4dev
```

Then, in `appsettings.json` beside the executable:

* `AllowRemoteConnections: false` - **it ships as `true`; change it.** Loopback only.
* SMTP on 25, POP3 on 110, **IMAP disabled** (nothing here needs it).
* `AuthenticationRequired: false`, `SecureConnectionRequired: false`, `TlsMode: "None"`.
* Leave `Mailboxes: []` so the catch-all is created automatically.
* `Urls: "http://localhost:5000"` for its web UI.

Register it as a service, which is what keeps it windowless and running before Outlook starts:

```
smtp4dev --install-service
sc.exe start Smtp4dev
```

Use `Rnwood.Smtp4dev.exe`, **not** `Rnwood.Smtp4dev.Desktop.exe`: the Desktop build creates a
window, which this machine must never do.

Before committing to ports 25 and 110, check they are free and not inside a reserved block -
Hyper-V and WinNAT genuinely do reserve ranges on a VM:

```
netsh interface ipv4 show excludedportrange protocol=tcp
netstat -ano -p tcp | findstr ":25 "
```

Nothing about the tests needs the well-known numbers; 2525 and 1110 are fine, and the settings
file carries whichever you pick. **Create no inbound firewall rule.** Loopback traffic is not
filtered, so a listener that needs a rule is a listener bound to `0.0.0.0`, which on a test VM
is an open relay.

### 2.8 The dummy account

Add a POP3 account in the tier profile, pointing at the sink:

* incoming POP3 `127.0.0.1:110`, outgoing SMTP `127.0.0.1:25`, encryption **None**, any
  credentials (the sink accepts anything);
* **"Deliver new messages to" must be the hub PST.** This is the failure to bet on: a POP3
  account delivers to the profile's DEFAULT store unless told otherwise, the arrival assertions
  read the hub store's Inbox, and a misrouted delivery looks exactly like a sink that is not
  working - a 180-second timeout with no diagnostic pointing anywhere useful.
* **"Leave a copy of messages on the server" OFF.** Outlook then issues `DELE`, the sink drains
  to zero after each test, and the whole class of stale-UIDL bugs becomes impossible.
* Set `Send Mail Immediately` to `1` under `HKCU\Software\Microsoft\Office\16.0\Outlook\
  Options\Mail`. A `0` there is the documented cause of mail sitting in the Outbox until
  somebody presses F9.

Do **not** try to shorten Outlook's send/receive interval by registry. The interval lives in a
binary `.srs` file, no Microsoft-documented value for it was found, and the suite does not need
one: `LiveInboxArrival` re-issues `NameSpace.SendAndReceive(false)` while it waits.

**Prove the sink once, by hand, before wiring any test to it.** Send a self-addressed mail with
an attachment and a body containing a line that starts with a period, confirm it arrives in the
hub Inbox intact, restart the smtp4dev service, and confirm nothing re-downloads. If that
passes, the one real objection to smtp4dev - that its POP3 side is much less exercised than its
SMTP side - is retired.

### 2.8b The identity account - a BUILD step, not a TODO

**Add a SECOND mail account, give it its own delivery PST, and leave that PST out of
`bystanderStoreDisplayNames`.** That is the whole of the `IdentityAccount` capability: a non-hub
primary the write allowlist grants an identity draft in. Section 1.3's three-store layout is the
floor and deliberately has none - both of its non-hub stores are declared bystanders - so a
machine built only to the floor runs the two identity tests, iterates nothing, and announces that
it proved nothing. Build this and they assert instead.

**Why it is here rather than on a list for later.** The guests are being created from scratch,
where this costs minutes: one more account in the same wizard, one more PST, one more name in a
settings file. Retrofitting it into a machine that is already built is a profile edit, an Outlook
restart, a settings change and a checkpoint that no longer describes the machine. The suite
already tells the reader to do exactly this - `IdentityDraftCoverageReport` ends its announcement
with "give this machine a second mail account and leave it out of `bystanderStoreDisplayNames`" -
and this section is that instruction put where a rebuilder will act on it, instead of only where a
run mentions it afterwards.

**What it has to be, read off the two tests rather than invented.** They are
`LiveDraftTests.IdentityDrafts_BusinessAccounts_...` and
`LiveDraftOptionsTests.NewDraft_BusinessAccounts_...`, and both iterate the granted stores by
DISPLAY NAME and then hand that same string to `NewDraft` as an address:

* **A real Outlook account, not a bare PST.** The first test asserts `AccountResolved` and reads
  `SendUsingAccountSmtp` back off the saved draft, so Outlook must have an `Account` object to pin.
  Add it exactly as section 2.8 adds the dummy account - POP3 against the same sink, which is
  catch-all, so a new address needs no provisioning anywhere.
* **The store display name IS the address**, as it is for the hub. `identity@vm.invalid` is the
  obvious choice, and `.invalid` keeps it unroutable by RFC 2606. Section 8 item 2 used to gate
  this one as well; it no longer does - the `@` is accepted, measured, and that item carries the
  evidence.
* **Its own delivery store.** "Deliver new messages to" must be this account's own PST: the draft
  is asserted to land in *that account's own Drafts folder*. Not the hub's, and never a corpus -
  section 2.6 says why an account delivering into a corpus store locks the generator out of it.
* **A signature configured on the account.** The second test asserts `SignatureInjected` and then
  reads the injected HTML back, so the account needs a signature assigned for New mail in
  Outlook's own settings. It is ordinary user data on this machine and is **not** one of the
  `OutlookAI-McpTest-` signatures the suite creates and deletes; the SHA-256 signature snapshot
  requires it to come back bit-identical, which it will, because nothing in the suite writes to it.
* **Declared in `expectedStoreDisplayNames`, and named NOWHERE in `bystanderStoreDisplayNames`.**
  The first is what censuses it; the second is the entire point. A declared bystander is refused
  every write, and that refusal is what empties the identity list. This is the one store on the
  machine whose absence from the bystander list is deliberate rather than an oversight, so record
  that fact beside the settings file - the half-declared state elsewhere is a refusal precisely
  because nobody can tell those two apart by looking.

That makes the machine **four stores**: hub, corpus, bystander, identity. The tripwire's bystander
is still a separate store and still the only one the guard can decide on - the identity store is
written to, so it can never do that job.

In section 2.10's settings file the delta is one name, in one list:

```
"expectedStoreDisplayNames":  [ "test@vm.invalid", "Corpus A", "OutlookAI Bystander", "identity@vm.invalid" ],
"bystanderStoreDisplayNames": [ "OutlookAI Bystander", "Corpus A" ],
```

**The example in 2.10 is deliberately left at the three-store floor, and so is section 1.3's
table.** That floor is the shape of the committed `Testbed/live-test-settings.example.json`, which
`T1/IdentityDraftCoverageTests` reads in order to pin what a machine WITHOUT an identity account
does - it is the machine the announcement exists for. Build the fourth store; do not edit the
example to match it.

### 2.9 The seed corpus

Corpus work runs under the **no-accounts profile** (section 1.2). The generator is
`McpServer/OutlookAI.RemediationTools`, a plain `Exe`; it is not installed or aliased, so
invoke it by project path or by its built exe.

```
:: 0. the expectation sheet. Pure - no Outlook, runnable anywhere, including the host.
dotnet run --project McpServer/OutlookAI.RemediationTools/OutlookAI.RemediationTools.csproj -- \
  corpus-plan --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000

:: 1. probe placement and dates. Creates and deletes a handful of throwaway items.
dotnet run --project <as above> -- corpus-probe \
  --store "Corpus A" --allow-store "Corpus A" \
  --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000

:: 2. dry run. NB it runs NEITHER probe, because both create items.
dotnet run --project <as above> -- corpus-build \
  --store "Corpus A" --allow-store "Corpus A" \
  --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000 \
  --manifest D:\corpus\vm1.jsonl

:: 3. build. Resumable and idempotent: it builds the ordinals the manifest lacks.
dotnet run --project <as above> -- corpus-build ... --progress-every 250 --execute

:: 4. check what actually landed. Read-only; the build runs this on itself.
dotnet run --project <as above> -- corpus-census \
  --store "Corpus A" --allow-store "Corpus A" \
  --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000 \
  --manifest D:\corpus\vm1.jsonl
```

Before letting a build proceed, confirm in its own output that the store line and the profile
line both say accepted, that `profile accounts: 0`, and that the placement probe and the date
probe each named a **verified** rung. A build that had to be talked past either of those guards
is a build whose measurements mean something other than what they say.

**On a guest, `Testbed/guest/Build-Corpus.ps1` runs all of this and refuses first** unless three
preconditions hold: the default profile has no **mail** account, Outlook is already running and
warm (180 s), and `ImportPRF` is not set (section 2.5a - fail-closed, and `-SkipPreflight` does not
skip it). Its preflight first met a guest on OAI-UNINDEXED on 2026-09-24, from CP-05, and refused
the correct setup: it counted every entry under the profile's account-manager key, and Outlook
lists the profile's data file and address book there too, so the account-less corpus profile
read as "2 account entries". It now counts mail accounts only. With that fixed, and
`CorpusProfile` default with Outlook up 208 s, it ran the whole of the above against that guest's
existing corpus and exited 0: both probes verified (`DraftsThenMoveWithSentFlag`,
`PropertyAccessorDates`), `Build finished: created 0, already present 20,000, failed 0`, and a
census that found every ordinal exactly once, in the folder the plan names. With the tier profile
default and Outlook closed it refused with both reasons and wrote nothing; with `ImportPRF` set it
gave the third.

Repeat for Corpus B under the other Windows account, with **a different `--corpus-id` and a
different manifest path**. Whether the two corpora should share a seed and anchor is not
settled; sharing them makes the two stores directly comparable, which is probably what you
want.

**One corpus id per guest, and the ids are assigned, not invented at the keyboard**:
`vm-indexed` on `OutlookAI-Indexed`, `vm-unindexed` on `OutlookAI-Unindexed`, and `vm2` stays
`vm2` on the outgoing guest. The manifest is `corpus-<corpusId>.jsonl`, so `--corpus-id` and
`--manifest` change together, always. `Testbed/host/Copy-FromGuest.ps1` pulls every guest's
manifest into one shared directory - on purpose, so a reused id still collides where a human can
see it - so two guests sharing an id means the second pull replaces the first's manifest, and that
manifest is the only allowlist `corpus-teardown` will delete from a real store. (`measure.jsonl`
and the logs go to a per-guest subdirectory instead; the two halves of the pull are settled
differently and `Testbed/README.md` section 3 says why.) The convention, the reasoning and the
backstop that refuses such an overwrite are in `Testbed/README.md` section 3.

**The manifest is the only thing that can tear the corpus down**, and it is also what the
freshness check reads. Copy it somewhere outside the guest. Losing it means `corpus-reindex`
and a human inspecting the result.

Budget roughly 400 MB of body text for 40,000 items before Outlook's own overhead, and about
12 minutes at ~50 items/s when the chosen placement rung needs no move. A rung that moves each
item writes it twice; budget double.

**Checkpoint `CP-06-PRE-CORPUS` before the build and a fresh one after it.** Snapshot after the
corpus exists and before any measurement, so a measurement can be repeated against the same
population.

### 2.10 The settings files

Create `McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json`. It is
gitignored and must stay that way: it names real stores and this repository is public. Without
it the whole live tier refuses to start.

```json
{
  "machineProfile": "Portable",
  "testHubStoreDisplayName": "test@vm.invalid",
  "expectedStoreDisplayNames": [ "test@vm.invalid", "Corpus A", "OutlookAI Bystander" ],
  "bystanderStoreDisplayNames": [ "OutlookAI Bystander", "Corpus A" ],
  "expectedDelegateStoreDisplayNames": [],
  "corpus": {
    "storeDisplayName": "Corpus A",
    "manifestPath": "D:\\corpus\\vm1.jsonl",
    "corpusId": "vm1",
    "seed": 4242,
    "anchorUtc": "2026-08-01T00:00:00Z",
    "itemCount": 40000,
    "windowDays": [ 7, 30, 60 ]
  },
  "mailSink": {
    "submitHost": "127.0.0.1",
    "submitPort": 25,
    "retrieveHost": "127.0.0.1",
    "retrievePort": 110
  }
}
```

| Field | What it is | Required |
| --- | --- | --- |
| `machineProfile` | `Production` or `Portable`. Absent means `Production`, so an older settings file keeps the validation it was written under. Accepted as a string or a number. | no |
| `testHubStoreDisplayName` | Display name of the store the suite may write to, exactly as Outlook shows it. Doubles as an SMTP address. | **yes** |
| `expectedStoreDisplayNames` | Every store the count tripwire watches. Include the hub. | **yes** |
| `bystanderStoreDisplayNames` | Stores the write allowlist must **refuse**. Every name here must also be in `expectedStoreDisplayNames`; naming a store in only one of the two refuses the tier. Both corpus stores belong here. | no |
| `expectedDelegateStoreDisplayNames` | Delegate/shared mailboxes. Watched, never written, folder hierarchy allowed to come and go. Empty here. | no |
| `probeTerm` | A word proven to hit this machine's search index. | Production only |
| `subjectOnlyProbe` | Coordinates of a population whose term is in the subject and not the body. Four fields, all or none. | Production only |
| `delegateNestedFolderProbe` | A delegate folder Outlook nests and the index publishes flat. | never |
| `corpus` | Where the measurement corpus is and what it was generated from, so the tier can prove it is still measurable. Six fields plus optional `windowDays`. | no, all or none |
| `mailSink` | Loopback submission and retrieval endpoints. **Absent is AMBIGUOUS and that is a known hazard: it means EITHER this machine has real transport OR it has none at all.** The testbed guests are the second case (decided 2026-09-15, no sink), the maintainer's machine is the first, and the settings file cannot currently tell them apart - so a guest with no transport reads exactly like a machine with perfect transport. See the no-sink decision in section 1.4. | no, all or none |

A block that is present must be **complete**: three fields out of four reads as configured and
behaves as absent, which is the exact silence these checks exist to remove.

`windowDays` is how the machine declares which measurement windows it actually asks about. Left
empty it means all of them, including the one-day window - which forces a rebuild every day.
Name the windows your tests use.

The same file is read by the remediation console's `audit`/`refile`/`purge`/`dedupe` verbs,
which require the hub to appear in `expectedStoreDisplayNames`. The `corpus-*` verbs do not read
it at all; they take everything on the command line.

### 2.11 Checkpoints

Names in use: `CP-01-WIN-CLEAN`, `CP-02-INSTALLER-STAGED`, `CP-03-OUTLOOKAI-INSTALLED`,
`CP-04-OFFICE-GOLD`, `CP-05-ADDIN-TRUSTED`, `CP-06-PRE-CORPUS`. Take another after the corpus
and another after the sink and dummy account exist, because those two are the steps most likely
to need redoing.

---

## 3. Keeping the corpus usable

**A corpus goes quietly out of date, and this is the failure mode to understand before
anything else.** The corpus is generated against a FIXED anchor. Every test asking about "the
last N days" selects against the CLOCK. Six weeks after generation a seven-day window selects
nothing at all - and every test asking about that window still **passes**, because selecting
nothing is a valid answer about an empty window. Nothing goes red. The suite stops measuring
and keeps reporting that it measured.

Two things now prevent that.

**The check.** `corpus-verify` is pure: no Outlook, no store, runnable on the host.

```
dotnet run --project McpServer/OutlookAI.RemediationTools/OutlookAI.RemediationTools.csproj -- \
  corpus-verify --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000 \
  --manifest D:\corpus\vm1.jsonl --window 7 --window 60
```

It derives the shift the store already carries from the manifest, counts what each window
selects now against what it selected at the anchor, and exits non-zero when any window under
test has emptied. The live tier runs the same check at fixture time from the `corpus` settings
block, fail-closed, beside the count tripwire.

**The repair is to REBUILD, not to re-anchor. `corpus-reanchor` is retired and refuses to
run.** It prints a notice and exits non-zero before it even checks its arguments, so a typo
still gets told the command is gone rather than being quietly interpreted.

**Why it is retired, since a shift-the-dates verb sounds obviously cheaper than regenerating
40,000 items.** It destroyed a corpus. Pointed at an existing 20,000-item population it wrote
**wall-clock** dates onto every item and reported `failed 0` while doing it - so the corpus was
silently flattened to a single instant and the tool said the run was clean. It was recovered
only because a checkpoint existed. The defect was fixed and pinned by tests, but the shape of
the thing does not change: it is a write path across the whole corpus whose failure mode is
invisible, in service of an outcome a rebuild reaches with no write path at all.

**And the objection that a rebuild gives you a different population is simply false.** The plan
is a pure function of `(seed, ordinal, field)` - there is no clock anywhere in it. Same seed,
same corpus id, same count, same body text, same subjects, same recipients, same distribution.
**Only the anchor moves**, which is exactly and solely what a re-anchor was for.

So the maintenance path is:

```
:: 1. remove the old population - refuses anything without BOTH the EntryID and the ordinal tag
dotnet run --project <as above> -- corpus-teardown ... --execute

:: 2. rebuild it against today. Resumable and idempotent, as in section 2.9.
dotnet run --project <as above> -- corpus-build \
  --store "Corpus A" --allow-store "Corpus A" \
  --corpus-id vm-indexed --seed 7777 --anchor <today> --count 20000 \
  --manifest D:\corpus\corpus-vm-indexed.jsonl --progress-every 250 --execute
```

**The manifest file name is `corpus-<corpusId>.jsonl` and that is not cosmetic.** An earlier
revision of this section said to write `vm1-<today>.jsonl`, which **`Copy-FromGuest.ps1` would
never have collected** - its default include is `corpus-*.jsonl`, so the rebuilt manifest would
have stayed on the guest and nobody would have been told. The manifest is the EntryID allowlist
`corpus-teardown` requires and the file `corpus-verify` reads, so a manifest that silently fails
to come off the guest is the worst of the available outcomes. See `Testbed/testbed.json`'s
`corpusIdConvention` for the id rules.

**Do not put the date in the file name to make it "new".** A fresh anchor is a fresh corpus, but
what distinguishes corpora is the **id**, not the date - and the id is what appears in every
subject, in the teardown match and in the comparison scope of every measurement. If you want the
old manifest kept, move it aside yourself; the rebuild writes the canonical name.

**Faster still, and the reason the checkpoints exist:** a corpus lives in its own local `.pst`,
which the store guard already proves. Deleting that file removes the population completely and
with certainty - no predicate, no allowlist, no partial run - and step 2 rebuilds it.

**Rebuild after every checkpoint restore.** A restored checkpoint puts the corpus back where it
was on the day it was taken, which is by definition older than today.

**A rebuild replaces every item, so the index will re-crawl Corpus A.** Let it settle before
taking an index measurement. This is the one real cost of rebuilding over shifting dates, and
it applied to re-anchoring too, which also touched every item.

**The one remaining use of the old verb** is `--diagnose-write-path`, named after the only
thing it is still good for: establishing whether date writes land on an existing item on a
given machine. It prints the retirement notice as well.

---

## 4. Running the tier

```
dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
  --filter "Category=Live&Requires!=DelegateStore"
```

That filter IS the VM bucket, spelled out: everything live except the tests naming a capability
this machine cannot be given. There is no separate "which bucket" trait to keep in step with it -
see section 5.

**Do NOT narrow this filter to quieten a machine that lacks a capability - and specifically, do
not add `&Requires!=IdentityAccount`.** Two tests carry that trait,
`LiveDraftTests.IdentityDrafts_BusinessAccounts_...` and
`LiveDraftOptionsTests.NewDraft_BusinessAccounts_...`. On a machine with no identity account they
still run: they iterate an empty list and print `PROVED NOTHING:` naming the reason and every
store the write allowlist withheld. **That line is the only record anywhere that the identity path
is unverified on this machine.** Deselect the two tests and the line goes with them, and the run reports a clean
pass over a gap nobody is told about - which is the exact vacuous-green failure both the trait and
the announcement were added to stop, cancelling each other out.

**The trait is for a machine that HAS the account**, so such a machine can select those tests and
mean it. It is not a way to silence one that does not. A machine without the account leaves the
filter exactly as written above and reads the `PROVED NOTHING:` line in the output. Section 2.8b
is how to stop needing that line at all - by building the account, which is the only thing that
turns those two tests from an announcement into a verification.

To run one class:

```
dotnet test <csproj> --filter "Category=Live&FullyQualifiedName~LiveTableSortProbeTests"
```

**A filtered run is fully guarded.** It takes the census, runs the health preflight, checks
corpus freshness and sink reachability, and verifies at the end of whichever collection the
filter left last. That was not true before 2026-08-19: verification lived in one collection's
teardown, so any run that did not include that collection paid for a baseline and threw it
away.

To see the sets without running anything - `--list-tests` discovers and does not execute, so it
is safe against any mailbox:

```
dotnet test <csproj> --list-tests --filter "Category=Live"                          # 127
dotnet test <csproj> --list-tests --filter "Category=Live&Requires!=DelegateStore"  # 121
dotnet test <csproj> --list-tests --filter "Category=Live&Requires=DelegateStore"   # 6
```

Treat those numbers as "what they were when this was written" - measured 2026-08-24. The traits
are the authority; the counts in a document drift. `Requires!=X` means "no value of `Requires` on
this test equals X", which is what makes a multi-valued trait usable as an exclusion.

---

## 5. Which tests are in which bucket, and how to find out

The classification is **two traits on the test itself**, not a list in a document that can drift -
and not three traits either. It used to be three, and the third one was the problem.

* **`Category=Live`** means "this test needs a mailbox". It is the CI gate, and it survives the
  existence of this VM because CI runs on a GitHub Windows runner with no Outlook at all.
* **`Requires`** says *what of a machine* the test needs, from one closed vocabulary, declared
  **per method**. Nothing else is declared: which bucket a test is in is a question asked of
  `Requires` at filter time.

**The three buckets, all computed:**

| Bucket | How it is selected | Size |
| --- | --- | --- |
| CI | `--filter "Category!=Live"` | 2,226 cases |
| VM | `--filter "Category=Live&Requires!=DelegateStore"` | 121 |
| production-only | `--filter "Category=Live&Requires=DelegateStore"` | 6 |

**The vocabulary, all eleven values.** Ten of them this VM can be given; one it cannot.

| Capability | What the machine must have |
| --- | --- |
| `OutlookInstance` | an Outlook to attach to, and nothing more specific. The floor: a live test that needs nothing else says this rather than saying nothing |
| `InteractiveDesktop` | a real desktop session - Outlook windows and screenshots cannot be driven from session 0. Declared only by tests that PUT SOMETHING ON SCREEN |
| `AddInRegistry` | the add-in installed and run once, so its tuning values exist |
| `SearchIndex` | a populated Windows Search index - Corpus A |
| `MailAccount` | a mail account rather than a bare PST - the dummy account |
| `Transport` | mail that actually goes out and comes back - the local sink |
| `MultipleStores` | more than one store mounted - all three |
| `IdentityAccount` | a second mail account the write allowlist grants an identity draft in - a non-hub primary left OUT of `bystanderStoreDisplayNames`. **The three-store floor in section 1.3 does not have one** and the tests naming it then prove nothing and say so; **section 2.8b builds one**, which is how a machine stops needing that announcement. Both states are real: 1.3 is the minimum that runs, 2.8b is what a complete guest has |
| `SmallHubStore` | a hub small enough that a paging assertion means something |
| `ProbePopulation` | the hand-curated population named in the settings file |
| **`DelegateStore`** | **a delegate/shared mailbox. The one capability no test machine can be given** |

`.github/scripts/check-pinned-constants.ps1` fails the build if any of those eleven names stops
appearing in this file, so the table above is load-bearing text and not decoration.

**Why `DelegateStore` is the only production-only capability.** A delegate/shared mailbox is
indexed with its folder hierarchy FLATTENED - an item in the delegate's `Archive/SomeFolder` is
published as `<host>/1/<delegate>/SomeFolder`, every intermediate folder dropped. A local PST
cannot be made to have that property, and faking it would manufacture confidence in the one area
this product has most often been surprised by. The six capabilities that used to sit beside it
(`SearchIndex`, `MailAccount`, `Transport`, `MultipleStores`, `SmallHubStore`, `ProbePopulation`)
stopped being production-only the moment this machine's shape was settled: sections 1 and 2 build
every one of them.

**The third axis is gone, and this is what it was.** A `LiveTier` trait held `Portable` or
`ProfileBound` and had to be kept in agreement with `Requires` by hand - a computed value
maintained manually, which is the exact drift `T1/LiveTierInventoryTests` exists to prevent. It
was paired with CLASS-level `Requires`, so a class read as the union of everything any one of its
methods needed. Between them they reported **96 tests that could not leave the maintainer's
machine**. Re-read method by method, the real floor is **six** - the six the production-only
filter selects. `LiveTierInventoryTests` now refuses the retired trait outright and refuses a
class-level `Requires`, so neither can come back quietly.

**The tier-3 correction, and the two mechanisms that hold it.** The T3 stdio classes spawn the
real server, which spawns a COM host, which attaches to whatever Outlook is on the machine - so
tests calling `outlook_health`, `list_accounts` or `search` were reaching a real mailbox from a
run filtered `Category!=Live`. They are now `ComHostSupervisionLiveTests`,
`OutlookAvailabilityLiveTests` and `OutlookHealthLiveToolShapeTests`, needing only
`OutlookInstance`: what they need is an Outlook, not *this* Outlook. Two mechanisms hold that,
both described in `McpServer/README.md`. `McpStdioClient` refuses to send a `tools/call` for
`outlook_health`, `list_accounts` or `list_folders` unless the test hands it a contact token, and
`LiveTierInventoryTests.EveryStdioTestReachingOutlook_DeclaresIt` reads that token back out of
the compiled IL, so a new method in an old class is caught as well as a new class. That pin now
also catches the opposite error - a live class that names one of those tools and *forgets* the
token, which throws on its first call, in a tier no CI run ever executes. Three classes were in
exactly that state.

`T1/LiveTierInventoryTests` enforces all of it in CI, together with the rule that every live
class sits in a registered collection.

---

## 6. What the guards do, and what to check afterwards

Six guards arm themselves; none needs remembering.

1. **Health preflight** (`LiveOutlookPreflight`). Asks Windows whether Outlook's UI thread is
   servicing its message queue before any COM call. Refuses the tier in milliseconds when it is
   not. Exists because a wedged Outlook once turned a live run into a 22-minute hang, and an
   aborted run skips its cleanup - which is how tagged items were left in a real mailbox.
2. **Store-count tripwire** (`LiveStoreCountTripwire`). Censuses every watched store before the
   first live collection and after the last. Fail-closed: no census, no live tier.
3. **Corpus freshness** (`LiveCorpusFreshness`). Refuses the tier when a measurement window the
   corpus is meant to fill now selects nothing. Reads the manifest, never the mailbox, so it can
   run before Outlook is started. Silent when the settings declare no corpus.
4. **Mail sink reachability** (`LiveMailSink`). TCP-probes both sink endpoints before anything
   is sent, and refuses when the profile's Outbox is not already empty - mail left queued by an
   earlier run is indistinguishable at teardown from mail this run failed to clean up. Silent
   when the settings declare no sink.
5. **Write allowlist** (`StoreWriteAllowlist`). A write aimed outside the hub throws instead of
   running.
6. **Signature snapshot** (`SignatureDirectorySnapshot`). SHA-256 before and after; the user's
   real signatures must be bit-identical.

**What to read in the output.**

* `[tripwire] live-test settings: machineProfile=..., stores=N, ..., corpus=..., mailSink=...` -
  the first line. If it names the wrong machine's settings, stop there.
* `[tripwire] watch soundness: N declared bystander(s), M store(s) this census can fail on,
  K watched store(s) the suite may still write to` - printed straight after the settings line.
  **`M=0` is the machine-readable form of "the guard runs and proves nothing"** (section 7), and
  it says so in the same words rather than leaving you to infer it from a zero elsewhere.
* `[corpus] Freshness: OK - anchor ... Windows now/at-anchor: 7d=3,180/3,180, ...` The `now`
  side is what the tests will actually see. A window at zero is a refusal, not a warning.
* `[sink] submission 127.0.0.1:25 and retrieval 127.0.0.1:110 both answering.`
* `[tripwire] baseline: 3 stores, 21 mail folders, identified 8 folder(s)/312 item(s), 431 ms.`
  The identified count is what was walked ITEM BY ITEM, and it is the number that says how much
  of the guard is live. **Zero identified items means the guard can only see counts**, which
  means the bystander store is empty or the hub is the only store.
* `[tripwire] post-run census in T ms (identified ...); 0 failure(s), K note(s).` Notes are
  benign; failures throw.
* `PROVED NOTHING:` - a test that ran but found no population to test. On a machine declaring
  `machineProfile: "Portable"`
  that is expected for the handful of tests that discover their own population; on a Production
  machine it throws instead.

**And check that a verification happened at all.** A run that prints a `baseline` line and no
`post-run census` line did not compare anything.

**Afterwards, every time:** zero tagged artifacts across Drafts, Inbox, Sent Items, **Outbox**,
Deleted Items and the Sync Issues subtree; `0 failure(s)` from the post-run census; the
signature directory bit-identical; the Outbox empty.

---

## 7. The count tripwire's first run: what to expect

The tripwire was rewritten on 2026-08-19 and **has never completed a baseline-and-verify pair**.
It has run: on 2026-08-20 it refused the live tier outright when a per-store census on the
maintainer's real profile exceeded its STA budget, which is the guard behaving correctly and is
also why its census now reads a table instead of opening every message. What follows is
predicted from the code, so treat it as something to check.

**With one PST that is also the hub.** `PlanFor` gives the hub a count-only plan and `Evaluate`
exempts it: every mail folder is counted, `0 folder(s), 0 item(s) identified`, and no failure is
reachable. The guard runs and proves nothing. This is the configuration to avoid, and it is why
section 1.3 insists on a bystander.

**With the three-store layout.** Everything the guard decides happens on the bystander. Its
folders are inside both budgets, so each is walked item by item - the rewritten half of the
guard, exercised for real. Cost: a table row count per folder plus four late-bound property
reads per identified item, so well under a second against local PSTs. On the maintainer's
five-store profile it is a different number entirely and has never been measured.

**Two things to watch.**

* **The Junk folder may not be marked self-pruning.** The census marks volatile folders by
  asking the store for its default Deleted Items, Junk and sync-issue folders. A PST may refuse
  `GetDefaultFolder` for Junk, in which case a generator-made "Junk Email" folder is treated as
  ordinary and a decrease in it would FAIL rather than be noted. Nothing prunes it here, so this
  should stay theoretical - but if the tripwire ever fires on Junk, this is why.
* **A move whose destination was only counted cannot be exonerated.** The census can prove an
  item was filed rather than deleted only when BOTH folders were walked item by item. An item
  moved from a small folder into one above the budget is reported as removed.

**Correction: there is no retry ladder on this machine (2026-08-24).** An earlier version of
this section predicted that a suspected loss would be re-censused twice and the run repeated
before anything failed. That is not what a `Portable` profile does. `TripwireRetryPolicy.None`
applies here, so **a suspected loss fails on the FIRST reading**, with `NO RE-CENSUS IS
CONFIGURED` rather than a ladder. The re-census-then-re-run policy exists, but it is a
`Production` behaviour; on the VM the first reading is the verdict.

---

## 8. What a rebuilder still has to establish

This list is a deliverable in its own right: the point of naming a gap is that a rebuilder
should not have to discover it is missing.

**Closed items are kept, with the answer, rather than deleted.** A numbered gap that vanishes
reads as one that was never there, and the answer is usually the more useful half. Items 6, 7,
14, 15, 18 and 20 are now closed and say where the answer lives; the rest are genuinely
unrecorded or unverified.

**Verify before building anything else**

1. ~~That one Windows account's Outlook profile can be excluded from the index while another's
   is not~~ - **NO LONGER LOAD-BEARING, 2026-09-15.** This was the riskiest item on the list: the
   whole three-store layout rested on it, and it was *derived from how the `mapi16://{SID}/` scope
   is addressed, never measured*. It is now moot, because **the fallback this entry itself named
   has been taken** - "two VMs, and the store layout collapses to one corpus per machine". Index
   state is a property of the **machine** now, which Indexing Options controls unambiguously.
   **Nobody has to verify it, and nobody has to build around it.** See 1.1a. Note the distinction:
   the assumption was not disproved, we simply stopped depending on it, and the two-guest build
   was adopted for an unrelated reason.
2. ~~That Outlook accepts `@` in a store display name~~ - **ANSWERED 2026-09-15: YES, MEASURED.**
   This was called out as gating the whole draft family, and no documentation or community source
   stated a restriction either way. It came out as a side effect of the PRF spike rather than from
   the probe written for it: the profile Outlook built from `Testbed/guest/tier-profile.prf` on
   Office LTSC 2024 build 16.0.17932 carries a `MSUPST MS` service whose `Account Name` reads
   literally **`tier@vm.invalid`**. The `@` is accepted, stored and read back unchanged.

   **The second caveat this item used to carry is now closed as well**, and by a second
   independent measurement. It read: *"this is the name on the profile's PST service; whether
   `Store.DisplayName` reports the same string over COM has not been read back yet, and that is
   the property the live tier keys on."* It has been read back, twice:
   - from the PRF-built profile, `Store.DisplayName` reports `tier@vm.invalid` **over COM**;
   - `Testbed/guest/Rename-OutlookStore.ps1` renamed a store's **root folder** from `Outlook Data
     File` to `tier@vm.invalid` and `Store.DisplayName` FOLLOWED - which also settles a question
     four research passes could not answer in any source - and the account afterwards still
     reported `SmtpAddress='tier@vm.invalid'` and `DeliveryStore='tier@vm.invalid'`, so the rename
     does not break the binding.

   **One caveat stands.** The name was set **at profile-creation time through the PRF**, or by a
   scripted folder rename - never typed into Data File Properties - so a validation rule in *that
   dialog* is still untested. It is irrelevant while stores are always created by script, which is
   the plan.

   **THIS ITEM IS THE SINGLE PLACE THIS ANSWER LIVES.** Sections 2.6 and 2.8b used to restate the
   question as open and now point here instead. A document that contradicts itself is worse than
   one that is merely out of date, so if this ever changes, change it here.
3. **Whether smtp4dev's POP3 side maps an arbitrary `USER` to the catch-all mailbox**, or
   whether the username must match a configured mailbox name. If the latter, add an explicit
   `Mailboxes` entry with `Recipients: "*"` and use its name as the POP3 username.
4. **That the corpus store's `IsDataFileStore` stays true once the profile has an account**
   (section 2.6). If it flips, the generator is locked out of that store permanently.

**Not recorded anywhere**

5. Hyper-V generation, Secure Boot, TPM, vCPU, RAM, disk size, checkpoint type (production
   checkpoints use VSS and behave differently with Outlook mid-run).
6. ~~Windows edition, build, ISO, licensing~~ - **RECORDED 2026-08-25 in `Testbed/MEDIA.md`**:
   Windows 11 Pro from the consumer multi-edition English International image, unactivated
   (which watermarks and nags but, unlike the Enterprise evaluation it replaces, **never
   expires**). The locale the guests are built to is recorded there as a measured table, and it
   deliberately matches the maintainer's own machine rather than a neutral en-US default,
   because that configuration is where the userbase sits: display en-GB, formats **nl-NL**,
   system locale en-US, keyboard **US-International**, `W. Europe Standard Time`. **Expect
   `4.000,50` for four thousand and `25-8-2026` for a date, on purpose.**
   Still unrecorded: computer name, and Defender exclusions - an indexer, a 400 MB PST and
   real-time AV interact, and nobody has measured how much.
7. ~~Office version, channel, bitness, install method~~ - **RECORDED in `Testbed/MEDIA.md`**:
   Office Deployment Tool with `ProPlus2024Volume` on `PerpetualVL2024`, 64-bit, and
   **`ExcludeApp OutlookForWindows`, which is load-bearing** - it suppresses the new Outlook,
   which offers no COM object model. The guest build was read on 2026-09-15 and is
   16.0.17932.20884. Office's out-of-box grace is **30 days, not 90** - but **that clock turned
   out not to matter, and the monthly rebuild cadence it justified was RETIRED on 2026-09-15**.
   **Past grace, Office does NOT lose functionality**: the documented state is "Unlicensed
   notification" - nags and a red title bar - the guest was measured at `LicenseStatus=5` with
   every COM read this project uses still working, and **a cold COM start completed in 3.7 s with
   no dialog**. The grace clock also exists only because the guest is a KMS client; the
   maintainer's Office is MAK-activated with no clock at all. Rebuilds now have two triggers,
   neither a calendar: `corpus-verify` refusing, and a rebuild before a release. Still
   unrecorded: how the first-run wizard is suppressed. Still unmeasured, and stated as a gap in
   evidence rather than a known risk: the **write** path, which mailbox-safety rule 1 keeps out
   of a probe's reach.
8. The two Windows account names and their roles; whether both need a clone and an SDK; whether
   checkpoints must be taken with both logged on.
9. Outlook profile names, how they are created, which is default, and how the switch between the
   no-accounts profile and the tier profile is automated.
   **HALF-ANSWERED, and the other half is CLOSED-NEGATIVE.** Profiles, PSTs with exact display
   names, and the default-profile switch all have scripts - `Testbed/README.md` section 4b
   indexes them and `Docs/research/profile-automation-research.md` is the evidence. **The first
   route they were all built on, Extended MAPI's `IProfAdmin`, is measured broken on this Office
   build** (2026-09-16, `E_NOINTERFACE` on `IID_IProfAdmin`) and all three have been moved onto
   routes that work: a `.prf` import for the profile and its named stores, `AddStoreEx` plus a
   root-folder rename for a store going into a profile that already exists, and the registry
   `DefaultProfile` value for the switch. The default switch **has run on a guest**; the other two
   have not, so this item stays open until they have - what exists for them today is a `-SelfTest`
   over their decision logic, which is not the same claim. Section 2.5a carries the one hazard a
   rebuilder must not skip.
   **The mail account is the closed-negative half: there is no free programmatic route to creating
   one.** The object model has no `Accounts.Add` and `Account.DeliveryStore` is read-only; MAPI has
   no POP3 message service, because account administration moved behind the undocumented
   `IOlkAccountManager`; the registry has no published recipe on 16.x and DPAPI-seals its stored
   passwords per user per machine; and a `.prf` cannot carry `PROP_ACCT_DELIVERY_STORE`, which is a
   binary EntryID. `Testbed/guest/New-PopAccountPrf.ps1` is the spike that tests the last of those
   and is expected to fail. **So sections 2.8 and 2.8b keep ONE GUI pass per guest** - two accounts
   and their delivery stores - and the mitigation is a checkpoint immediately after it.
10. The scheduled-task recipe for session 1: task name, principal, working directory, argument
    line, output redirection and exit-code capture. Only "`-LogonType Interactive`" is recorded,
    and an elevated process's stdout cannot reach the caller, so output must go to a file.
11. The exact PST file paths and names for all four stores, and the mapping from file name to
    display name.
12. .NET SDK version, clone path, build configuration, and how the built server exe reaches the
    path the tier-3 tests expect.
13. How results, screenshots and logs get out of the guest, and where `ScreenCapture` writes.
14. **CLOSED, with a caveat that matters.** `Docs/v3-probes/soakfix13-probe-sweep-cost.ps1` was
    gitignored, so it lived on one machine and is gone. It has been **reconstructed in the
    repository** as `Testbed/guest/Measure-SweepCost.ps1`, written from the shipped sweep's own
    source rather than from memory, and read-only by construction. **It has never been
    executed.** Read it before trusting a number out of it, and replace its banner with what it
    actually did once it has run.
15. **CLOSED 2026-08-24.** The real parameters are `vm2 / 7777 / 2026-08-19 / 20000`, recorded
    machine-readably in `Testbed/testbed.json` together with the expected plan, the per-folder
    and per-window counts, the store path and the build cost. They are not an example: they were
    read out of the recovered manifest header and re-derived on the host with `corpus-plan`,
    which reproduced Inbox=10,912 / Sent=4,964 / Deleted=2,461 / Junk=1,663 and 7d=1,612 exactly
    - the figures `Docs/magic-numbers.md` quotes. `check-testbed-references.ps1` pins the four
    values across `testbed.json`, `Build-Corpus.ps1` and `corpus-measurement-plan.md`.
    **The `vm1 / 4242 / 2026-08-01 / 40000` set that appears in this document's command
    examples is an EXAMPLE and always was.** Do not build from it.
16. Whether Corpus A and Corpus B should share a seed and anchor, and how their manifests are
    named apart.

**Known gaps in what the corpus contains**

17. The generator writes `IPM.Note` with a subject, a body, a read state, message flags and two
    date properties. **No senders, no recipients, no attachments, no HTML, no categories, no
    flags, no subfolders, no other message classes.** A test needing any of those is in the VM
    bucket and will still fail here - because of the corpus, not because of the machine, and
    nothing in the traits says so. Widening the generator is queued work.

**Open behaviour**

18. **CLOSED 2026-08-24, in both halves.** The policy is decided *and* built - and the answer for
    this machine is that there is no ladder at all. `TripwireRetryPolicy.None` applies to a
    `Portable` profile, so a suspected loss fails on the first reading with `NO RE-CENSUS IS
    CONFIGURED`. Section 7 says so; the re-census-then-re-run behaviour is `Production` only.
19. **STILL OPEN, and narrower than it was.** `machineProfile: "Portable"` still turns "found
    nothing to test" into a pass, and there is still **no assertion-counting hook anywhere in the
    suite** - confirmed by reading, not assumed: xunit 2.9.3, no `BeforeAfterTestAttribute`, no
    analyzer that fails an assertion-free `[Fact]`. What changed is that the tests known to be
    exposed now announce it: the `RequireProductionPopulation` + `PROVED NOTHING:` idiom throws
    on a `Production` profile and prints on a `Portable` one, and the two identity tests were
    converted to it on 2026-08-25 after being found green-while-iterating-nothing.
    **Three more silently-empty iterations are known and unfixed**, listed in `TODO.md`; the
    strongest is `LiveFolderScopeTests.DelegateFirstLevelFolders_StillResolve_...`, which
    iterates `expectedDelegateStoreDisplayNames` with no guard while a sibling two methods above
    it asserts non-emptiness first - so the omission reads as an oversight, and that list is
    empty on every Portable machine including this one.
20. **CLOSED 2026-08-24 - and worth reading how, because the obvious fix was the wrong one.** The
    grant was NOT narrowed: two live tests legitimately need draft-create in a non-hub store.
    Instead a store is now **declared** a bystander in `bystanderStoreDisplayNames`, the write
    allowlist checks that list *ahead of* the identity grant, and the tripwire verifies the
    declaration. Sections 1.3 and 2.6 carry the rule. The corpus stores are declared too, which
    is what stopped the identity tests drafting into the measurement corpus.

 21. **OPEN, and deliberately named rather than folded into item 2 - a guest whose index is
    genuinely UNREACHABLE.** Corpus B is an *unindexed store on a working indexer*: the frontier
    probe runs and returns no rows, which is the shape `ResolveSweepWindows` handles through
    `IndexFrontierMissing` / `GapNoIndexFrontier` and the seven-day fallback window. A machine
    with **no catalog at all** is a different thing entirely - the probe *throws*, and the
    product goes down its "SystemIndex is unreachable" branch, which **no test in this
    repository exercises today**: `MailService.Search` does not wrap `_index.Value.Search` or
    `GetStaleness` in a catch, and all four `Unindexed*Tests` classes use a client that answers
    and merely holds nothing. Nobody has established whether `Search.CollatorDSO` throws,
    returns empty, or serves stale rows with the Windows Search service stopped; measuring it
    would mean stopping that service on the maintainer's workstation, so it has not been
    measured. **Conflating this with Corpus B is the mistake section 2.4 exists to prevent, and
    leaving it unnamed was a different one.** If it is ever built, it is a *third* guest shape,
    not a setting on the second.

---

## 9. Known limits, honestly

* **`testHubStoreDisplayName` doubles as an SMTP address**, which is why the hub PST has to be
  named after the dummy account. It is a constraint the tests impose on the machine, not a
  design anybody chose.
* **Several tests assume a tiny hub** and would break their paging assertions
  against a 20,000-item one (`Phase7LiveMcpToolShapeTests` asserts the hub holds between 2 and
  99 items; `LiveMailServiceTests.ListFolders...` asserts the hub tree fits one page). They
  carry `Requires=SmallHubStore`. This is the reason the two machines cannot share one settings
  shape.
* **The VM bucket does not prove the delegate-store paths at all**, and no test machine can:
  `Requires=DelegateStore` needs a mailbox somebody else owns. Six tests, named by the
  production-only filter in section 5.
* **The guest's SDK is PINNED to whatever the host was running when the payload was staged**,
  and nothing enforces that they stay equal. 10.0.401 was chosen for sameness rather than for any
  requirement - no `global.json` exists and CI asks only for `10.0.x` - so the two can drift the
  moment the host updates, and the first symptom would be a guest measurement that differs from a
  host one for a reason nobody is looking for. `Testbed/MEDIA.md` records the pinned version; it is
  the thing to check when host and guest disagree about something that should not depend on the
  toolchain.

* **Nobody has yet run the VM bucket end to end anywhere.** The 121 read as runnable there; that
  is not the same as having run there. The count moved from 31 to 121 by re-reading what each
  test needs method by method - no test was changed to make it fit.

* **The VM runs a different Office from the maintainer's machine, by 3,598 builds, and that is
  accepted rather than fixed.** Measured 2026-09-15: the guest is `ProPlus2024Volume` on
  `PerpetualVL2024`, build 16.0.**17932**.20884; the maintainer's machine is
  `ProPlusSPLA2021Volume` on `Production::LTSC2021`, build 16.0.**14334**.20848. **DECIDED
  2026-09-15: stay on Office 2024.**

  This is a **deliberate exception to the principle the testbed is otherwise built on**. That
  principle - stated in `Testbed/MEDIA.md` - is that the guests match the maintainer's
  configuration on purpose, because that configuration is where the userbase sits; it is what
  justified carrying the maintainer's locale over verbatim rather than building a tidy en-US box.
  The Office version does not follow it, and saying so is better than leaving the inconsistency
  for a reader to find and mistake for an oversight.

  **The consequence is the one section 2.1 already names: an Outlook build difference is the
  first thing to suspect when a live test behaves differently on the VM than on the maintainer's
  machine.** With 3,598 builds between them that suspicion is well founded, and this entry is why
  section 2.1 says it. Anything that reproduces on the VM and not on the host, or the reverse, is
  a build-difference candidate until ruled out.

* **The guest's 30-day Office grace clock is an artefact of the testbed, not something a user
  ever experiences — and it turned out to cost nothing.** The guest is a KMS client that has
  never reached a KMS host, so it runs on out-of-box grace. The maintainer's own Office is
  **MAK-activated, `LicenseStatus=1`, with no grace clock of any kind**, so this was never a
  property of the userbase - though it had previously been discussed as though it were.
  **Measured past grace on 2026-09-15: every COM read works, and a cold COM start completes in
  3.7 s with no dialog.** Nothing stops. The monthly rebuild cadence this clock justified is
  therefore **retired**, and the `TODO.md` preflight item is **closed** - both of its
  justifications were measured false. Rebuilds now trigger on `corpus-verify` refusing and on a
  release, neither of which is a calendar.

* **The clock that DID bite was the corpus, and this is the one to respect.** A three-week
  absence left it 27 days stale with its 1-day and 7-day windows selecting nothing, while the
  Office clock everyone was watching cost nothing at all. `corpus-verify` catches it fail-closed.
  The general lesson is worth more than the instance: **the schedule was attached to the visible
  clock rather than the harmful one.**
