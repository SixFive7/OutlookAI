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
>
> **WHO WRITES THE RULE, AND WHETHER AN EXCLUSION LASTS - MEASURED 2026-09-24 (Q69).** Outlook
> writes it itself - a search root and a user INCLUDE rule for `mapi16://{SID}/`, within seconds of
> starting - but **only when it runs NOT elevated**: an elevated Outlook never touches Windows
> Search, which is why neither guest ever had the rule (section 8 item 22). The testbed now writes
> the same rule, value for value, through the Crawl Scope Manager API
> (`Set-OutlookIndexingDisabled.ps1 -Enable -Execute`), and the durability half is no longer
> underived. **A user EXCLUDE rule written through the API holds**: in three runs a non-elevated
> Outlook - the one that registers the scope - ran five minutes on top of it and left it excluded,
> nothing queued and no row back, and it held through the reboots after. **The policy alone is not
> an exclusion the service knows about**: with only `PreventIndexingOutlook = 1` the service still
> reports the scope IN and purges nothing - which is why the rule is now the exclusion and the
> policy the second layer. On a guest that was never indexed the policy alone does stop a
> non-elevated Outlook registering itself - six minutes, nothing added - but it cannot undo a scope
> something else included. And rows crawled BEFORE an exclusion go only if the service
> is left alone while it purges them: a restart in that window loses the purge - every row was still
> there after a reboot (section 2.4, the Q69 block, item 3).

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

### 1.3 The stores, the two lists, and what every store holds

**Rewritten 2026-09-24** for two decisions the maintainer made that day (Q70): the generator now
builds a small, tagged **population** into the hub and the bystander (and the identity store), and
the one list that used to mean two things is **split into a watched list and an indexed list**.
Everything below describes that final shape.

| Store | Named | Watched | Indexed list | What is in it |
| --- | --- | --- | --- | --- |
| Hub | as its account's address, e.g. `tier@vm.invalid` | yes | **first**, on the indexed guest | the **hub population** (56 items, section 3b) plus whatever a run is writing |
| Bystander | as an address, e.g. `bystander@vm.invalid` | yes, and a **declared bystander** | **second**, on the indexed guest | the **bystander population** (300 items, section 3b) - never written by a test |
| Corpus A / Corpus B | anything - `Corpus A` in the examples | yes, and a **declared bystander** | **last**, on the indexed guest | the measurement corpus: at least **160,000** items on the indexed guest (`minimumItemCount` in `Testbed/testbed.json`); no minimum is recorded for the unindexed guest's |

These three are the floor, and the shape of the committed example (section 2.8b says why the table
stays there). **Section 2.8b adds a fourth store, the identity account's**: named as its address, e.g.
`identity@vm.invalid`; watched and **not** a declared bystander, which is the point of it; **not** in
the indexed list; holding the **identity population** (8 items, section 3b) plus the identity tests'
transient drafts.

**The indexed list is empty on the unindexed guest** - the exclusion is per machine (section 1.1a),
so every store there is unindexed, hub and bystander included. That is intended and costs nothing,
because the tests that need an index run on the other guest; a rebuilder should not read an empty
`index.perStore[]` row there as a fault.

#### The two lists (split 2026-09-24)

`expectedStoreDisplayNames` used to mean two things at once, and a machine could only be described
truthfully while the two happened to coincide. They are now two lists:

* **`expectedStoreDisplayNames` - the WATCHED list.** Every store the tier profile mounts. It is
  what the count tripwire censuses (with the bystander and delegate lists unioned in), what the
  identity-draft grant is drawn from, what `list_accounts` exactness counts, what `outlook_health`'s
  reachability check reads, and what the archive-resolution test walks.
* **`indexedStoreDisplayNames` - the INDEXED list.** The stores the index-tier tests measure, in
  order; each must be discoverable in the search index with mail in it. **Absent, it means exactly
  the watched list** - which is what every index test read before the split, so a settings file
  written earlier behaves as it did. Present, it may name only watched stores, never a delegate
  mailbox, and must include the hub; empty is allowed only on a Portable machine, and there every
  index test **refuses** rather than iterate nothing (`LiveTestSettings.RequireIndexedStores`).

The identity store is the reason the split had to happen: it is watched, written to and granted,
and holds almost nothing - one list could not say "watch it, but do not demand an index scope of it".
The admission check (`TripwireWatchSoundness`) does not read the indexed list at all.
`T1/StoreListSplitTests` pins both the loader rules and, from the compiled IL, which consumer reads
which list.

**The ORDER of the indexed list is read.** Several index tests read its first entry - the post-filter,
the filter shapes, the sender filter - so the hub, whose population carries attachments, unread mail
and senders, comes first. The exclude-subfolders measurement takes the first non-hub entry and needs a
mail folder with populated children, which only the bystander population has - so the bystander comes
second and the corpus last. `Testbed/host/New-LiveTestSettings.ps1` refuses any other order.

**Every indexed store except the corpus is named as an `.invalid` address.** Not tidiness: the product
finds a small store in the index only through mail addressed to it (`TryDiscoverStoreScopeByAddress`),
and its own search and `outlook_health` try that only for a name shaped like an address
(`MailService.ResolveFolderScope`, `ProbeStoreInIndex`). A 300-item bystander named `OutlookAI
Bystander` beside a 160,000-item corpus is missed by the 2000-row discovery sample, and health then
reports it missing from the index. The corpus dominates that sample and may keep a plain name. Every
population item is addressed to or sent from its store's owner for the same reason.

#### The bystander

The tripwire **exempts the hub**, because the hub is where the suite writes; a machine whose only
store is the hub gets a guard that censuses, reports zero failures, and is structurally incapable of
reporting anything else - and the tier refuses to start on it (`NO STORE THIS CENSUS WATCHES CAN
PRODUCE A FAILURE`). The bystander must therefore be a store **no test ever touches**, and it must hold
a few hundred items rather than none, because an empty store exercises the item-by-item identity path
over nothing. A corpus is the wrong shape for that: the identity budget is 500 items per folder and
3,000 per store, and a corpus is over both in every populated folder, so it falls back to bare counts.

**It is populated by the generator** (`corpus-build --population bystander`, section 3b) - 300 tagged
items in four folders, every one inside the identity budget, two of them subfolders of the Inbox. The
earlier objection that "the generator tags everything it creates, and the bystander's whole job is to
be untouched" does not hold: the tag is the CORPUS tag, which no artifact sweep can select, and no test
writes to the store either way.

#### What each list refuses, and what it does not (corrected 2026-09-24, Q77)

This section used to say that a bystander named in only one of `expectedStoreDisplayNames` and
`bystanderStoreDisplayNames` refuses the tier. **The code does not do that, and on the maintainer's
decision the documents now say what the code does:**

* **`bystanderStoreDisplayNames` is what DECLARES a bystander.** It denies the store every kind of
  write and keeps it in the census - the census watches every declared bystander **whether or not
  `expectedStoreDisplayNames` names it**, deliberately, so the declaration alone is sufficient
  (`T1/TripwireBystanderStoreTests.TheCensusWatchesEveryDeclaredBystanderEvenOneNoOtherListNames`).
* **A store in `expectedStoreDisplayNames` and not declared is inside the identity-draft grant.** That
  is the identity account's legitimate shape (section 2.8b), not a mistake. The tier refuses it only in
  the one case that matters: when no watched store is left that is both non-hub and denied every
  write, because then the census can fail on nothing.
* **On a test guest, name every mounted store in `expectedStoreDisplayNames` as well** - declared
  bystanders included. The tier would not refuse the omission, but `outlook_health`'s reachability
  check and `list_accounts` exactness read only that list, and a store missing from it is one they
  never look for. The renderer holds guests to this; the maintainer's own file is not held to it.

**Both corpus stores are declared bystanders.** They are stores no test may write to, which is exactly
what the declaration means. Before that was true, the identity tests resolved to `Corpus A` and drafted
**into the measurement corpus**.

**`Corpus B` is deliberately NOT declared in the indexed guest's settings file.** It lives on the other
guest, and a declared bystander the running profile does not mount is censused, not found, and refuses
the tier. It belongs in *that* guest's settings file.

### 1.4 A dummy account AND a loopback sink - decided 2026-09-24, reversing 2026-09-15

**DECISION (2026-09-24): the testbed guests get a mail sink - Inbucket 3.1.1 on `127.0.0.1`,
started with the guest.** The maintainer's direction is that every live test moves to the guests
**without compromising on the tests**, and that the guests stay deterministic, reproducible and
rebuildable from scratch. Thirteen live methods put mail on the wire (section 1.4a). Twelve could
have been rewritten to seed items into the PST instead - but that is a change to the tests - and
the thirteenth, the product's own two-step `send`, cannot be replaced at all. A sink runs all
thirteen as written. Asked to build one, the maintainer asked instead for "a simple ready made
open source tool... as simple as possible"; section 2.7 says why that is Inbucket, and
`Testbed/MEDIA.md` records it as the one third-party program on the guests, under a carve-out of
`CLAUDE.md`'s Dependencies rule that the maintainer confirmed.

**What the 2026-09-15 decision weighed, and what answers each point now:**

| Declined on 2026-09-15 because | Answered on 2026-09-24 by |
| --- | --- |
| A sink is a third media precondition, which the Dependencies rule forbids | the maintainer's carve-out for this one tool. It is media in `Testbed/MEDIA.md`: staged on the host by `Testbed/host/Get-MailSinkMedia.ps1`, pinned by a hash its maintainers publish, never fetched on a guest |
| smtp4dev's POP3 refuses an empty password and the `.prf` carries none, so somebody types one per guest | Inbucket's POP3 accepts any password, **including none** [SOURCE]. What is left is Outlook's half - whether it sends a PASS without prompting - and that is one guest measurement, section 2.7 |
| smtp4dev deletes at `DELE` and renumbers the mailbox mid-session, and advertises `TOP` without implementing it | Inbucket deletes at `QUIT`, keeps message numbers fixed and implements `TOP` [SOURCE]; `Testbed/guest/Install-MailSink.ps1 -Verify` asserts each one on the guest |
| a sink buys exactly one method nothing else covers | still true, and no longer the question: the decision is not to rewrite the other twelve |

`[SOURCE]` means read in Inbucket's source at tag `v3.1.1`, not yet observed; the first `-Verify`
on a guest turns each one into a measurement or a named failure.

**What this gives back.** The Outbox is a canary on the guests again. With delivery really
happening, mail that stays in the Outbox is a send-path fault, so `LiveMailSink.EnsureOutboxDrained`
and the zero-artifact sweep over folder 4 mean on a guest what they already meant on the
maintainer's machine. The no-sink decision had made that guard vacuous there; this undoes it.

**What it costs, stated plainly.** One third-party program on the guests, never in the product; a
SYSTEM scheduled task that starts it at boot; three loopback listeners (`25`, `110`, and the web
UI's `9000`, which cannot be switched off); and one question only a guest can settle - whether
Outlook, holding no stored POP3 password, logs in or prompts (section 2.7).

**NOTHING HERE HAS RUN ON A GUEST YET.** `Install-MailSink.ps1` passes its `-SelfTest` on the host
and has never been executed on a guest. Until its `-Verify` reports `SINK-READY` on one, every
statement in this section about how Inbucket behaves is read from source, not measured.

### 1.4a Why the account points at `127.0.0.1`

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

**So the account is configured for a local sink that delivers back, and since 2026-09-24 the
guests have one.** It points at `127.0.0.1` on `25` and `110`, which is exactly where
`Testbed/guest/Install-MailSink.ps1` binds Inbucket, and the tier profile needs no change of
server, port or encryption for it (section 2.7 lists what it must say). All 13 methods that put
mail on the wire run on a guest as written, and section 4's VM filter - which already selects
them - becomes correct rather than a list of tests that could only fail.

**Until a guest's sink is installed and its `-Verify` reports `SINK-READY`, do not send on that
guest.** A send with nothing listening queues in the Outbox, and the tier's Outbox check then
refuses every later run until that item is removed - through the suite's own helpers, never from a
shell (`CLAUDE.md`, mailbox-safety rule 1).

**`OutlookAI-Unindexed` reports `SINK-READY` since 2026-09-24**, the first guest to - installed,
restarted and verified again, starting with the guest (section 2.7). Section 1.4's closing paragraph,
"nothing here has run on a guest yet", predates that run. And **`SINK-READY` is not the whole of it:**
Outlook must also be able to LOG IN to the sink, and with no stored POP3 password it does not - it
prompts before it ever connects (measured, section 2.7). A guest's accounts therefore also need
`Testbed/guest/New-TierProfile.ps1 -StoreSinkPassword -Execute` before a run that sends.

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

**Checkpoint `CP-01-WIN-CLEAN` - with BOTH discs ejected first, then delete the answer ISO.** Use
`Testbed/host/New-TestbedVm.ps1 -Name <vm> -CompleteInstall -Execute`, which waits for first logon to
finish, ejects the Windows ISO and the answer disc, takes the checkpoint, confirms it holds no disc,
and only then deletes the answer ISO. The order matters because the answer ISO carries the guest
password in clear text, and a checkpoint taken with the disc still attached references the file: it
then cannot be deleted without breaking that checkpoint's restore. Measured 2026-09-24 - both current
guests' `CP-01-WIN-CLEAN` reference their answer ISO, which is why those two ISOs are kept until the
from-scratch rebuild replaces the guests (see `Testbed/README.md` section 1b).

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
* **The add-in, on BOTH guests, built from a named commit - a scripted build step since
  2026-09-24** (`Testbed/README.md` section 1, step 5b). This line used to say "install the add-in
  and let it run once", and the script-built guests never had it: nothing in the build installed it.

**Which tests need it, and exactly what they read.** Two, both through
`McpServer/OutlookAI.Core/Services/HealthReporting.cs` (`ReadTuningState`), over
`HKCU\Software\OutlookAI\Tuning`, with the value names of `Services/AddInServerContract.cs`:

| Test | Declares | Asserts | Reads |
| --- | --- | --- | --- |
| `T3/Phase7LiveMcpToolShapeTests.Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning` | `AddInRegistry` | `tuning.managed` is true | `Initialized`, a nonzero **REG_DWORD** |
| `T2/LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail` | `SearchIndex`, `MultipleStores` - **not** `AddInRegistry`, which it needs | `Tuning.Managed`, `Tuning.Enabled`, `Tuning.LastReconcileUtc` not null | `Initialized` and `Enabled` as nonzero **REG_DWORD**s, `LastReconcileUtc` as a **REG_SZ** |

**The type matters as much as the value**: `HealthReporting`'s `AsBool` accepts a boxed `int` and
nothing else, so a REG_QWORD 1 reads as not-managed. The third test that declares `AddInRegistry`,
`T2/LiveUiSearchBackendTests.FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates`, writes the
value it tests itself and does not need the add-in. **Both guests need it** because the Phase-7 test
declares nothing that keeps it off the unindexed one - and there it asserts
`index.wSearchStartMode == "automatic"`, which `Testbed/guest/Set-OutlookIndexingDisabled.ps1`
deliberately keeps true.

**Why a build from a commit, not a released installer.** The add-in and the server agree about those
values through one file compiled into both. The suite on a guest is built from a named commit
(`Testbed/host/Publish-LiveTierPayload.ps1`); pairing it with the last release's add-in would test a
contract nobody changed together. So `Testbed/host/Publish-AddInPayload.ps1` builds the add-in from a
commit too, and records the hashes of the two contract files; the guest compares them with the
suite's and reports a mismatch as `BROKEN`. A pinned release installer remains the fallback for a rebuilder with
no Visual Studio - it is what users get, and it is exactly the drift this avoids.

**A plain build of the add-in REGISTERS it on the machine that builds it**, and on the maintainer's
workstation that repoints his own Outlook at a build output. Read in the VSTO build targets: every
build runs `RegisterOfficeAddin`, which rewrites `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAI`,
and `SetInclusionListEntry`, which writes a VSTO trust entry. The workstation already carries such
trust entries for builds made under this repository's agent worktrees. The host script builds with
three guards - the registration target off the chain, the writing tasks replaced by logging
stand-ins, and a before/after snapshot of the host that fails on any trace - and three runs left the
workstation identical. Its banner is the record.

**On the guest, `Testbed/guest/Install-OutlookAIAddIn.ps1 -Execute`, through the interactive task**:

1. **The VSTO runtime** from staged media (`Testbed/MEDIA.md`), because `Installer.iss` installs
   prerequisites only when it is NOT silent - and an unattended guest can only run it silently.
2. **The product's own installer**, `/VERYSILENT`: per-user, `|vstolocal` registration, the
   slow-add-in exemption, the signing certificate into the user's TrustedPublisher store - exactly
   what a user gets.
3. **Trust, written rather than clicked.** A self-signed certificate in TrustedPublisher does not
   retire the ClickOnce trust prompt, and on an unattended guest a prompt is a hang - the hand-built
   guest's `CP-05-ADDIN-TRUSTED` was somebody clicking through it. The script writes the inclusion
   entry the prompt itself writes: `HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid>` with
   `Url` and `PublicKey`. Microsoft Learn says accepting the prompt creates an entry holding a URL
   and a public key, and Microsoft's Japan Office support blog gives its registry path and value
   names; the maintainer's workstation holds an entry of exactly that shape for this installer's own
   install path; and the runtime's own IL stores and compares entries that way - by URI, and by key
   blob. The key is read out of the installed deployment manifest and must match the installed
   certificate and the payload.
4. **`VSTO_LOGALERTS=1`**, so a load failure writes `<app>\OutlookAI.vsto.log` instead of
   vanishing.
5. **Outlook started once**, over COM, headless, in a watchdogged child job - never quit, never
   killed.
6. **Proof, not exit codes**: `LastReconcileUtc` written AFTER that start, `Initialized` and
   `Enabled` of the right type, `LoadBehavior` still 3, nothing in Outlook's disabled list, the
   add-in connected and answering a call into it, no Claude Code registration question pending (it
   would surface as a modal dialog mid-tier), and the installed build the payload's.

**It does not disturb the index exclusion or the corpora.** The add-in's tuning service
(`Services/OutlookTuningService.cs`) writes only under HKCU - Outlook's Search key (four search-box
preferences), the Cached Mode user and policy keys (Exchange sync settings), and the PST key (a
larger file-size cap). `Set-OutlookIndexingDisabled.ps1` writes only HKLM - the Windows Search
`PreventIndexingOutlook` policy and the crawl-scope rule. The two sets are disjoint, none of the four
search values decides whether a store is indexed, and nothing in the add-in's startup path touches
an item or a store. That is read from both sources; the guest script also snapshots the exclusion
state before and after its Outlook start and says if anything moved. **Order it before step 7b** so
7b's own `-Verify` certifies the exclusion with the add-in present; on a guest already past 7b, re-run
`Set-OutlookIndexingDisabled.ps1 -Verify` after it.

**Never executed on a guest yet.** Everything above that says "measured" was measured on the host.

**Checkpoint `CP-03-OUTLOOKAI-INSTALLED` once `-Execute` prints `ADDIN-READY`.** `CP-05-ADDIN-TRUSTED`
is no longer a separate manual step - the trust entry is part of the scripted install - and the name
survives only as the hand-built guest's history.

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

**WHAT ACTUALLY SETS A GUEST'S INDEX STATE, now that there is one account per guest.** The
sentence above named a GUI on a machine that is driven headlessly, and that sentence was the
entire specification of half the testbed. The step is
**`Testbed/guest/Set-OutlookIndexingDisabled.ps1`**, with `Testbed/guest/SearchCrawlScope.cs`
staged beside it, Outlook closed. On the unindexed guest **`-Execute`** writes the documented Group
Policy `PreventIndexingOutlook = 1` **and** a user EXCLUDE rule for `mapi16://{SID}/` through the
Crawl Scope Manager API - the writer Indexing Options itself uses, which runs inside the service and
so is not stopped by the registry ACL that stops an administrator. On the indexed guest
**`-Enable -Execute`** writes the inverse: policy 0, a search root and a user INCLUDE rule. Either
way the script then asks the service whether the scope is now in or out, and throws if it does
not agree. Both **leave the indexer running**, because stopping the Windows Search service produces
a machine with no search rather than a mailbox search has not been told about, and the product
takes a different, untested code path there. Section 8 item 21 keeps that distinction from
collapsing.

> **MEASURED 2026-09-24 on `OutlookAI-Indexed` (Q69) - why no guest was ever indexed, the fix, and
> the exclusion measured on a guest that is.** Evidence in `.work/aa5e-2026-09-24-q69-index-scope/`
> and the script's banner.
>
> 1. **The cause: every Outlook the testbed ever started was elevated.** An elevated Outlook does
>    not use Windows Search - no rule, no store pushed, `Store.IsInstantSearchEnabled = False` -
>    while the same Outlook started without elevation adds a root and a user INCLUDE rule for
>    `mapi16://{SID}/` within seconds and pushes every item of its profile's stores. The testbed's
>    only route into session 1, `Testbed/guest/Register-InteractiveTask.ps1`, runs at `RunLevel
>    Highest`. Controlled A/B and the evidence: section 8 item 22.
> 2. **The fix is the documented writer, in both directions.** `-Enable -Execute` adds the root and
>    the INCLUDE rule through the Crawl Scope Manager API and the service confirms it (`included=True
>    reason=USER`); the registry afterwards holds, value for value, what the non-elevated Outlook
>    wrote and what the maintainer's workstation holds. **A rule alone crawls nothing**: with the
>    scope in, the catalog stayed at zero Outlook rows for four minutes with Outlook closed and for
>    six more with Outlook running ELEVATED - the index moves only while a NON-elevated Outlook
>    runs, which is why `Testbed/README.md` section 1 step 8c starts it with
>    `Testbed/guest/Start-OutlookUnelevated.ps1`. Started that way, Outlook queued ~19,800 item
>    notifications in its first minute and the corpus was fully crawled 7.6 to 9.6 minutes after its
>    start (four runs); "finished" is the catalog's queues at 0, its status back to IDLE and the row count
>    standing still, all three, which is what `-Verify` now requires for `INDEXED` (section 8 item 22).
> 3. **The exclusion, from the indexed checkpoint - and the ORDER is the finding.** Five ways, each
>    from `CP-10-INDEXED` (20,048 Outlook rows: corpus 20,028, identity store 4, tier store 16),
>    Outlook closed, then watched with Outlook closed and - all but R - with a NON-elevated Outlook
>    running on top, the one that registers itself (`task3-variant-*.txt`):
>
>    | | what was written | the scope, per the service | the 20,048 rows | with the self-registering Outlook, then a reboot |
>    | --- | --- | --- | --- | --- |
>    | U | the user EXCLUDE rule alone | out (`USER`) at once | **purged by the indexer itself**: 0 rows 4.2 min after the rule | stayed out, 0 rows, nothing queued |
>    | R | the same rule, then a WSearch restart a second later | out (`USER`) | **nothing purged** - 20,048 after 6.5 min | (not run) |
>    | S | the script as it stood: policy 1 + rule + restart, in one go | out (`USER`) | **nothing purged** - 20,048 after 6 min | stayed out; still 20,048 after 5 min of Outlook and a reboot |
>    | P | `PreventIndexingOutlook = 1` alone, + restart | **still IN** (`USER`) | nothing purged | still in, still 20,048 |
>    | N | the script now: rule, wait for the purge, then policy and restart | out (`USER`) | **0 rows 4.5 min after the rule**, then `UNINDEXED` | stayed out, 0 rows; `UNINDEXED` again after the reboot |
>
>    So the purge of a newly excluded scope is work the service holds and **loses in a restart** -
>    and does not redo afterwards: R and S still held every row after a reboot. `-Execute` therefore
>    writes the rule, waits for the purge, and only then writes the policy and restarts (its
>    `-SelfTest` pins that order); a purge that does not finish in `-PurgeMinutes` stops it before
>    either. The way out when rows outlived an exclusion anyway is `-Execute -RebuildCatalog`:
>    `ISearchCatalogManager::Reset` took S's 20,048 rows out at once, and the catalog re-crawled its
>    other 431 items in about two minutes. **The policy is not a crawl-scope rule**: on its own it
>    leaves the scope IN and purges nothing - what it does is make Outlook's own
>    `Store.IsInstantSearchEnabled` read `False`, as did every exclusion above. **And the exclusion
>    holds against Outlook**: in U, S and N a non-elevated Outlook - which registers an INCLUDED
>    scope within seconds - ran five minutes on top and left it excluded. On a guest that was never
>    indexed (`CP-09`, no rule), the policy ALONE kept a non-elevated Outlook from registering at
>    all - six minutes, no rule, no row, the same after a reboot - so the unindexed guest's recorded
>    state (policy only) already holds against either kind of Outlook; what the policy cannot do is
>    exclude, or purge, a scope something else included. Not established:
>    whether the policy would also stop the purge (R shows the restart alone does), and why
>    `EnumerateScopeRules` stops listing the mapi16 rule once it excludes while
>    `IncludedInCrawlScopeEx` and `WorkingSetRules` both show it - `-Verify` judges by the former
>    and prints the registry beside it.
> 4. **What `OutlookAI-Unindexed` needs - written down, not done: that guest was not touched.** Its
>    recorded state (2026-09-16, not re-read here) is the policy only: no rule, no Outlook row. Item
>    3 says that holds against either kind of Outlook as it stands, but it is the one exclusion the
>    service does not report. To make it the real one: stage `Set-OutlookIndexingDisabled.ps1` and
>    `SearchCrawlScope.cs` from this change into `C:\OutlookAI-Q5\`; close Outlook with
>    `Testbed/host/Restart-Guest.ps1 -VMName OutlookAI-Unindexed -Execute` (add `-CancelLogonPrompt`
>    if its tier profile is up); then, elevated over PowerShell Direct, `-Execute` - it adds the user
>    EXCLUDE rule, finds no row to purge, keeps the policy and restarts the service - and `-Verify`,
>    which should now say `UNINDEXED` with reason `USER` and **without** the policy-only caveat.
>    Checkpoint it. If `-Execute` finds Outlook rows after all (a non-elevated Outlook ran there
>    before the policy), let it wait them out and do not restart the guest meanwhile - or use
>    `-RebuildCatalog`. Nothing else - no `Start-OutlookUnelevated.ps1` there - and once the rule is
>    in, that guest's index state no longer depends on how its Outlook is started (items 1 and 3).
>
> **MEASURED 2026-09-24 on `OutlookAI-Indexed` (Q68) - the registry half was exercised, and three
> things changed.** Full evidence in the script's banner.
>
> 1. **An administrator cannot write the crawl-scope rule at all.** The crawl scope manager's
>    keys (`SystemIndex`, `WorkingSetRules`, each `WorkingSetRules\<n>`, `SearchRoots`,
>    `DefaultRules`) grant `BUILTIN\Administrators` **ReadKey only**; `SetValue` belongs to
>    `SYSTEM`, `NT SERVICE\WSearch` and `TrustedInstaller`. The exact write the script made -
>    `New-ItemProperty -Name Include -PropertyType DWord` on a rule key - failed from an elevated
>    session with *"System.Security.SecurityException: Requested registry access is not
>    allowed"* (tried as a no-op on an existing `csc://` rule). **So the script's second layer
>    could never have worked**, and on the first machine that had a rule, `-Execute` would have
>    written the policy and then thrown - before the restart and before any verification. It is
>    now read-only; `-SelfTest` pins that the file writes nothing under those keys.
> 2. **The "indexed" guest has never been indexed.** No `mapi16` rule in any container, and
>    `mapiRows=0 mailRows=0` on two readings ten minutes apart, nine days after its 20,000-item
>    corpus was built. Outlook did not put itself into the crawl scope with its UI open for 10
>    minutes on the tier profile, 8 minutes on the corpus profile (Explorer fully loaded), or 5
>    more with the policy set; nor after the community-reported 25H2 key
>    (`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Search\SetupCompletedSuccessfully = 1`,
>    absent on the maintainer's workstation too, which has the rule); nor after the documented
>    *Default indexed paths* policy value was written for `mapi16://{SID}/` and `gpupdate /force`
>    run. The maintainer's workstation, on the same Windows release, **does** carry the rule.
>    Why a freshly built guest never gets one was **not established** here - it is now: every one
>    of those Outlook starts was elevated (the Q69 block above, and section 8 item 22).
>    `-Verify` used to call this state `SETTLING`, "wait and re-run", which never converges; it now
>    has its own verdict, **`NOT-IN-SCOPE`**, reported as a failure.
> 3. **So the 2026-09-16 `UNINDEXED` on the other guest proves less than it read.** It proves that
>    guest holds no Outlook row. It does not prove the policy excluded anything: its twin shows
>    the same zero rows with no policy at all. `-Verify` now prints that caveat under `UNINDEXED`
>    whenever no `mapi16` rule exists.
>
> Also measured the same day: `-RebuildCatalog`'s recipe (`SetupCompletedSuccessfully = 0` under
> `HKLM\SOFTWARE\Microsoft\Windows Search`, then a service restart) **does** trigger a full reset -
> Windows Search logged 1008 and 1004 *"{Reason: Full Index Reset}"*, and `Windows.db` went from
> 31.6 MB to 0.6 MB. That step was INFERRED until now.

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
>
> **STILL NOT ESTABLISHED, 2026-09-24 - and now known to be unmeasurable on the guests as they
> stand.** The experiment was set up on `OutlookAI-Indexed` precisely to answer it: exclude a
> guest that already holds Outlook rows and see whether they go. It could not run, because that
> guest holds **no** Outlook rows and never has (the note above). The durability question -
> does the exclusion survive Outlook? - is unanswered for the same reason: with the policy set,
> Outlook ran for 5 minutes with its UI up, the guest was restarted, and the policy was still `1`
> and no rule had appeared - but no rule had appeared **without** the policy either, so that
> reading distinguishes nothing. Both questions need a machine with Outlook **in** the crawl
> scope, and neither guest is one (section 8 item 22). Ordering therefore still matters for the
> day that changes; on today's guests it happens to be moot.
>
> **ANSWERED 2026-09-24 (Q69), on a guest that IS indexed** - both questions, and the ordering
> question with them (item 3 of the Q69 block above): `PreventIndexingOutlook = 1` removes
> **nothing** - it is not even a crawl-scope rule; a user EXCLUDE rule written through the API
> **does** remove the rows, the indexer purging them itself in about four and a half minutes - but
> only if nothing restarts the service in the meantime; and the exclusion **survives** a
> non-elevated Outlook, the one that would otherwise register the scope. Excluding first is still
> the cheaper order - nothing to purge, nothing to wait for - but it is no longer the only safe
> one: `-Execute` now waits the purge out itself, and `-RebuildCatalog` is the documented reset
> when rows outlived an exclusion made in the wrong order.

**Three durability risks the script does not defend against**, all established 2026-09-16 and
none of them a reason to avoid it - they are the reason the Group Policy layer is written *as
well as* the registry rule. **The first two are HISTORY since 2026-09-24**: they are risks to a
registry-written rule, and the script no longer writes one (an administrator cannot - see the
note above). They still describe what would undo an exclusion made through Indexing Options:

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
installer to repair Outlook". Close it with **`Testbed/host/Restart-Guest.ps1 -VMName <guest>
-Execute`**, which quits Outlook the way mailbox-safety rule 7 describes and then restarts the guest
without forcing anything - **never** `taskkill` it, and never `shutdown /r /t 5`, which is a forced
close by another name (Microsoft: a timeout above 0 implies `/f`).

**Verify by asking the INDEX, not the registry.** `Set-OutlookIndexingDisabled.ps1 -Verify` runs
a control probe, a scoped-MAPI probe and a scope-free mail probe through the same
`Search.CollatorDSO` provider the product uses, **twice**, `-SettleMinutes` apart - because a
machine believed unindexed while it is quietly still indexing produces measurements that look
fine and mean nothing. It returns five verdicts, and **three of them are not answers**:
`SETTLING` and `NO-INDEXER` both mean "ask again", not "pass", and `NOT-IN-SCOPE` (added
2026-09-24) means "asking again will not help - Outlook is not in the crawl scope at all".
Reading back the settings that were written proves nothing; neither does a single zero.
Cross-check with `outlook_health`'s `index.perStore[]`, which is the instrument section 1.1
names. `-SelfTest` drives the verdict table with synthetic readings on any machine.

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
| The tier profile and its POP3 account | `Testbed/guest/New-TierProfile.ps1 -Execute` - `tier-profile-forcepst.prf`, the default since 2026-09-24, and the `ForcePSTPath` it now writes - then one Outlook start, then `Testbed/guest/Rename-OutlookStore.ps1` | **Works from the scripts alone - proven from `CP-02` twice on 2026-09-24.** Imported at Outlook's first start on the machine: `C:\OutlookAI-Tier\Outlook.pst` minted and bound in that start (COM `DeliveryStore` and its Drafts resolve), renamed `tier@vm.invalid`, every `-Verify` check passing. Imported at a later start instead, the account stayed **unbound** until the start after - `-Verify` failed on it until then. |

So **the corpus profile is `/PIM` plus `Add-OutlookPstStore.ps1`**, then - after a guest restart,
because it refuses while Outlook runs - `Set-DefaultOutlookProfile.ps1` to make it the default. That
is also exactly how both guests' corpus profiles already exist.

**And the tier profile comes FIRST, at Outlook's first start on the machine** (Testbed/README.md
section 1 has the order and the two runs behind it). Three things about it a rebuilder meets:

* **Until 2026-09-24 it could not be rebuilt from the repository.** The working route needs
  `ForcePSTPath` - the directory Outlook mints an unnamed PST into, `REG_EXPAND_SZ` under
  `HKCU\Software\Microsoft\Office\16.0\Outlook` - and nothing under `Testbed/` set it; both guests
  got it from hand-run scratch scripts, which also defaulted to the wrong template by passing the
  right one explicitly. `New-TierProfile.ps1 -Execute` now writes it and reads it back, refuses a
  template that names a PST service or `DefaultStore` (the shape measured to leave `DeliveryStore`
  NULL), and `-Verify` requires the account's delivery store to be the profile's only PST, under
  `ForcePSTPath`. **The hand-run scripts also did damage worth knowing about:** they created the
  Outlook key with `New-Item -Force`, which on an existing key deletes every value under it - and
  `CP-05` is missing exactly the three values `Set-OfficeFirstRunSuppressed.ps1` writes inside that
  key (the classic-Outlook pins), while everything it writes elsewhere is present.
* **`ForcePSTPath` outlives the tier build.** It is per-user, so the corpus profile's `/PIM` store is
  minted there too: in the tier-first rehearsal and on `CP-05` it is `C:\OutlookAI-Tier\Outlook Data
  File - CorpusProfile.pst` (made corpus-first, before the value existed, it went to
  `Documents\Outlook Files`). Harmless - the tier's `-Verify` judges the profile's own PSTs, not the
  directory - but it is why a corpus lives in a directory called `OutlookAI-Tier`.
* **The account binds late when the import is not Outlook's first start.** In the corpus-first
  rehearsal the import start showed Office's one-time "Check out our new look" dialog, reached no
  account (no POP3 prompt), and left `DeliveryStore` NULL; the next start raised the POP3 prompt and
  bound it. The dialog is the likely cause, not a proven one. If a rebuild must import late:
  restart, start once more, and re-run `-Verify`.

**The hub has no Archive folder until something asks for one** (Q75, measured read-only
2026-09-24 on the rehearsed tier profile). The product resolved a store's designated Archive folder
with the undocumented `Store.GetDefaultFolder(39)` (`McpServer/OutlookAI.Core/Com/ArchiveFolderResolution.cs`).
On the hub PST - which had no `Archive` folder - that call **returned a folder named `Archive` at
the store's root, which it had just created**, and it passed the product's own verification (same
store, a mail folder, none of the core defaults). The verification step then also created
`Junk Email`: its `GetDefaultFolder(23)` is the only call there that names that folder. A second
resolution returned the same folder and created nothing. The same `GetDefaultFolder(23)` ran in
every search's freshness sweep, and in the count tripwire's census.

**Fixed the same day (Q84, maintainer decision (c)): a read-only lookup never creates a folder.**
On a store that is not Exchange, the product and the live tier now look a default folder up
without asking Outlook for it, and ask only once it is proven to exist (`SpecialFolders.Resolve`:
the store's `PR_VALID_FOLDER_MASK` for Inbox, Outbox, Sent Items and Deleted Items; the entry ids
designated on its Inbox for Drafts, Archive, Junk Email and the Sync Issues folders). An Exchange
store is asked exactly as before. What that changes here:

* `LiveMoveArchiveTests.ArchiveResolution_AllFiveStores_ReadOnly` answers
  `NoDesignatedArchiveFolder` for a PST without one, and asserts that the store's folder list
  reads the same after the lookup as before it. It is read-only on a PST now.
* `archive_mail` is the one path still allowed to create the folder, because it moves mail into
  it - and it reports that in `createdFolders`. The two archiving tests (the T2 move chain and the
  T3 stdio test) expect `createdFolders` exactly when the hub had no Archive folder at their start,
  so on a fresh guest the first of them to run creates it and the second finds it.
* The census and the artifact sweeps look default folders up the same non-creating way, so neither
  adds `Junk Email` - or a Sync Issues folder - to a bystander any more. A folder the zero-artifact
  count cannot prove exists on a non-Exchange store fails the count loudly instead of reading as
  empty.
* So do two checks inside write tools (decision A, direction 1): `update_draft`'s and
  `discard_draft`'s "the item is in Drafts", and `move_mail`'s "the target is not Deleted Items or
  the Outbox". Neither creates the folder it compares against, and a check that cannot be made
  refuses - `drafts_folder_unreadable`, `TargetGuardUnreadable`; the move check used to let the move
  through. On a guest, `LiveUpdateDiscardTests` now depend on the Drafts entry id designated on the
  hub's Inbox: if they refuse the hub's own drafts as `not_in_drafts_folder`, the designation is
  not where MS-OXOSFLD 2.2.3 puts it on this kind of store.

Whether a PST that is not an account's delivery store (the bystander, a corpus) keeps these
designations where the specification puts them was not measured, and neither was the hub after
Q84: the first guest run after it is where both get measured.

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
* **ONE ADDITION, written out because Q70 needs it (2026-09-24): a store this testbed created may be
  attached to the guest's OTHER profile as well - while it still holds no items.** The generator
  builds only in the account-less corpus profile and refuses, with no override and rightly, any
  profile holding an account; the tests run in the tier profile. So every store that gets a
  generated population - the hub, the bystander, the identity store - has to be mounted in both,
  as a corpus store is. And the hub cannot simply be created in the corpus profile, because only a
  store Outlook mints in the tier profile becomes the dummy account's delivery store. The line is
  drawn at EMPTY, as the one above is: attach the store to its second profile straight after it is
  created, before any population is built into it and before any test has run against it. Section
  3b gives the order. `Add-OutlookPstStore.ps1` is what does it - `AddStoreEx` opens a file that is
  already there instead of creating one - and it still never reads, moves or deletes an item.
  **A store that already holds items is still never attached by script, whatever its origin:** if
  a profile ever has to be rebuilt around a populated store, add that store through the GUI and
  say so in the session log. **This ADDS to the line drawn above rather than reading it more
  generously, and the maintainer should see it as an addition.**

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
  will write anything. In the TIER profile, that is. Whether the same PST attached to the
  account-less corpus profile reads `IsDataFileStore = true` there - it has no account delivering
  into it in that profile - is what the hub population's first build establishes (section 3b); the
  code says it should, and nothing has run it.
* **The bystander is named as an `.invalid` address too**, e.g. `bystander@vm.invalid`, and so is the
  identity store. Section 1.3 says why: the product finds a small store in the index only through
  mail addressed to it, and only for a name shaped like an address. Only the corpus - big enough to
  dominate the index's discovery sample - may keep a plain name.
* **The bystander is DECLARED, and is never the hub.** It goes in `bystanderStoreDisplayNames`,
  which denies it every write and keeps it censused; on a test guest it goes in
  `expectedStoreDisplayNames` as well, because that is the list `outlook_health` and `list_accounts`
  read. **Naming it in only the bystander list does NOT refuse the tier** - the census adds declared
  bystanders back in - and a store named only in `expectedStoreDisplayNames` is inside the
  identity-draft grant, which is the identity account's legitimate shape; section 1.3 has the one
  case that does refuse. The same declaration applies to `Corpus A`.

**Populate the hub and the bystander with the generator**, not by hand: `corpus-build --population
hub` and `--population bystander` build a small, tagged, deterministic population into each (section
3b). The identity store gets `--population identity` once section 2.8b has built it.

### 2.7 The mail sink - Inbucket 3.1.1

**The sink is a ready-made open-source program, not code in this repository, and that is the
maintainer's decision (2026-09-24).** A loopback SMTP-plus-POP3 server is a few hundred lines of
RFC 1939 whose failure modes - dot-stuffing a body line that begins with a period, UIDL identities
that move when the store is recreated, `STAT` octet counts - produce INTERMITTENT wrong answers
against Outlook, the fussiest POP3 client there is. Asked for a sink, the maintainer asked for "a
simple ready made open source tool... as simple as possible" rather than a new source of those.
It is media, recorded in `Testbed/MEDIA.md` under the Dependencies carve-out they confirmed.

**Which one, and why.** Three were read at the source of the release that would be pinned:

| | **Inbucket 3.1.1** | Mailpit 1.31.2 | smtp4dev 3.15.0 |
| --- | --- | --- | --- |
| Licence | MIT | MIT | BSD-3-Clause |
| Shape | one static Go binary; 11 MB zip | one static Go binary | self-contained .NET app; 77 MB zip |
| POP3 `PASS` with no password | **accepted** | refused, and the connection closed | refused |
| POP3 with no login configured | on | **off** - it starts only once a login is set | on, with authentication off |
| Which mail a POP3 login sees | **only the mailbox its `USER` names** | every message, to every login | every message, to every login |
| `DELE` | applied at `QUIT`, numbers fixed (RFC 1939) | applied at `QUIT`, numbers fixed | applied **at once**, and the mailbox renumbers |
| `TOP` | implemented | implemented | advertised in `CAPA`, not implemented |
| Hash published by its maintainers | `checksums.txt` on every release | none - only GitHub's computed digest | none - only GitHub's computed digest |
| Starts as a Windows service | no; a scheduled task does it | no | yes |
| Latest release, 2026-09-24 | 2025-12-06 | 2026-09-19 | 2026-03-03 |

Where each row was read: Inbucket `pkg/server/pop3/handler.go` (the `PASS` case checks only that a
`USER` came first; deletes run in `processDeletes` on `QUIT`), `pkg/policy/address.go` (mailbox
naming) and `pkg/config/config.go`; Mailpit `internal/pop3/server.go` (`PASS` with no argument
answers `-ERR must supply a password` and returns, closing the connection; `Run` returns at once
when no POP3 credentials are loaded) and `internal/pop3/functions.go` (every login is served the
latest 100 messages of the whole store); smtp4dev `Server/Pop3/CommandHandlers/PassCommand.cs`,
`DeleCommand.cs` and `CapaCommand.cs`, with no `TOP` handler in the tree.

**Two rows decide it, and neither is taste.**

* **The password.** On an unattended guest an Outlook logon prompt is a hang, and the tier
  profile's POP3 account deliberately stores no password. Inbucket accepts whatever arrives -
  `PASS` with an argument, with an empty one, or with none - so the sink half of the problem
  disappears. With either of the other two a password has to be stored in Outlook, and on Mailpit it
  would also have to match one configured in the sink.
* **Two accounts.** Section 2.8b adds an identity account beside the dummy one, both POP3 against
  this sink. A sink that shows every login every message lets whichever account polls first
  download the other's mail - an intermittent misdelivery that reads exactly like "the mail never
  arrived". Inbucket files each message under its recipient's local part and a login reads only
  that mailbox, so each account sees its own mail and nothing else.

What Inbucket costs in return: it is not a Windows service, so a SYSTEM scheduled task starts it -
Task Scheduler ships with Windows, so this adds nothing - and it releases about twice a year, which
does not matter for a pinned version and is why `-Verify` exists for the day it is bumped.

**Also set aside:** MailHog (MIT; SMTP and a web UI, no POP3, last release 2020-08-11); Papercut
SMTP (Apache-2.0; SMTP capture and a viewer, no POP3); MailDev (MIT; needs Node.js, no POP3);
MailCatcher (MIT; a Ruby gem, no POP3); GreenMail (Apache-2.0; has POP3, needs a Java runtime); and
full mail servers - hMailServer, Stalwart, mox - each far more machine than a sink.

**How it is configured.** Inbucket reads its configuration from `INBUCKET_*` environment variables
and nothing else, and a scheduled task cannot set environment for what it starts, so the task runs
a launcher, `C:\OutlookAI-Sink\run-sink.cmd`, that `Testbed/guest/Install-MailSink.ps1` generates
and whose `-Verify` fails on any hand edit. The dry run prints every value with its reason; the ones
that matter:

| Setting | Value | Why |
| --- | --- | --- |
| `INBUCKET_SMTP_ADDR` / `_POP3_ADDR` / `_WEB_ADDR` | `127.0.0.1:25` / `:110` / `:9000` | loopback only - anything else is an open relay on a test VM. The web UI cannot be switched off, so it is confined too |
| `INBUCKET_MAILBOXNAMING` | `local` | a message for `NAME@anything` is filed under mailbox `name` |
| `INBUCKET_SMTP_TLSENABLED`, `INBUCKET_POP3_TLSENABLED` | `false` | the tier profile sets `SMTPUseSSL=0` and `POP3UseSSL=0`; no `STARTTLS`, no `STLS` is offered |
| `INBUCKET_STORAGE_TYPE` / `_PARAMS` | `file` / `path:C$\OutlookAI-Sink\store` | survives a restart, and its message ids are timestamps - the memory store would reuse UIDLs after every restart |
| `INBUCKET_STORAGE_RETENTIONPERIOD`, `_MAILBOXMSGCAP` | `24h`, `500` | the defaults, written down: uncollected mail is purged after a day |

SMTP accepts mail for any domain, and `AUTH PLAIN`/`LOGIN` with any credentials or none, so the
tier profile's `SMTPUseAuth=0` works as it stands. Messages up to 10,240,000 bytes are accepted; a
live run's largest is a few kilobytes.

**THE MAILBOX RULE, WHICH EVERY ACCOUNT ON THIS SINK MUST FOLLOW.** A POP3 login's `USER` is used
*verbatim* as the mailbox name, and a mailbox name is the recipient's local part, lowercased, cut at
any `+`. So an account's POP3 user name must be exactly its lowercase local part: `tier` for
`tier@vm.invalid` - which is what `Testbed/guest/New-TierProfile.ps1` already sets - and `identity`
for section 2.8b's `identity@vm.invalid`. Get it wrong and nothing fails: the account reads an empty
mailbox for ever, and every arrival wait times out pointing nowhere. `-Verify` prints each account's
mailbox name so it can be compared.

**THE PASSWORD: SETTLED, BOTH HALVES - MEASURED 2026-09-24 ON `OutlookAI-Unindexed`.**

* **The sink half: it accepts anything, measured.** There is no credential store at all; `-Verify`
  logged in with `PASS` alone, `PASS ` and `PASS <anything>`, and all three reached TRANSACTION.
* **The Outlook half: OUTLOOK PROMPTS, AND NEVER CONNECTS.** With the sink up at `-LogLevel debug`,
  two starts of the tier profile each raised "Internet Email - tier" ("Enter your user name and
  password for the following server", User Name `tier`, Password empty), and the sink's log shows
  **no POP3 or SMTP session at all** for the five minutes Outlook ran - only its once-a-minute
  retention scans. Outlook asks before it connects, so the sink's leniency never comes into play.
  The dialog does not block COM, but no mail is ever collected, and every arrival wait would time out.
* **THE FIX, PROVEN: A STORED PASSWORD, WRITTEN BY SCRIPT.** `Testbed/guest/New-TierProfile.ps1
  -StoreSinkPassword -Execute`, with Outlook closed, seals a password - any value - into every account
  of the tier profile whose POP3 server is the sink: the community-documented storage below, with the
  plaintext NUL-terminated. At the next start the sink logged `read CAPA`, `read USER tier`,
  `read PASS any-value`, `Entering state TRANSACTION mailbox=tier`, `STAT`, `QUIT` and
  `Processing deletes mailbox=tier`, and no logon dialog appeared. `New-TierProfile.ps1 -Verify` now
  fails an account that polls the sink without such a password. Run it again once section 2.8b's
  identity account exists: it covers every sink account in the profile.
* **READ THE SINK'S LOG KNOWING IT IS BUFFERED.** Inbucket writes `inbucket.log` through a 4 KB
  buffer - the file grew in exactly 4,096-byte steps and held session lines back for six minutes - and
  `Install-MailSink.ps1`'s restart kills the process, losing what the buffer held. The first read
  above, 75 s into the no-password start, showed nothing at all; only the later flush told "no
  connection" apart from "not written yet". Read it after more has been logged.
* **How it was settled, with no mail anywhere** - kept as the procedure. Install the sink with `-LogLevel debug`. Build the
  tier profile and start Outlook in session 1 as usual (`Testbed/guest/Register-InteractiveTask.ps1`);
  it connects for its start-up send/receive, and a `Namespace.SendAndReceive` asks again if it does
  not. Then read `C:\OutlookAI-Sink\inbucket.log`:
    * `read USER tier`, then `read PASS...`, then `Processing deletes` with `mailbox=tier` - Outlook
      logged in holding no password. **Settled; nothing to change.**
    * the session stops after `USER` - `Client closed connection (state AUTHORIZATION)` - or never
      opens, and a window like "Internet E-mail - tier@vm.invalid" sits in session 1 - **Outlook
      prompts.** Close it; do not type into it, because a typed password is a step no rebuild repeats.
* **The storage the fix writes** - on the Outlook side, in the tier profile script, because the sink
  compares the password with nothing. [COMMUNITY: SecurityXploded's Outlook password notes for
  2002-2013, LaZagne's Outlook module, a 2022 borncity write-up placing it under the 16.0 hive]: a
  `REG_BINARY` value named `POP3 Password` in the account's subkey under
  `HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles\<profile>\9375CFF0413111d3B88A00104B2A6676\<n>`,
  holding a `0x02` byte followed by a DPAPI blob (current user, no entropy) of the password as
  UTF-16LE. **The open byte-layout question is answered for one layout: the plaintext WITH a
  terminating NUL** is what produced `read PASS any-value` (a bare plaintext was not tried). On this
  build the account's other values - `POP3 Server`, `POP3 User`, `Email` - are `REG_SZ`, so there
  was no binary sibling to copy the shape from. DPAPI seals it to `vmadmin` on this guest, so it is
  written on the guest and dies with an admin password reset (`Testbed/README.md` section 4).

**Installing it.** `Testbed/MEDIA.md`, "The mail sink", has the staging and copy-in lines. On the
guest, elevated - PowerShell Direct is fine, nothing here touches COM:

```
C:\OutlookAI-Q5\Install-MailSink.ps1                                    # the plan; reads and writes nothing
C:\OutlookAI-Q5\Install-MailSink.ps1 -ExpectedSha256 <hash> -Execute    # install, start, and the whole -Verify
C:\OutlookAI-Q5\Install-MailSink.ps1 -Verify                            # again, any time
```

Then **reboot the guest and run `-Verify` once more**: it reports how many seconds after boot the
sink process started, which is the evidence that it starts WITH the guest rather than because the
installer started it. Checkpoint after that. `-Uninstall -Execute` takes everything back out.

**RUN ON `OutlookAI-Unindexed`, 2026-09-24, and it worked first time.** `-SelfTest` 104 assertions, 0
failures on the guest's Windows PowerShell 5.1; `-ExpectedSha256 <pin> -LogLevel debug -Execute` over
PowerShell Direct with Outlook closed: every check passed, `VERDICT: SINK-READY`. After a graceful
restart (Outlook quit first, then `shutdown /r /t 0`) `-Verify` again: `SINK-READY`, the sink process
started **7 s after boot**. Checkpoint `CP-08-MAIL-SINK`. `-Uninstall` has not run on a guest.

**What `-Verify` proves** is in the script's own banner, check by check. In short: the task and the
launcher are what the script writes; the process runs from the install root and owns all three
listeners on `127.0.0.1`; a message submitted over SMTP comes back over POP3 byte-intact through
every dot-stuffing case and a base64 attachment; `TOP` works; one mailbox cannot see another's mail;
numbers stay fixed after `DELE`; a `DELE` without `QUIT` loses nothing; and after a restart nothing
deleted comes back and no id is reused. It writes to and deletes from only two mailboxes of its own,
which no account reads, so it may run while Outlook is open - it then skips the restart.

**Known deviations from RFC 1939, none of which this tier trips** [SOURCE]:

* `STAT` and `LIST` report the stored size, with LF line endings, which is smaller than the CRLF
  bytes `RETR` sends. Outlook reads to the terminator.
* `RETR` and `TOP` still serve a message marked for deletion, where RFC 1939 wants `-ERR`. Outlook
  never asks.
* A single line over 64 KB ends a `RETR` with an error. Outlook's own MIME never writes one.

**Ports.** Nothing needs the well-known numbers. If `25`, `110` or `9000` is taken or inside a
Windows reserved range - Hyper-V and WinNAT do reserve ranges on a VM - `-Execute` says so and writes
nothing; pick others and carry them into the tier profile and the settings `mailSink` block, which
must agree with the script. **Create no inbound firewall rule.** Loopback traffic is not filtered,
so a listener that needs a rule is a listener bound to `0.0.0.0`, which on a test VM is an open
relay; `-Verify` fails if a rule names the sink.

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

> **BUILT BY SCRIPT, 2026-09-24, on `OutlookAI-Indexed` - checkpoint `CP-09-IDENTITY-ACCOUNT`.**
> `Testbed/guest/Add-IdentityAccount.ps1`, in four phases because they alternate between Outlook
> closed and Outlook running, and the script never starts or stops Outlook itself:
>
> ```
> .\Add-IdentityAccount.ps1 -Phase Import -Execute          # Outlook CLOSED: .prf + ImportPRF
> <start Outlook once - it imports the account at start-up and clears ImportPRF itself>
> .\Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName identity@vm.invalid `
>                           -Path C:\OutlookAI-Tier\identity.pst -Execute      # Outlook RUNNING
> .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute    # Outlook RUNNING: read the store's IDs
> <restart the guest: Testbed/host/Restart-Guest.ps1 -VMName OutlookAI-Indexed -Execute -CancelLogonPrompt - never taskkill Outlook>
> .\Add-IdentityAccount.ps1 -Phase Bind -Execute            # Outlook CLOSED: bind the account
> .\Add-IdentityAccount.ps1 -Phase Verify                   # session 1, COM
> ```
>
> **What it produced, read over COM and again after a guest restart:** two POP3 accounts and two
> stores - `OutlookAI tier sink` delivering to `tier@vm.invalid` (`C:\OutlookAI-Tier\Outlook.pst`),
> `OutlookAI identity sink` delivering to its **own** `identity@vm.invalid`
> (`C:\OutlookAI-Tier\identity.pst`), each Drafts folder resolving, no store shared.
>
> **Why four phases and not one `.prf`, measured the same day:** a `.prf` creates the account
> (`OverwriteProfile=Append`, `Testbed/guest/identity-account.prf`) but Outlook binds it to the
> profile's **default** store - the tier store; a single `.prf` carrying both accounts got one
> minted store bound to both. So the identity PST is attached afterwards and the account is
> re-pointed at it by writing the two registry values in which Outlook itself records the binding
> (`Delivery Store EntryID`, `Delivery Folder EntryID`), with the exact bytes COM reports for that
> store. **That write is the one undocumented step** - the object model's `Account.DeliveryStore`
> is read-only and Microsoft documents only the GUI's *Change Folder*.
> `Docs/research/profile-automation-research.md` §4 has the full evidence.
>
> **Still to do on this machine, and NOT done by that script:** the **signature** below
> (`Testbed/guest/Set-AccountSignature.ps1`, never executed); the settings-file declaration below;
> and reading `Account.SmtpAddress` over COM, which the Object Model Guard blocks on these guests
> (section 8 item 23). The rest of this section is the specification the script was built to.

> **BUILT AGAIN, 2026-09-24, on `OutlookAI-Unindexed` - from `CP-09-ADDIN-READY` to checkpoint
> `CP-10-IDENTITY-ACCOUNT`.** Same script, same result: two POP3 accounts, `OutlookAI tier sink`
> on `tier@vm.invalid` and `OutlookAI identity sink` on its own `identity@vm.invalid`, each Drafts
> resolving, two distinct delivery stores - read over COM, then again after Outlook's own graceful
> quit and a fresh start. Before `-Phase Bind` the identity account sat on the tier store's
> EntryID, exactly as the Q64 measurement predicts. What that run added:
>
> * **The identity account PROMPTS for its POP3 password**, like the tier account (section 2.7):
>   the start that imports it raised an `Internet Email - identity` logon dialog. The build order
>   is therefore Bind, then - Outlook still closed - `New-TierProfile.ps1 -StoreSinkPassword
>   -Execute`, which stores the password on every account that polls the sink. Proven: the next
>   start raised no dialog, and the sink's debug log shows `read USER identity`,
>   `read PASS any-value`, `STAT`.
> * **Closing Outlook for `-Phase Bind` needs no guest restart**: the graceful quit of
>   `Testbed/README.md` step 4d, after cancelling that logon dialog (IDCANCEL; nothing typed).
> * **`Account.SmtpAddress` reads** - `tier@vm.invalid` and `identity@vm.invalid`, no prompt -
>   because this guest has `Set-OutlookProgrammaticAccess.ps1` (Q80). That closes the third
>   "still to do" above for any guest built with it.
> * **The signature is NOT done, and the shipped tool said it was.** `Set-AccountSignature.ps1
>   -Account identity@vm.invalid -Execute` (its first run anywhere) wrote the `Identity` file set
>   and printed `Verified` - but `manage_signature` had written `New Signature` onto the identity
>   PST's **data-file** entry (subkey `00000005`, clsid `{ED475414-...}`, whose `Account Name` is
>   the store's name `identity@vm.invalid`), not onto the POP3 account (`00000004`, `Account Name`
>   `OutlookAI identity sink`). `ProfileSignatureDefaultsStore` selects every subkey whose
>   `Account Name` contains `@`, whatever its clsid, and `list_signatures` reads back from the same
>   wrong place. Two consequences: the identity account has no signature Outlook will inject, so
>   `NewDraft_BusinessAccounts_BodyAboveTheirOwnIntactHtmlSignature` is expected to fail on this
>   guest until it is fixed; and on a real machine whose PST or data file is named after its
>   address - Outlook's own default for a POP3/IMAP account - the tool can write a user's default
>   signature onto the wrong subkey and report success. The fix is a product decision, not a
>   testbed one; it is left in place on this guest so the live run shows what Outlook does with it.
> * **The account's delivery FOLDER is the PST's hidden root, not an Inbox** - found at step 6
>   (section 4.1, defect 4). `identity.pst`, attached by `AddStoreEx`, has no Inbox, and
>   `GetDefaultFolder(6)` on it returns the root folder (NID `0x122`, no display name), so
>   `-Phase CaptureStore` captured that and `-Phase Bind` bound it. The same script did the same on
>   `OutlookAI-Indexed`. The draft tests are unaffected - they use the store's Drafts - but mail
>   delivered to this account lands where Outlook's folder tree does not show it. How a secondary
>   PST gets real default folders is the open question section 4.1 names.

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

> **ON A TEST GUEST, RENDER THIS FILE - DO NOT WRITE IT BY HAND (2026-09-24).**
> `Testbed/host/New-LiveTestSettings.ps1 -VMName <guest>` builds it on the host from the
> tokens-only `Testbed/live-test-settings.template.json` and that guest's section of
> `Testbed/testbed.json` (`liveTestSettings.<VMName>`), writes it into gitignored `.work/`, and
> prints the `Testbed/host/Copy-ToGuest.ps1` line that lands it at
> `C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json`,
> which is where the guest builds the suite. It refuses while any value in that section is still a
> placeholder, naming each; it refuses everything the rules below forbid; and it refuses the
> documented rules the tier does not enforce itself. The one thing it cannot check from the host is
> the one that matters most - that the guest's tier profile really mounts every store it names,
> under exactly those names - which is why those values have to be read off the guest over COM.
> **Nobody has done that yet, so no guest's file has been rendered.** The renderer **never writes
> into a `live-fixtures` directory on the host**, however the path is spelled, because that is
> where the maintainer's own hand-written file lives. Everything below still describes the file,
> and is still how the maintainer's own is written.
>
> **The table below used to be stricter than the code; since 2026-09-24 (Q77) it says what the code
> does.** It said naming a store in only one of `expectedStoreDisplayNames` and
> `bystanderStoreDisplayNames` refuses the tier. It does not: a bystander left out of
> `expectedStoreDisplayNames` is still censused, because the census adds declared bystanders back in
> on purpose (`T1/TripwireBystanderStoreTests.TheCensusWatchesEveryDeclaredBystanderEvenOneNoOtherListNames`),
> and a store in `expectedStoreDisplayNames` that is not declared is inside the identity-draft grant,
> which is exactly the shape section 2.8b's identity account is meant to have. The tier refuses only
> when that leaves the count tripwire nothing it could fail on. The maintainer chose to correct the
> documents rather than tighten the loader. **The renderer still holds a guest to the stricter
> shape** - every declared bystander named in `expectedStoreDisplayNames` too - because that list is
> also what `outlook_health` and `list_accounts` read, and on a machine this project builds there is
> no reason to leave a store out of it.
>
> **And the one list that used to mean two things is two lists (Q70, 2026-09-24):**
> `expectedStoreDisplayNames` is the WATCHED list and `indexedStoreDisplayNames` the INDEXED one.
> Section 1.3 says why, and what reads which.

Create `McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json`. It is
gitignored and must stay that way: it names real stores and this repository is public. Without
it the whole live tier refuses to start.

```json
{
  "machineProfile": "Portable",
  "testHubStoreDisplayName": "test@vm.invalid",
  "expectedStoreDisplayNames": [ "test@vm.invalid", "bystander@vm.invalid", "Corpus A" ],
  "indexedStoreDisplayNames": [ "test@vm.invalid", "bystander@vm.invalid", "Corpus A" ],
  "bystanderStoreDisplayNames": [ "bystander@vm.invalid", "Corpus A" ],
  "expectedDelegateStoreDisplayNames": [],
  "probeTerm": "invoice",
  "subjectOnlyProbe": {
    "storeDisplayName": "test@vm.invalid",
    "folderPath": "Inbox/OutlookAI-Corpus-Folder-Notices",
    "subjectTerm": "bulletin",
    "senderFragment": "noticebot"
  },
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
| `expectedStoreDisplayNames` | The **WATCHED** list: every store the tier profile mounts. The count tripwire censuses it (with the bystander and delegate lists unioned in); the identity-draft grant, `list_accounts` exactness, `outlook_health`'s reachability check and the archive-resolution test read it. Include the hub. | **yes** |
| `indexedStoreDisplayNames` | The **INDEXED** list: the stores the index-tier tests measure, in order - the hub first, the bystander second, the corpus last (section 1.3 says why the order is read). Every entry must be watched; never a delegate mailbox; must include the hub unless empty. **Absent means "the same as the watched list"**, which is what every file written before the split meant. Empty only on a `Portable` machine, and there every index test refuses rather than iterate nothing. | no |
| `bystanderStoreDisplayNames` | Stores the write allowlist must **refuse**. Declaring a store here is sufficient to have it censused - the census adds declared bystanders back in - so leaving one out of `expectedStoreDisplayNames` does not refuse the tier; name it there too all the same, because `outlook_health` and `list_accounts` read only that list, and the renderer holds a test guest to it. Both corpus stores belong here, and so does the bystander. Never the hub. | no |
| `expectedDelegateStoreDisplayNames` | Delegate/shared mailboxes. Watched, never written, folder hierarchy allowed to come and go. Empty here. | no |
| `probeTerm` | A word proven to hit this machine's search index, in every indexed store and in one text attachment of the hub. On a test guest it is the generator's `CorpusPopulation.ProbeTerm`, `invoice`; empty on a guest with no index. | Production only |
| `subjectOnlyProbe` | Coordinates of a population whose term is in the subject and not the body. Four fields, all or none. On the indexed test guest it is the hub population's Notices folder, and three of the four values are generator constants (`corpus-plan --population hub` prints them); only the store - the hub - is read off the guest. Absent on a guest with no index. | Production only |
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

## 3b. Keeping the fixture populations usable

**Added 2026-09-24 for Q70. NOTHING IN THIS SECTION HAS RUN ON A GUEST.** The generator half is
built and pinned on the host - `T1/CorpusPopulationTests`, 50 cases, no Outlook - and every step
below that opens a store is guest-only and unmeasured. The first build of each population is what
answers the questions at the end of this section.

**What a population is.** A small, curated, deterministic set of items the corpus generator builds
into a store the measurement corpus cannot serve: `corpus-build --population hub|bystander|identity`.
Same generator, same guards - the store allowlist, the four store facts, no account in the profile,
none of them overridable - the same `[OutlookAI-Corpus]` tag, the same manifest and the same two-key
teardown. What it adds is exactly what the corpus leaves out on purpose: senders and recipients,
attachments, conversations with more than one member, created subfolders, and every item addressed
to or sent from the store's owner, which is what lets the index find a small store at all (section
1.3).

| Population | Built into | Items | What it carries, and for whom |
| --- | --- | --- | --- |
| `hub` | the hub, `testHubStoreDisplayName` | 56 | Four conversations of four, alternating Inbox and Sent Items, whose newest member is the newest item in the store - one minute before the anchor. Sixteen received and six sent singles; six items in `Inbox/OutlookAI-Corpus-Folder-Projects`. Eleven attachments - PNG, `.ics`, `.eml` and text, and one mail carrying three - with the probe term `invoice` in one text attachment and in its parent's body. Read and unread mail. And the SF-6 subject-only population: twelve items in `Inbox/OutlookAI-Corpus-Folder-Notices`, all from `Noticebot Relay <noticebot@alerts.invalid>` and nobody else, `bulletin` in every subject and in no body. For the tests that read, page, cap or walk the hub; the index tests that read the first indexed store; SF-6; the attachment-kind recall; the conversation walk; the staleness frontier. |
| `bystander` | the plain bystander | 300 | Six conversations of three; 160 received and 50 sent singles spread over two years, one in eight received with an attachment; two populated subfolders of the Inbox, `-Projects` (40) and `-Suppliers` (32). Every folder inside the census identity budget. For the count tripwire's item-by-item path and the exclude-subfolders measurement. |
| `identity` | the identity account's delivery store (section 2.8b) | 8 | Five received, three sent, all to or from its owner - so the index knows the store exists and `outlook_health` does not report it missing. |

**It costs nothing to look at one.** `corpus-plan` is pure - no Outlook, runnable on the host - and
for a hub population it also prints the values a settings file must carry:

```
dotnet run --project McpServer/OutlookAI.RemediationTools/OutlookAI.RemediationTools.csproj -- \
  corpus-plan --population hub --store tier@vm.invalid \
  --corpus-id hub-indexed --seed 8181 --anchor 2026-09-24T08:00:00Z
```

`--store` is required even here, because the store's name decides the owner every item is addressed
to. `--count` may be left out: a population's size is part of what it is, and a different number is
refused rather than building part of one.

**The ids and seeds are assigned**, in `Testbed/testbed.json` under
`corpusIdConvention.populations` - `hub-indexed` 8181, `bystander-indexed` 8282, `identity-indexed`
8383, and `hub-unindexed` / `bystander-unindexed` on the other guest - for the same reason the corpus
ids are (section 2.9): the manifest is `corpus-<id>.jsonl`, every guest's manifest lands in one pull
directory, and a manifest is the only allowlist teardown will delete from. **The anchor is not
fixed**, and for the hub that is the point; see below.

### Building them the first time

Every store that gets a population must be mounted in BOTH profiles - built in the account-less one,
read in the tier one - and attached to its second profile **while it is still empty** (section 2.6,
the addition written out there). So the order is:

1. **In the tier profile**, after sections 2.5 and 2.8: the hub exists, minted by
   `New-TierProfile.ps1` (by default `C:\OutlookAI-Tier\tier.pst`) and named by
   `Rename-OutlookStore.ps1`. Create the bystander here too, new and empty, with
   `Add-OutlookPstStore.ps1`. Once section 2.8b has minted the identity account's store, it is here
   as well.
2. **Switch the default to the corpus profile** (`Set-DefaultOutlookProfile.ps1`) and restart
   Outlook - gracefully, under mailbox-safety rule 7.
3. **Attach each of those stores to the corpus profile by its path, before anything is in it**:
   `Add-OutlookPstStore.ps1 -ProfileName <corpus profile> -DisplayName tier@vm.invalid -Path
   C:\OutlookAI-Tier\tier.pst -Execute`, and the same for the bystander and the identity store. The
   name must come back byte-identical to the one the tier profile shows; the script checks.
4. **Build each population there**, as for a corpus: `corpus-probe`, then `corpus-build` dry, then
   `corpus-build --execute`, which runs its own census. For the hub:

   ```
   dotnet run --project <as above> -- corpus-build --population hub \
     --store tier@vm.invalid --allow-store tier@vm.invalid \
     --corpus-id hub-indexed --seed 8181 --anchor <now, UTC, to the second> \
     --manifest C:\OutlookAI-Q5\corpus-hub-indexed.jsonl --execute
   ```

   Before letting it proceed, read its own output: the store and profile lines accepted, `profile
   accounts: 0`, both corpus probes verified, and **`== enrichment probe ==` reporting sender,
   recipients, attachment and conversation index all written**. That third probe is new: a
   population is built only where one throwaway item proved every write it depends on, with no
   override. After the build, the census reads every item back - sender, recipients, attachments,
   conversation - and fails the build on any difference.
5. **Switch the default back to the tier profile**, restart Outlook, and on the indexed guest let
   the indexer settle. `outlook_health`'s `index.perStore[]` must then list every store in
   `indexedStoreDisplayNames` with rows.

The hub's build prints the `probeTerm` and `subjectOnlyProbe` values for the settings file; three of
the four `subjectOnlyProbe` fields are generator constants that `Testbed/testbed.json` already
carries, and the fourth is the hub's own name.

### The hub is rebuilt before every run

`LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier` asserts that the index frontier is not
in the future. On a guest whose newest item is weeks old, a product that misread local time as UTC
would still pass it; with the newest item one minute old, the same misreading puts the frontier an
hour or two in the future - the guest's UTC offset - and fails it. So the hub population is built
against **the moment the run starts**, which is why its anchor is not recorded in
`Testbed/testbed.json`: the manifest header records it, and the next rebuild reads it from there.

The rebuild is the same switch as above, and it tears down with the anchor the manifest records -
the anchor is part of the shape key, so a teardown given today's anchor is refused as a different
population:

```
:: in the corpus profile
dotnet run --project <as above> -- corpus-teardown --population hub \
  --store tier@vm.invalid --allow-store tier@vm.invalid \
  --corpus-id hub-indexed --seed 8181 --anchor <the anchor the manifest header records> \
  --manifest C:\OutlookAI-Q5\corpus-hub-indexed.jsonl --execute
dotnet run --project <as above> -- corpus-build --population hub ... --anchor <now> --execute
:: then back to the tier profile, let the index take the 56 items, and start the run
```

Teardown removes the population's items AND the two folders it created, and drains Deleted Items
behind itself: a delete in a PST is a soft one that re-issues the EntryID, so teardown re-scans and
deletes again, still by both keys. **The frontier test has to run within the guest's UTC offset of
the build, less its own five-minute tolerance** - under 55 minutes in winter and 115 in summer on a
`W. Europe` guest, counting the indexer's crawl and everything the run does before it gets there - or
it is back to proving what it proved before.

The bystander and identity populations have no such clock: no test reads their dates, so they are
built once and left alone. A checkpoint restored from before they were built needs them built again.

### What only a guest can answer

These are INFERRED from the code, and each is settled by the first build of the population it names:

1. **Does the hub, mounted in the account-less profile, pass the store guard?** `IsDataFileStore`
   should read true there - nothing delivers into it in that profile - and the other three facts do
   not depend on the profile. If it reads false, the hub population cannot be built by this route,
   and the answer is to report it, not to relax the guard.
2. **Does every enrichment write land?** The sender goes on through `PropertyAccessor`
   (`PR_SENDER_*` and `PR_SENT_REPRESENTING_*`), recipients through `Recipients.Add` and `Resolve`,
   attachments by value from a temporary file, and the conversation through `PR_CONVERSATION_INDEX`
   and `PR_CONVERSATION_TOPIC`. The enrichment probe refuses the build if any does not.
3. **Does the index carry what the tests read?** `FromAddress`/`FromName`, `ToAddress`, one
   attachment row per attachment, and a `ConversationID` shared by each conversation's members.
4. **Does a store mounted in two profiles give the index two scopes?** The per-store scope URL is
   `mapi16://{SID}/StoreDisplayName($Hash)/`. Corpus A has always been in the same position, so the
   answer - whatever it is - is not new to the populations.
5. **How long does the indexer take over a fresh hub?** It bounds how soon after the rebuild the run
   can start, and so how much of the UTC-offset margin is left for the run itself.

**What a population does not carry, and what that costs - OPEN, for the maintainer.** Every
population item is a dated `IPM.Note` with a plain-text body: no appointment, contact or task, no
unsent item, no HTML. So on a guest the index holds **no undated rows** - bar whatever drafts other
tests happen to have left in the hub at that moment, which is nothing to measure against - and the three
`LiveOrderKeyCollationTests` - which exist because rows with no `System.Message.DateReceived` share
the `ORDER BY ... DESC` cut with mail - run, pass, and say nothing about undated rows:
`NullCollation_UnderDateReceivedDescending_IsMeasured` reports `no-undated-rows-in-sample` instead
of the provider's NULL collation, `OrderKeyFloorPredicate_IsAccepted_AndAdmitsOnlyDatedRows` excludes
undated rows from a sample that has none, and
`WidenedSearch_NeverReturnsFewerRowsThanTheOldMailKindShape` compares two shapes that return the
same rows. Before the populations they could not run at all; now they run weak. Giving the hub a few
undated items of other classes would make them real, and is a change to what a population is - so it
is recorded here rather than made quietly.

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

### 4.1 The first live-tier run on a VM - `OutlookAI-Unindexed`, 2026-09-24

**Why this section exists.** Until this date the live tier had never run on a test machine (section 9,
`Testbed/README.md` section 6 item 13). This is the record of bringing `OutlookAI-Unindexed` from
`CP-05-CORPUS-B-CLEAN-UNINDEXED` to the full design, one scripted step at a time, each step proven
on the guest and checkpointed, and then running the tier there. Raw logs of every step were kept
outside the repository (`.work\g2-buildout\` in the main checkout).

**The build-out, step by step:**

| Step | What ran | Verdict | Checkpoint |
| --- | --- | --- | --- |
| Restart route | CP-05's saved memory carries an Object Model Guard prompt with no client left; `Application.Quit()` was ignored while it stood. Answered **Deny** by the dialog's own `WM_COMMAND` (a `BM_CLICK` was ignored), then a graceful `Quit()` - exited in 2 s | - | - |
| 1. First-run settings | `Set-OfficeFirstRunSuppressed.ps1 -Verify` (10 OK, 3 FAIL: exactly the three classic-Outlook values), `-Execute`, `-Verify` | 13 of 13, still 13 after two Outlook restarts | `CP-06-FIRSTRUN-REPAIRED` |
| 2. Programmatic access (Q80) | `Set-OutlookProgrammaticAccess.ps1`: a control `-Verify` with nothing written, `-Execute`, `-Verify` | control `PROMPTED` in 0.7 s; then `NO-PROMPT`, `SmtpAddress` in 16 ms | `CP-07-PROGRAMMATIC-ACCESS` |
| 3. Mail sink | `Install-MailSink.ps1 -LogLevel debug -Execute`, graceful restart, `-Verify`; then the password question (section 2.7) | `SINK-READY` twice, started 7 s after boot; Outlook PROMPTED and never connected, so `New-TierProfile.ps1 -StoreSinkPassword` - then `read USER tier` / `read PASS any-value` | `CP-08-MAIL-SINK` |
| 4. Add-in | `Publish-AddInPayload.ps1` on the host (`fe65ced`, host unchanged), `Install-OutlookAIAddIn.ps1` `-SelfTest`, `-Verify`, `-Execute`, restart, `-Execute`; then `Set-OutlookIndexingDisabled.ps1 -Verify` | `NOT-INSTALLED`, then `ADDIN-READY` twice (tuning state 3.5 s and 3 s after the start; trust entry kept the second time); the index verify `NO-INDEXER` when its first reading fell 74 s after boot, before Windows Search's delayed start, then `UNINDEXED` on a re-run | `CP-09-ADDIN-READY` (the identity import of step 5 already pending in it) |
| 5. Identity account | `Add-IdentityAccount.ps1` Import (before CP-09), a start that imported it, `Add-OutlookPstStore.ps1` + `-Phase CaptureStore`, graceful quit, `-Phase Bind`, `New-TierProfile.ps1 -StoreSinkPassword -Execute`, start, `-Phase Verify -TrySmtpAddress`; quit, start, Verify again. Then `Set-AccountSignature.ps1 -Execute` | the import start raised the identity account's POP3 logon dialog (cancelled); after Bind + the stored password no dialog, the sink logged `read USER identity` / `read PASS any-value`; `OK` - 2 accounts, 2 distinct delivery stores, `SmtpAddress` of both read with no prompt - twice. The signature printed `Verified` and landed on the wrong subkey (section 2.8b) | `CP-10-IDENTITY-ACCOUNT` |
| 6. Fixture populations (section 3b) | In the tier profile `Add-OutlookPstStore.ps1` made `bystander@vm.invalid` (`C:\OutlookAI-Tier\bystander.pst`); all three population stores read 0 items. Default to `CorpusProfile`, the hub (`Outlook.pst`), bystander and identity PSTs attached there by path, names byte-identical. `corpus-probe --population hub`, then per population a dry run and `--execute` - `hub-unindexed` 8181, `bystander-unindexed` 8282, `identity-unindexed` 8383 (the last id is not yet in `testbed.json`), one anchor `2026-09-24T16:08:00Z` - then `corpus-census` of all three and of Corpus B. Default back to `OutlookAI-Tier` | **NOT CLEAN - every population build exited 1.** Placement census clean: 56, 300 and 8 items, each ordinal once, in the folder the manifest records. Read-back FAULTS: 42 of 56, 244 of 300 and 5 of 8 items - exactly the RECEIVED ones - carry no owner recipient. And three more defects, below. Corpus B: its own census clean (20,000), but 12 probe items now sit in its Drafts | `CP-11-POPULATIONS-BUILT-WITH-FAULTS` - evidence, not a base to build on |
| 7. SDK and suite | `Publish-LiveTierPayload.ps1` on the host (source archived from `329925d`; the 54-package feed reused), both archives expanded on the guest, `Install-DotnetSdk.ps1 -ExpectedSha512 <the published hash> -Execute` over PowerShell Direct; then `-Verify` from a new session | `TEST-READY` both times: the installer exited 0 after 74 s, the offline probe ran (`OUTLOOKAI-SDK-PROBE-OK 10.0.12 x64`), **2,794 tests discovered**, 17 executed and passed; 148 s end to end. The Release build baked `McpServerExePath` = `C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer\bin\Release\net10.0-windows\OutlookAI.McpServer.exe`, and that file exists - question 12 answered by construction, as the installer's banner predicted | `CP-12-SDK-TEST-READY` |
| 8. Settings | Store names read over COM in the tier profile (three stores, the two accounts on their own stores); `OutlookAI-Unindexed`'s section of `Testbed/testbed.json` filled - watched `tier@`, `bystander@`, `identity@vm.invalid`, indexed `[]`, bystander `bystander@vm.invalid`, `corpus` null (the tier profile does not mount Corpus B), `mailSink` loopback 25/110 with a 2,000 ms connect timeout; `New-LiveTestSettings.ps1 -VMName OutlookAI-Unindexed`; the file copied to the suite's `live-fixtures` | rendered and admitted (Portable; hub `tier@vm.invalid`; identity `identity@vm.invalid` draft-and-delete only); SHA-256 `401A4BA8...5C6BF1` on host and guest alike | `CP-13-LIVE-SETTINGS` |
| 9. First live run | **ON HOLD, not started** - by instruction: the live tier's own census, `LiveOutlookTestMailer.CaptureMailFolderCensus`, calls `GetDefaultFolder` for Deleted Items and ids 19-23 on every store, and on a PST lacking one of those folders that CREATES it (step 6 above measured the same thing from the generator's side). A run now would write into the bystander PST and measure the harness. It waits for the non-creating resolver and the matching harness fix | - |

**Step 6's defects, measured on the guest, none fixed here** (the generator is outside this work's
files; each needs its owner's fix and a rebuild, and two need a decision first):

1. **The owner is never a resolved recipient.** For a store named as an address the generator makes
   the owner `Name = Address = tier@vm.invalid` and adds the recipient as the spec
   `tier@vm.invalid <tier@vm.invalid>`; Outlook's `Resolve()` refuses that string, so every received
   item carries one UNRESOLVED To row whose `Address` is empty and whose `Name` is the whole spec.
   Sent items (`Kester Wren <kester.wren@margie.invalid>`) resolve to SMTP one-offs. The enrichment
   probe passed because its throwaway item is addressed to correspondents, never to the owner.
2. **A PST attached with `AddStoreEx` has no Inbox and no Sent Items, and the generator does not
   notice.** `Store.GetDefaultFolder(olFolderInbox)` on such a PST returns the PST's non-IPM ROOT
   folder (NID `0x122`, display name empty), so the bystander's 172 Inbox items and the identity
   store's 5 sit there, invisible in Outlook's folder tree - and the bystander's two "Inbox"
   subfolders were created under that root, outside the IPM subtree. Sent Items fell back to a
   created `OutlookAI-Corpus-Folder-5` at the top of the IPM tree (bystander 56, identity 3). The
   placement probe reported `target= landedIn=` - empty names - and VERIFIED it, and the census,
   which checks the recorded folder by EntryID, calls it clean. The hub is unaffected: it is the
   tier profile's minted default store and has every default folder. Deciding how a secondary PST
   gets real default folders before anything is built into it is the open question.
3. **Failed probe rungs strand their item in the DEFAULT store's Drafts - a store not on
   `--allow-store`.** Every probe session left three items (`placement InPlaceWithSentFlag`,
   `placement InPlaceOnly`, `date ObjectModel` - exactly the three rungs that failed) in the Drafts
   of `Outlook Data File`, Corpus B, which was the account-less profile's default store and was
   named on no allowlist: 12 items from four sessions, including both hub sessions. The purge only
   scans the target store. Nothing here deletes them - mailbox-safety rule 1, and no helper covers
   it.
4. **The identity account delivers into that invisible root.** `Add-IdentityAccount.ps1 -Phase
   CaptureStore` reads the "Inbox EntryID" with `GetDefaultFolder(6)`, so on this guest - and by the
   same script on `OutlookAI-Indexed` - the account's `Delivery Folder EntryID` is the PST root
   (`...22010000`), not an Inbox. Drafts resolves, which is all section 2.8b's two tests use; mail
   delivered to the account would land where nothing shows it.

Also measured here, and the reason for the harness fix the first live run now waits for:
`GetDefaultFolder` CREATES some missing special folders on such a PST. The bystander PST held only
`Deleted Items` before its build and gained `Drafts` and `Junk Email` during it - the generator's
store scan calls `GetDefaultFolder` for Drafts, Inbox, Sent Items, Junk Email, Outbox and Deleted
Items (`ScanFolderIds`) - while Inbox came back as the root and neither Outbox nor Sent Items was
created. The identity PST gained `Junk Email` the same way; its `Drafts` was already there, very
likely from `Add-IdentityAccount.ps1`'s own `GetDefaultFolder(16)` reads (CaptureStore, Verify),
which are therefore not the pure reads that script's banner calls them.

**What the build-out found, beyond the step verdicts:**

* **An Outlook started by COM (`-Embedding`) is not in the Running Object Table**, so
  `GetActiveObject` fails (`MK_E_UNAVAILABLE`) while `New-Object -ComObject Outlook.Application` from
  the same session and integrity level attaches to it. And attaching to an Outlook started as a
  PROGRAM launches a transient `OUTLOOK.EXE -Embedding` that hands off and exits within seconds.
* **A modal dialog swallows `Application.Quit()`** - measured with the guard prompt: `Quit()` returned,
  and Outlook was still up four minutes later.
* **Office LTSC 2024 does not register the VSTO runtime** Installer.iss looks for (`v4R` absent; only
  `v4` `10.0.60910`), so the add-in installer's own runtime step is load-bearing on a guest.
* **Inbucket's log is written through a 4 KB buffer** (section 2.7): read it after more has been logged.
  An SMTP session that only says `EHLO` and `QUIT` - no `MAIL`, no `RCPT`, no `DATA`, so no message
  and no mailbox - adds enough debug lines to push the buffer out; one such session was enough.
* **`manage_signature` can bind a signature to a data file** (section 2.8b): it picks the profile
  subkey by an `@` in `Account Name`, and a PST named after an address has one.

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
   **CORRECTED 2026-09-24, three times over.** (a) **The closed-negative half is not negative:**
   both accounts are built by script with no GUI - the tier account by `New-TierProfile.ps1` with
   `tier-profile-forcepst.prf` (both guests, 2026-09-15/16), the identity account with its own
   delivery store by `Testbed/guest/Add-IdentityAccount.ps1` (`OutlookAI-Indexed`, one
   undocumented registry step - section 2.8b and the research doc's §4). There is no GUI pass left
   in sections 2.8 and 2.8b. (b) **`IProfAdmin` was never "measured broken on this Office build".**
   The deleted interop declared it with the wrong IID - `00020379` where Microsoft's `MAPIGuid.h`
   has `IID_IProfAdmin` = `0002031C` - and a `QueryInterface` for the right one succeeds on this
   build (`Testbed/guest/OutlookMapiInterop.ps1`, REOPEN section). Nothing was rebuilt on MAPI; the
   routes above stay. (c) **The hazard section 2.5a carries does not occur as described on this
   build:** Outlook deletes `ImportPRF` after importing and writes `First-Run` back - observed on
   the tier profile's own nine-day-old import and on four controlled imports on 2026-09-24. What
   does stay true is that an `OverwriteProfile=Yes` import drops every store the file does not
   name; a repeated import still costs the corpus, it just has to be repeated by someone.
   `Add-OutlookPstStore.ps1` has now also run on a guest (2026-09-24, `AddStoreEx` returned
   promptly, the rename carried) - so of the profile scripts only `New-OutlookProfile.ps1` is
   still unexecuted as far as this guest's record goes.
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

22. **CLOSED 2026-09-24 (Q69) - `OutlookAI-Indexed` was never indexed because every Outlook the
    testbed ever started on it was ELEVATED, and an elevated Outlook does not use Windows Search at
    all.** The session-1 route, `Testbed/guest/Register-InteractiveTask.ps1`, registers its task at
    `RunLevel Highest`, so every Outlook started through it - every UI watch, every
    `Build-Corpus.ps1` run, every COM start - inherited an elevated token. A non-elevated Outlook
    adds itself to the crawl scope within seconds - and the maintainer's workstation has exactly that
    rule (its Outlook's integrity level was not inspected: the host is read-only registry and files). Measured on `OutlookAI-Indexed` from the same clean
    checkpoint (`CP-09-IDENTITY-ACCOUNT`: no `mapi16` rule, zero Outlook rows), the same profile
    (`CorpusProfile`), a graceful restart in between, and nothing but the integrity level differing
    (`.work/aa5e-2026-09-24-q69-index-scope/ab-control.txt`):

    | | ELEVATED (`RunLevel Highest`, the testbed's route) | NOT elevated (`RunLevel Limited`) |
    | --- | --- | --- |
    | crawl scope | no rule in 8 minutes, Explorer up the whole time | a search root and a user INCLUDE rule for `mapi16://{SID}/` within 8 s of the start |
    | catalog | 0 Outlook rows, nothing queued | crawl running at once: 19,874 item notifications queued within 70 s (the whole corpus took 8.5 min - below) |
    | `Store.IsInstantSearchEnabled` - Outlook's own view, read over COM at the same level | `False` | `True` |
    | `mssprxy.dll` - the Windows Search interfaces' proxy - loaded in OUTLOOK.EXE | no | yes |
    | per-store marker in `HKCU\...\Outlook\Search` (a DWORD named by the PST path) | never written | written - the same kind of value the maintainer's machine carries per store |

    Microsoft's own wording for the state, as support pages quote it: *"Instant Search is not
    available when Outlook is running with administrator permissions."* Nothing announces it: no
    event, no prompt, no log line.

    **Ruled out, each by evidence rather than by assumption.** The protocol handler: `Mapi16` is
    registered (`ProtocolHandlers\Mapi16\0` -> `Outlook.Search.MAPI16Handler.1` -> CLSID
    `{F8E61EDD-...}` -> Click-to-Run's `Interceptor.dll` -> `MAPIPH.DLL`), identically on the
    maintainer's workstation, and it crawls the moment a rule exists. Outlook's search settings: the
    guest's `HKCU\...\Outlook\Search` holds only `IndexAvailableBody=0`; there is no disabling value.
    Policy: nothing under `Windows Search` or `Outlook\Search` in HKLM or HKCU. Event logs: no
    protocol-handler or gatherer failure. The 25H2 key and *Default indexed paths* (Q68). **And the
    comparison with the maintainer's workstation settles the shape**: the rule a non-elevated Outlook
    wrote on the guest - `WorkingSetRules` `Include=1 Suppress=0 Default=0 Policy=0 NoContent=0
    Container=0`, `SearchRoots` `ProvidesNotifications=1 Container=0` - is value for value the
    workstation's (which carries one more, `IntelligentlyAdded=0`). The workstation's key timestamps
    cannot say when its rule was made: every crawl-scope key there carries the same last-write time,
    because the service rewrites them all on every save.

    **The fix is two steps, and both are now in the build order** (`Testbed/README.md` section 1,
    7b and 8c). The scope is written deterministically through the documented API -
    `Set-OutlookIndexingDisabled.ps1 -Enable -Execute` (the Crawl Scope Manager interop in
    `Testbed/guest/SearchCrawlScope.cs`), which leaves exactly the shape above - and the stores are
    crawled with Outlook running NOT elevated (`Testbed/guest/Start-OutlookUnelevated.ps1`): **the
    20,000-item corpus was fully crawled 7.6 to 9.6 minutes after Outlook's start**, four times over:
    9.6, 8.6, 7.6 and 8.8 minutes to the first reading with nothing left, readings a minute apart
    (the last two are the committed build step replayed from the clean checkpoint -
    `.work/aa5e-2026-09-24-q69-index-scope/task2b-build-step-and-checkpoint.txt`), about 3,500 items a
    minute at the peak; the tier profile's two small stores then took under three. **How
    "finished" is known**: three signals fell due in the same minute - the catalog's
    `NumberOfItemsToIndex` queues at 0, `GetCatalogStatus` back to `IDLE` (it had gone
    `INCREMENTAL_CRAWL` -> `PROCESSING_NOTIFICATIONS`), and the Outlook row count no longer moving
    (20,030 = 20,000 items plus the store's folders). `Set-OutlookIndexingDisabled.ps1 -Verify`
    now requires all three across two readings before it says `INDEXED`, and `-WaitMinutes` waits
    for them. **A rule alone crawls nothing**: with the scope in, the count stayed at zero for four
    minutes with Outlook closed and for six with Outlook running ELEVATED, whose
    `IsInstantSearchEnabled` still read `False`. The guest is INDEXED and checkpointed as
    `CP-10-INDEXED`. The exclusion was then measured from that checkpoint - five ways, and the order
    it needs - in section 2.4's Q69 block, item 3; the guest was restored to `CP-10-INDEXED` after
    it and re-verified `INDEXED` (20,048 rows over three stores, the catalog IDLE, nothing queued).

    **What follows for the live tier on the guests - a decision, not taken here.** The index tests
    need the index to MOVE while they run - a test creates an item and waits for the index to show it
    - and on these guests the index moves only while a NON-elevated Outlook runs. Run through `Register-InteractiveTask.ps1` as the tier would be
    today, Outlook is elevated, `IsInstantSearchEnabled` is `False`, and nothing a test creates is
    ever indexed: the "not indexed yet" state, permanently, which the suite reads as a slow indexer.
    So the tier's session-1 route must run Outlook - and the test host with it, because an elevation
    mismatch breaks COM attach (v3.MD S8) - at `RunLevel Limited`. Options: a `-RunLevel` parameter
    on `Register-InteractiveTask.ps1`; a second, Limited task for the tier; or Outlook started by
    `Start-OutlookUnelevated.ps1` with the suite in a Limited task beside it. Not built here -
    `Register-InteractiveTask.ps1` was not this round's file, and which route is the user's call.

    **A side finding for the identity tests, not chased.** The index names a store by the name in
    its profile's service, not by the root-folder name COM reports. The tier store appears as
    `tier@vm.invalid($93f42b43)`, but the identity store - named by `Add-OutlookPstStore.ps1` through a
    root-folder rename - appears as `Outlook Data File($b25ac20a)`, its folders under
    `/Outlook Data File`, while `Store.DisplayName` reads `identity@vm.invalid` (identified by
    elimination: it is the tier profile's only other store). `IndexSearchService.TryDiscoverStoreScopeByAddress`
    accepts a store only when the index's name EQUALS the address, so it cannot find the identity
    store here. Whether anything the identity tests exercise goes through that path is not
    established.
23. **OPEN - Outlook's Object Model Guard prompts on the guests, and the live tier reads protected
    members.** Windows Security Center reports Defender's signatures out of date (dated 2025-09-17,
    372 days on 2026-09-24; the guests have no network), so Outlook treats every out-of-process COM
    caller as untrusted and raises *"A program is trying to access email address information stored
    in Outlook"* - a modal dialog on the guest's desktop that nothing answers. Measured 2026-09-24 on
    `OutlookAI-Indexed`: `Account.UserName` (not on Microsoft's published list of protected members)
    raised it and blocked; `Account.SmtpAddress` (on the list) blocked in one run and returned while
    another run's prompt was already pending; `DisplayName`, `AccountType`, `DeliveryStore` and its
    folders never did. Microsoft's list also covers `MailItem.Body`, `HTMLBody`, `SenderEmailAddress`,
    `Recipients`, `PropertyAccessor` and more - members the product reads - so **expect the live tier
    to stall on a prompt on these guests**. It did not show on 2026-09-15, when the account probe read
    `SmtpAddress` freely; what changed in between is not established. Options, each a decision: the
    documented Outlook security policy that auto-approves programmatic access, updated signatures
    staged offline, or the Trust Center's programmatic-access setting (HKLM, per machine).
    **Still open and still the user's decision, 2026-09-24 (Q69) - one thing changed around it.** A
    pending prompt used to be cleared by a forced guest restart. `Testbed/host/Restart-Guest.ps1`
    will not do that: it refuses to quit Outlook behind ANY visible Outlook dialog and names it, and
    it never answers a security prompt - its one exception, `-CancelLogonPrompt`, cancels only the
    POP3 "Internet Email - <account>" logon prompt. So a guest with this prompt up now needs a person
    (or the decision above) before it can be restarted gracefully. Q69 read no address property and
    raised no prompt: `Store.DisplayName`, `FilePath` and `IsInstantSearchEnabled` never did.

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

  **DECIDED 2026-09-24, and it settles the version policy for good (Q73).** The maintainer's
  words: *never* a subscription - no Microsoft 365 Apps, on any machine. So:

  * **The guests stay on Office LTSC 2024**, the perpetual volume-licensed release.
  * **The maintainer's workstation stays on Office LTSC 2021 until he chooses to move it**, and
    until then it counts as **a production user on an older version**: a defect that reproduces
    only there is a real defect a real user on that version would hit, reported by the one user
    who happens to be the maintainer. LTSC 2021 leaves Microsoft support on 13 October 2026; that
    is his call, not the testbed's.
  * **There is no current-channel coverage anywhere, and that is accepted.** Most users on
    Microsoft 365 run builds newer than either machine. Nothing in this repository can close that
    gap without a subscription, and the subscription is ruled out.
  * **LTSC 2024 is not the last perpetual release.** Microsoft has committed to another
    on-premises release after it (see `Testbed/MEDIA.md` for sources). When it ships, moving the
    guests is a change of staged media in a scripted build, not a rebuild by hand.

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
