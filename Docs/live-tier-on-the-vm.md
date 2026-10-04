# The live-tier test VM: building it from nothing, and running it

**Who this is for.** Someone rebuilding this machine after it has been deleted, corrupted or
moved to another host, with nothing but this repository and a Windows ISO. It assumes no
knowledge of how the tier grew up. It is also the reference for running the tier once the
machine exists.

**Read `AGENTS.md`'s Mailbox Safety section first.** Nothing here overrides it. Every rule in
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
| Hub | as its account's address, e.g. `tier@vm.invalid` | yes | **first**, on the indexed guest | the **hub population** (56 items - generator v2 with its 12 undated items switched off, Q98 (a) 2026-10-03; section 3b) plus whatever a run is writing |
| Bystander | as an address, e.g. `bystander@vm.invalid` | yes, and a **declared bystander** | **second**, on the indexed guest | the **bystander population** (300 items - its 42 undated items switched off, Q98 (a); section 3b) - never written by a test. It has **no Inbox and no Sent Items**, on purpose (below) |
| Corpus A / Corpus B | anything - `Corpus A` in the examples | yes, and a **declared bystander** | **last**, on the indexed guest | the measurement corpus: at least **160,000** items on the indexed guest (`minimumItemCount` in `Testbed/testbed.json`) - **built there at 160,000 on 2026-10-03, as `Corpus A`, mounted in both profiles (section 4.2d)**; no minimum is recorded for the unindexed guest's |

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

**It is populated by the generator** (`corpus-build --population bystander`, section 3b) - 300 dated
items in four mail folders, every one inside the identity budget, two of them subfolders of the one
the "Inbox" items go into. (Generator v2 also defines 42 undated appointments, contacts and tasks for it;
they are switched off since 2026-10-03 - Q98 (a), section 3b.) The
earlier objection that "the generator tags everything it creates, and the bystander's whole job is to
be untouched" does not hold: the tag is the CORPUS tag, which no artifact sweep can select, and no test
writes to the store either way.

**And it keeps the shape it was chosen for: no arrival folders.** The bystander is also the
absent-arrival-folders store (Q5): a PST attached with `AddStoreEx` has Deleted Items and none of
Inbox, Sent Items, Drafts, Junk Email, Calendar, Contacts or Tasks - measured on `OutlookAI-Unindexed`,
2026-09-24 - and nothing is to give it any. The generator used to ask Outlook for them anyway: it
CREATED Drafts and Junk Email in it, and for its Inbox got the PST's hidden root, where 172 items then
sat out of sight (section 4.1, step 6). Since generator v2 no lookup creates a folder, and a folder the
store lacks is replaced by a visible **stand-in** under the store's root - `OutlookAI-Corpus-Folder-6`
for the Inbox items and its two subfolders, `-5` for Sent Items, and `-9`/`-10`/`-13` (typed Calendar,
Contacts and Tasks folders) for the undated kinds while those are switched on, which since 2026-10-03
they are not - recorded in the manifest, so teardown removes it.
None is a default folder, so the store still has no arrival folders. Whether any TEST needed the
bystander's Inbox was checked, test by test, before deciding to keep this shape: section 3b.

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
`AGENTS.md`'s Dependencies rule that the maintainer confirmed.

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
**Corrected 2026-09-27: it has run, on both guests.** `-Verify` reported `SINK-READY` on
`OutlookAI-Unindexed` on 2026-09-24 (section 2.7, `CP-08-MAIL-SINK`) and on `OutlookAI-Indexed` the
same day and again on 2026-09-27 (section 4.2 step 3, section 4.2b step 4.3). So the `[SOURCE]` rows
of the table above that `-Verify` asserts - any password or none accepted, deletes applied at
`QUIT`, numbers fixed, `TOP` implemented - are measured on both guests. The RFC deviations section
2.7 lists are still read from source.

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
shell (`AGENTS.md`, mailbox-safety rule 1).

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

* Hyper-V guests, named `OutlookAI-Indexed` and `OutlookAI-Unindexed`; the build VM, which has no
  Outlook, is `OutlookAI-Build` (`Testbed/README.md` section 1c). Every host script takes the name
  explicitly and none defaults to one (`Testbed/README.md` section 4a). `OutlookAI-TestVM` was the
  original, hand-built guest - retired and deleted on 2026-10-03 (Q105 (a)) - and stays in
  `Testbed/testbed.json` only as the provenance of the published measurements: do not build under
  that name. Generation, firmware, vCPU, RAM and disk size are **not recorded anywhere and are
  yours to choose**; see section 8.
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
  *(2026-09-27: recorded since - `Testbed/MEDIA.md`'s Office section and section 8 item 7 below:
  the Office Deployment Tool with `ProPlus2024Volume` on `PerpetualVL2024`, 64-bit
  (`OfficeClientEdition="64"`) - `Testbed.xml` for the script-built guests, `VoIPFabric.xml` for
  the original one - and a guest build read on 2026-09-15 of 16.0.17932.20884. `Testbed.xml` also
  sets `<Updates Enabled="FALSE" />`, the pin the last bullet asks for. What stays untested is
  whether Office MUST be x64: every guest is 64-bit by configuration, and a 32-bit Office has
  never been tried.)*
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
> *(2026-09-27: both have the SDK now - installed by the first bullet below on
> `OutlookAI-Indexed` from 2026-09-17 and on `OutlookAI-Unindexed` from 2026-09-24, `TEST-READY` on
> each (sections 4.1c and 4.2b). Still no git, no clone and no network; and the live tier still has
> not run on either, for the reasons section 4 records.)*

* **.NET SDK 10.0.401, win-x64, installed from STAGED media** by `Testbed/guest/Install-DotnetSdk.ps1`
  (`Testbed/MEDIA.md` declares the precondition). Nothing pins a feature band - there is no
  `global.json` in this repository - so any .NET 10 SDK would compile. 10.0.401 is chosen because it is what the host runs, and the host publishes the
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
  2026-09-24, in two halves since 2026-10-03** (`Testbed/README.md` section 1, steps 5b and 7c). This
  line used to say "install the add-in and let it run once", and the script-built guests never had
  it: nothing in the build installed it.

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
workstation identical. Its banner is the record. Since Q81 (2026-09-27) `OutlookAI.csproj` itself
stops both writers outside Visual Studio; the script keeps its guards, because it can build a commit
from before that change.

**On the guest, `Testbed/guest/Install-OutlookAIAddIn.ps1`, in TWO PHASES at TWO RUN LEVELS, both through the
interactive task** - decided by the maintainer 2026-10-03 (Q100, option 3):

    .\Register-InteractiveTask.ps1 -Script "& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase Install -Execute"
    .\Register-InteractiveTask.ps1 -RunLevel Limited -Script "& 'C:\OutlookAI-Q5\Install-OutlookAIAddIn.ps1' -Phase FirstRun -Execute"

Until then one `-Execute` did all six steps below from the elevated task, so the Outlook of step 5
started elevated - and an elevated Outlook never feeds Windows Search (section 8 item 22), which
broke the indexed guest's rule that every Outlook there starts unelevated. Now `-Phase Install`
(elevated, the task's default `RunLevel Highest`) does steps 1 to 4 and records what it installed and
when, never starts Outlook, and ends `INSTALLED-NEVER-RAN` (exit 2) by design; `-Phase FirstRun`
(`-RunLevel Limited`) does steps 5 and 6 and installs and writes nothing. Each refuses the other's
token, and `-Execute` without `-Phase` is refused. `ADDIN-READY` means what it meant - the state the
tests read, written by THIS start - and only FirstRun reaches it; `-Verify` compares the state with
the install record, so a state older than the last `-Phase Install` reads `INSTALLED-NEVER-RAN`.

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
5. **Outlook started once, NOT elevated** (FirstRun), over COM, headless, in a watchdogged child
   job - never quit, never killed. Before it, a preflight: installed, registered, trusted, the
   payload's build, the runtime present and not hard-disabled - or it stops with Outlook untouched,
   because without the trust entry the start would put the trust prompt on the console. While the
   job still holds Outlook it reads the started OUTLOOK.EXE's token, as
   `Start-OutlookUnelevated.ps1` does, and an elevated one is `BROKEN`.
6. **Proof, not exit codes**: `LastReconcileUtc` written AFTER that start, `Initialized` and
   `Enabled` of the right type, `LoadBehavior` still 3, nothing in Outlook's disabled list, the
   add-in connected and answering a call into it, no Claude Code registration question pending (it
   would surface as a modal dialog mid-tier), no window left on screen by the run, and the installed
   build the payload's.

**It does not disturb the index exclusion or the corpora.** The add-in's tuning service
(`Services/OutlookTuningService.cs`) writes only under HKCU - Outlook's Search key (four search-box
preferences), the Cached Mode user and policy keys (Exchange sync settings), and the PST key (a
larger file-size cap). `Set-OutlookIndexingDisabled.ps1` writes only HKLM - the Windows Search
`PreventIndexingOutlook` policy and the crawl-scope rule. The two sets are disjoint, none of the four
search values decides whether a store is indexed, and nothing in the add-in's startup path touches
an item or a store. That is read from both sources; the guest script also snapshots the exclusion
state before and after its Outlook start and says if anything moved. **The order, since the
split: the install before step 7b**, so 7b's own `-Verify` certifies the exclusion with the add-in
present, and **the first run after it** (step 7c): its Outlook is NOT elevated, and a non-elevated
Outlook adds itself to the index within a minute of starting, so on the unindexed guest the
exclusion must already be there. Then `Set-OutlookIndexingDisabled.ps1 -Verify` on either guest.

**Never executed on a guest yet.** Everything above that says "measured" was measured on the host.
**Corrected 2026-09-27: executed on both guests since, and it printed `ADDIN-READY`** -
`OutlookAI-Unindexed` on 2026-09-24 (section 4.1 step 4, `CP-09-ADDIN-READY`), `OutlookAI-Indexed`
the same day and again on 2026-09-27 (section 4.2 step 4, section 4.2b step 4.4): `NOT-INSTALLED`
first, then `ADDIN-READY` twice each time, the trust entry kept on the second run, and the index
exclusion state unchanged by it.

**The two phases RAN on a guest - `OutlookAI-Unindexed`, 2026-10-03 - and work as designed. Their
first run found a PRODUCT DEFECT: in a NOT elevated Outlook the add-in's tuning reconcile never
finishes, so on a fresh guest step 7c ends `BROKEN`, not `ADDIN-READY` (below). Fixed the same day (Q128,
the last record of this section): the same proof from `CP-08` now ends `ADDIN-READY` with no registry
step.**
Until that day their proof was the host's: `-SelfTest`, 168 assertions under Windows PowerShell 5.1
and PowerShell 7 - seven reading the script's own syntax tree - eleven rules broken on purpose in
scratch copies, each caught, and the four `Tools/Checks` guards.

*The run.* The guest restored to `CP-08-MAIL-SINK` - the checkpoint before `CP-09-ADDIN-READY`: no
add-in, no VSTO `v4R`, the tier profile the default with its POP3 password stored, the Q80 policy in,
the index excluded by the POLICY alone (that line predates the scope rule), Outlook not running. The
payload built on the host from `883ec5f` by `Testbed/host/Publish-AddInPayload.ps1` (guard 3
`UNCHANGED` over 305 host lines), staged with the pinned `vstor_redist.exe` and `883ec5f`'s guest
scripts. Every phase through `Register-InteractiveTask.ps1` in session 1 and every plain `-Verify`
over PowerShell Direct, while a READ-ONLY poller in session 0 sampled OUTLOOK.EXE every 0.5 s: its token
(TokenElevation, TokenElevationType, integrity level), command line, parent and modules. Raw logs,
one file per step below, the poller and both build-VM runs: `.work\q100-proof\` in the main checkout.

| # | What ran | Verdict, exit code, and what else was seen |
| --- | --- | --- |
| 1 | `-SelfTest` (guest), `-Verify` | 138 assertions, 0 failures (the contract section SKIPs off the repository); `NOT-INSTALLED`, 3 - `v4R` absent, `v4` 10.0.60910 |
| 2 | `-Execute` with no `-Phase`, default task | refused, 1: `REFUSING: -Execute needs -Phase since 2026-10-03 (Q100)`, both phase commands printed; nothing installed, no record |
| 3 | `-Phase Install -Execute`, default task (elevated) | **`INSTALLED-NEVER-RAN`, 2**, 66 s: the runtime installed (`v4R` 10.0.60917, 35 s), the installer 14 s, the trust entry and `install-addin-record.json` written, the exclusion state `UNCHANGED`. **No OUTLOOK.EXE at any sample, and none after** |
| 4 | `-Verify` | `INSTALLED-NEVER-RAN`, 2 - "the add-in has NOT run since" |
| 5 | `-Phase FirstRun -Execute` from the DEFAULT task; `-Phase Install -Execute` at `-RunLevel Limited` | both refused, 1, in under 4 s - `REFUSING TO RUN -Phase FirstRun: this session is ELEVATED` and `REFUSING TO RUN -Phase Install: this session is NOT elevated` - no Outlook, the record untouched |
| 6 | `-Phase FirstRun -Execute`, `-RunLevel Limited` | **`BROKEN`, 1**, 253 s. COM start 3.5 s, MAPI ok; `OUTLOOK.EXE pid 9064: token NOT elevated`; `Connect = True`, the add-in answered `GetRestartNeeded()`; `Initialized` and `Enabled` written - and **no `LastReconcileUtc` in 240 s**. Exclusion `UNCHANGED`. Outlook left running headless; it had closed by itself 3.5 minutes later |
| 7 | `-Verify -WithOutlook` at Limited, then `-Verify` | `INSTALLED-NEVER-RAN`, 2, both - Outlook had closed, so the COM half was not checked. **Wrong: the add-in had run.** Fixed in the script (below) |
| 8 | the fixed script: `-SelfTest`, `-Verify`, then FirstRun again | 151, 0 failures; `BROKEN`, 1 - "the add-in HAS run since ... registration reconcile wrote ...Mcp\LastReconcileUtc at 16:01:25Z ... Tuning\Applied records 4 of the 13 Desired values"; FirstRun `BROKEN`, 1: "tuning state (Tuning\LastReconcileUtc) written NEVER, in 240 s; registration reconcile (Mcp\LastReconcileUtc) written after 2.4 s" |
| 9 | **Control:** the five Cached Mode policy values written from session 0, elevated, equal to `Tuning\Desired` - as a GPO would set them | - |
| 10 | `-Phase FirstRun -Execute`, Limited | **`ADDIN-READY`, 0**, 8.4 s: COM 2 s, `LastReconcileUtc` 2.1 s in, `tuning walk: ... 13 of the 13`, `token NOT elevated`, `GetRestartNeeded() = True`, exclusion `UNCHANGED` |
| 11 | `-Verify` | `ADDIN-READY`, 0 - "the add-in has run since" |
| 12 | `-Phase Install -Execute` again, default task | **`INSTALLED-NEVER-RAN`, 2, over that valid state**, 12 s: "already registered: v4R 10.0.60917 - not reinstalling", "kept the existing entry ... same URL, same key", the record rewritten; no OUTLOOK.EXE |
| 13 | `-Verify`; FirstRun, Limited | `INSTALLED-NEVER-RAN`, 2 - "LastReconcileUtc '...16:20:11Z' is from before" the install at 16:20:55Z; then `ADDIN-READY`, 0, 5.7 s, `GetRestartNeeded() = False` |
| 14 | `Testbed/host/Restart-Guest.ps1 -Execute`, then `Set-OutlookIndexingDisabled.ps1 -Verify` | restarted in 22 s (Outlook had closed by itself); **`UNINDEXED`**: two readings 10 minutes apart, `mapiRows=0 outlookRowsTotal=0`, the catalog `IDLE` with 430 other items - after three NOT elevated Outlook starts, with the policy-only caveat that line carries |
| 15 | `Restore-VMSnapshot CP-13B-LIVE-GREEN`, saved, lease released | no checkpoint kept from this run |

*That the add-in loaded in the NOT elevated Outlook, independently of the script.* For each of the four
FirstRun starts the poller read OUTLOOK.EXE as `TokenElevation=0 TokenElevationType=3` (Limited) at
integrity `0x2000` (Medium), command line `OUTLOOK.EXE -Embedding`, parent `svchost.exe -k DcomLaunch` -
started by COM, not by the task - with `VSTOLoader.dll` and `vstoee.dll` mapped into it within 0.8 s.
`Addins\OutlookAI\LoadBehavior` stayed 3 after every run (Outlook sets 2 on a load that fails), and
Outlook's own `...\16.0\Outlook\AddInLoadTimes` gained an `OutlookAI` value at the first start and then
recorded loads of 829 ms and 625 ms. `OutlookAI.dll` itself never appeared in the module list; a
managed assembly need not.

*The defect, read on the guest and in the source.* `OutlookTuningService.Reconcile` walks its catalog
in order: four `search.*` values in Outlook's user Search key, then D25's five `caching.policy.*`
values under `HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode`, then two user Cached
Mode values and two PST size values, and `LastReconcileUtc` last. `HKCU\Software\Policies` grants the
user ReadKey only - Administrators and SYSTEM hold FullControl, and a filtered token does not use the
Administrators group (its ACL read on the guest; the maintainer's workstation's has the same shape).
So in a NOT elevated Outlook the first policy write throws, the reconcile's catch-all swallows it to
the debugger, and nothing after it runs: `Tuning\Applied` held the four search values and nothing
else, the policy key and the user Cached Mode key were absent, and the reconcile's own bookkeeping -
updating `RestartNeeded`, writing `PolicyConflicts` and `LastReconcileUtc` - never ran. The class's own summary says "Everything is
HKCU - no elevation is ever required"; for the Policies hive that is not so. **What hid it:** the
single elevated `-Execute` this split replaced - its Outlook could write there, and every later
start, elevated or not, found the five values in sync. Both live guests were installed by that
`-Execute`, so they should carry the five values too - not read on them in this run. **What it means for users:** the installer is
per-user and Outlook runs NOT elevated, so on any machine where an administrator, a GPO or an earlier
elevated Outlook has not already set those five values, the reconcile never completes - the user
Cached Mode and OST size values never apply either, and `outlook_health` reports
`tuning.lastReconcileUtc` null. The maintainer's workstation carries the five values and a fresh
`LastReconcileUtc` (read 2026-10-03), so it is not affected; who set them there is not recorded.
**The control (step 9) isolates it:** the same build, guest, token level and command, and only the
five values differ - `BROKEN` without them, `ADDIN-READY` with them. It is a control, not a build
step: setting them in the testbed would hide the defect exactly as the elevated `-Execute` did.

*The script, fixed from this run (step 8).* `-Verify` judged "run since the install" by the tuning
state alone. It now also reads the add-in's second startup marker, `Mcp\LastReconcileUtc` - written by
the registration reconcile at every start, failed or not - and an add-in that started since the
install without finishing a tuning reconcile is `BROKEN`, with when it started and how far its walk
got (`Tuning\Applied` against `Tuning\Desired`, printed as `tuning walk`). FirstRun watches that marker
during its wait, so a tuning state that never comes says which failure it was, and its timing is no
longer read after the 240 s wait (step 6 printed "registration reconcile after 245.4s"). `-SelfTest`:
187 assertions, 0 failures, on the build VM under Windows PowerShell 5.1, with all 21 scripts'
self-tests passing (`Testbed/host/Invoke-TestsOnBuildVm.ps1 9bfc135 -SkipSuite`); 151 on the guest,
whose copy has no repository for the contract section.

**Not settled by it:** `ADDIN-READY` on a fresh guest, which waited on the defect - settled by the Q128 run
below; the indexed guest -
that an unelevated first run there feeds the index and leaves it `INDEXED`; `-Verify -WithOutlook`
against a running Outlook, which had closed each time before the attach was tried; and the two live
tests on a guest installed this way.

**The defect, fixed (Q128), and the proof again from `CP-08` - `OutlookAI-Unindexed`, 2026-10-03.**
Decided by the maintainer the same day (Q128): *"Allow reading it and changing it from the gui. If
the change requires admin and the user is not admin generate a uac prompt."* What changed:

* **The reconcile never stops early** (`Services/TuningReconciler.cs`, split out of
  `OutlookTuningService` so the test project can pin it - T1 `TuningReconcilerTests`). A value Windows
  refuses to write for lack of rights is skipped and listed in `Tuning\NeedsAdministrator` - a REG_SZ
  of `;`-joined entry ids, `Services/AddInServerContract.cs` - which `outlook_health` reports as
  `tuning.needsAdministrator`; every other value is still applied; any other failed write is skipped
  too; and the bookkeeping - `RestartNeeded`, `PolicyConflicts`, `NeedsAdministrator`, then
  `LastReconcileUtc` - is written in a `finally`, each write on its own.
* **OutlookAI Settings shows the five policy values**, one row each: current, desired (a list - the
  choice is stored at once as the desired value), and whether it is in effect or needs an
  administrator. **Apply as administrator** starts `OutlookAI.PolicyWriter.exe` through ShellExecuteEx
  `runas` - one UAC prompt for every value that needs it.
* **The helper** (`PolicyWriter/`) is the one program in the product that runs elevated, and a
  privilege boundary: it accepts `--sid <SID> --office <major>` and some of the five value names with
  the values the dialog offers (`Services/CachedModePolicy.cs`, `Services/PolicyWriterRequest.cs`),
  refuses anything else whole before writing, refuses unless the SID is the user of the process that
  started it - read from the session manager, in the helper's own session - and writes
  `HKEY_USERS\<SID>\Software\Policies\Microsoft\Office\<major>\Outlook\Cached Mode`, never HKCU: a
  standard user who approves the prompt with an ADMINISTRATOR's credentials gets a helper running as
  that administrator, whose HKCU is the administrator's hive. It never loads a hive. It is built by the
  add-in's own build (a ProjectReference), listed with its SHA-256 in the signed
  `OutlookAI.dll.manifest`, flattened beside `OutlookAI.dll` and installed to `{app}` by the
  installer's `publish\*` rule; `requireAdministrator` in its manifest, so it cannot run with a
  filtered token at all.

*The run.* `CP-08-MAIL-SINK` restored and started; the payload built on the host from `94f115f` by
`Testbed/host/Publish-AddInPayload.ps1` (guard 3 `UNCHANGED` over 305 host lines, no build warning;
`publish\OutlookAI.PolicyWriter.exe` 26,624 bytes, SHA-256 `AEC67A1B...`, a `<file>` with that hash in
the signed manifest) and staged with the pinned `vstor_redist.exe` and `94f115f`'s guest scripts. Every
phase through `Register-InteractiveTask.ps1` in session 1; in session 0, read-only, a poller sampled
OUTLOOK.EXE every 0.5 s and a WMI process-start trace recorded every `OutlookAI.PolicyWriter.exe` and
`consent.exe` with its PARENT process id. Raw logs, one file per step, the harness and the build-VM
runs: `.work\q128-proof\` in the main checkout (the second pass in `second-pass\`).

| # | What ran | Verdict, exit code, and what else was seen |
| --- | --- | --- |
| 1 | `-SelfTest` (guest), `-Verify` | 158 assertions, 0 failures (the contract section SKIPs off the repository); `NOT-INSTALLED`, 3 |
| 2 | `-Phase Install -Execute`, default task (elevated) | **`INSTALLED-NEVER-RAN`, 2**, 99.6 s: `v4R` 10.0.60917 installed, the installer, the trust entry and the record written; no OUTLOOK.EXE at any sample |
| 3 | `-Phase FirstRun -Execute`, `-RunLevel Limited` | **`ADDIN-READY`, 0**, 16.6 s, **with no registry step of any kind**: COM 6.4 s, `LastReconcileUtc` 7.2 s after the start, `token NOT elevated`, `Connect = True`, `GetRestartNeeded() = True`, `tuning walk: 8 of the 13`, and the note *"the add-in skipped 5 value(s) it may not write without an administrator, and finished its reconcile"* - `Tuning\NeedsAdministrator` = the five `caching.policy.*` ids; the policy key absent; the user Cached Mode values 0/0 and the PST values 102400/96256 written; exclusion `UNCHANGED` |
| 4 | OutlookAI Settings, opened through the add-in's automation hook (`OpenSettings`) and read with UI Automation, Limited | the five rows: current `(not set)`, desired `All (0)`, `Not used (0)`, `On (1)` three times, state `Needs administrator`; **Apply as administrator...** enabled. (A second Outlook start meanwhile: its startup reconcile wrote nothing, cleared `RestartNeeded`, and listed the five again) |
| 5 | the helper, `CreateProcess` from a Limited task, a valid request | did not start: *"The requested operation requires elevation"*; nothing written |
| 6 | the helper from the elevated task: 14 requests it must refuse | 12 refused **exit 2** (arguments): none at all, a sixth name, a name in other case, `SyncWindowSetting=2`, `DownloadSharedFolders=2`, `0x0`, `--office 18.0`, `--hive HKLM`, a value twice, no `--sid`, `S-1-5-18`, `<SID>_Classes`; 2 refused **exit 3** (user): this machine's built-in Administrator (`...-500`) and `S-1-5-21-1-2-3-4` - *"the process that started this helper runs as ...-1000"*. Afterwards the policy key in no loaded hive under `HKEY_USERS` |
| 7 | the helper elevated, valid: `SyncWindowSetting=0 SyncWindowSettingDays=0` for vmadmin's SID | **exit 0**, *"WRITTEN under HKEY_USERS\S-1-5-21-...-1000\Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode"*; of every loaded hive under `HKEY_USERS` only that one holds the key, both values REG_DWORD |
| 8 | **Apply as administrator...** pressed (UI Automation), the UAC prompt ended from session 0 by `Stop-Process` | *"Windows did not start the administrator helper: Unknown error (0xffffffff) (-1)."* ShellExecuteEx hands back the consent process's own exit code, and `Stop-Process` ends it with -1 - not what a user's No does. Nothing changed |
| 9 | the same, the prompt ended with `ERROR_CANCELLED` (1223) - the code `consent.exe` ends with on No | **"Cancelled at the administrator prompt. Nothing was changed."**, 3.5 s after the press, in the dialog's secondary colour, not as an error; no helper started; the three values still `Needs administrator` |
| 10 | the same with `ConsentPromptBehaviorAdmin` 0 ("elevate without prompting") for the step, 5 restored after | **"The values were written to your Outlook policy settings. Restart Outlook for them to take effect."** 0.6 s after the press, and the restart line. The trace: **`OutlookAI.PolicyWriter.exe` started with parent `OUTLOOK.EXE -Embedding`** - AppInfo makes the requester the elevated process's parent, which the helper's user check relies on. All five rows `In effect`, the button disabled, `NeedsAdministrator` empty, `Applied` 13 of 13 |
| 11 | `-Phase FirstRun -Execute` again, Limited | **`ADDIN-READY`, 0**, 8 s: `LastReconcileUtc` 2.4 s in, `tuning walk: 13 of the 13`, `needsAdministrator=` empty, `GetRestartNeeded() = False`, token NOT elevated, exclusion `UNCHANGED`; `-Verify` `ADDIN-READY`, 0 |
| 12 | a probe: one byte appended to `{app}\OutlookAI.PolicyWriter.exe`, then FirstRun | `ADDIN-READY`, `Connect = True`: **the VSTO runtime does not check that file's manifest hash when it loads the add-in** - the signed manifest records the helper, it does not guard it on disk |
| 13 | `Restore-VMSnapshot CP-13B-LIVE-GREEN`, saved, lease released | no checkpoint kept from this run |
| 14 | a second pass from `CP-08`, the same payload: `-Phase Install`, then FirstRun at Limited | `INSTALLED-NEVER-RAN`, 2, in 115.6 s; then `ADDIN-READY`, 0, in 14.4 s, the five listed again |
| 15 | in OutlookAI Settings, **12 months** chosen in `SyncWindowSetting`'s list - by the list's own `CB_SETCURSEL` and the `CBN_SELCHANGE` a pick sends - then **Apply as administrator...**, prompt-free consent | the row read desired `12 months (12)`, still `Needs administrator` (the reconcile tried and was refused); then *"The values were written..."* in 1.3 s, all five `In effect`, `SyncWindowSetting` 12 in the registry, the helper's parent `OUTLOOK.EXE` again |
| 16 | **All** chosen again, Apply as administrator | that row alone `Needs administrator`, then written: `SyncWindowSetting` 0, all five `In effect`; restored to `CP-13B-LIVE-GREEN`, saved, lease released |

**Not settled by it:** a standard user approving the prompt with ANOTHER account's credentials - every
guest has one account, so step 10's requester and helper were the same user; the helper's check and
its `HKEY_USERS\<SID>` target are what make that case right, pinned by T1 `PolicyWriterRunTests`, not
measured. The prompt itself was never clicked: steps 8 to 10 and 15 to 16 stood in for it from
session 0, and steps 15 and 16 chose from the list by its own messages, not by a mouse. And, as before,
the indexed guest and the two live tests on a guest installed this way.

**Checkpoint `CP-03-OUTLOOKAI-INSTALLED` once `-Phase FirstRun` prints `ADDIN-READY`.** `CP-05-ADDIN-TRUSTED`
is no longer a separate manual step - the trust entry is part of the scripted install - and the name
survives only as the hand-built guest's history.

### 2.4 The two Windows accounts - NOT NEEDED, skip to 2.5

**SUPERSEDED 2026-09-15. Do not do this step.** Each guest has one Windows account, `vmadmin`,
created by the answer file, and index state is a property of the **machine** rather than of an
account - see 1.1a. Creating a second account, and installing the repository, the SDK and a built
server exe under it, is work with nothing behind it.

The original text is kept below because it explains a design the old guest, `OutlookAI-TestVM`,
was built towards (it was retired and deleted on 2026-10-03), and because if the two-guest
arrangement is ever collapsed back to one machine this is what it would have to become again.

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
>    only route into session 1, `Testbed/guest/Register-InteractiveTask.ps1`, ran everything at
>    `RunLevel Highest` - still its default; `-RunLevel Limited` exists since the same day, and the
>    live tier runs at it (`Testbed/README.md` section 4c). Controlled A/B and the evidence: section
>    8 item 22.
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
> 4. **`OutlookAI-Unindexed` - DONE 2026-09-24, exactly as this item prescribed, and checkpointed
>    `CP-14-EXCLUDED-BY-SCOPE-RULE`** (child of `CP-10-IDENTITY-ACCOUNT`, which it was restored to
>    first: the checkpoints after it carry faulty populations - section 4.1).
>    Before: the policy only - `PreventIndexingOutlook = 1`, the service reporting the scope
>    `included=False reason=UNKNOWNSCOPE`, no Outlook row, and `-Verify` printing the policy-only
>    caveat. The steps: `Set-OutlookIndexingDisabled.ps1` and `SearchCrawlScope.cs` staged into
>    `C:\OutlookAI-Q5\`; `Testbed/host/Restart-Guest.ps1 -Execute -CancelLogonPrompt` (Outlook was not
>    running; the guest restarted without force, 27 s); `-Execute` elevated over PowerShell Direct -
>    the user EXCLUDE rule, and the service answered `included=False reason=USER` at once, though
>    that guest has NO search root for the scope (the indexed guest's exclusions all had one); `no
>    Outlook row in the catalog: nothing to purge`; the policy already 1; `WSearch` restarted. Then
>    `-Verify`, twice (3 minutes apart, and again 1 minute apart immediately before the checkpoint):
>    **`UNINDEXED`, reason `USER`, no caveat**, 0 Outlook rows, the catalog IDLE with 430 other items,
>    `WorkingSetRules\18 include=0` in the registry, Outlook closed.
>    `.work/aa5e-2026-09-24-q69b-runlevel-unindexed/`. Nothing else was run there - no
>    `Start-OutlookUnelevated.ps1` - and with the rule in, that guest's index state no longer depends
>    on how its Outlook is started (items 1 and 3). **The live tier's populations, the SDK and the
>    settings file are NOT in `CP-14`** - they were built after `CP-10` - so steps 8a to 9 of
>    `Testbed/README.md` section 1 run again from it.
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
> *(Answered 2026-09-24 by the Q69 block above, noted 2026-09-27: the policy on its own removes
> nothing - the scope stays IN and every row stays (row P). It is the crawl-scope EXCLUDE rule that
> takes a scope out, and the indexer then purges the rows itself within minutes, unless a service
> restart interrupts the purge, which it does not redo (rows U, R and N). Excluding before the
> corpus is built still keeps the question from arising; on a guest already crawled, `-Execute`
> now writes the rule, waits for the purge and only then restarts (row N).)*
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
`MAPIAdminProfiles` both succeeding; cause not established). *(Corrected 2026-09-27: the cause was
established on 2026-09-24 and it was not Office - the interop declared the wrong IID, and a
`QueryInterface` for Microsoft's `IID_IProfAdmin` succeeds on this build; section 8 item 9 (b),
`Testbed/guest/OutlookMapiInterop.ps1`'s REOPEN section. Nothing was rebuilt on it.)* The
rewrites then ran on `OAI-UNINDEXED` on **2026-09-24**, each from a fresh restore of
`CP-05-CORPUS-B-CLEAN-UNINDEXED`:

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
  `AGENTS.md` mailbox-safety rule 1 remains the authority: item mutation goes through the tested
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
> is measured broken (2026-09-16, `E_NOINTERFACE` on `IID_IProfAdmin` - on the interop's wrong IID,
> not on Office, as established 2026-09-24; section 8 item 9 (b)). So it does now modify a
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
  will write anything. In the TIER profile, that is. **Attached to the account-less corpus profile
  the same PST passes - ANSWERED on `OutlookAI-Unindexed`, 2026-09-24** (section 4.1, step 6): the
  hub, the bystander and the identity store were each "accepted as a corpus target (bound profile:
  'CorpusProfile')", `profile accounts: 0`, `delivering into this store: 0`.
* **A PST attached with `AddStoreEx` has almost no default folders, and nothing may ask Outlook for
  one.** Measured the same day: the bystander held Deleted Items and nothing else - no Inbox, no Sent
  Items, no Drafts, no Junk Email - and `Store.GetDefaultFolder` then CREATED Drafts and Junk Email in
  it and, for the Inbox, handed back the PST's hidden non-IPM root. Only a store Outlook mints as a
  profile's default (the hub, `C:\OutlookAI-Tier\Outlook.pst`) has the full set. The generator no
  longer asks: it finds a default folder only through the store's own designations (the product's
  non-creating `SpecialFolders.Resolve`), and gives a store that lacks one a visible stand-in
  (section 3b). For the bystander that absence is its point (section 1.3); for the identity store it
  is a defect with a decision attached (section 3b, "The identity store has no Inbox").
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
3b). The identity store gets `--population identity` once section 2.8b has built it - and, since
2026-09-24, once it has a real Inbox (section 3b).

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

> **THE ROUTE SINCE 2026-09-27 (Q87 (a)): THE IDENTITY STORE IS MINTED, AND THE ACCOUNT DELIVERS
> INTO A REAL INBOX.** Built on `OutlookAI-Unindexed` from `CP-09-ADDIN-READY` twice - once by hand
> as the measurement, once as the script's phases - checkpoint `CP-10B-IDENTITY-REAL-INBOX` (section
> 4.1b). A PST that `AddStoreEx` creates has no Inbox (the boxes below); one Outlook MINTS as a
> profile's default store has every default folder, so the identity store is minted in a throwaway
> account-less profile and then attached:
>
> ```
> .\Add-IdentityAccount.ps1 -Phase Mint -Execute           # Outlook CLOSED; no ImportPRF pending
> <start OUTLOOK.EXE /PIM IdentityMint, ~90 s: 'Outlook Today', no dialog>
> .\Add-IdentityAccount.ps1 -Phase CaptureMint -Execute    # the one new default store: Inbox, 0 items
> .\Rename-OutlookStore.ps1 -StoreFilePath '<minted>' -DisplayName identity@vm.invalid -Execute
> <quit: Testbed/host/Restart-Guest.ps1 -VMName <guest> -Execute>
> .\Add-IdentityAccount.ps1 -Phase Import -Execute         # skip if the account exists
> <start Outlook on the tier profile>
> .\Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName identity@vm.invalid `
>                           -Path '<minted>' -Execute
> .\Add-IdentityAccount.ps1 -Phase CaptureStore -Execute
> <quit: Restart-Guest.ps1 -VMName <guest> -Execute -CancelLogonPrompt>
> .\Add-IdentityAccount.ps1 -Phase Bind -Execute
> .\New-TierProfile.ps1 -StoreSinkPassword -Execute
> <start Outlook on the tier profile>
> .\Add-IdentityAccount.ps1 -Phase Verify -TrySmtpAddress
> ```
>
> `<minted>` is `C:\OutlookAI-Tier\Outlook Data File - IdentityMint.pst` on this build - Outlook's
> name for it, read off the guest by CaptureMint and recorded in `C:\OutlookAI-Tier\identity-mint.json`,
> from which CaptureStore, Bind and Verify take it. A guest whose `ImportPRF` is already pending
> (`CP-09-ADDIN-READY` on both guests) completes that import first - one start of the tier profile,
> then a quit with `-CancelLogonPrompt` - because `-Phase Mint` refuses while it is pending: the
> `/PIM` start would process it too, and that is unmeasured. What it produced, both times: the
> minted store's folder mask `0xFF`, its Inbox named, visible and designated, 14 folders and no
> item; CaptureStore from the tier profile `Inbox designated=True, EntryID 24 bytes, name 'Inbox',
> visible=True; Drafts designated`; Verify `the identity account delivers into 'Inbox' (visible=True,
> PST node id 0x8082)`, two accounts on two distinct stores, both `SmtpAddress` reads, no logon
> dialog - and the sink's log `read USER identity`, `read PASS any-value`. It leaves the mint
> profile behind, and cannot replace an `AddStoreEx` identity store that is already attached under
> the same name (the script's banner says both).
>
> **Two preconditions the indexed guest added** (section 4.2b, `CP-09C-IDENTITY-REAL-INBOX`, the same
> lines from `CP-08B`): run `Set-OutlookProgrammaticAccess.ps1 -Execute` BEFORE the route - without
> it CaptureStore's COM read met the Object Model Guard prompt about two minutes after a boot and
> blocked behind it, the phase having no deadline; and on that guest start every Outlook NOT
> elevated and run every phase that attaches through `Register-InteractiveTask.ps1 -RunLevel
> Limited` - the `/PIM` start included, as a Limited job running `Start-Process OUTLOOK.EXE /PIM
> IdentityMint`, because `Start-OutlookUnelevated.ps1` opens only a profile that exists.
>
> **The boxes below are the route as it was built on 2026-09-24, kept as the record: both guests'
> identity accounts from their `CP-10` on are bound to that PST's hidden root.**

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
> **Still to do on this machine, and NOT done by that script:** the **signature** below; the
> settings-file declaration below; and reading `Account.SmtpAddress` over COM, which the Object
> Model Guard blocked on these guests (section 8 item 23). **Two of the three are done since**
> (section 4.2): `Set-OutlookProgrammaticAccess.ps1` made the `SmtpAddress` read work
> (`CP-12-PROGRAMMATIC-ACCESS`), and `Testbed/guest/Set-AccountSignature.ps1`, run with the fixed
> server on 2026-09-27, put `New Signature` = `Identity` on the identity ACCOUNT's own entry
> (`00000004`), where the product read it back (`CP-15-SIGNATURE-SUITE-STAGED`). The settings file
> is not rendered yet. The rest of this section is the specification the script was built to.

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
>   **Fixed in the product since** (the merge "manage_signature writes to the mail account, never
>   to a data file named like an address"): on `OutlookAI-Indexed`, 2026-09-27, the fixed server
>   wrote the ACCOUNT entry and left the data-file entry alone (section 4.2). This guest's
>   checkpoints from `CP-10-IDENTITY-ACCOUNT` on still carry the misplaced value.
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

In section 2.10's settings file the delta is one name, in one list - the watched one. The identity
store is in neither the indexed list nor the bystander list:

```
"expectedStoreDisplayNames":  [ "test@vm.invalid", "bystander@vm.invalid", "Corpus A", "identity@vm.invalid" ],
"indexedStoreDisplayNames":   [ "test@vm.invalid", "bystander@vm.invalid", "Corpus A" ],
"bystanderStoreDisplayNames": [ "bystander@vm.invalid", "Corpus A" ],
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

**On the indexed guest the corpus is at least 160,000 items - decided 2026-09-24 (question E,
option (a), keeping the proposed minimum).** `Testbed/testbed.json` records it as
`corpusIdConvention.minimumItemCount`, and `Testbed/host/New-LiveTestSettings.ps1` refuses to render
that guest's settings while the recorded corpus is smaller. Build it with `--count 160000`: at the
recorded 24.8 items/s that is about 1 h 50 min, plus the indexer's crawl. **Built 2026-10-03 (section
4.2d): 1 h 21 min at 33.0 items/s, through `PostAsNote` into a new store that is not the corpus
profile's default, with an unelevated Outlook pushing every item into the index as it went** - and at
that size `Set-OutlookIndexingDisabled.ps1 -Verify` needed a longer timeout on its row count, which it
now has. The reason is the same
day's question B, option (c): every index-tier latency bound is now timed against the LARGEST
indexed store (`T2/LiveLatencyTarget`, pinned by `T1/LatencyTargetTests`), which on a guest is this
corpus - and against a 20,000-item store an idle VM meets a 2-second bound by construction, where
the maintainer's profile, which the bound was set against, holds about 160,000.

**Attach the corpus store to the tier profile while it is still EMPTY - decided the same day
(question C, option (a)).** It is built in the account-less profile and read by the tier in the
tier profile, so it is mounted in both, exactly as a population's store is (section 2.6's
addition, section 3b's order): create it in the corpus profile, attach it to the tier profile by
path with `Testbed/guest/Add-OutlookPstStore.ps1` - names byte-identical - and only THEN build.
Never the other way round: a store that already holds 160,000 items is not attached by script,
whatever made them.

Repeat for Corpus B under the other Windows account, with **a different `--corpus-id` and a
different manifest path**. Whether the two corpora should share a seed and anchor is not
settled; sharing them makes the two stores directly comparable, which is probably what you
want.

**One corpus id per guest, and the ids are assigned, not invented at the keyboard**:
`vm-indexed` on `OutlookAI-Indexed`, `vm-unindexed` on `OutlookAI-Unindexed`, and `vm2` stays
`vm2`, reserved, although that corpus is gone - deleted with its guest, `OutlookAI-TestVM`, on
2026-10-03. The manifest is `corpus-<corpusId>.jsonl`, so `--corpus-id` and
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
> **That was done for `OutlookAI-Unindexed` on 2026-09-24 (section 4.1, step 8), and only there.**
> Since the same day a guest's section must also name its hub population's manifest,
> `hubPopulationManifestPath` - the renderer refuses one that does not. The renderer **never writes
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
  "hubPopulationManifestPath": "C:\\OutlookAI-Q5\\corpus-hub-indexed.jsonl",
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
| `hubPopulationManifestPath` | **A test guest's field (2026-09-24).** The absolute path of the hub population's manifest, `corpus-<id>.jsonl` for the hub id `Testbed/testbed.json` assigns the guest. `Testbed/guest/Reset-HubPopulation.ps1` rebuilds the hub from it before every run, and `LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier` reads its anchor and FAILS on a hub too old to catch a local-time frontier with, naming that script. Blank is refused; absent - the maintainer's machine, whose hub is real mail - leaves the test as it was. The renderer requires it on every guest. | on a test guest |
| `expectedStoreDisplayNames` | The **WATCHED** list: every store the tier profile mounts. The count tripwire censuses it (with the bystander and delegate lists unioned in); the identity-draft grant, `list_accounts` exactness, `outlook_health`'s reachability check and the archive-resolution test read it. Include the hub. | **yes** |
| `indexedStoreDisplayNames` | The **INDEXED** list: the stores the index-tier tests measure, in order - the hub first, the bystander second, the corpus last (section 1.3 says why the order is read). Every entry must be watched; never a delegate mailbox; must include the hub unless empty. **Absent means "the same as the watched list"**, which is what every file written before the split meant. Empty only on a `Portable` machine, and there every index test refuses rather than iterate nothing. | no |
| `bystanderStoreDisplayNames` | Stores the write allowlist must **refuse**. Declaring a store here is sufficient to have it censused - the census adds declared bystanders back in - so leaving one out of `expectedStoreDisplayNames` does not refuse the tier; name it there too all the same, because `outlook_health` and `list_accounts` read only that list, and the renderer holds a test guest to it. Both corpus stores belong here, and so does the bystander. Never the hub. | no |
| `expectedDelegateStoreDisplayNames` | Delegate/shared mailboxes. Watched, never written, folder hierarchy allowed to come and go. Empty here. | no |
| `probeTerm` | A word proven to hit this machine's search index, in every indexed store and in one text attachment of the hub. On a test guest it is the generator's `CorpusPopulation.ProbeTerm`, `invoice`; empty on a guest with no index. | Production only |
| `subjectOnlyProbe` | Coordinates of a population whose term is in the subject and not the body. Four fields, all or none. On the indexed test guest it is the hub population's Notices folder, and three of the four values are generator constants (`corpus-plan --population hub` prints them); only the store - the hub - is read off the guest. Absent on a guest with no index. | Production only |
| `delegateNestedFolderProbe` | A delegate folder Outlook nests and the index publishes flat. | never |
| `corpus` | Where the measurement corpus is and what it was generated from, so the tier can prove it is still measurable. Six fields plus optional `windowDays`. | no, all or none |
| `mailSink` | Loopback submission and retrieval endpoints. **Absent is AMBIGUOUS and that is a known hazard: it means EITHER this machine has real transport OR it has none at all.** The maintainer's machine is the first case. The testbed guests were the second (decided 2026-09-15, no sink) - **corrected 2026-09-27: since the 2026-09-24 decision (Q71, section 1.4) they get a loopback sink, and a guest declares it here - `127.0.0.1` on 25 and 110, as both guests' sections of `Testbed/testbed.json` now do - once `Testbed/guest/Install-MailSink.ps1 -Verify` has reported `SINK-READY` on it.** A guest without a verified sink still leaves the block out and is still the second case, and the settings file still cannot tell the two apart - so a guest with no transport reads exactly like a machine with perfect transport. Section 1.4 has both decisions. | no, all or none |

A block that is present must be **complete**: three fields out of four reads as configured and
behaves as absent, which is the exact silence these checks exist to remove.

`windowDays` is how the machine declares which measurement windows it actually asks about. Left
empty it means all of them, including the one-day window - which forces a rebuild every day.
Name the windows your tests use. The indexed test guest names 30 and 60 days, not the example's
7, 30 and 60 - none of its live tests asks its corpus a 7-day question, and declaring one made the
tier refuse a week after every rebuild (decided 2026-10-03, section 4.2d).

The same file is read by the remediation console's `audit`/`refile`/`purge`/`dedupe` verbs,
which require the hub to appear in `expectedStoreDisplayNames`. The `corpus-*` verbs do not read
it at all; they take everything on the command line.

### 2.11 Checkpoints

Names in use: `CP-01-WIN-CLEAN`, `CP-02-INSTALLER-STAGED`, `CP-03-OUTLOOKAI-INSTALLED`,
`CP-04-OFFICE-GOLD`, `CP-05-ADDIN-TRUSTED`, `CP-06-PRE-CORPUS`. Take another after the corpus
and another after the sink and dummy account exist, because those two are the steps most likely
to need redoing.

The build VM, `OutlookAI-Build`, has three (sections 4.3 and 4.3a): `CP-01-WIN-CLEAN`,
`CP-02-SDK-TEST-READY` - the base until Q126 (a), in W. Europe Standard Time - and its child
`CP-03-SDK-TEST-READY-UTC`, the same machine in UTC, which `Testbed/host/Invoke-TestsOnBuildVm.ps1`
restores before and after every run. Both bases were taken RUNNING, because a running checkpoint
resumes in seconds. Never take a checkpoint of that VM by hand while a run might be using it - hold
its lease, as section 4.3a did - and never delete `CP-03-SDK-TEST-READY-UTC` without taking its
replacement in the same sitting: the runner refuses without it. `CP-02` is kept: a branch whose
runner still names it runs from it, in W. Europe.

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

**Added 2026-09-24 for Q70; generator v2 the same evening. v1 RAN ON A GUEST ONCE AND WAS NOT CLEAN;
v2's PROBES RAN ON ONE, AND NO v2 POPULATION IS BUILT.** `OutlookAI-Unindexed` built all three
populations with v1 (section 4.1, step 6): every build exited 1, and the four defects it found are
what v2 fixes - "What v2 changed", below. On 2026-09-27 v2 probed the hub and the bystander on the
same guest (section 4.1a): placement, dates and enrichment verified on both, and **the undated items
refused on both** - an appointment, a contact or a task saved into a PST is DATED, and Outlook will not
remove the date. That is open, for the maintainer - "The undated kinds are dated in a PST", below.
The generator half is pinned on the host - `T1/CorpusPopulationTests` (79 cases), `T1/CorpusDefaultFolderTests`
(16), the placement cases of `T1/CorpusGeneratorTests`, `T1/CorpusProbeResidueCensusTests`,
`T1/CorpusUndatedTableTests` and `T1/CorpusUndatedWritePathTests`, no Outlook - and every step below
that opens a store is guest-only. "What only a guest can answer" lists what the probes settled.

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
| `hub` | the hub, `testHubStoreDisplayName` | 56 | Four conversations of four, alternating Inbox and Sent Items, whose newest member is the newest item in the store - one minute before the anchor. Sixteen received and six sent singles; six items in `Inbox/OutlookAI-Corpus-Folder-Projects`. Eleven attachments - PNG, `.ics`, `.eml` and text, and one mail carrying three - with the probe term `invoice` in one text attachment and in its parent's body. Read and unread mail. The SF-6 subject-only population: twelve items in `Inbox/OutlookAI-Corpus-Folder-Notices`, all from `Noticebot Relay <noticebot@alerts.invalid>` and nobody else, `bulletin` in every subject and in no body. **On the INDEXED guest, plus twelve undated CONTACTS, ordinals 57-68, in its Contacts folder - 68 items** (Q98 (f), below). (Generator v2 also defines twelve undated items of three kinds - four appointments, four contacts, four tasks at 57-68 - **switched off since 2026-10-03**, Q98 (a).) For the tests that read, page, cap or walk the hub; the index tests that read the first indexed store; SF-6; the attachment-kind recall; the conversation walk; the staleness frontier. |
| `bystander` | the plain bystander | 300 | Six conversations of three; 160 received and 50 sent singles spread over two years, one in eight received with an attachment; two populated subfolders, `-Projects` (40) and `-Suppliers` (32), of the folder its received mail is filed in. Every folder inside the census identity budget. **On the INDEXED guest, plus 42 undated CONTACTS, ordinals 301-342 - 342 items** (Q98 (f), below). (Generator v2 also defines 42 undated items of three kinds, fourteen each - switched off, Q98 (a).) It has no Inbox, Sent Items, Calendar, Contacts or Tasks and gets none: each folder it is built into is a visible **stand-in** under the store root - `OutlookAI-Corpus-Folder-6` and `-5`, and on the indexed guest `-10`, a contacts folder, for the contacts (`-9` and `-13` only while version 2's full set is switched on). For the count tripwire's item-by-item path and the exclude-subfolders measurement. |
| `identity` | the identity account's delivery store (section 2.8b) | 8 | Five received, three sent, all to or from its owner - so the index knows the store exists and `outlook_health` does not report it missing. Its received items go into the store's real Inbox, since the store is minted (Q87 (a), 2026-09-27). |

**The undated items (decided 2026-09-24, question A, option (a)).** The index tier sorts by
`System.Message.DateReceived DESC`, and a store-scoped search admits every item class since gap B3, so
where the provider sorts a row with NO received date decides whether such rows crowd mail out of a
`TOP n` - which is what the three `LiveOrderKeyCollationTests` measure, and on a store of dated mail
only they measured nothing (`no-undated-rows-in-sample`). Every one of them uses a kind filter that
admits every class, so appointments, contacts and tasks are what they read. **The decision named
unsent drafts too, and there are none - a deviation, with its reason:** measured on
`OutlookAI-Unindexed` the same day, a new unsent mail item's first save is filed in the profile's
DEFAULT store's Drafts, whichever store's folder created it, and a population is never built into the
default store - so a draft could only be made in one by writing into another store first, which is
the defect below. The undated kinds are never unsent, and the undated probe checks, per kind, that its
save stayed in the target store before any is built. **The deviation is accepted - decided 2026-09-27
(Q89 (a)): no drafts among the undated items, and nothing changed in the code.**

**And the three kinds turned out not to be undated in a PST** - measured on `OutlookAI-Unindexed`,
2026-09-27; "The undated kinds are dated in a PST", below, has the lines. **So they are switched off -
decided 2026-10-03 (Q98 (a)):** the populations are built without them, behind one plan parameter,
`CorpusPlanOptions.IncludeUndatedItems`, which defaults to off and which no command-line option sets.
The undated items are each population's LAST ordinals and every other item is the same either way, so
a population without them is a prefix of one with them, under the same shape key - T1 pins both, and
keeps the undated probe, the census's undated checks and their tests running on plans that switch
them on. They come back only if the measurement on the indexed guest - Q98 (f) - says it is worth it.

**It said so for CONTACTS alone - decided on the maintainer's behalf 2026-10-03 (Q98 (f); D62 of
`Docs/overnight-review-2026-10-03.md`): undated contacts, on the indexed guest only.** Measured that
night on `OutlookAI-Indexed` (section 8 item 24, Q99's `corpus-probe --undated-index-wait`): an
appointment and a task saved into a PST are indexed WITH `System.Message.DateReceived` - their creation
time - and a contact WITHOUT one (`<null>`), though the store gives all three a delivery time Outlook
will not remove. So a contact is the one undated row `LiveOrderKeyCollationTests` can have, and those
tests are `Requires=SearchIndex`: they run only on the indexed guest. Appointments and tasks stay out
there too - indexed as dated at build time, they would become the hub's newest rows, under the
frontier the staleness test reads. The plan option is `CorpusPlanOptions.IncludeUndatedContacts`,
`--undated-contacts` on the command line:

* **the same undated row counts version 2 sized, every one a contact:** the hub twelve (57-68 - every
  one a search hit, and the hub's top-100 search must stay under 100), the bystander forty-two (301-342 -
  more than the 35 rows of over-fetch room a scoped `TOP 25` has, so the widened-search guarantee can be
  told apart from luck). Fourteen, the bystander's contacts in the full set, would not be. The dated
  items are untouched - each population with them is the default plus its contacts, T1-pinned
  (digest `26D9CD68...`, and the dated half equal ordinal for ordinal);
* **IN the shape key**, as `|u:contacts` at its end - unlike `IncludeUndatedItems`: the contacts take
  the ordinals where the full set puts appointments, so neither population may be continued as the
  other. Every verb on such a manifest needs the option - teardown, census, the index wait - and is
  refused as another population without it;
* **undated in the INDEX, not the store** (`CorpusUndatedCriterion.IndexHoldsNoDate`): the write path
  does not try to remove the delivery time (a PST refuses it, measured), the undated probe requires the
  folder, the target store, the tag and the class but not the store's date, the read-back counts the
  store's delivery time instead of failing on it, and `corpus-indexed` - the index wait - says how
  many undated rows the index holds with no received date;
* **`Reset-HubPopulation.ps1` keeps what it tears down**: it reads `|u:contacts` off the manifest's
  header and passes `--undated-contacts` to every verb; a first build gives an indexed hub its contacts
  and an unindexed one none; and on a guest whose hub is indexed it refuses to `-Execute` ELEVATED -
  every Outlook start there is NOT elevated, the rebuild's own included (`-RunLevel Limited`).

On the UNINDEXED guest nothing changes: its populations are built without them, and the three tests are
deselected there (`Requires!=SearchIndex`). Where they run without undated rows - a hub built without
the option - they still print `verdict=no-undated-rows-in-sample`, `coverage: 0 ... on this machine`
and a `PROVED NOTHING:` line, which now says to rebuild with `--undated-contacts`.

**And all three kinds again - decided by the maintainer 2026-10-03 (his answer (b) to D62); how, decided on
his behalf (D126-D129 of `Docs/overnight-review-2026-10-03.md`).** The indexed guest's hub and bystander
carry version 2's full set once more, at its own ordinals and counts - the hub four appointments, four
contacts and four tasks (57-68), the bystander fourteen of each (301-342) - as
`CorpusPlanOptions.IncludeAllKinds`, `--all-kinds` on the command line, each kind given the date the index
can be held to:

* **a contact stays undated in the index**, exactly as under `--undated-contacts`: the store keeps the
  delivery time a PST gives it, nothing tries to remove it, and the index gives it no
  `System.Message.DateReceived`. The contacts are the undated rows `LiveOrderKeyCollationTests` measure;
* **an appointment and a task are DATED BY THE PLAN** - the deterministic part. Their delivery time is
  written after the first save, through the PropertyAccessor in UTC, exactly as a mail item is dated, to
  `CorpusPopulation.PlannedDeliveryUtc`: one day older than the oldest dated item the population can hold,
  one hour further back per ordinal - the hub's from 61 days back, the bystander's from 731. So under the
  index's `DateReceived DESC` they sort after all of the population's mail and before its undated
  contacts: never the frontier, never a "most recent" hit, never in a date window short of all the mail.
  **Measured before a single one was built** (section 4.2e, phase P1): the undated probe wrote
  `2026-08-03T09:32:50Z` on a throwaway appointment and task in the hub, read it back from the store, and
  the index dated both at exactly that instant 9 s after their save (`DATED AS WRITTEN`) - while their
  `System.Message.DateSent` and `System.DateModified` stayed at the save itself. So the index takes an
  appointment's and a task's `System.Message.DateReceived` from `PR_MESSAGE_DELIVERY_TIME`, which a plan
  can choose; the creation time Q98 (f) saw was only what a PST stamps into that property at the first save;
* **the build holds every one to it**: the undated probe writes the youngest planned instant on its own
  appointment and task and refuses the build unless both read it back; the build records each item
  before it judges the read-back (`RequirePlannedDeliveryTime`); the population read-back counts
  `UndatedDatedAsPlanned` and fails on a mismatch; and `corpus-indexed` is NOT complete until the index
  dates every planned appointment and task at its planned instant - so the per-run hub rebuild's index
  wait refuses a hub whose order is not the plan's;
* **IN the shape key**, `|u:all-kinds`, exclusive with both other undated options: an appointment of this
  population carries a date the full set's does not. `Testbed/guest/Reset-HubPopulation.ps1` tears a hub
  down by the marker its manifest carries and BUILDS it the decided way - all three kinds where the hub is
  indexed, none where it is not (D129) - so the next per-run rebuild moves a contacts-only hub over, and
  says so;
* **the hub's search budget is unchanged**: twelve non-mail rows, as with the twelve contacts, so the
  top-100 hub search (`Phase7LiveMcpToolShapeTests`) sees the same number of hits from them.

On the UNINDEXED guest nothing changes here either. Where the widened-search test can now contest a store
is no longer a margin either - see "the sized contest" in section 4.2e (D74).

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
8383, and `hub-unindexed` / `bystander-unindexed` / `identity-unindexed` with the same seeds on the
other guest - for the same reason the corpus ids are (section 2.9): the manifest is
`corpus-<id>.jsonl`, every guest's manifest lands in one pull directory, and a manifest is the only
allowlist teardown will delete from. **The anchor is not fixed**, and for the hub that is the point;
see below.

### Building them the first time

Every store that gets a population must be mounted in BOTH profiles - built in the account-less one,
read in the tier one - and attached to its second profile **while it is still empty** (section 2.6,
the addition written out there). So the order is:

1. **In the tier profile**, after sections 2.5 and 2.8: the hub exists, minted by
   `New-TierProfile.ps1` as the profile's DEFAULT store (`C:\OutlookAI-Tier\Outlook.pst` on both
   guests - measured) and named by `Rename-OutlookStore.ps1`; it is the only population store with
   every default folder. Create the bystander here too, new and empty, with
   `Add-OutlookPstStore.ps1`. Once section 2.8b has attached the identity account's store, it is here
   as well.
2. **Switch the default to the corpus profile** (`Set-DefaultOutlookProfile.ps1`) and restart
   Outlook - gracefully, under mailbox-safety rule 7.
3. **Attach each of those stores to the corpus profile by its path, before anything is in it**:
   `Add-OutlookPstStore.ps1 -ProfileName <corpus profile> -DisplayName tier@vm.invalid -Path
   C:\OutlookAI-Tier\Outlook.pst -Execute`, and the same for the bystander and the identity store. The
   name must come back byte-identical to the one the tier profile shows; the script checks.
4. **Build each population there**, as for a corpus: `corpus-probe`, then `corpus-build` dry, then
   `corpus-build --execute`, which runs its own census. For the hub:

   ```
   dotnet run --project <as above> -- corpus-build --population hub \
     --store tier@vm.invalid --allow-store tier@vm.invalid \
     --corpus-id hub-indexed --seed 8181 --anchor <now, UTC, to the second> \
     --manifest C:\OutlookAI-Q5\corpus-hub-indexed.jsonl --execute
   ```

   Before letting it proceed, read its own output, which since v2 says more than it used to:
   * the store and profile lines accepted, `profile accounts: 0`;
   * `Cross-store residue sweep (before the probes)` - and again after them - naming any store other
     than the target in which it found this corpus's probe items. It should find none;
   * `== placement probe ==`: `target store: NOT the profile's default store` for every population
     store (in the corpus profile the default is the corpus PST), `target folder:` the store's own
     Inbox for the hub or a STAND-IN for a store without one, then one line per rung - on a
     non-default store only `PostAsNote`, the one rung measured to keep its first save there - with
     `visible=True` and `store=target`. A rung whose first save landed in another store says
     `store=OTHER` and is unusable, whatever else it achieved;
   * the date probe verified;
   * **`== enrichment probe ==` reporting sender, recipients, attachment and conversation index all
     written** - and the recipient it checks is now the store's OWNER, the one v1 left unresolved;
   * the undated probe: with the undated items switched off (Q98 (a)) it probes nothing and prints
     `Undated probe: this population carries no undated item; nothing to probe.` On the indexed guest,
     with `--undated-contacts` (Q98 (f)), one line for the contact, with `folder=True inFolder=True
     tag=True class=True inTargetStore=True` - and `undated=False tableUndated=False`, the store's own
     date, reported and not judged - then `Undated probe verified for contact: ...`. With version 2's
     full set switched on, an `== undated probe ==` block where each kind must show `inFolder=True`,
     `undated=True`, `tableUndated=True`, `inTargetStore=True`, and no `removalRefused=` - which on a
     PST no kind does (measured 2026-09-27, below).

   A population is built only where one throwaway item proved every write it depends on, with no
   override. After the build, the census reads every item back - sender, recipients (the owner
   included), attachments, conversation, and each undated item's class and missing delivery time -
   and fails the build on any difference.
5. **Switch the default back to the tier profile**, restart Outlook, and on the indexed guest let
   the indexer settle: `corpus-indexed` (the command `Testbed/guest/Reset-HubPopulation.ps1` runs)
   waits until the index holds every item of a population. `outlook_health`'s `index.perStore[]`
   must then list every store in `indexedStoreDisplayNames` with rows.

The hub's build prints the `probeTerm` and `subjectOnlyProbe` values for the settings file; three of
the four `subjectOnlyProbe` fields are generator constants that `Testbed/testbed.json` already
carries, and the fourth is the hub's own name.

### The hub is rebuilt before every run

**On a frozen guest it is not - Q130 (a), decided 2026-10-03 (section 4.5).** There the frozen
checkpoint holds the hub as it was built and every run restores it at the same instant, so the hub's
newest item is exactly as old on every run as on the first - minutes, well inside the frontier margin.
The rebuild below is how a NEW frozen checkpoint's hub is made. The frontier test still reads the
manifest and still fails on a hub older than the margin, which on a frozen guest means its clock moved.

`LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier` asserts that the index frontier is not
in the future. On a guest whose newest item is weeks old, a product that misread local time as UTC
would still pass it; with the newest item one minute old, the same misreading puts the frontier an
hour or two in the future - the guest's UTC offset - and fails it. So the hub population is built
against **the moment the run starts**, which is why its anchor is not recorded in
`Testbed/testbed.json`: the manifest header records it, and the next rebuild reads it from there.

**The rebuild is a script step, not a paragraph - decided 2026-09-24 (question D, option (a)), the
pattern chosen for `ImportPRF`.** `Testbed/guest/Reset-HubPopulation.ps1 -Execute`, through
`Register-InteractiveTask.ps1 -TimeoutSeconds 3600`, right after a guest restart through
Testbed/host/Restart-Guest.ps1 (section 4 has the line). It reads everything from the guest's own live-test settings - the hub, the manifest
(`hubPopulationManifestPath`, whose file name is the population id) and whether the hub is indexed -
and the seed and the old anchor from the manifest's header, and then:

1. refuses a Production profile, a manifest of any other population or store, a running Outlook (it
   quits none it did not start) and a pending `ImportPRF` (it starts Outlook twice);
2. makes the account-less profile the default, starts Outlook and waits 180 s;
3. `corpus-teardown --execute` by the anchor the manifest records - the anchor is part of the shape
   key, so a teardown given today's would be refused as another population - and requires 0 left;
4. moves the torn-down manifest to `hub-history\<id>.<anchor>.jsonl` beside it, a name
   `Copy-FromGuest.ps1` never mistakes for a live manifest;
5. `corpus-build --execute` against now, to the second: every probe, the build, its census and its
   read-back, exit 0 required - then reads the new header back and requires that anchor;
6. quits the Outlook it started, under mailbox-safety rule 7 - attached through the Running Object
   Table, every Outbox proven empty by its store's folder mask, no item window open, and OUTLOOK.EXE
   the one process it started; a quit that does not complete (a modal dialog swallows `Quit()`,
   measured) stops with "`Testbed/host/Restart-Guest.ps1 -Execute`, then `-SkipRebuild`";
7. makes the tier profile the default again, starts Outlook on it NOT ELEVATED
   (`Start-OutlookUnelevated.ps1` - an elevated Outlook never feeds the index, section 8 item 22,
   and the run attaches at the user's own level), and on the indexed guest runs `corpus-indexed`
   until the index holds every item of the new population;
8. prints the newest item's instant, the minutes the frontier test has left, and this guest's
   opt-in value and `dotnet test` filter for the run (`Testbed/README.md` section 4c).

Its own Outlook, the one the build runs against in step 2, is at the task's level: the corpus tool
runs in the same task and attaches only to an Outlook at its own integrity level. **On the indexed
guest that level is NOT elevated** - run it through `Register-InteractiveTask.ps1 -RunLevel Limited`,
and it refuses to `-Execute` elevated there - because an elevated Outlook never feeds the index, and
every Outlook start on that guest is unelevated; it then starts both Outlooks directly, at its own
level. On the unindexed guest either level works. **And it builds the non-mail kinds the decision names
for the guest** (D129, 2026-10-03): the TEARDOWN takes whatever the manifest's shape key says the old hub
carried (`|u:all-kinds` → `--all-kinds`, `|u:contacts` → `--undated-contacts`), and the BUILD gives a hub
that is indexed here all three kinds (`--all-kinds`, D62 (b), above) and one that is not none - so a
contacts-only hub is moved over by its next rebuild, which says so in a note. Until 2026-10-03 it kept
what it tore down.

Teardown removes the population's items AND the folders it created, and drains Deleted Items behind
itself: a delete in a PST is a soft one that re-issues the EntryID, so teardown re-scans and deletes
again, still by both keys. **The frontier test has to run within the guest's UTC offset of the build,
less its own five-minute tolerance** - under 55 minutes in winter and 115 in summer on a `W. Europe`
guest, counting the indexer's crawl and everything the run does before it gets there. It no longer
depends on anyone remembering: the frontier test reads the manifest and FAILS, naming the script, on a
hub too old to catch a local-time frontier with.

`-SelfTest` covers every decision above (79 assertions, Windows PowerShell 5.1 and 7), and the guest
guard refuses the script on the workstation. **It has run on a guest since 2026-10-03** - on
`OutlookAI-Unindexed`, twice, every step above but the index wait, which that guest skips (section
4.1d): each teardown `0 corpus item(s) remaining`, each build 56 of 56 censused and read back, each
quit 2 s, Outlook back on the tier profile NOT elevated, and a margin of 110 minutes printed with the
opt-in and the filter. No live run has read the result yet. What the teardown leaves in Deleted Items
is section 3b's item 7.

The bystander and identity populations have no such clock: no test reads their dates, so they are
built once and left alone. A checkpoint restored from before they were built needs them built again.

### What v2 changed, after v1's first guest build

`OutlookAI-Unindexed`, 2026-09-24 (section 4.1, step 6): every v1 build exited 1. v2 answers each
fault, host-side, and each answer is a probe or a check that fails the build rather than a hope:

| v1 fault, measured | v2 |
| --- | --- |
| Every RECEIVED item carried one unresolved To row with an EMPTY address: the owner was added as `tier@vm.invalid <tier@vm.invalid>`, which Outlook's resolver refuses. The enrichment probe passed, because its item was addressed to correspondents only | A recipient whose name IS an address is added as the bare address (`CorpusCorrespondent.ToRecipientSpec`), and the enrichment probe's To row is now the store's OWNER - the case that failed |
| A PST attached with `AddStoreEx` has no Inbox and no Sent Items; `GetDefaultFolder(6)` returned its hidden root, the bystander's 172 and the identity store's 5 "Inbox" items went there, and the placement probe printed `target= landedIn=` and said VERIFIED | Every default folder is found without a creating lookup (`SpecialFolders.Resolve` and the Calendar/Contacts/Tasks designations), and returned only when it is a NAMED folder under the store's root (`CorpusFolderVisibility`). A store without one gets a visible stand-in under its root, recorded in the manifest. The placement probe refuses an invisible target |
| Failed probe rungs stranded 12 items in Corpus B's Drafts - the corpus profile's DEFAULT store, on no allowlist - because an unsent mail item's first save goes to the default store's Drafts, and the delete looked in the target | A store that is not the profile's default is probed only with `PostAsNote` (a post, created in the sent state, converted to a note) - the one rung measured to keep its first save there. v2 shipped with `InPlaceReceived` (the flags written before the first save) beside it; on `OutlookAI-Unindexed`, 2026-09-27, its first save landed in the corpus profile's default store (`store=OTHER`), so it was retired from every ladder the same day. Every first save's store is read before anything else, and an item that landed elsewhere is deleted THERE; every other store's Drafts and Deleted Items are swept of this corpus's probe items before the probes, after them, and after teardown, by the two-key rule - `ComCorpusMailbox.SweepProbeResidueOutsideTarget` |
| `GetDefaultFolder` created Drafts and Junk Email in the bystander during its build | No lookup creates a folder: the scans, the probes and the build all go through the non-creating resolver |

**The Drafts-then-Move rungs are not probed on a non-default store either, and that is measured, not
assumed.** The third stranded item of each session was the date probe's `ObjectModel` item, created
in the TARGET store's own Drafts for `DraftsThenMoveWithSentFlag`: it failed before its move and was
found in Corpus B's Drafts. So even a Drafts-created item's first save lands in the default store, and
every item those rungs built in v1 passed through Corpus B on its way.

**If no rung stays in the target store, the build refuses - there is no override.** The directions
were (a) build each population with its store as the profile's DEFAULT store - a throwaway
account-less profile per store, created with `/PIM`, which cannot be deleted without the GUI; (b)
accept a transient write into the default store for the Drafts-then-Move rungs, with the transit check
and the cross-store sweep behind it - which the maintainer's rule "a probe must never write outside its
target store" forbids as written; (c) Extended MAPI's `IMessage` creation in the target folder, which
the Dependencies rule allows but the repository chose not to rebuild on (Q67, knowledge only).
**Decided 2026-09-27 (Q88 (a)) - and not needed:** `PostAsNote` keeps its first save in the target, on
both stores probed (section 4.1a):

```
hub (its own Inbox):
  PostAsNote  target=Inbox visible=True store=target landedIn=Inbox parentMatches=True inFolderTable=True sentFlag=True usable=True
bystander (a stand-in):
  PostAsNote  target=OutlookAI-Corpus-Folder-6 visible=True store=target landedIn=OutlookAI-Corpus-Folder-6 parentMatches=True inFolderTable=True sentFlag=True usable=True
```

(a) stands as the route to take if a later probe ever refuses `PostAsNote` too; the refusal message
names it.

### The bystander keeps its shape - checked test by test

The bystander's role is the count tripwire's watched store AND the absent-arrival-folders shape
(section 1.3), so before deciding to keep it without an Inbox, every test that reads it was checked:

* **`LiveFolderScopeTests.PrimaryStore_ExcludeSubfolders_NarrowsExactly_AndCostsNothing`** takes the
  first non-hub entry of the indexed list - the bystander, by design - and walks its folder tree for
  ANY mail folder of 5 to 20,000 items with populated children. It never asks for an Inbox. v1's
  "Inbox" and its two subfolders were under the hidden root, so it could not have found them; v2's
  stand-in `OutlookAI-Corpus-Folder-6` holds 172 items with `-Projects` (40) and `-Suppliers` (32)
  under it - exactly the shape it needs. **Kept on the bystander; the test is unchanged.** Moving it to
  the hub would mean choosing a different store than the one the test is written to measure.
* **The count tripwire** censuses every mail folder of the store, stand-ins included; it needs items
  inside the identity budget, not an Inbox.
* **`LiveOrderKeyCollationTests`** read the bystander's undated rows through a store-scoped index
  query; no folder is named. (Since Q98 (a) there are none to read on a guest - above.)

No test needed the bystander's Inbox, so none was moved and none was weakened.

### The undated kinds are dated in a PST - DECIDED 2026-10-03 (Q98): (a) now, then (f)

**Decided by the maintainer 2026-10-03: (a) now, then (f).** The populations are built WITHOUT the
undated items now - `CorpusPlanOptions.IncludeUndatedItems`, off (above). (f), added to the directions
below by the decision: another agent measures on the indexed guest whether the index dates those kinds;
they come back only if that says it is worth it. The question, the guest's lines and the directions
stay below as the record.

**(f) measured the same night, and decided on the maintainer's behalf: undated CONTACTS, on the
indexed guest only** (D62; "The undated items", above, has the option and what it changes). The index
dates an appointment and a task at their creation time and leaves a contact with no
`System.Message.DateReceived`, so the indexed guest's hub carries twelve contacts and its bystander
forty-two; `OutlookAI-Indexed` was built that way the same morning, and `corpus-indexed` found every
one of the 54 in the index with no received date (section 4.2c).

**The question.** The hub's 12 and the bystander's 42 undated items exist so the three
`LiveOrderKeyCollationTests` have index rows with NO `System.Message.DateReceived` to measure (question
A (a), 2026-09-24; drafts left out, Q89 (a)). Their premise was that an appointment, a contact or a task
carries no received date. **In a PST it does, and Outlook will not take it off**, so the undated probe
refuses every kind and no population that carries undated items - the hub, the bystander - can be
built, on either guest. The identity population carries none and is not affected.

**What the guest showed** (`OutlookAI-Unindexed`, 2026-09-27, section 4.1a; the bystander, in its
stand-ins `-9`, `-10` and `-13`, printed the same three lines):

```
== undated probe ==
  appointment  folder=True inFolder=True tag=True undated=False tableUndated=False class=True inTargetStore=True removalRefused=UnauthorizedAccessException: The property "http://schemas.microsoft.com/mapi/proptag/0x0E060040" does not support this operation.
  contact      folder=True inFolder=True tag=True undated=False tableUndated=False class=True inTargetStore=True removalRefused=UnauthorizedAccessException: The property "http://schemas.microsoft.com/mapi/proptag/0x0E060040" does not support this operation.
  task         folder=True inFolder=True tag=True undated=False tableUndated=False class=True inTargetStore=True removalRefused=UnauthorizedAccessException: The property "http://schemas.microsoft.com/mapi/proptag/0x0E060040" does not support this operation.
```

* **The date is STORED, not supplied by the object model.** `undated=False` is the PropertyAccessor's
  read; `tableUndated=False` is the folder's own table, restricted by the store - it returns the item for
  `NOT ("urn:schemas:httpmail:datereceived" IS NULL)` and not for `IS NULL`. So every DASL path of the
  product sees these items as dated, and so would anything reading the message through MAPI.
* **Outlook refuses to delete it**, for all three kinds: `PropertyAccessor.DeleteProperty` on
  PR_MESSAGE_DELIVERY_TIME throws E_ACCESSDENIED.
* **What the INDEX makes of it is not known.** Whether Windows Search derives `System.Message.DateReceived`
  from that property for a calendar item, a contact or a task is item 4 below - the indexed guest only.
  `LiveOrderKeyCollationTests`' own comment says those kinds carry none; nothing here has measured it.
* Everything else held: each kind landed in its own folder of the target store, kept its tag and its
  class, and the probes left no item behind - only the bystander's four stand-ins, empty (section 4.1a).

| Direction | How | For | Against |
| --- | --- | --- | --- |
| **(a) Defer the undated items** | Build v2 without them - the hub 56 items, the bystander 300 - and keep the probe, the census check and their T1 tests, so a later route re-enables them by giving the two populations their undated kinds back | Unblocks every v2 build on both guests today; builds nothing on an unmeasured premise; the order-key tests are exactly where they were before v2 (`no-undated-rows-in-sample`), no worse | Gives up, for now, the measurement question A added them for; the population counts, digests, `UndatedRemedy`'s wording and these docs change |
| (b) Build them dated, and let the index answer | For the three kinds the probe's criterion becomes folder, tag, class and store - not "no delivery time" - and `corpus-indexed` reports per kind whether the index dates them | Keeps the items; the one reading the tests depend on, the index's, is taken on the indexed guest as a by-product; if the index leaves those kinds undated, the tests measure at last | Builds on an unverified premise - if the index dates them, 54 items give the tests nothing; the hub's newest indexed row becomes an item saved at build time, which the frontier test's reasoning never planned for; census, `corpus-indexed` and the docs change |
| (c) Remove the date through Extended MAPI | `IMAPIProp::DeleteProps` and `SaveChanges` on the saved item | Permitted by the Dependencies rule; the probe's two reads would say at once whether it held | `item.MAPIOBJECT` hands out an in-process MAPI interface and the corpus tool runs out of process, so the dependable form is a MAPI session of the tool's own (`MAPIInitialize`, `MAPILogonEx`, `OpenMsgStore`, `OpenEntry`) - the interop this repository deleted and kept as knowledge only (Q67); and whether the PST provider stamps the date again on `SaveChanges` is unknown |
| (d) Undated drafts, where the store IS the default | The Q88 (a) route - a throwaway `/PIM` profile per store - so an unsent mail's first save stays in the store; drafts become the undated rows (reopens Q89 (a)) | Outlook's own object model only; an unsent mail is the textbook row with no received date | The heaviest: a profile per store, none deletable without the GUI; the bystander would gain the full default folder set it is designed not to have; that a draft in a PST carries no delivery time is itself unmeasured |
| (e) Measure the premise first, read-only | On the maintainer's workstation: content-free COUNTS of the index's `mapi` rows by kind, with and without `System.Message.DateReceived` - no Outlook, no MAPI, no mailbox, like the three SELECTs of 2026-08-18. Exchange is believed to give every item a delivery time, calendar items and contacts included (not measured here), so undated calendar and contact rows there would mean the index does not date those kinds at all | About ten minutes; decides between (a) and (b) on evidence | A read of the workstation's index about real mail, counts only - the maintainer's to allow; and it speaks for Exchange items, not PST ones |

**Recommended: (a) now, and (e) as the follow-up that decides whether (b) is worth building.** The NULL
collation is a cost question, not a correctness one: `IndexOrderGuard` is sound under any collation by
construction (QUESTIONS.md, 2026-08-18), so nothing the product guarantees waits on it - while every v2
population and the per-run hub rebuild do. ~~Nothing is changed until this is decided; the probes refuse,
as designed, and `CP-14A-POPULATIONS-V2-PROBED` holds the probed guest.~~ Decided: (a), then (f) - the
top of this subsection.

### The identity store has no Inbox - DECIDED (a) 2026-09-27 (Q87), MEASURED AND BUILT

**Measured and built the same day on `OutlookAI-Unindexed`** (section 4.1b; checkpoint
`CP-10B-IDENTITY-REAL-INBOX`): the minted store kept its designated Inbox as a secondary store of
the tier profile - folder mask `0xFF` in both profiles - so `-Phase CaptureStore` PASSED from the
tier profile, and the identity account now delivers into `Inbox` (PST node id `0x8082`). The route
is scripted: section 2.8b, first box. What (a) was up against, and the four other directions, stay
below as the record of the decision.

**The question.** The identity account delivers into the identity PST (section 2.8b), and a delivery
store needs a real, designated Inbox - one the store's own folder mask and receive folder name, that
Outlook's tree shows, that `SpecialFolders.Resolve` finds. `identity.pst` is attached by `AddStoreEx`
and has none, so both guests bound the account to its hidden root (section 4.1, defect 4), and v1
built the identity population's received items there too. `Add-IdentityAccount.ps1 -Phase CaptureStore`
now refuses such a store, so the identity account - and the identity population, whose received items
belong in that Inbox - wait on this decision.

| Direction | How | For | Against |
| --- | --- | --- | --- |
| **(a) Mint it as a default store, then attach it** | `OUTLOOK.EXE /PIM <scratch profile>` mints a PST under `ForcePSTPath` as that profile's DEFAULT store, with every default folder - the mechanism that gave the hub its full set, measured on both guests. Rename it (`Rename-OutlookStore.ps1`), then attach it to the tier profile (`Add-OutlookPstStore.ps1`) and run CaptureStore / Bind / Verify, which now pass on a designated Inbox | Outlook's own folders, designations and receive folder; only existing, measured scripts; nothing undocumented beyond the Bind that already runs | One scratch profile left on the guest (a profile has no free delete route, section 1); the minted file name must be read off the guest; whether the designations survive the attach as a secondary store is the measurement |
| (b) Outlook's own "Change Folder" | Account Settings > Change Folder > New Folder, the only route Microsoft documents | Documented | GUI only: UIAutomation that `Dump-UiaTree.ps1` has not even classified; and whether it DESIGNATES the folder, or only binds it, is unknown |
| (c) Extended MAPI | Create the folder, `IMsgStore::SetReceiveFolder`, the IPM designations and the folder mask through `mapi32.dll` from `Add-Type` | Permitted by the Dependencies rule; complete control | The interop this repository deleted (the `IProfAdmin` IIDs were wrong, Q67 knowledge only); writes store internals nothing else writes; the most code to get wrong |
| (d) A plain folder named Inbox | `Folders.Add("Inbox", 6)` under the IPM root, and bind the account to it | Cheap, scriptable, visible | NOT designated: the store still has no Inbox to the mask, `SpecialFolders.Resolve` or the product's sweep; it fixes the symptom the tests never read and leaves the defect |
| (e) Let POP3 delivery make one | Bind the account to the store with an empty `Delivery Folder EntryID` and see what the first send/receive does | Nothing to write if it works | Unknown; the plausible outcome is delivery falling back to the default store's Inbox - the tier's - which is exactly what the identity account must not share |

**Decided 2026-09-27: (a) (Q87)**, measured before anything is written into `Add-IdentityAccount.ps1`. The
measurement, from a checkpoint before the identity account (`CP-09-ADDIN-READY` on
`OutlookAI-Unindexed`), every step in session 1 and every Outlook close a graceful quit: first
re-apply the index exclusion, which CP-09 predates - `Set-OutlookIndexingDisabled.ps1 -Execute`, then
`-Verify` must say `UNINDEXED` with reason `USER`; list `C:\OutlookAI-Tier\*.pst`;
`OUTLOOK.EXE /PIM IdentityMint`, wait 90 s, list again - the new file is the minted store; read over
COM that store's `PR_VALID_FOLDER_MASK` (Inbox bit set?) and, only if it is, its Inbox's name and parent
chain; quit; confirm the default profile is still `OutlookAI-Tier`; then section 2.8b's sequence with
`-PstPath` and `Add-OutlookPstStore.ps1 -Path` naming the minted file (renamed to
`identity@vm.invalid` first), where `-Phase CaptureStore` must now PASS - and read the mask and the
Inbox again from the tier profile, which is the question (a) turns on. **Record the store's display
name at each stage - after the mint, after the rename, after the attach** (Q92, still with the
maintainer). **Run 2026-09-27, and it passed** - section 4.1b has the lines and the names.

### What only a guest can answer

**Answered by v1's build on `OutlookAI-Unindexed`, 2026-09-24:** the store guard accepts the hub, the
bystander and the identity store in the account-less profile; sender, correspondent recipients,
attachment and conversation index all land; the store computes one conversation id per conversation;
and the four faults above. **Answered by v2's probes on the same guest, 2026-09-27 (section 4.1a):**
items 1, 2, 3 and 8. **Answered by v2's build and the first two hub rebuilds there, 2026-10-03
(section 4.1d):** item 7, and the half of item 8 the probes left. **Answered on the indexed guest the
same morning (section 4.2c):** item 6, and item 4's undated rows; **and item 5 by the 160,000-item
corpus's two-profile mount (section 4.2d).** **Still open:** the rest of item 4, which the first live
run there answers.

1. **ANSWERED: `PostAsNote` does; `InPlaceReceived` does NOT.** In a store that is not the profile's
   default, `PostAsNote` kept its first save in the target, parented it in the target folder and showed
   it in that folder's table - the hub's own Inbox and the bystander's stand-in alike (the lines are
   under "What v2 changed"). `InPlaceReceived`'s first save landed in another store -
   `InPlaceReceived  target=Inbox visible=True store=OTHER landedIn=(unknown) parentMatches=False
   inFolderTable=False sentFlag=False usable=False error=its first save landed in another store` - and
   the sweep after the probes found and deleted it there (`1 probe item(s) of 'hub-unindexed' found in
   'Outlook Data File', which is NOT the target; 1 deleted by the two-key rule`). Retired from every
   ladder the same day.
2. **ANSWERED: yes.** The enrichment probe, whose To row is the owner, on both stores:
   `sender=True recipients=True attachment=True conversationIndex=True conversationId=(computed)`.
3. **ANSWERED: they stay - and they are dated.** On the hub (its own Calendar, Contacts and Tasks) and
   the bystander (stand-ins), every kind `folder=True inFolder=True tag=True class=True
   inTargetStore=True`; but `undated=False tableUndated=False`, and Outlook refused to remove the date.
   Decided 2026-10-03, above: built without them (Q98 (a)).
4. **Does the index carry what the tests read** - `FromAddress`/`FromName`, `ToAddress` (the owner's
   now), one attachment row per attachment, a `ConversationID` shared by each conversation's members,
   and the undated rows with NO `System.Message.DateReceived`? The indexed guest only. **The last part
   ANSWERED 2026-10-03** (Q98 (f); section 8 item 24, and section 4.2c): a contact is indexed with no
   received date, an appointment and a task with one - and every one of the indexed guest's 54 built
   contacts reads back that way through `corpus-indexed`. The rest waits for the first live run there.
5. **Does a store mounted in two profiles give the index two scopes?** The per-store scope URL is
   `mapi16://{SID}/StoreDisplayName($Hash)/`. Corpus A has always been in the same position, so the
   answer - whatever it is - is not new to the populations. **Half answered 2026-09-27 (Q92, section
   8 item 24):** `$Hash` is computed from the store's entry ID, which for a PST holds its file path
   and nothing else that varies, and the name is the store's own, not the profile's - so one file
   mounted from the same path in two profiles should give ONE scope. Predicted from the measured
   rule; the two-profile mount itself was not made. **ANSWERED 2026-10-03: ONE scope** (section
   4.2d). On `OutlookAI-Indexed` every population store and the new 160,000-item Corpus A are mounted
   in both profiles, each built - and pushed into the index - by the corpus profile's Outlook. With the
   TIER profile's Outlook then running for over thirteen minutes, mounting all of them, the index
   still held exactly one scope per store - `Corpus A($996dc7a9)` 160,006 rows, `bystander@vm.invalid
   ($38ebfb53)` 374 - with nothing queued and the catalog idle: no second scope, no second crawl.
6. **ANSWERED 2026-10-03: no time at all, when the rebuild's Outlook is NOT elevated.** It bounds how
   soon after the rebuild the run can start, and so how much of the UTC-offset margin is left for the run
   itself. On `OutlookAI-Indexed` (section 4.2c) the rebuild ran at `-RunLevel Limited`, so its own Outlook
   pushed every item to the index as it was created: `corpus-indexed` asked once, `[0 s]`, and found all 68
   - and the same held for the first builds of all three populations (hub 68, bystander 342, identity 8,
   each complete at its first ask). The whole rebuild took about six and a half minutes - six of them its
   two 180-second Outlook warm-ups - and left 110 minutes of margin.
7. **ANSWERED: yes - and they pile up, two per rebuild.** Teardown removes a created folder with
   `Folder.Delete()`, and in a PST that MOVES the emptied folder into the store's Deleted Items. Measured
   on `OutlookAI-Unindexed`, 2026-10-03, with `corpus-folders` of the hub after each of the first two
   `Reset-HubPopulation.ps1` runs (section 4.1d). After the first:

   ```
   Deleted Items  items=0  mail
     OutlookAI-Corpus-Folder-Projects  items=0  mail  [made by the corpus tool]
     OutlookAI-Corpus-Folder-Notices  items=0  mail  [made by the corpus tool]
   ```

   After the second, the same two and `OutlookAI-Corpus-Folder-Projects (2)` and
   `OutlookAI-Corpus-Folder-Notices (2)`, also empty: Outlook renames a moved folder whose name is taken,
   so nothing collides and nothing stops them accumulating. Both teardowns printed `folders removed 2`,
   which counts `Delete()` calls that did not throw - not folders gone from the store. The items are
   drained (`considered 112, deleted 112`: each of the 56 deleted, then deleted again from Deleted
   Items, and the hub's Deleted Items reads `items=0`); the folders are not. **Harmless to the tests**:
   the live tier's own test folders are `OutlookAI-McpTest-Folder*`, a different string that neither
   contains nor is contained in `OutlookAI-Corpus-Folder` (`CountLiveTestFolders` would not count these),
   the subject-only probe names its folder by path (`Inbox/OutlookAI-Corpus-Folder-Notices`), and the
   count tripwire exempts the hub. What it costs: the hub's tree grows by two folders per run, and the
   teardown's count says something that did not happen. **Decided overnight on the maintainer's behalf,
   2026-10-03, for his review: (a), and not built yet** - nothing about it blocks a live run, and building
   it starts with a measurement (below). The directions:

   | Direction | How | For | Against |
   | --- | --- | --- | --- |
   | **(a) Drain the folders as the items are drained** | After `Folder.Delete()`, look the folder up again by its manifest EntryID; if it now sits directly in the store's Deleted Items and its name still ordinal-contains `OutlookAI-Corpus-Folder`, `Delete()` it again - from Deleted Items that is permanent. Count "removed" only when the second lookup fails, and print "moved to Deleted Items" otherwise | Both keys, as for every deletion here; mirrors the item drain the teardown already does; the count becomes true | Rests on a PST folder KEEPING its EntryID across the move - unmeasured (an item's does not); if it does not, the folder must be found by name under Deleted Items, which is one key, not two |
   | (b) Keep the folders, empty them only | The hub teardown deletes the items and leaves its two subfolders in place; the build's code already adopts an existing child of that name, as item 8 measured for a stand-in | Nothing moves, nothing piles up, no second delete | A teardown for good leaves two empty folders under the Inbox; "torn down" stops meaning "gone", which the bystander and identity teardowns would have to say differently |
   | (c) Leave it, and say so | Change only the printed line - `folders moved to Deleted Items 2` - and let `Reset-HubPopulation.ps1` report how many corpus folders the hub's Deleted Items holds | No deletion logic changes | The pile grows on any guest not restored from a checkpoint between runs |
   | (d) A separate sweep of old corpus folders | `Reset-HubPopulation.ps1` hands the torn-down manifests in `hub-history\` to a verb that deletes those folders from Deleted Items, by their recorded EntryIDs and the prefix | Clears the pile left by every past run, not just the current one | A second deletion path to keep correct; same EntryID question as (a) |

   **Why (a)**: it is the only direction that both stops the pile and keeps two keys on every deletion.
   Its EntryID question comes first, read-only and through the corpus tool rather than ad-hoc COM:
   `corpus-folders` prints no EntryIDs, so it needs an option that resolves a manifest's recorded folder
   EntryIDs and prints where each one sits now. The four folders in the hub's Deleted Items came from
   the two torn-down manifests in `hub-history\`, and `CP-12B-POPULATIONS-V2` holds both - the
   measurement is ready to run. Nothing was changed here.
8. **ANSWERED: a probe that makes a stand-in and a build that then refuses leave that stand-in, empty,
   in the target store.** The bystander, before its probe: `Deleted Items  items=0` and nothing else.
   After it, and after the undated probe refused: `OutlookAI-Corpus-Folder-6  items=0  mail`, `-9
   calendar`, `-10 contacts`, `-13 tasks`, every one `items=0` and `[made by the corpus tool]`, under
   the store root - and no Inbox, Drafts or Junk Email created. **And the next build adopts a stand-in by
   name - measured 2026-10-03 (section 4.1d):** on a new, empty bystander the probe made
   `OutlookAI-Corpus-Folder-6`, and the build's own placement probe and its 172 items used it - one `-6`
   in the tree afterwards, beside the `-5` the build made for the sent items.

---

## 4. Running the tier

**The run lines live in ONE place: `Testbed/README.md` section 4c.** This section keeps no copy, so
the two cannot drift. There the filter below runs on a guest through
`Testbed/guest/Register-InteractiveTask.ps1 -RunLevel Limited` - NOT elevated, because an elevated
Outlook never feeds the index (section 8 item 22) - with the per-run opt-in `OUTLOOKAI_LIVE_OPT_IN`
set to that guest's computer name inside the task's own script, and `-c Release`. The filter:

`Category=Live&Requires!=DelegateStore&Requires!=CachedExchange`

That filter IS the VM bucket, spelled out: everything live except the tests naming a capability
this machine cannot be given. There is no separate "which bucket" trait to keep in step with it -
see section 5. It is not typed anywhere either: `McpServer/OutlookAI.McpServer.Tests/T2/LiveRunFilters.cs`
derives it from the vocabulary, and `T1/LiveTierInventoryTests` fails the build if a copy of it - this
one included - stops short of the derived string. `CachedExchange` joined `DelegateStore` on
2026-10-03 (Q74 C1).

**On a FROZEN guest the run starts with the restore of its frozen checkpoint instead - decided by the
maintainer 2026-10-03 (Q130 (a); section 4.5).** Both Outlook guests are frozen once that work is
merged: time synchronisation off, every run restoring the checkpoint `Testbed/testbed.json` names
under `frozenClocks`, staging, passing `Testbed/host/Set-GuestClockFrozen.ps1 -Verify` and starting
Outlook NOT elevated on the tier profile before the suite - with no restart and no hub rebuild after the
restore (`Testbed/README.md` section 4c has the
order). The hub rebuild below is then the first step of making a NEW frozen checkpoint, not of a run;
until the merge it stays the first step of every run on the unfrozen checkpoints.

**On a test guest, every run starts with the hub rebuild - a script step, decided 2026-09-24
(question D, option (a)).** Restart the guest gracefully - Testbed/host/Restart-Guest.ps1 -VMName <guest>
-Execute, the only restart this runbook allows - then in session 1:

```
.\Register-InteractiveTask.ps1 -TimeoutSeconds 3600 -Script "& 'C:\OutlookAI-Q5\Reset-HubPopulation.ps1' -Execute"
```

`Testbed/guest/Reset-HubPopulation.ps1` tears the hub population down by the anchor its manifest
records, builds it again against now, switches back to the tier profile, and on the indexed guest
waits until the index holds all of it (section 3b has the steps and every refusal). It ends by
printing how many minutes `LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier` has left -
the guest's UTC offset less five minutes from the newest item, 55 in a W. Europe winter - and this
guest's filter. **Start the run inside that margin.** The frontier test reads the manifest named by
`hubPopulationManifestPath` and FAILS on a hub nobody rebuilt, naming the script: skipping the step
is a red run, not a quietly weaker one. The maintainer's own machine has no such field and no such
step.

*Until 2026-10-03 a second step followed here - `Reset-ThrowawayStore.ps1`, which
recreated a throwaway data file for the created-folder proof (`T2/LiveCreatedFolderTests`). Both
were removed that day with the requirement they proved: the maintainer dropped Q85's "must report"
(section 8 item 25). A run record that names step 9a-ii predates that.*

**On `OutlookAI-Unindexed` the filter also deselects the index tier:**
`Category=Live&Requires!=DelegateStore&Requires!=CachedExchange&Requires!=SearchIndex`. That guest has no index by design
(section 1.1a), so every test carrying `Requires=SearchIndex` would refuse there, not measure. The
rebuild script prints the right filter for the guest it runs on, from whether the hub is in
`indexedStoreDisplayNames`.

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

**The maintainer's workstation runs no live test at all** since 2026-10-03 (Q116 (a)) - its profile
refuses them in `LiveTestSettings.Load` (`Testbed/README.md` section 4d). **The Exchange test VM runs
a different filter, and only that one:** read-only, Exchange-only - `Testbed/README.md` section 4e,
and section 5 below for the `Writes=Nothing` trait it rests on.

To run one class - on a test guest, inside the same interactive-task script as `Testbed/README.md`
section 4c's lines, opt-in included:

```
dotnet test <csproj> --filter "Category=Live&FullyQualifiedName~LiveTableSortProbeTests"
```

**Where each kind of run happens, since 2026-10-03 (Q94, `AGENTS.md`):** a live run on a test guest,
as above; the Exchange-only read-only subset on the Exchange test VM (section 4.4; until Q116 (a) it
ran on the maintainer's workstation, which runs no live test now); and the NON-live suite on none of
them - it runs on the build VM, through
`Testbed/host/Invoke-TestsOnBuildVm.ps1` (section 4.3).

**A filtered run is fully guarded.** It takes the census, runs the health preflight, checks
corpus freshness and sink reachability, and verifies at the end of whichever collection the
filter left last. That was not true before 2026-08-19: verification lived in one collection's
teardown, so any run that did not include that collection paid for a baseline and threw it
away.

To see the sets without running anything - `--list-tests` discovers and does not execute, so it
is safe against any mailbox:

```
dotnet test <csproj> --list-tests --filter "Category=Live"                                                       # 130
dotnet test <csproj> --list-tests --filter "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange"      # 123
dotnet test <csproj> --list-tests --filter "Category=Live&(Requires=DelegateStore|Requires=CachedExchange)"      # 7
```

Treat those numbers as "what they were when this was written" - measured 2026-10-03, after Q74,
again the same day after Q96 (iv) added `LiveCreatedFolderTests` (128 and 121 before it), and once
more after the Q99 folder-name proof added `LiveFolderNameEncodingTests` (129 and 122 before it). The
traits are the authority; the counts in a document drift. `Requires!=X` means "no value of `Requires` on
this test equals X", which is what makes a multi-valued trait usable as an exclusion.

### 4.1 The unindexed guest's build-out - `OutlookAI-Unindexed`, 2026-09-24 (the live run itself: on hold)

**Why this section exists.** Until this date the live tier had never run on a test machine (section 9,
`Testbed/README.md` section 6 item 13). This is the record of bringing `OutlookAI-Unindexed` from
`CP-05-CORPUS-B-CLEAN-UNINDEXED` to the full design, one scripted step at a time, each step proven
on the guest and checkpointed, and then running the tier there.
*(Retitled 2026-09-27: it was called "The first live-tier run on a VM", and that run never
happened - step 9 below was put on hold, and the tier has still not run on any guest (section 8
item 22). What the section records is the build-out.)* Raw logs of every step were kept
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
| 6. Fixture populations (section 3b) | In the tier profile `Add-OutlookPstStore.ps1` made `bystander@vm.invalid` (`C:\OutlookAI-Tier\bystander.pst`); all three population stores read 0 items. Default to `CorpusProfile`, the hub (`Outlook.pst`), bystander and identity PSTs attached there by path, names byte-identical. `corpus-probe --population hub`, then per population a dry run and `--execute` - `hub-unindexed` 8181, `bystander-unindexed` 8282, `identity-unindexed` 8383 (the last id is not yet in `testbed.json`; *2026-09-27: it is - added the same evening, `5d22c7b`, under `corpusIdConvention.populations`*), one anchor `2026-09-24T16:08:00Z` - then `corpus-census` of all three and of Corpus B. Default back to `OutlookAI-Tier` | **NOT CLEAN - every population build exited 1.** Placement census clean: 56, 300 and 8 items, each ordinal once, in the folder the manifest records. Read-back FAULTS: 42 of 56, 244 of 300 and 5 of 8 items - exactly the RECEIVED ones - carry no owner recipient. And three more defects, below. Corpus B: its own census clean (20,000), but 12 probe items now sit in its Drafts | `CP-11-POPULATIONS-BUILT-WITH-FAULTS` - evidence, not a base to build on |
| 7. SDK and suite | `Publish-LiveTierPayload.ps1` on the host (source archived from `329925d`; the 54-package feed reused), both archives expanded on the guest, `Install-DotnetSdk.ps1 -ExpectedSha512 <the published hash> -Execute` over PowerShell Direct; then `-Verify` from a new session | `TEST-READY` both times: the installer exited 0 after 74 s, the offline probe ran (`OUTLOOKAI-SDK-PROBE-OK 10.0.12 x64`), **2,794 tests discovered**, 17 executed and passed; 148 s end to end. The Release build baked `McpServerExePath` = `C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer\bin\Release\net10.0-windows\OutlookAI.McpServer.exe`, and that file exists - question 12 answered by construction, as the installer's banner predicted | `CP-12-SDK-TEST-READY` |
| 8. Settings | Store names read over COM in the tier profile (three stores, the two accounts on their own stores); `OutlookAI-Unindexed`'s section of `Testbed/testbed.json` filled - watched `tier@`, `bystander@`, `identity@vm.invalid`, indexed `[]`, bystander `bystander@vm.invalid`, `corpus` null (the tier profile does not mount Corpus B), `mailSink` loopback 25/110 with a 2,000 ms connect timeout; `New-LiveTestSettings.ps1 -VMName OutlookAI-Unindexed`; the file copied to the suite's `live-fixtures` | rendered and admitted (Portable; hub `tier@vm.invalid`; identity `identity@vm.invalid` draft-and-delete only); SHA-256 `401A4BA8...5C6BF1` on host and guest alike | `CP-13-LIVE-SETTINGS` |
| 9. First live run | **ON HOLD, not started** - by instruction: the live tier's own census, `LiveOutlookTestMailer.CaptureMailFolderCensus`, calls `GetDefaultFolder` for Deleted Items and ids 19-23 on every store, and on a PST lacking one of those folders that CREATES it (step 6 above measured the same thing from the generator's side). A run now would write into the bystander PST and measure the harness. It waits for the non-creating resolver and the matching harness fix. *(2026-09-27: that fix landed the same evening - Q84, `4f82004`: the census and the sweeps now resolve those folders without creating them. The run itself has still not happened; section 8 item 22.)* | - |

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

Also measured here, and the reason for the harness fix the first live run now waits for (landed
the same day - Q84, `4f82004`):
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

### 4.1a Populations v2 on `OutlookAI-Unindexed`, 2026-09-27 - probed, NOT built

**Why this section exists.** The first guest run of generator v2 (section 3b), from
`CP-14-EXCLUDED-BY-SCOPE-RULE` - `CP-10-IDENTITY-ACCOUNT` plus the index exclusion - on the one guest
this work owned. Every mailbox write went through the corpus tool; nothing else wrote to a store.
Raw logs: `.work\g2-populations-v2\` in the main checkout.

| Step | What ran | Verdict | Checkpoint |
| --- | --- | --- | --- |
| 1. Restore and stage | `Restore-VMSnapshot CP-14-EXCLUDED-BY-SCOPE-RULE`; the tools and server published from `685a2be`, the settings rendered by `New-LiveTestSettings.ps1`, and `Reset-HubPopulation.ps1`, `Start-OutlookUnelevated.ps1`, `Register-InteractiveTask.ps1`, `Add-OutlookPstStore.ps1`, `Set-DefaultOutlookProfile.ps1` copied in | restored `Running`; the default profile `OutlookAI-Tier`, no Outlook, no pending `ImportPRF` | - |
| 2. Still unindexed | `Set-OutlookIndexingDisabled.ps1 -Verify` | `the service says: included=False reason=USER`; two catalog readings ten minutes apart, `mapiRows=0` both; `VERDICT: UNINDEXED` | - |
| 3. Stores | Outlook started on the tier profile; `Add-OutlookPstStore.ps1` made `bystander@vm.invalid` there (`before : Store.DisplayName='Outlook Data File'`, `after : Store.DisplayName='bystander@vm.invalid'`); `Restart-Guest.ps1 -Execute`; the default to `CorpusProfile`, Outlook started, the hub and the bystander attached by path while empty | quit in 0 s, restart 25 s; both names came back byte-identical; `corpus-folders` on each: the hub every default folder, all `items=0`; the bystander `Deleted Items` only | - |
| 4. Probes | `corpus-probe --population hub` and `--population bystander`, four tool builds (below) | placement, dates and enrichment VERIFIED on both; **undated REFUSED on both** | - |
| 5. After | `corpus-folders` of the hub, the bystander and Corpus B; `Restart-Guest.ps1 -Execute`; the state read back | hub all `items=0`; the bystander four empty stand-ins; Corpus B unchanged - `Inbox 10912`, `Sent Items 4964`, `Deleted Items 2461`, `Junk Email 1663`, `Drafts 0`, as before the probes; Outlook not running, default `CorpusProfile` | `CP-14A-POPULATIONS-V2-PROBED` |

**Four tool builds, because the guest proved three fixes necessary - each made host-side first, with
T1 tests and controls:**

1. `685a2be`: `InPlaceReceived`'s first save landed in the corpus profile's default store
   (`store=OTHER`, and the sweep after the probes deleted it there); `PostAsNote` stayed. And every
   undated kind read back CARRYING a delivery time. **`185113a`** retired `InPlaceReceived` from every
   ladder, leaving `PostAsNote` the one rung for a non-default store, and made the undated write path
   remove the delivery time after the first save.
2. `185113a`: the undated probe ended the run - `FATAL: InvalidOperationException: corpus undated probe
   failed.` - with its cause dropped, and the sweep after the probes skipped. **`0ab9232`** names the
   whole inner chain on the FATAL line and runs that sweep in a `finally`, as the build and re-anchor
   already did.
3. `0ab9232`: `... <- caused by: UnauthorizedAccessException: The property
   "http://schemas.microsoft.com/mapi/proptag/0x0E060040" does not support this operation.` - the
   PropertyAccessor refuses the removal, and E_ACCESSDENIED is not in the COM-failure set the probe
   caught. **`c3529c9`** returns a refused removal instead of throwing it, reports it per kind, reads
   the date a second way - the folder's own table - and has the build record an undated item before it
   refuses one.
4. `c3529c9`: every probe ran to the end and refused cleanly; the lines are in section 3b, "The undated
   kinds are dated in a PST", which is where this run stops.

**Stopped there, as designed:** no population was built, so `CP-15-POPULATIONS-V2` was not taken, the
hub rebuild (`Reset-HubPopulation.ps1`) did not run, and the live tier did not run. Apart from the
first run's one `InPlaceReceived` item, deleted where it landed, no sweep found a probe item in any
other store - before or after the probes; the `185113a` run skipped its sweep after them (the defect
`0ab9232` fixed), and the next run's sweep before them found nothing.

### 4.1b The identity store minted - `OutlookAI-Unindexed`, 2026-09-27 (Q87 (a))

**Why this section exists.** Direction (a) of section 3b, "The identity store has no Inbox": measured
first with the repository's existing scripts, then scripted into `Add-IdentityAccount.ps1` and run
again as its phases. Every step in session 1 at `RunLevel Highest`, every Outlook close
`Testbed/host/Restart-Guest.ps1 -Execute` (with `-CancelLogonPrompt` whenever the identity account's
POP3 prompt was up), every read through a repository tool. Raw logs: `.work\g2-identity-q87\` in the
main checkout.

| Step | What ran | Verdict | Checkpoint |
| --- | --- | --- | --- |
| Restore | `Restore-VMSnapshot CP-09-ADDIN-READY` | default `OutlookAI-Tier`, no Outlook, `ImportPRF` = `C:\OutlookAI-Tier\identity-account.prf` (the identity import pending since before CP-09), policy `PreventIndexingOutlook` 1, no scope rule | - |
| Restage | master `751b5dc`'s guest scripts, and the tools published from it | every staged hash equal to the host's | - |
| Exclusion | `Set-OutlookIndexingDisabled.ps1 -Execute` (session 0, Outlook closed) | `before: ... included=False reason=UNKNOWNSCOPE`, the rule added, `no Outlook row in the catalog: nothing to purge`, then its own verify: `included=False reason=USER`, `mapiRows=0` twice, `VERDICT: UNINDEXED` | `CP-09B-EXCLUDED-RESTAGED` |
| Measure (by hand) | one tier start completing the import; `OUTLOOK.EXE /PIM IdentityMint`; `Add-OutlookPstStore.ps1 -ListOnly`, CaptureStore as a dry run and `corpus-folders` in the mint profile; `Rename-OutlookStore.ps1`; the attach to the tier profile; CaptureStore, Bind, `-StoreSinkPassword`, Verify | CaptureStore **PASSED from the tier profile**; Verify `OK` | not kept |
| Script | `-Phase Mint` and `-Phase CaptureMint`, the path taken from the mint record (`3474396`) | `-SelfTest` 61/0 under 5.1 and 7 | - |
| Prove | from `CP-09B` again, as the phases: Mint (refused while the import was pending, as written), the tier start, Mint, `/PIM`, CaptureMint, rename, attach, CaptureStore, Bind, `-StoreSinkPassword`, Verify; then `Set-OutlookIndexingDisabled.ps1 -Verify` | every phase as below; `VERDICT: UNINDEXED`, reason `USER`, after the four Outlook starts of this pass | `CP-10B-IDENTITY-REAL-INBOX` |

**What the guest printed, pass 2 (pass 1 printed the same lines):**

```
-Phase Mint, the import pending:  REFUSING to prepare the mint: an ImportPRF is pending ('C:\OutlookAI-Tier\identity-account.prf') ...
-Phase Mint -Execute:             mint profile 'IdentityMint' (does not exist yet); ForcePSTPath 'C:\OutlookAI-Tier' holds 2 PST(s): Outlook Data File - CorpusProfile.pst, Outlook.pst
OUTLOOK.EXE /PIM IdentityMint:    'Outlook Today - Outlook', no dialog; default profile still 'OutlookAI-Tier'
-Phase CaptureMint -Execute:      minted store 'Outlook Data File' at C:\OutlookAI-Tier\Outlook Data File - IdentityMint.pst (profile 'IdentityMint'): mask 0xFF; Inbox designated=True, name 'Inbox', visible=True; Drafts designated; 14 folder(s) counted, 0 item(s)
Rename-OutlookStore.ps1:          current: Store.DisplayName='Outlook Data File' ... after: Store.DisplayName='identity@vm.invalid'
Add-OutlookPstStore.ps1 (tier):   before : Store.DisplayName='identity@vm.invalid' ... after : Store.DisplayName='identity@vm.invalid'
-Phase CaptureStore -Execute:     identity store: C:\OutlookAI-Tier\Outlook Data File - IdentityMint.pst (from the mint record)
                                  store 'identity@vm.invalid' at ...: StoreID 164 bytes; mask 0xFF; Inbox designated=True, EntryID 24 bytes, name 'Inbox', visible=True; Drafts designated
-Phase Bind -Execute:             before: Delivery Store EntryID = <the tier store's> ... after: bound to 'identity@vm.invalid' - both values read back byte-identical
-Phase Verify -TrySmtpAddress:    COM: the identity account delivers into 'Inbox' (visible=True, PST node id 0x8082)
                                  COM: 2 account(s), 2 distinct delivery store(s)
                                  COM: account 'OutlookAI identity sink' SmtpAddress='identity@vm.invalid'
                                  OK: 'OutlookAI identity sink' delivers to its own store 'identity@vm.invalid' ...
the sink's log:                   5:14AM DBG read USER identity / read PASS any-value
```

In the mint profile, before the rename, `corpus-folders` (pass 1) listed Inbox, Sent Items, Deleted
Items, Outbox, Drafts, Calendar, Contacts and Tasks - every one `items=0`, and Junk Email `ABSENT`.

**The display names, for Q92.** Over COM (`Store.DisplayName`): after the mint `Outlook Data File`,
after the rename `identity@vm.invalid`, after the attach `identity@vm.invalid`. The name the index
files a store under is the store's OWN `PR_DISPLAY_NAME` - its root folder's name, which is what
`Rename-OutlookStore.ps1` renames (section 8 item 24, which corrects item 22) - and that rename ran in
the mint profile, before the attach: so this store's own name is `identity@vm.invalid` from then on,
and the index should file it as `identity@vm.invalid($<hash>)`, where both guests' AddStoreEx stores
were `Outlook Data File(...)`. Unmeasured: this guest has no index. Also read, in the registry after
pass 2: the name each PROFILE keeps for the file in its service section (`PR_DISPLAY_NAME_W`) is
`identity@vm.invalid` in the tier profile and in the mint profile, and `Outlook Data File` for Corpus B
in `CorpusProfile` - which, by item 24, the index does not use.

**What is left.** The mint profile, which nothing opens again. The identity account's signature:
`CP-10B` has none, where `CP-10` had one on the wrong subkey - put back in section 4.1c. And
`OutlookAI-Indexed`, whose identity account is still the old one.

### 4.1c Signature, SDK and suite on the minted line - `OutlookAI-Unindexed`, 2026-09-27

**Why this section exists.** What the old `CP-11`/`CP-12` line had, put on the `CP-10B` line: the
identity account's signature, the .NET SDK and the suite. From `CP-10B-IDENTITY-REAL-INBOX`, master
`af56efc` staged first. No population (Q98 is open) and no live tier. Raw logs:
`.work\g2-cp11b\` in the main checkout.

| Step | What ran | Verdict |
| --- | --- | --- |
| Restage | `Publish-GuestPayload.ps1` at `af56efc`, the server and the tools expanded (the old ones kept as `server.fe65ced`, `tools.751b5dc`); master's guest scripts copied | server and tools `99.99.99.0+af56efc...`; every staged script hash equal to the host's. `Register-InteractiveTask.ps1` and `Reset-HubPopulation.ps1` were already current on this line - `CP-09B` was restaged from `751b5dc` - and are unchanged since |
| 1. Signature | the account entries read from the registry; Outlook started NOT elevated (`Start-OutlookUnelevated.ps1 -Profile OutlookAI-Tier`, up in 6 s); `Set-AccountSignature.ps1 -Account identity@vm.invalid -Execute` in a `RunLevel Limited` task (Medium Mandatory Level); the entries read again; `list_accounts` and `list_signatures` asked separately, read-only; `Restart-Guest.ps1 -Execute` (its quit task picked Limited from Outlook's token) | before: every entry `New Signature=''`. `signature 'Identity' -> create`, `account: identity@vm.invalid / new message: Identity`, `Verified`. After: `New Signature='Identity'` on `00000004` (clsid `{ED475411-...}`, `OutlookAI identity sink`, `identity@vm.invalid`) and on nothing else - `00000005`, the data-file entry named `identity@vm.invalid`, still `''`; `Identity.htm` 152 B, `.rtf` 113 B, `.txt` 72 B. The separate read-back: `account 'identity@vm.invalid' newMessage='Identity'`, `account 'tier@vm.invalid' newMessage=''`. Only those four tools were called - none creates an item - and every Outbox read 0 at the quit |
| 2. SDK | `Install-DotnetSdk.ps1 -ExpectedSha512 <MEDIA.md's> -Execute`, then `-Verify` from a new session | `hash matches`, `installer exited 0` in 40 s, `wrote C:\OutlookAI-Q5\src\NuGet.config`; both runs `VERDICT: TEST-READY` - SDK 10.0.401, the probe `OUTLOOKAI-SDK-PROBE-OK 10.0.12 x64`, **3,105 tests discovered**, 17 run and passed |
| 3. Suite | `Publish-LiveTierPayload.ps1 -Ref af56efc` on the host (54 packages, 75.7 MB; the offline feed restores all five projects with every other source cleared), `Source.zip` and `NuGet.zip` expanded into `src` and `nuget-offline` - before step 2, which builds from them | 451 source files, 54 packages; the guest-built `McpServerExePath` target exists |
| 4. Index | `Set-OutlookIndexingDisabled.ps1 -Verify`, Outlook closed - after step 1's NOT elevated Outlook, the kind that registers itself where nothing excludes it | `the service says: included=False reason=USER`; `mapiRows=0 ... outlookRowsTotal=0` at both readings, ten minutes apart; `VERDICT: UNINDEXED`. The catalog itself grew from 1,397 to 3,365 items while it watched - files, none of them Outlook's |
| 5. Checkpoint | Outlook not running, default profile `OutlookAI-Tier` | `CP-11B-SIGNATURE-SDK-SUITE` |

**Order.** Step 3's archives went in before step 2 ran: `Install-DotnetSdk.ps1` writes the suite's
`NuGet.config` into `src` and builds the suite from it, so `TEST-READY` needs both there first.

**Not on this line yet, and not asked for:** the live-test settings file (`Testbed/README.md` step 9 -
the old `CP-13` had one), any population (Q98), and a live run. *(2026-10-03: the settings and the
populations are on it now - section 4.1d. The live run is not.)*

### 4.1d The populations built - `OutlookAI-Unindexed`, 2026-10-03 (Q98 (a))

**Why this section exists.** Q98 decided (a): the populations are built without their undated items
(section 3b). From `CP-11B-SIGNATURE-SDK-SUITE`, on this guest only: the three populations, the hub
rebuild, the settings, a checkpoint - and no live run, which another agent runs. The tools were
published from this work's branch - master `ea40cc8` plus the Q98 switch (`ace34e0`) and the Corpus B
record (`c256817`), because master alone did not have the switch yet. Every mailbox write went through
the corpus tool; every Outlook close was a graceful quit (`Restart-Guest.ps1 -Execute`, or the hub
rebuild's own); every step that starts or drives Outlook ran in session 1. Raw logs:
`.work\g2-cp12b\` in the main checkout.

| Step | What ran | Verdict |
| --- | --- | --- |
| Restore and restage | `Restore-VMSnapshot CP-11B-SIGNATURE-SDK-SUITE` (the VM was Off, so the restore left it Saved and `Start-VM` resumed it); `Publish-GuestPayload.ps1` at `c256817` on the host and only the tools expanded on the guest - the old ones kept as `tools.af56efc`, the server left at `af56efc`'s; the branch's 17 guest files copied (scripts, `.prf` templates, `SearchCrawlScope.cs`) | tools `99.99.99.0+c2568171867b04d1c0f2dfd7e0357ca48e8d237e`; the eight scripts this run calls hash the same on the guest as on the host (`Reset-HubPopulation.ps1` `52445BBBFC8255AB`, `Register-InteractiveTask.ps1` `82B07836611D97E3`, ...) |
| 1. The bystander | Outlook on the tier profile; `Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName bystander@vm.invalid -Path C:\OutlookAI-Tier\bystander.pst -Execute`, then `-ListOnly` | `before : Store.DisplayName='Outlook Data File'`, `after : Store.DisplayName='bystander@vm.invalid'`; the tier profile holds 3 stores - `tier@vm.invalid` (`Outlook.pst`), `identity@vm.invalid` (`Outlook Data File - IdentityMint.pst`), `bystander@vm.invalid` (`bystander.pst`) - and 2 accounts |
| 2. The corpus profile | `Restart-Guest.ps1 -Execute`; `Set-DefaultOutlookProfile.ps1 -Name CorpusProfile -Execute`; Outlook started (elevated, as the corpus tool is), 180 s; the three stores attached by path while empty; `corpus-folders` of all four | every name byte-identical; hub every default folder, all `items=0`; the bystander `Deleted Items` only; the identity store every default folder but Junk Email, all `items=0`; Corpus B `Inbox 10912`, `Sent Items 4964`, `Deleted Items 2461`, `Junk Email 1663`, `Drafts 0` |
| 3. Hub | `corpus-probe`, `corpus-build` dry, `--execute` - `hub-unindexed` 8181, anchor `2026-10-03T01:16:19Z` | below |
| 4. Bystander | the same - `bystander-unindexed` 8282, anchor `2026-10-03T01:17:22Z` | below |
| 5. Identity | the same - `identity-unindexed` 8383, anchor `2026-10-03T01:18:26Z`, into the store's real Inbox | below |
| 6. After the builds | `corpus-folders` of all four | hub `Inbox 24` (`-Projects 6`, `-Notices 12`), `Sent Items 14`, `Deleted Items 0`; bystander `OutlookAI-Corpus-Folder-6 172` (`-Projects 40`, `-Suppliers 32`), `-5 56`, `Deleted Items 0`, still no Inbox, Drafts or Junk Email; identity `Inbox 5`, `Sent Items 3`; Corpus B unchanged, `Drafts 0` |
| 7. Settings | `New-LiveTestSettings.ps1 -VMName OutlookAI-Unindexed` on the host; `Copy-ToGuest.ps1` with the line it printed | rendered and admitted - Portable, hub `tier@vm.invalid`, 3 watched, bystander `bystander@vm.invalid`, no index, no corpus, the sink on loopback; SHA-256 `7AC0A324AAAD8117501AA908997B66A0BA9D2B77B3ABCC75D82AD84A8A32186D` on host and guest alike |
| 8. Hub rebuild | `Restart-Guest.ps1 -Execute`; `Reset-HubPopulation.ps1 -SelfTest` (Windows PowerShell 5.1.26100 - the guest has no PowerShell 7), the dry run, then `-Execute` through `Register-InteractiveTask.ps1 -TimeoutSeconds 3600` | `79 assertion(s), 0 failure(s)`; the dry run's plan sheet `items : 56` and the command lines; `-Execute` below |
| 9. Item 7 | `Restart-Guest.ps1 -Execute`, the corpus profile, `corpus-folders` of all four; the restart, `Reset-HubPopulation.ps1 -Execute` a SECOND time, and the same again | section 3b, item 7: two emptied folders in the hub's Deleted Items after the first rebuild, four after the second; bystander, identity store and Corpus B unchanged both times |
| 10. State | `Restart-Guest.ps1 -Execute`; `Set-DefaultOutlookProfile.ps1 -Name OutlookAI-Tier -Execute`; the state read; `Set-OutlookIndexingDisabled.ps1 -Verify`, Outlook closed - after the two rebuilds' NOT elevated Outlooks, the kind that registers itself where nothing excludes it; then this section's banner-only edit of `Reset-HubPopulation.ps1` copied in | `outlook processes: 0`, `ImportPRF: ''`, `DefaultProfile: 'OutlookAI-Tier'`; `the service says: included=False reason=USER`, `mapiRows=0 ... outlookRowsTotal=0` at both readings, ten minutes apart, `VERDICT: UNINDEXED`; the copied script's SHA-256 the same on guest and host (`2754193F...`) |

**What each build printed** - the probes and the build of every population in one run, the same lines
for all three but the target folder and the counts:

```
Cross-store residue sweep (before the probes): no probe item of 'hub-unindexed' in any other store.
  target folder: 'Inbox' - the store's own Inbox
  PostAsNote                   target=Inbox visible=True store=target landedIn=Inbox parentMatches=True inFolderTable=True sentFlag=True usable=True
Date fidelity: VERIFIED via PropertyAccessorDates. Received dates drive DASL selection.
  sender=True recipients=True attachment=True conversationIndex=True conversationId=(computed)
Undated probe: this population carries no undated item; nothing to probe.
Cross-store residue sweep (after the probes): no probe item of 'hub-unindexed' in any other store.
  undated items         : none - switched off since 2026-10-03 (Q98 (a)): in a PST these kinds are dated, and Outlook will not remove it
Build finished: created 56, already present 0, failed 0, 40,286 body bytes in 00:00:02 (20.9 items/s).
Census: 56 item(s) found for 56 planned; per folder found/planned: Sent Items=14/14, Inbox=24/24, Inbox/OutlookAI-Corpus-Folder-Projects=6/6, Inbox/OutlookAI-Corpus-Folder-Notices=12/12. Every ordinal exists exactly once, in the folder the plan names.
Population read-back: 56 of 56 item(s) read; 4 of 4 conversation(s) grouped by the store under one id. Every item carries the sender, recipients, attachments and conversation the plan names.

bystander:
  target folder: 'OutlookAI-Corpus-Folder-6' - a STAND-IN: the store has no visible Inbox of its own, and keeps none
Build finished: created 300, already present 0, failed 0, 260,041 body bytes in 00:00:13 (22.2 items/s).
Census: 300 item(s) found for 300 planned; per folder found/planned: Sent Items=56/56, Inbox=172/172, Inbox/OutlookAI-Corpus-Folder-Projects=40/40, Inbox/OutlookAI-Corpus-Folder-Suppliers=32/32. Every ordinal exists exactly once, in the folder the plan names.
Population read-back: 300 of 300 item(s) read; 6 of 6 conversation(s) grouped by the store under one id. ...

identity:
  target folder: 'Inbox' - the store's own Inbox
Build finished: created 8, already present 0, failed 0, 5,501 body bytes in 00:00:00 (21.1 items/s).
Census: 8 item(s) found for 8 planned; per folder found/planned: Sent Items=3/3, Inbox=5/5. Every ordinal exists exactly once, in the folder the plan names.
Population read-back: 8 of 8 item(s) read; 0 of 0 conversation(s) grouped by the store under one id. ...
```

Every probe and build exited 0, and no sweep - before the probes, after them, or after a teardown -
found a probe item in any other store. The bystander's census labels its folders by the plan's names
(`Inbox=172/172`); the items are in the stand-ins `corpus-folders` lists. The manifests:
`corpus-hub-unindexed.jsonl` 59 lines (header, 56 items, 2 folders), `corpus-bystander-unindexed.jsonl`
305 (header, 300 items, 4 folders - `-6`, `-5` and the two subfolders), `corpus-identity-unindexed.jsonl`
9 (header and 8 items: the store has every folder its population needs).

**The hub rebuild's first run** (the second printed the same lines, with anchor `2026-10-03T01:43:20Z`):

```
Manifest records 56 item(s) and 2 created folder(s).
Teardown: considered 112, deleted 112, refused by rule 0, already gone 0, failed 0, folders removed 2.
Post-teardown scan finds 0 corpus item(s) remaining (expected 0).
Cross-store residue sweep (teardown): no probe item of 'hub-unindexed' in any other store.
  torn-down manifest kept as C:\OutlookAI-Q5\hub-history\hub-unindexed.20261003T011619Z.jsonl
=== corpus-build ... --anchor 2026-10-03T01:29:52Z ... --execute
Build finished: created 56, already present 0, failed 0, 40,286 body bytes in 00:00:02 (21.4 items/s).
Census: 56 item(s) found for 56 planned; ... Every ordinal exists exactly once, in the folder the plan names.
Population read-back: 56 of 56 item(s) read; 4 of 4 conversation(s) grouped by the store under one id. ...
=== quitting the Outlook this script started (pid 1440), under mailbox-safety rule 7
  not in the Running Object Table - attached through the class factory
  profile 'CorpusProfile', item windows open: 0
  OUTLOOK.EXE left 2s after the Quit.
OK: OUTLOOK.EXE pid 544 in session 1, NOT elevated, profile 'OutlookAI-Tier'. Left running - close it with Testbed/host/Restart-Guest.ps1.
Hub population 'hub-unindexed' in 'tier@vm.invalid' is anchored 2026-10-03T01:29:52Z; its newest item is 2026-10-03T01:28:52Z.
The frontier test can catch a local-time misreading until 2026-10-03T03:23:52Z - 110 min from now (115 min on this guest's UTC offset of 02:00:00). START THE RUN NOW, by Testbed/README.md section 4c, with:
  the opt-in   $env:OUTLOOKAI_LIVE_OPT_IN = 'OAI-UNINDEXED'
  the filter   --filter "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange&Requires!=SearchIndex"
```

(Shown as the script prints it since Q74. The guest's copy that night predated Q74 and printed the
same filter without `Requires!=CachedExchange`, which Q74 added a few hours later.)

**Why it ran twice.** Section 3b's item 7 asked what the teardown leaves in the hub's Deleted Items; the
first run answered "two emptied folders", and only a second could say whether a second pair collides,
fails or piles up. It piles up. The checkpoint therefore holds the hub as the SECOND rebuild left it -
anchor `2026-10-03T01:43:20Z`, four empty corpus folders in its Deleted Items, both torn-down manifests
in `hub-history\`. A run from the checkpoint starts with a third rebuild in any case.

| Checkpoint | State |
| --- | --- |
| `CP-12B-POPULATIONS-V2` (parent `CP-11B-SIGNATURE-SDK-SUITE`; taken with the guest running, 2026-10-03 04:04 local) | Outlook not running; default profile `OutlookAI-Tier`; no `ImportPRF`; `UNINDEXED`, reason `USER`; the three populations built; the hub at anchor `2026-10-03T01:43:20Z` (manifest SHA-256 `86141A7E...`), the bystander at `01:17:22Z`, the identity store at `01:18:26Z`; Corpus B untouched (manifest SHA-256 `B7373BA0...`, 20,001 lines); the live-test settings staged; the tools from `c256817`, the server and the suite still `af56efc`'s from `CP-11B` |

**For the run, which is not part of this:** the suite on the guest is `af56efc`'s, staged at `CP-11B`.
It predates the Q98 switch, and master has changed several live test files since; re-stage it from
master before the run (`Testbed/README.md` step 8b). The one live check that reads the hub
population's plan - the frontier test's `LiveHubPopulationFreshness` - judges the newest DATED item,
which is the same with or without the undated items, so the old suite would not misread this hub.

### 4.1e The live tier's first runs - `OutlookAI-Unindexed`, 2026-10-03

**Why this section exists.** The first time the VM bucket ran end to end anywhere (section 9 said
nobody had). From `CP-12B-POPULATIONS-V2` (section 4.1d), on this guest only, by the route section 4
and `Testbed/README.md` section 4c give: `Restart-Guest.ps1 -Execute`, the hub rebuild
(`Reset-HubPopulation.ps1 -Execute`, default level), from run 2 the throwaway data file
(`Reset-ThrowawayStore.ps1 -Execute`, `-RunLevel Limited` - merged from master that morning), then
the suite through `Register-InteractiveTask.ps1 -RunLevel Limited` with `OUTLOOKAI_LIVE_OPT_IN =
'OAI-UNINDEXED'` set inside the task's script and
`dotnet test ... -c Release --filter "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange&Requires!=SearchIndex"`,
plus a TRX logger, the console logger at `detailed` and `--diag`; the results came back through
`Copy-FromGuest.ps1`. Every fix went in on the host first with T1 tests, proven on the build VM
(`Invoke-TestsOnBuildVm.ps1`: 3,478 of 3,478 and every self-test at `67c4afb`, `1059f4a` and `0f4bbfa`), and before runs 2 to 5
the guest was restored to `CP-12B` and restaged from the new commit (`Publish-GuestPayload.ps1`,
`Publish-LiveTierPayload.ps1`, `Install-DotnetSdk.ps1 -Execute`, `-Verify`: TEST-READY each time), the
settings re-rendered once for the throwaway key (SHA-256 `43A77749...`). Raw logs, TRX files and
console captures: `.work\g2-live-green\` in worktree `agent-a634a99d582147265`.

| Run | Revision | Total | Passed | Failed | Skipped | Suite time |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | `3e7b861` (master) | 80 | 66 | 14 | 0 | 17.6 min |
| 2 | `3a0fb05` (fixes F1-F7, F9; master `cd8a984` merged) | 80 | 77 | 3 | 0 | 7.1 min |
| 3 | `67c4afb` (F10; F9's change reverted) | 80 | 78 | 2 | 0 | 7.3 min |
| 4 | `1059f4a` (F8 second attempt) | 80 | 77 | 3 | 0 | 7.1 min |
| 5 | `0f4bbfa` (F8 third attempt) | 80 | 78 | 2 | 0 | 7.2 min |
| 6 | `f0cc4a2` (this branch + master `21bbdfd` merged) | 80 | 72 | 8 | 0 | 7.1 min |
| 7 | `f0cc4a2` again, no restore | 80 | 71 | 9 | 0 | 6.7 min |
| 8 | master `21bbdfd` alone, for isolation | 81 | 67 | 14 | 0 | 14.3 min |

**Runs 6 to 8: merging master `21bbdfd` made OUTLOOK.EXE crash.** With master's twenty newer commits
merged in (`f0cc4a2`, kept on branch `a634-merge-21bbdfd-outlook-crash`; the non-live suite passed it,
3,513 of 3,513), Outlook crashed in the Phase-4 collection in both runs - the guest's Application log:
`Faulting application name: OUTLOOK.EXE, version: 16.0.17932.20996 ... Faulting module name: ntdll.dll
... Exception code: 0xc0000005 Fault offset: 0x0000000000078cad`, and every Phase-4 test after it failed
with `RPC_S_SERVER_UNAVAILABLE`. Run 6 crashed just after `LiveDraftTests.IdentityDrafts` passed, run 7
inside its `new_draft` into `identity@vm.invalid`; that test took 7.4 s and 4.5 s there against 0.28 to
0.38 s in every other run. Neither half crashes alone: this branch before the merge ran four times
without it (runs 2 to 5), and master alone (run 8, from a separate worktree, restored `CP-12B`) ran
`IdentityDrafts` in 0.28 s with no crash and no `APPCRASH` - its 14 failures are run 1's kinds again,
F1 to F10, which confirms each of those diagnoses a second time. So the crash comes from the
two together, and the branch tip was put back to `c1b67d9`, the last state the guest ran clean; the
merge needs that interaction found before it lands (`TODO.md`).

**The crash, found (same day, later runs; every one from `CP-12B`, restaged).**

| Run | Revision | Total | Passed | Failed | Outlook |
| --- | --- | --- | --- | --- | --- |
| E1 | `f0cc4a2` + timing diagnostics, `LiveDraftTests` only | 5 | 5 | 0 | no crash |
| E2 | the same, whole filter | 80 | 69 | 11 | crashed (`IdentityDrafts`) |
| E3 | without the true-root read (`5ac1d85`) | 80 | 51 | 29 | crashed, `OLMAPI32.DLL` `0x2e411` (`NewDraft_WithAttachments`) |
| E4b | without the snapshot's `PR_CONVERSATION_INDEX_TRACKING` read | 80 | 76 | 4 | no crash |
| 9 | `5d94851` - master `60fba07` merged, tracking read removed (`3f8cfc3`) | 80 | 54 | 26 | crashed, `OLMAPI32.DLL` `0x2e411` |
| 10 | `c1b67d9` - the branch alone, a control | 80 | 78 | 2 | no crash |
| 11b | `e0cbe0a` - merged, after the host's restart | 80 | 72 | 8 | crashed, `OUTLOOK.EXE` `0x1571e5` |
| 13 | `ace09f9` - merged, COM child objects released (`9664aa0`) | 80 | 79 | 1 | no crash |

E4b's one clean run made the tracking read look like the trigger, and `3f8cfc3` took it out; run 9
crashed anyway, so it was not. The control (run 10, the branch alone, after the host had been
restarted) made it five clean runs of five without master and six crashes in seven with it - and
master's whole product delta since `cd8a984` (`59fad08`, Q96's follow-ups) adds no COM call on any
path that succeeds: it was compared line by line, and the hub rebuild's logs were identical. **What
differs between two builds whose COM calls are identical is when the .NET garbage collector runs**, and
three things in this process were left for it to release: the `Column` that `Table.Columns.Add`
returns (the scan's and the sweep's date columns, and the tripwire census's - which runs first, on 30
folders), the `Bookmark` that `Bookmarks.Add` returns in the compose paths (and a `bm.Parent` document
reference), and the `Attachment` that `Attachments.Add` returns in the test mailer - each released,
whenever a collection happened, after the table, the document or the mail it belongs to was gone,
inside Outlook. The faults moved with every run (`ntdll.dll`, `OLMAPI32.DLL`, `OUTLOOK.EXE` itself),
which is what a damaged heap looks like. `9664aa0` holds and releases every one of them and
`T1/ComChildObjectReleaseTests` pins it from the sources. **Run 13 was the first merged run without a
crash**, and the next two were clean too:

| Run | Revision | Total | Passed | Failed | Outlook | Suite time |
| --- | --- | --- | --- | --- | --- | --- |
| 13 | `ace09f9` | 80 | 79 | 1 (D49, before its decision) | no crash | 7.9 min |
| 18 | `fd2c58b` | 80 | 80 | 0 | no crash | 7.5 min |
| 19 | `c1a72f1` (`585a9f6` + this record) | 80 | 80 | 0 | no crash | 7.2 min |

Three clean merged runs of three after `9664aa0`, against six crashes in seven before it. The mechanism
is inferred from that and from the code - no crash dump was taken (no debugger on the guests, by the
dependency rule) - so it is the strongest available reading, not an observed one. In between, D49
was found and decided (F9 below) with Explorer-only probes and a ninety-second subset run (the
show-me tests plus `LiveDisconnectRecoveryTests`, no hub rebuild) that reproduced the full suite's
failure where five probe designs had not. **`CP-13B-LIVE-GREEN` was taken** right after run 18
(16:49 local, a standard checkpoint, parent `CP-12B-POPULATIONS-V2`): `fd2c58b` staged, the hub as
that run left it, Outlook not running. Run 19 ran from `CP-12B` as every run does.

**Every run, the guards.** The sink probe answered from run 2 on (`[sink] submission 127.0.0.1:25 and
retrieval 127.0.0.1:110 both answering` - run 1 never armed it, F1). The count tripwire's baseline:
`3 stores, 30 mail folders, identified 11 folder(s)/308 item(s)` - the bystander's 300 items read item by
item in 4 folders, the identity store's 8 in 7; its post-run census: `0 failure(s), 1 note(s)`, the note
the hub's `Archive` folder the move tests make (`folder newly enumerated`). `post-suite: 0 tagged
artifacts (incl. Archive), 0 test folders` from run 2 on; no fixture clean-up failed, so the signature
snapshots of the Phase-4 and signature-management fixtures held. **Run 1 did not end clean:** two
tagged move seeds stayed in the hub's `Archive` (the sweeps could not find a PST's Archive folder, F2)
and one tagged send round-trip stayed undelivered in the sink; the restore to `CP-12B` before run 2
removed both. One `PROVED NOTHING` line every run, the same one:
`OutlookAvailabilityLiveTests.ATransientOutlookState_AnswersFastAndCarriesRetryGuidance` - "the
retry-guidance check iterated nothing - this machine has no a transient Outlook state (starting, not
responding or unavailable) for list_accounts to report".

**What failed, and what each was.**

| # | Failing test(s) | Cause | Kind | Fix |
| --- | --- | --- | --- | --- |
| F1 | `LiveSweepScopeTests.ControlledCorpus`, `LiveDraftOptionsTests.ForwardDraft`, `Phase5LiveMcpToolShapeTests.SendTool` | the sink probe, the Outbox check and the corpus freshness check were armed only by `LivePhase1Fixture`, which this guest's filter never selects - so no arrival wait nudged Outlook, and the sink's log shows the POP3 fetch closing before the SMTP submission was stored; two of the three waited through private, non-nudging copies of the wait | harness | `f290845` - armed in `LiveStoreCountTripwire.EnsureBaseline`, every sender waits through `LiveInboxArrival` |
| F2 | `LiveMoveArchiveTests.MoveChain`, `MoveArchiveLiveMcpToolTests`, `LiveFolderScopeTests` x2 (and run 1's two left-behind seeds) | a PST's Archive folder, made by `GetDefaultFolder(39)`, is designated in block `0x800F` of the Inbox's `PR_ADDITIONAL_REN_ENTRYIDS_EX` - undocumented, measured byte for byte - and not in `PR_IPM_ARCHIVE_ENTRYID` | product | `5a2f9da` |
| F3 | `LiveResumableScanTests.APagedScan` | Q11 measured: a `Table` column added by its explicit name reports LOCAL time, by its namespace reference UTC; the scan read both as UTC, so its cursor sat one offset late and re-admitted ordinal 24 | product | `b09041b` |
| F4 | `LiveExhaustiveSearchTests.Exhaustive_FolderBounded` | asked for "exactly that folder" without `IncludeSubfolders=false`: 3 folders scanned, 42 hits against 24 | test | `7bb4626` - asks for the folder alone, and checks the default against the subtree's own ground truth |
| F5 | `LiveSweepCacheTests.RapidSearches` | the sweep cache is keyed on the index frontier, so with no indexed mail it never hits - by design (`T1/SweepCacheKeyTests`) | test trait | `7bb4626` - `Requires=SearchIndex` |
| F6 | `OutlookHealthLiveToolShapeTests.CarriesTheFreshnessBlock` | required advice whenever the index is reachable; a reachable index with no mail is reported as a problem | test | `7bb4626` |
| F7 | `LiveUpdateDiscardTests.DiscardDraft` | a PST keeps a draft's EntryID across the soft delete (run 2: `parent='Deleted Items', re-located id is the same id`); the re-locate scan excluded the old id | product | `593e668` |
| F8 | `LiveDraftOptionsTests.DerivedDrafts` | the renamed reply's ConversationId differs from the seed's and the plain reply's. Measured in E4b (`T2/ConversationIdHashes`): it is MD5 over the upper-cased KEPT topic in UTF-16LE - not a hash of the new subject - while the seed's and the plain reply's are their index header's GUID; `PR_CONVERSATION_ID` refuses a write | product, decided | three attempts were taken out again first (`67c4afb`, `6bd9d60`, `0f4bbfa`). DECIDED (coordinator, on the maintainer's behalf; `QUESTIONS.md`): the same-id promise is Exchange's; elsewhere the test holds the renamed reply - and now the renamed forward - to the kept topic's hash, and the subject hint says so (`b4aa28f`). Passed from run 13 on |
| F9 | `LiveDisconnectRecoveryTests` | "D49 regression: Outlook exited when its last window closed". Two causes, measured with Explorer-only probes and a 90-second subset run: Office 2024's `Explorers.Add` on the folder the hidden lifetime pin shows hands back THE PIN, so the promotion displayed it (`explorers=1`); and Office 2024 raises `Application.Quit` when the user closes the last visible window, so a session that started Outlook drops and closes its pin as it leaves (two scratch builds: without the Quit sink the test passed, with only the first fix it failed) | product, Office LTSC 2024; decided | `ffc6529` - the show-me path never returns an Explorer that already existed (`T1/ShowMeExplorerPinTests`); DECIDED for the Quit (`QUESTIONS.md`): keep honouring it - it is what lets the user's own Exit end an Outlook OutlookAI started - and `fd2c58b` holds the test to exactly that: an exit passes only if the promoting session started Outlook, two Explorers stood before the close and the quit event, not a process exit, ended the session; the reattach must still work. The earlier attempt (`01d81ec`) stays reverted |
| F11 | (none - the hub rebuild before run 11, and E4's first attempt) | `corpus-teardown` finished cleanly, then the move of its manifest into `hub-history\` met "being used by another process"; the run stopped with the hub torn down | test bed | `292dbd8` - the move is retried for up to 60 s on a sharing or lock violation only |
| F12 | `LiveDisconnectRecoveryTests`, step 3b (first reached in a D49-only run) | the degraded-search check needs index results to fall back to; with no catalog the search failed with `0x80041820` | test | `58d3707` - the step goes through `LivePopulationCoverage.Require` over the indexed stores: `PROVED NOTHING` here, refused on a Production profile, unchanged where the hub is indexed |
| F10 | `LiveCreatedFolderTests` (run 2, its first run anywhere) | the Drafts folder a reply made in the throwaway data file is designated in `PR_IPM_DRAFTS_ENTRYID` on the store's TRUE root folder (NID `0x122`, the parent of the IPM subtree) and nowhere the lookup read | product | `5ac1d85` - passed in run 3 |

**Section 8 item 25, the created-folder proof's first runs.** `Reset-ThrowawayStore.ps1 -Execute`,
every run: `verify : 1 store(s) named 'throwaway@vm.invalid', file
C:\OutlookAI-Tier\Throwaway\throwaway-<stamp>.pst, Drafts designation NotFound, top-level folders [Deleted
Items]`, and `detach : none` - each run started from `CP-12B`, which has no throwaway. Run 3's four lines:
`before: 'throwaway@vm.invalid' Drafts=DefaultFolderAbsent`, `reply_draft: store='throwaway@vm.invalid'
folder='Drafts' createdFolders=[throwaway@vm.invalid/Drafts]`, `after: the non-creating lookup sees
Drafts='Drafts'`, `discard_draft: discarded=True to='Deleted Items' createdFolders=[]`. Run 2's `after:`
line was `DefaultFolderAbsent` - Q85's open question answered: Outlook designates a Drafts folder it
makes in a data file with no Inbox on that file's true root folder (F10). A read-only probe after run 2
found the throwaway's Drafts and Deleted Items empty and nothing tagged in it. **The second run's
detach**, read after run 5 without a restore (Outlook first started on the tier profile with
`Start-OutlookUnelevated.ps1`, because the suite's lifecycle tests end with Outlook closed and the
script refuses to start one): `detach : 'throwaway@vm.invalid' <- ...\throwaway-20261003T081352Z.pst`,
`attach : ...\throwaway-20261003T082322Z.pst`, the same `verify` line, then `delete :
...\throwaway-20261003T081352Z.pst` - the freshly started Outlook did not hold the old file, so it went.

**What stayed red.** Nothing, from run 18 on. F8 and F9 were decided rather than loosened - each
test now holds its store kind or its Office build to what was measured, and says so in its output -
and two checks print `PROVED NOTHING` every run on this guest, by design: the retry-guidance check
(no transient Outlook state to report) and D49's step 3b (no index to degrade to). Q74 C3's PST half
(`ShortDecodedId_OpensAsTheItemItself_OnAPstStore`) carries `Requires=SearchIndex`, so this guest's
filter never selects it. Q101's Inspector/Outbox ordering in `LiveDisconnectRecoveryTests` never bit:
every run read `no inspectors, outbox empty` before it closed the parked window.

### 4.1f The guest clock with time synchronisation off - `OutlookAI-Unindexed`, 2026-10-03 (Q130, measured only)

**Why this section exists.** Q108 (`Docs/overnight-review-2026-10-03.md`) asks whether freezing an
Outlook guest's clock - Hyper-V time synchronisation off, every run from a checkpoint - would stop the
test data ever going stale. This is what the guest's clock actually does, measured from the host over
PowerShell Direct (guest UTC minus host UTC, the host's reading at the midpoint of the call). Nothing was
left changed: every step began and ended on `CP-13B-LIVE-GREEN` with time sync back ON, the one
temporary checkpoint (`CP-TEMP-CLOCK-PROBE-a24876cb`) was deleted, and the guest was saved with its 20
checkpoints. No Outlook was started and no store opened. Raw logs: `.work\g1-d62\logs\` of the agent
worktree `a24876cb`, phases `p4g2b` and `p4b`.

| Step | Guest clock | Reading |
| --- | --- | --- |
| `CP-13B-LIVE-GREEN` restored, time sync ON (its own setting) | agrees with the host | skew 0.2 s at the first answer, 3 s after the restore; 0.1 s 20 s later |
| Time sync OFF, `Set-Date` 10:00Z | stays where it was put | skew -25,443 s, unchanged 90 s later; `w32tm /query /source`: "the service has not been started" - nothing else sets the clock on this guest |
| Saved 90 s, resumed | stopped while saved | lost 94.9 s against the host |
| Checkpoint taken with time sync OFF; restored twice, 60 s apart | **the same instant every restore** | 1.7 s and 1.8 s before the checkpoint's instant; time sync still OFF after each restore - the setting travels with the checkpoint |
| `CP-13B-LIVE-GREEN` restored after that | back to the host's time | time sync ON again (that checkpoint's own setting); skew 0.1 s |
| COLD boot (graceful `Stop-VM`, `Start-VM`) from the time-sync-OFF checkpoint | host time plus the offset the guest last WROTE | before: skew -25,702 s (the restored instant); after: -25,445 s - the `Set-Date` offset, not the restored instant |
| OS restart through `Testbed/host/Restart-Guest.ps1`, time sync OFF, no restore before it | keeps the offset it had | moved by -2.2 s; the restart script proved the restart by the later boot time as usual |
| Clock set to 2026-08-01, then signatures checked | - | Authenticode of `dotnet.exe` and of the staged SDK installer `Valid`; `dotnet nuget verify --all` of `Microsoft.Extensions.Logging.Abstractions 10.0.10` exit 0, only the offline revocation warnings |

**What it means for Q108 (a).** A checkpoint taken with time sync off starts every run at the same instant,
and saving or resuming does not move it - the mechanism (a) needs is real. Two rules come with it: a run
must not restart or cold-boot the guest after the restore (both leave the frozen instant - the restart case
after a restore inferred from the cold boot, not measured), and the build VM can never be frozen (its runner
requires the host's clock within 2 s). Not measured: installing a freshly built add-in on a frozen guest,
and MSBuild with files the host dated after the guest's clock. *Both measured since, and the restart after
a restore too - section 4.5.*

### 4.2 The indexed guest's build-out - `OutlookAI-Indexed`, 2026-09-24 and 2026-09-27

**Why this section exists.** The same build-out as section 4.1, on the guest that must stay
INDEXED: from `CP-10-INDEXED` (20,048 Outlook rows over three stores, the catalog IDLE), one step
at a time, each checkpointed, with `Set-OutlookIndexingDisabled.ps1 -Verify -SettleMinutes 1
-MinimumOutlookRows 20000` after every step that started Outlook or installed something. No
population was built and the live tier did not run: both wait on the generator and harness fixes.
*(2026-09-27: both have landed - the harness's on 2026-09-24 (Q84, `4f82004`), generator v2 and its
guest fixes by 2026-09-27 (section 4.1a). Section 4.2b's rebuild ran neither a population nor the
tier. **2026-10-03: the populations - with undated contacts, Q98 (f) - and this guest's settings are
built, section 4.2c; the live run is not.**)*
Raw logs: `.work\g1-buildout\` in the main checkout.

**Every COM caller ran at Outlook's own integrity level.** Outlook on this guest is started NOT
elevated (`Testbed/guest/Start-OutlookUnelevated.ps1`), because only that Outlook feeds the index
(section 8 item 22). An elevated caller cannot attach to it, and `Register-InteractiveTask.ps1`
still registered only `RunLevel Highest` in the tree these steps ran from, so the steps that talk
to Outlook - `Set-OutlookProgrammaticAccess.ps1 -Verify`, `Set-AccountSignature.ps1`, the product's
read-back - ran in a one-shot `RunLevel Limited` interactive task, a scratch helper (both tokens
read integrity 0x2000). Closing Outlook was `Testbed/host/Restart-Guest.ps1 -Execute
-CancelLogonPrompt` every time; it picks its quit task's run level from Outlook's own token. The
committed route landed on master meanwhile - `Register-InteractiveTask.ps1 -RunLevel Limited`, the
live tier's route (section 4 and section 8 item 22) - and it replaces that scratch helper: a
rebuild runs these steps through it.

| Step | What ran | Verdict | Index after | Checkpoint |
| --- | --- | --- | --- | --- |
| 1. First-run settings | `Set-OfficeFirstRunSuppressed.ps1 -Verify`, `-Execute`, `-Verify` | 10 OK and 3 FAIL - the same three classic-Outlook values as on the other guest, its `Options\General` key missing altogether - then 13 of 13 | no Outlook start; `INDEXED` at the baseline | `CP-11-FIRSTRUN-REPAIRED` |
| 2. Programmatic access (Q80) | `-SelfTest` 43/0, `-Execute`, Outlook started unelevated on the tier profile, `-Verify` from a Limited task | `NO-PROMPT`, both accounts' `SmtpAddress` in 1.5 s all told. **No control, and one would have proved nothing here**: Windows Security Center reported the antivirus up to date (`0x061100`) while Defender reported its signatures 372 days old - so the guard's trigger was absent (the script's banner) | `INDEXED`, 20,048 | `CP-12-PROGRAMMATIC-ACCESS` |
| 3. Mail sink | `Install-MailSink.ps1 -LogLevel debug -Execute`, graceful restart, `-Verify`; `New-TierProfile.ps1 -StoreSinkPassword -Execute`; Outlook started unelevated | `SINK-READY` twice, the sink up with the boot; a password stored on both accounts; no logon dialog; the sink logged `read USER tier` and `read USER identity`, each with `read PASS any-value` and a `STAT` | `INDEXED`, 20,048 | `CP-13-MAIL-SINK` |
| 4. Add-in | `Publish-AddInPayload.ps1` on the host from `98e050e` (host unchanged), `Install-OutlookAIAddIn.ps1` `-SelfTest` 85/0, `-Verify`, `-Execute`, graceful restart, `-Execute` | `NOT-INSTALLED`, then `ADDIN-READY` twice: VSTO `v4R` absent, then 10.0.60917; tuning state 8.3 s and 3.6 s after the start; the trust entry kept the second time; the index exclusion state `UNCHANGED` by both runs. The installer's first-run Outlook is started over COM from the elevated task, so it is elevated - it did not disturb the index | `INDEXED`, 20,048 | `CP-14-ADDIN-READY` |
| 5. Signature | the server published from `4e23866`'s McpServer tree; Outlook started unelevated; the two profile entries read; `Set-AccountSignature.ps1 -Account identity@vm.invalid -Execute` from a Limited task; the entries read again; the product's `list_signatures` asked separately | `New Signature` = `Identity` on `00000004` (clsid `{ED475411-...}`, `Email` `identity@vm.invalid`), nothing on `00000005` (the `identity.pst` data-file entry); `Identity.htm`, `.rtf`, `.txt` written; the separate read-back: account `identity@vm.invalid`, newMessage `Identity`, and `tier@vm.invalid` with none | `INDEXED`, 20,048 | - |
| 6. Suite | `Publish-LiveTierPayload.ps1 -Ref 4e23866` (54 packages; the feed restores everything), the guest's old `src` and feed kept as `*.2026-09-17` and fresh ones expanded; `Install-DotnetSdk.ps1 -Execute` - the SDK present, the installer skipped; it writes the `NuGet.config` a fresh `src` lacks, without which `-Verify` stops at `SDK-ONLY` - then `-Verify` from a new session | `TEST-READY` both times: **3,041 tests discovered**, 17 run and passed, 35 s and 11 s; the `McpServerExePath` the guest's build bakes in names an exe that exists | `INDEXED`, 20,048 | `CP-15-SIGNATURE-SUITE-STAGED` |

**The interruption, and why step 5 ran twice.** On 2026-09-24 the first pass of step 5, with a
server from `98e050e` that already wrote the account entry, also ran a scratch COM probe: an
unsaved MailItem in the identity store's Drafts with `SendUsingAccount` pinned, `GetInspector`,
`HTMLBody` read, `Close(olDiscard)`. It saw the signature inserted, and a census of every folder of
both stores before and after found nothing added. The session then stopped, and both guests were
saved to disk for three days. **That probe created an item from ad-hoc code, which mailbox-safety
rule 1 forbids on the guests as well**, so on 2026-09-27 the guest was restored to
`CP-14-ADDIN-READY` - discarding the probe and everything after it - and step 5 was redone creating
no item at all; `CP-15` descends from that pass. Whether Outlook inserts the signature into a new
mail is left to the live tier (`LiveDraftOptionsTests`). The resumed guest, before the restore, still
read `INDEXED`, with Outlook running from the probe's pass and its clock three days behind until
time sync caught it up within seconds; after the restore, `SINK-READY` and `INDEXED` both held with
nothing repaired.

**Also found:** `Install-MailSink.ps1 -Verify` on a guest resumed from saved state reported the sink
as started "-199,807 s after boot" - `LastBootUpTime` moves with the clock jump. Cosmetic; every
check passed.

### 4.2b The indexed guest rebuilt on the Q87 route - `OutlookAI-Indexed`, 2026-09-27

**Why this section exists.** The identity account of every checkpoint above, `CP-09-IDENTITY-ACCOUNT`
on, delivers into the hidden root of an `AddStoreEx` PST, and a minted store cannot be attached under
the same name where that one already is (section 2.8b). So the guest was rebuilt from
`CP-08B-RESTORED-BEFORE-IDENTITY` on the Q87 route and taken back through section 4.2's steps, from
master `af56efc`. Every Outlook start NOT elevated but one - the add-in installer's own first-run
Outlook in step 4.4, which it starts over COM from its `RunLevel Highest` task, as it did in section
4.2 - everything that attached to Outlook through `Register-InteractiveTask.ps1 -RunLevel Limited`,
every close `Testbed/host/Restart-Guest.ps1`. No population, no live tier. Raw logs:
`.work\g1-rebuild-q87\` in the main checkout.

| Step | What ran | Verdict | Checkpoint |
| --- | --- | --- | --- |
| Restore | `Restore-VMSnapshot CP-08B-RESTORED-BEFORE-IDENTITY`, master's guest scripts staged (every hash equal) | default `OutlookAI-Tier`, no identity account, no `ImportPRF` pending, `PreventIndexingOutlook` absent, no scope rule, 0 Outlook rows (`NOT-IN-SCOPE`); WSC `0x061100` | - |
| 1. Identity (Q87) | `-Phase Mint -Execute`; `OUTLOOK.EXE /PIM IdentityMint` started by a `-RunLevel Limited` job; CaptureMint; `Rename-OutlookStore.ps1`; graceful restart; `-Phase Import -Execute`; the tier profile started (it imported the account); the attach; **CaptureStore - blocked, see below**; restart; `Set-OutlookProgrammaticAccess.ps1 -Execute` (brought forward from step 4.2); start; CaptureStore; restart; Bind; `New-TierProfile.ps1 -StoreSinkPassword -Execute`; start; `-Phase Verify -TrySmtpAddress` | as section 4.1b, line for line: the minted store `Outlook Data File`, mask `0xFF`, Inbox designated and visible, 14 folders, 0 items; CaptureStore PASSED from the tier profile; Verify `delivers into 'Inbox' (visible=True, PST node id 0x8082)`, 2 accounts on 2 stores, both `SmtpAddress` reads, no logon dialog | `CP-09C-IDENTITY-REAL-INBOX` |
| 2. Indexed | `Set-OutlookIndexingDisabled.ps1 -Enable -Execute`; `Start-OutlookUnelevated.ps1 -Profile CorpusProfile` and `-Verify -SettleMinutes 1 -WaitMinutes 30 -MinimumOutlookRows 20000`; restart; the same on `OutlookAI-Tier`; restart | the scope was already IN - the step-1 Outlooks, NOT elevated, had added the rule themselves (`included=True reason=USER`, 32 rows); `-Enable` wrote policy 0 and re-asserted the rule; with the corpus profile open the Outlook rows reached 20,000 10.3 min in and 20,059 by 11.4 min, `INDEXED` at 12.5 min: **20,059 rows** - corpus 20,028, tier 16, identity 15 - unchanged by the tier profile's start after it | `CP-10C-INDEXED` |
| 3. Q92 evidence | the Q92 agent's read-only probe (`DIRECTORY='mapi16://{SID}/'`, `SCOPE` counts), and the store hash computed on the host from each store's `StoreID` | below | - |
| 4.1 First-run | `Set-OfficeFirstRunSuppressed.ps1` `-Verify`, `-Execute`, `-Verify` | 10 OK and the same 3 FAIL, then 13 of 13 | `CP-11C-FIRSTRUN-REPAIRED` |
| 4.2 Programmatic access | (`-Execute` in step 1) `-Verify` from a Limited job | `NO-PROMPT`, both `SmtpAddress` in 1.1 s | `CP-12C-PROGRAMMATIC-ACCESS` |
| 4.3 Mail sink | `Install-MailSink.ps1 -LogLevel debug -Execute`; restart; `-Verify`; `-StoreSinkPassword -Execute` (both `ours` already); the tier profile started | `SINK-READY` twice, the sink up 4 s after the boot; no logon dialog; `read USER tier` and `read USER identity`, each with `read PASS any-value` and a `STAT` | `CP-13C-MAIL-SINK` |
| 4.4 Add-in | built from `af56efc` (host unchanged); `-SelfTest` 85/0, `-Verify`, `-Execute`, restart, `-Execute` - from a `RunLevel Highest` task, as the VSTO runtime needs, so the installer's first-run Outlook is the one start here that is not unelevated | `NOT-INSTALLED`, then `ADDIN-READY` twice (tuning state 3.5 s after the start both times; trust entry kept the second time; the index exclusion state `UNCHANGED`); the installer's headless Outlook had exited by itself within 11 s and 20 s; `INDEXED`, 20,059, after it | `CP-14C-ADDIN-READY` |
| 4.5 Signature | the server published from `af56efc`; entries `00000004` and `00000005` read; `Set-AccountSignature.ps1 -Execute` from a Limited job; read again; `list_signatures` asked separately | `New Signature` = `Identity` on `00000004` (the POP3 account) and nothing on `00000005` (the minted file's entry, `Account Name` `identity@vm.invalid`); the three files written; the read-back `identity@vm.invalid` newMessage `Identity`. No item created | - |
| 4.6 Suite | `Publish-LiveTierPayload.ps1 -Ref af56efc`, old `src` and feed kept as `*.2026-09-17`, fresh ones expanded, `Install-DotnetSdk.ps1 -Execute` (writes `NuGet.config`), `-Verify` from a new session | `TEST-READY` twice: **3,105 tests discovered**, 17 run; `McpServerExePath` names an exe that exists | `CP-15C-SIGNATURE-SUITE-STAGED` |

`INDEXED` held at every check after step 2 - 20,059 rows, the catalog `IDLE`, nothing queued - with
Outlook closed and with it running NOT elevated on `OutlookAI-Tier`, and after the add-in
installer's first-run Outlook (started from its elevated task; that Outlook's own token was not
read).

**That one elevated start is gone from the procedure since 2026-10-03 (Q100, option 3), and its
guest proof is PENDING.** `Install-OutlookAIAddIn.ps1` now runs as `-Phase Install` from the
elevated task, which never starts Outlook, and `-Phase FirstRun` from a `-RunLevel Limited` one,
which starts Outlook NOT elevated and reads that Outlook's token (section 2.3; `Testbed/README.md`
section 1, steps 5b and 7c). Step 4.4 above is the old single `-Execute`, and this guest still
carries what it installed - `-Verify` reads it as before, since it has no install record. The
two-phase form has not run here or on `OutlookAI-Unindexed`: both were busy when it was written, so
it waits for the next rebuild of this guest or a free slot, and section 2.3 lists what that run must
record. Until it has run, "every Outlook start on this guest NOT elevated" holds for the procedure
as written, not yet for a build that followed it.

**The guard prompt, and why Q80 came first.** CaptureStore's first attempt, about two minutes after
the restart that closed the mint profile, raised the Object Model Guard prompt ("A program is trying
to access email address information stored in Outlook") and its COM read blocked behind it - the
phase has no deadline - until the job's own time limit ended the job 15 minutes later; the orphaned
prompt was then answered **Deny** through its own `WM_COMMAND` (README step 4d; Deny grants nothing).
Windows Security Center read `0x061100` - signatures up to date - both before and after, its
timestamp 16 s after that read began. So a guest without `Set-OutlookProgrammaticAccess.ps1` can
prompt at any moment, whatever one WSC reading says, and the route of section 2.8b needs the values
first. `OutlookAI-Unindexed` never met this: it had them from `CP-07`.

**The display names, stage by stage.** Over COM (`Store.DisplayName`, the root's name the same): after
the mint `Outlook Data File`; after the rename `identity@vm.invalid`; after the attach
`identity@vm.invalid`; at CaptureStore and Verify `identity@vm.invalid`. In the registry, each
profile's service section `PR_DISPLAY_NAME_W`: `IdentityMint` - `Outlook Data File` after the mint,
`identity@vm.invalid` after the rename; `OutlookAI-Tier` - `identity@vm.invalid` from the attach on.

**What the index says (Q92).** `DIRECTORY='mapi16://{SID}/'` lists three store roots in 2 ms:
`identity@vm.invalid($be889d8b)`, `Outlook Data File($23a27f0d)` (the corpus, another profile's store)
and `tier@vm.invalid($93f42b43)`. The identity store is filed under **its address and a NEW hash**,
`be889d8b` - section 8 item 24's rule exactly: the name is the store's own `PR_DISPLAY_NAME`, now the
address, and the hash follows the file's path, which moved from `identity.pst` (`b25ac20a`) to
`Outlook Data File - IdentityMint.pst`. The host reproduces all three hashes (`be889d8b`, `93f42b43`,
`23a27f0d`) - the identity store's from the `Store.StoreID` CaptureStore recorded, the other two
from their profile `PR_ENTRYID`s, which item 24 found equal to `Store.StoreID` for a PST. **No two stores share a name any more** - before the rebuild the
identity store and the corpus were both `Outlook Data File`. The mint profile mounts the same file, so
it adds no fourth root. The identity store's 15 rows are all folders; the product's discovery sample
(`TOP 2000 ... Kind='email'`) saw only the corpus, as section 8 item 24 found.

### 4.2c The populations built - `OutlookAI-Indexed`, 2026-10-03 (Q98 (f))

**Why this section exists.** The indexed guest's populations, the morning Q98 (f) was measured and
decided (section 3b): undated CONTACTS, on this guest only. From `CP-15C-SIGNATURE-SUITE-STAGED`, on
this guest only: the bystander, the three populations, the store names read off the tier profile, the
settings, the hub rebuild with its index wait, a checkpoint - and no live run, which another agent runs.
**Every Outlook start NOT elevated** and every step that attaches to Outlook in a `RunLevel Limited` job
(each printed `run level: Limited - this token is NOT elevated`), every close
`Testbed/host/Restart-Guest.ps1 -Execute`, every mailbox write through the corpus tool or
`Add-OutlookPstStore.ps1`. The tools were published from this work's branch (`6d01e72`, the
contacts option); master had not got it yet. Raw logs: `.work\g1-cp16c\` in the main checkout.

| Step | What ran | Verdict |
| --- | --- | --- |
| Restage | `Publish-GuestPayload.ps1` at `6d01e72` on the host, the tools expanded on the guest (the old ones kept as `tools.2b5e2c6`, the server left at `af56efc`'s), ten guest scripts copied | tools `99.99.99.0+6d01e7288f9d46bab8bebd956eac31c18dacfa90`; every script's hash the same on guest and host (`Reset-HubPopulation.ps1` `BFBE3B7B56FB43E5`, `6d01e72`'s; the committed one differs from it in its banner comment alone, rewritten after the run) |
| Baseline | `Set-OutlookIndexingDisabled.ps1 -Verify -SettleMinutes 1 -MinimumOutlookRows 20000`, Outlook closed | `INDEXED`, 20,059 rows: Corpus A 20,028, tier 16, identity 15 |
| 1. The bystander | `Start-OutlookUnelevated.ps1 -Profile OutlookAI-Tier`; `Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName bystander@vm.invalid -Path C:\OutlookAI-Tier\bystander.pst -Execute`, then `-ListOnly` | `before : Store.DisplayName='Outlook Data File'`, `after : Store.DisplayName='bystander@vm.invalid'`; the tier profile holds 3 stores - `tier@vm.invalid` (`Outlook.pst`), `identity@vm.invalid` (`Outlook Data File - IdentityMint.pst`), `bystander@vm.invalid` (`bystander.pst`) - and 2 accounts. **Not Corpus A**, which only the corpus profile mounts |
| 2. The corpus profile | `Restart-Guest.ps1`; `Set-DefaultOutlookProfile.ps1 -Name CorpusProfile -Execute`; `Start-OutlookUnelevated.ps1 -Profile CorpusProfile`; the three stores attached by path while empty; `corpus-folders` of all four | every name byte-identical; 4 stores, 0 accounts; hub every default folder, all `items=0`; the bystander `Deleted Items` only; the identity store all but Junk Email, all `items=0`; Corpus A `Inbox 10912`, `Sent Items 4964`, `Deleted Items 2473`, `Junk Email 1663`, `Drafts 0` - its Deleted Items twelve over its plan's 2,461, at `CP-15C` already |
| 3-5. Builds | `corpus-probe`, `corpus-build` dry, `--execute`: `hub-indexed` 8181 and `bystander-indexed` 8282 with `--undated-contacts`, `identity-indexed` 8383 without | below |
| 6. After | `corpus-folders` of all four | hub `Inbox 24` (`-Projects 6`, `-Notices 12`), `Sent Items 14`, `Contacts 12`; bystander `OutlookAI-Corpus-Folder-6 172` (`-Projects 40`, `-Suppliers 32`), `OutlookAI-Corpus-Folder-10 42 contacts`, `-5 56`, still no Inbox, Drafts or Junk Email; identity `Inbox 5`, `Sent Items 3`; Corpus A unchanged |
| 7. The index | `corpus-indexed` for each, `--wait-seconds 900`, Outlook still up NOT elevated; then `-Verify` | each complete at its first ask - below; `INDEXED`, 20,522 rows: bystander 374, tier 97, identity 23, Corpus A 20,028 |
| 8. Settings | `Restart-Guest.ps1`; `testbed.json`'s section filled from step 1's read; `New-LiveTestSettings.ps1 -VMName OutlookAI-Indexed`; `Copy-ToGuest.ps1` | rendered and admitted - Portable, hub `tier@vm.invalid`, 3 watched, indexed `tier@vm.invalid, bystander@vm.invalid`, bystander `bystander@vm.invalid`, throwaway `throwaway@vm.invalid`, probe term `invoice`, the SF-6 probe on the hub, **no corpus**, the sink on loopback; SHA-256 `BBC18D92A050847B53547A2989569B3EF980C8486DEEDFAA4D1B84034B2697AE` on host and guest |
| 9. Hub rebuild | `Reset-HubPopulation.ps1 -SelfTest` (Windows PowerShell 5.1), the dry run, then `-Execute` through `Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 3600` | `99 assertion(s), 0 failure(s)`; the dry run's `undated     the hub carries undated CONTACTS (Q98 (f)) - every verb gets --undated-contacts` and `run level   NOT elevated`; `-Execute` below |
| 10. State | `Restart-Guest.ps1`; the state read; `-Verify` | `outlook processes: 0`, `DefaultProfile: 'OutlookAI-Tier'`, `ImportPRF: ''`; `INDEXED`, 20,524 rows - the tier store's 99 two more than before, the folder rows of the two emptied subfolders the teardown moved into its Deleted Items (section 3b item 7) |

**What the builds printed** - the hub's whole run; the bystander and the identity store printed the same
lines but their own target folders and counts:

```
== undated probe ==
  undated in the INDEX, not the store (Q98 (f)): no delivery-time removal is attempted, and the store's date is reported, not judged - corpus-indexed checks the index
  contact      folder=True inFolder=True tag=True undated=False tableUndated=False class=True inTargetStore=True
Undated probe verified for contact: each lands in its own folder of the target store, keeps its tag and its class. ...
  undated items         : contact=12 - dated in the store, UNDATED in the index (Q98 (f)), for LiveOrderKeyCollationTests on the indexed guest
Build finished: created 68, already present 0, failed 0, 41,316 body bytes in 00:00:03 (19.4 items/s).
Census: 68 item(s) found for 68 planned; per folder found/planned: Sent Items=14/14, Inbox=24/24, Contacts=12/12, Inbox/OutlookAI-Corpus-Folder-Projects=6/6, Inbox/OutlookAI-Corpus-Folder-Notices=12/12. Every ordinal exists exactly once, in the folder the plan names.
Population read-back: 68 of 68 item(s) read; 4 of 4 conversation(s) grouped by the store under one id. ... 12 of them UNDATED IN THE INDEX - the right kind of item; the store gives 12 of them a delivery time Outlook will not remove, which the index does not use for this kind (Q98 (f)) - corpus-indexed checks the index.

bystander (anchor 2026-10-03T06:38:42Z):
Build finished: created 342, already present 0, failed 0, 263,818 body bytes in 00:00:14 (24.4 items/s).
Census: 342 item(s) found for 342 planned; per folder found/planned: Sent Items=56/56, Inbox=172/172, Contacts=42/42, Inbox/OutlookAI-Corpus-Folder-Projects=40/40, Inbox/OutlookAI-Corpus-Folder-Suppliers=32/32. ...
Population read-back: 342 of 342 item(s) read; 6 of 6 conversation(s) grouped by the store under one id. ... 42 of them UNDATED IN THE INDEX ...

identity (anchor 2026-10-03T06:39:44Z):
Undated probe: this population carries no undated item; nothing to probe.
Build finished: created 8, already present 0, failed 0, 5,501 body bytes in 00:00:00 (22.3 items/s).
Census: 8 item(s) found for 8 planned; per folder found/planned: Sent Items=3/3, Inbox=5/5. ...

corpus-indexed:
[    1 s] Index coverage: 68 of 68 population item(s) indexed (12 of them undated; 12 of those in the index, 12 with no received date); newest planned 2026-10-03T06:36:52Z, newest indexed 2026-10-03T06:36:52Z.
[    0 s] Index coverage: 342 of 342 population item(s) indexed (42 of them undated; 42 of those in the index, 42 with no received date); newest planned 2026-10-02T06:38:42Z, newest indexed 2026-10-02T06:38:42Z.
[    0 s] Index coverage: 8 of 8 population item(s) indexed (0 of them undated); newest planned 2026-09-27T23:13:42Z, newest indexed 2026-09-27T23:13:42Z.
```

No sweep found a probe item in any other store, before the probes, after them or after the teardown. The
bystander's contacts went into a typed stand-in, `OutlookAI-Corpus-Folder-10`, recorded in its manifest
(348 lines: the header, 342 items, 5 folders); the hub's into its own Contacts folder.

**The hub rebuild** (`Reset-HubPopulation.ps1 -Execute`, NOT elevated):

```
  undated     the hub carries undated CONTACTS (Q98 (f)) - every verb gets --undated-contacts
  run level   NOT elevated - the rebuild's Outlook starts at this level
=== corpus-teardown ... --anchor 2026-10-03T06:37:52Z --undated-contacts --manifest C:\OutlookAI-Q5\corpus-hub-indexed.jsonl --execute
Manifest records 68 item(s) and 2 created folder(s).
Teardown: considered 136, deleted 136, refused by rule 0, already gone 0, failed 0, folders removed 2.
Post-teardown scan finds 0 corpus item(s) remaining (expected 0).
=== corpus-build ... --anchor 2026-10-03T06:48:34Z --undated-contacts ... --execute
Build finished: created 68, already present 0, failed 0, 41,316 body bytes in 00:00:03 (22.6 items/s).
  OUTLOOK.EXE left 2s after the Quit.
=== starting Outlook for the live run, on 'OutlookAI-Tier', NOT elevated - this script runs NOT elevated, so directly: ...OUTLOOK.EXE /profile OutlookAI-Tier
[    0 s] Index coverage: 68 of 68 population item(s) indexed (12 of them undated; 12 of those in the index, 12 with no received date); newest planned 2026-10-03T06:47:34Z, newest indexed 2026-10-03T06:47:34Z.
The frontier test can catch a local-time misreading until 2026-10-03T08:42:34Z - 110 min from now (115 min on this guest's UTC offset of 02:00:00). START THE RUN NOW, by Testbed/README.md section 4c, with:
  the opt-in   $env:OUTLOOKAI_LIVE_OPT_IN = 'OAI-INDEXED'
  the filter   --filter "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange"
```

| Checkpoint | State |
| --- | --- |
| `CP-16C-POPULATIONS-V2` (parent `CP-15C-SIGNATURE-SUITE-STAGED`; taken with the guest running, 2026-10-03 08:55 local) | Outlook not running; default profile `OutlookAI-Tier`; no `ImportPRF`; `INDEXED`, 20,524 rows; the hub at anchor `2026-10-03T06:48:34Z` with its twelve contacts (manifest SHA-256 `7A999122...`, its predecessor in `hub-history\`), the bystander at `06:38:42Z` with its forty-two (`E0D55AFB...`), the identity store at `06:39:44Z` (`EB2F0B06...`); Corpus A untouched (`0BCFA3DA...`, 20,001 lines); the live-test settings staged; the tools from `6d01e72`, the server and the suite still `af56efc`'s |

**For the run, which is not part of this:** the suite on the guest is `af56efc`'s; re-stage it from
master first (`Testbed/README.md` step 8b). The run starts with step 9a at `-RunLevel Limited`. With
no corpus declared, the corpus tests print that they
have none, and the index tests measure the hub and the bystander - where `LiveOrderKeyCollationTests`
should now find the 12 and 42 undated rows `corpus-indexed` counted; that run is what shows they do.

**Next, planned here and NOT started: Corpus A at 160,000 items** (the decision of 2026-09-24, question E).
Estimated from tonight's rates: **about 2.5 to 3.5 hours in all**, within the 6-hour bar. The route the
rules leave (section 2.9): a NEW store, `Corpus A`, created empty in the corpus profile and attached to
the tier profile by path while still empty, then `corpus-build --count 160000` into it with a fresh
anchor - not an extension of the 20,000 in `Outlook Data File`, which no script may attach to the tier
profile now that it holds items, and whose 2026-09-15 anchor has left its 7-day window empty anyway.
A store that is not the profile's default builds through `PostAsNote` alone: 19-24 items/s tonight on
the populations' small items (the 2026-09-15 default-store build ran about 36 items/s through
Drafts-then-Move), so **about 1.8 to 2.2 hours of build**. The index keeps up with an unelevated
Outlook as it goes (tonight: every population complete at the first ask; the earlier crawl rate about
2,000 rows a minute, 80 minutes for 160,000 should it lag). The PST grows to about 8.6 GB (Corpus A is
1.07 GB for 20,000); the guest has 96 GB free, the host 284 GB. Then the settings gain the corpus block
and the indexed list its third entry, `corpusIdConvention` records the new seed, anchor and 160,000,
and a checkpoint follows. **BUILT the same morning by exactly that route - section 4.2d, checkpoint
`CP-17C-CORPUS-160K`.** The estimate held: 1 h 21 min of build at 33.0 items/s, and 2 h 32 min from
the first step to the checkpoint; the PST came to 8.5 GB. The settings, the hub rebuild and the
"For the run" note above are superseded there.

### 4.2d Corpus A at 160,000 items - `OutlookAI-Indexed`, 2026-10-03 (question E)

**Why this section exists.** The production-scale corpus decided on 2026-09-24 (question E, option (a);
section 2.9), built the same morning as section 4.2c - decided on the maintainer's behalf (D80 of
`Docs/overnight-review-2026-10-03.md`): guest one would otherwise sit idle while guest two's live run
finds the fixes both guests need. From `CP-16C-POPULATIONS-V2`, on this guest only, by the route
section 4.2c planned: a NEW store, `Corpus A`, created empty in the corpus profile and attached by path
to the tier profile while still empty, then built with a fresh anchor. **Every Outlook start NOT
elevated**, every step that attaches to Outlook in a `RunLevel Limited` job, every close
`Testbed/host/Restart-Guest.ps1 -Execute`, every mailbox write through the corpus tool or
`Add-OutlookPstStore.ps1`. The tools were section 4.2c's (`6d01e72`, the corpus code master has);
`Build-Corpus.ps1` and `Reset-HubPopulation.ps1` were restaged from master (`d2b13a7`) - the guest's
`Build-Corpus.ps1` was still the 2026-09-15 one - and `Set-OutlookIndexingDisabled.ps1` with the fix
below. Raw logs, both manifests and the build's whole output: `.work\g1-cp17c\` in the main checkout.

| Step | What ran | Verdict |
| --- | --- | --- |
| 0. The old manifest aside | Session 0, Outlook and the corpus tool not running: `corpus-vm-indexed.jsonl` moved to `C:\OutlookAI-Q5\corpus-history\vm-indexed.20260915T000000Z.jsonl`, its build log and the 2026-09-15 `Build-Corpus.ps1` beside it; the manifest copied off the guest | SHA-256 `0BCFA3DA...` before and after, on guest and host; nothing named `corpus-vm-indexed.jsonl` left - the id is the new corpus's now, and a manifest is named after its id |
| 1. Create it | `Set-DefaultOutlookProfile.ps1 -Name CorpusProfile -Execute`; `Start-OutlookUnelevated.ps1 -Profile CorpusProfile`; `Add-OutlookPstStore.ps1 -ProfileName CorpusProfile -DisplayName 'Corpus A' -Path C:\OutlookAI-Tier\corpus-a.pst -Execute`, then `-ListOnly`; `corpus-folders` | `decision : AddThenRename`, `after : Store.DisplayName='Corpus A'`; 5 stores, `Accounts.Count : 0`; a 271,360-byte PST holding `Deleted Items items=0` and no other default folder |
| 2. Attach it to the tier profile, empty | `Restart-Guest.ps1`; default `OutlookAI-Tier`; `Start-OutlookUnelevated.ps1 -Profile OutlookAI-Tier`; `Add-OutlookPstStore.ps1 -ProfileName OutlookAI-Tier -DisplayName 'Corpus A' -Path C:\OutlookAI-Tier\corpus-a.pst -Execute`, then `-ListOnly` | `before : Store.DisplayName='Corpus A'` - the name is the file's own (its root folder's), so the second profile needed no rename; the tier profile reads 4 stores - `tier@vm.invalid`, `identity@vm.invalid`, `bystander@vm.invalid`, `Corpus A` - and 2 accounts |
| 3. Dry run | `Restart-Guest.ps1`; default `CorpusProfile`; Outlook NOT elevated, 200 s; `corpus-folders` of all five stores; `Build-Corpus.ps1 -Store 'Corpus A' -CorpusId vm-indexed -Seed 7777 -Anchor 2026-10-03 -Count 160000 -Manifest C:\OutlookAI-Q5\corpus-vm-indexed.jsonl` | the plan below; the preflight `OK - the default profile has no accounts, Outlook is up and warm, and ImportPRF is not set`; `Dry-run complete; 160,000 item(s) would be created` |
| 4. Build | The same with `-Execute`, through `Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 28800 -TaskName OutlookAI-Corpus160k` - a task name of its own, because every job unregisters the shared one; watched from session 0 by the manifest's line count, the lease renewed as it went | 09:38 to 11:12 local: both probes verified, `created 160,000 ... failed 0` in 01:20:53, two censuses clean - below |
| 5. After | `corpus-folders` of all five stores | Corpus A `Deleted Items 19292`, `OutlookAI-Corpus-Folder-6 88037`, `-5 39709`, `OutlookAI-Corpus-Folder-Junk 12962`; the old corpus's `Deleted Items 2461` - twelve fewer, below - and its other folders as before; the hub, the bystander and the identity store unchanged |
| 6. The index | `Set-OutlookIndexingDisabled.ps1 -Verify -SettleMinutes 2 -WaitMinutes 90 -MinimumOutlookRows 180000`, Outlook still up on the corpus profile; again after the fix below; then a read-only census of Corpus A's index scope | first `NO-INDEXER`, wrongly - below; with the fix `INDEXED`, 180,518 rows on both readings, `store Corpus A($996dc7a9): 160006 row(s)`, the catalog `IDLE` with nothing queued; the census: 160,000 corpus rows, ordinals 1-160,000 each exactly once, every one with a received date, and 5 folder rows |
| 7. Settings | `testbed.json` (below); `New-LiveTestSettings.ps1 -VMName OutlookAI-Indexed`; `Copy-ToGuest.ps1`; `corpus-verify ... --window 7 --window 30 --window 60` | rendered and admitted - watched 4, indexed `tier@vm.invalid, bystander@vm.invalid, Corpus A`, bystanders `bystander@vm.invalid, Corpus A`, `corpus vm-indexed in 'Corpus A' ... windows 7/30/60 day(s)`; SHA-256 `9674ED3E7343D23856E7DBEC59F910F53482798F1FC3A45FEB1CFD63DF2E67DA` on host and guest; `Freshness: OK - anchor 2026-10-03T00:00:00Z (never re-anchored) ... Windows now/at-anchor: 7d=12,191/12,849, 30d=24,396/24,596, 60d=39,707/39,903`. Restaged the same day with the windows 30 and 60 only (`688FDB99...`, "How long it stays usable", below) |
| 8. Hub rebuild | `Restart-Guest.ps1`; `Reset-HubPopulation.ps1 -SelfTest`, the dry run, then `-Execute` at `-RunLevel Limited` | `99 assertion(s), 0 failure(s)`; teardown `considered 136, deleted 136 ... folders removed 2`, `0 corpus item(s) remaining`; `created 68`, census 68/68, read-back 68 of 68; `OUTLOOK.EXE left 2s after the Quit`; the index wait `[1 s] 0 of 68`, then `[16 s] 68 of 68 ... 12 with no received date`; anchored `2026-10-03T09:32:50Z`, 109 minutes of margin; exit 0 in 532 s |
| 9. Two profiles, one scope | The tier profile's Outlook, started by step 8, left running; `-Verify` at about 10 and 14 minutes | `INDEXED`, 180,520 rows, still exactly one `Corpus A($996dc7a9)` at 160,006 and one scope for every other store, nothing queued - item 5 of section 3b's "What only a guest can answer", answered |
| 10. State | `Restart-Guest.ps1`; the state; `-Verify` | `outlook processes: 0`, `DefaultProfile: 'OutlookAI-Tier'`, `ImportPRF: ''`, guest C: 85.7 GB free; `INDEXED`, 180,520 rows on both readings |

**What the build printed** (Build-Corpus's own log and the job's output, both in `.work\g1-cp17c\manifests\`):

```
  items                 : 160,000
  body bytes (total)    : 1,766,578,340  (mean 11,041)
  received range        : 2022-10-04T00:00:36Z .. 2026-10-02T23:59:16Z
  per folder            : Deleted Items=19,292, Sent Items=39,709, Inbox=88,037, Junk Email=12,962
  selected by window    : 1d=3,311, 7d=12,849, 30d=24,596, 60d=39,903, 90d=44,713, 365d=88,075
Cross-store residue sweep (before the probes): 12 probe item(s) of 'vm-indexed' found in 'Outlook Data File', which is NOT the target; 12 deleted by the two-key rule.
  target store: NOT the profile's default store - only the rungs whose item is never unsent (PostAsNote) are probed; every other rung files its item in the default store's Drafts first
  target folder: 'OutlookAI-Corpus-Folder-6' - a STAND-IN: the store has no visible Inbox of its own, and keeps none
  PostAsNote                   target=OutlookAI-Corpus-Folder-6 visible=True store=target landedIn=OutlookAI-Corpus-Folder-6 parentMatches=True inFolderTable=True sentFlag=True usable=True
Placement: VERIFIED via PostAsNote. Items will live in the folders the plan names.
  PropertyAccessorDates        requested 2026-09-03T00:00:00Z wrote 2026-09-03T00:00:00Z readBack 2026-09-03T00:00:00Z daslIn=True daslOut=True usable=True
Date fidelity: VERIFIED via PropertyAccessorDates. Received dates drive DASL selection.
Cross-store residue sweep (after the probes): no probe item of 'vm-indexed' in any other store.
  progress: created 50,000, skipped 0, failed 0, remaining 110,000, 549,758,351 body bytes, 00:24:17 elapsed
  progress: created 100,000, skipped 0, failed 0, remaining 60,000, 1,110,739,011 body bytes, 00:50:00 elapsed
Build finished: created 160,000, already present 0, failed 0, 1,766,578,340 body bytes in 01:20:53 (33.0 items/s).
Census: 160,000 item(s) found for 160,000 planned; per folder found/planned: Deleted Items=19,292/19,292, Sent Items=39,709/39,709, Inbox=88,037/88,037, Junk Email=12,962/12,962. Every ordinal exists exactly once, in the folder the plan names.
manifest: 23557073 bytes, 160004 line(s), sha256 AB395B8157AACD0401518CFCFE32DD06C5534EF178B1BA0C090B790C3D944D42
pst: 8479220736 bytes
```

The rate held between 27 and 36 items/s for the whole build, the corpus tool's working set growing
from 53 to 194 MB and Outlook's staying under 90 MB; the guest's C: went from 96.0 to 86.5 GB free and
the host's E: from 136.1 to 125.9 GB. The index kept up behind it: at 10:10, with about 65,000 items
built, it held 53,932 of them and had 10,528 notifications queued, and by 11:15 - fifteen minutes after
the last item - the catalog was `IDLE` with nothing queued. The manifest has 160,000 item lines and three
folder lines: the stand-ins for the plan's Inbox (`OutlookAI-Corpus-Folder-6`), Sent Items (`-5`) and
Junk Email (`OutlookAI-Corpus-Folder-Junk`); a store attached with `AddStoreEx` has only Deleted Items of
its own (section 2.6), and the plan's Deleted Items went there.

**The index check that said NO-INDEXER, and its fix.** `Set-OutlookIndexingDisabled.ps1 -Verify` counts
every Outlook row in the catalog, per store, on each of its two readings. It did that through ADO with the
default command timeout, 30 s, and at 180,989 catalog items both readings failed at exactly 30 s with
`QUERY_E_TIMEDOUT` (`0x80041607`) - and a probe that fails is `NO-INDEXER`, so the script called a
complete, idle index absent. The same had happened once, mid-build, at 10:09 (74,444 rows on the next
reading, in 26 s). The count now runs under its own timeout, `$CountCommandTimeoutSeconds` = 900; with it
each reading took about 70 s and the verdict was `INDEXED`. The TOP 1 probes keep the default.

**The twelve items over plan in the old corpus.** `CP-16C` recorded that the 20,000-item corpus's Deleted
Items held 2,473 against a plan of 2,461 (section 4.2c). The build's cross-store residue sweep found what
they were: probe items of corpus `vm-indexed` - by every sign the 2026-09-15 build's own, whose probes ran
twice against that store (`Build-Corpus.ps1`'s probe and the build's), six rungs each, and whose deletes
left them in its Deleted Items, where nothing swept - and it deleted them by the two-key rule (each one's
subject parses as the corpus's reserved probe ordinal, and its EntryID came from that very scan). That store now holds its plan
exactly - Inbox 10,912, Sent Items 4,964, Deleted Items 2,461, Junk Email 1,663 - and its index scope
20,016 rows, twelve fewer. Nothing else in it changed.

**The old corpus stays where it is, inert** - decided on the maintainer's behalf. It is the corpus
profile's default store, so it cannot be detached, and nothing reads it: the tier profile does not mount
it, and its rows sit under their own scope, `Outlook Data File($23a27f0d)`. Its manifest is in
`corpus-history\` (and off the guest), so `corpus-teardown` can still empty it.

**How long it stays usable.** The tier refuses to start once a window the settings declare selects
nothing (`T2/LiveCorpusFreshness`). As built, this guest declared 7, 30 and 60 days - the example's
windows - and the 7-day window, which held 12,191 that day, would have emptied when the newest item,
`2026-10-02T23:59:16Z`, was a week old: from 2026-10-09 23:59:16 UTC the tier would have refused here
until the corpus was rebuilt, every week. A rebuild is the teardown of 160,000 items (or a new store)
and about 1 h 35 min of `Build-Corpus.ps1`.

**Decided the same day, on the maintainer's behalf: this guest declares the 30- and 60-day windows
only** (D103 of `Docs/overnight-review-2026-10-03.md`, the option recommended above). No live test asks
this corpus a question by window - it is the largest indexed store the latency bounds are timed
against, and a bystander the count tripwire censuses; `Settings.Corpus` is read only by the freshness
check - and the step-10 measurement scripts take their own window per run, so their 7-day default does
not bind the tier. (It does bind them: `Testbed/guest/Measure-SweepCost.ps1` measures the last 7 days
unless told otherwise, and after 2026-10-09 that window of this corpus is empty - on this guest, pass
`-WindowDays` explicitly.) **The corpus is now fresh until 2026-11-01 23:59:16 UTC**, when its newest item leaves
the 30-day window; after that the tier refuses on this guest until it is rebuilt. `Testbed/testbed.json`
records the windows and the date in the guest's corpus block, and `T1/LiveTestSettingsTemplateTests`
pins both: the windows `[30, 60]`, and the freshness check's own verdict on the committed seed, anchor
and count - fresh at 2026-11-01 23:59, the 30-day window empty from 2026-11-02 (and the example's
windows refusing from 2026-10-10). Rendered from `23c7527` and staged, read back on the guest, and
checked there against the manifest with the same windows:

```
after: 1902 bytes, sha256 688FDB9944C081D1D4F6C1F1F89D52BF3E66A779E12DCDE10F1B52628E300226
host file: 1902 bytes, sha256 688FDB9944C081D1D4F6C1F1F89D52BF3E66A779E12DCDE10F1B52628E300226
corpus-verify --corpus-id vm-indexed --seed 7777 --anchor 2026-10-03T00:00:00Z --count 160000 --manifest C:\OutlookAI-Q5\corpus-vm-indexed.jsonl --window 30 --window 60
Manifest records 160,000 item(s), 160,000 of them dated; 159,988 agree on the shift now applied.
Freshness: OK - anchor 2026-10-03T00:00:00Z (never re-anchored), 10h 12m behind the clock. Windows now/at-anchor: 30d=24,374/24,596, 60d=39,695/39,903.
```

**That staged file is the guest's only change since the checkpoint**, so the checkpoint was not
retaken: `CP-17C-CORPUS-160K` holds the settings declaring 7, 30 and 60 (`9674ED3E...`), and the running
guest holds the 30-and-60 file (`688FDB99...`). Revert to the checkpoint, and that file must be staged
again - `New-LiveTestSettings.ps1 -VMName OutlookAI-Indexed`, then the `Copy-ToGuest.ps1` line it prints.

| Checkpoint | State |
| --- | --- |
| `CP-17C-CORPUS-160K` (parent `CP-16C-POPULATIONS-V2`; taken with the guest running, 2026-10-03 11:56 local) | Outlook not running; default profile `OutlookAI-Tier`; no `ImportPRF`; `INDEXED`, 180,520 rows - `Corpus A($996dc7a9)` 160,006; Corpus A at `C:\OutlookAI-Tier\corpus-a.pst` (8,520,360,960 bytes), mounted in both profiles, its manifest `AB395B81...` (160,004 lines); the old corpus inert, its manifest in `corpus-history\`; the hub at anchor `2026-10-03T09:32:50Z` with its twelve contacts (`8A1257E9...`), the bystander and the identity store as at `CP-16C`; the live-test settings with the corpus staged (`9674ED3E...`, windows 7, 30 and 60 - since replaced on the running guest by `688FDB99...`, windows 30 and 60, above); `Set-OutlookIndexingDisabled.ps1` with the count fix, `Build-Corpus.ps1` and `Reset-HubPopulation.ps1` from master; the tools `6d01e72`'s, the server and the suite still `af56efc`'s |

**For the run, which is not part of this:** as section 4.2c says - re-stage the suite from master first
(its hub-freshness check must know the `|u:contacts` marker), then step 9a at `-RunLevel Limited`.
The suite then finds Corpus A in the settings: the freshness check runs at start, the corpus is
the largest indexed store the latency bounds are timed against, and the count tripwire censuses it as a
bystander. **Before 2026-11-01 23:59 UTC**, or after a rebuild.

### 4.2e All three kinds, the sized contest and Corpus A's stand-ins - `OutlookAI-Indexed`, 2026-10-03 (D62 (b), D74, D101)

**Why this section exists.** The maintainer's answers of 2026-10-03 to D62 (all three kinds), D74 ("no
luck") and D101 ("measure if relevant"), on guest one; decisions D126-D133 of
`Docs/overnight-review-2026-10-03.md`. Two phases, each from `CP-17C-CORPUS-160K` and back to it with the
resting 30/60 settings staged, the guest saved and the lease released: **P1** (the premise, before anything
was built) with this branch's tools at `6c528af`, and **P2** (the builds, the D101 measurements, a live run
and a checkpoint) with its server, tools and suite at `0e018bf`. Every Outlook start NOT elevated, every step
that attaches to Outlook in a `RunLevel Limited` job, every close `Testbed/host/Restart-Guest.ps1`, every
mailbox write through the corpus tool or the hub rebuild. Raw logs: `.work\g1-d62\` of the agent worktree
`a24876cb`.

| Step | What ran | Verdict |
| --- | --- | --- |
| P1 premise | `corpus-probe --population hub --store tier@vm.invalid --all-kinds --undated-index-wait 180` (throwaway items, deleted by the two-key rule) | the appointment and the task took `2026-08-03T09:32:50Z` and read it back; the index dated both **at that instant, 9 s after their save** (`DATED AS WRITTEN`) while their `System.Message.DateSent` and `System.DateModified` stayed at the save itself; the contact `<null>`; exit 0 |
| P1 folders | `corpus-folders` of Corpus A, read-only | defaults: Deleted Items only (Inbox, Sent Items, Outbox, Drafts, Junk Email, Calendar, Contacts, Tasks ABSENT); tree: Deleted Items 19,292, `OutlookAI-Corpus-Folder-6` 88,037, `-5` 39,709, `-Junk` 12,962 - the 160,000 |
| P1 sweep cost | `Measure-SweepCost.ps1 -Store 'Corpus A' -WindowDays 30` as it stood | the new resolver right (`PR_VALID_FOLDER_MASK 0xC9`: Inbox and Sent Items ABSENT, Junk Email ABSENT through the absent Inbox, Deleted Items its own); then `Method 'System.Object[].Count' not found` - the unroll, fixed in `a857704` |
| P2 bystander | the contacts population torn down by its own manifest (`--undated-contacts`), that manifest kept as `bystander-history\bystander-indexed.20261003T063842Z.contacts.jsonl`, the all-kinds one built at the SAME anchor, `2026-10-03T06:38:42Z` | teardown 684 deleted, 5 folders removed, 0 left; probe: both planned `2024-10-02T06:38:42Z` and read back; 342 built in 11 s; census 342 of 342 (Calendar 14, Contacts 14, Tasks 14); read-back 28 appointments and tasks at their planned instant, the 14 contacts store-dated; **index 342 of 342 at its first ask - 14 with no received date, 28 at their planned instant**, newest `2026-10-02T06:38:42Z` as before; manifest `259A0F27...` (350 lines) |
| P2 D101 | on the same Outlook: `corpus-folders`, `Measure-SweepCost.ps1` (30 days, 2 passes, without and with `-OpenItems`), the product's sweep of Corpus A through the shipped server (read-only tools, a scratch stdio driver), `Invoke-GuestMeasure.ps1 -Store 'Corpus A' -ScanFolder OutlookAI-Corpus-Folder-6` | below |
| P2 hub | `Reset-HubPopulation.ps1 -Execute` (this branch's), `RunLevel Limited` | "torn down as undated contacts only (--undated-contacts); built with appointments, contacts and tasks (--all-kinds)", the move noted; teardown 136 deleted, 2 folders; probe planned `2026-08-03T17:26:22Z`, read back; 68 built; **index 68 of 68 after 16 s: 4 with no received date, 8 at their planned instant; newest planned = newest indexed = `2026-10-03T17:25:22Z`** - the frontier untouched; 109 min of margin |
| P2 live | the guest's filter, `OAI-INDEXED` opt-in, `RunLevel Limited`, 626 s | 123 tests: 107 passed, 16 failed - **every one of the 16 also failed in another agent's run of master's suite on the contacts populations the same afternoon**, with the same message (12 x `Store 'Corpus A' not found among 3 discovered index scopes` and one `store scope not discoverable in the index` - store discovery, being fixed elsewhere; the hub's top-100 search at exactly 100, as with the twelve contacts; the apostrophe-in-a-folder-name search; the Outlook-exit COM case); that run's 17th, a timing race, passed here; 0 Outlook crashes |
| P2 checkpoint | `Restart-Guest.ps1`, then `CP-18C-ALL-KINDS` | below |

**The sized contest and the collation, from that live run** - the order-key tests measure the hub and the
bystander before they reach Corpus A, where store discovery stops them:

```
NullCollation  store=tier@vm.invalid rows=103 undated=28 firstUndated=75 lastDated=74 verdict=NULLS LAST (guard rarely fires)
               store=bystander@vm.invalid rows=380 undated=28 firstUndated=352 lastDated=351 verdict=NULLS LAST (guard rarely fires)
Floor          store=tier@vm.invalid rows=75 undated=0 undatedWithoutTheFloor=28
               store=bystander@vm.invalid rows=352 undated=0 undatedWithoutTheFloor=28
WidenedSearch  store=tier@vm.invalid contest: 28 undated row(s) against Top 17, fetched as TOP 44 (27 row(s) of over-fetch room) - if the undated rows sort ahead, the statement can hold at most 16 dated row(s), fewer than 17
               store=tier@vm.invalid top=17 sqlTop=44 statementRows=44 statementDated=39 predictedDated=39 widened=17 (dated 17) mailKindOnly=17 (dated 17) ... guard=not needed (...)
               store=bystander@vm.invalid contest: 28 undated row(s) against Top 17, fetched as TOP 44 ...
               store=bystander@vm.invalid top=17 sqlTop=44 statementRows=44 statementDated=44 predictedDated=44 widened=17 (dated 17) mailKindOnly=17 (dated 17) ... guard=not needed (...)
```

So the provider sorts a row with no `System.Message.DateReceived` LAST under `DESC` inside a mapi `SCOPE`,
and accepts the `1601` floor literal, which excludes exactly those rows - the two facts the order-key guard
was written not to need (`Docs/magic-numbers.md`). The undated rows are more than the contacts: 28 in each
store, the rest folder rows, which carry no received date either; the contest counts whatever the index
leaves undated. Both stores are contested (twelve rows or more), the unguarded statement held exactly what
the wider sample predicted, and the guard was not load-bearing - on this provider it never is, and the test
now says so every run. The frontier test and the completeness oracle passed on the all-kinds hub (the oracle
56 of 56 for each of three terms).

**D101, measured on Corpus A** (Outlook on the account-less profile, which mounts it):

* **The product's sweep is right about it.** A search scoped to Corpus A sweeps
  `"folders":["Corpus A/Deleted Items"]`, `foldersSwept 1`, `foldersAbsent 3`, `foldersSkipped 0`, from the
  store's frontier `2026-10-02T23:59:16Z` less its 10-minute margin; the index tier returns hits from all four
  folders, stand-ins included, and a search scoped to the stand-in `OutlookAI-Corpus-Folder-6` sweeps that
  folder. No default folder is created, none is mistaken for another, and the absent three are not reported as
  gaps - a data file with no Inbox has nothing arriving in one.
* **`Measure-SweepCost.ps1`, first run** (fixed): every folder sorted; 28.5-35.9 ms a row walking the table,
  34.9-43.3 ms with `-OpenItems` - PowerShell's late-binding cost more than Outlook's, so only the difference,
  about +7 ms a row, speaks for opening an item; the stand-ins cost what Deleted Items costs.
* **`Invoke-GuestMeasure.ps1`**: its unscoped sweeps read 12 items across the five stores (every store here is
  indexed, so each is swept from its frontier) - they measure an unindexed corpus only, which its header now
  says; its scan of the stand-in took 372 ms and the whole-store 365-day scan 461 ms, neither timed out.

| Checkpoint | State |
| --- | --- |
| `CP-18C-ALL-KINDS` (parent `CP-17C-CORPUS-160K`; taken with the guest running, 2026-10-03 19:42 local) | Outlook not running (the restart closed it); default profile `OutlookAI-Tier`; the hub all-kinds at anchor `2026-10-03T17:26:22Z` (`D08B4C19...`), the bystander all-kinds at `06:38:42Z` (`259A0F27...`, its contacts manifest in `bystander-history\`), the identity store as at `CP-16C` (`EB2F0B06...`), Corpus A as at `CP-17C` (`AB395B81...`); the server, tools and suite at `0e018bf`; the live-test settings rendered from this branch, windows 30 and 60 (`4702979C...`); E: 227.7 GB free |

**Guest one RESTS on `CP-17C-CORPUS-160K`** with the 30/60 settings (`FDDA110B...`), as before: the contacts
populations stay its resting state until this branch is merged. **After the merge**, restore
`CP-18C-ALL-KINDS` instead - its hub is rebuilt by every run anyway (the rebuild keeps all three kinds from
then on), and its bystander is the one that needs the all-kinds build. The same date applies: **before
2026-11-01 23:59 UTC** for Corpus A's 30-day window (Q108). *Once Q130 (a) is merged, runs restore
`CP-20C-FROZEN-CLOCK` instead (section 4.5): CP-18C-ALL-KINDS frozen at its own instant, whose clock never
reaches that date.*

### 4.2f The live tier's first runs - `OutlookAI-Indexed`, 2026-10-03

**Why this section exists.** The indexed guest's first full live runs: the index tier
(`Requires=SearchIndex`) on a guest for the first time, with the 160,000-item Corpus A mounted, and the
first runs through `Testbed/host/Invoke-LiveTierOnGuest.ps1` (`Testbed/README.md` section 4c), written
after run 1 had been driven by hand. Every run: the start checkpoint restored, the commit's payloads,
settings (30/60) and guest scripts staged, TEST-READY, `Restart-Guest.ps1 -Execute`, step 9a at
`-RunLevel Limited` (and 9a-ii while it existed), the suite at `-RunLevel Limited` with
`OUTLOOKAI_LIVE_OPT_IN = 'OAI-INDEXED'` and
`--filter "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange"`, then the resting
checkpoint restored, the settings staged on it, the guest saved, the lease released. Every fix went
in on the host first with T1 tests and was proven on the build VM (`Invoke-TestsOnBuildVm.ps1`: 3,613
of 3,613 at `d6b8f89`, 3,619 at `19e6021`, 3,558 at the merge `dc1b5c5`, every self-test each time).
Results: `.work\g1-live-green\` (run 1) and `.work\guest-live-runs\<run>\` (the rest) in worktree
`agent-a634a99d582147265`.

| Run | Revision | From | Total | Passed | Failed | Suite | Outlook |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | `d4e31fe` (master), by hand | `CP-17C` | 123 | 106 | 17 | 10.9 min | no crash |
| 2 | `e5bec94` (fixes 1-4 below) | `CP-17C` | 123 | 122 | 1 | 11.5 min | no crash |
| diag | `2ff9e4f`, `-FilterSuffix` two tests, `-SkipHubReset` | `CP-17C` | 2 | 1 | 1 | - | no crash |
| 3 | `19e6021` (fix 5) | `CP-17C` | **123** | **123** | 0 | 11.4 min | no crash |
| 4 | `dc1b5c5` - master `af1fd3f` merged | `CP-18C-ALL-KINDS` | 122 | 119 | 3 | 15.8 min | **crashed** (`wwlib.dll`) |
| 5 | `dc1b5c5` again | `CP-18C-ALL-KINDS` | **122** | **122** | 0 | 11.1 min | no crash - **GREEN**, `CP-19C-LIVE-GREEN` |

**Run 1's seventeen, and what answered them** (each decided in `QUESTIONS.md`'s decision log, "the
indexed guest's first live runs"):

1. **Thirteen tests: "Store 'Corpus A' not found among 3 discovered index scopes".** The 2000-row
   discovery sample held `Outlook Data File(1941)` - the account-less profile's 20,000-item store -
   `tier@vm.invalid(59)` and `bystander@vm.invalid(1)`, and none of Corpus A's 160,006 rows; no
   address names Corpus A. `T2/LiveIndexScopes` adds the store-root listing as the third step. Run 2:
   all thirteen ran, and passed.
2. **`LiveFreshModeTests`**: the index served the arrival with its frontier at `16:02:34.0000000Z` for a
   send at `16:02:34.3804579Z` - whole seconds. Compared at that precision; run 2:
   `sweptLive=False indexCaughtUp=True inboxFindMs=3343`.
3. **`LiveDisconnectRecoveryTests`**: closing the parked window quit Outlook (D49's Office 2024
   finding), and the pin release then reconnected into it (`0x800706BA`). Released only on a live
   session now; run 2: "no live session left to release a pin on - Outlook quit as its last visible
   window closed (gone signal: quit event)", passed.
4. **`LiveFolderScopeTests.ApostropheInAFolderName`** asserted the zero-row guard's pre-G5 contract.
   Held to G5; run 2: "zero-row guard=spoke (the folder is new to the index)", passed.
5. **`Phase7LiveMcpToolShapeTests.Search_TopOne_OnHubStore`: 100 hits for a hub of 68 - a product
   defect.** Counted by store, tier, folder and class (run 2, then the diag run): 67 mail items, 12
   contact cards and **21 folder rows** (`kind:folder` - every folder of the hub, its root, Calendar,
   Quick Step Settings, the emptied subfolders in Deleted Items). `IndexRowFilter.IsFolderRow` drops a
   row of kind `folder` with no item segment (`19e6021`, T1 `IndexRowFilterTests`). Run 3: passed, and
   the whole tier with it - **123 of 123, the indexed guest green on `CP-17C`.** Its checkpoint
   (`CP-18C-LIVE-GREEN`) was deleted again: the coordinator moved guest one to `CP-18C-ALL-KINDS`
   while it ran.

**What the brief asked to confirm.**

* **Q74 C3's PST half** - `LiveDecodeVerifyTests.ShortDecodedId_OpensAsTheItemItself_OnAPstStore`
  passed in runs 1 to 3: "short-id open on a PST: result=opened ... decoded=00000000CD4732829AAD7048B384708E60CA19F2241D2000
  opened=00000000CD4732829AAD7048B384708E60CA19F2241D2000" - the 24-byte id decoded from the index URL
  opens as the item itself on a PST. Its `TODO.md` item is gone.
* **The order-key tests with undated contacts (D62, D74)** - all three passed in runs 2 and 3. NULL
  collation measured under `DateReceived DESC`: the hub 103 rows, 36 undated (its 12 contacts and the
  folder rows above), **NULLS LAST**; the bystander 373 rows, 49 undated, NULLS LAST; Corpus A 500
  rows, none undated.
* **Corpus A's census inside the tripwire's time** - counted, never identified (4 folders measured, 0
  identified) and never swept (a declared bystander): baseline 543 / 643 / 1,114 ms (runs 1, 2, 4),
  post-run 597 / 509 ms; the whole post-run census 1,331 to 3,024 ms, every one `0 failure(s)`.

**Safety, every run:** the artifact sweep ran and every store ended at `taggedArtifacts=0` (Corpus A
and the bystander counted, not swept); the post-run census reported 0 failures; one PROVED NOTHING a
run - the retry-guidance check, which has no transient Outlook state to report on an idle guest.

**The three failures another agent saw here (coordinator heads-up).** A folder-path run staged its
suite onto `CP-17C`'s running Outlook - no graceful restart, no step 9a - and failed
`LiveMoveArchiveTests.MoveChain` ("NoDesignatedArchiveFolder", then five later hub checks),
`LiveSweepScopeTests.ControlledCorpus` (its self-sent mail never arrived) and the apostrophe test
(item 4). Here MoveChain passed in all five runs and in the diag run, which restarted the guest but
skipped 9a; ControlledCorpus in all five. Not reproduced under the procedure, whose graceful restart
the runner makes unconditional; MoveChain's is kept open in `TODO.md`.

**Run 4 - Outlook crashed, in Word.** On the new base and the merged master, `OUTLOOK.EXE` died at
20:56:49 local inside `LiveDraftOptionsTests.NewDraft_Hub_SignatureOverride_BodyAboveTheSignature_OutsideTheSignatureBookmark`
- `Faulting module name: wwlib.dll, version: 16.0.17932.20996 ... Exception code: 0xc0000005 Fault
offset: 0x000000000007ba1a` - and the next two tests of the class failed on `0x800706BA` and
`CO_E_SERVER_EXEC_FAILURE` while it went. The guest-two crash (section 4.1e) faulted in `ntdll.dll` and
`OLMAPI32.DLL` and stopped when every COM child object was released; the signature path this test
drives captures and releases every Word object it touches. Run 5, the same commit from the same
checkpoint, did not crash. One crash in five full runs here, none in the guest-two runs since the
release fix; no dump was written (the guests do not keep any), which is the first thing to change
(`TODO.md`). *They keep full dumps since 2026-10-04 - section 4.6.*

**Green, and where guest one rests.** Run 5 is the indexed guest green on the merged master: 122 of 122,
artifacts 0, census 0 failures, no crash. Its checkpoint:

| Checkpoint | State |
| --- | --- |
| `CP-19C-LIVE-GREEN` (parent `CP-18C-ALL-KINDS`; taken with the guest running, 2026-10-03 21:35 local, by `Invoke-LiveTierOnGuest.ps1 -GreenCheckpoint`) | The green run's end state: `dc1b5c5` staged (server, tools, suite, 30/60 settings), the hub as the run left it after its rebuild and sweep, Outlook as the suite left it |

**Guest one RESTS on `CP-18C-ALL-KINDS`** with the 30/60 settings staged (the runner leaves it so after
every run), not on the green checkpoint: `CP-18C-ALL-KINDS` is the base every agent's phase restores
(the coordinator moved guest one there during this work), and the green state is a run's end, with a
hub the next run tears down and rebuilds anyway. `CP-19C-LIVE-GREEN` is the evidence, kept beside it.
*Since Q130 (a) (section 4.5) the runner's default for guest one is `CP-20C-FROZEN-CLOCK` - a second
child of `CP-18C-ALL-KINDS`, taken three minutes after `CP-19C-LIVE-GREEN` and holding CP-18C's own state
frozen at its own instant; it was made as `CP-19C-FROZEN-CLOCK` and renamed the same evening so no two
checkpoints share a number.*

### 4.3 The build VM - `OutlookAI-Build`, 2026-10-03 (Q94, Q102)

**Why this section exists.** The third machine, and not a live-tier guest: it runs the non-live
suite and the script self-tests for `Testbed/host/Invoke-TestsOnBuildVm.ps1`, so that nothing runs
on the maintainer's workstation but the Exchange-only read-only live tests (Q94; `AGENTS.md`) - and
since Q116 (a), the same day, not even those (section 4.4).
No Office, no mailbox, no sink, no network. `Testbed/README.md` section 1c is the procedure and how
to use it; this is the record of building it, every step from the committed scripts and the media
`Testbed/MEDIA.md` names. Raw logs: `.work\q102-build-vm\` in the main checkout.

| Step | What ran | Verdict | Checkpoint |
| --- | --- | --- | --- |
| 1. Answer volume | `New-AnswerFile.ps1 -VMName OutlookAI-Build -ComputerName OAI-BUILD` | 675,840 bytes, built with oscdimg | - |
| 2. Create and boot | `New-TestbedVm.ps1 -Name OutlookAI-Build ... -ProcessorCount 4 -MemoryStartupBytes 6GB -Execute -Start` | Generation 2, Secure Boot, vTPM, 128 GB dynamic VHDX, the adapter disconnected; 27 keystrokes typed through the boot prompt | - |
| 3. Unattended install | nobody | the first-logon log `DONE. All 14 step(s) succeeded` 10 minutes after the start: en-NL then nl-NL, GeoId 176, system locale en-US, formats nl-NL, W. Europe Standard Time, Windows 11 Pro 10.0.26200 | - |
| 4. Finish the install | `New-TestbedVm.ps1 -CompleteInstall -Execute`, three times - its first real run | Attempt 1 read the DONE line, ejected both discs, then refused: its read-back went through the VM object fetched before the eject and still listed both ISOs (by name, a minute later, both drives were empty). Attempt 2, after that fix, took the checkpoint and refused again: `Get-VMSnapshot` straight after `Checkpoint-VM` listed no checkpoint of the name. Attempt 3 found it, without a disc, and deleted the answer ISO. Each refusal left everything a restore needs; both read-backs now go by name and poll | `CP-01-WIN-CLEAN` |
| 5. Payload | `Publish-LiveTierPayload.ps1 -Ref e4b00fa -ExpectedSha512 <MEDIA.md's>`; `Copy-ToGuest.ps1` four times | the SDK hash MATCHES; `Source.zip` 3.2 MB, 54 packages, 75.7 MB; the feed restores all five projects with every other source cleared; 48 s to copy in | - |
| 6. SDK | `Install-DotnetSdk.ps1 -ExpectedSha512 <hash> -Execute` over PowerShell Direct | first: "running scripts is disabled on this system" - a fresh guest's execution policy is Restricted on every scope; with `Set-ExecutionPolicy -Scope Process Bypass -Force` first, the installer exited 0 in 84 s and `TEST-READY`, 3,137 tests discovered, 17 run | - |
| 7. Memory | a graceful `shutdown.exe /s /t 0` inside, `Set-VMMemory -DynamicMemoryEnabled $false -StartupBytes 6GB`, `Start-VM` | `New-VM` had made it DYNAMIC - 512 MB to 1 TB - and the spec file recorded only the 6 GB; static since, and `New-TestbedVm.ps1 -StaticMemory` does it at creation now (proved on a throwaway VM, deleted) | - |
| 8. Restart and verify | `Restart-Guest.ps1 -VMName OutlookAI-Build -Execute`; `Install-DotnetSdk.ps1 -Verify` from a new session | the restart in 26 s, no Outlook to quit - the script fits a VM without Office unchanged; `TEST-READY` again | - |
| 9. Base checkpoint | the VM's CPU at 0 % for three minutes, no PowerShell Direct session open; `Checkpoint-VM` with the VM running | 4.6 s; the saved memory is 1,578 MB on disk | `CP-02-SDK-TEST-READY` |
| 10. First proof | `Invoke-TestsOnBuildVm.ps1 -Ref e4b00fa` | **PASS: 3,005 total, 3,005 passed, 0 failed, 0 skipped**; 18 of 18 self-tests; 4 min 05 s | - |
| 11. Timed second run | the same, again | **PASS, the same counts**, 3 min 56 s: 8 s restore and resume, 1 s connect, 1 s stage; in the VM 2 s expand, 3 s restore, 23 s build, 2 min 19 s test, 39 s self-tests; 1 s fetch, 12 s restore and save | - |
| 12. Master | `Invoke-TestsOnBuildVm.ps1 origin/master` at `3e7b861` (Q74, Q86, Q93 and Q98 merged) | **PASS: 3,346 total, 3,346 passed, 0 failed, 0 skipped** - the count measured on the workstation the same night - and 18 of 18 self-tests, 3 min 50 s | - |
| 13. A failing test | a scratch commit adding one test that fails, `-Filter` on it and one real class | **FAIL, exit 1**: 19 total, 18 passed, 1 failed, the failing test and the first line of its message in the summary | - |
| 14. A compile error | a scratch commit that does not compile | **BUILD, exit 2**, in 1 min 02 s, the compiler's error lines in `vm\build.out.txt` | - |
| 15. A package the feed lacks | a scratch commit adding `Humanizer.Core` 2.14.1 to the test project | the guest named it (`PACKAGES-MISSING`), the host staged that commit's closure in 45 s (55 packages; the feed check green), added the one new package to the VM's feed, and the second attempt **PASSED** - 3 min 12 s in all | - |
| 16. Two callers at once | the branch's own HEAD, and 9 s later the failing-test commit | the second printed who held the VM every minute, ran 3 min 50 s later, and finished as above; the first PASSED - 3,089 tests and 20 of 20 self-tests, the runner's and the guest script's own among them | - |
| 17. A lease taken by hand | `Set-TestbedLease.ps1 -VMName OutlookAI-Build`, then a run under Windows PowerShell 5.1 with `-QueueTimeoutMinutes 1` | held back for the minute, then **INFRA, exit 3**, the VM untouched (still saved) | - |
| 18. A caller killed part-way | a run under Windows PowerShell 5.1, its process stopped mid-build | the lock went with the process. The VM stayed RUNNING - nothing on this host saves it, the idle-saver task not being registered here - until the next run, three hours later, restored the checkpoint over it (Running to Running), passed, and saved it | - |
| 19. The same, with a janitor | every run now starts one (the runner's banner, A CALLER THAT DIES); a run under Windows PowerShell 5.1 stopped mid-build again | the janitor - a child of `WmiPrvSE`, not of the caller - saw the run end 9 s after the kill with no end line in `runs.log`, took the lock, restored the base checkpoint over the running VM and saved it, released the dead run's lease, and logged each step: the VM saved, holding no RAM, **22 s after the kill**. After a run that finished, its janitor exited without a word | - |
| 20. This work, merged with master | the branch's HEAD after merging `7933c5c` | **PASS: 3,346 total, 3,346 passed, 0 failed, 0 skipped**, 20 of 20 self-tests, 3 min 44 s; the VM saved, its janitor gone, no lease left | - |
| 21. A finished run called "not tested" | Q96's run `20261003-091919-d7e58af90183`: the mutant commit `d7e58af`, three test classes, `-SkipSelfTests` | **INFRA, exit 3** - with its suite finished, 238 total, 237 passed, 1 failed, and its `run.json` and TRX file back. The host's read of `run.log` had failed against the guest script's own append: Windows PowerShell 5.1's `Add-Content` opens a file so that nobody else may read it, and fails itself beside a reader (measured the same morning). A caller reading the exit code took a failing run for an untested one | - |
| 22. The same, on master's runner (`8103f58`) | the same commit and arguments; a second PowerShell Direct session reading `run.log` with sharing for its cue, to hold it once the build began | **INFRA, exit 3, in 27 s** - with row 21's very error, in the restore phase, before the hold had even begun: the race is that easy to lose. Nothing came back but an empty `guest.out.txt` | - |
| 23. The fix, with `run.log` held | this branch's runner at `028019a`, the same commit and arguments; `run.log` held with `FileShare.None` for 12 s from the build's first line | **FAIL, exit 1: 238 total, 237 passed, 1 failed** - the run row 21 reported as INFRA. The runner said the log was held at its first busy poll and that it read again after 4; 55 s in all. `summary.json`'s `failed` a one-element array, `skipped` an empty one (in row 21's, an object and null) | - |
| 24. A results file held, on the first fix's runner (`89f27a9`) | the same commit and arguments; a file in the guest's `results\` held with `FileShare.None` from the build's first line until `done.txt` appeared | **INFRA, exit 3**, "no run.json came back": the guest's zip failed on the held file and left a zip of the two files it had reached, `build.err.txt` and `build.out.txt`, which the host took for the whole result | - |
| 25. The same, on `028019a` | as row 24 | **FAIL, exit 1: 238 total, 237 passed, 1 failed**, 55 s - the guest removed its partial zip, and the host read `run.log`, `run.json`, `trx\suite.trx`, `restore.out.txt` and `build.out.txt` one by one, each with sharing | - |
| 26. The whole suite again, after the fix | `Invoke-TestsOnBuildVm.ps1` at `b78d91f` - master `8103f58` and these commits, which touch no test | **PASS: 3,491 total, 3,491 passed, 0 failed, 0 skipped**, 21 of 21 self-tests, 3 min 40 s; the VM saved, no lease left | - |

**No test behaves differently here than on the maintainer's workstation**, and the two that could
were checked. The runs matched master exactly: 3,005 / 0 / 0 at `e4b00fa`, 3,346 / 0 / 0 at
`3e7b861`. `T1.SweepSortWiringTests.AnAbsentTableDateFallsBackToTheItemValueCONVERTED`
failed on GitHub CI at this commit and passes here: at `e4b00fa` it still read the machine's own time
zone, CI's runners are UTC and this VM was W. Europe Standard Time like the workstation (fixed on master
since, Q95 `3cd62c0`, by giving the test a zone of its own; the VM is UTC since Q126 (a), section 4.3a). And about 200 tests take an "Outlook is
not running" branch through `ComGateway.IsOutlookRunning` and the installer mutex - here as on CI,
which has no Office either; on the workstation they may take the other. Both branches pass (Q94's
research, decision D3 of `Docs/overnight-review-2026-10-03.md`). The suite ran in session 0, over
PowerShell Direct: nothing in it needs a desktop. Memory, sampled every 5 s by the guest script: the
lowest free 2,998 MB of 6,144, the highest committed 2,847 MB.

**What decides a run's verdict, since rows 21 to 25.** The guest's `run.json` and the suite's TRX
file, wherever they came back - in the guest's zip or, when it left none, read one at a time with
sharing. An error on the host side after that is a note beside the verdict, never the verdict; a
log or a file something else holds is waited out, not taken for a failure; and a TRX file from a run
that stopped part-way counts for the tests it shows failing, never for a pass. INFRA, exit 3, is left
for a run from which neither came back - one that was not tested.

**Three Hyper-V behaviours the runner rests on, measured here.** A checkpoint applied to a RUNNING VM
resumes it at once from the checkpoint's state (Running to Running, 9 s); so the runner, ending a
run, restores the base and then saves it - restoring alone would leave it running and holding its
RAM. A `VirtualMachine` object keeps the state it was read with and has no `Refresh()`, so every
wait re-reads the VM by name. And a running checkpoint's memory is stored sparse: 1.6 GB for 6 GB.

### 4.3a The build VM in UTC - `OutlookAI-Build`, 2026-10-03 (Q126 (a))

**Why this section exists.** Decision D4 (`Docs/overnight-review-2026-10-03.md`) built the build VM
in W. Europe Standard Time, like the workstation and the two Outlook guests, while GitHub CI's UTC
runners covered the other zone - which is how Q95's bug was found, a test that only passed where the
local zone was not UTC. CI was removed the same day, and with it every run outside W. Europe, so the
maintainer moved the build VM to UTC (Q126 (a)): the non-live suite now runs in a zone other than
the workstation's. `Testbed/README.md` section 1c, row B9, is the procedure; this is the record.
Every step under one lease, with nothing else on the VM.

| Step | What ran | Verdict | Checkpoint |
| --- | --- | --- | --- |
| 1. Lease | read first: no live lease on `OutlookAI-Build`, and the runner's lock opened (no run going); then `Set-TestbedLease.ps1 -VMName OutlookAI-Build -Minutes 60` at 22:03:24Z | read again 20 s later: still this lease, the lock still free - so no run had passed its lease check in between, and every later run queues | - |
| 2. Restore | `Restore-VMSnapshot` `CP-02-SDK-TEST-READY` (Saved to Saved), `Start-VM` | heartbeat OK 11 s after the restore. Read over PowerShell Direct: `W. Europe Standard Time`, daylight saving on, `tzutil /g` and the registry's `TimeZoneKeyName` the same; the automatic time zone service `tzautoupdate` Stopped and Disabled, so nothing on the guest sets the zone back | - |
| 3. The zone | `Set-TimeZone -Id UTC`, 22:04:50Z | read from a NEW session, a new process with nothing cached: `UTC`, "(UTC) Co-ordinated Universal Time", offset 0, no daylight saving; `tzutil /g` `UTC`; `TimeZoneKeyName` `UTC`, `Bias` and `ActiveTimeBias` 0; `TimeZoneInfo.Local` `UTC`; local time minus UTC, 0 min | - |
| 4. Restart | `Restart-Guest.ps1 -VMName OutlookAI-Build -Execute` - so that no process keeps the old zone cached, as a .NET process does | 26 s, no Outlook to quit; booted 22:05:25Z; `UTC` again from a new session | - |
| 5. Verify | `Install-DotnetSdk.ps1 -Verify` over PowerShell Direct, as B6 | `TEST-READY`: SDK 10.0.401, the probe built and ran, 3,137 tests discovered, 17 executed and passed; its log stamped in UTC (22:06:21 to 22:06:48) | - |
| 6. Base checkpoint | the VM's CPU sampled every 10 s until 18 samples in a row read 0 % (211 s: one read 3 %), no PowerShell Direct session open; `Checkpoint-VM` with the VM running | 3.1 s, 22:10:40Z, a child of `CP-02-SDK-TEST-READY`; its saved memory 1,620 MB on disk | `CP-03-SDK-TEST-READY-UTC` |
| 7. Proving restore | `Restore-VMSnapshot` `CP-03-SDK-TEST-READY-UTC` over the running VM; then the runner's way of resting - restore it again, `Save-VM` | running again in 6.5 s, `UTC`, its clock 0.1 s from the host's; then saved. The lease released at 22:11:44Z | - |
| 8. First run in UTC | `Invoke-TestsOnBuildVm.ps1` at `e5bbb4b` (the runner and `testbed.json` pointed at the new base; no test changed), run `20261004-001535-e5bbb4bd0bc3` | **PASS, exit 0: 3,818 total, 3,818 passed, 0 failed, 0 skipped**, 27 of 27 self-tests - exactly the W. Europe run of `289030e` (no test differs between the two commits), whose test phase took 3 min 23 s to this one's 4 min 35 s; the W. Europe runs of the same evening took 2 min 12 s to 4 min 53 s, and the extra seconds sit mostly in the classes that start processes (`AuditLogStressTests`, `AuditLogTests`, the T3 server tests), not in any one test. The runner's line: `VM: OAI-BUILD as vmadmin, time zone UTC, clock within 0 s`; the summary's `machine` line: `UTC` | - |

**No test failed in UTC**, so Q126 found no time-zone bug on its first run: the suite at `e5bbb4b`
passes in both zones. What UTC cannot do is stand in for W. Europe. Tests that use the machine's own
zone as the "other" zone - converting a `Local` value and comparing it with the `Utc` one - prove
only the identity on a UTC machine, and since Q94 nothing else runs the non-live suite: some of them
were written knowing it (`T1.ComDateValueTests` pins its conversion in a zone of its own and reads
its overload's IL for the same reason), others were not. That is an open question, not a defect of
this switch - recorded with the decision (`Docs/overnight-review-2026-10-03.md`, Q126).

**What enforces the zone.** The base checkpoint holds it - Windows keeps the zone on the disk and
in the saved memory alike - and the runner refuses a guest that reads otherwise: its identity check
asks for `(Get-TimeZone).Id` beside the computer and the account, and `Get-GuestRefusal` turns
anything but `testbed.json`'s `buildVm.timeZone` into a refusal before a byte is copied in, INFRA,
exit 3. `-SelfTest` holds that value equal to the runner's `$BuildVmTimeZone` and checks the
refusal, `GMT Standard Time` included - the zone that reads like UTC and keeps British summer time.
`CP-02-SDK-TEST-READY` keeps W. Europe and stays: a branch whose runner still names it runs from it
as before, until it merges this.

### 4.4 The Exchange VM - `OutlookAI-Exchange`, 2026-10-03 (Q108 to Q111, Q113 (b), Q116 (a))

**Why this section exists.** The fourth machine: ONE real Microsoft 365 mailbox, `telefonie@xxlnet.nl`,
cached and indexed, so that the live tests needing an Exchange profile leave the maintainer's
workstation, whose Outlook holds far more critical mailboxes (Q108). `Testbed/README.md` section 1d
is the procedure, 4e the run; this is the record of the build-out, every step from the committed
scripts. It has internet (Q111) and no PST, no sink, no population. Raw logs: `.work\exchange-vm\`
of the worktree that built it.

**The build, 2026-10-03, in the order it ran** (times local, UTC+2):

1. 19:21 - `New-AnswerFile.ps1`, then `New-TestbedVm.ps1 -Execute -Start` with its own disk folder,
   4 vCPU, 8 GB static, no switch. 19:27:39 first logon `DONE`; `-CompleteInstall` ejected both discs,
   took `CP-01-WIN-CLEAN` and deleted the answer ISO.
2. The Default Switch connected; `outlook.office365.com:443` reachable. Windows Update policy
   `NoAutoRebootWithLoggedOnUsers` 1 with `AUOptions` 4.
3. Office from `.work/office-odt/Testbed.xml`, unchanged, online: 2.3 min, 16.0.17932.21000. The
   configuration file deleted from the guest; first-run suppression `-Execute` and `-Verify`, 13 OK.
   `CP-02-OFFICE-INSTALLED`. The licence read after Outlook's first start: `VOLUME_KMSCLIENT`,
   out-of-box grace, 30 days - no KMS host reachable, nothing activated.
4. Outlook's first start, NOT elevated, no profile: its "Email Account Setup" dialog; the address
   typed, the Microsoft sign-in's password page (a WebView in an `ApplicationFrameWindow`, UI Automation
   ids `i0118` and `idSIButton9`), the password typed from the host - no MFA page that time -, then
   "Sign in to all apps and websites on this device?" answered "No, this app only", then Outlook's
   "Account successfully added", its Outlook Mobile box cleared and Done, the last three through MSAA:
   Outlook's NetUI shows the managed UI Automation client a pane and nothing in it. First by hand with
   scratch helpers, then as the committed `host/Invoke-ExchangeSignIn.ps1` from `CP-02` again: 2 min
   51 s, `SIGNED-IN` - profile `Outlook`, account type Exchange, `ExchangeConnectionMode` 700, the
   store `IsCachedExchange` and `IsInstantSearchEnabled`. `dsregcmd`: AzureAdJoined NO, WorkplaceJoined NO.
5. `Set-OutlookIndexingDisabled.ps1 -Verify`: `INDEXED`, 183 rows under `telefonie@xxlnet.nl($65e0d53e)`.
   The mailbox is small: 25 mail folders, 5 mail items (Inbox 1, Sent Items 2, Deleted Items 2).
   `CP-03-EXCHANGE-SIGNED-IN`.
6. The suite staged from the branch's commit (`Install-DotnetSdk.ps1 -Execute`: `TEST-READY`), the
   settings rendered (`machineProfile` `ExchangeGuest`).
7. The read-only runs below; `CP-04-SUITE-READONLY-RUN`.

**The read-only runs.** Every one through `guest/Register-InteractiveTask.ps1 -RunLevel Limited`, opted
in for `OAI-EXCHANGE`, against the one Outlook, unelevated. The count tripwire, on every run:
`watch soundness: 0 declared bystander(s), 1 store(s) this census can fail on` - the hub, censused
item by item because the machine is read-only - and afterwards `0 failure(s), 0 note(s)`.

| Run | Filter | Result |
| --- | --- | --- |
| A | the Exchange VM's filter, branch at `dbc8b44` | 1 of 1: the short decoded id rejected, `0x80040107` |
| B | every `Writes=Nothing` test outside `DelegateStore`, to learn which make sense on Exchange | 49 run, 29 passed, 20 failed - every failure but one a test that needs what this VM does not have: the hub population and its attachments, its probe term and subject-only probe (null), a PST, three stores, 25 mail hits, mail in the last 30 days or over 100 KB, or the add-in, which this VM does not have (`LiveHealthTests` reads its tuning state without declaring `AddInRegistry`). The one that is about Exchange: `RoundTrip_SearchThenRead_TenHitsAcrossStores` read 9 hits and then met one whose item is no longer in its folder - an index row for an item gone since (the product's message says so), on a live mailbox where Outlook prunes its own sync logs |
| C | the Exchange VM's filter, with the two tests below added | 3 of 4: `T2/LiveExchangeStoreHashTests` and the short id pass; `T2/LiveExchangeHubArtifactTests` FAILS on one item tagged `[OutlookAI-McpTest]` in Sent Items |
| Q99 | the same filter, on a LOCAL-ONLY merge of `q99-name-encoding-followup` (`a6f4371`) with this branch, which the one-mailbox VM needs to run at all | 3 of 4: `T2/LiveExchangeFolderPathTests` PASSES - 40 folders walked, 16 nested paths at depth 2 each resolved to itself, and a missing child answered NotFound with one place to build: Exchange answers a missing folder name with MAPI_E_NOT_FOUND. The gate for that branch holds. The artifact count fails as in C |

**The Object Model Guard, met at 21:01.** After `CP-02`'s proof run the VM was restored to `CP-04`,
and the next run from the branch's head refused twice: the baseline census timed out at 300 s with two
folders measured. A COM read hung as well. Outlook's windows, read by title and class only: six
"Microsoft Outlook" dialogs stacked on its disabled main window, each the guard's "A program is
trying to access email address information stored in Outlook" - Defender's signatures were the
image's, 381 days old, and Security Center reported them out of date (`productState` 0x061110).
`Update-MpSignature` (26 s) brought it to 0x061100; the six prompts were answered Deny through their
own `WM_COMMAND`, the hung calls failed, and the same run then passed as in C - 2 of 3, the leftover
failing, the tripwire clean. `host/Invoke-ExchangeSignIn.ps1 -Mode Preflight` now does that before
every run.

**Q113 (b), measured:** `outlook_health`'s row for the cached Exchange store reads `matchedBy=storeHash
matchedInput=profileMappingSignature inLocalIndex=True`, index root `telefonie@xxlnet.nl($65e0d53e)`,
`storesNotInProfile` 0. Outlook hashes the profile's `PR_MAPPING_SIGNATURE`, as Microsoft documents,
so for a cached Exchange store the name fallback did not decide anything. One store, one profile, one
measurement: a delegate store's row and a second profile are untested here.

**The leftover in Sent Items.** One item whose subject carries `OutlookAI-McpTest`, found by the
read-only count; nothing this VM ran created it - the census identifies the same items before and
after every run. It is test data, so under the maintainer's rule (Q129 (a)) it must go; on this
read-only machine nothing removes it, and nothing will by hand.

**Phase 2 - PROPOSED 2026-10-03, NOT APPROVED.** What writing in this real mailbox would rest on, each
part in code. Nothing below is built; the maintainer decides first.

1. **The rule (Q129 (a))**: tests create, change and delete only their own tagged items, and remove
   every one of them; no untagged item is touched, lost or buried.
2. **A recipient allowlist in code (Q110)**: every outgoing address - To, Cc, Bcc, a reply's and a
   forward's - must be the hub's own address, or the run refuses before the item is saved: a new
   `RecipientAllowlist` asked by `LiveOutlookTestMailer` before every save and send, and by the stdio
   client before every write-capable tool call whose arguments name a recipient. Reply-all, replies
   to real mail and forwards are refused unless the source item is a tagged item of this run.
3. **Writable only as an Exchange guest with the allowlist**: `LiveWriteAccess` would let
   `ExchangeGuest` write in the hub alone, and only while that allowlist is armed; the hub stays
   censused item by item, with the run's own tagged items the only departures and arrivals the
   tripwire accepts: a departure of an UNTAGGED item fails, an arrival is noted - real mail arrives.
4. **Every item tagged twice**: the subject tag and a run marker - `[OutlookAI-McpTest]` plus the
   run's id - and every created item's EntryID recorded the moment it is saved (the allowlist the
   sweep deletes by: EntryID AND ordinal tag, both required).
5. **Self-sent mail**: a send is addressed to the hub itself; its Sent Items copy and the delivered
   Inbox copy both carry the tag and the marker and are both swept; the sweep waits for the delivery
   (the existing stable-zero wait) so no copy lands after it.
6. **Purged, not left in Deleted Items**: the sweep deletes each item, then deletes that copy again
   from Deleted Items, so no test mail lingers in a folder the owner reads. Exchange then keeps it in
   the hidden Recoverable Items folder for its retention period (14 days by default) - out of the
   owner's sight, and nothing a test touches.
7. **The sweep's folders**: Drafts, Inbox, Sent Items, Outbox, Deleted Items and the Sync Issues
   subtree (Conflicts, Local Failures, Server Failures), plus the test folders it created,
   deepest first - the existing `HubSweepFolderIdsWithArchive` set.
8. **Populations**: created per run, under the run marker, and removed at the end - never kept in
   the mailbox between runs.
9. **An aborted run**: the next run's preflight counts tagged items; with any present it refuses to
   start until the leftover sweep - by tag AND marker of a recorded run, never by a subject pattern -
   has removed them.
10. **Never**: an untagged item touched, a deletion by subject pattern, a send to anyone but the hub.

**Which writes can move to the shared test mailbox once it exists (Q110) - read from the product's
code on 2026-10-03, to be confirmed on it.** `new_draft` creates the draft with `Items.Add` in the
SENDING ACCOUNT's Drafts and pins `SendUsingAccount` from an Account object
(`OutlookComSession.TryCreateNewDraft`); a shared mailbox reached through the account is a delegate
store, not an account, so new drafts - and `update_draft` and `discard_draft` on them - stay in
telefonie's Drafts. `send` sends from the account, so its Sent Items copy stays in telefonie's Sent
Items, and under the recipient allowlist the delivered copy lands in telefonie's Inbox. `reply_draft`,
`replyall_draft` and `forward_draft` pin the account whose delivery store holds the source item and
save "into that store's Drafts" (`TryCreateDerivedDraft`): for a source item in the shared mailbox no
account has that store, `SendUsingAccount` is left to Outlook, and whether the draft then stays in the
shared mailbox's Drafts is exactly what the first run there must read. `move_mail` and `archive_mail`
move items wherever they are, so a population built in the shared mailbox can stay there.
`manage_signature` writes no mailbox at all.

### 4.5 The frozen guest clocks - both Outlook guests, 2026-10-03 (Q130 (a) and (b))

**Why this section exists.** The maintainer's answer of 2026-10-03 to Q130
(`Docs/overnight-review-2026-10-03.md`): (a) freeze the clocks of `OutlookAI-Indexed` and
`OutlookAI-Unindexed` - time synchronisation off, every run from a checkpoint whose clock stands just after
the guest's data was built, so the data never ages; (b) anchor the one corpus-dependent timing test's window
to the data; and first, a 20-minute measurement of the one step that could stop (a): installing the add-in
on a frozen guest. Not the build VM (its runner needs the host's clock within 2 s) and not `OutlookAI-Exchange`
(Microsoft 365 sign-in needs real time). Raw logs: `.work\frozen-clock\` of the agent worktree `a5fd1dc3`.

**1. The measurement: the two-phase add-in install on a frozen guest - it does not block (a).** On
`OutlookAI-Unindexed`, from `CP-08-MAIL-SINK` - the checkpoint before `CP-09-ADDIN-READY`, which section 2.3's
unfrozen run of the same day started from - with the payload built on the host from `af1fd3f` by
`Testbed/host/Publish-AddInPayload.ps1` (guard 3 `UNCHANGED` over 305 host lines; its throwaway signing
certificate `8C328F09...` valid from 2026-10-03T18:13:17Z). Every phase through `Register-InteractiveTask.ps1`,
as section 2.3 runs them:

| Step | What ran | Result |
| --- | --- | --- |
| restore | `CP-08` restored onto the saved VM, time sync turned OFF while it was saved, started | `Disable-VMIntegrationService` accepts a saved VM. The guest resumed at 2026-09-24T15:19:34.3Z - 1.8 s before `CP-08` was taken, 790,668 s (9.15 days) behind the host - and held that offset to 0.1 s over 20 s, and to the second through every step below |
| signatures | `Get-AuthenticodeSignature` in the guest | `vstor_redist.exe` `Valid` (its signer 2023-11-16 to 2024-11-14, timestamped); the product's installer `NotSigned` - a testbed build stops at the unsigned installer by design (`Publish-AddInPayload.ps1`, THE RELEASE BUILD) |
| Install | `-Phase Install -Execute`, default task (elevated) | **`INSTALLED-NEVER-RAN`, exit 2** - the VSTO runtime installed (`v4R` 10.0.60917), the trust entry and the install record written: as unfrozen (section 2.3 row 3) |
| manifest | the installed `OutlookAI.vsto`'s certificate against the guest's clock | `CN=OutlookAI Testbed`, notBefore 2026-10-03T18:13:17Z - **NOT YET VALID on the guest's clock, by 9.1 days** |
| FirstRun | `-Phase FirstRun -Execute`, `-RunLevel Limited` | **`BROKEN`, exit 1**, 250 s: COM start 3.4 s, `Connect = True`, the add-in answered `GetRestartNeeded()`, OUTLOOK.EXE NOT elevated, its registration reconcile wrote `Mcp\LastReconcileUtc` 3.6 s in, the tuning walk 4 of 13 - the product defect section 2.3 recorded that day, exactly as unfrozen (rows 6 and 8); this payload was `af1fd3f`'s, from before Q128 fixed it (section 2.3, "The defect, fixed"). No window appeared during the run (the script's own check), no VSTO or Office event in the Application log, no VSTO alert log (`VSTO_LOGALERTS` is 1) |
| control | the five Cached Mode policy values from session 0, elevated, as row 9 there | - |
| FirstRun | again, Limited | **`ADDIN-READY`, exit 0**, 8 s: `LastReconcileUtc` 2.2 s in (written as the frozen clock's 2026-09-24T15:26:12Z), 13 of 13, token NOT elevated: as unfrozen (row 10) |
| `-Verify` | session 0 | `ADDIN-READY`, exit 0 - "the add-in has run since" |

So the VSTO runtime trusted and loaded an add-in whose manifests are signed by a certificate that is not valid
on the guest's clock for another nine days - the trust the installer writes is the inclusion list, keyed on the
public key - and raised no trust prompt; both phases reached the verdicts they reach unfrozen. **The future
risk the question named - a certificate renewed later, with a notBefore after the frozen date - is this case**:
every testbed build signs with a throwaway certificate made at build time, so every add-in installed on a frozen
guest from now on is "newer than its clock", and nothing in its trust path read the dates. It does not block (a).
**Not measured:** an Authenticode-signed installer whose certificate starts after the frozen date - a release
installer signed after a renewal of the maintainer's certificate, or a future SDK or VSTO redistributable. The
testbed installs none of those on a frozen guest today; `vstor_redist.exe` is timestamped, and section 4.1f found
two months of clock lag harmless to the SDK's signatures.

**2. A restart of a frozen guest, measured on the same guest.** `Testbed/host/Restart-Guest.ps1` without
`-Refreeze` refused, exit 2, dry run and `-Execute` alike, nothing changed. With `-Refreeze` it restarted the
guest in 27 s and **the guest came back 610,668 s (7.07 days) ahead of its own time** - at the host's time less
about 50 hours, whatever offset its clock last held - even with its present written to its clock (`Set-Date`)
just before the restart, so that write was taken out again; the correction after the boot left it -0.2 s off.
The shipped version, once more: the same 610,668.6 s, corrected to -0.0 s, and the guest's offset from the host
afterwards what it had been before both restarts. So a restart after a restore does not come back to the frozen
instant - measured now, where section 4.1f inferred it - and the only way to keep a frozen guest's time across one
is to set it after the boot, which is what `-Refreeze` does, for work outside a run only.

**3. The frozen checkpoints**, both made by `Testbed/host/Set-GuestClockFrozen.ps1 -FromCheckpoint
<base> -NewCheckpoint <name> -Execute`: the VM saved, the base restored onto it (Saved), time sync turned off
while saved, started - it resumes at the base's own instant - its offset from the host held to 0.1 s over 20 s,
OUTLOOK.EXE not running, the checkpoint taken between two clock readings, a proving restore, and the VM left
saved on it. `Testbed/testbed.json`, `frozenClocks`, records both:

| Guest | Frozen checkpoint | Made from | Frozen instant | The data at that instant (read off the frozen guest) |
| --- | --- | --- | --- | --- |
| `OutlookAI-Unindexed` | `CP-14B-FROZEN-CLOCK` | `CP-13B-LIVE-GREEN` (14:49:51Z, after green run 18) | **2026-10-03T14:50:12Z** (1791039012); the proving restore came back 0.3 s before it | the hub at anchor 14:37:29Z (`39EAB0AC...`, 56 items) - its newest item 13.7 min old; the bystander 01:17:22Z (`C0A5B75C...`), the identity store 01:18:26Z (`A14EAE6E...`); no corpus declared |
| `OutlookAI-Indexed` | `CP-20C-FROZEN-CLOCK` | `CP-18C-ALL-KINDS` (17:42:43Z) | **2026-10-03T17:43:06Z** (1791049386); the proving restore came back 2.2 s before it | the hub all-kinds at anchor 17:26:22Z (`D08B4C19...`) - its newest dated item 17.7 min old; the bystander all-kinds 06:38:42Z (`259A0F27...`), the identity store 06:39:44Z (`EB2F0B06...`), Corpus A `AB395B81...`, its newest item 2026-10-02T23:59:16Z |

**Why these instants hold for every run.** A run's suite must start within 45 minutes of the frozen instant -
`Set-GuestClockFrozen.ps1 -Verify` refuses later - and `T1/FrozenGuestClockTests` proves, from the records and
with the tier's own code, that every date check holds to 90 minutes after it, which leaves the suite 45 minutes
(guest one's full suite took 16 minutes in the runner's unfrozen run of the same hour): the guest's UTC offset is
+2 h at both ends - no run can reach 2026-10-25 - so the frontier margin is 115 minutes, and the hub's newest item
is 107.7 (guest one) and 103.7 (guest two) minutes old at the end; Corpus A's 30- and 60-day windows are fresh at
both ends, where a real clock would have refused the tier from 2026-11-02; and the unindexed guest's hub has the
same items inside its seven-day reach at its anchor and at the end. **Why each base's own instant** (decided on the
maintainer's behalf): it moves nothing the guest wrote. A clock set back nearer the hub's anchor would put the
base's own writes - the live run's index entries, the PSTs' times, the logs - in the guest's future, which
nothing measured; and a fresh hub rebuild before freezing would have given guest one a hub its resting checkpoint
never had. The cost is the 14 to 18 minutes already on each hub, which the 90-minute proof includes.

**4. What keeps a run on the frozen instant.**

* **The order of a run** - `Testbed/README.md` section 4c: lease, restore the frozen checkpoint, stage,
  `Testbed/host/Set-GuestClockFrozen.ps1 -VMName <guest> -Verify` (exit 0 `FROZEN` or no suite), Outlook started
  NOT elevated on the tier profile (`Testbed/guest/Start-OutlookUnelevated.ps1`), the suite, rest. No restart and
  no hub rebuild after the restore: the restore puts the hub back as it was built and wipes every artifact, which
  is why the per-run rebuild (section 3b) goes. The Outlook start is the other thing step 9a used to hand the
  suite, besides its rebuild; section 5 says why it is there.
* **The guard**, `-Verify`, read-only: time sync off, the guest restored from the recorded checkpoint, and its
  clock no more than 2 minutes before the frozen instant and no more than 45 minutes after it. A restart or cold
  boot after the restore, time sync turned back on, or a hand-set clock each fail it, naming the cause.
* **`Restart-Guest.ps1` refuses a frozen guest** (time sync off), exit 2, before anything changes; `-Refreeze` is
  for work outside a run. It is the only script here that restarts a guest - audited: the installers pass
  `/norestart` (the product's installer `/NORESTART`, `vstor_redist.exe` `/norestart`, the SDK's `/norestart`,
  whose 3010 leaves the reboot to the operator), `Set-OutlookIndexingDisabled.ps1` restarts the `WSearch` service,
  not the guest, and every guest script that needs Outlook closed or a reboot says to use `Restart-Guest.ps1`. The
  idle-saver only saves, and a saved guest's clock stops (section 4.1f); a host restart saves every testbed VM. So
  what is left - a cold boot by hand, or time sync re-enabled - is what the guard catches.
* **The freshness and frontier checks stay as guards**, unchanged in what they check: on a frozen guest a hub that
  reads stale or future, or an emptied corpus window, now means the clock moved, and their refusals say so and name
  the restore and the guard (`T1/FrozenGuestClockTests`).
* **Making a new frozen checkpoint** - after changing a frozen guest: restore its frozen checkpoint, do the work
  (`Restart-Guest.ps1 -Refreeze` if it needs a restart), rebuild the hub if it should be fresh (step 9a), close
  Outlook with a restart, take an ordinary running checkpoint, `Set-GuestClockFrozen.ps1 -FromCheckpoint` it,
  record it in `frozenClocks` - the T1 pins then prove its instant - and make it the resting checkpoint.
* **`Testbed/host/Invoke-LiveTierOnGuest.ps1`**, the guest runner (merged from the guest-one work at `3e0ad4f`,
  while this was being done): its default start and resting checkpoints are now the frozen ones, and its
  `-SelfTest` holds them to `frozenClocks`; a start checkpoint whose time sync is off - read off the checkpoint
  itself - makes its PREPARE step the clock guard and the Outlook start alone: no `Restart-Guest.ps1`, no hub
  rebuild, and `Set-GuestClockFrozen.ps1 -Verify` must say `FROZEN` or the run stops as not tested; it reads the
  guest's crash events from the guest's own clock (section 5 says why); `-RestOnGreen` from a
  frozen checkpoint is refused, because a green checkpoint stands a run later than the recorded instant; and
  `-Checkpoint` with an unfrozen checkpoint still runs the old way, restart and step 9a included. Had it run a
  frozen checkpoint unchanged, its PREPARE's restart would have been refused (exit 2) and the run stopped -
  loudly, not on a moved clock.

**5. The first runs from the frozen checkpoints** - this branch's suite, first (`3fb921e` on guest two, `a6e90de`
on guest one) staged and run in README section 4c's order by a scratch driver that does what the runner's STAGE and
RUN do, with the guard between them; then, once the runner was merged, through `Invoke-LiveTierOnGuest.ps1` with the
section 4 changes (`1b5da1e`, `af3ba68`, `2f42c74`) - from the second on with `-RestingCheckpoint` naming the
guest's unfrozen resting checkpoint, which stays the resting one until this work is merged;
and X and Y, the scratch driver again, to compare the frozen and the real clock on one commit:

| Guest | Run | Guard before the suite | Suite | Safety |
| --- | --- | --- | --- | --- |
| `OutlookAI-Unindexed` | 1 | `FROZEN`, 110 s after the instant | 79: 54 passed, 25 failed - OUTLOOK.EXE crashed 3.5 min in (`OLMAPI32.DLL`, `0xc0000005`) and every later compose test met `RPC server is unavailable` | 0 tagged artifacts (2 late sent copies purged); tripwire 0 failures, 0 notes |
| `OutlookAI-Unindexed` | 2 | `FROZEN` | **79 of 79 passed**, no crash | 0 tagged artifacts; tripwire 0 failures, 0 notes |
| `OutlookAI-Indexed` | 1 (`a6e90de`, before master's guest-one fixes) | `FROZEN`, 2 min after the instant | 122: 101 passed, 21 failed - 13 `Store 'Corpus A' not found among 3 discovered index scopes` (and one `store scope not discoverable`), 7 `0x80041607` (`QUERY_E_TIMEDOUT`) on folder-scoped searches and `Search_TopOne` at exactly 100: the failures section 4.2e met on the real clock and master fixed at `3e0ad4f` (section 4.2f) - none a date check, and none `RPC server is unavailable`. The frontier test passed - `Hub population fresh ... its newest item is 26 min old ... under 115 min` - the corpus gate said `Freshness: OK`, and the date-range test asked for `the 30 days before corpus 'vm-indexed''s anchor`, 10 rows in 208 ms | 0 tagged artifacts; tripwire 0 failures, 0 notes |
| `OutlookAI-Unindexed` | runner A (`1b5da1e`, master merged; Outlook not started before the suite) | `FROZEN`, 99 s after the instant | 81: 55 passed, 26 failed - Outlook lost from `UpdateDraft_AddsAndRemovesAttachments` on, every later compose test `RPC server is unavailable`; the runner counted 0 crashes because it searched the guest's log from the HOST's clock (fixed, below) | 0 tagged artifacts; tripwire 0 failures |
| `OutlookAI-Unindexed` | runner B (`af3ba68`, Outlook started NOT elevated before the suite) | `FROZEN` | 81: 55 passed, 26 failed - the same, from the same test | 0 tagged artifacts; tripwire 0 failures |
| `OutlookAI-Unindexed` | X (`af3ba68`, the scratch driver, Outlook started, the guest's events read before the rest) | `FROZEN` | **81 of 81**, no crash event | 0 tagged artifacts; tripwire 0 failures |
| `OutlookAI-Unindexed` | Y (`af3ba68`, the same state on the REAL clock: `CP-13B-LIVE-GREEN`, time sync on, no restart, no rebuild, Outlook started) | - | **81 of 81** | 0 tagged artifacts; tripwire 0 failures |
| `OutlookAI-Indexed` | runner (`af3ba68`) | `FROZEN`, 85 s after | 127: 114 passed, 13 failed - 12 compose tests on `RPC server is unavailable`, and `LiveFolderIdentityTests.FolderId_Survives...` (a moved folder not yet found under its new parent - COM folder calls, no clock in them) | 0 tagged artifacts; tripwire 0 failures; frontier `24 min old`, corpus `OK` |
| `OutlookAI-Indexed` | runner (`2f42c74`), alone on the host | `FROZEN`, 77 s after | **127 of 127**, 0 crashes searched from the guest's clock | 0 tagged artifacts; tripwire 0 failures, 0 notes; frontier `24 min old ... under 115 min`, corpus `Freshness: OK`, the window from the data, 10 rows in 42 ms |
| `OutlookAI-Unindexed` | runner (`2f42c74`), alone on the host | `FROZEN`, 77 s after | **81 of 81**, 0 crashes searched from the guest's clock; the two `PROVED NOTHING` lines this guest always prints | 0 tagged artifacts; tripwire 0 failures |

**The compose failures are intermittent, and nothing ties them to the clock.** The same commit failed through the
runner from the frozen checkpoint (B) and passed 81 of 81 from it 25 minutes later (X), and 81 of 81 on the real clock
from the same disk state (Y); on guest one one frozen run failed and the next was 127 of 127; and the runner's run of
master on guest one on the real clock with the hub rebuilt, the same hour (`20261003-202603-indexed-dc1b5c5d51cd`,
section 4.2f), failed the same way with an OUTLOOK.EXE crash. In every failing run Outlook went away mid-compose; in
run 1 the guest's log recorded the crash (`OLMAPI32.DLL`, `0xc0000005`). Neither host load nor the runner separates
the failing runs from the green ones. *Its rate, its cause and its fix are section 4.6 (2026-10-04).*

**Two things the first runner runs taught, both fixed here.** (1) On a frozen guest the runner's crash count was
blind: it searched the guest's Application log from the HOST's clock, hours ahead of every event the guest wrote -
`Get-CrashEventsSince` now reads the guest's own clock before the suite, and `-SelfTest` pins it. Any other script
that hands a host instant to a frozen guest has the same trap; none does today. (2) A frozen start must still hand
the suite the state step 9a used to: Outlook running NOT elevated on the tier profile, its window up - the frozen
checkpoint holds Outlook closed so that staging can swap the commit in, and the suite then started Outlook itself, by
COM and without a window. The runner now starts it with `Testbed/guest/Start-OutlookUnelevated.ps1` after the guard.
(It did not stop runner B's failure; it restores the established hand-over, which every green run before Q130 had.)

**MSBuild with files the host dated after the guest's clock** - the other unmeasured item of section 4.1f: every run
above staged the commit's source with the host's file times, hours after the frozen clock, and every one built it,
proved it `TEST-READY` and ran it. Whether `dotnet test` then rebuilds every run - its outputs can never be newer
than those inputs - was not read off a build log, and the suite times do not say (7 min 37 s to 8 min 56 s frozen,
7 min 55 s on the real clock).

**6. Q130 (b), the window.** `LiveIndexSearchTests.ProbeParity_DateRangeQuery_HitsUnder2s` now asks for the 30
days BEFORE the declared corpus's anchor - `[2026-09-03, 2026-10-03)` for Corpus A, which select the same 24,596
of its items on every run - and, only where no corpus is declared, the last 30 days of the clock
(`T2/LiveDataWindow.cs`; `T1/LiveDataWindowTests` pins both branches, the 24,596, and from the compiled IL that the
test computes no window of its own). On a frozen guest the two coincide; on a real clock the test now keeps timing
a date predicate over the big store after 2026-11-01.

### 4.6 Crash dumps, and the compose crash they caught - both Outlook guests, 2026-10-04

**Why this section exists.** OUTLOOK.EXE died now and then in the compose tests on both guests - `wwlib.dll`
once on guest one (4.2f, run 4), `OLMAPI32.DLL` on guest two (4.5 run 1; 4.1e's E3 and run 9 at `0x2E411`),
and other runs lost Outlook mid-compose ("RPC server is unavailable") - and the guests kept no dump, so
there was nothing to read. Raw material: `.work\crash-repro\` and `.work\guest-live-runs\` of the agent
worktree `a467ede5` - every run's `summary.txt` and the report beside each of its 32 dumps (`<dump>.txt`). Three dumps
are kept whole - the first from master's compose runs (`20261004-011320-indexed-...`), the first soak's
(`20261004-013517-unindexed-...`) and the `MailItem.Close` variant's (`20261004-041412-unindexed-...`); the other 29,
about 1 GB each, were deleted for disk space once read.

**1. The guests keep full dumps now.** `Testbed/guest/Set-OutlookCrashDumps.ps1` writes Windows Error
Reporting's `LocalDumps` key for `OUTLOOK.EXE` and `WINWORD.EXE` - `DumpType` 2 (full), `DumpCount` 10 -
into `C:\OutlookAI-Q5\crash-dumps`, Authenticated Users Modify, because Outlook runs NOT elevated and a
filtered token cannot write where only administrators may. Nothing is installed. `Invoke-LiveTierOnGuest.ps1`
runs it after every restore with its own checkout's copy - no checkpoint changed, so the frozen instants of
4.5 and their T1 pins stand - fetches every dump into `dumps\<guest>\` before the guest rests, counts an
Outlook dump as a crash even when its event was missed, and puts `host/Read-CrashDump.ps1`'s first look into
`summary.txt`. Proved on both guests (`-ProveCrashDumps`): a throwaway copy of `rundll32.exe`, crashed on
`DebugBreak` in session 1 NOT elevated, as Outlook runs, left a 30 MB full dump within seconds. Outlook's
own dumps are about 1 GB (989 MB the first); every one of the 30 taken here was fetched and read.

**2. Reading them with what ships with Windows.** `host/Read-CrashDump.ps1` parses the minidump itself and
unwinds the faulting thread with `dbghelp.dll`'s `StackWalk64`, fed the dump's memory and each module's
`.pdata`; a frame is named `module!export+0xN` when its function is exported, else module, offset and the
start of its function. Checked first on a dump of a running `PING.EXE` on the host: every thread unwound to
`RtlUserThreadStart`, the export names right. Office has no exports for most of itself and no RTTI, and
Microsoft's symbol server publishes no PDB for this build (404 for `outlook.pdb` and `olmapi32.pdb` by their
GUIDs, while Windows' `combase.pdb` answered) - so Office frames stay module and offset, and what follows was
read off the dump's memory: the faulting code's bytes, the objects it touched, and the stack arguments of the
COM call in progress. No debugger was needed (QUESTIONS.md, decision of 2026-10-04, item 2).

**3. What every dump says** - 30 of 30, both guests, the master build and every variant below:

* `access violation (0xC0000005)` READING address 0 at `OLMAPI32.DLL+0x2E411`, on Outlook's MAIN thread.
* The faulting function (`OLMAPI32+0x2E3F8`) resets a holder - `if (m0) { m8->Release(); m8 = 0; ...;
  m0 = 0; }` - whose `m0` was `0x40000` and whose `m8` was NULL: its own invariant broken. The holder sits at
  `+0xA0` of an OLMAPI32 object whose `Release` (stack frame 3) is destroying it, beside the string
  `\REGISTRY\USER\<SID>\Software\Policies\Microsoft\Office\`; another object of the class, with the same
  string, lay destroyed cleanly in the heap.
* The destruction is reached from an INCOMING COM call: the message loop, `combase!CStdStubBuffer_Invoke`,
  `rpcrt4!NdrStubCall2`, oleaut32, then `OUTLOOK.EXE+0x207C60` - which sits in slot 6, `IDispatch::Invoke`,
  of dozens of Outlook's vtables - then a shared `Release` and some 25 Outlook frames into OLMAPI32.
* The call: that `Invoke` frame's stack arguments read `wFlags` 3 (a method) and one argument, `VT_I4` 0.
  In `new_draft` that is one call - `Inspector.Close(0)`, olSave, on the hidden compose inspector of the
  item `Items.Add` had just returned - and the tests' own error, "a draft may have been saved", is raised only
  past that point. A variant that closed with `Close(1)` crashed with `VT_I4` 1 (item 5).

So the COM objects involved: the hidden compose inspector of a just-created mail item, its close, and an
Office-policy watcher in OLMAPI32 that the teardown destroys half-built. The fault is inside Outlook.

**4. Reproducing it.** The compose classes alone - the 36 tests of `LiveUpdateDiscardTests`,
`LiveHtmlDraftTests`, `LiveDraftOptionsTests`, `LiveDraftTests`, `LiveHeadlessComposeParityTests`,
`LiveSignatureTests`, `Phase4LiveMcpToolShapeTests` and `LiveHeadlessGuaranteeTests`, the runner's
`-FilterSuffix '&(FullyQualifiedName~...|...)'` - on master `9d2cb8c`: guest one crashed in 2 runs of 10,
guest two in none of 4; guest two's full tier in none of 3. A scratch soak test, never merged (QUESTIONS.md,
item 5): a `LivePhase4Fixture` class whose one test runs `new_draft` with a signature override 400 times on the
hub, revises every second draft with `update_draft` and deletes each through `LiveOutlookTestMailer`, run with
`-FilterSuffix '&FullyQualifiedName~LiveComposeSoakTests'` from a commit built outside the branch. It crashed
after 64 iterations on guest one and after 37, 17 and 49 on guest two: 4 crashes in 171 new drafts, about one in 40.

**5. What it is not** - each a scratch variant of the compose, run in the same soak:

| Variant | Iterations to the crash, each run |
| --- | --- |
| the build as it was | 64; 37, 17, 49 |
| the Word document released before the close | 26, 164 |
| a 250 ms pumped wait before the close | 52, 244 |
| a 2 s pumped wait before it | 237 |
| OutlookAI's add-in disconnected first (`COMAddIns('OutlookAI').Connect = False`, read back False) | 70, 4 |
| `CurrentItem.Save` and `Close(olDiscard)` instead of `Close(olSave)` | 13, 25, 7 - in `Close(1)` |
| the editor promoted first (park, `Activate`, pumped settle), as update_draft does | 59, 40 |
| the item saved before its inspector was opened | 18, 39 |
| no signature override | 73, 16 |
| no signature override and no picture embedding | 49, 126 |
| no `SendUsingAccount` | 121, 56 |
| the `Items` collection released right after `Add` | 46, 36 |
| the inspector used only for Outlook's own signature injection, closed with no Word work in it | 24, 28 |
| the same, ended by `MailItem.Close(olDiscard)` instead of the inspector's close | 308, and the second run - in `MailItem.Close(1)` |
| `update_draft` alone, 400 revisions of one draft (an item opened by EntryID) | none in 400 |
| saved, released, re-opened by EntryID, and revised the update_draft way | none in 4 runs of 400 |

A plain re-open followed by the old compose also never crashed (2 runs of 400), but it committed nothing - 400
bodies of 400 missing - because a saved draft's hidden inspector edits nothing without the update_draft path,
as `ReviseHeldDocument`'s notes say. So: closing an inspector that was ever opened on the object a creator
returned - whatever was done in it - crashes Outlook about once in 40; an inspector on an item opened by
EntryID never did. Why the policy watcher is half-built is Outlook's to answer; OutlookAI's part is never to
hand it that close.

**6. The fix** (`c4595b1`, `c31a9d9`). `new_draft`, `reply_draft`, `replyall_draft` and `forward_draft` never
open an inspector on the object `Items.Add`, `Reply`, `ReplyAll` or `Forward` returned. `ComposeReopened`
saves it as it stands, re-opens it by EntryID, releases the original, and composes through
`ReviseHeldDocument` - the update_draft path: one held inspector, the editor promoted, the signature and body
placed in Word, `CurrentItem.Save`, `Close(olDiscard)`; the HTML splice stays the fallback. The old in-place
compose and `CloseHiddenInspector` are gone. One behaviour had to move with it: Outlook injects an account's
default signature only into an inspector on the creator's object, so `ComposeReopened` inserts the account's
configured default itself - the profile's "New Signature" or "Reply-Forward Signature", in Outlook's default
profile, as `list_signatures` reads them - only when the agent named no signature and Outlook injected none
(QUESTIONS.md, item 4). `T1/ComposeOnReopenedDraftTests` pins it from the sources: no creator does inspector
work, `ComposeReopened` releases the creator's object before the revision and hands the revision the re-opened
item, and inspectors are opened only in `ReviseHeldDocument` and `TryApplySignatureOverrideToDraft`, on items
opened by EntryID.
*The trade-off:* a re-opened draft's hidden inspector edits nothing until it is activated, so new drafts now take
update_draft's editor promotion - the window parked off-screen, `Activate`, anything shown hidden again - on a
windowed Outlook too, where D49's compose left a windowed Outlook's windows alone. Nothing reaches the screen
(`LiveHeadlessComposeParityTests` and `LiveHeadlessGuaranteeTests` pass), and update_draft has always done this.

**7. After** - `c31a9d9` (`c4595b1` where named), the same guests, the same frozen starts:

| What | Before (`9d2cb8c`) | After |
| --- | --- | --- |
| The soak, 400 new drafts a run | 4 crashes in 171 drafts (4 runs) | none in 2,400 - 4 runs of 400 on `c31a9d9`, 2 on `c4595b1`; 0 bodies missing |
| The compose classes (36 tests) | guest one 2 crashes in 10 runs, guest two 0 in 4 | 36 of 36 in 10 runs of 10 on each guest, no crash |
| The full tier | guest two 81 of 81 in 3 runs of 3 | guest one 127 of 127 in 2 runs of 2, guest two 81 of 81 in 2 runs of 2, no crash |

Had the rate stayed at 4 in 171, 2,400 clean drafts would have had a probability of about 10^-25; at the lowest
rate any crashing variant showed (one in 309), about 0.04 %. The suite-level comparison alone would not prove it
(two crashes in 14 runs against none in 20); the soak does. The business-account signature test
(`NewDraft_BusinessAccounts_BodyAboveTheirOwnIntactHtmlSignature`) failed on `c4595b1` - the re-opened draft gets
no signature from Outlook - and passes on `c31a9d9`, which inserts it (item 6).

**8. What stays open.** The `wwlib.dll` crash of 4.2f run 4 was never caught in a dump. It struck in
`NewDraft_Hub_SignatureOverride`, inside the compose this replaced, so it is most likely the same close
reaching Word's teardown first - inferred, not observed; every dump here was `OLMAPI32+0x2E411`, and the
guests now keep the next one. `TODO.md` carries it. The default-signature insertion reads the default
profile's settings: a signature Outlook would pick by other means - roaming signatures in an Exchange Online
mailbox, a Group Policy MailSettings value - or an Outlook running on another profile was not measured, and
no guest has one.

---
## 5. Which tests are in which bucket, and how to find out

The classification is **two traits on the test itself**, not a list in a document that can drift -
and not three traits either. It used to be three, and the third one was the problem.

* **`Category=Live`** means "this test needs a mailbox". It is the gate of the non-live run on
  the build VM, `OutlookAI-Build`, which has no Outlook at all - as GitHub's runners had none
  while CI existed (until 2026-10-03).
* **`Requires`** says *what of a machine* the test needs, from one closed vocabulary, declared
  **per method**. Nothing else is declared: which bucket a test is in is a question asked of
  `Requires` at filter time.

**And one trait that is not a bucket: `Writes=Nothing` (Q74, 2026-10-03).** It says what a test DOES to
the machine - nothing: no mail item or folder, no signature, no registry value, nothing on the
user's screen - declared **per method**, with that one value, and absence meaning "may write". It
is not the retired third axis come back: that one restated `Requires` by hand and could only drift,
while this one cannot be derived from `Requires` at all, and it is not trusted either -
`T1/ReadOnlyLiveTestTests` walks the compiled code of every carrier, its fixtures included, and
fails the build on any way it can reach a write. It exists for one machine: the Exchange test VM
(section 4.4) runs only live tests that need an Exchange profile AND carry `Writes=Nothing` until the
maintainer approves its Phase 2 write-safety design, and every test needing Exchange must carry it.
Until Q116 (a), 2026-10-03, the machine was the maintainer's workstation, which runs no live test
now. The run is `Testbed/README.md` section 4e; its filter, derived in
`McpServer/OutlookAI.McpServer.Tests/T2/LiveRunFilters.cs` and pinned there and here, is:

`Category=Live&Writes=Nothing&Requires=CachedExchange&Requires!=DelegateStore`

- the cached-Exchange carriers, and none that needs a delegate store until the shared test mailbox
exists (Q109). Fifty-four live tests carried the trait on the morning of 2026-10-03; seven of them
needed Exchange, and those seven were the workstation run.

**The three buckets, all computed:**

| Bucket | How it is selected | Size |
| --- | --- | --- |
| non-live (the build VM) | `--filter "Category!=Live"` | 2,226 cases |
| VM | `--filter "Category=Live&Requires!=DelegateStore&Requires!=CachedExchange"` | 121 |
| Exchange-only (the Exchange VM, section 4.4; the workstation before Q116 (a)) | `--filter "Category=Live&(Requires=DelegateStore\|Requires=CachedExchange)"` | 7 on the morning of 2026-10-03; 9 once section 4.4 added two |

**The vocabulary, all twelve values.** Ten of them this VM can be given; two it cannot - both are an
Exchange profile.

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
| **`DelegateStore`** | **a delegate/shared mailbox. One of the two capabilities no test machine can be given** |
| **`CachedExchange`** | **a cached Exchange mailbox, whose entry ids are Exchange's 70-byte form. The other one (Q74 C1, 2026-10-03)** |

`Tools/Checks/check-pinned-constants.ps1` fails the build if any of those twelve names stops
appearing in this file, so the table above is load-bearing text and not decoration.

**Why `DelegateStore` and `CachedExchange` are the only production-only capabilities.** A delegate/shared mailbox is
indexed with its folder hierarchy FLATTENED - an item in the delegate's `Archive/SomeFolder` is
published as `<host>/1/<delegate>/SomeFolder`, every intermediate folder dropped. A local PST
cannot be made to have that property, and faking it would manufacture confidence in the one area
this product has most often been surprised by. `CachedExchange` is the same kind of fact one level
down: a cached Exchange store's object model hands out 70-byte Exchange entry ids, so the 24-byte id
decoded from an index URL is REFUSED there, while on a PST those 24 bytes are the entry id itself and
should open. One check had both facts in it and could pass on neither kind of machine but one; it is
two halves now - `LiveDecodeVerifyTests.ShortDecodedId_IsRejectedByGetItemFromID_DiscoveryRecorded`
(`Requires=CachedExchange`, the workstation) and `..._OpensAsTheItemItself_OnAPstStore` (the guests;
INFERRED from `Mapi/EntryIdCodec.cs` and section 8 item 24, not yet run - its first guest run confirms
it or fails it). The six capabilities that used to sit beside `DelegateStore`
(`SearchIndex`, `MailAccount`, `Transport`, `MultipleStores`, `SmallHubStore`, `ProbePopulation`)
stopped being production-only the moment this machine's shape was settled: sections 1 and 2 build
every one of them.

**The third axis is gone, and this is what it was.** A `LiveTier` trait held `Portable` or
`ProfileBound` and had to be kept in agreement with `Requires` by hand - a computed value
maintained manually, which is the exact drift `T1/LiveTierInventoryTests` exists to prevent. It
was paired with CLASS-level `Requires`, so a class read as the union of everything any one of its
methods needed. Between them they reported **96 tests that could not leave the maintainer's
machine**. Re-read method by method, the real floor is **six** - the six the production-only
filter selected then; seven since `CachedExchange` (Q74 C1) named the one more that could not leave. `LiveTierInventoryTests` now refuses the retired trait outright and refuses a
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

`T1/LiveTierInventoryTests` enforces all of it in the non-live suite, together with the rule that every live
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
  machine it throws instead. **Since 2026-09-27 (Q76) that handful includes all four
  `LiveResumableScanTests`**, on a hub whose exhaustive scan fits one page of two: they need a
  chain of pages, so at least five mail items in the hub, which the rebuilt hub population has -
  on a guest, read it first as a hub nobody rebuilt. It also includes `LiveStaleIndexRowTests`
  when the delegate's folder tree does not list the probe folder nested at that moment, a
  delegate test and so in practice a refusal on a Production machine. Both used to return green
  with a line that did not say so.
* **A non-empty Outbox fails `LiveDisconnectRecoveryTests` on every profile** - not a skip and
  not a `PROVED NOTHING:` (Q76, 2026-09-27). The test closes Outlook, which S7 forbids while
  anything is queued, and on a guest nothing but a test ever queues mail, so a count above zero
  there is residue that guard 4 above exists to refuse: find out what was not delivered before
  re-running. On a working profile it may be the user's own unsent mail. An unreadable count
  fails the same way. **Since 2026-10-03 (Q101) the Outbox and Inspector counts are taken before
  the test looks at windows at all**, so a run that starts with no Outlook window open is covered
  too - it used to make Outlook exit with neither count taken.
* **Since 2026-10-03 (Q101) the `PROVED NOTHING:` handful also includes** `T3/OutlookAvailabilityLiveTests`
  - its retry-guidance check whenever Outlook is healthy, and its freshness contract when this
  machine's Windows Search index cannot be reached at all - `T3/ComHostSupervisionLiveTests.NoComHostSurvivesTheServer`
  when no COM host was started (Outlook not running), `LiveUiSearchBackendTests` when a policy-hive
  `DisableServerAssistedSearch` overrides the value it flips, and `LiveHeadlessGuaranteeTests`' read
  and thread window checks when no hub hit opens or the one that did carries no conversation. Each
  used to pass green, most without a line. On a `Production` profile each now FAILS instead - and
  **read the first one before trusting a Production run**: the retry-guidance check can only run
  while Outlook is starting, hung or unavailable, so on a machine whose Outlook is healthy that test
  now fails by design. Its search test also FAILS on any `search` error other than an unreachable
  index (outlook_health reporting `index.provider` as `unavailable: ...`); it used to pass on every
  error, the one thing "search must degrade, never fail" is there to catch.
* `SKIP (user protection):` - `LiveDisconnectRecoveryTests` standing aside for a person: recent
  keyboard or mouse input with Outlook windows open, an open Inspector, a window appearing while
  Outlook restarts headless, a second window beside the one the test opened, or the installer mutex
  already held (that one skips step 3b only). **Only on a `Production` profile** (Q101, 2026-10-03),
  where a real user may be at the keyboard. On a `Portable` profile - and on any profile value not
  yet classified - each of those FAILS, naming what it saw: nobody uses a test guest during a run, so
  the same state there is residue or interference, never a person.
* `COM leaf matches after N walk(s) over T s` - `LiveStaleIndexRowTests` re-walks the delegate folder
  tree every 15 s for up to five minutes before it refuses (Q101, 2026-10-03), because Exchange syncs
  that hierarchy lazily. A run that needed most of the five minutes is the evidence for moving the
  bound.

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
  *(Corrected 2026-09-27: the mechanism is no longer this, and its premise was measured wrong. A
  PST does not refuse `GetDefaultFolder` for Junk - it CREATES the folder (measured on the guests'
  PSTs, 2026-09-24; section 4.1). Since Q84 (`4f82004`, the same day) the census never calls it on a
  non-Exchange store: it marks Junk self-pruning only when the store's Inbox designates a Junk
  E-mail folder in `PR_ADDITIONAL_REN_ENTRYIDS` and that entry id opens
  (`McpServer/OutlookAI.Core/Com/SpecialFolders.cs`). The conclusion stands for that reason instead: a Junk
  folder the Inbox does not designate is counted as ordinary, and a decrease in it fails rather
  than being noted. Whether the Junk Email folders on the guests' stores carry that designation
  has not been read.)*
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
   *(2026-09-27: the computer name IS recorded for the script-built guests, by construction:
   `Testbed/host/New-AnswerFile.ps1` sets it from the VM name, `OutlookAI-` shortened to `OAI-`,
   unless `-ComputerName` is given, refuses any name not starting `OAI-`, and writes it into
   `Testbed/guest/autounattend.template.xml`'s `<ComputerName>`; the guests report `OAI-INDEXED`
   and `OAI-UNINDEXED`, and the first-logon script refuses a guest named otherwise
   (`Testbed/README.md` section 1b). It is still unrecorded for the original guest,
   `OutlookAI-TestVM`, and stays so: that guest was retired and deleted on 2026-10-03. Defender
   exclusions stay open on every guest: no script here adds one, and none has been read off a
   guest - README section 6 item 2.)*
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
10. ~~The scheduled-task recipe for session 1~~ - **ANSWERED by `Testbed/guest/Register-InteractiveTask.ps1`,
    whose banner is the recipe - with ONE CHECK STILL OWED ON A GUEST (Q91, 2026-09-27).** Task
    `OutlookAI-Interactive`, principal `-LogonType Interactive` at `-RunLevel Highest` (the default)
    or `Limited`, a job directory `C:\OutlookAI-Q5\jobs\<id>` holding `cmd.ps1` (the wrapper), the
    inline work as `work.ps1`, `out.txt` and `exit.txt` - output goes to a file because a task's
    stdout never reaches the caller. **Exit-code capture changed on 2026-09-27:** an inline `-Script`
    used to be pasted into `cmd.ps1`, where its `exit N` ended the wrapper itself and `exit.txt`
    recorded 0 - a failure read as success - and a body that did not parse left no `exit.txt` at all,
    so the caller waited out `-TimeoutSeconds`. It now runs as its own `work.ps1`, the shape
    `Testbed/host/Restart-Guest.ps1`'s quit task has used since Q69, whose `exit 10` and `exit 16`
    reached `exit.txt` on `OAI-INDEXED`. **Proven on the host only**: 15 cases - `exit N`, `exit` in a
    function, a bare `exit`, a throw, `-ErrorAction Stop`, `$ErrorActionPreference = 'Stop'`, a
    non-terminating error, a trailing native exit code, a parse error, non-ASCII text - for `-Script`
    and `-ScriptPath`, at both run levels, each job written by and run under both PowerShells, with
    the old wrapper as the control (it recorded `exit 7` as 0 and wrote no `exit.txt` for the parse
    error); `-ScriptPath` wrote byte-identical `cmd.ps1` files old and new
    (`.work/aa5e-2026-09-27-q91-inline-exit/`). **PENDING - at the next use of either guest, before
    anything relies on an inline job's exit code** (a minute; no Outlook, no mailbox): stage the new
    script beside `OutlookMapiInterop.ps1` in `C:\OutlookAI-Q5` and, over PowerShell Direct, run
    `.\Register-InteractiveTask.ps1 -Script "'q91'; exit 7"` - expect `exit 7` printed and returned,
    and a kept job directory holding `work.ps1` - then the same with `-RunLevel Limited` (`exit 7`),
    with `-Script "'q91'; throw 'boom'"` (`exit 1`, `WRAPPER CAUGHT`), and with `-Script "'q91'"`
    (`exit 0`). Record the result here and in the script's banner.
11. The exact PST file paths and names for all four stores, and the mapping from file name to
    display name.
12. .NET SDK version, clone path, build configuration, and how the built server exe reaches the
    path the tier-3 tests expect.
13. How results, screenshots and logs get out of the guest, and where `ScreenCapture` writes.
14. **CLOSED, with a caveat that matters.** `Docs/v3-probes/soakfix13-probe-sweep-cost.ps1` was
    gitignored, so it lived on one machine and is gone. It has been **reconstructed in the
    repository** as `Testbed/guest/Measure-SweepCost.ps1`, written from the shipped sweep's own
    source rather than from memory, and read-only by construction. **It first ran on
    2026-10-03** (section 4.2e), against Corpus A on `OutlookAI-Indexed`: that run found two
    faults - default folders resolved by the call that creates them, and every COM collection
    unrolled by PowerShell - and fixed both; its banner now says what it did, and that its
    absolute milliseconds are PowerShell's rather than the sweep's.
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

17. **NARROWED 2026-09-24 - the measurement CORPUS is unchanged; the fixture POPULATIONS carry the
    rest.** The corpus still writes `IPM.Note` with a subject, a body, a read state, message flags
    and two date properties, and nothing else, on purpose: its shape key is what every published
    measurement rests on. What tests need beyond that is built into the hub, the bystander and the
    identity store as generated populations (section 3b): senders and recipients - the store's
    owner a resolved recipient of every received item since generator v2 - attachments of four
    kinds, conversations, created subfolders, and, since v2, UNDATED items of other classes
    (appointments, contacts and tasks; hub 12, bystander 42) for the three
    `LiveOrderKeyCollationTests`. **Still in neither: HTML bodies, categories, follow-up flags, and
    unsent drafts** - the last deliberately: a new unsent mail item's first save lands in the
    profile's DEFAULT store's Drafts (measured on `OutlookAI-Unindexed`, 2026-09-24), and a
    population is never built into the default store. A test needing one of those is in the VM
    bucket and still fails here - because of what the generator builds, not because of the machine,
    and nothing in the traits says so. **None of the v2 populations has been built on a guest yet**
    (section 4.1 step 6 is v1's build, which was not clean).

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
    **The silently-empty cases found since are fixed the same way.** Three iterations on
    2026-09-15 (`LiveFolderScopeTests`, `LiveSignatureTests`, `LiveIndexSearchTests`, all through
    `T2/LivePopulationCoverage`; `TODO.md` has them). On 2026-09-27 (Q76) the early returns of
    `LiveResumableScanTests` - four, not the three first counted, plus its acceptance, which on a
    one-page hub compared two sets that agree by construction - and the delegate-tree return of
    `LiveStaleIndexRowTests`, pinned by `T1/LiveEarlyReturnGuardTests`; the same decision made
    `LiveDisconnectRecoveryTests`' non-empty Outbox a FAILURE on every profile rather than an
    announcement (section 6 says why). **The maintainer decided the rest on 2026-10-03 (Q101)**,
    case by case. `T3/OutlookAvailabilityLiveTests` (both tests), `T3/ComHostSupervisionLiveTests`'
    no-host branch, `LiveUiSearchBackendTests`' policy-hive skip and `LiveHeadlessGuaranteeTests`'
    read/thread checks went through `T2/LivePopulationCoverage` the same way - and the search test
    now FAILS on every error but an unreachable index, which it used to pass on whatever went wrong.
    `LiveDisconnectRecoveryTests`' user-protection stops (a user active in the last three minutes, an
    open Inspector, a window appearing mid-scenario, a second window beside its own, the installer
    mutex) took the INVERSE: a `SKIP (user protection):` line on `Production`, where a person may be
    at the keyboard, and a failure on `Portable`, where nobody is
    (`LivePopulationCoverage.StandAsideForAUser`). The same decision moved that test's Inspector and
    Outbox counts ahead of its window check, and gave `LiveStaleIndexRowTests` a bounded wait for the
    delegate tree. All pinned by `T1/LiveEarlyReturnGuardTests`. What still passes on a skip:
    `LiveAttachmentKindRecallTests`' two `SKIP (the parent-open assertion ONLY ...)` returns, kept on
    purpose on 2026-09-15 - that line is accurate, and the test has asserted recall above them - and
    `T3/OutlookHealthLiveToolShapeTests`' advice assertion, skipped where the index is unreachable.
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

    **What follows for the live tier on the guests - DECIDED 2026-09-24 and built: the tier runs at
    `RunLevel Limited`.** The index tests need the index to MOVE while they run - a test creates an
    item and waits for the index to show it - and on these guests the index moves only while a
    NON-elevated Outlook runs. Run through `Register-InteractiveTask.ps1` as it stood, Outlook is
    elevated, `IsInstantSearchEnabled` is `False`, and nothing a test creates is ever indexed: the
    "not indexed yet" state, permanently, which the suite reads as a slow indexer. So the tier's
    session-1 route must run Outlook - and the test host with it, because an elevation mismatch
    breaks COM attach (v3.MD S8) - at `RunLevel Limited`. Of the three routes (a `-RunLevel`
    parameter on `Register-InteractiveTask.ps1`; a second, Limited task for the tier; Outlook started
    by `Start-OutlookUnelevated.ps1` with the suite in a Limited task beside it) the first was
    chosen: `Register-InteractiveTask.ps1 -RunLevel Highest|Limited`, **Highest by default** because
    every existing caller expects it and the installers refuse without it. `Testbed/README.md`
    section 4c carries the run lines, why, and the elevation audit - nothing in the live tier needs
    elevation; three things need the SAME level as Outlook. Measured on `OAI-UNINDEXED`: the default
    gave the job `High Mandatory Level`, `-RunLevel Limited` gave `Medium Mandatory Level`, the
    suite's source tree rebuilt from scratch at Limited and `--list-tests` discovered the live
    tests there; a live test was NOT executed at Limited - the tier has not yet run on any guest.

    **A side finding for the identity tests, not chased.** ~~The index names a store by the name in
    its profile's service, not by the root-folder name COM reports.~~ **CORRECTED 2026-09-27 (Q92,
    item 24), by measurement - it is the other way round:** the index names a store by the store's
    OWN display name (`PR_DISPLAY_NAME` on the store, which is its root folder's name), while
    `Store.DisplayName` reports the name in the profile. At `CP-15` the identity store's profile
    sections both say `identity@vm.invalid`; its root folder and `PR_DISPLAY_NAME` say
    `Outlook Data File` - so whatever named it changed the profile and not the root folder, since a
    root-folder rename through `Rename-OutlookStore.ps1` changes both (measured on the tier store,
    item 24). The tier store appears as
    `tier@vm.invalid($93f42b43)`, but the identity store - named by `Add-OutlookPstStore.ps1` through a
    root-folder rename - appears as `Outlook Data File($b25ac20a)`, its folders under
    `/Outlook Data File`, while `Store.DisplayName` reads `identity@vm.invalid` (identified since Q92
    by its hash - `ComputeHash(Store.StoreID)` is `b25ac20a` - not by elimination). `IndexSearchService.TryDiscoverStoreScopeByAddress`
    accepts a store only when the index's name EQUALS the address, so it cannot find the identity
    store here. **Checked by reading, 2026-09-24 - nothing changed:** the identity tests themselves
    (`LiveDraftTests.IdentityDrafts_BusinessAccounts_...`, `LiveDraftOptionsTests.NewDraft_BusinessAccounts_...`)
    are COM only and never reach the index, and every test caller of the method passes the hub, the
    SF-6 probe store (the hub) or an INDEXED-list entry, which the settings keep the identity store
    out of. **One path does need the index name to be the address: `outlook_health`.** For every
    Outlook store whose name contains `@` and that the index sample does not name, it asks this
    method (`MailService.AddStoresMissingFromIndex` -> `StoreHasIndexRows` -> `ProbeStoreInIndex`), a
    `false` becomes a problem ("the local index holds nothing for identity@vm.invalid"), and any
    problem makes the verdict `degraded`. `LiveHealthTests.Health_OnThisMachine_ReportsOkWithFullDetail`
    (Requires `SearchIndex` and `MultipleStores`, so the VM filter keeps it on the indexed guest)
    asserts `"ok"` - so on `OutlookAI-Indexed`, with the identity store indexed under
    `Outlook Data File`, it should fail. INFERRED from the code and the measured index name; not run.
    `ListAccounts_ExactAccountsDelegatesAndFlags` would fail the same way (`InLocalIndex` for every
    watched store) but carries `Requires: DelegateStore` and is never selected on a guest. The corpus
    tool asks the same question (`CorpusCommands.cs`, `corpus-indexed` for a store named like an
    address), so waiting for an identity population to reach the index would never end either. ~~The
    index name is fixed where the store is made: a PST named through its profile service, not by a
    root-folder rename, would be indexed under its address - untested, and a question for however
    `identity.pst` gets its real Inbox (section 3b).~~ **Superseded by item 24:** the index name is
    not fixed - it is the store's own name, a root-folder rename changes it, and the index follows
    the rename (by dropping the store and indexing it again). A name set only in the profile service
    is exactly what the index does NOT use. The product fix is Q92: find a store's slice by its hash.
    **Since 2026-09-27 the identity store is minted and renamed BEFORE it is attached to the tier
    profile** (section 4.1b): `Rename-OutlookStore.ps1` renames the store itself - item 24 measured
    that exact script changing the name the index files a store under - so by item 24's rule the
    index should file it under `identity@vm.invalid`. The tier profile's service section names it
    so too, but that is not what the index reads. The indexed guest has not been rebuilt this way
    yet, so this is still unmeasured.
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
24. **MEASURED 2026-09-27 (Q92) - how the index identifies a store, and what a rename does to it.**
    On `OutlookAI-Indexed` from `CP-15-SIGNATURE-SUITE-STAGED` (own checkpoint `CP-15a-Q92-BEFORE`,
    restored to `CP-15` afterwards), Outlook started NOT elevated through `Start-OutlookUnelevated.ps1`,
    the index read only through `SELECT` statements, COM read only (store properties and one table
    row), the renames only through `Rename-OutlookStore.ps1`. Findings, each measured:

    - **`($Hash)` is Microsoft's documented store hash** (*Algorithm to Calculate the Store Hash
      Number*, MAPI reference; the same code is MFCMAPI's `ComputeStoreHash`): `h = h*33 + x` over the
      blob's little-endian DWORDs, then its trailing bytes. The blob is the store's `PR_ENTRYID` - which
      is `Store.StoreID`, byte for byte - for a PST, and the profile's `PR_MAPPING_SIGNATURE` for a
      cached Exchange store (`'.PUB'` mixed in for public folders). All three stores on this guest
      reproduce exactly from `Store.StoreID`: `93f42b43` (tier), `b25ac20a` (identity), `23a27f0d`
      (corpus). A PST's entry ID is fixed provider bytes plus the file's full path, so its hash
      changes when the file moves and never when the store is renamed.
    - **The name in the URL is the store's own** `PR_DISPLAY_NAME` (= its root folder's name), not
      `Store.DisplayName`, which reports the profile's name. The identity store differs (item 22).
      Two stores share a name on this guest: `Outlook Data File($b25ac20a)` (identity) and
      `Outlook Data File($23a27f0d)` (the corpus, another profile's store - the index is per Windows
      user, not per profile, and lists every profile's stores).
    - **`DIRECTORY='mapi16://{SID}/'` lists exactly one row per store root, in 2-3 ms**, including the
      two stores that hold no mail. The product's discovery sample (`TOP 2000 ... System.Kind='email'`,
      314-370 ms) saw only the corpus store, on every reading.
    - **The store UID in every item URL** (EntryID bytes 4..19) is the PST's `PR_RECORD_KEY`
      (= `PR_STORE_RECORD_KEY` = `PR_MAPPING_SIGNATURE` on a PST): `026047AA...` on all 20,012 corpus
      rows. A store with no item rows has no UID in the index at all.
    - **A rename drops the store from the index at once.** Renaming the corpus store's root folder,
      Outlook running: `Outlook Data File($23a27f0d)`'s root and folder rows were gone within 30 s and
      the catalog fell from 20,482 items to 455 within a minute, while Outlook re-pushed the store
      under `q92-corpus-renamed($23a27f0d)` - same hash - and the index held all 20,028 rows again
      13 minutes after the rename (about 2,000 items a minute). Renaming the tier store, which holds
      only folders: its 16 rows were gone within two seconds and nothing came back in 5 minutes;
      after a graceful restart and a new Outlook start, only the store-root row returned
      (`q92-renamed@vm.invalid($93f42b43)`), and the 15 folders had not returned after 8 minutes.
      `SCOPE='<that root>'` then matches 0 rows while `DIRECTORY` lists it.

    Not measured, and not measurable on a guest: the cached-Exchange half (whether the
    `PR_MAPPING_SIGNATURE` readable through `PropertyAccessor` equals the profile's, which the hash
    uses) and the delegate `/1/<name>` subtrees - both are the maintainer's workstation's shape.
    `Store.PropertyAccessor` reads raised no Object Model Guard prompt at `CP-15` (item 23).

    **Borne out by the Q87 rebuild of the same guest** (section 4.2b, `CP-10C-INDEXED`): the minted
    identity store, renamed `identity@vm.invalid`, is filed as `identity@vm.invalid($be889d8b)` - its
    own name, and a new hash because its file is a new path; the host reproduces `be889d8b` from its
    `StoreID`. `DIRECTORY` lists the three roots in 2 ms, and no two stores share a name any more.

    **Extended 2026-10-03 (Q99) - every edge case a PST can take, and what the product now does
    with it.** Same guest from `CP-15C-SIGNATURE-SUITE-STAGED` (own checkpoint `CP-15C-q99-A-base`,
    restored to `CP-15C` afterwards), Outlook NOT elevated, stores attached only through
    `guest/Add-OutlookPstStore.ps1` (which gained `-Format Ansi` for this), items created only by the
    corpus tool, the index read only through `SELECT` statements. Ten stores in all, every one filed
    under exactly `ComputeHash(Store.StoreID)`, four of them predicted on the host from the file path
    before Outlook was asked:

    - **ANSI PST** (`q99-ansi.pst`, an 84-byte 8-bit entry ID): `($54556ee0)` = the hash of its
      `StoreID`. Its `Store.DisplayName` came back ONE CHARACTER SHORT (`q99ansi@vm.invali`) while
      the root folder and the index say `q99ansi@vm.invalid` - a name lookup cannot find it at all.
      `IsInstantSearchEnabled` read False for it while its items were indexed within seconds, so that
      flag is no evidence either way.
    - **A copy at another path is another store**: `moved\q98-scratch.pst` filed as `($befbe850)`
      beside the original's `($65d10200)`, both under the name `q98scratch@vm.invalid` - two roots of
      one name that only the hash tells apart. Attached to ONE profile together, Outlook re-keyed the
      second (`PR_RECORD_KEY` `C4B808E7...` instead of the copy's `44616A9B...`); the hash, which is
      over the path, did not move. The original's root stayed listed while the store was in no open
      profile.
    - **One PST in two profiles is one root** (`identity@vm.invalid($be889d8b)` from `OutlookAI-Tier`
      and from `IdentityMint`): the hash is over the file, not the profile.
    - **The URL spelling**: the hash is lowercase hex WITHOUT leading zeros - `($ce9d6e4)` for
      `0x0CE9D6E4` (`q99-lz-12.pst`, whose path was searched for that property) - so it is compared as
      a number. A store named `q99 50% off*?x` is filed `q99 50%25 off%2A%3Fx($5159380d)`: `%`, `*`
      and `?` are percent-encoded, a space is not.
    - **A catalog reset** (`ISearchCatalogManager::Reset`, elevated, with the tier profile open): all
      three roots were listed again within 30 s and the open profile's folders re-pushed (464 items);
      the corpus store, whose profile was not open, kept only its ROOT - `SCOPE` on it matched nothing,
      because `SCOPE` matches what lies below a URL and never the URL's own row. Opening its profile
      (non-elevated) re-pushed it under the same name and hash: 8,877 rows nine minutes later, still
      climbing. A non-elevated reader can read the catalog's status, counters, crawl-scope rules and
      the roots.

    **What the product does with it** (`StoreIndexMatcher`, `MailService.TryGetStoreIndexMap`): it
    lists the roots with one `DIRECTORY` statement, computes each store's candidate hashes, and ties a
    store to the root that carries one of them and that no other store claims. A PST no root carries
    the hash of is NOT INDEXED - never searched through a same-named root of another store or profile.
    Every store whose hash input is not measured - cached Exchange (the documented input is the
    profile's `PR_MAPPING_SIGNATURE`; the product reads it from the store and from the profile section
    named by `PR_EMSMDB_SECTION_UID`, and also tries the documented entry-ID-plus-`.ost`-path variant),
    an IMAP or Outlook.com `.ost`, a store whose id would not read - is tied by its hash when one fits,
    and otherwise resolved by the name rule exactly as before, as are delegates (under the owner's
    `/1/<name>`) and every store when no map can be built. `outlook_health` says which, per store.

    **Q98(f), measured the same day** with the corpus tool's undated-item probe
    (`--undated-index-wait`), in a scratch Unicode PST and in the ANSI one: an appointment, a contact
    and a task saved into a PST were in the index within 3-6 s. The APPOINTMENT and the TASK carry
    `System.Message.DateReceived` - their creation time - and so sort and window like mail (the
    appointment's `System.ItemDate` is its start); the CONTACT has NO `DateReceived` (NULL), only
    `DateCreated`/`DateModified`/`ItemDate`, all its creation time. Every column the product's
    `ORDER BY` reads was present for all three except the contact's `DateReceived`. Kinds:
    `calendar|communication`, `contact|communication`, `task|communication`.

    **End to end, through the product's own server, before and after** (same day, same guest
    state). The server built from this change, and the guest's own staged server of 2026-09-27 (the
    name rule it replaces), were driven with the same read-only calls - `outlook_health`,
    `list_accounts`, `search` - over raw stdio from the console session at the Outlook's own level,
    the way `guest/Invoke-GuestMeasure.ps1` drives it. Every outcome was written down before the run.
    - **On `CorpusProfile`** (the corpus store, `q99lz@vm.invalid($ce9d6e4)` and
      `q99 50%25 off%2A%3Fx($5159380d)`): the new server tied all three by `storeHash` / `entryId` to
      exactly those segments and listed the tier profile's two stores under `storesNotInProfile`. The
      old one called the two empty stores "the local index holds nothing for", with advice to
      add them to Indexing Options, listed them `onlineOnly` in `list_accounts`, and answered a search
      scoped to either with `storeNotIndexed: true` - all three false: both stores were indexed, they
      just held no mail. Results for the corpus store, scoped or not, were identical.
    - **The decoy, on `OutlookAI-Tier`**, after a new empty PST named `Outlook Data File` was attached
      there (filed as `Outlook Data File($580470ed)`, the hash predicted from its path): a search
      scoped to `Outlook Data File` was answered by the old server with three hits from the CORPUS
      store of `CorpusProfile` - another profile's mail, under this profile's store name, as
      `freshness: "live"` - and its `outlook_health` gave that store the corpus's frontier. The new
      server searched the decoy's own root (no hits; no mail frontier, so the widest sweep window and
      `degraded: true`) and listed `Outlook Data File($23a27f0d)` under `storesNotInProfile`.
    - **Unchanged, and still open**: an UNSCOPED search is not scoped by store, so on both servers it
      returned the corpus hits too - another profile's mail, under a name this profile also uses, and
      not openable from this profile. Whether to drop or flag such hits is open (`TODO.md`); doing
      either from the map alone would also catch an Exchange store the hash did not decide.

25. **ANSWERED 2026-10-03, and its proof REMOVED the same day - where Outlook registers the Drafts
    folder it creates in a data file with no Inbox.** On the store's TRUE root folder - the parent of
    the IPM subtree, reached through the top folder's `PR_PARENT_ENTRYID` - and nowhere the object model
    hands out: not the store object, not the top folder, and there is no Inbox (bisect run E4b and run
    13, section 4.1e; F10). The non-creating lookup reads it there since `5ac1d85`, which is what lets
    `discard_draft` and `update_draft` accept a draft a reply filed in such a file. The measurement came
    from the created-folder proof (`T2/LiveCreatedFolderTests`, its throwaway data file and
    `Reset-ThrowawayStore.ps1`), which ran green on that guest and was then removed with the
    requirement it proved: the maintainer dropped Q85's "may create, must report" for the draft tools
    and `discard_draft` (`Docs/overnight-review-2026-10-03.md`, D120-D125). The true-root reading keeps
    its T1 pins (`T1/ReadOnlyFolderLookupTests`); nothing on a guest re-measures it any more.
26. **MEASURED 2026-10-03 (the Q99 folder finding) - how the index spells a FOLDER name holding
    `% / \ * ?`, and what a folder-scoped search does with it.** Microsoft documents those five as
    percent-encoded "if they are in the store or folder display name" (*About MAPI URLs for
    Notification-Based Indexing*), and item 24 measured it for a store; for a folder it had never
    been measured, while the product built a folder scope from the RAW name. On `OutlookAI-Indexed`
    from `CP-17C-CORPUS-160K`, the 30-and-60 settings staged again after the restore and
    `corpus-verify --window 30 --window 60` OK, the suite staged from `b0d35a1`
    (`Testbed/README.md` step 8b, `TEST-READY`, 3,627 tests), Outlook started NOT elevated on
    `OutlookAI-Tier` by `Start-OutlookUnelevated.ps1`, and `T2/LiveFolderNameEncodingTests` run alone
    through `Register-InteractiveTask.ps1 -RunLevel Limited` with the opt-in: **1 passed, 1 min
    10 s**, the tripwire's post-run census 0 failures, the hub reconciled, zero tagged artifacts and
    zero test folders. Every write in the hub, through `move_mail` with `create_folder` or, for the
    one name it cannot make, `LiveOutlookTestMailer.FileTaggedItemInNewTestFolder`; the index read
    only through `SELECT` statements. Then `CP-17C` restored and the 30-and-60 settings staged again
    (`corpus-verify` OK). Raw output: `.work\q99-folders\` of that worktree. Findings:

    - **Outlook accepts all five in a folder name** (a PST, Office LTSC 2024): `move_mail` made
      `OutlookAI-McpTest-Folder 50% off`, `... star*`, `... why?`, `... back\slash`,
      `... 100%*? mix`, `... %2A not a star` and, inside the first, `... inner*?`; `Folders.Add`
      made `... a/b` and kept the name exactly. `list_folders` lists each under its name - the last
      as `OutlookAI-McpTest-Folder-Enc/OutlookAI-McpTest-Folder a/b`, which reads like a nested path.
    - **The URL encodes each exactly as documented**: `50%25 off`, `star%2A`, `why%3F`,
      `back%5Cslash`, `a%2Fb`, `100%25%2A%3F mix`, `%252A not a star`, and the nested one
      `.../OutlookAI-McpTest-Folder 50%25 off/OutlookAI-McpTest-Folder inner%2A%3F`. Every item was
      in the index 6 s after it was filed.
    - **The display paths do not encode**: `System.ItemFolderPathDisplay` reads
      `/tier@vm.invalid/OutlookAI-McpTest-Folder-Enc/OutlookAI-McpTest-Folder 50% off`, and
      `System.ItemPathDisplay` the same with the subject after it - names, for every one of the five.
    - **The scope the product used to build addressed nothing**: the raw-name `SCOPE`, alone or with
      its folder-path equality, matched **0** rows for every name. The encoded `SCOPE` matched the
      item, alone and with the display path spelled as names; with the display path spelled as the
      URL spells it, 0.
    - **Through the product, with the fix** (an index-only search, so the sweep could not cover for
      the index): every name found by a folder search, recursive and not, reported under its real
      folder name, `read` locating it by the URL's own folder path (`urlSegments`), and the zero-row
      guard silent; the nested item found from its parent with subfolders and not without. The `/`
      name is found by its parent's recursive search (folder
      `OutlookAI-McpTest-Folder-Enc/OutlookAI-McpTest-Folder a/b`, `urlSegments`) but cannot be
      named in `folder`, which splits on `/`: asked for that way the search finds nothing and the
      guard says the path "matched NOTHING" - a limit of the product's path syntax, not of the index
      (`TODO.md`).
    - **Two oddities in the folders' OWN index rows**, which no search returns: the `back\slash`
      folder's `System.ItemNameDisplay` is `slash`, and the `a/b` folder's
      `System.ItemFolderPathDisplay` is `.../OutlookAI-McpTest-Folder a` - the index splits a display
      path on both separators. The item rows inside both folders read right.

    **What the product does with it** (`Mapi/MapiUrlSegment`, `McpServer/README.md` load-bearing
    fact 17): a folder scope spells every segment the way the index does, a delegate's
    `/1/<name>` likewise, and every URL the product reads back - a hit's folder segments, its store
    name, a display path derived from a URL - is decoded to names, one pass each way. A name holding
    none of the five builds byte for byte the scope it always did. Not measured, and not measurable
    on a guest as it stands: a STORE name in a display path (no store here has one of the five;
    decoded there by the folder evidence), a delegate's name (Exchange), and an attachment's file
    name in an `/at=` URL, which the product does not decode.
31. **MEASURED 2026-10-03 (Q114/Q115) - what a FOLDER's id is, what it survives, and what the index
    keeps of it.** (27 to 30 are the Q99 follow-up's items, on its own branch.) The maintainer decided
    that tools address a folder only by a unique id, its name kept for display (Q114), and asked
    whether a folder-scoped search can be answered correctly from the index alone while Outlook is
    down (Q115). Measured on both guests, Office LTSC 2024 (16.0.17932), every Outlook start NOT
    elevated: `T2/LiveFolderIdentityTests` through `Invoke-LiveTierOnGuest.ps1` with
    `-FilterSuffix '&FullyQualifiedName~LiveFolderIdentityTests' -SkipHubReset` - on
    `OutlookAI-Unindexed` from and back to `CP-13B-LIVE-GREEN` (1 of 1, at `a5ce5cb`), on
    `OutlookAI-Indexed` 3 of 3 at `a5ce5cb` from and back to `CP-17C-CORPUS-160K`, and 3 of 3 again at
    `b7673d1` and `e906f85` from and back to `CP-18C-ALL-KINDS`, its resting checkpoint by then, the
    settings staged each time; and a read-only COM probe of every store's ids on `OutlookAI-Unindexed` around two
    graceful restarts (`Restart-Guest.ps1`) and a byte copy of a scratch PST attached with
    `Add-OutlookPstStore.ps1`, the guest restored to `CP-13B` afterwards (raw output: `.work\q114\` of
    that worktree). Every write in the hub, through `move_mail` with `create_folder` and three new
    tested helpers (`LiveOutlookTestMailer.RenameTestFolder`, `MoveTestFolder`, `SoftDeleteTestFolder`,
    which refuse anything that is not a test folder); the index read only through `SELECT`. Findings:

    - **A PST folder's id** is `Folder.EntryID` = its `PR_ENTRYID`, 24 bytes: four zero flag bytes (a
      long-term id), the store's `PR_RECORD_KEY` (16) and the folder's node id (4, little-endian,
      type 0x02) - [MS-PST] 2.4.3.2. It survived a rename, a move within the store, a soft delete
      (`Folder.Delete`, into Deleted Items - the old id then opens the folder THERE) and two graceful
      restarts (all 37 ids of 4 stores, their `StoreID`s and record keys identical); an item inside
      kept its own id through its folder's rename and move. A folder deleted and made again under the
      same name and parent got a new id ([MS-PST] 2.2.2.6: node ids come from a per-type counter).
      `GetFolderFromID` opened every id WITHOUT a store id - nine stores on the two guests, the default
      among them - and from lower-case hex. The note in `LiveOutlookTestMailer.RemoveEmptyTestFolder`
      that a soft delete gives a NEW EntryID was wrong for a PST, and is corrected.
    - **A node id is not unique across stores.** Every PST gives its default folders the same node
      ids - root 0x8022, Deleted Items 0x8062, Inbox 0x8082, Outbox 0x80A2, Sent Items 0x80C2 - so only
      the store UID tells two PSTs' Inboxes apart, and one hex digit separates two sibling folders'
      ids (Inbox `...82800000`, Outbox `...A2800000`).
    - **The other candidates fall away.** A PST folder's `PR_RECORD_KEY` is its 4-byte node id
      (store-scoped, as MAPI documents for folders); `PR_SOURCE_KEY`, `PR_PARENT_SOURCE_KEY` and
      `PR_LONGTERM_ENTRYID_FROM_TABLE` are absent on a PST folder object (`0x8004010F`); `StoreID` is
      the store's `PR_ENTRYID`, 114 to 172 bytes for these Unicode PSTs because it holds the file path.
    - **A byte copy of a PST**, attached to the profile while its original was open, was RE-KEYED by
      Outlook: a new record key, so every folder id of the copy differs from the original's in the UID
      part while the node ids stay equal; the original kept its key, and both keys survived a further
      restart.
    - **The index keeps no folder's whole id, but every row's own node id.** `System.ProviderItemID`
      is `N` and the row's node id in ten decimal digits on every folder and item row compared
      (`Inbox` `N0000032898` = 0x8082; an item row its own message node id, never its folder's) - the
      "provider item ID" a store pushes with each MAPI URL (*About MAPI URLs for Notification-Based
      Indexing*: "send only the provider item ID for folders"). A folder's own row (`System.ItemType`
      `MAPI/Folder`) has the names path as its URL and no id segment, as documented; an item row names
      its folder only by that path. Every one of the 1,688 property descriptions the property system
      names was asked of each row (any the index refused as a column was dropped): none holds the
      24-byte folder id in hex, base64 or the URL encoding.
    - **So a PST folder id finds its index scope with no Outlook call**: the one root whose item URLs
      carry the id's store UID, then the one row under it whose `System.ProviderItemID` is the id's
      node id. Seven folder ids of the three indexed PSTs - six of up to 172 items and Corpus A's
      largest, 88,037 items in a 160,000-item store - mapped exactly to the URL built from their names,
      in 59 to 114 ms each (422 ms for the first, cold). A one-item folder's item row moved to the new
      path 6 s after a rename and 10 s after a move (first run); in the second run the item row and the
      folder's own row - found at its new URL by the same `System.ProviderItemID` - moved together, in
      10 s and 8 s. Each row moved at once: never at both paths, never at neither, at a 2 s poll.

    Not measured, and not measurable on these guests: Exchange - 46-byte ids ([MS-OXCDATA] 2.2.4.1:
    the mailbox GUID and the folder's FID), whose `System.ProviderItemID` in an OST's rows is unknown
    and, if it is an OST node id, is not held in that id; IMAP and Outlook.com stores; a move to
    another store (documented: a new id); export and import; a mailbox seen through two stores. The
    test is written to run as it stands on the Exchange guest.

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
* **The VM bucket does not prove the delegate-store or cached-Exchange paths at all**, and no test
  machine can: `Requires=DelegateStore` needs a mailbox somebody else owns, and `Requires=CachedExchange`
  an Exchange server. Seven tests, named by the production-only filter in section 5. **Corrected
  2026-10-03:** the Exchange test VM can - it has an Exchange server (section 4.4); the delegate half
  waits for the shared test mailbox (Q109).
* **The guest's SDK is PINNED to whatever the host was running when the payload was staged**,
  and nothing enforces that they stay equal. 10.0.401 was chosen for sameness rather than for any
  requirement - no `global.json` exists - so the two can drift the
  moment the host updates, and the first symptom would be a guest measurement that differs from a
  host one for a reason nobody is looking for. `Testbed/MEDIA.md` records the pinned version; it is
  the thing to check when host and guest disagree about something that should not depend on the
  toolchain.

* **The VM bucket has run end to end on ONE guest, `OutlookAI-Unindexed`, on 2026-10-03 (section
  4.1e) - the unindexed filter, 80 or 81 tests a run - and is GREEN there: 80 of 80 in runs 18 and 19,
  checkpoint `CP-13B-LIVE-GREEN`.** Green includes two decisions taken on the maintainer's behalf, each
  in `QUESTIONS.md`: a renamed derived draft's ConversationId is held to Exchange's promise on Exchange
  and to the kept topic's hash elsewhere (F8), and on Office LTSC 2024 the user's close of the last
  window may end an Outlook OutlookAI started (F9). The count moved from 31 to 121 by re-reading what
  each test needs method by method - no test was changed to make it fit.
* **And on `OutlookAI-Indexed`, the index tier included, it is GREEN too: 122 of 122 on the merged
  master (`dc1b5c5`, from `CP-18C-ALL-KINDS`), checkpoint `CP-19C-LIVE-GREEN` (section 4.2f)** -
  after one product fix (a folder's own index row was being returned as a search hit) and four tests
  brought up to the indexed guest. One of its five full runs crashed Outlook inside Word (`wwlib.dll`)
  during a signature-override draft and left no dump; that crash is open in `TODO.md`.

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
  therefore **retired**, and the `TODO.md` preflight item was **closed** (and has since been removed from `TODO.md`) - both of its
  justifications were measured false. Rebuilds now trigger on `corpus-verify` refusing and on a
  release, neither of which is a calendar.

* **The clock that DID bite was the corpus, and this is the one to respect.** A three-week
  absence left it 27 days stale with its 1-day and 7-day windows selecting nothing, while the
  Office clock everyone was watching cost nothing at all. `corpus-verify` catches it fail-closed.
  The general lesson is worth more than the instance: **the schedule was attached to the visible
  clock rather than the harmful one.**
