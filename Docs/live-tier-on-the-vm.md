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

Windows Search does not index Outlook per data file. It indexes a MAPI scope, and that scope
is expressed as **one URL per Windows user account**, of the form `mapi16://{SID}/`, covering
that account's whole Outlook profile. There is no per-store URL underneath it, so Indexing
Options offers exactly one switch per account: the profile is indexed, or it is not. Two
stores in one profile are therefore indexed together or excluded together.

So the split is per **Windows account**: one account whose Outlook profile is indexed and one
whose profile is not. That is the whole reason this machine has two logons.

> This is the load-bearing assumption of the layout and it is **derived from how the scope is
> addressed, not measured on this VM**. Verify it before building anything else: add a store,
> let the indexer settle, and read `outlook_health`'s `index.perStore[]` on both accounts.
> Section 8 says what to do if it turns out to be wrong.

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

### 1.4 A dummy account with a LOCAL SINK, because an unroutable one poisons the teardown

The account exists because `NewDraft` resolves an `Account` object by SMTP address and refuses
when none matches, which is what puts the entire draft, update/discard, HTML-draft and send
families out of reach of an account-less machine.

Pointing it at an unroutable server was the first plan and it is wrong. A send **queues** and
never leaves, the Outbox is in the mandatory zero-artifact sweep, and so every run that sent
anything would fail its own teardown forever on residue nothing could remove. Six live methods
also need the mail to genuinely arrive.

So the account points at a **local sink that delivers back**: submissions on loopback, and the
same messages served back over POP3 to the same profile, so self-addressed mail round-trips
into the Inbox. Section 2.7 says which sink and how to install it.

---

## 2. Building the machine from nothing

Do these in order. Checkpoint where the section says to; a checkpoint is much cheaper than
redoing the step above it.

### 2.1 The hypervisor and the guest

* Hyper-V guest, named `OutlookAI-TestVM` by convention. Generation, firmware, vCPU, RAM and
  disk size are **not recorded anywhere and are yours to choose**; see section 8.
* Windows 11, edition and build unrecorded. Whatever you choose, record it beside the VM: an
  Outlook build difference is the first thing to suspect when a live test behaves differently
  here than on the maintainer's machine.
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

* .NET SDK (version unrecorded; it must build `net10.0-windows` and `net48`), git, and a clone
  of this repository.
* Build once so the server exe exists where the tier-3 tests look for it. That path is baked
  into the test assembly at build time as `AssemblyMetadata("McpServerExePath")` and points at
  `McpServer\OutlookAI.McpServer\bin\<Config>\net10.0-windows\OutlookAI.McpServer.exe`.
* Install the add-in and let it run once. Tests carrying `Requires=AddInRegistry` read tuning
  state the add-in writes on first run; without it they have nothing to read.

**Checkpoint `CP-03-OUTLOOKAI-INSTALLED`, then `CP-05-ADDIN-TRUSTED`.**

### 2.4 The two Windows accounts

Create two local accounts. Section 1.1 says why. Suggested roles, since neither is recorded:

* an **indexed** account, whose Outlook profile carries Corpus A and the dummy account, and
  where the index tier and the send path run;
* an **unindexed** account, whose Outlook profile carries Corpus B, with its `mapi16://{SID}/`
  scope removed from Indexing Options.

Both accounts need the repository, the SDK and a built server exe, or the tier can only run
under one of them. Whether that is a clone each or one clone with both accounts granted access
is your call; record which.

**Verify the split before going further.** On each account, open Indexing Options, confirm the
Outlook entry is present or absent as intended, let the indexer settle, then read
`outlook_health` and check `index.perStore[]`. Establish it; do not assume it.

### 2.5 The Outlook profiles

Two profiles are needed on the account that does corpus work:

* a **corpus profile with no mail accounts at all** (section 1.2), and
* a **tier profile** with the dummy account.

Set Outlook to "always use this profile" and switch by changing that setting, not by prompting:
a prompting profile cannot be driven over COM. How the profiles are created and switched is
**not recorded**; the Mail control panel works and is the obvious route.

### 2.6 The stores

Add every store **through Outlook itself** (File > Account Settings > Data Files > Add). Do not
improvise a script. Creating stores is not something the tested helpers do, and mailbox
mutation from ad-hoc shell code is the thing that once destroyed real mail.

Naming matters more than it looks:

* **The hub store must be named exactly the dummy account's SMTP address.** Several tests use
  `testHubStoreDisplayName` as an address (`NewDraft(Hub, Hub, ...)`, `FindAccountBySmtp(Hub)`),
  so the hub PST has to be called something like `test@vm.invalid` literally. Whether Outlook
  accepts `@` in a store display name is **untested and it gates the whole draft family** - try
  it first, it costs five minutes. `.invalid` is guaranteed unresolvable by RFC 2606, so a
  misconfiguration cannot leak mail anywhere.
* **The dummy account's delivery store must be a separate throwaway PST**, not a corpus store.
  An account delivering into the corpus store can flip that store's `IsDataFileStore`, and
  `CorpusSafety` reads that property as one of four independent facts it requires before it
  will write anything. Get it wrong and the generator refuses that store permanently.
* The bystander is listed in `expectedStoreDisplayNames` and is never the hub.

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

Repeat for Corpus B under the other Windows account, with **a different `--corpus-id` and a
different manifest path**. Whether the two corpora should share a seed and anchor is not
settled; sharing them makes the two stores directly comparable, which is probably what you
want.

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
| `expectedDelegateStoreDisplayNames` | Delegate/shared mailboxes. Watched, never written, folder hierarchy allowed to come and go. Empty here. | no |
| `probeTerm` | A word proven to hit this machine's search index. | Production only |
| `subjectOnlyProbe` | Coordinates of a population whose term is in the subject and not the body. Four fields, all or none. | Production only |
| `delegateNestedFolderProbe` | A delegate folder Outlook nests and the index publishes flat. | never |
| `corpus` | Where the measurement corpus is and what it was generated from, so the tier can prove it is still measurable. Six fields plus optional `windowDays`. | no, all or none |
| `mailSink` | Loopback submission and retrieval endpoints. **Absent means this machine has real transport.** | no, all or none |

A block that is present must be **complete**: three fields out of four reads as configured and
behaves as absent, which is the exact silence these checks exist to remove.

`windowDays` is how the machine declares which measurement windows it actually asks about. Left
empty it means all of them, including the one-day window - which forces a re-anchor every day.
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

**The repair.** `corpus-reanchor` shifts every item's received and submit instants forward.

```
dotnet run --project <as above> -- corpus-reanchor \
  --store "Corpus A" --allow-store "Corpus A" \
  --corpus-id vm1 --seed 4242 --anchor 2026-08-01 --count 40000 \
  --manifest D:\corpus\vm1.jsonl --to now --execute
```

It runs under the **no-accounts profile**, like every other corpus verb. It never creates,
moves or removes an item; it opens each one by the EntryID the manifest records and writes two
date properties, and it touches an item only when the EntryID is in the manifest, the subject
still carries both tags, and the ordinal in that subject is the one being addressed.

The target is ABSOLUTE, not incremental, so running it twice is a no-op and an interrupted run
is finished by running it again. The manifest's anchor is deliberately not rewritten - it is
half the corpus's identity, and every later `--anchor` argument depends on it - so the shift is
derived from the item lines rather than recorded in the header, and the re-anchor appends a
replacement line per item. Expect the manifest to roughly double in size per re-anchor.

**Re-anchor after every checkpoint restore.** A restored checkpoint puts the corpus back where
it was on the day it was taken, which is by definition older than today.

**Re-anchoring changes every item, so the index will re-crawl Corpus A.** Let it settle before
taking an index measurement.

---

## 4. Running the tier

```
dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
  --filter "Category=Live&Requires!=DelegateStore"
```

That filter IS the VM bucket, spelled out: everything live except the tests naming a capability
this machine cannot be given. There is no separate "which bucket" trait to keep in step with it -
see section 5.

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

**The vocabulary, all ten values.** Nine of them this VM can be given; one it cannot.

| Capability | What the machine must have |
| --- | --- |
| `OutlookInstance` | an Outlook to attach to, and nothing more specific. The floor: a live test that needs nothing else says this rather than saying nothing |
| `InteractiveDesktop` | a real desktop session - Outlook windows and screenshots cannot be driven from session 0. Declared only by tests that PUT SOMETHING ON SCREEN |
| `AddInRegistry` | the add-in installed and run once, so its tuning values exist |
| `SearchIndex` | a populated Windows Search index - Corpus A |
| `MailAccount` | a mail account rather than a bare PST - the dummy account |
| `Transport` | mail that actually goes out and comes back - the local sink |
| `MultipleStores` | more than one store mounted - all three |
| `SmallHubStore` | a hub small enough that a paging assertion means something |
| `ProbePopulation` | the hand-curated population named in the settings file |
| **`DelegateStore`** | **a delegate/shared mailbox. The one capability no test machine can be given** |

`.github/scripts/check-pinned-constants.ps1` fails the build if any of those ten names stops
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

---

## 8. What a rebuilder still has to establish

Everything in this list is genuinely unrecorded or unverified. It is a deliverable in its own
right: the point of naming it is that a rebuilder should not have to discover it is missing.

**Verify before building anything else**

1. **That one Windows account's Outlook profile can be excluded from the index while another's
   is not** (section 1.1). The whole layout rests on it and it is derived rather than measured.
   If it turns out to be false, the fallback is two VMs, and the store layout collapses to one
   corpus per machine.
2. **That Outlook accepts `@` in a store display name** (section 2.6). It gates the draft
   family and costs five minutes.
3. **Whether smtp4dev's POP3 side maps an arbitrary `USER` to the catch-all mailbox**, or
   whether the username must match a configured mailbox name. If the latter, add an explicit
   `Mailboxes` entry with `Recipients: "*"` and use its name as the POP3 username.
4. **That the corpus store's `IsDataFileStore` stays true once the profile has an account**
   (section 2.6). If it flips, the generator is locked out of that store permanently.

**Not recorded anywhere**

5. Hyper-V generation, Secure Boot, TPM, vCPU, RAM, disk size, checkpoint type (production
   checkpoints use VSS and behave differently with Outlook mid-run).
6. Windows edition, build, ISO, licensing, computer name, Defender exclusions (an indexer, a
   400 MB PST and real-time AV interact), power and sleep policy.
7. Office version, channel, bitness, install method, and how the first-run wizard is suppressed.
8. The two Windows account names and their roles; whether both need a clone and an SDK; whether
   checkpoints must be taken with both logged on.
9. Outlook profile names, how they are created, which is default, and how the switch between the
   no-accounts profile and the tier profile is automated.
10. The scheduled-task recipe for session 1: task name, principal, working directory, argument
    line, output redirection and exit-code capture. Only "`-LogonType Interactive`" is recorded,
    and an elevated process's stdout cannot reach the caller, so output must go to a file.
11. The exact PST file paths and names for all four stores, and the mapping from file name to
    display name.
12. .NET SDK version, clone path, build configuration, and how the built server exe reaches the
    path the tier-3 tests expect.
13. How results, screenshots and logs get out of the guest, and where `ScreenCapture` writes.
14. `Docs/v3-probes/soakfix13-probe-sweep-cost.ps1` is gitignored and is required by step 2 of
    the measurement plan. It has to be copied in by hand.
15. **The parameters of the corpus that is currently on the VM.** Only an example
    (`vm1 / 4242 / 2026-08-01 / 40000`) is written down, while the corpus every published
    measurement rests on is a 20,000-item one. Record the real parameters beside the manifest.
16. Whether Corpus A and Corpus B should share a seed and anchor, and how their manifests are
    named apart.

**Known gaps in what the corpus contains**

17. The generator writes `IPM.Note` with a subject, a body, a read state, message flags and two
    date properties. **No senders, no recipients, no attachments, no HTML, no categories, no
    flags, no subfolders, no other message classes.** A test needing any of those is in the VM
    bucket and will still fail here - because of the corpus, not because of the machine, and
    nothing in the traits says so. Widening the generator is queued work.

**Open behaviour**

18. The tripwire's re-census-then-re-run policy on a suspected loss is decided and not built.
19. `machineProfile: "Portable"` turns "found nothing to test" from a failure into a pass. There
    is no check on how many assertions actually fired, so a run that proved nothing looks like a
    run that passed.
20. Non-hub stores in `expectedStoreDisplayNames` are granted draft-create and draft-delete by
    `StoreWriteAllowlist`. That was harmless while the identity tests were all pinned to the dev
    machine; now that 121 of 127 select onto this one, the bystander - the one store the tripwire
    needs untouched - is inside that grant.

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
* **Nobody has yet run the VM bucket end to end anywhere.** The 121 read as runnable there; that
  is not the same as having run there. The count moved from 31 to 121 by re-reading what each
  test needs method by method - no test was changed to make it fit.
