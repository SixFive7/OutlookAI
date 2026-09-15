# Rebuilding the live-tier test VM

**Start here if the test machine is gone.**

This directory holds the runnable half of the testbed: parameter sets, scripts and templates
that a person with this git repository, a Windows box and an Office licence needs in order to
rebuild the machine the `Category=Live` tests run on. The reasoning half - why the machine is
shaped this way, what each guard does, what to read in a run's output - is
`Docs/live-tier-on-the-vm.md`. Read that alongside this; neither is complete on its own.

**Why this exists.** Three times in one week, material this project depended on was kept only
in a scratch directory on one machine, and that directory was cleared without warning. What was
in it was not notes: it was the corpus manifest that makes a 20,000-item synthetic corpus
verifiable and removable, the scripts that built the guest, and the parameters that every
published measurement is a statement about. The manifest was recovered on 2026-08-24 from the
guest itself, by luck rather than by design. **Nothing in here should exist only on a machine.**

---

## 0. Media is a precondition — read `MEDIA.md` first

You cannot start without Windows installation media, and this machine did not have any when the
question was first asked. `Testbed/MEDIA.md` records what is needed, **where on this machine it
actually is** (`.work/media/` for Windows, `.work/office-odt/` for Office — gitignored scratch,
and the right home precisely because Downloads has been purged without warning once already), the
Office deployment settings, and the licence clocks — including the correction that matters most:
**Office's 30-day grace turned out to cost nothing.** Past grace the guest keeps working — every
COM read succeeds and a cold COM start completes in 3.7 s with no dialog — so the monthly rebuild
cadence it once justified was **retired on 2026-09-15**. Rebuilds now trigger on `corpus-verify`
refusing and before a release. The clock that actually bit was the **corpus**, not the licence.

Both are **preconditions, not artefacts**: nothing here regenerates them, `.work/` is scratch, and
a rebuilder who finds either directory empty must re-stage it before step 4.

It also carries the rule that came out of nearly getting this wrong: **never destroy a working
testbed before its replacement runs.**

## 1. The order to do things in

| # | Step | What runs it | Where |
| --- | --- | --- | --- |
| 1 | Build the answer volume | `host/New-AnswerFile.ps1` | host |
| 2 | Create the guest, attach both ISOs, boot it | `host/New-TestbedVm.ps1` | host |
| 3 | Windows installs itself - edition, disk, account, autologon, locale, power | nobody: the answer file | guest, unattended |
| 4 | Install Office, the accounts and the profiles | Office and the Windows accounts by hand - `Docs/live-tier-on-the-vm.md` §2.2-2.4, with `.work/office-odt/Testbed.xml` (`MEDIA.md`). **Profiles and PSTs now have drafted scripts - §4b** | guest |
| 5 | Give yourself a way to reach session 1 | `guest/Register-InteractiveTask.ps1` | guest, once |
| 6 | Build the server and the tools, and copy them in | `host/Publish-GuestPayload.ps1` | host |
| 7 | Install the mail sink, the dummy account and the identity account | Sink: `guest/Install-MailSink.ps1`, from a staged package (§6 item 11 - and read it first, §2.7 of the runbook is wrong in four places). Accounts: `guest/New-TierProfile.ps1`, then the POP3 password **by hand, once**, because the sink refuses an empty one. The signature has a script | guest |
| 7b | **Decide whether this guest is the indexed one or the unindexed one, and do it BEFORE the corpus exists** | `guest/Set-OutlookIndexingDisabled.ps1` on `OutlookAI-Unindexed`; nothing on `OutlookAI-Indexed`. Order is the whole point: exclude Outlook first and no row is ever crawled, so there is nothing to purge and nothing to wait for | guest |
| 8 | Build the corpus | `guest/Build-Corpus.ps1` | guest, session 1 |
| 9 | Write the live-test settings file | copy `live-test-settings.example.json` - and read §3b, or the tier refuses to start | host or guest |
| 10 | Take the measurements | `guest/Invoke-GuestMeasure.ps1`, `guest/Measure-SweepCost.ps1` | guest, session 1 |
| 11 | Get the results out | `host/Copy-FromGuest.ps1` | host |

Every script takes `-WhatIf`-style caution seriously: the ones that write take an explicit
`-Execute`, and print what they would do without it.

---

## 1b. Steps 1 to 3 need nobody: the Windows install is unattended

**Two commands and a wait.** Per guest:

    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName OutlookAI-Indexed
    pwsh -File Testbed/host/New-TestbedVm.ps1  -Name OutlookAI-Indexed `
        -IsoPath      .work\media\Win11_25H2_EnglishInternational_x64_v2.iso `
        -AnswerIsoPath .work\testbed-answer\OutlookAI-Indexed\OutlookAI-Indexed-unattend.iso `
        -Execute -Start

The first builds a small ISO carrying `autounattend.xml` and the first-logon script. The second
creates a Generation 2 VM with Secure Boot and a vTPM, attaches the Windows ISO and the answer
volume as two DVD drives, and boots it. Nothing else is typed. Do it twice, with the two names,
and both guests come up identically - which is the point: a difference between the indexed and
unindexed guests should come from the store layout, never from someone having answered an
installer differently on a Tuesday.

**What the answer file sets, and why those values.** The guests must match the maintainer's own
machine, because that configuration is where most of the userbase sits and is therefore what the
live tier should be testing against. The measured host survey is in `MEDIA.md`; the short version:

| Setting | Value | Set by |
| --- | --- | --- |
| Edition | Windows 11 Pro, from the multi-edition ISO, no product key | answer file, `/IMAGE/NAME` |
| Disk | GPT: 260 MB EFI, 16 MB MSR, Windows fills the rest, disk 0 wiped | answer file |
| Setup UI and display language | en-GB | answer file |
| System locale (non-Unicode) | en-US | answer file, re-asserted at first logon |
| User locale / formats | nl-NL - so `25-8-2026` and `4.000,50` | answer file, re-asserted at first logon |
| Time zone | `W. Europe Standard Time` | answer file |
| Keyboard | `00020409`, United States-International, on every language | answer file, forced at first logon |
| Preferred languages | **en-NL then nl-NL** | first-logon script only - see below |
| Home location | Netherlands, GeoId 176 | first-logon script |
| Account | one local administrator, from the gitignored credential | answer file |
| Autologon | on, 999 logons | answer file |
| Sleep, hibernate, fast startup, screen saver | all off | first-logon script |

**The one thing an answer file cannot say.** `en-NL` (English, Netherlands) is a *transient*
language: Windows allocates its LCID out of the `0x2000` block at runtime, so there is no
constant to write down and no way to name it in XML. `guest/Complete-FirstLogon.ps1` therefore
sets the final language list with `Set-WinUserLanguageList` after the account exists, and rewrites
each entry's keyboard to `00020409` while keeping whatever LCID Windows assigned. It logs
everything it did, and then reads every setting back, to `C:\Windows\Setup\first-logon.log` on the
guest. **Read that file before trusting a guest** - diff it against the table in `MEDIA.md`.

**No Windows 11 requirement is bypassed.** There is deliberately no `LabConfig`,
`BypassTPMCheck`, `BypassSecureBootCheck` or `BypassRAMCheck` anywhere. A Generation 2 Hyper-V VM
with a vTPM and Secure Boot on satisfies Windows 11 natively, and a guest built by switching those
checks off is not the machine the userbase runs.

**What the generator needs.** A gitignored `vm-credentials.json` (§4 - the same file everything
else uses, loaded through the same `host/Get-GuestCredential.ps1`), and something to build an ISO
with. **The credential's `vmName` is now empty, and that is a setting rather than an omission** -
see §4a. The loader refuses a credential whose `vmName` names a *different* VM, so a pinned file
would refuse two of the three guests; the check only fires when the field has a value, and with
one shared account there is no wrong machine for it to protect against. It prefers `oscdimg.exe` from the Windows ADK's Deployment Tools and falls back to the
IMAPI2FS COM object that ships with Windows, so a machine with no ADK can still build the volume.
If neither works it says which is missing and writes nothing; it never leaves half an ISO, because
a broken answer volume looks exactly like an ordinary interactive Setup and tells you nothing.

**Rebuilding after a change.** Edit `guest/autounattend.template.xml` or
`guest/Complete-FirstLogon.ps1`, re-run `host/New-AnswerFile.ps1`, and create the VM again. The
generated ISO is disposable; the template is the record.

**The credential never lands in this repository.** The committed template carries placeholder
tokens where the password goes. The generator substitutes them and writes only into gitignored
`.work/`; it refuses an output path under `Testbed/`, and refuses any path inside the repository
that is not under `.work/`. `.github/scripts/check-testbed-references.ps1` check 7 fails the build
if the template ever stops holding placeholders, or if a filled `autounattend.xml` is ever
tracked. **The generated ISO holds the password in clear text** - it is in scratch, keep it there,
and delete it once the guest is built.

**One keystroke this cannot avoid, and what is done about it.** Microsoft's retail ISO boots
through a loader that prints "Press any key to boot from CD or DVD" and gives up after a few
seconds; that prompt is inside the ISO, so no VM setting removes it. `-Start` types at the guest's
synthetic keyboard over WMI (`Msvm_Keyboard`) for the first few seconds - no window, no focus
change. If that route is blocked on a host, press a key in the console once; everything after it
is unattended either way.

---

## 2. Two facts that govern everything else

**The guest has a .NET runtime and no SDK.** Nothing can be compiled there. Every binary the
testbed runs is published on the host and copied in - that is what `host/Publish-GuestPayload.ps1`
is for, and it is why the guest's working directory holds `McpServer.zip` and `Tools.zip`.
(Verified on the guest 2026-08-24: runtime 10.0.10, no SDK.)

**PowerShell Direct lands in session 0, and Outlook can never finish starting there.** Anything
that touches COM - every corpus verb, the MCP server, the live suite - has to run in the
interactive console session. The route is a scheduled task registered with
`LogonType=Interactive`, which lands in session 1; `guest/Register-InteractiveTask.ps1` is that
recipe, and it is the single piece of knowledge that was hardest to reconstruct. Autologon is
enabled on the guest (`AutoAdminLogon=1`, user `vmadmin`), which is what keeps a console session
alive for the task to land in - and on a guest built from `guest/autounattend.template.xml` that
comes out of the answer file rather than from someone remembering to switch it on.

A corollary worth stating because it costs an afternoon otherwise: **an elevated or
scheduled-task process's stdout cannot reach the caller.** Output goes to a file, and the caller
polls the file. Every script here does that.

---

## 3. The corpus parameters, and why they are committed

`testbed.json` carries four values that matter more than anything else in this directory:

```
corpusId  vm2      seed  7777      anchor  2026-08-19      itemCount  20000
```

Those four, plus the generator's default shape, deterministically reproduce the corpus that
**every published sweep and frame measurement in this repository is a statement about** - the
~12 s-per-store sweep behind `SweepBudgetMs`, the 10,734,599-byte frame high-water behind
`SweepBodyBytesBudget`, and the seven-day window those numbers were taken over - which this
corpus fills with 1,612 items across four folders, enough that the 200-per-folder cap engages
and the sweep actually reads 758 of them.

They were not written down anywhere until 2026-08-24. `Docs/live-tier-on-the-vm.md` and
`Docs/corpus-measurement-plan.md` both used `vm1 / 4242 / 2026-08-01 / 40000` as a worked
example, and a reader had no way to know the real corpus differed in all four.

**How they were established, so the claim can be checked rather than believed.** The header line
of the recovered manifest carries them, and `corpus-plan` re-run on the host with exactly these
four arguments reproduced the per-folder counts the docs quote (Inbox 10,912 / Sent Items 4,964 /
Deleted Items 2,461 / Junk Email 1,663) and the seven-day window count (1,612). Those numbers
appear in `Docs/magic-numbers.md` and `Docs/autonomous-session-log.md` as measured facts, so the
match is between two independently recorded things rather than a tautology.

`testbed.json` also carries the whole expected plan output, so a rebuilder can tell a correct
rebuild from a subtly different one without needing the old store to compare against.

**Its `vmName` says `OutlookAI-TestVM`, and it stays that way.** That field is *provenance* - the
guest these numbers were taken on - not a build target, and the two are now labelled apart in the
file itself. Repointing it at `OutlookAI-Indexed` today would assert that measurements were taken
on a machine nobody has built yet, which is worse than a name that looks stale; emptying it would
throw away the one record of where the numbers came from. It changes when the measurements
themselves are retaken, and the whole record changes with it. The guests to *build* are named in
§4a, and nothing defaults to either of them.

**The manifest itself is deliberately NOT committed.** It is 2.9 MB of EntryIDs describing one
machine's mailbox state, a build regenerates it, and it belongs in the gitignored
`McpServer/OutlookAI.McpServer.Tests/live-fixtures/` directory - which is where the recovered
copy now lives, as `live-fixtures/vm-corpus/corpus-vm2.jsonl`.

### Every guest gets its OWN corpus id, and the manifest is named after it

**Two guests, one manifest file, and the file is the only thing that can remove a corpus.** A
manifest is named `corpus-<corpusId>.jsonl` - after the **corpus**, not after the guest - and
`host/Copy-FromGuest.ps1` lands every guest's manifest in the same shared directory. With one
guest that is harmless. With two it is not: the second pull replaces the first's manifest, and a
manifest is the EntryID allowlist `corpus-teardown` requires (EntryID **and** ordinal tag, both,
always) as well as the file `corpus-verify` reads for freshness. Lose one and that corpus becomes
items in a real store that nothing is entitled to delete.

**The fix is different ids, not different directories.** A per-guest subdirectory would let the
two corpora keep one name and hide the collision behind a path. They are not one corpus: different
machines, different index state, genuinely different populations - and they already have to be
told apart in each machine's settings file and in every measurement that quotes one. The id is
what names the population, so the id is what has to differ.

**"Measurements and logs get a directory per guest" below gives those files exactly the directory
this paragraph refuses the manifest, and that is not a contradiction** - the two halves of the
pull are wrong in different ways. Read them together before changing either.

| Corpus id | Guest | Indexed | State |
| --- | --- | --- | --- |
| `vm2` | `OutlookAI-TestVM` | **unrecorded** | **built and measured** - section 3 above is its record |
| `vm-indexed` | `OutlookAI-Indexed` | yes | not built; reserved name. `Corpus A` in the store layout |
| `vm-unindexed` | `OutlookAI-Unindexed` | no | not built; reserved name. `Corpus B` in the store layout |

**Why those two names.** They name **the property that actually differs** - the index state - which
is the only reason the two guests exist; a rebuilder reading `corpus-vm-indexed.jsonl` in a shared
directory knows what it is without opening it, which is exactly what failed here. They map
one-to-one onto the guest names, so there is no second mapping to remember or get wrong. They are
hyphenated rather than underscored on purpose: `_` is DASL's single-character wildcard, and an id
containing one turns `CorpusPlan.DaslSubjectFragment` from an exact match into a superset -
harmless, because every caller re-checks the EntryID, but there is no reason to spend it. And
**neither id contains the other**, the same discipline `T1/CorpusTagSeparationTests` enforces on
the two subject tags, so a substring search for one corpus can never select the other. `vm3` and
`vm4` would have satisfied "distinct" and told a reader nothing, which is the modelling error this
decision exists to avoid.

`vm2` keeps its name. It is provenance - the corpus every published measurement is a statement
about - and renaming it to fit a convention invented afterwards would break the pin
`.github/scripts/check-testbed-references.ps1` holds across `testbed.json`,
`guest/Build-Corpus.ps1` and `Docs/corpus-measurement-plan.md`, for no gain.

`testbed.json` carries this as **`corpusIdConvention`, a separate top-level key** rather than
extra fields on `corpus`. That separation is the point: `corpus` is a record of one corpus that
exists and was measured, and everything in it was read off a machine. The convention block is a
naming rule plus reserved ids, and its seed, anchor and count are explicitly `null` because those
corpora **do not exist yet**. Folding the two together would put unmeasured placeholders in the
one place this repository treats as measured fact.

**Whether the two new corpora share a seed and anchor is still open** (`Docs/live-tier-on-the-vm.md`
section 8, item 16). Nothing here settles it, and nothing here needs to: two corpora may share
every generator parameter and still must not share an id, because the id names the population and
those are two different stores either way.

**The backstop, for the day somebody reuses an id anyway.** `host/Copy-FromGuest.ps1` refuses to
replace a file that is already in the destination and holds different content, unless `-Force` is
passed. It hashes the guest's copy before it copies anything, so the refusal happens instead of
the overwrite rather than after it, and the message says what the file is for. **The guard now
covers every file it pulls**, not only manifests - but the manifest is the reason the guard has to
be able to fire at all, and that is what keeps manifests in the shared root (next section).

### Measurements and logs get a directory per guest, and that is a different fix

`measure.jsonl` and the `*.log` files had fixed names in a shared directory, so with two guests
**every second pull silently overwrote the first**. That is not a tidiness problem: comparing the
indexed and the unindexed guest is the entire reason there are two of them, and replacing guest
one's transcript with guest two's destroys exactly the comparison the two-VM design exists to
produce. The loss looks like nothing at all.

**Decided 2026-09-15: the pull lands per guest.**

| What | Where it lands | Named after |
| --- | --- | --- |
| `corpus-<corpusId>.jsonl` | `<Destination>\` - the shared root | the **corpus** |
| `measure.jsonl`, `*.log` | `<Destination>\<VMName>\` | the **machine** |

**Why a directory here when the section above refused one for manifests.** Because the two files
are wrong in different ways, and the fix has to match the fault:

- For the **corpora** the **identity** was wrong. Two genuinely different populations shared one
  id, and the id is what appears in every subject, in the teardown match and in every measurement
  that quotes one. A directory would have let them keep that one name and hidden a modelling error
  behind a path. So the *names* had to change.
- For **measurements and logs** the identity is correct. There is one thing called "the
  measurement transcript" and this really is it; what differs between two copies is **which
  machine produced them**, which is precisely what a directory expresses. Stamping the machine
  into the file name would invent an identity these files do not have.

Rename what is genuinely two different things; separate by directory what is one kind of thing
arriving from two different places.

**The manifest deliberately stays in the shared root.** Moving it down with the rest would disarm
the backstop above: that refusal can only fire when two different contents arrive at **one path**,
and a per-guest directory guarantees they never do. Two guests mistakenly given the same corpus id
would quietly stop colliding and the modelling error would go back to being silent - the state the
id decision was made to end. The cost is that one pull lands in two places; that is the price of a
guard that can still fire.

**An existing flat destination from an earlier pull is left exactly where it is.** A destination
used before this change may hold a `measure.jsonl` or a `.log` at the root. The script does not
move them and does not delete them, because **which guest wrote them is recorded nowhere** - that
is the defect being fixed, and filing them under a guess would turn a stale file into a false
attribution in the very comparison this exists to protect. It names them on every pull until they
are gone. Move them into the right guest's directory if you know which one it was; delete them if
you do not.

### What is actually in the store on the VM

Reconciling the plan against the census taken straight after the build, because the two do not
match exactly and the difference is not a fault:

| Folder | Plan | Store | Difference |
| --- | --- | --- | --- |
| Inbox | 10,912 | 10,912 | - |
| Sent Items | 4,964 | 4,964 | - |
| Junk Email | 1,663 | 1,663 | - |
| Deleted Items | 2,461 | 2,467 | +6, all unread - the probe items `corpus-probe` creates and deletes |
| Outbox | 0 | 2,761 | +2,761, **all unread** |

The Outbox residue is the known `MSGFLAG_SUBMIT` defect, and this is the second independent
confirmation of its identity: 2,761 is *exactly* the plan's unread count, not approximately. The
first confirmation was the 40,000-item build's 5,532. The build now clears `MSGFLAG_SUBMIT` on
every item, so a rebuild today should leave the Outbox empty - **and if it does not, that is the
signal that the fix did not take**, because the count is predictable in advance.

### A stale corpus is REBUILT, not re-anchored

A corpus is anchored on a fixed date and every test asks its question against the clock, so
roughly six weeks after a build the narrow measurement windows select nothing - while every test
asking about them still passes, because selecting nothing is a valid answer about an empty
window. `corpus-verify` is pure, runs on the host, and refuses when that has happened.

**The repair is a rebuild: `corpus-teardown --execute` (or delete the `.pst`), then
`corpus-build`.** It is deterministic, and the recorded build was 20,000 items in 13m25s.

**`corpus-reanchor` is retired as of 2026-08-25 and refuses when invoked.** Its date writes do
not land on already-delivered items: the write method is chosen by a probe that creates
*throwaway* items - the dry run says so outright - and is then reused, unverified, on existing
ones. A run over 20,000 items reported `rewritten 20,000, refused 0, failed 0` while dating every
item inside the six minutes the tool had been running, destroying the age-band structure the
corpus exists for and overwriting the manifest with the read-back values. The command is kept
rather than deleted because the per-item write-landed guard added afterwards is what now stops it
on the first item, and that evidence is worth keeping. Run it and it tells you all of this.

---

## 3b. The settings file must declare a BYSTANDER, and the corpus store is one

**The live tier refuses to start without one.** `bystanderStoreDisplayNames` in
`live-test-settings.json` names the stores the count tripwire watches and **nothing** writes to.
If that leaves no watched store which is both non-hub and denied every write, the tier refuses at
the top of the run and names the two keys to edit. There is no flag that turns the refusal off.

That is not bureaucracy. The tripwire exempts the hub, because the hub is where the suite writes.
A configuration with nothing else to look at still censuses every folder, still identifies
nothing, and still prints `0 failure(s)` - a line produced by arithmetic that could not have
reached any other answer, and which then sits in a run report looking exactly like an earned one.

**The corpus store is declared a bystander too, and this one is load-bearing.** No live test
writes to a corpus: the freshness check reads the manifest and never the store, and re-anchoring
is an operator action run from the accountless profile. But a corpus store has to appear in
`expectedStoreDisplayNames` to be censused at all, and every non-hub entry of that list is inside
the identity-draft grant unless something says otherwise. Left undeclared, two different code
paths write into the measurement corpus:

* **the identity tests** create one draft per granted store - so they would draft into the corpus
  the moment this machine gains the dummy mail account it is getting;
* **the post-run artifact sweep** counts subjects carrying `[OutlookAI-McpTest]` and deletes what
  it finds - and the corpus generator *used to* put that exact tag at the front of every corpus
  subject, so the sweep would have found the whole corpus and tried to remove it.

Declaring the store turns both into a refusal at the write guard instead.

**Since 2026-08-25 the sweep half is closed at source as well: corpus items carry their own tag,
`[OutlookAI-Corpus]`.** The bystander declaration was a configuration fix for something that
should never have been possible by construction, and it left the sweep *failing* on the corpus
store rather than skipping it - a store deliberately full of tagged items can never satisfy "zero
tagged artifacts", and a run whose normal outcome is a refusal gets muted. A corpus subject now
contains the text `OutlookAI-McpTest` nowhere at all, so it cannot match the sweep's DASL
prefilter or its `Contains`, whatever any settings file says. `T1/CorpusTagSeparationTests` fails
the build if the two tags are ever made equal again, or if either starts containing the other's
bracket-free fragment.

**Keep declaring the corpus store a bystander anyway.** It is what stops the identity tests
drafting into the corpus, which is a separate path and still open.

**Declare only corpus stores this Outlook profile actually mounts.** A declared bystander is
watched whether or not any other list names it, so a name the profile does not have gets
censused, is not found, and refuses the tier. `Corpus B` lives in the *other* Windows account's
profile (`Docs/live-tier-on-the-vm.md` §1.1), so it belongs in that machine's settings file,
declared the same way - not in the indexed account's.

**One consequence to know about before the first run.** With the hub, the corpus store and the
plain bystander all accounted for, the **three-store floor** leaves the identity tests no store
they may draft in, so they iterate an empty list. **They no longer pass silently when that
happens** - as of 2026-09-15 they announce `PROVED NOTHING:` with the reason and refuse outright
on a `Production` profile, and they carry `Requires=IdentityAccount` as well as
`Requires=MailAccount`.

**The fix is a build step, not a caveat to live with:** `Docs/live-tier-on-the-vm.md` §2.8b adds
a second mail account, declared in `expectedStoreDisplayNames` and named nowhere in
`bystanderStoreDisplayNames`. Build it and the tests prove something instead of announcing that
they cannot. **Do not "fix" an announcing machine by adding `&Requires!=IdentityAccount` to the
filter** - that deselects the tests and deletes the only record that the identity path is
unverified, which is the vacuous green both mechanisms exist to prevent.

---

## 4. Credentials

**Never in this repository. It is public, and a guest password has already been published from
it once** - it survives in git history and had to be rotated.

| What | Where it lives | How to create it |
| --- | --- | --- |
| Guest account password (PowerShell Direct, autologon) | `McpServer/OutlookAI.McpServer.Tests/live-fixtures/vm-credentials.json`, gitignored | Set it when you create the account. Set it to **never expire**: a maximum password age silently breaks the tier and recreates this problem. |
| Live-test machine coordinates (store names, manifest path, sink ports) | `McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json`, gitignored | Copy `Testbed/live-test-settings.example.json` and fill it in. |
| Dummy mail account password | wherever the sink is configured; the sink accepts anything | Anything. It is a loopback sink with no authentication. |

`McpServer/**/live-fixtures/` is gitignored, and
`.github/scripts/check-testbed-references.ps1` asserts that the rule still covers every path
declared absent because it is machine-local - the credential file, the settings file and the
corpus manifest among them. An ignore rule that is deleted is silent until the day something
lands.

**If you rotate the guest password after the dummy mail account exists, do not use an admin
reset.** An admin reset destroys that account's DPAPI master key, which takes Outlook's saved
account password with it. That was free in August 2026 only because the profile had no mail
accounts yet.

---

## 4a. Nothing guesses which guest you mean

**THREE MACHINES COEXIST during the changeover**: `OutlookAI-Indexed` and `OutlookAI-Unindexed`
are being built, and `OutlookAI-TestVM` - the guest every published measurement was taken on -
stays until its replacements are proved (`MEDIA.md`, "never destroy a working testbed before its
replacement runs"). **A default that silently picks one of three is the exact shape of mistake
this testbed keeps making**, so as of 2026-09-15 there are no VM-name defaults left in
`host/`.

**Two changes, decided together.**

1. **The credential's `vmName` is empty.** `host/Get-GuestCredential.ps1` refuses a credential
   pinned to a different VM, and that refusal fires *before anything else runs* - it stopped
   `New-AnswerFile.ps1 -VMName OutlookAI-Indexed` at step one while the file still said
   `OutlookAI-TestVM`. The loader only enforces the match when the field has a value, so an
   empty string means "usable for any guest", which is what one shared account across three
   machines actually needs. Pinning stays available for anyone who genuinely has one credential
   per machine.
2. **`-VMName` / `-Name` is `[Parameter(Mandatory)]`** in `Copy-FromGuest.ps1`,
   `Copy-ToGuest.ps1`, `Get-GuestCredential.ps1`, `New-AnswerFile.ps1`, `New-TestbedVm.ps1` and
   `Set-TestbedLease.ps1`. Omit it and the script asks, or fails; it never assumes.
   `Publish-GuestPayload.ps1` takes no VM name at all and correctly does not need one - it only
   builds on the host, and one payload serves all three guests. The naming happens at the
   copy-in.

**One deliberate exception: `host/Invoke-TestbedIdleSave.ps1` keeps its default of all three
names.** Its `-VMName` is an **allowlist, not a target** - it bounds the set the saver may touch
at all rather than picking one to act on, and naming every known guest *is* the intent. Its
failure direction is the safe one too: a wrong or stale entry makes it do less (a VM is not
saved, the host keeps its RAM, and `Get-VM` shows it), where a wrong default elsewhere acts on a
machine nobody was looking at. Making it mandatory would also break the scheduled task outright -
`host/Register-IdleSaveTask.ps1` invokes it with `-NonInteractive` and no arguments, so a
mandatory parameter cannot prompt; it would throw every fifteen minutes for ever and the only
symptom would be a host that never reclaims its RAM. **Keep that list in step with the guests
that exist:** drop `OutlookAI-TestVM` when the old guest goes, and add any new guest the day it
is built, because a testbed VM missing from the list is simply never saved.

---

## 4b. Profiles, PSTs and accounts WITHOUT the GUI - and the one part that cannot be

**Where this came from.** Everything after the Windows install was "by hand", and §2.5 of
`Docs/live-tier-on-the-vm.md` said outright that profile creation is "not recorded" with the Mail
control panel as the assumed route. That route had been driven once, through a vision model, and
it was slow and expensive enough to be worth not repeating.

**THE SCRIPTS IN §5 HAVE NEVER BEEN EXECUTED.** They were written by an agent forbidden to run
them - the machine they were written on is the maintainer's workstation, with a real profile and
real delegate mailboxes on it - and verified by PARSING alone. Treat them as a draft to iterate
against a checkpoint, not as a build step that works. Each carries its own banner saying so, and
each is designed to **fail loudly rather than half-succeed**, because a script that silently
half-works costs a revert and a rebuild.

**Four of the five things are automatable. One is not.**

| What | How | Where |
| --- | --- | --- |
| Profile with **no mail accounts** - the corpus profile, which `corpus-build` requires | `IProfAdmin::CreateProfile`, no default services | `guest/New-OutlookProfile.ps1` |
| **PST with an exact display name**, `@` included | `IMsgServiceAdmin::ConfigureMsgService` carrying `PR_DISPLAY_NAME` **at creation** | `guest/Add-OutlookPstStore.ps1` |
| **Default-profile switch**, no prompt | `IProfAdmin::SetDefaultProfile` + `PickLogonProfile` | `guest/Set-DefaultOutlookProfile.ps1` |
| The identity account's **signature** | the shipped `manage_signature` tool | `guest/Set-AccountSignature.ps1` |
| **A POP3 mail account, and its delivery store** | **nothing free can.** GUI, once per guest | `guest/New-PopAccountPrf.ps1` is the spike that proves it |

**Why the mail account is closed rather than merely hard**, because somebody will want to try
again: the object model has no `Accounts.Add` and `Account.DeliveryStore` is read-only; MAPI has
no POP3 message service to create, because account administration moved behind the undocumented
`IOlkAccountManager`; the registry has no published working recipe on 16.x and its stored
passwords are DPAPI-sealed per user per machine; and the one documented text format, a `.prf`,
cannot carry `PROP_ACCT_DELIVERY_STORE` because that is a binary EntryID and an INI file is text.
Three routes, three unrelated reasons, same answer. A paid component (Redemption) is the only
thing that plausibly closes it, this project has no third-party dependency today, and adding one
is the maintainer's call rather than a script's.

**So take a checkpoint straight after the GUI pass.** That is the whole consolation prize, and it
is a real one: everything before it is scripted and reproducible, so a checkpoint there turns the
only unscripted step from a per-rebuild cost into a **per-guest-lifetime** one.

**Why Extended MAPI rather than the registry.** `IProfAdmin` and `IMsgServiceAdmin` are documented
and supported; the registry layout underneath is reverse engineering with no published end-to-end
recipe on 16.x. The stronger reason is specific to this testbed: the PST provider's configure call
applies `PR_DISPLAY_NAME` **when the store is created**, which is the only route anybody found to
a store named exactly what we asked for - `Store.DisplayName` is read-only in the object model, and
`Namespace.AddStoreEx` adds a PST perfectly well but cannot name it.

**The preflight that says this is worth trying at all**, measured on a guest running
16.0.17932.20996: `DLLPathEx` resolves to a real `msmapi32.dll` under the Click-to-Run `root\VFS`
tree, Office is x64 and PowerShell is a 64-bit process, so Extended MAPI loads and the bitness
matches. `DLLPath` being an unresolvable bare filename beside it is the **healthy** shape on
Click-to-Run, not a fault.

**Run `New-OutlookProfile.ps1 -Preflight` first, on a checkpoint you are willing to lose.** It
initialises MAPI, reads the profile table and stops. If the C# interop is going to take the
PowerShell process down with it - the most likely first failure, and one that produces an access
violation with no error text - that is where it happens, with nothing at stake.

**The guard that keeps these off the wrong machine.** Every one of them refuses unless the session
is logged on as `vmadmin`, the guests' autologon account (§2). It is not silenceable by a flag:
the only way past it is `-ExpectedUser <name>`, which is a thing nobody does by accident. The
second Windows account of `Docs/live-tier-on-the-vm.md` §2.4 will need exactly that.

**One question these close on first run, and it is worth running them for that alone.**
`Add-OutlookPstStore.ps1 -NameProbe` answers §6 item 10 - whether Outlook accepts `@` in a store
display name - in a throwaway profile it creates and deletes. It reports **accepted**, **rejected**
or **transformed**, and the third is why a probe is better than trying it on the real store: a
silently-renamed store is one the tests cannot find by name, on a machine that looks correctly
built.

The findings behind all of this, with every claim labelled Microsoft-documented,
community-reported, guest-measured or inferred - and a list of what could **not** be established -
are in `.work/profile-automation-research.md`. That file is gitignored scratch: if it matters
tomorrow, it has to move somewhere tracked.

---

## 5. What is in here

| Path | What it is |
| --- | --- |
| `testbed.json` | The parameter set. Corpus quad, expected plan output, build cost, guest layout, and an explicit list of what is still unrecorded. Its `vmName` is the guest this was **measured on**, not a guest to build (§3). |
| `live-test-settings.example.json` | Complete example of the gitignored settings file, every field present, placeholders only. |
| `guest/autounattend.template.xml` | The unattended-install answer file. Locale, disk layout, local account, autologon - and placeholder tokens where the password goes. Contains no credential and must never contain one. |
| `host/New-AnswerFile.ps1` | Fills that template from the gitignored credential and packages it as a small ISO. Writes into gitignored scratch only, and refuses anywhere else. `-VMName` is mandatory (§4a). |
| `guest/Complete-FirstLogon.ps1` | The first-logon fix-ups the answer file cannot express: the en-NL language list, the home location, the locales, no sleep, no fast startup. Logs and reads back everything it set. |
| `host/New-TestbedVm.ps1` | Creates the Hyper-V guest, attaches the Windows ISO and the answer volume, boots it, and records the spec it chose. `-Name` is mandatory (§4a). |
| `host/Publish-GuestPayload.ps1` | Publishes the MCP server and the remediation tools on the host and zips them for copy-in. Host-only, so it takes no VM name; one payload serves all three guests. |
| `host/Get-GuestCredential.ps1` | Loads the guest credential from the gitignored fixtures directory. Documents the one place a credential may live; contains none. `-VMName` is mandatory (§4a). |
| `host/Copy-ToGuest.ps1` | Copies a file or a zip into the guest over PowerShell Direct. `-VMName` is mandatory (§4a). |
| `host/Copy-FromGuest.ps1` | Gets results, logs and the corpus manifest back out. `-VMName` is mandatory (§4a), and it also names the subdirectory results land in. **The manifest goes to the shared root** - safe because each guest has its own corpus id, and kept there so a reused id still collides visibly; **`measure.jsonl` and the logs go to `<Destination>\<VMName>\`**, because what differs about a transcript is the machine that produced it (§3). It refuses to replace any pulled file whose content differs, unless `-Force` says you mean it. |
| `guest/Register-InteractiveTask.ps1` | The session-1 scheduled-task recipe. Everything COM-touching goes through it. |
| `guest/OutlookMapiInterop.ps1` | Shared Extended MAPI layer for the four profile scripts below - `IProfAdmin`, `IMsgServiceAdmin`, the table readers, and the guest guard. **Dot-sourced, never run.** **Never executed** - see its banner, and §4b. |
| `guest/New-OutlookProfile.ps1` | Creates an Outlook profile with no GUI: account-less (the corpus profile) or carrying named PSTs. `-Preflight` checks the MAPI interop alone, first. **Never executed.** |
| `guest/Add-OutlookPstStore.ps1` | Adds a PST to a profile with an **exact** display name, and `-NameProbe` settles §6 item 10 - whether Outlook accepts `@` in one. **Never executed.** |
| `guest/Set-DefaultOutlookProfile.ps1` | Switches the default profile and switches the profile prompt off. Closes §6 item 5. **Never executed.** |
| `guest/Dump-UiaTree.ps1` | **Read-only.** Dumps the UIAutomation tree of an open dialog and prints a VERDICT: whether Outlook's account wizard is a classic Win32 property sheet (addressable by locale-invariant numeric `AutomationId`) or Office's own DirectUI chrome (no stable ids - dead for a PowerShell client). Two minutes, and it decides the whole GUI-automation route. **Never executed.** |
| `guest/Rename-OutlookStore.ps1` | Renames a store to an exact display name, which the tier profile needs because Outlook names the store it mints 'Outlook Data File' and section 2.6 requires the hub store to be named as an SMTP address. **MEASURED WORKING 2026-09-15** - and it settles a question no documentation could: `Store.DisplayName` is read-only, but it DOES follow a rename of the store's root folder, `@` and all, without breaking the account's delivery-store binding. |
| `guest/tier-profile-forcepst.prf` | **The one that works.** A PRF with no PST service and no `DefaultStore`, used with `ForcePSTPath`, so Outlook mints the POP3 account's delivery store itself - a store Outlook mints is a store Outlook binds, and binding is the step a text file cannot perform. Measured 2026-09-15: `SmtpAddress`, `DeliveryStore` and its Drafts folder all resolve. Prefer this over `tier-profile.prf`, which leaves `DeliveryStore` NULL. |
| `guest/Set-OfficeFirstRunSuppressed.ps1` | Suppresses Office's own first-run dialogs and pins CLASSIC Outlook. `ImportPRF` silences the profile wizard and nothing else - a guest with a perfect profile still came up on "Your privacy matters", and a dialog on an unattended guest is a hang, not a prompt. Ends by saying registry values prove nothing and to start Outlook and look. **Never executed.** |
| `guest/New-PopAccountPrf.ps1` | A **spike**, not a route: the one free candidate for creating a POP3 account, plus the read-back that says how far it got. Expected to fail; §4b says why. **Never executed.** |
| `guest/Set-AccountSignature.ps1` | Gives the identity account its signature, by driving the shipped `manage_signature` tool rather than improvising. Runs **after** the accounts exist. **Never executed.** |
| `guest/Install-MailSink.ps1` | Installs the loopback sink the POP3 account points at, from a **staged** package - never a download - pinned by SHA-256. Then **proves the round trip over raw sockets** rather than reporting two open ports, which is all the suite's own `T2/LiveMailSink.cs` probe does: submission accepted, message retrievable, dot-stuffing intact, `DELE` honoured, message numbers stable, nothing re-served after a restart. Two of those are **expected to fail** against smtp4dev as shipped and are asserted rather than assumed - see §6 item 11. Touches no Outlook, no MAPI and no mail item. **Never executed.** |
| `guest/Set-OutlookIndexingDisabled.ps1` | Takes a guest's Outlook **out of** the Windows Search index, which is the only thing that makes `OutlookAI-Unindexed` different from `OutlookAI-Indexed` - both are built from the same answer file and both come up indexed. It never disables the Windows Search **service**: a machine with no indexer is not a store with no index frontier, it is the product's "index unreachable" branch, and it removes `index.perStore[]` - the one instrument the runbook says to read. `-Verify` probes the **catalog** rather than reading registry values back, twice, and returns one of four verdicts; two of them are "not an answer" on purpose. **Run it BEFORE `guest/Build-Corpus.ps1`** (§1 step 8) or you also own purging what was already crawled. **Never executed.** |
| `guest/Build-Corpus.ps1` | plan, probe, build, census - with the committed parameters as defaults. |
| `guest/Invoke-GuestMeasure.ps1` | The measurement driver, recovered from the guest. Produced the numbers now in `Docs/magic-numbers.md`. |
| `guest/Measure-SweepCost.ps1` | Per-folder / per-item sweep cost, out of band. **Reconstructed, never executed** - see its banner. |
| `guest/tier-profile.prf` | The tier profile as an Outlook .prf: one Unicode PST plus one POP3 account on the loopback sink. A template with `{{...}}` tokens - `New-TierProfile.ps1` renders it. **Never imported by Outlook.** |
| `guest/New-TierProfile.ps1` | Creates the tier profile by importing that .prf, then reads the profile hive back and asserts whether Outlook honoured it. Three modes: dry run, `-Execute`, `-Verify`. **Never executed** - see its banner and §5c. |
| `guest/Set-AccountWizardClassic.ps1` | Restores Outlook's classic account wizard and stops AutoDiscover reaching the network. Prerequisite for the UI Automation route. **Never executed.** |
| `guest/Dump-UiaTree.ps1` | Dumps the UI Automation tree of a dialog that is already open. Read-only; invokes nothing. Decides whether the account wizard can be driven at all. **Never executed against Outlook.** |

---

## 5c. The last three scripts are an EXPERIMENT, not the supported path

Step 7 above still says "by hand", and it still means it. `New-TierProfile.ps1`,
`Set-AccountWizardClassic.ps1` and `Dump-UiaTree.ps1` exist because creating a POP3 account
programmatically turned out to have no proven free route - the object model is read-only for
accounts, Extended MAPI can no longer create POP3 services, and the one component that can is
excluded by the repository's Dependencies rule. They are the two remaining candidates, written so
that a guest can settle them in one checkpoint cycle each.

**None of them has ever run.** Each carries a banner saying so, and each verifies its own result
and exits non-zero rather than reporting a success it did not check. Read the banner before you
trust an output, and replace it with what actually happened once one of them has run.

**Two questions decide everything, and both are cheap:**

* **Does Outlook 16.x process a .prf's internet-account sections at all?** Every literal POP3 .prf
  Microsoft ever published is 2000-2007 era. `New-TierProfile.ps1 -Verify` fails loudly with that
  exact diagnosis if no account subkey appears.
* **Does the classic account wizard expose non-empty, numeric `AutomationId`s?** If it does, the
  `Dump-UiaTree.ps1` output *is* the specification for a driver. If the ids are empty or the
  framework reports `DirectUI`, that route is dead for a PowerShell 5.1 client and no driver
  should be written.

**And there is a third answer, which is to need less.** Most of what the POP3 account was for is
satisfied by *any* account whose SMTP address matches - the product never reads
`Account.AccountType` - and most of the rest is seedable by writing into a PST directly, which the
corpus generator already does. The analysis behind all three, with a source for every claim, is in
`.work/pop3-account-routes.md`, which is gitignored scratch: move it under `Docs/` if it should
outlive the session that produced it.

---

## 5b. Idle VMs are SAVED, so the host gets its resources back

The testbed VMs are not left running. A scheduled task saves any testbed VM nobody is using,
and this section is the contract for it.

**Saved, not paused, and the difference is the whole point.** `Suspend-VM` freezes a VM but
keeps its memory resident: the host gets its CPU back and none of its RAM. `Save-VM` writes the
guest's memory to disk and releases the RAM entirely, and resuming is still far faster than a
boot because the guest never shut down - Outlook is still running, the profile is still open,
the search service is still warm. Resources back plus a fast restart is the requirement, and
saving is the only thing that gives both. The cost is disk: a saved VM's memory file is roughly
its assigned RAM.

**"In use" is DECLARED, never inferred, and this is the part worth understanding before
changing anything.** The obvious signal - watch the guest's CPU - is wrong here in a way that
would be found as a mystery rather than as a bug. A live tier run is ~27 minutes of driving
Outlook through COM, and Outlook spends much of that waiting: on a store to open, a folder to
enumerate, a save to commit. A guest that looks idle for two minutes mid-run is entirely
ordinary. Saving it there suspends a COM call, and what the operator sees afterwards is a test
that timed out for no reason on a machine that looks fine.

So anything that intends to use a VM takes a lease and renews it while it works:

    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Minutes 45 -Reason 'live tier'
    # ... work, renewing before it expires ...
    Testbed/host/Set-TestbedLease.ps1 -VMName OutlookAI-Indexed -Release

A VM is saved only when it is one of the named testbed VMs, is Running, holds no live lease,
and has been up past a grace period - so a VM someone has just started for work that has not
taken its lease yet is not immediately put back to sleep.

| Script | What it does |
| --- | --- |
| `host/Set-TestbedLease.ps1` | Take, renew or release a lease. `-VMName` is mandatory (§4a). |
| `host/TestbedLeasePath.ps1` | Where leases live, and how one is read. Dot-sourced by both sides so they cannot disagree. |
| `host/Invoke-TestbedIdleSave.ps1` | The saver. `-WhatIf` reports without changing anything. The one script here that keeps a VM-name default, because its `-VMName` is an allowlist rather than a target (§4a). |
| `host/Register-IdleSaveTask.ps1` | Registers it as the invoking user at ordinary privilege, every 15 minutes. Needs local `Hyper-V Administrators` membership, **not** elevation - `Save-VM` does not require it, and a standing elevated task would outlive the reason it was created. |

**Deliberate failure directions**, each chosen so the wrong answer is visible rather than silent:

- **An unreadable or expired lease does not protect the VM.** Treating a corrupt lease as live
  would let one truncated write pin a VM awake for ever. Getting it wrong this way saves a VM
  someone was using, which is immediately visible and fixed by resuming.
- **A lease whose holder died stops protecting when it expires.** That is why leases are short
  and renewed rather than long. Do not take an eight-hour lease to avoid renewing.
- **The saver refuses loudly when it cannot see Hyper-V.** It checks before the loop, because
  the per-VM "not on this host" skip swallows a permissions failure exactly as it swallows an
  absent VM - and an unelevated run would then skip every VM, print nothing, exit 0, and look
  like a machine where nothing was ever idle. That is not hypothetical; it is what the first
  version of this did.
- **The task only ever saves.** It never starts, stops, checkpoints or deletes a VM.

**One trap this cost, recorded so nobody repeats it.** The lease expiry is compared as a Unix
second count, not as `expiresUtc`. PowerShell's `ConvertFrom-Json` silently coerces an ISO-8601
string into a `DateTime` of Kind `Unspecified`; re-parsing that object's rendering drops the UTC
marker, and `ToUniversalTime()` then treats a UTC instant as local and shifts it by the offset.
Measured here: every lease read as having expired two hours before it was written, so the saver
would have suspended VMs that were in active use. A number cannot be coerced into anything but
a number.

## 6. What is still unknown

A runbook that implies completeness it does not have is worse than one that names its holes.
These are the holes. Each is a **question a rebuilder must answer for themselves**, not a step
that was left out.

**Things only the maintainer can answer, because only the VM knows them**

1. **Hyper-V generation, Secure Boot, TPM, vCPU, RAM, disk size, checkpoint type.**
   `host/New-TestbedVm.ps1` picks defaults and says loudly that it is picking them; it is not a
   record of the original.
2. **Defender exclusions, and everything about the ORIGINAL guest's Windows.** An indexer, a
   400 MB PST and real-time AV interact, and nobody has recorded whether an exclusion is in
   place. Note what changed here and what did not: for guests built from
   `guest/autounattend.template.xml` the edition, the ISO, the licensing stance, the computer
   name, the locale and the time zone are all decided in that file and are therefore a record.
   For the guest the published measurements were actually taken on, they remain unknown, and no
   answer file written afterwards can turn that into knowledge.
3. **Office version, channel, bitness and install method**, and how the first-run wizard is
   suppressed. Bitness in particular: the test host is x64 because the `Search.CollatorDSO`
   provider needs an x64 host, but whether Office itself must be x64 is untested.
4. **The second Windows account.** The layout in `Docs/live-tier-on-the-vm.md` §1.1 needs two
   Windows accounts, one indexed and one not. The guest has `vmadmin`. Whether a second account
   exists, what it is called, and whether it has its own clone and runtime, is unknown.
5. **Which Outlook profile is default, and how the switch is automated.** Two profiles are known
   to exist, `Outlook` and `OutlookAITest`. The switch is a registry value under
   `HKCU\...\Outlook` and Outlook must not be running when it changes - but the exact value and
   whether anything automates it is unrecorded.
   **Half-answered 2026-09-15:** `guest/Set-DefaultOutlookProfile.ps1` now does it through
   `IProfAdmin::SetDefaultProfile` and switches the profile prompt off with `PickLogonProfile`,
   verifying both. It has **never been executed**, so this stays open until it has - and it says
   nothing about which profile the ORIGINAL guest had as default, which remains unrecorded.
6. **Whether the three-store layout exists at all.** Everything measured so far was taken against
   ONE PST named `Outlook Data File`. Corpus B, the bystander store, the hub named after the
   dummy address, and the dummy account itself are a design in a document; no evidence in this
   repository shows any of them has been built.
7. **Whether the mail sink is installed**, which build, and on which ports.
8. **The PST file paths of every store except the corpus one.** The corpus store's path is
   recorded in `testbed.json` because the manifest header carried it; nothing records the others.

**Things nobody has answered yet, which a rebuilder will hit**

9. ~~**Does one Windows account's Outlook profile really stay out of the index while another's is
   in it?**~~ - **NO LONGER LOAD-BEARING** (two guests now, `Docs/live-tier-on-the-vm.md` §1.1a),
   and the half of it that still matters is **MEASURED, 2026-09-16**: the crawl scope manager on a
   real Windows 11 machine carries exactly one Outlook rule, `mapi16://{SID}/`, under
   `HKLM\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\WorkingSetRules`,
   with `Include=1` and `Default=0` - a **user** rule, one URL per Windows account. Above it sits a
   documented Group Policy, `PreventIndexingOutlook`, which outranks any user rule and is what makes
   an exclusion durable. **And one thing the runbook says is wrong**: `Docs/live-tier-on-the-vm.md`
   §1.1 justifies "one switch per account" by claiming there is no per-store URL, and there is -
   Microsoft documents excluding a single store through the Crawl Scope Manager with
   `mapi16://{SID}/StoreDisplayName($Hash)/`. The conclusion holds for the **dialog**, which offers
   one tick and cannot even display the difference; the reason given for it does not. What is still
   **open** is whether Outlook re-creates its rule on the next start.
   `guest/Set-OutlookIndexingDisabled.ps1` writes the policy as well as the rule for exactly that
   reason, and its `-Verify` is what settles it. Full evidence, every claim labelled by source and
   with an explicit list of what could not be established, in `.work/unindexed-guest.md`.
10. **Does Outlook accept `@` in a store display name?** The hub store must be named after the
    dummy account's SMTP address because several tests use the display name as an address. It
    gates the whole draft family and costs five minutes to settle.
    **Still open, but now one command:** `guest/Add-OutlookPstStore.ps1 -NameProbe -Execute
    -VerifyWithOutlook` settles it in a throwaway profile it creates and deletes, and reports
    **accepted**, **rejected** or **transformed**. Nothing in Microsoft's documentation or in any
    community source states a character restriction either way, which is weak evidence and is why
    this is a probe rather than an answer. Watch for **transformed**: a silently-renamed store is
    one no test can find by name, on a machine that looks correctly built.
11. ~~**Does smtp4dev actually serve POP3?**~~ - **ANSWERED 2026-09-15: YES, and pin 3.15.0.**
    The doubt was well-founded and the conclusion is the other way. Issue #155, "Support
    retrieval of messages using POP3", really was closed **as not planned** - in 2022. POP3 then
    shipped **three years later** from an unrelated pull request (#1888, merged 2025-10-03) with
    no issue behind it, and first appears in a stable release at **3.11.0** (2025-12-06);
    `Server/Pop3/Pop3Server.cs` is absent at 3.10.3 and present at 3.11.0. **Pin 3.15.0**, not
    3.11.0: 3.14.0 added the ability to disable POP3/IMAP by a null port, and 3.15.0 added a POP3
    `NOOP` handler. So the documented sink design **is** buildable - but POP3 here is an
    eleven-month-old, single-PR feature with one end-to-end test, and reading its source found
    two RFC 1939 violations (`TOP` advertised in `CAPA` and not implemented; `DELE` deleting
    immediately and **renumbering the mailbox mid-session**, so `DELE 1; DELE 2` deletes the
    wrong item). `guest/Install-MailSink.ps1` asserts both rather than assuming either way.

    **Four corrections fall out of this and are not yet made**, because which of them survives
    depends on a decision that is the maintainer's - a sink is a third media precondition and
    `MEDIA.md` names two, so `.work/mail-sink.md` §8 frames it rather than acting on it.
    `Docs/live-tier-on-the-vm.md` §2.7 names the winget package `RnwoodLtd.smtp4dev`; **the id is
    `Rnwood.Smtp4dev`**, so that command cannot work. Its configuration keys all live under a
    **`ServerOptions`** root object, and a key written at the file root is read by nothing. A port
    of **`0` means auto-assign**, not disabled - only `null` disables a listener. And its TLS
    settings are incomplete: **`Pop3TlsMode` is a separate key**, documented in the source as
    independent of the global `TlsMode`, so setting only the latter leaves POP3 advertising
    `STLS` to a client the PRF configured for no encryption.

    **One new collision, which nothing had recorded.** smtp4dev's POP3 checks credentials against
    nothing, but its `PASS` handler **refuses an empty password** - while
    `guest/tier-profile-forcepst.prf` deliberately carries no password key at all. POP3 has no
    anonymous mode, and on an unattended guest a credential prompt is a hang rather than a
    prompt. The likely answer is a one-time manual entry with *Remember password*, preserved by
    the checkpoint, exactly as the accounts themselves are - but §2.8 does not say so.

    **And it settles §8 item 3 of the runbook:** with `AuthenticationRequired: false`, POP3
    **never consults the username** and always serves the auto-created catch-all mailbox. No
    mailbox needs provisioning for the account name.

    Full evidence, with a source per claim and an explicit list of what could not be
    established - including whether Outlook issues `TOP` at all - is in `.work/mail-sink.md`,
    which is gitignored scratch: move it under `Docs/research/` if it should outlive the session
    that produced it. **IMAP is not a fallback**: an IMAP account gets its own store,
    `Account.DeliveryStore` is read-only, and Microsoft documents "deliver to an existing Outlook
    Data File" only for POP accounts - so the hub-PST arrival assertion is unsatisfiable with one.
12. **How the built server exe reaches the path the tier-3 tests expect.** The path is baked into
    the test assembly at build time as `AssemblyMetadata("McpServerExePath")` and points into the
    repository's `bin` tree - but the guest cannot build, so nothing puts a binary there.
    `host/Publish-GuestPayload.ps1` stages `C:\OutlookAI-Q5\server\`, which is where the recovered
    guest scripts point, and that is *not* the same path. Running tier 3 on the guest needs this
    resolved.
13. **Whether the live suite has ever run on the guest at all.** `Docs/live-tier-on-the-vm.md` §9
    says nobody has run the `Portable` subset end to end anywhere. Everything measured on the VM
    so far was driven by `guest/Invoke-GuestMeasure.ps1` talking raw stdio to the server, not by
    `dotnet test` - which the guest cannot run, having no SDK.

**Things nobody can put in a repository**

14. A Windows licence and an Office licence.
15. A host with Hyper-V and enough disk for a guest plus a ~400 MB PST plus checkpoints.
16. The guest account password (§4).

---

## 7. Keeping this honest

`.github/scripts/check-testbed-references.ps1` runs in CI beside the other two checks. It fails
when:

* a tracked document references a repository path that does not exist and is not on its declared
  list of intentionally-absent paths;
* an intentionally-absent path stops being gitignored (which would mean the next commit could
  publish it);
* the corpus parameters stop agreeing across `testbed.json`, `guest/Build-Corpus.ps1` and
  `Docs/corpus-measurement-plan.md`;
* a script in this directory is not named by the table in §5, or the table names one that does
  not exist;
* a script in this directory fails to parse;
* something credential-shaped appears in a tracked file under `Testbed/`;
* the answer-file template stops holding placeholders where the password belongs, or a filled
  `autounattend.xml` becomes tracked.

The last two are not paranoia. The value they look for has been committed to this repository
before, and an unattend password is `<Password><Value>...</Value></Password>` - which does not
read as an assignment, so the credential-shaped check would walk straight past it.
