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
question was first asked. `Testbed/MEDIA.md` records what is needed, what exists, the Office
deployment settings the current guest was built with, and the licence clocks — including the one
correction that matters: **Office's grace is 30 days, not 90**, so it expires before Windows and
a "rebuild when it expires" policy means rebuilding monthly unless the guest can reach a KMS host.

It also carries the rule that came out of nearly getting this wrong: **never destroy a working
testbed before its replacement runs.**

## 1. The order to do things in

| # | Step | What runs it | Where |
| --- | --- | --- | --- |
| 1 | Build the answer volume | `host/New-AnswerFile.ps1` | host |
| 2 | Create the guest, attach both ISOs, boot it | `host/New-TestbedVm.ps1` | host |
| 3 | Windows installs itself - edition, disk, account, autologon, locale, power | nobody: the answer file | guest, unattended |
| 4 | Install Office, the accounts and the profiles | by hand - `Docs/live-tier-on-the-vm.md` §2.2-2.6 | guest |
| 5 | Give yourself a way to reach session 1 | `guest/Register-InteractiveTask.ps1` | guest, once |
| 6 | Build the server and the tools, and copy them in | `host/Publish-GuestPayload.ps1` | host |
| 7 | Install the mail sink and the dummy account | by hand - `Docs/live-tier-on-the-vm.md` §2.7-2.8 | guest |
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
with. **One thing to know before building the second guest:** that loader refuses a credential
whose `vmName` names a different VM, so a single credential file serving both guests must leave
`vmName` empty. The check only fires when the field has a value, and with one shared account
there is no wrong machine for it to protect against. It prefers `oscdimg.exe` from the Windows ADK's Deployment Tools and falls back to the
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

**The manifest itself is deliberately NOT committed.** It is 2.9 MB of EntryIDs describing one
machine's mailbox state, a build regenerates it, and it belongs in the gitignored
`McpServer/OutlookAI.McpServer.Tests/live-fixtures/` directory - which is where the recovered
copy now lives, as `live-fixtures/vm-corpus/corpus-vm2.jsonl`.

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
  it finds - and the corpus generator puts that exact tag at the front of *every* corpus subject,
  so the sweep would find the whole corpus and try to remove it.

Declaring the store turns both into a refusal at the write guard instead.

**Declare only corpus stores this Outlook profile actually mounts.** A declared bystander is
watched whether or not any other list names it, so a name the profile does not have gets
censused, is not found, and refuses the tier. `Corpus B` lives in the *other* Windows account's
profile (`Docs/live-tier-on-the-vm.md` §1.1), so it belongs in that machine's settings file,
declared the same way - not in the indexed account's.

**One consequence to know about before the first run.** With the hub, the corpus store and the
plain bystander all accounted for, the three-store layout leaves the identity tests no store they
may draft in, so they iterate an empty list and pass without proving anything. They are marked
`Requires=MailAccount`; the VM's single mail account is the hub's. `TODO.md` carries this as an
open item - do not read a green identity test on this machine as evidence.

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

## 5. What is in here

| Path | What it is |
| --- | --- |
| `testbed.json` | The parameter set. Corpus quad, expected plan output, build cost, guest layout, and an explicit list of what is still unrecorded. |
| `live-test-settings.example.json` | Complete example of the gitignored settings file, every field present, placeholders only. |
| `guest/autounattend.template.xml` | The unattended-install answer file. Locale, disk layout, local account, autologon - and placeholder tokens where the password goes. Contains no credential and must never contain one. |
| `host/New-AnswerFile.ps1` | Fills that template from the gitignored credential and packages it as a small ISO. Writes into gitignored scratch only, and refuses anywhere else. |
| `guest/Complete-FirstLogon.ps1` | The first-logon fix-ups the answer file cannot express: the en-NL language list, the home location, the locales, no sleep, no fast startup. Logs and reads back everything it set. |
| `host/New-TestbedVm.ps1` | Creates the Hyper-V guest, attaches the Windows ISO and the answer volume, boots it, and records the spec it chose. |
| `host/Publish-GuestPayload.ps1` | Publishes the MCP server and the remediation tools on the host and zips them for copy-in. |
| `host/Get-GuestCredential.ps1` | Loads the guest credential from the gitignored fixtures directory. Documents the one place a credential may live; contains none. |
| `host/Copy-ToGuest.ps1` | Copies a file or a zip into the guest over PowerShell Direct. |
| `host/Copy-FromGuest.ps1` | Gets results, logs and the corpus manifest back out. |
| `guest/Register-InteractiveTask.ps1` | The session-1 scheduled-task recipe. Everything COM-touching goes through it. |
| `guest/Build-Corpus.ps1` | plan, probe, build, census - with the committed parameters as defaults. |
| `guest/Invoke-GuestMeasure.ps1` | The measurement driver, recovered from the guest. Produced the numbers now in `Docs/magic-numbers.md`. |
| `guest/Measure-SweepCost.ps1` | Per-folder / per-item sweep cost, out of band. **Reconstructed, never executed** - see its banner. |

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
| `host/Set-TestbedLease.ps1` | Take, renew or release a lease. |
| `host/TestbedLeasePath.ps1` | Where leases live, and how one is read. Dot-sourced by both sides so they cannot disagree. |
| `host/Invoke-TestbedIdleSave.ps1` | The saver. `-WhatIf` reports without changing anything. |
| `host/Register-IdleSaveTask.ps1` | Registers it as SYSTEM in session 0, every 15 minutes. Needs elevation. |

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
6. **Whether the three-store layout exists at all.** Everything measured so far was taken against
   ONE PST named `Outlook Data File`. Corpus B, the bystander store, the hub named after the
   dummy address, and the dummy account itself are a design in a document; no evidence in this
   repository shows any of them has been built.
7. **Whether the mail sink is installed**, which build, and on which ports.
8. **The PST file paths of every store except the corpus one.** The corpus store's path is
   recorded in `testbed.json` because the manifest header carried it; nothing records the others.

**Things nobody has answered yet, which a rebuilder will hit**

9. **Does one Windows account's Outlook profile really stay out of the index while another's is
   in it?** The entire two-account layout rests on this and it is derived from how the MAPI scope
   is addressed, not measured. Verify it before building anything else.
10. **Does Outlook accept `@` in a store display name?** The hub store must be named after the
    dummy account's SMTP address because several tests use the display name as an address. It
    gates the whole draft family and costs five minutes to settle.
11. **Does smtp4dev actually serve POP3?** `Docs/live-tier-on-the-vm.md` §2.7 specifies POP3 on
    port 110 and §2.8 adds the dummy account as POP3, and `MailSinkSettings.RetrievePort`
    documents itself as POP3 - but smtp4dev v3 is usually described as SMTP plus **IMAP**. If it
    is IMAP-only, the sink section is wrong in a way that only shows up as mail sitting in a
    sink nobody can retrieve from. **Settle this before installing anything**, and if it is
    IMAP-only, the choices are an IMAP dummy account or a different sink.
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
