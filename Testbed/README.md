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
| 3a | **Eject both discs, THEN take the first checkpoint, then delete the answer ISO** | `host/New-TestbedVm.ps1 -Name <guest> -CompleteInstall -Execute` - run it straight after step 2; it waits for the guest's `first-logon.log` to end in its DONE line (the answer volume carries `guest/Complete-FirstLogon.ps1`, which runs from it at first logon, so it is needed until then), ejects the Windows ISO and the answer volume, takes `CP-01-WIN-CLEAN` with no disc in it, and deletes the answer ISO - which holds the guest password in clear text - once nothing on the host references it. **Order is the point:** a checkpoint taken with a disc attached references that file for as long as the checkpoint exists, so the file cannot be deleted without breaking a restore. Never take `CP-01-WIN-CLEAN` by hand. **Not yet run** - §1b says what that means | host |
| 4 | Install Office - **and do not start Outlook** | By hand - `Docs/live-tier-on-the-vm.md` §2.2-2.4 (§2.4, a second Windows account, is no longer needed), with `.work/office-odt/Testbed.xml` (`MEDIA.md`). Checkpoint here (`CP-02-OFFICE-INSTALLED`). Outlook's first start is spent on step 4c, deliberately | guest |
| 4a | **Reach session 1, and stage the guest scripts** | `host/Copy-ToGuest.ps1 -VMName <guest> -RepoRoot <the main checkout, where the credential is> -Path <scripts> -Destination C:\OutlookAI-Q5` - every guest script the steps below run, **plus `guest/OutlookMapiInterop.ps1` and `guest/Register-InteractiveTask.ps1`**: every writing script dot-sources the first for its `vmadmin` guard and fails loudly without it - since 2026-09-24 that includes the second, so the two always travel together - and everything that touches Outlook runs in session 1 through the second. There was no step for this; it used to sit at 5, after steps that already needed it | host, then guest |
| 4b | Suppress Office's first-run dialogs | `guest/Set-OfficeFirstRunSuppressed.ps1 -Execute`, then `-Verify`. It also creates the `...\Office\16.0\Outlook` key that step 4c's preflight requires. It does not cover the one-time "Check out our new look" dialog at Outlook's **second** start - COM works with it on screen | guest, session 1 |
| 4b-ii | **Auto-approve Outlook's programmatic-access prompts (Q80, offline test guests only)** | `guest/Set-OutlookProgrammaticAccess.ps1 -Execute`, Outlook closed: the documented per-user Group Policy - `AdminSecurityMode` 3 and eight `PromptOOM*` values at 2, "Automatically approve". Then, once a profile with a mail account runs (after 4c), `-Verify` in session 1: it reads `Account.SmtpAddress` over COM inside a time limit, and a prompt fails it. Without this, a guest with no network has out-of-date antivirus signatures, and Outlook's Object Model Guard then stops every out-of-process read of an address, a body or a recipient on a modal Allow/Deny prompt nobody answers. **RUN 2026-09-24 on `OAI-UNINDEXED`: `NO-PROMPT`**, against a control run with the values absent that `PROMPTED` 0.7 s after the read | guest, session 1 |
| 4c | **The tier profile - at Outlook's FIRST start** | `guest/New-TierProfile.ps1 -Execute` (the forcepst template is the default, and it writes `ForcePSTPath`), then **start `OUTLOOK.EXE` once, in session 1**, and let it settle (~90 s: it imports, mints `C:\OutlookAI-Tier\Outlook.pst` and binds the account to it; the POP3 password dialog appears and does not block COM), then `guest/Rename-OutlookStore.ps1 -StoreFilePath C:\OutlookAI-Tier\Outlook.pst -DisplayName tier@vm.invalid -Execute`, then - Outlook quit gracefully (4d) - `New-TierProfile.ps1 -StoreSinkPassword -Execute` (without a stored POP3 password Outlook prompts at every start and never reaches the sink, measured; step 7), then `New-TierProfile.ps1 -Verify`, which fails an account with none. Why first: the paragraph under this table | guest, session 1 |
| 4d | Restart the guest | **Quit Outlook first, gracefully, in session 1**: attach, refuse while an Inspector is open or any Outbox holds an item, release every COM reference but the Application, `Application.Quit()`, and wait for `OUTLOOK.EXE` to exit - never kill it. A modal dialog can make Outlook ignore that `Quit()` - measured 2026-09-24 with the Object Model Guard prompt CP-05 carries in its saved memory: `Quit()` returned, and Outlook was still up 240 s later. So answer any open dialog first with its harmless button - Deny on a guard prompt, Cancel on the POP3 logon prompt (whether that one also swallows `Quit()` was not tested; it is cancelled first on the same reasoning) - sent as the dialog's own `WM_COMMAND`: the guard prompt ignored a `BM_CLICK` while it was not the active window. Then `shutdown /r /t 0` from session 0, and wait for Heartbeat `OK` and `query session` showing `console ... Active`. **Never plain `shutdown /r`**: its time-out defaults to 30 s, and Microsoft documents that any time-out above 0 implies `/f`, which force-closes Outlook - the same harm as `taskkill`. Step 4e's `/PIM` needs a fresh start, and step 4f refuses while Outlook runs | guest |
| 4e | The corpus profile, with no account | Start **`OUTLOOK.EXE /PIM CorpusProfile`** in session 1 (its store, `Outlook Data File`, is minted in `C:\OutlookAI-Tier` - `ForcePSTPath` is per-user and outlives step 4c), create the directory for any named PST by hand (`Add-OutlookPstStore.ps1` refuses a missing one), `guest/Add-OutlookPstStore.ps1 -ProfileName CorpusProfile -DisplayName <name> -Path <pst> -Execute` per store, then `guest/New-OutlookProfile.ps1 -Name CorpusProfile -Store <name>=<pst> -Verify -WithOutlook`. **Not** `New-OutlookProfile.ps1 -Execute`: an account-less profile imported from a `.prf` stops on "Email Account Setup" at every start (§4b) | guest, session 1 |
| 4f | Restart, then make the corpus profile the default | `guest/Set-DefaultOutlookProfile.ps1 -Name CorpusProfile -Execute`. The tier import made `OutlookAI-Tier` the default and `/PIM` does not change it, and `guest/Build-Corpus.ps1` refuses unless the default profile has no mail account. **Step 9 and every live run need the default switched back to `OutlookAI-Tier` the same way** | guest |
| 5 | ~~Clear `ImportPRF`~~ - **automatic now; nothing to run** | Outlook removes `ImportPRF` itself within ~5 s of the start that imports it (measured 2026-09-24, three times, §4b-i). The protection no longer depends on anyone remembering a step: each `-Verify` removes a lingering value that names its own `.prf` once the import has run, and `guest/Build-Corpus.ps1` refuses to build while one is set, `-SkipPreflight` or not | - |
| 5b | **Put the add-in on the guest - BOTH guests - built from a named commit.** Two live tests read state only the add-in writes on its first run, and one of them runs on the unindexed guest too. Build and stage on the host with `host/Publish-AddInPayload.ps1`, copy `AddIn.zip` and the staged `vstor_redist.exe` in, then run `guest/Install-OutlookAIAddIn.ps1 -Execute` **through the interactive task** - it installs the VSTO runtime and the product's own installer silently, writes the VSTO trust entry the trust prompt would have written, starts Outlook once, and verifies the exact registry state the tests read. Outlook must be closed when it starts, and may still be running headless when it ends - restart the guest before 7b, which refuses while Outlook is up. Do it BEFORE 7b, so the index exclusion is certified with the add-in already present; on a guest already past 7b, re-run `guest/Set-OutlookIndexingDisabled.ps1 -Verify` afterwards. `Docs/live-tier-on-the-vm.md` §2.3. **RUN 2026-09-24 on `OAI-UNINDEXED` (past 7b): `ADDIN-READY` at the first `-Execute` and again after a restart.** Stage the guest script beside `guest/OutlookMapiInterop.ps1` in `C:\OutlookAI-Q5` - the suite's source, where the printed command looks for it, does not arrive until step 8b | host, then guest, session 1 |
| 6 | Build the server and the tools, and copy them in | `host/Publish-GuestPayload.ps1` | host |
| 7 | Install the mail sink, the dummy account and the identity account | **Sink first**, so Outlook's first send/receive finds it: `host/Get-MailSinkMedia.ps1 -Execute` stages the pinned Inbucket release on the host (`MEDIA.md`, "The mail sink"), `host/Copy-ToGuest.ps1` carries it in, then `guest/Install-MailSink.ps1 -ExpectedSha256 <hash> -Execute` registers it to start with the guest and ends by proving an SMTP-to-POP3 round trip - `SINK-READY`. Reboot and `-Verify` once more. Accounts: `guest/New-TierProfile.ps1`. **The POP3 password - MEASURED 2026-09-24 on `OAI-UNINDEXED`: the sink accepts none and any, but Outlook holding none PROMPTS and never connects**, so after the account exists, with Outlook closed: `guest/New-TierProfile.ps1 -StoreSinkPassword -Execute` - proven by the sink's debug log (`read USER tier`, `read PASS any-value`), runbook §2.7 - and again once the identity account exists (measured: its first start PROMPTS too). `SINK-READY` first time on that guest; after a graceful restart the sink started 7 s after boot. Identity account: `guest/Add-IdentityAccount.ps1`, phases in its banner - **built on both guests** (`CP-09-IDENTITY-ACCOUNT` on `OutlookAI-Indexed`, `CP-10-IDENTITY-ACCOUNT` on `OutlookAI-Unindexed`). The signature has a script, `guest/Set-AccountSignature.ps1` - **run on `OAI-UNINDEXED` 2026-09-24 it printed `Verified` and bound the signature to the identity PST's data-file entry instead of the account**, so that account still has no signature Outlook injects (runbook §2.8b; a product defect, not fixed here) | guest |
| 7b | **Decide whether this guest is the indexed one or the unindexed one, and do it BEFORE the corpus exists** | `guest/Set-OutlookIndexingDisabled.ps1` on `OutlookAI-Unindexed`; nothing on `OutlookAI-Indexed`. Order is the whole point: exclude Outlook first and no row is ever crawled, so there is nothing to purge and nothing to wait for. **Do not run its `-Verify` in the first minutes after a boot:** Windows Search is a delayed-start service, and on 2026-09-24 a `-Verify` whose first reading fell 74 s after boot found the catalog unreachable and said `NO-INDEXER` - a re-run once `WSearch` was Running said `UNINDEXED` | guest |
| 8 | Build the corpus | `guest/Build-Corpus.ps1` | guest, session 1 |
| 8a | **Build the fixture populations** into the hub, the bystander and - once it exists - the identity store: small, tagged, deterministic, and what the hub and index tests read | `OutlookAI.RemediationTools corpus-build --population hub\|bystander\|identity`, in the corpus profile, with every such store attached to BOTH profiles while it is still empty - `Docs/live-tier-on-the-vm.md` §3b has the order and the commands, and the hub is rebuilt before every run. **FIRST RUN, `OutlookAI-Unindexed` 2026-09-24: NOT CLEAN** (runbook §4.1, step 6; checkpoint `CP-11-POPULATIONS-BUILT-WITH-FAULTS`). The hub - the tier profile's own default store - built right: 56 items in its real folders. But every build exited 1, because no received item carries a resolved owner recipient; a PST attached with `AddStoreEx` has no Inbox or Sent Items, so the bystander's and identity store's "Inbox" items went into the PST's hidden root; and failed probe rungs left 12 items in the account-less profile's DEFAULT store (Corpus B's Drafts), which no `--allow-store` named. **Do not build a population into a PST that is not a minted default store until those are fixed** | guest, session 1 |
| 8b | **Make the guest able to RUN the suite at all** - it has no .NET, no git and no clone. Stage on the host, then install: `host/Publish-LiveTierPayload.ps1` then `guest/Install-DotnetSdk.ps1`. The SDK alone is not enough; the source and an offline NuGet feed travel with it. **`TEST-READY` on `OutlookAI-Unindexed` too, 2026-09-24** (runbook §4.1 step 7, `CP-12-SDK-TEST-READY`): 2,794 tests discovered, 17 run, 148 s end to end, and again from a new session - and the guest-built test assembly's `McpServerExePath` names a server exe that exists, so the T3 path question needs nothing more | host, then guest |
| 9 | Write the live-test settings file | **render it, do not write it by hand**: read the guest's store names over COM into its `liveTestSettings` section of `testbed.json` - the watched list, and on the indexed guest the INDEXED list in the order hub, bystander, corpus - then `host/New-LiveTestSettings.ps1 -VMName <guest>`, then the `host/Copy-ToGuest.ps1` line it prints. It refuses while any value is still a placeholder, naming each. `probeTerm` and three of `subjectOnlyProbe`'s fields are already filled in: they are the population generator's own constants. Read §3b first, or the tier refuses to start. **Done for `OutlookAI-Unindexed`, 2026-09-24** (§4.1 step 8, `CP-13-LIVE-SETTINGS`): three watched stores, one bystander, no index, no corpus, the sink on loopback | host, then guest |
| 9b | **Run the tier - and opt in, for that run and this machine** | Through `guest/Register-InteractiveTask.ps1`, with `$env:OUTLOOKAI_LIVE_OPT_IN` set to THIS guest's computer name inside the script it runs, right before `dotnet test`. Without it every live test refuses at its fixture and nothing touches Outlook. The exact lines are in §4c; the filter is the runbook's (`Docs/live-tier-on-the-vm.md` §4) | guest, session 1 |
| 10 | Take the measurements | `guest/Invoke-GuestMeasure.ps1`, `guest/Measure-SweepCost.ps1` | guest, session 1 |
| 11 | Get the results out | `host/Copy-FromGuest.ps1` | host |

Every script takes `-WhatIf`-style caution seriously: the ones that write take an explicit
`-Execute`, and print what they would do without it.

**Why the tier profile comes first - both orders were run from `CP-02`, 2026-09-24, on
`OAI-UNINDEXED`.** Until then this table put the corpus profile first; the guests were in fact
built tier-first, and only one of the two orders works as written:

- **Tier first (steps 4b-4f as above): every step passed as written.** The import at Outlook's
  first-ever start bound the account in that same start - COM read `DeliveryStore` as the minted
  `C:\OutlookAI-Tier\Outlook.pst`, Drafts resolving - and every `-Verify` passed, the corpus
  profile's included, with its `-Store` naming only the store it was asked to add.
- **Corpus first: three things went wrong, none of them loudly until a `-Verify` ran.** (1) `/PIM`
  on a machine with no profile at all also made an empty profile called `Outlook` - and made
  *that* the default. It cannot be scripted away: deleting a profile has no free route. (2) The
  corpus profile's `/PIM` store landed in `Documents\Outlook Files`, and
  `New-OutlookProfile.ps1 -Verify` failed on it as a stray unless its `-Store` named that store
  too. (3) The tier import happened at Outlook's **second** start - the one that shows the
  one-time "Check out our new look" dialog - and the POP3 account came up **unbound** (`DeliveryStore`
  NULL, `New-TierProfile.ps1 -Verify` failing on it). It bound at the next start, with nothing
  else changed. Whether the dialog is the cause is inferred from one run of each order; that the
  binding waits for a later start is measured.

**If a rebuild cannot give the tier import Outlook's first start** - a guest where Outlook has
already run - restart after the import, start Outlook once more, and re-run
`New-TierProfile.ps1 -Verify`: it fails, by name, until the account is bound.

---

## 1b. Steps 1 to 3 need nobody: the Windows install is unattended

**Three commands and a wait.** Per guest:

    pwsh -File Testbed/host/New-AnswerFile.ps1 -VMName OutlookAI-Indexed
    pwsh -File Testbed/host/New-TestbedVm.ps1  -Name OutlookAI-Indexed `
        -IsoPath      .work\media\Win11_25H2_EnglishInternational_x64_v2.iso `
        -AnswerIsoPath .work\testbed-answer\OutlookAI-Indexed\OutlookAI-Indexed-unattend.iso `
        -Execute -Start
    pwsh -File Testbed/host/New-TestbedVm.ps1  -Name OutlookAI-Indexed -CompleteInstall -Execute

The first builds a small ISO carrying `autounattend.xml` and the first-logon script. The second
creates a Generation 2 VM with Secure Boot and a vTPM, attaches the Windows ISO and the answer
volume as two DVD drives, and boots it. The third waits for the first logon to finish - about
seven minutes on the 2026-09-15 build - then ejects both discs, takes `CP-01-WIN-CLEAN` with no
disc in it, and deletes the answer ISO (step 3a; the password paragraph below says why in that
order). Run it with `-RepoRoot <the main checkout>` from a worktree: it reads the first-logon log
over PowerShell Direct, with the credential that lives only there. **The third has not run yet** -
it was written 2026-09-24 by an agent forbidden to touch Hyper-V, and only its `-SelfTest` (28
assertions, both shells) has executed; run it without `-Execute` first, which reads and reports and
changes nothing. Nothing else is typed. Do it twice, with the two names,
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
**Since 2026-09-24 it refuses to run anywhere but a guest** - logged on as `vmadmin` AND named
`OAI-*`, the answer file's own account and naming convention - and it checks before it opens that
log, so **a guest with no `first-logon.log` at all is one where it refused**, or - less likely - died
before its first line (if the script is on no drive, the answer file's command writes a log saying
so instead). Run it by hand there to read which.
A guest built outside the convention - `New-AnswerFile.ps1 -ComputerName` without the `OAI-`
prefix, or a VM name not starting `OutlookAI-` - refuses at first logon by design.

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
generated ISO is disposable; the template is the record. `New-AnswerFile.ps1` refuses to replace an
ISO that is still there - on this route one survives only an unfinished build or a guest whose
checkpoints reference it - so pass `-Replace` once nothing does, or build elsewhere with `-OutDir`.

**The credential never lands in this repository.** The committed template carries placeholder
tokens where the password goes. The generator substitutes them and writes only into gitignored
`.work/`; it refuses an output path under `Testbed/`, and refuses any path inside the repository
that is not under `.work/`. `.github/scripts/check-testbed-references.ps1` check 7 fails the build
if the template ever stops holding placeholders, or if a filled `autounattend.xml` is ever
tracked. **The generated ISO holds the password in clear text** - it is in scratch, keep it there,
and let step 3a delete it: `New-TestbedVm.ps1 -CompleteInstall` removes it once Windows setup has
consumed it, AFTER both discs are out and `CP-01-WIN-CLEAN` is taken without them. Not by hand, and
not before: a checkpoint taken with the disc attached references the ISO, and a restore of that
checkpoint needs the file back.

**DO NOT DELETE THE TWO ANSWER ISOs ALREADY IN `.work/testbed-answer/`** -
`OutlookAI-Indexed-unattend.iso` and `OutlookAI-Unindexed-unattend.iso`, built 2026-09-15. Those two
guests were built before step 3a existed: each one's `CP-01-WIN-CLEAN` was taken with its answer
disc attached and references the file (read off both guests by the coordinator 2026-09-24; the
change that wrote this paragraph did not touch a VM), so deleting either breaks a restore of that
checkpoint, and of any later one taken with the disc still in. They stay, password and all, until
the from-scratch rebuild replaces those two VMs. Until then `-CompleteInstall` is written to refuse
on those guests without changing anything (its first check; like the rest of it, never yet run
against a VM), and `New-AnswerFile.ps1` refuses to replace either ISO unless told `-Replace`. Once the guest an ISO belongs to has been deleted, with its checkpoints, nothing
references the file and it can go: a successor built under the same name replaces it (`-Replace`)
and its own step 3a deletes what replaced it, and one built under another name leaves it for you to
delete then.

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

Reconciling the plan against the census taken straight after the build. **Both differences this
section used to describe are now fixed, so the expectation is that plan and store MATCH EXACTLY** -
the middle column is history, kept because each number is what makes the fix checkable:

| Folder | Plan | Store, 2026-08-19 | Expect today |
| --- | --- | --- | --- |
| Inbox | 10,912 | 10,912 | 10,912 |
| Sent Items | 4,964 | 4,964 | 4,964 |
| Junk Email | 1,663 | 1,663 | 1,663 |
| Deleted Items | 2,461 | 2,467 (+6 probe items) | **2,461** |
| Outbox | 0 | 2,761 (+2,761, all unread) | **0** |

**The Outbox residue** is the known `MSGFLAG_SUBMIT` defect, and 2,761 is the second independent
confirmation of its identity: it is *exactly* the plan's unread count, not approximately. The
first confirmation was the 40,000-item build's 5,532. The build now clears `MSGFLAG_SUBMIT` on
every item, so a rebuild today should leave the Outbox empty - **and if it does not, that is the
signal that the fix did not take**, because the count is predictable in advance.

**The Deleted Items residue was the probes' own litter, and the figure this table used to print
was too low.** Each rung of the placement ladder (4 rungs) and the date ladder (2 rungs, plus up
to 2 compensated re-tries) creates one throwaway item and deletes it in a `finally` - and
`MailItem.Delete()` is a SOFT delete, so every "deleted" probe item became a permanent resident of
Deleted Items under an EntryID no manifest records. **One `corpus-build --execute` therefore left 6
to 8 of them, and following the committed `Testbed/guest/Build-Corpus.ps1 -Execute` left 12 to
16** - that script runs a standalone `corpus-probe --execute` of its own and the build then probes
again. The +6 in the table is from the 2026-08-19 hand-run, which did not run the standalone probe.
A probe interrupted between the item's save and the capture of its EntryID additionally stranded
one item in **Drafts** permanently, because the delete was gated on having an id.

**Since 2026-09-16 the expected figure is 0.** Both probes purge their own residue - once at the
start of each pass, which is what heals a store built before this change, and again after each
item's soft delete - selecting by the reserved probe ordinal and deleting under the same two-key
rule as a teardown. The census counts probe items in their own number, excludes them from every
corpus statistic, and gives them their own sentence, so a non-zero count is reported as
`N throwaway probe item(s) were left behind` rather than as a duplicated ordinal. **A non-zero
count today means a store built before 2026-09-16, or a probe killed mid-run; re-running
`corpus-probe --execute` clears it.**

**"All unread" was a property of the tooling of the day, and is no longer one.** On 2026-08-19 the
probes wrote a read state only on the two rungs that clear `MSGFLAG_UNSENT`, so four of the six
items should have been READ and two UNREAD - which already does not match "all unread". Since
2026-08-24 both probes write `MSGFLAG_READ` on **every** probe item and save it, so a probe item is
created read. **What is NOT established** is whether that bit survives the soft delete into Deleted
Items: nothing on the host can settle it, no run since has recorded it, and a mis-recorded
observation fits the evidence just as well. It is moot while the expected count is 0.

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
is an operator action run from the accountless profile. But a guest names its corpus store in
`expectedStoreDisplayNames` - that is the list `outlook_health` and `list_accounts` read - and every
non-hub entry of that list is inside the identity-draft grant unless something says otherwise. Left
undeclared, two different code paths write into the measurement corpus:

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

**What the tier refuses, and what it does not - corrected 2026-09-24 (Q77).** An earlier version of
this section and of the runbook said a store named in only one of `expectedStoreDisplayNames` and
`bystanderStoreDisplayNames` refuses the tier. The code does not do that, and the maintainer chose to
correct the documents rather than the loader. A store declared in `bystanderStoreDisplayNames` alone
is still censused - the census adds every declared bystander back in, on purpose
(`T1/TripwireBystanderStoreTests.TheCensusWatchesEveryDeclaredBystanderEvenOneNoOtherListNames`) - and
a store named in `expectedStoreDisplayNames` alone is inside the identity-draft grant, which is
exactly the identity account's shape. The tier refuses only when no watched store is left that is
both non-hub and denied every write. `host/New-LiveTestSettings.ps1` still holds a guest to naming
every declared bystander in `expectedStoreDisplayNames` as well, because that list is what
`outlook_health`'s reachability check and `list_accounts` read.

**Two lists since 2026-09-24 (Q70): WATCHED and INDEXED.** `expectedStoreDisplayNames` used to be
both the stores the census watches and the stores every index test demands an index scope of, and
the identity store - watched, written to, holding almost nothing - could not be described
truthfully by one list. So `indexedStoreDisplayNames` now says which stores the index tier
measures, in order: **the hub first, the bystander second, the corpus last**, because several index
tests read entry 0 and the exclude-subfolders measurement reads the first non-hub entry. Absent, it
means the watched list, which is what every older settings file meant. On the unindexed guest it is
empty, and every index test there refuses rather than iterate nothing. The renderer holds the order,
and requires every indexed store except the corpus to be named as an `.invalid` address: the product
finds a small store in the index only through mail addressed to it, and only for a name shaped like
an address. `Docs/live-tier-on-the-vm.md` §1.3 and §2.10 have the whole of it.

**The bystander, the hub and the identity store are POPULATED BY THE GENERATOR** -
`corpus-build --population bystander|hub|identity`, 300, 56 and 8 tagged items, deterministic from a
seed recorded in `testbed.json` under `corpusIdConvention.populations`. The bystander is named as an
address too (`bystander@vm.invalid` in the example). The old objection - that the generator tags
everything and a bystander must be untouched - does not hold: the tag is `[OutlookAI-Corpus]`, which
no artifact sweep can select, and no test writes to the store either way. `Docs/live-tier-on-the-vm.md`
§3b is the procedure, and says what only a guest can answer.

---

## 4. Credentials

**Never in this repository. It is public, and a guest password has already been published from
it once** - it survives in git history and had to be rotated.

| What | Where it lives | How to create it |
| --- | --- | --- |
| Guest account password (PowerShell Direct, autologon) | `McpServer/OutlookAI.McpServer.Tests/live-fixtures/vm-credentials.json`, gitignored | Set it when you create the account. Set it to **never expire**: a maximum password age silently breaks the tier and recreates this problem. |
| Live-test machine coordinates (store names, manifest path, sink ports) | `McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json`, gitignored | **On a test guest:** rendered by `host/New-LiveTestSettings.ps1` from `Testbed/live-test-settings.template.json` and the guest's `liveTestSettings` section of `testbed.json` - committable only because a guest's stores are synthetic. **Anywhere else, the maintainer's machine above all:** copy `Testbed/live-test-settings.example.json` and fill it in by hand. The renderer never writes into any `live-fixtures` directory on the host. |
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

**THE FIRST VERSION OF THESE SCRIPTS WAS BUILT ON A CALL THAT DOES NOT WORK HERE, and finding
that out cost one guest run.** All three profile scripts went through Extended MAPI's
`IProfAdmin`. The first time any of that code executed anywhere - 2026-09-16, on `OAI-UNINDEXED`,
Office LTSC 2024 16.0.17932.20884 - it threw `E_NOINTERFACE` on `IID_IProfAdmin`, with
`MAPIInitialize` and `MAPIAdminProfiles` both succeeding and only the QueryInterface failing. The
cause has **not** been established. All three have since been moved onto routes that do work; the
shared MAPI interop they stood on is deleted, and there is **no MAPI call left anywhere in this
repository**.

**The lesson, because it is the expensive half:** those scripts carried "verified by PARSING"
banners and they parsed perfectly. A script verified by parsing is a script whose syntax is
verified. What replaced that is a `-SelfTest` switch on each one - pure decision functions driven
against synthetic inputs, reading no registry and making no COM call, so they run on any machine
including the maintainer's workstation. That still does not prove a guest run: each self-test ends
by printing the list of things only a guest can settle, and those lists are the honest statement
of what is unproven.

> **IF YOU EXTEND THOSE SELF-TESTS, DO NOT ASSERT WITH `-like`. USE `String.Contains`.**
> PowerShell's `-like` treats `[...]` as a **character class**, not as literal brackets - so
> `$prf -like '*[Account1]*'` asks "does this text contain any one of `A c o u n t 1`", which is
> true of very nearly any text. **Every section header in a `.prf` is bracketed**, so this hits the
> most interesting assertions and hits them silently: the test passes, and it passes for the wrong
> reason. Four assertions in `New-OutlookProfile.ps1` were written with `-like` and did exactly
> that before they were caught.
>
> It is the same wildcard trap as `CLAUDE.md` mailbox-safety rule 2, which exists because
> `-like "*[tag]*"` once matched nearly every subject in a real mailbox and destroyed real mail.
> Here it costs a false green rather than data - but it is the same operator, the same
> misunderstanding, and worth recognising on sight. `Contains` is ordinal and case-sensitive, which
> is what an exact-name assertion wants anyway.

**Five things. Four are automatable, one is not, and two of the four now rest partly on inference
rather than on measurement** - marked below, because a reader must be able to tell which claims
have been run and which have been reasoned.

| What | How | Where |
| --- | --- | --- |
| Profile with **no mail accounts** - the corpus profile, which `corpus-build` requires | **`outlook.exe /PIM <name>`, then `guest/Add-OutlookPstStore.ps1` per named store** - the unattended route, [MEASURED] end to end on 2026-09-24 from `CP-02` (`accounts=0`, no dialog, every store under its exact name) and how both guests' corpus profiles were made. `/PIM` names its own store `Outlook Data File` and does **not** change the default profile; run it after the tier profile exists, or on a machine with no profile it also makes - and defaults to - an empty profile called `Outlook` (§1). The **`.prf` import** (`guest/New-OutlookProfile.ps1 -Execute`) is NOT the unattended route: [MEASURED] 2026-09-24 it imports within 5 s and makes the profile and its PSTs, and the account-less profile it makes then stops on "Email Account Setup" at every start, so `-Execute` refuses without `-AcceptAccountWizard`. ~~`IProfAdmin::CreateProfile`~~ is measured broken here | `guest/Add-OutlookPstStore.ps1` |
| **PST with an exact display name**, `@` included | Two routes, both off MAPI. **At profile creation:** the same `.prf`'s `[ServiceN] Name=`, which maps to `PT_UNICODE,0x3001` - `PR_DISPLAY_NAME_W`. [MEASURED]: `Store.DisplayName` read back over COM as exactly `tier@vm.invalid`. **Into a profile that already exists:** `NameSpace.AddStoreEx` [MS-DOC] then a **root-folder rename** [MEASURED, `guest/Rename-OutlookStore.ps1`]. [INFERRED] only for more than one PST in one `.prf` - `UniqueService=No` is the documented key that permits it, and it has not been run with two. ~~`IMsgServiceAdmin::ConfigureMsgService`~~ is behind the broken gateway | `guest/Add-OutlookPstStore.ps1` |
| **Default-profile switch**, no prompt | **`HKCU\...\Outlook\DefaultProfile` (REG_SZ) + `PickLogonProfile`** - NOT `IProfAdmin::SetDefaultProfile`, which is measured broken here (2026-09-16: `E_NOINTERFACE` on `IID_IProfAdmin`, Office LTSC 2024 16.0.17932.20884). The registry value is measured WORKING on the same guest the same day | `guest/Set-DefaultOutlookProfile.ps1` |
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

**Why Extended MAPI WAS chosen, and why the argument collapsed.** `IProfAdmin` and
`IMsgServiceAdmin` are documented and supported, and the registry layout underneath them is
reverse engineering with no published end-to-end recipe on 16.x - so the API was the contract and
the registry was only the storage. The stronger reason was specific to this testbed: the PST
provider's configure call applies `PR_DISPLAY_NAME` **when the store is created**, and that was
believed to be the only route to a store named exactly what we asked for, because
`Store.DisplayName` is read-only in the object model and **`Namespace.AddStoreEx` adds a PST
perfectly well but cannot name it**.

That last sentence is still true and **no longer matters**, which is what unblocked all of this.
`guest/Rename-OutlookStore.ps1` measured, on this build, that `Store.DisplayName` FOLLOWS a rename
of the store's root folder - `@` included, and without breaking an account's delivery-store
binding. Four research passes had failed to settle that either way in any source, Microsoft's or
the community's. So the store can be added unnamed and named afterwards, and the one capability
that justified reaching for MAPI is reachable without it.

**A preflight that passed, and proved less than it looked.** Measured on a guest running
16.0.17932.20996: `DLLPathEx` resolves to a real `msmapi32.dll` under the Click-to-Run `root\VFS`
tree, Office is x64 and PowerShell is a 64-bit process, so Extended MAPI loads and the bitness
matches. `DLLPath` being an unresolvable bare filename beside it is the **healthy** shape on
Click-to-Run, not a fault. All of that is still correct - and every bit of it was true on the
machine where `MAPIAdminProfiles` then failed. **It proved the DLL resolves. It did not prove the
interface is reachable**, and nothing short of calling it would have.

**`New-OutlookProfile.ps1 -Preflight` still exists and now checks the route that shipped**: the
Office hive, the `Setup` key, the profiles already present, the current `ImportPRF` value, whether
`FirstRun`/`First-Run` are absent, and whether Outlook is running. It reads only - no COM call and
no MAPI call, because a preflight that could itself start Outlook, or itself hang, would be the
fault it is checking for. Run it first; it costs nothing and needs no checkpoint.

**The guard that keeps these off the wrong machine.** Every one of them refuses unless the session
is logged on as `vmadmin`, the guests' autologon account (§2). It is not silenceable by a flag:
the only way past it is `-ExpectedUser <name>`, which is a thing nobody does by accident. The
second Windows account of `Docs/live-tier-on-the-vm.md` §2.4 will need exactly that.

**Until 2026-09-24, "every one of them" was not true.** `guest/Rename-OutlookStore.ps1` - whose
banner claimed an identity check it did not have - `guest/Set-AccountWizardClassic.ps1` and
`guest/Set-OfficeFirstRunSuppressed.ps1` wrote with no guard at all; run on the maintainer's
workstation, the first renames a real mailbox's store. All three now call it before anything
else, and on the workstation each refused with every registry key it names unchanged, last-write
time included, and no Outlook COM object requested. **`.github/scripts/check-testbed-references.ps1`
check 9** now fails the build when a guest script that writes does not call the guard, or calls it
after its first write. The four writers outside that change - `guest/Build-Corpus.ps1`,
`guest/Complete-FirstLogon.ps1`, `guest/Measure-SweepCost.ps1` and
`guest/Register-InteractiveTask.ps1` - were declared in the check by name and guarded later the same
day, each refusal proven on the workstation the same way, so the check's declared list is empty and
it enforces every writer it recognises. `Complete-FirstLogon.ps1` carries a restated guard, because
the answer volume carries it alone. The list stays a ratchet: adding an entry is a visible edit.

### 4b-i. `ImportPRF` could rebuild the profile at every Outlook start - Outlook clears it, and it is cleared anyway

**Read this before building a corpus on a profile made by `New-OutlookProfile.ps1`.** It is a
data-shaped hazard on a machine whose entire purpose is a 20,000-item corpus that takes about
thirteen minutes to rebuild.

`ImportPRF` is a registry value under `HKCU\...\Outlook\Setup`. Outlook reads it **at every
start**, and the `.prf` this project writes carries `OverwriteProfile=Yes` - which is correct for
a rebuild, because it is what makes a repeat import converge instead of producing a
`Backup Of <name>` profile.

**WHETHER OUTLOOK CLEARS THE VALUE AFTER PROCESSING IT: MEASURED - IT DOES, WITHIN ~5 s.** On
`OAI-UNINDEXED`, 2026-09-24, Office 16.0.17932: sampled every 5 s through a plain first start and
through a `/PIM` first start (`guest/New-OutlookProfile.ps1` imports), and again through the tier
import of the §1 rehearsal (`guest/New-TierProfile.ps1`), the value was gone at the first sample
each time, with the new profile already listed. The tier profile built on 2026-09-15 had already
answered it after the fact: set at 19:54:37, imported at 19:54:43, absent ever since, while
nothing in the repository or the scratch that drove that guest removes it. This used to read
"unknown - nothing in this repository has ever checked".

**What it would have cost if it had not**, which is why it is still guarded: every later Outlook
start would re-import the file and REBUILD the profile from it (`OverwriteProfile=Yes`). No
`.pst` would be deleted - nothing here deletes a data file - but a store attached or filled after
the import would **stop being part of the profile**; on the corpus guest, a corpus that has
apparently vanished, and a rebuild.

**Cleared anyway, with nothing to remember** (the maintainer's call: do not make the protection
depend on the answer, or on a flag). `New-OutlookProfile.ps1 -Verify` and `New-TierProfile.ps1
-Verify` each remove a lingering `ImportPRF` that names their own `.prf` - only once First-Run is
back, because removing it before Outlook has read it cancels the import - and
`guest/Build-Corpus.ps1` refuses to build while one is set, `-SkipPreflight` or not. The old
manual step, `.\New-OutlookProfile.ps1 -ClearImportPrf -Execute`, still exists and is no longer
part of the build (§1 step 5).

### 4b-ii. Where the evidence is

The findings behind all of this, with every claim labelled Microsoft-documented,
community-reported, guest-measured or inferred - and a list of what could **not** be established -
are in `Docs/research/profile-automation-research.md`. Read its §1 table first: it summarises the
answer per capability, and it now carries the measured-broken marker on the rows that named MAPI.

---

## 4c. Every live run opts in - for that run, on its own machine

**No `Category=Live` test starts unless the run sets `OUTLOOKAI_LIVE_OPT_IN` to the computer name
of the machine it runs on.** Decided 2026-09-24, after a targeted `dotnet test` whose name filter
lacked `Category!=Live` selected three live tests: they failed only because that worktree had no
settings file, and the same command in the main checkout would have run them - writes included -
against the maintainer's real mailbox. A settings file is permanent, so it cannot be what says a
run is MEANT to reach a mailbox; the opt-in says that, once per run. It is checked first thing in
`LiveTestSettings.Load()`, which every live fixture starts with, and `T1/LiveRunOptInTests`
proves from the compiled code that no live class - the stdio tier's included - can start without
it.

On a guest, the tier runs through `guest/Register-InteractiveTask.ps1` (session 0 cannot finish
starting Outlook), so the opt-in goes INTO the script it is handed, right before `dotnet test`:

```powershell
.\Register-InteractiveTask.ps1 -TimeoutSeconds 7200 -Script @'
$env:OUTLOOKAI_LIVE_OPT_IN = 'OAI-INDEXED'   # THIS guest's computer name - $env:COMPUTERNAME prints it
Set-Location C:\OutlookAI-Q5\src
dotnet test McpServer\OutlookAI.McpServer.Tests\OutlookAI.McpServer.Tests.csproj -c Release --filter "Category=Live&Requires!=DelegateStore"
'@
```

`C:\OutlookAI-Q5\src` is where step 8b expands the source, and `-c Release` is the tree its
`-Verify` already built (without it `dotnet test` builds a second, Debug, tree beside it); the
filter is the runbook's (`Docs/live-tier-on-the-vm.md` §4). Three rules, each refused with a
message that says why:

* **The value is the computer name, not "1" or "true"**, so an opt-in carried to another machine -
  a copied script, a roaming profile - opens nothing there. Spell it out rather than pasting
  `$env:COMPUTERNAME` into a script you might run somewhere else.
* **Set it in the session, for the run.** A value saved with `setx`, or in System Properties, is
  refused even on its own machine: it would open every later run there, the accidental ones
  included. Do not put it in a PowerShell profile either - nothing can detect that, and it defeats
  the point the same way.
* **It changes nothing else.** It says a live run was intended; it makes no test read-only and
  chooses no tests. **The maintainer's workstation is read-only for live tests, always** (CLAUDE.md,
  Mailbox Safety); how that is enforced in code is the maintainer's open decision (Q74), not this.

Without the opt-in, every live collection fails at its fixture with `LIVE TEST REFUSED` and nothing
touches Outlook. `--list-tests` still discovers everything without it, because discovery never
builds a fixture.

---

## 5. What is in here

| Path | What it is |
| --- | --- |
| `testbed.json` | The parameter set. Corpus quad, expected plan output, build cost, guest layout, and an explicit list of what is still unrecorded. Its `vmName` is the guest this was **measured on**, not a guest to build (§3). `corpusIdConvention` also records the indexed guest's **minimum corpus size, 160,000 items** - a requirement, so the index tier's latency bounds are tests there - and, under `populations`, the ids and seeds of the generated hub, bystander and identity populations (none built yet). Its `liveTestSettings` section holds each guest's live-test settings values, keyed by VM name - **placeholders until they are read off the guest over COM**, which nobody has done yet - except the values that are recorded decisions, which `_decided` lists: the probe term and three of the subject-only probe's fields are the population generator's own constants. |
| `live-test-settings.example.json` | Complete example of the gitignored settings file, every field present, placeholders only - the watched and the indexed list, the probe term and the subject-only probe as a test guest carries them. Every address in it is under `.invalid`, and `.github/scripts/check-testbed-references.ps1` check 8 fails the build if one ever is not. |
| `live-test-settings.template.json` | The same shape as the example, holding **tokens only**: every value is a double-brace token spelling its own JSON path in a guest's `liveTestSettings` section. `.github/scripts/check-testbed-references.ps1` check 8 fails the build if its fields stop matching the example's or a value in it stops being a token, and `T1/LiveTestSettingsTemplateTests` renders it with synthetic values through the live tier's own loader. |
| `host/New-LiveTestSettings.ps1` | Renders one guest's gitignored `live-test-settings.json` from that template and the guest's section of `testbed.json`, into `.work/`. `-VMName` is mandatory (§4a). **Refuses** while any value is still a placeholder (naming each), anything the tier would refuse at start, and the documented rules the tier does not itself enforce - a bystander must also be in `expectedStoreDisplayNames`, the corpus store must be a bystander, the hub must be shaped as an address; on the indexed guest the indexed list must run hub, bystander, corpus, with every entry but the corpus named as an address, and carry the generator's probe term and a complete subject-only probe on the hub; on the unindexed guest the indexed list, the probe term and the probe block must be empty or absent; every address must be under `.invalid`; and a corpus below the size `corpusIdConvention` requires is refused. **Never writes into a `live-fixtures` directory on the host**, however the path is spelled - checked first, on the resolved path and again on the path Windows reports - and inside a git working tree only under `.work/`. Prints the `Copy-ToGuest.ps1` line for the guest's destination, and what it could not check: that the guest's tier profile actually mounts every store it names. **Never rendered a real guest's file** - their values are still placeholders; `-SelfTest` is 168 assertions over its decisions, 0 failures, under Windows PowerShell 5.1 and PowerShell 7 (2026-09-24). |
| `guest/autounattend.template.xml` | The unattended-install answer file. Locale, disk layout, local account, autologon - and placeholder tokens where the password goes. Contains no credential and must never contain one. |
| `host/New-AnswerFile.ps1` | Fills that template from the gitignored credential and packages it as a small ISO. Writes into gitignored scratch only, and refuses anywhere else. `-VMName` is mandatory (§4a). The ISO holds the guest password in clear text; step 3a deletes it once setup has consumed it. **Refuses to replace an ISO already at its output path** unless `-Replace` (2026-09-24): on this route one survives only an unfinished build or a guest whose checkpoints reference it - both current guests' do (§1b) - and replacing deletes the old file first. That refusal was run on the workstation against a stand-in file: it refused before reading the credential or writing anything. |
| `guest/Complete-FirstLogon.ps1` | The first-logon fix-ups the answer file cannot express: the en-NL language list, the home location, the locales, no sleep, no fast startup. Logs and reads back everything it set. **Guarded since 2026-09-24** by a restated two-axis guard - `vmadmin` AND a computer name starting `OAI-`, the answer file's own convention - because the answer volume carries it alone; it refuses before opening its log, so a refusal leaves a guest with no `first-logon.log` at all. Refusal proven on the workstation (no arguments and `-SkipPower`: nothing written, every setting it names unchanged); its guard's inputs were measured at a real first logon on 2026-09-15 (`on OAI-UNINDEXED as vmadmin`), but the guarded script has not yet run at one. |
| `host/New-TestbedVm.ps1` | Creates the Hyper-V guest, attaches the Windows ISO and the answer volume, boots it, and records the spec it chose. `-Name` is mandatory (§4a). **`-CompleteInstall`** (step 3a, 2026-09-24) finishes it: waits for the first-logon log's DONE line, ejects both discs, takes `CP-01-WIN-CLEAN` with no disc in it, and deletes the answer ISO once no VM or checkpoint on the host references it; it refuses on a guest whose `CP-01-WIN-CLEAN` already holds a disc - both current guests - changing nothing. **Never run** against a VM; `-SelfTest` covers its decisions, 28 assertions in both shells. |
| `host/Publish-GuestPayload.ps1` | Publishes the MCP server and the remediation tools on the host and zips them for copy-in. Host-only, so it takes no VM name; one payload serves all three guests. |
| `host/Get-GuestCredential.ps1` | Loads the guest credential from the gitignored fixtures directory. Documents the one place a credential may live; contains none. `-VMName` is mandatory (§4a). |
| `host/Copy-ToGuest.ps1` | Copies a file or a zip into the guest over PowerShell Direct. `-VMName` is mandatory (§4a). |
| `host/Copy-FromGuest.ps1` | Gets results, logs and the corpus manifest back out. `-VMName` is mandatory (§4a), and it also names the subdirectory results land in. **The manifest goes to the shared root** - safe because each guest has its own corpus id, and kept there so a reused id still collides visibly; **`measure.jsonl` and the logs go to `<Destination>\<VMName>\`**, because what differs about a transcript is the machine that produced it (§3). It refuses to replace any pulled file whose content differs, unless `-Force` says you mean it. |
| `guest/Register-InteractiveTask.ps1` | The session-1 scheduled-task recipe. Everything COM-touching goes through it. **Guarded since 2026-09-24** (`Assert-TestbedGuest`, so `guest/OutlookMapiInterop.ps1` must sit beside it - §1 step 4a): on the workstation it refused before creating a job directory or touching a task; for `vmadmin` on a guest the guard returns silently and every line after it is unchanged, which has not yet been run on a guest. |
| `guest/OutlookMapiInterop.ps1` | **The name is historical: there is no MAPI in it any more.** It held ~550 lines of C# Extended MAPI interop - `IProfAdmin`, `IMsgServiceAdmin`, the table readers - whose central call was RUN 2026-09-16 AND FAILED (`MAPIAdminProfiles` returned S_OK; the QueryInterface for `IID_IProfAdmin` returned `E_NOINTERFACE`, Office LTSC 2024 16.0.17932.20884). It is **deleted** rather than kept as a fallback, because an unexercised second route is the bug this file was just bitten by - and every script dot-sourcing it was paying an `Add-Type` compile for code nothing could call. It is one `git show 8b610c2` away. What remains is the shared guest layer: the `vmadmin` guard, the PST path normaliser, and one Outlook COM session helper whose release path is shared rather than copied. **Its REOPEN section records the ten-line experiment, RUN 2026-09-24 on `OutlookAI-Indexed`: the hypothesis is REFUTED.** The interop had declared `IProfAdmin` as `{00020379-...}`; Microsoft's `MAPIGuid.h` has `IID_IProfAdmin` = `{0002031C-...}`, and on this build `QueryInterface` for `00020379` returns `0x80004002` while `0002031C` returns `S_OK` with the same pointer. The `E_NOINTERFACE` was the declaration, not Office - `IMsgServiceAdmin` and `IMsgServiceAdmin2` were declared with wrong IIDs too. Nothing was rebuilt on it (knowledge only, by decision). **Dot-sourced, never run directly.** |
| `guest/New-OutlookProfile.ps1` | Creates an Outlook profile with no GUI by writing a `.prf` and pointing `ImportPRF` at it, each PST under an exact display name; also `-Preflight` (read-only), `-Verify` (`-WithOutlook` adds the COM read), `-ClearImportPrf` and `-SelfTest`. **RUN 2026-09-24 on `OAI-UNINDEXED`, from CP-05. The import WORKS** - `ImportPRF` consumed within 5 s, the profile created and made default, every named PST created (two in one `.prf`: the repeated service block is confirmed) - **and the account-less profile it makes does NOT open unattended**: Outlook stops on "Email Account Setup" at every start (First-Run absent, and put back) and COM reads "You are not connected". So `-Execute` now **refuses unless `-AcceptAccountWizard`** and names the unattended route, measured end to end the same day: `OUTLOOK.EXE /PIM <name>`, then `guest/Add-OutlookPstStore.ps1` per store - which `-Verify -WithOutlook` then passes. The run also found `-Verify` dead on its first line (a script-scope `foreach ($store ...)` was the `[string[]]` `-Store` parameter; `-SelfTest` now reads its own syntax tree for that), a stray-PST scan blind to `ForcePSTPath`, and a mail-account count that counted the profile's data files. `-Verify` removes a lingering `ImportPRF` once the import has run (§4b-i). `-SelfTest`: 154 assertions, 0 failures. |
| `guest/Add-OutlookPstStore.ps1` | Adds a PST to the profile Outlook is running, with an **exact** display name: `NameSpace.AddStoreEx` and then a rename of the store's root folder. Needs Outlook **running**, works on the profile Outlook has open and refuses if that is not the one named, matches stores on `FilePath`. `-NameProbe` is **gone** (§6 item 10 answered it). **RUN 2026-09-24 from CP-05, every path, unchanged: it works** - `AddStoreEx` returned at once (no spin), `Store.DisplayName` followed the rename, `@` included; a re-run is a no-op, a new name renames without a second attach, a taken name and the wrong profile are refused before anything runs; and into a fresh `/PIM` profile it made the corpus-profile shape `guest/New-OutlookProfile.ps1` now points at. `-SelfTest`: 39 assertions, 0 failures. |
| `guest/Set-DefaultOutlookProfile.ps1` | Switches the default profile and switches the profile prompt off, through the HKCU `DefaultProfile` and `PickLogonProfile` values (the `IProfAdmin` route is measured broken, see the interop row). Closes §6 item 5. **RUN 2026-09-24 from CP-05 and it works end to end**: a profile that does not exist is refused with the key's own last-write time unchanged, a running Outlook is refused, both values read back, and after a restart Outlook opened the named profile with no prompt - confirmed over COM. One output fix: an `Office\8.0\Outlook` key holding only `First-Run` was offered as "`-OfficeVersion 8.0`". `-SelfTest`: 64 assertions. Matters more than it looks: the corpus tool LOGS ON with the default profile and does not attach to whatever Outlook is running, so the wrong default makes a corpus build refuse. |
| `guest/Dump-UiaTree.ps1` | **Read-only.** Dumps the UIAutomation tree of an open dialog and prints a VERDICT: whether Outlook's account wizard is a classic Win32 property sheet (addressable by locale-invariant numeric `AutomationId`) or Office's own DirectUI chrome (no stable ids - dead for a PowerShell client). Two minutes, and it decides the whole GUI-automation route. **Never executed.** |
| `guest/Rename-OutlookStore.ps1` | Renames a store to an exact display name, which the tier profile needs because Outlook names the store it mints 'Outlook Data File' and section 2.6 requires the hub store to be named as an SMTP address. **MEASURED WORKING 2026-09-15** - and it settles a question no documentation could: `Store.DisplayName` is read-only, but it DOES follow a rename of the store's root folder, `@` and all, without breaking the account's delivery-store binding. **Guarded since 2026-09-24** - its banner claimed an identity check it did not have, so on the workstation it would have renamed a real store; it now calls the `vmadmin` guard first and needs `guest/OutlookMapiInterop.ps1` staged beside it. Refused on the workstation with no COM object requested; renamed the minted store as vmadmin in both §1 rehearsals, and is a no-op when run twice. |
| `guest/tier-profile-forcepst.prf` | **The one that works, and `New-TierProfile.ps1`'s default since 2026-09-24.** A PRF with no PST service and no `DefaultStore`: Outlook mints the POP3 account's delivery store itself under `ForcePSTPath` - which `New-TierProfile.ps1 -Execute` now writes - and a store Outlook mints is a store Outlook binds, the step a text file cannot perform. Imported on both guests 2026-09-15/16, and from `CP-02` twice on 2026-09-24 by the scripts alone: `DeliveryStore` and its Drafts folder resolve (§1 says which order binds at the first start). Its banner used to say it had never been imported. |
| `guest/Set-OfficeFirstRunSuppressed.ps1` | Suppresses Office's own first-run dialogs and pins CLASSIC Outlook. `ImportPRF` silences the profile wizard and nothing else - a guest with a perfect profile still came up on "Your privacy matters", and a dialog on an unattended guest is a hang, not a prompt. **Run 2026-09-24 from `CP-02`, twice: all 13 values written and read back; at Outlook's first start after it (every visible window enumerated, in one of the two runs) the only dialog was the POP3 password prompt.** It does not cover the one-time "Check out our new look" dialog at the **second** start (§1). Guarded since 2026-09-24 (it had written Office policy values on any machine). It had also run before, from hand-run scratch that then deleted the three values it writes inside the Outlook key - `CP-05` lacks exactly those (`guest/New-TierProfile.ps1` banner). **Repaired on `OAI-UNINDEXED` 2026-09-24**: `-Verify` showed exactly those three missing, `-Execute` put them back, and they read 13 of 13 again after two Outlook starts and exits (`CP-06-FIRSTRUN-REPAIRED`). |
| `guest/Set-OutlookProgrammaticAccess.ps1` | **Q80: auto-approves Outlook's programmatic-access prompts on an offline test guest.** Writes the documented per-user Outlook security Group Policy under `HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security` - `AdminSecurityMode` 3 ("use the security policy from the GPO settings") and eight `PromptOOM*` values at 2 ("Automatically approve"), every object-model category Microsoft documents (KB 926512; the Intune Office baseline reference says Outlook Security Mode must be set for any of them to apply) - rather than the machine-wide Trust Center value `ObjectModelGuard` (KB 3189806), which switches the guard off altogether and whose HKLM home is virtualised under Click-to-Run; the banner has the sources and the reasons. `-Execute` refuses while Outlook runs; `-Verify` is not a read-back: in session 1 it reads `Account.SmtpAddress` over COM from a child job inside a time limit, watches the session for a guard prompt, answers any prompt Deny so it never leaves a hang, and fails on it. Guarded on two axes (`vmadmin` and an `OAI-` computer name). **RUN 2026-09-24 on `OAI-UNINDEXED`**: a control with the values absent `PROMPTED` 0.7 s after the read; with them, `NO-PROMPT` (16 ms); every mode refused on the maintainer's workstation with its policy key unchanged. `-SelfTest`: 43 checks under 5.1 and 7. |
| `guest/New-PopAccountPrf.ps1` | A **spike**, not a route: the one free candidate for creating a POP3 account, plus the read-back that says how far it got. Expected to fail; §4b says why. **Never executed.** |
| `guest/Set-AccountSignature.ps1` | Gives the identity account its signature, by driving the shipped `manage_signature` tool rather than improvising. Runs **after** the accounts exist. **Never executed.** |
| `guest/Install-MailSink.ps1` | Installs **Inbucket 3.1.1** as the loopback sink the POP3 accounts point at, from a **staged** package - never a download - pinned by a mandatory, undefaulted SHA-256. A SYSTEM scheduled task starts it at boot through a generated launcher, because Inbucket takes its configuration from the environment only and is not a Windows service. `-Verify` **proves the round trip over raw sockets** rather than reporting open ports: an empty POP3 password accepted, a submitted message returned byte-intact through every dot-stuffing case and a base64 attachment, `TOP` working, one mailbox blind to another's mail, numbers fixed after `DELE`, deletes only at `QUIT`, nothing resurrected and no id reused after a restart. It writes to and deletes from two probe mailboxes of its own only - no account reads them - and opens the accounts' mailboxes read-only. Touches no Outlook, no MAPI and no mail item. **RUN 2026-09-24 on `OAI-UNINDEXED`: `SINK-READY` at the first `-Execute`**, every check passing, and again after a graceful restart with the sink started 7 s after boot; `-SelfTest` 104 assertions on the guest's 5.1. Its own log is written through a 4 KB buffer, so read it after more has been logged (its banner). `-Uninstall` not yet run on a guest. Replaces the smtp4dev version, which never ran. |
| `guest/Set-OutlookIndexingDisabled.ps1` | Takes a guest's Outlook **out of** the Windows Search index, which is the only thing that makes `OutlookAI-Unindexed` different from `OutlookAI-Indexed` - both are built from the same answer file and both come up indexed. It never disables the Windows Search **service**: a machine with no indexer is not a store with no index frontier, it is the product's "index unreachable" branch, and it removes `index.perStore[]` - the one instrument the runbook says to read. `-Verify` probes the **catalog** rather than reading registry values back, twice, and returns one of four verdicts; two of them are "not an answer" on purpose. **Run it BEFORE `guest/Build-Corpus.ps1`** (§1 step 8) or you also own purging what was already crawled. **RUN 2026-09-16 AND IT WORKED** - `VERDICT: UNINDEXED`, indexer still running. But it reported *no mapi rule existed to change*, so the **Group Policy layer did all the work and the registry half is still unexercised**; it said so rather than claiming success for a write it never made. It now also refuses while Outlook is running - changing the index state with a PST open is a documented route to an Office repair. **EXERCISED 2026-09-24 on `OutlookAI-Indexed`, and three things changed:** (1) the registry half **cannot write** - administrators hold ReadKey only on the crawl scope manager's keys, and the script's own write failed with a `SecurityException` from an elevated session - so that layer is now **read-only** and the policy value is the only write; (2) `OutlookAI-Indexed` has **never been indexed** (no `mapi16` rule, zero Outlook rows, and Outlook would not add its scope even with its UI open), a state the old `-Verify` called `SETTLING` and now calls **`NOT-IN-SCOPE`**; (3) so the 2026-09-16 `UNINDEXED` proves that guest has no Outlook rows, **not** that the policy excludes anything, and `-Verify` now prints that caveat. `-Execute -RebuildCatalog`'s recipe is measured to trigger a *Full Index Reset* (events 1008/1004). `-SelfTest` (14 checks, pure) pins the verdict table and that the file writes nothing under the crawl scope keys. |
| `guest/Add-IdentityAccount.ps1` | Builds the **identity account** of `Docs/live-tier-on-the-vm.md` §2.8b - a second POP3 account in the tier profile with its **own** delivery store, `identity@vm.invalid`. Four phases, because they alternate between Outlook closed and running and the script never starts or stops Outlook: `-Phase Import` (renders `identity-account.prf`, `ImportPRF`), then `guest/Add-OutlookPstStore.ps1` attaches and names `identity.pst`, `-Phase CaptureStore` (reads that store's `StoreID` and Inbox `EntryID` over COM), `-Phase Bind` (writes those bytes into the account's `Delivery Store EntryID` / `Delivery Folder EntryID` - **the one undocumented step**), `-Phase Verify` (every account's delivery store and Drafts over COM, no two sharing). **RUN 2026-09-24 ON `OutlookAI-Indexed` END TO END, AND IT WORKS** - two accounts, two stores, each on its own, surviving a restart; checkpoint `CP-09-IDENTITY-ACCOUNT`. **RUN AGAIN on `OutlookAI-Unindexed` the same day, same result** (`CP-10-IDENTITY-ACCOUNT`), adding three measured facts now in its banner: the identity account prompts for its POP3 password until `New-TierProfile.ps1 -StoreSinkPassword -Execute` runs after Bind; closing Outlook needs a graceful quit, not a guest restart; and with Q80 applied `-TrySmtpAddress` reads both addresses with no prompt. A `.prf` alone cannot do it: measured the same day, Outlook binds an appended account to the profile's default store. `-SelfTest` (20 checks, pure). Needs `OutlookMapiInterop.ps1` beside it. |
| `guest/identity-account.prf` | The `.prf` template `Add-IdentityAccount.ps1` renders: `OverwriteProfile=Append` into the tier profile, one POP3 account, **no** service and **no** `DefaultStore`. Sections 6 and 7 verbatim from `tier-profile-forcepst.prf`. **Imported by Outlook on `OutlookAI-Indexed`, 2026-09-24.** |
| `guest/Build-Corpus.ps1` | plan, probe, build, census - with the committed parameters as defaults. **Run many times; it works.** It REFUSES EARLY, before anything touches COM, unless three preconditions hold: the default profile has no **mail** account, Outlook is already running and warm, and **`ImportPRF` is not set** (added 2026-09-24, §4b-i: fail-closed, and `-SkipPreflight` does not skip it). The first two were unchecked until the only build that ever succeeded was found to have satisfied both BY ACCIDENT - which cost five failed builds. It refuses and prints the remedy; it deliberately does NOT fix either condition, because silently rewriting the default profile changes the machine out from under the next step. **First guest run of the preflight, 2026-09-24 from CP-05:** with the tier profile default and Outlook closed it refused with both reasons and wrote nothing (manifest hash, corpus `.pst`, registry unchanged - only its own log); **and it refused the real corpus profile too** - it counted every entry of the account-manager key, and Outlook lists the profile's data file and address book there. Fixed to count mail accounts only, with the measured shapes pinned in `-SelfTest` (81 assertions). **After the fix, from a fresh CP-05** (`CorpusProfile` default, Outlook up 208 s) the preflight passed and the whole pipeline ran against that guest's existing corpus and exited 0: both probes verified, `created 0, already present 20,000, failed 0`, and the census found every ordinal exactly once, in the folder the plan names. With `ImportPRF` set it refused with three reasons, and with `-SkipPreflight` it still refused, on that one. **Guarded since 2026-09-24** (`Assert-TestbedGuest`): it used to rely on the preflight and the tool's own store guard to keep it off the workstation, and `-SkipPreflight` removes the first; now it refuses anywhere but a guest before the log, the expectation sheet, the preflight or the dry run - proven on the workstation - with only `-SelfTest` ahead of it. |
| `host/Get-MailSinkMedia.ps1` | Stages the mail sink on the host: downloads the pinned Inbucket release zip and the checksums file its maintainers attach to the release, and calls it staged only when three values agree - the file's SHA-256, the maintainers' published line for it, and the hash this repository recorded in `MEDIA.md`, which it also checks is still the one this script holds. A download lands under a temporary name and is renamed only once it verifies; it refuses a destination inside the repository that is not under `.work`. Prints the `Copy-ToGuest.ps1` lines. **RUN 2026-09-24 on the host under Windows PowerShell 5.1: `STAGED AND VERIFIED`**, both from a hand-staged copy and from its own `-Execute` download; `-SelfTest` passes under 5.1 and 7. |
| `host/Publish-LiveTierPayload.ps1` | **RUN 2026-09-17, first time, and it worked** - 21 s, 5 projects, 54 packages / 75.7 MB, feed check green on all five with every other source cleared. Stages the two halves an SDK does **not** buy, because the guest has no network and no clone: the suite's source as a `git archive` of a **named commit** (not the working tree, so no stale `obj/` carrying the host's package paths), and every NuGet package in the restore closure as an offline folder feed. It then re-restores the whole suite through a config that clears every source and fallback folder, so a feed short of one transitive package fails **on the host** rather than after 300 MB of copy-in. Host-only, so it takes no VM name. |
| `guest/Install-DotnetSdk.ps1` | Installs the .NET SDK from **staged** media - never a download - pinned by SHA-512, with the hash **mandatory and undefaulted**: a number nobody has compared against anything reads as verified. `-Verify` is four rungs rather than a read-back, because an installer exiting zero is not a machine that can run the suite: the SDK reports a 10.x version and an x64 RID; a throwaway project with no package reference restores, builds and **runs** (which separates "SDK broken" from "feed missing"); `dotnet test --list-tests` discovers the real suite; and one pure T1 class actually executes. Four verdicts, only `TEST-READY` exits zero. Rung 4's filter is `Category!=Live` ANDed **in code**, and the script refuses a filter that is empty or mentions `Live`. **RUN 2026-09-17 on `OutlookAI-Indexed`: VERDICT TEST-READY** - 2,698 tests discovered, 17 executed, 0 failed, entirely offline. It reported `BROKEN` on the first attempt on a machine where everything had worked: `Start-Process -PassThru` leaves `.ExitCode` reading `$null` and every rung compared that against 0. The tell was an EMPTY exit code in the message, not a non-zero one. Fixed; see its banner. |
| `host/Publish-AddInPayload.ps1` | Builds the Outlook add-in on the HOST from a **named commit** (a `git archive`, never the working tree), compiles the product's own `Installer.iss` around it, and stages `AddIn.zip` plus a manifest pinning every hash the guest checks - including the two contract files the add-in and the server agree through, so the guest can tell whether the add-in it installs writes the names the suite reads. **A plain build of a VSTO add-in REGISTERS it on the machine that builds it** (`RegisterOfficeAddin` rewrites `HKCU\...\Outlook\Addins\OutlookAI` to point at the build output; `SetInclusionListEntry` writes a trust entry), so on the maintainer's workstation it would silently repoint his own Outlook. Three guards: the registration target is taken off the chain by a global property; the three registry-writing VSTO tasks are replaced by logging stand-ins (`Override="true"`); and the host's watched keys and certificate stores are snapshotted before and after, with any trace of the build a failure. Signs with a throwaway self-signed key deleted before it exits. Version 99.99.99.0, so the add-in's updater never replaces it. **RUN 2026-09-24 on the maintainer's workstation, three times.** Run 1 refused twice over, both correctly: a `%3B`-escaped target list named one bogus target (MSB4057), and guard 3 caught the throwaway certificate left in `Cert:\CurrentUser\CA` - `New-SelfSignedCertificate` puts a copy there - which was removed and is now cleaned up by the script. Runs 2 and 3: built in 4 s with 0 warnings, both stand-ins logged, registration never scheduled, **376 host lines identical before and after**, installer 41.1 MB. `-SelfTest` 96 assertions, 0 failures, under 5.1 and 7, and each of four rules broken on purpose in a scratch copy was caught. **No guest has installed its payload yet.** |
| `guest/Install-OutlookAIAddIn.ps1` | Puts that payload on a guest and **proves the exact registry state the live tests read**, not that an installer exited 0. Run through `guest/Register-InteractiveTask.ps1`; it refuses session 0, a running Outlook, an unelevated or 32-bit shell, and any machine that is not a guest. Installs the VSTO runtime from staged media (a silent `Installer.iss` skips it), runs the product's installer `/VERYSILENT`, writes the VSTO inclusion-list entry the trust prompt would have written (key path and value names from Microsoft's Japan Office support blog, format from the entry of that shape on the host, comparison rule from the runtime's own IL), sets `VSTO_LOGALERTS=1`, starts Outlook once over COM in a watchdogged child job, and requires `HKCU\Software\OutlookAI\Tuning` to have been written AFTER that start - and the add-in to answer a call into it. Mirrors `HealthReporting.ReadTuningState` down to value types (a REG_QWORD 1 is NOT managed), snapshots the index exclusion before and after, and compares the add-in's contract files with the suite's. Four verdicts; only `ADDIN-READY` exits 0. **RUN 2026-09-24 on `OAI-UNINDEXED`, payload from `fe65ced`: `ADDIN-READY` at the first `-Execute`** - the VSTO runtime installed (v4R absent before, `10.0.60917` after; Office LTSC 2024 brings only v4 `10.0.60910`), the add-in wrote its tuning state 3.5 s after a COM start that took 3.2 s, and connected and answered; no dialog; index exclusion unchanged - **and again after a graceful restart**, keeping the existing trust entry and not reinstalling the runtime. Its guest `-SelfTest` died on first contact when staged at `C:\OutlookAI-Q5` (an empty repository path handed to `Join-Path`); fixed, it now skips the 30 source-reading checks there. `-SelfTest` 115 assertions on the host under 5.1 and 7. |
| `guest/Invoke-GuestMeasure.ps1` | The measurement driver, recovered from the guest. Produced the numbers now in `Docs/magic-numbers.md`. |
| `guest/Measure-SweepCost.ps1` | Per-folder / per-item sweep cost, out of band. **Reconstructed, never executed** - see its banner. **Guarded since 2026-09-24**: read-only, but binding Outlook on the workstation opens the real profile, so it refuses there before touching `-OutFile` or COM (proven on the workstation). |
| `guest/tier-profile.prf` | **RETIRED 2026-09-24 - not a route; `New-TierProfile.ps1` refuses it, and any template carrying a PST service or `DefaultStore`, before writing anything.** One named Unicode PST plus one POP3 account, with `[General] DefaultStore` pointing at the PST. It WAS imported - on `OutlookAI-Indexed`, 2026-09-15 - and left the account's `DeliveryStore` NULL: `DefaultStore` binds the profile's default store, not the account's. Kept as evidence rather than deleted: it is what showed a `.prf`'s `[ServiceN] Name=` carries an exact display name, `@` included (§6 item 10), and `New-OutlookProfile.ps1` builds its PST service blocks from it. |
| `guest/New-TierProfile.ps1` | Creates the tier profile by importing that .prf, then reads the profile hive back and asserts whether Outlook honoured it. Modes: dry run, `-Execute`, `-Verify`, and (2026-09-24) `-SelfTest`. **RUN ON BOTH GUESTS 2026-09-15/16 AND IT WORKS** - first attempt on the second guest, from the committed scripts. Its `-Verify` had a bug worth knowing about: it read the legacy Windows Messaging Subsystem hive and so reported a WORKING route dead five times while its own dump contradicted it. Fixed. **2026-09-24:** the profile it built answered the `ImportPRF` question after the fact - set 19:54:37, gone ever since, Outlook removed it (§4b-i); `-Verify` now removes a lingering one naming the tier `.prf`, but only once the import has run; it carries the `vmadmin` guard it lacked (it writes the Setup key) and so needs `guest/OutlookMapiInterop.ps1` staged beside it; and its account check counts **mail** accounts, not every entry, and no longer fails the working `ForcePSTPath` route for having no `tier.pst`. **Later on 2026-09-24 it became the whole route rather than half of it:** the forcepst template is the default, `-Execute` writes `ForcePSTPath` (REG_EXPAND_SZ, read back) - which nothing under `Testbed/` set before; hand-run scratch did, and wiped part of the Outlook key doing it - a template naming a PST service or `DefaultStore` is refused before any write, and `-Verify` checks that the delivery store, read out of the account's own EntryID, is the one PST in the profile, under `ForcePSTPath`. Proven from `CP-02` in both orders, by the scripts alone (§1: tier first binds at the first start; corpus first left the account unbound until the next one, and `-Verify` said so). **And `-StoreSinkPassword` (2026-09-24):** with the sink up, Outlook holding no POP3 password PROMPTED at every start and never connected, so this seals one - `POP3 Password`, `0x02` plus a DPAPI blob of a NUL-terminated UTF-16LE value, the community-documented layout - into every account of the profile that polls the sink; proven by the sink's debug log (`read USER tier`, `read PASS any-value`, `Processing deletes mailbox=tier`) with no dialog left on screen, and `-Verify` now fails a sink account without one. `-SelfTest`: 55 assertions. See its banner and §5c. |
| `guest/Set-AccountWizardClassic.ps1` | Restores Outlook's classic account wizard and stops AutoDiscover reaching the network. Prerequisite for the UI Automation route. **Guarded since 2026-09-24** (it wrote 18 HKCU values, half under Policies, on any machine). **The registry half ran on a guest the same day** - 18 values written and read back, idempotent, reverted - but **whether Outlook then shows the classic wizard is still unverified**: nobody has opened the Mail applet afterwards. |

---

## 5c. The last three scripts are an EXPERIMENT, not the supported path

Step 7 above still says "by hand", and it still means it. `New-TierProfile.ps1`,
`Set-AccountWizardClassic.ps1` and `Dump-UiaTree.ps1` exist because creating a POP3 account
programmatically turned out to have no proven free route - the object model is read-only for
accounts, Extended MAPI can no longer create POP3 services, and the one component that can is
excluded by the repository's Dependencies rule. They are the two remaining candidates, written so
that a guest can settle them in one checkpoint cycle each.

**ONE OF THEM RAN, AND IT WORKED - so this section is no longer three unrun experiments.**
`New-TierProfile.ps1` built the tier profile on BOTH guests (2026-09-15/16), first attempt on the
second guest from the committed scripts, untouched by hand - and since 2026-09-24 it needs no
hand-run step at all (§1 step 4c). `Set-AccountWizardClassic.ps1` has run only its registry half
(2026-09-24: written, read back, reverted); whether the classic wizard then appears is still
unlooked-at. `Dump-UiaTree.ps1` has never run. Each carries a banner saying which, and each
verifies its own result and exits non-zero rather than reporting a success it did not check.

**Of the two questions that decided this, one is answered:**

* ~~**Does Outlook 16.x process a .prf's internet-account sections at all?**~~ - **YES, MEASURED
  2026-09-15.** Every literal POP3 .prf Microsoft ever published is 2000-2007 era, so this was a
  complete absence of evidence rather than evidence of breakage. The profile Outlook built carries
  a genuine POP3 account: `clsid={ED475411-...}` = `CLSID_OlkPOP3Account`, with the sink's host,
  user and address all landed from section 5. What a `.prf` still **cannot** do is bind the
  account's delivery store - `PROP_ACCT_DELIVERY_STORE` is a binary EntryID and an INI file is
  text - and the way around that is measured too: name no PST service, set `ForcePSTPath`, and let
  Outlook MINT the store, because a store Outlook mints is a store Outlook binds. That is
  `guest/tier-profile-forcepst.prf`.
* **Does the classic account wizard expose non-empty, numeric `AutomationId`s?** **Still open, and
  now much less urgent** - the PRF route works, so the UI Automation route is a fallback nobody
  needs yet. If the ids are non-empty the `Dump-UiaTree.ps1` output *is* the specification for a
  driver; if they are empty or the framework reports `DirectUI`, that route is dead for a
  PowerShell 5.1 client and no driver should be written.

**And there is a third answer, which is to need less.** Most of what the POP3 account was for is
satisfied by *any* account whose SMTP address matches - the product never reads
`Account.AccountType` - and most of the rest is seedable by writing into a PST directly, which the
corpus generator already does. The analysis behind all three, with a source for every claim, is in
`Docs/research/pop3-account-routes.md` - it was gitignored scratch when this paragraph was written
and has since been moved under `Docs/`, which is where it now lives.

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
   **ANSWERED 2026-09-16/17, and the first answer was wrong.** `guest/Set-DefaultOutlookProfile.ps1`
   was written to do this through `IProfAdmin::SetDefaultProfile`. The first time it was ever run it
   threw `E_NOINTERFACE` on `IID_IProfAdmin` (Office LTSC 2024, 16.0.17932.20884) - `MAPIInitialize`
   and `MAPIAdminProfiles` both succeeded, only the QueryInterface failed, and the cause has not been
   established. The script now writes the documented `HKCU\...\Outlook\DefaultProfile` REG_SZ and
   refuses a profile that does not exist; that route is **measured working** on the same guest the
   same day. It matters more than it looks: the corpus tool LOGS ON with the default profile rather
   than attaching to whatever Outlook is running, so the wrong default makes a corpus build refuse.
   Still unrecorded: which profile the ORIGINAL guest had as default.
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
10. ~~**Does Outlook accept `@` in a store display name?**~~ - **ANSWERED: YES, MEASURED TWICE,
    2026-09-15, Office LTSC 2024 16.0.17932.** The hub store must be named after the dummy
    account's SMTP address because several tests use the display name as an address, and this
    gated the whole draft family. Nothing in Microsoft's documentation or in any community source
    states a character restriction either way, so it could only ever be settled empirically - and
    it was, twice, by two different routes, neither of them the probe written for it:
    - through a **`.prf` import** - the profile Outlook built from `guest/tier-profile.prf` carries
      a store whose `Store.DisplayName` reads back **over COM** as literally `tier@vm.invalid`;
    - through a **root-folder rename** - `guest/Rename-OutlookStore.ps1` renamed a store from
      `Outlook Data File` to `tier@vm.invalid` and `Store.DisplayName` followed, without breaking
      the account's delivery-store binding.

    The `@` is accepted, stored, and read back unchanged through both. `Docs/live-tier-on-the-vm.md`
    §8 item 2 carries the same answer with the full evidence.
    **`Add-OutlookPstStore.ps1 -NameProbe` has been REMOVED** rather than left pointing at a
    settled question - and it could not have survived anyway, because it created and deleted a
    throwaway profile, and **deleting a profile has no free route** now that `IProfAdmin` is gone.
    That missing capability is recorded in `guest/OutlookMapiInterop.ps1`'s banner rather than
    quietly dropped; nothing in the build needs it.
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
  `autounattend.xml` becomes tracked;
* the live-test settings template stops holding tokens only, or its fields stop matching
  `live-test-settings.example.json`'s, or a guest section of `testbed.json` names different ones,
  or the renderer's guest destination stops agreeing with where `host/Publish-LiveTierPayload.ps1`
  builds the suite, or any file called `live-test-settings.json` becomes tracked, or an
  address-shaped value in `live-test-settings.example.json` stops ending in `.invalid` (Q79,
  2026-09-24: this repository is public, and an address that can resolve is one somebody could
  own).

The last two are not paranoia. The value they look for has been committed to this repository
before, and an unattend password is `<Password><Value>...</Value></Password>` - which does not
read as an assignment, so the credential-shaped check would walk straight past it.

**And one about the scripts themselves (check 9, 2026-09-24).** It fails when a script under
`guest/` that writes - the registry, the mailbox through any Outlook COM session, a scheduled task,
a service, a locale or power setting, or a process launch - does not call the `vmadmin` guard, or
calls it only after its first write. It was written the day three such scripts were found, one of
them renaming stores behind a banner that claimed an identity check it did not have. The four
writers it found unguarded outside that change are declared in the check by name, each with its
reason, and that list can only shrink (§4b).
