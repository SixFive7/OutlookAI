# Installation media — the precondition nobody wrote down

**Staging media is step zero.** The rest of `Testbed/README.md` assumes you can install Windows
and Office; neither is in this repository and neither can be, so this file records what is
needed, what this machine already has, and what has to be obtained.

This exists because the gap was discovered the expensive way: a rebuild was authorised, and the
survey done immediately before it found **no Windows installation media anywhere on the host**.
Had the old VM been destroyed first, as the instruction literally said, the testbed would have
been unrebuildable. Hence the rule at the bottom of this file.

## What is needed

| | Needed | On this machine (checked 2026-08-24) |
| --- | --- | --- |
| Windows | A Windows 11 x64 image | **STAGED 2026-08-24**: `.work/media/Win11_25H2_EnglishInternational_x64_v2.iso` (7.9 GB, gitignored). Consumer multi-edition, volume label `CCCOMA_X64FRE_EN-GB_DV9`, so it carries Pro. |
| Office | Office Deployment Tool + a configuration | **STAGED**: `.work/office-odt/` (gitignored), holding `setup.exe` and `VoIPFabric.xml`. A testbed-specific `Testbed.xml` sits beside them — see below. |

### Windows — staged, and it is NOT the edition the old guest ran

`.work/media/Win11_25H2_EnglishInternational_x64_v2.iso`, verified: ISO 9660 signature present,
volume label `CCCOMA_X64FRE_EN-GB_DV9` — Microsoft's consumer multi-edition x64 image, which
includes **Pro**. It sits in gitignored scratch; it is 7.9 GB and must never be committed.

**Two deliberate differences from the machine it replaces, both of which change something.**

**1. Consumer Pro, not Enterprise LTSC Evaluation.** The old guest was `TIMEBASED_EVAL` with a
hard 90-day expiry. An unactivated consumer Pro install has **no expiry at all** — it watermarks,
blocks personalisation and nags, but it does not stop. Since the decided rebuild cadence is
driven by Office's 30-day grace, the Windows clock was doing no useful work, and removing it
means one fewer way for the testbed to die silently. The evaluation route also needed a
registration form; this image did not.

**2. `EN-GB`, not `EN-US` — VERIFIED 2026-08-25, AND IT IS THE RIGHT IMAGE.** This entry used to
say "verify this before trusting any measurement", because English International was assumed to
be a mismatch that had crept in. It is not. The host was surveyed on 2026-08-25 and its effective
display language **is en-GB** — see the next section for the whole table and the command behind
each row. The mechanism is worth understanding rather than memorising: the host's first preferred
language is **en-NL**, English (Netherlands), and **Windows ships no MUI for en-NL**, so the
display language falls back to **en-GB**. That is why `Get-UICulture` reports en-GB on a machine
whose language list never mentions it, and it makes `CCCOMA_X64FRE_EN-GB_DV9` the *correct* base
image for a guest that is supposed to look like this host. Nothing to correct.


## The host configuration the guests match

**The guests are built to match the maintainer's own machine, deliberately.** Not a clean
en-US default, not a "sensible" configuration: that machine's configuration is where most of the
userbase sits, so it is what the live tier should be testing against. A testbed set up the tidy
way would be a testbed that cannot reproduce the bugs the userbase hits.

**Measured on the host (PC657) on 2026-08-25.** These are readings, not intentions — the command
behind each is in the last column so any of them can be checked rather than believed.

| Setting | Value | Read with |
| --- | --- | --- |
| OS | Windows 11 Pro, 10.0.26200 (25H2) | `Win32_OperatingSystem` |
| Base install language | en-US (`OSLanguage` 1033, `Locale` 0409) | `Win32_OperatingSystem` |
| MUI languages present | en-US, en-GB, nl-NL | `Win32_OperatingSystem.MUILanguages` |
| **Effective display language** | **en-GB** | `Get-UICulture` |
| **Preferred language list** | **en-NL** (English, Netherlands), then **nl-NL** | `Get-WinUserLanguageList` |
| **System locale** (non-Unicode / ANSI) | **en-US** | `Get-WinSystemLocale` |
| **User locale / formats** | **nl-NL** | `Get-Culture` |
| Date format | `d-M-yyyy` — 2026-08-25 renders `25-8-2026` | `Get-Culture` |
| Number format | decimal `,`, group `.` — `4000.5` renders `4.000,50` | `Get-Culture` |
| Currency / first day of week | `€` / Monday | `Get-Culture` |
| Home location | Netherlands, **GeoId 176** | `Get-WinHomeLocation` |
| **Keyboard, both languages** | KLID **`00020409`**, United States-International | input tips `2000:00020409` and `0413:00020409` |
| Time zone | **`W. Europe Standard Time`** (UTC+01:00 Amsterdam), DST on | `Get-TimeZone` |

**The en-NL line is the load-bearing one.** en-NL has no MUI, so the display language falls back
to en-GB — which is why the row above it says en-GB, and why English International is the right
image. It is also why the language list cannot live in the answer file: **en-NL is a transient
language**, handed an LCID out of the `0x2000` block at runtime (currently `2000`, but that is an
allocation, not an identity). `Testbed/guest/Complete-FirstLogon.ps1` sets it after the account
exists, and `Testbed/guest/autounattend.template.xml` says so where a reader will hit it.

**THE GUESTS WILL RENDER `4.000,50`, AND THAT IS THE POINT.** This project has been bitten by
locale once already: the remediation console printed `4.000` for four thousand on a Dutch-locale
machine, and was pinned to the invariant culture precisely because its output is compared across
machines. A nl-NL user locale reproduces exactly the conditions that found that, so the caution is
not "check whether the guest is safe" but **"the corpus, the assertions and every rendered payload
have to survive it"**. Anything that only passes on an en-US box is a defect on the maintainer's
machine too, and the testbed exists to say so before a user does.

Outlook's default folder names (`Inbox`, `Sent Items`, `Deleted Items`, `Junk Email`) are
identical between en-GB and en-US, so folder resolution is unaffected either way. What is still
not established, and is a question for whoever next runs the tier on a fresh guest, is whether any
assertion, corpus date parse or rendered payload is culture-sensitive in a way nobody has hit yet.
The guest is now the place that would show it.

### Office — the method, which was the actual unknown

The Office Deployment Tool (`setup.exe`) plus an XML configuration.

**WHERE IT ACTUALLY IS: `.work/office-odt/`.** This file used to say the Office media sat "in an
archive under the maintainer's Downloads". That was **wrong, and wrong in the dangerous
direction** — it named a volatile location as the home of a precondition. On 2026-08-23 everything
in that Downloads directory older than roughly five days was deleted with no warning and no
prompt, taking thirteen scratch directories with it. A runbook that points a rebuilder at
Downloads is a runbook that eventually points at nothing.

`.work/` is the repository's own gitignored scratch directory, and it is the **right** home for
this precisely because Downloads is volatile: it sits beside the checkout, it is visible to anyone
who clones and looks, nothing outside this project prunes it, and `.gitignore`'s `.work/` rule
keeps its contents — a product key among them — out of a public repository. That rule is what
makes staging here safe; confirm it with `git check-ignore -v .work/office-odt/` before adding
anything.

| File | What it is |
| --- | --- |
| `.work/office-odt/setup.exe` | The Office Deployment Tool. 7.2 MB, dated 2024-08-09. |
| `.work/office-odt/VoIPFabric.xml` | The **maintainer's workstation** configuration. Carries product keys. |
| `.work/office-odt/Testbed.xml` | The **testbed** configuration, narrower — see below. Carries a product key. |

**THIS IS A PRECONDITION, NOT AN ARTEFACT.** Nothing in this repository regenerates it, and
`.work/` is scratch — a tidy-up, a fresh clone or a disk swap leaves the directory empty and
nothing announces it. **If `.work/office-odt/` is missing, re-stage it before step 4 of the
runbook:** download the Office Deployment Tool from Microsoft, extract its `setup.exe` there, and
recreate the configuration from the table below plus the product key, which lives with the
maintainer and nowhere else. The Windows ISO in `.work/media/` has exactly the same standing.

**The configuration files contain product keys and are therefore not reproduced here.** What the
**existing** guest was built with — `VoIPFabric.xml`, the workstation configuration, because at
the time there was no other:

| Setting | Value |
| --- | --- |
| Product | `ProPlus2024Volume` (alongside Visio, Project and proofing tools, which the guest did not need) |
| Channel | `PerpetualVL2024` |
| Edition | 64-bit (`OfficeClientEdition="64"`) |
| Languages | `en-us`, `nl-nl`, `MatchOS` |
| Excluded apps | `Lync`, `OneDrive`, **`OutlookForWindows`** |
| Activation | `AUTOACTIVATE=1` |

**`ExcludeApp OutlookForWindows` is load-bearing, not cosmetic.** It suppresses the *new* Outlook.
Everything this project does goes through classic Outlook's COM object model, which the new
client does not provide.

**New guests use `Testbed.xml` instead.** The section below is what to install them with.

### The testbed configuration — `.work/office-odt/Testbed.xml`

**A testbed configuration is narrower than the maintainer's**, and as of 2026-09-15 it exists as
its own file rather than as an argument in this document. `VoIPFabric.xml` installs Visio, Project
and proofing tools because it is a workstation configuration. A testbed wants `ProPlus2024Volume`
alone — Outlook is required, Word is worth keeping because HTML signatures are rendered through
it, and the rest is install time and disk for nothing.

**Install with it like this, on the guest:**

    .\setup.exe /configure Testbed.xml

**It carries a product key, so it lives in `.work/` and nowhere else.** Never under `Testbed/`,
never anywhere tracked. `.github/scripts/check-testbed-references.ps1` fails the build on a
credential-shaped literal under `Testbed/`, but do not rely on that as the guard — the file simply
does not belong in the repository at all.

**What is in it, so a rebuilder can reconstruct it without the key.** It is `VoIPFabric.xml` with
the three extra products deleted; nothing else is changed except the two cosmetic identifiers in
the last row:

| | Value |
| --- | --- |
| Products | `ProPlus2024Volume` **only** — `VisioPro2024Volume`, `ProjectPro2024Volume` and `ProofingTools` removed |
| `PIDKEY` | the same volume key `VoIPFabric.xml` carries for `ProPlus2024Volume`. **Not written down anywhere in this repository** |
| Channel | `PerpetualVL2024` |
| Edition | `OfficeClientEdition="64"` |
| Languages | `en-us`, `MatchOS`, `nl-nl` |
| Excluded apps | `Lync`, `OneDrive`, **`OutlookForWindows`** |
| Properties | `SharedComputerLicensing=0`, `FORCEAPPSHUTDOWN=TRUE`, `DeviceBasedLicensing=0`, `SCLCacheOverride=0`, `AUTOACTIVATE=1`, `PinIconsToTaskbar=FALSE` |
| Other elements | `<Updates Enabled="TRUE" />`, `<RemoveMSI />`, `<Display Level="Full" AcceptEULA="TRUE" />`, and the `AppSettings` block (company name, default save formats) |
| Cosmetic, and the only other difference | a fresh `Configuration ID` GUID, and an `Info Description` naming this as the testbed configuration. Neither affects the install; they exist so the two files cannot be mistaken for each other. |

**Why the product set is the ONLY difference.** This testbed's whole design principle is that the
guest matches the maintainer's machine — see "The host configuration the guests match" above —
because that is where the userbase sits. So every property, language and app setting is carried
over verbatim; deviating on any of them would build a guest that is tidy rather than
representative, and would make any difference between guest and host a suspect rather than a
finding. The extra *products* are exempt from that argument because nothing under test touches
Visio or Project.

**`ExcludeApp OutlookForWindows` must stay in whatever you reconstruct** — see above for why. A
guest that ends up with the new client is a guest the live tier cannot run on, and the failure
reads as Outlook automation being broken rather than as a wrong install.

## The licence clocks, and the corrections worth reading

Both clocks were measured on the guest on 2026-08-24, and the Office one again on 2026-09-15:

| | Channel | 2026-08-24 | 2026-09-15 |
| --- | --- | --- | --- |
| Windows | `TIMEBASED_EVAL` | ~82 days | not re-read; expiry lands mid-November |
| Office | KMS client, **out-of-box grace** | ~16 days | **expired** — `LicenseStatus=5`, `GracePeriodRemaining=0`, 7 days past |

**The decision was to treat the VM as disposable and rebuild when a clock expires. That does not
work for Office as configured, and the arithmetic says why.** The guest was installed on
2026-08-09; 15.7 days of Office grace remained on 2026-08-24. That is a **30-day** out-of-box
grace begun at install — not a 90-day one. **Office is the binding constraint, and it binds
roughly monthly.**

The old evaluation Windows image reset to 90 days, which is where "Windows resets to 90" came
from; the staged replacement is consumer Pro and **never expires at all** (see "Windows — staged"
above). Either way Windows never becomes the reason to rebuild — the cadence is Office's.

**DECIDED 2026-08-24: accept the monthly rebuild. RETIRED 2026-09-15 — the cadence solved a
problem that does not exist.**

The monthly rebuild was adopted because Office's 30-day grace was believed to disable the guest.
**It does not.** Measured past grace on 2026-09-15: every COM read this project uses still works,
and a cold COM start completes in 3.7 s with no dialog. Nothing stops working when that clock
runs out.

**The clock that actually bit was the corpus, and it already had a guard.** During a three-week
absence the corpus went 27 days stale and its 1-day and 7-day windows emptied — while the Office
clock, which the cadence was attached to, cost nothing. `corpus-verify` already refuses the tier
fail-closed when a declared window has emptied. So the schedule was watching the harmless clock
while the harmful one was handled.

**There are now two triggers, and neither is a calendar:**

1. **`corpus-verify` refuses.** Driven by the windows the machine actually declares in
   `windowDays`, so it fires exactly when a measurement has stopped being possible.
2. **Before a release.** Not because anything expires, but because **a from-nothing rebuild
   playbook rots exactly as quietly as a stale corpus**, and this is the moment that matters. A
   calendar is the wrong instrument for that: the guests can go untouched for months, and a
   schedule nobody needs is a schedule that gets skipped and then distrusted.

Not chosen, and still worth knowing why they were on the table: making a KMS host reachable so
`AUTOACTIVATE=1` succeeds would remove the Office clock entirely but depends on guest networking
nobody has verified; licensing the guest another way spends a licence on a machine meant to be
thrown away. **Both now buy nothing**, because the clock they would remove costs nothing.

### What past grace actually does — and it is NOT "reduced functionality"

**This file used to say that a KMS client past grace "drops Office into reduced functionality".
That is wrong, and it is wrong in the direction that changes a decision** — it made the rebuild
cadence read as a hard stop on the machine whose only purpose is driving Outlook, when the
measured behaviour is a nag.

**The term does not describe this product.** "Reduced functionality" does not appear in
Microsoft's volume-activation documentation for Office LTSC 2024 at all — not on *Overview of
volume activation of Office*, not on *Activate volume licensed versions of Office by using KMS*,
not on *Activate volume editions of Office*, all three of which state that they apply to
LTSC 2024. The phrase is two other things:

* An **Office 2007 / Windows Vista era** term, from a licensing model this product does not use.
* Separately, a live **Microsoft 365 Apps subscription** term — an unlicensed or deactivated
  subscription install, where "users can only view and print their documents. All features for
  editing or creating new documents are disabled." That is a subscription concept and cannot
  arise on a volume KMS client. Its closest documented app list is **viewer mode**, which is
  supported for "Version 1902 or later of Word, Excel, and PowerPoint" and "Version 2005 or later
  of Project and Visio". **Outlook has never been on that list.**

**The documented terminal state for a volume KMS client is "Unlicensed notification"**, and
Microsoft's own KMS licence-state table for LTSC 2024 describes it in one sentence: *"Users then
see notifications that request activation and a red title bar."* No functional loss is described.

*Activate volume editions of Office* — the page that covers 2016 through 2024 and **lists Outlook
by name** among the applications it applies to — goes further: **"there is no functionality loss
even if the licenses for KMS clients cannot be renewed."**

**State the caveat honestly, because it is a real one.** That sentence sits in a paragraph about
the **180-day renewal** path — a client that activated once and then lost its KMS host. This
guest is on the **out-of-box** path: it has never reached a KMS host at all. Both paths terminate
in `LicenseStatus 5`, which is why the sentence very likely covers this case too, but **that is
inference and not a quotation about this machine.**

**MEASURED ON THE GUEST, 2026-09-15** — 7 days past grace, `LicenseStatus=5`,
`GracePeriodRemaining=0`, `LicenseStatusReason 0xC004F056`, SKU
`Office24ProPlus2024VL_KMS_Client_AE`, channel `VOLUME_KMSCLIENT`. Outlook was running
**through** the expiry — 28 days up, `Responding=True` — and **every COM read this project uses
still works**: `CreateObject`, `GetNamespace`, `Stores`, `Accounts`, `GetDefaultFolder`,
`GetTable`, `Restrict`, `Sort`, `PropertyAccessor`. The corpus counts came back intact and
matching `Testbed/testbed.json`.

**Bound that claim to what it covers: reads.** `CreateItem`, `Save`, `Move`, `Delete` and `Send`
were **not** exercised, because mailbox-safety rule 1 forbids mutating items from ad-hoc shell
code. So "no functional loss" is measured for the read path and *inferred* for the write path
from Microsoft's wording. If a write path did degrade past grace, this measurement would not have
seen it.

**THE COLD START WAS THE ACTUAL RISK, AND IT HAS NOW BEEN MEASURED. IT DOES NOT HAPPEN.**

The concern was a modal activation prompt appearing when Outlook is **started** on a past-grace
guest. Microsoft's server-side automation guidance says of a blocking dialog that "the
`CreateObject` function and the `CoCreateInstance` function stop responding and never finish, or
take a long time to return", and the unattended-automation article says an Office dialog "might
result in the application appearing to 'hang'". That is a hang rather than an error, which is the
expensive shape — it reads as a wedged suite rather than as a licence.

**Measured 2026-09-15 on the guest**, `LicenseStatus=5`, `GracePeriodRemaining=0`, 7 days past
grace. The guest was **restarted** so that `OUTLOOK.EXE` was genuinely not running; the probe
asserted that before touching COM and would have refused otherwise, because attaching to a
running instance is what made the earlier attempt uninformative.

```
--- precondition: Outlook must NOT be running ---
   confirmed: OUTLOOK.EXE is not running
--- COLD CreateObject, watchdogged at 150 s ---
   CreateObject returned in 3.7 s; full bind in 4.4 s - v16.0.0.17932, 1 store(s), 0 account(s)
   VERDICT: a cold COM start COMPLETED under an expired licence.
```

**No dialog appeared** — the probe enumerated titled windows in session 1 before and after and
found none either time. So the documented "no functionality loss" holds for the one path that
would actually have cost an evening.

**Two bounds on that result, both real.** It started into the **configured** `OutlookAITest`
profile. A *bare* profile — this guest's other one, which references no data file — would raise
Outlook's account-setup wizard, and that **is** a modal dialog that would hang a COM start. That
is a profile fault, not a licence fault, but the symptom is identical, so a rebuilder who sees a
cold start hang should check the profile before blaming the licence. And the **write path**
(`CreateItem`/`Save`/`Move`/`Delete`/`Send`) is still unmeasured, because mailbox-safety rule 1
puts it out of reach of a probe.

### Reading the licence state — the query, and why it is shaped this way

Measured 2026-09-15. Written down rather than left to be re-derived, because it is non-obvious in
three separate places:

    SELECT Name, Description, LicenseStatus, LicenseStatusReason, GracePeriodRemaining, PartialProductKey
    FROM   SoftwareLicensingProduct
    WHERE  ApplicationID = '0ff1ce15-a989-479d-af46-f275c6370663'
      AND  PartialProductKey IS NOT NULL

* **It works unelevated** — verified on a token where `IsInRole(Administrator)` is `False`. A
  preflight that needed elevation would not be a preflight.
* **Filter in the query, not afterwards.** **241 ms** filtered, against **10,679 ms** enumerating
  the class and filtering in PowerShell — a 44× difference. The naive form is unusable in a
  preflight; the filtered form is cheap enough that there is no argument against running it.
* **`PartialProductKey IS NOT NULL` is load-bearing.** Without it, a perfectly healthy machine
  returns keyless SKU-catalogue rows carrying `LicenseStatus = 0`, so a check phrased as "any row
  that is not Licensed" fires on **every** healthy machine, forever.
* `0ff1ce15-a989-479d-af46-f275c6370663` is Office's `ApplicationID`. It is the same on every
  machine and every Office version.

`LicenseStatus` values, and what a preflight should do with each:

| | Meaning | Preflight |
| --- | --- | --- |
| 0 | Unlicensed | refuse |
| 1 | Licensed | proceed |
| **2** | **Out-of-box grace** | **warn** — `GracePeriodRemaining` is the countdown, in **minutes** |
| 3 | Out-of-tolerance grace | warn |
| 4 | Non-genuine grace | warn |
| **5** | **Notification** | **refuse** — grace is at zero |
| 6 | Extended grace | warn |

**Branch on the status before reading the number.** On a *licensed* KMS client
`GracePeriodRemaining` is the 180-day renewal countdown, not an expiry — read naively it reports
a healthy machine as one about to die.

**A trap worth recording: the query can return zero rows transiently while `sppsvc` is starting.**
"The query failed" and "no Office is installed" must not collapse into one answer, or the check
fails **open** on exactly the fault it exists to catch.

### The preflight check was NOT built, and the query above is kept anyway

**DECIDED 2026-09-15: not building it.** Both justifications it was ever given were measured
false. It was first written on "past grace, Office drops into reduced functionality" — a term
that does not apply to this product at all. It was rewritten on "the startup-hang risk is
unquantified" — and the cold start was then measured at 3.7 s with no dialog. What survived was
only "241 ms is cheap", which is an argument for adding a check to anything, and was not the
argument that was authorised. The standing instruction is **"if it does not impact our work, do
nothing"**, and it does not.

**The argument that once carried it, and why it no longer applies.** It used to read: *a cadence
that depends on remembering gets skipped exactly once, and then the tier stops with no visible
cause* — and a three-week absence duly skipped it, leaving the guest 7 days past grace. That
reasoning was sound while the licence was believed to stop the tier. It doesn't. **The tier did
not stop; nothing stopped.** The clock worth watching was the corpus, and `corpus-verify` was
already watching it fail-closed.

**Why the query above is kept in this file even so.** It is a measured, non-obvious recipe — the
unelevated access, the 44x cost difference, the load-bearing `PartialProductKey` clause, the
branch-before-the-number rule and the `sppsvc` fail-open trap are all things somebody would
otherwise have to rediscover. Keeping the knowledge is free; building the check was not
justified. **If a future measurement shows the write path or an escalated licence state does
break something, this is the recipe to build the gate from** — the reason to close was the
absence of impact, not the absence of a mechanism.

**The consequence reaches further than the check.** The monthly rebuild cadence was adopted
because Office's 30-day grace was believed to disable the guest. It does not. Nothing measured so
far stops working when that clock runs out, so **the thing that actually forces a rebuild is
corpus staleness, not the licence** — and that already has a fail-closed guard in
`corpus-verify`, which refuses the tier when a measurement window has emptied. Left open in
`TODO.md` pending the maintainer's decision, because closing a mailbox-adjacent safety item and
retiring a rebuild cadence are both theirs to make, not an agent's.

## The Office version gap between host and guest — an ACCEPTED KNOWN LIMIT

**Measured 2026-09-15.**

| | Host (the maintainer's machine) | Guest (the testbed) |
| --- | --- | --- |
| `ProductReleaseIds` | `ProPlusSPLA2021Volume` (plus Project and Visio) | `ProPlus2024Volume` |
| Build | 16.0.**14334**.20848 | 16.0.**17932**.20884 |
| Audience / channel | `Production::LTSC2021` | `PerpetualVL2024` |
| Licensing | **MAK, `LicenseStatus=1`, no grace clock at all** | KMS client, `LicenseStatus=5` |

**DECIDED 2026-09-15: stay on Office 2024.** Recorded here as a **known limit**, not as a
non-issue, because three things have to survive the decision.

**1. It is a deliberate exception to this testbed's own governing principle.** "The host
configuration the guests match" above says the guests copy the maintainer's machine on purpose,
because that is where the userbase sits — and that argument is what justified carrying the locale
over verbatim rather than tidying it to en-US. The Office *version* does not follow it. There are
**3,598 builds** between the machine the live tier is validated on and the machine the maintainer
runs. Saying so plainly is better than leaving a reader to find the inconsistency and assume it
was an accident.

**2. Practical consequence: an Outlook build difference is the FIRST thing to suspect** when a
live test behaves differently on the VM than on the maintainer's machine. `Docs/live-tier-on-the-vm.md`
section 2.1 already says as much; this table is the measured reason it says it, and section 9
records the limit.

**3. The grace clock is an artefact of the guest being a KMS client — the userbase never sees
it.** The maintainer's own Office is MAK-activated: `LicenseStatus=1`, no clock of any kind, no
renewal countdown, nothing to expire. So the 30-day rebuild cadence above is a property of **the
testbed's licensing choice**, not of Office 2024 and not of anything a user experiences. That
matters because the cadence has previously been discussed as though it described reality.

## The rule

**Never destroy a working testbed before the replacement runs.**

Build the new machine alongside the old one, prove it, and only then remove the old. The overlap
costs disk; the alternative cost, on a machine with no staged media, is the testbed itself. This
is not hypothetical caution — it is what the survey on 2026-08-24 prevented.
