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
| .NET SDK | The .NET 10 SDK, **win-x64**, as the `.exe` installer | **STAGED 2026-09-17**: `.work/media/dotnet-sdk-10.0.401-win-x64.exe` (215,437,248 bytes, gitignored), SHA-512 **matched against Microsoft's published hash**. It is what lets a guest run `dotnet test` at all — see "The .NET SDK" below for the version, the source and the hash. |

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
| Other elements | **`<Updates Enabled="FALSE" />`**, `<RemoveMSI />`, **`<Display Level="None" AcceptEULA="TRUE" />`**, and the `AppSettings` block (company name, default save formats) |
| Cosmetic, and the only other difference | a fresh `Configuration ID` GUID, and an `Info Description` naming this as the testbed configuration. Neither affects the install; they exist so the two files cannot be mistaken for each other. |

**TWO elements deliberately do NOT match the maintainer's file, corrected 2026-09-15 while
actually building a guest. Both were carried over verbatim and both are wrong for a testbed:**

* **`Display Level` was `Full`.** That puts the Office installer's UI on the guest, which defeats
  an unattended build outright — somebody has to be watching a console. `None` with
  `AcceptEULA="TRUE"` is the only setting compatible with the rest of this machine being built by
  script.
* **`Updates` was `Enabled="TRUE"`.** The runbook is explicit in section 2.2: *"Pin the update
  channel. An Office auto-update invalidates the Office checkpoint silently."* A testbed whose
  Office moves underneath its own checkpoints cannot be rebuilt to a known state, and the failure
  is silent — `CP-04-OFFICE-GOLD` stops meaning what it says with nothing to announce it.

**These are the exception that shows where the matching principle stops.** It governs what the
guest *is* — locale, languages, bitness, app settings — not how it is *built* or how it is kept
still. Deviating on a rendering or licensing property would make a guest that is tidy rather than
representative; deviating on the installer's display level makes no difference to any behaviour
under test, and pinning updates is what makes the guest reproducible at all.

**Why the product set is otherwise the ONLY difference.** This testbed's whole design principle is
that the guest matches the maintainer's machine — see "The host configuration the guests match"
above — because that is where the userbase sits. So every property, language and app setting is
carried over verbatim; deviating on any of them would build a guest that is tidy rather than
representative, and would make any difference between guest and host a suspect rather than a
finding. The extra *products* are exempt from that argument because nothing under test touches
Visio or Project.

**`ExcludeApp OutlookForWindows` must stay in whatever you reconstruct** — see above for why. A
guest that ends up with the new client is a guest the live tier cannot run on, and the failure
reads as Outlook automation being broken rather than as a wrong install.

## The .NET SDK — the precondition the live tier has been blocked on

**A guest cannot run the live tier without one, and until 2026-09-17 nothing in this repository
said so in a place a rebuilder would hit.** `Testbed/README.md` question 13 records the shape of
it: everything ever measured on a guest was driven by `Testbed/guest/Invoke-GuestMeasure.ps1`
talking raw stdio to the server, never by `dotnet test` — "which the guest cannot run, having no
SDK". The live tier is an xUnit suite and `dotnet test` is the only supported way to start it,
because the safety machinery — the per-store count tripwire, the `StoreWriteAllowlist`, the
signature-directory snapshot, the zero-artifact sweep — lives in **xUnit fixtures**. Any other
launcher runs the tier with its guards absent, against a real mailbox. That is worse than not
running it, and it is why the answer is an SDK on the guest rather than a cleverer runner.

**This is media, not a dependency.** The repository's `## Dependencies` rule in `CLAUDE.md`
forbids "anything a rebuilder would have to download and install beyond the media
`Testbed/MEDIA.md` already names as preconditions" — so naming it here is exactly the mechanism
that rule points at, the same one the Windows ISO, the Office Deployment Tool and the mail sink
package already use. Microsoft's own SDK, free, no licence key, no third-party component.

### Which version, and why

| | Value |
| --- | --- |
| Product | **.NET SDK 10**, Windows, **x64** |
| Pinned version | **10.0.401** — what the host runs, read with `dotnet --list-sdks` on 2026-09-17 |
| File | `dotnet-sdk-10.0.401-win-x64.exe` — **215,437,248 bytes exactly** (the earlier ~250 MB was an estimate; it is ~205 MiB) |
| Where it comes from | Microsoft's .NET 10 download page, the win-x64 **Installer** under SDK. Not a zip, not `dotnet-install.ps1` (that one downloads, and the guest has no network). |
| Staged at | `.work/media/dotnet-sdk-10.0.401-win-x64.exe` on the host — beside the Windows ISO, for the reason the Office section gives about volatile directories |
| Verified by | **SHA-512**, which is what Microsoft publishes for .NET installers. `Testbed/host/Publish-LiveTierPayload.ps1` prints the hash of the file you staged; compare it against Microsoft's before recording it here. |
| Recorded hash | `f0d8f8e7ec24efb05172a65dd80c4a9b1ef17efcebdbf0f57c15f436eea417960a7eeb6726c473a042373d6a8b94ac1adc7d680decbf7d2c45fa5c5662d62265` |
| Hash provenance | **Matched against Microsoft's own published value**, not merely computed from the file received. Taken from `https://builds.dotnet.microsoft.com/dotnet/release-metadata/10.0/releases.json`, release **10.0.12**, `sdk.files[]`. **The file is named `dotnet-sdk-win-x64.exe` in that metadata - WITHOUT the version** - so a lookup keyed on the download's filename finds nothing and silently reports "not found", which is exactly what happened on the first attempt. |
| Installed with | `Testbed/guest/Install-DotnetSdk.ps1`, which runs it `/install /quiet /norestart` |

**Why 10, and whether a newer one would do.** Every project under `McpServer/` targets
`net10.0-windows`. `OutlookAI.Core` additionally targets `net48`, and that needs **no** separate
install: `McpServer/OutlookAI.Core/OutlookAI.Core.csproj` carries
`Microsoft.NETFramework.ReferenceAssemblies`, which is a NuGet package and travels in the offline
feed below — the csproj says as much, "lets `dotnet build` compile the net48 target without a
Visual Studio install". **There is no `global.json` anywhere in this repository**, so nothing
pins a feature band, and `.github/workflows/mcpserver.yml` asks `actions/setup-dotnet` for
`10.0.x` — meaning CI itself floats. So any .NET 10 SDK would compile the suite, and a .NET 11
SDK almost certainly would too.

**Pin it to the host's version anyway.** The host is the machine that publishes the payload the
guest measures with, and one toolchain across both is one fewer difference to suspect when a
guest behaves unlike the host. That is the same argument this file makes about locale, and the
same one the Office version-gap section makes in reverse when it accepts a 3,598-build gap *and
writes it down as a known limit*. Bump this row when the host is bumped; do not let the guest
drift ahead of it by accident.

**x64 is not optional.** Both the server and the test project set `PlatformTarget x64`, and the
index tier reads the `Search.CollatorDSO` OLE DB provider, which has no 32-bit story here. The
`-Verify` in the guest script reads the RID out of `dotnet --info` and refuses anything else.

### The hash WAS deliberately blank, and is now filled - the reasoning is kept

> **FILLED 2026-09-17.** The installer is staged and its SHA-512 is in the table above, matched
> against Microsoft's published value rather than merely computed from what arrived. The argument
> below is kept because it is why the parameter stays **mandatory and undefaulted** in
> `Install-DotnetSdk.ps1`: a hash that travels with the thing it checks proves nothing, and a
> rebuilder staging a different build must be made to look the number up rather than inherit it.


`Testbed/guest/Install-MailSink.ps1` carries a default SHA-256 taken from a published manifest
and says honestly that it has never been compared against a downloaded file. **This entry carries
no hash at all**, because the agent that wrote it could neither stage nor download the installer,
and a number nobody has compared against anything is worse than an empty field — it reads as
verified.

So the guest script has **no default and refuses without one**:

    .\Install-DotnetSdk.ps1 -ExpectedSha512 <hash> -Execute

**Fill this in on the first real run.** Stage the installer, run
`Testbed/host/Publish-LiveTierPayload.ps1` (it prints the SHA-512 of the staged file), compare
that against the checksum Microsoft publishes beside the download, and write the value into the
table above. From then on it is a one-line check rather than a piece of work.

### The other two halves, which an SDK alone does not buy

**A guest with an SDK still cannot run `dotnet test`.** It also needs the source — the suite is
built from source, there is no prebuilt test assembly — and it needs **54 NuGet packages, ~76 MB**
(counted from the host's restore graph on 2026-09-17), on a machine with no route to nuget.org.
`Testbed/host/Publish-LiveTierPayload.ps1` stages both:

| Artefact | Guest path | What it is |
| --- | --- | --- |
| `Source.zip` | `C:\OutlookAI-Q5\src` | `git archive` of a named commit. Built from a commit, not the working tree, so a payload is reproducible and carries no stale `obj/`. |
| `NuGet.zip` | `C:\OutlookAI-Q5\nuget-offline` | A flat folder feed of every `.nupkg` the restore graph names. The script re-restores the whole suite through a config that clears every other source, so a feed that is short of a package fails on the **host**, where the fix is a minute. |

**Those two are ARTEFACTS, not media** — a script regenerates them from this repository, which is
the distinction this file draws everywhere else. They are listed here because they are
preconditions for the same act, and because a reader who stages only the SDK will get the guest
script's `SDK-ONLY` verdict and should know in advance what it means.

**One thing the payload cannot carry:**
`McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json`. It is gitignored
because it names real stores and this repository is public, so each guest needs its own.
`Testbed/live-test-settings.example.json` is the committed shape. Without it the tier has no
write allowlist, which is a refusal rather than a pass — by design. **For a guest it is rendered,
not written**: `Testbed/host/New-LiveTestSettings.ps1` builds it from the guest's section of
`Testbed/testbed.json` once that guest's store names have been read off it, and prints the
`Copy-ToGuest.ps1` line that lands it.

### Installing it, and where in the build order it goes

    pwsh -File Testbed/host/Publish-LiveTierPayload.ps1
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path .work\media\dotnet-sdk-10.0.401-win-x64.exe -Destination C:\OutlookAI-Q5\media\dotnet-sdk-10.0.401-win-x64.exe
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path .work\testbed-livetier-payload\Source.zip -Destination C:\OutlookAI-Q5\Source.zip
    pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName <guest> -Path .work\testbed-livetier-payload\NuGet.zip  -Destination C:\OutlookAI-Q5\NuGet.zip

then on the guest, elevated — **PowerShell Direct is fine for this, and only for this**: an SDK
install touches no COM and no Outlook, so it does not need
`Testbed/guest/Register-InteractiveTask.ps1`. The `dotnet test` that follows *does*, because
Outlook cannot finish starting in session 0.

    Expand-Archive C:\OutlookAI-Q5\Source.zip -DestinationPath C:\OutlookAI-Q5\src            -Force
    Expand-Archive C:\OutlookAI-Q5\NuGet.zip  -DestinationPath C:\OutlookAI-Q5\nuget-offline  -Force
    .\Install-DotnetSdk.ps1 -ExpectedSha512 <hash> -Execute

**Budget ~2.5 GB on the guest's C:** — roughly 900 MB installed SDK, ~300 MB extracted package
cache, and a Release build of the suite on top.

**Take the checkpoint after this, not before.** An SDK, a machine `PATH` edit and a package cache
are a real change to the machine, and the whole point of the checkpoint discipline in
`Testbed/README.md` is that a named checkpoint describes a state somebody can return to.

**Neither script has ever been run.** Both carry the banner saying so. Replace those banners with
what actually happened the first time either of them runs on a guest.

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

**4. The policy is settled: perpetual releases only, never a subscription (decided 2026-09-24).**
The guests stay on LTSC 2024. The maintainer's workstation stays on LTSC 2021 until he chooses to
move it, and meanwhile counts as a production user on an older version. There is therefore **no
current-channel coverage anywhere**, and that is accepted — see `Docs/live-tier-on-the-vm.md`
section 9 for the full reasoning.

**LTSC 2024 is not the last perpetual release.** In its April 2024 *preview* announcement of
LTSC 2024, Microsoft said it is committed to another on-premises release after it; an independent
analyst reads the usual three-year cadence as pointing to around 2027. When that ships, moving the
guests is a change of staged media here, not a rebuild by hand. **Evidence class, stated
honestly:** the commitment is recorded here as REPORTED - Microsoft's page renders client-side and
could not be read verbatim when this was written (2026-09-24), so the sentence was not quoted from
it directly. Sources:
[Upcoming preview of Microsoft Office LTSC 2024 (Microsoft Tech Community)](https://techcommunity.microsoft.com/blog/microsoft_365blog/upcoming-preview-of-microsoft-office-ltsc-2024/4082963),
[Microsoft's subscription-free 'perpetual' Office LTSC 2024 (Directions on Microsoft)](https://www.directionsonmicrosoft.com/microsofts-subscription-free-perpetual-office-ltsc-2024-to-ship-this-year/),
[Office 2024 and Office LTSC 2024 FAQ (Microsoft Support)](https://support.microsoft.com/en-us/office/lifecycle/office-2024-and-office-ltsc-2024-faq).

## The rule

**Never destroy a working testbed before the replacement runs.**

Build the new machine alongside the old one, prove it, and only then remove the old. The overlap
costs disk; the alternative cost, on a machine with no staged media, is the testbed itself. This
is not hypothetical caution — it is what the survey on 2026-08-24 prevented.
