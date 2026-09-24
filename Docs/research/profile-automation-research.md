# Creating Outlook profiles, stores and accounts without a GUI

**What this is.** The testbed's guests are built unattended through Windows installation, and then
stop: `Testbed/README.md` step 4 and step 7 both say "by hand", and
`Docs/live-tier-on-the-vm.md` section 2.5 says outright that how the profiles are created "is
**not recorded**; the Mail control panel works and is the obvious route." The maintainer has done
it once by driving that GUI through a vision model, which was slow and expensive. This document is
the research for doing it programmatically, and the draft scripts under `Testbed/guest/` are the
other half.

**Nothing in here was executed on the machine that wrote it.** That machine is the maintainer's own
workstation, with a real Outlook profile and delegate mailboxes on it. Every PowerShell file this
work produced was verified by **parsing** (`[Parser]::ParseFile`), never by running. The only
empirical facts below are ones the maintainer measured on a guest, and they are labelled as such.

## How to read the labels

| Label | Means |
| --- | --- |
| **[MS-DOC]** | Documented by Microsoft. |
| **[COMMUNITY]** | Reported by a blog, forum, Stack Overflow answer, or a third-party tool's documentation. |
| **[GUEST-MEASURED]** | Measured by the maintainer on a testbed guest, date and build given. |
| **[REPO]** | Read out of this repository by me, in this session. Checkable by anyone with the checkout. |
| **[INFERRED]** | My inference. The thing it is inferred from is always named. |

An unlabelled sentence is structure, not a claim.

---

## 1. The answer, per thing that has to be creatable

| # | What | Route | Confidence |
| --- | --- | --- | --- |
| 1 | Profile with **no mail accounts** (the corpus profile) | ~~Extended MAPI `IProfAdmin::CreateProfile`, `ulFlags = 0`~~ **measured broken 2026-09-16** - same `E_NOINTERFACE` gateway as row 5. The route that shipped is a **`.prf` import** (`ImportPRF`), which builds the profile and its named stores in one pass | **[MEASURED]** for the mechanism - it is what `New-TierProfile.ps1` does on both guests. **[INFERRED]** for the account-less part: an empty `[Internet Account List]` has not been run. Fallback, **[MEASURED]**: `outlook.exe /PIM <name>` gives `accounts=0`. See §3.2. |
| 2 | Profile with a **POP3/SMTP account** (the tier profile) | ~~No free programmatic route exists.~~ **A `.prf` naming no PST service and no `DefaultStore`, plus `ForcePSTPath`: Outlook mints the store and binds the account to it.** `New-TierProfile.ps1` with `tier-profile-forcepst.prf` | **[MEASURED] on both guests, 2026-09-15/16.** §4. |
| 3 | **Third mail account** + own delivery store + signature | Account: **the same `.prf` route creates it** (`OverwriteProfile=Append`), but binds it to the profile's default store. **Its own store:** `AddStoreEx` + rename, then the account's two `Delivery ... EntryID` registry values set to bytes read over COM. `Add-IdentityAccount.ps1`. **Signature: solved** — this repository already ships the tool. | **[MEASURED] 2026-09-24**, with one **undocumented** step (the registry binding). Signature tool never executed. §4. |
| 4 | **PST store with an exact display name** (incl. `@`) | ~~Extended MAPI `IMsgServiceAdmin::CreateMsgService("MSUPST MS")` + `ConfigureMsgService`~~ **behind the same broken gateway**. Two routes replace it: the `.prf`'s `[ServiceN] Name=` at creation, and `NameSpace.AddStoreEx` + a **root-folder rename** into a profile that already exists | **[MEASURED] both, and the `@` is ANSWERED: yes.** `Store.DisplayName` read back over COM as `tier@vm.invalid` by each route independently. §5 has the evidence; it is no longer an open question. |
| 5 | **Switch the default profile**, no prompt | ~~`IProfAdmin::SetDefaultProfile`~~ **measured broken 2026-09-16** - `E_NOINTERFACE` on `IID_IProfAdmin`, Office LTSC 2024. The working route is the documented `HKCU\...\Outlook\DefaultProfile` REG_SZ, plus `PickLogonProfile = 0` | **Good, but not by the mechanism this table originally named.** See §7. |

**The headline.** Four of the five are automatable and one is not. The one that is not is the mail
account, and it is not a gap in this research — it is a capability Microsoft removed from every
public interface. §4 is the evidence, from three independent directions, and it is the section to
read if you read only one.

> **[MEASURED] 2026-09-24 — THE HEADLINE ABOVE IS WRONG: ALL FIVE ARE AUTOMATABLE.** The mail
> account was built by script on both guests on 2026-09-15/16, and the second (identity) account,
> with its own delivery store, on `OutlookAI-Indexed` on 2026-09-24 - one step of the latter is an
> undocumented registry write, stated as such in §4. And rows 1, 4 and 5's "measured broken"
> `IProfAdmin` was this project's wrong IID, not Office (§3's first note). §4 is still the section
> to read if you read only one - for the opposite reason.

---

## 2. The preflight, which is the reason any of this is worth writing

**[GUEST-MEASURED]** on a fresh guest, Office LTSC 2024 build **16.0.17932.20996**, on the day this
was written:

```
DLLPath        = mapi32.dll                                   [file exists: False]
DLLPathEx      = C:\Program Files\Microsoft Office\root\VFS\ProgramFilesCommonX64\system\msmapi\1033\msmapi32.dll
                                                              [file exists: TRUE]
MSIComponentID = {6DB1921F-8B40-4406-A18B-E906DBEEF0C9}
Office Platform         = x64
PowerShell Is64BitProcess = True
```

That is the whole "if this fails, nothing below matters" check, and it passes. It says three things:

* **Extended MAPI still resolves on this build.** **[MS-DOC]** The stub `mapi32.dll` in `System32`
  finds the real provider through `HKLM\SOFTWARE\Clients\Mail\Microsoft Outlook`, preferring
  `DLLPathEx` over `DLLPath`
  (<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-stub-library>).
  `DLLPath` being an unresolvable bare filename while `DLLPathEx` carries a full path that exists
  is the **expected, healthy** shape on Click-to-Run, not a fault — Click-to-Run installs into a
  virtualised `root\VFS\` tree that the old value was never able to name.
* **The bitness matches.** MAPI is in-process **[MS-DOC]**, so an x86 host cannot load an x64
  provider. Windows PowerShell 5.1 on this guest is a 64-bit process and Office is x64, so the
  helper loads. A mismatch fails at `MAPIInitialize` with a load error, not a subtle wrong answer.
* **`msmapi32.dll` is under `1033`** — the English resource directory. **[INFERRED]** from the
  path shape: the guest's *display* language is en-GB and its *formats* are nl-NL, and MAPI resolved
  to the 1033 tree anyway, so nothing in the profile work should be locale-sensitive. Worth knowing
  because this guest is deliberately not en-US (`Testbed/MEDIA.md`).

**What the preflight does not prove.** It proves the DLL resolves. It does not prove
`MAPIInitialize` succeeds, and it does not prove `CreateProfile` writes anything. Those are the
first two things the scripts do, and they are the first two things that can fail.

---

## 3. Profiles and PST stores: Extended MAPI

> **[MEASURED] 2026-09-24 - THE BANNER BELOW HAS THE WRONG CAUSE. THE `E_NOINTERFACE` WAS THIS
> PROJECT'S DECLARATION, NOT OFFICE.** The deleted `[ComImport]` interop declared `IProfAdmin` as
> `{00020379-0000-0000-C000-000000000046}`. Microsoft's own MAPI headers
> (`github.com/microsoft/MAPIStubLibrary`, `include/MAPIGuid.h`) define
> `DEFINE_OLEGUID(IID_IProfAdmin, 0x0002031C, 0, 0)` - `{0002031C-...}` - and neither `00020379`
> nor the `0002037A` it gave `IMsgServiceAdmin` (real: `0002031D`) appears in them;
> `IMsgServiceAdmin2` was declared `0002038A`, which is `IID_IMessageRaw` (real: `00020387`,
> `MAPIAux.h`). Run on `OAI-INDEXED`, both in session 0 and session 1, 64-bit Windows PowerShell
> 5.1, Click-to-Run 16.0.17932.20996: `MAPIInitialize` = `0x00000000`, `MAPIAdminProfiles(0, out
> IntPtr)` = `0x00000000` with a non-null pointer, `QueryInterface({00020379-...})` =
> `0x80004002` (the 2026-09-16 failure, reproduced exactly), `QueryInterface({0002031C-...})` =
> `0x00000000` returning **the same pointer**, and `IUnknown` likewise. So MAPI's
> profile-administration object answers `IID_IProfAdmin` perfectly well: "`IProfAdmin` is
> unreachable from PowerShell on this build" was never measured - a declaration carrying a GUID
> that names no interface was. What this does **not** establish is that the rest of the deleted
> interop would have worked: no call ever got past the cast, so its vtable order and marshalling
> were never executed. **Nothing was rebuilt on this result, by decision** (Q67: knowledge only).
> The routes that shipped instead - `.prf`, `AddStoreEx` + rename, the registry default switch -
> are measured and stay. Full record: the REOPEN section of `Testbed/guest/OutlookMapiInterop.ps1`.
>
> **[MEASURED] 2026-09-16 - NOTHING IN THIS SECTION IS REACHABLE FROM POWERSHELL ON OFFICE LTSC
> 2024.** *(Superseded as to the cause by the note above; kept as the record.)* The first execution of this interop anywhere threw *"Unable to cast COM object … to
> interface type `IProfAdmin` … QueryInterface … `{00020379-0000-0000-C000-000000000046}` … No
> such interface supported (`E_NOINTERFACE`)"* on `OAI-UNINDEXED`, 64-bit elevated PowerShell 5.1,
> build 16.0.17932.20884. `MAPIInitialize` and `MAPIAdminProfiles` BOTH SUCCEEDED; only the
> QueryInterface failed. The cause was not established.
>
> **Everything below is kept because it is correct as documentation of the API.** The property
> tags, the flags, the `MSUPST MS` service name, the `CreateMsgService`/`ConfigureMsgService`
> sequence and the `PST_CONFIG_PRESERVE_DISPLAY_NAME` corroboration are all accurate and all still
> the right reference if this is ever revisited from C++ or from a language whose COM interop does
> not force a `QueryInterface`. What is *not* true any more is the framing: that this was the route
> to take. It was not, here.
>
> **What shipped instead**, and each is called out again in place below:
>
> * **a profile, and its PST stores named exactly** — a `.prf` import (`ImportPRF`), §3.2a;
> * **a PST into a profile that already exists** — `NameSpace.AddStoreEx` plus a rename of the
>   store's root folder, §3.4 — which is the same §3.4 that used to rule that route out;
> * **the default-profile switch** — the registry value, §7.
>
> The shared interop that wrapped all of this is **deleted**, not kept as a fallback: an
> unexercised second route is exactly the defect that produced this banner. It is in git at
> `8b610c2`. `Testbed/guest/OutlookMapiInterop.ps1` carries the leading hypothesis for the
> `E_NOINTERFACE`, the ten-line experiment that would confirm it, and the statement that nobody
> has run it.

### 3.1 Why this rather than the registry

**[MS-DOC]** `IProfAdmin` and `IMsgServiceAdmin` are the supported, documented interface for
profile administration
(<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/iprofadmin-iunknown>,
<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imsgserviceadmin-iunknown>).
The registry layout underneath is not documented, has never been documented, and is reverse
engineering all the way down — the useful published work on it is MFCMAPI and OutlookSpy, which are
diagnostic tools that *read* it. **[COMMUNITY]** Nobody has published a working end-to-end "create a
profile purely by registry writes" recipe for Office 16.x.

There is a second reason, specific to this project, and it is the stronger one: **the display name
has to be right, and MAPI sets it at creation time.** See §5.

**[MEASURED, 2026-09-17] BOTH REASONS HAVE SINCE FAILED, and the paragraphs above are kept as the
record of an argument rather than as advice.** The first reason — that the API is the supported
contract and the registry is reverse engineering — is still true *about the API*, and is simply
moot when the API cannot be called from here. The second reason has been **disproved directly**:
`Store.DisplayName` follows a rename of the store's **root folder**, `@` included, measured by
`Testbed/guest/Rename-OutlookStore.ps1` on this build, so MAPI is not the only way to a store named
exactly what was asked for. §3.4 carries that, and it is the same section that used to rule the
object-model route out on precisely this point.

What the project actually reaches for now, in the order it needs them: a **`.prf`** (§3.2a) for a
profile and its stores, and **`AddStoreEx` + rename** (§3.4) for a store going into a profile that
already exists. Both are off the registry *and* off MAPI, so neither inherits the objection this
subsection was written to make.

### 3.2 The account-less profile

**[MS-DOC]** `HRESULT IProfAdmin::CreateProfile(LPTSTR lpszProfileName, LPTSTR lpszPassword,
ULONG_PTR ulUIParam, ULONG ulFlags)`. Passing `ulFlags = 0` creates the profile **without**
`MAPI_DEFAULT_SERVICES`, i.e. without the default service set. Passing `MAPI_DEFAULT_SERVICES`
(`0x00000001`) is what adds them
(<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/iprofadmin-createprofile>).
Passing neither `MAPI_DIALOG` nor a window handle is what keeps it silent.

**This is the mandatory one.** **[REPO]** `corpus-build` refuses any profile with an account and
has no override — `CorpusSafety.EvaluateProfile` returns `ProfileCanSend` when the count is above
zero, and `ProfileUnprovable` when it cannot be read at all
(`McpServer/OutlookAI.RemediationTools/CorpusSafety.cs`).

**And here is the thing worth knowing before you build it.** **[REPO]** The predicate is
`Outlook.Application.GetNamespace("MAPI").Accounts.Count`, read over COM — I read the gatherer,
`ComCorpusMailbox.ReadProfileFacts`. It is **not** the `9375CFF0413111d3B88A00104B2A6676` registry
subkey and it is **not** the MAPI service list. That distinction matters because of a known trap:

**[COMMUNITY]** MAPI injects the Contact Address Book provider (`CONTAB`) into a bare profile the
first time that profile is opened. So a profile you created with zero services does not stay at
zero *services*.

**[INFERRED]** That should be harmless here, because `CONTAB` is an address-book provider and the
Outlook object model's `Accounts` collection contains `Account` objects — `olExchange`, `olPop3`,
`olImap`, `olHttp`, `olEas`, `olOtherAccount` — with no member for an address book provider. The
inference is from the OM's own type surface, not from a measurement.

**Settle it in one line rather than trusting the inference.** On the guest, in session 1, with the
corpus profile default:

```powershell
$o = New-Object -ComObject Outlook.Application; $o.GetNamespace('MAPI').Accounts.Count
```

Zero means the corpus profile is usable. Anything else means it is not, and the number tells you
how far off. `Testbed/guest/New-OutlookProfile.ps1 -Verify -WithOutlook` runs exactly this read and
fails on a non-zero answer, because "the generator will refuse this profile" is much cheaper to
learn now than thirteen minutes into a build.

### 3.2a What actually creates the account-less profile: a `.prf` import

**[MEASURED]** The mechanism. Write an Outlook Profile file, set
`HKCU\...\Office\<major>\Outlook\Setup\ImportPRF` to its path, delete `FirstRun` and `First-Run`
(`ImportPRF` is ignored while either exists, with no diagnostic anywhere), then start Outlook once.
`Testbed/guest/New-TierProfile.ps1` does exactly this and built its profile on **both** guests,
2026-09-15/16, Office LTSC 2024 16.0.17932 — first attempt on the second guest, from the committed
scripts, untouched by hand.

**[MEASURED]** The stores come with it, named exactly. `[ServiceN] Name=` maps through section 6's
`[Unicode Personal Folders]` block to `PT_UNICODE,0x3001` — `PR_DISPLAY_NAME_W`, the same property
§3.3's `ConfigureMsgService` route was chosen for. On that run the PST the file named was the one
Outlook used, Outlook minted no data file of its own, and `Store.DisplayName` read back **over
COM** as exactly `tier@vm.invalid`.

**[INFERRED], twice, and both are asserted by `-Verify` rather than assumed.** The measured file
carried **one** PST service and **one** POP3 account. The corpus profile's file is a strict subset
of it plus a repetition of that one service block, so:

1. **that an empty `[Internet Account List]` yields `Accounts.Count = 0`.** Corroboration, and it
   is decent: **[REPO]** `Build-Corpus.ps1`'s preflight comment records that on this project's own
   guest a profile with a PST store and no internet account held *nothing* under
   `9375CFF0413111d3B88A00104B2A6676`. **Fallback if it is false, [MEASURED]:** `outlook.exe /PIM
   <name>` produces `accounts=0` on this build, and is how both guests' corpus profiles were
   actually made — it simply cannot name a store.
2. **that `[Service List]` may name `Unicode Personal Folders` more than once.** **[MS-DOC]**
   `UniqueService=No` — which the measured template carries verbatim — is the key whose documented
   job is to permit a service block to repeat. It has not been run with two PSTs. **Fallback:** one
   store through the `.prf`, the rest through §3.4's `AddStoreEx` + rename, which needs no `.prf`
   at all.

> **[MEASURED] 2026-09-24 - OUTLOOK CLEARS IT.** On `OutlookAI-Indexed`, Office 16.0.17932:
> the tier profile's own import (2026-09-15, `ImportPRF` set and `First-Run` deleted by
> `New-TierProfile.ps1`) had left `Setup` with **no `ImportPRF` value and `First-Run` present**
> (`REG_BINARY`, 16 bytes) nine days later, with no script on record that cleared it. Then four
> controlled imports the same day (the identity-account `.prf` three times - twice from the
> experiment's copy, once rendered by `Testbed/guest/Add-IdentityAccount.ps1` - and a two-account
> variant once), each with `ImportPRF` set and `First-Run` deleted beforehand: after each Outlook
> start `ImportPRF` was **gone** and `First-Run` was **back**, with the identical 16 bytes each
> time. So the import is one-shot on this build and the "rebuilt at every start"
> hazard below does not occur; `-ClearImportPrf` is belt and braces rather than a necessity.
> **What the measurement also showed, and it is the half of the hazard that stays true:** an
> import with `OverwriteProfile=Yes` rebuilds the profile from the file and **drops every store
> the file does not name** - the two-account variant left the old tier PST outside the rebuilt
> profile. So a *repeated* import still costs the corpus; it just has to be repeated by someone.

**[UNKNOWN] `ImportPRF` is read at EVERY Outlook start, and nobody has checked whether Outlook
clears it.** This is the one hazard in the route that can cost data. The file carries
`OverwriteProfile=Yes` — correct for a rebuild, because it is what makes a repeat import converge
instead of leaving a `Backup Of <name>` profile. If the value is not cleared, every later start
**re-imports and rebuilds the profile**, and a store attached or filled afterwards stops being part
of it. No `.pst` is deleted; the corpus simply is not in the profile any more, and the repair is a
~13-minute rebuild of 20,000 items. Nothing in this repository has ever looked —
`New-TierProfile.ps1` did not either. The remedy is
`New-OutlookProfile.ps1 -ClearImportPrf -Execute`, and `-Verify` warns by name while the value is
still set. `Testbed/README.md` §4b-i and `Docs/live-tier-on-the-vm.md` §2.5a both carry it.

**[MS-DOC]** One authoring rule worth repeating because breaking it fails silently: sections 6 and
7 are the property mappings and Microsoft says *"You typically do not modify existing entries"* —
a property with no mapping there is written **nowhere**, with no error. `New-OutlookProfile.ps1`
generates its `.prf` rather than substituting into a committed template, because the number of
`[ServiceN]` blocks varies and a token substitution cannot express that; sections 6 and 7 are
emitted verbatim from the file Outlook measurably processed. The generator is a pure function, so
the whole emitted file is asserted by `-SelfTest` without a guest.

### 3.3 Adding a PST, and the display name

**[MS-DOC]** The PST message service is named **`MSUPST MS`** for Unicode PSTs and `MSPST MS` for
ANSI ones (<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pst-providers>).
Unicode is the only sensible choice here: **[REPO]** the corpus store is ~400 MB and `Build-Corpus`
writes 20,000 items, which is far past the 2 GB ANSI ceiling's comfort zone and past the point
where the ANSI format is a good idea at all.

**[MS-DOC]** `IMsgServiceAdmin::CreateMsgService(LPTSTR lpszService, LPTSTR lpszDisplayName,
ULONG_PTR ulUIParam, ULONG ulFlags)` adds the service;
`IMsgServiceAdmin::ConfigureMsgService(LPMAPIUID lpUID, ULONG_PTR ulUIParam, ULONG ulFlags,
ULONG cValues, LPSPropValue lpProps)` configures it. `ulFlags = 0` on both suppresses the property
sheet; `MAPI_DIALOG` is what would raise one
(<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imsgserviceadmin-configuremsgservice>).

The properties that matter, **[MS-DOC]** from the PST provider's own header prose (`MSPST.H`, quoted
in <https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pst-configuration>):

| Property | Tag | What it does |
| --- | --- | --- |
| `PR_PST_PATH_W` | `0x6700001F` | The `.pst` file path. Creates the file when absent. |
| `PR_DISPLAY_NAME_W` | `0x3001001F` | **The name Outlook shows in the folder list.** |
| `PR_PST_REMEMBER_PW` | `0x67010003` | Not used here. |

**The load-bearing sentence, and the reason this route beats the object model** — *it does not beat
it any more, and it cannot be called from here at all; see §3's banner and §3.4. The property fact
itself is correct and is the same `PR_DISPLAY_NAME_W` that a `.prf`'s `[ServiceN] Name=` writes
(§3.2a), which is why the `.prf` route inherits the behaviour this paragraph describes:* **[MS-DOC]**
the
provider's documented behaviour is that the configure routine takes `PR_DISPLAY_NAME` and *sets it
on the message store object*. **[MS-DOC]** The corroboration is a flag that exists only to switch
that behaviour off: `PST_CONFIG_PRESERVE_DISPLAY_NAME`. A flag whose whole job is "do not apply the
display name this time" is proof that not passing it applies the display name.

**[COMMUNITY]** `CreateMsgService` does not hand back the UID of the service it created. The
published pattern is to snapshot `GetMsgServiceTable` before and after and diff on
`PR_SERVICE_UID`. **[MS-DOC]** `IMsgServiceAdmin2::CreateMsgServiceEx` returns the UID directly and
avoids the diff — but it is a later interface, obtained by `QueryInterface`, and if that IID is
wrong or unavailable you get `E_NOINTERFACE`. The draft script asks for `IMsgServiceAdmin2` first
and falls back to create-then-diff, and says in its output which one it used, because those two
paths have different failure modes and an operator should not have to guess which one ran.

### 3.4 `Namespace.AddStoreEx` — the easy route that does not do the hard part

**[MS-DOC]** `NameSpace.AddStoreEx(Store, Type)` adds a PST to the *running* profile;
`olStoreUnicode = 2`
(<https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.addstoreex>).

It is genuinely easier, and it is the right answer for a store whose name does not matter. It is
**not** the answer here, for two reasons:

* **It cannot name the store.** **[MS-DOC]** `Store.DisplayName` is read-only. **[COMMUNITY]** The
  usual workaround is to rename the store's *root folder*, and whether `Store.DisplayName` then
  follows could not be established by anybody. **[COMMUNITY]** `PropertyAccessor.SetProperty` on
  `http://schemas.microsoft.com/mapi/proptag/0x3001001F` at store level is reported blocked.
* **It cannot bootstrap.** It needs Outlook already running on a loaded profile, so it can never
  create the first one.

**[INFERRED]** Its real use here is **verification**, not construction: after the MAPI route has
run, opening Outlook on that profile and reading `Session.Stores` is the only check that answers
"what does Outlook actually show", which is the only form of the question the tests care about.

> **[MEASURED] 2026-09-15 — THE FIRST BULLET IS NOW FALSE, AND THIS IS THE ROUTE THAT SHIPPED.**
> The question nobody could settle has been settled by running it.
> `Testbed/guest/Rename-OutlookStore.ps1`, Office LTSC 2024 build 16.0.17932:
>
> ```
> set root.Name        was 'Outlook Data File', now 'tier@vm.invalid'
> Store.DisplayName    'tier@vm.invalid'
> account afterwards   SmtpAddress='tier@vm.invalid' DeliveryStore='tier@vm.invalid'
> ```
>
> **`Store.DisplayName` FOLLOWS a root-folder rename**, the `@` survives, and the rename does not
> break an account's delivery-store binding. Everything else in the two bullets stands:
> `Store.DisplayName` is still read-only, `PropertyAccessor` is still blocked, and `AddStoreEx`
> still cannot name anything *itself*. What changed is that it no longer has to — **add unnamed,
> then rename** is two calls and needs no MAPI.
>
> The second bullet is also still true and is now a *scoping* statement rather than an objection:
> `AddStoreEx` cannot bootstrap, so it is used for a store going into a profile that **already
> exists**, and §3.2a's `.prf` is what creates the first one. That is exactly the split between
> `Testbed/guest/Add-OutlookPstStore.ps1` and `New-OutlookProfile.ps1`.
>
> **Two things to carry if you use this route.** **[MEASURED, once, unexplained]** `AddStoreEx` has
> been seen *spinning* on this project's guest — `Docs/autonomous-session-log.md` records it beside
> one other unexplained COM block and is careful to say spinning rather than blocked. A COM call
> owns its thread, so nothing can time it out from PowerShell; do **not** `taskkill` OUTLOOK.EXE
> (mailbox-safety rule 7), and revert the checkpoint if it never returns. And **[COMMUNITY]** the
> idempotency detail below is load-bearing here rather than optional: match on `Store.FilePath`,
> resolved and case-insensitive, because the display name is the thing being set.

**[COMMUNITY]** One idempotency detail worth copying: match an already-added store on
`Store.FilePath`, resolved and compared case-insensitively — **not** on `DisplayName`, which is
exactly the thing under test and is user-editable besides.

---

## 4. The mail account — SUPERSEDED: both accounts are created by script, measured

> **[MEASURED] 2026-09-15/16 and 2026-09-24 — THE CONCLUSION OF THIS SECTION WAS WRONG. BOTH
> POP3 ACCOUNTS THE TESTBED NEEDS ARE CREATED BY SCRIPT, WITH NO GUI AND NO PAID COMPONENT.** Its
> three reasons are each still true about what they describe; none of them turned out to be the
> constraint. The original section is kept below, unchanged, as the record of the argument.
>
> **The tier account (the first) — [MEASURED] on both guests, 2026-09-15/16.**
> `Testbed/guest/New-TierProfile.ps1` with `Testbed/guest/tier-profile-forcepst.prf` and
> `HKCU\...\Outlook\ForcePSTPath` produced `Accounts.Count = 1`, `SmtpAddress = 'tier@vm.invalid'`,
> `AccountType = 2` (POP3), a bound `DeliveryStore` and its Drafts folder resolving - on
> `OutlookAI-Indexed` after two variants, on `OutlookAI-Unindexed` at the first attempt, from the
> committed scripts. **The trick this section could not predict:** the `.prf` names **no** PST
> service and **no** `DefaultStore`. The profile then has no store at all, Outlook **mints** one
> (`Outlook.pst`, "Outlook Data File", in `ForcePSTPath`) and **binds** the account to it - the one
> step a text file cannot perform. Naming the store in the `.prf` is exactly what left
> `DeliveryStore` NULL on the first attempt. The minted store gets its name afterwards, by the
> root-folder rename of §3.4.
>
> **Where the binding lives — [MEASURED] 2026-09-24, `OutlookAI-Indexed`.** In the account's own
> subkey under `...\Profiles\<profile>\9375CFF0413111d3B88A00104B2A6676`, as two `REG_BINARY`
> values Outlook writes at import time:
>
> | Value | Content, compared byte for byte with COM |
> | --- | --- |
> | `Delivery Store EntryID` | the store's `Store.StoreID` - identical (114 bytes for `C:\OutlookAI-Tier\Outlook.pst`: a wrapped PST store EntryID naming `mspst.dll`, the PST provider UID and then the file path in UTF-16) |
> | `Delivery Folder EntryID` | that store's Inbox `EntryID` - identical (24 bytes) |
>
> They are `PROP_ACCT_DELIVERY_STORE` and `PROP_ACCT_DELIVERY_FOLDER`, and on this build the
> registry values carry those **names**, not the `00180102`-style tags - so a verifier looking for
> a value named `00180102` reports a working account as unbound. `New-TierProfile.ps1 -Verify`
> does exactly that.
>
> **The identity account (the second) — the same way CREATES it, and cannot give it its own
> store. [MEASURED] 2026-09-24 on `OutlookAI-Indexed`**, both variants with no PST service and no
> `DefaultStore`:
>
> 1. **`OverwriteProfile=Append` into the existing tier profile.** The account appears (a new
>    POP3 subkey, `clsid {ED475411-...}`), the tier account and its store are untouched, and no
>    `Backup Of` profile is made - but Outlook binds it, at import time, **to the profile's default
>    store**: the tier store, same `StoreID`. Nothing is minted.
> 2. **One `.prf` carrying both accounts into a rebuilt profile (`OverwriteProfile=Yes`).**
>    Outlook minted **one** store (`Outlook1.pst`, because `Outlook.pst` already existed) and
>    bound **both** accounts to it; the previous tier store dropped out of the profile.
>
> So "a store Outlook mints is a store Outlook binds" holds, but Outlook mints **one store per
> profile, and only when the profile has none**, and binds every account a `.prf` adds to the
> profile's default store. This section's own last line on the `.prf` route - *"Even a PRF that
> works leaves both accounts delivering to the profile default"* - was exactly right about the
> second account.
>
> **What gives the identity account its own store — [MEASURED] 2026-09-24, and ONE step of it is
> undocumented.** `Testbed/guest/Add-IdentityAccount.ps1`, in phases, from checkpoint
> `CP-08B-RESTORED-BEFORE-IDENTITY`:
>
> 1. append the account with `Testbed/guest/identity-account.prf` (the variant-1 file) - [MEASURED];
> 2. `Testbed/guest/Add-OutlookPstStore.ps1` attaches `C:\OutlookAI-Tier\identity.pst` by
>    `NameSpace.AddStoreEx` and names it `identity@vm.invalid` by a root-folder rename -
>    [MEASURED], and the first guest run of that script: `AddStoreEx` returned promptly, no spin;
> 3. read that store's `StoreID` and its Inbox's `EntryID` over COM - [MEASURED], read-only;
> 4. with Outlook closed, write those two byte strings into the identity account's
>    `Delivery Store EntryID` and `Delivery Folder EntryID` - **[MEASURED] to work, [UNDOCUMENTED]
>    as a mechanism.** The object model's `Account.DeliveryStore` is read-only and Microsoft
>    documents only the GUI's *Change Folder*. It writes only bytes Outlook itself reported, only
>    into values Outlook itself created, and reads them back; it is the step to re-check first on
>    another Office build.
>
> Result, read over COM and read again after a guest restart (so across Outlook's own shutdown and
> a fresh start): `Accounts.Count = 2`, `Stores.Count = 2`; `OutlookAI tier sink` delivering to
> `tier@vm.invalid` (`C:\OutlookAI-Tier\Outlook.pst`), `OutlookAI identity sink` delivering to
> `identity@vm.invalid` (`C:\OutlookAI-Tier\identity.pst`); each Drafts folder resolving; two
> distinct delivery stores for two accounts. The registry `Email` values are `tier@vm.invalid`
> and `identity@vm.invalid`. **`Account.SmtpAddress` itself could not be relied on over COM on these
> guests** - see the Object Model Guard note in `Docs/live-tier-on-the-vm.md` section 8.
>
> **What still stands from the reasoning below, precisely.** (1) The object model has no
> `Accounts.Add` and `Account.DeliveryStore` is read-only - TRUE, and it is why step 4 is a
> registry write. (2) MAPI has no POP3 message service - TRUE, and irrelevant: the `.prf` processor
> creates POP3 accounts through Outlook's own account manager. (3) The registry has no published
> recipe for a whole account and DPAPI-seals passwords - TRUE, and irrelevant: nothing here
> synthesises an account or a password; Outlook creates the account from the `.prf`, and the one
> registry write copies a binding Outlook has already expressed in bytes. (4) A `.prf` cannot carry
> the binary EntryID - TRUE, and got round twice: by letting Outlook mint and bind (first account)
> and by copying the bytes Outlook reports (second). **The Redemption paragraph is moot**: the
> repository's Dependencies rule (decided 2026-09-15) excludes it outright, whatever it costs.
>
> *The original section follows, unchanged.*

This is the finding the maintainer most needs, and it is a refusal rather than a recipe. Three
independent lines of evidence converge, which is why it is stated this strongly.

**1. The object model is read-only for accounts.** **[MS-DOC]** `Namespace.Accounts` returns an
`Accounts` collection with `Count` and `Item` and no `Add`
(<https://learn.microsoft.com/en-us/office/vba/api/outlook.accounts>). **[COMMUNITY]** Microsoft's
own published sample carries a comment saying the account objects cannot be created from the OM.
`Account.DeliveryStore` is **[MS-DOC]** read-only, which matters specifically: "deliver new messages
to" is the one setting `Docs/live-tier-on-the-vm.md` §2.8 calls "the failure to bet on", and the OM
can read it but cannot set it.

**2. MAPI cannot, because POP3 stopped being a MAPI message service.** **[COMMUNITY]** Account
administration on Outlook 2010 and later lives behind `IOlkAccountManager`, which is undocumented,
and whose creation method sits in one of fourteen placeholder vtable slots. That is why §3's
`IMsgServiceAdmin` route can add a *store* and cannot add an *account*: they stopped being the same
kind of object. There is no `MSPOP3` message service to name.

**3. Direct registry synthesis has no working published recipe on 16.x, and the passwords are
DPAPI.** **[COMMUNITY]** The account subkeys under `9375CFF0413111d3B88A00104B2A6676` are readable
and their names are well known, but the password blobs are DPAPI-protected per Windows user per
machine, so they cannot be authored offline or copied between guests. **[REPO]** This repository
already corroborates the read half independently: `SignatureCatalog.ReadProfileAccountValueSets`
walks `HKCU\<OutlookRoot>\Profiles\<DefaultProfile>\9375CFF0413111d3B88A00104B2A6676\*` and reads
`Account Name`, `New Signature` and `Reply-Forward Signature`, decoding REG_SZ **or** REG_BINARY
UTF-16LE. So the shape is confirmed from inside this codebase — what is missing is a recipe for
*writing* a whole working account, not knowledge of where one lives.

**The one free candidate, and why it is a spike rather than a route.** **[MS-DOC]** PRF files
document POP3 sections, including `PROP_ACCT_POP3_SERVER`, `PROP_ACCT_POP3_USER_NAME`,
`PROP_ACCT_SMTP_SERVER` and friends. The documentation is the Office Resource Kit's, and it stops
at Office 2013 — **[COMMUNITY]** there is no source at all discussing PRF POP3 sections on 16.x,
which is not evidence that it broke but is a complete absence of evidence that it works.

**The structural objection is worse than the staleness.** **[MS-DOC]**
`PROP_ACCT_DELIVERY_STORE` is `PT_BINARY` — an EntryID. A `.prf` is an INI file of text
assignments, and the EntryID of a store does not exist until that store exists in that profile on
that machine. **[INFERRED]** So a PRF structurally cannot carry the delivery-store binding, and
that binding is precisely the one `Docs/live-tier-on-the-vm.md` §2.8 says is the failure to bet on
and §2.8b requires again for the identity account. Even a PRF that works leaves both accounts
delivering to the profile default.

**What this means for the build.** Steps 4 and 7 of `Testbed/README.md` do not both become
automated. Step 4 does. Step 7 keeps a GUI pass whose scope is now small and precisely bounded:
**add two POP3 accounts in the Account Settings wizard and set each one's "Deliver new messages
to".** Everything around it — both profiles, all four PSTs with their exact names, the default
switch, the signature — is scripted. That is a much better position than "step 4 and step 7 are by
hand", and it is an honest one.

**Not designed around, deliberately: Redemption.** **[COMMUNITY]** Dmitry Streblechenko's
Redemption exposes profile and account administration (`RDOProfMan`) and is the one component that
plausibly closes this gap. It is **commercial**, it is a redistributable dependency, and this
project has none today. Whether to buy it is the maintainer's call and is being put to them
separately. Nothing in the draft scripts assumes it, and if the answer is yes, the seam to replace
is one script.

---

## 5. The `@` question — ANSWERED: yes, measured twice

> **[MEASURED] 2026-09-15 — OUTLOOK ACCEPTS `@` IN A STORE DISPLAY NAME.** Settled empirically, on
> Office LTSC 2024 build 16.0.17932, by **two independent routes**, and by neither of the probes
> this section proposed:
>
> 1. **through a `.prf` import** — the profile Outlook built from `Testbed/guest/tier-profile.prf`
>    carries a store whose `Store.DisplayName` reads back **over COM** as literally
>    `tier@vm.invalid`;
> 2. **through a root-folder rename** — `Testbed/guest/Rename-OutlookStore.ps1` renamed a store
>    from `Outlook Data File` to `tier@vm.invalid`, `Store.DisplayName` followed, and the account
>    afterwards still reported `SmtpAddress` and `DeliveryStore` as that name.
>
> Both readings are over COM, which is the reading this section's last paragraph correctly insisted
> on. The `[INFERRED]` guess below — that `PR_DISPLAY_NAME` is free text and `@` would simply be
> accepted — turned out right, but it was a weak claim when it was made and it is superseded rather
> than vindicated.
>
> **The probe described below has been REMOVED from `Add-OutlookPstStore.ps1`.** Two reasons, and
> the second is the one that would have forced it regardless: the question is answered, and the
> probe created and deleted a throwaway profile — and **deleting a profile has no free route** now
> that `IProfAdmin` is unreachable (§3's banner). A probe that cannot clean up after itself is
> worse than no probe.
>
> `Docs/live-tier-on-the-vm.md` §8 item 2 is the single authoritative statement of this answer and
> carries the one caveat that still stands (a validation rule in the Data File Properties *dialog*
> remains untested, which is irrelevant while stores are only ever created by script).
> Everything below is kept as the record of what was known before it was run.

### 5-historical. The `@` question as it stood — still open, and now cheap to settle

`Docs/live-tier-on-the-vm.md` §8 item 2 and `Testbed/README.md` §6 item 10 both carry it: the hub
store must be named literally `test@vm.invalid` and the identity store `identity@vm.invalid`,
because several tests hand the store's display name to `NewDraft` as an address. Whether Outlook
accepts `@` in a store display name gates the whole draft family.

**Nothing found settles it.** No Microsoft documentation states a character restriction on a store
display name, and no community report describes `@` being rejected or transformed. **[INFERRED]**
`PR_DISPLAY_NAME` is a free-text MAPI string property with no documented character class, so the
likeliest outcome is that it is simply accepted — but "I found no evidence of a restriction" is a
weak claim and should not be built on.

**What changed is who can answer it and how fast.** It was a five-minute GUI job that nobody had
done. It is now a script: `Testbed/guest/Add-OutlookPstStore.ps1 -NameProbe` creates a throwaway
PST, asks for a display name containing `@`, reads back what MAPI and then Outlook actually report,
and removes it again. The probe reports one of three outcomes, and the third is the interesting one:

1. the name comes back **exactly** — the gate opens, nothing else changes;
2. the call **fails** — the gate closes, and the hub and identity stores need a different naming
   scheme, which is a test-side change;
3. the name comes back **silently transformed** — the worst case, because it would otherwise be
   found much later as a store the tests cannot find by name.

**[INFERRED]** Run it against MAPI *and* over COM. MAPI reporting the name it was given proves the
property was stored; only Outlook reporting it proves the tests will see it, because
`Store.DisplayName` is what they read.

---

## 6. The signature: already solved, inside this repository

§2.8b requires the identity account to carry a signature configured for New mail, because
`LiveDraftOptionsTests` asserts `SignatureInjected` and reads the injected HTML back.

**[REPO]** This does not need a new script, because the product already does it, tested:

* `SignatureManager` writes the file set — `.htm` + `.txt` + `.rtf` — under
  `%APPDATA%\Microsoft\Signatures`, deriving the renditions it was not given.
* `ProfileSignatureDefaultsStore.WriteDefault` writes the per-account binding as **REG_SZ** into
  `New Signature` / `Reply-Forward Signature`, under the account's subkey of
  `9375CFF0413111d3B88A00104B2A6676`. Its writes are deliberately narrow: only those two value
  names, only on subkeys carrying an SMTP-shaped `Account Name`, never creating a subkey.
* Both are reachable from the shipped MCP tool `manage_signature`, whose `set_default_for` argument
  takes `{account, scope}` with scope `new` | `reply` | `both`.

**[REPO]** There is a comment in `SignatureManager` worth repeating because it is exactly this
machine: on Microsoft 365 Apps 2303+ roaming signatures can overrule local files unless
`DisableRoamingSignatures = 1`; **on Office LTSC, local files are authoritative**. The guests are
LTSC 2024, so the local file set is the whole story and no roaming setting is needed.

`Testbed/guest/Set-AccountSignature.ps1` therefore does not implement anything — it drives the
shipped tool over stdio, using the JSON-RPC pattern already proven in
`Testbed/guest/Invoke-GuestMeasure.ps1`. **That is the right shape under this project's own
mailbox-safety rule 1**, which says mutation happens only through tested helper code or the shipped
MCP tools, never improvised. It also has an ordering dependency that will bite: the binding is
written onto an account subkey, so **the account must exist first** — which means after the GUI
pass of §4, not before.

One thing to get right, from `Docs/live-tier-on-the-vm.md` §2.8b: this signature is **ordinary user
data**, deliberately *not* one of the `OutlookAI-McpTest-` prefixed signatures the suite creates and
deletes. The SHA-256 signature-directory snapshot requires it to come back bit-identical, which it
will, because nothing in the suite writes to it. Do not name it with that prefix.

---

## 7. Switching the default profile

> **[MEASURED] 2026-09-24 - "DOES NOT WORK ON OFFICE LTSC 2024" BELOW IS THE WRONG DIAGNOSIS.**
> The `E_NOINTERFACE` came from the interop declaring `IProfAdmin` with the GUID `00020379`;
> `IID_IProfAdmin` is `0002031C` (Microsoft's `MAPIGuid.h`), and a `QueryInterface` for it on the
> object `MAPIAdminProfiles` returns succeeds on this build. §3's first note has the HRESULTs.
> `SetDefaultProfile` is therefore *untested*, not broken. The registry route stays the one in
> use: it is measured, and nothing needs the MAPI one.
>
> **[MEASURED] 2026-09-16 - THE MAPI ROUTE BELOW DOES NOT WORK ON OFFICE LTSC 2024.** The first
> execution of `Invoke-WithProfAdmin` anywhere threw *"Unable to cast COM object … to interface type
> `IProfAdmin` … QueryInterface … `{00020379-0000-0000-C000-000000000046}` … No such interface
> supported (`E_NOINTERFACE`)"* on `OAI-UNINDEXED`, 64-bit elevated PowerShell 5.1, build
> 16.0.17932.20884. `MAPIInitialize` and `MAPIAdminProfiles` BOTH SUCCEEDED; only the QueryInterface
> failed. The cause was not established and was deliberately not chased.
>
> **The registry half of this section is the route that shipped**, and it is measured working on the
> same guest the same day. Everything below about `SetDefaultProfile` is kept because it is correct
> as documentation of the API - it is simply not reachable from PowerShell on this build. The same
> gateway blocks `New-OutlookProfile.ps1` and `Add-OutlookPstStore.ps1`, which is tracked separately.

**[MS-DOC]** `IProfAdmin::SetDefaultProfile(LPTSTR lpszProfileName, ULONG ulFlags)`, `ulFlags = 0`
(<https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/iprofadmin-setdefaultprofile>).

**[REPO]** The registry side is confirmed from inside this repository twice over, which is unusual
and worth using: `OutlookProfileRegistry` builds the Outlook root at
`HKCU\Software\Microsoft\Office\<major>\Outlook` with the major **detected at runtime**, and
`DefaultProfile` is a value directly under it — `SignatureCatalog` and `SignatureManager` both read
it that way in shipped code. `McpServer/README.md` records a measurement taken on the dev machine
on 2026-08-17: the real `16.0` key held 31 subkeys and 4 values *including* `DefaultProfile` and
`Profiles`, while the decoy `15.0` and `17.0` keys held one subkey each and no values. So on Office
16.x the hive is the Office one, not the old Windows Messaging Subsystem path.

**The prompt is a separate setting from the default**, and both have to be right, because
`Docs/live-tier-on-the-vm.md` §2.5 is blunt about it: a prompting profile cannot be driven over COM.
**[COMMUNITY]** The value is `PickLogonProfile` (DWORD) under the Outlook root — `0` for "always use
this profile", `1` for "prompt for a profile to be used". **[INFERRED]** It is not documented by
Microsoft as a supported setting; it is however trivially verifiable by reading it back, which is
what the script does, and its failure direction is benign — get it wrong and Outlook prompts, which
is loud rather than silent.

**Outlook must not be running when this changes.** **[COMMUNITY]** The profile is read at logon and
a running Outlook will write its own view back at shutdown, silently reverting you.
`Set-DefaultOutlookProfile.ps1` refuses outright when it sees `OUTLOOK` running, rather than doing
the work and hoping. **[REPO]** It does **not** kill it: the project's mailbox-safety rule 7 forbids
`taskkill` on `OUTLOOK.EXE` outright, so refusing and telling the operator is the only correct
behaviour.

---

## 8. Getting Extended MAPI out of PowerShell 5.1

**[COMMUNITY]** `Add-Type -TypeDefinition` works on a machine with **no .NET SDK**, because
`csc.exe` ships with the .NET Framework itself. That is the constraint that made this approach
viable at all — **[REPO]** `Testbed/README.md` §2 is explicit that the guest has a runtime and no
SDK, and that nothing can be compiled there.

Five constraints, all of which the draft interop obeys, and each of which produces a distinctive
failure if broken:

1. **C# 5 only.** The in-box compiler is old. No string interpolation, no expression-bodied
   members, no `nameof`. Breaking this gives a compile error at dot-source time — loud, immediate,
   harmless.
2. **A type cannot be redefined in a live session.** Re-dot-sourcing after an edit throws "type
   already exists". The interop guards on the type being present and skips `Add-Type`; while
   iterating the C# itself you must start a fresh PowerShell.
3. **Declare every vtable slot, in order, including deprecated ones.** A COM interface declaration
   is positional. Omitting a method — even one nobody calls, even a deprecated one — shifts every
   method after it and you call the wrong function pointer. **[INFERRED]** That failure does not
   look like a mistake: it looks like MAPI returning nonsense, or an access violation that takes
   the whole PowerShell process down with no error text. This is the single most likely cause of an
   unexplained crash when someone first runs these scripts.
4. **`ULONG_PTR` is `IntPtr`, not `uint`.** The `ulUIParam` parameter is pointer-sized. Declaring it
   `uint` on x64 corrupts the stack for every argument after it.
5. **ANSI throughout (`CharSet.Ansi` / `LPStr`).** **[MS-DOC]** `MAPI_UNICODE` is documented as not
   supported on the service-admin calls. Passing wide strings without the flag hands MAPI a string
   it reads as ANSI; passing the flag gets `MAPI_E_BAD_CHARWIDTH`.

**Threading.** **[COMMUNITY]** Windows PowerShell 5.1 hosts in STA by default, which is what MAPI
wants. Launching it `-MTA` gives `RPC_E_CHANGED_MODE`. **[INFERRED]** Keep every MAPI call on one
thread: no `Start-Job`, no runspaces, no `ForEach-Object -Parallel` (which 5.1 does not have
anyway). The scripts are written as straight-line single-threaded code for this reason and not by
accident.

**`MAPIInitialize` / `MAPIUninitialize` must be paired.** **[MS-DOC]** Every script does this in a
`finally`.

---

## 9. What I could not establish

This list is a deliverable. Somebody is going to run these scripts on a guest, and the useful thing
is knowing where the unknowns are before the first revert.

1. **Whether Outlook accepts `@` in a store display name.** §5. Untestable without a machine; the
   probe exists and is one command.
2. **The exact IID of `IMsgServiceAdmin2`.** I could not verify it against a primary source. The
   script treats `QueryInterface` failing as expected-and-handled and falls back to
   create-then-diff, so a wrong IID costs a code path, not a run.
   **ANSWERED 2026-09-24, and the fear was justified one line higher up.** Microsoft's
   `MAPIAux.h` (MAPIStubLibrary) has `DEFINE_OLEGUID(IID_IMsgServiceAdmin2, 0x00020387, 0, 0)`;
   the interop had `0002038A`, which is `IID_IMessageRaw`. But the IIDs of `IProfAdmin`
   (`0002031C`, declared `00020379`) and `IMsgServiceAdmin` (`0002031D`, declared `0002037A`)
   were wrong too, and those had no fallback - which is the whole 2026-09-16 `E_NOINTERFACE`
   (§3's first note).
3. **Whether `CreateProfile` with `ulFlags = 0` really yields a profile Outlook opens without a
   wizard.** The flag semantics are documented; what Outlook's *first run* does when it meets a
   service-less profile is not. This is where I would bet on the first surprise.
4. **How Outlook's first-run wizard is suppressed on LTSC 2024.** Still unrecorded — it is
   `Docs/live-tier-on-the-vm.md` §8 item 7's own remaining gap and this work did not close it. The
   commonly cited values (`DisableOfficeFirstRun`, `Outlook\Setup\First-Run`, `ZeroConfigExchange`)
   are **[COMMUNITY]** and version-drifted, and I will not assert them for 16.0.17932.
5. **Whether `PR_DISPLAY_NAME` passed to `ConfigureMsgService` survives Outlook's first open of that
   profile.** The documentation says it is set on the store object. Whether Outlook later rewrites
   it from the PST's internal name is not documented either way. **[INFERRED]** the existence of
   `PST_CONFIG_PRESERVE_DISPLAY_NAME` argues it does survive, but the verify step reads it back
   through Outlook rather than trusting that.
6. **Whether a `.prf` import does anything at all on 16.x.** §4. No source newer than 2013.
   **ANSWERED: it does** - it built the tier profile and its POP3 account on both guests
   (2026-09-15/16), and appended a second account to an existing profile (2026-09-24). §4.
7. **Whether `CONTAB` injection changes `Accounts.Count`.** §3.2. Inferred to be no; one line
   settles it.
8. **Every timing and ordering question.** Whether `ConfigureMsgService` needs the profile to be
   closed, whether two PSTs can be added in one MAPI session, whether the default-profile switch
   takes effect without a logoff. All unknown, all cheap to learn on the guest.

---

## 10. Where I expect these scripts to fail first

> **[MEASURED] SCORED 2026-09-16/17. Item 1 happened, in a way this list did not predict; items 3,
> 4 and 5 are now moot.** Worth keeping as a calibration record rather than rewriting: the list was
> right that MAPI was the first thing to break and wrong about how. It predicted a *compile* failure
> or an access violation from a vtable off-by-one. What actually happened was a clean, well-behaved
> `E_NOINTERFACE` from the CLR's `QueryInterface`, with `MAPIInitialize` and `MAPIAdminProfiles`
> both returning S_OK — **the failure mode the mitigations were least designed for, because every
> one of them assumed the problem would be loud.** The `-Preflight` described in item 1 would not
> have caught it either: it resolved the DLL, which was never in doubt.

In order of my confidence that it will happen:

1. **The C# interop will not compile, or will compile and crash the process.** §8 constraints 1 and
   3. A vtable off-by-one is an access violation with no error text. **Mitigation in the draft:**
   every interface declares every documented slot with the unused ones named and commented, the
   interop is dot-sourced separately so a compile failure happens before anything touches MAPI, and
   `New-OutlookProfile.ps1 -Preflight` does `MAPIInitialize` / `MAPIUninitialize` and nothing else —
   run that first, on a checkpoint you are willing to lose.
   **[MEASURED] HAPPENED, differently.** It compiled and it did not crash; the QueryInterface
   failed. **[MEASURED] 2026-09-24: because the `[Guid]` on the declaration was wrong** -
   `00020379` where Microsoft's `MAPIGuid.h` has `IID_IProfAdmin` = `0002031C` (§3's first note).
   So this prediction was right in kind after all - a wrong constant in hand-written interop that
   fails without looking like a mistake - just on the IID rather than on a vtable slot. The
   interop is now deleted (§3's banner) and `-Preflight` no longer touches MAPI at all —
   it reads the Office hive, the `Setup` key, the profiles present, the current `ImportPRF` value
   and whether Outlook is running, and makes no COM call, so it needs no checkpoint.
2. **`CreateProfile` succeeds and Outlook still shows a wizard.** §9 items 3 and 4.
   **Still a live risk, against the `.prf` route rather than this one** — `ImportPRF` is ignored
   while `FirstRun` or `First-Run` exists, with no diagnostic anywhere, which is the same shape of
   failure. `New-OutlookProfile.ps1` deletes both and reads back that they are gone.
3. **The display name does not come back the way it went in.** §5 outcome 3. **[MEASURED] MOOT** —
   the name comes back exactly, over COM, by both surviving routes.
4. **The PRF spike does nothing, silently.** Fully expected; it is labelled a spike for that reason.
   **[COMMUNITY]** `outlook.exe /importprf` starting the full Outlook UI rather than running
   headless is a specific risk worth watching for on a guest where nothing may open a window.
   **[MEASURED] WRONG, and this is the single largest correction in this document.** The PRF route
   worked on both guests, and it is no longer a spike — it is how a profile and its named stores
   are created (§3.2a). The headless worry did not arise because the project never calls
   `/importprf`: it sets the `ImportPRF` value and lets the next ordinary Outlook start consume it.
5. **`DeleteProfile` returns `S_OK` and deletes nothing** when the profile is in use. The scripts
   never trust that HRESULT and verify through the profile table instead.
   **[MEASURED] MOOT, and replaced by something worse: there is no way to delete a profile at all.**
   `IProfAdmin::DeleteProfile` went with the rest of the interface, removing the key by hand is
   reverse engineering with no published recipe, and the Mail control panel is a human. Nothing in
   the build needs it — the only caller was the `@` probe — so it is recorded as **a capability with
   no route** rather than worked around.
6. **Path casing.** **[COMMUNITY]** A PST path spelled differently in two profiles gives
   `MAPI_E_FAILONEPROVIDER`. The scripts normalise once with `GetFullPathName` and reuse that exact
   string. **Still true and still done**, and it matters for a second reason now: Outlook normalises
   the `FilePath` it reports back, and store identity is matched on that.

---

## 11. What is not automatable, stated plainly

> **[MEASURED] 2026-09-24 - NOTHING IN THE BUILD SEQUENCE BELOW IS LEFT TO A GUI.** The paragraph
> under this note, and the table's "GUI, once per guest" row, are superseded by §4: the tier account
> is built by `New-TierProfile.ps1` (both guests), and the identity account with its own delivery
> store by `Testbed/guest/Add-IdentityAccount.ps1` (`OutlookAI-Indexed`, with one undocumented
> registry step, §4). The maintainer's call that the note at the end of this section waited for was
> made (Q64, option 4): §4 is rewritten, with the original kept as history. What is still NOT
> scripted anywhere is deleting a profile - see `Testbed/guest/OutlookMapiInterop.ps1`, whose REOPEN
> section now records that the documented route (`IProfAdmin::DeleteProfile`) was never actually
> dead, only unbuilt.

**Creating a POP3 mail account, and binding its delivery store.** §4. Not "hard" — closed. The
object model has no create, MAPI has no service to create, the registry has no published recipe and
DPAPI-sealed passwords besides, and the one documented text format cannot express a binary EntryID.
Three routes, three different reasons, same answer.

So the honest build sequence for a guest is:

| Step | How |
| --- | --- |
| Corpus profile, account-less, **with its PSTs named exactly** | `New-OutlookProfile.ps1` — a `.prf` import (§3.2a). Then **`-ClearImportPrf -Execute`**, which is not optional: see §3.2a's `ImportPRF` hazard |
| Tier profile, **and its POP3 account** | `New-TierProfile.ps1` with `tier-profile-forcepst.prf` — **[MEASURED] working on both guests**, not `New-OutlookProfile.ps1`, which this row named by mistake |
| A further PST into a profile that already exists | `Add-OutlookPstStore.ps1` — `AddStoreEx` + root-folder rename (§3.4) |
| Default-profile switches | `Set-DefaultOutlookProfile.ps1` — the registry value (§7) |
| The **identity** account + its delivery store | ~~**GUI, once per guest**, as far as this document establishes — but see the note below~~ **`Add-IdentityAccount.ps1`** (phases Import / CaptureStore / Bind / Verify, with `Add-OutlookPstStore.ps1` between Import and CaptureStore) — **[MEASURED] 2026-09-24**, one step undocumented (§4) |
| Identity signature | `Set-AccountSignature.ps1`, *after* the accounts exist |
| Renaming any store afterwards | `Rename-OutlookStore.ps1` — **[MEASURED]** |

> **AN OPEN QUESTION FOR THE MAINTAINER, FLAGGED RATHER THAN ANSWERED HERE.** §4 concludes there is
> **no** free programmatic route to a POP3 account with a bound delivery store, and its three
> reasons — the OM has no `Accounts.Add`, MAPI has no POP3 message service, the registry has no
> recipe and DPAPI-seals the passwords — are each still correct. Its **conclusion**, however, looks
> superseded by a measurement taken after it was written: `Testbed/guest/tier-profile-forcepst.prf`
> plus `ForcePSTPath` produced `Accounts.Count = 1`, `SmtpAddress = 'tier@vm.invalid'`
> (`AccountType = 2`, POP3), a **bound** `DeliveryStore`, and `DeliveryStore.GetDefaultFolder(Drafts)`
> resolving — on both guests, with no GUI and no paid component. The trick §4 could not have
> predicted is that the `.prf` must name **no** PST service, so that Outlook **mints** the account's
> delivery store: a store Outlook mints is a store Outlook **binds**, which is the one step a text
> file cannot perform, and naming the store in the file is exactly what left `DeliveryStore` NULL on
> the first attempt.
>
> **§4's body is deliberately left untouched pending that call.** How much of it should be rewritten
> — and whether the second (identity) account is now scripted too, which nobody has tried — is a
> decision for the maintainer rather than for this edit. `TODO.md` already records the tier account
> as **DONE, built by script**.
>
> **ANSWERED 2026-09-24.** The maintainer chose to rewrite §4 with the measured answers and keep the
> original as history; that is done. The identity account **has** now been tried: the same `.prf`
> route creates it but binds it to the tier store, and giving it its own store took an
> `AddStoreEx` plus a registry write of the binding Outlook reports over COM - §4 has both halves.

**[INFERRED]** One consolation that is worth more than it looks: a Hyper-V checkpoint taken
immediately after the GUI pass makes it a **once-per-guest-lifetime** cost rather than a
once-per-rebuild one. `Docs/live-tier-on-the-vm.md` §2.11 already names checkpoints around exactly
the steps most likely to need redoing, and this is now the strongest candidate on the list:
everything before it is scripted and reproducible, so the checkpoint only has to preserve the part
that is not.
