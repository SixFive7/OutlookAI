# Running the live tier somewhere other than the maintainer's machine

**Who this is for.** Someone who has a Windows machine with Outlook on it and wants to run
OutlookAI's `Category=Live` tests without touching a production mailbox. It assumes no
knowledge of how the tier grew up.

**Read `CLAUDE.md`'s Mailbox Safety section first.** Nothing here overrides it. Every rule in
it applies on a test machine too, and the guards described below are what enforce it.

---

## 1. The two facts that shape everything else

**The live tier writes to exactly one mailbox and reads several.** The one it writes to is the
**hub**, named in a gitignored settings file. Every other store the settings mention is watched
by the store-count tripwire, and what a test may do to it is decided in code by
`StoreWriteAllowlist` rather than by anybody remembering: the other configured PRIMARY stores get
draft-create and draft-delete and nothing else (one tagged, never-displayed draft each, for the
identity tests), and delegate/shared mailboxes and any store not in the settings at all get
nothing, ever.

**Most live tests cannot honestly run without a real profile.** Of 127 live test methods, 31
are `LiveTier=Portable` and 96 are `LiveTier=ProfileBound`. That is not a policy choice; it is
what the tests do. Drafts need a resolvable mail Account object, searches need the Windows
Search index to have published the store, delegate tests need a delegate mailbox, and one
test actually sends mail. A machine without those cannot run them, and pretending otherwise
would mean tests that pass without testing anything.

So the split is: **the VM runs the Portable subset routinely, and the whole tier runs on the
maintainer's own profile before a release.** Section 5 is the second half of that.

---

## 2. What a second machine needs

### 2.1 The machine

* Windows with Outlook installed and a **mail profile that opens without prompting**.
  The suite connects in-process over COM and cannot answer a dialog.
* An **interactive logon session**. Outlook never finishes starting in session 0, so anything
  driving it must run in session 1. On a Hyper-V guest reached by PowerShell Direct that means
  a scheduled task with `-LogonType Interactive`, not a direct remote call.
* At least one store the suite may write to (the hub). A PST is fine. Section 2.3 says which
  store that should be, and it is not the one most people would pick.
* The .NET SDK, and a clone of this repository.

### 2.2 The settings file

Create `McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json`. It is
gitignored, and it must stay that way: it names real stores and this repository is public.
Without it the whole live tier refuses to start.

A test machine declares itself with `machineProfile: "Portable"`:

```json
{
  "machineProfile": "Portable",
  "testHubStoreDisplayName": "Outlook Data File",
  "expectedStoreDisplayNames": [ "Outlook Data File", "OutlookAI Bystander" ],
  "expectedDelegateStoreDisplayNames": []
}
```

| Field | What it is | Required |
| --- | --- | --- |
| `machineProfile` | `Production` or `Portable`. Defaults to `Production` when absent, so an older settings file keeps the validation it was written under. | no |
| `testHubStoreDisplayName` | Display name of the store the suite may write to. Exactly as Outlook shows it. | **yes** |
| `expectedStoreDisplayNames` | Every store the count tripwire watches. Include the hub. | **yes** |
| `expectedDelegateStoreDisplayNames` | Delegate/shared mailboxes. Watched, never written, and their folder hierarchy is allowed to appear and disappear (it syncs lazily). Empty on a test machine. | no |
| `probeTerm` | A word proven to hit this machine's search index. | Production only |
| `subjectOnlyProbe` | Coordinates of a real population whose term is in the subject and not the body. | Production only |
| `delegateNestedFolderProbe` | A delegate folder Outlook nests and the index publishes flat. | never |

The last three describe real mail. A test machine has none of it, which is exactly why
`machineProfile` exists: before it, every machine had to supply all three, and a requirement
that cannot be met honestly gets met dishonestly. A block that is present must be **complete**
on any profile: three fields out of four reads as configured and behaves as absent.

### 2.3 The store layout, and why two PSTs

The tripwire **exempts the hub**: the hub is where the suite writes, its churn is tagged, and
the zero-artifact sweep polices it. So a machine whose only store IS the hub gives the tripwire
nothing to watch. It will census, report zero failures, and be structurally incapable of
reporting anything else. That is a real configuration and it is not wrong, but it must not be
mistaken for the guard having passed.

**Recommended layout: two PSTs, with the CORPUS as the hub.**

* `Outlook Data File` - the generated corpus, and the **hub**. This is the way round it has to
  be: the Portable subset is mostly scans and sweeps, and every one of them targets the hub
  store. `LiveResumableScanTests` pages through the hub, `LiveExhaustiveSearchTests` bounds a
  scan to a hub folder, `LiveSweepScopeTests` sweeps the hub's arrival-path folders. Point the
  hub at an empty store and they all degrade to their "corpus too small" early return: green,
  and proving nothing. The suite's writes into the corpus store are tagged, swept and
  recoverable from a checkpoint.
* `OutlookAI Bystander` - a second, small store, listed in `expectedStoreDisplayNames` and never
  the hub. It exists so the tripwire has something to watch, and because it is small it is walked
  ITEM BY ITEM rather than counted, which is the half of the guard that was rewritten. A few
  hundred items in it is better than none: an empty store exercises the identity path over
  nothing.

Add the second PST through Outlook itself (File > Account Settings > Data Files > Add). Do not
improvise a script: creating stores is not something the tested helpers do, and mailbox
mutation from ad-hoc shell code is the thing that once destroyed real mail.

One consequence worth knowing: stores in `expectedStoreDisplayNames` other than the hub are
granted **draft-create and draft-delete** rights by `StoreWriteAllowlist`, because the identity
tests create one tagged, never-displayed draft in each business account and delete it again.
On a Portable machine those tests are `ProfileBound` and will not run, so the grant is unused.

---

## 3. Running the Portable subset

```
dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
  --filter "Category=Live&LiveTier=Portable"
```

19 tests. Nothing in that set needs a mail account, the search index, a delegate mailbox or
transport. Some of them do need the interactive desktop (`Requires=InteractiveDesktop`): they
open Outlook windows and take screenshots.

To run one class - the two probes the session is blocked on, for example:

```
dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
  --filter "Category=Live&FullyQualifiedName~LiveTableSortProbeTests"

dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
  --filter "Category=Live&FullyQualifiedName~LiveResumableScanTests"
```

**A filtered run is fully guarded.** It takes the census, runs the health preflight, and
verifies at the end of whichever collection the filter left last. That was not true before
2026-08-19: verification lived in one collection's teardown, so any run that did not include
that collection paid for a baseline and threw it away. See section 6.

### 3.1 Which tests are in which bucket, and how to find out

The classification is two traits on the test itself, not a list in a document that can drift:

* `LiveTier` is the selector. Exactly one value per live test: `Portable` or `ProfileBound`.
* `Requires` is the reason, and it carries weight. A `ProfileBound` test must name at least one
  capability a test machine cannot have - `SearchIndex`, `MailAccount`, `Transport`,
  `MultipleStores`, `DelegateStore`, `SmallHubStore`, `ProbePopulation` - and a `Portable` test
  must name none of them. Three further values - `InteractiveDesktop`, `AddInRegistry` and
  `OutlookInstance` - describe things a test machine CAN have and constrain only how the run is
  launched.

`OutlookInstance` arrived on 2026-08-23 with the tier-3 correction, and it is why the Portable
subset grew by eleven. The T3 stdio classes spawn the real server, which spawns a COM host,
which attaches to whatever Outlook is on the machine - so eleven tests that called
`outlook_health`, `list_accounts` or `search` were reaching a real mailbox from a run filtered
`Category!=Live`. They are now `ComHostSupervisionLiveTests`, `OutlookAvailabilityLiveTests` and
`OutlookHealthLiveToolShapeTests`, all `Portable`: what they need is an Outlook, not this
Outlook, so the VM runs them unchanged and they are among the easiest things it can prove.

Two mechanisms keep that from drifting back, both described in `McpServer/README.md`:
`McpStdioClient` refuses to send a `tools/call` for `outlook_health`, `list_accounts` or
`list_folders` unless the test declares mailbox contact, and
`LiveTierInventoryTests.EveryStdioTestReachingOutlook_DeclaresIt` reads that declaration back out
of the compiled IL, so a new method in an old class is caught as well as a new class.

`T1/LiveTierInventoryTests` enforces all of that in CI, together with the rule that every live
class sits in a registered collection. So a live test added later cannot be left unclassified,
and a test cannot be quietly reclassified to make the VM subset look bigger: the reason has to
be named, and the reason is checked.

To see the sets without running anything:

```
dotnet test <csproj> --list-tests --filter "Category=Live"                     # 127
dotnet test <csproj> --list-tests --filter "Category=Live&LiveTier=Portable"   # 31
```

`--list-tests` discovers and does not execute, so it is safe against any mailbox.

---

## 4. What the guards do, and what to check afterwards

Four guards arm themselves; none needs remembering.

1. **Health preflight** (`LiveOutlookPreflight`). Asks Windows whether Outlook's UI thread is
   servicing its message queue before any COM call. Refuses the tier in milliseconds when it is
   not. Exists because a wedged Outlook once turned a live run into a 22-minute hang, and an
   aborted run skips its cleanup - which is how tagged items were left in a real mailbox.
2. **Store-count tripwire** (`LiveStoreCountTripwire`). Censuses every watched store before the
   first live collection and after the last. Fail-closed: no census, no live tier.
3. **Write allowlist** (`StoreWriteAllowlist`). A write aimed outside the hub throws instead of
   running.
4. **Signature snapshot** (`SignatureDirectorySnapshot`). SHA-256 before and after; the user's
   real signatures must be bit-identical.

**What to read in the output.**

* `[tripwire] live-test settings: machineProfile=..., stores=N, ...` - the first line. If it
  names the wrong machine's settings, stop there.
* `[tripwire] baseline: 2 stores, 14 mail folders, identified 8 folder(s)/312 item(s), 431 ms.`
  The identified count is what was walked ITEM BY ITEM, and it is the number that says how much
  of the guard is live. **Zero identified items means the guard can only see counts**, which on
  a two-store machine means the bystander store is empty or the hub is the only store. Section 6
  says what to expect.
* `[tripwire] post-run census in T ms (identified ...); 0 failure(s), K note(s).` Notes are
  benign - mail arriving, filing, hub churn. Failures throw.
* `PROVED NOTHING:` - a test that ran but found no population to test. On a Portable machine
  that is expected for the handful of tests that discover their own population; on a Production
  machine it now throws instead.

**And check that a verification happened at all.** A run that prints a `baseline` line and no
`post-run census` line did not compare anything.

---

## 5. Running the whole tier against the maintainer's own profile before a release

This is unchanged, and it is deliberately still available: the Portable subset is 31 tests of
115, and the other 96 are the ones that exercise the index, the accounts, the delegate stores
and the send path.

```
dotnet test McpServer/OutlookAI.McpServer.Tests/OutlookAI.McpServer.Tests.csproj \
  --filter "Category=Live"
```

Before starting:

* the settings file on that machine says `machineProfile: "Production"` (or omits the field,
  which means the same), and carries `probeTerm` and a complete `subjectOnlyProbe`;
* Outlook is running and responsive, with **no unsent compose windows open and an empty
  Outbox** - the lifecycle collection quits and restarts Outlook gracefully, and it refuses to
  do so otherwise;
* nobody is going to work in those mailboxes for the duration. A person filing mail during the
  run reads to a before/after census exactly like a runaway test, so the tripwire fails and
  hands over the EntryIDs rather than guessing.

Afterwards:

* zero tagged artifacts, proven by the teardown sweep across Drafts, Inbox, Sent Items,
  **Outbox**, Deleted Items and the Sync Issues subtree;
* `0 failure(s)` from the post-run census;
* the signature directory verified bit-identical;
* no `PROVED NOTHING:` lines. On a Production profile a missing population now throws, so if
  one appears it means the machine or the settings have drifted.

---

## 6. The count tripwire's first run: what to expect

The tripwire was rewritten on 2026-08-19 and, as of this document, **has never executed**. Its
first run will be on a test machine, which is the right place for one. What follows is
predicted by reading the code, so treat it as something to check rather than something known.

**With one PST that is also the hub.** `PlanFor` gives the hub a count-only plan and `Evaluate`
exempts it, so: every mail folder is counted, `0 folder(s), 0 item(s) identified`, and no
failure is reachable. The guard runs and proves nothing. This is the configuration to avoid.

**With the recommended two PSTs.** The corpus store is the hub, so it is counted and exempt -
which is fine, because it is the store the suite writes to. Everything the guard actually
decides happens on the bystander store.

* The identity budget is 500 items per folder and 3,000 per store, and a small bystander store
  is entirely inside both, so every one of its folders is walked item by item. That is the
  rewritten half of the guard, exercised for real.
* Cost: the census reads `Folder.Items.Count` per folder, which is a table row count and not a
  walk, plus four late-bound property reads per identified item. Against two local PSTs expect
  the whole baseline in well under a second. On the maintainer's five-store profile it is a
  different number entirely - up to 3,000 items walked per non-hub store, twice per run - and it
  has never been measured.
* Nothing about the corpus should make it fire falsely: no mail arrives on a machine with no
  accounts, and the suite writes only to the hub.

**Were the corpus made a NON-hub store instead**, the numbers are worth knowing because they say
what a 20,000-item store buys the guard: nothing. All four populated folders (Inbox 10,912 /
Sent 4,964 / Deleted 2,467 / Junk 1,663) are above the 500-item per-folder limit, so all four
fall back to counts; Deleted Items and Junk are self-pruning and excluded from identity anyway.
The identity path would walk only the store's small and empty folders and the 3,000-item store
budget would go almost entirely unspent. **A corpus is the wrong shape for this guard**; a few
hundred items in a small store is the right one.

**Two things to watch on the first run.**

* **The Junk folder may not be marked self-pruning.** The census marks volatile folders by
  asking the store for its default Deleted Items, Junk and sync-issue folders. A PST may refuse
  `GetDefaultFolder` for Junk, in which case a generator-made "Junk Email" folder is treated as
  an ordinary folder and a decrease in it would FAIL rather than be noted. Nothing prunes it on
  a VM, so this should stay theoretical - but if the tripwire ever fires on Junk, this is why.
* **A move whose destination was only counted cannot be exonerated.** The census can prove an
  item was filed rather than deleted only when BOTH folders were walked item by item. An item
  moved from a small folder into one above the budget will be reported as removed. On a test
  machine nobody is moving mail, so this matters on the production run, not here.

---

## 7. Known limits, honestly

* **The Portable subset is 31 tests.** It contains the two acceptances the project is currently
  blocked on (`LiveTableSortProbeTests`, `LiveResumableScanTests`), the sweep-scope and
  sweep-cache behaviour, the exhaustive folder-bounded scan, the signature lifecycle and the
  show-me UI paths. It does not contain anything that proves the index, the accounts, the
  delegate stores or the send path.
* **A machine with no mail accounts cannot create a draft at all.** `NewDraft` resolves an
  Account object by SMTP address and refuses when none matches, which is what puts the whole of
  the draft, update/discard, HTML-draft and send families in `ProfileBound` regardless of
  anything else. Adding one dummy account would move a large part of the tier onto a test
  machine - and it is a decision with a catch, because the corpus generator refuses to run at
  all unless the profile has **no accounts whatsoever**. Generate the corpus first, checkpoint,
  then add the account.
* **`testHubStoreDisplayName` doubles as an SMTP address** in several tests (`to: Hub`,
  `FindAccountBySmtp(Hub)`). A PST display name is not an address, so those tests could not pass
  on a PST hub even with an account present.
* **Several `ProfileBound` tests assume a tiny hub** and would break their paging assertions
  against a 20,000-item one (`Phase7LiveMcpToolShapeTests` asserts the hub holds between 2 and 99
  items; `LiveMailServiceTests.ListFolders...` asserts the hub tree fits one page). They carry
  `Requires=SmallHubStore` and are excluded from the Portable subset, so this is a conflict only
  if the whole tier is ever pointed at a corpus hub. **It is the reason the two machines cannot
  share one settings shape**, not a reason to move the corpus off the hub on a test machine: the
  Portable scans and sweeps all target the hub and prove nothing against an empty one.
