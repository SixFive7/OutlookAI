# Changelog

Everything notable that has happened to OutlookAI, newest first. The format is
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

**A version here is a release tag, and nothing else.** `Tools/Publish-Release.ps1` derives it when a
release is cut: the latest release's major, minor and patch plus the bump asked for, then the
commit count. It stamps that version and the date on the section below `## Unreleased`, as
`## v<version> - <date>`. No version is typed into a project file, into this file or anywhere else.

**Entries are written as the work lands, never reconstructed at release time.** The release script
refuses a release while `## Unreleased` holds no entries.

**Every entry opens with one icon from a fixed palette, then a bold one-sentence headline, then the
whole of what happened.** The icon says what the entry *is*, not which group it sits under, so a fix
to a test carries the test icon and a fix that removes something carries the removal icon. The
groups are Keep a Changelog's: Added, Changed, Deprecated, Removed, Fixed and Security, in that
order, each at most once in a release. `Tools/Publish-Release.ps1` turns each headline into one line
of the GitHub release, with a *read more* link to that entry's lines in this file, and refuses an
entry that is not in this shape. The releases up to v3.1.0.325 were written one line per change,
before this format, and stay as they were published.

| Icon | Meaning | Icon | Meaning |
|---|---|---|---|
| ✨ | new capability | 🐛 | fix |
| 🔧 | behaviour or configuration change | 🔒 | security or permissions |
| 🗑️ | removal or deprecation | 💥 | breaking, or the reader must act |
| 📝 | documentation | ✅ | tests and the gate |
| 📦 | packaging, installer, release pipeline | ⚡ | performance |
| ♻️ | refactor with no behaviour change | ⬆️ | dependency move |

## Unreleased

OutlookAI is an Outlook add-in with an AI writing sidebar. It also installs a local mail server that
lets an AI assistant such as Claude Code search and read your mail and prepare drafts for you to
send.

Most of the changes since v3.1.0.325 are in that server. When it cannot finish something, it now
says what it did and whether your mail changed. It works on POP3, IMAP and local data-file (.pst)
mailboxes without creating folders in them, and it no longer hangs when Outlook stops answering. The
assistant can now read OutlookAI's audit log and continue an exhaustive search where it stopped.

In the add-in, OutlookAI Settings has tabs, and you can edit every prompt and quick button. You can
also pick the Claude model, check for updates yourself, and see and change Outlook's Cached Mode
settings without administrator rights.

Releases are now built and signed on the maintainer's own computer, so they no longer carry GitHub's
build-provenance attestation. Many of the entries are about the test machines the project now runs
its mailbox tests on, and change nothing you install.

### Added

- ✨ **The add-in's AI chooses a mail's signature by itself, by the language of the mail it answers.**
  Decided by the maintainer on 2026-11-17 (Q158, Q162 (b)): OutlookAI Settings has one tick box on
  its Claude Code tab, "Allow the AI to automatically choose the correct signature", on unless you
  untick it, and no signature editor, list or menu, since Outlook has its own. With it on, the AI
  picks the signature when it writes or rewrites the body in the task pane and when you start a reply
  or a forward, from the signatures the sending account can use - by the language of the mail being
  answered, so that a Dutch mail gets the Dutch signature, for one. A draft opened again, yours or
  one an agent wrote, keeps its signature until the AI rewrites its body. It changes a signature
  Outlook or the AI put in; it adds none to a mail that has none, chooses none while a policy or a
  deployed setting names a default, and never replaces one a person chose or changed, which it
  checks again after the AI answered. Nothing happens as a mail is sent, and nothing in a compose
  window nobody sees, such as one an agent's draft is written in. The choice uses the rule the MCP
  server's drafts use for the account's signatures (`SignatureSelector.Candidates`) and the prompt
  of "Select the best signature"; when it
  asks the AI at all is `SignatureAutoChoice`, in OutlookAI.Core, which T1 `SignatureAutoChoiceTests`
  pins. A reply or forward costs one more request to Claude Code as it opens.

- ✨ **A stuck Outlook is reported with its process ID, how to end it, and what ending it risks (Q177).**
  When Outlook stops responding, or has been starting for two minutes without finishing, every mail
  tool answers `OutlookUnresponsive` with a report that says what is stuck and why, the process ID,
  how to end it in Task Manager and what that can cost: unsaved text in a compose window, mail
  waiting in the Outbox, a data file checked at the next start. For an Outlook with no window, the
  kind OutlookAI starts for an agent, it says the user cannot see or close it and that starting
  Outlook does nothing until it is ended, and no longer that restarting Outlook fixes it or that it
  clears by itself. The advice tells the assistant to tell the user and to end Outlook only if the
  user asks; OutlookAI never ends it (Q177 (a)). An Outlook that hangs before it has created its
  windows used to read as starting for ever; past 120 seconds it is reported as stuck, and so is a
  connection that has waited that long in one start-up step, while an Outlook that had windows and
  closed them is shutting down, not stuck (Q176 (c)). `outlook_health` gains an `outlook.start`
  block: the process and its start time, whether OutlookAI started it and for which server, whether
  it ever showed a window, since when it has not responded, the start-up step of the last
  connection, whether the add-in has started, and the start log's path (Q178 (e)). Decided by the
  maintainer on 2026-11-18, after the hang of 2026-11-15: an Outlook OutlookAI had started with no
  window deadlocked, and the user could neither see it nor close it.

- ✨ **A start log records every step of starting Outlook and every hang, with no mail content (Q178).**
  `%LOCALAPPDATA%\OutlookAI\outlook-start.log`, beside the audit log, gets a line for each step a
  connection takes into Outlook, each wait for it to settle and why the wait ended, each change of
  the verdict in front of a request, and each COM host ended at its deadline. Keys are constants and
  every value is reduced to a token, so no sentence, subject or store name can be written; a failure
  is logged by exception type and HRESULT, never by its message. About 1 MB is kept. A per-user
  record, `HKCU\Software\OutlookAI\OutlookStart`, holds which server started the running Outlook,
  which start-up step a connection is in and since when, and when Outlook was first seen not
  responding or with a window, so every server can report it (Q176 (c)). Nothing OutlookAI wrote
  could say what happened inside the Outlook that hung on 2026-11-15, nor what started it again.

- ✨ **`find_folder` finds folders by name and gives their IDs, and never acts (Q136).**
  Every tool now addresses a folder by its ID, so an assistant asked about "my Projects folder" first
  looks the folder up. `find_folder` takes a name, part of one, or the end of a path, and returns
  candidates: each folder's ID, name, store and store ID, and its path for reading, with exact name
  matches first, then folders whose path ends with the text, then names that contain it. It changes
  nothing and acts on nothing; when two folders share a name, its advice tells the assistant to pick
  by store and path, and to ask the user when that does not settle it. It walks the same folder tree
  `list_folders` does, and can be narrowed to one store with `store_id`.

- ✨ **`create_folder` makes one mail folder under a folder given by its ID (Q134).**
  `create_folder(parent_id, name)` makes exactly one folder and returns its ID, ready for
  `move_mail`. A `/` in the name is part of the name, never a path separator: `Q3/Q4` makes one
  folder called `Q3/Q4`. A top-level folder is made under the store's root folder, whose ID
  `list_accounts` and `list_folders` now show as `rootFolderId`. It refuses, with nothing created, a
  parent in Deleted Items or under it, the Outbox, a folder that is not a mail folder, and a parent
  that already has a folder of that name, comparing names as Outlook does, without regard to case;
  that refusal gives the existing folder's ID. If Outlook fails while making the folder, the parent
  is looked at again, so the answer says whether a folder was made. Every folder made is in the
  audit log.

- ✨ **A search of one folder in a local data file works while Outlook is not running (Q132).**
  The folder's ID carries the folder's node id, and the Windows search index keeps that number on
  the folder's own row (`System.ProviderItemID`), so the search finds the folder in the index
  without Outlook: the store by the identity its mail carries, then the folder's row. It searches
  where the index files the folder's mail at that moment, also in the seconds after a rename or move
  before the index catches up, which a search by the folder's new name could not. Measured on the
  indexed test machine in October: correct after renames and moves, about a tenth of a second. A
  folder of an Exchange mailbox still needs Outlook, which is started when needed; while Outlook is
  not answering, such a search is refused with a message saying so. Measured on the Exchange test
  machine in November, that index row carries the offline file's own node number, which an
  Exchange folder's ID does not hold and which changes when the file is rebuilt. Nothing remembers
  where a folder was: the ID is looked up every time.

- ✨ **A signature can be read in full and changed in place, pictures included.**
  list_signatures with a name returns all three versions Outlook keeps of a signature - the HTML one
  for HTML mail, the RTF one for Rich Text mail and the plain text - and every picture with its file
  path. manage_signature then changes only what it is asked to: the versions sent back replace their
  own files, and an update that sends one sends every version the signature has; a replace edit
  changes a piece of text in every version that has it and keeps its formatting; and pictures are
  added, replaced or removed one file at a time. Each file keeps the encoding it was stored in. In
  the RTF version the pictures are embedded, and they show as [[picture:1]] tokens: each token left
  in place becomes the same picture again, byte for byte, and [[picture-file:logo.png]] embeds a new
  one. A new signature can be made with pictures, and its RTF version then embeds them too.

- ✨ **A default signature can be set or cleared on its own, for a named profile and account.**
  manage_signature has two new actions, set_default and clear_default. A default could only be set
  while creating or rewriting the signature itself. Every default change names its Outlook profile and
  its account, because one address can have different defaults in two profiles, and nothing is
  guessed. clear_default clears only the default it names. The answer gives the value before each
  change, so it can be put back.

- ✨ **A draft can go without a signature, and every draft says which signature it got and why.**
  The draft tools take signature "none" for no signature, and update_draft takes it to remove the
  draft's signature. Left out, a draft gets the sending account's default; a named signature must be
  one the sending account can use. The answer names the signature, where it came from (asked for, the
  account's default, or none), the reason in one sentence and the profile, and suggests the
  account's own default when the draft got something else. A reply from a shared or
  delegate mailbox gets no default, as before, and now says so. The tool descriptions and the
  server's instructions spell all of this out.

- ✨ **`audit_log` pages through any number of entries.**
  When more match than fit on one page, the answer carries nextToken, and passing it back as
  resume_token returns the next older page, until no nextToken comes back. Entries recorded in the
  same millisecond are no longer skipped, which paging by time could not avoid, and a token is
  refused with a clear reason if the log was archived or edited in between.

- ✨ **OutlookAI Settings shows Outlook's Cached Mode settings and who sets them, and can change them.**
  There are four: how much mail Outlook keeps offline in accounts you add later, caching shared
  folders and other people's mail in new profiles, and keeping shared mailboxes fully cached. Each
  row shows what Outlook uses, the value OutlookAI sets, and whether that is in effect. The offline
  period offers the eleven periods Microsoft documents, 3 days to All, 3 and 5 years among them, and
  nothing else (Q143). A value your organization sets by policy, through Microsoft 365 Cloud Policy
  or Group Policy, comes first in Outlook, so the row says "Set by your organization" and cannot be
  changed. An account you already have keeps its own offline period, and the window says how to set
  it in Outlook's Account Settings. None of it needs administrator rights (Q145 (a)).

- ✨ **Your AI assistant can read OutlookAI's audit log, so it can tell you what it did.**
  The new `audit_log` tool returns the record OutlookAI keeps of every change it makes - drafts
  created, revised and discarded, attachments saved, mail moved or archived, each step of a send,
  signature changes, mails opened on screen - newest first, from every session on this computer. It
  can narrow that to a time window, to particular kinds of change, or to one message, following the
  message through a move even though moving changes its id; that also lets it check whether a
  request whose outcome was reported as unknown actually took effect. It only reads: it never
  changes the log, never gets in the way of OutlookAI adding to it, never starts or touches Outlook,
  works with Outlook closed, and reads only the current log - never an older copy that has been
  renamed or archived. What it returns is the log's own metadata, never a subject or a message body,
  and lines it cannot read (such as the log's oldest lines, from before it had its current format)
  are counted rather than guessed at. The tool is marked read-only in the information it gives the
  assistant, so tools that sort read-only from write-capable actions can recognise it.

- ✨ **Every failure carries a plain outcome value an assistant can act on without reading the wording.**
  The values are "unchanged" (your mail is exactly as it was), "applied" (it DID happen - only the
  reporting failed) or "unknown" (nobody can say; look before repeating the call). It is the only
  way to state the middle case at all: a move that succeeded and could not be written to the audit
  log was being reported as a move that failed, because the only field available was a yes/no.

- ✨ **An exhaustive search can be continued where it stopped.**
  That mode reads your folders through Outlook directly, and it stops when it hits its ten-minute
  allowance or its hundred-result limit - on a real mailbox with a 108,000-item archive it reached
  three folders out of thirty-two before the clock ran out. Until now the answer said so honestly
  and there was nothing you could do about it: raising the result limit is capped, and running the
  same search again re-walks the same folders and stops in the same place. Every scan that stops
  early now hands back a short continuation handle; passing it back to the same search picks up
  exactly where it left off, and you repeat that until the handle stops coming, which is the signal
  that the whole scope really was covered. It is not a shortcut: each page is still reported as
  incomplete, because the page in your hand is not the whole answer.

- ✨ **A continued scan reports how far it has got in terms you can act on without the handle.**
  It reports which folder it stops in, how many folders of the total are finished, how many results
  the whole chain has returned so far, and the date it will resume from. If the handle expires or
  the server restarts, that is enough to carry on by hand with the ordinary folder and date
  arguments, so nothing already paid for is thrown away.

- ✨ **A continued scan tells you when the folder tree moved underneath it between pages.**
  That covers folders added, renamed, moved or deleted. Folders that appeared are scanned rather
  than skipped, folders that left had already been covered, and the one case that really does lose
  mail, the folder the scan was part-way through disappearing, is named as such instead of passing
  silently.

- ✨ **The breadth of the new-mail check is a plain value an assistant can act on.**
  Unless you name a folder, the check covers Inbox, Sent Items, Deleted Items and Junk Email only,
  and not their subfolders - so mail a server-side rule files elsewhere is invisible until the index
  catches up. That was stated only as a sentence in the answer; it is now also a plain value beside
  it, so an assistant can notice the limit and search the folder directly instead of assuming it
  looked everywhere.

- ✨ **`outlook_health` reports the largest answer from Outlook so far, beside its size limit.**
  It also counts anything refused for exceeding that limit, so a limit that was set by estimate can
  be judged against what really happens on your mailbox.

- ✨ **A search result that is not an ordinary mail says what it is.**
  Bounce reports, read receipts, meeting requests, responses, posts and sharing invitations now come
  back labelled, and the answer says how many of its hits are not plain mail - so the assistant does
  not tell you "you have four mails about the invoice" when one of them is a delivery receipt.

- ✨ **A search reports what the index left out of it.**
  It gives the number of rows it looked at, the number it refused, and - this one matters - whether
  it ran out of candidates while filtering, which is the one way a search list can be short of
  matches without saying so. That last fact was previously buried in a sentence rather than being
  something the assistant could act on.

- ✨ **Every AI instruction can be your own, in new Prompts and Buttons tabs in OutlookAI Settings.**
  Rename the sidebar's quick buttons, rewrite the instruction behind any of them, reorder them,
  delete the ones you never use and add your own; edit the four prompts that wrap every request; and
  reset a single button, a single prompt, or the whole button set back to how it shipped. A button
  is its name, so renaming a shipped button gives you a custom one. Only what you actually change is
  stored, which means anything you leave alone keeps following the built-in default and still
  improves with updates, while anything you edit stays exactly as you wrote it. Changes apply to
  your next action without restarting Outlook, and every open compose window picks them up at once.

- 🔒 **The prompt editor warns before you weaken a prompt.**
  If an edited preamble no longer tells the AI to ignore instructions hidden inside an email you
  received, or no longer demands plain text with no markdown or HTML, the editor says so next to the
  text. It is advice and never blocks a save - the first protects you from a malicious email
  steering the assistant, the second is what keeps code fences and stray tags out of your mail.

- ✨ **You choose which Claude model writes your mail, in a Model group on the Claude Code tab.**
  By default OutlookAI now chooses nothing and lets Claude Code decide, following your own model
  setting and picking up new models without an update - where before it was locked to one model that
  would have stopped working when that model retired. You can pin a family instead (opus, sonnet,
  haiku, fable), or type a specific model id. If Claude Code rejects your choice it quietly answers
  on its own default; OutlookAI now tells you when that happens instead of letting a substituted
  model pass unnoticed.

- ✨ **AI assistants drafting through the mail tools write mail your way too.**
  When an assistant drafts a reply or a new mail through the mail tools, it is now handed your own
  writing prompt - the same text the sidebar uses, exactly as you edited it - and asked to compose
  the body again to follow it. This happens once per session and again whenever you change your
  rules, so an edit takes effect immediately instead of waiting for the next session. Your rules can
  be as long as you like and are never copied into the tool definitions, so nothing gets silently
  cut.

- ✨ **You can check for a new version yourself, from the sidebar or OutlookAI Settings.**
  There is now a "check for updates" link under the version in the writing sidebar, and a "Check for
  updates" button in OutlookAI Settings, so you no longer have to wait up to ten minutes for the
  next automatic check. Both say "checking…" while one is running - including one started from the
  other place - and go quiet again when it finishes.

- ✨ **OutlookAI Settings shows the version and update state, as the sidebar does.**
  A new "Version and updates" section tells you which version you are running and when OutlookAI
  last managed to look for a newer one, the same as the writing sidebar has always shown. When a
  check fails, the reason is written out in full here rather than hidden behind a link. Both
  indicators now take their wording from one place, so they cannot disagree.

- ✅ **A hunting run puts a debugger on Outlook to catch the mso40uiwin32client.dll crash in a dump.**
  Two Outlook crashes in that module on 2026-11-18 left no dump, though Windows Error Reporting's
  local dumps were armed, and the evidence says the dump was lost inside WER.
  `Invoke-LiveTierOnGuest.ps1 -HuntFirstChance` attaches
  `Testbed/guest/Watch-OutlookFirstChance.ps1`, Windows' own debugging API through `Add-Type`, to
  every Outlook of the suite's session. It writes one full dump at the first first-chance access
  violation in `mso40uiwin32client.dll`, before any handler runs, and one at any second-chance
  exception, and passes every exception back to Outlook; kill-on-exit is off, so Outlook outlives
  it. Meanwhile the suite repeats the tests the two crashes struck in, the Q149 typed-text tests and
  the reply with a signature beside the forward that failed after it, each in a dotnet test of its
  own, for up to 20 rounds or 60 minutes or until a dump. The dumps land beside the WER evidence,
  and the summary says how many rounds ran. Off by default, so no ordinary run has a debugger
  attached, and refused on the Exchange VM, with a filter or with a signature step; the runner's and
  the debugger's self-tests pin the switch, the module and every decision (runbook section 4.15).
  `-HuntVariant` narrows the trigger down. Legs, the default, is the run above, and LegsOtherFolder
  runs the same tests after a step that moves the main window to the Calendar. Every other variant
  runs `T2/LiveHuntDerivedDraftTests` instead, which makes derived drafts in a loop for
  `-HuntLegMinutes` (10 by default) or until Outlook ends, and writes one line per draft, so a run
  counts the drafts before a crash. ProductLoop, the control, makes the product's replies and forwards
  in the windowed Outlook. OtherFolder does the same with the main window on the Calendar, and
  NoWindow in an Outlook that COM started with no window. NoFormPage follows the product's direct
  route through the object model without the "Message" form page, and BareCreators calls Reply() or
  Forward() and releases the result unsaved. Only the hunt's suite script names a variant, so the loop
  makes nothing in any other run. With `-HuntShowEach` every draft of a loop comes after a compose
  window shown and closed: a reply or forward made through the object model, promoted the product's
  own way and closed discarding, which is what the typed-text tests do before each of the product's
  drafts, and with `-HuntTypeEach` that window is displayed on screen and typed into as those tests
  do. PersonTyping keeps a compose window open on screen and types a line into it before each of
  the product's drafts, five drafts a window, as a person writes while an agent drafts. With
  `-HuntAblate` the typed-text tests leave out one of their steps: the new mail beside the reply and
  the forward, reading Word's layout, the test signature, or a source mail of their own. With
  `-HuntTraceClr` the .NET runtime's own exception events are traced on the guest with Windows'
  logman and tracerpt, and their XML is fetched. With `-HuntPromoteAfter` each of the product's
  drafts is re-opened, promoted the product's own way and closed as soon as it is made, which is how
  the typed-text tests read every draft of the product's. With `-HuntReleaseAfterClose` the typed
  window keeps its Word document until the window is closed and lets go of it after, as the
  typed-text tests' window and the product's own compose do; the loop otherwise let go of it first.
  Three of the ablations still crashed the same way, so the crash needs neither the new mails, nor
  the layout read, nor the test signature. PersonTyping made 2,431 of the product's replies and
  forwards beside a compose window a line was typed into before each, and Outlook did not crash.
  Every hunting run also records Office's build and
  the hunted module's file, and the summary shows both. `-HuntWithoutAddIn` turns OutlookAI's own add-in off
  for the run, LoadBehavior 0 before Outlook starts, and the restore after the run brings the
  checkpoint's value back. Every Outlook guest's checkpoint holds the add-in installed and loading,
  and inside Outlook it hooks every inspector and puts its task pane on each one that is activated.
  Each round says whether Outlook holds the .NET runtime the add-in needs. A hunt with the add-in off
  crashed the same way in round 18, so the crash does not need it. The hunt's minutes and the
  debugger's time limit are counted on a stopwatch: two unindexed guests' frozen clocks jumped 46
  days to the real date in the middle of a hunt, and a deadline on the guest's clock ended both the
  rounds and the debugger's watch there.

- ✅ **A live run keeps Windows' own report of every Outlook crash beside its dumps.**
  Two Outlook crashes on 2026-11-18, both in `mso40uiwin32client.dll` while a draft was written,
  left an Application Error event and no dump, though the run had armed Windows Error Reporting's
  local dumps. On a test machine the arming checked out: the keys hold what
  `Set-OutlookCrashDumps.ps1` writes, the folder takes a dump from Outlook's own account, and
  programs crashed by hand or reporting through `ReportFault` left a full dump with the same event.
  A program that files its crash through WER's own calls leaves no local dump, but no Application
  Error event either, and WER keeps no dump with any report on these machines. So both crashes took
  Windows' own crash path and lost the dump somewhere in it, and what WER recorded about them went
  with the restore. A run now also keeps each crash's event 1001, which names WER's report and where
  it filed it, in `guest\crash-events.log`, and copies that report's files from the machine's and
  each user's WER queue and archive into `wer\<guest>\`; the summary counts both. Whether the next
  such crash leaves a dump is not settled by this (runbook section 4.15).

- ✅ **A live test drives the task pane on a test machine, with a stand-in for the Claude Code CLI.**
  Decided by the maintainer on 2026-11-18 (Q168 (b)): nothing tested the pane, so the code-sharing
  refactor's stage S4 could not show that it still worked. T2 `LivePaneTests` has the MCP server open
  a tagged draft with a test signature in a compose window, types an instruction into the pane and
  clicks "Edit current draft" with the Win32 messages typing and clicking send, and reads the open
  window back, read only. The pane's requests go to `ClaudeCliStandIn`, which the test compiles with
  the .NET Framework's own C# compiler to where the add-in runs Claude Code, answers the body and a
  signature name the test chose, and keeps every prompt for the test to read. It goes only onto a
  test machine - the live opt-in and a profile that is not Production come first - never over a real
  CLI, and is removed when the test ends, so the product needs no automation entry point. The test
  checks that the pane still writes one paragraph a line, that its AI's signature choice (Q162 (b))
  follows the body and replaces the agent's signature, that the chosen signature's picture is
  embedded (D47), and that "Select the best signature" then puts in the signature the stand-in
  names next. `LiveOutlookTestMailer` gained the read of an open draft and the close that
  discards it, both refusing a draft without this run's tag. A test machine's checkpoint holds the
  add-in of the day it was made, so `Testbed/host/Invoke-LiveTierOnGuest.ps1 -WithAddIn` builds the
  tested commit's add-in with `Publish-AddInPayload.ps1`, installs it on the frozen guest before
  Outlook starts - elevated, then one start NOT elevated to `ADDIN-READY` - and marks the run for the
  test; a run without it gets a PROVED NOTHING line from the test, as a machine without a population
  does. On `OutlookAI-Unindexed` on 2026-11-18, with the add-in of `8cf5431`, the pane wrote the
  stand-in's four lines as 4 paragraphs, its AI replaced the agent's signature with the one the
  stand-in named, and that signature's picture was embedded, 0 of 1 left a link (runbook section
  2.3). Beside it, the headless tests that count Outlook's windows found two more at their end than
  at their start in two of eleven runs, both on an indexed guest. Named, they were a tooltip and its
  shadow, there from the moment the pane test closed its compose window. A tooltip comes up at the
  pointer, which nobody moves on a test machine, so the pane test now moves the pointer to the
  screen's bottom-left corner before its compose window comes up and after it closes, and lists
  what it leaves on screen with the headless tests' own window listing.

- ✅ **A live run can set a signature policy before Outlook starts, and measure what Outlook inserts.**
  Outlook reads an administrator's signature policy, a default deployed in its general mail settings
  and an account's own defaults when it starts, and the live tests run without administrator rights,
  where a policy cannot be written. So the live runner takes `-SignatureSettings GroupPolicy` or
  `-SignatureSettings Preference` (Q200 (a) and Q201, decided by the maintainer 2026-11-18): after
  the clock guard and before Outlook starts, a new guest script, `Set-SignatureSettings.ps1`, run as
  administrator, writes Group Policy's NewSignature and ReplySignature, or the same two in Outlook's
  own MailSettings, makes the test signatures they name and gives the hub account its own defaults
  beside them, through the shipped manage_signature. The run is narrowed to one live test, which
  opens a new mail, a reply and a forward the way Outlook's own commands do, reads which signature
  Outlook put in and whether its Signature button is open, and holds OutlookAI to the same. Those
  compose windows are the kind that crashed Outlook on 2026-10-04, and the runner keeps full crash
  dumps. The checkpoint the guest is restored to afterwards removes the setting, so the step runs
  only from a frozen checkpoint, never on the Exchange test machine and never with a green
  checkpoint. The values are written as expandable strings unless `-SignatureValueKind String` asks
  for plain ones. The guest scripts that drive the MCP server share one stdio client now.
  A third scenario, `-SignatureSettings PreferenceNoOwnNew`, gives the hub account its own default
  for replies under the deployed setting, so a new mail shows what that setting does for an account
  with no new-mail default of its own, and a fourth, `NoSetting`, gives it both own defaults and makes
  no setting at all, as the control. With `-SignatureDialog` a run reads Outlook's own Signatures
  dialog before the suite instead, through `Measure-SignatureDialog.ps1`, which presses nothing but
  Cancel: whether its default boxes, its signature list and its buttons are open, which a second
  live test holds OutlookAI's refusals to. It runs apart from the insert test because opening the
  dialog under a setting rewrites the accounts' own defaults. The two probes that read Outlook's
  dialogs share one copy of that code, `OutlookDialogUia.ps1`.

- ✅ **Tests pin the start policy, lock, record, log and report, and live tests time the settle wait.**
  T1 runs every branch of the start policy over every combination of its facts, the record's and the
  log's formats, the rule that no free text reaches the log, the add-in stamp rule, the start lock
  between threads and after a holder dies, every variant of the report, the agent's error and the
  start block of `outlook_health`, the in-flight rule for a cancelled request and the rebuild that
  only attaches. A stdio test proves a cancelled request's wedged COM host is still ended at its
  deadline. On the Outlook test machines a live test connects to the running Outlook and checks the
  record, the log and health, and the disconnect test's two cold starts print what the settle wait
  cost and must have settled on every condition. On the Exchange test machine, whose Outlook syncs
  real mailboxes as it starts, a cold start only has to have waited until Outlook was 20 seconds old.

- 📝 **The Outlook hang of 2026-11-15 is written up: its timeline, its cause and what was measured.**
  `Docs/research/outlook-start-hang-2026-11-15.md` keeps what the investigation item in `TODO.md`
  held, now that the item is gone: what happened minute by minute, the deadlock between OutlookAI's
  connection and the Teams Meeting add-in as far as it was established, what was ruled out, which of
  Q175 to Q182 answers which part of it, and what the settle wait costs on the test machines. It
  names no person, mailbox or machine; the full report stays outside the repository.

- 📝 **The README says what to do when Outlook does not open or an agent says it is not responding.**
  A troubleshooting entry explains the Outlook without a window that hangs, gives the Task Manager
  steps and what ending Outlook can cost, and asks for the start log in a report. The section on
  Outlook's lifetime says when a launch is handed to a hung Outlook and that the server waits about
  20 seconds after any Outlook start; the server's README and `McpServer/Docs/com-host.md` describe
  the settle wait, the start lock, the limit, the record, the log and the report.

- ✅ **Test runs wait while the PC is 70 % busy or more, and a pause switch holds new ones back.**
  With every testbed machine running at once, the host's 32 logical processors were 82 to 100 % busy:
  the maintainer's own work slowed down, and every run took about twice as long. He decided on
  2026-11-18 (Q195 (e)) that a run starts only while the processors were less than 70 % busy over the
  last 30 seconds, and that he can stop new runs when he needs the PC, for gaming for example.
  `pwsh -File Testbed/host/Set-TestbedPause.ps1 -On` holds every new run back until `-Off`, and runs
  already going finish. The build runner looks at the switch and the load on every turn of its queue,
  before it touches a build machine; the live runner before its build and on every turn of its wait
  for a guest; the clone maker before it takes its source. A waiting run holds no machine's lock or
  lease, keeps its place in the queue, sleeps until a file of the testbed changes or 30 seconds pass,
  says what it waits for and since when once a minute, and is listed by `Set-TestbedPause.ps1`.
  A machine another run took in the last two minutes counts at its full size, so the runs that waited
  do not all start in the second the load dips. The load comes from kernel32's GetSystemTimes, which
  counts the virtual machines here because Hyper-V runs its root scheduler on this host: it read 83 %
  where the hypervisor's own counter read 85 %, of which 29 points were the machines'.

- ✅ **The compose measurement also runs the product's own drafting, Outlook's typed text, and cold starts.**
  As the maintainer decided on 2026-11-18, the Q149 suite stays (Q210 (a)) and gains seven cells.
  P1 runs the product's drafts in an Outlook with no window, three starts of ten drafts each,
  recording each draft's route, every window shown and every change of foreground, then ends the
  session and watches Outlook end. P2 has two sessions draft in one such Outlook and checks that the
  first to leave keeps the hidden window and the last closes it; P3 drafts in an Outlook a program
  started in which a window was then shown, where nothing may be activated. In these a small window
  of the test's own holds the foreground before each draft, standing in for the program a person
  types in, and each draft records how long the foreground was elsewhere and where it went. P4 and
  P5 repeat P1 and P3 while that window types a character no draft holds, 25 times a second, and
  count the keys that did not reach it and those found in the stored drafts. Q211 opens drafts in
  Outlook's own compose window, types into it through Word, and reads the style and spacing of the
  typed paragraphs beside the product's drafts of the same text (Q211 (b)). Every cold start in these
  runs is logged for Q180 (b): the time until Outlook answered, Windows' verdict on it every 100 ms,
  and any hang. They run like the rest of the suite, alone. Run on 2026-11-18, they found typed text
  in Word's "Normal" style with no space before or after, as in the product's drafts without a
  signature, while with a signature the product's text takes the signature's "Normal (Web)", 5 pt
  before and after; 22 cold starts answered in 2.3 to 4.1 s and none hung
  (`Docs/research/q149-draft-text`, stage 4).

- ✅ **A tested one-off took the confirmed test leftover out of the bystander mailbox's Sent Items.**
  The maintainer confirmed it by its EntryID on 2026-11-18 (Q138): an item the live tests left in that
  mailbox's Sent Items in July, when they still ran on his own computer. It is the only write that
  mailbox ever gets. A new live test, which no machine's filter selects and which runs only when named,
  finds the item by its EntryID and its exact subject, both required, moves it to the mailbox's Deleted
  Items, checks the copy the move hands back, deletes it there, and reads both folders again. The
  EntryID is pinned by its SHA-256, so the repository holds no identifier of the real mailbox. The count
  tripwire, which censuses that mailbox item by item around the run, accepts exactly that departure: a
  tagged item, by EntryID, from a folder it identified item by item, once. Any other change there still
  fails the run. Two runs on 2026-11-18 found two things. A folder table on a cached Exchange store
  gives a short-term id, so the first run matched nothing and changed nothing; the item's own long-term
  id is read now. The second run moved the item to Deleted Items, but the delete there did not purge it,
  so the tripwire failed that run on the one departure, as it should. The item waits in the bystander's
  Deleted Items, under a new EntryID, for the maintainer's decision.

- ✅ **A live run records the guest's clock and time zone around every call it makes into the guest.**
  In two measurement runs on a frozen-clock test machine, the machine's clock jumped about two hours
  forward during the suite, and waits of 9,000 and 7,200 seconds ran out after 35 to 40 minutes of real
  time. Nobody knows why yet. Each call into the guest now reads the guest's UTC time and time zone as
  its first and last step, and the runner sets the two readings beside its own clock in
  `guest-clock.log`. A call over which the two clocks ran more than 60 seconds apart, or the time zone
  changed, is a warning in `run.log`, `summary.txt` and `summary.json`; the run's verdict does not
  change.

- ✅ **A probe on the Exchange test machine measures what the folder-ID fixes of November rest on.**
  `Testbed/guest/Measure-FolderIdResolution.ps1` reads, changing nothing, what three of the fixes
  needed. Reading the top folder of the mailbox Exchange automapped online takes 2 to 3 ms, as it does
  for the cached stores. That entry and the test account's own cached store carry one StoreID and one
  provider UID. A folder ID opens in the cached entry whichever StoreID is passed beside it, and an ID
  no folder has opens there too, as an empty folder whose every property is missing. And the test
  account's newest index rows are holiday appointments, whose index time is their creation time and
  whose subjects come back up to five times, created tenths of a second apart. It prints no name,
  address, ID or subject: those are compared through short hashes (Docs/live-tier-on-the-vm.md
  section 4.9).

- ✅ **A test-machine probe measures how Windows and the installer remove trust in the leaked key.**
  `Testbed/host/Invoke-TrustRemovalProbe.ps1` runs `Testbed/guest/Measure-TrustRemoval.ps1` on an
  Outlook test machine under its lease, from its frozen checkpoint, and puts the checkpoint back
  afterwards. One scenario adds a throwaway certificate to the user's Root store and removes it with
  certutil started exactly as the installer starts it, and records every question Windows asks: its
  text, whether it is shown, its default answer, what Yes and No do, and whether Windows still trusts
  the certificate afterwards. It found that every removal asks, that certutil waits for the answer
  (still waiting after two minutes), and that deleting the certificate's registry key instead leaves
  Windows' list of confirmed roots naming it, so writing the key back makes it trusted again with no
  question. The other scenario runs the built installer and its uninstaller against planted trust
  entries and checks that exactly the leaked key's go. Both pass (Docs/live-tier-on-the-vm.md
  section 4.8).

- ✅ **The tests address folders and stores by ID, as an assistant does (Q131-Q136).**
  The mailbox tests on the test machines look IDs up through `list_accounts`, `list_folders` and
  `find_folder`, and make their test folders with `create_folder`, which refuses any name without
  the test-folder prefix. New checks pin the ID format: every single mistyped character and every
  swap of two characters up to 30 places apart is refused, a name or path is refused with the
  `find_folder` call that would find it, and an ID round-trips to the folder's whole entry id. Others
  pin how a search finds its folder in the index with no Outlook call, the refusal of an Exchange
  folder while Outlook is not answering, `find_folder`'s ordering and cap, every `create_folder`
  outcome, and the refusal of every argument this release removed.

- 📝 **The probe the D49 design rests on, and the raw Q118 pin logs, are kept in Docs/research.**
  The 2026-07-31 probe behind the invisible compose surface and the lifetime pin, recovered from its
  session transcript, and the raw logs of the Q118 pin measurement of 2026-10-03 were each in one
  scratch folder until now. Identifiers are replaced, and the probe's scripts and sources are kept as
  text so nothing runs them. Read against the probe's own results, the D49 record overstates in four
  ways, now written down beside them: it gives no sample sizes (22 runs in all), three of its runs
  attached to an Outlook an earlier run had started instead of a fresh one, rows 3, 7, 8 and 9
  overstate survival, and its 53-79 ms to an editor leaves out the 314-611 ms Activate before it. A
  TODO note saying an Outlook that outlived its holder without a pin weakened the pin's case is
  corrected: a holder's death releases references, and the pin guards against a window closing.

- ✅ **The Q149 measurement of how drafts are written runs as live tests on the test machines.**
  Three T2 classes in a collection of their own measure, on Office LTSC 2024, whether closing a
  draft's hidden compose window ends an Outlook that has no window when no lifetime pin holds it,
  whether Word's editor can be had, or an edit committed, without activating that window, and
  whether a draft written as HTML alone matches today's when opened in Outlook and when received
  through the test machine's mail sink. Every attempt is one JSON line carrying its HRESULT. The
  suite ends and restarts Outlook dozens of times, so it runs only in a run that selects it alone
  and reports NOT RUN in any other. Its runs of 2026-11-17, on both kinds of test guest, found that
  without the pin the compose window's close ended a window-less Outlook in 40 of 40 drafts. An
  inspector never activated refused its editor in every trial (86 of 86, waits of 30 seconds and of
  180 included), and no Windows-level show of a window changed that, but after one activation in a
  process every later inspector handed it out (120 of 120). An edit made that way reached the draft
  when committed with ModifiedFormPages and Close(olSave) (60 of 60 in a window-less Outlook, 40 of
  40 in a windowed one), and not with CurrentItem.Save (0 of 100). No draft written as HTML alone
  looked the same as today's, opened in Outlook or as received. The merged plan, its decision rules,
  every attempt line and the results are in Docs/research/q149-draft-text.

- ✅ **Tests that can read the machine's time zone run again in Central European time (Q144).**
  The build machine runs the suite in UTC (Q126), where the machine's zone and UTC are one zone, so
  a test that compares a local time with a UTC one only proved them equal. After the suite, every
  run now switches the machine to W. Europe Standard Time, runs the tests marked
  `[Trait("TimeZone", "Machine")]` once more, and switches it back before the self-checks. The
  summary, its JSON and the exit code report both runs. A new check, `T1/TimeZoneMarkerTests`,
  fails any unmarked test whose code, or the code it calls in the product, reads the machine's zone,
  and names the call chain to the read. It found 323 of the 2,676 test methods, in 39 classes; 145
  of them reach the zone only through a product conversion that reads it for a local time alone,
  which a reading of the code cannot tell apart. All 323 are marked; their 383 test cases took
  1.2 s between them in a recent full run. As the maintainer decided, a marked test is rewritten to
  set its own zone when it is next changed, and then loses its mark.

- ✅ **The signature tests try the profile, folder and default settings on a test machine.**
  Three live tests set a value the user could set themselves, check what OutlookAI does with it,
  record what Outlook's own signature editor reports, and put the value back: Outlook running on a
  profile other than the one recorded as its default, a default deployed for every account in
  Outlook's settings, and a signature folder moved elsewhere. A snapshot of the registry, beside the
  snapshot of the signature files, now fails any signature test run that leaves a signature setting
  changed: the defaults in every profile, the folder, the policies, the roaming switches and the
  default profile. Their first run, on the unindexed test machine on 2026-11-17: with Outlook on
  `OutlookAI-Tier` and `CorpusProfile` recorded as the default, a draft got the running profile's
  default; a folder moved while Outlook ran was listed, written to and used by OutlookAI, and
  Outlook's own list of signatures showed what was in it. What Outlook makes of a default deployed
  for every account could not be read that way, written while Outlook ran: its signature settings
  reported no default with it or without it. The live runner's -SignatureSettings step measured it
  on 2026-11-18, with the setting made before Outlook starts.

- ✅ **The Exchange test machine checks folders by ID without writing: 66-character IDs and find_folder.**
  Folders and stores addressed by ID (Q131-Q136) reached a real Exchange mailbox through T1 and one
  measurement only. Three live tests that write nothing now run on that machine: every folder of the
  test account carries a 66-character folder ID of kind X that reads back to an Exchange entry id,
  find_folder finds the Inbox by its name under the ID list_folders gives, and a search scoped to that
  ID answers while an Exchange-shaped ID no folder has is refused. Their first run, on 2026-11-18,
  confirmed the IDs and found two things for the folders-by-ID code to settle: the test account's
  storeId names both its cached store and the online entry Exchange automapped, so list_folders and
  find_folder show its folders twice, and an Exchange folder ID no folder has was answered, not
  refused. Both are settled: the mailbox is listed and resolved once, through its cached entry, and
  the ID no folder has is refused. A fourth test checks that every store of the profile has a
  28-character store ID, the automapped entry's too. create_folder's refusals and the refusal while Outlook does not answer stay with T1:
  reaching them there would take a write or a stopped Outlook. The test-side MCP client on that machine passes find_folder and lets create_folder
  make a folder only in the test account, and the guarded session follows the new folder interface.

- ✅ **The Exchange test machine writes in a dedicated test account, and only there.**
  A Microsoft 365 account the maintainer gave for extensive testing joined the machine's Outlook
  profile beside the mailbox it already held, and is the hub its live tests may write in (Q137,
  settled 2026-10-11); that mailbox, real mail, is a read-only bystander the count tripwire watches.
  The write-safety design proposed on 2026-10-03 is built for that account:
  every address a test hands to Outlook, in-process or through the MCP server, must be the test
  account's own; every item a run makes carries the test tag and a run marker the host records before
  the run starts; a reply, a send, a move or a delete touches only the run's own tagged items in the
  test account; and the sweep takes back only items of a recorded run, by tag and marker, so an aborted
  run's leftovers go with the next run and never by a subject pattern. Exchange also maps the test
  mailbox into the profile a second time, as an online entry reached through the real mailbox's Full
  Access; before every baseline the tripwire now proves by the store's record key that the hub's name
  is the test account's own store and that the second entry is neither censused nor written through,
  or the run refuses. Neither address is in the repository: `Testbed/testbed.json` names each account
  by its role, a run reads the address from that account's gitignored credential file, and the same
  check refuses a run whose settings name an account the machine's Outlook has not signed in. Another
  test machine's settings are refused only for such a placeholder in their own section, never for the
  accounts the Exchange machine records, which had stopped every Outlook guest's live run. The
  live runner runs this machine with the broadest filter it can honestly run, carries
  its state from run to run, because a restored offline data file is older than the server, and sets
  the data files aside when a checkpoint is restored, has Windows Search rebuild its index, and waits
  until each indexed mailbox has one index root before the tests start. A tagged item that was
  already in a read-only
  mailbox before a run is reported and left alone instead of failing every run.

- ✅ **A second Microsoft 365 account signs in on the Exchange test machine, and an MFA page stops it.**
  The sign-in script opens Outlook's own Add Account dialog in the running Outlook and types the
  address and the password through the machine's virtual keyboard, as for the first account; each
  account keeps its own encrypted credential file, picked by address, and a verification-code secret
  is optional. After the password the script reads the page Microsoft shows before it types anything:
  a request for a second factor, or for one to be registered, stops it with a message for the
  maintainer, and nothing is clicked on his behalf. The test account signed in on 2026-10-08 without a
  second factor; its code secret has been stored since, so a code page would be answered. Since
  2026-11-17 the same script also signs the test account, and only it, in to Outlook on the web in an
  InPrivate Edge window on that machine, for the roaming-signature measurement the maintainer allowed
  (Q161 (c)), answers "Stay signed in?" with No, and signs it out again.

- ✅ **Three probes measure signatures, the offline window and index ids on the Exchange test machine.**
  `Measure-SignatureRoaming.ps1` reads the machine's Office build, licence and signature settings,
  opens Outlook's Signatures dialog and closes it with Cancel, and looks for the hidden folder roaming
  signatures live in: on 2026-11-17 Office LTSC 2024 16.0.17932.21000 showed the classic dialog with no
  sign of roaming, and a signature made the test account's default in Outlook on the web had no copy
  in Outlook an hour later or after a restart (Q147 R3). That was with the testbed's policies that
  switch Office's connected experiences off, and Office past its activation grace. On 2026-11-18, with
  those four policy values removed and the grace re-armed (Q194 (b)), the same kind of signature
  arrived about three minutes after it was saved: a `<name> (<address>)` copy in the signature folder,
  set as the test account's default for new mail, still there after an hour and a restart. With the
  policies back and the grace kept, it did not come, and Outlook's Options no longer offered "Store my
  Outlook settings in the cloud" at all. So this Outlook roams signatures unless connected experiences
  are switched off, and Outlook's own copies show it, as OutlookAI's evidence rule expects. The probe also reads now whether "Store my Outlook settings in
  the cloud" is on: it was. The hidden folder could not be seen on the server even while either
  signature existed, so the probe's server-side reading proves nothing.
  `Measure-OfflineWindow.ps1` moves the test account's "Download email for the past" slider through its
  eleven positions and reads the profile after each: the months are property 0x6649 and the days
  0x665A, always written as a pair, and LTSC 2024 offers 3 and 5 years; a pair written into the profile
  while Outlook runs, as the add-in would, is kept when Outlook quits and shown at its next start.
  `Measure-ExchangeIndexIds.ps1`
  found that a cached Exchange folder's id does not lead to its Windows Search row the way a PST
  folder's does: the row carries the offline file's own node id, which the id lacks and which changes
  when the file is rebuilt. Each probe has a self-test, and the slider probe changes nothing but the
  test account's own setting, on that machine, before its checkpoint is restored.

- ✅ **The OutlookAI add-in is installed on the Exchange test machine, and its tests run there.**
  Installed on 2026-11-17 at the maintainer's decision (Q140 (b)) through the testbed's own two-phase
  installer, it reached ADDIN-READY with 8 of its 13 Outlook settings applied, the five Cached Mode
  policy values left to an administrator as everywhere else. The machine's live-run filter stops
  leaving out the tests that read the add-in's registry state. Checkpoint `CP-08-ADDIN-READY`.

- ✅ **Three build machines and two of each Outlook test machine now take test runs at once.**
  Eight agents were queueing on one build machine and the two Outlook test machines, and the
  maintainer asked for as much parallel work as the host allows, extra machines included. The build
  machine has two copies, `OutlookAI-Build2` and `OutlookAI-Build3`, and each Outlook test machine
  one, `OutlookAI-Indexed2` and `OutlookAI-Unindexed2`, each made in about two minutes from the
  checkpoint its source's runner restores, with the same computer name, account, mailboxes and
  frozen clock (`Testbed/host/New-TestbedClone.ps1`). The build runner takes whichever build machine
  is free and names it in its summary; the live runner takes a machine by name as before, or the
  first free one of a kind with `-Kind Indexed` or `-Kind Unindexed`. A copy is started only while
  the host has its memory and 16 GB more available, so the maintainer's own work keeps its memory;
  a resting copy holds disk and no memory. The idle-saver watches every copy. The Exchange test
  machine is never copied: it holds a real Microsoft 365 sign-in.

- ✅ **A fifth guard fails a release or a build-VM run that finds personal data in the repository.**
  An audit of 2026-11-17 found a third party's shared mailbox, a delegate's name, a colleague's first
  name and the maintainer's own user name, computer name and mailbox details in this public
  repository, most of them written there by agents. `Tools/Checks/check-personal-data.ps1` looks
  for them in every tracked file: the known values from a private list kept only in the maintainer's
  checkout - nothing about them is committed, not even a hash, which a dictionary would turn back
  into a short word (Q196) - and, by shape, an e-mail address outside the reserved domains, a profile
  path with a real-looking user name, a private key, a key or certificate file, a Windows default
  computer name, and the names of the machine it runs on. A finding names the file, the line and the
  rule, never the value. The release runs it under both PowerShells and refuses to pass without the
  private list. Its self-test runs on the build VM, which gets a copy of the list for each run: it
  goes straight from the checkout to the VM, is never fetched back, and is reverted with the VM, and
  a run that finds no list says so on its summary's first line. A baseline in
  `Tools/Checks/personal-data.json` lets HEAD keep, per file, the occurrences it still has and not
  one more. `.gitignore` now also keeps out `*.p12`, `*.snk`, `*.pem` and `*.key` beside `*.pfx`,
  and the guard fails if those rules go (Q190).

- ✅ **Git hooks refuse a commit, a commit message or a push that would publish personal data.**
  `Tools/Install-GitHooks.ps1 -Execute` puts pre-commit, commit-msg and pre-push hooks into the main
  repository's `.git/hooks`, which every worktree shares, with a copy of the guard for them to run,
  so a worktree on an older branch is checked as well. They read the private list of real values in
  the maintainer's checkout, from any worktree. A commit fails only when it adds an
  occurrence to a file, so an edit next to a value HEAD still holds goes through. Tried in a scratch
  repository on 2026-11-17: a value in a file, in a commit message and in a pushed commit was
  refused, from the main checkout and from a linked worktree, and a clean commit and push went
  through (Q190).

- 📝 **AGENTS.md says to write placeholders, never personal data, in everything committed.**
  That covers every tracked file, commit message, TODO item and release note: `<user>`,
  `<computer>`, an address under a reserved domain, or "the Exchange VM's mailbox". The real values
  are kept only in the maintainer's checkout, in `.work/personal-data/identifiers.json`. AGENTS.md
  also holds the one-off exception for `git filter-repo`, the tool of the coming history rewrite
  (Q187), and QUESTIONS.md records the maintainer's answers of 2026-11-17 on personal data in the
  public history, Q183 to Q193.

- ✅ **The Outlook test machines keep a full crash dump whenever Outlook crashes in a live run.**
  The one-call runner switches Windows Error Reporting's own crash dumps on at every run, on the
  frozen checkpoints as they are, brings every dump back with the run's results, and reads each one
  with Windows' own debugging library: its summary names the fault, the code it struck in, the
  faulting thread's call stack and the classes of the objects that thread was handling. Nothing is
  installed on the machines and no checkpoint changed.

- 📦 **A release runs CodeQL's C# security queries and a check for vulnerable NuGet packages (Q125).**
  They take the place of the CodeQL scan and the dependency review that ran on GitHub until
  2026-10-03, as step 7 of `Tools/Publish-Release.ps1`, and each also runs on its own.
  `Tools/Invoke-CodeQL.ps1` fetches CodeQL 2.27.1, held to the SHA-256 GitHub publishes for it, and
  runs every C# security query over the commit; a result fails the release unless it is triaged,
  with its reason, in `Tools/codeql-accepted.json`. CodeQL is free but not open source, so
  `AGENTS.md` records it as a third exception to the Dependencies rule, with what its licence
  allows: this repository, while it is open source and on GitHub.com.
  `Tools/Test-VulnerablePackages.ps1` asks nuget.org about every package of the server's projects,
  direct or brought in by another, and a run that could not ask fails too. The first runs found
  nothing: no CodeQL result in 432 C# files, and no known-vulnerable package in six projects.
  CodeQL had little chance to, though: by default it trusts everything but the network, and
  OutlookAI reads nothing from the network. So since Q146 (decided 2026-10-11) it also distrusts
  standard input, files, the command line and the registry. Environment variables stay trusted:
  over 0ccf5df they alone made 168 of the 174 results, nearly all paths under the user's own
  profile and TEMP. And `Tools/CodeQL` tells it that every argument an agent passes to an MCP tool
  is untrusted, because the agent fills them in from mail other people wrote. CodeQL cannot see
  that for itself: the MCP SDK hands the arguments over inside a library it does not analyse. A
  census query on every run counts the tools that model reached, and the run fails if one is
  missing, so a model that stopped matching cannot pass a release unnoticed: 22 of 22 on
  2026-11-17. That run found six results. Five are safe, and `Tools/codeql-accepted.json` says
  why: the search index's SQL, which six arguments of the search and thread tools reach only
  through `WsSqlBuilder`'s allow-list and doubled quotes, and four reads of registry values that
  only the user or an administrator can set. The sixth was real: the elevated helper's install
  path, which Q145 removes along with the helper. What the scan still cannot see, measured: the
  work the COM host process does with the arguments, and the eight list arguments past the
  server's first call (TODO.md).

- ✅ **Every build runs the .NET SDK's security analysers, and any finding fails the build (Q125).**
  GitHub's CI went on 2026-10-03, and the CodeQL scan with it; the maintainer chose the analysers
  that ship with the .NET SDK in its place. Every rule in their Security category that .NET 10
  ships runs as an error: in the MCP server's projects through `McpServer/Directory.Build.props`,
  and in the add-in and its helper through `NetSecurityAnalyzers.targets`, which imports the SDK's
  own analyser files into the two old-style projects. Switched on, they found 78 places in the
  server and one in the add-in. The DLL search (CA5392) is fixed, under Security. The rest were
  judged safe and are suppressed where they stand, each with its reason: the search index's SQL
  (CA2100), MD5 and a fixed-seed random in two tests (CA5351, CA5394), and the updater's TLS 1.2
  (CA5386), where the analysers' own fix was measured to stop update checks reaching GitHub.

- ✅ **A one-call runner runs the live tests on the two test machines.**
  It builds the chosen commit, waits its turn for the machine, prepares it, runs the tests the safe
  way, reads the suite's own safety checks as part of the verdict, and always puts the machine back
  to rest and saves it - so a test run no longer depends on someone repeating dozens of manual steps
  correctly. Nothing that ships to users is affected.

- ✅ **Which date the search index gives appointments, contacts and tasks in a data file is measured (Q98).**
  Appointments and tasks get a received date - their creation time - so they sort and filter like
  mail, and contacts get none. The test machines' measurement-mailbox tool can now hold each probe
  item until the index has it and record every column a search sorts by, and the script that adds a
  data file to a test machine can create an older-format (ANSI) one. Nothing that ships to users is
  affected.

- ✅ **A third test machine runs the non-mailbox tests and every self-check, off the maintainer's computer.**
  It is a small virtual machine with no Office, no mailbox and no network, driven by one script. The
  script tests a named commit or branch - an agent's unfinished work included, once committed - on a
  machine put back to the same saved state before and after every run, so no run can inherit
  anything from the one before; callers queue and run one at a time, each gets back a summary, the
  full test results and every log, and a caller stopped part-way leaves nothing running: a watcher
  it starts puts the machine back and saves it. Its runs reproduced the main branch's passing tests
  exactly - 3,005, and 3,346 at a later commit - in about four minutes each. Building it also found
  two faults in the step that finishes a test machine's Windows install - it checked its own work a
  moment too early, saw the install discs still attached and the first checkpoint not yet there, and
  stopped - and one in the step that creates a machine, which recorded a fixed memory size for
  machines that in fact had memory growing on demand up to a terabyte; all three are fixed. Nothing
  that ships to users is affected.

- ✅ **Every mailbox check that changes nothing is labelled, and the build proves it (Q74).**
  Fifty-four of the checks that run against a real Outlook now carry a "writes nothing" label; the
  build reads their compiled code - including everything that runs before and after them - and fails
  if any of them could create, change, move, send or delete anything, or reach a part of the product
  nobody has shown to be read-only. The maintainer's own machine runs only the seven labelled checks
  that need an Exchange mailbox, selected by a filter computed from those labels rather than typed
  by hand, and the instructions for that run, the test-machine run book and the project rules now
  describe it. A short list of manual checks before a release covers the Exchange-only writes no
  automated check can reach. Nothing that ships to users is affected.

- ✅ **On-behalf sends and shared or delegate Exchange drafts get checks that need no mailbox (Q74).**
  The drafts include replies. No mailbox check that runs against a real Outlook can reach these any
  more - the test machines have no Exchange, and the maintainer's own machine is read-only - so the
  decisions behind them (which account a mail goes out from, which on-behalf name it carries, when a
  send is refused) now have checks that need no mailbox at all. How the product behaves is
  unchanged.

- ✅ **The test-mailbox tool lists a test machine's mailbox folders, read-only.**
  So what a build or a clean-up left behind - folders Outlook was asked to create, emptied folders
  moved to Deleted Items - can be checked with the project's own guarded tool instead of an
  improvised script.

- ✅ **The unindexed test machine is built out towards its full design, one checkpointed step at a time.**
  Each step is scripted and proven on the machine. The test machines now answer Outlook's "a program
  is trying to access email address information" security prompt by themselves: they have no
  network, so their antivirus is permanently out of date, and Outlook then stops every outside
  program that reads an address, a message body or a recipient on an Allow/Deny prompt that nobody
  is there to click - which would hang the mailbox tests. A documented Outlook security policy now
  approves those reads on the test machines only, and its check does what the tests do - it reads an
  account's address - and fails if the prompt appears; with the policy absent the same check caught
  the prompt within a second, so the check is known to work. The three Outlook settings that keep
  the classic, scriptable Outlook in place, lost on that machine by an earlier hand-run script, are
  restored and stay put across Outlook restarts. And the rebuild instructions no longer tell anyone
  to restart a test machine with a plain `shutdown /r`, which force-closes Outlook exactly as
  killing it would; they give the graceful route that was used instead. The machine's local mail
  server was installed and proved itself on the first attempt, including starting on its own after a
  restart - and settling the question of whether Outlook logs in to it without a stored password
  showed that it does not: Outlook asks for one before it ever connects, so mail would never have
  been collected. The test machines' account setup now stores a password (any value; the local
  server compares it with nothing), and the server's own log shows Outlook logging in with it. The
  OutlookAI add-in was installed on the machine from the same commit as the tests, trusted without a
  prompt, and confirmed running inside Outlook, twice; its installer's self-check no longer crashes
  when it is copied to the machine on its own. The machine's second mail account - the one the tests
  use to check that a draft is written from the right account, into that account's own mailbox - was
  added and verified there too, surviving an Outlook restart; it also asked for a password until one
  was stored, so the setup steps now store it for every test account. Setting up that account's
  signature exposed a real defect in the signature tool: asked to make a signature the default for
  an account, it can attach it to a mail file that merely carries the account's address as its name,
  then report success. That is recorded for a fix rather than worked around. The first build of the
  small test mailboxes the tests read from worked for the main test mailbox but exposed four defects
  in the test-data generator and the account setup, all recorded for their owners rather than
  patched around: received test messages are never addressed to a resolved mailbox owner; a mail
  file added to a profile after the fact has no Inbox, so the messages meant for it landed in a
  hidden folder no one can see; its throwaway test items could be left behind in a different mail
  file than the one it was told it may touch; and the second account's incoming mail would be filed
  in that same hidden folder. The machine can now build and run the test suite by itself - installed
  from the staged media with no network, proven by discovering all 2,794 tests and running a sample,
  twice - and its settings file, generated from the recorded store names rather than typed, is in
  place and accepted by the suite's own loader. The first live run on it is deliberately held back
  until a defect in the tests' own safety census, which can create folders in the mail files it is
  meant only to count, is fixed.

- ✅ **The indexed test machine is built out the same way, with its mail searchable throughout.**
  Each step is scripted and checkpointed: its lost Outlook settings restored, Outlook's security
  prompt approved automatically, the local mail server installed with both test accounts logging in
  to it, and the add-in installed from the current code and confirmed running twice - with the
  search index checked unchanged after every step that started Outlook. Outlook on that machine has
  to run without administrator rights for its mail to be searchable, so every check that talks to it
  now runs without them as well. Its second test account now has its own default signature, set by
  the corrected signature tool on the account itself and read back from there rather than from a
  mail file carrying the same name, and the machine's copy of the test suite was brought up to date
  and proven to build and run.

- ✅ **The test machines get a local mail server, so the thirteen tests that send mail run as written.**
  The product's own send tool is among those tests. The test machines have no network, so the server
  runs on each machine itself: it accepts what Outlook sends and hands it straight back when Outlook
  collects its mail. It is Inbucket, a small free open-source program under the MIT licence, used as
  it comes rather than written here, at the maintainer's request and as the one deliberate exception
  to the project's rule against external applications. It was chosen over the two usual alternatives
  for two reasons that matter on an unattended machine: it lets Outlook collect mail without a
  stored password, where the others refuse one and a password prompt nobody can answer would stall
  the machine; and it keeps each test account's mail separate, where the others would let two
  accounts take each other's mail at random. It is never downloaded on a test machine: a new script
  fetches the pinned release on the maintainer's computer and accepts it only when its fingerprint
  matches both the one its makers publish and the one recorded in this project, and the installer on
  the test machine refuses any copy that does not match. The installer starts the server whenever
  the machine starts, and then proves it works by sending test messages through it and reading them
  back - lines that start with a full stop, an attachment, deletion, a restart, and one mailbox
  unable to see another's mail - instead of only checking that something is listening. The test
  suite itself now also refuses to start when the mail server's ports answer in the wrong language,
  which would otherwise show up minutes later as mail that never arrived. None of this has run on a
  test machine yet, and one question can only be answered there: whether Outlook collects mail
  without asking for a password it does not have. The documentation says how to find out without
  sending a single message, and what to change if it does ask.

- ✅ **The test machines can be given the OutlookAI add-in, as a scripted step of building them.**
  The mailbox tests that check what the add-in sets up in Outlook can then run there instead of
  failing. Two of those tests read settings only the add-in writes the first time it runs, and the
  test machines had never had it - nothing in their build installed it, and one of the two runs on
  both machines. A new script on the maintainer's machine builds the add-in from a named, committed
  version of the project - the same version the test suite on the machine is built from, so the
  tests check the add-in they are meant to check rather than whatever was last released - and wraps
  it in the product's own installer. A second script on the test machine installs it silently,
  trusts it without the security prompt that would otherwise freeze an unattended machine, starts
  Outlook once, and then checks the exact settings the tests read, down to their type, and that the
  add-in wrote them during that start and answers when called. **Building the add-in turned out to
  register it on the machine doing the building**, quietly pointing that machine's own Outlook at
  the fresh build; on the maintainer's machine that is the Outlook he works in. The build script
  stops that three separate ways and compares the machine before and after, and three builds left it
  exactly as it was - the first also caught, and removed, a certificate the signing step had left
  behind in a place nobody expected. The add-in writes nothing that affects which mailboxes are
  searchable or what is in them, so the machine kept deliberately out of the search index stays that
  way. The install has not yet run on a test machine.

- ✅ **The test machines' small mailboxes get real, reproducible contents, so 36 more checks can run there.**
  Those 36 could never run on a test machine, because the mailboxes they read were empty, or held no
  attachments, senders or conversations; and 4 that ran but proved little should now prove what they
  claim. Three of the 36 will run but still say nothing about items with no received date, such as
  appointments, which these contents leave out; whether to add some is an open question. The
  project's measurement-mailbox builder can now also build three small, fixed sets of made-up mail:
  56 messages for the mailbox the tests write to, 300 for the mailbox no test may ever touch, and 8
  for the second account's mailbox. Unlike the large measurement mailbox they carry what the checks
  actually read - senders and recipients, attachments of the kinds search used to drop (pictures,
  calendar invites, forwarded messages) as well as plain documents, conversations with several
  replies, read and unread mail, folders inside folders, and a folder of notices whose search word
  appears only in the subject. Each set is rebuilt identically from a recorded number, is tagged
  like the measurement mailbox so no test clean-up can ever select it, and is removed by the same
  two-key teardown. Before building, the builder proves on one throwaway message that every kind of
  write it needs actually lands, and refuses otherwise; after building, it reads every message back
  and fails if anything differs from the plan. The settings file's single list of mailboxes, which
  meant both "watch these for lost mail" and "measure search on these", is split into two lists,
  because the second account's mailbox has to be watched but holds too little to measure search on;
  a settings file written before the split behaves exactly as it did, and search checks on a machine
  with nothing indexed now refuse instead of passing having looked at nothing. The indexed test
  machine's measurement mailbox must now be built at the size of a real mailbox, 160,000 messages,
  before its settings will be produced, so the search speed limits are real tests there rather than
  formalities. The documents now say what the suite actually refuses about the untouchable mailbox,
  answering the question the previous entry left open: the documents were wrong, not the suite. The
  build also fails if the example settings file ever contains an email address that could belong to
  someone, as this repository is public. **None of this has run on a test machine yet**: the new
  contents are built and checked on the host without Outlook, and the first build on each machine is
  what settles the handful of open questions the setup guide lists.

- ✅ **A script produces each test machine's suite settings file, no longer written by hand.**
  That removes the last manual step between the test machines and the suite's first-ever run on one
  of them. That file names every mailbox on the machine and says which one the tests may write to
  and which they must never touch; without it the suite refuses to start, and because it names real
  mailboxes it can never be committed as it stands. So the project now keeps a template of it that
  holds nothing but placeholders, beside a per-machine list of each test machine's own mailbox
  names, and a new script puts the two together for one named machine: it refuses outright while any
  value is still unknown, naming every one; checks the result against the same rules the suite
  applies before it touches a mailbox, plus the documented ones the suite does not check itself;
  writes it to scratch space; and prints the exact command that copies it onto that machine. **It
  will not write anywhere the maintainer's own hand-written settings file lives, however the path is
  spelled** - relative paths, forward slashes, junctions and short names included - and that was
  proved against the real file, which was left byte-for-byte unchanged. The template cannot quietly
  drift either: the build now fails if its fields stop matching the documented example's, and a new
  check feeds it, filled with made-up values, through the suite's own settings loader, alongside
  deliberately broken versions the loader must reject. The test machines' real mailbox names have
  not been filled in yet - they have to be read off each machine - and until they are, the script
  says exactly which ones are missing. One documented rule turned out to be enforced less strictly
  by the suite than the documentation says: a mailbox declared untouchable but left off the list of
  mailboxes to watch is watched anyway, not refused; the new script applies the stricter reading,
  and which of the two should change is an open question.

- ✅ **The mailbox test suite can be run on a test machine, which it never could before.**
  The two scripts written for it were run for the first time: one packed up the project's source and
  every component the build needs into a form a machine with no internet can build from, and proved
  that pack was complete before shipping it; the other installed the development tools on the test
  machine and then demonstrated the result by building and running part of the suite there - finding
  all 2,698 checks and executing a sample of them with no network at all. The installer got it wrong
  the first time in an instructive way: everything had in fact worked, but it read the installer's
  result code as blank rather than as success and declared the machine broken, while the tool it had
  just installed was reporting its own version two lines further down. It now distinguishes 'this
  failed' from 'I could not tell', because those are not the same and only one of them is worth
  worrying about. Running the mailbox checks themselves still needs a per-machine settings file that
  names real mailboxes, which no script can write.

- ✅ **Two scripts bring the mailbox test suite to the bare, offline test machines.**
  The test machines are built from scripts and are deliberately bare: no development tools, no copy
  of the project, and no network connection. The suite needs all three, so it had never been run on
  one - which is why twenty-odd checks that need a real mailbox have been waiting. There are now two
  scripts for it. One runs on the developer's machine and packs up everything the test machine
  cannot fetch for itself: the project's source taken from a named point in its history rather than
  from whatever is lying around, and every third-party component the build needs, gathered into a
  folder the test machine can build from offline. It then rehearses the whole build against that
  folder before shipping it, so a missing component is found in seconds on the fast machine instead
  of after a slow copy onto the slow one. The other installs the development tools on the test
  machine from a copy staged beforehand, checks the installer is the exact file expected, and then
  proves the result by actually building and running part of the suite - rather than concluding from
  an installer exiting quietly that anything works. It reports one of four outcomes and only one of
  them counts as ready. Neither has been run yet, and both say so. Nothing that ships to users is
  affected.

- ✅ **A test-machine script takes a machine's mailbox out of Windows Search.**
  That is the only thing that makes the second test machine different from the first. The test setup
  is two machines, one whose mail is indexed and one whose is not, and half the mailbox checks
  depend on that difference being real; but both are built from the same file and both come up
  indexed, so nothing had yet made the second one different and both halves would have been
  measuring the same machine without anyone noticing. The obvious shortcut - switching the whole
  Windows search service off - turns out to be the wrong answer and the script refuses it: that
  produces a machine with no search at all rather than a mailbox search has not been told about,
  which the product handles down a completely different and untested path, and it removes the very
  report the setup notes say to read to check the machine was built right. One existing check would
  simply have failed on such a machine. So the script takes the mailbox out of what Windows Search
  looks at and leaves the indexer running, and then proves the result by asking the index itself -
  twice, minutes apart, so "not indexed" can be told apart from "not indexed yet" - instead of
  reading back the settings it wrote. It gives four answers, two of which say in plain words that no
  answer has been reached yet, because a machine believed unindexed while it is quietly still
  indexing produces measurements that look fine and mean nothing. It has deliberately never been
  run. Nothing that ships to users is affected.

- ✅ **The test machine's mail server is fit for the job, and a script installs it and proves it.**
  The test machine's mail account was pointed at a local mail server that nobody had checked could
  hand mail back - the project's own notes flagged that the chosen server may only speak the other
  retrieval protocol, and that if so the whole design was unbuildable. It can: the feature was
  turned down in 2019 and added anyway five years later, so there is a version to fix on and the
  notes now say which. Three instructions for installing it were wrong, including one command that
  cannot work at all, and are corrected. Reading the server's own source also found two places where
  it departs from the mail standard, and one collision nobody had noticed: it refuses a blank
  password, while the test machine's account is set up with none - so one password has to be typed
  by hand, once, when the machine is built. The new script installs the server from a copy staged
  beforehand rather than downloading anything, checks the file is the exact one expected, and then
  actually sends a message and fetches it back instead of merely checking that something is
  listening - including the awkward cases that break mail servers quietly. It has deliberately never
  been run, and two of its checks are expected to fail, which is why they are checks. Nothing that
  ships to users is affected.

- ✅ **Draft scripts build a test machine's Outlook profiles and data files without the Mail control panel.**
  They also switch the default profile. That part of rebuilding the test machine was still done by
  hand, once at considerable cost by driving the screen with a vision model. Four of the five things
  a test machine needs can be scripted: a profile with no mail accounts, a data file carrying an
  exact name, the switch between profiles with no prompt, and the identity signature. The fifth
  cannot: **nothing free can create a mail account**, for three unrelated reasons, so adding the two
  test accounts stays a manual step - now a precisely bounded one, taken once per machine and
  preserved by a checkpoint. The scripts have deliberately **never been run**: they were written on
  a machine with a real mailbox on it, so they were checked by reading rather than by trying, and
  each says so at the top and refuses to run on any machine that is not a test guest. Nothing that
  ships to users is affected.

- ✅ **Three test-machine scripts try to build the missing mail account, and none has been run yet.**
  Setting up that account by hand is the one step of the rebuild nothing could automate, and the
  only off-the-shelf component that can do it is one this project will not take. Two free routes
  remain - describing the account in a file Outlook reads at startup, or driving the setup window
  the way a screen reader does - and there is now a script for each, plus one that answers in two
  minutes whether the second route is possible at all. Each checks its own result and fails loudly
  rather than reporting a success it did not verify, and each says in its own header what it expects
  the machine to look like first.

- ✅ **The test machines get an Office configuration of their own, narrower than the workstation's.**
  It installs Outlook and the rest of Office, without Visio, Project or the proofing tools a test
  guest never uses. It carries a licence key, so it stays in the scratch directory and never reaches
  the repository; what it contains is written down in the rebuild notes so it can be reconstructed
  from the key alone.

- ✅ **A test machine can say whether it has a second mail account for the identity checks.**
  That account is one the identity checks can create a draft in, so a machine without one can leave
  those checks out of a run instead of running them to be told they proved nothing. What such a run
  says is unchanged: a machine that runs them anyway and still finds no account to use says so in
  the log rather than passing quietly.

### Changed

- 🔧 **The task pane chooses and inserts a signature with the mail server's code, pictures embedded.**
  "Select the best signature" lists the signatures, and puts the chosen one in the mail, with the
  code the MCP server's drafts use (the code-sharing refactor's stage S4, Q154). The signature's
  pictures are embedded in the mail now, as the server's drafts embed them: the pane inserted them as
  links to the signature's folder on this computer, which a recipient never sees (D47). The excerpts
  the AI chooses by are read the way the server reads them: each file in its own encoding, the HTML
  through the server's converter to text, and a signature that has only a Rich Text rendition
  excerpted too. The recipients the pane tells the AI about carry the SMTP address of an
  Exchange recipient where Outlook knows it, where the pane gave the Exchange (X.500) address. The
  draft the pane writes is laid out as before, one paragraph a line.

- 📦 **Setup puts the mail server's own .NET 11 runtime in your profile, with no administrator rights.**
  The mail server runs on .NET 11 since the maintainer moved OutlookAI to it on 2026-11-18 (Q172),
  and the installer no longer installs a .NET runtime for the whole machine. On every install and
  every silent update it downloads Microsoft's runtime package, about 40 MB, checks its length and
  SHA-512 against what Microsoft publishes, and unpacks it into `%LOCALAPPDATA%\OutlookAI\dotnet`;
  once the runtime is there, it downloads nothing. A silent update used to skip the runtime,
  because installing it for the whole machine needed an administrator's answer. The installer stays
  under the updater's 50 MB limit, since the runtime is never inside it. On a machine without
  internet, Microsoft's `dotnet-runtime-11.0.0-win-x64.zip` beside the installer is used instead,
  checked the same way. The server looks for its own copy first and then for a .NET 11 installed
  for the whole machine, and uninstalling removes its copy. If setup could not put the runtime in
  place, OutlookAI Settings says so and what to do. On a test machine whose only .NET is 10, a fresh
  install and a silent update from v3.1.0.325 each left a mail server that answered on its own
  runtime, with no administrator rights and nothing installed for the whole machine. Uninstalled
  there as Settings runs it, both of the uninstaller's phases ran, its own runtime went, and the
  machine-wide .NET stayed as it was.

- 🔧 **OutlookAI opens no mailbox in a newly started Outlook until it has settled (Q175).**
  Settled means at least 20 seconds old, the OutlookAI add-in finished starting inside it, and its
  main thread idle: under a tenth of a 500 ms sample, and answering a message within a second. On
  2026-11-15 the server started Outlook over COM and opened mailboxes in its first second, while its
  add-ins were still loading; the Teams Meeting add-in called back into Outlook and its main thread
  deadlocked. The wait narrows that race; another add-in's own calls are outside OutlookAI's reach,
  so it cannot close it. Only one OutlookAI connection starts or attaches to Outlook at a time,
  across every server of the user, and the others wait for it (Q175 (e)). Meanwhile mail tools
  answer `OutlookStarting` at once; the retry guidance after a start is 25 seconds, up from 15. An
  Outlook that started more than two minutes ago counts as settled, so connecting to it costs one
  look. The add-in writes a start-up stamp, `HKCU\Software\OutlookAI\AddIn`, at the end of its
  start; an older add-in is recognised by its tuning reconcile. A folder search of an Exchange
  mailbox that starts Outlook (Q212 (a)) goes through the same wait. Decided by the maintainer on
  2026-11-18. Measured on the two Outlook test machines, which have no add-in that could deadlock:
  the disconnect test's eight cold starts, six on the unindexed machine and two on the indexed one,
  waited 18.0 to 18.6 s for Outlook to settle, the 20 s age being the last condition met, and
  connected in 21.2 to 22.2 s, where a cold start took about 4 s before; a connection to an Outlook
  already minutes old waited under 40 ms. In an Outlook opened with its window, as a user opens it,
  the add-in had started 6.2 s, 14.2 s and 27.9 s in on the unindexed machine and 171 s in on the
  indexed one, on a day its Outlook started slowly, so such an Outlook can settle later than one
  OutlookAI starts; the two-minute limit ends the wait regardless.

- 🔧 **A draft puts a window on screen at most once per Outlook start, and never while Outlook is open.**
  OutlookAI writes every draft in a compose window it does not show. In an Outlook with no window of
  its own, Word lends that window its editor only after a compose window has been activated in the
  process, and the activation put the window on screen at the top left for 130 to 650 ms: for every
  draft, 50 of 50 measured on Office LTSC 2024 (Q149). As the maintainer decided (Q208 B), only the
  first draft after each Outlook start activates its window now, and every later one reads its
  editor straight away, which Q149 measured to work after one activation in 120 of 120 drafts. While
  Outlook has a window open, no draft activates anything at all (40 of 40). An edit read that way is
  kept by adding Outlook's Message form page before the editor is read and saving as the window
  closes, which kept all 100 measured edits and changed no message's class or form. The activation is
  recorded per Outlook process for every OutlookAI session of the user, and decided under a shared
  lock, so agents drafting at the same moment make one between them. Measured on the change itself,
  170 drafts in 22 starts of Outlook: the first draft of each of the 18 starts with no window showed
  its window, for 81 to 410 ms, no other draft showed one, every edit was kept, and a person typing in
  another program throughout 30 of the drafts lost none of 1,090 keys. Windows announced Outlook's
  windows as the foreground during most drafts, but never made one the foreground window.

- 🔧 **Connecting to Outlook adds no hidden window; the first draft that needs one does (Q175 (b)).**
  A session that found Outlook with no window used to add a hidden Outlook window of its own, the
  lifetime pin, as it connected. The pin is what keeps an Outlook with no window running when a
  draft's compose window closes (without it, 40 of 40 such closes ended Outlook), so it is now made
  by the first draft that needs it and shared by every session. A session that only searches and
  reads makes none; an Outlook it started then ends by itself some minutes after the last program
  lets go of it, as before pins existed.

- 🔧 **The add-in changes only its own entry in `~/.claude.json`, and leaves every other server alone.**
  When it registers its mail server for all of your projects, the add-in read Claude Code's
  settings file with .NET's `JavaScriptSerializer` and wrote the whole `mcpServers` section back, so
  every other server's entry came back reformatted each time. It now writes its `outlookai` entry
  into the file as it stands, the way it has always written a project's `.mcp.json`, and every other
  byte stays as it was; the new text is checked before it replaces the file, as before. An entry the
  add-in wrote keeps the keys added to it since, such as `env`, and loses those only a remote server
  has; an entry anyone else wrote is replaced whole, as before. Ours is written as
  `{ "type": "stdio", "command": "...", "args": [], "env": {} }`. Two kinds of file it used to
  rewrite are left untouched now, and the registration reports an error that says so: one whose
  `mcpServers` is `null`, and one that is not well-formed JSON although `JavaScriptSerializer` read
  it - single quotes, unquoted names, leading zeros or a line break inside a string, none of which
  Claude Code can read either. The code is OutlookAI.Core's since the add-in began to use it (Q154),
  and T1 `McpUserConfigTests` pins it. Measured on a test machine on 2026-11-18, over a settings file
  holding a server laid out by hand and our entry as an earlier version wrote it: the earlier add-in
  wrote the whole `mcpServers` section back on one line and unescaped the hand-written server's `é`,
  and this one changed its own entry and nothing else (runbook section 2.3).

- 🔧 **`read` prints a mail's folder as every other tool does: its names joined by `/`.**
  It printed Outlook's own folder path, with the store's name in front, backslashes between the
  names and Outlook's codes for `%`, `/` and `\` inside them, while `list_folders`, `find_folder`,
  `search` and `move_mail` print `Inbox/Projects`. The store and its `storeId` are beside it, and
  `folderId` is the folder's address. Decided 2026-11-18.

- 💥 **Every tool addresses folders and stores by ID; names and paths are for display only (Q131-Q135).**
  Decided by the maintainer on 2026-10-11, after research found a folder's own Outlook ID unique
  and stable: it survives renames, moves within the mail file, deletion into Deleted Items and
  restarts, and is never handed out again. A folder ID is that ID, shortened: one letter for its
  kind, the rest in Crockford's base32 with the bytes every such ID shares left out, and a last
  character that catches a mistyped ID. It is 34 characters for a folder in a data file and 66 for
  an Exchange folder; a store ID is 28. Nothing is stored and there is no table: the ID is the
  folder's own, so it opens the folder directly. `search` takes `store_id` and `folder_id`,
  `thread` and `list_folders` take `store_id`, `goto_folder` and `show_search_results` take
  `folder_id` or `store_id`, and `move_mail` takes `folder_id`. `list_accounts` shows each store's
  `storeId` and `rootFolderId`; `list_folders` and `find_folder` show every folder's `folderId`;
  `read` shows the mail's `folderId` and `storeId`; search hits carry their `storeId`, and hits found
  through Outlook their `folderId`; `move_mail`, `archive_mail` and `discard_draft` give
  `fromFolderId`, so an undo is `move_mail` with that ID. Display names still come with every
  answer, so a person can read along. A store whose name another store shares is told apart now,
  where a name could only be refused. A name or path where an ID belongs is refused, with the
  `find_folder` call that finds it; so is an argument this release removed, rather than being
  ignored, which for `search` would have searched every mailbox. Every store's ID is 28 characters:
  the top folder of an Exchange store Outlook does not cache is read too, once each time the list of
  stores is refreshed (2 to 3 ms, measured on the Exchange test machine), where its ID used to be
  its whole entry id, which changed when the store's caching mode did. While Outlook says it is
  offline, that top folder is not asked for, since a call that needs the server can then wait, and
  the store keeps its long ID until Outlook is connected again. A mailbox the profile holds twice,
  as its own cached store and as the online entry Exchange automaps for an account with Full Access
  to it, has one store ID and is listed and resolved once, through the cached entry;
  `outlook_health` calls the online entry a second view of that mailbox, not a store the index holds
  nothing for, `list_accounts` names the mailbox it shows (`secondViewOf`), and a search covers the
  mailbox once, through the cached entry, without checking the same mail live through the online one
  (decision 207). An Exchange folder ID no folder has is refused: Outlook in cached mode opens such
  an ID as an empty folder, which read as the top of the mailbox, so a search scoped to it answered
  for the whole mailbox.

- 🔧 **OutlookAI writes Outlook's Cached Mode settings as your own settings, never as policy.**
  Versions 3.0.0.319 to 3.1.0.325 wrote five of them under Windows' policy key for Outlook, which on
  most computers only an administrator may change. Where Windows refused, the tuning stopped; where
  it let them through, they took the offline period slider in Outlook's Account Settings away from
  you, and could overwrite a value your organization had set there. OutlookAI now writes them where
  your own Outlook settings live, and only reads the two policy keys, Microsoft 365 Cloud Policy's
  and Group Policy's, to show and leave alone a value your organization sets. Cloud Policy comes
  first, as it does in Outlook; it used to be ignored. Values an earlier OutlookAI wrote as policy
  stay where they are, shown as "Set by an earlier OutlookAI". The health report's policyConflicts
  names every tuning value a policy sets, found by reading both keys, where it used to name only a
  value OutlookAI had applied and then seen changed.

- 🔧 **Drafts use the defaults of the profile Outlook is running, and every profile is listed.**
  OutlookAI read only the profile recorded as Outlook's default, and when none was recorded it merged
  every profile into one list and took whichever account matched first. A draft now gets the default
  of the profile Outlook has open. list_signatures names every profile, marks the default and the
  running one, and shows each account's own defaults beside the ones a draft actually gets. It asks a
  running Outlook which profile it has open, and still never starts Outlook.

- 🔧 **OutlookAI follows the administrator's signature settings the way Outlook itself does.**
  When Group Policy or Cloud Policy forbids signatures, a draft gets none and manage_signature changes
  nothing. A default that a policy, or Outlook's own settings (Common\MailSettings), names for new
  mail or for replies and forwards no longer decides which signature a draft gets: the account's own
  default does. Outlook did that on the test machines on 2026-11-18 (Q200 (a), Q201), with the
  settings made before it started, as plain and as expandable strings: it put the account's own
  default in a new mail, a reply and a forward, put nothing in a reply or a forward of an account
  with no reply default of its own, and kept its Signature button open under a policy, so one message
  may have another signature or none there too. manage_signature changes no account's default, and
  deletes no signature, that a policy or a deployed default names (Q156): Outlook's own Signatures
  dialog shuts its default choice, its signature list and its New button under either, and rewrites
  every account's own defaults to the setting's as it opens, so after that the setting's signature
  goes in. With no setting the dialog leaves all of it open and changes nothing. OutlookAI never
  writes a policy.

- 🔧 **Signatures kept in an Exchange mailbox are left to Outlook.**
  Microsoft 365 Apps from Version 2303 can keep an Exchange Online account's signatures in its
  mailbox, with a copy on the computer named "Name (address)". OutlookAI treats that as possible on
  Microsoft 365 Apps unless DisableRoamingSignatures is set, and on perpetual Office, such as the
  Office LTSC 2024 the test machines run, only where Outlook's own copies show it. LTSC 2024 roams
  too: on the Exchange test machine it showed no sign of it while the testbed's connected-experience
  policy values were set, and had a signature from Outlook on the web within about three minutes once
  they were removed, as such a copy. Where an account's signatures may live in its mailbox,
  OutlookAI does not edit or delete their copies or change the account's defaults, because Outlook
  would overwrite the change or copy it into the mailbox, and says to use Outlook instead. Roaming on
  Microsoft 365 Apps itself is not tested: the test machines run perpetual Office only.

- ♻️ **A failed COM call is described in one place, where nine catch blocks wrote out its text (Q154).**
  The COM session and the mail service named a failed call as "COMException 0x" and its HRESULT in
  eight digits in nine catch blocks, each writing out the text of the shared
  `ComObjects.DescribeFailure` - which the add-in's compose code and the rest of the COM session
  already call - rather than calling it. They call it now, and say what they said before. The hex
  stage of the code-sharing refactor found them.

- ♻️ **Bytes are written as hex and read back in one place, and no id or hash changes case (Q154).**
  The code-sharing refactor's hex stage. Nine writers of bytes as hexadecimal, six of a SHA-256,
  two of a number and three readers each had their own copy - in the entry and store ids, the index
  search's rows, the special and archive folders, the profile's sections, the COM session's tables
  and properties, the send confirmation, the update_draft intents, the page tokens, the audit log,
  a signature's RTF pictures and the private runtime's download check - and they all use Core's
  `Hex` now, as do the server's and the corpus tool's calls of .NET's own writer. Each caller names
  the case it writes, the one it wrote before: uppercase for every id, as Outlook writes an entry
  id, and lowercase for the hashes and tokens that were. Five of those outputs were pinned first, to
  values computed outside OutlookAI, and held before the fold and after it: the send confirmation's
  hash, the update_draft intent key, a page token, the audit log's writer mutex and the picture
  SHA-256 list_signatures reports. The reader of a conversation index, which .NET's hex number
  parsing let take a space inside a byte, refuses one now, as the store id's reader did: all three
  of its sources are Outlook's own `ConversationIndex` property, which Outlook writes as two digits
  a byte with nothing between them. Three checks of hexadecimal digits that stage S6 had missed -
  in a profile section's registry key, a resumable search's token and the audit log's reading of a
  CRC - use `Hex` too, and the COM session reads a binary value that comes as an array of
  single bytes in one place, where it had two copies. What stays is .NET's own formatting and
  parsing of a number in a message or an escape, such as an HRESULT, a character's code or a store
  hash's digits.

- ♻️ **The mail server keeps one copy of what it had written two to five times over (Q154).**
  The code-sharing refactor's stage S6 folds the server's own repetition that the survey of
  2026-11-17 found, its groups S01 to S14. The index search and the exhaustive scan check a search
  term through one rule, `SearchTermRules`, where each had a copy that differed only in how it
  refused a bare `*`: both now say "Bare '*' is not a valid term." The index search renders its
  dates with the exhaustive scan's `DaslDateLiteral`, the same format and the same rule for a local
  time. Hexadecimal digits are checked in one place, `Hex`, in the add-in's part of OutlookAI.Core,
  where the add-in's editor of `~/.claude.json` finds it too; the five copies differed in what they
  made of an empty string, and each caller now says. The send confirmations, the resumable searches
  and the update_draft intents make their random tokens with `Hex.NewRandom` and list their expired
  entries with `ExpiringEntries`, where each had its own copy of both. Everything that lives under
  `%LOCALAPPDATA%\OutlookAI` - the audit log, the session registry, Outlook's start log, attachment
  downloads, signature backups, the add-in's temporary HTML and the private .NET runtime - finds
  that folder through `SharedState.Root`, where six places built it; the check that no test names
  the real audit log watches that name too. The COM session opens an item by its EntryID in one
  place, `OpenItem`, where 21 work items wrote out the two calls; checks a mail item through
  `OpenMailItem`; and reads where a moved item is through `ReadMoveOrigin`, for both moves.
  move_mail and archive_mail walk their batch through one `MoveBatch`, and the index store map
  decides a store found by name and one found by identity through one `DecideStoreMatch`. The corpus
  tool's PNG writer takes its CRC-32 from the audit log's `AuditCrc32` and its designated Calendar,
  Contacts and Tasks from the product's `SpecialFolders.OpenDesignated`, where it kept copies of
  both; its build and re-index write one manifest header, its two populations one conversation
  message, and its date and placement probes pick their rung through one `CorpusProbeLadder`. The
  draft tools' descriptions that two or three tools wrote out word for word are constants, and the
  writing rules each tool checked before its call are one `WritingDraftAsync`. Left as they are,
  with their reasons in the TODO list: the two draft results' shared fields and the tools' and the
  COM host's parameter lists. The corpus tool's thirteen operations, which each opened Outlook, its
  namespace and its stores and released them the same way, do it through one `WithOutlook` since
  S5b, the twelve that work in one store through `WithCorpusStore` on it. The product's session, the
  live tests' mailer, the compose drafts and the corpus tool start Outlook and find a store by its
  name through one `OutlookAutomation`, and the mailer and the corpus tool release a COM object
  through Core's `ComObjects.Release`, where each kept a copy. Since S6, Core declares each Win32
  call once: the audit log's opening of its file and the attachment writer's of a folder or a file
  were one declaration written twice, as were the session registry's and the start observer's
  closing of a handle, and the audit log declared the current process's id beside the one the
  add-in's window calls declare.

- ♻️ **The live tests end and restart Outlook through one shared helper.**
  The compose measurement's stop in `LiveComposeProbe`, whose `Application.Quit()` from a client
  holding nothing else was private there, and the quit protocol `LiveDisconnectRecoveryTests`
  carries are folded into `T2/LiveOutlookRestart.cs`, so AGENTS.md's rule 7 for Outlook is kept in
  one place. The Inspector and Outbox counts come first, and an Outlook that ends by itself while
  they are read ends the stop: in the Q149 smoke on 2026-11-18 a cell's stop met the Outlook the
  cell before had just let go of, and the product's start guard reported it as an Outlook that
  exited while it started. An Outlook that ends by itself while the stop first waits for it gets the
  same five seconds before the next start as one a rung ended, where it used to get none: in that
  smoke a start less than a second after such an exit was followed by a crash in
  `mso40uiwin32client.dll`, though whether the one caused the other is not known. A step the stop
  and the compose measurement bound in time throws its own failure, as a direct call would.
  `Task.Wait` handed it on inside an `AggregateException`, which no catch around such a step looks
  for, so the stop's catches never matched, and a compose that threw failed the whole Q149 test
  instead of being recorded as its cell's result. A person's windows are closed only after three
  minutes without input. The rungs go in order: WM_CLOSE to Outlook's windows, its hidden Explorers
  closed, then `Application.Quit()`, and Outlook is never killed. One rung is new, for a window-less
  Outlook that a client still holds, which `Application.Quit()` would only park: one Explorer shown
  through the product's goto surface and closed the way a user closes it, as the disconnect test
  ends Outlook. It counts only this Windows session's Outlook, by the product's own rule. A start
  opens Outlook with its window through the shell, as the guest runner does, so the Outlook inherits
  no handle of the test host: one the test host started itself held a live run's output pipe open
  for 50 minutes after its suite had finished (Exchange VM, 2026-11-18). The measurement stops and
  starts Outlook through it. The disconnect test keeps its own scenario, which is what it tests, and
  takes from it its window probe, its Outbox refusal, its user-protection meanings and its one test
  for a session that died under a call, the `ObjectDisposedException` it met on the Exchange test
  machine among them (runbook section 4.4a, run 8). `T1/LiveOutlookRestartTests` pins the order of
  the rungs, that test for a dead session, that a bounded step throws its own failure and not a
  wrapper, that an Outlook that ended by itself is settled before the stop returns, and that nothing
  in it kills Outlook.

- ✅ **The testbed can build an installer that carries the mail server, to test the installed product.**
  `Testbed/host/Publish-AddInPayload.ps1 -WithServer` publishes the mail server and
  `OutlookAI.RuntimeSetup` into a testbed build's installer, as a release does, signed with the
  throwaway certificate and versioned 99.99.99.0 like every testbed build. It proved the private
  .NET 11 runtime on `OutlookAI-Unindexed2` (Q172, runbook section 4.12): a fresh install and a
  silent update from v3.1.0.325, each run as the logged-on user without administrator rights, left
  the runtime in place, the server answering `initialize` and `tools/list` on it, and the COM host
  starting on it through `DOTNET_ROOT`, which without that variable refused to start there.

- ⬆️ **The mail server, its tests and its tools build for .NET 11, and the add-in as C# 15 (Q172).**
  The maintainer moved OutlookAI to .NET 11 on 2026-11-18. Every project under `McpServer/` targets
  `net11.0-windows`, `OutlookAI.Core` and the audit-log stress writer keep their `net48` target
  beside it, and the security analysers run at `11.0-all`, whose 94 rules are the ones `10.0-all`
  ran, so the build found nothing new. System.Data.OleDb, Microsoft.Extensions.Hosting and
  System.Drawing.Common move from 10.0.10 to 11.0.0 with it - OleDb's 11.0.0 still carries the .NET
  Framework asset Core's `net48` build uses - and `Tools/Test-VulnerablePackages.ps1` finds no known
  vulnerability in any package of the seven projects. The VSTO add-in stays on .NET Framework 4.8,
  because a VSTO add-in cannot run on modern .NET, and compiles as C# 15 instead of the C# 7.3 a
  .NET Framework project gets by default; `Properties/LanguageFeatureStandIns.cs` declares, as
  source, the few types records, init accessors, required members and argument-expression guards
  look up by name, and a feature that needs the modern runtime the compiler refuses there. What a
  user sees of it is the runtime setup now brings for the mail server.

- ✅ **A compose test names each window it finds on screen, at its start and in its failure.**
  `LiveHeadlessComposeParityTests` failed twice on 2026-11-18 with "user-visible windows 1 -> 3" and
  nothing else: it counted the windows, so nobody could say what the two new ones were. And some runs
  found the two already showing when the tests began, counted 3 -> 3 and passed. Both of its tests
  now list the counted windows in their output as they start, and again after composing, through
  `ComposeSurface.UserVisibleWindowHandles`, a read-only list made with the count's own filter; when
  the count goes up, the failure names them too, as does the first test's check that a window-less
  Outlook starts with none. Each window is given by its class (`rctrl_renwnd32` is Outlook's own, a
  compose window among them), its title with everything but the window's kind reduced to the test
  tag and a hash, its process, its rectangle, whether it is cloaked or minimised, when it was read,
  and whether it is new. The count and the list are set side by side, so a drift between their
  filters shows. What the tests assert is unchanged.

- ✅ **The test machines carry the .NET 11 SDK beside .NET 10, on new checkpoints (Q172).**
  OutlookAI is moving to .NET 11 (Q172), and nothing on .NET 11 can be tested until the test
  machines have its SDK. Each of the three build machines has a new base checkpoint,
  `CP-04-SDK11-TEST-READY-UTC`, which the build runner now restores: SDK 11.0.100 and, beside it,
  .NET 10's SDK 10.0.404, both staged as media and held to the SHA-512 in Microsoft's release
  metadata. SDK 10.0.404 is there because a branch still on .NET 10 is built by the newest SDK on
  the machine, and SDK 11 builds .NET 10 against the 10.0.15 targeting packs, which a machine with
  no network has only from that SDK; it replaced 10.0.401 in place. Each build machine made its own
  checkpoint by the same steps, one at a time, each under the runner's lock and a lease. The
  Exchange test machine has both SDKs too, in a new checkpoint, `CP-10-SDK11`, made on the one with
  Office re-armed. The two Outlook test machines have new frozen-clock checkpoints with
  SDK 11 beside 10.0.401 (10.0.404's install took 7 minutes, more than the indexed machine's clock
  margin allows); the live runner restores them since the code moved to .NET 11, because there a
  .NET 10 branch would lack those packs, and the clones of those two machines are new copies of
  them, `OutlookAI-Indexed3` and `OutlookAI-Unindexed3`. Every old checkpoint, and the clones that
  copy the old ones, stay for the branches whose runners still name them. `Install-DotnetSdk.ps1`
  installs SDK 11 by default, builds its probe for the major it is asked about and pins that SDK
  there, and installs without verifying for a frozen machine, where every second before the
  checkpoint moves its clock. The live runner
  names the SDK each machine's install step must find, and a relative output folder no longer
  breaks the live-tier payload's offline feed check.

- 📝 **The README's section on Outlook's lifetime and tray icon describes the lifetime pin.**
  It was written before pins existed. It now says that the first draft written in an Outlook with no
  window adds one hidden window, which every agent session shares and connecting never adds; that the
  last session to end closes it, whichever made it, and Outlook then ends at once; that your exit, by
  File > Exit or by closing the last window you opened, makes the sessions close it; and that a
  session whose process was killed leaves its hidden window to the others or to the next session,
  with File > Exit ending Outlook meanwhile and `outlook_health` naming it. The 10-12 minute exit is
  now given as what it was: one measurement, on Office LTSC 2021, before pins. What Exit Now does with
  a pin present is said to be unmeasured, where the old text claimed it disconnects and exits at once.
  And where it said every draft is written in a compose window nobody sees, it now says what Q149
  measured on Office LTSC 2024 and what changed with it: with no Outlook window open, the first draft
  after each start of Outlook shows its compose window near the top left of the screen for 80 to 650
  ms, which on the test machines took neither the foreground nor a key from a person typing in
  another program, and no later draft shows anything.

- ♻️ **The scripts share one copy of the helpers they restated, in Testbed/guest/ScriptHelpers.ps1 (Q166).**
  Seventeen scripts each declared the same `Invoke-NativeCommand`, the wrapper every program call
  goes through so that Windows PowerShell 5.1 cannot end a script on the first line a program writes
  to stderr. All seventeen now dot-source it from `Testbed/guest/ScriptHelpers.ps1`, where
  decision 215 (a) put the helpers the scripts share, so that every route that carries the guest
  scripts to a test machine carries the file too. It is a dot-sourced file and not a module
  because a module does not pass the caller's error settings through: measured 2026-11-18, a script
  running a program through the wrapper from a module died on that first stderr line under Windows
  PowerShell 5.1, and came back with the program's exit code when the same function was dot-sourced.
  `Tools/Checks/check-powershell-51.ps1` now follows dot-sourced files, so a call through the shared
  wrapper counts as one through a copy always did, and its new check 4 fails a copy of a shared
  helper anywhere but in a script marked `# travels-alone:`, and there a copy that differs from the
  shared one. Six cases that break those rules were run against it here, and it failed each.
  `Testbed/README.md` names the file, and its step 4a stages it with the other guest scripts. The
  build and release scripts' own helpers went the same way: the run of a program into log files
  under a deadline (`Invoke-Logged`, five copies), the tail of such a log (`Show-LogTail`, four),
  the quoting of one argument (`Format-Argument`, two) and the elevation test (`Test-IsElevated`,
  two). The build-VM runner, which reaches the VM on its own, keeps its own run of a program,
  which quotes its arguments and stops the whole process tree, as `Invoke-LoggedStep`. And three
  `Tools/` scripts that restated the import of the running edition's own Security and Utility
  modules now dot-source `Testbed/host/OwnEditionModules.ps1`, as the host scripts do. The
  progress line `Say` had 27 copies in 12 bodies: the 18 that print the time and the message, and
  append the line to the script's log where it keeps one, are one function now, in the shared
  file, and the SDK installer, which reaches a fresh machine on its own, keeps an identical copy
  the check holds to it. Five scripts that collect their lines into a report and print them
  without the time now share `Write-ReportLine`. The mail sink's installer, whose log starts only
  once its guards have passed, calls its own `Write-SinkLog`, and the build-VM runner, which
  reaches the VM on its own and appends to its log in a way the host can read while it grows, its
  own `Write-RunLog`. `Say` appends only to the log of the script that calls it, read as
  `$script:LogPath`: it used to take `$LogPath` from whichever script up the chain had one, so
  `Publish-LiveTierPayload.ps1`, which keeps no log, wrote its progress into the live runner's
  run.log as well, which a live run showed on 2026-11-18. The live tier's runner and the
  trust-removal probe, the last two runners with copies, dot-source the file too, for their program
  calls, `Say` and `Stop-Refused`, and share with the build-VM runner `Format-Seconds`, which in the
  live runner had shown 59.6 s as "60 s", the run directory's name (`New-RunId`, which takes the
  guest's kind or the probe's scenario as a label), the TRX reader (`ConvertFrom-TrxText`), the lease
  decision (`Get-LeaseDecision`, now handed the run's whole lease reason) and the exit code of a
  task on a guest (`Get-TaskExit`, whose line the probe now writes as the live runner does). The
  clone maker's `Stop-Refused` went the same way. What the live runner and the probe each wrote out
  about a guest - its Hyper-V id, its state, the wait for its heartbeat and the restore of a
  checkpoint over it - is `Get-TestbedVmId`, `Get-TestbedVmState`, `Wait-TestbedHeartbeat` and
  `Restore-TestbedCheckpoint` in `Testbed/host/TestbedPool.ps1`, beside the narrow queries they ask,
  and the build-VM runner waits for its heartbeat through the same function. So the probe now saves
  a running guest before it restores a checkpoint over it, as the live runner has since 2026-11-18,
  and asks for a heartbeat with one narrow query where it asked 23 WMI operations. `Restart-Guest.ps1`
  and `Set-GuestClockFrozen.ps1`, which print what a heartbeat reads, share `Get-VmHeartbeatStatus`.
  Check 4 now holds the host's pool and lease files, `TestbedPool.ps1` and `TestbedLeasePath.ps1`, to
  one copy as well. The guest guard went the same way: whether a session is a test guest's - its
  account and the prefix of its computer name - and the refusal a script throws on any other machine
  are `Test-GuestIdentity` and `Assert-TestbedGuestIdentity` in the shared file, which each script
  hands its own words for what it would have done there. Five guest scripts call them from the file;
  five that reach a machine on their own - the first-logon script, the SDK and mail sink installers,
  the crash-dump arming and the unelevated Outlook start - keep copies check 4 holds identical. Where
  a guard took an empty prefix to match every computer name, the shared one refuses it, and the
  Exchange VM's three measurements and the Outlook start now refuse in the same words as the rest.
  `check-testbed-references` check 9 accepts the shared guard only after the script dot-sources the
  file, and fails a shared guard that stops reading the account or stops throwing; both failures were
  made on a scratch copy, and it reported each. The Outlook profile scripts' helpers went the same
  way: whether a registry key is a real Outlook hive or the installer's shell, the note on the other
  hives, the hives and the profiles on a machine, the profile a .prf names and whether a profile
  entry is a mail account were restated in `New-OutlookProfile.ps1`, `Set-DefaultOutlookProfile.ps1`,
  `New-TierProfile.ps1` and `Build-Corpus.ps1`, two to four copies each, and are one copy in the
  shared file now; each script keeps the Office constants they read. `Build-Corpus.ps1`'s
  `Write-Preflight` and `Reset-HubPopulation.ps1`'s `Write-Line`, one three-line function under two
  names, are the shared `Write-LogLine`. And check 4 now reads a function defined inside another one,
  as a self-test's helpers are, as well as one at the top of a script: a scratch copy that hid a copy
  of a shared helper inside a self-test failed it. The other guest guards went the same way: whether
  the session is elevated, whether PowerShell is 64-bit and whether Outlook is running are
  `Assert-Elevated`, `Assert-Bitness` and `Assert-OutlookNotRunning` in the shared file, each refusing
  in one form with the words its script hands it, in place of ten copies in six files, two of them
  named `Assert-OutlookClosed` and one written inline. The profile writers' reason is said once, in
  `OutlookMapiInterop.ps1`, and the SDK and mail sink installers keep copies check 4 holds identical. The choice of the hive a profile script works in (`Select-OutlookHive`,
  two copies and an older one in `Build-Corpus.ps1`, which now reads the shared one's choice) and
  what `-Verify` does about ImportPRF (`Resolve-ImportPrfClearance`, two copies) are one copy too,
  and so is the one helper of the self-tests with a single form, `ConvertTo-SelfTestText`, which four
  of them compared values by. The measurements' short hash went the same way: the two scripts on the
  Exchange machine that print a value they may not show as a SHA-256 prefix each had their own
  `Get-ShortHash`, one of 12 digits that ignores case and one of 8 that keeps it, and both call the
  shared one now, each saying which it needs, so every hash they print is the one they printed before.
  The second's hash call is proven by its self-test alone, which pins the exact digits: that script
  runs only on the Exchange machine. The attach to a running Outlook went the same way. Four scripts
  attached each their own way - `Reset-HubPopulation.ps1` through the Running Object Table and then
  the class factory, `OutlookDialogUia.ps1`'s dialog helper through the class factory once it had
  seen Outlook run in its session, and `Measure-FolderIdResolution.ps1` and
  `Measure-SignatureRoaming.ps1` through the table alone, which an Outlook that has not yet lost the
  focus is not in: the first could not attach on a test machine where Outlook had just started. All
  four use `Connect-RunningOutlook` now, which asks the table first and the class factory only once an
  OUTLOOK.EXE runs in the caller's own session, so it never starts one, and says which way it went.
  The MCP server's stdio client is one copy too: `Invoke-GuestMeasure.ps1`, kept as
  found until then, and `Measure-PinWindows.ps1`, whose sessions were a C# class of its own, both use
  `Testbed/guest/McpStdioClient.ps1`. It took on what the second did that it did not: a read still
  waiting when a reply's time runs out is kept for the next one, where the next read used to throw; a
  server starts without the initialize for a caller that times it; a request goes out now and its
  reply is read later; a silence can give back nothing instead of throwing; a reply must carry a
  result or an error; and the end of the server's standard error is waited for. A server that cannot
  be initialized is stopped, and the error says what it wrote. `Set-SignatureSettings.ps1`'s
  self-test now talks to a stand-in server through it as well as checking its pure parts.

- 📝 **If Office's own grace reset fails, the Exchange test machine may get Microsoft 365 Apps (Q206).**
  The maintainer's decision Q206, 2026-11-18, now in AGENTS.md beside the machine's internet
  exception. Its Office LTSC 2024 was never activated, and since the grace ran out in late October
  Outlook there creates no item through its object model, so the machine runs only the tests that
  write nothing. In his words: *"a with fallback to b with fallback to reporting to me for spla mak
  keys"*. First Office's own grace-period reset, `ospprearm.exe`; if that does not bring item
  creation back, Microsoft 365 Apps, activated by the test account's own licence, a licensed
  component allowed on that machine alone; and failing that, a volume key (SPLA MAK) that the
  maintainer supplies. Every other condition of the exception stands. `QUESTIONS.md` records the
  decision with his other answers of that day, to Q138, Q163 to Q182 and Q194 to Q215.

- ♻️ **The test suite keeps one copy of each stand-in, helper and Win32 call its files share (Q154).**
  Nineteen test files each declared a stand-in for Outlook's connection that runs an operation
  straight against a fake session, seventeen of them byte for byte, and the suite held six more
  copies of it and of two other such stand-ins under four other names. Thirteen declared a stand-in
  for the search index that answers the questions the product asks around a search, in nine bodies
  that had drifted apart one option at a time. Twenty-five files read the test project's folder out
  of the test assembly in their own words, and three parsed the committed example live-test settings
  the same way. `McpServer/OutlookAI.McpServer.Tests/Support/` now holds one of each:
  `DirectGateway`, which also takes the two positions two of its copies had drifted into (Outlook
  reported as not connected while an operation still runs, and every operation refused as by an
  Outlook that does not answer), `DeadGateway`, `RecordingGateway`, `StubIndexClient`, which the
  three suites whose index answers in its own way derive from, `IndexRows`, `RepoPaths`, moved
  there, and `ExampleLiveSettings`. `StoreIndexHashScopeTests` keeps its own index stand-in, which
  answers different questions. Thirteen of the checks that read the suite's own compiled code or
  source carried private copies of the same reading helpers - a walk over a method's instructions,
  the Category=Live trait, a collection attribute's argument, a live test's source file and the body
  of one member in it - and now share `IlReader`, moved there from `T1/`, `Traits` and `SourceText`;
  three of them read the instructions byte by byte, and now walk the opcodes as the others did.
  Smaller copies went the same way: the in-memory prompt registry (`FakePromptRegistry`, which also
  replaces the narrower one `PromptLintTests` kept), the setter of a stand-in's out parameter
  (`ProxyArgs`), settings naming only the hub (`HubOnlyLiveSettings`), a PST store's EntryID
  (`TestFolders.PstStoreEntryId`) and the checks of an over-long subject's refusal, which an
  in-process and a stdio test each wrote out (`DraftRefusals`). Six suites around the sweep and the
  exhaustive scan spelled out the same one-mailbox profile - its SID, store, index prefix and
  frontier - and three of them its mail, its sweep, an exhaustive search of its Inbox and a session
  that answers the store list, the sweep and the scan as well: `AliceMailbox` holds it once, and
  `ThreadWalks` the conversation walks three thread suites each built. Five more suites take that
  SID, store, prefix and frontier from `AliceMailbox`, where they spelled them out, with the index
  that holds the store alone and the one that holds none, which four and three of them each built,
  and `ThreadWalks.Member` the thread member two of them wrote. The two thread suites' three pairs
  of parallel tests - a named store raises its list's code, the sentence names the stores, a capped
  list says how many more - keep their names and assert through `ThreadWalks`, each failure naming
  its list. The suite's eighteen stand-ins for the COM session are two: eleven that answered the
  same members - the profile's name, its stores, a folder's description and, all but one, the
  freshness sweep - three of them word for word, are one `StandInSession`, on which a test scripts
  any other member it needs, such as the exhaustive scan, an item read, a conversation walk or the
  folder lists; and seven that handed every call to a fake of the test's own, three of them giving
  each unanswered out argument and result its default, are one `ForwardingSession`. The live tier
  went the same way. Its test tag is spelled once, in `LiveOutlookTestMailer.TagText`, which the
  bracketed subject tag and the test-folder prefix are made of, and the mailer's twelve inline
  ordinal tag checks, and four more elsewhere, now ask one ordinal predicate, `CarriesSubjectTag`;
  the reading filters that leave test items out of a search, a census or a ground truth ask
  `MentionsSubjectTag`, which ignores letter case as each of them did. A test's tagged subject is
  built once (`TaggedSubject`), the subject the signature work's Outlook-made draft gets since among
  them. `LiveDraftChecks` holds the draft suites' reads of a draft and of where its body sits, which
  seven suites restated in up to four copies each; `LiveDraftCleanup` deletes their drafts and
  proves none is left; the folder suites' clean-up and zero-artifact proof are methods of the
  fixture they share, the move/archive suite's and its stdio test's since too; the two post-suite
  sweeps run one proof, `LiveArtifactSweep.ProveNoTaggedArtifactRemains`; `AuditLines` reads the
  audit lines six suites counted, each still naming which log; the stdio suites read the tool
  surface through `ToolShapes` and write and read the wire through one `McpWire`, which the stdio
  client and the conformance test each wrote out; the corpus helpers live in `HubCorpus`, scratch
  files go through `ScratchFiles`, a store's id through `LiveIds`, an index scope through
  `LiveIndexScopes` and the registry through `RawRegistryValues`; and four fixtures read the test
  project's folder through `RepoPaths`. The live tests make the window calls the product makes too
  through Core's `NativeWindows`, where six of their files declared their own - 24 declarations,
  with two copies of its `RECT`, five of its callback and four of its constants - and `TestWin32`
  declares the calls only tests make: the screen capture's, and those two or three test files each
  declared, a foreground window, a child walk, a window's text and a posted message among them. Two
  test helpers that wrote bytes as hex by hand use .NET's own writer. No test changed what it
  checks: the build machine ran the same tests, by name, before and after each step, and the 164
  live test methods are the same 164.

- ✅ **The draft tests read where Word lays out a reply's text, and check no draft became a form.**
  The reply, reply-all and forward tests now read Word's own layout of each draft and fail when its
  first paragraph has a top border or its text or signature is not above the quoted header block
  (Q209 (a)), with and without a signature. They also read every new draft's, reply's and forward's
  message class and form-definition properties, because a draft is now committed through Outlook's
  Message form page and then saved again, and a one-off form would be sent as an attachment. The
  pin test checks that connecting makes no hidden window and that a session is listed while it is
  connected and gone once it ends; the disconnect test expects the close of Outlook's only window to
  end it, with no pin to hold it. `T1/OutlookSessionRegistryTests` pins every rule of the session
  list and its file, `T1/QuoteSlotRulesTests` where the text goes, and the source checks that
  connecting makes no pin and that a draft gets its editor only through the decided route.

- ✅ **A live run that starts Outlook closes Office's activation wizard and records it.**
  Past Office's 30-day activation grace, which the Exchange test machine's real clock has crossed,
  every start of Outlook with a window raises the modal Microsoft Office Activation Wizard before the
  main window, and the main window waits for it: one start sat 277 seconds on 2026-11-17 until it was
  closed by hand. The script that starts Outlook for a run now closes that one window with its own
  Close button, says so, and fails if it stays up, and the run's summary records how many times it
  was closed. On that machine the runner also starts Outlook no sooner than 240 seconds after a boot:
  started 38 and 79 seconds after one, Outlook put its programmatic-access prompt in front of the
  first address read although the antivirus read up to date, and started 2 and 8 minutes after one it
  did not. Past the grace Outlook there also created no item through its object model, from 2026-11-17
  until the grace was re-armed on 2026-11-18.

- ✅ **The Exchange test machine leaves out the signature tests until signature roaming is measured.**
  A signature or a default signature set on the test account could be carried into the mailbox by an
  Outlook that roams signatures, where no sweep removes it. Until that is measured for this Office
  (Q147, decided by the maintainer 2026-11-17), the tests that write a signature or touch an account's
  default declare a new capability, `LocalSignatures`, which the machine's filter leaves out, and the
  test-side MCP client there refuses `manage_signature` whatever its action.

- 📝 **An agent may sign in to Outlook on the web on the Exchange test machine, on six conditions.**
  The maintainer's decision Q161 (c), 2026-11-17, now in AGENTS.md beside the machine's internet
  exception: that machine only; the test account only, never the real mailbox; Windows' own browser,
  Edge; the password and the code typed as keystrokes from the host's encrypted store, as the sign-in
  script does; "Stay signed in" and device registration declined; signed out afterwards and the
  machine restored to its checkpoint. It exists to measure roaming signatures.

- 📦 **The release page lists each change's icon and headline, linked to its full entry (Q123).**
  The maintainer decided it on 2026-10-03, in his words: *"copy the release notes style and system
  and rules from the BrowserAI repo"*. Every changelog entry now opens with one of twelve icons,
  then a bold one-sentence headline, then the whole of what happened, under Keep a Changelog's
  groups. The GitHub release no longer pastes the section: `Tools/Publish-Release.ps1` makes the
  body from each entry's icon and headline, with a *read more* link to that entry's own lines in the
  tagged changelog, so a release page stays readable however much a release holds. On the day of
  the change this release's section was 157,366 characters, over the 125,000 GitHub accepts for a
  release body; its body is about 43,000. The script refuses an entry, a group or a legend that
  breaks the rules, a released section that changed, and a character on the release page that a
  person does not type, such as an em dash; `-CheckChangelog` checks the changelog on its own, and the
  build VM checks it on every run. The 216 entries written since v3.1.0.325 were converted: each
  one's opening sentence became its headline, and the rest of its text was kept word for word.

- 📦 **Releases are built, tested, signed and published on the maintainer's own computer.**
  GitHub's servers no longer build them. Installers are signed with the same OutlookAI certificate
  as before, so installing and updating work exactly as they did; what goes is the GitHub
  build-provenance attestation that releases used to carry, which only GitHub's own build service
  can make.

- 📝 **The subject option says how a renamed reply or forward threads in POP3, IMAP and .pst mailboxes.**
  When you give a reply or forward its own subject, OutlookAI keeps the draft in its conversation by
  restoring the conversation details Outlook resets when a subject changes, and in an Exchange or
  Microsoft 365 mailbox the draft also keeps the original's conversation id. A test machine measured
  that a data file differs: there Outlook gives the renamed draft a conversation id worked out from
  the conversation's topic rather than the original's, and refuses to let it be changed - so the
  draft still threads for its recipients and keeps its topic, but a lookup by the original's
  conversation id does not find it. The subject option's description now says so, and the check that
  runs against a real Outlook holds each kind of mailbox to what it actually does.

- 🔧 **Each mail file is found in the Windows search index by its identity, not its name (Q99).**
  Windows files every Outlook mail file in its index under the file's own name plus a number Outlook
  computes from the file's identity - Microsoft documents the formula - and OutlookAI now computes
  that number too. A search limited to one data file (.pst) therefore searches that file and nothing
  else: also when the file's own name differs from the name Outlook shows for it, as it can for a
  file named only in the Outlook profile or for an older-format file, and never another file that
  shares its name - one in another Outlook profile, or a copy at another location. A data file the
  index does not hold is now said to be unindexed instead of being searched through a same-named
  stranger, and search results name the mail file the way Outlook does, the name the other tools
  accept. `outlook_health` now says for every mail file how it was found in the index (`matchedBy`),
  under which index name, and lists the indexed mail files that belong to none of this profile's.
  Exchange, IMAP and Outlook.com mailboxes, whose numbers could not be tested here, are found the
  same way when the number fits and otherwise exactly as before, as are shared and delegate
  mailboxes.

- 🔧 **Each mail file is matched to the search index by its name and identity number together (Q113).**
  It is never matched by a guess. The index's address for a mail file carries both, so a file is now
  matched only where both are its own: a renamed data file is no longer searched through the entry
  the index still keeps under its old name, and two files that share a number are told apart by
  their names. Where the index itself cannot tell two of your mail files apart - they share both
  name and number - OutlookAI no longer picks one: a search limited to either is refused with the
  reason, a search across all your mail leaves that index entry out and names the two files
  (`index.storesUnmatched`) instead of mixing their mail in, and `outlook_health` marks them
  `ambiguous`. A search limited to a name two of your mail files share is refused the same way, and
  so is a search limited to a data file while the index's list of mail files cannot be read. IMAP
  and Outlook.com mail files are held to the same rule instead of falling back to their names, and
  `outlook_health` now reports such a match as `nameAndHash` (it said `storeHash`). Exchange
  mailboxes keep their current matching - by number, or else by name - until it has been measured on
  an Exchange test machine.

- 🔧 **An exhaustive search walks folders in a stated order, the one the folder listing promises.**
  Outlook does not define what order it hands folders back in, so a search that stopped part-way
  through was missing whole folders chosen at random - and stopping in a reproducible place is what
  makes continuing from it mean anything.

- 🔧 **An exhaustive search says which limit stopped it, because the two need different answers.**
  Running out of time means continuing is the only route to the rest; running out of result slots
  means continuing works but narrowing the folder or date range is cheaper. Both could be true at
  once and the answer never distinguished them, so half the advice an assistant gave about it was
  the wrong half.

- 🔧 **A continuation whose search has changed is refused, and no longer answers a different question.**
  Change the terms, the account, the folder, a date bound or a filter and the call is refused with
  the changed argument named; the page size may still differ from page to page. A handle that has
  expired, been superseded by a later page, or was never issued by this server each get their own
  explanation and their own way back.

- 🔧 **The check for new mail reports a refused sort apart from an unavailable column.**
  The two failures were caught together, so the flag that decides whether "the oldest mail is what
  is missing" is true could not say which had happened. They are now told apart and counted, which
  is what will settle whether that sentence has ever been true.

- 🔧 **An interrupted draft revision is recorded in the audit log.**
  That way an outcome nobody can state leaves a line saying so rather than a gap where a line should
  be.

- 🔧 **Outlook gets far more time to answer, so slow work finishes instead of failing.**
  Every mail operation now has minutes rather than seconds: a search of your whole mailbox can spend
  up to a minute in Windows Search and ten minutes checking Outlook for new mail, an exhaustive scan
  can run for ten minutes, and an ordinary operation has five minutes before the connection to
  Outlook is considered stuck. The reason is measured, not guessed - on a mailbox of this size one
  ordinary check for new mail across five accounts takes about a minute, which the previous
  thirty-second limit could not hold, so it failed on real searches and restarted the connection to
  Outlook every time. Slower answers are the deliberate trade: when you ask an assistant to do this
  work you are not sitting watching it, and a complete answer that took two minutes beats a partial
  one that took ten seconds.

- 🔧 **An exhaustive search has a time limit of its own, without slowing down everything else.**
  It used to share one limit with every other operation, so giving the scan ten minutes would also
  have meant waiting ten minutes to notice that Outlook had stopped responding to a simple read. The
  scan now has a limit of its own, and everything else keeps a much shorter one.

- 🔧 **The check for new mail may run for ten minutes, up from three, without slowing down anything else.**
  That check is the part of every search that asks Outlook directly about mail too new to be
  indexed, and on a large mailbox three minutes was not enough to finish it - so a search quietly
  came back having covered only some of your accounts. It now has a time limit of its own, so
  raising it does not raise the time every other tool waits before deciding Outlook is stuck:
  reading a mail, moving mail, creating a draft and listing folders are all unchanged at five
  minutes. The same ten minutes applies to fetching a conversation, which does the same kind of live
  check. The old three minutes came from a measurement taken while the sort defect above was still
  present, so it described the wrong work; ten minutes is a deliberate ceiling until the real cost
  has been measured, on the principle that a complete answer you waited for beats a fast answer with
  holes in it.

- 🔧 **An exhaustive search explains why it returned two results and still said there were more.**
  That mode stops after a fixed number of matching mails, and any `from`, unread or has-attachments
  filter you passed is applied afterwards - so a scan can stop at 25 candidates, discard 24 of them,
  and hand back one result while thousands more match further on. The answer now says the limit
  counted candidates rather than results, how many were discarded, and that the way to narrow it is
  a tighter folder or date bound rather than a filter that runs after the limit.

- 🔧 **Index freshness is reported for the worst account, not just the newest moment anything was indexed.**
  On a profile where one account lags hours behind another, the single figure hid exactly the
  account whose recent mail was at risk.

- 🔧 **The list of accounts a search reports as missing from the index is capped, and says so.**
  It was the one list in the mail server with no limit, so a profile with many local data files
  would have printed the whole list twice over in every search answer.

- 🔧 **How far behind the index is gets measured per account, not across your whole profile.**
  A busy account no longer hides how stale a quiet account is, so searches in a quiet account now
  check the right stretch of time for new mail instead of a window sized by your busiest one.

- 🔧 **Every setting is in one resizable OutlookAI Settings window with tabs.**
  OutlookAI Settings is now resizable with tabs - Outlook, Claude Code, Prompts, Buttons, Updates -
  instead of one tall fixed dialog, so it stops running out of room and nothing is cut off at any
  display scale.

- 🔧 **A mail search that only partly checked for new mail says so in a form software notices.**
  If the live check ran but could not cover everything it was asked to - a folder it could not read,
  a folder with more new mail than it reads in one pass, or simply running out of time - the answer
  now says so in a way your AI assistant can act on, instead of looking complete unless it happened
  to read the explanation. It already said so in words; now it says so in a form software notices.

- 🔧 **A draft with an over-long subject explains itself well enough for the assistant to fix it.**
  A draft with a subject past Outlook's limit now comes back saying the limit, how long the subject
  actually was, and that nothing was created - so the assistant shortens it and retries rather than
  reporting a failure to you.

- 🔧 **Update progress shows the moment it changes, in OutlookAI Settings and the sidebar alike.**
  Starting a check in OutlookAI Settings now updates the sidebar immediately and vice versa, instead
  of both places noticing within a second by checking over and over.

- 🔧 **Uninstalling OutlookAI keeps your settings in the registry.**
  Your prompts, quick buttons, Outlook tuning preferences and mail-server registration state now
  stay in the registry instead of being removed, so reinstalling picks up where you left off. This
  deliberately reverses the cleanup added in v3.1.0 - prompt text you wrote yourself is worth
  keeping, and nothing exports it.

- 🔧 **Every sidebar writing action carries a rule against drafts that read like AI wrote them.**
  Every writing action in the sidebar now carries the rule "Ensure there is no trace of AI both in
  wording and character use." It sits in the always-sent prompt, so it applies to every button and
  every instruction you type, and you can edit or remove it like any other rule.

- ⚡ **A stuck or starting Outlook is noticed at once, by asking Windows whether it responds.**
  The server now asks Windows directly whether Outlook is responding - which costs nothing and
  cannot itself get stuck - before trying to use it. A stuck Outlook is reported in a fraction of a
  second rather than after a long wait, and a search returns your indexed mail immediately instead
  of stalling first.

- ⚡ **Mail tools answer at once while Outlook starts up, saying roughly how long to wait.**
  If Outlook is closed or still starting, mail tools now answer straight away saying so and roughly
  how many seconds to wait, instead of blocking for up to a minute and a half. Outlook is started in
  the background meanwhile, and searches keep working from the index.

- ⚡ **Once Outlook is known to be stuck, mail requests answer at once instead of each paying the wait.**
  Previously each request discovered the problem on its own, so a stuck Outlook made every search
  and every account listing take up to two minutes, over and over. After two failures in a row the
  server answers straight away instead, and quietly re-checks Outlook every half minute so it
  recovers on its own the moment Outlook responds again - restarting Outlook still fixes it
  instantly. In this state a search returns your indexed mail in a fraction of a second, and tells
  you the live check was skipped.

- 🔧 **Each of the three search paths says what "search the body" means for it.**
  The index covers the body plus the text inside attachments; the live check for new mail reads the
  body as Outlook renders it; an exhaustive search reads the stored plain-text body only. Each now
  states its own answer in the result. The one that can cost you mail is the last: an exhaustive
  search is the mode you pick BECAUSE it is thorough, and it is the only one that cannot see inside
  attachments at all, and may not match a mail whose body exists only as HTML. It now says both, and
  says the remedy is the counter-intuitive one - search again WITHOUT the exhaustive option.

- 🔧 **Each part of a search says whether it matched whole words or fragments.**
  The index matches whole words; the live check for new mail matches fragments, so it returns
  slightly more. That is the safe direction, and it is also why a just-arrived mail can appear in a
  search and then drop out of the identical search once it has been indexed - which now has an
  answer in the result instead of looking like a bug.

- 🔧 **A conversation fetch says when it could have narrowed to one account and did not.**
  The field that used to report "the product chose this narrowing, you did not" had nothing left to
  describe once the product stopped narrowing on its own, so it now reports the decision that
  replaced it: you pointed at a mail, the account that mail sits in was there to be used, and the
  lookup covered your whole profile anyway. Nothing is missing because of it and there is nothing to
  do about it - it is there to explain replies coming back from accounts other than the one you
  pointed at, and to say where the narrowing went if you wanted it (name an account yourself). The
  advice attached to the old meaning went with it: a lookup you narrowed yourself is still told to
  drop the account, and the second remedy - which existed only to escape a narrowing nobody had
  asked for - would now name an argument that changes nothing.

- 🔧 **The update check waits out Outlook's startup, and an offline machine backs off.**
  The first check used to fire the instant the add-in loaded, which is the busiest moment Outlook
  has - it is opening your mailbox and starting every other add-in at the same time - so it now
  waits half a minute. And a machine that cannot reach the update server used to keep asking every
  ten minutes indefinitely: after three failures in a row the wait doubles each time, up to two
  hours, which takes a disconnected machine from 144 requests a day to twelve. Nothing changes on a
  machine that can reach the server. Recovery is immediate rather than delayed: reconnecting to a
  network clears the backoff and checks shortly after, and "Check for updates" always runs straight
  away.

- ♻️ **The signature and compose-document code is in OutlookAI.Core's add-in part (Q164).**
  The signature catalog; the signature environment, which is the folder, the policy, roaming and the
  Office build; the rule that chooses a draft's signature; the readers of the profile's accounts and
  of Outlook's profile keys; the conversion of HTML to text; and the readers and writers of a
  signature's three renditions and its pictures move from `OutlookAI.Core.Services` and
  `OutlookAI.Core.Text` into `OutlookAI.Core.AddIn` (`McpServer/OutlookAI.Core/AddIn/`), the one part
  of Core the add-in may use (the code-sharing refactor's stage S4, Q154). So does the code that reads
  and writes the regions of a
  compose document - the draft, the signature region Outlook marks `_MailAutoSig` and the quoted
  original it marks `_MailOriginal` - which the server had ported from the pane and fixed since, now
  `ComposeDocument`, with a choice between the server's soft line breaks and the pane's paragraphs,
  and with the rule that puts a reply's own text above the quoted message's header block
  (`QuoteSlotRules`, Q209 (a));
  the release of a COM reference, the test for a failed COM call and a COM object's identity, now
  `ComObjects`, to which `OutlookComSession`'s members of those names forward; the reading of a mail's
  recipients with the SMTP address where Outlook knows one, now `MailRecipients`, which also replaces
  the server's two copies of that loop; and the Win32 window calls, now `NativeWindows`, which three
  files declared for themselves, each a different subset. The server and the tests use all of it from
  there, unchanged, and T1 `ComChildObjectReleaseTests` reads `ComposeDocument` for a discarded
  `Bookmarks.Add` as it read `OutlookComSession`. The add-in uses it in the place of its own copies:
  `SignatureStore` is gone, the task pane reads and writes a draft through `ComposeDocument`, reads
  recipients through `MailRecipients` and releases through `ComObjects`, and the MCP registration
  prompt finds Outlook's windows through `NativeWindows`. Inside the add-in, the five handlers that
  each handed an event from another thread to their window share `UiThread.Post`, the pane's three
  edits and "Select the best signature" run in one scaffold around the compose document,
  `RunDocumentActionAsync`, where each wrote it out, and the settings dialog's prompt sections build
  their row of controls with the dialog's own button-row helper. Nothing in the namespace attaches
  to Outlook, starts a COM thread or opens the search index, and check 15 holds that for every file
  in it.
  The word the draft tools read as "no signature" and the refusal a signature request can end in
  move with the rule, out of `SignatureManager`, which stays where it is: it writes signature files,
  and only the server does.

- ♻️ **The fourteen files the add-in shared by linking are OutlookAI.Core's, in one namespace (Q163).**
  The MCP server, Core and the test project compiled fourteen of the add-in's source files as
  linked copies: the registry contract between the add-in and the server, the Office versions, the
  signature folder rule, the offline periods, the MCP configuration surgery and the decision when to
  act on it, the prompt store, its defaults and its checks, the update schedule, the tuning
  reconcile, and the Cached Mode settings with the seam for an account's offline period and the
  removal of what an earlier version wrote as policy. They are OutlookAI.Core's own now, public in
  `OutlookAI.Core.AddIn` (`McpServer/OutlookAI.Core/AddIn/`); the add-in uses them from the
  `OutlookAI.Core.dll` installed beside it, and nothing is linked into Core any more (Q163 (a)).
  `outlook_health` reads `~/.claude.json`, expands `${VAR}` references and compares paths with the
  add-in's `McpConfigEditor` where it had copies of all three, so `EnvironmentExpansionParityTests`,
  which held the two expansions equal, is gone; `ComposeSurface` takes its park coordinates from
  `McpRegistrationDecision`, so `ParkCoordinateMirrorTests` is gone too. The add-in's registry root
  is one constant, `AddInServerContract.AppKeyPath`, which the tuning, registration and prompt keys
  are built from. The add-in's own reader of `~/.claude.json`, built on `JavaScriptSerializer` where
  no test could load it, is gone as well: whether the file may be touched, what the `outlookai`
  entry is and the rewrite are `McpConfigEditor`'s, and T1 `McpUserConfigTests` pins them. Check 15
  holds the namespace to its rule with fourteen files in it, and the guest installer's 41 pinned
  lines, the payload's two contract files and checks 4, 12 and 13 follow the files.
  `Tools/Switch-AddInBuild.ps1 -BuildOnly` built the stage's three code commits and both merges with
  master, and left the host unchanged. On `OutlookAI-Unindexed`, in four passes alternating with the
  build from before the add-in referenced Core, every start reached `ADDIN-READY`, and the load times
  Outlook recorded differed more between two passes of one build than between the builds: 3.9 and
  4.6 s for the first start after an install and 0.8 to 1.7 s after it, against 4.2 and 5.9 s and 0.7
  to 1.4 s (runbook section 2.3).

- ♻️ **The add-in references OutlookAI.Core and may use one namespace of it (Q164, Q167).**
  Until now the add-in kept its own copies of logic the MCP server has in Core, because it did not
  reference Core (v3's "Phase-6 no-coupling choice"). Q154 (b) ended that, and this was the
  refactor's first stage, plumbing only, which changed nothing the add-in does.
  `OutlookAI.csproj` references Core's net48 build (Q167 (a)) and keeps two of the scripted
  builds' global properties out of Core's build: `PrepareForRunDependsOn` names a VSTO target Core
  does not have, and `DefineConstants` would replace Core's own `NETFRAMEWORK` symbol. Both add-in
  build scripts, `Testbed/host/Publish-AddInPayload.ps1` and `Tools/Switch-AddInBuild.ps1`, restore
  Core's packages first with MSBuild's `/restore`; guard 1's evaluation still runs no target and
  restores nothing, and their self-tests check both. `OutlookAI.slnx` holds Core, so F5 in Visual
  Studio builds it. The installer carries `OutlookAI.Core.dll` and
  `System.Reflection.DispatchProxy.dll` beside `OutlookAI.dll`, both listed and hashed in the signed
  add-in manifest, and a release stamps its version into that copy of Core as well. Core's
  `System.Data.OleDb` package adds nothing there: on .NET Framework its assembly only forwards to
  the framework's own `System.Data`, which Core's net48 build references instead. The add-in may
  use only the namespace `OutlookAI.Core.AddIn`, and nothing in it may attach to Outlook, start a
  COM apartment thread, open a MAPI session or open the search index (Q164 (a)), because the add-in
  runs inside Outlook's own process. `Tools/Checks/check-pinned-constants.ps1` check 15 holds that
  rule, and T1 `AddInCoreReferenceTests` the shape of the reference. Measured 2026-11-18: both build
  scripts built the add-in and left the host unchanged over 318 lines, the installer grew from
  43,159,206 to 43,435,319 bytes, and a commit from before the reference still builds, its restore
  finding nothing to do. On `OutlookAI-Unindexed` the build installed from `CP-08` and reached
  `ADDIN-READY` three times in a row, as master's did; Outlook recorded its loads at 2.11 and 2.13 s
  against master's 1.75 and 1.72 s, on a host running three build VMs, which two samples do not settle
  (runbook section 2.3, the last record).

- 📝 **The documents and the TODO list name no real person, address or machine any more (Q189).**
  They quoted the maintainer's computer name, his user name in paths under his profile, his
  personal and business addresses, the item counts of his mailboxes, a colleague's first name, and,
  in AGENTS.md, the Exchange VM's real mailbox. They now say "the maintainer's workstation",
  `C:\Users\<user>\`, addresses under reserved domains, rounded counts, "a colleague" and "the
  Exchange VM's mailbox". The Exchange VM's runbook section, `Testbed/README.md` and
  `Testbed/testbed.json` call its real mailbox `bystander@exchange-vm.invalid`, the name its live
  tests use for it. `Testbed/MEDIA.md` and the runbook call the Office configuration the workstation
  was built from `Workstation.xml`.

- ✅ **Tests use placeholder stores, names and addresses instead of real ones (Q189).**
  Twelve tests named their example stores after the maintainer's own Outlook profile: his address,
  his name and a real archive's file name. They now use `maintainer@personal.invalid`, `Alex
  Example` and `Archive Old.pst`, which test the same things - case, spaces and dots in a store
  name. Eight more used a third party's shared mailbox as the store in audit-log lines, a delegate's
  first and full name as store names, the maintainer's name in a signature and his surname as a
  sender to search for, and toy addresses at real domains, as two newer signature tests did too.
  They now use `shared-mailbox@exchange-org.invalid`, `Ann` and `Ann van Example`, `Alex Example`,
  `example` and addresses under `example.com` and `b.example`. Five live-test comments say "the hub" where they named the shared mailbox, and
  one no longer gives the item counts of a real mailbox. The testbed scripts' self-tests refuse a
  made-up workstation, `devuser` on `DEV-PC`, instead of the maintainer's own user and computer
  names, and the CodeQL script's licence fixture names `Alex Example`. The baseline of
  `Tools/Checks/check-personal-data.ps1` comes down with every such change, and is empty once the
  tests and the testbed need no real value at all.

- ✅ **The Outlook test machines run every test at a frozen time, so their test mail never ages (Q130).**
  Each machine rests on a checkpoint taken with Hyper-V's time synchronisation off, so every run
  starts at the same instant just after its test data's date: `CP-20C-FROZEN-CLOCK` on the indexed
  machine, `CP-14B-FROZEN-CLOCK` on the unindexed one. Before, the indexed machine's 160,000-item
  measurement mailbox would have stopped every run after 2026-11-01 until a 1.5-hour rebuild, and
  the hub mailbox aged out within days. A guard checks the clock before a run starts, a restart is
  refused or re-frozen, and the 30-day timing check reads the 30 days before the measurement
  mailbox's own date. Both machines passed their whole live tier from the frozen checkpoints, 127
  of 127 and 81 of 81.

- ✅ **The build machine runs the test suite in UTC, and refuses to run in any other time zone (Q126).**
  GitHub's UTC runner used to catch time-zone bugs, the way the Q95 one was found, and it went with
  the CI. `OutlookAI-Build` now rests on `CP-03-SDK-TEST-READY-UTC`, the runner refuses a guest in
  any other zone, and the first runs passed 3,818 of 3,818 with no zone bug found. Tests that use
  the machine's own zone as the other one prove less there, so they run again in Central European time (Q144).

- ✅ **The indexed test machine's test mailboxes hold calendar entries, contacts and tasks again.**
  There are four of each in the hub and fourteen in the bystander, as the maintainer decided (D62
  (b)). The search index gives a contact no received date, so the contacts are the undated items the
  search-order checks measure; it dates a calendar entry or a task like mail, so those are now dated
  by the mailbox's own plan, older than every mail item in it, instead of by the moment they were
  built - they can never become the newest item a check reads, and where each one sorts is the same
  on every rebuild. Building one refuses if Outlook will not keep the date written to it, and the
  wait for the search index refuses until the index dates every one where the plan does. The per-run
  hub rebuild now builds the decided set on that machine whatever the previous hub held.

- ✅ **The check that undated items cannot push mail out of a search decides by arithmetic (D74).**
  It used to search for 25 results and rely on the bystander mailbox holding 42 undated items, seven
  more than the 35 spare results such a search fetches - true only while three unrelated numbers
  stayed put. It now counts the undated items a mailbox holds and sizes its search so that they
  always out-number the spare results, runs the same search without the safeguard to see whether the
  safeguard decided the answer, says so, and fails if two otherwise identical searches disagree
  about where undated items sort.

- ✅ **The test machines stay switched off unless they are in use.**
  New ones no longer start with the computer, and are saved to disk rather than shut down when it
  stops, so a restart leaves them using no memory.

- ✅ **The test machines install the add-in in two steps, so Outlook's first start is not elevated (Q100).**
  The installation itself with administrator rights, then Outlook's first start with the add-in -
  the start that proves it runs - without them. The first start used to inherit the installer's
  administrator rights, and an Outlook started that way never adds mail to Windows Search, so on the
  test machine whose mail must stay searchable it broke the rule that Outlook only ever starts as an
  ordinary user. The installation step now never starts Outlook and always reports that the add-in
  has not run yet; the first-start step checks that Outlook really started without administrator
  rights and is the only one that can report the add-in ready; and the old one-step command is
  refused instead of silently doing half the job. Checked on the developer's machine only so far -
  not yet run on a test machine.

- ✅ **The check that a search-index item id cannot be opened directly is split into its two halves (Q74).**
  On an Exchange mailbox such an id is refused, as recorded in 2026-07; on a local mailbox file the
  same id is the item's own id and should open - so the old single check failed on every test
  machine and could only run on the maintainer's own. The Exchange half now runs only where there is
  an Exchange mailbox, and a new half asserts the local-file behaviour on the test machines; that
  half follows from how the id is built and is confirmed or refuted by its first run there. The run
  lines for the test machines exclude Exchange-only checks by a name that is computed, not typed,
  and the build fails if a written copy stops matching it. Nothing that ships to users is affected.

- ✅ **The two delegate-folder search checks move into a group of mailbox checks that writes nothing (Q74).**
  They can only run on the maintainer's own machine, which is now read-only, and the group they were
  in creates and removes test folders when it starts - so on that machine they would have failed
  before they began. What they check is unchanged. Nothing that ships to users is affected.

- ✅ **Mailbox checks on the maintainer's machine refuse every tool not known to be read-only (Q74).**
  That closes the second way such a check could still write there. The checks that talk to a real
  copy of the server over its standard connection now refuse, before anything is sent, every tool
  that is not known to be read-only - drafting, sending, moving, archiving, signatures, saving
  attachments and opening windows - whenever the machine is the read-only one. A tool nobody has
  classified is refused the same way, so a new tool has to be classified on purpose before such a
  check may use it there. Nothing that ships to users is affected.

- ✅ **The mail-loss guard no longer re-runs other mailbox checks on the maintainer's machine (Q74).**
  When it suspects mail went missing it still counts the mailboxes again twice before deciding, but
  it no longer starts a second run of every check that had just run - which ignored the original
  selection, included checks that write, and could never have turned the failure into a pass anyway.
  Nothing that ships to users is affected.

- ✅ **The maintainer's machine is read-only for mailbox checks against a real Outlook, in code (Q74).**
  Until now that was a convention. A machine whose test settings declare it the reference machine -
  or declare nothing, as his do - now refuses every write a check attempts, his designated test
  mailbox included, and the permission that let two checks create and delete a draft in his other
  business mailboxes is gone. Only a dedicated test machine may still write, to its own test
  mailbox. Nothing that ships to users is affected.

- ✅ **The mailbox test suite runs on the test machines without administrator rights.**
  That is how the product itself is built to run. The second test machine also gets a real exclusion
  from the search index. The tool that runs work in a test machine's logged-in session could only
  start it with administrator rights, and an Outlook started that way never uses Windows Search - so
  on the machine meant to have an index, nothing a test created would ever have been found, and the
  suite would have read that as a slow indexer. That tool now takes a run level: administrator
  rights stay the default, because the installers that go through it need them, and the suite now
  runs without them. It warns when an Outlook is already running at the other level, which the work
  could not talk to, and work asked to run without administrator rights refuses to start at all if
  it gets them anyway. Every test and helper the suite runs was checked, and none needs
  administrator rights. The machine meant to have no index was also taken back to its last clean
  saved state and given the exclusion rule that Windows Search itself reports, instead of relying
  only on a policy setting that turned out not to be an exclusion at all; it then verified as not
  indexed, for that reason, and was saved in that state. Nothing that ships to users is affected.

- ✅ **The undated-item checks measure real undated contacts in the indexed test machine's mailboxes.**
  The plan gave the test machines' hub and bystander mailboxes calendar entries, contacts and tasks
  meant to carry no received date - 12 in the hub, 42 in the bystander - so the three checks of how
  the search index sorts undated items against mail can measure something, instead of passing on
  mailboxes that held none. Unsent drafts were left out on purpose: Outlook would first write each
  one into a different mailbox. The first trial on a test machine found that Outlook gives every
  such entry a received date when it is saved into a data file, and refuses to remove it; so they
  are left out for now - the hub and bystander mailboxes are built with their mail alone, 56 and 300
  items - and they come back only if a measurement shows the search index treats those entries as
  undated after all. The measurement found that it does for contacts alone - calendar entries and
  tasks get a received date in the index, contacts get none - so the test machine with a search
  index now carries 12 undated contacts in its hub and 42 in its bystander, and the three checks
  measure them there; the other test machine, where those checks do not run, keeps its mailboxes
  without them.

- ✅ **One script rebuilds the hub mailbox before every test run, and the time-zone check insists on it.**
  On a machine whose hub was not rebuilt, the time-zone check fails and names that script; it used
  to pass without proving anything. The check that the search index reports its newest mail at the
  right time can only catch a time-zone mistake while that mail is under an hour or two old. The
  settings file now names the mailbox the hub is rebuilt from, and the settings script refuses a
  test machine whose settings do not.

- ✅ **The search-index speed checks are timed against the largest indexed mailbox on the machine.**
  They used to take the first one listed. On a test machine the first is a mailbox of a few dozen
  mails, where a two-second limit is met by construction; the indexed test machine's measurement
  mailbox stays at 160,000 items so the limit means what it does on a real one. That mailbox now
  exists: 160,000 items, built into a new data file on the indexed test machine in under an hour and
  a half, and every one of them in its search index. It stays usable for a month between rebuilds
  rather than a week: that machine's tests now check only its 30- and 60-day spans of mail for
  staleness, because none of them asks it about the last seven days. The script that checks a test
  machine's index state now gives an index of that size the time it needs to be counted, instead of
  giving up after 30 seconds and reporting the index as missing.

- ✅ **Mailbox tests that act on a real Outlook profile need a go-ahead for that run and computer.**
  A test selection that merely happened to include one of those tests could reach a real mailbox by
  accident - and nearly did: a run meant for a handful of offline checks picked up three of them,
  stopped only because that copy of the project had no mailbox settings in it. Those tests now
  refuse to start, before touching anything, unless the run sets `OUTLOOKAI_LIVE_OPT_IN` to the name
  of the computer it is running on; a value saved permanently in the environment, or naming another
  computer, is refused as well. The refusal says what the setting is for, how to give it on a test
  machine, and that the maintainer's own computer is read-only for these tests; a new check proves
  from the compiled tests that none of them can start any other way. Nothing that ships to users is
  affected.

- ✅ **A test machine's install disc with a computer name the machine would refuse is not built.**
  The refusal happens on the developer's machine. The first-logon script now checks it is on a test
  machine by its name, and a disc naming anything else produced a machine that stopped partway
  through an unattended install about seven minutes in. Nothing that ships to users is affected.

- ✅ **The last two gaps in how the test machines protect the computers around them are closed.**
  First, four test-machine scripts still changed whatever computer ran them without checking that it
  was a test machine: the one that builds the 20,000-item measurement mailbox, the one that times a
  sweep of it, the one that finishes a freshly installed machine's language, regional and power
  settings, and the one every Outlook step on a test machine runs through, which installs a
  scheduled task. All four now refuse to run anywhere but a test machine, before they touch
  anything - even when only asked what they would do - and the project's automated checks now fail
  the build if any of them loses that check; the list of scripts excused from it is empty. The
  first-logon script travels alone on the installation disc and cannot borrow the shared check, so
  it carries its own, matched to the account and machine names the test machines are built with.
  Each refusal was proven on the maintainer's own computer with every command that could change
  something replaced by a tripwire: every script refused, nothing reached a tripwire, and every
  setting each would have touched read back unchanged. Second, building a test machine from scratch
  left its password on disk in plain text for good: the disc image that installs Windows carries it,
  and the machine's first checkpoint was taken with that disc still in, so the checkpoint depended
  on the file and it could never be deleted. The build now ends with a step that waits for Windows
  to finish setting itself up, takes both discs out, takes that first checkpoint without them, and
  only then deletes the password-carrying image - and refuses to delete it while anything on the
  computer still depends on it. The two existing test machines were built before this, so their
  images stay until those machines are rebuilt, and the image builder now refuses to overwrite one
  by accident. The new finishing step has not yet run against a real machine; the parts that make
  its decisions are tested without one.

- ✅ **Two holes in how the test machines are built, found rehearsing a rebuild from nothing, are closed.**
  First, three of the test-machine setup scripts changed a computer's settings without checking that
  it was a test machine. One of them renames a data file in whatever Outlook profile is open, and
  its own description claimed it checked when it did not, so run by mistake on the maintainer's own
  computer it would have renamed a real mailbox. All three now refuse to run anywhere but a test
  machine. That was proven on the maintainer's computer: each refused, and every setting it would
  have touched read back unchanged. The project's automated checks now fail the build if any
  test-machine script that changes something stops checking first, or checks only after it has
  already changed something; four older scripts that still do not check are named in that check,
  each with its reason, until they are fixed. Second, the test machine's mail profile could not
  actually be rebuilt from the project's scripts. The one setting that makes its test mail account
  usable had only ever been set by hand, and the scripts still defaulted to a version of the profile
  known to leave that account unusable. The scripts now set it themselves and refuse the broken
  version, and the profile part of a rebuild was run twice from a freshly installed Office, once in
  each possible order. That showed the order matters: the profile with the mail account has to be
  created at Outlook's very first start, or its account stays unusable until Outlook is started
  again. The rebuild instructions are corrected to match. The runs also found that the hand-made
  setup had quietly deleted three of the settings meant to keep the old-style Outlook in place, and
  answered an open question for one mailbox test: a plain Outlook data file has no archive folder
  until the product asks for one, and asking creates it.

- ✅ **Three open questions about the test machines are settled by running them, and the notes corrected.**
  The Windows mail component that the profile scripts were abandoned for on these machines' version
  of Office was never broken: the scripts had asked it for an interface under the wrong identifier,
  and asking under the right one works - recorded as knowledge, with nothing rebuilt on it. The
  script that takes a test machine's mail out of Windows Search had one half that could never have
  worked, because Windows does not let even an administrator change the setting it was writing, so
  on the first machine where that setting existed the script would have stopped halfway; it now only
  reads that setting and reports it, and a self-check proves it never writes it again. Exercising it
  also showed that the test machine meant to have its mail INDEXED never has - Outlook never put
  itself into Windows Search there, not even with its window open - so the check now names that
  state instead of telling you to wait for a search index that will never fill, and the earlier
  "unindexed" result on the other machine is now reported as proving less than it seemed. And the
  second mail account the test machine needs, the one that lets two account-specific draft checks
  prove something, is now built by a script with its own data file, where the notes said no free way
  existed; one step writes a setting Microsoft does not document, and the script and the notes say
  exactly which. Two side findings are recorded for the next step: Outlook clears its one-shot
  profile-import setting by itself, and the test machines' out-of-date antivirus makes Outlook stop
  and ask permission when a program reads account details, which the mailbox test run will need an
  answer to.

- ✅ **The test machines' default-profile switcher is replaced, because the way it worked did not.**
  The first time it was ever run it failed outright on the test machine's version of Office - the
  component it asks Windows for exists and answers, but refuses to hand back the interface the
  script needs, and why is still unknown. It now sets the default the plain documented way instead,
  which is measured working on the same machine the same day, and it refuses to name a profile that
  does not exist: that write would have succeeded, and a default pointing at a profile nobody
  created surfaces much later looking like an entirely different fault. Two sibling scripts are on
  the same broken route and are marked as expected to fail until they are fixed.

- ✅ **The measurement-mailbox builder refuses up front when it cannot possibly succeed.**
  It needs two things that nothing checked: the default mail profile must be the one with no mail
  account in it, and Outlook must already be running and warmed up. The only build that ever worked
  satisfied both by accident, which is why it looked repeatable and was not - and why five
  consecutive attempts failed for two different reasons that looked alike. It now checks both before
  touching anything, reports every failing condition rather than the first, and prints the four-step
  recipe that works. It deliberately does not fix either one: quietly rewriting which profile is the
  default would change the machine out from under whatever runs next.

- ✅ **The measurement tools name the mail profile they used when they refuse.**
  They refuse to write into a mailbox they cannot vouch for, which is right, but the refusal named
  only the mailbox - and the tools pick their profile themselves rather than using whichever Outlook
  happens to be open, so the useful fact was the missing one. It cost five failed attempts to work
  that out by elimination. Both the refusal and the success now name the profile and say where it
  came from.

- ✅ **The guard that watches for lost mail names its real reason in every failure message.**
  It counts every folder of every mailbox before and after a test run, and walks the smaller ones
  item by item so it can tell a mail that was FILED from a mail that was DELETED. When it cannot
  afford the item-by-item reading it falls back to counting - and every failure message then blamed
  the same cause, "folder above the identity budget", which is right in one case out of six and
  sends the reader to the wrong remedy in the other five. The reason is now carried with the reading
  and named in the message. It also says, for the first time, when the reading taken AFTER a run was
  weaker than the one taken before it: that is the case where the same untouched mailbox can be
  reported as a loss on one pass and as ordinary filing on the next, and it used to be
  indistinguishable from a real deletion. Which runs fail is unchanged.

- 📝 **The notes record that the test machine needs far less than was thought.**
  The notes said the whole draft, revise, discard and send families were out of reach without a
  working mail server on the machine; read against the tests themselves, one ordinary mail account
  of any kind reaches 26 of those 34 checks, and only one single check anywhere genuinely needs a
  mail to travel and arrive. The note claiming six checks needed real delivery was counting files
  rather than checks - the real figure is 13 - and 12 checks are labelled as needing a mail server
  when they never send anything.

- ♻️ **The test suite says which machine a test needs with two labels, down from three overlapping ones.**
  A test now carries one flag meaning "this needs a mailbox" and one short list of what it needs of
  that machine, written on the test itself rather than on the file it sits in. The third label was a
  summary of the other two, kept up to date by hand, and it disagreed with them: it reported that 96
  of the mailbox tests could never run anywhere but the maintainer's own PC. Read properly, the real
  figure is six - the ones that need a shared mailbox belonging to somebody else, which no test
  machine can be given. Nothing about any test changed; only what it says about itself.

- ✅ **The live test suite runs under the same time limits as the product.**
  Those tests ran against a code path where none of the limits existed, which is both a coverage
  hole and the likely reason a live test run once hung with nothing to stop it.

- ✅ **The refused-sort report is confirmed against a purpose-built 20,000-item mailbox.**
  When Outlook refuses to sort a folder, the check for new mail correctly reports that the mail it
  could not cover is unknown rather than claiming the oldest mail is what is missing.

- ♻️ **The tuning and mail-server status in `outlook_health` is harder to get wrong.**
  The add-in and the server now share one definition of every registry key and configuration name
  they exchange, so renaming one can no longer leave the report quietly describing something that no
  longer exists.

- ✅ **The test-machine scripts no longer guess which machine you mean.**
  Three test guests exist side by side while the old one is replaced, and the scripts that build,
  copy into and reach into a guest each quietly assumed one of them - so a command aimed at a new
  guest was refused before it started, by a stored credential still pinned to the machine being
  retired. Naming the guest is now required everywhere, the stored credential no longer pins itself
  to one machine, and the recorded name in the parameter file is labelled as the guest the published
  measurements were taken on rather than a guest to build. The one script that sweeps every guest to
  reclaim the host's memory keeps its list of all three, because sweeping all of them is the point.

- 📝 **The rebuild notes warn that leaving the identity checks out also deletes the record of the gap.**
  Those checks can now be excluded by name, and on a machine with no second mail account the line
  they print - saying in writing that nothing about that path was verified there - is the one place
  the gap is written down anywhere. The rebuild notes now say this beside the command they would be
  excluded from, and say that the label exists for machines that have the account rather than as a
  way to quieten one that does not.

- 📝 **The build instructions give the test machine a second mail account now, not later.**
  The machines are about to be created from scratch, where it is one more account and one more data
  file; adding it afterwards means editing a profile, restarting Outlook and changing a settings
  file on a machine whose saved snapshots no longer describe it. The instructions now say what the
  account has to be, read off the two checks that need it rather than guessed: a real account and
  not a bare data file, named after its own address, delivering into its own data file, with a
  signature configured, and deliberately left off the list of accounts nothing may write to.

- 📝 **The notes record that the test machine runs a newer Office than the maintainer's, by choice.**
  It is about 3,600 builds newer, and that is accepted, not fixed. The test machines otherwise copy
  the maintainer's setup deliberately, so this is an exception and is now written down as one,
  together with its practical consequence: a difference in Outlook's version is the first thing to
  suspect when a test behaves differently on the test machine than on the maintainer's. Also
  recorded: the monthly rebuild the test machine needs comes from how that one machine was licensed,
  and is not something anyone using this product experiences.

### Removed

- 🗑️ **The add-in no longer runs a Claude Code session every time Outlook starts (Q182).**
  Its warm-up ran `claude -p` from inside OUTLOOK.EXE at each Outlook start, the starts OutlookAI's
  server makes for an agent included, in system32 when Outlook was started over COM, and the process
  did not stay warm: on 2026-11-15 it sent the prompt "-" to the model and exited 3.3 s later. The
  first sidebar action after Outlook starts now starts its own process, and every action still
  warms one up for the next. Decided by the maintainer on 2026-11-18 (Q182 (a)).

- 🗑️ **`move_mail` no longer creates folders, and no tool takes a folder path or a store name (Q134, Q135).**
  `move_mail`'s `create_folder` and `store` arguments are gone: a folder is made with
  `create_folder`, and the target folder's ID names its store. Its target is checked once, before
  any mail moves, so a move into Deleted Items or the Outbox is refused with nothing moved; a move by
  path used to make the missing folders before that check, and could leave a new folder in Deleted
  Items. The `store` and `folder` arguments of `search`, `thread`, `list_folders`, `goto_folder` and
  `show_search_results` are gone too, replaced by `store_id` and `folder_id`. A call that still
  passes one is refused, naming what replaced it.

- 🗑️ **The operator console keeps its corpus commands and loses the July 2026 incident ones (Q170).**
  The maintainer closed that incident, so `OutlookAI.RemediationTools` no longer has `audit`,
  `refile`, `purge` and `dedupe`, nor the decision rules, mailbox helper and MCP client only they
  used, nor their T1 test. What the test machines use stays: the corpus and fixture-population
  commands, their COM session and their STA runner, all of which still build. The corpus tests
  that proved the corpus tag apart from the artifact tag now read the live tier's own tag and
  sweep fragment, the definitions the sweep matches with, instead of the incident tool's copy.

### Fixed

- 🐛 **A `\u` escape or a numeric HTML entity with a space or a sign in it stays text.**
  The add-in's reader of the values in `~/.claude.json` parsed a `\u` escape's four characters with
  .NET's hex number parsing, which allows spaces around the digits, so `\u 41 ` read as `A`, though
  JSON has no such escape. OutlookAI's reader of HTML as text, which gives a mail's body and a
  signature's text, parsed a numeric entity's digits the same way, and a decimal one's with a sign
  allowed too, so `&# 65;`, `&#+65;` and `&#x 41;` each read as `A`, where a browser shows them as
  written. Both read hexadecimal digits through Core's `Hex` now, and a decimal entity's digits as
  the ten digits alone, as the check that guards every write to `~/.claude.json` always read an
  escape; anything else keeps each reader's fallback, the text as it stands. The code-sharing
  refactor's hex stage found them, and T1 theories pin both.

- 🐛 **The task pane finds a Claude Code installed while Outlook runs, without a restart.**
  The pane remembered that Claude Code was missing when Outlook started, and every request then
  failed with "not found" until Outlook was restarted, however soon Claude Code was installed. When
  that is the error it remembers, it now looks for the program again before each request, and goes
  ahead once it is there. Found while building the live pane test (Q168 (b)), whose stand-in for the
  CLI arrives after the test machine's Outlook has started.

- 🐛 **Another user's Outlook on the same computer is no longer taken for yours.**
  OutlookAI counted every OUTLOOK.EXE on the computer. With someone else signed in to it, by fast
  user switching or on a terminal server, their Outlook, whose windows are on another desktop,
  looked like yours still starting, and health reported their process as yours. Only an OUTLOOK.EXE
  in your own Windows session counts now: in the checks before a mailbox is opened, in health, in
  the watch for Outlook exiting and in the count of its windows. Decided on 2026-11-18.

- 🐛 **An Outlook that closes during a call is reported as a COM failure, with advice and an outcome.**
  When Outlook went away under a call and the mail server's exit watcher had released the session
  first, the agent was told "Cannot access a disposed object", with no advice and no outcome. That
  shape, and a COM object already separated from the session, are now reported exactly as the COM
  disconnect they are: the `ComFailure` error, the advice for the call that was interrupted, and
  whether it may have changed anything, for a read and a write alike. A read meets this more often
  since its retry no longer starts Outlook again (Q179). Decided on 2026-11-18.

- 🐛 **A request whose caller gave up no longer leaves a stuck call holding the COM host (Q179).**
  The cancel path disposed the request's deadline, against the supervisor's own description, so a
  call stuck inside Outlook that nobody awaited any more held the COM host for good, and the host
  serves one call at a time. On 2026-11-15 such a request stayed stuck for twelve minutes. Its
  deadline now still ends the COM host. `ComHostPolicy.DecideInFlight`, which put a cancellation
  above everything and was called by nothing, now says a cancellation stops the waiting and not the
  deadline, and the deadline watchdog acts on it (Q179 (b)).

- 🐛 **A retry after Outlook went away attaches to Outlook and never starts it again (Q179).**
  When Outlook went away under a read, the COM host rebuilt its connection and ran the read once
  more, and that connection could start Outlook. On 2026-11-15 it started Outlook 67 ms after the
  hung one was ended, for a request abandoned twelve minutes before. The rebuild now only attaches;
  with no Outlook running the read fails with the disconnect that ended it, and the next request
  starts Outlook if it is still wanted (Q179 (c), Q150 (a)).

- 🐛 **A paged exhaustive search of an Exchange mailbox returns no mail twice after Outlook restarts.**
  An exhaustive search hands out its results a page at a time, and each page remembers which mail
  the pages before it returned, so that none comes back. It remembered the ids the folder's table
  gave, and on an Exchange mailbox Outlook keeps locally those are short-term ids, valid only until
  Outlook restarts: a page after a restart could not recognise what earlier pages had returned.
  Measured read-only on the Exchange test machine on 2026-11-18: the table's ids were 24 bytes, the
  same in two tables of one session, and the column PR_LONGTERM_ENTRYID_FROM_TABLE gave each row's
  70-byte long-term id, equal to the mail's own id in 25 of 25 rows. The scan reads that column and
  keys its memory on it; a .pst data file's table has no such column, and its ids are long-term
  already. Found by stream E2's Exchange runs.

- 🐛 **OutlookAI's hidden window is no longer left behind, keeping Outlook running (Q118 (a)).**
  The hidden window that keeps an Outlook with no window running while agents draft was closed only
  by the session that had started Outlook. So it was left behind by a session that connected while
  another was starting Outlook, or whose process was killed, and a window left behind kept Outlook
  running past Application.Quit() and past closing a window opened over it (measured 2026-10-03);
  and a starter that ended first closed it under the others. Every session now joins a list kept in a
  file under the OutlookAI folder and changed only under a shared lock: Outlook has one such window,
  recorded by its window handle and shared by every session; the last session to end closes it,
  whoever made it; a session that hears Outlook quit closes it, so your exit wins for every session;
  and a session whose process died is dropped by the next one, which takes its window over. Only a
  recorded window that is still hidden is ever closed. `outlook_health` lists the sessions attached
  to Outlook, whether OutlookAI started it, the hidden window and who made it, and what will end
  Outlook, and calls a window that no session holds any more a problem, with File > Exit as the way
  out.

- 🐛 **A reply's or forward's text and signature sit above the quoted message, not under its line.**
  `Reply()` and `Forward()` leave two empty paragraphs above the quoted message's header block for
  the new text, and the block's first paragraph carries the separator line as its top border.
  OutlookAI put the signature, and then the text, at the start of that bordered paragraph, so Word
  laid them out under the line, inside the header block: all 12 replies and forwards Q149 measured on
  2026-11-17, the first paragraph in Times New Roman. As decided (Q209 (a)), the text and the
  signature now go into the empty paragraphs above the block, and the last of them stays as the blank
  line before it; with text right above the block, an empty paragraph is split off it first. A
  revision of such a draft replaces only what is above the block. Nothing is written into the quoted
  message any more, so the marker that bounds it no longer swallows the new text either.

- 🐛 **The health report says when Claude Code's `~/.claude.json` cannot be read whole.**
  `outlook_health` called the MCP registration `absent` when the file was there but read back
  empty, which can happen in the moment Claude Code rewrites it. A file cut off part-way, or with
  more after its closing brace, could be called `absent` too, or have its command read out of the
  part before the break and be reported `ok` or `drifted`. Both are `unreadable` now: the server
  reads the file with the add-in's code, by the rule the add-in applies before it writes there, and
  only a file that is not there at all is `absent`. T1 `McpRegistrationHealthTests` pins the rule.

- 🐛 **A move names a folder whose name holds `/`, `\` or `%` the way list_folders does.**
  `move_mail` and `archive_mail` printed the folders an item left and went to as Outlook writes them,
  with those three characters coded: a move into the folder `a/b` that `create_folder` had just made
  answered `Parent/a%2Fb`, while `list_folders` and `find_folder` show `Parent/a/b`. Outlook codes
  only those three and codes `%` itself, so undoing them is exact (measured on the indexed test
  machine, runbook section 8 item 28), and a move now prints a folder as those two do.

- 🐛 **A search of one folder whose name ends with a space finds that folder's mail in the index.**
  Outlook allows such a name and the index keeps the space, in the folder's address and in its path,
  but OutlookAI trimmed it, so a search of that folder alone found nothing in the index. The space
  is kept now, and a folder whose name ends with one is narrowed through its parent, the one way the
  index matches it. Measured on the indexed test machine in the follow-up to the folder-name work
  of October (branch `q99-name-encoding-followup`), carried over here.

- 🐛 **A search never takes the stale copy the index keeps of a rebuilt mailbox for the mailbox.**
  After an Exchange offline file (.ost) is rebuilt, Windows Search keeps the old copy of the mailbox
  beside the new one, under the same name and with the same folders; measured on the Exchange test
  machine in November. Where nothing ties the mailbox to one of them by its index number, a search
  scoped to it took the first of the two, which could be the copy that no longer changes. It is
  refused now, naming both, with what to do: read the mailbox with `exhaustive`, or rebuild the
  Windows Search index. The same holds for a shared mailbox found under its owner's store. A search
  of every mailbox read both copies and said nothing. Where the index ties the mailbox to one copy
  by its hash, the other is now left out of it, and the advice says which and why
  (`index.storeCopiesLeftOut`); a data file's leftover from an earlier path and another Outlook
  profile's store of the same name are left out the same way. Where nothing ties it, both are still
  read, and the advice says the results can hold mail moved or deleted since, or the same mail twice
  (`index.storeCopiesSearched`). `outlook_health` lists such copies with their store
  (`storeCopies`), and counts the second case as a problem. Decided 2026-11-18.

- 🐛 **Mail in a folder with a namesake beside it opens, and a search of it says what it can include.**
  Two folders of the same name can sit side by side, as Exchange's two top-level "Conversation
  Action Settings" folders do on the Exchange test machine, and the search index keeps one path for
  both, so a search of one through the index can return the other's mail. Outlook now reports such
  a folder when its ID is looked up, and the search says so in its advice and with
  `scope.pathShared`; `find_folder` lists both folders, each with its own ID, and says they share a
  path. Opening a search result from such a folder (read, thread, save_attachment, open_in_outlook)
  looked for it in the first folder of that name only, so mail in the second could not be opened.
  When the first does not hold it, every other folder at the path is tried now, read by its own
  name, which also finds a folder whose name holds `/`. Decided 2026-11-18.

- 🐛 **A draft's audit line says Outlook's editor was promoted only when Outlook had no window.**
  Since the compose crash fix every draft is written through an activated inspector, and the audit
  line's composeSurface read wordEditorPromoted for all of them, so it said nothing. It now reads
  wordEditorPromoted only when no Outlook window was on screen as the editor was promoted, and
  wordEditor otherwise. A refused editor keeps its cause as well: NoWordEditor and ActivateFailed carry
  the failure's type and HRESULT after a colon - in the audit line, on the wire and in update_draft's
  compose_surface_unavailable refusal - where the bare word made every refusal look the same.

- 🐛 **A draft written without Word's editor no longer advises that opening a window fixes it for good.**
  The advice said opening any Outlook window makes the full compose path available again. Measured
  on Office LTSC 2021 (D48), the editor came back with a window and was lost again when that window
  closed. It now says a window helps only while it stays open, and points at composeSurfaceError,
  which carries Outlook's own error code, and at outlook_health.

- 🐛 **Signatures are read from and written to the folder Outlook uses.**
  Outlook records the folder it keeps signatures in, which an Outlook in another language can name in
  that language and an administrator can move. OutlookAI always used %APPDATA%\Microsoft\Signatures
  instead, in the mail server, in the add-in's signature button and in a test, so on such a computer
  it found no signatures and wrote new ones where Outlook never looks. All three now read the folder
  Outlook records, each time they need it.

- 🐛 **Changing a signature keeps its pictures and its other versions.**
  An update deleted the signature's picture folder and rebuilt its RTF version from the plain text, so
  editing a signature Outlook had made with a logo lost the logo, and Rich Text mail lost its
  formatting. An update now touches only the files it is asked to change and converts nothing from
  one version to another. Because Outlook puts the HTML version in HTML mail, the RTF version in Rich
  Text mail and the text version in plain-text mail, an update that replaces a version whole must
  send every version the signature has, and one that leaves a version out is refused with the
  missing ones named (Q205 (d), decided by the maintainer 2026-11-18). Until then such an update wrote
  the one version and said in its answer which versions still read as before. The test machines'
  `Set-AccountSignature.ps1` takes an RTF body beside the HTML one for the same reason: run again on
  a signature it made, it has to send all three.

- 🐛 **A draft whose signature could not be added says it has none.**
  Since the crash fix, a draft whose chosen signature could not be put in came out with no signature
  at all, while the tool descriptions said the account's default would stay. The answer now says the
  draft has no signature and how to retry, and the descriptions say the same.

- 🐛 **An Exchange account added under another of its addresses gets its default signature.**
  OutlookAI found an account's defaults by its email address, and an Exchange account added under an
  alias is filed in the profile under the alias, so its drafts got no default signature and nothing
  said why. The account is now also found by the name Outlook shows it under, and a draft whose
  account cannot be found says so.

- 🐛 **Writing a new draft, reply or forward no longer crashes Outlook now and then.**
  OutlookAI wrote each draft in a hidden compose window opened on the item Outlook had just created,
  and closed that window when it was done. Closing it crashed Outlook, inside Outlook itself, about once
  in every 40 drafts written back to back, and in two of fourteen runs of the draft tests on the test
  machines. OutlookAI now saves the new item first, opens it again as a saved draft and writes into
  that, the way it already revised existing drafts, which never crashed. 2,400 drafts written back to
  back on the two test machines then ran without one crash. Outlook adds an account's default signature only in the window that crashed,
  so OutlookAI now adds it itself: the signature the account is set to use for new messages, or for
  replies and forwards, in the profile Outlook is running - the settings list_signatures shows. A
  signature you name for the draft takes its place.

- 🐛 **The mail server no longer uses CPU when files change in a Claude Code project folder.**
  Claude Code starts its own copy of the server for every session, in the session's project folder,
  and each copy watched that folder and everything under it for a settings file the server never
  reads. Every file created, changed or deleted there cost each copy some CPU, and every renamed
  folder a walk of everything in it: on the maintainer's machine one copy used up to 0.9 of a core
  and 3,000 to 4,900 CPU-seconds, and all of them together 2.39 cores in a two-minute trace while
  the machine lagged. The server now keeps to its own folder and watches no file at all. Creating,
  changing and deleting 1,000 files and renaming a 1,000-file folder ten times cost the old server
  3.5 to 4.0 CPU-seconds and costs the new one nothing Windows can measure. The tests now start the
  server in a scratch folder the way Claude Code does, and fail if file activity there costs it
  CPU. A session already running keeps its old server until OutlookAI is updated and the session
  is restarted.

- 🐛 **A search of a mailbox or folder no longer returns the folders themselves as results.**
  Where Windows Search indexes the mailbox, a search with no search words listed one "hit" for every
  folder - Calendar, Quick Step Settings, the mailbox's top level, emptied folders in Deleted
  Items - beside the mail, and none of them could be opened. On the test machine they were 21 of the
  first 100 results.

- ✅ **A test machine's lease is never taken or released while its holder still has it.**
  `Testbed/host/Set-TestbedLease.ps1` wrote over whatever lease was there, and its `-Release`
  deleted whatever lease was there, so a person who leased a machine a run was using took it from
  the run without a word, and a live-tier run then stopped at its next renewal, which found the lease
  somebody else's; the load investigation of 2026-11-20 found it. The script now refuses both, and
  says whose lease it is, which script and process took it and until when. It still takes or
  releases a lease that has expired, one with the caller's own `-Reason`, which is how the runners
  tell their own lease, and one whose process has ended - a stopped run, a host that restarted -
  judged as the queue's tickets are, by the process id and the second it started. Every script that
  releases a lease now gives the `-Reason` it took it with. The runners judge a lease the same way,
  through one function, `Get-LeaseDecision` in `Testbed/guest/ScriptHelpers.ps1`: a lease whose
  process has ended no longer holds the live-tier runs, the build runs, the trust-removal probe or the
  clone maker off until it expires, so a host restart leaves no run waiting out a dead run's lease. A
  lease taken by hand, from a prompt, has no process that lasts as long as its holder, so only its
  expiry or a release with its `-Reason` ends it; a lease written before this change is judged by
  its process alone. By hand, the read and the write are one step under the lease's claim, as the
  runners' are. The script's `-SelfTest` holds each case on a scratch directory with two hidden helper
  processes and checks that every release a script makes gives its `-Reason`; the runners' and the
  probe's self-tests pin how they judge a dead holder's lease. On this workstation a lease for a
  machine that does not exist was refused, taken over and released as above under both PowerShells.

- ✅ **A build-machine run cut off by a host restart gets a line in runs.log from the next run.**
  The restart of 2026-11-18 at 20:34 stopped run `20261118-203146-3ca41927bf3c` and its janitor
  together, and left its start line in `runs.log` with no end and its owner file naming a process
  that was gone. A run that takes a build machine's lock now reads the owner file the run before it
  left: when that run's process has ended and `runs.log` holds no end line for it, it writes a
  `gone` line for it - not an end line, so a janitor still watching that run puts the machine back
  as before. A queued run and a janitor waiting for a lock no longer name such a run as its holder,
  and say "another run". Owner files record the second their process started, as queue tickets do,
  so a process id Windows has given to another process since does not pass for the run. The build
  runner's `-SelfTest` pins the line, the names, and the run's and the janitor's owner files.

- ✅ **The guests' sweep measurement names its store and folder by ID, as the server asks.**
  `Testbed/guest/Invoke-GuestMeasure.ps1` still passed the store's display name and the one-folder
  scan's folder name, which `list_folders` and `search` refuse since stores and folders became
  addressed by ID, so its folder inventory and both exhaustive scans measured a refusal. It now
  looks the store up through `list_accounts`, by its display name, and the folder through
  `list_folders`, by its path in that store, and passes `store_id` and `folder_id`. A name that
  matches no store or folder, or more than one, is written down, and the scans that need it are not
  run.

- ✅ **Two package checks started in the same second get an output folder each.**
  `Tools/Test-VulnerablePackages.ps1`, the release's check for NuGet packages with a known
  vulnerability, named its output folder for the second it started and made it with
  `New-Item -Force`, which hands a second check started in that second the first one's folder and
  logs: the bug the test runners' results folders had until 2026-11-18. A check that names no
  `-OutDir` now claims its folder through `New-RunDirectory` in `Testbed/guest/ScriptHelpers.ps1`,
  so a taken name gets a number on the end, 2 and up, until one is free; a release names its own
  folder and is unchanged. The script's self-test claims one second twice and gets two folders, and
  checks that the run takes its folder this way.

- ✅ **The guest scripts that drive the MCP server start it again, with its COM host beside it.**
  Since the move to .NET 11 (19d9993) none of them could: `Set-AccountSignature.ps1`,
  `Set-SignatureSettings.ps1` (the live runner's `-SignatureSettings` step, whose tests printed
  PROVED NOTHING in every run since), `Invoke-GuestMeasure.ps1` and `Measure-PinWindows.ps1` each
  found a server that exited at once, saying "The library 'hostpolicy.dll' required to execute the
  application was not found in 'C:\Program Files\dotnet\'". `McpServer/Directory.Build.props` gives
  the server's published apphost the search the installed product needs, its private runtime and
  then the machine-wide one, and `Testbed/host/Publish-GuestPayload.ps1` publishes the guests' copy
  self-contained, with its runtime beside it: under that search its apphost took the machine-wide
  .NET, which cannot run a self-contained app. The props now set that search only when the server is
  not published self-contained, so the guests' copy keeps the default search, which looks beside the
  apphost first. A search property on the publish's command line does not do: measured on
  2026-11-20, `-p:AppHostDotNetSearch=AppLocal` started the server and lost its COM host, and
  `-p:AppHostRelativeDotNet` alone did the same. A property given there reaches the COM host's
  project too, and the SDK leaves a referenced executable whose apphost has search options out of
  the publish, so the guests' copy had no `OutlookAI.ComHost.exe` and `list_folders` on
  OutlookAI-Unindexed3 answered that it was missing. `Tools/Checks/check-pinned-constants.ps1` holds
  the condition and refuses either property on a script's command line. The installed product is
  unchanged: published with `Testbed/host/Publish-AddInPayload.ps1`'s arguments, the server's
  apphost still looks in the private runtime and then the machine-wide one, and
  `OutlookAI.ComHost.exe` still ships beside it.

- ✅ **Two test runs started in the same second get a results folder each.**
  A run's folder is named for the second it started, its commit and the kind of machine it ran on, and
  the live runner, the build-machine runner and the trust-removal probe each made it with
  `New-Item -Force`, which hands a second run the first one's folder without a word. Stream K2 had two
  runs on the indexed machines start at 12:14:18 on 2026-11-18: they wrote one log, one run's
  `git worktree add` failed, and the other's build failed when the first deleted the tree. All three
  runners now claim the folder through `New-RunDirectory` in `Testbed/guest/ScriptHelpers.ps1`, which
  makes it under a name nobody else uses and renames it to the run's name, a rename that fails when
  the name is taken; a taken name gets a number on the end, 2 and up, until one is free. The run goes
  by that name from then on: its build tree, its folder on the guest, its lease, its queue ticket and
  its summary, and the build-machine runner's janitor accepts it. The build-machine runner's
  self-test pins the helper, two processes racing to claim one name 25 times each among its checks,
  and each runner's self-test that it claims its folder this way before anything is named after the
  run. `Tools/Switch-AddInBuild.ps1`, which names its build folder and the dev build it puts in
  Outlook the same way, claims both with the same helper: two builds of one commit started in the
  same second get a folder each, and a dev build's folder that another run made first refuses the
  second, where it used to be copied into.

- ✅ **The Exchange machine's preflight says when Outlook has refused an old .ost file.**
  On 2026-11-18 the preflight stopped two live runs on the Exchange test machine: the real mailbox was
  no longer signed in, it said, and it asked for `-Mode Reauth`, which then found no sign-in page to
  answer. The machine had been restored to a checkpoint outside the runner with its .ost files left in
  place, and at Outlook's next start Office raised "Outlook is using an old copy of your Outlook data
  file (.ost)" (OAlerts event 300). From then on Outlook held that mailbox disconnected
  (ExchangeConnectionMode 400), and the Windows AAD log held no token error at all. The preflight now
  reads Office's alerts since the running Outlook started, and for an account at 400 after that alert
  it says `OLD-OST` and names the remedy instead: the live runner's `-Checkpoint` path, or
  `guest/Reset-OutlookOfflineStore.ps1`, each of which sets both .ost files aside. Its `-SelfTest` pins
  the decision with that day's alert text and no real address. `Testbed/README.md` says a restore
  outside the runner must set both .ost files aside as well: a restored .ost is older than the
  server, and Outlook refuses it.

- ✅ **A build-machine run that has to stage packages checks the known personal data again.**
  When the build machine's offline feed lacked a package a revision needs, the runner staged it and
  ran the machine's half a second time. That second run expanded the source afresh, after the first
  had moved the private list of personal data into it, so the known identifiers went unchecked and
  the run's summary said so. Every run of the move to .NET 11 stages its 11.0.0 packages, so it hit
  every one of them (Q172, 2026-11-18). The runner stages the list again before the second run, and
  its self-test checks that it does; the next run that staged packages checked the list on its second
  run.

- ✅ **A queued test run no longer stops when another run's ticket goes as it reads the queue.**
  Stream M's build run 20261118-092047 ended as not tested before it started: "An object at the
  specified path ...\queue\<ticket>.json does not exist". Two runs had found the same ended run's
  ticket and both deleted it, and PowerShell's Remove-Item throws for a file that goes between its own
  look and its delete - and under Windows PowerShell 5.1 says "Access is denied" for one already
  pending delete - whatever -ErrorAction says. On a scratch queue whose tickets came and went, it threw
  in 183 of 600 turns under PowerShell 7.7 and in 150 of 600 under 5.1. The queue, the gate's waiting
  records, its pause switch and leaving the queue now list file names alone, count a file they cannot
  read as gone, and delete through one helper that never throws: the same 600 turns threw in none. A
  run whose own ticket has gone writes it again, and a write that cannot be made at that moment - the
  old file still pending delete, which Windows PowerShell 5.1 answered with "Access is denied" in 5 of
  3,000 turns - waits for the next turn instead of stopping the run. The self-tests pin a ticket
  deleted between the list and the read, one ticket deleted twice, a run's own ticket that cannot be
  written again, and a queue whose tickets come and go while 100 turns read it, as other runs join,
  leave and delete the tickets of runs that ended. That churn used to delete the run's own ticket too,
  which no run does, and it failed 1 build-VM run in 100. It now prints the exception of any turn that
  throws, and it threw in none of 1,000 turns under either PowerShell.

- ✅ **The disconnect test no longer fails when Outlook quits just before its pin is released.**
  Closing Outlook's last window can make it quit, so the test releases the lifetime pin only on a
  session that is still connected. When Outlook quit between that check and the call, the test
  accepted a COM disconnect but not a session the mail server's exit watcher had already disposed,
  and on 2026-11-18 it failed that way on the Exchange test machine. It now accepts every shape the
  mail server itself counts as that disconnect.

- ✅ **The build machine's run summaries never print a duration as "1 m 60 s"; seconds round down.**
  `Testbed/host/Invoke-TestsOnBuildVm.ps1` rounded the seconds of a duration to the nearest whole
  one after taking the minutes, so 119.6 s read "1 m 60 s" in `summary.txt`, the run log and the
  history line; under a minute it rounded too, so 59.6 s read "60 s". The live-tier runner's copy of
  the same function rounded down. One copy is left, in `Testbed/guest/ScriptHelpers.ps1` with the
  other helpers the scripts share, rounding down at every step, and the build-VM runner uses it; its
  self-test now pins 59.6 s and 119.6 s. The live-tier runner keeps its copy until its own helpers
  fold (S7d), which `Tools/Checks/check-powershell-51.ps1` check 4 lists.

- 📝 **The draft tools' descriptions say which drafts the tools can reach.**
  `update_draft` can revise any unsent draft in a Drafts folder, including one you wrote yourself,
  and a draft it revised can then be discarded by the same session. The descriptions used to say
  OutlookAI could only ever change or delete drafts it had made itself.

- 🐛 **Outlook tuning no longer stops halfway where Outlook runs without administrator rights.**
  That is how almost everyone runs it. Five of the Cached Mode settings OutlookAI kept applied were
  Outlook *policy* settings, which Windows lets you read but only an administrator change, and the
  first of them stopped the whole run: the Cached Mode and mailbox-size settings after it were never
  applied, and the mail server's health report said the tuning had never run. OutlookAI now writes
  those settings as your own instead, and a setting Windows refuses for any reason is skipped while
  everything else is applied, the run is always recorded, and the health report names what was
  refused.

- 🐛 **OutlookAI releases the Outlook objects it adds as soon as it has added them.**
  They are a column it adds to one of Outlook's tables while it searches, and a bookmark it places
  in a draft while it writes. They used to be left for .NET's garbage collector, which releases them
  at a moment nobody chooses: after the table or the draft they belong to has already gone, inside
  Outlook. The project's own mailbox checks did the same with the attachments they add and the
  folder collections they count. Found while chasing Outlook crashes on a test machine.

- 🐛 **On Office LTSC 2024, showing you a folder no longer reuses OutlookAI's invisible window.**
  The same goes for search results. That invisible window is the one OutlookAI keeps open so that
  Outlook stays running for it. Asked for a new window on the same folder - your Inbox - Office 2024
  hands that invisible window back instead of making a new one, so OutlookAI showed it, and closing
  it took away what kept Outlook running. OutlookAI now always opens a window of its own and moves
  it to the folder you asked for. Closing it still ends Outlook when OutlookAI had started Outlook
  itself: Office 2024 treats closing its last window as quitting, and OutlookAI lets it quit - as it
  must, or your own Exit could not close an Outlook that OutlookAI started - and starts it again,
  without a window, the next time it needs it. Measured on a test machine with Office LTSC 2024; on
  Office builds that make a new window every time, nothing changes.

- 🐛 **A draft OutlookAI saved in a data file with no Inbox can be discarded or revised.**
  Replying to a message in such a data file makes Outlook create a Drafts folder there, but Outlook
  records that folder only in the data file's hidden top folder, which OutlookAI did not read, so it
  then answered that the data file had no Drafts folder and refused to discard or revise the draft
  it had just saved. Found on a test machine.

- 🐛 **Discarding a draft in a .pst data file reports the id of the draft it just moved to Deleted Items.**
  In a local data file, such as a POP3 account's, it used to report no id, or the id of an older
  discarded draft with the same subject. A data file keeps a message's id when the message is
  deleted, and OutlookAI only looked for a new one, as an Exchange mailbox hands out; on any mailbox
  that is not Exchange it now first checks whether the old id opens in Deleted Items.

- 🐛 **A complete search handed out page by page no longer repeats or skips mail.**
  It could repeat a message on a later page, or, on a computer west of UTC, skip mail and still call
  itself complete. Outlook reports a message's received time in local time under the column name
  OutlookAI asks for first, and in UTC only under a second spelling of the same column; OutlookAI
  read both as UTC, so where each page stopped was off by the computer's distance from UTC. Measured
  on the first full run on a test machine, which settles an open question about the zone those
  columns report in (Q11).

- 🐛 **The Archive folder Outlook makes in a .pst data file is found without asking Outlook for it.**
  When mail is archived in a POP3, IMAP or data-file mailbox that has no Archive folder, Outlook
  creates one and records it in a place OutlookAI did not read, so OutlookAI's read-only lookup kept
  answering "no Archive folder" about the folder the mail had just been archived into. The first
  full run on a test machine found it: the mailbox checks that look for the Archive folder that way
  expected a new one on every archive, and their clean-up never reached the test mail inside it. The
  place Outlook uses is not in Microsoft's documentation; it was measured on that machine and is
  read only on non-Exchange mailboxes.

- 🐛 **A search of one folder finds its mail when the folder's name contains `%`, `*`, `?` or `\`.**
  Until now it found nothing in the Windows search index (Q99 finding). Windows writes those
  characters - and `/` - as codes in the index's address for a folder, so a folder called `50% off`
  is filed as `50%25 off`, and OutlookAI looked for the name as written: the index answered nothing,
  and only the last few days' mail came back. OutlookAI now spells folder names the way the index
  does, measured on a test machine for all five characters. Results from such folders now also show
  the folder's real name instead of the coded one, open directly, and are no longer listed twice
  when the index and OutlookAI's check of recent mail both find them. A folder whose name contains
  `/` can still only be searched through the folder above it, because folder paths use `/` to
  separate folders.

- 🐛 **A discard that failed before the delete was attempted says the draft was NOT deleted.**
  That is a failure while Outlook was still reading the draft and its folder. It used to say that
  whether the draft was deleted is unknown, and send you to look in Deleted Items. Only a failure
  during the delete itself still says the outcome is unknown.

- 🐛 **A draft that cannot have been saved is no longer reported as possibly saved.**
  A new draft whose Drafts folder could not be opened, and a new draft, reply or forward that failed
  before Outlook was given anything to save, now say that no draft was created and that retrying
  cannot leave a second one, instead of sending you to look for one in Drafts.

- 🐛 **The audit log no longer loses lines when two sessions write at once, and it cannot be garbled.**
  Every session runs its own copy of the mail server, and when two of them added a line at the same
  instant, one line could overwrite the other without any error - measured, two copies writing 3,000
  lines each at once kept 5,883 of the 6,000. Every line is now added in one write that Windows
  itself places at the end of the file, so no copy can overwrite another's line, even one that
  cannot wait its turn; the copies still take turns, so the log stays in time order. Each line also
  records which process wrote it, that writer's own line number and a checksum, and reading the log
  back reports any line that was damaged and any line missing by those numbers, instead of answering
  as if the log were complete. A line left unfinished by a crash no longer swallows the next one,
  and outlook_health counts how often this server had to write without waiting its turn.

- 🐛 **A default signature lands on the right account, and is read back before success is reported.**
  OutlookAI used to put it on the wrong entry of the Outlook profile and tell you it had worked.
  Outlook names a POP3 or IMAP account's data file after its email address by default, and making a
  signature the default for that address could write it onto the data file instead of the account:
  the account was left with no default signature, yet both the tool's answer and the signature list
  reported it as set, because the check that read it back looked in the same wrong place. It was
  found on a test machine whose account had a name of its own. A default is now written only onto
  the mail account whose address matches exactly - never a data file or address book, whatever it is
  called - it is refused rather than guessed when no account or more than one has that address, and
  it is read back from that account before success is reported. The signature list now shows each
  account's own defaults, including accounts whose name in Outlook is not their address, which it
  used to leave out. A default an earlier version put on a data file is left where it is; setting
  the default again puts it on the account.

- 🐛 **OutlookAI no longer adds folders you never asked for to a POP3, IMAP or .pst mailbox.**
  Finding out where such a mailbox keeps its Archive folder, and checking that it was the right
  folder, asked Outlook for the mailbox's standard folders in a way that makes Outlook create any
  that are missing: measured on a POP3 test mailbox, one lookup added an "Archive" folder and a
  "Junk Email" folder to a mailbox that had neither - and every search asked for Junk Email the same
  way. Looking things up now never creates anything: OutlookAI first checks the mailbox's own record
  of which standard folders it has, and a folder that is not there is reported as not there.
  Archiving mail is the one action that may still create the Archive folder, because you asked for
  mail to be moved into it, and it now tells you when it did instead of promising that nothing is
  ever created. Checking that a message is still a draft before revising or discarding it, and that
  a move is not into Deleted Items or the Outbox, now looks those folders up the same way; if such a
  check cannot be made, on any kind of mailbox, the change is refused and says why - a move used to
  go ahead unchecked. Microsoft 365 and Exchange mailboxes, which always have these folders,
  otherwise work exactly as before.

- 🐛 **The check for new mail keeps the newest 200 mails per folder.**
  It kept an ARBITRARY 200 instead, on every account, for every user, since the feature shipped.
  When a folder holds more just-arrived mail than one pass reads, the check keeps 200 of it, and the
  whole design rests on those being the newest 200: that is what makes "the oldest of this window is
  what is missing" true, and it is what makes a tier that exists to find mail too new to be indexed
  worth having. It asked Outlook to sort the folder by date first, in a form Outlook does not accept
  for sorting, and the refusal was ignored - so the sort never happened and the 200 kept were
  whatever order the folder came back in. Measured against a real mailbox across five accounts: the
  accepted form sorts on all five and puts today's mail first, while the shipped form was refused on
  all five and started at mail years old. Both forms are still tried, the accepted one first.

- 🐛 **A Sent Items folder is sorted by when mail was SENT, no longer by when it was received.**
  Mail you sent was never received, so a sent copy may carry no delivery time at all - and those
  sorted below everything else, which made them the first thing the 200-mail limit dropped. That is
  exactly backwards for a check whose purpose is finding the newest.

- 🐛 **A continued exhaustive search can no longer skip mail.**
  Each page reports the date the next one resumes from, and that date was being worked out two
  different ways in two different places - one of which shifted it by your machine's time-zone
  offset. Shifted the wrong way it makes the next page start past mail that was never returned, and
  the scan still reports itself complete, in the one search mode you choose because completeness
  matters. Both places now read the date through one piece of code, and that code takes the reading
  which can only ever re-read mail, never skip it. There was a second copy of the same mistake in
  the same place: when a mail's date could not be read from the folder listing, the value taken from
  the mail itself was used without converting it at all.

- 🐛 **A failure message no longer says nothing happened when something may have.**
  Every message this product produces about a failure was checked - all thirty-one of them - and
  sixteen were claiming, or implying, that your mail was untouched when the code could not know
  that. The pattern was always the same: a sentence written for the checks that run BEFORE anything
  is changed had been attached to the catch-all that runs after everything. So a draft revision
  interrupted half-way said "Nothing was changed or deleted", a discard that failed during the
  delete said "Nothing was changed", a move that failed said "retry", and a reply that was refused
  because the answer was too big to return said the work had succeeded and nothing had been changed
  in the same breath. Each of those now says what is actually known, and says what to do about it -
  look in Deleted Items, look in Drafts, find the item before moving it again.

- 🐛 **A failed draft creation no longer leaves a draft nothing can clean up.**
  Creating a draft saves it and then does several more things to it - closing the editor, filing it
  in Drafts, opening it on screen - and if any of those failed you were told the draft could not be
  created, while the draft itself sat in your mailbox. Worse, the assistant never learned its
  identity, so the one tool that could have thrown it away was structurally unable to reach it. The
  identity is now recorded the moment the draft exists, the failure message names it, and discarding
  it works.

- 🐛 **A refused move names the folders it created before it refused.**
  Asking to move mail into a new folder creates that folder first and checks the destination second,
  so a move into Deleted Items created the folder and then refused with no mention of it anywhere.
  Any folders created are now named in the failure and listed in the result. This server cannot
  delete folders, so an unwanted one still has to be removed in Outlook.

- 🐛 **A failed attachment save names the file it left behind.**
  If saving an attachment failed part-way, a partial file was left on disk at a name you were never
  told - the name is chosen by the server and gets a numeric suffix when it clashes. The failure now
  names the exact path. Separately, a save that COMPLETED could be reported as a failure because
  reading the finished file's size failed; measuring the file is no longer allowed to fail the save.

- 🐛 **A failed "show me this mail" says it may have shown the mail anyway.**
  Opening a mail in Outlook is the last step of the sequence, so a failure reported over it can
  arrive after the window opened and the mail was marked read. The same applies to navigating to a
  folder or driving Outlook's search box: the window may have moved even though the call failed.

- 🐛 **An aborted send says it has changed the draft.**
  When sending is stopped because the account could not be verified, nothing is sent - that was and
  remains true - but the attempt writes the sending account onto the draft first and does not put it
  back. The message now says so, instead of implying the draft is as you left it.

- 📝 **The discard tool's description warns that a failed discard is not guaranteed to be a clean no-op.**
  The tool description said any refusal comes back as a clear error and it never silently does
  nothing, which is true of refusals and was silently untrue of failures: if Outlook fails during
  the delete itself, whether the draft was deleted is unknown.

- 🐛 **A failed signature change names the half-written backup it can leave behind.**
  Every signature update or delete copies the whole signature aside first, and if that copy fails
  the operation is abandoned with your signatures untouched - that part was already true. What it
  did not say is that the backup folder itself may already have been created and partly filled. It
  is now named in the message when one exists, so there is nothing to discover later.

- 🐛 **Revising a draft is safe to retry when Outlook is interrupted.**
  Editing a draft is not one instruction to Outlook but about twenty in a row, and if the connection
  is restarted part-way through - the same time limit that protects you from a wedged Outlook - the
  draft could be left half-changed. The worst of that was real loss: replacing an attachment removes
  the old file and then adds the new one, so an interruption between the two took your file off the
  draft and put nothing back. Now the request is written down before anything is touched, so
  repeating the identical call FINISHES it instead of doing it again: a file already attached is not
  attached twice, a file already removed stays removed, and the draft's place in its conversation is
  put back from what was recorded rather than from what the interrupted attempt left behind. The
  failure message says so - it used to tell you not to retry, because retrying was the thing that
  made it worse.

- 🐛 **Replacing an attachment attaches the new file before deleting the old one.**
  The order was the other way round, so the moment of danger was a draft that had lost your file and
  not yet gained its replacement. Now the worst an interruption can leave is a duplicate, which you
  can see and undo.

- 🐛 **A draft revision that failed part-way no longer claims nothing was changed.**
  Only some failures prove that, and this was not one of them: it is the one raised when Outlook
  fails somewhere inside the sequence, possibly after the text was already rewritten. It now says
  the outcome is unknown and that repeating the same call will finish it.

- 🐛 **A slow check for new mail is no longer mistaken for a broken Outlook.**
  The safeguard that trips when Outlook stops responding compares how long a request asked for
  against the limit for that kind of request, and with the check for new mail now allowed ten
  minutes it would have looked like a stuck Outlook every time it ran long - taking the product
  offline for thirty seconds for no reason. It is judged against its own limit now.

- 🐛 **The check for new mail keeps what it found when it runs out of time.**
  It used to fail outright, which also restarted the connection to Outlook - so on a large mailbox a
  search lost the whole freshness check rather than getting part of it, and the results already
  gathered were discarded. It now stops at the next folder, hands back the folders it did cover, and
  says plainly that it ran out of time, which accounts it did not reach, and that naming an account
  or a folder will let it finish.

- 🐛 **Two ordinary slow searches no longer make every following request fail for half a minute.**
  The safeguard that trips when Outlook stops responding counted a search running past its own time
  allowance as evidence that Outlook was broken - so on a big mailbox the product could take itself
  offline for thirty seconds with nothing wrong at all. Running past a time allowance the search
  itself chose is no longer treated as Outlook failing to answer; the safeguard still trips on a
  genuinely unresponsive Outlook exactly as before.

- 🐛 **An interrupted send says what may have happened to the mail.**
  If Outlook does not answer within its time limit while sending, the connection is restarted - and
  the message may already have gone out or be sitting in the Outbox. That case used to report only
  that Outlook had not responded, which reads like nothing happened and invites sending again. It
  now says the outcome is unknown, tells you to look in Sent Items and the Outbox before doing
  anything else, and records the unknown outcome in the audit log so the gap is visible afterwards
  rather than merely being a missing line.

- 🐛 **After a restarted Outlook connection, a request that changes mail is no longer told to retry.**
  When one request wedges, the others in flight fail with it - and they were all told "this was not
  your fault, retry it", which is right for a search and wrong for a move, a draft edit or a send,
  any of which may already have taken effect. Reads are still told retrying is safe; anything that
  changes mail is now told the outcome is unknown and to check before repeating it.

- 🐛 **A move or archive batch keeps to its own time limit.**
  The limit on a whole batch was the same as the limit on a single item, so a batch could start one
  more item just under the wire and run for twice as long as its own limit allowed. Each item now
  runs on what is left of the batch's time, and an item interrupted by that limit is reported
  individually - saying that whether it moved is unknown - instead of failing the whole call.

- 🐛 **Outlook's health check keeps the five-second promise in its own description.**
  It says it gives up after five seconds, but on a cold start it could spend ten seconds setting up
  the connection before its own clock even started - so the one tool meant to answer when Outlook is
  wedged had the longest wait in the product. A tool that asks for a short time limit now gets one.

- 🐛 **The Outlook helper process gets a moment to close itself cleanly before it is terminated.**
  It always had a clean shutdown path; the termination followed so closely that the path was never
  actually used, so its own cleanup never ran.

- 🐛 **The check for new mail never builds an answer too big for the connection to Outlook to carry.**
  That check reads the FULL text of every mail it finds, and until now nothing limited how much of
  it crossed - so on an account Windows Search has not indexed, where the check reads a whole week
  of mail, one answer could grow past what the connection can send and the search would come back
  with nothing from Outlook at all. Each mail's text is now cut at 500,000 characters, and one round
  of checking will not carry more than 32 MB of message text in total. Both limits are far beyond
  ordinary mail and are unreachable in normal use. Because that text is what brand-new mail is
  searched against, a cut is never silent: the answer says how many mails were cut, how many of
  those did not match your terms - the only ones where a word further down could have been missed -
  and what to do about it, and the search is marked as possibly incomplete only in that case. A cut
  copy is never mistaken for the whole mail: reading the mail itself still returns all of it.

- 🐛 **A draft no longer hides recipients that could not be looked up.**
  When more than 20 of the addresses on a draft fail to resolve, only the first 20 were listed and
  nothing said the list had been cut - so an assistant told "these 20 addresses did not resolve" out
  of 27 reported a complete-looking list, and the seven it never saw were never raised with you.
  Every one of them stays on the draft and fails on send. The answer now says how many really
  failed, that the list was cut, and how to see the rest.

- 🐛 **A conversation says when it may be missing part of itself because an account is not indexed.**
  Outlook can only follow a conversation inside a single account, so members in another account come
  from the search index - and an account Windows Search has never indexed (a local data file or
  archive, typically) is covered by neither. That was reported only when the index happened to hold
  something for the other account, which on exactly the profiles at risk it does not. Fetching a
  conversation now names the accounts that are outside both checks, marks the answer incomplete, and
  says how to check them: it is careful to say that whether the conversation reaches into them is
  unknown, not that it does not.

- 🐛 **An exhaustive search no longer takes the Outlook connection down on a damaged folder tree.**
  That mode walked folders without any depth limit, so a looping or absurdly deep tree ended the
  helper that talks to Outlook and your assistant was told only that Outlook had gone away. It now
  stops at the same 64-level limit the other folder walks use, keeps scanning everything else, and
  reports plainly that folders below that depth were never opened - so a scan that skipped part of
  the tree can no longer read as a complete one.

- 🐛 **The check for new mail no longer tells you the wrong mail is missing.**
  When a folder holds more new mail than one pass reads, the answer says which folder was cut short
  and that the OLDEST mail there is the part not covered - advice you act on by narrowing the date
  range. That sentence assumed Outlook had sorted the folder newest-first, and when Outlook refused
  to sort it the failure was ignored: the cut was then arbitrary, so the missing mail could be
  anything, and the answer still said it was the oldest. Both cases are now told apart and reported
  separately, and the arbitrary one says plainly that which mail is missing is unknown, so nothing
  sends you looking in the wrong place.

- 🐛 **A named-folder search says when the index cannot see that folder at all.**
  That is typically a renamed or localised folder. The answer already checked for this, but only
  when it came back completely empty, so a single just-arrived mail from the live check hid it
  entirely: the search then looked like a thin result rather than one where everything older than
  the last few minutes was never searched. It is now reported whenever the index contributes nothing
  for the folder, in a field as well as in words.

- 🐛 **A search says it did not look inside attachments of mail from the last few minutes.**
  Text inside attachments is only searchable once Windows Search has indexed the mail; the live
  check reads subject and body only. A search that asked for an attachment-only match already said
  so, but the ordinary case said nothing and reported results as up to date - so a word inside a PDF
  that arrived moments ago was simply absent from an answer that looked complete. The answer now
  says so whenever there is new mail it could apply to.

- 🐛 **An account Outlook cannot name no longer vanishes from the folder list and the new-mail check.**
  If Outlook refuses to report an account's display name - a damaged profile entry, a data file that
  will not open cleanly - that account used to vanish: its whole folder tree was missing from the
  folder listing with nothing saying so, and the check for new mail abandoned it, so recent mail
  sitting in it was invisible to every search. It is now listed and searched under a clearly marked
  placeholder like `(unnamed store 2)`, the answer says how many accounts are in that state, and
  asking about that placeholder by name gets an explanation instead of a "no such account" that
  would have sent you hunting for a typo. Asking for an account that really is absent now also says
  when the profile holds unnameable accounts it could not rule out.

- 🐛 **The folder listing says when it is not the whole tree.**
  Walking the folders stops at built-in safety limits (10,000 folders, 64 levels deep) and the
  answer used to report itself as complete anyway, because "is there more" was worked out from the
  already-shortened list. It now reports that the listing is short, which limit stopped it, and that
  paging cannot reach the rest - list one account at a time instead.

- 🐛 **A shared or delegate mailbox search says when an incomplete folder list narrowed it.**
  Folder scopes in those mailboxes can only be matched by folder NAME, so if the folder walk behind
  them stopped early, folders were searched by nothing at all while the answer looked complete. Such
  a search is now marked as covering less than it was asked to, and says how to get the rest.

- 🐛 **A search for older mail no longer skips an unindexed account while reporting itself complete.**
  Such an account is typically a local data file or archive, which is where older mail actually
  lives. When the requested period ended before the live check would even start, the search answered
  from the index alone and said results were up to date, naming nothing. It now names the account
  the index holds nothing for and marks the answer incomplete, whichever way the live check ended:
  not needed, refused because of the filters you used, or failed outright. Previously only a search
  whose live check ran said so, and only on a profile where nothing at all was indexed.

- 🐛 **An over-large answer from Outlook no longer takes the whole mail connection down with it.**
  When a request produced more data than the connection to Outlook can carry in one piece - most
  likely on an account Windows Search has not indexed, where the check for new mail reads a week of
  mail with full message text - the helper that talks to Outlook exited, every request in flight
  failed, and your assistant was told only that Outlook had gone away. That single request is now
  refused with an explanation naming what was asked for and by how much it was too big, the
  connection stays up, and the assistant is told to ask for less rather than to try the same thing
  again.

- 🐛 **Searching one named account works when Windows Search has not indexed it.**
  That account can be a local Outlook data file (.pst), an archive, a fresh install, or a machine
  where indexing is switched off, excluded by policy or still catching up. Naming that account in a
  search returned "Store 'X' was not found in the local index. Known stores: " - an empty list,
  whose suggested fix was to look up the account name, which gave back the same name that had just
  failed. Searching without naming an account already worked on the very same profile, and so did an
  exhaustive search of that account, so the failure hit the ordinary case only. The account name is
  now checked against the accounts Outlook actually has: an account you have searches normally - the
  index contributes nothing, the live check covers it, and the answer says so and points at the
  exhaustive search for full coverage - while a name you do not have is still refused, now listing
  the accounts you DO have. A name that does not resolve is never quietly widened to your whole
  profile.

- 🐛 **A failed read, open, save, send or archive no longer retries the mail in every other account.**
  Only a mail that genuinely was not found where the assistant looked is now searched for elsewhere.
  Two of those retries were not free: a mail Outlook refused to display could open a window in every
  account before the call reported failure, and a failed attachment save could write a file per
  account. The same tightening was already made for drafts; this completes it.

- 🐛 **An unexpected Outlook error reaches your assistant under its own name.**
  An unexpected error from Outlook reached your assistant labelled with the name of the internal
  connection it had crossed rather than with what went wrong, so a missing file and a refused
  operation read identically. The failure now arrives under its own name.

- 🐛 **A search asked to order by size is no longer re-sorted by date before you see it.**
  The request reached the index correctly and was then undone one layer later. Size ordering is not
  reachable from the mail tools today, so no answer you have had was affected.

- 🐛 **Calendar and contact entries no longer take places away from real mail in a search.**
  Now that a search no longer excludes items by type, an appointment or a contact can turn up among
  the results - but those carry no received date, so nothing decided where they belonged in a list
  sorted newest first. On a mailbox with a big calendar they could in principle have filled the
  whole list and left no mail in it at all, with nothing in the answer to say mail was missing. An
  entry with no date can now never take a place from one that has a date, and when that cannot be
  ruled out the search asks the index again for the dated mail rather than handing back a short
  answer.

- 🐛 **Mail searches no longer lose bounce messages, read receipts and meeting requests.**
  Which of those you got back used to depend on which engine answered your search - the exhaustive
  scan could not see any of them, the index could not see meeting requests, and the check for
  brand-new mail returned all of them - so the same search gave different answers at different times
  and nothing said so. All three now return everything: "did my mail bounce?" is answerable, and a
  meeting request found today no longer disappears once it is indexed. What a search returns can be
  opened too: read, thread and the other tools that open a search hit matched it to its item by the
  time it was received, which an appointment or a task in the results does not have. The index gives
  such an item its creation time, so that is what it is matched on now, and when several items share
  the subject, the nearest in time wins: on the Exchange test machine, holiday appointments of one
  name were created tenths of a second apart.

- 🐛 **A mail search says when it left something out.**
  Three holes could not be seen from the answer at all: mail the check found but could not open,
  mail dropped because a filter you asked for could not be read on it, and mail skipped by an
  exhaustive scan. All three are now counted in the answer, and where a filter is the cause the
  answer names which of your own search options failed to apply - so the assistant can re-run
  without it instead of reporting an incomplete answer as a complete one. A folder in which every
  mail failed to open used to report as fully checked with nothing in it.

- 🐛 **Asking for the folders of an account that does not exist says the account was not found.**
  A typo in the name used to get an empty folder list. It now says so and lists the accounts you do
  have.

- 🐛 **A failed reply, forward or draft from a mail no longer retries in every other account.**
  Each retry could leave a stray draft that the assistant never learns about. It now only retries
  when the original mail genuinely was not in the account it looked in.

- 🐛 **Editing or discarding a draft outside your main account no longer fails with a raw error code.**
  The recovery that was supposed to look in your other accounts never ran.

- 🐛 **Mail tools no longer report every problem as the same unhelpful sentence.**
  An error raised while talking to Outlook - a folder that does not exist, an argument the tool
  rejected, a real Outlook failure with its own error code - reached your AI assistant as "Exception
  has been thrown by the target of an invocation." The assistant could not tell one cause from
  another, so it retried blindly instead of doing the obvious thing, like listing the folders.
  Errors now arrive with their own type and message. The same defect had quietly disabled the mail
  server's recovery from Outlook dropping out, which had therefore never once run.

- 🐛 **Mail searches no longer answer about the wrong dates on a day-first machine.**
  Date limits were written into Outlook's query language in US month-first order, while Outlook
  reads them in your own Windows date format. On a day-first machine every date whose day number is
  12 or lower - roughly 12 days in every month - had its day and month swapped, so a search bounded
  to 1-5 August returned 158 mails received between January and June and not one from August, while
  the same search for 13-15 August was correct. The check for brand-new mail was hit hardest and
  least visibly: its window could land in the future, so it looked at nothing at all and still
  reported that it had checked every folder and that results were up to date. Dates are now written
  year-first, which no regional setting can misread, and the same limits now return exactly the mail
  inside them.

- 🐛 **A search of one account is no longer flagged incomplete over another account's unreadable folder.**
  The freshness check now reports its coverage per account instead of pooling every account's
  problems into one answer.

- 🐛 **A search bounded to older mail is no longer reported as possibly missing recent mail.**
  When the requested period ends before anything new could exist, the freshness check is now
  recorded as unnecessary rather than as failed.

- 🐛 **The sidebar reads the Office colour theme from the right Office version.**
  It used to read the wrong one on machines that have had more than one Office installed. It also
  falls back to the Windows light/dark setting when no Office is detected.

- 🐛 **A search no longer calls itself incomplete because an account has no Junk Email or Deleted Items.**
  A folder an account does not have is no longer counted as a folder the freshness check failed to
  look in.

- 🐛 **The AI sees everyone a mail is addressed to when it picks a signature.**
  It was previously told only about the first 20 recipients, with nothing to say the list had been
  cut, so it could pick a signature for the wrong audience.

- 🐛 **The update line's "checked 4 minutes ago" stays right after a clock change.**
  The update line measured elapsed time against the wall clock, so putting the clock back, a
  daylight-saving change or a time sync could make it read as hours ago or freeze at "just now".

- 🐛 **Outlook tuning works on Outlook 2013 and on future Outlook versions.**
  The settings were written to a location only Outlook 2016 and later read, so on other versions the
  dialog showed "(not set)" forever and the restart notice never cleared. OutlookAI now works out
  which Office version is actually installed, rather than assuming, and the mail server does the
  same - it reports the version it found in its health check, so an Office it does not support says
  so plainly instead of looking like an empty mailbox.

- 🐛 **The writing sidebar grows with its contents at 125% and 150% display scaling.**
  It stayed 280 pixels wide at those scalings while its contents grew.

- 🐛 **Update downloads and setup no longer hang on a stalled connection.**
  A stalled download now gives up and says so, instead of leaving the version line on "checking..."
  for the rest of the session. Setup also no longer hangs indefinitely when Windows or a download
  server stops responding.

- 🐛 **Searching is more reliable when Outlook is slow to start.**
  The first search after Outlook launches is given the time it needs instead of failing, and an
  exhaustive search that runs long now returns the results it found with a note, rather than timing
  out and taking the connection down with it.

- 🐛 **The mail-search instructions reach your AI assistant in full.**
  The guidance the search tool sends was nearly twice the size Claude Code accepts, so it was
  silently cut in the middle - losing, among other things, the rule that tells the assistant to warn
  you when search results are incomplete. It now fits, with the detail moved onto the individual
  search options, where it arrives in full.

- 🐛 **Text in OutlookAI Settings is no longer cut off at any display scale.**
  The explanation under "Make available in all my Claude Code projects" lost its last words, and on
  a display scaled above 100% several other lines in the dialog were cut off too. Every wrapped line
  is now measured against the font actually in use and the dialog grows to fit it, so nothing is
  clipped at any display scale, and if it ever grew taller than your screen it scrolls instead of
  hiding the bottom. The dialog is also a little shorter, because space that was permanently
  reserved for two notices that are almost never shown is now only taken when they appear.

- 🐛 **Mail tools no longer hang forever when Outlook stops responding.**
  If Outlook got into a state where it accepted requests but never replied, every mail tool -
  search, read, drafts, and even the health check meant to diagnose it - would wait silently until
  your AI assistant gave up half an hour later, and the server stayed stuck that way until it was
  restarted. Outlook is now driven from a separate helper process that the server can restart, so a
  stuck Outlook produces a clear, quick error naming what happened instead of silence, and the very
  next request starts from a clean slate. Searches still return your indexed mail while Outlook is
  unavailable.

- 🐛 **Mail tool failures are reported as real errors.**
  Mail tools that fail now mark the response as an error rather than returning a normal-looking
  result that merely contained an error message inside it. Assistants that did not know OutlookAI's
  particular convention could previously mistake a failure for a successful answer.

- 🐛 **Search results say clearly when they are incomplete because the live check could not run.**
  When the live check against Outlook cannot run, results are marked as incomplete and the assistant
  is told in plain words to pass that on, so an answer missing the last few minutes of mail can no
  longer look like a complete one.

- 🐛 **The server no longer restarts Outlook repeatedly in the background.**
  The server could previously start Outlook again moments after a previous copy began shutting down,
  which appears to be what left Outlook stuck and unresponsive in the first place. It now waits
  before starting Outlook again.

- ⚡ **The health check answers in about five seconds even when Outlook is not responding.**
  Checking Outlook's health could itself take over two minutes on a machine where Outlook had
  stopped answering - the one moment the check is worth running. It now reports in about five
  seconds, says plainly that Outlook did not answer, and still gives you everything that does not
  depend on Outlook.

- 🐛 **OutlookAI no longer leaves stray server and helper processes behind.**
  Server and helper processes now shut down with the program that started them instead of
  accumulating in the background - 18 had built up on one machine, one of them stuck holding Outlook
  open.

- 🐛 **A conversation fetch says when it looked in only one of your accounts.**
  Asking for a thread narrows the lookup to a single account for speed - and it does that even when
  you did not ask it to, deriving the account from whichever mail you pointed at. That narrowing was
  invisible, and worse, it silenced the very warning that exists to say a conversation reaches
  further than the answer does: the warning is worked out from what the search index returns, and
  narrowing the lookup narrows that too. So a reply sitting in your second mailbox was both missing
  from the answer and unmentioned by it. The answer now names the account it was narrowed to, lists
  the accounts nobody asked about, marks the result incomplete, and gives the fix that clears it -
  drop the account you named.

- 🐛 **A search that reused a seconds-old check for new mail says so and marks itself incomplete.**
  Repeat searches share one live check of Outlook for up to ten seconds so that refining a search
  stays fast; mail arriving inside those seconds is in neither the index nor the reused check. The
  answer always carried how old the reused check was - sitting next to the word that says results
  are fully up to date. It now says the result is incomplete, states how wide the gap is in seconds,
  and says it closes by itself.

- 🐛 **A snippet length you asked for is no longer shortened silently.**
  Asking for more context than the limit allows gave you the limit's worth with no mention of it, so
  a snippet that had been cut read like a whole one - and judging a mail irrelevant from it is a
  decision made on text nobody said was truncated. Both directions are now reported. An exhaustive
  search does not shorten this setting but drops it entirely, which was equally silent and is now
  said plainly.

- 🐛 **Asking for a conversation by mail fetches the whole of it, across every account.**
  Asking for a thread with just a mail's id - the shape an assistant reaches for straight out of a
  search - made this product quietly narrow the lookup to that mail's account, so replies filed in a
  second account were missing from an answer that presents itself as the full conversation. Nobody
  asked for that narrowing; it was taken because the mail had to be opened anyway. It is gone: a
  lookup you did not scope now covers every account. Naming an account yourself still narrows it,
  because that is a choice you made and can drop, and the answer still says what the choice cost.
  The side effect is that the warning about members in an account that was never looked at now fires
  only when you really did narrow the search, and the stronger warning - "this conversation
  demonstrably continues in another account" - can finally be raised, because the evidence for it is
  no longer being filtered away.

- ⚡ **Whole-mailbox searches reuse a recent check of Outlook, as account-scoped ones always did.**
  Repeat searches are meant to share one live check for up to ten seconds, so that refining a query
  runs at index speed - but that sharing only ever worked when you named an account. A search across
  the whole mailbox worked out what it recognises the check by from the clock, on every call, so two
  identical searches a second apart never matched and each of them walked your folders again, which
  is the slowest part of a search. Whole-mailbox searches now reuse the check exactly as
  account-scoped ones always did. Nothing about what is searched changes, and the check is still
  thrown away the moment Windows Search takes in new mail, so a reused answer can never be older
  than the index it stands in for.

- 🐛 **OutlookAI lets go of Outlook properly when the assistant's connection to it closes.**
  This product talks to Outlook through a small helper program, and until now that helper was always
  ended the abrupt way - even when nothing was wrong and your session was simply finishing. An
  abrupt ending runs none of its own tidying up, and one of the things it leaves behind is a hidden
  Outlook window the helper opens when it has to start Outlook for you. Outlook counts that window
  even though you can never see it, so it accumulates, and a window Outlook thinks is open is a
  reason for Outlook not to close when you tell it to. The helper is now asked to leave and given a
  moment to do so, which is when it closes that window and hands back everything it was holding; if
  it does not go within two seconds it is still ended the old way, so nothing can make closing
  slower than it was. To be clear about what this is NOT: leftovers from the abrupt route were
  measured against a real Outlook and did not make Outlook slow or unusable for anything that came
  next - the cost was the accumulating hidden window, not a broken Outlook. This also cannot help
  when the assistant kills the helper outright, which is what happens when Outlook has genuinely
  stopped responding and when a session ends abruptly.

- ✅ **On a busy host the test machines give Outlook 20 minutes to open, and say why when it does not.**
  Eight live runs on the two indexed test machines stopped on 2026-11-18, between 03:49 and 05:16
  host time, because Outlook had not shown its window within the 10 minutes
  `Testbed/guest/Start-OutlookUnelevated.ps1` gave it. Outlook there took 94 to 585 s on 2026-11-17
  and 125 s at 02:50 that night; in that window the host's processors were saturated, the unindexed
  machines' Outlook took 13 to 45 s instead of 6, and copying a run in took 6 to 12 minutes instead
  of under 3. The starter waits 20 minutes now, and when the time runs out it records Outlook's
  processor time, how much of it fell in the last minute, and its visible windows, so a slow start
  can be told from one a dialog blocks. Its `-SelfTest` pins the default, the report and that no
  wait reads the wall clock.

- ✅ **The test machines' waits run on a stopwatch, so a guest clock that moves cannot end them early.**
  Twice a live suite was taken for finished while it still ran: on the unindexed test machine, whose
  clock is frozen, `Testbed/guest/Register-InteractiveTask.ps1` reported "No exit.txt after 7200 s"
  40 min 17 s into one run and "after 9000 s" 34 min 50 s into another, because it read its deadline
  off the guest's wall clock and that clock had moved ahead. Why it moved is open
  (`Docs/live-tier-on-the-vm.md` section 4.5). That wait, and every wait in
  `Start-OutlookUnelevated.ps1`, measure elapsed time with a stopwatch now.

- ✅ **A test machine's .NET check no longer fails a working SDK because it could not read an exit code.**
  `Testbed/guest/Install-DotnetSdk.ps1 -Verify` stopped two live runs on the indexed test machines
  on 2026-11-18 with "dotnet --list-sdks exited " and an empty code, while the SDK it wanted was
  listed on the next line. `Start-Process` keeps no handle to the process it starts, and reading
  `.Handle` straight away narrows that gap without closing it: on a busy host, dotnet had already
  gone. Each check passes on exit 0, or now on an unread code when its own output proves the step
  (the SDK listed, the probe's sentinel printed, the tests counted) and says so in its log. An
  unread code without that proof still fails, and the message reads `(unread)` instead of nothing.
  Keeping the handle instead would have meant reading dotnet's output through pipes, which a build
  server that inherits one holds open after dotnet exits.

- ✅ **Test runs that wait for a test machine no longer ask Hyper-V every few seconds.**
  Every queued run asked Hyper-V about every machine it could use every 5 seconds, through the WMI
  calls `Get-VM` makes, and read the host's free memory through WMI as well. On the night of
  2026-11-18, with eighteen runners alive and at most eight of them running, the maintainer's PC
  lagged, and a trace of 15 seconds counted 32 WMI calls a second, 91 % of them Hyper-V's. A run that
  waits now takes a ticket in a queue of files and tries only the machines no earlier run wants, first
  come first served, where a newer run could take a machine before an older one. It tells a free
  machine from its lock and lease files and the host's memory from Windows' `GlobalMemoryStatusEx`,
  asks Hyper-V once, before it waits, which of its machines exist, and sleeps until a lock, a lease or
  a ticket changes, or 30 seconds pass. A running run reads its machine's state and heartbeat with one
  narrow query each, by the machine's id, instead of `Get-VM` and `Get-VMIntegrationService` (4 and 23
  WMI calls), and asks its machine how far it got every 15 seconds instead of every 3. A live run makes
  all its own calls into the guest over one PowerShell Direct session instead of one each, and the
  idle-saver reads the leases first and asks Hyper-V about the rest in one query. Measured on the same
  host that morning, a run waiting the old way cost about 0.43 of a logical processor: 0.34 in
  WmiPrvSE, nearly all of it the memory read, and 0.08 in vmms and the WMI service, which answer
  Hyper-V's calls. A run waiting the new way costs nothing measurable, and a trace of 15 seconds found
  no WMI call from it, where each of four queued runs on the old runner made 30 to 34.

- ✅ **A test machine that is running is saved before a checkpoint is restored over it.**
  Both runners restored a checkpoint over a test machine in whatever state it was, and at the end of a
  run it is running. In another session's runs that night, a guest restored that way stayed out of
  PowerShell Direct's reach for six minutes, where one saved first and then restored answered in
  seconds. The build-VM runner and the live-tier runner now save a running or paused machine first.

- ✅ **The Exchange test machine makes drafts and mail again, so its write tests run there again.**
  Its Office LTSC 2024 was never activated, and past its 30-day grace Outlook there refused to create
  any item through its object model: on 2026-11-17, 41 of 87 tests failed at their first draft or mail
  while every test that only reads passed, and the runner was narrowed to the tests that write nothing.
  On 2026-11-18, as the maintainer decided (206 (a)), the grace was reset once with Office's own
  `OSPPREARM.EXE`: the licence went from notifications to 30 days of grace, and its re-arms left from
  three to two. A draft test then made and removed its draft at the first try, Outlook started without
  Office's activation wizard, and the runner gives the machine its whole read-and-write selection again.
  Its first whole run made every item it tried to: 95 of 106 tests passed, the sweeps left nothing
  tagged in the test account, the count tripwire found the bystander mailbox unchanged, and none of
  the 11 failures was Office's. The machine rests on a new checkpoint holding the re-armed grace. The
  grace runs out again on 2026-12-18; a run after that is narrowed to the tests that write nothing with
  `-FilterSuffix '&Writes=Nothing'` until it is re-armed.

- ✅ **The index-refile measurement no longer fails when Outlook lists a moved folder late.**
  The live test that times how fast the index follows a folder rename and move reads the folder's
  ID under its new parent right after moving it. On the copy of the indexed test machine the parent
  read before the move did not list the folder yet, and the test failed on that read alone. The
  test helper now opens the new parent again and looks once more, for at most five seconds.

- ✅ **The audit log's eight-writer test passes on a busy machine, and still checks the lock's promise.**
  With the writer lock on, it demanded that no line was ever written without the lock and that the
  whole file was in time order. On 2026-11-17, with three build runs and four Outlook test machines
  on one host, writers waited longer than the lock's 2 s and appended without it, as the product
  does by design, marking the line `lock="timeout"`: 12 of 1,000 lines on one build machine and 1 on
  another, and the test failed on both. It now checks what the product promises instead: every line
  written without the lock says so, exactly as many per writer as that writer's own count of
  appends that ran out the wait, the lines written under the lock are in time order, and most lines
  were written under it. How many lines waited it out goes to the test's output. A new test holds
  the lock until three writers have each run out the wait, so that path is checked on every build,
  not only on a busy machine. The lock's wait is unchanged.

- ✅ **The signature test swaps one signature for another and checks the first is gone.**
  The test meant to prove that a draft's signature can be replaced applied the same signature a
  second time, through a test-only path that saves in a way measured to keep nothing, so it passed
  whether or not anything changed. It now replaces the signature with a different one through
  update_draft, checks that the first is gone and the second is there once, then removes it with
  "none". The test-only path is deleted.

- ✅ **The test machines' account password never expires, so runs on the build machine log on again.**
  Windows expires a password after 42 days, and nothing in the testbed said otherwise. The build
  machine was installed on 2026-10-03, so its base checkpoint would have refused every run's
  PowerShell Direct logon from 2026-11-14, the day the Exchange machine's did ("The credential is
  invalid"). The Outlook test machines' passwords would have expired on 2026-10-27, after their
  frozen clocks' instants, so only a run from an unfrozen checkpoint met it. The flag is now set in
  every checkpoint a runner restores: the build machine's `CP-03-SDK-TEST-READY-UTC`, made again
  under its own name from the old one with its clock held back so the password still worked, and
  the Outlook machines' new frozen checkpoints `CP-22C-FROZEN-CLOCK` and `CP-16B-FROZEN-CLOCK`,
  whose frozen instants moved 27 and 28 seconds. A new test machine sets it at its first logon, and
  the answer-file builder refuses to build one whose first-logon script would not. On the Exchange
  test machine, which no runner restores from a checkpoint, it was set from its console on
  2026-11-17, the password itself unchanged, and checkpoint `CP-07-TEST-ACCOUNT-READY` holds it.

- ✅ **Two index-coverage tests pass on any date, not only before 2026-11-16.**
  `UnindexedStoreReportingTests` searched for mail older than today minus 90 days while its stub index
  stays at 2026-08-18, so from 2026-11-16 the search reached past the stub's frontier and the two
  tests that expect no sweep failed on every machine, master included. The search now ends 90 days
  before the stub's frontier, as the file's store-scoped test already did.

- ✅ **A test-machine restart refuses Outlook's own Office dialogs instead of waiting behind them.**
  The restart script refused only classic Windows dialogs, so Outlook's start-up "Check out our new
  look" theme dialog went unseen on the Exchange test machine on 2026-10-08: Outlook ignored the
  request to quit for half an hour, and had never added the new mailbox to Windows Search either. The
  script now names and refuses that kind of dialog too.

- ✅ **The live-test runner deletes the build it made for each run instead of leaving 180 MB behind.**
  Git unregistered the commit's build tree after every run but left its directory, so every run of the
  test machines' runner kept about 180 MB of build output in its results folder. The runner now deletes
  what git left, and keeps the build tool's worker processes from holding the files open past the run.

- ✅ **The live tests find a mail store in the search index the way the server itself does.**
  They now find it also when no sample of indexed mail reaches it and no address names it. On the
  test machine with the 160,000-item measurement store, thirteen tests stopped before measuring
  anything because of this. Nothing that ships to users is affected.

- ✅ **The sweep-cost measurement script no longer creates missing folders or times the wrong one (D101).**
  It used to ask Outlook for a data file's standard folders in a way that creates the missing ones.
  The measurement mailbox on the indexed test machine has no Inbox, Sent Items or Junk Email of its
  own - its mail sits in stand-in folders - and the way the script asked Outlook for them is the way
  that, on the test machines' other data files, created a missing Junk Email folder and answered an
  Inbox request with the file's hidden root. The script now finds folders the way OutlookAI's own
  sweep does, says which ones that sweep skips, and times the stand-ins separately, labelled as
  folders the sweep never reads; the other measurement script can be pointed at a stand-in folder.
  The script had never run: its first run, against that mailbox, also found that it could not walk a
  single folder - PowerShell turned every Outlook collection it handed on into a plain list - which
  is fixed too. OutlookAI's own sweep was measured against the same mailbox and was right: it reads
  the one standard folder the file has, Deleted Items, and does not count the three it lacks as
  gaps.

- ✅ **Building the add-in for the test machines works again.**
  It stopped at once after the release script arrived, saying Visual Studio's Office build files
  were missing when they were not.

- ✅ **The test machines' hub rebuild waits for a list file another program still holds.**
  It waits for up to a minute after the rebuild's first step, and no longer stops with the test
  mailbox emptied and the test run lost. It happened twice in a day. Nothing that ships to users is
  affected.

- ✅ **Three mailbox checks the first full test-machine run caught asking the wrong question are fixed.**
  The check that a complete search of one folder returns exactly that folder's mail never said "this
  folder only", so since folders came to include their subfolders by default it compared a whole
  subtree against one folder's mail; it now asks for the folder alone and, separately, checks the
  default against the subtree. The check that repeated searches are answered from memory now runs
  only where the search index holds mail, because without one every search is meant to look afresh.
  And the health check's test now expects what the tool says when the search index is reachable but
  holds no mail - a problem, not advice.

- ✅ **Mailbox checks work on a test machine that runs none of the search-index checks.**
  The check that the local mail server answers, the check that no mail is left queued from an
  earlier run, and the nudge that asks Outlook to fetch mail while a check waits for it now switch
  on for every run, not only when an index check happens to be selected - and every check that sends
  itself mail now waits through that nudging helper. On the first full run on a test machine, three
  checks timed out with their mail sitting on the local mail server, because Outlook's one fetch had
  finished a moment before the mail arrived and nothing asked again. Nothing that ships to users is
  affected.

- ✅ **The remaining mailbox checks that could pass having checked nothing no longer do.**
  The maintainer decided them one by one. The check that a search always answers passed on any error
  at all - the very failure it exists to catch - and now fails on every error except the one where
  the machine's search index cannot be reached. That case, and the checks that Outlook says when to
  retry while it is starting or not responding, that the helper process Outlook work runs in never
  outlives the server, that Outlook's search setting can be switched both ways, and that reading a
  message and following its conversation open no window, now say when there was nothing to check,
  the way the project's other mailbox checks already do: on the project's reference machine they
  fail, and on a test machine they print a PROVED NOTHING line naming what was missing - so on the
  reference machine the retry check now fails whenever Outlook is healthy, by design. The check that
  Outlook recovers after it exits used to stand aside, green, whenever a person might be using the
  computer; it still does on the maintainer's own machine, with a line saying so, but on a test
  machine, where nobody is at the keyboard, the same situation now fails - and it now checks for
  unsent mail and open compose windows before making Outlook exit even when no Outlook window was
  open, which it used to skip. The check that a shared mailbox's messages open from a folder the
  search index lists flat now waits up to five minutes for Exchange to finish listing that folder
  before it gives up. Pinned by new checks that need no mailbox, with a control that fails against
  the old code and a check for each safeguard that fails when it is removed. Nothing that ships to
  users is affected.

- ✅ **The project's own tests no longer write into your real OutlookAI audit log.**
  The audit log (`%LOCALAPPDATA%\OutlookAI\audit.log`) is OutlookAI's record of every draft, saved
  attachment, move, archive, send and signature change it makes, and a draft or send whose line
  cannot be written is reported as an error - but the project's automated checks drive those same
  paths with stand-in mail, and every one of them added a line to the real log of the computer that
  ran them: when it was last counted, almost two lines in three were test noise. Every test run now
  writes its audit lines to a throwaway folder of its own under the temporary folder, deleted when
  the run ends, and a test run can no longer write the real log at all - anywhere outside the
  temporary folder is refused before anything is created. One check went further and made the mail
  server itself write a line to the real log on every run; it now runs inside the test process, and
  the checks that talk to the mail server refuse to ask it for that kind of call. All of this is
  proven without opening the real log, which your own OutlookAI may be writing to while the tests
  run, and the proof was shown to fail on the old code. Where the log lives, what OutlookAI writes
  to it and what it reports about it are unchanged; lines already in the log are left as they are.

- ✅ **The test machines' record says both 20,000-item measurement mailboxes are built.**
  It still said neither had been built. Both were built in mid-September, as the notes on building
  the machines record; the record now says so and cites where, and notes that the indexed machine's
  is still below the 160,000 items decided for it. The tool that writes each test machine's settings
  no longer calls such a mailbox "not built" just because its build parameters have not been copied
  into the record yet. Nothing that ships to users is affected.

- ✅ **Three mailbox checks against a real Outlook no longer pass having checked nothing.**
  The checks that a search of the whole test mailbox handed out one page at a time returns exactly
  what a single search does - no message skipped, none twice - and that a continued search refuses a
  changed or outdated continuation all need the test mailbox to hold more than a page of mail; on a
  smaller one three of them quietly ended green, and the fourth compared two one-page answers that
  could not disagree. The check that a delegate mailbox's messages still open from a folder the
  search index lists flat did the same whenever Outlook had not yet listed that folder. Each now
  says so, the way the project's other mailbox checks already did: on the project's reference
  machine it fails, because there it means something has drifted, and on a test machine it prints a
  PROVED NOTHING line naming what was missing and how to supply it. The check that Outlook recovers
  cleanly after it exits used to skip, green, whenever anything was waiting in the Outbox; it now
  fails on every machine instead - it cannot run without closing Outlook, which is never done while
  mail is queued, and on a test machine nothing but a test ever queues mail, so a full Outbox there
  is a stuck send or a previous run's leftovers that should not exist. Pinned by 32 new checks that
  need no mailbox, with a control that fails against the old code and a check for each safeguard
  that fails when that safeguard is removed. Nothing that ships to users is affected.

- ✅ **The test machines' session runner no longer reports success for work that exited with a failure.**
  Work handed to it as a piece of script text used to be pasted into the runner's own script, so a
  deliberate exit with an error code ended the runner itself and the result it recorded was success;
  and text with a syntax error left no result at all, so whoever started it waited out the whole
  time limit. That work now runs as a script of its own, the way work handed over as a file always
  has, so its exit code, its errors and its syntax errors all come back as they are, and work handed
  over as a file behaves exactly as before. Shown on the maintainer's machine under both versions of
  PowerShell, with the old runner alongside as the control; the same check on a test machine is
  recorded as still to do.

- 📦 **A build of the add-in outside Visual Studio no longer switches the builder's own Outlook to it.**
  Every build used to register itself with the Outlook on the machine that built it and mark itself
  as trusted, so a developer's Outlook ended up loading whichever build folder had been built last -
  and cleaning a build removed the add-in from Outlook altogether, installed copy included. Building
  inside Visual Studio still registers, because debugging needs it. A new developer script puts a
  build on your own Outlook only when you ask for one - copied out of the build folder first, so
  later builds cannot change it underneath you - takes it off again by putting the installed release
  back, and reports which build Outlook will load at its next start.

- ✅ **Four faults the first real build of the test machines' small mailboxes found are fixed.**
  The mailbox owner was never a proper recipient of the mail addressed to them, because Outlook
  refused the form their address was written in; mail meant for a data file's Inbox was filed in a
  hidden folder nobody could see, because such a data file has no Inbox and Outlook handed back its
  hidden top instead; trial items the builder writes and deletes before a build were left in a
  different mailbox nobody had allowed, because Outlook files a new unsent mail in the default
  mailbox's Drafts; and asking Outlook for standard folders created ones the mailbox was meant not
  to have. The builder now writes the owner in a form Outlook accepts and checks exactly that before
  building, files mail only into folders Outlook shows (a visible stand-in folder where a data file
  has none), writes mail into a mailbox that is not the default one only in the one way a test
  machine showed stays in it, checks where every first save landed and removes it there, clears its
  own trial items from every other mailbox before and after, and never makes Outlook create a
  folder. The unindexed test machine's trial runs confirmed each fix; the mailboxes themselves are
  not built yet (see the entry on undated items).

- ✅ **The test machines' identity account has a real Inbox, no longer a hidden folder.**
  Its data file used to be created empty and attached, and such a file has no Inbox, so the
  account's mail would have gone where Outlook shows nothing. Outlook now makes that data file
  itself, as the default data file of a throwaway profile - which gives it every standard folder -
  and the setup checks it has a real Inbox and holds nothing before it is named, attached, and made
  the account's delivery folder. The setup script also refuses a data file with no real Inbox,
  checks the delivery folder when it verifies the account, and no longer creates a Drafts folder
  while it is only meant to be reading. Built this way on both test machines, where the account now
  delivers into its own Inbox: on the unindexed one first, then on the indexed one, rebuilt from a
  checkpoint taken before its old account and taken through the rest of its setup again - the search
  index, the local mail server, the add-in, the signature and the test suite - each proven again,
  with its mail searchable throughout. The search index there now files that account's mail file
  under the account's address, instead of under the generic name another mail file on the machine
  also had. The rebuild also showed that the test machines' automatic approval of Outlook's security
  prompt must be in place before this setup runs, not after: without it, one of the setup's reads
  stopped on the prompt about two minutes after a restart, although Windows reported the antivirus
  as up to date, so the build instructions now put that step first.

- ✅ **The mailbox tests no longer add folders to the test machines' mailboxes.**
  The check that every test run leaves each mailbox as it found it, and the clean-up that removes
  the tests' own messages, asked Outlook for each mailbox's standard folders in the same way the
  product used to - the way that makes Outlook create any that are missing. On a test machine that
  added a "Junk Email" folder to mailboxes the tests are never allowed to change, before the check
  had even recorded what it would compare against. Both now look folders up without ever creating
  one, exactly as the product itself does since this release's fix for POP3, IMAP and data-file
  mailboxes; a folder they cannot confirm exists makes the leftover-message count fail loudly rather
  than count it as empty; the read-only archive test now also proves that no mailbox's folder list
  changed; and the two archiving tests expect the Archive folder to be created, and reported,
  exactly when a test machine's mailbox had none. Nothing that ships to users is affected.

- ✅ **The project's own tooling runs on the PowerShell that ships with Windows.**
  The test machines therefore never need anything extra installed. Ten scripts declared they worked
  on it and did not: they worked out where the project lives in a way that version leaves blank at
  that moment, so they failed before doing anything - including three of the checks the project's
  build runs on every change. They now run under both that version and the newer one, verified by
  running each under both. A second, quieter failure is fixed as well: when a script ran another
  program and captured that program's error output, the older version stopped the whole script at
  the first line the program wrote there - even a progress report - before the script could see
  whether the program had succeeded. The tool that makes a test machine's install disk did exactly
  that on its own "0% complete", and the measurement-mailbox builder would have lost its tool's real
  exit code and message the first time the tool printed a warning. Six scripts now read that output
  safely under both versions and still stop on a real failure exactly as before; the privacy check
  and the pre-release measurement gate no longer need the newer PowerShell installed to run their
  own sub-checks; and one character the older version would have misread was replaced. A new build
  check fails on all three kinds of mistake in any script, and the build now runs every one of its
  checks under the version that ships with Windows as well as the newer one, so none of this can
  quietly come back. Nothing that ships to users is affected.

- ✅ **The indexed test machine's mail is in Windows Search at last, and the other is really kept out.**
  The machine had never held a single one of its 20,000 test mails in its search index, and the
  reason turned out to be the way every tool on the test machines starts Outlook: with administrator
  rights. Outlook started like that quietly switches its Windows Search integration off - no error,
  no prompt, no log entry, and Outlook's own "is search working" flag simply reads false - while the
  same Outlook started the ordinary way registers itself with Windows Search within seconds and had
  the whole test mailbox indexed in under ten minutes; that was shown on one machine, from one saved
  starting point, changing nothing but how Outlook was launched. The fix has two parts. Whether a
  test machine's mail belongs in the index is now set through the documented Windows interface the
  Indexing Options window itself uses - reached through hand-written declarations of that interface,
  checked against Microsoft's own definition line by line and held in place by two independent
  checks, each shown to fail when a single declaration is moved or altered. And the mail is then
  indexed by starting Outlook the ordinary way, which a new script does, because the setting on its
  own indexes nothing: Windows Search only takes in what a running, non-elevated Outlook hands it.
  The check that reports a machine's index state now asks Windows Search itself whether the mail is
  in or out and why, counts the indexed mail per mailbox, only calls a machine indexed once indexing
  has actually finished rather than merely started, and no longer reports a working index as missing
  when it runs in the first seconds after a restart. Taking mail out of the index now actually does
  so, in an order that was measured rather than assumed: the Windows policy setting the old script
  relied on turned out not to remove anything - Windows Search still counts the mail as included -
  so the script now writes the same kind of exclusion rule the Indexing Options window writes, waits
  while Windows Search deletes the mail it had already indexed (about four and a half minutes for
  20,000 mails), and only then applies the policy and restarts the search service, because
  restarting it during that wait was found to make Windows Search drop the deletion - the mail was
  all still there after a reboot. An Outlook started on top of the exclusion left it in place, and
  on a machine that was never indexed the policy on its own did keep Outlook from adding itself.
  Test machines are also now restarted gracefully: the restarts on record used a Windows command
  that, as Microsoft documents, force-closes every open program whenever it is given a delay - and
  most were given one - so any Outlook still running was shut down with its mail files open. A new
  script closes Outlook the way the project's mailbox-safety rules require - refusing while a draft,
  a dialog or an unsent message is open, and never killing it - and only then restarts the machine
  without force. Nothing that ships to users is affected.

- ✅ **The four profile-building scripts ran on a real test machine, and what that found is fixed.**
  They build a test machine's Outlook profiles and its measurement mailbox, and had not run since
  they were rewritten. The open question about Outlook's profile-import setting is settled: Outlook
  clears it itself, and it is now cleared regardless. Each script ran from the same saved starting
  point, restored fresh before every one. The script that switches the default profile worked end to
  end - including refusing a profile that does not exist without touching anything, which was shown
  rather than assumed - and Outlook then opened the chosen profile with no prompt. The script that
  adds a named data file to a profile worked on every path, and the stall its notes warned about
  never happened. The profile-creation script's import worked too, but three things were wrong with
  it. Its check-up mode crashed on its very first line, because a loop variable shared its name with
  one of the script's own settings and PowerShell silently turned every data file it checked into
  plain text; its self-check now reads the script's own source for exactly that mistake. Its search
  for stray data files could not see the folder these machines actually put them in. And - the one
  that matters - a profile with no mail account, built from a profile file, stops on Outlook's "set
  up your email" window at every single start, so no unattended tool can ever use it; the script now
  refuses to build one unless someone will be at the screen to click through, and points at the
  route that needs no one - Outlook's own no-email mode followed by the data-file script - which was
  run end to end on the same machine the same day and verified clean. The measurement-mailbox
  builder's up-front check turned out to refuse the machine's real, correct profile, because it
  counted the profile's own data file and address book as mail accounts; it now counts only mail
  accounts. With that fixed, a full run from the same starting point passed its checks and confirmed
  all 20,000 items of that machine's measurement mailbox present, each exactly once and in the
  folder it belongs in. Whether Outlook leaves its "import this profile file" setting behind - which
  would rebuild the profile at a later start and cut the measurement mailbox out of it - is now
  measured rather than feared: Outlook removes it within five seconds of acting on it. It is removed
  anyway: the step that checks an import now clears a leftover one, but only once the import has
  really happened, never before; and the measurement-mailbox builder refuses to start while one is
  set, a check its bypass switch deliberately cannot skip. The tier-profile script, which uses the
  same setting, gained the same clean-up, the guard against being run anywhere but a test machine
  that the other profile scripts already had, and fixes to four of its checks that were failing the
  one tier profile that works. Two things found along the way are recorded for whoever runs the
  mailbox tests next: the tier profile raises a password prompt at every start, and on these offline
  machines reading an email address through Outlook's programming interface can raise Outlook's
  security prompt, which waits for a click that never comes - the saved starting point already has
  one on screen. Nothing that ships to users is affected.

- ✅ **Building a test machine from nothing works again: the last two broken scripts are fixed.**
  They create the Outlook mail profile the project's measurement mailbox lives in, and the data
  files inside it, and both were built on the same Windows component that the profile switcher had
  already found unusable on these machines' version of Office - so a rebuild from scratch would have
  failed at its first step, and rebuilding from scratch is precisely how these machines are meant to
  be proved. Both now use routes that have been measured working on the machines themselves: the
  profile and its data files come from a settings file Outlook reads once when it starts, which is
  the same mechanism that already builds the project's other test profile; and a data file added to
  a profile that already exists is attached first and given its exact name afterwards, by renaming
  it - which was measured to work and settles a question four rounds of research could not answer
  either way. The unusable component and the several hundred lines of glue around it are deleted
  outright rather than kept as a fallback nobody exercises, because an unexercised second route is
  the defect that caused all of this; it remains in the project's history, and the note left behind
  records the best current guess at why it fails and the small experiment that would confirm it,
  while saying plainly that nobody has run that experiment. One diagnostic mode went with it - it
  asked whether Outlook tolerates an "@" in a data file's name, which has since been answered twice
  over by measurement, and it needed to delete a throwaway profile, which nothing can do any more;
  that missing ability is written down as a gap rather than quietly dropped, and nothing in the
  build needs it. Both scripts can now check their own decision-making on any machine without
  touching Outlook at all, a hundred and forty checks between them, and each prints the list of
  things only a real test machine can settle so that a clean run is not mistaken for proof. **One
  hazard travels with the new route and is worth knowing about**: the setting that tells Outlook to
  read that profile file is consulted at every start, and whether Outlook clears it afterwards is
  unknown - if it does not, every later start rebuilds the profile and would detach a measurement
  mailbox that had since been filled, costing a thirteen-minute rebuild of twenty thousand items.
  The remedy is a single command, it is now a numbered step in the build order rather than a
  footnote, and the scripts warn by name while the setting is still there. The three documents that
  describe all of this have been brought back into line with it, including several places that had
  gone on recommending the route that does not work, or restating as unanswered a question that had
  been measured and answered elsewhere in the same file.

- ✅ **A timed-out mailbox operation in the measurement tools stops, no longer left running.**
  The tools that build the project's measurement mailbox drive Outlook on a dedicated worker, under
  a time limit. When that limit ran out the tool reported failure and returned - but the worker
  carried on, still holding Outlook and still working, until the whole program exited. For a read
  that wasted effort; for anything that WRITES it meant an abandoned session still changing the
  mailbox after the caller had been told it had stopped, and the obvious next move - run it again -
  would then have put two writers on the same mailbox. The worker is now asked to stop at a safe
  point, the tool waits for it, and it reports which of two things happened: it stopped cleanly, and
  how far it got; or it did not answer, in which case it says so loudly and says not to re-run.
  Those are different situations inviting opposite responses, and they now also produce different
  exit codes, so a script can tell them apart. The check that makes this work is built into the
  loops themselves rather than added alongside them, so the mistake cannot quietly return.
  Separately, the time limit was being spent on starting Outlook before any work began - which is
  what actually ran out on the test machine twice - so starting up now has its own allowance, and
  the limit on the work itself only counts time in which nothing progressed. Nothing that ships to
  users is affected.

- ✅ **A live index check that never managed to ask its question fails on the reference machine.**
  One live check seeds a mail, waits for the machine's search index to notice it, and then asks the
  index whether it has - and each of those questions is given a time limit, because a question that
  runs too long says more about the machine being busy than about the feature. If every question ran
  out of time, the index was never actually asked anything, and the check still passed - printing a
  line saying it had proved nothing, and passing anyway. On the machine the project treats as its
  reference machine, that now FAILS instead, because a green result there means the machine is as
  expected and this one could not have known. On a laptop or a spare machine, where a saturated
  indexer is ordinary, it still passes and says the same thing just as loudly - the finding is
  identical, only whether it should stop the run differs. The message also now explains what a
  genuinely slow indexer looks like instead, since that is the thing it is most likely to be
  mistaken for: the questions COMPLETE and simply keep coming back empty. Same treatment two other
  checks in this suite already had. Nothing that ships to users is affected.

- ✅ **The corpus builder's trial litter no longer fails every build or reads as a corrupt corpus.**
  Before writing anything the builder runs two short trials against the target data file - one to
  find out where that file will really file mail, one to find out whether it will accept back-dated
  mail - and each trial creates a throwaway mail and deletes it again. But deleting mail in Outlook
  does not remove it, it moves it to Deleted Items; so every one of those throwaway items stayed in
  the data file permanently, under an identity nothing had written down. One build left six to eight
  of them, and building the documented way left twelve to sixteen, because that route runs the
  trials twice. The self-check at the end of the build then counted them as corpus mail and
  announced that the corpus held an item more than once and that twelve items were in the wrong
  folder - the sentence that means every measurement taken against that corpus is worthless - about
  a corpus that was in fact perfect. That check had never once managed to run, because it tripped
  over its own file lock, and the moment that was fixed the builder started failing on every
  successful build and the corpus could not be rebuilt at all. Three things are fixed. The trials
  now clear up after themselves, both after each item and once before they start, so a data file
  carrying residue from an older version is cleaned by the next run that touches it. The identity of
  each throwaway item is now recorded the instant the item exists rather than three steps later,
  closing the window in which an interruption left one that nothing could ever remove. And the
  self-check counts them separately and names them for what they are, which gives the duplicate
  warning back the only meaning it was ever written for: a build interrupted so late that it created
  the same mail twice. The rule deciding which mail the clean-up may touch is pinned by ten new
  tests that need no mailbox at all - it deletes mail, so what chooses its targets has to be
  checkable on a machine that has none - and the two tests that used to pin the wrong behaviour now
  pin the right one, with the faulty report they were written from kept beside them as evidence.
  Nothing that ships to users is affected.

- ✅ **The corpus builder no longer fails at the end of a successful run, and its self-check runs again.**
  The builder writes a corpus of test mail into a test machine's data file, records every item in a
  manifest file as it goes, and then reads that manifest back to check what it actually produced -
  the check that exists because an early run once created every single item in the wrong folder and
  reported success. It could never do it: it was still holding the manifest open for writing when it
  tried to read it, and Windows refuses that combination outright. So a complete, correct build
  ended in a file-in-use error and reported failure, and the self-check that was supposed to decide
  whether the corpus is usable had never once run. Both halves are fixed - the reader now tolerates
  a file someone else is writing, and the builder closes the manifest before checking itself - and
  three tests pin it, including one that deliberately reproduces the refusal so the fix cannot be
  quietly undone. Nothing that ships to users is affected.

- ✅ **Four mailbox tests that could pass having checked nothing no longer do.**
  Each walks a list the machine supplies - the shared mailboxes it opens, the mail accounts its
  signature settings name, the mail it holds carrying an attachment - and an empty list meant the
  loop body never ran, or the check ran over no items at all, and the test still passed. On a
  machine that is supposed to have those things an empty list means the machine or its settings have
  drifted, so the run now refuses; on a machine that genuinely has none it prints a line saying in
  plain words that nothing was verified and what a green result there does and does not mean. Same
  treatment two other tests already had. Every run now also prints how much each of these actually
  covered, so a pass says what it looked at instead of leaving it to be guessed from the test's
  name.

- ✅ **The leftover-mail check covers every watched mailbox and never deletes where no test may write.**
  After every run against real mailboxes the suite proves that no test-created mail remains
  anywhere. It only ever looked in the mailboxes tests are allowed to write to - while the separate
  guard that watches for mail going MISSING watches those plus every shared mailbox and every
  reference mailbox the profile opens. So a stray test mail landing in somebody else's mailbox was
  invisible to both: one of them only looks for mail disappearing, and this would be mail appearing.
  The check now visits exactly the same set. In any mailbox no test may write to it counts and
  reports rather than deleting, and a count above zero fails the run naming the mailbox and saying
  why nothing else would have told you - because a delete aimed at those is the shape that once came
  within one count of destroying twenty thousand items of reference mail. Nothing that ships to
  users is affected.

- ✅ **One mailbox test no longer blames Windows Search for its own query running out of time.**
  The test creates a mail with three kinds of attachment and then waits up to 90 seconds for Windows
  to index it, asking every 5 seconds - but each question was allowed to take 60 of those 90
  seconds, so a single slow one consumed most of the wait and the test then reported that the
  indexer had not got round to the mail. Each question is now bounded at 15 seconds; one that runs
  out costs a single missed attempt rather than failing the run, and the run reports how many
  attempts it actually got. A real failure of the search provider still fails the test, told apart
  from a question that simply ran long. And if every attempt is lost the test says in plain words
  that it proved nothing, instead of stating something about the indexer it never measured. Nothing
  that ships to users is affected.

- ✅ **Three mailbox tests that could not have run at all declare what they need, and can run.**
  A safety check added earlier makes the test client refuse to call the tools that always reach
  Outlook unless the test has declared that it means to - and three tests that do mean to never made
  the declaration, so they would have failed on their first call. They are in the tier that no
  automated run executes, which is why nobody found out. Fixed, and the check that catches the
  opposite mistake now catches this one too.

- 📦 **The mail server's shared library builds again for the .NET Framework target the add-in reuses.**
  Two pieces of code added the day before used APIs that target does not have. Nothing you can see
  changes; the build gate that exists to catch exactly this had caught it.

- 📝 **The rebuild notes say where the Office installation media really lives.**
  The rebuild notes pointed at the maintainer's Downloads folder, which was purged without warning
  once already; the media is in the repository's own scratch directory, which nothing outside this
  project prunes. The notes now say so, and say plainly that it is a precondition someone has to put
  back if that directory is ever cleaned.

- ✅ **The suite's final cleanup never points a delete at accounts no test may write to.**
  That cleanup walks every account the run watches and deletes anything left carrying the suite's
  own tag - and the list it walks deliberately includes the 20,000-item measurement mailbox and the
  untouched account the suite watches as its evidence that nothing strayed, because an account has
  to be on that list to be counted at all. Those two carry a different tag now, so nothing matches;
  what kept them safe was the data rather than the code. They are counted and never deleted from,
  every run states in writing which accounts it left alone, and a tagged item found in one fails the
  run naming the account and how many - which nothing else would have reported, because the guard
  that watches for lost mail only fires when items disappear and this is items appearing.

- ✅ **Each test machine names its own synthetic mail population, so two no longer overwrite one record.**
  That record is the only thing that can remove such a population from a real mailbox afterwards,
  and it is also what says whether the population is still recent enough to measure against - yet
  every machine's results are collected into one shared folder, under a file name taken from the
  population rather than from the machine. With one machine that was harmless; with two, the second
  collection silently replaced the first. The two machines being built now have distinct names
  chosen to say which is which, and the collection step refuses to replace an existing record whose
  contents differ unless explicitly told to, explaining what it is protecting instead of stopping
  without a reason.

- 📝 **The test-machine notes say what really happens when the test machine's Office licence runs out.**
  They claimed Office drops into "reduced functionality" - a term that does not exist for this
  edition of Office, borrowed from a much older product and from a different kind of licence
  entirely, and one whose view-and-print restriction has never applied to Outlook. Measured on the
  test machine a week after its licence expired: Outlook still runs and everything this project asks
  of it still works. What actually happens is nagging and a red title bar. The correction matters
  because the false version made a monthly rebuild look like a hard deadline. The notes now record
  the real reason to check the licence before a long test run - it is a one-off quarter-second
  question worth asking, and the one genuinely unknown risk, whether an activation prompt can hang
  Outlook as it starts, is written down as unknown rather than guessed at. The exact way to read the
  licence state is recorded too, including the three easy ways to get it wrong.

- ✅ **Each test machine's measurements and logs go to a folder of its own, no longer overwritten.**
  Comparing the two machines is the entire reason there are two, and those files had fixed names in
  one shared folder, so every second collection overwrote the first and the loss looked like nothing
  at all. The record that can remove a synthetic mail population deliberately stays in the shared
  folder: it is named after the population rather than the machine, and keeping it there is what
  makes two populations wrongly given the same name collide where a person can see it. Collecting
  now also refuses to replace ANY file whose contents differ, not just that one, and says which kind
  of loss it is preventing. Files left in the old shared folder by an earlier collection are named
  on every run and never moved, because which machine produced them was never recorded and a guess
  would be worse than a gap.

### Security

- 🔒 **The draft tools refuse an attachment path that crashes Outlook or that Outlook will not attach.**
  They used to hand Outlook any absolute path to a readable file, and a `\\?\` path crashes
  OUTLOOK.EXE inside `Attachments.Add`. On a test machine it did so four times out of four, whatever
  the file was called, so one call to `new_draft`, `reply_draft`, `replyall_draft`, `forward_draft` or
  `update_draft` could crash your Outlook. Outlook also refused a file named after a Windows device
  (`CON.txt`), an alternate data stream (`file.txt:stream`) and a path of 304 characters, and the
  draft was made without the file. These are now refused before Outlook is asked anything, and the
  rest of the call with them: a `\\?\` path, with the ordinary path to give instead; a file named
  after a device; a `:` that is not a drive letter's; and a path longer than 259 characters.
  `\\.\C:\...` and `\??\C:\...` attached the right file and are still accepted, and any other device
  path, such as a pipe, is refused. Each refusal says what to do instead, and so does the
  `attachments` description of each draft tool. The rules about Windows names are shared with
  `save_attachment` and `manage_signature` (`WindowsFileNames`). On the test machine the refused calls
  made no draft and left Outlook running, and ordinary files still attached, a path of exactly 259
  characters among them (Docs/live-tier-on-the-vm.md section 4.11).

- 🔒 **`save_attachment` saves only into OutlookAI's own download folder and folders inside it.**
  It used to save into any folder the assistant named. The sender of a mail chooses an attachment's
  content and its name, and the assistant that reads the mail chooses the folder, so a mail could
  talk it into saving its attachment in the Startup folder, where Windows opens it at the next
  sign-in. Now it saves in `%LOCALAPPDATA%\OutlookAI\scratch\attachments`, or in a folder inside it
  named by a relative path such as `invoices`, and refuses anything else with nothing written. The
  refusal tells the assistant to copy the file elsewhere with its own file tools, under their own
  permission rules (decision 198 (a)), and the tool's description says the same. A folder is judged
  by where Windows really puts it. `..` is resolved; network and device paths (`\\server\share`,
  `\\?\`, `\\.\`), a `:` that is not a drive letter's, and folder names Windows would change are
  refused on their text; junctions, symbolic links and 8.3 short names are followed to the folder's
  final path before it is compared. The server checks the folder before Outlook is asked anything,
  its Outlook process checks it again before it opens the mail, and the writer checks it before and
  after it makes the folder. The sender's file name is made safe: separators, `:` and the other
  characters Windows cannot store become `_`, a device name such as `CON` gets a leading `_`,
  trailing dots and spaces go, and so do the Unicode controls that can make `exe.pdf` read as
  `fdp.exe`. A file is never replaced. The name is claimed by creating an empty file that must not
  exist yet, without following a link of that name, and a taken name gets ` (2)`, ` (3)` and so on.
  All of the tool's file writing is in one class, `AttachmentFileWriter`. A test machine showed it
  end to end, through the shipped server too: Outlook wrote each attachment into the file claimed
  for it, a name holding a direction control was saved under its safe name, a name too long for the
  folder was shortened with its extension kept, and a junction in the download folder that led out
  was refused by the server and again by its Outlook process. `manage_signature` now also refuses a
  device name before a dot, such as `nul.backup`, which Windows reads as a device too.

- 🔒 **The installer removes the trust Office saved for add-ins signed with the leaked old key.**
  Clicking Install on Office's "Publisher cannot be verified" prompt for OutlookAI v1.0.0.1 to
  v2.3.1.111 saved a trust entry under `HKCU\Software\Microsoft\VSTO\Security\Inclusion`: where
  the add-in was, and the public key that signed it. From then on Office loads any add-in at that
  place signed with that key without asking. It never looks at the certificate, and that key's
  private half is public, so anybody can sign with it. Every install, update and uninstall now
  removes each entry that holds the leaked key, found by the key alone and wherever it points.
  Entries for other keys stay, among them the one for the current release. Microsoft documents the
  entry as a location and a public key, and the runtime's own code compares those two and nothing
  else. On a test machine the installer and its uninstaller each removed three planted entries for
  the leaked key, spelled three ways, and kept every other entry (Docs/live-tier-on-the-vm.md
  section 4.8).

- 🔒 **The installer removes the leaked old signing certificate from every certificate store it can change.**
  That certificate signed OutlookAI v1.0.0.1 to v2.3.1.111, and its private key became public, so no
  computer should trust it. Installing or updating already removed it from your Trusted Publishers
  store. Every install, update and uninstall now also removes it from your Intermediate
  Certification Authorities and Trusted People stores, and from the computer's stores too when the
  installer runs with administrator rights. It is found by its thumbprint,
  E205060633DD7062D4F90033130542948A69D068, and nothing else. Your Trusted Root Certification
  Authorities store works differently: Windows asks you to confirm every removal from it, in a
  window titled Root Certificate Store whose default answer is No, and the removal waits for your
  answer; on a test machine it was still waiting after two minutes. So when the certificate is in
  that store, an install or uninstall you run yourself first says why in one sentence and then
  waits for your answer, and an automatic update leaves it for the next install you run. A test
  machine showed all of this (Docs/live-tier-on-the-vm.md section 4.8). The installer carries no
  copy of the certificate and never adds it to the Disallowed store (decision 193 (b)).

- 🔒 **The add-in and the MCP server load Windows DLLs only from System32.**
  Each of their P/Invoke declarations names a Windows DLL by file name alone, and Windows looked for
  such a DLL in the program's own folder first: for an installed copy, the per-user install folder
  under `%LOCALAPPDATA%`, which any of the user's own processes can write. Every assembly of the
  add-in and the server now carries `DefaultDllImportSearchPaths(System32)`. The .NET SDK's security
  analysers found it (CA5392) the day they were turned on.

- 🔒 **A send confirmation token no longer outlives its two-minute limit when the clock moves backwards.**
  That can be a time sync, a daylight-saving change or a virtual machine resuming. The token's
  expiry, and every other timer inside the mail server, is now measured with a clock that only moves
  forward.

## v3.1.0.325 - 2026-08-15

- Make registering the mail server with Claude Code your choice instead of something that just happens: OutlookAI no longer adds itself to your personal Claude Code configuration on its own. A "Make available in all my Claude Code projects" tick box in OutlookAI Settings turns it on, and unticking it removes the entry again - which is also the tidy way to take it out before uninstalling. If OutlookAI was already registered before this update it stays registered and the box starts ticked, so nothing you already had working changes.
- Add the mail server to a single project instead of all of them: a new button in OutlookAI Settings lets you pick a folder and writes a `.mcp.json` there, so the server is available in that project alone. The entry is written with a portable path rather than one pointing inside your own user folder, so the file is safe to commit - a colleague who has OutlookAI installed gets a working server, and one who does not simply sees it fail to connect rather than anything breaking. Claude Code asks you to approve the server the first time you open that folder; that is its own security prompt and is meant to happen. There is also a Copy CLI command button if you would rather add it by hand.
- Fix a mail-server registration you wrote yourself being overwritten every time Outlook started: an entry that left out the optional type field was treated as damaged and rewritten on each start. Such an entry is now recognised as valid and left alone.
- Fix folder-scoped searches being able to run on far longer than intended: sweeping a folder and its subfolders counted only the folders it successfully read against its own limit, so calendar, contacts and unreadable folders were walked for free and a deep or wide mailbox could keep a search going with no bound at all. The sweep now counts every folder it visits, stops at a sensible depth and time limit, and says in its result when it stopped early rather than quietly returning less.
- Fix the setup instructions for the Claude Code CLI sending you to the wrong installer: OutlookAI looks for the CLI where the official installer puts it, but the instructions told you to install it with npm, which puts it somewhere else entirely - so the writing sidebar reported Claude Code as not installed even though `claude` worked perfectly in your terminal. The instructions now point at the official installer, the Node.js requirement is gone (it was only ever needed for the npm route), and when the CLI still cannot be found the message names the exact location OutlookAI looked in and says plainly that your PATH is not searched.
- Ship the Office runtime inside the installer instead of downloading it: the Visual Studio Tools for Office runtime OutlookAI needs on a machine without Office is now carried in the installer itself, so setup no longer depends on a Microsoft download address staying valid - which is exactly what broke. The installer is larger as a result, and still comfortably small enough for automatic updates to keep working.
- Explain the security prompt you see the first time Outlook starts after installing: Windows asks you to confirm OutlookAI's customization and shows the publisher as unverified, because OutlookAI is signed with our own certificate rather than a commercial one. Clicking Install is expected and happens only once on a machine. The README now describes the prompt, gives the certificate fingerprint so you can check it yourself before accepting, and explains how to recover if you clicked "Don't Install" and the add-in never appeared.
- Fix setup silently failing to install the Office runtime it depends on: the Microsoft address setup downloaded the Visual Studio Tools for Office runtime from stopped being a file and became a web page, so setup saved that page as if it were the installer and ran it, and the runtime was never installed. Setup now downloads it from a direct address, and every prerequisite it fetches is checked to be a real Windows program before being run - so if an address breaks like this again you are told, instead of setup appearing to succeed. This only ever affected machines with no Office installed, because Office brings that runtime with it.
- Remove the .NET Framework 4.8 download from setup: every Windows version OutlookAI supports has included it since Windows 10 1903, so that download could never actually run. If it is somehow missing, setup now tells you and points you at Windows Update instead of quietly trying to fetch it.
- Clean up after yourself when uninstalling: uninstalling now removes OutlookAI's own registry settings - your Outlook tuning preferences and the mail server's registration state - instead of leaving them behind for good. Your Outlook configuration itself is still left exactly as it is, as it always has been. One consequence worth knowing: because the tuning preferences go too, uninstalling and reinstalling starts you from the defaults again.

## v3.0.1.321 - 2026-08-08

- Fix the add-in not loading at all after installing: since v2.3.3.141 the installer put the add-in's files in a folder layout Outlook could not load them from, so an installed OutlookAI silently disabled itself on the first Outlook start and left you with no ribbon button and no writing sidebar. The installer now places them where Outlook looks, and an install is checked at build time so this cannot ship again. If you are on v2.3.3.141, v2.3.4.145 or v3.0.0.319 you must install this update by hand - a disabled add-in can never update itself.

## v3.0.0.319 - 2026-08-08

- Add the OutlookAI MCP server: a local server that lets AI agents such as Claude Code work with your mail - it searches all locally indexed Outlook mail (all accounts, delegate mailboxes and attachment contents) in milliseconds.
- Install the mail server with the add-in: it now ships in the same installer, under a McpServer folder next to the add-in, and both carry the same version so one update covers both. The server needs the .NET 10 runtime (the base runtime, not the desktop one), which the installer downloads and installs for you like it already does for .NET Framework 4.8 and the VSTO runtime. As with those, an automatic background update does not install prerequisites - it runs unattended after Outlook closes and cannot ask you for permission - so if the runtime is missing OutlookAI Settings says so and links to the download; the add-in itself is unaffected either way.
- Register the mail server automatically, and keep it registered: on every Outlook start OutlookAI points Claude Code at the installed server and repairs the entry if it has drifted, so there is no setup command to run any more. Your Claude Code configuration is treated as the file it is - a configuration that cannot be read is never rewritten, everything unrelated to OutlookAI is left byte for byte as it was, the previous file is kept as a backup, and an entry that is already correct is not touched at all. If Claude Code is not installed, or the .NET 10 runtime is missing, nothing is changed and the reason is shown in OutlookAI Settings. The health report also states whether the registration matches the server that is actually running.
- Replace running mail-server processes cleanly on update: an agent session starts its own copy of the server, so several are usually running when an update arrives. Setup now stops the ones it is about to replace before replacing them; a session simply starts a fresh copy on its next mail request.
- Make every search always fresh: results automatically include mail that arrived seconds ago, sweeping the Inbox, Sent Items, Deleted Items and Junk Email - or, when you search one folder, that folder and its subfolders - so mail a mail-server rule files straight into another folder is found immediately instead of only once the search index catches up. The sweep is cached for about 10 seconds so rapid repeated searches stay instant, falls back to index results with a clear freshness warning when it cannot run (for example during an add-in update) instead of failing, and every result reports which folders it covered.
- Match search words across the subject and the body: a word is found in either, and in a multi-word search each word may match in a different one, so subject-phrased terms (alert prefixes, ticket tags, invoice numbers) and mail with one word in the subject line and another in the message text are no longer missed. A search_in option narrows a search to just the subject or just the body when a word is noisy in the other.
- Search inside every kind of attachment, not just documents: images, forwarded .msg/.eml messages, calendar invites, audio and video are now searched too, so a mail can be found by what is written inside anything attached to it, in shared mailboxes as much as in your own - and searching got slightly faster rather than slower.
- Add an include_subfolders option to mail search, on by default: a folder search covers that folder and everything under it, and switching it off searches that one folder only - identically in all three search paths, so the exhaustive scan of a folder is never quietly narrower than a normal search of it.
- Add exhaustive search mode, a true/false option on the same search tool: a folder- or date-bounded scan straight through Outlook that bypasses the search index, for when the index is stale or correctness matters more than speed. Searches advise when the index is stale enough to warrant it, and the scan honours its 2-minute limit per folder and reports how far it got instead of running on and still claiming success.
- Fix mail in a shared (delegate) mailbox's subfolders being effectively out of reach: searching a path like "Inbox/2025 invoices" there returned nothing at all, with no error, as if the folder were simply empty, and mail found in such a folder could not be read, opened in Outlook, followed as a thread or have its attachments saved. Both now work, and two consequences of how shared mailboxes are indexed are reported rather than hidden: a search that cannot be narrowed to the subfolders you asked for covers the whole shared mailbox and says so, and if that mailbox has two folders with the same name the results may include both.
- Fix a search for words inside attachments also returning ordinary mail that merely matched on subject or body: such a search is now purely attachment matches, and the assistant is told that the freshest mail is not covered by it until that mail has been indexed.
- Fix mail search failing outright on a folder whose name contains an apostrophe (for example "O'Brien").
- Fix following the conversation of a mail that belongs to no conversation failing with a raw error: such a mail is now simply reported as having no conversation to follow.
- Explain what is actually wrong when a search result cannot be opened - the item's folder is gone from Outlook, the mailbox is not open in this profile, or the item was moved or deleted after it was indexed - each with the step that actually resolves it, instead of always suggesting a retry that cannot help.
- Report every way a result can be incomplete instead of letting it look complete: a folder whose mail Outlook would not read counts as failed rather than checked, the freshness sweep names its per-folder and folder limits and any omitted or skipped folders, a folder search that finds nothing says whether the folder path resolved at all, and search, conversation, recipient and attachment lists say when a cap cut them - with has-more markers and guidance to raise it or narrow the query, while operations keep using the full lists. The assistant is also told plainly how search behaves: what search words are matched against, that sender and recipient are matched by the separate from/to filters, which follow-up actions a result can be used for, and which searches cover subfolders and which do not.
- Add show-me tools: the assistant can open a mail in an Outlook window for you, jump your Outlook to any folder, and run a search in Outlook's own search box so you see the result list on screen (works even if Outlook was not running).
- Warn when Outlook's own search is running online (server-assisted): showing search results in Outlook now says the on-screen list may differ from what the assistant itself finds - online results are capped and ranked differently - and recommends switching local search back on, the health report states which search backend Outlook's search box is actually using, and the Search group in OutlookAI Settings spells out what turning local search tuning off costs.
- Add mail reading: full message read with safe truncation, sender and recipient details, conversation (thread) view, and saving attachments to disk so the assistant can open them. Very long mails can be read in consecutive windows that continue exactly where the previous one ended, and a mail's stored HTML can be returned - the only way to see formatting, where the signature starts and where the quoted conversation begins - which also works on a draft directly, without the search index.
- Add mailbox insight: account and store listing (delegate and online-only stores flagged), and the complete folder tree of a mailbox with unread counts in a stable order, with paging in the unlikely case a profile exceeds 1000 folders.
- Add an outlook_health tool that reports the assistant's whole Outlook picture in one call, without ever starting Outlook: Outlook state and installed version, whether it is running invisibly in the background (tray icon only) or as a normal window, connection health verified at report time rather than inferred from a stale session, mail store reachability, search-index freshness overall and per mailbox with advice, Windows Search service state, audit-log writability, and the current Outlook tuning state.
- Add email drafting: the assistant can prepare a new mail, reply, reply-all or forward for you - the draft opens on screen with the right account identity, that account's own signature and the assistant's text above the quoted conversation, ready for you to review and send yourself (nothing is ever sent automatically). Drafts are built the way Outlook's own compose window builds them, so your text sits above the account's real signature (HTML, logos and all) and keeps the message's own font and layout, signature images are embedded in the mail itself so they survive revisions and arrive intact for the recipient, an account without a configured signature still gets a body-only draft, and a draft written while Outlook has no window open is identical to one written with Outlook on screen - with nothing appearing on your screen while it happens.
- Add signature steering: the assistant can list your installed email signatures (with a short excerpt so it can tell their language and purpose apart) and apply a specific one to any draft it prepares - for example a signature matching the recipient's language - while omitting the choice keeps the account's default; per-account defaults are reported where Outlook records them.
- Add signature management: the assistant can create, update and delete Outlook email signatures (writing all three formats - HTML, plain text and RTF - and deriving whichever you did not supply) and optionally record one as an account's default for new mail and/or replies. Previous signature files are automatically backed up under %LOCALAPPDATA%\OutlookAI\signature-backups before any update or delete and the backup location is reported, and deleting a signature also cleans up default assignments that pointed at it.
- Add a "Select the best signature" button to the AI writing sidebar: the AI looks at your draft, the quoted thread and the recipients, picks the most fitting of your installed signatures (for example by matching the language) and applies it, leaving your draft text and the quoted conversation untouched. With a single installed signature it is applied directly without an AI call, and the button is available only when at least one signature exists.
- Add formatted drafts: the assistant can write a message as real HTML - headings, bold, bulleted and numbered lists, tables, links and inline styling - so a proper letter finally comes out looking like one instead of a wall of plain text, with the formatting confined to your message and the signature and quoted conversation left exactly as Outlook made them; plain text remains the default. Anything a mail cannot safely carry (scripts, style blocks, embedded frames, images, tracking attributes) is removed, unknown tags are unwrapped with their text kept and half-finished markup is repaired rather than rejected, so a small mistake in the assistant's output can never mangle the message or swallow your signature.
- Add Cc and Bcc to every draft the assistant creates - on replies and forwards the addresses are added to the recipients Outlook already filled in rather than replacing them, so a reply-all keeps everyone - and any address Outlook cannot recognize is reported back with the draft instead of being silently dropped.
- Add a subject override for reply, reply-all and forward drafts, for when a thread needs renaming: the draft stays part of the original conversation, still grouping with the thread in Outlook and keeping its reply chain, which Outlook does not do on its own when a subject is changed.
- Add importance (low / normal / high) and a read-receipt request to all four draft tools.
- Add file attachments to drafts: the assistant can attach files from anywhere on the machine to any draft it creates or revises, and remove them again by name. Every path is checked before anything is written, so if one file is missing, unreadable or a folder, nothing is attached at all and the answer names each problem file - instead of quietly sending a mail with a file missing - and the result reports which files ended up on the draft, with their sizes.
- Add draft revising: the assistant can rewrite the text, rename the subject, change who it goes to, switch the signature or add a file on a draft it already made, instead of starting a second one. Only your message text is replaced - the signature and any quoted conversation stay put, a renamed subject keeps the draft in its original conversation, recipients are replaced rather than added to so a wrong address can finally be taken off, and anything not passed is left untouched; only unsent drafts in a Drafts folder can be revised (anything else is refused with a clear reason and left alone), and a draft whose signature image is still only a link to a file on this machine cannot be repaired in place, so revising it now says so and asks for the signature to be re-applied instead of losing the image quietly.
- Add draft discarding, so a draft that turned out wrong does not have to be cleaned up by hand. It is deliberately narrow: only an unsent draft the assistant itself created in the same session and still sitting in Drafts can be thrown away - mail you received, mail you wrote, anything already sent, anything outside Drafts and drafts from an earlier session are all out of reach - and it works exactly like pressing Delete in Outlook, moving the draft to Deleted Items with the result saying how to put it back, never deleting permanently or emptying anything.
- Add a deliberately high-friction send tool: the assistant can send a saved draft only via a two-step confirmation - a one-time token bound to that exact draft and its current content, expiring after two minutes and voided by any change to its text, recipients, attached files or formatting - with the sending account hard-verified before transport and every step audit-logged; drafting for you to review and send yourself remains the default path.
- Report a draft honestly instead of quietly handing back a worse one: the attachment sizes reported are the ones the recipient will actually receive, and the store, folder and sending account no longer come back empty when Outlook has only just started or is momentarily unresponsive. If a draft could not be written the proper way it is still created and nothing is lost, but the answer now names what that costs - the message layout, a requested signature, embedded signature images - and how to get the full result; when nothing went wrong, nothing is reported.
- Add mail moving: the assistant can move mail (1-50 at a time) to another folder within the same account - refiling into project folders or restoring items from Deleted Items - creating the target folder on request. Every move is audit-logged and reversible (results carry the source folder and the item's old and new ids), and moving to Deleted Items or the Outbox is refused: the assistant still cannot delete mail.
- Add one-click archiving: the assistant can archive mail exactly like Outlook's own Archive button (Backspace), with each item going to its own account's designated Archive folder, resolved per mailbox even when it has a localized name (for example "Archiveren"). Accounts without a designated Archive folder get a clear error and nothing is created; archiving is audited and reversible like any move.
- Add a structured audit log: every draft the assistant creates and every attachment it saves is recorded locally under %LOCALAPPDATA%\OutlookAI.
- Add automatic Outlook tuning: OutlookAI now keeps your proven Outlook configuration applied on every start - fast local search settings, full mailbox caching (sync slider = All, shared folders included), and raised OST size limits so a fully cached mailbox never stalls. Only actual differences are written, values enforced by your organization's group policy are respected and flagged, and you are told when an Outlook restart is needed for a change to take effect.
- Add an OutlookAI Settings dialog: a large button in the OutlookAI group on Outlook's main Mail ribbon (carrying the same icon as the AI Assistant) opens a settings window matching your Office light or dark theme, where each tuning group can be switched on or off, the current effective values are shown, and a clear indicator tells you when an Outlook restart is still needed. Turning a group off stops managing it and leaves your Outlook settings as they are.
- Guarantee background operation: only the explicit show-me tools (open a mail, jump to a folder, show search results) and draft windows ever open Outlook UI - every other assistant operation now provably leaves Outlook window-less, verified by tests.
- Keep the assistant and Outlook in step: the server watches the Outlook process and releases every held connection the moment Outlook exits (user close, crash or logoff), so a dead Outlook is never held open by stale references and the next mail request starts Outlook fresh without errors. An Outlook exit in the middle of an operation is handled as a recoverable condition instead of a raw error, Outlook is kept running rather than exiting after each draft, and a search that started Outlook itself no longer reports it as not running.
- Add mail-server usage instructions that load into every Claude session automatically, so the assistant knows it can search and read your Outlook mail (and that drafts open for review while sending stays gated) even before any mail tool is used.
- Update the README with the MCP server and Settings documentation and with Outlook's lifetime around the assistant - headless background start, the tray icon meaning, the roughly 10-12 minute self-exit after the last agent disconnects, promoting it to a normal Outlook by just launching it, and what closing the window does - and add developer documentation for the MCP server.
- Remove a superseded OutlookAI signing certificate from your Trusted Publishers store during install and update — its private key was briefly exposed publicly, so it should no longer be trusted. Current releases are signed with a different, unaffected certificate.

## v2.3.4.145 - 2026-07-22

- Fix: AI error responses (such as hitting the turn limit) are no longer written into your draft as if they were the reply — you now get a clear message instead.
- Fix: the assistant no longer gets stuck with its buttons disabled if the email window is closed while a request is in progress.
- Fix: a momentary "couldn't access the email editor" no longer causes your next action to silently blank the draft.
- Fix: the AI now always uses the current email's signature and quoted thread as context (previously it could reuse a previous email's context when Outlook recycled a window).
- Fix: if writing the AI result into the email fails partway through, the signature and quoted-thread markers are now restored instead of being lost.
- Fix: setup or sign-in problems now surface a clear message quickly instead of retrying silently in the background.
- Improve: the assistant now ignores instructions embedded in quoted/original email text (the random "fences" added previously are now actually described to the model as untrusted data).
- Improve: downloaded updates are now genuinely verified — valid Authenticode signature (via WinVerifyTrust) plus a pinned publisher-certificate thumbprint — before they run.
- Fix: the auto-updater no longer gets permanently stuck after a malformed GitHub response, no longer leaves partial installer files behind, pins the installer asset by name, and shows clearer status.
- Fix: uninstalling now correctly removes the OutlookAI signing certificate from the trusted store.
- Fix: resolved several memory/handle leaks (email editor, text selection, inspector/explorer windows) and a duplicated inspector COM release.
- Improve: the assistant follows Windows/Office light–dark theme changes live while a pane is open, and the White Office theme is now correctly shown as light.
- Fix: AI panes opened for inline replies are cleaned up when that window closes, instead of lingering hidden (with a background timer) until Outlook exits.
- Change: update checks retry on every Outlook start rather than giving up after repeated failures.
- Fix: the Claude helper no longer spawns extra background processes under rapid use, and Outlook no longer briefly stalls on shutdown.
- Fix: a COM handle leak when reading the draft, signature, and quoted text.

## v2.3.3.141 - 2026-05-31

- Fix 48 bugs across security, correctness, resource management, and UI
- Add Authenticode signature verification on downloaded installer before execution
- Add installer code signing step to release workflow
- Fix predictable temp file path for update downloads (now uses random filename)
- Fix PowerShell command injection in update launcher via -EncodedCommand
- Fix prompt injection vulnerability by replacing static delimiters with random fences
- Fix COM object lifetime issues in ribbon callbacks that could crash Outlook
- Fix writing AI output to wrong email when user switches during processing
- Fix bookmark offset calculation that could produce invalid Word ranges
- Fix semver arithmetic in release workflow (minor bump now resets patch)
- Fix COM identity comparison using IUnknown pointer equality instead of RCW reference
- Fix duplicate explorer event hookups and task pane lifecycle leaks
- Fix ribbon toggle state not refreshing across all ribbon contexts
- Fix race between inspector Activate and Close handlers
- Fix process tree not being killed on timeout (orphaned Node.js processes)
- Fix overly broad error keyword matching causing false authentication errors
- Fix COM Range object leaks throughout Word document operations
- Fix UI thread race condition in InvokeOnUI that could throw ObjectDisposedException
- Fix released COM inspector reference causing crashes on async operations
- Fix thread safety on update service shared state with volatile fields
- Fix JSON parse errors showing cryptic messages instead of helpful diagnostics
- Fix Office theme detection to check multiple Office versions (15.0, 16.0, 17.0)
- Fix static constructor crash risk in ThemeService with safe fallback
- Fix release workflow creating tag before changelog commit
- Fix release workflow push failure by adding git pull --rebase
- Fix misleading version_bump description suggesting 0.0.0 is valid
- Fix bare catch blocks swallowing all exceptions including critical ones
- Fix unbounded download into memory (now capped at 50 MB)
- Add runtime theme change detection via SystemEvents
- Add Anchor layout to task pane controls for proper resize behavior
- Fix lblStatus truncating long error messages (AutoEllipsis enabled)
- Fix ToolTip GDI handle leak
- Fix timer tick firing on disposed controls by correcting dispose order
- Fix debug mode activating from casual clicks (now requires 7 clicks within 3 seconds)
- Fix version comparison with mixed 3/4-component versions
- Fix dev builds triggering false auto-updates (placeholder version now 99.99.99.0)
- Fix installer not prompting when Outlook is running during manual install
- Upgrade NuGet setup action to v3 for Node.js 24 compatibility
- Fix add-in silently failing to appear in Outlook on fresh systems by replacing VSTOInstaller.exe dependency with direct registry-based add-in registration
- Add on-demand VSTO Runtime download and install with elevation if missing on target system
- Add on-demand .NET Framework 4.8 download and install with elevation if missing on target system
- Add signing certificate to user's Trusted Publishers store at install time to prevent silent trust failures
- Prevent Outlook from disabling the add-in due to slow startup or crashes via DoNotDisableAddinList registry key
- Add uninstall support via Add/Remove Programs, including certificate and registry cleanup
- Fix update race condition where restarting Outlook during an update could cause file lock errors
- Fix auto-update retries getting stuck after a failed update attempt until Outlook restart
- Stop auto-update retries after 3 consecutive failures instead of retrying indefinitely
- Fix potential socket exhaustion from creating a new HTTP client on every update check
- Fix installer UI freezing during prerequisite downloads by running downloads out-of-process
- Surface download error details (DNS, proxy, SSL) in installer error messages instead of generic failure
- Fix concurrent update checks when a download takes longer than the 10-minute poll interval
- Increase auto-update download timeout from 15 seconds to 5 minutes for slow connections

## v2.3.2.134 - 2026-04-09

- Move signing certificate out of repository into GitHub Actions secrets to prevent unauthorized use
- Fix CI release build by switching Office interop references to PIAs (previous attempt to remove dead conditional code broke CI builds)
- Remove redundant build CI workflow (release workflow already covers the same build steps)
- Update default Visual Studio version fallback from VS 2010 to VS 2022 in project file
- Fix VSTO project metadata targeting Office 2013 instead of 2016, matching the actual minimum supported Outlook version
- Remove dead USEOFFICEINTEROP conditional branch from project file (VSTO template boilerplate never used)
- Fix rare double-processing when rapidly clicking action buttons by adding a reentrancy guard
- Fix potential crash from unhandled exceptions in async button handlers and pre-try logic in ProcessAction
- Fix process handle leak when warm-up detects missing prerequisites and exits early
- Fix Outlook UI freezing for up to 1.5 seconds when prerequisites were missing at startup
- Fix non-ASCII email content being garbled when sent to Claude CLI by writing stdin as UTF-8
- Fix potential stdout corruption by attaching async output readers immediately after process start
- Rewrite README with comprehensive documentation of all features, limitations, context awareness, iterative editing, dark mode, auto-updates, debug mode, inline responses, and troubleshooting
- Fix incorrect Unblock-File path in README troubleshooting and clarify that Outlook restart is usually not needed after fixing prerequisites
- Rename installer.iss to Installer.iss for consistent file casing

## v2.3.1.111 - 2026-03-24

- Switch from HTML parsing to Word Object Model for email structure detection, using bookmarks for reliable signature and thread boundary identification
- Preserve email signature and thread formatting across all Outlook versions by never modifying them directly
- Fix upgrade error ("another version is currently installed") by uninstalling previous VSTO registration before installing new version
- Disable built-in VSTO update checks that conflicted with the add-in's own update mechanism
- Enable full feature parity for inline responses (reading pane replies) including selection editing
- Add descriptive tooltips to all action buttons
- Improve quick action prompts for more reliable results on short emails
- Simplify UI by merging Instruction and Custom Action into a single text box with three buttons

## v2.1.1.103 - 2026-03-23

- Fix signature layout being destroyed when drafting emails by switching to HTML-native editing
- Ensure "Draft new email" starts with a clean slate, excluding any previous AI-generated content
- Fix release workflow failing to push version bump by deriving version from release tags instead of workflow files

## v2.0.0.94 - 2026-03-23

- Add dark mode support matching Outlook's theme setting
- Preserve email signature formatting when drafting or editing with AI
- Include signature as context in AI prompts to prevent duplicate sign-offs
- Add "Draft new email", "Edit current draft", and "Edit selection only" buttons
- Add "Run Custom Action on Selection" for targeted edits
- Apply AI results immediately instead of showing a preview panel
- Remove result preview panel, Apply, and Discard buttons
- Split CI into build-only (every push) and release (on demand) workflows
- Update CI actions to fix Node.js 20 deprecation warnings

## v2.0.0.87 - 2026-03-23

## v2.0.0.84 - 2026-03-23

- Add iterative draft editing with full conversation history across multiple AI turns
- Add "Edit Draft" button for refining an existing AI draft without starting over
- Add "Run Custom Action on Selection" button to apply instructions to selected text only
- Preserve email formatting by reading and writing HTMLBody instead of plain text Body
- Detect Outlook reply boundaries to separate draft from quoted thread
- Send email thread context to AI only once, not on every subsequent edit
- Capture manual edits made in the Outlook editor between AI turns
- Replace Insert/Replace buttons with a single Apply button for cleaner workflow
- Fix installer failing to find VSTOInstaller.exe by using correct Common Files path
- Bump major version to 2.0.0

## v1.1.0.82 - 2026-03-23

- Fix update installer showing a popup by using silent VSTO installation

## v1.1.0.79 - 2026-03-23

- Fix inspector COM leak for non-compose windows (read emails, calendar items)
- Fix inspector COM ownership: Close handler releases when no task pane owns it, task pane releases on dispose
- Fix COM leak when CurrentItem or ActiveInlineResponse returns a non-MailItem object

## v1.1.0.77 - 2026-03-23

- Fix double-release race condition on inspector COM object between Close handler and task pane disposal

## v1.1.0.75 - 2026-03-23

- Fix COM object leaks from repeated CustomTaskPane.Window access

## v1.1.0.73 - 2026-03-23

- Fix COM object leaks by releasing all Outlook COM references after use

## v1.1.0.71 - 2026-03-23

- Install updates automatically when Outlook closes without force-quitting
- Show update status in sidebar: downloading, ready, up to date
- Prevent multiple installer instances from running simultaneously

## v1.1.0.68 - 2026-03-23

- Fix update check failing with TLS error on .NET Framework 4.8

## v1.1.0.66 - 2026-03-23

- Fix update error link never appearing when all update checks fail
- Show "checking..." state before first update check completes

## v1.1.0.63 - 2026-03-23

- Check for updates every 10 minutes using conditional requests to avoid API rate limits
- Show time since last update check in the sidebar version label
- Show clickable error link in sidebar when update check fails
- Show progress bar during automatic updates instead of running fully hidden

## v1.1.0.61 - 2026-03-23

- Add automatic silent updates via GitHub Releases on Outlook close
- Show current version at the bottom of the AI sidebar
- Bump minor version to 1.1.0

## v1.0.0.59 - 2026-03-23

- Remove empty release entries from changelog
- Fix CI creating empty changelog headings when no unreleased entries exist

## v1.0.0.57 - 2026-03-23

- Add ribbon toggle button to show/hide the AI sidebar in compose windows and inline replies
- Toggle button reflects sidebar state (pressed when open, unpressed when closed)
- Sidebar auto-shows for each new composition and can be toggled off via button or close control
- Fix sidebar staying hidden when opening a new compose window after closing a previous one
- Fix Claude Code not being found by using its default install path instead of relying on PATH

## v1.0.0.49 - 2026-03-11

- Add AI Assistant sidebar for inline replies in the reading pane
- Auto-show the AI sidebar whenever composing an email (new, reply, forward)
- Remove the ribbon button — the sidebar now appears automatically
- Fix crash when closing a compose window while an AI request is in-flight
- Fix wrong email being processed when multiple compose windows are open

## v1.0.0.1 - 2026-03-10

- Replace Anthropic API with Claude Code CLI as AI backend, using pre-warmed subprocess for zero-latency requests
- Add Custom Action feature for arbitrary user instructions on email content
- Add automated build and release pipeline via GitHub Actions
- Add single-file Inno Setup installer published as GitHub Release on every push to master
- Add assembly version synchronization with release version
- Add changelog with AI-maintained release notes
- Upgrade target framework from .NET 4.7.2 to .NET 4.8
- Remove voice input and OpenAI Whisper integration (users dictate via external tools)
- Remove configuration UI and API key management (no longer needed with Claude Code CLI)
- Remove enterprise RDS deployment scripts in favor of single-user ClickOnce install
- Simplify project structure: flatten directory layout, remove dead code and unused references
