# Q149 (d): how OutlookAI should write draft text - the measurement

**The question.** OutlookAI writes the text of every draft through Outlook's Word editor, in a compose
window nobody sees. In an Outlook that a program started without a window, that window's close can end
Outlook, which is why a session holds a hidden "lifetime pin" while it drafts. On 2026-11-17 the
maintainer chose to measure first (Q149 (d)) and then choose between:

- **(b)** keep the Word editor, with pins made only while drafting and the pin lifecycle work around
  them (a shared session list, a lock, taking over a crashed session's pin, the last session out
  closing it); or
- **(c)** write the text as HTML through the object model: no compose window, no activation, no pin,
  ever, with OutlookAI doing itself what Word does today - Outlook's mail stylesheet, the signature
  with its pictures and its `_MailAutoSig` marker, the text above a reply's quoted original, and
  update_draft's replacement.

(c) replaces Word only if its drafts measure as indistinguishable from today's, both opened in Outlook
and as received. Otherwise (b).

## Sources

- The two plans this one merges, both written 2026-10-11 and both unrun: the pins plan (is the pin
  needed, and when) and the Word-editor plan (can the editor be had without a window; can an edit be
  committed without Activate; what would the HTML route lose).
- The 2026-07-31 probe that the D49 design rests on, recovered from its session transcript:
  `probe-2026-07-31/` here, with identifiers removed. Its sample sizes and the corrections to how it
  was recorded are in [Corrections to the D49 record](#corrections-to-the-d49-record).
- The Q118 pin measurements of 2026-10-03, raw logs in `../q118-pin-measure/`, summarised in
  `McpServer/Docs/com-host.md`.

## The staged plan

One merged plan, decisive cells first. Every cell runs on an Outlook test guest (Office LTSC 2024,
16.0.17932) through `Testbed/host/Invoke-LiveTierOnGuest.ps1` - first `OutlookAI-Unindexed`, then
`OutlookAI-Indexed` - as T2 live tests. Office LTSC 2021, which the 2026-07-31 probe measured, cannot be
measured again: no live test runs on the maintainer's workstation (Q116 (a)).

### Stage 1 - an Outlook with no window

Each trial ends Outlook the S7 way and cold-starts it over COM: `OUTLOOK.EXE -Embedding`, no Explorer,
no Inspector, nothing on screen - the state an agent session meets when it is the one that starts
Outlook. Class `T2/LiveQ149WindowlessComposeTests`.

| Cell | What it does | n | From |
| --- | --- | --- | --- |
| **L1** | The shipped compose (`ReviseHeldDocument`, reached through update_draft) in a session with **no lifetime pin**. Does closing the never-shown compose window end the window-less Outlook? | 20 | pins A1 (decisive) |
| L1c | The same with the product's pin - the control | 5 | pins A4 |
| L2 | An inspector fetched on a re-opened draft, never activated: the editor read once (HRESULT), the item's HTML before and after the inspector touch, `Close(olDiscard)`, Outlook's lifetime | 10 | pins A2/A3, editor A1, C1 |
| W2 | As L2, the editor re-read every 250 ms (pumped) for 30 s; never a second GetInspector | 10 | editor A2 |
| W3 | As W2, the first read 180 s after the start | 3 | editor A3 |
| W4 | The inspector's own window parked off-screen, shown by Windows (`SW_SHOWNOACTIVATE`), hidden again - no Outlook `Activate` - then read | 10 | editor A4 |
| W5 | The same Windows-level show of the pin Explorer's window, then the inspector read | 10 | editor A5 |
| W6 | The product's promotion on the first draft of a process, then nine more drafts read and committed with no Activate; then a displayed Explorer opened and closed, and one more read | 3 x 10 | editor A6, A6b, B1 after a warm-up |

L1 goes through update_draft rather than new_draft for one reason: new_draft sets the subject after the
compose, so an Outlook that ends at the compose's close would leave an untagged draft behind (mailbox
rule 2). The inspector work - `GetInspector`, the promotion, the edit, `CurrentItem.Save`,
`Close(olDiscard)` - is the same method either way.

### Stage 2 - committing an edit on a re-opened draft without Activate

Since 2026-10-04 every compose works on the draft re-opened by its EntryID, and an edit on such a draft
did not commit without `Inspector.Activate()` (D46: 3 of 3; runbook 4.6 variant m30c: 400 of 400 on
Office LTSC 2024). In a windowed Outlook, 20 drafts per variant, each re-read through a separate client:
the new text present and the old text gone. Class `T2/LiveQ149CommitTests`.

| Cell | Commit | n |
| --- | --- | --- |
| B0 | The product's update_draft, with its Activate - the control | 20 |
| B1 | Editor read without Activate; `CurrentItem.Save()`; `Close(olDiscard)` | 20 |
| B2 | Editor read without Activate; `Close(olSave)` | 20 |
| B3 | `ModifiedFormPages.Add("Message")` first; `Close(olSave)`; then the message class and the form-definition properties | 20 |
| B4 | Word's `Document.Save()`; `Close(olDiscard)` | 20 |

### Stage 3 - the HTML-only draft against today's

The same body and the same signature (the test signature, with a picture) composed both ways and
compared. The HTML-only draft borrows today's whole `<head>` - its best case; what an item made with no
inspector holds of its own is recorded beside it. In a windowed Outlook. Class
`T2/LiveQ149HtmlOnlyParityTests`.

| Cell | What is compared | n |
| --- | --- | --- |
| C2 | A new mail: the stylesheet, the body inside `WordSection1`, the signature with its pictures, Word's layout when opened | 5 pairs (3 HTML bodies, 2 text) |
| C1 | What a new item made with no inspector holds: a stylesheet, `WordSection1` | with every C2 pair |
| C4 | `Reply()` and `Forward()` with no inspector: the quoted original, its header, the stylesheet; then the HTML-only reply and forward against today's | 3 + 3 |
| C3 | update_draft: Word's update against the HTML-only update, on a Word draft and on an HTML-only draft, and Word's update on an HTML-only draft | 2 sets |
| C5 | Both sent through the guest's loopback mail sink with the product's own two-step send, to the hub itself, and the received copies compared | 2 pairs |

"How it looks when opened in Outlook" is Word's layout of the item, paragraph by paragraph: the text,
the style, the font and size, bold and italic, the colour, the alignment, the spacing before and after,
the line spacing, the indent, and the pictures with their size - what Word renders in the compose window
and the reading pane.

### Decision rules

- **D1** (the pins plan's). L1 exits in 18 or more of 20: the pin is required for as long as drafts go
  through an inspector. L1 survives in 18 or more of 20: the pin is unnecessary for the shipped compose.
  Anything between: timing-dependent, keep the pin.
- **D2** (the Word-editor plan's phase A, in its order). W2 or W3 at 9 of 10 or better with no window
  shown: replace the window with a bounded re-read. Else W4 or W5 at 9 of 10: one Windows-level show
  per process, no Outlook Activate. Else W6's direct reads at 9 of 10: promote once per process, not per
  compose. Else the current design is justified on this Office, and only (c) is window-free.
- **D3** (phase B). A variant that commits every draft with no side effect lets a windowed Outlook drop
  the per-compose Activate.
- **D4** (the decision itself). (c) replaces Word only if every compared aspect of stage 3 is the same,
  opened in Outlook and as received. Otherwise (b).

### What is not run, and why

- The editor plan's A7, an in-process read through the add-in: whether the add-in loads in a COM-started
  Outlook is unmeasured, and the guests' add-in install is outside this stream.
- The pins plan's A5 (release without closing, watched 15 minutes), A6 (`MailItem.Close`), A7 (Display
  plus a user's close), A9 and A10 (a second COM client during A1 and A2; Quit from it) and A11 (the exit
  delay after the last release, watched 30 minutes): none of them decides (b) against (c). A11 is what
  (c) would leave the user with - Outlook's own delay before a window-less Outlook ends - and is listed
  among the open points below.
- The editor plan's phase D (WINWORD.EXE as a separate process): it needs a ruling under the "no
  external applications" rule first.

## How it runs

The three classes sit in one collection, `LiveComposeSurface`, whose fixture holds back every test in a
run that did not select the measurement alone - it bounces Outlook dozens of times and takes over an
hour - so a whole-tier run reports each of them `NOT RUN`. A run selects it with:

```
pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Unindexed <commit> -FilterSuffix '&FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveQ149&FullyQualifiedName!~Smoke' -SkipHubReset -RunTimeoutMinutes 180
```

and each class's `Smoke_*` test runs every cell once (`&FullyQualifiedName~Smoke` instead).

**Every attempt is one JSON line** (`Q149-ATTEMPT {...}`), in the test output and the console capture,
with the fields the plans name: the cell and iteration, Outlook's version, process id, whether it was a
COM start (`-Embedding`), the time since Outlook started, the Explorer and Inspector counts,
`IsWordMail` and `EditorType`, the editor's outcome with its exception type, HRESULT and message, the
milliseconds to it and the number of reads, the `OpusApp` count, every top-level Outlook window Windows
showed during the trial (`EVENT_OBJECT_SHOW`) and whether it was on screen, foreground changes, the
presence samples (every 50 ms to 2 s, then 5, 10, 30 s), the exit time and exit code, the Quit event,
the lifetime class (survived, parked, orderly exit, silent exit, crash), and for every commit a separate
client's re-read: the new text present, the old text gone, the HTML's length, `WordSection1`,
`MsoNormal`, the stylesheet's hash, the marker offsets, the pictures, `file:///` links and VML shapes.
No subject or body ever enters a line: only the run's own tokens, counts, offsets and hashes.

**Mailbox safety (AGENTS.md rules 1-8).** Every item is a draft in the test hub, tagged
`[OutlookAI-McpTest]` and the run's marker from the moment it is first saved, and recorded in the run's
EntryID allowlist; a change checks the tag, the marker and the allowlist first, a delete goes through
`LiveOutlookTestMailer.DeleteItemByEntryId`, and the fixture ends with the tested artifact sweep and the
count tripwire. Nothing is sent to anyone but the hub. OUTLOOK.EXE is never killed: it is ended by
WM_CLOSE to its windows, by closing its Explorers once none is visible, or by `Application.Quit()` from a
client holding nothing else, each only after the S7 counts, and a stop that does not end it fails the
cell. No inspector is ever opened on an object `Items.Add`, `Reply` or `Forward` returned.

## Results

Not run yet.

## Corrections to the D49 record

The D49 probe matrix of 2026-07-31 lives in the maintainer's untracked plan file, not in this
repository; the code comments that cite it (`ComposeSurface.cs`, `OutlookComSession.cs`) are what is
tracked. Read against the probe's own output (`probe-2026-07-31/`), it overstates in four ways:

- **It records no sample sizes.** There were 22 valid runs in all, on one machine, one build (LTSC 2021
  16.0.14334): rows 1, 2, 8 and 10 2 runs each; rows 3, 4, 5, 6, 9 and 11 1 each; row 7 3; row 12 1,
  plus 2 lost to probe bugs; the saved-draft activate 1; and 3 outside the matrix (a pinned off-screen
  activate 2, the pin-then-launch check 1).
- **"Every row ... against a FRESH COM-autostarted Outlook" is false for three runs**: row 3, one row-8
  run and the saved-draft run connected in about 90 ms, so they attached to the previous run's Outlook.
  Row 3's only run used an Outlook in which a WordEditor read had already failed, and the saved-draft
  run attached to an instance still holding the previous run's pin.
- **Rows 3, 7, 8 and 9 overstate survival.** Row 7 ("varies"): 3 of 3 exited, one alive at 300 ms and
  dead by the next call (0x800706BA). Row 8 ("varies"): its one cold start exited; the only survival was
  an instance a previous probe had started. Row 9 ("yes"): one run, alive only through the verify read;
  its delete then failed - inconclusive. Row 3 ("yes"): attached to an instance holding the previous
  trial's pin.
- **"53-79 ms to a live editor" leaves out the `Activate()` itself**, which took 314-611 ms.

Every survival was a single check 300 ms after the close, so "survives" there means alive for a few
seconds, not for good.
