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
| W7 | W6 with B3's commit on every later draft: `ModifiedFormPages.Add("Message")` before the read, `Close(olSave)` after the edit | 3 x 10 | added on 2026-11-17, when W6's reads worked and its commits did not |

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
| B4 | Word's `Document.Save()`; `Close(olDiscard)` | 1 (see the results) |

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

### Stage 4 - after the decision (2026-11-18)

The maintainer decided on these results: keep Word, with one promotion per Outlook process, no per-draft
activation while Outlook shows a window, and the B3 commit (Q208 B); no pin at connect, the pin made when a
draft first needs it (Q175 (b)); a reply's and forward's text above the quoted header block (Q209 (a)); keep
this suite (Q210 (a)); measure what Outlook's own compose window makes of typed text (Q211 (b)); and log every
cold start of that run (Q180 (b)). These cells measure the product as it then drafts, and the two questions.

| Cell | What it does | n | Class |
| --- | --- | --- | --- |
| **P1** | The product's drafts in a window-less Outlook, through update_draft: each draft's route (direct or the promotion), every window Windows showed, every change of foreground; then the session ends - the last session out - and Outlook's exit is watched; the commits re-read | 3 starts x 10 | `LiveQ149WindowlessComposeTests` |
| P2 | Two product sessions in one window-less Outlook: the first drafts and so makes the pin, the second drafts on it, the first leaves - Outlook must stay up - and the last out closes a pin it did not make | 1 start, 5 drafts | the same |
| P3 | A COM-started Outlook in which an Explorer was then displayed (parked off-screen): every product draft must read directly, activate nothing and make no pin | 5 | the same |
| P4 | P1 while a person types: a window of the test's own holds the foreground and types a character no draft holds, every 40 ms, from just before each draft until it returns; where every key went - that window, Outlook, or the stored draft | 2 starts x 5 | the same |
| P5 | P3 while a person types, as in P4 | 5 | the same |
| **Q211** | A draft saved with no inspector (new mail, reply, forward), re-opened and displayed - Outlook's own compose window - typed into at Outlook's insertion point through Word (`Selection.TypeText`, `TypeParagraph`), its layout read and the window closed discarding; beside it the product's drafts of the same text, with no signature and with the test signature | 2 rounds x 3 kinds x 3 | `LiveQ149TypedTextTests` |
| Q180 | Every cold start of these cells: the time until Outlook first answered, `OutlookLiveness`'s verdict every 100 ms from the start request until 2 s after, the milestones, and any step that did not return within its budget; and every OUTLOOK.EXE that started during the run, from its creation until it answered | every start | both |

In P1 to P5 a small window of the test's own takes the foreground before each draft, standing in for the
program a person is typing in, and a hook on every foreground change of the desktop records where the
foreground went during the draft and for how long it was not on that window.

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

The four classes sit in one collection, `LiveComposeSurface`, whose fixture holds back every test in a
run that did not select the measurement alone - it bounces Outlook dozens of times and takes over an
hour - so a whole-tier run reports each of them `NOT RUN`. A run selects it with:

```
pwsh -File Testbed/host/Invoke-LiveTierOnGuest.ps1 -Kind Unindexed <commit> -FilterSuffix '&FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveQ149&FullyQualifiedName!~Smoke' -SkipHubReset -RunTimeoutMinutes 180
```

and each class's `Smoke_*` test runs every cell once (`&FullyQualifiedName~Smoke` instead). Every
window-less cell is a test of its own (`L1_...`, `L1c_...`, `L2_...`, `W2_...` to `W7_...`), so a run can
take a group of them by name - `&(FullyQualifiedName~.L1_|FullyQualifiedName~.L1c_)` - and stay short.
The measurement was run that way: one group of cells per run, with `-Kind Unindexed` or `-Kind Indexed`
so that the runs spread over both guests of a kind. Stage 4 runs as two groups,
`-FilterSuffix '&(FullyQualifiedName~.P1_|FullyQualifiedName~.P2_|FullyQualifiedName~.P3_|FullyQualifiedName~.P4_|FullyQualifiedName~.P5_)'`
and `-FilterSuffix '&FullyQualifiedName~.Q211_'`: a run whose test process does not end is fetched with
nothing, so a Q211 that blocked would take the P cells' lines with it. On the frozen checkpoints of 2026-10-03 a run's guest
account stopped answering about 40 minutes after the restore, and the first smoke run lost its results
that way; every run in the results is on the checkpoints recorded since, `CP-16B` and `CP-22C`.

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

Measured on 2026-11-17 on Office LTSC 2024 (16.0.17932), on the four Outlook guests: `OutlookAI-Unindexed`
and its clone `OutlookAI-Unindexed2` from `CP-16B-FROZEN-CLOCK`, `OutlookAI-Indexed` and `OutlookAI-Indexed2`
from `CP-22C-FROZEN-CLOCK`. One run per group of cells and kind of guest; every run passed. Every attempt
line of every run is in `results/`, one file per run, named after the run's folder.

| Cells | Unindexed guest | Indexed guest |
| --- | --- | --- |
| L1, L1c, L2 | `20261117-185607-unindexed-916755eda809` | `20261117-185655-indexed-916755eda809` |
| W2, W3, W4 | `20261117-190436-unindexed-3c5a63e47a1e` | `20261117-190553-indexed-3c5a63e47a1e` |
| W5, W6 | `20261117-190725-unindexed-3c5a63e47a1e` | `20261117-190841-indexed-3c5a63e47a1e` |
| W7 | `20261117-193747-unindexed-a7d41bc39928` | `20261117-193829-indexed-a7d41bc39928` |
| B0-B3, C1-C5 | `20261117-190159-unindexed-3c5a63e47a1e` | `20261117-190318-indexed-3c5a63e47a1e` |
| B4, once | - | `20261117-181203-indexed-ee5fc276b528` (its one B4 line) |

The L runs are of `916755e`, before master's signature work was merged into `3c5a63e`. That merge added
a signature decision ahead of `ReviseHeldDocument`'s inspector work and left the inspector work itself -
`GetInspector`, the promotion, the edit, `CurrentItem.Save`, `Close(olDiscard)` - as it was, so they were
not run again. The two kinds of guest agree in every cell. Where two numbers stand together below, they
are the unindexed guest's and the indexed guest's.

### Stage 1: an Outlook with no window

| Cell | n per guest | Outlook after the close | The editor | On screen | The edit stored |
| --- | --- | --- | --- | --- | --- |
| L1 | 20 | ended in 20 and 20 | through the product's promotion | the compose window, 20 and 20 | 20 and 20 |
| L1c | 5 | running in 5 and 5 | through the product's promotion | the compose window, 5 and 5 | 5 and 5 |
| L2 | 10 | ended in 10 and 10 | refused in 10 and 10 | nothing | - |
| W2 | 10 | ended in 10 and 10 | refused in 10 and 10, 114-116 reads over 30 s | nothing | - |
| W3 | 3 | ended in 3 and 3 | refused in 3 and 3, reads from 180 s after the start | nothing | - |
| W4 | 10 | ended in 10 and 10 | refused in 10 and 10 | the inspector's window, off-screen, 10 and 10 | - |
| W5 | 10 | running in 10 and 10 (the pin) | refused in 10 and 10 | the pin's window in 1 and 1 trial | - |
| W6 | 3 x 11 | running (the pin) | the promotion 3 and 3; then a direct read 30 and 30 of 30 | the compose window at each promotion | promoted 3 and 3; read directly 0 and 0 of 30 |
| W7 | 3 x 11 | running (the pin) | the promotion 3 and 3; then a direct read 30 and 30 of 30 | the compose window at each promotion | all 33 and 33, no side effect |

- **L1, the decisive cell.** Without the pin, the window-less Outlook ended at the close of the compose
  window in 20 of 20 trials on each guest. Its Quit event fired and the process exited with code 0,
  1.9 to 4.0 seconds after the compose began. Every edit reached the stored draft, but update_draft
  reported a failure in all 40, because Outlook was gone when it read the draft back:
  `MissingMemberException` 38 times, `COMException 0x800706BE` twice. With the pin (L1c) Outlook kept
  running in 10 of 10, and every compose returned, placed by Word.
- **The promotion is on screen.** Every product compose in a window-less Outlook, the 50 of L1 and
  L1c, showed the compose window at (25,25), 768 x 519, for 158-335 ms (median 207) on the unindexed
  guest and 130-646 ms (median 147) on the indexed one, and gave Outlook the foreground. The promotion
  parks the window off-screen before `Activate()`, and `Activate()` puts it back where Outlook keeps
  it; the sweep after it parks and hides it again. After the run's first compose on each guest the
  foreground stayed on an Outlook window although none was visible. In a windowed Outlook the window
  stayed off-screen (B0, 40 of 40). That first compose, L1c's first trial, took 13.0 and 46.4 s; every
  other one 1.9 to 4.8 s.
- **L2.** An inspector never activated refused its editor in 10 of 10 on each guest:
  `COMException 0x80004005`, "The operation failed.", within 39 ms, although `IsWordMail()` was true,
  `EditorType` was 4 (Word) and Word's `OpusApp` window existed. Fetching the inspector changed
  nothing in the item: 241 bytes of HTML and no stylesheet, before and after. Its close ended Outlook
  in 10 of 10, with its Quit event and exit code 0: an inspector that was never activated ends a
  window-less Outlook as surely as the shipped compose.
- **W2, W3.** Waiting did not help. Read every 250 ms for 30 seconds, 114 to 116 reads, the editor
  was refused at every read of every trial, 10 of 10 on each guest, and so it was when the reads began
  180 seconds after the cold start (W3, 3 of 3). Outlook ended at every close.
- **W4.** Nor did Windows showing the inspector's own window, found through `IOleWindow`, off-screen
  and without activating it, with no change of foreground: refused in 10 of 10 on each guest, over 5
  seconds of reads. Outlook ended at every close. In one trial on the indexed guest its Quit event and
  exit came 8.2 and 9.7 s into a treatment that took 14.9 s, before the trial's close had returned.
- **W5.** Showing the pin's window through Windows (`SW_SHOWNOACTIVATE`, no change of foreground) did
  not unlock the editor: refused in 10 of 10 on each guest, after 19 or 20 reads over 5 seconds. The
  window was shown off-screen except in the first trial on each guest, where it was on screen at
  Outlook's own place for 244 and 133 ms.
- **W6.** After the product's promotion of a process's first draft, every later inspector that was
  never activated handed out its editor at the first read - 27 of 27 on each guest, with no window
  shown - and still did after a displayed Explorer was opened and closed (3 of 3). But none of those
  edits reached the stored draft through `CurrentItem.Save()` and `Close(olDiscard)`: 0 of 30 on each
  guest, which is B1's result in a window-less Outlook. The three promoted drafts' edits did. (The W6
  runs' `oldAbsent` field means nothing: the old text was keyed by the draft's number and checked by its
  iteration, fixed in `a7d41bc`. `tokenPresent`, the new text, is right.)
- **W7.** W6 with B3's commit: every later draft's edit reached the stored draft, 30 of 30 on each
  guest, with the old text gone, the message class `IPM.Note`, no form-definition property, and no
  window shown; with the promoted drafts, 33 of 33. One promotion per Outlook process is enough, the
  commit included.

**D1: the pin is required for as long as drafts go through an inspector.** L1 ended Outlook in 20 of 20
on each guest.

**D2: promote once per process, not per compose.** Waiting did not bring the editor (W2 and W3, 0 of 13
on each guest), and nor did Windows showing a window (W4 and W5, 0 of 20 on each guest), so the rule's
first two branches fail. After one promotion, every later inspector in the process handed out its
editor (W6 and W7, 60 of 60 on each guest), and W7 committed each such edit with B3's commit. The
compose window then shows once per Outlook process instead of once per draft, as long as the process
outlives the draft: a session-long pin, as today, gives that, and a pin made only while drafting would
not, since closing the last pin ends a window-less Outlook (Q118).

### Stage 2: committing an edit on a re-opened draft without Activate

| Cell | n per guest | Committed | Editor read without Activate | Message class or form properties changed |
| --- | --- | --- | --- | --- |
| B0 | 20 | 20 and 20 | - (it activates) | 0 and 0 |
| B1 | 20 | 0 and 0 | 20 and 20 | 0 and 0 |
| B2 | 20 | 0 and 0 | 20 and 20 | 0 and 0 |
| B3 | 20 | 20 and 20 | 20 and 20 | 0 and 0 |
| B4 | 1, indexed only | 0 | 1 | - |

- B1 and B2 read the editor and made the edit with no error, and the stored draft still held the old
  text.
- B3 committed 40 of 40. The message class stayed `IPM.Note` and no form-definition property
  appeared. One B3 trial on the indexed guest counted two windows on screen; the commit cells count
  windows and do not log them, so which they were is not known.
- B4 ran once. Word's `Document.Save()` put one window on screen, did not return within the step's
  180 seconds, and committed nothing. Word documents `Save()` on a document never saved as a file as
  prompting with its Save As dialog, which waits for a person; it was not repeated.

**D3: B3 commits every draft with no side effect, so a windowed Outlook can drop the per-compose
`Activate()`.**

### Stage 3: the HTML-only draft against today's

Two ways of styling the HTML-only draft were measured, each against the same draft of today's:
**Message** gives every bare paragraph the message's `MsoNormal` class, and **AsImported** leaves the
markup as it came, as Word's HTML import receives it. Both guests gave the same result in every aspect,
down to the number of differences.

| Aspect | New mail (5 pairs) | Reply (3) | Forward (3) | update_draft (2 sets) | As received (2 pairs) |
| --- | --- | --- | --- | --- | --- |
| The stylesheet | same in 5 (borrowed) | same in 0 | same in 0 | - | same in 2 |
| The body inside `WordSection1` | same in 5 | same in 3 | same in 3 | - | - |
| The signature with its pictures | same in 5 | same in 3 | same in 3 | kept in all | the pictures same in 2 |
| Looks the same, Message | 0 (16 or 19 differences) | 0 (36) | 0 (36) | 0 (6 or 16) | 0 (16 or 19) |
| Looks the same, AsImported | 0 (4 or 7) | 0 (21) | 0 (21) | - | 0 (4 or 7) |

- **A new mail, AsImported, the closest.** Today's draft has one paragraph more, an empty one after
  the signature, and its blank line between the body and the signature is Word's "Normal (Web)" style,
  5 pt before and after, where the HTML-only draft's is "Normal" with none. With a plain-text body
  the text's own paragraph differs the same way. That is all of it: 4 differences with an HTML body,
  7 with a text body.
- **A new mail, Message.** The same, and every paragraph of an HTML body differs that way too: today's
  draft lays out the body's paragraphs as "Normal (Web)", 5 pt before and after.
- **A reply and a forward.** Today's drafts put the body and the signature inside the quoted original's
  header block, under the line that separates it: Word lays out their first eleven paragraphs with that
  block's top border, and the first of them, the body's heading, in Times New Roman where the HTML-only
  draft's is Aptos. The HTML-only drafts leave the body where `Reply()` and `Forward()` leave room for
  it, above the header block. Both differ from today's in every pair: 36 aspects (Message) and 21
  (AsImported).
- **What `Reply()` and `Forward()` hand over without an inspector (C4).** The quoted original and its
  header block whole (6 of 6), no `_MailAutoSig` (6 of 6), and the original's stylesheet, not the one
  Word writes into today's reply (0 of 6 the same).
- **update_draft (C3).** Every update replaced the text and kept the signature and its pictures. Against
  Word's update of a Word draft, today's, the HTML-only update of a Word draft differs in 6 aspects
  (the new paragraphs' style), the HTML-only update of an HTML-only draft in 16, and Word's own update
  of an HTML-only draft in 16 as well.
- **As received (C5)**, sent through the loopback sink with the product's two-step send: the same
  differences as opened in Outlook, the stylesheet and the pictures the same.
- **What a new item made with no inspector holds (C1).** 241 bytes of HTML from `Items.Add`, no
  stylesheet and no `WordSection1` (0 of 5): (c) would have to bring the whole `<head>` itself. The
  HTML-only drafts here borrowed today's, their best case.

**D4: no HTML-only draft looks the same as today's, opened in Outlook or as received, so by the rule
the answer is (b).**

### Found on the way

- **Today's replies and forwards put the agent's text under the quoted original's separator line**
  (above). `ApplySignatureToDocument` places a signature at the start of `_MailOriginal` when the
  draft has no `_MailAutoSig`, and a reply made with no inspector has none; that anchor sits inside the
  header block's bordered `div`, and the body then goes above the signature, inside the same block. A
  reply without a signature was not measured.
- **Today's body paragraphs are "Normal (Web)", 5 pt before and after,** where a paragraph of the
  message's own `MsoNormal` class is "Normal" with none. What Outlook's own compose window makes of a
  typed paragraph was not measured.
- **The promotion is visible in a window-less Outlook** (above). The code and the D49 record call it
  invisible; on Office LTSC 2024 the window is on screen for 130-646 ms and takes the foreground.
- **The first new_draft of the parity class fell back to the HTML path in both stage 2-3 runs**:
  `NoWordEditor` with `COMException 0x80004005`, a draft without its signature; the same call again
  went through Word. It came right after the commit class had closed 60 inspectors without activating
  them in the same Outlook; the cause was not isolated.

### Not measured

- A third HTML-only style copying today's markup - a bare `<p>` for the body and the blank line, and
  an empty `MsoNormal` paragraph after the signature. The differences above are exactly those, so it
  may match a new mail; a reply could match today's only by copying its misplacement.
- What Outlook's own compose window makes of the same text, as a yardstick beside today's drafts.
- Office LTSC 2021, and Microsoft 365 Apps: the guests run Office LTSC 2024 only.

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
