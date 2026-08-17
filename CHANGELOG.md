# Changelog

## Unreleased

- Make every AI instruction your own: OutlookAI Settings now has Prompts and Buttons tabs. Rename the sidebar's quick buttons, rewrite the instruction behind any of them, reorder them, delete the ones you never use and add your own; edit the four prompts that wrap every request; and reset a single button, a single prompt, or the whole button set back to how it shipped. A button is its name, so renaming a shipped button gives you a custom one. Only what you actually change is stored, which means anything you leave alone keeps following the built-in default and still improves with updates, while anything you edit stays exactly as you wrote it. Changes apply to your next action without restarting Outlook, and every open compose window picks them up at once.
- Warn before you weaken a prompt: if an edited preamble no longer tells the AI to ignore instructions hidden inside an email you received, or no longer demands plain text with no markdown or HTML, the editor says so next to the text. It is advice and never blocks a save - the first protects you from a malicious email steering the assistant, the second is what keeps code fences and stray tags out of your mail.
- Put every setting in one window: OutlookAI Settings is now resizable with tabs - Outlook, Claude Code, Prompts, Buttons, Updates - instead of one tall fixed dialog, so it stops running out of room and nothing is cut off at any display scale.
- Choose which Claude model writes your mail: OutlookAI Settings has a Model group on the Claude Code tab. By default OutlookAI now chooses nothing and lets Claude Code decide, following your own model setting and picking up new models without an update - where before it was locked to one model that would have stopped working when that model retired. You can pin a family instead (opus, sonnet, haiku, fable), or type a specific model id. If Claude Code rejects your choice it quietly answers on its own default; OutlookAI now tells you when that happens instead of letting a substituted model pass unnoticed.
- Fix Outlook tuning silently doing nothing on Outlook 2013 and on future Outlook versions: the settings were written to a location only Outlook 2016 and later read, so on other versions the dialog showed "(not set)" forever and the restart notice never cleared.
- Fix the writing sidebar staying 280 pixels wide at 125% and 150% display scaling while its contents grew.
- Improve update downloads: a stalled download now gives up and says so, instead of leaving the version line on "checking..." for the rest of the session. Setup also no longer hangs indefinitely when Windows or a download server stops responding.
- Make searching more reliable when Outlook is slow to start: the first search after Outlook launches is given the time it needs instead of failing, and an exhaustive search that runs long now returns the results it found with a note, rather than timing out and taking the connection down with it.
- Make AI assistants write mail your way too: when an assistant drafts a reply or a new mail through the mail tools, it is now handed your own writing prompt - the same text the sidebar uses, exactly as you edited it - and asked to compose the body again to follow it. This happens once per session and again whenever you change your rules, so an edit takes effect immediately instead of waiting for the next session. Your rules can be as long as you like and are never copied into the tool definitions, so nothing gets silently cut.
- Fix half the mail-search instructions never reaching your AI assistant: the guidance the search tool sends was nearly twice the size Claude Code accepts, so it was silently cut in the middle - losing, among other things, the rule that tells the assistant to warn you when search results are incomplete. It now fits, with the detail moved onto the individual search options, where it arrives in full.
- Keep your settings when you uninstall: your prompts, quick buttons, Outlook tuning preferences and mail-server registration state now stay in the registry instead of being removed, so reinstalling picks up where you left off. This deliberately reverses the cleanup added in v3.1.0 - prompt text you wrote yourself is worth keeping, and nothing exports it.
- Stop drafts reading like AI wrote them: every writing action in the sidebar now carries the rule "Ensure there is no trace of AI both in wording and character use." It sits in the always-sent prompt, so it applies to every button and every instruction you type, and you can edit or remove it like any other rule.
- Check for a new version yourself instead of waiting for OutlookAI to get round to it: there is now a "check for updates" link under the version in the writing sidebar, and a "Check for updates" button in OutlookAI Settings, so you no longer have to wait up to ten minutes for the next automatic check. Both say "checking…" while one is running - including one started from the other place - and go quiet again when it finishes.
- Show the version and update state in OutlookAI Settings: a new "Version and updates" section tells you which version you are running and when OutlookAI last managed to look for a newer one, the same as the writing sidebar has always shown. When a check fails, the reason is written out in full here rather than hidden behind a link. Both indicators now take their wording from one place, so they cannot disagree.

- Fix text being cut off in OutlookAI Settings: the explanation under "Make available in all my Claude Code projects" lost its last words, and on a display scaled above 100% several other lines in the dialog were cut off too. Every wrapped line is now measured against the font actually in use and the dialog grows to fit it, so nothing is clipped at any display scale, and if it ever grew taller than your screen it scrolls instead of hiding the bottom. The dialog is also a little shorter, because space that was permanently reserved for two notices that are almost never shown is now only taken when they appear.
- Fix mail tools hanging forever with no answer when Outlook stops responding: if Outlook got into a state where it accepted requests but never replied, every mail tool - search, read, drafts, and even the health check meant to diagnose it - would wait silently until your AI assistant gave up half an hour later, and the server stayed stuck that way until it was restarted. Outlook is now driven from a separate helper process that the server can restart, so a stuck Outlook produces a clear, quick error naming what happened instead of silence, and the very next request starts from a clean slate. Searches still return your indexed mail while Outlook is unavailable.
- Report failures as real errors: mail tools that fail now mark the response as an error rather than returning a normal-looking result that merely contained an error message inside it. Assistants that did not know OutlookAI's particular convention could previously mistake a failure for a successful answer.
- Notice a stuck or starting Outlook instantly instead of waiting to find out: the server now asks Windows directly whether Outlook is responding - which costs nothing and cannot itself get stuck - before trying to use it. A stuck Outlook is reported in a fraction of a second rather than after a long wait, and a search returns your indexed mail immediately instead of stalling first.
- Stop making you wait while Outlook starts up: if Outlook is closed or still starting, mail tools now answer straight away saying so and roughly how many seconds to wait, instead of blocking for up to a minute and a half. Outlook is started in the background meanwhile, and searches keep working from the index.
- Say clearly when search results are incomplete: when the live check against Outlook cannot run, results are marked as incomplete and the assistant is told in plain words to pass that on, so an answer missing the last few minutes of mail can no longer look like a complete one.
- Stop restarting Outlook repeatedly in the background: the server could previously start Outlook again moments after a previous copy began shutting down, which appears to be what left Outlook stuck and unresponsive in the first place. It now waits before starting Outlook again.
- Stop every mail request paying the full wait once Outlook is known to be stuck: previously each request discovered the problem on its own, so a stuck Outlook made every search and every account listing take up to two minutes, over and over. After two failures in a row the server answers straight away instead, and quietly re-checks Outlook every half minute so it recovers on its own the moment Outlook responds again - restarting Outlook still fixes it instantly. In this state a search returns your indexed mail in a fraction of a second, and tells you the live check was skipped.
- Make the health check answer quickly even when Outlook is not responding: checking Outlook's health could itself take over two minutes on a machine where Outlook had stopped answering - the one moment the check is worth running. It now reports in about five seconds, says plainly that Outlook did not answer, and still gives you everything that does not depend on Outlook.
- Stop leaving stray OutlookAI processes behind: server and helper processes now shut down with the program that started them instead of accumulating in the background - 18 had built up on one machine, one of them stuck holding Outlook open.

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
