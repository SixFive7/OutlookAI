# Changelog

## Unreleased

- Fix mail search missing mail whose search word appears only in the subject: search words were matched against message bodies and attachment contents only, so alert prefixes, ticket tags, invoice numbers and other subject-phrased terms simply did not come back (measured at roughly 1 in 30 mails store-wide). Searching now covers the subject and the body together in every search path - the index, the freshness sweep of just-arrived mail, and the exhaustive folder scan.
- Fix multi-word mail searches missing mail whose words are spread across the subject and the body: searching for two words only returned mail where both appeared in the same part, so a mail with one word in the subject line and the other in the message text came back empty. Each word is now looked for in the subject and the body independently, which finds those mails at no measurable cost (measured: a few percent more results on real mailboxes, no slower).
- Extend the freshness sweep of just-arrived mail to follow what you are searching: searching one folder now also sweeps that folder and its subfolders, so mail a mail-server rule files straight into another folder (for example alerts routed into Deleted Items) is found immediately instead of only after the search index catches up. Searches that are not folder-specific now sweep Deleted Items and Junk Email as well as the Inbox and Sent Items, and every result reports which folders the freshness sweep covered.
- Add a search_in option to mail search: keep the default and a word is found in the subject or the body, or narrow it to just the subject or just the body when a word is noisy in the other one (for example a term that appears in every quoted thread or every alert subject line). The search tool description is now a plain usage guide for the assistant: what search words are matched against (whole words, each of which may match in either the subject or the body), that sender and recipient are matched by the separate from/to filters instead, which follow-up actions a result can be used for, and how freshness, capped result lists and the exhaustive folder scan behave.
- Add mail moving: the assistant can now move mail (1-50 at a time) to another folder within the same account - refile into project folders or restore items from Deleted Items - creating the target folder on request. Every move is audit-logged and reversible: results carry the source folder and the item's old and new ids so any move can be undone, and moving to Deleted Items or the Outbox is refused (the assistant still cannot delete mail).
- Add one-click archiving: the assistant can archive mail exactly like Outlook's own Archive button (Backspace) - each item goes to its own account's designated Archive folder, resolved per mailbox even when it has a localized name (for example "Archiveren"); accounts without a designated Archive folder get a clear error and nothing is created. Audited and reversible like any move.
- Add a "Select the best signature" button to the AI writing sidebar: the AI looks at your draft, the quoted thread, and the recipients, picks the most fitting of your installed signatures (for example by matching the language), and applies it - your draft text and the quoted conversation stay untouched. With a single installed signature it is applied directly without an AI call, and the button is available only when at least one signature exists.
- Add signature management to the assistant: it can now create, update, and delete Outlook email signatures (writing all three formats - HTML, plain text, and RTF - and deriving whichever you did not supply) and optionally record one as an account's default for new mail and/or replies. Before any update or delete the previous signature files are automatically backed up under %LOCALAPPDATA%\OutlookAI\signature-backups and the backup location is reported; deleting a signature also cleans up default assignments that pointed at it.
- Add signature steering: the assistant can now list your installed email signatures (with a short excerpt so it can tell their language and purpose apart) and apply a specific one to any draft it prepares - for example a signature matching the recipient's language - while omitting the choice keeps the account's default; per-account defaults are reported where Outlook records them.
- Add body paging to mail reading: very long mails can now be read in consecutive windows that continue exactly where the previous one ended, without re-reading the whole body from the start each time.
- Improve folder listing: it now always returns the complete folder tree of a mailbox (no depth setting to get wrong) in a stable order, with paging available in the unlikely case a profile exceeds 1000 folders.
- Merge the assistant diagnostics into one outlook_health tool: a single call now reports Outlook state and version, connection health, store reachability, search-index freshness overall and per mailbox with advice, Windows Search service state, audit-log writability, and tuning state; the separate health, index_status, and echo tools are gone.
- Add mail-server usage instructions that load into every Claude session automatically, so the assistant knows it can search and read your Outlook mail (and that drafts open for review while sending stays gated) even before any mail tool is used.
- Add a clear warning when Outlook's search runs online (server-assisted): showing search results in Outlook now tells the agent that the on-screen list may differ from what the agent itself finds (online results are capped and ranked differently) and recommends switching local search back on, and the health report now states which search backend Outlook's own search box is actually using ("local" or "server-assisted").
- Add a caution line to the Search group in OutlookAI Settings explaining what turning local search tuning off means: slower online search, capped results, and "show me" results that may no longer match what the agent finds.
- Simplify mail search to one always-fresh tool: the fast/fresh mode choice is gone - every search now automatically includes mail that arrived seconds ago. The freshness sweep is cached for about 10 seconds so rapid repeated searches stay instant, and when the sweep cannot run (for example during an add-in update) the search still returns index results with a clear freshness warning instead of failing. The exhaustive folder scan is now a simple true/false option on the same search tool.
- Fix search results contradicting themselves right after an automatic Outlook start: the "Outlook running" indicator now reflects the state after the freshness sweep, not before it.
- Fix the assistant server keeping a closed Outlook "connected": it now watches the Outlook process and releases every held connection the moment Outlook exits (user close, crash, or logoff), so a dead Outlook is never held open by stale references and the next mail request starts Outlook fresh without errors.
- Fix the health report claiming a connection to an Outlook that already exited: connectivity is now actively verified at report time, never inferred from a stale held session.
- Add a "headless" indicator to the health report: it now tells you whether Outlook is running invisibly in the background (tray icon only) or as a normal window.
- Improve draft results while Outlook is starting up, busy, or restarting: the reported store, folder, and sending account no longer come back empty when Outlook has only just started or is momentarily unresponsive, and an Outlook exit in the middle of an operation is now handled as a recoverable condition instead of surfacing a raw error.
- Guarantee background operation: only the explicit show-me tools (open a mail, jump to a folder, show search results) and draft windows ever open Outlook UI - every other assistant operation now provably leaves Outlook window-less, verified by tests.
- Document Outlook's lifetime around the assistant in the README: headless background start, the tray icon meaning, the ~10-12 minute self-exit after the last agent disconnects, promoting to a normal Outlook by just launching it, and what closing the window does.
- Add a health check tool: one call reports Outlook state and installed version, mail store reachability, search-index freshness, Windows Search service state, audit-log writability, and the current Outlook tuning state - without ever starting Outlook.
- Improve assistant payload discipline: search and conversation results now say explicitly when the requested cap cut the list (with guidance to raise it or narrow the query), and very long recipient or attachment lists are capped with clear has-more markers while operations keep using the full lists.
- Add an OutlookAI Settings dialog: a large button in the OutlookAI group on Outlook's main Mail ribbon (carrying the same icon as the AI Assistant) opens a settings window (matching your Office light or dark theme) where each tuning group can be switched on or off, the current effective values are shown, and a clear indicator tells you when an Outlook restart is still needed. Turning a group off stops managing it and leaves your Outlook settings as they are.
- Add automatic Outlook tuning: OutlookAI now keeps your proven Outlook configuration applied on every start - fast local search settings, full mailbox caching (sync slider = All, shared folders included), and raised OST size limits so a fully cached mailbox never stalls. Only actual differences are written, values enforced by your organization's group policy are respected and flagged, and you are told when an Outlook restart is needed for a change to take effect.
- Add a deliberately high-friction send tool: the assistant can send a saved draft only via a two-step confirmation (a one-time token bound to that exact draft and its current content, expiring after two minutes and voided by any draft change), with the sending account hard-verified before transport and every step audit-logged; drafting for you to review and send yourself remains the default path.
- Add email drafting tools: the assistant can now prepare a new mail, reply, reply-all, or forward for you - the draft opens on screen with the right account identity, that account's own signature, and the assistant's text above the quoted conversation, ready for you to review and send yourself (nothing is ever sent automatically).
- Add a structured audit log: every draft the assistant creates and every attachment it saves is recorded locally under %LOCALAPPDATA%\OutlookAI.
- Add show-me tools: the assistant can now open a mail in an Outlook window for you, jump your Outlook to any folder, and run a search in Outlook's own search box so you see the result list on screen (works even if Outlook was not running).
- Add exhaustive search mode: a folder- or date-bounded scan straight through Outlook that bypasses the search index, for when the index is stale or correctness matters more than speed; searches now also advise when the index is stale enough to warrant it.
- Add the OutlookAI MCP server: a local server that lets AI agents such as Claude Code work with your mail - it searches all locally indexed Outlook mail (all accounts, delegate mailboxes, and attachment contents) in milliseconds, with a fresh mode that also catches mail that just arrived and is not yet indexed.
- Add mail reading tools: full message read with safe truncation, sender/recipient details, conversation (thread) view, and saving attachments to disk so the assistant can open them.
- Add mailbox insight tools: account and store listing (delegate and online-only stores flagged), folder trees with unread counts, and an index freshness self-report with actionable advice.
- Update the README with the MCP server and Settings documentation, and add developer documentation for the MCP server.

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
