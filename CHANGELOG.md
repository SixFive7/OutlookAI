# Changelog

## Unreleased

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
