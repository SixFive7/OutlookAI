# Changelog

## Unreleased

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
