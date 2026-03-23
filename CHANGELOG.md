# Changelog

## Unreleased

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
