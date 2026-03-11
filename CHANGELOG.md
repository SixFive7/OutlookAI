# Changelog

## Unreleased

## v1.0.0.51 - 2026-03-11

## v1.0.0.49 - 2026-03-11

- Add AI Assistant sidebar for inline replies in the reading pane
- Auto-show the AI sidebar whenever composing an email (new, reply, forward)
- Remove the ribbon button — the sidebar now appears automatically
- Fix crash when closing a compose window while an AI request is in-flight
- Fix wrong email being processed when multiple compose windows are open

## v1.0.0.47 - 2026-03-10

## v1.0.0.3 - 2026-03-10

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
