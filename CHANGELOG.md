# Changelog

## Unreleased

- Use Inno Setup for a proper single-file installer experience

## v1.0.0.5 - 2026-03-10 (yanked)

- Add automated build and release pipeline via GitHub Actions
- Add ClickOnce installer published as GitHub Release on every push to master
- Add assembly version synchronization with release version
- Rename signing certificate to `OutlookAI.pfx` and track in repository
- Add changelog with CLAUDE.md rules for AI-maintained release notes
- Replace Anthropic API with Claude Code CLI as AI backend, using pre-warmed subprocess for zero-latency requests
- Add Custom Action feature for arbitrary user instructions on email content
- Upgrade target framework from .NET 4.7.2 to .NET 4.8
- Remove voice input and OpenAI Whisper integration (users dictate via external tools)
- Remove configuration UI and API key management (no longer needed with Claude Code CLI)
- Remove enterprise RDS deployment scripts in favor of single-user ClickOnce install
- Simplify project structure: flatten directory layout, remove dead code and unused references
