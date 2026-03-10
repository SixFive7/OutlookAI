# Changelog

## Unreleased

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
