# Project Instructions

## Changelog

When committing changes, **always update `CHANGELOG.md`** under the `## Unreleased` section.

Rules:
- Write entries as **user-facing summaries**, not developer jargon. Describe what changed from the user's or project's perspective.
- Keep entries concise — one line per change, starting with a verb (Add, Fix, Remove, Update, Improve).
- Group related commits into a single entry when they are part of the same logical change.
- Do not include CI fixes, typo fixes, or internal refactoring unless they affect user-visible behavior.
- Never modify released sections (any `## v...` heading). Only add to `## Unreleased`.
- If the Unreleased section already has entries from earlier in the session, add to it rather than replacing it.

## Build and Release

- **Build** runs automatically on every pull request (`.github/workflows/build.yml`), and can be triggered on demand. It only compiles — no releases, no tags, no changelog changes. On pull requests it also runs a dependency review.
- **Release** is triggered on demand (`.github/workflows/release.yml`) via `gh workflow run release`. It extracts the Unreleased changelog section, builds, creates an installer, publishes a GitHub Release, and stamps the changelog.
- The release workflow **fails if the Unreleased section is empty** — you must have release notes before creating a release.
- Version is derived from the latest GitHub release tag (base version) + commit count. No hardcoded version in the repo.
- The release workflow requires a `version_bump` input in `major.minor.patch` format (e.g. `1.0.0` for major bump, `0.1.0` for minor, `0.0.1` for patch). This input is **required** — the workflow will not run without it. `0.0.0` is rejected — every release must bump at least one version component.
- **After committing, ALWAYS ask the user if they want to create a release.** If yes:
  1. **ALWAYS ask the version bump question.** Get the current version from the latest release tag via `gh release view --json tagName -q .tagName` and present options in A/B/C format showing current → new version. Example with latest tag v2.1.0.103:
     - A) Patch — 2.1.0 → 2.1.1
     - B) Minor — 2.1.0 → 2.2.0
     - C) Major — 2.1.0 → 3.0.0
  2. Run: `gh workflow run release -f version_bump=X.X.X` with the user's chosen bump value.
  3. Monitor with `gh run watch`.
- After a release, pull the stamped changelog commit before continuing work: `git pull --rebase`.

## MCP Server (`McpServer/`)

- `McpServer/` holds the MCP server projects (`OutlookAI.Core`, `OutlookAI.McpServer`, `OutlookAI.McpServer.Tests`). Build them with `dotnet build` **by explicit csproj path** — never via `OutlookAI.slnx`, which only contains the VSTO add-in (MSBuild-only).
- Their CI is `.github/workflows/mcpserver.yml` (windows runner, dotnet only; runs `dotnet test --filter "Category!=Live"`). Tests marked `Category=Live` need the real Windows Search index plus Outlook and only run on a configured dev machine.
- Developer documentation: `McpServer/README.md`.

## Mailbox Safety (MANDATORY — live tests touch REAL mailboxes)

`Category=Live` tests run against the developer's **real production Outlook profile**: real mail accounts plus delegate/shared mailboxes **to which the profile has full write access**. Treat every live run as an operation on production data. A past incident mass-deleted real mail (fully recovered) because an agent improvised a cleanup script — these rules exist so that never repeats. They are non-negotiable and apply to every agent, every session, whether or not live tests are the task:

1. **Never mutate mailbox items from ad-hoc shell code.** No PowerShell, no raw COM one-liners, no throwaway scripts. Creating, deleting, moving or editing an item happens **only** through the project's tested helper code (`LiveOutlookTestMailer` and the live fixtures) or the shipped MCP tools. If cleanup needs something the helpers cannot do, extend the helpers with tests — do not improvise.
2. **Never pattern-match subjects shell-side.** PowerShell's `-like "*[tag]*"` treats `[...]` as a character-class wildcard, so it matches nearly every subject — that is exactly how real mail was destroyed. Deletion selection is **EntryID allowlist AND ordinal tag match, both required**; every test-created item carries the `[OutlookAI-McpTest]` subject tag.
3. **Writes only in the designated test mailbox** (see the gitignored live-test settings). Every other account and **all** delegate/shared mailboxes are read-only for tests — no exceptions. The `StoreWriteAllowlist` guard enforces this in code: a write aimed anywhere else throws instead of running. Logs and failure messages never print other stores' subjects or bodies.
4. **Every live run must end with zero tagged artifacts**, proven by the post-run sweep — which covers Drafts, Inbox, Sent Items, **Outbox**, Deleted Items and the **Sync Issues subtree** (Conflicts / Local Failures / Server Failures), plus test folders removed deepest-first.
5. **A live run may not lose mail anywhere.** The per-store count tripwire snapshots every store's mail folders before and after; any item-count **decrease**, or any folder added/removed, outside the test mailbox fails the suite loudly. No snapshot ⇒ the live tier refuses to run.
6. **Signatures are user data.** Tests may only create/update/delete signatures prefixed `OutlookAI-McpTest-`; the `SignatureDirectorySnapshot` guard (SHA-256 before/after) must run and the suite must leave the user's real signatures bit-identical. `manage_signature` tests restore any registry defaults they touch.
7. **Outlook lifecycle:** never `taskkill` OUTLOOK.EXE. Graceful `Application.Quit()` only when no unsent compose windows are open and the Outbox is empty — and release COM references BEFORE quitting (quitting while refs are held zombifies the process). Prefer leaving Outlook headless.
8. **Run live tests only via the suite** (`dotnet test <Tests csproj> --filter "Category=Live"`); its fixtures enforce the snapshots, allowlists, tripwire and zero-artifact sweeps. Never perform mailbox operations outside it during testing.
9. If a gitignored `v3.MD` exists at the repo root, read its §0 safety envelope before any live-test or mailbox-touching work — it is the authoritative, more detailed contract.
