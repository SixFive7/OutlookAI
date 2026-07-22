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
