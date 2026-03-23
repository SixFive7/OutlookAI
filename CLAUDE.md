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

- **Build** runs automatically on every push to master (`.github/workflows/build.yml`). It only compiles — no releases, no tags, no changelog changes.
- **Release** is triggered on demand (`.github/workflows/release.yml`) via `gh workflow run release`. It extracts the Unreleased changelog section, builds, creates an installer, publishes a GitHub Release, and stamps the changelog.
- The release workflow **fails if the Unreleased section is empty** — you must have release notes before creating a release.
- Version is `BASE_VERSION` (in both workflow files) + commit count. The release workflow handles bumping automatically.
- The release workflow requires a `version_bump` input in `major.minor.patch` format (e.g. `0.0.0` for no change, `1.0.0` for major bump, `0.1.0` for minor, `0.0.1` for patch). This input is **required** — the workflow will not run without it.
- **After committing, ALWAYS ask the user if they want to create a release.** If yes:
  1. **ALWAYS ask the version bump question.** Read the current `BASE_VERSION` from `.github/workflows/release.yml` and show the user: current version, and what major/minor/patch/no-change would produce. This question is mandatory because the release workflow requires the `version_bump` input.
  2. Run: `gh workflow run release -f version_bump=X.X.X` with the user's chosen bump value.
  3. Monitor with `gh run watch`.
- After a release, pull the stamped changelog commit before continuing work: `git pull --rebase`.
