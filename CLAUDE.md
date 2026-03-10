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

- The project builds via GitHub Actions on push to master (`.github/workflows/release.yml`).
- Version is `BASE_VERSION` (in the workflow) + CI run number. To bump major/minor, edit `BASE_VERSION`.
- The CI extracts the Unreleased changelog section as release notes, then commits a stamped changelog back.
- The workflow ignores commits from `github-actions[bot]` to avoid infinite loops.
