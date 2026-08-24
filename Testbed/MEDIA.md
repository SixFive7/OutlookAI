# Installation media — the precondition nobody wrote down

**Staging media is step zero.** The rest of `Testbed/README.md` assumes you can install Windows
and Office; neither is in this repository and neither can be, so this file records what is
needed, what this machine already has, and what has to be obtained.

This exists because the gap was discovered the expensive way: a rebuild was authorised, and the
survey done immediately before it found **no Windows installation media anywhere on the host**.
Had the old VM been destroyed first, as the instruction literally said, the testbed would have
been unrebuildable. Hence the rule at the bottom of this file.

## What is needed

| | Needed | On this machine (checked 2026-08-24) |
| --- | --- | --- |
| Windows | Windows 11 Enterprise LTSC **Evaluation**, build 26100 | **ABSENT.** Searched C:, D:, E: — the only ISO images are a WinPE build and a game disc. |
| Office | Office Deployment Tool + a configuration | **PRESENT**, in an archive under the maintainer's Downloads. |

### Windows

The existing guest is Windows 11 Enterprise LTSC Evaluation, build 26100 — a `TIMEBASED_EVAL`
channel install with a 90-day clock. Evaluation media is a free download from Microsoft's
evaluation centre, but that normally requires filling in a registration form, so it does not
fetch unattended. **Someone has to obtain the ISO and stage it.** Record where it lands in the
local settings, not here.

Why evaluation media rather than a licensed edition: the testbed is disposable by decision — see
the licence section below — and an evaluation image makes that stance explicit instead of
quietly consuming a licence.

### Office — the method, which was the actual unknown

The Office Deployment Tool (`setup.exe`) plus an XML configuration. **The configuration file
contains product keys and is therefore not reproduced here**; it lives with the maintainer's
media, not in this repository. What a rebuilder needs from it, and what the existing guest was
built with:

| Setting | Value |
| --- | --- |
| Product | `ProPlus2024Volume` |
| Channel | `PerpetualVL2024` |
| Edition | 64-bit (`OfficeClientEdition="64"`) |
| Languages | `en-us`, `nl-nl`, `MatchOS` |
| Excluded apps | `Lync`, `OneDrive`, **`OutlookForWindows`** |
| Activation | `AUTOACTIVATE=1` |

**`ExcludeApp OutlookForWindows` is load-bearing, not cosmetic.** It suppresses the *new* Outlook.
Everything this project does goes through classic Outlook's COM object model, which the new
client does not provide.

**A testbed configuration should be narrower than the maintainer's.** Theirs installs Visio,
Project and proofing tools because it is a workstation configuration. A testbed wants
`ProPlus2024Volume` alone — Outlook is required, Word is worth keeping because HTML signatures
are rendered through it, and the rest is install time and disk for nothing.

## The licence clocks, and a correction worth reading

Both clocks were measured on the guest on 2026-08-24:

| | Channel | Remaining |
| --- | --- | --- |
| Windows | `TIMEBASED_EVAL` | ~82 days |
| Office | KMS client, **out-of-box grace** | ~16 days |

**The decision was to treat the VM as disposable and rebuild when a clock expires. That does not
work for Office as configured, and the arithmetic says why.** The guest was installed on
2026-08-09; 15.7 days of Office grace remained on 2026-08-24. That is a **30-day** out-of-box
grace begun at install — not a 90-day one. So a rebuild resets Office to 30 days, while Windows
resets to 90. **Office, not Windows, is the binding constraint, and it binds roughly monthly.**

**DECIDED 2026-08-24: accept the monthly rebuild.** The testbed is disposable by design and a
30-day cadence is the price of that stance. Not chosen, and worth knowing why they were on the
table: making a KMS host reachable so `AUTOACTIVATE=1` succeeds would remove the clock entirely
but depends on guest networking nobody has verified; licensing the guest another way spends a
licence on a machine meant to be thrown away.

Since Windows resets to 90 days and Office to 30, **the rebuild cadence is Office's**. Windows
never becomes the reason to rebuild.

**A cadence that depends on remembering gets skipped exactly once, and then the tier stops with
no visible cause** - a KMS client past grace drops Office into reduced functionality, on the
machine whose only purpose is driving Outlook. So the deadline is checked where it bites: the
**live tier's own preflight** refuses to run when the guest's Office licence is nearly out of
grace, instead of letting the run produce failures that look like anything except a licence.
That is the right home because it fires exactly when it matters and nobody has to remember
anything; a release-time check would not, since releases can be further apart than 30 days.
Tracked in `TODO.md`.

## The rule

**Never destroy a working testbed before the replacement runs.**

Build the new machine alongside the old one, prove it, and only then remove the old. The overlap
costs disk; the alternative cost, on a machine with no staged media, is the testbed itself. This
is not hypothetical caution — it is what the survey on 2026-08-24 prevented.
