#Requires -Version 5.1
<#
    ============================================================================================
    WRITTEN 2026-10-03, WHEN THE GITHUB CI WAS REMOVED. DRY-RUN ON THE MAINTAINER'S WORKSTATION.
    -Execute, WHICH PUBLISHES, HAS NEVER RUN.
    ============================================================================================

    The first dry run (from edc9dfd) stopped at step 8, and rightly: Get-AuthenticodeSignature
    calls every signature by the self-signed release key UnknownError, because its root is not
    trusted, so step 8 now asks WinVerifyTrust exactly as the shipped updater does. The second
    (from 4d9d353, 2026-10-03 18:25-18:32, 7 minutes) ran steps 1-10 through: the four guards under
    PowerShell 7.6.6 and 5.1; D7 (c) 8/0 and 3/0 against Visual Studio Community 2026 18.10.3's
    targets; the release build of 3.1.1.825, guard 3 UNCHANGED over 305 host lines, the add-in DLL
    and both server executables stamped 3.1.1.825; the installer signed, WinVerifyTrust 0x800B0109
    (the result the updater accepts), DigiCert timestamp, 43.2 MB of the 50 MB cap; the build VM's
    run of 4d9d353 PASS - 3607 of 3607 tests and 23 of 23 self-tests, this script's 72/0 among
    them; and the stamp commit made, two lines of CHANGELOG.md and nothing else. It also measured
    the notes at 151,497 characters, over the 125,000 GitHub accepts as a release body. Since Q123
    (2026-10-03) the body is no longer the section: step 3 generates it, BrowserAI's way, from
    each entry's headline and a link to that entry's lines.

.SYNOPSIS
    Builds, tests, signs and publishes an OutlookAI release from the maintainer's workstation -
    what .github/workflows/release.yml did on a GitHub runner until 2026-10-03.

.DESCRIPTION
    WHY THIS EXISTS. Decided by the maintainer 2026-10-03, in his words: "Remove the git CI
    pipeline. I want only the build and test suite to run on my pc. No CI pipelines on github."
    The release workflow was the one pipeline that produced something. This is its port, step for
    step, plus what a workstation needs that a fresh runner did not. The workflow itself stays in
    git history - `git show 883ec5f:.github/workflows/release.yml` - and each step below names the
    workflow step it replaces.

    WHAT ONE RUN DOES, IN ORDER. Every step before PUBLISH only reads, builds under .work\, or
    writes git objects no ref points at. A failure anywhere stops the run with nothing published.

       1. PRECONDITIONS. Not elevated - an elevated process can carry another account's HKCU, and
          the build runs next to this user's Outlook. A clean working tree, because what the guards
          read must be what is built. `git fetch origin master`. With -Execute, HEAD must be
          origin/master's tip: the release workflow released master, and the stamped changelog
          goes onto it as a fast-forward.
       2. VERSION (release.yml's "Validate version bump input" and "Apply version bump"). The base
          is the latest GitHub release's tag plus -VersionBump: a major bump resets minor and
          patch, a minor bump resets patch. The fourth part is HEAD's commit count plus one - the
          stamp commit this run makes, as every release so far (v3.1.0.325 is commit 325). The tag
          must not exist yet.
       3. RELEASE NOTES (release.yml's "Extract changelog", replaced by BrowserAI's system - Q123):
          the "## Unreleased" section of CHANGELOG.md, held to the changelog's rules (AGENTS.md,
          Changelog). No entries refuses. So does an entry that is not '- <icon> **One-sentence
          headline.** the detail', an icon outside the legend, a headline over 100 characters or of
          two sentences, a detail that opens by repeating its headline, groups that are not Keep a
          Changelog's in its order, a legend that is not the palette, a character nobody types in
          what the release page shows, and a released section that differs from what the latest
          release's tag carries. With -Execute a section with no preamble refuses too. Then the
          stamp is worked out (step 10 commits that very text) and the GITHUB RELEASE BODY is made
          from it: the preamble, each entry's icon and headline with a "read more" link to that
          entry's own lines in CHANGELOG.md as the tag will carry it, the legend, and a link to the
          whole section. A body over GitHub's 125,000-character limit drops the links; one still
          over it refuses with -Execute, which the workflow would have found only after pushing
          the stamp commit.
       4. THE SIGNING CERTIFICATE: the thumbprint OutlookAI.csproj pins, from
          Cert:\CurrentUser\My, with its private key. Expired refuses; under 30 days warns
          (release.yml's "Import signing certificate", minus the import - it is already here).
       5. GUARDS: every Tools/Checks/check-*.ps1 under PowerShell 7 and again under Windows
          PowerShell 5.1, each in its own process (the release ran two of them, the build
          workflow all four under both shells); then check-pinned-constants.ps1
          -ExpectedSigningThumbprint with the certificate that will sign.
       6. D7 (c): Testbed/host/Publish-AddInPayload.ps1 -CompareInstalledTargets and
          Tools/Switch-AddInBuild.ps1 -CompareInstalledTargets - the two read-only comparisons of
          the stand-ins that keep a build from registering the add-in in this Outlook (Q81) with
          Visual Studio's real VSTO targets. The build VM has no Visual Studio, so their
          self-tests skip it there, and a Visual Studio update could reopen that hole without a
          word. Decided by the maintainer 2026-10-03: they run here before every release, the one
          exception to "no self-test on the workstation". Either failing refuses the release.
       7. BUILD (release.yml's "Stamp assembly version" through "Create installer"):
          Testbed/host/Publish-AddInPayload.ps1 -ReleaseSigningThumbprint, from a `git archive`
          of HEAD - the Q81 guards included, the manifests signed with the release key, the MCP
          server published into the payload, the VSTO runtime compiled in.
       8. SIZE AND SIGNATURE (release.yml's size gate and "Sign installer"): signtool with the
          pinned certificate BY THUMBPRINT and an RFC 3161 timestamp from the first server that
          answers; the signature read back the way the shipped updater reads it (WinVerifyTrust,
          accepting the self-signed key's untrusted root, and the signer's thumbprint pinned); and
          the signed file no larger than the updater's 50 MB cap.
       9. TESTS: Testbed/host/Invoke-TestsOnBuildVm.ps1 <HEAD> - the whole non-live suite and
          every script self-test, on the build VM. Anything but exit 0 refuses. Last before the
          stamp because it is the slow one: the cheap failures come first.
      10. STAMP COMMIT (release.yml's "Stamp changelog" and "Commit stamped changelog"): HEAD's
          tree with only CHANGELOG.md changed - "## Unreleased", a blank line, "## v<version> -
          <date>" - committed with git plumbing, so neither the working tree nor a branch moves.
          The diff is checked: two lines added to CHANGELOG.md and nothing else; and the committed
          CHANGELOG.md is read back, because the body's line ranges are true of that text only.
      11. PUBLISH - ONLY WITH -Execute. origin/master must still be HEAD; the stamp commit is pushed
          to master (a fast-forward, never forced - if master moved, the push is refused and
          nothing is released, as in the workflow); `gh release create v<version>` with the
          signed installer, targeting that commit, notes from step 3; the release is read back;
          and a local master that sits on HEAD is fast-forwarded to the stamp commit.

    WITHOUT -Execute IT IS A DRY RUN, AND A DRY RUN IS THE WHOLE RELEASE EXCEPT PUBLISHING: steps
    1-10 run for real - the guards, D7 (c), the build, the signature, the build VM's tests, the
    stamp commit object - and it ends by printing the push and the gh release create it would have
    run. No tag, no release, no push.

    WHAT THE WORKFLOW DID THAT THIS DOES NOT, ON PURPOSE:
      * the build-provenance attestation (actions/attest-build-provenance). It needs the OIDC token
        that exists only inside a GitHub Actions run, so no local release can make one.
      * uploading the installer as a workflow artifact. The release asset is the installer.
      * importing the certificate from the PFX_BASE64 secret. It is in this user's store.
      * downloading the VSTO runtime on every release. It is staged media here (Testbed/MEDIA.md)
        and Publish-AddInPayload.ps1 holds it to the same SHA-256 and length.

    WHAT THIS DOES THAT THE WORKFLOW DID NOT: the four guards, D7 (c) and the build VM's suite gate
    the release; signtool picks the certificate by thumbprint, where the runner's `/a` took the
    only one it had and this machine can hold more; the signature and the stamp diff are read
    back; the tag targets the exact commit that was pushed.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no `??`, and ASCII apart from the
    changelog palette's twelve icons, which is why the file is saved as UTF-8 WITH a byte order mark:
    without one, Windows PowerShell 5.1 reads it in the ANSI code page (Tools/Checks/
    check-powershell-51.ps1, check 2). It needs both installed all the same: step 5 runs the guards
    under each.

.PARAMETER VersionBump
    major.minor.patch, as the workflow's version_bump input: 0.0.1 a patch, 0.1.0 a minor, 1.0.0 a
    major. Required except with -SelfTest. 0.0.0 is refused - every release bumps something.

.PARAMETER Execute
    Publish. Without it the run is a dry run (above).

.PARAMETER VstoRuntimePath
    The staged VSTO runtime redistributable. Default: .work\media\vstor_redist.exe, then
    Redist\vstor_redist.exe - the two places Testbed/MEDIA.md names.

.PARAMETER WorkRoot
    Where releases are built: <WorkRoot>\v<version>. Default .work\release (gitignored).

.PARAMETER RepoRoot
    Repository root. Defaults to the folder above this script's.

.PARAMETER TestTimeoutMinutes
    How long the build VM's run may take, queueing included. Default 200: the runner queues for
    up to 120 minutes behind other callers and gives a run 60.

.PARAMETER SelfTest
    Run the pure decisions against synthetic inputs and exit. No git, no gh, no network, no
    certificate, no Visual Studio - it runs on the build VM like every other self-test. D7 (c)
    itself is not in it: only its wiring is. The wiring includes this repository's own
    CHANGELOG.md, held to the changelog's rules, so a malformed entry fails the build VM's run.

.PARAMETER CheckChangelog
    Check CHANGELOG.md against the changelog's rules (step 3's, minus the release itself) and
    write the release body the next release would get to .work\release\changelog-check\. Reads
    the file and asks git for the newest release tag; builds, tests and publishes nothing, so it
    is for the workstation: run it after writing an entry. Exit 0 clean, 1 a rule is broken.

.EXAMPLE
    pwsh -File Tools/Publish-Release.ps1 -VersionBump 0.0.1             # dry run
    pwsh -File Tools/Publish-Release.ps1 -VersionBump 0.0.1 -Execute    # publishes
    pwsh -File Tools/Publish-Release.ps1 -CheckChangelog                # the changelog alone
    pwsh -File Tools/Publish-Release.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string] $VersionBump,
    [switch] $Execute,
    [string] $VstoRuntimePath,
    [string] $WorkRoot,
    [string] $RepoRoot,
    [int]    $TestTimeoutMinutes = 200,
    [switch] $SelfTest,
    [switch] $CheckChangelog
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}
# The running edition's OWN Security and Utility modules, by path - Testbed/host/OwnEditionModules.ps1
# says why (Windows PowerShell 5.1 started from PowerShell 7 resolves both to 7's copies and loses
# Get-FileHash, the Cert: drive and Get-AuthenticodeSignature). Restated inline, as
# Tools/Switch-AddInBuild.ps1 does, because this script lives outside Testbed/.
foreach ($ownEditionModuleName in @('Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility')) {
    Import-Module (Join-Path $PSHOME "Modules\$ownEditionModuleName\$ownEditionModuleName.psd1")
}

# Defaults that need $PSScriptRoot are set HERE, not in param(): Windows PowerShell 5.1 leaves it
# empty while param() defaults are evaluated (Q78).
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
if (-not $WorkRoot) { $WorkRoot = Join-Path $RepoRoot '.work\release' }

# ---------------------------------------------------------------------------------------------
# Everything this script names, in one place.
# ---------------------------------------------------------------------------------------------
$Remote = 'origin'
$ReleaseBranch = 'master'

# The largest installer the shipped auto-updater accepts (UpdateService.MaxDownloadBytes). An asset
# over it silently stops auto-update on every installed copy, so the release refuses it instead.
# Tools/Checks/check-pinned-constants.ps1 #3 holds this figure equal to the updater's.
$InstallerCapMB = 50

# RFC 3161 timestamp servers, tried in order, so one server's outage does not fail a release.
$TimestampServers = @('http://timestamp.digicert.com', 'http://timestamp.sectigo.com')
$CertificateWarnDays = 30

# GitHub refuses a release body longer than this: "body is too long (maximum is 125000 characters)".
# gh release create runs AFTER the stamp commit is pushed, so a body over it would leave master saying
# a version was released that never was. Step 3 measures the body it generates against it up front.
$ReleaseBodyLimit = 125000

# THE CHANGELOG'S RULES. Decided by the maintainer 2026-10-03 (Q123), in his words: "copy the release
# notes style and system and rules from the BrowserAI repo". Each one below is BrowserAI's - from its
# CHANGELOG.md, RELEASING.md item 10, build/New-ReleaseNotes.ps1 and ChangelogTests - and AGENTS.md
# (Changelog) states them for whoever writes an entry.
#
# The palette an entry opens with, in the order the legend lists it: BrowserAI's, which the maintainer
# approved there (its Q192b). This list is the DECISION and the legend at the top of CHANGELOG.md is the
# PUBLICATION; step 3 holds the two identical. The icons are written as themselves, which is why this
# file carries a byte order mark (Q78). The self-test holds each one to its code points, so an editor
# that drops a variation selector is caught.
$ChangelogPalette = @(
    [pscustomobject]@{ Icon = '✨'; Means = 'new capability' },
    [pscustomobject]@{ Icon = '🐛'; Means = 'fix' },
    [pscustomobject]@{ Icon = '🔧'; Means = 'behaviour or configuration change' },
    [pscustomobject]@{ Icon = '🔒'; Means = 'security or permissions' },
    [pscustomobject]@{ Icon = '🗑️'; Means = 'removal or deprecation' },
    [pscustomobject]@{ Icon = '💥'; Means = 'breaking, or the reader must act' },
    [pscustomobject]@{ Icon = '📝'; Means = 'documentation' },
    [pscustomobject]@{ Icon = '✅'; Means = 'tests and the gate' },
    [pscustomobject]@{ Icon = '📦'; Means = 'packaging, installer, release pipeline' },
    [pscustomobject]@{ Icon = '⚡'; Means = 'performance' },
    [pscustomobject]@{ Icon = '♻️'; Means = 'refactor with no behaviour change' },
    [pscustomobject]@{ Icon = '⬆️'; Means = 'dependency move' }
)
# Keep a Changelog's groups, in the order that format fixes them; each at most once in a section.
$ChangelogGroups = @('Added', 'Changed', 'Deprecated', 'Removed', 'Fixed', 'Security')
# The longest a headline may be, its full stop included. BrowserAI's budget, CHOSEN and not measured:
# about the width of the bold line a reader scans on a release page, and short enough that a headline
# cannot become the entry.
$HeadlineBudget = 100
# How many of its headline's opening words an entry's detail may repeat. An entry about a named thing
# opens with that name; five words in a row is the headline said twice.
$RestatementBudget = 4
# The eight characters a person typing into a text box does not produce, refused in everything the
# release page shows - the preamble, every headline, the legend and the body's own words. A backticked
# code span is exempt: it quotes something that exists.
# Written as [char] numbers, never as themselves: PowerShell's parser reads a curly quote as a quote and
# a dash as a hyphen, so the characters cannot sit in this file's own string literals.
$UntypedCharacters = @(
    [pscustomobject]@{ Character = [char]0x2014; Says = 'an em dash (U+2014); write a hyphen, a comma or two sentences' },
    [pscustomobject]@{ Character = [char]0x2013; Says = 'an en dash (U+2013); write a hyphen or the word "to"' },
    [pscustomobject]@{ Character = [char]0x2018; Says = 'a curly opening quote (U+2018); write a straight apostrophe' },
    [pscustomobject]@{ Character = [char]0x2019; Says = 'a curly closing quote (U+2019); write a straight apostrophe' },
    [pscustomobject]@{ Character = [char]0x201C; Says = 'a curly opening double quote (U+201C); write a straight double quote' },
    [pscustomobject]@{ Character = [char]0x201D; Says = 'a curly closing double quote (U+201D); write a straight double quote' },
    [pscustomobject]@{ Character = [char]0x2026; Says = 'an ellipsis character (U+2026); write three full stops' },
    [pscustomobject]@{ Character = [char]0x00A0; Says = 'a non-breaking space (U+00A0); write an ordinary space' }
)
# Where -CheckChangelog writes the body it would publish.
$ChangelogCheckDirectory = 'changelog-check'

# CERT_E_UNTRUSTEDROOT, 0x800B0109. The certificate is SELF-SIGNED (CN=OutlookAI), so WinVerifyTrust
# ends a perfectly good signature in an untrusted root, and the shipped updater accepts exactly that
# result and 0, then pins the signer (Services/UpdateService.cs, VerifySignature). So does step 8.
$CertEUntrustedRoot = [int]-2146762487

# The updater's WinVerifyTrust call (Services/UpdateService.cs, WinVerifyTrustFile), restated in the
# C# that Windows PowerShell 5.1's Add-Type compiles: the same action, the same flags - no UI, no
# revocation check, WTD_SAFER_FLAG - so step 8 asks the question every installed copy will ask.
$InstallerTrustSource = @'
using System;
using System.Runtime.InteropServices;

namespace OutlookAIRelease
{
    public static class InstallerTrust
    {
        public static int Verify(string path)
        {
            WINTRUST_FILE_INFO fileInfo = new WINTRUST_FILE_INFO();
            fileInfo.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO));
            fileInfo.pcwszFilePath = path;
            fileInfo.hFile = IntPtr.Zero;
            fileInfo.pgKnownSubject = IntPtr.Zero;
            IntPtr pFile = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
            IntPtr pData = IntPtr.Zero;
            try
            {
                Marshal.StructureToPtr(fileInfo, pFile, false);
                WINTRUST_DATA data = new WINTRUST_DATA();
                data.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_DATA));
                data.dwUIChoice = 2;              // WTD_UI_NONE
                data.fdwRevocationChecks = 0;     // WTD_REVOKE_NONE
                data.dwUnionChoice = 1;           // WTD_CHOICE_FILE
                data.pFile = pFile;
                data.dwStateAction = 0;           // WTD_STATEACTION_IGNORE
                data.dwProvFlags = 0x00000100;    // WTD_SAFER_FLAG
                data.dwUIContext = 0;
                pData = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_DATA)));
                Marshal.StructureToPtr(data, pData, false);
                return WinVerifyTrust(IntPtr.Zero, new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE"), pData);
            }
            finally
            {
                if (pData != IntPtr.Zero) Marshal.FreeHGlobal(pData);
                Marshal.DestroyStructure(pFile, typeof(WINTRUST_FILE_INFO));
                Marshal.FreeHGlobal(pFile);
            }
        }

        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, IntPtr pWVTData);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;
            [MarshalAs(UnmanagedType.LPWStr)] public string pcwszFilePath;
            public IntPtr hFile;
            public IntPtr pgKnownSubject;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINTRUST_DATA
        {
            public uint cbStruct;
            public IntPtr pPolicyCallbackData;
            public IntPtr pSIPClientData;
            public uint dwUIChoice;
            public uint fdwRevocationChecks;
            public uint dwUnionChoice;
            public IntPtr pFile;
            public uint dwStateAction;
            public IntPtr hWVTStateData;
            public IntPtr pwszURLReference;
            public uint dwProvFlags;
            public uint dwUIContext;
            public IntPtr pSignatureSettings;
        }
    }
}
'@

# The guards, and the two read-only comparisons D7 (c) runs here.
$GuardDirectory = 'Tools\Checks'
$GuardPattern = 'check-*.ps1'
$ExpectedGuardCount = 4
$PinnedConstantsGuard = 'check-pinned-constants.ps1'
$D7Comparisons = @('Testbed\host\Publish-AddInPayload.ps1', 'Tools\Switch-AddInBuild.ps1')
$D7Switch = 'CompareInstalledTargets'
$BuildScript = 'Testbed\host\Publish-AddInPayload.ps1'
$TestRunner = 'Testbed\host\Invoke-TestsOnBuildVm.ps1'

function Say([string] $m) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }

# Every native call goes through this: under Windows PowerShell 5.1 with 'Stop', the first line a
# program writes to stderr ends the script when anything redirects it, and this script is as
# likely to run from a job or behind a redirection as from a console (Tools/Checks/
# check-powershell-51.ps1, check 3). 'Continue' holds only in here; the try keeps "program not
# found" terminating; callers judge the call by $LASTEXITCODE. Restated, as in the other scripts.
function Invoke-NativeCommand {
    param([Parameter(Mandatory = $true)] [scriptblock] $NativeCommand)

    $ErrorActionPreference = 'Continue'
    try {
        & $NativeCommand | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { $_ }
        }
    }
    catch {
        throw
    }
}

# Start a program with its output going STRAIGHT TO FILES, never through this process's pipe, and a
# deadline: a build that leaves children behind can hold a pipe open after it has finished, and a
# declared timeout then never fires. Restated from Testbed/host/Publish-AddInPayload.ps1.
function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)] [string] $FilePath,
        [Parameter(Mandatory = $true)] [string[]] $ArgumentList,
        [Parameter(Mandatory = $true)] [string] $LogStem,
        [int] $TimeoutMinutes = 20,
        [string] $WorkingDirectory
    )
    $out = "$LogStem.out.txt"
    $err = "$LogStem.err.txt"
    foreach ($f in @($out, $err)) { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force } }
    $startArgs = @{
        FilePath               = $FilePath
        ArgumentList           = ($ArgumentList -join ' ')
        RedirectStandardOutput = $out
        RedirectStandardError  = $err
        NoNewWindow            = $true
        PassThru               = $true
    }
    if ($WorkingDirectory) { $startArgs['WorkingDirectory'] = $WorkingDirectory }
    $p = Start-Process @startArgs
    $null = $p.Handle
    $exited = $p.WaitForExit($TimeoutMinutes * 60 * 1000)
    if ($exited) { $p.WaitForExit() }
    if (-not $exited) {
        try { $p.Kill() } catch { }
        return [pscustomobject]@{ ExitCode = -1; TimedOut = $true; Out = $out; Err = $err }
    }
    return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false; Out = $out; Err = $err }
}

function Show-LogTail([string] $Path, [int] $Lines = 30) {
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        foreach ($l in (Get-Content -LiteralPath $Path -Tail $Lines)) { Write-Host "      | $l" }
    }
}

# One argument for Start-Process's joined command line.
function Format-Argument([string] $Value) {
    if ($Value -match '\s') { return '"' + $Value + '"' }
    return $Value
}

# =============================================================================================
# PURE DECISIONS - everything -SelfTest checks
# =============================================================================================

# release.yml's "Validate version bump input". $null when it is acceptable.
function Test-VersionBumpText {
    param([string] $Bump)
    if (-not $Bump) { return '-VersionBump is required: major.minor.patch, e.g. 0.0.1 for a patch, 0.1.0 for a minor, 1.0.0 for a major release.' }
    if ($Bump -notmatch '^\d+\.\d+\.\d+$') { return "Invalid -VersionBump '$Bump'. It must be major.minor.patch (e.g. 1.0.0, 0.1.0, 0.0.1)." }
    if ($Bump -eq '0.0.0') { return 'A -VersionBump of 0.0.0 is not valid: every release must bump at least the patch version (e.g. 0.0.1).' }
    return $null
}

# The base version a release tag carries: v3.1.0.325 -> 3.1.0. A tag that is not one - or no release
# at all - is 0.0.0, as in the workflow.
function Get-BaseVersionFromTag {
    param([string] $Tag)
    if ($Tag -match '^v?(\d+\.\d+\.\d+)') { return $Matches[1] }
    return '0.0.0'
}

# release.yml's arithmetic: major resets minor and patch, minor resets patch, a patch adds.
function Get-NextBaseVersion {
    param([string] $Current, [string] $Bump)
    $cv = $Current -split '\.'
    $b = $Bump -split '\.'
    $bumpMajor = [int]$b[0]
    $bumpMinor = [int]$b[1]
    $bumpPatch = [int]$b[2]
    $newMajor = [int]$cv[0] + $bumpMajor
    if ($bumpMajor -gt 0) { $newMinor = $bumpMinor } else { $newMinor = [int]$cv[1] + $bumpMinor }
    if ($bumpMajor -gt 0 -or $bumpMinor -gt 0) { $newPatch = $bumpPatch } else { $newPatch = [int]$cv[2] + $bumpPatch }
    return "$newMajor.$newMinor.$newPatch"
}

# The full version: the base plus HEAD's commit count PLUS ONE, the stamp commit this run makes. A
# part over 65535 is not an assembly version, so it is refused rather than truncated by the build.
function Get-FullVersion {
    param([string] $Base, [int] $CommitCount)
    $full = "$Base.$($CommitCount + 1)"
    foreach ($part in ($full -split '\.')) {
        if ([long]$part -gt 65535) { throw "Version $full has a part over 65535, which no assembly version can hold." }
    }
    return $full
}

# release.yml's "Extract changelog": the text between "## Unreleased" and the next "## " heading.
function Get-UnreleasedNotes {
    param([string] $ChangelogText)
    if ($ChangelogText -match '(?s)## Unreleased\r?\n(.*?)(?=\r?\n## |\z)') { return $Matches[1].Trim() }
    return ''
}

# A body GitHub would refuse. New-ReleaseBody has already dropped the per-entry links to get under the
# limit, so this is a release whose HEADLINES alone do not fit. Publishing refuses it; a dry run notes
# it and goes on.
function Get-NotesLengthVerdict {
    param([string] $Notes, [int] $Limit, [bool] $Publishing)
    if ($Notes.Length -le $Limit) { return [pscustomobject]@{ Problem = $null; Note = $null } }
    $text = "the release body is $($Notes.Length) characters with every per-entry link dropped, and GitHub refuses a release body over $Limit ('body is too long'). gh release create would fail after the stamp commit had been pushed. Release before the section grows this large, or shorten its headlines."
    if ($Publishing) { return [pscustomobject]@{ Problem = $text; Note = $null } }
    return [pscustomobject]@{ Problem = $null; Note = $text + ' -Execute would refuse; the dry run goes on.' }
}

# A release page opens with the section's preamble - the prose above its first group, written when the
# release is cut, for somebody who has never seen OutlookAI. Publishing refuses a section without one; a
# dry run notes it and goes on.
function Get-PreambleVerdict {
    param([string] $Preamble, [bool] $Publishing)
    if ($Preamble -and $Preamble.Trim()) { return [pscustomobject]@{ Problem = $null; Note = $null } }
    $text = "the Unreleased section has no preamble, so the release page would open with its first entry. Write a few short paragraphs above its first group - what OutlookAI is and what this release changes, for somebody who has never seen it (AGENTS.md, Changelog)."
    if ($Publishing) { return [pscustomobject]@{ Problem = $text; Note = $null } }
    return [pscustomobject]@{ Problem = $null; Note = $text + ' -Execute would refuse; the dry run goes on.' }
}

# release.yml's "Stamp changelog": "## Unreleased" stays, empty, and the entries under it become the
# new version's. Only the first such heading, only at the start of a line, and in the file's own
# line endings. $null when there is no such heading.
function Set-ChangelogStamp {
    param([string] $ChangelogText, [string] $Version, [string] $Date)
    $m = [regex]::Match($ChangelogText, '(?m)^## Unreleased(?<eol>\r?)$')
    if (-not $m.Success) { return $null }
    $nl = $m.Groups['eol'].Value + "`n"
    $insert = $nl + $nl + "## v$Version - $Date"
    return $ChangelogText.Substring(0, $m.Index + '## Unreleased'.Length) + $insert + $ChangelogText.Substring($m.Index + '## Unreleased'.Length)
}

# The stamp commit must differ from HEAD in CHANGELOG.md alone, by two added lines: the blank line
# and the version heading. `git diff --numstat` lines in, problems out.
function Test-StampDiff {
    param([string[]] $NumstatLines)
    $lines = @($NumstatLines | Where-Object { $_ -and $_.Trim() })
    if ($lines.Count -ne 1) { return "the stamp commit changes $($lines.Count) file(s); it must change CHANGELOG.md alone." }
    if ($lines[0] -notmatch '^2\t0\tCHANGELOG\.md$') { return "the stamp commit's change is '$($lines[0])'; it must be two lines added to CHANGELOG.md and none removed." }
    return $null
}

# ---------------------------------------------------------------------------------------------
# THE CHANGELOG AND THE RELEASE BODY (Q123) - BrowserAI's system, ported
# ---------------------------------------------------------------------------------------------
# CHANGELOG.md is the record: every entry an icon, a bold one-sentence headline and the whole of what
# happened. The GitHub release body is MADE from it and is not it: the section's preamble, each entry's
# icon and headline with a "read more" link to that entry's own lines in the tagged file, the legend,
# and a link to the whole section. A line range points AT the record instead of copying it, so the body
# stays a page however much a release holds - BrowserAI measured its 1.0.0 at 41,288 characters linked,
# where the same section folded into the body was 288,437.

# The text with LF line ends, so line i of its split is line i+1 of the file whatever its endings.
function ConvertTo-LfText {
    param([string] $Text)
    return ($Text -replace "`r`n", "`n")
}

# The legend: the last block above the first "## " heading in which every line is a table row - a block
# and not a line scan, because the prose above it may mention a pipe. $null when there is none. Rows are
# the body rows' cells, Pairs the icon-and-meaning pairs in reading order, Table the block as written.
function Get-ChangelogLegend {
    param([string] $ChangelogText)
    $text = ConvertTo-LfText $ChangelogText
    $first = [regex]::Match($text, '(?m)^## ')
    $head = $text
    if ($first.Success) { $head = $text.Substring(0, $first.Index) }
    $block = $null
    foreach ($paragraph in ($head -split "`n`n")) {
        $rows = @(($paragraph -split "`n") | Where-Object { $_.Trim().Length -gt 0 })
        $prose = @($rows | Where-Object { -not $_.TrimStart().StartsWith('|') })
        if ($rows.Count -ge 3 -and $prose.Count -eq 0) { $block = $rows }
    }
    if ($null -eq $block) { return $null }
    $cells = New-Object System.Collections.ArrayList
    $pairs = New-Object System.Collections.ArrayList
    foreach ($row in @($block | Select-Object -Skip 2)) {
        $c = @($row.Trim().Trim('|').Split('|') | ForEach-Object { $_.Trim() })
        $null = $cells.Add($c)
        for ($i = 0; $i + 1 -lt $c.Count; $i += 2) {
            if ($c[$i]) { $null = $pairs.Add($c[$i] + ' ' + $c[$i + 1]) }
        }
    }
    return [pscustomobject]@{ Table = (@($block | ForEach-Object { $_.Trim() }) -join "`n"); Rows = $cells.ToArray(); Pairs = $pairs.ToArray() }
}

# What is wrong with the legend: nothing when it is the approved palette, in its order, as a table of two
# icon-and-meaning pairs per row - six rows for the twelve. The body ends with this table as written.
function Get-LegendProblems {
    param([string] $ChangelogText)
    $legend = Get-ChangelogLegend $ChangelogText
    if ($null -eq $legend) {
        return @("CHANGELOG.md has no icon legend above its first '## ' heading: a Markdown table - a heading row, a delimiter row and one row per two icons - which every release body ends with.")
    }
    $problems = @()
    $expected = @($ChangelogPalette | ForEach-Object { $_.Icon + ' ' + $_.Means }) -join ' | '
    $published = @($legend.Pairs) -join ' | '
    if (-not [string]::Equals($expected, $published, [StringComparison]::Ordinal)) {
        $problems += "CHANGELOG.md's legend is not the approved palette in its order (`$ChangelogPalette in Tools/Publish-Release.ps1). It lists: $published"
    }
    $rows = @($legend.Rows)
    $short = @($rows | Where-Object { @($_).Count -ne 4 })
    if ($rows.Count -ne 6 -or $short.Count -gt 0) {
        $problems += "CHANGELOG.md's legend is not six rows of two icon-and-meaning pairs each."
    }
    return $problems
}

# One section of the changelog, read the way a release body is made from it: its heading, the prose above
# its first group (the preamble), its groups in order, and every entry with its first and last LINE
# NUMBER IN THE FILE. An entry is a line starting "- " and every following line that is blank or
# indented, and it ends on its last line that carries text, so a range never highlights the blank line
# before the next entry. An entry above the first group, and a paragraph under a group that is not an
# entry, are Problems: the body is made group by group, and has no place for either. $null when no line
# matches $HeadingPattern.
function Read-ChangelogSection {
    param([string] $ChangelogText, [string] $HeadingPattern)
    $lines = (ConvertTo-LfText $ChangelogText) -split "`n"
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match $HeadingPattern) { $start = $i; break }
    }
    if ($start -lt 0) { return $null }
    $preamble = New-Object System.Collections.ArrayList
    $groups = New-Object System.Collections.ArrayList
    $entries = New-Object System.Collections.ArrayList
    $problems = New-Object System.Collections.ArrayList
    $group = $null
    $entry = $null
    for ($i = $start + 1; $i -le $lines.Count; $i++) {
        $line = $null
        if ($i -lt $lines.Count) { $line = $lines[$i] }
        $isEnd = ($null -eq $line) -or $line.StartsWith('## ')
        $groupMatch = $null
        if (-not $isEnd) { $groupMatch = [regex]::Match($line, '^### +(\S.*?)\s*$') }
        $isGroup = ($null -ne $groupMatch) -and $groupMatch.Success
        $isEntry = (-not $isEnd) -and ($line.StartsWith('- ') -or $line.StartsWith("-`t"))
        $continues = ($null -ne $entry) -and -not ($isEnd -or $isGroup -or $isEntry)
        if ($continues -and ($line.Trim().Length -eq 0 -or $line.StartsWith(' ') -or $line.StartsWith("`t"))) {
            $null = $entry.Lines.Add($line)
            if ($line.Trim().Length -gt 0) { $entry.Last = $i + 1 }
            continue
        }
        if ($null -ne $entry) {
            if ($null -eq $entry.Group) {
                $null = $problems.Add("CHANGELOG.md:$($entry.First): an entry above the first group heading of '$($lines[$start])'. Every entry sits under one of $($ChangelogGroups -join ', ').")
            }
            $null = $entries.Add($entry)
            $entry = $null
        }
        if ($isEnd) { break }
        if ($isGroup) {
            $group = $groupMatch.Groups[1].Value
            $null = $groups.Add([pscustomobject]@{ Name = $group; Line = $i + 1 })
            continue
        }
        if ($isEntry) {
            $entry = [pscustomobject]@{ First = $i + 1; Last = $i + 1; Group = $group; Lines = (New-Object System.Collections.ArrayList) }
            $null = $entry.Lines.Add($line)
            continue
        }
        if ($null -eq $group) {
            $null = $preamble.Add($line)
            continue
        }
        if ($line.Trim().Length -gt 0) {
            $null = $problems.Add("CHANGELOG.md:$($i + 1): a paragraph under '### $group' that is not an entry. A group holds entries and nothing else; prose belongs in the preamble, above the first group.")
        }
    }
    return [pscustomobject]@{
        Heading     = $lines[$start]
        HeadingLine = $start + 1
        Preamble    = ((@($preamble.ToArray()) -join "`n").Trim())
        Groups      = $groups.ToArray()
        Entries     = $entries.ToArray()
        Problems    = $problems.ToArray()
    }
}

# An entry's lines as one line with its whitespace collapsed: the file wraps an entry, and its shape is
# read whole.
function Join-ChangelogEntry {
    param($Entry)
    return ((@($Entry.Lines) -join ' ') -replace '\s+', ' ').Trim()
}

# An entry's icon, headline and detail, or $null when it is not "- <icon> **Headline.** the detail".
function Split-ChangelogEntry {
    param([string] $Joined)
    $m = [regex]::Match($Joined, '^-\s+(?<icon>\S+)\s+\*\*(?<headline>.+?)\*\*\s*(?<detail>.*)$')
    if (-not $m.Success) { return $null }
    return [pscustomobject]@{ Icon = $m.Groups['icon'].Value; Headline = $m.Groups['headline'].Value.Trim(); Detail = $m.Groups['detail'].Value }
}

# How many words in a row a detail's opening shares with its headline's, case and markup aside.
function Get-SharedOpeningWords {
    param([string] $Headline, [string] $Detail)
    $left = @([regex]::Matches(($Headline -replace '[`*]', ''), "[A-Za-z0-9']+") | ForEach-Object { $_.Value.ToUpperInvariant() })
    $right = @([regex]::Matches(($Detail -replace '[`*]', ''), "[A-Za-z0-9']+") | ForEach-Object { $_.Value.ToUpperInvariant() })
    $n = 0
    while ($n -lt $left.Count -and $n -lt $right.Count -and [string]::Equals($left[$n], $right[$n], [StringComparison]::Ordinal)) { $n++ }
    return $n
}

# What is wrong with one entry, given its joined text and the line it starts on. The icon is compared
# ORDINALLY: a culture comparison ignores the variation selector three of the icons carry.
function Get-EntryProblems {
    param([string] $Joined, [int] $Line)
    $where = "CHANGELOG.md:$Line"
    $excerpt = $Joined
    if ($excerpt.Length -gt 90) { $excerpt = $excerpt.Substring(0, 90) + '...' }
    $e = Split-ChangelogEntry $Joined
    if ($null -eq $e) { return @("${where}: not in the shape '- <icon> **One-sentence headline.** the detail' - $excerpt") }
    $problems = @()
    $known = @($ChangelogPalette | Where-Object { [string]::Equals($_.Icon, $e.Icon, [StringComparison]::Ordinal) })
    if ($known.Count -eq 0) { $problems += "${where}: '$($e.Icon)' is not one of the legend's twelve icons - $excerpt" }
    $h = $e.Headline
    if (-not $h.EndsWith('.')) { $problems += "${where}: the headline does not end in a full stop - $h" }
    elseif ($h.Substring(0, $h.Length - 1) -match '[.!?]\s') { $problems += "${where}: the headline is more than one sentence - $h" }
    if ($h.Length -gt $HeadlineBudget) { $problems += "${where}: the headline is $($h.Length) characters, over the $HeadlineBudget a headline may be - $h" }
    $shared = Get-SharedOpeningWords $h $e.Detail
    if ($shared -gt $RestatementBudget) { $problems += "${where}: the detail opens with $shared of the headline's own words in a row, which is the headline said twice - $excerpt" }
    return $problems
}

# Group headings that are not Keep a Changelog's, appear twice, or run out of its order.
function Get-GroupProblems {
    param($Groups, [string] $Heading)
    $problems = @()
    $seen = New-Object System.Collections.ArrayList
    foreach ($g in @($Groups)) {
        $name = [string]$g.Name
        $index = [array]::IndexOf($ChangelogGroups, $name)
        if ($index -lt 0) { $problems += "CHANGELOG.md:$($g.Line): '### $name' is not a Keep a Changelog group: $($ChangelogGroups -join ', ')."; continue }
        if ($seen.Contains($name)) { $problems += "CHANGELOG.md:$($g.Line): '### $name' appears twice under '$Heading'; its entries belong in one list."; continue }
        if ($seen.Count -gt 0) {
            $previous = [string]$seen[$seen.Count - 1]
            if ($index -lt [array]::IndexOf($ChangelogGroups, $previous)) { $problems += "CHANGELOG.md:$($g.Line): '### $name' comes after '### $previous', out of Keep a Changelog's order: $($ChangelogGroups -join ', ')." }
        }
        $null = $seen.Add($name)
    }
    return $problems
}

# Every character of $UntypedCharacters in $Text, outside its code spans, as sentences naming $Where.
function Get-UntypedCharacterProblems {
    param([string] $Text, [string] $Where)
    $plain = [regex]::Replace($Text, '`[^`]*`', '')
    $problems = @()
    foreach ($u in $UntypedCharacters) {
        if ($plain.IndexOf([string]$u.Character, [StringComparison]::Ordinal) -ge 0) { $problems += "$Where carries $($u.Says)." }
    }
    return $problems
}

# Everything step 3 refuses about the Unreleased section, as sentences, with what it found: the number of
# entries and the preamble. The legend, the groups and every entry are read; the preamble and every
# headline are what the release page shows, so they are held to the untyped-character rule too. A
# detail never reaches the page, and is not.
function Test-UnreleasedSection {
    param([string] $ChangelogText)
    $problems = @(Get-LegendProblems $ChangelogText)
    $section = Read-ChangelogSection $ChangelogText '^## Unreleased\s*$'
    if ($null -eq $section) {
        return [pscustomobject]@{ Problems = @($problems + "CHANGELOG.md has no '## Unreleased' heading."); Entries = 0; Preamble = '' }
    }
    $problems += @($section.Problems)
    $problems += @(Get-GroupProblems $section.Groups $section.Heading)
    foreach ($e in @($section.Entries)) {
        $joined = Join-ChangelogEntry $e
        $problems += @(Get-EntryProblems -Joined $joined -Line $e.First)
        $parts = Split-ChangelogEntry $joined
        if ($null -ne $parts) { $problems += @(Get-UntypedCharacterProblems $parts.Headline "The headline at CHANGELOG.md:$($e.First)") }
    }
    if ($section.Preamble) { $problems += @(Get-UntypedCharacterProblems $section.Preamble "The Unreleased section's preamble") }
    return [pscustomobject]@{ Problems = $problems; Entries = @($section.Entries).Count; Preamble = $section.Preamble }
}

# The anchor GitHub gives a heading, by the rule BrowserAI's link checker applies: code spans lose their
# angle brackets, a link becomes its text, inline HTML goes, code and emphasis markers go; then letters,
# digits, hyphens and underscores stay in lower case, a space becomes a hyphen, and the rest goes.
function Get-GitHubAnchor {
    param([string] $Heading)
    $text = $Heading
    $spans = [regex]::Matches($text, '`([^`]*)`')
    for ($k = $spans.Count - 1; $k -ge 0; $k--) {
        $s = $spans[$k]
        $text = $text.Substring(0, $s.Index) + ($s.Groups[1].Value -replace '[<>]', '') + $text.Substring($s.Index + $s.Length)
    }
    $text = [regex]::Replace($text, '\[([^\]]*)\]\([^)]*\)', '$1')
    $text = [regex]::Replace($text, '<[^>]*>', '')
    $text = [regex]::Replace($text, '[`*]', '')
    $slug = New-Object System.Text.StringBuilder
    foreach ($ch in $text.Trim().ToCharArray()) {
        if ([char]::IsLetterOrDigit($ch) -or $ch -eq [char]'-' -or $ch -eq [char]'_') { $null = $slug.Append([char]::ToLowerInvariant($ch)) }
        elseif ($ch -eq [char]' ' -or $ch -eq [char]"`t") { $null = $slug.Append('-') }
        else {
            $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
            if ($category -eq [System.Globalization.UnicodeCategory]::NonSpacingMark -or
                $category -eq [System.Globalization.UnicodeCategory]::SpacingCombiningMark -or
                $category -eq [System.Globalization.UnicodeCategory]::EnclosingMark) { $null = $slug.Append($ch) }
        }
    }
    return $slug.ToString()
}

# The GitHub release body for $Version, made from the STAMPED changelog - the text the stamp commit
# carries and the tag will point at - so every line range is true of the file a reader opens. ONE shape,
# BrowserAI's: the preamble, then each group's entries as icon, headline and a "read more" link to that
# entry's own lines (CHANGELOG.md?plain=1#L<first>-L<last>, the source view, which highlights exactly
# those lines), then the legend and a link to the whole section. Over $Limit the links go and the
# headlines stay - a fallback for a release several times the size of any either project has cut - and
# Shape says which body this is.
function New-ReleaseBody {
    param([string] $StampedText, [string] $Version, [string] $RepositoryUrl, [int] $Limit)
    $result = [pscustomobject]@{ Body = ''; Shape = ''; Entries = 0; Groups = 0; LinkedLength = 0; Problems = @() }
    $section = Read-ChangelogSection $StampedText ('^## v' + [regex]::Escape($Version) + ' - ')
    if ($null -eq $section) {
        $result.Problems = @("the stamped CHANGELOG.md has no '## v$Version - ' section to make a release body from.")
        return $result
    }
    $problems = @($section.Problems)
    $legend = Get-ChangelogLegend $StampedText
    if ($null -eq $legend) { $problems += 'CHANGELOG.md has no icon legend table for the release body to end with.' }
    $items = New-Object System.Collections.ArrayList
    foreach ($e in @($section.Entries)) {
        $parts = Split-ChangelogEntry (Join-ChangelogEntry $e)
        if ($null -eq $parts) { $problems += "CHANGELOG.md:$($e.First): not in the shape '- <icon> **One-sentence headline.** the detail'."; continue }
        $null = $items.Add([pscustomobject]@{ Group = $e.Group; Icon = $parts.Icon; Headline = $parts.Headline; First = $e.First; Last = $e.Last })
    }
    if ($problems.Count -gt 0) {
        $result.Problems = $problems
        return $result
    }
    $file = $RepositoryUrl + '/blob/v' + $Version + '/CHANGELOG.md'
    $anchor = Get-GitHubAnchor ($section.Heading -replace '^##\s+', '')
    $paragraphs = @()
    if ($section.Preamble) {
        $paragraphs = @(($section.Preamble -split "\n[ \t]*\n") | ForEach-Object { ($_ -replace '\s+', ' ').Trim() } | Where-Object { $_.Length -gt 0 })
    }
    $used = New-Object System.Collections.ArrayList
    foreach ($g in @($section.Groups)) {
        $name = [string]$g.Name
        if ($used.Contains($name)) { continue }
        if (@($items | Where-Object { [string]::Equals([string]$_.Group, $name, [StringComparison]::Ordinal) }).Count -gt 0) { $null = $used.Add($name) }
    }
    $bodies = @{}
    foreach ($linked in @($true, $false)) {
        $out = New-Object System.Collections.ArrayList
        foreach ($p in $paragraphs) { $null = $out.Add($p); $null = $out.Add('') }
        foreach ($name in $used) {
            $null = $out.Add('### ' + $name)
            $null = $out.Add('')
            foreach ($it in @($items | Where-Object { [string]::Equals([string]$_.Group, $name, [StringComparison]::Ordinal) })) {
                $line = '- ' + $it.Icon + ' **' + $it.Headline + '**'
                if ($linked) { $line += ' [read more](' + $file + '?plain=1#L' + $it.First + '-L' + $it.Last + ')' }
                $null = $out.Add($line)
            }
            $null = $out.Add('')
        }
        $null = $out.Add('---')
        $null = $out.Add('')
        $null = $out.Add($legend.Table)
        $null = $out.Add('')
        $null = $out.Add('The full changelog for this release: [CHANGELOG.md](' + $file + '#' + $anchor + ')')
        $bodies[$linked] = ((@($out.ToArray()) -join "`n") -replace "`n{3,}", "`n`n").Trim() + "`n"
    }
    $result.Body = $bodies[$true]
    $result.Shape = 'linked'
    if ($result.Body.Length -gt $Limit) {
        $result.Body = $bodies[$false]
        $result.Shape = 'headlines'
    }
    $result.Entries = $items.Count
    $result.Groups = $used.Count
    $result.LinkedLength = $bodies[$true].Length
    return $result
}

# Everything from the first released heading ("## v") to the end of the file: the part of the changelog
# nothing may change.
function Get-ReleasedChangelog {
    param([string] $ChangelogText)
    $text = ConvertTo-LfText $ChangelogText
    $m = [regex]::Match($text, '(?m)^## v')
    if (-not $m.Success) { return '' }
    return $text.Substring($m.Index).TrimEnd()
}

# $null when the released sections are exactly what the latest release's tag carries, and otherwise the
# section where they part. BrowserAI seals each released section by length and SHA-256 in a table; here
# every release's stamp commit is the seal on everything below it, so the table is the tags themselves.
function Test-ReleasedSectionsUnchanged {
    param([string] $ChangelogText, [string] $TaggedText, [string] $Tag)
    $now = Get-ReleasedChangelog $ChangelogText
    $then = Get-ReleasedChangelog $TaggedText
    if ([string]::Equals($now, $then, [StringComparison]::Ordinal)) { return $null }
    $a = @($now -split '(?m)^(?=## )')
    $b = @($then -split '(?m)^(?=## )')
    $where = ''
    for ($k = 0; $k -lt [Math]::Max($a.Count, $b.Count); $k++) {
        $x = ''
        $y = ''
        if ($k -lt $a.Count) { $x = $a[$k].TrimEnd() }
        if ($k -lt $b.Count) { $y = $b[$k].TrimEnd() }
        if (-not [string]::Equals($x, $y, [StringComparison]::Ordinal)) {
            $where = ($x -split "`n")[0]
            if (-not $where) { $where = ($y -split "`n")[0] }
            break
        }
    }
    return "CHANGELOG.md's released sections differ from what $Tag carries, from '$where' on. A released section is a record and is never rewritten: put it back as $Tag has it (git diff $Tag -- CHANGELOG.md)."
}

# owner/name from the origin URL, https or ssh.
function Get-GitHubRepository {
    param([string] $RemoteUrl)
    if ($RemoteUrl -match 'github\.com[:/](?<repo>[^/\s]+/[^/\s]+?)(?:\.git)?/?$') { return $Matches['repo'] }
    return $null
}

# HEAD against origin/master. Publishing needs them equal; a dry run only notes the difference.
function Get-HeadVerdict {
    param([string] $Head, [string] $RemoteTip, [bool] $Publishing)
    if ($Head -eq $RemoteTip) { return [pscustomobject]@{ Problem = $null; Note = $null } }
    $text = "HEAD $Head is not $Remote/$ReleaseBranch ($RemoteTip). A release is master's tip, and its stamp commit goes onto master as a fast-forward."
    if ($Publishing) { return [pscustomobject]@{ Problem = $text; Note = $null } }
    return [pscustomobject]@{ Problem = $null; Note = $text + ' -Execute would refuse; the dry run goes on.' }
}

# The certificate's remaining life: expired refuses, under $CertificateWarnDays warns.
function Get-CertificateVerdict {
    param([datetime] $NotAfter, [datetime] $Now)
    $days = [int][math]::Floor(($NotAfter - $Now).TotalDays)
    if ($NotAfter -le $Now) { return [pscustomobject]@{ Problem = "the signing certificate EXPIRED on $($NotAfter.ToString('yyyy-MM-dd')). Renew it - and ship the new thumbprint in an update first - before releasing."; Warning = $null; Days = $days } }
    if ($days -lt $CertificateWarnDays) { return [pscustomobject]@{ Problem = $null; Warning = "the signing certificate expires in $days day(s), on $($NotAfter.ToString('yyyy-MM-dd'))."; Days = $days } }
    return [pscustomobject]@{ Problem = $null; Warning = $null; Days = $days }
}

function Test-InstallerSize {
    param([long] $Bytes, [int] $CapMB)
    if ($Bytes -gt ([long]$CapMB * 1MB)) {
        return "the installer is $([math]::Round($Bytes / 1MB, 2)) MB, over the auto-updater's $CapMB MB download cap. Raise UpdateService.MaxDownloadBytes and ship that first."
    }
    return $null
}

# The updater's acceptance rule, restated: WinVerifyTrust 0 or CERT_E_UNTRUSTEDROOT, and the signer
# pinned. Plus a timestamp, without which the signature would die with the certificate. $null when
# every installed copy would accept the installer.
function Get-SignatureVerdict {
    param([int] $TrustResult, [string] $SignerThumbprint, [string] $PinnedThumbprint, [bool] $Timestamped)
    if ($TrustResult -ne 0 -and $TrustResult -ne $CertEUntrustedRoot) {
        return ('WinVerifyTrust answers 0x{0:X8}: the signature is invalid or missing, and every installed copy would refuse this installer.' -f $TrustResult)
    }
    if ($SignerThumbprint -ne $PinnedThumbprint) { return "the installer is signed by '$SignerThumbprint', not the pinned $PinnedThumbprint - every installed copy would refuse it." }
    if (-not $Timestamped) { return 'the signature carries no timestamp, so it would stop being valid when the certificate expires.' }
    return $null
}

# signtool, by thumbprint - never /a, which picks whichever certificate it likes best in the store.
function Get-SignToolArgumentList {
    param([string] $Thumbprint, [string] $TimestampUrl, [string] $File)
    return @('sign', '/sha1', $Thumbprint, '/fd', 'SHA256', '/tr', $TimestampUrl, '/td', 'SHA256', (Format-Argument $File))
}

function Get-GhReleaseArgumentList {
    param([string] $Repository, [string] $Tag, [string] $Asset, [string] $Target, [string] $NotesFile)
    return @('release', 'create', $Tag, $Asset, '--repo', $Repository, '--target', $Target,
        '--title', "OutlookAI $Tag", '--notes-file', $NotesFile)
}

# The pinned thumbprint, as OutlookAI.csproj holds it (check-pinned-constants.ps1 #1 holds the
# updater's copy equal to it).
function Get-PinnedThumbprint {
    param([string] $ProjectText)
    $m = [regex]::Match($ProjectText, '<ManifestCertificateThumbprint>\s*([0-9A-Fa-f]{40})\s*</ManifestCertificateThumbprint>')
    if (-not $m.Success) { return $null }
    return $m.Groups[1].Value.ToUpperInvariant()
}

# The type of a script's parameter, read from its param() block by the parser and never by running
# the script - how the self-test proves D7 (c)'s wiring where D7 (c) itself cannot run. $null when the
# script does not parse or declares no such parameter.
function Get-ScriptParameterType {
    param([string] $Path, [string] $Name)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0 -or -not $ast.ParamBlock) { return $null }
    foreach ($p in $ast.ParamBlock.Parameters) {
        if ($p.Name.VariablePath.UserPath -eq $Name) { return $p.StaticType.Name }
    }
    return $null
}

# The results directory Invoke-TestsOnBuildVm.ps1 names in its log ("results   <dir>").
function Get-TestRunDirectory {
    param([string[]] $LogLines)
    foreach ($l in @($LogLines)) {
        if ($l -match '\]\s+results\s+(?<dir>\S.*)$') { return $Matches['dir'].Trim() }
    }
    return $null
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()

    # Compared with -ceq. Never -like: a bracket in a -like pattern is a character class
    # (Testbed/README.md section 4b; AGENTS.md mailbox rule 2).
    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:Checks++
        $e = [string]$Expected
        $a = [string]$Actual
        if ($Expected -is [System.Array]) { $e = ($Expected -join ' | ') }
        if ($Actual -is [System.Array]) { $a = ($Actual -join ' | ') }
        if ($e -ceq $a) { Write-Host "  OK   $What" }
        else {
            $script:Failures += "$What : expected [$e], got [$a]"
            Write-Host "  FAIL $What - expected [$e], got [$a]"
        }
    }

    Write-Host "Tools/Publish-Release.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    Write-Host ''
    Write-Host '== the version bump (release.yml''s "Validate version bump input") =='
    Test-Case 'a patch bump is accepted' $true ($null -eq (Test-VersionBumpText '0.0.1'))
    Test-Case 'a major bump is accepted' $true ($null -eq (Test-VersionBumpText '1.0.0'))
    Test-Case 'no bump is refused' $true ([string](Test-VersionBumpText '')).Contains('required')
    Test-Case '0.0.0 is refused' $true ([string](Test-VersionBumpText '0.0.0')).Contains('at least the patch')
    Test-Case 'two parts are refused' $true ([string](Test-VersionBumpText '0.1')).Contains('major.minor.patch')
    Test-Case 'four parts are refused' $true ([string](Test-VersionBumpText '0.0.1.0')).Contains('major.minor.patch')
    Test-Case 'a v prefix is refused' $true ([string](Test-VersionBumpText 'v0.0.1')).Contains('major.minor.patch')

    Write-Host ''
    Write-Host '== the version (release.yml''s "Apply version bump") =='
    Test-Case 'the base of the latest tag' '3.1.0' (Get-BaseVersionFromTag 'v3.1.0.325')
    Test-Case 'a tag without the v' '2.3.4' (Get-BaseVersionFromTag '2.3.4.145')
    Test-Case 'no release at all is 0.0.0' '0.0.0' (Get-BaseVersionFromTag '')
    Test-Case 'a tag that is not a version is 0.0.0' '0.0.0' (Get-BaseVersionFromTag 'nightly')
    Test-Case 'a patch adds to the patch' '3.1.1' (Get-NextBaseVersion '3.1.0' '0.0.1')
    Test-Case 'two patches add two' '3.1.7' (Get-NextBaseVersion '3.1.5' '0.0.2')
    Test-Case 'a minor resets the patch' '3.2.0' (Get-NextBaseVersion '3.1.5' '0.1.0')
    Test-Case 'a minor bump keeps its own patch' '3.2.1' (Get-NextBaseVersion '3.1.5' '0.1.1')
    Test-Case 'a major resets minor and patch' '4.0.0' (Get-NextBaseVersion '3.1.5' '1.0.0')
    Test-Case 'a major bump keeps its own minor and patch' '4.2.3' (Get-NextBaseVersion '3.1.5' '1.2.3')
    Test-Case 'the fourth part is the commit count plus the stamp commit' '3.1.1.325' (Get-FullVersion '3.1.1' 324)
    $threw = $false
    try { $null = Get-FullVersion '3.1.1' 65535 } catch { $threw = $true }
    Test-Case 'a fourth part over 65535 is refused' $true $threw

    Write-Host ''
    Write-Host '== the changelog (release.yml''s "Extract changelog" and "Stamp changelog") =='
    $log = "# Changelog`n`n## Unreleased`n`n- Fix one thing`n- Add another`n`n### Notes`n`ndetail`n`n## v3.1.0.325 - 2026-08-15`n`n- Older`n"
    Test-Case 'the notes are the Unreleased section, sub-headings included' "- Fix one thing`n- Add another`n`n### Notes`n`ndetail" (Get-UnreleasedNotes $log)
    Test-Case 'an empty Unreleased section has no notes' '' (Get-UnreleasedNotes "# Changelog`n`n## Unreleased`n`n## v3.1.0.325 - 2026-08-15`n`n- Older`n")
    Test-Case 'no Unreleased section has no notes' '' (Get-UnreleasedNotes "# Changelog`n`n## v1 - x`n")
    Test-Case 'a last section runs to the end of the file' '- Only' (Get-UnreleasedNotes "## Unreleased`n- Only`n")
    Test-Case 'GitHub''s release body limit' 125000 $ReleaseBodyLimit
    $v = Get-NotesLengthVerdict -Notes ('x' * 125000) -Limit 125000 -Publishing $true
    Test-Case 'notes at the limit may publish' $true ($null -eq $v.Problem -and $null -eq $v.Note)
    $v = Get-NotesLengthVerdict -Notes ('x' * 125001) -Limit 125000 -Publishing $true
    Test-Case 'one character over refuses to publish' $true ([string]$v.Problem).Contains('125001 characters')
    $v = Get-NotesLengthVerdict -Notes ('x' * 125001) -Limit 125000 -Publishing $false
    Test-Case 'and a dry run of it goes on, noted' $true ($null -eq $v.Problem -and ([string]$v.Note).Contains('dry run goes on'))
    $stamped = Set-ChangelogStamp $log '3.1.1.325' '2026-10-03'
    Test-Case 'the stamp leaves Unreleased empty and heads the entries with the version' $true ($stamped.StartsWith("# Changelog`n`n## Unreleased`n`n## v3.1.1.325 - 2026-10-03`n`n- Fix one thing`n"))
    Test-Case 'and changes nothing else' $log ($stamped.Replace("`n`n## v3.1.1.325 - 2026-10-03", ''))
    Test-Case 'after it, the Unreleased section is empty' '' (Get-UnreleasedNotes $stamped)
    $crlf = $log.Replace("`n", "`r`n")
    Test-Case 'a CRLF file is stamped in CRLF' $true ((Set-ChangelogStamp $crlf '1.2.3.4' '2026-10-03').Contains("## Unreleased`r`n`r`n## v1.2.3.4 - 2026-10-03`r`n"))
    Test-Case 'only the first Unreleased heading is stamped' 1 ([regex]::Matches((Set-ChangelogStamp ($log + "## Unreleased`n") '1.2.3.4' 'd'), '## v1\.2\.3\.4').Count)
    Test-Case 'a heading that merely contains it is not it' $true ($null -eq (Set-ChangelogStamp "### Unreleased`n- x`n" '1.2.3.4' 'd'))
    Test-Case 'the stamp diff that is right' $true ($null -eq (Test-StampDiff @("2`t0`tCHANGELOG.md")))
    Test-Case 'a stamp that touches another file is refused' $true ([string](Test-StampDiff @("2`t0`tCHANGELOG.md", "1`t1`tProperties/AssemblyInfo.cs"))).Contains('2 file(s)')
    Test-Case 'a stamp that removes a line is refused' $true ([string](Test-StampDiff @("2`t1`tCHANGELOG.md"))).Contains('none removed')
    Test-Case 'no change at all is refused' $true ([string](Test-StampDiff @())).Contains('0 file(s)')

    Write-Host ''
    Write-Host '== the changelog''s rules and the release body (Q123, BrowserAI''s system) =='
    # Code points, not glyphs: a culture comparison ignores the variation selector three icons carry.
    function Get-CodePointText([string] $Text) {
        $points = @()
        for ($n = 0; $n -lt $Text.Length; $n++) {
            $point = [char]::ConvertToUtf32($Text, $n)
            if ($point -gt 0xFFFF) { $n++ }
            $points += ('{0:X}' -f $point)
        }
        return ($points -join ' ')
    }
    $palette = @($ChangelogPalette)
    Test-Case 'the palette is twelve icons' 12 $palette.Count
    Test-Case 'exactly these code points, in the approved order' '2728|1F41B|1F527|1F512|1F5D1 FE0F|1F4A5|1F4DD|2705|1F4E6|26A1|267B FE0F|2B06 FE0F' (@($palette | ForEach-Object { Get-CodePointText $_.Icon }) -join '|')
    Test-Case 'each with its meaning, the first and the last' 'new capability|dependency move' ($palette[0].Means + '|' + $palette[11].Means)
    $icon = @{}
    foreach ($p in $palette) { $icon[$p.Means] = $p.Icon }
    $newIcon = $icon['new capability']
    $fixIcon = $icon['fix']
    $testIcon = $icon['tests and the gate']
    $legendLines = @('| Icon | Meaning | Icon | Meaning |', '|---|---|---|---|')
    for ($k = 0; $k -lt $palette.Count; $k += 2) { $legendLines += ('| ' + $palette[$k].Icon + ' | ' + $palette[$k].Means + ' | ' + $palette[$k + 1].Icon + ' | ' + $palette[$k + 1].Means + ' |') }
    $legendText = $legendLines -join "`n"
    $fixture = @('# Changelog', '', 'Prose above the legend that mentions a | pipe.', '', $legendText, '', '## Unreleased', '',
        'This release does two things.', 'It wraps.', '', 'A second paragraph.', '',
        '### Added', '', "- $newIcon **A new thing.**", '  The detail of it,', '  over two lines.', '',
        "- $testIcon **A test of it.** Detail on the headline's own line.", '',
        '### Fixed', '', "- $fixIcon **A fixed thing.**", '',
        '## v1.0.0.1 - 2026-01-01', '', '- Older, in the old shape.', '') -join "`n"
    $clean = Test-UnreleasedSection $fixture
    Test-Case 'a changelog in the shape breaks no rule' '' (@($clean.Problems) -join ' | ')
    Test-Case 'its entries are counted' 3 $clean.Entries
    Test-Case 'and its preamble is the prose above the first group' "This release does two things.`nIt wraps.`n`nA second paragraph." $clean.Preamble
    $broken = @(
        @('an entry with no icon', "- $fixIcon **A fixed thing.**", '- **A fixed thing.**', 'not in the shape'),
        @('an icon the palette does not have', "- $fixIcon **A fixed thing.**", '- 🎉 **A fixed thing.**', 'twelve icons'),
        @('a headline nobody made bold', "- $fixIcon **A fixed thing.**", "- $fixIcon A fixed thing.", 'not in the shape'),
        @('a headline of two sentences', "- $fixIcon **A fixed thing.**", "- $fixIcon **Two sentences. That is one too many.**", 'more than one sentence'),
        @('a headline with no full stop', "- $fixIcon **A fixed thing.**", "- $fixIcon **A fixed thing**", 'full stop'),
        @('a headline one character over budget', "- $fixIcon **A fixed thing.**", ("- $fixIcon **" + ('x' * $HeadlineBudget) + '.**'), 'over the 100'),
        @('a detail that opens with its headline', "- $fixIcon **A fixed thing.**", "- $fixIcon **The packer's own asset list no longer disagrees.** The packer's own asset list no longer disagrees, because it is written once.", 'the headline said twice'),
        @('a group Keep a Changelog does not have', '### Fixed', '### Improved', 'not a Keep a Changelog group'),
        @('a group out of its order', '### Added', '### Security', 'out of Keep a Changelog''s order'),
        @('the same group twice', "- $fixIcon **A fixed thing.**", "- $fixIcon **A fixed thing.**`n`n### Fixed`n`n- $fixIcon **Fixed again.**", 'appears twice'),
        @('an entry above the first group', 'A second paragraph.', "A second paragraph.`n`n- $newIcon **Above every group.**", 'above the first group heading'),
        @('a paragraph under a group', "  over two lines.`n", "  over two lines.`n`nA paragraph nobody indented.`n", 'a paragraph under'),
        @('an em dash in a headline', '**A fixed thing.**', ('**A fixed ' + [char]0x2014 + ' thing.**'), 'an em dash'),
        @('a curly quote in the preamble', 'It wraps.', ('It' + [char]0x2019 + 's wrapped.'), 'curly closing quote'),
        @('a legend on one line', $legendText, (@($palette | ForEach-Object { $_.Icon + ' ' + $_.Means }) -join ' - '), 'no icon legend'),
        @('a legend missing an icon', ('| ' + $palette[10].Icon + ' | ' + $palette[10].Means + ' | ' + $palette[11].Icon + ' | ' + $palette[11].Means + ' |'), ('| ' + $palette[10].Icon + ' | ' + $palette[10].Means + ' | | |'), 'not the approved palette')
    )
    foreach ($b in $broken) {
        $found = @((Test-UnreleasedSection ($fixture.Replace($b[1], $b[2]))).Problems) -join ' | '
        Test-Case "refused: $($b[0])" $true ($found.Contains($b[3]))
    }
    Test-Case 'a headline of exactly 100 characters is allowed' '' (@((Test-UnreleasedSection ($fixture.Replace('**A fixed thing.**', ('**' + ('x' * ($HeadlineBudget - 1)) + '.**')))).Problems) -join ' | ')
    Test-Case 'a short subject shared with the headline is ordinary English' 3 (Get-SharedOpeningWords 'The upload set is data.' ' The upload set was a judgement until today.')
    Test-Case 'a dash in a detail is not refused: no detail reaches the release page' '' (@((Test-UnreleasedSection ($fixture.Replace('  over two lines.', ('  over two lines ' + [char]0x2014 + ' and on.')))).Problems) -join ' | ')
    Test-Case 'nor one inside a code span in a headline' '' (@((Test-UnreleasedSection ($fixture.Replace('**A fixed thing.**', ('**A fixed `a' + [char]0x2014 + 'b` thing.**')))).Problems) -join ' | ')
    $emptied = "# Changelog`n`n$legendText`n`n## Unreleased`n`nA preamble alone.`n`n### Added`n`n## v1.0.0.1 - 2026-01-01`n`n- Older.`n"
    Test-Case 'a preamble and a group heading with no entry under it is no entries' 0 (Test-UnreleasedSection $emptied).Entries
    $v = Get-PreambleVerdict -Preamble '' -Publishing $true
    Test-Case 'no preamble refuses to publish' $true ([string]$v.Problem).Contains('no preamble')
    $v = Get-PreambleVerdict -Preamble '' -Publishing $false
    Test-Case 'and a dry run of it goes on, noted' $true ($null -eq $v.Problem -and ([string]$v.Note).Contains('dry run goes on'))
    $v = Get-PreambleVerdict -Preamble 'What this release is.' -Publishing $true
    Test-Case 'a preamble may publish' $true ($null -eq $v.Problem -and $null -eq $v.Note)
    $stampedFixture = Set-ChangelogStamp $fixture '1.2.3.4' '2026-10-03'
    $made = New-ReleaseBody -StampedText $stampedFixture -Version '1.2.3.4' -RepositoryUrl 'https://github.com/o/r' -Limit 125000
    $file = 'https://github.com/o/r/blob/v1.2.3.4/CHANGELOG.md'
    $expectedBody = (@('This release does two things. It wraps.', '', 'A second paragraph.', '', '### Added', '',
            ('- ' + $newIcon + ' **A new thing.** [read more](' + $file + '?plain=1#L25-L27)'),
            ('- ' + $testIcon + ' **A test of it.** [read more](' + $file + '?plain=1#L29-L29)'),
            '', '### Fixed', '',
            ('- ' + $fixIcon + ' **A fixed thing.** [read more](' + $file + '?plain=1#L33-L33)'),
            '', '---', '', $legendText, '',
            ('The full changelog for this release: [CHANGELOG.md](' + $file + '#v1234---2026-10-03)')) -join "`n") + "`n"
    Test-Case 'the body: preamble unwrapped, each headline with its own lines, the legend, the section' $expectedBody $made.Body
    Test-Case 'the same, compared ordinally' $true ([string]::Equals($expectedBody, $made.Body, [StringComparison]::Ordinal))
    Test-Case 'linked, three entries, two groups' 'linked 3 2' "$($made.Shape) $($made.Entries) $($made.Groups)"
    $stampedLines = (ConvertTo-LfText $stampedFixture) -split "`n"
    Test-Case 'a range starts on its entry''s first line' ('- ' + $newIcon + ' **A new thing.**') $stampedLines[24]
    Test-Case 'and ends on its last line with text, not the blank after it' '  over two lines.' $stampedLines[26]
    $tight = New-ReleaseBody -StampedText $stampedFixture -Version '1.2.3.4' -RepositoryUrl 'https://github.com/o/r' -Limit ($made.Body.Length - 1)
    Test-Case 'a body over the limit drops every link and keeps the headlines' 'headlines False 3' "$($tight.Shape) $($tight.Body.Contains('[read more]')) $($tight.Entries)"
    Test-Case 'and says how long the linked one would have been' $made.Body.Length $tight.LinkedLength
    $none = New-ReleaseBody -StampedText $fixture -Version '9.9.9.9' -RepositoryUrl 'https://github.com/o/r' -Limit 125000
    Test-Case 'a version with no section is refused, not given an empty body' $true ((@($none.Problems) -join ' ').Contains("no '## v9.9.9.9 - ' section") -and $none.Body -eq '')
    Test-Case 'the anchor GitHub gives a release heading' 'v310325---2026-08-15' (Get-GitHubAnchor 'v3.1.0.325 - 2026-08-15')
    Test-Case 'and one with a code span, a link and emphasis' 'a-x-link-b' (Get-GitHubAnchor 'A `<x>` [link](u) *b*')
    Test-Case 'released sections as the tag has them pass' $true ($null -eq (Test-ReleasedSectionsUnchanged -ChangelogText $fixture -TaggedText $fixture -Tag 'v1.0.0.1'))
    Test-Case 'a new entry under Unreleased is no released section' $true ($null -eq (Test-ReleasedSectionsUnchanged -ChangelogText ($fixture.Replace('### Fixed', "### Fixed`n`n- $fixIcon **Another.**")) -TaggedText $fixture -Tag 'v1.0.0.1'))
    $rewritten = Test-ReleasedSectionsUnchanged -ChangelogText ($fixture.Replace('- Older, in the old shape.', '- Older, quietly reworded.')) -TaggedText $fixture -Tag 'v1.0.0.1'
    Test-Case 'a released section rewritten is refused, by its heading' $true ([string]$rewritten).Contains("from '## v1.0.0.1 - 2026-01-01' on")
    Test-Case 'and CRLF against LF is no rewrite' $true ($null -eq (Test-ReleasedSectionsUnchanged -ChangelogText ($fixture.Replace("`n", "`r`n")) -TaggedText $fixture -Tag 'v1.0.0.1'))

    Write-Host ''
    Write-Host '== publishing =='
    Test-Case 'the repository from an https remote' 'SixFive7/OutlookAI' (Get-GitHubRepository 'https://github.com/SixFive7/OutlookAI.git')
    Test-Case 'from an ssh remote' 'SixFive7/OutlookAI' (Get-GitHubRepository 'git@github.com:SixFive7/OutlookAI.git')
    Test-Case 'without .git' 'SixFive7/OutlookAI' (Get-GitHubRepository 'https://github.com/SixFive7/OutlookAI')
    Test-Case 'not GitHub is nothing' $true ($null -eq (Get-GitHubRepository 'https://example.com/a/b.git'))
    $v = Get-HeadVerdict -Head 'aaa' -RemoteTip 'aaa' -Publishing $true
    Test-Case 'HEAD on origin/master may publish' $true ($null -eq $v.Problem -and $null -eq $v.Note)
    $v = Get-HeadVerdict -Head 'aaa' -RemoteTip 'bbb' -Publishing $true
    Test-Case 'HEAD off origin/master may not publish' $true ([string]$v.Problem).Contains('fast-forward')
    $v = Get-HeadVerdict -Head 'aaa' -RemoteTip 'bbb' -Publishing $false
    Test-Case 'but a dry run of it goes on, noted' $true ($null -eq $v.Problem -and ([string]$v.Note).Contains('dry run goes on'))
    $now = [datetime]'2026-10-03'
    Test-Case 'a certificate valid for years passes quietly' $true ($null -eq (Get-CertificateVerdict ([datetime]'2031-03-24') $now).Problem -and $null -eq (Get-CertificateVerdict ([datetime]'2031-03-24') $now).Warning)
    Test-Case 'one expiring within 30 days warns' $true ([string](Get-CertificateVerdict ([datetime]'2026-10-20') $now).Warning).Contains('expires in 17 day')
    Test-Case 'an expired one refuses' $true ([string](Get-CertificateVerdict ([datetime]'2026-10-01') $now).Problem).Contains('EXPIRED')
    Test-Case 'an installer at the cap passes' $true ($null -eq (Test-InstallerSize (50MB) 50))
    Test-Case 'one byte over it refuses' $true ([string](Test-InstallerSize (50MB + 1) 50)).Contains('over the auto-updater')
    Test-Case 'CERT_E_UNTRUSTEDROOT is 0x800B0109' '800B0109' ('{0:X8}' -f $CertEUntrustedRoot)
    $pin = '2578F7B869383572E751DD6B61B5374C55C6E995'
    Test-Case 'a trusted signature by the pinned key is accepted' $true ($null -eq (Get-SignatureVerdict 0 $pin $pin $true))
    Test-Case 'so is the self-signed key''s untrusted root - as the updater accepts it' $true ($null -eq (Get-SignatureVerdict $CertEUntrustedRoot $pin.ToLowerInvariant() $pin $true))
    Test-Case 'a bad digest is refused' $true ([string](Get-SignatureVerdict ([int]-2146869232) $pin $pin $true)).Contains('0x80096010')
    Test-Case 'no signature is refused' $true ([string](Get-SignatureVerdict ([int]-2146762496) $pin $pin $true)).Contains('0x800B0100')
    Test-Case 'another signer is refused' $true ([string](Get-SignatureVerdict 0 ('A' * 40) $pin $true)).Contains('not the pinned')
    Test-Case 'no timestamp is refused' $true ([string](Get-SignatureVerdict $CertEUntrustedRoot $pin $pin $false)).Contains('no timestamp')
    Add-Type -TypeDefinition $InstallerTrustSource
    $selfTrust = [OutlookAIRelease.InstallerTrust]::Verify($PSCommandPath)
    Test-Case 'the WinVerifyTrust interop compiles and refuses this unsigned script' $true ($null -ne (Get-SignatureVerdict $selfTrust $pin $pin $true))
    $sign = Get-SignToolArgumentList -Thumbprint 'ABCDEF' -TimestampUrl 'http://ts' -File 'C:\r\.work\release\v1\OutlookAI-v1.exe'
    Test-Case 'signtool signs by thumbprint, SHA-256, RFC 3161' 'sign /sha1 ABCDEF /fd SHA256 /tr http://ts /td SHA256 C:\r\.work\release\v1\OutlookAI-v1.exe' ($sign -join ' ')
    Test-Case 'and never with /a' $false ($sign -contains '/a')
    Test-Case 'a file with a space is quoted' '"C:\a b\x.exe"' (Get-SignToolArgumentList 'A' 'http://ts' 'C:\a b\x.exe')[-1]
    $gh = Get-GhReleaseArgumentList -Repository 'SixFive7/OutlookAI' -Tag 'v3.1.1.325' -Asset 'C:\r\OutlookAI-v3.1.1.325.exe' -Target ('a' * 40) -NotesFile 'C:\r\notes.md'
    Test-Case 'gh release create, the tag, the asset' 'release create v3.1.1.325 C:\r\OutlookAI-v3.1.1.325.exe' (($gh[0..3]) -join ' ')
    Test-Case 'targeting the stamp commit, in this repository' $true ((($gh -join ' ').Contains("--repo SixFive7/OutlookAI --target $('a' * 40)")))
    Test-Case 'titled as every release has been' 'OutlookAI v3.1.1.325' $gh[([array]::IndexOf($gh, '--title') + 1)]
    Test-Case 'the pinned thumbprint is read from the project' '2578F7B869383572E751DD6B61B5374C55C6E995' (Get-PinnedThumbprint '<PropertyGroup><ManifestCertificateThumbprint>2578f7b869383572e751dd6b61b5374c55c6e995</ManifestCertificateThumbprint></PropertyGroup>')
    Test-Case 'a project without one gives nothing' $true ($null -eq (Get-PinnedThumbprint '<Project />'))
    Test-Case 'the build VM run directory is read from its log' 'C:\r\.work\build-vm-runs\20261003-1' (Get-TestRunDirectory @('[12:00:00] == Invoke-TestsOnBuildVm: run x ==', '[12:00:01]   results   C:\r\.work\build-vm-runs\20261003-1'))
    Test-Case 'a log without one gives nothing' $true ($null -eq (Get-TestRunDirectory @('nothing')))

    Write-Host ''
    Write-Host '== the wiring: what the release runs exists, read and never run =='
    $projectPath = Join-Path $RepoRoot 'OutlookAI.csproj'
    Test-Case 'OutlookAI.csproj pins a thumbprint this script can read' $true ($null -ne (Get-PinnedThumbprint ([System.IO.File]::ReadAllText($projectPath))))
    $guards = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot $GuardDirectory) -Filter $GuardPattern -File)
    Test-Case "the guards: $ExpectedGuardCount under $GuardDirectory" $ExpectedGuardCount $guards.Count
    Test-Case 'check-pinned-constants.ps1 is one of them' $true (@($guards | Where-Object { $_.Name -eq $PinnedConstantsGuard }).Count -eq 1)
    Test-Case 'and it takes the certificate the release signs with' 'String' (Get-ScriptParameterType -Path (Join-Path $RepoRoot "$GuardDirectory\$PinnedConstantsGuard") -Name 'ExpectedSigningThumbprint')
    foreach ($rel in $D7Comparisons) {
        Test-Case "D7 (c): $($rel.Replace('\', '/')) declares the switch -$D7Switch" 'SwitchParameter' (Get-ScriptParameterType -Path (Join-Path $RepoRoot $rel) -Name $D7Switch)
    }
    Test-Case 'the build script takes the release certificate' 'String' (Get-ScriptParameterType -Path (Join-Path $RepoRoot $BuildScript) -Name 'ReleaseSigningThumbprint')
    Test-Case 'the build VM runner takes a revision' 'String' (Get-ScriptParameterType -Path (Join-Path $RepoRoot $TestRunner) -Name 'Ref')
    Test-Case 'a parameter nobody declares is nothing' $true ($null -eq (Get-ScriptParameterType -Path (Join-Path $RepoRoot $BuildScript) -Name 'NoSuchParameter'))
    Write-Host '  SKIP D7 (c) itself: it needs Visual Studio, so it runs in a release, on the workstation (AGENTS.md).'

    Write-Host ''
    Write-Host '== the wiring: this repository''s own CHANGELOG.md, held to the same rules =='
    $realText = [System.IO.File]::ReadAllText((Join-Path $RepoRoot 'CHANGELOG.md'))
    Test-Case 'its legend is the approved palette, as a table' '' (@(Get-LegendProblems $realText) -join ' | ')
    $real = Test-UnreleasedSection $realText
    Test-Case 'every Unreleased entry keeps the rules (-CheckChangelog lists any that do not)' '' (@($real.Problems) -join ' | ')
    if ($real.Entries -gt 0) {
        $realStamped = Set-ChangelogStamp $realText '0.0.0.0' '2000-01-01'
        $realBody = New-ReleaseBody -StampedText $realStamped -Version '0.0.0.0' -RepositoryUrl 'https://github.com/SixFive7/OutlookAI' -Limit $ReleaseBodyLimit
        Test-Case "a body made from its $($real.Entries) entries is linked, within GitHub's limit" 'linked' $realBody.Shape
        $realLines = (ConvertTo-LfText $realStamped) -split "`n"
        $ranges = @([regex]::Matches($realBody.Body, '\?plain=1#L(\d+)-L(\d+)\)'))
        $astray = @($ranges | Where-Object { -not $realLines[[int]$_.Groups[1].Value - 1].StartsWith('- ') -or $realLines[[int]$_.Groups[2].Value - 1].Trim().Length -eq 0 })
        Test-Case 'one range per entry, each from its entry''s first line to its last with text' "$($real.Entries) 0" "$($ranges.Count) $($astray.Count)"
    }
    else { Write-Host '  SKIP no entries under ## Unreleased, as right after a release - a release refuses that itself' }

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a run on the workstation can settle these - its dry run does, all but the last:'
    Write-Host '  * that gh, git, signtool, the certificate and the build VM answer as assumed'
    Write-Host '  * that D7 (c) passes against the Visual Studio installed there'
    Write-Host '  * that GitHub renders release-notes.md as written, and highlights each read-more range on github.com'
    Write-Host '  * that the push and gh release create succeed - which only -Execute does'
    if ($script:Failures.Count -gt 0) {
        Write-Host ''
        foreach ($f in $script:Failures) { Write-Host "  $f" }
        return 1
    }
    return 0
}

# =============================================================================================
# I/O
# =============================================================================================
function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-Git {
    param([Parameter(Mandatory = $true)] [string[]] $Arguments, [switch] $AllowFailure)
    $out = @(Invoke-NativeCommand { & git -C $RepoRoot @Arguments 2>&1 })
    $code = $LASTEXITCODE
    if ($code -ne 0 -and -not $AllowFailure) { throw "git $($Arguments -join ' ') failed (exit $code): $(($out | Out-String).Trim())" }
    return [pscustomobject]@{ ExitCode = $code; Lines = $out; Text = (($out | Out-String).Trim()) }
}

function Invoke-Gh {
    param([Parameter(Mandatory = $true)] [string[]] $Arguments)
    $out = @(Invoke-NativeCommand { & gh @Arguments 2>&1 })
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Lines = $out; Text = (($out | Out-String).Trim()) }
}

# One file as a revision holds it, read as the UTF-8 it is - or $null when git cannot produce it. Not
# through Invoke-Git: a PowerShell pipe decodes a program's output in the console's code page, and
# CHANGELOG.md carries emoji and dashes that would arrive garbled and compare unequal to themselves.
# Invoke-Logged hands git a FILE for its output, so the bytes land as git wrote them.
function Get-GitFileText {
    param([string] $Revision, [string] $Path, [string] $ScratchDirectory)
    New-Item -ItemType Directory -Force -Path $ScratchDirectory | Out-Null
    $stem = Join-Path $ScratchDirectory ('git-show-' + ($Revision -replace '[^A-Za-z0-9.]', '_'))
    $r = Invoke-Logged -FilePath 'git' -ArgumentList @('-C', (Format-Argument $RepoRoot), 'show', ($Revision + ':' + $Path)) -LogStem $stem -TimeoutMinutes 2
    if ($r.TimedOut -or $r.ExitCode -ne 0) { return $null }
    return [System.IO.File]::ReadAllText($r.Out, (New-Object System.Text.UTF8Encoding($false))).TrimStart([char]0xFEFF)
}

function Resolve-ShellExe([string] $Edition) {
    if ($Edition -eq 'Desktop') {
        $p = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (Test-Path -LiteralPath $p) { return $p }
        throw 'REFUSING: Windows PowerShell 5.1 (powershell.exe) was not found, and the guards must pass under it.'
    }
    if ($PSVersionTable.PSEdition -eq 'Core') { return (Join-Path $PSHOME 'pwsh.exe') }
    $onPath = @(Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue)
    if ($onPath.Count -gt 0) { return $onPath[0].Source }
    throw 'REFUSING: PowerShell 7 (pwsh.exe) was not found, and the guards must pass under it as well as under 5.1.'
}

function Resolve-SignTool {
    $root = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    $found = @(Get-ChildItem -LiteralPath $root -Recurse -Filter 'signtool.exe' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match '\\x64\\' } | Sort-Object FullName -Descending)
    if ($found.Count -eq 0) { throw "REFUSING: signtool.exe (x64) was not found under $root - it comes with the Windows SDK, which Visual Studio installs." }
    return $found[0].FullName
}

# A child PowerShell running one script with -File, logged, judged by its exit code.
function Invoke-ScriptStep {
    param([string] $ShellExe, [string] $ScriptPath, [string[]] $ScriptArguments, [string] $LogStem, [int] $TimeoutMinutes = 20)
    $argList = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', (Format-Argument $ScriptPath)) + @($ScriptArguments)
    return Invoke-Logged -FilePath $ShellExe -ArgumentList $argList -LogStem $LogStem -TimeoutMinutes $TimeoutMinutes -WorkingDirectory $RepoRoot
}

# -CheckChangelog: step 3's checks on the working tree's CHANGELOG.md, its released sections against the
# newest release tag HEAD descends from, and the body the next release would get - written out, never
# published. Version 0.0.0.0 stands in for the one a release derives: it changes the links' tag and
# nothing else, and no line range depends on it.
function Invoke-ChangelogCheck {
    $path = Join-Path $RepoRoot 'CHANGELOG.md'
    $text = [System.IO.File]::ReadAllText($path)
    $check = Test-UnreleasedSection $text
    $problems = @($check.Problems)
    Write-Host "Tools/Publish-Release.ps1 -CheckChangelog: $path"
    Write-Host "  $($check.Entries) entr(y/ies) under ## Unreleased"
    if ($check.Entries -eq 0) { Write-Host '  NOTE no entries: a release refuses until there are some. Right after a release that is expected.' }
    if ($check.Entries -gt 0 -and -not $check.Preamble) { Write-Host '  NOTE no preamble yet: a release with -Execute refuses until the section opens with one.' }
    $tag = Invoke-Git -Arguments @('describe', '--tags', '--abbrev=0', '--match', 'v*', 'HEAD') -AllowFailure
    if ($tag.ExitCode -eq 0 -and $tag.Text) {
        $tagged = Get-GitFileText -Revision $tag.Text -Path 'CHANGELOG.md' -ScratchDirectory (Join-Path $WorkRoot $ChangelogCheckDirectory)
        if ($null -eq $tagged) { $problems += "git could not read CHANGELOG.md at $($tag.Text)." }
        else {
            $moved = Test-ReleasedSectionsUnchanged -ChangelogText $text -TaggedText $tagged -Tag $tag.Text
            if ($moved) { $problems += $moved } else { Write-Host "  the released sections are exactly what $($tag.Text) carries" }
        }
    }
    else { Write-Host '  NOTE no release tag is reachable from HEAD, so the released sections were not compared with one.' }
    if ($check.Entries -gt 0 -and $problems.Count -eq 0) {
        $repository = $null
        $remoteUrl = Invoke-Git -Arguments @('remote', 'get-url', $Remote) -AllowFailure
        if ($remoteUrl.ExitCode -eq 0) { $repository = Get-GitHubRepository $remoteUrl.Text }
        if (-not $repository) { $repository = 'SixFive7/OutlookAI' }
        $version = '0.0.0.0'
        $stamped = Set-ChangelogStamp -ChangelogText $text -Version $version -Date (Get-Date -Format 'yyyy-MM-dd')
        $made = New-ReleaseBody -StampedText $stamped -Version $version -RepositoryUrl "https://github.com/$repository" -Limit $ReleaseBodyLimit
        $problems += @($made.Problems)
        if ($made.Body) {
            $problems += @(Get-UntypedCharacterProblems $made.Body 'The release body')
            $lengthVerdict = Get-NotesLengthVerdict -Notes $made.Body -Limit $ReleaseBodyLimit -Publishing $true
            if ($lengthVerdict.Problem) { $problems += $lengthVerdict.Problem }
            $directory = Join-Path $WorkRoot $ChangelogCheckDirectory
            New-Item -ItemType Directory -Force -Path $directory | Out-Null
            $bodyFile = Join-Path $directory 'release-body.md'
            [System.IO.File]::WriteAllText($bodyFile, $made.Body, (New-Object System.Text.UTF8Encoding($false)))
            Write-Host "  release body $($made.Shape.ToUpperInvariant()): $($made.Body.Length) characters of the $ReleaseBodyLimit GitHub accepts, $($made.Entries) entries in $($made.Groups) group(s) -> $bodyFile"
            Write-Host "  (v$version stands in for the version a release derives; no line range depends on it)"
        }
    }
    if ($problems.Count -gt 0) {
        Write-Host ''
        Write-Host "$($problems.Count) problem(s). AGENTS.md (Changelog) says what each rule is:"
        foreach ($p in $problems) { Write-Host "  $p" }
        return 1
    }
    Write-Host 'The changelog keeps every rule.'
    return 0
}

# =============================================================================================
# MAIN
# =============================================================================================
if ($SelfTest) { exit (Invoke-SelfTest) }
if ($CheckChangelog) { exit (Invoke-ChangelogCheck) }

$bumpProblem = Test-VersionBumpText $VersionBump
if ($bumpProblem) { throw $bumpProblem }
if (Test-IsElevated) {
    throw 'REFUSING TO RUN ELEVATED. An elevated process can carry another account''s HKCU, and this build runs beside your own Outlook. Run it from an ordinary, unelevated shell.'
}
$RepoRoot = [System.IO.Path]::GetFullPath($RepoRoot)
$WorkRoot = [System.IO.Path]::GetFullPath($WorkRoot)
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$started = Get-Date
$record = [ordered]@{ schema = 1; producedBy = 'Tools/Publish-Release.ps1'; mode = ''; started = $started.ToString('o') }
if ($Execute) { $record.mode = 'publish' } else { $record.mode = 'dry run' }

$modeText = 'DRY RUN - the whole release except publishing'
if ($Execute) { $modeText = 'PUBLISH' }
Say "== Tools/Publish-Release.ps1 - $modeText =="
Say "  repository $RepoRoot"

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 1. Preconditions =='
$dirty = Invoke-Git -Arguments @('status', '--porcelain')
if (@($dirty.Lines | Where-Object { $_ }).Count -gt 0) {
    throw "REFUSING: the working tree has uncommitted changes. The guards read the working tree and the build archives the commit, so they must be the same:`n$($dirty.Text)"
}
$head = (Invoke-Git -Arguments @('rev-parse', '--verify', 'HEAD^{commit}')).Text
$headSubject = (Invoke-Git -Arguments @('log', '-1', '--format=%s', $head)).Text
Say "  HEAD       $head - $headSubject"
$null = Invoke-Git -Arguments @('fetch', '--quiet', $Remote, $ReleaseBranch)
$remoteTip = (Invoke-Git -Arguments @('rev-parse', '--verify', "refs/remotes/$Remote/$ReleaseBranch^{commit}")).Text
Say "  $Remote/$ReleaseBranch $remoteTip"
$verdict = Get-HeadVerdict -Head $head -RemoteTip $remoteTip -Publishing ([bool]$Execute)
if ($verdict.Problem) { throw "REFUSING: $($verdict.Problem)" }
if ($verdict.Note) { Say "  NOTE $($verdict.Note)" }
$repository = Get-GitHubRepository (Invoke-Git -Arguments @('remote', 'get-url', $Remote)).Text
if (-not $repository) { throw "REFUSING: $Remote is not a GitHub repository, so there is nowhere to publish." }
$auth = Invoke-Gh -Arguments @('auth', 'status', '--hostname', 'github.com')
if ($auth.ExitCode -ne 0) { throw "REFUSING: gh is not logged in to github.com (gh auth status, exit $($auth.ExitCode)):`n$($auth.Text)" }
Say "  GitHub     $repository, gh logged in"
$record.head = $head
$record.remoteTip = $remoteTip
$record.repository = $repository

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 2. Version =='
$latest = Invoke-Gh -Arguments @('release', 'view', '--repo', $repository, '--json', 'tagName', '-q', '.tagName')
if ($latest.ExitCode -ne 0) {
    # The workflow turned any failure here into 0.0.0, which would have released 0.0.1.<n> on a
    # network blip. Only "there is no release yet" may mean that.
    if ($latest.Text -match 'release not found') { $latestTag = '' }
    else { throw "REFUSING: gh could not read the latest release (exit $($latest.ExitCode)): $($latest.Text)" }
}
else { $latestTag = $latest.Text }
$currentBase = Get-BaseVersionFromTag $latestTag
$base = Get-NextBaseVersion -Current $currentBase -Bump $VersionBump
$count = [int](Invoke-Git -Arguments @('rev-list', '--count', $head)).Text
$version = Get-FullVersion -Base $base -CommitCount $count
$tag = "v$version"
Say "  latest release $latestTag; $currentBase + $VersionBump = $base; HEAD is commit $count, the stamp commit $($count + 1)"
Say "  VERSION    $version"
$tagOnRemote = Invoke-Git -Arguments @('ls-remote', '--tags', $Remote, "refs/tags/$tag")
if ($tagOnRemote.Text) { throw "REFUSING: the tag $tag already exists on $Remote." }
$existing = Invoke-Gh -Arguments @('release', 'view', $tag, '--repo', $repository, '--json', 'tagName')
if ($existing.ExitCode -eq 0) { throw "REFUSING: a release $tag already exists on GitHub." }
$record.latestRelease = $latestTag
$record.versionBump = $VersionBump
$record.version = $version

$outDir = Join-Path $WorkRoot $tag
if (Test-Path -LiteralPath $outDir) { Remove-Item -LiteralPath $outDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$logDir = New-Item -ItemType Directory -Force -Path (Join-Path $outDir 'logs')
Say "  output     $outDir"

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 3. Release notes =='
$changelogPath = Join-Path $RepoRoot 'CHANGELOG.md'
$changelogBytes = [System.IO.File]::ReadAllBytes($changelogPath)
$changelogHasBom = ($changelogBytes.Length -ge 3 -and $changelogBytes[0] -eq 0xEF -and $changelogBytes[1] -eq 0xBB -and $changelogBytes[2] -eq 0xBF)
$changelogText = [System.IO.File]::ReadAllText($changelogPath)
$notes = Get-UnreleasedNotes $changelogText
# EMPTY MEANS NO ENTRIES, not no characters: a section holding nothing but its group headings, or only
# a preamble, is what a changelog nobody wrote looks like.
$check = Test-UnreleasedSection $changelogText
if ($check.Entries -eq 0) { throw "REFUSING: no entries under '## Unreleased' in CHANGELOG.md. Add release notes before creating a release." }
if (@($check.Problems).Count -gt 0) {
    throw ("REFUSING: CHANGELOG.md breaks the changelog's rules (AGENTS.md, Changelog) in $(@($check.Problems).Count) place(s); pwsh -File Tools/Publish-Release.ps1 -CheckChangelog lists them without a release:`n  " + (@($check.Problems) -join "`n  "))
}
Say "  $($check.Entries) entr(y/ies) under ## Unreleased, every one an icon from the legend, a one-sentence headline and its detail, in Keep a Changelog's groups and order"
$preambleVerdict = Get-PreambleVerdict -Preamble $check.Preamble -Publishing ([bool]$Execute)
if ($preambleVerdict.Problem) { throw "REFUSING: $($preambleVerdict.Problem)" }
if ($preambleVerdict.Note) { Say "  NOTE $($preambleVerdict.Note)" }
if ($latestTag) {
    $null = Invoke-Git -Arguments @('fetch', '--quiet', '--no-tags', $Remote, "+refs/tags/${latestTag}:refs/tags/${latestTag}")
    $taggedChangelog = Get-GitFileText -Revision "refs/tags/$latestTag" -Path 'CHANGELOG.md' -ScratchDirectory $logDir
    if ($null -eq $taggedChangelog) { throw "REFUSING: git could not read CHANGELOG.md at $latestTag, the latest release, so nothing shows its released sections are unchanged." }
    $moved = Test-ReleasedSectionsUnchanged -ChangelogText $changelogText -TaggedText $taggedChangelog -Tag $latestTag
    if ($moved) { throw "REFUSING: $moved" }
    Say "  the released sections are exactly what $latestTag carries"
}
else { Say '  no release yet, so no released section to hold unchanged' }
# THE STAMP IS WORKED OUT HERE, AND STEP 10 COMMITS THIS VERY TEXT: every line range in the body is a
# line of the stamped file, and true of no other. The date is the day this run reads the notes - the
# day of the cut - even when the run ends after midnight.
$date = Get-Date -Format 'yyyy-MM-dd'
$stampedText = Set-ChangelogStamp -ChangelogText $changelogText -Version $version -Date $date
if ($null -eq $stampedText) { throw "CHANGELOG.md has no '## Unreleased' heading at the start of a line to stamp." }
$made = New-ReleaseBody -StampedText $stampedText -Version $version -RepositoryUrl "https://github.com/$repository" -Limit $ReleaseBodyLimit
if (@($made.Problems).Count -gt 0) { throw ("REFUSING: the release body cannot be made:`n  " + (@($made.Problems) -join "`n  ")) }
$untyped = @(Get-UntypedCharacterProblems $made.Body 'The release body')
if ($untyped.Count -gt 0) { throw ('REFUSING: ' + ($untyped -join ' ')) }
$notesFile = Join-Path $outDir 'release-notes.md'
[System.IO.File]::WriteAllText($notesFile, $made.Body, (New-Object System.Text.UTF8Encoding($false)))
if ($made.Shape -eq 'linked') {
    Say "  release body LINKED: $($made.Body.Length) characters of the $ReleaseBodyLimit GitHub accepts; $($made.Entries) entries in $($made.Groups) group(s), each its headline and its lines in CHANGELOG.md at $tag -> $notesFile"
}
else {
    Say "  NOTE the linked body would be $($made.LinkedLength) characters, over the $ReleaseBodyLimit GitHub accepts, so this one is HEADLINES ONLY: $($made.Body.Length) characters, $($made.Entries) entries, no per-entry links - the footer's link to the section is the only way into the detail -> $notesFile"
}
$notesVerdict = Get-NotesLengthVerdict -Notes $made.Body -Limit $ReleaseBodyLimit -Publishing ([bool]$Execute)
if ($notesVerdict.Problem) { throw "REFUSING: $($notesVerdict.Problem)" }
if ($notesVerdict.Note) { Say "  NOTE $($notesVerdict.Note)" }
$record.notesCharacters = $notes.Length
$record.releaseBody = [ordered]@{ shape = $made.Shape; characters = $made.Body.Length; linkedCharacters = $made.LinkedLength; limit = $ReleaseBodyLimit; entries = $made.Entries; groups = $made.Groups; preamble = [bool]$check.Preamble }

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 4. The signing certificate =='
$thumb = Get-PinnedThumbprint ([System.IO.File]::ReadAllText((Join-Path $RepoRoot 'OutlookAI.csproj')))
if (-not $thumb) { throw 'REFUSING: OutlookAI.csproj pins no ManifestCertificateThumbprint this script can read.' }
$cert = Get-Item -LiteralPath "Cert:\CurrentUser\My\$thumb" -ErrorAction SilentlyContinue
if (-not $cert) { throw "REFUSING: the release certificate $thumb, which OutlookAI.csproj and the updater pin, is not in Cert:\CurrentUser\My." }
if (-not $cert.HasPrivateKey) { throw "REFUSING: Cert:\CurrentUser\My\$thumb has no private key on this machine, so it cannot sign." }
$certVerdict = Get-CertificateVerdict -NotAfter $cert.NotAfter -Now (Get-Date)
if ($certVerdict.Problem) { throw "REFUSING: $($certVerdict.Problem)" }
if ($certVerdict.Warning) { Say "  WARNING $($certVerdict.Warning)" }
Say "  $thumb  $($cert.Subject), private key here, valid $($certVerdict.Days) more day(s), until $($cert.NotAfter.ToString('yyyy-MM-dd'))"
$record.signingThumbprint = $thumb

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 5. Guards, under PowerShell 7 and Windows PowerShell 5.1 =='
$pwshExe = Resolve-ShellExe 'Core'
$desktopExe = Resolve-ShellExe 'Desktop'
$guards = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot $GuardDirectory) -Filter $GuardPattern -File | Sort-Object Name)
if ($guards.Count -ne $ExpectedGuardCount) { throw "REFUSING: found $($guards.Count) guard(s) under $GuardDirectory, expected $ExpectedGuardCount - a guard that went missing proves nothing." }
$guardResults = @()
foreach ($shell in @(@{ Name = 'pwsh'; Exe = $pwshExe }, @{ Name = 'powershell'; Exe = $desktopExe })) {
    foreach ($g in $guards) {
        $r = Invoke-ScriptStep -ShellExe $shell.Exe -ScriptPath $g.FullName -ScriptArguments @() -LogStem (Join-Path $logDir "guard-$($g.BaseName)-$($shell.Name)")
        $guardResults += [ordered]@{ guard = $g.Name; shell = $shell.Name; exitCode = $r.ExitCode }
        if ($r.TimedOut -or $r.ExitCode -ne 0) {
            Show-LogTail $r.Out 25
            throw "REFUSING: $($g.Name) failed under $($shell.Name) (exit $($r.ExitCode)). Log: $($r.Out)"
        }
        Say "  OK  $($g.Name) under $($shell.Name)"
    }
}
$r = Invoke-ScriptStep -ShellExe $pwshExe -ScriptPath (Join-Path $RepoRoot "$GuardDirectory\$PinnedConstantsGuard") -ScriptArguments @('-ExpectedSigningThumbprint', $thumb) -LogStem (Join-Path $logDir 'guard-pinned-constants-certificate')
if ($r.TimedOut -or $r.ExitCode -ne 0) {
    Show-LogTail $r.Out 25
    throw "REFUSING: $PinnedConstantsGuard -ExpectedSigningThumbprint $thumb failed (exit $($r.ExitCode)): the certificate that would sign is not the one the updater pins."
}
Say "  OK  $PinnedConstantsGuard -ExpectedSigningThumbprint $thumb"
$record.guards = $guardResults

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 6. D7 (c): the build''s stand-ins against Visual Studio''s real VSTO targets =='
$d7Results = @()
foreach ($rel in $D7Comparisons) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($rel)
    $r = Invoke-ScriptStep -ShellExe $pwshExe -ScriptPath (Join-Path $RepoRoot $rel) -ScriptArguments @("-$D7Switch") -LogStem (Join-Path $logDir "d7-$name")
    $summary = @(Get-Content -LiteralPath $r.Out -ErrorAction SilentlyContinue | Where-Object { $_ -match 'assertion\(s\) across' })
    $d7Results += [ordered]@{ script = $rel.Replace('\', '/'); exitCode = $r.ExitCode; summary = ($summary -join ' ') }
    if ($r.TimedOut -or $r.ExitCode -ne 0) {
        Show-LogTail $r.Out 25
        throw "REFUSING TO RELEASE: $($rel.Replace('\', '/')) -$D7Switch failed (exit $($r.ExitCode)). The stand-ins that keep a build from registering the add-in in this Outlook no longer match Visual Studio's targets (D7 (c), Q81). Log: $($r.Out)"
    }
    Say "  OK  $($rel.Replace('\', '/')) -$D7Switch : $($summary -join ' ')"
}
$record.d7 = $d7Results

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 7. Build (Testbed/host/Publish-AddInPayload.ps1, the release build) =='
$vsto = $VstoRuntimePath
if (-not $vsto) {
    foreach ($c in @((Join-Path $RepoRoot '.work\media\vstor_redist.exe'), (Join-Path $RepoRoot 'Redist\vstor_redist.exe'))) {
        if (Test-Path -LiteralPath $c) { $vsto = $c; break }
    }
}
if (-not $vsto -or -not (Test-Path -LiteralPath $vsto)) {
    throw 'REFUSING: the VSTO runtime redistributable is not staged - not at .work\media\vstor_redist.exe, not at Redist\vstor_redist.exe. Testbed/MEDIA.md says where it comes from; pass -VstoRuntimePath if it is elsewhere.'
}
$buildDir = Join-Path $outDir 'build'
$buildArgs = @('-RepoRoot', (Format-Argument $RepoRoot), '-Ref', $head, '-Version', $version, '-OutDir', (Format-Argument $buildDir),
    '-ReleaseSigningThumbprint', $thumb, '-VstoRuntimePath', (Format-Argument $vsto))
Say "  $($BuildScript.Replace('\', '/')) $($buildArgs -join ' ')"
$r = Invoke-ScriptStep -ShellExe $pwshExe -ScriptPath (Join-Path $RepoRoot $BuildScript) -ScriptArguments $buildArgs -LogStem (Join-Path $logDir 'build') -TimeoutMinutes 60
if ($r.TimedOut -or $r.ExitCode -ne 0) {
    Show-LogTail $r.Out 40
    throw "REFUSING: the release build failed (exit $($r.ExitCode), timed out: $($r.TimedOut)). Log: $($r.Out)"
}
foreach ($l in @(Get-Content -LiteralPath $r.Out | Where-Object { $_ -match 'UNCHANGED|WARN|stamped|carries|OutlookAI\.(McpServer|ComHost)\.exe|payload|RELEASE BUILD DONE|sha256' })) { Write-Host "      | $l" }
$unsigned = Join-Path $buildDir "installer\OutlookAI-$tag.exe"
if (-not (Test-Path -LiteralPath $unsigned)) { throw "The build reported success but left no installer at $unsigned." }

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 8. Sign the installer, read the signature back, hold it to the cap =='
$installer = Join-Path $outDir "OutlookAI-$tag.exe"
Copy-Item -LiteralPath $unsigned -Destination $installer -Force
$signtool = Resolve-SignTool
Say "  signtool   $signtool"
$signed = $false
$usedServer = $null
foreach ($ts in $TimestampServers) {
    $r = Invoke-Logged -FilePath $signtool -ArgumentList (Get-SignToolArgumentList -Thumbprint $thumb -TimestampUrl $ts -File $installer) -LogStem (Join-Path $logDir 'signtool') -TimeoutMinutes 5
    if (-not $r.TimedOut -and $r.ExitCode -eq 0) { $signed = $true; $usedServer = $ts; break }
    Show-LogTail $r.Out 10
    Show-LogTail $r.Err 10
    Say "  timestamp server $ts failed (exit $($r.ExitCode)); trying the next"
    Start-Sleep -Seconds 5
}
if (-not $signed) { throw 'REFUSING: signing failed with every timestamp server.' }
Add-Type -TypeDefinition $InstallerTrustSource
$trust = [OutlookAIRelease.InstallerTrust]::Verify($installer)
$signer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 ([System.Security.Cryptography.X509Certificates.X509Certificate]::CreateFromSignedFile($installer))
$sig = Get-AuthenticodeSignature -LiteralPath $installer
$signatureProblem = Get-SignatureVerdict -TrustResult $trust -SignerThumbprint $signer.Thumbprint -PinnedThumbprint $thumb -Timestamped ($null -ne $sig.TimeStamperCertificate)
if ($signatureProblem) { throw "REFUSING: $signatureProblem" }
$installerItem = Get-Item -LiteralPath $installer
$sizeProblem = Test-InstallerSize -Bytes $installerItem.Length -CapMB $InstallerCapMB
if ($sizeProblem) { throw "REFUSING: $sizeProblem" }
$installerHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
Say "  $installer"
Say "  signed by $($signer.Subject) ($thumb), timestamped by $($sig.TimeStamperCertificate.Subject) via $usedServer"
Say ("  WinVerifyTrust 0x{0:X8}, the result the shipped updater accepts, and the signer it pins" -f $trust)
Say "  $([math]::Round($installerItem.Length / 1MB, 2)) MB of the $InstallerCapMB MB cap; sha256 $installerHash"
$record.installer = [ordered]@{ file = $installer; bytes = $installerItem.Length; sha256 = $installerHash; signer = $signer.Subject; winVerifyTrust = ('0x{0:X8}' -f $trust); timestampServer = $usedServer }

# ---------------------------------------------------------------------------------------------
Say ''
Say "== 9. Tests: the non-live suite and every self-test of $($head.Substring(0, 12)), on the build VM =="
$r = Invoke-ScriptStep -ShellExe $pwshExe -ScriptPath (Join-Path $RepoRoot $TestRunner) -ScriptArguments @($head) -LogStem (Join-Path $logDir 'build-vm') -TimeoutMinutes $TestTimeoutMinutes
$runDir = Get-TestRunDirectory @(Get-Content -LiteralPath $r.Out -ErrorAction SilentlyContinue)
$record.tests = [ordered]@{ exitCode = $r.ExitCode; timedOut = $r.TimedOut; runDirectory = $runDir }
if ($runDir -and (Test-Path -LiteralPath (Join-Path $runDir 'summary.txt'))) {
    foreach ($l in (Get-Content -LiteralPath (Join-Path $runDir 'summary.txt'))) { Write-Host "      | $l" }
}
else { Show-LogTail $r.Out 25 }
if ($r.TimedOut -or $r.ExitCode -ne 0) {
    throw "REFUSING TO RELEASE: the build VM's verdict on $head is exit $($r.ExitCode) (timed out: $($r.TimedOut)) - 1 a failure, 2 does not build, 3 not tested, 4 refused. Log: $($r.Out)"
}
Say '  exit 0: the whole non-live suite and every self-test passed'

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 10. The stamp commit - made, not pushed =='
# $stampedText and $date are step 3's: the release body was made from this very text.
$stampedPath = Join-Path $outDir 'CHANGELOG.stamped.md'
[System.IO.File]::WriteAllText($stampedPath, $stampedText, (New-Object System.Text.UTF8Encoding($changelogHasBom)))
$blob = (Invoke-Git -Arguments @('hash-object', '-w', '--path=CHANGELOG.md', $stampedPath)).Text
$entry = (Invoke-Git -Arguments @('ls-tree', $head, 'CHANGELOG.md')).Text
if ($entry -notmatch '^(?<mode>\d{6}) blob ') { throw "CHANGELOG.md is not a file in $head - ls-tree said '$entry'." }
$changelogMode = $Matches['mode']
$indexFile = Join-Path $outDir 'stamp.index'
$previousIndex = $env:GIT_INDEX_FILE
try {
    # A private index: the repository's own index, working tree and branches stay untouched.
    $env:GIT_INDEX_FILE = $indexFile
    $null = Invoke-Git -Arguments @('read-tree', $head)
    $null = Invoke-Git -Arguments @('update-index', '--cacheinfo', "$changelogMode,$blob,CHANGELOG.md")
    $tree = (Invoke-Git -Arguments @('write-tree')).Text
}
finally {
    $env:GIT_INDEX_FILE = $previousIndex
    if (Test-Path -LiteralPath $indexFile) { Remove-Item -LiteralPath $indexFile -Force }
}
$stampCommit = (Invoke-Git -Arguments @('commit-tree', $tree, '-p', $head, '-m', "Release $tag")).Text
$diffProblem = Test-StampDiff -NumstatLines @((Invoke-Git -Arguments @('diff', '--numstat', $head, $stampCommit)).Lines)
if ($diffProblem) { throw "REFUSING: $diffProblem" }
# Read back: the release body's line ranges were computed from $stampedText and are true of no other text.
$committedChangelog = Get-GitFileText -Revision $stampCommit -Path 'CHANGELOG.md' -ScratchDirectory $logDir
if ($null -eq $committedChangelog -or -not [string]::Equals((ConvertTo-LfText $committedChangelog), (ConvertTo-LfText $stampedText), [StringComparison]::Ordinal)) {
    throw "REFUSING: CHANGELOG.md in the stamp commit $stampCommit is not the text the release body's line ranges were computed from."
}
Say "  $stampCommit  Release $tag"
Say "  parent $head; CHANGELOG.md +2 lines (## v$version - $date), nothing else - read back, and the release body's line ranges are its lines"
$record.stampCommit = $stampCommit

$pushArgs = @('push', $Remote, "${stampCommit}:refs/heads/$ReleaseBranch")
$ghArgs = Get-GhReleaseArgumentList -Repository $repository -Tag $tag -Asset $installer -Target $stampCommit -NotesFile $notesFile
$recordPath = Join-Path $outDir 'release.json'

if (-not $Execute) {
    $record.finished = (Get-Date).ToString('o')
    $record.published = $false
    Set-Content -LiteralPath $recordPath -Value ($record | ConvertTo-Json -Depth 6) -Encoding UTF8
    Say ''
    Say "DRY RUN COMPLETE in $([int]((Get-Date) - $started).TotalMinutes) minute(s). Steps 1-10 ran for real; nothing was published. -Execute would now run:"
    Say "  git $($pushArgs -join ' ')"
    Say "  gh $($ghArgs -join ' ')"
    Say "  and fast-forward a local $ReleaseBranch that sits on $($head.Substring(0, 12))."
    Say "  Record: $recordPath"
    exit 0
}

# ---------------------------------------------------------------------------------------------
Say ''
Say '== 11. PUBLISH =='
$null = Invoke-Git -Arguments @('fetch', '--quiet', $Remote, $ReleaseBranch)
$remoteNow = (Invoke-Git -Arguments @('rev-parse', '--verify', "refs/remotes/$Remote/$ReleaseBranch^{commit}")).Text
if ($remoteNow -ne $head) { throw "REFUSING: $Remote/$ReleaseBranch moved to $remoteNow while the release was built from $head. Nothing was published; run the release again." }
$push = Invoke-Git -Arguments $pushArgs -AllowFailure
if ($push.ExitCode -ne 0) { throw "REFUSING: the push of the stamp commit was rejected (exit $($push.ExitCode)); nothing was released:`n$($push.Text)" }
Say "  pushed $stampCommit to $Remote/$ReleaseBranch"
$create = Invoke-Gh -Arguments $ghArgs
if ($create.ExitCode -ne 0) {
    throw "The stamp commit is on $Remote/$ReleaseBranch but gh release create failed (exit $($create.ExitCode)):`n$($create.Text)`nRetry exactly this - the commit and the installer are both still right:`n  gh $($ghArgs -join ' ')"
}
Say "  $($create.Text)"
$view = Invoke-Gh -Arguments @('release', 'view', $tag, '--repo', $repository, '--json', 'tagName,assets')
if ($view.ExitCode -ne 0) { throw "gh release view $tag failed after the release was created (exit $($view.ExitCode)): $($view.Text)" }
$published = $view.Text | ConvertFrom-Json
$asset = @($published.assets | Where-Object { $_.name -eq [System.IO.Path]::GetFileName($installer) })
if ($asset.Count -ne 1 -or [long]$asset[0].size -ne $installerItem.Length) { throw "The release $tag exists, but its asset is not the $($installerItem.Length)-byte installer: $($view.Text)" }
Say "  $tag published with $($asset[0].name), $($asset[0].size) bytes"
$branch = (Invoke-Git -Arguments @('rev-parse', '--abbrev-ref', 'HEAD')).Text
if ($branch -eq $ReleaseBranch) {
    $null = Invoke-Git -Arguments @('merge', '--ff-only', '--quiet', $stampCommit)
    Say "  local $ReleaseBranch fast-forwarded to the stamp commit"
}
else { Say "  HEAD is on '$branch', not ${ReleaseBranch}: run 'git pull --rebase' on $ReleaseBranch before continuing work there." }
$record.finished = (Get-Date).ToString('o')
$record.published = $true
Set-Content -LiteralPath $recordPath -Value ($record | ConvertTo-Json -Depth 6) -Encoding UTF8
Say ''
Say "RELEASED $tag in $([int]((Get-Date) - $started).TotalMinutes) minute(s). Record: $recordPath"
exit 0
