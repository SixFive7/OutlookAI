#Requires -Version 5.1
<#
    ============================================================================================
    NEVER RUN AGAINST A REAL GUEST'S VALUES - THERE ARE NONE YET. WRITTEN 2026-09-24. WHAT HAS
    RUN IS ITS SELF-TEST, A SYNTHETIC RENDER AND ITS REFUSALS, ALL ON THE HOST.
    ============================================================================================

    What was executed, on the maintainer's workstation and on no guest:

      * -SelfTest under Windows PowerShell 5.1 and under PowerShell 7: 142 assertions, 0 failures
        on both. It was also run with its live-fixtures rule deliberately broken, and all 15
        assertions about that rule failed - so they are not decorative.
      * 2026-09-24, after the watched/indexed list split and the probe fields: -SelfTest 168
        assertions, 0 failures under both editions, and a synthetic INDEXED-guest render from a
        filled copy of testbed.json in .work/ under both, identical apart from the render time -
        which is also the first run of the git provenance calls under 5.1 with their stderr
        guarded (see Get-GitProvenance).
      * A full render from a SYNTHETIC values file in .work/, under both editions: written, read
        back and parsed again. The two files were byte-identical apart from the render time.
      * The refusal this script exists for. -OutPath was pointed at the main checkout's REAL
        McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json four ways: the
        plain path under both editions (once with -Force), as ..\..\..\ from a worktree, and in
        forward slashes, upper case and a trailing dot. All four were refused by the live-fixtures
        rule before anything was read. The file's LastWriteTime, length and SHA-256, and its
        directory's LastWriteTime and entry count, were identical before and after. This script
        never read its contents and nothing printed them.
      * A render for each of the two guests from Testbed/testbed.json's own sections: both
        refused, each naming all eight of its placeholders, and nothing written.

    What has NOT run: no guest's file has been rendered, because nobody has yet read a guest's
    store names off it - so none has been copied into a guest, and the live tier has never loaded
    one. T1/LiveTestSettingsTemplateTests proves the TEMPLATE renders into something the loader
    accepts; nothing yet proves a GUEST's values do.

    Replace this banner with what it actually did once it has rendered a real guest's file and the
    tier has started on it.

.SYNOPSIS
    Renders one test guest's gitignored live-test-settings.json from the committed template and
    that guest's section of Testbed/testbed.json, checks it the way the live tier will, and
    writes it into gitignored scratch. HOST ONLY. It never writes into a live-fixtures directory.

.DESCRIPTION
    WHY THIS EXISTS. The live tier reads a machine-local settings file naming the machine's
    stores: the hub it may write to, the bystanders nothing may touch, the delegate stores, the
    machine profile, the corpus, the mail sink. Without it the tier has no write allowlist and
    refuses to start - correctly. It names real stores, so it is gitignored and no script could
    supply it; on every guest it had to be written by hand, and a hand-written file is exactly
    where the rules below get broken. So this takes the shape New-AnswerFile.ps1 already has:

      Testbed/live-test-settings.template.json   committed, TOKENS ONLY
      Testbed/testbed.json, liveTestSettings      committed per-guest values, keyed by VM name -
                                                  safe to commit only because a test guest's
                                                  stores are synthetic, named after nothing real
      this script                                 puts the two together for ONE guest, refuses
                                                  anything the tier would refuse, and writes it
                                                  to .work/

    THE HAZARD THIS SCRIPT IS BUILT AROUND. On the host,
    McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json is the maintainer's
    REAL settings file, for his production profile. This script never writes, overwrites, moves,
    reads or prints it, and it REFUSES ANY OUTPUT PATH THAT RESOLVES INTO A DIRECTORY CALLED
    live-fixtures - wider than McpServer/**/live-fixtures/ on purpose. That check runs first,
    before a single file is read, and it runs twice:

      * on the path as spelled, made absolute WITHOUT touching the disk - so '..', relative
        paths, forward slashes, case, trailing dots and spaces, and a '\\?\' or admin-share
        spelling are all refused on their text alone;
      * then on the path WINDOWS reports for the deepest directory that exists
        (GetFinalPathNameByHandle, on a handle opened for no access at all), which is what
        defeats a junction, a symbolic link, a SUBST or mapped drive and an 8.3 short name.
        -SelfTest exercises the junction and short-name cases on a synthetic directory; the
        SUBST and mapped-drive cases rest on that call's documented behaviour and have not been
        exercised here.

    The contents of an existing output file are never read; only what it IS gets asked - a
    directory, a link, a file with more than one name. Inside ANY git working tree it may only
    land under that tree's .work\ - which also keeps it out of the main checkout's tracked files
    when run from a worktree - and an existing file is replaced by writing a new file beside it
    and renaming it over the old NAME, so a hard link or symbolic link at the destination is
    unlinked rather than written through. It refuses those anyway, because one there is almost
    certainly a mistake worth looking at.

    The rendered file reaches the GUEST's live-fixtures directory through
    Testbed/host/Copy-ToGuest.ps1, and this prints the exact command line for it.

    WHAT IT REFUSES, before anything is written and with every reason at once:

      * a guest whose testbed.json values are not ready - any value still beginning '<FILL', any
        field the template needs that the section lacks, any field the section has that the
        template does not (a typo there would be silently ignored by the loader), and any block
        that is neither complete nor null. Each is named.
      * a rendered file with a token left in it, or that does not parse as JSON.
      * everything the live tier itself refuses before it touches a mailbox, checked here so it
        fails in seconds on the host instead of at the top of a guest run: no hub, no store to
        watch, a machineProfile other than Portable (a test guest is Portable by decision - see
        testbed.json's _decided), a partial subjectOnlyProbe, corpus or mailSink block, a delegate
        store inside the identity-draft grant, the hub declared a bystander, a configuration that
        leaves the count tripwire no store it could fail on, and an INDEXED list that names a store
        the census does not watch, a delegate mailbox, or leaves out the hub.
      * and the rules a TEST GUEST's file keeps that the tier does not enforce, because a rendered
        file should not lean on the tier's leniency: every bystander ALSO named in
        expectedStoreDisplayNames - the tier censuses a bystander missing from that list anyway,
        and does not refuse it (Docs/live-tier-on-the-vm.md section 2.10, corrected 2026-09-24 on
        Q77), but outlook_health's reachability check and list_accounts read only that list; the
        corpus store declared a bystander; the hub shaped as an SMTP address, because tests hand
        it to NewDraft as one; every store named as an address ending in .invalid; the indexed
        list ordered hub first and corpus last, with every entry but the corpus named as an
        address; an indexed guest carrying a probe term and a subject-only probe in the hub, and an
        unindexed guest carrying neither and no indexed store; one spelling per store across every
        list; a sink on loopback only; the manifest named corpus-<corpusId>.jsonl; and corpusId,
        seed, anchor and item count agreeing with testbed.json's corpusIdConvention - including the
        indexed guest's minimum corpus size.

    WHAT IT CANNOT CHECK, and it prints this on every render: whether the guest's TIER profile
    actually mounts every declared store under exactly those names. A declared store the running
    profile does not mount is censused, is not found, and refuses the tier. Only the guest can
    answer that, over COM - which is where the placeholder values have to come from.

    Windows PowerShell 5.1 compatible: no ternary, no '??', no three-argument Join-Path. It also
    runs under PowerShell 7, where ConvertFrom-Json would otherwise turn an ISO-8601 anchor into
    a DateTime; that is prevented, and refused where it cannot be undone exactly.

.PARAMETER VMName
    MANDATORY. The guest to render for - a key of the liveTestSettings section of
    Testbed/testbed.json. No default, for the reason Testbed/README.md section 4a gives: several
    guests coexist, and a default that silently picks one is the mistake this testbed keeps
    making. Here it would be a settings file describing one machine's stores copied onto another.

.PARAMETER OutPath
    Where the rendered file goes. Default: .work\live-test-settings\<VMName>\live-test-settings.json
    under the repository root, which is gitignored. Refused anywhere under a directory called
    live-fixtures, and anywhere inside a git working tree except that tree's .work\.

.PARAMETER Force
    Replace an existing file OUTSIDE every git working tree. Inside one - which means under its
    .work\ - a previous render is replaced without it.

.PARAMETER SelfTest
    Run the decision tests and exit. Reads the committed template and testbed.json, and writes
    only a scratch directory under .work\ that it removes again: no guest, no COM, no registry,
    and never a live-fixtures directory other than a synthetic one inside that scratch.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.

.PARAMETER TemplatePath
    The template. Default: Testbed/live-test-settings.template.json.

.PARAMETER TestbedJsonPath
    The file holding the liveTestSettings section. Default: Testbed/testbed.json.

.EXAMPLE
    pwsh -File Testbed/host/New-LiveTestSettings.ps1 -SelfTest

.EXAMPLE
    pwsh -File Testbed/host/New-LiveTestSettings.ps1 -VMName OutlookAI-Indexed
#>
[CmdletBinding(DefaultParameterSetName = 'Render')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Render')] [string] $VMName,
    [Parameter(ParameterSetName = 'Render')] [string] $OutPath,
    [Parameter(ParameterSetName = 'Render')] [switch] $Force,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest,
    [string] $RepoRoot,
    [string] $TemplatePath,
    [string] $TestbedJsonPath
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# Defaulted HERE rather than in param(): Windows PowerShell 5.1 leaves $PSScriptRoot EMPTY inside
# the param() defaults of an advanced script run with -File (measured 2026-09-24; PowerShell 7 fills
# it), so the usual one-liner there throws before the script starts.
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }

# Every placeholder in testbed.json begins with this. Matched ordinally at the START of a string,
# never with -like: '<' is harmless there, but Testbed/README.md section 4b says why this
# directory does not assert with wildcards at all.
$script:PlaceholderMarker = '<FILL'

# A double-brace token. Its name is the JSON path of its own value.
$script:TokenPattern = '\{\{[A-Za-z0-9_.]+\}\}'

# The directory this script never writes into, anywhere, however it is spelled.
$script:ForbiddenDirectory = 'live-fixtures'

# Where the suite is BUILT on a guest - Testbed/host/Publish-LiveTierPayload.ps1 expands Source.zip
# here, and .github/scripts/check-testbed-references.ps1 check 8 pins the two values together.
$script:GuestSourceRoot = 'C:\OutlookAI-Q5\src'
$script:GuestSettingsRelative = 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json'

# What the rules below know about. The template must name exactly these; -SelfTest asserts it and
# a render refuses a template that has drifted, because a field nobody wrote a rule for is a field
# nobody checked.
$script:TopLevelFields = @('machineProfile', 'testHubStoreDisplayName', 'expectedStoreDisplayNames',
    'indexedStoreDisplayNames', 'expectedDelegateStoreDisplayNames', 'bystanderStoreDisplayNames', 'probeTerm',
    'subjectOnlyProbe', 'corpus', 'mailSink')
$script:BlockFields = @{
    subjectOnlyProbe = @('storeDisplayName', 'folderPath', 'subjectTerm', 'senderFragment')
    corpus           = @('storeDisplayName', 'manifestPath', 'corpusId', 'seed', 'anchorUtc', 'itemCount', 'windowDays')
    mailSink         = @('submitHost', 'submitPort', 'retrieveHost', 'retrievePort', 'connectTimeoutMs')
}

# A store name shaped like an SMTP address. Matched with -cmatch against a literal pattern, never -like.
$script:AddressPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

# MailSinkSettings says "Loopback, always", and Docs/live-tier-on-the-vm.md section 2.7 says why: a
# listener that needs anything else is a listener bound to 0.0.0.0, which on a test VM is an open relay.
$script:LoopbackHosts = @('127.0.0.1', 'localhost', '::1')

$script:Invariant = [System.Globalization.CultureInfo]::InvariantCulture

# =============================================================================================
# WINDOWS' OWN ANSWER TO "WHERE IS THIS, REALLY". kernel32 only, compiled by Add-Type - which on
# Windows PowerShell 5.1 is the .NET Framework's own csc.exe, so nothing is installed for it (the
# repository's Dependencies rule). C# 5, because that is what that compiler speaks.
# =============================================================================================

if (-not ('OutlookAI.Testbed.LiveSettingsPathProbe' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace OutlookAI.Testbed
{
    public static class LiveSettingsPathProbe
    {
        private const uint ShareAll = 0x00000007;             // read | write | delete
        private const uint OpenExisting = 3;
        private const uint BackupSemantics = 0x02000000;      // needed to open a directory at all

        [StructLayout(LayoutKind.Sequential)]
        private struct FileInformation
        {
            public uint FileAttributes;
            public uint CreationTimeLow;
            public uint CreationTimeHigh;
            public uint LastAccessTimeLow;
            public uint LastAccessTimeHigh;
            public uint LastWriteTimeLow;
            public uint LastWriteTimeHigh;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes,
            uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandleW(
            SafeFileHandle file, StringBuilder filePath, uint filePathLength, uint flags);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle file, out FileInformation information);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetShortPathNameW(string longPath, StringBuilder shortPath, uint bufferLength);

        // Desired access 0: the handle can answer questions about the object and do nothing else -
        // no byte of it is read or written.
        private static SafeFileHandle OpenForMetadata(string path)
        {
            SafeFileHandle handle = CreateFileW(path, 0, ShareAll, IntPtr.Zero, OpenExisting, BackupSemantics, IntPtr.Zero);
            if (handle.IsInvalid)
            {
                int error = Marshal.GetLastWin32Error();
                handle.Dispose();
                throw new IOException("Windows would not open '" + path + "' to say where it really is (Win32 error " + error + ").");
            }
            return handle;
        }

        public static string FinalPath(string path)
        {
            using (SafeFileHandle handle = OpenForMetadata(path))
            {
                StringBuilder buffer = new StringBuilder(512);
                uint length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
                if (length >= buffer.Capacity)
                {
                    buffer = new StringBuilder((int)length + 1);
                    length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
                }
                if (length == 0 || length >= buffer.Capacity)
                {
                    throw new IOException("Windows could not report the final path of '" + path + "' (Win32 error " + Marshal.GetLastWin32Error() + ").");
                }
                return buffer.ToString();
            }
        }

        public static uint LinkCount(string path)
        {
            using (SafeFileHandle handle = OpenForMetadata(path))
            {
                FileInformation information;
                if (!GetFileInformationByHandle(handle, out information))
                {
                    throw new IOException("Windows could not report the link count of '" + path + "' (Win32 error " + Marshal.GetLastWin32Error() + ").");
                }
                return information.NumberOfLinks;
            }
        }

        // Null when the volume keeps no 8.3 names, or the path does not exist.
        public static string ShortPath(string path)
        {
            StringBuilder buffer = new StringBuilder(1024);
            uint length = GetShortPathNameW(path, buffer, (uint)buffer.Capacity);
            if (length == 0 || length >= buffer.Capacity) { return null; }
            return buffer.ToString();
        }
    }
}
'@
}

# =============================================================================================
# PURE DECISIONS. Everything this script decides is decided here, from arguments, so -SelfTest
# decides it too. The only ones that touch the disk are the path checks, and those only ask
# Windows about directories; they never open the output file.
# =============================================================================================

function Test-IsInteger {
    param($Value)
    if ($null -eq $Value) { return $false }
    return ($Value -is [int] -or $Value -is [long] -or $Value -is [int16] -or $Value -is [byte] -or
        $Value -is [sbyte] -or $Value -is [uint16] -or $Value -is [uint32] -or $Value -is [uint64])
}

function Test-IsNumber {
    param($Value)
    if ($null -eq $Value) { return $false }
    if (Test-IsInteger $Value) { return $true }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) { return $true }
    # By name: System.Numerics is not loaded in every Windows PowerShell session.
    return ($Value.GetType().FullName -eq 'System.Numerics.BigInteger')
}

function Test-IsJsonObject {
    param($Value)
    return ($null -ne $Value -and $Value -is [System.Management.Automation.PSCustomObject])
}

function Test-IsPlaceholder {
    param($Value)
    if ($Value -is [string]) {
        return $Value.StartsWith($script:PlaceholderMarker, [System.StringComparison]::OrdinalIgnoreCase)
    }
    if ($Value -is [System.Array]) {
        foreach ($item in $Value) { if (Test-IsPlaceholder $item) { return $true } }
    }
    return $false
}

<#
    A property by its EXACT name. PowerShell's own property lookup ignores case, and a field
    spelled differently from the template's is a typo this script exists to catch rather than
    quietly accept.
#>
function Get-ExactProperty {
    param($Object, [string] $Name)
    if (-not (Test-IsJsonObject $Object)) { return $null }
    foreach ($property in $Object.PSObject.Properties) {
        if ($property.Name -ceq $Name) { return $property }
    }
    return $null
}

<#
    Undoes PowerShell 7's date coercion. ConvertFrom-Json there turns "2026-08-19T00:00:00Z" into
    a DateTime, and a DateTime written back out loses its 'Z' - the exact trap
    Testbed/host/TestbedLeasePath.ps1 records. -DateKind String prevents it on 7.5 and later;
    this repairs the rest, and refuses the one case it cannot repair exactly.
#>
function Repair-JsonDates {
    param($Node, [string] $Path = '')
    if ($null -eq $Node) { return $null }
    if ($Node -is [datetime]) {
        if ($Node.Kind -ne [System.DateTimeKind]::Utc -or ($Node.Ticks % [TimeSpan]::TicksPerSecond) -ne 0) {
            throw "'$Path' was read as a date and time that is not a whole-second UTC instant, so it cannot be written back exactly as it was. Write it yyyy-MM-ddTHH:mm:ssZ."
        }
        return $Node.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $script:Invariant)
    }
    if (Test-IsJsonObject $Node) {
        foreach ($property in $Node.PSObject.Properties) {
            $childPath = $property.Name
            if ($Path) { $childPath = $Path + '.' + $property.Name }
            $property.Value = Repair-JsonDates -Node $property.Value -Path $childPath
        }
        return $Node
    }
    if ($Node -is [System.Array]) {
        $copy = New-Object object[] $Node.Length
        for ($i = 0; $i -lt $Node.Length; $i++) {
            $copy[$i] = Repair-JsonDates -Node $Node[$i] -Path ('{0}[{1}]' -f $Path, $i)
        }
        return , $copy
    }
    return $Node
}

function ConvertFrom-JsonText {
    param([AllowEmptyString()] [string] $Text, [string] $What)
    if ([string]::IsNullOrWhiteSpace($Text)) { throw "$What is empty." }
    $command = Get-Command -Name ConvertFrom-Json
    try {
        if ($command.Parameters.ContainsKey('DateKind')) {
            $parsed = ConvertFrom-Json -InputObject $Text -DateKind String
        }
        else {
            $parsed = ConvertFrom-Json -InputObject $Text
        }
    }
    catch {
        throw "$What is not valid JSON: $($_.Exception.Message)"
    }
    return (Repair-JsonDates -Node $parsed)
}

function ConvertTo-JsonString {
    param([AllowEmptyString()] [string] $Text)
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    foreach ($character in $Text.ToCharArray()) {
        $code = [int]$character
        if ($code -eq 0x22) { [void]$builder.Append('\"') }
        elseif ($code -eq 0x5C) { [void]$builder.Append('\\') }
        elseif ($code -lt 0x20 -or $code -gt 0x7E) {
            # Everything outside printable ASCII is escaped, so the rendered file is pure ASCII and no
            # tool on the guest can misread its encoding.
            [void]$builder.Append('\u').Append($code.ToString('x4', $script:Invariant))
        }
        else { [void]$builder.Append($character) }
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

<#
    A small JSON writer, rather than ConvertTo-Json, for three reasons that each bit somebody: its
    default depth of 2 silently flattens a nested list into the text "System.Object[]"; its layout
    and escaping differ between Windows PowerShell and PowerShell 7, so one input rendered two ways;
    and a one-element list piped through it can come out as a bare string, which the loader then
    refuses as the wrong type. This writes the few shapes a settings file has, identically on both.
#>
function ConvertTo-JsonLiteral {
    param($Value, [int] $Level = 0)
    $indent = '  ' * ($Level + 1)
    $closing = '  ' * $Level
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [string]) { return (ConvertTo-JsonString -Text $Value) }
    if ($Value -is [bool]) {
        if ($Value) { return 'true' }
        return 'false'
    }
    if (Test-IsNumber $Value) { return [System.Convert]::ToString($Value, $script:Invariant) }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Count -eq 0) { return '{}' }
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($key in $Value.Keys) {
            $lines.Add($indent + (ConvertTo-JsonString -Text ([string]$key)) + ': ' + (ConvertTo-JsonLiteral -Value $Value[$key] -Level ($Level + 1)))
        }
        return "{`n" + ($lines -join ",`n") + "`n" + $closing + '}'
    }
    if (Test-IsJsonObject $Value) {
        $ordered = [ordered]@{}
        foreach ($property in $Value.PSObject.Properties) { $ordered[$property.Name] = $property.Value }
        return (ConvertTo-JsonLiteral -Value $ordered -Level $Level)
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $parts = New-Object System.Collections.Generic.List[string]
        $allScalar = $true
        foreach ($item in $Value) {
            if ($item -is [System.Collections.IDictionary] -or (Test-IsJsonObject $item) -or
                ($item -is [System.Collections.IEnumerable] -and -not ($item -is [string]))) {
                $allScalar = $false
            }
            $parts.Add((ConvertTo-JsonLiteral -Value $item -Level ($Level + 1)))
        }
        if ($parts.Count -eq 0) { return '[]' }
        if ($allScalar) { return '[ ' + ($parts -join ', ') + ' ]' }
        return "[`n" + $indent + ($parts -join (",`n" + $indent)) + "`n" + $closing + ']'
    }
    throw "Cannot write a value of type $($Value.GetType().FullName) as JSON."
}

<#
    Reads the template's shape: which fields it has, in order, and whether every value in it is a
    token spelling its own path. Returns Fields (Path, Block, Key) and Problems.
#>
function Get-TemplateFields {
    param($Template)
    $fields = New-Object System.Collections.Generic.List[object]
    $problems = New-Object System.Collections.Generic.List[string]
    if (-not (Test-IsJsonObject $Template)) {
        $problems.Add('the template is not a JSON object.')
        return [pscustomobject]@{ Fields = $fields.ToArray(); Problems = $problems.ToArray() }
    }
    foreach ($property in $Template.PSObject.Properties) {
        if ($property.Name.StartsWith('_')) { continue }
        if (Test-IsJsonObject $property.Value) {
            foreach ($inner in $property.Value.PSObject.Properties) {
                if ($inner.Name.StartsWith('_')) { continue }
                $path = $property.Name + '.' + $inner.Name
                if (Test-IsJsonObject $inner.Value) {
                    $problems.Add("'$path' nests a block inside a block. The renderer handles one level, which is all a settings file has.")
                    continue
                }
                if (-not ($inner.Value -is [string]) -or $inner.Value -cne ('{{' + $path + '}}')) {
                    $problems.Add("'$path' is not the token for its own path, {{$path}}. The template holds tokens only - never a value - and every token spells where its value comes from.")
                }
                $fields.Add([pscustomobject]@{ Path = $path; Block = $property.Name; Key = $inner.Name })
            }
            continue
        }
        $path = $property.Name
        if (-not ($property.Value -is [string]) -or $property.Value -cne ('{{' + $path + '}}')) {
            $problems.Add("'$path' is not the token for its own path, {{$path}}. The template holds tokens only - never a value - and every token spells where its value comes from.")
        }
        $fields.Add([pscustomobject]@{ Path = $path; Block = $null; Key = $path })
    }
    return [pscustomobject]@{ Fields = $fields.ToArray(); Problems = $problems.ToArray() }
}

<#
    Whether the template names exactly the fields the rules below were written for. A difference
    means somebody changed the template (and so the example, which check 8 holds it to) without
    teaching this script - and a field nobody wrote a rule for is a field nobody checked.
#>
function Get-TemplateDriftProblems {
    param($Fields)
    $problems = New-Object System.Collections.Generic.List[string]
    $expected = New-Object System.Collections.Generic.List[string]
    foreach ($name in $script:TopLevelFields) {
        if ($script:BlockFields.ContainsKey($name)) {
            foreach ($inner in $script:BlockFields[$name]) { $expected.Add($name + '.' + $inner) }
        }
        else { $expected.Add($name) }
    }
    $actual = @($Fields | ForEach-Object { $_.Path })
    foreach ($path in $expected) {
        if ($actual -cnotcontains $path) { $problems.Add("the template has no '$path', which this script has a rule for.") }
    }
    foreach ($path in $actual) {
        if ($expected -cnotcontains $path) { $problems.Add("the template has '$path', which this script has no rule for.") }
    }
    return , $problems.ToArray()
}

<#
    Matches a guest's section against the template's fields. Returns ByPath (path -> value),
    Omitted (blocks declared null) and Problems - every problem, not the first, because a guest
    being filled in has several gaps at once and finding them one render at a time is slow.
#>
function Resolve-GuestValues {
    param($Fields, $Guest, [string] $Where)
    $byPath = @{}
    $omitted = New-Object System.Collections.Generic.List[string]
    $problems = New-Object System.Collections.Generic.List[string]

    if (-not (Test-IsJsonObject $Guest)) {
        $problems.Add("$Where is not a JSON object.")
        return [pscustomobject]@{ ByPath = $byPath; Omitted = $omitted.ToArray(); Problems = $problems.ToArray() }
    }

    $topNames = New-Object System.Collections.Generic.List[string]
    $blockKeys = @{}
    foreach ($field in $Fields) {
        if ($null -eq $field.Block) { $topNames.Add($field.Key); continue }
        if (-not $blockKeys.ContainsKey($field.Block)) {
            $topNames.Add($field.Block)
            $blockKeys[$field.Block] = New-Object System.Collections.Generic.List[string]
        }
        $blockKeys[$field.Block].Add($field.Key)
    }

    foreach ($property in $Guest.PSObject.Properties) {
        if ($property.Name.StartsWith('_')) { continue }
        if ($topNames -cnotcontains $property.Name) {
            $problems.Add("$($property.Name): not a field of the template. A misspelt field is ignored by the loader, so the value never arrives - check the spelling against Testbed/live-test-settings.template.json.")
        }
    }

    foreach ($name in $topNames) {
        $property = Get-ExactProperty $Guest $name
        if ($null -eq $property) {
            $problems.Add("${name}: missing. Every field of the template needs a value here (a block may be null).")
            continue
        }
        $value = $property.Value

        if (-not $blockKeys.ContainsKey($name)) {
            if (Test-IsPlaceholder $value) {
                $problems.Add("${name}: still a placeholder.")
            }
            else { $byPath[$name] = $value }
            continue
        }

        # A block: an object with every field, or null for "none on this guest".
        if ($null -eq $value) { $omitted.Add($name); continue }
        if (Test-IsPlaceholder $value) {
            $problems.Add("${name}: still a placeholder - give it an object carrying every field, or null.")
            continue
        }
        if (-not (Test-IsJsonObject $value)) {
            $problems.Add("${name}: must be an object carrying every field, or null. It is neither.")
            continue
        }
        foreach ($inner in $value.PSObject.Properties) {
            if ($inner.Name.StartsWith('_')) { continue }
            if ($blockKeys[$name] -cnotcontains $inner.Name) {
                $problems.Add("$name.$($inner.Name): not a field of the template's $name block. The loader would ignore it.")
            }
        }
        foreach ($key in $blockKeys[$name]) {
            $path = $name + '.' + $key
            $innerProperty = Get-ExactProperty $value $key
            if ($null -eq $innerProperty) {
                $problems.Add("${path}: missing. A $name block is all or nothing - the loader refuses a partial one.")
                continue
            }
            if (Test-IsPlaceholder $innerProperty.Value) {
                $problems.Add("${path}: still a placeholder.")
                continue
            }
            $byPath[$path] = $innerProperty.Value
        }
    }

    return [pscustomobject]@{ ByPath = $byPath; Omitted = $omitted.ToArray(); Problems = $problems.ToArray() }
}

function New-RenderedDocument {
    param(
        $Fields, $Resolved, [string] $VMName, [string] $Provenance, [datetime] $RenderedAtUtc,
        [string] $TemplateShown = 'Testbed/live-test-settings.template.json',
        [string] $ValuesShown = 'Testbed/testbed.json'
    )
    $stamp = $RenderedAtUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $script:Invariant)
    $document = [ordered]@{}
    $document['_rendered'] = "Rendered for guest '$VMName' by Testbed/host/New-LiveTestSettings.ps1 from $TemplateShown and the liveTestSettings.$VMName section of $ValuesShown at $stamp, from $Provenance. Do not edit it where it lands: correct the values and render again, or the next render silently undoes the edit."
    $document['_what'] = "The live tier's machine-local settings for this guest: the one store it may write to, the stores it must never touch, and what the count tripwire watches. Gitignored where it lands, like every live-test-settings.json. Every field is described in Docs/live-tier-on-the-vm.md section 2.10."
    if ($Resolved.Omitted -contains 'subjectOnlyProbe') {
        $document['_subjectOnlyProbe'] = 'Absent: testbed.json declares no subject-only probe population for this guest - it has no search index, and every test reading one measures the index.'
    }
    if ($Resolved.Omitted -contains 'corpus') {
        $document['_corpus'] = 'Absent: testbed.json declares no corpus for this guest, so the tier skips the corpus freshness check.'
    }
    if ($Resolved.Omitted -contains 'mailSink') {
        $document['_mailSink'] = "Absent: testbed.json declares no sink for this guest (Docs/live-tier-on-the-vm.md section 1.4). The loader reads absent as 'this machine has real transport' - the ambiguity section 2.10 names. Nothing listens here, so a send would queue in the Outbox: do not send."
    }
    foreach ($field in $Fields) {
        if ($null -eq $field.Block) {
            $document[$field.Key] = $Resolved.ByPath[$field.Path]
            continue
        }
        if ($Resolved.Omitted -contains $field.Block) { continue }
        if (-not $document.Contains($field.Block)) { $document[$field.Block] = [ordered]@{} }
        $document[$field.Block][$field.Key] = $Resolved.ByPath[$field.Path]
    }
    return $document
}

function Get-FieldValue {
    param($Object, [string] $Name)
    $property = Get-ExactProperty $Object $Name
    if ($null -eq $property) { return [pscustomobject]@{ Present = $false; Value = $null } }
    return [pscustomobject]@{ Present = $true; Value = $property.Value }
}

function Test-StoreName {
    param($Name, [string] $Where)
    if (-not ($Name -is [string])) { return "${Where}: holds a value that is not a string. Store display names are strings." }
    if ([string]::IsNullOrWhiteSpace($Name)) { return "${Where}: holds an empty store name." }
    if (Test-IsPlaceholder $Name) { return "${Where}: still holds a placeholder." }
    if ($Name -cne $Name.Trim()) {
        return "${Where}: '$Name' has leading or trailing whitespace. Outlook trims a store's name (it is its root folder's name), so the store would never match."
    }
    if ($Name.IndexOfAny([char[]]@('\', '/')) -ge 0) {
        return "${Where}: '$Name' contains a slash, which Outlook does not allow in a store's name."
    }
    return $null
}

function Get-StoreNameList {
    param($Settings, [string] $Field, [switch] $AllowEmpty, [System.Collections.Generic.List[string]] $Problems)
    $names = New-Object System.Collections.Generic.List[string]
    $found = Get-FieldValue $Settings $Field
    if (-not $found.Present) {
        $Problems.Add("${Field}: missing.")
        return , $names.ToArray()
    }
    if (-not ($found.Value -is [System.Array])) {
        $Problems.Add("${Field}: must be a list of store display names.")
        return , $names.ToArray()
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $found.Value) {
        $problem = Test-StoreName -Name $item -Where $Field
        if ($null -ne $problem) { $Problems.Add($problem); continue }
        if (-not $seen.Add($item)) {
            $Problems.Add("${Field}: names '$item' twice. The tier matches store names ignoring case, so these are one store.")
            continue
        }
        $names.Add($item)
    }
    if ($found.Value.Count -eq 0 -and -not $AllowEmpty) { $Problems.Add("${Field}: is empty.") }
    return , $names.ToArray()
}

function Test-NameIn {
    param([string] $Name, [string[]] $List)
    foreach ($item in $List) {
        if ([string]::Equals($item, $Name, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function ConvertFrom-AnchorText {
    param($Text)
    if (-not ($Text -is [string])) { return $null }
    if ($Text -cnotmatch '^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}Z)?$') { return $null }
    $parsed = [datetime]::MinValue
    $formats = [string[]]@('yyyy-MM-dd', "yyyy-MM-dd'T'HH:mm:ss'Z'")
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    if ([datetime]::TryParseExact($Text, $formats, $script:Invariant, $styles, [ref]$parsed)) { return $parsed }
    return $null
}

<#
    Every rule a rendered settings file must satisfy, applied to the PARSED ARTEFACT rather than to
    the values that went into it - what is checked is what will be read. $Assigned is testbed.json's
    corpusIdConvention.assigned.
#>
function Add-SettingsProblems {
    param(
        $Settings,
        [string] $VMName,
        $Assigned,
        [System.Collections.Generic.List[string]] $Problems
    )

    if (-not (Test-IsJsonObject $Settings)) {
        $Problems.Add('the settings are not a JSON object.')
        return
    }

    foreach ($property in $Settings.PSObject.Properties) {
        if ($property.Name.StartsWith('_')) { continue }
        if ($script:TopLevelFields -cnotcontains $property.Name) {
            $Problems.Add("$($property.Name): not a field this script has a rule for.")
        }
    }

    # machineProfile - Portable, spelled as the documented example spells it.
    $machineProfile = Get-FieldValue $Settings 'machineProfile'
    if (-not $machineProfile.Present) {
        $Problems.Add("machineProfile: missing. Absent means Production to the loader, which this template cannot satisfy.")
    }
    elseif ($machineProfile.Value -is [string] -and $machineProfile.Value -ceq 'Portable') { }
    elseif ($machineProfile.Value -is [string] -and $machineProfile.Value -ceq 'Production') {
        $Problems.Add("machineProfile: 'Production' is the maintainer's own machine, whose missing populations REFUSE a run. A test guest is Portable - recorded as a decision in testbed.json's _decided - so a population it lacks is announced as PROVED NOTHING instead, and its probe values come from the generator rather than from real mail.")
    }
    else {
        $Problems.Add("machineProfile: must be the string 'Portable' (or 'Production', which this template cannot carry). Spell it exactly.")
    }

    # testHubStoreDisplayName - where the suite writes, and an address.
    $hub = $null
    $hubValue = Get-FieldValue $Settings 'testHubStoreDisplayName'
    if (-not $hubValue.Present) {
        $Problems.Add('testHubStoreDisplayName: missing. The hub is the one store the suite may write to; the loader refuses a file without one.')
    }
    else {
        $problem = Test-StoreName -Name $hubValue.Value -Where 'testHubStoreDisplayName'
        if ($null -ne $problem) { $Problems.Add($problem) }
        elseif ($hubValue.Value -cnotmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $Problems.Add("testHubStoreDisplayName: '$($hubValue.Value)' is not shaped like an SMTP address. It doubles as one - tests call NewDraft(Hub, Hub, ...) and FindAccountBySmtp(Hub) - so the hub store must be named exactly as its account's address (Docs/live-tier-on-the-vm.md section 2.6).")
        }
        else { $hub = $hubValue.Value }
    }

    $expected = Get-StoreNameList -Settings $Settings -Field 'expectedStoreDisplayNames' -Problems $Problems
    $indexed = Get-StoreNameList -Settings $Settings -Field 'indexedStoreDisplayNames' -AllowEmpty -Problems $Problems
    $delegates = Get-StoreNameList -Settings $Settings -Field 'expectedDelegateStoreDisplayNames' -AllowEmpty -Problems $Problems
    $bystanders = Get-StoreNameList -Settings $Settings -Field 'bystanderStoreDisplayNames' -AllowEmpty -Problems $Problems

    # The guest's own record in corpusIdConvention.assigned - which corpus id it may use, and whether
    # it is the INDEXED guest. Looked up once: the indexed-list rules and the corpus rules both read it.
    $mine = @(@($Assigned) | Where-Object { (Test-IsJsonObject $_) -and $_.guest -is [string] -and $_.guest -ceq $VMName })
    $entry = $null
    $guestIndexed = $null
    if ($mine.Count -ne 1) {
        $Problems.Add("corpusIdConvention.assigned in testbed.json has $($mine.Count) entries for guest '$VMName', where exactly one is needed - it is the record that says which corpus id this guest may use and whether it is the indexed guest.")
    }
    else {
        $entry = $mine[0]
        $indexedProperty = Get-ExactProperty $entry 'indexed'
        if ($null -ne $indexedProperty -and $indexedProperty.Value -is [bool]) { $guestIndexed = [bool]$indexedProperty.Value }
        else { $Problems.Add("corpusIdConvention.assigned's entry for '$VMName' carries no true/false 'indexed', so nothing says whether this guest's index is on - and the indexed list, the probe term and the subject-only probe all depend on it.") }
    }

    # The corpus store's name, peeked at here: the indexed list's ORDER is stated against it.
    $corpusName = $null
    $corpusPeek = Get-FieldValue $Settings 'corpus'
    if ($corpusPeek.Present -and (Test-IsJsonObject $corpusPeek.Value)) {
        $peeked = (Get-FieldValue $corpusPeek.Value 'storeDisplayName').Value
        if ($peeked -is [string]) { $corpusName = $peeked }
    }

    if ($null -ne $hub -and $expected.Count -gt 0 -and -not (Test-NameIn $hub $expected)) {
        $Problems.Add("expectedStoreDisplayNames: does not name the hub '$hub'. It is the list the census walks, and the remediation console requires the hub in it.")
    }

    foreach ($store in $delegates) {
        if ($null -ne $hub -and [string]::Equals($store, $hub, [System.StringComparison]::OrdinalIgnoreCase)) {
            $Problems.Add("expectedDelegateStoreDisplayNames: names the hub '$store'. A delegate mailbox is read-only for tests; the hub is where they write.")
        }
        elseif (Test-NameIn $store $expected) {
            $Problems.Add("expectedDelegateStoreDisplayNames: '$store' is also in expectedStoreDisplayNames, which puts a read-only store inside the identity-draft grant - the write allowlist refuses to build at all.")
        }
    }

    foreach ($store in $bystanders) {
        if ($null -ne $hub -and [string]::Equals($store, $hub, [System.StringComparison]::OrdinalIgnoreCase)) {
            $Problems.Add("bystanderStoreDisplayNames: names the hub '$store'. The hub is where the suite writes; the tier refuses a hub declared a bystander.")
            continue
        }
        if (-not (Test-NameIn $store $expected)) {
            $Problems.Add("bystanderStoreDisplayNames: '$store' is not ALSO in expectedStoreDisplayNames. The tier would still census it - it adds declared bystanders back in, deliberately - but on a test guest every mounted store goes in the watched list, because outlook_health's reachability check and list_accounts exactness read that list and nothing else (Docs/live-tier-on-the-vm.md section 2.10). Name it in both.")
        }
    }

    # The tripwire's own refusal, restated: something watched must be non-hub AND denied every write.
    $policed = New-Object System.Collections.Generic.List[string]
    foreach ($store in (@($bystanders) + @($delegates))) {
        if ($null -ne $hub -and [string]::Equals($store, $hub, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not (Test-NameIn $store $policed.ToArray())) { $policed.Add($store) }
    }
    if ($policed.Count -eq 0) {
        $Problems.Add("bystanderStoreDisplayNames: declares nothing the count tripwire could fail on - no watched store that is both non-hub and denied every write. The tier refuses exactly this at start ('NO STORE THIS CENSUS WATCHES CAN PRODUCE A FAILURE'). Declare the plain bystander, and the corpus store, in bystanderStoreDisplayNames.")
    }

    # indexedStoreDisplayNames - the stores the index tier measures. The loader's own rules first
    # (watched, never a delegate, the hub included), then the ORDER and NAMING a test guest keeps.
    foreach ($store in $indexed) {
        if (Test-NameIn $store $delegates) {
            $Problems.Add("indexedStoreDisplayNames: '$store' is a delegate mailbox. A delegate is indexed under its owner's subtree and has no index scope of its own; the loader refuses it.")
        }
        elseif (-not (Test-NameIn $store (@($expected) + @($bystanders)))) {
            $Problems.Add("indexedStoreDisplayNames: '$store' is in neither expectedStoreDisplayNames nor bystanderStoreDisplayNames, so the tests would read a store the count tripwire never censuses. The loader refuses it.")
        }
        elseif ($null -ne $corpusName -and [string]::Equals($store, $corpusName, [System.StringComparison]::OrdinalIgnoreCase)) { }
        elseif ($store -cnotmatch $script:AddressPattern) {
            $Problems.Add("indexedStoreDisplayNames: '$store' is not named as an address. Only the corpus - big enough to dominate the index's 2000-row discovery sample - may be named otherwise: a small store is found only through mail addressed to it, and the product's own search and outlook_health try that only for a name shaped like an address (MailService.ResolveFolderScope, ProbeStoreInIndex). Name the store as an .invalid address.")
        }
    }
    if ($indexed.Count -gt 0) {
        if ($null -ne $hub -and -not [string]::Equals($indexed[0], $hub, [System.StringComparison]::OrdinalIgnoreCase)) {
            $Problems.Add("indexedStoreDisplayNames: the hub '$hub' must come FIRST. Several index tests read entry 0 - the post-filter, the filter shapes, the sender filter - and the hub population is the one that carries attachments, unread mail and senders. The loader also refuses a non-empty list without the hub.")
        }
        if ($null -ne $corpusName -and (Test-NameIn $corpusName $indexed) -and
            -not [string]::Equals($indexed[$indexed.Count - 1], $corpusName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $Problems.Add("indexedStoreDisplayNames: the corpus store '$corpusName' must come LAST. The exclude-subfolders measurement takes the first non-hub entry and needs a mail folder with populated children - which the bystander population has and the measurement corpus does not.")
        }
    }
    if ($guestIndexed -eq $true -and $indexed.Count -eq 0) {
        $Problems.Add("indexedStoreDisplayNames: empty, on the guest testbed.json records as INDEXED. Name the hub, the bystander and the corpus store - in that order.")
    }
    if ($guestIndexed -eq $false -and $indexed.Count -gt 0) {
        $Problems.Add("indexedStoreDisplayNames: names $($indexed.Count) store(s) on a guest testbed.json records as UNINDEXED. Nothing is indexed there, so every test measuring these would fail by construction; the list is [] and the index tests are deselected with Requires!=SearchIndex.")
    }

    # probeTerm - one word, set exactly when the guest has an index.
    $probeTerm = Get-FieldValue $Settings 'probeTerm'
    $probeText = $null
    if (-not $probeTerm.Present) { $Problems.Add('probeTerm: missing. It is a string - the generator''s probe term on the indexed guest, empty on the unindexed one.') }
    elseif (-not ($probeTerm.Value -is [string])) { $Problems.Add('probeTerm: must be a string.') }
    else {
        $probeText = [string]$probeTerm.Value
        if ($probeText.Length -gt 0 -and $probeText -cnotmatch '^[A-Za-z]{3,}$') {
            $Problems.Add("probeTerm: '$probeText' is not one plain word. The tests put it inside a CONTAINS phrase and a search query as it stands.")
        }
        if ($guestIndexed -eq $true -and $probeText.Length -eq 0) {
            $Problems.Add('probeTerm: empty on the INDEXED guest. Six index tests search for it - use the population generator''s own term, which corpus-plan --population hub prints.')
        }
        if ($guestIndexed -eq $false -and $probeText.Length -gt 0) {
            $Problems.Add("probeTerm: '$probeText' on a guest with no index. Leave it empty: it names a word 'proven to hit this machine's search index', and there is no index.")
        }
    }

    # subjectOnlyProbe - the SF-6 population, which the hub population holds.
    $subjectProbe = Get-FieldValue $Settings 'subjectOnlyProbe'
    $subjectProbeStore = $null
    if ($subjectProbe.Present) {
        $block = $subjectProbe.Value
        if (-not (Test-IsJsonObject $block)) {
            $Problems.Add('subjectOnlyProbe: must be an object carrying every field. An absent one is left out of the file, never written as null or anything else.')
        }
        else {
            foreach ($inner in $block.PSObject.Properties) {
                if ($inner.Name.StartsWith('_')) { continue }
                if ($script:BlockFields['subjectOnlyProbe'] -cnotcontains $inner.Name) { $Problems.Add("subjectOnlyProbe.$($inner.Name): not a field this script has a rule for.") }
            }
            foreach ($name in @($script:BlockFields['subjectOnlyProbe'] | Where-Object { -not (Get-FieldValue $block $_).Present })) {
                $Problems.Add("subjectOnlyProbe.${name}: missing. The block is all or nothing - the loader refuses a partial one.")
            }
            $storeName = (Get-FieldValue $block 'storeDisplayName').Value
            $problem = Test-StoreName -Name $storeName -Where 'subjectOnlyProbe.storeDisplayName'
            if ($null -ne $problem) { $Problems.Add($problem) }
            else {
                $subjectProbeStore = $storeName
                if ($null -ne $hub -and -not [string]::Equals($storeName, $hub, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $Problems.Add("subjectOnlyProbe.storeDisplayName: '$storeName' is not the hub '$hub'. The generated population that holds the subject-only probe folder is the HUB's.")
                }
                if (-not (Test-NameIn $storeName $indexed)) {
                    $Problems.Add("subjectOnlyProbe.storeDisplayName: '$storeName' is not in indexedStoreDisplayNames, and every SF-6 test reads it through the index.")
                }
            }
            $folderPath = (Get-FieldValue $block 'folderPath').Value
            if (-not ($folderPath -is [string]) -or $folderPath.Length -eq 0 -or $folderPath.StartsWith('/') -or $folderPath.EndsWith('/') -or $folderPath.Contains('\')) {
                $Problems.Add('subjectOnlyProbe.folderPath: must be a store-relative path with forward slashes and none at either end, like Inbox/OutlookAI-Corpus-Folder-Notices.')
            }
            $subjectTerm = (Get-FieldValue $block 'subjectTerm').Value
            if (-not ($subjectTerm -is [string]) -or $subjectTerm -cnotmatch '^[A-Za-z]{5,}$') {
                $Problems.Add('subjectOnlyProbe.subjectTerm: must be one word of at least five letters - the prefix-stem test shortens it by two.')
            }
            $senderFragment = (Get-FieldValue $block 'senderFragment').Value
            if (-not ($senderFragment -is [string]) -or $senderFragment -cnotmatch '^[A-Za-z0-9]{3,}$') {
                $Problems.Add('subjectOnlyProbe.senderFragment: must be one word of letters and digits, at least three long.')
            }
        }
    }
    if ($guestIndexed -eq $true -and -not $subjectProbe.Present) {
        $Problems.Add('subjectOnlyProbe: absent on the INDEXED guest. Five SF-6 tests read it, and the hub population holds exactly that population - corpus-plan --population hub prints the four values.')
    }
    if ($guestIndexed -eq $false -and $subjectProbe.Present) {
        $Problems.Add('subjectOnlyProbe: present on a guest with no index. Every test that reads it measures the index; leave it out (null in testbed.json).')
    }

    # Every store named as an address is a synthetic one, under RFC 2606's .invalid - a guest's
    # stores are named after nothing real, and an address that resolved would be one somebody owns.
    $allNamed = @()
    if ($null -ne $hub) { $allNamed += $hub }
    $allNamed += @($expected) + @($indexed) + @($delegates) + @($bystanders)
    if ($null -ne $corpusName) { $allNamed += $corpusName }
    if ($null -ne $subjectProbeStore) { $allNamed += $subjectProbeStore }
    $reported = @()
    foreach ($store in $allNamed) {
        if ($store -is [string] -and $store -cmatch $script:AddressPattern -and
            -not $store.EndsWith('.invalid', [System.StringComparison]::OrdinalIgnoreCase) -and $reported -notcontains $store) {
            $reported += $store
            $Problems.Add("'$store' is named as an address outside the RFC 2606 .invalid domain. A test guest's stores are synthetic; an address that can resolve is one somebody could own.")
        }
    }

    # corpus - all or nothing, a declared bystander, and the corpus the testbed record says it is.
    $corpusValue = Get-FieldValue $Settings 'corpus'
    $corpusStore = $null
    if ($corpusValue.Present) {
        $corpus = $corpusValue.Value
        if (-not (Test-IsJsonObject $corpus)) {
            $Problems.Add('corpus: must be an object carrying every field. An absent corpus is left out of the file, never written as null or anything else.')
        }
        else {
            foreach ($inner in $corpus.PSObject.Properties) {
                if ($inner.Name.StartsWith('_')) { continue }
                if ($script:BlockFields['corpus'] -cnotcontains $inner.Name) { $Problems.Add("corpus.$($inner.Name): not a field this script has a rule for.") }
            }
            $missing = @($script:BlockFields['corpus'] | Where-Object { -not (Get-FieldValue $corpus $_).Present })
            foreach ($name in $missing) { $Problems.Add("corpus.${name}: missing. A corpus block is all or nothing - the loader refuses a partial one.") }

            $storeName = (Get-FieldValue $corpus 'storeDisplayName').Value
            $problem = Test-StoreName -Name $storeName -Where 'corpus.storeDisplayName'
            if ($null -ne $problem) { $Problems.Add($problem) }
            else {
                $corpusStore = $storeName
                if (-not (Test-NameIn $storeName $bystanders)) {
                    $Problems.Add("corpus.storeDisplayName: '$storeName' is not declared in bystanderStoreDisplayNames. No test writes to a corpus; undeclared, it sits inside the identity-draft grant and the identity tests would draft into the measurement corpus (Testbed/README.md section 3b).")
                }
                if (-not (Test-NameIn $storeName $expected)) {
                    $Problems.Add("corpus.storeDisplayName: '$storeName' is not in expectedStoreDisplayNames.")
                }
            }

            $corpusId = (Get-FieldValue $corpus 'corpusId').Value
            $idUsable = $false
            if (-not ($corpusId -is [string]) -or $corpusId -cnotmatch '^[A-Za-z0-9_-]+$') {
                $Problems.Add("corpus.corpusId: must be ASCII letters, digits, '-' or '_' - the id lands in every corpus subject, and CorpusPlan refuses anything else.")
            }
            else { $idUsable = $true }

            $manifest = (Get-FieldValue $corpus 'manifestPath').Value
            if (-not ($manifest -is [string]) -or $manifest -cnotmatch '^[A-Za-z]:\\') {
                $Problems.Add('corpus.manifestPath: must be an absolute path on the guest, like C:\OutlookAI-Q5\corpus-<corpusId>.jsonl.')
            }
            elseif ($idUsable) {
                $leaf = $manifest.Substring($manifest.LastIndexOf('\') + 1)
                if (-not [string]::Equals($leaf, "corpus-$corpusId.jsonl", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $Problems.Add("corpus.manifestPath: the file must be named corpus-$corpusId.jsonl, not '$leaf'. Copy-FromGuest.ps1 collects corpus-*.jsonl and nothing else, and the manifest is the only thing that can tear the corpus down (Docs/live-tier-on-the-vm.md section 3).")
                }
            }

            $seed = (Get-FieldValue $corpus 'seed').Value
            if (-not (Test-IsInteger $seed)) { $Problems.Add('corpus.seed: must be an integer, written without quotes or a decimal point.') }

            $anchorText = (Get-FieldValue $corpus 'anchorUtc').Value
            $anchor = ConvertFrom-AnchorText $anchorText
            if ($null -eq $anchor) { $Problems.Add('corpus.anchorUtc: must be yyyy-MM-dd or yyyy-MM-ddTHH:mm:ssZ - the two forms the corpus tooling reads as a UTC instant.') }

            $itemCount = (Get-FieldValue $corpus 'itemCount').Value
            if (-not (Test-IsInteger $itemCount) -or [decimal]$itemCount -lt 1 -or [decimal]$itemCount -gt [int]::MaxValue) {
                $Problems.Add('corpus.itemCount: must be a whole number above zero. The loader treats zero as "no corpus configured".')
            }

            $windows = (Get-FieldValue $corpus 'windowDays').Value
            if (-not ($windows -is [System.Array]) -or $windows.Count -eq 0) {
                $Problems.Add('corpus.windowDays: must list the measurement windows this guest asks about, in days. Empty means every window, including the one-day one, which forces a corpus rebuild every day (Docs/live-tier-on-the-vm.md section 2.10).')
            }
            else {
                $seenWindows = @()
                foreach ($day in $windows) {
                    if (-not (Test-IsInteger $day) -or [decimal]$day -lt 1 -or [decimal]$day -gt [int]::MaxValue) {
                        $Problems.Add('corpus.windowDays: every entry must be a whole number of days above zero.')
                    }
                    elseif ($seenWindows -contains [long]$day) { $Problems.Add("corpus.windowDays: lists $day twice.") }
                    else { $seenWindows += [long]$day }
                }
            }

            # One corpus, one record. The testbed record assigns each guest its id, and holds the
            # build parameters once a build has produced them. The record itself was looked up
            # above, and a guest with none has already been told so.
            if ($null -ne $entry) {
                $minimum = (Get-FieldValue $entry 'minimumItemCount').Value
                if ((Test-IsInteger $minimum) -and (Test-IsInteger $itemCount) -and [decimal]$itemCount -lt [decimal]$minimum) {
                    $Problems.Add("corpus.itemCount: $itemCount, below the $minimum corpusIdConvention requires for '$($entry.corpusId)'. The index tier's latency bounds are tests only against an index of production scale - see that entry's _minimumItemCount - so this guest's corpus is rebuilt at that size, not rendered at this one.")
                }
                if ($idUsable -and $entry.corpusId -cne $corpusId) {
                    $Problems.Add("corpus.corpusId: '$corpusId' is not the id corpusIdConvention assigns to '$VMName', which is '$($entry.corpusId)'. Two guests sharing an id overwrite each other's manifest (Testbed/README.md section 3).")
                }
                if ($null -eq $entry.seed -or $null -eq $entry.anchor -or $null -eq $entry.itemCount) {
                    $Problems.Add("corpus: corpusIdConvention still records corpus '$($entry.corpusId)' as not built - its seed, anchor and itemCount are null there. If it is built, record it there first, from the manifest's header line, so the testbed record and this guest's settings describe one corpus. If it is not, set this guest's corpus to null.")
                }
                else {
                    if ((Test-IsInteger $seed) -and (Test-IsInteger $entry.seed) -and [decimal]$seed -ne [decimal]$entry.seed) {
                        $Problems.Add("corpus.seed: $seed, where corpusIdConvention records $($entry.seed) for '$($entry.corpusId)'.")
                    }
                    if ((Test-IsInteger $itemCount) -and (Test-IsInteger $entry.itemCount) -and [decimal]$itemCount -ne [decimal]$entry.itemCount) {
                        $Problems.Add("corpus.itemCount: $itemCount, where corpusIdConvention records $($entry.itemCount) for '$($entry.corpusId)'.")
                    }
                    $recordedAnchor = ConvertFrom-AnchorText $entry.anchor
                    if ($null -eq $recordedAnchor) {
                        $Problems.Add("corpus: corpusIdConvention's anchor for '$($entry.corpusId)' is not a yyyy-MM-dd date, so nothing can be compared with it.")
                    }
                    elseif ($null -ne $anchor -and $anchor -ne $recordedAnchor) {
                        $Problems.Add("corpus.anchorUtc: $anchorText, where corpusIdConvention records $($entry.anchor) for '$($entry.corpusId)'.")
                    }
                }
            }
        }
    }

    # mailSink - all or nothing, loopback only.
    $sinkValue = Get-FieldValue $Settings 'mailSink'
    if ($sinkValue.Present) {
        $sink = $sinkValue.Value
        if (-not (Test-IsJsonObject $sink)) {
            $Problems.Add('mailSink: must be an object carrying every field. An absent sink is left out of the file, never written as null or anything else.')
        }
        else {
            foreach ($inner in $sink.PSObject.Properties) {
                if ($inner.Name.StartsWith('_')) { continue }
                if ($script:BlockFields['mailSink'] -cnotcontains $inner.Name) { $Problems.Add("mailSink.$($inner.Name): not a field this script has a rule for.") }
            }
            foreach ($name in @('submitHost', 'retrieveHost')) {
                $hostValue = (Get-FieldValue $sink $name).Value
                if (-not ($hostValue -is [string]) -or $script:LoopbackHosts -notcontains $hostValue) {
                    $Problems.Add("mailSink.${name}: must be loopback ($($script:LoopbackHosts -join ', ')). A sink reachable any other way is an open relay on a test VM (Docs/live-tier-on-the-vm.md section 2.7).")
                }
            }
            foreach ($name in @('submitPort', 'retrievePort')) {
                $port = (Get-FieldValue $sink $name).Value
                if (-not (Test-IsInteger $port) -or [decimal]$port -lt 1 -or [decimal]$port -gt 65535) {
                    $Problems.Add("mailSink.${name}: must be a port number, 1 to 65535. The loader treats zero as 'not configured' and refuses the block as partial.")
                }
            }
            $timeout = (Get-FieldValue $sink 'connectTimeoutMs').Value
            if (-not (Test-IsInteger $timeout) -or [decimal]$timeout -lt 1 -or [decimal]$timeout -gt [int]::MaxValue) {
                $Problems.Add('mailSink.connectTimeoutMs: must be a whole number of milliseconds above zero.')
            }
        }
    }

    # One spelling per store. The tier compares names ignoring case, so two spellings are one store
    # written two ways - which reads as two stores to anybody checking the file by eye.
    $spellings = @{}
    $named = New-Object System.Collections.Generic.List[object]
    if ($null -ne $hub) { $named.Add(@('testHubStoreDisplayName', $hub)) }
    foreach ($store in $expected) { $named.Add(@('expectedStoreDisplayNames', $store)) }
    foreach ($store in $indexed) { $named.Add(@('indexedStoreDisplayNames', $store)) }
    foreach ($store in $delegates) { $named.Add(@('expectedDelegateStoreDisplayNames', $store)) }
    foreach ($store in $bystanders) { $named.Add(@('bystanderStoreDisplayNames', $store)) }
    if ($null -ne $corpusStore) { $named.Add(@('corpus.storeDisplayName', $corpusStore)) }
    if ($null -ne $subjectProbeStore) { $named.Add(@('subjectOnlyProbe.storeDisplayName', $subjectProbeStore)) }
    foreach ($pair in $named) {
        $key = $pair[1].ToLowerInvariant()
        if (-not $spellings.ContainsKey($key)) { $spellings[$key] = $pair; continue }
        if ($spellings[$key][1] -cne $pair[1]) {
            $Problems.Add("$($pair[0]): spells '$($pair[1])', where $($spellings[$key][0]) spells it '$($spellings[$key][1])'. The tier treats them as one store; spell it one way.")
        }
    }
}

<#
    The stores a settings file lets the suite write to besides the hub - draft and delete only, for
    the identity tests. Reported on every render, because it is the one list a reader most needs and
    the file never states outright.
#>
function Get-IdentityGrant {
    param($Settings)
    $hub = (Get-FieldValue $Settings 'testHubStoreDisplayName').Value
    $bystanders = @((Get-FieldValue $Settings 'bystanderStoreDisplayNames').Value)
    $grant = New-Object System.Collections.Generic.List[string]
    foreach ($store in @((Get-FieldValue $Settings 'expectedStoreDisplayNames').Value)) {
        if ([string]::Equals($store, $hub, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if (Test-NameIn $store $bystanders) { continue }
        $grant.Add($store)
    }
    return , $grant.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Where the file may go.
# ---------------------------------------------------------------------------------------------

<#
    The refusal for a path with a live-fixtures directory in it, or $null. Pure text: each segment
    is compared the way Windows will resolve it - case ignored, trailing dots and spaces dropped,
    and anything after a ':' that is not a drive letter treated as the alternate-data-stream suffix
    it would be.
#>
function Get-ForbiddenSegmentProblem {
    param([string] $Resolved, [string] $Asked, [string] $How)
    $segments = $Resolved -split '[\\/]'
    for ($i = 0; $i -lt $segments.Count; $i++) {
        $segment = $segments[$i]
        $colon = $segment.IndexOf(':')
        if ($colon -ge 0 -and -not ($colon -eq 1 -and $segment.Length -eq 2)) { $segment = $segment.Substring(0, $colon) }
        $segment = $segment.TrimEnd('.', ' ')
        if ([string]::Equals($segment, $script:ForbiddenDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
            return ("REFUSING to write live-test settings to '$Asked': $How '$Resolved', which is inside a directory called " +
                "live-fixtures. On the host, McpServer/**/live-fixtures/ holds the maintainer's REAL live-test-settings.json, for his " +
                "production profile, and this script never writes, overwrites or moves anything in a live-fixtures directory, however " +
                "the path is spelled. Nothing was read and nothing was written. Render into scratch - leave -OutPath off for " +
                ".work\live-test-settings\<VMName>\ - and land the file in the GUEST's live-fixtures directory with the " +
                "Copy-ToGuest.ps1 line this script prints.")
        }
    }
    return $null
}

function Find-GitWorkTree {
    param([string] $Directory)
    $dir = $Directory
    while ($dir) {
        $dotGit = [System.IO.Path]::Combine($dir, '.git')
        if ([System.IO.Directory]::Exists($dotGit) -or [System.IO.File]::Exists($dotGit)) { return $dir }
        $parent = [System.IO.Path]::GetDirectoryName($dir)
        if (-not $parent -or $parent -eq $dir) { return $null }
        $dir = $parent
    }
    return $null
}

<#
    Where Windows says a path really is. Walks up to the deepest DIRECTORY that exists - the output
    file itself is never opened, even when it exists - asks for that directory's final path, and
    puts the rest back on. Returns $null when no part of the path exists.
#>
function Get-CanonicalPath {
    param([string] $FullPath)
    $leaf = [System.IO.Path]::GetFileName($FullPath)
    $tail = New-Object System.Collections.Generic.List[string]
    $dir = [System.IO.Path]::GetDirectoryName($FullPath)
    while ($dir -and -not [System.IO.Directory]::Exists($dir)) {
        $tail.Insert(0, [System.IO.Path]::GetFileName($dir))
        $next = [System.IO.Path]::GetDirectoryName($dir)
        if (-not $next -or $next -eq $dir) { $dir = $null; break }
        $dir = $next
    }
    if (-not $dir) { return $null }

    $final = [OutlookAI.Testbed.LiveSettingsPathProbe]::FinalPath($dir)
    if ($final.StartsWith('\\?\UNC\', [System.StringComparison]::OrdinalIgnoreCase)) { $final = '\\' + $final.Substring(8) }
    elseif ($final.StartsWith('\\?\')) { $final = $final.Substring(4) }

    $result = $final
    foreach ($segment in $tail) { $result = [System.IO.Path]::Combine($result, $segment) }
    return [System.IO.Path]::Combine($result, $leaf)
}

<#
    The refusal for an output path, or $null. The live-fixtures rule runs first and on the text
    alone, so the one path that matters most is refused before this asks the disk anything.
#>
function Get-OutputPathProblem {
    param([string] $Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return 'REFUSING: no output path was given.' }

    try {
        $provider = $null
        $drive = $null
        $unresolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path, [ref]$provider, [ref]$drive)
        if ($null -eq $provider -or $provider.Name -ne 'FileSystem') {
            return "REFUSING to write live-test settings to '$Path': it is not a file-system path."
        }
        $full = [System.IO.Path]::GetFullPath($unresolved)
    }
    catch {
        return "REFUSING to write live-test settings to '$Path': it cannot be resolved to a plain file path ($($_.Exception.Message)), so nothing can say where it would land."
    }

    $problem = Get-ForbiddenSegmentProblem -Resolved $full -Asked $Path -How 'it resolves to'
    if ($null -ne $problem) { return $problem }

    if ($full.EndsWith('\') -or $full.EndsWith('/') -or [string]::IsNullOrEmpty([System.IO.Path]::GetFileName($full))) {
        return "REFUSING to write live-test settings to '$Path': it names a directory, not a file."
    }

    try { $canonical = Get-CanonicalPath -FullPath $full }
    catch {
        return "REFUSING to write live-test settings to '$Path': Windows could not say where '$full' really is ($($_.Exception.Message)). A path this script cannot resolve is a path it cannot prove is safe."
    }
    if ($null -eq $canonical) {
        return "REFUSING to write live-test settings to '$Path': no part of '$full' exists, not even its drive."
    }

    $problem = Get-ForbiddenSegmentProblem -Resolved $canonical -Asked $Path -How 'through a junction, a symbolic link, a substituted or mapped drive or a short name, it really is'
    if ($null -ne $problem) { return $problem }

    $tree = Find-GitWorkTree -Directory ([System.IO.Path]::GetDirectoryName($canonical))
    if ($null -ne $tree) {
        $scratch = $tree.TrimEnd('\') + '\.work\'
        if (-not $canonical.StartsWith($scratch, [System.StringComparison]::OrdinalIgnoreCase)) {
            return ("REFUSING to write live-test settings to '$Path': it lands in the git working tree at '$tree', outside that " +
                "tree's gitignored scratch. Inside a working tree this script writes only under .work\ - so a render can never " +
                "replace a tracked file, in this checkout or in another one. Leave -OutPath off, or point it under '$scratch'.")
        }
    }
    return $null
}

<#
    Whether an existing file may be replaced, or $null. Never reads it: the questions are what it
    is, not what it holds.
#>
function Get-OverwriteProblem {
    param([string] $FullPath, [bool] $InsideWorkTree, [bool] $Force)
    if ([System.IO.Directory]::Exists($FullPath)) {
        return "REFUSING to write live-test settings to '$FullPath': a directory is already there."
    }
    if (-not [System.IO.File]::Exists($FullPath)) { return $null }

    $attributes = [System.IO.File]::GetAttributes($FullPath)
    if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        return "REFUSING to replace '$FullPath': it is a link. Replacing it would unlink it rather than write through it, but a link here is almost certainly a mistake worth looking at first. Remove it yourself if it is not."
    }
    $links = [OutlookAI.Testbed.LiveSettingsPathProbe]::LinkCount($FullPath)
    if ($links -gt 1) {
        return "REFUSING to replace '$FullPath': it is one of $links hard links to the same file. Replacing it would detach this name rather than change the others, but a hard link here is almost certainly a mistake worth looking at first. Remove it yourself if it is not."
    }
    if (-not $InsideWorkTree -and -not $Force) {
        return "REFUSING to replace '$FullPath', which already exists: it is outside every git working tree, so this script cannot tell a previous render from anything else. Pass -Force if replacing it is what you mean."
    }
    return $null
}

function Format-CopyToGuestLine {
    param([string] $VMName, [string] $RenderedPath)
    $destination = $script:GuestSourceRoot.TrimEnd('\') + '\' + $script:GuestSettingsRelative
    $pathArgument = $RenderedPath
    if ($pathArgument.Contains(' ')) { $pathArgument = '"' + $pathArgument + '"' }
    return "pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName $VMName -Path $pathArgument -Destination $destination"
}

<#
    Writes a NEW file beside the target and renames it over the target's name. A hard link or a
    symbolic link at the destination is therefore unlinked, never written through - though both are
    refused before this runs. The bytes are ASCII: the writer above escapes everything else.
#>
function Write-RenderedFile {
    param([string] $FullPath, [string] $Text)
    if ($Text -match '[^\x00-\x7F]') { throw 'Internal error: the rendered text is not pure ASCII, so it was not written.' }
    $directory = [System.IO.Path]::GetDirectoryName($FullPath)
    if (-not [System.IO.Directory]::Exists($directory)) { [void][System.IO.Directory]::CreateDirectory($directory) }
    $temporary = [System.IO.Path]::Combine($directory, '.' + [System.IO.Path]::GetFileName($FullPath) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($Text)
    $stream = New-Object System.IO.FileStream($temporary, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    }
    finally { $stream.Dispose() }
    try {
        if ([System.IO.File]::Exists($FullPath)) { [System.IO.File]::Delete($FullPath) }
        [System.IO.File]::Move($temporary, $FullPath)
    }
    catch {
        if ([System.IO.File]::Exists($temporary)) { [System.IO.File]::Delete($temporary) }
        throw
    }
}

<#
    A path as a reader of the rendered file should see it: repository-relative with forward slashes
    when it is inside $Root, in full when it is not.
#>
function Get-DisplayPath {
    param([string] $Path, [string] $Root)
    $full = [System.IO.Path]::GetFullPath($Path)
    $prefix = $Root.TrimEnd('\') + '\'
    if ($full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($prefix.Length).Replace('\', '/')
    }
    return $full
}

function Get-GitProvenance {
    param([string] $Root, [string[]] $Inputs)
    # Provenance only: a render is right or wrong by its checks, not by this line. So git missing or
    # failing is reported as 'unknown' rather than stopping anything.
    $inside = @($Inputs | Where-Object { -not [System.IO.Path]::IsPathRooted($_) })
    $outside = @($Inputs | Where-Object { [System.IO.Path]::IsPathRooted($_) })

    # WINDOWS POWERSHELL 5.1: with the preference at 'Stop', the FIRST line a native program writes
    # to stderr is a terminating NativeCommandError - whether it is sent to $null, merged with 2>&1 or
    # left alone to a file - and git writes warnings there. PowerShell 7 does not do this. So each git
    # call runs under 'Continue', restored in a finally, and success is judged by $LASTEXITCODE alone.
    # With 'Continue' the redirected lines arrive as ErrorRecord objects rather than strings; 2>$null
    # discards them, and only the success stream is read.
    $text = $null
    try {
        $saved = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { $head = @(& git -C $Root rev-parse HEAD 2>$null) }
        finally { $ErrorActionPreference = $saved }
        if ($LASTEXITCODE -ne 0 -or $head.Count -eq 0) { $text = 'an unknown repository commit (git did not answer)' }
        else {
            $text = "repository commit $(([string]$head[0]).Trim())"
            if ($inside.Count -gt 0) {
                $saved = $ErrorActionPreference
                $ErrorActionPreference = 'Continue'
                try { $dirty = @(& git -C $Root status --porcelain -- $inside 2>$null) }
                finally { $ErrorActionPreference = $saved }
                if ($LASTEXITCODE -eq 0 -and $dirty.Count -gt 0) {
                    $text += " WITH UNCOMMITTED CHANGES to $($inside -join ' or '), so that commit does not fully describe these values"
                }
            }
        }
    }
    catch { $text = 'an unknown repository commit (git is not available)' }
    if ($outside.Count -gt 0) { $text += "; $($outside -join ' and ') is outside the repository, so no commit describes it" }
    return $text
}

function Read-TextFile {
    param([string] $Path, [string] $What)
    if (-not [System.IO.File]::Exists($Path)) { throw "$What is not at $Path." }
    return [System.IO.File]::ReadAllText($Path)
}

# =============================================================================================
# SELF-TEST. Reads the committed template and testbed.json; writes only a scratch directory under
# .work\ and removes it. No guest, no COM, no registry, no real live-fixtures directory.
# =============================================================================================

function ConvertTo-SelfTestText {
    param($Value)
    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Array]) {
        $parts = @()
        foreach ($item in $Value) { $parts += (ConvertTo-SelfTestText $item) }
        return '[' + ($parts -join ', ') + ']'
    }
    return [string] $Value
}

function Invoke-SelfTest {
    $script:SelfTestChecks = 0
    $script:SelfTestFailures = @()

    # Assertions compare with -ceq and search with String.Contains. Never -like: see Testbed/README.md
    # section 4b for the false greens a bracket in a -like pattern produced in a sibling script.
    function Test-Case {
        param([string] $What, $Expected, $Actual)
        $script:SelfTestChecks++
        $expectedText = ConvertTo-SelfTestText $Expected
        $actualText = ConvertTo-SelfTestText $Actual
        if ($expectedText -ceq $actualText) {
            Write-Host ("  OK   {0}" -f $What)
        }
        else {
            $script:SelfTestFailures += "$What : expected $expectedText, got $actualText"
            Write-Host ("  FAIL {0} - expected {1}, got {2}" -f $What, $expectedText, $actualText)
        }
    }

    function Test-Refused {
        param([string] $What, $Message, [string] $Needle)
        # Case ignored for the needle only: some needles are paths, and Windows reports a path in the
        # case it is stored in, not the case it was typed in.
        $refused = ($null -ne $Message -and ([string]$Message).StartsWith('REFUSING') -and
            ([string]$Message).IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
        if (-not $refused -and $null -ne $Message) { Write-Host ("       (message was: {0})" -f $Message) }
        Test-Case $What $true $refused
    }

    function Test-HasProblem {
        param([string] $What, [string[]] $Problems, [string] $Needle)
        $hit = $false
        foreach ($problem in $Problems) { if ($problem.Contains($Needle)) { $hit = $true } }
        if (-not $hit) { Write-Host ("       (problems were: {0})" -f ($Problems -join ' | ')) }
        Test-Case $What $true $hit
    }

    $repoFull = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $vm = 'OutlookAI-Synthetic'

    $guestText = @'
{
  "_note": "comment keys are ignored at every level",
  "machineProfile": "Portable",
  "testHubStoreDisplayName": "hub@render.invalid",
  "expectedStoreDisplayNames": [ "hub@render.invalid", "Synthetic Corpus", "bystander@render.invalid", "identity@render.invalid" ],
  "indexedStoreDisplayNames": [ "hub@render.invalid", "bystander@render.invalid", "Synthetic Corpus" ],
  "expectedDelegateStoreDisplayNames": [],
  "bystanderStoreDisplayNames": [ "bystander@render.invalid", "Synthetic Corpus" ],
  "probeTerm": "invoice",
  "subjectOnlyProbe": {
    "_note": "the hub population's notices folder",
    "storeDisplayName": "hub@render.invalid",
    "folderPath": "Inbox/OutlookAI-Corpus-Folder-Notices",
    "subjectTerm": "bulletin",
    "senderFragment": "noticebot"
  },
  "corpus": {
    "_note": "ignored too",
    "storeDisplayName": "Synthetic Corpus",
    "manifestPath": "C:\\OutlookAI-Q5\\corpus-vm-synthetic.jsonl",
    "corpusId": "vm-synthetic",
    "seed": 4242,
    "anchorUtc": "2026-08-01T00:00:00Z",
    "itemCount": 1000,
    "windowDays": [ 7, 30, 60 ]
  },
  "mailSink": {
    "submitHost": "127.0.0.1",
    "submitPort": 2525,
    "retrieveHost": "127.0.0.1",
    "retrievePort": 1110,
    "connectTimeoutMs": 1500
  }
}
'@
    $conventionText = '{ "assigned": [ { "corpusId": "vm-synthetic", "guest": "OutlookAI-Synthetic", "indexed": true, "seed": 4242, "anchor": "2026-08-01", "itemCount": 1000, "minimumItemCount": 1000 }, { "corpusId": "vm-other", "guest": "OutlookAI-Other", "indexed": false, "seed": null, "anchor": null, "itemCount": null }, { "corpusId": "vm-unknown", "guest": "OutlookAI-Unknown", "seed": null, "anchor": null, "itemCount": null } ] }'
    $assigned = (ConvertFrom-JsonText -Text $conventionText -What 'the synthetic convention').assigned

    function New-Guest { return (ConvertFrom-JsonText -Text $guestText -What 'the synthetic guest') }

    Write-Host 'New-LiveTestSettings self-test. No guest, no COM, no registry; the only thing written is a'
    Write-Host 'scratch directory under .work\, removed at the end.'

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== the live-fixtures refusal, on the text alone =='
    $real = Join-Path $repoFull 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json'
    Test-Refused 'the settings file''s own path is refused' (Get-OutputPathProblem -Path $real) 'inside a directory called live-fixtures'
    Test-Refused 'and the refusal says nothing was read or written' (Get-OutputPathProblem -Path $real) 'Nothing was read and nothing was written'
    Test-Refused 'forward slashes are refused' (Get-OutputPathProblem -Path ($real -replace '\\', '/')) 'inside a directory called live-fixtures'
    Test-Refused 'upper case is refused' (Get-OutputPathProblem -Path ($real -replace 'live-fixtures', 'LIVE-FIXTURES')) 'inside a directory called live-fixtures'
    Test-Refused 'a .. walk back into it is refused' (Get-OutputPathProblem -Path (Join-Path $repoFull '.work\..\McpServer\OutlookAI.McpServer.Tests\live-fixtures\x.json')) 'inside a directory called live-fixtures'
    Test-Refused 'a trailing dot on the directory is refused' (Get-OutputPathProblem -Path ($real -replace 'live-fixtures', 'live-fixtures.')) 'inside a directory called live-fixtures'
    Test-Refused 'a trailing space on the directory is refused' (Get-OutputPathProblem -Path ($real -replace 'live-fixtures', 'live-fixtures ')) 'inside a directory called live-fixtures'
    Test-Refused 'a subdirectory of it is refused' (Get-OutputPathProblem -Path (Join-Path $repoFull 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\sub\x.json')) 'inside a directory called live-fixtures'
    Test-Refused 'a live-fixtures directory anywhere is refused, not only under McpServer' (Get-OutputPathProblem -Path 'C:\elsewhere\live-fixtures\x.json') 'inside a directory called live-fixtures'
    Test-Refused 'an admin-share spelling is refused before any network access' (Get-OutputPathProblem -Path ('\\localhost\C$' + $real.Substring(2))) 'inside a directory called live-fixtures'
    Test-Refused 'a stream suffix on the directory is refused' (Get-OutputPathProblem -Path (Join-Path $repoFull 'McpServer\OutlookAI.McpServer.Tests\live-fixtures:x')) 'REFUSING'
    Test-Refused 'a \\?\ spelling is refused' (Get-OutputPathProblem -Path ('\\?\' + $real)) 'REFUSING'
    Test-Refused 'a registry path is refused' (Get-OutputPathProblem -Path 'HKCU:\Software\OutlookAI\x.json') 'not a file-system path'
    Test-Refused 'an empty path is refused' (Get-OutputPathProblem -Path ' ') 'no output path'

    $scratchRoot = Join-Path $repoFull '.work'
    if (-not [System.IO.Directory]::Exists($scratchRoot)) { [void][System.IO.Directory]::CreateDirectory($scratchRoot) }
    Push-Location -LiteralPath $scratchRoot
    try {
        Test-Refused 'a relative .. path is resolved from where you stand, and refused' (Get-OutputPathProblem -Path '..\McpServer\OutlookAI.McpServer.Tests\live-fixtures\x.json') 'inside a directory called live-fixtures'
    }
    finally { Pop-Location }

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== where a file may land =='
    $default = Join-Path $repoFull '.work\live-test-settings\OutlookAI-Indexed\live-test-settings.json'
    Test-Case 'the default under .work\ is allowed' '<null>' (Get-OutputPathProblem -Path $default)
    Test-Refused 'a tracked directory of this working tree is refused' (Get-OutputPathProblem -Path (Join-Path $repoFull 'Testbed\live-test-settings.json')) 'outside that tree''s gitignored scratch'
    Test-Refused 'the repository root is refused' (Get-OutputPathProblem -Path (Join-Path $repoFull 'rendered.json')) 'gitignored scratch'
    Test-Refused 'a path naming a directory is refused' (Get-OutputPathProblem -Path ($scratchRoot + '\')) 'names a directory'
    $outside = [System.IO.Path]::Combine([System.IO.Path]::GetPathRoot($repoFull), 'new-live-test-settings-selftest-' + [guid]::NewGuid().ToString('N'), 'x.json')
    Test-Case 'a path outside every working tree is allowed (nothing there is created)' '<null>' (Get-OutputPathProblem -Path $outside)

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== links, short names and existing files - on a synthetic directory in .work\ =='
    $sandbox = Join-Path $scratchRoot ('new-live-test-settings-selftest-' + [guid]::NewGuid().ToString('N'))
    $fixtures = Join-Path $sandbox 'live-fixtures'
    $junction = Join-Path $sandbox 'through-a-junction'
    try {
        [void][System.IO.Directory]::CreateDirectory($fixtures)
        [void](New-Item -ItemType Junction -Path $junction -Target $fixtures)
        Test-Refused 'a junction into a live-fixtures directory is refused' (Get-OutputPathProblem -Path (Join-Path $junction 'x.json')) 'really is'
        Test-Refused 'and the refusal names the directory it really leads to' (Get-OutputPathProblem -Path (Join-Path $junction 'x.json')) $fixtures
        Test-Refused 'a path that does not exist yet, below the junction, is refused too' (Get-OutputPathProblem -Path (Join-Path $junction 'new\deeper\x.json')) 'really is'

        $short = [OutlookAI.Testbed.LiveSettingsPathProbe]::ShortPath($fixtures)
        if ($null -ne $short -and -not ($short -split '\\' | Where-Object { $_ -ieq 'live-fixtures' })) {
            # Either layer may catch this one: .NET Core's GetFullPath expands a short name itself, so on
            # PowerShell 7 the text check refuses it; on Windows PowerShell only the final-path check can.
            Test-Refused 'an 8.3 short name for it is refused' (Get-OutputPathProblem -Path (Join-Path $short 'x.json')) 'inside a directory called live-fixtures'
        }
        else {
            Write-Host '  SKIP an 8.3 short name for it - this volume keeps no short name for that directory, so there is none to try.'
        }

        $plain = Join-Path $sandbox 'plain'
        [void][System.IO.Directory]::CreateDirectory($plain)
        Test-Case 'an ordinary directory in scratch is allowed' '<null>' (Get-OutputPathProblem -Path (Join-Path $plain 'x.json'))
        $canonical = Get-CanonicalPath -FullPath (Join-Path $plain 'x.json')
        Test-Case 'and Windows reports it where it is' $true ([string]::Equals($canonical, (Join-Path $plain 'x.json'), [System.StringComparison]::OrdinalIgnoreCase))

        $existing = Join-Path $plain 'existing.json'
        [System.IO.File]::WriteAllText($existing, '{}')
        Test-Case 'an earlier render inside a working tree may be replaced' '<null>' (Get-OverwriteProblem -FullPath $existing -InsideWorkTree $true -Force $false)
        Test-Refused 'an existing file outside every working tree needs -Force' (Get-OverwriteProblem -FullPath $existing -InsideWorkTree $false -Force $false) '-Force'
        Test-Case 'and -Force allows it' '<null>' (Get-OverwriteProblem -FullPath $existing -InsideWorkTree $false -Force $true)
        Test-Refused 'a directory in the way is refused' (Get-OverwriteProblem -FullPath $plain -InsideWorkTree $true -Force $true) 'a directory is already there'

        $second = Join-Path $plain 'second-name.json'
        [void](New-Item -ItemType HardLink -Path $second -Target $existing)
        Test-Case 'Windows counts two names for a hard-linked file' '2' ([string][OutlookAI.Testbed.LiveSettingsPathProbe]::LinkCount($existing))
        Test-Refused 'a hard link at the destination is refused even with -Force' (Get-OverwriteProblem -FullPath $second -InsideWorkTree $true -Force $true) 'hard links'

        $written = Join-Path $plain 'written.json'
        Write-RenderedFile -FullPath $written -Text "{`n}`n"
        Test-Case 'a new file is written' "{`n}`n" ([System.IO.File]::ReadAllText($written))
        Write-RenderedFile -FullPath $written -Text "{ `"a`": 1 }`n"
        Test-Case 'and replaced through the rename' "{ `"a`": 1 }`n" ([System.IO.File]::ReadAllText($written))
        Test-Case 'with no temporary file left behind' 0 ([System.IO.Directory]::GetFiles($plain, '*.tmp').Count)
        $threw = $false
        try { Write-RenderedFile -FullPath $written -Text ('{ "a": "caf' + [char]0x00E9 + '" }') } catch { $threw = $true }
        Test-Case 'text that is not pure ASCII is never written' $true $threw
        Test-Case 'and the file it would have replaced is untouched' "{ `"a`": 1 }`n" ([System.IO.File]::ReadAllText($written))
    }
    finally {
        # The junction first, as a link: removing it recursively could follow it.
        if ([System.IO.Directory]::Exists($junction)) { [System.IO.Directory]::Delete($junction, $false) }
        if ([System.IO.Directory]::Exists($sandbox)) { Remove-Item -LiteralPath $sandbox -Recurse -Force }
    }
    Test-Case 'the scratch directory is gone again' $false ([System.IO.Directory]::Exists($sandbox))

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== the committed template =='
    $templateFile = $TemplatePath
    if (-not $templateFile) { $templateFile = Join-Path $repoFull 'Testbed\live-test-settings.template.json' }
    $template = ConvertFrom-JsonText -Text (Read-TextFile $templateFile 'The template') -What 'the template'
    $shape = Get-TemplateFields -Template $template
    Test-Case 'every value in it is the token for its own path' '' ($shape.Problems -join ' | ')
    Test-Case 'it names exactly the fields this script has rules for' '' ((Get-TemplateDriftProblems -Fields $shape.Fields) -join ' | ')
    Test-Case 'in order' 'machineProfile,testHubStoreDisplayName,expectedStoreDisplayNames,indexedStoreDisplayNames,expectedDelegateStoreDisplayNames,bystanderStoreDisplayNames,probeTerm,subjectOnlyProbe.storeDisplayName,subjectOnlyProbe.folderPath,subjectOnlyProbe.subjectTerm,subjectOnlyProbe.senderFragment,corpus.storeDisplayName,corpus.manifestPath,corpus.corpusId,corpus.seed,corpus.anchorUtc,corpus.itemCount,corpus.windowDays,mailSink.submitHost,mailSink.submitPort,mailSink.retrieveHost,mailSink.retrievePort,mailSink.connectTimeoutMs' ((@($shape.Fields | ForEach-Object { $_.Path })) -join ',')

    $bad = ConvertFrom-JsonText -Text '{ "machineProfile": "Portable", "corpus": { "seed": "{{corpus.itemCount}}", "deep": { "x": "{{corpus.deep.x}}" } } }' -What 'a bad template'
    $badShape = Get-TemplateFields -Template $bad
    Test-HasProblem 'a literal value in a template is refused' $badShape.Problems "'machineProfile' is not the token"
    Test-HasProblem 'a token in the wrong slot is refused' $badShape.Problems "'corpus.seed' is not the token"
    Test-HasProblem 'a block inside a block is refused' $badShape.Problems 'nests a block inside a block'
    Test-HasProblem 'and a template that lost a field is noticed' (Get-TemplateDriftProblems -Fields $badShape.Fields) "has no 'testHubStoreDisplayName'"

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== matching a guest''s values against the template =='
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest (New-Guest) -Where 'the synthetic guest'
    Test-Case 'a complete guest has no problems' '' ($resolved.Problems -join ' | ')
    Test-Case 'and no block left out' 0 $resolved.Omitted.Count
    Test-Case 'and every field has a value' 23 $resolved.ByPath.Count

    $guest = New-Guest
    $guest.testHubStoreDisplayName = '<FILL: from the guest>'
    $guest.expectedStoreDisplayNames = @('hub@render.invalid', '<FILL: more>')
    $guest.corpus.seed = '<FILL: an integer>'
    $guest.mailSink = '<FILL: a block or null>'
    $guest.PSObject.Properties.Remove('bystanderStoreDisplayNames')
    $guest | Add-Member -NotePropertyName 'bystanderStoreDisplayName' -NotePropertyValue @('Synthetic Bystander')
    $guest.corpus.PSObject.Properties.Remove('windowDays')
    $guest.corpus | Add-Member -NotePropertyName 'windowsDays' -NotePropertyValue @(7)
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guest -Where 'the synthetic guest'
    Test-HasProblem 'a placeholder hub is named' $resolved.Problems 'testHubStoreDisplayName: still a placeholder'
    Test-HasProblem 'a placeholder inside a list is named' $resolved.Problems 'expectedStoreDisplayNames: still a placeholder'
    Test-HasProblem 'a placeholder inside a block is named by its path' $resolved.Problems 'corpus.seed: still a placeholder'
    Test-HasProblem 'a placeholder standing for a whole block is named' $resolved.Problems 'mailSink: still a placeholder'
    Test-HasProblem 'a missing field is named' $resolved.Problems 'bystanderStoreDisplayNames: missing'
    Test-HasProblem 'a misspelt field is named, not ignored' $resolved.Problems 'bystanderStoreDisplayName: not a field of the template'
    Test-HasProblem 'a field missing from a block is named' $resolved.Problems 'corpus.windowDays: missing'
    Test-HasProblem 'a misspelt field inside a block is named' $resolved.Problems 'corpus.windowsDays: not a field'
    Test-Case 'and every one of them is reported at once' 8 $resolved.Problems.Count

    $guest = New-Guest
    $guest.PSObject.Properties.Remove('machineProfile')
    $guest | Add-Member -NotePropertyName 'MachineProfile' -NotePropertyValue 'Portable'
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guest -Where 'the synthetic guest'
    Test-HasProblem 'a field spelt in the wrong case is not accepted as the field' $resolved.Problems 'machineProfile: missing'

    $guest = New-Guest
    $guest.corpus = $null
    $guest.mailSink = $null
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guest -Where 'the synthetic guest'
    Test-Case 'blocks declared null are no problem' '' ($resolved.Problems -join ' | ')
    Test-Case 'and are left out' 'corpus,mailSink' ($resolved.Omitted -join ',')

    $guest = New-Guest
    $guest.corpus = 42
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guest -Where 'the synthetic guest'
    Test-HasProblem 'a block that is neither an object nor null is refused' $resolved.Problems 'corpus: must be an object carrying every field, or null'

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== the committed testbed.json =='
    $testbedFile = $TestbedJsonPath
    if (-not $testbedFile) { $testbedFile = Join-Path $repoFull 'Testbed\testbed.json' }
    $testbed = ConvertFrom-JsonText -Text (Read-TextFile $testbedFile 'testbed.json') -What 'testbed.json'
    $section = Get-ExactProperty $testbed 'liveTestSettings'
    Test-Case 'it has a liveTestSettings section' $true ($null -ne $section)
    if ($null -ne $section) {
        foreach ($entry in $section.Value.PSObject.Properties) {
            if ($entry.Name.StartsWith('_')) { continue }
            $check = Resolve-GuestValues -Fields $shape.Fields -Guest $entry.Value -Where $entry.Name
            $structural = @($check.Problems | Where-Object { -not $_.Contains('still a placeholder') })
            Test-Case "$($entry.Name) names exactly the template's fields" '' ($structural -join ' | ')
            $waiting = @($check.Problems | Where-Object { $_.Contains('still a placeholder') }).Count
            if ($waiting -gt 0) { Write-Host ("       ({0} still waits on {1} value(s) from the guest - a render refuses it, naming each)" -f $entry.Name, $waiting) }
            else { Write-Host ("       ({0} is filled in)" -f $entry.Name) }
        }
    }

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== writing JSON =='
    Test-Case 'a quote and a backslash are escaped' '"a\"b\\c"' (ConvertTo-JsonString -Text 'a"b\c')
    Test-Case 'a newline is escaped' '"a\u000ab"' (ConvertTo-JsonString -Text "a`nb")
    Test-Case 'anything outside ASCII is escaped' '"caf\u00e9"' (ConvertTo-JsonString -Text ('caf' + [char]0x00E9))
    Test-Case 'integers are written plainly' '7777' (ConvertTo-JsonLiteral -Value 7777)
    Test-Case 'a large one too' '9007199254740993' (ConvertTo-JsonLiteral -Value ([long]9007199254740993))
    Test-Case 'null is null' 'null' (ConvertTo-JsonLiteral -Value $null)
    Test-Case 'true is true' 'true' (ConvertTo-JsonLiteral -Value $true)
    Test-Case 'an empty list stays a list' '[]' (ConvertTo-JsonLiteral -Value @())
    Test-Case 'a ONE-element list stays a list' '[ "only" ]' (ConvertTo-JsonLiteral -Value @('only'))
    Test-Case 'a list of numbers is written on one line' '[ 7, 30, 60 ]' (ConvertTo-JsonLiteral -Value @(7, 30, 60))
    $nested = [ordered]@{ a = 1; b = [ordered]@{ c = @('x') } }
    Test-Case 'a nested object is laid out and indented' "{`n  `"a`": 1,`n  `"b`": {`n    `"c`": [ `"x`" ]`n  }`n}" (ConvertTo-JsonLiteral -Value $nested)
    $threw = $false
    try { [void](ConvertTo-JsonLiteral -Value (New-Object System.Object)) } catch { $threw = $true }
    Test-Case 'a type it does not know is refused, not stringified' $true $threw

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== dates, which PowerShell 7 would otherwise rewrite =='
    Test-Case 'a UTC instant goes back to exactly the text it came from' '2026-08-19T00:00:00Z' (Repair-JsonDates -Node ([datetime]::SpecifyKind([datetime]'2026-08-19', [System.DateTimeKind]::Utc)) -Path 'x')
    $threw = $false
    try { [void](Repair-JsonDates -Node ([datetime]::SpecifyKind([datetime]'2026-08-19', [System.DateTimeKind]::Local)) -Path 'corpus.anchorUtc') } catch { $threw = $_.Exception.Message.Contains('corpus.anchorUtc') }
    Test-Case 'a local time is refused, naming the field' $true $threw
    Test-Case 'a string is left alone' '2026-08-19T00:00:00Z' (Repair-JsonDates -Node '2026-08-19T00:00:00Z')
    $parsedDate = ConvertFrom-JsonText -Text '{ "anchorUtc": "2026-08-19T00:00:00Z" }' -What 'a date'
    Test-Case 'and on this PowerShell an anchor reads back as the same string' '2026-08-19T00:00:00Z|System.String' ('{0}|{1}' -f $parsedDate.anchorUtc, $parsedDate.anchorUtc.GetType().FullName)

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== a whole render, in memory =='
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest (New-Guest) -Where 'the synthetic guest'
    $document = New-RenderedDocument -Fields $shape.Fields -Resolved $resolved -VMName $vm -Provenance 'a self-test' -RenderedAtUtc ([datetime]::UtcNow)
    $text = (ConvertTo-JsonLiteral -Value $document) + "`n"
    Test-Case 'no token is left in it' 0 ([regex]::Matches($text, $script:TokenPattern)).Count
    Test-Case 'it is pure ASCII' $false ($text -match '[^\x00-\x7F]')
    $parsed = ConvertFrom-JsonText -Text $text -What 'the rendered text'
    Test-Case 'it parses' $true (Test-IsJsonObject $parsed)
    Test-Case 'the hub arrives' 'hub@render.invalid' $parsed.testHubStoreDisplayName
    Test-Case 'the store list arrives whole, in order' 'hub@render.invalid|Synthetic Corpus|bystander@render.invalid|identity@render.invalid' ($parsed.expectedStoreDisplayNames -join '|')
    Test-Case 'the indexed list arrives whole, in order' 'hub@render.invalid|bystander@render.invalid|Synthetic Corpus' ($parsed.indexedStoreDisplayNames -join '|')
    Test-Case 'the probe term and the subject-only probe arrive' 'invoice|bulletin|noticebot' ('{0}|{1}|{2}' -f $parsed.probeTerm, $parsed.subjectOnlyProbe.subjectTerm, $parsed.subjectOnlyProbe.senderFragment)
    Test-Case 'an empty list arrives as a list' $true ($parsed.expectedDelegateStoreDisplayNames -is [System.Array])
    Test-Case 'the corpus arrives with its numbers as numbers' '4242|1000|7,30,60' ('{0}|{1}|{2}' -f $parsed.corpus.seed, $parsed.corpus.itemCount, ($parsed.corpus.windowDays -join ','))
    Test-Case 'and its anchor as the string it was' '2026-08-01T00:00:00Z' $parsed.corpus.anchorUtc
    Test-Case 'the sink arrives' '127.0.0.1:2525' ('{0}:{1}' -f $parsed.mailSink.submitHost, $parsed.mailSink.submitPort)
    Test-Case 'the provenance note says where it came from' $true ($parsed._rendered.Contains('liveTestSettings.OutlookAI-Synthetic'))
    $problems = New-Object System.Collections.Generic.List[string]
    Add-SettingsProblems -Settings $parsed -VMName $vm -Assigned $assigned -Problems $problems
    Test-Case 'and it breaks no rule' '' ($problems -join ' | ')
    Test-Case 'the identity grant is the one store left out of the bystanders' 'identity@render.invalid' ((Get-IdentityGrant -Settings $parsed) -join ',')

    $guest = New-Guest
    $guest.corpus = $null
    $guest.mailSink = $null
    # A corpus that is not built is not indexed either, so it leaves the indexed list with its block.
    $guest.indexedStoreDisplayNames = @('hub@render.invalid', 'bystander@render.invalid')
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guest -Where 'the synthetic guest'
    $text = (ConvertTo-JsonLiteral -Value (New-RenderedDocument -Fields $shape.Fields -Resolved $resolved -VMName $vm -Provenance 'a self-test' -RenderedAtUtc ([datetime]::UtcNow))) + "`n"
    $parsed = ConvertFrom-JsonText -Text $text -What 'the rendered text'
    Test-Case 'a block declared null is absent from the file, not null in it' 'False|False' ('{0}|{1}' -f ($null -ne (Get-ExactProperty $parsed 'corpus')), ($null -ne (Get-ExactProperty $parsed 'mailSink')))
    Test-Case 'and a note says why the sink is absent' $true ($null -ne (Get-ExactProperty $parsed '_mailSink'))
    $problems = New-Object System.Collections.Generic.List[string]
    Add-SettingsProblems -Settings $parsed -VMName $vm -Assigned $assigned -Problems $problems
    Test-Case 'a guest with neither block breaks no rule' '' ($problems -join ' | ')

    # The unindexed guest's shape: nothing indexed, no probe term, no subject-only probe - rendered,
    # with the probe block left OUT, and accepted for a guest the record calls unindexed.
    $guest = New-Guest
    $guest.corpus = $null
    $guest.indexedStoreDisplayNames = @()
    $guest.probeTerm = ''
    $guest.subjectOnlyProbe = $null
    $resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guest -Where 'the synthetic guest'
    $text = (ConvertTo-JsonLiteral -Value (New-RenderedDocument -Fields $shape.Fields -Resolved $resolved -VMName 'OutlookAI-Other' -Provenance 'a self-test' -RenderedAtUtc ([datetime]::UtcNow))) + "`n"
    $parsed = ConvertFrom-JsonText -Text $text -What 'the rendered text'
    Test-Case 'an unindexed guest''s probe block is absent from the file' $false ($null -ne (Get-ExactProperty $parsed 'subjectOnlyProbe'))
    Test-Case 'and a note says why' $true ($null -ne (Get-ExactProperty $parsed '_subjectOnlyProbe'))
    Test-Case 'and its indexed list is an empty LIST' $true ($parsed.indexedStoreDisplayNames -is [System.Array])
    $problems = New-Object System.Collections.Generic.List[string]
    Add-SettingsProblems -Settings $parsed -VMName 'OutlookAI-Other' -Assigned $assigned -Problems $problems
    Test-Case 'the unindexed guest''s shape breaks no rule' '' ($problems -join ' | ')

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== every rule fires =='
    function Get-RuleProblems {
        param([scriptblock] $Change, [string] $ForVm = $vm)
        $settings = New-Guest
        & $Change $settings
        $found = New-Object System.Collections.Generic.List[string]
        Add-SettingsProblems -Settings $settings -VMName $ForVm -Assigned $assigned -Problems $found
        return , $found.ToArray()
    }

    Test-HasProblem 'Production is refused, and says why' (Get-RuleProblems { param($s) $s.machineProfile = 'Production' }) 'A test guest is Portable'
    Test-HasProblem 'a profile in the wrong case is refused' (Get-RuleProblems { param($s) $s.machineProfile = 'portable' }) "must be the string 'Portable'"
    Test-HasProblem 'a numeric profile is refused' (Get-RuleProblems { param($s) $s.machineProfile = 1 }) "must be the string 'Portable'"
    Test-HasProblem 'a missing profile is refused' (Get-RuleProblems { param($s) $s.PSObject.Properties.Remove('machineProfile') }) 'machineProfile: missing'
    Test-HasProblem 'a blank hub is refused' (Get-RuleProblems { param($s) $s.testHubStoreDisplayName = ' ' }) 'empty store name'
    Test-HasProblem 'a hub that is not an address is refused' (Get-RuleProblems { param($s) $s.testHubStoreDisplayName = 'Hub Store'; $s.expectedStoreDisplayNames = @('Hub Store', 'Synthetic Corpus', 'Synthetic Bystander') }) 'not shaped like an SMTP address'
    Test-HasProblem 'a hub with a trailing space is refused' (Get-RuleProblems { param($s) $s.testHubStoreDisplayName = 'hub@render.invalid ' }) 'leading or trailing whitespace'
    Test-HasProblem 'a hub missing from the census list is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames = @('Synthetic Corpus', 'Synthetic Bystander') }) 'does not name the hub'
    Test-HasProblem 'an empty census list is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames = @() }) 'expectedStoreDisplayNames: is empty'
    Test-HasProblem 'a census list that is not a list is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames = 'hub@render.invalid' }) 'must be a list'
    Test-HasProblem 'a store named twice is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames = @('hub@render.invalid', 'Synthetic Corpus', 'synthetic corpus', 'Synthetic Bystander') }) 'twice'
    Test-HasProblem 'a store name with a slash is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames = @('hub@render.invalid', 'Synthetic Corpus', 'Synthetic Bystander', 'A/B') }) 'contains a slash'
    Test-HasProblem 'a bystander missing from the census list is refused - the guest convention' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames = @('hub@render.invalid', 'Synthetic Corpus') }) "'bystander@render.invalid' is not ALSO in expectedStoreDisplayNames"
    Test-HasProblem 'the hub declared a bystander is refused' (Get-RuleProblems { param($s) $s.bystanderStoreDisplayNames = @('Synthetic Bystander', 'Synthetic Corpus', 'hub@render.invalid') }) 'names the hub'
    Test-HasProblem 'no bystander at all is refused, in the tier''s own words' (Get-RuleProblems { param($s) $s.bystanderStoreDisplayNames = @(); $s.corpus = $null }) 'NO STORE THIS CENSUS WATCHES CAN PRODUCE A FAILURE'
    Test-HasProblem 'a delegate that is also in the census list is refused' (Get-RuleProblems { param($s) $s.expectedDelegateStoreDisplayNames = @('identity@render.invalid') }) 'inside the identity-draft grant'
    Test-HasProblem 'the hub declared a delegate is refused' (Get-RuleProblems { param($s) $s.expectedDelegateStoreDisplayNames = @('hub@render.invalid') }) 'names the hub'
    Test-HasProblem 'a corpus store that is not a bystander is refused' (Get-RuleProblems { param($s) $s.bystanderStoreDisplayNames = @('Synthetic Bystander') }) 'is not declared in bystanderStoreDisplayNames'
    Test-HasProblem 'a corpus store outside the census list is refused' (Get-RuleProblems { param($s) $s.corpus.storeDisplayName = 'Elsewhere'; $s.bystanderStoreDisplayNames = @('Synthetic Bystander', 'Elsewhere') }) "corpus.storeDisplayName: 'Elsewhere' is not in expectedStoreDisplayNames"
    Test-HasProblem 'a manifest not named after its corpus is refused' (Get-RuleProblems { param($s) $s.corpus.manifestPath = 'C:\OutlookAI-Q5\vm-synthetic.jsonl' }) 'must be named corpus-vm-synthetic.jsonl'
    Test-HasProblem 'a relative manifest path is refused' (Get-RuleProblems { param($s) $s.corpus.manifestPath = 'corpus-vm-synthetic.jsonl' }) 'must be an absolute path'
    Test-HasProblem 'a corpus id with a space is refused' (Get-RuleProblems { param($s) $s.corpus.corpusId = 'vm synthetic' }) 'ASCII letters, digits'
    Test-HasProblem 'a corpus id the record gives another guest is refused' (Get-RuleProblems { param($s) $s.corpus.corpusId = 'vm-other'; $s.corpus.manifestPath = 'C:\OutlookAI-Q5\corpus-vm-other.jsonl' }) "is not the id corpusIdConvention assigns"
    Test-HasProblem 'a corpus the record still calls not built is refused' (Get-RuleProblems -ForVm 'OutlookAI-Other' { param($s) $s.corpus.corpusId = 'vm-other'; $s.corpus.manifestPath = 'C:\OutlookAI-Q5\corpus-vm-other.jsonl' }) 'as not built'
    Test-HasProblem 'a guest the record assigns nothing is refused' (Get-RuleProblems -ForVm 'OutlookAI-Nobody' { param($s) }) 'has 0 entries for guest'
    Test-HasProblem 'a seed that disagrees with the record is refused' (Get-RuleProblems { param($s) $s.corpus.seed = 7 }) 'where corpusIdConvention records 4242'
    Test-HasProblem 'an item count that disagrees with the record is refused' (Get-RuleProblems { param($s) $s.corpus.itemCount = 999 }) 'where corpusIdConvention records 1000'
    Test-HasProblem 'an anchor that disagrees with the record is refused' (Get-RuleProblems { param($s) $s.corpus.anchorUtc = '2026-08-02T00:00:00Z' }) 'where corpusIdConvention records 2026-08-01'
    Test-Case 'the date-only spelling of the same anchor is accepted' '' ((Get-RuleProblems { param($s) $s.corpus.anchorUtc = '2026-08-01' }) -join ' | ')
    Test-HasProblem 'an anchor with an offset is refused' (Get-RuleProblems { param($s) $s.corpus.anchorUtc = '2026-08-01T00:00:00+02:00' }) 'corpus.anchorUtc: must be'
    Test-HasProblem 'a quoted seed is refused' (Get-RuleProblems { param($s) $s.corpus.seed = '4242' }) 'corpus.seed: must be an integer'
    Test-HasProblem 'a fractional seed is refused' (Get-RuleProblems { param($s) $s.corpus.seed = 4242.5 }) 'corpus.seed: must be an integer'
    Test-HasProblem 'an item count of zero is refused' (Get-RuleProblems { param($s) $s.corpus.itemCount = 0 }) 'corpus.itemCount: must be a whole number above zero'
    Test-HasProblem 'no measurement windows is refused' (Get-RuleProblems { param($s) $s.corpus.windowDays = @() }) 'forces a corpus rebuild every day'
    Test-HasProblem 'a window of zero days is refused' (Get-RuleProblems { param($s) $s.corpus.windowDays = @(0, 7) }) 'every entry must be a whole number of days'
    Test-HasProblem 'a window listed twice is refused' (Get-RuleProblems { param($s) $s.corpus.windowDays = @(7, 7) }) 'lists 7 twice'
    Test-HasProblem 'a partial corpus is refused' (Get-RuleProblems { param($s) $s.corpus.PSObject.Properties.Remove('manifestPath') }) 'corpus.manifestPath: missing'
    Test-HasProblem 'a corpus written as null is refused - absent means left out' (Get-RuleProblems { param($s) $s.corpus = $null }) 'never written as null'
    Test-HasProblem 'a sink that is not on loopback is refused' (Get-RuleProblems { param($s) $s.mailSink.submitHost = '0.0.0.0' }) 'must be loopback'
    Test-HasProblem 'a sink port of zero is refused' (Get-RuleProblems { param($s) $s.mailSink.retrievePort = 0 }) 'must be a port number'
    Test-HasProblem 'a sink port past 65535 is refused' (Get-RuleProblems { param($s) $s.mailSink.submitPort = 70000 }) 'must be a port number'
    Test-HasProblem 'a sink timeout of zero is refused' (Get-RuleProblems { param($s) $s.mailSink.connectTimeoutMs = 0 }) 'connectTimeoutMs'
    Test-HasProblem 'a field nobody wrote a rule for is refused' (Get-RuleProblems { param($s) $s | Add-Member -NotePropertyName 'delegateNestedFolderProbe' -NotePropertyValue 'x' }) 'delegateNestedFolderProbe: not a field this script has a rule for'
    Test-HasProblem 'one store spelt two ways is refused' (Get-RuleProblems { param($s) $s.bystanderStoreDisplayNames = @('BYSTANDER@render.invalid', 'Synthetic Corpus') }) "where expectedStoreDisplayNames spells it 'bystander@render.invalid'"
    Test-Case 'faults are all reported together' $true ((Get-RuleProblems { param($s) $s.machineProfile = 'Production'; $s.mailSink.submitHost = '10.0.0.1'; $s.corpus.itemCount = 0 }).Count -ge 3)

    # The indexed list, the probe values and the addresses.
    Test-HasProblem 'an indexed store the census does not watch is refused' (Get-RuleProblems { param($s) $s.indexedStoreDisplayNames = @('hub@render.invalid', 'nobody@render.invalid', 'Synthetic Corpus') }) 'in neither expectedStoreDisplayNames nor bystanderStoreDisplayNames'
    Test-HasProblem 'a delegate in the indexed list is refused' (Get-RuleProblems { param($s) $s.expectedDelegateStoreDisplayNames = @('shared@render.invalid'); $s.indexedStoreDisplayNames = @('hub@render.invalid', 'shared@render.invalid', 'Synthetic Corpus') }) 'is a delegate mailbox'
    Test-HasProblem 'the hub must come first in the indexed list' (Get-RuleProblems { param($s) $s.indexedStoreDisplayNames = @('bystander@render.invalid', 'hub@render.invalid', 'Synthetic Corpus') }) 'must come FIRST'
    Test-HasProblem 'an indexed list without the hub is refused' (Get-RuleProblems { param($s) $s.indexedStoreDisplayNames = @('bystander@render.invalid', 'Synthetic Corpus') }) 'must come FIRST'
    Test-HasProblem 'the corpus must come last in the indexed list' (Get-RuleProblems { param($s) $s.indexedStoreDisplayNames = @('hub@render.invalid', 'Synthetic Corpus', 'bystander@render.invalid') }) 'must come LAST'
    Test-HasProblem 'a small indexed store not named as an address is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames += 'Plain Store'; $s.bystanderStoreDisplayNames += 'Plain Store'; $s.indexedStoreDisplayNames = @('hub@render.invalid', 'Plain Store', 'Synthetic Corpus') }) 'is not named as an address'
    Test-HasProblem 'an address outside .invalid is refused' (Get-RuleProblems { param($s) $s.expectedStoreDisplayNames += 'mail@example.com'; $s.bystanderStoreDisplayNames += 'mail@example.com' }) 'outside the RFC 2606 .invalid domain'
    Test-HasProblem 'the indexed guest names indexed stores' (Get-RuleProblems { param($s) $s.indexedStoreDisplayNames = @() }) 'records as INDEXED'
    Test-HasProblem 'the unindexed guest names none' (Get-RuleProblems -ForVm 'OutlookAI-Other' { param($s) $s.corpus = $null; $s.PSObject.Properties.Remove('corpus'); $s.probeTerm = ''; $s.PSObject.Properties.Remove('subjectOnlyProbe') }) 'records as UNINDEXED'
    Test-HasProblem 'a record that does not say whether the guest is indexed is refused' (Get-RuleProblems -ForVm 'OutlookAI-Unknown' { param($s) $s.PSObject.Properties.Remove('corpus') }) "carries no true/false 'indexed'"
    Test-HasProblem 'a probe term of two words is refused' (Get-RuleProblems { param($s) $s.probeTerm = 'two words' }) 'is not one plain word'
    Test-HasProblem 'the indexed guest needs a probe term' (Get-RuleProblems { param($s) $s.probeTerm = '' }) 'empty on the INDEXED guest'
    Test-HasProblem 'the unindexed guest carries none' (Get-RuleProblems -ForVm 'OutlookAI-Other' { param($s) $s.PSObject.Properties.Remove('corpus'); $s.indexedStoreDisplayNames = @(); $s.PSObject.Properties.Remove('subjectOnlyProbe') }) 'on a guest with no index'
    Test-HasProblem 'the subject-only probe lives in the hub' (Get-RuleProblems { param($s) $s.subjectOnlyProbe.storeDisplayName = 'bystander@render.invalid' }) 'is not the hub'
    Test-HasProblem 'a partial subject-only probe is refused' (Get-RuleProblems { param($s) $s.subjectOnlyProbe.PSObject.Properties.Remove('subjectTerm') }) 'subjectOnlyProbe.subjectTerm: missing'
    Test-HasProblem 'a subject term too short to stem is refused' (Get-RuleProblems { param($s) $s.subjectOnlyProbe.subjectTerm = 'abc' }) 'at least five letters'
    Test-HasProblem 'a folder path with a leading slash is refused' (Get-RuleProblems { param($s) $s.subjectOnlyProbe.folderPath = '/Inbox/x' }) 'forward slashes and none at either end'
    Test-HasProblem 'the indexed guest needs the subject-only probe' (Get-RuleProblems { param($s) $s.PSObject.Properties.Remove('subjectOnlyProbe') }) 'absent on the INDEXED guest'
    Test-HasProblem 'a subject-only probe written as null is refused - absent means left out' (Get-RuleProblems { param($s) $s.subjectOnlyProbe = $null }) 'never written as null'
    Test-HasProblem 'a corpus below the recorded minimum size is refused' (Get-RuleProblems { param($s) $s.corpus.itemCount = 999 }) 'below the 1000 corpusIdConvention requires'

    # ---------------------------------------------------------------------------------------
    Write-Host ''
    Write-Host '== the line that copies it into the guest =='
    Test-Case 'is exact' 'pwsh -File Testbed/host/Copy-ToGuest.ps1 -VMName OutlookAI-Indexed -Path C:\x\live-test-settings.json -Destination C:\OutlookAI-Q5\src\McpServer\OutlookAI.McpServer.Tests\live-fixtures\live-test-settings.json' (Format-CopyToGuestLine -VMName 'OutlookAI-Indexed' -RenderedPath 'C:\x\live-test-settings.json')
    Test-Case 'and quotes a path with a space' $true ((Format-CopyToGuestLine -VMName 'OutlookAI-Indexed' -RenderedPath 'C:\a b\s.json').Contains('-Path "C:\a b\s.json"'))

    Write-Host ''
    Write-Host "$($script:SelfTestChecks) assertion(s), $($script:SelfTestFailures.Count) failure(s)."
    Write-Host ''
    Write-Host 'NOT COVERED HERE. Only a guest can settle these, and nothing above stands in for them:'
    Write-Host '  * that the guest''s TIER profile mounts every declared store, under exactly those names'
    Write-Host '  * that the hub is an account''s delivery store and that account''s SmtpAddress is the hub''s name'
    Write-Host '  * that the corpus manifest is where the settings say, and the corpus still fresh'
    Write-Host '  * Copy-ToGuest.ps1 landing the file, and the live tier loading it and starting'
    Write-Host '  * SUBST and mapped drives resolving through GetFinalPathNameByHandle - documented, not exercised'

    if ($script:SelfTestFailures.Count -gt 0) {
        Write-Host ''
        foreach ($failure in $script:SelfTestFailures) { Write-Host "  $failure" }
        return 1
    }
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE RENDER.
# =============================================================================================

# 1. The guest name, before it is used to build any path.
if ($VMName -cnotmatch '^[A-Za-z0-9][A-Za-z0-9-]{0,63}$') {
    throw "REFUSING: '$VMName' is not a VM name this script will use - letters, digits and hyphens only. It becomes part of a path."
}

$repoFull = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
if (-not $TemplatePath) { $TemplatePath = Join-Path $repoFull 'Testbed\live-test-settings.template.json' }
if (-not $TestbedJsonPath) { $TestbedJsonPath = Join-Path $repoFull 'Testbed\testbed.json' }
if (-not $OutPath) { $OutPath = Join-Path $repoFull (Join-Path '.work\live-test-settings' (Join-Path $VMName 'live-test-settings.json')) }

# 2. Where it may go - FIRST, before anything is read, so the most dangerous mistake is refused
#    before any other can be reported.
$pathProblem = Get-OutputPathProblem -Path $OutPath
if ($null -ne $pathProblem) { throw $pathProblem }
$outFull = [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutPath))
$outCanonical = Get-CanonicalPath -FullPath $outFull
$insideWorkTree = ($null -ne (Find-GitWorkTree -Directory ([System.IO.Path]::GetDirectoryName($outCanonical))))
$overwriteProblem = Get-OverwriteProblem -FullPath $outFull -InsideWorkTree $insideWorkTree -Force ([bool]$Force)
if ($null -ne $overwriteProblem) { throw $overwriteProblem }

# 3. The template, and the guest's values.
$template = ConvertFrom-JsonText -Text (Read-TextFile $TemplatePath 'The template') -What "The template ($TemplatePath)"
$shape = Get-TemplateFields -Template $template
# Assigned, never wrapped in @(): the function returns its array with the comma operator, and
# @() around that nests it - an empty list would then count as one problem.
$drift = Get-TemplateDriftProblems -Fields $shape.Fields
$templateProblems = @($shape.Problems) + $drift
if ($templateProblems.Count -gt 0) {
    $message = "REFUSING: the template at $TemplatePath has changed shape, and filling it now would produce a settings file nobody checked:`n  - " +
        ($templateProblems -join "`n  - ") + "`n.github/scripts/check-testbed-references.ps1 check 8 holds it to Testbed/live-test-settings.example.json; teach this script any new field before rendering."
    throw $message
}

$templateShown = Get-DisplayPath -Path $TemplatePath -Root $repoFull
$valuesShown = Get-DisplayPath -Path $TestbedJsonPath -Root $repoFull
$testbed = ConvertFrom-JsonText -Text (Read-TextFile $TestbedJsonPath 'testbed.json') -What "testbed.json ($TestbedJsonPath)"
$section = Get-ExactProperty $testbed 'liveTestSettings'
if ($null -eq $section -or -not (Test-IsJsonObject $section.Value)) {
    throw "REFUSING: $TestbedJsonPath has no liveTestSettings section, so there is nothing to render from."
}
$guestProperty = Get-ExactProperty $section.Value $VMName
if ($null -eq $guestProperty) {
    $known = @($section.Value.PSObject.Properties | Where-Object { -not $_.Name.StartsWith('_') } | ForEach-Object { $_.Name })
    throw "REFUSING: testbed.json's liveTestSettings has no section for '$VMName'. It has: $($known -join ', '). The name must match exactly - it is the Hyper-V VM name."
}

$resolved = Resolve-GuestValues -Fields $shape.Fields -Guest $guestProperty.Value -Where "liveTestSettings.$VMName"
if ($resolved.Problems.Count -gt 0) {
    $message = "REFUSING to render live-test settings for '$VMName': $($resolved.Problems.Count) value(s) in the liveTestSettings.$VMName section of $valuesShown are not ready:`n  - " +
        ($resolved.Problems -join "`n  - ") +
        "`nA placeholder is a value only the guest can supply: read it there - over COM, from the TIER profile the live tier runs under - and write it into $valuesShown. Nothing was written."
    throw $message
}

# 4. Render.
$provenance = Get-GitProvenance -Root $repoFull -Inputs @($valuesShown, $templateShown)
$document = New-RenderedDocument -Fields $shape.Fields -Resolved $resolved -VMName $VMName -Provenance $provenance -RenderedAtUtc ([datetime]::UtcNow) -TemplateShown $templateShown -ValuesShown $valuesShown
$text = (ConvertTo-JsonLiteral -Value $document) + "`n"

# 5. Check the ARTEFACT - what will be read - rather than the values that went into it.
$leftover = @([regex]::Matches($text, $script:TokenPattern) | ForEach-Object { $_.Value } | Sort-Object -Unique)
if ($leftover.Count -gt 0) {
    throw "REFUSING: the rendered file still holds token(s) $($leftover -join ', '). Nothing was written."
}
$parsed = ConvertFrom-JsonText -Text $text -What 'The rendered file'
$assigned = $null
$convention = Get-ExactProperty $testbed 'corpusIdConvention'
if ($null -ne $convention) {
    $assignedProperty = Get-ExactProperty $convention.Value 'assigned'
    if ($null -ne $assignedProperty) { $assigned = $assignedProperty.Value }
}
$problems = New-Object System.Collections.Generic.List[string]
Add-SettingsProblems -Settings $parsed -VMName $VMName -Assigned $assigned -Problems $problems
if ($problems.Count -gt 0) {
    $message = "REFUSING to write live-test settings for '$VMName': the rendered file breaks $($problems.Count) rule(s) the live tier or its documentation sets:`n  - " +
        ($problems -join "`n  - ") + "`nCorrect the liveTestSettings.$VMName section of $valuesShown. Nothing was written."
    throw $message
}

# 6. Write - after asking again where to, so nothing that changed in between is written through.
$pathProblem = Get-OutputPathProblem -Path $OutPath
if ($null -ne $pathProblem) { throw $pathProblem }
$overwriteProblem = Get-OverwriteProblem -FullPath $outFull -InsideWorkTree $insideWorkTree -Force ([bool]$Force)
if ($null -ne $overwriteProblem) { throw $overwriteProblem }
Write-RenderedFile -FullPath $outFull -Text $text

$readBack = [System.IO.File]::ReadAllText($outFull)
if ($readBack -cne $text) { throw "The file at $outFull does not read back as what was written. Do not copy it anywhere." }
[void](ConvertFrom-JsonText -Text $readBack -What 'The written file')
$hash = (Get-FileHash -LiteralPath $outFull -Algorithm SHA256).Hash

# 7. Say what it is, what it could not check, and what to do with it.
$hub = $parsed.testHubStoreDisplayName
$bystanderList = @($parsed.bystanderStoreDisplayNames)
$grant = Get-IdentityGrant -Settings $parsed
Write-Host ''
Write-Host "Live-test settings for '$VMName':"
Write-Host ("  file        {0}" -f $outFull)
Write-Host ("  sha256      {0}" -f $hash)
Write-Host ("  profile     {0}" -f $parsed.machineProfile)
Write-Host ("  hub         {0}   - the one store the suite may write to" -f $hub)
Write-Host ("  watched     {0} store(s): {1}" -f @($parsed.expectedStoreDisplayNames).Count, (@($parsed.expectedStoreDisplayNames) -join ', '))
$indexedList = @($parsed.indexedStoreDisplayNames)
if ($indexedList.Count -gt 0) {
    Write-Host ("  indexed     {0} store(s), in the order the index tests read them: {1}" -f $indexedList.Count, ($indexedList -join ', '))
}
else { Write-Host '  indexed     none - this guest has no index, and the index tests refuse rather than pass here' }
Write-Host ("  bystanders  {0}   - denied every write, and censused" -f ($bystanderList -join ', '))
if (([string]$parsed.probeTerm).Length -gt 0) { Write-Host ("  probeTerm   {0}" -f $parsed.probeTerm) }
if ($null -ne (Get-ExactProperty $parsed 'subjectOnlyProbe')) {
    Write-Host ("  SF-6 probe  '{0}' in {1}, term {2}, sender {3}" -f $parsed.subjectOnlyProbe.folderPath, $parsed.subjectOnlyProbe.storeDisplayName, $parsed.subjectOnlyProbe.subjectTerm, $parsed.subjectOnlyProbe.senderFragment)
}
if ($grant.Count -gt 0) {
    Write-Host ("  identity    {0}   - draft and delete only: the identity tests write here" -f ($grant -join ', '))
}
else {
    Write-Host '  identity    none - the two identity tests will announce PROVED NOTHING (Docs/live-tier-on-the-vm.md section 2.8b)'
}
if ($null -ne (Get-ExactProperty $parsed 'corpus')) {
    Write-Host ("  corpus      {0} in '{1}', manifest {2}, windows {3} day(s)" -f $parsed.corpus.corpusId, $parsed.corpus.storeDisplayName, $parsed.corpus.manifestPath, (@($parsed.corpus.windowDays) -join '/'))
}
else { Write-Host '  corpus      none declared - the freshness check is skipped' }
if ($null -ne (Get-ExactProperty $parsed 'mailSink')) {
    Write-Host ("  mailSink    submit {0}:{1}, retrieve {2}:{3}" -f $parsed.mailSink.submitHost, $parsed.mailSink.submitPort, $parsed.mailSink.retrieveHost, $parsed.mailSink.retrievePort)
}
else { Write-Host '  mailSink    none declared - the loader reads that as real transport; on a guest nothing listens, so do not send' }
Write-Host ''
Write-Host 'NOT CHECKED HERE - only the guest can answer these, and a wrong answer refuses the tier:'
Write-Host '  * that the TIER profile mounts every store named above, under exactly these names'
Write-Host '    (a declared store the running profile does not mount is censused, not found, and refused)'
Write-Host '  * that the hub is an account''s delivery store, and that account''s SmtpAddress is the hub''s name'
Write-Host '  * that the hub and bystander POPULATIONS are built and censused (Docs/live-tier-on-the-vm.md section 3b) -'
Write-Host '    every hub test, and the probe values above, read them'
if ($null -ne (Get-ExactProperty $parsed 'corpus')) {
    Write-Host ("  * that {0} exists on the guest, and the corpus is still fresh - the tier checks both at start" -f $parsed.corpus.manifestPath)
}
Write-Host ''
Write-Host 'Next - copy it into the guest (PowerShell Direct; the guest needs no network):'
Write-Host ("  {0}" -f (Format-CopyToGuestLine -VMName $VMName -RenderedPath $outFull))
Write-Host ''
Write-Host 'Then run the tier in session 1, through Testbed/guest/Register-InteractiveTask.ps1 - never straight'
Write-Host 'over PowerShell Direct, where Outlook cannot finish starting.'
