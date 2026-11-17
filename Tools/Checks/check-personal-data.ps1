#Requires -Version 5.1
<#
.SYNOPSIS
    Fails when personal data has reached the repository: a name, address or machine detail of the
    maintainer or of anybody else that has to stay out of this public repository.

.DESCRIPTION
    DECIDED BY THE MAINTAINER 2026-11-17 (Q185 and Q190, QUESTIONS.md). An audit of the whole history
    found a third party's shared mailbox and mail domain, a delegate's name, a colleague's first name,
    and the maintainer's own user name, computer name and mailbox details in documents, tests, TODO
    items and commit messages of this public repository - nearly all of them written by agents.
    Q185 (b) takes them out of the history. Q190 keeps them out: this check (a), git hooks that run it
    on every commit and push (b, installed by Tools/Install-GitHooks.ps1), the AGENTS.md rule that only
    placeholders are ever committed (c), and the key and certificate patterns in .gitignore (e).

    WHAT IT FINDS
      1. KNOWN IDENTIFIERS, BY HASH ONLY. Tools/Checks/personal-data.json lists the SHA-256 of each
         identifier, never the value. This check hashes every word of every file it reads and looks
         the hashes up. A word is a run of letters and digits, compared in lower case, so an address
         matches whatever punctuation joins its words, and a match never starts or ends inside a
         longer word ("majority" holds no name). An identifier of several words also lists the hash
         of its first word, and only where that word stands are the runs that start with it hashed:
         a whole tree is then a few tens of thousands of hashes, not millions. An identifier may be
         allowed in named files: the maintainer's name in LICENSE (Q186).
      2. THE PRIVATE LIST. The values themselves, each with the placeholder to write instead, live
         only in the maintainer's checkout, in .work\personal-data\identifiers.json, which git ignores
         (Q190 (c)). Where that file exists - on the workstation, so in a release and in the git
         hooks - it is read as well, so a value is guarded before its hash is committed.
      3. GENERIC SHAPES, which need no list: an e-mail address outside the reserved domains (RFC 2606
         and RFC 6761); a Windows profile path or a Claude project folder with a real-looking user
         name; a private key block; a key or certificate file (.pfx, .p12, .snk, .pem, .key); a
         Windows default computer name (DESKTOP-, LAPTOP-, WIN- and its random suffix).
      4. THE NAMES OF THE MACHINE IT RUNS ON: its user name, its computer name and its account SID,
         except a testbed VM's (OAI-*) and that VM's own account.
      5. AND .gitignore STILL COVERS the key and certificate files.

    THE BASELINE, UNTIL Q189 IS DONE. HEAD still carries values that tests and testbed files need until
    they read them from somewhere else (Q189). personal-data.json's baseline says, per file and per
    rule, how many there are: more fails, fewer is reported so that the ceiling comes down. Every
    ceiling reaches zero before the history is rewritten (Q191), and the baseline is then empty. The
    hooks need no baseline: they fail a commit that ADDS an occurrence to a file, and only that.

    MODES
      (default)            Every tracked file - git ls-files, or every file on disk under -RepoRoot
                           but build output where there is no git, as on the build VM. The release
                           runs it as a guard (Tools/Publish-Release.ps1 step 5), in both shells.
      -Hook pre-commit     What `git diff --cached` adds, per file. Run by the pre-commit hook.
      -Hook commit-msg F   The commit message in file F, up to git's scissors line.
      -Hook pre-push R U   Every commit the push sends to remote R that R does not have yet: its
                           message and what it adds, per file. Git's ref lines arrive on stdin.
      -HashIdentifiers     Show the committed list's identifiers as the private list makes them;
                           with -Execute, write them into personal-data.json. Prints no value.
      -SelfTest            Synthetic cases for every rule, then the commit's own files against the
                           committed list and its baseline. Runs on the build VM.

    IT PRINTS NO VALUE. A finding names the file, the line and column, and the identifier's id and
    kind; a generic finding shows its value masked. Its output can end up in a log, a TODO item or a
    commit message, and must not undo what it guards.

    WHAT IT CANNOT SEE, stated rather than implied. A value glued into a longer word ("xyzSmith"), a
    name the list does not hold, a fact about a mailbox that is no name at all (an item count), and
    binary files: a file with a NUL byte in its first 8 KB is skipped, except that a key or certificate
    file is found by its name. And since the hash of a short word can be found by hashing a dictionary
    of words, a committed hash protects a value from casual reading only.

    EXIT CODES: 0 clean; 1 personal data found; 3 not checked - nothing could be read, or git failed;
    4 refused - a malformed list or an unusable argument.

    WINDOWS POWERSHELL 5.1 AND POWERSHELL 7 BOTH. No ternary, no ??, ASCII only.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script; in a hook, to the working tree git
    runs the hook in.

.PARAMETER ListPath
    The committed list. Defaults to personal-data.json beside this script.

.PARAMETER IdentifiersFile
    The private list. Defaults to .work\personal-data\identifiers.json under -RepoRoot, then under the
    main checkout of the repository -RepoRoot belongs to (an agent's worktree finds the maintainer's).

.PARAMETER IncludeWorkingTree
    Also check untracked files git does not ignore - worth passing before a commit.

.PARAMETER ReportPath
    Also write every finding, baselined or not, as JSON to this file: per file and rule, the count
    and the lines. For tracking Q189's work; holds no value.

.PARAMETER Hook
    Run as the named git hook. Tools/Install-GitHooks.ps1 installs the hooks that pass it.

.PARAMETER HookArguments
    What git passes the hook: the message file for commit-msg, the remote's name and URL for
    pre-push.

.PARAMETER HashIdentifiers
    Compare the committed list with the private list; with -Execute, write the private list's
    hashes into the committed list.

.PARAMETER Execute
    With -HashIdentifiers: write the committed list.

.PARAMETER SelfTest
    Synthetic cases, then this commit's own files against the committed list. Build VM only.

.EXAMPLE
    pwsh -File Tools/Checks/check-personal-data.ps1
    powershell -NoProfile -File Tools/Checks/check-personal-data.ps1 -IncludeWorkingTree
#>
[CmdletBinding(DefaultParameterSetName = 'Tree', PositionalBinding = $false)]
param(
    [string] $RepoRoot,
    [string] $ListPath,
    [string] $IdentifiersFile,
    [Parameter(ParameterSetName = 'Tree')] [switch] $IncludeWorkingTree,
    [Parameter(ParameterSetName = 'Tree')] [string] $ReportPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Hook')] [ValidateSet('pre-commit', 'commit-msg', 'pre-push')] [string] $Hook,
    [Parameter(ParameterSetName = 'Hook', ValueFromRemainingArguments = $true)] [string[]] $HookArguments = @(),
    [Parameter(Mandatory = $true, ParameterSetName = 'Hash')] [switch] $HashIdentifiers,
    [Parameter(ParameterSetName = 'Hash')] [switch] $Execute,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)
# Defaults that need $PSScriptRoot are set in the body, not in param(): Windows PowerShell 5.1
# leaves it empty while param() defaults are evaluated (Tools/Checks/check-powershell-51.ps1, 1).
if (-not $PSBoundParameters.ContainsKey('ListPath')) { $ListPath = Join-Path $PSScriptRoot 'personal-data.json' }

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# ---------------------------------------------------------------------------------------------
# Constants. -SelfTest pins each of them.
# ---------------------------------------------------------------------------------------------

# Hashed in front of every value, so that a committed hash is not one any public table of
# SHA-256(word) already holds. Changing it changes every hash: personal-data.json is then
# rewritten with -HashIdentifiers -Execute.
$Salt = 'OutlookAI personal data v1|'

# Where the private list lives, relative to a checkout.
$PrivateListRelative = '.work\personal-data\identifiers.json'

# Reserved for documentation and testing (RFC 2606, RFC 6761): any address under them is fine.
$ReservedEmailDomains = @('example.com', 'example.net', 'example.org')
$ReservedTopLevelDomains = @('example', 'invalid', 'test', 'localhost')
# Real addresses that are public by design: git's own SSH user at GitHub, and the address agents'
# Co-Authored-By lines carry.
$AllowedEmailAddresses = @('git@github.com', 'noreply@anthropic.com')
$AllowedEmailDomainSuffixes = @('users.noreply.github.com')
# A last label that is a file's extension, not a top-level domain: x@example.com.ost is a file.
$FileExtensionLabels = @('ost', 'pst', 'nst', 'vsto', 'xml', 'json', 'txt', 'log', 'md', 'cs', 'ps1', 'dll', 'exe', 'msg', 'eml', 'zip', 'png')

# User names in a profile path that name nobody: placeholders, Windows' own profiles, the testbed
# VMs' account, and the stand-ins tests use. Anything else - and a name of three or more letters is
# a name - fails, so a test writes <user> or one of these.
$PlaceholderUserNames = @(
    'public', 'default', 'default user', 'all users', 'administrator', 'vmadmin', 'runneradmin',
    'me', 'you', 'user', 'username', 'someone', 'somebody', 'tester', 'test', 'testuser', 'example',
    'alice', 'bob', 'carol', 'dave', 'eve', 'mallory', 'trent'
)
# The testbed VMs' computer names (Testbed/testbed.json): public by design.
$TestbedComputerNamePrefix = 'OAI-'

# Key and certificate files: never tracked (Q190 (e)), and .gitignore must say so.
$KeyFileExtensions = @('.pfx', '.p12', '.snk', '.pem', '.key')
$KeyFileIgnoreRules = @('*.pfx', '*.p12', '*.snk', '*.pem', '*.key')

# A file is text unless it has a NUL byte in its first 8 KB (git's own test), or one of these.
$BinaryExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.ico', '.bmp', '.exe', '.dll', '.pdb', '.zip', '.7z', '.msi', '.cab', '.nupkg', '.snk', '.pfx', '.p12', '.der', '.cer')

# Directories a walk without git never enters: build output and scratch, never tracked here.
$WalkSkipDirectories = @('.git', 'bin', 'obj', '.vs', '.work', 'TestResults', 'node_modules', 'packages')

# The generic shapes, as .NET regular expressions.
$EmailPattern = '(?<![A-Za-z0-9._%+-])[A-Za-z0-9][A-Za-z0-9._%+-]*@(?<domain>[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,24})(?![A-Za-z0-9-])'
# A profile folder: Users, then the user's segment - doubled separators included, as JSON, C# and
# regex escaping write them. Whatever comes before (a drive, a variable, file:///) does not matter.
$UserPathPattern = '(?i)[\\/]+(?:Users|Documents and Settings)[\\/]+(?<user>[^\\/\s"''<>|:*?`,;)\]}]+)(?<after>[\\/]?)'
# Claude Code's project folders encode a path, C:\Users\<user>\x as C--Users-<user>-x.
$ClaudeSlugPattern = '(?i)(?<![A-Za-z0-9])[A-Z]--Users-(?<user>[A-Za-z0-9._]+?)(?:-|$|(?=[\\/"''\s]))'
$PrivateKeyPattern = '-----BEGIN (?:[A-Z0-9]+ )*PRIVATE KEY-----'
$DefaultComputerNamePattern = '(?<![A-Za-z0-9-])(?:DESKTOP-[A-Z0-9]{7}|LAPTOP-[A-Z0-9]{7}|WIN-[A-Z0-9]{11})(?![A-Za-z0-9-])'

# Compiled for a whole tree, where it pays for itself many times over; interpreted for a hook's few
# lines, where compiling would be most of the wait.
$RegexOptions = [System.Text.RegularExpressions.RegexOptions]::None
if ($PSCmdlet.ParameterSetName -eq 'Tree' -or $PSCmdlet.ParameterSetName -eq 'SelfTest') { $RegexOptions = [System.Text.RegularExpressions.RegexOptions]::Compiled }
$EmailRegex = New-Object System.Text.RegularExpressions.Regex($EmailPattern, $RegexOptions)
$UserPathRegex = New-Object System.Text.RegularExpressions.Regex($UserPathPattern, $RegexOptions)
$ClaudeSlugRegex = New-Object System.Text.RegularExpressions.Regex($ClaudeSlugPattern, $RegexOptions)
$PrivateKeyRegex = New-Object System.Text.RegularExpressions.Regex($PrivateKeyPattern, $RegexOptions)
$DefaultComputerNameRegex = New-Object System.Text.RegularExpressions.Regex($DefaultComputerNamePattern, $RegexOptions)

# Each shape runs only on a text that holds what it must contain: most files hold none of these,
# and a regular expression tried at every position of 14 MB is most of a release's wait.
function Test-Holds([string] $Text, [string[]] $Needles) {
    foreach ($n in $Needles) { if ($Text.IndexOf($n, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true } }
    return $false
}

$ExitClean = 0
$ExitFound = 1
$ExitNotChecked = 3
$ExitRefused = 4

# ---------------------------------------------------------------------------------------------
# The matcher. C# because a release hashes about two million word runs; Windows PowerShell 5.1's
# Add-Type compiles it with the .NET Framework's own csc (C# 5: no interpolation, no =>).
# ---------------------------------------------------------------------------------------------
$MatcherSource = @'
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAIPersonalData
{
    public sealed class WordSpan
    {
        public int Start;
        public int Length;
        public string Lower;
    }

    public sealed class IdentifierHit
    {
        public int Entry;
        public int Start;
        public int Length;
        public int Line;
        public int Column;
    }

    public static class TextShape
    {
        // Runs of letters and digits, lower-cased: what an identifier is compared word by word.
        public static List<WordSpan> Words(string text)
        {
            List<WordSpan> words = new List<WordSpan>();
            if (string.IsNullOrEmpty(text)) { return words; }
            int i = 0;
            int n = text.Length;
            while (i < n)
            {
                if (char.IsLetterOrDigit(text[i]))
                {
                    int start = i;
                    while (i < n && char.IsLetterOrDigit(text[i])) { i++; }
                    WordSpan w = new WordSpan();
                    w.Start = start;
                    w.Length = i - start;
                    w.Lower = text.Substring(start, i - start).ToLowerInvariant();
                    words.Add(w);
                }
                else { i++; }
            }
            return words;
        }

        public static string Normalize(string value)
        {
            List<WordSpan> words = Words(value);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < words.Count; i++)
            {
                if (i > 0) { sb.Append(' '); }
                sb.Append(words[i].Lower);
            }
            return sb.ToString();
        }

        public static int WordCount(string value)
        {
            return Words(value).Count;
        }

        public static int[] LineStarts(string text)
        {
            List<int> starts = new List<int>();
            starts.Add(0);
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\n') { starts.Add(i + 1); }
            }
            return starts.ToArray();
        }

        // The 1-based line and column of an offset.
        public static int[] Position(int[] lineStarts, int offset)
        {
            int lo = 0;
            int hi = lineStarts.Length - 1;
            while (lo < hi)
            {
                int mid = (lo + hi + 1) / 2;
                if (lineStarts[mid] <= offset) { lo = mid; } else { hi = mid - 1; }
            }
            return new int[] { lo + 1, offset - lineStarts[lo] + 1 };
        }

        public static string HashHex(string salt, string normalized)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(salt + normalized));
                StringBuilder sb = new StringBuilder(64);
                for (int i = 0; i < h.Length; i++) { sb.Append(h[i].ToString("x2")); }
                return sb.ToString();
            }
        }
    }

    public sealed class Matcher
    {
        private sealed class WordState
        {
            public int Entry = -1;          // a one-word identifier this word is
            public List<int> Starts;        // the several-word identifiers this word begins
        }

        private readonly byte[] saltBytes;
        private readonly SHA256 sha;
        private readonly List<byte[]> full = new List<byte[]>();
        private readonly int[] wordCounts;
        private readonly Dictionary<string, int> single = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<int>> firsts = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        private readonly Dictionary<string, WordState> seen = new Dictionary<string, WordState>(StringComparer.Ordinal);
        private byte[] buffer;

        // fullHashes[i]: the identifier's hash; firstHashes[i]: its first word's, or empty for one
        // word; wordCounts[i]: how many words it has.
        public Matcher(string salt, string[] fullHashes, string[] firstHashes, int[] wordCounts)
        {
            this.saltBytes = Encoding.UTF8.GetBytes(salt);
            this.sha = SHA256.Create();
            this.buffer = new byte[saltBytes.Length + 1024];
            Buffer.BlockCopy(saltBytes, 0, buffer, 0, saltBytes.Length);
            this.wordCounts = wordCounts;
            for (int i = 0; i < fullHashes.Length; i++)
            {
                string h = fullHashes[i].ToLowerInvariant();
                full.Add(FromHex(h));
                if (wordCounts[i] <= 1) { single[h] = i; continue; }
                string f = firstHashes[i];
                if (f == null || f.Length != 64) { throw new ArgumentException("an identifier of several words needs its first word's hash"); }
                f = f.ToLowerInvariant();
                List<int> list;
                if (!firsts.TryGetValue(f, out list)) { list = new List<int>(); firsts[f] = list; }
                list.Add(i);
            }
        }

        private static byte[] FromHex(string hex)
        {
            if (hex == null || hex.Length != 64) { throw new ArgumentException("a SHA-256 is 64 hex digits"); }
            byte[] b = new byte[32];
            for (int i = 0; i < 32; i++) { b[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16); }
            return b;
        }

        private byte[] Hash(string normalized)
        {
            int max = saltBytes.Length + Encoding.UTF8.GetMaxByteCount(normalized.Length);
            if (buffer.Length < max)
            {
                byte[] grown = new byte[max * 2];
                Buffer.BlockCopy(saltBytes, 0, grown, 0, saltBytes.Length);
                buffer = grown;
            }
            int len = Encoding.UTF8.GetBytes(normalized, 0, normalized.Length, buffer, saltBytes.Length);
            return sha.ComputeHash(buffer, 0, saltBytes.Length + len);
        }

        private static string Hex(byte[] h)
        {
            StringBuilder sb = new StringBuilder(64);
            for (int i = 0; i < h.Length; i++) { sb.Append(h[i].ToString("x2")); }
            return sb.ToString();
        }

        private static bool Same(byte[] a, byte[] b)
        {
            for (int j = 0; j < 32; j++) { if (a[j] != b[j]) { return false; } }
            return true;
        }

        // What one word is: hashed once per distinct word, however often it occurs.
        private WordState State(string word)
        {
            WordState s;
            if (seen.TryGetValue(word, out s)) { return s; }
            s = new WordState();
            string h = Hex(Hash(word));
            int e;
            if (single.TryGetValue(h, out e)) { s.Entry = e; }
            List<int> list;
            if (firsts.TryGetValue(h, out list)) { s.Starts = list; }
            if (seen.Count < 1000000) { seen[word] = s; }
            return s;
        }

        private IdentifierHit NewHit(int entry, List<WordSpan> words, int first, int last, ref int[] lineStarts, string text)
        {
            if (lineStarts == null) { lineStarts = TextShape.LineStarts(text); }
            IdentifierHit hit = new IdentifierHit();
            hit.Entry = entry;
            hit.Start = words[first].Start;
            hit.Length = words[last].Start + words[last].Length - words[first].Start;
            int[] pos = TextShape.Position(lineStarts, hit.Start);
            hit.Line = pos[0];
            hit.Column = pos[1];
            return hit;
        }

        // Every listed identifier in a text: the entry, the span, the line and the column.
        public List<IdentifierHit> Find(string text)
        {
            List<IdentifierHit> hits = new List<IdentifierHit>();
            if (full.Count == 0 || string.IsNullOrEmpty(text)) { return hits; }
            List<WordSpan> words = TextShape.Words(text);
            int[] lineStarts = null;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < words.Count; i++)
            {
                WordState s = State(words[i].Lower);
                if (s.Entry >= 0) { hits.Add(NewHit(s.Entry, words, i, i, ref lineStarts, text)); }
                if (s.Starts == null) { continue; }
                foreach (int entry in s.Starts)
                {
                    int n = wordCounts[entry];
                    if (i + n > words.Count) { continue; }
                    sb.Length = 0;
                    for (int k = 0; k < n; k++)
                    {
                        if (k > 0) { sb.Append(' '); }
                        sb.Append(words[i + k].Lower);
                    }
                    if (Same(Hash(sb.ToString()), full[entry])) { hits.Add(NewHit(entry, words, i, i + n - 1, ref lineStarts, text)); }
                }
            }
            return hits;
        }
    }
}
'@

function Initialize-Matcher {
    if (-not ('OutlookAIPersonalData.Matcher' -as [type])) {
        Add-Type -TypeDefinition $MatcherSource
    }
}

# ---------------------------------------------------------------------------------------------
# Pure helpers. No file, no git: -SelfTest pins every one of them.
# ---------------------------------------------------------------------------------------------

function ConvertTo-NormalizedIdentifier([string] $Value) {
    Initialize-Matcher
    return [OutlookAIPersonalData.TextShape]::Normalize($Value)
}

function Get-IdentifierHash([string] $Value) {
    Initialize-Matcher
    return [OutlookAIPersonalData.TextShape]::HashHex($Salt, (ConvertTo-NormalizedIdentifier $Value))
}

# The hash of a value's first word: what an identifier of several words is found by.
function Get-FirstWordHash([string] $Value) {
    $normalized = ConvertTo-NormalizedIdentifier $Value
    $first = @($normalized.Split(' '))[0]
    return [OutlookAIPersonalData.TextShape]::HashHex($Salt, $first)
}

# A value shown in a finding: its first character and its shape, never the rest.
function Get-MaskedValue([string] $Value) {
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    $at = $Value.IndexOf('@')
    if ($at -gt 0) {
        $domain = $Value.Substring($at + 1)
        $dot = $domain.LastIndexOf('.')
        $tld = ''
        if ($dot -ge 0) { $tld = $domain.Substring($dot) }
        $domainHead = ''
        if ($domain.Length -gt 0) { $domainHead = $domain.Substring(0, 1) }
        return $Value.Substring(0, 1) + '***@' + $domainHead + '***' + $tld
    }
    return $Value.Substring(0, 1) + '***'
}

function Test-ReservedEmailDomain([string] $Domain) {
    $d = $Domain.ToLowerInvariant()
    $labels = @($d.Split('.'))
    # x@example.com.ost is a file named after an address: judge the address.
    if ($labels.Count -gt 2 -and ($FileExtensionLabels -contains $labels[$labels.Count - 1])) {
        $d = ($labels[0..($labels.Count - 2)]) -join '.'
        $labels = @($d.Split('.'))
    }
    foreach ($r in $ReservedEmailDomains) {
        if ($d -eq $r -or $d.EndsWith('.' + $r)) { return $true }
    }
    return ($ReservedTopLevelDomains -contains $labels[$labels.Count - 1])
}

function Test-AllowedEmail([string] $Address, [string] $Domain) {
    if (Test-ReservedEmailDomain $Domain) { return $true }
    if ($AllowedEmailAddresses -contains $Address.ToLowerInvariant()) { return $true }
    foreach ($s in $AllowedEmailDomainSuffixes) {
        $d = $Domain.ToLowerInvariant()
        if ($d -eq $s -or $d.EndsWith('.' + $s)) { return $true }
    }
    return $false
}

# A user segment of a profile path that names nobody.
function Test-PlaceholderUserName([string] $Name, [bool] $LastSegment) {
    $n = $Name.Trim()
    if ($n.Length -le 2) { return $true }
    if ($n -match '^[<%$\{\*\.\[]') { return $true }
    if ($PlaceholderUserNames -contains $n.ToLowerInvariant()) { return $true }
    # C:\Users\x.vsto: a file directly under Users, not a user.
    if ($LastSegment -and $n -match '\.[A-Za-z0-9]{1,6}$') { return $true }
    return $false
}

function Test-KeyFilePath([string] $Path) {
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return ($KeyFileExtensions -contains $ext)
}

# Identifier entries a list contributes. Each: Id, Kind, Hash, First (its first word's hash, for
# several words), Words, AllowedIn (paths), Source.
function New-IdentifierEntry([string] $Id, [string] $Kind, [string] $Hash, [string] $First, [int] $Words, [string[]] $AllowedIn, [string] $Source) {
    $f = ''
    if ($First) { $f = $First.ToLowerInvariant() }
    return [pscustomobject]@{
        Id        = $Id
        Kind      = $Kind
        Hash      = $Hash.ToLowerInvariant()
        First     = $f
        Words     = [Math]::Max(1, $Words)
        AllowedIn = @($AllowedIn | Where-Object { $_ } | ForEach-Object { $_.Replace('\', '/') })
        Source    = $Source
    }
}

# An entry from a value: the private list's, the machine's, and the self-test's.
function New-EntryFromValue([string] $Id, [string] $Kind, [string] $Value, [string[]] $AllowedIn, [string] $Source) {
    Initialize-Matcher
    $words = [OutlookAIPersonalData.TextShape]::WordCount($Value)
    $first = ''
    if ($words -gt 1) { $first = Get-FirstWordHash $Value }
    return New-IdentifierEntry $Id $Kind (Get-IdentifierHash $Value) $first $words $AllowedIn $Source
}

# The names of the machine that runs the check, as identifiers - unless they are the testbed's.
function Get-MachineEntries([string] $UserName, [string] $ComputerName, [string] $UserSid) {
    $entries = @()
    if ($UserName -and -not (Test-PlaceholderUserName $UserName $false)) {
        $entries += New-EntryFromValue 'this-machine-user' 'the user name of the machine running this check' $UserName @() 'machine'
    }
    if ($ComputerName -and -not $ComputerName.ToUpperInvariant().StartsWith($TestbedComputerNamePrefix)) {
        $entries += New-EntryFromValue 'this-machine-name' 'the computer name of the machine running this check' $ComputerName @() 'machine'
    }
    # S-1-5-21-<a>-<b>-<c>-<rid>: <a> names the machine (or the domain) every local account of it
    # shares. Ten digits nothing else in this repository carries.
    if ($UserSid -match '^S-1-5-21-(\d{6,10})-\d+-\d+-\d+$') {
        $entries += New-EntryFromValue 'this-machine-sid' 'the account SID of the machine running this check' $Matches[1] @() 'machine'
    }
    return @($entries)
}

# The committed list, the private list and the machine's names as one catalog: entries deduplicated
# by hash (the committed entry wins, with its allowances), and the matcher over them.
function New-Catalog($Committed, $Private, $Machine) {
    $entries = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($group in @($Committed, $Private, $Machine)) {
        foreach ($e in @($group)) {
            if ($null -eq $e) { continue }
            if ($seen.ContainsKey($e.Hash)) { continue }
            $seen[$e.Hash] = $true
            [void]$entries.Add($e)
        }
    }
    Initialize-Matcher
    $hashes = [string[]]@($entries | ForEach-Object { $_.Hash })
    $firsts = [string[]]@($entries | ForEach-Object { $_.First })
    $words = [int[]]@($entries | ForEach-Object { $_.Words })
    return [pscustomobject]@{
        Entries = @($entries)
        Matcher = New-Object OutlookAIPersonalData.Matcher -ArgumentList $Salt, $hashes, $firsts, $words
    }
}

function Test-PathAllowed($Entry, [string] $Path) {
    if (-not $Path) { return $false }
    $p = $Path.Replace('\', '/')
    foreach ($a in @($Entry.AllowedIn)) {
        if ([string]::Equals($a, $p, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Get-TextPosition([int[]] $LineStarts, [int] $Offset) {
    return [OutlookAIPersonalData.TextShape]::Position($LineStarts, $Offset)
}

# Every finding in one text. $Path is repository-relative, or '' for a commit message. Each finding:
# Rule (an identifier's id or a generic rule), Kind, Line, Column, Shown (masked, or ''). Findings
# an allowance covers are not returned.
function Find-PersonalData([string] $Text, [string] $Path, $Catalog) {
    $findings = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    $spans = New-Object System.Collections.ArrayList
    foreach ($h in $Catalog.Matcher.Find($Text)) {
        $entry = $Catalog.Entries[$h.Entry]
        [void]$spans.Add(@($h.Start, ($h.Start + $h.Length)))
        if (Test-PathAllowed $entry $Path) { continue }
        [void]$findings.Add([pscustomobject]@{ Rule = $entry.Id; Kind = $entry.Kind; Line = $h.Line; Column = $h.Column; Shown = '' })
    }
    $lineStarts = $null

    $emails = @()
    if ($Text.IndexOf('@') -ge 0) { $emails = $EmailRegex.Matches($Text) }
    foreach ($m in $emails) {
        $domain = $m.Groups['domain'].Value
        if (Test-AllowedEmail $m.Value $domain) { continue }
        if ($null -eq $lineStarts) { $lineStarts = [OutlookAIPersonalData.TextShape]::LineStarts($Text) }
        $pos = Get-TextPosition $lineStarts $m.Index
        [void]$findings.Add([pscustomobject]@{ Rule = 'email'; Kind = 'an e-mail address outside the reserved domains'; Line = $pos[0]; Column = $pos[1]; Shown = (Get-MaskedValue $m.Value) })
    }
    $paths = @()
    if (Test-Holds $Text @('users', 'documents and settings')) { $paths = $UserPathRegex.Matches($Text) }
    foreach ($m in $paths) {
        $user = $m.Groups['user'].Value
        $last = ($m.Groups['after'].Value.Length -eq 0)
        if (Test-PlaceholderUserName $user $last) { continue }
        if ($null -eq $lineStarts) { $lineStarts = [OutlookAIPersonalData.TextShape]::LineStarts($Text) }
        $pos = Get-TextPosition $lineStarts $m.Groups['user'].Index
        [void]$findings.Add([pscustomobject]@{ Rule = 'user-path'; Kind = 'a user profile path with a real-looking user name'; Line = $pos[0]; Column = $pos[1]; Shown = (Get-MaskedValue $user) })
    }
    $slugs = @()
    if (Test-Holds $Text @('--users-')) { $slugs = $ClaudeSlugRegex.Matches($Text) }
    foreach ($m in $slugs) {
        $user = $m.Groups['user'].Value
        if (Test-PlaceholderUserName $user $false) { continue }
        if ($null -eq $lineStarts) { $lineStarts = [OutlookAIPersonalData.TextShape]::LineStarts($Text) }
        $pos = Get-TextPosition $lineStarts $m.Groups['user'].Index
        [void]$findings.Add([pscustomobject]@{ Rule = 'user-path'; Kind = 'a user profile path with a real-looking user name'; Line = $pos[0]; Column = $pos[1]; Shown = (Get-MaskedValue $user) })
    }
    $keys = @()
    if ($Text.IndexOf('PRIVATE KEY', [System.StringComparison]::Ordinal) -ge 0) { $keys = $PrivateKeyRegex.Matches($Text) }
    foreach ($m in $keys) {
        if ($null -eq $lineStarts) { $lineStarts = [OutlookAIPersonalData.TextShape]::LineStarts($Text) }
        $pos = Get-TextPosition $lineStarts $m.Index
        [void]$findings.Add([pscustomobject]@{ Rule = 'private-key'; Kind = 'a private key block'; Line = $pos[0]; Column = $pos[1]; Shown = '' })
    }
    $pcs = @()
    if (Test-Holds $Text @('DESKTOP-', 'LAPTOP-', 'WIN-')) { $pcs = $DefaultComputerNameRegex.Matches($Text) }
    foreach ($m in $pcs) {
        # The machine's own name is already an identifier hit: one finding per place.
        $covered = $false
        foreach ($s in $spans) { if ($m.Index -lt $s[1] -and ($m.Index + $m.Length) -gt $s[0]) { $covered = $true; break } }
        if ($covered) { continue }
        if ($null -eq $lineStarts) { $lineStarts = [OutlookAIPersonalData.TextShape]::LineStarts($Text) }
        $pos = Get-TextPosition $lineStarts $m.Index
        [void]$findings.Add([pscustomobject]@{ Rule = 'machine-name'; Kind = 'a Windows default computer name'; Line = $pos[0]; Column = $pos[1]; Shown = (Get-MaskedValue $m.Value) })
    }
    return @($findings)
}

# Findings counted per rule: rule -> @{ Count; Lines; Kind; Shown }.
function Group-Findings($Findings) {
    $byRule = @{}
    foreach ($f in @($Findings)) {
        if ($null -eq $f) { continue }
        if (-not $byRule.ContainsKey($f.Rule)) { $byRule[$f.Rule] = [pscustomobject]@{ Rule = $f.Rule; Kind = $f.Kind; Count = 0; Lines = New-Object System.Collections.ArrayList; Shown = New-Object System.Collections.ArrayList } }
        $g = $byRule[$f.Rule]
        $g.Count++
        [void]$g.Lines.Add($f.Line)
        if ($f.Shown -and -not $g.Shown.Contains($f.Shown)) { [void]$g.Shown.Add($f.Shown) }
    }
    return $byRule
}

# A finding as one line: where, which rule, what kind, the masked value. Never the value.
function Format-Finding([string] $Path, $Group, [string] $Extra) {
    $where = $Path
    if (-not $where) { $where = '(commit message)' }
    $lines = @($Group.Lines | Select-Object -First 8) -join ','
    $more = ''
    if ($Group.Lines.Count -gt 8) { $more = ",..." }
    $shown = ''
    if ($Group.Shown.Count -gt 0) { $shown = ' [' + (@($Group.Shown | Select-Object -First 3) -join ', ') + ']' }
    return "${where}:$lines$more  $($Group.Rule) - $($Group.Kind)$shown$Extra"
}

# The baseline's verdict for one file: rule -> ceiling, against what was found.
function Compare-Baseline([string] $Path, $Groups, $Baseline) {
    $over = @()
    $under = @()
    $ceilings = @{}
    foreach ($b in @($Baseline)) {
        if ($null -eq $b) { continue }
        if ([string]::Equals($b.path, $Path, [System.StringComparison]::OrdinalIgnoreCase)) { $ceilings[[string]$b.rule] = [int]$b.count }
    }
    foreach ($rule in @($Groups.Keys | Sort-Object)) {
        $g = $Groups[$rule]
        $ceiling = 0
        if ($ceilings.ContainsKey($rule)) { $ceiling = $ceilings[$rule] }
        if ($g.Count -gt $ceiling) { $over += [pscustomobject]@{ Group = $g; Ceiling = $ceiling } }
        elseif ($g.Count -lt $ceiling) { $under += [pscustomobject]@{ Rule = $rule; Count = $g.Count; Ceiling = $ceiling } }
    }
    foreach ($rule in @($ceilings.Keys | Sort-Object)) {
        if (-not $Groups.ContainsKey($rule) -and $ceilings[$rule] -gt 0) { $under += [pscustomobject]@{ Rule = $rule; Count = 0; Ceiling = $ceilings[$rule] } }
    }
    return [pscustomobject]@{ Over = @($over); Under = @($under) }
}

# A unified diff (git diff -U0, git log -p) as files: Path, Added (line, text), Removed (text), and
# whether the file is new. Renamed files are their new path.
function ConvertFrom-UnifiedDiff([string] $DiffText) {
    $files = New-Object System.Collections.ArrayList
    $current = $null
    $newLine = 0
    $inHunk = $false
    foreach ($raw in ($DiffText -split "`n")) {
        $line = $raw.TrimEnd("`r")
        if ($line.StartsWith('diff --git ')) {
            $current = [pscustomobject]@{ Path = ''; IsNew = $false; Added = New-Object System.Collections.ArrayList; Removed = New-Object System.Collections.ArrayList }
            [void]$files.Add($current)
            $inHunk = $false
            # "a/<p> b/<p>": the one path a file that is not renamed has - and the only place a
            # binary file, which has no --- and +++ lines, says it.
            $pair = $line.Substring(11)
            if ($pair.Length -ge 7 -and (($pair.Length - 5) % 2) -eq 0) {
                $half = ($pair.Length - 5) / 2
                $a = $pair.Substring(0, 2 + $half)
                $b = $pair.Substring(2 + $half)
                if ($a.StartsWith('a/') -and $b -ceq (' b/' + $a.Substring(2))) { $current.Path = $a.Substring(2) }
            }
            continue
        }
        if ($null -eq $current) { continue }
        if (-not $inHunk) {
            # The header: what it says about the file. A content line is never read as one.
            if ($line.StartsWith('new file mode')) { $current.IsNew = $true; continue }
            if ($line.StartsWith('rename to ')) { $current.Path = $line.Substring(10); $current.IsNew = $true; continue }
            if ($line.StartsWith('+++ ')) {
                $p = $line.Substring(4)
                if ($p.StartsWith('"') -and $p.EndsWith('"')) { $p = $p.Substring(1, $p.Length - 2) }
                if ($p.StartsWith('b/')) { $current.Path = $p.Substring(2) }
                continue
            }
            if ($line.StartsWith('--- ')) {
                # A deletion's only path: kept so that what it removes is counted.
                $p = $line.Substring(4)
                if ($p.StartsWith('"') -and $p.EndsWith('"')) { $p = $p.Substring(1, $p.Length - 2) }
                if ($p.StartsWith('a/') -and -not $current.Path) { $current.Path = $p.Substring(2) }
                continue
            }
        }
        if ($line.StartsWith('@@')) {
            $m = [regex]::Match($line, '^@@ -\d+(?:,\d+)? \+(\d+)(?:,\d+)? @@')
            if ($m.Success) { $newLine = [int]$m.Groups[1].Value; $inHunk = $true }
            continue
        }
        if (-not $inHunk) { continue }
        if ($line.StartsWith('+')) {
            [void]$current.Added.Add([pscustomobject]@{ Line = $newLine; Text = $line.Substring(1) })
            $newLine++
            continue
        }
        if ($line.StartsWith('-')) {
            [void]$current.Removed.Add($line.Substring(1))
            continue
        }
        if ($line.StartsWith(' ')) { $newLine++ }
    }
    return @($files)
}

# What a diff ADDS: per file and rule, the occurrences added minus those removed. Anything above
# zero fails - a line moved within a file, or edited around its value, adds nothing. A new key file,
# and a file whose name holds an identifier, fail by their name alone.
function Get-DiffFindings($Files, $Catalog) {
    $failures = @()
    foreach ($f in @($Files)) {
        if (-not $f.Path) { continue }
        if ($f.IsNew -and (Test-KeyFilePath $f.Path)) {
            $failures += "$($f.Path)  key-file - a key or certificate file; .gitignore keeps these out (Q190 (e))"
        }
        $nameFindings = @(Find-PersonalData $f.Path.Replace('/', ' / ') '' $Catalog | Where-Object { $_.Rule -notin @('email', 'user-path', 'machine-name', 'private-key') })
        if ($f.IsNew -and $nameFindings.Count -gt 0) {
            $failures += "$($f.Path)  the file's own name holds $($nameFindings[0].Rule) - $($nameFindings[0].Kind)"
        }
        $addedText = (@($f.Added | ForEach-Object { $_.Text }) -join "`n")
        $removedText = (@($f.Removed) -join "`n")
        $added = Group-Findings (Find-PersonalData $addedText $f.Path $Catalog)
        $removed = Group-Findings (Find-PersonalData $removedText $f.Path $Catalog)
        foreach ($rule in @($added.Keys | Sort-Object)) {
            $g = $added[$rule]
            $gone = 0
            if ($removed.ContainsKey($rule)) { $gone = $removed[$rule].Count }
            if ($g.Count -le $gone) { continue }
            # Lines of the added text are 1-based positions in Added: map them to the file's lines.
            $fileLines = New-Object System.Collections.ArrayList
            foreach ($l in $g.Lines) { [void]$fileLines.Add($f.Added[$l - 1].Line) }
            $g.Lines = $fileLines
            $failures += (Format-Finding $f.Path $g "  (+$($g.Count - $gone))")
        }
    }
    return @($failures)
}

# A commit message as git keeps it: everything up to the scissors line of `git commit -v`.
function Get-MessageText([string] $Raw) {
    $lines = @($Raw -split "`r?`n")
    $kept = New-Object System.Collections.ArrayList
    foreach ($l in $lines) {
        if ($l -match '^. -+ >8 -+$') { break }
        [void]$kept.Add($l)
    }
    return ($kept -join "`n")
}

function Get-MessageFindings([string] $Message, $Catalog, [string] $Label) {
    $groups = Group-Findings (Find-PersonalData $Message '' $Catalog)
    $out = @()
    foreach ($rule in @($groups.Keys | Sort-Object)) { $out += (Format-Finding $Label $groups[$rule] '') }
    return @($out)
}

# git's pre-push input: one line per ref, "<local ref> <local oid> <remote ref> <remote oid>".
function ConvertFrom-PrePushInput([string] $Text) {
    $refs = @()
    foreach ($raw in @($Text -split "`r?`n")) {
        $parts = @($raw.Trim() -split '\s+')
        if ($parts.Count -ne 4) { continue }
        if ($parts[1] -match '^0+$') { continue }   # a deletion sends nothing
        $refs += [pscustomobject]@{ LocalRef = $parts[0]; LocalOid = $parts[1]; RemoteRef = $parts[2]; RemoteOid = $parts[3] }
    }
    return @($refs)
}

# ---------------------------------------------------------------------------------------------
# Lists.
# ---------------------------------------------------------------------------------------------

function Read-CommittedList([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "the committed list $Path does not exist" }
    $json = [System.IO.File]::ReadAllText($Path) | ConvertFrom-Json
    $entries = @()
    $ids = @{}
    foreach ($i in @($json.identifiers)) {
        if ($null -eq $i) { continue }
        if (-not ($i.id -match '^[a-z][a-z0-9-]{1,30}$')) { throw "an identifier in $Path has the id '$($i.id)': lower-case letters, digits and hyphens" }
        if ($ids.ContainsKey($i.id)) { throw "the id $($i.id) is listed twice in $Path" }
        $ids[$i.id] = $true
        if (-not ([string]$i.sha256 -cmatch '^[0-9a-f]{64}$')) { throw "the identifier $($i.id) in $Path has no SHA-256 of 64 lower-case hex digits" }
        if (-not $i.kind) { throw "the identifier $($i.id) in $Path says no kind" }
        $words = 1
        $first = ''
        if ($null -ne $i.words) {
            $words = [int]$i.words
            if ($words -lt 2 -or $words -gt 6) { throw "the identifier $($i.id) in $Path has $words words; an entry that says its words has 2 to 6" }
            if (-not ([string]$i.first -cmatch '^[0-9a-f]{64}$')) { throw "the identifier $($i.id) in $Path has several words but no SHA-256 of its first" }
            $first = [string]$i.first
        }
        $entries += New-IdentifierEntry $i.id $i.kind $i.sha256 $first $words @($i.allowedIn) 'committed'
    }
    return [pscustomobject]@{ Entries = @($entries); Baseline = @($json.baseline | Where-Object { $null -ne $_ }); Raw = $json }
}

function Read-PrivateList([string] $Path) {
    $json = [System.IO.File]::ReadAllText($Path) | ConvertFrom-Json
    $entries = @()
    foreach ($i in @($json.identifiers)) {
        if ($null -eq $i -or -not $i.value) { continue }
        $id = [string]$i.id
        if (-not $id) { $id = 'private' }
        $kind = [string]$i.kind
        if (-not $kind) { $kind = 'an identifier on the private list' }
        $entries += New-EntryFromValue $id $kind ([string]$i.value) @($i.allowedIn) 'private'
    }
    return @($entries)
}

# The main checkout of the repository a path belongs to: the folder holding git's common directory.
function Get-MainCheckoutRoot([string] $Root) {
    $r = Invoke-Git -Arguments @('rev-parse', '--path-format=absolute', '--git-common-dir') -WorkingDirectory $Root
    if ($null -eq $r -or $r.ExitCode -ne 0) { return $null }
    $common = $r.Out.Trim()
    if (-not $common) { return $null }
    if ((Split-Path -Leaf $common) -ne '.git') { return $null }
    return (Split-Path -Parent $common)
}

function Resolve-PrivateListPath([string] $Root) {
    if ($IdentifiersFile) { return $IdentifiersFile }
    $local = Join-Path $Root $PrivateListRelative
    if (Test-Path -LiteralPath $local -PathType Leaf) { return $local }
    $main = Get-MainCheckoutRoot $Root
    if ($main) {
        $shared = Join-Path $main $PrivateListRelative
        if (Test-Path -LiteralPath $shared -PathType Leaf) { return $shared }
    }
    return $null
}

function Get-CurrentMachine {
    $sid = ''
    try { $sid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value } catch { $sid = '' }
    return [pscustomobject]@{ UserName = [Environment]::UserName; ComputerName = [Environment]::MachineName; UserSid = $sid }
}

# ---------------------------------------------------------------------------------------------
# Git, through System.Diagnostics.Process: UTF-8 both ways whatever the console's code page, stdin
# when a command reads it, and no window (AGENTS.md: no console window from a background process).
# Not through PowerShell's own native call, which decodes git's output in the console's code page
# under Windows PowerShell 5.1 and would mangle a path or a value that is not ASCII.
# ---------------------------------------------------------------------------------------------

function ConvertTo-ArgumentString([string[]] $Arguments) {
    $parts = foreach ($a in $Arguments) {
        if ($a -notmatch '[\s"]' -and $a.Length -gt 0) { $a; continue }
        # CommandLineToArgvW: backslashes before a quote are doubled, and so are those at the end.
        $escaped = [regex]::Replace($a, '(\\*)"', { param($m) ($m.Groups[1].Value * 2) + '\"' })
        $escaped = [regex]::Replace($escaped, '(\\+)$', { param($m) $m.Groups[1].Value * 2 })
        '"' + $escaped + '"'
    }
    return ($parts -join ' ')
}

function Invoke-Git([string[]] $Arguments, [string] $WorkingDirectory, [string] $InputText) {
    $git = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $git) { return $null }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $git.Path
    $psi.Arguments = ConvertTo-ArgumentString (@('-c', 'core.quotePath=false') + $Arguments)
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardInput = $true
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $psi.StandardErrorEncoding = New-Object System.Text.UTF8Encoding($false)
    $p = [System.Diagnostics.Process]::Start($psi)
    $errTask = $p.StandardError.ReadToEndAsync()
    # Written whole and closed before the output is read: the one command given input here, git log
    # --stdin, reads all of it before it writes anything, and waits for the end of it to begin.
    $stdin = New-Object System.IO.StreamWriter($p.StandardInput.BaseStream, (New-Object System.Text.UTF8Encoding($false)))
    if ($InputText) { $stdin.Write($InputText) }
    $stdin.Close()
    $out = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()
    return [pscustomobject]@{ ExitCode = $p.ExitCode; Out = $out; Err = $errTask.Result }
}

# ---------------------------------------------------------------------------------------------
# The tree.
# ---------------------------------------------------------------------------------------------

# Repository-relative paths, forward slashes: what git tracks, or - with no git, as on the build VM,
# where the commit arrives as a `git archive` - every file but build output and scratch.
function Get-TreeFiles([string] $Root, [bool] $WithUntracked) {
    $r = Invoke-Git -Arguments @('ls-files', '-z') -WorkingDirectory $Root
    if ($null -ne $r -and $r.ExitCode -eq 0) {
        $files = @($r.Out.Split([char]0) | Where-Object { $_ })
        if ($WithUntracked) {
            $u = Invoke-Git -Arguments @('ls-files', '-z', '--others', '--exclude-standard') -WorkingDirectory $Root
            if ($null -ne $u -and $u.ExitCode -eq 0) { $files += @($u.Out.Split([char]0) | Where-Object { $_ }) }
        }
        return [pscustomobject]@{ Files = $files; Source = 'git ls-files' }
    }
    $rootFull = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd('\')
    $files = New-Object System.Collections.ArrayList
    $stack = New-Object System.Collections.Stack
    $stack.Push($rootFull)
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        foreach ($d in [System.IO.Directory]::GetDirectories($dir)) {
            if ($WalkSkipDirectories -contains (Split-Path -Leaf $d)) { continue }
            $stack.Push($d)
        }
        foreach ($f in [System.IO.Directory]::GetFiles($dir)) {
            [void]$files.Add($f.Substring($rootFull.Length + 1).Replace('\', '/'))
        }
    }
    return [pscustomobject]@{ Files = @($files | Sort-Object); Source = 'a walk of the folder (no git repository)' }
}

# A file's text, or $null for a binary one.
function Read-TextFile([string] $FullPath) {
    $ext = [System.IO.Path]::GetExtension($FullPath).ToLowerInvariant()
    if ($BinaryExtensions -contains $ext) { return $null }
    $bytes = [System.IO.File]::ReadAllBytes($FullPath)
    $probe = [Math]::Min($bytes.Length, 8192)
    $isUtf16 = ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)))
    if (-not $isUtf16 -and $probe -gt 0 -and [Array]::IndexOf($bytes, [byte]0, 0, $probe) -ge 0) { return $null }
    $stream = New-Object System.IO.MemoryStream(, $bytes)
    $reader = New-Object System.IO.StreamReader($stream, (New-Object System.Text.UTF8Encoding($false)), $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Test-GitIgnoreCoversKeys([string] $Root) {
    $path = Join-Path $Root '.gitignore'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return @($KeyFileIgnoreRules) }
    $lines = @([System.IO.File]::ReadAllLines($path) | ForEach-Object { $_.Trim() })
    return @($KeyFileIgnoreRules | Where-Object { $lines -notcontains $_ })
}

# The whole check of a tree. Returns the verdict and what it saw; prints nothing.
function Invoke-TreeCheck([string] $Root, $Catalog, $Baseline, [bool] $WithUntracked) {
    $tree = Get-TreeFiles $Root $WithUntracked
    $failures = New-Object System.Collections.ArrayList
    $notices = New-Object System.Collections.ArrayList
    $report = New-Object System.Collections.ArrayList
    $scanned = 0
    foreach ($rel in @($tree.Files)) {
        $full = Join-Path $Root $rel
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { continue }
        if (Test-KeyFilePath $rel) {
            [void]$failures.Add("$rel  key-file - a key or certificate file is tracked; .gitignore keeps these out (Q190 (e))")
        }
        $text = Read-TextFile $full
        if ($null -eq $text) { continue }
        $scanned++
        $groups = Group-Findings (Find-PersonalData $text $rel $Catalog)
        foreach ($rule in @($groups.Keys | Sort-Object)) {
            [void]$report.Add([ordered]@{ path = $rel; rule = $rule; count = $groups[$rule].Count; lines = @($groups[$rule].Lines) })
        }
        $verdict = Compare-Baseline $rel $groups $Baseline
        foreach ($o in $verdict.Over) {
            $extra = ''
            if ($o.Ceiling -gt 0) { $extra = "  ($($o.Group.Count), baseline $($o.Ceiling))" }
            [void]$failures.Add((Format-Finding $rel $o.Group $extra))
        }
        foreach ($u in $verdict.Under) {
            [void]$notices.Add("${rel}: $($u.Rule) $($u.Count), below its baseline of $($u.Ceiling) - lower the ceiling in personal-data.json")
        }
    }
    # Baseline entries naming a file the tree no longer has.
    $present = @{}
    foreach ($rel in @($tree.Files)) { $present[$rel.ToLowerInvariant()] = $true }
    foreach ($b in @($Baseline)) {
        if ($null -eq $b) { continue }
        if (-not $present.ContainsKey(([string]$b.path).ToLowerInvariant())) { [void]$notices.Add("$($b.path): not in the tree any more - delete its baseline entry") }
    }
    $missingRules = @(Test-GitIgnoreCoversKeys $Root)
    if ($missingRules.Count -gt 0) {
        [void]$failures.Add(".gitignore  no longer ignores $($missingRules -join ', '): those rules keep key and certificate files out of every commit (Q190 (e))")
    }
    return [pscustomobject]@{ Failures = @($failures); Notices = @($notices); Report = @($report); Scanned = $scanned; Files = @($tree.Files).Count; Source = $tree.Source }
}

# ---------------------------------------------------------------------------------------------
# The list's maintenance: hashes from the private list. Prints ids and kinds, never a value.
# ---------------------------------------------------------------------------------------------

function ConvertTo-JsonString([string] $s) {
    return '"' + $s.Replace('\', '\\').Replace('"', '\"') + '"'
}

function ConvertTo-CommittedListJson($About, $Identifiers, $Baseline) {
    $nl = "`n"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('{' + $nl + '  "about": [' + $nl)
    $aboutLines = @($About)
    for ($i = 0; $i -lt $aboutLines.Count; $i++) {
        $comma = ','
        if ($i -eq $aboutLines.Count - 1) { $comma = '' }
        [void]$sb.Append('    ' + (ConvertTo-JsonString $aboutLines[$i]) + $comma + $nl)
    }
    [void]$sb.Append('  ],' + $nl + '  "identifiers": [' + $nl)
    $ids = @($Identifiers)
    for ($i = 0; $i -lt $ids.Count; $i++) {
        $e = $ids[$i]
        $allowed = @($e.AllowedIn | ForEach-Object { ConvertTo-JsonString $_ }) -join ', '
        $comma = ','
        if ($i -eq $ids.Count - 1) { $comma = '' }
        $several = ''
        if ($e.Words -gt 1) { $several = ', "words": ' + $e.Words + ', "first": ' + (ConvertTo-JsonString $e.First) }
        [void]$sb.Append('    { "id": ' + (ConvertTo-JsonString $e.Id) + ', "kind": ' + (ConvertTo-JsonString $e.Kind) + ', "sha256": ' + (ConvertTo-JsonString $e.Hash) + $several + ', "allowedIn": [' + $allowed + '] }' + $comma + $nl)
    }
    [void]$sb.Append('  ],' + $nl + '  "baseline": [' + $nl)
    $base = @($Baseline | Where-Object { $null -ne $_ })
    for ($i = 0; $i -lt $base.Count; $i++) {
        $b = $base[$i]
        $comma = ','
        if ($i -eq $base.Count - 1) { $comma = '' }
        [void]$sb.Append('    { "path": ' + (ConvertTo-JsonString ([string]$b.path)) + ', "rule": ' + (ConvertTo-JsonString ([string]$b.rule)) + ', "count": ' + [int]$b.count + ' }' + $comma + $nl)
    }
    [void]$sb.Append('  ]' + $nl + '}' + $nl)
    return $sb.ToString()
}

# ---------------------------------------------------------------------------------------------
# SELF-TEST
# ---------------------------------------------------------------------------------------------
function Invoke-SelfTest {
    $script:Checks = 0
    $script:Failures = @()

    # Compared with -ceq; never -like, whose brackets are a character class.
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

    Write-Host "Tools/Checks/check-personal-data.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    Write-Host ''
    Initialize-Matcher

    # Every synthetic value that would trip a generic rule is assembled here, from parts, so that this
    # file itself never holds one: the tree check at the end reads this file too.
    $at = '@'
    $realDomain = 'zqx' + 'corp' + '.nl'
    $realAddress = 'p.someone' + $at + $realDomain
    $realUser = 'wilma' + 'zandvoort'
    $keyBlock = '-----BEGIN ' + 'RSA PRIVATE KEY-----'
    $defaultPc = 'DESKTOP-' + 'AB12CD3'

    Write-Host '== words and hashes =='
    Test-Case 'a value is its lower-case words, one space apart' 'ab cd ef 12' (ConvertTo-NormalizedIdentifier 'Ab-Cd  EF.12')
    Test-Case 'an address is its words' 'someone example net' (ConvertTo-NormalizedIdentifier ('Someone' + $at + 'Example.NET'))
    Test-Case 'letters with diacritics are letters' ('zo' + [char]0x00EB + ' m' + [char]0x00FC + 'ller') (ConvertTo-NormalizedIdentifier ('Zo' + [char]0x00EB + ' M' + [char]0x00FC + 'ller'))
    Test-Case 'the salt is pinned: changing it changes every committed hash' 'OutlookAI personal data v1|' $Salt
    $pinned = Get-IdentifierHash 'Zebra Finch'
    Test-Case 'the hash is SHA-256 of the salt and the normalized words, pinned' 'c685da9104cb8e753c6079885f3e75e8c00777450cf040f8e357298b91ed6693' $pinned
    Test-Case 'the same words in another case and punctuation hash alike' $pinned (Get-IdentifierHash 'zebra-FINCH')

    $zebra = New-EntryFromValue 'pdx1' 'a synthetic bird' 'zebrafinch' @('LICENSE') 'committed'
    $pair = New-EntryFromValue 'pdx2' 'a synthetic pair' 'lorem ipsum' @() 'committed'
    Test-Case 'a value of two words is found by its first' "2 $(Get-IdentifierHash 'lorem')" "$($pair.Words) $($pair.First)"
    Test-Case 'a value of one word has no first' '1 ' "$($zebra.Words) $($zebra.First)"
    $catalog = New-Catalog @($zebra, $pair) @() @()

    Write-Host ''
    Write-Host '== known identifiers =='
    $f = @(Find-PersonalData "one`nThe ZebraFinch sang" 'Docs/x.md' $catalog)
    Test-Case 'a word is found whatever its case' 'pdx1 2 5' ("$($f[0].Rule) $($f[0].Line) $($f[0].Column)")
    Test-Case 'never inside a longer word' 0 @(Find-PersonalData 'zebrafinches and xzebrafinch' 'Docs/x.md' $catalog).Count
    Test-Case 'two words across punctuation' 'pdx2' @(Find-PersonalData ('lorem' + $at + 'ipsum.example') 'Docs/x.md' $catalog)[0].Rule
    Test-Case 'two words across a line break' 'pdx2 1' ("$(@(Find-PersonalData "lorem`r`n  ipsum" 'Docs/x.md' $catalog)[0].Rule) $(@(Find-PersonalData "lorem`r`n  ipsum" 'Docs/x.md' $catalog)[0].Line)")
    Test-Case 'not with a word between them' 0 @(Find-PersonalData 'lorem dolor ipsum' 'Docs/x.md' $catalog).Count
    Test-Case 'an allowance covers its file' 0 @(Find-PersonalData 'Copyright ZebraFinch' 'LICENSE' $catalog).Count
    Test-Case 'and only its file' 1 @(Find-PersonalData 'Copyright ZebraFinch' 'Docs/LICENSE.md' $catalog).Count
    Test-Case 'a finding carries no value' '' @(Find-PersonalData 'zebrafinch' 'x.cs' $catalog)[0].Shown
    $twice = New-Catalog @($zebra) @((New-EntryFromValue 'private' 'the same value, privately' 'ZEBRAFINCH' @() 'private')) @()
    Test-Case 'the private list adds no second entry for a committed hash' 1 @($twice.Entries).Count
    Test-Case 'and the committed one keeps its allowance' 0 @(Find-PersonalData 'zebrafinch' 'LICENSE' $twice).Count

    Write-Host ''
    Write-Host '== the machine running the check =='
    $machine = @(Get-MachineEntries $realUser ('PC' + '4711') 'S-1-5-21-1234567890-2345678901-3456789012-1001')
    Test-Case 'its user, its computer and its SID are identifiers' 'this-machine-user this-machine-name this-machine-sid' (@($machine | ForEach-Object { $_.Id }) -join ' ')
    $mc = New-Catalog @() @() $machine
    Test-Case 'its user name is found in a sentence' 'this-machine-user' @(Find-PersonalData "logged on as $realUser here" 'x.md' $mc)[0].Rule
    Test-Case 'its SID by the part that names the machine' 'this-machine-sid' @(Find-PersonalData 'S-1-5-21-1234567890-1-2-500' 'x.md' $mc)[0].Rule
    Test-Case 'a testbed VM and its account are not' 0 @(Get-MachineEntries 'vmadmin' 'OAI-BUILD' '').Count

    Write-Host ''
    Write-Host '== generic shapes =='
    Test-Case 'an address at a real domain is found, masked' ('p***@z***.nl') @(Find-PersonalData "mail $realAddress now" 'x.cs' $catalog)[0].Shown
    foreach ($ok in @(('a' + $at + 'example.com'), ('tier' + $at + 'vm.invalid'), ('x' + $at + 'mail.example'), ('y' + $at + 'host.test'), ('carol' + $at + 'example.com.ost'), ('git' + $at + 'github.com'), ('noreply' + $at + 'anthropic.com'), ('1+k' + $at + 'users.noreply.github.com'))) {
        Test-Case "reserved or public by design: $(Get-MaskedValue $ok)" 0 @(Find-PersonalData $ok 'x.cs' $catalog | Where-Object { $_.Rule -eq 'email' }).Count
    }
    Test-Case 'a profile path with a real-looking name is found' 'user-path' @(Find-PersonalData ('C:\Users\' + $realUser + '\AppData') 'x.md' $catalog)[0].Rule
    Test-Case 'escaped as JSON too' 'user-path' @(Find-PersonalData ('"C:\\Users\\' + $realUser + '\\Documents"') 'x.json' $catalog)[0].Rule
    Test-Case 'and as a file URL' 'user-path' @(Find-PersonalData ('file:///C:/Users/' + $realUser + '/x.vsto') 'x.ps1' $catalog)[0].Rule
    Test-Case 'and as a Claude project folder' 'user-path' @(Find-PersonalData ('C--Users-' + $realUser + '-Downloads') 'x.md' $catalog)[0].Rule
    foreach ($ok in @('C:\Users\<user>\AppData', 'C:\Users\%USERNAME%\x', 'C:\Users\$env:USERNAME\x', 'C:\Users\vmadmin\Documents', 'C:\Users\Public\x', 'C:\Users\x\AppData', 'file:///C:/Users/x.vsto', 'C:\Users\...\x', 'C:\Users\tester\AppData')) {
        Test-Case "a placeholder path: $ok" 0 @(Find-PersonalData $ok 'x.cs' $catalog).Count
    }
    Test-Case 'a private key block is found' 'private-key' @(Find-PersonalData "x`n$keyBlock`nMIIE" 'x.txt' $catalog)[0].Rule
    Test-Case 'a public certificate block is not' 0 @(Find-PersonalData ('-----BEGIN ' + 'CERTIFICATE-----') 'x.txt' $catalog).Count
    Test-Case 'a Windows default computer name is found' 'machine-name' @(Find-PersonalData "host $defaultPc." 'x.md' $catalog)[0].Rule
    Test-Case 'a testbed name is not' 0 @(Find-PersonalData 'WIN-CLEAN and OAI-BUILD' 'x.md' $catalog).Count
    foreach ($p in @('a.pfx', 'b/c.P12', 'k.snk', 'x.pem', 'y.key')) { Test-Case "a key file by name: $p" $true (Test-KeyFilePath $p) }
    Test-Case 'a keys.json is not' $false (Test-KeyFilePath 'keys.json')

    Write-Host ''
    Write-Host '== the baseline =='
    $groups = Group-Findings @([pscustomobject]@{ Rule = 'pdx1'; Kind = 'k'; Line = 3; Column = 1; Shown = '' }, [pscustomobject]@{ Rule = 'pdx1'; Kind = 'k'; Line = 9; Column = 1; Shown = '' })
    $base = @([pscustomobject]@{ path = 'T/a.cs'; rule = 'pdx1'; count = 2 })
    Test-Case 'as many as the ceiling passes' '0 0' ("$(@((Compare-Baseline 'T/a.cs' $groups $base).Over).Count) $(@((Compare-Baseline 'T/a.cs' $groups $base).Under).Count)")
    $low = @([pscustomobject]@{ path = 'T/a.cs'; rule = 'pdx1'; count = 1 })
    Test-Case 'one more fails' 1 @((Compare-Baseline 'T/a.cs' $groups $low).Over).Count
    $high = @([pscustomobject]@{ path = 'T/a.cs'; rule = 'pdx1'; count = 5 })
    Test-Case 'fewer is a notice, not a failure' '0 1' ("$(@((Compare-Baseline 'T/a.cs' $groups $high).Over).Count) $(@((Compare-Baseline 'T/a.cs' $groups $high).Under).Count)")
    Test-Case 'a ceiling is per file' 1 @((Compare-Baseline 'T/b.cs' $groups $base).Over).Count

    Write-Host ''
    Write-Host '== the hooks: what a diff adds, a message, a push =='
    $diff = @(
        'diff --git a/T/a.cs b/T/a.cs',
        'index 1111111..2222222 100644',
        '--- a/T/a.cs',
        '+++ b/T/a.cs',
        '@@ -10,1 +10,1 @@',
        '-    var s = "zebrafinch one";',
        '+    var s = "zebrafinch two";',
        '@@ -20,0 +21,2 @@',
        '+    // nothing here',
        '+    var t = "lorem ipsum";',
        'diff --git a/K/new.pfx b/K/new.pfx',
        'new file mode 100644',
        'index 0000000..3333333',
        'Binary files /dev/null and b/K/new.pfx differ',
        'diff --git a/D/zebrafinch-notes.md b/D/zebrafinch-notes.md',
        'new file mode 100644',
        '--- /dev/null',
        '+++ b/D/zebrafinch-notes.md',
        '@@ -0,0 +1 @@',
        '+nothing personal'
    ) -join "`n"
    $files = @(ConvertFrom-UnifiedDiff $diff)
    Test-Case 'a diff is its files' 'T/a.cs | K/new.pfx | D/zebrafinch-notes.md' (@($files | ForEach-Object { $_.Path }))
    Test-Case 'with the new-file lines numbered' '10 21 22' (@($files[0].Added | ForEach-Object { $_.Line }) -join ' ')
    $diffFailures = @(Get-DiffFindings $files $catalog)
    Test-Case 'an edited line that keeps its value adds nothing; a new pair does' 1 @($diffFailures | Where-Object { $_.StartsWith('T/a.cs:22  pdx2') }).Count
    Test-Case 'and nothing else in that file' 1 @($diffFailures | Where-Object { $_.StartsWith('T/a.cs') }).Count
    Test-Case 'a new key file fails by its name' 1 @($diffFailures | Where-Object { $_.StartsWith('K/new.pfx  key-file') }).Count
    Test-Case 'a new file whose name holds an identifier fails' 1 @($diffFailures | Where-Object { $_.StartsWith('D/zebrafinch-notes.md  the file''s own name') }).Count
    $moved = @(
        'diff --git a/T/b.cs b/T/b.cs', '--- a/T/b.cs', '+++ b/T/b.cs',
        '@@ -3,1 +2,0 @@', '-x zebrafinch', '@@ -9,0 +9,1 @@', '+x zebrafinch'
    ) -join "`n"
    Test-Case 'a line moved within a file adds nothing' 0 @(Get-DiffFindings @(ConvertFrom-UnifiedDiff $moved) $catalog).Count
    $msg = "Fix the zebrafinch thing`n`nCo-Authored-By: Claude <noreply$($at)anthropic.com>`n# ------------------------ >8 ------------------------`n+ lorem ipsum"
    Test-Case 'a message is cut at the scissors line' $false ((Get-MessageText $msg).Contains('lorem'))
    Test-Case 'and its identifiers are found, the trailer''s address is not' 'pdx1' (@(Get-MessageFindings (Get-MessageText $msg) $catalog '(commit message)') | ForEach-Object { ($_ -split '\s+')[1] })
    $push = "refs/heads/master 1234567890123456789012345678901234567890 refs/heads/master abcdefabcdefabcdefabcdefabcdefabcdefabcd`nrefs/heads/gone 0000000000000000000000000000000000000000 refs/heads/gone 1111111111111111111111111111111111111111`n"
    $refs = @(ConvertFrom-PrePushInput $push)
    Test-Case 'a push is its refs, deletions left out' '1 refs/heads/master 1234567890123456789012345678901234567890' ("$($refs.Count) $($refs[0].LocalRef) $($refs[0].LocalOid)")
    Test-Case 'an argument with a space is quoted for git' 'log "a b" "c\"d"' (ConvertTo-ArgumentString @('log', 'a b', 'c"d'))

    Write-Host ''
    Write-Host '== the committed list, read as the release reads it =='
    $repoRoot = $RepoRoot
    $listOk = $true
    try { $list = Read-CommittedList $ListPath } catch { $listOk = $false; Test-Case "personal-data.json reads: $($_.Exception.Message)" $true $false }
    if ($listOk) {
        Test-Case 'personal-data.json reads' $true $true
        Test-Case 'it lists identifiers' $true (@($list.Entries).Count -gt 0)
        Test-Case 'every allowance names a file in the tree' '' (@($list.Entries | ForEach-Object { $_.AllowedIn } | Where-Object { $_ -and -not (Test-Path -LiteralPath (Join-Path $repoRoot $_) -PathType Leaf) }) -join ' ')
        $known = @{}
        foreach ($e in $list.Entries) { $known[$e.Id] = $true }
        foreach ($r in @('email', 'user-path', 'private-key', 'machine-name')) { $known[$r] = $true }
        Test-Case 'every baseline entry names a known rule' '' (@($list.Baseline | Where-Object { -not $known.ContainsKey([string]$_.rule) } | ForEach-Object { "$($_.path) $($_.rule)" }) -join ' ')
        Test-Case 'and a count above zero' '' (@($list.Baseline | Where-Object { [int]$_.count -lt 1 } | ForEach-Object { "$($_.path) $($_.rule)" }) -join ' ')
        $round = ConvertTo-CommittedListJson $list.Raw.about $list.Entries $list.Baseline
        Test-Case 'and -HashIdentifiers -Execute would write it back byte for byte' ([System.IO.File]::ReadAllText($ListPath).Replace("`r`n", "`n")) $round

        Write-Host ''
        Write-Host '== this commit''s own files, against the committed list and its baseline =='
        $vm = Get-CurrentMachine
        $realCatalog = New-Catalog $list.Entries @() (Get-MachineEntries $vm.UserName $vm.ComputerName $vm.UserSid)
        $result = Invoke-TreeCheck $repoRoot $realCatalog $list.Baseline $false
        Write-Host "  read $($result.Scanned) text files of $($result.Files), by $($result.Source)"
        Test-Case 'it read the tree' $true ($result.Scanned -gt 100)
        foreach ($n in $result.Notices) { Write-Host "  NOTE $n" }
        Test-Case 'no personal data beyond the baseline' '' ($result.Failures -join ' | ')
    }

    Write-Host ''
    Write-Host "$($script:Checks) assertion(s), $($script:Failures.Count) failure(s)."
    if ($script:Failures.Count -gt 0) {
        foreach ($x in $script:Failures) { Write-Host "  FAILED: $x" }
        return 1
    }
    return 0
}

# ---------------------------------------------------------------------------------------------
# Modes.
# ---------------------------------------------------------------------------------------------

if ($SelfTest) {
    if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
    exit (Invoke-SelfTest)
}

Initialize-Matcher

if ($Hook) {
    if (-not $PSBoundParameters.ContainsKey('RepoRoot')) {
        $top = Invoke-Git -Arguments @('rev-parse', '--show-toplevel') -WorkingDirectory (Get-Location).ProviderPath
        if ($null -eq $top -or $top.ExitCode -ne 0) { Write-Host "personal-data hook: git did not say where the working tree is - not checked."; exit $ExitNotChecked }
        $RepoRoot = $top.Out.Trim()
    }
}
elseif (-not $PSBoundParameters.ContainsKey('RepoRoot')) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}

try { $committed = Read-CommittedList $ListPath }
catch { Write-Host "REFUSED: $($_.Exception.Message)"; exit $ExitRefused }

$privatePath = Resolve-PrivateListPath $RepoRoot
$private = @()
if ($privatePath) {
    try { $private = @(Read-PrivateList $privatePath) }
    catch { Write-Host "REFUSED: the private list $privatePath does not read: $($_.Exception.Message)"; exit $ExitRefused }
}
$machine = Get-CurrentMachine
$catalog = New-Catalog $committed.Entries $private (Get-MachineEntries $machine.UserName $machine.ComputerName $machine.UserSid)

if ($HashIdentifiers) {
    if (-not $privatePath) { Write-Host "REFUSED: no private list - looked for $PrivateListRelative under $RepoRoot and the main checkout."; exit $ExitRefused }
    $committedById = @{}
    foreach ($c in $committed.Entries) { $committedById[$c.Id] = $c }
    Write-Host "The private list has $($private.Count) identifier(s); the committed list $($committed.Entries.Count)."
    foreach ($p in $private) {
        $state = 'new'
        if ($committedById.ContainsKey($p.Id)) {
            if ($committedById[$p.Id].Hash -eq $p.Hash) { $state = 'same' } else { $state = 'changed' }
        }
        Write-Host ("  {0,-8} {1,-8} {2}" -f $p.Id, $state, $p.Kind)
    }
    foreach ($c in $committed.Entries) { if (@($private | Where-Object { $_.Id -eq $c.Id }).Count -eq 0) { Write-Host ("  {0,-8} {1,-8} {2}" -f $c.Id, 'dropped', $c.Kind) } }
    if (-not $Execute) { Write-Host 'Dry run: nothing written. -Execute writes these hashes into the committed list.'; exit $ExitClean }
    $json = ConvertTo-CommittedListJson $committed.Raw.about $private $committed.Baseline
    [System.IO.File]::WriteAllText($ListPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Wrote $ListPath."
    exit $ExitClean
}

if ($Hook) {
    $label = "personal-data $Hook"
    $failures = @()
    if ($Hook -eq 'commit-msg') {
        if ($HookArguments.Count -lt 1 -or -not (Test-Path -LiteralPath $HookArguments[0] -PathType Leaf)) { Write-Host "${label}: no message file - not checked."; exit $ExitNotChecked }
        $text = [System.IO.File]::ReadAllText($HookArguments[0])
        $failures = @(Get-MessageFindings (Get-MessageText $text) $catalog '(commit message)')
    }
    elseif ($Hook -eq 'pre-commit') {
        $d = Invoke-Git -Arguments @('diff', '--cached', '-U0', '--no-color', '--no-ext-diff', '-M', '--diff-filter=ACDMRT') -WorkingDirectory $RepoRoot
        if ($null -eq $d -or $d.ExitCode -ne 0) { Write-Host "${label}: git diff --cached failed - not checked."; exit $ExitNotChecked }
        $failures = @(Get-DiffFindings (ConvertFrom-UnifiedDiff $d.Out) $catalog)
    }
    else {
        $remote = ''
        if ($HookArguments.Count -ge 1) { $remote = $HookArguments[0] }
        $stdinText = [Console]::In.ReadToEnd()
        foreach ($ref in @(ConvertFrom-PrePushInput $stdinText)) {
            $kind = Invoke-Git -Arguments @('cat-file', '-t', $ref.LocalOid) -WorkingDirectory $RepoRoot
            if ($null -eq $kind -or $kind.ExitCode -ne 0) { Write-Host "${label}: git cat-file failed on $($ref.LocalRef) - not checked."; exit $ExitNotChecked }
            if ($kind.Out.Trim() -eq 'tag') {
                $tag = Invoke-Git -Arguments @('cat-file', 'tag', $ref.LocalOid) -WorkingDirectory $RepoRoot
                $body = ''
                $blank = $tag.Out.IndexOf("`n`n")
                if ($blank -ge 0) { $body = $tag.Out.Substring($blank + 2) }
                $failures += @(Get-MessageFindings $body $catalog "(tag $($ref.LocalRef))")
            }
            $exclude = @('--not', '--remotes')
            if ($remote -and $remote -notmatch '[:/\\]') { $exclude = @('--not', "--remotes=$remote") }
            $list = Invoke-Git -Arguments (@('rev-list', "$($ref.LocalOid)^{}") + $exclude) -WorkingDirectory $RepoRoot
            if ($null -eq $list -or $list.ExitCode -ne 0) { Write-Host "${label}: git rev-list failed on $($ref.LocalRef) - not checked."; exit $ExitNotChecked }
            $commits = @($list.Out -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($commits.Count -eq 0) { continue }
            $marker = 'OUTLOOKAI-PD-COMMIT '
            $log = Invoke-Git -Arguments @('log', '--no-walk=unsorted', '--stdin', '-p', '-U0', '--no-color', '--no-ext-diff', '-M', '--diff-merges=first-parent', "--format=$marker%H%n%B") -WorkingDirectory $RepoRoot -InputText (($commits -join "`n") + "`n")
            if ($null -eq $log -or $log.ExitCode -ne 0) { Write-Host "${label}: git log failed on $($ref.LocalRef) - not checked."; exit $ExitNotChecked }
            $chunks = @($log.Out -split "(?m)^$marker")
            foreach ($chunk in $chunks) {
                if (-not $chunk.Trim()) { continue }
                $nl = $chunk.IndexOf("`n")
                $sha = $chunk.Substring(0, $nl).Trim()
                $rest = $chunk.Substring($nl + 1)
                $diffAt = $rest.IndexOf("`ndiff --git ")
                $message = $rest
                $diffText = ''
                if ($rest.StartsWith('diff --git ')) { $message = ''; $diffText = $rest }
                elseif ($diffAt -ge 0) { $message = $rest.Substring(0, $diffAt); $diffText = $rest.Substring($diffAt + 1) }
                $short = $sha.Substring(0, [Math]::Min(10, $sha.Length))
                $failures += @(Get-MessageFindings $message $catalog "(message of $short)")
                foreach ($x in @(Get-DiffFindings (ConvertFrom-UnifiedDiff $diffText) $catalog)) { $failures += "$short $x" }
            }
        }
    }
    if ($failures.Count -eq 0) { exit $ExitClean }
    Write-Host ''
    Write-Host "${label}: REFUSED - this would publish personal data (AGENTS.md, Personal data; Q190):"
    foreach ($x in $failures) { Write-Host "  $x" }
    Write-Host ''
    Write-Host '  Write a placeholder instead: a reserved domain (example.com, *.invalid, *.test), <user>, or the'
    Write-Host '  placeholder the private list names for that identifier. The value itself belongs only in'
    Write-Host "  $PrivateListRelative. Never bypass this with --no-verify."
    if (-not $privatePath) { Write-Host "  (No private list was found, so only the committed hashes and the generic shapes were checked.)" }
    exit $ExitFound
}

# The tree.
Write-Host "Checking that no personal data has reached $RepoRoot"
$sourceNote = 'the committed list'
if ($privatePath) { $sourceNote = "the committed list and the private list ($($private.Count))" }
Write-Host "  identifiers: $(@($catalog.Entries).Count) from $sourceNote and this machine's names; generic shapes: e-mail, profile paths, private keys, key files, default computer names"
$result = Invoke-TreeCheck $RepoRoot $catalog $committed.Baseline ([bool]$IncludeWorkingTree)
if ($ReportPath) {
    $dir = Split-Path -Parent $ReportPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $json = ConvertTo-Json -InputObject @($result.Report) -Depth 5
    [System.IO.File]::WriteAllText($ReportPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "  report: $ReportPath"
}
Write-Host "  read $($result.Scanned) text files of $($result.Files), by $($result.Source)"
foreach ($n in $result.Notices) { Write-Host "  NOTE $n" }
Write-Host ''
if ($result.Scanned -eq 0) {
    Write-Host '::error::PERSONAL DATA - read no file at all, which cannot be right: this check has switched itself off.'
    exit $ExitNotChecked
}
if ($result.Failures.Count -gt 0) {
    foreach ($f in $result.Failures) { Write-Host "::error::PERSONAL DATA - $f" }
    Write-Host "$($result.Failures.Count) finding(s) of personal data beyond the baseline. Replace each with a placeholder (AGENTS.md, Personal data); the values belong only in $PrivateListRelative."
    exit $ExitFound
}
Write-Host "No personal data beyond the baseline in $($result.Scanned) files."
exit $ExitClean
