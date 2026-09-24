#Requires -Version 5.1
<#
.SYNOPSIS
    Fails when the testbed documentation points at something this repository does not contain.

.DESCRIPTION
    A runbook that nothing verifies drifts, and this repository has already been bitten by
    exactly that twice: a runbook one test count behind that also recommended a superseded store
    layout, and - three times in one week - material that a document depended on kept only in a
    scratch directory that was then cleared without warning. What was lost was not notes. It was
    the corpus manifest that makes a 20,000-item synthetic corpus verifiable and removable, the
    scripts that built the guest, and the parameters that every published measurement is a
    statement about.

    So this script asks one question of the tree: DOES THE THING THE DOCUMENT NAMES EXIST HERE?

    EIGHT CHECKS.

    1. NO DANGLING REPOSITORY REFERENCE. Every repository-relative path named in a tracked
       document or testbed script must exist, or be on the declared list below with a reason.
       The matcher is deliberately narrow - a reference counts only when it starts with a real
       top-level directory of this repository - because a heuristic that also matches `tools/call`
       and `now/at-anchor` produces noise, and a check people learn to ignore is worse than none.

    2. INTENTIONALLY-ABSENT PATHS ARE STILL GITIGNORED. Every entry on that list that claims to
       be absent BECAUSE it is machine-local is checked against `git check-ignore`. An ignore rule
       that gets deleted is silent until the day something lands - and one of these files is a
       credential, in a repository that has published one before.

    3. THE CORPUS PARAMETERS AGREE EVERYWHERE. corpusId, seed, anchor and item count exist in
       three places that cannot see each other: a JSON parameter file, a PowerShell script's
       defaults, and a prose document. Those four values reproduce the corpus every published
       sweep and frame measurement rests on, so a copy that drifts is a measurement nobody can
       reproduce.

    4. THE INDEX AND THE DIRECTORY AGREE. Every script under Testbed/ is named by the table in
       Testbed/README.md, and every path that table names exists. Both directions: an unlisted
       script is one a rebuilder will not find, and a listed one that is gone is the drift this
       whole script exists to catch.

    5. EVERY TESTBED SCRIPT PARSES. CI cannot run them - there is no Hyper-V, no Outlook and no
       guest on a runner - but a script that does not parse cannot be run by a rebuilder either,
       and the parser is free.

    6. NOTHING CREDENTIAL-SHAPED UNDER Testbed/. Not paranoia: the test VM's guest password was
       committed to this public repository, survives in 32 commits of history, and had to be
       rotated. Testbed/ is where a rebuilder is most tempted to write one down.

    7. THE ANSWER-FILE TEMPLATE STILL HOLDS PLACEHOLDERS, NOT VALUES. The unattended install needs
       a password in the answer file Windows Setup reads, and the fastest way to make an
       unattended install work is to type one into the template and commit it. Check 6 would not
       catch that: an unattend password is <Password><Value>...</Value></Password>, which is not
       an assignment and does not read as one. So this check parses the template and requires
       EVERY password value in it to be exactly the placeholder token - not merely
       placeholder-shaped, not empty, not base64. It also fails on a committed autounattend.xml
       anywhere in the tree, because the filled file is the one that carries the credential and it
       belongs in gitignored scratch. Testbed/host/New-AnswerFile.ps1 is what fills the template,
       and it refuses to write its output anywhere but there.

    8. THE LIVE-TEST SETTINGS TEMPLATE HOLDS TOKENS ONLY, AND ITS FIELDS ARE THE EXAMPLE'S. A test
       guest's gitignored live-test-settings.json is RENDERED, by Testbed/host/New-LiveTestSettings.ps1,
       from Testbed/live-test-settings.template.json and that guest's section of Testbed/testbed.json.
       Three committed files describe one shape - the documented example, the template, and the
       per-guest values - and nothing else ties them together. So this requires the template's
       field set to equal the example's (_-prefixed notes aside) and every value in it to be the
       double-brace token spelling its own path, never a value; every guest section of testbed.json
       to name exactly the template's fields; the renderer's guest destination to sit under the
       same source root Testbed/host/Publish-LiveTierPayload.ps1 stages the suite into; every
       address-shaped value in the example - fields, lists and notes alike - to end in the RFC 2606
       .invalid domain, so the documented shape can never carry a real mailbox (Q79, 2026-09-24);
       and no file called live-test-settings.json to be tracked anywhere, since the real one names
       real stores.
       It lives here rather than in a T1 test because every one of those files is under Testbed/,
       and the workflow that runs T1 only triggers on McpServer/ - this script runs on every pull
       request. T1/LiveTestSettingsTemplateTests covers the other half: that the template still
       renders into something the live tier's loader accepts.

    Run it from anywhere:
        pwsh -File .github/scripts/check-testbed-references.ps1

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script.
#>
[CmdletBinding()]
param(
    [string] $RepoRoot
)

$ErrorActionPreference = 'Stop'

# Defaulted HERE rather than in param(): Windows PowerShell 5.1 leaves $PSScriptRoot EMPTY inside the
# param() defaults of a script run with -File, so the one-liner that used to sit there threw before the
# first check ran - on 5.1 this guard never checked anything. PowerShell 7 fills it, which is why CI
# never noticed.
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$script:Failures = @()
$script:Checks = 0

function Fail([string] $invariant, [string] $detail) {
    $script:Failures += "$invariant`n    $detail"
}
function Pass([string] $invariant, [string] $detail) {
    Write-Host "  OK   $invariant - $detail"
}

Write-Host "Checking testbed references under $RepoRoot"
Write-Host ''

# ---------------------------------------------------------------------------------------------
# The declared list of paths that documents may name and the repository deliberately does not
# contain. Every entry needs a reason, and a reason that is a decision rather than an excuse.
# ---------------------------------------------------------------------------------------------
$intentionallyAbsent = @(
    @{
        Path       = 'Docs/live-test-inventory.txt'
        MustIgnore = $false
        Why        = 'DELETED 2026-09-15. A tracked snapshot of the live-tier traits that had gone stale in a way that misled: it printed the RETIRED LiveTier trait, and printed Requires as a PER-CLASS UNION - the exact shape that once turned a real floor of six impossible tests into a reported 96. Nothing generated it and nothing checked it. The traits on the tests are the authority and "dotnet test --list-tests" answers the same question from the source of truth. Docs/research/pop3-account-routes.md still names it because that note is a dated snapshot and is deliberately not edited to stay current.'
    }
    @{
        Path       = 'McpServer/OutlookAI.McpServer.Tests/live-fixtures/live-test-settings.json'
        MustIgnore = $true
        Why        = 'Machine-local. Names real stores, and this repository is public. Testbed/live-test-settings.example.json is the committed shape.'
    }
    @{
        Path       = 'McpServer/OutlookAI.McpServer.Tests/live-fixtures/vm-credentials.json'
        MustIgnore = $true
        Why        = 'A CREDENTIAL. The only place one may live. Never committed, never printed. Testbed/host/Get-GuestCredential.ps1 documents its shape.'
    }
    @{
        Path       = 'McpServer/OutlookAI.McpServer.Tests/live-fixtures/vm-corpus/corpus-vm2.jsonl'
        MustIgnore = $true
        Why        = '2.9 MB of EntryIDs describing one machine mailbox state, regenerated by a build. The four PARAMETERS that reproduce it are committed instead, in Testbed/testbed.json.'
    }
    @{
        Path       = 'Docs/v3-probes/soakfix13-probe-sweep-cost.ps1'
        MustIgnore = $true
        Why        = 'LOST with its scratch directory. Testbed/guest/Measure-SweepCost.ps1 is the reconstruction, written from the shipped sweep source. Docs/corpus-measurement-plan.md step 2 still names the original; the entry stays until that reference is retired.'
    }
    @{
        Path       = 'Docs/v3-probes'
        MustIgnore = $true
        Why        = 'The v3 probe directory: machine-specific and personal data, gitignored on purpose (public repository).'
    }
    @{
        Path       = 'Redist/vstor_redist.exe'
        MustIgnore = $true
        Why        = '40 MB unmodified third-party binary, fetched and hash-verified by release.yml. Committing it would put 40 MB into git history permanently.'
    }
)

$absentIndex = @{}
foreach ($e in $intentionallyAbsent) { $absentIndex[$e.Path.ToLowerInvariant()] = $e }

# ---------------------------------------------------------------------------------------------
# 1. No dangling repository reference.
# ---------------------------------------------------------------------------------------------
$script:Checks++

# Tracked files AND untracked-but-not-ignored ones. The second half matters locally rather than in
# CI: a document added in the same commit as the file it names should be checked at the moment the
# mistake is still cheap, not one commit later.
Push-Location $RepoRoot
try {
    $tracked = @(& git ls-files --cached --others --exclude-standard | Sort-Object -Unique)
}
finally {
    Pop-Location
}
if (-not $tracked -or $tracked.Count -eq 0) {
    Fail 'tracked file list' 'git ls-files returned nothing - this check would prove nothing, so it fails instead.'
    $tracked = @()
}

# A reference counts only when it begins with a real top-level entry of this repository. That is
# what keeps `tools/call`, `T1/SomeTests` and `now/at-anchor` out of the results.
$topLevel = @($tracked | ForEach-Object { ($_ -split '/')[0] } | Sort-Object -Unique)
$prefixPattern = ($topLevel | ForEach-Object { [regex]::Escape($_) }) -join '|'

$scanned = @($tracked | Where-Object {
        $_ -like '*.md' -or ($_ -like 'Testbed/*' -and $_ -like '*.ps1') -or ($_ -like '.github/scripts/*.ps1')
    })

# WHERE A REFERENCE IS ALLOWED TO RESOLVE FROM. Not only the repository root: this project's
# documents legitimately use project-relative shorthand - `Services/MailService.cs` means the one
# under McpServer/OutlookAI.Core, and McpServer/README.md's own links are relative to McpServer/.
# Resolving against every project directory as well as the root is a rule rather than a fudge, and
# it is what keeps this check at zero false positives, which is the difference between a check
# people act on and one they learn to skip.
$resolutionRoots = @('') + @(
    $tracked | Where-Object { $_ -like '*.csproj' } |
        ForEach-Object { Split-Path -Parent $_ } |
        Sort-Object -Unique
)

# Backticked `path/like/this`, or a markdown link target (path/like/this).
$refPattern = '(?:`|\]\()(' + '(?:' + $prefixPattern + ')' + '/[^`\)\s|]*)'

$dangling = @()
$referenced = @{}
foreach ($rel in $scanned) {
    $full = Join-Path $RepoRoot $rel
    if (-not (Test-Path -LiteralPath $full)) { continue }
    $ownDir = Split-Path -Parent $rel
    $text = Get-Content -LiteralPath $full -Raw
    foreach ($m in [regex]::Matches($text, $refPattern)) {
        $p = $m.Groups[1].Value
        # A trailing line or line-range citation is a coordinate, not part of the path.
        $p = [regex]::Replace($p, ':\d+(-\d+)?$', '')
        $p = $p.TrimEnd('.', ',', ':', ';', ')', '/')
        if (-not $p) { continue }
        # A reference with a wildcard or a placeholder is a shape, not a path.
        if ($p -match '[\*\<\>]') { continue }

        # A CLASS OR MEMBER IS NOT AN ARTEFACT. This project's documents write symbols the same
        # way they write paths - `Services/DraftUpdateIntents`, `Com/StoreNaming.UnnamedStorePrefix`,
        # `T1/LiveTierInventoryTests.ProductionOnlyCapabilities` - and no filesystem check can
        # settle those. This check is about files and directories the repository must CONTAIN, so
        # a reference qualifies only when its last segment carries a known file extension, or the
        # whole thing resolves as a directory.
        $leaf = ($p -split '/')[-1]
        $looksLikeFile = $leaf -match '\.(cs|md|ps1|psd1|psm1|json|jsonl|ya?ml|csproj|slnx|props|targets|rsp|txt|xml|iss|png|exe|editorconfig|gitignore|gitattributes)$'
        if (-not $looksLikeFile) {
            $isDirectory = $false
            foreach ($root in (@($ownDir) + $resolutionRoots)) {
                $candidate = if ($root) { Join-Path $RepoRoot (Join-Path $root $p) } else { Join-Path $RepoRoot $p }
                if (Test-Path -LiteralPath $candidate -PathType Container) { $isDirectory = $true; break }
            }
            if (-not $isDirectory -and -not $absentIndex.ContainsKey($p.ToLowerInvariant())) { continue }
        }

        $referenced[$p] = $true
        if ($absentIndex.ContainsKey($p.ToLowerInvariant())) { continue }

        $found = $false
        foreach ($root in (@($ownDir) + $resolutionRoots)) {
            $candidate = if ($root) { Join-Path $RepoRoot (Join-Path $root $p) } else { Join-Path $RepoRoot $p }
            if (Test-Path -LiteralPath $candidate) { $found = $true; break }
        }
        if (-not $found) { $dangling += "$rel -> $p" }
    }
}

if ($dangling.Count -gt 0) {
    Fail 'no dangling repository reference' (
        "$($dangling.Count) reference(s) name a path this repository does not contain:`n        " +
        (($dangling | Sort-Object -Unique) -join "`n        ") +
        "`n    Either add the file, fix the reference, or - if it is absent on purpose - add it to " +
        '$intentionallyAbsent in this script WITH A REASON.')
}
else {
    Pass 'no dangling repository reference' "$($referenced.Count) repository path(s) named across $($scanned.Count) tracked file(s), all present or declared"
}

# ---------------------------------------------------------------------------------------------
# 2. Intentionally-absent paths are still gitignored.
# ---------------------------------------------------------------------------------------------
$script:Checks++

$ignoreProblems = @()
Push-Location $RepoRoot
try {
    foreach ($e in $intentionallyAbsent) {
        if (-not $e.MustIgnore) { continue }
        # Both spellings: a `dir/` rule in .gitignore only matches a path git can tell is a
        # directory, and it cannot tell for a path that does not exist on this checkout.
        & git check-ignore -q -- $e.Path
        $ok = ($LASTEXITCODE -eq 0)
        if (-not $ok) {
            & git check-ignore -q -- ($e.Path.TrimEnd('/') + '/')
            $ok = ($LASTEXITCODE -eq 0)
        }
        if (-not $ok) {
            $ignoreProblems += "$($e.Path) is NOT ignored - the next commit could publish it. Reason it must not be: $($e.Why)"
        }
    }
}
finally {
    Pop-Location
}

if ($ignoreProblems.Count -gt 0) {
    Fail 'declared-absent paths stay ignored' (($ignoreProblems -join "`n        "))
}
else {
    Pass 'declared-absent paths stay ignored' "$(@($intentionallyAbsent | Where-Object { $_.MustIgnore }).Count) path(s) confirmed covered by .gitignore"
}

# ---------------------------------------------------------------------------------------------
# 3. The corpus parameters agree everywhere.
# ---------------------------------------------------------------------------------------------
$script:Checks++

function Read-Text([string] $relative) {
    $full = Join-Path $RepoRoot $relative
    if (-not (Test-Path -LiteralPath $full)) {
        Fail 'corpus parameters agree' "$relative does not exist - this check now proves nothing."
        return $null
    }
    return Get-Content -LiteralPath $full -Raw
}

$paramsOk = $true
$jsonText = Read-Text 'Testbed/testbed.json'
$buildText = Read-Text 'Testbed/guest/Build-Corpus.ps1'
$planText = Read-Text 'Docs/corpus-measurement-plan.md'

if ($null -eq $jsonText -or $null -eq $buildText -or $null -eq $planText) {
    $paramsOk = $false
}
else {
    $corpus = ($jsonText | ConvertFrom-Json).corpus
    $expected = [ordered]@{
        corpusId  = [string]$corpus.corpusId
        seed      = [string]$corpus.seed
        anchor    = [string]$corpus.anchor
        itemCount = [string]$corpus.itemCount
    }

    foreach ($k in @('corpusId', 'seed', 'anchor', 'itemCount')) {
        if ([string]::IsNullOrWhiteSpace($expected[$k])) {
            Fail 'corpus parameters agree' "Testbed/testbed.json corpus.$k is empty."
            $paramsOk = $false
        }
    }

    # The script's defaults. Read as literals rather than by running it: CI has no toolchain for
    # the guest and no reason to execute a testbed script.
    $scriptChecks = @(
        @{ What = 'corpusId';  Pattern = '\$CorpusId\s*=\s*''([^'']+)'''; Expect = $expected['corpusId'] }
        @{ What = 'seed';      Pattern = '\$Seed\s*=\s*(\d+)';            Expect = $expected['seed'] }
        @{ What = 'anchor';    Pattern = '\$Anchor\s*=\s*''([^'']+)''';   Expect = $expected['anchor'] }
        @{ What = 'itemCount'; Pattern = '\$Count\s*=\s*(\d+)';           Expect = $expected['itemCount'] }
    )
    foreach ($c in $scriptChecks) {
        $m = [regex]::Match($buildText, $c.Pattern)
        if (-not $m.Success) {
            Fail 'corpus parameters agree' "Could not find the $($c.What) default in Testbed/guest/Build-Corpus.ps1 - the file changed shape and this check no longer proves anything. Pattern: $($c.Pattern)"
            $paramsOk = $false
        }
        elseif ($m.Groups[1].Value -ne $c.Expect) {
            Fail 'corpus parameters agree' "Testbed/guest/Build-Corpus.ps1 has $($c.What) = '$($m.Groups[1].Value)'; Testbed/testbed.json says '$($c.Expect)'."
            $paramsOk = $false
        }
    }

    # The prose side. One canonical sentence, so there is exactly one place to update and the
    # regex cannot quietly stop matching without failing.
    $prose = "--corpus-id $($expected['corpusId']) --seed $($expected['seed']) --anchor $($expected['anchor']) --count $($expected['itemCount'])"
    if ($planText -notmatch [regex]::Escape($prose)) {
        Fail 'corpus parameters agree' "Docs/corpus-measurement-plan.md does not contain the canonical parameter line '$prose'. The document must quote the committed parameters verbatim at least once, or a reader has no way to know which corpus the numbers on that page describe."
        $paramsOk = $false
    }
}

if ($paramsOk) {
    Pass 'corpus parameters agree' "testbed.json, Build-Corpus.ps1 and corpus-measurement-plan.md all say $($expected['corpusId']) / $($expected['seed']) / $($expected['anchor']) / $($expected['itemCount'])"
}

# ---------------------------------------------------------------------------------------------
# 4. The index and the directory agree.
# ---------------------------------------------------------------------------------------------
$script:Checks++

$readmeText = Read-Text 'Testbed/README.md'
$indexOk = $true
if ($null -eq $readmeText) {
    $indexOk = $false
}
else {
    $onDisk = @($tracked | Where-Object { $_ -like 'Testbed/*' -and $_ -like '*.ps1' } |
            ForEach-Object { $_.Substring('Testbed/'.Length) })

    # Also count what is in the working tree but not yet tracked, so a script added in the same
    # commit as its README row is checked rather than silently skipped.
    $untracked = @(Get-ChildItem -Path (Join-Path $RepoRoot 'Testbed') -Recurse -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
            ForEach-Object { $_.FullName.Substring((Join-Path $RepoRoot 'Testbed').Length + 1).Replace('\', '/') })
    $onDisk = @($onDisk + $untracked | Sort-Object -Unique)

    $listed = @([regex]::Matches($readmeText, '\|\s*`(host/[^`]+\.ps1|guest/[^`]+\.ps1)`\s*\|') |
            ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)

    if ($listed.Count -eq 0) {
        Fail 'testbed index matches the directory' 'The script table in Testbed/README.md matched nothing. It changed shape, and this check no longer proves anything.'
        $indexOk = $false
    }

    $missingFromIndex = @($onDisk | Where-Object { $listed -notcontains $_ })
    $missingFromDisk = @($listed | Where-Object { $onDisk -notcontains $_ })

    if ($missingFromIndex.Count -gt 0) {
        Fail 'testbed index matches the directory' ("Not named by the table in Testbed/README.md section 5: " + ($missingFromIndex -join ', ') + '. A script a rebuilder cannot find is a script that does not exist.')
        $indexOk = $false
    }
    if ($missingFromDisk.Count -gt 0) {
        Fail 'testbed index matches the directory' ("Named by Testbed/README.md but absent: " + ($missingFromDisk -join ', '))
        $indexOk = $false
    }

    if ($indexOk) { Pass 'testbed index matches the directory' "$($onDisk.Count) script(s), all listed and all present" }
}

# ---------------------------------------------------------------------------------------------
# 5. Every testbed script parses.
# ---------------------------------------------------------------------------------------------
$script:Checks++

$parseProblems = @()
$scripts = @(Get-ChildItem -Path (Join-Path $RepoRoot 'Testbed') -Recurse -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
foreach ($s in $scripts) {
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($s.FullName, [ref]$null, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $first = $errors[0]
        $parseProblems += "$($s.Name) line $($first.Extent.StartLineNumber): $($first.Message)"
    }
}
if ($parseProblems.Count -gt 0) {
    Fail 'testbed scripts parse' (($parseProblems -join "`n        "))
}
elseif ($scripts.Count -eq 0) {
    Fail 'testbed scripts parse' 'No scripts found under Testbed/ - either they moved or this check is switched off.'
}
else {
    Pass 'testbed scripts parse' "$($scripts.Count) script(s)"
}

# ---------------------------------------------------------------------------------------------
# 6. Nothing credential-shaped under Testbed/.
# ---------------------------------------------------------------------------------------------
$script:Checks++

# An assignment of a literal to something password-shaped. Placeholders in angle brackets, empty
# strings and reads out of a variable or a file are all fine - what is not fine is a value.
$secretPatterns = @(
    '(?im)^\s*(?!#)[^\r\n]*\b(password|passwd|pwd|secret|apikey|api_key)\b\s*[:=]\s*[''"]([^''"<>\s]{4,})[''"]'
    '(?im)ConvertTo-SecureString\s+[''"][^''"<>]{4,}[''"]'
)
$secretHits = @()
$testbedFiles = @(Get-ChildItem -Path (Join-Path $RepoRoot 'Testbed') -Recurse -File -ErrorAction SilentlyContinue)
foreach ($f in $testbedFiles) {
    $text = Get-Content -LiteralPath $f.FullName -Raw
    foreach ($p in $secretPatterns) {
        foreach ($m in [regex]::Matches($text, $p)) {
            $line = ($text.Substring(0, $m.Index) -split "`n").Count
            # Never echo the match: this output goes into a public build log.
            $secretHits += "$($f.Name) line $line"
        }
    }
}
if ($secretHits.Count -gt 0) {
    Fail 'no credential-shaped literal under Testbed/' (
        "Something that reads as a hard-coded secret is at: " + (($secretHits | Sort-Object -Unique) -join ', ') +
        ". The value is NOT printed here on purpose. Credentials live only in " +
        'McpServer/OutlookAI.McpServer.Tests/live-fixtures/, which is gitignored. If this is a false positive, ' +
        'restate it so it does not read as an assignment of a literal.')
}
else {
    Pass 'no credential-shaped literal under Testbed/' "$($testbedFiles.Count) file(s) scanned"
}

# ---------------------------------------------------------------------------------------------
# 7. The answer-file template still holds placeholders, not values.
# ---------------------------------------------------------------------------------------------
$script:Checks++

$answerToken = '{{GUEST_PASSWORD}}'
$templateRelative = 'Testbed/guest/autounattend.template.xml'
$templateFull = Join-Path $RepoRoot $templateRelative
$answerProblems = @()

if (-not (Test-Path -LiteralPath $templateFull)) {
    $answerProblems += "$templateRelative does not exist. Either the unattended install was removed - in which case delete this check and say so - or the template moved and this check now proves nothing."
}
else {
    $templateText = Get-Content -LiteralPath $templateFull -Raw

    if (-not $templateText.Contains($answerToken)) {
        $answerProblems += "$templateRelative no longer contains the $answerToken placeholder at all. A password had to come from somewhere for the install to work, so assume it is now in the file."
    }

    $doc = New-Object System.Xml.XmlDocument
    $parsed = $true
    try { $doc.LoadXml($templateText) }
    catch {
        $parsed = $false
        $answerProblems += "$templateRelative is not well-formed XML ($($_.Exception.Message)). Windows Setup would ignore it, and this check cannot inspect it."
    }

    if ($parsed) {
        # Namespace-agnostic on purpose: the unattend schema puts everything in a default
        # namespace, and matching on LocalName keeps this working if that ever changes.
        $secretNodes = @($doc.SelectNodes('//*') | Where-Object {
                $_.LocalName -eq 'Value' -and $_.ParentNode -and
                ($_.ParentNode.LocalName -eq 'Password' -or $_.ParentNode.LocalName -eq 'AdministratorPassword')
            })

        if ($secretNodes.Count -lt 2) {
            $answerProblems += "$templateRelative declares $($secretNodes.Count) password value(s). The template needs the local account's and the autologon's, so this check is no longer looking at what it thinks it is."
        }

        foreach ($node in $secretNodes) {
            if ($node.InnerText -ne $answerToken) {
                # The VALUE IS NEVER PRINTED: this output goes into a public build log. The XPath
                # is enough to find it.
                $answerProblems += ("A password value under <{0}> in {1} is not the {2} placeholder. Its content is deliberately not printed here. Put the placeholder back; Testbed/host/New-AnswerFile.ps1 fills it from the gitignored credential at build time." -f
                    $node.ParentNode.ParentNode.LocalName, $templateRelative, $answerToken)
            }
        }
    }
}

# A FILLED answer file must never be tracked. This looks at git's view rather than the disk, so a
# generated one sitting in gitignored scratch is fine and a committed one is not.
$committedAnswerFiles = @($tracked | Where-Object { [IO.Path]::GetFileName($_).ToLowerInvariant() -eq 'autounattend.xml' })
if ($committedAnswerFiles.Count -gt 0) {
    $answerProblems += ("A filled answer file is tracked: " + ($committedAnswerFiles -join ', ') +
        '. That file carries the guest password. It belongs in .work/ only - see Testbed/host/New-AnswerFile.ps1.')
}

if ($answerProblems.Count -gt 0) {
    Fail 'answer-file template holds placeholders, not values' (($answerProblems | Sort-Object -Unique) -join "`n        ")
}
else {
    Pass 'answer-file template holds placeholders, not values' "$templateRelative checked, no filled answer file tracked"
}

# ---------------------------------------------------------------------------------------------
# 8. The live-test settings template holds tokens only, and its fields are the example's.
# ---------------------------------------------------------------------------------------------
$script:Checks++

$exampleRelative = 'Testbed/live-test-settings.example.json'
$settingsTemplateRelative = 'Testbed/live-test-settings.template.json'
$rendererRelative = 'Testbed/host/New-LiveTestSettings.ps1'
$stagerRelative = 'Testbed/host/Publish-LiveTierPayload.ps1'
$settingsProblems = @()

# Every field of a settings-shaped object as (Path, Value), notes skipped, one block deep - which is
# all a settings file has. A value is only ever compared, never printed: if a real store name were
# ever typed into the template, this output goes into a public build log.
function Get-SettingsLeaves($node, [string] $prefix) {
    $leaves = @()
    foreach ($p in $node.PSObject.Properties) {
        if ($p.Name.StartsWith('_')) { continue }
        $path = $p.Name
        if ($prefix) { $path = $prefix + '.' + $p.Name }
        if ($p.Value -is [System.Management.Automation.PSCustomObject]) { $leaves += Get-SettingsLeaves $p.Value $path }
        else { $leaves += [pscustomobject]@{ Path = $path; Value = $p.Value } }
    }
    return $leaves
}

function Read-SettingsJson([string] $relative) {
    $full = Join-Path $RepoRoot $relative
    if (-not (Test-Path -LiteralPath $full)) {
        $script:settingsProblems += "$relative does not exist. Either the rendered-settings route was removed - in which case delete this check and say so - or it moved and this check now proves nothing."
        return $null
    }
    try { return (Get-Content -LiteralPath $full -Raw | ConvertFrom-Json) }
    catch {
        $script:settingsProblems += "$relative is not valid JSON ($($_.Exception.Message))."
        return $null
    }
}

$exampleJson = Read-SettingsJson $exampleRelative
$settingsTemplateJson = Read-SettingsJson $settingsTemplateRelative
$testbedJson = Read-SettingsJson 'Testbed/testbed.json'

$templateLeaves = @()
if ($null -ne $exampleJson -and $null -ne $settingsTemplateJson) {
    $exampleLeaves = @(Get-SettingsLeaves $exampleJson '')
    $templateLeaves = @(Get-SettingsLeaves $settingsTemplateJson '')
    $examplePaths = @($exampleLeaves | ForEach-Object { $_.Path })
    $templatePaths = @($templateLeaves | ForEach-Object { $_.Path })

    $notInTemplate = @($examplePaths | Where-Object { $templatePaths -cnotcontains $_ })
    $notInExample = @($templatePaths | Where-Object { $examplePaths -cnotcontains $_ })
    if ($notInTemplate.Count -gt 0) {
        $settingsProblems += "$settingsTemplateRelative lacks field(s) the example has: $($notInTemplate -join ', '). The two must name the same fields - add the token, and teach $rendererRelative its rule."
    }
    if ($notInExample.Count -gt 0) {
        $settingsProblems += "$settingsTemplateRelative has field(s) the example does not: $($notInExample -join ', '). The example is the documented shape; add the field there first, or take it out of the template."
    }

    foreach ($leaf in $templateLeaves) {
        if (-not ($leaf.Value -is [string]) -or $leaf.Value -cne ('{{' + $leaf.Path + '}}')) {
            # The value is deliberately NOT printed - see Get-SettingsLeaves.
            $settingsProblems += "$settingsTemplateRelative field '$($leaf.Path)' is not the token for its own path. The template holds tokens and nothing else; its content is not printed here."
        }
    }
    foreach ($p in $settingsTemplateJson.PSObject.Properties) {
        if ($p.Name.StartsWith('_') -and $p.Value -is [string] -and $p.Value.Contains('{{')) {
            $settingsProblems += "$settingsTemplateRelative note '$($p.Name)' contains a double brace. A note that spells a token reads as an unreplaced one to anything scanning for them."
        }
    }
}

# The example is what a rebuilder copies, and it is committed to a PUBLIC repository: it must never
# hold a real address (decided 2026-09-24, Q79). So every address-shaped value in it - at any depth,
# in any list, in the _-notes as much as in the fields - must end in the RFC 2606 .invalid domain,
# which cannot resolve. Only WHERE is reported, never the value: a real address found here goes into
# a public build log otherwise.
function Get-AddressLeaks($node, [string] $path) {
    $leaks = @()
    if ($null -eq $node) { return $leaks }
    if ($node -is [string]) {
        foreach ($m in [regex]::Matches($node, '[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+')) {
            if (-not $m.Value.EndsWith('.invalid', [System.StringComparison]::OrdinalIgnoreCase)) { $leaks += $path }
        }
        return $leaks
    }
    if ($node -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $node.PSObject.Properties) {
            $child = $p.Name
            if ($path) { $child = $path + '.' + $p.Name }
            $leaks += Get-AddressLeaks $p.Value $child
        }
        return $leaks
    }
    if ($node -is [System.Array]) {
        for ($i = 0; $i -lt $node.Count; $i++) { $leaks += Get-AddressLeaks $node[$i] ('{0}[{1}]' -f $path, $i) }
    }
    return $leaks
}

$exampleAddressCount = 0
if ($null -ne $exampleJson) {
    $exampleRaw = Get-Content -LiteralPath (Join-Path $RepoRoot $exampleRelative) -Raw
    $exampleAddressCount = [regex]::Matches($exampleRaw, '[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+').Count
    foreach ($where in @(Get-AddressLeaks $exampleJson '' | Select-Object -Unique)) {
        $settingsProblems += "$exampleRelative '$where' holds an address outside the RFC 2606 .invalid domain. The example is committed to a public repository and must never name a real mailbox - use one under .invalid (the value is not printed here)."
    }
    if ($exampleAddressCount -eq 0) {
        $settingsProblems += "$exampleRelative holds no address at all, so the .invalid rule checks nothing - the example's hub is an address by design (tests hand it to NewDraft as one). Either the example changed shape or this pattern stopped matching."
    }
}

# The per-guest values must name exactly the template's fields. A block may be null (none on that
# guest) or a placeholder still waiting to be read off the guest - but never a different shape.
$guestCount = 0
if ($null -ne $testbedJson -and $templateLeaves.Count -gt 0) {
    $section = $testbedJson.PSObject.Properties | Where-Object { $_.Name -ceq 'liveTestSettings' }
    if (-not $section -or -not ($section.Value -is [System.Management.Automation.PSCustomObject])) {
        $settingsProblems += "Testbed/testbed.json has no liveTestSettings object, so $rendererRelative has nothing to render from."
    }
    else {
        $topLevel = @($templateLeaves | ForEach-Object { ($_.Path -split '\.')[0] } | Select-Object -Unique)
        $blocks = @($templateLeaves | Where-Object { $_.Path.Contains('.') } | ForEach-Object { ($_.Path -split '\.')[0] } | Select-Object -Unique)
        foreach ($guest in $section.Value.PSObject.Properties) {
            if ($guest.Name.StartsWith('_')) { continue }
            $guestCount++
            $where = "Testbed/testbed.json liveTestSettings.$($guest.Name)"
            if (-not ($guest.Value -is [System.Management.Automation.PSCustomObject])) {
                $settingsProblems += "$where is not an object."
                continue
            }
            $guestTop = @($guest.Value.PSObject.Properties | Where-Object { -not $_.Name.StartsWith('_') } | ForEach-Object { $_.Name })
            foreach ($name in @($topLevel | Where-Object { $guestTop -cnotcontains $_ })) { $settingsProblems += "$where lacks '$name'." }
            foreach ($name in @($guestTop | Where-Object { $topLevel -cnotcontains $_ })) { $settingsProblems += "$where has '$name', which the template does not." }

            foreach ($block in $blocks) {
                $value = ($guest.Value.PSObject.Properties | Where-Object { $_.Name -ceq $block }).Value
                if ($null -eq $value) { continue }
                if ($value -is [string] -and $value.StartsWith('<FILL')) { continue }
                if (-not ($value -is [System.Management.Automation.PSCustomObject])) {
                    $settingsProblems += "$where.$block must be an object carrying every field, null, or a placeholder."
                    continue
                }
                $want = @($templateLeaves | Where-Object { $_.Path.StartsWith($block + '.') } | ForEach-Object { $_.Path.Substring($block.Length + 1) })
                $have = @($value.PSObject.Properties | Where-Object { -not $_.Name.StartsWith('_') } | ForEach-Object { $_.Name })
                foreach ($name in @($want | Where-Object { $have -cnotcontains $_ })) { $settingsProblems += "$where.$block lacks '$name'." }
                foreach ($name in @($have | Where-Object { $want -cnotcontains $_ })) { $settingsProblems += "$where.$block has '$name', which the template does not." }
            }
        }
        if ($guestCount -eq 0) {
            $settingsProblems += "Testbed/testbed.json liveTestSettings names no guest, so this part of the check proves nothing."
        }
    }
}

# Where the rendered file lands on a guest must be where the suite is built there.
$rendererText = $null
$stagerText = $null
foreach ($pair in @(@($rendererRelative, 'renderer'), @($stagerRelative, 'stager'))) {
    $full = Join-Path $RepoRoot $pair[0]
    if (-not (Test-Path -LiteralPath $full)) {
        $settingsProblems += "$($pair[0]) does not exist, so nothing proves where a rendered settings file lands on a guest."
        continue
    }
    if ($pair[1] -eq 'renderer') { $rendererText = Get-Content -LiteralPath $full -Raw }
    else { $stagerText = Get-Content -LiteralPath $full -Raw }
}
if ($null -ne $rendererText -and $null -ne $stagerText) {
    $rootPattern = '\$(?:script:)?GuestSourceRoot\s*=\s*''([^'']+)'''
    $rendererRoot = [regex]::Match($rendererText, $rootPattern)
    $stagerRoot = [regex]::Match($stagerText, $rootPattern)
    if (-not $rendererRoot.Success -or -not $stagerRoot.Success) {
        $settingsProblems += "Could not find `$GuestSourceRoot in both $rendererRelative and $stagerRelative - one of them changed shape, and this check no longer proves the rendered file lands where the suite is built."
    }
    elseif ($rendererRoot.Groups[1].Value -cne $stagerRoot.Groups[1].Value) {
        $settingsProblems += "$rendererRelative sends the rendered file under '$($rendererRoot.Groups[1].Value)', but $stagerRelative builds the suite under '$($stagerRoot.Groups[1].Value)'. The tier on the guest would not find its settings."
    }
}

# A real settings file names real stores. None may ever be tracked - wherever it was copied to.
$trackedSettings = @($tracked | Where-Object { [IO.Path]::GetFileName($_).ToLowerInvariant() -eq 'live-test-settings.json' })
if ($trackedSettings.Count -gt 0) {
    $settingsProblems += ("A live-test-settings.json is tracked: " + ($trackedSettings -join ', ') +
        '. That file names real stores. Rendered ones belong in .work/, and the real one only in the gitignored live-fixtures directory.')
}

if ($settingsProblems.Count -gt 0) {
    Fail 'live-test settings template holds tokens only, with the example''s fields' (($settingsProblems | Sort-Object -Unique) -join "`n        ")
}
else {
    Pass 'live-test settings template holds tokens only, with the example''s fields' "$($templateLeaves.Count) token(s) matching $exampleRelative, $guestCount guest section(s) in testbed.json of the same shape, the guest root agreeing with the stager, $exampleAddressCount address(es) in the example all under .invalid, no live-test-settings.json tracked"
}

# ---------------------------------------------------------------------------------------------

Write-Host ''
if ($script:Failures.Count -gt 0) {
    Write-Host "FAILED $($script:Failures.Count) of $($script:Checks) check(s):" -ForegroundColor Red
    foreach ($f in $script:Failures) {
        Write-Host ''
        Write-Host "  $f" -ForegroundColor Red
    }
    Write-Host ''
    exit 1
}

Write-Host "All $($script:Checks) testbed checks passed." -ForegroundColor Green
exit 0
