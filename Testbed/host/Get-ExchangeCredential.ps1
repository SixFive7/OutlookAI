#Requires -Version 5.1
<#
.SYNOPSIS
    Reads the Exchange test VM's sign-in credential from the gitignored live-fixtures directory,
    and makes a fresh TOTP code from it. Never prints, logs or writes either.

.DESCRIPTION
    THIS FILE CONTAINS NO CREDENTIAL AND MUST NEVER CONTAIN ONE. This repository is public. It
    names the one place the Exchange VM's credential may live and the shape it takes there, so
    that no other script has an excuse to hold one (decided by the maintainer 2026-10-03, Q108).

    The file:

        McpServer/OutlookAI.McpServer.Tests/live-fixtures/exchange-credentials.json

    gitignored by the `McpServer/**/live-fixtures/` rule, beside vm-credentials.json. Its shape:

        {
          "account":       "<the mailbox's sign-in address>",
          "passwordDpapi": "<ConvertFrom-SecureString of the password>",
          "totpUriDpapi":  "<ConvertFrom-SecureString of the otpauth://totp/... URI>",
          "protection":    "<a note>", "storedUtc": "<when>", "why": "<why>"
        }

    The two secrets are Windows DPAPI blobs in the CurrentUser scope - what ConvertFrom-SecureString
    writes with no -Key - so only the Windows user who stored them, on the machine they were stored
    on, can read them back. Nothing here writes that file: the maintainer stored it.

    WHERE THE SECRETS GO, AND WHERE THEY NEVER GO. They are decrypted only here, in memory, on the
    host, at the moment of a sign-in. The password leaves this script as a SecureString; the TOTP
    secret never leaves it at all - only a six-digit code does, made at the moment it is asked for.
    Neither is ever written to a stream, a log, a file or the console: not by this script, and not by
    its caller, Testbed/host/Invoke-ExchangeSignIn.ps1, which types them into the guest's sign-in
    dialog through the VM's synthetic keyboard, so that no guest process or file holds a copy. On
    the VM the only copy is the token cache Windows keeps for Outlook.

    THE CODE IS RFC 6238 TOTP: HMAC-SHA1 over the 30-second step count since the Unix epoch, dynamic
    truncation (RFC 4226 section 5.3), six digits - what Microsoft's authenticator enrolment issues.
    Built from .NET's own HMACSHA1, nothing else. A URI that asks for anything else - another
    algorithm, digit count or period, or a counter-based (hotp) code - is REFUSED rather than
    approximated: a wrong code is a failed sign-in at best and a locked-out account at worst.

    MODES

      (default)   returns [pscustomobject] @{ Account = <string>; Password = <SecureString> }.
                  The TOTP secret is not part of it.
      -TotpCode   returns [pscustomobject] @{ Code = <6 digits>; ValidUntilUtc = <DateTime> }. If the
                  current code would expire within -MinimumValiditySeconds, it waits for the next
                  step first, so a code is never typed with a second left on it.
      -SelfTest   the published test vectors - RFC 4648 base32, RFC 4226 appendix D, RFC 6238
                  appendix B - and the URI rules. Reads no credential file; runs anywhere.

.PARAMETER RepoRoot
    Repository root. Defaults to two levels above this script. From a worktree, point it at the
    main checkout: the credential is gitignored and lives only there.

.PARAMETER TotpCode
    Return a fresh TOTP code instead of the credential.

.PARAMETER MinimumValiditySeconds
    With -TotpCode: the least time the returned code must still be valid for. Default 10.

.PARAMETER SelfTest
    Run the test vectors and exit 0 (all pass) or 1.

.EXAMPLE
    $cred = & Testbed/host/Get-ExchangeCredential.ps1 -RepoRoot C:\Source\SixFive7\OutlookAI

.EXAMPLE
    (& Testbed/host/Get-ExchangeCredential.ps1 -RepoRoot C:\Source\SixFive7\OutlookAI -TotpCode).Code

.EXAMPLE
    powershell.exe -NoProfile -File Testbed/host/Get-ExchangeCredential.ps1 -SelfTest
#>
[CmdletBinding(DefaultParameterSetName = 'Credential')]
param(
    [Parameter(ParameterSetName = 'Credential')]
    [Parameter(ParameterSetName = 'Totp')]
    [string] $RepoRoot,
    [Parameter(Mandatory = $true, ParameterSetName = 'Totp')] [switch] $TotpCode,
    [Parameter(ParameterSetName = 'Totp')] [ValidateRange(0, 25)] [int] $MinimumValiditySeconds = 10,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)
# Defaults that need $PSScriptRoot are set HERE, not in param(): Windows PowerShell 5.1 leaves it
# empty while param() defaults are evaluated (Q78).
if (-not $PSBoundParameters.ContainsKey('RepoRoot')) { $RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) }

$ErrorActionPreference = 'Stop'

# =============================================================================================
# PURE FUNCTIONS. No file, no credential, no clock unless one is passed in - -SelfTest pins them.
# =============================================================================================

# RFC 4648 section 6 base32, as an otpauth URI carries the secret: case-insensitive, padding and
# spaces ignored. Any other character is refused - a mistyped secret must not decode to a key.
function ConvertFrom-Base32 {
    param([Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Text)

    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'
    $clean = ($Text -replace '[\s=]', '').ToUpperInvariant()
    $bytes = New-Object System.Collections.Generic.List[byte]
    $buffer = 0
    $bits = 0
    foreach ($ch in $clean.ToCharArray()) {
        $value = $alphabet.IndexOf($ch)
        if ($value -lt 0) { throw "Not a base32 character at position $($bytes.Count): the secret is malformed (the character is not shown)." }
        $buffer = (($buffer -shl 5) -bor $value) -band 0xFFFF
        $bits += 5
        if ($bits -ge 8) {
            $bits -= 8
            $bytes.Add([byte](($buffer -shr $bits) -band 0xFF))
        }
    }
    return , $bytes.ToArray()
}

# RFC 4226 section 5.3: HMAC-SHA1 over the 8-byte big-endian counter, dynamic truncation, modulo
# 10^Digits, left-padded with zeros.
function Get-HotpValue {
    param(
        [Parameter(Mandatory = $true)] [byte[]] $Key,
        [Parameter(Mandatory = $true)] [long] $Counter,
        [ValidateSet(6, 8)] [int] $Digits = 6
    )

    $message = New-Object byte[] 8
    $c = $Counter
    for ($i = 7; $i -ge 0; $i--) {
        $message[$i] = [byte]($c -band 0xFF)
        $c = $c -shr 8
    }
    $hmac = New-Object System.Security.Cryptography.HMACSHA1 (, $Key)
    try { $hash = $hmac.ComputeHash($message) } finally { $hmac.Dispose() }

    $offset = $hash[$hash.Length - 1] -band 0x0F
    $binary = (([long]($hash[$offset] -band 0x7F)) -shl 24) -bor
              (([long]$hash[$offset + 1]) -shl 16) -bor
              (([long]$hash[$offset + 2]) -shl 8) -bor
              ([long]$hash[$offset + 3])
    $modulus = [long][math]::Pow(10, $Digits)
    return ($binary % $modulus).ToString(('D' + $Digits), [Globalization.CultureInfo]::InvariantCulture)
}

# RFC 6238: the HOTP value of floor(unixSeconds / period).
function Get-TotpValue {
    param(
        [Parameter(Mandatory = $true)] [byte[]] $Key,
        [Parameter(Mandatory = $true)] [long] $UnixSeconds,
        [int] $Period = 30,
        [ValidateSet(6, 8)] [int] $Digits = 6
    )
    return Get-HotpValue -Key $Key -Counter ([long][math]::Floor($UnixSeconds / $Period)) -Digits $Digits
}

# The second at which the code valid at $UnixSeconds stops being valid.
function Get-TotpStepEnd {
    param([Parameter(Mandatory = $true)] [long] $UnixSeconds, [int] $Period = 30)
    return ([long][math]::Floor($UnixSeconds / $Period) + 1) * $Period
}

# Reads an otpauth://totp/ URI (the Key Uri Format authenticator apps share) into its key and
# parameters, and refuses every parameter this generator does not implement exactly: HMAC-SHA1,
# 6 digits, 30 seconds. The error never quotes the URI - it holds the secret.
function Read-OtpAuthUri {
    param([Parameter(Mandatory = $true)] [string] $Uri)

    $m = [regex]::Match($Uri, '^otpauth://([A-Za-z]+)/[^?]*\?(.*)$')
    if (-not $m.Success) { throw 'The TOTP entry is not an otpauth:// URI (the value is not shown).' }
    if ($m.Groups[1].Value -ne 'totp') { throw "The TOTP entry is an otpauth '$($m.Groups[1].Value)' URI; only 'totp' is supported." }

    $query = @{}
    foreach ($pair in ($m.Groups[2].Value -split '&')) {
        if (-not $pair) { continue }
        $kv = $pair -split '=', 2
        $name = [Uri]::UnescapeDataString($kv[0]).ToLowerInvariant()
        $value = ''
        if ($kv.Count -gt 1) { $value = [Uri]::UnescapeDataString($kv[1]) }
        $query[$name] = $value
    }

    $encoded = $query['secret']
    if ([string]::IsNullOrWhiteSpace($encoded)) { throw 'The TOTP URI carries no key.' }

    $algorithm = 'SHA1'
    if ($query.ContainsKey('algorithm') -and $query['algorithm']) { $algorithm = $query['algorithm'].ToUpperInvariant() }
    if ($algorithm -ne 'SHA1') { throw "The TOTP URI asks for algorithm '$algorithm'; this generator implements HMAC-SHA1 only and refuses rather than guesses." }

    $digits = 6
    if ($query.ContainsKey('digits') -and $query['digits']) { $digits = [int]$query['digits'] }
    if ($digits -ne 6) { throw "The TOTP URI asks for $digits digits; this generator makes 6-digit codes only." }

    $period = 30
    if ($query.ContainsKey('period') -and $query['period']) { $period = [int]$query['period'] }
    if ($period -ne 30) { throw "The TOTP URI asks for a $period-second period; this generator uses 30 seconds only." }

    $key = ConvertFrom-Base32 -Text $encoded
    if ($key.Length -lt 10) { throw "The TOTP key decodes to $($key.Length) byte(s); RFC 4226 requires at least 128 bits." }

    return [pscustomobject]@{ Key = $key; Algorithm = $algorithm; Digits = $digits; Period = $period; Issuer = $query['issuer'] }
}

# SecureString to a plain string for the one moment it is needed. The BSTR is zeroed at once; the
# managed string lives until the garbage collector takes it, which PowerShell cannot prevent.
function ConvertFrom-SecureStringInMemory {
    param([Parameter(Mandatory = $true)] [System.Security.SecureString] $Secure)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# =============================================================================================
# SELF-TEST
# =============================================================================================
if ($SelfTest) {
    Write-Host "Get-ExchangeCredential.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Reads no credential."
    $script:failures = 0
    $script:passes = 0
    function Check {
        param([string] $Name, $Expected, $Actual)
        if ("$Expected" -ceq "$Actual") { $script:passes++ ; Write-Host "  PASS  $Name" }
        else { $script:failures++ ; Write-Host "  FAIL  $Name - expected '$Expected', got '$Actual'" }
    }
    function CheckThrows {
        param([string] $Name, [scriptblock] $Block, [string] $Like)
        try { & $Block | Out-Null; $script:failures++ ; Write-Host "  FAIL  $Name - did not refuse" }
        catch {
            if ($_.Exception.Message -like $Like) { $script:passes++ ; Write-Host "  PASS  $Name" }
            else { $script:failures++ ; Write-Host "  FAIL  $Name - refused with '$($_.Exception.Message)'" }
        }
    }
    $ascii = [Text.Encoding]::ASCII

    # RFC 4648 section 10.
    $base32Vectors = [ordered]@{ '' = ''; 'MY======' = 'f'; 'MZXQ====' = 'fo'; 'MZXW6===' = 'foo'; 'MZXW6YQ=' = 'foob'; 'MZXW6YTB' = 'fooba'; 'MZXW6YTBOI======' = 'foobar' }
    foreach ($k in $base32Vectors.Keys) {
        Check "base32 '$k'" $base32Vectors[$k] ($ascii.GetString((ConvertFrom-Base32 -Text $k)))
    }
    Check 'base32 lower case and spaces' 'foobar' ($ascii.GetString((ConvertFrom-Base32 -Text 'mzxw 6ytb oi')))
    CheckThrows 'base32 refuses a character outside the alphabet' { ConvertFrom-Base32 -Text 'MZXW1' } '*Not a base32 character*'

    # RFC 4226 appendix D: the 20-byte ASCII key "12345678901234567890", counters 0 to 9.
    $rfcKey = $ascii.GetBytes('12345678901234567890')
    $hotp = @('755224', '287082', '359152', '969429', '338314', '254676', '287922', '162583', '399871', '520489')
    for ($i = 0; $i -lt $hotp.Count; $i++) { Check "RFC 4226 HOTP counter $i" $hotp[$i] (Get-HotpValue -Key $rfcKey -Counter $i) }

    # RFC 6238 appendix B, the SHA1 rows: eight digits as published, and the six-digit code the
    # same truncation gives (the published value modulo 10^6).
    $totp = [ordered]@{ '59' = '94287082'; '1111111109' = '07081804'; '1111111111' = '14050471'; '1234567890' = '89005924'; '2000000000' = '69279037'; '20000000000' = '65353130' }
    foreach ($t in $totp.Keys) {
        Check "RFC 6238 TOTP T=$t, 8 digits" $totp[$t] (Get-TotpValue -Key $rfcKey -UnixSeconds ([long]$t) -Digits 8)
        Check "RFC 6238 TOTP T=$t, 6 digits" $totp[$t].Substring(2) (Get-TotpValue -Key $rfcKey -UnixSeconds ([long]$t))
    }
    Check 'step end at T=59' 60 (Get-TotpStepEnd -UnixSeconds 59)
    Check 'step end at T=60' 90 (Get-TotpStepEnd -UnixSeconds 60)

    # The URI: built from parts, with the RFC key, never a real one.
    $rfcKeyBase32 = 'GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ'
    Check 'the RFC key in base32 decodes to the RFC key' '12345678901234567890' ($ascii.GetString((ConvertFrom-Base32 -Text $rfcKeyBase32)))
    $keyParam = '{0}={1}' -f 'secret', $rfcKeyBase32
    $uri = 'otpauth://totp/Example:someone%40example.invalid?' + $keyParam + '&issuer=Example'
    $parsed = Read-OtpAuthUri -Uri $uri
    Check 'URI: defaults are SHA1, 6 digits, 30 s' 'SHA1/6/30' ("$($parsed.Algorithm)/$($parsed.Digits)/$($parsed.Period)")
    Check 'URI: the issuer is read' 'Example' $parsed.Issuer
    Check 'URI: the code at T=59 is the RFC one' '287082' (Get-TotpValue -Key $parsed.Key -UnixSeconds 59)
    Check 'URI: explicit SHA1/6/30 is accepted' 'SHA1' ((Read-OtpAuthUri -Uri ($uri + '&algorithm=sha1&digits=6&period=30')).Algorithm)
    CheckThrows 'URI: SHA256 is refused' { Read-OtpAuthUri -Uri ($uri + '&algorithm=SHA256') } '*HMAC-SHA1 only*'
    CheckThrows 'URI: 8 digits are refused' { Read-OtpAuthUri -Uri ($uri + '&digits=8') } '*6-digit codes only*'
    CheckThrows 'URI: a 60 s period is refused' { Read-OtpAuthUri -Uri ($uri + '&period=60') } '*30 seconds only*'
    CheckThrows 'URI: hotp is refused' { Read-OtpAuthUri -Uri ($uri -replace '^otpauth://totp/', 'otpauth://hotp/') } "*only 'totp'*"
    CheckThrows 'URI: no key is refused' { Read-OtpAuthUri -Uri 'otpauth://totp/Example:x?issuer=Example' } '*carries no key*'
    CheckThrows 'URI: a short key is refused' { Read-OtpAuthUri -Uri ('otpauth://totp/x?' + ('{0}={1}' -f 'secret', 'MZXW6YTB')) } '*at least 128 bits*'
    CheckThrows 'URI: not a URI at all is refused, without quoting it' { Read-OtpAuthUri -Uri 'GEZDGNBVGY3TQOJQ' } '*not shown*'

    # SecureString round trip, in memory.
    $ss = New-Object System.Security.SecureString
    foreach ($ch in 'abc'.ToCharArray()) { $ss.AppendChar($ch) }
    Check 'SecureString reads back in memory' 'abc' (ConvertFrom-SecureStringInMemory -Secure $ss)

    Write-Host ''
    Write-Host "$script:passes passed, $script:failures failed."
    if ($script:failures -gt 0) { exit 1 }
    exit 0
}

# =============================================================================================
# THE CREDENTIAL
# =============================================================================================

# ConvertTo-SecureString is a Security-module cmdlet; 5.1 started from 7 cannot load 7's copy.
. (Join-Path $PSScriptRoot 'OwnEditionModules.ps1')

$path = Join-Path $RepoRoot 'McpServer\OutlookAI.McpServer.Tests\live-fixtures\exchange-credentials.json'
if (-not (Test-Path -LiteralPath $path)) {
    throw @"
No Exchange credential at:
    $path
That directory is gitignored and machine-local, so a fresh clone or a worktree never has it - pass
-RepoRoot <the main checkout>. The maintainer stores it; see the header of this script for its shape.
"@
}

$json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
foreach ($field in @('account', 'passwordDpapi', 'totpUriDpapi')) {
    if (-not $json.PSObject.Properties.Name.Contains($field) -or [string]::IsNullOrWhiteSpace([string]$json.$field)) {
        throw "exchange-credentials.json is missing '$field'. See the header of $PSCommandPath."
    }
}

function ConvertFrom-DpapiField {
    param([string] $Name, [string] $Blob)
    try { return (ConvertTo-SecureString -String $Blob) }
    catch {
        throw ("exchange-credentials.json's '$Name' does not decrypt for this Windows user on this machine. DPAPI " +
            "in the CurrentUser scope opens only for the user who stored it, where it was stored - the maintainer " +
            "stores it again (the value is not shown).")
    }
}

if ($TotpCode) {
    $uriSecure = ConvertFrom-DpapiField -Name 'totpUriDpapi' -Blob $json.totpUriDpapi
    $parsed = Read-OtpAuthUri -Uri (ConvertFrom-SecureStringInMemory -Secure $uriSecure)
    try {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $end = Get-TotpStepEnd -UnixSeconds $now -Period $parsed.Period
        if (($end - $now) -lt $MinimumValiditySeconds) {
            # Too close to the end of the step to be typed and submitted in time: wait for the next.
            Start-Sleep -Seconds ([int]($end - $now) + 1)
            $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
            $end = Get-TotpStepEnd -UnixSeconds $now -Period $parsed.Period
        }
        $code = Get-TotpValue -Key $parsed.Key -UnixSeconds $now -Period $parsed.Period -Digits $parsed.Digits
    }
    finally {
        [Array]::Clear($parsed.Key, 0, $parsed.Key.Length)
    }
    return [pscustomobject]@{
        Code          = $code
        ValidUntilUtc = [DateTimeOffset]::FromUnixTimeSeconds($end).UtcDateTime
    }
}

# The account is shown because it is worth showing ("wrong account" otherwise reads as "wrong
# password"); the password never is.
Write-Verbose "Exchange credential loaded for '$($json.account)'."
return [pscustomobject]@{
    Account  = [string]$json.account
    Password = (ConvertFrom-DpapiField -Name 'passwordDpapi' -Blob $json.passwordDpapi)
}
