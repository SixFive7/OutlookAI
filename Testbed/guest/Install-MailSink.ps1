#Requires -Version 5.1
<#
    ============================================================================================
    THIS SCRIPT HAS NEVER BEEN EXECUTED.
    ============================================================================================

    Written by an agent forbidden to run it: the machine it was written on is the maintainer's
    workstation, with real Outlook and real mail on it, and this script installs a Windows
    service that listens on the SMTP and POP3 ports. Verified by PARSING only - the same check
    .github/scripts/check-testbed-references.ps1 applies to every script under Testbed/. Nothing
    below has run anywhere, and no smtp4dev binary was downloaded, unpacked or started to write
    it. Replace this banner with what it actually did once it has run on a guest, and say which
    of the POP3 assertions passed - three of them exist to settle open questions and two of them
    are expected to FAIL against smtp4dev as shipped. See EXPECTED FAILURES below.

.SYNOPSIS
    Installs the loopback mail sink the live tier's dummy account points at, and then PROVES the
    round trip over raw sockets instead of assuming it.

.DESCRIPTION
    RUN ON THE GUEST. Windows PowerShell 5.1 - no ternary, no `??`, no `-p` on mkdir.
    NEVER run this on the maintainer's workstation. It refuses there; see THE GUARD below.

    WHY A SINK AT ALL. Docs/live-tier-on-the-vm.md section 1.4. The tier profile's POP3 account
    exists because NewDraft resolves an Account by SMTP address and refuses when none matches.
    Pointing that account at an unroutable server makes every send QUEUE and never leave, and
    the Outbox is inside the mandatory zero-artifact sweep - so every run that sent anything
    would fail its own teardown forever, on residue nothing could remove.

    WHY THIS PROVES THE ROUND TRIP RATHER THAN REPORTING TWO OPEN PORTS. The suite's own check
    is deliberately shallow: McpServer/OutlookAI.McpServer.Tests/T2/LiveMailSink.cs connects a
    TcpClient to each port and disconnects. It speaks no protocol. ANYTHING THAT BINDS THOSE TWO
    PORTS PASSES IT - including a sink that accepts mail and cannot hand it back, which is the
    exact failure that design exists to avoid. The real proof otherwise happens only inside a
    120-second arrival wait whose failure mode is a silent timeout pointing nowhere useful. So
    -Verify here speaks SMTP and POP3 itself.

    WHICH SINK, AND THE VERSION FLOOR. smtp4dev (rnwood/smtp4dev, BSD-3-Clause). Its POP3 server
    is REAL but RECENT: it was added by PR #1888 and first appears in a stable release at
    3.11.0. Two later releases matter and are why the floor here is 3.15.0 rather than 3.11.0 -
    3.14.0 added the ability to disable POP3/IMAP by a null port, and 3.15.0 added a POP3 NOOP
    handler (before it, NOOP answers "-ERR Unknown command", and NOOP is what clients use as a
    session keepalive). Note for anyone reading the issue tracker and concluding otherwise:
    issue #155, "Support retrieval of messages using POP3", WAS closed as not planned - in 2022.
    POP3 shipped three years later from an unrelated pull request with no issue behind it. The
    closure is real and it is not the current state.

    EXPECTED FAILURES, AND WHY THEY ARE ASSERTIONS RATHER THAN COMMENTS. Two defects were found
    by reading smtp4dev's POP3 source, not by running it, so this script asserts them rather
    than assuming either way:

      1. TOP IS ADVERTISED IN CAPA AND NOT IMPLEMENTED. The CAPA handler writes TOP
         unconditionally; no TOP handler is registered, so it falls through to "-ERR Unknown
         command". A client that trusts CAPA gets a hard error on a verb the server said it had.
      2. DELE DELETES IMMEDIATELY AND RENUMBERS MID-SESSION. RFC 1939 requires deletion in the
         UPDATE state at QUIT, and requires message numbers to be STABLE for the whole session.
         smtp4dev deletes at once and re-lists the mailbox per command, so after DELE 1 the
         message that was 2 becomes 1. A client issuing DELE 1; DELE 2 deletes the wrong item.

    Neither is necessarily fatal here - the tier PRF sets LeaveOnServer=0x0, so Outlook should
    download-and-delete and never need TOP - but "should" is what this script exists to replace.
    -Verify reports each as a named PASS or FAIL so the answer lands in a log instead of in a
    180-second timeout six weeks later.

    THE PASSWORD PROBLEM, WHICH IS THE ONE REAL COLLISION AND IS NOT FIXABLE HERE.
    smtp4dev's POP3 never checks credentials against anything - but its PASS handler refuses
    when EITHER the username or the password is EMPTY. Meanwhile
    Testbed/guest/tier-profile-forcepst.prf deliberately carries no password key at all, and
    New-TierProfile.ps1 prints the account as "no stored password". Those two facts cannot both
    hold: POP3 has no anonymous mode, and on an unattended guest a credential prompt is a hang
    rather than a prompt. -Verify establishes the SINK half - it tries an empty PASS first, then
    any non-empty one, and REPORTS which the sink accepted. The Outlook half needs a guest.
    The likely shape of the answer is that the password is typed once by hand with "remember
    password" ticked and preserved by the checkpoint, exactly as the two mail accounts already
    are; this script does not assume that and does not do it.

    WHAT IT NEVER DOES. It never starts Outlook, never creates a COM object, never touches MAPI,
    never reads or writes an Outlook profile registry key, and never touches a mail item in any
    store. Everything it does is a service, a directory, a JSON file and some TCP conversations
    with a listener on 127.0.0.1.

    THE GUARD. Same rule and same fail-closed shape as Assert-TestbedGuest in
    Testbed/guest/OutlookMapiInterop.ps1: the cheapest reliable difference between the
    maintainer's machine and a guest is WHO IS LOGGED ON, and the guests autologon as vmadmin.
    It is restated here rather than dot-sourced ON PURPOSE, and the reason is concrete: that
    file compiles Extended MAPI interop with Add-Type at dot-source time, and its own constraint
    2 records that re-dot-sourcing it in a live session throws "type already exists". A mail
    sink installer that dies with a MAPI compile error has failed for a reason that has nothing
    to do with anything it does. The duplication fails SAFE - if the shared guard is ever
    widened this copy stays narrow, and the only way past either is to name the account you
    mean, which is a thing you cannot do by accident.

    STARTING STATE IT EXPECTS.
      * A guest checkpoint where Windows is installed and you are logged on as -ExpectedUser.
      * An ELEVATED session. Registering a Windows service needs it, and the script asserts it
        rather than failing halfway with an access-denied nobody can interpret.
      * OUTLOOK.EXE not running - see WHY IT REFUSES WHILE OUTLOOK IS RUNNING.
      * The sink package STAGED at -PackagePath. It is not downloaded here and it is not in this
        repository - see WHERE THE PACKAGE COMES FROM.
      * Ports -SmtpPort and -Pop3Port free, and not inside a Windows reserved port range.
        Hyper-V and WinNAT genuinely reserve ranges on a VM, and a reservation is not the same
        thing as something listening: netstat shows nothing and the bind fails anyway.

    IDEMPOTENT. Run it twice against the same checkpoint and the second run reports "already"
    for everything and still runs the full verification. Re-running -Execute over an existing
    install rewrites the configuration and restarts the service; it does not stack a second
    service or a second copy of the payload. -Uninstall takes it back to the starting state so a
    checkpoint can be re-taken.

    WHERE THE PACKAGE COMES FROM, AND WHY THAT IS A DECISION RATHER THAN AN OMISSION. This
    project's Dependencies rule forbids "anything a rebuilder would have to download and install
    beyond the media Testbed/MEDIA.md already names as preconditions". A mail sink is exactly
    such a thing, and MEDIA.md does not currently name one. So this script treats the sink the
    way MEDIA.md treats the Windows ISO and the Office Deployment Tool: as STAGED MEDIA, taken
    from -PackagePath, version-pinned and hash-checked, never fetched from the network at
    install time. That keeps a rebuild reproducible and offline, and it puts the decision where
    it belongs - in MEDIA.md, with the maintainer, not inside a script.

    DO NOT INSTALL IT WITH winget. Two reasons, and the first one is a correction: the package
    id is `Rnwood.Smtp4dev`, not `RnwoodLtd.smtp4dev` as Docs/live-tier-on-the-vm.md section 2.7
    currently says, so that command fails outright. The second is structural - winget installs
    this package as a PORTABLE under a version-stamped path in %LOCALAPPDATA% and puts a PATH
    shim in front of it, while service registration bakes the real binary path into the service.
    A staged zip has neither problem.

    THE APPSETTINGS TRAP, WHICH IS THE SILENT ONE. smtp4dev reads appsettings.json from the
    install directory AND THEN from {AppData}/smtp4dev/appsettings.json, and the second one
    WINS. A service registered by --install-service runs as LocalSystem, whose %APPDATA% is
    C:\Windows\System32\config\systemprofile\AppData\Roaming - so a settings file edited in
    either the installer's directory or the interactive user's profile can be silently overridden
    or silently ignored. This script closes that off by appending --nousersettings to the
    service's binary path and VERIFYING it reads back, and by reporting any LocalSystem-profile
    overlay it finds. --install-service itself cannot do this: it hardcodes the binPath as
    `"<exe>" --service` and passes nothing else through.

    WHY IT REFUSES WHILE OUTLOOK IS RUNNING. -Verify submits probe messages and then retrieves
    and deletes them. A running Outlook polls the same sink on its own schedule, so it can take
    a probe first - which fails the verification for a reason that is not a sink fault, AND puts
    the probe into a mailbox. It does NOT kill Outlook: mailbox-safety rule 7 forbids taskkill
    on OUTLOOK.EXE outright.

    THE PROBE MESSAGES, AND WHY THEY CARRY THE LIVE-TIER TAG. If the script dies between
    submitting and deleting, a probe sits in the sink and the next Outlook poll lands it in the
    hub Inbox. Tagging the subject [OutlookAI-McpTest] makes that self-healing: the live tier's
    mandatory zero-artifact sweep matches that tag ordinally across Inbox, Drafts, Sent Items,
    Outbox, Deleted Items and the Sync Issues subtree, so a stray probe is removed by a guard
    that already exists. An untagged stray would be invisible to it and would sit there forever.
    The corpus tag [OutlookAI-Corpus] is deliberately NOT used and must never be - the two
    strings are kept apart so an artifact sweep can never select a corpus item.

.PARAMETER PackagePath
    The staged sink package - Rnwood.Smtp4dev-win-x64-<version>.zip, the SELF-CONTAINED build,
    which needs no .NET runtime on the guest. Required for -Execute. Never downloaded here.

.PARAMETER ExpectedSha256
    Pin the package. When given, the file's hash must match or nothing is unpacked. The default
    is the SHA-256 the winget manifest publishes for the 3.15.0 win-x64 zip; it has NOT been
    verified against a downloaded file on this machine, so a mismatch means "check the hash",
    not "the package is bad". It fails CLOSED, which is the safe direction.

.PARAMETER InstallRoot
    Where the payload is unpacked. The win-x64 zip puts Rnwood.Smtp4dev.exe and appsettings.json
    at its ROOT, so this is also the directory the settings file is written to. Space-free on
    purpose: it goes into a service binary path.

.PARAMETER ServiceName
    The Windows service. smtp4dev registers itself as `Smtp4dev`; changing this only changes
    what is looked for, not what --install-service creates.

.PARAMETER SinkHost
    The loopback address both listeners must bind, and the only address they may bind.

.PARAMETER SmtpPort
    Submission port. Must match the live-test settings' mailSink.submitPort and the PRF SMTPPort.

.PARAMETER Pop3Port
    Retrieval port. Must match mailSink.retrievePort and the PRF POP3Port.

.PARAMETER ImapPort
    0 DISABLES IMAP, which is what this machine wants. Note the translation this parameter
    performs and why it exists: in smtp4dev's own configuration, `null` disables a listener and
    `0` means AUTO-ASSIGN A FREE PORT - the exact opposite. A 0 written straight through would
    silently bind IMAP to an unpredictable port.

.PARAMETER WebUiUrl
    The sink's web UI. Loopback only. No test uses it; it is what a human looks at when the
    round trip fails.

.PARAMETER ProbeAddress
    The fabricated address -Verify submits to and retrieves as. Use the tier account's address:
    that is what proves the catch-all covers it.

.PARAMETER ProbeUser
    The POP3 username -Verify logs in with. With authentication off, smtp4dev ignores it
    entirely and always serves the catch-all mailbox - which is the answer to
    Docs/live-tier-on-the-vm.md section 8 item 3 - but it must still be non-empty.

.PARAMETER ExpectedUser
    The Windows account this may run as. The default IS the guard; do not widen it.

.PARAMETER Execute
    Actually install. Without it nothing is written and the plan is printed instead.

.PARAMETER Verify
    Run the proof only. Writes nothing but its own log and the probe messages, which it removes.

.PARAMETER Uninstall
    Stop and remove the service and delete the install root. Needs -Execute to do anything.

.PARAMETER Force
    Allow -Execute to replace an existing payload.

.EXAMPLE
    .\Install-MailSink.ps1
    .\Install-MailSink.ps1 -PackagePath C:\staging\Rnwood.Smtp4dev-win-x64-3.15.0.zip -Execute
    .\Install-MailSink.ps1 -Verify
    .\Install-MailSink.ps1 -Uninstall -Execute
#>
[CmdletBinding()]
param(
    [string]   $PackagePath,
    [string]   $ExpectedSha256 = '9ED061316772445D54F09B4BD69B97DC79E830312C8F32A8C1717C24FD7068DC',
    [string]   $InstallRoot    = 'C:\OutlookAI-Sink',
    [string]   $ServiceName    = 'Smtp4dev',
    [string]   $SinkHost       = '127.0.0.1',
    [int]      $SmtpPort       = 25,
    [int]      $Pop3Port       = 110,
    [int]      $ImapPort       = 0,
    [string]   $WebUiUrl       = 'http://127.0.0.1:5000',
    [string]   $ProbeAddress   = 'tier@vm.invalid',
    [string]   $ProbeUser      = 'tier',
    [string[]] $ExpectedUser   = @('vmadmin'),
    [string]   $LogPath        = 'C:\OutlookAI-Sink\install-mail-sink.log',
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $Uninstall,
    [switch]   $Force
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Constants. Everything this script writes, runs or asserts is named here and nowhere else, so
# a reader can see the whole blast radius in one place.
# ---------------------------------------------------------------------------------------------

# The live tier's artifact tag, matched ORDINALLY by the zero-artifact sweep. See THE PROBE
# MESSAGES in the banner. It is NOT the corpus tag and must never be.
$probeSubjectTag = '[OutlookAI-McpTest]'

# The windowless executable. The Desktop build creates a WINDOW - which this machine must never
# do - and additionally refuses --install-service, --urls and --basepath, and always picks its
# own port on localhost. Named here as something to refuse rather than left to chance.
$serviceExeName   = 'Rnwood.Smtp4dev.exe'
$forbiddenExeName = 'Rnwood.Smtp4dev.Desktop.exe'

# Where the service reads configuration from. See THE APPSETTINGS TRAP in the banner: this is
# the LOWER-precedence of two files, and --nousersettings is what makes it the only one.
$settingsFileName = 'appsettings.json'

# The overlay that would otherwise win. LocalSystem's roaming profile, not the logged-on user's.
$localSystemOverlay = 'C:\Windows\System32\config\systemprofile\AppData\Roaming\smtp4dev\appsettings.json'

# The flag that makes the install-directory settings file authoritative. --install-service
# cannot pass it, so it is appended to the service binary path afterwards and verified.
$noUserSettingsFlag = '--nousersettings'

# How long any single socket connect or read may take. Loopback: a wait that is not instant is
# a wait that is not going to end.
$socketTimeoutMs = 10000

# How long the whole submit-then-retrieve proof may take. The suite's TIGHTEST arrival deadline
# is 120 s (T2/LiveFreshModeTests, which is stricter than LiveInboxArrival's 180 s), and that
# budget also has to cover Outlook. A sink needing more than 30 s on loopback has already failed.
$roundTripBudgetSeconds = 30

$script:Failures = @()
$script:Lines    = @()
$script:SinkPid  = -1

function Say {
    param([string] $Text)
    $script:Lines += $Text
    Write-Host $Text
}

function Pass {
    param([string] $What, [string] $Detail)
    $suffix = ''
    if ($Detail) { $suffix = " - $Detail" }
    Say ("  OK   {0}{1}" -f $What, $suffix)
}

function Fail {
    param([string] $What, [string] $Why)
    $script:Failures += "$What : $Why"
    Say ("  FAIL {0} - {1}" -f $What, $Why)
}

function Save-Log {
    # A dry run creates NOTHING, including its own log directory. A "dry run" that leaves a
    # directory behind is a dry run whose promise is already false.
    if (-not ($Execute -or $Verify)) { return }

    $dir = Split-Path -Parent $LogPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Set-Content -LiteralPath $LogPath -Value $script:Lines -Encoding UTF8
    Write-Host ''
    Write-Host "Log: $LogPath"
}

# ---------------------------------------------------------------------------------------------
# Guards. First calls in every path that writes, and in -Verify too, because -Verify submits mail.
# ---------------------------------------------------------------------------------------------

function Assert-TestbedGuestLocal {
    $who = $env:USERNAME
    foreach ($candidate in $ExpectedUser) {
        if ($who -eq $candidate) { return }
    }

    throw @"
REFUSING TO RUN. This session is logged on as '$who', which is not one of: $($ExpectedUser -join ', ').

This script registers a Windows SERVICE that listens on the SMTP and POP3 ports, and its
verification submits mail. On the maintainer's workstation that would put an unauthenticated
listener on port $SmtpPort of a machine holding real mail.

The testbed guests autologon as 'vmadmin' (Testbed/README.md section 2). If you are building a
guest whose account has a different name, pass it explicitly:

    -ExpectedUser <that account's username>

Do not 'fix' this by widening the default. The default is the guard.
"@
}

function Assert-Elevated {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw @"
REFUSING TO RUN. This session is not elevated, and registering a Windows service needs it.

Asserted here rather than discovered halfway through: a service install that fails at the
registration step leaves an unpacked payload, no service, and an access-denied message that reads
like a packaging fault.
"@
    }
}

function Assert-OutlookNotRunning {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        throw @"
REFUSING: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')).

This script's verification submits probe messages to the sink and then retrieves and deletes
them. A running Outlook polls the same sink on its own schedule, so it can take a probe first -
which fails the verification for a reason that is not a sink fault, AND puts the probe into a
mailbox.

Close Outlook properly and run this again. DO NOT taskkill it - mailbox-safety rule 7 forbids
that outright.
"@
    }
}

# ---------------------------------------------------------------------------------------------
# Ports. A reserved range is NOT the same thing as something listening: netstat shows nothing
# and the bind fails anyway. Hyper-V and WinNAT reserve ranges on a VM as a matter of course,
# which is exactly the machine this runs on.
# ---------------------------------------------------------------------------------------------

function Get-ExcludedPortRange {
    $ranges = @()
    try {
        $raw = & netsh interface ipv4 show excludedportrange protocol=tcp 2>&1
    }
    catch {
        return $null
    }
    if ($LASTEXITCODE -ne 0) { return $null }

    foreach ($line in $raw) {
        $m = [regex]::Match([string] $line, '^\s*(\d+)\s+(\d+)\s*$')
        if ($m.Success) {
            $ranges += [pscustomobject]@{
                Start = [int] $m.Groups[1].Value
                End   = [int] $m.Groups[2].Value
            }
        }
    }
    return ,$ranges
}

<#
    What is listening on a port, and ON WHICH ADDRESS. The address is the point: loopback
    traffic is not filtered, so a listener that needs an inbound firewall rule is a listener
    bound to 0.0.0.0 - which on a test VM is an open relay.
#>
function Get-ListeningEndpoint {
    param([int] $Port)

    $conn = @()
    try {
        $conn = @(Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction Stop |
                ForEach-Object {
                    [pscustomobject]@{
                        LocalAddress  = [string] $_.LocalAddress
                        OwningProcess = [int] $_.OwningProcess
                    }
                })
    }
    catch {
        # Fall back to netstat. Get-NetTCPConnection needs the NetTCPIP module, and a guest that
        # has lost it should produce a diagnosis rather than a missing-cmdlet error.
        $raw = & netstat -ano -p tcp 2>$null
        foreach ($line in $raw) {
            $m = [regex]::Match([string] $line, '^\s*TCP\s+(\S+):(\d+)\s+\S+\s+LISTENING\s+(\d+)\s*$')
            if ($m.Success -and [int] $m.Groups[2].Value -eq $Port) {
                $conn += [pscustomobject]@{
                    LocalAddress  = $m.Groups[1].Value
                    OwningProcess = [int] $m.Groups[3].Value
                }
            }
        }
    }

    if ($conn.Count -eq 0) { return $null }

    $name = '(exited)'
    try { $name = (Get-Process -Id $conn[0].OwningProcess -ErrorAction Stop).ProcessName } catch { }

    return [pscustomobject]@{
        LocalAddress  = [string] $conn[0].LocalAddress
        OwningProcess = [int] $conn[0].OwningProcess
        ProcessName   = $name
        AllAddresses  = @($conn | ForEach-Object { [string] $_.LocalAddress })
    }
}

function Test-PortUsable {
    param([int] $Port, [string] $What, $Ranges)

    if ($null -eq $Ranges) {
        Say "  NOTE $What port $Port - the reserved-range list could not be read; a bind failure here would look like a packaging fault."
    }
    else {
        foreach ($r in $Ranges) {
            if ($Port -ge $r.Start -and $Port -le $r.End) {
                Fail "$What port $Port is usable" (
                    "It is inside the Windows reserved TCP range $($r.Start)-$($r.End). Nothing is listening there and " +
                    'a bind will still fail. Pick another port and carry it in BOTH the live-test settings mailSink ' +
                    'block and the tier PRF - nothing needs the well-known numbers.')
                return $false
            }
        }
    }

    $listener = Get-ListeningEndpoint -Port $Port
    if ($null -ne $listener -and $listener.OwningProcess -ne $script:SinkPid) {
        Fail "$What port $Port is free" (
            "Something is already listening on $($listener.LocalAddress):$Port (pid $($listener.OwningProcess), " +
            "$($listener.ProcessName)). Stop it or choose another port.")
        return $false
    }

    return $true
}

# ---------------------------------------------------------------------------------------------
# Configuration. Every key below is a REAL smtp4dev ServerOptions key; none is invented. The
# nesting under "ServerOptions" is load-bearing - a key written at the root is read by nothing
# and reported by nothing.
# ---------------------------------------------------------------------------------------------

function New-SinkSetting {
    # PowerShell 5.1 has no ternary. An ordered hashtable keeps the generated file readable and
    # diffable, which matters because the whole point is that a human can check it.
    $imap = $null
    if ($ImapPort -gt 0) { $imap = $ImapPort }

    $options = [ordered]@{
        # Web UI. Loopback only; nothing under test uses it.
        Urls                         = $WebUiUrl

        # SMTP. Not nullable in smtp4dev - submission cannot be switched off, which is fine.
        Port                         = $SmtpPort

        # POP3. null disables; a NUMBER is what enables it. See the -ImapPort help for why 0 is
        # translated rather than written through.
        Pop3Port                     = $Pop3Port
        ImapPort                     = $imap

        # LOOPBACK ONLY, AND THIS DEFAULTS THE OTHER WAY. smtp4dev ships with
        # AllowRemoteConnections = true, i.e. bound to every interface. On a test VM that is an
        # open relay. It is set false here and the binding is asserted afterwards, because a
        # setting that did not take looks exactly like a setting that did.
        AllowRemoteConnections       = $false
        DisableIPv6                  = $false

        # No TLS anywhere. The tier PRF sets POP3UseSSL=0 and SMTPUseSSL=0, so a sink offering
        # STARTTLS is a sink Outlook may try to negotiate with and fail. Note that POP3 has its
        # own key: Pop3TlsMode is INDEPENDENT of TlsMode, and setting only the latter leaves POP3
        # advertising STLS.
        TlsMode                      = 'None'
        Pop3TlsMode                  = 'None'
        SecureConnectionRequired     = $false

        # Authentication. False here means SMTP takes anything, and means POP3 always serves the
        # catch-all mailbox and IGNORES the username. It does NOT mean POP3 will accept an empty
        # password - see THE PASSWORD PROBLEM in the banner.
        AuthenticationRequired       = $false
        SmtpAllowAnyCredentials      = $true

        # Empty on purpose. A catch-all mailbox with Recipients="*" is created automatically as
        # the last mailbox, so there is nothing to provision per address and nothing to
        # re-provision after a checkpoint restore. Declaring one here would only give it a
        # different name to get wrong.
        Users                        = @()
        Mailboxes                    = @()

        # A file database at a path this script chose, rather than the relative default. The
        # alternative is "" for in-memory, which would also work and would drop everything on a
        # restart - rejected because it makes the restart assertion below vacuous, and because a
        # deterministic path is what a checkpoint captures.
        Database                     = (Join-Path $InstallRoot 'sink.db')

        # Retention. smtp4dev prunes silently at its default of 100 per mailbox. The tier deletes
        # every message it retrieves (the PRF sets LeaveOnServer=0x0), so this is headroom rather
        # than a working limit - but a silent prune is worth not having.
        NumberOfMessagesToKeep       = 500
    }

    return [ordered]@{ ServerOptions = $options }
}

# ---------------------------------------------------------------------------------------------
# The socket layer. Raw SMTP and POP3, because the whole point of -Verify is to speak the
# protocols the suite's own probe does not.
# ---------------------------------------------------------------------------------------------

<#
    Connects with a REAL timeout. TcpClient.Connect() blocks for the OS connect timeout, and a
    setup script that hangs is worse than one that fails.
#>
function Connect-SinkSocket {
    param([string] $HostName, [int] $Port)

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($socketTimeoutMs)) {
            $client.Close()
            throw "Nothing answered ${HostName}:${Port} within $socketTimeoutMs ms."
        }
        $client.EndConnect($iar)
    }
    catch {
        $client.Close()
        throw
    }

    $client.ReceiveTimeout = $socketTimeoutMs
    $client.SendTimeout    = $socketTimeoutMs

    $stream = $client.GetStream()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII)
    $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::ASCII)
    $writer.NewLine   = "`r`n"
    $writer.AutoFlush = $true

    return [pscustomobject]@{ Client = $client; Reader = $reader; Writer = $writer }
}

function Close-SinkSocket {
    param($Socket)
    if ($null -eq $Socket) { return }
    try { $Socket.Writer.Dispose() } catch { }
    try { $Socket.Reader.Dispose() } catch { }
    try { $Socket.Client.Close() }   catch { }
}

<#
    One SMTP reply, which may be several lines. A continuation line has '-' in column 4; the
    final one has a space. The whole thing is returned so a failure can quote what the server
    actually said rather than just its code.
#>
function Read-SmtpReply {
    param($Socket)

    $lines = @()
    while ($true) {
        $line = $Socket.Reader.ReadLine()
        if ($null -eq $line) { throw 'The SMTP server closed the connection mid-reply.' }
        $lines += $line
        if ($line.Length -lt 4 -or $line[3] -ne '-') { break }
    }

    $last = $lines[$lines.Count - 1]
    $code = 0
    if ($last.Length -ge 3) { [void][int]::TryParse($last.Substring(0, 3), [ref] $code) }

    return [pscustomobject]@{ Code = $code; Text = ($lines -join ' | ') }
}

function Invoke-SmtpCommand {
    param($Socket, [string] $Command, [int] $ExpectCode)

    $Socket.Writer.WriteLine($Command)
    $reply = Read-SmtpReply -Socket $Socket
    if ($reply.Code -ne $ExpectCode) {
        throw "SMTP '$Command' expected $ExpectCode, got $($reply.Code): $($reply.Text)"
    }
    return $reply
}

function Read-Pop3Reply {
    param($Socket)

    $line = $Socket.Reader.ReadLine()
    if ($null -eq $line) { throw 'The POP3 server closed the connection mid-reply.' }
    return [pscustomobject]@{ Ok = $line.StartsWith('+OK'); Text = $line }
}

function Invoke-Pop3Command {
    param($Socket, [string] $Command, [switch] $AllowError)

    $Socket.Writer.WriteLine($Command)
    $reply = Read-Pop3Reply -Socket $Socket
    if (-not $reply.Ok -and -not $AllowError) {
        # Never echo a PASS argument: this log is read out loud and pasted into issues.
        $shown = $Command -replace '^(PASS)\s.*$', '$1 <not shown>'
        throw "POP3 '$shown' failed: $($reply.Text)"
    }
    return $reply
}

<#
    A POP3 multi-line response, UNSTUFFED. RFC 1939 section 3: the body ends at a line containing
    only '.', and any body line that began with '.' was sent with an extra one prepended. Getting
    this wrong in the CLIENT would make a correctly-behaving server look broken, which is the
    opposite of what this script is for - so it is done explicitly, and whether the SERVER did
    its half is one of the assertions below.
#>
function Read-Pop3MultiLine {
    param($Socket)

    $lines = @()
    while ($true) {
        $line = $Socket.Reader.ReadLine()
        if ($null -eq $line) { throw 'The POP3 server closed the connection inside a multi-line response.' }
        if ($line -eq '.') { break }
        if ($line.StartsWith('.')) { $line = $line.Substring(1) }
        $lines += $line
    }
    return ,$lines
}

# ---------------------------------------------------------------------------------------------
# The probe message. Deliberately awkward in the ways a POP3 implementation gets wrong, because
# a probe that only proves "hello world survives" proves nothing about Outlook's mail.
# ---------------------------------------------------------------------------------------------

function New-ProbeMessage {
    param([string] $Address, [string] $Marker)

    $subject  = "$probeSubjectTag sink probe $Marker"
    $date     = (Get-Date).ToUniversalTime().ToString(
        'ddd, dd MMM yyyy HH:mm:ss +0000', [Globalization.CultureInfo]::InvariantCulture)
    $boundary = "sinkprobe$Marker"

    # Base64 computed at run time so the literal and the assertion cannot drift apart.
    $attachmentB64 = [Convert]::ToBase64String(
        [System.Text.Encoding]::ASCII.GetBytes('OutlookAI sink probe attachment'))

    $lines = @(
        "From: <$Address>"
        "To: <$Address>"
        "Subject: $subject"
        "Date: $date"
        "Message-ID: <$Marker@vm.invalid>"
        'MIME-Version: 1.0'
        "Content-Type: multipart/mixed; boundary=`"$boundary`""
        ''
        "--$boundary"
        'Content-Type: text/plain; charset=us-ascii'
        ''
        # (1) A line that is EXACTLY a period. Un-stuffed by a broken server this terminates the
        #     DATA block early, and the message is truncated silently and only sometimes.
        '.'
        # (2) A line that BEGINS with a period - the classic dot-stuffing bug.
        '.leading period must survive'
        # (3) Two leading periods, which must come back as two.
        '..two leading periods must survive'
        "marker $Marker"
        ''
        "--$boundary"
        'Content-Type: application/octet-stream; name="probe.bin"'
        'Content-Transfer-Encoding: base64'
        'Content-Disposition: attachment; filename="probe.bin"'
        ''
        $attachmentB64
        ''
        "--$boundary--"
    )

    return [pscustomobject]@{
        Subject       = $subject
        Marker        = $Marker
        Lines         = $lines
        AttachmentB64 = $attachmentB64
    }
}

<#
    Submits one probe over SMTP, dot-stuffing on the way out. Throws on any unexpected reply: a
    partial submission is not a state worth continuing from.
#>
function Submit-ProbeMessage {
    param([string] $Address, $Message)

    $socket = $null
    try {
        $socket = Connect-SinkSocket -HostName $SinkHost -Port $SmtpPort

        $greeting = Read-SmtpReply -Socket $socket
        if ($greeting.Code -ne 220) { throw "SMTP greeting was not 220: $($greeting.Text)" }

        $socket.Writer.WriteLine('EHLO localhost')
        $ehlo = Read-SmtpReply -Socket $socket
        if ($ehlo.Code -ne 250) {
            Say "  NOTE EHLO was refused ($($ehlo.Code)); falling back to HELO."
            [void](Invoke-SmtpCommand -Socket $socket -Command 'HELO localhost' -ExpectCode 250)
        }

        [void](Invoke-SmtpCommand -Socket $socket -Command "MAIL FROM:<$Address>" -ExpectCode 250)
        [void](Invoke-SmtpCommand -Socket $socket -Command "RCPT TO:<$Address>"   -ExpectCode 250)
        [void](Invoke-SmtpCommand -Socket $socket -Command 'DATA'                 -ExpectCode 354)

        foreach ($line in $Message.Lines) {
            $out = $line
            if ($out.StartsWith('.')) { $out = '.' + $out }
            $socket.Writer.WriteLine($out)
        }
        $socket.Writer.WriteLine('.')

        $accepted = Read-SmtpReply -Socket $socket
        if ($accepted.Code -ne 250) {
            throw "The sink refused the message at end-of-DATA: $($accepted.Text)"
        }

        $socket.Writer.WriteLine('QUIT')
        return $accepted.Text
    }
    finally {
        Close-SinkSocket -Socket $socket
    }
}

<#
    Logs in and returns the open socket plus WHICH credential shape worked. The order is the
    point: the tier PRF stores no password, so an empty PASS is the shape Outlook is nearest to,
    and if the sink demands more than that it is a finding rather than a detail.

    smtp4dev's PASS handler refuses when either field is empty and otherwise checks nothing, so
    the expected outcome here is that the empty attempt FAILS and the non-empty one succeeds.
    That is reported rather than hidden, because it is half the answer to a question the runbook
    has open.
#>
function Connect-Pop3Authenticated {
    $socket = Connect-SinkSocket -HostName $SinkHost -Port $Pop3Port

    $greeting = Read-Pop3Reply -Socket $socket
    if (-not $greeting.Ok) {
        Close-SinkSocket -Socket $socket
        throw "POP3 greeting was not +OK: $($greeting.Text)"
    }

    $user = Invoke-Pop3Command -Socket $socket -Command "USER $ProbeUser" -AllowError
    if (-not $user.Ok) {
        Close-SinkSocket -Socket $socket
        throw ("The sink rejected USER '$ProbeUser': $($user.Text). With authentication off it should accept any " +
            'non-empty username and serve the catch-all mailbox regardless.')
    }

    # Shape 1: an EMPTY PASS - what an account with nothing stored is nearest to.
    $empty = Invoke-Pop3Command -Socket $socket -Command 'PASS ' -AllowError
    if ($empty.Ok) {
        return [pscustomobject]@{ Socket = $socket; Credential = 'empty PASS' }
    }

    # A server that refused PASS may or may not still be in AUTHORIZATION state. Reconnecting is
    # the only shape correct against both, and costs one loopback connect.
    Close-SinkSocket -Socket $socket
    $socket = Connect-SinkSocket -HostName $SinkHost -Port $Pop3Port
    [void](Read-Pop3Reply -Socket $socket)
    [void](Invoke-Pop3Command -Socket $socket -Command "USER $ProbeUser")

    # Shape 2: any non-empty value. Not a credential - smtp4dev checks nothing - it is the probe
    # for "does this server need the field populated at all".
    $any = Invoke-Pop3Command -Socket $socket -Command "PASS $ProbeUser" -AllowError
    if ($any.Ok) {
        return [pscustomobject]@{ Socket = $socket; Credential = 'any non-empty PASS' }
    }

    Close-SinkSocket -Socket $socket
    throw @"
The sink accepted USER '$ProbeUser' and refused BOTH an empty PASS and a non-empty one.

That is decisive and it is bad news: POP3 has no anonymous mode, Testbed/guest/tier-profile-forcepst.prf
deliberately carries no password key, and on an unattended guest a credential prompt is a hang
rather than a prompt. Either the sink is configured to require real credentials - check
AuthenticationRequired in $InstallRoot\$settingsFileName - or this is not the sink this script
was written for.
"@
}

# ---------------------------------------------------------------------------------------------
# Verification. Every check has a name; failures are COLLECTED rather than thrown, so one run
# reports all the faults instead of the first one.
# ---------------------------------------------------------------------------------------------

function Test-ListenerBinding {
    param([int] $Port, [string] $What)

    $endpoint = Get-ListeningEndpoint -Port $Port
    if ($null -eq $endpoint) {
        Fail "$What listener is up" (
            "Nothing is listening on port $Port. The service can be Running and misconfigured at the same time - " +
            "check $InstallRoot\$settingsFileName, and remember that a settings file in a profile's " +
            'smtp4dev directory overrides it unless the service runs with ' + $noUserSettingsFlag + '.')
        return
    }

    Pass "$What listener is up" "$($endpoint.LocalAddress):$Port (pid $($endpoint.OwningProcess), $($endpoint.ProcessName))"

    $wide = @($endpoint.AllAddresses | Where-Object { $_ -eq '0.0.0.0' -or $_ -eq '::' -or $_ -eq '[::]' })
    if ($wide.Count -gt 0) {
        Fail "$What listener is loopback-only" (
            "It is bound to $($wide -join ', '), not $SinkHost. smtp4dev ships with AllowRemoteConnections = TRUE, so " +
            'this is what an unconfigured install looks like, and on a test VM it is an open relay reachable from the ' +
            'host. Set AllowRemoteConnections false and restart. Do NOT "fix" it by adding a firewall rule.')
    }
    else {
        Pass "$What listener is loopback-only" ($endpoint.AllAddresses -join ', ')
    }
}

function Test-NoInboundFirewallRule {
    # Not a fault on its own - a rule may predate this script - but a rule naming the sink is
    # evidence somebody worked around a 0.0.0.0 bind instead of fixing it.
    try {
        $rules = @(Get-NetFirewallRule -Direction Inbound -Enabled True -ErrorAction Stop |
                Where-Object { $_.DisplayName -like "*$ServiceName*" -or $_.DisplayName -like '*mtp4dev*' })
    }
    catch {
        Say '  NOTE the inbound firewall rules could not be read; a rule created for the sink would not be reported.'
        return
    }

    if ($rules.Count -eq 0) {
        Pass 'no inbound firewall rule for the sink' 'none found, which is correct - loopback traffic is not filtered'
    }
    else {
        Fail 'no inbound firewall rule for the sink' (
            'These inbound rules name it: ' + (($rules | ForEach-Object { $_.DisplayName }) -join ', ') +
            '. A sink that needs one is bound to the wrong address; remove the rule and fix the binding.')
    }
}

<#
    CAPA against reality. smtp4dev writes TOP into its capability list unconditionally and
    registers no TOP handler, so this is expected to FAIL - and it is asserted rather than
    assumed because "expected" is exactly the word this script exists to remove. A client that
    trusts CAPA and issues TOP gets "-ERR Unknown command" on a verb the server advertised.
#>
function Test-Pop3Capabilities {
    param($Socket)

    $capa = Invoke-Pop3Command -Socket $Socket -Command 'CAPA' -AllowError
    if (-not $capa.Ok) {
        Say '  NOTE CAPA is not supported; capability claims cannot be checked against behaviour.'
        return
    }

    $caps = Read-Pop3MultiLine -Socket $Socket
    Say "  ...  CAPA advertises: $($caps -join ', ')"

    $advertisesTop = @($caps | Where-Object { $_ -match '^TOP\b' }).Count -gt 0
    if (-not $advertisesTop) {
        Pass 'CAPA does not over-claim TOP' 'TOP is not advertised, so no client will try it'
        return
    }

    $top = Invoke-Pop3Command -Socket $Socket -Command 'TOP 1 0' -AllowError
    if ($top.Ok) {
        [void](Read-Pop3MultiLine -Socket $Socket)
        Pass 'advertised TOP actually works' 'CAPA and behaviour agree'
    }
    else {
        Fail 'advertised TOP actually works' (
            "CAPA advertises TOP and 'TOP 1 0' answered '$($top.Text)'. This is a known smtp4dev defect, not a " +
            'configuration fault: the capability list is written unconditionally and no TOP handler is registered. ' +
            'It matters only if Outlook uses TOP - which it should not here, because the tier PRF sets ' +
            'LeaveOnServer=0x0 so mail is downloaded and deleted. If a guest ever shows Outlook failing to fetch, ' +
            'this is the first thing to suspect.')
    }
}

function Invoke-RoundTripProof {
    $marker  = [Guid]::NewGuid().ToString('N').Substring(0, 12)
    $message = New-ProbeMessage -Address $ProbeAddress -Marker $marker
    $started = Get-Date

    Say ''
    Say "Round trip: submitting to $SinkHost`:$SmtpPort and retrieving from $SinkHost`:$Pop3Port."

    try {
        $accepted = Submit-ProbeMessage -Address $ProbeAddress -Message $message
        Pass 'the sink accepts a loopback submission' $accepted
    }
    catch {
        Fail 'the sink accepts a loopback submission' $_.Exception.Message
        return
    }

    $session = $null
    try {
        $session = Connect-Pop3Authenticated
        Pass 'POP3 lets the tier account in' "logged in as '$ProbeUser' with $($session.Credential)"
        if ($session.Credential -ne 'empty PASS') {
            Say "  NOTE the sink REQUIRED a non-empty password. The tier PRF stores none, so Outlook will prompt"
            Say '       unless one is entered by hand once and remembered. That is a real step, not a detail -'
            Say '       on an unattended guest a credential prompt is a hang. See section 8 item 3 of the runbook.'
        }
    }
    catch {
        Fail 'POP3 lets the tier account in' $_.Exception.Message
        return
    }

    $socket = $session.Socket
    try {
        Test-Pop3Capabilities -Socket $socket

        $stat = Invoke-Pop3Command -Socket $socket -Command 'STAT'
        $m = [regex]::Match($stat.Text, '^\+OK\s+(\d+)\s+(\d+)')
        if (-not $m.Success) {
            Fail 'STAT reports the submitted message' "Could not parse the STAT reply: $($stat.Text)"
            return
        }

        $count  = [int] $m.Groups[1].Value
        $octets = [int] $m.Groups[2].Value
        if ($count -ne 1) {
            Fail 'STAT reports the submitted message' (
                "STAT says $count message(s), expected exactly 1. More than one means a previous run left residue in " +
                'the sink; zero means the submission was accepted and dropped, which is the failure this whole script ' +
                'exists to catch.')
            # Drain before leaving. Returning here with mail still in the sink is how this run's
            # probe ends up in the hub Inbox at the next poll - which is the one side effect this
            # script must never leave behind.
            Clear-SinkMailbox -Socket $socket
            return
        }
        if ($octets -le 0) {
            Fail 'STAT reports a non-zero octet count' "STAT says $octets octets for 1 message."
        }
        else {
            # Deliberately a lower-bound check, not equality: smtp4dev announces the STORED byte
            # length and then transmits added stuffing dots, so the announced count is smaller
            # than what arrives. RFC 1939 wants it exact; most clients read to the terminator and
            # never notice. Asserting equality here would fail on a sink that works.
            Pass 'STAT reports a non-zero octet count' "1 message, $octets octets announced"
        }

        $uidl = Invoke-Pop3Command -Socket $socket -Command 'UIDL 1' -AllowError
        if ($uidl.Ok -and $uidl.Text -match '^\+OK\s+1\s+(\S+)') {
            Pass 'UIDL answers for the message' $Matches[1]
        }
        else {
            Fail 'UIDL answers for the message' (
                "UIDL 1 returned '$($uidl.Text)'. Outlook keys its already-seen state on UIDL; without it, whether " +
                'mail re-downloads is undefined.')
        }

        $null = Invoke-Pop3Command -Socket $socket -Command 'RETR 1'
        $body = Read-Pop3MultiLine -Socket $socket
        $text = $body -join "`n"

        if ($text.Contains($message.Subject)) {
            Pass 'RETR returns the message with its subject intact' "$($body.Count) line(s)"
        }
        else {
            Fail 'RETR returns the message with its subject intact' (
                "The retrieved message does not contain the submitted subject. $($body.Count) line(s) came back.")
        }

        $bare   = @($body | Where-Object { $_ -eq '.' }).Count
        $single = @($body | Where-Object { $_ -eq '.leading period must survive' }).Count
        $double = @($body | Where-Object { $_ -eq '..two leading periods must survive' }).Count

        if ($bare -eq 1 -and $single -eq 1 -and $double -eq 1) {
            Pass 'dot-stuffing survives the round trip' 'a bare period, one leading period and two all came back unchanged'
        }
        else {
            Fail 'dot-stuffing survives the round trip' (
                "bare-period lines=$bare (expected 1), single-leading=$single (expected 1), double-leading=$double " +
                '(expected 1). This truncates mail silently and intermittently, which is the single worst failure ' +
                'shape available to a test sink.')
        }

        if ($text.Contains($message.AttachmentB64)) {
            Pass 'the base64 attachment part survives' 'byte-identical'
        }
        else {
            Fail 'the base64 attachment part survives' (
                'The base64 part did not come back intact. The live tier sends mail with attachments; a sink that ' +
                'rewrites MIME is a sink whose failures land in the attachment tests.')
        }

        [void](Invoke-Pop3Command -Socket $socket -Command 'DELE 1')
        [void](Invoke-Pop3Command -Socket $socket -Command 'QUIT' -AllowError)
    }
    finally {
        Close-SinkSocket -Socket $socket
    }

    try {
        $after = Connect-Pop3Authenticated
        try {
            $stat2 = Invoke-Pop3Command -Socket $after.Socket -Command 'STAT'
            if ($stat2.Text -match '^\+OK\s+0\s') {
                Pass 'DELE is honoured' 'the mailbox is empty on reconnect'
            }
            else {
                Fail 'DELE is honoured' (
                    "STAT still reports '$($stat2.Text)' after DELE and QUIT. The tier PRF sets LeaveOnServer=0x0 and " +
                    'relies on the sink draining; a sink that ignores DELE re-delivers the same mail forever and leaves ' +
                    "this run's probe behind.")
            }
            [void](Invoke-Pop3Command -Socket $after.Socket -Command 'QUIT' -AllowError)
        }
        finally {
            Close-SinkSocket -Socket $after.Socket
        }
    }
    catch {
        Fail 'DELE is honoured' $_.Exception.Message
    }

    $elapsed = ((Get-Date) - $started).TotalSeconds
    if ($elapsed -le $roundTripBudgetSeconds) {
        Pass 'the round trip fits the budget' ("{0:N1} s, budget {1} s" -f $elapsed, $roundTripBudgetSeconds)
    }
    else {
        Fail 'the round trip fits the budget' (
            ("{0:N1} s against a {1} s budget. The suite's tightest arrival deadline is 120 s " -f $elapsed, $roundTripBudgetSeconds) +
            '(T2/LiveFreshModeTests) and that has to cover Outlook as well as the sink.')
    }
}

<#
    Two messages, one DELE, and the question RFC 1939 section 5 settles: are message numbers
    STABLE for the whole session? smtp4dev deletes immediately rather than at QUIT and re-lists
    the mailbox on every command, so the message that was 2 is expected to become 1 - which means
    a client issuing DELE 1 then DELE 2 deletes the wrong item and leaves the right one.

    This matters here because more than one seeded mail can be in the sink at once, and because
    the failure is silent: the wrong mail is deleted, the arrival wait matches a stale subject,
    and nothing anywhere says a number moved. It is asserted so the answer lands in a log.

    It drains the mailbox to zero whatever it finds, because a probe left behind is a probe that
    turns up in the hub Inbox later.
#>
function Invoke-OrdinalStabilityProof {
    Say ''
    Say 'Ordinal stability: two messages, one DELE, does the survivor keep its number?'

    $markerA = [Guid]::NewGuid().ToString('N').Substring(0, 12)
    $markerB = [Guid]::NewGuid().ToString('N').Substring(0, 12)

    try {
        [void](Submit-ProbeMessage -Address $ProbeAddress -Message (New-ProbeMessage -Address $ProbeAddress -Marker $markerA))
        [void](Submit-ProbeMessage -Address $ProbeAddress -Message (New-ProbeMessage -Address $ProbeAddress -Marker $markerB))
    }
    catch {
        Fail 'two messages can be submitted back to back' $_.Exception.Message
        return
    }

    $session = $null
    try {
        $session = Connect-Pop3Authenticated
    }
    catch {
        Fail 'ordinal stability can be measured' $_.Exception.Message
        return
    }

    $socket = $session.Socket
    try {
        $stat = Invoke-Pop3Command -Socket $socket -Command 'STAT'
        if ($stat.Text -notmatch '^\+OK\s+2\s') {
            Fail 'both messages are held' "STAT reports '$($stat.Text)', expected 2 messages."
            Clear-SinkMailbox -Socket $socket
            return
        }
        Pass 'both messages are held' '2 messages'

        $u1 = Invoke-Pop3Command -Socket $socket -Command 'UIDL 1' -AllowError
        $u2 = Invoke-Pop3Command -Socket $socket -Command 'UIDL 2' -AllowError
        if (-not ($u1.Ok -and $u2.Ok)) {
            Fail 'ordinal stability can be measured' "UIDL 1 => '$($u1.Text)', UIDL 2 => '$($u2.Text)'."
            Clear-SinkMailbox -Socket $socket
            return
        }
        $uid2 = ($u2.Text -split '\s+')[2]

        [void](Invoke-Pop3Command -Socket $socket -Command 'DELE 1')

        # RFC 1939: after DELE 1 the survivor is STILL message 2, and message 1 answers -ERR.
        $afterAt2 = Invoke-Pop3Command -Socket $socket -Command 'UIDL 2' -AllowError
        $afterAt1 = Invoke-Pop3Command -Socket $socket -Command 'UIDL 1' -AllowError

        if ($afterAt2.Ok -and ($afterAt2.Text -split '\s+')[2] -eq $uid2) {
            Pass 'message numbers are stable within a session' 'the survivor is still message 2, as RFC 1939 requires'
        }
        elseif ($afterAt1.Ok -and ($afterAt1.Text -split '\s+')[2] -eq $uid2) {
            Fail 'message numbers are stable within a session' (
                'After DELE 1 the surviving message has become message 1. RFC 1939 section 5 requires numbers to stay ' +
                'fixed for the whole session, and smtp4dev deletes immediately and re-lists per command instead of ' +
                'deferring to QUIT. CONSEQUENCE: a client that issues DELE 1 then DELE 2 deletes the wrong item and ' +
                'leaves the right one. It bites only when more than one message is in the sink at once - which happens ' +
                'whenever two seeded tests overlap or a previous run left residue. Keep Outlook fetching one message ' +
                'at a time, or accept the risk knowingly.')
        }
        else {
            Fail 'message numbers are stable within a session' (
                "After DELE 1, UIDL 1 => '$($afterAt1.Text)' and UIDL 2 => '$($afterAt2.Text)'. Neither matches the " +
                'survivor recorded before the delete, so this sink does something a third way.')
        }

        Clear-SinkMailbox -Socket $socket
        [void](Invoke-Pop3Command -Socket $socket -Command 'QUIT' -AllowError)
    }
    finally {
        Close-SinkSocket -Socket $socket
    }

    # Whatever happened above, the sink MUST be empty now: a probe left behind turns up in the
    # hub Inbox at the next poll.
    try {
        $final = Connect-Pop3Authenticated
        try {
            $stat = Invoke-Pop3Command -Socket $final.Socket -Command 'STAT'
            if ($stat.Text -match '^\+OK\s+0\s') {
                Pass 'the sink is drained afterwards' 'no probe left behind'
            }
            else {
                Fail 'the sink is drained afterwards' (
                    "STAT reports '$($stat.Text)'. Probe messages are still in the sink and will arrive in the hub " +
                    'Inbox at the next poll. They carry the ' + $probeSubjectTag + ' tag, so the live tier''s ' +
                    'zero-artifact sweep will remove them - but remove them yourself rather than relying on that.')
            }
            [void](Invoke-Pop3Command -Socket $final.Socket -Command 'QUIT' -AllowError)
        }
        finally {
            Close-SinkSocket -Socket $final.Socket
        }
    }
    catch {
        Fail 'the sink is drained afterwards' $_.Exception.Message
    }
}

<#
    Deletes everything the mailbox currently holds, re-reading STAT after each delete rather than
    counting down from the first answer. That is deliberate: this runs against a server whose
    numbering may shift under it, so the only safe loop is "ask, delete number 1, ask again".
#>
function Clear-SinkMailbox {
    param($Socket)

    for ($i = 0; $i -lt 50; $i++) {
        $stat = Invoke-Pop3Command -Socket $Socket -Command 'STAT' -AllowError
        if (-not $stat.Ok) { return }
        if ($stat.Text -match '^\+OK\s+0\s') { return }
        $del = Invoke-Pop3Command -Socket $Socket -Command 'DELE 1' -AllowError
        if (-not $del.Ok) { return }
    }
}

function Test-ServiceSurvivesRestart {
    Say ''
    Say "Restarting $ServiceName to prove nothing is re-served afterwards."

    try {
        Restart-Service -Name $ServiceName -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
    }
    catch {
        Fail 'the service restarts cleanly' $_.Exception.Message
        return
    }

    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $status = 'absent'
    if ($null -ne $svc) { $status = [string] $svc.Status }
    if ($status -ne 'Running') {
        Fail 'the service restarts cleanly' "After a restart the service status is '$status'."
        return
    }
    Pass 'the service restarts cleanly' 'Running'

    try {
        $session = Connect-Pop3Authenticated
        try {
            $stat = Invoke-Pop3Command -Socket $session.Socket -Command 'STAT'
            if ($stat.Text -match '^\+OK\s+0\s') {
                Pass 'nothing is re-served after a restart' 'the mailbox is still empty'
            }
            else {
                Fail 'nothing is re-served after a restart' (
                    "STAT reports '$($stat.Text)' after a service restart. A sink that resurrects deleted mail on " +
                    'restart re-delivers every past seed into the hub Inbox, and every arrival assertion then matches ' +
                    'the wrong item.')
            }
            [void](Invoke-Pop3Command -Socket $session.Socket -Command 'QUIT' -AllowError)
        }
        finally {
            Close-SinkSocket -Socket $session.Socket
        }
    }
    catch {
        Fail 'nothing is re-served after a restart' $_.Exception.Message
    }
}

function Test-SettingsArePrecedent {
    # The silent one. See THE APPSETTINGS TRAP in the banner.
    $binPath = $null
    try {
        $binPath = (Get-CimInstance -ClassName Win32_Service -Filter "Name='$ServiceName'" -ErrorAction Stop).PathName
    }
    catch {
        Say '  NOTE the service binary path could not be read; whether the install-directory settings file wins is unknown.'
        return
    }

    if ($binPath -and $binPath.Contains($noUserSettingsFlag)) {
        Pass 'the install-directory settings file is authoritative' "the service runs with $noUserSettingsFlag"
    }
    else {
        Fail 'the install-directory settings file is authoritative' (
            "The service binary path is '$binPath' and does not carry $noUserSettingsFlag. smtp4dev layers a settings " +
            'file from the running account''s roaming profile ON TOP of the install-directory one, and the service ' +
            'account is not the logged-on user - so a configuration edited here can be silently overridden or ' +
            'silently ignored. Re-run with -Execute -Force, which sets it.')
    }

    if (Test-Path -LiteralPath $localSystemOverlay) {
        Say "  NOTE an overlay settings file exists at $localSystemOverlay."
        Say "       With $noUserSettingsFlag it is ignored; without it, it WINS over $InstallRoot\$settingsFileName."
    }
}

# ---------------------------------------------------------------------------------------------
# Modes.
# ---------------------------------------------------------------------------------------------

function Show-Plan {
    $packageText = '<not given - -Execute needs -PackagePath>'
    if ($PackagePath) { $packageText = $PackagePath }
    $pinText = '<unpinned>'
    if ($ExpectedSha256) { $pinText = $ExpectedSha256 }
    $imapText = 'disabled (written as a null port, NOT as 0 - 0 means auto-assign)'
    if ($ImapPort -gt 0) { $imapText = "$SinkHost`:$ImapPort" }

    Say 'PLAN (nothing below has been done - this is a dry run).'
    Say ''
    Say "  package             : $packageText"
    Say "  pinned to sha256    : $pinText"
    Say "  install root        : $InstallRoot"
    Say "  service             : $ServiceName  (registered with $noUserSettingsFlag)"
    Say "  submission          : $SinkHost`:$SmtpPort   (no auth, no TLS)"
    Say "  retrieval           : $SinkHost`:$Pop3Port   (POP3, no TLS)"
    Say "  IMAP                : $imapText"
    Say "  web UI              : $WebUiUrl"
    Say "  probe address       : $ProbeAddress  (POP3 user '$ProbeUser')"
    Say ''
    Say '  Three places must agree on the ports, or the tier fails in a way nothing diagnoses:'
    Say "    - the live-test settings mailSink block (submitPort $SmtpPort, retrievePort $Pop3Port)"
    Say '    - Testbed/guest/tier-profile-forcepst.prf SMTPPort / POP3Port'
    Say '    - this script'
}

function Invoke-Execute {
    Assert-TestbedGuestLocal
    Assert-Elevated
    Assert-OutlookNotRunning

    if (-not $PackagePath) {
        throw '-Execute needs -PackagePath. This script never downloads anything; see WHERE THE PACKAGE COMES FROM in the banner.'
    }
    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) {
        throw "The staged package '$PackagePath' does not exist. Stage it first and record it in Testbed/MEDIA.md."
    }

    # -- the pin ------------------------------------------------------------------------------

    $actualHash = (Get-FileHash -LiteralPath $PackagePath -Algorithm SHA256).Hash
    if ($ExpectedSha256) {
        if ($actualHash -ne $ExpectedSha256.Trim().ToUpperInvariant()) {
            throw @"
REFUSING: the staged package does not match the pin.

  expected  $($ExpectedSha256.Trim().ToUpperInvariant())
  actual    $actualHash

An unpinned sink that silently changes behaviour in an update is the same class of problem as an
Office auto-update invalidating a checkpoint. Re-stage the pinned version, or change the pin
deliberately and record it in Testbed/MEDIA.md.

NOTE: this script's DEFAULT pin is the hash the winget manifest publishes for the 3.15.0 win-x64
zip. It has never been checked against a downloaded file, so if you are staging that exact
release and see this message, verify the hash before assuming the package is wrong.
"@
        }
        Pass 'the package matches its pin' $actualHash
    }
    else {
        Say "  NOTE the package is UNPINNED. Its sha256 is $actualHash - record it and pass -ExpectedSha256 next time."
    }

    # -- ports, before anything is written ----------------------------------------------------

    $existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -ne $existing -and $existing.Status -eq 'Running') {
        # The idempotent path: this install's own listeners must not be reported as somebody
        # else's.
        $ep = Get-ListeningEndpoint -Port $SmtpPort
        if ($null -ne $ep) { $script:SinkPid = $ep.OwningProcess }
    }

    $ranges  = Get-ExcludedPortRange
    $portsOk = $true
    if (-not (Test-PortUsable -Port $SmtpPort -What 'submission' -Ranges $ranges)) { $portsOk = $false }
    if (-not (Test-PortUsable -Port $Pop3Port -What 'retrieval'  -Ranges $ranges)) { $portsOk = $false }
    if (-not $portsOk) {
        throw 'Refusing to install: the ports above cannot be used. Nothing has been written.'
    }
    Pass 'both ports are usable' "$SmtpPort and $Pop3Port, outside every reserved range and free"

    if ($null -ne $existing) {
        Say "  NOTE $ServiceName already exists (status $($existing.Status)). Reconfiguring it in place."
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    }

    # -- unpack -------------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $InstallRoot)) {
        New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
        Say "Created $InstallRoot"
    }

    $exePath = Join-Path $InstallRoot $serviceExeName
    if ($Force -or (-not (Test-Path -LiteralPath $exePath))) {
        # Expand-Archive ships with PowerShell 5.1 (Microsoft.PowerShell.Archive).
        Expand-Archive -LiteralPath $PackagePath -DestinationPath $InstallRoot -Force
        Say "Unpacked $PackagePath into $InstallRoot"
    }
    else {
        Say "$serviceExeName is already present; leaving the payload alone (pass -Force to replace it)."
    }

    if (-not (Test-Path -LiteralPath $exePath)) {
        throw @"
REFUSING: '$serviceExeName' is not in $InstallRoot after unpacking.

The win-x64 zip puts that executable and appsettings.json at its ROOT, so either this is a
different package - the Desktop zip, the noruntime zip, or another project entirely - or it
unpacks into a subdirectory. This script does not guess: point -InstallRoot at the directory that
actually holds the executable, or stage the right package.
"@
    }
    Pass 'the windowless executable is present' $exePath

    if (Test-Path -LiteralPath (Join-Path $InstallRoot $forbiddenExeName)) {
        Say "  NOTE $forbiddenExeName is also present. It creates a WINDOW, refuses --install-service, and always"
        Say '       chooses its own port. It must never be the executable registered as the service.'
    }

    # -- configure ----------------------------------------------------------------------------

    $settingsPath = Join-Path $InstallRoot $settingsFileName
    $json = New-SinkSetting | ConvertTo-Json -Depth 10

    # UTF-8 WITHOUT A BOM, and written through .NET rather than Set-Content. Windows PowerShell
    # 5.1's `-Encoding UTF8` emits a byte-order mark, and a BOM in front of a JSON document is a
    # coin-flip: some readers skip it and some report "'0xEF' is an invalid start of a value".
    # This file is read by another program, not by this script, so the coin is not ours to flip -
    # and the failure would arrive as a service that will not start, with nothing naming the
    # cause. The repository's other scripts use Set-Content for their LOGS, which is fine.
    [System.IO.File]::WriteAllText($settingsPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Say "Wrote $settingsPath"

    $readBack = Get-Content -LiteralPath $settingsPath -Raw
    if ($readBack.Trim() -ne $json.Trim()) {
        throw 'The configuration read back differently from what was written. Refusing to continue.'
    }
    Pass 'the configuration reads back byte-identical' $settingsPath

    # -- register, then make the install-directory settings authoritative ----------------------

    if ($null -eq (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue)) {
        & $exePath --install-service 2>&1 | ForEach-Object { Say "    $_" }
        if ($LASTEXITCODE -ne 0) {
            throw "Registering the service failed (exit $LASTEXITCODE). Nothing was started."
        }
        Say "Registered service $ServiceName"
    }
    else {
        Say "Service $ServiceName already registered"
    }

    # --install-service hardcodes the binary path as `"<exe>" --service` and passes nothing else
    # through, so the flag that closes the settings-overlay trap has to be added afterwards.
    #
    # WHY THIS WRITES ImagePath RATHER THAN CALLING sc.exe. The value contains embedded double
    # quotes around the executable path, and Windows PowerShell 5.1 mangles quotes when it hands
    # an argument to a native executable - so `sc.exe config <name> binPath= "<value>"` can
    # silently register a DIFFERENT path from the one intended, which produces a service that
    # fails to start for a reason nothing in the output explains. ImagePath under the service key
    # IS the storage sc.exe writes, the write is exact, and it is read back below through
    # Win32_Service.PathName - the same property -Verify asserts on - so the loop is closed by
    # measurement rather than by trusting either tool.
    $serviceKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
    $binPath = (Get-CimInstance -ClassName Win32_Service -Filter "Name='$ServiceName'" -ErrorAction Stop).PathName
    if (-not $binPath.Contains($noUserSettingsFlag)) {
        $newBinPath = "$binPath $noUserSettingsFlag"
        Set-ItemProperty -LiteralPath $serviceKey -Name 'ImagePath' -Value $newBinPath -ErrorAction Stop
        Say "Set the service binary path to: $newBinPath"

        $confirmed = (Get-CimInstance -ClassName Win32_Service -Filter "Name='$ServiceName'" -ErrorAction Stop).PathName
        if ($confirmed -ne $newBinPath) {
            throw @"
REFUSING: the service binary path did not read back as written.

  wrote  $newBinPath
  read   $confirmed

Without $noUserSettingsFlag the service layers a settings file from its own account's roaming
profile over the one this script just wrote, so the configuration would be silently ignored.
"@
        }
        Pass 'the service binary path reads back as written' $confirmed
    }
    else {
        Say "The service binary path already carries $noUserSettingsFlag"
    }

    Start-Service -Name $ServiceName -ErrorAction Stop
    # A service can report Running before its listeners are bound. A short settle is cheaper than
    # a spurious failure in the verification that follows.
    Start-Sleep -Seconds 3

    $svc = Get-Service -Name $ServiceName -ErrorAction Stop
    if ($svc.Status -ne 'Running') {
        throw "Service $ServiceName is '$($svc.Status)' after Start-Service."
    }
    Pass 'the service is running' $ServiceName

    Invoke-Verify -SkipGuards
}

function Invoke-Verify {
    param([switch] $SkipGuards)

    if (-not $SkipGuards) {
        Assert-TestbedGuestLocal
        Assert-OutlookNotRunning
    }

    Say ''
    Say 'VERIFY.'

    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -eq $svc) {
        Fail 'the service exists' "No service named '$ServiceName'. Run with -Execute first."
        return
    }
    if ($svc.Status -ne 'Running') {
        Fail 'the service is running' "Status is '$($svc.Status)'."
        return
    }
    Pass 'the service is running' $ServiceName

    Test-SettingsArePrecedent
    Test-ListenerBinding -Port $SmtpPort -What 'submission'
    Test-ListenerBinding -Port $Pop3Port -What 'retrieval'

    if ($ImapPort -le 0) {
        $imap = Get-ListeningEndpoint -Port 143
        if ($null -eq $imap) {
            Pass 'IMAP is off' 'nothing is listening on 143'
        }
        else {
            Fail 'IMAP is off' (
                "Something is listening on 143 ($($imap.LocalAddress), pid $($imap.OwningProcess)). smtp4dev enables " +
                'IMAP by DEFAULT and only a null port disables it - a 0 means auto-assign. Nothing here retrieves over ' +
                'IMAP, and IMAP would not work for this suite anyway: an IMAP account gets its own store and cannot ' +
                'deliver into the hub PST.')
        }
    }

    Test-NoInboundFirewallRule
    Invoke-RoundTripProof
    Invoke-OrdinalStabilityProof
    Test-ServiceSurvivesRestart
}

function Invoke-Uninstall {
    Assert-TestbedGuestLocal
    Assert-Elevated

    if ($null -ne (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue)) {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        $exePath = Join-Path $InstallRoot $serviceExeName
        if (Test-Path -LiteralPath $exePath) {
            & $exePath --uninstall-service 2>&1 | ForEach-Object { Say "    $_" }
        }
        else {
            & sc.exe delete $ServiceName 2>&1 | ForEach-Object { Say "    $_" }
        }
        Say "Removed service $ServiceName"
    }
    else {
        Say "Service $ServiceName already absent"
    }

    if (Test-Path -LiteralPath $InstallRoot) {
        Remove-Item -LiteralPath $InstallRoot -Recurse -Force
        Say "Removed $InstallRoot"
    }
    else {
        Say "$InstallRoot already absent"
    }

    if ($null -ne (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue)) {
        Fail 'the service is gone' "'$ServiceName' still exists after the uninstall."
    }
    else {
        Pass 'the service is gone' $ServiceName
    }
}

# ---------------------------------------------------------------------------------------------
# Entry point.
# ---------------------------------------------------------------------------------------------

Say "Install-MailSink.ps1 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Say ''

if ($Verify -and $Execute) {
    throw 'Pass -Execute or -Verify, not both: -Execute runs the full verification when it finishes.'
}
if ($Uninstall -and $Verify) {
    throw 'Pass -Uninstall or -Verify, not both.'
}

if ($Uninstall) {
    if (-not $Execute) {
        Say 'PLAN: stop and remove the service, then delete the install root.'
        Say ''
        Say "  service      : $ServiceName"
        Say "  install root : $InstallRoot"
        Say ''
        Say 'Dry run. Nothing removed. Re-run with -Uninstall -Execute.'
    }
    else {
        Invoke-Uninstall
    }
}
elseif ($Verify) {
    Invoke-Verify
}
elseif ($Execute) {
    Invoke-Execute
}
else {
    Show-Plan
    Say ''
    Say 'Dry run. Nothing written. Re-run with -PackagePath <zip> -Execute.'
}

Say ''
if ($script:Failures.Count -gt 0) {
    Say ("{0} check(s) FAILED:" -f $script:Failures.Count)
    foreach ($f in $script:Failures) { Say "  - $f" }
    Save-Log
    exit 1
}

if ($Verify -or $Execute) {
    Say 'All checks passed.'
}
Save-Log
exit 0
