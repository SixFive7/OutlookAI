#Requires -Version 5.1
<#
    ============================================================================================
    THIS SCRIPT HAS NEVER RUN ON A GUEST. ITS -SelfTest HAS, ON THE HOST.
    ============================================================================================

    Written 2026-09-24 by an agent with no guest available and forbidden to run a mail sink on
    the machine it worked on - the maintainer's workstation, with real Outlook and real mail on
    it. So what is established, and how, is exactly this:

      * it PARSES under Windows PowerShell 5.1 and PowerShell 7, and -SelfTest passes under both
        (the count is in the -SelfTest output). -SelfTest drives every pure decision in this file
        - the launcher it writes, the mailbox naming rule, dot-stuffing both ways, the probe
        message and the checks on what comes back, the scheduled-task audit, the install-root
        refusals - against synthetic inputs, with no socket, no process and no file;
      * the PACKAGE is real: inbucket_3.1.1_windows_amd64.zip was downloaded on the host and its
        SHA-256 matched three ways - the maintainers' own checksums file, GitHub's asset digest
        and the file itself (Testbed/MEDIA.md, "The mail sink");
      * every claim below about how Inbucket BEHAVES is marked [SOURCE]: read in its source at
        tag v3.1.1, never observed. -Verify exists to turn each one into [MEASURED] or into a
        named FAIL on the first guest run.

    Replace this banner with what actually happened the first time -Execute runs on a guest,
    and say which checks passed.

.SYNOPSIS
    Installs Inbucket as the testbed's loopback mail sink, starts it with the guest, and then
    PROVES an SMTP-to-POP3 round trip over raw sockets instead of reporting open ports.

.DESCRIPTION
    RUN ON THE GUEST, ELEVATED. Windows PowerShell 5.1 - no ternary, no `??`.
    NEVER run this on the maintainer's workstation. It refuses there; see THE GUARD below.

    WHY A SINK. Docs/live-tier-on-the-vm.md section 1.4. Thirteen live methods put mail on the
    wire and need it to come back into the hub store's Inbox, and one of them - the product's
    own two-step `send` - cannot be replaced by seeding at all. The maintainer decided on
    2026-09-24 that every test moves to the guests without compromise, so the guests need
    transport, and a guest has no network: the transport has to be a server on 127.0.0.1 that
    accepts what Outlook submits and hands it straight back over POP3.

    WHY INBUCKET, AND NOT THE TWO THAT WERE ON THE TABLE. All three are free, open source and
    permissively licensed. What separates them is what this testbed actually needs:

                          Inbucket 3.1.1     Mailpit 1.31.2        smtp4dev 3.15.0
      POP3 empty PASS     ACCEPTED           refused, and the      refused
                                             connection closed
      POP3 without a      yes                no - POP3 is off      yes
      configured login                       until one is set
      whose mail a POP3   the mailbox named  EVERY message, to     every message, to
      login sees          by the USER        every login           every login
      DELE                at QUIT (RFC 1939) at QUIT (RFC 1939)    immediately, and the
                                                                   mailbox renumbers
      maintainer-         yes, goreleaser    no - only GitHub's    no - only GitHub's
      published hash      checksums.txt      computed digest       computed digest
      runtime             none (Go)          none (Go)             none (self-contained
                                                                   .NET, 77 MB zip)

    Two of those rows decide it. (1) THE PASSWORD. Outlook must never prompt on an unattended
    guest - a prompt is a hang - and the tier profile's account stores no password. A sink that
    accepts any PASS, including none, removes the sink half of that problem entirely; the other
    two require a password that Outlook would then have to hold, and that no free route writes.
    (2) TWO ACCOUNTS. Docs/live-tier-on-the-vm.md section 2.8b adds an identity account beside
    the dummy one, both POP3 against this sink. With a catch-all sink whichever account polls
    first downloads the other's mail - an intermittent misdelivery that reads as "the mail never
    arrived". Inbucket files each message under the recipient's local part and a POP3 login
    reads only the mailbox its USER names, so each account sees its own mail and nothing else.

    WHAT ONLY A GUEST CAN SETTLE - THE OUTLOOK HALF OF THE PASSWORD. The sink accepts whatever
    Outlook sends [SOURCE]. Whether Outlook, holding NO stored password, sends an empty PASS or
    raises its "Internet E-mail" logon prompt instead is not documented anywhere this agent
    could find, and it cannot be measured on the host. Docs/live-tier-on-the-vm.md section 2.7
    gives the measurement - it needs no mail, only an Outlook start with this sink at
    -LogLevel debug - and the fallback if Outlook does prompt.

    WHAT -Verify PROVES, AND WHAT IT DELIBERATELY DOES NOT TOUCH. It speaks SMTP and POP3 itself,
    and it only ever writes to and deletes from two mailboxes of its OWN - the probe mailbox and
    its isolation twin - which no Outlook account logs in to. So it can prove the round trip
    while Outlook is running, and a probe it leaves behind can never be downloaded into a store.
    The accounts' mailboxes are opened read-only, to REPORT mail waiting for Outlook, and are
    never emptied: there is exactly one code path here that issues DELE, and it refuses any
    mailbox that is not a probe mailbox. Checks, each named in the output:

      * the scheduled task that starts the sink with the guest is registered as it must be -
        SYSTEM, at startup, no time limit (the default kills a task after three days);
      * the launcher on disk is byte-for-byte what this script would write, so a hand edit is
        drift rather than configuration;
      * the sink process is running from the install root, and all three listeners are bound to
        127.0.0.1 and owned by it;
      * POP3 accepts an empty PASS, a PASS with nothing after it, and any non-empty PASS;
      * a message submitted over SMTP comes back over POP3 with its subject, a bare "." line,
        a leading "." and a leading ".." intact, and a base64 attachment byte-identical;
      * TOP, which the capability list advertises, actually works;
      * a message for one mailbox is invisible to another;
      * message numbers stay fixed after a DELE, and deletes happen at QUIT - so a client that
        drops the connection mid-session loses nothing;
      * after a restart through the scheduled task nothing deleted comes back, and message ids
        are never reused - the file store's ids are timestamps. Skipped while Outlook runs,
        because a restart under a polling Outlook is exactly the nondeterminism this avoids.

    THE PROBE MESSAGES CARRY THE LIVE-TIER TAG ANYWAY. They go to a mailbox no account reads, so
    they cannot reach a store. If that ever stopped being true, the subject tag
    [OutlookAI-McpTest] puts them inside the live tier's mandatory zero-artifact sweep, which
    matches it ordinally. The corpus tag [OutlookAI-Corpus] is deliberately NOT used and must
    never be.

    WHY A SCHEDULED TASK AND A LAUNCHER, NOT A SERVICE. inbucket.exe is a plain Go console
    program that does not speak to the Service Control Manager, so registering it as a service
    fails with a start timeout. Task Scheduler ships with Windows: the task runs as SYSTEM in
    session 0 at every boot - before autologon and before Outlook - so nothing draws a window.
    Inbucket reads its configuration from the ENVIRONMENT and nowhere else [SOURCE], and a
    scheduled task cannot set environment variables for what it starts, so the task runs a
    generated run-sink.cmd that sets them and then runs the exe.

    WHAT IT NEVER DOES. It never starts Outlook, never creates a COM object, never touches MAPI,
    never reads or writes an Outlook profile, and never touches a mail item in any store.
    Everything it does is a directory, a launcher, a scheduled task, a process, and TCP
    conversations with 127.0.0.1.

    THE GUARD. Two axes, both required, same shape as Testbed/guest/Install-DotnetSdk.ps1:
    logged on as the guests' autologon account, AND a computer name carrying the prefix
    Testbed/host/New-AnswerFile.ps1 gives every guest. Restated here rather than dot-sourced
    from the shared guest layer, which carries an Outlook COM helper a mail sink installer has
    no business loading. The only way past either axis is to name the value you mean.

    STARTING STATE IT EXPECTS.
      * An ELEVATED session on the guest. Registering a SYSTEM task, and stopping and starting
        it in -Verify, needs it; this asserts it rather than failing halfway.
      * For -Execute and -Uninstall: OUTLOOK.EXE not running. Swapping the sink out from under
        a polling Outlook turns a quiet send/receive into an error nobody asked for. It is not
        killed: mailbox-safety rule 7 forbids that outright.
      * The package STAGED at -PackagePath - by Testbed/host/Get-MailSinkMedia.ps1 on the host,
        then Testbed/host/Copy-ToGuest.ps1. It is never downloaded here; the guest has no
        network.
      * Its SHA-256 in -ExpectedSha256. There is no default, for the reason Testbed/MEDIA.md
        gives for the .NET SDK: a number that travels inside the script it checks is a number
        nobody looks up. The recorded value is in Testbed/MEDIA.md.
      * Ports -SmtpPort, -Pop3Port and -WebPort free on 127.0.0.1 and outside every Windows
        reserved TCP range. A reservation is not a listener: netstat shows nothing and the bind
        fails anyway.

    IDEMPOTENT. Re-running -Execute stops the sink, keeps the unpacked payload unless -Force,
    rewrites the launcher, re-registers the task, starts it and runs the whole verification.
    -Uninstall -Execute takes it all back out, so a checkpoint can be re-taken.

.PARAMETER PackagePath
    The staged release zip on the guest. Never downloaded here.

.PARAMETER ExpectedSha256
    The package's SHA-256, from Testbed/MEDIA.md. REQUIRED for -Execute; no default.

.PARAMETER InstallRoot
    Where the payload, the launcher, the mail store and the sink's own log live. Space-free on
    purpose: it goes into a command line and into a setting whose syntax reserves ':' and ','.

.PARAMETER TaskName
    The scheduled task that starts the sink at boot.

.PARAMETER SmtpPort
    Submission port. Must equal the tier profile's SMTPPort and the live-test settings'
    mailSink.submitPort.

.PARAMETER Pop3Port
    Retrieval port. Must equal the tier profile's POP3Port and mailSink.retrievePort.

.PARAMETER WebPort
    Inbucket's web UI and REST API. It cannot be switched off, so it is bound to loopback like
    the rest. No test uses it; it is what a human opens when a round trip fails.

.PARAMETER LogLevel
    Inbucket's own log level. `debug` logs every POP3 command line an account sends - including
    its PASS, which on this sink is not a secret - and is what the first Outlook start should
    run with, because that line is the evidence for the Outlook half of the password question.

.PARAMETER ProbeAddress
    The address -Verify submits to and retrieves as. Its mailbox, and an isolation twin derived
    from it, are the ONLY mailboxes this script ever deletes from. Refused if it maps to the
    same mailbox as any -AccountAddress.

.PARAMETER AccountAddress
    The addresses of the Outlook accounts that poll this sink. Their mailboxes are reported,
    read-only, and never emptied - and a probe address that would land in one is refused.

.PARAMETER ExpectedUser
    Accounts this may run as. The default IS the guard; do not widen it.

.PARAMETER ExpectedComputerNamePrefix
    Computer-name prefix this may run on. New-AnswerFile.ps1 names guests OAI-*.

.PARAMETER LogPath
    This script's transcript. Outside -InstallRoot, so -Uninstall does not delete its own record
    and Testbed/host/Copy-FromGuest.ps1 collects it with the other logs.

.PARAMETER Execute
    Install, or with -Uninstall, remove. Without it nothing is read, written or started.

.PARAMETER Verify
    Run the proof only. Writes nothing but its own log and probe messages, which it removes.

.PARAMETER Uninstall
    Stop the sink, unregister the task, delete the install root. Needs -Execute.

.PARAMETER Force
    Let -Execute replace an already unpacked payload.

.PARAMETER SelfTest
    Exercise every pure decision in this file against synthetic inputs. Touches nothing - no
    socket, no process, no file, no registry - so it runs anywhere, the host included.

.EXAMPLE
    .\Install-MailSink.ps1
    .\Install-MailSink.ps1 -ExpectedSha256 <hash from Testbed/MEDIA.md> -Execute
    .\Install-MailSink.ps1 -ExpectedSha256 <hash> -LogLevel debug -Execute
    .\Install-MailSink.ps1 -Verify
    .\Install-MailSink.ps1 -Uninstall -Execute
    powershell.exe -NoProfile -File Testbed\guest\Install-MailSink.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string]   $PackagePath                = 'C:\OutlookAI-Q5\media\inbucket_3.1.1_windows_amd64.zip',
    [string]   $ExpectedSha256,
    [string]   $InstallRoot                = 'C:\OutlookAI-Sink',
    [string]   $TaskName                   = 'OutlookAI-MailSink',
    [int]      $SmtpPort                   = 25,
    [int]      $Pop3Port                   = 110,
    [int]      $WebPort                    = 9000,
    [ValidateSet('debug', 'info', 'warn', 'error')] [string] $LogLevel = 'info',
    [string]   $ProbeAddress               = 'outlookai-sink-probe@vm.invalid',
    [string[]] $AccountAddress             = @('tier@vm.invalid', 'identity@vm.invalid'),
    [string[]] $ExpectedUser               = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-',
    [string]   $LogPath                    = 'C:\OutlookAI-Q5\install-mail-sink.log',
    [switch]   $Execute,
    [switch]   $Verify,
    [switch]   $Uninstall,
    [switch]   $Force,
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'
if (Test-Path Variable:\PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

# Captured here: inside a function $PSBoundParameters is that FUNCTION's, not the script's.
$ScriptBoundParameters = $PSBoundParameters

# ---------------------------------------------------------------------------------------------
# Constants. Everything this script writes, runs or asserts is named here, so the blast radius
# is readable without reading the code.
# ---------------------------------------------------------------------------------------------

# Loopback, and only loopback. Not a parameter: a sink reachable any other way is an open relay
# on a test VM, and there is no configuration of this testbed that wants one.
$SinkHost = '127.0.0.1'

# The live tier's artifact tag, matched ORDINALLY by the zero-artifact sweep. NOT the corpus tag.
$ProbeSubjectTag = '[OutlookAI-McpTest]'

# What a POP3 login sends when -Verify wants a non-empty password. Any value works on this sink
# [SOURCE]; it is a probe of "does the field need to be populated", not a credential.
$NonEmptyPassValue = 'any-value'

$ExeName          = 'inbucket.exe'
$LauncherName     = 'run-sink.cmd'
$StoreDirName     = 'store'
$SinkLogName      = 'inbucket.log'
$LuaScriptName    = 'inbucket.lua'
$CmdExe           = Join-Path $env:SystemRoot 'System32\cmd.exe'

# Loopback is instant. A connect or a read that is not is one that is not going to finish.
$SocketTimeoutMs = 10000

# The suite's tightest arrival deadline is 120 s and it also has to cover Outlook. A sink that
# needs more than this on loopback has already failed.
$RoundTripBudgetSeconds = 30

# How long a start or a stop may take before it is reported as not having happened.
$ListenerWaitSeconds = 30

# The verdicts. Three, because "not installed" and "installed and wrong" want different fixes.
$VerdictReady  = 'SINK-READY'
$VerdictAbsent = 'SINK-ABSENT'
$VerdictBroken = 'BROKEN'

$script:Failures = @()
$script:SeenIds  = @()
$script:LogReady = $false

# =============================================================================================
# PURE DECISIONS. Everything below until the socket layer is decided from arguments alone, so
# -SelfTest can drive it on any machine. No function in this block reads the machine.
# =============================================================================================

<#
    Inbucket's mailbox for an address, with MailboxNaming=local [SOURCE: pkg/policy/address.go,
    parseMailboxName]: the local part, lowercased, cut at the first '+'. A POP3 USER is used
    VERBATIM as a mailbox name [SOURCE: pkg/server/pop3/handler.go, loadMailbox] - so an
    account's POP3 user name must be exactly what this returns for its address.
#>
function Get-InbucketMailboxName {
    param([Parameter(Mandatory = $true)] [string] $Address)

    $at = $Address.LastIndexOf([char] '@')
    if ($at -lt 0) { throw "'$Address' is not an address: it has no '@'." }
    $local = $Address.Substring(0, $at)
    if ($local.Length -eq 0) { throw "'$Address' has an empty local part." }

    $name = $local.ToLowerInvariant()
    foreach ($c in $name.ToCharArray()) {
        $ok = ($c -ge [char]'a' -and $c -le [char]'z') -or ($c -ge [char]'0' -and $c -le [char]'9') -or
              ("!#$%&'*+-=/?^_``.{|}~".IndexOf($c) -ge 0)
        if (-not $ok) { throw "'$Address' carries a character Inbucket refuses in a mailbox name: '$c'." }
    }
    $plus = $name.IndexOf('+')
    if ($plus -ge 0) { $name = $name.Substring(0, $plus) }
    if ($name.Length -eq 0) { throw "'$Address' names an empty mailbox once its +tag is removed." }
    return $name
}

# The second probe-owned mailbox, used to prove one mailbox cannot see another's mail. Never a
# '+' suffix: the naming rule above would fold that back into the probe mailbox itself.
function Get-IsolationAddress {
    param([Parameter(Mandatory = $true)] [string] $Address)
    $at = $Address.LastIndexOf([char] '@')
    if ($at -lt 1) { throw "'$Address' is not an address." }
    return $Address.Substring(0, $at) + '-isolation' + $Address.Substring($at)
}

# The probe must never share a mailbox with an Outlook account: its messages would then be
# downloaded into a real store, and its deletes would take an account's mail.
function Get-ProbeMailboxProblem {
    param([string] $Probe, [string[]] $Accounts)

    try {
        $probeBox = Get-InbucketMailboxName -Address $Probe
        $isolationBox = Get-InbucketMailboxName -Address (Get-IsolationAddress -Address $Probe)
    }
    catch { return "The probe address is unusable: $($_.Exception.Message)" }

    foreach ($account in @($Accounts)) {
        if (-not $account) { continue }
        $accountBox = $null
        try { $accountBox = Get-InbucketMailboxName -Address $account }
        catch { return "The account address is unusable: $($_.Exception.Message)" }
        if ($accountBox -ceq $probeBox) {
            return "The probe address '$Probe' lands in mailbox '$probeBox', which is the mailbox of account '$account'. A probe there would be downloaded into that account's store, and its cleanup would delete that account's mail."
        }
        if ($accountBox -ceq $isolationBox) {
            return "The probe's isolation mailbox '$isolationBox' is the mailbox of account '$account'. Choose another -ProbeAddress."
        }
    }
    return $null
}

# The single gate in front of DELE. Any mailbox that is not probe-owned is refused, whatever
# called it - an account's mail is never this script's to delete.
function Assert-ProbeMailboxForDelete {
    param([string] $Mailbox, [string[]] $ProbeMailboxes)
    foreach ($p in @($ProbeMailboxes)) {
        if ($p -and ($Mailbox -ceq $p)) { return }
    }
    throw "REFUSING to DELE in mailbox '$Mailbox': only the probe mailboxes ($(@($ProbeMailboxes) -join ', ')) may ever be emptied by this script."
}

# Inbucket's storage setting is key:value pairs separated by ',' - so a Windows path's ':' is
# written as '$' and replaced back [SOURCE: pkg/storage/file/fstore.go, getMailPath].
function ConvertTo-InbucketStoragePath {
    param([Parameter(Mandatory = $true)] [string] $Path)
    if ($Path -notmatch '^[A-Za-z]:\\') { throw "'$Path' is not an absolute drive path." }
    if ($Path.Contains(',') -or $Path.Contains('$')) { throw "'$Path' contains ',' or '$', which Inbucket's storage setting cannot carry." }
    return $Path.Substring(0, 1) + '$' + $Path.Substring(2)
}

# What cmd.exe would interpret rather than pass through. A value carrying one of these would be
# written into the launcher as something other than what it says.
function Get-CmdUnsafeReason {
    param([string] $Value)
    if ($null -eq $Value) { return 'it is null' }
    foreach ($c in @('%', '^', '&', '|', '<', '>', '"', '!', "`r", "`n")) {
        if ($Value.Contains($c)) {
            $shown = $c
            if ($c -eq "`r") { $shown = 'CR' }
            if ($c -eq "`n") { $shown = 'LF' }
            return "it contains '$shown', which cmd.exe interprets"
        }
    }
    return $null
}

# Refuses install roots whose recursive deletion (by -Uninstall) or whose use in a command line
# would be dangerous: drive roots, system and profile directories, UNC paths, relative paths.
function Get-InstallRootProblem {
    param([string] $Path)
    if (-not $Path) { return 'it is empty' }
    if ($Path -notmatch '^[A-Za-z]:\\[^\\]+') { return 'it must be an absolute path at least one directory below a drive root (no UNC, no relative path)' }
    if ($Path -match '\s') { return 'it contains whitespace, and it goes into a command line' }
    $trimmed = $Path.TrimEnd('\')
    $forbidden = @('Windows', 'Program Files', 'Program Files (x86)', 'ProgramData', 'Users', 'System Volume Information')
    foreach ($f in $forbidden) {
        $root = $trimmed.Substring(0, 3) + $f
        if ([string]::Equals($trimmed, $root, [System.StringComparison]::OrdinalIgnoreCase)) { return "it is the system directory '$root'" }
        if ($trimmed.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) { return "it is inside the system directory '$root'" }
    }
    $unsafe = Get-CmdUnsafeReason -Value $Path
    if ($unsafe) { return $unsafe }
    return $null
}

<#
    The environment the launcher sets, in order, each with the reason it is set. Inbucket is
    configured through INBUCKET_* variables and nothing else [SOURCE: pkg/config/config.go].
#>
function Get-SinkEnvironment {
    param([string] $Root, [int] $Smtp, [int] $Pop3, [int] $Web, [string] $Level)

    $store = ConvertTo-InbucketStoragePath -Path (Join-Path $Root $StoreDirName)
    return @(
        [pscustomobject]@{ Name = 'INBUCKET_LOGLEVEL';                Value = $Level;                 Why = 'debug logs every POP3 command line an account sends, PASS included - the evidence for the Outlook half of the password question.' }
        [pscustomobject]@{ Name = 'INBUCKET_MAILBOXNAMING';           Value = 'local';                Why = 'A message for NAME@anything is filed under mailbox "name", and POP3 USER name reads exactly that. This is what stops two accounts taking each other''s mail.' }
        [pscustomobject]@{ Name = 'INBUCKET_SMTP_ADDR';               Value = "${SinkHost}:$Smtp";    Why = 'Submission. Loopback only: any other address is an open relay on a test VM.' }
        [pscustomobject]@{ Name = 'INBUCKET_POP3_ADDR';               Value = "${SinkHost}:$Pop3";    Why = 'Retrieval. Loopback only.' }
        [pscustomobject]@{ Name = 'INBUCKET_WEB_ADDR';                Value = "${SinkHost}:$Web";     Why = 'Web UI and REST API. It cannot be switched off, so it is confined to loopback. Nothing under test uses it.' }
        [pscustomobject]@{ Name = 'INBUCKET_SMTP_TLSENABLED';         Value = 'false';                Why = 'The tier profile sets SMTPUseSSL=0. The default already, written down so it cannot drift.' }
        [pscustomobject]@{ Name = 'INBUCKET_POP3_TLSENABLED';         Value = 'false';                Why = 'The tier profile sets POP3UseSSL=0; with this false the capability list offers no STLS.' }
        [pscustomobject]@{ Name = 'INBUCKET_STORAGE_TYPE';            Value = 'file';                 Why = 'Survives a restart, and its message ids are timestamps - the memory store numbers from 1 again after every restart, which would reuse UIDLs.' }
        [pscustomobject]@{ Name = 'INBUCKET_STORAGE_PARAMS';          Value = "path:$store";          Why = 'Where the file store lives. The $ stands for the drive colon; this setting reserves ":" to separate key from value.' }
        [pscustomobject]@{ Name = 'INBUCKET_STORAGE_RETENTIONPERIOD'; Value = '24h';                  Why = 'Mail nobody collected is purged after a day, so a failed run cannot leave residue in the sink for good. The default, written down.' }
        [pscustomobject]@{ Name = 'INBUCKET_STORAGE_MAILBOXMSGCAP';   Value = '500';                  Why = 'Past this the oldest message in a mailbox is dropped silently. The default, written down; a live run sends thirteen.' }
    )
}

<#
    The launcher the scheduled task runs. ASCII, CRLF, no BOM - it is read by cmd.exe. -Verify
    re-renders it and fails on any difference, so it is a record rather than a starting point.
#>
function ConvertTo-LauncherText {
    param(
        [Parameter(Mandatory = $true)] [string] $ExePath,
        [Parameter(Mandatory = $true)] [string] $AppDir,
        [Parameter(Mandatory = $true)] [string] $SinkLogPath,
        [Parameter(Mandatory = $true)] $Environment
    )

    foreach ($pair in @(@('the exe path', $ExePath), @('the app directory', $AppDir), @('the sink log path', $SinkLogPath))) {
        $why = Get-CmdUnsafeReason -Value $pair[1]
        if ($why) { throw "Refusing to write a launcher: $($pair[0]) '$($pair[1])' is unusable - $why." }
    }

    $lines = @(
        '@echo off'
        'rem WRITTEN BY Testbed/guest/Install-MailSink.ps1 -Execute. DO NOT EDIT BY HAND.'
        'rem -Verify re-renders this file and FAILS on any difference, so an edit here is reported'
        'rem as drift instead of quietly becoming the configuration. Change the script''s'
        'rem parameters and re-run -Execute.'
        'rem'
        'rem Inbucket takes its configuration from the environment and from nowhere else, and a'
        'rem scheduled task cannot set environment variables for what it starts: hence this file.'
        'setlocal'
    )
    foreach ($e in @($Environment)) {
        $why = Get-CmdUnsafeReason -Value ([string] $e.Value)
        if ($why) { throw "Refusing to write a launcher: $($e.Name) is unusable - $why." }
        $lines += ('set "{0}={1}"' -f $e.Name, $e.Value)
    }
    $lines += ('cd /d "{0}"' -f $AppDir)
    $lines += ('"{0}" -logfile "{1}"' -f $ExePath, $SinkLogPath)

    return (($lines -join "`r`n") + "`r`n")
}

<#
    The folder inside the release zip that holds inbucket.exe. The goreleaser zip puts everything
    under one versioned folder; this reads it from the entries rather than trusting a file name,
    and refuses anything it cannot place exactly.
#>
function Get-ZipAppFolder {
    param([string[]] $EntryNames)

    $hits = @()
    foreach ($raw in @($EntryNames)) {
        if (-not $raw) { continue }
        $name = $raw.Replace('\', '/')
        $leaf = ($name -split '/')[-1]
        if ([string]::Equals($leaf, $ExeName, [System.StringComparison]::OrdinalIgnoreCase)) { $hits += $name }
    }
    if ($hits.Count -eq 0) { throw "The package holds no $ExeName. It is not the Inbucket Windows release." }
    if ($hits.Count -gt 1) { throw "The package holds $($hits.Count) copies of ${ExeName}: $($hits -join ', '). Refusing to guess which one runs." }

    $parts = $hits[0] -split '/'
    if ($parts.Count -eq 1) { return '' }
    if ($parts.Count -eq 2) { return $parts[0] }
    throw "$ExeName sits $($parts.Count - 1) folders deep in the package ($($hits[0])). The release puts it one folder deep; this is not that package."
}

function Get-VersionFromFolderName {
    param([string] $FolderName)
    $m = [regex]::Match([string] $FolderName, '^inbucket_(\d+\.\d+\.\d+)_windows_amd64$')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Get-Sha256Problem {
    param([string] $Expected, [string] $Actual)
    if ([string]::IsNullOrWhiteSpace($Expected)) {
        return '-ExpectedSha256 was not supplied. It is mandatory and has no default: take it from Testbed/MEDIA.md.'
    }
    $want = $Expected.Trim().Replace('-', '').ToUpperInvariant()
    if ($want -notmatch '^[0-9A-F]{64}$') { return "-ExpectedSha256 '$Expected' is not a SHA-256 (64 hex digits)." }
    $have = ([string] $Actual).Trim().ToUpperInvariant()
    if ($want -cne $have) { return "the package does not match its pin.`n  expected  $want`n  actual    $have" }
    return $null
}

# netsh's reserved-range table. A line is "start end", optionally followed by '*' for a range an
# administrator reserved.
#
# A NOTE ON EVERY FUNCTION HERE THAT RETURNS A LIST: it returns the items, never `,$list`, and
# every caller wraps the call in @(). Returning `,$list` hands the caller ONE object that is the
# list, and a caller that then writes @(...) around it gets a one-element list whose element is
# the list - so an empty result counts as 1. This file does it one way, everywhere.
function Get-ExcludedPortRangeFromText {
    param([string[]] $Lines)
    foreach ($line in @($Lines)) {
        $m = [regex]::Match([string] $line, '^\s*(\d+)\s+(\d+)(\s+\*)?\s*$')
        if ($m.Success) {
            [pscustomobject]@{ Start = [int] $m.Groups[1].Value; End = [int] $m.Groups[2].Value }
        }
    }
}

function Test-PortInRange {
    param([int] $Port, $Ranges)
    foreach ($r in @($Ranges)) {
        if ($null -eq $r) { continue }
        if ($Port -ge $r.Start -and $Port -le $r.End) { return $true }
    }
    return $false
}

# One SMTP reply from its lines. A continuation line has '-' in column 4; the last has a space.
function ConvertFrom-SmtpReplyLines {
    param([string[]] $Lines)
    $all = @($Lines)
    $code = 0
    if ($all.Count -gt 0) {
        $last = [string] $all[$all.Count - 1]
        if ($last.Length -ge 3) { [void][int]::TryParse($last.Substring(0, 3), [ref] $code) }
    }
    return [pscustomobject]@{ Code = $code; Text = ($all -join ' | '); LineCount = $all.Count }
}

# RFC 5321 section 4.5.2 on the way out: a line beginning with '.' gets one more.
function ConvertTo-DotStuffedLines {
    param([string[]] $Lines)
    foreach ($line in @($Lines)) {
        if ($line.StartsWith('.')) { '.' + $line } else { $line }
    }
}

# RFC 1939 section 3 on the way back: stop at a line that is exactly '.', drop one leading '.'
# from every other line that has one. Reports whether the terminator was actually seen - a
# response that simply stops is a truncation, not an empty body.
function ConvertFrom-DotStuffedLines {
    param([string[]] $Lines)
    $out = @()
    $terminated = $false
    foreach ($line in @($Lines)) {
        if ($line -ceq '.') { $terminated = $true; break }
        if ($line.StartsWith('.')) { $out += $line.Substring(1) } else { $out += $line }
    }
    return [pscustomobject]@{ Lines = $out; Terminated = $terminated }
}

function ConvertFrom-Pop3Stat {
    param([string] $Line)
    $m = [regex]::Match([string] $Line, '^\+OK\s+(\d+)\s+(\d+)')
    if (-not $m.Success) { return $null }
    return [pscustomobject]@{ Count = [int] $m.Groups[1].Value; Octets = [long] $m.Groups[2].Value }
}

function ConvertFrom-Pop3Uidl {
    param([string] $Line)
    $m = [regex]::Match([string] $Line, '^\+OK\s+(\d+)\s+(\S+)\s*$')
    if (-not $m.Success) { return $null }
    return [pscustomobject]@{ Number = [int] $m.Groups[1].Value; Id = $m.Groups[2].Value }
}

<#
    The probe message. Deliberately awkward in the ways a POP3 path gets wrong, because a probe
    that only proves "hello world survives" proves nothing about Outlook's mail. The Date header
    is formatted with the invariant culture: the guests run nl-NL formats on purpose.
#>
function New-ProbeMessage {
    param([string] $Address, [string] $Marker, [datetime] $DateUtc)

    $subject = "$ProbeSubjectTag sink probe $Marker"
    $date = $DateUtc.ToString('ddd, dd MMM yyyy HH:mm:ss +0000', [System.Globalization.CultureInfo]::InvariantCulture)
    $boundary = "sinkprobe$Marker"
    $attachmentB64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("OutlookAI sink probe attachment $Marker"))

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
        # (1) EXACTLY a period. Un-stuffed by a broken path this ends the message early.
        '.'
        # (2) One leading period - the classic dot-stuffing bug.
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
        DateHeader    = "Date: $date"
    }
}

# What is wrong with a retrieved probe, as a list. Empty means it came back intact.
function Get-ProbeBodyProblems {
    param($Message, [string[]] $Retrieved)

    $problems = @()
    $body = @($Retrieved)
    if (@($body | Where-Object { $_ -ceq ('Subject: ' + $Message.Subject) }).Count -ne 1) {
        $problems += 'the subject header did not come back exactly once'
    }
    $bare   = @($body | Where-Object { $_ -ceq '.' }).Count
    $single = @($body | Where-Object { $_ -ceq '.leading period must survive' }).Count
    $double = @($body | Where-Object { $_ -ceq '..two leading periods must survive' }).Count
    if ($bare -ne 1 -or $single -ne 1 -or $double -ne 1) {
        $problems += "dot-stuffing did not survive: bare-period lines=$bare, single-leading=$single, double-leading=$double (each must be 1)"
    }
    if (@($body | Where-Object { $_ -ceq ('marker ' + $Message.Marker) }).Count -ne 1) {
        $problems += 'the marker line after the dotted lines is missing - the body was cut short'
    }
    if (@($body | Where-Object { $_ -ceq $Message.AttachmentB64 }).Count -ne 1) {
        $problems += 'the base64 attachment part did not come back byte-identical'
    }
    return $problems
}

<#
    The scheduled task, audited from a snapshot so the rules can be tested without Task
    Scheduler. Each rule is one a default gets wrong: a task with the default time limit is
    KILLED after three days, a task registered for a user only runs while that user is logged
    on (and then draws a console window), and a logon trigger is not a boot trigger.
#>
function Get-TaskDefinitionProblems {
    param($Snapshot, [string] $LauncherPath)

    $problems = @()
    if ($null -eq $Snapshot -or -not $Snapshot.Exists) { return 'the scheduled task does not exist' }
    if (-not $Snapshot.Enabled) { $problems += 'the task is disabled, so the sink will not start with the guest' }

    $triggers = @($Snapshot.TriggerKinds)
    if ($triggers.Count -ne 1 -or $triggers[0] -cne 'MSFT_TaskBootTrigger') {
        $problems += "the task must have exactly one trigger, at startup; it has: $($triggers -join ', ')"
    }

    $system = @('SYSTEM', 'S-1-5-18', 'NT AUTHORITY\SYSTEM')
    $userOk = $false
    foreach ($s in $system) {
        if ([string]::Equals([string] $Snapshot.UserId, $s, [System.StringComparison]::OrdinalIgnoreCase)) { $userOk = $true }
    }
    if (-not $userOk) { $problems += "the task runs as '$($Snapshot.UserId)', not SYSTEM - it would start only with that user's logon, and draw a console window when it did" }

    $actionCount = [int] $Snapshot.ActionCount
    if ($actionCount -ne 1) { $problems += "the task must have exactly one action; it has $actionCount" }
    $exe = [string] $Snapshot.Execute
    if (-not $exe.EndsWith('\cmd.exe', [System.StringComparison]::OrdinalIgnoreCase)) { $problems += "the task runs '$exe', not cmd.exe" }
    $actionArguments = [string] $Snapshot.Arguments
    if ($LauncherPath -and ($actionArguments.IndexOf($LauncherPath, [System.StringComparison]::OrdinalIgnoreCase) -lt 0)) {
        $problems += "the task's arguments '$actionArguments' do not name the launcher $LauncherPath"
    }

    $limit = [string] $Snapshot.ExecutionTimeLimit
    if ($limit -and $limit -cne 'PT0S') {
        $problems += "the task has an execution time limit of $limit. Task Scheduler's default is three days (PT72H), after which it KILLS the sink; it must be PT0S"
    }
    if ([string] $Snapshot.MultipleInstances -cne 'IgnoreNew') {
        $problems += "MultipleInstances is '$($Snapshot.MultipleInstances)'; it must be IgnoreNew, or a second start races the first for the ports"
    }
    return $problems
}

<#
    -f formats in the CURRENT culture, and the guests run nl-NL formats on purpose
    (Testbed/MEDIA.md, "The host configuration the guests match") - so 1.5 prints as "1,5" and
    1234 as "1.234". Every number or date this script prints goes through here instead, so a
    guest's log reads the same as the host's and as the documentation.
#>
function Format-Invariant {
    param([string] $Format, [object[]] $Values)
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, $Format, $Values)
}

# =============================================================================================
# LOGGING AND GUARDS.
# =============================================================================================

function Say {
    param([string] $Text)
    Write-Host $Text
    if (-not $script:LogReady) { return }
    try { Add-Content -LiteralPath $LogPath -Value $Text -Encoding UTF8 } catch { }
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

function Note {
    param([string] $Text)
    Say ("  NOTE $Text")
}

function Open-Log {
    $dir = Split-Path -Parent $LogPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $script:LogReady = $true
    $given = @()
    foreach ($k in $ScriptBoundParameters.Keys) { $given += ('-' + $k) }
    Say ''
    Say (Format-Invariant 'Install-MailSink.ps1 - {0:yyyy-MM-dd HH:mm:ss} - {1} on {2} as {3}' @((Get-Date), ($given -join ' '), $env:COMPUTERNAME, $env:USERNAME))
}

function Assert-TestbedGuestLocal {
    $who = $env:USERNAME
    $userOk = $false
    foreach ($candidate in $ExpectedUser) { if ($who -eq $candidate) { $userOk = $true } }
    $machine = $env:COMPUTERNAME
    $machineOk = [bool] ($ExpectedComputerNamePrefix -and $machine -and
        $machine.StartsWith($ExpectedComputerNamePrefix, [System.StringComparison]::OrdinalIgnoreCase))
    if ($userOk -and $machineOk) { return }

    throw @"
REFUSING TO RUN.

  logged on as : '$who'          (allowed: $($ExpectedUser -join ', '))
  computer name: '$machine'      (must start with: '$ExpectedComputerNamePrefix')

This script registers a SYSTEM task that listens on the SMTP and POP3 ports and submits mail to
it. On the maintainer's workstation that would put an unauthenticated listener on port $SmtpPort
of a machine holding real mail. The testbed guests autologon as 'vmadmin' and are named 'OAI-*'
by Testbed/host/New-AnswerFile.ps1. If you built a guest differently, say so:

    -ExpectedUser <username> -ExpectedComputerNamePrefix <prefix>

Do not 'fix' this by widening either default. The defaults are the guard.
"@
}

function Assert-Elevated {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'REFUSING TO RUN: this session is not elevated. A SYSTEM task can only be registered, started and stopped by an administrator; asserted here rather than failing halfway with an access-denied that reads like a packaging fault.'
    }
}

function Test-OutlookRunning {
    return (@(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue).Count -gt 0)
}

function Assert-OutlookNotRunning {
    $running = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return }
    throw @"
REFUSING: OUTLOOK.EXE is running (pid $(($running | ForEach-Object { $_.Id }) -join ', ')).

Installing or removing the sink stops and starts its listeners, and a polling Outlook that finds
them gone turns a quiet send/receive into an error. Close Outlook properly and run this again.
DO NOT taskkill it - mailbox-safety rule 7 forbids that outright. (-Verify alone may run while
Outlook is open: it only uses mailboxes no account reads, and it skips its restart step.)
"@
}

# =============================================================================================
# READING THE MACHINE. Read-only.
# =============================================================================================

<#
    Runs a native program and hands back its stdout lines, its stderr lines and its exit code.

    WHY THIS EXISTS, MEASURED ON THIS PROJECT'S HOST 2026-09-24: under Windows PowerShell 5.1 with
    $ErrorActionPreference = 'Stop', the FIRST line a native program writes to stderr becomes a
    terminating NativeCommandError - for 2>$null, 2>&1 and *> alike. PowerShell 7 does not do
    that, which is how it goes unnoticed on a workstation and then kills a guest run, where 5.1 is
    the only PowerShell there is. So every native call in this file comes through here: it runs
    under 'Continue', separates the ErrorRecords that stderr lines arrive as from the strings that
    stdout lines arrive as, and is judged by the exit code - never by whether PowerShell complained.
#>
function Invoke-NativeLines {
    param([Parameter(Mandatory = $true)] [string] $FilePath, [string[]] $Arguments = @())
    $saved = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $raw = @()
    $code = $null
    try {
        $raw = @(& $FilePath @Arguments 2>&1)
        $code = $LASTEXITCODE
    }
    catch {
        return [pscustomobject]@{ ExitCode = -1; Lines = @(); ErrorLines = @([string] $_.Exception.Message) }
    }
    finally {
        $ErrorActionPreference = $saved
    }
    $out = @()
    $err = @()
    foreach ($item in $raw) {
        if ($item -is [System.Management.Automation.ErrorRecord]) { $err += $item.ToString() }
        else { $out += [string] $item }
    }
    return [pscustomobject]@{ ExitCode = $code; Lines = $out; ErrorLines = $err }
}

# Readable = $false is "could not ask", which must not collapse into "nothing is reserved".
function Get-ExcludedPortRange {
    $r = Invoke-NativeLines -FilePath 'netsh.exe' -Arguments @('interface', 'ipv4', 'show', 'excludedportrange', 'protocol=tcp')
    if ($r.ExitCode -ne 0) { return [pscustomobject]@{ Readable = $false; Ranges = @() } }
    return [pscustomobject]@{ Readable = $true; Ranges = @(Get-ExcludedPortRangeFromText -Lines $r.Lines) }
}

# What is listening on a port and on which addresses. The address is the point: a listener that
# is not on 127.0.0.1 is reachable from the host.
function Get-ListeningEndpoint {
    param([int] $Port)

    $conn = @()
    if (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue) {
        $conn = @(Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction SilentlyContinue |
                ForEach-Object { [pscustomobject]@{ LocalAddress = [string] $_.LocalAddress; OwningProcess = [int] $_.OwningProcess } })
    }
    else {
        foreach ($line in @((Invoke-NativeLines -FilePath 'netstat.exe' -Arguments @('-ano', '-p', 'tcp')).Lines)) {
            $m = [regex]::Match([string] $line, '^\s*TCP\s+(\S+):(\d+)\s+\S+\s+LISTENING\s+(\d+)\s*$')
            if ($m.Success -and [int] $m.Groups[2].Value -eq $Port) {
                $conn += [pscustomobject]@{ LocalAddress = $m.Groups[1].Value; OwningProcess = [int] $m.Groups[3].Value }
            }
        }
    }
    if ($conn.Count -eq 0) { return $null }

    $name = '(exited)'
    try { $name = (Get-Process -Id $conn[0].OwningProcess -ErrorAction Stop).ProcessName } catch { }
    return [pscustomobject]@{
        OwningProcess = [int] $conn[0].OwningProcess
        ProcessName   = $name
        AllAddresses  = @($conn | ForEach-Object { $_.LocalAddress })
        Pids          = @($conn | ForEach-Object { $_.OwningProcess } | Sort-Object -Unique)
    }
}

function Get-SinkLayout {
    # Exactly one inbucket.exe, one folder below the install root - where -Execute unpacks it.
    $root = $InstallRoot.TrimEnd('\')
    $exes = @()
    if (Test-Path -LiteralPath $root) {
        foreach ($dir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
            $candidate = Join-Path $dir.FullName $ExeName
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $exes += $candidate }
        }
    }
    $exe = $null
    $appDir = $null
    if ($exes.Count -eq 1) {
        $exe = $exes[0]
        $appDir = Split-Path -Parent $exe
    }
    return [pscustomobject]@{
        Root         = $root
        Exe          = $exe
        AppDir       = $appDir
        ExeCount     = $exes.Count
        Launcher     = Join-Path $root $LauncherName
        Store        = Join-Path $root $StoreDirName
        SinkLog      = Join-Path $root $SinkLogName
    }
}

# The sink's processes: inbucket.exe running from -ExePath exactly - or, with -AnyUnderRoot, from
# anywhere under the install root. Stopping must use the second: an earlier install may have
# unpacked a different version into a different folder, and its process holds the ports all the
# same. Nothing outside the install root is ever matched, whatever it is called.
function Get-SinkProcess {
    param([string] $ExePath, [switch] $AnyUnderRoot)
    $rootPrefix = $InstallRoot.TrimEnd('\') + '\'
    return @(Get-CimInstance -ClassName Win32_Process -Filter "Name='$ExeName'" -ErrorAction SilentlyContinue |
            Where-Object {
                $path = [string] $_.ExecutablePath
                if (-not $path) { $false }
                elseif ($AnyUnderRoot) { $path.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase) }
                else { [bool] $ExePath -and [string]::Equals($path, $ExePath, [System.StringComparison]::OrdinalIgnoreCase) }
            })
}

function Get-TaskSnapshot {
    $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -eq $t) { return [pscustomobject]@{ Exists = $false } }
    $actions = @($t.Actions)
    $first = $null
    if ($actions.Count -gt 0) { $first = $actions[0] }
    return [pscustomobject]@{
        Exists             = $true
        State              = [string] $t.State
        Enabled            = [bool] $t.Settings.Enabled
        TriggerKinds       = @($t.Triggers | ForEach-Object { $_.CimClass.CimClassName })
        UserId             = [string] $t.Principal.UserId
        ActionCount        = $actions.Count
        Execute            = [string] $(if ($first) { $first.Execute } else { '' })
        Arguments          = [string] $(if ($first) { $first.Arguments } else { '' })
        ExecutionTimeLimit = [string] $t.Settings.ExecutionTimeLimit
        MultipleInstances  = [string] $t.Settings.MultipleInstances
    }
}

# =============================================================================================
# STARTING AND STOPPING. Through the scheduled task, because that is how the guest starts it.
# =============================================================================================

function Wait-SinkListeners {
    param([string] $ExePath)
    $deadline = (Get-Date).AddSeconds($ListenerWaitSeconds)
    while ((Get-Date) -lt $deadline) {
        $procs = @(Get-SinkProcess -ExePath $ExePath)
        if ($procs.Count -gt 0) {
            $pids = @($procs | ForEach-Object { [int] $_.ProcessId })
            $all = $true
            foreach ($port in @($SmtpPort, $Pop3Port, $WebPort)) {
                $ep = Get-ListeningEndpoint -Port $port
                if ($null -eq $ep -or -not ($pids -contains $ep.OwningProcess)) { $all = $false }
            }
            if ($all) { return $true }
        }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

function Stop-Sink {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $task) { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue }

    # Ending the task ends cmd.exe; whether Windows also ends the child it started is not
    # something to rely on. Our own exe, from under the install root, and nothing else.
    foreach ($p in @(Get-SinkProcess -AnyUnderRoot)) {
        Stop-Process -Id ([int] $p.ProcessId) -Force -ErrorAction SilentlyContinue
    }

    $deadline = (Get-Date).AddSeconds($ListenerWaitSeconds)
    while ((Get-Date) -lt $deadline) {
        $busy = $false
        if (@(Get-SinkProcess -AnyUnderRoot).Count -gt 0) { $busy = $true }
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($null -ne $task -and [string] $task.State -eq 'Running') { $busy = $true }
        if (-not $busy) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

function Start-Sink {
    param([string] $ExePath)
    Start-ScheduledTask -TaskName $TaskName
    return (Wait-SinkListeners -ExePath $ExePath)
}

# =============================================================================================
# THE SOCKET LAYER. Raw SMTP and POP3, because the point of -Verify is to speak the protocols
# rather than to find two open ports.
# =============================================================================================

function Connect-SinkSocket {
    param([int] $Port)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($SinkHost, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($SocketTimeoutMs)) {
            throw "Nothing answered ${SinkHost}:$Port within $SocketTimeoutMs ms."
        }
        $client.EndConnect($iar)
    }
    catch {
        $client.Close()
        throw
    }
    $client.ReceiveTimeout = $SocketTimeoutMs
    $client.SendTimeout = $SocketTimeoutMs

    # Latin-1 maps every byte to one char and back, so nothing is lost or invented in transit.
    $encoding = [System.Text.Encoding]::GetEncoding(28591)
    $stream = $client.GetStream()
    $reader = New-Object System.IO.StreamReader($stream, $encoding)
    $writer = New-Object System.IO.StreamWriter($stream, $encoding)
    $writer.NewLine = "`r`n"
    $writer.AutoFlush = $true
    return [pscustomobject]@{ Client = $client; Reader = $reader; Writer = $writer }
}

function Close-SinkSocket {
    param($Socket)
    if ($null -eq $Socket) { return }
    try { $Socket.Writer.Dispose() } catch { }
    try { $Socket.Reader.Dispose() } catch { }
    try { $Socket.Client.Close() } catch { }
}

function Read-SmtpReply {
    param($Socket)
    $lines = @()
    for ($i = 0; $i -lt 50; $i++) {
        $line = $Socket.Reader.ReadLine()
        if ($null -eq $line) { throw 'The SMTP server closed the connection mid-reply.' }
        $lines += $line
        if ($line.Length -lt 4 -or $line[3] -ne '-') { break }
    }
    return (ConvertFrom-SmtpReplyLines -Lines $lines)
}

function Invoke-SmtpCommand {
    param($Socket, [string] $Command, [int] $ExpectCode)
    $Socket.Writer.WriteLine($Command)
    $reply = Read-SmtpReply -Socket $Socket
    if ($reply.Code -ne $ExpectCode) { throw "SMTP '$Command' expected $ExpectCode, got $($reply.Code): $($reply.Text)" }
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
    if (-not $reply.Ok -and -not $AllowError) { throw "POP3 '$Command' failed: $($reply.Text)" }
    return $reply
}

function Read-Pop3MultiLine {
    param($Socket)
    $raw = @()
    for ($i = 0; $i -lt 100000; $i++) {
        $line = $Socket.Reader.ReadLine()
        if ($null -eq $line) { break }
        $raw += $line
        if ($line -ceq '.') { break }
    }
    $decoded = ConvertFrom-DotStuffedLines -Lines $raw
    if (-not $decoded.Terminated) { throw 'The POP3 server ended a multi-line response without its terminating ".".' }
    return $decoded.Lines
}

<#
    Submits one probe, dot-stuffing on the way out. Throws on any unexpected reply: a partial
    submission is not a state worth continuing from.
#>
function Submit-ProbeMessage {
    param([string] $Address, $Message)
    $socket = $null
    try {
        $socket = Connect-SinkSocket -Port $SmtpPort
        $greeting = Read-SmtpReply -Socket $socket
        if ($greeting.Code -ne 220) { throw "The SMTP greeting was not 220: $($greeting.Text)" }
        [void](Invoke-SmtpCommand -Socket $socket -Command 'EHLO sink-probe.vm.invalid' -ExpectCode 250)
        [void](Invoke-SmtpCommand -Socket $socket -Command "MAIL FROM:<$Address>" -ExpectCode 250)
        [void](Invoke-SmtpCommand -Socket $socket -Command "RCPT TO:<$Address>" -ExpectCode 250)
        [void](Invoke-SmtpCommand -Socket $socket -Command 'DATA' -ExpectCode 354)
        foreach ($line in (ConvertTo-DotStuffedLines -Lines $Message.Lines)) { $socket.Writer.WriteLine($line) }
        $accepted = Invoke-SmtpCommand -Socket $socket -Command '.' -ExpectCode 250
        $socket.Writer.WriteLine('QUIT')
        return $accepted.Text
    }
    finally {
        Close-SinkSocket -Socket $socket
    }
}

<#
    Opens a POP3 session on one mailbox with one of three password shapes. None, Empty and
    NonEmpty are the three things a client with no stored password might plausibly send: the
    bare command, the command with an empty argument, or whatever it holds.
#>
function Open-Pop3Session {
    param([string] $Mailbox, [ValidateSet('None', 'Empty', 'NonEmpty')] [string] $PassShape = 'Empty')
    $socket = Connect-SinkSocket -Port $Pop3Port
    try {
        $greeting = Read-Pop3Reply -Socket $socket
        if (-not $greeting.Ok) { throw "The POP3 greeting was not +OK: $($greeting.Text)" }
        [void](Invoke-Pop3Command -Socket $socket -Command "USER $Mailbox")
        $command = 'PASS '
        if ($PassShape -eq 'None') { $command = 'PASS' }
        if ($PassShape -eq 'NonEmpty') { $command = "PASS $NonEmptyPassValue" }
        $pass = Invoke-Pop3Command -Socket $socket -Command $command -AllowError
        if (-not $pass.Ok) { throw "USER $Mailbox then '$command' was refused: $($pass.Text)" }
    }
    catch {
        Close-SinkSocket -Socket $socket
        throw
    }
    return [pscustomobject]@{ Socket = $socket; Mailbox = $Mailbox }
}

function Close-Pop3Session {
    param($Session, [switch] $WithoutQuit)
    if ($null -eq $Session) { return }
    if (-not $WithoutQuit) {
        try { [void](Invoke-Pop3Command -Socket $Session.Socket -Command 'QUIT' -AllowError) } catch { }
    }
    Close-SinkSocket -Socket $Session.Socket
}

function Get-Pop3StatOf {
    param($Session)
    $reply = Invoke-Pop3Command -Socket $Session.Socket -Command 'STAT'
    $stat = ConvertFrom-Pop3Stat -Line $reply.Text
    if ($null -eq $stat) { throw "Could not read the STAT reply: $($reply.Text)" }
    return $stat
}

function Get-Pop3UidOf {
    param($Session, [int] $Number)
    $reply = Invoke-Pop3Command -Socket $Session.Socket -Command "UIDL $Number" -AllowError
    return (ConvertFrom-Pop3Uidl -Line $reply.Text)
}

# THE ONLY DELE IN THIS FILE. See Assert-ProbeMailboxForDelete.
function Remove-Pop3Message {
    param($Session, [int] $Number, [string[]] $ProbeMailboxes)
    Assert-ProbeMailboxForDelete -Mailbox $Session.Mailbox -ProbeMailboxes $ProbeMailboxes
    return (Invoke-Pop3Command -Socket $Session.Socket -Command "DELE $Number" -AllowError)
}

# Empties a probe mailbox and reports how much it held. A non-zero count is residue from an
# earlier run that died mid-proof - reported, because it says something went wrong once.
function Clear-ProbeMailbox {
    param([string] $Mailbox, [string[]] $ProbeMailboxes)
    $session = Open-Pop3Session -Mailbox $Mailbox
    try {
        $stat = Get-Pop3StatOf -Session $session
        for ($n = 1; $n -le $stat.Count; $n++) {
            $del = Remove-Pop3Message -Session $session -Number $n -ProbeMailboxes $ProbeMailboxes
            if (-not $del.Ok) { throw "DELE $n in '$Mailbox' was refused: $($del.Text)" }
        }
        [void](Invoke-Pop3Command -Socket $session.Socket -Command 'QUIT')
        return $stat.Count
    }
    finally {
        Close-SinkSocket -Socket $session.Socket
    }
}

function Get-MailboxCount {
    param([string] $Mailbox)
    $session = Open-Pop3Session -Mailbox $Mailbox
    try { return (Get-Pop3StatOf -Session $session).Count }
    finally { Close-Pop3Session -Session $session }
}

# =============================================================================================
# THE PROOFS. Each check has a name, and failures are COLLECTED rather than thrown, so one run
# reports every fault instead of the first.
# =============================================================================================

function New-ProbeMarker {
    return [Guid]::NewGuid().ToString('N').Substring(0, 12)
}

function Test-LoginShapes {
    param([string] $Mailbox)
    Say ''
    Say "POP3 logins with no stored password - mailbox '$Mailbox':"
    foreach ($shape in @('Empty', 'None', 'NonEmpty')) {
        $label = @{ Empty = "'PASS ' (an empty argument)"; None = "'PASS' (no argument at all)"; NonEmpty = "'PASS <any value>'" }[$shape]
        $session = $null
        try {
            $session = Open-Pop3Session -Mailbox $Mailbox -PassShape $shape
            [void](Get-Pop3StatOf -Session $session)
            Pass "POP3 accepts $label" 'and the session reaches TRANSACTION'
        }
        catch {
            Fail "POP3 accepts $label" ($_.Exception.Message + ' - the tier account stores no password, so a sink that refuses this makes Outlook prompt, and on an unattended guest a prompt is a hang.')
        }
        finally {
            Close-Pop3Session -Session $session
        }
    }
}

function Invoke-RoundTripProof {
    param([string] $Address, [string] $Mailbox, [string[]] $ProbeMailboxes)
    Say ''
    Say "Round trip: SMTP ${SinkHost}:$SmtpPort -> mailbox '$Mailbox' -> POP3 ${SinkHost}:$Pop3Port."

    $message = New-ProbeMessage -Address $Address -Marker (New-ProbeMarker) -DateUtc ((Get-Date).ToUniversalTime())
    $started = Get-Date
    try {
        $accepted = Submit-ProbeMessage -Address $Address -Message $message
        Pass 'the sink accepts a loopback submission' $accepted
    }
    catch {
        Fail 'the sink accepts a loopback submission' $_.Exception.Message
        return
    }

    $session = $null
    try {
        $session = Open-Pop3Session -Mailbox $Mailbox
        $stat = Get-Pop3StatOf -Session $session
        if ($stat.Count -ne 1) {
            Fail 'STAT reports exactly the submitted message' "STAT says $($stat.Count) message(s); the mailbox was emptied first, so it must be 1. Zero means the submission was accepted and dropped."
            return
        }
        Pass 'STAT reports exactly the submitted message' "1 message, $($stat.Octets) octets announced"

        $capa = Invoke-Pop3Command -Socket $session.Socket -Command 'CAPA' -AllowError
        if ($capa.Ok) {
            $caps = @(Read-Pop3MultiLine -Socket $session.Socket)
            Say "  ...  CAPA advertises: $($caps -join ', ')"
            if (@($caps | Where-Object { $_ -match '^STLS\b' }).Count -gt 0) {
                Fail 'CAPA offers no STLS' 'The tier profile sets POP3UseSSL=0; a sink that offers STLS is one Outlook may try to negotiate with.'
            }
            else { Pass 'CAPA offers no STLS' 'plain POP3, as the tier profile expects' }
        }
        else { Note "CAPA is not supported ($($capa.Text)); nothing advertised can be checked against behaviour." }

        $uid = Get-Pop3UidOf -Session $session -Number 1
        if ($null -eq $uid) { Fail 'UIDL answers for the message' 'UIDL 1 did not return an id. Outlook keys its already-seen state on UIDL.' }
        else {
            Pass 'UIDL answers for the message' $uid.Id
            $script:SeenIds += $uid.Id
        }

        $top = Invoke-Pop3Command -Socket $session.Socket -Command 'TOP 1 0' -AllowError
        if ($top.Ok) {
            $head = @(Read-Pop3MultiLine -Socket $session.Socket)
            $hasSubject = @($head | Where-Object { $_ -ceq ('Subject: ' + $message.Subject) }).Count -eq 1
            $hasBody = @($head | Where-Object { $_ -ceq ('marker ' + $message.Marker) }).Count -gt 0
            if ($hasSubject -and -not $hasBody) { Pass 'TOP 1 0 returns the headers and no body' "$($head.Count) line(s)" }
            else { Fail 'TOP 1 0 returns the headers and no body' "subject present=$hasSubject, body present=$hasBody" }
        }
        else { Fail 'TOP 1 0 returns the headers and no body' "TOP was refused: $($top.Text). CAPA advertises it, so a client that trusts CAPA would hit this." }

        [void](Invoke-Pop3Command -Socket $session.Socket -Command 'RETR 1')
        $body = @(Read-Pop3MultiLine -Socket $session.Socket)
        $problems = @(Get-ProbeBodyProblems -Message $message -Retrieved $body)
        if ($problems.Count -eq 0) {
            Pass 'RETR returns the message intact' "$($body.Count) line(s): subject, a bare '.', '.x', '..x' and the base64 part all unchanged"
        }
        else {
            Fail 'RETR returns the message intact' (($problems -join '; ') + '. This truncates or alters mail silently, which is the worst failure shape a test sink can have.')
        }

        $del = Remove-Pop3Message -Session $session -Number 1 -ProbeMailboxes $ProbeMailboxes
        if (-not $del.Ok) { Fail 'DELE is accepted' $del.Text }
        [void](Invoke-Pop3Command -Socket $session.Socket -Command 'QUIT')
    }
    catch {
        Fail 'the round trip completes' $_.Exception.Message
        return
    }
    finally {
        Close-SinkSocket -Socket $session.Socket
    }

    try {
        $after = Get-MailboxCount -Mailbox $Mailbox
        if ($after -eq 0) { Pass 'DELE is honoured at QUIT' 'the mailbox is empty on reconnect' }
        else { Fail 'DELE is honoured at QUIT' "STAT still says $after after DELE and QUIT. The tier profile leaves nothing on the server and relies on the sink draining; one that ignores DELE re-delivers the same mail for ever." }
    }
    catch { Fail 'DELE is honoured at QUIT' $_.Exception.Message }

    $elapsed = ((Get-Date) - $started).TotalSeconds
    if ($elapsed -le $RoundTripBudgetSeconds) { Pass 'the round trip fits its budget' (Format-Invariant '{0:N1} s of {1} s' @($elapsed, $RoundTripBudgetSeconds)) }
    else { Fail 'the round trip fits its budget' (Format-Invariant '{0:N1} s against {1} s. The suite''s tightest arrival deadline is 120 s and has to cover Outlook too.' @($elapsed, $RoundTripBudgetSeconds)) }
}

function Invoke-IsolationProof {
    param([string] $Address, [string] $Mailbox, [string] $OtherMailbox, [string[]] $ProbeMailboxes)
    Say ''
    Say "Isolation: mail for '$Mailbox' must be invisible to '$OtherMailbox'."
    try {
        [void](Submit-ProbeMessage -Address $Address -Message (New-ProbeMessage -Address $Address -Marker (New-ProbeMarker) -DateUtc ((Get-Date).ToUniversalTime())))
        $mine = Get-MailboxCount -Mailbox $Mailbox
        $theirs = Get-MailboxCount -Mailbox $OtherMailbox
        if ($mine -eq 1 -and $theirs -eq 0) {
            Pass 'a message is visible only to its own mailbox' "'$Mailbox' holds 1, '$OtherMailbox' holds 0"
        }
        else {
            Fail 'a message is visible only to its own mailbox' ("'$Mailbox' holds $mine (expected 1), '$OtherMailbox' holds $theirs (expected 0). " +
                'A sink that shows every login every message lets the dummy and identity accounts take each other''s mail, whichever polls first.')
        }
    }
    catch { Fail 'a message is visible only to its own mailbox' $_.Exception.Message }
    finally {
        try { [void](Clear-ProbeMailbox -Mailbox $Mailbox -ProbeMailboxes $ProbeMailboxes) } catch { Fail 'the probe mailbox drains after the isolation proof' $_.Exception.Message }
    }
}

<#
    RFC 1939 section 5: message numbers are fixed for the whole session and deletes happen at
    QUIT. The failure it rules out is silent - a sink that renumbers after DELE makes a client's
    "DELE 1; DELE 2" delete the wrong message.
#>
function Invoke-OrdinalStabilityProof {
    param([string] $Address, [string] $Mailbox, [string[]] $ProbeMailboxes)
    Say ''
    Say 'Ordinal stability: two messages, DELE 1 - does the survivor keep its number?'
    $session = $null
    try {
        foreach ($i in 1..2) {
            [void](Submit-ProbeMessage -Address $Address -Message (New-ProbeMessage -Address $Address -Marker (New-ProbeMarker) -DateUtc ((Get-Date).ToUniversalTime())))
        }
        $session = Open-Pop3Session -Mailbox $Mailbox
        $stat = Get-Pop3StatOf -Session $session
        if ($stat.Count -ne 2) {
            Fail 'both messages are held' "STAT says $($stat.Count); expected 2."
            return
        }
        $u2 = Get-Pop3UidOf -Session $session -Number 2
        if ($null -eq $u2) { Fail 'ordinal stability can be measured' 'UIDL 2 returned no id.'; return }
        $script:SeenIds += $u2.Id
        $u1 = Get-Pop3UidOf -Session $session -Number 1
        if ($null -ne $u1) { $script:SeenIds += $u1.Id }

        $del = Remove-Pop3Message -Session $session -Number 1 -ProbeMailboxes $ProbeMailboxes
        if (-not $del.Ok) { Fail 'ordinal stability can be measured' "DELE 1 was refused: $($del.Text)"; return }

        $after2 = Get-Pop3UidOf -Session $session -Number 2
        $after1 = Get-Pop3UidOf -Session $session -Number 1
        if ($null -ne $after2 -and $after2.Id -ceq $u2.Id -and $null -eq $after1) {
            Pass 'message numbers stay fixed after a DELE' 'the survivor is still message 2 and message 1 answers -ERR, as RFC 1939 requires'
        }
        elseif ($null -ne $after1 -and $after1.Id -ceq $u2.Id) {
            Fail 'message numbers stay fixed after a DELE' 'After DELE 1 the survivor became message 1: this sink renumbers mid-session, so a client issuing DELE 1; DELE 2 deletes the wrong message.'
        }
        else {
            Fail 'message numbers stay fixed after a DELE' "After DELE 1: UIDL 1 -> $(if ($after1) { $after1.Id } else { '-ERR' }), UIDL 2 -> $(if ($after2) { $after2.Id } else { '-ERR' }); expected -ERR and $($u2.Id)."
        }

        $stat = Get-Pop3StatOf -Session $session
        if ($stat.Count -eq 1) { Pass 'STAT stops counting a message marked for deletion' '1 left' }
        else { Fail 'STAT stops counting a message marked for deletion' "STAT says $($stat.Count); expected 1." }

        $del = Remove-Pop3Message -Session $session -Number 2 -ProbeMailboxes $ProbeMailboxes
        if (-not $del.Ok) { Fail 'the survivor can be deleted by its original number' $del.Text }
        [void](Invoke-Pop3Command -Socket $session.Socket -Command 'QUIT')
    }
    catch { Fail 'ordinal stability can be measured' $_.Exception.Message }
    finally {
        if ($session) { Close-SinkSocket -Socket $session.Socket }
        try {
            $left = Get-MailboxCount -Mailbox $Mailbox
            if ($left -eq 0) { Pass 'the probe mailbox is empty afterwards' 'both deletes landed at QUIT' }
            else {
                Fail 'the probe mailbox is empty afterwards' "$left message(s) remain; draining them."
                [void](Clear-ProbeMailbox -Mailbox $Mailbox -ProbeMailboxes $ProbeMailboxes)
            }
        }
        catch { Fail 'the probe mailbox is empty afterwards' $_.Exception.Message }
    }
}

# A session that dies after DELE and before QUIT must lose nothing: the delete belongs to the
# UPDATE state, which only QUIT reaches. This is what makes an interrupted fetch re-fetchable
# rather than lost.
function Invoke-AbortSemanticsProof {
    param([string] $Address, [string] $Mailbox, [string[]] $ProbeMailboxes)
    Say ''
    Say 'Abort semantics: DELE, then drop the connection without QUIT.'
    try {
        [void](Submit-ProbeMessage -Address $Address -Message (New-ProbeMessage -Address $Address -Marker (New-ProbeMarker) -DateUtc ((Get-Date).ToUniversalTime())))
        $session = Open-Pop3Session -Mailbox $Mailbox
        $uid = Get-Pop3UidOf -Session $session -Number 1
        if ($uid) { $script:SeenIds += $uid.Id }
        [void](Remove-Pop3Message -Session $session -Number 1 -ProbeMailboxes $ProbeMailboxes)
        Close-Pop3Session -Session $session -WithoutQuit
        Start-Sleep -Milliseconds 500
        $count = Get-MailboxCount -Mailbox $Mailbox
        if ($count -eq 1) { Pass 'a DELE without QUIT deletes nothing' 'the message is still there after the dropped session' }
        else { Fail 'a DELE without QUIT deletes nothing' "The mailbox holds $count; expected 1. A sink that deletes before QUIT loses mail whenever a fetch is interrupted." }
    }
    catch { Fail 'a DELE without QUIT deletes nothing' $_.Exception.Message }
    finally {
        try { [void](Clear-ProbeMailbox -Mailbox $Mailbox -ProbeMailboxes $ProbeMailboxes) } catch { Fail 'the probe mailbox drains after the abort proof' $_.Exception.Message }
    }
}

# Read-only. Mail waiting in an account's mailbox is mail Outlook has not collected yet; outside
# a run that is an earlier run's undelivered residue. It is reported and NEVER deleted here.
function Show-AccountMailboxes {
    Say ''
    Say 'Account mailboxes (read-only - opened, counted, closed; never emptied):'
    foreach ($address in @($AccountAddress)) {
        if (-not $address) { continue }
        try {
            $box = Get-InbucketMailboxName -Address $address
            $count = Get-MailboxCount -Mailbox $box
            $tail = 'nothing waiting'
            if ($count -gt 0) { $tail = "$count message(s) waiting for Outlook - if no run is in progress, that is an earlier run's undelivered mail" }
            Say ("  ...  {0,-24} mailbox '{1}': {2}" -f $address, $box, $tail)
        }
        catch { Note "could not read the mailbox for ${address}: $($_.Exception.Message)" }
    }
    Say '       An account''s POP3 user name must be exactly its mailbox name above - lowercase, no'
    Say '       domain, no +tag. The sink uses the USER verbatim as the mailbox to open.'
}

function Invoke-RestartProof {
    param($Layout, [string] $Address, [string] $Mailbox, [string[]] $ProbeMailboxes)
    Say ''
    Say "Restart through the scheduled task '$TaskName'."
    if (Test-OutlookRunning) {
        Note 'SKIPPED: Outlook is running, and restarting the sink under a polling Outlook is the nondeterminism this avoids. Close Outlook and re-run -Verify to prove the restart.'
        return
    }

    if (-not (Stop-Sink)) {
        Fail 'the sink stops' "It was still running $ListenerWaitSeconds s after the task was ended and its process stopped."
        return
    }
    if (-not (Start-Sink -ExePath $Layout.Exe)) {
        Fail 'the sink starts through its scheduled task' "Its three listeners were not all up $ListenerWaitSeconds s after Start-ScheduledTask. Read $($Layout.SinkLog)."
        return
    }
    Pass 'the sink starts through its scheduled task' "listening on $SmtpPort, $Pop3Port and $WebPort again"

    try {
        $count = Get-MailboxCount -Mailbox $Mailbox
        if ($count -eq 0) { Pass 'nothing deleted comes back after a restart' 'the probe mailbox is still empty' }
        else { Fail 'nothing deleted comes back after a restart' "The probe mailbox holds $count after a restart. A sink that resurrects deleted mail re-delivers every past message into the hub Inbox." }

        [void](Submit-ProbeMessage -Address $Address -Message (New-ProbeMessage -Address $Address -Marker (New-ProbeMarker) -DateUtc ((Get-Date).ToUniversalTime())))
        $session = Open-Pop3Session -Mailbox $Mailbox
        try { $uid = Get-Pop3UidOf -Session $session -Number 1 }
        finally { Close-Pop3Session -Session $session }
        if ($null -eq $uid) { Fail 'message ids are not reused across a restart' 'UIDL 1 returned no id after the restart.' }
        elseif ($script:SeenIds -contains $uid.Id) {
            Fail 'message ids are not reused across a restart' "The first message after the restart got id '$($uid.Id)', already issued before it. Outlook keys its seen-state on UIDL; a reused id can make it skip new mail."
        }
        else { Pass 'message ids are not reused across a restart' "'$($uid.Id)' is new; $(@($script:SeenIds).Count) id(s) were issued before the restart" }
    }
    catch { Fail 'the sink works after a restart' $_.Exception.Message }
    finally {
        try { [void](Clear-ProbeMailbox -Mailbox $Mailbox -ProbeMailboxes $ProbeMailboxes) } catch { Fail 'the probe mailbox drains after the restart proof' $_.Exception.Message }
    }
}

# =============================================================================================
# MODES.
# =============================================================================================

function Show-Plan {
    $root = $InstallRoot.TrimEnd('\')
    $environment = @(Get-SinkEnvironment -Root $root -Smtp $SmtpPort -Pop3 $Pop3Port -Web $WebPort -Level $LogLevel)
    $pin = '<REQUIRED: -ExpectedSha256, recorded in Testbed/MEDIA.md>'
    if ($ExpectedSha256) { $pin = $ExpectedSha256 }

    Say 'PLAN. Dry run: nothing has been read, written or started.'
    Say ''
    Say "  package           : $PackagePath"
    Say "  pinned sha256     : $pin"
    Say "  install root      : $root   (payload one folder down, as the release zip lays it out)"
    Say "  launcher          : $(Join-Path $root $LauncherName)"
    Say "  scheduled task    : $TaskName   SYSTEM, at startup, no time limit, runs the launcher via cmd.exe"
    Say "  submission        : ${SinkHost}:$SmtpPort   SMTP, no TLS; AUTH optional and never checked"
    Say "  retrieval         : ${SinkHost}:$Pop3Port   POP3, no TLS; any password, including none"
    Say "  web UI            : ${SinkHost}:$WebPort   cannot be switched off, so loopback only"
    Say "  sink log          : $(Join-Path $root $SinkLogName)   (level $LogLevel)"
    Say "  this script's log : $LogPath"
    Say ''
    Say '  Mailboxes. A message for NAME@anything is filed under mailbox "name"; POP3 USER name reads it.'
    foreach ($a in @($AccountAddress)) {
        $box = '<unusable address>'
        try { $box = Get-InbucketMailboxName -Address $a } catch { }
        Say ("    account  {0,-32} -> '{1}'   read-only here, never emptied" -f $a, $box)
    }
    $isolation = Get-IsolationAddress -Address $ProbeAddress
    foreach ($a in @($ProbeAddress, $isolation)) {
        $box = '<unusable address>'
        try { $box = Get-InbucketMailboxName -Address $a } catch { }
        Say ("    probe    {0,-32} -> '{1}'   the only mailboxes -Verify deletes from" -f $a, $box)
    }
    $problem = Get-ProbeMailboxProblem -Probe $ProbeAddress -Accounts $AccountAddress
    if ($problem) { Say "    REFUSED: $problem" }
    Say ''
    Say '  Environment the launcher sets, and why:'
    foreach ($e in $environment) {
        Say ("    {0}={1}" -f $e.Name, $e.Value)
        Say ("        {0}" -f $e.Why)
    }
    Say ''
    Say '  Three places must agree on the ports, or the tier fails in a way nothing diagnoses:'
    Say "    - the live-test settings mailSink block   (submitPort $SmtpPort, retrievePort $Pop3Port)"
    Say "    - the tier profile's POP3Port / SMTPPort  (Testbed/guest/New-TierProfile.ps1)"
    Say '    - this script'
}

function Invoke-Execute {
    Assert-TestbedGuestLocal
    Assert-Elevated
    Assert-OutlookNotRunning
    Open-Log
    Say '== Execute =='

    $rootProblem = Get-InstallRootProblem -Path $InstallRoot
    if ($rootProblem) { throw "REFUSING: -InstallRoot '$InstallRoot' is unusable - $rootProblem." }
    $probeProblem = Get-ProbeMailboxProblem -Probe $ProbeAddress -Accounts $AccountAddress
    if ($probeProblem) { throw "REFUSING: $probeProblem" }

    # -- the package and its pin ------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) {
        throw "REFUSING: no package at $PackagePath. It is STAGED MEDIA: Testbed/host/Get-MailSinkMedia.ps1 stages it on the host and prints the Copy-ToGuest.ps1 line that lands it here. It is never downloaded on the guest."
    }
    $actual = (Get-FileHash -LiteralPath $PackagePath -Algorithm SHA256).Hash
    $hashProblem = Get-Sha256Problem -Expected $ExpectedSha256 -Actual $actual
    if ($hashProblem) { throw "REFUSING: $hashProblem`nA mismatch means 'check the hash', not necessarily 'the file is bad' - confirm which release you staged against Testbed/MEDIA.md." }
    Pass 'the package matches its pin' $actual

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try { $entries = @($zip.Entries | ForEach-Object { $_.FullName }) }
    finally { $zip.Dispose() }
    $appFolder = Get-ZipAppFolder -EntryNames $entries
    if (-not $appFolder) { throw "REFUSING: $ExeName is at the root of the package, not in a versioned folder. That is not the release zip this script and Testbed/MEDIA.md describe." }
    Pass 'the package holds one inbucket.exe, one folder deep' "$appFolder ($($entries.Count) entries)"
    $version = Get-VersionFromFolderName -FolderName $appFolder
    if ($version) { Say "  version from the package layout: $version" }

    $root = $InstallRoot.TrimEnd('\')
    $appDir = Join-Path $root $appFolder
    $exe = Join-Path $appDir $ExeName

    # -- ports, before anything is stopped or written ----------------------------------------
    # An earlier install's own listeners are not "somebody else's": exempt every sink process
    # under the install root, whatever version folder it runs from.
    $running = @(Get-SinkProcess -AnyUnderRoot | ForEach-Object { [int] $_.ProcessId })
    $reserved = Get-ExcludedPortRange
    if (-not $reserved.Readable) { Note 'the reserved-port list could not be read; a bind failure would look like a packaging fault.' }
    $portsOk = $true
    foreach ($pair in @(@('submission', $SmtpPort), @('retrieval', $Pop3Port), @('web UI', $WebPort))) {
        if ($reserved.Readable -and (Test-PortInRange -Port $pair[1] -Ranges $reserved.Ranges)) {
            Fail "$($pair[0]) port $($pair[1]) is usable" 'It is inside a Windows reserved TCP range: nothing listens there and the bind fails anyway. Choose another port and carry it into the tier profile and the settings mailSink block too.'
            $portsOk = $false
            continue
        }
        $ep = Get-ListeningEndpoint -Port $pair[1]
        if ($null -ne $ep -and -not ($running -contains $ep.OwningProcess)) {
            Fail "$($pair[0]) port $($pair[1]) is free" "pid $($ep.OwningProcess) ($($ep.ProcessName)) already listens on $($ep.AllAddresses -join ', ')."
            $portsOk = $false
        }
    }
    if (-not $portsOk) { throw 'REFUSING to install: the ports above cannot be used. Nothing has been written.' }
    Pass 'all three ports are usable' "$SmtpPort, $Pop3Port and $WebPort - free, and outside every reserved range"

    # -- stop whatever is there ---------------------------------------------------------------
    if ($null -ne (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) -or $running.Count -gt 0) {
        Say "  an earlier install is present; stopping it to reconfigure in place"
        if (-not (Stop-Sink)) { throw "The existing sink did not stop within $ListenerWaitSeconds s." }
    }

    # -- unpack ------------------------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null; Say "  created $root" }
    if ($Force -or -not (Test-Path -LiteralPath $exe)) {
        Expand-Archive -LiteralPath $PackagePath -DestinationPath $root -Force
        Get-ChildItem -LiteralPath $appDir -Recurse -File | Unblock-File
        Say "  unpacked into $appDir"
    }
    else { Say "  $exe is already present; leaving the payload alone (-Force replaces it)" }
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "REFUSING: $exe is not there after unpacking." }
    Pass 'the payload is in place' ("$exe, sha256 " + (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash)

    $lua = Join-Path $appDir $LuaScriptName
    if (Test-Path -LiteralPath $lua) { throw "REFUSING: $lua exists. Inbucket runs a Lua script of that name from its working directory on every message, and nothing in this testbed puts one there." }

    $store = Join-Path $root $StoreDirName
    if (-not (Test-Path -LiteralPath $store)) { New-Item -ItemType Directory -Path $store -Force | Out-Null; Say "  created $store" }

    # -- the launcher ------------------------------------------------------------------------
    $layout = Get-SinkLayout
    if ($layout.ExeCount -ne 1) { throw "REFUSING: $($layout.ExeCount) copies of $ExeName sit one folder below $root. Remove the stale one (or -Uninstall -Execute first) so there is no doubt which one runs." }
    $environment = @(Get-SinkEnvironment -Root $root -Smtp $SmtpPort -Pop3 $Pop3Port -Web $WebPort -Level $LogLevel)
    $launcherText = ConvertTo-LauncherText -ExePath $layout.Exe -AppDir $layout.AppDir -SinkLogPath $layout.SinkLog -Environment $environment
    [System.IO.File]::WriteAllText($layout.Launcher, $launcherText, (New-Object System.Text.ASCIIEncoding))
    if ([System.IO.File]::ReadAllText($layout.Launcher) -cne $launcherText) { throw 'The launcher read back differently from what was written. Refusing to continue.' }
    Pass 'the launcher reads back byte-identical' $layout.Launcher

    # -- the scheduled task ------------------------------------------------------------------
    $action = New-ScheduledTaskAction -Execute $CmdExe -Argument ('/d /c "{0}"' -f $layout.Launcher) -WorkingDirectory $layout.AppDir
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -Priority 4
    [void](Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force `
        -Description 'Loopback SMTP + POP3 mail sink (Inbucket) for the OutlookAI live test tier. Written by Testbed/guest/Install-MailSink.ps1.')
    $taskProblems = @(Get-TaskDefinitionProblems -Snapshot (Get-TaskSnapshot) -LauncherPath $layout.Launcher)
    if ($taskProblems.Count -gt 0) { throw ("The task read back wrong: " + ($taskProblems -join '; ')) }
    Pass 'the scheduled task reads back as registered' "$TaskName - SYSTEM, at startup, no time limit"

    if (-not (Start-Sink -ExePath $layout.Exe)) {
        throw "The sink did not bring up all three listeners within $ListenerWaitSeconds s of Start-ScheduledTask. Read $($layout.SinkLog) - Inbucket logs its startup there, including a failed bind."
    }
    Pass 'the sink is running' "started through the task; listening on $SmtpPort, $Pop3Port and $WebPort"

    return (Invoke-Verify -SkipGuards)
}

function Invoke-Verify {
    param([switch] $SkipGuards)
    if (-not $SkipGuards) {
        Assert-TestbedGuestLocal
        Assert-Elevated
        Open-Log
    }
    Say ''
    Say '== Verify =='

    $probeProblem = Get-ProbeMailboxProblem -Probe $ProbeAddress -Accounts $AccountAddress
    if ($probeProblem) { throw "REFUSING: $probeProblem" }
    $probeBox = Get-InbucketMailboxName -Address $ProbeAddress
    $isolationAddress = Get-IsolationAddress -Address $ProbeAddress
    $isolationBox = Get-InbucketMailboxName -Address $isolationAddress
    $probeBoxes = @($probeBox, $isolationBox)

    # -- what is installed -------------------------------------------------------------------
    $layout = Get-SinkLayout
    $snapshot = Get-TaskSnapshot
    if ($layout.ExeCount -eq 0 -and -not $snapshot.Exists) {
        Say "  no $ExeName under $($layout.Root) and no task '$TaskName'."
        return $VerdictAbsent
    }
    if ($layout.ExeCount -ne 1) {
        Fail "exactly one $ExeName is installed" "$($layout.ExeCount) found one folder below $($layout.Root)."
        return $VerdictBroken
    }
    Pass "exactly one $ExeName is installed" $layout.Exe

    try {
        $versionRun = Invoke-NativeLines -FilePath $layout.Exe -Arguments @('-version')
        $versionOut = [string] ((@($versionRun.Lines) + @($versionRun.ErrorLines)) -join ' ')
        $expectedVersion = Get-VersionFromFolderName -FolderName (Split-Path -Leaf $layout.AppDir)
        if ($expectedVersion -and $versionOut -notmatch [regex]::Escape($expectedVersion)) {
            $m = [regex]::Match($versionOut, '\d+\.\d+\.\d+')
            if ($m.Success) { Fail 'the exe reports the version its folder names' "folder says $expectedVersion, the exe says '$versionOut'" }
            else { Note "the exe reports version '$versionOut'; its folder names $expectedVersion" }
        }
        else { Pass 'the exe reports the version its folder names' $versionOut.Trim() }
    }
    catch { Note "'$ExeName -version' could not be run: $($_.Exception.Message)" }

    $taskProblems = @(Get-TaskDefinitionProblems -Snapshot $snapshot -LauncherPath $layout.Launcher)
    if ($taskProblems.Count -eq 0) { Pass 'the scheduled task starts the sink with the guest' "$TaskName - SYSTEM, at startup, no time limit, state $($snapshot.State)" }
    else { foreach ($p in $taskProblems) { Fail 'the scheduled task starts the sink with the guest' $p } }

    $environment = @(Get-SinkEnvironment -Root $layout.Root -Smtp $SmtpPort -Pop3 $Pop3Port -Web $WebPort -Level $LogLevel)
    $wanted = ConvertTo-LauncherText -ExePath $layout.Exe -AppDir $layout.AppDir -SinkLogPath $layout.SinkLog -Environment $environment
    if (-not (Test-Path -LiteralPath $layout.Launcher)) { Fail 'the launcher is what this script writes' "$($layout.Launcher) does not exist." }
    elseif ([System.IO.File]::ReadAllText($layout.Launcher) -cne $wanted) {
        Fail 'the launcher is what this script writes' "$($layout.Launcher) differs from what these parameters render (-LogLevel $LogLevel, ports $SmtpPort/$Pop3Port/$WebPort). Re-run -Execute with the parameters you mean; do not edit the file."
    }
    else { Pass 'the launcher is what this script writes' $layout.Launcher }

    if (Test-Path -LiteralPath (Join-Path $layout.AppDir $LuaScriptName)) {
        Fail 'no Lua script sits in the working directory' "$LuaScriptName would run on every message. Nothing in this testbed puts one there."
    }
    else { Pass 'no Lua script sits in the working directory' 'the Lua extension is inert' }

    # -- the running process and its listeners --------------------------------------------------
    $procs = @(Get-SinkProcess -ExePath $layout.Exe)
    if ($procs.Count -ne 1) {
        Fail 'the sink process is running' "$($procs.Count) $ExeName process(es) from $($layout.Exe). Start-ScheduledTask '$TaskName', or read $($layout.SinkLog)."
        return $VerdictBroken
    }
    $proc = $procs[0]
    $boot = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $sinceBoot = ($proc.CreationDate - $boot).TotalSeconds
    Pass 'the sink process is running' (Format-Invariant 'pid {0}, started {1:yyyy-MM-dd HH:mm:ss}, {2:N0} s after boot' @($proc.ProcessId, $proc.CreationDate, $sinceBoot))
    if ($sinceBoot -gt 300) { Note 'it was started after boot (by -Execute or a restart). A reboot followed by -Verify is what proves it starts WITH the guest.' }

    foreach ($pair in @(@('submission', $SmtpPort), @('retrieval', $Pop3Port), @('web UI', $WebPort))) {
        $ep = Get-ListeningEndpoint -Port $pair[1]
        if ($null -eq $ep) { Fail "the $($pair[0]) listener is up" "nothing listens on port $($pair[1])."; continue }
        if (-not ($ep.Pids -contains [int] $proc.ProcessId) -or $ep.Pids.Count -ne 1) {
            Fail "the $($pair[0]) listener is the sink's" "port $($pair[1]) is held by pid(s) $($ep.Pids -join ', '), not by the sink (pid $($proc.ProcessId))."
            continue
        }
        $wide = @($ep.AllAddresses | Where-Object { $_ -cne $SinkHost })
        if ($wide.Count -gt 0) { Fail "the $($pair[0]) listener is loopback-only" "it is also bound to $($wide -join ', '): an open relay on a test VM. Do NOT answer that with a firewall rule." }
        else { Pass "the $($pair[0]) listener is loopback-only" "${SinkHost}:$($pair[1]), pid $($proc.ProcessId)" }
    }

    try {
        $rules = @(Get-NetFirewallApplicationFilter -ErrorAction Stop |
                Where-Object { $_.Program -and [string]::Equals([string] $_.Program, $layout.Exe, [System.StringComparison]::OrdinalIgnoreCase) })
        if ($rules.Count -eq 0) { Pass 'no firewall rule names the sink' 'none needed: loopback traffic is not filtered' }
        else { Fail 'no firewall rule names the sink' "$($rules.Count) rule(s) name $($layout.Exe). A loopback listener needs none; a rule is evidence somebody worked around a wide bind." }
    }
    catch { Note "the firewall rules could not be read ($($_.Exception.Message)); a rule for the sink would not be reported." }

    # -- the protocol proofs ------------------------------------------------------------------
    Say ''
    Say "Probe mailboxes: '$probeBox' and '$isolationBox' - the only mailboxes this script empties."
    foreach ($box in $probeBoxes) {
        try {
            $residue = Clear-ProbeMailbox -Mailbox $box -ProbeMailboxes $probeBoxes
            if ($residue -gt 0) { Note "'$box' held $residue message(s) from an earlier run that died mid-proof; removed." }
        }
        catch { Fail "the probe mailbox '$box' can be emptied" $_.Exception.Message }
    }

    Test-LoginShapes -Mailbox $probeBox
    Invoke-RoundTripProof -Address $ProbeAddress -Mailbox $probeBox -ProbeMailboxes $probeBoxes
    Invoke-IsolationProof -Address $ProbeAddress -Mailbox $probeBox -OtherMailbox $isolationBox -ProbeMailboxes $probeBoxes
    Invoke-OrdinalStabilityProof -Address $ProbeAddress -Mailbox $probeBox -ProbeMailboxes $probeBoxes
    Invoke-AbortSemanticsProof -Address $ProbeAddress -Mailbox $probeBox -ProbeMailboxes $probeBoxes
    Show-AccountMailboxes
    Invoke-RestartProof -Layout $layout -Address $ProbeAddress -Mailbox $probeBox -ProbeMailboxes $probeBoxes

    if ($script:Failures.Count -gt 0) { return $VerdictBroken }
    return $VerdictReady
}

function Invoke-Uninstall {
    Assert-TestbedGuestLocal
    Assert-Elevated
    Assert-OutlookNotRunning
    Open-Log
    Say '== Uninstall =='

    $rootProblem = Get-InstallRootProblem -Path $InstallRoot
    if ($rootProblem) { throw "REFUSING: -InstallRoot '$InstallRoot' is unusable - $rootProblem." }
    $layout = Get-SinkLayout

    if (Test-Path -LiteralPath $layout.Root) {
        $looksLikeOurs = (Test-Path -LiteralPath $layout.Launcher) -or ($layout.ExeCount -gt 0)
        if (-not $looksLikeOurs) { throw "REFUSING to delete $($layout.Root): it holds neither $LauncherName nor an $ExeName, so it does not look like a sink this script installed." }
    }

    [void](Stop-Sink)
    if ($null -ne (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Say "  unregistered task $TaskName"
    }
    else { Say "  task $TaskName already absent" }

    if (Test-Path -LiteralPath $layout.Root) {
        Remove-Item -LiteralPath $layout.Root -Recurse -Force
        Say "  removed $($layout.Root)"
    }
    else { Say "  $($layout.Root) already absent" }

    if ($null -ne (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) { Fail 'the task is gone' "'$TaskName' still exists." }
    else { Pass 'the task is gone' $TaskName }
    if (Test-Path -LiteralPath $layout.Root) { Fail 'the install root is gone' "$($layout.Root) still exists." }
    else { Pass 'the install root is gone' $layout.Root }
    foreach ($port in @($SmtpPort, $Pop3Port, $WebPort)) {
        $ep = Get-ListeningEndpoint -Port $port
        if ($null -ne $ep -and $ep.ProcessName -eq 'inbucket') { Fail "nothing of the sink listens on $port" "pid $($ep.OwningProcess) still does." }
    }
}

# =============================================================================================
# SELF-TEST. Pure: no socket, no process, no file, no registry.
# =============================================================================================

function Invoke-SelfTest {
    $script:StChecks = 0
    $script:StFailures = @()

    function Show-Value($v) {
        if ($null -eq $v) { return '<null>' }
        if ($v -is [array]) { return ('[' + ((@($v) | ForEach-Object { Show-Value $_ }) -join '|') + ']') }
        return [string] $v
    }
    function Test-Case([string] $What, $Expected, $Actual) {
        $script:StChecks++
        $e = Show-Value $Expected
        $a = Show-Value $Actual
        if ($e -cne $a) { $script:StFailures += "$What : expected $e, got $a" }
    }
    function Test-Throws([string] $What, [scriptblock] $Block, [string] $Fragment) {
        $script:StChecks++
        $threw = $false
        $message = ''
        try { & $Block | Out-Null } catch { $threw = $true; $message = $_.Exception.Message }
        if (-not $threw) { $script:StFailures += "$What : expected a refusal, got none" }
        elseif ($Fragment -and -not $message.Contains($Fragment)) { $script:StFailures += "$What : refused, but the message lacks '$Fragment': $message" }
    }

    # --- mailbox naming, mirroring Inbucket's parseMailboxName --------------------------------
    Test-Case 'a plain address names its local part'            'tier'                  (Get-InbucketMailboxName 'tier@vm.invalid')
    Test-Case 'mailbox names are lowercased'                    'tier'                  (Get-InbucketMailboxName 'Tier@VM.Invalid')
    Test-Case 'a +tag is cut off'                               'tier'                  (Get-InbucketMailboxName 'tier+extra@vm.invalid')
    Test-Case 'the identity account gets its own mailbox'       'identity'              (Get-InbucketMailboxName 'identity@vm.invalid')
    Test-Case 'the probe mailbox'                               'outlookai-sink-probe'  (Get-InbucketMailboxName 'outlookai-sink-probe@vm.invalid')
    Test-Throws 'an address with no @ is refused'               { Get-InbucketMailboxName 'no-at-sign' } "no '@'"
    Test-Throws 'a space in the local part is refused'          { Get-InbucketMailboxName 'bad space@vm.invalid' } 'character'
    Test-Throws 'an empty local part is refused'                { Get-InbucketMailboxName '@vm.invalid' } 'empty local part'
    Test-Throws 'a local part that is only a +tag is refused'   { Get-InbucketMailboxName '+x@vm.invalid' } 'empty mailbox'

    # --- the probe never shares a mailbox with an account ------------------------------------
    Test-Case 'the isolation twin'                              'outlookai-sink-probe-isolation@vm.invalid' (Get-IsolationAddress 'outlookai-sink-probe@vm.invalid')
    Test-Case 'the defaults are disjoint'                       '<null>'                (Get-ProbeMailboxProblem -Probe 'outlookai-sink-probe@vm.invalid' -Accounts @('tier@vm.invalid', 'identity@vm.invalid'))
    $clash = Get-ProbeMailboxProblem -Probe 'tier@elsewhere.invalid' -Accounts @('tier@vm.invalid')
    Test-Case 'a probe in an account mailbox is refused'        $true                   ([bool] $clash -and $clash.Contains("mailbox 'tier'"))
    Test-Case 'a probe that folds into an account after the naming rules is refused' $true ([bool] (Get-ProbeMailboxProblem -Probe 'TIER+probe@vm.invalid' -Accounts @('tier@vm.invalid')))
    Test-Case 'an isolation twin that is an account is refused' $true                   ([bool] (Get-ProbeMailboxProblem -Probe 'x@vm.invalid' -Accounts @('x-isolation@vm.invalid')))
    Test-Throws 'DELE in an account mailbox is refused'         { Assert-ProbeMailboxForDelete -Mailbox 'tier' -ProbeMailboxes @('outlookai-sink-probe', 'outlookai-sink-probe-isolation') } 'REFUSING to DELE'
    Test-Throws 'DELE matches mailbox names ordinally'          { Assert-ProbeMailboxForDelete -Mailbox 'Outlookai-Sink-Probe' -ProbeMailboxes @('outlookai-sink-probe') } 'REFUSING to DELE'
    $allowed = $true
    try { Assert-ProbeMailboxForDelete -Mailbox 'outlookai-sink-probe' -ProbeMailboxes @('outlookai-sink-probe', 'outlookai-sink-probe-isolation') } catch { $allowed = $false }
    Test-Case 'DELE in a probe mailbox is allowed'              $true                   $allowed

    # --- the storage path and cmd safety -------------------------------------------------------
    Test-Case 'the drive colon becomes $'                       'C$\OutlookAI-Sink\store' (ConvertTo-InbucketStoragePath 'C:\OutlookAI-Sink\store')
    Test-Throws 'a comma is refused'                            { ConvertTo-InbucketStoragePath 'C:\a,b' } "','"
    Test-Throws 'a relative path is refused'                    { ConvertTo-InbucketStoragePath 'relative\store' } 'absolute'
    Test-Throws 'a $ in the path is refused'                    { ConvertTo-InbucketStoragePath 'C:\x$y' } "','"
    Test-Case 'an ordinary path is cmd-safe'                    '<null>'                (Get-CmdUnsafeReason 'C:\OutlookAI-Sink\run-sink.cmd')
    Test-Case 'a % is not'                                      $true                   ([bool] (Get-CmdUnsafeReason 'a%b'))
    Test-Case 'an & is not'                                     $true                   ([bool] (Get-CmdUnsafeReason 'a&b'))
    Test-Case 'a double quote is not'                           $true                   ([bool] (Get-CmdUnsafeReason 'a"b'))
    Test-Case 'a line break is not'                             $true                   ([bool] (Get-CmdUnsafeReason "a`r`nb"))

    # --- install-root refusals, because -Uninstall deletes it recursively ------------------------
    Test-Case 'the default install root is fine'                '<null>'                (Get-InstallRootProblem 'C:\OutlookAI-Sink')
    Test-Case 'a drive root is refused'                         $true                   ([bool] (Get-InstallRootProblem 'C:\'))
    Test-Case 'the Windows directory is refused'                $true                   ([bool] (Get-InstallRootProblem 'C:\Windows'))
    Test-Case 'a directory inside Program Files is refused'     $true                   ([bool] (Get-InstallRootProblem 'C:\Program Files\Sink'))
    Test-Case 'the Users directory is refused'                  $true                   ([bool] (Get-InstallRootProblem 'C:\Users'))
    Test-Case 'a relative path is refused'                      $true                   ([bool] (Get-InstallRootProblem 'OutlookAI-Sink'))
    Test-Case 'a UNC path is refused'                           $true                   ([bool] (Get-InstallRootProblem '\\server\share\sink'))
    Test-Case 'whitespace is refused'                           $true                   ([bool] (Get-InstallRootProblem 'C:\Outlook Sink'))

    # --- the environment and the launcher ------------------------------------------------------
    $envList = @(Get-SinkEnvironment -Root 'C:\OutlookAI-Sink' -Smtp 25 -Pop3 110 -Web 9000 -Level 'info')
    $byName = @{}
    foreach ($e in $envList) { $byName[$e.Name] = $e.Value }
    Test-Case 'the environment, in order' 'INBUCKET_LOGLEVEL,INBUCKET_MAILBOXNAMING,INBUCKET_SMTP_ADDR,INBUCKET_POP3_ADDR,INBUCKET_WEB_ADDR,INBUCKET_SMTP_TLSENABLED,INBUCKET_POP3_TLSENABLED,INBUCKET_STORAGE_TYPE,INBUCKET_STORAGE_PARAMS,INBUCKET_STORAGE_RETENTIONPERIOD,INBUCKET_STORAGE_MAILBOXMSGCAP' (($envList | ForEach-Object { $_.Name }) -join ',')
    Test-Case 'every setting says why'                          $envList.Count          (@($envList | Where-Object { $_.Why -and $_.Why.Length -gt 20 }).Count)
    Test-Case 'SMTP binds loopback'                             '127.0.0.1:25'          $byName['INBUCKET_SMTP_ADDR']
    Test-Case 'POP3 binds loopback'                             '127.0.0.1:110'         $byName['INBUCKET_POP3_ADDR']
    Test-Case 'the web UI binds loopback'                       '127.0.0.1:9000'        $byName['INBUCKET_WEB_ADDR']
    Test-Case 'mailboxes are named by local part'               'local'                 $byName['INBUCKET_MAILBOXNAMING']
    Test-Case 'the store is on disk'                            'file'                  $byName['INBUCKET_STORAGE_TYPE']
    Test-Case 'the store path is written in the $ syntax'       'path:C$\OutlookAI-Sink\store' $byName['INBUCKET_STORAGE_PARAMS']
    Test-Case 'TLS is off on both protocols'                    'false|false'           ($byName['INBUCKET_SMTP_TLSENABLED'] + '|' + $byName['INBUCKET_POP3_TLSENABLED'])

    $exe = 'C:\OutlookAI-Sink\inbucket_3.1.1_windows_amd64\inbucket.exe'
    $appDir = 'C:\OutlookAI-Sink\inbucket_3.1.1_windows_amd64'
    $launcher = ConvertTo-LauncherText -ExePath $exe -AppDir $appDir -SinkLogPath 'C:\OutlookAI-Sink\inbucket.log' -Environment $envList
    $lfCount = ([regex]::Matches($launcher, "`n")).Count
    $crlfCount = ([regex]::Matches($launcher, "`r`n")).Count
    Test-Case 'the launcher ends every line with CRLF'          $lfCount                $crlfCount
    Test-Case 'the launcher is ASCII'                           $true                   ($launcher -cmatch '^[\x00-\x7F]*$')
    Test-Case 'the launcher sets each variable exactly once'    $envList.Count          (@($envList | Where-Object { ([regex]::Matches($launcher, [regex]::Escape(('set "{0}={1}"' -f $_.Name, $_.Value)))).Count -eq 1 }).Count)
    $launcherLines = @($launcher.TrimEnd("`r", "`n") -split "`r`n")
    Test-Case 'the launcher runs the exe last, logging to a file' ('"' + $exe + '" -logfile "C:\OutlookAI-Sink\inbucket.log"') $launcherLines[-1]
    Test-Case 'and changes into its folder first'               ('cd /d "' + $appDir + '"') $launcherLines[-2]
    $other = ConvertTo-LauncherText -ExePath $exe -AppDir $appDir -SinkLogPath 'C:\OutlookAI-Sink\inbucket.log' -Environment (Get-SinkEnvironment -Root 'C:\OutlookAI-Sink' -Smtp 25 -Pop3 110 -Web 9000 -Level 'debug')
    Test-Case 'a different log level renders a different launcher, so drift is visible' $true ($other -cne $launcher)
    Test-Throws 'an exe path cmd would interpret is refused'    { ConvertTo-LauncherText -ExePath 'C:\a%b\inbucket.exe' -AppDir 'C:\a' -SinkLogPath 'C:\a\l.log' -Environment $envList } 'Refusing to write a launcher'
    Test-Throws 'a setting cmd would interpret is refused'      { ConvertTo-LauncherText -ExePath $exe -AppDir $appDir -SinkLogPath 'C:\l.log' -Environment @([pscustomobject]@{ Name = 'X'; Value = 'a&b' }) } 'X is unusable'

    # --- the package layout ------------------------------------------------------------------
    $entries = @('inbucket_3.1.1_windows_amd64/LICENSE', 'inbucket_3.1.1_windows_amd64/ui/dist/index.html', 'inbucket_3.1.1_windows_amd64/inbucket.exe')
    Test-Case 'the release zip holds the exe one folder deep'   'inbucket_3.1.1_windows_amd64' (Get-ZipAppFolder $entries)
    Test-Case 'an exe at the root is reported as such'          ''                      (Get-ZipAppFolder @('inbucket.exe', 'LICENSE'))
    Test-Case 'backslash separators are understood'             'inbucket_3.1.1_windows_amd64' (Get-ZipAppFolder @('inbucket_3.1.1_windows_amd64\inbucket.exe'))
    Test-Case 'the exe name is matched case-insensitively'      'x'                     (Get-ZipAppFolder @('x/Inbucket.EXE'))
    Test-Throws 'a package with no exe is refused'              { Get-ZipAppFolder @('README.md') } 'holds no'
    Test-Throws 'a package with two exes is refused'            { Get-ZipAppFolder @('a/inbucket.exe', 'b/inbucket.exe') } 'Refusing to guess'
    Test-Throws 'an exe buried deeper is refused'               { Get-ZipAppFolder @('a/b/inbucket.exe') } 'folders deep'
    Test-Case 'the version is read off the folder name'         '3.1.1'                 (Get-VersionFromFolderName 'inbucket_3.1.1_windows_amd64')
    Test-Case 'and nothing is invented for another name'        '<null>'                (Get-VersionFromFolderName 'something-else')

    # --- the hash --------------------------------------------------------------------------------
    $good = '232fb49c92f88505be1feceb4be90b70ca59bb7853216dee7c8b2814c85235d0'
    Test-Case 'a matching hash in either case passes'           '<null>'                (Get-Sha256Problem -Expected $good -Actual $good.ToUpperInvariant())
    Test-Case 'surrounding whitespace is ignored'               '<null>'                (Get-Sha256Problem -Expected (" $good ") -Actual $good)
    Test-Case 'a missing hash is refused, not defaulted'        $true                   ((Get-Sha256Problem -Expected '' -Actual $good).Contains('mandatory'))
    Test-Case 'something that is not a SHA-256 is refused'      $true                   ((Get-Sha256Problem -Expected 'abc' -Actual $good).Contains('not a SHA-256'))
    $mismatch = Get-Sha256Problem -Expected $good -Actual ('0' * 64)
    Test-Case 'a mismatch names both values'                    $true                   ($mismatch.Contains($good.ToUpperInvariant()) -and $mismatch.Contains('0' * 64))

    # --- reserved port ranges ----------------------------------------------------------------------
    $netsh = @('', 'Protocol tcp Port Exclusion Ranges', '', 'Start Port    End Port', '----------    --------', '      5357        5357', '     50000       50059     *', '     49709       49808', '', '* - Administered port exclusions.')
    $ranges = @(Get-ExcludedPortRangeFromText -Lines $netsh)
    Test-Case 'three reserved ranges are read, the administered one included' 3 $ranges.Count
    Test-Case 'with their bounds'                               '5357-5357|50000-50059|49709-49808' (($ranges | ForEach-Object { '{0}-{1}' -f $_.Start, $_.End }) -join '|')
    Test-Case 'a port inside one is caught'                     $true                   (Test-PortInRange -Port 50010 -Ranges $ranges)
    Test-Case 'the SMTP port is not'                            $false                  (Test-PortInRange -Port 25 -Ranges $ranges)
    Test-Case 'an empty table has no ranges'                    0                       @(Get-ExcludedPortRangeFromText -Lines @()).Count

    # --- SMTP replies and dot-stuffing -------------------------------------------------------------
    Test-Case 'a one-line reply'                                220                     (ConvertFrom-SmtpReplyLines @('220 inbucket Inbucket SMTP ready')).Code
    $multi = ConvertFrom-SmtpReplyLines @('250-inbucket Hello', '250-8BITMIME', '250 SIZE 10240000')
    Test-Case 'a multi-line reply takes its code from the last line' '250|3'           ('{0}|{1}' -f $multi.Code, $multi.LineCount)
    Test-Case 'garbage has no code'                             0                       (ConvertFrom-SmtpReplyLines @('xyz')).Code
    $plain = @('.', '.a', '..b', 'c', '')
    $stuffed = @(ConvertTo-DotStuffedLines $plain)
    Test-Case 'every leading period gains one on the way out'   @('..', '..a', '...b', 'c', '') $stuffed
    Test-Case 'a single line comes back as one line'            @('..x')               @(ConvertTo-DotStuffedLines @('.x'))
    $decoded = ConvertFrom-DotStuffedLines (@($stuffed) + @('.'))
    Test-Case 'and loses exactly one on the way back'           $plain                  $decoded.Lines
    Test-Case 'the terminator is recognised'                    $true                   $decoded.Terminated
    Test-Case 'a response that just stops is not terminated'    $false                  (ConvertFrom-DotStuffedLines @('a', 'b')).Terminated

    # --- POP3 status lines -----------------------------------------------------------------------
    $stat = ConvertFrom-Pop3Stat '+OK 2 320'
    Test-Case 'STAT is read'                                    '2|320'                 ('{0}|{1}' -f $stat.Count, $stat.Octets)
    Test-Case 'an empty mailbox is read'                        0                       (ConvertFrom-Pop3Stat '+OK 0 0').Count
    Test-Case 'an -ERR is not a STAT'                           '<null>'                (ConvertFrom-Pop3Stat '-ERR STAT command must have no arguments')
    $uid = ConvertFrom-Pop3Uidl '+OK 1 20260924T150405-0000'
    Test-Case 'a UIDL line is read'                             '1|20260924T150405-0000' ('{0}|{1}' -f $uid.Number, $uid.Id)
    Test-Case 'a deleted message has no UIDL'                   '<null>'                (ConvertFrom-Pop3Uidl '-ERR You deleted message 1')

    # --- the probe message, built under the guests' own formats culture ---------------------------
    $saved = [System.Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object System.Globalization.CultureInfo('nl-NL')
        $probe = New-ProbeMessage -Address 'outlookai-sink-probe@vm.invalid' -Marker 'abc123def456' -DateUtc (New-Object DateTime(2026, 9, 24, 13, 5, 9, [DateTimeKind]::Utc))
    }
    finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $saved }
    Test-Case 'the Date header is RFC 5322 whatever the culture' 'Date: Thu, 24 Sep 2026 13:05:09 +0000' $probe.DateHeader
    $saved = [System.Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object System.Globalization.CultureInfo('nl-NL')
        $printed = Format-Invariant '{0:N1} s, {1:N0} s after boot' @(1.5, 1234)
    }
    finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $saved }
    Test-Case 'printed numbers ignore the guests'' nl-NL formats' '1.5 s, 1,234 s after boot' $printed
    Test-Case 'the subject starts with the live-tier tag'       $true                   $probe.Subject.StartsWith('[OutlookAI-McpTest] ', [System.StringComparison]::Ordinal)
    Test-Case 'and never carries the corpus tag'                $false                  $probe.Subject.Contains('OutlookAI-Corpus')
    Test-Case 'the body has one bare period'                    1                       @($probe.Lines | Where-Object { $_ -ceq '.' }).Count
    Test-Case 'an intact retrieval has no problems'             0                       @(Get-ProbeBodyProblems -Message $probe -Retrieved $probe.Lines).Count
    $cut = @($probe.Lines[0..10])
    Test-Case 'a body cut at the bare period is caught'         $true                   (@(Get-ProbeBodyProblems -Message $probe -Retrieved $cut).Count -ge 2)
    $unstuffedWrong = @($probe.Lines | ForEach-Object { if ($_ -ceq '.leading period must survive') { '..leading period must survive' } else { $_ } })
    Test-Case 'a path that did not unstuff is caught'           $true                   (((Get-ProbeBodyProblems -Message $probe -Retrieved $unstuffedWrong) -join ' ').Contains('dot-stuffing'))
    $noAttachment = @($probe.Lines | Where-Object { $_ -cne $probe.AttachmentB64 })
    Test-Case 'a lost attachment is caught'                     $true                   (((Get-ProbeBodyProblems -Message $probe -Retrieved $noAttachment) -join ' ').Contains('attachment'))
    $noSubject = @($probe.Lines | Where-Object { -not $_.StartsWith('Subject:') })
    Test-Case 'a lost subject is caught'                        $true                   (((Get-ProbeBodyProblems -Message $probe -Retrieved $noSubject) -join ' ').Contains('subject'))

    # --- the scheduled-task audit ------------------------------------------------------------------
    $launcherPath = 'C:\OutlookAI-Sink\run-sink.cmd'
    $goodTask = [pscustomobject]@{
        Exists = $true; State = 'Running'; Enabled = $true; TriggerKinds = @('MSFT_TaskBootTrigger'); UserId = 'SYSTEM'
        ActionCount = 1; Execute = 'C:\WINDOWS\System32\cmd.exe'; Arguments = '/d /c "C:\OutlookAI-Sink\run-sink.cmd"'
        ExecutionTimeLimit = 'PT0S'; MultipleInstances = 'IgnoreNew'
    }
    function Copy-Task($t, [hashtable] $changes) {
        $c = $t.PSObject.Copy()
        foreach ($k in $changes.Keys) { $c.$k = $changes[$k] }
        return $c
    }
    Test-Case 'a correct task has no problems'                  0                       @(Get-TaskDefinitionProblems -Snapshot $goodTask -LauncherPath $launcherPath).Count
    Test-Case 'the SID spelling of SYSTEM is accepted'          0                       @(Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ UserId = 'S-1-5-18' }) -LauncherPath $launcherPath).Count
    Test-Case 'a missing task is reported'                      'the scheduled task does not exist' ((Get-TaskDefinitionProblems -Snapshot ([pscustomobject]@{ Exists = $false }) -LauncherPath $launcherPath) -join '')
    Test-Case 'the three-day default time limit is caught'      $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ ExecutionTimeLimit = 'PT72H' }) -LauncherPath $launcherPath) -join ' ').Contains('three days'))
    Test-Case 'a task for a user is caught'                     $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ UserId = 'vmadmin' }) -LauncherPath $launcherPath) -join ' ').Contains('not SYSTEM'))
    Test-Case 'a logon trigger is caught'                       $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ TriggerKinds = @('MSFT_TaskLogonTrigger') }) -LauncherPath $launcherPath) -join ' ').Contains('at startup'))
    Test-Case 'a second trigger is caught'                      $true                   (@(Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ TriggerKinds = @('MSFT_TaskBootTrigger', 'MSFT_TaskLogonTrigger') }) -LauncherPath $launcherPath).Count -ge 1)
    Test-Case 'a task running another launcher is caught'       $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ Arguments = '/d /c "C:\elsewhere\run.cmd"' }) -LauncherPath $launcherPath) -join ' ').Contains('do not name the launcher'))
    Test-Case 'a disabled task is caught'                       $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ Enabled = $false }) -LauncherPath $launcherPath) -join ' ').Contains('disabled'))
    Test-Case 'two actions are caught'                          $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ ActionCount = 2 }) -LauncherPath $launcherPath) -join ' ').Contains('exactly one action'))
    Test-Case 'parallel instances are caught'                   $true                   (((Get-TaskDefinitionProblems -Snapshot (Copy-Task $goodTask @{ MultipleInstances = 'Parallel' }) -LauncherPath $launcherPath) -join ' ').Contains('IgnoreNew'))

    Write-Host ''
    Write-Host ("SELF-TEST: {0} assertion(s), {1} failure(s)." -f $script:StChecks, $script:StFailures.Count)
    foreach ($f in $script:StFailures) { Write-Host "  FAIL $f" }
    Write-Host ''
    Write-Host 'What -SelfTest cannot prove, and only a guest run can:'
    Write-Host '  - that Inbucket 3.1.1 behaves as its source reads: an empty PASS accepted, DELE at QUIT,'
    Write-Host '    numbers fixed after DELE, TOP working, mailboxes isolated, ids never reused (-Verify);'
    Write-Host '  - that the scheduled task starts it at boot as SYSTEM with no window (reboot, then -Verify);'
    Write-Host '  - whether Outlook, holding no stored password, logs in or prompts (Docs/live-tier-on-the-vm.md 2.7).'
    return $script:StFailures.Count
}

# =============================================================================================
# ENTRY POINT.
# =============================================================================================

if ($SelfTest) {
    # The last thing Invoke-SelfTest emits is its failure count; taking [-1] keeps a stray
    # emission elsewhere from turning the count into a list that compares wrongly.
    $failed = [int] (@(Invoke-SelfTest)[-1])
    if ($failed -gt 0) { exit 1 }
    exit 0
}

if ($Verify -and $Execute) { throw 'Pass -Execute or -Verify, not both: -Execute runs the whole verification when it finishes.' }
if ($Uninstall -and $Verify) { throw 'Pass -Uninstall or -Verify, not both.' }

$verdict = $null
if ($Uninstall) {
    if (-not $Execute) {
        Say 'PLAN: stop the sink, unregister its task, delete the install root. Dry run: nothing done.'
        Say "  task         : $TaskName"
        Say "  install root : $InstallRoot"
        Say 'Re-run with -Uninstall -Execute.'
        exit 0
    }
    Invoke-Uninstall
}
elseif ($Verify) {
    $verdict = Invoke-Verify
}
elseif ($Execute) {
    $verdict = Invoke-Execute
}
else {
    Show-Plan
    Say ''
    Say 'Dry run. Re-run with -ExpectedSha256 <hash from Testbed/MEDIA.md> -Execute.'
    exit 0
}

Say ''
if ($script:Failures.Count -gt 0) {
    Say ("{0} check(s) FAILED:" -f $script:Failures.Count)
    foreach ($f in $script:Failures) { Say "  - $f" }
}
if ($verdict) {
    Say "VERDICT: $verdict"
    switch ($verdict) {
        $VerdictReady  { Say '  The sink answers SMTP and POP3 on loopback and hands back exactly what it was given. Whether Outlook logs in to it without a stored password is a separate question - Docs/live-tier-on-the-vm.md section 2.7.' }
        $VerdictAbsent { Say '  Nothing is installed. Run with -ExpectedSha256 <hash> -Execute.' }
        $VerdictBroken { Say "  Something that should work does not. The FAIL lines say which; the sink's own log is $(Join-Path $InstallRoot.TrimEnd('\') $SinkLogName)." }
    }
}
if ($script:LogReady) { Write-Host ''; Write-Host "Log: $LogPath" }

if ($script:Failures.Count -gt 0) { exit 1 }
if ($verdict -eq $VerdictAbsent) { exit 3 }
exit 0
