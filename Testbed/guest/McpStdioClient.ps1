#Requires -Version 5.1
<#
.SYNOPSIS
    Dot-sourced by the guest scripts that drive the shipped MCP server over stdio: start it, send a request and read
    its reply, call a tool, stop it. Defines functions; runs nothing.

.DESCRIPTION
    One copy (AGENTS.md, Shared code, Q154) of the client Testbed/guest/Invoke-GuestMeasure.ps1 proved in 2026-08 and
    Set-AccountSignature.ps1 copied: newline-delimited JSON-RPC 2.0 over the server's stdin and stdout, UTF-8 with no
    byte-order mark, each reply matched to its request by id and every other line skipped. Set-AccountSignature.ps1,
    Set-SignatureSettings.ps1, Invoke-GuestMeasure.ps1 and Measure-PinWindows.ps1 use it; the last two had clients of
    their own until 2026-11-18, and what they did that this did not, it does now. Set-SignatureSettings.ps1's
    -SelfTest holds it: the pure parts, and a stand-in server it talks to.

    The server's standard error is read to its end in the background from the start: it is where the server logs,
    and a pipe nobody reads fills and stops the server mid-reply. A read of its standard output that is still waiting
    when a reply's time runs out is kept and taken up by the next read, so no line is lost and the stream is never
    read twice at once - a second read beside a waiting one throws.

    A session is a hashtable: Process, Stdin, Stderr (that read), Transcript - a file every line sent (">>> ") and
    read ("<<< ") is appended to, or $null - TimeoutSeconds for each reply, NextId, PendingRead (the read still
    waiting, or $null) and LastReplyLine (the line the last reply came in, as the server wrote it). A caller may add
    keys of its own, as Measure-PinWindows.ps1 adds Name.

    Windows PowerShell 5.1: no ternary, no `??`.
#>

# The line a request or a notification is sent as. Pure.
function ConvertTo-McpMessage {
    param([string] $Method, $Params, $Id, [switch] $Notification)
    $message = [ordered]@{ jsonrpc = '2.0'; method = $Method }
    if (-not $Notification) { $message['id'] = $Id }
    if ($null -ne $Params) { $message['params'] = $Params }
    return ($message | ConvertTo-Json -Depth 20 -Compress)
}

# The reply a line of the server's output is when it answers request $Id - a response, which carries a result or an
# error - or $null for any other line: another request's reply, a notification, a blank line, text that is not JSON.
# Pure.
function Get-McpReplyFor {
    param([string] $Line, $Id)
    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
    $parsed = $null
    try { $parsed = $Line | ConvertFrom-Json } catch { return $null }
    if ($null -eq $parsed) { return $null }
    $names = @($parsed.PSObject.Properties.Name)
    if (-not ($names -contains 'id')) { return $null }
    if ("$($parsed.id)" -ne "$Id") { return $null }
    if (-not (($names -contains 'result') -or ($names -contains 'error'))) { return $null }
    return $parsed
}

# A tool call's answer, parsed from its text content - or a throw carrying what the server said: a JSON-RPC error,
# or a tool's own refusal (isError, and an 'error' object in the answer). Pure.
function Get-McpToolResult {
    param($Reply, [string] $Tool)
    if ($Reply.PSObject.Properties.Name -contains 'error') { throw "$Tool failed: $($Reply.error | ConvertTo-Json -Depth 10 -Compress)" }
    $text = [string]@($Reply.result.content)[0].text
    $answer = $null
    try { $answer = $text | ConvertFrom-Json } catch { throw "$Tool answered with text that is not JSON: $text" }
    $refused = ($Reply.result.PSObject.Properties.Name -contains 'isError') -and ($Reply.result.isError -eq $true)
    if ($refused -or ($null -ne $answer -and ($answer.PSObject.Properties.Name -contains 'error'))) { throw "$Tool refused: $text" }
    return $answer
}

# Starts the server and initializes the session - or, with -NoInitialize, only starts it, for a caller that times or
# reports the initialize itself (Measure-PinWindows.ps1). -ArgumentList is the server's command line: none for the
# shipped server; Set-SignatureSettings.ps1's self-test starts a stand-in with one. A server that cannot be
# initialized is stopped before the error is thrown, and the error carries what it wrote to standard error.
function Start-McpStdioServer {
    param(
        [Parameter(Mandatory = $true)] [string] $ServerExe,
        [string] $ArgumentList,
        [string] $ClientName = 'outlookai-guest',
        [int] $TimeoutSeconds = 300,
        [string] $Transcript,
        [switch] $NoInitialize
    )
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $ServerExe
    if ($ArgumentList) { $psi.Arguments = $ArgumentList }
    $psi.WorkingDirectory = Split-Path $ServerExe -Parent
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
    $psi.StandardErrorEncoding = New-Object Text.UTF8Encoding($false)
    $process = [Diagnostics.Process]::Start($psi)
    $stdin = New-Object IO.StreamWriter($process.StandardInput.BaseStream, (New-Object Text.UTF8Encoding($false)))
    $stdin.AutoFlush = $true
    $session = @{
        Process = $process; Stdin = $stdin; Stderr = $process.StandardError.ReadToEndAsync(); Transcript = $Transcript
        TimeoutSeconds = $TimeoutSeconds; NextId = 1; PendingRead = $null; LastReplyLine = $null
    }
    if ($NoInitialize) { return $session }
    try {
        [void](Invoke-McpRequest -Session $session -Method 'initialize' -Params @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = $ClientName; version = '1.0' } })
        Send-McpMessage -Session $session -Method 'notifications/initialized' -Notification
    }
    catch {
        $failure = $_.Exception.Message
        $stderr = Stop-McpStdioServer -Session $session
        if ($stderr) { $failure = "$failure`nthe server's standard error:`n$($stderr.TrimEnd())" }
        throw $failure
    }
    return $session
}

function Send-McpMessage {
    param([hashtable] $Session, [string] $Method, $Params, $Id, [switch] $Notification)
    $json = ConvertTo-McpMessage -Method $Method -Params $Params -Id $Id -Notification:$Notification
    if ($Session.Transcript) { Add-Content -Path $Session.Transcript -Value ">>> $json" }
    $Session.Stdin.WriteLine($json)
}

# Sends a request and returns its id - the session's next one unless -Id names it - for a caller that reads the
# reply later with Read-McpReply, as Measure-PinWindows.ps1 sends to three sessions before it reads from any.
function Send-McpRequest {
    param([hashtable] $Session, [string] $Method, $Params, $Id)
    if ($null -eq $Id) {
        $Id = $Session.NextId
        $Session.NextId = $Session.NextId + 1
    }
    Send-McpMessage -Session $Session -Method $Method -Params $Params -Id $Id
    return $Id
}

# The reply to request $Id, read within the session's timeout - or -TimeoutSeconds - and kept in LastReplyLine as the
# line it came in. Throws when the server ends or the time runs out; with -NullWhenNone it returns $null then, for a
# caller that reports a silence rather than stops on one. A read still waiting when the time runs out is kept in
# PendingRead for the next call.
function Read-McpReply {
    param([hashtable] $Session, $Id, [int] $TimeoutSeconds = 0, [switch] $NullWhenNone)
    $process = $Session.Process
    $seconds = $Session.TimeoutSeconds
    if ($TimeoutSeconds -gt 0) { $seconds = $TimeoutSeconds }
    $deadline = (Get-Date).AddSeconds($seconds)
    $why = "no answer to request $Id within $seconds s"
    while ((Get-Date) -lt $deadline) {
        if ($null -eq $Session.PendingRead) {
            if ($process.HasExited -and $process.StandardOutput.EndOfStream) { $why = "the server exited $($process.ExitCode) before answering request $Id"; break }
            $Session.PendingRead = $process.StandardOutput.ReadLineAsync()
        }
        $remaining = [int]([Math]::Max(1000, ($deadline - (Get-Date)).TotalMilliseconds))
        if (-not $Session.PendingRead.Wait($remaining)) { break }
        $line = $Session.PendingRead.Result
        $Session.PendingRead = $null
        if ($null -eq $line) { $why = "the server closed its output before answering request $Id"; break }
        if ($line.Trim() -eq '') { continue }
        if ($Session.Transcript) { Add-Content -Path $Session.Transcript -Value "<<< $line" }
        $reply = Get-McpReplyFor -Line $line -Id $Id
        if ($null -ne $reply) {
            $Session.LastReplyLine = $line
            return $reply
        }
    }
    if ($NullWhenNone) { return $null }
    throw $why
}

# One request and its reply. Without -Id the session numbers it.
function Invoke-McpRequest {
    param([hashtable] $Session, [string] $Method, $Params, $Id)
    $Id = Send-McpRequest -Session $Session -Method $Method -Params $Params -Id $Id
    return (Read-McpReply -Session $Session -Id $Id)
}

# One tool call, its answer parsed (Get-McpToolResult).
function Invoke-McpTool {
    param([hashtable] $Session, [string] $Tool, [hashtable] $Arguments = @{}, $Id)
    $reply = Invoke-McpRequest -Session $Session -Method 'tools/call' -Params @{ name = $Tool; arguments = $Arguments } -Id $Id
    return (Get-McpToolResult -Reply $reply -Tool $Tool)
}

# How a client ends a session normally: it closes the server's input, and the server exits when it sees that.
# Measure-PinWindows.ps1 closes three sessions' at once this way before it waits for any.
function Close-McpStdioInput {
    param([hashtable] $Session)
    if ($null -ne $Session -and $null -ne $Session.Stdin) { try { $Session.Stdin.Dispose() } catch { } }
}

# Ends the session: its input closed, and the server given -WaitMilliseconds to exit rather than killed - it owns a
# COM host attached to Outlook. -KillIfStillRunning kills a server that has not exited by then, as
# Invoke-GuestMeasure.ps1 did when it was found. Returns what the server wrote to standard error - waiting up to 5 s
# for the end of it once the server has exited - or ''.
function Stop-McpStdioServer {
    param([hashtable] $Session, [int] $WaitMilliseconds = 30000, [switch] $KillIfStillRunning)
    if ($null -eq $Session) { return '' }
    Close-McpStdioInput -Session $Session
    $process = $Session.Process
    if ($null -ne $process -and -not $process.HasExited) {
        if (-not $process.WaitForExit($WaitMilliseconds) -and $KillIfStillRunning) {
            try { $process.Kill() } catch { }
            [void]$process.WaitForExit(5000)
        }
    }
    $stderr = ''
    if ($null -ne $Session.Stderr) {
        if ($null -ne $process -and $process.HasExited -and -not $Session.Stderr.IsCompleted) { [void]$Session.Stderr.Wait(5000) }
        if ($Session.Stderr.IsCompleted) { try { $stderr = [string]$Session.Stderr.Result } catch { } }
    }
    return $stderr
}
