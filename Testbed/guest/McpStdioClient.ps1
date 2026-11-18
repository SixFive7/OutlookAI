#Requires -Version 5.1
<#
.SYNOPSIS
    Dot-sourced by the guest scripts that drive the shipped MCP server over stdio: start it, send a request and read
    its reply, call a tool, stop it. Defines functions; runs nothing.

.DESCRIPTION
    One copy (AGENTS.md, Shared code, Q154) of the client Testbed/guest/Invoke-GuestMeasure.ps1 proved in 2026-08 and
    Set-AccountSignature.ps1 copied: newline-delimited JSON-RPC 2.0 over the server's stdin and stdout, UTF-8 with no
    byte-order mark, each reply matched to its request by id and every other line skipped. Set-AccountSignature.ps1
    and Set-SignatureSettings.ps1 use it; Invoke-GuestMeasure.ps1 keeps its own copy because it is a recovered
    artefact, kept as it was found (its header). Set-SignatureSettings.ps1's -SelfTest holds the pure parts here.

    The server's standard error is read to its end in the background from the start: it is where the server logs,
    and a pipe nobody reads fills and stops the server mid-reply.

    A session is a hashtable: Process, Stdin, Stderr (that read), Transcript - a file every line sent (">>> ") and
    read ("<<< ") is appended to, or $null - TimeoutSeconds for each reply, and NextId.

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

# The reply a line of the server's output is when it answers request $Id, or $null for any other line - another
# request's reply, a notification, a blank line, text that is not JSON. Pure.
function Get-McpReplyFor {
    param([string] $Line, $Id)
    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
    $parsed = $null
    try { $parsed = $Line | ConvertFrom-Json } catch { return $null }
    if ($null -eq $parsed -or -not ($parsed.PSObject.Properties.Name -contains 'id')) { return $null }
    if ("$($parsed.id)" -ne "$Id") { return $null }
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

# Starts the server and initializes the session.
function Start-McpStdioServer {
    param([Parameter(Mandatory = $true)] [string] $ServerExe, [string] $ClientName = 'outlookai-guest', [int] $TimeoutSeconds = 300, [string] $Transcript)
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $ServerExe
    $psi.WorkingDirectory = Split-Path $ServerExe -Parent
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
    $process = [Diagnostics.Process]::Start($psi)
    $stdin = New-Object IO.StreamWriter($process.StandardInput.BaseStream, (New-Object Text.UTF8Encoding($false)))
    $stdin.AutoFlush = $true
    $session = @{ Process = $process; Stdin = $stdin; Stderr = $process.StandardError.ReadToEndAsync(); Transcript = $Transcript; TimeoutSeconds = $TimeoutSeconds; NextId = 1 }
    [void](Invoke-McpRequest -Session $session -Method 'initialize' -Params @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = $ClientName; version = '1.0' } })
    Send-McpMessage -Session $session -Method 'notifications/initialized' -Notification
    return $session
}

function Send-McpMessage {
    param([hashtable] $Session, [string] $Method, $Params, $Id, [switch] $Notification)
    $json = ConvertTo-McpMessage -Method $Method -Params $Params -Id $Id -Notification:$Notification
    if ($Session.Transcript) { Add-Content -Path $Session.Transcript -Value ">>> $json" }
    $Session.Stdin.WriteLine($json)
}

# The reply to request $Id, read within the session's timeout; throws when the server ends or the time runs out.
function Read-McpReply {
    param([hashtable] $Session, $Id)
    $process = $Session.Process
    $deadline = (Get-Date).AddSeconds($Session.TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($process.HasExited -and $process.StandardOutput.EndOfStream) { throw "the server exited $($process.ExitCode) before answering request $Id" }
        $remaining = [int]([Math]::Max(1000, ($deadline - (Get-Date)).TotalMilliseconds))
        $task = $process.StandardOutput.ReadLineAsync()
        if (-not $task.Wait($remaining)) { throw "no answer to request $Id within $($Session.TimeoutSeconds) s" }
        $line = $task.Result
        if ($null -eq $line) { throw "the server closed its output before answering request $Id" }
        if ($line.Trim() -eq '') { continue }
        if ($Session.Transcript) { Add-Content -Path $Session.Transcript -Value "<<< $line" }
        $reply = Get-McpReplyFor -Line $line -Id $Id
        if ($null -ne $reply) { return $reply }
    }
    throw "no answer to request $Id within $($Session.TimeoutSeconds) s"
}

# One request and its reply. Without -Id the session numbers it.
function Invoke-McpRequest {
    param([hashtable] $Session, [string] $Method, $Params, $Id)
    if ($null -eq $Id) {
        $Id = $Session.NextId
        $Session.NextId = $Session.NextId + 1
    }
    Send-McpMessage -Session $Session -Method $Method -Params $Params -Id $Id
    return (Read-McpReply -Session $Session -Id $Id)
}

# One tool call, its answer parsed (Get-McpToolResult).
function Invoke-McpTool {
    param([hashtable] $Session, [string] $Tool, [hashtable] $Arguments = @{}, $Id)
    $reply = Invoke-McpRequest -Session $Session -Method 'tools/call' -Params @{ name = $Tool; arguments = $Arguments } -Id $Id
    return (Get-McpToolResult -Reply $reply -Tool $Tool)
}

# Ends the session: the server exits when its input closes. It is given time rather than killed - it owns a COM host
# attached to Outlook. Returns what it wrote to standard error, or ''.
function Stop-McpStdioServer {
    param([hashtable] $Session, [int] $WaitMilliseconds = 30000)
    if ($null -eq $Session) { return '' }
    if ($null -ne $Session.Stdin) { try { $Session.Stdin.Dispose() } catch { } }
    if ($null -ne $Session.Process -and -not $Session.Process.HasExited) { [void]$Session.Process.WaitForExit($WaitMilliseconds) }
    $stderr = ''
    if ($null -ne $Session.Stderr -and $Session.Stderr.IsCompleted) { try { $stderr = [string]$Session.Stderr.Result } catch { } }
    return $stderr
}
