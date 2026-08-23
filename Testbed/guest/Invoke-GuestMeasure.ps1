<#
    RECOVERED ARTEFACT. This is the driver that produced the sweep and frame numbers now quoted
    in Docs/magic-numbers.md and Docs/autonomous-session-log.md. It existed only on the guest, at
    C:\OutlookAI-Q5\guest-measure.ps1, and was pulled back to the host on 2026-08-24 during the
    recovery that also found the corpus manifest. It is committed here so that never depends on
    one machine again.

    Changed from the guest copy, and nothing else:
      * the header comment, which described the EARLIER 40,000-item build and said every item was
        dated roughly now. That was true of the run it was written for and is false of the corpus
        it actually measured;
      * the store display name, which was hard-coded twice, is now a parameter.

    Runs INSIDE the guest console session, Windows PowerShell 5.1. Reach session 1 with
    Register-InteractiveTask.ps1 - Outlook can never finish starting in session 0, and this script
    spawns the MCP server, which spawns a COM host, which attaches to Outlook.

    WHAT IT MEASURES, against the corpus whose parameters are in Testbed/testbed.json
    (vm2 / 7777 / 2026-08-19 / 20000):

      * four freshness sweeps - one with a term, one filter-only, one immediate repeat to see the
        10 s cache, one with a term that matches nothing. The unindexed store has no index
        frontier, so the sweep takes the seven-day fallback window, which this corpus fills with
        1,612 items across four folders - enough that the 200-per-folder cap engages and
        `item_cap_unsorted` is raised. That is the shape a sweep budget has to survive.
      * the COM host's frame high-water, before and after, off outlook_health.
      * two exhaustive scans, both with a term matching nothing so the scan runs to its budget
        rather than stopping early at the result cap: one folder, then the whole store over a
        365-day window.

    WHAT IT PRODUCED, 2026-08-19, so a re-run can be compared rather than merely read:
      sweeps      13,624 / 11,818 / 10,652 / 11,889 ms, itemsSeen 758 on all four
      frame       572 bytes at start, 10,734,599 bytes after the sweeps (limit 67,108,864)
      index       perStore = Outlook Data File False   <- the store is OUT of the index, as intended
      scans       Inbox-only 1 folder; whole-store 365d 10 folders; neither timed out

    Output is JSON Lines at -OutFile with `###` marker lines a human reads. It is measurement data
    about one machine: it belongs in the gitignored fixtures directory, never in this repository.
    .github/scripts/check-measurement-privacy.ps1 fails a build if one lands.
#>

param(
    [string]$ServerExe = 'C:\OutlookAI-Q5\server\OutlookAI.McpServer.exe',
    [string]$OutFile   = 'C:\OutlookAI-Q5\measure.jsonl',
    [string]$Store     = 'Outlook Data File',
    [int]$ReplyTimeoutSeconds = 600
)

$ErrorActionPreference = 'Stop'
Remove-Item $OutFile -Force -ErrorAction SilentlyContinue

$psi = New-Object Diagnostics.ProcessStartInfo
$psi.FileName = $ServerExe
$psi.WorkingDirectory = Split-Path $ServerExe -Parent
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)

$proc = [Diagnostics.Process]::Start($psi)
$stdin = New-Object IO.StreamWriter($proc.StandardInput.BaseStream, (New-Object Text.UTF8Encoding($false)))
$stdin.AutoFlush = $true
$stderrTask = $proc.StandardError.ReadToEndAsync()

function Send-Rpc { param([string]$Method, $Params, $Id, [switch]$Notification)
    $msg = [ordered]@{ jsonrpc = '2.0'; method = $Method }
    if (-not $Notification) { $msg['id'] = $Id }
    if ($null -ne $Params) { $msg['params'] = $Params }
    $json = $msg | ConvertTo-Json -Depth 20 -Compress
    Add-Content -Path $OutFile -Value ">>> $json"
    $stdin.WriteLine($json)
}
function Read-Reply { param($ForId)
    $deadline = (Get-Date).AddSeconds($ReplyTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($proc.HasExited -and $proc.StandardOutput.EndOfStream) { throw "server exited $($proc.ExitCode)" }
        $remaining = [int]([Math]::Max(1000, ($deadline - (Get-Date)).TotalMilliseconds))
        $task = $proc.StandardOutput.ReadLineAsync()
        if (-not $task.Wait($remaining)) { throw "read timed out for id $ForId" }
        $line = $task.Result
        if ($null -eq $line) { throw "stdout closed before id $ForId" }
        if ($line.Trim() -eq '') { continue }
        Add-Content -Path $OutFile -Value "<<< $line"
        $o = $null
        try { $o = $line | ConvertFrom-Json } catch { continue }
        if (($o.PSObject.Properties.Name -contains 'id') -and ("$($o.id)" -eq "$ForId")) { return $o }
    }
    throw "no reply to id $ForId"
}
function Call { param([int]$Id, [string]$Tool, $CallArgs, [string]$Label)
    Add-Content -Path $OutFile -Value "### CASE $Id : $Label"
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Send-Rpc -Method 'tools/call' -Id $Id -Params @{ name = $Tool; arguments = $CallArgs }
    $r = Read-Reply -ForId $Id
    $sw.Stop()
    Add-Content -Path $OutFile -Value "### WALL $Id : $($sw.ElapsedMilliseconds) ms"
    return $r
}
function NoteSweep { param($Reply, [int]$Id)
    try {
        $d = $Reply.result.content[0].text | ConvertFrom-Json
        $s = $d.sweep
        Add-Content -Path $OutFile -Value ("### SWEEP {0} : performed={1} elapsedMs={2} foldersSwept={3} itemsSeen={4} cached={5} gaps={6} bodiesCapped={7} err={8}" -f `
            $Id, $s.performed, $s.elapsedMs, $s.foldersSwept, $s.itemsSeen, $s.cached, ($s.coverageGaps -join ','), $s.itemsBodyCapped, ($s.error -replace '\s+',' '))
        Add-Content -Path $OutFile -Value ("### TOP   {0} : degraded={1} freshness={2} hits={3}" -f $Id, $d.degraded, $d.freshness, $d.hits.Count)
    } catch { Add-Content -Path $OutFile -Value "### SWEEP $Id : unreadable" }
}
function NoteHealth { param($Reply, [string]$When)
    try {
        $h = $Reply.result.content[0].text | ConvertFrom-Json
        $c = $h.outlook.comHost
        Add-Content -Path $OutFile -Value ("### METER {0} : largest={1} limit={2} refused={3} restarts={4} state={5}" -f `
            $When, $c.largestFrameBytes, $c.frameLimitBytes, $c.framesRefusedTooLarge, $c.restartCount, $h.outlook.state)
        Add-Content -Path $OutFile -Value ("### INDEX {0} : perStore={1}" -f $When, ((($h.index.perStore) | ForEach-Object { "$($_.store)=$($_.inLocalIndex)" }) -join ' '))
    } catch { Add-Content -Path $OutFile -Value "### METER $When : unreadable" }
}

try {
    Send-Rpc -Method 'initialize' -Id 1 -Params @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = 'measure'; version = '1.0' } }
    Read-Reply -ForId 1 | Out-Null
    Send-Rpc -Method 'notifications/initialized' -Notification

    $wid = 100; $ok = $false; $until = (Get-Date).AddMinutes(5)
    while ((Get-Date) -lt $until) {
        Send-Rpc -Method 'tools/call' -Id $wid -Params @{ name = 'outlook_health'; arguments = @{} }
        $h = Read-Reply -ForId $wid
        try { if (($h.result.content[0].text | ConvertFrom-Json).outlook.comConnected) { $ok = $true } } catch { }
        if ($ok) { break }
        $wid++; Start-Sleep -Seconds 10
    }
    Add-Content -Path $OutFile -Value "### COM connected: $ok"
    NoteHealth (Call -Id 9 -Tool 'outlook_health' -CallArgs @{} -Label 'baseline') 'start'

    $fl = Call -Id 10 -Tool 'list_folders' -CallArgs @{ store = $Store } -Label 'folder inventory'
    try {
        $d = $fl.result.content[0].text | ConvertFrom-Json
        foreach ($f in ($d.stores[0].folders | Where-Object { $_.items -gt 0 })) {
            Add-Content -Path $OutFile -Value ("### FOLDER : {0} items={1} unread={2}" -f $f.path, $f.items, $f.unread)
        }
    } catch { Add-Content -Path $OutFile -Value '### FOLDER : unreadable' }

    # The sweep. A term-bearing search, then a filter-only one, then a repeat to see the cache.
    # Waits between runs so the 10 s sweep cache cannot serve a stale answer as a fast one.
    NoteSweep (Call -Id 20 -Tool 'search' -CallArgs @{ query = 'corpus'; top = 20 } -Label 'sweep, term') 20
    Start-Sleep -Seconds 12
    NoteSweep (Call -Id 21 -Tool 'search' -CallArgs @{ top = 20 } -Label 'sweep, filter-only') 21
    NoteSweep (Call -Id 22 -Tool 'search' -CallArgs @{ top = 20 } -Label 'sweep, immediate repeat (expect cached)') 22
    Start-Sleep -Seconds 12
    NoteSweep (Call -Id 23 -Tool 'search' -CallArgs @{ query = 'zzzznomatch'; top = 20 } -Label 'sweep, term matching nothing') 23

    NoteHealth (Call -Id 30 -Tool 'outlook_health' -CallArgs @{} -Label 'after sweeps') 'after-sweeps'

    # Scan throughput: a term matching nothing, so the scan runs to its budget rather than
    # stopping early at the result cap.
    # An exhaustive search needs a bound: a folder, or a date. Both shapes are measured, because
    # they stress different things - one folder's rows against the whole store's folder walk.
    $scans = @(
        @{ id = 40; label = 'exhaustive, Inbox only, no match';       args = @{ query = 'zzzznomatch'; store = $Store; folder = 'Inbox'; include_subfolders = $false; exhaustive = $true; top = 100 } },
        @{ id = 41; label = 'exhaustive, whole store, 365-day window'; args = @{ query = 'zzzznomatch'; store = $Store; after = '2025-08-19'; exhaustive = $true; top = 100 } }
    )
    foreach ($sc in $scans) {
        $r = Call -Id $sc.id -Tool 'search' -CallArgs $sc.args -Label $sc.label
        try {
            $d = $r.result.content[0].text | ConvertFrom-Json
            $e = $d.exhaustive
            if ($null -eq $e) {
                Add-Content -Path $OutFile -Value ("### SCAN {0} : no exhaustive block - {1}" -f $sc.id, ($d.error.message))
            } else {
                Add-Content -Path $OutFile -Value ("### SCAN {0} : foldersScanned={1} foldersSkipped={2} rowsDropped={3} timedOut={4} truncated={5} gaps={6}" -f `
                    $sc.id, $e.foldersScanned, $e.foldersSkipped, $e.rowsDropped, $e.timedOut, $e.truncated, ($e.coverageGaps -join ','))
            }
        } catch { Add-Content -Path $OutFile -Value ("### SCAN {0} : unreadable" -f $sc.id) }
    }

    NoteHealth (Call -Id 50 -Tool 'outlook_health' -CallArgs @{} -Label 'final') 'final'
}
catch { Add-Content -Path $OutFile -Value "### ERROR: $($_.Exception.Message)" }
finally {
    try { $stdin.Close() } catch { }
    if (-not $proc.WaitForExit(20000)) { try { $proc.Kill() } catch { } }
    $err = ''; try { $err = $stderrTask.Result } catch { }
    if ($err) { Add-Content -Path $OutFile -Value "### stderr:`r`n$err" }
    Add-Content -Path $OutFile -Value '### done'
}
