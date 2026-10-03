#Requires -Version 5.1
<#
.SYNOPSIS
    Q118 probe: counts Outlook's Explorers and windows around REAL OutlookAI server sessions, and
    times whether OUTLOOK.EXE then ends after Application.Quit() or a user's File > Exit.

.DESCRIPTION
    RUN ON A TEST GUEST, in the interactive session, NOT elevated - through
    guest/Register-InteractiveTask.ps1 at -RunLevel Limited, staged beside it and beside
    guest/OutlookMapiInterop.ps1 (Testbed/README.md section 4c says why the level matters: COM does
    not attach across integrity levels, and a user's Outlook is not elevated):

        .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 1800 -Script "& 'C:\OutlookAI-Q5\q118\Measure-PinWindows.ps1' -Scenario Visible"

    WHAT IT ANSWERS (Q118, 2026-10-03). Every OutlookAI server session runs its own COM host, and
    OutlookComSession.Connect adds a never-displayed "lifetime pin" Explorer when Outlook has NO
    Explorer at all (D49, ComposeSurface.TryPinProcess); a session closes its pin on a normal exit
    only if that session STARTED Outlook. The claim under test: pins left behind accumulate, and stop
    a later Application.Quit() - or the user's File > Exit - from ending OUTLOOK.EXE. And S6: whether
    Application.ActiveExplorer() hands back such a hidden Explorer, which the show-me path would then
    display as the user's window.

    WHAT IT DOES. Drives the staged server (-ServerExe) over stdio exactly as an MCP client does -
    initialize, then list_folders, a read-only tool - and ends each session the normal way (stdin
    closed) or by killing ITS COM HOST ONLY. Around each step a separate child process - so nothing
    it attaches can outlive the reading - counts Application.Explorers and each one's window
    (IOleWindow, IsWindowVisible), what ActiveExplorer() returns, the Inspectors, every store's
    Outbox, and every top-level window of OUTLOOK.EXE by class. Then it ends Outlook the graceful way
    the scenario names, and times whether and when OUTLOOK.EXE leaves.

    WHAT IT NEVER DOES. No mail item is created, changed or moved; the only tool called is
    list_folders. OUTLOOK.EXE is never killed (mailbox-safety rule 7): Application.Quit(), File >
    Exit, a window close and the closing of HIDDEN Explorers each run only with no Inspector open,
    every Outbox empty and no visible Outlook dialog; closing hidden Explorers is refused while any
    Outlook window is visible; and an Outlook that does not end is LEFT RUNNING and reported.
    Stop-Process is used on exactly two kinds of process, each checked by name first: a COM host
    (OutlookAI.ComHost) the scenario kills on purpose, and this script's own child probe when it
    outruns its time limit.

    SCENARIOS (-Scenario):
      Control        visible Outlook, no OutlookAI session ever attached: Quit; then File > Exit.
      ControlClose   the same, ended by the window's close button (UI Automation WindowPattern).
      Visible        visible Outlook: sessions one at a time, three at once, one whose COM host is
                     killed; Quit. Then a second round ending with File > Exit.
      Started        no Outlook: a session starts it headless, a second attaches at once, both end
                     normally; then, while it lingers, three at once and two one at a time; Quit.
      Sequential     no Outlook: a session starts it, waits until it is ready and ends normally;
                     a second arrives while it lingers and ends normally; a third; Quit.
      Killed         no Outlook: a session starts it and its COM host is killed; more sessions; Quit.
      KilledVisible  as Killed, then the user opens Outlook's window; File > Exit. Then the same
                     twice more, ended by Quit and by the window's close button.
      ShowMe         as KilledVisible, the user closes the window; then the two calls the show-me path
                     (EnsureVisibleExplorer) makes on what ActiveExplorer() returns - un-minimise and
                     Activate() - and the user's close of whatever that showed. A UI change only.
      Headless       a plain COM client (this script, -Mode Hold) starts Outlook headless and holds
                     it; three sessions at once, then one; the holder lets go; Quit.
    A scenario that leaves Outlook running tries to close the hidden Explorers last, under the same
    preconditions, and says whether that ended it.

    OUTPUT. Readable lines on stdout (out.txt of the interactive task) and one JSON object per event
    in <LogDir>\q118-<scenario>-<stamp>.jsonl.

    THE GUARD. vmadmin (Assert-TestbedGuest, guest/OutlookMapiInterop.ps1) AND a computer name
    starting OAI- that is not the build VM's, before anything else but -SelfTest runs.

    RUN 2026-10-03 on OAI-UNINDEXED for Q118 - the results are in the commit that added this file.

.PARAMETER Scenario
    One of the scenarios above. Required in the default mode.

.PARAMETER Mode
    Scenario (default). State, Quit, CloseHidden and Hold are the child modes the scenarios start;
    they are not meant to be run by hand.

.PARAMETER OutFile
    Child modes: where the child writes its JSON result.

.PARAMETER ReleaseFlag
    -Mode Hold: the holder lets Outlook go when this file appears (or after 30 minutes).

.PARAMETER ServerExe
    The staged OutlookAI.McpServer.exe. Its COM host sits beside it.

.PARAMETER SelfTest
    Pure: compiles the helper types, checks the JSON-RPC framing, the rule-7 decision and the
    scenario table, and reads this file's syntax tree for the two Stop-Process targets. Any machine.

.EXAMPLE
    .\Register-InteractiveTask.ps1 -RunLevel Limited -TimeoutSeconds 1800 -Script "& 'C:\OutlookAI-Q5\q118\Measure-PinWindows.ps1' -Scenario Control"
    .\Measure-PinWindows.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [ValidateSet('Scenario', 'State', 'Quit', 'CloseHidden', 'Hold', 'ShowActive')] [string] $Mode = 'Scenario',
    [ValidateSet('', 'Control', 'ControlClose', 'Visible', 'Started', 'Sequential', 'Killed', 'KilledVisible', 'ShowMe', 'Headless')] [string] $Scenario = '',
    [string]   $OutFile,
    [string]   $ReleaseFlag,
    [string]   $ServerExe = 'C:\OutlookAI-Q5\server\OutlookAI.McpServer.exe',
    [string]   $Profile = 'OutlookAI-Tier',
    [string]   $OutlookExe = 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
    [string]   $LogDir = 'C:\OutlookAI-Q5\q118',
    [int]      $ExitWaitSeconds = 120,
    [string[]] $ExpectedUser = @('vmadmin'),
    [switch]   $SelfTest
)

$ErrorActionPreference = 'Stop'

$script:ProbeSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

[ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOaiPinOleWindow
{
    [PreserveSig] int GetWindow(out IntPtr phwnd);
    [PreserveSig] int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
}

public sealed class OaiPinWindow
{
    public long Handle; public string Class; public bool Visible; public bool Minimized; public string Title;
    public int Left; public int Top; public int Right; public int Bottom;
}

public static class OaiPinNative
{
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);

    public static List<OaiPinWindow> Snapshot(int[] pids)
    {
        HashSet<uint> set = new HashSet<uint>();
        List<OaiPinWindow> list = new List<OaiPinWindow>();
        if (pids == null) { return list; }
        foreach (int p in pids) { set.Add((uint)p); }
        if (set.Count == 0) { return list; }
        EnumWindows(delegate (IntPtr h, IntPtr l)
        {
            uint pid;
            GetWindowThreadProcessId(h, out pid);
            if (!set.Contains(pid)) { return true; }
            StringBuilder c = new StringBuilder(256); GetClassName(h, c, 256);
            StringBuilder t = new StringBuilder(256); GetWindowText(h, t, 256);
            RECT r; GetWindowRect(h, out r);
            OaiPinWindow w = new OaiPinWindow();
            w.Handle = h.ToInt64(); w.Class = c.ToString(); w.Visible = IsWindowVisible(h); w.Minimized = IsIconic(h);
            w.Title = t.ToString(); w.Left = r.Left; w.Top = r.Top; w.Right = r.Right; w.Bottom = r.Bottom;
            list.Add(w);
            return true;
        }, IntPtr.Zero);
        return list;
    }
}

public sealed class OaiPinExplorer
{
    public int Index; public long Hwnd; public string Visible; public string Caption; public int WindowState; public long Identity;
}

public sealed class OaiPinComState
{
    public bool Attached; public string AttachError;
    public int ExplorerCount = -1; public List<OaiPinExplorer> Explorers = new List<OaiPinExplorer>();
    public int InspectorCount = -1;
    public int DefaultOutbox = -1; public int OutboxTotal = -1; public int OutboxStores; public string OutboxErrors = "";
    public string Active = "not read"; public int ActiveIndex = -1; public long ActiveHwnd; public string ActiveVisible = "unknown";
    public string Errors = "";
    public string QuitResult = ""; public int Closed = -1;
}

public static class OaiPinCom
{
    public static object Get(object o, string name) { return o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, null); }
    public static object Call(object o, string name, object[] args) { return o.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, o, args); }

    public static void Release(object o)
    {
        if (o != null && Marshal.IsComObject(o)) { try { Marshal.FinalReleaseComObject(o); } catch (Exception) { } }
    }

    public static long Hwnd(object o)
    {
        try
        {
            IOaiPinOleWindow w = o as IOaiPinOleWindow;
            if (w == null) { return 0; }
            IntPtr h;
            return w.GetWindow(out h) == 0 ? h.ToInt64() : 0;
        }
        catch (Exception) { return -1; }
    }

    public static long Identity(object o)
    {
        IntPtr p = Marshal.GetIUnknownForObject(o);
        try { return p.ToInt64(); } finally { Marshal.Release(p); }
    }

    // The product's own attach (OutlookComSession.Connect): the class object of the RUNNING
    // instance. The Running Object Table is not used - Office joins it only once its window loses
    // focus (KB 238610), which on an unattended guest may be never.
    public static object Attach()
    {
        Type t = Type.GetTypeFromProgID("Outlook.Application");
        if (t == null) { throw new InvalidOperationException("Outlook.Application is not registered."); }
        return Activator.CreateInstance(t);
    }

    static void Collect() { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers(); }

    // Reads; changes nothing. Every reference is released before it returns.
    public static OaiPinComState Read()
    {
        OaiPinComState s = new OaiPinComState();
        object app = null, explorers = null, inspectors = null, ns = null, stores = null, active = null;
        List<object> held = new List<object>();
        try
        {
            try { app = Attach(); s.Attached = true; }
            catch (Exception ex) { s.AttachError = ex.GetType().Name + ": " + ex.Message; return s; }

            try
            {
                explorers = Get(app, "Explorers");
                s.ExplorerCount = (int)Get(explorers, "Count");
                for (int i = 1; i <= s.ExplorerCount; i++)
                {
                    object e = Call(explorers, "Item", new object[] { i });
                    held.Add(e);
                    OaiPinExplorer x = new OaiPinExplorer();
                    x.Index = i;
                    x.Hwnd = Hwnd(e);
                    x.Visible = x.Hwnd > 0 ? (OaiPinNative.IsWindowVisible(new IntPtr(x.Hwnd)) ? "visible" : "hidden") : "unknown";
                    try { x.Caption = (string)Get(e, "Caption"); } catch (Exception ex) { x.Caption = "<" + ex.GetType().Name + ">"; }
                    try { x.WindowState = (int)Get(e, "WindowState"); } catch (Exception) { x.WindowState = -1; }
                    x.Identity = Identity(e);
                    s.Explorers.Add(x);
                }
            }
            catch (Exception ex) { s.Errors += "explorers: " + ex.GetType().Name + "; "; }

            try
            {
                active = Call(app, "ActiveExplorer", new object[0]);
                if (active == null) { s.Active = "null"; }
                else
                {
                    s.Active = "explorer";
                    long id = Identity(active);
                    foreach (OaiPinExplorer x in s.Explorers) { if (x.Identity == id) { s.ActiveIndex = x.Index; } }
                    s.ActiveHwnd = Hwnd(active);
                    s.ActiveVisible = s.ActiveHwnd > 0 ? (OaiPinNative.IsWindowVisible(new IntPtr(s.ActiveHwnd)) ? "visible" : "hidden") : "unknown";
                }
            }
            catch (Exception ex) { s.Active = "error " + ex.GetType().Name; }

            try
            {
                inspectors = Get(app, "Inspectors");
                s.InspectorCount = (int)Get(inspectors, "Count");
            }
            catch (Exception ex) { s.Errors += "inspectors: " + ex.GetType().Name + "; "; }

            // The Outboxes, the way host/Restart-Guest.ps1 reads them before its Quit: the DEFAULT
            // store's must be readable (fail closed); any other store's where it has one - a data
            // file that is nobody's delivery store may have none, and then nothing can be queued in it.
            try
            {
                ns = Call(app, "GetNamespace", new object[] { "MAPI" });
                object dob = null, doi = null;
                try { dob = Call(ns, "GetDefaultFolder", new object[] { 4 }); doi = Get(dob, "Items"); s.DefaultOutbox = (int)Get(doi, "Count"); }
                catch (Exception ex) { s.Errors += "default outbox: " + ex.GetType().Name + "; "; }
                finally { Release(doi); Release(dob); }
                stores = Get(ns, "Stores");
                int n = (int)Get(stores, "Count");
                int total = 0;
                for (int i = 1; i <= n; i++)
                {
                    object store = null, outbox = null, items = null;
                    try
                    {
                        store = Call(stores, "Item", new object[] { i });
                        outbox = Call(store, "GetDefaultFolder", new object[] { 4 }); // olFolderOutbox
                        items = Get(outbox, "Items");
                        total += (int)Get(items, "Count");
                        s.OutboxStores++;
                    }
                    catch (Exception ex) { s.OutboxErrors += "store " + i + ": " + ex.GetType().Name + "; "; }
                    finally { Release(items); Release(outbox); Release(store); }
                }
                s.OutboxTotal = total;
            }
            catch (Exception ex) { s.Errors += "outbox: " + ex.GetType().Name + "; "; }
        }
        finally
        {
            Release(active);
            foreach (object e in held) { Release(e); }
            Release(stores); Release(ns); Release(inspectors); Release(explorers); Release(app);
            Collect();
        }
        return s;
    }

    // Application.Quit() and nothing else. The caller has already applied rule 7.
    public static string Quit()
    {
        object app = null;
        try
        {
            app = Attach();
            Call(app, "Quit", new object[0]);
            return "quit-called";
        }
        catch (Exception ex) { return "quit-failed " + ex.GetType().Name + ": " + ex.Message; }
        finally { Release(app); Collect(); }
    }

    // Closes every Explorer. The caller has proved no Outlook window is visible, so every one of
    // them is a hidden Explorer - which is exactly what OutlookComSession.TryCloseInvisibleExplorers
    // assumes too.
    public static int CloseAllExplorers()
    {
        object app = null, explorers = null;
        int closed = 0;
        try
        {
            app = Attach();
            explorers = Get(app, "Explorers");
            int count = (int)Get(explorers, "Count");
            for (int i = count; i >= 1; i--)
            {
                object e = null;
                try { e = Call(explorers, "Item", new object[] { i }); Call(e, "Close", new object[0]); closed++; }
                catch (Exception) { }
                finally { Release(e); }
            }
        }
        finally { Release(explorers); Release(app); Collect(); }
        return closed;
    }

    // What OutlookComSession.EnsureVisibleExplorer does with an Explorer that ActiveExplorer() hands
    // back and ComposeSurface.IsPin does not know - and IsPin never knows another process's pin:
    // un-minimise it, then Activate() it. Returns what it found and what became of the window.
    public static string ShowActiveExplorer()
    {
        object app = null, active = null;
        try
        {
            app = Attach();
            active = Call(app, "ActiveExplorer", new object[0]);
            if (active == null) { return "ActiveExplorer() returned null; nothing shown"; }
            long hwnd = Hwnd(active);
            string before = hwnd > 0 ? (OaiPinNative.IsWindowVisible(new IntPtr(hwnd)) ? "visible" : "hidden") : "unknown";
            int state = -1;
            try { state = (int)Get(active, "WindowState"); } catch (Exception) { }
            if (state == 1) { try { active.GetType().InvokeMember("WindowState", BindingFlags.SetProperty, null, active, new object[] { 2 }); } catch (Exception) { } }
            string activate = "Activate() returned";
            try { Call(active, "Activate", new object[0]); } catch (Exception ex) { activate = "Activate() threw " + ex.GetType().Name; }
            System.Threading.Thread.Sleep(1500);
            string after = hwnd > 0 ? (OaiPinNative.IsWindowVisible(new IntPtr(hwnd)) ? "visible" : "hidden") : "unknown";
            return "ActiveExplorer() window was " + before + " (WindowState " + state + "); " + activate + "; window now " + after;
        }
        catch (Exception ex) { return "failed " + ex.GetType().Name + ": " + ex.Message; }
        finally { Release(active); Release(app); Collect(); }
    }

    // A plain COM client, as any other program on the machine might be: starts Outlook headless
    // when none runs, holds Application and NameSpace until told to let go.
    public static string Hold(string readyFile, string releaseFlag, int maxMinutes)
    {
        object app = null, ns = null;
        try
        {
            app = Attach();
            ns = Call(app, "GetNamespace", new object[] { "MAPI" });
            try { Call(ns, "Logon", new object[] { Missing.Value, Missing.Value, false, false }); } catch (Exception) { }
            File.WriteAllText(readyFile, "holding");
            Stopwatch sw = Stopwatch.StartNew();
            while (!File.Exists(releaseFlag) && sw.Elapsed.TotalMinutes < maxMinutes) { System.Threading.Thread.Sleep(250); }
            return File.Exists(releaseFlag) ? "released-on-flag" : "released-on-timeout";
        }
        finally { Release(ns); Release(app); Collect(); }
    }
}

// One MCP stdio session against the real server, newline-delimited JSON-RPC as every client
// speaks it. Writes raw UTF-8 bytes (no BOM) to the pipe and never touches the StreamWriter, so
// nothing but the frames reaches the server.
public sealed class OaiMcpSession
{
    readonly Process _p;
    readonly Stream _in;
    readonly Task<string> _stderr;
    Task<string> _pendingRead;
    int _id;
    public string Name;

    public OaiMcpSession(string exe, string name)
    {
        Name = name;
        ProcessStartInfo psi = new ProcessStartInfo(exe);
        psi.UseShellExecute = false;
        psi.RedirectStandardInput = true;
        psi.RedirectStandardOutput = true;
        psi.RedirectStandardError = true;
        psi.CreateNoWindow = true;
        psi.WorkingDirectory = Path.GetDirectoryName(exe);
        psi.StandardOutputEncoding = new UTF8Encoding(false);
        psi.StandardErrorEncoding = new UTF8Encoding(false);
        _p = Process.Start(psi);
        _in = _p.StandardInput.BaseStream;
        _stderr = _p.StandardError.ReadToEndAsync();
    }

    public int Pid { get { return _p.Id; } }
    public bool HasExited { get { return _p.HasExited; } }
    public int ExitCode { get { return _p.HasExited ? _p.ExitCode : -999; } }

    void Write(string json)
    {
        byte[] b = new UTF8Encoding(false).GetBytes(json + "\n");
        _in.Write(b, 0, b.Length);
        _in.Flush();
    }

    public static string Frame(int id, string method, string paramsJson)
    {
        return "{\"jsonrpc\":\"2.0\",\"id\":" + id + ",\"method\":\"" + method + "\",\"params\":" + paramsJson + "}";
    }

    public static bool IsResponseTo(string line, int id)
    {
        if (line == null) { return false; }
        string compact = line.Replace(" ", "");
        bool idMatches = compact.Contains("\"id\":" + id + ",") || compact.Contains("\"id\":" + id + "}");
        return idMatches && (compact.Contains("\"result\":") || compact.Contains("\"error\":"));
    }

    public int Send(string method, string paramsJson)
    {
        int id = ++_id;
        Write(Frame(id, method, paramsJson));
        return id;
    }

    public void Notify(string method) { Write("{\"jsonrpc\":\"2.0\",\"method\":\"" + method + "\"}"); }

    public string Await(int id, int timeoutMs)
    {
        Stopwatch sw = Stopwatch.StartNew();
        while (true)
        {
            int left = timeoutMs - (int)sw.ElapsedMilliseconds;
            if (left <= 0) { return null; }
            if (_pendingRead == null) { _pendingRead = _p.StandardOutput.ReadLineAsync(); }
            if (!_pendingRead.Wait(left)) { return null; }
            string line = _pendingRead.Result;
            _pendingRead = null;
            if (line == null) { return null; }
            if (IsResponseTo(line, id)) { return line; }
        }
    }

    // How a client ends a session normally: it closes stdin.
    public void CloseInput() { try { _in.Close(); } catch (Exception) { } }
    public bool WaitExit(int ms) { return _p.WaitForExit(ms); }

    public string StderrTail(int max)
    {
        if (!_stderr.Wait(3000)) { return "<stderr still open>"; }
        string s = _stderr.Result ?? "";
        return s.Length <= max ? s : s.Substring(s.Length - max);
    }
}
'@

function Add-ProbeTypes {
    if (-not ('OaiPinCom' -as [type])) {
        Add-Type -TypeDefinition $script:ProbeSource -Language CSharp
    }
}

# Rule 7, as a pure decision: $null means the graceful end may run.
function Get-Rule7Refusal {
    param([int] $InspectorCount, [int] $DefaultOutbox, [int] $OutboxTotal, [int] $VisibleDialogs, [bool] $Attached = $true)
    if (-not $Attached) { return 'could not attach to Outlook, so its Inspectors and Outbox are unread' }
    if ($InspectorCount -ne 0) { return "an Inspector is open ($InspectorCount) - possibly an unsent compose window" }
    if ($DefaultOutbox -ne 0) { return "the default store's Outbox is not proven empty ($DefaultOutbox; -1 is unreadable)" }
    if ($OutboxTotal -ne 0) { return "an Outbox is not proven empty (total $OutboxTotal over the stores that have one; -1 is unread)" }
    if ($VisibleDialogs -ne 0) { return "a visible Outlook dialog ($VisibleDialogs) - Quit behind one parks" }
    return $null
}

function Invoke-SelfTest {
    $script:p = 0; $script:f = 0
    function Check([string] $name, [bool] $ok) { if ($ok) { $script:p++; Write-Host "  PASS  $name" } else { $script:f++; Write-Host "  FAIL  $name" } }

    $compiled = $true
    try { Add-ProbeTypes } catch { $compiled = $false; Write-Host "  compile error: $($_.Exception.Message)" }
    Check 'the helper types compile under this PowerShell' $compiled

    if ($compiled) {
        $frame = [OaiMcpSession]::Frame(10, 'tools/call', '{"name":"list_folders","arguments":{}}')
        $parsed = $null
        try { $parsed = $frame | ConvertFrom-Json } catch { }
        Check 'a request frame is valid JSON' ($null -ne $parsed)
        Check 'a request frame carries its id and method' ($parsed.id -eq 10 -and $parsed.method -eq 'tools/call' -and $parsed.params.name -eq 'list_folders')
        Check 'the response matcher takes its own id' ([OaiMcpSession]::IsResponseTo('{"result":{},"id":1,"jsonrpc":"2.0"}', 1))
        Check 'the response matcher does not take id 10 for id 1' (-not [OaiMcpSession]::IsResponseTo('{"result":{},"id":10,"jsonrpc":"2.0"}', 1))
        Check 'the response matcher takes an error answer' ([OaiMcpSession]::IsResponseTo('{"jsonrpc":"2.0","id":3,"error":{"code":-32601}}', 3))
        Check 'the response matcher ignores a notification' (-not [OaiMcpSession]::IsResponseTo('{"jsonrpc":"2.0","method":"notifications/message","params":{"id":1}}', 1))
        $w = [OaiPinNative]::Snapshot([int[]]@())
        Check 'a window snapshot of no process is empty' ($w.Count -eq 0)
    }

    Check 'rule 7 allows a clean state' ($null -eq (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 0 -OutboxTotal 0 -VisibleDialogs 0))
    Check 'rule 7 refuses an open Inspector' ($null -ne (Get-Rule7Refusal -InspectorCount 1 -DefaultOutbox 0 -OutboxTotal 0 -VisibleDialogs 0))
    Check 'rule 7 refuses a non-empty default Outbox' ($null -ne (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 1 -OutboxTotal 1 -VisibleDialogs 0))
    Check 'rule 7 refuses an unreadable default Outbox' ($null -ne (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox -1 -OutboxTotal 0 -VisibleDialogs 0))
    Check 'rule 7 refuses an item in another store''s Outbox' ($null -ne (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 0 -OutboxTotal 2 -VisibleDialogs 0))
    Check 'rule 7 refuses Outboxes it never counted' ($null -ne (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 0 -OutboxTotal -1 -VisibleDialogs 0))
    Check 'rule 7 refuses a visible dialog' ($null -ne (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 0 -OutboxTotal 0 -VisibleDialogs 1))
    Check 'rule 7 refuses an Outlook it could not attach to' ($null -ne (Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 0 -OutboxTotal 0 -VisibleDialogs 0 -Attached $false))

    $ast = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref] $null, [ref] $null)
    $defined = @{}
    foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fd.Name] = $true }
    $declared = @((Get-Command -Name $PSCommandPath).Parameters['Scenario'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | ForEach-Object { $_.ValidValues } | Where-Object { $_ })
    $missing = @($declared | Where-Object { -not $defined.ContainsKey("Invoke-Scenario$_") })
    Check 'every -Scenario value has its Invoke-Scenario function' ($declared.Count -eq 9 -and $missing.Count -eq 0)
    $dispatch = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst] -and $n.Condition.Extent.Text -eq '($Scenario)' }, $true))
    $clauses = @()
    if ($dispatch.Count -eq 1) { $clauses = @($dispatch[0].Clauses | ForEach-Object { $_.Item1.Value }) }
    Check 'the dispatcher has one clause per scenario' (($dispatch.Count -eq 1) -and (@($declared | Where-Object { $clauses -notcontains $_ }).Count -eq 0))

    # Stop-Process appears only inside the two functions allowed to stop a process, and each of them
    # checks the process name first. Nothing in this file names taskkill.
    $stops = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Stop-Process' }, $true))
    $owners = @($stops | ForEach-Object {
            $q = $_.Parent
            while ($q -and -not ($q -is [System.Management.Automation.Language.FunctionDefinitionAst])) { $q = $q.Parent }
            if ($q) { $q.Name } else { '<script>' }
        } | Sort-Object -Unique)
    Check 'Stop-Process is called only by Stop-ComHost and Invoke-ChildProbe' (($owners -join ',') -eq 'Invoke-ChildProbe,Stop-ComHost')
    $text = Get-Content -LiteralPath $PSCommandPath -Raw
    Check 'nothing here calls taskkill' (-not ($text -match ('task' + 'kill\.exe|task' + 'kill /')))

    Write-Host ''
    Write-Host "SelfTest: $script:p passed, $script:f failed"
    if ($script:f -gt 0) { exit 1 }
    exit 0
}

if ($SelfTest) { Invoke-SelfTest }

# THE GUARD, before anything that touches a process, a file or COM.
. "$PSScriptRoot\OutlookMapiInterop.ps1"
Assert-TestbedGuest -ExpectedUser $ExpectedUser
if ($env:COMPUTERNAME -notlike 'OAI-*' -or $env:COMPUTERNAME -eq 'OAI-BUILD') {
    throw "REFUSING TO RUN on '$env:COMPUTERNAME': this probe starts and ends Outlook, and runs only on an Outlook test guest (OAI-*, not the build VM)."
}

Add-ProbeTypes

function Get-OutlookPids { return @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }

# ------------------------------------------------------------------------------------------------
# Child modes. Each attaches, does its one thing, releases and exits - so a reading can never leave
# a reference behind in the process that drives the scenario.
# ------------------------------------------------------------------------------------------------
if ($Mode -ne 'Scenario') {
    $result = $null
    switch ($Mode) {
        'State' { $result = [OaiPinCom]::Read() }
        'Quit' { $result = @{ quit = [OaiPinCom]::Quit() } }
        'CloseHidden' { $result = @{ closed = [OaiPinCom]::CloseAllExplorers() } }
        'Hold' { $result = @{ hold = [OaiPinCom]::Hold($OutFile + '.ready', $ReleaseFlag, 30) } }
        'ShowActive' { $result = @{ show = [OaiPinCom]::ShowActiveExplorer() } }
    }
    $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutFile -Encoding UTF8
    exit 0
}

if (-not $Scenario) { throw '-Scenario is required: Control, ControlClose, Visible, Started, Sequential, Killed, KilledVisible, ShowMe or Headless.' }
if (-not (Test-Path -LiteralPath $ServerExe)) { throw "The staged server is not at $ServerExe." }
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$script:Stamp = (Get-Date).ToString('yyyyMMddTHHmmss')
$script:LogPath = Join-Path $LogDir "q118-$Scenario-$script:Stamp.jsonl"
$script:Clock = [System.Diagnostics.Stopwatch]::StartNew()
$script:ChildSeq = 0

function Write-Event([string] $Kind, [string] $Text, $Data) {
    $t = [Math]::Round($script:Clock.Elapsed.TotalSeconds, 1)
    $o = [ordered]@{ t = $t; utc = [DateTime]::UtcNow.ToString('o'); scenario = $Scenario; kind = $Kind; text = $Text; data = $Data }
    Add-Content -LiteralPath $script:LogPath -Value ($o | ConvertTo-Json -Compress -Depth 8) -Encoding UTF8
    # Write-Host, not Write-Output: a function that logs must not return its log lines. The task
    # wrapper captures every stream (*>&1), so out.txt still gets them.
    Write-Host ('[+{0,6:N1}s] {1,-9} {2}' -f $t, $Kind, $Text)
}

function Get-WindowSummary {
    $pids = Get-OutlookPids
    $wins = @([OaiPinNative]::Snapshot([int[]]$pids))
    $frames = @($wins | Where-Object { $_.Class -eq 'rctrl_renwnd32' })
    $visible = @($wins | Where-Object { $_.Visible })
    $dialogs = @($visible | Where-Object { $_.Class -eq '#32770' })
    return [ordered]@{
        pids          = $pids
        windows       = $wins.Count
        visible       = $visible.Count
        frames        = $frames.Count
        framesVisible = @($frames | Where-Object { $_.Visible }).Count
        dialogs       = $dialogs.Count
        visibleList   = @($visible | ForEach-Object { '{0} "{1}"{2}' -f $_.Class, ($(if ($_.Title.Length -gt 60) { $_.Title.Substring(0, 60) } else { $_.Title })), $(if ($_.Minimized) { ' (minimized)' } else { '' }) })
    }
}

function Get-OutlookProcessInfo {
    $out = @()
    foreach ($id in (Get-OutlookPids)) {
        $w = Get-CimInstance Win32_Process -Filter "ProcessId = $id"
        $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($w.ParentProcessId)" -ErrorAction SilentlyContinue
        $out += [ordered]@{ pid = $id; commandLine = $w.CommandLine; parent = $(if ($parent) { $parent.Name } else { "pid $($w.ParentProcessId) (gone)" }) }
    }
    return $out
}

# Runs this file again in a child mode, under a time limit, and returns its JSON result.
function Invoke-ChildProbe([string] $ChildMode, [int] $TimeoutSeconds = 90, [string[]] $Extra = @()) {
    $script:ChildSeq++
    $out = Join-Path $LogDir ("child-{0}-{1:000}-{2}.json" -f $script:Stamp, $script:ChildSeq, $ChildMode)
    $psi = New-Object System.Diagnostics.ProcessStartInfo (Join-Path $PSHOME 'powershell.exe')
    $argList = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath), '-Mode', $ChildMode, '-OutFile', ('"{0}"' -f $out)) + $Extra
    $psi.Arguments = $argList -join ' '
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $child = [System.Diagnostics.Process]::Start($psi)
    if (-not $child.WaitForExit($TimeoutSeconds * 1000)) {
        $name = (Get-Process -Id $child.Id -ErrorAction SilentlyContinue).ProcessName
        if ($name -eq 'powershell') { Stop-Process -Id $child.Id -Force }
        return [pscustomobject]@{ timedOut = $true; mode = $ChildMode }
    }
    if (-not (Test-Path -LiteralPath $out)) { return [pscustomobject]@{ failed = $true; mode = $ChildMode; exitCode = $child.ExitCode } }
    return (Get-Content -LiteralPath $out -Raw | ConvertFrom-Json)
}

# One reading: COM (in a child) and Win32. The line it prints is the measurement.
function Measure-PinState([string] $Label) {
    $pids = Get-OutlookPids
    $com = $null
    if ($pids.Count -gt 0) { $com = Invoke-ChildProbe -ChildMode State -TimeoutSeconds 90 }
    $win = Get-WindowSummary
    $pidsAfter = Get-OutlookPids
    $note = ''
    if ($pids.Count -eq 0 -and $pidsAfter.Count -gt 0) { $note = ' NOTE: an Outlook appeared during the reading' }
    if ($pids.Count -gt 0 -and $pidsAfter.Count -eq 0) { $note = ' NOTE: Outlook left during the reading' }
    $text = ''
    if ($pids.Count -eq 0) { $text = 'no OUTLOOK.EXE' }
    elseif ($null -eq $com -or $com.timedOut -or $com.failed -or -not $com.Attached) {
        $why = 'unknown'
        if ($null -ne $com) { if ($com.timedOut) { $why = 'timed out' } elseif ($com.failed) { $why = 'failed' } else { $why = $com.AttachError } }
        $text = "COM reading failed ($why); windows: frames $($win.framesVisible) visible / $($win.frames), dialogs $($win.dialogs)"
    }
    else {
        $hidden = @($com.Explorers | Where-Object { $_.Visible -eq 'hidden' }).Count
        $vis = @($com.Explorers | Where-Object { $_.Visible -eq 'visible' }).Count
        $unk = @($com.Explorers | Where-Object { $_.Visible -eq 'unknown' }).Count
        $active = $com.Active
        if ($com.Active -eq 'explorer') { $active = "Explorer #$($com.ActiveIndex), $($com.ActiveVisible)" }
        $text = "Explorers $($com.ExplorerCount) (visible $vis, hidden $hidden$(if ($unk) { ", unknown $unk" })); ActiveExplorer(): $active; Inspectors $($com.InspectorCount); Outbox default $($com.DefaultOutbox), all $($com.OutboxTotal) over $($com.OutboxStores) stores; frames $($win.framesVisible) visible / $($win.frames) rctrl_renwnd32; top-level windows $($win.visible) visible / $($win.windows); dialogs $($win.dialogs)"
        if ($com.Errors) { $text += "; errors: $($com.Errors)" }
    }
    Write-Event 'state' ("{0}: {1}{2}" -f $Label, $text, $note) ([ordered]@{ label = $Label; pids = $pids; com = $com; windows = $win; process = (Get-OutlookProcessInfo) })
    return $com
}

function Wait-OutlookExit([int] $Seconds) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $Seconds) {
        if ((Get-OutlookPids).Count -eq 0) { return [Math]::Round($sw.Elapsed.TotalSeconds, 1) }
        Start-Sleep -Milliseconds 250
    }
    return $null
}

function Start-OutlookVisible([switch] $AllowRunning, [string] $Label = 'user') {
    $before = Get-OutlookPids
    if ($before.Count -gt 0 -and -not $AllowRunning) {
        Write-Event 'refused' "Outlook is already running (pid $($before -join ',')); not starting another" $null
        return $false
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Start-Process -FilePath $OutlookExe -ArgumentList @('/profile', ('"{0}"' -f $Profile)) | Out-Null
    $up = $null
    while ($sw.Elapsed.TotalSeconds -lt 180) {
        $win = Get-WindowSummary
        if ($win.framesVisible -gt 0) { $up = [Math]::Round($sw.Elapsed.TotalSeconds, 1); break }
        Start-Sleep -Milliseconds 500
    }
    if ($null -eq $up) {
        Write-Event 'outlook' "$Label start: NO visible Outlook window within 180 s" (Get-WindowSummary)
        return $false
    }
    # Settled: 15 s, then until Outlook uses under 1 s of CPU in 5 s (at most 60 s more).
    Start-Sleep -Seconds 15
    $quietAfter = $null
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw2.Elapsed.TotalSeconds -lt 60) {
        $pid0 = (Get-OutlookPids | Select-Object -First 1)
        if (-not $pid0) { break }
        $c0 = (Get-Process -Id $pid0).TotalProcessorTime.TotalSeconds
        Start-Sleep -Seconds 5
        $p1 = Get-Process -Id $pid0 -ErrorAction SilentlyContinue
        if (-not $p1) { break }
        if (($p1.TotalProcessorTime.TotalSeconds - $c0) -lt 1.0) { $quietAfter = [Math]::Round($sw.Elapsed.TotalSeconds, 1); break }
    }
    Write-Event 'outlook' ("{0} start: window up after {1} s, settled at {2} s (pid {3})" -f $Label, $up, $(if ($quietAfter) { $quietAfter } else { 'not quiet in 60 s' }), ((Get-OutlookPids) -join ',')) ([ordered]@{ windowUp = $up; settled = $quietAfter; process = (Get-OutlookProcessInfo) })
    return $true
}

function Get-ComHostOf($Session) {
    return @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $($Session.Pid)" | Where-Object { $_.Name -eq 'OutlookAI.ComHost.exe' })
}

function Start-Session([string] $Name) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $s = New-Object OaiMcpSession($ServerExe, $Name)
    $init = $s.Await($s.Send('initialize', '{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"OutlookAI.Q118Probe","version":"0.1"}}'), 60000)
    if ($null -eq $init) { Write-Event 'session' "$Name : initialize got NO answer in 60 s" $null; return $s }
    $s.Notify('notifications/initialized')
    Write-Event 'session' ("{0}: server pid {1} initialized in {2:N0} ms" -f $Name, $s.Pid, $sw.Elapsed.TotalMilliseconds) ([ordered]@{ name = $Name; pid = $s.Pid })
    return $s
}

# list_folders on each session at once: every request is sent before any answer is awaited.
# -UntilReady repeats a call answered OutlookStarting (the server answers at once while its COM host
# goes on starting Outlook) every 2 s for up to 90 s, as the answer's own retry guidance says.
function Invoke-ListFolders([object[]] $Sessions, [switch] $UntilReady) {
    $sent = @()
    foreach ($s in $Sessions) { $sent += [pscustomobject]@{ s = $s; id = $s.Send('tools/call', '{"name":"list_folders","arguments":{}}'); sw = [System.Diagnostics.Stopwatch]::StartNew() } }
    foreach ($x in $sent) {
        $resp = $x.s.Await($x.id, 180000)
        $tries = 1
        while ($UntilReady -and $resp -and $resp -match 'OutlookStarting' -and $x.sw.Elapsed.TotalSeconds -lt 90) {
            Write-Event 'tool' ("{0}: list_folders answered OutlookStarting at {1} ms (try {2}); retrying in 2 s" -f $x.s.Name, [Math]::Round($x.sw.Elapsed.TotalMilliseconds), $tries) $null
            Start-Sleep -Seconds 2
            $tries++
            $resp = $x.s.Await($x.s.Send('tools/call', '{"name":"list_folders","arguments":{}}'), 180000)
        }
        $ms = [Math]::Round($x.sw.Elapsed.TotalMilliseconds)
        $comHost = Get-ComHostOf $x.s
        if ($null -eq $resp) { $verdict = 'NO ANSWER in 180 s' }
        elseif ($resp -match '"isError"\s*:\s*true') { $verdict = 'isError' }
        elseif ($resp -match '"error"\s*:\s*\{') { $verdict = 'JSON-RPC error' }
        else { $verdict = 'ok' }
        $snippet = ''
        if ($verdict -ne 'ok' -and $resp) { $snippet = $resp.Substring(0, [Math]::Min(300, $resp.Length)) }
        Write-Event 'tool' ("{0}: list_folders {1} in {2} ms, {3} bytes; COM host pid {4}" -f $x.s.Name, $verdict, $ms, $(if ($resp) { $resp.Length } else { 0 }), (($comHost | ForEach-Object { $_.ProcessId }) -join ',')) ([ordered]@{ name = $x.s.Name; verdict = $verdict; ms = $ms; snippet = $snippet })
    }
}

# The normal end: stdin closed on every session at once, then each awaited.
function Stop-SessionNormally([object[]] $Sessions) {
    $hosts = @{}
    foreach ($s in $Sessions) { $hosts[$s.Name] = @(Get-ComHostOf $s | ForEach-Object { $_.ProcessId }) }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($s in $Sessions) { $s.CloseInput() }
    foreach ($s in $Sessions) {
        $exited = $s.WaitExit(60000)
        $ms = [Math]::Round($sw.Elapsed.TotalMilliseconds)
        $left = @($hosts[$s.Name] | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue })
        Write-Event 'end' ("{0}: stdin closed; server {1} after {2} ms (exit code {3}); its COM host {4}" -f $s.Name, $(if ($exited) { 'exited' } else { 'STILL RUNNING' }), $ms, $s.ExitCode, $(if ($hosts[$s.Name].Count -eq 0) { 'was never seen' } elseif ($left.Count -eq 0) { 'gone' } else { 'STILL RUNNING' })) ([ordered]@{ name = $s.Name; exited = $exited; ms = $ms; exitCode = $s.ExitCode; stderrTail = $s.StderrTail(400) })
    }
}

# The COM host only - never Outlook, which is why the name is checked before the kill.
function Stop-ComHost($Session) {
    $targets = @(Get-ComHostOf $Session)
    foreach ($t in $targets) {
        $proc = Get-Process -Id $t.ProcessId -ErrorAction SilentlyContinue
        if ($proc -and $proc.ProcessName -eq 'OutlookAI.ComHost') { Stop-Process -Id $proc.Id -Force }
    }
    Start-Sleep -Milliseconds 500
    $left = @($targets | Where-Object { Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue })
    Write-Event 'kill' ("{0}: COM host {1} terminated ({2})" -f $Session.Name, (($targets | ForEach-Object { $_.ProcessId }) -join ','), $(if ($targets.Count -eq 0) { 'NONE FOUND' } elseif ($left.Count -eq 0) { 'gone' } else { 'STILL RUNNING' })) $null
}

function Invoke-UiaElement($Element) {
    $pattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref] $pattern)) { $pattern.Invoke(); return 'Invoke' }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref] $pattern)) { $pattern.Select(); return 'SelectionItem.Select' }
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref] $pattern)) { $pattern.Expand(); return 'ExpandCollapse.Expand' }
    return $null
}

function Find-UiaByName($Root, [string[]] $Names) {
    foreach ($n in $Names) {
        $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $n)
        $el = $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
        if ($el) { return $el }
    }
    return $null
}

# A user's File > Exit, through UI Automation on Outlook's own window.
function Invoke-UiFileExit {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $pids = Get-OutlookPids
    $frame = @([OaiPinNative]::Snapshot([int[]]$pids) | Where-Object { $_.Class -eq 'rctrl_renwnd32' -and $_.Visible -and -not $_.Minimized }) | Select-Object -First 1
    if (-not $frame) { return 'no visible Outlook window to drive' }
    $AE = [System.Windows.Automation.AutomationElement]
    $root = $AE::FromHandle([IntPtr]$frame.Handle)
    $file = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, 'FileTabButton')))
    if (-not $file) { $file = Find-UiaByName $root @('File Tab', 'File') }
    if (-not $file) { return 'the File tab was not found' }
    $how1 = Invoke-UiaElement $file
    $exit = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $exit -and $sw.Elapsed.TotalSeconds -lt 20) {
        Start-Sleep -Milliseconds 500
        foreach ($w in @([OaiPinNative]::Snapshot([int[]](Get-OutlookPids)) | Where-Object { $_.Visible })) {
            $exit = Find-UiaByName ($AE::FromHandle([IntPtr]$w.Handle)) @('Exit')
            if ($exit) { break }
        }
    }
    if (-not $exit) {
        $names = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition) | Select-Object -First 400 | ForEach-Object { $_.Current.ControlType.ProgrammaticName + ' "' + $_.Current.Name + '" id=' + $_.Current.AutomationId } | Where-Object { $_ -notmatch '"" id=$' })
        Write-Event 'uia' 'backstage dump (first 400 named elements)' ([ordered]@{ elements = $names })
        return "File invoked ($how1) but no 'Exit' element appeared in 20 s"
    }
    $how2 = $null
    try { $how2 = Invoke-UiaElement $exit } catch { $how2 = 'invoked, then ' + $_.Exception.GetType().Name }
    return ("File ({0}) then Exit ({1}), Exit was {2} id={3}" -f $how1, $how2, $exit.Current.ControlType.ProgrammaticName, $exit.Current.AutomationId)
}

# A user's click on the close button of Outlook's main window: UI Automation's WindowPattern.Close.
function Invoke-UiWindowClose {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $frames = @([OaiPinNative]::Snapshot([int[]](Get-OutlookPids)) | Where-Object { $_.Class -eq 'rctrl_renwnd32' -and $_.Visible })
    if ($frames.Count -ne 1) { return "expected one visible Outlook window, found $($frames.Count); nothing closed" }
    $el = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$frames[0].Handle)
    $pattern = $null
    if (-not $el.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref] $pattern)) { return 'the window offers no WindowPattern; nothing closed' }
    try { $pattern.Close() } catch { return 'WindowPattern.Close, then ' + $_.Exception.GetType().Name }
    return 'WindowPattern.Close on the one visible Outlook window'
}

# Ends Outlook the graceful way named, under rule 7, and times it. Leaves it running if it stays.
function Invoke-OutlookEnd([ValidateSet('Quit', 'FileExit', 'WindowClose', 'CloseHidden')] [string] $How, [string] $Label) {
    if ((Get-OutlookPids).Count -eq 0) { Write-Event 'end-ol' "${Label}: $How skipped - Outlook is not running" $null; return $true }
    $com = Invoke-ChildProbe -ChildMode State -TimeoutSeconds 90
    $win = Get-WindowSummary
    $attached = ($null -ne $com) -and -not $com.timedOut -and -not $com.failed -and $com.Attached
    $refusal = $null
    if ($attached) { $refusal = Get-Rule7Refusal -InspectorCount $com.InspectorCount -DefaultOutbox $com.DefaultOutbox -OutboxTotal $com.OutboxTotal -VisibleDialogs $win.dialogs }
    else { $refusal = Get-Rule7Refusal -InspectorCount 0 -DefaultOutbox 0 -OutboxTotal 0 -VisibleDialogs 0 -Attached $false }
    if (-not $refusal -and $How -eq 'CloseHidden' -and $win.framesVisible -gt 0) { $refusal = "an Outlook window is visible ($($win.framesVisible)) - closing Explorers is only for an Outlook nobody can see" }
    if ($refusal) { Write-Event 'refused' "${Label}: $How REFUSED under rule 7: $refusal" $null; return $false }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $action = ''
    switch ($How) {
        'Quit' { $r = Invoke-ChildProbe -ChildMode Quit -TimeoutSeconds 90; $action = "Application.Quit(): $(if ($r.timedOut) { 'the call did not return in 90 s' } else { $r.quit })" }
        'FileExit' { $action = 'File > Exit by UI Automation: ' + (Invoke-UiFileExit) }
        'WindowClose' { $action = 'window close by UI Automation: ' + (Invoke-UiWindowClose) }
        'CloseHidden' { $r = Invoke-ChildProbe -ChildMode CloseHidden -TimeoutSeconds 90; $action = "closed $($r.closed) hidden Explorer(s)" }
    }
    $actionMs = [Math]::Round($sw.Elapsed.TotalMilliseconds)
    $gone = Wait-OutlookExit $ExitWaitSeconds
    if ($null -ne $gone) {
        Write-Event 'end-ol' ("{0}: {1} -> OUTLOOK.EXE EXITED {2} s after the action returned ({3} ms to return)" -f $Label, $action, $gone, $actionMs) ([ordered]@{ how = $How; exited = $true; seconds = $gone; actionMs = $actionMs })
        return $true
    }
    Write-Event 'end-ol' ("{0}: {1} -> OUTLOOK.EXE STILL RUNNING {2} s after the action ({3} ms to return)" -f $Label, $action, $ExitWaitSeconds, $actionMs) ([ordered]@{ how = $How; exited = $false; seconds = $ExitWaitSeconds; actionMs = $actionMs; windows = (Get-WindowSummary) })
    Measure-PinState "${Label}: after $How, Outlook still up (this reading attaches)" | Out-Null
    return $false
}

# Last resort for a scenario that left Outlook up: close the hidden Explorers, gracefully.
function Close-Leftover([string] $Label) {
    if ((Get-OutlookPids).Count -eq 0) { return }
    $win = Get-WindowSummary
    if ($win.framesVisible -gt 0) { Write-Event 'leftover' "${Label}: Outlook left running WITH a visible window; nothing closed" $null; return }
    Invoke-OutlookEnd -How CloseHidden -Label "$Label leftover" | Out-Null
    if ((Get-OutlookPids).Count -gt 0) { Write-Event 'leftover' "${Label}: OUTLOOK.EXE is STILL RUNNING; left as it is (rule 7) - restore the checkpoint" $null }
}

function Assert-NoOutlook([string] $Label) {
    if ((Get-OutlookPids).Count -gt 0) { Write-Event 'refused' "${Label}: needs NO Outlook running and one is (pid $((Get-OutlookPids) -join ',')); scenario not run" $null; return $false }
    return $true
}

# ------------------------------------------------------------------------------------------------
# Scenarios.
# ------------------------------------------------------------------------------------------------
function Invoke-ScenarioControl {
    if (-not (Start-OutlookVisible)) { return }
    Measure-PinState 'control: baseline, no OutlookAI session ever attached' | Out-Null
    if (-not (Invoke-OutlookEnd -How Quit -Label 'control')) { Close-Leftover 'control'; return }
    if (-not (Start-OutlookVisible)) { return }
    Measure-PinState 'control: second baseline' | Out-Null
    if (-not (Invoke-OutlookEnd -How FileExit -Label 'control')) { Close-Leftover 'control' }
}

function Invoke-ScenarioVisible {
    if (-not (Start-OutlookVisible)) { return }
    Measure-PinState 'visible: baseline' | Out-Null
    foreach ($i in 1..3) {
        $s = Start-Session "seq$i"; Invoke-ListFolders @($s); Stop-SessionNormally @($s)
        Measure-PinState "visible: after sequential session $i ended normally" | Out-Null
    }
    $many = @(Start-Session 'con1'; Start-Session 'con2'; Start-Session 'con3')
    Invoke-ListFolders $many
    Measure-PinState 'visible: three sessions alive at once' | Out-Null
    Stop-SessionNormally $many
    Measure-PinState 'visible: after the three ended normally' | Out-Null
    $k = Start-Session 'kill1'; Invoke-ListFolders @($k); Stop-ComHost $k
    Measure-PinState 'visible: after kill1 lost its COM host' | Out-Null
    Stop-SessionNormally @($k)
    Measure-PinState 'visible: after kill1 server ended' | Out-Null
    if (-not (Invoke-OutlookEnd -How Quit -Label 'visible')) { Close-Leftover 'visible'; return }

    if (-not (Start-OutlookVisible)) { return }
    Measure-PinState 'visible2: baseline' | Out-Null
    $s = Start-Session 'seq4'; Invoke-ListFolders @($s); Stop-SessionNormally @($s)
    $two = @(Start-Session 'con4'; Start-Session 'con5'); Invoke-ListFolders $two; Stop-SessionNormally $two
    $k = Start-Session 'kill2'; Invoke-ListFolders @($k); Stop-ComHost $k; Stop-SessionNormally @($k)
    Measure-PinState 'visible2: after one sequential, two at once and one killed' | Out-Null
    if (-not (Invoke-OutlookEnd -How FileExit -Label 'visible2')) { Close-Leftover 'visible2' }
}

function Invoke-ScenarioStarted {
    if (-not (Assert-NoOutlook 'started')) { return }
    $a = Start-Session 'A'; Invoke-ListFolders @($a)
    Measure-PinState 'started: A started Outlook (A alive)' | Out-Null
    $b = Start-Session 'B'; Invoke-ListFolders @($b)
    Measure-PinState 'started: B attached (A and B alive)' | Out-Null
    Stop-SessionNormally @($b)
    Measure-PinState 'started: B ended normally (A alive)' | Out-Null
    Stop-SessionNormally @($a)
    Measure-PinState 'started: A, the session that started Outlook, ended normally' | Out-Null
    $gone = Wait-OutlookExit 45
    if ($null -ne $gone) { Write-Event 'end-ol' "started: OUTLOOK.EXE exited $gone s after that reading, with no session attached" $null; return }
    Write-Event 'note' 'started: Outlook is still up 45 s after its starter ended - the lingering headless Outlook' $null
    $many = @(Start-Session 'C1'; Start-Session 'C2'; Start-Session 'C3')
    Invoke-ListFolders $many
    Measure-PinState 'started: C1-C3 alive at once on the lingering Outlook' | Out-Null
    Stop-SessionNormally $many
    Measure-PinState 'started: C1-C3 ended normally' | Out-Null
    foreach ($n in @('D', 'E')) {
        $s = Start-Session $n; Invoke-ListFolders @($s); Stop-SessionNormally @($s)
        Measure-PinState "started: $n ended normally" | Out-Null
    }
    if (-not (Invoke-OutlookEnd -How Quit -Label 'started')) { Close-Leftover 'started' }
}

function Invoke-ScenarioSequential {
    if (-not (Assert-NoOutlook 'sequential')) { return }
    $a = Start-Session 'A'; Invoke-ListFolders @($a) -UntilReady
    Measure-PinState 'sequential: A started Outlook and is ready (A alive)' | Out-Null
    Stop-SessionNormally @($a)
    Measure-PinState 'sequential: A, the session that started Outlook, ended normally' | Out-Null
    $gone = Wait-OutlookExit 20
    if ($null -ne $gone) { Write-Event 'end-ol' "sequential: OUTLOOK.EXE exited $gone s after A's end was read, with no session attached" $null; return }
    Write-Event 'note' 'sequential: Outlook is still up 20 s after its starter ended' $null
    $b = Start-Session 'B'; Invoke-ListFolders @($b)
    Measure-PinState 'sequential: B attached to the lingering Outlook (B alive)' | Out-Null
    Stop-SessionNormally @($b)
    Measure-PinState 'sequential: B ended normally' | Out-Null
    $c = Start-Session 'C'; Invoke-ListFolders @($c); Stop-SessionNormally @($c)
    Measure-PinState 'sequential: C ended normally' | Out-Null
    if (-not (Invoke-OutlookEnd -How Quit -Label 'sequential')) { Close-Leftover 'sequential' }
}

function Invoke-KilledStart([string] $Label) {
    $a = Start-Session 'A'; Invoke-ListFolders @($a) -UntilReady
    Measure-PinState "${Label}: A started Outlook (A alive)" | Out-Null
    Stop-ComHost $a
    Measure-PinState "${Label}: A's COM host killed" | Out-Null
    Stop-SessionNormally @($a)
    Start-Sleep -Seconds 5
    Measure-PinState "${Label}: A's server ended, 5 s on" | Out-Null
}

function Invoke-ScenarioKilled {
    if (-not (Assert-NoOutlook 'killed')) { return }
    Invoke-KilledStart 'killed'
    if ((Get-OutlookPids).Count -eq 0) { Write-Event 'note' 'killed: Outlook did not survive its starter' $null; return }
    $s = Start-Session 'B'; Invoke-ListFolders @($s); Stop-SessionNormally @($s)
    Measure-PinState 'killed: B ended normally' | Out-Null
    $two = @(Start-Session 'C'; Start-Session 'D'); Invoke-ListFolders $two
    Measure-PinState 'killed: C and D alive at once' | Out-Null
    Stop-SessionNormally $two
    Measure-PinState 'killed: C and D ended normally' | Out-Null
    if (-not (Invoke-OutlookEnd -How Quit -Label 'killed')) { Close-Leftover 'killed' }
}

function Invoke-ScenarioKilledVisible {
    # The leaked pin, then the user's own window over it, ended three ways: File > Exit, then (after
    # the leak is made again) Application.Quit(), then the window's close button.
    foreach ($how in @('FileExit', 'Quit', 'WindowClose')) {
        $label = "killedvisible-$how"
        if (-not (Assert-NoOutlook $label)) { return }
        Invoke-KilledStart $label
        if ((Get-OutlookPids).Count -eq 0) { Write-Event 'note' "${label}: Outlook did not survive its starter" $null; return }
        if (-not (Start-OutlookVisible -AllowRunning -Label 'user opens the window of the running Outlook')) { Close-Leftover $label; return }
        Measure-PinState "${label}: the user has the window open" | Out-Null
        if (-not (Invoke-OutlookEnd -How $how -Label $label)) { Close-Leftover $label }
        if ((Get-OutlookPids).Count -gt 0) { return }
    }
}

function Invoke-ScenarioShowMe {
    # S6 to its end: the leaked pin, the user's window opened and closed over it, then the two calls
    # the show-me path makes on what ActiveExplorer() returns, then the user's close of what that showed.
    $label = 'showme'
    if (-not (Assert-NoOutlook $label)) { return }
    Invoke-KilledStart $label
    if ((Get-OutlookPids).Count -eq 0) { Write-Event 'note' "${label}: Outlook did not survive its starter" $null; return }
    if (-not (Start-OutlookVisible -AllowRunning -Label 'user opens the window of the running Outlook')) { Close-Leftover $label; return }
    Measure-PinState "${label}: the user has the window open" | Out-Null
    if (Invoke-OutlookEnd -How WindowClose -Label "$label user close") { return }
    $r = Invoke-ChildProbe -ChildMode ShowActive -TimeoutSeconds 60
    Write-Event 'showme' ("{0}: EnsureVisibleExplorer's ActiveExplorer branch, emulated: {1}" -f $label, $(if ($r.timedOut) { 'timed out' } else { $r.show })) $null
    Measure-PinState "${label}: after the emulated show-me call" | Out-Null
    $win = Get-WindowSummary
    if ($win.framesVisible -eq 1) {
        if (-not (Invoke-OutlookEnd -How WindowClose -Label "$label close of the shown pin")) { Close-Leftover $label }
    }
    else { Close-Leftover $label }
}

function Invoke-ScenarioControlClose {
    if (-not (Start-OutlookVisible)) { return }
    Measure-PinState 'controlclose: baseline, no OutlookAI session in this Outlook' | Out-Null
    if (-not (Invoke-OutlookEnd -How WindowClose -Label 'controlclose')) { Close-Leftover 'controlclose' }
}

function Invoke-ScenarioHeadless {
    if (-not (Assert-NoOutlook 'headless')) { return }
    $flag = Join-Path $LogDir "release-$script:Stamp.flag"
    $holdOut = Join-Path $LogDir "hold-$script:Stamp.json"
    $psi = New-Object System.Diagnostics.ProcessStartInfo (Join-Path $PSHOME 'powershell.exe')
    $psi.Arguments = ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode Hold -OutFile "{1}" -ReleaseFlag "{2}"' -f $PSCommandPath, $holdOut, $flag)
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $holder = [System.Diagnostics.Process]::Start($psi)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path -LiteralPath ($holdOut + '.ready')) -and $sw.Elapsed.TotalSeconds -lt 180 -and -not $holder.HasExited) { Start-Sleep -Milliseconds 500 }
    Write-Event 'holder' ("headless: plain COM client pid {0} holds Outlook ({1} s)" -f $holder.Id, [Math]::Round($sw.Elapsed.TotalSeconds, 1)) $null
    Start-Sleep -Seconds 20
    Measure-PinState 'headless: Outlook started by a plain COM client, no OutlookAI session yet' | Out-Null
    $many = @(Start-Session 'H1'; Start-Session 'H2'; Start-Session 'H3')
    Invoke-ListFolders $many
    Measure-PinState 'headless: H1-H3 alive at once' | Out-Null
    Stop-SessionNormally $many
    Measure-PinState 'headless: H1-H3 ended normally' | Out-Null
    $s = Start-Session 'H4'; Invoke-ListFolders @($s); Stop-SessionNormally @($s)
    Measure-PinState 'headless: H4 ended normally' | Out-Null
    Set-Content -LiteralPath $flag -Value 'release'
    $released = $holder.WaitForExit(60000)
    Write-Event 'holder' ("headless: holder {0}" -f $(if ($released) { 'let go and exited' } else { 'DID NOT EXIT in 60 s' })) $null
    Start-Sleep -Seconds 10
    Measure-PinState 'headless: holder gone, 10 s on' | Out-Null
    if (-not (Invoke-OutlookEnd -How Quit -Label 'headless')) { Close-Leftover 'headless' }
}

Write-Event 'begin' ("scenario {0} on {1} as {2}; server {3} ({4}); Outlook {5}" -f $Scenario, $env:COMPUTERNAME, $env:USERNAME, $ServerExe, (Get-Item -LiteralPath $ServerExe).VersionInfo.ProductVersion, (Get-Item -LiteralPath $OutlookExe).VersionInfo.ProductVersion) $null
switch ($Scenario) {
    'Control' { Invoke-ScenarioControl }
    'Visible' { Invoke-ScenarioVisible }
    'Started' { Invoke-ScenarioStarted }
    'Sequential' { Invoke-ScenarioSequential }
    'Killed' { Invoke-ScenarioKilled }
    'KilledVisible' { Invoke-ScenarioKilledVisible }
    'ControlClose' { Invoke-ScenarioControlClose }
    'ShowMe' { Invoke-ScenarioShowMe }
    'Headless' { Invoke-ScenarioHeadless }
}
Write-Event 'done' ("scenario {0}; OUTLOOK.EXE {1}" -f $Scenario, $(if ((Get-OutlookPids).Count -gt 0) { 'running (pid ' + ((Get-OutlookPids) -join ',') + ')' } else { 'not running' })) $null
Write-Output "log: $script:LogPath"
