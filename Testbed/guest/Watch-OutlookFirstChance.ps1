#Requires -Version 5.1
<#
.SYNOPSIS
    Hunts one crash of OUTLOOK.EXE with a debugger: attached to every OUTLOOK.EXE of this Windows session, it writes
    ONE full dump at the first first-chance access violation in one module - mso40uiwin32client.dll by default - and
    one at any second-chance exception, and passes every exception back to Outlook. Windows' own debugging API through
    Add-Type; nothing installed.

.DESCRIPTION
    RUN ON A TEST GUEST, in session 1, NOT elevated - as Outlook runs. Windows PowerShell 5.1 - no ternary, no `??`.

    WHY (2026-11-18). Two crashes of OUTLOOK.EXE in mso40uiwin32client.dll that day left an Application Error event
    and no dump, though Windows Error Reporting's LocalDumps keys were armed; on a test guest every crash path that
    writes that event wrote the dump, so these two lost it inside WER (Docs/live-tier-on-the-vm.md section 4.15). A
    debugger is told of an access violation before any handler runs - Outlook's, Office's or WER's - so a dump it
    writes then cannot be lost that way. The coordinator decided the hunt that day: an opt-in run, never an ordinary
    one, because a debugger changes what it is attached to (IsDebuggerPresent answers yes, and every first-chance
    exception, thread and module load stops Outlook for a moment).

    ONLY IN A HUNTING RUN. Testbed/host/Invoke-LiveTierOnGuest.ps1 -HuntFirstChance starts it, in the suite's own task,
    and stops it after the suite through -StopFile. No other run starts it.

    WHAT IT DOES TO OUTLOOK, AND WHAT IT NEVER DOES.
      - Attaches to each OUTLOOK.EXE of its own session as it appears (DebugActiveProcess, polled four times a second),
        with DebugSetProcessKillOnExit(FALSE): if this script ends for any reason, Windows detaches it and Outlook goes
        on. It never kills, suspends or writes to Outlook.
      - Continues the attach's own breakpoint as handled, as every debugger does, and passes EVERY other exception back
        to Outlook unhandled, so Outlook's own handlers run as they would.
      - At the first first-chance access violation whose faulting address lies in -Module, writes a full dump of the
        process at that instant - OUTLOOK.EXE.<pid>.firstchance.dmp in -DumpFolder, with the exception record and the
        faulting thread's context - then passes the exception on. At the first second-chance exception of any kind -
        Outlook about to die of it, which with a debugger attached Windows Error Reporting does not see - writes
        OUTLOOK.EXE.<pid>.secondchance.dmp. One of each per process at most.
      - Detaches at -StopFile, after -MaxMinutes, or when Outlook exits, and logs per process what it saw: the events,
        the exceptions by code, the dumps, the exit code.
    A full dump holds Outlook's memory, mail included: like Set-OutlookCrashDumps.ps1, it refuses to run anywhere but a
    test guest, and never on the Exchange VM, whose Outlook holds a real mailbox.

    MODES:
      (default) The watch: -DumpFolder, -StopFile, -Log, -Module, -MaxMinutes.
      -SelfTest Pure: compiles the debugger and checks its decisions - which exception is dumped, how each is continued,
                the dump names, the event layout - and that nothing in it ends a process. No process is attached.

    EXIT CODES: 0 the watch ended (stop file, time limit); 1 it could not start; REFUSING TO RUN off a test guest (it
    throws).

.EXAMPLE
    # What Invoke-LiveTierOnGuest.ps1 -HuntFirstChance starts in the suite's task (session 1, NOT elevated):
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'C:\OutlookAI-Q5\Watch-OutlookFirstChance.ps1', '-StopFile', 'C:\OutlookAI-Q5\live-runs\<run>\first-chance.stop', '-Log', 'C:\OutlookAI-Q5\live-runs\<run>\first-chance.log'
    # Anywhere, pure:
    powershell.exe -NoProfile -File Testbed\guest\Watch-OutlookFirstChance.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [string]   $DumpFolder = 'C:\OutlookAI-Q5\crash-dumps',
    [string]   $StopFile = '',
    [string]   $Log = '',
    [string]   $Module = 'mso40uiwin32client.dll',
    [int]      $MaxMinutes = 120,
    [switch]   $SelfTest,
    [string[]] $ExpectedUser = @('vmadmin'),
    [string]   $ExpectedComputerNamePrefix = 'OAI-'
)

$ErrorActionPreference = 'Stop'

# The debugger. One thread does everything - attach, wait, dump, continue, detach - because Windows ties a debuggee to
# the thread that attached it. C# 5: Windows PowerShell 5.1 compiles it with the .NET Framework's own csc.
$DebuggerSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class OutlookFirstChance
{
    public const uint AccessViolation = 0xC0000005;
    public const uint Breakpoint = 0x80000003;
    public const uint DbgContinue = 0x00010002;
    public const uint DbgExceptionNotHandled = 0x80010001;

    // DEBUG_EVENT on x64: the code, the process and thread ids, then the union at 16. For an exception the union is
    // EXCEPTION_DEBUG_INFO: the EXCEPTION_RECORD (152 bytes: code at 0, address at 16, parameters at 32) and
    // dwFirstChance at 152. CREATE_PROCESS and LOAD_DLL hold a file handle at 16 that the debugger must close;
    // EXIT_PROCESS holds the exit code at 16.
    public const int EventSize = 176;
    public const int UnionOffset = 16;
    public const int RecordSize = 152;
    public const int AddressOffset = UnionOffset + 16;
    public const int ParametersOffset = UnionOffset + 32;
    public const int FirstChanceOffset = UnionOffset + RecordSize;

    // Full memory, handles, unloaded modules, the memory map and thread times: what host\Read-CrashDump.ps1 reads.
    public const uint FullDump = 0x2 | 0x4 | 0x20 | 0x800 | 0x1000;

    // PURE. The exception the hunt dumps: an access violation, at first chance, whose faulting address lies in the module.
    public static bool IsHunted(uint code, bool firstChance, string faultingModule, string module)
    {
        return code == AccessViolation && firstChance && !string.IsNullOrEmpty(faultingModule) && !string.IsNullOrEmpty(module)
            && string.Equals(faultingModule, module, StringComparison.OrdinalIgnoreCase);
    }

    // PURE. How an exception event is continued: the attach's own breakpoint as handled, every other exception back
    // to the process unhandled - its own handlers decide, as they would with no debugger.
    public static uint ContinueStatus(uint code, bool firstChance, bool attachBreakpointPending)
    {
        if (firstChance && code == Breakpoint && attachBreakpointPending) { return DbgContinue; }
        return DbgExceptionNotHandled;
    }

    // PURE. A dump's file name: what LocalDumps would name it, with the moment between - so the runner fetches it with
    // the rest and never counts it as a LocalDumps crash dump.
    public static string DumpName(int pid, bool firstChance)
    {
        return "OUTLOOK.EXE." + pid.ToString() + (firstChance ? ".firstchance.dmp" : ".secondchance.dmp");
    }

    [DllImport("kernel32.dll", SetLastError = true)] static extern bool DebugActiveProcess(int pid);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool DebugActiveProcessStop(int pid);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool DebugSetProcessKillOnExit(bool killOnExit);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool WaitForDebugEvent(IntPtr debugEvent, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool ContinueDebugEvent(int pid, int tid, uint status);
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenThread(uint access, bool inherit, int tid);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool GetThreadContext(IntPtr thread, IntPtr context);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool K32EnumProcessModulesEx(IntPtr process, [Out] IntPtr[] modules, int cb, out int needed, uint filter);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool K32GetModuleInformation(IntPtr process, IntPtr module, out ModuleInfo info, int cb);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] static extern int K32GetModuleBaseNameW(IntPtr process, IntPtr module, StringBuilder name, int size);
    [DllImport("dbghelp.dll", SetLastError = true)] static extern bool MiniDumpWriteDump(IntPtr process, int pid, Microsoft.Win32.SafeHandles.SafeFileHandle file, uint type, ref MinidumpExceptionInformation exception, IntPtr userStreams, IntPtr callback);

    [StructLayout(LayoutKind.Sequential)]
    struct ModuleInfo { public IntPtr Base; public int Size; public IntPtr Entry; }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    struct MinidumpExceptionInformation { public int ThreadId; public IntPtr ExceptionPointers; public int ClientPointers; }

    sealed class Debuggee
    {
        public int Pid;
        public IntPtr Process;
        public bool AttachBreakpointPending = true;
        public bool FirstChanceDumped;
        public bool SecondChanceDumped;
        public long HuntBase;
        public long HuntEnd;
        public long Events;
        public readonly Dictionary<uint, int> Exceptions = new Dictionary<uint, int>();
    }

    static string logPath = "";

    static void Say(string line)
    {
        string text = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") + "  " + line;
        Console.WriteLine(text);
        if (logPath.Length > 0) { try { File.AppendAllText(logPath, text + Environment.NewLine); } catch (IOException) { } }
    }

    // The module whose image holds the address, by name; the hunted module's range is kept once seen.
    static string ModuleOf(Debuggee d, long address, string module)
    {
        if (d.HuntEnd > d.HuntBase && address >= d.HuntBase && address < d.HuntEnd) { return module; }
        IntPtr[] modules = new IntPtr[2048];
        int needed;
        if (!K32EnumProcessModulesEx(d.Process, modules, modules.Length * IntPtr.Size, out needed, 0x3)) { return null; }
        string found = null;
        int count = Math.Min(modules.Length, needed / IntPtr.Size);
        for (int i = 0; i < count; i++)
        {
            ModuleInfo info;
            if (!K32GetModuleInformation(d.Process, modules[i], out info, Marshal.SizeOf(typeof(ModuleInfo)))) { continue; }
            long start = info.Base.ToInt64();
            long end = start + info.Size;
            bool holds = address >= start && address < end;
            StringBuilder name = new StringBuilder(260);
            if (holds || d.HuntEnd <= d.HuntBase)
            {
                if (K32GetModuleBaseNameW(d.Process, modules[i], name, name.Capacity) <= 0) { continue; }
                if (string.Equals(name.ToString(), module, StringComparison.OrdinalIgnoreCase)) { d.HuntBase = start; d.HuntEnd = end; }
                if (holds) { found = name.ToString(); }
            }
        }

        return found;
    }

    static void WriteDump(Debuggee d, int tid, IntPtr debugEvent, bool firstChance, string folder)
    {
        string path = Path.Combine(folder, DumpName(d.Pid, firstChance));
        IntPtr record = Marshal.AllocHGlobal(RecordSize);
        IntPtr contextRaw = Marshal.AllocHGlobal(1232 + 16);
        IntPtr pointers = Marshal.AllocHGlobal(16);
        IntPtr thread = IntPtr.Zero;
        try
        {
            byte[] copy = new byte[RecordSize];
            Marshal.Copy(new IntPtr(debugEvent.ToInt64() + UnionOffset), copy, 0, RecordSize);
            Marshal.Copy(copy, 0, record, RecordSize);
            Marshal.WriteIntPtr(record, 8, IntPtr.Zero);
            IntPtr context = new IntPtr((contextRaw.ToInt64() + 15) & ~15L);
            Marshal.Copy(new byte[1232], 0, context, 1232);
            Marshal.WriteInt32(context, 0x30, 0x0010001F);
            thread = OpenThread(0x0008 | 0x0040, false, tid);
            bool haveContext = thread != IntPtr.Zero && GetThreadContext(thread, context);
            Marshal.WriteIntPtr(pointers, 0, record);
            Marshal.WriteIntPtr(pointers, IntPtr.Size, haveContext ? context : IntPtr.Zero);
            MinidumpExceptionInformation exception = new MinidumpExceptionInformation();
            exception.ThreadId = tid;
            exception.ExceptionPointers = pointers;
            exception.ClientPointers = 0;
            Stopwatch clock = Stopwatch.StartNew();
            bool written;
            using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                written = MiniDumpWriteDump(d.Process, d.Pid, file.SafeFileHandle, FullDump, ref exception, IntPtr.Zero, IntPtr.Zero);
            }

            int error = written ? 0 : Marshal.GetLastWin32Error();
            long bytes = File.Exists(path) ? new FileInfo(path).Length : 0;
            Say("pid " + d.Pid + ": " + (written ? "DUMP " : "DUMP FAILED (error " + error + ") ") + Path.GetFileName(path) + ", " + bytes + " bytes, "
                + clock.ElapsedMilliseconds + " ms, thread " + tid + ", context " + (haveContext ? "read" : "NOT read"));
        }
        catch (Exception ex)
        {
            Say("pid " + d.Pid + ": DUMP FAILED - " + ex.GetType().Name + ": " + ex.Message);
        }
        finally
        {
            if (thread != IntPtr.Zero) { CloseHandle(thread); }
            Marshal.FreeHGlobal(pointers);
            Marshal.FreeHGlobal(contextRaw);
            Marshal.FreeHGlobal(record);
        }
    }

    static string Summary(Debuggee d)
    {
        StringBuilder text = new StringBuilder();
        foreach (KeyValuePair<uint, int> pair in d.Exceptions) { text.Append(" 0x" + pair.Key.ToString("X8") + "x" + pair.Value); }
        return d.Events + " debug event(s); exceptions by code:" + (text.Length == 0 ? " none" : text.ToString());
    }

    public static int Run(string module, string folder, string stopFile, string log, int maxMinutes)
    {
        logPath = log ?? "";
        int session = Process.GetCurrentProcess().SessionId;
        Dictionary<int, Debuggee> debuggees = new Dictionary<int, Debuggee>();
        HashSet<int> refused = new HashSet<int>();
        // The time limit and the look interval on a monotonic clock: a frozen test guest's clock can jump to the real
        // date in the middle of a hunt - 46 days on two unindexed guests on 2026-11-18 - and a deadline on it ended the
        // watch there.
        Stopwatch watched = Stopwatch.StartNew();
        TimeSpan limit = TimeSpan.FromMinutes(maxMinutes);
        TimeSpan nextLook = TimeSpan.Zero;
        IntPtr debugEvent = Marshal.AllocHGlobal(256);
        Say("watching session " + session + " for OUTLOOK.EXE; hunting first-chance access violations in " + module + "; dumps to " + folder);
        try
        {
            while (true)
            {
                if (watched.Elapsed >= nextLook)
                {
                    nextLook = watched.Elapsed + TimeSpan.FromMilliseconds(250);
                    if (stopFile.Length > 0 && File.Exists(stopFile)) { Say("stop file seen"); break; }
                    if (watched.Elapsed >= limit) { Say("time limit of " + maxMinutes + " minute(s) reached"); break; }
                    foreach (Process p in Process.GetProcessesByName("OUTLOOK"))
                    {
                        using (p)
                        {
                            int inSession;
                            try { inSession = p.SessionId; } catch (InvalidOperationException) { continue; }
                            if (inSession != session || debuggees.ContainsKey(p.Id) || refused.Contains(p.Id)) { continue; }
                            if (!DebugActiveProcess(p.Id))
                            {
                                refused.Add(p.Id);
                                Say("pid " + p.Id + ": could not attach (error " + Marshal.GetLastWin32Error() + ")");
                                continue;
                            }

                            // Never kill on exit: should this process end, Windows detaches it and Outlook goes on.
                            DebugSetProcessKillOnExit(false);
                            Debuggee d = new Debuggee();
                            d.Pid = p.Id;
                            d.Process = OpenProcess(0x0400 | 0x0010 | 0x0040, false, p.Id);
                            debuggees[p.Id] = d;
                            Say("pid " + p.Id + ": attached");
                        }
                    }
                }

                if (debuggees.Count == 0) { Thread.Sleep(100); continue; }
                if (!WaitForDebugEvent(debugEvent, 100)) { continue; }
                uint kind = (uint)Marshal.ReadInt32(debugEvent, 0);
                int pid = Marshal.ReadInt32(debugEvent, 4);
                int tid = Marshal.ReadInt32(debugEvent, 8);
                uint status = DbgContinue;
                Debuggee e;
                debuggees.TryGetValue(pid, out e);
                if (e != null) { e.Events++; }
                if (kind == 1)
                {
                    uint code = (uint)Marshal.ReadInt32(debugEvent, UnionOffset);
                    bool firstChance = Marshal.ReadInt32(debugEvent, FirstChanceOffset) != 0;
                    status = DbgExceptionNotHandled;
                    if (e != null)
                    {
                        int seen;
                        e.Exceptions.TryGetValue(code, out seen);
                        e.Exceptions[code] = seen + 1;
                        status = ContinueStatus(code, firstChance, e.AttachBreakpointPending);
                        if (firstChance && code == Breakpoint && e.AttachBreakpointPending)
                        {
                            e.AttachBreakpointPending = false;
                        }
                        else if (firstChance && code == AccessViolation && !e.FirstChanceDumped)
                        {
                            long address = Marshal.ReadInt64(debugEvent, AddressOffset);
                            string where = ModuleOf(e, address, module);
                            if (IsHunted(code, firstChance, where, module))
                            {
                                e.FirstChanceDumped = true;
                                long how = Marshal.ReadInt64(debugEvent, ParametersOffset);
                                long what = Marshal.ReadInt64(debugEvent, ParametersOffset + 8);
                                Say("pid " + pid + ": FIRST-CHANCE access violation in " + where + " at 0x" + address.ToString("X") + " (+0x" + (address - e.HuntBase).ToString("X")
                                    + "), " + (how == 0 ? "reading" : (how == 1 ? "writing" : "executing")) + " 0x" + what.ToString("X") + ", thread " + tid);
                                WriteDump(e, tid, debugEvent, true, folder);
                            }
                        }
                        else if (!firstChance && !e.SecondChanceDumped)
                        {
                            e.SecondChanceDumped = true;
                            Say("pid " + pid + ": SECOND-CHANCE exception 0x" + code.ToString("X8") + " - Outlook will not survive it; thread " + tid);
                            WriteDump(e, tid, debugEvent, false, folder);
                        }
                    }
                }
                else if (kind == 3 || kind == 6)
                {
                    IntPtr file = Marshal.ReadIntPtr(debugEvent, UnionOffset);
                    if (file != IntPtr.Zero) { CloseHandle(file); }
                }
                else if (kind == 5 && e != null)
                {
                    Say("pid " + pid + ": exited with 0x" + ((uint)Marshal.ReadInt32(debugEvent, UnionOffset)).ToString("X8") + "; " + Summary(e));
                    if (e.Process != IntPtr.Zero) { CloseHandle(e.Process); }
                    debuggees.Remove(pid);
                }

                ContinueDebugEvent(pid, tid, status);
            }
        }
        finally
        {
            foreach (Debuggee d in debuggees.Values)
            {
                bool stopped = DebugActiveProcessStop(d.Pid);
                Say("pid " + d.Pid + ": " + (stopped ? "detached" : "detach failed (error " + Marshal.GetLastWin32Error() + ")") + "; " + Summary(d));
                if (d.Process != IntPtr.Zero) { CloseHandle(d.Process); }
            }

            Marshal.FreeHGlobal(debugEvent);
            Say("watch ended");
        }

        return 0;
    }
}
'@

# =============================================================================================
# PURE DECISIONS. -SelfTest pins them.
# =============================================================================================

function Invoke-SelfTest {
    $script:stFailures = New-Object System.Collections.Generic.List[string]
    $script:stChecks = 0
    function Check([string] $Name, $Expected, $Actual) {
        $script:stChecks++
        if ([string]$Expected -cne [string]$Actual) { $script:stFailures.Add("$Name - expected '$Expected', got '$Actual'") }
    }
    Write-Host "Watch-OutlookFirstChance.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). Pure: compiles the debugger, attaches to nothing."
    Add-Type -TypeDefinition $DebuggerSource
    $t = [OutlookFirstChance]
    $av = [Convert]::ToUInt32('C0000005', 16)
    $bp = [Convert]::ToUInt32('80000003', 16)

    Write-Host '== which exception is dumped =='
    Check 'an access violation at first chance in the module is the hunted one' $true ($t::IsHunted($av, $true, 'mso40uiwin32client.dll', 'mso40uiwin32client.dll'))
    Check 'the module is matched without regard to case' $true ($t::IsHunted($av, $true, 'MSO40UIWin32Client.DLL', 'mso40uiwin32client.dll'))
    Check 'not in another module' $false ($t::IsHunted($av, $true, 'OLMAPI32.DLL', 'mso40uiwin32client.dll'))
    Check 'not where no module holds the address' $false ($t::IsHunted($av, $true, $null, 'mso40uiwin32client.dll'))
    Check 'not at second chance - that dump is the second-chance one' $false ($t::IsHunted($av, $false, 'mso40uiwin32client.dll', 'mso40uiwin32client.dll'))
    Check 'not another exception in the module' $false ($t::IsHunted([Convert]::ToUInt32('E06D7363', 16), $true, 'mso40uiwin32client.dll', 'mso40uiwin32client.dll'))
    $paramAst = [System.Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$null, [ref]$null).ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Module' }
    Check 'the default module is mso40uiwin32client.dll, the module both crashes of 2026-11-18 named' 'mso40uiwin32client.dll' $paramAst.DefaultValue.Value

    Write-Host '== how each exception is continued =='
    $continue = [Convert]::ToUInt32('00010002', 16)
    $notHandled = [Convert]::ToUInt32('80010001', 16)
    Check 'the attach''s own breakpoint is continued as handled' $continue ($t::ContinueStatus($bp, $true, $true))
    Check 'any later breakpoint goes back to Outlook unhandled' $notHandled ($t::ContinueStatus($bp, $true, $false))
    Check 'an access violation goes back to Outlook unhandled, dumped or not' $notHandled ($t::ContinueStatus($av, $true, $true))
    Check 'a second-chance exception goes back unhandled - the debugger never rescues Outlook' $notHandled ($t::ContinueStatus($av, $false, $false))
    Check 'the constants are Windows'' own DBG_CONTINUE and DBG_EXCEPTION_NOT_HANDLED' '65538|2147549185' "$($t::DbgContinue)|$($t::DbgExceptionNotHandled)"

    Write-Host '== the dumps =='
    Check 'a first-chance dump is named for its moment' 'OUTLOOK.EXE.4242.firstchance.dmp' ($t::DumpName(4242, $true))
    Check 'and a second-chance one' 'OUTLOOK.EXE.4242.secondchance.dmp' ($t::DumpName(4242, $false))
    Check 'neither reads as a LocalDumps dump to the runner (<program>.<pid>.dmp)' $false ($t::DumpName(4242, $true) -match '^(.+?)\.\d+\.dmp$')
    Check 'a full dump: full memory, handles, unloaded modules, the memory map, thread times' ('0x{0:X}' -f (0x2 -bor 0x4 -bor 0x20 -bor 0x800 -bor 0x1000)) ('0x{0:X}' -f $t::FullDump)

    Write-Host '== the debug event, as Windows lays it out on x64 =='
    Check 'the event is 176 bytes, the exception record at 16, its address at 32, first chance at 168' '176|16|32|48|168' "$($t::EventSize)|$($t::UnionOffset)|$($t::AddressOffset)|$($t::ParametersOffset)|$($t::FirstChanceOffset)"
    Check 'this PowerShell is 64-bit, as Outlook is - the layout above is x64''s' 8 ([IntPtr]::Size)

    Write-Host '== what it never does =='
    $code = ($DebuggerSource -split "`n" | Where-Object { $_ -notmatch '^\s*//' }) -join "`n"
    Check 'nothing in it ends a process' $false ($code -match 'TerminateProcess|\.Kill\(|DebugSetProcessKillOnExit\(true\)')
    Check 'it never writes into Outlook or suspends it' $false ($code -match 'WriteProcessMemory|SuspendThread|SetThreadContext|DebugBreakProcess')
    Check 'every attach is followed by kill-on-exit OFF' $true ($code -match '(?s)DebugActiveProcess\(p\.Id\).*?DebugSetProcessKillOnExit\(false\)')
    Check 'it watches its own session only' $true ($code.Contains('inSession != session'))
    $self = [System.IO.File]::ReadAllText($PSCommandPath)
    Check 'it refuses the Exchange VM, whose Outlook holds a real mailbox' $true ($self.Contains("-match 'EXCHANGE'"))

    if ($script:stFailures.Count -eq 0) { Write-Host "SELF-TEST PASS - $($script:stChecks) check(s)."; return 0 }
    Write-Host "SELF-TEST FAIL - $($script:stFailures.Count) of $($script:stChecks) check(s):"
    foreach ($f in $script:stFailures) { Write-Host "  - $f" }
    return 1
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# =============================================================================================
# THE GUEST.
# =============================================================================================

# travels-alone: Invoke-LiveTierOnGuest.ps1 -HuntFirstChance stages THIS checkout's copy beside any commit's guest scripts.
# The guest guard: Test-GuestIdentity and Assert-TestbedGuestIdentity, copies of ScriptHelpers.ps1's that
# Tools/Checks/check-powershell-51.ps1 check 4 holds identical, then what this script would have done elsewhere.
function Test-GuestIdentity {
    param([string] $UserName, [string] $ComputerName, [string[]] $Users, [string] $Prefix)
    $userOk = $false
    foreach ($u in $Users) { if ($UserName -eq $u) { $userOk = $true } }
    $machineOk = [bool]($Prefix -and $ComputerName -and $ComputerName.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase))
    return ($userOk -and $machineOk)
}

function Assert-TestbedGuestIdentity {
    param([string] $Why)
    $users = @($script:ExpectedUser)
    $prefix = [string]$script:ExpectedComputerNamePrefix
    if (Test-GuestIdentity -UserName $env:USERNAME -ComputerName $env:COMPUTERNAME -Users $users -Prefix $prefix) { return }
    throw ("REFUSING TO RUN.`n`n  logged on as : '{0}'          (allowed: {1})`n  computer name: '{2}'      (must start with: '{3}')`n`n{4}" -f $env:USERNAME, ($users -join ', '), $env:COMPUTERNAME, $prefix, $Why)
}

$GuestGuardWhy = @"
This script attaches a debugger to Outlook and writes Outlook's whole memory - mail included - to a folder. That
belongs on a test guest, never on a machine with real mail on it. The testbed guests autologon as 'vmadmin' and are
named 'OAI-...' (Testbed/host/New-AnswerFile.ps1). Do not 'fix' this by widening either default.
"@

Assert-TestbedGuestIdentity -Why $GuestGuardWhy
if ($env:COMPUTERNAME -match 'EXCHANGE') {
    throw "REFUSING TO RUN on $($env:COMPUTERNAME): the Exchange VM's Outlook holds a real mailbox, and a full dump of it would hold that mail."
}

try {
    Add-Type -TypeDefinition $DebuggerSource
    if (-not (Test-Path -LiteralPath $DumpFolder)) { New-Item -ItemType Directory -Force -Path $DumpFolder | Out-Null }
    exit ([OutlookFirstChance]::Run($Module, $DumpFolder, $StopFile, $Log, $MaxMinutes))
}
catch {
    $line = 'WATCH COULD NOT START: ' + $_.Exception.Message
    Write-Host $line
    if ($Log) { Add-Content -LiteralPath $Log -Value $line }
    exit 1
}
