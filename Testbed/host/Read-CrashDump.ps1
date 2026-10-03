#Requires -Version 5.1
<#
.SYNOPSIS
    Reads a Windows crash dump (a minidump - what Windows Error Reporting's LocalDumps writes) with
    nothing but what ships with Windows: the exception, the faulting module and offset, the faulting
    thread's stack unwound by Windows' own dbghelp.dll, and the C++ classes of the objects on it.

.DESCRIPTION
    Written 2026-10-04 for the OUTLOOK.EXE crashes on the Outlook test guests (Docs/live-tier-on-the-vm.md
    section 4.6), whose dumps Testbed/host/Invoke-LiveTierOnGuest.ps1 now fetches. The Debugging Tools
    for Windows (WinDbg, cdb) are not on this machine and would need an exception to the Dependencies
    rule (AGENTS.md), so this does what a first look in a debugger does, from the dump file alone:

      * the header, the system, the process, and every stream the dump holds;
      * the exception: code (named), flags, address, its module and offset, and for an access violation
        whether it read, wrote or executed, and where;
      * the faulting thread's stack, UNWOUND - dbghelp.dll's StackWalk64 (System32, every Windows) on
        the dump's own memory, each function's unwind data read from its module's .pdata in that memory;
        a frame is named module!export+0xN when it is inside an exported function, else module+0xRVA
        with the start of the function it is in - Office ships no exports for most of itself and no
        symbols are fetched, so that is as far as names go;
      * a SCAN of the same stack for return addresses (a value just after a CALL into a module) - the
        cross-check when the unwind stops early;
      * the C++ class of every object a register or a stack slot points at, from MSVC's run-time type
        information in the dump's memory (the vtable, its complete-object locator, the type
        descriptor's decorated name - ".?AVCDocument@@" is class CDocument). COM objects are C++
        objects with vtables, so this names what was being called or released, when the module keeps RTTI;
      * every other thread's top frames, briefly;
      * each faulting-stack module's version and CodeView record (PDB name, GUID, age) - what a symbol
        server would be asked for, if fetching symbols is ever allowed.

    Read-only: it maps the dump file and nothing else. Prints the report; -OutFile writes it too.

.PARAMETER Path
    The dump file.

.PARAMETER OutFile
    Also write the report here (UTF-8).

.PARAMETER OtherThreadFrames
    How many top frames each other thread gets. Default 6; 0 leaves the other threads out.

.PARAMETER SelfTest
    Builds a small minidump in memory - header, system, module, thread, exception and memory streams,
    a module image with an export and .pdata, a stack and an RTTI-carrying object - and holds the reader
    to what it put there. No file, no process.

.EXAMPLE
    pwsh -File Testbed/host/Read-CrashDump.ps1 .work\guest-live-runs\<run>\dumps\OutlookAI-Indexed\OUTLOOK.EXE.4242.dmp
    pwsh -File Testbed/host/Read-CrashDump.ps1 -SelfTest
#>
[CmdletBinding(DefaultParameterSetName = 'Read')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Read', Position = 0)] [string] $Path,
    [Parameter(ParameterSetName = 'Read')] [string] $OutFile,
    [Parameter(ParameterSetName = 'Read')] [int] $OtherThreadFrames = 6,
    [Parameter(Mandatory = $true, ParameterSetName = 'SelfTest')] [switch] $SelfTest
)

$ErrorActionPreference = 'Stop'

# C# 5 on purpose: Windows PowerShell 5.1's Add-Type compiles with the .NET Framework's csc, which
# stops at C# 5 - no interpolation, no ?., no nameof - and the build VM runs -SelfTest there.
$Source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookAICrashDump
{
    public sealed class Module
    {
        public ulong Base; public uint Size; public string Name; public string Version;
        public string Pdb; public string PdbId;
        public uint ExceptionDirRva; public uint ExceptionDirSize;
        public List<KeyValuePair<uint, string>> Exports;
        public uint[] FunctionBegin; public uint[] FunctionEnd; public uint[] FunctionUnwind;
        public string ShortName { get { int i = Name.LastIndexOf('\\'); return i >= 0 ? Name.Substring(i + 1) : Name; } }
        public bool Contains(ulong a) { return a >= Base && a < Base + Size; }
    }

    public sealed class ThreadInfo
    {
        public uint Id; public ulong StackStart; public uint StackSize; public uint StackRva;
        public uint ContextRva; public uint ContextSize; public string Name;
    }

    public sealed class Range64 { public ulong Start; public ulong Size; public long FileOffset; }

    public sealed class Frame { public ulong Pc; public ulong Sp; public string Text; }

    public sealed class Dump : IDisposable
    {
        private readonly byte[] _bytes;   // the whole file when small (self-test); else null
        private readonly FileStream _file;
        public readonly List<string> Notes = new List<string>();
        public uint Signature, Version, StreamCount, DirRva, TimeDateStamp; public ulong Flags;
        public readonly Dictionary<uint, KeyValuePair<uint, uint>> Streams = new Dictionary<uint, KeyValuePair<uint, uint>>();
        public readonly List<Module> Modules = new List<Module>();
        public readonly List<ThreadInfo> Threads = new List<ThreadInfo>();
        public readonly List<Range64> Memory = new List<Range64>();
        public ushort Architecture; public uint OsMajor, OsMinor, OsBuild;
        public uint ProcessId; public uint ProcessCreateTime; public bool HasMisc;
        public bool HasException; public uint ExceptionThread, ExceptionCode, ExceptionFlags, ExceptionParamCount;
        public ulong ExceptionAddress; public ulong[] ExceptionParams = new ulong[15]; public uint ExceptionContextRva, ExceptionContextSize;

        public Dump(string path) { _file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16, FileOptions.RandomAccess); Parse(); }
        public Dump(byte[] bytes) { _bytes = bytes; Parse(); }
        public void Dispose() { if (_file != null) _file.Dispose(); }

        public long Length { get { return _bytes != null ? _bytes.Length : _file.Length; } }

        public byte[] ReadFile(long offset, int count)
        {
            if (offset < 0 || count < 0 || offset + count > Length) return null;
            byte[] b = new byte[count];
            if (_bytes != null) { Buffer.BlockCopy(_bytes, (int)offset, b, 0, count); return b; }
            _file.Position = offset;
            int got = 0;
            while (got < count) { int n = _file.Read(b, got, count - got); if (n <= 0) return null; got += n; }
            return b;
        }
        private uint U32(long o) { byte[] b = ReadFile(o, 4); return b == null ? 0 : BitConverter.ToUInt32(b, 0); }
        private ulong U64(long o) { byte[] b = ReadFile(o, 8); return b == null ? 0 : BitConverter.ToUInt64(b, 0); }
        private string MinidumpString(long rva)
        {
            uint len = U32(rva);
            if (len == 0 || len > 4096) return string.Empty;
            byte[] b = ReadFile(rva + 4, (int)len);
            return b == null ? string.Empty : Encoding.Unicode.GetString(b);
        }

        private void Parse()
        {
            Signature = U32(0); Version = U32(4); StreamCount = U32(8); DirRva = U32(12); TimeDateStamp = U32(20); Flags = U64(24);
            if (Signature != 0x504D444D) throw new InvalidDataException("not a minidump (no MDMP signature)");
            for (uint i = 0; i < StreamCount && i < 4096; i++)
            {
                long d = DirRva + i * 12L;
                uint type = U32(d);
                if (type != 0 && !Streams.ContainsKey(type)) Streams[type] = new KeyValuePair<uint, uint>(U32(d + 4), U32(d + 8));
            }
            KeyValuePair<uint, uint> s;
            if (Streams.TryGetValue(7, out s)) { byte[] b = ReadFile(s.Value, 24); if (b != null) { Architecture = BitConverter.ToUInt16(b, 0); OsMajor = BitConverter.ToUInt32(b, 8); OsMinor = BitConverter.ToUInt32(b, 12); OsBuild = BitConverter.ToUInt32(b, 16); } }
            if (Streams.TryGetValue(15, out s) && s.Key >= 16) { HasMisc = true; uint flags1 = U32(s.Value + 4); if ((flags1 & 1) != 0) ProcessId = U32(s.Value + 8); if ((flags1 & 2) != 0) ProcessCreateTime = U32(s.Value + 12); }
            if (Streams.TryGetValue(9, out s))
            {
                ulong n = U64(s.Value); long data = (long)U64(s.Value + 8);
                for (ulong i = 0; i < n && i < 1000000; i++)
                {
                    long d = s.Value + 16 + (long)i * 16;
                    Range64 r = new Range64(); r.Start = U64(d); r.Size = U64(d + 8); r.FileOffset = data;
                    data += (long)r.Size; Memory.Add(r);
                }
            }
            if (Streams.TryGetValue(5, out s))
            {
                uint n = U32(s.Value);
                for (uint i = 0; i < n && i < 1000000; i++)
                {
                    long d = s.Value + 4 + i * 16L;
                    Range64 r = new Range64(); r.Start = U64(d); r.Size = U32(d + 8); r.FileOffset = U32(d + 12); Memory.Add(r);
                }
            }
            Memory.Sort(delegate (Range64 a, Range64 b) { return a.Start.CompareTo(b.Start); });
            if (Streams.TryGetValue(4, out s))
            {
                uint n = U32(s.Value);
                for (uint i = 0; i < n && i < 10000; i++)
                {
                    long d = s.Value + 4 + i * 108L;
                    Module m = new Module();
                    m.Base = U64(d); m.Size = U32(d + 8); m.Name = MinidumpString(U32(d + 20));
                    uint ms = U32(d + 24 + 8), ls = U32(d + 24 + 12);
                    m.Version = string.Format("{0}.{1}.{2}.{3}", ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF);
                    uint cvSize = U32(d + 76), cvRva = U32(d + 80);
                    if (cvSize >= 24 && U32(cvRva) == 0x53445352)
                    {
                        byte[] g = ReadFile(cvRva + 4, 16);
                        uint age = U32(cvRva + 20);
                        byte[] nameBytes = ReadFile(cvRva + 24, (int)Math.Min(cvSize - 24, 1024));
                        if (g != null && nameBytes != null)
                        {
                            int z = Array.IndexOf(nameBytes, (byte)0); if (z < 0) z = nameBytes.Length;
                            m.Pdb = Encoding.UTF8.GetString(nameBytes, 0, z);
                            m.PdbId = new Guid(g).ToString("N").ToUpperInvariant() + age.ToString("X");
                        }
                    }
                    Modules.Add(m);
                }
                Modules.Sort(delegate (Module a, Module b) { return a.Base.CompareTo(b.Base); });
            }
            if (Streams.TryGetValue(3, out s))
            {
                uint n = U32(s.Value);
                for (uint i = 0; i < n && i < 100000; i++)
                {
                    long d = s.Value + 4 + i * 48L;
                    ThreadInfo t = new ThreadInfo();
                    t.Id = U32(d); t.StackStart = U64(d + 24); t.StackSize = U32(d + 32); t.StackRva = U32(d + 36);
                    t.ContextSize = U32(d + 40); t.ContextRva = U32(d + 44);
                    Threads.Add(t);
                }
            }
            if (Streams.TryGetValue(24, out s))
            {
                uint n = U32(s.Value);
                for (uint i = 0; i < n && i < 100000; i++)
                {
                    long d = s.Value + 4 + i * 12L;
                    uint id = U32(d); ulong rva = U64(d + 4);
                    string name = MinidumpString((long)rva);
                    foreach (ThreadInfo t in Threads) if (t.Id == id && name.Length > 0) t.Name = name;
                }
            }
            if (Streams.TryGetValue(6, out s))
            {
                HasException = true;
                ExceptionThread = U32(s.Value);
                long r = s.Value + 8;
                ExceptionCode = U32(r); ExceptionFlags = U32(r + 4); ExceptionAddress = U64(r + 16); ExceptionParamCount = U32(r + 24);
                for (int i = 0; i < 15; i++) ExceptionParams[i] = U64(r + 32 + i * 8);
                ExceptionContextSize = U32(s.Value + 160); ExceptionContextRva = U32(s.Value + 164);
            }
            foreach (Module m in Modules) LoadImageTables(m);
        }

        // Reading the process's memory, as the dump holds it.
        public int FindRange(ulong a)
        {
            int lo = 0, hi = Memory.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2; Range64 r = Memory[mid];
                if (a < r.Start) hi = mid - 1; else if (a >= r.Start + r.Size) lo = mid + 1; else return mid;
            }
            return -1;
        }
        public byte[] Read(ulong address, int count)
        {
            if (count <= 0) return null;
            byte[] result = new byte[count];
            int done = 0;
            while (done < count)
            {
                int i = FindRange(address + (ulong)done);
                if (i < 0) return null;
                Range64 r = Memory[i];
                ulong offsetIn = address + (ulong)done - r.Start;
                int take = (int)Math.Min((ulong)(count - done), r.Size - offsetIn);
                byte[] b = ReadFile(r.FileOffset + (long)offsetIn, take);
                if (b == null) return null;
                Buffer.BlockCopy(b, 0, result, done, take);
                done += take;
            }
            return result;
        }
        public bool TryU64(ulong a, out ulong v) { byte[] b = Read(a, 8); v = b == null ? 0 : BitConverter.ToUInt64(b, 0); return b != null; }
        public bool TryU32(ulong a, out uint v) { byte[] b = Read(a, 4); v = b == null ? 0 : BitConverter.ToUInt32(b, 0); return b != null; }
        public string ReadAscii(ulong a, int max)
        {
            byte[] b = Read(a, max);
            if (b == null) { for (int n = max / 2; n >= 16 && b == null; n /= 2) b = Read(a, n); }
            if (b == null) return null;
            int z = Array.IndexOf(b, (byte)0); if (z < 0) z = b.Length;
            return Encoding.ASCII.GetString(b, 0, z);
        }

        private void LoadImageTables(Module m)
        {
            uint lfanew;
            if (!TryU32(m.Base + 0x3C, out lfanew) || lfanew > 4096) return;
            uint pe; if (!TryU32(m.Base + lfanew, out pe) || pe != 0x4550) return;
            ulong opt = m.Base + lfanew + 24;
            uint magic; if (!TryU32(opt, out magic) || (magic & 0xFFFF) != 0x20B) return;
            uint expRva, expSize, excRva, excSize;
            TryU32(opt + 112, out expRva); TryU32(opt + 116, out expSize);
            TryU32(opt + 136, out excRva); TryU32(opt + 140, out excSize);
            m.ExceptionDirRva = excRva; m.ExceptionDirSize = excSize;
            if (excRva != 0 && excSize >= 12 && excSize < 64 * 1024 * 1024)
            {
                byte[] t = Read(m.Base + excRva, (int)excSize);
                if (t != null)
                {
                    int n = t.Length / 12;
                    m.FunctionBegin = new uint[n]; m.FunctionEnd = new uint[n]; m.FunctionUnwind = new uint[n];
                    for (int i = 0; i < n; i++) { m.FunctionBegin[i] = BitConverter.ToUInt32(t, i * 12); m.FunctionEnd[i] = BitConverter.ToUInt32(t, i * 12 + 4); m.FunctionUnwind[i] = BitConverter.ToUInt32(t, i * 12 + 8); }
                }
            }
            m.Exports = new List<KeyValuePair<uint, string>>();
            if (expRva != 0 && expSize >= 40)
            {
                byte[] e = Read(m.Base + expRva, 40);
                if (e != null)
                {
                    uint nNames = BitConverter.ToUInt32(e, 24), aFuncs = BitConverter.ToUInt32(e, 28), aNames = BitConverter.ToUInt32(e, 32), aOrds = BitConverter.ToUInt32(e, 36);
                    uint nFuncs = BitConverter.ToUInt32(e, 20);
                    if (nNames < 200000 && nFuncs < 200000)
                    {
                        byte[] names = Read(m.Base + aNames, (int)nNames * 4), ords = Read(m.Base + aOrds, (int)nNames * 2), funcs = Read(m.Base + aFuncs, (int)nFuncs * 4);
                        if (names != null && ords != null && funcs != null)
                        {
                            for (int i = 0; i < nNames; i++)
                            {
                                uint nameRva = BitConverter.ToUInt32(names, i * 4); ushort ord = BitConverter.ToUInt16(ords, i * 2);
                                if (ord >= nFuncs) continue;
                                uint fn = BitConverter.ToUInt32(funcs, ord * 4);
                                if (fn >= expRva && fn < expRva + expSize) continue; // a forwarder
                                string name = ReadAscii(m.Base + nameRva, 256);
                                if (!string.IsNullOrEmpty(name)) m.Exports.Add(new KeyValuePair<uint, string>(fn, name));
                            }
                            m.Exports.Sort(delegate (KeyValuePair<uint, string> a, KeyValuePair<uint, string> b) { return a.Key.CompareTo(b.Key); });
                        }
                    }
                }
            }
        }

        public Module ModuleAt(ulong a)
        {
            int lo = 0, hi = Modules.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2; Module m = Modules[mid];
                if (a < m.Base) hi = mid - 1; else if (a >= m.Base + m.Size) lo = mid + 1; else return m;
            }
            return null;
        }

        // The index of the .pdata entry whose function holds this RVA, or -1 (a leaf, or no table).
        public static int FunctionIndex(Module m, uint rva)
        {
            if (m == null || m.FunctionBegin == null) return -1;
            int lo = 0, hi = m.FunctionBegin.Length - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (rva < m.FunctionBegin[mid]) hi = mid - 1; else if (rva >= m.FunctionEnd[mid]) lo = mid + 1; else return mid;
            }
            return -1;
        }

        // The start of the PRIMARY function holding this RVA - following chained unwind info, so a
        // function split into pieces is named by its entry.
        public uint FunctionStart(Module m, uint rva)
        {
            int i = FunctionIndex(m, rva);
            if (i < 0) return uint.MaxValue;
            uint begin = m.FunctionBegin[i], unwind = m.FunctionUnwind[i];
            for (int hop = 0; hop < 32; hop++)
            {
                byte[] head = Read(m.Base + (unwind & ~1u), 4);
                if (head == null) break;
                byte flags = (byte)(head[0] >> 3), codes = head[2];
                if ((flags & 4) == 0) break; // UNW_FLAG_CHAININFO
                ulong chained = m.Base + (unwind & ~1u) + 4 + (ulong)(((codes + 1) & ~1) * 2);
                byte[] rf = Read(chained, 12);
                if (rf == null) break;
                begin = BitConverter.ToUInt32(rf, 0); unwind = BitConverter.ToUInt32(rf, 8);
            }
            return begin;
        }

        public string Describe(ulong a)
        {
            Module m = ModuleAt(a);
            if (m == null) return "0x" + a.ToString("X16");
            uint rva = (uint)(a - m.Base);
            uint start = FunctionStart(m, rva);
            string name = null;
            if (m.Exports != null && m.Exports.Count > 0)
            {
                int lo = 0, hi = m.Exports.Count - 1, best = -1;
                while (lo <= hi) { int mid = (lo + hi) / 2; if (m.Exports[mid].Key <= rva) { best = mid; lo = mid + 1; } else hi = mid - 1; }
                // Only when the export IS the function this address is in - a nearest export
                // elsewhere in the image would name the wrong code with confidence.
                if (best >= 0 && (start == uint.MaxValue ? rva - m.Exports[best].Key < 0x40 : m.Exports[best].Key == start)) name = m.Exports[best].Value + "+0x" + (rva - m.Exports[best].Key).ToString("X");
            }
            if (name != null) return m.ShortName + "!" + name;
            if (start != uint.MaxValue) return m.ShortName + "+0x" + rva.ToString("X") + " (fn " + m.ShortName + "+0x" + start.ToString("X") + " +0x" + (rva - start).ToString("X") + ")";
            return m.ShortName + "+0x" + rva.ToString("X");
        }

        // The C++ class of the object at this address, from MSVC RTTI: vtable, complete-object
        // locator (x64: signature 1, image-relative), type descriptor, decorated name. Null when not.
        public string ClassOf(ulong obj)
        {
            ulong vtable;
            if (!TryU64(obj, out vtable)) return null;
            Module vm = ModuleAt(vtable);
            if (vm == null) return null;
            ulong col;
            if (!TryU64(vtable - 8, out col)) return null;
            Module cm = ModuleAt(col);
            if (cm == null) return null;
            byte[] c = Read(col, 24);
            if (c == null || BitConverter.ToUInt32(c, 0) != 1) return null;
            uint offset = BitConverter.ToUInt32(c, 4), td = BitConverter.ToUInt32(c, 12), self = BitConverter.ToUInt32(c, 20);
            if (cm.Base + self != col) return null;
            string decorated = ReadAscii(cm.Base + td + 16, 256);
            if (decorated == null || !decorated.StartsWith(".?A")) return null;
            string name = Undecorate(decorated);
            if (offset != 0) name += " (at +0x" + offset.ToString("X") + " in the object)";
            return name + " [" + vm.ShortName + "]";
        }

        public static string Undecorate(string d)
        {
            // ".?AVCDocument@@" -> "CDocument"; ".?AVCFoo@NS@@" -> "NS::CFoo"; templates are left decorated.
            string body = d.Substring(4);
            if (body.EndsWith("@@")) body = body.Substring(0, body.Length - 2);
            if (body.IndexOf('?') >= 0 || body.IndexOf('$') >= 0) return d;
            string[] parts = body.Split('@');
            Array.Reverse(parts);
            return string.Join("::", parts);
        }

        public byte[] Context(uint rva, uint size) { return size < 1232 ? null : ReadFile(rva, 1232); }
    }

    public static class Walker
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RuntimeFunction { public uint Begin; public uint End; public uint Unwind; }

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate bool ReadMemory(IntPtr process, ulong address, IntPtr buffer, uint size, out uint read);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate IntPtr FunctionTableAccess(IntPtr process, ulong address);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate ulong GetModuleBase(IntPtr process, ulong address);

        [DllImport("dbghelp.dll", SetLastError = true)]
        private static extern bool StackWalk64(uint machine, IntPtr process, IntPtr thread, IntPtr frame, IntPtr context,
            ReadMemory read, FunctionTableAccess table, GetModuleBase moduleBase, IntPtr translate);

        // StackWalk64 on the dump: dbghelp's own x64 unwinder, fed the dump's memory and each
        // module's .pdata. Returns the frames (pc, sp), at most maxFrames.
        public static List<Frame> Unwind(Dump dump, byte[] context, int maxFrames, out string stop)
        {
            stop = null;
            List<Frame> frames = new List<Frame>();
            if (context == null) { stop = "no thread context in the dump"; return frames; }
            List<IntPtr> keep = new List<IntPtr>();
            Dictionary<ulong, IntPtr> tables = new Dictionary<ulong, IntPtr>();
            IntPtr ctx = Marshal.AllocHGlobal(1232 + 16), frame = Marshal.AllocHGlobal(4096);
            try
            {
                // CONTEXT must be 16-byte aligned for x64.
                IntPtr ctxAligned = new IntPtr((ctx.ToInt64() + 15) & ~15L);
                Marshal.Copy(context, 0, ctxAligned, 1232);
                byte[] zero = new byte[4096]; Marshal.Copy(zero, 0, frame, 4096);
                ulong rip = BitConverter.ToUInt64(context, 248), rsp = BitConverter.ToUInt64(context, 152), rbp = BitConverter.ToUInt64(context, 160);
                Marshal.WriteInt64(frame, 0, (long)rip); Marshal.WriteInt32(frame, 12, 3);    // AddrPC, flat
                Marshal.WriteInt64(frame, 32, (long)rbp); Marshal.WriteInt32(frame, 44, 3);   // AddrFrame
                Marshal.WriteInt64(frame, 48, (long)rsp); Marshal.WriteInt32(frame, 60, 3);   // AddrStack
                ReadMemory read = delegate (IntPtr p, ulong a, IntPtr buf, uint size, out uint got)
                {
                    got = 0;
                    byte[] b = dump.Read(a, (int)size);
                    if (b == null)
                    {
                        // A partial read is still useful to the unwinder; give what is there.
                        for (int n = (int)size - 1; n > 0 && b == null; n--) { b = dump.Read(a, n); }
                        if (b == null) return false;
                    }
                    Marshal.Copy(b, 0, buf, b.Length); got = (uint)b.Length; return true;
                };
                FunctionTableAccess table = delegate (IntPtr p, ulong a)
                {
                    Module m = dump.ModuleAt(a);
                    if (m == null) return IntPtr.Zero;
                    int i = Dump.FunctionIndex(m, (uint)(a - m.Base));
                    if (i < 0) return IntPtr.Zero;
                    ulong key = m.Base + m.FunctionBegin[i];
                    IntPtr entry;
                    if (!tables.TryGetValue(key, out entry))
                    {
                        entry = Marshal.AllocHGlobal(12); keep.Add(entry);
                        Marshal.WriteInt32(entry, 0, (int)m.FunctionBegin[i]); Marshal.WriteInt32(entry, 4, (int)m.FunctionEnd[i]); Marshal.WriteInt32(entry, 8, (int)m.FunctionUnwind[i]);
                        tables[key] = entry;
                    }
                    return entry;
                };
                GetModuleBase moduleBase = delegate (IntPtr p, ulong a) { Module m = dump.ModuleAt(a); return m == null ? 0UL : m.Base; };
                ulong lastSp = 0;
                for (int n = 0; n < maxFrames; n++)
                {
                    if (!StackWalk64(0x8664, new IntPtr(1), new IntPtr(1), frame, ctxAligned, read, table, moduleBase, IntPtr.Zero))
                    {
                        stop = frames.Count > 0 && frames[frames.Count - 1].Text.Contains("RtlUserThreadStart") ? "end of stack" : "the unwinder stopped (StackWalk64 returned FALSE)";
                        break;
                    }
                    ulong pc = (ulong)Marshal.ReadInt64(frame, 0), sp = (ulong)Marshal.ReadInt64(frame, 48);
                    if (pc == 0) { stop = "end of stack"; break; }
                    if (n > 0 && sp <= lastSp && sp != 0) { stop = "the stack pointer stopped growing - the unwind is unreliable from here"; break; }
                    lastSp = sp;
                    Frame f = new Frame(); f.Pc = pc; f.Sp = sp; f.Text = dump.Describe(pc); frames.Add(f);
                }
                GC.KeepAlive(read); GC.KeepAlive(table); GC.KeepAlive(moduleBase);
            }
            finally
            {
                Marshal.FreeHGlobal(ctx); Marshal.FreeHGlobal(frame);
                foreach (IntPtr k in keep) Marshal.FreeHGlobal(k);
            }
            return frames;
        }

        // Return-address candidates on a stack: a value inside a module's executable code that
        // directly follows a CALL instruction. The cross-check for an unwind that stopped early.
        public static List<Frame> Scan(Dump dump, ulong rsp, ulong stackEnd, int max)
        {
            List<Frame> found = new List<Frame>();
            if (stackEnd <= rsp) return found;
            ulong span = Math.Min(stackEnd - rsp, 512UL * 1024);
            byte[] stack = dump.Read(rsp, (int)span);
            if (stack == null) return found;
            for (int off = 0; off + 8 <= stack.Length && found.Count < max; off += 8)
            {
                ulong v = BitConverter.ToUInt64(stack, off);
                Module m = dump.ModuleAt(v);
                if (m == null || Dump.FunctionIndex(m, (uint)(v - m.Base)) < 0 && m.FunctionBegin != null) continue;
                byte[] before = dump.Read(v - 7, 7);
                if (before == null || !FollowsCall(before)) continue;
                Frame f = new Frame(); f.Pc = v; f.Sp = rsp + (ulong)off; f.Text = dump.Describe(v); found.Add(f);
            }
            return found;
        }

        // The bytes before a return address end with a CALL: E8 rel32, FF /2 (ModRM forms), or
        // FF 15 disp32 / 41 FF Dx / FF D0-D7 and their REX variants.
        public static bool FollowsCall(byte[] b)
        {
            int n = b.Length;
            if (n >= 5 && b[n - 5] == 0xE8) return true;
            if (n >= 6 && b[n - 6] == 0xFF && b[n - 5] == 0x15) return true;
            if (n >= 7 && b[n - 7] >= 0x40 && b[n - 7] <= 0x4F && b[n - 6] == 0xFF && b[n - 5] == 0x15) return true;
            if (n >= 2 && b[n - 2] == 0xFF && (b[n - 1] & 0xF8) == 0xD0) return true;
            if (n >= 3 && b[n - 3] >= 0x40 && b[n - 3] <= 0x4F && b[n - 2] == 0xFF && (b[n - 1] & 0xF8) == 0xD0) return true;
            if (n >= 3 && b[n - 3] == 0xFF && ((b[n - 2] & 0x38) == 0x10) && (b[n - 2] & 0xC0) == 0x40) return true;
            if (n >= 6 && b[n - 6] == 0xFF && ((b[n - 5] & 0x38) == 0x10) && (b[n - 5] & 0xC0) == 0x80) return true;
            if (n >= 4 && b[n - 4] >= 0x40 && b[n - 4] <= 0x4F && b[n - 3] == 0xFF && ((b[n - 2] & 0x38) == 0x10) && (b[n - 2] & 0xC0) == 0x40) return true;
            if (n >= 2 && b[n - 2] == 0xFF && ((b[n - 1] & 0x38) == 0x10) && (b[n - 1] & 0xC0) == 0x00) return true;
            if (n >= 3 && b[n - 3] == 0xFF && ((b[n - 2] & 0x38) == 0x10) && (b[n - 2] & 0xC7) == 0x04) return true;
            return false;
        }
    }

    public static class Report
    {
        public static string ExceptionName(uint code)
        {
            switch (code)
            {
                case 0xC0000005: return "access violation";
                case 0xC0000374: return "heap corruption";
                case 0xC0000409: return "stack buffer overrun / fail-fast";
                case 0xC000001D: return "illegal instruction";
                case 0xC00000FD: return "stack overflow";
                case 0x80000003: return "breakpoint";
                case 0xC0000602: return "fail-fast exception";
                case 0xE06D7363: return "C++ exception";
                case 0xE0434352: return "CLR exception";
                case 0xC0020001: return "RPC: string binding is invalid";
                default: return "0x" + code.ToString("X8");
            }
        }

        public static string Build(Dump d, int otherThreadFrames)
        {
            StringBuilder s = new StringBuilder();
            DateTime when = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(d.TimeDateStamp);
            s.AppendLine("DUMP       " + when.ToString("o") + " UTC (the dump's own clock - a frozen guest's, which is not the host's)");
            s.AppendLine("           flags 0x" + d.Flags.ToString("X") + ((d.Flags & 2) != 0 ? " (full memory)" : " (NOT full memory)") + ", " + d.Streams.Count + " streams, " + d.Memory.Count + " memory ranges, " + d.Modules.Count + " modules, " + d.Threads.Count + " threads");
            s.AppendLine("SYSTEM     Windows " + d.OsMajor + "." + d.OsMinor + "." + d.OsBuild + ", architecture " + d.Architecture + (d.Architecture == 9 ? " (x64)" : ""));
            if (d.HasMisc) s.AppendLine("PROCESS    pid " + d.ProcessId + (d.ProcessCreateTime != 0 ? ", started " + new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(d.ProcessCreateTime).ToString("o") + " UTC (up " + Math.Round((d.TimeDateStamp - (double)d.ProcessCreateTime) / 60.0, 1) + " min)" : ""));
            if (!d.HasException) { s.AppendLine("EXCEPTION  none recorded - not a crash dump"); return s.ToString(); }
            s.AppendLine("EXCEPTION  " + ExceptionName(d.ExceptionCode) + " (0x" + d.ExceptionCode.ToString("X8") + "), flags 0x" + d.ExceptionFlags.ToString("X") + ", on thread " + d.ExceptionThread + ThreadName(d, d.ExceptionThread));
            s.AppendLine("           at 0x" + d.ExceptionAddress.ToString("X16") + "  " + d.Describe(d.ExceptionAddress));
            if (d.ExceptionCode == 0xC0000005 && d.ExceptionParamCount >= 2)
            {
                string kind = d.ExceptionParams[0] == 0 ? "READING" : d.ExceptionParams[0] == 1 ? "WRITING" : d.ExceptionParams[0] == 8 ? "EXECUTING" : "accessing";
                ulong target = d.ExceptionParams[1];
                s.AppendLine("           " + kind + " 0x" + target.ToString("X16") + (target < 0x10000 ? " - a null (or near-null) pointer" : "") + Where(d, target));
            }
            else if (d.ExceptionParamCount > 0)
            {
                StringBuilder p = new StringBuilder();
                for (int i = 0; i < d.ExceptionParamCount && i < 15; i++) p.Append(" 0x" + d.ExceptionParams[i].ToString("X"));
                s.AppendLine("           parameters" + p);
            }
            byte[] ctx = d.Context(d.ExceptionContextRva, d.ExceptionContextSize);
            if (ctx != null)
            {
                s.AppendLine();
                s.AppendLine("REGISTERS  (the faulting instruction's) - and the class of what each points at, by RTTI");
                string[] names = { "rax", "rcx", "rdx", "rbx", "rsp", "rbp", "rsi", "rdi", "r8", "r9", "r10", "r11", "r12", "r13", "r14", "r15", "rip" };
                int[] offs = { 120, 128, 136, 144, 152, 160, 168, 176, 184, 192, 200, 208, 216, 224, 232, 240, 248 };
                for (int i = 0; i < names.Length; i++)
                {
                    ulong v = BitConverter.ToUInt64(ctx, offs[i]);
                    string cls = d.ClassOf(v), code = d.ModuleAt(v) != null ? d.Describe(v) : null;
                    s.AppendLine("  " + names[i].PadRight(4) + "0x" + v.ToString("X16") + (cls != null ? "  -> object of class " + cls : code != null ? "  = " + code : Where(d, v)));
                }
                byte[] at = d.Read(d.ExceptionAddress, 16);
                if (at != null) s.AppendLine("  code at rip: " + BitConverter.ToString(at).Replace("-", " "));
            }
            ThreadInfo faulting = null;
            foreach (ThreadInfo t in d.Threads) if (t.Id == d.ExceptionThread) faulting = t;
            s.AppendLine();
            s.AppendLine("STACK      the faulting thread, unwound by dbghelp.dll from the exception context");
            string stop;
            List<Frame> frames = Walker.Unwind(d, ctx, 128, out stop);
            for (int i = 0; i < frames.Count; i++) s.AppendLine("  " + i.ToString("00") + "  " + frames[i].Text);
            s.AppendLine("  (" + frames.Count + " frame(s); " + (stop ?? "the frame limit") + ")");
            if (faulting != null && ctx != null)
            {
                ulong rsp = BitConverter.ToUInt64(ctx, 152);
                ulong end = faulting.StackStart + faulting.StackSize;
                List<Frame> scan = Walker.Scan(d, rsp, end, 64);
                s.AppendLine();
                s.AppendLine("STACK SCAN return addresses on the faulting thread's stack (values just after a CALL), top first");
                foreach (Frame f in scan) s.AppendLine("  sp+0x" + (f.Sp - rsp).ToString("X").PadLeft(5, '0') + "  " + f.Text);
                s.AppendLine();
                s.AppendLine("OBJECTS    every distinct C++ class a slot of the faulting stack points at (RTTI), top first");
                byte[] stack = d.Read(rsp, (int)Math.Min(end > rsp ? end - rsp : 0, 256UL * 1024));
                Dictionary<string, bool> seen = new Dictionary<string, bool>();
                if (stack != null)
                {
                    for (int off = 0; off + 8 <= stack.Length && seen.Count < 60; off += 8)
                    {
                        string cls = d.ClassOf(BitConverter.ToUInt64(stack, off));
                        if (cls != null && !seen.ContainsKey(cls)) { seen[cls] = true; s.AppendLine("  sp+0x" + off.ToString("X").PadLeft(5, '0') + "  " + cls); }
                    }
                }
                if (seen.Count == 0) s.AppendLine("  none found - the modules may carry no RTTI");
            }
            s.AppendLine();
            s.AppendLine("MODULES    those on the faulting stack - version, and the PDB a symbol server would be asked for");
            Dictionary<Module, bool> used = new Dictionary<Module, bool>();
            Module fm = d.ModuleAt(d.ExceptionAddress); if (fm != null) used[fm] = true;
            foreach (Frame f in frames) { Module m = d.ModuleAt(f.Pc); if (m != null) used[m] = true; }
            foreach (Module m in d.Modules) if (used.ContainsKey(m)) s.AppendLine("  " + m.ShortName.PadRight(22) + " " + m.Version.PadRight(18) + " 0x" + m.Base.ToString("X") + "  " + (m.Pdb ?? "-") + " " + (m.PdbId ?? "") + (m.FunctionBegin == null ? "  (no .pdata in the dump)" : ""));
            if (otherThreadFrames > 0)
            {
                s.AppendLine();
                s.AppendLine("THREADS    every other thread's top " + otherThreadFrames + " frames");
                foreach (ThreadInfo t in d.Threads)
                {
                    if (t.Id == d.ExceptionThread) continue;
                    string tstop;
                    List<Frame> tf = Walker.Unwind(d, d.Context(t.ContextRva, t.ContextSize), otherThreadFrames, out tstop);
                    StringBuilder line = new StringBuilder();
                    foreach (Frame f in tf) { if (line.Length > 0) line.Append(" <- "); line.Append(f.Text); }
                    s.AppendLine("  " + t.Id.ToString().PadLeft(6) + ThreadName(d, t.Id) + ": " + line);
                }
            }
            return s.ToString();
        }

        private static string ThreadName(Dump d, uint id)
        {
            foreach (ThreadInfo t in d.Threads) if (t.Id == id && !string.IsNullOrEmpty(t.Name)) return " (\"" + t.Name + "\")";
            return string.Empty;
        }

        private static string Where(Dump d, ulong a)
        {
            if (d.ModuleAt(a) != null) return "  (in " + d.ModuleAt(a).ShortName + ")";
            foreach (ThreadInfo t in d.Threads) if (a >= t.StackStart && a < t.StackStart + t.StackSize + 0x100000UL && a >= t.StackStart - 0x100000UL) return "  (near thread " + t.Id + "'s stack)";
            return d.FindRange(a) >= 0 ? "  (in the dump's memory)" : "  (NOT in the dump's memory - unmapped, freed or never committed)";
        }
    }
}
'@

function Import-CrashDumpReader {
    if (-not ('OutlookAICrashDump.Dump' -as [type])) { Add-Type -TypeDefinition $Source -Language CSharp }
}

# A small minidump, built here, holding what the reader reads: x64 system info, misc info, one module
# whose image (in memory) has an export directory and .pdata, one thread with a context and a stack,
# an access-violation exception on it, and an object whose vtable carries an MSVC RTTI locator.
function New-SelfTestDump {
    $base = [uint64]0x7FF600000000
    $image = New-Object byte[] 0x3000
    # PE: DOS header -> NT at 0x80; optional header magic 0x20B; export dir at 0x1000 (0x60 bytes); exception dir at 0x1100 (24 bytes).
    [BitConverter]::GetBytes([uint32]0x80).CopyTo($image, 0x3C)
    [BitConverter]::GetBytes([uint32]0x4550).CopyTo($image, 0x80)
    [BitConverter]::GetBytes([uint16]0x20B).CopyTo($image, 0x98)
    [BitConverter]::GetBytes([uint32]0x1000).CopyTo($image, 0x98 + 112); [BitConverter]::GetBytes([uint32]0x60).CopyTo($image, 0x98 + 116)
    [BitConverter]::GetBytes([uint32]0x1100).CopyTo($image, 0x98 + 136); [BitConverter]::GetBytes([uint32]24).CopyTo($image, 0x98 + 140)
    # Export directory: one function 'DoWork' at RVA 0x2000.
    [BitConverter]::GetBytes([uint32]1).CopyTo($image, 0x1000 + 20); [BitConverter]::GetBytes([uint32]1).CopyTo($image, 0x1000 + 24)
    [BitConverter]::GetBytes([uint32]0x1040).CopyTo($image, 0x1000 + 28); [BitConverter]::GetBytes([uint32]0x1044).CopyTo($image, 0x1000 + 32); [BitConverter]::GetBytes([uint32]0x1048).CopyTo($image, 0x1000 + 36)
    [BitConverter]::GetBytes([uint32]0x2000).CopyTo($image, 0x1040); [BitConverter]::GetBytes([uint32]0x1050).CopyTo($image, 0x1044); [BitConverter]::GetBytes([uint16]0).CopyTo($image, 0x1048)
    [System.Text.Encoding]::ASCII.GetBytes("DoWork`0").CopyTo($image, 0x1050)
    # .pdata: DoWork 0x2000-0x2100 (unwind 0x1200), and an unexported function 0x2200-0x2300 (unwind 0x1210).
    foreach ($e in @(@(0x1100, 0x2000, 0x2100, 0x1200), @(0x110C, 0x2200, 0x2300, 0x1210))) { [BitConverter]::GetBytes([uint32]$e[1]).CopyTo($image, $e[0]); [BitConverter]::GetBytes([uint32]$e[2]).CopyTo($image, $e[0] + 4); [BitConverter]::GetBytes([uint32]$e[3]).CopyTo($image, $e[0] + 8) }
    # UNWIND_INFO for both: version 1, no prolog, no codes - a frameless function; the return address is at rsp.
    $image[0x1200] = 1; $image[0x1210] = 1
    # A CALL (E8 rel32) at 0x2205, so 0x220A is a return address into the unexported function.
    $image[0x2205] = 0xE8
    # RTTI: type descriptor at 0x2800 (name at +16), complete-object locator at 0x2900, vtable at 0x2A08 (locator pointer at 0x2A00).
    [System.Text.Encoding]::ASCII.GetBytes(".?AVCTestDocument@@`0").CopyTo($image, 0x2810)
    [BitConverter]::GetBytes([uint32]1).CopyTo($image, 0x2900); [BitConverter]::GetBytes([uint32]0x2800).CopyTo($image, 0x2900 + 12); [BitConverter]::GetBytes([uint32]0x2900).CopyTo($image, 0x2900 + 20)
    [BitConverter]::GetBytes([uint64]($base + 0x2900)).CopyTo($image, 0x2A00)

    $stackBase = [uint64]0x0000009F00000000
    $stack = New-Object byte[] 0x100
    [BitConverter]::GetBytes([uint64]($base + 0x220A)).CopyTo($stack, 0)     # DoWork returns into the unexported function
    [BitConverter]::GetBytes([uint64]0x0000020000001000).CopyTo($stack, 8)  # the object below
    $heapBase = [uint64]0x0000020000001000
    $heap = New-Object byte[] 0x20
    [BitConverter]::GetBytes([uint64]($base + 0x2A08)).CopyTo($heap, 0)

    $context = New-Object byte[] 1232
    [BitConverter]::GetBytes([uint32]0x10001F).CopyTo($context, 48)
    [BitConverter]::GetBytes([uint64]$heapBase).CopyTo($context, 128)                 # rcx -> the object
    [BitConverter]::GetBytes([uint64]$stackBase).CopyTo($context, 152)                # rsp
    [BitConverter]::GetBytes([uint64]($base + 0x2010)).CopyTo($context, 248)          # rip in DoWork

    $ms = New-Object System.IO.MemoryStream
    $w = New-Object System.IO.BinaryWriter($ms)
    $nStreams = 6
    $w.Write([uint32]0x504D444D); $w.Write([uint32]0xA793); $w.Write([uint32]$nStreams); $w.Write([uint32]32); $w.Write([uint32]0); $w.Write([uint32]1791049386); $w.Write([uint64]0x1826)
    $dirAt = 32; $w.Write((New-Object byte[] ($nStreams * 12)))
    $dir = New-Object System.Collections.Generic.List[object]
    # system info (7)
    $at = $ms.Position; $w.Write([uint16]9); $w.Write((New-Object byte[] 6)); $w.Write([uint32]10); $w.Write([uint32]0); $w.Write([uint32]26100); $w.Write((New-Object byte[] 36)); $dir.Add(@(7, 56, $at))
    # misc info (15): flags1 = 1|2, pid 4242, created an hour before the dump
    $at = $ms.Position; $w.Write([uint32]24); $w.Write([uint32]3); $w.Write([uint32]4242); $w.Write([uint32](1791049386 - 3600)); $w.Write([uint32]0); $w.Write([uint32]0); $dir.Add(@(15, 24, $at))
    # module name string
    $nameAt = $ms.Position; $n = [System.Text.Encoding]::Unicode.GetBytes('C:\Program Files\Test\wwtest.dll'); $w.Write([uint32]$n.Length); $w.Write($n); $w.Write([uint16]0)
    # context
    $ctxAt = $ms.Position; $w.Write($context)
    # module list (4)
    $at = $ms.Position; $w.Write([uint32]1); $w.Write([uint64]$base); $w.Write([uint32]0x3000); $w.Write([uint32]0); $w.Write([uint32]0); $w.Write([uint32]$nameAt)
    $vs = New-Object byte[] 52; [BitConverter]::GetBytes([uint32]0x00100000).CopyTo($vs, 8); [BitConverter]::GetBytes([uint32]0x460C4E44).CopyTo($vs, 12); $w.Write($vs)
    $w.Write((New-Object byte[] 32)); $dir.Add(@(4, 4 + 108, $at))
    # memory 64 list (9): the image, the stack, the heap object - data placed after the descriptors
    $at = $ms.Position; $ranges = @(@($base, $image), @($stackBase, $stack), @($heapBase, $heap))
    $dataAt = $at + 16 + 16 * $ranges.Count
    $w.Write([uint64]$ranges.Count); $w.Write([uint64]$dataAt)
    foreach ($r in $ranges) { $w.Write([uint64]$r[0]); $w.Write([uint64]$r[1].Length) }
    foreach ($r in $ranges) { $w.Write([byte[]]$r[1]) }
    $dir.Add(@(9, 16 + 16 * $ranges.Count, $at))
    # thread list (3): thread 77, its stack in the memory list too
    $at = $ms.Position; $w.Write([uint32]1); $w.Write([uint32]77); $w.Write([uint32]0); $w.Write([uint32]0); $w.Write([uint32]0); $w.Write([uint64]0)
    $w.Write([uint64]$stackBase); $w.Write([uint32]$stack.Length); $w.Write([uint32]($dataAt + $image.Length)); $w.Write([uint32]1232); $w.Write([uint32]$ctxAt); $dir.Add(@(3, 52, $at))
    # exception (6): an access violation writing 0x18 on thread 77
    $at = $ms.Position; $w.Write([uint32]77); $w.Write([uint32]0); $w.Write([uint32]0xC0000005); $w.Write([uint32]0); $w.Write([uint64]0); $w.Write([uint64]($base + 0x2010)); $w.Write([uint32]2); $w.Write([uint32]0)
    $p = New-Object uint64[] 15; $p[0] = 1; $p[1] = 0x18; foreach ($v in $p) { $w.Write([uint64]$v) }
    $w.Write([uint32]1232); $w.Write([uint32]$ctxAt); $dir.Add(@(6, 168, $at))
    $w.Flush()
    $bytes = $ms.ToArray()
    for ($i = 0; $i -lt $dir.Count; $i++) {
        [BitConverter]::GetBytes([uint32]$dir[$i][0]).CopyTo($bytes, $dirAt + 12 * $i)
        [BitConverter]::GetBytes([uint32]$dir[$i][1]).CopyTo($bytes, $dirAt + 12 * $i + 4)
        [BitConverter]::GetBytes([uint32]$dir[$i][2]).CopyTo($bytes, $dirAt + 12 * $i + 8)
    }
    return , $bytes
}

function Invoke-SelfTest {
    $script:stChecks = 0
    $script:stFailures = New-Object System.Collections.Generic.List[string]
    function Check([string] $What, $Expected, $Actual) {
        $script:stChecks++
        if ([string]$Expected -ceq [string]$Actual) { Write-Host "  OK   $What" }
        else { $script:stFailures.Add("$What : expected [$Expected], got [$Actual]"); Write-Host "  FAIL $What - expected [$Expected], got [$Actual]" }
    }
    Write-Host "Read-CrashDump.ps1 -SelfTest under PowerShell $($PSVersionTable.PSVersion). An in-memory dump; no file, no process."
    Import-CrashDumpReader
    $d = New-Object OutlookAICrashDump.Dump(, (New-SelfTestDump))
    Check 'the header: a full-memory minidump with six streams' 'True|6' @((($d.Flags -band 2) -ne 0), $d.Streams.Count)
    Check 'the system: Windows 10.0.26100, x64' '10|0|26100|9' (@($d.OsMajor, $d.OsMinor, $d.OsBuild, $d.Architecture) -join '|')
    Check 'the process id' 4242 $d.ProcessId
    Check 'one module, by its file name' 'wwtest.dll' $d.Modules[0].ShortName
    Check 'its version from VS_FIXEDFILEINFO' '16.0.17932.20036' $d.Modules[0].Version
    Check 'its export, read from the image in memory' 'DoWork' $d.Modules[0].Exports[0].Value
    Check 'its .pdata, both functions' 2 $d.Modules[0].FunctionBegin.Length
    Check 'the exception: an access violation on thread 77' '3221225477|77' (@($d.ExceptionCode, $d.ExceptionThread) -join '|')
    Check 'an address inside an exported function is named by it' 'wwtest.dll!DoWork+0x10' ($d.Describe([uint64]0x7FF600002010))
    Check 'one inside an unexported function is named by module, RVA and function start' 'wwtest.dll+0x220A (fn wwtest.dll+0x2200 +0xA)' ($d.Describe([uint64]0x7FF60000220A))
    Check 'one outside every module is a bare address' '0x0000000000001234' ($d.Describe([uint64]0x1234))
    Check 'RTTI names the object a register points at' 'CTestDocument [wwtest.dll]' ($d.ClassOf([uint64]0x0000020000001000))
    Check 'and nothing for a pointer that is no object' '' ([string]$d.ClassOf([uint64]0x0000009F00000000))
    Check 'a decorated name in a namespace is undecorated' 'NS::CFoo' ([OutlookAICrashDump.Dump]::Undecorate('.?AVCFoo@NS@@'))
    Check 'a template is left decorated rather than guessed' '.?AV?$CList@H@@' ([OutlookAICrashDump.Dump]::Undecorate('.?AV?$CList@H@@'))
    Check 'E8 rel32 is a call' $true ([OutlookAICrashDump.Walker]::FollowsCall([byte[]](0, 0, 0xE8, 1, 2, 3, 4)))
    Check 'FF 15 disp32 is a call' $true ([OutlookAICrashDump.Walker]::FollowsCall([byte[]](0, 0xFF, 0x15, 1, 2, 3, 4)))
    Check 'FF D0 (call rax) is a call' $true ([OutlookAICrashDump.Walker]::FollowsCall([byte[]](0, 0, 0, 0, 0, 0xFF, 0xD0)))
    Check 'a plain mov is not' $false ([OutlookAICrashDump.Walker]::FollowsCall([byte[]](0x48, 0x8B, 0x45, 0x10, 0x90, 0x90, 0x90)))
    $scan = @([OutlookAICrashDump.Walker]::Scan($d, [uint64]0x0000009F00000000, [uint64]0x0000009F00000100, 8))
    Check 'the stack scan finds the return address after the CALL' 'wwtest.dll+0x220A (fn wwtest.dll+0x2200 +0xA)' (@($scan | ForEach-Object { $_.Text }) -join '|')
    $stop = $null
    $frames = @([OutlookAICrashDump.Walker]::Unwind($d, $d.Context(1232 * 0 + $d.ExceptionContextRva, 1232), 8, [ref]$stop))
    Check 'dbghelp unwinds the exception context: the faulting function, then its caller' 'wwtest.dll!DoWork+0x10|wwtest.dll+0x220A (fn wwtest.dll+0x2200 +0xA)' ((@($frames | Select-Object -First 2 | ForEach-Object { $_.Text })) -join '|')
    $report = [OutlookAICrashDump.Report]::Build($d, 4)
    Check 'the report names the fault' $true ($report.Contains('EXCEPTION  access violation (0xC0000005)') -and $report.Contains('WRITING 0x0000000000000018 - a null (or near-null) pointer'))
    Check 'and the class rcx points at' $true $report.Contains('rcx 0x0000020000001000  -> object of class CTestDocument [wwtest.dll]')
    $threw = $false; try { New-Object OutlookAICrashDump.Dump(, [byte[]](1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32)) | Out-Null } catch { $threw = $true }
    Check 'a file that is no minidump is refused' $true $threw

    Write-Host ''
    if ($script:stFailures.Count -gt 0) {
        Write-Host "SelfTest: $($script:stChecks - $script:stFailures.Count) passed, $($script:stFailures.Count) failed."
        $script:stFailures | ForEach-Object { Write-Host "  $_" }
        return 1
    }
    Write-Host "SelfTest: $($script:stChecks) passed, 0 failed."
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

Import-CrashDumpReader
$full = [System.IO.Path]::GetFullPath($Path)
$dump = New-Object OutlookAICrashDump.Dump($full)
try {
    $text = "CRASH DUMP $full`n" + [OutlookAICrashDump.Report]::Build($dump, $OtherThreadFrames)
}
finally { $dump.Dispose() }
if ($OutFile) { [System.IO.File]::WriteAllText($OutFile, $text, (New-Object System.Text.UTF8Encoding($false))) }
Write-Output $text
