using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using OutlookAI.AuditLogStress;
using OutlookAI.Core.Audit;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The audit log under many writer PROCESSES at once (Q117, 2026-10-03) - the case the writer exists
/// for, which threads in this test host cannot stand in for. Every run starts copies of
/// <c>OutlookAI.AuditLogStress</c>, the .NET 10 and the .NET Framework 4.8 build side by side, against
/// one throwaway log under %TEMP%; lines are up to 64 KB, and every line's content is recomputed here
/// from (seed, writer, n) by the same <see cref="StressLines"/> the writers used, then compared byte
/// for byte after the read-back.
/// <para>
/// What is asserted, run by run: with the lock, every line whole, each writer's lines complete and in
/// its own order, and the file in time order; with the lock DISABLED, every line still whole and none
/// lost - the measurement of whether two append-only writes can interleave their bytes, which Windows
/// does not document; and with a writer killed in the middle of an append (holding the lock, half a
/// line written), exactly one fragment, detected, the next line on a line of its own, the lock passed
/// on as abandoned and recorded, and nothing else lost. The numbers each run measured go to the test
/// output, which the build VM's TRX file keeps.
/// </para>
/// </summary>
public sealed class AuditLogStressTests : IDisposable
{
    private const int Seed = 117;
    private const string Net10 = "net10.0-windows";
    private const string Net48 = "net48";
    private static readonly TimeSpan StartTimeout = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan RunTimeout = TimeSpan.FromSeconds(240);

    private readonly ITestOutputHelper _output;
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogStress-" + Guid.NewGuid().ToString("N"));
    private readonly List<Process> _processes = new();

    public AuditLogStressTests(ITestOutputHelper output)
    {
        _output = output;
        Directory.CreateDirectory(_dir);
    }

    private string LogPath => Path.Combine(_dir, AuditLog.LogFileName);

    public void Dispose()
    {
        foreach (Process process in _processes)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill();
                    process.WaitForExit(10_000);
                }
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception)
            {
            }

            process.Dispose();
        }

        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
        {
        }
    }

    [Fact]
    public void EightWriterProcesses_OnBothRuntimes_WithTheLock_KeepEveryLineWhole_AndTheFileInTimeOrder()
    {
        Writer[] writers = StartWriters("lockon", lockOn: true, net10: 6, net48: 2, lines: 250);
        Stopwatch elapsed = Go(writers);
        WaitAll(writers);
        elapsed.Stop();

        List<AuditLogEntry> lines = VerifyEveryLine(writers, out AuditLogScan scan);
        int inversions = CountTimeInversions(lines, lockedOnly: false);
        Report("lock on", writers, scan, elapsed, inversions);

        // Ordered: the timestamp is taken while the lock is held, so the file is in time order.
        Assert.Equal(0, scan.LinesWithoutWriterLock);
        Assert.Equal(0, inversions);
    }

    [Fact]
    public void EightWriterProcesses_WithTheLockDisabled_StillOverwriteNothing_AndInterleaveNothing()
    {
        // The measurement Q117 asked for. Nothing orders these appends: the only thing between eight
        // processes writing lines of up to 64 KB to one file is FILE_APPEND_DATA. Every line must still
        // arrive whole - a byte of one line inside another would fail its checksum - and none may be
        // lost. Order is NOT asserted (that is what the lock is for); how much of it was lost is reported.
        Writer[] writers = StartWriters("lockoff", lockOn: false, net10: 6, net48: 2, lines: 250);
        Stopwatch elapsed = Go(writers);
        WaitAll(writers);
        elapsed.Stop();

        List<AuditLogEntry> lines = VerifyEveryLine(writers, out AuditLogScan scan);
        int inversions = CountTimeInversions(lines, lockedOnly: false);
        Report("lock DISABLED", writers, scan, elapsed, inversions);

        Assert.Equal(lines.Count, scan.LinesWithoutWriterLock);
        Assert.All(lines, line => Assert.Equal(AuditLog.LockDisabled, line.WriterLock));
    }

    [Fact]
    public void AWriterKilledMidAppend_LeavesOneDetectedFragment_TheLockIsPassedOn_AndNothingElseIsLost()
    {
        // 1. A writer that takes the lock, writes half a line and is killed there - the crash case.
        string readyName = EventName("torn");
        Process tear;
        using (EventWaitHandle torn = new(false, EventResetMode.ManualReset, readyName))
        {
            tear = Launch(Net10, "tear", Quote(_dir), readyName);
            Assert.True(torn.WaitOne(StartTimeout), "the tearing writer never reported its half line");
        }

        // 2. Four ordinary writers and one that appends until it is killed - all of them now waiting on
        //    the lock the tearing writer holds.
        Writer[] finite = StartWriters("killrun", lockOn: true, net10: 3, net48: 1, lines: 150);
        Writer endless = StartWriter("endless", Net10, lockOn: true, lines: -1, finite[0].StartEvent);
        Writer[] all = finite.Concat(new[] { endless }).ToArray();
        WaitReady(all);
        Stopwatch elapsed = Go(all);
        Thread.Sleep(300);
        tear.Kill();
        Assert.True(tear.WaitForExit(30_000), "the tearing writer did not die");

        // 3. Let the ordinary writers finish, make sure the endless one got well under way, and kill it
        //    too - wherever it is in its append.
        WaitAll(finite);
        DateTime deadline = DateTime.UtcNow + RunTimeout;
        while (CountLinesOf("endless") < 60 && DateTime.UtcNow < deadline)
        {
            Thread.Sleep(100);
        }

        endless.Process.Kill();
        Assert.True(endless.Process.WaitForExit(30_000), "the endless writer did not die");

        // 4. One more writer after both deaths: whatever they left, its line must start clean.
        Writer after = StartWriter("after", Net48, lockOn: true, lines: 1, EventName("after"));
        WaitReady(new[] { after });
        Go(new[] { after });
        WaitAll(new[] { after });
        elapsed.Stop();

        List<AuditLogEntry> lines = VerifyEveryLine(finite.Concat(new[] { endless, after }).ToArray(), out AuditLogScan scan);
        Report("writer killed mid-append", all.Concat(new[] { after }).ToArray(), scan, elapsed, CountTimeInversions(lines, lockedOnly: true));

        // The deliberate fragment, and nothing else: it does not parse, so it is malformed (not a
        // checksum failure), and the line after it was written on a line of its own.
        Assert.Equal(1, scan.MalformedLines);
        Assert.Equal(0, scan.ChecksumFailures);
        byte[] head = ReadHead(64 * 1024);
        int firstBreak = Array.IndexOf(head, (byte)'\n');
        Assert.True(firstBreak > 0, "no line break after the fragment");
        Assert.DoesNotContain(" crc=\"", Encoding.UTF8.GetString(head, 0, firstBreak), StringComparison.Ordinal);
        Assert.Equal("ts=", Encoding.ASCII.GetString(head, firstBreak + 1, 3));

        // The lock the tearing writer died holding was passed on, and the line that inherited it says
        // so - and it is the first whole line in the file.
        Assert.Equal(AuditLog.LockAbandoned, lines[0].WriterLock);
        Assert.InRange(scan.LinesAfterAbandonedLock, 1, 2);
        _output.WriteLine("abandoned-lock lines: " + scan.LinesAfterAbandonedLock.ToString(CultureInfo.InvariantCulture)
            + " (2 means the endless writer was killed holding the lock too)");

        // The endless writer's lines stop where it was killed, with no hole before that point.
        long endlessLines = lines.Count(l => Field(l, "writer") == "endless");
        Assert.True(endlessLines >= 60, "the endless writer wrote only " + endlessLines.ToString(CultureInfo.InvariantCulture));
        Assert.Equal(0, scan.MissingLines);
    }

    [Fact]
    public void WriteThroughCost_IsMeasured_OnBothRuntimes()
    {
        // FILE_FLAG_WRITE_THROUGH was switched on by measurement, not by assumption: this prints the
        // open-append-close latency with and without it, and the product's whole append, per runtime.
        foreach (string runtime in new[] { Net10, Net48 })
        {
            string dir = Path.Combine(_dir, "bench-" + runtime);
            Process bench = Launch(runtime, "bench", Quote(dir), "300", "400");
            List<string> output = new();
            bench.OutputDataReceived += (_, e) =>
            {
                if (e.Data != null)
                {
                    lock (output)
                    {
                        output.Add(e.Data);
                    }
                }
            };
            bench.BeginOutputReadLine();
            Assert.True(bench.WaitForExit((int)RunTimeout.TotalMilliseconds), "the benchmark did not finish");
            bench.WaitForExit();
            Assert.Equal(0, bench.ExitCode);
            lock (output)
            {
                Assert.Equal(3, output.Count(l => l.StartsWith("bench ", StringComparison.Ordinal)));
                foreach (string line in output)
                {
                    _output.WriteLine(runtime + ": " + line);
                }
            }
        }
    }

    // ================================================================== running writers

    private Writer[] StartWriters(string prefix, bool lockOn, int net10, int net48, long lines)
    {
        string startEvent = EventName(prefix);
        List<Writer> writers = new();
        for (int i = 0; i < net10 + net48; i++)
        {
            string runtime = i < net10 ? Net10 : Net48;
            string name = prefix + i.ToString("00", CultureInfo.InvariantCulture) + (runtime == Net48 ? "fx" : "core");
            writers.Add(StartWriter(name, runtime, lockOn, lines, startEvent));
        }

        WaitReady(writers);
        return writers.ToArray();
    }

    private Writer StartWriter(string name, string runtime, bool lockOn, long lines, string startEvent)
    {
        // The event exists before any writer opens it, and lives as long as the writer.
        EventWaitHandle start = new(false, EventResetMode.ManualReset, startEvent, out bool _);
        Process process = Launch(runtime, "write", Quote(_dir), name, lines.ToString(CultureInfo.InvariantCulture),
            Seed.ToString(CultureInfo.InvariantCulture), lockOn ? "on" : "off", startEvent);
        Writer writer = new(name, runtime, lines, process, start, startEvent);
        process.OutputDataReceived += (_, e) => writer.Receive(e.Data);
        process.BeginOutputReadLine();
        return writer;
    }

    private static void WaitReady(IEnumerable<Writer> writers)
    {
        foreach (Writer writer in writers)
        {
            Assert.True(writer.Ready.Wait(StartTimeout), writer.Name + " never became ready: " + writer.Transcript());
        }
    }

    private static Stopwatch Go(IEnumerable<Writer> writers)
    {
        foreach (Writer writer in writers)
        {
            writer.Start.Set();
        }

        return Stopwatch.StartNew();
    }

    private static void WaitAll(IEnumerable<Writer> writers)
    {
        foreach (Writer writer in writers)
        {
            Assert.True(writer.Process.WaitForExit((int)RunTimeout.TotalMilliseconds), writer.Name + " did not finish");
            writer.Process.WaitForExit();
            Assert.True(writer.Process.ExitCode == 0, writer.Name + " failed: " + writer.Transcript());
        }
    }

    private Process Launch(string runtime, params string[] arguments)
    {
        string exe = Path.Combine(StressBinDir, runtime, "OutlookAI.AuditLogStress.exe");
        Assert.True(File.Exists(exe), "the stress writer is not built at '" + exe + "'");
        ProcessStartInfo info = new(exe, string.Join(" ", arguments))
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = false,
        };
        Process process = Process.Start(info) ?? throw new InvalidOperationException("could not start " + exe);
        _processes.Add(process);
        return process;
    }

    private static string StressBinDir =>
        typeof(AuditLogStressTests).Assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
            .Single(a => a.Key == "AuditLogStressBinDir").Value!;

    private static string EventName(string what) => @"Local\OutlookAI.AuditLogStress." + what + "." + Guid.NewGuid().ToString("N");

    private static string Quote(string value) => "\"" + value + "\"";

    // ================================================================== reading back

    /// <summary>
    /// Pages through the whole log the way audit_log does, checks the pages add up to the file, and
    /// checks every line against what its writer must have written. Returns the lines in FILE order.
    /// </summary>
    private List<AuditLogEntry> VerifyEveryLine(Writer[] writers, out AuditLogScan scan)
    {
        AuditLogFilter all = new(null, null, null, null);
        List<AuditLogEntry> newestFirst = new();
        AuditLogPosition? next = null;
        AuditLogScan? first = null;
        do
        {
            AuditLogScan page = AuditLogReader.Read(LogPath, all, 100, next);
            Assert.NotEqual(AuditLogResumeStatus.PositionNotFound, page.Resume);
            first ??= page;
            newestFirst.AddRange(page.Entries);
            next = page.NextPosition;
        }
        while (next != null);

        scan = first!;
        Assert.Equal(scan.Matched, newestFirst.Count);
        Assert.Equal(0, scan.UnverifiedLines);
        Assert.Equal(0, scan.DuplicateLines);
        Assert.False(scan.IncompleteLastLine);

        List<AuditLogEntry> lines = Enumerable.Reverse(newestFirst).ToList();
        Dictionary<string, Writer> byName = writers.ToDictionary(w => w.Name);
        Dictionary<string, long> lastN = new();
        foreach (AuditLogEntry line in lines)
        {
            Assert.Equal(StressLines.Operation, line.Operation);
            Assert.True(line.Checksummed);
            string name = Field(line, "writer");
            Writer writer = byName[name];
            long n = long.Parse(Field(line, "n"), CultureInfo.InvariantCulture);

            // Byte for byte what this writer wrote as its line n - and its n-th line in this file.
            Assert.Equal(StressLines.Payload(Seed, name, n), Field(line, "data"));
            Assert.Equal(n, line.Seq);
            Assert.Equal(writer.Pid, line.Pid);
            Assert.True(lastN.TryGetValue(name, out long previous) ? n == previous + 1 : n == 1,
                name + " line " + n.ToString(CultureInfo.InvariantCulture) + " is out of its writer's order");
            lastN[name] = n;
        }

        foreach (Writer writer in writers.Where(w => w.Lines > 0))
        {
            Assert.Equal(writer.Lines, lastN.TryGetValue(writer.Name, out long n) ? n : 0);
        }

        Assert.Equal(lines.Count, scan.VerifiedLines);
        Assert.Equal(0, scan.MissingLines);
        Assert.Equal(lastN.Count, scan.Writers);
        return lines;
    }

    private long CountLinesOf(string writer)
    {
        if (!File.Exists(LogPath))
        {
            return 0;
        }

        AuditLogScan scan = AuditLogReader.Read(LogPath, new AuditLogFilter(null, null, null, null), 100_000);
        return scan.Entries.Count(e => Field(e, "writer") == writer);
    }

    private static int CountTimeInversions(List<AuditLogEntry> lines, bool lockedOnly)
    {
        int inversions = 0;
        DateTime previous = DateTime.MinValue;
        foreach (AuditLogEntry line in lines)
        {
            if (lockedOnly && line.WriterLock != null && line.WriterLock != AuditLog.LockAbandoned)
            {
                continue;
            }

            if (line.TimestampUtc < previous)
            {
                inversions++;
            }
            else
            {
                previous = line.TimestampUtc;
            }
        }

        return inversions;
    }

    private byte[] ReadHead(int max)
    {
        using FileStream stream = new(LogPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        byte[] buffer = new byte[(int)Math.Min(max, stream.Length)];
        int read = 0;
        while (read < buffer.Length)
        {
            int n = stream.Read(buffer, read, buffer.Length - read);
            if (n == 0)
            {
                break;
            }

            read += n;
        }

        return buffer;
    }

    private void Report(string run, Writer[] writers, AuditLogScan scan, Stopwatch elapsed, int inversions)
    {
        long bytes = new FileInfo(LogPath).Length;
        _output.WriteLine(run + ": " + writers.Length.ToString(CultureInfo.InvariantCulture) + " writer processes ("
            + writers.Count(w => w.Runtime == Net10).ToString(CultureInfo.InvariantCulture) + " .NET 10, "
            + writers.Count(w => w.Runtime == Net48).ToString(CultureInfo.InvariantCulture) + " .NET Framework 4.8), "
            + scan.LinesScanned.ToString(CultureInfo.InvariantCulture) + " lines, "
            + bytes.ToString(CultureInfo.InvariantCulture) + " bytes, "
            + elapsed.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms");
        _output.WriteLine("  verified=" + scan.VerifiedLines.ToString(CultureInfo.InvariantCulture)
            + " malformed=" + scan.MalformedLines.ToString(CultureInfo.InvariantCulture)
            + " checksumFailures=" + scan.ChecksumFailures.ToString(CultureInfo.InvariantCulture)
            + " missing=" + scan.MissingLines.ToString(CultureInfo.InvariantCulture)
            + " duplicates=" + scan.DuplicateLines.ToString(CultureInfo.InvariantCulture)
            + " withoutLock=" + scan.LinesWithoutWriterLock.ToString(CultureInfo.InvariantCulture)
            + " afterAbandonedLock=" + scan.LinesAfterAbandonedLock.ToString(CultureInfo.InvariantCulture)
            + " writers=" + scan.Writers.ToString(CultureInfo.InvariantCulture)
            + " timeInversions=" + inversions.ToString(CultureInfo.InvariantCulture));
        foreach (Writer writer in writers)
        {
            _output.WriteLine("  " + writer.Runtime + " " + writer.Transcript());
        }
    }

    private static string Field(AuditLogEntry entry, string key) => entry.Fields.Single(f => f.Key == key).Value;

    private sealed class Writer
    {
        private readonly List<string> _lines = new();

        internal Writer(string name, string runtime, long lines, Process process, EventWaitHandle start, string startEvent)
        {
            Name = name;
            Runtime = runtime;
            Lines = lines;
            Process = process;
            Start = start;
            StartEvent = startEvent;
        }

        internal string Name { get; }

        internal string Runtime { get; }

        internal long Lines { get; }

        internal Process Process { get; }

        internal EventWaitHandle Start { get; }

        internal string StartEvent { get; }

        internal int Pid => Process.Id;

        internal ManualResetEventSlim Ready { get; } = new(false);

        internal void Receive(string? line)
        {
            if (line == null)
            {
                return;
            }

            lock (_lines)
            {
                _lines.Add(line);
            }

            if (line.StartsWith("ready ", StringComparison.Ordinal))
            {
                Ready.Set();
            }
        }

        internal string Transcript()
        {
            lock (_lines)
            {
                return string.Join(" | ", _lines);
            }
        }
    }
}
