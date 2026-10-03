using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;
using OutlookAI.Core.Audit;

namespace OutlookAI.AuditLogStress
{
    /// <summary>
    /// The audit log's multi-process stress writer (Q117, 2026-10-03) - test tooling, never shipped.
    /// <c>T1/AuditLogStressTests</c> starts copies of it, the .NET 10 and the .NET Framework 4.8 build
    /// side by side, against one throwaway log under %TEMP%: many PROCESSES appending to one file is
    /// the case the writer exists for, and threads in one test host cannot stand in for it.
    /// <code>
    /// write &lt;dir&gt; &lt;writer&gt; &lt;lines | -1&gt; &lt;seed&gt; &lt;lock: on | off&gt; &lt;start event&gt;
    ///     waits for the start event, then appends lines 1..lines (-1: until killed) through
    ///     AuditLog.AppendTo - with the writer lock, or with it deliberately skipped - and prints its
    ///     append latencies.
    /// tear &lt;dir&gt; &lt;ready event&gt;
    ///     takes the writer lock, appends the first half of a line and stops there, holding the lock:
    ///     a writer that dies mid-write, once it is killed.
    /// bench &lt;dir&gt; &lt;count&gt; &lt;bytes&gt;
    ///     the cost of FILE_FLAG_WRITE_THROUGH: open-append-close latencies with and without it, and
    ///     the product's whole append.
    /// </code>
    /// It refuses any directory outside %TEMP%, so it can never be pointed at the real log.
    /// </summary>
    internal static class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                if (args.Length == 7 && args[0] == "write")
                {
                    return Write(args[1], args[2], long.Parse(args[3], CultureInfo.InvariantCulture),
                        int.Parse(args[4], CultureInfo.InvariantCulture), args[5] == "on", args[6]);
                }

                if (args.Length == 3 && args[0] == "tear")
                {
                    return Tear(args[1], args[2]);
                }

                if (args.Length == 4 && args[0] == "bench")
                {
                    return Bench(args[1], int.Parse(args[2], CultureInfo.InvariantCulture), int.Parse(args[3], CultureInfo.InvariantCulture));
                }

                Console.Error.WriteLine("usage: write <dir> <writer> <lines|-1> <seed> <on|off> <event> | tear <dir> <event> | bench <dir> <count> <bytes>");
                return 2;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("failed: " + ex);
                return 1;
            }
        }

        private static int Write(string dir, string writer, long lines, int seed, bool useLock, string startEvent)
        {
            RequireThrowaway(dir);
            using (EventWaitHandle start = EventWaitHandle.OpenExisting(startEvent))
            {
                Say("ready " + writer + " pid=" + AuditLog.WriterPid.ToString(CultureInfo.InvariantCulture)
                    + " run=" + AuditLog.WriterRun + " runtime=" + RuntimeInformation.FrameworkDescription);
                start.WaitOne();
            }

            List<long> ticks = new List<long>();
            Stopwatch total = Stopwatch.StartNew();
            long n = 0;
            while (lines < 0 || n < lines)
            {
                n++;
                (string, string?)[] fields =
                {
                    ("writer", writer),
                    ("n", n.ToString(CultureInfo.InvariantCulture)),
                    ("data", StressLines.Payload(seed, writer, n)),
                };
                long before = Stopwatch.GetTimestamp();
                AuditLog.AppendTo(dir, StressLines.Operation, fields, useLock);
                if (ticks.Count < 100000)
                {
                    ticks.Add(Stopwatch.GetTimestamp() - before);
                }
            }

            Say("done " + writer + " lines=" + n.ToString(CultureInfo.InvariantCulture)
                + " ms=" + total.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " " + Percentiles(ticks)
                + " lockTimeouts=" + AuditLog.WriterLockTimeouts.ToString(CultureInfo.InvariantCulture)
                + " lockAbandoned=" + AuditLog.WriterLockAbandoned.ToString(CultureInfo.InvariantCulture));
            return 0;
        }

        private static int Tear(string dir, string readyEvent)
        {
            RequireThrowaway(dir);
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, AuditLog.LogFileName);

            // Held to the end: killed holding it, the mutex is ABANDONED, which is what the next writer sees.
            AuditLog.AuditWriterLock held = AuditLog.AuditWriterLock.Acquire(path);
            if (!held.Owned)
            {
                throw new InvalidOperationException("the writer lock could not be taken: " + held.LineState);
            }

            string line = AuditLog.FormatSealedLine(
                DateTime.UtcNow,
                StressLines.Operation,
                new (string, string?)[] { ("writer", "tear"), ("n", "1"), ("data", new string('x', 4000)) },
                AuditLog.WriterPid,
                AuditLog.WriterRun,
                1,
                null);
            byte[] bytes = new UTF8Encoding(false).GetBytes(line);
            if (!AuditLogFile.TryOpenForAppend(path, AuditLog.WriteThrough, out SafeFileHandle? handle, out int error))
            {
                throw AuditLogFile.ErrorFor(error, path);
            }

            using (handle!)
            {
                // Half the line - well inside the 4,000-character value - and no terminator.
                AuditLogFile.Append(handle!, bytes, bytes.Length / 2);
            }

            using (EventWaitHandle ready = EventWaitHandle.OpenExisting(readyEvent))
            {
                ready.Set();
            }

            Say("torn pid=" + AuditLog.WriterPid.ToString(CultureInfo.InvariantCulture) + " run=" + AuditLog.WriterRun);
            Thread.Sleep(Timeout.Infinite);
            GC.KeepAlive(held);
            return 0;
        }

        private static int Bench(string dir, int count, int bytes)
        {
            RequireThrowaway(dir);
            Directory.CreateDirectory(dir);
            byte[] line = Encoding.ASCII.GetBytes(new string('b', Math.Max(1, bytes - 1)) + "\n");
            foreach (bool writeThrough in new[] { false, true })
            {
                string path = Path.Combine(dir, writeThrough ? "bench-writethrough.log" : "bench-cached.log");
                List<long> ticks = new List<long>(count);
                for (int i = 0; i < count; i++)
                {
                    long before = Stopwatch.GetTimestamp();
                    if (!AuditLogFile.TryOpenForAppend(path, writeThrough, out SafeFileHandle? handle, out int error))
                    {
                        throw AuditLogFile.ErrorFor(error, path);
                    }

                    using (handle!)
                    {
                        AuditLogFile.Append(handle!, line, line.Length);
                    }

                    ticks.Add(Stopwatch.GetTimestamp() - before);
                }

                Say("bench open-append-close writeThrough=" + (writeThrough ? "on" : "off") + " count="
                    + count.ToString(CultureInfo.InvariantCulture) + " bytes=" + line.Length.ToString(CultureInfo.InvariantCulture)
                    + " " + Percentiles(ticks));
            }

            string payload = new string('p', Math.Max(1, bytes - 200));
            List<long> product = new List<long>(count);
            for (int i = 0; i < count; i++)
            {
                long before = Stopwatch.GetTimestamp();
                AuditLog.AppendTo(dir, "bench_append", new (string, string?)[] { ("data", payload) });
                product.Add(Stopwatch.GetTimestamp() - before);
            }

            Say("bench product-append writeThrough=" + (AuditLog.WriteThrough ? "on" : "off") + " count="
                + count.ToString(CultureInfo.InvariantCulture) + " " + Percentiles(product));
            return 0;
        }

        private static void RequireThrowaway(string dir)
        {
            string full = Path.GetFullPath(dir).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string temp = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!full.StartsWith(temp, StringComparison.OrdinalIgnoreCase) || full.Length == temp.Length)
            {
                throw new ArgumentException("refusing '" + dir + "': the stress writer writes only below " + temp);
            }
        }

        private static string Percentiles(List<long> ticks)
        {
            if (ticks.Count == 0)
            {
                return "p50us=0 p90us=0 p99us=0 maxus=0";
            }

            long[] sorted = ticks.ToArray();
            Array.Sort(sorted);
            return "p50us=" + Micros(sorted[sorted.Length / 2]) + " p90us=" + Micros(sorted[(int)(sorted.Length * 0.9)])
                + " p99us=" + Micros(sorted[Math.Min(sorted.Length - 1, (int)(sorted.Length * 0.99))])
                + " maxus=" + Micros(sorted[sorted.Length - 1]);
        }

        private static string Micros(long ticks)
        {
            return (ticks * 1000000L / Stopwatch.Frequency).ToString(CultureInfo.InvariantCulture);
        }

        private static void Say(string line)
        {
            Console.Out.WriteLine(line);
            Console.Out.Flush();
        }
    }
}
