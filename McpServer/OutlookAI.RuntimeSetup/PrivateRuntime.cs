using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using OutlookAI.Core.AddIn;

namespace OutlookAI.RuntimeSetup
{
    /// <summary>
    /// THE ONE CLASS THAT FETCHES, CHECKS AND UNPACKS THE MAIL SERVER'S PRIVATE .NET RUNTIME (Q172,
    /// decided by the maintainer 2026-11-18). Every network call, file write, folder move and delete
    /// this helper makes is in here, so there is one place to vet them (stream S's rule that risky
    /// calls live in named classes).
    /// <list type="number">
    /// <item>The package: Microsoft's own win-x64 runtime zip for <see cref="ServerRuntime.Version"/>,
    /// from its fixed address (<see cref="PinnedUrl"/>) - or, when one lies beside the installer under
    /// Microsoft's own name, that file - copied into a work folder beside the runtime folder, never
    /// more than one byte past the pinned length.</item>
    /// <item>The check, before anything is unpacked: the length and the SHA-512 Microsoft's release
    /// metadata publishes for that file (<see cref="PinnedLength"/>, <see cref="PinnedSha512"/>).</item>
    /// <item>The unpack, into the work folder: every entry's place is worked out first, against the
    /// work folder's full path, and a package with one entry that would land outside it is refused
    /// whole, with nothing written (zip-slip). Then <see cref="ServerRuntime.IsComplete"/> must hold
    /// for what was unpacked.</item>
    /// <item>The swap: the old runtime folder, when there is one, is renamed into the work folder, and
    /// the new one renamed into its place, so the runtime folder is the old copy or the whole new one,
    /// never half of either. The work folder is deleted at the end, and a work folder an earlier run
    /// left behind at the start.</item>
    /// </list>
    /// It writes and deletes nowhere else: the runtime folder (%LOCALAPPDATA%\OutlookAI\dotnet), the work
    /// folders beside it (<c>dotnet.setup-&lt;time&gt;</c>), and its log beside them
    /// (<c>dotnet-setup.log</c>). It starts no program.
    /// </summary>
    public sealed class PrivateRuntime : IDisposable
    {
        /// <summary>
        /// Microsoft's runtime package for <see cref="ServerRuntime.Version"/>, as
        /// https://builds.dotnet.microsoft.com/dotnet/release-metadata/11.0/releases.json lists it: release
        /// 11.0.0 (2026-11-10), <c>runtime.files[]</c>, the entry named <c>dotnet-runtime-win-x64.zip</c>.
        /// Read 2026-11-18; its length is that address's Content-Length the same day. A new version is a
        /// deliberate edit of all three from that file, never a hash computed from a download.
        /// </summary>
        public const string PinnedUrl = "https://builds.dotnet.microsoft.com/dotnet/Runtime/" + ServerRuntime.Version
            + "/" + PackageFileName;

        /// <summary>The package's file name, as Microsoft publishes it - and as setup looks for it beside the installer.</summary>
        public const string PackageFileName = "dotnet-runtime-" + ServerRuntime.Version + "-win-x64.zip";

        /// <summary>The package's length in bytes (<see cref="PinnedUrl"/>).</summary>
        public const long PinnedLength = 39856058;

        /// <summary>The package's SHA-512, as the release metadata prints it (<see cref="PinnedUrl"/>).</summary>
        public const string PinnedSha512 = "48d200527a30d2b21e361d1e91c9b71df736e7da76383d7f7b4e4102a1bc78d1e4690c6540f308598ed036ca30e92de90ac80628f4ca498c55e826d86b1fe313";

        /// <summary>The download's limit before the response starts, and the whole run's.</summary>
        private static readonly TimeSpan ConnectTimeout = TimeSpan.FromMinutes(10);
        private static readonly TimeSpan RunTimeout = TimeSpan.FromMinutes(20);

        /// <summary>What setup reads from the exit code. The log says which of the two ways 0 came about.</summary>
        public enum ExitCode
        {
            Ready = 0,
            BadArguments = 2,
            DownloadFailed = 10,
            PackageRejected = 11,
            UnpackFailed = 12,
            SwapFailed = 13,
            Failed = 14,
        }

        /// <summary>What <see cref="Ensure"/> did.</summary>
        public enum Outcome
        {
            AlreadyInstalled,
            Installed,
            DownloadFailed,
            PackageRejected,
            UnpackFailed,
            SwapFailed,
        }

        /// <summary>A package: where it comes from, and the length and SHA-512 it must have.</summary>
        public sealed class Package
        {
            public Package(Uri url, long length, string sha512)
            {
                Url = url;
                Length = length;
                Sha512 = sha512;
            }

            public Uri Url { get; }

            public long Length { get; }

            public string Sha512 { get; }

            /// <summary>The one setup installs.</summary>
            public static Package Pinned { get; } = new Package(new Uri(PinnedUrl), PinnedLength, PinnedSha512);
        }

        private readonly string _root;
        private readonly Package _package;
        private readonly Func<Uri, CancellationToken, Stream> _download;
        private readonly StreamWriter _log;

        /// <param name="root">The runtime folder: <see cref="ServerRuntime.DefaultRoot"/> in setup, a scratch folder in a test.</param>
        /// <param name="package">What to fetch, and what it must hash to.</param>
        /// <param name="download">Opens the package's bytes: an HTTP GET in setup, a file in a test.</param>
        public PrivateRuntime(string root, Package package, Func<Uri, CancellationToken, Stream> download)
        {
            _root = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            _package = package;
            _download = download;
            Directory.CreateDirectory(Path.GetDirectoryName(_root)!);
            _log = new StreamWriter(LogPath(_root), append: false, new UTF8Encoding(false)) { AutoFlush = true };
        }

        /// <summary>The log of the last run, beside the runtime folder: <c>dotnet-setup.log</c>.</summary>
        public static string LogPath(string root) => root + "-setup.log";

        /// <summary>The prefix of a work folder's name, beside the runtime folder: <c>dotnet.setup-</c>.</summary>
        public static string WorkFolderPrefix(string root) => Path.GetFileName(root) + ".setup-";

        /// <summary>
        /// What setup runs: the runtime into <see cref="ServerRuntime.DefaultRoot"/>, from the package
        /// in <paramref name="from"/> (<see cref="PackageFileName"/>) when that file exists and from the
        /// network otherwise. Never throws: setup is waiting, and an unhandled exception could leave a
        /// crash dialog it waits on.
        /// </summary>
        public static ExitCode Run(string? from) => Run(from, ServerRuntime.DefaultRoot());

        /// <summary>
        /// <see cref="Run(string)"/> into another runtime folder: the whole of what setup gets - the real
        /// download, the checks, the unpack and the swap - for a proof against a scratch folder (Q172
        /// phase 4). Setup's helper never names one: its only argument is <c>--from</c>.
        /// </summary>
        public static ExitCode Run(string? from, string root)
        {
            PrivateRuntime? runtime = null;
            try
            {
                using var http = new HttpClient { Timeout = ConnectTimeout };
                runtime = new PrivateRuntime(root, Package.Pinned, HttpDownload(http));
                using var cancel = new CancellationTokenSource(RunTimeout);
                string? staged = string.IsNullOrEmpty(from) ? null : Path.Combine(from, PackageFileName);
                Outcome outcome = runtime.Ensure(staged, cancel.Token);
                runtime.Log("outcome: " + outcome);
                return ToExitCode(outcome);
            }
            catch (Exception ex)
            {
                runtime?.Log("failed: " + ex.GetType().Name + ": " + ex.Message);
                return ExitCode.Failed;
            }
            finally
            {
                runtime?.Dispose();
            }
        }

        /// <summary>The exit code setup reads for an outcome.</summary>
        public static ExitCode ToExitCode(Outcome outcome)
        {
            switch (outcome)
            {
                case Outcome.AlreadyInstalled:
                case Outcome.Installed:
                    return ExitCode.Ready;
                case Outcome.DownloadFailed:
                    return ExitCode.DownloadFailed;
                case Outcome.PackageRejected:
                    return ExitCode.PackageRejected;
                case Outcome.UnpackFailed:
                    return ExitCode.UnpackFailed;
                case Outcome.SwapFailed:
                    return ExitCode.SwapFailed;
                default:
                    return ExitCode.Failed;
            }
        }

        /// <summary>
        /// Leaves a complete runtime in the runtime folder, or the folder as it was. Does nothing more
        /// when the folder already holds <see cref="ServerRuntime.Version"/>.
        /// </summary>
        /// <param name="stagedPackage">A package on disk to use instead of the network, when it exists.</param>
        public Outcome Ensure(string? stagedPackage, CancellationToken cancel)
        {
            Log("OutlookAI.RuntimeSetup: .NET " + ServerRuntime.Version + " into " + _root);
            RemoveWorkFolders();
            if (ServerRuntime.IsComplete(_root))
            {
                Log("already installed");
                return Outcome.AlreadyInstalled;
            }

            string work = _root + ".setup-" + DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
            string package = Path.Combine(work, "package.zip");
            string unpacked = Path.Combine(work, "runtime");
            string retired = Path.Combine(work, "retired");
            try
            {
                Directory.CreateDirectory(work);
                Outcome? failed = Fetch(stagedPackage, package, cancel) ?? Check(package) ?? Unpack(package, unpacked);
                return failed ?? Swap(unpacked, retired);
            }
            finally
            {
                DeleteTree(work);
            }
        }

        public void Dispose() => _log.Dispose();

        /// <summary>
        /// Where an archive entry named <paramref name="entryName"/> lands under <paramref name="target"/>,
        /// or null when its name would put it anywhere else: outside the folder by <c>..</c> or a rooted
        /// name, on the folder itself, or into a stream of a file (a colon). Checked against the folder's
        /// full path, for every entry, before any is written.
        /// </summary>
        public static string? PlaceOf(string target, string entryName)
        {
            if (string.IsNullOrEmpty(entryName) || entryName.IndexOf(':') >= 0)
                return null;

            try
            {
                string targetFull = Path.GetFullPath(target).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;
                string place = Path.GetFullPath(Path.Combine(targetFull, entryName));
                if (!place.StartsWith(targetFull, StringComparison.OrdinalIgnoreCase))
                    return null;
                return place.Length > targetFull.Length ? place : null;
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
            {
                return null;
            }
        }

        private Outcome? Fetch(string? stagedPackage, string package, CancellationToken cancel)
        {
            try
            {
                if (!string.IsNullOrEmpty(stagedPackage) && File.Exists(stagedPackage))
                {
                    Log("package from " + stagedPackage);
                    File.Copy(stagedPackage, package);
                    return null;
                }

                if (!string.IsNullOrEmpty(stagedPackage))
                    Log("no package at " + stagedPackage);
                Log("downloading " + _package.Url);
                using (Stream source = _download(_package.Url, cancel))
                using (var target = new FileStream(package, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    CopyAtMost(source, target, _package.Length + 1, cancel);
                }

                return null;
            }
            catch (Exception ex) when (ex is IOException || ex is HttpRequestException || ex is OperationCanceledException
                                       || ex is UnauthorizedAccessException || ex is InvalidOperationException)
            {
                Log("download failed: " + ex.GetType().Name + ": " + ex.Message);
                return Outcome.DownloadFailed;
            }
        }

        private Outcome? Check(string package)
        {
            long length = new FileInfo(package).Length;
            if (length != _package.Length)
            {
                Log("rejected: " + length.ToString(CultureInfo.InvariantCulture) + " bytes, where "
                    + _package.Length.ToString(CultureInfo.InvariantCulture) + " are pinned");
                return Outcome.PackageRejected;
            }

            string actual = Sha512Of(package);
            if (!string.Equals(actual, _package.Sha512, StringComparison.OrdinalIgnoreCase))
            {
                Log("rejected: SHA-512 " + actual + ", where " + _package.Sha512 + " is pinned");
                return Outcome.PackageRejected;
            }

            Log("package checked: " + length.ToString(CultureInfo.InvariantCulture) + " bytes, SHA-512 as pinned");
            return null;
        }

        private Outcome? Unpack(string package, string target)
        {
            try
            {
                using (ZipArchive zip = ZipFile.OpenRead(package))
                {
                    var places = new List<KeyValuePair<ZipArchiveEntry, string>>();
                    foreach (ZipArchiveEntry entry in zip.Entries)
                    {
                        string? place = PlaceOf(target, entry.FullName);
                        if (place == null)
                        {
                            Log("rejected: the entry '" + entry.FullName + "' would land outside " + target + "; nothing was unpacked");
                            return Outcome.PackageRejected;
                        }

                        places.Add(new KeyValuePair<ZipArchiveEntry, string>(entry, place));
                    }

                    foreach (KeyValuePair<ZipArchiveEntry, string> item in places)
                    {
                        if (item.Key.FullName.EndsWith("/", StringComparison.Ordinal) || item.Key.FullName.EndsWith("\\", StringComparison.Ordinal))
                        {
                            Directory.CreateDirectory(item.Value);
                            continue;
                        }

                        Directory.CreateDirectory(Path.GetDirectoryName(item.Value)!);
                        item.Key.ExtractToFile(item.Value, overwrite: false);
                    }

                    Log("unpacked " + places.Count.ToString(CultureInfo.InvariantCulture) + " entries");
                }
            }
            catch (Exception ex) when (ex is IOException || ex is InvalidDataException || ex is UnauthorizedAccessException
                                       || ex is NotSupportedException || ex is ArgumentException)
            {
                Log("unpack failed: " + ex.GetType().Name + ": " + ex.Message);
                return Outcome.UnpackFailed;
            }

            if (!ServerRuntime.IsComplete(target))
            {
                Log("unpack failed: the package holds no runtime " + ServerRuntime.Version);
                return Outcome.UnpackFailed;
            }

            return null;
        }

        private Outcome Swap(string unpacked, string retired)
        {
            bool movedAway = false;
            try
            {
                if (Directory.Exists(_root))
                {
                    Directory.Move(_root, retired);
                    movedAway = true;
                    Log("the earlier runtime folder is set aside");
                }

                Directory.Move(unpacked, _root);
                Log("installed");
                return Outcome.Installed;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Log("swap failed: " + ex.GetType().Name + ": " + ex.Message);
                if (movedAway && !Directory.Exists(_root))
                {
                    try
                    {
                        Directory.Move(retired, _root);
                        Log("the earlier runtime folder is back");
                    }
                    catch (Exception back) when (back is IOException || back is UnauthorizedAccessException)
                    {
                        Log("the earlier runtime folder could not be put back: " + back.Message);
                    }
                }

                return Outcome.SwapFailed;
            }
        }

        private void RemoveWorkFolders()
        {
            string parent = Path.GetDirectoryName(_root)!;
            string prefix = WorkFolderPrefix(_root);
            if (!Directory.Exists(parent))
                return;

            foreach (string folder in Directory.GetDirectories(parent, prefix + "*"))
            {
                // The pattern also matches short 8.3 names; the folder's own name must carry the prefix.
                if (!Path.GetFileName(folder).StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                DeleteTree(folder);
                Log("removed " + folder + ", which an earlier run left behind");
            }
        }

        private void DeleteTree(string folder)
        {
            try
            {
                if (Directory.Exists(folder))
                    Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Log("could not delete " + folder + ": " + ex.Message + " - the next run tries again");
            }
        }

        private void Log(string line)
        {
            _log.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture) + " " + line);
        }

        private static Func<Uri, CancellationToken, Stream> HttpDownload(HttpClient http)
        {
            return (uri, cancel) =>
            {
                HttpResponseMessage response = http.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancel).GetAwaiter().GetResult();
                response.EnsureSuccessStatusCode();
                return response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
            };
        }

        /// <summary>Copies at most <paramref name="limit"/> bytes: a package longer than pinned is refused by its length anyway.</summary>
        private static void CopyAtMost(Stream source, Stream target, long limit, CancellationToken cancel)
        {
            byte[] buffer = new byte[81920];
            long total = 0;
            while (total < limit)
            {
                int read = source.ReadAsync(buffer, 0, buffer.Length, cancel).GetAwaiter().GetResult();
                if (read <= 0)
                    break;
                int keep = (int)Math.Min(read, limit - total);
                target.Write(buffer, 0, keep);
                total += keep;
            }
        }

        private static string Sha512Of(string path)
        {
            using (SHA512 sha = SHA512.Create())
            using (FileStream stream = File.OpenRead(path))
            {
                byte[] hash = sha.ComputeHash(stream);
                var text = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                    text.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return text.ToString();
            }
        }
    }
}
