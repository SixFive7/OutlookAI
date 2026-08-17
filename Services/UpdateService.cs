using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OutlookAI.Services
{
    internal static class UpdateService
    {
        private const string GitHubOwner = "SixFive7";
        private const string GitHubRepo = "OutlookAI";
        private static readonly TimeSpan PollInterval = TimeSpan.FromMinutes(10);

        /// <summary>
        /// How long the GitHub API call may take. Generous for a JSON GET, and deliberately
        /// unchanged from the ambient <c>HttpClient.Timeout</c> it replaced - the client no
        /// longer carries one, because an ambient timeout said nothing useful about the
        /// installer download below (see <see cref="DownloadTimeout"/>).
        /// </summary>
        private static readonly TimeSpan ApiTimeout = TimeSpan.FromMinutes(5);

        /// <summary>
        /// Largest installer the updater will accept. Mirrored by the release workflow, which
        /// names this constant and fails the release rather than shipping an asset the whole
        /// installed base would silently refuse - see the "Create installer" step in
        /// <c>.github/workflows/release.yml</c>, and the drift check in
        /// <c>.github/scripts/check-pinned-constants.ps1</c> that compares the two.
        /// </summary>
        private const long MaxDownloadBytes = 50L * 1024 * 1024; // 50 MB

        /// <summary>
        /// Slowest transfer the updater will sit through. Only ever used to derive
        /// <see cref="DownloadTimeout"/>; it is not enforced moment to moment.
        /// </summary>
        private const long MinDownloadBytesPerSecond = 64 * 1024; // 64 KB/s

        /// <summary>
        /// How long the installer download may take, DERIVED from the cap it guards rather
        /// than guessed at: at <see cref="MinDownloadBytesPerSecond"/> the largest asset
        /// <see cref="MaxDownloadBytes"/> allows still finishes inside it, so this can no
        /// longer be exceeded by the very operation it is supposed to bound. The old fixed
        /// five minutes demanded ~170 KB/s sustained of a 50 MB download or the update aborted.
        /// </summary>
        private static readonly TimeSpan DownloadTimeout =
            TimeSpan.FromSeconds(MaxDownloadBytes / MinDownloadBytesPerSecond);

        /// <summary>
        /// How long the handed-off script waits after Outlook exits before running the
        /// installer, giving Windows time to release the add-in DLLs. A guess, but a bounded
        /// one: too short and the silent install fails, and the next check retries it.
        /// </summary>
        private const int InstallerGraceSeconds = 2;

        // Named because DescribeState has to tell "nothing to report" apart from every other
        // status, and a second spelling of the same words in a different file is a bug waiting.
        private const string StatusUpToDate = "up to date";
        private const string StatusDeveloperBuild = "developer build";
        private const string StatusChecking = "checking…";

        private static readonly HttpClient _httpClient = CreateHttpClient();

        private static volatile string _etag;
        private static Timer _timer;
        private static Process _updateProcess;
        private static int _checking;

        private static volatile string _lastChecked;
        private static volatile string _lastError;
        private static volatile string _status;

        public static DateTime? LastChecked
        {
            get { var s = _lastChecked; return s == null ? (DateTime?)null : DateTime.ParseExact(s, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind); }
            private set { _lastChecked = value?.ToString("o"); }
        }
        public static string LastError
        {
            get { return _lastError; }
            private set { _lastError = value; }
        }
        public static string Status
        {
            get { return _status; }
            private set { _status = value; }
        }

        /// <summary>
        /// Whether a check is in flight right now - the <see cref="PollInterval"/> poll's, or
        /// one the user asked for. Both version indicators read it, so a check from the settings
        /// dialog shows up in the sidebar too.
        /// </summary>
        public static bool IsChecking
        {
            get { return Volatile.Read(ref _checking) != 0; }
        }

        /// <summary>
        /// How often <see cref="Start"/> polls, spelled out for a human. Every sentence in the
        /// product that says how often OutlookAI looks for an update builds itself from this,
        /// rather than restating the number: the interval used to exist once in code and four
        /// times in English, one of those on screen in a tooltip.
        /// </summary>
        public static string PollIntervalDescription
        {
            get { return DescribeInterval(PollInterval); }
        }

        /// <summary>
        /// How often a version indicator has to re-read <see cref="VersionLine"/>. It lives
        /// here rather than in either piece of UI for the same reason the wording does: the
        /// sidebar and the settings dialog were ticking at the same rate by coincidence, in two
        /// files, and one of them could have been changed without the other.
        ///
        /// One second, and that is deliberately faster than the once-a-minute rollover it
        /// serves: the tick is also how a check started in the OTHER indicator becomes visible
        /// here, so slowing it slows down a status line the user is watching. Both readers make
        /// an unchanged tick free (neither repaints, and the dialog does not re-lay out), which
        /// is what makes the rate affordable. <c>TODO.md</c> tracks replacing the polling half
        /// of this with a StateChanged event, after which only the rollover needs a timer at
        /// all and it can drop to 30-60s.
        /// </summary>
        public const int VersionLineTickMs = 1000;

        private static string DescribeInterval(TimeSpan span)
        {
            int minutes = (int)Math.Round(span.TotalMinutes);
            if (minutes >= 1)
                return minutes == 1 ? "1 minute" : minutes + " minutes";
            int seconds = (int)Math.Round(span.TotalSeconds);
            return seconds == 1 ? "1 second" : seconds + " seconds";
        }

        /// <summary>
        /// The single line both version indicators show: "v3.1.0.325 - checked 4m ago", or what
        /// the updater is busy with. It lives here rather than in either piece of UI so the
        /// sidebar and the settings dialog cannot word the same state two different ways.
        /// </summary>
        public static string VersionLine()
        {
            var version = "v" + Assembly.GetExecutingAssembly().GetName().Version;
            var state = DescribeState();
            return state != null ? version + " - " + state : version;
        }

        /// <summary>
        /// The part after the version: what the updater is doing, or how long ago it last
        /// managed to ask. Null when there is nothing worth saying — a first check that failed,
        /// where the error line carries it instead.
        /// </summary>
        private static string DescribeState()
        {
            var lastChecked = LastChecked;
            var error = LastError;
            var status = Status;

            // An update being downloaded or waiting to install outranks everything below.
            if (status != null && status != StatusUpToDate)
                return status;
            if (IsChecking)
                return StatusChecking;
            if (lastChecked == null)
                return error != null ? null : StatusChecking;

            var ago = DateTime.Now - lastChecked.Value;
            if (ago < TimeSpan.Zero) ago = TimeSpan.Zero;
            if (ago.TotalSeconds < 60)
                return "checked just now";
            if (ago.TotalMinutes < 60)
                return $"checked {(int)ago.TotalMinutes}m ago";
            if (ago.TotalHours < 24)
                return $"checked {(int)ago.TotalHours}h ago";
            return $"checked {(int)ago.TotalDays}d ago";
        }

        private static HttpClient CreateHttpClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "OutlookAI-Updater");
            // No ambient timeout: every request below carries its own deadline as a
            // CancellationToken instead. HttpClient.Timeout is the wrong tool for the installer
            // download - that one is fetched with HttpCompletionOption.ResponseHeadersRead and
            // the body is copied by hand, which is outside whatever the ambient timeout covers.
            // Rather than depend on exactly where net48 draws that line, nothing here relies on
            // it at all.
            client.Timeout = Timeout.InfiniteTimeSpan;
            return client;
        }

        public static void Start()
        {
            // Fire immediately, then once per PollInterval.
            _timer = new Timer(_ => _ = RunCheckAsync(), null, TimeSpan.Zero, PollInterval);
        }

        public static void Stop()
        {
            _timer?.Dispose();
            _timer = null;
        }

        /// <summary>
        /// Runs the check now instead of waiting out the rest of the poll interval, and
        /// completes when that check is done. Safe to call from a UI click handler, and a
        /// no-op while a check is already running — so an impatient second click costs nothing.
        /// </summary>
        public static Task CheckNowAsync()
        {
            return RunCheckAsync();
        }

        /// <summary>
        /// Claims the one-at-a-time guard, then runs the check on a worker thread.
        ///
        /// The guard is claimed on the CALLING thread deliberately, which is why it moved out
        /// of <see cref="CheckForUpdateAsync"/>: "Check for updates" repaints the instant this
        /// returns, and a guard claimed only after the first await would leave that repaint
        /// saying "checked 4m ago" about a check the user had just started.
        ///
        /// The worker thread matters too. Nothing below uses ConfigureAwait(false), so a check
        /// started inline from a click handler would capture Outlook's message pump and resume
        /// every continuation of a 50 MB download on the UI thread.
        /// </summary>
        private static Task RunCheckAsync()
        {
            // Developer/from-source builds carry the 99.99.99.0 placeholder (CI stamps the real
            // version at release). Never auto-update such a build, and skip the network check.
            // Answered here, ahead of the guard: it is a return that never reaches the finally
            // below, so inside the guarded region it would leak the guard and wedge the updater.
            if (Assembly.GetExecutingAssembly().GetName().Version.Major == 99)
            {
                Status = StatusDeveloperBuild;
                return Task.CompletedTask;
            }

            if (Interlocked.CompareExchange(ref _checking, 1, 0) != 0)
                return Task.CompletedTask;

            try
            {
                return Task.Run((Func<Task>)CheckForUpdateAsync);
            }
            catch (Exception ex)
            {
                // Only reachable if the work could not even be queued, but releasing the guard
                // here is what stops that being permanent: claimed and never released, it wedges
                // the updater for the session and leaves both indicators stuck on "checking…".
                Interlocked.Exchange(ref _checking, 0);
                LastError = ex.Message;
                return Task.CompletedTask;
            }
        }

        // Entered with _checking already claimed by RunCheckAsync, and releases it in its
        // finally. Never call it directly.
        private static async Task CheckForUpdateAsync()
        {
            string tempPath = null;
            bool installerHandedOff = false;
            try
            {
                var apiUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
                Dictionary<string, object> release;
                string newEtag = null;
                using (var request = new HttpRequestMessage(HttpMethod.Get, apiUrl))
                using (var apiCts = new CancellationTokenSource(ApiTimeout))
                {
                    if (_etag != null)
                        request.Headers.IfNoneMatch.Add(new EntityTagHeaderValue(_etag));

                    // Default completion option, so the whole JSON body is buffered inside this
                    // call and the ReadAsStringAsync below is memory-only - which is what keeps
                    // the API half fully inside ApiTimeout.
                    using (var response = await _httpClient.SendAsync(request, apiCts.Token))
                    {
                        if (response.StatusCode == HttpStatusCode.NotModified)
                        {
                            LastChecked = DateTime.Now;
                            LastError = null;
                            return;
                        }

                        if (!response.IsSuccessStatusCode)
                        {
                            LastError = $"GitHub API returned {(int)response.StatusCode} {response.ReasonPhrase}";
                            return;
                        }

                        LastChecked = DateTime.Now;
                        LastError = null;

                        newEtag = response.Headers.ETag?.Tag;

                        var json = await response.Content.ReadAsStringAsync();
                        var serializer = new JavaScriptSerializer();
                        release = serializer.Deserialize<Dictionary<string, object>>(json);
                    }
                }

                if (release == null || !release.TryGetValue("tag_name", out var tagObj) || !(tagObj is string tagName))
                {
                    LastError = "GitHub release response was missing a tag_name.";
                    return;
                }
                if (!Version.TryParse(tagName.TrimStart('v'), out var rv))
                {
                    LastError = $"Could not parse release version from tag '{tagName}'.";
                    return;
                }

                // Only cache the ETag once the body parsed into a well-formed release.
                // Caching it earlier means a garbled or non-release response would be
                // remembered and every later poll would get 304 Not Modified, wedging
                // the updater until the upstream release changes.
                _etag = newEtag;

                var remoteVersion = new Version(rv.Major, rv.Minor, rv.Build < 0 ? 0 : rv.Build, 0);
                var lv = Assembly.GetExecutingAssembly().GetName().Version;
                var localVersion = new Version(lv.Major, lv.Minor, lv.Build, 0);

                if (remoteVersion <= localVersion)
                {
                    Status = StatusUpToDate;
                    return;
                }

                // Skip while an update process is still waiting to install (Outlook open).
                // If a previous one already exited, dispose it and allow a fresh attempt —
                // we retry on every check/restart with no cross-restart state.
                if (_updateProcess != null)
                {
                    if (!_updateProcess.HasExited)
                        return;
                    _updateProcess.Dispose();
                    _updateProcess = null;
                }

                Status = $"downloading v{remoteVersion}\u2026";

                // Find the installer asset, pinned to the expected name (OutlookAI-<tag>.exe);
                // fall back to any OutlookAI*.exe. The signature/thumbprint check is the real
                // anchor, but pinning the name avoids grabbing an unrelated .exe asset.
                string downloadUrl = null;
                string fallbackUrl = null;
                var expectedName = "OutlookAI-" + tagName + ".exe";
                var assets = (ArrayList)release["assets"];
                foreach (Dictionary<string, object> asset in assets)
                {
                    var name = asset["name"] as string;
                    if (name == null || !name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (string.Equals(name, expectedName, StringComparison.OrdinalIgnoreCase))
                    {
                        downloadUrl = (string)asset["browser_download_url"];
                        break;
                    }
                    if (fallbackUrl == null && name.StartsWith("OutlookAI", StringComparison.OrdinalIgnoreCase))
                        fallbackUrl = (string)asset["browser_download_url"];
                }
                if (downloadUrl == null)
                    downloadUrl = fallbackUrl;

                if (downloadUrl == null)
                {
                    LastError = "No installer asset found in latest release";
                    return;
                }

                tempPath = Path.Combine(Path.GetTempPath(), "OutlookAI-" + Path.GetRandomFileName() + ".exe");
                using (var downloadCts = new CancellationTokenSource(DownloadTimeout))
                using (var response2 = await _httpClient.GetAsync(
                    downloadUrl, HttpCompletionOption.ResponseHeadersRead, downloadCts.Token))
                {
                    response2.EnsureSuccessStatusCode();
                    var contentLength = response2.Content.Headers.ContentLength;
                    if (contentLength.HasValue && contentLength.Value > MaxDownloadBytes)
                    {
                        LastError = "Installer download too large";
                        return;
                    }
                    // Disposing the response tears down its content stream, which is what makes
                    // a read that has stopped delivering bytes actually fail instead of waiting
                    // forever. Without it a stalled socket hangs this copy loop, and because the
                    // one-at-a-time guard is only released in the finally below, both version
                    // indicators would sit on "checking..." for the rest of the session.
                    using (downloadCts.Token.Register(() => { try { response2.Dispose(); } catch { } }))
                    {
                        try
                        {
                            using (var src = await response2.Content.ReadAsStreamAsync())
                            using (var dst = File.Create(tempPath))
                            {
                                // The .NET default CopyTo buffer size.
                                var buf = new byte[81920];
                                long total = 0;
                                int read;
                                while ((read = await src.ReadAsync(buf, 0, buf.Length, downloadCts.Token)) > 0)
                                {
                                    total += read;
                                    if (total > MaxDownloadBytes)
                                    {
                                        dst.Close();
                                        File.Delete(tempPath);
                                        LastError = "Installer download too large";
                                        return;
                                    }
                                    await dst.WriteAsync(buf, 0, read, downloadCts.Token);
                                }
                            }
                        }
                        catch (Exception) when (downloadCts.IsCancellationRequested)
                        {
                            // The deadline fired and pulled the response out from under the
                            // read. Whatever the stream threw on the way out names the
                            // mechanism, not the problem - so report the problem.
                            throw new OperationCanceledException(downloadCts.Token);
                        }
                    }
                }

                if (!VerifySignature(tempPath, out var sigError))
                {
                    File.Delete(tempPath);
                    LastError = sigError;
                    return;
                }

                // Spawn a hidden process that waits for Outlook to exit,
                // then runs the installer. -Wait keeps this launcher alive until
                // the installer finishes, so the guard above won't spawn another
                // while one is still pending.
                var installerArgs = "/SILENT /SP- /NOCANCEL /NORESTART /NORESTARTAPPLICATIONS";
                var safePath = tempPath.Replace("'", "''");
                var script = $"Get-Process outlook -ErrorAction SilentlyContinue | Wait-Process; Start-Sleep -Seconds {InstallerGraceSeconds}; if (-not (Test-Path 'HKCU:\\Software\\Microsoft\\Office\\Outlook\\Addins\\OutlookAI')) {{ exit }}; Start-Process '{safePath}' -ArgumentList '{installerArgs}' -Wait";
                var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
                _updateProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-WindowStyle Hidden -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });

                installerHandedOff = true;
                Status = $"v{remoteVersion} ready - installs on close";
            }
            catch (OperationCanceledException)
            {
                // Only one thing cancels anything here: a deadline of ours expiring. Said
                // plainly, because "The operation was canceled." on the update line tells a
                // user nothing about which operation or why.
                LastError = "Update check timed out.";
                Status = null;
            }
            catch (Exception ex)
            {
                LastError = ex.InnerException?.Message ?? ex.Message;
                Status = null;
            }
            finally
            {
                if (!installerHandedOff && tempPath != null)
                {
                    try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                }
                Interlocked.Exchange(ref _checking, 0);
            }
        }

        // === Installer signature verification ===
        // The installer is Authenticode-signed with a SELF-SIGNED certificate (CN=OutlookAI).
        // We require BOTH: (1) the signature is cryptographically valid (file not tampered) and
        // (2) the signer is exactly our certificate (thumbprint pin). Because the cert is
        // self-signed, WinVerifyTrust reports CERT_E_UNTRUSTEDROOT even for a perfectly valid
        // file, so we accept ONLY that specific "valid hash / untrusted root" result and then
        // pin the thumbprint. Every other outcome is rejected. Fail-closed: any error treats
        // the installer as unverified (auto-update stops; manual update still works).
        //
        // THE SAME 40 HEX CHARACTERS ARE IN OutlookAI.csproj AS ManifestCertificateThumbprint.
        // Rotating the signing certificate has to change BOTH. The csproj half fails loudly at
        // build time; this half fails closed and silently - every future installer is rejected
        // as "not signed by the expected OutlookAI certificate" and auto-update stops across the
        // entire installed base, with nothing to show for it but a line in the update-error
        // label. So the two are compared mechanically rather than by memory:
        // .github/scripts/check-pinned-constants.ps1 fails the build (and the release) when they
        // drift, and the release workflow additionally checks both against the certificate it
        // has just imported.
        private const string ExpectedCertThumbprint = "2578F7B869383572E751DD6B61B5374C55C6E995";

        private static bool VerifySignature(string path, out string error)
        {
            error = null;

            int trust = WinVerifyTrustFile(path);
            const int CERT_E_UNTRUSTEDROOT = unchecked((int)0x800B0109);
            if (trust != 0 && trust != CERT_E_UNTRUSTEDROOT)
            {
                error = "Installer signature is invalid or missing (0x" + trust.ToString("X8") + ").";
                return false;
            }

            try
            {
                using (var cert = new X509Certificate2(X509Certificate.CreateFromSignedFile(path)))
                {
                    if (!string.Equals(cert.Thumbprint, ExpectedCertThumbprint, StringComparison.OrdinalIgnoreCase))
                    {
                        error = "Installer was not signed by the expected OutlookAI certificate.";
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "Could not read installer signature: " + ex.Message;
                return false;
            }

            return true;
        }

        private static readonly Guid WINTRUST_ACTION_GENERIC_VERIFY_V2 =
            new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");

        private static int WinVerifyTrustFile(string path)
        {
            var fileInfo = new WINTRUST_FILE_INFO
            {
                cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)),
                pcwszFilePath = path,
                hFile = IntPtr.Zero,
                pgKnownSubject = IntPtr.Zero
            };

            IntPtr pFile = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
            IntPtr pData = IntPtr.Zero;
            try
            {
                Marshal.StructureToPtr(fileInfo, pFile, false);

                var data = new WINTRUST_DATA
                {
                    cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_DATA)),
                    pPolicyCallbackData = IntPtr.Zero,
                    pSIPClientData = IntPtr.Zero,
                    dwUIChoice = 2,            // WTD_UI_NONE
                    fdwRevocationChecks = 0,   // WTD_REVOKE_NONE
                    dwUnionChoice = 1,         // WTD_CHOICE_FILE
                    pFile = pFile,
                    dwStateAction = 0,         // WTD_STATEACTION_IGNORE
                    hWVTStateData = IntPtr.Zero,
                    pwszURLReference = IntPtr.Zero,
                    dwProvFlags = 0x00000100,  // WTD_SAFER_FLAG
                    dwUIContext = 0,
                    pSignatureSettings = IntPtr.Zero
                };

                pData = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_DATA)));
                Marshal.StructureToPtr(data, pData, false);
                return WinVerifyTrust(IntPtr.Zero, WINTRUST_ACTION_GENERIC_VERIFY_V2, pData);
            }
            catch
            {
                // Fail closed: any marshaling/P-Invoke problem means "not verified".
                return unchecked((int)0x80004005); // E_FAIL
            }
            finally
            {
                if (pData != IntPtr.Zero) Marshal.FreeHGlobal(pData);
                Marshal.DestroyStructure(pFile, typeof(WINTRUST_FILE_INFO));
                Marshal.FreeHGlobal(pFile);
            }
        }

        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, IntPtr pWVTData);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;
            [MarshalAs(UnmanagedType.LPWStr)] public string pcwszFilePath;
            public IntPtr hFile;
            public IntPtr pgKnownSubject;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINTRUST_DATA
        {
            public uint cbStruct;
            public IntPtr pPolicyCallbackData;
            public IntPtr pSIPClientData;
            public uint dwUIChoice;
            public uint fdwRevocationChecks;
            public uint dwUnionChoice;
            public IntPtr pFile;
            public uint dwStateAction;
            public IntPtr hWVTStateData;
            public IntPtr pwszURLReference;
            public uint dwProvFlags;
            public uint dwUIContext;
            public IntPtr pSignatureSettings;
        }
    }
}
