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
        /// Whether a check is in flight right now — the ten-minute poll's, or one the user
        /// asked for. Both version indicators read it, so a check started from the settings
        /// dialog shows up in the sidebar too.
        /// </summary>
        public static bool IsChecking
        {
            get { return Volatile.Read(ref _checking) != 0; }
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
            client.Timeout = TimeSpan.FromMinutes(5);
            return client;
        }

        public static void Start()
        {
            // Fire immediately, then every 10 minutes
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
                {
                    if (_etag != null)
                        request.Headers.IfNoneMatch.Add(new EntityTagHeaderValue(_etag));

                    using (var response = await _httpClient.SendAsync(request))
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
                const long maxDownloadBytes = 50 * 1024 * 1024; // 50 MB
                using (var response2 = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
                {
                    response2.EnsureSuccessStatusCode();
                    var contentLength = response2.Content.Headers.ContentLength;
                    if (contentLength.HasValue && contentLength.Value > maxDownloadBytes)
                    {
                        LastError = "Installer download too large";
                        return;
                    }
                    using (var src = await response2.Content.ReadAsStreamAsync())
                    using (var dst = File.Create(tempPath))
                    {
                        var buf = new byte[81920];
                        long total = 0;
                        int read;
                        while ((read = await src.ReadAsync(buf, 0, buf.Length)) > 0)
                        {
                            total += read;
                            if (total > maxDownloadBytes)
                            {
                                dst.Close();
                                File.Delete(tempPath);
                                LastError = "Installer download too large";
                                return;
                            }
                            await dst.WriteAsync(buf, 0, read);
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
                var script = $"Get-Process outlook -ErrorAction SilentlyContinue | Wait-Process; Start-Sleep -Seconds 2; if (-not (Test-Path 'HKCU:\\Software\\Microsoft\\Office\\Outlook\\Addins\\OutlookAI')) {{ exit }}; Start-Process '{safePath}' -ArgumentList '{installerArgs}' -Wait";
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
