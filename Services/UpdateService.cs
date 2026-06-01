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
        private static readonly HttpClient _httpClient = CreateHttpClient();

        private const int MaxUpdateFailures = 3;

        private static volatile string _etag;
        private static Timer _timer;
        private static Process _updateProcess;
        private static int _updateFailures;
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
            _timer = new Timer(_ => _ = CheckForUpdateAsync(), null, TimeSpan.Zero, PollInterval);
        }

        public static void Stop()
        {
            _timer?.Dispose();
            _timer = null;
        }

        private static async Task CheckForUpdateAsync()
        {
            if (Interlocked.CompareExchange(ref _checking, 1, 0) != 0)
                return;
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
                    Status = "up to date";
                    return;
                }

                // Skip if an update process is still running
                if (_updateProcess != null && !_updateProcess.HasExited)
                    return;

                // If a previous update process exited without the version changing,
                // it failed. Count failures and stop retrying after the limit.
                if (_updateProcess != null && _updateProcess.HasExited)
                {
                    _updateFailures++;
                    _updateProcess = null;
                    if (_updateFailures >= MaxUpdateFailures)
                    {
                        LastError = $"Update failed {MaxUpdateFailures} times, not retrying";
                        return;
                    }
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
                // then runs the installer. -Wait keeps the process alive until
                // the installer finishes, so we can detect completion/failure.
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
