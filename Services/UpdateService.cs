using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
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

        private static string _etag;
        private static Timer _timer;
        private static Process _updateProcess;
        private static int _updateFailures;
        private static int _checking;

        public static DateTime? LastChecked { get; private set; }
        public static string LastError { get; private set; }
        public static string Status { get; private set; }

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
            try
            {
                var apiUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
                Dictionary<string, object> release;
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

                        if (response.Headers.ETag != null)
                            _etag = response.Headers.ETag.Tag;

                        var json = await response.Content.ReadAsStringAsync();
                        var serializer = new JavaScriptSerializer();
                        release = serializer.Deserialize<Dictionary<string, object>>(json);
                    }
                }

                var tagName = (string)release["tag_name"];
                var remoteVersion = Version.Parse(tagName.TrimStart('v'));
                var localVersion = Assembly.GetExecutingAssembly().GetName().Version;

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

                // Find the .exe installer asset
                string downloadUrl = null;
                var assets = (ArrayList)release["assets"];
                foreach (Dictionary<string, object> asset in assets)
                {
                    var name = (string)asset["name"];
                    if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        downloadUrl = (string)asset["browser_download_url"];
                        break;
                    }
                }

                if (downloadUrl == null)
                {
                    LastError = "No installer asset found in latest release";
                    return;
                }

                var tempPath = Path.Combine(Path.GetTempPath(), "OutlookAI-Update.exe");
                var bytes = await _httpClient.GetByteArrayAsync(downloadUrl);
                File.WriteAllBytes(tempPath, bytes);

                // Spawn a hidden process that waits for Outlook to exit,
                // then runs the installer. -Wait keeps the process alive until
                // the installer finishes, so we can detect completion/failure.
                var installerArgs = "/SILENT /SP- /NOCANCEL /NORESTART /NORESTARTAPPLICATIONS";
                var script = $"Get-Process outlook -ErrorAction SilentlyContinue | Wait-Process; Start-Sleep -Seconds 2; if (-not (Test-Path 'HKCU:\\Software\\Microsoft\\Office\\Outlook\\Addins\\OutlookAI')) {{ exit }}; Start-Process '{tempPath}' -ArgumentList '{installerArgs}' -Wait";
                _updateProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-WindowStyle Hidden -ExecutionPolicy Bypass -Command \"{script}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });

                Status = $"v{remoteVersion} ready - installs on close";
            }
            catch (Exception ex)
            {
                LastError = ex.InnerException?.Message ?? ex.Message;
            }
            finally
            {
                Interlocked.Exchange(ref _checking, 0);
            }
        }
    }
}
