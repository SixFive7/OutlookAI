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

        private static string _etag;
        private static Timer _timer;
        private static bool _updateLaunched;

        public static DateTime? LastChecked { get; private set; }
        public static string LastError { get; private set; }
        public static string Status { get; private set; }

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
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "OutlookAI-Updater");
                    client.Timeout = TimeSpan.FromSeconds(15);

                    var apiUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
                    var request = new HttpRequestMessage(HttpMethod.Get, apiUrl);
                    if (_etag != null)
                        request.Headers.IfNoneMatch.Add(new EntityTagHeaderValue(_etag));

                    var response = await client.SendAsync(request);

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
                    var release = serializer.Deserialize<Dictionary<string, object>>(json);

                    var tagName = (string)release["tag_name"];
                    var remoteVersion = Version.Parse(tagName.TrimStart('v'));
                    var localVersion = Assembly.GetExecutingAssembly().GetName().Version;

                    if (remoteVersion <= localVersion)
                    {
                        Status = "up to date";
                        return;
                    }

                    if (_updateLaunched)
                        return;

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
                    var bytes = await client.GetByteArrayAsync(downloadUrl);
                    File.WriteAllBytes(tempPath, bytes);

                    // Spawn a hidden process that waits for Outlook to exit,
                    // then runs the installer. This avoids depending on the
                    // unreliable VSTO Shutdown event and never force-closes Outlook.
                    var installerArgs = "/SILENT /SP- /NOCANCEL /NORESTART /NORESTARTAPPLICATIONS";
                    var script = $"Get-Process outlook -ErrorAction SilentlyContinue | Wait-Process; Start-Process '{tempPath}' -ArgumentList '{installerArgs}'";
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-WindowStyle Hidden -ExecutionPolicy Bypass -Command \"{script}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });

                    _updateLaunched = true;
                    Status = $"v{remoteVersion} ready - installs on close";
                }
            }
            catch (Exception ex)
            {
                LastError = ex.InnerException?.Message ?? ex.Message;
            }
        }
    }
}
