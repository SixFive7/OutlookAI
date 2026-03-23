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

        private static string _stagedInstallerPath;
        private static string _etag;
        private static Timer _timer;

        public static DateTime? LastChecked { get; private set; }

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
                        return;
                    }

                    if (!response.IsSuccessStatusCode)
                        return;

                    LastChecked = DateTime.Now;

                    if (response.Headers.ETag != null)
                        _etag = response.Headers.ETag.Tag;

                    var json = await response.Content.ReadAsStringAsync();
                    var serializer = new JavaScriptSerializer();
                    var release = serializer.Deserialize<Dictionary<string, object>>(json);

                    var tagName = (string)release["tag_name"];
                    var remoteVersion = Version.Parse(tagName.TrimStart('v'));
                    var localVersion = Assembly.GetExecutingAssembly().GetName().Version;

                    if (remoteVersion <= localVersion)
                        return;

                    // Already have this update staged
                    if (_stagedInstallerPath != null && File.Exists(_stagedInstallerPath))
                        return;

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
                        return;

                    var tempPath = Path.Combine(Path.GetTempPath(), "OutlookAI-Update.exe");
                    var bytes = await client.GetByteArrayAsync(downloadUrl);
                    File.WriteAllBytes(tempPath, bytes);
                    _stagedInstallerPath = tempPath;
                }
            }
            catch
            {
                // Silent failure — try again next poll
            }
        }

        public static void ApplyIfReady()
        {
            try
            {
                var path = _stagedInstallerPath;
                if (path == null || !File.Exists(path))
                    return;

                Process.Start(new ProcessStartInfo
                {
                    FileName = path,
                    Arguments = "/VERYSILENT",
                    UseShellExecute = true
                });
            }
            catch
            {
                // Silent failure
            }
        }
    }
}
