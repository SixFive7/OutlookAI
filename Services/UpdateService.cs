using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OutlookAI.Services
{
    internal static class UpdateService
    {
        private const string GitHubOwner = "SixFive7";
        private const string GitHubRepo = "OutlookAI";

        private static string _stagedInstallerPath;

        public static async Task CheckForUpdateAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "OutlookAI-Updater");
                    client.Timeout = TimeSpan.FromSeconds(15);

                    var apiUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
                    var json = await client.GetStringAsync(apiUrl);
                    var serializer = new JavaScriptSerializer();
                    var release = serializer.Deserialize<Dictionary<string, object>>(json);

                    var tagName = (string)release["tag_name"];
                    var remoteVersion = Version.Parse(tagName.TrimStart('v'));
                    var localVersion = Assembly.GetExecutingAssembly().GetName().Version;

                    if (remoteVersion <= localVersion)
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
                // Silent failure — try again next session
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
