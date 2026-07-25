using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAI.Services
{
    /// <summary>
    /// Read-only enumeration of the installed Outlook signatures for the task pane's
    /// "Select the best signature" action (D38). Deliberately self-contained in the
    /// add-in (NO OutlookAI.Core reference - the Phase-6 no-coupling choice): reads
    /// %APPDATA%\Microsoft\Signatures, groups the rendition files by base name, and
    /// derives a short plain-text excerpt (language/purpose detection hint for the AI)
    /// from the .txt rendition (BOM-aware) or a crude HTML strip of the .htm.
    /// Exception-safe: every public member degrades to empty instead of throwing.
    /// </summary>
    internal static class SignatureStore
    {
        /// <summary>Maximum excerpt length passed to the AI per signature.</summary>
        internal const int ExcerptMaxChars = 160;

        /// <summary>One installed signature (name + AI hint + insertable file).</summary>
        internal sealed class SignatureOption
        {
            public SignatureOption(string name, string excerpt, string filePath)
            {
                Name = name;
                Excerpt = excerpt;
                FilePath = filePath;
            }

            /// <summary>Signature name (file base name, what Outlook's pickers show).</summary>
            public string Name { get; }

            /// <summary>Short plain-text excerpt ("" when unreadable).</summary>
            public string Excerpt { get; }

            /// <summary>Best rendition for Word InsertFile: .htm preferred, then .rtf, then .txt.</summary>
            public string FilePath { get; }
        }

        /// <summary>The Outlook signatures directory (%APPDATA%\Microsoft\Signatures).</summary>
        internal static string SignatureDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Signatures");
            }
        }

        /// <summary>
        /// Enumerates the installed signatures, name-sorted. Empty on a missing
        /// directory or any IO trouble - never throws.
        /// </summary>
        internal static List<SignatureOption> ListSignatures()
        {
            var result = new List<SignatureOption>();
            try
            {
                string root = SignatureDirectory;
                if (!Directory.Exists(root))
                    return result;

                var byName = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                foreach (string file in Directory.GetFiles(root))
                {
                    string extension = Path.GetExtension(file);
                    if (!extension.Equals(".htm", StringComparison.OrdinalIgnoreCase)
                        && !extension.Equals(".html", StringComparison.OrdinalIgnoreCase)
                        && !extension.Equals(".rtf", StringComparison.OrdinalIgnoreCase)
                        && !extension.Equals(".txt", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string name = Path.GetFileNameWithoutExtension(file);
                    if (string.IsNullOrEmpty(name))
                        continue;

                    Dictionary<string, string> files;
                    if (!byName.TryGetValue(name, out files))
                    {
                        files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        byName[name] = files;
                    }

                    files[extension.ToLowerInvariant()] = file;
                }

                foreach (var pair in byName.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
                {
                    string html = pair.Value.ContainsKey(".htm") ? pair.Value[".htm"]
                        : pair.Value.ContainsKey(".html") ? pair.Value[".html"] : null;
                    string rtf = pair.Value.ContainsKey(".rtf") ? pair.Value[".rtf"] : null;
                    string text = pair.Value.ContainsKey(".txt") ? pair.Value[".txt"] : null;
                    string insertable = html ?? rtf ?? text;
                    if (insertable == null)
                        continue;

                    result.Add(new SignatureOption(pair.Key, ReadExcerpt(text, html), insertable));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SignatureStore.ListSignatures: " + ex.Message);
                result.Clear();
            }

            return result;
        }

        /// <summary>True when at least one signature is installed (button enablement).</summary>
        internal static bool AnySignatureInstalled()
        {
            return ListSignatures().Count > 0;
        }

        /// <summary>Finds an option by name (case-insensitive, trimmed); null when absent.</summary>
        internal static SignatureOption FindByName(List<SignatureOption> options, string name)
        {
            if (options == null || string.IsNullOrWhiteSpace(name))
                return null;

            string trimmed = name.Trim().Trim('"', '\'', '.');
            return options.FirstOrDefault(o => string.Equals(o.Name, trimmed, StringComparison.OrdinalIgnoreCase));
        }

        private static string ReadExcerpt(string textPath, string htmlPath)
        {
            try
            {
                string plain = null;
                if (textPath != null && File.Exists(textPath))
                {
                    // Outlook writes .txt renditions as UTF-16 LE with BOM; BOM-aware
                    // detection also copes with UTF-8 files from other writers.
                    using (var reader = new StreamReader(textPath, Encoding.UTF8, true))
                        plain = reader.ReadToEnd();
                }
                else if (htmlPath != null && File.Exists(htmlPath))
                {
                    string html;
                    using (var reader = new StreamReader(htmlPath, Encoding.UTF8, true))
                        html = reader.ReadToEnd();
                    plain = StripHtml(html);
                }

                if (string.IsNullOrWhiteSpace(plain))
                    return "";

                var lines = new List<string>(2);
                foreach (string rawLine in plain.Split('\n'))
                {
                    string line = rawLine.Trim().Trim('\uFEFF', '\uFFFE');
                    if (line.Length > 0)
                    {
                        lines.Add(line);
                        if (lines.Count == 2)
                            break;
                    }
                }

                if (lines.Count == 0)
                    return "";

                string excerpt = string.Join(" / ", lines);
                return excerpt.Length > ExcerptMaxChars ? excerpt.Substring(0, ExcerptMaxChars) : excerpt;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SignatureStore.ReadExcerpt: " + ex.Message);
                return "";
            }
        }

        /// <summary>
        /// Crude tag strip for excerpt purposes only (head/style/script dropped, tags
        /// removed, core entities decoded) - NOT a general HTML-to-text converter.
        /// </summary>
        private static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
                return "";

            string s = Regex.Replace(html, @"<(head|style|script)[^>]*>.*?</\1>", " ",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            s = Regex.Replace(s, @"<(br|/p|/div|/tr)[^>]*>", "\n", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, "<[^>]+>", " ");
            s = s.Replace("&nbsp;", " ").Replace("&amp;", "&").Replace("&lt;", "<")
                .Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&#39;", "'");
            s = Regex.Replace(s, @"&#(\d+);", m =>
            {
                int code;
                return int.TryParse(m.Groups[1].Value, out code) && code > 0 && code <= 0x10FFFF
                    ? char.ConvertFromUtf32(code)
                    : " ";
            });
            return Regex.Replace(s, @"[ \t]+", " ");
        }
    }
}
