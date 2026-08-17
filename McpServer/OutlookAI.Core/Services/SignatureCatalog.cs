using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// One signature as installed in the Outlook Signatures folder (data only).
    /// </summary>
    public sealed class SignatureInfo
    {
        /// <summary>Creates a signature descriptor.</summary>
        public SignatureInfo(string name, string? excerpt, string? htmlPath, string? rtfPath, string? textPath)
        {
            Name = name;
            Excerpt = excerpt;
            HtmlPath = htmlPath;
            RtfPath = rtfPath;
            TextPath = textPath;
        }

        /// <summary>Signature name (file base name - what Outlook shows in its pickers).</summary>
        public string Name { get; }

        /// <summary>Short plain-text excerpt (first lines) for language/purpose detection.</summary>
        public string? Excerpt { get; }

        /// <summary>Full path of the .htm rendition (preferred for insertion), when present.</summary>
        public string? HtmlPath { get; }

        /// <summary>Full path of the .rtf rendition, when present.</summary>
        public string? RtfPath { get; }

        /// <summary>Full path of the .txt rendition, when present.</summary>
        public string? TextPath { get; }

        /// <summary>The best rendition for Word InsertFile: .htm, then .rtf, then .txt.</summary>
        public string? PreferredFilePath => HtmlPath ?? RtfPath ?? TextPath;
    }

    /// <summary>Registry-determined default-signature assignment of one mail account.</summary>
    public sealed class SignatureAssignment
    {
        /// <summary>Creates an assignment row.</summary>
        public SignatureAssignment(string account, string? newMessageSignature, string? replyForwardSignature)
        {
            Account = account;
            NewMessageSignature = newMessageSignature;
            ReplyForwardSignature = replyForwardSignature;
        }

        /// <summary>Account name as the profile registry records it (the SMTP address).</summary>
        public string Account { get; }

        /// <summary>Signature name assigned for new messages (null = not recorded in the registry - unknown, never guessed).</summary>
        public string? NewMessageSignature { get; }

        /// <summary>Signature name assigned for replies/forwards (null = not recorded - unknown).</summary>
        public string? ReplyForwardSignature { get; }
    }

    /// <summary>
    /// Reads the Outlook signature landscape (soak fix D37, R5 signature steering):
    /// the signature files under %APPDATA%\Microsoft\Signatures (names + short
    /// plain-text excerpts, content never logged) and the per-account default
    /// assignments from the profile registry ("New Signature"/"Reply-Forward
    /// Signature" under the profile's 9375CFF0... key). Assignment reading DEGRADES
    /// GRACEFULLY: values can be absent (an account without a signature, or
    /// Office-roaming-managed state - Phase 4 found them unreadable on this machine at
    /// the time) - absent means UNKNOWN, never a guess. Pure filesystem + registry:
    /// no COM, safe on any machine, host-neutral (v3.MD section 0.5.2).
    /// </summary>
    public static class SignatureCatalog
    {
        /// <summary>Maximum excerpt length (payload discipline).</summary>
        public const int ExcerptMaxChars = 160;

        /// <summary>Prefix of the throwaway test signatures live tests may create (S3 cleanup set).</summary>
        public const string TestSignaturePrefix = "OutlookAI-McpTest-";

        private static readonly string[] SignatureExtensions = { ".htm", ".html", ".rtf", ".txt" };

        /// <summary>The Outlook signatures directory (%APPDATA%\Microsoft\Signatures).</summary>
        public static string DefaultSignatureDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Signatures");

        /// <summary>
        /// Enumerates the installed signatures (name-sorted). Missing directory =>
        /// empty list. <paramref name="directory"/> overrides the default for tests.
        /// </summary>
        public static IReadOnlyList<SignatureInfo> ListSignatures(string? directory = null)
        {
            string root = directory ?? DefaultSignatureDirectory;
            if (!Directory.Exists(root))
            {
                return Array.Empty<SignatureInfo>();
            }

            Dictionary<string, Dictionary<string, string>> byName =
                new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            foreach (string file in Directory.GetFiles(root))
            {
                string extension = Path.GetExtension(file);
                if (!SignatureExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                string name = Path.GetFileNameWithoutExtension(file);
                if (name.Length == 0)
                {
                    continue;
                }

                if (!byName.TryGetValue(name, out Dictionary<string, string>? files))
                {
                    files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    byName[name] = files;
                }

                files[extension.ToLowerInvariant()] = file;
            }

            List<SignatureInfo> result = new List<SignatureInfo>(byName.Count);
            foreach (KeyValuePair<string, Dictionary<string, string>> pair in byName.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                string? html = pair.Value.TryGetValue(".htm", out string? htm) ? htm
                    : pair.Value.TryGetValue(".html", out string? htmlAlt) ? htmlAlt : null;
                string? rtf = pair.Value.TryGetValue(".rtf", out string? rtfPath) ? rtfPath : null;
                string? text = pair.Value.TryGetValue(".txt", out string? txtPath) ? txtPath : null;
                result.Add(new SignatureInfo(pair.Key, TryReadExcerpt(text, html), html, rtf, text));
            }

            return result;
        }

        /// <summary>
        /// Resolves a signature by name (case-insensitive). Returns null when no
        /// signature of that name exists.
        /// </summary>
        public static SignatureInfo? TryResolve(string name, string? directory = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            return ListSignatures(directory)
                .FirstOrDefault(s => string.Equals(s.Name, name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Reads per-account default-signature assignments from the profile registry.
        /// Never throws: unreadable/absent state yields an empty list (unknown - the
        /// Phase-4 degradation contract). <paramref name="valuesReader"/> is the T1
        /// injection seam - production passes null for the live registry.
        /// </summary>
        public static IReadOnlyList<SignatureAssignment> ReadAccountAssignments(
            Func<IReadOnlyList<IReadOnlyDictionary<string, object?>>>? valuesReader = null)
        {
            try
            {
                IReadOnlyList<IReadOnlyDictionary<string, object?>> accountValueSets =
                    valuesReader != null ? valuesReader() : ReadProfileAccountValueSets();

                List<SignatureAssignment> result = new List<SignatureAssignment>();
                foreach (IReadOnlyDictionary<string, object?> values in accountValueSets)
                {
                    string? account = DecodeRegistryString(values.TryGetValue("Account Name", out object? a) ? a : null);

                    // Only mail accounts: profile subkeys also cover address books and
                    // other providers; the SMTP-shaped Account Name is the mail marker.
                    if (account == null || account.IndexOf('@') < 0)
                    {
                        continue;
                    }

                    string? newSignature = DecodeRegistryString(values.TryGetValue("New Signature", out object? n) ? n : null);
                    string? replyForward = DecodeRegistryString(values.TryGetValue("Reply-Forward Signature", out object? r) ? r : null);
                    result.Add(new SignatureAssignment(account, EmptyToNull(newSignature), EmptyToNull(replyForward)));
                }

                return result;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // Unknown, never guessed: callers report no assignment data.
                return Array.Empty<SignatureAssignment>();
            }
        }

        /// <summary>
        /// Reads the raw per-account value sets from
        /// HKCU\...\Outlook\Profiles\&lt;default profile&gt;\9375CFF0413111d3B88A00104B2A6676\*.
        /// </summary>
        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> ReadProfileAccountValueSets()
        {
            // Not a const any more: the Office major in the Outlook root is detected at runtime.
            string outlookRoot = OutlookProfileRegistry.OutlookRootKeyPath;
            const string accountsSubKey = OutlookProfileRegistry.AccountsSubKeyName;

            List<IReadOnlyDictionary<string, object?>> sets = new List<IReadOnlyDictionary<string, object?>>();
            using RegistryKey? outlook = Registry.CurrentUser.OpenSubKey(outlookRoot);
            if (outlook == null)
            {
                return sets;
            }

            string? defaultProfile = outlook.GetValue("DefaultProfile") as string;
            using RegistryKey? profiles = outlook.OpenSubKey("Profiles");
            if (profiles == null)
            {
                return sets;
            }

            IEnumerable<string> profileNames = defaultProfile != null
                ? new[] { defaultProfile }
                : profiles.GetSubKeyNames();
            foreach (string profileName in profileNames)
            {
                using RegistryKey? accounts = profiles.OpenSubKey(profileName + "\\" + accountsSubKey);
                if (accounts == null)
                {
                    continue;
                }

                foreach (string subKeyName in accounts.GetSubKeyNames())
                {
                    using RegistryKey? account = accounts.OpenSubKey(subKeyName);
                    if (account == null)
                    {
                        continue;
                    }

                    Dictionary<string, object?> values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    foreach (string valueName in account.GetValueNames())
                    {
                        values[valueName] = account.GetValue(valueName);
                    }

                    sets.Add(values);
                }
            }

            return sets;
        }

        /// <summary>
        /// Registry values of the profile store come as REG_SZ or REG_BINARY
        /// (UTF-16LE, NUL-terminated) depending on the writer - decode both.
        /// </summary>
        public static string? DecodeRegistryString(object? value)
        {
            if (value is string s)
            {
                return s.TrimEnd('\0');
            }

            if (value is byte[] bytes && bytes.Length >= 2)
            {
                try
                {
                    return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
                }
                catch (ArgumentException)
                {
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Short plain-text excerpt of a signature: the first two non-empty lines of
        /// the .txt rendition (BOM-aware decode), else of the HTML converted to text.
        /// Capped at <see cref="ExcerptMaxChars"/>. Null when nothing is readable.
        /// </summary>
        public static string? TryReadExcerpt(string? textPath, string? htmlPath)
        {
            try
            {
                string? plain = null;
                if (textPath != null && File.Exists(textPath))
                {
                    using StreamReader reader = new StreamReader(textPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    plain = reader.ReadToEnd();
                }
                else if (htmlPath != null && File.Exists(htmlPath))
                {
                    using StreamReader reader = new StreamReader(htmlPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    plain = Text.HtmlToText.Convert(reader.ReadToEnd());
                }

                if (string.IsNullOrWhiteSpace(plain))
                {
                    return null;
                }

                List<string> lines = new List<string>(2);
                foreach (string rawLine in plain!.Split('\n'))
                {
                    string line = rawLine.Trim().Trim('\uFEFF', '\uFFFE');
                    if (line.Length > 0)
                    {
                        lines.Add(line);
                        if (lines.Count == 2)
                        {
                            break;
                        }
                    }
                }

                if (lines.Count == 0)
                {
                    return null;
                }

                string excerpt = string.Join(" / ", lines);
                if (excerpt.Length > ExcerptMaxChars)
                {
                    excerpt = excerpt.Substring(0, ExcerptMaxChars);
                }

                return excerpt;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                return null;
            }
        }

        private static string? EmptyToNull(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
        }
    }
}
