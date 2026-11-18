using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// One signature as installed in the Outlook Signatures folder (data only).
    /// </summary>
    public sealed class SignatureInfo
    {
        /// <summary>Creates a signature descriptor.</summary>
        public SignatureInfo(
            string name,
            string? excerpt,
            string? htmlPath,
            string? rtfPath,
            string? textPath,
            string? resourceDirectory = null,
            int pictureCount = 0)
        {
            Name = name;
            Excerpt = excerpt;
            HtmlPath = htmlPath;
            RtfPath = rtfPath;
            TextPath = textPath;
            ResourceDirectory = resourceDirectory;
            PictureCount = pictureCount;
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

        /// <summary>The <c>&lt;name&gt;_files</c> folder holding the HTML rendition's pictures, when it exists.</summary>
        public string? ResourceDirectory { get; }

        /// <summary>How many picture files <see cref="ResourceDirectory"/> holds.</summary>
        public int PictureCount { get; }

        /// <summary>The best rendition for Word InsertFile: .htm, then .rtf, then .txt.</summary>
        public string? PreferredFilePath => HtmlPath ?? RtfPath ?? TextPath;

        /// <summary>
        /// The account this signature belongs to when its name is shaped like Outlook's local copy of
        /// a roaming signature, <c>Name (address)</c>; null for a signature of this device.
        /// </summary>
        public string? ScopeAccount => SignatureRoaming.AccountSuffixOf(Name);

        /// <summary>The renditions present, in the answers' vocabulary: html, rtf, text.</summary>
        public IReadOnlyList<string> Formats
        {
            get
            {
                List<string> formats = new List<string>(3);
                if (HtmlPath != null)
                {
                    formats.Add(SignatureFormats.Html);
                }

                if (RtfPath != null)
                {
                    formats.Add(SignatureFormats.Rtf);
                }

                if (TextPath != null)
                {
                    formats.Add(SignatureFormats.Text);
                }

                return formats;
            }
        }
    }

    /// <summary>The three renditions Outlook keeps of a signature, one per mail format, by name.</summary>
    public static class SignatureFormats
    {
        /// <summary>The .htm rendition, used by HTML mail.</summary>
        public const string Html = "html";

        /// <summary>The .rtf rendition, used by Rich Text mail.</summary>
        public const string Rtf = "rtf";

        /// <summary>The .txt rendition, used by plain-text mail.</summary>
        public const string Text = "text";

        /// <summary>All three, in the order they are reported.</summary>
        public static readonly IReadOnlyList<string> All = Array.AsReadOnly(new[] { Html, Rtf, Text });
    }

    /// <summary>
    /// Reads the signature files in the folder Outlook uses (<see cref="SignatureFolder"/>): names,
    /// renditions, pictures and short plain-text excerpts. Content is never logged. Pure
    /// filesystem - no COM, no registry beyond the folder value, safe on any machine, host-neutral
    /// (v3.MD section 0.5.2). Which account uses which signature is <see cref="SignatureEnvironment"/>.
    /// </summary>
    public static class SignatureCatalog
    {
        /// <summary>Maximum excerpt length (payload discipline).</summary>
        public const int ExcerptMaxChars = 160;

        /// <summary>Prefix of the throwaway test signatures live tests may create (S3 cleanup set).</summary>
        public const string TestSignaturePrefix = "OutlookAI-McpTest-";

        /// <summary>The suffix of a signature's picture folder: <c>&lt;name&gt;_files</c>.</summary>
        public const string ResourceDirectorySuffix = "_files";

        private static readonly string[] SignatureExtensions = { ".htm", ".html", ".rtf", ".txt" };

        private static readonly string[] PictureExtensions =
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".emz", ".wmz", ".svg", ".webp",
        };

        /// <summary>
        /// The Outlook signatures directory, resolved NOW from the value Outlook records
        /// (<see cref="SignatureFolder"/>) - <c>%APPDATA%\Microsoft\Signatures</c> only when none is.
        /// </summary>
        public static string DefaultSignatureDirectory => SignatureFolder.CurrentPath;

        /// <summary>True when a file name is a picture by its extension.</summary>
        public static bool IsPictureFile(string? fileName)
        {
            string extension = Path.GetExtension(fileName ?? string.Empty);
            return PictureExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Enumerates the installed signatures (name-sorted). Missing directory =>
        /// empty list. <paramref name="directory"/> overrides the resolved folder for tests.
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
                string resourceDirectory = Path.Combine(root, pair.Key + ResourceDirectorySuffix);
                bool hasResources = Directory.Exists(resourceDirectory);
                int pictures = hasResources ? CountPictures(resourceDirectory) : 0;
                result.Add(new SignatureInfo(
                    pair.Key, TryReadExcerpt(text, html, rtf), html, rtf, text, hasResources ? resourceDirectory : null, pictures));
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
        /// Short plain-text excerpt of a signature: the first two non-empty lines of the .txt
        /// rendition, else of the HTML converted to text, else of the RTF's text - each decoded the
        /// way it was written (<see cref="SignatureText"/>). Capped at <see cref="ExcerptMaxChars"/>.
        /// Null when nothing is readable.
        /// </summary>
        public static string? TryReadExcerpt(string? textPath, string? htmlPath, string? rtfPath = null)
        {
            try
            {
                string? plain = null;
                if (textPath != null && File.Exists(textPath))
                {
                    plain = SignatureText.ReadText(textPath).Text;
                }
                else if (htmlPath != null && File.Exists(htmlPath))
                {
                    plain = HtmlToText.Convert(SignatureText.ReadHtml(htmlPath).Text);
                }
                else if (rtfPath != null && File.Exists(rtfPath))
                {
                    plain = SignatureRtf.ToPlainText(SignatureText.ReadRtf(rtfPath).Text);
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
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException or ArgumentException)
            {
                return null;
            }
        }

        private static int CountPictures(string resourceDirectory)
        {
            try
            {
                return Directory.GetFiles(resourceDirectory).Count(f => IsPictureFile(Path.GetFileName(f)));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return 0;
            }
        }
    }
}
