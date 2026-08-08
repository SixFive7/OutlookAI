using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// One file accepted for attachment to a draft: the absolute path that will be handed
    /// to <c>Attachments.Add</c>, plus the name and size the caller gets reported back.
    /// </summary>
    public sealed class DraftAttachmentFile
    {
        /// <summary>Creates the accepted file record.</summary>
        public DraftAttachmentFile(string path, string fileName, long sizeBytes)
        {
            Path = path;
            FileName = fileName;
            SizeBytes = sizeBytes;
        }

        /// <summary>Absolute path on disk (rooted, existing, readable at validation time).</summary>
        public string Path { get; }

        /// <summary>File name as it will appear on the mail.</summary>
        public string FileName { get; }

        /// <summary>Size in bytes, measured at validation time.</summary>
        public long SizeBytes { get; }
    }

    /// <summary>
    /// PRE-COM validation of the draft tools' <c>attachments</c> argument (v3.MD D46/C3).
    /// The user granted NO path restrictions - any absolute path the server process can
    /// read may be attached - so the checks here are about the file being ATTACHABLE, not
    /// about where it lives: rooted path, exists, is a file and not a directory, is
    /// actually readable, and the set stays within a sane size/count budget.
    /// <para>
    /// FAIL-CLOSED AND WHOLE-SET (the deliberate design choice, D46): if ANY entry is bad
    /// the whole call is refused before a single COM object is touched, and the message
    /// names EVERY offending path with its own reason. Attaching the good subset would
    /// produce a mail that silently misses a file the agent believes it sent, which is
    /// worse than a refusal the agent can correct in one retry.
    /// </para>
    /// Pure logic apart from the file-system probes, so T1 pins it directly.
    /// </summary>
    public static class DraftAttachments
    {
        /// <summary>Maximum number of files one draft call may attach.</summary>
        public const int MaxFiles = 20;

        /// <summary>Maximum total size of one call's attachment set (bytes).</summary>
        public const long MaxTotalBytes = 150L * 1024 * 1024;

        /// <summary>Maximum number of names one update_draft call may remove.</summary>
        public const int MaxRemoveNames = 50;

        /// <summary>
        /// Validates the requested attachment paths. Returns an empty list for null/empty
        /// input (the parameter is optional everywhere it appears); throws
        /// <see cref="ArgumentException"/> listing every rejected path otherwise.
        /// </summary>
        public static IReadOnlyList<DraftAttachmentFile> Validate(IReadOnlyList<string>? paths, string parameterName = "attachments")
        {
            if (paths == null || paths.Count == 0)
            {
                return Array.Empty<DraftAttachmentFile>();
            }

            if (paths.Count > MaxFiles)
            {
                throw new ArgumentException(
                    parameterName + " holds " + paths.Count.ToString(CultureInfo.InvariantCulture)
                    + " paths; at most " + MaxFiles.ToString(CultureInfo.InvariantCulture)
                    + " files can be attached in one call.",
                    parameterName);
            }

            List<string> problems = new List<string>();
            List<DraftAttachmentFile> accepted = new List<DraftAttachmentFile>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            long total = 0;

            for (int i = 0; i < paths.Count; i++)
            {
                string raw = paths[i];
                if (string.IsNullOrWhiteSpace(raw))
                {
                    problems.Add("entry " + (i + 1).ToString(CultureInfo.InvariantCulture) + ": blank path");
                    continue;
                }

                string path = raw.Trim();
                string? problem = Describe(path, out string fullPath, out string fileName, out long size);
                if (problem != null)
                {
                    problems.Add("'" + path + "': " + problem);
                    continue;
                }

                if (!seen.Add(fullPath))
                {
                    problems.Add("'" + path + "': the same file is listed more than once");
                    continue;
                }

                total += size;
                accepted.Add(new DraftAttachmentFile(fullPath, fileName, size));
            }

            if (problems.Count > 0)
            {
                throw new ArgumentException(
                    "Nothing was attached and no draft was changed - " + parameterName + " has "
                    + problems.Count.ToString(CultureInfo.InvariantCulture) + " unusable entr"
                    + (problems.Count == 1 ? "y" : "ies") + ": " + string.Join("; ", problems)
                    + ". Supply ABSOLUTE paths to existing, readable files (directories cannot be attached).",
                    parameterName);
            }

            if (total > MaxTotalBytes)
            {
                throw new ArgumentException(
                    "The attachment set is " + FormatBytes(total) + ", over the "
                    + FormatBytes(MaxTotalBytes) + " limit for one call. Attach fewer or smaller files.",
                    parameterName);
            }

            return accepted;
        }

        /// <summary>
        /// Validates update_draft's <c>remove_attachments</c> names: plain file names as
        /// reported by read/the draft tools, trimmed, de-duplicated case-insensitively.
        /// </summary>
        public static IReadOnlyList<string> ValidateRemoveNames(IReadOnlyList<string>? names, string parameterName = "remove_attachments")
        {
            if (names == null || names.Count == 0)
            {
                return Array.Empty<string>();
            }

            if (names.Count > MaxRemoveNames)
            {
                throw new ArgumentException(
                    parameterName + " holds " + names.Count.ToString(CultureInfo.InvariantCulture)
                    + " names; at most " + MaxRemoveNames.ToString(CultureInfo.InvariantCulture) + " can be removed in one call.",
                    parameterName);
            }

            List<string> cleaned = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string raw in names)
            {
                if (string.IsNullOrWhiteSpace(raw))
                {
                    throw new ArgumentException(
                        parameterName + " contains a blank name. Use the attachment file names exactly as read/the draft tools report them.",
                        parameterName);
                }

                string name = raw.Trim();
                if (seen.Add(name))
                {
                    cleaned.Add(name);
                }
            }

            return cleaned;
        }

        /// <summary>Total size of an accepted set (bytes).</summary>
        public static long TotalBytes(IReadOnlyList<DraftAttachmentFile> files)
        {
            return files == null ? 0 : files.Sum(f => f.SizeBytes);
        }

        /// <summary>Human-readable byte size used in refusal messages and outcomes.</summary>
        public static string FormatBytes(long bytes)
        {
            if (bytes < 1024)
            {
                return bytes.ToString(CultureInfo.InvariantCulture) + " B";
            }

            if (bytes < 1024 * 1024)
            {
                return (bytes / 1024.0).ToString("F1", CultureInfo.InvariantCulture) + " KB";
            }

            return (bytes / (1024.0 * 1024.0)).ToString("F1", CultureInfo.InvariantCulture) + " MB";
        }

        private static string? Describe(string path, out string fullPath, out string fileName, out long size)
        {
            fullPath = string.Empty;
            fileName = string.Empty;
            size = 0;

            try
            {
                if (!Path.IsPathRooted(path))
                {
                    return "not an absolute path (the server has no working directory an agent can rely on)";
                }

                fullPath = Path.GetFullPath(path);
            }
            catch (ArgumentException)
            {
                return "not a usable path (invalid characters)";
            }
            catch (NotSupportedException)
            {
                return "not a usable path";
            }
            catch (PathTooLongException)
            {
                return "path is too long";
            }

            if (Directory.Exists(fullPath))
            {
                return "is a directory - directories cannot be attached, name the individual files";
            }

            if (!File.Exists(fullPath))
            {
                return "no such file";
            }

            try
            {
                FileInfo info = new FileInfo(fullPath);
                size = info.Length;
                fileName = info.Name;
            }
            catch (IOException ex)
            {
                return "could not be inspected (" + ex.GetType().Name + ")";
            }
            catch (UnauthorizedAccessException)
            {
                return "access denied";
            }

            if (size == 0)
            {
                return "is empty (0 bytes) - Outlook drops empty attachments";
            }

            // Readability is PROVED, not assumed: a path can exist and still be locked by
            // another process or denied by ACL, and Outlook would fail deep inside COM
            // with an opaque error after the draft was already half-built.
            try
            {
                using (FileStream stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    stream.ReadByte();
                }
            }
            catch (UnauthorizedAccessException)
            {
                return "cannot be read (access denied)";
            }
            catch (IOException ex)
            {
                return "cannot be read (" + ex.GetType().Name + " - it may be locked by another program)";
            }

            return null;
        }
    }
}
