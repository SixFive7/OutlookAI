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
    /// actually readable, the set stays within a sane size/count budget, and the path is
    /// one Outlook can be handed.
    /// <para>
    /// FAIL-CLOSED AND WHOLE-SET (the deliberate design choice, D46): if ANY entry is bad
    /// the whole call is refused before a single COM object is touched, and the message
    /// names EVERY offending path with its own reason. Attaching the good subset would
    /// produce a mail that silently misses a file the agent believes it sent, which is
    /// worse than a refusal the agent can correct in one retry.
    /// </para>
    /// <para>
    /// <b>The paths Outlook cannot be handed</b>, measured on the unindexed test guest
    /// 2026-11-18 (Outlook 16.0.17932.20996, Windows 10.0.26200; runbook section 4.10), one draft
    /// per path, are refused on their text, before the disk is asked anything:
    /// a <c>\\?\</c> path crashed OUTLOOK.EXE inside <c>Attachments.Add</c> every time, four of four,
    /// whatever the file was called - an ordinary name, <c>CON.txt</c>, a name ending in dots and
    /// spaces, a path over 260 characters; a file called after a Windows device, an alternate data
    /// stream (<c>file.txt:stream</c>) and a path longer than <see cref="WindowsFileNames.MaxPathChars"/>
    /// Outlook refused, so a draft would have been made without them. <c>\\.\C:\...</c> and
    /// <c>\??\C:\...</c> attached the right file and stay accepted; the rest of the device namespace
    /// - pipes, devices, volumes, shares reached through it - names no file on a drive and is refused,
    /// and so is a <c>\??\</c> path to a name ending in a dot or a space, which only such a path can
    /// reach. Every refusal says what to give instead.
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

                if (!seen.Add(OnItsDrive(fullPath)))
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

                // On the text first, before the disk is asked anything - and again on the full path Windows
                // makes of it, which can be a device: C:\x\NUL becomes \\.\NUL (Windows 11, measured
                // 2026-11-18), and by Microsoft's naming rules C:\x\CON.txt becomes \\.\CON on Windows 10.
                string? refused = RefuseOnItsText(path);
                if (refused != null)
                {
                    return refused;
                }

                fullPath = Path.GetFullPath(path);
                refused = RefuseOnItsText(fullPath);
                if (refused != null)
                {
                    return refused;
                }
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

            if (fullPath.Length > WindowsFileNames.MaxPathChars)
            {
                return "is " + fullPath.Length.ToString(CultureInfo.InvariantCulture) + " characters long, and Outlook attaches a file "
                    + "only from a path of at most " + WindowsFileNames.MaxPathChars.ToString(CultureInfo.InvariantCulture)
                    + " - copy it to a shorter path with your own file tools, and attach the copy";
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

        /// <summary>
        /// Why Outlook cannot be handed <paramref name="path"/>, judged on its text alone (the class
        /// summary says what was measured); null when nothing on the text is wrong. Each reason ends in
        /// what to give instead. Internal so T1 can hold every spelling to it without a disk.
        /// </summary>
        internal static string? RefuseOnItsText(string path)
        {
            string onItsDrive = path;
            if (WindowsFileNames.IsDeviceNamespacePath(path))
            {
                // \\?\ and //?/ have the '?' third; \??\ has it second too.
                if (path[2] == '?' && path[1] != '?')
                {
                    string? ordinary = OrdinaryPathOf(path.Substring(4));
                    return @"is a \\?\ path, and Outlook crashes when it attaches a file from one - give the file's ordinary path instead"
                        + (ordinary == null ? @", such as 'C:\folder\file.pdf'" : ", '" + ordinary + "'");
                }

                string rest = path.Substring(4);
                if (!StartsWithDrive(rest))
                {
                    return @"is a device path that names no file on a drive - a pipe, a device, a volume or a share reached through \\.\ "
                        + @"or \??\ - give the file's ordinary path, such as 'C:\folder\file.pdf' or '\\server\share\file.pdf'";
                }

                // \??\ hands the rest to Windows untouched, so it reaches a name Windows drops the end of
                // anywhere else; \\.\ is normalized like an ordinary path.
                if (path[1] == '?' && WindowsFileNames.EndsInDotOrSpace(Path.GetFileName(rest)))
                {
                    return @"names a file whose name ends in a dot or a space, which only a device path can reach, and Outlook crashed "
                        + @"attaching such a file from a \\?\ path - copy it under a name without them, with your own file tools, and "
                        + "attach the copy";
                }

                onItsDrive = rest;
            }

            if (WindowsFileNames.HasColonBeyondTheDrive(onItsDrive))
            {
                return "holds a ':' that is not a drive letter's - it names an alternate data stream of a file, which Outlook will not "
                    + "attach, or a path relative to a drive's current folder - give the full path of the file itself, or copy the "
                    + "stream into a file of its own with your own file tools and attach that";
            }

            string name = Path.GetFileName(onItsDrive);
            if (WindowsFileNames.IsReservedDeviceName(name))
            {
                return "is called after a Windows device ('" + name + "'), and Outlook will not attach a file of that name - copy it "
                    + "under another name, such as '_" + name + "', with your own file tools, and attach the copy";
            }

            return null;
        }

        /// <summary>The ordinary path a <c>\\?\</c> path's rest names - <c>C:\...</c> or <c>\\server\share\...</c> - or null.</summary>
        private static string? OrdinaryPathOf(string rest)
        {
            if (StartsWithDrive(rest))
            {
                return rest;
            }

            if (rest.Length > 4 && rest.StartsWith("UNC", StringComparison.OrdinalIgnoreCase) && (rest[3] == '\\' || rest[3] == '/'))
            {
                return @"\\" + rest.Substring(4);
            }

            return null;
        }

        /// <summary>A path's comparable form: a <c>\\.\C:\</c> or <c>\??\C:\</c> path as its ordinary <c>C:\</c> path.</summary>
        private static string OnItsDrive(string fullPath)
        {
            return WindowsFileNames.IsDeviceNamespacePath(fullPath) && StartsWithDrive(fullPath.Substring(4))
                ? fullPath.Substring(4)
                : fullPath;
        }

        private static bool StartsWithDrive(string path)
        {
            return path.Length >= 3 && WindowsFileNames.IsAsciiLetter(path[0]) && path[1] == ':' && (path[2] == '\\' || path[2] == '/');
        }
    }
}
