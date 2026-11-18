using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
using OutlookAI.Core.AddIn;
using OutlookAI.Core.Interop;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Everything <c>save_attachment</c> writes to disk, and the check that keeps it inside
    /// OutlookAI's own download folder (Q198 (a), decided by the maintainer 2026-11-18). Nothing else
    /// in the product writes an attachment to a file: the folder check, making the folder, choosing a
    /// name no file has, the write itself and the clean-up after a write that failed all happen here,
    /// so this is the one class a rule about where files may be written has to allow for it (decision
    /// 173 (b)).
    /// <para>
    /// <b>Why only that folder.</b> The sender of a mail chooses an attachment's content and its
    /// name, and the agent that reads the mail chooses the folder. A mail could talk an agent into
    /// saving its attachment into the Startup folder, where Windows opens it at the next sign-in.
    /// So a save goes into <see cref="DownloadFolder"/> or a folder inside it and nowhere else. An
    /// agent that needs the file elsewhere copies it with its own file tools, under their own
    /// permission rules - and every refusal says so.
    /// </para>
    /// <para>
    /// <b>How a folder is checked.</b> First on its text, before the disk is asked anything: a
    /// network share or device path (<c>\\server\share</c>, <c>\\?\</c>, <c>\\.\</c>), a path
    /// relative to the current drive or folder (<c>\x</c>, <c>C:x</c>), a <c>:</c> that is not a
    /// drive's (an alternate data stream), a character no folder name can have, a name ending in a
    /// dot or a space (Windows drops those, so the folder made would not be the one named) and a
    /// device name (<see cref="WindowsFileNames"/>) are refused. <c>.</c> and <c>..</c> are resolved
    /// as Windows resolves them, on the text. Then on the path Windows itself reports: the deepest
    /// folder of the path that exists is opened and asked for its final path
    /// (<c>GetFinalPathNameByHandle</c>), which resolves every junction, symbolic link and 8.3 short
    /// name in it, and the rest of the path is put back on. That is compared with the download
    /// folder's own final path, found the same way. A folder a link leads out of is outside, however
    /// it is spelled; one that is really inside is inside, however it is spelled.
    /// </para>
    /// <para>
    /// <b>The write.</b> The folder is checked again on the disk right before anything is written,
    /// made, and checked a third time once it exists, and the file is written there under its final
    /// path, not under the spelling the agent gave. The name is the attachment's own, made safe by
    /// <see cref="SafeFileName"/>, and it never replaces anything: the name is claimed by creating an
    /// empty file that must not already exist (<c>CREATE_NEW</c>), without following a link of that
    /// name, and Outlook then writes the attachment into that file. A name already taken - by a file,
    /// a folder or a link - gets " (2)", " (3)" and so on. A write that fails leaves the empty file
    /// removed and a partly written one in place, and the caller is told where it is.
    /// </para>
    /// <para>
    /// <b>What it does not stop</b>: a program already running as the user that swaps a folder for a
    /// junction in the instant between the last check and the write. Such a program can write to the
    /// Startup folder itself, so it gains nothing from this one.
    /// </para>
    /// </summary>
    public static class AttachmentFileWriter
    {
        /// <summary>The longest path Windows' ordinary file calls take - 260 with the terminating null.</summary>
        public const int MaxPathChars = WindowsFileNames.MaxPathChars;

        /// <summary>The room a folder must leave for a file name, so a saved file stays within <see cref="MaxPathChars"/>.</summary>
        public const int MinFileNameChars = 32;

        /// <summary>The longest folder path a save goes into.</summary>
        public const int MaxFolderPathChars = MaxPathChars - 1 - MinFileNameChars;

        /// <summary>The longest single file or folder name NTFS stores.</summary>
        public const int MaxNameChars = 255;

        /// <summary>The longest ending treated as a file's extension when a name is shortened or numbered.</summary>
        public const int MaxExtensionChars = 16;

        /// <summary>The highest number a taken name is given before a save gives up.</summary>
        public const int MaxNumberedCopy = 9999;

        /// <summary>The name a file gets when the attachment's own is empty, or nothing of it is left once made safe.</summary>
        public const string FallbackFileName = "attachment";

        // Never in a folder name; ':' and the separators are dealt with separately, and characters
        // below U+0020 are refused wherever they appear.
        private const string InvalidPathCharacters = "\"<>|?*";

        // Never in a saved file's name: the path characters, ':' (an alternate data stream), and both
        // separators - the sender's name is one name, never a path.
        private const string InvalidNameCharacters = "\"<>|?*:\\/";

        private const char Separator = '\\';

        /// <summary>
        /// OutlookAI's download folder: <c>%LOCALAPPDATA%\OutlookAI\scratch\attachments</c>. Every
        /// attachment <c>save_attachment</c> saves is in it or in a folder inside it.
        /// </summary>
        public static string DownloadFolder => Path.Combine(SharedState.Root, "scratch", "attachments");

        /// <summary>
        /// The folder a save into <paramref name="targetDirectory"/> would write to - its final path,
        /// as Windows reports it - or <see cref="AttachmentFolderRefusedException"/> when it is not
        /// inside <see cref="DownloadFolder"/>. Reads the disk and writes nothing: a folder that does
        /// not exist yet is not made.
        /// </summary>
        /// <param name="targetDirectory">
        /// Null or blank for the download folder itself; a relative path for a folder inside it; or a
        /// full path that must lead into it.
        /// </param>
        public static string CheckFolder(string? targetDirectory)
        {
            return Resolve(targetDirectory, DownloadFolder, WindowsProbe.Instance).Folder;
        }

        /// <summary>
        /// Writes <paramref name="attachment"/> - an Outlook <c>Attachment</c> - into
        /// <paramref name="targetDirectory"/>, under its own file name made safe, and returns the file's
        /// final path and its size (0 when the size could not be read back). Refuses a folder outside
        /// <see cref="DownloadFolder"/> before anything is written, makes a folder that is missing, and
        /// never replaces a file.
        /// </summary>
        /// <param name="attachment">The attachment; its <c>FileName</c> is read and its <c>SaveAsFile</c> called.</param>
        /// <param name="targetDirectory">As <see cref="CheckFolder"/>.</param>
        /// <param name="reserved">
        /// Told the file's path once it is claimed and before anything is written into it, so a caller
        /// whose save then fails can say where the remains are.
        /// </param>
        public static (string Path, long SizeBytes) Save(object attachment, string? targetDirectory, Action<string>? reserved = null)
        {
            return Save(attachment, targetDirectory, reserved, DownloadFolder, WindowsProbe.Instance);
        }

        /// <summary>
        /// The sender's file name, made one name Windows stores as written: every character a file name
        /// cannot hold - the separators, <c>:</c>, <c>* ? " &lt; &gt; |</c>, control characters and the
        /// Unicode direction controls that can make <c>exe.pdf</c> read as <c>fdp.exe</c> - becomes
        /// <c>_</c>; leading and trailing spaces and trailing dots go, which Windows would drop anyway;
        /// a device name gets a leading <c>_</c>; and nothing left becomes <see cref="FallbackFileName"/>.
        /// </summary>
        public static string SafeFileName(string? attachmentFileName)
        {
            if (string.IsNullOrEmpty(attachmentFileName))
            {
                return FallbackFileName;
            }

            StringBuilder safe = new StringBuilder(attachmentFileName!.Length);
            foreach (char c in attachmentFileName)
            {
                safe.Append(IsUnsafeInName(c) ? '_' : c);
            }

            string name = safe.ToString().Trim(' ').TrimEnd('.', ' ');
            if (name.Length == 0)
            {
                return FallbackFileName;
            }

            return WindowsFileNames.IsReservedDeviceName(name) ? "_" + name : name;
        }

        /// <summary>The disk as Windows answers - what the product asks, and what T1's real-disk tests ask too.</summary>
        internal static IFolderProbe WindowsDisk => WindowsProbe.Instance;

        /// <summary>The check and the write, against any download folder and any way of asking the disk - T1's seam.</summary>
        internal static (string Path, long SizeBytes) Save(
            object attachment, string? targetDirectory, Action<string>? reserved, string downloadFolder, IFolderProbe probe)
        {
            if (attachment == null)
            {
                throw new ArgumentNullException(nameof(attachment));
            }

            ResolvedFolder resolved = Resolve(targetDirectory, downloadFolder, probe);
            string folder = resolved.Folder;

            // The check stands next to the write it guards, whatever Resolve answered.
            if (!IsWithin(folder, resolved.DownloadFolder))
            {
                throw Refuse(targetDirectory, "is outside OutlookAI's download folder", resolved.DownloadFolder);
            }

            Directory.CreateDirectory(folder);

            // Made or found: where it is NOW, before a byte goes into it. A folder of the path that a
            // junction replaced since the check is caught here, and nothing is written into it.
            string made = probe.FinalPath(folder);
            if (!IsWithin(made, resolved.DownloadFolder))
            {
                throw new IOException(
                    "The folder '" + folder + "' changed while it was being made: Windows now puts it at '" + made
                    + "', outside OutlookAI's download folder '" + resolved.DownloadFolder + "'. Nothing was saved in it.");
            }

            string path = ClaimFreeName(made, SafeFileName(ReadFileName(attachment)));
            reserved?.Invoke(path);

            bool written = false;
            try
            {
                ((dynamic)attachment).SaveAsFile(path);
                written = true;
            }
            finally
            {
                if (!written)
                {
                    DiscardIfEmpty(path);
                }
            }

            return (path, MeasureSaved(path));
        }

        /// <summary>The check alone, against any download folder and any way of asking the disk - T1's seam.</summary>
        internal static ResolvedFolder Resolve(string? targetDirectory, string downloadFolder, IFolderProbe probe)
        {
            if (downloadFolder == null)
            {
                throw new ArgumentNullException(nameof(downloadFolder));
            }

            if (probe == null)
            {
                throw new ArgumentNullException(nameof(probe));
            }

            string asked = (targetDirectory ?? string.Empty).Trim();
            if (!TrySplitRooted(downloadFolder, out string homePrefix, out List<string> homeParts))
            {
                throw new InvalidOperationException("OutlookAI's download folder '" + downloadFolder + "' is not a full path.");
            }

            string home = Join(homePrefix, homeParts);
            string lexical = asked.Length == 0 ? home : Lexical(asked, homePrefix, homeParts, home);
            if (lexical.Length > MaxFolderPathChars)
            {
                throw Refuse(asked, TooLong(lexical), home);
            }

            string realHome = Real(home, probe, asked, home);
            string real = asked.Length == 0 ? realHome : Real(lexical, probe, asked, realHome);
            if (!IsWithin(real, realHome))
            {
                string really = string.Equals(real, lexical, StringComparison.OrdinalIgnoreCase)
                    ? string.Empty
                    : " - through a junction, a symbolic link or a short name it is really '" + real + "'";
                throw Refuse(asked, "is outside OutlookAI's download folder" + really, realHome);
            }

            if (real.Length > MaxFolderPathChars)
            {
                throw Refuse(asked, TooLong(real), realHome);
            }

            return new ResolvedFolder(real, realHome);
        }

        /// <summary>
        /// True when <paramref name="path"/> is <paramref name="folder"/> or inside it. Both are final
        /// paths; case is ignored as Windows ignores it, and a sibling that merely starts with the same
        /// letters (<c>attachments-old</c>) is not inside.
        /// </summary>
        internal static bool IsWithin(string path, string folder)
        {
            if (string.Equals(path, folder, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string prefix = folder.EndsWith(@"\", StringComparison.Ordinal) ? folder : folder + @"\";
            return path.Length > prefix.Length && path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// <paramref name="stem"/> and <paramref name="ending"/> within <paramref name="budget"/>
        /// characters: the stem is shortened when it must, never in the middle of a surrogate pair, and
        /// never so that it ends in a dot or a space.
        /// </summary>
        internal static string Fit(string stem, string ending, int budget)
        {
            if (stem.Length + ending.Length <= budget)
            {
                return stem + ending;
            }

            int keep = budget - ending.Length;
            if (keep < 1)
            {
                // An ending this long is not an extension worth keeping (SplitExtension caps it), so
                // this is the whole name being cut.
                return Fit(stem + ending, string.Empty, budget);
            }

            if (char.IsHighSurrogate(stem[keep - 1]))
            {
                keep--;
            }

            string cut = stem.Substring(0, keep).TrimEnd(' ', '.');
            return (cut.Length == 0 ? "_" : cut) + ending;
        }

        // ------------------------------------------------------------------ the folder, on its text

        private static string Lexical(string asked, string homePrefix, List<string> homeParts, string home)
        {
            // The two prefixes first: \\?\ and \\.\ hold characters the scan below refuses too, and
            // "a device path" is the answer that says what is wrong with them.
            string path = asked.Replace('/', Separator);
            if (path.StartsWith(@"\\", StringComparison.Ordinal))
            {
                throw Refuse(asked, @"is a network share or a device path (\\server\share, \\?\, \\.\)", home);
            }

            if (path[0] == Separator)
            {
                throw Refuse(asked, "starts at the root of whatever drive is current, which this server does not have", home);
            }

            foreach (char c in path)
            {
                if (c < ' ' || InvalidPathCharacters.IndexOf(c) >= 0)
                {
                    throw Refuse(asked, "holds a character no folder name can have (" + Describe(c) + ")", home);
                }
            }

            bool absolute = path.Length >= 3 && WindowsFileNames.IsAsciiLetter(path[0]) && path[1] == ':' && path[2] == Separator;
            if (WindowsFileNames.HasColonBeyondTheDrive(path))
            {
                throw Refuse(
                    asked,
                    "holds a ':' that is not a drive letter's - it would name an alternate data stream of a file, "
                    + "or a folder relative to a drive's current folder",
                    home);
            }

            string prefix;
            List<string> parts;
            string rest;
            if (absolute)
            {
                prefix = char.ToUpperInvariant(path[0]) + ":";
                parts = new List<string>();
                rest = path.Substring(3);
            }
            else
            {
                prefix = homePrefix;
                parts = new List<string>(homeParts);
                rest = path;
            }

            foreach (string part in rest.Split(Separator))
            {
                if (part.Length == 0 || part == ".")
                {
                    continue;
                }

                if (part == "..")
                {
                    if (parts.Count == 0)
                    {
                        throw Refuse(asked, "climbs above the root of its drive", home);
                    }

                    parts.RemoveAt(parts.Count - 1);
                    continue;
                }

                if (WindowsFileNames.EndsInDotOrSpace(part))
                {
                    throw Refuse(
                        asked,
                        "has a folder name ending in a dot or a space ('" + part + "'), which Windows drops - the folder made would not be the one named",
                        home);
                }

                if (WindowsFileNames.IsReservedDeviceName(part))
                {
                    throw Refuse(asked, "names a Windows device ('" + part + "'), not a folder", home);
                }

                if (part.Length > MaxNameChars)
                {
                    throw Refuse(asked, "has a folder name longer than " + MaxNameChars.ToString(CultureInfo.InvariantCulture) + " characters", home);
                }

                parts.Add(part);
            }

            return Join(prefix, parts);
        }

        // ------------------------------------------------------------------ the folder, on the disk

        /// <summary>
        /// Where Windows puts <paramref name="lexical"/>: the final path of its deepest folder that
        /// exists, with the part that does not exist yet put back on as written.
        /// </summary>
        private static string Real(string lexical, IFolderProbe probe, string asked, string home)
        {
            string current = lexical;
            List<string> tail = new List<string>();
            while (true)
            {
                FolderEntry entry = probe.Inspect(current, out int win32Error);
                if (entry == FolderEntry.Directory)
                {
                    break;
                }

                if (entry == FolderEntry.NotADirectory)
                {
                    throw Refuse(asked, "cannot be made: '" + current + "' is a file, not a folder", home);
                }

                if (entry == FolderEntry.Unreadable)
                {
                    throw Refuse(
                        asked,
                        "cannot be checked: Windows would not say what '" + current + "' is ("
                        + new Win32Exception(win32Error).Message + ")",
                        home);
                }

                if (!TrySplitRooted(current, out string prefix, out List<string> parts) || parts.Count == 0)
                {
                    throw Refuse(asked, "is on a drive that is not there", home);
                }

                tail.Insert(0, parts[parts.Count - 1]);
                parts.RemoveAt(parts.Count - 1);
                current = Join(prefix, parts);
            }

            string final;
            try
            {
                final = probe.FinalPath(current);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                throw Refuse(
                    asked,
                    "cannot be checked: Windows could not say where '" + current + "' really is (" + ex.Message
                    + ") - a link that leads nowhere, or a folder this account may not open",
                    home);
            }

            string result = final;
            foreach (string part in tail)
            {
                result = Append(result, part);
            }

            return result;
        }

        /// <summary>
        /// Splits a full path into its root - <c>C:</c>, or <c>\\server\share</c> - and its folder
        /// names, with <c>.</c> and <c>..</c> resolved. False for anything else.
        /// </summary>
        private static bool TrySplitRooted(string path, out string prefix, out List<string> parts)
        {
            parts = new List<string>();
            string normalized = path.Replace('/', Separator);
            string rest;
            if (normalized.Length >= 3 && WindowsFileNames.IsAsciiLetter(normalized[0]) && normalized[1] == ':' && normalized[2] == Separator)
            {
                prefix = char.ToUpperInvariant(normalized[0]) + ":";
                rest = normalized.Substring(3);
            }
            else if (normalized.Length == 2 && WindowsFileNames.IsAsciiLetter(normalized[0]) && normalized[1] == ':')
            {
                prefix = char.ToUpperInvariant(normalized[0]) + ":";
                rest = string.Empty;
            }
            else if (normalized.StartsWith(@"\\", StringComparison.Ordinal)
                && !normalized.StartsWith(@"\\?\", StringComparison.Ordinal)
                && !normalized.StartsWith(@"\\.\", StringComparison.Ordinal))
            {
                string[] unc = normalized.Substring(2).Split(new[] { Separator }, 3);
                if (unc.Length < 2 || unc[0].Length == 0 || unc[1].Length == 0)
                {
                    prefix = string.Empty;
                    return false;
                }

                prefix = @"\\" + unc[0] + @"\" + unc[1];
                rest = unc.Length == 3 ? unc[2] : string.Empty;
            }
            else
            {
                prefix = string.Empty;
                return false;
            }

            foreach (string part in rest.Split(Separator))
            {
                if (part.Length == 0 || part == ".")
                {
                    continue;
                }

                if (part == "..")
                {
                    if (parts.Count > 0)
                    {
                        parts.RemoveAt(parts.Count - 1);
                    }

                    continue;
                }

                parts.Add(part);
            }

            return true;
        }

        private static string Join(string prefix, List<string> parts)
        {
            string joined = string.Join(@"\", parts);
            if (prefix.EndsWith(":", StringComparison.Ordinal))
            {
                return prefix + @"\" + joined;
            }

            return parts.Count == 0 ? prefix : prefix + @"\" + joined;
        }

        private static string Append(string folder, string name)
        {
            return folder.EndsWith(@"\", StringComparison.Ordinal) ? folder + name : folder + @"\" + name;
        }

        // ------------------------------------------------------------------ the file

        private static string? ReadFileName(object attachment)
        {
            try
            {
                return (string?)((dynamic)attachment).FileName;
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>
        /// Claims a name in <paramref name="folder"/> by creating an empty file under it - one that
        /// must not exist, through no link - and returns its path. A name that is taken gets the next
        /// number; a name too long for the folder is shortened, keeping its extension.
        /// </summary>
        private static string ClaimFreeName(string folder, string name)
        {
            int budget = Math.Min(MaxNameChars, MaxPathChars - folder.Length - 1);
            if (budget < MinFileNameChars)
            {
                throw new IOException(
                    "The folder '" + folder + "' leaves no room for a file name within Windows' " + (MaxPathChars + 1).ToString(CultureInfo.InvariantCulture)
                    + "-character limit on a path. Nothing was written.");
            }

            SplitExtension(name, out string stem, out string extension);
            for (int copy = 1; copy <= MaxNumberedCopy; copy++)
            {
                string ending = (copy == 1 ? string.Empty : " (" + copy.ToString(CultureInfo.InvariantCulture) + ")") + extension;
                string candidate = Append(folder, Fit(stem, ending, budget));
                int error = WindowsProbe.TryCreateNew(candidate);
                if (error == 0)
                {
                    return candidate;
                }

                // A file answers "exists"; a folder or a link of that name answers "access denied" -
                // measured 2026-11-18 - and so does a real refusal, which is told apart by asking
                // whether anything is there.
                if (error == WindowsProbe.ErrorFileExists
                    || error == WindowsProbe.ErrorAlreadyExists
                    || (error == WindowsProbe.ErrorAccessDenied && WindowsProbe.EntryExists(candidate)))
                {
                    continue;
                }

                throw WindowsProbe.Failure(error, candidate);
            }

            throw new IOException(
                "Every name from '" + name + "' to its copy (" + MaxNumberedCopy.ToString(CultureInfo.InvariantCulture)
                + ") is taken in '" + folder + "'. Nothing was written.");
        }

        private static void SplitExtension(string name, out string stem, out string extension)
        {
            int dot = name.LastIndexOf('.');
            if (dot > 0 && name.Length - dot <= MaxExtensionChars && name.IndexOf(' ', dot) < 0)
            {
                stem = name.Substring(0, dot);
                extension = name.Substring(dot);
                return;
            }

            stem = name;
            extension = string.Empty;
        }

        /// <summary>
        /// After a failed write: removes the file this class claimed when nothing went into it, and
        /// leaves it when something did - a partial file is reported, never silently deleted.
        /// Best-effort: a file it cannot remove stays, and the caller reports its path either way.
        /// </summary>
        private static void DiscardIfEmpty(string path)
        {
            try
            {
                FileInfo claimed = new FileInfo(path);
                if (claimed.Exists
                    && (claimed.Attributes & (FileAttributes.ReparsePoint | FileAttributes.Directory)) == 0
                    && claimed.Length == 0)
                {
                    claimed.Delete();
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        /// <summary>
        /// The saved file's size. Past the write the file is saved, and measuring it is reporting,
        /// not saving: a size that cannot be read is 0 and never turns a completed save into a
        /// reported failure.
        /// </summary>
        private static long MeasureSaved(string path)
        {
            try
            {
                return new FileInfo(path).Length;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
            {
                return 0;
            }
        }

        // ------------------------------------------------------------------ words

        private static AttachmentFolderRefusedException Refuse(string? asked, string reason, string downloadFolder)
        {
            string subject = string.IsNullOrWhiteSpace(asked)
                ? "OutlookAI's download folder"
                : "target_dir '" + asked!.Trim() + "'";
            return new AttachmentFolderRefusedException(
                subject + " " + reason + ". Nothing was written. save_attachment saves only into OutlookAI's download folder, '"
                + downloadFolder + "', and folders inside it, because the sender chose the file and its name. Leave target_dir "
                + "out, or name a folder inside it, such as 'invoices'. To use the file somewhere else, save it there and copy "
                + "it with your own file tools, under their own permission rules.");
        }

        private static string TooLong(string path)
        {
            return "is too long: a file saved in '" + path + "' would pass Windows' "
                + (MaxPathChars + 1).ToString(CultureInfo.InvariantCulture) + "-character limit on a path";
        }

        private static string Describe(char c)
        {
            return c < ' '
                ? "the control character U+" + ((int)c).ToString("X4", CultureInfo.InvariantCulture)
                : "'" + c + "'";
        }

        private static bool IsUnsafeInName(char c)
        {
            return c < ' ' || c == '\u007F' || InvalidNameCharacters.IndexOf(c) >= 0 || IsDirectionControl(c);
        }

        // The Unicode controls that reorder what follows them - U+202E turns "invoice<U+202E>fdp.exe" into
        // something that reads "invoiceexe.pdf" in Explorer.
        private static bool IsDirectionControl(char c)
        {
            return c == '\u061C' || c == '\u200E' || c == '\u200F'
                || (c >= '\u202A' && c <= '\u202E')
                || (c >= '\u2066' && c <= '\u2069');
        }

        // ------------------------------------------------------------------ Windows

        /// <summary>
        /// The disk as Windows answers, through its own calls: <c>GetFileAttributesW</c>, which does not
        /// follow a link at the end of the path, so a link reads as what it is; and
        /// <c>GetFinalPathNameByHandle</c> on a handle opened for no access at all, which follows every
        /// one. Declared here and nowhere else, so every file this class creates is created in it.
        /// </summary>
        private sealed class WindowsProbe : IFolderProbe
        {
            internal const int ErrorFileNotFound = 2;
            internal const int ErrorPathNotFound = 3;
            internal const int ErrorAccessDenied = 5;
            internal const int ErrorFileExists = 80;
            internal const int ErrorAlreadyExists = 183;

            private const uint InvalidFileAttributes = 0xFFFFFFFF;
            private const uint FileAttributeDirectory = 0x10;
            private const uint FileAttributeNormal = 0x80;
            private const uint GenericWrite = 0x40000000;
            private const uint ShareAll = 0x00000007;
            private const uint CreateNew = 1;
            private const uint OpenExisting = 3;
            private const uint FileFlagBackupSemantics = 0x02000000;
            private const uint FileFlagOpenReparsePoint = 0x00200000;

            internal static readonly WindowsProbe Instance = new WindowsProbe();

            public FolderEntry Inspect(string path, out int win32Error)
            {
                uint attributes = GetFileAttributesW(path);
                if (attributes == InvalidFileAttributes)
                {
                    win32Error = Marshal.GetLastWin32Error();
                    return win32Error == ErrorFileNotFound || win32Error == ErrorPathNotFound
                        ? FolderEntry.Missing
                        : FolderEntry.Unreadable;
                }

                win32Error = 0;
                return (attributes & FileAttributeDirectory) != 0 ? FolderEntry.Directory : FolderEntry.NotADirectory;
            }

            public string FinalPath(string directory)
            {
                // Access 0: the handle answers questions about the folder and can do nothing else.
                // Backup semantics is what lets CreateFileW open a folder at all.
                using (SafeFileHandle handle = Kernel32.CreateFileW(directory, 0, ShareAll, IntPtr.Zero, OpenExisting, FileFlagBackupSemantics, IntPtr.Zero))
                {
                    if (handle.IsInvalid)
                    {
                        throw Failure(Marshal.GetLastWin32Error(), directory);
                    }

                    char[] buffer = new char[512];
                    uint length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Length, 0);
                    if (length >= buffer.Length)
                    {
                        buffer = new char[length + 1];
                        length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Length, 0);
                    }

                    if (length == 0 || length >= buffer.Length)
                    {
                        throw Failure(Marshal.GetLastWin32Error(), directory);
                    }

                    return WithoutWin32Prefix(new string(buffer, 0, (int)length));
                }
            }

            /// <summary>0 when <paramref name="path"/> was created, empty, by this call; otherwise the Win32 error.</summary>
            internal static int TryCreateNew(string path)
            {
                // CREATE_NEW fails when anything has the name; FILE_FLAG_OPEN_REPARSE_POINT makes a link
                // of that name count as the name rather than lead the new file somewhere else.
                using (SafeFileHandle handle = Kernel32.CreateFileW(
                    path, GenericWrite, ShareAll, IntPtr.Zero, CreateNew, FileAttributeNormal | FileFlagOpenReparsePoint, IntPtr.Zero))
                {
                    return handle.IsInvalid ? Marshal.GetLastWin32Error() : 0;
                }
            }

            internal static bool EntryExists(string path)
            {
                return GetFileAttributesW(path) != InvalidFileAttributes;
            }

            internal static Exception Failure(int win32Error, string path)
            {
                string message = new Win32Exception(win32Error).Message + " (" + path + ")";
                switch (win32Error)
                {
                    case ErrorAccessDenied:
                        return new UnauthorizedAccessException(message);
                    case ErrorFileNotFound:
                        return new FileNotFoundException(message, path);
                    case ErrorPathNotFound:
                        return new DirectoryNotFoundException(message);
                    default:
                        return new IOException(message, unchecked((int)0x80070000 | (win32Error & 0xFFFF)));
                }
            }

            private static string WithoutWin32Prefix(string final)
            {
                string path = final;
                if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
                {
                    path = @"\\" + path.Substring(8);
                }
                else if (path.StartsWith(@"\\?\", StringComparison.Ordinal))
                {
                    path = path.Substring(4);
                }

                // A drive's root keeps its separator ("C:\"); every other folder loses a trailing one.
                if (path.Length > 3 && path.EndsWith(@"\", StringComparison.Ordinal))
                {
                    path = path.TrimEnd(Separator);
                }

                return path;
            }

            [DllImport("kernel32.dll", EntryPoint = "GetFileAttributesW", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern uint GetFileAttributesW(string lpFileName);

            [DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW", CharSet = CharSet.Unicode, SetLastError = true)]
            private static extern uint GetFinalPathNameByHandleW(
                SafeFileHandle hFile, [Out] char[] lpszFilePath, uint cchFilePath, uint dwFlags);
        }
    }

    /// <summary>
    /// <c>save_attachment</c>'s refusal of a folder that is not inside OutlookAI's download folder
    /// (Q198 (a)) - always decided before anything is written. The message says why, names the
    /// download folder, and tells the agent it can copy the file elsewhere with its own tools.
    /// </summary>
    public sealed class AttachmentFolderRefusedException : ArgumentException
    {
        /// <summary>Creates the refusal.</summary>
        public AttachmentFolderRefusedException(string message)
            : base(message)
        {
        }
    }

    /// <summary>What is at a path, as <see cref="IFolderProbe.Inspect"/> answers.</summary>
    internal enum FolderEntry
    {
        /// <summary>Nothing has the name, or a folder above it is missing.</summary>
        Missing,

        /// <summary>A folder - or a link to one, which is followed later.</summary>
        Directory,

        /// <summary>A file.</summary>
        NotADirectory,

        /// <summary>Windows would not say: access denied, a bad name, a device.</summary>
        Unreadable,
    }

    /// <summary>How <see cref="AttachmentFileWriter"/> asks the disk about a folder - Windows in the product, a stand-in in T1.</summary>
    internal interface IFolderProbe
    {
        /// <summary>What is at <paramref name="path"/>, without following a link at its end.</summary>
        FolderEntry Inspect(string path, out int win32Error);

        /// <summary>
        /// The final path of the existing folder <paramref name="directory"/> - every junction,
        /// symbolic link and short name in it resolved - without a <c>\\?\</c> prefix. Throws
        /// <see cref="IOException"/> or <see cref="UnauthorizedAccessException"/> when Windows cannot say.
        /// </summary>
        string FinalPath(string directory);
    }

    /// <summary>A checked folder: where a save would write, and the download folder it is inside, both as final paths.</summary>
    internal sealed class ResolvedFolder
    {
        internal ResolvedFolder(string folder, string downloadFolder)
        {
            Folder = folder;
            DownloadFolder = downloadFolder;
        }

        /// <summary>The folder's final path.</summary>
        internal string Folder { get; }

        /// <summary>The download folder's final path.</summary>
        internal string DownloadFolder { get; }
    }
}
