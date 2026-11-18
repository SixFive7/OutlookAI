using System;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// The names and paths Windows treats specially, written once (Q154): device names, device-namespace
    /// paths, a ':' that names a stream, names ending in a dot or a space, and the longest ordinary path.
    /// <c>manage_signature</c> refuses a signature named after a device
    /// (<see cref="SignatureManager.ValidateName"/>); <c>save_attachment</c> renames a sender's file that
    /// is called after one, and refuses a target folder that is or that breaks the other rules
    /// (<see cref="AttachmentFileWriter"/>); the draft tools refuse an attachment path that breaks them
    /// (<see cref="DraftAttachments"/>).
    /// <para>
    /// <b>Device names</b>, from Microsoft's "Naming Files, Paths, and Namespaces": CON, PRN, AUX, NUL,
    /// COM0 to COM9 and LPT0 to LPT9 - the superscript digits 1, 2 and 3 count as digits there - are
    /// devices whatever their case, alone and followed by an extension: <c>NUL.txt</c> and
    /// <c>nul.tar.gz</c> are the NUL device too. Trailing spaces are dropped before the comparison,
    /// so <c>CON </c> is CON. <c>CONIN$</c> and <c>CONOUT$</c> open the console the same way and are
    /// treated alike. On Windows 10 a file saved under such a name writes to the device and leaves
    /// nothing on disk, or opens a port. Windows 11 reads fewer of them so: measured 2026-11-18 on
    /// 10.0.26300, only NUL in a folder (<c>C:\x\NUL</c> is <c>\\.\NUL</c>, <c>C:\x\CON.txt</c> an
    /// ordinary file), and on the test guests' 10.0.26200 a file called <c>CON.txt</c> read back through
    /// its ordinary path. The rule is kept whole all the same: the documented one, for every Windows,
    /// and Outlook itself refuses to attach a file called <c>CON.txt</c>.
    /// </para>
    /// </summary>
    public static class WindowsFileNames
    {
        /// <summary>The longest path Windows' ordinary file calls take - 260 with the terminating null.</summary>
        public const int MaxPathChars = 259;

        private static readonly string[] DeviceNames =
        {
            "CON", "PRN", "AUX", "NUL",
            "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "COM¹", "COM²", "COM³",
            "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
            "LPT¹", "LPT²", "LPT³",
            "CONIN$", "CONOUT$",
        };

        /// <summary>
        /// True when Windows reads <paramref name="name"/> - one file or folder name, no separators -
        /// as a device: a device name alone, before an extension, or before trailing spaces.
        /// </summary>
        public static bool IsReservedDeviceName(string? name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            string stem = name!;
            int dot = stem.IndexOf('.');
            if (dot >= 0)
            {
                stem = stem.Substring(0, dot);
            }

            stem = stem.TrimEnd(' ');
            foreach (string device in DeviceNames)
            {
                if (string.Equals(stem, device, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// True when <paramref name="path"/> is written in Windows' device namespace, not as a drive or
        /// a share: <c>\\?\</c> and <c>\??\</c>, which hand the rest to Windows untouched, and <c>\\.\</c>
        /// - with either separator, so <c>//?/</c> and <c>//./</c> too. Windows' own path rules do not
        /// apply after <c>\\?\</c> or <c>\??\</c>: a name it would read as a device, one ending in a dot
        /// or a space, and a path longer than <see cref="MaxPathChars"/> are taken as written. A network
        /// share, <c>\\server\share</c>, is not one.
        /// </summary>
        public static bool IsDeviceNamespacePath(string? path)
        {
            if (path == null || path.Length < 4)
            {
                return false;
            }

            if (path[0] == '\\' && path[1] == '?' && path[2] == '?' && path[3] == '\\')
            {
                return true;
            }

            return IsSeparator(path[0]) && IsSeparator(path[1]) && (path[2] == '?' || path[2] == '.') && IsSeparator(path[3]);
        }

        /// <summary>
        /// True when <paramref name="path"/> holds a ':' that is not the one after the drive letter of
        /// a full path (<c>C:\</c>): after a file or folder name it names an alternate data stream of
        /// it (<c>report.pdf:hidden</c>, <c>report.pdf::$DATA</c>), and <c>C:report.pdf</c> is a path
        /// relative to whatever folder is current on that drive. The drive is looked for at the start
        /// only, so a device-namespace path is judged without its prefix (<c>C:\...</c> of <c>\\.\C:\...</c>).
        /// </summary>
        public static bool HasColonBeyondTheDrive(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            bool drive = path.Length >= 3 && IsAsciiLetter(path[0]) && path[1] == ':' && IsSeparator(path[2]);
            return path.IndexOf(':', drive ? 2 : 0) >= 0;
        }

        /// <summary>
        /// True when <paramref name="name"/> - one file or folder name - ends in a dot or a space. Windows
        /// drops those from every name of an ordinary path, so a file or folder made under such a name is
        /// not called that, and one that is called that only a device-namespace path can reach.
        /// </summary>
        public static bool EndsInDotOrSpace(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return name.Length > 0 && (name[name.Length - 1] == '.' || name[name.Length - 1] == ' ');
        }

        /// <summary>A drive letter: A to Z, either case, and nothing else.</summary>
        internal static bool IsAsciiLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        private static bool IsSeparator(char c)
        {
            return c == '\\' || c == '/';
        }
    }
}
