using System;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// The names Windows will not keep as a file or folder name, written once (Q154): the device
    /// names it reads as a device in every folder. <c>manage_signature</c> refuses a signature named
    /// after one (<see cref="SignatureManager.ValidateName"/>); <c>save_attachment</c> renames a
    /// sender's file that is called after one, and refuses a target folder that is
    /// (<see cref="AttachmentFileWriter"/>).
    /// <para>
    /// <b>The rule</b>, from Microsoft's "Naming Files, Paths, and Namespaces": CON, PRN, AUX, NUL,
    /// COM0 to COM9 and LPT0 to LPT9 - the superscript digits 1, 2 and 3 count as digits there - are
    /// devices whatever their case, alone and followed by an extension: <c>NUL.txt</c> and
    /// <c>nul.tar.gz</c> are the NUL device too. Trailing spaces are dropped before the comparison,
    /// so <c>CON </c> is CON. <c>CONIN$</c> and <c>CONOUT$</c> open the console the same way and are
    /// treated alike. A file saved under such a name writes to the device and leaves nothing on disk,
    /// or opens a port.
    /// </para>
    /// </summary>
    public static class WindowsFileNames
    {
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
    }
}
