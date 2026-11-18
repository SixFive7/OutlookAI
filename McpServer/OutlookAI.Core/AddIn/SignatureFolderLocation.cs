using System;
using System.IO;

using Microsoft.Win32;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// WHERE OUTLOOK KEEPS THE SIGNATURE FILES, IN ONE PLACE.
    /// <para>
    /// Outlook does not always use <c>%APPDATA%\Microsoft\Signatures</c>. It records the folder it
    /// uses in <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Common\General</c>, value
    /// <c>Signatures</c> (REG_SZ), as a path RELATIVE TO <c>%APPDATA%\Microsoft</c>; every profile of
    /// the Windows user shares that one folder. A localized Office names it in its own language
    /// (<c>Handtekeningen</c>, <c>Signaturen</c>), and an administrator or a deployment tool can point
    /// it anywhere. The folder used to be the literal <c>%APPDATA%\Microsoft\Signatures</c> in three
    /// places - the server's catalog, the add-in's task pane and a live-test fixture - so on such a
    /// machine every signature read and write went to a folder Outlook never looks at (research
    /// 2026-10-11, Q147 section 8 item 5).
    /// </para>
    /// <para>
    /// The value is read FRESH on every call: Outlook can change it while the server runs, and a
    /// cached folder would keep writing to the old one.
    /// </para>
    /// <para>
    /// One rule for both: OutlookAI.Core's, in the add-in namespace (Q163 (a)), which the add-in's
    /// task pane uses and Core re-exports through <c>SignatureFolder</c>. The add-in compiles as
    /// C# 7.3, so nothing here may need a newer language to call.
    /// </para>
    /// </summary>
    public static class SignatureFolderLocation
    {
        /// <summary>The subkey of Office's <c>Common</c> key that carries the folder value.</summary>
        public const string GeneralSubKeyName = "General";

        /// <summary>The value naming the folder, relative to <c>%APPDATA%\Microsoft</c>.</summary>
        public const string ValueName = "Signatures";

        /// <summary>The folder name when no value is recorded - what an English Outlook creates.</summary>
        public const string DefaultFolderName = "Signatures";

        /// <summary>HKCU path of the key holding <see cref="ValueName"/> for one Office major.</summary>
        public static string GeneralKeyPath(string officeVersion)
        {
            return OfficeVersions.CommonKeyPath(officeVersion) + @"\" + GeneralSubKeyName;
        }

        /// <summary>
        /// The folder a recorded value names, or the default folder when the value is blank or not a
        /// usable path. Relative values are taken from <c>&lt;appData&gt;\Microsoft</c>, as Outlook takes
        /// them; an absolute value stands as it is. Pure, so every shape is assertable without a
        /// registry. <paramref name="fromRecordedValue"/> says which of the two it was.
        /// </summary>
        public static string Resolve(string appDataRoaming, string recordedValue, out bool fromRecordedValue)
        {
            string microsoftDirectory = Path.Combine(appDataRoaming ?? string.Empty, "Microsoft");
            string value = recordedValue == null ? string.Empty : recordedValue.Trim().Trim('\0').Trim();
            if (value.Length > 0 && value.IndexOfAny(Path.GetInvalidPathChars()) < 0)
            {
                try
                {
                    string resolved = Path.GetFullPath(Path.Combine(microsoftDirectory, value));
                    fromRecordedValue = true;
                    return resolved;
                }
                catch (ArgumentException)
                {
                }
                catch (NotSupportedException)
                {
                }
                catch (PathTooLongException)
                {
                }
            }

            fromRecordedValue = false;
            return Path.Combine(microsoftDirectory, DefaultFolderName);
        }

        /// <summary>
        /// The recorded value as it stands in the registry, read now. False, with
        /// <paramref name="value"/> empty, when the key or the value is absent, blank or unreadable -
        /// a REG_EXPAND_SZ value comes back expanded, as Outlook would use it.
        /// </summary>
        public static bool TryReadRecordedValue(string officeVersion, out string value)
        {
            value = string.Empty;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(GeneralKeyPath(officeVersion), false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    var text = key.GetValue(ValueName) as string;
                    if (text == null || text.Trim().Length == 0)
                    {
                        return false;
                    }

                    value = text;
                    return true;
                }
            }
            catch (System.Security.SecurityException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            catch (IOException)
            {
                return false;
            }
        }

        /// <summary>The folder Outlook uses for this Office major, resolved now.</summary>
        public static string CurrentFolder(string officeVersion, out bool fromRecordedValue)
        {
            string recorded;
            TryReadRecordedValue(officeVersion, out recorded);
            return Resolve(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), recorded, out fromRecordedValue);
        }
    }
}
