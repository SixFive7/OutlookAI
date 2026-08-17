using Microsoft.Win32;

namespace OutlookAI.Services
{
    /// <summary>
    /// WHICH OFFICE MAJOR VERSIONS THIS ADD-IN KNOWS ABOUT, IN ONE PLACE.
    ///
    /// Office puts everything per-user under <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\</c>,
    /// and the major is not a number the add-in can infer - 16.0 is Outlook 2016 through
    /// Microsoft 365, 15.0 is 2013, and 17.0 is whatever comes next. Three separate places used
    /// to spell that list out for themselves: the theme probe listed 16/17/15, the installer's
    /// resiliency exemption listed 16/15/17 with a comment claiming it matched "versions checked
    /// elsewhere", and the tuning service silently assumed 16.0 in all four of its registry
    /// paths. The last of those was a real defect rather than untidiness: on Outlook 2013 or a
    /// future 17.0 it wrote its values into a hive Outlook never reads, so the settings dialog
    /// showed every value as pending forever and the user was told to restart Outlook
    /// indefinitely.
    ///
    /// <para>
    /// <see cref="Supported"/> is that list, once. It is checked against
    /// <c>Installer.iss</c>'s resiliency entries by
    /// <c>.github/scripts/check-pinned-constants.ps1</c>, because the installer is Pascal and
    /// cannot read a C# constant.
    /// </para>
    /// </summary>
    internal static class OfficeVersions
    {
        /// <summary>
        /// Office major versions this add-in supports, NEWEST FIRST. The order is the probe
        /// order: the first one actually present on the machine wins, so a machine with both
        /// 2013 and 2016 installed is treated as the newer of the two.
        /// </summary>
        internal static readonly string[] Supported = { "16.0", "17.0", "15.0" };

        /// <summary>
        /// The version used when nothing can be detected. Current Outlook, so a machine whose
        /// registry has not been probed successfully behaves exactly as the add-in did before
        /// detection existed.
        /// </summary>
        internal const string Fallback = "16.0";

        /// <summary>
        /// The Office major version this Outlook is running as, detected from which
        /// <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook</c> key exists, newest first.
        ///
        /// Registry rather than <c>Application.Version</c> on purpose: the callers are
        /// registry-only services that must stay usable off the UI thread and without COM, and
        /// Outlook writes this key as soon as it has a profile. A machine where none of them
        /// exists gets <see cref="Fallback"/>.
        /// </summary>
        internal static string DetectOutlookVersion()
        {
            foreach (string version in Supported)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Office\" + version + @"\Outlook"))
                    {
                        if (key != null)
                            return version;
                    }
                }
                catch
                {
                    // A key we cannot open tells us nothing; try the next version.
                }
            }
            return Fallback;
        }
    }
}
