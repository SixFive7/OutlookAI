using System;

using Microsoft.Win32;

namespace OutlookAI.Services
{
    /// <summary>
    /// WHICH OFFICE MAJOR VERSIONS THIS PRODUCT KNOWS ABOUT, IN ONE PLACE.
    ///
    /// Office puts everything per-user under <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\</c>,
    /// and the major is not a number the code can infer - 16.0 is Outlook 2016 through
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
    /// The MCP server had the SAME defect, read-side: its Outlook Search key, that key's Policies
    /// mirror and the whole profile registry (accounts, signature defaults) were hardcoded to
    /// 16.0, so a 15.0 or 17.0 machine got an empty answer that looked exactly like a broken
    /// install. It shares this file rather than reimplementing it, which is what makes the two
    /// halves probe the SAME versions in the SAME order by construction:
    ///  - the add-in (<c>OutlookAI.csproj</c>) compiles this file directly;
    ///  - <c>OutlookAI.Core</c> LINKS it (see <c>OutlookAI.Core.csproj</c>) and re-exports what it
    ///    needs through <c>OutlookProfileRegistry</c>'s public surface.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL, and the intersection is narrow - the same one
    /// <c>Services\AddInServerContract.cs</c> documents: this compiles as net48 (the add-in, which
    /// sets no <c>LangVersion</c> and therefore gets C# 7.3, nullable off) and as net48 AND net10
    /// inside Core, nullable-enabled with warnings as errors. So no <c>string?</c>, no
    /// <c>return null</c> from a string member (which is why detection reports "nothing found" as
    /// a <c>bool</c> plus an <c>out</c> that is always assigned), no target-typed <c>new</c>, and
    /// nothing newer than C# 7.3.
    /// </para>
    ///
    /// <para>
    /// INTERNAL, and it has to stay internal. A PUBLIC type in a linked file compiled into two
    /// assemblies that can see each other is CS0436, an error here (<c>TreatWarningsAsErrors</c>).
    /// Core's copy is invisible to the test assembly precisely BECAUSE it is internal - Core
    /// grants no <c>InternalsVisibleTo</c> - so the tests reach this through
    /// <c>OutlookProfileRegistry</c> and <c>HealthReporting</c> instead. For the same reason this
    /// file is NOT compiled into <c>OutlookAI.McpServer</c>: that project DOES open its internals
    /// to the test assembly, so a copy there would collide with Core's.
    /// </para>
    ///
    /// <para>
    /// <see cref="Supported"/> is also checked against <c>Installer.iss</c>'s resiliency entries
    /// by <c>.github/scripts/check-pinned-constants.ps1</c> (#4), because the installer is Pascal
    /// and cannot read a C# constant.
    /// </para>
    /// </summary>
    internal static class OfficeVersions
    {
        /// <summary>
        /// Office major versions this product supports, NEWEST FIRST. The order is the probe
        /// order: the first one actually present on the machine wins, so a machine with both
        /// 2013 and 2016 installed is treated as the newer of the two.
        /// </summary>
        internal static readonly string[] Supported = { "16.0", "17.0", "15.0" };

        /// <summary>
        /// The version used when nothing can be detected. Current Outlook, so a machine whose
        /// registry has not been probed successfully behaves exactly as the add-in did before
        /// detection existed. Callers that need to TELL a fallback from a detection - the health
        /// report does, because "the hive is empty" and "there is no such hive" look identical
        /// from the outside - use <see cref="TryDetectOutlookVersion(out string)"/>, whose bool
        /// says which of the two happened.
        /// </summary>
        internal const string Fallback = "16.0";

        /// <summary>
        /// HKCU path of Outlook's own key for one Office major. One spelling of that hive root,
        /// shared by the probe below and by every path the server builds from the detected
        /// version, so the two cannot drift into reading different hives.
        /// </summary>
        internal static string OutlookKeyPath(string version)
        {
            return @"Software\Microsoft\Office\" + version + @"\Outlook";
        }

        /// <summary>
        /// The Office major version this Outlook is running as, or <see cref="Fallback"/> when
        /// none is detected. Convenience over <see cref="TryDetectOutlookVersion(out string)"/>
        /// for the callers that only need a hive to write into.
        /// </summary>
        internal static string DetectOutlookVersion()
        {
            string version;
            TryDetectOutlookVersion(out version);
            return version;
        }

        /// <summary>
        /// Detects the Office major version from which
        /// <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook</c> key is a REAL Outlook hive
        /// (<see cref="IsOutlookHive"/>, not merely present), newest first.
        /// Returns false when NONE of <see cref="Supported"/> is present, with
        /// <paramref name="version"/> set to <see cref="Fallback"/> - a reportable state, not an
        /// exception, because a probe that threw would take the health report down with it.
        ///
        /// Registry rather than <c>Application.Version</c> on purpose: the callers are
        /// registry-only services that must stay usable off the UI thread and without COM (the
        /// server has no Outlook object model at all when Outlook is closed), and Outlook writes
        /// this key as soon as it has a profile.
        /// </summary>
        internal static bool TryDetectOutlookVersion(out string version)
        {
            return TryDetectOutlookVersion(HasOutlookKey, out version);
        }

        /// <summary>
        /// The subkey this product's OWN installer creates under EVERY supported major, which is
        /// why the mere existence of an <c>...\Office\&lt;major&gt;\Outlook</c> key proves nothing.
        /// <c>Installer.iss</c> writes
        /// <c>...\Office\{16.0,15.0,17.0}\Outlook\Resiliency\DoNotDisableAddinList</c> on every
        /// install - deliberately, so the add-in survives a slow start on whichever Outlook is
        /// really there - and creating that value creates the whole path above it. Measured on
        /// this developer machine (2026-08-17): the 15.0 and 17.0 keys exist, each holding
        /// <c>Resiliency</c> and NOTHING else, while the real 16.0 key holds 31 subkeys and 4
        /// values. A probe that only asked "does the key exist?" would therefore answer 16.0 on
        /// every machine this product is installed on, INCLUDING the 2013 and 17.0 machines the
        /// detection exists for.
        /// </summary>
        internal const string InstallerFootprintSubKeyName = "Resiliency";

        /// <summary>
        /// Whether an <c>...\Office\&lt;major&gt;\Outlook</c> key is a real Outlook hive rather
        /// than the shell our own installer leaves behind: it must carry at least one VALUE, or
        /// at least one subkey that is not <see cref="InstallerFootprintSubKeyName"/>. A used
        /// Outlook has both in quantity (Profiles, Options, Setup, Security, DefaultProfile...);
        /// a version we merely wrote a resiliency exemption for has neither.
        /// <para>
        /// Pure, and public through <c>OutlookProfileRegistry.IsOutlookHive</c>, because this is
        /// the rule that decides the answer and the only way to test it against a hive shape this
        /// machine does not have.
        /// </para>
        /// </summary>
        internal static bool IsOutlookHive(string[] valueNames, string[] subKeyNames)
        {
            if (valueNames != null && valueNames.Length > 0)
            {
                return true;
            }

            if (subKeyNames == null)
            {
                return false;
            }

            foreach (string name in subKeyNames)
            {
                if (!string.Equals(name, InstallerFootprintSubKeyName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// The pure half of the detection above: same list, same order, same fallback, with the
        /// registry replaced by <paramref name="outlookKeyExists"/>, which is handed the FULL key
        /// path (<see cref="OutlookKeyPath"/>) rather than the bare version and answers the
        /// <see cref="IsOutlookHive"/> question rather than a bare key-exists one. This is the
        /// seam tests use - this machine has one Office version really installed, so the 15.0,
        /// 17.0 and nothing-found paths are not otherwise reachable without a second Outlook.
        /// A predicate that throws counts as "not present" and the probe carries on.
        /// </summary>
        internal static bool TryDetectOutlookVersion(Func<string, bool> outlookKeyExists, out string version)
        {
            if (outlookKeyExists == null)
            {
                throw new ArgumentNullException(nameof(outlookKeyExists));
            }

            foreach (string candidate in Supported)
            {
                bool present;
                try
                {
                    present = outlookKeyExists(OutlookKeyPath(candidate));
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    // A key we cannot look at tells us nothing; try the next version.
                    present = false;
                }

                if (present)
                {
                    version = candidate;
                    return true;
                }
            }

            version = Fallback;
            return false;
        }

        /// <summary>
        /// Live registry predicate: is there a REAL Outlook hive at this HKCU path? Deliberately
        /// not a bare key-exists check - see <see cref="InstallerFootprintSubKeyName"/> for why
        /// that answer would be yes for every supported major on every machine we ship to.
        /// </summary>
        private static bool HasOutlookKey(string keyPath)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key == null)
                {
                    return false;
                }

                return IsOutlookHive(key.GetValueNames(), key.GetSubKeyNames());
            }
        }
    }
}
