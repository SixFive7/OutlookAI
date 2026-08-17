using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

// The Office major versions this product supports and the probe that finds the installed one.
// Services\OfficeVersions.cs is LINKED into this project (see OutlookAI.Core.csproj), so this is
// not a copy of the add-in's list - it is the add-in's list, probed in the add-in's order.
using OfficeVersions = global::OutlookAI.Services.OfficeVersions;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// The Outlook profile registry addresses this product depends on, in one place.
    /// <para>
    /// The accounts GUID subkey in particular was written out three times - twice in the
    /// signature code and once in a live test fixture - as a bare 32-character literal with
    /// no name attached. A typo in one copy is invisible on inspection and produces "no
    /// accounts found" rather than an error, so it exists here instead, named.
    /// </para>
    /// <para>
    /// The Office MAJOR in those addresses used to be a hardcoded 16.0, which was the same
    /// defect one level up: on Outlook 2013 (15.0) or a future 17.0 every read landed in a hive
    /// Outlook never writes, so accounts, signature defaults and the Outlook search settings all
    /// came back EMPTY - identical in the payload to a healthy machine with nothing configured,
    /// and with no diagnostic anywhere. It is now detected once per process from the registry,
    /// and <c>outlook_health</c> reports both the version found and the fact that none was.
    /// </para>
    /// </summary>
    public static class OutlookProfileRegistry
    {
        /// <summary>
        /// Office majors this product looks for, NEWEST FIRST - the probe order, so a machine
        /// with two Office versions installed is treated as the newer of them. The same array
        /// the add-in probes and the same one <c>Installer.iss</c>'s resiliency exemptions are
        /// checked against; public here because <c>HealthReporting</c>'s diagnostic has to name
        /// what it looked for, and the type holding it is internal (CS0436 - see that file).
        /// </summary>
        public static readonly IReadOnlyList<string> SupportedOfficeVersions =
            new ReadOnlyCollection<string>((string[])OfficeVersions.Supported.Clone());

        /// <summary>
        /// The major whose hives are read when detection finds nothing. Current Outlook, so a
        /// machine whose registry cannot be probed behaves exactly as the server did before
        /// detection existed - it just now says so instead of reporting an empty profile.
        /// </summary>
        public const string FallbackOfficeVersion = OfficeVersions.Fallback;

        /// <summary>
        /// The Office major actually registered on this machine, or null when NONE of
        /// <see cref="SupportedOfficeVersions"/> is. Null is the reportable state the health
        /// report turns into a problem line; every other caller wants
        /// <see cref="OfficeVersion"/>, which substitutes the fallback.
        /// <para>
        /// Detected ONCE per process. Outlook's major does not change under a running server,
        /// and the alternative is a registry open on every signature read.
        /// </para>
        /// </summary>
        public static readonly string? DetectedOfficeVersion;

        /// <summary>
        /// The Office major every registry path below is built from: the detected one, or
        /// <see cref="FallbackOfficeVersion"/> when nothing was detected.
        /// </summary>
        public static readonly string OfficeVersion;

        /// <summary>
        /// HKCU root of Outlook's own settings, for the Office major this machine actually has.
        /// No longer a <c>const</c>: it is computed from <see cref="OfficeVersion"/>, so callers
        /// that used to copy it into a <c>const</c> of their own hold a <c>static readonly</c>
        /// now.
        /// </summary>
        public static readonly string OutlookRootKeyPath;

        static OutlookProfileRegistry()
        {
            string detected;
            bool found = OfficeVersions.TryDetectOutlookVersion(out detected);
            DetectedOfficeVersion = found ? detected : null;
            OfficeVersion = detected;
            OutlookRootKeyPath = BuildOutlookRootKeyPath(detected);
        }

        /// <summary>Profiles container, relative to <see cref="OutlookRootKeyPath"/>.</summary>
        public const string ProfilesSubKeyName = "Profiles";

        /// <summary>
        /// The per-account container inside a profile. An Outlook-internal GUID, stable
        /// across every Outlook version this product supports, and not documented anywhere
        /// but the profile itself.
        /// </summary>
        public const string AccountsSubKeyName = "9375CFF0413111d3B88A00104B2A6676";

        /// <summary>
        /// The HKCU Outlook root for an arbitrary Office major. Pure, so the 15.0 and 17.0
        /// shapes are assertable on a machine that has neither installed.
        /// </summary>
        public static string BuildOutlookRootKeyPath(string officeVersion)
        {
            if (officeVersion == null)
            {
                throw new ArgumentNullException(nameof(officeVersion));
            }

            return OfficeVersions.OutlookKeyPath(officeVersion);
        }

        /// <summary>
        /// Whether an <c>...\Office\&lt;major&gt;\Outlook</c> key with these values and subkeys is
        /// a REAL Outlook hive, or just the shell this product's own installer leaves behind.
        /// <para>
        /// This is not pedantry, it is the whole detection. <c>Installer.iss</c> writes a
        /// resiliency exemption under EVERY supported major, and creating that value creates the
        /// <c>Outlook</c> key above it - so on every machine this product is installed on, all
        /// three keys exist and a bare "does the key exist?" probe answers the FIRST version
        /// tried, 16.0, on a 2013 or 17.0 machine as readily as on a 2016 one. That is the very
        /// defect the detection was added to remove. The rule: at least one value, or at least
        /// one subkey that is not <c>Resiliency</c>.
        /// </para>
        /// </summary>
        public static bool IsOutlookHive(string[] valueNames, string[] subKeyNames)
        {
            return OfficeVersions.IsOutlookHive(valueNames, subKeyNames);
        }

        /// <summary>
        /// The subkey name the rule above discounts, because this product's installer creates it
        /// for majors that are not installed. Public so a test can state the trap by name.
        /// </summary>
        public const string InstallerFootprintSubKeyName = OfficeVersions.InstallerFootprintSubKeyName;

        /// <summary>
        /// Detects the Office major from a caller-supplied "is there a real Outlook hive here?"
        /// predicate (handed the FULL HKCU path), returning null when none of
        /// <see cref="SupportedOfficeVersions"/> qualifies. The pure seam behind
        /// <see cref="DetectedOfficeVersion"/>: this machine has exactly one Office version
        /// really installed, so the 15.0, 17.0 and nothing-found paths are unreachable in a real
        /// registry without a second Outlook installation.
        /// </summary>
        public static string? DetectOfficeVersion(Func<string, bool> outlookKeyExists)
        {
            string detected;
            return OfficeVersions.TryDetectOutlookVersion(outlookKeyExists, out detected) ? detected : null;
        }
    }
}
