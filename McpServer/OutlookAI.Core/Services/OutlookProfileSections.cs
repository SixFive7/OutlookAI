using System;
using System.Globalization;
using System.Text;

using Microsoft.Win32;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// READ-ONLY access to one value Microsoft's store-hash code reads from an Outlook profile
    /// section: the account's <c>PR_MAPPING_SIGNATURE</c>, which is the documented input of a
    /// cached Exchange store's index hash (<see cref="OutlookAI.Core.Mapi.StoreHash"/>).
    /// <para>
    /// Microsoft's code reaches it through Extended MAPI - the store's
    /// <c>PR_EMSMDB_SECTION_UID</c>, then <c>IMAPISession::OpenProfileSection</c>. This server has
    /// no MAPI session of its own, so it reads where Outlook 2013 and later keep profile sections:
    /// <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;profile&gt;\&lt;section uid&gt;</c>,
    /// one value per property, named by its tag (<c>01020ff8</c> for <c>PR_MAPPING_SIGNATURE</c>).
    /// That layout is how Outlook persists the documented sections, not a documented interface in
    /// itself, so a value this cannot read is simply not a candidate; nothing here ever writes.
    /// NOT MEASURED: no test guest can have Exchange (Docs/live-tier-on-the-vm.md section 8
    /// item 24); <c>outlook_health</c> reports which input matched each store, which is how a
    /// first run on an Exchange machine says whether this one does.
    /// </para>
    /// </summary>
    public static class OutlookProfileSections
    {
        /// <summary>The value name a profile section stores <c>PR_MAPPING_SIGNATURE</c> (PT_BINARY 0x0FF8) under.</summary>
        public const string MappingSignatureValueName = "01020ff8";

        /// <summary>
        /// The section's key path relative to HKCU, or null when either part could reach outside
        /// the profile: a section uid must be exactly 32 hex digits, a profile name must hold no
        /// path separator. Pure.
        /// </summary>
        public static string? SectionKeyPath(string outlookRootKeyPath, string? profileName, string? sectionUidHex)
        {
            if (string.IsNullOrEmpty(outlookRootKeyPath) || string.IsNullOrEmpty(profileName)
                || profileName!.IndexOf('\\') >= 0 || profileName.IndexOf('/') >= 0
                || sectionUidHex == null || sectionUidHex.Length != 32)
            {
                return null;
            }

            foreach (char c in sectionUidHex)
            {
                bool hex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!hex)
                {
                    return null;
                }
            }

            return outlookRootKeyPath + "\\" + OutlookProfileRegistry.ProfilesSubKeyName + "\\" + profileName
                + "\\" + sectionUidHex.ToLowerInvariant();
        }

        /// <summary>
        /// The section's <c>PR_MAPPING_SIGNATURE</c> as upper-case hex, or null when the section,
        /// the value or the registry would not answer. Never throws for an unreadable profile: a
        /// missing candidate only means this route cannot vouch for the store.
        /// </summary>
        public static string? TryReadMappingSignatureHex(string? profileName, string? sectionUidHex)
        {
            string? path = SectionKeyPath(OutlookProfileRegistry.OutlookRootKeyPath, profileName, sectionUidHex);
            if (path == null)
            {
                return null;
            }

            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(path, writable: false))
                {
                    if (!(key?.GetValue(MappingSignatureValueName) is byte[] bytes) || bytes.Length == 0)
                    {
                        return null;
                    }

                    var hex = new StringBuilder(bytes.Length * 2);
                    foreach (byte b in bytes)
                    {
                        hex.Append(b.ToString("X2", CultureInfo.InvariantCulture));
                    }

                    return hex.ToString();
                }
            }
            catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException
                || ex is System.IO.IOException || ex is ArgumentException)
            {
                return null;
            }
        }
    }
}
