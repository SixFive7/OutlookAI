using System;
using System.Collections;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Pure helpers for designated-Archive-folder resolution (D39, research 2026-07-26).
    ///
    /// How the one-click Archive folder is found (live-verified on all 5 stores of the
    /// dev machine, including the Dutch-localized delegate whose folder is
    /// "Archiveren"):
    ///
    ///  1. PRIMARY - <c>Store.GetDefaultFolder(39)</c>. 39 is an UNDOCUMENTED
    ///     OlDefaultFolders value (the public enum and this build's own type library
    ///     stop at olFolderSuggestedContacts=30) that Outlook's implementation
    ///     nevertheless honors: it returns exactly the mailbox's designated Archive
    ///     folder - the one the Archive button/Backspace, mobile swipe-archive and OWA
    ///     use, with the server-designated (localized) name. Because it is
    ///     undocumented, callers feature-detect per store (catch) and VERIFY the
    ///     result (same store, mail folder, not a core default folder) before use.
    ///  2. FALLBACK - the store object's PR_IPM_ARCHIVE_ENTRYID
    ///     (0x35FF0102, Exchange InternalSchema "ArchiveFolderEntryId") read via
    ///     PropertyAccessor, then opened with GetFolderFromID. Live probe result: the
    ///     property is NOT exposed on this machine's cached-Exchange stores (property
    ///     accessor reports unknown/not found on all 5), so the fallback exists for
    ///     other store configurations, not because it fires here.
    ///  3. NEVER by folder NAME - localization makes name guessing wrong by design.
    ///
    /// Researched-and-rejected carriers (documented so nobody re-walks this path):
    /// PR_ADDITIONAL_REN_ENTRYIDS_EX (0x36D90102, PersistData blocks per MS-OXOSFLD) -
    /// the documented PersistID list (RSF_PID_* up to 0x800B, spec 2024-04-16) has no
    /// archive value and the property is absent from all 5 live stores;
    /// PR_ADDITIONAL_REN_ENTRYIDS (0x36D81102, on the Inbox) - live-probed: carries
    /// exactly the classic 5 slots (Conflicts/Sync Issues/Local Failures/Server
    /// Failures/Junk) plus one 4-byte non-EntryID trailer, no archive slot.
    /// </summary>
    public static class ArchiveFolderResolution
    {
        /// <summary>
        /// Undocumented OlDefaultFolders value resolving the designated Archive folder
        /// (see class remarks; live-proven on this build, feature-detected per store).
        /// </summary>
        public const int OlFolderArchive = 39;

        /// <summary>PR_IPM_ARCHIVE_ENTRYID (ArchiveFolderEntryId) PropertyAccessor schema name - the fallback carrier.</summary>
        public const string ArchiveEntryIdPropertySchema = "http://schemas.microsoft.com/mapi/proptag/0x35FF0102";

        /// <summary>
        /// Interprets a PropertyAccessor.GetProperty result as ONE EntryID hex string:
        /// accepts a byte[] (PT_BINARY), the first non-empty byte[] of an array
        /// (PT_MV_BINARY), or an existing plausible hex string. Returns null for
        /// anything else (missing, empty, junk) - resolution then falls through to the
        /// per-item "no designated archive folder" error instead of guessing.
        /// </summary>
        public static string? TryReadEntryIdHex(object? propertyValue)
        {
            switch (propertyValue)
            {
                case null:
                    return null;
                case byte[] bytes:
                    return ToHexOrNull(bytes);
                case string text:
                    string trimmed = text.Trim();
                    return trimmed.Length >= 8 && trimmed.Length % 2 == 0 && IsHex(trimmed)
                        ? trimmed.ToUpperInvariant()
                        : null;
                case IEnumerable values:
                    foreach (object? entry in values)
                    {
                        if (entry is byte[] entryBytes)
                        {
                            string? hex = ToHexOrNull(entryBytes);
                            if (hex != null)
                            {
                                return hex;
                            }
                        }
                    }

                    return null;
                default:
                    return null;
            }
        }

        private static string? ToHexOrNull(byte[] bytes)
        {
            // A real folder EntryID is dozens of bytes; anything shorter than 4 is a
            // marker/filler (live-probed: the Inbox slot list ends with a 4-byte
            // non-EntryID trailer), not an id.
            if (bytes.Length < 4)
            {
                return null;
            }

            StringBuilder sb = new StringBuilder(bytes.Length * 2);
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }

        private static bool IsHex(string value)
        {
            foreach (char c in value)
            {
                bool ok = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!ok)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
