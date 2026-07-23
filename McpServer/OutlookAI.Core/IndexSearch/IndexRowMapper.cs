using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// Maps raw SystemIndex rows (column name -> provider value) to compact
    /// <see cref="IndexHit"/> records, including the EntryID decode and attachment
    /// parent mapping. Tolerant of nulls/DBNull and of provider type differences
    /// (System.Data.OleDb vs the ADODB COM fallback return arrays and integers in
    /// slightly different shapes).
    /// </summary>
    public static class IndexRowMapper
    {
        /// <summary>Maps one row. Rows without a parsable System.ItemUrl still map (URL fields null).</summary>
        public static IndexHit Map(IReadOnlyDictionary<string, object?> row)
        {
            if (row == null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            IndexHit hit = new IndexHit
            {
                ItemUrl = AsString(Get(row, "System.ItemUrl")) ?? string.Empty,
                Kinds = AsStringList(Get(row, "System.Kind")),
                Subject = AsString(Get(row, "System.Subject")),
                FromAddress = AsString(Get(row, "System.Message.FromAddress")),
                FromName = AsString(Get(row, "System.Message.FromName")),
                ToAddresses = AsStringList(Get(row, "System.Message.ToAddress")),
                DateReceivedUtc = AsUtcDateTime(Get(row, "System.Message.DateReceived")),
                ItemPathDisplay = AsString(Get(row, "System.ItemPathDisplay")),
                ItemNameDisplay = AsString(Get(row, "System.ItemNameDisplay")),
                AutoSummary = AsString(Get(row, "System.Search.AutoSummary")),
                SizeBytes = AsInt64(Get(row, "System.Size")),
                IsRead = AsBoolean(Get(row, "System.IsRead")),
                HasAttachments = AsBoolean(Get(row, "System.Message.HasAttachments")),
                ConversationId = AsString(Get(row, "System.Message.ConversationID")),
            };

            if (MapiItemUrl.TryParse(hit.ItemUrl, out MapiItemUrl? parsed) && parsed != null)
            {
                hit.StoreDisplayName = parsed.StoreDisplayName;
                hit.StorePrefix = parsed.StorePrefix;
                hit.StoreType = parsed.StoreType;
                hit.FolderSegments = parsed.FolderSegments;
                hit.IsAttachmentHit = parsed.IsAttachment;
                hit.AttachmentFileName = parsed.AttachmentFileName;
                hit.ParentItemUrl = parsed.ParentItemUrl;

                if (parsed.TryDecodeEntryId(out DecodedEntryId? decoded) && decoded != null)
                {
                    hit.EntryIdHex = decoded.EntryIdHex;
                    hit.StoreUidHex = decoded.StoreUidHex;
                }
            }

            return hit;
        }

        private static object? Get(IReadOnlyDictionary<string, object?> row, string column)
        {
            return row.TryGetValue(column, out object? value) ? value : null;
        }

        private static string? AsString(object? value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }

            if (value is string s)
            {
                return s;
            }

            if (value is byte[] bytes)
            {
                StringBuilder sb = new StringBuilder(bytes.Length * 2);
                foreach (byte b in bytes)
                {
                    sb.Append(b.ToString("X2", CultureInfo.InvariantCulture));
                }

                return sb.ToString();
            }

            if (value is Array array)
            {
                List<string> parts = new List<string>();
                foreach (object? element in array)
                {
                    string? part = AsString(element);
                    if (part != null)
                    {
                        parts.Add(part);
                    }
                }

                return parts.Count > 0 ? string.Join("; ", parts) : null;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static IReadOnlyList<string> AsStringList(object? value)
        {
            if (value == null || value is DBNull)
            {
                return Array.Empty<string>();
            }

            if (value is string single)
            {
                return new[] { single };
            }

            if (value is Array array)
            {
                List<string> result = new List<string>(array.Length);
                foreach (object? element in array)
                {
                    string? s = AsString(element);
                    if (s != null)
                    {
                        result.Add(s);
                    }
                }

                return result;
            }

            string? converted = AsString(value);
            return converted != null ? new[] { converted } : Array.Empty<string>();
        }

        private static DateTime? AsUtcDateTime(object? value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }

            if (value is DateTime dt)
            {
                // Windows Search SQL reports datetime values in UTC; providers hand them
                // over with Kind=Unspecified. Stamp UTC (verified live in T2).
                return dt.Kind == DateTimeKind.Unspecified
                    ? DateTime.SpecifyKind(dt, DateTimeKind.Utc)
                    : dt.ToUniversalTime();
            }

            return null;
        }

        private static long? AsInt64(object? value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }

            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            }
            catch (OverflowException)
            {
                return null;
            }
            catch (FormatException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
        }

        private static bool? AsBoolean(object? value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }

            if (value is bool b)
            {
                return b;
            }

            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
        }
    }
}
