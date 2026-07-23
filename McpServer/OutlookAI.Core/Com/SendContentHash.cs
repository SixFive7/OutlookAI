using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Canonical content hash a send confirm-token is bound to (v3.MD D4): subject +
    /// recipients + plain-text body + the requested send-on-behalf-of address. Any
    /// change to the draft between token issue and confirmation changes the hash and
    /// invalidates the token. Pure logic (T1-tested); recomputed INSIDE the STA send
    /// operation right before <c>Send()</c> so the token->send gap is covered too.
    /// </summary>
    public static class SendContentHash
    {
        /// <summary>Computes the lowercase-hex SHA-256 canonical content hash.</summary>
        public static string Compute(
            string? subject,
            IReadOnlyList<ComRecipientInfo> recipients,
            string? bodyText,
            string? sentOnBehalfOf)
        {
            if (recipients == null)
            {
                throw new ArgumentNullException(nameof(recipients));
            }

            StringBuilder canonical = new StringBuilder(256);
            canonical.Append("subject=").Append(subject ?? string.Empty).Append('\n');

            // Recipient order as COM reports it is stable, but sort anyway so the hash
            // depends on the SET of recipients, not enumeration order.
            IEnumerable<string> recipientLines = recipients
                .Select(r => (r.Kind ?? "to") + ":" + ((r.Address ?? r.Name) ?? string.Empty).ToLowerInvariant())
                .OrderBy(line => line, StringComparer.Ordinal);
            foreach (string line in recipientLines)
            {
                canonical.Append("rcpt=").Append(line).Append('\n');
            }

            canonical.Append("onbehalf=").Append((sentOnBehalfOf ?? string.Empty).Trim().ToLowerInvariant()).Append('\n');
            canonical.Append("body=").Append(NormalizeBody(bodyText));

            byte[] bytes = Encoding.UTF8.GetBytes(canonical.ToString());
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(bytes);
                StringBuilder hex = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                {
                    hex.Append(b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
                }

                return hex.ToString();
            }
        }

        /// <summary>
        /// Line-ending normalization: the object model can report CRLF or LF for the
        /// same unchanged body depending on the read path - the hash must not flip on
        /// that.
        /// </summary>
        private static string NormalizeBody(string? bodyText)
        {
            if (string.IsNullOrEmpty(bodyText))
            {
                return string.Empty;
            }

            return bodyText!.Replace("\r\n", "\n").Replace('\r', '\n');
        }
    }
}
