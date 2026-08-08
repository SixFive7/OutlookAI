using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Canonical content hash a send confirm-token is bound to (v3.MD D4, widened by
    /// D46): subject + recipients + plain-text body + the requested send-on-behalf-of
    /// address + THE ATTACHMENT SET + a digest of the HTML body. Any change to the draft
    /// between token issue and confirmation changes the hash and invalidates the token.
    /// Pure logic (T1-tested); recomputed INSIDE the STA send operation right before
    /// <c>Send()</c> so the token-&gt;send gap is covered too.
    /// <para>
    /// WHY ATTACHMENTS ARE IN HERE (D46/C3, the mandatory interlock): once drafts can
    /// carry files, "the content the user confirmed" includes them. A hash over text
    /// alone would let a file be added to - or removed from - a draft AFTER the user
    /// confirmed the send, and the token would still validate. Names AND sizes are
    /// hashed, so a same-named file swapped for different content also invalidates.
    /// </para>
    /// <para>
    /// WHY AN HTML DIGEST IS IN HERE (D46/C3): Outlook derives <c>MailItem.Body</c> from
    /// <c>HTMLBody</c>, and that derivation is LOSSY - markup-only edits (a hyperlink
    /// retargeted while its visible label stays, text recoloured to invisible, a table
    /// restructured) leave the plain text byte-identical, so a body-text-only hash could
    /// not see them. The HTML is folded in as a SHA-256 digest rather than raw markup so
    /// the state snapshot stays small; a null digest (HTML unreadable) is canonicalized
    /// as empty and simply contributes nothing.
    /// </para>
    /// </summary>
    public static class SendContentHash
    {
        /// <summary>Computes the lowercase-hex SHA-256 canonical content hash.</summary>
        public static string Compute(
            string? subject,
            IReadOnlyList<ComRecipientInfo> recipients,
            string? bodyText,
            string? sentOnBehalfOf,
            IReadOnlyList<ComAttachmentInfo>? attachments,
            string? bodyHtmlDigest)
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

            // Attachment SET, order-independent like the recipients above: Outlook can
            // report the collection in a different order after a reopen, so the hash must
            // depend on WHICH files are attached, not on how they enumerate. Name AND
            // size, so replacing a file with a same-named different one is a change.
            IEnumerable<string> attachmentLines = (attachments ?? Array.Empty<ComAttachmentInfo>())
                .Select(a => (a.FileName ?? string.Empty).Trim().ToLowerInvariant()
                    + "|" + (a.SizeBytes?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "?"))
                .OrderBy(line => line, StringComparer.Ordinal);
            foreach (string line in attachmentLines)
            {
                canonical.Append("att=").Append(line).Append('\n');
            }

            canonical.Append("htmldigest=").Append((bodyHtmlDigest ?? string.Empty).Trim().ToLowerInvariant()).Append('\n');
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
        /// Digest of the stored HTML body, folded into <see cref="Compute"/> so a
        /// markup-only change (which leaves <c>MailItem.Body</c> byte-identical) still
        /// invalidates a pending confirm token. Null/empty HTML yields null, which
        /// canonicalizes to an empty contribution.
        /// </summary>
        public static string? DigestHtml(string? html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return null;
            }

            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(html!));
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
