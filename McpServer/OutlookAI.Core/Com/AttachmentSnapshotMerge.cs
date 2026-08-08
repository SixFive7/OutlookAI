using System.Collections.Generic;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Pure rules for choosing between two attachment snapshots of the SAME item, taken
    /// moments apart through different COM references.
    /// <para>
    /// Soak fix 21, and the rules exist because of a live-reported defect: the item
    /// reference a draft flow holds while composing - the one whose hidden Inspector was
    /// just closed - reports <c>Attachment.Size</c> as ZERO for an attachment Outlook
    /// materialized during that composition (a signature's inline logo), and in the
    /// HTMLBody-fallback shape it reports <c>Attachments.Count</c> as zero as well. The
    /// bytes are on disk and correct the entire time. Re-opening the saved item by EntryID
    /// answers truthfully, so the flows take BOTH snapshots and keep the better one -
    /// which is what these rules decide. "Better" is deliberately monotone: a re-read can
    /// only ever improve a result, never replace known bytes with unknown ones.
    /// </para>
    /// </summary>
    public static class AttachmentSnapshotMerge
    {
        /// <summary>
        /// True when at least one attachment has no positive size - the shape that means
        /// "this snapshot may be premature", not "this file is empty" (a genuinely empty
        /// attachment is indistinguishable here, which is why the caller's retry is
        /// bounded to one attempt).
        /// </summary>
        public static bool HasUnsizedAttachment(IReadOnlyList<ComAttachmentInfo> attachments)
        {
            if (attachments == null)
            {
                return false;
            }

            for (int i = 0; i < attachments.Count; i++)
            {
                long? size = attachments[i].SizeBytes;
                if (!(size.HasValue && size.Value > 0))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>Total of the sizes that are actually known (positive) in a snapshot.</summary>
        public static long KnownBytes(IReadOnlyList<ComAttachmentInfo> attachments)
        {
            if (attachments == null)
            {
                return 0;
            }

            long total = 0;
            for (int i = 0; i < attachments.Count; i++)
            {
                long? size = attachments[i].SizeBytes;
                if (size.HasValue && size.Value > 0)
                {
                    total += size.Value;
                }
            }

            return total;
        }

        /// <summary>
        /// True when <paramref name="candidate"/> should replace <paramref name="current"/>:
        /// it sees MORE attachments, or the same number with MORE known bytes. Equal
        /// snapshots never swap, and a snapshot that lost an attachment never wins.
        /// </summary>
        public static bool IsBetter(
            IReadOnlyList<ComAttachmentInfo> candidate,
            IReadOnlyList<ComAttachmentInfo> current)
        {
            if (candidate == null)
            {
                return false;
            }

            if (current == null)
            {
                return true;
            }

            if (candidate.Count != current.Count)
            {
                return candidate.Count > current.Count;
            }

            return KnownBytes(candidate) > KnownBytes(current);
        }
    }
}
