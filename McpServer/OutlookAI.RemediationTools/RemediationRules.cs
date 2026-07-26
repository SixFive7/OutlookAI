namespace OutlookAI.RemediationTools;

/// <summary>
/// Pure decision logic for the 2026-07-25 incident remediation (v3.MD SOAK FIX LOG
/// entry 7): tagged-artifact purge matching, telefonie refile origin classification
/// with the deletion-log cross-check, and the jori duplicate-delete verification.
/// Deliberately COM-free so the T1 tier pins every rule (the standing post-incident
/// rule: deletions only through tested C# helper code with ORDINAL tag matching -
/// never shell-side patterns, whose wildcard classes caused the incident).
/// </summary>
public static class RemediationRules
{
    /// <summary>The S3 subject tag every test artifact carries (ordinal match only).</summary>
    public const string SubjectTag = "[OutlookAI-McpTest]";

    /// <summary>
    /// Bracket-free DASL LIKE prefilter fragment (the suite's proven pattern for fast
    /// GetTable counts). A superset of tag matches by construction - every subject
    /// containing the full tag contains this fragment - so LIKE count 0 proves tagged
    /// count 0. The authoritative per-item decision is always
    /// <see cref="IsTagged(string?)"/> on the re-read subject.
    /// </summary>
    public const string DaslCountFragment = "OutlookAI-McpTest";

    /// <summary>Ordinal full-tag check - the ONLY sanctioned delete predicate for the purge.</summary>
    public static bool IsTagged(string? subject)
        => subject != null && subject.Contains(SubjectTag, StringComparison.Ordinal);

    /// <summary>Classification of a telefonie Deleted Items item for the refile (step 1).</summary>
    public enum TelefonieOrigin
    {
        /// <summary>Hub-sent mail - belongs in Sent Items.</summary>
        SentOrigin,

        /// <summary>Received mail - belongs in the Inbox.</summary>
        InboxOrigin,
    }

    /// <summary>
    /// Classifies one untagged telefonie Deleted Items item by two INDEPENDENT
    /// signals that must agree: the sender SMTP being the hub itself (sent-origin)
    /// and the PR_RECEIVED_BY presence (received-origin). Disagreement returns null -
    /// the caller aborts instead of guessing (nothing is moved on ambiguity).
    /// </summary>
    public static TelefonieOrigin? ClassifyOrigin(string? senderSmtp, bool receivedByPresent, string hubSmtpAddress)
    {
        bool senderIsHub = senderSmtp != null
            && string.Equals(senderSmtp.Trim(), hubSmtpAddress.Trim(), StringComparison.OrdinalIgnoreCase);
        if (senderIsHub && !receivedByPresent)
        {
            return TelefonieOrigin.SentOrigin;
        }

        if (!senderIsHub && receivedByPresent)
        {
            return TelefonieOrigin.InboxOrigin;
        }

        return null; // signals disagree - ambiguous, never guessed
    }

    /// <summary>One line of the preserved incident deletion log.</summary>
    public readonly record struct DeletionLogEntry(string Store, int FolderId, string Prefix);

    /// <summary>
    /// Parses the preserved incident-deletion-log.txt lines
    /// (<c>delete: store=... folder=N markerPrefix=P</c>). Unparseable lines throw:
    /// the log is the authoritative ground truth and must be read exactly.
    /// </summary>
    public static IReadOnlyList<DeletionLogEntry> ParseDeletionLog(IEnumerable<string> lines)
    {
        List<DeletionLogEntry> entries = new();
        int lineNo = 0;
        foreach (string raw in lines)
        {
            lineNo++;
            string line = raw.Trim();
            if (line.Length == 0)
            {
                continue;
            }

            string? store = null;
            int? folderId = null;
            string? prefix = null;
            foreach (string token in line.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                if (token.StartsWith("store=", StringComparison.Ordinal))
                {
                    store = token.Substring("store=".Length);
                }
                else if (token.StartsWith("folder=", StringComparison.Ordinal))
                {
                    folderId = int.Parse(token.Substring("folder=".Length), System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (token.StartsWith("markerPrefix=", StringComparison.Ordinal))
                {
                    prefix = token.Substring("markerPrefix=".Length);
                }
            }

            if (store == null || folderId == null || prefix == null)
            {
                throw new FormatException($"Deletion log line {lineNo} is not in the expected format.");
            }

            entries.Add(new DeletionLogEntry(store, folderId.Value, prefix));
        }

        return entries;
    }

    /// <summary>
    /// The expected subject-prefix multiset for one store+folder of the deletion log
    /// (e.g. telefonie folder 5 = the 21 Sent-origin items; folder 6 = the 2
    /// Inbox-origin items).
    /// </summary>
    public static List<string> ExpectedPrefixes(IReadOnlyList<DeletionLogEntry> log, string store, int folderId)
        => log.Where(e => string.Equals(e.Store, store, StringComparison.OrdinalIgnoreCase) && e.FolderId == folderId)
            .Select(e => e.Prefix)
            .ToList();

    /// <summary>
    /// Consumes one expected log prefix that matches <paramref name="subject"/>
    /// (ordinal StartsWith after leading-whitespace trim - the log recorded the first
    /// subject token truncated to 8 chars, so prefix-of-subject is the robust
    /// direction). Returns the consumed prefix, or null when nothing in the remaining
    /// multiset matches. The telefonie prefix set (RE:, FW:, Actie, Update, Telefoni,
    /// Trunk, test) is mutually non-prefixing, so greedy consumption is exact there.
    /// </summary>
    public static string? TryConsumePrefixMatch(List<string> remainingPrefixes, string? subject)
    {
        if (subject == null)
        {
            return null;
        }

        string trimmed = subject.TrimStart();
        for (int i = 0; i < remainingPrefixes.Count; i++)
        {
            if (trimmed.StartsWith(remainingPrefixes[i], StringComparison.Ordinal))
            {
                string consumed = remainingPrefixes[i];
                remainingPrefixes.RemoveAt(i);
                return consumed;
            }
        }

        return null;
    }

    /// <summary>Verdict for one jori Deleted Items item in the duplicate delete (step 3).</summary>
    public enum DedupeDecision
    {
        /// <summary>Verified duplicate: untagged, has a Message-ID, and that Message-ID exists in the Inbox right now.</summary>
        Delete,

        /// <summary>Tagged artifact - the purge owns it, never the dedupe.</summary>
        SkipTagged,

        /// <summary>PR_INTERNET_MESSAGE_ID empty/absent - verification impossible, kept.</summary>
        SkipEmptyMessageId,

        /// <summary>No item with the same Message-ID currently in the Inbox - kept.</summary>
        SkipNoInboxTwin,
    }

    /// <summary>
    /// The delete-time verification of step 3: ONLY an untagged item whose non-empty
    /// PR_INTERNET_MESSAGE_ID has a twin CURRENTLY in the store's Inbox may be
    /// deleted; every other outcome is a reported skip. Message-IDs compare ordinal
    /// after trimming (exact same id, no case folding).
    /// </summary>
    public static DedupeDecision DecideDuplicateDelete(
        string? subject,
        string? internetMessageId,
        IReadOnlySet<string> inboxMessageIds)
    {
        if (IsTagged(subject))
        {
            return DedupeDecision.SkipTagged;
        }

        string? id = internetMessageId?.Trim();
        if (string.IsNullOrEmpty(id))
        {
            return DedupeDecision.SkipEmptyMessageId;
        }

        return inboxMessageIds.Contains(id) ? DedupeDecision.Delete : DedupeDecision.SkipNoInboxTwin;
    }

    /// <summary>Normalizes a Message-ID for set membership (trim only - ordinal identity).</summary>
    public static string? NormalizeMessageId(string? internetMessageId)
    {
        string? id = internetMessageId?.Trim();
        return string.IsNullOrEmpty(id) ? null : id;
    }
}
