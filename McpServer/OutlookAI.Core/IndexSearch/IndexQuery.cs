using System;
using System.Collections.Generic;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>Where free-text terms are matched (v3.MD sections 4/5 + Phase-1 probes).</summary>
    public enum TermScope
    {
        /// <summary>
        /// Bare <c>CONTAINS('...')</c>: matches every full-text-indexed property - body
        /// content, subject, sender/recipient names and addresses, attachment names. The
        /// product default: broadest recall per iteration.
        /// </summary>
        AllProperties = 0,

        /// <summary>
        /// <c>(CONTAINS(System.Subject, ...) OR CONTAINS(System.Search.Contents, ...))</c>:
        /// subject + body content only. Exact-parity mode used by the completeness oracle,
        /// where ground truth is computed from Subject/Body via COM.
        /// </summary>
        SubjectAndBody = 1,
    }

    /// <summary>System.Kind filter shape - a query is never emitted without one (v3.MD section 12).</summary>
    public enum KindFilter
    {
        /// <summary>
        /// <c>(System.Kind='email' OR System.Kind='document')</c>: messages plus indexed
        /// attachment-content entries. The default - email-only queries miss
        /// attachment-content hits (v3.MD section 12).
        /// </summary>
        EmailAndDocuments = 0,

        /// <summary><c>System.Kind='email'</c>: messages only (completeness-oracle parity mode).</summary>
        EmailOnly = 1,

        /// <summary><c>System.Kind='document'</c>: attachment-content entries only (probe R1 shape).</summary>
        DocumentsOnly = 2,
    }

    /// <summary>
    /// Parameters for one SystemIndex search. Translated to Windows Search SQL by
    /// <see cref="WsSqlBuilder"/>, which enforces the v3.MD section-12 anti-pattern guards.
    /// </summary>
    public sealed class IndexQuery
    {
        /// <summary>
        /// SCOPE predicate value restricting the search to a store or folder subtree, e.g.
        /// <c>mapi16://{SID}/account($hash)</c> or <c>.../account($hash)/0/Inbox</c>.
        /// Must use a mapi scheme when set - path filtering via
        /// <c>System.ItemPathDisplay LIKE</c> is a 9-10 s property scan (v3.MD section 12).
        /// Null searches the whole index (all stores; anything else indexed is excluded by
        /// the kind filter).
        /// </summary>
        public string? Scope { get; set; }

        /// <summary>Free-text terms, ANDed. Each may end in '*' for prefix matching.</summary>
        public IReadOnlyList<string>? Terms { get; set; }

        /// <summary>Where <see cref="Terms"/> are matched. Default: all indexed properties.</summary>
        public TermScope TermScope { get; set; } = TermScope.AllProperties;

        /// <summary>Which System.Kind values the query covers. Default: email plus documents.</summary>
        public KindFilter Kinds { get; set; } = KindFilter.EmailAndDocuments;

        /// <summary>
        /// Sender filter, matched with <c>CONTAINS(System.Message.FromAddress, ...)</c>.
        /// Phase-1 probe result: equality and LIKE on FromAddress are multi-second property
        /// scans; per-column CONTAINS is the only index-backed shape (~60 ms).
        /// </summary>
        public string? FromAddressContains { get; set; }

        /// <summary>
        /// Recipient filter, matched with CONTAINS over To and Cc address columns.
        /// </summary>
        public string? RecipientContains { get; set; }

        /// <summary>Lower bound (inclusive) on System.Message.DateReceived, UTC.</summary>
        public DateTime? ReceivedOnOrAfterUtc { get; set; }

        /// <summary>Upper bound (exclusive) on System.Message.DateReceived, UTC.</summary>
        public DateTime? ReceivedBeforeUtc { get; set; }

        /// <summary>Filter on read state (System.IsRead).</summary>
        public bool? IsRead { get; set; }

        /// <summary>Filter on attachment presence (System.Message.HasAttachments).</summary>
        public bool? HasAttachments { get; set; }

        /// <summary>
        /// Maximum rows (SELECT TOP n). Compact-payload discipline: default 25, hard cap
        /// 5000 (the cap exists for the completeness oracle and store discovery, not for
        /// agent-facing result lists).
        /// </summary>
        public int Top { get; set; } = 25;
    }
}
