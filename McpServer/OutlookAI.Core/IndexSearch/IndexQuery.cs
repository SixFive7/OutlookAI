using System;
using System.Collections.Generic;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// Where free-text terms are matched (v3.MD sections 4/5, Phase-1 probes, D40/SF-6).
    /// <para>
    /// SF-6 (measured 2026-07-26): a bare <c>CONTAINS('term')</c> is NOT an all-properties
    /// match - Windows Search documents the unqualified CONTAINS predicate as searching
    /// <c>System.Search.Contents</c> alone ("the body of the document"), and the contents
    /// stream carries no subject text. Mail whose term appears only in the subject was
    /// therefore invisible to the term predicate (~3.4% of items store-wide). The bare
    /// shape is gone; every scope below names its column(s) explicitly.
    /// </para>
    /// </summary>
    public enum SearchIn
    {
        /// <summary>
        /// <c>(CONTAINS(System.Subject, ...) OR CONTAINS(System.Search.Contents, ...))</c>:
        /// subject OR body/attachment content. THE DEFAULT (user order 2026-07-26) and the
        /// completeness oracle's parity shape, where ground truth is computed from
        /// Subject/Body via COM.
        /// </summary>
        SubjectAndBody = 0,

        /// <summary>
        /// <c>CONTAINS(System.Subject, ...)</c>: subject line only - useful when a term is
        /// noisy in body text (quoted threads, footers, signatures).
        /// </summary>
        SubjectOnly = 1,

        /// <summary>
        /// <c>CONTAINS(System.Search.Contents, ...)</c>: body + attachment content only -
        /// useful when a term is noisy in subjects (alert prefixes, ticket tags).
        /// </summary>
        BodyOnly = 2,
    }

    /// <summary>
    /// Wire names for <see cref="SearchIn"/> and the host-neutral parser behind the
    /// <c>search_in</c> tool argument (v3.MD section 0.5.2: Core carries no MCP types).
    /// </summary>
    public static class SearchInValues
    {
        /// <summary>Wire name of <see cref="SearchIn.SubjectAndBody"/> - the default.</summary>
        public const string SubjectAndBodyName = "subject_and_body";

        /// <summary>Wire name of <see cref="SearchIn.SubjectOnly"/>.</summary>
        public const string SubjectName = "subject";

        /// <summary>Wire name of <see cref="SearchIn.BodyOnly"/>.</summary>
        public const string BodyName = "body";

        /// <summary>The scope used when the argument is omitted.</summary>
        public const SearchIn Default = SearchIn.SubjectAndBody;

        /// <summary>The three accepted wire values, in schema order.</summary>
        public static readonly IReadOnlyList<string> WireNames = new[]
        {
            SubjectAndBodyName, SubjectName, BodyName,
        };

        /// <summary>
        /// Parses a <c>search_in</c> argument. Null/blank yields <see cref="Default"/>;
        /// matching is case- and whitespace-insensitive and tolerates the two obvious
        /// near-misses (<c>subject_only</c>/<c>body_only</c>) rather than failing a whole
        /// tool call over a naming guess. Anything else throws with the valid values.
        /// </summary>
        public static SearchIn Parse(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return Default;
            }

            switch (value!.Trim().ToLowerInvariant())
            {
                case SubjectAndBodyName:
                    return SearchIn.SubjectAndBody;
                case SubjectName:
                case "subject_only":
                    return SearchIn.SubjectOnly;
                case BodyName:
                case "body_only":
                    return SearchIn.BodyOnly;
                default:
                    throw new ArgumentException(
                        "search_in must be one of: " + string.Join(", ", WireNames)
                        + " (default " + SubjectAndBodyName + " = the term must appear in the subject or in the body/attachment content).",
                        nameof(value));
            }
        }

        /// <summary>Wire name of a scope (inverse of <see cref="Parse"/>).</summary>
        public static string ToWireName(SearchIn scope)
        {
            switch (scope)
            {
                case SearchIn.SubjectAndBody:
                    return SubjectAndBodyName;
                case SearchIn.SubjectOnly:
                    return SubjectName;
                case SearchIn.BodyOnly:
                    return BodyName;
                default:
                    throw new ArgumentException("Unknown SearchIn value.", nameof(scope));
            }
        }
    }

    /// <summary>
    /// Which ROW SHAPES an index query covers - message-level rows, attachment-content
    /// rows, or both. A query is never emitted without one (v3.MD section 12).
    /// <para>
    /// THE NAMES CHANGED WITH THE RULE (gap B3, maintainer decision 2026-08-18). They used
    /// to be <c>EmailAndDocuments</c> / <c>EmailOnly</c> / <c>DocumentsOnly</c> /
    /// <c>MessagesAnyClass</c>, from a time when a message row had to be kind <c>email</c>
    /// to be admitted at all. Item class no longer narrows any tier
    /// (<see cref="OutlookAI.Core.Mapi.MailItemAdmission"/>), so those names would now
    /// describe a filter that is gone - which is the same defect as <c>FromAddressContains</c>
    /// matching names: a field name the next reader believes instead of the code.
    /// </para>
    /// </summary>
    public enum KindFilter
    {
        /// <summary>
        /// Message-level rows of EVERY item class plus attachment-content rows of every
        /// kind - the widest shape there is, and what <c>search</c> asks for by default.
        /// Emits no Kind predicate under a mapi SCOPE (the namespace is the guard) and the
        /// enumerated kind list without one.
        /// </summary>
        MessagesAndAttachments = 0,

        /// <summary>
        /// Message-level rows of every item class, and no attachment-content rows: what
        /// <c>search</c> asks for when <c>include_attachment_hits</c> is false, and what
        /// <c>thread</c> always asks for (gap C2 - a meeting request carries the
        /// surrounding mail's ConversationID and indexes as <c>calendar</c>, so a
        /// kind-narrowed thread dropped real members of a conversation the tool promises
        /// whole). Admission is <see cref="IndexRowFilter.Keep"/>'s "not an attachment row"
        /// alone.
        /// </summary>
        MessagesOnly = 1,

        /// <summary>Attachment-content entries only, any kind (probe R1 shape / attachment_hits_only).</summary>
        AttachmentsOnly = 2,

        /// <summary>
        /// <c>System.Kind='email'</c>: the one narrow shape left, and no search uses it.
        /// It exists for store-scope DISCOVERY
        /// (<see cref="IndexSearchService.TryDiscoverStoreScopeByAddress"/>), which needs a
        /// row that is certainly a mail message in order to read a store prefix off it, and
        /// for the completeness oracle's parity mode. Narrowing a user-facing search with
        /// it is what gap B3 was.
        /// </summary>
        MailKindOnly = 3,
    }

    /// <summary>Result ordering for an index query.</summary>
    public enum IndexOrder
    {
        /// <summary>ORDER BY System.Message.DateReceived DESC - the default.</summary>
        DateReceivedDescending = 0,

        /// <summary>ORDER BY System.Size DESC (Phase-2 truncation tests: find big mails).</summary>
        SizeDescending = 1,
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

        /// <summary>
        /// Non-recursive folder narrowing: the exact <c>System.ItemFolderPathDisplay</c>
        /// values a row may carry, ORed. Combined with <see cref="Scope"/> this selects
        /// specific folders WITHOUT their subfolders, including their attachment-content
        /// rows (those inherit the parent message's folder display path).
        /// <para>
        /// Equality on this column is index-backed (<c>isColumn</c> in the property
        /// schema; measured 6-49 ms) - it is NOT the section-12 <c>LIKE</c>-on-
        /// <c>System.ItemPathDisplay</c> property scan. The shallow <c>DIRECTORY=</c>
        /// predicate is deliberately NOT used: it returns zero <c>System.Kind='document'</c>
        /// rows, dropping up to 41% of a folder's hits (v3.MD section 12).
        /// </para>
        /// <para>
        /// Null = no folder narrowing (a bare <see cref="Scope"/> is recursive). Values
        /// are derived from the scope URL by
        /// <see cref="OutlookAI.Core.Mapi.MapiItemUrl.TryBuildFolderPathDisplay"/> -
        /// never from a store display name.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? FolderPathsAnyOf { get; set; }

        /// <summary>Free-text terms, ANDed. Each may end in '*' for prefix matching.</summary>
        public IReadOnlyList<string>? Terms { get; set; }

        /// <summary>Where <see cref="Terms"/> are matched. Default: subject OR body/attachment content (D40/SF-6).</summary>
        public SearchIn SearchIn { get; set; } = SearchInValues.Default;

        /// <summary>Which row shapes the query covers. Default: every message row plus every attachment row.</summary>
        public KindFilter Kinds { get; set; } = KindFilter.MessagesAndAttachments;

        /// <summary>
        /// Sender filter, matched with CONTAINS over the sender ADDRESS and the sender
        /// display NAME - the same "address or name fragment" the tool description promises
        /// and the same thing the freshness sweep and the exhaustive scan have always
        /// matched. Phase-1 probe result: equality and LIKE on either column are
        /// multi-second property scans; per-column CONTAINS is the only index-backed shape.
        /// <para>
        /// It was <c>FromAddressContains</c> and matched the address alone, while
        /// <c>System.Message.FromName</c> was SELECTed and never used in a predicate. So a
        /// caller filtering by a person's display name got NOTHING from the index tier and
        /// whatever the freshness sweep window happened to catch from the other two - an
        /// answer built from minutes of mail, reported as complete. MEASURED on this
        /// machine's index (read-only, 2026-08-18): of 419 distinct senders in a 3 000-row
        /// sample, 218 (52%) have a display-name token that appears nowhere in their
        /// address, so for half the correspondents in the mailbox the index tier could not
        /// be reached by name at all.
        /// </para>
        /// <para>
        /// The name changed with the behaviour on purpose: a field called
        /// <c>FromAddressContains</c> that also matches names is the same defect one level
        /// down, where the next reader believes the field name instead of the SQL.
        /// </para>
        /// </summary>
        public string? SenderContains { get; set; }

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
        /// Exact-match filter on System.Message.ConversationID (thread tool). NOTE: '='
        /// on a non-CONTAINS column is a property scan - always combine with
        /// <see cref="Scope"/> when the store is known, and keep Top small.
        /// </summary>
        public string? ConversationIdEquals { get; set; }

        /// <summary>Result ordering; DateReceived DESC unless stated otherwise.</summary>
        public IndexOrder OrderBy { get; set; } = IndexOrder.DateReceivedDescending;

        /// <summary>
        /// Maximum rows (SELECT TOP n). Compact-payload discipline: default 25, hard cap
        /// 5000 (the cap exists for the completeness oracle and store discovery, not for
        /// agent-facing result lists).
        /// </summary>
        public int Top { get; set; } = 25;
    }
}
