namespace OutlookAI.RemediationTools;

/// <summary>
/// The facts about one store that decide whether a corpus may be written into it. Every
/// field is nullable because every one of them is a late-bound COM read that can fail,
/// and "could not read it" must never be confused with "it is fine".
/// </summary>
/// <param name="DisplayName">Store.DisplayName.</param>
/// <param name="IsDataFileStore">Store.IsDataFileStore - true only for a store not tied to an account.</param>
/// <param name="ExchangeStoreType">Raw OlExchangeStoreType (0 primary, 1 delegate, 2 public folders, 3 not Exchange).</param>
/// <param name="FilePath">Store.FilePath - the backing file, when there is one.</param>
public sealed record CorpusStoreFacts(
    string? DisplayName,
    bool? IsDataFileStore,
    int? ExchangeStoreType,
    string? FilePath);

/// <summary>Why a store was refused as a corpus target. <see cref="None"/> is the only value that permits a write.</summary>
public enum CorpusStoreRefusal
{
    /// <summary>Permitted.</summary>
    None = 0,

    /// <summary>The caller passed no allowlist, or an empty one.</summary>
    NoAllowlist = 1,

    /// <summary>The store has no readable display name, so it cannot be matched against anything.</summary>
    NoStoreName = 2,

    /// <summary>The store's name is not on the allowlist the caller passed in.</summary>
    NotOnAllowlist = 3,

    /// <summary>Store.IsDataFileStore is false - the store belongs to a mail account.</summary>
    NotADataFileStore = 4,

    /// <summary>OlExchangeStoreType says this is an Exchange store of some kind.</summary>
    ExchangeStore = 5,

    /// <summary>There is no backing file, or it is not a .pst.</summary>
    NotAPstFile = 6,

    /// <summary>One or more of the facts could not be read, so nothing about the store is proven.</summary>
    Unprovable = 7,
}

/// <summary>
/// Every refusal the corpus generator makes, as pure functions over facts. Nothing here
/// touches COM, so the T1 tier pins the answers - which matters more here than anywhere
/// else in this project, because the thing being decided is whether it is safe to write
/// tens of thousands of items into a mailbox.
/// <para>
/// The rules follow CLAUDE.md's mailbox-safety section. Two of them are load-bearing and
/// neither has an exception:
/// </para>
/// <list type="number">
/// <item>
/// A write happens only into a store the CALLER named on an allowlist AND that four
/// independent COM facts agree is a local .pst data file. The allowlist alone is not
/// enough: display names are user-editable and a profile can carry two stores with the
/// same one, so a typo or a renamed mailbox could otherwise put a corpus into live mail.
/// </item>
/// <item>
/// A delete happens only when the item's EntryID is in the run's manifest AND the item's
/// re-read subject carries the tags ordinally. Both, always. Rule 2 of the mailbox-safety
/// section exists because a shell-side subject pattern - where the tag's own brackets
/// became a character class - matched nearly every subject and destroyed real mail.
/// </item>
/// </list>
/// </summary>
public static class CorpusSafety
{
    /// <summary>OlExchangeStoreType value meaning "not an Exchange store" - the only acceptable one.</summary>
    public const int NotExchangeStoreType = 3;

    /// <summary>The extension a corpus target must have.</summary>
    public const string PstExtension = ".pst";

    /// <summary>
    /// Decides whether <paramref name="facts"/> describe a store a corpus may be written
    /// into. FAIL-CLOSED: an unreadable fact refuses, because a store nothing is known
    /// about is exactly the store that must not be written to.
    /// </summary>
    public static CorpusStoreRefusal EvaluateStore(CorpusStoreFacts facts, IReadOnlyCollection<string>? allowlist)
    {
        ArgumentNullException.ThrowIfNull(facts);
        if (allowlist == null || allowlist.Count == 0)
        {
            return CorpusStoreRefusal.NoAllowlist;
        }

        if (string.IsNullOrWhiteSpace(facts.DisplayName))
        {
            return CorpusStoreRefusal.NoStoreName;
        }

        bool named = false;
        foreach (string allowed in allowlist)
        {
            // Ordinal, case-insensitive, whole name - never a prefix and never a pattern.
            if (!string.IsNullOrWhiteSpace(allowed)
                && string.Equals(allowed.Trim(), facts.DisplayName!.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                named = true;
                break;
            }
        }

        if (!named)
        {
            return CorpusStoreRefusal.NotOnAllowlist;
        }

        if (facts.IsDataFileStore == null || facts.ExchangeStoreType == null || facts.FilePath == null)
        {
            return CorpusStoreRefusal.Unprovable;
        }

        if (facts.IsDataFileStore != true)
        {
            return CorpusStoreRefusal.NotADataFileStore;
        }

        if (facts.ExchangeStoreType != NotExchangeStoreType)
        {
            return CorpusStoreRefusal.ExchangeStore;
        }

        if (!facts.FilePath!.EndsWith(PstExtension, StringComparison.OrdinalIgnoreCase))
        {
            return CorpusStoreRefusal.NotAPstFile;
        }

        return CorpusStoreRefusal.None;
    }

    /// <summary>
    /// The refusal message. Names only the offending store, never any other store's name,
    /// subject or content - the same discipline the live tier's guard keeps.
    /// </summary>
    public static string Explain(CorpusStoreRefusal refusal, CorpusStoreFacts facts)
    {
        ArgumentNullException.ThrowIfNull(facts);
        string target = string.IsNullOrWhiteSpace(facts.DisplayName) ? "(unnamed store)" : facts.DisplayName!;
        string why = refusal switch
        {
            CorpusStoreRefusal.None => "permitted",
            CorpusStoreRefusal.NoAllowlist => "no --allow-store was given; the corpus tool never picks a target itself",
            CorpusStoreRefusal.NoStoreName => "the store has no readable display name, so nothing can vouch for it",
            CorpusStoreRefusal.NotOnAllowlist => "it is not on the --allow-store list this run was given",
            CorpusStoreRefusal.NotADataFileStore => "Store.IsDataFileStore is false - this store belongs to a mail account",
            CorpusStoreRefusal.ExchangeStore => "OlExchangeStoreType says this is an Exchange store",
            CorpusStoreRefusal.NotAPstFile => "its backing file is not a .pst",
            CorpusStoreRefusal.Unprovable => "one of IsDataFileStore / ExchangeStoreType / FilePath could not be read, "
                + "so nothing about this store is proven",
            _ => "unrecognised refusal",
        };

        return refusal == CorpusStoreRefusal.None
            ? $"Store '{target}' accepted as a corpus target."
            : $"REFUSING to build a corpus in store '{target}': {why}. A corpus may only be written into a LOCAL "
                + ".pst named explicitly by the caller. See the mailbox-safety rules in CLAUDE.md; widen the "
                + "--allow-store list, never the guard.";
    }

    /// <summary>
    /// The ONLY sanctioned delete predicate. True when both independent conditions hold:
    /// the EntryID is one this run's manifest recorded creating, AND the subject just
    /// re-read from the item carries the mailbox-safety tag and this corpus's tag,
    /// matched ordinally.
    /// <para>
    /// Both are required because each covers the other's failure. An EntryID alone would
    /// delete whatever now lives at a recycled or mistyped id; a tag alone is a content
    /// match on user-editable text, which is the shape of the deletion that destroyed
    /// real mail. Neither is a pattern: the tag check is
    /// <see cref="string.Contains(string, StringComparison)"/> with
    /// <see cref="StringComparison.Ordinal"/>, so every character - including the tag's
    /// brackets - is a literal.
    /// </para>
    /// </summary>
    public static bool MayDelete(string? entryId, string? subject, ISet<string> entryIdAllowlist, string corpusId)
    {
        ArgumentNullException.ThrowIfNull(entryIdAllowlist);
        if (string.IsNullOrWhiteSpace(entryId) || subject == null)
        {
            return false;
        }

        return entryIdAllowlist.Contains(entryId) && CorpusPlan.TryParseOrdinal(subject, corpusId, out _);
    }

    /// <summary>
    /// The only sanctioned way to build the allowlist <see cref="MayDelete"/> consults.
    /// It exists so the comparer is not left to the call site: EntryIDs are hex and are
    /// sometimes reported in a different case than they were recorded in, and an allowlist
    /// built with the default comparer would silently fail to match its own entries -
    /// which fails safe here (nothing is deleted) but leaves items behind that the operator
    /// was told had been removed.
    /// </summary>
    public static HashSet<string> BuildEntryIdAllowlist(IEnumerable<string> entryIds)
    {
        ArgumentNullException.ThrowIfNull(entryIds);
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string id in entryIds)
        {
            if (!string.IsNullOrWhiteSpace(id))
            {
                set.Add(id);
            }
        }

        return set;
    }
}
