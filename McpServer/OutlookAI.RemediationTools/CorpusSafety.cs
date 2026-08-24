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

/// <summary>
/// What the PROFILE looks like from the point of view of "could a build queue mail for
/// delivery". Separate from <see cref="CorpusStoreFacts"/> because it is a fact about the
/// profile, not about the store, and because the store half proved insufficient on its own:
/// a local .pst can be an account's delivery store.
/// </summary>
/// <param name="AccountCount">Session.Accounts.Count, or null when it could not be read.</param>
/// <param name="AccountsDeliveringToTarget">
/// Accounts whose <c>Account.DeliveryStore</c> is the target store, compared by StoreID
/// rather than by display name - names are user-editable and two stores may share one.
/// </param>
/// <param name="AccountsWithUnreadableDeliveryStore">
/// Accounts whose delivery store could not be read at all. Any of these makes the profile
/// unprovable, because an account that cannot be examined is an account that might deliver
/// into the target.
/// </param>
public sealed record CorpusProfileFacts(
    int? AccountCount,
    int AccountsDeliveringToTarget,
    int AccountsWithUnreadableDeliveryStore);

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

    /// <summary>
    /// An account in this profile delivers into the target store. Building here would create
    /// mail inside a store a transport provider services, and an Outbox in such a store is
    /// not inert.
    /// </summary>
    TargetIsAccountDeliveryStore = 8,

    /// <summary>
    /// The profile has at least one account, so something in it can send. See
    /// <see cref="CorpusSafety.EvaluateProfile"/> for why this is refused even when no
    /// account delivers into the target.
    /// </summary>
    ProfileCanSend = 9,

    /// <summary>The account list, or one account's delivery store, could not be read.</summary>
    ProfileUnprovable = 10,
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
/// <item>
/// A write happens only into a profile that has NO mail accounts at all. Added 2026-08-19
/// after the first real build: a 40 000 item run put 5 532 items into the target store's
/// Outbox. On that VM they were inert because the profile had no account and nothing could
/// send - but the same build on a profile with an account would have queued 5 532 messages
/// for delivery. The store half of this guard did not catch it, and could not: "local .pst"
/// and "an account's delivery store" are not mutually exclusive.
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
    /// Whether the PROFILE is one a corpus may be built in. The rule is: NO MAIL ACCOUNTS AT
    /// ALL.
    /// <para>
    /// That is stricter than "no account delivers into the target store", and the strictness
    /// is forced rather than chosen. The requirement is "no account able to SEND FROM the
    /// target store", and the object model cannot express it: a draft's location does not
    /// constrain which account sends it, because <c>SendUsingAccount</c> is set per item and
    /// any account may send a message that lives anywhere. So "no account can send from this
    /// store" is, through the OM, indistinguishable from "no account can send" - which is
    /// <c>Accounts.Count == 0</c>. Anything weaker would be a guard that reads as a proof and
    /// is not one.
    /// </para>
    /// <para>
    /// It costs nothing operationally: a measurement VM has no mail account, which is the
    /// entire reason its Outbox is inert. There is deliberately NO override flag - the one
    /// safety verdict this tool has already had overridden was overridden on reasoning that
    /// turned out to be wrong, and this is the verdict where being wrong queues real mail.
    /// </para>
    /// <para>
    /// FAIL-CLOSED, like the store facts: an account list that cannot be read, or one account
    /// whose delivery store cannot be read, refuses. An account nothing can examine is an
    /// account that might deliver into the target.
    /// </para>
    /// </summary>
    public static CorpusStoreRefusal EvaluateProfile(CorpusProfileFacts facts)
    {
        ArgumentNullException.ThrowIfNull(facts);
        if (facts.AccountCount == null || facts.AccountsWithUnreadableDeliveryStore > 0)
        {
            return CorpusStoreRefusal.ProfileUnprovable;
        }

        // Named separately from the general case even though both refuse: this is the one
        // that actually happened, and an operator reading the message deserves to be told
        // which of the two they are looking at.
        if (facts.AccountsDeliveringToTarget > 0)
        {
            return CorpusStoreRefusal.TargetIsAccountDeliveryStore;
        }

        return facts.AccountCount > 0 ? CorpusStoreRefusal.ProfileCanSend : CorpusStoreRefusal.None;
    }

    /// <summary>
    /// The whole gate: the store must pass <see cref="EvaluateStore"/> AND the profile must
    /// pass <see cref="EvaluateProfile"/>. The store is judged first so its message is the
    /// one an operator sees when the target is simply wrong.
    /// </summary>
    public static CorpusStoreRefusal Evaluate(
        CorpusStoreFacts storeFacts, CorpusProfileFacts profileFacts, IReadOnlyCollection<string>? allowlist)
    {
        CorpusStoreRefusal storeVerdict = EvaluateStore(storeFacts, allowlist);
        return storeVerdict != CorpusStoreRefusal.None ? storeVerdict : EvaluateProfile(profileFacts);
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
            CorpusStoreRefusal.TargetIsAccountDeliveryStore => "an account in this profile DELIVERS into this store, "
                + "so a transport provider services it. A build here would put items into a live Outbox",
            CorpusStoreRefusal.ProfileCanSend => "this profile has at least one mail account, so something in it can "
                + "send. A corpus is built only in a profile with NO accounts - the object model cannot prove the "
                + "narrower 'no account may send from this store', because any account may send an item that lives "
                + "anywhere",
            CorpusStoreRefusal.ProfileUnprovable => "the account list, or one account's delivery store, could not be "
                + "read, so it is not proven that nothing here can send",
            _ => "unrecognised refusal",
        };

        if (refusal == CorpusStoreRefusal.None)
        {
            return $"Store '{target}' accepted as a corpus target.";
        }

        bool profileRefusal = refusal is CorpusStoreRefusal.TargetIsAccountDeliveryStore
            or CorpusStoreRefusal.ProfileCanSend
            or CorpusStoreRefusal.ProfileUnprovable;
        string remedy = profileRefusal
            ? "Use a profile with no mail accounts. There is no flag for this: a build creates unsent items in bulk, "
                + "and a real build put 5 532 of them into the target store's Outbox - inert only because that "
                + "profile could not send."
            : "A corpus may only be written into a LOCAL .pst named explicitly by the caller. Widen the "
                + "--allow-store list, never the guard.";

        return $"REFUSING to build a corpus in store '{target}': {why}. {remedy} "
            + "See the mailbox-safety rules in CLAUDE.md.";
    }

    /// <summary>
    /// The ONLY sanctioned delete predicate. True when both independent conditions hold:
    /// the EntryID is one this run's manifest recorded creating, AND the subject just
    /// re-read from the item carries <see cref="CorpusPlan.SubjectTag"/> and this corpus's
    /// per-item tag, matched ordinally.
    /// <para>
    /// The tag half is the CORPUS tag, which since 2026-08-25 is deliberately not the live
    /// tier's artifact tag: an artifact sweep must not be able to select a corpus item, and a
    /// corpus teardown must not be able to select an artifact. Rule 2's requirement - two
    /// independent keys, one of them an ordinal tag match - is unchanged; only which tag is
    /// matched. A subject carrying the OLD corpus tag returns false here, deliberately.
    /// </para>
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
    /// The predicate guarding an in-place REWRITE of an existing item - what
    /// <c>corpus-reanchor</c> does when it moves a corpus forward in time. It is
    /// <see cref="MayDelete"/> with one more key: the ordinal parsed out of the subject must
    /// be the ordinal the caller believes it is addressing.
    /// <para>
    /// A rewrite is guarded exactly like a delete because the blast radius is the same
    /// shape. Writing a delivery time onto somebody's mail is not recoverable from a manifest
    /// and would not even be visible as damage; the extra ordinal check exists because a
    /// rewrite, unlike a delete, is addressed PER ITEM from a plan, so an off-by-one in the
    /// caller would otherwise write item N's dates onto item M and leave a corpus that
    /// nothing could tell was wrong.
    /// </para>
    /// </summary>
    public static bool MayRewrite(
        string? entryId, string? subject, ISet<string> entryIdAllowlist, string corpusId, int expectedOrdinal)
    {
        ArgumentNullException.ThrowIfNull(entryIdAllowlist);
        if (string.IsNullOrWhiteSpace(entryId) || subject == null || expectedOrdinal < 1)
        {
            return false;
        }

        return entryIdAllowlist.Contains(entryId)
            && CorpusPlan.TryParseOrdinal(subject, corpusId, out int ordinal)
            && ordinal == expectedOrdinal;
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
