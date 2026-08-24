namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What a test intends to do to a store. Used by <see cref="StoreWriteAllowlist"/>.</summary>
public enum StoreWriteKind
{
    /// <summary>Transport: sending a mail from that store.</summary>
    Send = 0,

    /// <summary>Creating or editing a draft item in that store.</summary>
    Draft = 1,

    /// <summary>Deleting an item from that store.</summary>
    Delete = 2,

    /// <summary>Moving an item within that store.</summary>
    Move = 3,

    /// <summary>Creating or removing a folder in that store.</summary>
    Folder = 4,
}

/// <summary>
/// Code-enforced answer to "which mailbox may a test write to". Every mailbox-mutating
/// helper asks this first; a store outside the allowlist throws instead of running, so a
/// write to a delegate/shared or business mailbox is not a policy the agent must remember
/// but a state the process cannot reach.
/// <para>
/// Tiers (v3.MD S2 + the Q-it2-3a exception, nothing wider):
/// <list type="bullet">
/// <item><b>hub</b> - the designated test mailbox from the gitignored live-test settings:
/// every kind of write.</item>
/// <item><b>bystander stores</b> - stores DECLARED as watched-and-never-written: nothing,
/// ever, and the declaration beats the identity-draft grant rather than losing to it. The
/// count tripwire's whole value rests on there being a store whose contents no test can
/// explain, so the guarantee has to come from a declaration rather than from a store
/// happening not to appear in another list.</item>
/// <item><b>identity-draft stores</b> - the other configured primary accounts, granted ONLY
/// so the identity tests can create one tagged, never-displayed draft each and clean it
/// up: <see cref="StoreWriteKind.Draft"/> and <see cref="StoreWriteKind.Delete"/> only, no
/// send, no move, no folder work.</item>
/// <item><b>everything else</b> - delegate/shared mailboxes and any store not in the
/// settings: nothing, ever.</item>
/// </list>
/// </para>
/// <para>
/// Pure and CI-testable: it holds names, not COM handles, so T1 pins its behaviour without
/// Outlook and without any real store name reaching this PUBLIC repo (S6).
/// </para>
/// </summary>
public sealed class StoreWriteAllowlist
{
    private readonly HashSet<string> _identityDraftStores;
    private readonly HashSet<string> _bystanders;
    private readonly HashSet<string> _denied;

    /// <summary>Builds an allowlist around one hub store.</summary>
    /// <param name="hubStoreDisplayName">The designated test mailbox; the only store with full write rights.</param>
    /// <param name="identityDraftStores">Stores granted draft+delete for the identity tests. May be null.</param>
    /// <param name="knownReadOnlyStores">
    /// Stores known to be off limits (delegate/shared mailboxes). Purely for a louder error
    /// message - a store absent from every list is refused just as hard.
    /// </param>
    /// <param name="bystanderStores">
    /// Stores DECLARED watched-and-never-written, denied every kind of write.
    /// <para>
    /// These are NOT rejected when they also appear in <paramref name="identityDraftStores"/>,
    /// and that is the point: the runbook has the bystander listed in
    /// <c>expectedStoreDisplayNames</c> - it has to be, or the census never visits it and
    /// <c>list_accounts</c> exactness never counts it - which is exactly how it ended up inside
    /// the identity grant. The overlap is the ordinary configuration, so the declaration wins
    /// and nothing is said about it. A read-only store in the grant stays a hard refusal below:
    /// that one has no legitimate shape.
    /// </para>
    /// </param>
    public StoreWriteAllowlist(
        string hubStoreDisplayName,
        IEnumerable<string>? identityDraftStores = null,
        IEnumerable<string>? knownReadOnlyStores = null,
        IEnumerable<string>? bystanderStores = null)
    {
        if (string.IsNullOrWhiteSpace(hubStoreDisplayName))
        {
            throw new ArgumentException("A write allowlist needs a hub store.", nameof(hubStoreDisplayName));
        }

        HubStoreDisplayName = hubStoreDisplayName;
        _denied = new HashSet<string>(knownReadOnlyStores ?? [], StringComparer.OrdinalIgnoreCase);
        _bystanders = new HashSet<string>(
            (bystanderStores ?? []).Where(s => !string.IsNullOrWhiteSpace(s)), StringComparer.OrdinalIgnoreCase);
        _identityDraftStores = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string store in identityDraftStores ?? [])
        {
            if (string.IsNullOrWhiteSpace(store) || IsHub(store))
            {
                continue;
            }

            if (_denied.Contains(store))
            {
                // A store cannot be both granted and read-only; refuse to build a
                // contradictory allowlist rather than resolve it silently.
                throw new ArgumentException(
                    "A read-only store may not appear in the identity-draft grant.", nameof(identityDraftStores));
            }

            _identityDraftStores.Add(store);
        }
    }

    /// <summary>The designated test mailbox.</summary>
    public string HubStoreDisplayName { get; }

    /// <summary>Stores granted draft+delete only.</summary>
    public IReadOnlyCollection<string> IdentityDraftStores => _identityDraftStores;

    /// <summary>Stores declared watched-and-never-written. Denied every kind of write.</summary>
    public IReadOnlyCollection<string> Bystanders => _bystanders;

    /// <summary>True when <paramref name="storeDisplayName"/> was declared a bystander.</summary>
    public bool IsBystander(string? storeDisplayName)
    {
        return storeDisplayName != null && _bystanders.Contains(storeDisplayName);
    }

    /// <summary>
    /// Which of <paramref name="candidateStores"/> the identity tests may actually draft in:
    /// everything this allowlist grants <see cref="StoreWriteKind.Draft"/> to, minus the hub,
    /// in the order given and without repeats.
    /// <para>
    /// The two identity tests iterate this rather than "the configured stores that are not the
    /// hub", so the list of stores they write to and the list of stores they are PERMITTED to
    /// write to are one answer from one place. Derived the other way they can disagree, and a
    /// disagreement is a live test throwing at the guard halfway through - or, before the
    /// bystander tier existed, not throwing at all.
    /// </para>
    /// </summary>
    public IReadOnlyList<string> IdentityAccountsAmong(IEnumerable<string>? candidateStores)
    {
        List<string> accounts = new();
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        foreach (string store in candidateStores ?? [])
        {
            if (!IsHub(store) && IsAllowed(store, StoreWriteKind.Draft) && seen.Add(store))
            {
                accounts.Add(store);
            }
        }

        return accounts;
    }

    /// <summary>True when <paramref name="storeDisplayName"/> is the hub.</summary>
    public bool IsHub(string? storeDisplayName)
    {
        return storeDisplayName != null
            && string.Equals(storeDisplayName, HubStoreDisplayName, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>True when <paramref name="kind"/> is permitted against <paramref name="storeDisplayName"/>.</summary>
    public bool IsAllowed(string? storeDisplayName, StoreWriteKind kind)
    {
        if (string.IsNullOrWhiteSpace(storeDisplayName))
        {
            return false;
        }

        if (IsHub(storeDisplayName))
        {
            // The hub wins even over a bystander declaration, and the count tripwire refuses
            // the run when the two collide (TripwireWatchSoundness). Resolving it the other
            // way would deny every write in the suite and fail 100-odd tests far from the
            // mistake; resolving it this way produces one refusal that names it.
            return true;
        }

        // Ahead of the identity grant, not after it. A bystander is normally IN that grant -
        // the runbook lists it in expectedStoreDisplayNames - so a declaration checked second
        // is a declaration that never applies to the one store it was written for.
        if (_bystanders.Contains(storeDisplayName))
        {
            return false;
        }

        if (!_identityDraftStores.Contains(storeDisplayName))
        {
            return false;
        }

        return kind is StoreWriteKind.Draft or StoreWriteKind.Delete;
    }

    /// <summary>Throws unless <paramref name="kind"/> is permitted; returns the store name so call sites read as one expression.</summary>
    public string Assert(string? storeDisplayName, StoreWriteKind kind, string operation)
    {
        if (IsAllowed(storeDisplayName, kind))
        {
            return storeDisplayName!;
        }

        throw new InvalidOperationException(Explain(storeDisplayName, kind, operation));
    }

    /// <summary>The refusal message. Content-free: it names no mailbox but the offending one.</summary>
    public string Explain(string? storeDisplayName, StoreWriteKind kind, string operation)
    {
        string target = string.IsNullOrWhiteSpace(storeDisplayName) ? "(no store)" : storeDisplayName!;
        string why = _bystanders.Contains(target)
            ? "that store is a declared BYSTANDER - the count tripwire watches it precisely "
                + "because nothing writes to it, so no test may write to it"
            : _denied.Contains(target)
                ? "that store is a delegate/shared mailbox and is READ-ONLY for tests"
                : _identityDraftStores.Contains(target)
                    ? "that store is granted draft+delete only (identity tests), not "
                        + kind.ToString().ToLowerInvariant()
                    : "only the designated test mailbox may be written to";

        return "REFUSING '" + operation + "' (" + kind.ToString().ToLowerInvariant() + ") on store '" + target
            + "': " + why + ". See the mailbox-safety rules in CLAUDE.md; widen the live-test settings, never the guard.";
    }
}
