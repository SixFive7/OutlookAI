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
/// <item><b>the throwaway data file</b> (Q96 (iv), 2026-10-03) - the data file with no Drafts
/// folder that <c>Testbed/guest/Reset-ThrowawayStore.ps1</c> recreates before every live run,
/// granted the same two kinds and nothing else: the created-folder proof saves one tagged post
/// in it, replies to that post - which makes the product create the store's Drafts folder - and
/// discards the reply. It may be none of the other tiers; a contradictory declaration refuses to
/// build.</item>
/// <item><b>everything else</b> - delegate/shared mailboxes and any store not in the
/// settings: nothing, ever.</item>
/// </list>
/// </para>
/// <para>
/// <b>Above all four tiers: a READ-ONLY MACHINE (Q74, 2026-10-03).</b> An allowlist built by
/// <see cref="RefusingEveryWrite"/> refuses every kind of write to every store - the hub too, and
/// ahead of the hub check, because the hub is the one store the tiers above would otherwise let
/// through. <see cref="LiveStoreWriteGuard.Build"/> builds one whenever the settings' machine
/// profile is read-only (<see cref="LiveWriteAccess.RefusesEveryWrite"/>), which since Q72 is the
/// maintainer's workstation: there, an in-process write throws instead of reaching his mailbox,
/// and the identity grant that used to open draft+delete on his other primary mailboxes is gone
/// with it.
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
    private readonly HashSet<string> _throwaways;

    /// <summary>
    /// Why every write is refused, or null on a machine that may write. Set only by
    /// <see cref="RefusingEveryWrite"/>.
    /// </summary>
    private readonly string? _everyWriteRefusedBecause;

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
    /// <param name="throwawayStores">
    /// The throwaway data file(s) - granted draft+delete for the created-folder proof (Q96 (iv)).
    /// May be null. One that is also the hub, a bystander, a read-only store or an identity-draft
    /// store is a contradiction and refuses to build.
    /// </param>
    public StoreWriteAllowlist(
        string hubStoreDisplayName,
        IEnumerable<string>? identityDraftStores = null,
        IEnumerable<string>? knownReadOnlyStores = null,
        IEnumerable<string>? bystanderStores = null,
        IEnumerable<string>? throwawayStores = null)
        : this(hubStoreDisplayName, identityDraftStores, knownReadOnlyStores, bystanderStores, throwawayStores, everyWriteRefusedBecause: null)
    {
    }

    private StoreWriteAllowlist(
        string hubStoreDisplayName,
        IEnumerable<string>? identityDraftStores,
        IEnumerable<string>? knownReadOnlyStores,
        IEnumerable<string>? bystanderStores,
        IEnumerable<string>? throwawayStores,
        string? everyWriteRefusedBecause)
    {
        _everyWriteRefusedBecause = everyWriteRefusedBecause;
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

        _throwaways = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string store in throwawayStores ?? [])
        {
            if (string.IsNullOrWhiteSpace(store))
            {
                continue;
            }

            if (IsHub(store) || _bystanders.Contains(store) || _denied.Contains(store) || _identityDraftStores.Contains(store))
            {
                // The throwaway is granted draft+delete BECAUSE it holds nothing but test artifacts
                // and is recreated before every run. Every other tier says something else about the
                // same store, so a store in two of them is a configuration nobody can mean.
                throw new ArgumentException(
                    "The throwaway data file '" + store + "' may not also be the hub, a bystander, a read-only store "
                    + "or an identity-draft store.",
                    nameof(throwawayStores));
            }

            _throwaways.Add(store);
        }
    }

    /// <summary>
    /// An allowlist for a READ-ONLY MACHINE: the same declarations as the public constructor -
    /// so <see cref="IsHub"/>, <see cref="IsBystander"/> and <see cref="IsKnownReadOnly"/> still
    /// answer, and a contradictory configuration is still refused - but every kind of write to
    /// every store is refused, the hub included, with <paramref name="reason"/> as the explanation.
    /// </summary>
    /// <param name="reason">Why nothing may be written here - normally <see cref="LiveWriteAccess.ReadOnlyReason"/>.</param>
    /// <param name="hubStoreDisplayName">The designated test mailbox. Named, and refused like the rest.</param>
    /// <param name="identityDraftStores">The other primaries. Granted nothing here.</param>
    /// <param name="knownReadOnlyStores">Delegate/shared mailboxes, for the error message.</param>
    /// <param name="bystanderStores">Declared bystanders, for the error message.</param>
    /// <param name="throwawayStores">A declared throwaway data file. Granted nothing here either.</param>
    public static StoreWriteAllowlist RefusingEveryWrite(
        string reason,
        string hubStoreDisplayName,
        IEnumerable<string>? identityDraftStores = null,
        IEnumerable<string>? knownReadOnlyStores = null,
        IEnumerable<string>? bystanderStores = null,
        IEnumerable<string>? throwawayStores = null)
    {
        if (string.IsNullOrWhiteSpace(reason))
        {
            throw new ArgumentException(
                "A read-only allowlist must say why nothing may be written - the refusal is the only thing whoever hits it reads.",
                nameof(reason));
        }

        return new StoreWriteAllowlist(hubStoreDisplayName, identityDraftStores, knownReadOnlyStores, bystanderStores, throwawayStores, reason);
    }

    /// <summary>True when this allowlist refuses every write to every store, the hub included.</summary>
    public bool RefusesEveryWrite => _everyWriteRefusedBecause != null;

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

    /// <summary>The declared throwaway data file(s) - granted draft+delete only (Q96 (iv)).</summary>
    public IReadOnlyCollection<string> ThrowawayStores => _throwaways;

    /// <summary>True when <paramref name="storeDisplayName"/> was declared the throwaway data file.</summary>
    public bool IsThrowaway(string? storeDisplayName)
    {
        return storeDisplayName != null && _throwaways.Contains(storeDisplayName);
    }

    /// <summary>
    /// True when <paramref name="storeDisplayName"/> was declared a delegate/shared mailbox -
    /// read-only for tests, per mailbox-safety rule 3.
    /// <para>
    /// Exposed for the same reason <see cref="IsBystander"/> is: a store the sweep may not delete
    /// from has to be able to say WHICH kind of off-limits it is, and "a delegate mailbox" and
    /// "a store nobody granted anything" read very differently to whoever hits the refusal.
    /// <see cref="Explain"/> already made that distinction inside a message; this makes it
    /// available to a caller deciding what to say.
    /// </para>
    /// </summary>
    public bool IsKnownReadOnly(string? storeDisplayName)
    {
        return storeDisplayName != null && _denied.Contains(storeDisplayName);
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
            // The throwaway data file is granted a draft for the created-folder proof, not for the
            // identity tests: it is no account's store, and it is in no candidate list anyway.
            if (!IsHub(store) && !IsThrowaway(store) && IsAllowed(store, StoreWriteKind.Draft) && seen.Add(store))
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

        // A read-only machine, ahead of the hub check: the hub is the one store every tier below
        // lets through, and on the maintainer's workstation it is his real mailbox too.
        if (RefusesEveryWrite)
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

        if (!_identityDraftStores.Contains(storeDisplayName) && !_throwaways.Contains(storeDisplayName))
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
        if (RefusesEveryWrite)
        {
            // Not "widen the live-test settings": on a read-only machine that advice is exactly the
            // edit that must never be made.
            return "REFUSING '" + operation + "' (" + kind.ToString().ToLowerInvariant() + ") on store '" + target
                + "': " + _everyWriteRefusedBecause + " Do NOT change machineProfile to make this pass - run the test on a "
                + "test guest.";
        }

        string why = _bystanders.Contains(target)
            ? "that store is a declared BYSTANDER - the count tripwire watches it precisely "
                + "because nothing writes to it, so no test may write to it"
            : _denied.Contains(target)
                ? "that store is a delegate/shared mailbox and is READ-ONLY for tests"
                : _identityDraftStores.Contains(target)
                    ? "that store is granted draft+delete only (identity tests), not "
                        + kind.ToString().ToLowerInvariant()
                    : _throwaways.Contains(target)
                        ? "that store is the throwaway data file, granted draft+delete only (the created-folder "
                            + "proof), not " + kind.ToString().ToLowerInvariant()
                        : "only the designated test mailbox may be written to";

        return "REFUSING '" + operation + "' (" + kind.ToString().ToLowerInvariant() + ") on store '" + target
            + "': " + why + ". See the mailbox-safety rules in AGENTS.md; widen the live-test settings, never the guard.";
    }
}
