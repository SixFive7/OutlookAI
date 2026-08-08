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
    private readonly HashSet<string> _denied;

    /// <summary>Builds an allowlist around one hub store.</summary>
    /// <param name="hubStoreDisplayName">The designated test mailbox; the only store with full write rights.</param>
    /// <param name="identityDraftStores">Stores granted draft+delete for the identity tests. May be null.</param>
    /// <param name="knownReadOnlyStores">
    /// Stores known to be off limits (delegate/shared mailboxes). Purely for a louder error
    /// message - a store absent from every list is refused just as hard.
    /// </param>
    public StoreWriteAllowlist(
        string hubStoreDisplayName,
        IEnumerable<string>? identityDraftStores = null,
        IEnumerable<string>? knownReadOnlyStores = null)
    {
        if (string.IsNullOrWhiteSpace(hubStoreDisplayName))
        {
            throw new ArgumentException("A write allowlist needs a hub store.", nameof(hubStoreDisplayName));
        }

        HubStoreDisplayName = hubStoreDisplayName;
        _denied = new HashSet<string>(knownReadOnlyStores ?? [], StringComparer.OrdinalIgnoreCase);
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
            return true;
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
        string why = _denied.Contains(target)
            ? "that store is a delegate/shared mailbox and is READ-ONLY for tests"
            : _identityDraftStores.Contains(target)
                ? "that store is granted draft+delete only (identity tests), not " + kind.ToString().ToLowerInvariant()
                : "only the designated test mailbox may be written to";

        return "REFUSING '" + operation + "' (" + kind.ToString().ToLowerInvariant() + ") on store '" + target
            + "': " + why + ". See the mailbox-safety rules in CLAUDE.md; widen the live-test settings, never the guard.";
    }
}
