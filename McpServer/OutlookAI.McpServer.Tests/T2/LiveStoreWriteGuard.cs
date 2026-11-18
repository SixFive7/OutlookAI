namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Process-wide access to the <see cref="StoreWriteAllowlist"/> built from the gitignored
/// live-test settings. Every mailbox-mutating helper calls this before touching COM, so a
/// write outside the allowlist throws in the test process instead of reaching Outlook.
/// <para>
/// The allowlist is derived, never hand-written: hub = the designated test mailbox,
/// identity-draft grant = the other configured primary accounts, read-only = the
/// configured delegate/shared mailboxes, and denied outright = the declared BYSTANDER
/// stores, which the count tripwire watches on the premise that nothing writes to them.
/// The bystander declaration is passed LAST and wins over the grant above it, because the
/// runbook has the bystander listed among the primary accounts and would otherwise be
/// declaring something the allowlist ignores.
/// </para>
/// </summary>
public static class LiveStoreWriteGuard
{
    private static readonly object Gate = new();
    private static StoreWriteAllowlist? _allowlist;

    /// <summary>The active allowlist, loading the live-test settings on first use.</summary>
    public static StoreWriteAllowlist Allowlist
    {
        get
        {
            lock (Gate)
            {
                return _allowlist ??= Build(LiveTestSettings.Load());
            }
        }
    }

    /// <summary>
    /// Builds the allowlist for <paramref name="settings"/> (pure - used by tests too).
    /// <para>
    /// On a READ-ONLY machine (<see cref="LiveWriteAccess.RefusesEveryWrite"/> - since Q72 the
    /// maintainer's workstation, whose settings declare Production or nothing at all) it is the
    /// allowlist that refuses EVERY store, the hub included (Q74 layer 2). The same declarations go
    /// into it either way, so the tripwire's soundness check and the artifact sweep still see which
    /// store is the hub, a bystander or a delegate - they just find nothing writable.
    /// </para>
    /// </summary>
    public static StoreWriteAllowlist Build(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        if (LiveWriteAccess.RefusesEveryWrite(settings.MachineProfile))
        {
            return StoreWriteAllowlist.RefusingEveryWrite(
                LiveWriteAccess.ReadOnlyReason(settings.MachineProfile),
                settings.TestHubStoreDisplayName,
                settings.ExpectedStoreDisplayNames,
                settings.ExpectedDelegateStoreDisplayNames,
                settings.BystanderStoreDisplayNames);
        }

        // The Exchange VM's test account: the hub and nothing else. No identity grant at all - every other
        // store there is real mail - so a store the settings forgot to declare a bystander is refused as hard
        // as one they did declare (LiveTestSettings.Validate refuses such a file anyway).
        if (settings.MachineProfile == LiveMachineProfile.ExchangeTestAccount)
        {
            return new StoreWriteAllowlist(
                settings.TestHubStoreDisplayName,
                identityDraftStores: null,
                settings.ExpectedDelegateStoreDisplayNames,
                settings.BystanderStoreDisplayNames);
        }

        return new StoreWriteAllowlist(
            settings.TestHubStoreDisplayName,
            settings.ExpectedStoreDisplayNames,
            settings.ExpectedDelegateStoreDisplayNames,
            settings.BystanderStoreDisplayNames);
    }

    /// <summary>Throws unless the write is permitted; returns the store name.</summary>
    public static string Assert(string? storeDisplayName, StoreWriteKind kind, string operation)
    {
        return Allowlist.Assert(storeDisplayName, kind, operation);
    }

    /// <summary>Fluent form for call sites: <c>Service.NewDraft(Writable(store, ...), ...)</c>.</summary>
    public static string Writable(string? storeDisplayName, StoreWriteKind kind, string operation)
    {
        return Assert(storeDisplayName, kind, operation);
    }
}
