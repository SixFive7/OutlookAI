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

    /// <summary>Builds the allowlist for <paramref name="settings"/> (pure - used by tests too).</summary>
    public static StoreWriteAllowlist Build(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
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
