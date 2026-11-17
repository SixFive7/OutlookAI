using OutlookAI.Core.Mapi;
using OutlookAI.McpServer.Tests.T2;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The store and folder IDs a wire test passes (Q131-Q136), computed the way the product computes them
/// (<see cref="OutlookIds"/>) from what Outlook reports through the test's own read-only COM helper - so a
/// test that may not call list_accounts over the wire can still address a store by ID, and every ID the
/// wire then answers is checked against one computed independently of the server.
/// </summary>
internal static class WireIds
{
    /// <summary>The store's ID: <c>S</c> and the provider UID its root folder's id carries.</summary>
    internal static string StoreId(string storeDisplayName)
    {
        StoreIdentity store = LiveOutlookTestMailer.ReadStoreIdentity(storeDisplayName);
        return OutlookIds.StoreIdOf(store.RootFolderEntryId, store.StoreId)
            ?? throw new InvalidOperationException("Outlook reports no usable ids for '" + storeDisplayName + "'.");
    }

    /// <summary>The ID of the store's root folder - parent_id for a top-level create_folder.</summary>
    internal static string RootFolderId(string storeDisplayName)
    {
        return OutlookIds.FolderIdFromEntryId(LiveOutlookTestMailer.ReadStoreIdentity(storeDisplayName).RootFolderEntryId)
            ?? throw new InvalidOperationException("Outlook reports no root folder id for '" + storeDisplayName + "'.");
    }
}
