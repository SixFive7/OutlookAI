using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 tests for the System.ItemPathDisplay fallback derivation (v3.MD section 4):
/// "/StoreDisplayName/FolderA/.../ItemName" -> store + folder path + item name by
/// stripping the first and last segments. Synthetic paths only (S6).
/// </summary>
public sealed class ItemPathFallbackTests
{
    [Fact]
    public void TryDerive_NestedPath_SplitsStoreFoldersAndItem()
    {
        Assert.True(ItemPathFallback.TryDerive("/alice@example.com/Inbox/Projects/Q3 report", out ItemPathFallback? fallback));

        Assert.Equal("alice@example.com", fallback!.StoreDisplayName);
        Assert.Equal(new[] { "Inbox", "Projects" }, fallback.FolderPath);
        Assert.Equal("Q3 report", fallback.ItemDisplayName);
    }

    [Fact]
    public void TryDerive_ItemDirectlyUnderStore_HasEmptyFolderPath()
    {
        Assert.True(ItemPathFallback.TryDerive("/alice@example.com/root item", out ItemPathFallback? fallback));

        Assert.Empty(fallback!.FolderPath);
        Assert.Equal("root item", fallback.ItemDisplayName);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("no-leading-slash/Inbox/x")]
    [InlineData("/only-store")]
    [InlineData("//starts-empty/x")]
    public void TryDerive_RejectsMalformedPaths(string? path)
    {
        Assert.False(ItemPathFallback.TryDerive(path, out ItemPathFallback? fallback));
        Assert.Null(fallback);
    }
}
