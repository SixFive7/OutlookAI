using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 argument-validation coverage for the Phase-3 surface: every rule below fires
/// BEFORE any COM or index access, so these run anywhere (CI included). The gateway is
/// constructed with autostart disabled as a tripwire - if validation ever regressed
/// into touching Outlook, these tests would fail with OutlookUnavailableException
/// instead of ArgumentException.
/// </summary>
public sealed class Phase3ValidationTests : IDisposable
{
    private readonly MailService _service = new(new ComGateway(allowStartingOutlook: false));

    public void Dispose()
    {
        _service.Dispose();
    }

    [Fact]
    public void Exhaustive_WithoutStore_Throws()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => _service.Search(new SearchRequest
        {
            Exhaustive = true,
            Query = "anything",
        }));
        Assert.Contains("store", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Exhaustive_WithoutFolderOrAfterBound_Throws()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => _service.Search(new SearchRequest
        {
            Exhaustive = true,
            Query = "anything",
            Store = "someone@example.com",
        }));
        Assert.Contains("bound", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Exhaustive_WithRecipientFilter_Throws()
    {
        Assert.Throws<ArgumentException>(() => _service.Search(new SearchRequest
        {
            Exhaustive = true,
            Store = "someone@example.com",
            Folder = "Inbox",
            To = "someone",
        }));
    }

    [Fact]
    public void Exhaustive_WithAttachmentHitsOnly_Throws()
    {
        Assert.Throws<ArgumentException>(() => _service.Search(new SearchRequest
        {
            Exhaustive = true,
            Store = "someone@example.com",
            Folder = "Inbox",
            AttachmentHitsOnly = true,
        }));
    }

    [Theory]
    [InlineData("current_folder", 0)]
    [InlineData("all_folders", 1)]
    [InlineData("all_outlook", 2)]
    [InlineData("subfolders", 3)]
    [InlineData("  Current_Folder  ", 0)]
    public void MapSearchScope_KnownValues(string scope, int expected)
    {
        Assert.Equal(expected, MailService.MapSearchScope(scope));
    }

    [Theory]
    [InlineData("everything")]
    [InlineData("")]
    public void MapSearchScope_UnknownValues_Throw(string scope)
    {
        Assert.Throws<ArgumentException>(() => MailService.MapSearchScope(scope));
    }

    [Fact]
    public void ShowSearchResults_BlankQuery_Throws()
    {
        Assert.Throws<ArgumentException>(() => _service.ShowSearchResults("  "));
    }

    [Fact]
    public void ShowSearchResults_ControlCharacters_Throw()
    {
        Assert.Throws<ArgumentException>(() => _service.ShowSearchResults("a\nb"));
    }

    [Fact]
    public void ShowSearchResults_FolderWithoutStore_Throws()
    {
        Assert.Throws<ArgumentException>(() => _service.ShowSearchResults("term", "current_folder", null, "Inbox"));
    }

    [Fact]
    public void ShowSearchResults_UnknownScope_Throws()
    {
        Assert.Throws<ArgumentException>(() => _service.ShowSearchResults("term", "everywhere"));
    }

    [Fact]
    public void GotoFolder_BlankStore_Throws()
    {
        Assert.Throws<ArgumentException>(() => _service.GotoFolder("  "));
    }

    [Fact]
    public void OpenInOutlook_UnknownHitId_Throws()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => _service.OpenInOutlook("h424242"));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }
}
