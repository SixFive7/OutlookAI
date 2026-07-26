using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the D39 move_mail/archive_mail live tier: the MailService under
/// test, an independent verify session, and one unique run marker (S3 double-match).
/// ALL move/archive writes target the test hub only (S2); the other four stores are
/// touched READ-ONLY (archive-folder resolution checks). Disposal purges this run's
/// tagged artifacts (Drafts/Inbox/Sent/Archive/Deleted) AND any leftover test folders.
/// </summary>
public sealed class LiveMoveArchiveFixture : IDisposable
{
    private readonly Lazy<OutlookComSession> _verifySession;

    public LiveMoveArchiveFixture()
    {
        Settings = LiveTestSettings.Load();
        Service = MailService.CreateDefault();
        RunMarker = "d39" + Guid.NewGuid().ToString("N").Substring(0, 14);
        _verifySession = new Lazy<OutlookComSession>(
            () => OutlookComSession.Connect(allowStartingOutlook: true),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Per-run unique marker; every artifact subject carries tag + marker (S3).</summary>
    public string RunMarker { get; }

    /// <summary>Independent COM session for verification (and read-only resolution checks).</summary>
    public OutlookComSession VerifySession => _verifySession.Value;

    /// <summary>Builds a tagged subject: [OutlookAI-McpTest] + run marker + label.</summary>
    public string TaggedSubject(string label)
    {
        return LiveOutlookTestMailer.SubjectTag + " " + RunMarker + " " + label;
    }

    /// <summary>This run's test folder name (carries the folder name prefix, D39).</summary>
    public string TestFolderName => LiveOutlookTestMailer.TestFolderNamePrefix;

    /// <summary>StoreID of a store by display name (via the verify session).</summary>
    public string GetStoreId(string storeDisplayName)
    {
        return VerifySession.GetStores()
                .FirstOrDefault(s => string.Equals(s.DisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId
            ?? throw new InvalidOperationException("Store not found by display name.");
    }

    public void Dispose()
    {
        try
        {
            // Folders first (their tagged contents purge into Deleted Items), then
            // the item sweep - the same order the tests use.
            LiveOutlookTestMailer.DeleteTestFolders(Settings.TestHubStoreDisplayName);
        }
        catch (Exception)
        {
            // Best-effort final belt for folders.
        }

        try
        {
            LiveOutlookTestMailer.DeleteTaggedArtifacts(Settings.TestHubStoreDisplayName, RunMarker);
        }
        catch (Exception)
        {
            // Best-effort - each test already cleaned up and asserted in finally.
        }

        if (_verifySession.IsValueCreated)
        {
            // Releases COM references only - Outlook keeps running (S7: never kill/close).
            _verifySession.Value.Dispose();
        }

        Service.Dispose();
    }
}

[CollectionDefinition("LiveMoveArchive")]
public sealed class LiveMoveArchiveCollection : ICollectionFixture<LiveMoveArchiveFixture>
{
}
