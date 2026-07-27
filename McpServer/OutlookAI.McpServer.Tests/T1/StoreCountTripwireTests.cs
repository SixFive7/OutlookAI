using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the per-store count tripwire (soak fix 16, part A3). The point of these tests is
/// the FIRING behaviour: a live run may add nothing and remove nothing outside the
/// designated test mailbox, while mail arriving during the run must never fail the suite.
/// </summary>
public sealed class StoreCountTripwireTests
{
    private const string Hub = "hub@example.test";
    private const string Other = "other@example.test";
    private const string DelegateStore = "Someone Else";

    private static Dictionary<string, IReadOnlyDictionary<string, int>> Census(
        params (string Store, (string Folder, int Count)[] Folders)[] stores)
    {
        Dictionary<string, IReadOnlyDictionary<string, int>> census = new(StringComparer.OrdinalIgnoreCase);
        foreach ((string store, (string Folder, int Count)[] folders) in stores)
        {
            Dictionary<string, int> byFolder = new(StringComparer.OrdinalIgnoreCase);
            foreach ((string folder, int count) in folders)
            {
                byFolder[folder] = count;
            }

            census[store] = byFolder;
        }

        return census;
    }

    private static Dictionary<string, IReadOnlyDictionary<string, int>> Baseline()
    {
        return Census(
            (Hub, new[] { ("Inbox", 2), ("Sent Items", 21), ("Deleted Items", 0), ("Outbox", 0) }),
            (Other, new[] { ("Inbox", 171), ("Sent Items", 4866), ("Sync Issues/Conflicts", 40) }),
            (DelegateStore, new[] { ("Inbox", 99), ("Archive", 6153), ("Deleted Items", 19525) }));
    }

    [Fact]
    public void NoChange_Passes()
    {
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), Baseline(), Hub);

        Assert.False(verdict.Failed);
        Assert.Empty(verdict.Failures);
        Assert.Empty(verdict.Notes);
    }

    [Fact]
    public void ItemsLostInADelegateStore_FiresAndNamesStoreFolderAndDelta()
    {
        // THE CASE THE TRIPWIRE EXISTS FOR (incident 7): items disappearing from a
        // mailbox the suite may only read.
        var after = Baseline();
        ((Dictionary<string, int>)after[DelegateStore])["Archive"] = 6130;

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        string message = verdict.Describe();
        Assert.Contains("ITEMS LOST", message, StringComparison.Ordinal);
        Assert.Contains(DelegateStore, message, StringComparison.Ordinal);
        Assert.Contains("Archive", message, StringComparison.Ordinal);
        Assert.Contains("6153 -> 6130", message, StringComparison.Ordinal);
        Assert.Contains("-23", message, StringComparison.Ordinal);
    }

    [Fact]
    public void ItemsLostInABusinessStore_Fires()
    {
        var after = Baseline();
        ((Dictionary<string, int>)after[Other])["Inbox"] = 170;

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("ITEMS LOST", StringComparison.Ordinal));
    }

    [Fact]
    public void MailArrivingElsewhereDuringTheRun_IsNotedNotFailed()
    {
        // An 8-minute live run happens while real mail arrives; increases outside the hub
        // are the normal case and must never fail the suite.
        var after = Baseline();
        ((Dictionary<string, int>)after[Other])["Inbox"] = 174;
        ((Dictionary<string, int>)after[DelegateStore])["Inbox"] = 101;

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.False(verdict.Failed);
        Assert.Equal(2, verdict.Notes.Count);
        Assert.Contains(verdict.Notes, n => n.Contains("171 -> 174 (+3)", StringComparison.Ordinal));
        Assert.Contains(verdict.Notes, n => n.Contains("99 -> 101 (+2)", StringComparison.Ordinal));
    }

    [Fact]
    public void FolderRemovedOutsideTheHub_Fires()
    {
        var after = Baseline();
        ((Dictionary<string, int>)after[Other]).Remove("Sync Issues/Conflicts");

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("FOLDER REMOVED", StringComparison.Ordinal)
            && f.Contains("Sync Issues/Conflicts", StringComparison.Ordinal));
    }

    [Fact]
    public void FolderAddedOutsideTheHub_Fires()
    {
        var after = Baseline();
        ((Dictionary<string, int>)after[DelegateStore])["OutlookAI-McpTest-Folder"] = 1;

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("FOLDER ADDED", StringComparison.Ordinal));
    }

    [Fact]
    public void HubChurn_IsToleratedByThisGuard()
    {
        // The hub legitimately gains, loses and reshapes tagged items; the zero-artifact
        // sweep and the move/archive reconciliation police it instead.
        var after = Baseline();
        ((Dictionary<string, int>)after[Hub])["Inbox"] = 5;
        ((Dictionary<string, int>)after[Hub])["Sent Items"] = 19;
        ((Dictionary<string, int>)after[Hub])["OutlookAI-McpTest-Folder"] = 1;
        ((Dictionary<string, int>)after[Hub]).Remove("Outbox");

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.False(verdict.Failed);
    }

    [Fact]
    public void AStoreThatVanishesOrAppears_Fires()
    {
        var missing = Baseline();
        missing.Remove(DelegateStore);
        Assert.True(StoreCountTripwire.Evaluate(Baseline(), missing, Hub).Failed);

        var extra = Baseline();
        extra["surprise@example.test"] = new Dictionary<string, int> { ["Inbox"] = 1 };
        Assert.True(StoreCountTripwire.Evaluate(Baseline(), extra, Hub).Failed);
    }

    [Fact]
    public void WatchedStores_CoverEveryPrimaryAndEveryDelegate()
    {
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Other },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore, "Another Person" },
        };

        IReadOnlyList<string> watched = LiveStoreCountTripwire.WatchedStores(settings);

        Assert.Equal(4, watched.Count);
        Assert.Contains(DelegateStore, watched);
        Assert.Contains("Another Person", watched);
    }
}
