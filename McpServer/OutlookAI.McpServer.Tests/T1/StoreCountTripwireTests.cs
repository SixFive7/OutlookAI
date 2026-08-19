using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the per-store tripwire (soak fix 16, part A3). The point of these tests is the
/// FIRING behaviour: a live run may add nothing and remove nothing outside the designated
/// test mailbox, while mail arriving during the run must never fail the suite.
/// <para>
/// The identity half was added after the 2026-08-18 false alarm, where three count
/// decreases in two mailboxes turned out to be the maintainer reading and deleting his own
/// mail during a 27-minute run. A count cannot tell that from a runaway test - nothing can,
/// from a before/after reading - so these tests pin the two things that CAN be decided:
/// mail that was filed rather than removed is not loss, and mail removed while other mail
/// arrives is still loss even though the count never moved.
/// </para>
/// </summary>
public sealed class StoreCountTripwireTests
{
    private const string Hub = "hub@example.test";
    private const string Other = "other@example.test";
    private const string DelegateStore = "Someone Else";
    private const string Volatile = StoreCountTripwire.VolatilePrefix;

    private static Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> Census(
        params (string Store, (string Folder, int Count)[] Folders)[] stores)
    {
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> census =
            new(StringComparer.OrdinalIgnoreCase);
        foreach ((string store, (string Folder, int Count)[] folders) in stores)
        {
            Dictionary<string, FolderCensus> byFolder = new(StringComparer.OrdinalIgnoreCase);
            foreach ((string folder, int count) in folders)
            {
                byFolder[folder] = FolderCensus.CountOnly(count);
            }

            census[store] = byFolder;
        }

        return census;
    }

    private static Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> Baseline()
    {
        return Census(
            (Hub, new[] { ("Inbox", 2), ("Sent Items", 21), ("Deleted Items", 0), ("Outbox", 0) }),
            (Other, new[] { ("Inbox", 171), ("Sent Items", 4866), (Volatile + "Sync Issues/Conflicts", 40) }),
            (DelegateStore, new[] { ("Inbox", 99), ("Archive", 6153), (Volatile + "Deleted Items", 19525) }));
    }

    private static void SetCount(
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> census,
        string store, string folder, int count)
    {
        ((Dictionary<string, FolderCensus>)census[store])[folder] = FolderCensus.CountOnly(count);
    }

    private static void SetItems(
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> census,
        string store, string folder, params CensusItem[] items)
    {
        ((Dictionary<string, FolderCensus>)census[store])[folder] = FolderCensus.WithItems(items);
    }

    private static CensusItem Item(string id, string? fingerprint = null, bool tagged = false)
    {
        return new CensusItem(id, fingerprint ?? "fp-" + id, tagged);
    }

    [Fact]
    public void NoChange_Passes()
    {
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), Baseline(), Hub);

        Assert.False(verdict.Failed);
        Assert.Empty(verdict.Failures);
        Assert.Empty(verdict.Notes);
        Assert.Null(verdict.Attribution);
    }

    [Fact]
    public void ItemsLostInADelegateStore_FiresAndNamesStoreFolderAndDelta()
    {
        // THE CASE THE TRIPWIRE EXISTS FOR (incident 7): items disappearing from a
        // mailbox the suite may only read.
        var after = Baseline();
        SetCount(after, DelegateStore, "Archive", 6130);

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
        SetCount(after, Other, "Inbox", 170);

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("ITEMS LOST", StringComparison.Ordinal));
    }

    [Fact]
    public void ACountedOnlyLoss_SaysItCannotNameTheItems()
    {
        // A folder above the identity budget is still guarded, and the message has to admit
        // what it cannot show - otherwise the reader assumes the detail simply was not
        // written down, and stops believing the ones that DO carry detail.
        var after = Baseline();
        SetCount(after, Other, "Sent Items", 4860);

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.Contains(
            verdict.Failures,
            f => f.Contains("WHICH items left is not known", StringComparison.Ordinal));
    }

    [Fact]
    public void MailArrivingElsewhereDuringTheRun_IsNotedNotFailed()
    {
        // An 8-minute live run happens while real mail arrives; increases outside the hub
        // are the normal case and must never fail the suite.
        var after = Baseline();
        SetCount(after, Other, "Inbox", 174);
        SetCount(after, DelegateStore, "Inbox", 101);

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
        ((Dictionary<string, FolderCensus>)after[Other]).Remove("Sent Items");

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("FOLDER REMOVED", StringComparison.Ordinal)
            && f.Contains("Sent Items", StringComparison.Ordinal));
    }

    [Fact]
    public void FolderAddedOutsideTheHub_Fires()
    {
        var after = Baseline();
        SetCount(after, DelegateStore, "OutlookAI-McpTest-Folder", 1);

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("FOLDER ADDED", StringComparison.Ordinal));
    }

    [Fact]
    public void SelfPruningFolders_ShrinkWithoutFailingTheSuite()
    {
        // Deleted Items ages out, junk mail expires, and Outlook writes/removes sync-issue
        // reports on its own. A tripwire that fails on those gets ignored, which is the one
        // outcome that must not happen - so they are noted, while ordinary mail folders
        // still fail.
        var after = Baseline();
        SetCount(after, DelegateStore, Volatile + "Deleted Items", 19000);
        ((Dictionary<string, FolderCensus>)after[Other]).Remove(Volatile + "Sync Issues/Conflicts");

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Baseline(), after, Hub);

        Assert.False(verdict.Failed);
        Assert.Contains(verdict.Notes, n => n.Contains("self-pruning", StringComparison.Ordinal));
        Assert.True(StoreCountTripwire.IsVolatile(Volatile + "Deleted Items"));
        Assert.False(StoreCountTripwire.IsVolatile("Inbox"));
    }

    [Fact]
    public void DelegateHierarchyChurn_IsNotedNotFailed_ButItemLossStillFails()
    {
        // A delegate/shared mailbox syncs its folder tree lazily: a real 450-item folder
        // was absent from one census and present again in the next. Folders appearing and
        // disappearing there says something about the hierarchy cache, not about deletion.
        var after = Baseline();
        ((Dictionary<string, FolderCensus>)after[DelegateStore]).Remove("Archive");
        SetCount(after, DelegateStore, "Archive/Sub", 12);

        Assert.False(StoreCountTripwire.Evaluate(Baseline(), after, Hub, new[] { DelegateStore }).Failed);

        // ...but the same store still fails on ITEM LOSS in a folder seen in both censuses -
        // that is the shape a mass deletion actually has.
        var lost = Baseline();
        SetCount(lost, DelegateStore, "Inbox", 3);
        Assert.True(StoreCountTripwire.Evaluate(Baseline(), lost, Hub, new[] { DelegateStore }).Failed);

        // Without the lazy grant, the same disappearance is a hard failure.
        Assert.True(StoreCountTripwire.Evaluate(Baseline(), after, Hub).Failed);
    }

    [Fact]
    public void HubChurn_IsToleratedByThisGuard()
    {
        // The hub legitimately gains, loses and reshapes tagged items; the zero-artifact
        // sweep and the move/archive reconciliation police it instead.
        var after = Baseline();
        SetCount(after, Hub, "Inbox", 5);
        SetCount(after, Hub, "Sent Items", 19);
        SetCount(after, Hub, "OutlookAI-McpTest-Folder", 1);
        ((Dictionary<string, FolderCensus>)after[Hub]).Remove("Outbox");

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
        extra["surprise@example.test"] = new Dictionary<string, FolderCensus> { ["Inbox"] = FolderCensus.CountOnly(1) };
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

    [Fact]
    public void ARemovedItem_IsNamedByEntryIdAndSaysWhereItWent()
    {
        // What the maintainer needs the moment this fires: WHICH items, and where they are
        // now. "-7" is unfalsifiable; an EntryID that is sitting in Deleted Items can be
        // confirmed or refuted in seconds.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a"), Item("id-b"), Item("id-c"));
        SetItems(before, Other, Volatile + "Deleted Items", Item("id-old"));

        var after = Baseline();
        SetItems(after, Other, "Inbox", Item("id-a"), Item("id-b"));
        SetItems(after, Other, Volatile + "Deleted Items", Item("id-old"), Item("id-c-moved", "fp-id-c"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        string message = verdict.Describe();
        Assert.Contains("ITEMS REMOVED", message, StringComparison.Ordinal);
        Assert.Contains("id-c", message, StringComparison.Ordinal);
        Assert.Contains("now in 'Deleted Items (self-pruning)'", message, StringComparison.Ordinal);
    }

    [Fact]
    public void MailFiledIntoAnotherFolderOfTheSameStore_IsNotLoss()
    {
        // The one exoneration a before/after census can actually PROVE, and the one a count
        // can never see: the item is still there, in a folder the same census walked.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a"), Item("id-b"));
        SetItems(before, Other, "Projects");

        var after = Baseline();
        SetItems(after, Other, "Inbox", Item("id-a"));
        SetItems(after, Other, "Projects", Item("id-b-reissued", "fp-id-b"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.False(verdict.Failed);
        Assert.Contains(verdict.Notes, n => n.Contains("filed (not loss)", StringComparison.Ordinal)
            && n.Contains("Projects", StringComparison.Ordinal));
    }

    [Fact]
    public void ARemovalMaskedByAnArrival_StillFires()
    {
        // The half nobody had looked at: on a busy machine a count can stay put while an
        // item is destroyed and another arrives. This is invisible to the count rule.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a"), Item("id-b"));

        var after = Baseline();
        SetItems(after, Other, "Inbox", Item("id-a"), Item("id-new"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("id-b", StringComparison.Ordinal));
        Assert.Contains(
            verdict.Failures,
            f => f.Contains("not found in any folder this census identified", StringComparison.Ordinal));
    }

    [Fact]
    public void AnItemAlreadySittingInDeletedItems_CannotExplainADeparture()
    {
        // The relocation index is built from ARRIVALS only. If it were built from the whole
        // after-census, any old item with a colliding fingerprint would excuse a deletion.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a", "same-fp"));
        SetItems(before, Other, "Projects", Item("id-b", "same-fp"));

        var after = Baseline();
        SetItems(after, Other, "Inbox");
        SetItems(after, Other, "Projects", Item("id-b", "same-fp"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("id-a", StringComparison.Ordinal));
    }

    [Fact]
    public void OneArrivalCannotExcuseTwoDepartures()
    {
        // Each arrival is consumed by at most one departure, so a single unrelated move
        // cannot account for a folder that lost two items.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a", "fp-x"), Item("id-b", "fp-x"));
        SetItems(before, Other, "Projects");

        var after = Baseline();
        SetItems(after, Other, "Inbox");
        SetItems(after, Other, "Projects", Item("id-moved", "fp-x"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("lost 1 item(s)", StringComparison.Ordinal));
        Assert.Contains(verdict.Notes, n => n.Contains("filed (not loss)", StringComparison.Ordinal));
    }

    [Fact]
    public void ATaggedItemLeavingAMailboxTheSuiteMayNotWriteTo_IsAttributedToTheSuite()
    {
        // The only attribution the evidence supports outright: the live tier's own tag,
        // in a store the write allowlist forbids it from touching.
        var before = Baseline();
        SetItems(before, DelegateStore, "Inbox", Item("id-a"), Item("id-tagged", tagged: true));

        var after = Baseline();
        SetItems(after, DelegateStore, "Inbox", Item("id-a"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        Assert.NotNull(verdict.Attribution);
        Assert.Contains("ATTRIBUTION: THE SUITE", verdict.Attribution!, StringComparison.Ordinal);
        Assert.Contains(LiveOutlookTestMailer.SubjectTag, verdict.Describe(), StringComparison.Ordinal);
        Assert.Contains(verdict.Failures, f => f.Contains("TEST-TAGGED", StringComparison.Ordinal));
    }

    [Fact]
    public void AnUntaggedRemoval_SaysTheActorIsUndecidableAndStillFails()
    {
        // The 2026-08-18 reading. It fails, because it must - and it says why it cannot
        // tell the maintainer from a runaway test rather than implying it can.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a"), Item("id-b"));

        var after = Baseline();
        SetItems(after, Other, "Inbox", Item("id-a"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        Assert.NotNull(verdict.Attribution);
        Assert.Contains("ATTRIBUTION: undecidable", verdict.Attribution!, StringComparison.Ordinal);
        Assert.Contains(
            "cannot name the actor", verdict.Attribution!, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderIdentifiedOnOnlyOneSide_FallsBackToTheCountRule()
    {
        // Half a comparison is not a comparison. The count rule still guards the folder.
        var before = Baseline();
        SetItems(before, Other, "Inbox", Item("id-a"), Item("id-b"), Item("id-c"));

        var after = Baseline();
        SetCount(after, Other, "Inbox", 2);

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.True(verdict.Failed);
        Assert.Contains(verdict.Failures, f => f.Contains("ITEMS LOST", StringComparison.Ordinal));

        // ...and the same fallback must not invent a loss when the count held.
        var held = Baseline();
        SetCount(held, Other, "Inbox", 3);
        Assert.False(StoreCountTripwire.Evaluate(before, held, Hub).Failed);
    }

    [Fact]
    public void DeparturesInSelfPruningFoldersAndInTheHub_AreNotFailures()
    {
        // Identity does not widen where the guard fires: Deleted Items still prunes itself,
        // and the hub is still policed by the zero-artifact sweep instead.
        var before = Baseline();
        SetItems(before, Other, Volatile + "Sync Issues/Conflicts", Item("id-x"), Item("id-y"));
        SetItems(before, Hub, "Inbox", Item("id-hub-1"), Item("id-hub-2"));

        var after = Baseline();
        SetItems(after, Other, Volatile + "Sync Issues/Conflicts", Item("id-x"));
        SetItems(after, Hub, "Inbox", Item("id-hub-1"));

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, Hub);

        Assert.False(verdict.Failed);
    }

    [Fact]
    public void ManyRemovals_ReportABoundedListAndSayHowManyMore()
    {
        // A mass deletion must not bury its own headline under a thousand EntryIDs.
        List<CensusItem> before = new();
        for (int i = 0; i < 40; i++)
        {
            before.Add(Item("id-" + i));
        }

        var baseline = Baseline();
        SetItems(baseline, Other, "Inbox", before.ToArray());

        var after = Baseline();
        SetItems(after, Other, "Inbox");

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(baseline, after, Hub);

        Assert.True(verdict.Failed);
        string failure = Assert.Single(verdict.Failures, f => f.Contains("ITEMS REMOVED", StringComparison.Ordinal));
        Assert.Contains("lost 40 item(s)", failure, StringComparison.Ordinal);
        Assert.Contains(
            "and " + (40 - StoreCountTripwire.MaxReportedDepartures) + " more",
            failure,
            StringComparison.Ordinal);
    }

    [Fact]
    public void AConfirmingCensus_MatchesOnTheSubjectOfTheFailure_NotOnItsTallies()
    {
        // A suspected loss is re-censused and only what fails BOTH times is reported. That
        // comparison used to be made on the rendered message, which carries the folder's
        // before/after counts and how many items arrived - and every one of those moves when
        // a single mail lands between the two censuses. On a real profile during a 27-minute
        // run that is ordinary, so a genuine mass deletion could be dismissed as "enumeration
        // noise". The key names the SUBJECT of the failure and nothing that moves.
        var baseline = Baseline();
        SetItems(baseline, Other, "Inbox", Item("id-a"), Item("id-b"));

        var after = Baseline();
        SetItems(after, Other, "Inbox", Item("id-a"));

        var confirmation = Baseline();
        SetItems(confirmation, Other, "Inbox", Item("id-a"), Item("id-arrived-since"));

        TripwireVerdict first = StoreCountTripwire.Evaluate(baseline, after, Hub);
        TripwireVerdict second = StoreCountTripwire.Evaluate(baseline, confirmation, Hub);

        Assert.True(first.Failed);
        Assert.True(second.Failed);

        // The messages differ - the second census saw an arrival - so a string intersection
        // would report nothing and wave the loss through.
        Assert.NotEqual(first.Failures.Single(), second.Failures.Single());
        Assert.Empty(first.Failures.Intersect(second.Failures, StringComparer.Ordinal));

        // The keys agree, so the loss is confirmed.
        Assert.Equal(first.FailureRecords.Single().Key, second.FailureRecords.Single().Key);
    }

    [Fact]
    public void FailureKeys_SeparateTheKindTheStoreAndTheFolder()
    {
        var baseline = Baseline();
        SetItems(baseline, Other, "Inbox", Item("id-a"));

        var after = Baseline();
        SetItems(after, Other, "Inbox");

        TripwireFailure failure = Assert.Single(StoreCountTripwire.Evaluate(baseline, after, Hub).FailureRecords);

        Assert.Equal("items-removed|" + Other + "|Inbox", failure.Key);
        Assert.Contains("ITEMS REMOVED", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FolderCensus_CountsWhatItWalked()
    {
        FolderCensus walked = FolderCensus.WithItems(new[] { Item("a"), Item("b") });
        Assert.True(walked.HasIdentities);
        Assert.Equal(2, walked.Count);

        FolderCensus counted = FolderCensus.CountOnly(9);
        Assert.False(counted.HasIdentities);
        Assert.Equal(9, counted.Count);
        Assert.Null(counted.Items);
    }
}
