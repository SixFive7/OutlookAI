using System.Reflection;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the throwaway data file of Q96 (iv) (2026-10-03) everywhere CI can reach it: what the live-test
/// settings accept, what the write allowlist grants it - and, on the maintainer's read-only workstation,
/// does not - and every conclusion the live created-folder proof (<see cref="LiveCreatedFolderTests"/>)
/// may draw, decided in <see cref="ThrowawayStoreProof"/>. The live proof itself runs only on a test
/// guest, after <c>Testbed/guest/Reset-ThrowawayStore.ps1</c> has recreated the store; what no pure
/// function can carry - that the live test still ASKS each judge, writes only behind the guard and
/// cleans up whatever happens - is read out of its source.
/// <para>
/// Synthetic names only (S6). Nothing here touches Outlook, a mailbox or a machine-local settings file.
/// </para>
/// </summary>
public sealed class ThrowawayStoreTests
{
    private const string Hub = "hub@example.test";
    private const string Identity = "identity@example.test";
    private const string Bystander = "bystander@example.test";
    private const string Throwaway = "throwaway@example.test";

    /// <summary>A test guest's settings with a throwaway data file - the shape the renderer writes.</summary>
    private static LiveTestSettings Guest()
    {
        return new LiveTestSettings
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Bystander, Identity },
            BystanderStoreDisplayNames = new List<string> { Bystander },
            ThrowawayStoreDisplayName = Throwaway,
        };
    }

    // ------------------------------------------------------------------ the settings

    [Fact]
    public void AGuestThatDeclaresAThrowawayDataFile_Loads_AndOneThatDeclaresNoneLoadsToo()
    {
        LiveTestSettings.Validate(Guest());
        Assert.Contains("throwawayStore=set", Guest().Describe(), StringComparison.Ordinal);

        LiveTestSettings none = Guest();
        none.ThrowawayStoreDisplayName = null;
        LiveTestSettings.Validate(none);
        Assert.Contains("throwawayStore=none", none.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheSettingsFileSpellingIsRead()
    {
        LiveTestSettings parsed = LiveTestSettings.Parse(
            "{ \"machineProfile\": \"Portable\", \"testHubStoreDisplayName\": \"" + Hub + "\", "
            + "\"expectedStoreDisplayNames\": [\"" + Hub + "\", \"" + Bystander + "\"], "
            + "\"bystanderStoreDisplayNames\": [\"" + Bystander + "\"], "
            + "\"throwawayStoreDisplayName\": \"" + Throwaway + "\" }");

        Assert.Equal(Throwaway, parsed.ThrowawayStoreDisplayName);
    }

    [Fact]
    public void ABlankThrowaway_IsRefused()
    {
        LiveTestSettings settings = Guest();
        settings.ThrowawayStoreDisplayName = " ";

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(settings));
        Assert.Contains("blank 'throwawayStoreDisplayName'", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheHubAsTheThrowaway_IsRefused_BecauseItHasADraftsFolderAndWouldProveNothing()
    {
        LiveTestSettings settings = Guest();
        settings.ThrowawayStoreDisplayName = "HUB@example.test";

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(settings));
        Assert.Contains("name the hub 'HUB@example.test' as 'throwawayStoreDisplayName'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("proved nothing", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("expectedStoreDisplayNames")]
    [InlineData("indexedStoreDisplayNames")]
    [InlineData("expectedDelegateStoreDisplayNames")]
    [InlineData("bystanderStoreDisplayNames")]
    public void TheThrowawayInAnyOtherList_IsRefused(string list)
    {
        // Watched, the tripwire fails the run over the Drafts folder the proof creates; a primary, the
        // identity tests draft in it; a bystander or a delegate, nothing may write to it; indexed, the
        // index tests demand mail it never has. Each list, separately.
        LiveTestSettings settings = Guest();
        switch (list)
        {
            case "expectedStoreDisplayNames":
                settings.ExpectedStoreDisplayNames.Add(Throwaway);
                break;
            case "indexedStoreDisplayNames":
                settings.ExpectedStoreDisplayNames.Add(Throwaway);
                settings.IndexedStoreDisplayNames = new List<string> { Hub, Throwaway };
                break;
            case "expectedDelegateStoreDisplayNames":
                settings.ExpectedDelegateStoreDisplayNames.Add(Throwaway);
                break;
            default:
                settings.BystanderStoreDisplayNames.Add(Throwaway);
                break;
        }

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(settings));
        Assert.Contains("throwaway data file '" + Throwaway + "' in '", ex.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the write allowlist

    [Fact]
    public void OnATestGuest_TheThrowawayMayTakeADraftAndADelete_AndNothingElse()
    {
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(Guest());

        Assert.True(allowlist.IsThrowaway(Throwaway));
        Assert.Equal(new[] { Throwaway }, allowlist.ThrowawayStores);
        foreach (StoreWriteKind kind in Enum.GetValues<StoreWriteKind>())
        {
            bool granted = kind is StoreWriteKind.Draft or StoreWriteKind.Delete;
            Assert.Equal(granted, allowlist.IsAllowed(Throwaway, kind));
            if (!granted)
            {
                InvalidOperationException refused = Assert.Throws<InvalidOperationException>(() => allowlist.Assert(Throwaway, kind, "unit"));
                Assert.Contains("throwaway data file, granted draft+delete only", refused.Message, StringComparison.Ordinal);
            }
        }

        // It is no identity account: the identity tests never draft in it, even when handed it.
        Assert.Equal(new[] { Identity }, allowlist.IdentityAccountsAmong(new[] { Hub, Identity, Throwaway, Bystander }));
    }

    [Fact]
    public void OnTheReadOnlyWorkstation_TheThrowawayIsRefusedLikeEveryOtherStore()
    {
        // Q74 must not be weakened by the new grant: a Production profile - or none, which reads as
        // Production - refuses every write, the throwaway included, even if its settings named one.
        LiveTestSettings workstation = Guest();
        workstation.MachineProfile = LiveMachineProfile.Production;
        workstation.ProbeTerm = "term";
        workstation.SubjectOnlyProbe = new SubjectOnlyProbeSettings
        {
            StoreDisplayName = Hub,
            FolderPath = "Inbox",
            SubjectTerm = "term",
            SenderFragment = "term",
        };
        LiveTestSettings.Validate(workstation);

        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(workstation);

        Assert.True(allowlist.RefusesEveryWrite);
        Assert.True(allowlist.IsThrowaway(Throwaway));
        foreach (StoreWriteKind kind in Enum.GetValues<StoreWriteKind>())
        {
            Assert.False(allowlist.IsAllowed(Throwaway, kind));
            InvalidOperationException refused = Assert.Throws<InvalidOperationException>(() => allowlist.Assert(Throwaway, kind, "unit"));
            Assert.Contains(LiveWriteAccess.ReadOnlyMachine, refused.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AThrowawayThatIsAlsoAnotherTier_RefusesToBuild()
    {
        Assert.Throws<ArgumentException>(() => new StoreWriteAllowlist(Hub, throwawayStores: new[] { Hub }));
        Assert.Throws<ArgumentException>(() => new StoreWriteAllowlist(Hub, bystanderStores: new[] { Throwaway }, throwawayStores: new[] { Throwaway }));
        Assert.Throws<ArgumentException>(() => new StoreWriteAllowlist(Hub, knownReadOnlyStores: new[] { Throwaway }, throwawayStores: new[] { Throwaway }));
        Assert.Throws<ArgumentException>(() => new StoreWriteAllowlist(Hub, identityDraftStores: new[] { Throwaway }, throwawayStores: new[] { Throwaway }));
        Assert.Throws<ArgumentException>(() => StoreWriteAllowlist.RefusingEveryWrite(
            "unit reason", Hub, bystanderStores: new[] { Throwaway }, throwawayStores: new[] { Throwaway }));
    }

    [Fact]
    public void TheThrowawayIsNotWatched_SoTheTripwireAndTheArtifactSweepAreUnchanged()
    {
        // Decided on the maintainer's behalf (Q96 (iv)): the store holds nothing but this proof's
        // artifacts and is recreated before the next run, and watching it would fail the run over
        // the Drafts folder the proof exists to make. The census and the sweep plan stay exactly
        // what they were without it; the proof asserts its own zero-artifact end.
        LiveTestSettings with = Guest();
        LiveTestSettings without = Guest();
        without.ThrowawayStoreDisplayName = null;

        Assert.Equal(LiveStoreCountTripwire.WatchedStores(without), LiveStoreCountTripwire.WatchedStores(with));
        Assert.DoesNotContain(Throwaway, LiveStoreCountTripwire.WatchedStores(with), StringComparer.OrdinalIgnoreCase);
        Assert.Equal(
            ArtifactSweepPolicy.Assess(without).Steps.Select(s => s.Store + ":" + s.Action),
            ArtifactSweepPolicy.Assess(with).Steps.Select(s => s.Store + ":" + s.Action));
        Assert.True(TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(with), LiveStoreWriteGuard.Build(with), with.BystanderStoreDisplayNames).Usable);
    }

    // ------------------------------------------------------------------ the proof's judges

    [Fact]
    public void ThePrecondition_PassesOnlyWhenTheLookupProvesDraftsAbsent()
    {
        Assert.Null(ThrowawayStoreProof.DraftsAbsentBefore(Throwaway, "DefaultFolderAbsent"));

        string present = ThrowawayStoreProof.DraftsAbsentBefore(Throwaway, null)!;
        Assert.Contains("already HAS a Drafts folder", present, StringComparison.Ordinal);
        Assert.Contains(ThrowawayStoreProof.ResetScript, present, StringComparison.Ordinal);

        string unmounted = ThrowawayStoreProof.DraftsAbsentBefore(Throwaway, "StoreNotFound")!;
        Assert.Contains("is not mounted", unmounted, StringComparison.Ordinal);
        Assert.Contains(ThrowawayStoreProof.ResetScript, unmounted, StringComparison.Ordinal);

        // Unreadable is never read as absent - the rule SpecialFolders.Resolve itself keeps.
        foreach (string unknown in new[] { "DefaultFolderUnreadable", "COMException 0x80004005" })
        {
            string refused = ThrowawayStoreProof.DraftsAbsentBefore(Throwaway, unknown)!;
            Assert.Contains("could not be established", refused, StringComparison.Ordinal);
            Assert.Contains(unknown, refused, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TheReply_PassesOnlyWhenItReportsExactlyTheDraftsFolderItWasFiledIn()
    {
        string label = SpecialFolders.CreatedFolderLabel(Throwaway, "Drafts");
        Assert.Equal(Throwaway + "/Drafts", label);
        Assert.Null(ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { label }, Throwaway, "Drafts"));
        Assert.Null(ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { label }, "THROWAWAY@example.test", "Drafts"));

        // THE DEFECT Q85 exists for, on a real store: the folder made, the result silent.
        string silent = ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, null, Throwaway, "Drafts")!;
        Assert.Contains("reports creating nothing", silent, StringComparison.Ordinal);
        Assert.Contains("reports creating nothing", ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, Array.Empty<string>(), Throwaway, "Drafts")!, StringComparison.Ordinal);

        // Filed elsewhere - the hub's Drafts, where the first save lands - exercises nothing.
        Assert.Contains("not in the throwaway data file", ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { label }, Hub, "Drafts")!, StringComparison.Ordinal);
        Assert.Contains("names no folder", ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { label }, Throwaway, null)!, StringComparison.Ordinal);

        // Anything but exactly the one folder is a wrong report.
        Assert.Contains("where exactly", ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { label, Throwaway + "/Junk Email" }, Throwaway, "Drafts")!, StringComparison.Ordinal);
        Assert.Contains("where exactly", ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { Hub + "/Drafts" }, Throwaway, "Drafts")!, StringComparison.Ordinal);
        Assert.Contains("where exactly", ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(Throwaway, new[] { Throwaway + "/drafts" }, Throwaway, "Drafts")!, StringComparison.Ordinal);
    }

    [Fact]
    public void AfterTheCreation_TheProductsOwnLookupMustSeeTheFolder()
    {
        Assert.Null(ThrowawayStoreProof.DraftsVisibleAfter(Throwaway, null, "Drafts", "Drafts"));

        string blind = ThrowawayStoreProof.DraftsVisibleAfter(Throwaway, "DefaultFolderAbsent", null, "Drafts")!;
        Assert.Contains("does not see it", blind, StringComparison.Ordinal);
        Assert.Contains("discard_draft and update_draft", blind, StringComparison.Ordinal);
        Assert.Contains("Q85", blind, StringComparison.Ordinal);

        Assert.Contains("resolves the Drafts folder", ThrowawayStoreProof.DraftsVisibleAfter(Throwaway, null, "Other", "Drafts")!, StringComparison.Ordinal);
    }

    [Fact]
    public void TheDiscard_ReportsACreatedDeletedItemsExactlyWhenThereWasNone()
    {
        Assert.Null(ThrowawayStoreProof.DiscardReported(Throwaway, deletedItemsBefore: true, discarded: true, createdFolders: null));
        Assert.Contains("nothing made true", ThrowawayStoreProof.DiscardReported(Throwaway, true, true, new[] { Throwaway + "/Deleted Items" })!, StringComparison.Ordinal);

        Assert.Null(ThrowawayStoreProof.DiscardReported(Throwaway, deletedItemsBefore: false, discarded: true, createdFolders: new[] { Throwaway + "/Deleted Items" }));
        Assert.Contains("must report the one it created", ThrowawayStoreProof.DiscardReported(Throwaway, false, true, null)!, StringComparison.Ordinal);
        Assert.Contains("must report the one it created", ThrowawayStoreProof.DiscardReported(Throwaway, false, true, new[] { Hub + "/Deleted Items" })!, StringComparison.Ordinal);

        Assert.Contains("did not report the draft discarded", ThrowawayStoreProof.DiscardReported(Throwaway, true, false, null)!, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ Q96 question 2 (c): the top-level comparison

    [Fact]
    public void TheTopLevelComparison_PassesOnlyWhenWhatAppearedIsExactlyWhatWasReported()
    {
        string[] before = { "Deleted Items", "Deleted Items/old", "Search Folders" };
        string[] afterReply = { "Deleted Items", "Deleted Items/old", "Search Folders", "Drafts" };
        string drafts = SpecialFolders.CreatedFolderLabel(Throwaway, "Drafts");

        // Exactly what appeared was reported - as created, or as having appeared - and a call
        // that changed nothing reported nothing. Folder names compare as Outlook compares them.
        Assert.Null(ThrowawayStoreProof.TopLevelChangeReported(Throwaway, "reply_draft", before, afterReply, new[] { drafts }, null));
        Assert.Null(ThrowawayStoreProof.TopLevelChangeReported(Throwaway, "reply_draft", before, afterReply, null, new[] { drafts }));
        Assert.Null(ThrowawayStoreProof.TopLevelChangeReported(Throwaway, "discard_draft", afterReply, afterReply, null, null));
        Assert.Null(ThrowawayStoreProof.TopLevelChangeReported(
            Throwaway, "reply_draft", before, new[] { "deleted items", "Deleted Items/old", "search folders", "Drafts" }, new[] { drafts.ToUpperInvariant() }, null));

        // Below the top is not the top: the product's comparison is of the top level, and so is this.
        Assert.Null(ThrowawayStoreProof.TopLevelChangeReported(
            Throwaway, "reply_draft", before, afterReply.Append("Deleted Items/new").ToArray(), new[] { drafts }, null));

        // THE MEASUREMENT: a folder appeared beside the one the success path named, and nothing said so.
        string unreported = ThrowawayStoreProof.TopLevelChangeReported(
            Throwaway, "reply_draft", before, afterReply.Append("Junk Email").ToArray(), new[] { drafts }, null)!;
        Assert.Contains(Throwaway + "/Junk Email", unreported, StringComparison.Ordinal);
        Assert.Contains("Q96 question 2", unreported, StringComparison.Ordinal);
        Assert.Contains("on success too", unreported, StringComparison.Ordinal);
        Assert.Contains("reported none of them", ThrowawayStoreProof.TopLevelChangeReported(
            Throwaway, "reply_draft", before, afterReply, null, null)!, StringComparison.Ordinal);

        // A report nothing made true, and a folder that went missing, both fail.
        Assert.Contains("a claim nothing made true", ThrowawayStoreProof.TopLevelChangeReported(
            Throwaway, "discard_draft", afterReply, afterReply, new[] { Throwaway + "/Deleted Items" }, null)!, StringComparison.Ordinal);
        Assert.Contains("vanished", ThrowawayStoreProof.TopLevelChangeReported(
            Throwaway, "discard_draft", afterReply, new[] { "Deleted Items", "Search Folders" }, null, null)!, StringComparison.Ordinal);
    }

    [Fact]
    public void TheTopLevelRecord_SaysEverySide()
    {
        string line = ThrowawayStoreProof.DescribeTopLevelChange(
            Throwaway, "reply_draft", new[] { "Deleted Items", "Deleted Items/old" }, new[] { "Deleted Items", "Deleted Items/old", "Drafts" },
            new[] { Throwaway + "/Drafts" }, null);

        Assert.Contains("around reply_draft", line, StringComparison.Ordinal);
        Assert.Contains("before [Deleted Items]", line, StringComparison.Ordinal);
        Assert.Contains("after [Deleted Items, Drafts]", line, StringComparison.Ordinal);
        Assert.Contains("appeared [Drafts]; vanished []", line, StringComparison.Ordinal);
        Assert.Contains("reported created [" + Throwaway + "/Drafts], appeared []", line, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ Q96 question 3: where the Drafts folder is registered

    [Fact]
    public void TheDesignationRecord_SaysWhatEachPlaceNames()
    {
        const string DraftsId = "00000000AABBCCDD0011223344556677";
        byte[] draftsBytes = Convert.FromHexString(DraftsId);
        byte[] otherBytes = Convert.FromHexString("00000000FFEEDDCC0011223344556677");

        Assert.Equal("THE CREATED DRAFTS FOLDER", ThrowawayStoreProof.DescribeDesignation(
            new DesignationRead("store object", PropertyReadStatus.Found, draftsBytes), DraftsId));
        Assert.Equal("THE CREATED DRAFTS FOLDER", ThrowawayStoreProof.DescribeDesignation(
            new DesignationRead("store object", PropertyReadStatus.Found, draftsBytes), DraftsId.ToLowerInvariant()));
        Assert.StartsWith("another folder (00000000FFEEDDCC", ThrowawayStoreProof.DescribeDesignation(
            new DesignationRead("top folder", PropertyReadStatus.Found, otherBytes), DraftsId), StringComparison.Ordinal);
        Assert.StartsWith("another folder", ThrowawayStoreProof.DescribeDesignation(
            new DesignationRead("top folder", PropertyReadStatus.Found, draftsBytes), null), StringComparison.Ordinal);
        Assert.Equal("not set", ThrowawayStoreProof.DescribeDesignation(new DesignationRead("top folder", PropertyReadStatus.NotFound, null), DraftsId));
        Assert.Equal("could not be read", ThrowawayStoreProof.DescribeDesignation(new DesignationRead("Inbox", PropertyReadStatus.Failed, null), DraftsId));
        Assert.Equal("no such folder in this store", ThrowawayStoreProof.DescribeDesignation(new DesignationRead("Inbox", null, null), DraftsId));
        Assert.Equal("set, naming no folder", ThrowawayStoreProof.DescribeDesignation(
            new DesignationRead("store object", PropertyReadStatus.Found, Array.Empty<byte>()), DraftsId));
        Assert.Equal("set to a value that is no entry id", ThrowawayStoreProof.DescribeDesignation(
            new DesignationRead("store object", PropertyReadStatus.Found, 42), DraftsId));

        DesignationRead[] reads =
        {
            new DesignationRead("store object", PropertyReadStatus.Found, draftsBytes),
            new DesignationRead("top folder", PropertyReadStatus.NotFound, null),
            new DesignationRead("Inbox", null, null),
        };
        string line = ThrowawayStoreProof.DescribeDraftsDesignation(Throwaway, new DraftsDesignationReading(DraftsId, reads));
        Assert.Contains("store object = THE CREATED DRAFTS FOLDER; top folder = not set; Inbox = no such folder in this store", line, StringComparison.Ordinal);
        Assert.DoesNotContain("could not be read in that store", line, StringComparison.Ordinal);
        Assert.Contains("could not be read in that store", ThrowawayStoreProof.DescribeDraftsDesignation(
            Throwaway, new DraftsDesignationReading(null, reads)), StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the live test still asks

    [Fact]
    public void TheLiveProof_ObtainsItsStoreThroughTheCoverageGuard_AndAsksEveryJudge()
    {
        string source = LiveSource();

        Assert.Contains("LivePopulationCoverage.Require(", source, StringComparison.Ordinal);
        Assert.Contains("ThrowawayStoreProof.Population", source, StringComparison.Ordinal);
        foreach (string judge in new[]
        {
            "ThrowawayStoreProof.DraftsAbsentBefore(",
            "ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(",
            "ThrowawayStoreProof.DraftsVisibleAfter(",
            "ThrowawayStoreProof.DiscardReported(",
        })
        {
            Assert.Contains("Require(" + judge, source, StringComparison.Ordinal);
        }

        // The precondition is asked BEFORE the first write, and the writes in the order the proof needs.
        int precondition = source.IndexOf("Require(ThrowawayStoreProof.DraftsAbsentBefore(", StringComparison.Ordinal);
        int seed = source.IndexOf("LiveOutlookTestMailer.SaveTaggedPostInDeletedItems(", StringComparison.Ordinal);
        int reply = source.IndexOf("Service.ReplyDraft(", StringComparison.Ordinal);
        int discard = source.IndexOf("Service.DiscardDraft(", StringComparison.Ordinal);
        Assert.True(precondition > 0 && precondition < seed && seed < reply && reply < discard, "the proof's steps are out of order");
    }

    [Fact]
    public void TheLiveProof_MeasuresTheTopLevelAroundBothCalls_RecordsTheDesignation_AndJudgesTheComparisonLast()
    {
        // Q96 questions 2 (c) and 3: measured around each call, recorded before the first verdict
        // that could end the run, and the comparison judged only after the cleanup and the
        // zero-artifact proof - so a run that fails on it has answered everything else.
        string source = LiveSource();

        int seed = source.IndexOf("LiveOutlookTestMailer.SaveTaggedPostInDeletedItems(", StringComparison.Ordinal);
        int beforeReply = source.IndexOf("FolderPathsOf(store)", StringComparison.Ordinal);
        int reply = source.IndexOf("Service.ReplyDraft(", StringComparison.Ordinal);
        int afterReply = source.IndexOf("FolderPathsOf(store)", reply, StringComparison.Ordinal);
        int designation = source.IndexOf("RecordDraftsDesignation(store, reply.EntryId);", StringComparison.Ordinal);
        int firstVerdict = source.IndexOf("Require(ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(", StringComparison.Ordinal);
        int discard = source.IndexOf("Service.DiscardDraft(", StringComparison.Ordinal);
        int afterDiscard = source.IndexOf("FolderPathsOf(store)", discard, StringComparison.Ordinal);
        int nothingLeft = source.IndexOf("AssertNothingLeft(store, sourceEntryId, reply, discarded);", StringComparison.Ordinal);
        int replyJudged = source.IndexOf("Require(replyComparison);", StringComparison.Ordinal);
        int discardJudged = source.IndexOf("Require(discardComparison);", StringComparison.Ordinal);

        Assert.True(seed > 0 && seed < beforeReply && beforeReply < reply, "the top level must be listed before the reply");
        Assert.True(reply < afterReply && afterReply < firstVerdict, "the top level must be listed after the reply, before any verdict");
        Assert.True(reply < designation && designation < firstVerdict, "the designation must be recorded before any verdict");
        Assert.True(discard < afterDiscard, "the top level must be listed after the discard");
        Assert.True(nothingLeft > 0 && nothingLeft < replyJudged && nothingLeft < discardJudged, "the comparison is judged last");
        Assert.Equal(2, CountOf(source, "ThrowawayStoreProof.TopLevelChangeReported("));
        Assert.Contains("LiveFolderProbe.FolderTree(Service, store)", source, StringComparison.Ordinal);
        Assert.Contains("LiveOutlookTestMailer.ReadDraftsDesignations(store, draftEntryId)", source, StringComparison.Ordinal);
    }

    private static int CountOf(string text, string needle)
    {
        int count = 0;
        for (int at = text.IndexOf(needle, StringComparison.Ordinal); at >= 0; at = text.IndexOf(needle, at + needle.Length, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    [Fact]
    public void TheLiveProof_WritesOnlyBehindTheGuard_ThroughTheProductOrTheTestedHelpers_AndAlwaysCleansUp()
    {
        string source = LiveSource();

        // Every product write asks the guard first (mailbox-safety rules 1 and 3).
        Assert.True(
            source.IndexOf("LiveStoreWriteGuard.Writable(store, StoreWriteKind.Draft, \"reply_draft\")", StringComparison.Ordinal)
                < source.IndexOf("Service.ReplyDraft(", StringComparison.Ordinal),
            "reply_draft is not guarded");
        Assert.True(
            source.IndexOf("LiveStoreWriteGuard.Writable(store, StoreWriteKind.Delete, \"discard_draft\")", StringComparison.Ordinal)
                < source.IndexOf("Service.DiscardDraft(", StringComparison.Ordinal),
            "discard_draft is not guarded");

        // No COM of its own, and above all not the lookup that CREATES folders.
        foreach (string forbidden in new[] { ".GetDefaultFolder(", "dynamic ", "Outlook.Application", ".Delete()", ".Move(", "PermanentlyDelete" })
        {
            Assert.DoesNotContain(forbidden, source, StringComparison.Ordinal);
        }

        // The cleanup runs whatever happens, and only then is the zero-artifact end asserted.
        int finallyAt = source.IndexOf("finally", StringComparison.Ordinal);
        Assert.True(finallyAt > 0 && source.IndexOf("CleanUp(store, reply, discarded);", finallyAt, StringComparison.Ordinal) > finallyAt);
        Assert.Contains("LiveOutlookTestMailer.DeleteTaggedArtifacts(store, Marker, ThrowawayFolders)", source, StringComparison.Ordinal);
        Assert.Contains("CountTaggedArtifactsAfterPurgingStragglers(", source, StringComparison.Ordinal);
        Assert.Contains("TryGetMailInfo(entryId, storeId, out _)", source, StringComparison.Ordinal);
    }

    [Fact]
    public void TheLiveProof_IsALiveTest_ThatNamesItsCapability_AndClaimsNoReadOnlyTrait()
    {
        // It writes, so it must never carry Writes=Nothing - the workstation filter selects on that.
        MethodInfo test = typeof(LiveCreatedFolderTests).GetMethods()
            .Single(m => m.GetCustomAttributes(typeof(FactAttribute), false).Length > 0);
        List<(string Key, string Value)> traits = test.GetCustomAttributesData()
            .Where(a => a.AttributeType == typeof(TraitAttribute))
            .Select(a => ((string)a.ConstructorArguments[0].Value!, (string)a.ConstructorArguments[1].Value!))
            .ToList();
        Assert.Contains(("Requires", "MultipleStores"), traits);
        Assert.DoesNotContain(traits, t => t.Key == LiveRunFilters.WritesTrait);

        CollectionAttribute collection = Assert.Single(typeof(LiveCreatedFolderTests).GetCustomAttributes<CollectionAttribute>());
        Assert.Contains(typeof(LiveCreatedFolderTests).GetCustomAttributesData(), a =>
            a.AttributeType == typeof(TraitAttribute)
            && (string)a.ConstructorArguments[0].Value! == "Category"
            && (string)a.ConstructorArguments[1].Value! == "Live");
        Assert.NotNull(collection);
    }

    private static string LiveSource()
    {
        string testProjectDir = typeof(LiveCreatedFolderTests).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .First(a => a.Key == "TestProjectDir").Value!;
        string path = Path.Combine(testProjectDir, "T2", "LiveCreatedFolderTests.cs");
        Assert.True(File.Exists(path), "the live proof's source is missing: " + path);

        // Whole-line comments out, so a sentence about a call is not read as the call.
        return string.Join(
            "\n",
            File.ReadAllLines(path).Where(line => !line.TrimStart().StartsWith("//", StringComparison.Ordinal)
                && !line.TrimStart().StartsWith("///", StringComparison.Ordinal)));
    }
}
