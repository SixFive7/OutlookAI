using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// A store the profile has and the local index holds no mail for must be NAMED on every
/// unscoped search, not only on the one path that happened to have Outlook's store list in
/// hand (gap A1, residue).
/// <para>
/// WHAT WAS ALREADY CLOSED, and why this looked finished. 79c1827 gave the sweep one window
/// per store, raised <c>no_index_frontier</c>, and named the offending stores in
/// <c>sweep.storesWithoutIndex</c>. On the measured all-PST machine that fired correctly,
/// because there the PROFILE-wide frontier probe itself came back empty and the flag was set
/// before anything else ran.
/// </para>
/// <para>
/// THE RESIDUE, which that shape hides. On a MIXED profile - one indexed Exchange store plus
/// an unindexed data file, the ordinary "archive PST" shape - the profile-wide frontier is
/// measured fine, and the per-store loop only walks the index's own CATALOG, which has never
/// heard of the data file. The only thing that noticed it was a pass over the store names the
/// SWEEP RESULT carried, which runs after the sweep. So on all three paths where the sweep
/// does not run, the store was invisible:
/// </para>
/// <list type="number">
/// <item><description>
/// <b>The sweep was not needed.</b> A <c>before</c> bound older than the fallback window ends
/// before the sweep would start, so the search answered from the index alone - and called
/// itself <c>freshness: "live"</c> with <c>degraded</c> absent, over a profile where the one
/// store old mail actually lives in contributed nothing. That is the exact sentence
/// 79c1827's own message says it fixed, still true one profile shape over, and it is the
/// worst of the three: a <c>before</c>-bounded search is how an agent looks for OLD mail.
/// </description></item>
/// <item><description>
/// <b>The sweep was refused</b> (a recipient filter, or attachment-only): reported as
/// <c>index-only</c> and degraded, but nothing said which store had no index behind it.
/// </description></item>
/// <item><description>
/// <b>The sweep failed</b> (Outlook timed out, or COM threw): same.
/// </description></item>
/// </list>
/// <para>
/// FreshMerge already states the rule this violates - the code "survives a sweep that never
/// ran", because "the sweep could not run" and "the index has nothing here" are independent
/// facts and an answer missing both tiers must say both. It survived only when the profile
/// probe found nothing ANYWHERE.
/// </para>
/// <para>
/// Driven through the real <see cref="MailService"/> with a stand-in index client whose
/// catalog holds one of the profile's two stores, and a stand-in Outlook session. No mailbox
/// and no Windows Search index are touched.
/// </para>
/// </summary>
[Trait("TimeZone", "Machine")]
public sealed class UnindexedStoreReportingTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    /// <summary>An ordinary Exchange store, present in both the profile and the index.</summary>
    private const string IndexedStore = "alice@example.com";

    private const string IndexedPrefix = "mapi16://" + Sid + "/" + IndexedStore + "($deadbeef)";

    /// <summary>The archive the profile has mounted and Windows Search has never opened.</summary>
    private const string PstStore = "Archive Old.pst";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] ProfileStores =
    {
        new ComStoreDetail(IndexedStore, "store-alice", 0, true),
        new ComStoreDetail(PstStore, "store-pst", 3, null),
    };

    // ============================================== (1) the sweep that was not needed

    [Fact]
    public void ASearchBoundedPastTheFallbackWindow_NamesTheStoreNeitherTierCovers()
    {
        // "Find that invoice from January." The index answers for the Exchange store; the
        // PST - where a 2019 archive's mail actually is - is in neither tier, because the
        // index has nothing for it and the sweep window ends 7 days back. The answer used to
        // be freshness:"live" with degraded absent over exactly that.
        using MailService service = Service(OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(OldMailRequest());

        Assert.True(outcome.Sweep!.NotNeeded);
        Assert.True(outcome.Sweep.IndexFrontierMissing);
        Assert.Equal(new[] { PstStore }, outcome.Sweep.StoresWithoutIndex);
        Assert.Contains(FreshMerge.GapNoIndexFrontier, outcome.Sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.True(outcome.Degraded);
    }

    [Fact]
    public void ThatSearch_SaysWhichStore_InProseAsWellAsInFields()
    {
        // The code and the sentence are two renderings of one decision, so a caller that
        // reads advice and a caller that branches on fields learn the same thing.
        using MailService service = Service(OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(OldMailRequest());

        Assert.Contains(
            outcome.Advice!,
            a => a.Contains(PstStore, StringComparison.Ordinal)
                && a.Contains("NEITHER tier", StringComparison.Ordinal));
    }

    // ================================================ (2) the sweep that was refused

    [Fact]
    public void ASweepRefusedByAnAttachmentOnlyFilter_StillNamesIt()
    {
        // Two independent holes: the sweep never opens an attachment, AND the index holds no
        // mail for one store. The answer has to say both - it is missing both tiers there.
        using MailService service = Service(OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            AttachmentHitsOnly = true,
            Top = 25,
            SnippetChars = 0,
        });

        Assert.Equal(FreshMerge.AttachmentContentNotSweepable, outcome.Sweep!.Error);
        Assert.Equal(FreshMerge.FreshnessIndexOnly, outcome.Freshness);
        Assert.True(outcome.Sweep.IndexFrontierMissing);
        Assert.Equal(new[] { PstStore }, outcome.Sweep.StoresWithoutIndex);
        Assert.Contains(FreshMerge.GapNoIndexFrontier, outcome.Sweep.CoverageGaps!);
    }

    // ================================================== (3) the sweep that could not run

    [Fact]
    public void ASweepThatFailed_StillNamesTheStoreTheIndexHoldsNothingFor()
    {
        // Outlook answered for the store list and then threw on the sweep itself - a timeout
        // or a COM fault mid-walk. "The sweep failed" does not make "the index has nothing
        // for this store" go away, and the remedy differs: retry versus exhaustive:true.
        using MailService service = Service(OnlyAliceIndexed(), FailingSweep);

        SearchOutcome outcome = service.Search(Request());

        Assert.NotNull(outcome.Sweep!.Error);
        Assert.Equal(FreshMerge.FreshnessIndexOnly, outcome.Freshness);
        Assert.True(outcome.Sweep.IndexFrontierMissing);
        Assert.Equal(new[] { PstStore }, outcome.Sweep.StoresWithoutIndex);
    }

    // ======================================================== no false alarms

    [Fact]
    public void AFullyIndexedProfile_RaisesNothing_OnAnyOfThosePaths()
    {
        // The flag that matters most is the one that does not cry wolf: every store here has
        // a frontier of its own, so no window fell back and nothing is missing.
        using MailService service = Service(EverythingIndexed());

        SearchOutcome notNeeded = service.Search(OldMailRequest());

        Assert.True(notNeeded.Sweep!.NotNeeded);
        Assert.Null(notNeeded.Sweep.IndexFrontierMissing);
        Assert.Null(notNeeded.Sweep.StoresWithoutIndex);
        Assert.Equal(FreshMerge.FreshnessLive, notNeeded.Freshness);
        Assert.Null(notNeeded.Degraded);
    }

    [Fact]
    public void OutlookUnreachable_MakesNoClaimEitherWay()
    {
        // Silence here means "not established", never "indexed". With no store list there is
        // no evidence, and inventing a missing store from an absent list would be the same
        // defect pointing the other way. The answer is still honestly index-only.
        using MailService service = new MailService(new DeadGateway(), null, OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(Request());

        Assert.Equal(FreshMerge.FreshnessIndexOnly, outcome.Freshness);
        Assert.True(outcome.Degraded);
        Assert.Null(outcome.Sweep!.IndexFrontierMissing);
        Assert.Null(outcome.Sweep.StoresWithoutIndex);
    }

    [Fact]
    public void AStoreScopedSearch_IsSettledByItsOwnFrontierProbe_AndAsksOutlookForNoStoreList()
    {
        // A scoped search already knows: its frontier probe is scoped to the one store in
        // scope, so a null there IS the answer. Asking Outlook for a profile-wide list would
        // be a round trip that cannot change anything.
        using MailService service = new MailService(
            new DirectGateway(StandInSession.Create(ProfileStores, Sweep, null, out StandInSession session)), null, OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            StoreId = TestFolders.StoreId(PstStore),
            BeforeUtc = Frontier.AddDays(-90),
            Top = 25,
            SnippetChars = 0,
        });

        Assert.True(outcome.Sweep!.IndexFrontierMissing);
        Assert.Equal(new[] { PstStore }, outcome.Sweep.StoresWithoutIndex);

        // One read only: the store-scope resolution that A4 added - which is also where the store's ID
        // is looked up since Q133. Not a second one for a list this path cannot use.
        Assert.Equal(1, session.StoreListReads);
    }

    // =================================================== the path that already worked

    [Fact]
    public void TheOrdinarySweepPath_StillNamesIt_FromTheSweepsOwnStoreList()
    {
        // The regression guard on what 79c1827 shipped: when the sweep DOES run, its own
        // per-store counters are the store list, and no extra round trip is spent.
        using MailService service = Service(OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(Request());

        Assert.True(outcome.Sweep!.Performed);
        Assert.True(outcome.Sweep.IndexFrontierMissing);
        Assert.Equal(new[] { PstStore }, outcome.Sweep.StoresWithoutIndex);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
    }

    // =================================================================== fixtures

    private static SearchRequest Request()
    {
        return new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 };
    }

    /// <summary>
    /// A search for mail older than the fallback window: its <c>before</c> bound ends before
    /// any sweep would start, which is what makes the sweep "not needed".
    /// </summary>
    private static SearchRequest OldMailRequest()
    {
        return new SearchRequest
        {
            Query = "test",
            BeforeUtc = Frontier.AddDays(-90),
            Top = 25,
            SnippetChars = 0,
        };
    }

    private static MailService Service(StubIndexClient index, Func<string?, ComSweepResult>? sweep = null)
    {
        return new MailService(
            new DirectGateway(StandInSession.Create(ProfileStores, sweep ?? Sweep)), null, index);
    }

    /// <summary>The mixed profile: the Exchange store indexed, the data file not.</summary>
    private static StubIndexClient OnlyAliceIndexed()
    {
        return new StubIndexClient(new[] { IndexedPrefix }, Frontier);
    }

    /// <summary>Both stores in the catalog, both with a frontier - nothing to report.</summary>
    private static StubIndexClient EverythingIndexed()
    {
        return new StubIndexClient(new[] { IndexedPrefix, PstPrefix }, Frontier);
    }

    private const string PstPrefix = "mapi16://" + Sid + "/" + PstStore + "($feedface)";

    /// <summary>A sweep that reaches both stores and finds one mail in the data file.</summary>
    private static ComSweepResult Sweep(string? onlyStore)
    {
        IReadOnlyList<string> stores = onlyStore == null
            ? new[] { IndexedStore, PstStore }
            : new[] { onlyStore };

        return new ComSweepResult(
            new[]
            {
                new ComMailBrief(
                    entryId: "AA1",
                    storeDisplayName: stores[stores.Count - 1],
                    storeId: "store-pst",
                    folderName: "Inbox",
                    folderKind: "inbox",
                    subject: "a test mail",
                    senderName: "Bob",
                    senderAddress: "bob@example.com",
                    receivedTime: Frontier,
                    isRead: true,
                    hasAttachments: false,
                    sizeBytes: 2048,
                    body: "test body"),
            },
            foldersSwept: 4 * stores.Count,
            foldersSkipped: 0,
            sweptFolders: stores.Select(s => s + "/Inbox").ToList(),
            perStore: stores
                .Select(s => new ComStoreSweepCounters(s, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0))
                .ToList());
    }

    /// <summary>Outlook answered the store list and then threw on the walk itself.</summary>
    private static ComSweepResult FailingSweep(string? onlyStore)
    {
        throw new TimeoutException("The freshness sweep exceeded its 30s budget.");
    }
}
