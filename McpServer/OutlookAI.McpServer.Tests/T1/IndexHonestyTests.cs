using System;
using System.Collections.Generic;

using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: outlook_health must be honest about the index, because it is the report an operator
/// opens to DISCOVER that searches are falling back to a fixed window - and it denied the
/// condition existed.
/// <para>
/// (A2) A NULL FRONTIER WAS READ AS "NO LAG". An index holding zero mail rows reports a null
/// newest-indexed timestamp and a null age, which is byte-identical to an index nobody asked
/// about. The report keyed on the age alone, found nothing to complain about, and printed
/// "Index is current; searches run at index speed" with <c>status: "ok"</c> over an index
/// that could not serve a single search.
/// </para>
/// <para>
/// (A3) TWO LISTS THAT NEVER MET. <c>index.perStore</c> came from the index-derived catalog
/// and <c>outlook.stores</c> from COM, both already in hand, and nothing compared them - so a
/// store Outlook has mounted and the index has never seen appeared in neither, which is
/// exactly the store whose searches fall back to a fixed window.
/// </para>
/// <para>
/// WHAT IS PROVEN HERE AND WHAT IS NOT. Both decisions are pure functions - one over the
/// probe result, one over the two lists plus an injected per-store probe - and every branch
/// is pinned below. What needs the testbed is the PROBE itself answering correctly against a
/// real Windows Search index (delegate mailboxes indexed under an owner subtree, an '@' store
/// the discovery sample missed): that is exercised by list_accounts' own live coverage, which
/// reports the same verdict through the same probe.
/// </para>
/// </summary>
public sealed class IndexHonestyTests
{
    private static readonly DateTime FrontierUtc = new(2026, 08, 18, 07, 20, 09, DateTimeKind.Utc);

    // ------------------------------------------------ (A2) three states behind one null

    [Fact]
    public void AnIndexWithNoMailAtAll_IsNotCurrent_AndIsNotUnavailableEither()
    {
        // The state that had no name. Reachable provider, zero mail rows: not a lag, not a
        // quiet mailbox, not an outage - the index tier simply cannot answer anything.
        Assert.Equal(MailService.IndexCurrency.NoMailAtAll, MailService.ClassifyIndexCurrency("OleDb", null));
    }

    [Fact]
    public void AMeasuredFrontier_IsMeasured_HoweverOldItIs()
    {
        Assert.Equal(MailService.IndexCurrency.Measured, MailService.ClassifyIndexCurrency("OleDb", FrontierUtc));
        Assert.Equal(
            MailService.IndexCurrency.Measured,
            MailService.ClassifyIndexCurrency("AdodbCom", FrontierUtc.AddYears(-3)));
    }

    [Fact]
    public void AnUnreachableIndex_EstablishesNothing_AndIsNeverCalledEmpty()
    {
        // Same null frontier as the empty index, different fact: saying "the index holds no
        // mail" over an index nobody could read would be a claim about data nobody saw.
        Assert.Equal(
            MailService.IndexCurrency.Unavailable,
            MailService.ClassifyIndexCurrency("unavailable: COMException", null));
        Assert.Equal(
            MailService.IndexCurrency.Unavailable,
            MailService.ClassifyIndexCurrency("unavailable: InvalidOperationException", FrontierUtc));
    }

    // ------------------------------------------------ (A3) the comparison nobody made

    [Fact]
    public void AStoreOutlookHasAndTheIndexDoesNot_IsNamed()
    {
        IReadOnlyList<string> missing = MailService.StoresMissingFromIndex(
            new[] { "jori@huisman.io", "Archive 2019.pst" },
            new[] { "jori@huisman.io" },
            store => store == "Archive 2019.pst" ? false : true);

        Assert.Equal(new[] { "Archive 2019.pst" }, missing);
    }

    [Fact]
    public void AStoreTheProbeCouldNotSettle_IsNotReportedEitherWay()
    {
        // The budget ran out, or the index stopped answering. "Not established" must never
        // become "not indexed": this report drives a problem and a degraded status.
        Assert.Empty(MailService.StoresMissingFromIndex(
            new[] { "Archive 2019.pst" }, Array.Empty<string>(), _ => null));
    }

    [Fact]
    public void AbsenceFromTheCatalog_IsNeverEvidenceOnItsOwn()
    {
        // The catalog comes from an unordered 2000-row sample that misses small stores, and a
        // delegate mailbox is indexed under its OWNER's subtree so it never appears under its
        // own name at all. A name comparison alone would report every shared mailbox on a
        // delegate-heavy profile as unindexed - a flag that cries wolf is worse than no flag.
        Assert.Empty(MailService.StoresMissingFromIndex(
            new[] { "Shared Mailbox", "Tiny.pst" },
            Array.Empty<string>(),
            _ => true));
    }

    [Fact]
    public void KnownStores_AreMatchedCaseInsensitively_AndAreNeverProbedTwice()
    {
        List<string> probed = new();
        IReadOnlyList<string> missing = MailService.StoresMissingFromIndex(
            new[] { "JORI@HUISMAN.IO", "Archive.pst", "archive.pst" },
            new[] { "jori@huisman.io" },
            store =>
            {
                probed.Add(store);
                return false;
            });

        Assert.Equal(new[] { "Archive.pst" }, missing);
        Assert.Equal(new[] { "Archive.pst" }, probed);
    }

    [Fact]
    public void NoComStoreList_MakesNoClaimAtAll()
    {
        // Outlook was not running or did not answer: half a list would invent missing stores.
        Assert.Empty(MailService.StoresMissingFromIndex(null, new[] { "jori@huisman.io" }, _ => false));
        Assert.Empty(MailService.StoresMissingFromIndex(Array.Empty<string>(), null, _ => false));
    }
}
