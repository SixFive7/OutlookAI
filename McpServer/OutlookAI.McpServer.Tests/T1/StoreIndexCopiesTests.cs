using System.Globalization;
using System.Text;
using System.Text.Json;

using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what the store map makes of an index store that carries a profile store's name while nothing ties it to the
/// store - the stale copy Windows Search keeps of an Outlook data file that was rebuilt (measured by stream A on the
/// Exchange guest, 2026-11-17: the old copy of the mailbox as a second root of the same name and folder paths beside
/// the live one), or another profile's store of the same name. Decided 2026-11-18, the TODO item's (a) and (b)
/// together: when another index store of that name IS tied to the store by its hash, that one is the live copy and an
/// unscoped search leaves the untied one out, saying so; when none is tied, nothing is left out and the copies are only
/// reported; outlook_health reports both. Every name, address and hash here is synthetic.
/// </summary>
public sealed class StoreIndexCopiesTests
{
    private const string Root = "mapi16://{S-1-5-21-1-2-3-1000}/";
    private const string Signature = "0A0B0C0D";

    private static byte[] Pst(string path)
    {
        byte[] header = Convert.FromHexString(
            "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000" + "0000");
        return header.Concat(Encoding.Unicode.GetBytes(path + "\0")).ToArray();
    }

    private static StoreIndexIdentity PstStore(string displayName, string path, string ownName)
        => new(displayName, 3, StoreHash.Candidates(Convert.ToHexString(Pst(path)), false, false, null, null, null),
            filePath: path, ownName: ownName);

    /// <summary>A cached Exchange mailbox whose mapping signature is <see cref="Signature"/>.</summary>
    private static StoreIndexIdentity Mailbox(string displayName = "alice@example.invalid", string signature = Signature)
        => new(displayName, 0, StoreHash.Candidates("00112233", true, false, signature, null, null), filePath: @"C:\Users\a\alice.ost");

    private static StoreScopeInfo IndexRoot(string name, uint hash)
        => StoreScopeInfo.FromStorePrefix(Root + name + "($" + hash.ToString("x", CultureInfo.InvariantCulture) + ")")!;

    private static readonly StoreScopeInfo LiveCopy = IndexRoot("alice@example.invalid", StoreHash.Compute(Convert.FromHexString(Signature)));
    private static readonly StoreScopeInfo OldCopy = IndexRoot("alice@example.invalid", 0x9B7B502);

    [Fact]
    public void TheRebuiltOstsOldCopy_IsLeftOut_WhenTheLiveOneIsTiedByTheStoresHash()
    {
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { Mailbox() }, new[] { OldCopy, LiveCopy });

        Assert.Same(LiveCopy, map.Stores[0].Root);
        StoreIndexCopies copies = Assert.Single(map.Copies);
        Assert.Same(map.Stores[0], copies.Store);
        Assert.Same(LiveCopy, copies.Live);
        Assert.Same(OldCopy, Assert.Single(copies.Copies));
        Assert.True(copies.LeftOut);
        Assert.Null(copies.SharedNameWith);
        Assert.Single(map.CopiesLeftOut);
        Assert.Empty(map.CopiesSearched);

        // Control: the live copy alone - nothing to leave out, nothing to say.
        Assert.Empty(StoreIndexMatcher.Match(new[] { Mailbox() }, new[] { LiveCopy }).Copies);
    }

    [Fact]
    public void ADataFilesLeftoverAtAnOldPath_OrAnotherProfilesStoreOfItsName_IsLeftOutTheSameWay()
    {
        // A PST tied by name and hash; the index also holds a root of its own name under another hash - the file
        // at an earlier path, or another profile's store called the same (runbook section 8 item 24's decoy).
        StoreScopeInfo own = IndexRoot("Outlook Data File", StoreHash.Compute(Pst(@"C:\OutlookAI-Tier\Outlook.pst")));
        StoreScopeInfo other = IndexRoot("Outlook Data File", StoreHash.Compute(Pst(@"C:\Other\Outlook.pst")));
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("tier@example.invalid", @"C:\OutlookAI-Tier\Outlook.pst", "Outlook Data File") }, new[] { other, own });

        StoreIndexCopies copies = Assert.Single(map.Copies);
        Assert.Same(own, copies.Live);
        Assert.Same(other, Assert.Single(copies.Copies));
        Assert.True(copies.LeftOut);
    }

    [Fact]
    public void WhenNothingTiesTheStore_ItsCopiesAreSearchedAndReported_NeverLeftOut()
    {
        // An Exchange store the hash does not decide (the open exception, Q113 (b)): two index stores of its name,
        // nothing to say which is live.
        StoreScopeInfo a = IndexRoot("bob@example.invalid", 0x2468ACE0);
        StoreScopeInfo b = IndexRoot("bob@example.invalid", 0x13579BDF);
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { Mailbox("bob@example.invalid", "0E0F1011") }, new[] { a, b });

        Assert.Equal(StoreIndexMatchKind.NameRule, map.Stores[0].Kind);
        StoreIndexCopies copies = Assert.Single(map.Copies);
        Assert.Null(copies.Live);
        Assert.Equal(new[] { a, b }, copies.Copies);
        Assert.False(copies.LeftOut);
        Assert.Single(map.CopiesSearched);
        Assert.Empty(map.CopiesLeftOut);

        // Control: one index store of its name is the name rule's ordinary answer, not a copy.
        Assert.Empty(StoreIndexMatcher.Match(new[] { Mailbox("bob@example.invalid", "0E0F1011") }, new[] { a }).Copies);
    }

    [Fact]
    public void AnotherStoreTiedToNothingThatCarriesTheName_KeepsTheCopiesIn_NeverAGuess()
    {
        // The tied mailbox's untied namesake could be filed under the other root: it is searched and reported.
        StoreIndexIdentity namesake = Mailbox("alice@example.invalid", "0E0F1011");
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { Mailbox(), namesake }, new[] { OldCopy, LiveCopy });

        StoreIndexCopies copies = Assert.Single(map.Copies, c => c.Live != null);
        Assert.False(copies.LeftOut);
        Assert.Equal("alice@example.invalid", copies.SharedNameWith);

        // Control: a PST of that own name tied to nothing is not in the index (its rule is measured), so it keeps
        // nothing in.
        StoreIndexIdentity untiedPst = PstStore("p", @"C:\x\p.pst", "alice@example.invalid");
        StoreIndexMap withPst = StoreIndexMatcher.Match(new[] { Mailbox(), untiedPst }, new[] { OldCopy, LiveCopy });
        Assert.True(Assert.Single(withPst.Copies).LeftOut);
    }

    [Fact]
    public void ARootAnotherStoreIsTiedTo_IsNeverACopy()
    {
        // Two data files of one own name, each tied to its own root.
        StoreIndexIdentity first = PstStore("one", @"C:\a\one.pst", "Archive");
        StoreIndexIdentity second = PstStore("two", @"C:\b\two.pst", "Archive");
        StoreScopeInfo firstRoot = IndexRoot("Archive", StoreHash.Compute(Pst(@"C:\a\one.pst")));
        StoreScopeInfo secondRoot = IndexRoot("Archive", StoreHash.Compute(Pst(@"C:\b\two.pst")));

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { first, second }, new[] { firstRoot, secondRoot });

        Assert.Equal(StoreIndexMatchKind.NameAndHash, map.Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.NameAndHash, map.Stores[1].Kind);
        Assert.Empty(map.Copies);
    }

    // ============================================================ what search and outlook_health say

    [Fact]
    public void TheWordsLeftOut_NameTheStore_TheLiveCopyAndTheOldOne()
    {
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { Mailbox() }, new[] { OldCopy, LiveCopy });

        string search = MailService.DescribeStoreCopiesLeftOut(map.CopiesLeftOut, forHealth: false);
        Assert.StartsWith("Left out of this search:", search, StringComparison.Ordinal);
        Assert.Contains("'alice@example.invalid' is searched through " + LiveCopy.StoreSegment, search, StringComparison.Ordinal);
        Assert.Contains(OldCopy.StoreSegment + " left out", search, StringComparison.Ordinal);
        Assert.Contains("outlook_health lists them per store under storeCopies", search, StringComparison.Ordinal);

        string health = MailService.DescribeStoreCopiesLeftOut(map.CopiesLeftOut, forHealth: true);
        Assert.StartsWith("Unscoped searches leave out", health, StringComparison.Ordinal);
        Assert.DoesNotContain("outlook_health", health, StringComparison.Ordinal);
    }

    [Fact]
    public void TheWordsSearched_SayNothingTellsWhichIsLive_AndWhatResultsCanThenHold()
    {
        StoreScopeInfo a = IndexRoot("bob@example.invalid", 0x2468ACE0);
        StoreScopeInfo b = IndexRoot("bob@example.invalid", 0x13579BDF);
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { Mailbox("bob@example.invalid", "0E0F1011") }, new[] { a, b });

        string search = MailService.DescribeStoreCopiesSearched(map.CopiesSearched, forHealth: false);
        Assert.StartsWith("This search read every index store", search, StringComparison.Ordinal);
        Assert.Contains("'bob@example.invalid': " + a.StoreSegment + ", " + b.StoreSegment + ", none tied to it", search, StringComparison.Ordinal);
        Assert.Contains("mail moved or deleted since, or the same mail twice", search, StringComparison.Ordinal);

        string health = MailService.DescribeStoreCopiesSearched(map.CopiesSearched, forHealth: true);
        Assert.Contains("a search scoped to such a store is refused", health, StringComparison.Ordinal);
    }

    [Fact]
    public void TheFields_CapAtTheStoreListCap_AndTravelUnderTheirWireNames()
    {
        List<StoreScopeInfo> roots = Enumerable.Range(0, MailService.UnindexedStoreListCap + 3)
            .Select(i => IndexRoot("carol@example.invalid", (uint)(0x1000 + i)))
            .ToList();
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { Mailbox("carol@example.invalid", "0E0F1011") }, roots);

        IReadOnlyList<string>? segments = MailService.CopySegments(map.CopiesSearched);
        Assert.Equal(MailService.UnindexedStoreListCap, segments!.Count);
        Assert.Null(MailService.CopySegments(Array.Empty<StoreIndexCopies>()));

        string index = JsonSerializer.Serialize(
            new IndexTierInfo { StoreCopiesLeftOut = new[] { "x($1)" }, StoreCopiesSearched = new[] { "y($2)" } },
            new JsonSerializerOptions(JsonSerializerDefaults.Web));
        Assert.Contains("\"storeCopiesLeftOut\":[\"x($1)\"]", index, StringComparison.Ordinal);
        Assert.Contains("\"storeCopiesSearched\":[\"y($2)\"]", index, StringComparison.Ordinal);

        string row = JsonSerializer.Serialize(
            new StoreStaleness { Store = "s", StoreCopies = new[] { "x($1)" }, StoreCopiesLeftOut = true },
            new JsonSerializerOptions(JsonSerializerDefaults.Web));
        Assert.Contains("\"storeCopies\":[\"x($1)\"]", row, StringComparison.Ordinal);
        Assert.Contains("\"storeCopiesLeftOut\":true", row, StringComparison.Ordinal);
    }
}
