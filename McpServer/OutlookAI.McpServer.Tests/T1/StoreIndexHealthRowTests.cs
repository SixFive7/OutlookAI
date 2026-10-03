using System.Text;
using System.Text.Json;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what outlook_health SAYS about how each store was tied to the index (Q92/Q99, Q113):
/// <c>index.perStore[].matchedBy / matchedInput / indexStore / matchNote</c> and
/// <c>index.storesNotInProfile</c>. Built from the real matcher over one profile holding every
/// kind of store the rule tells apart, so each wire value is pinned against the case that
/// produces it - and against the neighbouring case that must not. <c>nameAndHash</c>, <c>none</c>
/// and <c>ambiguous</c> are the rule (Q113 (a)); <c>storeHash</c> and <c>displayName</c> are the
/// one open exception, Exchange (Q113 (b)).
/// </summary>
public sealed class StoreIndexHealthRowTests
{
    private const string Root = "mapi16://{S-1-5-21-1-2-3-1000}/";

    private static byte[] Pst(string path)
    {
        byte[] header = Convert.FromHexString(
            "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000" + "0000");
        return header.Concat(Encoding.Unicode.GetBytes(path + "\0")).ToArray();
    }

    private static StoreIndexIdentity PstStore(string displayName, string path, string ownName)
        => new(displayName, 3, StoreHash.Candidates(Convert.ToHexString(Pst(path)), false, false, null, null, null),
            filePath: path, ownName: ownName);

    private static StoreScopeInfo IndexRoot(string ownName, uint hash)
        => StoreScopeInfo.FromStorePrefix(Root + ownName + "($" + hash.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")")!;

    // One profile: a PST indexed under its own (different) name, a PST the index lacks, an
    // Exchange store tied by its signature hash, an Exchange store the hash does not decide, a
    // delegate - and, in the index, another profile's store of the missing PST's name.
    private static readonly StoreIndexIdentity IdentityPst = PstStore("identity@vm.invalid", @"C:\OutlookAI-Tier\identity.pst", "Outlook Data File");
    private static readonly StoreIndexIdentity MissingPst = PstStore("archive", @"C:\OutlookAI-Tier\archive.pst", "archive");
    private static readonly StoreIndexIdentity HashedExchange = new("alice@example.com", 0,
        StoreHash.Candidates("00112233", true, false, "0A0B0C0D", null, null), filePath: @"C:\Users\a\alice.ost");
    private static readonly StoreIndexIdentity UndecidedExchange = new("bob@example.com", 4,
        StoreHash.Candidates("44556677", true, false, "0E0F1011", null, null), filePath: @"C:\Users\a\bob.ost");
    private static readonly StoreIndexIdentity Delegate = new("Jan", 1, Array.Empty<StoreHashCandidate>());

    private static readonly StoreScopeInfo IdentityRoot = IndexRoot("Outlook Data File", StoreHash.Compute(Pst(@"C:\OutlookAI-Tier\identity.pst")));
    private static readonly StoreScopeInfo AliceRoot = IndexRoot("alice@example.com", StoreHash.Compute(new byte[] { 0x0A, 0x0B, 0x0C, 0x0D }));
    private static readonly StoreScopeInfo BobRoot = IndexRoot("bob@example.com", 0x13579BDF);
    private static readonly StoreScopeInfo OtherProfiles = IndexRoot("archive", StoreHash.Compute(Pst(@"C:\Other\archive.pst")));

    private static StoreIndexMap Map() => StoreIndexMatcher.Match(
        new[] { IdentityPst, MissingPst, HashedExchange, UndecidedExchange, Delegate },
        new[] { IdentityRoot, AliceRoot, BobRoot, OtherProfiles });

    [Fact]
    public void EachStoreSaysHowItWasTied()
    {
        StoreIndexMap map = Map();

        // The rule: name and hash, both the store's own.
        StoreStaleness identity = MailService.DescribeStoreMatch(map.Stores[0], null);
        Assert.Equal(("identity@vm.invalid", "nameAndHash", "entryId", "Outlook Data File($b25ac20a)"),
            (identity.Store, identity.MatchedBy, identity.MatchedInput, identity.IndexStore));
        Assert.Null(identity.MatchNote);

        // The PST the index lacks says so - not "displayName", although a store of its name is
        // indexed - and its note names the near miss, so nobody has to wonder why.
        StoreStaleness missing = MailService.DescribeStoreMatch(map.Stores[1], null);
        Assert.Equal(("none", (string?)null, (string?)null), (missing.MatchedBy, missing.MatchedInput, missing.IndexStore));
        Assert.Contains(OtherProfiles.StoreSegment, missing.MatchNote, StringComparison.Ordinal);
        Assert.Contains("carries another hash", missing.MatchNote, StringComparison.Ordinal);

        // The open exception: an Exchange store by its hash alone.
        StoreStaleness alice = MailService.DescribeStoreMatch(map.Stores[2], null);
        Assert.Equal(("storeHash", "mappingSignature", "alice@example.com($" + AliceRoot.StoreHash!.Value.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")"),
            (alice.MatchedBy, alice.MatchedInput, alice.IndexStore));

        StoreStaleness jan = MailService.DescribeStoreMatch(map.Stores[4], null);
        Assert.Equal(("delegateFolder", (string?)null), (jan.MatchedBy, jan.IndexStore));
    }

    [Fact]
    public void AnExchangeStoreTheHashDidNotDecide_SaysWhatTheNameRuleFound()
    {
        StoreIndexMatch bob = Map().Stores[3];
        Assert.Equal(StoreIndexMatchKind.NameRule, bob.Kind);

        StoreStaleness found = MailService.DescribeStoreMatch(bob, BobRoot);
        Assert.Equal(("displayName", (string?)null, "bob@example.com($13579bdf)"), (found.MatchedBy, found.MatchedInput, found.IndexStore));

        // Control: nothing found by name either.
        StoreStaleness notFound = MailService.DescribeStoreMatch(bob, null);
        Assert.Equal(("none", (string?)null), (notFound.MatchedBy, notFound.IndexStore));

        // And a name-rule root is never reported for a store the rule DID decide.
        StoreStaleness identity = MailService.DescribeStoreMatch(Map().Stores[0], BobRoot);
        Assert.Equal("Outlook Data File($b25ac20a)", identity.IndexStore);
        StoreStaleness missing = MailService.DescribeStoreMatch(Map().Stores[1], OtherProfiles);
        Assert.Equal(("none", (string?)null), (missing.MatchedBy, missing.IndexStore));
    }

    [Fact]
    public void TwoStoresTheIndexCannotTellApart_SayAmbiguous_AndNameEachOther()
    {
        // Two PSTs sharing both the name and the hash of one index store (the same entry ID, as a
        // stand-in for a 32-bit collision under one name).
        string path = @"C:\OutlookAI-Tier\shared.pst";
        StoreScopeInfo shared = IndexRoot("Archive", StoreHash.Compute(Pst(path)));
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("Archive A", path, "Archive"), PstStore("Archive B", path, "Archive") }, new[] { shared });

        StoreStaleness a = MailService.DescribeStoreMatch(map.Stores[0], null);
        Assert.Equal(("ambiguous", (string?)null, shared.StoreSegment), (a.MatchedBy, a.MatchedInput, a.IndexStore));
        Assert.Contains("'Archive B'", a.MatchNote, StringComparison.Ordinal);

        StoreStaleness b = MailService.DescribeStoreMatch(map.Stores[1], null);
        Assert.Equal("ambiguous", b.MatchedBy);
        Assert.Contains("'Archive A'", b.MatchNote, StringComparison.Ordinal);

        // Not "not in this profile": the index store is this profile's, only not attributable.
        Assert.Empty(MailService.DescribeStoresNotInProfile(map, Array.Empty<string>()));

        // The problem outlook_health raises names both, and says what happens to searches.
        string problem = MailService.DescribeUnmatchableStores(new[] { "Archive A", "Archive B" });
        Assert.Contains("Archive A, Archive B", problem, StringComparison.Ordinal);
        Assert.Contains("refused", problem, StringComparison.Ordinal);
        Assert.Contains("leave that index store's mail out", problem, StringComparison.Ordinal);
    }

    [Fact]
    public void ARenamesLeftoverRoot_IsNotTheStores_AndTheRowExplains()
    {
        // The store's hash under its OLD name only: not tied - the pre-Q113 rule took that root -
        // and the note says what the leftover is.
        StoreScopeInfo leftover = IndexRoot("old name", StoreHash.Compute(Pst(@"C:\OutlookAI-Tier\identity.pst")));
        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(new[] { IdentityPst }, new[] { leftover }).Stores);

        StoreStaleness row = MailService.DescribeStoreMatch(match, null);

        Assert.Equal("none", row.MatchedBy);
        Assert.Contains("hash only under another name (old name($b25ac20a))", row.MatchNote, StringComparison.Ordinal);

        // Control: with its current root listed too, it is tied there and nothing is noted.
        StoreIndexMatch tied = Assert.Single(StoreIndexMatcher.Match(new[] { IdentityPst }, new[] { leftover, IdentityRoot }).Stores);
        StoreStaleness tiedRow = MailService.DescribeStoreMatch(tied, null);
        Assert.Equal(("nameAndHash", "Outlook Data File($b25ac20a)"), (tiedRow.MatchedBy, tiedRow.IndexStore));
        Assert.Null(tiedRow.MatchNote);
    }

    [Fact]
    public void StoresNotInProfile_AreTheUnclaimedOnes_LessWhatTheNameRuleTied()
    {
        StoreIndexMap map = Map();

        // Unclaimed: bob's root (his hash did not decide) and the other profile's 'archive'.
        Assert.Equal(new[] { "bob@example.com($13579bdf)", OtherProfiles.StoreSegment },
            MailService.DescribeStoresNotInProfile(map, Array.Empty<string>()));

        // Bob was found under his root by name, so it is his - only the other profile's remains.
        Assert.Equal(new[] { OtherProfiles.StoreSegment },
            MailService.DescribeStoresNotInProfile(map, new[] { BobRoot.StorePrefix }));
    }

    [Fact]
    public void TheMatchFields_TravelUnderTheirWireNames()
    {
        // The tools' own serializer settings (OutlookTools: web defaults, nulls omitted).
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        };
        StoreStaleness row = MailService.DescribeStoreMatch(Map().Stores[0], null);
        string json = JsonSerializer.Serialize(row, options);

        Assert.Contains("\"matchedBy\":\"nameAndHash\"", json, StringComparison.Ordinal);
        Assert.Contains("\"matchedInput\":\"entryId\"", json, StringComparison.Ordinal);
        Assert.Contains("\"indexStore\":\"Outlook Data File($b25ac20a)\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("matchNote", json, StringComparison.Ordinal);
    }
}
