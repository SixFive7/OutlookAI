using System.Text;
using System.Text.Json;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what outlook_health SAYS about how each store was tied to the index (Q92/Q99):
/// <c>index.perStore[].matchedBy / matchedInput / indexStore / matchNote</c> and
/// <c>index.storesNotInProfile</c>. Built from the real matcher over one profile holding every
/// kind of store the rule tells apart, so each wire value is pinned against the case that
/// produces it - and against the neighbouring case that must not.
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

    private static StoreIndexIdentity PstStore(string displayName, string path)
        => new(displayName, 3, StoreHash.Candidates(Convert.ToHexString(Pst(path)), false, false, null, null, null), filePath: path);

    private static StoreScopeInfo IndexRoot(string ownName, uint hash)
        => StoreScopeInfo.FromStorePrefix(Root + ownName + "($" + hash.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")")!;

    // One profile: a PST indexed under its own (different) name, a PST the index lacks, an
    // Exchange store tied by its signature hash, an Exchange store the hash does not decide, a
    // delegate - and, in the index, another profile's store.
    private static readonly StoreIndexIdentity IdentityPst = PstStore("identity@vm.invalid", @"C:\OutlookAI-Tier\identity.pst");
    private static readonly StoreIndexIdentity MissingPst = PstStore("archive", @"C:\OutlookAI-Tier\archive.pst");
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

        StoreStaleness identity = MailService.DescribeStoreMatch(map.Stores[0], null);
        Assert.Equal(("identity@vm.invalid", "storeHash", "entryId", "Outlook Data File($b25ac20a)"),
            (identity.Store, identity.MatchedBy, identity.MatchedInput, identity.IndexStore));

        // The PST the index lacks says so - not "displayName", although a store of its name is indexed.
        StoreStaleness missing = MailService.DescribeStoreMatch(map.Stores[1], null);
        Assert.Equal(("none", (string?)null, (string?)null), (missing.MatchedBy, missing.MatchedInput, missing.IndexStore));

        StoreStaleness alice = MailService.DescribeStoreMatch(map.Stores[2], null);
        Assert.Equal(("storeHash", "mappingSignature", "alice@example.com($" + AliceRoot.StoreHash!.Value.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")"),
            (alice.MatchedBy, alice.MatchedInput, alice.IndexStore));

        StoreStaleness jan = MailService.DescribeStoreMatch(map.Stores[4], null);
        Assert.Equal(("delegateFolder", (string?)null), (jan.MatchedBy, jan.IndexStore));
    }

    [Fact]
    public void AStoreTheHashDidNotDecide_SaysWhatTheNameRuleFound()
    {
        StoreIndexMatch bob = Map().Stores[3];
        Assert.Equal(StoreIndexMatchKind.NameRule, bob.Kind);

        StoreStaleness found = MailService.DescribeStoreMatch(bob, BobRoot);
        Assert.Equal(("displayName", (string?)null, "bob@example.com($13579bdf)"), (found.MatchedBy, found.MatchedInput, found.IndexStore));

        // Control: nothing found by name either.
        StoreStaleness notFound = MailService.DescribeStoreMatch(bob, null);
        Assert.Equal(("none", (string?)null), (notFound.MatchedBy, notFound.IndexStore));

        // And a name-rule root is never reported for a store the hash DID decide.
        StoreStaleness identity = MailService.DescribeStoreMatch(Map().Stores[0], BobRoot);
        Assert.Equal("Outlook Data File($b25ac20a)", identity.IndexStore);
    }

    [Fact]
    public void ARefusedHash_IsExplainedInTheRow()
    {
        StoreScopeInfo first = IndexRoot("old name", StoreHash.Compute(Pst(@"C:\OutlookAI-Tier\identity.pst")));
        StoreIndexMatch contested = Assert.Single(StoreIndexMatcher.Match(new[] { IdentityPst }, new[] { first, IdentityRoot }).Stores);

        StoreStaleness row = MailService.DescribeStoreMatch(contested, null);

        Assert.Equal("none", row.MatchedBy);
        Assert.Contains("more than one index store carries this store's hash", row.MatchNote, StringComparison.Ordinal);
        Assert.Contains("Outlook Data File($b25ac20a)", row.MatchNote, StringComparison.Ordinal);
    }

    [Fact]
    public void StoresNotInProfile_AreTheUnclaimedOnes_LessWhatTheNameRuleTied()
    {
        StoreIndexMap map = Map();

        // Unclaimed by hash: bob's root and the other profile's 'archive'.
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

        Assert.Contains("\"matchedBy\":\"storeHash\"", json, StringComparison.Ordinal);
        Assert.Contains("\"matchedInput\":\"entryId\"", json, StringComparison.Ordinal);
        Assert.Contains("\"indexStore\":\"Outlook Data File($b25ac20a)\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("matchNote", json, StringComparison.Ordinal);
    }
}
