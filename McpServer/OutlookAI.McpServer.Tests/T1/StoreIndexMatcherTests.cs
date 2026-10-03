using System.Text;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="StoreIndexMatcher"/> - which index store is which Outlook store (Q92/Q99).
/// Every PST shape here is one MEASURED on the indexed test guest (Docs/live-tier-on-the-vm.md
/// section 8 item 24), driven through the pure matcher with index roots built the way the index
/// spells them: a store whose own name differs from its profile name, two stores sharing a name,
/// a copy of a PST at another path, another profile's store, a store the index does not hold.
/// The stores whose hash input is NOT measured - Exchange, an IMAP or Outlook.com .ost - are
/// pinned to the other half of the rule: matched when the hash says so, otherwise handed back to
/// the name rule the product used before, never declared unindexed. Each test pairs the case with
/// its control, so the rule is pinned from both sides.
/// </summary>
public sealed class StoreIndexMatcherTests
{
    private const string Root = "mapi16://{S-1-5-21-1-2-3-1000}/";

    private static byte[] Pst(string path)
    {
        byte[] header = Convert.FromHexString(
            "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000" + "0000");
        return header.Concat(Encoding.Unicode.GetBytes(path + "\0")).ToArray();
    }

    private static string Hex(byte[] bytes) => Convert.ToHexString(bytes);

    /// <summary>A PST as Outlook reports it: profile name, StoreID, not Exchange, a .pst path.</summary>
    private static StoreIndexIdentity PstStore(string displayName, string path)
        => new(displayName, 3, StoreHash.Candidates(Hex(Pst(path)), false, false, null, null, null), filePath: path);

    /// <summary>An index root as the DIRECTORY listing returns it: the store's own name and its hash.</summary>
    private static StoreScopeInfo IndexRoot(string ownName, uint hash)
        => StoreScopeInfo.FromStorePrefix(Root + ownName + "($" + hash.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")")!;

    private static StoreScopeInfo IndexRootOf(string ownName, string path) => IndexRoot(ownName, StoreHash.Compute(Pst(path)));

    [Fact]
    public void AStoreWhoseOwnNameDiffersFromItsProfileName_IsFoundByHash()
    {
        // The measured identity store: Store.DisplayName 'identity@vm.invalid', the index's
        // 'Outlook Data File'. The name lookup could never find it; the hash cannot miss it.
        StoreScopeInfo root = IndexRootOf("Outlook Data File", @"C:\OutlookAI-Tier\identity.pst");
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("identity@vm.invalid", @"C:\OutlookAI-Tier\identity.pst") }, new[] { root });

        StoreIndexMatch match = Assert.Single(map.Stores);
        Assert.Equal(StoreIndexMatchKind.StoreHash, match.Kind);
        Assert.Equal(StoreHashInput.EntryId, match.Input);
        Assert.Same(root, match.Root);
        Assert.Empty(map.Unclaimed);
        Assert.Equal("identity@vm.invalid", map.StoreNameForPrefix(root.StorePrefix));
    }

    [Fact]
    public void APstTheIndexDoesNotHold_IsNotIndexed_EvenWhenAnotherStoreOfThatNameIs()
    {
        // The guest's index held another PROFILE's 'Outlook Data File'. A name lookup scoped this
        // profile's store to that one - answering with the other profile's mail. By hash, the
        // store is simply not in the index, and the other root is left alone.
        StoreScopeInfo otherProfiles = IndexRootOf("Outlook Data File", @"C:\OutlookAI-Tier\Outlook Data File - CorpusProfile.pst");
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("Outlook Data File", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { otherProfiles });

        StoreIndexMatch match = Assert.Single(map.Stores);
        Assert.Equal(StoreIndexMatchKind.None, match.Kind);
        Assert.Null(match.Root);
        Assert.Null(match.Note);
        Assert.Same(otherProfiles, Assert.Single(map.Unclaimed));

        // Control: the same store, indexed - found, and the other root still unclaimed.
        StoreScopeInfo own = IndexRootOf("Outlook Data File", @"C:\OutlookAI-Tier\Outlook.pst");
        map = StoreIndexMatcher.Match(
            new[] { PstStore("Outlook Data File", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { otherProfiles, own });
        Assert.Same(own, map.Stores[0].Root);
        Assert.Same(otherProfiles, Assert.Single(map.Unclaimed));
    }

    [Fact]
    public void TwoStoresWithOneName_EachFindTheirOwnRoot()
    {
        // Measured: a PST and its copy at another path, both named 'q98scratch@vm.invalid', two
        // roots of that name. Only the hash tells them apart.
        StoreScopeInfo original = IndexRootOf("q98scratch@vm.invalid", @"C:\OutlookAI-Tier\q98-scratch.pst");
        StoreScopeInfo copy = IndexRootOf("q98scratch@vm.invalid", @"C:\OutlookAI-Tier\moved\q98-scratch.pst");
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[]
            {
                PstStore("q98scratch@vm.invalid", @"C:\OutlookAI-Tier\moved\q98-scratch.pst"),
                PstStore("q98scratch@vm.invalid", @"C:\OutlookAI-Tier\q98-scratch.pst"),
            },
            new[] { original, copy });

        Assert.Same(copy, map.Stores[0].Root);
        Assert.Same(original, map.Stores[1].Root);
        Assert.All(map.Stores, m => Assert.Equal(StoreIndexMatchKind.StoreHash, m.Kind));

        // Two stores share the name, so the name picks out neither.
        Assert.Null(map.ForStore("q98scratch@vm.invalid"));
    }

    [Fact]
    public void AnIndexNameWithPercentEncoding_DoesNotMatterToTheHash()
    {
        // Measured: 'q99 50% off*?x' is spelled 'q99 50%25 off%2A%3Fx($5159380d)' in the index.
        StoreScopeInfo root = StoreScopeInfo.FromStorePrefix(Root + "q99 50%25 off%2A%3Fx($5159380d)")!;
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("q99 50% off*?x", @"C:\OutlookAI-Tier\q99-special.pst") }, new[] { root });

        Assert.Equal(StoreIndexMatchKind.StoreHash, map.Stores[0].Kind);
    }

    [Fact]
    public void AnExchangeStoreWithNoHashMatch_IsLeftToTheNameRule_NeverCalledUnindexed()
    {
        // The cached-Exchange inputs are documented, not measured. A miss therefore proves
        // nothing, and the store goes back to the rule it had - the matcher itself never matches
        // by name, so the same-named root stays unclaimed for that rule to find.
        var exchange = new StoreIndexIdentity("alice@example.com", 0,
            StoreHash.Candidates("00112233", true, false, "0A0B0C0D", null, null), filePath: @"C:\Users\a\alice.ost");
        StoreScopeInfo byName = IndexRoot("alice@example.com", 0xDEADBEEF);
        StoreIndexMap map = StoreIndexMatcher.Match(new[] { exchange }, new[] { byName });

        Assert.Equal(StoreIndexMatchKind.NameRule, map.Stores[0].Kind);
        Assert.Null(map.Stores[0].Root);
        Assert.Same(byName, Assert.Single(map.Unclaimed));

        // Control: the same store with its signature hash in the index - tied by it.
        StoreHash.TryDecodeHex("0A0B0C0D", out byte[] signature);
        StoreScopeInfo own = IndexRoot("alice@example.com", StoreHash.Compute(signature));
        map = StoreIndexMatcher.Match(new[] { exchange }, new[] { byName, own });
        Assert.Equal(StoreIndexMatchKind.StoreHash, map.Stores[0].Kind);
        Assert.Equal(StoreHashInput.MappingSignature, map.Stores[0].Input);
        Assert.Same(own, map.Stores[0].Root);
    }

    [Fact]
    public void OnlyAPstFile_IsAuthoritative_AnImapOrOutlookComOstIsNot()
    {
        // An IMAP or Outlook.com store is "not Exchange" exactly like a PST, but it lives in an
        // .ost whose hash input nobody here has measured. A miss leaves it to the name rule; only
        // a .pst - the measured family - is declared unindexed by one.
        string path = @"C:\Users\a\AppData\Local\Microsoft\Outlook\bob@example.com.ost";
        var ost = new StoreIndexIdentity("bob@example.com", 3,
            StoreHash.Candidates(Hex(Pst(path)), false, false, null, null, null), filePath: path);
        StoreScopeInfo byName = IndexRoot("bob@example.com", 0x0BADF00D);

        Assert.False(ost.HashIsAuthoritative);
        Assert.Equal(StoreIndexMatchKind.NameRule, StoreIndexMatcher.Match(new[] { ost }, new[] { byName }).Stores[0].Kind);

        // Control: the same store as a .pst - authoritative, and not indexed.
        string pstPath = Path.ChangeExtension(path, ".PST");
        var pst = new StoreIndexIdentity("bob@example.com", 3,
            StoreHash.Candidates(Hex(Pst(pstPath)), false, false, null, null, null), filePath: pstPath);
        Assert.True(pst.HashIsAuthoritative);
        Assert.Equal(StoreIndexMatchKind.None, StoreIndexMatcher.Match(new[] { pst }, new[] { byName }).Stores[0].Kind);

        // And a .pst whose path would not read is not known to be one.
        var unknownPath = new StoreIndexIdentity("bob@example.com", 3,
            StoreHash.Candidates(Hex(Pst(pstPath)), false, false, null, null, null));
        Assert.False(unknownPath.HashIsAuthoritative);
    }

    [Fact]
    public void AnExchangeStoreWhoseSignatureHashIsInTheIndex_IsFoundByIt_WhateverItsName()
    {
        const string signature = "8899AABBCCDDEEFF0011223344556677";
        StoreHash.TryDecodeHex(signature, out byte[] sig);
        StoreScopeInfo root = IndexRoot("someone else's label", StoreHash.Compute(sig));
        var exchange = new StoreIndexIdentity("alice@example.com", 0,
            StoreHash.Candidates("00112233", true, false, null, signature, @"C:\Users\a\alice.ost"));

        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(new[] { exchange }, new[] { root }).Stores);
        Assert.Equal(StoreIndexMatchKind.StoreHash, match.Kind);
        Assert.Equal(StoreHashInput.ProfileMappingSignature, match.Input);
    }

    [Fact]
    public void AHashTwoStoresClaim_TiesNeither_AndSaysWhy()
    {
        // A 32-bit hash colliding between two of one user's stores is vanishingly rare - and the
        // one case where picking would answer with the wrong mail. Refused, never resolved, and
        // left to the name rule: the hash is present, so even a PST is not "not indexed".
        StoreScopeInfo root = IndexRootOf("x", @"C:\a.pst");
        StoreIndexIdentity a = PstStore("a", @"C:\a.pst");
        var b = new StoreIndexIdentity("b", 0, StoreHash.Candidates(Hex(Pst(@"C:\a.pst")), true, false, null, null, null));

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { a, b }, new[] { root });

        Assert.All(map.Stores, m =>
        {
            Assert.Equal(StoreIndexMatchKind.NameRule, m.Kind);
            Assert.Null(m.Root);
            Assert.Contains("another Outlook store", m.Note, StringComparison.Ordinal);
        });
        Assert.Empty(map.Unclaimed);

        // Control: either store alone is tied by the hash.
        Assert.Equal(StoreIndexMatchKind.StoreHash, StoreIndexMatcher.Match(new[] { a }, new[] { root }).Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.StoreHash, StoreIndexMatcher.Match(new[] { b }, new[] { root }).Stores[0].Kind);
    }

    [Fact]
    public void AStoreWhoseHashTwoIndexStoresCarry_IsTiedToNeither_AndLeftToTheNameRule()
    {
        StoreScopeInfo first = IndexRootOf("old name", @"C:\OutlookAI-Tier\Outlook.pst");
        StoreScopeInfo second = IndexRootOf("new name", @"C:\OutlookAI-Tier\Outlook.pst");

        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(
            new[] { PstStore("tier@vm.invalid", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { first, second }).Stores);

        Assert.Equal(StoreIndexMatchKind.NameRule, match.Kind);
        Assert.Null(match.Root);
        Assert.Contains("more than one index store", match.Note, StringComparison.Ordinal);

        // Control: one of the two alone ties it.
        Assert.Same(first, StoreIndexMatcher.Match(
            new[] { PstStore("tier@vm.invalid", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { first }).Stores[0].Root);
    }

    [Fact]
    public void ADelegateStore_IsNeverAStoreOfItsOwn()
    {
        var host = PstStore("host", @"C:\h.pst");
        var delegated = new StoreIndexIdentity("Jan", 1, StoreHash.Candidates(Hex(Pst(@"C:\h.pst")), true, false, null, null, null));
        StoreScopeInfo root = IndexRootOf("host", @"C:\h.pst");

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { host, delegated }, new[] { root });

        Assert.Equal(StoreIndexMatchKind.StoreHash, map.Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.Delegate, map.Stores[1].Kind);
        Assert.Null(map.Stores[1].Root);
    }

    [Fact]
    public void AStoreWithoutAReadableId_IsLeftToTheNameRule_EvenAPst()
    {
        // No id, no hash: nothing measured can be said about it either way.
        StoreScopeInfo root = IndexRoot("Archive", 0xABCDEF01);
        var noId = new StoreIndexIdentity("Archive", 3, Array.Empty<StoreHashCandidate>(), filePath: @"C:\Archive.pst");
        Assert.False(noId.HashIsAuthoritative);
        Assert.Equal(StoreIndexMatchKind.NameRule, StoreIndexMatcher.Match(new[] { noId }, new[] { root }).Stores[0].Kind);
    }

    [Fact]
    public void ARootWithoutAHash_IsNeverClaimed()
    {
        StoreScopeInfo noHash = StoreScopeInfo.FromStorePrefix(Root + "Legacy")!;
        Assert.Null(noHash.StoreHash);
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("Legacy", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { noHash });
        Assert.Equal(StoreIndexMatchKind.None, map.Stores[0].Kind);
        Assert.Same(noHash, Assert.Single(map.Unclaimed));
    }
}
