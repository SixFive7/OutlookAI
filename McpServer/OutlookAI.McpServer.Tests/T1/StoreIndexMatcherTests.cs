using System.Text;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="StoreIndexMatcher"/> - which index store is which Outlook store (Q92/Q99) - to
/// the maintainer's Q113 (a) decision of 2026-10-03: DETERMINISTIC MATCHING, NEVER A GUESS. A store
/// Outlook reports as not Exchange (a PST, an IMAP or Outlook.com .ost) is tied to the index store
/// whose NAME is its own name AND whose HASH is its hash; the index URL carries both, so where two
/// stores differ in either the match is certain, and where two share both, the store is refused
/// (<see cref="StoreIndexMatchKind.Ambiguous"/>) - never picked, and never handed to a name lookup.
/// An Exchange store keeps its pre-Q113 behaviour - THE ONE OPEN EXCEPTION, until Q113 (b) measures
/// the cached-Exchange input - and that is pinned here too, so removing it is a visible change.
/// <para>
/// Every PST shape is one MEASURED on the indexed test guest (Docs/live-tier-on-the-vm.md section 8
/// item 24), driven through the pure matcher with index roots built the way the index spells them.
/// Each test pairs the case with its control, so the rule is pinned from both sides.
/// </para>
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

    /// <summary>A PST as Outlook reports it: profile name, StoreID, not Exchange, a .pst path, and its own (root folder) name.</summary>
    private static StoreIndexIdentity PstStore(string displayName, string path, string? ownName = null)
        => new(displayName, 3, StoreHash.Candidates(Hex(Pst(path)), false, false, null, null, null),
            filePath: path, ownName: ownName ?? displayName);

    /// <summary>An index root as the DIRECTORY listing returns it: the store's own name and its hash.</summary>
    private static StoreScopeInfo IndexRoot(string ownName, uint hash)
        => StoreScopeInfo.FromStorePrefix(Root + ownName + "($" + hash.ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")")!;

    private static StoreScopeInfo IndexRootOf(string ownName, string path) => IndexRoot(ownName, StoreHash.Compute(Pst(path)));

    // ============================================================ the rule: name AND hash

    [Fact]
    public void NameAndHashBothMatch_TheStoreIsTiedToThatIndexStore_WhateverTheProfileCallsIt()
    {
        // The measured identity store: Store.DisplayName 'identity@vm.invalid', own name (root
        // folder, and the index's) 'Outlook Data File'. The match is on the OWN name.
        StoreScopeInfo root = IndexRootOf("Outlook Data File", @"C:\OutlookAI-Tier\identity.pst");
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("identity@vm.invalid", @"C:\OutlookAI-Tier\identity.pst", "Outlook Data File") }, new[] { root });

        StoreIndexMatch match = Assert.Single(map.Stores);
        Assert.Equal(StoreIndexMatchKind.NameAndHash, match.Kind);
        Assert.Equal(StoreHashInput.EntryId, match.Input);
        Assert.Same(root, match.Root);
        Assert.Null(match.Note);
        Assert.Empty(map.Unclaimed);
        Assert.Equal("identity@vm.invalid", map.StoreNameForPrefix(root.StorePrefix));

        // Control: matched against the PROFILE's name instead, the same root is not this store's.
        StoreIndexMatch byProfileName = Assert.Single(StoreIndexMatcher.Match(
            new[] { PstStore("identity@vm.invalid", @"C:\OutlookAI-Tier\identity.pst", "identity@vm.invalid") }, new[] { root }).Stores);
        Assert.Equal(StoreIndexMatchKind.None, byProfileName.Kind);
    }

    [Fact]
    public void TheHashMatchesButTheNameDiffers_TheIndexStoreIsNotTheStores_AndTheNoteSaysWhy()
    {
        // A rename: the index still holds the store's hash under its OLD name (measured: the old
        // root goes within seconds, the new one comes back under the new name). Same hash, other
        // name - not this store's index store as it is now.
        StoreScopeInfo oldName = IndexRootOf("old name", @"C:\OutlookAI-Tier\Outlook.pst");
        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(
            new[] { PstStore("tier@vm.invalid", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { oldName }).Stores);

        Assert.Equal(StoreIndexMatchKind.None, match.Kind);
        Assert.Null(match.Root);
        Assert.Contains("hash only under another name (old name($", match.Note, StringComparison.Ordinal);

        // Control: with the new root there too, the store is tied to it - the case the hash alone
        // could not decide (two index stores carried the hash, and the pre-Q113 rule refused both).
        StoreScopeInfo newName = IndexRootOf("tier@vm.invalid", @"C:\OutlookAI-Tier\Outlook.pst");
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("tier@vm.invalid", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { oldName, newName });
        Assert.Equal(StoreIndexMatchKind.NameAndHash, map.Stores[0].Kind);
        Assert.Same(newName, map.Stores[0].Root);
        Assert.Same(oldName, Assert.Single(map.Unclaimed));
    }

    [Fact]
    public void TheNameMatchesButTheHashDiffers_TheIndexStoreIsAnothers_AndIsNeverUsed()
    {
        // The guest's index held another PROFILE's 'Outlook Data File'. A name lookup scoped this
        // profile's store to that one - answering with the other profile's mail.
        StoreScopeInfo otherProfiles = IndexRootOf("Outlook Data File", @"C:\OutlookAI-Tier\Outlook Data File - CorpusProfile.pst");
        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(
            new[] { PstStore("Outlook Data File", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { otherProfiles }).Stores);

        Assert.Equal(StoreIndexMatchKind.None, match.Kind);
        Assert.Null(match.Root);
        Assert.Contains("the index store of its name (Outlook Data File($", match.Note, StringComparison.Ordinal);
        Assert.Contains("carries another hash", match.Note, StringComparison.Ordinal);

        // Control: the same store, indexed - tied to its own, and the other root still unclaimed.
        StoreScopeInfo own = IndexRootOf("Outlook Data File", @"C:\OutlookAI-Tier\Outlook.pst");
        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("Outlook Data File", @"C:\OutlookAI-Tier\Outlook.pst") }, new[] { otherProfiles, own });
        Assert.Equal(StoreIndexMatchKind.NameAndHash, map.Stores[0].Kind);
        Assert.Same(own, map.Stores[0].Root);
        Assert.Same(otherProfiles, Assert.Single(map.Unclaimed));
    }

    [Fact]
    public void TwoStoresSharingBothNameAndHash_AreRefused_NeitherIsPicked()
    {
        // The index itself cannot tell these apart: one index store carries the name and hash of
        // both. Picking one would answer with mail that may be the other's.
        string path = @"C:\OutlookAI-Tier\shared.pst";
        StoreScopeInfo shared = IndexRootOf("Archive", path);
        StoreIndexIdentity a = PstStore("Archive A", path, "Archive");
        StoreIndexIdentity b = PstStore("Archive B", path, "Archive");

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { a, b }, new[] { shared });

        Assert.All(map.Stores, m =>
        {
            Assert.Equal(StoreIndexMatchKind.Ambiguous, m.Kind);
            Assert.Null(m.Root);
            Assert.Same(shared, Assert.Single(m.Contested));
            Assert.Contains("cannot tell this store apart from", m.Note, StringComparison.Ordinal);
        });
        Assert.Equal(new[] { "Archive B" }, map.Stores[0].SharedWith);
        Assert.Equal(new[] { "Archive A" }, map.Stores[1].SharedWith);
        Assert.Equal(2, map.Unmatchable.Count);
        Assert.Null(map.StoreNameForPrefix(shared.StorePrefix));
        Assert.Empty(map.Unclaimed); // the index store is this profile's - only not attributable

        // Control: either store alone is tied by name and hash.
        Assert.Equal(StoreIndexMatchKind.NameAndHash, StoreIndexMatcher.Match(new[] { a }, new[] { shared }).Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.NameAndHash, StoreIndexMatcher.Match(new[] { b }, new[] { shared }).Stores[0].Kind);
    }

    [Fact]
    public void TwoStoresSharingTheHashButNotTheName_AreEachTiedToTheirOwn()
    {
        // The other half of "the URL carries both": a hash two stores share (a 32-bit collision)
        // no longer refuses both, because the name tells them apart.
        string path = @"C:\OutlookAI-Tier\collide.pst";
        StoreScopeInfo first = IndexRootOf("first", path);
        StoreScopeInfo second = IndexRootOf("second", path);

        StoreIndexMap map = StoreIndexMatcher.Match(
            new[] { PstStore("one", path, "first"), PstStore("two", path, "second") }, new[] { first, second });

        Assert.Same(first, map.Stores[0].Root);
        Assert.Same(second, map.Stores[1].Root);
        Assert.All(map.Stores, m => Assert.Equal(StoreIndexMatchKind.NameAndHash, m.Kind));
        Assert.Empty(map.Unmatchable);
    }

    [Fact]
    public void TwoStoresWithOneDisplayName_EachFindTheirOwnRoot_ButTheNamePicksNeither()
    {
        // Measured: a PST and its copy at another path, both named 'q98scratch@vm.invalid', two
        // roots of that name. The hash tells them apart; the display name cannot.
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
        Assert.All(map.Stores, m => Assert.Equal(StoreIndexMatchKind.NameAndHash, m.Kind));
        Assert.Null(map.ForStore("q98scratch@vm.invalid"));
        Assert.Equal(2, map.StoresNamed("q98scratch@vm.invalid").Count);
    }

    [Fact]
    public void NamesCompareAsTheIndexSpellsThem_PercentEncodingCaseAndEdgeSpace()
    {
        // Measured: 'q99 50% off*?x' is spelled 'q99 50%25 off%2A%3Fx($5159380d)' in the index.
        StoreScopeInfo root = StoreScopeInfo.FromStorePrefix(Root + "q99 50%25 off%2A%3Fx($5159380d)")!;
        var special = new StoreIndexIdentity("q99 50% off*?x", 3,
            new[] { new StoreHashCandidate(StoreHashInput.EntryId, 0x5159380d) }, filePath: @"C:\q99-special.pst", ownName: "q99 50% off*?x");
        Assert.Equal(StoreIndexMatchKind.NameAndHash, StoreIndexMatcher.Match(new[] { special }, new[] { root }).Stores[0].Kind);

        // Case and a trailing space do not count (the server treats index prefixes as
        // case-insensitive, and the URL parser trims a trailing space); any other letter does.
        var cased = new StoreIndexIdentity("x", 3,
            new[] { new StoreHashCandidate(StoreHashInput.EntryId, 0x5159380d) }, filePath: @"C:\x.pst", ownName: "Q99 50% OFF*?X ");
        Assert.Equal(StoreIndexMatchKind.NameAndHash, StoreIndexMatcher.Match(new[] { cased }, new[] { root }).Stores[0].Kind);
        var other = new StoreIndexIdentity("x", 3,
            new[] { new StoreHashCandidate(StoreHashInput.EntryId, 0x5159380d) }, filePath: @"C:\x.pst", ownName: "q99 50% off*?y");
        Assert.Equal(StoreIndexMatchKind.None, StoreIndexMatcher.Match(new[] { other }, new[] { root }).Stores[0].Kind);
    }

    [Fact]
    public void AStoreWithoutAReadableIdOrOwnName_IsTiedToNothing_NeverLeftToTheNameRule()
    {
        // No id: no hash. No own name: no name. Either way there is no name-and-hash to match, and
        // a store under the rule never falls back to its name alone (Q113 (a): "for PSTs, never").
        StoreScopeInfo root = IndexRoot("Archive", 0xABCDEF01);
        var noId = new StoreIndexIdentity("Archive", 3, Array.Empty<StoreHashCandidate>(), filePath: @"C:\Archive.pst", ownName: "Archive");
        StoreIndexMatch noIdMatch = StoreIndexMatcher.Match(new[] { noId }, new[] { root }).Stores[0];
        Assert.Equal(StoreIndexMatchKind.None, noIdMatch.Kind);
        Assert.Contains("store ID would not read", noIdMatch.Note, StringComparison.Ordinal);

        StoreIndexIdentity noName = PstStore("Archive", @"C:\Archive.pst", ownName: "");
        Assert.Null(noName.OwnName);
        StoreIndexMatch noNameMatch = StoreIndexMatcher.Match(new[] { noName }, new[] { IndexRootOf("Archive", @"C:\Archive.pst") }).Stores[0];
        Assert.Equal(StoreIndexMatchKind.None, noNameMatch.Kind);
        Assert.Contains("own name (its root folder's) would not read", noNameMatch.Note, StringComparison.Ordinal);
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

    // ============================================================ IMAP / Outlook.com .ost

    [Fact]
    public void AnImapOrOutlookComOst_IsMatchedByNameAndHash_LikeAPst_AndNeverByNameAlone()
    {
        // Outlook reports an IMAP or Outlook.com (EAS) store as olNotExchange, like a PST, in an
        // .ost. It is held to the same rule: tied when its own name and its entry-ID hash are both
        // in the index, tied to NOTHING otherwise - never handed to the name rule (D55's fallback,
        // superseded by Q113 (a)). Its input is documented, not measured, and its note says so.
        string path = @"C:\Users\a\AppData\Local\Microsoft\Outlook\bob@example.com.ost";
        var ost = new StoreIndexIdentity("bob@example.com", 3,
            StoreHash.Candidates(Hex(Pst(path)), false, false, null, null, null), filePath: path, ownName: "bob@example.com");
        Assert.False(ost.InExchangeException);
        Assert.False(ost.IsPstFile);

        StoreScopeInfo byNameOnly = IndexRoot("bob@example.com", 0x0BADF00D);
        StoreIndexMatch untied = StoreIndexMatcher.Match(new[] { ost }, new[] { byNameOnly }).Stores[0];
        Assert.Equal(StoreIndexMatchKind.None, untied.Kind);
        Assert.Contains("documented, not measured", untied.Note, StringComparison.Ordinal);

        // Control: its own name with its hash - tied.
        StoreScopeInfo own = IndexRootOf("bob@example.com", path);
        StoreIndexMatch tied = StoreIndexMatcher.Match(new[] { ost }, new[] { byNameOnly, own }).Stores[0];
        Assert.Equal(StoreIndexMatchKind.NameAndHash, tied.Kind);
        Assert.Same(own, tied.Root);

        // And the PST of the same name, missing from the index, says nothing about measurement.
        string pstPath = Path.ChangeExtension(path, ".PST");
        StoreIndexIdentity pst = PstStore("bob@example.com", pstPath);
        Assert.True(pst.IsPstFile);
        Assert.DoesNotContain("not measured", StoreIndexMatcher.Match(new[] { pst }, new[] { byNameOnly }).Stores[0].Note ?? string.Empty, StringComparison.Ordinal);
    }

    // ============================================================ THE ONE OPEN EXCEPTION: Exchange

    [Fact]
    public void TheExchangeException_IsEveryStoreOutlookDoesNotReportAsNotExchange_ExceptADelegate()
    {
        // The one switch that routes a store to the pre-Q113 behaviour. When Q113 (b) is measured
        // this becomes false for all of them, and this test is what changes.
        Assert.True(new StoreIndexIdentity("primary", 0, Array.Empty<StoreHashCandidate>()).InExchangeException);
        Assert.True(new StoreIndexIdentity("public", 2, Array.Empty<StoreHashCandidate>()).InExchangeException);
        Assert.True(new StoreIndexIdentity("additional", 4, Array.Empty<StoreHashCandidate>()).InExchangeException);
        Assert.True(new StoreIndexIdentity("type unreadable", null, Array.Empty<StoreHashCandidate>()).InExchangeException);
        Assert.False(new StoreIndexIdentity("delegate", 1, Array.Empty<StoreHashCandidate>()).InExchangeException);
        Assert.False(new StoreIndexIdentity("pst or ost", 3, Array.Empty<StoreHashCandidate>()).InExchangeException);
    }

    [Fact]
    public void AnExchangeStoreWithNoHashMatch_IsLeftToTheNameRule_TheOpenException()
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

        // Control: the same store with its signature hash in the index - tied by the hash ALONE,
        // under a name that is not its own: the exception does not ask for the name.
        StoreHash.TryDecodeHex("0A0B0C0D", out byte[] signature);
        StoreScopeInfo own = IndexRoot("someone else's label", StoreHash.Compute(signature));
        map = StoreIndexMatcher.Match(new[] { exchange }, new[] { byName, own });
        Assert.Equal(StoreIndexMatchKind.StoreHash, map.Stores[0].Kind);
        Assert.Equal(StoreHashInput.MappingSignature, map.Stores[0].Input);
        Assert.Same(own, map.Stores[0].Root);
    }

    [Fact]
    public void AnExchangeStoreWhoseProfileSignatureHashIsInTheIndex_IsFoundByIt()
    {
        const string signature = "8899AABBCCDDEEFF0011223344556677";
        StoreHash.TryDecodeHex(signature, out byte[] sig);
        StoreScopeInfo root = IndexRoot("alice@example.com", StoreHash.Compute(sig));
        var exchange = new StoreIndexIdentity("alice@example.com", 0,
            StoreHash.Candidates("00112233", true, false, null, signature, @"C:\Users\a\alice.ost"));

        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(new[] { exchange }, new[] { root }).Stores);
        Assert.Equal(StoreIndexMatchKind.StoreHash, match.Kind);
        Assert.Equal(StoreHashInput.ProfileMappingSignature, match.Input);
    }

    [Fact]
    public void AnExchangeHashTwoIndexStoresCarry_IsLeftToTheNameRule_AsBefore()
    {
        // D57, kept for the exception: a hash on two index stores is refused and the store takes
        // the name rule, with the reason in the note.
        StoreHash.TryDecodeHex("0A0B0C0D", out byte[] signature);
        uint hash = StoreHash.Compute(signature);
        var exchange = new StoreIndexIdentity("alice@example.com", 0,
            StoreHash.Candidates("00112233", true, false, "0A0B0C0D", null, null));

        StoreIndexMatch match = Assert.Single(StoreIndexMatcher.Match(
            new[] { exchange }, new[] { IndexRoot("one", hash), IndexRoot("two", hash) }).Stores);

        Assert.Equal(StoreIndexMatchKind.NameRule, match.Kind);
        Assert.Contains("more than one index store carries this store's hash", match.Note, StringComparison.Ordinal);
    }

    [Fact]
    public void APstAndAnExchangeStoreClaimingOneIndexStore_AreBothRefused()
    {
        // The Exchange store's claim is its hash alone, so the name cannot clear it: the PST is
        // refused as ambiguous (never picked), the Exchange store takes the name rule (D57).
        string path = @"C:\a.pst";
        StoreScopeInfo root = IndexRootOf("a", path);
        StoreIndexIdentity pst = PstStore("a", path);
        var exchange = new StoreIndexIdentity("b", 0, StoreHash.Candidates(Hex(Pst(path)), true, false, null, null, null));

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { pst, exchange }, new[] { root });

        Assert.Equal(StoreIndexMatchKind.Ambiguous, map.Stores[0].Kind);
        Assert.Equal(new[] { "b" }, map.Stores[0].SharedWith);
        Assert.Equal(StoreIndexMatchKind.NameRule, map.Stores[1].Kind);
        Assert.Contains("another Outlook store's hash is also", map.Stores[1].Note, StringComparison.Ordinal);
        Assert.Empty(map.Unclaimed);

        // Control: either alone is tied - the PST by name and hash, the Exchange store by hash.
        Assert.Equal(StoreIndexMatchKind.NameAndHash, StoreIndexMatcher.Match(new[] { pst }, new[] { root }).Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.StoreHash, StoreIndexMatcher.Match(new[] { exchange }, new[] { root }).Stores[0].Kind);
    }

    [Fact]
    public void ADelegateStore_IsNeverAStoreOfItsOwn()
    {
        StoreIndexIdentity host = PstStore("host", @"C:\h.pst");
        var delegated = new StoreIndexIdentity("Jan", 1, StoreHash.Candidates(Hex(Pst(@"C:\h.pst")), true, false, null, null, null));
        StoreScopeInfo root = IndexRootOf("host", @"C:\h.pst");

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { host, delegated }, new[] { root });

        Assert.Equal(StoreIndexMatchKind.NameAndHash, map.Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.Delegate, map.Stores[1].Kind);
        Assert.Null(map.Stores[1].Root);
    }
}
