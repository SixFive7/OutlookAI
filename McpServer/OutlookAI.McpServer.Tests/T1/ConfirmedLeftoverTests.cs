using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the one write a bystander ever gets (Q138): the confirmed leftover's pin and selection - its EntryID AND
/// its subject, both exactly - the gate that allows it only on the Exchange VM's test-account profile and only in
/// its one bystander, the check of the copy a move hands back, the count tripwire accepting that one tagged
/// departure and no other, the sweep's pre-run count coming down by it, and the one-off test that no derived
/// filter selects. Nothing here touches an Outlook: the EntryIDs are made up.
/// </summary>
public sealed class ConfirmedLeftoverTests
{
    private const string Hub = "test-account@example.invalid";
    private const string Bystander = "bystander@example.invalid";

    // Made-up 70-byte Exchange-shaped EntryIDs: flags, a provider id, the message type, then two halves.
    private const string Synthetic =
        "00000000AAAABBBBCCCCDDDDEEEEFFFF0000111107001111222233334444555566667777888800000000010A000011112222333344445555666677778888000000001234567800";
    private const string Other =
        "00000000AAAABBBBCCCCDDDDEEEEFFFF0000111107001111222233334444555566667777888800000000010A000011112222333344445555666677778888000000009876543200";

    private static readonly string TaggedSubject = LiveOutlookTestMailer.SubjectTag + " abc123def4567890 x1 seed";

    private static ConfirmedLeftover Pin() => ConfirmedLeftover.Create("T", 5, TaggedSubject, ConfirmedLeftover.HashEntryId(Synthetic));

    [Fact]
    public void TheQ138Pin_IsTaggedTestData_InSentItems_AndCarriesTheEntryIdOnlyAsAHash()
    {
        ConfirmedLeftover q138 = ConfirmedLeftover.Q138;
        Assert.StartsWith(LiveOutlookTestMailer.SubjectTag + " ", q138.Subject, StringComparison.Ordinal);
        Assert.Equal("[OutlookAI-McpTest] d3994a2ad25309442 d39 move seed", q138.Subject);
        Assert.Equal(5, q138.DefaultFolderId);
        Assert.Equal("Q138", q138.Decision);
        Assert.Matches("^[0-9a-f]{64}$", q138.EntryIdSha256);
        string described = q138.Describe();
        Assert.Contains(q138.EntryIdSha256[..12], described, StringComparison.Ordinal);
        Assert.DoesNotContain(q138.EntryIdSha256, described, StringComparison.Ordinal);
    }

    [Fact]
    public void TheHash_IsTheSha256OfTheTrimmedUpperCaseHex_InLowerCaseHex()
    {
        string expected = Convert.ToHexString(SHA256.HashData(Encoding.ASCII.GetBytes(Synthetic))).ToLowerInvariant();
        Assert.Equal(expected, ConfirmedLeftover.HashEntryId(Synthetic));
        Assert.Equal(expected, ConfirmedLeftover.HashEntryId(Synthetic.ToLowerInvariant()));
        Assert.Equal(expected, ConfirmedLeftover.HashEntryId("  " + Synthetic + "\r\n"));
        Assert.NotEqual(expected, ConfirmedLeftover.HashEntryId(Other));
    }

    [Theory]
    [InlineData(Synthetic, true, true)]
    [InlineData(Synthetic, true, false)]
    [InlineData(Other, false, false)]
    public void TheItem_IsItsEntryIdAndItsExactSubject(string entryId, bool expected, bool lowerCaseId)
    {
        Assert.Equal(expected, Pin().IsTheItem(lowerCaseId ? entryId.ToLowerInvariant() : entryId, TaggedSubject));
    }

    [Fact]
    public void ASubjectThatIsNotExactlyTheConfirmedOne_IsNeverTheItem_WhateverTheEntryId()
    {
        ConfirmedLeftover pin = Pin();
        Assert.False(pin.IsTheItem(Synthetic, TaggedSubject.ToUpperInvariant()));
        Assert.False(pin.IsTheItem(Synthetic, TaggedSubject + " "));
        Assert.False(pin.IsTheItem(Synthetic, " " + TaggedSubject));
        Assert.False(pin.IsTheItem(Synthetic, "RE: " + TaggedSubject));
        Assert.False(pin.IsTheItem(Synthetic, TaggedSubject.Replace("[OutlookAI-McpTest]", "[OutlookAI-Corpus]", StringComparison.Ordinal)));
        Assert.False(pin.IsTheItem(Synthetic, null));
        Assert.False(pin.IsTheItem(null, TaggedSubject));
        Assert.False(pin.IsTheItem(string.Empty, TaggedSubject));
        Assert.False(pin.IsTheItem("   ", TaggedSubject));
    }

    [Fact]
    public void APin_IsRefusedWithoutTheTagAtTheStart_OrWithAMalformedHash()
    {
        string hash = ConfirmedLeftover.HashEntryId(Synthetic);
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create("T", 5, "abc123def4567890 x1 seed", hash));
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create("T", 5, "OutlookAI-McpTest x1 seed", hash));
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create("T", 5, "RE: " + TaggedSubject, hash));
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create("T", 5, TaggedSubject, hash.ToUpperInvariant()));
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create("T", 5, TaggedSubject, hash[..63]));
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create("T", 5, TaggedSubject, Synthetic));
        Assert.Throws<ArgumentException>(() => ConfirmedLeftover.Create(" ", 5, TaggedSubject, hash));
    }

    [Fact]
    public void TheSelection_FindsExactlyOne_CountsTheRest_AndNeverGoesBySubjectAlone()
    {
        ConfirmedLeftover pin = Pin();

        ConfirmedLeftoverSelection none = pin.Select(Array.Empty<(string?, string?)>());
        Assert.Equal(ConfirmedLeftoverOutcome.NotPresent, none.Outcome);
        Assert.Null(none.EntryId);
        Assert.Equal(0, none.OtherTagged);

        ConfirmedLeftoverSelection one = pin.Select(new (string?, string?)[] { (Other, TaggedSubject), (Synthetic, TaggedSubject), (null, null) });
        Assert.Equal(ConfirmedLeftoverOutcome.Found, one.Outcome);
        Assert.Equal(Synthetic, one.EntryId);
        Assert.Equal(2, one.OtherTagged);

        // The same subject under another EntryID is not the item: the subject alone never selects.
        ConfirmedLeftoverSelection sameSubject = pin.Select(new (string?, string?)[] { (Other, TaggedSubject) });
        Assert.Equal(ConfirmedLeftoverOutcome.NotPresent, sameSubject.Outcome);
        Assert.Equal(1, sameSubject.OtherTagged);

        // One EntryID cannot be two items; if it ever reads so, nothing is chosen.
        ConfirmedLeftoverSelection twice = pin.Select(new (string?, string?)[] { (Synthetic, TaggedSubject), (Synthetic.ToLowerInvariant(), TaggedSubject) });
        Assert.Equal(ConfirmedLeftoverOutcome.Ambiguous, twice.Outcome);
        Assert.Null(twice.EntryId);
    }

    [Fact]
    public void TheGate_IsTheExchangeTestAccountsOneBystander_AndNothingElse()
    {
        ConfirmedLeftover pin = Pin();
        string[] one = { Bystander };
        Assert.Null(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, one, Hub, Bystander));
        Assert.Null(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, one, Hub, Bystander.ToUpperInvariant()));

        foreach (LiveMachineProfile profile in Enum.GetValues<LiveMachineProfile>().Where(p => p != LiveMachineProfile.ExchangeTestAccount))
        {
            Assert.StartsWith("REFUSING", pin.RefusalFor(profile, one, Hub, Bystander), StringComparison.Ordinal);
        }

        Assert.NotNull(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, Array.Empty<string>(), Hub, Bystander));
        Assert.NotNull(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, new[] { Bystander, "other@example.invalid" }, Hub, Bystander));
        Assert.NotNull(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, one, Hub, "other@example.invalid"));
        Assert.NotNull(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, one, Hub, Hub));
        Assert.NotNull(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, one, Hub, string.Empty));
        Assert.NotNull(pin.RefusalFor(LiveMachineProfile.ExchangeTestAccount, one, Bystander, Bystander));
    }

    [Fact]
    public void TheMovedCopy_IsDeletedOnlyWhenItIsStillTheItem_InDeletedItems_UnderANewEntryId()
    {
        ConfirmedLeftover pin = Pin();
        Assert.Null(pin.RefusalForMovedCopy(Other, TaggedSubject, inDeletedItems: true, Synthetic));
        Assert.NotNull(pin.RefusalForMovedCopy(null, TaggedSubject, inDeletedItems: true, Synthetic));
        Assert.NotNull(pin.RefusalForMovedCopy("  ", TaggedSubject, inDeletedItems: true, Synthetic));
        Assert.NotNull(pin.RefusalForMovedCopy(Other, TaggedSubject + " ", inDeletedItems: true, Synthetic));
        Assert.NotNull(pin.RefusalForMovedCopy(Other, null, inDeletedItems: true, Synthetic));
        Assert.NotNull(pin.RefusalForMovedCopy(Other, TaggedSubject, inDeletedItems: false, Synthetic));
        Assert.NotNull(pin.RefusalForMovedCopy(Synthetic.ToLowerInvariant(), TaggedSubject, inDeletedItems: true, Synthetic));
    }

    // ---- the count tripwire ----

    private static Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> Census(
        IReadOnlyList<CensusItem> bystanderSent, IReadOnlyList<CensusItem>? bystanderInbox = null)
    {
        return new Dictionary<string, IReadOnlyDictionary<string, FolderCensus>>(StringComparer.OrdinalIgnoreCase)
        {
            [Hub] = new Dictionary<string, FolderCensus>(StringComparer.Ordinal) { ["Inbox"] = FolderCensus.WithItems(Array.Empty<CensusItem>()) },
            [Bystander] = new Dictionary<string, FolderCensus>(StringComparer.Ordinal)
            {
                ["Sent Items"] = FolderCensus.WithItems(bystanderSent),
                ["Inbox"] = FolderCensus.WithItems(bystanderInbox ?? new[] { new CensusItem("REAL-1", "fp-r1", false) }),
            },
        };
    }

    private static readonly CensusItem Leftover = new(Synthetic, "fp-leftover", tagged: true);
    private static readonly CensusItem RealSent = new("REAL-2", "fp-r2", tagged: false);

    [Fact]
    public void WithoutAConfirmation_TheTaggedDepartureFromTheBystanderFails_AsTheSuites()
    {
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(Census(new[] { Leftover, RealSent }), Census(new[] { RealSent }), Hub);
        Assert.True(verdict.Failed);
        Assert.StartsWith("ATTRIBUTION: THE SUITE", verdict.Attribution, StringComparison.Ordinal);
    }

    [Fact]
    public void AConfirmedTaggedRemoval_IsNoted_NotFailed()
    {
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(
            Census(new[] { Leftover, RealSent }), Census(new[] { RealSent }), Hub, null,
            new[] { new ConfirmedRemoval(Bystander.ToUpperInvariant(), Synthetic.ToLowerInvariant()) });
        Assert.False(verdict.Failed, string.Join(Environment.NewLine, verdict.Failures));
        Assert.Null(verdict.Attribution);
        Assert.Contains(verdict.Notes, n => n.Contains("removed as confirmed", StringComparison.Ordinal) && n.Contains(Synthetic, StringComparison.Ordinal));
    }

    [Fact]
    public void AConfirmation_ExoneratesNothingElse()
    {
        ConfirmedRemoval[] confirmed = { new(Bystander, Synthetic) };

        // Another item leaving beside it still fails.
        TripwireVerdict alsoReal = StoreCountTripwire.Evaluate(Census(new[] { Leftover, RealSent }), Census(Array.Empty<CensusItem>()), Hub, null, confirmed);
        Assert.True(alsoReal.Failed);
        Assert.Single(alsoReal.Failures);
        Assert.Contains("REAL-2", alsoReal.Failures[0], StringComparison.Ordinal);
        Assert.DoesNotContain(Synthetic, alsoReal.Failures[0], StringComparison.Ordinal);

        // The confirmed EntryID on an item the census did NOT see tagged is no test item: it fails.
        CensusItem untagged = new(Synthetic, "fp-leftover", tagged: false);
        Assert.True(StoreCountTripwire.Evaluate(Census(new[] { untagged }), Census(Array.Empty<CensusItem>()), Hub, null, confirmed).Failed);

        // Confirmed for another store: it fails here.
        Assert.True(StoreCountTripwire.Evaluate(
            Census(new[] { Leftover }), Census(Array.Empty<CensusItem>()), Hub, null, new[] { new ConfirmedRemoval("other@example.invalid", Synthetic) }).Failed);

        // Another EntryID confirmed: it fails.
        Assert.True(StoreCountTripwire.Evaluate(
            Census(new[] { Leftover }), Census(Array.Empty<CensusItem>()), Hub, null, new[] { new ConfirmedRemoval(Bystander, Other) }).Failed);
    }

    [Fact]
    public void AConfirmation_IsUsedOnce_AndNeverForAFolderReadOnlyAsACount()
    {
        ConfirmedRemoval[] confirmed = { new(Bystander, Synthetic) };

        // The same EntryID leaving two folders: the second departure is not covered.
        CensusItem inInbox = new(Synthetic, "fp-other", tagged: true);
        TripwireVerdict twice = StoreCountTripwire.Evaluate(
            Census(new[] { Leftover }, new[] { inInbox }), Census(Array.Empty<CensusItem>(), Array.Empty<CensusItem>()), Hub, null, confirmed);
        Assert.True(twice.Failed);
        Assert.Single(twice.Failures);

        // A folder counted, not identified: a decrease fails as before, confirmation or not.
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> before = Census(Array.Empty<CensusItem>());
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> after = Census(Array.Empty<CensusItem>());
        before[Bystander] = new Dictionary<string, FolderCensus>(StringComparer.Ordinal) { ["Sent Items"] = FolderCensus.CountOnly(1, CensusCountReason.AbovePerFolderLimit) };
        after[Bystander] = new Dictionary<string, FolderCensus>(StringComparer.Ordinal) { ["Sent Items"] = FolderCensus.CountOnly(0, CensusCountReason.AbovePerFolderLimit) };
        TripwireVerdict counted = StoreCountTripwire.Evaluate(before, after, Hub, null, confirmed);
        Assert.True(counted.Failed);
        Assert.Contains("ITEMS LOST", counted.Failures[0], StringComparison.Ordinal);
    }

    // ---- the sweep's pre-run count ----

    [Fact]
    public void TheSweepsPreRunCount_ComesDownByEachConfirmedRemoval_NeverBelowZero_AndOnlyInItsStore()
    {
        Dictionary<string, int> one = new(StringComparer.OrdinalIgnoreCase) { [Bystander] = 1 };
        Assert.Equal(0, LiveArtifactBaseline.AdjustedFor(one, new[] { new ConfirmedRemoval(Bystander.ToUpperInvariant(), Synthetic) })[Bystander]);
        Assert.Equal(1, LiveArtifactBaseline.AdjustedFor(one, Array.Empty<ConfirmedRemoval>())[Bystander]);

        IReadOnlyDictionary<string, int> elsewhere = LiveArtifactBaseline.AdjustedFor(one, new[] { new ConfirmedRemoval("other@example.invalid", Synthetic) });
        Assert.Equal(1, elsewhere[Bystander]);
        Assert.False(elsewhere.ContainsKey("other@example.invalid"));

        Dictionary<string, int> zero = new(StringComparer.OrdinalIgnoreCase) { [Bystander] = 0 };
        Assert.Equal(0, LiveArtifactBaseline.AdjustedFor(zero, new[] { new ConfirmedRemoval(Bystander, Synthetic) })[Bystander]);

        Dictionary<string, int> two = new(StringComparer.OrdinalIgnoreCase) { [Bystander] = 2 };
        Assert.Equal(0, LiveArtifactBaseline.AdjustedFor(two, new[] { new ConfirmedRemoval(Bystander, Synthetic), new ConfirmedRemoval(Bystander, Other) })[Bystander]);
    }

    [Fact]
    public void AfterTheRemoval_TheSweepPassesTheBystanderAtZero_AndWouldFailItAtOne()
    {
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(new[] { Hub, Bystander }, new StoreWriteAllowlist(Hub, null, null, new[] { Bystander }));
        Dictionary<string, int> before = new(StringComparer.OrdinalIgnoreCase) { [Bystander] = 1 };
        IReadOnlyDictionary<string, int> adjusted = LiveArtifactBaseline.AdjustedFor(before, new[] { new ConfirmedRemoval(Bystander, Synthetic) });

        ArtifactSweepPolicy.Run(plan, _ => 0, _ => Assert.Fail("nothing to purge"), _ => { }, adjusted);
        Assert.Throws<InvalidOperationException>(() => ArtifactSweepPolicy.Run(plan, store => store == Bystander ? 1 : 0, _ => { }, _ => { }, adjusted));
    }

    // ---- the one-off test, and the filters ----

    [Fact]
    public void TheOneOffRemoval_IsOneTest_ThatNoDerivedFilterSelects()
    {
        List<MethodInfo> carriers = typeof(LiveCollections).Assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            .Where(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0)
            .Where(m => Traits(m, "Requires").Contains(LiveRunFilters.ConfirmedLeftover))
            .ToList();
        MethodInfo test = Assert.Single(carriers);
        Assert.Equal(typeof(LiveConfirmedLeftoverTests), test.DeclaringType);
        Assert.Contains("Live", Traits(test, "Category"));
        Assert.Contains(LiveRunFilters.CachedExchange, Traits(test, "Requires"));
        Assert.DoesNotContain(LiveRunFilters.WritesNothing, Traits(test, LiveRunFilters.WritesTrait));

        // The PST guests leave out CachedExchange; the read-only Exchange filter takes only Writes=Nothing; the
        // Exchange VM's write filter, and its narrowing, leave the capability out by name.
        Assert.Contains("Requires!=" + LiveRunFilters.CachedExchange, LiveRunFilters.Guest, StringComparison.Ordinal);
        Assert.Contains("Requires!=" + LiveRunFilters.CachedExchange, LiveRunFilters.GuestUnindexed, StringComparison.Ordinal);
        Assert.Contains(LiveRunFilters.WritesTrait + "=" + LiveRunFilters.WritesNothing, LiveRunFilters.ExchangeGuest, StringComparison.Ordinal);
        Assert.Contains("Requires!=" + LiveRunFilters.ConfirmedLeftover, LiveRunFilters.ExchangeTestAccount, StringComparison.Ordinal);
        Assert.Contains("Requires!=" + LiveRunFilters.ConfirmedLeftover, LiveRunFilters.ExchangeTestAccountWritesNothing, StringComparison.Ordinal);
    }

    private static List<string> Traits(MethodInfo method, string name)
    {
        return method.GetCustomAttributesData().Concat(method.DeclaringType!.GetCustomAttributesData())
            .Where(a => a.AttributeType == typeof(TraitAttribute) && a.ConstructorArguments.Count == 2)
            .Where(a => string.Equals(a.ConstructorArguments[0].Value as string, name, StringComparison.Ordinal))
            .Select(a => (string)a.ConstructorArguments[1].Value!)
            .ToList();
    }
}
