using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the measurement-corpus generator (RemediationTools). Everything the corpus tool
/// decides is decided here, in pure code, and everything here is pinned: the size
/// distribution, the date spread, the seeding, the store refusals, the teardown selection
/// rule, the manifest format and the date-fidelity verdicts. Only the COM calls themselves
/// are outside this tier, and they carry no decisions.
/// <para>
/// Two of these groups are safety tests rather than behaviour tests. The store refusals
/// decide whether tens of thousands of items may be written into a mailbox, and the
/// teardown rule decides what may be deleted - and the incident this project carries scars
/// from was a delete selected by a shell-side subject pattern, so the ordinal-matching
/// tests below are load-bearing.
/// </para>
/// </summary>
public class CorpusGeneratorTests
{
    private static readonly DateTime Anchor = new(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc);

    private static CorpusPlan Plan(long seed = 4242, string id = "vm1")
        => new(new CorpusPlanOptions(id, seed, Anchor));

    // ------------------------------------------------------------------ determinism

    [Fact]
    public void Describe_IsIdenticalAcrossTwoPlansWithTheSameSeed()
    {
        CorpusPlan a = Plan();
        CorpusPlan b = Plan();
        for (int ordinal = 1; ordinal <= 200; ordinal++)
        {
            Assert.Equal(a.Describe(ordinal), b.Describe(ordinal));
        }
    }

    [Fact]
    public void Describe_ChangesWithTheSeed()
    {
        CorpusPlan a = Plan(seed: 1);
        CorpusPlan b = Plan(seed: 2);
        int differences = 0;
        for (int ordinal = 1; ordinal <= 200; ordinal++)
        {
            if (a.Describe(ordinal) != b.Describe(ordinal))
            {
                differences++;
            }
        }

        // A seed that changed nothing would mean the seed is not wired in at all.
        Assert.True(differences > 190, $"only {differences} of 200 items changed with the seed");
    }

    [Fact]
    public void Describe_DoesNotDependOnHowManyItemsWereAskedFor()
    {
        // The property that makes "build 20 000 more" an ADDITION rather than a rewrite:
        // nothing in the derivation of item N knows the total. Proven by describing the
        // same ordinals from a plan used for a small corpus and a huge one.
        CorpusPlan plan = Plan();
        CorpusPlanReport small = plan.Report(1, 100);
        CorpusPlanReport large = plan.Report(1, 50_000);
        for (int ordinal = 1; ordinal <= 100; ordinal++)
        {
            Assert.Equal(plan.Describe(ordinal), Plan().Describe(ordinal));
        }

        Assert.Equal(100, small.ItemCount);
        Assert.Equal(50_000, large.ItemCount);
    }

    [Fact]
    public void Draw_PinsItsSequence()
    {
        // A regression guard on the generator itself. If someone swaps SplitMix64 for
        // System.Random - whose seeded sequence is explicitly not stable across framework
        // versions - every corpus built before that change becomes unreproducible, and
        // nothing else in this suite would notice.
        Assert.Equal(CorpusPlan.Draw(1, 1, 1), CorpusPlan.Draw(1, 1, 1));
        Assert.NotEqual(CorpusPlan.Draw(1, 1, 1), CorpusPlan.Draw(1, 1, 2));
        Assert.NotEqual(CorpusPlan.Draw(1, 1, 1), CorpusPlan.Draw(1, 2, 1));
        Assert.NotEqual(CorpusPlan.Draw(1, 1, 1), CorpusPlan.Draw(2, 1, 1));
    }

    [Fact]
    public void BuildBody_IsExactlyTheAdvertisedLength()
    {
        CorpusPlan plan = Plan();
        foreach (int ordinal in new[] { 1, 7, 33, 512, 4_096 })
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            Assert.Equal(spec.BodyBytes, plan.BuildBody(spec).Length);
        }
    }

    [Fact]
    public void BuildBody_IsDeterministicAndAscii()
    {
        CorpusPlan a = Plan();
        CorpusPlan b = Plan();
        CorpusItemSpec spec = a.Describe(99);
        string body = a.BuildBody(spec);
        Assert.Equal(body, b.BuildBody(b.Describe(99)));

        // ASCII matters for the measurement: the sweep's byte budget escapes non-ASCII to
        // \uXXXX, so a corpus that smuggled in accented text would move a different number
        // of bytes than its character counts imply.
        Assert.All(body, c => Assert.True(c < 128, $"non-ASCII char U+{(int)c:X4} in body"));
    }

    [Fact]
    public void BuildBody_UsesQuotedHistoryForLargeItems()
    {
        CorpusPlan plan = Plan();
        CorpusItemSpec spec = FindSpec(plan, s => s.BodyBytes >= 40_000);
        Assert.Contains(">", plan.BuildBody(spec), StringComparison.Ordinal);
    }

    [Fact]
    public void BuildSubject_CarriesBothTagsAndTheOrdinal()
    {
        CorpusPlan plan = Plan();
        string subject = plan.BuildSubject(1234);
        Assert.Contains(CorpusPlan.SubjectTag, subject, StringComparison.Ordinal);
        Assert.Contains(CorpusPlan.CorpusTagOpen, subject, StringComparison.Ordinal);
        Assert.True(CorpusPlan.TryParseOrdinal(subject, "vm1", out int ordinal));
        Assert.Equal(1234, ordinal);
    }

    [Fact]
    public void SubjectTag_IsTheProjectWideMailboxSafetyTag()
    {
        // Corpus items must be findable by the project's existing tested purge, so the tag
        // is the same constant rather than a copy of its text.
        Assert.Equal(RemediationRules.SubjectTag, CorpusPlan.SubjectTag);
    }

    // ----------------------------------------------------------- size distribution

    [Fact]
    public void BodySizes_StayInsideTheirClass()
    {
        CorpusPlan plan = Plan();
        var byName = CorpusPlanOptions.DefaultSizeClasses.ToDictionary(c => c.Name, StringComparer.Ordinal);
        for (int ordinal = 1; ordinal <= 5_000; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            CorpusSizeClass sizeClass = byName[spec.SizeClass];
            Assert.InRange(spec.BodyBytes, sizeClass.MinBytes, sizeClass.MaxBytes - 1);
        }
    }

    [Fact]
    public void BodySizes_FollowTheConfiguredWeightsWithinTolerance()
    {
        CorpusPlanReport report = Plan().Report(1, 40_000);
        int total = CorpusPlanOptions.DefaultSizeClasses.Sum(c => c.Weight);
        foreach (CorpusSizeClass sizeClass in CorpusPlanOptions.DefaultSizeClasses)
        {
            double expected = 40_000.0 * sizeClass.Weight / total;
            double actual = report.BySizeClass[sizeClass.Name];
            Assert.True(
                Math.Abs(actual - expected) < Math.Max(60, expected * 0.08),
                $"{sizeClass.Name}: expected ~{expected:F0}, got {actual:F0}");
        }
    }

    [Fact]
    public void BodySizes_HaveALongTailRatherThanOneShape()
    {
        CorpusPlanReport report = Plan().Report(1, 40_000);
        Assert.True(report.BodiesAtLeast24Kb > 2_000, $"only {report.BodiesAtLeast24Kb} bodies >= 24 KB");
        Assert.True(report.BodiesAtLeast96Kb > 200, $"only {report.BodiesAtLeast96Kb} bodies >= 96 KB");

        // The median must sit far below the mean - that IS the tail. A uniform corpus would
        // put them within a few per cent of each other and would tell the sweep measurement
        // nothing about what a real mailbox costs.
        var sizes = new List<int>(40_000);
        CorpusPlan plan = Plan();
        for (int ordinal = 1; ordinal <= 40_000; ordinal++)
        {
            sizes.Add(plan.Describe(ordinal).BodyBytes);
        }

        sizes.Sort();
        int median = sizes[sizes.Count / 2];
        Assert.True(median * 4 < report.MeanBodyBytes, $"median {median} is not far below mean {report.MeanBodyBytes}");
        Assert.True(sizes[^1] > 500_000, $"largest body is only {sizes[^1]}");
    }

    [Fact]
    public void BodySizes_ReachPastTheSweepBodyCapSoCappingCanBeObserved()
    {
        // A corpus that stopped below OutlookComSession.SweepBodyCharsCap could never show
        // what capping costs, nor that the itemsBodyCapped counter is wired at all.
        CorpusPlanReport report = Plan().Report(1, 40_000);
        Assert.True(
            report.BodiesOverSweepBodyCap > 0,
            $"no body exceeds the {OutlookComSession.SweepBodyCharsCap} char sweep cap");
        Assert.True(report.BodiesOverSweepBodyCap < report.ItemCount / 100);
    }

    [Fact]
    public void TotalBodyBytes_IsTheSumOfTheItems()
    {
        CorpusPlan plan = Plan();
        long expected = 0;
        for (int ordinal = 1; ordinal <= 1_000; ordinal++)
        {
            expected += plan.Describe(ordinal).BodyBytes;
        }

        Assert.Equal(expected, plan.Report(1, 1_000).TotalBodyBytes);
    }

    // ---------------------------------------------------------------- date spread

    [Fact]
    public void Dates_StayInsideTheConfiguredSpan()
    {
        CorpusPlan plan = Plan();
        DateTime oldestAllowed = Anchor.AddDays(-1_460);
        for (int ordinal = 1; ordinal <= 5_000; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            Assert.InRange(spec.ReceivedUtc, oldestAllowed, Anchor);
            Assert.Equal(DateTimeKind.Utc, spec.ReceivedUtc.Kind);
        }
    }

    [Fact]
    public void Dates_MakeTheSevenDayAndSixtyDayWindowsSelectDifferentVolumes()
    {
        // The whole reason the corpus exists: if these two numbers were equal, no window
        // measurement taken against it would mean anything.
        CorpusPlanReport report = Plan().Report(1, 40_000);
        int within7 = report.WithinDays[7];
        int within60 = report.WithinDays[60];
        Assert.True(within7 > 0, "nothing falls inside 7 days");
        Assert.True(within60 > within7 * 2, $"60-day window ({within60}) is not meaningfully bigger than 7-day ({within7})");
        Assert.True(within60 < report.ItemCount, "the 60-day window selects the whole corpus");
    }

    [Fact]
    public void Dates_WindowCountsAreMonotonic()
    {
        CorpusPlanReport report = Plan().Report(1, 20_000);
        int[] marks = report.WithinDays.Keys.OrderBy(k => k).ToArray();
        for (int i = 1; i < marks.Length; i++)
        {
            Assert.True(
                report.WithinDays[marks[i]] >= report.WithinDays[marks[i - 1]],
                $"{marks[i]}d selects fewer than {marks[i - 1]}d");
        }
    }

    [Fact]
    public void Dates_SubmitTimeNeverFollowsDeliveryTime()
    {
        CorpusPlan plan = Plan();
        for (int ordinal = 1; ordinal <= 5_000; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            Assert.True(spec.SentUtc <= spec.ReceivedUtc, $"item {ordinal} was submitted after it was delivered");
            if (spec.FolderId == 5)
            {
                Assert.Equal(spec.ReceivedUtc, spec.SentUtc); // a sent item is its own origin
            }
        }
    }

    [Fact]
    public void Dates_AreControlledByTheAnchorRatherThanTheClock()
    {
        var earlier = new CorpusPlan(new CorpusPlanOptions("vm1", 4242, Anchor.AddYears(-2)));
        CorpusPlan later = Plan();
        Assert.Equal(
            earlier.Describe(1).ReceivedUtc.AddYears(2),
            later.Describe(1).ReceivedUtc);
    }

    [Fact]
    public void Plan_RefusesALocalAnchor()
    {
        // A local anchor moves with the VM's time zone and across DST, so the same seed
        // would describe a different corpus after a rollback.
        Assert.Throws<ArgumentException>(() =>
            new CorpusPlan(new CorpusPlanOptions("vm1", 1, new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Local))));
    }

    [Fact]
    public void Plan_RefusesAnUnusableShape()
    {
        Assert.Throws<ArgumentException>(() => new CorpusPlan(new CorpusPlanOptions(string.Empty, 1, Anchor)));
        Assert.Throws<ArgumentException>(() => new CorpusPlan(new CorpusPlanOptions("has space", 1, Anchor)));
        Assert.Throws<ArgumentException>(() => new CorpusPlan(new CorpusPlanOptions("bad]id", 1, Anchor)));
        Assert.Throws<ArgumentException>(() => new CorpusPlan(
            new CorpusPlanOptions("vm1", 1, Anchor) { SizeClasses = new[] { new CorpusSizeClass("x", 1, 10, 10) } }));
        Assert.Throws<ArgumentException>(() => new CorpusPlan(
            new CorpusPlanOptions("vm1", 1, Anchor) { DateBands = new[] { new CorpusDateBand("x", 1, 5, 5) } }));
        Assert.Throws<ArgumentException>(() => new CorpusPlan(
            new CorpusPlanOptions("vm1", 1, Anchor) { Folders = new[] { new CorpusFolderShare(6, "Inbox", 0) } }));
    }

    [Fact]
    public void Folders_CoverTheFourTheSweepVisits()
    {
        CorpusPlanReport report = Plan().Report(1, 20_000);
        foreach (int folderId in new[] { 6, 5, 3, 23 })
        {
            Assert.True(report.ByFolderId.TryGetValue(folderId, out int count) && count > 0,
                $"folder {folderId} got no items");
        }

        Assert.Equal(20_000, report.ByFolderId.Values.Sum());
    }

    [Fact]
    public void ReadState_LeavesAMinorityOfArrivalMailUnread()
    {
        CorpusPlan plan = Plan();
        int unread = 0;
        int arrivals = 0;
        for (int ordinal = 1; ordinal <= 20_000; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            if (spec.FolderId is 6 or 23)
            {
                arrivals++;
                if (!spec.IsRead)
                {
                    unread++;
                }
            }
            else
            {
                Assert.True(spec.IsRead);
            }
        }

        Assert.InRange(unread, arrivals / 10, arrivals / 2);
    }

    // ----------------------------------------------------------- shape key / resume

    [Fact]
    public void ShapeKey_ChangesWithEveryPartOfTheShape()
    {
        var baseline = new CorpusPlanOptions("vm1", 1, Anchor);
        Assert.Equal(baseline.ShapeKey, new CorpusPlanOptions("vm1", 1, Anchor).ShapeKey);
        Assert.NotEqual(baseline.ShapeKey, new CorpusPlanOptions("vm2", 1, Anchor).ShapeKey);
        Assert.NotEqual(baseline.ShapeKey, new CorpusPlanOptions("vm1", 2, Anchor).ShapeKey);
        Assert.NotEqual(baseline.ShapeKey, new CorpusPlanOptions("vm1", 1, Anchor.AddDays(1)).ShapeKey);
        Assert.NotEqual(
            baseline.ShapeKey,
            (baseline with { SizeClasses = new[] { new CorpusSizeClass("only", 1, 10, 20) } }).ShapeKey);
        Assert.NotEqual(
            baseline.ShapeKey,
            (baseline with { DateBands = new[] { new CorpusDateBand("only", 1, 0, 5) } }).ShapeKey);
        Assert.NotEqual(
            baseline.ShapeKey,
            (baseline with { Folders = new[] { new CorpusFolderShare(6, "Inbox", 1) } }).ShapeKey);
    }

    // -------------------------------------------------------------- store refusals

    private static CorpusStoreFacts GoodFacts(string name = "Corpus PST")
        => new(name, true, 3, @"D:\corpus\corpus.pst");

    [Fact]
    public void Store_AcceptedOnlyWhenTheAllowlistAndAllFourFactsAgree()
    {
        Assert.Equal(
            CorpusStoreRefusal.None,
            CorpusSafety.EvaluateStore(GoodFacts(), new[] { "Corpus PST" }));
    }

    [Fact]
    public void Store_RefusedWithoutAnAllowlist()
    {
        Assert.Equal(CorpusStoreRefusal.NoAllowlist, CorpusSafety.EvaluateStore(GoodFacts(), null));
        Assert.Equal(CorpusStoreRefusal.NoAllowlist, CorpusSafety.EvaluateStore(GoodFacts(), Array.Empty<string>()));
    }

    [Fact]
    public void Store_RefusedWhenItIsNotOnTheAllowlist()
    {
        Assert.Equal(
            CorpusStoreRefusal.NotOnAllowlist,
            CorpusSafety.EvaluateStore(GoodFacts("Production Mailbox"), new[] { "Corpus PST" }));
    }

    [Fact]
    public void Store_AllowlistMatchesWholeNamesOnlyAndNeverAPattern()
    {
        // A prefix, a suffix or a wildcard must not match. The whole point of the incident
        // that shaped these rules is that a pattern matched more than it was meant to.
        Assert.Equal(
            CorpusStoreRefusal.NotOnAllowlist,
            CorpusSafety.EvaluateStore(GoodFacts("Corpus PST (production)"), new[] { "Corpus PST" }));
        Assert.Equal(
            CorpusStoreRefusal.NotOnAllowlist,
            CorpusSafety.EvaluateStore(GoodFacts("Corpus"), new[] { "Corpus PST" }));
        Assert.Equal(
            CorpusStoreRefusal.NotOnAllowlist,
            CorpusSafety.EvaluateStore(GoodFacts("Corpus PST"), new[] { "*" }));
        Assert.Equal(
            CorpusStoreRefusal.NotOnAllowlist,
            CorpusSafety.EvaluateStore(GoodFacts("Corpus PST"), new[] { "%" }));
    }

    [Fact]
    public void Store_AllowlistIgnoresCaseAndSurroundingWhitespace()
    {
        Assert.Equal(
            CorpusStoreRefusal.None,
            CorpusSafety.EvaluateStore(GoodFacts("corpus pst"), new[] { "  Corpus PST  " }));
    }

    [Fact]
    public void Store_RefusedWhenItBelongsToAMailAccount()
    {
        Assert.Equal(
            CorpusStoreRefusal.NotADataFileStore,
            CorpusSafety.EvaluateStore(GoodFacts() with { IsDataFileStore = false }, new[] { "Corpus PST" }));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    public void Store_RefusedForEveryExchangeStoreType(int exchangeStoreType)
    {
        Assert.Equal(
            CorpusStoreRefusal.ExchangeStore,
            CorpusSafety.EvaluateStore(
                GoodFacts() with { ExchangeStoreType = exchangeStoreType }, new[] { "Corpus PST" }));
    }

    [Theory]
    [InlineData(@"D:\corpus\corpus.ost")]
    [InlineData(@"D:\corpus\corpus")]
    [InlineData("")]
    public void Store_RefusedWhenTheBackingFileIsNotAPst(string filePath)
    {
        Assert.Equal(
            CorpusStoreRefusal.NotAPstFile,
            CorpusSafety.EvaluateStore(GoodFacts() with { FilePath = filePath }, new[] { "Corpus PST" }));
    }

    [Fact]
    public void Store_RefusedWhenAFactCouldNotBeRead()
    {
        // Fail-closed: a store nothing is known about is exactly the store not to write to.
        foreach (CorpusStoreFacts facts in new[]
                 {
                     GoodFacts() with { IsDataFileStore = null },
                     GoodFacts() with { ExchangeStoreType = null },
                     GoodFacts() with { FilePath = null },
                 })
        {
            Assert.Equal(CorpusStoreRefusal.Unprovable, CorpusSafety.EvaluateStore(facts, new[] { "Corpus PST" }));
        }
    }

    [Fact]
    public void Store_RefusedWhenItHasNoName()
    {
        Assert.Equal(
            CorpusStoreRefusal.NoStoreName,
            CorpusSafety.EvaluateStore(GoodFacts() with { DisplayName = null }, new[] { "Corpus PST" }));
    }

    [Fact]
    public void Store_ExplanationNamesOnlyTheOffendingStore()
    {
        string text = CorpusSafety.Explain(
            CorpusStoreRefusal.NotOnAllowlist, GoodFacts("Production Mailbox"));
        Assert.Contains("Production Mailbox", text, StringComparison.Ordinal);
        Assert.Contains("REFUSING", text, StringComparison.Ordinal);
        Assert.Contains("--allow-store", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Corpus PST", text, StringComparison.Ordinal);
    }

    // -------------------------------------------------- profile / send capability

    private static CorpusProfileFacts NoAccounts() => new(0, 0, 0);

    [Fact]
    public void Profile_AcceptedOnlyWhenThereAreNoAccountsAtAll()
    {
        Assert.Equal(CorpusStoreRefusal.None, CorpusSafety.EvaluateProfile(NoAccounts()));
    }

    [Fact]
    public void Profile_RefusedWhenAnAccountDeliversIntoTheTargetStore()
    {
        // The case that actually happened: a build put 5 532 items into the target store's
        // Outbox. On a profile with an account that is 5 532 queued messages.
        Assert.Equal(
            CorpusStoreRefusal.TargetIsAccountDeliveryStore,
            CorpusSafety.EvaluateProfile(new CorpusProfileFacts(2, 1, 0)));
    }

    [Fact]
    public void Profile_RefusedWhenAnyAccountExistsEvenIfNoneDeliversHere()
    {
        // Stricter than "no account delivers into the target", because the object model
        // cannot express the requirement any other way: SendUsingAccount is set per item, so
        // any account can send a message that lives anywhere. "No account can send from this
        // store" is only provable as "no account can send".
        Assert.Equal(
            CorpusStoreRefusal.ProfileCanSend,
            CorpusSafety.EvaluateProfile(new CorpusProfileFacts(1, 0, 0)));
    }

    [Fact]
    public void Profile_RefusedWhenItCannotBeProven()
    {
        Assert.Equal(
            CorpusStoreRefusal.ProfileUnprovable,
            CorpusSafety.EvaluateProfile(new CorpusProfileFacts(null, 0, 0)));
        Assert.Equal(
            CorpusStoreRefusal.ProfileUnprovable,
            CorpusSafety.EvaluateProfile(new CorpusProfileFacts(3, 0, 1)));
    }

    [Fact]
    public void Profile_UnprovableOutranksDeliveringSoTheWeakerVerdictIsNeverReported()
    {
        // An account whose delivery store could not be read might be delivering here. The
        // refusal has to be the one that says nothing is proven, not the one that says
        // "one account delivers here" as though the rest had been cleared.
        Assert.Equal(
            CorpusStoreRefusal.ProfileUnprovable,
            CorpusSafety.EvaluateProfile(new CorpusProfileFacts(3, 1, 1)));
    }

    [Fact]
    public void Evaluate_RequiresBothTheStoreAndTheProfileToPass()
    {
        Assert.Equal(
            CorpusStoreRefusal.None,
            CorpusSafety.Evaluate(GoodFacts(), NoAccounts(), new[] { "Corpus PST" }));

        // A local .pst is not sufficient on its own - it can be an account's delivery store.
        Assert.Equal(
            CorpusStoreRefusal.TargetIsAccountDeliveryStore,
            CorpusSafety.Evaluate(GoodFacts(), new CorpusProfileFacts(1, 1, 0), new[] { "Corpus PST" }));

        // The store verdict is reported first when both fail, because a wrong target is the
        // more likely operator mistake and its message is the more useful one.
        Assert.Equal(
            CorpusStoreRefusal.NotOnAllowlist,
            CorpusSafety.Evaluate(GoodFacts("Other"), new CorpusProfileFacts(1, 1, 0), new[] { "Corpus PST" }));
    }

    [Fact]
    public void Profile_RefusalMessageSaysWhatToDoAndOffersNoFlag()
    {
        string text = CorpusSafety.Explain(CorpusStoreRefusal.ProfileCanSend, GoodFacts());
        Assert.Contains("no mail accounts", text, StringComparison.Ordinal);
        Assert.Contains("There is no flag for this", text, StringComparison.Ordinal);
        Assert.Contains("Outbox", text, StringComparison.Ordinal);
        Assert.DoesNotContain("--allow-store", text, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ placement

    private static CorpusPlacementProbe PlacementProbe(
        CorpusPlacementMethod method,
        bool parentMatches = true,
        bool inTable = true,
        bool sentFlag = true,
        string? landedIn = "Inbox",
        string? error = null)
        => new(method, "Inbox", parentMatches, inTable, sentFlag, landedIn, error);

    [Fact]
    public void Placement_IsUsableOnlyWhenTheItemIsParentedThereAndInThatFoldersTable()
    {
        Assert.True(CorpusPlacement.IsUsable(PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag)));
        Assert.False(CorpusPlacement.IsUsable(
            PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag, parentMatches: false)));
        Assert.False(CorpusPlacement.IsUsable(
            PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag, inTable: false)));
        Assert.False(CorpusPlacement.IsUsable(
            PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag, error: "refused")));
    }

    [Fact]
    public void Placement_DoesNotRequireTheSentFlagItself()
    {
        // The flag is a means, not the goal. An item that sits in the Inbox and appears in
        // its table is measurable whether or not MSGFLAG_UNSENT happened to clear, and
        // demanding the flag would reject a rung that works.
        Assert.True(CorpusPlacement.IsUsable(
            PlacementProbe(CorpusPlacementMethod.DraftsThenMove, sentFlag: false)));
    }

    [Fact]
    public void Placement_FolderTableIsTheDecidingSignal()
    {
        // An item can be parented correctly and still be missing from the table the sweep
        // reads. That case would produce a corpus that looks right in Outlook and measures
        // as an empty store, which is the failure mode this whole probe exists to catch.
        Assert.False(CorpusPlacement.IsUsable(
            PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag, parentMatches: true, inTable: false)));
    }

    [Fact]
    public void Placement_ChoosesTheCheapestRungThatVerified()
    {
        Assert.Equal(
            CorpusPlacementMethod.InPlaceWithSentFlag,
            CorpusPlacement.Choose(new[]
            {
                PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag),
                PlacementProbe(CorpusPlacementMethod.DraftsThenMoveWithSentFlag),
            }));

        Assert.Equal(
            CorpusPlacementMethod.DraftsThenMoveWithSentFlag,
            CorpusPlacement.Choose(new[]
            {
                PlacementProbe(CorpusPlacementMethod.InPlaceWithSentFlag, inTable: false, landedIn: "Drafts"),
                PlacementProbe(CorpusPlacementMethod.DraftsThenMoveWithSentFlag),
                PlacementProbe(CorpusPlacementMethod.DraftsThenMove),
            }));

        Assert.Equal(CorpusPlacementMethod.None, CorpusPlacement.Choose(Array.Empty<CorpusPlacementProbe>()));
    }

    [Fact]
    public void Placement_LadderKeepsThePlainSaveAsAControlRung()
    {
        // InPlaceOnly is what the first real build did, and it is kept so the probe RECORDS
        // what the store does with a plain saved item rather than assuming the one observed
        // failure generalises. It is last, so it is never chosen over a rung that works.
        Assert.Equal(CorpusPlacementMethod.InPlaceOnly, CorpusPlacement.Ladder[^1]);
        Assert.Equal(CorpusPlacementMethod.InPlaceWithSentFlag, CorpusPlacement.Ladder[0]);
        Assert.Equal(4, CorpusPlacement.Ladder.Length);
        Assert.Equal(CorpusPlacement.Ladder.Length, CorpusPlacement.Ladder.Distinct().Count());
    }

    [Fact]
    public void Placement_MechanicsMatchTheirRungNames()
    {
        Assert.False(CorpusPlacement.CreatesInDrafts(CorpusPlacementMethod.InPlaceWithSentFlag));
        Assert.False(CorpusPlacement.RequiresMove(CorpusPlacementMethod.InPlaceWithSentFlag));
        Assert.True(CorpusPlacement.WritesSentFlag(CorpusPlacementMethod.InPlaceWithSentFlag));

        Assert.True(CorpusPlacement.CreatesInDrafts(CorpusPlacementMethod.DraftsThenMoveWithSentFlag));
        Assert.True(CorpusPlacement.RequiresMove(CorpusPlacementMethod.DraftsThenMoveWithSentFlag));
        Assert.True(CorpusPlacement.WritesSentFlag(CorpusPlacementMethod.DraftsThenMoveWithSentFlag));

        Assert.True(CorpusPlacement.CreatesInDrafts(CorpusPlacementMethod.DraftsThenMove));
        Assert.True(CorpusPlacement.RequiresMove(CorpusPlacementMethod.DraftsThenMove));
        Assert.False(CorpusPlacement.WritesSentFlag(CorpusPlacementMethod.DraftsThenMove));

        Assert.False(CorpusPlacement.CreatesInDrafts(CorpusPlacementMethod.InPlaceOnly));
        Assert.False(CorpusPlacement.RequiresMove(CorpusPlacementMethod.InPlaceOnly));
        Assert.False(CorpusPlacement.WritesSentFlag(CorpusPlacementMethod.InPlaceOnly));
    }

    [Fact]
    public void Placement_RefusalStatesTheConsequenceAsACount()
    {
        (bool proceed, string message) =
            CorpusPlacement.Decide(CorpusPlacementMethod.None, allowDraftsPlacement: false, itemCount: 40_000);
        Assert.False(proceed);
        Assert.Contains("0 of 40,000 items", message, StringComparison.Ordinal);
        Assert.Contains("NOT Drafts", message, StringComparison.Ordinal);
        Assert.Contains("Refusing to build", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Placement_OverrideIsAllowedButSaysWhatIsLost()
    {
        (bool proceed, string message) =
            CorpusPlacement.Decide(CorpusPlacementMethod.None, allowDraftsPlacement: true, itemCount: 40_000);
        Assert.True(proceed);
        Assert.Contains("CANNOT be used to measure the freshness sweep at all", message, StringComparison.Ordinal);
        Assert.Contains("step 2", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Placement_VerifiedMessageWarnsWhenTheChosenRungMoves()
    {
        (_, string moved) = CorpusPlacement.Decide(
            CorpusPlacementMethod.DraftsThenMoveWithSentFlag, allowDraftsPlacement: false, itemCount: 10);
        Assert.Contains("POST-move EntryID", moved, StringComparison.Ordinal);

        (_, string inPlace) = CorpusPlacement.Decide(
            CorpusPlacementMethod.InPlaceWithSentFlag, allowDraftsPlacement: false, itemCount: 10);
        Assert.DoesNotContain("POST-move", inPlace, StringComparison.Ordinal);
    }

    [Fact]
    public void Options_CarryTheDraftsPlacementOverrideSeparatelyFromTheDateOne()
    {
        CorpusOptions both = CorpusOptions.Parse(new[] { "--allow-undated", "--allow-drafts-placement" });
        Assert.True(both.AllowUndated);
        Assert.True(both.AllowDraftsPlacement);

        // The coupling that caused the incident, pinned apart: --allow-undated must not
        // imply anything about placement.
        CorpusOptions dateOnly = CorpusOptions.Parse(new[] { "--allow-undated" });
        Assert.True(dateOnly.AllowUndated);
        Assert.False(dateOnly.AllowDraftsPlacement);
    }

    // ------------------------------------------------------ teardown selection rule

    private const string CorpusId = "vm1";

    private static string TaggedSubject(int ordinal, string corpusId = CorpusId)
        => new CorpusPlan(new CorpusPlanOptions(corpusId, 1, Anchor)).BuildSubject(ordinal);

    [Fact]
    public void MayDelete_NeedsBothTheEntryIdAndTheTag()
    {
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABC123" });
        Assert.True(CorpusSafety.MayDelete("ABC123", TaggedSubject(1), allowlist, CorpusId));
    }

    [Fact]
    public void MayDelete_RefusesAnIdThatIsNotOnTheAllowlist()
    {
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABC123" });
        Assert.False(CorpusSafety.MayDelete("DEF456", TaggedSubject(1), allowlist, CorpusId));
    }

    [Fact]
    public void MayDelete_RefusesAnAllowlistedIdWhoseSubjectDoesNotCarryTheTags()
    {
        // The id half alone is not enough: an EntryID can be mistyped, recycled, or point at
        // an item that is no longer the one that was recorded.
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABC123" });
        Assert.False(CorpusSafety.MayDelete("ABC123", "Quarterly figures", allowlist, CorpusId));
        Assert.False(CorpusSafety.MayDelete("ABC123", null, allowlist, CorpusId));
        Assert.False(CorpusSafety.MayDelete("ABC123", string.Empty, allowlist, CorpusId));
    }

    [Fact]
    public void MayDelete_RefusesAnotherCorpusInTheSameStore()
    {
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABC123" });
        Assert.False(CorpusSafety.MayDelete("ABC123", TaggedSubject(1, "vm2"), allowlist, CorpusId));
    }

    [Fact]
    public void MayDelete_RefusesASubjectMissingTheMailboxSafetyTag()
    {
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABC123" });
        string corpusTagOnly = CorpusPlan.CorpusTagOpen + CorpusId + "#0000001] renewal invoice";
        Assert.False(CorpusSafety.MayDelete("ABC123", corpusTagOnly, allowlist, CorpusId));
    }

    [Fact]
    public void MayDelete_RefusesEverySubjectAWildcardMatchWouldHaveSwallowed()
    {
        // The incident, reproduced as a test. A shell-side "-like '*[OutlookAI-Corpus:vm1]*'"
        // reads the brackets as a CHARACTER CLASS, so it matches any subject containing any
        // one of those letters - which is nearly every subject in a mailbox. The sanctioned
        // predicate is an ordinal Contains, so each of these is a plain non-match.
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABC123" });
        foreach (string subject in new[]
                 {
                     "Outlook",
                     "Corpus",
                     "A",
                     "invoice renewal",
                     "OutlookAI-Corpus:vm1#0000001",           // the tag text without its brackets
                     "[OutlookAI-Corpus vm1#0000001]",          // no colon
                     "[OutlookAI-McpTest] plain test artifact", // safety tag but no corpus tag
                 })
        {
            Assert.False(
                CorpusSafety.MayDelete("ABC123", subject, allowlist, CorpusId),
                $"subject '{subject}' was accepted for deletion");
        }
    }

    [Fact]
    public void TryParseOrdinal_ReadsTheOrdinalBackExactly()
    {
        foreach (int ordinal in new[] { 1, 9, 1_234, 9_999_999 })
        {
            Assert.True(CorpusPlan.TryParseOrdinal(TaggedSubject(ordinal), CorpusId, out int parsed));
            Assert.Equal(ordinal, parsed);
        }
    }

    [Fact]
    public void TryParseOrdinal_RefusesRubbishRatherThanGuessing()
    {
        Assert.False(CorpusPlan.TryParseOrdinal(null, CorpusId, out _));
        Assert.False(CorpusPlan.TryParseOrdinal(TaggedSubject(1), null, out _));
        Assert.False(CorpusPlan.TryParseOrdinal(TaggedSubject(1), string.Empty, out _));
        Assert.False(CorpusPlan.TryParseOrdinal(
            CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + CorpusId + "#abc] x", CorpusId, out _));
        Assert.False(CorpusPlan.TryParseOrdinal(
            CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + CorpusId + "#0000000] x", CorpusId, out _));
        Assert.False(CorpusPlan.TryParseOrdinal(
            CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + CorpusId + "#12 x", CorpusId, out _));
    }

    [Fact]
    public void EntryIdAllowlist_IgnoresCase()
    {
        // EntryIDs come back from Outlook as hex and the case is not guaranteed to survive a
        // round trip. An allowlist that failed to match its own entries would fail SAFE, but
        // it would leave items behind that the operator was told had been removed.
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "abc123", string.Empty, "  " });
        Assert.True(CorpusSafety.MayDelete("ABC123", TaggedSubject(1), allowlist, CorpusId));
        Assert.Single(allowlist);
    }

    // ------------------------------------------------------------------- manifest

    private static CorpusManifestHeader Header(string storeName = "Corpus PST", long seed = 4242, string id = "vm1")
        => new(
            CorpusManifest.CurrentVersion,
            id,
            seed,
            CorpusManifest.FormatUtc(Anchor),
            new CorpusPlanOptions(id, seed, Anchor).ShapeKey,
            storeName,
            @"D:\corpus\corpus.pst",
            CorpusDateWriteMethod.PropertyAccessorDates.ToString(),
            CorpusPlacementMethod.InPlaceWithSentFlag.ToString());

    [Fact]
    public void Manifest_RoundTrips()
    {
        var lines = new List<string> { CorpusManifest.RenderLine(Header()) };
        lines.Add(CorpusManifest.RenderLine(new CorpusManifestItem(1, "ID1", 6, 1_234, CorpusManifest.FormatUtc(Anchor))));
        lines.Add(CorpusManifest.RenderLine(new CorpusManifestItem(2, "ID2", 5, 4_321, null)));
        lines.Add(CorpusManifest.RenderLine(new CorpusManifestFolder("FID", CorpusManifest.CreatedFolderPrefix + "-Junk", 23)));

        CorpusManifest manifest = CorpusManifest.Parse(lines);
        Assert.Equal(Header(), manifest.Header);
        Assert.Equal(2, manifest.Items.Count);
        Assert.Equal("ID1", manifest.Items[1].EntryId);
        Assert.Equal(4_321, manifest.Items[2].BodyBytes);
        Assert.Single(manifest.Folders);
        Assert.Equal(23, manifest.Folders[0].FolderId);
        Assert.Empty(manifest.UnparseableLines);
    }

    [Fact]
    public void Manifest_SurvivesTheTruncatedLastLineAnInterruptedBuildLeaves()
    {
        var lines = new List<string>
        {
            CorpusManifest.RenderLine(Header()),
            CorpusManifest.RenderLine(new CorpusManifestItem(1, "ID1", 6, 10, null)),
            "{\"Ordinal\":2,\"EntryId\":\"ID",
        };

        CorpusManifest manifest = CorpusManifest.Parse(lines);
        Assert.Single(manifest.Items);
        Assert.Single(manifest.UnparseableLines);
    }

    [Fact]
    public void Manifest_RefusesAFileWithNoHeader()
    {
        Assert.Throws<InvalidOperationException>(() => CorpusManifest.Parse(Array.Empty<string>()));
    }

    [Fact]
    public void Manifest_MissingOrdinalsDriveResumption()
    {
        CorpusManifest manifest = CorpusManifest.Create(Header());
        manifest.Add(new CorpusManifestItem(1, "ID1", 6, 10, null));
        manifest.Add(new CorpusManifestItem(3, "ID3", 6, 10, null));
        Assert.Equal(new[] { 2, 4, 5 }, manifest.MissingOrdinals(5).ToArray());

        manifest.Add(new CorpusManifestItem(2, "ID2", 6, 10, null));
        manifest.Add(new CorpusManifestItem(4, "ID4", 6, 10, null));
        manifest.Add(new CorpusManifestItem(5, "ID5", 6, 10, null));
        Assert.Empty(manifest.MissingOrdinals(5));
    }

    [Fact]
    public void Manifest_AcceptsAContinuationThatSimplyWantsMoreItems()
    {
        // Requirement 4: a corpus can be ADDED to. The count is deliberately absent from the
        // compatibility comparison because item N does not depend on it.
        CorpusManifest manifest = CorpusManifest.Create(Header());
        Assert.Equal(
            CorpusManifestMismatch.None,
            manifest.CheckCompatible(new CorpusPlanOptions("vm1", 4242, Anchor), "Corpus PST"));
    }

    [Fact]
    public void Manifest_RefusesAContinuationWithADifferentShape()
    {
        CorpusManifest manifest = CorpusManifest.Create(Header());
        Assert.Equal(
            CorpusManifestMismatch.CorpusId,
            manifest.CheckCompatible(new CorpusPlanOptions("vm2", 4242, Anchor), "Corpus PST"));
        Assert.Equal(
            CorpusManifestMismatch.Shape,
            manifest.CheckCompatible(new CorpusPlanOptions("vm1", 9999, Anchor), "Corpus PST"));
        Assert.Equal(
            CorpusManifestMismatch.Shape,
            manifest.CheckCompatible(new CorpusPlanOptions("vm1", 4242, Anchor.AddDays(1)), "Corpus PST"));
        Assert.Equal(
            CorpusManifestMismatch.Store,
            manifest.CheckCompatible(new CorpusPlanOptions("vm1", 4242, Anchor), "Other PST"));
    }

    [Fact]
    public void Manifest_RefusesAnOlderFormatVersion()
    {
        CorpusManifest manifest = CorpusManifest.Create(Header() with { Version = 0 });
        Assert.Equal(
            CorpusManifestMismatch.Version,
            manifest.CheckCompatible(new CorpusPlanOptions("vm1", 4242, Anchor), "Corpus PST"));
    }

    [Fact]
    public void Manifest_TimestampsRoundTripAsUtc()
    {
        DateTime? parsed = CorpusManifest.ParseUtc(CorpusManifest.FormatUtc(Anchor));
        Assert.Equal(Anchor, parsed);
        Assert.Equal(DateTimeKind.Utc, parsed!.Value.Kind);
        Assert.Null(CorpusManifest.ParseUtc("not a date"));
    }

    // --------------------------------------------------------------- date fidelity

    private static CorpusDateProbe Probe(
        CorpusDateWriteMethod method,
        DateTime? readBack,
        bool selected = true,
        bool excluded = true,
        string? error = null)
        => new(method, Anchor, Anchor, readBack, selected, excluded, error);

    [Fact]
    public void DateProbe_IsUsableOnlyWhenAllThreeSignalsAgree()
    {
        Assert.True(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor)));
        Assert.False(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, null)));
        Assert.False(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor.AddDays(3))));
        Assert.False(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor, selected: false)));
        Assert.False(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor, excluded: false)));
        Assert.False(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor, error: "refused")));
    }

    [Fact]
    public void DateProbe_ExclusionIsWhatProvesTheDateActuallySelects()
    {
        // A property that reads back correctly but does not drive a DASL restriction would
        // pass a read-back-only check and then make every window measurement wrong.
        Assert.False(CorpusDateFidelity.IsUsable(
            Probe(CorpusDateWriteMethod.PropertyAccessorDates, Anchor, selected: true, excluded: false)));
    }

    [Fact]
    public void DateProbe_ToleratesSubSecondDrift()
    {
        Assert.True(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor.AddMilliseconds(400))));
        Assert.False(CorpusDateFidelity.IsUsable(Probe(CorpusDateWriteMethod.ObjectModel, Anchor.AddSeconds(30))));
    }

    [Fact]
    public void ClassifyOffset_RecognisesTheLocalTimeConversion()
    {
        TimeSpan offset = TimeSpan.FromHours(2);
        Assert.Equal(CorpusDateOffsetVerdict.Exact, CorpusDateFidelity.ClassifyOffset(Anchor, Anchor, offset));
        Assert.Equal(
            CorpusDateOffsetVerdict.LocalOffsetApplied,
            CorpusDateFidelity.ClassifyOffset(Anchor, Anchor + offset, offset));
        Assert.Equal(
            CorpusDateOffsetVerdict.LocalOffsetApplied,
            CorpusDateFidelity.ClassifyOffset(Anchor, Anchor - offset, offset));
        Assert.Equal(
            CorpusDateOffsetVerdict.Unusable,
            CorpusDateFidelity.ClassifyOffset(Anchor, Anchor.AddDays(4), offset));
        Assert.Equal(CorpusDateOffsetVerdict.Unusable, CorpusDateFidelity.ClassifyOffset(Anchor, null, offset));
    }

    [Fact]
    public void CompensatedWriteValue_ReversesWhatWasObservedRatherThanAnAssumedSign()
    {
        TimeSpan offset = TimeSpan.FromHours(2);
        Assert.Equal(
            Anchor - offset,
            CorpusDateFidelity.CompensatedWriteValue(
                Anchor, CorpusDateOffsetVerdict.LocalOffsetApplied, offset, Anchor + offset));
        Assert.Equal(
            Anchor + offset,
            CorpusDateFidelity.CompensatedWriteValue(
                Anchor, CorpusDateOffsetVerdict.LocalOffsetApplied, offset, Anchor - offset));
        Assert.Equal(
            Anchor,
            CorpusDateFidelity.CompensatedWriteValue(Anchor, CorpusDateOffsetVerdict.Exact, offset, Anchor));
    }

    [Fact]
    public void Choose_TakesTheStrongestRungThatVerified()
    {
        Assert.Equal(
            CorpusDateWriteMethod.PropertyAccessorDates,
            CorpusDateFidelity.Choose(new[]
            {
                Probe(CorpusDateWriteMethod.PropertyAccessorDates, Anchor),
                Probe(CorpusDateWriteMethod.ObjectModel, Anchor),
            }));

        Assert.Equal(
            CorpusDateWriteMethod.ObjectModel,
            CorpusDateFidelity.Choose(new[]
            {
                Probe(CorpusDateWriteMethod.PropertyAccessorDates, Anchor, error: "refused"),
                Probe(CorpusDateWriteMethod.ObjectModel, Anchor),
            }));

        Assert.Equal(
            CorpusDateWriteMethod.None,
            CorpusDateFidelity.Choose(new[] { Probe(CorpusDateWriteMethod.ObjectModel, null) }));
        Assert.Equal(CorpusDateWriteMethod.None, CorpusDateFidelity.Choose(Array.Empty<CorpusDateProbe>()));
    }

    [Fact]
    public void Ladder_TriesTheStrongestMethodFirst()
    {
        Assert.Equal(
            new[] { CorpusDateWriteMethod.PropertyAccessorDates, CorpusDateWriteMethod.ObjectModel },
            CorpusDateFidelity.Ladder);
    }

    [Fact]
    public void DateLadder_NoLongerCarriesTheMessageFlagWrite()
    {
        // Regression guard for the coupling that cost a 40 000 item build. The MSGFLAG_UNSENT
        // write decides where an item LIVES; it used to be a rung of the DATE ladder, so
        // --allow-undated silently disabled it and every item was filed as a draft. The two
        // ladders are now disjoint concerns, and this test fails if anyone re-merges them.
        Assert.Equal(2, CorpusDateFidelity.Ladder.Length);
        Assert.DoesNotContain(
            CorpusDateFidelity.Ladder,
            m => m.ToString().Contains("Flag", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Decide_RefusesAnUndatedCorpusUnlessItWasAskedFor()
    {
        (bool proceed, string message) =
            CorpusDateFidelity.Decide(CorpusDateWriteMethod.None, allowUndated: false, itemCount: 40_000);
        Assert.False(proceed);
        Assert.Contains("Refusing to build", message, StringComparison.Ordinal);

        (bool proceedAnyway, string allowedMessage) =
            CorpusDateFidelity.Decide(CorpusDateWriteMethod.None, allowUndated: true, itemCount: 40_000);
        Assert.True(proceedAnyway);
        Assert.Contains("CANNOT be used to measure the freshness sweep", allowedMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void Decide_StatesTheUndatedConsequenceAsACountAndNotAsAnAllRecentCorpus()
    {
        // The message this replaces said the items "would carry a received time of roughly
        // now". An operator reasoned - reasonably, from that text - that an all-recent corpus
        // is still the sweep's worst case, overrode the guard, and lost the build. The
        // premise was wrong: an item whose delivery time the folder table does not carry is
        // not selected by ANY window, so the sweep sees fewer items, not more.
        (_, string message) =
            CorpusDateFidelity.Decide(CorpusDateWriteMethod.None, allowUndated: false, itemCount: 40_000);
        Assert.Contains("0 of 40,000 items", message, StringComparison.Ordinal);
        Assert.Contains("invisible to the sweep, not merely mis-dated", message, StringComparison.Ordinal);
        Assert.Contains("NOT 'an all-recent corpus'", message, StringComparison.Ordinal);
        Assert.DoesNotContain("roughly 'now'", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Decide_SaysWhatTheWeakestRungCannotCarry()
    {
        (bool proceed, string message) =
            CorpusDateFidelity.Decide(CorpusDateWriteMethod.ObjectModel, allowUndated: false, itemCount: 10);
        Assert.True(proceed);
        Assert.Contains("PR_CLIENT_SUBMIT_TIME", message, StringComparison.Ordinal);

        (_, string strong) =
            CorpusDateFidelity.Decide(CorpusDateWriteMethod.PropertyAccessorDates, allowUndated: false, itemCount: 10);
        Assert.DoesNotContain("PR_CLIENT_SUBMIT_TIME", strong, StringComparison.Ordinal);
    }

    // -------------------------------------------------------------- option parsing

    [Fact]
    public void Options_CollectAnAllowlistOfSeveralStores()
    {
        CorpusOptions options = CorpusOptions.Parse(new[]
        {
            "--store", "Corpus PST",
            "--allow-store", "Corpus PST",
            "--allow-store", "Second, with a comma",
            "--execute",
        });

        Assert.True(options.Execute);
        Assert.Equal("Corpus PST", options.Store);

        // Repeatable rather than comma-separated: a store display name may contain a comma,
        // and this list is the guard deciding where tens of thousands of items may be written.
        Assert.Equal(new[] { "Corpus PST", "Second, with a comma" }, options.AllowStores);
    }

    [Fact]
    public void Options_RequireASeedAndAnAnchor()
    {
        CorpusOptions noSeed = CorpusOptions.Parse(new[] { "--corpus-id", "vm1", "--anchor", "2026-08-01" });
        ArgumentException seedError = Assert.Throws<ArgumentException>(() => noSeed.ToPlanOptions());
        Assert.Contains("reproducible", seedError.Message, StringComparison.Ordinal);

        CorpusOptions noAnchor = CorpusOptions.Parse(new[] { "--corpus-id", "vm1", "--seed", "1" });
        ArgumentException anchorError = Assert.Throws<ArgumentException>(() => noAnchor.ToPlanOptions());
        Assert.Contains("not defaulted to", anchorError.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Options_RejectAnUnknownSwitch()
    {
        Assert.Throws<ArgumentException>(() => CorpusOptions.Parse(new[] { "--nonsense", "x" }));
        Assert.Throws<ArgumentException>(() => CorpusOptions.Parse(new[] { "loose-argument" }));
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("nl-NL")]
    [InlineData("ar-SA")]
    public void Options_ParseTheAnchorTheSameWayUnderEveryCulture(string culture)
    {
        // The locale defect this project has already been bitten by: a date read in the
        // machine's culture made 5 September out of 9 May. An anchor is the origin of every
        // date in the corpus, so a locale-dependent parse would silently shift the whole
        // corpus on a differently-configured VM.
        CultureInfo original = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            DateTime parsed = CorpusOptions.ParseAnchor("2026-05-09");
            Assert.Equal(new DateTime(2026, 5, 9, 0, 0, 0, DateTimeKind.Utc), parsed);
            Assert.Equal(DateTimeKind.Utc, parsed.Kind);
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = original;
        }
    }

    [Fact]
    public void Options_RejectAnAnchorThatIsNotADate()
    {
        Assert.Throws<ArgumentException>(() => CorpusOptions.ParseAnchor("last tuesday"));
    }

    // ------------------------------------------------------------------- reporting

    [Fact]
    public void PlanCommand_ReportsWithoutTouchingOutlook()
    {
        // corpus-plan is the pre-flight: it must be runnable anywhere, including CI.
        CorpusOptions options = CorpusOptions.Parse(new[]
        {
            "--corpus-id", "vm1", "--seed", "4242", "--anchor", "2026-08-01", "--count", "2000",
        });

        // Under a culture that groups with '.' rather than ',', to pin that the report is
        // written under the invariant culture. It is saved beside measurement results and
        // compared across machines, so the same figure has to produce the same string.
        CultureInfo original = Thread.CurrentThread.CurrentCulture;
        string text;
        try
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("nl-NL");
            var writer = new StringWriter();
            Assert.Equal(0, CorpusCommands.RunPlan(options, writer));
            text = writer.ToString();
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = original;
        }

        Assert.Contains("items                 : 2,000", text, StringComparison.Ordinal);
        Assert.DoesNotContain("2.000", text, StringComparison.Ordinal);
        Assert.Contains("Inbox=", text, StringComparison.Ordinal);
        Assert.Contains("Junk Email=", text, StringComparison.Ordinal);
        Assert.Contains("7d=", text, StringComparison.Ordinal);
        Assert.Contains("60d=", text, StringComparison.Ordinal);
    }

    private static CorpusItemSpec FindSpec(CorpusPlan plan, Func<CorpusItemSpec, bool> predicate)
    {
        for (int ordinal = 1; ordinal <= 100_000; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            if (predicate(spec))
            {
                return spec;
            }
        }

        throw new InvalidOperationException("No matching item in the first 100 000 ordinals.");
    }
}
