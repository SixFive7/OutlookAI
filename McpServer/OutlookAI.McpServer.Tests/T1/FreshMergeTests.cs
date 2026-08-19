using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: fresh-mode merge logic (v3.MD D19) - term re-application over gap-swept items
/// and boundary de-duplication against index hits. Index hits are fabricated through
/// the public IndexRowMapper (synthetic URLs, S6-safe).
/// </summary>
public sealed class FreshMergeTests
{
    private static readonly DateTime BaseUtc = new(2026, 07, 23, 10, 00, 00, DateTimeKind.Utc);

    private static ComMailBrief Brief(
        string subject = "Quarterly invoice",
        string store = "alice@example.com",
        string folder = "Inbox",
        DateTime? receivedLocal = null,
        string? body = null,
        string? senderName = null,
        string? senderAddress = null)
    {
        return new ComMailBrief(
            entryId: "AA" + Guid.NewGuid().ToString("N"),
            storeDisplayName: store,
            storeId: null,
            folderName: folder,
            folderKind: "inbox",
            subject: subject,
            senderName: senderName,
            senderAddress: senderAddress,
            receivedTime: receivedLocal ?? BaseUtc.ToLocalTime(),
            isRead: false,
            hasAttachments: false,
            sizeBytes: 1000,
            body: body);
    }

    private static IndexHit Hit(string subject = "Quarterly invoice", string store = "alice@example.com", string folder = "Inbox", DateTime? receivedUtc = null)
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] =
                $"mapi16://{{S-1-5-21-1-2-3-1001}}/{store}($abcd1234)/0/{folder}/{EntryIdCodecTests.SyntheticEncodedTail()}",
            ["System.Subject"] = subject,
            ["System.Message.DateReceived"] = (receivedUtc ?? BaseUtc).ToUniversalTime(),
        };
        return IndexRowMapper.Map(row);
    }

    // ---------------------------------------------------------------- MatchesTerms

    private const SearchIn Default = SearchInValues.Default;

    [Fact]
    public void MatchesTerms_NullOrEmpty_MatchesEverything()
    {
        Assert.True(FreshMerge.MatchesTerms(Brief(), null, Default));
        Assert.True(FreshMerge.MatchesTerms(Brief(), Array.Empty<string>(), Default));
    }

    [Fact]
    public void MatchesTerms_AndSemantics_AllTermsMustHit()
    {
        ComMailBrief item = Brief(subject: "Invoice for April", body: "total 100 euro");
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "invoice", "euro" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "invoice", "missingterm" }, Default));
    }

    [Fact]
    public void MatchesTerms_CaseInsensitive_AcrossSubjectAndBody()
    {
        ComMailBrief item = Brief(subject: "hello", body: "WORLD");
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "HELLO" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "world" }, Default));
    }

    [Fact]
    public void MatchesTerms_SenderIsNotMatchedByTerms_MatchingSenderAloneDoesNotHit()
    {
        // D40/SF-6 tier alignment: the index tier never matched senders by term, so the
        // sweep must not either - otherwise a hit would vanish once the frontier passed
        // the item. Sender matching is the 'from' filter's job (applied by MailService).
        ComMailBrief item = Brief(
            subject: "hello", body: "world", senderName: "Charlie", senderAddress: "c@example.com");
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "charlie" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "c@example.com" }, Default));
    }

    [Fact]
    public void MatchesTerms_PrefixStar_MatchesStem()
    {
        ComMailBrief item = Brief(subject: "factuur 2026-001");
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "fact*" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "xyz*" }, Default));
    }

    [Fact]
    public void MatchesTerms_NoBody_StillMatchesOnSubject()
    {
        ComMailBrief item = Brief(subject: "Order confirmation", body: null);
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "order" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "invoice" }, Default));
    }

    // ------------------------------------------------- search_in scopes (D40, user 2026-07-26)

    [Fact]
    public void MatchesTerms_SubjectOnlyScope_IgnoresBody()
    {
        ComMailBrief item = Brief(subject: "alert prefix", body: "requeued backend");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "alert" }, SearchIn.SubjectOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "requeued" }, SearchIn.SubjectOnly));
    }

    [Fact]
    public void MatchesTerms_BodyOnlyScope_IgnoresSubject()
    {
        ComMailBrief item = Brief(subject: "alert prefix", body: "requeued backend");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "requeued" }, SearchIn.BodyOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "alert" }, SearchIn.BodyOnly));
    }

    [Fact]
    public void MatchesTerms_DefaultScope_FindsSubjectOnlyAndBodyOnlyTerms()
    {
        // The SF-6 shape in miniature: a term that lives only in the subject must be
        // found by the default scope (that is the whole point of the fix).
        ComMailBrief item = Brief(subject: "alert prefix", body: "requeued backend");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "alert" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "requeued" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "alert", "requeued" }, Default));
    }

    [Fact]
    public void MatchesTerms_AndsAcrossSubjectAndBody_NotInsideOneOfThem()
    {
        // Tier parity with the index builder (soak fix 13): the sweep must match mail
        // carrying one term only in the subject and the other only in the body. This
        // tier already ANDed per term over the whole text - the pin keeps it that way.
        ComMailBrief item = Brief(subject: "Balans 2026", body: "verbruik energie per maand");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "balans", "energie" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "energie", "balans" }, Default));

        // Narrowing to one part must NOT find the cross-part pair.
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "balans", "energie" }, SearchIn.SubjectOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "balans", "energie" }, SearchIn.BodyOnly));
    }

    [Fact]
    public void MatchesTerms_ScopesHonorPrefixStems()
    {
        ComMailBrief item = Brief(subject: "factuur 2026-001", body: "betaling ontvangen");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "fact*" }, SearchIn.SubjectOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "fact*" }, SearchIn.BodyOnly));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "betal*" }, SearchIn.BodyOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "betal*" }, SearchIn.SubjectOnly));
    }

    // ---------------------------------------------------------------- IsDuplicate

    [Fact]
    public void IsDuplicate_SameStoreFolderSubjectAndTime_IsTrue()
    {
        Assert.True(FreshMerge.IsDuplicate(Brief(), Hit(), toleranceSeconds: 15));
    }

    [Fact]
    public void IsDuplicate_DifferentFolder_SentVsInboxCopy_IsFalse()
    {
        // A self-send: identical subject + near-identical time, but Sent Items vs Inbox
        // must remain two distinct hits.
        ComMailBrief sentCopy = Brief(folder: "Sent Items");
        Assert.False(FreshMerge.IsDuplicate(sentCopy, Hit(folder: "Inbox"), toleranceSeconds: 15));
    }

    [Fact]
    public void IsDuplicate_DifferentStore_IsFalse()
    {
        Assert.False(FreshMerge.IsDuplicate(Brief(store: "bob@example.com"), Hit(store: "alice@example.com"), 15));
    }

    [Fact]
    public void IsDuplicate_DifferentSubject_IsFalse()
    {
        Assert.False(FreshMerge.IsDuplicate(Brief(subject: "Other"), Hit(subject: "Quarterly invoice"), 15));
    }

    [Fact]
    public void IsDuplicate_TimeOutsideTolerance_IsFalse()
    {
        ComMailBrief item = Brief(receivedLocal: BaseUtc.AddMinutes(10).ToLocalTime());
        Assert.False(FreshMerge.IsDuplicate(item, Hit(receivedUtc: BaseUtc), toleranceSeconds: 15));
    }

    // ---------------------------------------------------------------- SelectFreshOnly

    [Fact]
    public void SelectFreshOnly_DropsIndexDuplicates_KeepsNewItems()
    {
        var swept = new List<ComMailBrief>
        {
            Brief(subject: "Quarterly invoice"),               // duplicate of the index hit
            Brief(subject: "Brand new mail", folder: "Inbox"), // genuinely fresh
        };
        var hits = new List<IndexHit> { Hit(subject: "Quarterly invoice") };

        IReadOnlyList<ComMailBrief> fresh = FreshMerge.SelectFreshOnly(swept, hits, 15, out int duplicates);

        Assert.Single(fresh);
        Assert.Equal("Brand new mail", fresh[0].Subject);
        Assert.Equal(1, duplicates);
    }

    [Fact]
    public void SelectFreshOnly_DropsRepeatedEntryIds()
    {
        ComMailBrief item = Brief(subject: "Once");
        var swept = new List<ComMailBrief> { item, item };

        IReadOnlyList<ComMailBrief> fresh = FreshMerge.SelectFreshOnly(swept, Array.Empty<IndexHit>(), 15, out int duplicates);

        Assert.Single(fresh);
        Assert.Equal(1, duplicates);
    }

    // ---------------------------------------------------------------- ResolveHitStore

    [Fact]
    public void ResolveHitStore_DelegateSubtree_UsesFirstFolderSegment()
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = "mapi16://{S-1-5-21-1-2-3-1001}/owner@example.com($abcd1234)/1/Delegate Name/Postvak IN/"
                + EntryIdCodecTests.SyntheticEncodedTail(),
            ["System.Subject"] = "x",
        };
        IndexHit hit = IndexRowMapper.Map(row);
        Assert.Equal(1, hit.StoreType);
        Assert.Equal("Delegate Name", FreshMerge.ResolveHitStore(hit));
    }

    [Fact]
    public void ResolveHitStore_PrimaryStore_UsesStoreDisplayName()
    {
        Assert.Equal("alice@example.com", FreshMerge.ResolveHitStore(Hit()));
    }

    // ============================================== D47: what the sweep cannot answer

    [Fact]
    public void Sweep_IsRefused_ForAnAttachmentOnlySearch_BecauseItNeverOpensAnAttachment()
    {
        // THE DEFECT THIS PINS: the sweep reads Subject/Body through COM, so every row it
        // can produce is a MESSAGE row. Merging those under an attachment-ONLY filter
        // returned precisely the rows the filter excludes.
        Assert.Equal(
            FreshMerge.AttachmentContentNotSweepable,
            FreshMerge.SweepRefusalReason(hasRecipientFilter: false, attachmentHitsOnly: true));
    }

    [Fact]
    public void Sweep_Runs_WhenAttachmentHitsAreMerelyExcluded_TheMirrorImageCase()
    {
        // The mirror image is NOT symmetric, and deliberately so: excluding attachment
        // hits asks for message rows, which is all the sweep produces. It must still run,
        // or freshness coverage would be lost for no reason.
        Assert.Null(FreshMerge.SweepRefusalReason(hasRecipientFilter: false, attachmentHitsOnly: false));
    }

    [Fact]
    public void Sweep_RecipientFilterRefusal_StillWins_WhenBothApply()
    {
        // Order is contractual: the recipient refusal is the older, more specific advice.
        Assert.Equal(
            FreshMerge.RecipientFilterNotSweepable,
            FreshMerge.SweepRefusalReason(hasRecipientFilter: true, attachmentHitsOnly: true));
        Assert.Equal(
            FreshMerge.RecipientFilterNotSweepable,
            FreshMerge.SweepRefusalReason(hasRecipientFilter: true, attachmentHitsOnly: false));
    }

    [Fact]
    public void SweepRefusalReasons_AreDistinctMachineReadableTokens()
    {
        Assert.NotEqual(FreshMerge.RecipientFilterNotSweepable, FreshMerge.AttachmentContentNotSweepable);
        Assert.DoesNotContain(" ", FreshMerge.AttachmentContentNotSweepable, StringComparison.Ordinal);
    }

    // ------------------------------------------------- freshness coverage (three states)
    //
    // THE DEFECT THIS SECTION PINS: degraded/freshness used to be set from sweep.performed
    // alone, so a sweep that RAN and covered part of its scope - a folder it could not
    // enumerate, a cap, a budget - reported freshness "live" with no degradation while
    // advice said in prose that coverage was partial. An agent reading fields rather than
    // prose, which is the sensible way to read a payload, was told a partial answer was
    // complete. Every hole below is reachable only with a real mailbox (a folder tree that
    // fails, truncates or runs long), so it is proven here on the payload block the COM
    // layer fills in, which is where the classification actually happens.

    private static SweepInfo Swept(int foldersSwept = 3)
    {
        return new SweepInfo { Performed = true, FoldersSwept = foldersSwept };
    }

    /// <summary>Every gap code declared on <see cref="FreshMerge"/>, read from the type itself.</summary>
    private static IReadOnlyList<string> AllGapCodes()
    {
        return typeof(FreshMerge)
            .GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(f => f.IsLiteral && f.FieldType == typeof(string) && f.Name.StartsWith("Gap", StringComparison.Ordinal))
            .Select(f => (string)f.GetRawConstantValue()!)
            .ToList();
    }

    [Fact]
    public void ASweepThatCoveredItsWholeScope_IsLive_AndReportsNoGaps()
    {
        SweepInfo sweep = Swept();
        Assert.Null(FreshMerge.DescribeCoverageGaps(sweep));
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void ASweepThatNeverRan_IsIndexOnly_AndIsNotReportedAsPartial()
    {
        // "Did not run" and "ran but covered part of it" are different states with
        // different remedies; a sweep that never ran must not borrow the partial vocabulary.
        SweepInfo sweep = new SweepInfo { Performed = false, Error = "OutlookUnavailable" };
        Assert.Null(FreshMerge.DescribeCoverageGaps(sweep));
        Assert.Equal(FreshMerge.FreshnessIndexOnly, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void NoSweepBlockAtAll_StaysLive_BecauseThatCallerAskedForIndexRows()
    {
        // The internal index-only escape hatch (SearchRequest.IndexOnly, not on the MCP
        // tool): nothing was withheld, so nothing is degraded.
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(null));
    }

    /// <summary>
    /// Every way a sweep that RAN can have covered less than its scope, paired with the
    /// code it must report. Written as data rather than as one test per hole so the set
    /// itself can be compared against the codes the type declares.
    /// </summary>
    private static List<(string Gap, SweepInfo Sweep)> CoverageHoles()
    {
        List<(string Gap, SweepInfo Sweep)> data = new();

        // 1. The sweep ran but reached no folder at all - a whole-scope miss that was
        //    silent in BOTH fields and prose whenever no folder was requested (a
        //    store name that matches nothing is skipped without even a skip count).
        data.Add((FreshMerge.GapNothingSwept, Swept(foldersSwept: 0)));

        // 2. Folders whose item enumeration failed: no freshness coverage there at all.
        data.Add((FreshMerge.GapFoldersFailed, new SweepInfo { Performed = true, FoldersSwept = 2, FoldersFailed = 1 }));

        // 3. The subtree walk stopped at MaxScopedSweepFolders.
        data.Add((FreshMerge.GapFolderCap, new SweepInfo { Performed = true, FoldersSwept = 40, FolderCapReached = true }));

        // 4. The subtree walk stopped at ScopedSweepTimeBudgetMs.
        data.Add((FreshMerge.GapTimeBudget, new SweepInfo { Performed = true, FoldersSwept = 7, TimeBudgetExceeded = true }));

        // 4b. The WHOLE sweep ran out of MailService.SweepWorkBudgetMs and stopped at a
        //     store or folder boundary. Its own code, because the remedy points the other
        //     way from 4: that one says one subtree is too wide (scope it, or drop
        //     include_subfolders), this one says the profile is too big for one sweep (name
        //     a store). It replaced an outcome rather than adding one - before the sweep had
        //     a budget of its own this arrived as a gateway timeout and a killed COM host,
        //     with every folder already swept thrown away.
        data.Add((
            FreshMerge.GapSweepBudget,
            new SweepInfo { Performed = true, FoldersSwept = 6, SweepBudgetExpired = true }));

        // 5. The subtree walk refused folders past the depth guard.
        data.Add((FreshMerge.GapDepthLimit, new SweepInfo { Performed = true, FoldersSwept = 9, DepthLimitReached = true }));

        // 6. Folders skipped because they could not be resolved or enumerated.
        data.Add((FreshMerge.GapFoldersSkipped, new SweepInfo { Performed = true, FoldersSwept = 3, FoldersSkipped = 2 }));

        // 7. The per-folder item cap truncated a folder's window (newest-first, so the
        //    OLDEST not-yet-indexed mail there is the part that is missing).
        data.Add((
            FreshMerge.GapItemCap,
            new SweepInfo { Performed = true, FoldersSwept = 3, ItemCappedFolders = new[] { "alice@example.com/Inbox" } }));

        // 7b. The same cap over a folder whose table Outlook would NOT sort, so the cut is
        //     arbitrary and "the oldest is missing" would be a false statement about it
        //     (gap H2). Its own code, because it leads somewhere different: nothing here
        //     tells the caller WHICH mail is absent.
        data.Add((
            FreshMerge.GapItemCapUnsorted,
            new SweepInfo
            {
                Performed = true,
                FoldersSwept = 3,
                ItemCappedFolders = new[] { "alice@example.com/Inbox" },
                ItemCappedFoldersUnsorted = new[] { "alice@example.com/Inbox" },
            }));

        // 8. A store in scope with NO index rows at all: there was no frontier to open the
        //    window from, so it fell back to a fixed span and everything older than that
        //    span is in neither tier. The sweep itself covered its whole scope, which is
        //    exactly why this was silent - every counter said "complete".
        data.Add((
            FreshMerge.GapNoIndexFrontier,
            new SweepInfo
            {
                Performed = true,
                FoldersSwept = 4,
                IndexFrontierMissing = true,
                StoresWithoutIndex = new[] { "Archive 2019.pst" },
            }));

        // 9. Rows inside a folder that WAS enumerated and could not be turned into items
        //    (gap H1). No folder counter can express this: the folder really was read, and
        //    such a row did not even count toward the per-folder cap - so a folder where
        //    every row failed reported itself swept, complete and empty.
        data.Add((
            FreshMerge.GapRowsUnreadable,
            new SweepInfo { Performed = true, FoldersSwept = 4, RowsUnreadable = 3 }));

        // 10. Items dropped because a filter the CALLER passed could not be evaluated on
        //     them (gap I1). The sweep covered every folder it was asked to; the loss is one
        //     level further in, at the item, and it was the last silent drop on this path.
        data.Add((
            FreshMerge.GapFilterUnreadable,
            new SweepInfo
            {
                Performed = true,
                FoldersSwept = 4,
                ItemsFilterUnreadable = 2,
                FiltersUnevaluated = new[] { "unread_only" },
            }));

        // 11. An item whose BODY was cut at the COM layer so one answer could not outgrow
        //     the frame carrying it, AND which then failed to match the terms. Narrower
        //     still than 10: the item was read, was inside the window and is in the sweep's
        //     own result set - it simply may have matched in the part that was not carried.
        //     Only the intersection raises it; a cut body on an item that matched anyway
        //     cost nothing (SweepBodyCapTests pins that direction).
        data.Add((
            FreshMerge.GapBodyCap,
            new SweepInfo
            {
                Performed = true,
                FoldersSwept = 4,
                ItemsBodyCapped = 3,
                ItemsBodyCappedUnmatched = 2,
            }));

        return data;
    }

    // ---------------------------------- (A1) no index frontier: a whole tier contributed nothing

    // THE DEFECT: the sweep window is the index frontier minus a safety margin, and a scope
    // with no indexed mail has no frontier, so the code substituted "seven days ago" and said
    // nothing. Everything older than that, in a store the index has never seen, was in
    // NEITHER tier - not indexed, and before the window - while the payload read
    // freshness:"live", no degraded, no coverage gaps, foldersSwept:4, and no advice. It is
    // the shape a local-PST-only profile takes whenever Windows Search has not indexed it.

    [Fact]
    public void AScopeWithNoIndexFrontier_IsPartialAndDegraded_EvenThoughTheSweepCoveredEverything()
    {
        SweepInfo sweep = new SweepInfo { Performed = true, FoldersSwept = 4, IndexFrontierMissing = true };

        // Every sweep counter says "complete", which is true and was the whole problem.
        Assert.Equal(new[] { FreshMerge.GapNoIndexFrontier }, FreshMerge.DescribeCoverageGaps(sweep));
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void NoIndexFrontier_SortsFirst_BecauseItIsTheWidestHole()
    {
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 2,
            FoldersFailed = 1,
            IndexFrontierMissing = true,
            ItemCappedFolders = new[] { "alice@example.com/Inbox" },
        };

        Assert.Equal(
            new[] { FreshMerge.GapNoIndexFrontier, FreshMerge.GapFoldersFailed, FreshMerge.GapItemCap },
            FreshMerge.DescribeCoverageGaps(sweep));
    }

    [Fact]
    public void NoIndexFrontier_SurvivesASweepThatCouldNotRun_BecauseItDescribesTheOtherTier()
    {
        // Two independent facts: the sweep did not run, AND the index holds nothing for this
        // scope. An answer missing both tiers has to say both, so the code is reported
        // alongside index-only rather than being swallowed by it.
        SweepInfo sweep = new SweepInfo { Performed = false, Error = "OutlookUnavailable", IndexFrontierMissing = true };

        Assert.Equal(new[] { FreshMerge.GapNoIndexFrontier }, FreshMerge.DescribeCoverageGaps(sweep));
        Assert.Equal(FreshMerge.FreshnessIndexOnly, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void ASweepThatWasNotNeeded_IsNoLongerLive_WhenTheIndexHasNothingForTheScope()
    {
        // "Not needed" asserts the INDEX already covers the requested window. Over a store
        // with no index rows that assertion is false, and the search would otherwise return
        // an empty list out of an unindexed store and call itself live. This is the one case
        // where notNeeded stops meaning complete.
        SweepInfo notNeeded = new SweepInfo { Performed = false, NotNeeded = true, IndexFrontierMissing = true };
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(notNeeded));

        // And with a frontier it is unchanged: still live, still no gaps.
        SweepInfo ordinary = new SweepInfo { Performed = false, NotNeeded = true };
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(ordinary));
        Assert.Null(FreshMerge.DescribeCoverageGaps(ordinary));
    }

    [Fact]
    public void TheNoIndexFrontierSentence_NamesTheStores_AndFallsBackToTheProfile()
    {
        SweepInfo named = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 4,
            IndexFrontierMissing = true,
            StoresWithoutIndex = new[] { "Archive 2019.pst", "Old mail.pst" },
            CoverageGaps = new[] { FreshMerge.GapNoIndexFrontier },
        };

        string line = Assert.Single(MailService.DescribeSweepCoverage(named, "12 minutes", folderScoped: false));
        Assert.Contains("Archive 2019.pst", line, StringComparison.Ordinal);
        Assert.Contains("Old mail.pst", line, StringComparison.Ordinal);

        // The unindexed-profile case knows the fact but has no catalog to name stores from.
        SweepInfo unnamed = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 4,
            IndexFrontierMissing = true,
            CoverageGaps = new[] { FreshMerge.GapNoIndexFrontier },
        };

        string profileLine = Assert.Single(MailService.DescribeSweepCoverage(unnamed, "12 minutes", folderScoped: false));
        Assert.Contains("this profile", profileLine, StringComparison.Ordinal);

        // The span it quotes is the constant, not a second copy of the number.
        Assert.Contains(
            MailService.EmptyIndexSweepWindow.TotalDays.ToString("F0", System.Globalization.CultureInfo.InvariantCulture)
                + " days",
            profileLine,
            StringComparison.Ordinal);
    }

    [Fact]
    public void ASweepWithAFrontier_NeverRaisesTheNoIndexCode()
    {
        // The flag is set from the frontier probe alone, so nothing else can raise this: a
        // code that fired on ordinary partial coverage would make the state meaningless.
        foreach ((string _, SweepInfo sweep) in CoverageHoles())
        {
            if (sweep.IndexFrontierMissing == true)
            {
                continue;
            }

            Assert.DoesNotContain(FreshMerge.GapNoIndexFrontier, FreshMerge.DescribeCoverageGaps(sweep)!);
        }
    }

    [Fact]
    public void EveryCoverageHole_MakesTheSweepPartial_AndNamesItself()
    {
        foreach ((string expectedGap, SweepInfo sweep) in CoverageHoles())
        {
            IReadOnlyList<string>? gaps = FreshMerge.DescribeCoverageGaps(sweep);
            Assert.True(gaps != null, $"{expectedGap}: a sweep with this hole must report coverage gaps");
            Assert.Contains(expectedGap, gaps!);

            // The whole point: the machine-readable pair must say partial, not "live".
            Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(sweep));
        }
    }

    [Fact]
    public void TheCoverageHoleSet_IsExactlyTheGapCodesDeclared_SoANewOneCannotBeAddedUntested()
    {
        List<string> covered = CoverageHoles().Select(row => row.Gap).OrderBy(c => c, StringComparer.Ordinal).ToList();
        List<string> declared = AllGapCodes().OrderBy(c => c, StringComparer.Ordinal).ToList();
        Assert.Equal(declared, covered);
    }

    [Fact]
    public void SkippedFolders_AreNotReportedTwice_WhenABoundStoppedTheWalk()
    {
        // A bound refuses the folders it did not reach and COUNTS them as skipped, so
        // reporting both would attribute a cap to unreadable folders. The bound's own code
        // still fires, so the answer is still partial - nothing is lost by the suppression.
        foreach (SweepInfo bounded in new[]
                 {
                     new SweepInfo { Performed = true, FoldersSwept = 40, FoldersSkipped = 12, FolderCapReached = true },
                     new SweepInfo { Performed = true, FoldersSwept = 5, FoldersSkipped = 12, TimeBudgetExceeded = true },
                     new SweepInfo { Performed = true, FoldersSwept = 5, FoldersSkipped = 12, DepthLimitReached = true },
                 })
        {
            IReadOnlyList<string> gaps = FreshMerge.DescribeCoverageGaps(bounded)!;
            Assert.DoesNotContain(FreshMerge.GapFoldersSkipped, gaps);
            Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(bounded));
        }
    }

    [Fact]
    public void SeveralHolesAtOnce_AreAllReported()
    {
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 40,
            FoldersFailed = 2,
            FolderCapReached = true,
            ItemCappedFolders = new[] { "alice@example.com/Inbox" },
        };

        IReadOnlyList<string> gaps = FreshMerge.DescribeCoverageGaps(sweep)!;
        Assert.Equal(
            new[] { FreshMerge.GapFoldersFailed, FreshMerge.GapFolderCap, FreshMerge.GapItemCap },
            gaps);
    }

    [Fact]
    public void GapCodesAndFreshnessValues_AreDistinctMachineReadableTokens()
    {
        IReadOnlyList<string> codes = AllGapCodes();
        Assert.Equal(codes.Count, codes.Distinct(StringComparer.Ordinal).Count());
        foreach (string code in codes)
        {
            Assert.DoesNotContain(" ", code, StringComparison.Ordinal);
            Assert.Equal(code.ToLowerInvariant(), code);
        }

        // The three freshness values stay distinct, and the two that already travelled on
        // the wire keep their exact spelling - callers pin them.
        Assert.Equal("live", FreshMerge.FreshnessLive);
        Assert.Equal("index-only", FreshMerge.FreshnessIndexOnly);
        Assert.Equal("partial", FreshMerge.FreshnessPartial);
    }

    // -------------------------------------------- every code earns one advice sentence

    [Fact]
    public void EveryGapCode_ProducesItsOwnAdviceSentence()
    {
        // Codes and prose are two renderings of one decision. A code with no sentence is a
        // partial result an agent can see but not explain to the user; a sentence with no
        // code is the original defect. This walks the codes declared on the type, so a new
        // one added without prose fails here rather than shipping silent.
        foreach (string code in AllGapCodes())
        {
            SweepInfo sweep = new SweepInfo
            {
                Performed = true,
                FoldersSwept = 4,
                FoldersFailed = 1,
                FoldersSkipped = 2,
                RowsUnreadable = 3,
                ItemsFilterUnreadable = 2,
                FiltersUnevaluated = new[] { "unread_only" },
                ItemCappedFolders = new[] { "alice@example.com/Inbox" },
                ItemsBodyCapped = 3,
                ItemsBodyCappedUnmatched = 2,
                CoverageGaps = new[] { code },
            };

            IReadOnlyList<string> advice = MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false);
            string line = Assert.Single(advice);
            Assert.StartsWith("Freshness sweep", line, StringComparison.Ordinal);
            Assert.DoesNotContain("no further detail available", line, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TheTotalMissSentence_IsLeftToTheFolderScopedCaller_WhichNamesTheFolder()
    {
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 0,
            CoverageGaps = new[] { FreshMerge.GapNothingSwept },
        };

        Assert.Empty(MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: true));
        Assert.Single(MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false));
    }

    [Fact]
    public void AnOmittedFolderList_IsReported_ButIsNotACoverageHole()
    {
        // The sweep covered those folders; only the LIST was dropped by its own cap. It
        // must be said (no silent caps) and it must not make a complete answer partial.
        SweepInfo sweep = new SweepInfo { Performed = true, FoldersSwept = 30, FolderListOmitted = true };
        sweep.CoverageGaps = FreshMerge.DescribeCoverageGaps(sweep);

        Assert.Null(sweep.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(sweep));
        Assert.Contains(
            MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false),
            line => line.Contains("swept-folder list is omitted", StringComparison.Ordinal));
    }
}
