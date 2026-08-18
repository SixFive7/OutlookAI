using System;
using System.Collections.Generic;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Gap B3, answered 2026-08-18: <b>the three search tiers admit the same item set, and item
/// class is not part of what any of them asks.</b>
/// <para>
/// The defect it closes was a product decision rather than a bug, which is why it needed an
/// answer: the index tier required <c>System.Kind</c> to contain <c>email</c>, the freshness
/// sweep filtered nothing, and the exhaustive COM scan filtered
/// <c>PR_MESSAGE_CLASS like 'IPM.Note%'</c>. So one query gave three different answers
/// depending on which engine reached the mail first, a meeting request the sweep returned
/// today vanished once it was indexed, and the mode that exists FOR completeness was the
/// only one that could not find a bounce report.
/// </para>
/// <para>
/// WHAT IS PROVEN HERE AND WHAT IS NOT. The rule and both of its payload renderings are pure
/// and pinned below, as is each tier's expression of it that can be read without Outlook:
/// the DASL the scan sends, the SQL the index statement carries, the row admission that runs
/// after it. What needs a real profile is that a meeting request or an NDR then actually
/// comes back from all three - that is T2's job, and this suite deliberately does not claim
/// it.
/// </para>
/// </summary>
public sealed class ItemClassAdmissionTests
{
    /// <summary>
    /// The rule itself. <c>Admits</c> cannot return false, and the test exists precisely
    /// because that is easy to "fix" by accident: a future narrowing has to delete a call
    /// site and this assertion, both of which say what is being given up.
    /// </summary>
    [Fact]
    public void EveryClassTheOldFiltersDropped_IsAdmitted()
    {
        Assert.NotEmpty(MailItemAdmission.ClassesTheOldFiltersDropped);
        foreach (string dropped in MailItemAdmission.ClassesTheOldFiltersDropped)
        {
            Assert.True(
                MailItemAdmission.Admits(dropped),
                dropped + " is mail a user asks about by name; no tier may filter it out again.");
        }

        // Including the classes nobody listed, which is the difference between this rule and
        // an allowlist: an allowlist admits what it knows, this admits what it meets.
        Assert.True(MailItemAdmission.Admits("IPM.Note"));
        Assert.True(MailItemAdmission.Admits("IPM.TaskRequest"));
        Assert.True(MailItemAdmission.Admits("IPM.Note.Microsoft.Voicemail.UM"));
        Assert.True(MailItemAdmission.Admits("IPM.Some.Class.Invented.In.2031"));
        Assert.True(MailItemAdmission.Admits(null));
        Assert.True(MailItemAdmission.Admits(string.Empty));
    }

    [Theory]
    [InlineData("IPM.Note", true)]
    [InlineData("ipm.note", true)]
    [InlineData("IPM.Note.SMIME", true)]
    [InlineData("IPM.Note.SMIME.MultipartSigned", true)]
    [InlineData("REPORT.IPM.Note.NDR", false)]
    [InlineData("IPM.Schedule.Meeting.Request", false)]
    [InlineData("IPM.Post", false)]
    [InlineData("IPM.Notification.Meeting", false)]
    [InlineData("IPM.Noteworthy", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void OrdinaryMail_IsPrefixMatchedAtAClassBoundary(string? messageClass, bool ordinary)
    {
        // The boundary matters twice over. Equality would report every signed mail as
        // something exotic; a bare StartsWith would swallow IPM.Notification.* and
        // IPM.Noteworthy, hiding exactly the items this reporting exists to name.
        Assert.Equal(ordinary, MailItemAdmission.IsOrdinaryMailClass(messageClass));
    }

    /// <summary>
    /// A class that could not be read is NOT ordinary mail: reporting an unidentified item
    /// as ordinary would be the payload asserting something nobody established. It is still
    /// admitted (above) - it just gets no claim made about it.
    /// </summary>
    [Fact]
    public void AnUnreadableClass_IsNeitherOrdinaryNorDescribed()
    {
        Assert.False(MailItemAdmission.IsOrdinaryMailClass(null));
        Assert.Null(MailItemAdmission.DescribeComItemClass(null));
        Assert.Null(MailItemAdmission.DescribeComItemClass("   "));
    }

    [Fact]
    public void ComHits_ReportTheirClassOnlyWhenItIsNotOrdinaryMail()
    {
        // Absent on the overwhelming majority of hits, so the field costs nothing and its
        // PRESENCE is the signal.
        Assert.Null(MailItemAdmission.DescribeComItemClass("IPM.Note"));
        Assert.Null(MailItemAdmission.DescribeComItemClass("IPM.Note.SMIME"));

        Assert.Equal("REPORT.IPM.Note.NDR", MailItemAdmission.DescribeComItemClass("REPORT.IPM.Note.NDR"));
        Assert.Equal(
            "IPM.Schedule.Meeting.Request",
            MailItemAdmission.DescribeComItemClass(" IPM.Schedule.Meeting.Request "));
    }

    [Fact]
    public void IndexHits_ReportAKindRatherThanAClass_BecauseTheIndexNeverOpensTheItem()
    {
        // The prefix is the honesty: this tier is guessing from System.Kind, and a bare
        // class name here would claim an authority it does not have.
        Assert.Equal("kind:calendar", MailItemAdmission.DescribeIndexRowClass(new[] { "calendar" }, false));
        Assert.Equal(
            "kind:calendar+document",
            MailItemAdmission.DescribeIndexRowClass(new[] { "calendar", "document" }, false));
        Assert.StartsWith(
            MailItemAdmission.IndexKindPrefix,
            MailItemAdmission.DescribeIndexRowClass(new[] { "contact" }, false),
            StringComparison.Ordinal);

        // Ordinary mail says nothing, whatever else the row is also tagged as.
        Assert.Null(MailItemAdmission.DescribeIndexRowClass(new[] { "email" }, false));
        Assert.Null(MailItemAdmission.DescribeIndexRowClass(new[] { "EMAIL" }, false));
        Assert.Null(MailItemAdmission.DescribeIndexRowClass(new[] { "document", "email" }, false));

        // An attachment-content row already says what it is via isAttachmentHit, and its
        // kind describes the ATTACHMENT rather than the mail carrying it.
        Assert.Null(MailItemAdmission.DescribeIndexRowClass(new[] { "picture" }, true));
    }

    [Fact]
    public void AnIndexRowWithNoKind_SaysSo_RatherThanSayingNothing()
    {
        // These rows used to be DROPPED, so silence about them would be the old filter
        // surviving as a blank field.
        Assert.Equal(MailItemAdmission.UnknownIndexKind, MailItemAdmission.DescribeIndexRowClass(null, false));
        Assert.Equal(
            MailItemAdmission.UnknownIndexKind,
            MailItemAdmission.DescribeIndexRowClass(Array.Empty<string>(), false));
        Assert.Equal(
            MailItemAdmission.UnknownIndexKind,
            MailItemAdmission.DescribeIndexRowClass(new[] { "  " }, false));
    }

    // ------------------------------------------------ the three tiers, one rule

    /// <summary>
    /// TIER 1 (index). The two shapes a search can ask for both admit a message row whatever
    /// its kind; the one shape that still tests a kind is store DISCOVERY, which no search
    /// reaches. Stated as one table so a re-narrowing of any search shape fails here with
    /// the reason attached.
    /// </summary>
    [Fact]
    public void IndexTier_AdmitsMessageRowsOfEveryClass_InEverySearchShape()
    {
        const string messageUrl = "mapi16://{SID}/alice@example.com($ab12)/0/Inbox/item";

        foreach (KindFilter shape in new[] { KindFilter.MessagesAndAttachments, KindFilter.MessagesOnly })
        {
            foreach (string kind in new[] { "email", "calendar", "document", "communication" })
            {
                Assert.True(
                    IndexRowFilter.Keep(MapRow(messageUrl, kind), shape),
                    shape + " must admit a message row of kind " + kind);
            }

            Assert.True(IndexRowFilter.Keep(MapRow(messageUrl), shape), shape + " must admit a kindless message row");
        }
    }

    /// <summary>
    /// TIER 3 (exhaustive scan). The DASL it sends carries no class predicate at all once it
    /// has anything else to restrict on - and where it has nothing, the predicate it emits
    /// matches every class rather than one.
    /// </summary>
    [Fact]
    public void ExhaustiveTier_SendsNoClassPredicate()
    {
        string bounded = ExhaustiveDaslFilter.Build(
            new[] { "factuur" },
            new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Utc),
            null,
            ExhaustiveEngine.CiPhraseMatch);

        Assert.DoesNotContain("IPM.", bounded, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("REPORT.", bounded, StringComparison.OrdinalIgnoreCase);

        string unbounded = ExhaustiveDaslFilter.Build(null, null, null, ExhaustiveEngine.Like);
        Assert.Contains(ExhaustiveDaslFilter.AdmitEveryClassClause, unbounded, StringComparison.Ordinal);
        Assert.EndsWith(" like '%'", ExhaustiveDaslFilter.AdmitEveryClassClause, StringComparison.Ordinal);
    }

    /// <summary>
    /// TIER 2 (freshness sweep) is the tier the other two were unified UP to: it has never
    /// filtered by class, so the pin that matters is that the snapshot it produces now
    /// CARRIES the class, which is what makes a widened result set legible.
    /// </summary>
    [Fact]
    public void SweepSnapshots_CarryTheMessageClass_SoAWidenedAnswerIsLegible()
    {
        ComMailBrief ndr = new ComMailBrief(
            "0000000000000000000000000000000000000000000000AB",
            "alice@example.com",
            null,
            "Inbox",
            "inbox",
            "Undeliverable: Invoice",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            "REPORT.IPM.Note.NDR");

        Assert.Equal("REPORT.IPM.Note.NDR", ndr.MessageClass);
        Assert.Equal("REPORT.IPM.Note.NDR", MailItemAdmission.DescribeComItemClass(ndr.MessageClass));
    }

    // ------------------------------------------------ what the answer says about it

    /// <summary>
    /// The widening announces itself in the one place it costs nothing: a sentence emitted
    /// only when the result set actually holds something that is not mail. An agent
    /// relaying "you have four mails about the invoice" over a list containing a delivery
    /// receipt is relaying something false, and until this existed nothing would have told
    /// it.
    /// </summary>
    [Fact]
    public void AnAnswerHoldingNonMail_SaysSo_AndNamesWhat()
    {
        string advice = MailService.DescribeNonMailHits(new[]
        {
            new HitSummary { Id = "h1" },
            new HitSummary { Id = "h2", ItemClass = "REPORT.IPM.Note.NDR" },
            new HitSummary { Id = "h3", ItemClass = "IPM.Schedule.Meeting.Request" },
            new HitSummary { Id = "h4", ItemClass = "kind:calendar" },
        })!;

        Assert.StartsWith("3 of these hits are not ordinary mail", advice, StringComparison.Ordinal);
        Assert.Contains("REPORT.IPM.Note.NDR", advice, StringComparison.Ordinal);
        Assert.Contains("IPM.Schedule.Meeting.Request", advice, StringComparison.Ordinal);
        Assert.Contains("kind:calendar", advice, StringComparison.Ordinal);
        Assert.Contains("itemClass", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOrdinaryAnswer_SaysNothing_SoTheSentenceCostsNothing()
    {
        Assert.Null(MailService.DescribeNonMailHits(null));
        Assert.Null(MailService.DescribeNonMailHits(Array.Empty<HitSummary>()));
        Assert.Null(MailService.DescribeNonMailHits(new[] { new HitSummary { Id = "h1" }, new HitSummary { Id = "h2" } }));
    }

    [Fact]
    public void TheClassListInTheSentence_IsCapped_AndSaysThatItIs()
    {
        // A cap in prose is still a cap. The per-hit itemClass fields are the complete
        // answer; this sentence is a heads-up.
        List<HitSummary> hits = new List<HitSummary>();
        for (int i = 0; i < MailService.NonMailClassAdviceCap + 3; i++)
        {
            hits.Add(new HitSummary { Id = "h" + i, ItemClass = "IPM.Class." + i });
        }

        string advice = MailService.DescribeNonMailHits(hits)!;

        Assert.StartsWith(
            (MailService.NonMailClassAdviceCap + 3).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " of these hits are not ordinary mail",
            advice,
            StringComparison.Ordinal);
        Assert.Contains(", ...)", advice, StringComparison.Ordinal);
        Assert.DoesNotContain("IPM.Class." + MailService.NonMailClassAdviceCap, advice, StringComparison.Ordinal);
    }

    [Fact]
    public void ExactlyTheCapManyClasses_DoesNotClaimThereAreMore()
    {
        // A "..." derived from "the list is full" would fire on exactly the cap, which is a
        // has-more flag that lies - the shape this payload discipline exists to prevent.
        List<HitSummary> hits = new List<HitSummary>();
        for (int i = 0; i < MailService.NonMailClassAdviceCap; i++)
        {
            hits.Add(new HitSummary { Id = "h" + i, ItemClass = "IPM.Class." + i });
        }

        string advice = MailService.DescribeNonMailHits(hits)!;

        Assert.DoesNotContain(", ...", advice, StringComparison.Ordinal);
        Assert.Contains("IPM.Class." + (MailService.NonMailClassAdviceCap - 1), advice, StringComparison.Ordinal);
    }

    [Fact]
    public void RepeatedClasses_AreNamedOnce_ButCountedEveryTime()
    {
        string advice = MailService.DescribeNonMailHits(new[]
        {
            new HitSummary { Id = "h1", ItemClass = "REPORT.IPM.Note.NDR" },
            new HitSummary { Id = "h2", ItemClass = "REPORT.IPM.Note.NDR" },
            new HitSummary { Id = "h3", ItemClass = "REPORT.IPM.Note.NDR" },
        })!;

        Assert.StartsWith("3 of these hits are not ordinary mail (REPORT.IPM.Note.NDR)", advice, StringComparison.Ordinal);
        Assert.DoesNotContain(", ...", advice, StringComparison.Ordinal);
    }

    private static IndexHit MapRow(string url, params string[] kinds)
    {
        Dictionary<string, object?> row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = url,
        };

        if (kinds.Length > 0)
        {
            row["System.Kind"] = kinds;
        }

        return IndexRowMapper.Map(row);
    }
}
