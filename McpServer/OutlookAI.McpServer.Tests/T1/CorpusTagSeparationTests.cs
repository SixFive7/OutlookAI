using System;
using System.Collections.Generic;
using System.IO;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The measurement corpus and the live tier's test artifacts carry DIFFERENT subject tags,
/// and this file is the mechanism that keeps it that way.
///
/// <para>
/// <b>The defect it closes.</b> <c>CorpusPlan.SubjectTag</c> was <c>RemediationRules.SubjectTag</c>
/// - literally the same constant - and <c>CorpusPlan.BuildSubject</c> put it at the front of
/// every corpus subject. The live tier's post-run artifact sweep walks every store in
/// <c>expectedStoreDisplayNames</c>, counts subjects containing that tag and calls
/// <c>DeleteTaggedArtifactsUntilStableZero</c> on anything above zero; its folder set covers
/// Inbox, Sent Items, Deleted Items and the Outbox - four of the corpus's five populated
/// folders, ~21 000 items on the documented layout. So a live run was one sweep away from
/// deleting the entire measurement corpus, through the tested helpers, inside the safety
/// rules, with nothing to stop it. Declaring the corpus store a bystander turned that into a
/// write-guard refusal - safe, but a run whose normal outcome is a refusal gets muted.
/// </para>
///
/// <para>
/// <b>Why a test and not a comment.</b> The two constants live in the same assembly and
/// nothing about the type system stops one being assigned from the other again; the audit
/// behind <c>Docs/magic-numbers.md</c> found a "keep these in step" comment that had already
/// become false. A test fails the build.
/// </para>
///
/// <para>
/// <b>And why it tests the FRAGMENTS, not just the tags.</b> The sweep does not match the
/// bracketed tag. It matches the bracket-free text, in a DASL <c>LIKE '%OutlookAI-McpTest%'</c>
/// prefilter and then in a <c>Contains</c> - because a <c>[</c> inside a LIKE pattern opens a
/// character class, which is the mechanism that destroyed real mail in the 2026-07-25 incident.
/// A test that only compared <c>"[OutlookAI-Corpus]" != "[OutlookAI-McpTest]"</c> would pass
/// for a corpus tag of <c>"[X-OutlookAI-McpTest-Corpus]"</c>, which the sweep would still eat.
/// </para>
/// </summary>
public sealed class CorpusTagSeparationTests
{
    private const string CorpusId = "vm1";
    private static readonly DateTime Anchor = new(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc);

    /// <summary>
    /// The bracket-free text the live sweep matches on, as a LITERAL rather than as
    /// <c>RemediationRules.DaslCountFragment</c>. The sweep's own call sites
    /// (<c>LiveDraftTests</c> / <c>LiveSendTests</c> <c>ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain</c>,
    /// via <c>LiveOutlookTestMailer.CountTaggedArtifacts</c>) pass this string spelled out, so
    /// pinning the constant instead would leave the real matcher unguarded.
    /// </summary>
    private const string SweepMatchText = "OutlookAI-McpTest";

    private static CorpusPlan Plan(string corpusId = CorpusId)
        => new(new CorpusPlanOptions(corpusId, 4242, Anchor));

    // ------------------------------------------------------------------ the two tags differ

    [Fact]
    public void TheCorpusTagAndTheArtifactTagAreDifferentStrings()
    {
        Assert.NotEqual(RemediationRules.SubjectTag, CorpusPlan.SubjectTag);

        // Neither may CONTAIN the other either. Equality alone would still permit
        // "[OutlookAI-McpTest-Corpus]", which an ordinal Contains on the artifact tag matches.
        Assert.DoesNotContain(RemediationRules.SubjectTag, CorpusPlan.SubjectTag, StringComparison.Ordinal);
        Assert.DoesNotContain(CorpusPlan.SubjectTag, RemediationRules.SubjectTag, StringComparison.Ordinal);
    }

    [Fact]
    public void NeitherTagCarriesTheOthersBracketFreeFragment()
    {
        // This is the property the sweep actually depends on - see the class remarks.
        Assert.DoesNotContain(RemediationRules.DaslCountFragment, CorpusPlan.SubjectTag, StringComparison.Ordinal);
        Assert.DoesNotContain(RemediationRules.DaslCountFragment, CorpusPlan.CorpusTagOpen, StringComparison.Ordinal);
        Assert.DoesNotContain(CorpusPlan.DaslCountFragment, RemediationRules.SubjectTag, StringComparison.Ordinal);
        Assert.NotEqual(RemediationRules.DaslCountFragment, CorpusPlan.DaslCountFragment);
    }

    [Fact]
    public void NoCorpusSubjectContainsTheTextTheArtifactSweepMatches()
    {
        CorpusPlan plan = Plan();
        foreach (int ordinal in new[] { 1, 2, 7, 99, 1_234, 20_000, CorpusPlan.ProbeOrdinal })
        {
            string subject = plan.BuildSubject(ordinal);
            Assert.DoesNotContain(SweepMatchText, subject, StringComparison.Ordinal);
            Assert.DoesNotContain(SweepMatchText, subject, StringComparison.OrdinalIgnoreCase);
            Assert.False(RemediationRules.IsTagged(subject), $"ordinal {ordinal} would be swept as an artifact");
        }
    }

    [Fact]
    public void NoCorpusSubjectContainsTheTextTheArtifactSweepMatchesForAnyCorpusId()
    {
        // The id is operator-supplied and lands inside the subject. It is restricted to ASCII
        // letters, digits, '-' and '_', which cannot spell the artifact fragment on their own -
        // but an id may legitimately contain "McpTest", so this pins the tag, not the id.
        foreach (string id in new[] { "vm1", "corpus-b", "a_b", "McpTest", "OutlookAI" })
        {
            string subject = Plan(id).BuildSubject(1);
            Assert.False(
                RemediationRules.IsTagged(subject),
                $"corpus id '{id}' produced a subject the artifact sweep would delete");
        }
    }

    [Fact]
    public void AnArtifactSubjectIsNotACorpusItem()
    {
        // The other direction. An artifact carrying its own tag must never parse as a corpus
        // item, or corpus-teardown could delete the live tier's leftovers - a store it was
        // never pointed at, deleting items no corpus manifest recorded.
        string artifact = RemediationRules.SubjectTag + " r7-draft-seed";
        Assert.False(CorpusPlan.TryParseOrdinal(artifact, CorpusId, out _));
        Assert.Equal(CorpusSubjectKind.NotCorpus, CorpusPlan.ClassifySubject(artifact, CorpusId, out _));
    }

    // ------------------------------------------------------------------ the load-bearing paths

    [Fact]
    public void TheDeleteAndRewriteGuardsFollowTheNewTag()
    {
        // corpus-teardown selects by EntryID allowlist AND ordinal tag match, both required
        // (CLAUDE.md mailbox-safety rule 2). Both keys are re-checked here against the CURRENT
        // tag, because a tag change that missed either turns a safety guard into a no-op.
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ID-1" });
        string subject = Plan().BuildSubject(42);

        Assert.True(CorpusSafety.MayDelete("ID-1", subject, allowlist, CorpusId));
        Assert.True(CorpusSafety.MayRewrite("ID-1", subject, allowlist, CorpusId, 42));

        // Same item, old tag: both refuse.
        string legacy = LegacySubject(42);
        Assert.False(CorpusSafety.MayDelete("ID-1", legacy, allowlist, CorpusId));
        Assert.False(CorpusSafety.MayRewrite("ID-1", legacy, allowlist, CorpusId, 42));

        // And the artifact tag on its own is not a corpus key at all.
        Assert.False(CorpusSafety.MayDelete("ID-1", RemediationRules.SubjectTag + " x", allowlist, CorpusId));
        Assert.False(CorpusSafety.MayRewrite("ID-1", RemediationRules.SubjectTag + " x", allowlist, CorpusId, 42));
    }

    [Fact]
    public void TheCensusAndProbeFiltersStillSelectACorpusSubject()
    {
        // The census walks a GetTable LIKE prefilter on DaslCountFragment and the placement
        // and date probes narrow to DaslSubjectFragment. Both are supersets by construction -
        // this is the assertion that they are still supersets of what BuildSubject writes.
        CorpusPlan plan = Plan();
        string subject = plan.BuildSubject(1_234);
        Assert.Contains(CorpusPlan.DaslCountFragment, subject, StringComparison.Ordinal);
        Assert.Contains(CorpusPlan.DaslSubjectFragment(CorpusId, 1_234), subject, StringComparison.Ordinal);

        // The probe item is deletable by the same two keys as a corpus item, so its subject
        // has to carry the current tag too.
        string probe = CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + CorpusId + "#"
            + CorpusPlan.ProbeOrdinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture)
            + "] placement probe";
        Assert.True(CorpusPlan.TryParseOrdinal(probe, CorpusId, out int probeOrdinal));
        Assert.Equal(CorpusPlan.ProbeOrdinal, probeOrdinal);
        Assert.False(RemediationRules.IsTagged(probe));
    }

    // ------------------------------------------------------------------ old-tagged corpora

    [Fact]
    public void AnOldTaggedCorpusIsRecognisedRatherThanUnmatched()
    {
        // The point of the Legacy value: without it an old corpus is simply not found, and a
        // census reports every ordinal missing - which reads as "the build never ran".
        Assert.Equal(
            CorpusSubjectKind.Legacy,
            CorpusPlan.ClassifySubject(LegacySubject(7), CorpusId, out int ordinal));
        Assert.Equal(7, ordinal);
    }

    [Fact]
    public void AnOldTaggedCorpusIsStillRefusedByEveryWritePredicate()
    {
        // Recognised is not the same as addressable. Nothing in this build may delete or
        // rewrite a legacy item, which is what makes the refusal safe.
        Assert.False(CorpusPlan.TryParseOrdinal(LegacySubject(7), CorpusId, out _));
    }

    [Fact]
    public void AnOldTaggedItemOfAnotherCorpusIsNotThisCorpusAtAll()
    {
        Assert.Equal(
            CorpusSubjectKind.NotCorpus,
            CorpusPlan.ClassifySubject(LegacySubject(7, "vm2"), CorpusId, out _));
    }

    [Fact]
    public void CarryingBothTagsCountsAsCurrent()
    {
        // A mail client, or a half-finished migration, could leave both on one item. The
        // current tag is the one this build wrote, and a later arrival does not take that
        // away - so the item stays deletable by its own manifest.
        string both = RemediationRules.SubjectTag + Plan().BuildSubject(7);
        Assert.Equal(CorpusSubjectKind.Current, CorpusPlan.ClassifySubject(both, CorpusId, out int ordinal));
        Assert.Equal(7, ordinal);
        Assert.True(CorpusPlan.TryParseOrdinal(both, CorpusId, out _));
    }

    [Fact]
    public void TheLegacyTagIsAFrozenLiteralNotAliasedToTheCurrentOne()
    {
        // If these were ever made equal the legacy branch would become unreachable and an old
        // corpus would go back to being silently unmatched.
        Assert.NotEqual(CorpusPlan.SubjectTag, CorpusPlan.LegacySubjectTag);
        Assert.Equal("[OutlookAI-McpTest]", CorpusPlan.LegacySubjectTag);
    }

    [Fact]
    public void TheCensusSaysWhatAnOldTaggedCorpusIsRatherThanCallingItMissing()
    {
        CorpusPlan plan = Plan();
        CorpusCensusReport report = CorpusCensus.Compare(
            plan, 100, Array.Empty<CorpusSighting>(), legacyTagged: 100);
        (bool clean, string message) = CorpusCensus.Decide(report);

        Assert.False(clean);
        Assert.Equal(100, report.LegacyTagged);
        Assert.Contains(CorpusPlan.LegacySubjectTag, message, StringComparison.Ordinal);
        Assert.Contains("REBUILD", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheLegacyRefusalNamesBothTagsAndTheRemedy()
    {
        string refusal = CorpusCommands.LegacyCorpusRefusal(20_000);
        Assert.Contains(CorpusPlan.LegacySubjectTag, refusal, StringComparison.Ordinal);
        Assert.Contains(CorpusPlan.SubjectTag, refusal, StringComparison.Ordinal);
        Assert.Contains("20,000", refusal, StringComparison.Ordinal);
        Assert.Contains(".pst", refusal, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ re-anchoring is retired

    [Fact]
    public void ReanchorRefusesAndSaysRebuildingIsTheMaintenancePath()
    {
        // Decision 2, 2026-08-25. The notice is in the TOOL, not only in the docs - and it is
        // printed before every argument check, so an operator who also mistyped an argument is
        // still told the command is retired rather than being sent to fix the typo.
        using var writer = new StringWriter();
        int exit = CorpusCommands.RunReanchor(CorpusOptions.Parse(Array.Empty<string>()), writer);

        Assert.Equal(1, exit);
        string printed = writer.ToString();
        Assert.Contains("RETIRED", printed, StringComparison.Ordinal);
        Assert.Contains("Rebuilding is the supported way", printed, StringComparison.Ordinal);
        Assert.Contains("13m25s", printed, StringComparison.Ordinal);
        Assert.Contains("THROWAWAY ITEMS", printed, StringComparison.Ordinal);
    }

    [Fact]
    public void ReanchorRefusesEvenWhenEveryArgumentIsPresent()
    {
        // No COM is reached: the refusal is ahead of Vet(), which is what makes this testable
        // from CI at all.
        using var writer = new StringWriter();
        int exit = CorpusCommands.RunReanchor(
            CorpusOptions.Parse(new[]
            {
                "--store", "Corpus A", "--allow-store", "Corpus A", "--corpus-id", CorpusId,
                "--seed", "4242", "--anchor", "2026-08-01", "--count", "20000",
                "--manifest", @"D:\corpus\vm1.jsonl", "--to", "now", "--execute",
            }),
            writer);

        Assert.Equal(1, exit);
        Assert.Contains("RETIRED", writer.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheDiagnosticEscapeHatchStillPrintsTheNotice()
    {
        // Kept so the write path can be diagnosed, which is the stated condition for
        // un-retiring the command. It fails on --count long before COM, which is enough to
        // prove the notice is printed on the way past.
        CorpusOptions options = CorpusOptions.Parse(new[] { "--diagnose-write-path" });
        Assert.True(options.DiagnoseWritePath);

        using var writer = new StringWriter();
        Assert.Throws<ArgumentException>(() => CorpusCommands.RunReanchor(options, writer));
        Assert.Contains("DIAGNOSIS ONLY", writer.ToString(), StringComparison.Ordinal);
    }

    private static string LegacySubject(int ordinal, string corpusId = CorpusId)
        => CorpusPlan.LegacySubjectTag + CorpusPlan.CorpusTagOpen + corpusId + "#"
            + ordinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture)
            + "] renewal invoice handover";
}
