using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what the census says when the only thing wrong with a store is PROBE RESIDUE, and
/// shows that what it says is indistinguishable from a real defect.
///
/// <para>
/// <b>The observation these tests explain.</b> A 20,000-item corpus built on the test guest
/// on 2026-09-16 produced a clean manifest (20,000 distinct EntryIDs, last ordinal 20,000)
/// and a <c>corpus-census</c> that reported three FAULTS: one item in Drafts, twelve in a
/// folder the plan does not name, and <b>"1 ordinal(s) exist more than once, so the corpus
/// holds more items than the plan describes and every per-item number measured against it is
/// wrong"</b>. That last sentence is the one that matters, and it is false here: not one
/// corpus ordinal exists twice. The thirteen extra items are the throwaway items the two
/// probes create, and every one of them carries <see cref="CorpusPlan.ProbeOrdinal"/>
/// (<see cref="int.MaxValue"/>) in its subject - so the scan sees thirteen sightings of ONE
/// ordinal, and <see cref="CorpusCensus.Compare"/> counts that as a duplicated ordinal.
/// </para>
///
/// <para>
/// <b>These are pure</b> - no Outlook, no COM, no mailbox, no settings file, and no
/// <c>Category=Live</c> trait. They are arithmetic over <see cref="CorpusCensus.Compare"/>
/// and <see cref="CorpusPlan.ClassifySubject"/>.
/// </para>
/// </summary>
public sealed class CorpusProbeResidueCensusTests
{
    // The guest's own corpus parameters, from Docs/corpus-measurement-plan.md and
    // Testbed/testbed.json: --corpus-id vm2 --seed 7777 --anchor 2026-08-19 --count 20000.
    private const string CorpusId = "vm2";
    private const int ItemCount = 20_000;
    private static readonly DateTime Anchor = new(2026, 8, 19, 0, 0, 0, DateTimeKind.Utc);

    private const int DeletedItemsFolderId = 3;
    private const int DraftsFolderId = 16;

    private static CorpusPlan Plan() => new(new CorpusPlanOptions(CorpusId, 7777, Anchor));

    /// <summary>
    /// The exact census line the guest printed on 2026-09-16, reproduced from nothing but the
    /// plan plus thirteen probe-ordinal sightings: twelve in Deleted Items (probe items whose
    /// <c>MailItem.Delete()</c> soft-deleted them there) and one in Drafts (a probe item whose
    /// delete never ran).
    /// </summary>
    [Fact]
    public void ThirteenProbeItemsReproduceTheGuestsCensusVerbatim()
    {
        CorpusPlan plan = Plan();
        var sightings = new List<CorpusSighting>(ItemCount + 13);
        for (int ordinal = 1; ordinal <= ItemCount; ordinal++)
        {
            sightings.Add(new CorpusSighting(ordinal, plan.Describe(ordinal).FolderId));
        }

        for (int i = 0; i < 12; i++)
        {
            sightings.Add(new CorpusSighting(CorpusPlan.ProbeOrdinal, DeletedItemsFolderId));
        }

        sightings.Add(new CorpusSighting(CorpusPlan.ProbeOrdinal, DraftsFolderId));

        CorpusCensusReport report = CorpusCensus.Compare(plan, ItemCount, sightings);
        (bool clean, string message) = CorpusCensus.Decide(report);

        Assert.False(clean);
        Assert.Equal(
            "Census: 20,013 item(s) found for 20,000 planned; per folder found/planned: "
            + "Deleted Items=2,473/2,461, Sent Items=4,964/4,964, Inbox=10,912/10,912, Drafts=1/0, "
            + "Junk Email=1,663/1,663. FAULTS: 1 item(s) are in DRAFTS, which the freshness sweep does not "
            + "cover - those items are invisible to the measurement this corpus exists for; 1 ordinal(s) "
            + "exist more than once, so the corpus holds more items than the plan describes and every "
            + "per-item number measured against it is wrong; 12 item(s) are in a folder the plan does not "
            + "put them in.",
            message);

        // The claim inside that message, tested directly: every corpus ordinal exists exactly
        // once. The "duplicate" is the reserved probe ordinal and nothing else.
        Assert.Equal(0, report.MissingOrdinals);
        Assert.Equal(ItemCount + 1, report.DistinctOrdinals);
        Assert.Equal(1, report.DuplicatedOrdinals);
        Assert.Equal(13, report.Misplaced);
    }

    /// <summary>
    /// The defect, stated as a test: a store whose ONLY flaw is probe residue and a store
    /// holding a genuinely double-created corpus item produce the SAME duplicate-ordinal
    /// fault, so the sentence cannot be used to tell them apart. A resumed build is the way
    /// the second one happens - an item created but not yet flushed to the manifest is
    /// re-created by the next run.
    /// </summary>
    [Fact]
    public void ARealDoubleCreateAndProbeResidueRaiseTheIdenticalFault()
    {
        CorpusPlan plan = Plan();
        var planned = new List<CorpusSighting>(ItemCount);
        for (int ordinal = 1; ordinal <= ItemCount; ordinal++)
        {
            planned.Add(new CorpusSighting(ordinal, plan.Describe(ordinal).FolderId));
        }

        var residue = new List<CorpusSighting>(planned)
        {
            new(CorpusPlan.ProbeOrdinal, DeletedItemsFolderId),
            new(CorpusPlan.ProbeOrdinal, DeletedItemsFolderId),
        };

        // Ordinal 7,431 built twice, both copies where the plan puts it: the shape a resumed
        // build leaves when the process died between the COM create and the manifest flush.
        var doubleCreated = new List<CorpusSighting>(planned)
        {
            new(7_431, plan.Describe(7_431).FolderId),
        };

        CorpusCensusReport residueReport = CorpusCensus.Compare(plan, ItemCount, residue);
        CorpusCensusReport doubleReport = CorpusCensus.Compare(plan, ItemCount, doubleCreated);

        Assert.Equal(1, residueReport.DuplicatedOrdinals);
        Assert.Equal(1, doubleReport.DuplicatedOrdinals);

        const string fault = "1 ordinal(s) exist more than once, so the corpus holds more items than the plan "
            + "describes and every per-item number measured against it is wrong";
        Assert.Contains(fault, CorpusCensus.Decide(residueReport).Message, StringComparison.Ordinal);
        Assert.Contains(fault, CorpusCensus.Decide(doubleReport).Message, StringComparison.Ordinal);

        // And only one of the two is actually true. The residue corpus holds 20,000 corpus
        // items, one per ordinal; the double-created one holds 20,001 in 20,000 ordinals.
        Assert.Equal(ItemCount, residueReport.Sightings - 2);
        Assert.Equal(ItemCount + 1, doubleReport.Sightings);
    }

    /// <summary>
    /// Why the scan counts probe items at all: a probe subject carries the CURRENT corpus tag
    /// and the corpus id, so <see cref="CorpusPlan.ClassifySubject"/> accepts it and hands
    /// back <see cref="CorpusPlan.ProbeOrdinal"/> - which is <see cref="int.MaxValue"/>, a
    /// perfectly legal ordinal as far as every predicate downstream is concerned.
    /// </summary>
    [Fact]
    public void AProbeSubjectParsesAsThisCorpusAndYieldsTheReservedOrdinal()
    {
        // Built exactly as ComCorpusMailbox.ProbeSubject builds it.
        string subject = CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + CorpusId + "#"
            + CorpusPlan.ProbeOrdinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture)
            + "] placement InPlaceOnly";

        Assert.Equal(CorpusSubjectKind.Current, CorpusPlan.ClassifySubject(subject, CorpusId, out int ordinal));
        Assert.Equal(CorpusPlan.ProbeOrdinal, ordinal);
        Assert.True(CorpusPlan.TryParseOrdinal(subject, CorpusId, out _));

        // Which is also what lets teardown's second phase delete them: the two-key rule needs
        // the subject to parse, and it does.
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABCD" });
        Assert.True(CorpusSafety.MayDelete("ABCD", subject, allowlist, CorpusId));

        // The ordinal is far outside any plan, so the census can never match it to a folder -
        // every probe sighting is counted as misplaced, whichever folder it is found in.
        Assert.True(CorpusPlan.ProbeOrdinal > ItemCount);
    }

    /// <summary>
    /// The probe ordinal renders to TEN digits, not seven, and the DASL fragment the probes
    /// select themselves with renders it the same way - so the fragment a probe searches for
    /// is the fragment its own subject carries. A "D7" that truncated would break both.
    /// </summary>
    [Fact]
    public void TheProbeOrdinalRendersIdenticallyInSubjectAndDaslFragment()
    {
        string rendered = CorpusPlan.ProbeOrdinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal("2147483647", rendered);
        Assert.Contains(
            rendered,
            CorpusPlan.DaslSubjectFragment(CorpusId, CorpusPlan.ProbeOrdinal),
            StringComparison.Ordinal);
    }
}
