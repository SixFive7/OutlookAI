using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what the census says when the only thing wrong with a store is PROBE RESIDUE, and
/// pins the predicate that selects that residue for deletion.
///
/// <para>
/// <b>The observation these tests explain.</b> A 20,000-item corpus built on the test guest
/// on 2026-09-16 produced a clean manifest (20,000 distinct EntryIDs, last ordinal 20,000)
/// and a <c>corpus-census</c> that reported three FAULTS: one item in Drafts, twelve in a
/// folder the plan does not name, and <b>"1 ordinal(s) exist more than once, so the corpus
/// holds more items than the plan describes and every per-item number measured against it is
/// wrong"</b>. That last sentence is the one that matters, and it was false: not one corpus
/// ordinal existed twice. The thirteen extra items were the throwaway items the two probes
/// create, and every one of them carries <see cref="CorpusPlan.ProbeOrdinal"/>
/// (<see cref="int.MaxValue"/>) in its subject - so the scan saw thirteen sightings of ONE
/// ordinal, and <see cref="CorpusCensus.Compare"/> counted that as a duplicated ordinal.
/// </para>
///
/// <para>
/// <b>What changed on 2026-09-16, and why these tests changed with it.</b> Two of the tests
/// below used to pin the DEFECT - one reproduced the guest's census line verbatim
/// (<c>ThirteenProbeItemsReproduceTheGuestsCensusVerbatim</c>), the other showed that probe
/// residue and a genuine double-create raise the identical fault
/// (<c>ARealDoubleCreateAndProbeResidueRaiseTheIdenticalFault</c>). Both now pin the fix: the
/// census counts probe items in their own number, excludes them from every corpus statistic,
/// and gives them their own sentence. The guest's measured line is kept below, in the test
/// that used to assert it, because it is evidence and nothing else records it.
/// </para>
///
/// <para>
/// <b>These are pure</b> - no Outlook, no COM, no mailbox, no settings file, and no
/// <c>Category=Live</c> trait. They are arithmetic over <see cref="CorpusCensus.Compare"/>,
/// <see cref="CorpusPlan.ClassifySubject"/> and
/// <see cref="ComCorpusMailbox.SelectProbeResidue"/>. That last one is the point of the
/// residue sweep's design: the sweep is a DELETE, and the predicate that decides which items
/// it addresses is a pure function over scan rows precisely so a machine with no mailbox can
/// pin it.
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
    /// The guest's thirteen probe items, put through the census again - and now reported as
    /// thirteen items of litter rather than as a corrupt corpus.
    ///
    /// <para>
    /// <b>The string this test used to assert</b>, printed by the guest on 2026-09-16 and
    /// reproduced here from nothing but the plan plus these same thirteen sightings. It is
    /// kept verbatim because it is the measured evidence, and because every clause of it that
    /// is wrong is wrong in a way the new line fixes:
    /// </para>
    /// <code>
    /// Census: 20,013 item(s) found for 20,000 planned; per folder found/planned: Deleted
    /// Items=2,473/2,461, Sent Items=4,964/4,964, Inbox=10,912/10,912, Drafts=1/0, Junk
    /// Email=1,663/1,663. FAULTS: 1 item(s) are in DRAFTS, which the freshness sweep does not
    /// cover - those items are invisible to the measurement this corpus exists for; 1
    /// ordinal(s) exist more than once, so the corpus holds more items than the plan describes
    /// and every per-item number measured against it is wrong; 12 item(s) are in a folder the
    /// plan does not put them in.
    /// </code>
    /// <para>
    /// Twelve of the thirteen were in Deleted Items - probe items whose <c>MailItem.Delete()</c>
    /// soft-deleted them there - and one was in Drafts, a probe item whose delete never ran.
    /// </para>
    /// </summary>
    [Fact]
    public void ThirteenProbeItemsAreCountedAsLitterAndNotAsACorruptCorpus()
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

        // Still not clean - thirteen items that nothing owns are still thirteen items that
        // nothing owns - but every number in the line is now true, and the fault names what
        // they actually are.
        Assert.False(clean);
        Assert.Equal(
            "Census: 20,000 item(s) found for 20,000 planned; per folder found/planned: "
            + "Deleted Items=2,461/2,461, Sent Items=4,964/4,964, Inbox=10,912/10,912, "
            + "Junk Email=1,663/1,663. FAULTS: 13 throwaway probe item(s) were left behind - the items the "
            + "placement and date probes create and delete, whose SOFT delete leaves them sitting in Deleted "
            + "Items under an EntryID no manifest records. They are not part of the corpus and every count in "
            + "this line excludes them. The probes now purge their own residue, at the start of each pass and "
            + "after each item, so a non-zero count here means a store built before 2026-09-16, or a probe "
            + "killed mid-run; re-running corpus-probe --execute clears it.",
            message);

        // What the old line got wrong, clause by clause.
        Assert.Equal(13, report.ProbeItems);
        Assert.Equal(ItemCount, report.Sightings);          // was 20,013
        Assert.Equal(ItemCount, report.DistinctOrdinals);   // was 20,001
        Assert.Equal(0, report.DuplicatedOrdinals);         // was 1, and it was false
        Assert.Equal(0, report.Misplaced);                  // was 13
        Assert.Equal(0, report.StrayDrafts);                // was 1 - it was a probe item
        Assert.Equal(0, report.MissingOrdinals);
    }

    /// <summary>
    /// A store whose only flaw is probe residue and a store holding a genuinely double-created
    /// corpus item now produce DIFFERENT faults, and the difference is the whole point: one is
    /// litter the next probe pass sweeps up, the other means every per-item number measured
    /// against the corpus is wrong.
    ///
    /// <para>
    /// <b>This test used to assert the opposite.</b> As
    /// <c>ARealDoubleCreateAndProbeResidueRaiseTheIdenticalFault</c> it pinned the defect - the
    /// two shapes raising one sentence, so an operator could not tell them apart - and it was
    /// written that way deliberately, to make the defect fail a test rather than live in a
    /// note. It is rewritten rather than deleted so the history stays attached to the code it
    /// is about.
    /// </para>
    /// <para>
    /// A resumed build is how the second shape happens: an item created but not yet flushed to
    /// the manifest is re-created by the next run, leaving an orphan copy outside the manifest
    /// that nothing can ever delete by id. That is the case the duplicate-ordinal sentence was
    /// written for, and it now says only that.
    /// </para>
    /// </summary>
    [Fact]
    public void ARealDoubleCreateAndProbeResidueNowRaiseDifferentFaults()
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

        Assert.Equal(0, residueReport.DuplicatedOrdinals);
        Assert.Equal(2, residueReport.ProbeItems);
        Assert.Equal(1, doubleReport.DuplicatedOrdinals);
        Assert.Equal(0, doubleReport.ProbeItems);

        const string duplicateFault = "1 ordinal(s) exist more than once, so the corpus holds more items than the "
            + "plan describes and every per-item number measured against it is wrong";
        const string probeFault = "2 throwaway probe item(s) were left behind";

        string residueMessage = CorpusCensus.Decide(residueReport).Message;
        string doubleMessage = CorpusCensus.Decide(doubleReport).Message;

        Assert.Contains(probeFault, residueMessage, StringComparison.Ordinal);
        Assert.DoesNotContain(duplicateFault, residueMessage, StringComparison.Ordinal);

        Assert.Contains(duplicateFault, doubleMessage, StringComparison.Ordinal);
        Assert.DoesNotContain("throwaway probe item", doubleMessage, StringComparison.Ordinal);

        // And the numbers now match the reality each message describes. The residue store
        // holds 20,000 corpus items, one per ordinal; the double-created one holds 20,001 in
        // 20,000 ordinals.
        Assert.Equal(ItemCount, residueReport.Sightings);
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
        string subject = ProbeSubject("placement InPlaceOnly");

        Assert.Equal(CorpusSubjectKind.Current, CorpusPlan.ClassifySubject(subject, CorpusId, out int ordinal));
        Assert.Equal(CorpusPlan.ProbeOrdinal, ordinal);
        Assert.True(CorpusPlan.TryParseOrdinal(subject, CorpusId, out _));

        // Which is also what lets teardown's second phase - and the probes' own residue sweep -
        // delete them: the two-key rule needs the subject to parse, and it does.
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ABCD" });
        Assert.True(CorpusSafety.MayDelete("ABCD", subject, allowlist, CorpusId));

        // The ordinal is far outside any plan, so the census can never match it to a folder.
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

    // ------------------------------------------------------------- the residue sweep's predicate

    /// <summary>
    /// The sweep selects exactly the probe rows out of a mixed scan, and nothing else. This is
    /// the predicate a DELETE is aimed with, so it is pinned by identity - the selected rows
    /// are compared as rows, not counted.
    /// </summary>
    [Fact]
    public void TheSweepSelectsEveryProbeRowAndOnlyProbeRows()
    {
        var probeInInbox = new ComCorpusMailbox.ScanRow(CorpusPlan.ProbeOrdinal, "P-INBOX", 6);
        var probeInDrafts = new ComCorpusMailbox.ScanRow(CorpusPlan.ProbeOrdinal, "P-DRAFTS", DraftsFolderId);
        var probeInDeleted = new ComCorpusMailbox.ScanRow(CorpusPlan.ProbeOrdinal, "P-DELETED", DeletedItemsFolderId);

        var rows = new List<ComCorpusMailbox.ScanRow>
        {
            new(1, "C-FIRST", 6),
            probeInInbox,
            new(7_431, "C-MIDDLE", 5),
            probeInDrafts,
            new(ItemCount, "C-LAST", 23),
            probeInDeleted,
        };

        IReadOnlyList<ComCorpusMailbox.ScanRow> selected = ComCorpusMailbox.SelectProbeResidue(rows);

        Assert.Equal(new[] { probeInInbox, probeInDrafts, probeInDeleted }, selected);
        Assert.All(selected, r => Assert.Equal(CorpusPlan.ProbeOrdinal, r.Ordinal));
    }

    /// <summary>
    /// The boundaries, one test each way. Ordinal 1, ordinal <c>count</c> and
    /// <c>int.MaxValue - 1</c> are all ordinary corpus items and none of them may be selected;
    /// <see cref="CorpusPlan.ProbeOrdinal"/> alone is.
    /// <para>
    /// <c>int.MaxValue - 1</c> is in here because it is the value an off-by-one in the
    /// predicate would catch, and because the probe ordinal is deliberately the largest one
    /// there is - a "greater than the plan" test would pass on both and delete a corpus item.
    /// </para>
    /// </summary>
    [Theory]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(ItemCount - 1, false)]
    [InlineData(ItemCount, false)]
    [InlineData(ItemCount + 1, false)]
    [InlineData(int.MaxValue - 1, false)]
    [InlineData(int.MaxValue, true)]
    public void OnlyTheReservedProbeOrdinalIsResidue(int ordinal, bool expected)
    {
        var row = new ComCorpusMailbox.ScanRow(ordinal, "ENTRY-" + ordinal, 6);
        Assert.Equal(expected, ComCorpusMailbox.IsProbeResidue(row));
        Assert.Equal(expected ? 1 : 0, ComCorpusMailbox.SelectProbeResidue(new[] { row }).Count);
    }

    /// <summary>
    /// A scan with no probe rows selects nothing, so the sweep's loop exits on its first pass
    /// and deletes nothing at all. That is the state of every store after the first purge, and
    /// it is what keeps the sweep's cost a single table walk rather than two.
    /// </summary>
    [Fact]
    public void ACleanScanSelectsNothing()
    {
        CorpusPlan plan = Plan();
        var rows = new List<ComCorpusMailbox.ScanRow>();
        for (int ordinal = 1; ordinal <= 500; ordinal++)
        {
            rows.Add(new ComCorpusMailbox.ScanRow(ordinal, "E-" + ordinal, plan.Describe(ordinal).FolderId));
        }

        Assert.Empty(ComCorpusMailbox.SelectProbeResidue(rows));
        Assert.Empty(ComCorpusMailbox.SelectProbeResidue(Array.Empty<ComCorpusMailbox.ScanRow>()));
    }

    /// <summary>
    /// The two keys, joined up: the allowlist the sweep hands
    /// <see cref="CorpusSafety.MayDelete"/> is built FROM the selected rows, so it names probe
    /// EntryIDs and no others - and a corpus item's id, re-read subject and all, is refused by
    /// the same call the sweep makes.
    /// <para>
    /// This is the reason the sweep builds its allowlist from a fresh enumeration rather than
    /// from anything it remembers: the ids it deletes are exactly the ids it just saw carrying
    /// the reserved ordinal.
    /// </para>
    /// </summary>
    [Fact]
    public void TheSweepsAllowlistNamesOnlyTheProbeRowsItJustEnumerated()
    {
        CorpusPlan plan = Plan();
        var rows = new List<ComCorpusMailbox.ScanRow>
        {
            new(1, "C-ONE", plan.Describe(1).FolderId),
            new(CorpusPlan.ProbeOrdinal, "P-ONE", DeletedItemsFolderId),
            new(2, "C-TWO", plan.Describe(2).FolderId),
            new(CorpusPlan.ProbeOrdinal, "P-TWO", DeletedItemsFolderId),
        };

        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(
            ComCorpusMailbox.SelectProbeResidue(rows).Select(r => r.EntryId));

        Assert.Equal(new[] { "P-ONE", "P-TWO" }, allowlist.OrderBy(id => id, StringComparer.Ordinal));

        // Key 1 holds for a probe item and key 2 holds for its subject, so it is deletable.
        Assert.True(CorpusSafety.MayDelete("P-ONE", ProbeSubject("placement InPlaceOnly"), allowlist, CorpusId));

        // A corpus item fails key 1 even though its subject parses perfectly - which is the
        // half of the rule that makes the sweep safe to run against a populated store.
        string corpusSubject = plan.Describe(1).Subject;
        Assert.True(CorpusPlan.TryParseOrdinal(corpusSubject, CorpusId, out _));
        Assert.False(CorpusSafety.MayDelete("C-ONE", corpusSubject, allowlist, CorpusId));

        // And an id on the allowlist whose subject does not parse fails key 2 - the case a
        // recycled EntryID would produce.
        Assert.False(CorpusSafety.MayDelete("P-ONE", "Quarterly numbers", allowlist, CorpusId));
    }

    /// <summary>
    /// The CROSS-STORE sweep (<see cref="ComCorpusMailbox.SweepProbeResidueOutsideTarget"/>), added after
    /// OAI-UNINDEXED 2026-09-24: twelve probe items of four probe sessions sat in Corpus B's Drafts - the
    /// profile's default store, named on no allowlist - because a failed rung's item was filed there on its
    /// first save and deleted, if at all, in the target. The sweep reads every OTHER store's Drafts and
    /// Deleted Items, and selects with the same predicate as the in-store purge - so in a store holding a
    /// whole measurement corpus, only the probe items of THIS corpus id are ever addressed.
    /// </summary>
    [Fact]
    public void TheCrossStoreSweep_ReadsOnlyDraftsAndDeletedItems_AndSelectsOnlyThisCorpussProbeItems()
    {
        Assert.Equal(new[] { DraftsFolderId, DeletedItemsFolderId }, ComCorpusMailbox.CrossStoreResidueFolderIds);

        // Corpus B's Drafts on the guest, as the sweep would scan it for a POPULATION's corpus id: its own
        // corpus items are another corpus id's and never reach a row (the scan keeps only subjects that
        // parse as the id it was given); the population's three stranded probe items do.
        const string population = "hub-unindexed";
        Assert.False(CorpusPlan.TryParseOrdinal(Plan().Describe(1).Subject, population, out _));
        Assert.Equal(CorpusSubjectKind.Current, CorpusPlan.ClassifySubject(
            CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + population + "#"
                + CorpusPlan.ProbeOrdinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture) + "] placement InPlaceOnly",
            population, out int probeOrdinal));
        Assert.Equal(CorpusPlan.ProbeOrdinal, probeOrdinal);

        var drafts = new List<ComCorpusMailbox.ScanRow>
        {
            new(CorpusPlan.ProbeOrdinal, "S-1", DraftsFolderId),
            new(CorpusPlan.ProbeOrdinal, "S-2", DraftsFolderId),
            new(CorpusPlan.ProbeOrdinal, "S-3", DraftsFolderId),

            // A population ITEM that somehow sat there is not the sweep's business: it is not a probe item.
            new(7, "S-ITEM", DraftsFolderId),
        };

        Assert.Equal(new[] { "S-1", "S-2", "S-3" }, ComCorpusMailbox.SelectProbeResidue(drafts).Select(r => r.EntryId));
    }

    /// <summary>Exactly what <c>ComCorpusMailbox.ProbeSubject</c> builds.</summary>
    private static string ProbeSubject(string what)
        => CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + CorpusId + "#"
            + CorpusPlan.ProbeOrdinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture)
            + "] " + what;
}
