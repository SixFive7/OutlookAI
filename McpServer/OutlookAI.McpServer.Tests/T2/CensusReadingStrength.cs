namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Why one folder of a census was compared by COUNT rather than item by item.
/// <para>
/// <b>It exists because the census used to guess, and guessed wrong.</b>
/// <c>StoreCountTripwire.EvaluateByCount</c>'s failure text said
/// <c>(folder above the identity budget)</c> unconditionally, and that is one of at least six
/// possible causes: it is the wrong sentence whenever the cause was the identity CLOCK, a
/// transient COM failure, an unusable table, a repeat pass that found no baseline reading to
/// match, or a plan that was never going to walk the folder at all. A failure message that names
/// the wrong cause sends the next reader to the wrong remedy, and this guard's messages are read
/// exactly once, in an emergency, by somebody who thinks mail has just been deleted.
/// </para>
/// </summary>
public enum CensusCountReason
{
    /// <summary>
    /// Not counted at all - the folder was walked item by item. Never valid on a
    /// <see cref="FolderCensus.CountOnly"/>, which refuses it.
    /// </summary>
    Walked = 0,

    /// <summary>The plan holds no identity budget at all - the hub's count-only census.</summary>
    PlanIsCountOnly,

    /// <summary>
    /// A self-pruning folder (Deleted Items, Junk, the sync-issue subtree). Identity would buy
    /// nothing there: a departure is never a failure, so the walk is pure cost.
    /// </summary>
    SelfPruningFolder,

    /// <summary>The folder held more items than one folder may spend on identity.</summary>
    AbovePerFolderLimit,

    /// <summary>The per-STORE item budget was already spent by earlier folders in the walk.</summary>
    StoreItemBudgetSpent,

    /// <summary>
    /// A repeat pass, and the baseline did not identify this folder - so there is nothing to
    /// compare item by item even if this pass could afford the walk.
    /// </summary>
    NotIdentifiedAtBaseline,

    /// <summary>The identity TIME budget for this store had expired before the folder was reached.</summary>
    IdentityClockExpired,

    /// <summary>
    /// The walk was attempted and abandoned: a COM call failed, a column was missing, a row
    /// carried no EntryID, an EntryID repeated, or the count moved under the read.
    /// </summary>
    TableUnusable,
}

/// <summary>
/// How strong one folder's two readings were, and what a failure over them may honestly say.
/// <para>
/// <b>The defect this closes, found by reading during the 2026-08-24 tripwire investigation.</b>
/// Whether a folder is compared BY IDENTITY or BY COUNT is decided independently on each pass, and
/// the post-run decision is timing-dependent: the identity clock, a folder that grew, or one
/// transient COM failure silently switches a folder from <c>EvaluateByIdentity</c> to
/// <c>EvaluateByCount</c>. The count rule cannot exonerate a filing and cannot see a departure
/// masked by an arrival, so the SAME mailbox state yields <c>note: filed (not loss)</c> on one
/// reading and <c>ITEMS LOST</c> on the next, purely because the second reading was weaker. That is
/// a census reporting differently twice with nothing having changed - the very thing the old
/// "treating as enumeration noise" message blamed on COM - and it is exactly the shape a re-census
/// 30 s later would clear, which is how it has been hiding behind the retry ladder.
/// </para>
/// <para>
/// <b>What is fixed here and what is not.</b> This is option (2) of the three the TODO records:
/// carry the reason, and say which reading was the weak one. It does not stop the false failure -
/// that is option (1), re-walking the degraded folder, which needs a COM call and therefore a live
/// run to land - and it does not refuse the comparison, which is option (3) and noisier by a lot.
/// What it removes is the wrong sentence and the silence: a degraded pair now SAYS it is degraded,
/// so a reader can tell "this folder lost items" from "this reading was weaker than the one it is
/// being compared against", which are different claims and used to print identically.
/// </para>
/// <para>
/// Pure - two censuses in, a sentence out. The caller is behind a COM census no CI test can
/// execute, which is why the decision lives here.
/// </para>
/// </summary>
public static class CensusReadingStrength
{
    /// <summary>
    /// True when the post-run reading of a folder is WEAKER than the baseline's: the baseline
    /// identified it and this pass could only count it. That is the one direction that changes a
    /// verdict, because the comparison then falls back to the count rule for a folder the census
    /// had the evidence to judge properly at the other end.
    /// </summary>
    public static bool Degraded(FolderCensus baseline, FolderCensus postRun)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(postRun);
        return baseline.HasIdentities && !postRun.HasIdentities;
    }

    /// <summary>
    /// The clause a count-rule verdict carries, naming which of the two readings was the weak one
    /// and why. Always begins with a space so a caller can append it to a sentence.
    /// </summary>
    public static string Explain(FolderCensus baseline, FolderCensus postRun)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(postRun);

        if (Degraded(baseline, postRun))
        {
            // The timing-dependent case, and the only one where the comparison itself is suspect.
            // Said in full because the reader's next question is always "so did it happen or not".
            return " The POST-RUN reading of this folder was WEAKER than the baseline's, which"
                + " identified it item by item (" + Describe(postRun.CountReason) + "), so this"
                + " verdict comes from the count rule for a folder the census could have judged"
                + " properly. The same mailbox state can read as a loss on one pass and as filed"
                + " on another for this reason alone - re-read this folder before concluding"
                + " anything.";
        }

        if (!baseline.HasIdentities && postRun.HasIdentities)
        {
            // Harmless, and worth saying: the pair is still count-only, but the weak end is the
            // one nothing can be done about now.
            return " The BASELINE reading was the weaker of the two ("
                + Describe(baseline.CountReason) + "), so the comparison is by count whatever this"
                + " pass managed.";
        }

        if (baseline.CountReason == postRun.CountReason)
        {
            return " Both readings were counts (" + Describe(baseline.CountReason) + ").";
        }

        return " Both readings were counts (baseline: " + Describe(baseline.CountReason)
            + "; after the run: " + Describe(postRun.CountReason) + ").";
    }

    /// <summary>One reason, in the words a reader can act on.</summary>
    public static string Describe(CensusCountReason reason)
    {
        return reason switch
        {
            CensusCountReason.Walked => "this folder was walked item by item",
            CensusCountReason.PlanIsCountOnly => "this store is censused by count only",
            CensusCountReason.SelfPruningFolder => "a self-pruning folder is never walked",
            CensusCountReason.AbovePerFolderLimit => "folder above the per-folder identity limit",
            CensusCountReason.StoreItemBudgetSpent => "the store's identity item budget was already spent",
            CensusCountReason.NotIdentifiedAtBaseline => "the baseline did not identify this folder",
            CensusCountReason.IdentityClockExpired => "the identity TIME budget expired",
            CensusCountReason.TableUnusable => "the folder's table could not be read for this pass",
            _ => "reason not recorded",
        };
    }
}
