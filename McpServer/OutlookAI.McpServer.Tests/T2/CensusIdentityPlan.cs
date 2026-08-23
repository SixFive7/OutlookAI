using System.Diagnostics;
using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// How much of a store the tripwire census may walk item by item.
/// <para>
/// Identity is what lets the guard say WHICH items left rather than only how many, and it
/// is the only way it can tell a filing from a deletion. It is optional because it is not
/// free: a real profile holds a 108,000-item Archive, a 10,000-item Sent Items and a 20,000
/// item Deleted Items, and walking all of that twice per run would cost more than the run.
/// So the census counts everything and walks what it can afford.
/// </para>
/// <para>
/// Since 2026-08-20 the walk is a BULK TABLE READ rather than an item-by-item one, so what
/// this budget bounds has changed: it used to bound round trips (five per item, which is
/// what exceeded the STA timeout on an Exchange store), and now it bounds bytes on the wire
/// and rows held in memory. The numbers were left where they were on purpose - moving them
/// changes what the guard proves, and that is the maintainer's call, not a side effect of
/// making the census affordable.
/// </para>
/// <para>
/// <b>Two budgets, because the item counts were always a proxy for the thing that actually
/// matters.</b> 500/3,000 bound the SIZE of the reading and therefore what the guard proves;
/// <see cref="DefaultIdentityTimeBudgetMs"/> bounds how long taking it may go on, which is
/// what 2026-08-20 actually ran out of. Before it existed the identity walk had no clock of
/// its own at all and simply shared the 3-minute STA join with the folder-tree walk, so the
/// only thing that could stop a slow walk was the join killing the whole store's census and
/// refusing the tier. Now the walk stops itself, says so, and the store is still counted.
/// </para>
/// <para>
/// Two shapes, and the difference matters. A BASELINE plan spends a budget and records what
/// it bought; the matching POST-RUN plan spends nothing and simply repeats the baseline's
/// choices, because a folder identified at one end and merely counted at the other cannot
/// be compared item by item at all. That is also why the repeat pass ignores the budget: a
/// folder that grew during the run must still be walked, up to
/// <see cref="RepeatGrowthHeadroom"/> times the per-folder limit, after which it degrades
/// to a count rather than stalling the teardown.
/// </para>
/// <para>
/// Used from one census at a time (each store's walk is a blocking STA call), so the
/// remaining budget is plain mutable state and needs no locking. It doubles as the census's
/// PROGRESS record, which is the one thing the caller can still read when that STA call
/// times out: the counters are plain <c>int</c> writes on the census thread and plain reads
/// on the caller's, so a reading taken after a timeout may be a moment stale but can never
/// be torn. Diagnostics only - nothing decides anything from these.
/// </para>
/// </summary>
public sealed class CensusIdentityPlan
{
    /// <summary>
    /// Largest folder the baseline will walk. Chosen so the folders a person actually works
    /// in are covered - the 2026-08-18 false alarm was an Inbox of 168, a Postvak IN of 52
    /// and a junk folder of 1 - while archives and Sent Items stay counted. Raising it buys
    /// precision in bigger folders and costs census time on every run.
    /// </summary>
    public const int DefaultPerFolderLimit = 500;

    /// <summary>
    /// Items one store may spend on identity per census. Bounds the whole profile at
    /// stores x this, which is the number that decides what the guard costs.
    /// </summary>
    public const int DefaultPerStoreItemBudget = 3_000;

    /// <summary>
    /// How much a folder may grow between the two censuses and still be walked the second
    /// time. Four times is far past ordinary arrival rates; past it the folder is counted,
    /// which is a weaker reading but never a wrong one.
    /// </summary>
    public const int RepeatGrowthHeadroom = 4;

    /// <summary>
    /// How long ONE STORE's census may go on identifying items before it stops walking and
    /// counts the rest. A CEILING, not a target: it is set at roughly seven times the only
    /// measurement that exists, and on that profile it never fires.
    /// <para>
    /// <b>What it governs, and what it does not.</b> It bounds the WALK only. Every folder is
    /// still counted, so the count rule still guards the whole store when this expires - the
    /// reading degrades exactly the way it already degrades when the per-store item budget
    /// runs out or a table comes back unusable, and the plan says so in
    /// <see cref="Describe"/>. It is not a deadline on the census as a whole and it must
    /// never be used as one: an unmeasured mailbox cannot be proven untouched, so a census
    /// that cannot COUNT still refuses the tier.
    /// </para>
    /// <para>
    /// <b>Why 120 s, and why the number is weak evidence.</b> The only trial ever taken is one
    /// run of the table-read census on the maintainer's real profile: 5 stores, 159 folders,
    /// 2,044 items, 16.9 s for the whole pass. One run is not a distribution, and the risk is
    /// asymmetric - a budget set too low kills an operation that was working, while one set
    /// too high costs nothing at all when the work finishes early. So this is deliberately
    /// generous and is to be NARROWED later from VM measurements rather than defended now.
    /// </para>
    /// <para>
    /// <b>It is one rung of a ladder.</b> The rung above it is the STA join the census runs
    /// under (<c>LiveOutlookTestMailer.CensusStaBudget</c>), which must stay strictly larger:
    /// this project has already shipped the failure where the outer timer killed an operation
    /// that was working fine inside its own budget, and a budget that can never expire because
    /// something above it fires first is not a budget. T1 pins the ordering.
    /// </para>
    /// </summary>
    public const int DefaultIdentityTimeBudgetMs = 120_000;

    private readonly HashSet<string>? _repeatFolders;
    private readonly int _perFolderLimit;
    private readonly TimeSpan _identityTimeBudget;
    private readonly Func<TimeSpan> _elapsed;
    private int _remaining;

    private CensusIdentityPlan(
        HashSet<string>? repeatFolders,
        int perFolderLimit,
        int budget,
        int identityTimeBudgetMs,
        Func<TimeSpan>? elapsed)
    {
        _repeatFolders = repeatFolders;
        _perFolderLimit = perFolderLimit;
        _remaining = budget;
        _identityTimeBudget = TimeSpan.FromMilliseconds(identityTimeBudgetMs);

        // Stopwatch, not the wall clock: this measures how long something has been going, and
        // a wall clock jumps (NTP, a VM resuming, a person setting the time). The same rule
        // LiveWaitBudget states for every live-tier wait.
        Stopwatch clock = Stopwatch.StartNew();
        _elapsed = elapsed ?? (() => clock.Elapsed);
    }

    /// <summary>Folders walked so far under this plan.</summary>
    public int IdentifiedFolders { get; private set; }

    /// <summary>Items walked so far under this plan - the number that costs census time.</summary>
    public int IdentifiedItems { get; private set; }

    /// <summary>
    /// Mail folders this census has reached at all, walked or merely counted. Exists so a
    /// census that RUNS OUT OF TIME can say where it was: on 2026-08-20 the live tier
    /// refused to start because one store's census exceeded the STA budget, and the refusal
    /// could not distinguish a slow folder tree from a slow item walk.
    /// </summary>
    public int MeasuredFolders { get; private set; }

    /// <summary>
    /// Folders this plan chose to walk and could not, so they were recorded as counts. A
    /// non-zero value is not a failure, but it IS the number that says how much of the
    /// identity reading a run actually got: a table missing its columns on every folder
    /// would disable the identity half of the guard, and it must not do that silently.
    /// </summary>
    public int FoldersDegradedToCount { get; private set; }

    /// <summary>
    /// Folders this plan wanted to walk and refused to, because
    /// <see cref="DefaultIdentityTimeBudgetMs"/> had run out. Separate from
    /// <see cref="FoldersDegradedToCount"/> on purpose: that one says the folder could not be
    /// read, this one says the census could not afford to. They point at different remedies.
    /// </summary>
    public int FoldersDeniedByClock { get; private set; }

    /// <summary>True when the identity time budget expired during this census.</summary>
    public bool IdentityClockExpired => FoldersDeniedByClock > 0;

    /// <summary>
    /// Counts only. Used for the designated test mailbox, which this guard exempts anyway
    /// (its churn is tagged and the zero-artifact sweep polices it), so walking it would buy
    /// nothing and it is the busiest store in the run.
    /// </summary>
    public static CensusIdentityPlan CountOnly()
    {
        return new CensusIdentityPlan(null, 0, 0, DefaultIdentityTimeBudgetMs, null);
    }

    /// <summary>A first census: walk what fits, in folder-tree order, until the budget runs out.</summary>
    public static CensusIdentityPlan Baseline(
        int perFolderLimit = DefaultPerFolderLimit,
        int perStoreItemBudget = DefaultPerStoreItemBudget,
        int identityTimeBudgetMs = DefaultIdentityTimeBudgetMs)
    {
        return new CensusIdentityPlan(null, perFolderLimit, perStoreItemBudget, identityTimeBudgetMs, null);
    }

    /// <summary>
    /// The matching second census: walk exactly the folders <paramref name="baseline"/>
    /// identified, and nothing else. Comparability outranks cost here.
    /// </summary>
    public static CensusIdentityPlan Repeating(
        IReadOnlyDictionary<string, FolderCensus> baseline,
        int perFolderLimit = DefaultPerFolderLimit,
        int identityTimeBudgetMs = DefaultIdentityTimeBudgetMs)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        HashSet<string> folders = new(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, FolderCensus> entry in baseline)
        {
            if (entry.Value.HasIdentities)
            {
                folders.Add(entry.Key);
            }
        }

        return new CensusIdentityPlan(folders, perFolderLimit, 0, identityTimeBudgetMs, null);
    }

    /// <summary>
    /// A plan whose clock is supplied rather than started, so CI can pin what the time budget
    /// does without spending the budget. Same assembly only - this is a test seam, not an API.
    /// </summary>
    internal static CensusIdentityPlan WithClock(
        Func<TimeSpan> elapsed,
        int perFolderLimit = DefaultPerFolderLimit,
        int perStoreItemBudget = DefaultPerStoreItemBudget,
        int identityTimeBudgetMs = DefaultIdentityTimeBudgetMs)
    {
        ArgumentNullException.ThrowIfNull(elapsed);
        return new CensusIdentityPlan(null, perFolderLimit, perStoreItemBudget, identityTimeBudgetMs, elapsed);
    }

    /// <summary>
    /// A REPEAT plan on a supplied clock, for the same reason. The repeat pass is the one that
    /// runs at the end of a 27-minute tier run, so it is also the one most likely to meet a
    /// profile that has gone slow.
    /// </summary>
    internal static CensusIdentityPlan RepeatingWithClock(
        IReadOnlyDictionary<string, FolderCensus> baseline,
        Func<TimeSpan> elapsed,
        int perFolderLimit = DefaultPerFolderLimit,
        int identityTimeBudgetMs = DefaultIdentityTimeBudgetMs)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(elapsed);
        HashSet<string> folders = new(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, FolderCensus> entry in baseline)
        {
            if (entry.Value.HasIdentities)
            {
                folders.Add(entry.Key);
            }
        }

        return new CensusIdentityPlan(folders, perFolderLimit, 0, identityTimeBudgetMs, elapsed);
    }

    /// <summary>Whether this folder should be walked as well as counted.</summary>
    /// <param name="folderKey">Census key, volatile prefix included.</param>
    /// <param name="isVolatile">True for folders the system prunes on its own.</param>
    /// <param name="itemCount">What the folder holds right now.</param>
    public bool ShouldIdentify(string folderKey, bool isVolatile, int itemCount)
    {
        if (!WantsToIdentify(folderKey, isVolatile, itemCount))
        {
            return false;
        }

        // The clock is asked LAST, so the counter below means what it says: folders this plan
        // wanted to walk and could not because the census had gone on too long. Asking first
        // would count every folder the plan was never going to walk anyway and turn a
        // diagnostic into noise.
        if (_elapsed() >= _identityTimeBudget)
        {
            FoldersDeniedByClock++;
            return false;
        }

        return true;
    }

    /// <summary>The size and comparability rules, with no clock in them.</summary>
    private bool WantsToIdentify(string folderKey, bool isVolatile, int itemCount)
    {
        if (_repeatFolders != null)
        {
            return _repeatFolders.Contains(folderKey) && itemCount <= _perFolderLimit * RepeatGrowthHeadroom;
        }

        // A self-pruning folder can shrink without anyone doing anything, so a departure
        // there is never a failure and identity would only cost time. Deleted Items is also
        // the largest folder in most stores, which is the other half of the reason.
        //
        // The limit is checked against zero as well: without that, a count-only plan (limit
        // and budget both zero) would still claim EMPTY folders, which walks nothing but
        // marks them as compared item by item - a reading the plan was told not to take.
        return _perFolderLimit > 0
            && !isVolatile
            && itemCount <= _perFolderLimit
            && itemCount <= _remaining;
    }

    /// <summary>Records a completed walk. Never called for a walk that had to be abandoned.</summary>
    public void Spend(int items)
    {
        IdentifiedFolders++;
        IdentifiedItems += items;
        _remaining -= items;
        if (_remaining < 0)
        {
            _remaining = 0;
        }
    }

    /// <summary>Records that one more mail folder was reached, before deciding what to do with it.</summary>
    public void NoteFolderMeasured()
    {
        MeasuredFolders++;
    }

    /// <summary>Records a folder this plan wanted to walk and had to record as a count instead.</summary>
    public void NoteDegradedToCount()
    {
        FoldersDegradedToCount++;
    }

    /// <summary>One line for the console, so the first live run reports what this actually cost.</summary>
    public string Describe()
    {
        string line = MeasuredFolders.ToString(CultureInfo.InvariantCulture) + " folder(s) measured, "
            + IdentifiedFolders.ToString(CultureInfo.InvariantCulture) + " folder(s), "
            + IdentifiedItems.ToString(CultureInfo.InvariantCulture) + " item(s) identified";
        if (FoldersDegradedToCount > 0)
        {
            line += ", " + FoldersDegradedToCount.ToString(CultureInfo.InvariantCulture)
                + " folder(s) fell back to counting";
        }

        // Loud, and only when it happened: a census that quietly stopped identifying would
        // leave the guard weaker than the log says it is.
        if (FoldersDeniedByClock > 0)
        {
            line += ", IDENTITY TIME BUDGET EXPIRED ("
                + _identityTimeBudget.TotalSeconds.ToString("0", CultureInfo.InvariantCulture) + " s) - "
                + FoldersDeniedByClock.ToString(CultureInfo.InvariantCulture)
                + " folder(s) counted instead of walked";
        }

        return line;
    }
}
