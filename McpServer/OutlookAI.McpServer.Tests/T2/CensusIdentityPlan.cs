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

    private readonly HashSet<string>? _repeatFolders;
    private readonly int _perFolderLimit;
    private int _remaining;

    private CensusIdentityPlan(HashSet<string>? repeatFolders, int perFolderLimit, int budget)
    {
        _repeatFolders = repeatFolders;
        _perFolderLimit = perFolderLimit;
        _remaining = budget;
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
    /// Counts only. Used for the designated test mailbox, which this guard exempts anyway
    /// (its churn is tagged and the zero-artifact sweep polices it), so walking it would buy
    /// nothing and it is the busiest store in the run.
    /// </summary>
    public static CensusIdentityPlan CountOnly()
    {
        return new CensusIdentityPlan(null, 0, 0);
    }

    /// <summary>A first census: walk what fits, in folder-tree order, until the budget runs out.</summary>
    public static CensusIdentityPlan Baseline(
        int perFolderLimit = DefaultPerFolderLimit, int perStoreItemBudget = DefaultPerStoreItemBudget)
    {
        return new CensusIdentityPlan(null, perFolderLimit, perStoreItemBudget);
    }

    /// <summary>
    /// The matching second census: walk exactly the folders <paramref name="baseline"/>
    /// identified, and nothing else. Comparability outranks cost here.
    /// </summary>
    public static CensusIdentityPlan Repeating(
        IReadOnlyDictionary<string, FolderCensus> baseline, int perFolderLimit = DefaultPerFolderLimit)
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

        return new CensusIdentityPlan(folders, perFolderLimit, 0);
    }

    /// <summary>Whether this folder should be walked as well as counted.</summary>
    /// <param name="folderKey">Census key, volatile prefix included.</param>
    /// <param name="isVolatile">True for folders the system prunes on its own.</param>
    /// <param name="itemCount">What the folder holds right now.</param>
    public bool ShouldIdentify(string folderKey, bool isVolatile, int itemCount)
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
        return FoldersDegradedToCount == 0
            ? line
            : line + ", " + FoldersDegradedToCount.ToString(CultureInfo.InvariantCulture)
                + " folder(s) fell back to counting";
    }
}
