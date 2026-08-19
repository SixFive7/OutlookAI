namespace OutlookAI.McpServer.Tests.T2;

/// <summary>Where one guarded collection sits in the run that is actually happening.</summary>
public enum GuardedCollectionPosition
{
    /// <summary>
    /// No run plan was published, or this collection is not in it. The caller must assume
    /// it could be the last one - verifying too often costs time, verifying never costs the
    /// guarantee.
    /// </summary>
    Unknown = 0,

    /// <summary>A later guarded collection will still run; the tripwire keeps its baseline.</summary>
    NotLast = 1,

    /// <summary>Nothing guarded runs after this one; this is where the run must be verified.</summary>
    Last = 2,
}

/// <summary>
/// What the current test run actually contains, published by
/// <see cref="SuiteCollectionOrderer"/> before any fixture is constructed.
/// <para>
/// <b>The defect this exists to fix.</b> The store-count tripwire took its baseline in every
/// live collection fixture but compared it in exactly one - <c>LiveLifecycleFixture</c>'s
/// dispose, on the strength of that collection being forced last. That holds for a whole-tier
/// run and fails for every FILTERED one. <c>--filter "FullyQualifiedName~LiveTableSortProbeTests"</c>
/// selects no LiveLifecycle test, so the fixture is never built, <c>Verify</c> is never called,
/// and the run pays for a census that is then thrown away: the guard that stands between the
/// suite and the incident that once destroyed real mail silently does nothing, and the run
/// reports green. Filtered runs are not an edge case here - they are how the tier is meant to
/// be used on a test machine, and they are the exact commands the session log hands the
/// maintainer for the two probes.
/// </para>
/// <para>
/// The collection orderer is the only vantage point that can fix it. xunit hands it the
/// collections that will run AFTER filtering, before anything is constructed, so it knows the
/// shape of this run when nothing else does. Each guarded fixture then asks, as it disposes,
/// whether anything guarded comes after it.
/// </para>
/// <para>
/// Process-wide mutable state, like the tripwire it serves, because a test run is one process
/// and the plan is one fact about it.
/// </para>
/// </summary>
public static class LiveTierRunPlan
{
    private static readonly object Gate = new();
    private static IReadOnlyList<string>? _ordered;

    /// <summary>
    /// Records the collections this run will execute, in the order they will execute.
    /// Idempotent: xunit orders once, but publishing the same list twice must not matter.
    /// </summary>
    public static void Publish(IEnumerable<string> orderedCollectionNames)
    {
        ArgumentNullException.ThrowIfNull(orderedCollectionNames);
        List<string> ordered = orderedCollectionNames.Where(name => name != null).ToList();
        lock (Gate)
        {
            _ordered = ordered;
        }
    }

    /// <summary>The published order, or null when nothing has been published.</summary>
    public static IReadOnlyList<string>? Current
    {
        get
        {
            lock (Gate)
            {
                return _ordered;
            }
        }
    }

    /// <summary>Where <paramref name="collectionName"/> sits in the run being executed.</summary>
    public static GuardedCollectionPosition PositionOf(string collectionName)
    {
        return PositionIn(Current, collectionName);
    }

    /// <summary>
    /// The rule, with the run plan injected so CI pins it without xunit having to run a
    /// live tier: the last GUARDED collection in the order is where verification belongs,
    /// and an order that does not mention the collection at all answers
    /// <see cref="GuardedCollectionPosition.Unknown"/> rather than guessing.
    /// </summary>
    public static GuardedCollectionPosition PositionIn(IReadOnlyList<string>? ordered, string collectionName)
    {
        ArgumentNullException.ThrowIfNull(collectionName);
        if (ordered == null || ordered.Count == 0)
        {
            return GuardedCollectionPosition.Unknown;
        }

        int index = -1;
        for (int i = 0; i < ordered.Count; i++)
        {
            if (string.Equals(ordered[i], collectionName, StringComparison.Ordinal))
            {
                index = i;
                break;
            }
        }

        if (index < 0)
        {
            return GuardedCollectionPosition.Unknown;
        }

        for (int i = index + 1; i < ordered.Count; i++)
        {
            if (LiveCollections.IsGuarded(ordered[i]))
            {
                return GuardedCollectionPosition.NotLast;
            }
        }

        return GuardedCollectionPosition.Last;
    }

    /// <summary>Test hook: forgets the plan so a self-test can drive every branch.</summary>
    internal static void ResetForTests()
    {
        lock (Gate)
        {
            _ordered = null;
        }
    }
}
