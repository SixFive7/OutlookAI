using System.Reflection;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T1;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// A COM session that answers what nearly every suite around search needs - the profile's name, its stores,
/// a folder's description and the freshness sweep - and whatever else a test scripts with
/// <see cref="Answering"/>, and refuses everything else, naming the member it was asked for. Six suites
/// declared it for themselves, three of them word for word (the code-sharing refactor's S5a, the survey's
/// T03). A <see cref="DispatchProxy"/> rather than a stub per member, so a member added to the contract needs
/// no change here. Not sealed: DispatchProxy derives from its TProxy at runtime and refuses a sealed one.
/// </summary>
internal class StandInSession : DispatchProxy
{
    private readonly Dictionary<string, Func<object?[]?, object?>> _answers = new(StringComparer.Ordinal);
    private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
    private Func<string?, ComSweepResult> _sweep = _ => throw new NotSupportedException("This test scripts no freshness sweep.");
    private int _storeListReads;

    /// <summary>How many times this session was asked for the profile's stores - per session, so suites running at once do not share it.</summary>
    internal int StoreListReads => Volatile.Read(ref _storeListReads);

    /// <summary>
    /// A session over <paramref name="stores"/> whose sweep answers with <paramref name="sweep"/>, given the one
    /// store it is scoped to or null; <paramref name="script"/> adds the members the test needs beyond those.
    /// </summary>
    internal static IOutlookSession Create(
        IReadOnlyList<ComStoreDetail> stores,
        Func<string?, ComSweepResult>? sweep = null,
        Action<StandInSession>? script = null) => Create(stores, sweep, script, out _);

    /// <summary>The same, handing back the stand-in itself, for a test that reads <see cref="StoreListReads"/>.</summary>
    internal static IOutlookSession Create(
        IReadOnlyList<ComStoreDetail> stores,
        Func<string?, ComSweepResult>? sweep,
        Action<StandInSession>? script,
        out StandInSession standIn)
    {
        object proxy = Create<IOutlookSession, StandInSession>()
            ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
        standIn = (StandInSession)proxy;
        standIn._stores = stores;
        if (sweep != null)
        {
            standIn._sweep = sweep;
        }

        script?.Invoke(standIn);
        return (IOutlookSession)proxy;
    }

    /// <summary>Answers <paramref name="member"/> of <see cref="IOutlookSession"/> with <paramref name="answer"/>, given its arguments.</summary>
    internal StandInSession Answering(string member, Func<object?[]?, object?> answer)
    {
        _answers[member] = answer;
        return this;
    }

    /// <inheritdoc />
    protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
    {
        string? member = targetMethod?.Name;
        if (member != null && _answers.TryGetValue(member, out Func<object?[]?, object?>? answer))
        {
            return answer(args);
        }

        switch (member)
        {
            case nameof(IOutlookSession.GetProfileName):
                return "T1 stand-in profile";

            case nameof(IOutlookSession.GetStoreDetails):
                Interlocked.Increment(ref _storeListReads);
                return TestFolders.WithRoots(_stores);

            case nameof(IOutlookSession.TryDescribeFolder):
                return TestFolders.Describe(args);

            // Argument 3 is onlyStoreDisplayName; see IOutlookSession.SweepFoldersNewerThan.
            case nameof(IOutlookSession.SweepFoldersNewerThan):
                return _sweep(args?[3] as string);

            default:
                throw new NotSupportedException(
                    "The stand-in session was asked for " + (member ?? "an unnamed member")
                    + ", which this test does not model.");
        }
    }
}
