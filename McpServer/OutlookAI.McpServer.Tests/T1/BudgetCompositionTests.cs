using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;

using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The budgets NEST, and the prose that quotes them tells the truth.
/// <para>
/// Every assertion here exists because the relationship it checks was previously carried by
/// two literals that happened to agree, or by a comment. Both fail silently: the exhaustive
/// scan's soft budget was written as its own <c>120_000</c>, equalled the COM host's hard
/// deadline exactly, and so could never produce the partial-results answer its own tool
/// description promises - the caller got a timeout and a killed host instead. Nothing
/// failed. Nothing could have.
/// </para>
/// <para>
/// The tool-description assertions are the other half. Those strings are attribute literals
/// (a <c>const int</c> cannot be concatenated into a <c>const string</c> in C#, so they
/// cannot be interpolated from the constants), which means the only way a number in prose
/// can be kept honest is a test that derives the expected text from the constant. Change a
/// budget without changing the description and this fails, naming both.
/// </para>
/// </summary>
public sealed class BudgetCompositionTests
{
    /// <summary>
    /// The inner scan budget is strictly INSIDE the outer hard deadline, by the declared
    /// headroom. Equality is the defect: the scan stops only once elapsed has PASSED its
    /// budget, then still has to serialize its results back over the pipe, while the
    /// watchdog fires at &gt;=.
    /// </summary>
    [Fact]
    public void ExhaustiveScanBudget_LeavesHeadroomInsideTheOperationDeadline()
    {
        Assert.True(
            MailService.ExhaustiveTimeBudgetMs < ComOperationBudgets.OperationDeadlineMs,
            $"the exhaustive scan's soft budget ({MailService.ExhaustiveTimeBudgetMs} ms) must be strictly inside the COM "
            + $"host's hard operation deadline ({ComOperationBudgets.OperationDeadlineMs} ms); equal means the documented "
            + "partial-results outcome is unreachable and a long scan becomes a timeout plus a host kill");

        Assert.Equal(
            ComOperationBudgets.OperationDeadlineMs - ComOperationBudgets.ResultReturnHeadroomMs,
            MailService.ExhaustiveTimeBudgetMs);

        Assert.True(
            ComOperationBudgets.ResultReturnHeadroomMs > 0,
            "the return-trip headroom must be positive - it is the whole mechanism");
    }

    /// <summary>
    /// The exhaustive advice text and the tool description quote the SAME number the
    /// constant holds. Both are rendered as whole seconds.
    /// </summary>
    [Fact]
    public void ExhaustiveBudget_IsQuotedTruthfullyInTheToolSurface()
    {
        string expected = (MailService.ExhaustiveTimeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture) + " s";
        string exhaustiveParameter = ParameterDescription(nameof(OutlookTools.Search), "exhaustive");

        Assert.Contains(
            expected,
            exhaustiveParameter,
            System.StringComparison.Ordinal);
    }

    /// <summary>
    /// The subject cap is ONE number, and the tool surface quotes THAT number.
    /// <para>
    /// It was three <c>&gt; 255</c> literals in MailService plus the number a fourth time as
    /// prose in the derived-subject hint, related by nothing - the same shape that produced a
    /// "these are pinned" comment nothing pinned. A number in an attribute literal cannot be
    /// interpolated from a <c>const int</c> in C#, so the only way to keep it honest is from
    /// the outside: the phrase is built here from the constant, so changing
    /// <see cref="MailService.SubjectCharsCap"/> without changing the hint fails and says so.
    /// </para>
    /// </summary>
    [Fact]
    public void SubjectCap_IsQuotedTruthfullyInTheToolSurface()
    {
        string expected = "Max " + MailService.SubjectCharsCap.ToString(CultureInfo.InvariantCulture) + " characters";

        // Every tool that advertises the cap, so a hint added later cannot quote a stale one.
        foreach (string tool in new[]
                 {
                     nameof(OutlookTools.ReplyDraft),
                     nameof(OutlookTools.ReplyAllDraft),
                     nameof(OutlookTools.ForwardDraft),
                 })
        {
            string hint = ParameterDescription(tool, "subject");
            Assert.True(
                hint.Contains(expected, System.StringComparison.Ordinal),
                $"{tool}'s 'subject' description must quote MailService.SubjectCharsCap "
                + $"({MailService.SubjectCharsCap}) as \"{expected}\", because that is the length the service "
                + $"actually enforces. The description reads: \"{hint}\"");
        }
    }

    /// <summary>
    /// outlook_health's description quotes its COM probe budget, and that budget is the
    /// supervisor's own - not a third independent 5 000.
    /// </summary>
    [Fact]
    public void HealthProbeBudget_IsOneValueAndIsQuotedTruthfully()
    {
        Assert.Equal(ComOperationBudgets.HealthProbeDeadlineMs, MailService.HealthProbeBudgetMs);
        Assert.Equal(ComOperationBudgets.HealthProbeDeadlineMs, (int)ComHostPolicy.HealthProbeDeadlineMilliseconds);

        string expected = "gives up after "
            + (MailService.HealthProbeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture) + " s";
        Assert.Contains(expected, ToolDescription(nameof(OutlookTools.OutlookHealth)), System.StringComparison.Ordinal);
    }

    /// <summary>
    /// One indexed search is one index statement plus one freshness sweep, and the pair fits
    /// inside the budget the sweep itself runs under. Search's own description calls it
    /// "sub-second and cheap"; on the index client's 30 s default with no tool-level bound
    /// above it, the composed worst case bore no relation to that.
    /// </summary>
    [Fact]
    public void SearchBudget_IsComposedFromItsPartsAndFitsTheOperationDeadline()
    {
        Assert.Equal((MailService.SearchIndexTimeoutSeconds * 1000) + MailService.SweepBudgetMs, MailService.SearchBudgetMs);
        Assert.True(
            MailService.SearchBudgetMs <= ComOperationBudgets.OperationDeadlineMs,
            $"one search ({MailService.SearchBudgetMs} ms of index + sweep) must fit inside the COM host operation "
            + $"deadline ({ComOperationBudgets.OperationDeadlineMs} ms)");
        Assert.True(
            MailService.SearchIndexTimeoutSeconds <= OleDbIndexClient.DefaultCommandTimeoutSeconds,
            "the search path must not ask for MORE index time than the client's own default");
    }

    /// <summary>
    /// A move/archive batch is bounded as a whole. Per-item deadlines bound one round trip;
    /// 50 ids at 2-3 round trips each had no aggregate bound at all.
    /// </summary>
    [Fact]
    public void MoveBatchBudget_BoundsTheWholeBatchNotJustOneItem()
    {
        Assert.True(MailService.MoveBatchBudgetMs > 0);
        Assert.Equal(ComOperationBudgets.OperationDeadlineMs, MailService.MoveBatchBudgetMs);
        Assert.True(
            MailService.MoveBatchBudgetMs < (long)MailService.MoveIdsCap * ComOperationBudgets.OperationDeadlineMs,
            "the batch budget must be smaller than the worst case it replaces");
    }

    /// <summary>
    /// The handshake follows the operation deadline, is never doubled, and never drops below
    /// the floor that keeps a shortened test budget from failing on child startup instead of
    /// on the path under test.
    /// </summary>
    [Theory]
    [InlineData(0L, ComHostPolicy.HandshakeFloorMilliseconds)]
    [InlineData(1L, ComHostPolicy.HandshakeFloorMilliseconds)]
    [InlineData(4_000L, ComHostPolicy.HandshakeFloorMilliseconds)]
    [InlineData(ComHostPolicy.HandshakeFloorMilliseconds, ComHostPolicy.HandshakeFloorMilliseconds)]
    [InlineData(ComHostPolicy.HandshakeFloorMilliseconds + 1, ComHostPolicy.HandshakeFloorMilliseconds + 1)]
    [InlineData(20_000L, 20_000L)]
    [InlineData(ComHostPolicy.HandshakeBudgetMilliseconds, ComHostPolicy.HandshakeBudgetMilliseconds)]
    [InlineData(ComHostPolicy.HandshakeBudgetMilliseconds + 1, ComHostPolicy.HandshakeBudgetMilliseconds)]
    [InlineData(120_000L, ComHostPolicy.HandshakeBudgetMilliseconds)]
    public void HandshakeBudget_FollowsTheDeadlineBetweenItsFloorAndItsCeiling(long deadline, long expected)
    {
        Assert.Equal(expected, ComHostPolicy.HandshakeBudgetFor(deadline));
    }

    /// <summary>Both ends of the pipe handshake are ONE number, not two that agree today.</summary>
    [Fact]
    public void HandshakeBudget_IsSharedByBothEndsOfThePipe()
    {
        Assert.Equal(ComOperationBudgets.HandshakeBudgetMs, (int)ComHostPolicy.HandshakeBudgetMilliseconds);

        // The child declares its own half as a private const; reading it back proves the two
        // ends still resolve to the same value rather than merely being written next to the
        // same comment.
        FieldInfo? childField = typeof(ComHostPolicy).Assembly
            .GetType("OutlookAI.ComHost.Program", throwOnError: false)
            ?.GetField("ConnectTimeoutMs", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(childField);
        Assert.Equal(ComOperationBudgets.HandshakeBudgetMs, (int)childField!.GetRawConstantValue()!);
    }

    /// <summary>
    /// A contract call is shrunk to what is left of the enclosing operation's aggregate, and
    /// a spent aggregate refuses to dispatch rather than sending a deadline so short it would
    /// kill a healthy host.
    /// </summary>
    [Theory]
    // No enclosing aggregate: the call keeps its own deadline.
    [InlineData(120_000L, null, 120_000L)]
    // An untouched aggregate must not shave the call, and a millisecond of wall clock must
    // not turn 4 000 into 3 999 in the message a human reads.
    [InlineData(120_000L, 120_000L, 120_000L)]
    [InlineData(4_000L, 3_999L, 4_000L)]
    [InlineData(120_000L, 119_500L, 120_000L)]
    // A spent aggregate does shrink the call.
    [InlineData(120_000L, 30_000L, 30_000L)]
    [InlineData(120_000L, 29_001L, 30_000L)]
    // The call's own budget still wins when it is the tighter of the two.
    [InlineData(5_000L, 30_000L, 5_000L)]
    // At and below the dispatch floor: dispatch, then refuse.
    [InlineData(120_000L, ComHostPolicy.MinimumDispatchDeadlineMilliseconds, ComHostPolicy.MinimumDispatchDeadlineMilliseconds)]
    [InlineData(120_000L, ComHostPolicy.MinimumDispatchDeadlineMilliseconds - 1, 0L)]
    [InlineData(120_000L, 0L, 0L)]
    public void EffectiveDeadline_ShrinksToTheAggregateAndRefusesBelowTheDispatchFloor(
        long callDeadline, long? remainingAggregate, long expected)
    {
        Assert.Equal(expected, ComHostPolicy.EffectiveDeadlineMilliseconds(callDeadline, remainingAggregate));
    }

    /// <summary>
    /// The raw-EntryID floor is derived from the shortest structurally valid MAPI entry id
    /// rather than picked. The code said 48 while its own comment said 140, with no constant
    /// and no test on either.
    /// </summary>
    [Fact]
    public void RawEntryIdFloor_IsDerivedFromTheMapiEntryIdLength()
    {
        Assert.Equal(EntryIdCodec.MessageEntryIdLength * 2, MailService.MinRawEntryIdHexChars);

        // And it must stay well clear of a hit id, which is what it exists to exclude.
        Assert.True(MailService.MinRawEntryIdHexChars > "h999999".Length * 2);
    }

    /// <summary>
    /// The two caches that describe the same profile expire together. The comment used to
    /// assert this equality; nothing enforced it.
    /// </summary>
    [Fact]
    public void StoreDetailsAndFolderPathCaches_ShareOneTimeToLive()
    {
        System.TimeSpan storeTtl = ReadPrivateTimeSpan("StoreDetailsCacheTtl");
        System.TimeSpan folderTtl = ReadPrivateTimeSpan("FolderPathCacheTtl");
        Assert.Equal(storeTtl, folderTtl);
        Assert.True(storeTtl > System.TimeSpan.Zero);
    }

    private static System.TimeSpan ReadPrivateTimeSpan(string name)
    {
        FieldInfo field = typeof(MailService).GetField(name, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new System.InvalidOperationException($"MailService.{name} not found");
        return (System.TimeSpan)field.GetValue(null)!;
    }

    private static string ToolDescription(string methodName)
    {
        MethodInfo method = typeof(OutlookTools).GetMethod(methodName, BindingFlags.Public | BindingFlags.Static)
            ?? throw new System.InvalidOperationException($"OutlookTools.{methodName} not found");
        return method.GetCustomAttribute<DescriptionAttribute>()?.Description
            ?? throw new System.InvalidOperationException($"OutlookTools.{methodName} carries no [Description]");
    }

    private static string ParameterDescription(string methodName, string parameterName)
    {
        MethodInfo method = typeof(OutlookTools).GetMethod(methodName, BindingFlags.Public | BindingFlags.Static)
            ?? throw new System.InvalidOperationException($"OutlookTools.{methodName} not found");
        ParameterInfo parameter = method.GetParameters().FirstOrDefault(p => p.Name == parameterName)
            ?? throw new System.InvalidOperationException($"OutlookTools.{methodName} has no '{parameterName}' parameter");
        return parameter.GetCustomAttribute<DescriptionAttribute>()?.Description
            ?? throw new System.InvalidOperationException($"'{parameterName}' carries no [Description]");
    }
}
