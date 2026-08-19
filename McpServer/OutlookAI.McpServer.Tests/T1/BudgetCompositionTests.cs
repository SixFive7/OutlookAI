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
    /// The inner scan budget is strictly INSIDE the hard deadline of its OWN class, by the
    /// declared headroom. Equality is the defect: the scan stops only once elapsed has
    /// PASSED its budget, then still has to serialize its results back over the pipe, while
    /// the watchdog fires at &gt;=.
    /// <para>
    /// The class is the second half of what this pins. The scan is the one operation whose
    /// budget expiry is a documented ANSWER, so it is allowed to be long - and every OTHER
    /// tool must keep the ordinary hang detector, which is only true while the two numbers
    /// are different. A future change that "simplified" this by dropping the class and
    /// raising the shared deadline would give <c>read</c> and <c>move_mail</c> a ten-minute
    /// wait to discover a wedged Outlook, and nothing else would notice.
    /// </para>
    /// </summary>
    [Fact]
    public void ExhaustiveScanBudget_LeavesHeadroomInsideItsOwnDeadlineClass()
    {
        Assert.True(
            MailService.ExhaustiveTimeBudgetMs < ComOperationBudgets.ExhaustiveScanDeadlineMs,
            $"the exhaustive scan's soft budget ({MailService.ExhaustiveTimeBudgetMs} ms) must be strictly inside the COM "
            + $"host's hard deadline for that class ({ComOperationBudgets.ExhaustiveScanDeadlineMs} ms); equal means the "
            + "documented partial-results outcome is unreachable and a long scan becomes a timeout plus a host kill");

        Assert.Equal(
            ComOperationBudgets.ExhaustiveScanDeadlineMs - ComOperationBudgets.ResultReturnHeadroomMs,
            MailService.ExhaustiveTimeBudgetMs);

        Assert.True(
            ComOperationBudgets.ResultReturnHeadroomMs > 0,
            "the return-trip headroom must be positive - it is the whole mechanism");

        // The class exists and carries that deadline, rather than the number living only in
        // the one call site that passes it.
        Assert.Equal(
            ComOperationBudgets.ExhaustiveScanDeadlineMs,
            (int)ComHostPolicy.DeadlineFor(ComHostOperationClass.ExhaustiveScan, null));

        // And it is a class of its OWN: the long scan must not be the price every other
        // tool pays for its hang detection.
        Assert.True(
            ComOperationBudgets.ExhaustiveScanDeadlineMs > ComOperationBudgets.OperationDeadlineMs,
            $"the exhaustive class ({ComOperationBudgets.ExhaustiveScanDeadlineMs} ms) exists precisely because it is "
            + $"longer than the ordinary operation deadline ({ComOperationBudgets.OperationDeadlineMs} ms); equal or "
            + "shorter means the class buys nothing and should not exist");
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
    /// The search result cap and its default are ONE pair of numbers, and the argument that
    /// advertises them quotes THAT pair.
    /// <para>
    /// <c>"1-100, default 25"</c> is prose in an attribute literal while the enforcement is
    /// <see cref="MailService.SearchTopCap"/> and <see cref="MailService.SearchTopDefault"/>
    /// - the same unguarded shape the exhaustive budget and the subject cap were in when
    /// each of them drifted. A wire pin exists (T3 SearchSchemaCiTests) but it asserts the
    /// LITERAL "1-100", so raising the cap would leave that test green over a description
    /// that lies. This one fails the moment the numbers stop agreeing, and names both.
    /// </para>
    /// <para>
    /// The floor is a genuine literal on both sides (<c>Clamp(request.Top, 1, ...)</c>), so
    /// it is written here as one too.
    /// </para>
    /// </summary>
    [Fact]
    public void SearchTopRange_IsQuotedTruthfullyInTheToolSurface()
    {
        string expected = "(1-" + MailService.SearchTopCap.ToString(CultureInfo.InvariantCulture)
            + ", default " + MailService.SearchTopDefault.ToString(CultureInfo.InvariantCulture) + ")";
        string hint = ParameterDescription(nameof(OutlookTools.Search), "top");

        Assert.True(
            hint.Contains(expected, System.StringComparison.Ordinal),
            $"search's 'top' description must quote MailService.SearchTopCap ({MailService.SearchTopCap}) and "
            + $"SearchTopDefault ({MailService.SearchTopDefault}) as \"{expected}\", because those are the values "
            + $"the service actually clamps to. The description reads: \"{hint}\"");

        // And the default the SCHEMA advertises is the same number, not a third copy.
        Assert.Equal(MailService.SearchTopDefault, DefaultValue<int>(nameof(OutlookTools.Search), "top"));
    }

    /// <summary>
    /// The attachment set's count and size caps are quoted from the constants that enforce
    /// them, on every tool that advertises them.
    /// <para>
    /// Both numbers are fail-closed refusals (<see cref="DraftAttachments.Validate"/> throws
    /// and attaches nothing), so a description that quotes a stale limit teaches an agent to
    /// build a call that cannot succeed - and the megabyte figure is doubly exposed, being a
    /// unit conversion of a byte constant rather than the constant itself.
    /// </para>
    /// </summary>
    [Fact]
    public void AttachmentCaps_AreQuotedTruthfullyInTheToolSurface()
    {
        string expected = "Max " + DraftAttachments.MaxFiles.ToString(CultureInfo.InvariantCulture)
            + " files and " + (DraftAttachments.MaxTotalBytes / (1024 * 1024)).ToString(CultureInfo.InvariantCulture)
            + " MB";

        foreach (string tool in new[]
                 {
                     nameof(OutlookTools.NewDraft),
                     nameof(OutlookTools.ReplyDraft),
                     nameof(OutlookTools.ReplyAllDraft),
                     nameof(OutlookTools.ForwardDraft),
                     nameof(OutlookTools.UpdateDraft),
                 })
        {
            string hint = ParameterDescription(tool, "attachments");
            Assert.True(
                hint.Contains(expected, System.StringComparison.Ordinal),
                $"{tool}'s 'attachments' description must quote DraftAttachments.MaxFiles ({DraftAttachments.MaxFiles}) "
                + $"and MaxTotalBytes ({DraftAttachments.MaxTotalBytes} bytes) as \"{expected}\", because those are the "
                + $"limits the validation actually refuses on. The description reads: \"{hint}\"");
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

        // STRICTLY below, and this assertion used to be Assert.Equal - i.e. the test
        // enforced the defect. An aggregate equal to its own unit of work is not an
        // aggregate: the budget is checked BEFORE each item, so a batch one millisecond
        // inside it could start one more item carrying a full operation deadline of its own
        // and run to twice the budget. The items are now dispatched with what is LEFT of
        // this budget, and this inequality is what keeps "the batch ran long" and "Outlook
        // stopped answering" two different events.
        Assert.True(
            MailService.MoveBatchBudgetMs < ComOperationBudgets.OperationDeadlineMs,
            $"the move batch budget ({MailService.MoveBatchBudgetMs} ms) must be strictly inside the COM host's hard "
            + $"operation deadline ({ComOperationBudgets.OperationDeadlineMs} ms)");

        Assert.True(
            MailService.MoveBatchBudgetMs < (long)MailService.MoveIdsCap * ComOperationBudgets.OperationDeadlineMs,
            "the batch budget must be smaller than the worst case it replaces");

        // And an item is never dispatched on a budget too small to be dispatched at all -
        // below this the item is reported as not attempted, which is legible, instead of
        // being refused by the COM host's own dispatch floor as a bare timeout.
        Assert.True(
            MailService.MinimumItemBudgetMs > 0 && MailService.MinimumItemBudgetMs < MailService.MoveBatchBudgetMs,
            "the per-item floor must be positive and inside the batch budget");
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

    /// <summary>
    /// A budget the CALLER declared outranks the handshake floor, on the same terms as it
    /// already outranks the cold-start connect floor.
    /// <para>
    /// The defect this pins: <c>outlook_health</c>'s description promises "gives up after
    /// 5 s", it asks the gateway for exactly 5 s - and the handshake, which runs BEFORE the
    /// deadline clock starts, took 10 s of floor anyway, twice over because health makes two
    /// gateway calls. The one tool that must always answer had the longest cold-host
    /// preamble in the product.
    /// </para>
    /// </summary>
    [Theory]
    // Health's own budget: honoured, not floored.
    [InlineData(ComOperationBudgets.HealthProbeDeadlineMs, ComOperationBudgets.HealthProbeDeadlineMs)]
    [InlineData(1L, 1L)]
    [InlineData(ComHostPolicy.HandshakeFloorMilliseconds, ComHostPolicy.HandshakeFloorMilliseconds)]
    // Above the floor nothing changes - the floor was never the binding rule there.
    [InlineData(20_000L, 20_000L)]
    // The ceiling still binds: a caller may not declare its way past one handshake budget.
    [InlineData(ComHostPolicy.HandshakeBudgetMilliseconds + 1, ComHostPolicy.HandshakeBudgetMilliseconds)]
    [InlineData(600_000L, ComHostPolicy.HandshakeBudgetMilliseconds)]
    // "No budget" must never become "no handshake".
    [InlineData(0L, ComHostPolicy.HandshakeFloorMilliseconds)]
    [InlineData(-1L, ComHostPolicy.HandshakeFloorMilliseconds)]
    public void HandshakeBudget_GivesWayToABudgetTheCallerDeclared(long deadline, long expected)
    {
        Assert.Equal(expected, ComHostPolicy.HandshakeBudgetFor(deadline, callerDeclaredBudget: true));

        // And the floor is untouched for everyone who did NOT declare one - which is the
        // test suite shortening the deadline to observe the timeout path, not the start path.
        Assert.Equal(
            ComHostPolicy.HandshakeFloorMilliseconds,
            ComHostPolicy.HandshakeBudgetFor(ComOperationBudgets.HealthProbeDeadlineMs, callerDeclaredBudget: false));
    }

    /// <summary>
    /// Only a HANG is evidence of a hang. A caller-declared work budget expiring says the
    /// work was big, and must not count toward the breaker.
    /// <para>
    /// The outage this prevents: the freshness sweep runs on an explicit budget on every
    /// search, so two ordinary slow searches on a large mailbox opened the breaker and every
    /// COM request then failed fast for the whole cooldown - caused by nothing but the size
    /// of the mailbox.
    /// </para>
    /// </summary>
    [Fact]
    public void OnlyAHangDetectorExpiring_CountsTowardTheBreaker()
    {
        // The sweep and the thread walk: explicit budgets, below the class deadline.
        Assert.False(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComHostOperationClass.Operation, MailService.SweepBudgetMs));
        Assert.False(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComHostOperationClass.Operation, MailService.ThreadWalkBudgetMs));
        Assert.False(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComHostOperationClass.Operation, MailService.MoveBatchBudgetMs));

        // No explicit budget at all: this IS the hang detector.
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComHostOperationClass.Operation, null));
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComHostOperationClass.Operation, 0));
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComHostOperationClass.Connect, null));

        // The health probe is the instrument: it is dispatched precisely to find out, so its
        // short explicit budget expiring is the answer rather than a work limit.
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComHostOperationClass.HealthProbe, ComOperationBudgets.HealthProbeDeadlineMs));

        // The exhaustive scan asks for its own class deadline, and its INNER budget stops it
        // gracefully long before - so reaching the outer one really is a wedge.
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComHostOperationClass.ExhaustiveScan, ComOperationBudgets.ExhaustiveScanDeadlineMs));
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

    private static T DefaultValue<T>(string methodName, string parameterName)
    {
        MethodInfo method = typeof(OutlookTools).GetMethod(methodName, BindingFlags.Public | BindingFlags.Static)
            ?? throw new System.InvalidOperationException($"OutlookTools.{methodName} not found");
        ParameterInfo parameter = method.GetParameters().FirstOrDefault(p => p.Name == parameterName)
            ?? throw new System.InvalidOperationException($"OutlookTools.{methodName} has no '{parameterName}' parameter");
        return (T)parameter.DefaultValue!;
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
