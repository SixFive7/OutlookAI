using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="ToolFailure.Describe"/> - the tool's FATAL line naming the CAUSE of a failure, not
/// only the wrapper the STA runner puts round it. On OAI-UNINDEXED, 2026-09-27, a probe ended in
/// "FATAL: InvalidOperationException: corpus undated probe failed." with the cause dropped.
/// </summary>
public sealed class ToolFailureTests
{
    [Fact]
    public void TheWholeChain_IsNamed_OuterFirst()
    {
        var cause = new System.Runtime.InteropServices.COMException("The property cannot be deleted.", unchecked((int)0x80070005));
        var wrapped = new InvalidOperationException("corpus undated probe failed.", cause);
        Assert.Equal(
            "InvalidOperationException: corpus undated probe failed. <- caused by: COMException: The property cannot be deleted.",
            ToolFailure.Describe(wrapped));
    }

    [Fact]
    public void ALoneException_IsNamedAsItWasBefore_TheControl()
    {
        // The control: an exception with no cause reads exactly like the line the tool always printed.
        Assert.Equal("ArgumentException: --store <display name> is required.", ToolFailure.Describe(new ArgumentException("--store <display name> is required.")));
    }

    [Fact]
    public void AnAggregate_IsFollowedIntoItsFirstInnerException()
    {
        var aggregate = new AggregateException("outer", new NotSupportedException("first"), new TimeoutException("second"));
        string text = ToolFailure.Describe(aggregate);
        Assert.Contains("<- caused by: NotSupportedException: first", text, StringComparison.Ordinal);

        // The aggregate's own message already quotes every inner message; only the first is FOLLOWED.
        Assert.DoesNotContain("TimeoutException", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ADeepChain_IsCut_AndSaysSo()
    {
        Exception ex = new InvalidOperationException("0");
        for (int i = 1; i < ToolFailure.MaxDepth + 3; i++)
        {
            ex = new InvalidOperationException(i.ToString(System.Globalization.CultureInfo.InvariantCulture), ex);
        }

        string text = ToolFailure.Describe(ex);
        Assert.EndsWith(" <- caused by: ...", text, StringComparison.Ordinal);
        Assert.Equal(ToolFailure.MaxDepth + 1, text.Split(" <- caused by: ").Length);
    }
}
