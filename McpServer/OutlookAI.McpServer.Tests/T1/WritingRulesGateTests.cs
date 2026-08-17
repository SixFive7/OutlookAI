using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The writing-rules gate (<c>OutlookAI.McpServer/Tools/WritingRulesGate.cs</c>), driven
/// against an INJECTED rules source so every branch is reachable without the developer's own
/// HKCU, without a server process, and without Outlook.
/// <para>
/// The gate exists because the user's writing rules cannot travel on the MCP schema - a
/// description is a compile-time constant and is capped at 2 KB by the client, while the rules
/// are user-edited, can run to pages, and can change between two tool calls. They travel in a
/// rejection instead: the first drafting call of a server process is refused with the rules
/// attached, and the retry goes through.
/// </para>
/// <para>
/// So the property this file pins is a cost, and it is the user's accepted cost: ONE rejection
/// per session, and one more per edit. Not two, not one per call, and never one for a call
/// that was not writing anything - each of those would be a tax on every draft the agent makes.
/// </para>
/// </summary>
public sealed class WritingRulesGateTests
{
    private const string Rules =
        "You are an email writing assistant.\r\n\r\nOutput format:\r\n- Return only the draft text.";

    /// <summary>A rules source whose text the test can change, as the settings dialog does.</summary>
    private sealed class FakeRules
    {
        internal FakeRules(string text)
        {
            Text = text;
        }

        internal string Text { get; set; }

        internal int Reads { get; private set; }

        internal Func<string> Source => () =>
        {
            Reads++;
            return Text;
        };
    }

    [Fact]
    public void FirstDraftingCall_IsGated_AndCarriesTheRulesVerbatim()
    {
        var rules = new FakeRules(Rules);
        var gate = new WritingRulesGate(rules.Source);

        Assert.True(gate.TryClaimDelivery(composesBody: true, out string delivered));

        // Verbatim: what the user typed is what the agent gets, line endings and all. The
        // user chose that over any filtered or reformatted variant.
        Assert.Equal(Rules, delivered);
    }

    [Fact]
    public void TheRetry_IsNotGated()
    {
        var rules = new FakeRules(Rules);
        var gate = new WritingRulesGate(rules.Source);

        Assert.True(gate.TryClaimDelivery(composesBody: true, out _));

        Assert.False(gate.TryClaimDelivery(composesBody: true, out string second));
        Assert.Equal(string.Empty, second);

        // And it stays that way - one rejection per session, not one per call.
        for (int i = 0; i < 5; i++)
        {
            Assert.False(gate.TryClaimDelivery(composesBody: true, out _));
        }
    }

    [Fact]
    public void EditedRules_CostExactlyOneMoreRejection()
    {
        var rules = new FakeRules(Rules);
        var gate = new WritingRulesGate(rules.Source);
        Assert.True(gate.TryClaimDelivery(composesBody: true, out _));
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));

        rules.Text = Rules + "\r\n- Always sign off with 'Kind regards'.";

        // The gate re-arms on the text itself, so an edit made while the server is running
        // reaches the agent without anything being restarted...
        Assert.True(gate.TryClaimDelivery(composesBody: true, out string delivered));
        Assert.Equal(rules.Text, delivered);

        // ...and costs one rejection, not one per call from then on.
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));
    }

    [Fact]
    public void RulesThatOnlyChangedTheirLineEndings_DoNotCostASecondRejection()
    {
        var rules = new FakeRules(Rules);
        var gate = new WritingRulesGate(rules.Source);
        Assert.True(gate.TryClaimDelivery(composesBody: true, out _));

        // A multiline text box hands back CRLF where the source had LF, and usually a
        // trailing newline with it. Not one word of what the model reads has changed, so it
        // may not cost the agent another failed call.
        rules.Text = Rules.Replace("\r\n", "\n") + "\n";
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));

        rules.Text = "  " + Rules.Replace("\r\n", "\r") + "  ";
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));
    }

    [Fact]
    public void ACallThatWritesNoBody_IsNeverGated_AndDoesNotSpendTheDelivery()
    {
        var rules = new FakeRules(Rules);
        var gate = new WritingRulesGate(rules.Source);

        // update_draft may legitimately change only recipients, a subject or attachments.
        Assert.False(gate.TryClaimDelivery(composesBody: false, out string nothing));
        Assert.Equal(string.Empty, nothing);
        Assert.Equal(0, rules.Reads);

        // And it must not have armed the gate behind the agent's back: the next call that
        // DOES write a body is still owed the rules.
        Assert.True(gate.TryClaimDelivery(composesBody: true, out string delivered));
        Assert.Equal(Rules, delivered);
    }

    [Fact]
    public async Task ConcurrentFirstCalls_GateExactlyOneOfThem()
    {
        // Tool calls overlap; two agents' first drafts could arrive at once. Both being let
        // through is the failure that matters (nobody sees the rules), and both being
        // rejected is a wasted call, so the answer has to be exactly one.
        for (int attempt = 0; attempt < 25; attempt++)
        {
            var rules = new FakeRules(Rules);
            var gate = new WritingRulesGate(rules.Source);

            const int Callers = 16;
            using var start = new ManualResetEventSlim(false);
            var claims = new bool[Callers];
            var callers = new Task[Callers];

            for (int i = 0; i < Callers; i++)
            {
                int index = i;
                callers[i] = Task.Run(() =>
                {
                    start.Wait();
                    claims[index] = gate.TryClaimDelivery(composesBody: true, out _);
                });
            }

            start.Set();
            await Task.WhenAll(callers);

            Assert.Equal(1, claims.Count(claimed => claimed));
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("   \r\n  ")]
    public void RulesTheUserCleared_AreNeverDelivered(string cleared)
    {
        var gate = new WritingRulesGate(new FakeRules(cleared).Source);

        // Nothing to deliver means nothing to refuse a draft for. A user who empties the
        // section has removed the rules, not asked for a rejection carrying nothing.
        Assert.False(gate.TryClaimDelivery(composesBody: true, out string delivered));
        Assert.Equal(string.Empty, delivered);
    }

    [Fact]
    public void ARulesSourceThatFails_CostsTheRulesAndNotTheDraft()
    {
        var gate = new WritingRulesGate(() => throw new InvalidOperationException("fabricated read failure"));

        // Failing open: the draft is written without the rules. Failing closed would mean an
        // unreadable registry key stops the user drafting anything at all, and would do it on
        // every call, forever.
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));
    }

    [Fact]
    public void TheRulesAreReadOnEveryGatedCall_NotCachedAtStartup()
    {
        var rules = new FakeRules(Rules);
        var gate = new WritingRulesGate(rules.Source);

        Assert.True(gate.TryClaimDelivery(composesBody: true, out _));
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));
        Assert.False(gate.TryClaimDelivery(composesBody: true, out _));

        // Three reads for three calls: an edit is noticed on the next call, not on the next
        // server start. That is the whole reason a mid-session edit costs one rejection.
        Assert.Equal(3, rules.Reads);
    }
}
