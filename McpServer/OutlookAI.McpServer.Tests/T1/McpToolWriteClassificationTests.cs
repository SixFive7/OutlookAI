using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
using OutlookAI.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins Q74 layer 3: on the read-only machine the test-side MCP client refuses every tool not
/// classified read-only BEFORE it is sent, so a write through the MCP server process - which the
/// in-process <see cref="StoreWriteAllowlist"/> cannot see - is refused too.
/// <para>
/// Three things are pinned: the classification is COMPLETE against the tools the server assembly
/// actually declares (so a new tool is a deliberate classification - the Q93 audit-log reader is
/// the first); the gate's decision for every posture; and that the shipped client consults it in
/// front of the wire. No server is started and no Outlook is touched: the client tests use a client
/// whose process never started, so anything the gate lets through fails at the transport.
/// </para>
/// </summary>
public sealed class McpToolWriteClassificationTests
{
    /// <summary>
    /// The read-only set, by value. Moving a tool INTO it is the one change that widens what the
    /// read-only machine may send, so it fails here until somebody writes the new name down on purpose.
    /// </summary>
    private static readonly string[] ReadOnlyTools =
    {
        "search", "thread", "read", "list_accounts", "list_folders", "list_signatures", "outlook_health",
    };

    private readonly ITestOutputHelper _output;

    public McpToolWriteClassificationTests(ITestOutputHelper output)
    {
        _output = output;
    }

    // ------------------------------------------------------------------ the classification

    [Fact]
    public void EveryToolTheServerDeclares_IsClassified_AndNothingElseIs()
    {
        List<string> declared = DeclaredToolNames();
        List<string> classified = McpToolWriteClassification.Tools.Keys.OrderBy(n => n, StringComparer.Ordinal).ToList();

        _output.WriteLine("server tools: " + declared.Count + ", classified: " + classified.Count);
        List<string> unclassified = declared.Except(classified, StringComparer.Ordinal).ToList();
        List<string> stale = classified.Except(declared, StringComparer.Ordinal).ToList();

        Assert.True(
            unclassified.Count == 0,
            "The server declares " + string.Join(", ", unclassified.Select(n => "'" + n + "'")) + " and "
            + nameof(McpToolWriteClassification) + ".Tools does not classify it. Add it there with a reason: "
            + nameof(McpToolEffect.ReadOnly) + " ONLY if it changes nothing outside the server process - no mail item or "
            + "folder, no signature, no registry value, no file, nothing on the user's screen - and "
            + nameof(McpToolEffect.Writes) + " otherwise. Until then the read-only machine refuses it (Q74 layer 3).");
        Assert.True(stale.Count == 0, "classified but no longer declared by the server: " + string.Join(", ", stale));

        // A reflection walk that found nothing would pass both checks above.
        Assert.True(declared.Count >= 21, "found only " + declared.Count + " [McpServerTool] methods - the scan stopped working");
    }

    [Fact]
    public void TheReadOnlySet_IsExactlyTheOneWrittenDownHere()
    {
        Assert.Equal(
            ReadOnlyTools.OrderBy(n => n, StringComparer.Ordinal),
            McpToolWriteClassification.Tools
                .Where(t => t.Value.Effect == McpToolEffect.ReadOnly)
                .Select(t => t.Key)
                .OrderBy(n => n, StringComparer.Ordinal));
    }

    [Fact]
    public void EveryEntry_IsClassifiedOneWayOrTheOther_AndSaysWhy()
    {
        Assert.All(McpToolWriteClassification.Tools, entry =>
        {
            Assert.NotEqual(McpToolEffect.Unclassified, entry.Value.Effect);
            Assert.False(string.IsNullOrWhiteSpace(entry.Value.Why), entry.Key + " carries no reason");
        });
    }

    // ------------------------------------------------------------------ the decision

    public static IEnumerable<object[]> WriteCapableTools()
    {
        return McpToolWriteClassification.Tools
            .Where(t => t.Value.Effect == McpToolEffect.Writes)
            .Select(t => new object[] { t.Key });
    }

    [Theory]
    [MemberData(nameof(WriteCapableTools))]
    public void UnderTheReadOnlyPosture_AWriteCapableToolIsRefused(string tool)
    {
        string? refusal = McpToolWriteClassification.RefusalFor(StdioWritePosture.ReadOnly, tool);

        Assert.NotNull(refusal);
        Assert.StartsWith(McpToolWriteClassification.Refused, refusal, StringComparison.Ordinal);
        Assert.Contains("'" + tool + "'", refusal, StringComparison.Ordinal);
        Assert.Contains("test guest", refusal, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("read_audit_log")]
    [InlineData("SEND")]
    [InlineData("")]
    [InlineData(null)]
    public void UnderTheReadOnlyPosture_AnUnclassifiedToolIsRefused_LikeAWrite(string? tool)
    {
        // A tool added after the list was last read, a misspelling, a name in the wrong case: none
        // of them has been shown to change nothing, so none of them goes out.
        string? refusal = McpToolWriteClassification.RefusalFor(StdioWritePosture.ReadOnly, tool);

        Assert.NotNull(refusal);
        Assert.Contains("not classified at all", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void UnderTheReadOnlyPosture_EveryReadOnlyToolGoesOut()
    {
        Assert.All(ReadOnlyTools, tool => Assert.Null(McpToolWriteClassification.RefusalFor(StdioWritePosture.ReadOnly, tool)));
    }

    [Theory]
    [InlineData(StdioWritePosture.Writable)]
    [InlineData(StdioWritePosture.NotALiveRun)]
    public void UnderAnyOtherPosture_TheGateRefusesNothing(StdioWritePosture posture)
    {
        foreach (string tool in McpToolWriteClassification.Tools.Keys.Append("read_audit_log"))
        {
            Assert.Null(McpToolWriteClassification.RefusalFor(posture, tool));
        }
    }

    // ------------------------------------------------------------------ the client

    [Theory]
    [InlineData("send")]
    [InlineData("new_draft")]
    [InlineData("move_mail")]
    [InlineData("manage_signature")]
    [InlineData("read_audit_log")]
    public async Task OnTheReadOnlyMachine_TheClientRefusesBeforeSending(string tool)
    {
        // Control: before Q74 the client had no write gate, so the call went to the transport and
        // failed there instead ("StandardIn has not been redirected" on this never-started process),
        // and the assertion on the refusal phrase failed.
        await using McpStdioClient client = McpStdioClient.Unstarted(StdioWritePosture.ReadOnly, outlookReachingToolsAllowed: true);

        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync(tool, new { id = "unit" }));

        Assert.StartsWith(McpToolWriteClassification.Refused, refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task OnTheReadOnlyMachine_ARawToolsCallEnvelope_IsRefusedToo()
    {
        // The supervision and availability tests build their own envelopes and hand them to
        // RoundTripAsync; the gate sits there, not in the convenience helpers.
        await using McpStdioClient client = McpStdioClient.Unstarted(StdioWritePosture.ReadOnly, outlookReachingToolsAllowed: true);

        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.RoundTripAsync("tools/call", new { name = "discard_draft", arguments = new { id = "unit" } }));
        Assert.StartsWith(McpToolWriteClassification.Refused, refused.Message, StringComparison.Ordinal);

        InvalidOperationException notified = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.NotifyAsync("tools/call"));
        Assert.StartsWith(McpToolWriteClassification.Refused, notified.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(StdioWritePosture.ReadOnly, "read")]
    [InlineData(StdioWritePosture.ReadOnly, "outlook_health")]
    [InlineData(StdioWritePosture.Writable, "send")]
    [InlineData(StdioWritePosture.NotALiveRun, "new_draft")]
    public async Task WhatTheGateLetsThrough_ReachesTheTransport(StdioWritePosture posture, string tool)
    {
        // The other half: the gate is selective. These pass it and fail only because no server
        // process stands behind this client - which is also the proof that nothing was sent.
        await using McpStdioClient client = McpStdioClient.Unstarted(posture, outlookReachingToolsAllowed: true);

        InvalidOperationException failed = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync(tool, new { id = "unit" }));

        Assert.DoesNotContain(McpToolWriteClassification.Refused, failed.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheWriteGate_ComesBeforeTheContactGate()
    {
        // A client that did NOT declare mailbox contact, asked for a write-capable tool on the
        // read-only machine, is refused for the write - the more fundamental of the two reasons.
        await using McpStdioClient client = McpStdioClient.Unstarted(StdioWritePosture.ReadOnly, outlookReachingToolsAllowed: false);

        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync("send", new { id = "unit" }));
        Assert.StartsWith(McpToolWriteClassification.Refused, refused.Message, StringComparison.Ordinal);

        // And a read-only tool that always reaches Outlook still meets the contact gate.
        InvalidOperationException undeclared = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync("list_accounts", new { }));
        Assert.Contains(nameof(McpStdioClient.OutlookReachingToolsAllowed), undeclared.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheShippedClient_TakesItsPostureFromTheMachine_NeverFromTheTest()
    {
        // StartAndInitializeAsync is the only way a test gets a client that talks to a server, and
        // it must read LiveWriteAccess.CurrentStdioPosture itself: a posture a test could pass in is
        // a posture a test could pass wrong. Read out of the compiled state machine.
        MethodInfo start = typeof(McpStdioClient).GetMethod(nameof(McpStdioClient.StartAndInitializeAsync))!;
        Assert.DoesNotContain(start.GetParameters(), p => p.ParameterType == typeof(StdioWritePosture));

        bool readsThePosture = IlReader.WithNestedTypes(typeof(McpStdioClient))
            .SelectMany(IlReader.BodiesOf)
            .SelectMany(IlReader.Read)
            .Any(i => i.Method?.DeclaringType == typeof(LiveWriteAccess)
                && i.Method.Name == "get_" + nameof(LiveWriteAccess.CurrentStdioPosture));
        Assert.True(readsThePosture, "McpStdioClient no longer reads LiveWriteAccess.CurrentStdioPosture");
    }

    [Fact]
    public void NoLiveTestSpawnsTheServer_OutsideTheGatedClient()
    {
        // The gate covers McpStdioClient and its callers. A live class that started the server exe
        // itself would reach it around the gate, so no live class may name the exe path at all -
        // neither the assembly-metadata key nor the client's property for it.
        List<string> problems = new();
        foreach (Type type in typeof(LiveCollections).Assembly.GetTypes().Where(IsLiveClass))
        {
            if (IlReader.StringsOf(type).Contains("McpServerExePath"))
            {
                problems.Add(type.Name + " names the McpServerExePath metadata key");
            }

            if (IlReader.MethodsNamedBy(type).Any(m => m.DeclaringType == typeof(McpStdioClient)
                && m.Name == "get_" + nameof(McpStdioClient.ServerExePath)))
            {
                problems.Add(type.Name + " reads McpStdioClient.ServerExePath");
            }
        }

        Assert.Empty(problems);
    }

    // ------------------------------------------------------------------ plumbing

    /// <summary>The wire name of every [McpServerTool] method in the server assembly.</summary>
    private static List<string> DeclaredToolNames()
    {
        return typeof(OutlookTools).Assembly.GetTypes()
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            .SelectMany(m => m.GetCustomAttributesData())
            .Where(a => a.AttributeType.Name == "McpServerToolAttribute")
            .Select(a => a.NamedArguments.FirstOrDefault(n => n.MemberName == "Name").TypedValue.Value as string)
            .Where(name => !string.IsNullOrEmpty(name))
            .Select(name => name!)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();
    }

    private static bool IsLiveClass(Type type)
    {
        return type.IsClass && type.GetCustomAttributesData().Any(a => a.AttributeType == typeof(TraitAttribute)
            && a.ConstructorArguments.Count == 2
            && string.Equals(a.ConstructorArguments[0].Value as string, "Category", StringComparison.Ordinal)
            && string.Equals(a.ConstructorArguments[1].Value as string, "Live", StringComparison.Ordinal));
    }
}
