using System.Text.Json;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 live acceptance for the Phase-3 tools (v3.MD section 0.6): open_in_outlook,
/// goto_folder, show_search_results and exhaustive search called over REAL stdio
/// MCP against the built server exe with golden-shape asserts. All UI targets the
/// test-hub store (S2/S5). The test uses its own COM session to verify/close windows
/// the server opened on its behalf (closing test-caused windows is allowed, S7) and to
/// clear the search UI afterwards. Output stays content-free for business stores (S4).
/// </summary>
[Trait("Category", "Live")]
public sealed class Phase3LiveMcpToolShapeTests
{
    private readonly ITestOutputHelper _output;
    private readonly LiveTestSettings _settings;

    public Phase3LiveMcpToolShapeTests(ITestOutputHelper output)
    {
        _output = output;
        _settings = LiveTestSettings.Load();
    }

    [Fact]
    public async Task ShowMe_And_ExhaustiveSearch_GoldenShapes_OverRealStdio()
    {
        string hub = _settings.TestHubStoreDisplayName;
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(6));
        using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: true);

        // --- goto_folder: golden shape + explorer folder path.
        JsonElement gone = await client.CallToolAsync("goto_folder", new { store = hub });
        Assert.True(gone.GetProperty("displayed").GetBoolean());
        string? explorerPath = gone.GetProperty("explorerFolderPath").GetString();
        Assert.False(string.IsNullOrEmpty(explorerPath));
        Assert.StartsWith("\\\\" + hub, explorerPath!, StringComparison.OrdinalIgnoreCase);
        _output.WriteLine($"goto_folder: displayed=true pathInHub=true");

        // --- open_in_outlook: search a hub hit, display it, verify + close the Inspector.
        JsonElement search = await client.CallToolAsync("search", new
        {
            store = hub,
            include_attachment_hits = false,
            top = 10,
        });
        JsonElement hits = search.GetProperty("hits");
        Assert.True(hits.GetArrayLength() >= 1, "no hub hits for open_in_outlook");

        // Skip test artifacts: deleted tagged items keep their index rows alive
        // (IncludeDeletedItems=1, Phase-2 fact 9) and cannot be opened - when the
        // Phase-4 draft tests run earlier in the same suite, such a row can be the
        // newest hub hit.
        string? hitId = null;
        foreach (JsonElement hit in hits.EnumerateArray())
        {
            string? subject = hit.TryGetProperty("subject", out JsonElement subjectProp) ? subjectProp.GetString() : null;
            if (subject == null || subject.IndexOf("[OutlookAI-McpTest]", StringComparison.OrdinalIgnoreCase) < 0)
            {
                hitId = hit.GetProperty("id").GetString();
                break;
            }
        }

        Assert.True(hitId != null, "no non-artifact hub hits for open_in_outlook");

        JsonElement opened = await client.CallToolAsync("open_in_outlook", new { id = hitId });
        Assert.True(opened.GetProperty("displayed").GetBoolean());
        string entryId = opened.GetProperty("entryId").GetString()!;
        Assert.True(entryId.Length >= 48, "open_in_outlook must return the real EntryID");

        try
        {
            bool inspectorSeen = false;
            for (int i = 0; i < 30 && !inspectorSeen; i++)
            {
                inspectorSeen = session.GetOpenInspectors()
                    .Any(x => x.EntryId != null && string.Equals(x.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
                if (!inspectorSeen)
                {
                    await Task.Delay(500);
                }
            }

            Assert.True(inspectorSeen, "no Inspector for the opened EntryID within 15 s");
            _output.WriteLine("open_in_outlook: inspector verified for the requested EntryID");
        }
        finally
        {
            session.TryCloseInspectorByEntryId(entryId, out _);
        }

        // --- show_search_results: drive the search UI (nonsense query - nothing shown).
        JsonElement shown = await client.CallToolAsync("show_search_results", new
        {
            query = "OutlookAiMcpNoSuchTerm7391",
            scope = "current_folder",
            store = hub,
        });
        Assert.True(shown.GetProperty("displayed").GetBoolean());
        Assert.Equal("current_folder", shown.GetProperty("scope").GetString());
        _output.WriteLine("show_search_results: displayed=true scope echoed");
        await Task.Delay(500);
        session.TryClearSearch(out _);

        // --- search exhaustive=true: golden shape (engine block, source, no sweep, no
        // mode field - D34).
        JsonElement exhaustive = await client.CallToolAsync("search", new
        {
            query = _settings.ProbeTerm,
            exhaustive = true,
            store = hub,
            after = "2000-01-01",
            top = 50,
        });
        Assert.False(exhaustive.TryGetProperty("mode", out _), "search results must not carry a mode field (D34)");
        JsonElement info = exhaustive.GetProperty("exhaustive");
        Assert.False(string.IsNullOrEmpty(info.GetProperty("engine").GetString()));
        Assert.True(info.GetProperty("foldersScanned").GetInt32() >= 1);
        Assert.True(info.TryGetProperty("instantSearchEnabled", out _));
        Assert.False(exhaustive.TryGetProperty("sweep", out _), "exhaustive must not carry sweep diagnostics");
        foreach (JsonElement hit in exhaustive.GetProperty("hits").EnumerateArray())
        {
            Assert.Equal("exhaustive", hit.GetProperty("source").GetString());
        }

        _output.WriteLine($"search exhaustive: engine={info.GetProperty("engine").GetString()} "
            + $"folders={info.GetProperty("foldersScanned").GetInt32()} hits={exhaustive.GetProperty("hits").GetArrayLength()}");

        Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
    }
}
