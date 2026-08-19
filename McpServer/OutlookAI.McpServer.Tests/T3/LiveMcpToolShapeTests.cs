using System.Text.Json;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 live acceptance (v3.MD section 0.6 Phase 2): EVERY L1/L2 tool called over REAL
/// stdio MCP against the built server exe with golden-shape asserts on live data. One
/// spawned server per test so hit ids stay within their own session cache. May start
/// Outlook (S7/D17). Output stays content-free for business stores (S4).
/// </summary>
[Trait("Category", "Live")]
[Trait("LiveTier", "ProfileBound")]
[Trait("Requires", "SearchIndex")]
[Trait("Requires", "MultipleStores")]
[Trait("Requires", "DelegateStore")]
[Collection(LiveCollections.McpToolShape)]
public sealed class LiveMcpToolShapeTests
{
    private readonly ITestOutputHelper _output;
    private readonly LiveTestSettings _settings;

    public LiveMcpToolShapeTests(ITestOutputHelper output)
    {
        _output = output;
        _settings = LiveTestSettings.Load();
    }

    [Fact]
    public async Task Search_Read_SaveAttachment_Thread_GoldenShapes_OverRealStdio()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(6));

        // --- search: hits with ids, staleness block, sweep diagnostics (D34: always
        // fresh - no mode field on the wire in either direction).
        JsonElement search = await client.CallToolAsync("search", new
        {
            query = _settings.ProbeTerm,
            top = 10,
        });
        JsonElement hits = search.GetProperty("hits");
        Assert.True(hits.GetArrayLength() >= 1, "probe term must hit");
        Assert.False(search.TryGetProperty("mode", out _), "search results must not carry a mode field (D34)");
        JsonElement firstIndexHit = hits.EnumerateArray()
            .First(h => h.GetProperty("source").GetString() == "index");
        string hitId = firstIndexHit.GetProperty("id").GetString()!;
        Assert.Matches("^h[0-9]+$", hitId);
        Assert.False(string.IsNullOrEmpty(firstIndexHit.GetProperty("store").GetString()));
        Assert.True(search.GetProperty("indexElapsedMs").GetInt64() >= 0);
        Assert.True(search.GetProperty("staleness").TryGetProperty("outlookRunning", out _));
        JsonElement sweep = search.GetProperty("sweep");
        Assert.True(sweep.GetProperty("performed").GetBoolean(), "gap sweep must run live");
        Assert.True(sweep.GetProperty("foldersSwept").GetInt32() >= 1);
        _output.WriteLine($"search: hits={hits.GetArrayLength()} indexMs={search.GetProperty("indexElapsedMs").GetInt64()} "
            + $"sweepMs={sweep.GetProperty("elapsedMs").GetInt64()} folders={sweep.GetProperty("foldersSwept").GetInt32()}");

        // --- search again (store-scoped, rapid): the D34 sweep cache serves the sweep.
        // A frontier advance between the two calls (new mail indexed on this live
        // machine) legitimately invalidates the cache, so retry the rapid pair.
        bool cachedProven = false;
        for (int attempt = 0; attempt < 3 && !cachedProven; attempt++)
        {
            _ = await client.CallToolAsync("search", new { query = _settings.ProbeTerm, top = 5 });
            JsonElement cachedSearch = await client.CallToolAsync("search", new
            {
                query = _settings.ProbeTerm,
                store = _settings.TestHubStoreDisplayName,
                top = 5,
            });
            JsonElement cachedSweep = cachedSearch.GetProperty("sweep");
            Assert.True(cachedSweep.GetProperty("performed").GetBoolean());
            cachedProven = cachedSweep.TryGetProperty("cached", out JsonElement cachedFlag) && cachedFlag.GetBoolean();
            if (cachedProven)
            {
                _output.WriteLine($"search (cached sweep, attempt {attempt + 1}): "
                    + $"cacheAgeSeconds={cachedSweep.GetProperty("cacheAgeSeconds").GetDouble()}");
            }
        }

        Assert.True(cachedProven, "a rapid follow-up search must be served from the sweep cache (D34)");

        // --- read: full golden shape on the first fast hit.
        JsonElement read = await client.CallToolAsync("read", new { id = hitId, max_body_chars = 1500 });
        Assert.True(read.GetProperty("entryId").GetString()!.Length >= 48);
        Assert.True(read.TryGetProperty("body", out _));
        Assert.True(read.GetProperty("bodyTotalChars").GetInt64() >= 0);
        Assert.True(read.TryGetProperty("bodyTruncated", out _));
        Assert.True(read.TryGetProperty("recipients", out JsonElement recipients));
        Assert.Equal(JsonValueKind.Array, recipients.ValueKind);
        Assert.True(read.TryGetProperty("attachments", out JsonElement attachments));
        Assert.Equal(JsonValueKind.Array, attachments.ValueKind);
        _output.WriteLine($"read: locatedVia={read.GetProperty("locatedVia").GetString()} bodyChars={read.GetProperty("bodyTotalChars").GetInt64()}");

        // --- thread: via the hit's conversation id (or the COM fallback via id).
        object threadArgs = firstIndexHit.TryGetProperty("conversationId", out JsonElement conv)
            ? new { conversation_id = conv.GetString(), id = hitId, store = firstIndexHit.GetProperty("store").GetString() }
            : new { id = hitId };
        JsonElement thread = await client.CallToolAsync("thread", threadArgs);
        Assert.True(thread.GetProperty("hits").GetArrayLength() >= 1);
        string threadSource = thread.GetProperty("source").GetString()!;
        Assert.True(threadSource is "index" or "com");
        _output.WriteLine($"thread: source={threadSource} members={thread.GetProperty("hits").GetArrayLength()} ms={thread.GetProperty("elapsedMs").GetInt64()}");

        // --- save_attachment: find a mail with attachments, save the first one.
        JsonElement withAttachments = await client.CallToolAsync("search", new
        {
            has_attachments = true,
            include_attachment_hits = false,
            top = 8,
        });
        string? savedPath = null;
        foreach (JsonElement hit in withAttachments.GetProperty("hits").EnumerateArray())
        {
            JsonElement candidate = await client.CallToolAsync("read", new
            {
                id = hit.GetProperty("id").GetString(),
                max_body_chars = 0,
            });
            if (candidate.TryGetProperty("error", out _))
            {
                continue;
            }

            if (candidate.GetProperty("attachments").GetArrayLength() >= 1)
            {
                int index = candidate.GetProperty("attachments")[0].GetProperty("index").GetInt32();
                JsonElement saved = await client.CallToolAsync("save_attachment", new
                {
                    id = hit.GetProperty("id").GetString(),
                    attachment_index = index,
                });
                if (saved.TryGetProperty("error", out JsonElement saveError))
                {
                    _output.WriteLine($"save_attachment skipped candidate: {saveError.GetProperty("type").GetString()}");
                    continue;
                }

                savedPath = saved.GetProperty("savedPath").GetString();
                Assert.True(saved.GetProperty("sizeBytes").GetInt64() >= 0);
                break;
            }
        }

        Assert.NotNull(savedPath);
        Assert.True(File.Exists(savedPath), "saved attachment must exist on disk");
        _output.WriteLine("save_attachment: file exists=true");
        File.Delete(savedPath!);

        // --- clean shutdown (no leaked processes per agent session).
        Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
    }

    [Fact]
    public async Task Status_Accounts_Folders_GoldenShapes_OverRealStdio()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(6));

        // --- outlook_health (the merged diagnostics tool, D37): index freshness block
        // with the per-store frontier rows and advice the old index_status carried.
        JsonElement status = await client.CallToolAsync("outlook_health", new { });
        JsonElement index = status.GetProperty("index");
        Assert.True(index.GetProperty("provider").GetString() is "OleDb" or "AdodbCom");
        Assert.True(index.TryGetProperty("newestIndexedUtc", out _));
        Assert.True(index.GetProperty("perStore").GetArrayLength() >= 3);
        Assert.True(status.GetProperty("advice").GetArrayLength() >= 1);
        _output.WriteLine($"outlook_health: perStore={index.GetProperty("perStore").GetArrayLength()}");

        // --- list_accounts: 3 accounts, delegates distinct, flags present.
        JsonElement accounts = await client.CallToolAsync("list_accounts", new { });
        Assert.Equal(3, accounts.GetProperty("accounts").GetArrayLength());
        JsonElement stores = accounts.GetProperty("stores");
        Assert.True(stores.GetArrayLength() >= 3);
        int delegates = 0;
        foreach (JsonElement store in stores.EnumerateArray())
        {
            Assert.False(string.IsNullOrEmpty(store.GetProperty("displayName").GetString()));
            Assert.True(store.TryGetProperty("locallySearchable", out _));
            Assert.True(store.TryGetProperty("onlineOnly", out _));
            if (store.GetProperty("isDelegate").GetBoolean())
            {
                delegates++;
            }
        }

        _output.WriteLine($"list_accounts: stores={stores.GetArrayLength()} delegates={delegates}");
        Assert.Equal(_settings.ExpectedDelegateStoreDisplayNames.Count, delegates);

        // --- list_folders scoped to the test hub (full tree, one page).
        JsonElement folders = await client.CallToolAsync("list_folders", new
        {
            store = _settings.TestHubStoreDisplayName,
        });
        JsonElement storeFolders = folders.GetProperty("stores")[0].GetProperty("folders");
        Assert.True(storeFolders.GetArrayLength() >= 1);
        Assert.False(string.IsNullOrEmpty(storeFolders[0].GetProperty("path").GetString()));
        Assert.Equal(JsonValueKind.False, folders.GetProperty("truncated").ValueKind);
        Assert.True(folders.GetProperty("folderTotal").GetInt32() >= storeFolders.GetArrayLength());
        _output.WriteLine($"list_folders: folders={storeFolders.GetArrayLength()} total={folders.GetProperty("folderTotal").GetInt32()}");

        // Offset paging over the wire: offset=1 = the same stable order minus the first folder.
        JsonElement page1 = await client.CallToolAsync("list_folders", new
        {
            store = _settings.TestHubStoreDisplayName,
            offset = 1,
        });
        Assert.Equal(1, page1.GetProperty("offset").GetInt32());
        Assert.Equal(
            storeFolders[1].GetProperty("path").GetString(),
            page1.GetProperty("stores")[0].GetProperty("folders")[0].GetProperty("path").GetString());
        _output.WriteLine("list_folders offset paging: wire-consistent");

        Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
    }
}
