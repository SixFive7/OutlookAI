using System.Globalization;
using System.Text.Json;

using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// D34 wire-schema pin (user decision 2026-07-24: "drop the fast mode. Keep everything
/// inside of 1 tool."): the search tool advertises NO mode enum - fresh (index +
/// freshness sweep) is THE behavior - and exhaustive survives as a boolean flag on the
/// same tool. CI-safe: tools/list needs no Outlook and no index.
/// </summary>
public sealed class SearchSchemaCiTests
{
    [Fact]
    public async Task SearchInputSchema_HasNoModeParameter_AndHasExhaustiveBoolean()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? searchTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "search")
            {
                searchTool = tool;
                break;
            }
        }

        Assert.True(searchTool != null, "the search tool must be advertised");
        JsonElement properties = searchTool!.Value.GetProperty("inputSchema").GetProperty("properties");

        Assert.False(properties.TryGetProperty("mode", out _),
            "the search schema must not expose a 'mode' parameter (removed by D34)");
        Assert.True(properties.TryGetProperty("exhaustive", out JsonElement exhaustive),
            "the search schema must expose the 'exhaustive' boolean (D34)");
        Assert.Equal("boolean", exhaustive.GetProperty("type").GetString());

        // The description must document the always-fresh contract + graceful degradation.
        string description = searchTool.Value.GetProperty("description").GetString()!;
        Assert.Contains("fresh", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("advice", description, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("mode=", description, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// D40 wire pin (user order 2026-07-26, SF-6 fix): the search_in argument exists,
    /// is a string, and the tool description states plainly what 'query' matches by
    /// default - subject AND body - instead of the old "search all ... mail" overpromise
    /// that hid a body-content-only predicate.
    /// </summary>
    [Fact]
    public async Task SearchSchema_ExposesSearchIn_AndDescribesWhatQueryMatches()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? searchTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "search")
            {
                searchTool = tool;
                break;
            }
        }

        Assert.True(searchTool != null, "the search tool must be advertised");
        JsonElement properties = searchTool!.Value.GetProperty("inputSchema").GetProperty("properties");

        Assert.True(properties.TryGetProperty("search_in", out JsonElement searchIn),
            "the search schema must expose the 'search_in' parameter (D40, renamed 2026-07-26)");
        Assert.Contains("string", DescribeJsonType(searchIn.GetProperty("type")), StringComparison.Ordinal);

        // The pre-rename name must be gone from the wire entirely.
        Assert.False(properties.TryGetProperty("term_scope", out _),
            "the search schema must not expose the old 'term_scope' name (renamed to 'search_in')");

        // The parameter description must name all three values and explain the default.
        string paramDescription = searchIn.GetProperty("description").GetString()!;
        foreach (string wireName in new[] { "subject_and_body", "subject", "body" })
        {
            Assert.Contains(wireName, paramDescription, StringComparison.Ordinal);
        }

        Assert.Contains("default", paramDescription, StringComparison.OrdinalIgnoreCase);

        // search_in must be optional - omitting it is the default subject+body search.
        if (searchTool.Value.GetProperty("inputSchema").TryGetProperty("required", out JsonElement required))
        {
            foreach (JsonElement name in required.EnumerateArray())
            {
                Assert.NotEqual("search_in", name.GetString());
            }
        }

        // The tool description must say what 'query' matches, and must not repeat the
        // pre-D40 overpromise ("Search all locally indexed Outlook mail ...").
        string description = searchTool.Value.GetProperty("description").GetString()!;
        Assert.Contains("subject", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("body", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("search_in", description, StringComparison.Ordinal);
        Assert.DoesNotContain("term_scope", description, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Search all locally indexed Outlook mail", description, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Description-rewrite pin (user order 2026-07-26): the description is a USAGE
    /// contract - every nuance that changes behavior must be stated. Each assertion
    /// below is a claim verified against the shipped code; changing the behavior
    /// without changing the sentence (or vice versa) must fail here.
    /// <para>
    /// Re-homed 2026-08-17 (client-truncation fix, see DescriptionBudgetCiTests): the tool
    /// description had grown to 3912 characters, and Claude Code truncates tool
    /// descriptions at 2 KB - positionally and silently - so from "only mail f|rom roughly
    /// the last few minutes" onward NOTHING below reached the model. Every claim pinned
    /// here still has to be on the wire; several are now pinned on the ARGUMENT that owns
    /// them (the schema's own, separately budgeted surface) instead of on the tool
    /// description, and the ones that only describe the ANSWER are pinned against the
    /// runtime payload that already reports them.
    /// </para>
    /// </summary>
    [Fact]
    public async Task SearchDescription_StatesEveryBehaviorChangingNuance()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? searchTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "search")
            {
                searchTool = tool;
                break;
            }
        }

        Assert.True(searchTool != null, "the search tool must be advertised");
        string description = searchTool!.Value.GetProperty("description").GetString()!;
        JsonElement parameters = searchTool.Value.GetProperty("inputSchema").GetProperty("properties");

        // Matching contract: whole words, terms may land in DIFFERENT parts (soak fix
        // 13 - the builder ANDs across the columns, one pair per term), sender/recipient
        // via from/to, prefix star, allowed charset. This is what a caller needs BEFORE
        // the call, so it stays on the tool description.
        Assert.Contains("whole words", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("different parts", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("include_attachment_hits", description, StringComparison.Ordinal);
        Assert.Contains("from / to", description, StringComparison.Ordinal);
        Assert.Contains("@.-_'+", description, StringComparison.Ordinal);

        // The retired claim must not come back: terms no longer have to share one part.
        Assert.DoesNotContain("same part", description, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("kind=document", description, StringComparison.OrdinalIgnoreCase);

        // Freshness contract (D34 + soak fix 13): the always-on sweep, its headless
        // autostart, and the sweep block that reports its coverage. The COVERAGE RULES
        // themselves are pinned on the folder argument below - they are how-to-call
        // detail, and the sweep block reports the realised coverage per call anyway.
        Assert.Contains("mail that arrived after the last index update", description, StringComparison.Ordinal);
        Assert.Contains("sweep block", description, StringComparison.Ordinal);
        Assert.Contains("headless", description, StringComparison.OrdinalIgnoreCase);

        // The retired freshness claim (sweep limited to Inbox + Sent Items) is gone.
        Assert.DoesNotContain("Inbox or Sent Items", description, StringComparison.Ordinal);

        // DEGRADED RESULTS is the one paragraph here that is a shipped BEHAVIOR the agent
        // must act on rather than documentation: degraded=true is a SUCCESSFUL result and
        // the user has to be told. It was the paragraph the 2 KB cut landed inside, so it
        // is now pinned word for word rather than by a single keyword.
        Assert.Contains("degraded=true", description, StringComparison.Ordinal);
        Assert.Contains("freshness=\"index-only\"", description, StringComparison.Ordinal);
        Assert.Contains("This is a SUCCESSFUL result, not an error", description, StringComparison.Ordinal);
        Assert.Contains("SAY SO TO THE USER when degraded is true", description, StringComparison.Ordinal);
        Assert.Contains("outlook_health", description, StringComparison.Ordinal);

        // ... and it must survive the cut, not merely exist. A client that truncates at
        // 2 KB must still receive the whole instruction.
        int degradedEnd = description.IndexOf("outlook_health gives the full picture", StringComparison.Ordinal);
        Assert.True(degradedEnd >= 0, "the degraded-results instruction must be intact");
        Assert.True(
            degradedEnd < DescriptionBudgetCiTests.ClientTruncationBudget,
            "the degraded-results instruction must land inside the client's truncation budget "
            + $"({DescriptionBudgetCiTests.ClientTruncationBudget}); it ends at {degradedEnd}");

        // D47: the freshness tier's REACH belongs to the flag that governs attachment
        // hits. The sweep reads Subject/Body through COM and never opens an attachment,
        // so attachment-content matching is index-only - the fact that makes an
        // attachment-only search index-only by construction.
        string attachmentFlagDescription =
            parameters.GetProperty("include_attachment_hits").GetProperty("description").GetString()!;
        Assert.Contains(
            "The sweep matches subject and body only - attachment text is matched by the index alone",
            attachmentFlagDescription,
            StringComparison.Ordinal);
        Assert.Contains("only once that mail is indexed", attachmentFlagDescription, StringComparison.Ordinal);

        // Soak fix 16: attachment matching covers EVERY attachment type. The old wording
        // implied documents only, which is exactly what the Kind='document' filter did -
        // 22.6% of attachment rows were unmatchable. Attachment hits are separate rows
        // flagged with isAttachmentHit=true.
        Assert.Contains("EVERY attachment type", attachmentFlagDescription, StringComparison.Ordinal);
        Assert.Contains("images, embedded messages, calendar invites", attachmentFlagDescription, StringComparison.Ordinal);
        Assert.Contains("isAttachmentHit=true", attachmentFlagDescription, StringComparison.Ordinal);

        // D47, the mirror image: excluding attachment hits drops ONLY those rows. The
        // sweep keeps running - it produces message rows, which is exactly what such a
        // caller asked for - so freshness coverage is not silently narrowed by the flag.
        Assert.Contains("drops only those hits", attachmentFlagDescription, StringComparison.Ordinal);
        Assert.Contains("freshness sweep's, are unaffected", attachmentFlagDescription, StringComparison.Ordinal);

        // Results contract: the id is the currency of every follow-up tool (D39 added
        // move_mail/archive_mail), and advice is for relaying.
        foreach (string followUpTool in new[] { "read", "thread", "save_attachment", "open_in_outlook", "move_mail", "archive_mail" })
        {
            Assert.Contains(followUpTool, description, StringComparison.Ordinal);
        }

        Assert.Contains("advice", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("scope", description, StringComparison.OrdinalIgnoreCase);

        // Truncation is DEFINITE (over-fetch by one) and belongs to the argument that
        // causes it; the runtime advice repeats it whenever it actually happens
        // ("Result list capped at N (top); more matches exist ...").
        string topDescription = parameters.GetProperty("top").GetProperty("description").GetString()!;
        Assert.Contains("truncated=true", topDescription, StringComparison.Ordinal);
        Assert.Contains("1-100", topDescription, StringComparison.Ordinal);

        // Folder-scope contract (soak fix 15) on the folder argument: ONE rule for every
        // mode - folder includes its subfolders unless include_subfolders=false - plus
        // the delegate caveats (name matching, widening, same-name collisions), the scope
        // block, and the sweep coverage that the folder bound also decides. Every one of
        // these is re-reported at runtime by MailService (FolderScopeResolver's widening /
        // collision advice, the scope block, and SweepInfo.Scope), so the schema states
        // them once, where the argument is chosen.
        string folderDescription = parameters.GetProperty("folder").GetProperty("description").GetString()!;
        Assert.Contains("include_subfolders=false", folderDescription, StringComparison.Ordinal);
        Assert.Contains("Delegate", folderDescription, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("WITHOUT", folderDescription, StringComparison.Ordinal);
        Assert.Contains("folder NAME", folderDescription, StringComparison.Ordinal);
        Assert.Contains("widens", folderDescription, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("same name", folderDescription, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("scope block", folderDescription, StringComparison.Ordinal);
        Assert.Contains("that folder and its subfolders", folderDescription, StringComparison.Ordinal);
        Assert.Contains("Inbox, Sent Items, Deleted Items and Junk Email", folderDescription, StringComparison.Ordinal);

        // The default four are swept NON-recursively (SweepFolder, not SweepFolderTree).
        Assert.Contains("those four folders only, not their subfolders", folderDescription, StringComparison.Ordinal);

        // Exhaustive contract on the exhaustive argument: bounds + the attachment-text
        // limitation. Soak fix 15 REMOVED the asymmetry the previous wording documented -
        // exhaustive now follows include_subfolders like every other mode - so the old
        // clauses must be gone and the cost warning must be present instead.
        string exhaustiveDescription = parameters.GetProperty("exhaustive").GetProperty("description").GetString()!;
        Assert.Contains("exhaustive=true", exhaustiveDescription, StringComparison.Ordinal);
        Assert.Contains("no attachment text", exhaustiveDescription, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("follows include_subfolders", exhaustiveDescription, StringComparison.Ordinal);

        // DERIVED from the constant, not a copy of the prose. This assertion used to read
        // Assert.Contains("120 s", ...) and pinned only the sentence: change the budget and
        // the test stayed green while the description lied to the agent about it. The
        // constant is public for exactly this reason.
        string budgetPhrase = (MailService.ExhaustiveTimeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture) + " s";
        Assert.Contains(budgetPhrase, exhaustiveDescription, StringComparison.Ordinal);
        Assert.Contains("foldersScanned/foldersSkipped", exhaustiveDescription, StringComparison.Ordinal);

        // The retired asymmetry claims (soak fix 14's wording) must not survive.
        Assert.DoesNotContain("ONLY the named folder - no subfolders", exhaustiveDescription, StringComparison.Ordinal);
        Assert.DoesNotContain("once per subfolder", exhaustiveDescription, StringComparison.Ordinal);
    }

    /// <summary>
    /// Soak-fix-15 pin: the folder ARGUMENT states ONE recursion rule for every mode
    /// (subfolders included unless include_subfolders=false), and the flag itself is on
    /// the wire as an optional boolean defaulting to true.
    /// </summary>
    [Fact]
    public async Task SearchFolderParameter_StatesOneSubfolderRule_ForEveryMode()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? searchTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "search")
            {
                searchTool = tool;
                break;
            }
        }

        Assert.True(searchTool != null, "the search tool must be advertised");

        JsonElement schema = searchTool!.Value.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        string folderDescription = properties.GetProperty("folder").GetProperty("description").GetString()!;
        Assert.Contains("subfolders", folderDescription, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("include_subfolders=false", folderDescription, StringComparison.Ordinal);

        // The soak-fix-14 exception clause is retired: exhaustive is no longer special.
        Assert.DoesNotContain("except with exhaustive", folderDescription, StringComparison.OrdinalIgnoreCase);

        Assert.True(
            properties.TryGetProperty("include_subfolders", out JsonElement includeSubfolders),
            "include_subfolders must be advertised on the search tool");
        Assert.Contains("boolean", DescribeJsonType(includeSubfolders.GetProperty("type")), StringComparison.Ordinal);

        string flagDescription = includeSubfolders.GetProperty("description").GetString()!;
        Assert.Contains("Default true", flagDescription, StringComparison.Ordinal);
        Assert.Contains("every mode", flagDescription, StringComparison.OrdinalIgnoreCase);

        // Optional argument: an existing caller that never passes it keeps working.
        if (schema.TryGetProperty("required", out JsonElement required))
        {
            foreach (JsonElement entry in required.EnumerateArray())
            {
                Assert.NotEqual("include_subfolders", entry.GetString());
            }
        }
    }

    private static string DescribeJsonType(JsonElement type)
    {
        if (type.ValueKind == JsonValueKind.Array)
        {
            return string.Join(",", type.EnumerateArray().Select(t => t.GetString()));
        }

        return type.GetString() ?? string.Empty;
    }
}
