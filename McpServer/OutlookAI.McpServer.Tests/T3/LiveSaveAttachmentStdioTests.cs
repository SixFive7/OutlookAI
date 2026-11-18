using System.Text.Json;

using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T1;
using OutlookAI.McpServer.Tests.T2;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// save_attachment through the shipped server over real stdio (Q198 (a)): the path a user's agent
/// takes, with the server checking the folder in its own process and the COM host checking it again
/// and writing the file in another. On a guest with or without a search index - the golden-shape
/// test that also saves an attachment needs the index - so the production path is proven on both.
/// <para>
/// SAFETY: the one item is a draft in the hub carrying the tag and this run's marker (S2, S3), made
/// and removed through <see cref="LiveOutlookTestMailer"/>; every file and folder carries the marker
/// and is removed by its own path.
/// </para>
/// </summary>
[Trait("Category", "Live")]
[Collection(LiveCollections.McpToolShape)]
public sealed class LiveSaveAttachmentStdioTests
{
    private readonly LiveMcpToolShapeFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSaveAttachmentStdioTests(LiveMcpToolShapeFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public async Task SaveAttachment_OverStdio_WritesInsideTheDownloadFolder_AndRefusesAnywhereElse()
    {
        string hub = _fixture.Settings.TestHubStoreDisplayName;
        string marker = LiveRunMarker.ForFixture("q198s");
        string tag = "OutlookAI-McpTest-q198s-" + marker;
        string sources = Path.Combine(Path.GetTempPath(), tag);
        string outside = Path.Combine(Path.GetTempPath(), tag + "-outside");
        string payload = "OutlookAI save_attachment stdio probe " + marker;
        string source = Path.Combine(sources, "q198s-" + marker + ".txt");
        string? entryId = null;
        string? savedPath = null;
        Directory.CreateDirectory(sources);
        Directory.CreateDirectory(outside);
        File.WriteAllText(source, payload);
        try
        {
            entryId = LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
                hub, LiveOutlookTestMailer.SubjectTag + " " + marker + " q198 save over stdio",
                "save_attachment stdio probe " + marker, new[] { source });

            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(6));

            // Inside: a folder of the download folder, by a relative path - made, and the file in it.
            JsonElement saved = await client.CallToolAsync("save_attachment", new { id = entryId, attachment_index = 1, target_dir = tag });
            Assert.False(saved.TryGetProperty("error", out JsonElement unexpected), "save_attachment failed: " + unexpected);
            savedPath = saved.GetProperty("savedPath").GetString();
            Assert.Equal(
                Path.Combine(AttachmentFileWriter.CheckFolder(null), tag), Path.GetDirectoryName(savedPath), ignoreCase: true);
            Assert.Equal(payload, File.ReadAllText(savedPath!));
            _output.WriteLine("saved over stdio: " + Path.GetFileName(savedPath));

            // Anywhere else: refused with the folder rule, provably unchanged, and nothing written.
            (JsonElement refusedPayload, bool isError) = await client.CallToolWithIsErrorAsync(
                "save_attachment", new { id = entryId, attachment_index = 1, target_dir = outside });
            Assert.True(isError, "a refused folder must come back as a tool error");
            JsonElement error = refusedPayload.GetProperty("error");
            Assert.Equal("SaveFolderRefused", error.GetProperty("type").GetString());
            Assert.Equal("unchanged", error.GetProperty("outcome").GetString());
            Assert.Contains("copy it with your own file tools", error.GetProperty("message").GetString(), StringComparison.Ordinal);
            Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
            _output.WriteLine("refused over stdio: " + error.GetProperty("type").GetString() + ", outcome " + error.GetProperty("outcome").GetString());

            Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
        }
        finally
        {
            if (savedPath != null)
            {
                ScratchFiles.TryDeleteFile(savedPath);
            }

            ScratchFiles.TryDeleteTree(Path.Combine(AttachmentFileWriter.DownloadFolder, tag));
            ScratchFiles.TryDeleteTree(outside);
            ScratchFiles.TryDeleteTree(sources);
            if (entryId != null)
            {
                LiveDraftCleanup.DeleteDraft(hub, marker, entryId);
            }
        }

        LiveDraftCleanup.AssertNoTaggedArtifactsRemain(hub, marker, _output);
    }
}
