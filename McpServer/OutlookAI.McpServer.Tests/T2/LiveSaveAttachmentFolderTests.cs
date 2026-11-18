using System.Globalization;
using System.Text;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T1;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// save_attachment on a real Outlook writes only into OutlookAI's download folder (Q198 (a)): what
/// T1 cannot show, because it needs Outlook's own <c>Attachment.SaveAsFile</c> writing into the
/// file the writer claimed, Outlook's own <c>Attachment.FileName</c> for names a sender can give an
/// attachment and no file on disk can have, and the COM host's own check of a folder a junction
/// leads out of. The disk cases themselves are pinned in T1 (<c>AttachmentFileWriterTests</c>,
/// <c>AttachmentFileWriterDiskTests</c>).
/// <para>
/// SAFETY: the one item each test makes is a draft in the hub carrying the tag and this run's
/// marker (S2, S3), made and removed through <see cref="LiveOutlookTestMailer"/>; nothing is sent.
/// Every file and folder a test makes carries the marker, is made in the download folder or the
/// temporary folder, and is removed by its own path - the junction as a link. The Startup folder
/// is only read.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LiveSaveAttachmentFolderTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSaveAttachmentFolderTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private string Tag => "OutlookAI-McpTest-q198-" + Marker;

    private static string Startup =>
        Environment.GetFolderPath(Environment.SpecialFolder.Startup, Environment.SpecialFolderOption.DoNotVerify);

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void SaveAttachment_WritesOnlyInsideTheDownloadFolder_AndNeverReplacesAFile()
    {
        string sources = Path.Combine(Path.GetTempPath(), Tag + "-sources");
        string outside = Path.Combine(Path.GetTempPath(), Tag + "-outside");
        string ours = Path.Combine(AttachmentFileWriter.DownloadFolder, Tag);
        string junction = Path.Combine(AttachmentFileWriter.DownloadFolder, Tag + "-junction");
        string payload = "OutlookAI save_attachment folder probe " + Marker;
        string source = Path.Combine(sources, "q198-" + Marker + ".txt");
        List<string> saved = new();
        string? entryId = null;
        Directory.CreateDirectory(sources);
        Directory.CreateDirectory(outside);
        File.WriteAllText(source, payload);
        try
        {
            entryId = LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
                Hub, _fixture.TaggedSubject("q198 save folder"), "save_attachment folder probe " + Marker, new[] { source });
            string downloads = AttachmentFileWriter.CheckFolder(null);

            // No target_dir: the download folder itself, Outlook writing into the file the writer claimed.
            SaveAttachmentOutcome first = Service.SaveAttachment(entryId, 1);
            saved.Add(first.SavedPath);
            Assert.Equal(downloads, Path.GetDirectoryName(first.SavedPath), ignoreCase: true);
            Assert.Equal(payload, File.ReadAllText(first.SavedPath));
            Assert.Equal(new FileInfo(first.SavedPath).Length, first.SizeBytes);

            // The same attachment again: a new name, and the first file as it was.
            SaveAttachmentOutcome second = Service.SaveAttachment(entryId, 1);
            saved.Add(second.SavedPath);
            Assert.Equal(
                Path.GetFileNameWithoutExtension(first.FileName) + " (2)" + Path.GetExtension(first.FileName), second.FileName);
            Assert.Equal(payload, File.ReadAllText(first.SavedPath));
            Assert.Equal(payload, File.ReadAllText(second.SavedPath));

            // A folder inside, by a relative path and by its full path - made when missing.
            SaveAttachmentOutcome nested = Service.SaveAttachment(entryId, 1, Tag + "/nested");
            saved.Add(nested.SavedPath);
            Assert.Equal(Path.Combine(downloads, Tag, "nested"), Path.GetDirectoryName(nested.SavedPath), ignoreCase: true);
            SaveAttachmentOutcome full = Service.SaveAttachment(entryId, 1, Path.Combine(AttachmentFileWriter.DownloadFolder, Tag, "full"));
            saved.Add(full.SavedPath);
            Assert.Equal(Path.Combine(downloads, Tag, "full"), Path.GetDirectoryName(full.SavedPath), ignoreCase: true);
            Assert.Equal(payload, File.ReadAllText(full.SavedPath));
            _output.WriteLine("saved inside: " + string.Join(", ", saved.Select(Path.GetFileName)));

            // Anywhere else: refused before Outlook is asked, and nothing written.
            string escape = @"..\..\..\" + Tag + "-escape";
            foreach (string asked in new[]
            {
                outside, Startup, escape, @"\\?\" + outside, Tag + ":stream", @"\\localhost\C$\" + Tag,
            })
            {
                AttachmentFolderRefusedException refused = Assert.Throws<AttachmentFolderRefusedException>(
                    () => Service.SaveAttachment(entryId, 1, asked));
                Assert.Contains("copy it with your own file tools", refused.Message, StringComparison.Ordinal);
                _output.WriteLine("refused: " + FirstSentence(refused.Message));
            }

            Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
            Assert.False(File.Exists(Path.Combine(Startup, first.FileName)), "a file of this run reached the Startup folder");
            Assert.False(
                Directory.Exists(Path.Combine(downloads, "..", "..", "..", Tag + "-escape")),
                "a folder of this run was made outside the download folder");

            // A junction inside the download folder that leads out: refused by the server, and by the
            // COM session on its own, before it opens the item.
            TestJunction.Make(junction, outside);
            AttachmentFolderRefusedException throughJunction = Assert.Throws<AttachmentFolderRefusedException>(
                () => Service.SaveAttachment(entryId, 1, Tag + "-junction"));
            Assert.Contains("it is really", throughJunction.Message, StringComparison.Ordinal);

            string? written = _fixture.VerifySession.TrySaveAttachment(
                entryId, _fixture.GetStoreId(Hub), 1, junction, out long size, out string? attempted, out string? error);
            Assert.Null(written);
            Assert.Null(attempted);
            Assert.Equal(0, size);
            Assert.True(ComErrorTokens.TryRead(error, ComErrorTokens.SaveFolderRefused, out _), "the COM session answered: " + error);
            Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
            _output.WriteLine("junction refused by the server and by the COM session; nothing written where it leads");
        }
        finally
        {
            foreach (string path in saved)
            {
                TryDeleteFile(path);
            }

            TestJunction.Remove(junction);
            TryDeleteTree(ours);
            TryDeleteTree(outside);
            TryDeleteTree(sources);
            if (entryId != null)
            {
                CleanupDraft(entryId);
            }
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void SaveAttachment_GivesTheSendersFileNamesASafeName_InsideTheFolder()
    {
        // Names a sender can put on an attachment and no file on disk can have. What Outlook reports
        // for each is the measurement; whatever it is, the file must land in the folder asked for,
        // under that name made safe.
        string[] sent =
        {
            @"..\..\..\Start Menu\Programs\Startup\q198-" + Marker + "-1.txt",
            "CON.txt",
            "q198-" + Marker + "-3.txt. . .",
            "q198-" + Marker + "-4.txt:hidden",
            "q198/" + Marker + "/5.txt",
            "q198-" + Marker + "-6\u202Efdp.txt",
        };
        string sources = Path.Combine(Path.GetTempPath(), Tag + "-names");
        string ours = Path.Combine(AttachmentFileWriter.DownloadFolder, Tag);
        List<string> paths = new();
        List<string> contents = new();
        List<string> saved = new();
        string? entryId = null;
        Directory.CreateDirectory(sources);
        for (int i = 0; i < sent.Length; i++)
        {
            string path = Path.Combine(sources, "q198-" + Marker + "-source-" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".txt");
            string content = "OutlookAI save_attachment name probe " + (i + 1).ToString(CultureInfo.InvariantCulture) + " " + Marker;
            File.WriteAllText(path, content);
            paths.Add(path);
            contents.Add(content);
        }

        try
        {
            entryId = LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
                Hub, _fixture.TaggedSubject("q198 save names"), "save_attachment name probe " + Marker, paths, sent);
            string folder = AttachmentFileWriter.CheckFolder(Tag + @"\names");

            ReadOutcome read = Service.Read(entryId, maxBodyChars: 0);
            Assert.Equal(sent.Length, read.Attachments.Count);
            foreach (AttachmentView attachment in read.Attachments)
            {
                SaveAttachmentOutcome outcome = Service.SaveAttachment(entryId, attachment.Index, Tag + "/names");
                saved.Add(outcome.SavedPath);
                string reported = attachment.FileName ?? string.Empty;
                _output.WriteLine("#" + attachment.Index.ToString(CultureInfo.InvariantCulture) + ": sent '" + Visible(sent[attachment.Index - 1])
                    + "', Outlook reports '" + Visible(reported) + "', saved as '" + Visible(outcome.FileName) + "'");

                Assert.Equal(folder, Path.GetDirectoryName(outcome.SavedPath), ignoreCase: true);
                Assert.Equal(AttachmentFileWriter.SafeFileName(reported), outcome.FileName);
                Assert.Equal(contents[attachment.Index - 1], File.ReadAllText(outcome.SavedPath));
            }

            Assert.Equal(sent.Length, Directory.EnumerateFiles(folder).Count());
            Assert.DoesNotContain(
                Directory.Exists(Startup) ? Directory.EnumerateFileSystemEntries(Startup) : Enumerable.Empty<string>(),
                entry => entry.Contains(Marker, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            foreach (string path in saved)
            {
                TryDeleteFile(path);
            }

            TryDeleteTree(ours);
            TryDeleteTree(sources);
            if (entryId != null)
            {
                CleanupDraft(entryId);
            }
        }

        AssertNoTaggedArtifactsRemain();
    }

    // ------------------------------------------------------------------ support

    private void CleanupDraft(string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
                LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    private void AssertNoTaggedArtifactsRemain()
    {
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, folderIds: null, out int stragglers);
        if (stragglers > 0)
        {
            _output.WriteLine($"cleanup[{Hub}]: {stragglers} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static void TryDeleteTree(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static string FirstSentence(string message)
    {
        int end = message.IndexOf(". ", StringComparison.Ordinal);
        return end < 0 ? message : message.Substring(0, end + 1);
    }

    /// <summary>A name with every control and direction character written as its code, so the log shows what is there.</summary>
    private static string Visible(string text)
    {
        StringBuilder shown = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            bool hidden = c < ' ' || c == '\u007F' || c == '\u061C' || c == '\u200E' || c == '\u200F'
                || (c >= '\u202A' && c <= '\u202E') || (c >= '\u2066' && c <= '\u2069');
            shown.Append(hidden ? "<U+" + ((int)c).ToString("X4", CultureInfo.InvariantCulture) + ">" : c.ToString());
        }

        return shown.ToString();
    }
}
