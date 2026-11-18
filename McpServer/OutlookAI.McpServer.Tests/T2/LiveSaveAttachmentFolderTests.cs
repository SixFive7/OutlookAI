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
                ScratchFiles.TryDeleteFile(path);
            }

            TestJunction.Remove(junction);
            ScratchFiles.TryDeleteTree(ours);
            ScratchFiles.TryDeleteTree(outside);
            ScratchFiles.TryDeleteTree(sources);
            if (entryId != null)
            {
                LiveDraftCleanup.DeleteDraft(Hub, Marker, entryId);
            }
        }

        LiveDraftCleanup.AssertNoTaggedArtifactsRemain(Hub, Marker, _output);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void SaveAttachment_GivesTheSendersFileNamesASafeName_InsideTheFolder()
    {
        // A name reaches an attachment through the file it is added from: Outlook will not rename an
        // attachment through PropertyAccessor - PR_ATTACH_LONG_FILENAME "does not support this
        // operation", measured on the unindexed guest 2026-11-18. So the names here are names any
        // program can give a file: a direction control, a leading space, and a name too long for the
        // folder. A device's name and trailing dots and spaces only a \\?\ path can give a file, and
        // asking Outlook to attach CON.txt from one crashed OUTLOOK.EXE (an access violation, the same
        // day, same guest) - so those, path separators and ':' are T1's alone
        // (AttachmentFileWriterDiskTests), with every other name.
        string sources = Path.Combine(Path.GetTempPath(), Tag + "-names");

        // As long as Outlook can still read the file it is added from - its path under 250
        // characters - and longer than the download folder leaves room for, so the save shortens it.
        string longPrefix = "q198-" + Marker + "-3-";
        int filler = Math.Max(1, 250 - (sources.Length + 1) - longPrefix.Length - ".txt".Length);
        string[] plainNames =
        {
            "q198-" + Marker + "-1\u202Efdp.txt",
            " q198-" + Marker + "-2 leading space.txt",
            longPrefix + new string('n', filler) + ".txt",
        };
        string ours = Path.Combine(AttachmentFileWriter.DownloadFolder, Tag);
        List<string> saved = new();
        List<string> drafts = new();
        Directory.CreateDirectory(sources);
        try
        {
            string folder = AttachmentFileWriter.CheckFolder(Tag + @"\names");

            IReadOnlyList<string> contents = WriteSources(sources, plainNames, out IReadOnlyList<string> paths);
            drafts.Add(LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
                Hub, _fixture.TaggedSubject("q198 save names"), "save_attachment name probe " + Marker, paths));
            SaveEachAndCheck(drafts[0], plainNames, contents, folder, saved);

            Assert.Equal(saved.Count, Directory.EnumerateFiles(folder).Count());
            Assert.DoesNotContain(
                Directory.Exists(Startup) ? Directory.EnumerateFileSystemEntries(Startup) : Enumerable.Empty<string>(),
                entry => entry.Contains(Marker, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            foreach (string path in saved)
            {
                ScratchFiles.TryDeleteFile(path);
            }

            ScratchFiles.TryDeleteTree(ours);
            ScratchFiles.TryDeleteTree(sources);
            foreach (string draft in drafts)
            {
                LiveDraftCleanup.DeleteDraft(Hub, Marker, draft);
            }
        }

        LiveDraftCleanup.AssertNoTaggedArtifactsRemain(Hub, Marker, _output);
    }

    /// <summary>
    /// Writes one source file per name, each with content of its own, and returns the contents; the
    /// paths come back in <paramref name="paths"/>.
    /// </summary>
    private IReadOnlyList<string> WriteSources(string folder, IReadOnlyList<string> names, out IReadOnlyList<string> paths)
    {
        List<string> contents = new();
        List<string> written = new();
        for (int i = 0; i < names.Count; i++)
        {
            string path = Path.Combine(folder, names[i]);
            string content = "OutlookAI save_attachment name probe " + (i + 1).ToString(CultureInfo.InvariantCulture) + " " + Marker;
            File.WriteAllText(path, content);
            written.Add(path);
            contents.Add(content);
        }

        paths = written;
        return contents;
    }

    /// <summary>
    /// Saves every attachment of <paramref name="entryId"/> into the test's folder and checks each one:
    /// in the folder asked for, under the name Outlook reports made safe - shortened when the path
    /// needs it, its extension kept - and holding what the file held.
    /// </summary>
    private void SaveEachAndCheck(
        string entryId, IReadOnlyList<string> fileNames, IReadOnlyList<string> contents, string folder, List<string> saved)
    {
        ReadOutcome read = Service.Read(entryId, maxBodyChars: 0);
        Assert.Equal(fileNames.Count, read.Attachments.Count);
        foreach (AttachmentView attachment in read.Attachments)
        {
            int i = attachment.Index - 1;
            SaveAttachmentOutcome outcome = Service.SaveAttachment(entryId, attachment.Index, Tag + "/names");
            saved.Add(outcome.SavedPath);
            string reported = attachment.FileName ?? string.Empty;
            string safe = AttachmentFileWriter.SafeFileName(reported);
            _output.WriteLine("#" + attachment.Index.ToString(CultureInfo.InvariantCulture) + ": file '" + Visible(fileNames[i])
                + "', Outlook reports '" + Visible(reported) + "', saved as '" + Visible(outcome.FileName) + "'");

            Assert.Equal(folder, Path.GetDirectoryName(outcome.SavedPath), ignoreCase: true);
            Assert.True(outcome.SavedPath.Length <= AttachmentFileWriter.MaxPathChars, outcome.SavedPath.Length + " characters");
            Assert.Equal(Path.GetExtension(safe), Path.GetExtension(outcome.FileName));
            Assert.StartsWith(Path.GetFileNameWithoutExtension(outcome.FileName), Path.GetFileNameWithoutExtension(safe), StringComparison.Ordinal);
            Assert.False(WindowsFileNames.IsReservedDeviceName(outcome.FileName), outcome.FileName);
            Assert.DoesNotContain(outcome.FileName, c => c == '\u202E');
            Assert.Equal(contents[i], File.ReadAllText(outcome.SavedPath));
        }
    }

    // ------------------------------------------------------------------ support

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
