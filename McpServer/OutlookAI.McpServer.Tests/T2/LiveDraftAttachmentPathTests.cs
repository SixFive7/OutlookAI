using System.Diagnostics;
using System.Globalization;

using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T1;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The draft tools refuse the attachment paths Outlook crashed on or would not attach on the test
/// guests, before Outlook is asked anything, and still attach ordinary files. Measured 2026-11-18
/// (runbook section 4.10): a <c>\\?\</c> path crashed OUTLOOK.EXE inside <c>Attachments.Add</c> four
/// times out of four, whatever the file was called; Outlook refused a file called <c>CON.txt</c>, an
/// alternate data stream and a 304-character path; it attached the right file from <c>\\.\C:\...</c>
/// and <c>\??\C:\...</c>. The spellings themselves are T1's (<c>DraftAttachmentsTests</c>); this is the
/// real Outlook: a refused call leaves no draft and the same OUTLOOK.EXE running, and a path at the
/// longest length allowed attaches.
/// <para>
/// SAFETY: the one draft is made in the hub with new_draft, addressed to the hub itself (Q110), tagged
/// with this run's marker (S2, S3), revised with update_draft and removed through
/// <see cref="LiveDraftCleanup"/>; nothing is sent. Every file carries the marker, lies in the temporary
/// folder and is removed by its own path. No path that crashed Outlook reaches it: that is the claim.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LiveDraftAttachmentPathTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDraftAttachmentPathTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void DraftTools_RefuseThePathsOutlookCannotTake_BeforeOutlookIsAsked_AndStillAttachOrdinaryFiles()
    {
        string folder = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-p2-" + Marker);
        string con = Path.Combine(folder, "CON.txt");
        string trail = Path.Combine(folder, "p2-trail.txt. . .");
        string refusedLabel = "p2 refused " + Guid.NewGuid().ToString("N").Substring(0, 8);
        DraftOutcome? draft = null;
        Directory.CreateDirectory(folder);
        try
        {
            string plain = Write(Path.Combine(folder, "p2-plain.txt"), "p2 plain file " + Marker);
            string dotted = Write(Path.Combine(folder, "p2-dot-path.txt"), "p2 file attached through \\\\.\\ " + Marker);
            string nt = Write(Path.Combine(folder, "p2-nt-path.txt"), "p2 file attached through \\??\\ " + Marker);
            string atTheLimit = WriteWithPathLength(folder, "p2-limit", WindowsFileNames.MaxPathChars);
            string over = WriteWithPathLength(folder, "p2-over", WindowsFileNames.MaxPathChars + 1);
            string stream = plain + ":p2stream";
            File.WriteAllText(stream, "p2 alternate data stream " + Marker);
            File.WriteAllText(@"\\?\" + con, "p2 file called CON.txt " + Marker);
            File.WriteAllText(@"\\?\" + trail, "p2 file whose name ends in dots and spaces " + Marker);

            _ = Service.ListAccounts();
            int[] outlookBefore = OutlookProcessIds();
            Assert.NotEmpty(outlookBefore);

            // Each beside an ordinary file, so the whole set is refused - and nothing is made.
            string[] refused =
            {
                @"\\?\" + plain, @"\\?\" + con, @"\\?\" + trail, @"\\?\" + over,
                con, stream, over, @"\??\" + trail, @"\\.\pipe\outlookai-p2",
            };
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
            foreach (string path in refused)
            {
                ArgumentException ex = Assert.Throws<ArgumentException>(() => Service.NewDraft(
                    Hub, Hub, cc: null, _fixture.TaggedSubject(refusedLabel), "p2 refused attachment " + Marker,
                    display: false, attachments: new[] { plain, path }));
                Assert.Contains("1 unusable entry: '" + path + "': ", ex.Message, StringComparison.Ordinal);
                Assert.Contains("no draft was changed", ex.Message, StringComparison.Ordinal);
                _output.WriteLine("new_draft refused " + Shape(path, folder) + ": " + Reason(ex.Message, path));
            }

            Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, refusedLabel));
            AssertStillRunning(outlookBefore);

            // Ordinary files still attach: by their ordinary path, through \\.\ and \??\, and at the longest path allowed.
            string[] attached = { plain, @"\\.\" + dotted, @"\??\" + nt, atTheLimit };
            draft = Service.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("p2 attached"), "p2 attachment paths " + Marker,
                display: false, attachments: attached);

            Assert.NotNull(draft.Attachments);
            Assert.Equal(attached.Length, draft.Attachments!.Count);
            foreach (string file in new[] { plain, dotted, nt, atTheLimit })
            {
                AttachmentView view = Assert.Single(draft.Attachments, a => a.FileName == Path.GetFileName(file));
                Assert.Equal(new FileInfo(file).Length, view.SizeBytes);
            }

            _output.WriteLine("new_draft attached " + attached.Length.ToString(CultureInfo.InvariantCulture)
                + " files: ordinary path, \\\\.\\, \\??\\, and a path of " + atTheLimit.Length.ToString(CultureInfo.InvariantCulture) + " characters");

            // update_draft holds the same line, and the draft keeps what it had.
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            ArgumentException update = Assert.Throws<ArgumentException>(
                () => Service.UpdateDraft(draft.EntryId, attachments: new[] { @"\\?\" + plain }, display: false));
            Assert.Contains("Outlook crashes", update.Message, StringComparison.Ordinal);
            Assert.Equal(attached.Length, Service.Read(draft.EntryId, maxBodyChars: 0).Attachments.Count);
            _output.WriteLine("update_draft refused a \\\\?\\ path; the draft still holds its " + attached.Length.ToString(CultureInfo.InvariantCulture) + " files");

            AssertStillRunning(outlookBefore);
        }
        finally
        {
            if (draft != null)
            {
                LiveDraftCleanup.DeleteDraft(Hub, Marker, draft.EntryId);
            }

            ScratchFiles.TryDeleteFile(@"\\?\" + con);
            ScratchFiles.TryDeleteFile(@"\\?\" + trail);
            ScratchFiles.TryDeleteTree(folder);
        }

        LiveDraftCleanup.AssertNoTaggedArtifactsRemain(Hub, Marker, _output);
    }

    private static string Write(string path, string content)
    {
        File.WriteAllText(path, content);
        return path;
    }

    /// <summary>A file in <paramref name="folder"/> whose full path is exactly <paramref name="length"/> characters.</summary>
    private string WriteWithPathLength(string folder, string label, int length)
    {
        string name = label + "-" + Marker + ".txt";
        int filler = length - folder.Length - 1 - name.Length;
        Assert.InRange(filler, 0, 200);
        string path = Path.Combine(folder, label + "-" + Marker + new string('n', filler) + ".txt");
        Assert.Equal(length, path.Length);
        return Write(path, "p2 file whose path is " + length.ToString(CultureInfo.InvariantCulture) + " characters " + Marker);
    }

    /// <summary>Every OUTLOOK.EXE that ran before is running still: nothing handed to Outlook ended it.</summary>
    private static void AssertStillRunning(int[] before)
    {
        int[] now = OutlookProcessIds();
        Assert.All(before, id => Assert.Contains(id, now));
    }

    private static int[] OutlookProcessIds()
    {
        Process[] processes = Process.GetProcessesByName("OUTLOOK");
        try
        {
            return processes.Select(p => p.Id).OrderBy(id => id).ToArray();
        }
        finally
        {
            foreach (Process process in processes)
            {
                process.Dispose();
            }
        }
    }

    /// <summary>The path's shape for the log: the test folder written as such.</summary>
    private static string Shape(string path, string folder)
    {
        return path.Replace(folder, "<folder>", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The reason given for <paramref name="path"/>, up to its advice.</summary>
    private static string Reason(string message, string path)
    {
        string head = "'" + path + "': ";
        int at = message.IndexOf(head, StringComparison.Ordinal);
        string reason = at < 0 ? message : message.Substring(at + head.Length);
        int advice = reason.IndexOf(" - ", StringComparison.Ordinal);
        return advice < 0 ? reason : reason.Substring(0, advice);
    }
}
