using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T1;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// MEASUREMENT PROBE, not a test the suite keeps (stream P2, 2026-11-18): which attachment paths
/// crash Outlook's <c>Attachments.Add</c>, which it refuses, and which it attaches - and what
/// <see cref="DraftAttachments.Validate"/> said about each before Outlook was asked. In live run
/// 20261118-050200 a draft whose attachments came from <c>\\?\</c> paths - <c>CON.txt</c> and a name
/// ending in dots and spaces - crashed OUTLOOK.EXE with an access violation inside the call, and the
/// draft tools hand Outlook any path the validation accepts.
/// <para>
/// Each case is its own draft with one attachment, so one case's crash or refusal hides no other's.
/// Outlook's process is watched across each call: a call that fails with an RPC error while the
/// OUTLOOK.EXE that took it exits is a crash, with its exit code; the next case starts a new Outlook
/// through the fixture's service. A draft that Outlook made is read back, its attachment saved through
/// save_attachment and compared with the file it came from, then deleted. The cases most likely to
/// crash come last, and a case Outlook hangs on ends the probe: the rest are reported NOT MEASURED.
/// </para>
/// <para>
/// SAFETY: every draft is made and removed through <see cref="LiveOutlookTestMailer"/> in the hub,
/// tagged with this run's marker (S2, S3); nothing is sent. Every file and folder carries the marker,
/// lies in the temporary folder or the download folder, and is removed by its own path. Outlook is
/// never ended by this probe; a crash is the measurement, and the run's crash dumps record it.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LiveDraftAttachmentPathProbeTests
{
    private const int RpcCallFailed = unchecked((int)0x800706BE);
    private const int RpcServerUnavailable = unchecked((int)0x800706BA);
    private const int RpcDisconnected = unchecked((int)0x80010108);
    private const int RpcServerDied = unchecked((int)0x80010007);
    private const int RpcServerDiedDne = unchecked((int)0x80010012);

    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDraftAttachmentPathProbeTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private string Tag => "OutlookAI-McpTest-p2-" + Marker;

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void AttachmentPaths_WhatOutlookDoesWithEach_AndWhatTheValidationSaidFirst()
    {
        string folder = Path.Combine(Path.GetTempPath(), Tag);
        string downloads = Path.Combine(AttachmentFileWriter.DownloadFolder, Tag);
        Directory.CreateDirectory(folder);
        List<string> verbatimFiles = new();
        List<string> drafts = new();
        List<string> lines = new();
        try
        {
            _output.WriteLine("Windows " + Environment.OSVersion.Version + " - " + RuntimeInformation.OSDescription
                + "; Outlook " + OutlookVersion());

            string plain = Path.Combine(folder, "p2-plain.txt");
            File.WriteAllText(plain, "p2 plain file " + Marker);
            string stream = plain + ":p2stream";
            File.WriteAllText(stream, "p2 alternate data stream, longer than the file " + Marker);

            string conFolder = Path.Combine(folder, "con");
            Directory.CreateDirectory(conFolder);
            string con = Path.Combine(conFolder, "CON.txt");
            File.WriteAllText(@"\\?\" + con, "p2 file called CON.txt " + Marker);
            verbatimFiles.Add(con);

            string trail = Path.Combine(folder, "p2-trail.txt. . .");
            File.WriteAllText(@"\\?\" + trail, "p2 file whose name ends in dots and spaces " + Marker);
            verbatimFiles.Add(trail);

            string longFolder = Path.Combine(folder, new string('d', 100), new string('e', 100));
            Directory.CreateDirectory(longFolder);
            string longFile = Path.Combine(longFolder, "p2-long-path.txt");
            File.WriteAllText(longFile, "p2 file whose path is longer than 260 characters " + Marker);

            // Safest first; a hang ends the probe, so the likeliest crashes come last.
            ProbeCase[] cases =
            {
                new("ordinary file, ordinary path", plain, File.ReadAllText(plain)),
                new("alternate data stream", stream, File.ReadAllText(stream)),
                new("CON.txt, ordinary path", con, File.ReadAllText(@"\\?\" + con)),
                new("path over 260 characters (" + longFile.Length.ToString(CultureInfo.InvariantCulture) + ")", longFile, File.ReadAllText(longFile)),
                new(@"ordinary file, \??\ path", @"\??\" + plain, File.ReadAllText(plain)),
                new(@"ordinary file, \\.\ path", @"\\.\" + plain, File.ReadAllText(plain)),
                new(@"ordinary file, \\?\ path", @"\\?\" + plain, File.ReadAllText(plain)),
                new(@"name ending in dots and spaces, \\?\ path", @"\\?\" + trail, File.ReadAllText(@"\\?\" + trail)),
                new(@"CON.txt, \\?\ path", @"\\?\" + con, File.ReadAllText(@"\\?\" + con)),
                new(@"path over 260 characters, \\?\ path", @"\\?\" + longFile, File.ReadAllText(longFile)),
            };

            bool wedged = false;
            int number = 0;
            foreach (ProbeCase probe in cases)
            {
                number++;
                string head = "#" + number.ToString(CultureInfo.InvariantCulture) + " " + probe.Label;
                if (wedged)
                {
                    lines.Add(head + ": NOT MEASURED - Outlook hung on an earlier case");
                    continue;
                }

                lines.Add(head + ": " + Measure(probe, number, drafts, downloads, out bool hung));
                _output.WriteLine(lines[^1]);
                wedged = hung;
            }
        }
        finally
        {
            foreach (string path in verbatimFiles)
            {
                ScratchFiles.TryDeleteFile(@"\\?\" + path);
            }

            ScratchFiles.TryDeleteTree(folder);
            ScratchFiles.TryDeleteTree(downloads);
            foreach (string draft in drafts)
            {
                LiveDraftCleanup.DeleteDraft(Hub, Marker, draft);
            }
        }

        _output.WriteLine(string.Empty);
        _output.WriteLine("SUMMARY - draft attachment paths on this guest:");
        foreach (string line in lines)
        {
            _output.WriteLine("  " + line);
        }

        LiveDraftCleanup.AssertNoTaggedArtifactsRemain(Hub, Marker, _output);
        Assert.Equal(10, lines.Count);
    }

    /// <summary>One case: the validation's verdict, then Outlook's, then what landed on the draft.</summary>
    private string Measure(ProbeCase probe, int number, List<string> drafts, string downloads, out bool hung)
    {
        hung = false;
        StringBuilder said = new StringBuilder();
        string handed = probe.Path;
        Task<string> validation = Task.Run(() =>
        {
            try
            {
                DraftAttachmentFile file = DraftAttachments.Validate(new[] { probe.Path })[0];
                handed = file.Path;
                return "validation ACCEPTS (name '" + file.FileName + "', " + file.SizeBytes.ToString(CultureInfo.InvariantCulture)
                    + " bytes, hands Outlook '" + Shown(file.Path) + "')";
            }
            catch (ArgumentException ex)
            {
                return "validation REFUSES (" + Reason(ex.Message) + ")";
            }
        });
        said.Append(validation.Wait(TimeSpan.FromSeconds(30)) ? validation.Result : "validation HUNG for 30 s");

        // An Outlook to watch: the previous case may have ended the last one.
        try
        {
            _ = Service.ListAccounts();
        }
        catch (Exception ex)
        {
            said.Append("; Outlook would not start (" + ex.GetType().Name + ": " + ex.Message + ")");
            return said.ToString();
        }

        Process[] watched = Process.GetProcessesByName("OUTLOOK");
        foreach (Process process in watched)
        {
            try
            {
                _ = process.SafeHandle;
            }
            catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
            {
            }
        }

        string? entryId = null;
        try
        {
            entryId = LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
                Hub, _fixture.TaggedSubject("p2 attachment path " + number.ToString(CultureInfo.InvariantCulture)),
                "p2 attachment path probe " + Marker, new[] { handed });
            drafts.Add(entryId);
            said.Append("; Outlook ATTACHED it");
        }
        catch (TimeoutException)
        {
            hung = true;
            said.Append("; Outlook HUNG - no answer within the mailer's budget");
            return said.ToString();
        }
        catch (InvalidOperationException ex)
        {
            Exception cause = ex.InnerException ?? ex;
            int hresult = cause.HResult;
            said.Append("; Outlook THREW " + cause.GetType().Name + " 0x" + hresult.ToString("X8", CultureInfo.InvariantCulture)
                + " (" + OneLine(cause.Message) + ")");
            if (hresult is RpcCallFailed or RpcServerUnavailable or RpcDisconnected or RpcServerDied or RpcServerDiedDne)
            {
                said.Append(Ended(watched));
            }
        }

        bool alive = watched.Length > 0 && watched.All(p => !HasExited(p));
        if (entryId != null)
        {
            said.Append(DescribeDraft(entryId, probe, downloads, number));
        }

        said.Append(alive ? "; OUTLOOK.EXE still running" : string.Empty);
        foreach (Process process in watched)
        {
            process.Dispose();
        }

        return said.ToString();
    }

    /// <summary>Whether the Outlook that took the call ended - waiting while Windows writes its dump.</summary>
    private static string Ended(Process[] watched)
    {
        if (watched.Length == 0)
        {
            return "; no OUTLOOK.EXE was running to watch";
        }

        StringBuilder ended = new StringBuilder();
        foreach (Process process in watched)
        {
            try
            {
                if (process.WaitForExit(TimeSpan.FromMinutes(3)))
                {
                    ended.Append("; OUTLOOK.EXE " + process.Id.ToString(CultureInfo.InvariantCulture) + " CRASHED, exit code 0x"
                        + process.ExitCode.ToString("X8", CultureInfo.InvariantCulture));
                }
                else
                {
                    ended.Append("; OUTLOOK.EXE " + process.Id.ToString(CultureInfo.InvariantCulture) + " still running after 3 minutes");
                }
            }
            catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
            {
                ended.Append("; OUTLOOK.EXE " + process.Id.ToString(CultureInfo.InvariantCulture) + " could not be watched (" + ex.Message + ")");
            }
        }

        return ended.ToString();
    }

    /// <summary>What the saved draft carries: the attachment's name and whose content it holds.</summary>
    private string DescribeDraft(string entryId, ProbeCase probe, string downloads, int number)
    {
        try
        {
            ReadOutcome read = Service.Read(entryId, maxBodyChars: 0);
            if (read.Attachments.Count != 1)
            {
                return "; the saved draft holds " + read.Attachments.Count.ToString(CultureInfo.InvariantCulture) + " attachments";
            }

            AttachmentView attachment = read.Attachments[0];
            SaveAttachmentOutcome saved = Service.SaveAttachment(entryId, 1, Path.Combine(downloads, number.ToString(CultureInfo.InvariantCulture)));
            string content = File.ReadAllText(saved.SavedPath);
            string match = string.Equals(content, probe.Content, StringComparison.Ordinal)
                ? "the file's own content"
                : "OTHER content ('" + OneLine(content) + "')";
            return "; the draft's attachment is '" + Shown(attachment.FileName ?? "(no name)") + "', holding " + match;
        }
        catch (Exception ex)
        {
            return "; the draft could not be read back (" + ex.GetType().Name + ": " + OneLine(ex.Message) + ")";
        }
    }

    private static bool HasExited(Process process)
    {
        try
        {
            return process.HasExited;
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            return true;
        }
    }

    private static string OutlookVersion()
    {
        foreach (Process process in Process.GetProcessesByName("OUTLOOK"))
        {
            using (process)
            {
                try
                {
                    return process.MainModule?.FileVersionInfo.FileVersion ?? "(no version)";
                }
                catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
                {
                    return "(version unreadable)";
                }
            }
        }

        return "(not running)";
    }

    /// <summary>The refusal's reason for this one path, without the guidance around it.</summary>
    private static string Reason(string message)
    {
        int quote = message.IndexOf("': ", StringComparison.Ordinal);
        string reason = quote < 0 ? message : message.Substring(quote + 3);
        int end = reason.IndexOf(". ", StringComparison.Ordinal);
        return OneLine(end < 0 ? reason : reason.Substring(0, end));
    }

    /// <summary>A path with the temporary folder's own part left out, so the log shows the shape.</summary>
    private static string Shown(string path)
    {
        string temp = Path.GetTempPath().TrimEnd('\\');
        int at = path.IndexOf(temp, StringComparison.OrdinalIgnoreCase);
        return at < 0 ? path : path.Substring(0, at) + "%TEMP%" + path.Substring(at + temp.Length);
    }

    private static string OneLine(string text)
    {
        return text.Replace('\r', ' ').Replace('\n', ' ');
    }

    private sealed record ProbeCase(string Label, string Path, string Content);
}
