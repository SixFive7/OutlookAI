using System;
using System.IO;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the one thing that stopped <c>corpus-build</c> ever checking its own work.
/// <para>
/// <b>What happened, measured on the unindexed test guest 2026-09-16.</b> A 20,000-item build
/// wrote every item and every manifest line - 20,001 lines, 0 unparseable, 20,000 distinct
/// EntryIDs, last ordinal 20000, a 1,024 MB store - and then died with
/// <c>IOException: The process cannot access the file '...corpus-vm-unindexed.jsonl' because it
/// is being used by another process</c>, exit 1.
/// </para>
/// <para>
/// The other process was itself. <c>RunBuild</c> held the manifest open for append through the
/// whole run (a using DECLARATION, disposed only at method exit) and then called the census,
/// which read the manifest back with <see cref="File.ReadLines(string)"/>. That overload opens
/// <c>FileAccess.Read, FileShare.Read</c>, and <c>FileShare.Read</c> refuses to coexist with an
/// existing WRITE handle - so the read could not succeed while the build was the one asking.
/// It is deterministic, not a race: no third party and no timing is needed to reproduce it.
/// </para>
/// <para>
/// The cost was not the exit code. The census is the guard that exists BECAUSE a build once
/// created 40,000 items, put every one of them in Drafts, and reported success - see the
/// comment above <c>RunCensusPass</c>. For as long as this stood, that guard had never once run
/// inside a build, and no test could have noticed: every test that reads a manifest reads one
/// nobody is holding.
/// </para>
/// <para>
/// Two fixes were applied and both are pinned here: the reader tolerates a live writer
/// (<see cref="CorpusManifest.ReadFile"/>), and the build scopes its writer so the census reads
/// a finished file. <see cref="ReadFile_ReadsAManifestThatIsStillOpenForAppend"/> is the one
/// that fails if the reader is put back to <c>File.ReadLines</c>;
/// <see cref="FileReadLines_CannotReadAFileHeldOpenForWrite"/> is the control that proves the
/// first test is testing something real rather than passing vacuously.
/// </para>
/// </summary>
public sealed class CorpusManifestSharedReadTests
{
    /// <summary>The header every manifest in this file starts with.</summary>
    private static string Header() => CorpusManifest.RenderLine(new CorpusManifestHeader(
        CorpusManifest.CurrentVersion,
        "share-read",
        7777,
        CorpusManifest.FormatUtc(new DateTime(2026, 9, 16, 0, 0, 0, DateTimeKind.Utc)),
        new CorpusPlanOptions("share-read", 7777, new DateTime(2026, 9, 16, 0, 0, 0, DateTimeKind.Utc)).ShapeKey,
        "Outlook Data File",
        null,
        "PropertyAccessorDates",
        "DraftsThenMoveWithSentFlag"));

    /// <summary>
    /// THE REGRESSION. A build reads its own manifest back while still appending to it, so the
    /// reader must accept a file another handle holds open for write. Revert
    /// <c>CorpusManifest.ReadFile</c> to <c>File.ReadLines</c> and this test throws the exact
    /// IOException the VM produced.
    /// </summary>
    [Fact]
    public void ReadFile_ReadsAManifestThatIsStillOpenForAppend()
    {
        string path = NewTempPath();
        try
        {
            // Exactly what CorpusCommands.OpenManifest does: StreamWriter on the path, header
            // written and flushed, then held open for the duration of the build.
            using (var writer = new StreamWriter(path, append: false))
            {
                writer.WriteLine(Header());
                writer.Flush();
                for (int ordinal = 1; ordinal <= 3; ordinal++)
                {
                    writer.WriteLine(CorpusManifest.RenderLine(new CorpusManifestItem(
                        ordinal,
                        "ENTRY" + ordinal.ToString("D8", System.Globalization.CultureInfo.InvariantCulture),
                        6,
                        1024,
                        CorpusManifest.FormatUtc(new DateTime(2026, 9, 1, 12, 0, 0, DateTimeKind.Utc)))));
                    writer.Flush();
                }

                // The census's read, performed while the writer above is still open.
                CorpusManifest read = CorpusManifest.ReadFile(path);

                Assert.Equal(3, read.Items.Count);
                Assert.Empty(read.UnparseableLines);
                Assert.Equal("share-read", read.Header.CorpusId);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// THE CONTROL. Proves the share mode is what the test above turns on: the same read
    /// through <see cref="File.ReadLines(string)"/>, against the same still-open writer, throws.
    /// Without this, a future change that made the writer close early would leave the
    /// regression test passing for the wrong reason.
    /// </summary>
    [Fact]
    public void FileReadLines_CannotReadAFileHeldOpenForWrite()
    {
        string path = NewTempPath();
        try
        {
            using (var writer = new StreamWriter(path, append: false))
            {
                writer.WriteLine(Header());
                writer.Flush();

                Assert.Throws<IOException>(() => CorpusManifest.Parse(File.ReadLines(path)));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// A manifest read mid-append sees a PREFIX of the truth, never a corrupt one - which is
    /// what makes tolerating the live writer safe. The writer flushes per line, so the reader
    /// sees whole lines only; a torn trailing line is already reported rather than thrown.
    /// </summary>
    [Fact]
    public void ReadFile_MidAppend_SeesThePrefixWrittenSoFar()
    {
        string path = NewTempPath();
        try
        {
            using (var writer = new StreamWriter(path, append: false))
            {
                writer.WriteLine(Header());
                writer.Flush();

                writer.WriteLine(CorpusManifest.RenderLine(new CorpusManifestItem(
                    1, "ENTRY00000001", 6, 512, CorpusManifest.FormatUtc(new DateTime(2026, 9, 1, 12, 0, 0, DateTimeKind.Utc)))));
                writer.Flush();
                Assert.Single(CorpusManifest.ReadFile(path).Items);

                writer.WriteLine(CorpusManifest.RenderLine(new CorpusManifestItem(
                    2, "ENTRY00000002", 6, 512, CorpusManifest.FormatUtc(new DateTime(2026, 9, 2, 12, 0, 0, DateTimeKind.Utc)))));
                writer.Flush();
                Assert.Equal(2, CorpusManifest.ReadFile(path).Items.Count);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string NewTempPath()
        => Path.Combine(Path.GetTempPath(), "corpus-share-" + Guid.NewGuid().ToString("N") + ".jsonl");
}
