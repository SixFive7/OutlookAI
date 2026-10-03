using System.Reflection;
using System.Text.RegularExpressions;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Every COM object a collection's <c>Add</c> hands back is held and released by the code that
/// asked for it - never discarded for the garbage collector to release later.
/// <para>
/// Why it matters (2026-10-03, the test guest's Outlook crashes): a discarded result - the
/// <c>Column</c> from <c>Table.Columns.Add</c>, the <c>Bookmark</c> from <c>Bookmarks.Add</c>, the
/// <c>Attachment</c> from <c>Attachments.Add</c> - stays referenced from this process until a
/// garbage collection releases it, at a moment nobody chose: after its table, its document or its
/// mail has already gone. Outlook then destroys a child whose parent is gone, inside its own process.
/// When that collection lands is decided by how much the process allocated, so it moves with every
/// build; that is the one difference left between the build that crashed the guest's Outlook five
/// times in six (OLMAPI32.DLL and ntdll heap faults, at varying places) and the one that never did
/// in five, after their COM calls were compared line by line and found identical.
/// </para>
/// <para>
/// Read out of the sources, because the failure needs a real Outlook. Nothing here touches one.
/// </para>
/// </summary>
public sealed class ComChildObjectReleaseTests
{
    /// <summary>
    /// A statement that calls <c>Add</c> on a receiver the scanned files use for a COM collection
    /// whose <c>Add</c> returns a child object, and keeps nothing. <c>Add(new ...)</c> is a managed
    /// list, never a COM call. (<c>recipients</c> is left out: the product names a managed list so,
    /// and its <c>Recipients.Add</c> goes through a receiver called <c>collection</c>, held.)
    /// </summary>
    private static readonly Regex DiscardedComAdd = new Regex(
        @"^\s*(\(\(dynamic\)\s*)?(columns|cols|bm|bookmarks|attachments)!?\)?\.Add\((?!new )",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static TheoryData<string> ScannedFiles => new TheoryData<string>
    {
        Path.Combine("OutlookAI.Core", "Com", "OutlookComSession.cs"),
        Path.Combine("OutlookAI.Core", "Com", "ComposeSurface.cs"),
        Path.Combine("OutlookAI.Core", "Com", "SpecialFolders.cs"),
        Path.Combine("OutlookAI.McpServer.Tests", "T2", "LiveOutlookTestMailer.cs"),
    };

    [Theory]
    [MemberData(nameof(ScannedFiles))]
    public void NoCollectionAdd_DiscardsTheComObjectItReturns(string relativePath)
    {
        string[] lines = File.ReadAllLines(Path.GetFullPath(Path.Combine(TestProjectDir(), "..", relativePath)));
        Assert.True(lines.Length > 100, relativePath + " is too short to be the file this test reads");

        List<string> offenders = new List<string>();
        for (int i = 0; i < lines.Length; i++)
        {
            if (DiscardedComAdd.IsMatch(lines[i]))
            {
                offenders.Add((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ": " + lines[i].Trim());
            }
        }

        Assert.True(offenders.Count == 0, relativePath + " discards what a COM Add returns:\n" + string.Join("\n", offenders));
    }

    [Fact]
    public void Control_ThePatternCatchesTheShapesThatLeaked()
    {
        // The exact lines the 2026-10-03 fix replaced - so the scan above is known to see them.
        Assert.Matches(DiscardedComAdd, "                    ((dynamic)columns!).Add(DateSortProperties[i]);");
        Assert.Matches(DiscardedComAdd, "                    bm.Add(\"_MailAutoSig\", newRange);");
        Assert.Matches(DiscardedComAdd, "                                attachments.Add(attachmentPath);");
        Assert.Matches(DiscardedComAdd, "                columns!.Add(spelling);");

        // And what replaced them, which keeps the object to release it.
        Assert.DoesNotMatch(DiscardedComAdd, "                    column = ((dynamic)columns!).Add(DateSortProperties[i]);");
        Assert.DoesNotMatch(DiscardedComAdd, "                    newMark = bm.Add(\"_MailAutoSig\", newRange);");
        Assert.DoesNotMatch(DiscardedComAdd, "                                object? added = attachments.Add(attachmentPath);");
    }

    [Fact]
    public void TheMailerHoldsEveryCollectionItCounts()
    {
        // The inline chain the same fix took out: folder.Items.Count left the Items collection for
        // the garbage collector. Collections are held and released one by one there.
        string mailer = File.ReadAllText(Path.GetFullPath(Path.Combine(
            TestProjectDir(), "..", "OutlookAI.McpServer.Tests", "T2", "LiveOutlookTestMailer.cs")));
        Assert.DoesNotMatch(new Regex(@"\b\w+\.(Items|Folders|Attachments|Recipients|Columns)\.(Count|Add|Item)\b"), mailer);
    }

    private static string TestProjectDir()
    {
        return typeof(ComChildObjectReleaseTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }
}
