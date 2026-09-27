using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins, from the source, what the undated write path does after its first save - the one part of
/// <c>ComCorpusMailbox.CreateUndatedItem</c> no pure test can reach, and the part the first guest run
/// of population v2 proved necessary.
/// <para>
/// <b>Why.</b> On OAI-UNINDEXED, 2026-09-27, the undated probe against the hub refused all three kinds:
/// an appointment, a contact and a task, each saved once into a PST, read back CARRYING
/// PR_MESSAGE_DELIVERY_TIME ("undated=False"). An undated item exists to carry none - the three
/// <c>LiveOrderKeyCollationTests</c> measure where rows with no received date sort - so the write path
/// now removes the property after the first save and saves again, and the probe and the build's
/// read-back re-open the item and check it stayed gone. This test holds the removal in place, and holds
/// it to the undated path alone: a dated item's delivery time is what dates it.
/// </para>
/// </summary>
public sealed class CorpusUndatedWritePathTests
{
    [Fact]
    public void TheUndatedWritePath_RemovesTheDeliveryTime_AfterItsFirstSave()
    {
        string body = MethodBody(Source(), "private static (string EntryId, string? SavedStoreId) CreateUndatedItem(");
        int save = body.IndexOf("item.Save();", StringComparison.Ordinal);
        int remove = body.IndexOf("RemoveDeliveryTime(", StringComparison.Ordinal);
        Assert.True(save >= 0, "CreateUndatedItem no longer saves the item");
        Assert.True(remove > save, "CreateUndatedItem must remove PR_MESSAGE_DELIVERY_TIME AFTER the first save - before it, there is nothing to remove");
        Assert.True(
            remove < body.IndexOf("return (", StringComparison.Ordinal),
            "the removal must happen before the EntryID is handed back to the probe and the build");
    }

    [Fact]
    public void TheRemoval_DeletesThatOneProperty_ToleratesItsAbsence_AndSavesAgain()
    {
        string body = MethodBody(Source(), "private static void RemoveDeliveryTime(");
        Assert.Contains("DeleteProperty(PrMessageDeliveryTime)", body, StringComparison.Ordinal);
        Assert.Contains("IsPropertyNotFound", body, StringComparison.Ordinal);
        Assert.Contains("item.Save();", body, StringComparison.Ordinal);

        // It touches nothing but the delivery time.
        Assert.DoesNotContain("PrClientSubmitTime", body, StringComparison.Ordinal);
        Assert.DoesNotContain("PrMessageFlags", body, StringComparison.Ordinal);
    }

    [Fact]
    public void ADatedItem_KeepsItsDeliveryTime_TheControl()
    {
        // The control: the dated write path never calls the removal. Its delivery time is written on
        // purpose, from the plan, and it is what the freshness sweep and every window count read.
        string source = Source();
        Assert.Single(
            System.Text.RegularExpressions.Regex.Matches(source, @"RemoveDeliveryTime\(item!\)"));
        foreach (string dated in new[] { "private static void ApplyMessageFlags(", "private static DateTime? ApplyDates(" })
        {
            Assert.DoesNotContain("RemoveDeliveryTime", MethodBody(source, dated), StringComparison.Ordinal);
        }
    }

    private static string Source()
        => File.ReadAllText(Path.Combine(RepoRoot(), "McpServer", "OutlookAI.RemediationTools", "ComCorpusMailbox.cs"));

    /// <summary>The text between a method's opening and closing brace at the class member indentation.</summary>
    private static string MethodBody(string source, string signatureStart)
    {
        int start = source.IndexOf(signatureStart, StringComparison.Ordinal);
        Assert.True(start >= 0, "not found in ComCorpusMailbox.cs: " + signatureStart);
        string newline = source.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
        int open = source.IndexOf(newline + "    {", start, StringComparison.Ordinal);
        int close = source.IndexOf(newline + "    }", open + 1, StringComparison.Ordinal);
        Assert.True(open > start && close > open, "could not find the body of " + signatureStart);
        return source.Substring(open, close - open);
    }

    private static string RepoRoot()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }
}
