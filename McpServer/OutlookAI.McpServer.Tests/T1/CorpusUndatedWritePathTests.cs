using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins, from the source, what the undated write path does after its first save - the one part of
/// <c>ComCorpusMailbox.CreateUndatedItem</c> no pure test can reach, and the part the first guest runs
/// of population v2 proved necessary.
/// <para>
/// <b>Why.</b> On OAI-UNINDEXED, 2026-09-27, the undated probe against the hub refused all three kinds:
/// an appointment, a contact and a task, each saved once into a PST, read back CARRYING
/// PR_MESSAGE_DELIVERY_TIME ("undated=False"). An undated item exists to carry none - the three
/// <c>LiveOrderKeyCollationTests</c> measure where rows with no received date sort - so the write path
/// removes the property after the first save and saves again, and the probe and the build's read-back
/// re-open the item and check it stayed gone. The next run measured that the PropertyAccessor REFUSES
/// that removal for an appointment (<c>UnauthorizedAccessException</c>, "does not support this
/// operation"), and the refusal escaped the probe and ended the run - so a refusal is now returned, the
/// probe reports it per kind, and the build records the item before it refuses. These tests hold all of
/// that in place, and hold the removal to the undated path alone: a dated item's delivery time is what
/// dates it.
/// </para>
/// </summary>
public sealed class CorpusUndatedWritePathTests
{
    private const string CreateSignature =
        "private static (string EntryId, string? SavedStoreId, string? RemovalRefused) CreateUndatedItem(";

    [Fact]
    public void TheUndatedWritePath_TriesTheRemoval_AfterItsFirstSave_WithTheEntryIdAlreadyInHand()
    {
        string body = MethodBody(Source(), CreateSignature);
        int save = body.IndexOf("item.Save();", StringComparison.Ordinal);
        int entryId = body.IndexOf("string entryId = (string)item!.EntryID;", StringComparison.Ordinal);
        int remove = body.IndexOf("TryRemoveDeliveryTime(", StringComparison.Ordinal);
        Assert.True(save >= 0, "CreateUndatedItem no longer saves the item");
        Assert.True(entryId > save, "the EntryID must be read right after the first save");
        Assert.True(
            remove > entryId,
            "the removal must come AFTER the EntryID is read - a refusal has to be able to hand the item back for deletion");
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
    public void ARefusedRemoval_IsReturned_NotThrown_AndAccessDeniedCounts()
    {
        string source = Source();
        string attempt = MethodBody(source, "private static string? TryRemoveDeliveryTime(");
        Assert.Contains("catch (Exception ex) when (IsUndatedWriteRefusal(ex))", attempt, StringComparison.Ordinal);
        Assert.Contains("return ToolFailure.Describe(ex);", attempt, StringComparison.Ordinal);

        // The measured shape: E_ACCESSDENIED arrives as UnauthorizedAccessException, which the shared COM
        // failure test does not include.
        string refusal = source.Substring(source.IndexOf("private static bool IsUndatedWriteRefusal(", StringComparison.Ordinal));
        refusal = refusal.Substring(0, refusal.IndexOf(';') + 1);
        Assert.Contains("OutlookComSession.IsComCallFailure(ex)", refusal, StringComparison.Ordinal);
        Assert.Contains("ex is UnauthorizedAccessException", refusal, StringComparison.Ordinal);

        // And the probe for one kind catches the same set, so no refusal ends the whole run again.
        string probe = MethodBody(source, "private static CorpusUndatedProbe RunOneUndatedProbe(");
        Assert.Contains("catch (Exception ex) when (IsUndatedWriteRefusal(ex))", probe, StringComparison.Ordinal);
        Assert.Contains("CorpusUndatedTable.Filter(corpusId, withReceivedDate: false)", probe, StringComparison.Ordinal);
        Assert.Contains("CorpusUndatedTable.Filter(corpusId, withReceivedDate: true)", probe, StringComparison.Ordinal);
    }

    [Fact]
    public void TheBuild_RecordsAnUndatedItem_BeforeItRefusesOne()
    {
        string source = Source();
        int create = source.IndexOf("CreateUndatedItem((object)undatedItems!", StringComparison.Ordinal);
        Assert.True(create >= 0, "the build no longer creates undated items through CreateUndatedItem");
        int record = source.IndexOf("record(undatedLine);", create, StringComparison.Ordinal);
        int refuse = source.IndexOf("RequireDeliveryTimeRemoved(ordinal, removalRefused);", create, StringComparison.Ordinal);
        Assert.True(record > create, "the build must record the undated item it created");
        Assert.True(
            refuse > record,
            "the build must RECORD the item before refusing it - an unrecorded item is one teardown can never delete");
    }

    [Fact]
    public void ADatedItem_KeepsItsDeliveryTime_TheControl()
    {
        // The control: the dated write path never calls the removal. Its delivery time is written on
        // purpose, from the plan, and it is what the freshness sweep and every window count read.
        string source = Source();
        Assert.Single(
            System.Text.RegularExpressions.Regex.Matches(source, @"TryRemoveDeliveryTime\(\(object\)item!\)"));
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
