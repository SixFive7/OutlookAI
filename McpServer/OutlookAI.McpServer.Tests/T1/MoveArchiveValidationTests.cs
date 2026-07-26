using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D39 pure-logic coverage: move_mail/archive_mail arg validation (pre-COM), the
/// same-store-guard and target-guard error texts, the EntryID-change advice pin, the
/// store-relative path derivation, and designated-Archive-folder resolution parsing
/// against synthetic property values (the live carriers are documented in
/// <see cref="ArchiveFolderResolution"/>).
/// </summary>
public sealed class MoveArchiveValidationTests
{
    // ------------------------------------------------------------ ids validation (pre-COM)

    [Fact]
    public void MoveIdsCap_IsPinnedAtFifty()
    {
        Assert.Equal(50, MailService.MoveIdsCap);
    }

    [Fact]
    public void ValidateMoveIds_NullOrEmpty_Throws()
    {
        Assert.Throws<ArgumentException>(() => MailService.ValidateMoveIds(null));
        Assert.Throws<ArgumentException>(() => MailService.ValidateMoveIds(Array.Empty<string>()));
    }

    [Fact]
    public void ValidateMoveIds_OverCap_Throws_AndCapExactlyPasses()
    {
        string[] tooMany = Enumerable.Range(0, MailService.MoveIdsCap + 1).Select(i => "h" + i).ToArray();
        ArgumentException ex = Assert.Throws<ArgumentException>(() => MailService.ValidateMoveIds(tooMany));
        Assert.Contains("50", ex.Message, StringComparison.Ordinal);

        string[] exactlyCap = Enumerable.Range(0, MailService.MoveIdsCap).Select(i => "h" + i).ToArray();
        Assert.Equal(MailService.MoveIdsCap, MailService.ValidateMoveIds(exactlyCap).Count);
    }

    [Fact]
    public void ValidateMoveIds_BlankEntry_Throws()
    {
        Assert.Throws<ArgumentException>(() => MailService.ValidateMoveIds(new[] { "h1", "  " }));
    }

    [Fact]
    public void ValidateMoveIds_DuplicateIds_Throw_CaseInsensitively()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => MailService.ValidateMoveIds(new[] { "ABCDEF", "abcdef" }));
        Assert.Contains("Duplicate", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidateMoveIds_TrimsAndPreservesOrder()
    {
        IReadOnlyList<string> cleaned = MailService.ValidateMoveIds(new[] { " h2 ", "h1" });
        Assert.Equal(new[] { "h2", "h1" }, cleaned);
    }

    // ------------------------------------------------------------ error text mapping

    [Fact]
    public void DescribeMoveFailure_CrossStore_NamesTheItemStore_AndV1Restriction()
    {
        string text = MailService.DescribeMoveFailure("CrossStoreTarget:Some Store", "Target", "Other Store", createFolder: false);
        Assert.Contains("Some Store", text, StringComparison.Ordinal);
        Assert.Contains("Other Store", text, StringComparison.Ordinal);
        Assert.Contains("same-store only", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DescribeMoveFailure_TargetNotFound_PointsAtCreateFolder_OnlyWhenNotCreating()
    {
        string withoutCreate = MailService.DescribeMoveFailure("TargetFolderNotFound", "A/B", null, createFolder: false);
        Assert.Contains("create_folder=true", withoutCreate, StringComparison.Ordinal);
        Assert.Contains("A/B", withoutCreate, StringComparison.Ordinal);

        string withCreate = MailService.DescribeMoveFailure("TargetFolderNotFound", "A/B", null, createFolder: true);
        Assert.DoesNotContain("create_folder=true", withCreate, StringComparison.Ordinal);
    }

    [Fact]
    public void DescribeMoveFailure_DeletedItemsTarget_IsRefusedAsDeletionSemantics()
    {
        string text = MailService.DescribeMoveFailure("TargetIsDeletedItems", "Deleted Items", null, createFolder: false);
        Assert.Contains("deletion semantics", text, StringComparison.Ordinal);
        Assert.Contains("no delete surface", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("ItemNotFound", "Re-run search")]
    [InlineData("NotAMailItem", "mail items")]
    [InlineData("TargetIsOutbox", "Outbox")]
    [InlineData("TargetNotAMailFolder", "not a mail folder")]
    [InlineData("AlreadyInTargetFolder", "already in the target folder")]
    [InlineData(null, "unknown")]
    public void DescribeMoveFailure_KnownReasons_MapToActionableText(string? reason, string expectedFragment)
    {
        string text = MailService.DescribeMoveFailure(reason, "T", null, createFolder: false);
        Assert.Contains(expectedFragment, text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DescribeArchiveResolutionFailure_NoDesignatedFolder_SaysNothingWasCreated()
    {
        string text = MailService.DescribeArchiveResolutionFailure("Store X", "NoDesignatedArchiveFolder");
        Assert.Contains("Store X", text, StringComparison.Ordinal);
        Assert.Contains("Nothing was created", text, StringComparison.Ordinal);
        Assert.Contains("Archive button", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DescribeArchiveResolutionFailure_VerificationFailure_RefusesWithReason()
    {
        string text = MailService.DescribeArchiveResolutionFailure("S", "ArchiveFolderVerificationFailed:coreDefault");
        Assert.Contains("ArchiveFolderVerificationFailed:coreDefault", text, StringComparison.Ordinal);
        Assert.Contains("refusing", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MoveEntryIdAdvice_CarriesUndoAndStalenessGuidance()
    {
        Assert.Contains("newEntryId", MailService.MoveEntryIdAdvice, StringComparison.Ordinal);
        Assert.Contains("fromFolder", MailService.MoveEntryIdAdvice, StringComparison.Ordinal);
        Assert.Contains("re-run search", MailService.MoveEntryIdAdvice, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("move_mail", MailService.MoveEntryIdAdvice, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------ store-relative paths

    [Theory]
    [InlineData("\\\\My Store\\Inbox", "My Store", "Inbox")]
    [InlineData("\\\\My Store\\A\\B", "My Store", "A/B")]
    [InlineData("\\\\my store\\A", "My Store", "A")] // store-name case-insensitive
    [InlineData("\\\\My Store", "My Store", "")] // the root itself
    [InlineData("\\\\Other Store\\Inbox", "My Store", "Inbox")] // unknown prefix: first segment dropped
    [InlineData(null, "My Store", "")]
    [InlineData("", "My Store", "")]
    public void ToStoreRelativeFolderPath_DerivesListFoldersConvention(string? folderPath, string store, string expected)
    {
        Assert.Equal(expected, OutlookComSession.ToStoreRelativeFolderPath(folderPath, store));
    }

    [Fact]
    public void ToStoreRelativeFolderPath_StoreNamePrefixCollision_FallsBackToSegmentDrop()
    {
        // "Jan" must not eat the front of "Jan van Linge" - the store name has to be
        // the WHOLE first segment.
        Assert.Equal("Inbox", OutlookComSession.ToStoreRelativeFolderPath("\\\\Jan van Linge\\Inbox", "Jan"));
    }

    // ------------------------------------------------------------ archive resolution parsing

    [Fact]
    public void ArchiveResolution_Constants_ArePinned()
    {
        // 39 = the undocumented-but-live-proven OlDefaultFolders archive value; the
        // fallback carrier is PR_IPM_ARCHIVE_ENTRYID on the store object.
        Assert.Equal(39, ArchiveFolderResolution.OlFolderArchive);
        Assert.EndsWith("0x35FF0102", ArchiveFolderResolution.ArchiveEntryIdPropertySchema, StringComparison.Ordinal);
    }

    [Fact]
    public void TryReadEntryIdHex_Binary_BecomesUppercaseHex()
    {
        Assert.Equal("01ABCDEF", ArchiveFolderResolution.TryReadEntryIdHex(new byte[] { 0x01, 0xAB, 0xCD, 0xEF }));
    }

    [Fact]
    public void TryReadEntryIdHex_TooShortBinary_IsRejected()
    {
        // Live-probed: slot lists end with a 4-byte non-EntryID trailer; anything under
        // 4 bytes is filler, never an id.
        Assert.Null(ArchiveFolderResolution.TryReadEntryIdHex(new byte[] { 0x00, 0x01, 0x02 }));
        Assert.Null(ArchiveFolderResolution.TryReadEntryIdHex(Array.Empty<byte>()));
    }

    [Fact]
    public void TryReadEntryIdHex_MultiValueBinary_TakesFirstUsableEntry()
    {
        object multiValue = new object[]
        {
            new byte[] { 0x01 }, // filler, skipped
            new byte[] { 0xAA, 0xBB, 0xCC, 0xDD },
        };
        Assert.Equal("AABBCCDD", ArchiveFolderResolution.TryReadEntryIdHex(multiValue));
    }

    [Theory]
    [InlineData("abcd1234", "ABCD1234")] // hex string passthrough, normalized
    [InlineData("  ABCD1234  ", "ABCD1234")]
    public void TryReadEntryIdHex_HexStrings_AreAcceptedNormalized(string value, string expected)
    {
        Assert.Equal(expected, ArchiveFolderResolution.TryReadEntryIdHex(value));
    }

    [Theory]
    [InlineData("not-hex-at-all")]
    [InlineData("ABC")] // odd length
    [InlineData("AB12")] // too short (< 8 chars)
    public void TryReadEntryIdHex_JunkStrings_AreRejected(string value)
    {
        Assert.Null(ArchiveFolderResolution.TryReadEntryIdHex(value));
    }

    [Fact]
    public void TryReadEntryIdHex_NullAndForeignTypes_AreRejected()
    {
        Assert.Null(ArchiveFolderResolution.TryReadEntryIdHex(null));
        Assert.Null(ArchiveFolderResolution.TryReadEntryIdHex(42));
        Assert.Null(ArchiveFolderResolution.TryReadEntryIdHex(new object[] { 42, "x" }));
    }
}
