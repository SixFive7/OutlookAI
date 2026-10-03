using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: folder and store NAMES holding <c>% / \ * ?</c> through every place the product turns a
/// name into an index URL or an index URL back into a name (the Q99 finding, measured by
/// <c>T2/LiveFolderNameEncodingTests</c>; <c>Docs/live-tier-on-the-vm.md</c> section 8 item 26).
/// <para>
/// The defect these pin: a folder scope was built from the RAW name, so a folder called
/// <c>50% off</c> was scoped as <c>.../0/50% off</c> against an index that files it as
/// <c>.../0/50%25 off</c> - zero rows, no error - and a hit from such a folder carried the URL's
/// spelling as its folder name, which named no folder in Outlook. Every ordinary name must stay
/// byte for byte what it was; that half is pinned here as hard as the fix.
/// </para>
/// <para>Fixtures are synthetic (S6) and shaped like the measured URLs.</para>
/// </summary>
public sealed class FolderNameEncodingTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";
    private const string Prefix = "mapi16://" + Sid + "/alice@example.com($deadbeef)";

    // ------------------------------------------------------- building the scope

    [Theory]
    [InlineData("Inbox")]
    [InlineData("Inbox/Fun/Immich")]
    [InlineData("Auto ongeluk")]
    [InlineData("Clients/O'Brien")]
    [InlineData("Projects [2026] {draft} #1 & co")]
    [InlineData("Postvak IN/Ελληνικά/가나다")]
    public void AnOrdinaryFolder_IsScopedByteForByteAsItAlwaysWas(string folder)
    {
        foreach (bool includeSubfolders in new[] { true, false })
        {
            FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(Prefix, folder, includeSubfolders);

            Assert.Equal(Prefix + "/0/" + folder, r.Scope);
            Assert.Equal(folder, r.RequestedFolder);
            if (!includeSubfolders)
            {
                Assert.Equal(new[] { "/alice@example.com/" + folder }, r.FolderPaths!);
            }
        }
    }

    [Theory]
    [InlineData("50% off", "50%25 off")]
    [InlineData("star*", "star%2A")]
    [InlineData("why?", "why%3F")]
    [InlineData("back\\slash", "back%5Cslash")]
    [InlineData("100%*? mix", "100%25%2A%3F mix")]
    [InlineData("%2A not a star", "%252A not a star")]
    public void AFolderNameHoldingOneOfTheFive_IsScopedAsTheIndexSpellsIt(string name, string spelled)
    {
        FolderScopeResolution recursive = FolderScopeResolver.ForPrimaryStore(Prefix, "Clients/" + name, true);
        Assert.Equal(FolderScopeKind.PrimaryRecursive, recursive.Kind);
        Assert.Equal(Prefix + "/0/Clients/" + spelled, recursive.Scope);
        Assert.Null(recursive.FolderPaths);

        // The sweep and the exhaustive scan look the folder up in Outlook by this, so it stays raw.
        Assert.Equal("Clients/" + name, recursive.RequestedFolder);

        FolderScopeResolution own = FolderScopeResolver.ForPrimaryStore(Prefix, "Clients/" + name, false);
        Assert.Equal(FolderScopeKind.PrimaryNonRecursive, own.Kind);
        Assert.Equal(Prefix + "/0/Clients/" + spelled, own.Scope);

        // The folder-path equality carries NAMES (System.ItemFolderPathDisplay - measured).
        Assert.Equal(new[] { "/alice@example.com/Clients/" + name }, own.FolderPaths!);
    }

    [Fact]
    public void ANestedPath_EncodesEverySegment_AndOnlyWhatNeedsIt()
    {
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(Prefix, "/Inbox/50% off/Q1*?\\x/plain/", false);

        Assert.Equal(Prefix + "/0/Inbox/50%25 off/Q1%2A%3F%5Cx/plain", r.Scope);
        Assert.Equal("Inbox/50% off/Q1*?\\x/plain", r.RequestedFolder);
        Assert.Equal(new[] { "/alice@example.com/Inbox/50% off/Q1*?\\x/plain" }, r.FolderPaths!);
    }

    [Fact]
    public void TheStoreHalf_IsTheIndexsOwnSpelling_AndItsDisplayPathIsTheName()
    {
        // The store prefix comes from the index's root listing, already spelled; the display path
        // derived from it carries the store's NAME.
        const string specialStore = "mapi16://" + Sid + "/q99 50%25 off%2A%3Fx($5159380d)";
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(specialStore, "Inbox/a*b", false);

        Assert.Equal(specialStore + "/0/Inbox/a%2Ab", r.Scope);
        Assert.Equal(specialStore, r.StoreScope);
        Assert.Equal(new[] { "/q99 50% off*?x/Inbox/a*b" }, r.FolderPaths!);
    }

    [Theory]
    [InlineData(Prefix + "/0/Inbox/50%25 off", "/alice@example.com/Inbox/50% off")]
    [InlineData(Prefix + "/0/a%2Fb/c%5Cd", "/alice@example.com/a/b/c\\d")]
    [InlineData(Prefix + "/0/%252A", "/alice@example.com/%2A")]
    [InlineData("mapi16://" + Sid + "/q99 50%25 off%2A%3Fx($5159380d)", "/q99 50% off*?x")]
    [InlineData(Prefix + "/1/Sam 50%25/Q%3F", "/alice@example.com/Sam 50%/Q?")]
    public void TheDisplayPathDerivedFromAUrl_HoldsNames(string url, string expected)
    {
        Assert.True(MapiItemUrl.TryBuildFolderPathDisplay(url, out string? path));
        Assert.Equal(expected, path);
    }

    // ------------------------------------------------------- the delegate half

    [Theory]
    [InlineData("Sam Delegate", Prefix + "/1/Sam Delegate")]
    [InlineData("Sam 50% *?", Prefix + "/1/Sam 50%25 %2A%3F")]
    public void ADelegateScope_SpellsTheDelegatesName_AsTheIndexDoes(string delegateName, string expected)
    {
        StoreScopeInfo owner = StoreScopeInfo.FromStorePrefix(Prefix)!;
        Assert.Equal(expected, MailService.DelegateScope(owner, delegateName));
    }

    [Fact]
    public void ADelegateFolderPath_UsesTheDelegatesName_NotItsUrlSpelling()
    {
        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            Prefix + "/1/Sam 50%25", "Inbox/Q1*?", includeSubfolders: false, comFolderPaths: null);

        Assert.Equal(Prefix + "/1/Sam 50%25", r.Scope);
        Assert.Equal(new[] { "/alice@example.com/Sam 50%/Q1*?" }, r.FolderPaths!);
    }

    // ------------------------------------------------------- reading a URL back

    [Fact]
    public void AHitsFolderSegmentsAndStoreName_AreTheNames_AndItsPrefixStaysSpelled()
    {
        string url = "mapi16://" + Sid + "/q99 50%25 off%2A%3Fx($5159380d)/0/Clients/50%25 off/Q1%2A%3F%5Cx/"
            + EntryIdCodecTests.SyntheticEncodedTail();

        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal("q99 50% off*?x", parsed!.StoreDisplayName);
        Assert.Equal("q99 50%25 off%2A%3Fx($5159380d)", parsed.StoreSegment);
        Assert.Equal("mapi16://" + Sid + "/q99 50%25 off%2A%3Fx($5159380d)", parsed.StorePrefix);
        Assert.Equal(new[] { "Clients", "50% off", "Q1*?\\x" }, parsed.FolderSegments);
        Assert.True(parsed.TryDecodeEntryId(out _));
    }

    [Fact]
    public void ANameWithASlash_ComesBackAsOneSegment()
    {
        // The index spells '/' inside a name as %2F, so the URL stays unambiguous; decoded, the
        // name is one segment holding a '/', which is what Outlook's folder is called.
        string url = Prefix + "/0/Parent/a%2Fb/" + EntryIdCodecTests.SyntheticEncodedTail();
        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal(new[] { "Parent", "a/b" }, parsed!.FolderSegments);
    }

    [Fact]
    public void ADelegateHitsStoreName_IsDecodedToo()
    {
        string url = Prefix + "/1/Sam 50%25/In%3Fbox/" + EntryIdCodecTests.SyntheticEncodedTail();
        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal(new[] { "Sam 50%", "In?box" }, parsed!.FolderSegments);
    }

    [Fact]
    public void AnOrdinaryUrl_ParsesExactlyAsBefore()
    {
        string url = Prefix + "/0/Inbox/Sub Folder/" + EntryIdCodecTests.SyntheticEncodedTail();
        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal("alice@example.com", parsed!.StoreDisplayName);
        Assert.Equal(new[] { "Inbox", "Sub Folder" }, parsed.FolderSegments);
        Assert.Equal(Prefix, parsed.StorePrefix);
    }

    [Fact]
    public void AHitFromAnEncodedFolder_IsLocatedByItsRealFolderNames()
    {
        IndexHit hit = Hit(Prefix + "/0/Clients/50%25 off/" + EntryIdCodecTests.SyntheticEncodedTail());

        Assert.True(HitLocator.TryMapUrlTarget(hit, out string? store, out IReadOnlyList<string>? folders));
        Assert.Equal("alice@example.com", store);
        Assert.Equal(new[] { "Clients", "50% off" }, folders);
    }

    [Fact]
    public void ASweptItem_IsRecognisedAsTheIndexHitOfTheSameMail()
    {
        // The duplicate check compares the swept item's folder NAME with the hit's last segment;
        // before the decode the two never matched, and the same mail came back twice.
        DateTime received = new(2026, 10, 3, 9, 0, 0, DateTimeKind.Utc);
        IndexHit hit = Hit(Prefix + "/0/Clients/50%25 off/" + EntryIdCodecTests.SyntheticEncodedTail(), received);
        ComMailBrief swept = new(
            entryId: "AA" + Guid.NewGuid().ToString("N"), storeDisplayName: "alice@example.com", storeId: null,
            folderName: "50% off", folderKind: null, subject: "Quarterly invoice", senderName: null, senderAddress: null,
            receivedTime: received.ToLocalTime(), isRead: false, hasAttachments: false, sizeBytes: 1000, body: null);

        Assert.True(FreshMerge.IsDuplicate(swept, hit, toleranceSeconds: 15));
    }

    [Fact]
    public void AStoreRootWithAnEncodedName_IsListedUnderItsName_AndKeepsItsSpelling()
    {
        StoreScopeInfo root = StoreScopeInfo.FromStorePrefix("mapi16://" + Sid + "/q99 50%25 off%2A%3Fx($5159380d)")!;

        Assert.Equal("q99 50% off*?x", root.StoreDisplayName);
        Assert.Equal("q99 50%25 off%2A%3Fx($5159380d)", root.StoreSegment);
        Assert.Equal(0x5159380Du, root.StoreHash);
    }

    // ------------------------------------------------------- the live proof's own guard

    [Fact]
    public void TheLiveHelper_MakesAnExactlyNamedFolder_OnlyInsideATestFolder()
    {
        string prefix = LiveOutlookTestMailer.TestFolderNamePrefix;
        string marker = "d39abcdef0123456";
        string[] parent = { prefix + "-Enc" };

        Assert.Null(LiveOutlookTestMailer.RefuseTestFolderFiling(prefix + " a/b", parent, marker));

        // A name the cleanup would not find.
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderFiling("a/b", parent, marker));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderFiling(" ", parent, marker));

        // A real folder must never gain a child: no top level, and every parent a test folder.
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderFiling(prefix + " a/b", Array.Empty<string>(), marker));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderFiling(prefix + " a/b", new[] { "Inbox" }, marker));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderFiling(prefix + " a/b", new[] { parent[0], "Inbox" }, marker));

        // A marker too weak for the S3 double match.
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderFiling(prefix + " a/b", parent, "short"));
    }

    private static IndexHit Hit(string url, DateTime? receivedUtc = null)
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = url,
            ["System.Subject"] = "Quarterly invoice",
            ["System.Message.DateReceived"] = receivedUtc ?? new DateTime(2026, 10, 3, 9, 0, 0, DateTimeKind.Utc),
        };
        return IndexRowMapper.Map(row);
    }
}
