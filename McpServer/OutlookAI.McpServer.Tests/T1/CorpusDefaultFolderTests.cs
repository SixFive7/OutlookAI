using OutlookAI.Core.Com;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="CorpusDefaultFolders"/> - how the corpus tool finds a store's default folders
/// without ever calling the lookup that CREATES a missing one on a PST (Q84, decision (c),
/// 2026-09-24) - and <see cref="CorpusFolderVisibility"/>, which refuses a folder nobody can see.
/// <para>
/// Both halves answer a defect measured on OAI-UNINDEXED the same day: <c>GetDefaultFolder</c>
/// created Drafts and Junk Email in the bystander PST during its build, and for that PST's Inbox it
/// returned the nameless non-IPM ROOT, where 172 items were then built out of sight.
/// </para>
/// <para>
/// Driven through a fake <see cref="ISpecialFolderStore"/> that records every
/// <c>GetDefaultFolder</c> call: on a PST the only ones allowed are for a folder
/// <c>PR_VALID_FOLDER_MASK</c> has proven present - <see cref="SpecialFolders"/>' own rule, which this
/// resolver goes through for every mail folder. No COM, no Outlook.
/// </para>
/// </summary>
public sealed class CorpusDefaultFolderTests
{
    private const int Calendar = CorpusItemKinds.CalendarFolderId;
    private const int Contacts = CorpusItemKinds.ContactsFolderId;
    private const int Tasks = CorpusItemKinds.TasksFolderId;
    private const int Drafts = CorpusItemKinds.DraftsFolderId;

    [Fact]
    public void TheDesignations_AreTheMsOxosfldPropertiesOfEachFolder()
    {
        // PidTagIpmAppointmentEntryId 0x36D0, PidTagIpmContactEntryId 0x36D1, PidTagIpmTaskEntryId
        // 0x36D4, each PT_BINARY (0x0102). A wrong tag reads "absent" for a Calendar that exists,
        // and the build would then file the population in a stand-in beside it.
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x36D00102", CorpusDefaultFolders.DesignationSchema(Calendar));
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x36D10102", CorpusDefaultFolders.DesignationSchema(Contacts));
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x36D40102", CorpusDefaultFolders.DesignationSchema(Tasks));

        // Every mail folder is SpecialFolders' own, and so has no designation here.
        foreach (int mailFolder in new[] { Drafts, SpecialFolders.OlFolderInbox, SpecialFolders.OlFolderSentMail,
            SpecialFolders.OlFolderDeletedItems, SpecialFolders.OlFolderOutbox, SpecialFolders.OlFolderJunk })
        {
            Assert.Null(CorpusDefaultFolders.DesignationSchema(mailFolder));
        }
    }

    [Theory]
    [InlineData(Calendar)]
    [InlineData(Contacts)]
    [InlineData(Tasks)]
    public void APstWhoseInboxDesignatesTheFolder_ResolvesIt_WithoutEverAskingForItByKind(int folderId)
    {
        FakeStore pst = FakeStore.PstWithInbox();
        object designated = pst.AddFolder(Bytes(folderId));
        pst.InboxProperties[CorpusDefaultFolders.DesignationSchema(folderId)!] = PropertyRead.Found(Bytes(folderId));

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, CorpusDefaultFolders.Resolve(pst, folderId, out object? folder));
        Assert.Same(designated, folder);
        Assert.Equal(new[] { SpecialFolders.OlFolderInbox }, pst.GetDefaultFolderCalls);
        Assert.Contains(pst.Inbox, pst.Released);
    }

    [Fact]
    public void AStoreWithNoInboxAndNoDesignation_IsAbsent_AndNothingIsAskedFor()
    {
        // The bystander's shape: a PST attached with AddStoreEx, whose mask proves no Inbox. The
        // build files the kind in a stand-in; GetDefaultFolder(9) would have MADE a Calendar.
        FakeStore pst = FakeStore.PstWithInbox();
        pst.ValidMask = SpecialFolders.FolderIpmWastebasketValid;

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(pst, Calendar, out object? folder));
        Assert.Null(folder);
        Assert.Empty(pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void AnInboxWithoutTheDesignation_IsAbsent_UnlessTheStoreObjectCarriesIt()
    {
        FakeStore pst = FakeStore.PstWithInbox();
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(pst, Contacts, out _));

        object designated = pst.AddFolder(Bytes(Contacts));
        pst.StoreProperties[CorpusDefaultFolders.ContactEntryIdSchema] = PropertyRead.Found(Bytes(Contacts));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, CorpusDefaultFolders.Resolve(pst, Contacts, out object? folder));
        Assert.Same(designated, folder);
        Assert.DoesNotContain(Contacts, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void ADesignationThatWillNotRead_IsUnreadable_NeverAbsent()
    {
        // Reading a failure as absence would let a census skip a folder whose items it should
        // count, and let a build put a stand-in beside a Calendar that is there.
        FakeStore pst = FakeStore.PstWithInbox();
        pst.InboxProperties[CorpusDefaultFolders.TaskEntryIdSchema] = PropertyRead.Failure();

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Unreadable, CorpusDefaultFolders.Resolve(pst, Tasks, out object? folder));
        Assert.Null(folder);
        Assert.DoesNotContain(Tasks, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void AMaskThatWillNotRead_IsUnreadable_AndTheInboxIsNotAskedForEither()
    {
        FakeStore pst = FakeStore.PstWithInbox();
        pst.ValidMask = null;

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Unreadable, CorpusDefaultFolders.Resolve(pst, Calendar, out _));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Unreadable, CorpusDefaultFolders.Resolve(pst, SpecialFolders.OlFolderInbox, out _));
        Assert.Empty(pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void ADesignatedFolderThatIsGone_IsAbsent_AndOneThatWillNotOpen_IsUnreadable()
    {
        FakeStore pst = FakeStore.PstWithInbox();
        pst.InboxProperties[CorpusDefaultFolders.AppointmentEntryIdSchema] = PropertyRead.Found(Bytes(Calendar));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(pst, Calendar, out _));

        pst.OpenFails = true;
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Unreadable, CorpusDefaultFolders.Resolve(pst, Calendar, out _));
        Assert.DoesNotContain(Calendar, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void Drafts_GoesThroughSpecialFolders_AndIsNotAskedForByKindEither()
    {
        FakeStore pst = FakeStore.PstWithInbox();
        object drafts = pst.AddFolder(Bytes(Drafts));
        pst.InboxProperties[SpecialFolders.DraftsEntryIdSchema] = PropertyRead.Found(Bytes(Drafts));

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, CorpusDefaultFolders.Resolve(pst, Drafts, out object? folder));
        Assert.Same(drafts, folder);
        Assert.DoesNotContain(Drafts, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void EveryMailFolder_IsSpecialFoldersAnswer_SoAPstIsNeverAskedForOneItLacks()
    {
        // The corpus scan, the probes and the build all come through here now - Inbox, Sent Items,
        // Deleted Items, Outbox, Drafts and Junk Email alike. On the bystander's shape (only Deleted
        // Items proven by the mask) the Inbox and Sent Items are ABSENT without a single creating call,
        // Deleted Items is asked for because the mask proves it, and Junk Email - identified through
        // the Inbox's PR_ADDITIONAL_REN_ENTRYIDS, which a store without an Inbox has none of - is absent.
        FakeStore bystander = FakeStore.PstWithInbox();
        bystander.ValidMask = SpecialFolders.FolderIpmWastebasketValid;

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(bystander, SpecialFolders.OlFolderInbox, out _));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(bystander, SpecialFolders.OlFolderSentMail, out _));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(bystander, SpecialFolders.OlFolderOutbox, out _));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, CorpusDefaultFolders.Resolve(bystander, SpecialFolders.OlFolderJunk, out _));
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, CorpusDefaultFolders.Resolve(bystander, SpecialFolders.OlFolderDeletedItems, out object? deleted));
        Assert.NotNull(deleted);
        Assert.Equal(new[] { SpecialFolders.OlFolderDeletedItems }, bystander.GetDefaultFolderCalls);

        // A store whose mask proves the Inbox is asked for exactly that.
        FakeStore hub = FakeStore.PstWithInbox();
        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, CorpusDefaultFolders.Resolve(hub, SpecialFolders.OlFolderInbox, out object? inbox));
        Assert.Same(hub.Inbox, inbox);
        Assert.Equal(new[] { SpecialFolders.OlFolderInbox }, hub.GetDefaultFolderCalls);
    }

    [Fact]
    public void AnExchangeMailbox_IsResolvedAsSpecialFoldersResolvesIt()
    {
        // Every folder of an Exchange mailbox is a server default folder; there the plain call is
        // the right one, exactly as in the product.
        FakeStore exchange = FakeStore.PstWithInbox();
        exchange.ExchangeStoreType = 0;

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, CorpusDefaultFolders.Resolve(exchange, Calendar, out object? folder));
        Assert.NotNull(folder);
        Assert.Equal(new[] { Calendar }, exchange.GetDefaultFolderCalls);
    }

    [Fact]
    public void TheStandInForAnyFolder_IsNamedSoReindexAttributesItToThatFolder()
    {
        // The build files a folder the store lacks - or has only invisibly - in the root child
        // StandInName(id), and a reindex with no manifest reads the id back off the name: both halves
        // must agree for every folder a population or its probes can need a stand-in for.
        foreach (int folderId in new[] { SpecialFolders.OlFolderInbox, SpecialFolders.OlFolderSentMail, Calendar, Contacts, Tasks, Drafts })
        {
            string name = CorpusFolderIds.StandInName(folderId);
            Assert.StartsWith(CorpusManifest.CreatedFolderPrefix + "-", name, StringComparison.Ordinal);
            Assert.Equal(folderId, CorpusFolderIds.ForCreatedFolder(null, name, null));
        }
    }

    // ================================================================ visibility

    private const string Root = "00000000ROOT";

    [Fact]
    public void AFolderUnderTheRootFolder_IsVisible_AtAnyDepth()
    {
        Assert.True(CorpusFolderVisibility.IsVisible("Inbox", new[] { Root }, Root));
        Assert.True(CorpusFolderVisibility.IsVisible("Projects", new[] { "INBOX", Root }, Root));
        Assert.True(CorpusFolderVisibility.IsVisible("Deep", new[] { "C", "B", "A", Root.ToLowerInvariant() }, Root));
    }

    [Fact]
    public void WhatGetDefaultFolderReturnedForTheBystandersInbox_IsNotVisible()
    {
        // OAI-UNINDEXED, 2026-09-24: the PST's non-IPM ROOT - display name empty, and its parent is
        // not a folder at all, so the walk ends before it reaches the root folder Outlook draws.
        Assert.False(CorpusFolderVisibility.IsVisible(string.Empty, new string?[] { null }, Root));
        Assert.False(CorpusFolderVisibility.IsVisible(null, Array.Empty<string?>(), Root));

        // A named folder hung off that invisible root - where the bystander's two "Inbox" subfolders
        // were created - is no more visible for having a name.
        Assert.False(CorpusFolderVisibility.IsVisible("Projects", new[] { "NONIPMROOT", null }, Root));
        Assert.False(CorpusFolderVisibility.IsVisible("Projects", new[] { "NONIPMROOT" }, Root));
    }

    [Fact]
    public void AFolderIsNotVisible_WhenTheRootCannotBeRead_OrItsNameIsBlank()
    {
        Assert.False(CorpusFolderVisibility.IsVisible("Inbox", new[] { Root }, null));
        Assert.False(CorpusFolderVisibility.IsVisible("Inbox", new[] { Root }, string.Empty));
        Assert.False(CorpusFolderVisibility.IsVisible("   ", new[] { Root }, Root));

        // A chain broken before the root is not a chain that reached it.
        Assert.False(CorpusFolderVisibility.IsVisible("Inbox", new string?[] { "A", null, Root }, Root));
        Assert.Throws<ArgumentNullException>(() => CorpusFolderVisibility.IsVisible("Inbox", null!, Root));
    }

    private static byte[] Bytes(int seed) => new byte[] { 0, 0, 0, 0, (byte)seed, 0x42, 0x42, 0x42 };

    /// <summary>A store that answers from dictionaries and records what it was asked.</summary>
    private sealed class FakeStore : ISpecialFolderStore
    {
        private readonly Dictionary<string, object> _foldersByHex = new(StringComparer.OrdinalIgnoreCase);

        public object Inbox { get; } = new object();

        public int? ValidMask { get; set; }

        public bool OpenFails { get; set; }

        public Dictionary<string, PropertyRead> InboxProperties { get; } = new(StringComparer.Ordinal);

        public Dictionary<string, PropertyRead> StoreProperties { get; } = new(StringComparer.Ordinal);

        public List<int> GetDefaultFolderCalls { get; } = new();

        public List<object?> Released { get; } = new();

        public int? ExchangeStoreType { get; set; } = SpecialFolders.OlNotExchange;

        public static FakeStore PstWithInbox() => new()
        {
            ValidMask = SpecialFolders.FolderIpmInboxValid | SpecialFolders.FolderIpmOutboxValid
                | SpecialFolders.FolderIpmWastebasketValid | SpecialFolders.FolderIpmSentmailValid,
        };

        public object AddFolder(byte[] entryId)
        {
            object folder = new object();
            _foldersByHex[Convert.ToHexString(entryId)] = folder;
            return folder;
        }

        public PropertyRead ReadStoreProperty(string schemaName)
        {
            if (schemaName == SpecialFolders.ValidFolderMaskSchema)
            {
                return ValidMask == null ? PropertyRead.Failure() : PropertyRead.Found(ValidMask.Value);
            }

            return StoreProperties.TryGetValue(schemaName, out PropertyRead read) ? read : PropertyRead.Missing();
        }

        public PropertyRead ReadFolderProperty(object folder, string schemaName)
            => ReferenceEquals(folder, Inbox) && InboxProperties.TryGetValue(schemaName, out PropertyRead read)
                ? read
                : PropertyRead.Missing();

        public object? GetDefaultFolder(int olDefaultFolderId)
        {
            GetDefaultFolderCalls.Add(olDefaultFolderId);
            return olDefaultFolderId == SpecialFolders.OlFolderInbox ? Inbox : new object();
        }

        public PropertyReadStatus OpenFolder(string entryIdHex, out object? folder)
        {
            folder = null;
            if (OpenFails)
            {
                return PropertyReadStatus.Failed;
            }

            if (_foldersByHex.TryGetValue(entryIdHex, out object? found))
            {
                folder = found;
                return PropertyReadStatus.Found;
            }

            return PropertyReadStatus.NotFound;
        }

        public IReadOnlyList<string>? ListRootChildEntryIds() => Array.Empty<string>();

        public string? EntryIdOf(object folder) => null;

        public SpecialFolderFacts Describe(object folder) => new(null, null, null, null, null);

        public void Release(object? comObject) => Released.Add(comObject);
    }
}
