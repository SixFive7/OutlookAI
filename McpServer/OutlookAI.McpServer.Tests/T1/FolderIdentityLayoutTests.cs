using OutlookAI.Core.Mapi;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The pure halves of the Q114 folder-identity pin (<c>T2/LiveFolderIdentityTests</c>): how a folder
/// entry id is taken apart, how an index value is searched for one, and when the test-folder rename,
/// move and delete helpers refuse. Every id here is synthetic - live ids never enter the repository.
/// </summary>
public sealed class FolderIdentityLayoutTests
{
    private static readonly byte[] ProviderUid =
    {
        0x02, 0x60, 0x47, 0xAA, 0x10, 0x20, 0x30, 0x40, 0x50, 0x60, 0x70, 0x80, 0x90, 0xA0, 0xB0, 0xC0,
    };

    private static byte[] PstFolderId(uint nid)
    {
        byte[] id = new byte[24];
        Array.Copy(ProviderUid, 0, id, 4, 16);
        BitConverter.GetBytes(nid).CopyTo(id, 20);
        return id;
    }

    private static byte[] ExchangeFolderId(ushort folderType, byte counter)
    {
        byte[] id = new byte[46];
        Array.Copy(ProviderUid, 0, id, 4, 16);
        id[20] = (byte)(folderType & 0xFF);
        id[21] = (byte)(folderType >> 8);
        for (int i = 0; i < 16; i++)
        {
            id[22 + i] = (byte)(0xD0 + i);
        }

        id[43] = counter;
        return id;
    }

    [Fact]
    public void APstFolderId_IsFlagsUidAndAFolderNid()
    {
        // NID 0x8022: index 0x401, type 0x02 (NID_TYPE_NORMAL_FOLDER) - little-endian in the id.
        FolderEntryIdLayout layout = Assert.IsType<FolderEntryIdLayout>(FolderEntryIdLayout.Parse(Convert.ToHexString(PstFolderId(0x8022))));

        Assert.Equal(FolderEntryIdKind.Pst, layout.Kind);
        Assert.Equal(24, layout.Length);
        Assert.Equal("00000000", layout.FlagsHex);
        Assert.Equal(Convert.ToHexString(ProviderUid), layout.ProviderUidHex);
        Assert.Equal("22800000", layout.FolderPartHex);
        Assert.Equal(FolderEntryIdLayout.NidTypeNormalFolder, layout.NidType);
        Assert.Null(layout.ExchangeFolderType);

        // The index spells the same node id in System.ProviderItemID: N and ten decimal digits
        // (0x8022 = 32802), measured on every row of the indexed guest (Q115).
        Assert.Equal(0x8022u, layout.Nid);
        Assert.Equal("N0000032802", layout.PstProviderItemId);
        Assert.Equal("N0002102180", FolderEntryIdLayout.ProviderItemIdOf(0x2013A4));
    }

    [Fact]
    public void AnExchangeFolderId_IsFlagsMailboxGuidTypeFidAndPad()
    {
        FolderEntryIdLayout layout = Assert.IsType<FolderEntryIdLayout>(
            FolderEntryIdLayout.Parse(Convert.ToHexString(ExchangeFolderId(FolderEntryIdLayout.ExchangePrivateFolderType, 0x2A))));

        Assert.Equal(FolderEntryIdKind.ExchangeFolder, layout.Kind);
        Assert.Equal(46, layout.Length);
        Assert.Equal(Convert.ToHexString(ProviderUid), layout.ProviderUidHex);
        Assert.Equal(FolderEntryIdLayout.ExchangePrivateFolderType, layout.ExchangeFolderType);
        Assert.Equal(44, layout.FolderPartHex!.Length); // database GUID + global counter: 22 bytes
        Assert.EndsWith("2A", layout.FolderPartHex, StringComparison.Ordinal);
        Assert.Null(layout.NidType);
        Assert.Null(layout.Nid);
        Assert.Null(layout.PstProviderItemId);
    }

    [Fact]
    public void AnIdThatFitsNoLayout_IsUnknown_AndNonHexIsNothing()
    {
        byte[] shortTerm = PstFolderId(0x8022);
        shortTerm[0] = 0x01; // a non-zero flag byte: not a long-term id
        Assert.Equal(FolderEntryIdKind.Unknown, FolderEntryIdLayout.Parse(Convert.ToHexString(shortTerm))!.Kind);

        byte[] padded = ExchangeFolderId(1, 1);
        padded[45] = 0x01; // the pad must be zero
        Assert.Equal(FolderEntryIdKind.Unknown, FolderEntryIdLayout.Parse(Convert.ToHexString(padded))!.Kind);

        Assert.Equal(FolderEntryIdKind.Unknown, FolderEntryIdLayout.Parse("0000000011")!.Kind);
        Assert.Null(FolderEntryIdLayout.Parse("not hex"));
        Assert.Null(FolderEntryIdLayout.Parse("000"));
        Assert.Null(FolderEntryIdLayout.Parse(null));
    }

    [Fact]
    public void SameBytes_IgnoresCase_ButNotLength()
    {
        string id = Convert.ToHexString(PstFolderId(0x8022));
        Assert.True(FolderEntryIdLayout.SameBytes(id, id.ToLowerInvariant()));
        Assert.False(FolderEntryIdLayout.SameBytes(id, id + "00"));
        Assert.False(FolderEntryIdLayout.SameBytes(id, Convert.ToHexString(PstFolderId(0x8042))));
        Assert.False(FolderEntryIdLayout.SameBytes(id, null));
    }

    [Fact]
    public void ValueCarries_FindsTheIdInEverySpelling_AndNowhereElse()
    {
        byte[] id = PstFolderId(0x8022);
        string hex = Convert.ToHexString(id);

        Assert.True(FolderEntryIdLayout.ValueCarries("x" + hex + "y", id));
        Assert.True(FolderEntryIdLayout.ValueCarries(hex.ToLowerInvariant(), id));
        Assert.True(FolderEntryIdLayout.ValueCarries(Convert.ToBase64String(id), id));
        Assert.True(FolderEntryIdLayout.ValueCarries("mapi16://{S-1-5-21-1}/s($1)/0/Inbox/" + EntryIdCodec.EncodeBytes(id), id));
        Assert.True(FolderEntryIdLayout.ValueCarries(new byte[] { 1, 2 }.Concat(id).ToArray(), id));
        Assert.True(FolderEntryIdLayout.ValueCarries(new object[] { "a", hex }, id));

        Assert.False(FolderEntryIdLayout.ValueCarries(null, id));
        Assert.False(FolderEntryIdLayout.ValueCarries(DateTime.UnixEpoch, id));
        Assert.False(FolderEntryIdLayout.ValueCarries("mapi16://{S-1-5-21-1}/s($1)/0/Inbox", id));
        Assert.False(FolderEntryIdLayout.ValueCarries(Convert.ToHexString(PstFolderId(0x8042)), id));

        // An item's id in the same store shares the provider UID but is NOT the folder's id.
        byte[] item = PstFolderId(0x200024);
        Assert.False(FolderEntryIdLayout.ValueCarries(EntryIdCodec.EncodeBytes(item), id));
    }

    [Fact]
    public void TestFolderChanges_AreRefusedUnlessEveryFolderInvolvedIsATestFolder()
    {
        string prefix = LiveOutlookTestMailer.TestFolderNamePrefix;

        Assert.Null(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", prefix + "-IdA-renamed", null));
        Assert.Null(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", null, prefix + "-IdB"));
        Assert.Null(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", null, null));

        // Only a test folder is changed at all.
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange("Inbox", prefix + "-x", null));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange(null, null, null));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange(" ", null, null));

        // A rename keeps the prefix, so the cleanup still finds the folder.
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", "Archive", null));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", " ", null));

        // A move goes only into another test folder: a real folder never gains a child.
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", null, "Inbox"));
        Assert.NotNull(LiveOutlookTestMailer.RefuseTestFolderChange(prefix + "-IdA", null, string.Empty));
    }
}
