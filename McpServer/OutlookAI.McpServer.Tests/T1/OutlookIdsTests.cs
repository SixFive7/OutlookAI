using OutlookAI.Core.Com;
using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The folder and store IDs every tool addresses folders and stores by (Q131 (b), Q133 (i)): the full
/// entry id re-encoded shorter in Crockford base32 with the constant bytes left out, and a check
/// character that catches a mistyped ID. Every id here is synthetic.
/// </summary>
public sealed class OutlookIdsTests
{
    private static readonly byte[] Uid =
    {
        0x02, 0x60, 0x47, 0xAA, 0x10, 0x20, 0x30, 0x40, 0x50, 0x60, 0x70, 0x80, 0x90, 0xA0, 0xB0, 0xC0,
    };

    private static byte[] PstFolder(uint nid)
    {
        byte[] id = new byte[24];
        Array.Copy(Uid, 0, id, 4, 16);
        BitConverter.GetBytes(nid).CopyTo(id, 20);
        return id;
    }

    private static byte[] ExchangeFolder(byte counter)
    {
        byte[] id = new byte[46];
        Array.Copy(Uid, 0, id, 4, 16);
        id[20] = 0x01;
        for (int i = 0; i < 16; i++)
        {
            id[22 + i] = (byte)(0xD0 + i);
        }

        id[43] = counter;
        return id;
    }

    [Fact]
    public void APstFolderId_Is34Characters_KindP_AndRoundTripsToTheWholeEntryId()
    {
        byte[] entryId = PstFolder(0x8082);
        string id = OutlookIds.FolderIdFromEntryId(entryId);

        Assert.Equal(OutlookIds.PstFolderIdLength, id.Length);
        Assert.Equal(34, id.Length);
        Assert.Equal('P', id[0]);
        Assert.All(id, c => Assert.Contains(c, OutlookIds.Alphabet));
        Assert.True(OutlookIds.TryDecodeFolderId(id, "folder_id", out string hex, out string? problem), problem);
        Assert.Equal(Convert.ToHexString(entryId), hex);
    }

    [Fact]
    public void AnExchangeFolderId_Is66Characters_KindX_AndKeepsTheFolderType()
    {
        byte[] entryId = ExchangeFolder(0x2A);
        string id = OutlookIds.FolderIdFromEntryId(entryId);

        Assert.Equal(66, id.Length);
        Assert.Equal('X', id[0]);
        Assert.True(OutlookIds.TryDecodeFolderId(id, "folder_id", out string hex, out _));
        Assert.Equal(Convert.ToHexString(entryId), hex);
    }

    [Theory]
    [InlineData(20)]
    [InlineData(31)]
    [InlineData(47)]
    public void AnyOtherEntryId_KeepsEveryByte_KindR(int length)
    {
        byte[] entryId = Enumerable.Range(1, length).Select(i => (byte)(i * 7)).ToArray();
        string id = OutlookIds.FolderIdFromEntryId(entryId);

        Assert.Equal('R', id[0]);
        Assert.True(OutlookIds.TryDecodeFolderId(id, "folder_id", out string hex, out _));
        Assert.Equal(Convert.ToHexString(entryId), hex);
    }

    [Fact]
    public void AShortRawPayload_IsNeverAnId_SoAWordThatPassesTheCheckIsStillRefused()
    {
        // Under four flag bytes and a provider UID no MAPI entry id exists; 'R' plus a short word is a name.
        string shortRaw = OutlookIds.FolderIdFromEntryId(new byte[] { 1, 2, 3, 4, 5 });
        Assert.False(OutlookIds.TryDecodeFolderId(shortRaw, "folder_id", out _, out string? problem));
        Assert.Contains("find_folder", problem, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadingBack_ForgivesCaseHyphensAndCrockfordsLookAlikes()
    {
        string id = OutlookIds.FolderIdFromEntryId(PstFolder(0x8082));
        string expected = Convert.ToHexString(PstFolder(0x8082));

        Assert.True(OutlookIds.TryDecodeFolderId(id.ToLowerInvariant(), "folder_id", out string lower, out _));
        Assert.Equal(expected, lower);
        Assert.True(OutlookIds.TryDecodeFolderId(id.Insert(10, "-").Insert(20, "-"), "folder_id", out string hyphened, out _));
        Assert.Equal(expected, hyphened);
        Assert.True(OutlookIds.TryDecodeFolderId("  " + id + "  ", "folder_id", out _, out _));

        // I and L read as 1, O as 0 (Crockford): an ID holding 1 or 0 survives either spelling.
        string withOne = Enumerable.Range(0, 4096).Select(n => OutlookIds.FolderIdFromEntryId(PstFolder(0x8002u + ((uint)n << 5))))
            .First(s => s.IndexOf('1', 1) > 0);
        int at = withOne.IndexOf('1', 1);
        Assert.True(OutlookIds.TryDecodeFolderId(withOne.Remove(at, 1).Insert(at, "l"), "folder_id", out _, out _));
        Assert.True(OutlookIds.TryDecodeFolderId(withOne.Remove(at, 1).Insert(at, "I"), "folder_id", out _, out _));
        Assert.Equal(withOne, OutlookIds.Canonical(withOne.ToLowerInvariant().Replace('1', 'l')));
    }

    [Fact]
    public void TheCheckCharacter_CatchesEverySingleWrongCharacter()
    {
        string id = OutlookIds.FolderIdFromEntryId(ExchangeFolder(0x2A));
        for (int position = 0; position < id.Length; position++)
        {
            foreach (char replacement in OutlookIds.Alphabet)
            {
                if (replacement == id[position])
                {
                    continue;
                }

                string typo = id.Substring(0, position) + replacement + id.Substring(position + 1);
                Assert.False(
                    OutlookIds.TryDecodeFolderId(typo, "folder_id", out _, out _),
                    "a wrong '" + replacement + "' at " + position + " was read as a folder ID");
            }
        }
    }

    [Fact]
    public void TheCheckCharacter_CatchesEverySwapOfTwoNeighbours_AndOfTwoCharactersUpTo30Apart()
    {
        string id = OutlookIds.FolderIdFromEntryId(ExchangeFolder(0x2A));
        for (int i = 0; i < id.Length; i++)
        {
            for (int distance = 1; distance <= 30 && i + distance < id.Length; distance++)
            {
                int j = i + distance;
                if (id[i] == id[j])
                {
                    continue;
                }

                char[] swapped = id.ToCharArray();
                (swapped[i], swapped[j]) = (swapped[j], swapped[i]);
                Assert.False(
                    OutlookIds.TryDecodeFolderId(new string(swapped), "folder_id", out _, out _),
                    "swapping " + i + " and " + j + " went unnoticed");
            }
        }
    }

    [Fact]
    public void ALostOrAddedCharacter_IsRefused()
    {
        string id = OutlookIds.FolderIdFromEntryId(PstFolder(0x8082));
        Assert.False(OutlookIds.TryDecodeFolderId(id.Substring(0, id.Length - 1), "folder_id", out _, out _));
        Assert.False(OutlookIds.TryDecodeFolderId(id.Remove(5, 1), "folder_id", out _, out _));
        Assert.False(OutlookIds.TryDecodeFolderId(id.Insert(5, "0"), "folder_id", out _, out _));
    }

    [Fact]
    public void AStoreId_Is28Characters_KindS_AndEveryFolderIdNamesItsStore()
    {
        string root = Convert.ToHexString(PstFolder(0x8022));
        string storeId = OutlookIds.StoreIdOf(root, "0000ABCD")!;

        Assert.Equal(28, storeId.Length);
        Assert.Equal('S', storeId[0]);
        Assert.Equal(storeId, OutlookIds.StoreIdOfEntryId(Convert.ToHexString(PstFolder(0x8082))));
        Assert.Equal(storeId, OutlookIds.StoreIdOfEntryId(Convert.ToHexString(ExchangeFolder(1))));
        Assert.True(OutlookIds.TryDecodeStoreId(storeId, "store_id", out StoreIdParts? parts, out _));
        Assert.Equal(Convert.ToHexString(Uid), parts!.ProviderUidHex);
        Assert.Null(parts.StoreEntryIdHex);
    }

    [Fact]
    public void AStoreWhoseRootWasNotRead_IsAddressedByItsOwnEntryId_KindT()
    {
        string storeId = OutlookIds.StoreIdOf(null, "0000000038A1BB1005E5101AA1BB08002B2A56C2")!;
        Assert.Equal('T', storeId[0]);
        Assert.True(OutlookIds.TryDecodeStoreId(storeId, "store_id", out StoreIdParts? parts, out _));
        Assert.Equal("0000000038A1BB1005E5101AA1BB08002B2A56C2", parts!.StoreEntryIdHex);
        Assert.Null(OutlookIds.StoreIdOf(null, "not hex"));
    }

    /// <summary>
    /// Every store's top folder is read for its S ID (decided 2026-11-18, option (b)) - but for an Exchange store
    /// Outlook does not cache while Outlook says it is offline or disconnected, whose top folder is a server
    /// round trip with no server to answer it: that store keeps its T ID until Outlook is connected again.
    /// <c>ExchangeStoreType</c> 1 is an online entry, 3 not Exchange; <c>ExchangeConnectionMode</c> 100 to 400 is
    /// offline or disconnected, 500 to 800 connected.
    /// </summary>
    [Theory]
    [InlineData(1, false, 700, true)]
    [InlineData(1, false, 800, true)]
    [InlineData(1, false, 500, true)]
    [InlineData(1, false, null, true)]
    [InlineData(1, false, 400, false)]
    [InlineData(1, false, 300, false)]
    [InlineData(1, false, 200, false)]
    [InlineData(1, false, 100, false)]
    [InlineData(0, false, 300, false)]
    [InlineData(1, null, 200, false)]
    [InlineData(4, true, 200, true)]
    [InlineData(0, true, 400, true)]
    [InlineData(3, null, 200, true)]
    [InlineData(null, null, 200, true)]
    public void AnExchangeStoreOutlookDoesNotCache_IsNotAskedForItsTopFolder_WhileOutlookSaysItIsOffline(
        int? exchangeStoreType, bool? isCachedExchange, int? connectionMode, bool readsTopFolder)
    {
        Assert.Equal(readsTopFolder, OutlookComSession.ReadsTopFolder(exchangeStoreType, isCachedExchange, connectionMode));
    }

    [Fact]
    public void APstFolderId_NamesItsIndexRow_ByTheNodeIdInTenDigits()
    {
        Assert.Equal("N0000032898", OutlookIds.PstProviderItemId(PstFolder(0x8082)));
        Assert.Null(OutlookIds.PstProviderItemId(ExchangeFolder(1)));
        Assert.Equal(FolderIdLayout.Pst, OutlookIds.Classify(PstFolder(1)));
        Assert.Equal(FolderIdLayout.Exchange, OutlookIds.Classify(ExchangeFolder(1)));
        byte[] padded = ExchangeFolder(1);
        padded[45] = 1;
        Assert.Equal(FolderIdLayout.Raw, OutlookIds.Classify(padded));
    }

    [Theory]
    [InlineData("Inbox", "\"Inbox\"")]
    [InlineData("Projects/2026", "\"2026\"")]
    [InlineData("\\\\tier@vm.invalid\\Inbox\\Sub", "\"Sub\"")]
    [InlineData("Inbox/Fun ", "\"Fun\"")]
    public void ANameOrPath_IsRefused_WithThePointerToFindFolder_AndTheNameToAskFor(string input, string leaf)
    {
        Assert.False(OutlookIds.TryDecodeFolderId(input, "folder_id", out _, out string? problem));
        Assert.Contains("find_folder", problem, StringComparison.Ordinal);
        Assert.Contains(leaf, problem, StringComparison.Ordinal);
        Assert.Contains("'folder_id'", problem, StringComparison.Ordinal);
    }

    [Fact]
    public void EachWrongKindOfValue_GetsItsOwnSentence()
    {
        string folderId = OutlookIds.FolderIdFromEntryId(PstFolder(0x8082));
        string storeId = OutlookIds.StoreIdOfEntryId(Convert.ToHexString(PstFolder(0x8082)))!;
        string typo = folderId.Substring(0, 5) + (folderId[5] == 'A' ? 'B' : 'A') + folderId.Substring(6);

        Assert.Contains("STORE ID", OutlookIds.DescribeFolderIdProblem("folder_id", storeId), StringComparison.Ordinal);
        Assert.Contains("FOLDER ID", OutlookIds.DescribeStoreIdProblem("store_id", folderId), StringComparison.Ordinal);
        Assert.Contains("check character", OutlookIds.DescribeFolderIdProblem("folder_id", typo), StringComparison.Ordinal);
        Assert.Contains(
            "raw Outlook entry id", OutlookIds.DescribeFolderIdProblem("folder_id", Convert.ToHexString(PstFolder(0x8082))),
            StringComparison.Ordinal);
        Assert.Contains("is empty", OutlookIds.DescribeFolderIdProblem("folder_id", "  "), StringComparison.Ordinal);
        Assert.Contains("list_accounts", OutlookIds.DescribeStoreIdProblem("store_id", "alice@example.com"), StringComparison.Ordinal);
        Assert.False(OutlookIds.TryDecodeStoreId(folderId, "store_id", out _, out _));
        Assert.False(OutlookIds.TryDecodeFolderId(storeId, "folder_id", out _, out _));
    }
}
