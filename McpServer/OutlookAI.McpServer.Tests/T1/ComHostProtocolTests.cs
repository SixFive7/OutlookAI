using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OutlookAI.ComHost.Host;
using OutlookAI.ComHost.Protocol;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the parent/child wire format.
/// <para>
/// Framing has one job that matters above the rest: it must be impossible to
/// desynchronise. A desync would leave the parent waiting on a response that never
/// parses - which presents exactly like the silent hang this whole architecture exists
/// to remove, one layer further down and harder to see.
/// </para>
/// </summary>
public sealed class ComHostProtocolTests
{
    [Fact]
    public void PerStoreSweepWindows_SurviveTheWire_WithTheirStoreNamesIntact()
    {
        // The freshness sweep now carries one window PER STORE, keyed on the store display
        // name, and that map crosses this pipe on every unscoped search. Two ways it could
        // fail silently and produce a sweep that simply looks narrow: a key policy mangling
        // the names (web defaults camel-case PROPERTY names - dictionary keys must pass
        // through verbatim, and a store really can be called "Archive 2019.pst"), and a
        // DateTime losing its UTC kind on the way, which would shift every window by the
        // machine's offset. Neither shows up anywhere in the payload.
        Dictionary<string, DateTime> windows = new(StringComparer.Ordinal)
        {
            ["Archive 2019.pst"] = new DateTime(2026, 08, 18, 07, 20, 09, DateTimeKind.Utc),
            ["Jori Huisman"] = new DateTime(2026, 08, 17, 22, 00, 00, DateTimeKind.Utc),
        };

        JsonElement wire = JsonSerializer.SerializeToElement(
            new { perStoreSinceUtc = windows }, ComHostProtocol.Json);
        IReadOnlyDictionary<string, DateTime>? back = wire
            .GetProperty("perStoreSinceUtc")
            .Deserialize<IReadOnlyDictionary<string, DateTime>>(ComHostProtocol.Json);

        Assert.NotNull(back);
        Assert.Equal(2, back!.Count);
        foreach (KeyValuePair<string, DateTime> expected in windows)
        {
            Assert.True(back.TryGetValue(expected.Key, out DateTime since), $"key '{expected.Key}' did not survive");
            Assert.Equal(expected.Value, since.ToUniversalTime());
        }

        // And the far side re-keys case-insensitively, because JSON gives it an ordinal map
        // while every store comparison in this server is case-insensitive.
        IReadOnlyDictionary<string, DateTime> normalized = OutlookComSession.NormalizeSweepWindows(back);
        Assert.Equal(
            windows["Archive 2019.pst"],
            OutlookComSession.WindowFor(normalized, "ARCHIVE 2019.PST", DateTime.MaxValue));
    }

    [Fact]
    public async Task Frame_RoundTrips()
    {
        ComHostRequest sent = new ComHostRequest
        {
            Id = 42,
            Operation = "TryReadItem",
            Arguments = JsonSerializer.SerializeToElement(new { entryIdHex = "ABC", storeId = (string?)null }, ComHostProtocol.Json),
        };

        using MemoryStream stream = new MemoryStream(ComHostProtocol.EncodeFrame(sent));
        ComHostRequest? read = await ComHostProtocol.ReadFrameAsync<ComHostRequest>(stream, CancellationToken.None);

        Assert.NotNull(read);
        Assert.Equal(42, read!.Id);
        Assert.Equal("TryReadItem", read.Operation);
        Assert.Equal("ABC", read.Arguments!.Value.GetProperty("entryIdHex").GetString());
    }

    [Fact]
    public async Task Frame_SurvivesPayloadsFullOfNewlines()
    {
        // The reason framing is length-prefixed rather than newline-delimited: real
        // payloads carry mail bodies, and those are full of newlines.
        string body = "line one\nline two\r\nline three\n\n{\"looks\":\"like json\"}\n";
        ComHostResponse sent = new ComHostResponse
        {
            Id = 7,
            Ok = true,
            Result = JsonSerializer.SerializeToElement(body, ComHostProtocol.Json),
        };

        using MemoryStream stream = new MemoryStream(ComHostProtocol.EncodeFrame(sent));
        ComHostResponse? read = await ComHostProtocol.ReadFrameAsync<ComHostResponse>(stream, CancellationToken.None);

        Assert.NotNull(read);
        Assert.Equal(body, read!.Result!.Value.GetString());
    }

    [Fact]
    public async Task TwoFramesBackToBack_ReadIndependently()
    {
        byte[] first = ComHostProtocol.EncodeFrame(new ComHostRequest { Id = 1, Operation = "A" });
        byte[] second = ComHostProtocol.EncodeFrame(new ComHostRequest { Id = 2, Operation = "B" });

        using MemoryStream stream = new MemoryStream([.. first, .. second]);
        ComHostRequest? a = await ComHostProtocol.ReadFrameAsync<ComHostRequest>(stream, CancellationToken.None);
        ComHostRequest? b = await ComHostProtocol.ReadFrameAsync<ComHostRequest>(stream, CancellationToken.None);

        Assert.Equal("A", a!.Operation);
        Assert.Equal("B", b!.Operation);
    }

    [Fact]
    public async Task CleanEndOfStream_ReturnsNullRatherThanBlocking()
    {
        // "Peer gone" must be distinguishable from "keep waiting". Getting this wrong is
        // how a dead child becomes an indefinite wait.
        using MemoryStream empty = new MemoryStream([]);
        Assert.Null(await ComHostProtocol.ReadFrameAsync<ComHostResponse>(empty, CancellationToken.None));
    }

    [Fact]
    public async Task TruncatedFrame_ReportsPeerGoneRatherThanHanging()
    {
        byte[] full = ComHostProtocol.EncodeFrame(new ComHostRequest { Id = 1, Operation = "Truncated" });
        using MemoryStream partial = new MemoryStream(full[..(full.Length / 2)]);

        Assert.Null(await ComHostProtocol.ReadFrameAsync<ComHostRequest>(partial, CancellationToken.None));
    }

    [Fact]
    public async Task ImplausibleDeclaredLength_IsRejectedInsteadOfAllocating()
    {
        byte[] header = BitConverter.GetBytes((uint)(ComHostProtocol.MaxFrameBytes + 1));
        using MemoryStream stream = new MemoryStream(header);

        ComHostProtocolException ex = await Assert.ThrowsAsync<ComHostProtocolException>(
            () => ComHostProtocol.ReadFrameAsync<ComHostRequest>(stream, CancellationToken.None));
        Assert.Contains("desynchronised", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ZeroLengthFrame_IsTreatedAsDesync()
    {
        using MemoryStream stream = new MemoryStream(BitConverter.GetBytes(0u));

        await Assert.ThrowsAsync<ComHostProtocolException>(
            () => ComHostProtocol.ReadFrameAsync<ComHostRequest>(stream, CancellationToken.None));
    }

    [Fact]
    public void EncodedFrame_IsLengthPrefixedUtf8Json()
    {
        byte[] frame = ComHostProtocol.EncodeFrame(new ComHostRequest { Id = 3, Operation = "Ping" });

        uint declared = BitConverter.ToUInt32(frame, 0);
        Assert.Equal((uint)(frame.Length - 4), declared);

        string json = Encoding.UTF8.GetString(frame, 4, frame.Length - 4);
        Assert.Contains("\"operation\":\"Ping\"", json, StringComparison.Ordinal);
    }

    // --------------------------------------------------------------- frame size meter

    [Fact]
    public void EncodingAFrame_RecordsHowBigItWas()
    {
        // Nobody had ever measured the largest frame this product actually produces, so
        // "64 MB is far above any real payload" was an assumption. This is the instrument
        // that turns it into evidence: outlook_health reports the mark and the limit
        // together, and the ratio between them is the headroom.
        string body = new string('m', 96 * 1024);
        byte[] frame = ComHostProtocol.EncodeFrame(new ComHostResponse
        {
            Id = 1,
            Ok = true,
            Result = JsonSerializer.SerializeToElement(body, ComHostProtocol.Json),
        });

        Assert.True(
            ComHostFrameMeter.Shared.LargestFrameBytes >= frame.Length - 4,
            "the meter must have seen a frame at least as big as the one just encoded");
    }

    [Fact]
    public void TheFrameMeter_IsAHighWaterMark_NotTheLastValue()
    {
        // The whole point is the LARGEST payload seen. A meter that tracked the most recent
        // one would read near zero at almost any moment it was asked, because the frame
        // before the health call is a health call.
        _ = ComHostProtocol.EncodeFrame(new ComHostResponse
        {
            Id = 2,
            Ok = true,
            Result = JsonSerializer.SerializeToElement(new string('h', 48 * 1024), ComHostProtocol.Json),
        });
        long peak = ComHostFrameMeter.Shared.LargestFrameBytes;

        _ = ComHostProtocol.EncodeFrame(new ComHostRequest { Id = 3, Operation = "Ping" });

        Assert.Equal(peak, ComHostFrameMeter.Shared.LargestFrameBytes);
    }

    [Fact]
    public void ARefusedFrame_DoesNotMoveTheHighWaterMark()
    {
        // A refused payload never crossed the pipe, so counting it as a frame would report a
        // size nothing ever carried and make every later reading look closer to the ceiling
        // than it was. It is counted separately instead, and its size is named in the error
        // the caller receives.
        long peak = ComHostFrameMeter.Shared.LargestFrameBytes;
        int refusals = ComHostFrameMeter.Shared.FramesRefusedTooLarge;

        _ = Assert.Throws<ComHostProtocolException>(() => ComHostProtocol.EncodeFrame(
            new ComHostResponse
            {
                Id = 4,
                Ok = true,
                Result = JsonSerializer.SerializeToElement(new string('r', 256 * 1024), ComHostProtocol.Json),
            },
            maxFrameBytes: 1024));

        Assert.Equal(peak, ComHostFrameMeter.Shared.LargestFrameBytes);
        Assert.Equal(refusals + 1, ComHostFrameMeter.Shared.FramesRefusedTooLarge);
    }

    [Fact]
    public async Task ReadingAFrame_RecordsItToo_BecauseTheBigOnesAreBuiltInTheOtherProcess()
    {
        // The answers that get large are encoded in the CHILD, whose counters die with it -
        // and it is restarted precisely on the degraded profiles that produce them. The
        // parent's only view of their size is the frame it reads back, so the read path has
        // to measure as well or health would only ever report the size of the REQUESTS this
        // process sends, which are tiny.
        byte[] frame = ComHostProtocol.EncodeFrame(
            new ComHostResponse
            {
                Id = 5,
                Ok = true,
                Result = JsonSerializer.SerializeToElement(new string('w', 192 * 1024), ComHostProtocol.Json),
            });

        using MemoryStream stream = new MemoryStream(frame);
        _ = await ComHostProtocol.ReadFrameAsync<ComHostResponse>(stream, CancellationToken.None);

        Assert.True(ComHostFrameMeter.Shared.LargestFrameBytes >= frame.Length - 4);
    }

    // ------------------------------------------------------- ComDraftBody round-trip

    [Fact]
    public void ComDraftBody_PlainText_RoundTrips()
    {
        // ComDraftBody has a private constructor and static factories, so the serializer
        // cannot build it without the hand-written converter. If that converter regresses,
        // every draft tool silently loses its body across the pipe.
        ComDraftBody original = ComDraftBody.FromText("Hello\nthere");

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComDraftBody? read = JsonSerializer.Deserialize<ComDraftBody>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.False(read!.IsHtml);
        Assert.Equal("Hello\nthere", read.Text);
        Assert.Equal(string.Empty, read.Html);
        Assert.Equal("text", read.FormatName);
    }

    [Fact]
    public void ComDraftBody_Html_RoundTrips()
    {
        ComDraftBody original = ComDraftBody.FromHtml("<p>Hi</p>");

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComDraftBody? read = JsonSerializer.Deserialize<ComDraftBody>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.True(read!.IsHtml);
        Assert.Equal("<p>Hi</p>", read.Html);
        Assert.Equal(string.Empty, read.Text);
        Assert.Equal("html", read.FormatName);
    }

    [Fact]
    public void ComDraftBody_DoesNotSerializeItsComputedFormatName()
    {
        // FormatName is derived from IsHtml. Writing it would round-trip a phantom
        // property that the factories would then have to ignore.
        string json = JsonSerializer.Serialize(ComDraftBody.FromText("x"), ComHostProtocol.Json);

        Assert.DoesNotContain("formatName", json, StringComparison.OrdinalIgnoreCase);
    }

    // -------------------------------------------------- sweep counters across the pipe

    [Fact]
    public void ComSweepResult_CarriesItsPerStoreCountersAcrossTheWire()
    {
        // The sweep runs in the CHILD and its counters decide degraded/freshness in the
        // PARENT, so per-store attribution only exists if it survives this hop. Nothing
        // hand-writes that: the proxy reflects over IOutlookSession and serializes whatever
        // the contract returns, so a shape the serializer cannot rebuild fails silently -
        // the counters would come back zeroed and every store-scoped search would report
        // itself as having swept nothing.
        ComSweepResult original = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 7,
            foldersSkipped: 1,
            sweptFolders: new[] { "alice@example.com/Inbox", "bob@example.com/Inbox" },
            foldersFailed: 1,
            foldersAbsent: 1,
            perStore: new[]
            {
                new ComStoreSweepCounters("alice@example.com", 4, 0, 0, 0),
                new ComStoreSweepCounters("bob@example.com", 3, 1, 1, 1),
            });

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComSweepResult? read = JsonSerializer.Deserialize<ComSweepResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(2, read!.PerStore.Count);
        Assert.Equal(original.FoldersSwept, read.FoldersSwept);
        Assert.Equal(original.FoldersFailed, read.FoldersFailed);

        ComStoreSweepCounters bob = read.PerStore[1];
        Assert.Equal("bob@example.com", bob.StoreDisplayName);
        Assert.Equal(3, bob.FoldersSwept);
        Assert.Equal(1, bob.FoldersSkipped);
        Assert.Equal(1, bob.FoldersFailed);
        Assert.Equal(1, bob.FoldersAbsent);
    }

    [Fact]
    public void ComSweepResult_CarriesItsUnreadableRowCountAcrossTheWire()
    {
        // Same hop, same reason, one level finer: rows lost INSIDE a folder that was swept
        // (gap H1) are counted in the child and decide degraded/freshness in the parent. A
        // counter that does not survive the pipe reads as zero, which is exactly the silence
        // it was added to remove.
        ComSweepResult original = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            perStore: new[]
            {
                new ComStoreSweepCounters("alice@example.com", 4, 0, 0, 0, rowsUnreadable: 6),
            },
            rowsUnreadable: 6);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComSweepResult? read = JsonSerializer.Deserialize<ComSweepResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(6, read!.RowsUnreadable);
        Assert.Equal(6, Assert.Single(read.PerStore).RowsUnreadable);
    }

    [Fact]
    public void ComSweepResult_CarriesTheUnsortedCappedFoldersAcrossTheWire()
    {
        // The sort failure is observed in the CHILD, where the table lives, and it decides
        // WHICH advice sentence the parent emits about the cap (gap H2). Lost on the hop, the
        // list reads as empty, the folder falls back into the ordinary capped set, and the
        // parent goes back to telling the caller the OLDEST mail is what is missing - which
        // is the false statement this whole field exists to retire.
        ComSweepResult original = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            itemCappedFolders: new[] { "alice@example.com/Inbox", "alice@example.com/Sent Items" },
            itemCappedFoldersUnsorted: new[] { "alice@example.com/Sent Items" });

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComSweepResult? read = JsonSerializer.Deserialize<ComSweepResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(2, read!.ItemCappedFolders.Count);
        Assert.Equal(new[] { "alice@example.com/Sent Items" }, read.ItemCappedFoldersUnsorted);
    }

    [Fact]
    public void ComSweepResult_CarriesTheBodyTruncationFactsAcrossTheWire()
    {
        // Both facts are measured in the CHILD, where the body bounds are applied, and both
        // decide what the PARENT says. Lost on the hop, the per-item flag reads as "not cut"
        // and the budget flag as "the per-item ceiling did it" - so a sweep that cut mail
        // reports itself complete, which is the exact species of silence the bound was added
        // WITH its reporting to avoid.
        ComSweepResult original = new ComSweepResult(
            new[] { CutBrief() },
            foldersSwept: 4,
            foldersSkipped: 0,
            bodiesTruncated: 3,
            bodyBudgetExhausted: true);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComSweepResult? read = JsonSerializer.Deserialize<ComSweepResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(3, read!.BodiesTruncated);
        Assert.True(read.BodyBudgetExhausted);
        Assert.True(Assert.Single(read.Items).BodyTruncated);
    }

    [Fact]
    public void ComMailBrief_OmitsTheBodyTruncationFlagWhenNothingWasCut()
    {
        // The flag is nullable for frame size, which is the very thing the bound exists to
        // control: an untruncated sweep of 800 items must carry this field zero times, not
        // 800 times. Checked on the wire rather than asserted about the type.
        string json = JsonSerializer.Serialize(CutBrief(cut: false), ComHostProtocol.Json);

        Assert.DoesNotContain("bodyTruncated", json, StringComparison.OrdinalIgnoreCase);
        Assert.Null(JsonSerializer.Deserialize<ComMailBrief>(json, ComHostProtocol.Json)!.BodyTruncated);
    }

    private static ComMailBrief CutBrief(bool cut = true)
    {
        return new ComMailBrief(
            "0000000000000000000000000000000000000000000000AB",
            "alice@example.com",
            "storeid",
            "Inbox",
            "inbox",
            "A very long thread",
            null,
            null,
            null,
            null,
            null,
            null,
            "the first 500 000 characters of it",
            null,
            cut ? true : (bool?)null);
    }

    [Fact]
    public void ComMailBrief_CarriesTheMessageClassAcrossTheWire()
    {
        // The snapshot is taken in the CHILD and the class decides a payload field in the
        // PARENT (`itemClass`). Since the tiers stopped filtering by class (gap B3) this is
        // the only thing that tells a caller a hit is a bounce report rather than mail, so
        // losing it on the hop would return the widened result set with no way to read it -
        // silently, because a null class is indistinguishable from ordinary mail on the wire.
        ComMailBrief original = new ComMailBrief(
            "0000000000000000000000000000000000000000000000AB",
            "alice@example.com",
            "storeid",
            "Inbox",
            "inbox",
            "Undeliverable: Invoice",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            "REPORT.IPM.Note.NDR");

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComMailBrief? read = JsonSerializer.Deserialize<ComMailBrief>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal("REPORT.IPM.Note.NDR", read!.MessageClass);
    }

    [Fact]
    public void ComExhaustiveResult_CarriesItsDroppedRowCountsAcrossTheWire()
    {
        // The scan runs in the child too (gap F5). Both numbers still cross the wire, and
        // their DIFFERENCE was the item-class filter gap B3 was about - which is now zero by
        // construction, because the scan admits every class. rowsUnreadable is what makes a
        // scan partial; rowsDropped stays the total and raises nothing.
        ComExhaustiveResult original = new ComExhaustiveResult(
            Array.Empty<ComMailBrief>(),
            foldersScanned: 12,
            foldersSkipped: 1,
            engine: "ci_phrasematch",
            instantSearchEnabled: true,
            truncated: false,
            timedOut: false,
            rowsDropped: 28,
            rowsUnreadable: 3,
            depthLimitReached: true);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComExhaustiveResult? read = JsonSerializer.Deserialize<ComExhaustiveResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(28, read!.RowsDropped);
        Assert.Equal(3, read.RowsUnreadable);

        // The depth guard latches in the CHILD (gap F4), so a flag that does not survive the
        // hop is a truncated scan reported as complete - the exact shape ComFolderTree's
        // bounds were fixed for one test below.
        Assert.True(read.DepthLimitReached);
    }

    [Fact]
    public void ComFolderTree_CarriesWhatBoundedTheWalkAcrossTheWire()
    {
        // The walk runs in the CHILD and its bounds decide list_folders' truncated flag,
        // its advice and whether a nextOffset is offered at all (gap G3). These used to be
        // no part of the return value, so there was nothing to lose on the hop; now there
        // is, and a bound that reads as false in the parent puts the payload back to
        // reporting a cut-off tree as a complete answer.
        ComFolderTree original = new ComFolderTree(
            new[] { new ComFolderInfo("alice@example.com", "Inbox", "Inbox", 12, 3, 0) },
            walkCapReached: true,
            depthLimitReached: true,
            storesUnnamed: 2,
            storesUnnamedExcluded: 1);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComFolderTree? read = JsonSerializer.Deserialize<ComFolderTree>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Single(read!.Folders);
        Assert.True(read.WalkCapReached);
        Assert.True(read.DepthLimitReached);
        Assert.Equal(2, read.StoresUnnamed);
        Assert.Equal(1, read.StoresUnnamedExcluded);
    }

    [Fact]
    public void ComFolderPathList_CarriesItsTruncationAcrossTheWire()
    {
        // Same hop, and this is the one where losing the flag costs MAIL rather than a
        // listing (gap G4): a delegate folder scope is an OR of folder names taken from this
        // list, so the parent has to know the list is short of the mailbox's real folder set.
        ComFolderPathList original = new ComFolderPathList(
            new[] { "Archive", "Archive/2024" }, walkCapReached: true, depthLimitReached: false);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComFolderPathList? read = JsonSerializer.Deserialize<ComFolderPathList>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(2, read!.Paths.Count);
        Assert.True(read.WalkCapReached);
        Assert.False(read.DepthLimitReached);
        Assert.True(read.Incomplete);
    }

    [Fact]
    public void ComSweepResult_CarriesItsUnnameableStoreCountAcrossTheWire()
    {
        // Gap G2. The sweep names such a store in the child (a StoreNaming label) and the
        // parent reports the count and the advice sentence, so this counter has the same
        // pipe dependency as every other one beside it.
        ComSweepResult original = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreNaming.LabelForUnnamedStore(2), 4, 0, 0, 0),
            },
            storesUnnamed: 1);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComSweepResult? read = JsonSerializer.Deserialize<ComSweepResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(1, read!.StoresUnnamed);
        Assert.Equal("(unnamed store 2)", Assert.Single(read.PerStore).StoreDisplayName);
    }

    [Fact]
    public void ComStoreDetail_CarriesTheUnreadableNameFlagAcrossTheWire()
    {
        // The flag decides whether list_accounts prints a label as a usable store name, and
        // whether list_folders' refusal claims a store is absent when it cannot know (G2).
        ComStoreDetail original = new ComStoreDetail(
            StoreNaming.LabelForUnnamedStore(3), "storeid", 3, null, nameUnreadable: true);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComStoreDetail? read = JsonSerializer.Deserialize<ComStoreDetail>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.True(read!.NameUnreadable);
        Assert.Equal("(unnamed store 3)", read.DisplayName);
    }

    [Fact]
    public void ComSweepResult_WithNoPerStoreCounters_RoundTripsAsEmptyNotNull()
    {
        // Every consumer reads PerStore without a null check, and "the sweep reached no
        // store" must stay a legible answer rather than a NullReferenceException in the
        // parent after a child that reported one.
        ComSweepResult original = new ComSweepResult(Array.Empty<ComMailBrief>(), foldersSwept: 0, foldersSkipped: 0);

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComSweepResult? read = JsonSerializer.Deserialize<ComSweepResult>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Empty(read!.PerStore);
    }

    // ------------------------------------------------------------- output parameters

    [Fact]
    public async Task Response_CarriesOutputParameters()
    {
        // Most of the contract reports failure through `out string? error`, and the
        // service layer branches on that string - retrying across stores only when it
        // reads "ItemNotFound". Losing these would change behaviour, not break the build.
        ComHostResponse sent = new ComHostResponse
        {
            Id = 5,
            Ok = true,
            Outputs = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
            {
                ["error"] = JsonSerializer.SerializeToElement("ItemNotFound", ComHostProtocol.Json),
                ["sizeBytes"] = JsonSerializer.SerializeToElement(4096L, ComHostProtocol.Json),
            },
        };

        using MemoryStream stream = new MemoryStream(ComHostProtocol.EncodeFrame(sent));
        ComHostResponse? read = await ComHostProtocol.ReadFrameAsync<ComHostResponse>(stream, CancellationToken.None);

        Assert.NotNull(read);
        Assert.Equal("ItemNotFound", read!.Outputs!["error"].GetString());
        Assert.Equal(4096L, read.Outputs["sizeBytes"].GetInt64());
    }

    [Fact]
    public async Task Error_PreservesTypeHResultAndReason()
    {
        // Guard branches on exception TYPE and ComGateway keys its disconnect-retry on
        // HRESULTs, so a flattened error would silently downgrade both.
        ComHostResponse sent = new ComHostResponse
        {
            Id = 9,
            Ok = false,
            Error = new ComHostError
            {
                Type = "COMException",
                Message = "boom",
                HResult = unchecked((int)0x80010108),
                Reason = "not_created_by_this_server",
            },
        };

        using MemoryStream stream = new MemoryStream(ComHostProtocol.EncodeFrame(sent));
        ComHostResponse? read = await ComHostProtocol.ReadFrameAsync<ComHostResponse>(stream, CancellationToken.None);

        Assert.NotNull(read);
        Assert.False(read!.Ok);
        Assert.Equal("COMException", read.Error!.Type);
        Assert.Equal(unchecked((int)0x80010108), read.Error.HResult);
        Assert.Equal("not_created_by_this_server", read.Error.Reason);
    }

    // ------------------------------------------------- what the CHILD puts on the wire
    //
    // The three tests above pin that a faithful error survives the wire. They cannot see
    // the defect found on 2026-08-18, which was one step earlier: the child was FILLING IN
    // that error from a reflection wrapper, so a perfectly transmitted frame carried
    // "TargetInvocationException" / "Exception has been thrown by the target of an
    // invocation." / no HRESULT / no reason. Every deliberate error the session raises read
    // the same, and the parent's exception-type mapping was unreachable in production.

    [Fact]
    public void AFailureFromTheSession_CrossesTheRoutingProxyAsItself()
    {
        // The routing proxy calls the session by reflection. Reflection wraps; the wrapper
        // must not escape, or the type and HRESULT the wire cares about are gone before the
        // wire is ever reached.
        IOutlookSession failing = ComHostFaultInjection.FaultingSession.Create("com");
        RecordingGateway gateway = new RecordingGateway(failing);
        IOutlookSession proxy = GatewayRoutingProxy.Create(gateway);

        COMException thrown = Assert.Throws<COMException>(() => proxy.GetProfileName());

        Assert.Equal(ComHostFaultInjection.SessionComMessage, thrown.Message);
        Assert.Equal(ComHostFaultInjection.SessionComHResult, thrown.HResult);
    }

    [Fact]
    public void TheGatewaySeesTheRealFailure_SoItsDisconnectRebuildStillFires()
    {
        // ComGateway.Run rebuilds a dead session exactly once, and decides whether to by
        // testing the exception it catches around the SAME call this records. Wrapped, that
        // test could never match: a TargetInvocationException is neither a COMException nor
        // an InvalidComObjectException, and carries no HRESULT to compare. So the one-shot
        // disconnect recovery was dead code inside the COM host, silently, for as long as
        // the wrapper escaped. Asserting the observed type and HRESULT rather than calling
        // the predicate keeps this a test of what happens, not a copy of the rule.
        IOutlookSession failing = ComHostFaultInjection.FaultingSession.Create("com");
        RecordingGateway gateway = new RecordingGateway(failing);
        IOutlookSession proxy = GatewayRoutingProxy.Create(gateway);

        _ = Assert.Throws<COMException>(() => proxy.GetProfileName());

        Assert.NotNull(gateway.Observed);
        COMException observed = Assert.IsType<COMException>(gateway.Observed);
        Assert.Equal(ComHostFaultInjection.SessionComHResult, observed.HResult);
    }

    [Fact]
    public void UnwrappingAFailure_KeepsTheStackItWasThrownWith()
    {
        // The unwrap is done with ExceptionDispatchInfo rather than `throw ex.InnerException`
        // precisely so this stays true. The stack inside the session is the only record of
        // WHERE a failure happened; a rethrow that resets it trades one kind of blindness
        // for another, and the COM host writes uncaught failures to stderr with their full
        // stack for exactly this reason.
        IOutlookSession failing = ComHostFaultInjection.FaultingSession.Create("folder");
        IOutlookSession proxy = GatewayRoutingProxy.Create(new RecordingGateway(failing));

        InvalidOperationException thrown = Assert.Throws<InvalidOperationException>(() => proxy.GetProfileName());

        Assert.Equal(ComHostFaultInjection.SessionFolderMessage, thrown.Message);
        Assert.Contains(
            nameof(ComHostFaultInjection.FaultingSession),
            thrown.StackTrace ?? string.Empty,
            StringComparison.Ordinal);
    }

    /// <summary>
    /// A gateway that runs the operation and remembers what came back out of it - standing
    /// in for the position <c>ComGateway.Run</c>'s exception filter occupies, without
    /// needing an Outlook to connect to.
    /// </summary>
    private sealed class RecordingGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal RecordingGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        internal Exception? Observed { get; private set; }

        /// <summary>The recovery the caller asked for on the last <see cref="Run{T}"/>.</summary>
        internal ComSessionRecovery? RequestedRecovery { get; private set; }

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        public T Run<T>(Func<IOutlookSession, T> operation)
        {
            return Run(operation, ComSessionRecovery.None);
        }

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery)
        {
            RequestedRecovery = recovery;
            try
            {
                return operation(_session);
            }
            catch (Exception ex)
            {
                Observed = ex;
                throw;
            }
        }

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
        {
            return Run(operation);
        }

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }
}
