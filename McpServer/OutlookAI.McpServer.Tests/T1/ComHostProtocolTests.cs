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

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        public T Run<T>(Func<IOutlookSession, T> operation)
        {
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
