using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookAI.ComHost.Protocol
{
    /// <summary>
    /// The wire contract between the MCP server (parent) and the COM host (child).
    /// <para>
    /// Framing is length-prefixed rather than newline-delimited on purpose: payloads
    /// carry mail bodies and HTML that contain newlines in abundance, and a length
    /// prefix cannot be desynchronised by content. A desync here would reproduce, one
    /// layer down, exactly the class of silent hang this architecture exists to remove.
    /// </para>
    /// <para>
    /// Every frame is: 4-byte little-endian unsigned length, then that many bytes of
    /// UTF-8 JSON. <see cref="MaxFrameBytes"/> bounds a single frame so a corrupt or
    /// hostile length cannot make the reader allocate without limit.
    /// </para>
    /// </summary>
    internal static class ComHostProtocol
    {
        /// <summary>
        /// Hard ceiling on one frame. `read` can legitimately return ~0.5 MB; 64 MB was
        /// chosen as far above any real payload and far below a denial-of-service
        /// allocation.
        /// <para>
        /// "Far above any real payload" was an assumption, and a sweep of an unindexed
        /// store can approach it (see TODO.md for the derivation). It is now checkable
        /// rather than asserted: <see cref="ComHostFrameMeter"/> records the largest frame
        /// actually seen and <c>outlook_health</c> reports it beside this number.
        /// </para>
        /// </summary>
        internal const int MaxFrameBytes = 64 * 1024 * 1024;

        /// <summary>Environment variable carrying the pipe name from parent to child.</summary>
        internal const string PipeNameVariable = "OUTLOOKAI_COMHOST_PIPE";

        /// <summary>Environment variable carrying the parent PID so the child can exit if orphaned.</summary>
        internal const string ParentPidVariable = "OUTLOOKAI_COMHOST_PARENT_PID";

        /// <summary>
        /// Shared serializer options. Web defaults (camelCase) to match the tool payload
        /// convention, nulls omitted to keep frames small, and no indentation.
        /// </summary>
        internal static readonly JsonSerializerOptions Json = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = false,
            Converters = { new ComDraftBodyConverter() },
        };

        /// <summary>
        /// Builds one frame, or refuses to.
        /// <para>
        /// <paramref name="maxFrameBytes"/> exists so the refusal branch can be reached
        /// without allocating a real 64 MB payload. Production never passes it: both
        /// callers take the default, so the shipped limit is still the constant above.
        /// The alternative was a test that allocates past the constant, in a tier that
        /// runs in about two minutes - and an untested refusal is how this branch came to
        /// kill the host in the first place.
        /// </para>
        /// </summary>
        internal static byte[] EncodeFrame<T>(T message, int maxFrameBytes = MaxFrameBytes)
        {
            byte[] payload = JsonSerializer.SerializeToUtf8Bytes(message, Json);
            if (payload.Length > maxFrameBytes)
            {
                ComHostFrameMeter.Shared.RecordRefusal();
                throw new ComHostProtocolException(
                    $"Frame of {payload.Length} bytes exceeds the {maxFrameBytes} byte limit.");
            }

            ComHostFrameMeter.Shared.RecordFrame(payload.Length);
            byte[] frame = new byte[4 + payload.Length];
            BitConverter.TryWriteBytes(frame.AsSpan(0, 4), (uint)payload.Length);
            payload.CopyTo(frame.AsSpan(4));
            return frame;
        }

        /// <summary>
        /// Reads one frame. Returns null on a clean end of stream (the peer closed the
        /// pipe), which callers must treat as "peer gone", never as "keep waiting".
        /// </summary>
        internal static async Task<TMessage?> ReadFrameAsync<TMessage>(Stream stream, CancellationToken cancellationToken)
            where TMessage : class
        {
            byte[] header = new byte[4];
            if (!await ReadExactlyAsync(stream, header, cancellationToken).ConfigureAwait(false))
            {
                return null;
            }

            uint length = BitConverter.ToUInt32(header, 0);
            if (length > MaxFrameBytes)
            {
                throw new ComHostProtocolException(
                    $"Declared frame length {length} exceeds the {MaxFrameBytes} byte limit; the stream is desynchronised.");
            }

            if (length == 0)
            {
                throw new ComHostProtocolException("Zero-length frame; the stream is desynchronised.");
            }

            byte[] payload = new byte[length];
            if (!await ReadExactlyAsync(stream, payload, cancellationToken).ConfigureAwait(false))
            {
                return null;
            }

            // Measured here as well as on encode because the two ends are different
            // PROCESSES: the answers that get big are built in the child, whose counters
            // die with it, and this is the only place the parent - the side with a health
            // surface - ever learns how big they were. Recorded after the payload is fully
            // read, so the number means "this much actually crossed".
            ComHostFrameMeter.Shared.RecordFrame(length);

            return JsonSerializer.Deserialize<TMessage>(payload, Json)
                ?? throw new ComHostProtocolException("Frame deserialized to null.");
        }

        private static async Task<bool> ReadExactlyAsync(Stream stream, byte[] buffer, CancellationToken cancellationToken)
        {
            int offset = 0;
            while (offset < buffer.Length)
            {
                int read = await stream
                    .ReadAsync(buffer.AsMemory(offset, buffer.Length - offset), cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0)
                {
                    // Clean EOF. Partway through a frame this is a truncated write by a
                    // peer that died mid-send - still "gone", not "retry".
                    return false;
                }

                offset += read;
            }

            return true;
        }

        /// <summary>Builds a per-process-unique pipe name.</summary>
        internal static string NewPipeName()
        {
            return "OutlookAI.ComHost." + Guid.NewGuid().ToString("N");
        }

        internal static string Describe(byte[] utf8)
        {
            return Encoding.UTF8.GetString(utf8);
        }
    }

    /// <summary>Raised when the pipe carries something that is not a well-formed frame.</summary>
    internal sealed class ComHostProtocolException : Exception
    {
        internal ComHostProtocolException(string message)
            : base(message)
        {
        }

        internal ComHostProtocolException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    /// <summary>
    /// Raised when the COM host produced an answer too large to fit in one frame, so it
    /// was refused instead of sent.
    /// <para>
    /// Public, and living beside the limit rather than with the supervision exceptions,
    /// because all three layers need to name it: the child stamps it on the wire error,
    /// <c>ComHostErrorMapper</c> rebuilds it, and the tool layer catches it to choose the
    /// advice an agent reads. Before this existed the refusal escaped the serve loop and
    /// ended the child process, so the caller was told the host had gone away - the one
    /// fact that says nothing about what to do next, and which invites a retry of the
    /// exact request that cannot succeed.
    /// </para>
    /// <para>
    /// Nothing failed in Outlook when this is raised: the work COMPLETED, and only the
    /// reply was refused. That is why it is not modelled as a transport fault - and it is
    /// also why the wording here used to be wrong. "Nothing was changed" is true of a
    /// search whose answer was too big and false of a draft that was created, so
    /// <see cref="Operation"/> carries the name the tool layer needs to tell the two apart.
    /// </para>
    /// </summary>
    public sealed class ComHostResponseTooLargeException : Exception
    {
        /// <summary>Creates the exception.</summary>
        public ComHostResponseTooLargeException(string message, string? operation = null)
            : base(message)
        {
            Operation = operation;
        }

        /// <summary>
        /// The contract operation whose answer was refused, or null when the parent could
        /// not attribute it. Null must not be read as "a read": callers state no outcome at
        /// all rather than guess one.
        /// </summary>
        public string? Operation { get; }
    }

    /// <summary>A parent -> child operation request.</summary>
    internal sealed class ComHostRequest
    {
        /// <summary>Correlation id, unique for the life of one child connection.</summary>
        public long Id { get; set; }

        /// <summary>Operation name - the <see cref="OutlookAI.Core.Com.IOutlookSession"/> method being invoked.</summary>
        public string Operation { get; set; } = string.Empty;

        /// <summary>JSON-encoded argument object for the operation, or null for no-arg operations.</summary>
        public JsonElement? Arguments { get; set; }
    }

    /// <summary>A child -> parent response, or an unsolicited event when <see cref="Event"/> is set.</summary>
    internal sealed class ComHostResponse
    {
        /// <summary>Correlation id echoing the request; 0 for unsolicited events.</summary>
        public long Id { get; set; }

        /// <summary>Set for unsolicited notifications (e.g. Outlook exited). Null for replies.</summary>
        public string? Event { get; set; }

        /// <summary>True when the operation completed; false when <see cref="Error"/> describes a failure.</summary>
        public bool Ok { get; set; }

        /// <summary>JSON-encoded return value; absent for void operations.</summary>
        public JsonElement? Result { get; set; }

        /// <summary>
        /// Values of <c>out</c>/<c>ref</c> parameters, keyed by parameter name.
        /// <para>
        /// Not an optional nicety: most of the contract reports failure detail through
        /// <c>out string? error</c> rather than by throwing, and the service layer
        /// branches on that string (for example retrying across stores only when it reads
        /// "ItemNotFound"). Dropping these on the floor would silently change behaviour
        /// rather than break the build.
        /// </para>
        /// </summary>
        public Dictionary<string, JsonElement>? Outputs { get; set; }

        /// <summary>Structured failure, faithfully carrying the child-side exception shape.</summary>
        public ComHostError? Error { get; set; }
    }

    /// <summary>
    /// A child-side failure, rendered so the parent can rethrow an equivalent exception.
    /// Preserving the type name and HRESULT matters: <c>ComGateway</c> keys its
    /// disconnect-retry and session-unusable decisions off exactly those.
    /// </summary>
    internal sealed class ComHostError
    {
        /// <summary>Exception type name, e.g. "COMException", "OutlookUnavailableException".</summary>
        public string Type { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        /// <summary>HRESULT for COM failures; null otherwise.</summary>
        public int? HResult { get; set; }

        /// <summary>Machine-readable refusal code for draft/send refusals.</summary>
        public string? Reason { get; set; }
    }

    /// <summary>Unsolicited child -> parent event names.</summary>
    internal static class ComHostEvents
    {
        /// <summary>The child's Outlook signalled Quit or its process exited; the parent must drop cached session state.</summary>
        internal const string OutlookGone = "outlookGone";

        /// <summary>The child finished starting and is ready to serve operations.</summary>
        internal const string Ready = "ready";
    }
}
