using System.Text.Json;
using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The test-side MCP client's gate on the Exchange VM's test account
/// (<see cref="StdioWritePosture.WritableHubOnly"/>, 2026-10-08): the out-of-process half of the Phase 2
/// write-safety design. The MCP server writes through its OWN Outlook session, which neither the in-process
/// <see cref="StoreWriteAllowlist"/> nor <see cref="LiveGuardedOutlookSession"/> can see, so every
/// write-capable <c>tools/call</c> is judged here before a byte of it is sent:
/// <list type="bullet">
/// <item><b>Addresses</b>: every <c>to</c>, <c>cc</c> and <c>bcc</c> must be the hub's own address, and a new
/// draft's <c>account</c> the hub's (<see cref="RecipientAllowlist"/>);</item>
/// <item><b>Subjects</b>: a new subject - a new draft's, an override - must carry the tag and this run's
/// marker;</item>
/// <item><b>Targets</b>: the item a reply, reply-all, forward, update, discard, send, move, archive or open
/// names must be THIS RUN'S, in the hub - a hit id this client saw in a search or thread result whose subject
/// carries the tag and the marker and whose store is the hub, or an EntryID the product's own reader, in this
/// process, finds in the hub with the tag and the marker; and for a reply, a reply-all and a send, its
/// sender and every recipient must be the hub (<see cref="LiveRunItemCheck"/>);</item>
/// <item>a tool nobody classified is refused, and the navigation tools, <c>save_attachment</c> and
/// <c>manage_signature</c>, which address no mail and touch no item's content, go through.</item>
/// </list>
/// </summary>
public sealed class StdioHubOnlyGate : IDisposable
{
    private static readonly Regex HitId = new("^h[0-9]+$", RegexOptions.CultureInvariant);
    private static readonly Regex EntryIdHex = new("^[0-9A-Fa-f]{40,}$", RegexOptions.CultureInvariant);

    private readonly LiveRunItemCheck _check;
    private readonly Func<string, string, bool, string?>? _entryIdRefusal;
    private readonly Dictionary<string, SeenHit> _hits = new(StringComparer.Ordinal);
    private readonly object _gate = new();
    private OutlookComSession? _session;

    /// <summary>The gate for one client.</summary>
    /// <param name="check">This run's item check.</param>
    /// <param name="entryIdRefusal">
    /// Judges an EntryID: (entryId, operation, addressesMail) to a refusal or null. Null to read it through
    /// the product's own session in this process - the live default; T1 passes a fake.
    /// </param>
    public StdioHubOnlyGate(LiveRunItemCheck check, Func<string, string, bool, string?>? entryIdRefusal = null)
    {
        _check = check ?? throw new ArgumentNullException(nameof(check));
        _entryIdRefusal = entryIdRefusal;
    }

    /// <summary>The phrase every refusal of this gate carries.</summary>
    public const string Refused = "HUB-ONLY GATE REFUSED";

    /// <summary>The hub's address - the one a write may come from and go to.</summary>
    public string Hub => _check.HubStoreDisplayName;

    private sealed record SeenHit(string? Subject, string? Store, string? FromAddress);

    /// <summary>
    /// Remembers the hits a result named - every object in it carrying a hit id, a subject and a store - so a
    /// later write naming one of them can be judged. Never throws.
    /// </summary>
    public void Observe(JsonElement payload)
    {
        lock (_gate)
        {
            Walk(payload, 0);
        }
    }

    private void Walk(JsonElement element, int depth)
    {
        if (depth > 12)
        {
            return;
        }

        if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement child in element.EnumerateArray())
            {
                Walk(child, depth + 1);
            }

            return;
        }

        if (element.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        if (element.TryGetProperty("id", out JsonElement id) && id.ValueKind == JsonValueKind.String && HitId.IsMatch(id.GetString()!))
        {
            _hits[id.GetString()!] = new SeenHit(StringOf(element, "subject"), StringOf(element, "store"), StringOf(element, "fromAddress"));
        }

        foreach (JsonProperty property in element.EnumerateObject())
        {
            if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
            {
                Walk(property.Value, depth + 1);
            }
        }
    }

    private static string? StringOf(JsonElement element, string name)
    {
        return element.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
    }

    /// <summary>The refusal for one <c>tools/call</c> of <paramref name="tool"/>, or null to send it.</summary>
    public string? RefusalFor(string? tool, JsonElement arguments)
    {
        switch (tool)
        {
            case "search":
            case "thread":
            case "read":
            case "list_accounts":
            case "list_folders":
            case "list_signatures":
            case "outlook_health":
            case "audit_log":
            case "goto_folder":
            case "show_search_results":
            case "save_attachment":
            case "manage_signature":
                return null;

            case "new_draft":
                return First(
                    AccountRefusal(arguments),
                    RecipientRefusal(tool, arguments, "to", "cc", "bcc"),
                    _check.RefusalForNewSubject(StringArgument(arguments, "subject"), tool));

            case "reply_draft":
            case "replyall_draft":
                return First(
                    TargetRefusal(tool, StringArgument(arguments, "id"), addressesMail: true),
                    RecipientRefusal(tool, arguments, "cc", "bcc"),
                    SubjectOverrideRefusal(tool, arguments));

            case "forward_draft":
                return First(
                    TargetRefusal(tool, StringArgument(arguments, "id"), addressesMail: false),
                    RecipientRefusal(tool, arguments, "to", "cc", "bcc"),
                    SubjectOverrideRefusal(tool, arguments));

            case "update_draft":
                return First(
                    TargetRefusal(tool, StringArgument(arguments, "id"), addressesMail: false),
                    RecipientRefusal(tool, arguments, "to", "cc", "bcc"),
                    SubjectOverrideRefusal(tool, arguments));

            case "discard_draft":
            case "open_in_outlook":
                return TargetRefusal(tool, StringArgument(arguments, "id"), addressesMail: false);

            case "send":
                return TargetRefusal(tool, StringArgument(arguments, "id"), addressesMail: true);

            case "move_mail":
            case "archive_mail":
                foreach (string? id in StringArrayArgument(arguments, "ids"))
                {
                    string? refusal = TargetRefusal(tool, id, addressesMail: false);
                    if (refusal != null)
                    {
                        return refusal;
                    }
                }

                return null;

            default:
                return Refused + ": '" + (tool ?? "(no tool name)") + "' is not a tool this gate knows, and on the Exchange VM's "
                    + "test account a tool nobody has judged is refused like a write to real mail.";
        }
    }

    private static string? First(params string?[] refusals)
    {
        return refusals.FirstOrDefault(r => r != null);
    }

    private string? AccountRefusal(JsonElement arguments)
    {
        string? account = StringArgument(arguments, "account");
        return _check.Recipients.IsAllowed(account)
            ? null
            : RecipientAllowlist.Refused + ": new_draft from the account '" + account + "' - on this machine a draft is made only in "
                + "the hub's own account.";
    }

    private string? RecipientRefusal(string tool, JsonElement arguments, params string[] fields)
    {
        return _check.Recipients.RefusalFor(fields.Select(f => StringArgument(arguments, f)), tool);
    }

    private string? SubjectOverrideRefusal(string tool, JsonElement arguments)
    {
        string? subject = StringArgument(arguments, "subject");
        return subject == null ? null : _check.RefusalForNewSubject(subject, tool);
    }

    /// <summary>Whether the item <paramref name="id"/> names is this run's, in the hub.</summary>
    private string? TargetRefusal(string tool, string? id, bool addressesMail)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            // The server refuses a missing id itself, before any Outlook work.
            return null;
        }

        if (HitId.IsMatch(id))
        {
            SeenHit? hit;
            lock (_gate)
            {
                _hits.TryGetValue(id, out hit);
            }

            if (hit == null)
            {
                return Refused + ": '" + tool + "' names the hit '" + id + "', which no search or thread result of this client "
                    + "showed - so nothing here can say it is this run's own item.";
            }

            if (tool == "send")
            {
                return Refused + ": 'send' by a hit id - on this machine a send names the draft's EntryID, so that the draft's "
                    + "own recipients can be read before it goes.";
            }

            // A hit carries its sender only when the index row names one; without it, the hit is judged by its
            // provenance alone - the hub, the tag, this run's marker - which only this run's guarded helpers can
            // have given an item, and they address nobody but the hub. A sender that IS named must be the hub.
            bool judgeSender = addressesMail && !string.IsNullOrWhiteSpace(hit.FromAddress);
            return _check.RefusalForReading(
                hit.Store, hit.Subject, judgeSender ? hit.FromAddress : null, Array.Empty<string?>(), tool, judgeSender);
        }

        if (EntryIdHex.IsMatch(id))
        {
            return (_entryIdRefusal ?? ReadThroughTheProduct)(id, tool, addressesMail);
        }

        // Neither a hit id nor an EntryID: the server rejects it as malformed before it reaches Outlook.
        return null;
    }

    /// <summary>
    /// The live judge: the item as the product's own reader, in this process, finds it - through one
    /// <see cref="OutlookComSession"/> held for the client's life (attaching to the running Outlook; the
    /// suite started it long before any write).
    /// </summary>
    private string? ReadThroughTheProduct(string entryId, string tool, bool addressesMail)
    {
        IOutlookSession session;
        lock (_gate)
        {
            session = _session ??= OutlookComSession.Connect(allowStartingOutlook: false);
        }

        if (tool == "send")
        {
            // The draft as it stands, read the way the send path reads it.
            ComSendableDraftState? state = session.TryGetSendableDraftState(entryId, null, out _);
            if (state == null && _check.HubStoreId(session) is { } hubStoreId)
            {
                state = session.TryGetSendableDraftState(entryId, hubStoreId, out _);
            }

            if (state != null)
            {
                return _check.RefusalForReading(
                    state.StoreDisplayName, state.Subject, state.ResolvedAccountSmtp, state.Recipients.Select(r => r.Address), tool, true);
            }
        }

        return _check.RefusalForItem(session, entryId, null, tool, addressesMail);
    }

    private static string? StringArgument(JsonElement arguments, string name)
    {
        return arguments.ValueKind == JsonValueKind.Object && arguments.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
    }

    private static IEnumerable<string?> StringArrayArgument(JsonElement arguments, string name)
    {
        if (arguments.ValueKind != JsonValueKind.Object || !arguments.TryGetProperty(name, out JsonElement value) || value.ValueKind != JsonValueKind.Array)
        {
            return Array.Empty<string?>();
        }

        return value.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() : null).ToList();
    }

    /// <summary>Releases the product session the live judge opened, if it opened one. Never quits Outlook.</summary>
    public void Dispose()
    {
        OutlookComSession? session;
        lock (_gate)
        {
            session = _session;
            _session = null;
        }

        session?.Dispose();
    }
}
