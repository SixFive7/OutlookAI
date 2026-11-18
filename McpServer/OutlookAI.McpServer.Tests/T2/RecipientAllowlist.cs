namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Code-enforced answer to "whom may a test address" on a machine whose mail reaches the internet - the
/// Exchange VM, since 2026-10-08 writing in its dedicated test account (Q110, the Phase 2 write-safety
/// design of <c>Docs/live-tier-on-the-vm.md</c> section 4.4): the hub's own address, and nothing else.
/// <para>
/// <b>Why an allowlist and not care.</b> Every other machine of the testbed is offline, and its mail ends
/// in a loopback sink. The Exchange VM's goes out through Exchange Online: one address typed wrong, one
/// reply-all to a real message, one forward with the wrong <c>to</c>, and a test has mailed a person. So
/// every address a test hands to Outlook - the To, Cc and Bcc of a new draft, a reply, a forward or an
/// update, and every recipient a draft holds when it is sent - is checked here first, and anything that is
/// not the hub's own address refuses the operation before it is saved or sent.
/// </para>
/// <para>
/// <b>Where it is asked.</b> <see cref="LiveOutlookTestMailer"/> before every self-send;
/// <see cref="LiveGuardedOutlookSession"/> - the in-process product's every draft and send; and
/// <see cref="StdioHubOnlyGate"/>, the test-side MCP client, before every write-capable tool call. One
/// rule, three doors.
/// </para>
/// <para>
/// Pure: strings in, a verdict out. CI pins every branch of it without Outlook.
/// </para>
/// </summary>
public sealed class RecipientAllowlist
{
    /// <summary>The phrase every refusal carries, so a run log and the T1 pins can find it.</summary>
    public const string Refused = "RECIPIENT REFUSED";

    private readonly HashSet<string> _allowed;

    /// <summary>
    /// A string that, inside an entry with no '@', marks that entry as a test's deliberately UNRESOLVABLE
    /// name - this run's marker, which no directory entry carries - or null to allow none.
    /// </summary>
    private readonly string? _unresolvableMarker;

    /// <summary>Builds the allowlist around the addresses a test may write to - normally the hub's one.</summary>
    /// <param name="allowedAddresses">Each must be an address (it contains '@'); compared ignoring case.</param>
    /// <param name="unresolvableMarker">
    /// When given (12 characters or more), an entry with no '@' that contains it is allowed too: a test's
    /// deliberately unresolvable name, which Outlook resolves to nobody because no directory entry carries
    /// this run's marker. Such a draft cannot be sent - its recipient has no address, which the send-time
    /// check refuses (<see cref="RefusalForItem"/>).
    /// </param>
    public RecipientAllowlist(IEnumerable<string> allowedAddresses, string? unresolvableMarker = null)
    {
        _unresolvableMarker = string.IsNullOrWhiteSpace(unresolvableMarker) || unresolvableMarker.Trim().Length < 12
            ? null
            : unresolvableMarker.Trim();
        ArgumentNullException.ThrowIfNull(allowedAddresses);
        _allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string address in allowedAddresses)
        {
            string trimmed = (address ?? string.Empty).Trim();
            if (trimmed.Length == 0 || !trimmed.Contains('@', StringComparison.Ordinal) || trimmed.IndexOfAny(Separators) >= 0)
            {
                throw new ArgumentException(
                    "A recipient allowlist holds plain addresses; '" + address + "' is not one.", nameof(allowedAddresses));
            }

            _allowed.Add(trimmed);
        }

        if (_allowed.Count == 0)
        {
            throw new ArgumentException("A recipient allowlist with no address would refuse every draft; name the hub.", nameof(allowedAddresses));
        }
    }

    /// <summary>The addresses a test may write to.</summary>
    public IReadOnlyCollection<string> Allowed => _allowed;

    private static readonly char[] Separators = { ';', ',' };

    /// <summary>
    /// The entries of one recipient field as Outlook and the MCP tools take it - separated by ';' or ',' -
    /// each reduced to its address (<see cref="AddressOf"/>), blanks dropped. Null or blank is no entry.
    /// </summary>
    public static IReadOnlyList<string> Split(string? recipients)
    {
        if (string.IsNullOrWhiteSpace(recipients))
        {
            return Array.Empty<string>();
        }

        List<string> entries = new();
        foreach (string part in recipients.Split(Separators))
        {
            string address = AddressOf(part);
            if (address.Length > 0)
            {
                entries.Add(address);
            }
        }

        return entries;
    }

    /// <summary>
    /// The address in one entry: what sits between '&lt;' and '&gt;' in a <c>Name &lt;address&gt;</c> entry, else
    /// the entry itself, trimmed, and an <c>smtp:</c> or <c>SMTP:</c> prefix removed. Never resolved
    /// against an address book: a display name is not an address, and refusing it is the safe answer.
    /// </summary>
    public static string AddressOf(string? entry)
    {
        string text = (entry ?? string.Empty).Trim();
        int open = text.LastIndexOf('<');
        int close = text.LastIndexOf('>');
        if (open >= 0 && close > open)
        {
            text = text.Substring(open + 1, close - open - 1).Trim();
        }

        if (text.StartsWith("smtp:", StringComparison.OrdinalIgnoreCase))
        {
            text = text.Substring(5).Trim();
        }

        return text;
    }

    /// <summary>
    /// True when <paramref name="address"/> - one entry, reduced by <see cref="AddressOf"/> - may be written to:
    /// an allowed address itself; a PLUS ADDRESS of one (<c>local+anything@domain</c>), which Exchange delivers
    /// to the same mailbox or, where plus addressing is off, bounces back to its sender - nobody else either
    /// way; or, when this allowlist was given a marker, a name with no '@' that carries it.
    /// </summary>
    public bool IsAllowed(string? address)
    {
        string reduced = AddressOf(address);
        if (reduced.Length == 0)
        {
            return false;
        }

        if (_allowed.Contains(reduced) || IsPlusAddressOfAllowed(reduced))
        {
            return true;
        }

        return _unresolvableMarker != null
            && !reduced.Contains('@', StringComparison.Ordinal)
            && reduced.Contains(_unresolvableMarker, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>True when <paramref name="address"/> is <c>local+tag@domain</c> of an allowed <c>local@domain</c>.</summary>
    private bool IsPlusAddressOfAllowed(string address)
    {
        int at = address.LastIndexOf('@');
        int plus = address.IndexOf('+', StringComparison.Ordinal);
        if (at <= 0 || plus <= 0 || plus > at || address.IndexOf('@') != at)
        {
            return false;
        }

        string unplussed = address.Substring(0, plus) + address.Substring(at);
        return _allowed.Contains(unplussed);
    }

    /// <summary>
    /// An address a test may give where it needs one DISTINCT from the hub's - a Cc, a Bcc: the hub's own
    /// plus address <c>local+tag@domain</c>. Every address in the allowlist is the hub's, so the first one
    /// is taken.
    /// </summary>
    public string PlusAddress(string tag)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tag);
        string hub = _allowed.First();
        int at = hub.LastIndexOf('@');
        return hub.Substring(0, at) + "+" + tag + hub.Substring(at);
    }

    /// <summary>
    /// The refusal for a set of recipient fields a TEST supplied - each a ';'/','-separated list, null for an
    /// absent field - or null when every entry may be written to. The offending entries are named: they are
    /// test code's own strings.
    /// </summary>
    /// <param name="recipientFields">The To, Cc and Bcc fields of one operation, in any order.</param>
    /// <param name="operation">What was about to happen, for the message.</param>
    public string? RefusalFor(IEnumerable<string?> recipientFields, string operation)
    {
        ArgumentNullException.ThrowIfNull(recipientFields);
        List<string> refused = new();
        foreach (string? field in recipientFields)
        {
            foreach (string entry in Split(field))
            {
                if (!IsAllowed(entry))
                {
                    refused.Add(entry);
                }
            }
        }

        if (refused.Count == 0)
        {
            return null;
        }

        return Refused + ": '" + operation + "' would address " + string.Join(", ", refused.Select(r => "'" + r + "'"))
            + ". On this machine mail reaches the internet, and a test may address only "
            + string.Join(", ", _allowed.Select(a => "'" + a + "'")) + " - the hub itself (Q110; "
            + "Docs/live-tier-on-the-vm.md section 4.4). Address the hub; never widen this list.";
    }

    /// <summary>
    /// The refusal for the recipients an ITEM holds - read off a draft before it is sent, or off a message
    /// before a reply-all copies them - or null when every one may be written to. Content-free (S4): an
    /// item's recipients may be a real correspondent's, so they are COUNTED, never named. An address that
    /// could not be read is refused like a stranger's.
    /// </summary>
    /// <param name="itemRecipientAddresses">Each recipient's address as read; null or blank when unreadable.</param>
    /// <param name="operation">What was about to happen, for the message.</param>
    public string? RefusalForItem(IEnumerable<string?> itemRecipientAddresses, string operation)
    {
        return RefusalForItem(itemRecipientAddresses, operation, allowMarkedNames: false);
    }

    /// <summary>
    /// <see cref="RefusalForItem(IEnumerable{string?}, string)"/>, and with <paramref name="allowMarkedNames"/> a name
    /// with no address that carries the run's marker is accepted - a test's deliberately unresolvable entry
    /// on a DRAFT, which can never be sent from here (the send-time check reads it as unreadable).
    /// </summary>
    public string? RefusalForItem(IEnumerable<string?> itemRecipientAddresses, string operation, bool allowMarkedNames)
    {
        ArgumentNullException.ThrowIfNull(itemRecipientAddresses);
        int total = 0;
        int strangers = 0;
        int unreadable = 0;
        foreach (string? address in itemRecipientAddresses)
        {
            total++;
            string reduced = AddressOf(address);
            if (allowMarkedNames && reduced.Length > 0 && !reduced.Contains('@', StringComparison.Ordinal) && IsAllowed(reduced))
            {
                continue;
            }

            if (reduced.Length == 0 || !reduced.Contains('@', StringComparison.Ordinal))
            {
                // No address at all - a name Outlook resolved to nobody: nothing this check can vouch for.
                unreadable++;
            }
            else if (!IsAllowed(reduced))
            {
                strangers++;
            }
        }

        if (strangers == 0 && unreadable == 0)
        {
            return null;
        }

        return Refused + ": '" + operation + "' - the item holds " + total + " recipient(s), " + strangers
            + " of them not the hub and " + unreadable + " whose address could not be read (not shown: an item's "
            + "recipients may be a real correspondent's). On this machine mail reaches the internet, and a test may "
            + "address only the hub itself (Q110).";
    }

    /// <summary>Throws <see cref="InvalidOperationException"/> with <see cref="RefusalFor"/>'s message, or returns.</summary>
    public void Assert(IEnumerable<string?> recipientFields, string operation)
    {
        string? refusal = RefusalFor(recipientFields, operation);
        if (refusal != null)
        {
            throw new InvalidOperationException(refusal);
        }
    }
}

/// <summary>
/// Process-wide access to the <see cref="RecipientAllowlist"/> of this machine, from its live-test
/// settings: armed only where <see cref="LiveWriteAccess.ArmsRecipientAllowlist"/> says so (the Exchange
/// VM's test account), null everywhere else - where the guards already in place are the whole story.
/// </summary>
public static class LiveRecipientGuard
{
    private static readonly Lazy<RecipientAllowlist?> CurrentAllowlist = new(
        () => For(LiveTestSettings.Load()), LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>This process's allowlist, or null when this machine does not arm one.</summary>
    public static RecipientAllowlist? Current => CurrentAllowlist.Value;

    /// <summary>The allowlist <paramref name="settings"/> arm, or null. Pure: used by T1.</summary>
    public static RecipientAllowlist? For(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return LiveWriteAccess.ArmsRecipientAllowlist(settings.MachineProfile)
            ? new RecipientAllowlist(new[] { settings.TestHubStoreDisplayName }, LiveRunMarker.Current)
            : null;
    }

    /// <summary>
    /// An address for a test that needs one distinct from the hub's: where the allowlist is armed, the hub's
    /// plus address carrying <paramref name="tag"/>; elsewhere <c>tag@example.invalid</c>, as the tests always
    /// used - a domain that resolves nowhere, on machines whose mail ends in a loopback sink.
    /// </summary>
    public static string AddressFor(string tag)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tag);
        RecipientAllowlist? armed = Current;
        return armed == null ? tag + "@example.invalid" : armed.PlusAddress(tag);
    }

    /// <summary>Throws when this machine arms an allowlist and a field addresses anyone but the hub.</summary>
    public static void Assert(IEnumerable<string?> recipientFields, string operation)
    {
        Current?.Assert(recipientFields, operation);
    }
}
