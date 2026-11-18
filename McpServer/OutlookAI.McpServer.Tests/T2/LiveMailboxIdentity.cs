using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>One store entry of the profile and the mailbox it is - content-free.</summary>
/// <param name="DisplayName">The profile's name for the entry (<c>Store.DisplayName</c>).</param>
/// <param name="ExchangeStoreType">Raw <c>OlExchangeStoreType</c>: 0 primary, 1 delegate, 2 public folders, 3 not Exchange, 4 an additional account's mailbox.</param>
/// <param name="IsCachedExchange"><c>Store.IsCachedExchange</c>.</param>
/// <param name="RecordKeyHex">The store's <c>PR_STORE_RECORD_KEY</c> in hex - the MAILBOX, however many entries the profile holds for it.</param>
public sealed record MailboxStoreEntry(string DisplayName, int? ExchangeStoreType, bool? IsCachedExchange, string? RecordKeyHex);

/// <summary>One account of the profile and the store entry its mail is delivered to.</summary>
/// <param name="DisplayName"><c>Account.DisplayName</c> - an Exchange account's is its address.</param>
/// <param name="DeliveryStoreDisplayName">Its <c>DeliveryStore.DisplayName</c>.</param>
/// <param name="DeliveryStoreRecordKeyHex">Its delivery store's <c>PR_STORE_RECORD_KEY</c> in hex.</param>
public sealed record MailboxAccountEntry(string DisplayName, string? DeliveryStoreDisplayName, string? DeliveryStoreRecordKeyHex);

/// <summary>What one read of the profile found: every store entry and every account.</summary>
public sealed record MailboxIdentityReading(IReadOnlyList<MailboxStoreEntry> Stores, IReadOnlyList<MailboxAccountEntry> Accounts);

/// <summary>The verdict on a <see cref="MailboxIdentityReading"/>: the hub's mailbox, its aliases, or a refusal.</summary>
public sealed class MailboxIdentityReport
{
    /// <summary>Why the tier may not run, or null when the identities line up.</summary>
    public string? Refusal { get; init; }

    /// <summary>A short, content-free name for the hub's mailbox: the first 8 hex digits of the SHA-256 of its record key.</summary>
    public string HubKeyShort { get; init; } = "-";

    /// <summary>The hub entry's Exchange store type and cached state, for the report line.</summary>
    public string HubEntry { get; init; } = "-";

    /// <summary>Every OTHER entry of the profile that is the hub's mailbox - an automapped copy, say. Never censused, never written through.</summary>
    public IReadOnlyList<MailboxStoreEntry> HubAliases { get; init; } = Array.Empty<MailboxStoreEntry>();

    /// <summary>How many distinct mailboxes the watched stores are.</summary>
    public int WatchedMailboxes { get; init; }

    /// <summary>One line for the run's log.</summary>
    public string Describe()
    {
        if (Refusal != null)
        {
            return Refusal;
        }

        string aliases = HubAliases.Count == 0
            ? "no other entry of the profile is that mailbox"
            : "the same mailbox is also entry " + string.Join(", ", HubAliases.Select(a =>
                "'" + a.DisplayName + "' (ExchangeStoreType " + (a.ExchangeStoreType?.ToString(CultureInfo.InvariantCulture) ?? "-")
                + ", " + (a.IsCachedExchange == true ? "cached" : "online") + ")"))
              + " - one mailbox: never censused as a second one, never written through";
        return "the hub is the test account's own store (mailbox " + HubKeyShort + ", " + HubEntry + "); " + aliases
            + "; the watched stores are " + WatchedMailboxes.ToString(CultureInfo.InvariantCulture) + " distinct mailboxes";
    }
}

/// <summary>
/// The Exchange VM's test mailbox by IDENTITY, not by name (2026-11-17). The profile there holds the test
/// account's mailbox twice: its own cached store, named by its address (ExchangeStoreType 4), and an online
/// entry Exchange automapped through the real mailbox's Full Access, named by the test account's display name
/// (ExchangeStoreType 1). Both carry the same <c>PR_STORE_RECORD_KEY</c>, and the bystander's differs - measured
/// on the VM that day. The coordinator's direction, the maintainer's question still open: the two entries are
/// ONE test mailbox; writes are allowlisted by the test account's identity, never by a store's name; nothing is
/// ever written through the online entry; and the tripwire never counts one mailbox as two stores.
/// <para>
/// So before the baseline census, on the Exchange VM only, <see cref="Require"/> reads every entry's record key
/// and every account's delivery store, and REFUSES the tier unless: the account named as the hub delivers to
/// exactly the store entry the settings name as the hub, and that entry is the account's own (not a delegate
/// entry); no store the tripwire watches or the settings declare a bystander is an alias of the hub's mailbox;
/// and no two watched stores are one mailbox. The write allowlist keeps writing by the hub's NAME - the name is
/// what every write call passes - and this check proves, before the first write, that the name is the test
/// account's own store and nothing else.
/// </para>
/// </summary>
public static class LiveMailboxIdentity
{
    /// <summary>The phrase every refusal carries.</summary>
    public const string Refused = "MAILBOX IDENTITY REFUSED";

    /// <summary>The value of <c>OlExchangeStoreType.olExchangeMailbox</c>: a delegate - another user's mailbox opened through this profile.</summary>
    public const int DelegateEntry = 1;

    private static readonly object Gate = new();
    private static MailboxIdentityReport? _current;

    /// <summary>The report <see cref="Require"/> made in this process, or null before it ran (or off the Exchange VM).</summary>
    public static MailboxIdentityReport? Current
    {
        get
        {
            lock (Gate)
            {
                return _current;
            }
        }
    }

    /// <summary>True on the machine profiles this check is for - the Exchange VM's.</summary>
    public static bool AppliesTo(LiveMachineProfile profile)
    {
        return profile is LiveMachineProfile.ExchangeTestAccount or LiveMachineProfile.ExchangeGuest;
    }

    /// <summary>
    /// Reads the profile's identities (read-only) and refuses the tier unless they line up - on the Exchange VM;
    /// every other machine returns at once.
    /// </summary>
    public static MailboxIdentityReport? Require(LiveTestSettings settings, IReadOnlyList<string> watched, Action<string> report)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(watched);
        ArgumentNullException.ThrowIfNull(report);
        if (!AppliesTo(settings.MachineProfile))
        {
            return null;
        }

        MailboxIdentityReading reading = LiveOutlookTestMailer.ReadMailboxIdentities();
        MailboxIdentityReport verdict = Judge(reading, settings.TestHubStoreDisplayName, watched, settings.BystanderStoreDisplayNames);
        lock (Gate)
        {
            _current = verdict;
        }

        report("[tripwire] mailbox identity: " + verdict.Describe() + ".");
        if (verdict.Refusal != null)
        {
            throw new InvalidOperationException("REFUSING to run the live tier: " + verdict.Refusal);
        }

        return verdict;
    }

    /// <summary>Pure: the verdict on one reading. See the class for the rules.</summary>
    public static MailboxIdentityReport Judge(
        MailboxIdentityReading reading,
        string hub,
        IReadOnlyList<string> watched,
        IReadOnlyList<string> bystanders)
    {
        ArgumentNullException.ThrowIfNull(reading);
        ArgumentNullException.ThrowIfNull(watched);
        ArgumentNullException.ThrowIfNull(bystanders);
        if (string.IsNullOrWhiteSpace(hub))
        {
            return Refuse("no hub is named, so no mailbox can be the test account's");
        }

        List<MailboxAccountEntry> hubAccounts = reading.Accounts.Where(a => Same(a.DisplayName, hub)).ToList();
        if (hubAccounts.Count != 1)
        {
            return Refuse(hubAccounts.Count.ToString(CultureInfo.InvariantCulture) + " accounts of the profile are named '" + hub
                + "' - the hub must be exactly one account's own mailbox");
        }

        MailboxAccountEntry account = hubAccounts[0];
        string? hubKey = Normal(account.DeliveryStoreRecordKeyHex);
        if (hubKey == null)
        {
            return Refuse("the record key of account '" + hub + "''s delivery store could not be read, so the hub has no identity to check");
        }

        if (!Same(account.DeliveryStoreDisplayName, hub))
        {
            return Refuse("account '" + hub + "' delivers to the store entry '" + (account.DeliveryStoreDisplayName ?? "?")
                + "', not to the entry the settings name as the hub");
        }

        List<MailboxStoreEntry> named = reading.Stores.Where(s => Same(s.DisplayName, hub)).ToList();
        if (named.Count != 1)
        {
            return Refuse(named.Count.ToString(CultureInfo.InvariantCulture) + " store entries are named '" + hub
                + "' - a write by that name must reach exactly one");
        }

        MailboxStoreEntry hubEntry = named[0];
        if (!string.Equals(Normal(hubEntry.RecordKeyHex), hubKey, StringComparison.Ordinal))
        {
            return Refuse("the store entry named '" + hub + "' is not the mailbox account '" + hub + "' delivers to (their record keys differ)");
        }

        if (hubEntry.ExchangeStoreType == DelegateEntry)
        {
            return Refuse("the store entry named '" + hub + "' is a delegate entry (ExchangeStoreType 1) - writes go only through the account's own store");
        }

        List<MailboxStoreEntry> aliases = reading.Stores
            .Where(s => !ReferenceEquals(s, hubEntry) && string.Equals(Normal(s.RecordKeyHex), hubKey, StringComparison.Ordinal))
            .ToList();
        foreach (MailboxStoreEntry alias in aliases)
        {
            if (watched.Any(w => Same(w, alias.DisplayName)))
            {
                return Refuse("the tripwire would watch '" + alias.DisplayName + "', which is the hub's mailbox under another entry - one mailbox counted as two stores, whose own churn would read as a loss");
            }

            if (bystanders.Any(b => Same(b, alias.DisplayName)))
            {
                return Refuse("'" + alias.DisplayName + "' is declared a bystander, but it is the hub's mailbox under another entry");
            }
        }

        Dictionary<string, string> seen = new(StringComparer.Ordinal);
        foreach (string store in watched)
        {
            List<MailboxStoreEntry> entries = reading.Stores.Where(s => Same(s.DisplayName, store)).ToList();
            if (entries.Count != 1)
            {
                return Refuse(entries.Count.ToString(CultureInfo.InvariantCulture) + " store entries are named '" + store
                    + "', which the tripwire watches - it must census exactly one");
            }

            string? key = Normal(entries[0].RecordKeyHex);
            if (key == null)
            {
                return Refuse("the record key of watched store '" + store + "' could not be read, so it cannot be told apart from the others");
            }

            if (seen.TryGetValue(key, out string? first))
            {
                return Refuse("the watched stores '" + first + "' and '" + store + "' are one mailbox - the tripwire would count it twice");
            }

            seen[key] = store;
        }

        foreach (string bystander in bystanders)
        {
            MailboxStoreEntry? entry = reading.Stores.FirstOrDefault(s => Same(s.DisplayName, bystander));
            if (entry != null && string.Equals(Normal(entry.RecordKeyHex), hubKey, StringComparison.Ordinal))
            {
                return Refuse("bystander '" + bystander + "' is the hub's own mailbox");
            }
        }

        return new MailboxIdentityReport
        {
            HubKeyShort = Short(hubKey),
            HubEntry = "ExchangeStoreType " + (hubEntry.ExchangeStoreType?.ToString(CultureInfo.InvariantCulture) ?? "-")
                + ", " + (hubEntry.IsCachedExchange == true ? "cached" : "not cached"),
            HubAliases = aliases,
            WatchedMailboxes = seen.Count,
        };
    }

    private static MailboxIdentityReport Refuse(string why)
    {
        return new MailboxIdentityReport { Refusal = Refused + ": " + why + "." };
    }

    private static bool Same(string? a, string? b)
    {
        return a != null && b != null && string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static string? Normal(string? hex)
    {
        return string.IsNullOrWhiteSpace(hex) ? null : hex.Trim().ToUpperInvariant();
    }

    /// <summary>The first 8 hex digits of the SHA-256 of a record key: enough to tell mailboxes apart in a log, and nothing more.</summary>
    public static string Short(string recordKeyHex)
    {
        byte[] hash = SHA256.HashData(Encoding.ASCII.GetBytes(recordKeyHex.Trim().ToUpperInvariant()));
        return Convert.ToHexString(hash, 0, 4).ToLowerInvariant();
    }
}
