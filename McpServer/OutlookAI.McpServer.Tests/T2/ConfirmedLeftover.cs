using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The one write the Exchange VM's bystander mailbox ever gets (Q138): the one item carrying the live-tier tag
/// that the workstation years left in its Sent Items, removed after the maintainer confirmed it by its EntryID
/// (2026-11-18). Nothing else in that mailbox is ever written - by this or by anything.
/// <para>
/// <b>What selects it: its EntryID AND its subject, both exactly</b> - AGENTS.md, Mailbox Safety rule 2, an
/// EntryID allowlist AND an ordinal tag match. The EntryID is pinned by its SHA-256, so that this public
/// repository carries no identifier of the real mailbox; the subject is test data the live tier itself once
/// wrote, pinned as it is, and starts with <see cref="LiveOutlookTestMailer.SubjectTag"/>. Nothing is chosen
/// by a subject pattern: the server-side tag restriction only narrows which rows are read, and every row is
/// judged by <see cref="IsTheItem"/>.
/// </para>
/// <para>
/// <b>How it goes</b> (<see cref="LiveOutlookTestMailer.RemoveConfirmedLeftover"/>): moved to the store's own
/// Deleted Items, the copy the move hands back checked again (<see cref="RefusalForMovedCopy"/>), then deleted
/// there, which takes it out of the owner's sight; Exchange keeps it in the hidden Recoverable Items folder for
/// its retention period. Only on the Exchange VM's test-account profile, only in its one declared bystander
/// (<see cref="RefusalFor"/>), and only inside the suite, whose count tripwire censuses the bystander item by
/// item around it and accepts this item's departure and no other (<see cref="ConfirmedRemovals"/>).
/// </para>
/// </summary>
public sealed class ConfirmedLeftover
{
    private ConfirmedLeftover(string decision, int defaultFolderId, string subject, string entryIdSha256)
    {
        Decision = decision;
        DefaultFolderId = defaultFolderId;
        Subject = subject;
        EntryIdSha256 = entryIdSha256;
    }

    /// <summary>
    /// Q138's leftover: in the bystander's Sent Items, sent 2026-07-26 by a live run on the maintainer's
    /// workstation, found by the Exchange VM's read-only count on 2026-10-03, and confirmed by the maintainer
    /// on 2026-11-18 by its 70-byte EntryID, whose SHA-256 this is.
    /// </summary>
    public static ConfirmedLeftover Q138 { get; } = Create(
        "Q138",
        defaultFolderId: 5,
        subject: "[OutlookAI-McpTest] d3994a2ad25309442 d39 move seed",
        entryIdSha256: "e9c3c7c43ba9729ba939b22fb0ab401b9c4ce2af07a8ac828ff2afb3711033fe");

    /// <summary>The decision that confirmed it.</summary>
    public string Decision { get; }

    /// <summary>The <c>OlDefaultFolders</c> number of the folder it was confirmed in (5, Sent Items).</summary>
    public int DefaultFolderId { get; }

    /// <summary>Its subject, exactly.</summary>
    public string Subject { get; }

    /// <summary>The SHA-256 of its EntryID's hex form in upper case (<see cref="HashEntryId"/>), lower-case hex.</summary>
    public string EntryIdSha256 { get; }

    /// <summary>
    /// A pin, refused unless the subject starts with the live-tier tag - an item without it is not test data -
    /// and the hash is a SHA-256 in lower-case hex.
    /// </summary>
    public static ConfirmedLeftover Create(string decision, int defaultFolderId, string subject, string entryIdSha256)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(decision);
        ArgumentNullException.ThrowIfNull(subject);
        ArgumentNullException.ThrowIfNull(entryIdSha256);
        if (!subject.StartsWith(LiveOutlookTestMailer.SubjectTag, StringComparison.Ordinal))
        {
            throw new ArgumentException("A confirmed leftover's subject must start with " + LiveOutlookTestMailer.SubjectTag
                + ": an item without the tag is not test data, and nothing but test data is ever removed.", nameof(subject));
        }

        if (entryIdSha256.Length != 64 || entryIdSha256.Any(c => !(c is (>= '0' and <= '9') or (>= 'a' and <= 'f'))))
        {
            throw new ArgumentException("The EntryID's pin must be a SHA-256 in lower-case hex.", nameof(entryIdSha256));
        }

        return new ConfirmedLeftover(decision, defaultFolderId, subject, entryIdSha256);
    }

    /// <summary>
    /// The SHA-256 of an EntryID's hex form, trimmed and in upper case - the form Outlook hands out - as
    /// lower-case hex. A value that is not hex still hashes, and then matches nothing.
    /// </summary>
    public static string HashEntryId(string entryIdHex)
    {
        ArgumentNullException.ThrowIfNull(entryIdHex);
        byte[] digest = SHA256.HashData(Encoding.ASCII.GetBytes(entryIdHex.Trim().ToUpperInvariant()));
        return Convert.ToHexString(digest).ToLowerInvariant();
    }

    /// <summary>True only for this item: its EntryID's hash AND its subject, ordinally, tag included.</summary>
    public bool IsTheItem(string? entryIdHex, string? subject)
    {
        return !string.IsNullOrWhiteSpace(entryIdHex)
            && subject != null
            && subject.StartsWith(LiveOutlookTestMailer.SubjectTag, StringComparison.Ordinal)
            && string.Equals(subject, Subject, StringComparison.Ordinal)
            && string.Equals(HashEntryId(entryIdHex), EntryIdSha256, StringComparison.Ordinal);
    }

    /// <summary>
    /// Which of a folder's tagged rows is this item. Exactly one is <see cref="ConfirmedLeftoverOutcome.Found"/>;
    /// none is <see cref="ConfirmedLeftoverOutcome.NotPresent"/>; more than one cannot happen - an EntryID names
    /// one item - and is <see cref="ConfirmedLeftoverOutcome.Ambiguous"/>, which removes nothing. Every other
    /// row is only counted.
    /// </summary>
    public ConfirmedLeftoverSelection Select(IEnumerable<(string? EntryId, string? Subject)> taggedRows)
    {
        ArgumentNullException.ThrowIfNull(taggedRows);
        List<string> matches = new();
        int others = 0;
        foreach ((string? entryId, string? subject) in taggedRows)
        {
            if (IsTheItem(entryId, subject))
            {
                matches.Add(entryId!.Trim());
            }
            else
            {
                others++;
            }
        }

        ConfirmedLeftoverOutcome outcome = matches.Count switch
        {
            0 => ConfirmedLeftoverOutcome.NotPresent,
            1 => ConfirmedLeftoverOutcome.Found,
            _ => ConfirmedLeftoverOutcome.Ambiguous,
        };
        return new ConfirmedLeftoverSelection(outcome, matches.Count == 1 ? matches[0] : null, others);
    }

    /// <summary>
    /// Whether this run may remove it from <paramref name="store"/> at all: only on the Exchange VM's
    /// test-account profile, whose write allowlist grants the hub alone; only when the settings declare exactly
    /// one bystander and <paramref name="store"/> is that one; and never the hub. Null when it may.
    /// </summary>
    public string? RefusalFor(LiveMachineProfile profile, IReadOnlyList<string> bystanders, string hub, string store)
    {
        ArgumentNullException.ThrowIfNull(bystanders);
        string prefix = "REFUSING to remove " + Decision + "'s confirmed leftover: ";
        if (profile != LiveMachineProfile.ExchangeTestAccount)
        {
            return prefix + "the machine's profile is " + profile + ", and only the Exchange VM's "
                + LiveMachineProfile.ExchangeTestAccount + " profile ever removes it.";
        }

        if (bystanders.Count != 1)
        {
            return prefix + "the settings declare " + bystanders.Count.ToString(CultureInfo.InvariantCulture)
                + " bystander(s), and it lives in the one bystander of the Exchange VM.";
        }

        if (string.IsNullOrWhiteSpace(store) || !string.Equals(store, bystanders[0], StringComparison.OrdinalIgnoreCase))
        {
            return prefix + "'" + store + "' is not the declared bystander.";
        }

        if (string.Equals(store, hub, StringComparison.OrdinalIgnoreCase))
        {
            return prefix + "'" + store + "' is the hub, which a bystander never is.";
        }

        return null;
    }

    /// <summary>
    /// Whether the copy the move into Deleted Items handed back is still this item - the same subject,
    /// ordinally, in that store's Deleted Items, under a new EntryID that was read - before it is deleted
    /// there. Null when it is.
    /// </summary>
    public string? RefusalForMovedCopy(string? movedEntryId, string? movedSubject, bool inDeletedItems, string originalEntryId)
    {
        ArgumentNullException.ThrowIfNull(originalEntryId);
        string prefix = "REFUSING to delete the moved copy of " + Decision + "'s confirmed leftover from Deleted Items: ";
        if (string.IsNullOrWhiteSpace(movedEntryId))
        {
            return prefix + "the move handed back no EntryID to delete it by.";
        }

        if (!string.Equals(movedSubject, Subject, StringComparison.Ordinal))
        {
            return prefix + "the moved copy's subject is not the confirmed subject.";
        }

        if (!inDeletedItems)
        {
            return prefix + "the moved copy is not in the store's Deleted Items.";
        }

        if (string.Equals(movedEntryId.Trim(), originalEntryId.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return prefix + "its EntryID did not change with the move, so it cannot be told from the original.";
        }

        return null;
    }

    /// <summary>One line for a log: the decision, the folder, the subject, and the hash's first twelve characters.</summary>
    public string Describe()
    {
        return Decision + "'s confirmed leftover - default folder " + DefaultFolderId.ToString(CultureInfo.InvariantCulture)
            + ", subject '" + Subject + "', EntryID SHA-256 " + EntryIdSha256[..12] + "...";
    }
}

/// <summary>What <see cref="ConfirmedLeftover.Select"/> found in a folder.</summary>
public enum ConfirmedLeftoverOutcome
{
    /// <summary>No row is the item.</summary>
    NotPresent = 0,

    /// <summary>Exactly one row is the item.</summary>
    Found = 1,

    /// <summary>More than one row matched - impossible for one EntryID, so nothing is removed.</summary>
    Ambiguous = 2,
}

/// <summary>A folder's reading: the outcome, the item's EntryID when found, and how many other tagged rows there were.</summary>
/// <param name="Outcome">Whether the item is there, once.</param>
/// <param name="EntryId">Its EntryID when <see cref="ConfirmedLeftoverOutcome.Found"/>, else null.</param>
/// <param name="OtherTagged">Tagged rows that are not the item - read, never touched.</param>
public sealed record ConfirmedLeftoverSelection(ConfirmedLeftoverOutcome Outcome, string? EntryId, int OtherTagged);

/// <summary>What <see cref="LiveOutlookTestMailer.RemoveConfirmedLeftover"/> did, content-free.</summary>
/// <param name="Removed">True when this call moved the item to Deleted Items and deleted it there.</param>
/// <param name="Gone">True when, read back afterwards, no row of its folder is the item and no row of Deleted Items carries its subject.</param>
/// <param name="OtherTaggedInFolder">Tagged rows of its folder that are not the item - read, never touched.</param>
/// <param name="CopiesInDeletedItems">Rows of Deleted Items carrying its subject, read back at the end.</param>
public sealed record ConfirmedLeftoverRemoval(bool Removed, bool Gone, int OtherTaggedInFolder, int CopiesInDeletedItems)
{
    /// <summary>One line for the test's output.</summary>
    public string Describe(ConfirmedLeftover leftover)
    {
        ArgumentNullException.ThrowIfNull(leftover);
        string what = Removed ? "moved to Deleted Items and deleted there" : "not present - nothing removed";
        return leftover.Describe() + ": " + what + "; read back: gone=" + Gone + ", other tagged in its folder="
            + OtherTaggedInFolder.ToString(CultureInfo.InvariantCulture) + ", copies in Deleted Items="
            + CopiesInDeletedItems.ToString(CultureInfo.InvariantCulture) + ".";
    }
}

/// <summary>
/// The confirmed removals this process made, by store and the EntryID each item had where it was confirmed -
/// written by <see cref="LiveOutlookTestMailer.RemoveConfirmedLeftover"/> alone, once the item reads back gone.
/// The count tripwire reads it (<see cref="StoreCountTripwire.Evaluate"/>): such an item's departure from its
/// folder is noted, not failed, and only when the census found it TAGGED; every other departure fails exactly
/// as before. The artifact sweep's pre-run counts of the read-only stores come down by it too
/// (<see cref="LiveArtifactBaseline.AdjustedFor"/>).
/// </summary>
public static class ConfirmedRemovals
{
    private static readonly object Gate = new();
    private static readonly List<ConfirmedRemoval> Removed = new();

    /// <summary>The removals so far, oldest first; empty in every run that made none.</summary>
    public static IReadOnlyList<ConfirmedRemoval> Current
    {
        get
        {
            lock (Gate)
            {
                return Removed.ToList();
            }
        }
    }

    /// <summary>Records one removal. Only the confirmed-leftover removal calls it, after reading the item back gone.</summary>
    internal static void Record(ConfirmedRemoval removal)
    {
        ArgumentNullException.ThrowIfNull(removal);
        lock (Gate)
        {
            Removed.Add(removal);
        }
    }
}

/// <summary>One confirmed removal: the store, and the EntryID the item had in the folder it was confirmed in.</summary>
/// <param name="Store">The store's display name, as the settings name the bystander.</param>
/// <param name="EntryId">The item's EntryID in the folder it was confirmed in - the id the census knows it by.</param>
public sealed record ConfirmedRemoval(string Store, string EntryId);
