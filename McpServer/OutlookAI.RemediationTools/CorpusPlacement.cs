using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>
/// How a corpus item is made to LIVE in the folder the plan asked for, in the order the
/// builder tries them.
/// </summary>
public enum CorpusPlacementMethod
{
    /// <summary>
    /// Nothing worked: an item ends up wherever Outlook decides to file it. Never chosen by
    /// <see cref="CorpusPlacement.Choose"/>; it is the value that means "refuse".
    /// </summary>
    None = 0,

    /// <summary>
    /// Create in the target folder, save, clear MSGFLAG_UNSENT (and set MSGFLAG_READ to the
    /// plan's read state) through the PropertyAccessor, save again. Cheapest rung by far -
    /// no move, so no second write of the whole item - which is why it is tried first.
    /// </summary>
    InPlaceWithSentFlag = 1,

    /// <summary>
    /// Create in Drafts, save, clear MSGFLAG_UNSENT, save, then <c>Move</c> to the target
    /// folder. For a store that insists on filing an item where it thinks it belongs at
    /// creation time, moving an item that is no longer unsent may be the only way it stays.
    /// <para>
    /// A move issues the item a NEW EntryID, so the builder records the post-move id - the
    /// manifest is the teardown allowlist, and an id recorded before a move would name
    /// nothing.
    /// </para>
    /// </summary>
    DraftsThenMoveWithSentFlag = 2,

    /// <summary>
    /// Create in Drafts, save, move to the target folder, without touching the message
    /// flags. Distinguishes "the move is what places it" from "the flag is what places it";
    /// without both rungs a passing probe would not say which half did the work.
    /// </summary>
    DraftsThenMove = 3,

    /// <summary>
    /// Create in the target folder, save, and nothing else. This is what the first real
    /// build did, and it is kept as a CONTROL rung rather than dropped: the probe should
    /// record what the store does with a plain saved item rather than assume the failure
    /// that was observed once generalises.
    /// </summary>
    InPlaceOnly = 4,
}

/// <summary>The outcome of one attempt to place a throwaway item in a target folder.</summary>
/// <param name="Method">Which rung this probe exercised.</param>
/// <param name="TargetFolderName">The folder the item was supposed to end up in.</param>
/// <param name="ParentIsTargetFolder">Whether the saved item's Parent is that folder, compared by folder EntryID.</param>
/// <param name="TargetFolderTableContainsIt">
/// Whether a <c>GetTable</c> on the target folder returned the item. The decisive signal:
/// the freshness sweep enumerates a folder through its table, so an item the table does not
/// carry does not exist as far as the measurement is concerned.
/// </param>
/// <param name="SentFlagSet">MailItem.Sent after the write - informational, not required.</param>
/// <param name="LandedInFolderName">Where the item actually ended up, so a failure says where it went.</param>
/// <param name="Error">Why the rung failed, when it did.</param>
public sealed record CorpusPlacementProbe(
    CorpusPlacementMethod Method,
    string TargetFolderName,
    bool ParentIsTargetFolder,
    bool TargetFolderTableContainsIt,
    bool SentFlagSet,
    string? LandedInFolderName,
    string? Error);

/// <summary>
/// Decides, from probe results alone, whether corpus items can be made to live where the
/// plan puts them - and refuses to build a corpus that cannot be measured if they cannot.
/// <para>
/// <b>Why this exists, and it is not a refinement.</b> The first real build created 40 000
/// items with <c>Items.Add</c> + <c>Save()</c> into Inbox, Sent Items, Deleted Items and
/// Junk Email. All 40 000 ended up in DRAFTS. An item created that way is an UNSENT message
/// (MSGFLAG_UNSENT), and Outlook files unsent messages as drafts regardless of the folder
/// they were added to. The freshness sweep covers Inbox, Sent Items, Deleted Items and Junk
/// Email and does not cover Drafts, so a sweep run against those 40 000 items selected SIX
/// of them in 234-367 ms. The corpus existed and the measurement it exists for could not be
/// taken.
/// </para>
/// <para>
/// <b>The design fault behind it, which is the part worth remembering.</b> The flag write
/// that would have cleared MSGFLAG_UNSENT was bundled into the DATE ladder as one rung of
/// <c>CorpusDateFidelity</c>. When the date probe correctly refused, the operator passed
/// <c>--allow-undated</c> - and that silently disabled the flag write too, because the two
/// had been coupled for no better reason than that both used the PropertyAccessor. Placement
/// is now probed and decided on its own, so a decision about dates cannot quietly become a
/// decision about where items live.
/// </para>
/// </summary>
public static class CorpusPlacement
{
    /// <summary>The rungs, cheapest-that-could-work first. The builder takes the first that fully verifies.</summary>
    public static readonly CorpusPlacementMethod[] Ladder =
    {
        CorpusPlacementMethod.InPlaceWithSentFlag,
        CorpusPlacementMethod.DraftsThenMoveWithSentFlag,
        CorpusPlacementMethod.DraftsThenMove,
        CorpusPlacementMethod.InPlaceOnly,
    };

    /// <summary>Whether this rung creates the item in Drafts rather than in the target folder.</summary>
    public static bool CreatesInDrafts(CorpusPlacementMethod method)
        => method is CorpusPlacementMethod.DraftsThenMoveWithSentFlag or CorpusPlacementMethod.DraftsThenMove;

    /// <summary>Whether this rung moves the item after saving it (and so changes its EntryID).</summary>
    public static bool RequiresMove(CorpusPlacementMethod method)
        => method is CorpusPlacementMethod.DraftsThenMoveWithSentFlag or CorpusPlacementMethod.DraftsThenMove;

    /// <summary>Whether this rung writes PR_MESSAGE_FLAGS to clear MSGFLAG_UNSENT.</summary>
    public static bool WritesSentFlag(CorpusPlacementMethod method)
        => method is CorpusPlacementMethod.InPlaceWithSentFlag or CorpusPlacementMethod.DraftsThenMoveWithSentFlag;

    /// <summary>
    /// A rung counts as usable only when the item's parent IS the target folder AND the
    /// target folder's TABLE returns it. Both, because they can disagree: an item can be
    /// parented correctly and still be absent from a table the sweep would read, and that
    /// second case is the one that would produce a corpus which looks right in Outlook and
    /// measures as empty.
    /// <para>
    /// <c>SentFlagSet</c> deliberately does NOT gate this. A still-unsent item that
    /// nonetheless sits in the Inbox and appears in its table is perfectly measurable; the
    /// flag is a means, not the goal, and requiring it would reject a working rung.
    /// </para>
    /// </summary>
    public static bool IsUsable(CorpusPlacementProbe probe)
    {
        ArgumentNullException.ThrowIfNull(probe);
        return probe.Error == null && probe.ParentIsTargetFolder && probe.TargetFolderTableContainsIt;
    }

    /// <summary>
    /// The rung to build with: the first in <see cref="Ladder"/> order that fully verified.
    /// <see cref="CorpusPlacementMethod.None"/> when none did.
    /// </summary>
    public static CorpusPlacementMethod Choose(IReadOnlyCollection<CorpusPlacementProbe> probes)
    {
        ArgumentNullException.ThrowIfNull(probes);
        foreach (CorpusPlacementMethod method in Ladder)
        {
            foreach (CorpusPlacementProbe probe in probes)
            {
                if (probe.Method == method && IsUsable(probe))
                {
                    return method;
                }
            }
        }

        return CorpusPlacementMethod.None;
    }

    /// <summary>
    /// Whether the build may proceed, and what to print either way.
    /// <para>
    /// The message states the consequence as a NUMBER rather than as a description. That is
    /// deliberate and it is a correction: the date guard's original refusal text described
    /// its consequence in prose ("every item would carry a received time of roughly now"),
    /// an operator drew a reasonable conclusion from it that happened to be wrong, overrode
    /// the guard, and lost a 12-minute build. "A freshness sweep will select 0 of 40 000
    /// items" cannot be reasoned around.
    /// </para>
    /// </summary>
    public static (bool Proceed, string Message) Decide(
        CorpusPlacementMethod chosen, bool allowDraftsPlacement, int itemCount)
    {
        if (chosen != CorpusPlacementMethod.None)
        {
            string extra = RequiresMove(chosen)
                ? " Items are moved after saving, so each one is written twice and the build is correspondingly "
                    + "slower; the manifest records the POST-move EntryID."
                : string.Empty;
            return (true, $"Placement: VERIFIED via {chosen}. Items will live in the folders the plan names." + extra);
        }

        string what = "Placement: NOT ACHIEVABLE on this store. No method left an item in the folder it was meant "
            + $"for and visible in that folder's table; every item would be filed as a draft instead. A freshness "
            + $"sweep covers Inbox, Sent Items, Deleted Items and Junk Email and NOT Drafts, so it would select "
            + $"0 of {itemCount.ToString("N0", CultureInfo.InvariantCulture)} items.";

        return allowDraftsPlacement
            ? (true, what + " Proceeding because --allow-drafts-placement was given. This corpus can still be used "
                + "for out-of-band per-item and per-folder timing (measurement plan step 2) and for exhaustive-scan "
                + "throughput against a named folder; it CANNOT be used to measure the freshness sweep at all.")
            : (false, what + " Refusing to build. The sweep measurement is what this corpus exists for, and a corpus "
                + "the sweep cannot see would look exactly like a good one from the outside.");
    }
}
