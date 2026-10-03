using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The decisions of the live created-folder proof (Q96 (iv), 2026-10-03), made where CI can pin
/// them: <see cref="LiveCreatedFolderTests"/> is <c>Category=Live</c> and runs only on a test guest,
/// so what it may conclude from each answer it gets is decided here, store names and folder lists
/// in, one failure sentence or null out. No COM, no settings file, no mailbox.
/// <para>
/// <b>What is being proven.</b> Since Q85 a draft tool that puts an item into a special folder a
/// POP3, IMAP or data-file mailbox lacks lets Outlook create the folder and REPORTS it in
/// <c>createdFolders</c>. Every T1 test of that runs against fakes; the guests' only writable store,
/// the hub, has a Drafts folder, so on a guest the "created" branch never ran. The throwaway data
/// file - attached by <c>Testbed/guest/Reset-ThrowawayStore.ps1</c> before every live run, holding
/// Deleted Items and nothing else - is where it runs: a reply to an item in it is filed in ITS
/// Drafts, which the product must create and report.
/// </para>
/// <para>
/// <b>Each judge returns the sentence the run fails with, or null.</b> A sentence names what was
/// seen, what it means and what to do, because the run that reads it is on a guest nobody is
/// watching.
/// </para>
/// </summary>
public static class ThrowawayStoreProof
{
    /// <summary>The script that recreates the throwaway data file - named in every remedy.</summary>
    public const string ResetScript = "Testbed/guest/Reset-ThrowawayStore.ps1 -Execute";

    /// <summary>What <see cref="LivePopulationCoverage.Require{T}"/> calls the population.</summary>
    public const string Population = "throwaway data file without a Drafts folder (throwawayStoreDisplayName)";

    /// <summary>What does not run when there is no throwaway data file.</summary>
    public const string WhatWouldNotRun =
        "the live proof that a reply into a data file with no Drafts folder reports the Drafts folder it creates (Q96 (iv))";

    /// <summary>How a machine gets the population.</summary>
    public const string Remedy =
        "Declare the guest's throwaway data file in 'throwawayStoreDisplayName' (Testbed/testbed.json, rendered by "
        + "Testbed/host/New-LiveTestSettings.ps1) and recreate it before the run with " + ResetScript
        + " (Testbed/README.md section 1, step 9a).";

    /// <summary>
    /// The precondition: before anything is written, the throwaway data file is mounted and has NO
    /// Drafts folder, by the product's own non-creating lookup.
    /// </summary>
    /// <param name="store">The throwaway data file's display name.</param>
    /// <param name="lookupError">
    /// What <c>OutlookComSession.TryGetDefaultFolderInfo(store, 16)</c> answered: null when it resolved a
    /// Drafts folder, otherwise its error - <c>DefaultFolderAbsent</c>, <c>DefaultFolderUnreadable</c>,
    /// <c>StoreNotFound</c> or a COM failure.
    /// </param>
    public static string? DraftsAbsentBefore(string store, string? lookupError)
    {
        if (lookupError == null)
        {
            return "The throwaway data file '" + store + "' already HAS a Drafts folder, so a reply into it creates "
                + "nothing and this proof would pass having proved nothing. It is recreated fresh - with Deleted Items "
                + "and nothing else - before every run: run " + ResetScript + " first.";
        }

        if (string.Equals(lookupError, "DefaultFolderAbsent", StringComparison.Ordinal))
        {
            return null;
        }

        if (string.Equals(lookupError, "StoreNotFound", StringComparison.Ordinal))
        {
            return "The throwaway data file '" + store + "' is declared in the live-test settings and is not mounted in "
                + "this profile. Attach it before the run with " + ResetScript + ".";
        }

        return "Whether the throwaway data file '" + store + "' has a Drafts folder could not be established without "
            + "asking Outlook for it, which would create one (" + lookupError + "). The proof refuses rather than guess; "
            + "recreate the store with " + ResetScript + " and run again.";
    }

    /// <summary>
    /// The reply: it was filed in the throwaway data file, in a folder of its own, and reported
    /// creating EXACTLY that folder - <c>store/path</c>, the form <c>archive_mail</c> uses
    /// (<see cref="SpecialFolders.CreatedFolderLabel"/>).
    /// </summary>
    /// <param name="store">The throwaway data file's display name.</param>
    /// <param name="createdFolders">The reply result's <c>createdFolders</c>.</param>
    /// <param name="draftStore">The store the result says the draft is in.</param>
    /// <param name="draftFolder">The folder the result says the draft is in.</param>
    public static string? ReplyReportedTheCreatedDrafts(
        string store, IReadOnlyList<string>? createdFolders, string? draftStore, string? draftFolder)
    {
        if (!string.Equals(draftStore, store, StringComparison.OrdinalIgnoreCase))
        {
            return "The reply was filed in '" + (draftStore ?? "(no store reported)") + "', not in the throwaway data file '"
                + store + "' its source lives in. A reply is filed in its SOURCE store's Drafts (Q85), so nothing about "
                + "creating that store's Drafts folder was exercised.";
        }

        if (string.IsNullOrEmpty(draftFolder))
        {
            return "The reply result names no folder for the draft, so what it reports creating cannot be checked against "
                + "where the draft is.";
        }

        string expected = SpecialFolders.CreatedFolderLabel(store, draftFolder!);
        if (createdFolders == null || createdFolders.Count == 0)
        {
            return "The reply was filed in '" + expected + "' in a data file that had NO Drafts folder before the call, and "
                + "its result reports creating nothing. That is the defect Q85 exists to prevent: a folder the product made "
                + "in the user's mailbox without saying so.";
        }

        if (createdFolders.Count != 1 || !string.Equals(createdFolders[0], expected, StringComparison.Ordinal))
        {
            return "The reply reports creating [" + string.Join(", ", createdFolders) + "], where exactly ['" + expected
                + "'] - the Drafts folder it was filed in - was expected.";
        }

        return null;
    }

    /// <summary>
    /// After the creation, the product's own non-creating lookup must SEE the new Drafts folder - the
    /// one <c>update_draft</c> and <c>discard_draft</c> gate on. Where Outlook designates a Drafts folder
    /// it makes in a data file is the question Q85 left open; this is where a guest answers it.
    /// </summary>
    /// <param name="store">The throwaway data file's display name.</param>
    /// <param name="lookupError">What <c>TryGetDefaultFolderInfo(store, 16)</c> answered: null when it resolved.</param>
    /// <param name="resolvedName">The name of the folder it resolved, when it did.</param>
    /// <param name="draftFolder">The folder the reply was filed in.</param>
    public static string? DraftsVisibleAfter(string store, string? lookupError, string? resolvedName, string? draftFolder)
    {
        if (lookupError != null)
        {
            return "The reply created and reported the Drafts folder of '" + store + "', but the product's own non-creating "
                + "lookup (SpecialFolders.Resolve) does not see it (" + lookupError + "): Outlook designated it somewhere "
                + "that lookup does not read - the question Q85 left open, answered. discard_draft and update_draft gate on "
                + "that lookup, so they refuse a draft the product itself just made there. Read where Outlook put the "
                + "designation before changing anything.";
        }

        if (!string.Equals(resolvedName, draftFolder, StringComparison.Ordinal))
        {
            return "The non-creating lookup resolves the Drafts folder of '" + store + "' as '" + (resolvedName ?? "(no name)")
                + "', while the reply was filed in '" + (draftFolder ?? "(no folder)") + "'.";
        }

        return null;
    }

    /// <summary>
    /// The discard: it moved the draft into the throwaway data file's Deleted Items, and reported
    /// creating a Deleted Items folder exactly when the store had none before - which a data file
    /// attached with <c>AddStoreEx</c> always has, so in practice it reports nothing.
    /// </summary>
    /// <param name="store">The throwaway data file's display name.</param>
    /// <param name="deletedItemsBefore">Whether the non-creating lookup found Deleted Items before the reply.</param>
    /// <param name="discarded">The discard result's <c>Discarded</c>.</param>
    /// <param name="createdFolders">The discard result's <c>createdFolders</c>.</param>
    public static string? DiscardReported(string store, bool deletedItemsBefore, bool discarded, IReadOnlyList<string>? createdFolders)
    {
        if (!discarded)
        {
            return "discard_draft did not report the draft discarded.";
        }

        if (deletedItemsBefore)
        {
            return createdFolders == null || createdFolders.Count == 0
                ? null
                : "discard_draft reports creating [" + string.Join(", ", createdFolders) + "] in '" + store + "', whose "
                    + "Deleted Items existed before the call - a creation claim nothing made true.";
        }

        return createdFolders != null && createdFolders.Count == 1
            && createdFolders[0].StartsWith(store + "/", StringComparison.Ordinal)
            ? null
            : "'" + store + "' had no Deleted Items before the discard, so discard_draft must report the one it created; "
                + "it reported [" + string.Join(", ", createdFolders ?? Array.Empty<string>()) + "].";
    }

    // ------------------------------------------------------------------ Q96 question 2 (c): the top-level comparison

    /// <summary>
    /// The before/after comparison of the throwaway data file's TOP-LEVEL folders around one call,
    /// against what the call reported (Q96 question 2 (c), decided 2026-10-03: in this test only).
    /// On SUCCESS the product names only the folder its lookup returned; it does not compare the
    /// top level the way a failed lookup's re-check does. This measures whether that is enough:
    /// every folder that appeared must be reported, as created or as having appeared, and nothing
    /// may be reported that did not appear. A folder that appeared and was not reported is the
    /// answer "the product needs the comparison on success too" - question 2's (b), the
    /// maintainer's to take.
    /// </summary>
    /// <param name="store">The throwaway data file's display name.</param>
    /// <param name="call">The tool the comparison is around, for the sentence.</param>
    /// <param name="before">Every folder path <c>list_folders</c> gave before the call (store-relative, '/'-separated).</param>
    /// <param name="after">The same, after the call.</param>
    /// <param name="createdFolders">The call's <c>createdFolders</c>.</param>
    /// <param name="appearedFolders">The call's <c>appearedFolders</c>.</param>
    public static string? TopLevelChangeReported(
        string store,
        string call,
        IReadOnlyList<string> before,
        IReadOnlyList<string> after,
        IReadOnlyList<string>? createdFolders,
        IReadOnlyList<string>? appearedFolders)
    {
        (IReadOnlyList<string> appeared, IReadOnlyList<string> vanished) = CompareTopLevel(before, after);
        List<string> reported = (createdFolders ?? Array.Empty<string>()).Concat(appearedFolders ?? Array.Empty<string>()).ToList();
        List<string> appearedLabels = appeared.Select(name => SpecialFolders.CreatedFolderLabel(store, name)).ToList();

        List<string> problems = new List<string>();
        if (vanished.Count > 0)
        {
            problems.Add("[" + string.Join(", ", vanished) + "] vanished from the top of '" + store + "' during " + call
                + " - nothing in this proof removes a folder.");
        }

        List<string> unproven = reported.Where(label => !appearedLabels.Contains(label, StringComparer.OrdinalIgnoreCase)).ToList();
        if (unproven.Count > 0)
        {
            problems.Add(call + " reports [" + string.Join(", ", unproven) + "], which did not appear at the top of '" + store
                + "' - a claim nothing made true.");
        }

        List<string> unreported = appearedLabels.Where(label => !reported.Contains(label, StringComparer.OrdinalIgnoreCase)).ToList();
        if (unreported.Count > 0)
        {
            problems.Add("[" + string.Join(", ", unreported) + "] appeared at the top of '" + store + "' during " + call
                + " and " + call + " reported none of them. On success the product names only the folder its lookup returned "
                + "(Q96 question 2): this run says it needs the top-level comparison on success too - the maintainer's (b) "
                + "of that question.");
        }

        return problems.Count == 0 ? null : string.Join(" ", problems);
    }

    /// <summary>The comparison as one line for the run's record, whatever it concluded.</summary>
    public static string DescribeTopLevelChange(
        string store,
        string call,
        IReadOnlyList<string> before,
        IReadOnlyList<string> after,
        IReadOnlyList<string>? createdFolders,
        IReadOnlyList<string>? appearedFolders)
    {
        (IReadOnlyList<string> appeared, IReadOnlyList<string> vanished) = CompareTopLevel(before, after);
        return "top level of '" + store + "' around " + call + ": before [" + string.Join(", ", TopLevel(before))
            + "]; after [" + string.Join(", ", TopLevel(after)) + "]; appeared [" + string.Join(", ", appeared)
            + "]; vanished [" + string.Join(", ", vanished) + "]; reported created [" + string.Join(", ", createdFolders ?? Array.Empty<string>())
            + "], appeared [" + string.Join(", ", appearedFolders ?? Array.Empty<string>()) + "]";
    }

    private static (IReadOnlyList<string> Appeared, IReadOnlyList<string> Vanished) CompareTopLevel(
        IReadOnlyList<string> before, IReadOnlyList<string> after)
    {
        List<string> topBefore = TopLevel(before);
        List<string> topAfter = TopLevel(after);
        return (
            topAfter.Where(name => !topBefore.Contains(name, StringComparer.OrdinalIgnoreCase)).ToList(),
            topBefore.Where(name => !topAfter.Contains(name, StringComparer.OrdinalIgnoreCase)).ToList());
    }

    /// <summary>The top-level folders of a <c>list_folders</c> tree: the paths with no '/' in them.</summary>
    private static List<string> TopLevel(IReadOnlyList<string> paths)
    {
        return paths.Where(path => !string.IsNullOrEmpty(path) && !path.Contains('/', StringComparison.Ordinal)).ToList();
    }

    // ------------------------------------------------------------------ Q96 question 3: where the Drafts folder is registered

    /// <summary>
    /// Where a store says which folder is its Drafts: <c>PR_IPM_DRAFTS_ENTRYID</c> read at each place
    /// the object model can reach, and what it named, for the run's record (Q96 question 3, decided
    /// 2026-10-03: (a) the first live run records it, then (b) - widen the non-creating lookup - is
    /// decided from it). A RECORD, not a verdict: the lookup the product gates on is judged by
    /// <see cref="DraftsVisibleAfter"/>; this says where the folder is registered, so that a blind
    /// lookup can be widened to the right place, or found to need a place the object model cannot read.
    /// </summary>
    /// <param name="store">The throwaway data file's display name.</param>
    /// <param name="reading">What <c>LiveOutlookTestMailer.ReadDraftsDesignations</c> read.</param>
    public static string DescribeDraftsDesignation(string store, DraftsDesignationReading reading)
    {
        ArgumentNullException.ThrowIfNull(reading);
        return "designation: where '" + store + "' registers its Drafts folder (PR_IPM_DRAFTS_ENTRYID) - "
            + string.Join("; ", reading.Reads.Select(read => read.Place + " = " + DescribeDesignation(read, reading.DraftsFolderEntryId)))
            + (reading.DraftsFolderEntryId == null
                ? " (the folder the reply sits in could not be read in that store, so no place can be said to name it)"
                : string.Empty);
    }

    /// <summary>What one designation read says, against the EntryID of the folder the reply was filed in.</summary>
    public static string DescribeDesignation(DesignationRead read, string? draftsFolderEntryId)
    {
        ArgumentNullException.ThrowIfNull(read);
        if (read.Status == null)
        {
            return "no such folder in this store";
        }

        if (read.Status.Value == PropertyReadStatus.NotFound)
        {
            return "not set";
        }

        if (read.Status.Value == PropertyReadStatus.Failed)
        {
            return "could not be read";
        }

        DesignatedEntryId designated = SpecialFolders.ReadEntryId(read.Value);
        if (designated.Kind == DesignatedEntryIdKind.None)
        {
            return "set, naming no folder";
        }

        if (designated.Kind == DesignatedEntryIdKind.Unrecognised)
        {
            return "set to a value that is no entry id";
        }

        return draftsFolderEntryId != null && string.Equals(designated.Hex, draftsFolderEntryId, StringComparison.OrdinalIgnoreCase)
            ? "THE CREATED DRAFTS FOLDER"
            : "another folder (" + designated.Hex!.Substring(0, Math.Min(16, designated.Hex.Length)) + "...)";
    }
}

/// <summary>
/// One place a store's Drafts designation was read: the place, and what the read gave -
/// <see cref="Status"/> null when the place does not exist in the store (a data file with no Inbox).
/// </summary>
public sealed record DesignationRead(string Place, PropertyReadStatus? Status, object? Value);

/// <summary>
/// Every designation read of one store, beside the EntryID of the folder the reply was filed in -
/// null when that could not be read in the store.
/// </summary>
public sealed record DraftsDesignationReading(string? DraftsFolderEntryId, IReadOnlyList<DesignationRead> Reads);
