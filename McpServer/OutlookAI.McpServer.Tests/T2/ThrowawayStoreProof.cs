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
}
