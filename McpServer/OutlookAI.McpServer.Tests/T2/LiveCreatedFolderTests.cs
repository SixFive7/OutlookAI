using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The live proof of the "created" path of Q85 - a draft tool that files a draft in a mailbox with
/// no Drafts folder lets Outlook create it, and REPORTS it - decided by the maintainer as Q96 (iv),
/// 2026-10-03. Test guests only: it writes, so it carries no <c>Writes=Nothing</c>, and the
/// maintainer's read-only workstation neither selects it nor lets it write (Q74).
/// <para>
/// <b>Where.</b> Not the hub: it has a Drafts folder, so on a guest the creating branch never ran
/// and every proof of it was against fakes. The THROWAWAY data file
/// (<see cref="LiveTestSettings.ThrowawayStoreDisplayName"/>) is recreated before every run by
/// <c>Testbed/guest/Reset-ThrowawayStore.ps1</c>, attached with <c>AddStoreEx</c>, so it holds
/// Deleted Items and nothing else (Docs/live-tier-on-the-vm.md section 1.3). The write allowlist
/// grants it draft and delete, and nothing else, ever, on a machine that may write at all.
/// </para>
/// <para>
/// <b>How.</b> A REPLY, not a new draft: <c>new_draft</c> files in an ACCOUNT's delivery store, and
/// no account delivers into a data file, whereas a reply is filed in its SOURCE item's store's
/// Drafts (Q85) - so a tagged post saved in the data file's Deleted Items is all the reply needs.
/// The reply's own first save lands in the DEFAULT store's Drafts, as every new unsent item's does
/// (measured, <c>CorpusPlacement</c>), so it is the product's lookup - and nothing before it - that
/// makes the data file's Drafts folder. Then <c>discard_draft</c>, which moves the reply to the data
/// file's Deleted Items.
/// </para>
/// <para>
/// <b>What it concludes</b> is decided in <see cref="ThrowawayStoreProof"/>, where CI pins every
/// branch; this class only gathers the answers. And it ends by proving it left nothing: the tagged
/// sweep of the data file and the hub, and every item it made looked up by EntryID - which no
/// designation question can blind. The data file itself is recreated or removed by the script
/// around the run, never by this test.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LiveCreatedFolderTests
{
    /// <summary>The folders this proof's artifacts can be in, in the throwaway data file: Drafts, then Deleted Items last.</summary>
    private static readonly int[] ThrowawayFolders = { SpecialFolders.OlFolderDrafts, SpecialFolders.OlFolderDeletedItems };

    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveCreatedFolderTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Marker => _fixture.RunMarker;

    [Fact]
    [Trait("Requires", "MultipleStores")]
    public void AReplyIntoADataFileWithNoDraftsFolder_ReportsTheDraftsFolderItCreates_AndLeavesNothingBehind()
    {
        string? declared = _fixture.Settings.ThrowawayStoreDisplayName;
        IReadOnlyList<string> throwaway = LivePopulationCoverage.Require(
            _fixture.Settings,
            string.IsNullOrWhiteSpace(declared) ? Array.Empty<string>() : new[] { declared! },
            ThrowawayStoreProof.Population,
            ThrowawayStoreProof.WhatWouldNotRun,
            ThrowawayStoreProof.Remedy,
            _output.WriteLine);
        foreach (string store in throwaway)
        {
            ProveTheCreatedDraftsFolderIsReported(store);
        }
    }

    private void ProveTheCreatedDraftsFolderIsReported(string store)
    {
        // Before anything is written: mounted, and no Drafts folder - by the product's own lookup,
        // which never creates one (Q84).
        ComDefaultFolderInfo? draftsBefore = _fixture.VerifySession.TryGetDefaultFolderInfo(
            store, SpecialFolders.OlFolderDrafts, out string? draftsBeforeError);
        _output.WriteLine($"before: '{store}' Drafts={(draftsBefore == null ? draftsBeforeError : "'" + draftsBefore.Name + "'")}");
        Require(ThrowawayStoreProof.DraftsAbsentBefore(store, draftsBefore == null ? draftsBeforeError ?? "unknown" : null));

        ComDefaultFolderInfo? deletedBefore = _fixture.VerifySession.TryGetDefaultFolderInfo(
            store, SpecialFolders.OlFolderDeletedItems, out string? deletedBeforeError);
        _output.WriteLine($"before: '{store}' Deleted Items={(deletedBefore == null ? deletedBeforeError : "'" + deletedBefore.Name + "'")}");

        string? sourceEntryId = null;
        DraftOutcome? reply = null;
        DiscardDraftOutcome? discarded = null;
        try
        {
            // The reply's source: one tagged post in the data file's Deleted Items.
            sourceEntryId = LiveOutlookTestMailer.SaveTaggedPostInDeletedItems(
                store, Marker, _fixture.TaggedSubject("q96-created-drafts-source"), "Source of the Q96 created-folder proof.");

            // THE CALL UNDER TEST: the product finds no Drafts folder in the source store, makes it,
            // files the reply in it, and must say so.
            LiveStoreWriteGuard.Writable(store, StoreWriteKind.Draft, "reply_draft");
            reply = Service.ReplyDraft(sourceEntryId, "Q96 created-folder proof " + Marker, display: false);
            _output.WriteLine(
                $"reply_draft: store='{reply.Store}' folder='{reply.Folder}' createdFolders=["
                + string.Join(", ", reply.CreatedFolders ?? Array.Empty<string>()) + "]");
            Require(ThrowawayStoreProof.ReplyReportedTheCreatedDrafts(store, reply.CreatedFolders, reply.Store, reply.Folder));

            // Where Outlook designated the folder it made - the question Q85 left open - read the way
            // the discard gate will read it.
            ComDefaultFolderInfo? draftsAfter = _fixture.VerifySession.TryGetDefaultFolderInfo(
                store, SpecialFolders.OlFolderDrafts, out string? draftsAfterError);
            _output.WriteLine($"after: the non-creating lookup sees Drafts={(draftsAfter == null ? draftsAfterError : "'" + draftsAfter.Name + "'")}");
            Require(ThrowawayStoreProof.DraftsVisibleAfter(
                store, draftsAfter == null ? draftsAfterError ?? "unknown" : null, draftsAfter?.Name, reply.Folder));

            LiveStoreWriteGuard.Writable(store, StoreWriteKind.Delete, "discard_draft");
            discarded = Service.DiscardDraft(reply.EntryId);
            _output.WriteLine(
                $"discard_draft: discarded={discarded.Discarded} to='{discarded.ToFolder}' createdFolders=["
                + string.Join(", ", discarded.CreatedFolders ?? Array.Empty<string>()) + "]");
            Require(ThrowawayStoreProof.DiscardReported(store, deletedBefore != null, discarded.Discarded, discarded.CreatedFolders));
        }
        finally
        {
            CleanUp(store, reply, discarded);
        }

        AssertNothingLeft(store, sourceEntryId, reply, discarded);
    }

    /// <summary>
    /// Removes what this proof made, through the tested helpers only (mailbox-safety rule 1): the
    /// reply by its EntryID when it was not discarded, then the tagged purge of the data file - its
    /// Drafts and, last, its Deleted Items, where the source post and the discarded reply are - and
    /// of the hub, whose Drafts held the reply's first save until it was moved.
    /// </summary>
    private void CleanUp(string store, DraftOutcome? reply, DiscardDraftOutcome? discarded)
    {
        if (reply != null && discarded == null)
        {
            Quietly(() => LiveOutlookTestMailer.DeleteItemByEntryId(store, reply.EntryId, Marker));
        }

        Quietly(() => LiveOutlookTestMailer.DeleteTaggedArtifacts(store, Marker, ThrowawayFolders));
        Quietly(() => LiveOutlookTestMailer.DeleteTaggedArtifacts(_fixture.Settings.TestHubStoreDisplayName, Marker));
    }

    /// <summary>
    /// Zero artifacts, proven twice: by the tagged count of the data file and the hub, and by
    /// EntryID for every item this proof made - a count reads folders through the lookup whose
    /// designation this proof is about, an EntryID does not.
    /// </summary>
    private void AssertNothingLeft(string store, string? sourceEntryId, DraftOutcome? reply, DiscardDraftOutcome? discarded)
    {
        foreach ((string where, int[]? folders) in new[] { (store, (int[]?)ThrowawayFolders), (_fixture.Settings.TestHubStoreDisplayName, (int[]?)null) })
        {
            int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(where, Marker, folders, out int stragglers);
            if (stragglers > 0)
            {
                _output.WriteLine($"cleanup[{where}]: {stragglers} late-materialized artifact(s) purged");
            }

            Assert.Equal(0, remaining);
        }

        string storeId = _fixture.GetStoreId(store);
        foreach (string? entryId in new[] { sourceEntryId, reply?.EntryId, discarded?.NewEntryId })
        {
            if (entryId != null)
            {
                ComDraftInfo? still = _fixture.VerifySession.TryGetMailInfo(entryId, storeId, out _);
                Assert.True(still == null, "an item this proof made can still be opened in '" + store + "' after the cleanup.");
            }
        }
    }

    private void Require(string? failure)
    {
        if (failure != null)
        {
            _output.WriteLine("FAILED: " + failure);
            Assert.Fail(failure);
        }
    }

    private void Quietly(Action cleanup)
    {
        try
        {
            cleanup();
        }
        catch (Exception ex)
        {
            // The assertions after the cleanup are what decide; a cleanup step that failed is
            // reported, and the count and the EntryID checks then say what it left.
            _output.WriteLine("cleanup step failed: " + ex.GetType().Name);
        }
    }
}
