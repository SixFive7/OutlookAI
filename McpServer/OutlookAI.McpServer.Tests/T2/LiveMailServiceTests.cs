using System.Diagnostics;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Phase-2 T2 live acceptance (v3.MD section 0.6): search-to-read round-trips on real
/// hits across all three stores, truncation flags on a >100 KB mail, attachment save
/// from an index document hit, exact list_accounts (3 accounts + delegates distinct +
/// online-only flags), list_folders, thread on both paths, and index_status. Logging is
/// content-free for business stores (S4): counts, ids, timings, booleans only.
/// </summary>
[Collection("LivePhase2")]
[Trait("Category", "Live")]
public sealed class LiveMailServiceTests
{
    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveMailServiceTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    [Fact]
    public void RoundTrip_SearchThenRead_TenHitsAcrossStores()
    {
        List<HitSummary> hits = new();
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            SearchOutcome outcome = Service.Search(new SearchRequest
            {
                Mode = SearchMode.Fast,
                Store = store,
                IncludeAttachmentHits = false,
                Top = 8,
            });
            hits.AddRange(outcome.Hits.Where(h => !string.IsNullOrEmpty(h.Subject)).Take(4));
            _output.WriteLine($"store hits sampled: {outcome.Hits.Count} (indexMs={outcome.IndexElapsedMs})");
        }

        hits = hits.Take(10).ToList();
        if (hits.Count < 10)
        {
            SearchOutcome extra = Service.Search(new SearchRequest
            {
                Mode = SearchMode.Fast,
                Store = _fixture.Settings.ExpectedStoreDisplayNames[0],
                IncludeAttachmentHits = false,
                Top = 30,
            });
            hits.AddRange(extra.Hits.Where(h => !string.IsNullOrEmpty(h.Subject) && hits.All(x => x.Id != h.Id))
                .Take(10 - hits.Count));
        }

        Assert.Equal(10, hits.Count);

        var readMillis = new List<long>();
        int succeeded = 0;
        foreach (HitSummary hit in hits)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            ReadOutcome read = Service.Read(hit.Id, maxBodyChars: 2000);
            stopwatch.Stop();
            readMillis.Add(stopwatch.ElapsedMilliseconds);

            Assert.True(read.EntryId.Length >= 48, "read must return a real EntryID");
            Assert.Equal(hit.Subject, read.Subject);
            if (hit.ReceivedUtc.HasValue && read.ReceivedUtc.HasValue)
            {
                Assert.True(Math.Abs((hit.ReceivedUtc.Value - read.ReceivedUtc.Value).TotalSeconds) <= 5,
                    "read ReceivedUtc must match the index row within 5 s");
            }

            Assert.True(read.BodyTotalChars >= read.Body.Length);
            succeeded++;
            _output.WriteLine($"read ok id={hit.Id} locate={read.LocatedVia} locateMs={read.LocateMs ?? 0} totalMs={stopwatch.ElapsedMilliseconds} bodyChars={read.BodyTotalChars}");
        }

        Assert.Equal(10, succeeded);
        _output.WriteLine($"read ms: avg={readMillis.Average():F0} max={readMillis.Max()}");

        // Located EntryIDs are cached: a repeat read must skip the locate cost entirely.
        Stopwatch cached = Stopwatch.StartNew();
        ReadOutcome again = Service.Read(hits[0].Id, maxBodyChars: 500);
        cached.Stop();
        Assert.Equal("cached", again.LocatedVia);
        _output.WriteLine($"cached re-read ms: {cached.ElapsedMilliseconds}");
        Assert.True(cached.ElapsedMilliseconds < 2000, "cached re-read should be fast");
    }

    [Fact]
    public void Truncation_MailOver100KB_FlagsAndTotalsCorrect()
    {
        const int cap = 20000;
        ReadOutcome? bigRead = null;
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            List<HitSummary> candidates;
            try
            {
                candidates = Service.Search(new SearchRequest
                {
                    Mode = SearchMode.Fast,
                    Store = store,
                    IncludeAttachmentHits = false,
                    OrderBySizeDescending = true,
                    HasAttachments = false,
                    Top = 15,
                }).Hits.ToList();
            }
            catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
            {
                // ORDER BY System.Size unsupported would surface here - fall back to a
                // date-ordered pull filtered by size.
                _output.WriteLine($"size ordering failed ({ex.GetType().Name}); falling back to client-side size filter");
                candidates = Service.Search(new SearchRequest
                {
                    Mode = SearchMode.Fast,
                    Store = store,
                    IncludeAttachmentHits = false,
                    HasAttachments = false,
                    Top = 100,
                }).Hits.Where(h => (h.SizeBytes ?? 0) > 100_000).ToList();
            }

            foreach (HitSummary hit in candidates.Where(h => (h.SizeBytes ?? 0) > 100_000).Take(6))
            {
                ReadOutcome read = Service.Read(hit.Id, maxBodyChars: cap);
                _output.WriteLine($"candidate size={hit.SizeBytes} bodyTotal={read.BodyTotalChars} origin={read.BodyOrigin}");
                if (read.BodyTotalChars > 100_000)
                {
                    bigRead = read;
                    break;
                }
            }

            if (bigRead != null)
            {
                break;
            }
        }

        Assert.NotNull(bigRead);
        Assert.True(bigRead!.BodyTruncated, "a >100 KB body must be flagged truncated at a 20 KB cap");
        Assert.Equal(cap, bigRead.Body.Length);
        Assert.True(bigRead.BodyTotalChars > 100_000);
        _output.WriteLine($"truncation verified: total={bigRead.BodyTotalChars} returned={bigRead.Body.Length} truncated={bigRead.BodyTruncated}");
    }

    [Fact]
    public void AttachmentHit_ReadParent_SaveToScratch()
    {
        HitSummary? attachmentHit = null;
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            SearchOutcome outcome = Service.Search(new SearchRequest
            {
                Mode = SearchMode.Fast,
                Query = _fixture.Settings.ProbeTerm,
                Store = store,
                AttachmentHitsOnly = true,
                Top = 10,
            });
            attachmentHit = outcome.Hits.FirstOrDefault(h => h.IsAttachmentHit && h.AttachmentFileName != null);
            if (attachmentHit != null)
            {
                break;
            }
        }

        Assert.NotNull(attachmentHit);
        _output.WriteLine($"attachment hit: id={attachmentHit!.Id} hasFileName={attachmentHit.AttachmentFileName != null}");

        // read resolves the PARENT mail of an attachment hit.
        ReadOutcome parent = Service.Read(attachmentHit.Id, maxBodyChars: 200);
        Assert.True(parent.Attachments.Count >= 1, "parent mail must list attachments");

        AttachmentView? match = parent.Attachments.FirstOrDefault(a =>
            string.Equals(a.FileName, attachmentHit.AttachmentFileName, StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(match);

        SaveAttachmentOutcome saved = Service.SaveAttachment(attachmentHit.Id, match!.Index);
        try
        {
            Assert.True(File.Exists(saved.SavedPath), "saved attachment file must exist");
            Assert.True(saved.SizeBytes > 0);
            _output.WriteLine($"saved: bytes={saved.SizeBytes} index={saved.AttachmentIndex} dirDefault={saved.SavedPath.StartsWith(MailService.DefaultAttachmentDirectory, StringComparison.OrdinalIgnoreCase)}");

            string extension = Path.GetExtension(saved.SavedPath).ToLowerInvariant();
            if (extension is ".txt" or ".csv" or ".htm" or ".html" or ".xml" or ".json" or ".eml")
            {
                string content = File.ReadAllText(saved.SavedPath);
                bool grep = content.Contains(_fixture.Settings.ProbeTerm, StringComparison.OrdinalIgnoreCase);
                _output.WriteLine($"text attachment grep for probe term: {grep}");
            }
            else
            {
                _output.WriteLine("binary attachment: content grep covered by the fresh-mode txt round-trip test");
            }
        }
        finally
        {
            TryDelete(saved.SavedPath);
        }
    }

    [Fact]
    public void ListAccounts_ExactAccountsDelegatesAndFlags()
    {
        AccountsOutcome outcome = Service.ListAccounts();

        // Exactly the three configured accounts.
        Assert.Equal(3, outcome.Accounts.Count);
        HashSet<string> expectedAccounts = new(_fixture.Settings.ExpectedStoreDisplayNames, StringComparer.OrdinalIgnoreCase);
        foreach (AccountView account in outcome.Accounts)
        {
            Assert.False(string.IsNullOrWhiteSpace(account.SmtpAddress));
            Assert.Contains(account.SmtpAddress!, expectedAccounts, StringComparer.OrdinalIgnoreCase);
        }

        // The three primary stores present and not delegate-flagged.
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreView view = Assert.Single(outcome.Stores, s =>
                string.Equals(s.DisplayName, store, StringComparison.OrdinalIgnoreCase));
            Assert.False(view.IsDelegate, "primary store must not be delegate-flagged");
            Assert.True(view.LocallySearchable);
            Assert.True(view.InLocalIndex == true, "primary store must be visible in the local index");
        }

        // Delegate stores listed DISTINCTLY and flagged.
        Assert.NotEmpty(_fixture.Settings.ExpectedDelegateStoreDisplayNames);
        foreach (string delegateStore in _fixture.Settings.ExpectedDelegateStoreDisplayNames)
        {
            StoreView view = Assert.Single(outcome.Stores, s =>
                string.Equals(s.DisplayName, delegateStore, StringComparison.OrdinalIgnoreCase));
            Assert.True(view.IsDelegate, "delegate cache store must be delegate-flagged");
            Assert.False(view.OnlineOnly);
        }

        // No unexpected stores; on this machine no online-only stores exist, and every
        // store must be flagged locally searchable (D22/D25 flag mechanism verified).
        int expectedTotal = _fixture.Settings.ExpectedStoreDisplayNames.Count
            + _fixture.Settings.ExpectedDelegateStoreDisplayNames.Count;
        Assert.Equal(expectedTotal, outcome.Stores.Count);
        Assert.All(outcome.Stores, s => Assert.False(s.OnlineOnly, "no online-only store expected on this machine"));
        Assert.All(outcome.Stores, s => Assert.True(s.LocallySearchable));

        _output.WriteLine($"accounts={outcome.Accounts.Count} stores={outcome.Stores.Count} "
            + $"delegates={outcome.Stores.Count(s => s.IsDelegate)} onlineOnly={outcome.Stores.Count(s => s.OnlineOnly)}");
    }

    [Fact]
    public void ListFolders_TestHub_ReturnsTreeWithCounts()
    {
        FoldersOutcome outcome = Service.ListFolders(_fixture.Settings.TestHubStoreDisplayName, depth: 3);

        StoreFoldersView store = Assert.Single(outcome.Stores);
        Assert.True(store.Folders.Count >= 3, "test hub should expose at least a few folders");
        Assert.All(store.Folders, f => Assert.False(string.IsNullOrWhiteSpace(f.Path)));
        Assert.Contains(store.Folders, f => f.Items.HasValue);
        _output.WriteLine($"folders={store.Folders.Count} withCounts={store.Folders.Count(f => f.Items.HasValue)} truncated={outcome.Truncated}");
    }

    [Fact]
    public void Thread_IndexPath_AndComFallback()
    {
        // A recent hit with a conversation id from a busy store.
        HitSummary? seed = null;
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            seed = Service.Search(new SearchRequest
            {
                Mode = SearchMode.Fast,
                Store = store,
                IncludeAttachmentHits = false,
                Top = 20,
            }).Hits.FirstOrDefault(h => h.ConversationId != null && !string.IsNullOrEmpty(h.Subject));
            if (seed != null)
            {
                break;
            }
        }

        Assert.NotNull(seed);

        // Index path.
        ThreadOutcome indexThread = Service.Thread(seed!.ConversationId, seed.Id, seed.Store);
        _output.WriteLine($"thread(index) source={indexThread.Source} members={indexThread.Hits.Count} ms={indexThread.ElapsedMs}");
        Assert.True(indexThread.Hits.Count >= 1);
        if (indexThread.Source == "index")
        {
            Assert.All(indexThread.Hits, h => Assert.Equal(seed.ConversationId, h.ConversationId));
        }

        // COM fallback path: a conversation id that cannot exist in the index forces the
        // Outlook Conversation walk over the referenced item.
        ThreadOutcome comThread = Service.Thread("zzzznonexistentconversationzzzz", seed.Id, seed.Store);
        _output.WriteLine($"thread(com) source={comThread.Source} members={comThread.Hits.Count} ms={comThread.ElapsedMs}");
        Assert.Equal("com", comThread.Source);
        Assert.True(comThread.Hits.Count >= 1, "COM conversation walk must return at least the item itself");
        Assert.All(comThread.Hits, h => Assert.False(string.IsNullOrEmpty(h.Id)));
    }

    [Fact]
    public void IndexStatus_Live_ReportsProviderStalenessAndStores()
    {
        IndexStatusOutcome status = Service.IndexStatus();

        Assert.True(status.Provider is "OleDb" or "AdodbCom", $"unexpected provider: {status.Provider}");
        Assert.NotNull(status.NewestIndexedUtc);
        Assert.True(status.IndexAgeMinutes >= 0);
        Assert.NotNull(status.PerStore);
        Assert.True(status.PerStore!.Count >= 3, "per-store staleness must cover at least the 3 account stores");
        Assert.NotEmpty(status.Advice);
        _output.WriteLine($"provider={status.Provider} outlookRunning={status.OutlookRunning} "
            + $"ageMin={status.IndexAgeMinutes:F1} perStore={status.PerStore.Count}");
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
