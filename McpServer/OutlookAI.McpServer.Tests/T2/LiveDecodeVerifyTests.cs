using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 hit-mapping verification (v3.MD section 0.6 Phase 1): 25 hits sampled across the
/// three stores; every hit must become an openable COM item whose Subject and
/// ReceivedTime match the index row, with at least 24 mapped by the module's primary
/// path and the remainder by the ItemPathDisplay fallback.
///
/// PHASE-1 REALITY (recorded in v3.MD section 0.8): on this machine all stores are
/// cached Exchange, whose object model exposes 70-byte Exchange EntryIDs -
/// GetItemFromID rejects the 24-byte OST-internal id decoded from the index URL
/// (0x80040107) in every store. The decode still yields the correct 16-byte store UID
/// (asserted below against the real EntryID bytes 4..19), and mapping goes through
/// <see cref="HitLocator"/>: URL folder segments -> narrow folder probe (primary),
/// ItemPathDisplay derivation (fallback). Every located item is then re-opened by its
/// REAL EntryID via GetItemFromID as the verify-on-open step.
///
/// May start Outlook (S7/D17); never stops it. Logging is content-free (S4): counts,
/// hex ids, hresults, timings - no subjects/bodies.
/// </summary>
[Collection("LivePhase1")]
[Trait("Category", "Live")]
public sealed class LiveDecodeVerifyTests
{
    private const int SampleTarget = 25;
    private const int TimeToleranceSeconds = 5;

    private readonly LivePhase1Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDecodeVerifyTests(LivePhase1Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void HitMapping_25SampledHits_AllOpenAndMatch_AtLeast24ViaPrimaryPath()
    {
        List<IndexHit> samples = SampleHitsAcrossStores();
        Assert.Equal(SampleTarget, samples.Count);

        int primary = 0;
        int fallback = 0;
        int verifiedOpens = 0;
        int uidConfirmed = 0;
        int uidMismatchNonDelegate = 0;
        int utcInterpretation = 0;
        int rawLocalInterpretation = 0;
        var perStoreCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var locateMillis = new List<long>();

        foreach (IndexHit hit in samples)
        {
            perStoreCounts[hit.StoreDisplayName!] =
                perStoreCounts.TryGetValue(hit.StoreDisplayName!, out int n) ? n + 1 : 1;

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            HitLocationResult location = HitLocator.Locate(_fixture.Session, hit, TimeToleranceSeconds);
            stopwatch.Stop();
            locateMillis.Add(stopwatch.ElapsedMilliseconds);

            if (location.Tier == HitLocationTier.Failed || location.Located == null)
            {
                _output.WriteLine($"locate FAILED shortId={hit.EntryIdHex} error={location.Error}");
                continue;
            }

            if (location.Tier == HitLocationTier.UrlSegments)
            {
                primary++;
            }
            else
            {
                fallback++;
            }

            // v3.MD section-4 store-UID claim: the real EntryID carries the decoded
            // 16-byte store UID at bytes 4..19 (hex chars 8..39).
            string realId = location.Located.EntryId;
            if (hit.StoreUidHex != null && realId.Length >= 40
                && string.Equals(realId.Substring(8, 32), hit.StoreUidHex, StringComparison.OrdinalIgnoreCase))
            {
                uidConfirmed++;
            }
            else
            {
                // Phase-1 finding: items indexed under the OWNER account's /1/ delegate
                // subtree carry the OWNER store's UID in the URL short id, while the item
                // physically lives in the delegate's own store (different UID in the real
                // EntryID). Routing must therefore use the folder-segment delegate rule,
                // never the UID, for store-type-1 hits.
                if (hit.StoreType != 1)
                {
                    uidMismatchNonDelegate++;
                }

                _output.WriteLine($"UID mismatch storeType={hit.StoreType} store={location.StoreDisplayName} "
                    + $"shortUid={hit.StoreUidHex} realUid={(realId.Length >= 40 ? realId.Substring(8, 32) : realId)}");
            }

            // Verify-on-open: re-open by the REAL EntryID and compare against the index row.
            string storeId = _fixture.GetComStoreId(location.StoreDisplayName!);
            ComOpenResult? opened = _fixture.Session.TryOpenItem(realId, storeId, out string? openError);
            if (opened == null)
            {
                _output.WriteLine($"re-open FAILED ({openError}) for located id len={realId.Length}");
                continue;
            }

            bool subjectMatches = string.Equals(opened.Subject ?? string.Empty, hit.Subject ?? string.Empty, StringComparison.Ordinal);
            bool timeMatches = ClassifyTimeMatch(opened.ReceivedTime, hit.DateReceivedUtc, ref utcInterpretation, ref rawLocalInterpretation);
            if (subjectMatches && timeMatches)
            {
                verifiedOpens++;
            }
            else
            {
                _output.WriteLine($"verify mismatch shortId={hit.EntryIdHex}: subject={subjectMatches} time={timeMatches}");
            }
        }

        _output.WriteLine("sampled per store: " + string.Join(", ", perStoreCounts.Select(kv => $"{kv.Key}={kv.Value}")));
        _output.WriteLine($"primary(UrlSegments)={primary} fallback(ItemPathDisplay)={fallback} "
            + $"verifiedOpens={verifiedOpens}/{SampleTarget} uidConfirmed={uidConfirmed}");
        _output.WriteLine($"receivedTime interpretation: utc={utcInterpretation} rawLocal={rawLocalInterpretation}");
        _output.WriteLine($"locate ms: avg={locateMillis.Average():F0} max={locateMillis.Max()}");

        Assert.True(primary >= SampleTarget - 1,
            $"only {primary}/{SampleTarget} hits mapped via the primary URL-segment path");
        Assert.Equal(SampleTarget, primary + fallback);
        Assert.Equal(SampleTarget, verifiedOpens);
        // Store-UID correspondence must hold for every NON-delegate hit; delegate-subtree
        // hits legitimately carry the owner store's UID (logged + recorded in v3.MD 0.8).
        Assert.Equal(0, uidMismatchNonDelegate);

        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            Assert.True(perStoreCounts.ContainsKey(store), $"no samples came from store {store}");
        }
    }

    [Fact]
    public void ShortDecodedId_IsRejectedByGetItemFromID_DiscoveryRecorded()
    {
        // Pins the Phase-1 platform finding so a future behavior change is noticed:
        // the 24-byte decoded id is NOT openable on cached Exchange stores.
        IndexHit? hit = _fixture.Service.Search(new IndexQuery
        {
            Scope = _fixture.GetScope(_fixture.Settings.ExpectedStoreDisplayNames[0]).StorePrefix,
            Kinds = KindFilter.EmailOnly,
            Top = 1,
        }).Hits.FirstOrDefault();
        Assert.NotNull(hit);
        Assert.NotNull(hit!.EntryIdHex);

        string storeId = _fixture.GetComStoreId(hit.StoreDisplayName!);
        ComOpenResult? opened = _fixture.Session.TryOpenItem(hit.EntryIdHex!, storeId, out string? error);

        _output.WriteLine($"short-id open: result={(opened == null ? "rejected" : "OPENED")} error={error}");
        Assert.Null(opened);
        Assert.Contains("80040107", error ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AttachmentHit_ParentMapping_OpensParentWithMatchingAttachment()
    {
        // Real attachment-content entries (kind=document under a mapi store scope).
        List<IndexHit> attachmentHits = new();
        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            IndexSearchResult result = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.DocumentsOnly,
                Top = 25,
            });
            attachmentHits.AddRange(result.Hits.Where(h =>
                h.IsAttachmentHit && h.EntryIdHex != null && !string.IsNullOrEmpty(h.AttachmentFileName)));
            if (attachmentHits.Count >= 15)
            {
                break;
            }
        }

        _output.WriteLine($"attachment-entry candidates: {attachmentHits.Count}");
        Assert.True(attachmentHits.Count > 0, "no indexed attachment entries found under any store scope");

        int verified = 0;
        foreach (IndexHit hit in attachmentHits)
        {
            Assert.NotNull(hit.ParentItemUrl); // parent mapping = strip the /at= segment

            HitLocationResult location = HitLocator.Locate(_fixture.Session, hit, toleranceSeconds: 120);
            if (location.Tier == HitLocationTier.Failed || location.Located == null)
            {
                _output.WriteLine($"parent locate failed ({location.Error})");
                continue;
            }

            string storeId = _fixture.GetComStoreId(location.StoreDisplayName!);
            IReadOnlyList<string>? names = _fixture.Session.TryGetAttachmentFileNames(
                location.Located.EntryId, storeId, out string? error);
            if (names == null)
            {
                _output.WriteLine($"parent open failed ({error})");
                continue;
            }

            bool fileNameMatches = names.Any(n =>
                string.Equals(n, hit.AttachmentFileName, StringComparison.OrdinalIgnoreCase));
            _output.WriteLine($"parent located tier={location.Tier} attachments={names.Count} fileNameMatch={fileNameMatches}");
            if (fileNameMatches)
            {
                verified++;
                break;
            }
        }

        Assert.True(verified > 0, "no attachment hit could be mapped to a parent message carrying that file name");
    }

    private static bool ClassifyTimeMatch(
        DateTime? comReceivedLocal,
        DateTime? indexUtc,
        ref int utcInterpretation,
        ref int rawLocalInterpretation)
    {
        if (!indexUtc.HasValue && !comReceivedLocal.HasValue)
        {
            return true;
        }

        if (!indexUtc.HasValue || !comReceivedLocal.HasValue)
        {
            return false;
        }

        DateTime local = comReceivedLocal.Value;
        double utcDelta = Math.Abs((DateTime.SpecifyKind(local, DateTimeKind.Local).ToUniversalTime() - indexUtc.Value).TotalSeconds);
        double rawDelta = Math.Abs((DateTime.SpecifyKind(local, DateTimeKind.Utc) - indexUtc.Value).TotalSeconds);
        if (utcDelta <= TimeToleranceSeconds)
        {
            utcInterpretation++;
            return true;
        }

        if (rawDelta <= TimeToleranceSeconds)
        {
            rawLocalInterpretation++;
            return true;
        }

        return false;
    }

    private List<IndexHit> SampleHitsAcrossStores()
    {
        // Recent mail per store (non-empty subjects keep the folder probe narrow); the
        // small test-hub store contributes what it has and the remainder is topped up
        // from the other stores.
        var samples = new List<IndexHit>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        const int perStore = 9;

        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            IndexSearchResult result = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.EmailOnly,
                Top = perStore * 2,
            });
            foreach (IndexHit hit in result.Hits)
            {
                if (samples.Count(s => string.Equals(s.StoreDisplayName, storeName, StringComparison.OrdinalIgnoreCase)) >= perStore)
                {
                    break;
                }

                if (hit.EntryIdHex != null && !string.IsNullOrEmpty(hit.Subject) && seen.Add(hit.EntryIdHex))
                {
                    samples.Add(hit);
                }
            }
        }

        if (samples.Count < SampleTarget)
        {
            StoreScopeInfo biggest = _fixture.StoreScopes.OrderByDescending(s => s.SampleCount).First();
            IndexSearchResult extra = _fixture.Service.Search(new IndexQuery
            {
                Scope = biggest.StorePrefix,
                Kinds = KindFilter.EmailOnly,
                Top = SampleTarget * 4,
            });
            foreach (IndexHit hit in extra.Hits)
            {
                if (samples.Count >= SampleTarget)
                {
                    break;
                }

                if (hit.EntryIdHex != null && !string.IsNullOrEmpty(hit.Subject) && seen.Add(hit.EntryIdHex))
                {
                    samples.Add(hit);
                }
            }
        }

        return samples.Take(SampleTarget).ToList();
    }
}
