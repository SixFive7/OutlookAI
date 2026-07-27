using System.Diagnostics;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Live proof of the attachment-recall fix (soak fix 16 / v3.MD block (q)): the shipped
/// <c>(System.Kind='email' OR System.Kind='document')</c> predicate kept only the
/// <c>document</c> slice of attachment-content rows, so a term living solely inside an
/// image, an embedded message or an <c>.ics</c> invite could never find its parent mail -
/// 709 of 3,139 attachment rows (22.6%) measured across this profile.
/// <para>
/// READ-ONLY against the real corpus (counts, kinds and timings only - never a subject,
/// sender or body, S4); the seeded end-to-end proof writes to the designated test mailbox
/// only and cleans up through the tested helpers.
/// </para>
/// </summary>
[Collection("LivePhase1")]
[Trait("Category", "Live")]
public sealed class LiveAttachmentKindRecallTests
{
    private readonly LivePhase1Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveAttachmentKindRecallTests(LivePhase1Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private static IIndexClient Client => IndexClientFactory.CreateAuto(out _);

    [Fact]
    public void DroppedAttachmentKinds_AreRecoveredByTheNewShape_AndTheGrowthIsMeasured()
    {
        IIndexClient client = Client;
        int totalOld = 0;
        int totalNew = 0;
        int recoveredAttachmentRows = 0;
        Dictionary<string, int> recoveredKinds = new(StringComparer.OrdinalIgnoreCase);

        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            string scope = _fixture.GetScope(store).StorePrefix;

            // Old shipped shape versus the shipped one. Same scope, same TOP, same order:
            // the ONLY difference is the kind predicate.
            IReadOnlyList<IReadOnlyDictionary<string, object?>> oldRows = client.ExecuteRows(
                "SELECT TOP 30000 System.ItemUrl, System.Kind FROM SystemIndex WHERE SCOPE='" + scope
                + "' AND (System.Kind='email' OR System.Kind='document')", 30000);
            IReadOnlyList<IReadOnlyDictionary<string, object?>> newRows = client.ExecuteRows(
                "SELECT TOP 30000 System.ItemUrl, System.Kind FROM SystemIndex WHERE SCOPE='" + scope + "'",
                30000);

            HashSet<string> oldUrls = new(oldRows.Select(Url).Where(u => u != null)!, StringComparer.OrdinalIgnoreCase);
            int storeRecovered = 0;
            foreach (IReadOnlyDictionary<string, object?> row in newRows)
            {
                string? url = Url(row);
                if (url == null || oldUrls.Contains(url) || !IndexRowFilter.IsAttachmentRow(url))
                {
                    continue;
                }

                storeRecovered++;
                foreach (string kind in Kinds(row))
                {
                    recoveredKinds[kind] = recoveredKinds.TryGetValue(kind, out int n) ? n + 1 : 1;
                }
            }

            totalOld += oldRows.Count;
            totalNew += newRows.Count;
            recoveredAttachmentRows += storeRecovered;
            bool capped = newRows.Count >= 30000 || oldRows.Count >= 30000;
            _output.WriteLine(
                $"store#{_fixture.Settings.ExpectedStoreDisplayNames.IndexOf(store)}: old {oldRows.Count} rows, "
                + $"new {newRows.Count} rows"
                + (capped ? " (CAPPED at the TOP window)" : $" (+{Percent(newRows.Count - oldRows.Count, oldRows.Count)}%)")
                + $", attachment rows recovered {storeRecovered}.");
        }

        _output.WriteLine(
            $"TOTAL: {totalOld} -> {totalNew} rows over the sampled window, "
            + $"{recoveredAttachmentRows} previously unmatchable attachment rows; kinds recovered: "
            + string.Join(", ", recoveredKinds.OrderByDescending(k => k.Value).Select(k => k.Key + "=" + k.Value)));
        _output.WriteLine(
            "NOTE: stores whose row count equals the TOP window are capped - read their growth from the "
            + "uncapped ones and from the recovered-row counts, not from the totals.");

        // The whole point: attachment rows that the old predicate could never return.
        Assert.True(
            recoveredAttachmentRows > 0,
            "the new shape must recover attachment rows the kind filter dropped");
        Assert.True(totalNew >= totalOld, "dropping the kind filter can only widen the candidate set");

        // ...and they are NOT documents - those were never dropped.
        Assert.DoesNotContain("document", recoveredKinds.Keys, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void PostFilter_KeepsMailOnly_AndAttachmentRowsOfEveryKind()
    {
        // Admission moved from SQL to code; prove the code decides the same thing the old
        // predicate did for MESSAGE rows (mail only - meeting requests stay out) while
        // keeping every attachment row.
        IIndexClient client = Client;
        string scope = _fixture.GetScope(_fixture.Settings.ExpectedStoreDisplayNames[0]).StorePrefix;
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = client.ExecuteRows(
            "SELECT TOP 20000 System.ItemUrl, System.Kind FROM SystemIndex WHERE SCOPE='" + scope + "'", 20000);

        int keptMessages = 0;
        int keptAttachments = 0;
        int droppedMessages = 0;
        foreach (IReadOnlyDictionary<string, object?> row in rows)
        {
            string? url = Url(row);
            if (url == null)
            {
                continue;
            }

            bool attachment = IndexRowFilter.IsAttachmentRow(url);
            bool keep = attachment || IndexRowFilter.HasEmailKind(Kinds(row).ToList());
            if (!keep)
            {
                droppedMessages++;
                // Everything dropped must be a non-mail MESSAGE row (calendar class).
                Assert.False(attachment);
                continue;
            }

            if (attachment)
            {
                keptAttachments++;
            }
            else
            {
                keptMessages++;
            }
        }

        _output.WriteLine(
            $"post-filter: {keptMessages} mail rows + {keptAttachments} attachment rows kept, "
            + $"{droppedMessages} non-mail message rows dropped (of {rows.Count}).");
        Assert.True(keptMessages > 0 && keptAttachments > 0);
    }

    [Fact]
    public void QuerySetLatency_IsUnchangedWithinNoise()
    {
        // The block-(q) claim is "no extra query, ~+6% rows". Measure the agent-sized
        // shape (TOP 26 + ORDER BY) old versus new, warm best-of-3, per store.
        IIndexClient client = Client;
        string term = _fixture.Settings.ProbeTerm;
        long oldTotal = 0;
        long newTotal = 0;

        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            string scope = _fixture.GetScope(store).StorePrefix;
            string contains = "(CONTAINS(System.Subject, '\"" + term + "\"') OR CONTAINS(System.Search.Contents, '\""
                + term + "\"'))";
            string tail = " ORDER BY System.Message.DateReceived DESC";
            string oldSql = "SELECT TOP 26 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + scope
                + "' AND (System.Kind='email' OR System.Kind='document') AND " + contains + tail;

            // The shipped shape over-fetches (TOP 62) because admission moved to code.
            string newSql = "SELECT TOP " + IndexRowFilter.ComputeSqlTop(26, scoped: true, maxTop: WsSqlBuilder.MaxTop)
                + " System.ItemUrl FROM SystemIndex WHERE SCOPE='" + scope + "' AND " + contains + tail;

            long oldMs = BestOfThree(client, oldSql, 26);
            long newMs = BestOfThree(client, newSql, 62);
            oldTotal += oldMs;
            newTotal += newMs;
            _output.WriteLine(
                $"store#{_fixture.Settings.ExpectedStoreDisplayNames.IndexOf(store)} latency: old {oldMs} ms -> new {newMs} ms.");
        }

        _output.WriteLine($"query-set latency: old {oldTotal} ms -> new {newTotal} ms across the store set.");

        // Generous bound: this is a regression tripwire, not a benchmark. A structural
        // slowdown (a property scan sneaking in) would be multiples, not percent.
        Assert.True(
            newTotal <= (oldTotal * 3) + 750,
            $"new shape must not be structurally slower (old {oldTotal} ms, new {newTotal} ms)");
    }

    [Fact]
    public void PreviouslyDroppedAttachmentKinds_SurfaceThroughSearch_AndReadOpensTheParent()
    {
        // Deterministic end-to-end proof through the PRODUCT, read-only on the real
        // corpus: attachment hits whose file type the old kind filter dropped now come
        // back from search, and reading one opens the PARENT mail.
        using MailService service = MailService.CreateDefault();
        string[] previouslyDropped =
        {
            ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".heic", ".webp",
            ".msg", ".eml", ".ics", ".vcs",
            ".mp3", ".wav", ".m4a", ".mp4", ".mov", ".avi", ".wmv",
        };

        HitSummary? probe = null;
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            SearchOutcome outcome = service.Search(new SearchRequest
            {
                Store = store,
                AttachmentHitsOnly = true,
                Top = 50,
                SnippetChars = 0,
            });

            int droppedKindHits = outcome.Hits.Count(h => HasExtension(h.AttachmentFileName, previouslyDropped));
            _output.WriteLine(
                $"store#{_fixture.Settings.ExpectedStoreDisplayNames.IndexOf(store)}: {outcome.Hits.Count} attachment "
                + $"hit(s), {droppedKindHits} of a type the old kind filter dropped.");

            probe ??= outcome.Hits.FirstOrDefault(h => HasExtension(h.AttachmentFileName, previouslyDropped));
            Assert.All(outcome.Hits, h => Assert.True(h.IsAttachmentHit));
        }

        Assert.NotNull(probe);

        // Attachment -> parent mapping through the product: read resolves the PARENT mail.
        ReadOutcome read = service.Read(probe!.Id, maxBodyChars: 0);
        _output.WriteLine($"parent mapping: locatedVia={read.LocatedVia} in {read.LocateMs} ms.");
        Assert.False(string.IsNullOrEmpty(read.EntryId));
        Assert.True(read.Attachments.Count > 0, "the parent of an attachment hit must carry attachments");
    }

    [Fact]
    public void SeededMixedAttachments_AreAdmittedOnceIndexed_AndAlwaysCleanedUp()
    {
        // Hub-only seed (S2): the distinctive token exists ONLY inside the attachments -
        // never in the subject, never in the body - so finding it requires an
        // attachment-content row of kind calendar/communication/document.
        //
        // The ADMISSION half is asserted unconditionally by the corpus tests above; what
        // this adds is the seeded end-to-end path. The Windows Search gatherer is
        // asynchronous and OS-driven, so whether it has crawled a brand-new hub item
        // inside the budget is not something a test can command: the crawl outcome is
        // REPORTED, and only what is under our control is asserted.
        string hub = _fixture.Settings.TestHubStoreDisplayName;
        string marker = "atk" + Guid.NewGuid().ToString("N")[..12];
        string token = "zqattach" + Guid.NewGuid().ToString("N")[..10];
        string subject = LiveOutlookTestMailer.SubjectTag + " " + marker + " attachment-kind recall";
        string directory = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-" + marker);
        Directory.CreateDirectory(directory);

        try
        {
            string ics = Path.Combine(directory, "invite-" + marker + ".ics");
            File.WriteAllText(
                ics,
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//OutlookAI//test//EN\r\nBEGIN:VEVENT\r\n"
                + "UID:" + marker + "@outlookai.test\r\nDTSTAMP:20260101T090000Z\r\nDTSTART:20260101T090000Z\r\n"
                + "DTEND:20260101T093000Z\r\nSUMMARY:" + token + "\r\nDESCRIPTION:" + token
                + "\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

            string eml = Path.Combine(directory, "embedded-" + marker + ".eml");
            File.WriteAllText(
                eml,
                "From: nobody@example.invalid\r\nTo: nobody@example.invalid\r\nSubject: " + token
                + "\r\nDate: Thu, 1 Jan 2026 09:00:00 +0000\r\nMIME-Version: 1.0\r\n"
                + "Content-Type: text/plain; charset=utf-8\r\n\r\n" + token + "\r\n");

            string txt = Path.Combine(directory, "note-" + marker + ".txt");
            File.WriteAllText(txt, token + Environment.NewLine);

            string entryId = LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
                hub, subject, "Attachment-kind recall probe (agent-authored). " + marker, new[] { ics, eml, txt });
            Assert.False(string.IsNullOrEmpty(entryId));

            IIndexClient client = Client;
            string scope = _fixture.GetScope(hub).StorePrefix;
            string sql = "SELECT TOP 200 System.ItemUrl, System.Kind FROM SystemIndex WHERE SCOPE='" + scope
                + "' AND CONTAINS(System.Search.Contents, '\"" + token + "\"')";

            List<IReadOnlyDictionary<string, object?>> rows = new();
            Stopwatch waited = Stopwatch.StartNew();
            while (waited.Elapsed < TimeSpan.FromSeconds(90))
            {
                rows = client.ExecuteRows(sql, 200).ToList();
                if (rows.Any(r => IndexRowFilter.IsAttachmentRow(Url(r))))
                {
                    break;
                }

                Thread.Sleep(5000);
            }

            List<IReadOnlyDictionary<string, object?>> attachmentRows =
                rows.Where(r => IndexRowFilter.IsAttachmentRow(Url(r))).ToList();
            _output.WriteLine(
                $"seeded probe: {rows.Count} row(s) after {waited.Elapsed.TotalSeconds:F0} s, "
                + $"{attachmentRows.Count} attachment row(s); kinds: "
                + string.Join(", ", attachmentRows.SelectMany(Kinds).Distinct(StringComparer.OrdinalIgnoreCase)));

            if (attachmentRows.Count == 0)
            {
                _output.WriteLine(
                    "the gatherer had not crawled the seeded item inside the budget - admission is proven "
                    + "unconditionally by the corpus tests in this class.");
            }
            else
            {
                // Every one of them is admitted by the shipped shape...
                Assert.All(
                    attachmentRows,
                    r => Assert.True(IndexRowFilter.Keep(IndexRowMapper.Map(r), KindFilter.EmailAndDocuments)));

                // ...and they all map to ONE parent message URL.
                List<string> parents = attachmentRows
                    .Select(r => Url(r)!)
                    .Select(u => u[..u.LastIndexOf("/at=", StringComparison.Ordinal)])
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                Assert.Single(parents);
            }
        }
        finally
        {
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
                hub, marker, window: TimeSpan.FromSeconds(120), stableFor: TimeSpan.FromSeconds(10),
                folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
            Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(
                hub, marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive));

            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch (IOException)
            {
            }
        }
    }

    private static bool HasExtension(string? fileName, IReadOnlyList<string> extensions)
    {
        if (string.IsNullOrEmpty(fileName))
        {
            return false;
        }

        foreach (string extension in extensions)
        {
            if (fileName!.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static long BestOfThree(IIndexClient client, string sql, int maxRows)
    {
        client.ExecuteRows(sql, maxRows); // warm
        long best = long.MaxValue;
        for (int i = 0; i < 3; i++)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            client.ExecuteRows(sql, maxRows);
            stopwatch.Stop();
            best = Math.Min(best, stopwatch.ElapsedMilliseconds);
        }

        return best;
    }

    private static string? Url(IReadOnlyDictionary<string, object?>? row)
    {
        return row != null && row.TryGetValue("System.ItemUrl", out object? value) ? value as string : null;
    }

    private static IEnumerable<string> Kinds(IReadOnlyDictionary<string, object?> row)
    {
        if (!row.TryGetValue("System.Kind", out object? value) || value == null)
        {
            yield break;
        }

        if (value is string single)
        {
            yield return single;
            yield break;
        }

        if (value is IEnumerable<object> many)
        {
            foreach (object item in many)
            {
                if (item is string text)
                {
                    yield return text;
                }
            }

            yield break;
        }

        if (value is string[] strings)
        {
            foreach (string text in strings)
            {
                yield return text;
            }
        }
    }

    private static string Percent(int delta, int baseline)
    {
        return baseline == 0 ? "0.0" : (100.0 * delta / baseline).ToString("F1");
    }
}
