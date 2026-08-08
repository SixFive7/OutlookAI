using System.Reflection;
using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the Phase-3 T2 live tier: the MailService under test (one
/// ComGateway / pumped STA / held-open Outlook session - may START Outlook per S7/D17,
/// never stops it) plus an INDEPENDENT OutlookComSession used only to verify Outlook's
/// UI state from outside the service (ActiveExplorer/Inspectors are process-global in
/// Outlook, so the second COM client observes exactly what the service changed) and to
/// close windows the tests themselves opened.
/// </summary>
public sealed class LivePhase3Fixture : IDisposable
{
    private readonly Lazy<OutlookComSession> _verifySession;
    private readonly Lazy<IReadOnlyList<ComWalkedItem>> _hubCorpus;

    public LivePhase3Fixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
        _verifySession = new Lazy<OutlookComSession>(
            () => OutlookComSession.Connect(allowStartingOutlook: true),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _hubCorpus = new Lazy<IReadOnlyList<ComWalkedItem>>(
            () => VerifySession.WalkStoreMailItems(Settings.TestHubStoreDisplayName)
                .Where(i => i.ReceivedTime.HasValue && !HubCorpus.IsTestArtifact(i.Subject))
                .ToList(),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Independent COM session for UI-state verification and test-owned window cleanup.</summary>
    public OutlookComSession VerifySession => _verifySession.Value;

    /// <summary>
    /// Ground-truth corpus of the tiny test-hub store (S2: only that store is ever
    /// walked): mail items with ReceivedTime, test artifacts excluded.
    /// </summary>
    public IReadOnlyList<ComWalkedItem> TestHubCorpus => _hubCorpus.Value;

    /// <summary>Gitignored screenshots directory (v3.MD S6): McpServer/**/screenshots/.</summary>
    public string ScreenshotsDirectory
    {
        get
        {
            string testProjectDir =
                typeof(LivePhase3Fixture).Assembly
                    .GetCustomAttributes<AssemblyMetadataAttribute>()
                    .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
                ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
            return Path.Combine(testProjectDir, "screenshots");
        }
    }

    public void Dispose()
    {
        // Releases COM references only - Outlook keeps running (S7: never kill/close).
        if (_verifySession.IsValueCreated)
        {
            _verifySession.Value.Dispose();
        }

        Service.Dispose();
    }
}

[CollectionDefinition("LivePhase3")]
public sealed class LivePhase3Collection : ICollectionFixture<LivePhase3Fixture>
{
}

/// <summary>
/// Corpus text helpers shared by the Phase-3 live tests, mirroring the Phase-1
/// completeness oracle's parity rules: probe terms whose every corpus occurrence is a
/// clean word occurrence (so regex ground truth, the index word breaker and DASL
/// ci_phrasematch all agree), and the S3 test-artifact subject tag filter.
/// </summary>
internal static class HubCorpus
{
    public const string TestArtifactTag = "[OutlookAI-McpTest]";

    public static bool IsTestArtifact(string? subject)
    {
        return subject != null && subject.Contains(TestArtifactTag, StringComparison.OrdinalIgnoreCase);
    }

    public static string TextOf(ComWalkedItem item)
    {
        return (item.Subject ?? string.Empty) + "\n" + (item.Body ?? string.Empty);
    }

    public static Regex WordRegex(string term)
    {
        return new Regex(
            "(?<![A-Za-z0-9])" + Regex.Escape(term) + "(?![A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    /// <summary>
    /// ASCII probe terms derived from the corpus, ordered by match count (descending,
    /// then alphabetically). A term qualifies when every occurrence is a clean word
    /// occurrence (no letter/digit neighbors).
    /// </summary>
    public static IReadOnlyList<string> RankedCleanTerms(IReadOnlyList<ComWalkedItem> corpus)
    {
        var tokenRegex = new Regex("[A-Za-z]{4,}", RegexOptions.CultureInvariant);
        List<string> texts = corpus.Select(TextOf).ToList();

        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string text in texts)
        {
            foreach (Match m in tokenRegex.Matches(text))
            {
                tokens.Add(m.Value.ToLowerInvariant());
            }
        }

        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (string token in tokens)
        {
            if (!AllOccurrencesAreCleanWords(texts, token))
            {
                continue;
            }

            Regex word = WordRegex(token);
            int count = texts.Count(t => word.IsMatch(t));
            if (count >= 1)
            {
                counts[token] = count;
            }
        }

        return counts
            .OrderByDescending(kv => kv.Value)
            .ThenBy(kv => kv.Key, StringComparer.Ordinal)
            .Select(kv => kv.Key)
            .ToList();
    }

    private static bool AllOccurrencesAreCleanWords(List<string> texts, string token)
    {
        foreach (string text in texts)
        {
            int start = 0;
            while (true)
            {
                int idx = text.IndexOf(token, start, StringComparison.OrdinalIgnoreCase);
                if (idx < 0)
                {
                    break;
                }

                bool beforeOk = idx == 0 || !char.IsLetterOrDigit(text[idx - 1]);
                int end = idx + token.Length;
                bool afterOk = end >= text.Length || !char.IsLetterOrDigit(text[end]);
                if (!beforeOk || !afterOk)
                {
                    return false;
                }

                start = idx + 1;
            }
        }

        return true;
    }
}
