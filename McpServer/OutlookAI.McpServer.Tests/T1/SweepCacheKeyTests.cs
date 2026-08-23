using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The sweep cache was unreachable for every UNSCOPED search, and the reason was that one
/// value was doing two jobs.
/// <para>
/// THE SHAPE. <c>MailService.ResolveSweepWindows</c> gives an unscoped search a FALLBACK
/// window of <c>UtcNow - EmptyIndexSweepWindow</c> - the widest window, for the store whose
/// own frontier could not be measured - and that same value was handed to the cache as the
/// key (<c>SweepCache.TryGetUsable</c> compares it for exact equality). A wall-clock value
/// differs on every call, so no two unscoped searches ever shared a key and none of them
/// could hit: every unscoped search paid a full COM sweep, the most expensive thing on the
/// search path. Before <c>c515565</c> the key was the profile frontier and it hit.
/// </para>
/// <para>
/// IT WAS COST, NOT CORRECTNESS, AND THE FIX MUST KEEP IT THAT WAY. Sweeping live errs
/// toward completeness, so nothing was ever reported wrongly - which means this change may
/// not make anything start being reported wrongly either. That is what the last two tests
/// here are for: a frontier advance still throws the entry away, and a profile whose index
/// knows nothing still sweeps live every time.
/// </para>
/// <para>
/// The two roles are now two fields, and both halves are pinned here: the cache is keyed on
/// the profile FRONTIER (stable while the index is stable, invalidated the moment it
/// ingests), while the fallback WINDOW handed to Outlook is still read off the WALL CLOCK -
/// keying that on the frontier would hand the one store whose window cannot be measured the
/// narrowest window on the profile, which is the defect the wall-clock fallback was
/// introduced to fix.
/// </para>
/// <para>
/// Driven through the real <see cref="MailService"/> against a stand-in session and a
/// stand-in index client. No mailbox and no Windows Search index are touched.
/// </para>
/// </summary>
public sealed class SweepCacheKeyTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string IndexedStore = "alice@example.com";

    private const string IndexedPrefix = "mapi16://" + Sid + "/" + IndexedStore + "($deadbeef)";

    /// <summary>A data file the index catalog has never heard of - the fallback window's own case.</summary>
    private const string PstStore = "Archive 2019.pst";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    /// <summary><c>MailService.SweepSafetyMargin</c>, which is private and subtracted from every window.</summary>
    private static readonly TimeSpan SafetyMargin = TimeSpan.FromMinutes(10);

    private static readonly ComStoreDetail[] ProfileStores =
    {
        new ComStoreDetail(IndexedStore, "store-alice", 0, true),
        new ComStoreDetail(PstStore, "store-pst", 3, null),
    };

    // ============================================================ the cache is reachable

    [Fact]
    public void ThreeUnscopedSearchesInsideTheTtl_PayForOneSweep()
    {
        // The defect, end to end: before the fix this recorded three COM sweeps and reported
        // cached: null three times, because each call keyed on its own reading of the wall
        // clock. The searches are separated by a real clock TICK on purpose - DateTime.UtcNow
        // moves in ~15.6 ms steps on Windows, so calls close enough together could share a
        // key by accident, and a test that passed for that reason would prove nothing.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, Index());

        SearchOutcome first = service.Search(Unscoped());
        WaitForTheWallClockToTick();
        SearchOutcome second = service.Search(Unscoped());
        WaitForTheWallClockToTick();
        SearchOutcome third = service.Search(Unscoped());

        Assert.Equal(1, session.SweepCount);
        Assert.Null(first.Sweep!.Cached);
        Assert.True(second.Sweep!.Cached);
        Assert.True(third.Sweep!.Cached);
    }

    [Fact]
    public void TheCachedAnswer_StillCarriesTheSweepsCoverage()
    {
        // Reuse has to mean the same answer, not merely a faster one - a cache that served a
        // hit while dropping the counters behind it would report coverage it did not have.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, Index());

        SearchOutcome first = service.Search(Unscoped());
        WaitForTheWallClockToTick();
        SearchOutcome second = service.Search(Unscoped());

        Assert.True(second.Sweep!.Cached);
        Assert.Equal(first.Sweep!.FoldersSwept, second.Sweep.FoldersSwept);
        Assert.Equal(first.Hits.Count, second.Hits.Count);

        // And it says so, rather than presenting a seconds-old live check as a current one.
        Assert.Contains(FreshMerge.GapCachedSweep, second.Sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, second.Freshness);
    }

    // ====================================================== the window is still wall clock

    [Fact]
    public void TheFallbackWindowHandedToOutlook_IsStillReadOffTheWallClock()
    {
        // The half that must NOT change. The fallback is the window for a store whose
        // frontier could not be measured, so it is the WIDEST window on the profile; keying
        // the cache on the frontier must not drag the window along with it, or the one store
        // that needs seven days would get the profile's newest instant instead.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, Index());

        DateTime before = DateTime.UtcNow;
        service.Search(Unscoped());
        DateTime after = DateTime.UtcNow;

        DateTime since = Assert.Single(session.SinceArguments);
        Assert.InRange(
            since,
            before - MailService.EmptyIndexSweepWindow - SafetyMargin,
            after - MailService.EmptyIndexSweepWindow - SafetyMargin);

        // Stated as an inequality too, because the range above would also be satisfied by a
        // frontier that happened to sit inside it: these are seven days apart by construction
        // and confusing them is exactly what this change had to avoid.
        Assert.NotEqual(Frontier - SafetyMargin, since);
    }

    [Fact]
    public void AStoreTheIndexCanName_StillGetsItsOwnFrontierWindow()
    {
        // The other side of the same statement: the per-store windows are unchanged, so the
        // catalogued store is still swept back minutes rather than days.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, Index());

        service.Search(Unscoped());

        IReadOnlyDictionary<string, DateTime> perStore = Assert.Single(session.PerStoreArguments)!;
        Assert.Equal(Frontier - SafetyMargin, perStore[IndexedStore]);
        Assert.False(perStore.ContainsKey(PstStore));
    }

    [Fact]
    public void AnUnscopedSweep_StillServesAFollowUpSearchOfONEStore()
    {
        // The cross-scope reuse the cache has always had - a store-scoped request served from
        // the all-stores entry, sound because every store gets the identical folder set - and
        // it turns on the key and the per-store windows meaning the same thing. Moving the key
        // without moving them would break this and nothing else would notice.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, Index());

        service.Search(Unscoped());
        WaitForTheWallClockToTick();
        SearchOutcome scoped = service.Search(
            new SearchRequest { Query = "test", Store = IndexedStore, Top = 25, SnippetChars = 0 });

        Assert.Equal(1, session.SweepCount);
        Assert.True(scoped.Sweep!.Cached);
    }

    // ================================================ the reuse rule is still conservative

    [Fact]
    public void AFrontierAdvance_ThrowsTheUnscopedEntryAway()
    {
        // The property that makes the key safe to reuse at all: the entry is keyed on what
        // the index has ingested, so the moment it ingests anything the entry stops matching
        // and the next search sweeps live. Without this the cache could outlive the state it
        // was taken against, which is the one way a cost fix turns into a recall bug.
        StandInSession session = new StandInSession();
        StubIndexClient index = Index();
        using MailService service = Service(session, index);

        service.Search(Unscoped());
        index.Frontier = Frontier.AddMinutes(3);
        service.Search(Unscoped());

        Assert.Equal(2, session.SweepCount);
    }

    [Fact]
    public void AProfileWithNoIndexedMailAtAll_StillSweepsLiveEveryTime()
    {
        // No frontier means nothing whose advance could invalidate an entry, so the key falls
        // back to the wall clock and every search pays a live sweep. That is the pre-existing
        // behaviour and the safe direction: on this profile the sweep is the ONLY tier that
        // can find anything, and a 10 s TTL is a weaker promise than the rest of this cache
        // makes.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, NothingIndexed());

        service.Search(Unscoped());
        WaitForTheWallClockToTick();
        service.Search(Unscoped());
        WaitForTheWallClockToTick();
        service.Search(Unscoped());

        Assert.Equal(3, session.SweepCount);
    }

    [Fact]
    public void AnAfterNarrowedSearch_IsStillNeverCached()
    {
        // Unchanged, and re-pinned here because the key moved: a request whose 'after' clamped
        // the window took a NARROWER sweep than an unclamped one would, so caching it would
        // poison every wider follow-up search.
        StandInSession session = new StandInSession();
        using MailService service = Service(session, Index());

        SearchRequest clamped = Unscoped();
        clamped.AfterUtc = DateTime.UtcNow.AddDays(-1);
        service.Search(clamped);
        WaitForTheWallClockToTick();
        service.Search(clamped);

        Assert.Equal(2, session.SweepCount);
    }

    // ======================================================================== fixtures

    private static SearchRequest Unscoped()
    {
        return new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 };
    }

    /// <summary>
    /// Blocks until <see cref="DateTime.UtcNow"/> reports a different value. The system clock
    /// advances in ~15.6 ms steps on Windows while <c>MonotonicClock</c> (which ages the cache
    /// entry) is stopwatch-backed and always moves, so without this two searches could share a
    /// wall-clock reading and the cache-hit assertions would pass for the wrong reason.
    /// </summary>
    private static void WaitForTheWallClockToTick()
    {
        DateTime start = DateTime.UtcNow;
        while (DateTime.UtcNow == start)
        {
            Thread.Sleep(1);
        }
    }

    private static MailService Service(StandInSession session, StubIndexClient index)
    {
        return new MailService(new DirectGateway(session.AsSession()), null, index);
    }

    /// <summary>The mixed profile: one store the catalog knows, one data file it does not.</summary>
    private static StubIndexClient Index() => new(new[] { IndexedPrefix });

    /// <summary>A profile Windows Search holds no mail for at all.</summary>
    private static StubIndexClient NothingIndexed() => new(Array.Empty<string>());

    /// <summary>
    /// A Windows Search stand-in that knows a chosen SET of store prefixes and answers the
    /// probe statements by shape. Its frontier is mutable so a test can make the index ingest
    /// something between two searches.
    /// </summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        private readonly IReadOnlyList<string> _knownPrefixes;

        internal StubIndexClient(IReadOnlyList<string> knownPrefixes)
        {
            _knownPrefixes = knownPrefixes;
        }

        /// <summary>The newest indexed instant this stand-in reports, profile-wide and per store.</summary>
        internal DateTime Frontier { get; set; } = SweepCacheKeyTests.Frontier;

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
            string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql.EndsWith(DiscoveryTail, StringComparison.Ordinal))
            {
                return _knownPrefixes
                    .Select(p => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.ItemUrl"] = p + "/0/Inbox/sampled-item",
                    })
                    .ToList();
            }

            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return Known(sql)
                    ? Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.Message.DateReceived"] = Frontier,
                    })
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return Known(sql)
                    ? Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.ItemUrl"] = _knownPrefixes[0] + "/0/Inbox/probed-item",
                    })
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            return Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(
            IReadOnlyDictionary<string, object?> row)
        {
            return new[] { row };
        }

        private bool Known(string sql)
        {
            int start = sql.IndexOf("SCOPE='", StringComparison.Ordinal);
            if (start < 0)
            {
                return _knownPrefixes.Count > 0;
            }

            start += "SCOPE='".Length;
            int end = sql.IndexOf('\'', start);
            string scope = end < 0 ? sql.Substring(start) : sql.Substring(start, end - start);

            return _knownPrefixes.Any(p =>
                string.Equals(scope, p, StringComparison.OrdinalIgnoreCase)
                || scope.StartsWith(p + "/0/", StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>Runs operations straight against the stand-in session (no COM host, no pipe).</summary>
    private sealed class DirectGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal DirectGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        public T Run<T>(Func<IOutlookSession, T> operation) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => operation(_session);

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }

    /// <summary>
    /// Records every sweep the service asks Outlook for, with the window arguments it was
    /// asked for them with. The COUNT is the whole point: a cache hit is invisible in the
    /// payload's window fields and visible only as a COM call that did not happen.
    /// </summary>
    private sealed class StandInSession
    {
        private readonly List<DateTime> _since = new List<DateTime>();
        private readonly List<IReadOnlyDictionary<string, DateTime>?> _perStore =
            new List<IReadOnlyDictionary<string, DateTime>?>();

        internal int SweepCount => _since.Count;

        internal IReadOnlyList<DateTime> SinceArguments => _since;

        internal IReadOnlyList<IReadOnlyDictionary<string, DateTime>?> PerStoreArguments => _perStore;

        internal IOutlookSession AsSession() => RecordingSession.Create(this);

        private ComSweepResult Sweep(DateTime sinceUtc, IReadOnlyDictionary<string, DateTime>? perStoreSinceUtc)
        {
            _since.Add(sinceUtc);
            _perStore.Add(perStoreSinceUtc);

            return new ComSweepResult(
                new[] { Mail("AA1", IndexedStore, "store-alice"), Mail("BB1", PstStore, "store-pst") },
                foldersSwept: 8,
                foldersSkipped: 0,
                sweptFolders: new[] { IndexedStore + "/Inbox", PstStore + "/Inbox" },
                perStore: new[]
                {
                    new ComStoreSweepCounters(IndexedStore, 4, 0, 0, 0),
                    new ComStoreSweepCounters(PstStore, 4, 0, 0, 0),
                });
        }

        private static ComMailBrief Mail(string entryId, string store, string storeId)
        {
            return new ComMailBrief(
                entryId: entryId,
                storeDisplayName: store,
                storeId: storeId,
                folderName: "Inbox",
                folderKind: "inbox",
                subject: "a test mail",
                senderName: "Bob",
                senderAddress: "bob@example.com",
                receivedTime: Frontier,
                isRead: true,
                hasAttachments: false,
                sizeBytes: 2048,
                body: "test body");
        }

        /// <summary>
        /// A <see cref="DispatchProxy"/> so a member added to the contract needs no change
        /// here. Not sealed: DispatchProxy derives from its TProxy at runtime.
        /// </summary>
        private class RecordingSession : DispatchProxy
        {
            private StandInSession _owner = null!;

            internal static IOutlookSession Create(StandInSession owner)
            {
                object proxy = Create<IOutlookSession, RecordingSession>()
                    ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
                ((RecordingSession)proxy)._owner = owner;
                return (IOutlookSession)proxy;
            }

            /// <inheritdoc />
            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
            {
                switch (targetMethod?.Name)
                {
                    case nameof(IOutlookSession.GetProfileName):
                        return "T1 stand-in profile";

                    case nameof(IOutlookSession.GetStoreDetails):
                        return ProfileStores;

                    // Arguments 0 and 6 are sinceUtc and perStoreSinceUtc; see
                    // IOutlookSession.SweepFoldersNewerThan.
                    case nameof(IOutlookSession.SweepFoldersNewerThan):
                        return _owner.Sweep(
                            (DateTime)args![0]!,
                            args[6] as IReadOnlyDictionary<string, DateTime>);

                    default:
                        throw new NotSupportedException(
                            "The stand-in session was asked for " + (targetMethod?.Name ?? "an unnamed member")
                            + ", which this test does not model.");
                }
            }
        }
    }
}
