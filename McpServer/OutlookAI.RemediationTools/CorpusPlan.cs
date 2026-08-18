using System.Globalization;
using System.Text;

namespace OutlookAI.RemediationTools;

/// <summary>One body-size class of the corpus mixture and how often it is drawn.</summary>
/// <param name="Name">Stable identifier; appears in the manifest and the plan report.</param>
/// <param name="Weight">Relative frequency. Only the ratio between weights matters.</param>
/// <param name="MinBytes">Inclusive lower bound of the body length in this class.</param>
/// <param name="MaxBytes">Exclusive upper bound of the body length in this class.</param>
public sealed record CorpusSizeClass(string Name, int Weight, int MinBytes, int MaxBytes);

/// <summary>
/// One age band, expressed as "between MinDaysAgo and MaxDaysAgo before the anchor".
/// Bands exist so a measurement can state up front how many items a 7-day window and a
/// 60-day window are each expected to select - with one smooth age curve those two
/// numbers would be arithmetic on a distribution rather than a property of the plan
/// that anyone can read off before a single item is created.
/// </summary>
/// <param name="Name">Stable identifier; appears in the plan report.</param>
/// <param name="Weight">Relative frequency. Only the ratio between weights matters.</param>
/// <param name="MinDaysAgo">Inclusive lower bound of the age in days.</param>
/// <param name="MaxDaysAgo">Exclusive upper bound of the age in days.</param>
public sealed record CorpusDateBand(string Name, int Weight, int MinDaysAgo, int MaxDaysAgo);

/// <summary>How much of the corpus lands in one folder of the target store.</summary>
/// <param name="FolderId">Outlook default-folder id (6 Inbox, 5 Sent Items, 3 Deleted Items, 23 Junk Email).</param>
/// <param name="Name">Human-readable folder name for reports.</param>
/// <param name="Weight">Relative frequency. Only the ratio between weights matters.</param>
public sealed record CorpusFolderShare(int FolderId, string Name, int Weight);

/// <summary>
/// Everything that decides WHAT the corpus is. Two runs with an equal
/// <see cref="ShapeKey"/> describe byte-identical items, which is what lets a
/// measurement be repeated after a VM rollback.
/// <para>
/// <see cref="AnchorUtc"/> is deliberately a REQUIRED parameter rather than "now". A
/// corpus anchored on the wall clock would drift every time it was extended or rebuilt,
/// so the same seed would stop meaning the same corpus and every window measurement
/// taken against it would be comparing different populations.
/// </para>
/// </summary>
/// <param name="CorpusId">Short identifier embedded in every subject; distinguishes corpora within one store.</param>
/// <param name="Seed">The only source of randomness. Same seed, same corpus.</param>
/// <param name="AnchorUtc">The instant ages are measured back from. Never the current time.</param>
public sealed record CorpusPlanOptions(string CorpusId, long Seed, DateTime AnchorUtc)
{
    /// <summary>
    /// Default body-size mixture: mostly small mail, with a deliberate tail. The two
    /// largest classes are a reason the corpus exists - the sweep body cap and the
    /// response frame budget are both sized against quoted threads of tens of KB and the
    /// occasional monster of a few hundred KB, and a uniform corpus never produces one.
    /// <para>
    /// The top class deliberately runs PAST
    /// <see cref="OutlookAI.Core.Com.OutlookComSession.SweepBodyCharsCap"/> (500 000
    /// chars), so roughly a fifth of it trips the per-item body cap. A corpus that
    /// stopped just below the cap could never show what capping costs or that the
    /// <c>itemsBodyCapped</c> counter is wired correctly.
    /// </para>
    /// </summary>
    public static IReadOnlyList<CorpusSizeClass> DefaultSizeClasses { get; } = new[]
    {
        new CorpusSizeClass("short", 45, 200, 1_200),
        new CorpusSizeClass("normal", 33, 1_200, 6_000),
        new CorpusSizeClass("long", 15, 6_000, 24_000),
        new CorpusSizeClass("quoted-thread", 6, 24_000, 96_000),
        new CorpusSizeClass("huge-thread", 1, 96_000, 640_000),
    };

    /// <summary>
    /// Default age mixture. The first three bands are cut exactly at the 1-day, 7-day and
    /// 60-day marks the freshness sweep and the exhaustive scan are measured on, so the
    /// expected selection of each window is a band count rather than an estimate.
    /// </summary>
    public static IReadOnlyList<CorpusDateBand> DefaultDateBands { get; } = new[]
    {
        new CorpusDateBand("last-24h", 2, 0, 1),
        new CorpusDateBand("1d-7d", 6, 1, 7),
        new CorpusDateBand("7d-60d", 17, 7, 60),
        new CorpusDateBand("60d-1y", 30, 60, 365),
        new CorpusDateBand("1y-4y", 45, 365, 1_460),
    };

    /// <summary>Default folder mixture across the four folders the measurement sweeps.</summary>
    public static IReadOnlyList<CorpusFolderShare> DefaultFolders { get; } = new[]
    {
        new CorpusFolderShare(6, "Inbox", 55),
        new CorpusFolderShare(5, "Sent Items", 25),
        new CorpusFolderShare(3, "Deleted Items", 12),
        new CorpusFolderShare(23, "Junk Email", 8),
    };

    /// <summary>Body-size mixture. Defaults to <see cref="DefaultSizeClasses"/>.</summary>
    public IReadOnlyList<CorpusSizeClass> SizeClasses { get; init; } = DefaultSizeClasses;

    /// <summary>Age mixture. Defaults to <see cref="DefaultDateBands"/>.</summary>
    public IReadOnlyList<CorpusDateBand> DateBands { get; init; } = DefaultDateBands;

    /// <summary>Folder mixture. Defaults to <see cref="DefaultFolders"/>.</summary>
    public IReadOnlyList<CorpusFolderShare> Folders { get; init; } = DefaultFolders;

    /// <summary>
    /// A stable digest of everything except the item COUNT, so a resumed or extended run
    /// can prove it is adding to the same corpus. The count is excluded on purpose: item
    /// N's description never depends on how many items were asked for, which is what
    /// makes "build 20 000 more" an addition rather than a silent rewrite of the first
    /// 20 000 (see <see cref="CorpusPlan.Describe"/>).
    /// </summary>
    public string ShapeKey
    {
        get
        {
            var sb = new StringBuilder();
            sb.Append("v1|").Append(CorpusId).Append('|').Append(Seed.ToString(CultureInfo.InvariantCulture))
                .Append('|')
                .Append(DateTime.SpecifyKind(AnchorUtc, DateTimeKind.Utc).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
            foreach (CorpusSizeClass c in SizeClasses)
            {
                sb.Append("|s:").Append(c.Name).Append(':').Append(c.Weight)
                    .Append(':').Append(c.MinBytes).Append(':').Append(c.MaxBytes);
            }

            foreach (CorpusDateBand b in DateBands)
            {
                sb.Append("|d:").Append(b.Name).Append(':').Append(b.Weight)
                    .Append(':').Append(b.MinDaysAgo).Append(':').Append(b.MaxDaysAgo);
            }

            foreach (CorpusFolderShare f in Folders)
            {
                sb.Append("|f:").Append(f.FolderId).Append(':').Append(f.Weight);
            }

            return sb.ToString();
        }
    }
}

/// <summary>One item the corpus is supposed to contain. Pure data - no COM, no I/O.</summary>
/// <param name="Ordinal">1-based position in the corpus; encoded in the subject.</param>
/// <param name="FolderId">Outlook default-folder id the item belongs in.</param>
/// <param name="Subject">Carries both subject tags and the ordinal.</param>
/// <param name="BodyBytes">Exact length of the body <see cref="CorpusPlan.BuildBody"/> produces.</param>
/// <param name="ReceivedUtc">Intended PR_MESSAGE_DELIVERY_TIME.</param>
/// <param name="SentUtc">Intended PR_CLIENT_SUBMIT_TIME; never later than the received instant.</param>
/// <param name="IsRead">Intended read state.</param>
/// <param name="SizeClass">Which size class produced the body length.</param>
/// <param name="DateBand">Which age band produced the received instant.</param>
public sealed record CorpusItemSpec(
    int Ordinal,
    int FolderId,
    string Subject,
    int BodyBytes,
    DateTime ReceivedUtc,
    DateTime SentUtc,
    bool IsRead,
    string SizeClass,
    string DateBand);

/// <summary>
/// The pure half of the corpus generator: given a seed and a shape it says exactly what
/// item number N is, with no clock, no COM and no I/O. Everything the T1 tier can pin
/// about the corpus lives here; <see cref="ComCorpusMailbox"/> only carries these
/// descriptions into a PST.
/// <para>
/// <b>Determinism.</b> Every field is derived by hashing (seed, ordinal, field) - there
/// is no mutable generator state and no <c>Random</c>, whose sequence for a given seed
/// is explicitly not a contract across framework versions. Two consequences the design
/// leans on: items may be produced in any order, or across any number of interrupted
/// runs, and still come out identical; and item N never changes when the requested item
/// COUNT changes, so extending a corpus adds to it instead of invalidating it.
/// </para>
/// </summary>
public sealed class CorpusPlan
{
    /// <summary>
    /// The mailbox-safety tag every test-created item carries (CLAUDE.md mailbox rule 2).
    /// Corpus items carry it as well as their own tag, so the project's existing tested
    /// purge can still find them if a corpus is ever built somewhere it should not have
    /// been. It is the same constant the remediation console matches, deliberately.
    /// </summary>
    public const string SubjectTag = RemediationRules.SubjectTag;

    /// <summary>Opening delimiter of the corpus tag: <c>[OutlookAI-Corpus:&lt;id&gt;#&lt;ordinal&gt;]</c>.</summary>
    public const string CorpusTagOpen = "[OutlookAI-Corpus:";

    /// <summary>
    /// Bracket-free fragment for a DASL LIKE prefilter, matching the pattern the live
    /// suite uses. A superset of real matches by construction, so a LIKE count of zero
    /// proves a corpus count of zero; the authoritative per-item test is always the
    /// ordinal <see cref="TryParseOrdinal"/> on the re-read subject.
    /// </summary>
    public const string DaslCountFragment = "OutlookAI-Corpus";

    private static readonly string[] Vocabulary =
    {
        "invoice", "renewal", "handover", "provisioning", "porting", "sip", "trunk", "outage",
        "maintenance", "window", "escalation", "contract", "quote", "migration", "rollout",
        "onboarding", "offboarding", "credentials", "firewall", "latency", "jitter", "codec",
        "voicemail", "routing", "dial", "plan", "tariff", "peering", "capacity", "handset",
        "provision", "ticket", "incident", "review", "approval", "budget", "forecast",
        "schedule", "deployment", "certificate", "expiry", "backup", "restore", "failover",
    };

    private static readonly string[] SubjectPrefixes = { "", "", "", "RE: ", "RE: ", "FW: " };

    private readonly CorpusPlanOptions _options;
    private readonly int _sizeWeightTotal;
    private readonly int _dateWeightTotal;
    private readonly int _folderWeightTotal;

    /// <summary>Builds a plan and validates the shape. Throws on a shape no corpus could satisfy.</summary>
    public CorpusPlan(CorpusPlanOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (string.IsNullOrWhiteSpace(options.CorpusId))
        {
            throw new ArgumentException("A corpus needs an id - it goes in every subject.", nameof(options));
        }

        foreach (char c in options.CorpusId)
        {
            // The id lands inside the subject tag that teardown matches ordinally, so it
            // must not contain a delimiter of that tag or a DASL wildcard character.
            if (!char.IsAsciiLetterOrDigit(c) && c != '-' && c != '_')
            {
                throw new ArgumentException("Corpus id must be ASCII letters, digits, '-' or '_'.", nameof(options));
            }
        }

        if (options.AnchorUtc.Kind == DateTimeKind.Local)
        {
            throw new ArgumentException("Anchor must be UTC, never local - a local anchor moves with the VM's time zone.", nameof(options));
        }

        _sizeWeightTotal = SumWeights(options.SizeClasses.Select(c => c.Weight), "size classes");
        _dateWeightTotal = SumWeights(options.DateBands.Select(b => b.Weight), "date bands");
        _folderWeightTotal = SumWeights(options.Folders.Select(f => f.Weight), "folders");

        foreach (CorpusSizeClass c in options.SizeClasses)
        {
            if (c.MinBytes < 1 || c.MaxBytes <= c.MinBytes)
            {
                throw new ArgumentException($"Size class '{c.Name}' has an empty byte range.", nameof(options));
            }
        }

        foreach (CorpusDateBand b in options.DateBands)
        {
            if (b.MinDaysAgo < 0 || b.MaxDaysAgo <= b.MinDaysAgo)
            {
                throw new ArgumentException($"Date band '{b.Name}' has an empty day range.", nameof(options));
            }
        }

        _options = options;
    }

    /// <summary>The shape this plan was built from.</summary>
    public CorpusPlanOptions Options => _options;

    /// <summary>
    /// What item <paramref name="ordinal"/> is. Depends only on the seed, the ordinal and
    /// the shape - never on the total count, the clock, or what has already been built.
    /// </summary>
    public CorpusItemSpec Describe(int ordinal)
    {
        if (ordinal < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(ordinal), "Ordinals are 1-based.");
        }

        CorpusFolderShare folder = PickWeighted(_options.Folders, f => f.Weight, _folderWeightTotal, Draw(ordinal, StreamFolder));
        CorpusSizeClass sizeClass = PickWeighted(_options.SizeClasses, c => c.Weight, _sizeWeightTotal, Draw(ordinal, StreamSizeClass));
        CorpusDateBand band = PickWeighted(_options.DateBands, b => b.Weight, _dateWeightTotal, Draw(ordinal, StreamDateBand));

        int bodyBytes = sizeClass.MinBytes
            + (int)(Draw(ordinal, StreamBodyBytes) % (ulong)(sizeClass.MaxBytes - sizeClass.MinBytes));

        // Age in whole seconds inside the band, so two items in the same band are almost
        // never simultaneous and an ORDER BY on the received date is well defined.
        long bandSeconds = (long)(band.MaxDaysAgo - band.MinDaysAgo) * 86_400L;
        long ageSeconds = ((long)band.MinDaysAgo * 86_400L)
            + (long)(Draw(ordinal, StreamAgeSeconds) % (ulong)bandSeconds);
        DateTime receivedUtc = DateTime.SpecifyKind(_options.AnchorUtc, DateTimeKind.Utc).AddSeconds(-ageSeconds);

        // A sent item is its own origin, so submit time equals delivery time there. For
        // everything else the submit time precedes delivery by a transport delay, which is
        // what a real message looks like to a filter that reads either date property - and
        // the freshness sweep's DASL filter reads both.
        bool isSentFolder = folder.FolderId == SentItemsFolderId;
        int transportSeconds = isSentFolder ? 0 : (int)(Draw(ordinal, StreamTransport) % 600UL);
        DateTime sentUtc = receivedUtc.AddSeconds(-transportSeconds);

        // Sent and deleted mail reads as read; a minority of inbox/junk mail stays unread,
        // so an unread-only filter selects a genuinely different population.
        bool isRead = isSentFolder
            || folder.FolderId == DeletedItemsFolderId
            || Draw(ordinal, StreamRead) % 100UL >= 22UL;

        return new CorpusItemSpec(
            ordinal,
            folder.FolderId,
            BuildSubject(ordinal),
            bodyBytes,
            receivedUtc,
            sentUtc,
            isRead,
            sizeClass.Name,
            band.Name);
    }

    /// <summary>
    /// The subject of item <paramref name="ordinal"/>: the mailbox-safety tag, then the
    /// corpus tag carrying the id and the zero-padded ordinal, then readable words. Both
    /// tags come first so an ordinal Contains still finds them if a mail client later
    /// prepends its own "RE:".
    /// </summary>
    public string BuildSubject(int ordinal)
    {
        ulong r = Draw(ordinal, StreamSubject);
        var sb = new StringBuilder(96);
        sb.Append(SubjectTag).Append(CorpusTagOpen).Append(_options.CorpusId).Append('#')
            .Append(ordinal.ToString("D7", CultureInfo.InvariantCulture)).Append("] ")
            .Append(SubjectPrefixes[(int)(r % (ulong)SubjectPrefixes.Length)]);
        int words = 3 + (int)((r >> 8) % 5UL);
        for (int i = 0; i < words; i++)
        {
            if (i > 0)
            {
                sb.Append(' ');
            }

            sb.Append(Vocabulary[(int)(Draw(ordinal, StreamSubject + 1 + i) % (ulong)Vocabulary.Length)]);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Reads the ordinal back out of a subject, ORDINALLY and with no pattern matching of
    /// any kind - this is the parse a teardown leans on, and the incident this codebase
    /// carries scars from was caused by treating a subject as a wildcard pattern where
    /// the tag's own brackets became a character class. Returns false unless the subject
    /// carries the mailbox-safety tag AND a corpus tag whose id matches exactly.
    /// </summary>
    public static bool TryParseOrdinal(string? subject, string? corpusId, out int ordinal)
    {
        ordinal = 0;
        if (subject == null || string.IsNullOrEmpty(corpusId))
        {
            return false;
        }

        if (!subject.Contains(SubjectTag, StringComparison.Ordinal))
        {
            return false;
        }

        string open = CorpusTagOpen + corpusId + "#";
        int start = subject.IndexOf(open, StringComparison.Ordinal);
        if (start < 0)
        {
            return false;
        }

        int digitsStart = start + open.Length;
        int close = subject.IndexOf(']', digitsStart);
        if (close <= digitsStart)
        {
            return false;
        }

        string digits = subject.Substring(digitsStart, close - digitsStart);
        return int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out ordinal) && ordinal >= 1;
    }

    /// <summary>
    /// The body of an item, exactly <see cref="CorpusItemSpec.BodyBytes"/> ASCII
    /// characters long. Larger bodies are built as a quoted reply chain rather than one
    /// wall of text, because that is the shape a real long thread has and it is quoted
    /// history - not prose - that the sweep body cap is sized against.
    /// </summary>
    public string BuildBody(CorpusItemSpec spec)
    {
        ArgumentNullException.ThrowIfNull(spec);
        var sb = new StringBuilder(spec.BodyBytes + 128);
        int paragraph = 0;
        while (sb.Length < spec.BodyBytes)
        {
            int quoteDepth = paragraph == 0 ? 0 : Math.Min(paragraph, 6);
            string quote = new('>', quoteDepth);
            if (quoteDepth > 0)
            {
                sb.Append(quote).Append(" On ")
                    .Append(spec.SentUtc.AddDays(-quoteDepth).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture))
                    .Append(" a correspondent wrote:\n");
            }

            int sentences = 2 + (int)(Draw(spec.Ordinal, StreamBody + paragraph) % 4UL);
            for (int s = 0; s < sentences && sb.Length < spec.BodyBytes; s++)
            {
                if (quoteDepth > 0)
                {
                    sb.Append(quote).Append(' ');
                }

                int words = 6 + (int)(Draw(spec.Ordinal, StreamBody + (paragraph * 32) + s + 1) % 12UL);
                for (int w = 0; w < words; w++)
                {
                    if (w > 0)
                    {
                        sb.Append(' ');
                    }

                    sb.Append(Vocabulary[
                        (int)(Draw(spec.Ordinal, StreamBody + (paragraph * 512) + (s * 32) + w) % (ulong)Vocabulary.Length)]);
                }

                sb.Append(".\n");
            }

            sb.Append('\n');
            paragraph++;
        }

        // Exact length is the contract: the measurement correlates elapsed time against
        // bytes moved, so "roughly this big" would blunt the number the corpus exists for.
        return sb.ToString(0, spec.BodyBytes);
    }

    /// <summary>
    /// Walks ordinals <paramref name="from"/>..<paramref name="to"/> and counts what the
    /// plan actually produces. Realised counts, not apportioned ones: they are computed by
    /// describing every item, so what the report says is what the corpus will contain.
    /// </summary>
    public CorpusPlanReport Report(int from, int to)
    {
        if (from < 1 || to < from)
        {
            throw new ArgumentOutOfRangeException(nameof(to), "Empty ordinal range.");
        }

        var byFolder = new SortedDictionary<int, int>();
        var bySizeClass = new SortedDictionary<string, int>(StringComparer.Ordinal);
        var byDateBand = new SortedDictionary<string, int>(StringComparer.Ordinal);
        var withinDays = new SortedDictionary<int, int>();
        long totalBodyBytes = 0;
        long largeBodies = 0;
        long hugeBodies = 0;
        long bodyCapTrippers = 0;
        DateTime oldest = DateTime.MaxValue;
        DateTime newest = DateTime.MinValue;

        foreach (int mark in WindowDayMarks)
        {
            withinDays[mark] = 0;
        }

        for (int ordinal = from; ordinal <= to; ordinal++)
        {
            CorpusItemSpec spec = Describe(ordinal);
            byFolder[spec.FolderId] = byFolder.TryGetValue(spec.FolderId, out int f) ? f + 1 : 1;
            bySizeClass[spec.SizeClass] = bySizeClass.TryGetValue(spec.SizeClass, out int s) ? s + 1 : 1;
            byDateBand[spec.DateBand] = byDateBand.TryGetValue(spec.DateBand, out int d) ? d + 1 : 1;
            totalBodyBytes += spec.BodyBytes;
            if (spec.BodyBytes >= LargeBodyBytes)
            {
                largeBodies++;
            }

            if (spec.BodyBytes >= HugeBodyBytes)
            {
                hugeBodies++;
            }

            if (spec.BodyBytes > OutlookAI.Core.Com.OutlookComSession.SweepBodyCharsCap)
            {
                bodyCapTrippers++;
            }

            if (spec.ReceivedUtc < oldest)
            {
                oldest = spec.ReceivedUtc;
            }

            if (spec.ReceivedUtc > newest)
            {
                newest = spec.ReceivedUtc;
            }

            foreach (int mark in WindowDayMarks)
            {
                if (spec.ReceivedUtc > _options.AnchorUtc.AddDays(-mark))
                {
                    withinDays[mark]++;
                }
            }
        }

        return new CorpusPlanReport(
            from, to, byFolder, bySizeClass, byDateBand, withinDays,
            totalBodyBytes, largeBodies, hugeBodies, bodyCapTrippers, oldest, newest);
    }

    /// <summary>Outlook default-folder id for Sent Items - its dates behave differently.</summary>
    internal const int SentItemsFolderId = 5;

    /// <summary>Outlook default-folder id for Deleted Items.</summary>
    internal const int DeletedItemsFolderId = 3;

    /// <summary>Body length at which an item counts as a large quoted thread in the report.</summary>
    internal const int LargeBodyBytes = 24_000;

    /// <summary>Body length at which an item counts as a frame-budget monster in the report.</summary>
    internal const int HugeBodyBytes = 96_000;

    /// <summary>
    /// Day marks the report always counts, chosen to match the windows under measurement:
    /// the sweep's fallback window (7), the scan windows (30/60/90) and a year.
    /// </summary>
    internal static readonly int[] WindowDayMarks = { 1, 7, 30, 60, 90, 365 };

    // Independent hash streams. Each field draws from its own so changing one field's
    // derivation later cannot silently reshuffle the others.
    private const int StreamFolder = 1;
    private const int StreamSizeClass = 2;
    private const int StreamDateBand = 3;
    private const int StreamBodyBytes = 4;
    private const int StreamAgeSeconds = 5;
    private const int StreamTransport = 6;
    private const int StreamRead = 7;
    private const int StreamSubject = 1_000;
    private const int StreamBody = 100_000;

    private ulong Draw(int ordinal, int stream) => Draw(_options.Seed, ordinal, stream);

    /// <summary>
    /// SplitMix64 over (seed, ordinal, stream). Chosen over <c>System.Random</c> on
    /// purpose: <c>Random</c>'s sequence for a given seed is explicitly not a contract
    /// across framework versions, so a corpus rebuilt on a newer runtime could differ
    /// from the one a measurement was taken against. It is also stateless, which is what
    /// lets an interrupted build resume without replaying what it already produced.
    /// <para>
    /// Public so the T1 tier can pin the sequence itself. Without that, swapping this for
    /// <c>Random</c> would break the reproducibility of every corpus already measured
    /// against and no test would notice, because every other property here is a
    /// distribution and a different generator satisfies the same distributions.
    /// </para>
    /// </summary>
    public static ulong Draw(long seed, int ordinal, int stream)
    {
        unchecked
        {
            ulong x = (ulong)seed;
            x = Mix(x + (0x9E3779B97F4A7C15UL * (ulong)(uint)ordinal));
            x = Mix(x + (0xBF58476D1CE4E5B9UL * (ulong)(uint)stream));
            return Mix(x);
        }
    }

    private static ulong Mix(ulong z)
    {
        unchecked
        {
            z += 0x9E3779B97F4A7C15UL;
            z = (z ^ (z >> 30)) * 0xBF58476D1CE4E5B9UL;
            z = (z ^ (z >> 27)) * 0x94D049BB133111EBUL;
            return z ^ (z >> 31);
        }
    }

    private static T PickWeighted<T>(IReadOnlyList<T> items, Func<T, int> weight, int total, ulong draw)
    {
        int target = (int)(draw % (ulong)total);
        int running = 0;
        foreach (T item in items)
        {
            running += weight(item);
            if (target < running)
            {
                return item;
            }
        }

        return items[items.Count - 1]; // unreachable while the weights sum to total
    }

    private static int SumWeights(IEnumerable<int> weights, string what)
    {
        int total = 0;
        foreach (int w in weights)
        {
            if (w < 0)
            {
                throw new ArgumentException($"Negative weight in {what}.");
            }

            total += w;
        }

        if (total <= 0)
        {
            throw new ArgumentException($"The {what} must carry at least one positive weight.");
        }

        return total;
    }
}

/// <summary>
/// What a plan produces over an ordinal range, computed by describing every item in it.
/// This is the sheet a measurement is read against: it says how many items a 7-day and a
/// 60-day window should select before anything is built.
/// </summary>
/// <param name="FromOrdinal">First ordinal covered.</param>
/// <param name="ToOrdinal">Last ordinal covered.</param>
/// <param name="ByFolderId">Item count per Outlook default-folder id.</param>
/// <param name="BySizeClass">Item count per size class.</param>
/// <param name="ByDateBand">Item count per age band.</param>
/// <param name="WithinDays">Item count newer than N days before the anchor, for the marks a measurement uses.</param>
/// <param name="TotalBodyBytes">Sum of all body lengths - the lower bound on what the PST will hold.</param>
/// <param name="BodiesAtLeast24Kb">Items whose body is 24 KB or larger (the quoted-thread population).</param>
/// <param name="BodiesAtLeast96Kb">Items whose body is 96 KB or larger (the frame-budget population).</param>
/// <param name="BodiesOverSweepBodyCap">
/// Items whose body exceeds <see cref="OutlookAI.Core.Com.OutlookComSession.SweepBodyCharsCap"/>,
/// so a sweep that reaches them must cut them and say so. The expected value of the
/// sweep's <c>itemsBodyCapped</c> counter, if the sweep window selects the whole corpus.
/// </param>
/// <param name="OldestReceivedUtc">Oldest intended received instant.</param>
/// <param name="NewestReceivedUtc">Newest intended received instant.</param>
public sealed record CorpusPlanReport(
    int FromOrdinal,
    int ToOrdinal,
    IReadOnlyDictionary<int, int> ByFolderId,
    IReadOnlyDictionary<string, int> BySizeClass,
    IReadOnlyDictionary<string, int> ByDateBand,
    IReadOnlyDictionary<int, int> WithinDays,
    long TotalBodyBytes,
    long BodiesAtLeast24Kb,
    long BodiesAtLeast96Kb,
    long BodiesOverSweepBodyCap,
    DateTime OldestReceivedUtc,
    DateTime NewestReceivedUtc)
{
    /// <summary>Number of items covered.</summary>
    public int ItemCount => ToOrdinal - FromOrdinal + 1;

    /// <summary>Mean body length in bytes.</summary>
    public long MeanBodyBytes => ItemCount == 0 ? 0 : TotalBodyBytes / ItemCount;
}
