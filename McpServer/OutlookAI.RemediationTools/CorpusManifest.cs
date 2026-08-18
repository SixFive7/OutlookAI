using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookAI.RemediationTools;

/// <summary>
/// First line of a manifest: what corpus this file records, and where. Everything here
/// except <see cref="StoreFilePath"/> and <see cref="DateWriteMethod"/> participates in
/// the compatibility check that guards a resumed or extended run.
/// </summary>
/// <param name="Version">Manifest format version.</param>
/// <param name="CorpusId">The corpus id embedded in every subject.</param>
/// <param name="Seed">The generator seed.</param>
/// <param name="AnchorUtc">The anchor, round-tripped as an ISO-8601 UTC string.</param>
/// <param name="ShapeKey">Digest of the whole shape - see <see cref="CorpusPlanOptions.ShapeKey"/>.</param>
/// <param name="StoreDisplayName">The store the items were written into.</param>
/// <param name="StoreFilePath">The .pst backing that store, recorded so a manifest can be matched to a file.</param>
/// <param name="DateWriteMethod">Which date-write rung the build verified before it started.</param>
/// <param name="PlacementMethod">
/// Which placement rung the build verified. Recorded because it decides whether the corpus
/// is measurable at all: a corpus placed as drafts is invisible to the freshness sweep, and
/// a measurement taken against one would otherwise look like a measurement of an empty
/// store. Null in manifests written before placement was probed.
/// </param>
public sealed record CorpusManifestHeader(
    int Version,
    string CorpusId,
    long Seed,
    string AnchorUtc,
    string ShapeKey,
    string StoreDisplayName,
    string? StoreFilePath,
    string DateWriteMethod,
    string? PlacementMethod = null);

/// <summary>One item this build created. The EntryID half of the teardown's two-key rule.</summary>
/// <param name="Ordinal">The item's ordinal in the plan.</param>
/// <param name="EntryId">EntryID as returned by Outlook after Save().</param>
/// <param name="FolderId">Outlook default-folder id it was created in.</param>
/// <param name="BodyBytes">Body length actually written.</param>
/// <param name="ReceivedUtc">The received instant Outlook reported after the write, or null when unverified.</param>
public sealed record CorpusManifestItem(
    int Ordinal,
    string EntryId,
    int FolderId,
    int BodyBytes,
    string? ReceivedUtc);

/// <summary>A folder this build created because the store had no such default folder.</summary>
/// <param name="EntryId">Folder EntryID.</param>
/// <param name="Name">Folder name; always carries <see cref="CorpusManifest.CreatedFolderPrefix"/>.</param>
/// <param name="FolderId">The default-folder id it stands in for.</param>
public sealed record CorpusManifestFolder(string EntryId, string Name, int FolderId);

/// <summary>Why a manifest may not be resumed against the parameters in hand.</summary>
public enum CorpusManifestMismatch
{
    /// <summary>Compatible - the run may continue this corpus.</summary>
    None = 0,

    /// <summary>The manifest was written by a different format version.</summary>
    Version = 1,

    /// <summary>A different corpus id.</summary>
    CorpusId = 2,

    /// <summary>A different shape: seed, anchor, size classes, date bands or folder shares.</summary>
    Shape = 3,

    /// <summary>A different target store.</summary>
    Store = 4,
}

/// <summary>
/// The append-only record of what a corpus build actually created: one JSON object per
/// line, header first. It is three things at once, and all three matter.
/// <list type="number">
/// <item><b>The teardown's allowlist.</b> Nothing is ever deleted that is not recorded here.</item>
/// <item><b>The resume point.</b> A build reads it, sees which ordinals exist, and creates only the rest - so an interrupted build is restarted rather than repeated, and a completed one re-run is a no-op.</item>
/// <item><b>The proof of shape.</b> The header carries the plan digest, so a run cannot append items from a different seed or anchor to an existing corpus and leave a file that describes neither.</item>
/// </list>
/// <para>
/// JSON Lines rather than one JSON document, deliberately: a build of tens of thousands
/// of items will be interrupted, and a half-written array is unreadable while a
/// half-written last line costs exactly one item. Parsing and rendering are pure and take
/// strings, so the T1 tier pins the format without touching a disk.
/// </para>
/// <para>
/// <b>If the manifest is lost, teardown cannot run.</b> That is the intended failure, not
/// an oversight - the alternative is deleting by subject match alone, which is the thing
/// the mailbox-safety rules forbid outright. <c>corpus-reindex</c> exists for that case:
/// it READS the store and writes a fresh candidate manifest, which an operator inspects
/// before handing it to teardown.
/// </para>
/// </summary>
public sealed class CorpusManifest
{
    /// <summary>Current manifest format version.</summary>
    public const int CurrentVersion = 1;

    /// <summary>Name prefix of any folder the builder creates. Ordinal-matched at teardown.</summary>
    public const string CreatedFolderPrefix = "OutlookAI-Corpus-Folder";

    private static readonly JsonSerializerOptions Json = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly Dictionary<int, CorpusManifestItem> _items = new();
    private readonly List<CorpusManifestFolder> _folders = new();

    private CorpusManifest(CorpusManifestHeader header)
    {
        Header = header;
    }

    /// <summary>What this manifest says the corpus is.</summary>
    public CorpusManifestHeader Header { get; }

    /// <summary>Items recorded, keyed by ordinal.</summary>
    public IReadOnlyDictionary<int, CorpusManifestItem> Items => _items;

    /// <summary>Folders the builder created.</summary>
    public IReadOnlyList<CorpusManifestFolder> Folders => _folders;

    /// <summary>Lines that could not be parsed. A non-empty list is reported, never ignored.</summary>
    public IReadOnlyList<string> UnparseableLines { get; private set; } = Array.Empty<string>();

    /// <summary>Every recorded EntryID - the allowlist half of <see cref="CorpusSafety.MayDelete"/>.</summary>
    public IReadOnlyCollection<string> EntryIds => _items.Values.Select(i => i.EntryId).ToList();

    /// <summary>Starts a manifest in memory for a fresh corpus.</summary>
    public static CorpusManifest Create(CorpusManifestHeader header)
    {
        ArgumentNullException.ThrowIfNull(header);
        return new CorpusManifest(header);
    }

    /// <summary>
    /// Parses a manifest from its lines. A trailing partial line - the shape an interrupted
    /// build leaves - is reported in <see cref="UnparseableLines"/> rather than throwing,
    /// because losing one item's record must not make the other 40 000 unreadable.
    /// </summary>
    public static CorpusManifest Parse(IEnumerable<string> lines)
    {
        ArgumentNullException.ThrowIfNull(lines);
        CorpusManifest? manifest = null;
        var bad = new List<string>();
        foreach (string raw in lines)
        {
            string line = raw.Trim();
            if (line.Length == 0)
            {
                continue;
            }

            if (manifest == null)
            {
                CorpusManifestHeader header = JsonSerializer.Deserialize<CorpusManifestHeader>(line, Json)
                    ?? throw new InvalidOperationException("Manifest header line is not a manifest header.");
                if (string.IsNullOrWhiteSpace(header.CorpusId))
                {
                    throw new InvalidOperationException("Manifest header carries no corpus id.");
                }

                manifest = new CorpusManifest(header);
                continue;
            }

            try
            {
                using JsonDocument doc = JsonDocument.Parse(line);
                if (doc.RootElement.TryGetProperty("Name", out _))
                {
                    CorpusManifestFolder folder = JsonSerializer.Deserialize<CorpusManifestFolder>(line, Json)!;
                    manifest._folders.Add(folder);
                }
                else
                {
                    CorpusManifestItem item = JsonSerializer.Deserialize<CorpusManifestItem>(line, Json)!;
                    if (item.Ordinal < 1 || string.IsNullOrWhiteSpace(item.EntryId))
                    {
                        bad.Add(line);
                        continue;
                    }

                    manifest._items[item.Ordinal] = item;
                }
            }
            catch (JsonException)
            {
                bad.Add(line);
            }
        }

        if (manifest == null)
        {
            throw new InvalidOperationException("Manifest is empty - it has no header line.");
        }

        manifest.UnparseableLines = bad;
        return manifest;
    }

    /// <summary>Records an item in memory. The caller appends the same line to the file.</summary>
    public void Add(CorpusManifestItem item)
    {
        ArgumentNullException.ThrowIfNull(item);
        _items[item.Ordinal] = item;
    }

    /// <summary>Records a created folder in memory.</summary>
    public void Add(CorpusManifestFolder folder)
    {
        ArgumentNullException.ThrowIfNull(folder);
        _folders.Add(folder);
    }

    /// <summary>Renders the header line.</summary>
    public static string RenderLine(CorpusManifestHeader header) => JsonSerializer.Serialize(header, Json);

    /// <summary>Renders one item line.</summary>
    public static string RenderLine(CorpusManifestItem item) => JsonSerializer.Serialize(item, Json);

    /// <summary>Renders one folder line.</summary>
    public static string RenderLine(CorpusManifestFolder folder) => JsonSerializer.Serialize(folder, Json);

    /// <summary>
    /// Whether a run described by <paramref name="options"/> against
    /// <paramref name="storeDisplayName"/> may append to this manifest. The requested item
    /// COUNT is deliberately not part of the comparison: item N's description does not
    /// depend on it, so "build 20 000 more" is a legal continuation of the same corpus.
    /// </summary>
    public CorpusManifestMismatch CheckCompatible(CorpusPlanOptions options, string storeDisplayName)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (Header.Version != CurrentVersion)
        {
            return CorpusManifestMismatch.Version;
        }

        if (!string.Equals(Header.CorpusId, options.CorpusId, StringComparison.Ordinal))
        {
            return CorpusManifestMismatch.CorpusId;
        }

        if (!string.Equals(Header.ShapeKey, options.ShapeKey, StringComparison.Ordinal))
        {
            return CorpusManifestMismatch.Shape;
        }

        if (!string.Equals(Header.StoreDisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))
        {
            return CorpusManifestMismatch.Store;
        }

        return CorpusManifestMismatch.None;
    }

    /// <summary>The refusal message for a mismatch.</summary>
    public static string Explain(CorpusManifestMismatch mismatch) => mismatch switch
    {
        CorpusManifestMismatch.None => "manifest is compatible",
        CorpusManifestMismatch.Version => "the manifest was written by a different format version",
        CorpusManifestMismatch.CorpusId => "the manifest records a different corpus id",
        CorpusManifestMismatch.Shape => "the manifest records a different shape (seed, anchor, size classes, "
            + "date bands or folder shares) - appending would produce a corpus that neither description fits",
        CorpusManifestMismatch.Store => "the manifest records a different target store",
        _ => "unrecognised mismatch",
    };

    /// <summary>
    /// The ordinals in 1..<paramref name="itemCount"/> that are NOT yet recorded, in order.
    /// This is the whole of resumption: it is derived from the manifest rather than from a
    /// saved cursor, so an interrupted build restarts correctly even if it died between
    /// creating an item and recording it (that item is simply created again - the plan is
    /// deterministic, so the duplicate is an identical item, and it too is recorded).
    /// </summary>
    public IEnumerable<int> MissingOrdinals(int itemCount)
    {
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            if (!_items.ContainsKey(ordinal))
            {
                yield return ordinal;
            }
        }
    }

    /// <summary>Formats an instant the way manifest lines carry it.</summary>
    public static string FormatUtc(DateTime utc)
        => DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

    /// <summary>Parses an instant a manifest line carries. Returns null when unreadable.</summary>
    public static DateTime? ParseUtc(string? text)
        => DateTime.TryParse(
            text,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
            out DateTime value)
            ? DateTime.SpecifyKind(value, DateTimeKind.Utc)
            : null;
}
