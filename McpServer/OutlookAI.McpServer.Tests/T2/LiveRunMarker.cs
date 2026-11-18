using System.Globalization;
using System.Text.RegularExpressions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The RUN MARKER: one string every item a live run creates carries in its subject beside the tag, so that
/// the run's items can be told from every other run's - and from anything else - by an ordinal match on
/// two strings, never by a pattern (the Phase 2 write-safety design, <c>Docs/live-tier-on-the-vm.md</c>
/// section 4.4; built 2026-10-08 for the Exchange VM's test account).
/// <para>
/// <b>Where it comes from.</b> The host's runner (<c>Testbed/host/Invoke-LiveTierOnGuest.ps1</c>) makes it
/// before the run starts, RECORDS it in its run ledger on the host, and hands it to the run in
/// <see cref="EnvironmentVariable"/> - with the markers of earlier runs that never reported a clean sweep in
/// <see cref="PriorEnvironmentVariable"/>. A checkpoint restore wipes the guest's disk but not the Exchange
/// server, so the record that survives an aborted run has to live on the host. Without the variable - a
/// PST guest, a run by hand - the process makes its own, which nothing records; the Exchange VM's test
/// account refuses such a run (<see cref="RequireRecordedOnExchange"/>).
/// </para>
/// <para>
/// <b>Where it goes.</b> Every fixture's own marker starts with a short prefix and ends with this one
/// (<see cref="ForFixture"/>), so a test's per-test cleanup - tag AND its fixture's marker - still selects
/// only its own items, and the run's final sweep - tag AND this marker - selects every item of the run and
/// nothing older.
/// </para>
/// </summary>
public static class LiveRunMarker
{
    /// <summary>The variable the runner sets, inside the guest task's script, for one run.</summary>
    public const string EnvironmentVariable = "OUTLOOKAI_LIVE_RUN_MARKER";

    /// <summary>The variable carrying earlier runs' recorded markers, comma-separated.</summary>
    public const string PriorEnvironmentVariable = "OUTLOOKAI_LIVE_PRIOR_RUN_MARKERS";

    /// <summary>
    /// What a marker looks like: lower-case letters and digits, 12 to 40 of them. Long enough that the
    /// mailer's delete filter accepts it on its own (it refuses anything under 12), plain enough to sit in
    /// a DASL LIKE literal, and nothing a person types into a subject by accident.
    /// </summary>
    private static readonly Regex WellFormed = new("^[a-z0-9]{12,40}$", RegexOptions.CultureInvariant);

    private static readonly Lazy<(string Marker, bool Recorded)> CurrentValue = new(
        () => Decide(Environment.GetEnvironmentVariable(EnvironmentVariable), DateTime.UtcNow, Guid.NewGuid()),
        LazyThreadSafetyMode.ExecutionAndPublication);

    private static readonly Lazy<IReadOnlyList<string>> PriorValue = new(
        () => ParsePrior(Environment.GetEnvironmentVariable(PriorEnvironmentVariable), Current),
        LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>This run's marker.</summary>
    public static string Current => CurrentValue.Value.Marker;

    /// <summary>True when the marker came from the runner, which recorded it on the host.</summary>
    public static bool Recorded => CurrentValue.Value.Recorded;

    /// <summary>Earlier runs' recorded markers, well-formed and distinct, this run's excluded.</summary>
    public static IReadOnlyList<string> Prior => PriorValue.Value;

    /// <summary>A fixture's own marker: its prefix, then this run's marker.</summary>
    public static string ForFixture(string prefix)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(prefix);
        return prefix + Current;
    }

    /// <summary>True when <paramref name="marker"/> has a marker's shape.</summary>
    public static bool IsWellFormed(string? marker)
    {
        return marker != null && WellFormed.IsMatch(marker);
    }

    /// <summary>
    /// The marker for this process: the runner's, when it set a well-formed one; otherwise one made here -
    /// <c>oaix</c>, the UTC time to the second and eight hex digits - which nothing recorded. Pure.
    /// </summary>
    internal static (string Marker, bool Recorded) Decide(string? fromRunner, DateTime utcNow, Guid random)
    {
        string? given = fromRunner?.Trim();
        if (IsWellFormed(given))
        {
            return (given!, true);
        }

        string made = "oaix" + utcNow.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)
            + random.ToString("N", CultureInfo.InvariantCulture).Substring(0, 8);
        return (made, false);
    }

    /// <summary>
    /// The prior markers in <paramref name="value"/>: separated by ',' or ';', trimmed, kept only when
    /// well-formed, without repeats and without <paramref name="current"/>. Pure.
    /// </summary>
    internal static IReadOnlyList<string> ParsePrior(string? value, string current)
    {
        List<string> markers = new();
        if (string.IsNullOrWhiteSpace(value))
        {
            return markers;
        }

        foreach (string part in value.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string marker = part.Trim();
            if (IsWellFormed(marker) && !string.Equals(marker, current, StringComparison.Ordinal) && !markers.Contains(marker))
            {
                markers.Add(marker);
            }
        }

        return markers;
    }

    /// <summary>
    /// Refuses a live run on the Exchange VM's test account whose marker nobody recorded: its items would
    /// carry a marker no later run could sweep by, if it were aborted. Every other machine returns.
    /// </summary>
    public static void RequireRecordedOnExchange(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        string? refusal = RefusalFor(settings.MachineProfile, Recorded);
        if (refusal != null)
        {
            throw new InvalidOperationException(refusal);
        }
    }

    /// <summary>The refusal <see cref="RequireRecordedOnExchange"/> throws, or null. Pure.</summary>
    internal static string? RefusalFor(LiveMachineProfile profile, bool recorded)
    {
        if (profile != LiveMachineProfile.ExchangeTestAccount || recorded)
        {
            return null;
        }

        return "LIVE TEST REFUSED: this is the Exchange VM's test account (machineProfile 'ExchangeTestAccount'), whose "
            + "runs write in a real Exchange mailbox that no checkpoint restores, and this run carries no recorded run "
            + "marker (" + EnvironmentVariable + " is not set). Every item a run makes there carries the marker the host "
            + "recorded, so that an aborted run's leftovers can be swept by it later. Run it through "
            + "Testbed/host/Invoke-LiveTierOnGuest.ps1 -VMName OutlookAI-Exchange, which records one.";
    }
}
