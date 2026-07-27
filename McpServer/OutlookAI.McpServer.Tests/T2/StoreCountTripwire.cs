using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>Verdict of one before/after comparison. Failures fail the suite; notes are informational.</summary>
public sealed class TripwireVerdict
{
    internal TripwireVerdict(IReadOnlyList<string> failures, IReadOnlyList<string> notes)
    {
        Failures = failures;
        Notes = notes;
    }

    /// <summary>Losses: an item-count DECREASE or a folder appearing/disappearing outside the hub.</summary>
    public IReadOnlyList<string> Failures { get; }

    /// <summary>Benign observations (mail arriving elsewhere during the run, hub churn).</summary>
    public IReadOnlyList<string> Notes { get; }

    /// <summary>True when anything must fail the suite.</summary>
    public bool Failed => Failures.Count > 0;

    /// <summary>One multi-line message naming every store/folder/delta.</summary>
    public string Describe()
    {
        return "STORE COUNT TRIPWIRE: the live tier changed mailboxes it may not touch."
            + Environment.NewLine + string.Join(Environment.NewLine, Failures);
    }
}

/// <summary>
/// Pure comparison behind the per-store count tripwire: a live run may add nothing and
/// remove nothing outside the designated test mailbox.
/// <para>
/// Tuned for low false positives, because an 8-minute run happens while real mail arrives:
/// <list type="bullet">
/// <item>item-count INCREASES outside the hub are normal (mail arriving, a rule filing it)
/// - noted, never failed;</item>
/// <item>item-count DECREASES outside the hub FAIL - nothing the suite does may remove an
/// item from another mailbox;</item>
/// <item>a folder ADDED or REMOVED outside the hub FAILS - the suite creates folders only
/// in the hub, and removes only its own;</item>
/// <item>the hub is exempt here: its churn is tagged, and the zero-tagged-artifact sweep
/// plus the move/archive reconciliation already police it.</item>
/// </list>
/// </para>
/// </summary>
public static class StoreCountTripwire
{
    /// <summary>Compares two per-store, per-folder censuses.</summary>
    public static TripwireVerdict Evaluate(
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, int>> before,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, int>> after,
        string hubStoreDisplayName)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);

        List<string> failures = new();
        List<string> notes = new();

        foreach (KeyValuePair<string, IReadOnlyDictionary<string, int>> store in before)
        {
            bool isHub = string.Equals(store.Key, hubStoreDisplayName, StringComparison.OrdinalIgnoreCase);
            if (!after.TryGetValue(store.Key, out IReadOnlyDictionary<string, int>? now))
            {
                // A store present at the start and gone at the end is never benign.
                failures.Add("  store '" + store.Key + "' could not be re-counted after the run (store missing).");
                continue;
            }

            foreach (KeyValuePair<string, int> folder in store.Value)
            {
                if (!now.TryGetValue(folder.Key, out int afterCount))
                {
                    if (isHub)
                    {
                        notes.Add("  hub folder removed: '" + folder.Key + "' (test folders are cleaned up).");
                    }
                    else
                    {
                        failures.Add("  FOLDER REMOVED: store '" + store.Key + "' folder '" + folder.Key
                            + "' existed before the run and is gone (" + Count(folder.Value) + " before).");
                    }

                    continue;
                }

                int delta = afterCount - folder.Value;
                if (delta == 0)
                {
                    continue;
                }

                if (delta < 0 && !isHub)
                {
                    failures.Add("  ITEMS LOST: store '" + store.Key + "' folder '" + folder.Key + "' "
                        + Count(folder.Value) + " -> " + Count(afterCount) + " (" + Count(delta) + ").");
                }
                else if (!isHub)
                {
                    notes.Add("  arrivals: store '" + store.Key + "' folder '" + folder.Key + "' "
                        + Count(folder.Value) + " -> " + Count(afterCount) + " (+" + Count(delta) + ").");
                }
            }

            foreach (KeyValuePair<string, int> folder in now)
            {
                if (store.Value.ContainsKey(folder.Key))
                {
                    continue;
                }

                if (isHub)
                {
                    notes.Add("  hub folder added: '" + folder.Key + "'.");
                }
                else
                {
                    failures.Add("  FOLDER ADDED: store '" + store.Key + "' folder '" + folder.Key
                        + "' did not exist before the run (" + Count(folder.Value) + " items).");
                }
            }
        }

        foreach (string storeName in after.Keys)
        {
            if (!before.ContainsKey(storeName))
            {
                failures.Add("  store '" + storeName + "' appeared during the run and was never baselined.");
            }
        }

        return new TripwireVerdict(failures, notes);
    }

    private static string Count(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}
