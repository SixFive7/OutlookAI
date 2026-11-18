namespace OutlookAI.RemediationTools;

/// <summary>
/// How the corpus tool picks the way it writes, once it has tried each: the first rung of a ladder, in the
/// ladder's own order, that a probe found usable. The date writer (<see cref="CorpusDateFidelity"/>) and the
/// item placement (<see cref="CorpusPlacement"/>) both climb one; each kept a copy until the code-sharing
/// refactor's stage S6.
/// </summary>
internal static class CorpusProbeLadder
{
    /// <summary>
    /// The first method of <paramref name="ladder"/> that one of <paramref name="probes"/> tried and found
    /// usable; <paramref name="none"/> when none was.
    /// </summary>
    public static TMethod FirstVerified<TMethod, TProbe>(
        IEnumerable<TMethod> ladder,
        IReadOnlyCollection<TProbe> probes,
        Func<TProbe, TMethod> methodOf,
        Func<TProbe, bool> isUsable,
        TMethod none)
        where TMethod : struct, Enum
    {
        ArgumentNullException.ThrowIfNull(ladder);
        ArgumentNullException.ThrowIfNull(probes);
        foreach (TMethod method in ladder)
        {
            foreach (TProbe probe in probes)
            {
                if (EqualityComparer<TMethod>.Default.Equals(methodOf(probe), method) && isUsable(probe))
                {
                    return method;
                }
            }
        }

        return none;
    }
}
