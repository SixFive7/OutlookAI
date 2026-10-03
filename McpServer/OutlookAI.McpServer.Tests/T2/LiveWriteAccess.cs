namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// THE ONE PLACE that decides whether a live run may write anything at all - decided by the
/// maintainer 2026-10-03 (Q74, "layers 1+2+3"), on top of his standing rule of 2026-09-24
/// (Q72): <b>the maintainer's workstation is read-only for live tests, always.</b>
/// <para>
/// <b>The semantics, and why they are the fail-safe ones.</b> Only a machine that DECLARES itself
/// <see cref="LiveMachineProfile.Portable"/> - a dedicated test guest, whose settings
/// <c>Testbed/host/New-LiveTestSettings.ps1</c> renders with that value and refuses to render
/// without it - may write. Everything else refuses every write, the designated test mailbox
/// included:
/// <list type="bullet">
/// <item><see cref="LiveMachineProfile.Production"/>. Since Q72 the only Production machine there
/// is, is the maintainer's workstation, and nothing may be written there. So "Production" no longer
/// means "a real mailbox the suite may write its hub in"; it means "the read-only workstation".</item>
/// <item><b>No profile declared at all</b>, which the loader reads as Production - the shape of the
/// workstation's own settings file, written before the field existed. The two answers this
/// decision could have had were "Production is read-only" and "no declaration is read-only"; the
/// first is chosen because it contains the second and closes the gap the second leaves open: under
/// "no declaration is read-only", typing <c>"machineProfile": "Production"</c> into a gitignored
/// file would have bought write access on the one machine that must never have it.</item>
/// <item><b>Any value added to the enum later</b>, or a number no member carries
/// (<c>"machineProfile": 7</c> parses). Nobody has thought about such a machine yet, and the
/// direction that fails safe is the one that writes nothing.</item>
/// </list>
/// </para>
/// <para>
/// <b>What reads it.</b> Three gates, and each asks this rather than the enum, so the three cannot
/// come to disagree about which machine is read-only:
/// <list type="number">
/// <item>the in-process write guard - <see cref="LiveStoreWriteGuard.Build"/> builds a
/// <see cref="StoreWriteAllowlist"/> that refuses EVERY store (layer 2);</item>
/// <item>the out-of-process gate - <c>T3.McpStdioClient</c> refuses every tool that is not
/// classified read-only before it sends it, so a write through the MCP server process, which the
/// in-process guard cannot see, is refused too (layer 3, <see cref="StdioPostureFor"/>);</item>
/// <item>the count tripwire's retry bounds - a read-only profile keeps its re-censuses and never
/// starts the bounded re-run (<see cref="TripwireRetryPolicy.For(LiveMachineProfile)"/>, A1).</item>
/// </list>
/// What a machine profile means for the POPULATIONS a test may expect is a different question, and
/// it is still answered in exactly one place: <see cref="LiveTestSettings.RequireProductionPopulation"/>,
/// through <see cref="LivePopulationCoverage"/>. Nothing here changes it.
/// </para>
/// </summary>
public static class LiveWriteAccess
{
    /// <summary>
    /// The phrase every refusal this decision causes carries, so a run log can be searched for it
    /// and the T1 pins can find it. One spelling, here.
    /// </summary>
    public const string ReadOnlyMachine = "READ-ONLY MACHINE";

    /// <summary>
    /// True when <paramref name="profile"/> refuses every write, the test mailbox included. Only
    /// <see cref="LiveMachineProfile.Portable"/> may write.
    /// </summary>
    public static bool RefusesEveryWrite(LiveMachineProfile profile)
    {
        return profile != LiveMachineProfile.Portable;
    }

    /// <summary>
    /// Why a write was refused on a read-only machine: the rule, where it is written down, and what
    /// to do instead. Content-free - it names a profile and a store at most, never mail.
    /// </summary>
    public static string ReadOnlyReason(LiveMachineProfile profile)
    {
        string declared = Enum.IsDefined(profile) ? "'" + profile + "'" : "the undefined value " + (int)profile;
        return ReadOnlyMachine + ": these live-test settings declare machineProfile " + declared
            + ", and only a machine that declares 'Portable' - a test guest - may write anything. "
            + "'Production' is the maintainer's workstation, which is read-only for live tests ALWAYS, the "
            + "designated test mailbox included (AGENTS.md, Mailbox Safety; Q72, enforced in code since Q74). "
            + "A test that writes runs on a test guest; on this machine only tests carrying Writes=Nothing are selected, "
            + "through the workstation filter in Testbed/README.md section 4d.";
    }

    /// <summary>
    /// The posture the test-side MCP client takes for this process - see
    /// <see cref="StdioPostureFor"/>. Decided once: the opt-in and the settings file cannot change
    /// during a run, and deciding it twice is a way for two clients in one run to disagree.
    /// </summary>
    public static StdioWritePosture CurrentStdioPosture => CurrentPosture.Value;

    private static readonly Lazy<StdioWritePosture> CurrentPosture = new(
        () => StdioPostureFor(LiveRunOptIn.CurrentVerdict(), LiveTestSettings.Load),
        LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>
    /// What the test-side MCP client may send in this process. Pure, so T1 pins every branch.
    /// <list type="bullet">
    /// <item><b>Not opted in</b> - no live run is happening: every live fixture refuses at its first
    /// line, so the only tests that can reach the client are the CI-tier ones, whose write-tool calls
    /// are built to be refused by the server before any COM work (and pinned so). Their rules are
    /// unchanged; <paramref name="loadSettings"/> is not even called, so a CI run never reads a
    /// settings file.</item>
    /// <item><b>Opted in</b> - a live run: the machine's own settings decide, through
    /// <see cref="RefusesEveryWrite"/>.</item>
    /// <item><b>Opted in, and the settings cannot be read</b> - refused as read-only. A posture
    /// nobody could establish must not be the one that writes.</item>
    /// </list>
    /// </summary>
    /// <param name="optIn">What <see cref="LiveRunOptIn"/> decided for this process.</param>
    /// <param name="loadSettings">Reads this machine's live-test settings; only called when opted in.</param>
    public static StdioWritePosture StdioPostureFor(LiveRunOptIn.Verdict optIn, Func<LiveTestSettings> loadSettings)
    {
        ArgumentNullException.ThrowIfNull(loadSettings);
        if (optIn != LiveRunOptIn.Verdict.Open)
        {
            return StdioWritePosture.NotALiveRun;
        }

        try
        {
            return RefusesEveryWrite(loadSettings().MachineProfile)
                ? StdioWritePosture.ReadOnly
                : StdioWritePosture.Writable;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return StdioWritePosture.ReadOnly;
        }
    }
}

/// <summary>
/// What the test-side MCP client (<c>T3.McpStdioClient</c>) may send in one process. Decided by
/// <see cref="LiveWriteAccess.StdioPostureFor"/>.
/// </summary>
public enum StdioWritePosture
{
    /// <summary>
    /// Every tool not classified read-only is refused before it is sent - an unclassified one
    /// included. Deliberately the ZERO value: a posture nobody set is the one that writes nothing.
    /// </summary>
    ReadOnly = 0,

    /// <summary>
    /// No live run: the run did not opt in, so no live-test settings are in play and the CI tier's
    /// rules apply unchanged.
    /// </summary>
    NotALiveRun = 1,

    /// <summary>A live run on a machine that declares <see cref="LiveMachineProfile.Portable"/>.</summary>
    Writable = 2,
}
