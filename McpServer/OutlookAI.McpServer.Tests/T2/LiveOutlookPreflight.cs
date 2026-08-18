using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What the preflight decides to do about one observed Outlook state.</summary>
public enum LivePreflightVerdict
{
    /// <summary>Outlook is usable (or absent, which the tier is allowed to fix by starting it).</summary>
    Proceed = 0,

    /// <summary>Outlook is mid-start. Wait for it to settle and look again.</summary>
    Settle = 1,

    /// <summary>Outlook is wedged. Refuse the tier before it touches COM.</summary>
    Refuse = 2,
}

/// <summary>
/// The live tier's health gate: asks Windows whether Outlook is actually responding
/// BEFORE any fixture opens a COM session, and refuses the whole tier when it is not.
/// <para>
/// This exists because of the 2026-08-18 incident. Outlook was completely wedged - it
/// could not be started or killed and had to be ended from Task Manager - and the live
/// tier discovered that by hanging: one test ran 22.5 minutes, then fixture setup sat for
/// 10 and then 15 minutes with no output at all, never even spawning a COM host. The runs
/// were aborted by hand, and an aborted run skips the teardown sweep, which is how 7
/// tagged items were left sitting in a real mailbox. "Outlook is not responding, refusing
/// the live tier" in milliseconds is an enormously better outcome than that.
/// </para>
/// <para>
/// It costs nothing to ask. <see cref="OutlookLiveness"/> reads Windows' own judgement of
/// whether Outlook's UI thread is servicing its message queue (<c>IsHungAppWindow</c>) in
/// microseconds, touching no COM, so the probe can never join the problem it reports on.
/// The shipped server already gates every request this way (<c>ComHostSupervisor</c>);
/// this is the same question asked one layer up, for a tier that has no supervisor.
/// </para>
/// <para>
/// WHY IT FAILS RATHER THAN SKIPS. Four reasons, and the first is mechanical: on
/// xunit 2.9.3 there is no dynamic skip, and a collection fixture constructor can only
/// return or throw, so a throw is the only signal available at the one place that runs
/// before every live collection. Beyond that, every other guard in this tier is
/// fail-closed by policy (no census, no live tier; no signature snapshot, no suite; no
/// settings file, no tier) and a health gate that skipped would be the single guard able
/// to report green on a run that tested nothing. The tier is excluded from CI
/// (<c>Category!=Live</c>) and only ever runs because a human asked for it on the dev
/// machine, so a quiet skip has no audience it would help. And the whole value here is
/// LOUDNESS: the failure names what is wrong, why the tier will not start, and what to do
/// about it.
/// </para>
/// </summary>
public static class LiveOutlookPreflight
{
    /// <summary>
    /// The same override the COM host supervision tests use to force an observed state
    /// (<c>ComHostSupervisor.LivenessOverrideVariable</c>). Shared deliberately: it means
    /// "what this process is to believe about Outlook", and having the tier and the
    /// supervisor disagree about that would be worse than either answer.
    /// <para>
    /// Read on every call rather than once at type load, unlike the supervisor's copy,
    /// because a test has to be able to set it AFTER the assembly is loaded in order to
    /// drive the refusal path on a machine whose Outlook is perfectly healthy.
    /// </para>
    /// </summary>
    public const string LivenessOverrideVariable = "OUTLOOKAI_COMHOST_LIVENESS";

    /// <summary>
    /// How long a still-starting Outlook is given to finish starting.
    /// <para>
    /// Generous on purpose. A cold start on a large profile takes a while, and refusing a
    /// tier because it was asked for eight seconds too early would be a worse gate than no
    /// gate at all. Only ever spent when Outlook is genuinely mid-start, and only once per
    /// run: the eight collections after the first re-ask the free question and skip the
    /// wait.
    /// </para>
    /// </summary>
    internal const int SettleBudgetMilliseconds = 30_000;

    /// <summary>Gap between settle polls. The probe is free, so this is only about not spinning.</summary>
    internal const int SettlePollMilliseconds = 500;

    /// <summary>
    /// Pause before the second opinion on a hung reading.
    /// <para>
    /// Confirm before crying, the same discipline the count tripwire applies to a
    /// suspected loss: <c>IsHungAppWindow</c> reports a window whose thread has not
    /// serviced its queue recently, and a single long synchronous operation on the UI
    /// thread can look exactly like that for a moment. A real wedge does not clear in two
    /// seconds; a busy moment does.
    /// </para>
    /// </summary>
    internal const int ConfirmationDelayMilliseconds = 2_000;

    /// <summary>
    /// What one observed state means for the tier. Pure and total, so the whole decision
    /// is pinned in CI without an Outlook of any kind.
    /// </summary>
    public static LivePreflightVerdict Decide(OutlookLivenessState state)
    {
        switch (state)
        {
            case OutlookLivenessState.Hung:
                // The only hard refusal. OutlookLiveness reports this only when EVERY
                // candidate UI window is hung, which is conservative by design.
                return LivePreflightVerdict.Refuse;

            case OutlookLivenessState.Starting:
                return LivePreflightVerdict.Settle;

            case OutlookLivenessState.NotRunning:
                // Not a fault: the fixtures connect with allowStartingOutlook, so a
                // missing Outlook is a thing the tier is entitled to fix (S7/D17). The
                // preflight must never start Outlook itself - a free probe that starts
                // processes is no longer free.
                return LivePreflightVerdict.Proceed;

            case OutlookLivenessState.Responsive:
            default:
                return LivePreflightVerdict.Proceed;
        }
    }

    /// <summary>The refusal text. Names the observation, the consequence, the history and the fix.</summary>
    public static string RefusalMessage(string detail)
    {
        return "REFUSING to run the live tier: Outlook is not responding (" + detail + ")."
            + Environment.NewLine
            + "Windows reports every Outlook UI window as hung, meaning its thread has stopped servicing its "
            + "message queue. Every COM call this tier makes would block instead of returning, and none of them "
            + "can be cut short: the fixtures, the count tripwire and LiveOutlookTestMailer all attach to Outlook "
            + "IN-PROCESS, and an in-process COM call is not cancellable."
            + Environment.NewLine
            + "This is the 2026-08-18 incident. A test ran 22.5 minutes, then fixture setup sat for 10 and then 15 "
            + "minutes with no output, never spawning a COM host. The runs were aborted by hand - and an aborted "
            + "run skips the teardown sweep, which is how 7 tagged items were left in a real mailbox. Refusing in "
            + "milliseconds is the better outcome."
            + Environment.NewLine
            + "Fix: restart Outlook, ending it from Task Manager if it will not close (on 2026-08-18 the wedged "
            + "instance could be neither started nor killed), let it finish loading, then run the tier again. To "
            + "override this gate deliberately, set " + LivenessOverrideVariable + "=Responsive.";
    }

    /// <summary>
    /// Whether this process has already spent its settle window. The wedge check is asked
    /// per collection because it is free; the WAITING is offered once per run, because a
    /// gate that can cost eight settle windows is a gate that gets blamed for the delay it
    /// was written to remove.
    /// </summary>
    private static int _settleSpent;

    /// <summary>
    /// Gates the live tier. Returns quickly when Outlook is usable; throws with
    /// <see cref="RefusalMessage"/> when it is wedged.
    /// </summary>
    public static void Require()
    {
        int budget = Interlocked.Exchange(ref _settleSpent, 1) == 0 ? SettleBudgetMilliseconds : 0;
        Require(ProbeOutlook, Wait, budget);
    }

    /// <summary>
    /// The gate with its clock, its probe and its settle budget injected, so the refusal
    /// path - including the settle window and the second opinion - is provable in CI
    /// without an unresponsive Outlook and without waiting out any real time.
    /// </summary>
    internal static void Require(
        Func<(OutlookLivenessState State, string Detail)> probe,
        Action<int> wait,
        int settleBudgetMilliseconds = SettleBudgetMilliseconds)
    {
        ArgumentNullException.ThrowIfNull(probe);
        ArgumentNullException.ThrowIfNull(wait);

        (OutlookLivenessState state, string detail) = probe();

        int settled = 0;
        while (Decide(state) == LivePreflightVerdict.Settle && settled < settleBudgetMilliseconds)
        {
            wait(SettlePollMilliseconds);
            settled += SettlePollMilliseconds;
            (state, detail) = probe();
        }

        if (Decide(state) != LivePreflightVerdict.Refuse)
        {
            // Still Starting after the whole budget is reported, not refused: a slow cold
            // start is not a wedge, and the fixtures' own connect will wait for it.
            Console.WriteLine(
                "[preflight] Outlook " + OutlookLiveness.Describe(state) + " (" + detail + ")"
                + (settled > 0 ? ", after " + settled + " ms settling" : string.Empty)
                + " - live tier may run.");
            return;
        }

        wait(ConfirmationDelayMilliseconds);
        (OutlookLivenessState second, string secondDetail) = probe();
        if (Decide(second) != LivePreflightVerdict.Refuse)
        {
            Console.WriteLine(
                "[preflight] Outlook looked hung (" + detail + ") but recovered on the second look ("
                + secondDetail + ") - treating the first reading as a busy moment.");
            return;
        }

        Console.WriteLine("[preflight] Outlook is NOT RESPONDING (" + secondDetail + ") - refusing the live tier.");
        throw new InvalidOperationException(RefusalMessage(secondDetail));
    }

    /// <summary>
    /// Parses the override value. Separate and pure so the parsing rules can be pinned
    /// without depending on what the machine's own Outlook happens to be doing.
    /// </summary>
    internal static bool TryReadOverride(string? raw, out OutlookLivenessState state)
    {
        // IsDefined as well as TryParse: TryParse happily accepts any integer, and a typo
        // silently becoming an undefined state is exactly the kind of quiet misreading
        // this gate exists to prevent. An unrecognised value means "no override".
        return Enum.TryParse(raw, ignoreCase: true, out state) && Enum.IsDefined(state);
    }

    /// <summary>The real probe, with the shared override applied.</summary>
    private static (OutlookLivenessState State, string Detail) ProbeOutlook()
    {
        OutlookLivenessState state = OutlookLiveness.Probe(out string detail);
        if (TryReadOverride(Environment.GetEnvironmentVariable(LivenessOverrideVariable), out OutlookLivenessState forced))
        {
            return (forced, "forced by " + LivenessOverrideVariable);
        }

        return (state, detail);
    }

    private static void Wait(int milliseconds)
    {
        Thread.Sleep(milliseconds);
    }
}
