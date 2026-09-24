using System.Security;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The explicit, per-run opt-in every <c>Category=Live</c> test passes before it can touch
/// anything. Checked by <see cref="Require"/>, which is the first statement of
/// <see cref="LiveTestSettings.Load"/> - and that is the first statement of every live
/// collection fixture, which xunit constructs before any test in the collection runs.
/// <c>T1/LiveRunOptInTests</c> proves both halves from the compiled code, for every live class.
/// <para>
/// <b>THE ACCIDENT IT CLOSES (2026-09-24).</b> A targeted <c>dotnet test</c> used a name filter
/// (<c>~MoveArchive</c>) without <c>Category!=Live</c>, and selected three live tests. They failed
/// only because that worktree had no settings file. The same command in the main checkout -
/// which holds the maintainer's real <c>live-test-settings.json</c> - would have run them,
/// writes included, against his real mailbox. A settings file is permanent, so it cannot be what
/// says "this run is meant to reach a mailbox": that has to be said per run, by the person or
/// agent starting it.
/// </para>
/// <para>
/// <b>WHY AN ENVIRONMENT VARIABLE WHOSE VALUE IS THE COMPUTER NAME.</b> An environment variable is
/// the one channel <c>dotnet test</c> carries into the test host, and the tripwire's bounded
/// re-run child, unchanged, and it belongs to the shell session that sets it - it ends with that
/// session. Its VALUE is the computer name, not "1" or "true", so that the opt-in names the one
/// machine it was given for: a value carried to another machine - a copied script, a roaming
/// profile, a CI configuration - opens nothing there. And it is refused when it is SAVED in the
/// user or machine environment (<c>setx</c>, System Properties), even on its own machine, because
/// a saved opt-in would open every later run there, the accidental ones included - which is the
/// exact thing it exists to stop. A PowerShell profile that sets it at every start defeats the
/// same property invisibly, and nothing here can see that; the documentation says not to.
/// </para>
/// <para>
/// <b>WHAT IT IS NOT.</b> It says only that a live run was meant. It does not make any test
/// read-only, and it does not decide which machine may run which test. The maintainer's
/// workstation is read-only for live tests, always (CLAUDE.md, Mailbox Safety); how that rule is
/// enforced in code is the maintainer's open decision (Q74), and nothing here pre-empts it.
/// </para>
/// </summary>
public static class LiveRunOptIn
{
    /// <summary>The variable a deliberate live run sets, for that run only, to its machine's computer name.</summary>
    public const string Variable = "OUTLOOKAI_LIVE_OPT_IN";

    /// <summary>What <see cref="Evaluate"/> decided.</summary>
    public enum Verdict
    {
        /// <summary>This run opted in, for this machine.</summary>
        Open = 0,

        /// <summary>The variable is not set for this run.</summary>
        Missing = 1,

        /// <summary>The variable names a different machine.</summary>
        OtherMachine = 2,

        /// <summary>The variable is saved in the user or machine environment, so it is not a per-run opt-in.</summary>
        Persisted = 3,
    }

    /// <summary>
    /// The whole decision, pure so every case is pinned in T1. A saved value refuses first,
    /// whatever it says: it would open every later run on this machine, not only this one.
    /// </summary>
    /// <param name="processValue">The variable as this process sees it.</param>
    /// <param name="persistedUserValue">The variable as saved in the user environment, if at all.</param>
    /// <param name="persistedMachineValue">The variable as saved in the machine environment, if at all.</param>
    /// <param name="machineName">This machine's computer name.</param>
    public static Verdict Evaluate(
        string? processValue,
        string? persistedUserValue,
        string? persistedMachineValue,
        string machineName)
    {
        if (!string.IsNullOrWhiteSpace(persistedUserValue) || !string.IsNullOrWhiteSpace(persistedMachineValue))
        {
            return Verdict.Persisted;
        }

        if (string.IsNullOrWhiteSpace(processValue))
        {
            return Verdict.Missing;
        }

        return string.Equals(processValue.Trim(), machineName, StringComparison.OrdinalIgnoreCase)
            ? Verdict.Open
            : Verdict.OtherMachine;
    }

    /// <summary>
    /// The gate. Returns only when this run opted in for this machine; otherwise throws, and a
    /// collection fixture that throws in its constructor stops every test in its collection
    /// before a test class is built or a test method runs.
    /// </summary>
    public static void Require()
    {
        string machine = Environment.MachineName;
        string? processValue = Environment.GetEnvironmentVariable(Variable, EnvironmentVariableTarget.Process);
        Verdict verdict = Evaluate(
            processValue,
            ReadSaved(EnvironmentVariableTarget.User),
            ReadSaved(EnvironmentVariableTarget.Machine),
            machine);
        if (verdict != Verdict.Open)
        {
            throw new InvalidOperationException(DescribeRefusal(verdict, machine, processValue));
        }
    }

    /// <summary>
    /// What a refused run is told: what the opt-in is for, how to give it for a deliberate run on
    /// a test guest, and that the maintainer's workstation is read-only for live tests. Pure, and
    /// public for T1.
    /// </summary>
    public static string DescribeRefusal(Verdict verdict, string machineName, string? processValue)
    {
        string lead;
        switch (verdict)
        {
            case Verdict.OtherMachine:
                lead = "LIVE TEST REFUSED: " + Variable + " is set to '" + (processValue ?? string.Empty).Trim()
                    + "', but this machine is '" + machineName + "' - an opt-in names the one machine it was given for.";
                break;
            case Verdict.Persisted:
                lead = "LIVE TEST REFUSED: " + Variable + " is saved in the user or machine environment (or that could "
                    + "not be read), so it would open every later run on this machine, not only this one. Remove it - "
                    + "[Environment]::SetEnvironmentVariable('" + Variable + "', $null, 'User'), and 'Machine' from an "
                    + "elevated shell - and set it in the session instead.";
                break;
            default:
                lead = "LIVE TEST REFUSED: this run did not opt in - " + Variable + " is not set.";
                break;
        }

        return lead
            + " Category=Live tests act on a real Outlook profile - its mail, its folders, its signatures - so every run "
            + "must opt in explicitly. The opt-in exists so that a test filter which merely happens to select live tests "
            + "(a name filter without Category!=Live, say) cannot reach a mailbox by accident."
            + " For a deliberate run ON A TEST GUEST, set " + Variable + " to that guest's own computer name, in the same "
            + "PowerShell session and for that run only, and start the run from there:"
            + " $env:" + Variable + " = '<the guest's computer name; $env:COMPUTERNAME prints it>'"
            + " then dotnet test McpServer\\OutlookAI.McpServer.Tests\\OutlookAI.McpServer.Tests.csproj --filter "
            + "\"Category=Live&Requires!=DelegateStore\" - on a guest both lines go into the -Script of "
            + "Testbed/guest/Register-InteractiveTask.ps1 (Testbed/README.md, section 4c)."
            + " The value must equal the computer name of the machine the run is on, so a value carried to another "
            + "machine opens nothing, and it must never be saved with setx or in the user or machine environment."
            + " THE MAINTAINER'S WORKSTATION IS READ-ONLY FOR LIVE TESTS, ALWAYS (CLAUDE.md, Mailbox Safety): this "
            + "variable only says a run was intended - it makes no test read-only, and it must never be set there to "
            + "run a test that can write.";
    }

    /// <summary>
    /// The variable as saved at <paramref name="target"/>. An environment that cannot be read
    /// answers a non-empty placeholder: "could not prove it is not saved" must refuse, not open.
    /// </summary>
    private static string? ReadSaved(EnvironmentVariableTarget target)
    {
        try
        {
            return Environment.GetEnvironmentVariable(Variable, target);
        }
        catch (Exception ex) when (ex is SecurityException || ex is UnauthorizedAccessException || ex is IOException)
        {
            return "<unreadable>";
        }
    }
}
