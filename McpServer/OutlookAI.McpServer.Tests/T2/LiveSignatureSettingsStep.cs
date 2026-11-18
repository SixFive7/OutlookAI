namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The live runner's <c>-SignatureSettings</c> step (Q200 (a), Q201, decided by the maintainer 2026-11-18), as the live
/// tier sees it: which signature setting an administrator made on this test guest BEFORE Outlook started -
/// <c>Testbed/guest/Set-SignatureSettings.ps1</c>, run elevated by <c>Testbed/host/Invoke-LiveTierOnGuest.ps1</c> - and
/// how what a compose window shows is told apart against the signatures that setting and the hub account's own
/// defaults name.
/// <para>
/// <b>Why a runner step and not the test.</b> Outlook reads an administrator's policy and its accounts' signature
/// settings when it starts; a value written while it runs was not seen (G1 and G2, 2026-11-17). And the live tier
/// runs NOT elevated, where <c>HKCU\Software\Policies</c> is read-only. So the runner writes them, elevated, before
/// it starts Outlook, and the guest's checkpoint restore at the end of the run undoes every one of them.
/// </para>
/// <para>
/// The runner names the step it took in <see cref="VariableName"/>, inside the guest task's script, for that run
/// only - as it names the Exchange VM's run marker. No variable: this run had no such step. Pure but for
/// <see cref="Current"/>.
/// </para>
/// </summary>
public static class LiveSignatureSettingsStep
{
    /// <summary>The variable the live runner sets inside the guest task for a run with the step.</summary>
    public const string VariableName = "OUTLOOKAI_LIVE_SIGNATURE_SETTINGS";

    /// <summary>Q200: Group Policy forces the default for new mail and for replies and forwards.</summary>
    public const string GroupPolicy = "GroupPolicy";

    /// <summary>Q201: a default deployed in Outlook's general mail settings, <c>Common\MailSettings</c>.</summary>
    public const string Preference = "Preference";

    /// <summary>Every scenario the guest script sets, as the runner's parameter spells it.</summary>
    public static IReadOnlyList<string> Scenarios { get; } = new[] { GroupPolicy, Preference };

    /// <summary>What a reading shows when no signature went in.</summary>
    public const string NoSignature = "(none)";

    /// <summary>What a reading shows when a signature went in that is none of the candidates.</summary>
    public const string AnotherSignature = "(another signature)";

    /// <summary>What a reading shows when more than one candidate went in.</summary>
    public const string SeveralSignatures = "(several signatures)";

    /// <summary>What a reading shows when nothing could be read.</summary>
    public const string NotRead = "(not read)";

    /// <summary>The scenario this run's step set, or null when it had none.</summary>
    public static string? Current => Parse(Environment.GetEnvironmentVariable(VariableName));

    /// <summary>
    /// The scenario a value of <see cref="VariableName"/> names: null for none, a scenario as spelled, or a refusal
    /// for anything else - a run whose step this suite cannot read must not pass as one without a step.
    /// </summary>
    public static string? Parse(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        string trimmed = value.Trim();
        return Scenarios.FirstOrDefault(s => string.Equals(s, trimmed, StringComparison.Ordinal))
            ?? throw new InvalidOperationException(
                VariableName + "='" + trimmed + "' names no signature setting the live runner makes - "
                + string.Join(" or ", Scenarios) + ". Refusing to read the run as one without the step.");
    }

    /// <summary>
    /// Which of <paramref name="candidates"/> - signature names - a compose window shows. Outlook writes a signature it
    /// inserts into the <c>_MailAutoSig</c> region, and the test signatures carry their own name as text, so the region
    /// is searched first; the whole document only when the region is empty or absent, which finds a signature put
    /// anywhere else. <see cref="NoSignature"/> when neither holds one, <see cref="AnotherSignature"/> when the region
    /// holds something that is none of them, <see cref="SeveralSignatures"/> when it holds more than one, and
    /// <see cref="NotRead"/> when there was nothing to read. Ordinal: the names are written as the step wrote them.
    /// </summary>
    public static string Classify(string? regionText, string? documentText, IReadOnlyList<string> candidates)
    {
        ArgumentNullException.ThrowIfNull(candidates);
        bool regionHolds = !string.IsNullOrWhiteSpace(regionText);
        string? where = regionHolds ? regionText : documentText;
        if (where == null)
        {
            return NotRead;
        }

        List<string> found = candidates
            .Where(c => !string.IsNullOrEmpty(c) && where.Contains(c, StringComparison.Ordinal))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        // A name that is part of another found name is that other one's text, not a signature of its own.
        found = found.Where(f => !found.Any(o => o.Length > f.Length && o.Contains(f, StringComparison.Ordinal))).ToList();
        if (found.Count == 1)
        {
            return found[0];
        }

        if (found.Count > 1)
        {
            return SeveralSignatures;
        }

        return regionHolds ? AnotherSignature : NoSignature;
    }
}
