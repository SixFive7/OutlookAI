using System.Text.Json;

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

    /// <summary>Q200: Group Policy names the default for new mail and for replies and forwards.</summary>
    public const string GroupPolicy = "GroupPolicy";

    /// <summary>Q201: a default deployed in Outlook's general mail settings, <c>Common\MailSettings</c>.</summary>
    public const string Preference = "Preference";

    /// <summary>
    /// Q201 again, with the hub account's own default for replies instead of new mail (decided 2026-11-18): a deployed
    /// new-mail value for an account with no new-mail default of its own, the case the first two leave open.
    /// </summary>
    public const string PreferenceNoOwnNew = "PreferenceNoOwnNew";

    /// <summary>Every scenario the guest script sets, as the runner's parameter spells it.</summary>
    public static IReadOnlyList<string> Scenarios { get; } = new[] { GroupPolicy, Preference, PreferenceNoOwnNew };

    /// <summary>
    /// The variable the live runner sets inside the guest task of a run with the step: where
    /// <c>Testbed/guest/Measure-SignatureDialog.ps1</c> wrote its reading of Outlook's own Signatures dialog, taken
    /// before the suite started.
    /// </summary>
    public const string DialogVariableName = "OUTLOOKAI_LIVE_SIGNATURE_DIALOG";

    /// <summary>
    /// Whether the step gave the hub account its one own default for replies and forwards - under
    /// <see cref="GroupPolicy"/> and <see cref="PreferenceNoOwnNew"/> - rather than for new mail, under
    /// <see cref="Preference"/>. The scope it left without one shows what the setting does for an account with none.
    /// </summary>
    public static bool HubOwnDefaultIsForReplies(string scenario) => scenario switch
    {
        GroupPolicy => true,
        Preference => false,
        PreferenceNoOwnNew => true,
        _ => throw new ArgumentException("'" + scenario + "' is no signature setting the live runner makes.", nameof(scenario)),
    };

    /// <summary>Where this run's reading of the Signatures dialog is, or null when the runner read none.</summary>
    public static string? DialogReadingPath
    {
        get
        {
            string? value = Environment.GetEnvironmentVariable(DialogVariableName);
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }

    /// <summary>
    /// A reading of Outlook's Signatures dialog as <c>Measure-SignatureDialog.ps1</c> writes it. Windows PowerShell 5.1's
    /// <c>ConvertTo-Json</c> writes a list of one as that one item, so a list is read from either.
    /// </summary>
    public static SignatureDialogReading ParseDialogReading(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        return new SignatureDialogReading(
            Refused: Text(root, "refused"),
            DialogFound: Flag(root, "dialogFound") == true,
            DialogClass: Text(root, "dialogClass"),
            Closed: Flag(root, "closed") == true,
            Account: Control(root, "account"),
            NewMessages: Control(root, "newMessages"),
            RepliesForwards: Control(root, "repliesForwards"),
            Delete: Control(root, "delete"),
            Signatures: Items(root, "signatures")
                .Select(e => new SignatureDialogEntry(Text(e, "name") ?? string.Empty, Flag(e, "selected") == true, Flag(e, "deleteEnabled"), Text(e, "error")))
                .ToList(),
            OtherSignatureItems: root.TryGetProperty("otherSignatureItems", out JsonElement other) && other.ValueKind == JsonValueKind.Number ? other.GetInt32() : 0,
            FolderUnchanged: Flag(root, "folderUnchanged"),
            FolderChanges: Items(root, "folderChanges").Select(e => e.ToString()).ToList(),
            Controls: Items(root, "controls").Select(ControlOf).ToList(),
            Errors: Items(root, "errors").Select(e => e.ToString()).ToList());
    }

    private static IEnumerable<JsonElement> Items(JsonElement parent, string name)
    {
        if (!parent.TryGetProperty(name, out JsonElement value))
        {
            return Array.Empty<JsonElement>();
        }

        return value.ValueKind switch
        {
            JsonValueKind.Array => value.EnumerateArray().ToList(),
            JsonValueKind.Null or JsonValueKind.Undefined => Array.Empty<JsonElement>(),
            _ => new[] { value },
        };
    }

    private static string? Text(JsonElement parent, string name) =>
        parent.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;

    private static bool? Flag(JsonElement parent, string name) =>
        parent.TryGetProperty(name, out JsonElement value) && (value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False)
            ? value.GetBoolean()
            : null;

    private static SignatureDialogControl? Control(JsonElement parent, string name) =>
        parent.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.Object ? ControlOf(value) : null;

    private static SignatureDialogControl ControlOf(JsonElement value) =>
        new(Text(value, "type"), Text(value, "name"), Flag(value, "enabled"), Text(value, "value"));

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

/// <summary>One control of Outlook's Signatures dialog: its UI Automation type, name, whether it is enabled, its value.</summary>
public sealed record SignatureDialogControl(string? Type, string? Name, bool? Enabled, string? Value);

/// <summary>One test signature selected in the dialog's list, and whether its Delete button was enabled then.</summary>
public sealed record SignatureDialogEntry(string Name, bool Selected, bool? DeleteEnabled, string? Error);

/// <summary>What <c>Measure-SignatureDialog.ps1</c> read off Outlook's Signatures dialog (Q156, 2026-11-18).</summary>
public sealed record SignatureDialogReading(
    string? Refused,
    bool DialogFound,
    string? DialogClass,
    bool Closed,
    SignatureDialogControl? Account,
    SignatureDialogControl? NewMessages,
    SignatureDialogControl? RepliesForwards,
    SignatureDialogControl? Delete,
    IReadOnlyList<SignatureDialogEntry> Signatures,
    int OtherSignatureItems,
    bool? FolderUnchanged,
    IReadOnlyList<string> FolderChanges,
    IReadOnlyList<SignatureDialogControl> Controls,
    IReadOnlyList<string> Errors);
