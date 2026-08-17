using System.Security.Cryptography;
using System.Text;
using OutlookAI.Services;

namespace OutlookAI.McpServer.Tools;

/// <summary>
/// THE USER'S OWN WRITING RULES, DELIVERED IN A REJECTION.
/// <para>
/// The user writes the rules their email is to follow in OutlookAI's prompt settings (the
/// "Always sent" section, <see cref="PromptSection.Preamble"/>). The sidebar sends that text
/// to the model on every writing request. An agent driving the MCP tools composed the body
/// itself and never saw a word of it, so the same person got two different voices out of one
/// product depending on which half of it they used.
/// </para>
/// <para>
/// It cannot travel on the schema. Everything a client reads as prose - the server
/// <c>instructions</c>, a tool <c>description</c>, a parameter <c>description</c> - is built
/// from compile-time constants, so it physically cannot carry text a user edits at runtime;
/// and those strings are capped anyway (Claude Code truncates them at 2 KB, silently and
/// mid-sentence, which <c>DescriptionBudgetCiTests</c> guards). The rules run to pages and
/// can change between two tool calls.
/// </para>
/// <para>
/// So they travel in an ERROR, which is the one channel that is per-call, unbudgeted, and
/// already understood as something to act on: the MCP specification's guidance is that a
/// tool error carries actionable feedback the model can self-correct from. The first drafting
/// call of a server process is refused with the rules attached and an instruction to compose
/// again; the retry goes through. The user accepted that cost explicitly. It is paid once per
/// session, and once more each time the rules change - the gate re-arms on the text itself,
/// so an edit mid-session costs exactly one more rejection and nothing has to be restarted.
/// </para>
/// <para>
/// Public, unlike most of this assembly, so the T1 suite can drive it against an injected
/// rules source instead of the developer's own HKCU.
/// </para>
/// </summary>
public sealed class WritingRulesGate
{
    /// <summary>
    /// Stable <c>error.type</c> for the rejection. Stable because an agent may reasonably
    /// branch on it, and because the T3 wire test tells "the gate refused this" from "the
    /// gate is done and something else refused it" by this string alone.
    /// </summary>
    public const string ErrorType = "WritingRulesRequired";

    /// <summary>
    /// What to do about it. The first sentence exists to stop an agent reporting a failure to
    /// the user: this is a retry, the call was not attempted, and nothing in the mailbox was
    /// touched.
    /// </summary>
    public const string RetryMessage =
        "Nothing failed and nothing was changed. Compose the body again so that it follows the user's "
        + "writing rules in writingRules below, then call this tool again with the same arguments. "
        + "This happens once per session, and again only when the user edits their rules.";

    /// <summary>
    /// What the rules are and how far they reach. The last half is not padding: the rules are
    /// written for the sidebar, which inserts plain text into the Word editor behind a compose
    /// window, so one of their lines bans HTML tags. Handed to an agent unqualified, that line
    /// reads as a ban on <c>body_html</c>, which is a supported and often correct way to use
    /// these tools.
    /// </summary>
    public const string Clarification =
        "writingRules is the user's own writing prompt, verbatim, from OutlookAI's prompt settings "
        + "(the \"Always sent\" section) - the same text the OutlookAI sidebar sends when the user writes "
        + "mail there. It governs the BODY you compose: language, tone, structure, and its own content "
        + "rules. It does not replace this tool's contract, which still holds: supply exactly one of body "
        + "or body_html. It was also written for the sidebar, which inserts PLAIN TEXT into the draft, so "
        + "read its no-code-fences/no-HTML-tags line as a rule about a plain-text body - body_html is an "
        + "HTML fragment and is still a supported way to call this tool. Retry now.";

    private readonly Func<string> _readRules;
    private readonly object _sync = new();

    /// <summary>
    /// Hash of the rules text last handed out, or null while none has been. Guarded by
    /// <see cref="_sync"/>: two tool calls can be in flight at once, and the whole point of
    /// the gate is that exactly one of them pays the rejection.
    /// </summary>
    private string? _deliveredHash;

    /// <param name="readRules">
    /// Reads the current rules text. Called on every gated tool call rather than cached, for
    /// the same reason the store itself caches nothing: the user can edit the text while the
    /// server runs, and stale rules are worse than the microseconds.
    /// </param>
    public WritingRulesGate(Func<string> readRules)
    {
        ArgumentNullException.ThrowIfNull(readRules);
        _readRules = readRules;
    }

    /// <summary>
    /// The process-wide gate the drafting tools consult, reading the user's live overrides and
    /// the shipped defaults from the one store the add-in itself uses.
    /// </summary>
    public static WritingRulesGate Shared { get; } =
        new(() => PromptStore.GetSection(PromptSection.Preamble));

    /// <summary>
    /// Decides whether this call must be refused so the rules can be delivered, and records
    /// the delivery in the same breath - the record is taken when the rejection is EMITTED,
    /// which is what makes the caller's retry succeed.
    /// </summary>
    /// <param name="composesBody">
    /// Whether this call actually supplies body text. False is never gated: <c>update_draft</c>
    /// may legitimately change only recipients, a subject or attachments, and rules about how
    /// to write have nothing to say about those.
    /// </param>
    /// <param name="rules">The text to deliver, verbatim as the user wrote it. Empty when false is returned.</param>
    /// <returns>True when the caller must return the rejection instead of doing the work.</returns>
    public bool TryClaimDelivery(bool composesBody, out string rules)
    {
        rules = string.Empty;
        if (!composesBody)
        {
            return false;
        }

        string current;
        try
        {
            current = _readRules() ?? string.Empty;
        }
        catch (Exception)
        {
            // A rules source that cannot answer must not cost the user their draft. The
            // shipped one (PromptStore) is documented never to throw, so this is the injected
            // and fabricated cases; either way, failing open means a draft written without the
            // rules, and failing closed would mean no draft at all.
            return false;
        }

        if (string.IsNullOrWhiteSpace(current))
        {
            // The user cleared their rules. There is nothing to deliver, so there is nothing
            // to refuse a call for.
            return false;
        }

        string hash = HashOf(current);
        lock (_sync)
        {
            if (string.Equals(_deliveredHash, hash, StringComparison.Ordinal))
            {
                return false;
            }

            _deliveredHash = hash;
        }

        rules = current;
        return true;
    }

    /// <summary>
    /// Identity of a rules text, for "has this already been delivered?".
    /// <para>
    /// Line endings are normalised (and the ends trimmed) before hashing, by the same
    /// <see cref="PromptDefaults.Normalize"/> the store uses to decide whether an edit is a
    /// real customisation. A multiline text box hands back CRLF where the source had LF and
    /// often a trailing newline, and neither changes a single word the model reads - so
    /// neither may cost the agent a second rejection.
    /// </para>
    /// </summary>
    private static string HashOf(string rules)
    {
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(PromptDefaults.Normalize(rules))));
    }
}
