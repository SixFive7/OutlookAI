using OutlookAI.McpServer.Tests.T2;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>What one MCP tool can do to the machine the server runs on.</summary>
public enum McpToolEffect
{
    /// <summary>
    /// Not in <see cref="McpToolWriteClassification.Tools"/>. Deliberately the ZERO value: a tool
    /// nobody classified is treated as one that writes, and refused under the read-only posture.
    /// </summary>
    Unclassified = 0,

    /// <summary>
    /// Changes nothing outside the server process: no mail item or folder, no signature, no
    /// registry value, no file, nothing on the user's screen. It may attach to Outlook, start it,
    /// and query the Windows Search index - all reads.
    /// </summary>
    ReadOnly = 1,

    /// <summary>Can change something the server process does not own. See each entry's reason.</summary>
    Writes = 2,
}

/// <summary>One classified tool: what it can do, and the reason, said once.</summary>
/// <param name="Effect">Whether the tool is read-only.</param>
/// <param name="Why">One sentence a reviewer can check against the tool's code.</param>
public sealed record McpToolClassification(McpToolEffect Effect, string Why);

/// <summary>
/// THE pinned classification of every tool the MCP server exposes, read-only or write-capable -
/// the one source the test-side client's write gate (Q74 layer 3) consults.
/// <para>
/// <b>Why a list and not the tools' own metadata.</b> The decision asked for the write-capable set
/// to be derived from the tools' read-only hints if they carried any. They do not, with one
/// exception: only <c>audit_log</c> (Q93) sets <c>ReadOnly</c>/<c>Destructive</c> in
/// <c>Tools/OutlookTools.cs</c>, so the wire carries no <c>readOnlyHint</c> for the rest. Adding them is a change
/// to what every MCP client of the product sees, which is not this change's to make (raised as a
/// question in the Q74 report). So: one explicit list, here, and
/// <c>T1.McpToolWriteClassificationTests</c> fails the build until it names EVERY tool the server
/// assembly declares and nothing else.
/// </para>
/// <para>
/// <b>A new tool is a deliberate classification, never a default.</b> Until somebody adds it here
/// the T1 pin is red, and at run time it is <see cref="McpToolEffect.Unclassified"/>, which the
/// read-only posture refuses exactly like a write. The audit-log reader added for Q93,
/// <c>audit_log</c>, was the first case: it reads the server's own log and changes nothing, so it
/// is classified <see cref="McpToolEffect.ReadOnly"/> below, by the edit that landed it with Q74.
/// </para>
/// <para>
/// <b>Read-only means the strict thing.</b> Not "writes no mail" but "changes nothing outside the
/// server process": the maintainer's workstation is read-only for live tests ALWAYS, and a test that
/// pops a window over his work or marks his mail read has changed his machine. Hence the three
/// navigation tools are <see cref="McpToolEffect.Writes"/>.
/// </para>
/// </summary>
public static class McpToolWriteClassification
{
    /// <summary>Every tool the server declares, by its wire name.</summary>
    public static IReadOnlyDictionary<string, McpToolClassification> Tools { get; } =
        new Dictionary<string, McpToolClassification>(StringComparer.Ordinal)
        {
            ["search"] = new(McpToolEffect.ReadOnly,
                "queries the Windows Search index and sweeps Outlook folders by reading them; its cache is in-process"),
            ["thread"] = new(McpToolEffect.ReadOnly,
                "reads a conversation from the index or through Outlook"),
            ["read"] = new(McpToolEffect.ReadOnly,
                "reads one item's properties and body; reads UnRead and never sets it"),
            ["list_accounts"] = new(McpToolEffect.ReadOnly,
                "reads the profile's accounts and stores"),
            ["list_folders"] = new(McpToolEffect.ReadOnly,
                "walks a folder tree by reading it (a read-only lookup never creates a folder since Q84)"),
            ["list_signatures"] = new(McpToolEffect.ReadOnly,
                "reads the signature directory and the profile's signature registry values"),
            ["outlook_health"] = new(McpToolEffect.ReadOnly,
                "reads store and index state; may start Outlook, never changes it"),
            ["audit_log"] = new(McpToolEffect.ReadOnly,
                "reads the server's own audit log, changes nothing; never touches Outlook (Q93)"),

            ["save_attachment"] = new(McpToolEffect.Writes,
                "writes the attachment to a file on disk"),
            ["move_mail"] = new(McpToolEffect.Writes,
                "moves items between folders, and can create the target folder"),
            ["archive_mail"] = new(McpToolEffect.Writes,
                "moves items to the Archive folder, creating it if it is missing"),
            ["manage_signature"] = new(McpToolEffect.Writes,
                "creates, updates or deletes signature files and the profile registry values that select them"),
            ["open_in_outlook"] = new(McpToolEffect.Writes,
                "opens an item in an Outlook window - which marks an unread item read - over the user's work"),
            ["goto_folder"] = new(McpToolEffect.Writes,
                "changes the folder the user's Outlook window shows"),
            ["show_search_results"] = new(McpToolEffect.Writes,
                "runs a search in the user's Outlook window, replacing what it shows"),
            ["new_draft"] = new(McpToolEffect.Writes, "creates a draft"),
            ["reply_draft"] = new(McpToolEffect.Writes, "creates a reply draft"),
            ["replyall_draft"] = new(McpToolEffect.Writes, "creates a reply-all draft"),
            ["forward_draft"] = new(McpToolEffect.Writes, "creates a forward draft"),
            ["update_draft"] = new(McpToolEffect.Writes, "rewrites an existing draft"),
            ["discard_draft"] = new(McpToolEffect.Writes, "moves a draft to Deleted Items"),
            ["send"] = new(McpToolEffect.Writes, "sends mail"),
        };

    /// <summary>The phrase every refusal carries, so a run log and the T1 pins can find it.</summary>
    public const string Refused = "WRITE-CAPABLE MCP TOOL REFUSED";

    /// <summary>What <paramref name="tool"/> can do, or <see cref="McpToolEffect.Unclassified"/>.</summary>
    public static McpToolEffect Classify(string? tool)
    {
        return tool != null && Tools.TryGetValue(tool, out McpToolClassification? entry)
            ? entry.Effect
            : McpToolEffect.Unclassified;
    }

    /// <summary>
    /// The refusal for one <c>tools/call</c> under <paramref name="posture"/>, or null to send it.
    /// Pure, so T1 drives every combination.
    /// <list type="bullet">
    /// <item><see cref="StdioWritePosture.ReadOnly"/>: only a <see cref="McpToolEffect.ReadOnly"/>
    /// tool goes out. A write-capable tool is refused, and so is an unclassified one - a name not in
    /// the list, a misspelling, a tool added after this list was last read - because nobody has
    /// shown it changes nothing.</item>
    /// <item>Any other posture: nothing is refused here. A test guest may write, and a run that did
    /// not opt in is the CI tier, whose own guards apply.</item>
    /// </list>
    /// </summary>
    public static string? RefusalFor(StdioWritePosture posture, string? tool)
    {
        if (posture != StdioWritePosture.ReadOnly)
        {
            return null;
        }

        McpToolEffect effect = Classify(tool);
        if (effect == McpToolEffect.ReadOnly)
        {
            return null;
        }

        string what = effect == McpToolEffect.Writes
            ? "'" + tool + "' is classified write-capable (" + Tools[tool!].Why + ")"
            : "'" + (tool ?? "(no tool name)") + "' is not classified at all, and a tool nobody has shown to be read-only "
                + "is refused like a write";
        return Refused + " before it was sent: " + what + ". This process is a live run on a READ-ONLY machine - "
            + "its live-test settings declare a profile that may write nothing (see LiveWriteAccess), which is the "
            + "maintainer's workstation - so the test-side client sends only tools classified read-only in "
            + nameof(McpToolWriteClassification) + ". A test that needs this tool runs on a test guest. A NEW read-only "
            + "tool is classified there, with its reason, never by loosening this gate.";
    }
}
