namespace OutlookAI.McpServer;

/// <summary>
/// Server-level MCP metadata. <see cref="Instructions"/> is returned verbatim as the
/// <c>instructions</c> field of the MCP <c>initialize</c> result. Claude Code injects it
/// passively into EVERY session at start - including sessions whose MCP tools are deferred
/// name-only by tool search (probe-proven, v3.MD D36) - so it must stay short: the budget
/// here is a fraction of Claude Code's 2 KB truncation cap, and T1 pins it. Written per the
/// official authoring guidance (task category + when to reach for the tools + key
/// capabilities, critical details first, keyword-rich for tool-search matching).
/// Pinned exactly by T3 (the initialize result must carry this string verbatim).
/// </summary>
public static class ServerMetadata
{
    public const string Instructions =
        "Local Outlook email access: search and read mail across all accounts, folders, and "
        + "delegate mailboxes; follow threads; save/read attachments; open results in Outlook. "
        + "Use these tools whenever an answer may live in the user's email or inbox. "
        + "Reply/forward/new drafts open in Outlook for the user to review; send is confirmation-gated.";
}
