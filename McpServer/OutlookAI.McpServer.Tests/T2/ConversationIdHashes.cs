using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Which string, if any, an Outlook ConversationId is an MD5 hash of - the derivation MS-OXOMSG
/// gives for an item whose id does not come from its conversation-index GUID. Pure, so T1 pins it;
/// the live drafts tests print what it finds, by candidate LABEL only (S4: never the text).
/// </summary>
public static class ConversationIdHashes
{
    /// <summary>The encodings and casings tried for every candidate, in the order they are reported.</summary>
    public static readonly IReadOnlyList<string> Variants = new[] { "utf16", "utf16-upper", "utf8", "utf8-upper" };

    /// <summary>
    /// The hex MD5 of <paramref name="text"/> under one of <see cref="Variants"/>, upper-case like
    /// Outlook's ConversationID.
    /// </summary>
    public static string Md5Hex(string text, string variant)
    {
        ArgumentNullException.ThrowIfNull(text);
        string cased = variant.EndsWith("-upper", StringComparison.Ordinal) ? text.ToUpperInvariant() : text;
        byte[] bytes = variant.StartsWith("utf16", StringComparison.Ordinal)
            ? Encoding.Unicode.GetBytes(cased)
            : Encoding.UTF8.GetBytes(cased);
        return Convert.ToHexString(MD5.HashData(bytes));
    }

    /// <summary>
    /// "label/variant" for every candidate whose hash equals <paramref name="conversationId"/>, joined
    /// by commas; "none" when no candidate matches, and "-" when there is no id to compare.
    /// </summary>
    public static string Describe(string? conversationId, params (string Label, string? Text)[] candidates)
    {
        if (string.IsNullOrEmpty(conversationId))
        {
            return "-";
        }

        List<string> matches = new List<string>();
        foreach ((string label, string? text) in candidates)
        {
            if (text == null)
            {
                continue;
            }

            foreach (string variant in Variants)
            {
                if (string.Equals(Md5Hex(text, variant), conversationId, StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add(label + "/" + variant);
                }
            }
        }

        return matches.Count == 0 ? "none" : string.Join(",", matches);
    }
}
