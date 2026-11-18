using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// What the draft suites check about a draft, through the verifying session (Q154): its mail facts and its
/// HTML, read or the test fails; where the agent's body sits in Outlook's own HTML; and an inspector that
/// opens or closes. The one copy of what <c>LiveDraftTests</c>, <c>LiveDraftOptionsTests</c>,
/// <c>LiveHtmlDraftTests</c>, <c>LiveUpdateDiscardTests</c>, <c>LiveSignatureTests</c>,
/// <c>LiveHeadlessComposeParityTests</c> and <c>LiveShowMeTests</c> each kept; a suite whose check needs its
/// fixture's session binds it in one line. Everything here reads; nothing writes.
/// </summary>
internal static class LiveDraftChecks
{
    /// <summary>The draft's mail facts, or the test fails saying why they could not be read.</summary>
    public static ComDraftInfo RequireMailInfo(OutlookComSession session, string entryId, string? storeId)
    {
        ComDraftInfo? info = session.TryGetMailInfo(entryId, storeId, out string? error);
        Assert.True(info != null, $"mail info unavailable: {error ?? "unknown"}");
        return info!;
    }

    /// <summary>The draft's HTML body, or the test fails saying why it could not be read.</summary>
    public static string RequireHtmlBody(OutlookComSession session, string entryId, string storeId)
    {
        string? html = session.TryGetHtmlBody(entryId, storeId, out string? error);
        Assert.True(!string.IsNullOrEmpty(html), $"draft HTML unavailable: {error ?? "empty"}");
        return html!;
    }

    /// <summary>Where <paramref name="needle"/> first stands in the draft HTML, or the test fails naming <paramref name="what"/>.</summary>
    public static int RequireIndexOf(string html, string needle, string what)
    {
        int at = html.IndexOf(needle, StringComparison.Ordinal);
        Assert.True(at >= 0, what + " must be present in the draft HTML");
        return at;
    }

    /// <summary>
    /// The body must live INSIDE Outlook's own document container, not be spliced in
    /// straight after &lt;body&gt; (which is what the retired HTMLBody string surgery did
    /// and why the agent text did not inherit the message style).
    /// </summary>
    public static void AssertBodyInsideWordSection(string html, int bodyAt)
    {
        int section = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        Assert.True(section >= 0, "Outlook's WordSection container must be present");
        Assert.True(bodyAt > section, $"agent body must sit INSIDE WordSection1 (section@{section} body@{bodyAt})");
    }

    /// <summary>
    /// The A1 contract on the RAW HTML: agent body first, then the signature region
    /// anchor, then the signature's own content - and the body is NOT inside the
    /// signature bookmark (the exact shape of the reported defect).
    /// </summary>
    public static void AssertBodyAboveIntactSignature(
        string html, string bodyMarker, string signatureMarker, string label, Action<string> output)
    {
        int bodyAt = RequireIndexOf(html, bodyMarker, "agent body");
        int anchorAt = RequireIndexOf(html, "_MailAutoSig", "signature region");
        int signatureAt = RequireIndexOf(html, signatureMarker, "signature content");

        Assert.True(bodyAt < anchorAt, $"{label}: agent body must precede the signature region (body@{bodyAt} sig@{anchorAt})");
        Assert.True(bodyAt < signatureAt, $"{label}: agent body must precede the signature content");

        // The signature keeps its own HTML: the image resource of the test signature
        // survives as a real <img> element rather than being flattened to text.
        Assert.Contains("<img", html, StringComparison.OrdinalIgnoreCase);
        AssertBodyInsideWordSection(html, bodyAt);
        output($"[{label}]: bodyAt={bodyAt} signatureAnchorAt={anchorAt} signatureContentAt={signatureAt} imgPreserved=true");
    }

    /// <summary>
    /// The open inspector showing <paramref name="entryId"/>, polled until it is there
    /// (<paramref name="present"/>) or gone, within <paramref name="timeout"/>; the last reading either way.
    /// </summary>
    public static ComInspectorInfo? PollForInspector(OutlookComSession session, string entryId, bool present, TimeSpan timeout)
    {
        LiveWaitBudget wait = LiveWaitBudget.Of(timeout);
        ComInspectorInfo? last = null;
        while (wait.HasTimeLeft)
        {
            last = session.GetOpenInspectors()
                .FirstOrDefault(i => i.EntryId != null && string.Equals(i.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
            if ((last != null) == present)
            {
                return last;
            }

            Thread.Sleep(500);
        }

        return last;
    }
}
