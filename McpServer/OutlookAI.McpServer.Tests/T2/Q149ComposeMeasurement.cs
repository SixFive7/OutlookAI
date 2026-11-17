using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The pure half of the Q149 (d) measurement: how OutlookAI should write draft text. Everything here is
/// names, strings and numbers in, strings and verdicts out - no COM, no window, no mailbox - so CI pins
/// every branch of it (<c>T1/Q149ComposeMeasurementTests</c>) while the live half
/// (<see cref="LiveComposeProbe"/> and the three <c>Live*ComposeTests</c> classes) can only run on a test
/// guest.
/// <para>
/// <b>The decision being measured</b> (chosen by the maintainer 2026-11-17, Q149 (d)): keep writing draft
/// text through Outlook's Word editor in a hidden compose window, with a lifetime pin while drafting (b);
/// or write it as HTML through the object model with no compose window, no activation and no pin ever
/// (c). (c) replaces Word only if its drafts measure as indistinguishable from today's, both opened in
/// Outlook and as received. The staged plan, the cells and the decision rules are in
/// <c>Docs/research/q149-draft-text/README.md</c>.
/// </para>
/// </summary>
public static class Q149Cells
{
    /// <summary>The xunit collection every Q149 class sits in.</summary>
    public const string Collection = LiveCollections.ComposeSurface;

    /// <summary>The prefix every attempt line carries in the test output, the console capture and the TRX.</summary>
    public const string LinePrefix = "Q149-ATTEMPT ";

    /// <summary>The prefix of the one summary line each cell writes.</summary>
    public const string SummaryPrefix = "Q149-SUMMARY ";

    // Stage 1 - a window-less Outlook, each trial a fresh COM cold start.

    /// <summary>The shipped compose (ReviseHeldDocument, reached through update_draft) with no lifetime pin. Decisive.</summary>
    public const string L1 = "S1-L1-shipped-compose-no-pin";

    /// <summary>The same compose with the pin the product makes - the control for L1.</summary>
    public const string L1Control = "S1-L1c-shipped-compose-with-pin";

    /// <summary>A fetched inspector on a re-opened draft, never activated, read once, closed (D49 row 1 on 2024).</summary>
    public const string L2 = "S1-L2-fetch-only-close";

    /// <summary>As L2, re-reading the editor every 250 ms for 30 s.</summary>
    public const string W2 = "S1-W2-wait-and-reread";

    /// <summary>As W2, the first read 180 s after the cold start.</summary>
    public const string W3 = "S1-W3-reread-after-180s";

    /// <summary>The inspector's own window shown by Windows, off-screen, without Activate.</summary>
    public const string W4 = "S1-W4-windows-show-inspector";

    /// <summary>The pin Explorer's window shown by Windows, off-screen, then the inspector read.</summary>
    public const string W5 = "S1-W5-windows-show-pin";

    /// <summary>The product promotion once per process, then later drafts read directly (and committed without Activate).</summary>
    public const string W6 = "S1-W6-promote-once-per-process";

    /// <summary>
    /// As W6, every later draft committed as B3 commits - ModifiedFormPages.Add("Message") before the editor is
    /// read, Close(olSave) after the edit - so one promotion per process is tried end to end, the commit included.
    /// </summary>
    public const string W7 = "S1-W7-promote-once-then-formpage-olsave";

    // Stage 2 - commits on re-opened drafts in a windowed Outlook.

    /// <summary>The product update_draft, with its Activate: the control.</summary>
    public const string B0 = "S2-B0-product-update";

    /// <summary>Editor read without Activate; CurrentItem.Save, Close(olDiscard).</summary>
    public const string B1 = "S2-B1-save-currentitem";

    /// <summary>Editor read without Activate; Close(olSave).</summary>
    public const string B2 = "S2-B2-close-olsave";

    /// <summary>ModifiedFormPages.Add("Message") first; Close(olSave).</summary>
    public const string B3 = "S2-B3-formpages-close-olsave";

    /// <summary>Word Document.Save, then Close(olDiscard).</summary>
    public const string B4 = "S2-B4-document-save";

    // Stage 3 - the HTML-only draft (c) against today's.

    /// <summary>A new mail: product draft against the HTML-only draft.</summary>
    public const string C2 = "S3-C2-new-mail-parity";

    /// <summary>Reply and Forward without any inspector: what the object model hands over.</summary>
    public const string C4 = "S3-C4-reply-forward-parity";

    /// <summary>update_draft across the two kinds of draft.</summary>
    public const string C3 = "S3-C3-update-parity";

    /// <summary>Both drafts sent through the guest's loopback sink and compared as received.</summary>
    public const string C5 = "S3-C5-received-parity";

    /// <summary>What a stage-3 item holds with no inspector ever: the stylesheet question.</summary>
    public const string C1 = "S3-C1-no-inspector-template";
}

/// <summary>
/// One attempt, written as one JSON line. Keys are kept sorted so two runs of the same cell diff cleanly.
/// Content-free by construction: callers put counts, offsets, hashes, booleans, error types and HRESULTs
/// here - and their own run tokens - never a subject or a body (S4).
/// </summary>
public sealed class Q149Attempt
{
    private readonly SortedDictionary<string, object?> _fields = new(StringComparer.Ordinal);

    public Q149Attempt(string cell, int iteration)
    {
        if (string.IsNullOrWhiteSpace(cell))
        {
            throw new ArgumentException("A cell is required.", nameof(cell));
        }

        _fields["cell"] = cell;
        _fields["iteration"] = iteration;
    }

    public string Cell => (string)_fields["cell"]!;

    /// <summary>Sets a field; a null value removes it, so an unmeasured field is absent rather than null.</summary>
    public Q149Attempt Set(string key, object? value)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            throw new ArgumentException("A key is required.", nameof(key));
        }

        if (value == null)
        {
            _fields.Remove(key);
        }
        else
        {
            _fields[key] = value;
        }

        return this;
    }

    public object? Get(string key) => _fields.TryGetValue(key, out object? value) ? value : null;

    public bool Has(string key) => _fields.ContainsKey(key);

    /// <summary>The line as written: the prefix and one JSON object, no line break inside it.</summary>
    public string ToLine()
    {
        return Q149Cells.LinePrefix + JsonSerializer.Serialize(_fields);
    }

    /// <summary>Reads a line written by <see cref="ToLine"/> back into its fields - for the summaries and T1.</summary>
    public static IReadOnlyDictionary<string, JsonElement>? TryParse(string line)
    {
        if (line == null)
        {
            return null;
        }

        int at = line.IndexOf(Q149Cells.LinePrefix, StringComparison.Ordinal);
        if (at < 0)
        {
            return null;
        }

        try
        {
            return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(line.Substring(at + Q149Cells.LinePrefix.Length));
        }
        catch (JsonException)
        {
            return null;
        }
    }
}

/// <summary>What happened to OUTLOOK.EXE after a trial's treatment.</summary>
public enum Q149Lifetime
{
    /// <summary>Not observed, or the observation itself failed.</summary>
    Unknown = 0,

    /// <summary>Still running at the end of the watch, and no Quit event.</summary>
    Survived = 1,

    /// <summary>Still running, but Outlook raised Quit: the documented park.</summary>
    Parked = 2,

    /// <summary>Exited after raising Quit.</summary>
    OrderlyExit = 3,

    /// <summary>Exited with exit code 0 and no Quit event seen.</summary>
    SilentExit = 4,

    /// <summary>Exited with an exception code (0xC0000000 and up).</summary>
    Crash = 5,
}

/// <summary>The lifetime classification, and the numbers it is read from.</summary>
public static class Q149LifetimeRules
{
    /// <summary>
    /// The classes the pins plan names: survived, orderly exit, silent exit, crash, parked. A NTSTATUS
    /// exception code reaches .NET's <c>Process.ExitCode</c> as a negative int.
    /// </summary>
    public static Q149Lifetime Classify(bool aliveAtEnd, int? exitCode, bool quitEventSeen)
    {
        if (aliveAtEnd)
        {
            return quitEventSeen ? Q149Lifetime.Parked : Q149Lifetime.Survived;
        }

        if (exitCode == null)
        {
            return Q149Lifetime.Unknown;
        }

        if (IsExceptionCode(exitCode.Value))
        {
            return Q149Lifetime.Crash;
        }

        return quitEventSeen ? Q149Lifetime.OrderlyExit : Q149Lifetime.SilentExit;
    }

    /// <summary>True for an exception exit code (STATUS_* with severity error, 0xC0000000 and up).</summary>
    public static bool IsExceptionCode(int exitCode)
    {
        return unchecked((uint)exitCode) >= 0xC0000000u;
    }

    /// <summary>True for every class in which OUTLOOK.EXE is gone.</summary>
    public static bool Exited(Q149Lifetime lifetime)
    {
        return lifetime is Q149Lifetime.OrderlyExit or Q149Lifetime.SilentExit or Q149Lifetime.Crash;
    }

    /// <summary>
    /// The presence samples the pins plan asks for, as offsets in milliseconds from the moment the
    /// treatment ended: every 50 ms to 2 s, then 5, 10, 30 and 60 s, cut at <paramref name="watchMs"/>.
    /// </summary>
    public static IReadOnlyList<int> SampleOffsets(int watchMs)
    {
        List<int> offsets = new();
        for (int ms = 0; ms <= 2000 && ms <= watchMs; ms += 50)
        {
            offsets.Add(ms);
        }

        foreach (int ms in new[] { 5000, 10000, 30000, 60000 })
        {
            if (ms <= watchMs)
            {
                offsets.Add(ms);
            }
        }

        return offsets;
    }

    /// <summary>
    /// Whether the process was alive at each offset, from the one fact that decides it: when it exited
    /// (null = it never did inside the watch). Derived rather than polled, because the exit time comes
    /// from the process handle, which is exact; a poll can only be later.
    /// </summary>
    public static IReadOnlyList<int[]> Samples(IReadOnlyList<int> offsets, long? exitedAtMs)
    {
        List<int[]> samples = new(offsets.Count);
        foreach (int offset in offsets)
        {
            bool alive = exitedAtMs == null || offset < exitedAtMs.Value;
            samples.Add(new[] { offset, alive ? 1 : 0 });
        }

        return samples;
    }
}

/// <summary>The decision rules of the two plans, merged, applied to tallies. Each verdict says which rule fired.</summary>
public static class Q149Decisions
{
    /// <summary>Rule D1 (the pins plan's): what the decisive cell L1 says about the pin.</summary>
    public static string Lifetime(int exits, int trials)
    {
        if (trials <= 0)
        {
            return "not measured";
        }

        if (exits * 10 >= trials * 9)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "D1: the window-less Outlook exited in {0} of {1} - the pin is required for as long as drafts go through an inspector",
                exits,
                trials);
        }

        if ((trials - exits) * 10 >= trials * 9)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "D1: the window-less Outlook survived in {0} of {1} - the pin is unnecessary for the shipped compose",
                trials - exits,
                trials);
        }

        return string.Format(
            CultureInfo.InvariantCulture,
            "D1: the window-less Outlook exited in {0} of {1} - timing-dependent, keep the pin",
            exits,
            trials);
    }

    /// <summary>
    /// Rule D2 (the Word-editor plan's phase A): how the editor can be had with no window. A cell "works"
    /// at 9 of 10 or better; the wait cells must also have shown no window at all.
    /// </summary>
    public static string Availability(
        (int Ok, int N, int Shown) wait,
        (int Ok, int N, int Shown) waitLate,
        (int Ok, int N) showInspector,
        (int Ok, int N) showPin,
        (int Ok, int N) afterOnePromotion)
    {
        static bool Works(int ok, int n) => n > 0 && ok * 10 >= n * 9;

        if ((Works(wait.Ok, wait.N) && wait.Shown == 0) || (Works(waitLate.Ok, waitLate.N) && waitLate.Shown == 0))
        {
            return "D2: a bounded re-read gives the editor with no window - replace the window with the re-read";
        }

        if (Works(showInspector.Ok, showInspector.N) || Works(showPin.Ok, showPin.N))
        {
            return "D2: a Windows-level show without Activate gives the editor - one such show per process, no Outlook Activate";
        }

        if (Works(afterOnePromotion.Ok, afterOnePromotion.N))
        {
            return "D2: only the product's promotion gives the editor, and it lasts for the process - promote once per process, not per compose";
        }

        return "D2: nothing short of the per-compose promotion gives the editor - the current design is justified on this Office, and only the HTML route (c) is window-free";
    }

    /// <summary>Rule D3 (phase B): a commit variant that works on every draft with no side effect lets a windowed Outlook skip the per-compose Activate.</summary>
    public static string Commit(IReadOnlyList<(string Variant, int Committed, int N, int SideEffects)> variants)
    {
        if (variants == null || variants.Count == 0)
        {
            return "not measured";
        }

        List<string> clean = variants
            .Where(v => v.N > 0 && v.Committed == v.N && v.SideEffects == 0)
            .Select(v => v.Variant)
            .ToList();
        return clean.Count > 0
            ? "D3: " + string.Join(", ", clean) + " committed every draft with no Activate and no side effect - a windowed Outlook can drop the per-compose Activate"
            : "D3: no variant committed every draft without Activate - keep the Activate";
    }

    /// <summary>Rule D4: (c) replaces Word only when every compared aspect is indistinguishable.</summary>
    public static string Parity(IReadOnlyList<(string Aspect, bool Same)> aspects)
    {
        if (aspects == null || aspects.Count == 0)
        {
            return "not measured";
        }

        List<string> different = aspects.Where(a => !a.Same).Select(a => a.Aspect).ToList();
        return different.Count == 0
            ? "D4: the HTML-only draft is indistinguishable on every aspect measured - (c) can replace Word"
            : "D4: the HTML-only draft differs in " + string.Join(", ", different) + " - (b)";
    }
}

/// <summary>Whether the Q149 measurement runs in this test run at all.</summary>
public static class Q149RunGate
{
    /// <summary>The line a test writes when the gate holds it back.</summary>
    public const string NotRunLine =
        "NOT RUN: the Q149 measurement runs only in a run that selects it alone - it bounces Outlook dozens of times "
        + "and takes over an hour - for example Invoke-LiveTierOnGuest.ps1 -FilterSuffix "
        + "'&FullyQualifiedName~OutlookAI.McpServer.Tests.T2.LiveQ149'.";

    /// <summary>
    /// True only when the published run plan holds this measurement's collection and nothing else. A
    /// whole-tier run must not pay for an hour of Outlook restarts, nor have its Outlook bounced under the
    /// other collections; an unpublished plan answers false for the same reason.
    /// </summary>
    public static bool ShouldRun(IReadOnlyList<string>? orderedCollections)
    {
        return orderedCollections != null
            && orderedCollections.Count > 0
            && orderedCollections.All(c => string.Equals(c, Q149Cells.Collection, StringComparison.Ordinal));
    }
}

/// <summary>
/// The refusals every Q149 write passes through before Outlook is asked anything (mailbox rules 1-3). Pure,
/// so T1 proves each one. The store itself is the allowlist's business (<see cref="LiveStoreWriteGuard"/>);
/// these are the item-level halves.
/// </summary>
public static class Q149ProbeGuards
{
    /// <summary>Null when <paramref name="subject"/> carries the test tag AND this run's marker, ordinally; else why not.</summary>
    public static string? RefuseSubject(string? subject, string? runMarker)
    {
        if (string.IsNullOrWhiteSpace(runMarker) || runMarker!.Length < 12)
        {
            return "the run marker is missing or too weak for a safe match (S3)";
        }

        if (string.IsNullOrEmpty(subject))
        {
            return "the item has no subject, so it cannot be this run's test item";
        }

        if (!subject!.Contains(LiveOutlookTestMailer.SubjectTag, StringComparison.Ordinal))
        {
            return "the subject does not carry the test tag " + LiveOutlookTestMailer.SubjectTag;
        }

        if (!subject.Contains(runMarker, StringComparison.Ordinal))
        {
            return "the subject does not carry this run's marker";
        }

        return null;
    }

    /// <summary>Null when <paramref name="entryId"/> is one this run created and recorded; else why not (rule 2's allowlist).</summary>
    public static string? RefuseEntryId(string? entryId, IReadOnlyCollection<string> allowlist)
    {
        ArgumentNullException.ThrowIfNull(allowlist);
        if (string.IsNullOrWhiteSpace(entryId))
        {
            return "no EntryID";
        }

        return allowlist.Contains(entryId!, StringComparer.OrdinalIgnoreCase)
            ? null
            : "the EntryID is not one this run created";
    }

    /// <summary>Null when every recipient is the hub itself (Q110: nothing is addressed to anyone else); else why not.</summary>
    public static string? RefuseRecipients(IReadOnlyList<string>? recipients, string hubAddress)
    {
        if (string.IsNullOrWhiteSpace(hubAddress))
        {
            return "the hub's own address is unknown";
        }

        if (recipients == null || recipients.Count == 0)
        {
            return "the item has no recipient";
        }

        foreach (string recipient in recipients)
        {
            if (!string.Equals(recipient?.Trim(), hubAddress.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return "a recipient other than the hub itself";
            }
        }

        return null;
    }
}

/// <summary>What one draft's raw HTML holds - offsets, counts and hashes, never text (S4).</summary>
public sealed record Q149HtmlFacts(
    int Length,
    int WordSectionAt,
    int MsoNormal,
    int StyleBlocks,
    string? StylesheetHash,
    int MailAutoSigAt,
    int MailOriginalAt,
    int HeaderBlockAt,
    int Images,
    int CidImages,
    int FileLinks,
    int VmlShapes,
    IReadOnlyDictionary<string, int> TokenAt)
{
    /// <summary>The fields as a JSON-friendly dictionary for an attempt line.</summary>
    public IReadOnlyDictionary<string, object> ToFields()
    {
        Dictionary<string, object> fields = new(StringComparer.Ordinal)
        {
            ["len"] = Length,
            ["wordSection1At"] = WordSectionAt,
            ["msoNormal"] = MsoNormal,
            ["styleBlocks"] = StyleBlocks,
            ["mailAutoSigAt"] = MailAutoSigAt,
            ["mailOriginalAt"] = MailOriginalAt,
            ["headerBlockAt"] = HeaderBlockAt,
            ["img"] = Images,
            ["cidImg"] = CidImages,
            ["fileLinks"] = FileLinks,
            ["vml"] = VmlShapes,
        };
        if (StylesheetHash != null)
        {
            fields["stylesheetHash"] = StylesheetHash;
        }

        foreach (KeyValuePair<string, int> token in TokenAt)
        {
            fields["at_" + token.Key] = token.Value;
        }

        return fields;
    }
}

/// <summary>Reads <see cref="Q149HtmlFacts"/> out of a draft's HTMLBody, and the stylesheet that decides its look.</summary>
public static class Q149Html
{
    private static readonly Regex StyleBlock = new(@"<style[^>]*>(?<css>.*?)</style>", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex ImgTag = new(@"<img\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex CidSrc = new(@"<img\b[^>]*\bsrc\s*=\s*[""']?cid:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex Whitespace = new(@"\s+", RegexOptions.CultureInvariant);

    /// <summary>
    /// Where Outlook's quoted original starts in a reply or forward: the bookmark anchor when Word wrote
    /// one, else the separator block Outlook puts above the original's header.
    /// </summary>
    private static readonly string[] HeaderBlockMarkers =
    {
        "id=\"divRplyFwdMsg\"",
        "id=divRplyFwdMsg",
        "border:none;border-top:solid #E1E1E1",
        "border:none;border-top:solid #B5C4DF",
    };

    public static Q149HtmlFacts Analyze(string? html, IReadOnlyDictionary<string, string>? tokens = null)
    {
        string text = html ?? string.Empty;
        Dictionary<string, int> at = new(StringComparer.Ordinal);
        if (tokens != null)
        {
            foreach (KeyValuePair<string, string> token in tokens)
            {
                at[token.Key] = string.IsNullOrEmpty(token.Value) ? -1 : text.IndexOf(token.Value, StringComparison.Ordinal);
            }
        }

        string? stylesheet = ExtractStylesheet(text);
        return new Q149HtmlFacts(
            text.Length,
            text.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase),
            CountOrdinalIgnoreCase(text, "class=MsoNormal") + CountOrdinalIgnoreCase(text, "class=\"MsoNormal\""),
            StyleBlock.Matches(text).Count,
            stylesheet == null ? null : Hash(stylesheet),
            IndexOfAnchor(text, "_MailAutoSig"),
            IndexOfAnchor(text, "_MailOriginal"),
            HeaderBlockAt(text),
            ImgTag.Matches(text).Count,
            CidSrc.Matches(text).Count,
            CountOrdinalIgnoreCase(text, "file:///"),
            CountOrdinalIgnoreCase(text, "<v:shape") + CountOrdinalIgnoreCase(text, "<v:rect") + CountOrdinalIgnoreCase(text, "<v:imagedata"),
            at);
    }

    /// <summary>
    /// Every &lt;style&gt; block's content, comment markers dropped and whitespace collapsed - the part of
    /// the stylesheet that decides how a reader renders the message. Null when there is none.
    /// </summary>
    public static string? ExtractStylesheet(string? html)
    {
        if (string.IsNullOrEmpty(html))
        {
            return null;
        }

        MatchCollection blocks = StyleBlock.Matches(html);
        if (blocks.Count == 0)
        {
            return null;
        }

        StringBuilder css = new();
        foreach (Match block in blocks)
        {
            string content = block.Groups["css"].Value.Replace("<!--", " ", StringComparison.Ordinal).Replace("-->", " ", StringComparison.Ordinal);
            css.Append(Whitespace.Replace(content, " ").Trim()).Append('\n');
        }

        return css.ToString().Trim();
    }

    /// <summary>The first 16 hex digits of a SHA-256 - enough to say two stylesheets are or are not the same.</summary>
    public static string Hash(string text)
    {
        byte[] digest = SHA256.HashData(Encoding.UTF8.GetBytes(text ?? string.Empty));
        return Convert.ToHexString(digest, 0, 8);
    }

    /// <summary>The offset of a bookmark anchor - <c>name="X"</c>, <c>name=X</c> or <c>id="X"</c> - or -1.</summary>
    public static int IndexOfAnchor(string html, string name)
    {
        int best = -1;
        foreach (string shape in new[] { "name=\"" + name + "\"", "name=" + name + ">", "name=" + name + " ", "name='" + name + "'", "id=\"" + name + "\"" })
        {
            int at = html.IndexOf(shape, StringComparison.OrdinalIgnoreCase);
            if (at >= 0 && (best < 0 || at < best))
            {
                best = at;
            }
        }

        return best;
    }

    /// <summary>Where the quoted original's header block starts, or -1.</summary>
    public static int HeaderBlockAt(string html)
    {
        int best = -1;
        foreach (string marker in HeaderBlockMarkers)
        {
            int at = html.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (at >= 0 && (best < 0 || at < best))
            {
                best = at;
            }
        }

        return best;
    }

    private static readonly Regex OutlineTag = new(@"<(?<close>/?)(?<tag>p|div|h[1-6]|table|ul|ol|li|a|img|hr|span|o:p)\b(?<attrs>[^>]*)>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex ClassAttr = new(@"\bclass\s*=\s*[""']?(?<c>[A-Za-z0-9_-]+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex NameAttr = new(@"\b(?:name|id)\s*=\s*[""']?(?<n>_Mail[A-Za-z]+|divRplyFwdMsg|WordSection1)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>
    /// The shape of a message body, content-free: from the start of <c>WordSection1</c> (or the body), every
    /// block, anchor and picture as a short token - <c>p.MsoNormal</c>, <c>div[border-top]</c>,
    /// <c>a#_MailAutoSig</c>, <c>img[cid]</c>, <c>h1</c> - with <c>{name}</c> where one of
    /// <paramref name="tokens"/> occurs. Only tag names, Word's class names and these marks; never text.
    /// </summary>
    public static string Outline(string? html, IReadOnlyDictionary<string, string>? tokens = null, int maxItems = 80)
    {
        if (string.IsNullOrEmpty(html))
        {
            return string.Empty;
        }

        int from = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        from = from >= 0
            ? Math.Max(0, html.LastIndexOf('<', from))
            : Math.Max(0, html.IndexOf("<body", StringComparison.OrdinalIgnoreCase));

        List<(int At, string Mark)> marks = new();
        foreach (Match m in OutlineTag.Matches(html, from))
        {
            string tag = m.Groups["tag"].Value.ToLowerInvariant();
            string attrs = m.Groups["attrs"].Value;
            if (m.Groups["close"].Value.Length > 0)
            {
                if (tag is "div" or "table" or "ul" or "ol")
                {
                    marks.Add((m.Index, "/" + tag));
                }

                continue;
            }

            if (tag is "span" or "o:p")
            {
                continue;
            }

            string mark = tag;
            Match c = ClassAttr.Match(attrs);
            if (c.Success)
            {
                mark += "." + c.Groups["c"].Value;
            }

            Match n = NameAttr.Match(attrs);
            if (n.Success)
            {
                mark += "#" + n.Groups["n"].Value;
            }
            else if (tag == "a")
            {
                continue;
            }

            if (attrs.Contains("border-top", StringComparison.OrdinalIgnoreCase))
            {
                mark += "[border-top]";
            }

            if (tag == "img")
            {
                mark += attrs.Contains("cid:", StringComparison.OrdinalIgnoreCase) ? "[cid]" : attrs.Contains("file:", StringComparison.OrdinalIgnoreCase) ? "[file]" : "[other]";
            }

            marks.Add((m.Index, mark));
        }

        if (tokens != null)
        {
            foreach (KeyValuePair<string, string> token in tokens)
            {
                if (string.IsNullOrEmpty(token.Value))
                {
                    continue;
                }

                for (int at = html.IndexOf(token.Value, from, StringComparison.Ordinal); at >= 0; at = html.IndexOf(token.Value, at + token.Value.Length, StringComparison.Ordinal))
                {
                    marks.Add((at, "{" + token.Key + "}"));
                }
            }
        }

        return string.Join(" ", marks.OrderBy(x => x.At).Take(maxItems).Select(x => x.Mark));
    }

    public static int CountOrdinalIgnoreCase(string text, string needle)
    {
        int count = 0;
        for (int at = text.IndexOf(needle, StringComparison.OrdinalIgnoreCase); at >= 0; at = text.IndexOf(needle, at + needle.Length, StringComparison.OrdinalIgnoreCase))
        {
            count++;
        }

        return count;
    }
}

/// <summary>
/// The HTML-only draft (c), prototyped so it can be measured - test code, never shipped. It does with
/// strings what Word does today with a document: Outlook's mail stylesheet, the body inside
/// <c>WordSection1</c> in <c>MsoNormal</c> paragraphs, the signature with its pictures as <c>cid:</c>
/// references and the <c>_MailAutoSig</c> marker, the text above a reply's quoted original, and
/// update_draft's replacement of everything above that boundary.
/// <para>
/// Its stylesheet comes from a reference: the <c>&lt;html&gt;</c> start tag, the whole <c>&lt;head&gt;</c>
/// and the <c>&lt;body&gt;</c> start tag of a draft Word wrote in the same profile. That is (c)'s best
/// case - the stylesheet byte for byte - so whatever still differs is what (c) would have to rebuild.
/// </para>
/// </summary>
public static class HtmlOnlyDraftComposer
{
    private static readonly Regex HtmlStart = new(@"<html\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex Head = new(@"<head\b[^>]*>.*?</head>", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex BodyStart = new(@"<body\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex BodyInner = new(@"<body\b[^>]*>(?<inner>.*)</body>", RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex WordSectionStart = new(@"<div\b[^>]*class\s*=\s*[""']?WordSection1[""']?[^>]*>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex ImgSrc = new(@"(<img\b[^>]*\bsrc\s*=\s*)(?<q>[""']?)(?<src>[^""'\s>]+)\k<q>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex ParagraphStart = new(@"<p\b(?<attrs>[^>]*)>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>The blank paragraph Word leaves between the body and the signature or the quote.</summary>
    public const string BlankParagraph = "<p class=MsoNormal><o:p>&nbsp;</o:p></p>";

    /// <summary>
    /// The start of the document a reference draft was written with: its <c>&lt;html&gt;</c> start tag
    /// (Office's namespaces), its <c>&lt;head&gt;</c> and its <c>&lt;body&gt;</c> start tag. Null when
    /// the reference lacks any of them.
    /// </summary>
    public static (string HtmlStart, string Head, string BodyStart)? ReadFrame(string? referenceHtml)
    {
        if (string.IsNullOrEmpty(referenceHtml))
        {
            return null;
        }

        Match html = HtmlStart.Match(referenceHtml);
        Match head = Head.Match(referenceHtml);
        Match body = BodyStart.Match(referenceHtml);
        if (!html.Success || !head.Success || !body.Success)
        {
            return null;
        }

        return (html.Value, head.Value, body.Value);
    }

    /// <summary>
    /// A plain-text body the way the Word path writes it today (<c>InsertBodyAboveSignature</c>): ONE
    /// paragraph, every newline a soft line break - Word's vertical tab, which it writes as &lt;br&gt; - so a
    /// blank line is two breaks in a row, not a paragraph of its own.
    /// </summary>
    public static string ParagraphsFromText(string text)
    {
        string normalized = (text ?? string.Empty).Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        StringBuilder html = new();
        html.Append("<p class=MsoNormal>");
        for (int i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                html.Append("<br>\n");
            }

            html.Append(WebUtility.HtmlEncode(lines[i]));
        }

        html.Append("<o:p></o:p></p>\n");
        return html.ToString();
    }

    /// <summary>
    /// An HTML fragment as (c) would place it: every bare <c>&lt;p&gt;</c> given Word's <c>MsoNormal</c>
    /// class, everything else - headings, lists, tables - left for the reader's own defaults.
    /// </summary>
    public static string ParagraphsFromHtml(string fragment)
    {
        return ParagraphsFromHtml(fragment, HtmlOnlyStyle.Message);
    }

    /// <summary>
    /// The fragment as the given <paramref name="style"/> places it: <see cref="HtmlOnlyStyle.Message"/> gives every
    /// bare paragraph the message's <c>MsoNormal</c>; <see cref="HtmlOnlyStyle.AsImported"/> leaves the markup as it
    /// came, the way Word's HTML import receives it when it inserts a file.
    /// </summary>
    public static string ParagraphsFromHtml(string fragment, HtmlOnlyStyle style)
    {
        if (style == HtmlOnlyStyle.AsImported)
        {
            return fragment ?? string.Empty;
        }

        return ParagraphStart.Replace(fragment ?? string.Empty, m =>
            m.Groups["attrs"].Value.Contains("class", StringComparison.OrdinalIgnoreCase)
                ? m.Value
                : "<p class=MsoNormal" + m.Groups["attrs"].Value + ">");
    }

    /// <summary>
    /// The signature as (c) would place it: the <c>&lt;body&gt;</c> content of its .htm, every picture's
    /// <c>src</c> re-pointed by <paramref name="srcToCid"/> (a <c>cid:</c> reference to an attachment the
    /// caller made, or null to leave it), and the whole of it marked <c>_MailAutoSig</c> the way Word
    /// marks its start.
    /// </summary>
    public static string SignatureBlock(string signatureHtm, Func<string, string?> srcToCid)
    {
        return SignatureBlock(signatureHtm, srcToCid, HtmlOnlyStyle.Message);
    }

    /// <summary><see cref="SignatureBlock(string, Func{string, string})"/> in the given <paramref name="style"/>.</summary>
    public static string SignatureBlock(string signatureHtm, Func<string, string?> srcToCid, HtmlOnlyStyle style)
    {
        ArgumentNullException.ThrowIfNull(srcToCid);
        Match inner = BodyInner.Match(signatureHtm ?? string.Empty);
        string content = inner.Success ? inner.Groups["inner"].Value : signatureHtm ?? string.Empty;
        content = ImgSrc.Replace(content, m =>
        {
            string? cid = srcToCid(m.Groups["src"].Value);
            return cid == null ? m.Value : m.Groups[1].Value + "\"cid:" + cid + "\"";
        });
        content = ParagraphsFromHtml(content, style);
        return "<a name=\"_MailAutoSig\"></a>" + content.Trim();
    }

    /// <summary>A new mail: the reference's frame, then <c>WordSection1</c> holding the body, a blank line and the signature.</summary>
    public static string ComposeNew((string HtmlStart, string Head, string BodyStart) frame, string bodyParagraphs, string? signatureBlock)
    {
        StringBuilder html = new();
        html.Append(frame.HtmlStart).Append('\n').Append(frame.Head).Append('\n').Append(frame.BodyStart).Append('\n');
        html.Append("<div class=WordSection1>\n");
        html.Append(bodyParagraphs);
        if (!string.IsNullOrEmpty(signatureBlock))
        {
            html.Append(BlankParagraph).Append('\n').Append(signatureBlock).Append('\n');
        }

        html.Append("</div>\n</body>\n</html>\n");
        return html.ToString();
    }

    /// <summary>
    /// A reply or forward Outlook made with no inspector, with the body placed at the top of its
    /// <c>WordSection1</c> - above the quoted original - followed by a blank line and, when given, the
    /// signature. Null when the item's HTML has no <c>WordSection1</c> to place it in.
    /// </summary>
    public static string? InsertAboveQuote(string? replyHtml, string bodyParagraphs, string? signatureBlock)
    {
        if (string.IsNullOrEmpty(replyHtml))
        {
            return null;
        }

        Match section = WordSectionStart.Match(replyHtml);
        if (!section.Success)
        {
            return null;
        }

        string insert = bodyParagraphs + BlankParagraph + "\n"
            + (string.IsNullOrEmpty(signatureBlock) ? string.Empty : signatureBlock + "\n" + BlankParagraph + "\n");
        int at = section.Index + section.Length;
        return replyHtml.Substring(0, at) + "\n" + insert + replyHtml.Substring(at);
    }

    /// <summary>
    /// update_draft without Word: everything between the start of <c>WordSection1</c> and the boundary -
    /// the <c>_MailAutoSig</c> anchor, else the <c>_MailOriginal</c> anchor, else the quoted original's
    /// header block - replaced by the new body and a blank line. With no boundary the whole section is the
    /// draft. The boundary is moved back to the start of the element that holds it, so a paragraph is never
    /// cut in half. Null when the HTML has no <c>WordSection1</c>.
    /// </summary>
    public static string? ReplaceDraftRegion(string? html, string newBodyParagraphs)
    {
        if (string.IsNullOrEmpty(html))
        {
            return null;
        }

        Match section = WordSectionStart.Match(html);
        if (!section.Success)
        {
            return null;
        }

        int start = section.Index + section.Length;
        int boundary = FirstBoundary(html, start);
        int end;
        if (boundary < 0)
        {
            end = html.LastIndexOf("</div>", StringComparison.OrdinalIgnoreCase);
            if (end < start)
            {
                return null;
            }

            return html.Substring(0, start) + "\n" + newBodyParagraphs + html.Substring(end);
        }

        end = StartOfEnclosingElement(html, start, boundary);
        return html.Substring(0, start) + "\n" + newBodyParagraphs + BlankParagraph + "\n" + html.Substring(end);
    }

    private static int FirstBoundary(string html, int from)
    {
        foreach (string name in new[] { "_MailAutoSig", "_MailOriginal" })
        {
            int at = Q149Html.IndexOfAnchor(html, name);
            if (at >= from)
            {
                return at;
            }
        }

        int header = Q149Html.HeaderBlockAt(html);
        return header >= from ? header : -1;
    }

    /// <summary>The start of the innermost &lt;p&gt; or &lt;div&gt; that opens between <paramref name="floor"/> and <paramref name="at"/> and holds it.</summary>
    private static int StartOfEnclosingElement(string html, int floor, int at)
    {
        int tagStart = html.LastIndexOf('<', Math.Max(0, at - 1));
        while (tagStart >= floor)
        {
            string head = html.Substring(tagStart, Math.Min(5, html.Length - tagStart)).ToLowerInvariant();
            if (head.StartsWith("<p ", StringComparison.Ordinal) || head.StartsWith("<p>", StringComparison.Ordinal) || head.StartsWith("<div", StringComparison.Ordinal))
            {
                return tagStart;
            }

            if (head.StartsWith("</p", StringComparison.Ordinal) || head.StartsWith("</div", StringComparison.Ordinal))
            {
                break;
            }

            tagStart = tagStart == 0 ? -1 : html.LastIndexOf('<', tagStart - 1);
        }

        // The boundary is not inside an element that opened after the floor: cut at the tag that holds it.
        int own = html.LastIndexOf('<', at);
        return own >= floor ? own : at;
    }
}

/// <summary>How the HTML-only draft styles what it places.</summary>
public enum HtmlOnlyStyle
{
    /// <summary>Bare paragraphs given the message's <c>MsoNormal</c> class, so they take the message's own style.</summary>
    Message = 0,

    /// <summary>The markup left as it came, as Word's HTML import receives it from an inserted file.</summary>
    AsImported = 1,
}

/// <summary>One paragraph as Word lays it out when Outlook opens the item - what a reader sees, not the markup.</summary>
public sealed record WordParagraphLook(
    string Text,
    int TopBorder,
    string Style,
    string Font,
    double Size,
    int Bold,
    int Italic,
    int Color,
    int Alignment,
    double SpaceBefore,
    double SpaceAfter,
    double LineSpacing,
    double LeftIndent,
    int InlineShapes,
    double ShapeWidth,
    double ShapeHeight)
{
    /// <summary>One line describing the paragraph, the text reduced to its length and a hash (S4).</summary>
    public string Describe()
    {
        return string.Format(
            CultureInfo.InvariantCulture,
            "len={0} h={1} border={16} style='{2}' font='{3}' {4}pt b={5} i={6} color={7} align={8} before={9} after={10} line={11} indent={12} shapes={13} {14}x{15}",
            Text.Length,
            Q149Html.Hash(Text),
            Style,
            Font,
            Size,
            Bold,
            Italic,
            Color,
            Alignment,
            SpaceBefore,
            SpaceAfter,
            LineSpacing,
            LeftIndent,
            InlineShapes,
            ShapeWidth,
            ShapeHeight,
            TopBorder);
    }
}

/// <summary>Compares two items as Word lays them out - the "how it looks when opened in Outlook" comparison.</summary>
public static class Q149WordLook
{
    /// <summary>
    /// Text as a reader sees it: Word's paragraph, cell and line marks, non-breaking spaces and runs of
    /// whitespace all reduced to one space, trimmed.
    /// </summary>
    public static string NormalizeText(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        StringBuilder normalized = new(text.Length);
        bool space = false;
        foreach (char c in text)
        {
            bool isSpace = char.IsWhiteSpace(c) || c == ' ' || c == '\a' || c == '\v' || c == '\r' || c == '\u000B' || c == '\u0007';
            if (isSpace)
            {
                if (!space && normalized.Length > 0)
                {
                    normalized.Append(' ');
                }

                space = true;
                continue;
            }

            if (c == '\u0001' || c == '\u0013' || c == '\u0014' || c == '\u0015')
            {
                // Word's field and object marks carry no text of their own.
                continue;
            }

            normalized.Append(c);
            space = false;
        }

        return normalized.ToString().Trim();
    }

    /// <summary>
    /// Every difference between two layouts, paragraph by paragraph, empty paragraphs included (they are
    /// the visible blank lines). At most <paramref name="max"/> are listed; the count is always exact.
    /// </summary>
    public static (int Count, IReadOnlyList<string> Listed) Compare(
        IReadOnlyList<WordParagraphLook> expected,
        IReadOnlyList<WordParagraphLook> actual,
        int max = 12)
    {
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(actual);
        List<string> listed = new();
        int count = 0;
        void Note(string difference)
        {
            count++;
            if (listed.Count < max)
            {
                listed.Add(difference);
            }
        }

        if (expected.Count != actual.Count)
        {
            Note(string.Format(CultureInfo.InvariantCulture, "paragraphs {0} vs {1}", expected.Count, actual.Count));
        }

        int shared = Math.Min(expected.Count, actual.Count);
        for (int i = 0; i < shared; i++)
        {
            WordParagraphLook e = expected[i];
            WordParagraphLook a = actual[i];
            string where = "p" + (i + 1).ToString(CultureInfo.InvariantCulture) + " ";
            if (!string.Equals(e.Text, a.Text, StringComparison.Ordinal))
            {
                Note(where + "text " + Q149Html.Hash(e.Text) + "/" + e.Text.Length + " vs " + Q149Html.Hash(a.Text) + "/" + a.Text.Length);
            }

            Field(where + "topBorder", e.TopBorder, a.TopBorder);
            Field(where + "style", e.Style, a.Style);
            Field(where + "font", e.Font, a.Font);
            Number(where + "size", e.Size, a.Size);
            Field(where + "bold", e.Bold, a.Bold);
            Field(where + "italic", e.Italic, a.Italic);
            Field(where + "color", e.Color, a.Color);
            Field(where + "align", e.Alignment, a.Alignment);
            Number(where + "before", e.SpaceBefore, a.SpaceBefore);
            Number(where + "after", e.SpaceAfter, a.SpaceAfter);
            Number(where + "line", e.LineSpacing, a.LineSpacing);
            Number(where + "indent", e.LeftIndent, a.LeftIndent);
            Field(where + "shapes", e.InlineShapes, a.InlineShapes);
            Number(where + "shapeW", e.ShapeWidth, a.ShapeWidth);
            Number(where + "shapeH", e.ShapeHeight, a.ShapeHeight);
        }

        return (count, listed);

        void Field<T>(string what, T e, T a)
        {
            if (!EqualityComparer<T>.Default.Equals(e, a))
            {
                Note(string.Format(CultureInfo.InvariantCulture, "{0} {1} vs {2}", what, e, a));
            }
        }

        void Number(string what, double e, double a)
        {
            if (Math.Abs(e - a) > 0.05)
            {
                Note(string.Format(CultureInfo.InvariantCulture, "{0} {1} vs {2}", what, e, a));
            }
        }
    }
}
