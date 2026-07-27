using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OutlookAI.Core.Text
{
    /// <summary>
    /// Result of normalizing an agent-supplied HTML body fragment.
    /// </summary>
    public sealed class HtmlNormalizationResult
    {
        internal HtmlNormalizationResult(string html, IReadOnlyList<string> adjustments, bool hasVisibleContent)
        {
            Html = html;
            Adjustments = adjustments;
            HasVisibleContent = hasVisibleContent;
        }

        /// <summary>The normalized, safe-to-inject HTML fragment (no html/head/body wrapper).</summary>
        public string Html { get; }

        /// <summary>
        /// Human-readable list of everything that was changed, so the tool result can tell the
        /// calling agent what happened to its markup instead of silently altering it.
        /// Capped; the last entry says how many further adjustments were omitted.
        /// </summary>
        public IReadOnlyList<string> Adjustments { get; }

        /// <summary>False when nothing renderable survived (empty input, or only dropped elements).</summary>
        public bool HasVisibleContent { get; }
    }

    /// <summary>
    /// Whitelist normalizer for the <c>body_html</c> draft argument (v3.MD D45).
    /// <para>
    /// Agent-authored HTML is injected into a real mail draft, so it is never trusted verbatim:
    /// a single unclosed tag from a model must not swallow the signature or the quoted thread,
    /// and nothing in it may load remote content or forge Outlook's own region markers.
    /// </para>
    /// <para>
    /// Policy (documented on the tool surface, pinned in T1):
    /// <list type="bullet">
    /// <item>allow-listed elements are kept; UNKNOWN/unsupported elements are UNWRAPPED - the tag
    /// is dropped and its text is kept, so no content is ever lost silently;</item>
    /// <item><c>script</c>/<c>style</c>/<c>iframe</c>/<c>object</c>/<c>embed</c>/<c>link</c>/<c>meta</c>
    /// and friends are dropped WITH their contents (they carry code or remote loads, not prose);</item>
    /// <item><c>img</c> is dropped - an inline image means a remote or attached resource, which this
    /// path does not own (the account signature keeps its own images);</item>
    /// <item>inline <c>style</c> attributes are KEPT (formatting is the point) minus declarations that
    /// fetch or execute; <c>id</c>/<c>name</c>/<c>class</c> and every <c>on*</c> handler are dropped -
    /// <c>id</c>/<c>name</c> because a forged <c>_MailAutoSig</c>/<c>_MailOriginal</c> anchor would
    /// corrupt the draft/signature/quote split;</item>
    /// <item><c>href</c> keeps only http/https/mailto/tel;</item>
    /// <item>malformed markup is REPAIRED, never rejected: stray <c>&lt;</c> is escaped, unclosed tags
    /// are closed at the end, mis-nested tags are unwound, stray end tags are dropped, and a stray
    /// <c>li</c>/<c>tr</c>/<c>td</c> gets the missing ancestors it needs.</item>
    /// </list>
    /// </para>
    /// Pure logic, no dependencies, no regex backtracking - a single forward scan.
    /// </summary>
    public static class HtmlFragmentNormalizer
    {
        /// <summary>Largest accepted <c>body_html</c> input; above this the caller is rejected pre-COM.</summary>
        public const int MaxInputChars = 512 * 1024;

        /// <summary>How many distinct adjustments are reported before the list is summarized.</summary>
        public const int MaxReportedAdjustments = 20;

        private const int MaxStyleChars = 1024;

        // Kept as-is (structure + formatting an agent needs for a formal letter).
        private static readonly HashSet<string> Allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "h1", "h2", "h3", "h4", "h5", "h6",
            "p", "div", "span", "br", "hr",
            "strong", "b", "em", "i", "u", "s", "strike", "del", "ins", "sub", "sup", "small",
            "code", "pre", "blockquote",
            "ol", "ul", "li", "dl", "dt", "dd",
            "table", "thead", "tbody", "tfoot", "tr", "th", "td", "caption", "colgroup", "col",
            "a",
        };

        // Dropped together with everything inside them.
        private static readonly HashSet<string> DropWithContent = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "script", "style", "iframe", "object", "embed", "applet", "frame", "frameset",
            "noframes", "noscript", "template", "svg", "math", "link", "meta", "base",
            "head", "title", "map", "area", "audio", "video", "source", "track", "canvas",
        };

        // Dropped, contents kept (also the fate of any unknown element).
        private static readonly HashSet<string> DropElementOnly = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "img", "input", "button", "select", "option", "textarea", "picture", "wbr",
        };

        private static readonly HashSet<string> VoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "br", "hr", "col",
        };

        private static readonly HashSet<string> GlobalAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "style", "title", "dir", "lang",
        };

        private static readonly HashSet<string> TableAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "align", "valign", "width", "height", "bgcolor", "border", "cellpadding", "cellspacing",
            "colspan", "rowspan", "nowrap", "span",
        };

        private static readonly HashSet<string> TableElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "table", "thead", "tbody", "tfoot", "tr", "th", "td", "caption", "colgroup", "col",
        };

        private static readonly HashSet<string> ListElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ol", "ul", "dl",
        };

        // Block-level elements that implicitly close an open <p>.
        private static readonly HashSet<string> ClosesOpenParagraph = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "hr", "pre", "blockquote",
            "ol", "ul", "dl", "li", "dt", "dd", "table", "thead", "tbody", "tfoot", "tr", "th", "td", "caption",
        };

        private static readonly string[] SafeUrlSchemes = { "http:", "https:", "mailto:", "tel:" };

        private static readonly string[] DangerousStyleTokens =
        {
            "url(", "expression(", "javascript:", "vbscript:", "behavior", "-moz-binding", "@import", "position:fixed",
        };

        /// <summary>
        /// Normalizes an agent-supplied HTML fragment into safe, well-formed HTML that can be
        /// injected into the draft region. Never throws on malformed input - it repairs.
        /// </summary>
        /// <exception cref="ArgumentException">Input longer than <see cref="MaxInputChars"/>.</exception>
        public static HtmlNormalizationResult Normalize(string? html)
        {
            if (html == null) html = string.Empty;
            if (html.Length > MaxInputChars)
            {
                throw new ArgumentException(
                    "body_html is " + html.Length.ToString(CultureInfo.InvariantCulture) +
                    " characters; the maximum is " + MaxInputChars.ToString(CultureInfo.InvariantCulture) +
                    ". Shorten the message or attach the long content instead.",
                    nameof(html));
            }

            var notes = new AdjustmentLog();
            var output = new StringBuilder(html.Length + 64);
            var open = new List<string>();
            bool sawVisible = false;

            int i = 0;
            while (i < html.Length)
            {
                char c = html[i];
                if (c != '<')
                {
                    int next = html.IndexOf('<', i);
                    if (next < 0) next = html.Length;
                    string text = html.Substring(i, next - i);
                    if (AppendText(text, output, open, notes)) sawVisible = true;
                    i = next;
                    continue;
                }

                if (StartsWith(html, i, "<!--"))
                {
                    int close = html.IndexOf("-->", i + 4, StringComparison.Ordinal);
                    i = close < 0 ? html.Length : close + 3;
                    notes.Add("removed an HTML comment");
                    continue;
                }

                if (StartsWith(html, i, "<!") || StartsWith(html, i, "<?"))
                {
                    int close = html.IndexOf('>', i);
                    i = close < 0 ? html.Length : close + 1;
                    notes.Add("removed a document-type or processing declaration");
                    continue;
                }

                bool isEnd = StartsWith(html, i, "</");
                int nameStart = i + (isEnd ? 2 : 1);
                if (nameStart >= html.Length || !IsNameStart(html[nameStart]))
                {
                    // Not a tag at all - a bare "<" in prose (e.g. "5 < 6"). Escape it.
                    output.Append("&lt;");
                    sawVisible = true;
                    notes.Add("escaped a stray \"<\"");
                    i++;
                    continue;
                }

                int nameEnd = nameStart;
                while (nameEnd < html.Length && IsNameChar(html[nameEnd])) nameEnd++;
                string name = html.Substring(nameStart, nameEnd - nameStart).ToLowerInvariant();

                int tagEnd = FindTagEnd(html, nameEnd);
                string rawAttributes = html.Substring(nameEnd, Math.Max(0, tagEnd - nameEnd));
                int afterTag = tagEnd < html.Length ? tagEnd + 1 : html.Length;

                if (isEnd)
                {
                    CloseTag(name, output, open, notes);
                    i = afterTag;
                    continue;
                }

                if (DropWithContent.Contains(name))
                {
                    i = SkipElement(html, afterTag, name);
                    notes.Add("removed <" + name + "> and its contents");
                    continue;
                }

                if (DropElementOnly.Contains(name))
                {
                    notes.Add(name == "img"
                        ? "removed <img> (images are not accepted in body_html)"
                        : "removed <" + name + "> (its text, if any, was kept)");
                    i = afterTag;
                    continue;
                }

                if (!Allowed.Contains(name))
                {
                    if (name == "html" || name == "body")
                        notes.Add("removed the <" + name + "> wrapper");
                    else
                        notes.Add("unwrapped unsupported <" + name + "> (its text was kept)");
                    i = afterTag;
                    continue;
                }

                bool selfClosed = rawAttributes.EndsWith("/", StringComparison.Ordinal);
                ApplyImplicitCloses(name, output, open, notes);
                EnsureAncestors(name, output, open, notes);

                output.Append('<').Append(name);
                AppendAttributes(name, rawAttributes, output, notes);
                output.Append('>');

                if (VoidElements.Contains(name))
                {
                    if (name == "hr") sawVisible = true;
                }
                else if (!selfClosed)
                {
                    open.Add(name);
                    if (name == "table" || name == "td" || name == "th") sawVisible = true;
                }
                else
                {
                    output.Append("</").Append(name).Append('>');
                }

                i = afterTag;
            }

            for (int k = open.Count - 1; k >= 0; k--)
            {
                output.Append("</").Append(open[k]).Append('>');
                notes.Add("closed an unclosed <" + open[k] + ">");
            }

            return new HtmlNormalizationResult(output.ToString(), notes.ToList(), sawVisible);
        }

        private static bool AppendText(string text, StringBuilder output, List<string> open, AdjustmentLog notes)
        {
            bool visible = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (!char.IsWhiteSpace(text[i])) { visible = true; break; }
            }

            if (visible)
            {
                // Text may not sit directly inside a table or a list - give it the cell/item it needs.
                EnsureAncestors("#text", output, open, notes);
            }
            else if (open.Count > 0 && (TableElements.Contains(open[open.Count - 1]) || ListElements.Contains(open[open.Count - 1])))
            {
                return false; // whitespace between structural tags: drop it, it only confuses Word's importer
            }

            AppendEscapedText(text, output);
            return visible;
        }

        private static void AppendEscapedText(string text, StringBuilder output)
        {
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '&')
                {
                    int len = EntityLength(text, i);
                    if (len > 0) { output.Append(text, i, len); i += len - 1; }
                    else output.Append("&amp;");
                }
                else if (c == '<') output.Append("&lt;");
                else if (c == '>') output.Append("&gt;");
                else output.Append(c);
            }
        }

        private static int EntityLength(string text, int start)
        {
            int i = start + 1;
            if (i < text.Length && text[i] == '#')
            {
                i++;
                if (i < text.Length && (text[i] == 'x' || text[i] == 'X')) i++;
                int digits = 0;
                while (i < text.Length && Uri.IsHexDigit(text[i])) { i++; digits++; }
                if (digits > 0 && i < text.Length && text[i] == ';') return i - start + 1;
                return 0;
            }

            int letters = 0;
            while (i < text.Length && IsNameChar(text[i]) && letters < 32) { i++; letters++; }
            if (letters > 0 && i < text.Length && text[i] == ';') return i - start + 1;
            return 0;
        }

        private static void CloseTag(string name, StringBuilder output, List<string> open, AdjustmentLog notes)
        {
            int at = open.LastIndexOf(name);
            if (at < 0)
            {
                notes.Add("dropped a stray </" + name + ">");
                return;
            }

            for (int k = open.Count - 1; k > at; k--)
            {
                output.Append("</").Append(open[k]).Append('>');
                notes.Add("closed a mis-nested <" + open[k] + ">");
                open.RemoveAt(k);
            }

            output.Append("</").Append(name).Append('>');
            open.RemoveAt(at);
        }

        private static void ApplyImplicitCloses(string name, StringBuilder output, List<string> open, AdjustmentLog notes)
        {
            while (open.Count > 0)
            {
                string top = open[open.Count - 1];
                bool close =
                    (name == "li" && top == "li") ||
                    ((name == "dt" || name == "dd") && (top == "dt" || top == "dd")) ||
                    ((name == "td" || name == "th") && (top == "td" || top == "th")) ||
                    (name == "tr" && (top == "td" || top == "th" || top == "tr")) ||
                    ((name == "thead" || name == "tbody" || name == "tfoot") &&
                        (top == "td" || top == "th" || top == "tr" || top == "thead" || top == "tbody" || top == "tfoot")) ||
                    (top == "p" && ClosesOpenParagraph.Contains(name));

                if (!close) break;
                output.Append("</").Append(top).Append('>');
                open.RemoveAt(open.Count - 1);
            }
        }

        private static void EnsureAncestors(string name, StringBuilder output, List<string> open, AdjustmentLog notes)
        {
            string top = open.Count > 0 ? open[open.Count - 1] : string.Empty;

            if (name == "li")
            {
                if (top != "ul" && top != "ol")
                {
                    output.Append("<ul>");
                    open.Add("ul");
                    notes.Add("added the missing <ul> around a stray <li>");
                }
                return;
            }

            if (name == "dt" || name == "dd")
            {
                if (top != "dl") { output.Append("<dl>"); open.Add("dl"); notes.Add("added the missing <dl>"); }
                return;
            }

            if (name == "td" || name == "th")
            {
                if (top != "tr") { EnsureAncestors("tr", output, open, notes); output.Append("<tr>"); open.Add("tr"); notes.Add("added the missing <tr>"); }
                return;
            }

            if (name == "tr" || name == "thead" || name == "tbody" || name == "tfoot" || name == "caption" || name == "colgroup")
            {
                if (!open.Contains("table")) { output.Append("<table>"); open.Add("table"); notes.Add("added the missing <table>"); }
                return;
            }

            if (name == "#text")
            {
                if (top == "ul" || top == "ol") { output.Append("<li>"); open.Add("li"); notes.Add("wrapped loose list text in <li>"); }
                else if (top == "dl") { output.Append("<dd>"); open.Add("dd"); notes.Add("wrapped loose list text in <dd>"); }
                else if (top == "table" || top == "thead" || top == "tbody" || top == "tfoot" || top == "tr")
                {
                    EnsureAncestors("td", output, open, notes);
                    output.Append("<td>");
                    open.Add("td");
                    notes.Add("wrapped loose table text in <td>");
                }
            }
        }

        private static void AppendAttributes(string element, string raw, StringBuilder output, AdjustmentLog notes)
        {
            int i = 0;
            while (i < raw.Length)
            {
                while (i < raw.Length && (char.IsWhiteSpace(raw[i]) || raw[i] == '/')) i++;
                if (i >= raw.Length) break;
                if (!IsNameStart(raw[i])) { i++; continue; }

                int nameStart = i;
                while (i < raw.Length && (IsNameChar(raw[i]) || raw[i] == ':')) i++;
                string name = raw.Substring(nameStart, i - nameStart).ToLowerInvariant();

                while (i < raw.Length && char.IsWhiteSpace(raw[i])) i++;
                string value = string.Empty;
                if (i < raw.Length && raw[i] == '=')
                {
                    i++;
                    while (i < raw.Length && char.IsWhiteSpace(raw[i])) i++;
                    if (i < raw.Length && (raw[i] == '"' || raw[i] == '\''))
                    {
                        char quote = raw[i++];
                        int valueStart = i;
                        while (i < raw.Length && raw[i] != quote) i++;
                        value = raw.Substring(valueStart, i - valueStart);
                        if (i < raw.Length) i++;
                    }
                    else
                    {
                        int valueStart = i;
                        while (i < raw.Length && !char.IsWhiteSpace(raw[i]) && raw[i] != '>') i++;
                        value = raw.Substring(valueStart, i - valueStart);
                    }
                }

                string? emitted = SanitizeAttribute(element, name, value, notes);
                if (emitted != null)
                {
                    output.Append(' ').Append(name).Append("=\"").Append(emitted).Append('"');
                }
            }
        }

        private static string? SanitizeAttribute(string element, string name, string value, AdjustmentLog notes)
        {
            if (name.StartsWith("on", StringComparison.Ordinal))
            {
                notes.Add("removed the event handler attribute " + name);
                return null;
            }

            if (name == "id" || name == "name")
            {
                // A forged _MailAutoSig/_MailOriginal anchor would redraw the draft/signature/quote
                // boundary Outlook and the writing sidebar both rely on. Never let one through.
                notes.Add("removed the " + name + " attribute (reserved for Outlook's own markers)");
                return null;
            }

            if (name == "style")
            {
                string style = SanitizeStyle(value, notes);
                return style.Length == 0 ? null : EscapeAttribute(style);
            }

            if (name == "href")
            {
                if (element != "a") { notes.Add("removed href from <" + element + ">"); return null; }
                string trimmed = value.Trim();
                if (trimmed.Length == 0) return null;
                if (trimmed.StartsWith("#", StringComparison.Ordinal))
                {
                    notes.Add("removed an in-page link (there is no page to jump to in a mail)");
                    return null;
                }
                bool safe = false;
                for (int k = 0; k < SafeUrlSchemes.Length; k++)
                {
                    if (trimmed.StartsWith(SafeUrlSchemes[k], StringComparison.OrdinalIgnoreCase)) { safe = true; break; }
                }
                if (!safe && trimmed.IndexOf(':') < 0) safe = true; // scheme-less relative form: harmless text
                if (!safe)
                {
                    notes.Add("removed a link with an unsupported URL scheme");
                    return null;
                }
                return EscapeAttribute(trimmed);
            }

            if (GlobalAttributes.Contains(name)) return EscapeAttribute(value);

            if ((TableElements.Contains(element) || element == "img") && TableAttributes.Contains(name))
                return EscapeAttribute(value);

            if ((element == "ol" || element == "ul" || element == "li") && (name == "start" || name == "type" || name == "value"))
                return EscapeAttribute(value);

            notes.Add("removed the unsupported attribute " + name);
            return null;
        }

        private static string SanitizeStyle(string value, AdjustmentLog notes)
        {
            var kept = new StringBuilder();
            string[] declarations = value.Split(';');
            for (int i = 0; i < declarations.Length; i++)
            {
                string declaration = declarations[i].Trim();
                if (declaration.Length == 0) continue;
                if (declaration.IndexOf(':') < 0) continue;

                bool dangerous = false;
                string lowered = declaration.ToLowerInvariant().Replace(" ", string.Empty);
                for (int k = 0; k < DangerousStyleTokens.Length; k++)
                {
                    if (lowered.IndexOf(DangerousStyleTokens[k], StringComparison.Ordinal) >= 0) { dangerous = true; break; }
                }
                if (dangerous)
                {
                    notes.Add("removed a CSS declaration that loads or executes something");
                    continue;
                }

                if (kept.Length + declaration.Length + 2 > MaxStyleChars)
                {
                    notes.Add("shortened an over-long style attribute");
                    break;
                }

                if (kept.Length > 0) kept.Append("; ");
                kept.Append(declaration);
            }

            return kept.ToString();
        }

        private static string EscapeAttribute(string value)
        {
            var sb = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == '&') sb.Append("&amp;");
                else if (c == '"') sb.Append("&quot;");
                else if (c == '<') sb.Append("&lt;");
                else if (c == '>') sb.Append("&gt;");
                else if (c == '\r' || c == '\n' || c == '\t') sb.Append(' ');
                else sb.Append(c);
            }
            return sb.ToString();
        }

        private static int SkipElement(string html, int from, string name)
        {
            string closing = "</" + name;
            int i = from;
            while (i < html.Length)
            {
                int at = html.IndexOf(closing, i, StringComparison.OrdinalIgnoreCase);
                if (at < 0) return html.Length;
                int after = at + closing.Length;
                if (after >= html.Length || !IsNameChar(html[after]))
                {
                    int close = html.IndexOf('>', at);
                    return close < 0 ? html.Length : close + 1;
                }
                i = after;
            }
            return html.Length;
        }

        private static int FindTagEnd(string html, int from)
        {
            bool inQuote = false;
            char quote = '"';
            for (int i = from; i < html.Length; i++)
            {
                char c = html[i];
                if (inQuote)
                {
                    if (c == quote) inQuote = false;
                }
                else if (c == '"' || c == '\'')
                {
                    inQuote = true;
                    quote = c;
                }
                else if (c == '>')
                {
                    return i;
                }
            }
            return html.Length;
        }

        private static bool StartsWith(string s, int at, string value)
        {
            return at + value.Length <= s.Length && string.CompareOrdinal(s, at, value, 0, value.Length) == 0;
        }

        private static bool IsNameStart(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        private static bool IsNameChar(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-' || c == '_';
        }

        private sealed class AdjustmentLog
        {
            private readonly List<string> _order = new List<string>();
            private readonly Dictionary<string, int> _counts = new Dictionary<string, int>(StringComparer.Ordinal);

            internal void Add(string message)
            {
                if (_counts.TryGetValue(message, out int count))
                {
                    _counts[message] = count + 1;
                    return;
                }
                _counts[message] = 1;
                _order.Add(message);
            }

            internal IReadOnlyList<string> ToList()
            {
                var result = new List<string>(Math.Min(_order.Count, MaxReportedAdjustments + 1));
                for (int i = 0; i < _order.Count && i < MaxReportedAdjustments; i++)
                {
                    string message = _order[i];
                    int count = _counts[message];
                    result.Add(count > 1
                        ? message + " (x" + count.ToString(CultureInfo.InvariantCulture) + ")"
                        : message);
                }
                if (_order.Count > MaxReportedAdjustments)
                {
                    result.Add("and " + (_order.Count - MaxReportedAdjustments).ToString(CultureInfo.InvariantCulture) +
                               " further adjustment(s) not listed");
                }
                return result;
            }
        }
    }
}
