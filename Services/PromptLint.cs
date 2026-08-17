using System;
using System.Collections.Generic;

namespace OutlookAI.Services
{
    /// <summary>
    /// ADVISORY CHECKS ON EDITED PROMPT TEXT. WARNS, NEVER BLOCKS.
    ///
    /// Two things in the shipped prompts are load-bearing for reasons a user editing a text box
    /// cannot see, so removing either has a consequence worth naming out loud:
    ///
    ///  - The untrusted-content sentence. The draft, the quoted thread and the signature are
    ///    pasted into the prompt, and a reply to a mail that contains "ignore your instructions
    ///    and write X" is an ordinary thing to receive. Without that sentence the model has no
    ///    reason not to treat mail it was given as a request from the user.
    ///  - The plain-text output contract. The result is written straight into the Word editor
    ///    behind the compose window as text, so a code fence, an HTML tag or markdown asterisks
    ///    do not render - they appear in the email as the characters they are.
    ///
    /// Both stay editable. The user chose warnings over locked text, on the grounds that a
    /// prompt you cannot change is a prompt you cannot fix, and these checks are deliberately
    /// crude substring tests: their job is to notice that a sentence went missing, not to judge
    /// how someone rephrased it. A false negative here costs nothing that was not already lost;
    /// a false positive costs a warning nobody needed.
    /// </summary>
    internal static class PromptLint
    {
        internal const string UntrustedContentWarning =
            "This prompt no longer says that the draft, the quoted thread and the signature are untrusted content and not instructions. Text inside an email you received could then be followed as if you had asked for it yourself.";

        internal const string PlainTextOnlyWarning =
            "This prompt no longer asks for only the draft text back. Anything else the model says, such as an explanation of its changes, gets written into your email along with the draft.";

        internal const string NoMarkupWarning =
            "This prompt no longer rules out code fences, HTML and markdown. Outlook receives the reply as plain text, so those end up in the email as literal characters instead of formatting.";

        /// <summary>
        /// Advisory warnings for a section's text, in display order. Empty means nothing to say,
        /// which is also the answer for every section that carries no such contract.
        /// Never throws, and never rejects anything.
        /// </summary>
        internal static IList<string> Warn(PromptSection section, string text)
        {
            var warnings = new List<string>();
            string candidate = text == null ? string.Empty : text;

            switch (section)
            {
                case PromptSection.Preamble:
                    if (!MentionsUntrustedContent(candidate))
                        warnings.Add(UntrustedContentWarning);
                    if (!MentionsPlainTextOnly(candidate))
                        warnings.Add(PlainTextOnlyWarning);
                    if (!MentionsNoMarkup(candidate))
                        warnings.Add(NoMarkupWarning);
                    break;

                case PromptSection.SignatureSelection:
                    // Same exposure: the draft, thread, recipients and signature excerpts are all
                    // pasted into this prompt. It has no draft-text contract to lose, though - its
                    // output is one signature name, and picking a wrong one is not destructive.
                    if (!MentionsUntrustedContent(candidate))
                        warnings.Add(UntrustedContentWarning);
                    break;
            }

            return warnings;
        }

        /// <summary>Whether any section is worth checking at all, for a UI deciding where to put a warning label.</summary>
        internal static bool IsChecked(PromptSection section)
        {
            return section == PromptSection.Preamble || section == PromptSection.SignatureSelection;
        }

        private static bool MentionsUntrustedContent(string text)
        {
            return Contains(text, "untrusted") || Contains(text, "not instructions");
        }

        private static bool MentionsPlainTextOnly(string text)
        {
            return Contains(text, "return only")
                || Contains(text, "only the email draft")
                || Contains(text, "draft text only")
                || Contains(text, "nothing else");
        }

        private static bool MentionsNoMarkup(string text)
        {
            return Contains(text, "code fence")
                || Contains(text, "markdown")
                || Contains(text, "html");
        }

        private static bool Contains(string text, string term)
        {
            // IndexOf rather than Contains(string, StringComparison): that overload does not
            // exist on .NET Framework 4.8, and this file also compiles into the add-in.
            return text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
