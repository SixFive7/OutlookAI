using System;
using System.Collections.Generic;

namespace OutlookAI.Services
{
    /// <summary>What a picker entry means once it is chosen.</summary>
    internal enum ModelChoiceKind
    {
        /// <summary>Store nothing, send no <c>--model</c>, let Claude Code decide.</summary>
        Default,

        /// <summary>Store <see cref="ModelChoice.Alias"/> and send it as <c>--model</c>.</summary>
        Alias,

        /// <summary>Store whatever the free-text box holds, once it is a usable model id.</summary>
        Custom,
    }

    /// <summary>
    /// One entry in the model picker: what it stores, and what the user reads. Its
    /// <see cref="ToString"/> is the label, because a combo box shows exactly that and holding
    /// the object as the item is what lets the picker hand back a choice rather than a string it
    /// then has to match back to one.
    /// </summary>
    internal sealed class ModelChoice
    {
        internal ModelChoice(ModelChoiceKind kind, string alias, string label)
        {
            Kind = kind;
            Alias = alias;
            Label = label;
        }

        internal ModelChoiceKind Kind { get; private set; }

        /// <summary>
        /// The value that goes in the registry and on the command line. Null for
        /// <see cref="ModelChoiceKind.Default"/> and <see cref="ModelChoiceKind.Custom"/>:
        /// neither of those has a value of its own, one stores nothing and one stores what was
        /// typed.
        /// </summary>
        internal string Alias { get; private set; }

        internal string Label { get; private set; }

        public override string ToString()
        {
            return Label;
        }
    }

    /// <summary>
    /// Everything OutlookAI claims to know about Claude models, in one place.
    ///
    /// TWO DECISIONS SHAPE IT:
    ///
    ///  1. THE PICKER OFFERS ALIASES, NOT MODEL IDS. <c>opus</c>, <c>sonnet</c> and the rest are
    ///     stable names that Claude Code resolves to the newest model of that family, so this
    ///     list does not go stale when a model ships. A list of dated ids would - and a shipped
    ///     list of dated ids goes stale on every installed copy at once. Anybody who does want a
    ///     specific build types it into the custom box instead.
    ///
    ///  2. THE ONLY THING WE CLAIM ABOUT A MODEL IS WHETHER WE RECOGNISE IT. There is no
    ///     capability matrix here, because there is no way to maintain one: Claude Code exposes
    ///     no way to list models or report capabilities, and the API that does needs a key this
    ///     product deliberately does not have. So <see cref="IsKnownFamily"/> answers one
    ///     question - "have we heard of this family" - and the note it drives says exactly that
    ///     and nothing more.
    /// </summary>
    internal static class ModelCatalog
    {
        /// <summary>
        /// The first entry, and the shipped default. Deliberately NOT the <c>default</c> alias:
        /// storing nothing lets Claude Code apply the user's own <c>model</c> setting first,
        /// whereas passing <c>--model "default"</c> would override that setting with the account
        /// default. Two entries that differ only in whether they quietly ignore a setting the
        /// user made elsewhere is not a distinction a dropdown line can carry, so the alias is
        /// left out and this entry is the one that means "not my business".
        /// </summary>
        internal const string DefaultLabel = "Use Claude Code's default (recommended)";

        /// <summary>The last entry: enables the free-text box, and stores what it holds.</summary>
        internal const string CustomLabel = "Custom - a specific model id...";

        /// <summary>
        /// Shown when a hand-typed model id is not one of <see cref="KnownFamilies"/>. Worded as
        /// an admission rather than a limitation on purpose: the model may well be perfectly
        /// capable, we simply cannot say, and claiming otherwise would be inventing a fact.
        /// </summary>
        internal const string UnverifiedModelNote =
            "OutlookAI cannot verify what this model supports - it is not one of the families "
            + "OutlookAI knows about. It is passed to Claude Code exactly as typed. If Claude "
            + "Code does not accept it, your requests quietly run on its own default instead, "
            + "and OutlookAI will tell you the first time that happens.";

        /// <summary>Shown for text that could not be a model id at all, so nothing is stored.</summary>
        internal const string MalformedModelNote =
            "Not saved. A model id can only contain letters, digits, dots, dashes, underscores "
            + "and square brackets.";

        /// <summary>
        /// The stable aliases, in the order the picker shows them. Each one tracks the newest
        /// model of its family, so what a user picks today keeps meaning the same thing after a
        /// model release without anybody editing this file.
        /// </summary>
        private static readonly ModelChoice[] Aliases =
        {
            new ModelChoice(ModelChoiceKind.Alias, "best",
                "Best available - the strongest model your plan allows"),
            new ModelChoice(ModelChoiceKind.Alias, "opus",
                "Opus - the deepest reasoning, and the slowest"),
            new ModelChoice(ModelChoiceKind.Alias, "sonnet",
                "Sonnet - balanced speed and quality"),
            new ModelChoice(ModelChoiceKind.Alias, "haiku",
                "Haiku - the fastest, for short edits"),
            new ModelChoice(ModelChoiceKind.Alias, "fable",
                "Fable - Anthropic's most capable model"),
            new ModelChoice(ModelChoiceKind.Alias, "opusplan",
                "Opus while planning, Sonnet for the work"),
            new ModelChoice(ModelChoiceKind.Alias, "opus[1m]",
                "Opus with the extended context window"),
            new ModelChoice(ModelChoiceKind.Alias, "sonnet[1m]",
                "Sonnet with the extended context window"),
        };

        /// <summary>
        /// THE ONE PLACE THE MODEL FAMILIES ARE NAMED. Matched as id prefixes, so a dated build
        /// (claude-opus-4-5-20251101) is recognised by its family.
        ///
        /// LAST CHECKED: 2026-08-17, against Anthropic's published model list.
        ///
        /// Updating it is one edit and it is never urgent: a family missing from here only means
        /// a user who typed that id sees <see cref="UnverifiedModelNote"/> saying OutlookAI
        /// cannot vouch for it. Nothing is blocked, nothing is refused, and no capability is
        /// claimed either way - which is why this list is allowed to be a little behind without
        /// ever being wrong.
        /// </summary>
        private static readonly string[] KnownFamilies =
        {
            "claude-fable-5",
            "claude-opus-5",
            "claude-opus-4-8",
            "claude-opus-4-7",
            "claude-opus-4-6",
            "claude-opus-4-5",
            "claude-sonnet-5",
            "claude-sonnet-4-6",
            "claude-sonnet-4-5",
            "claude-haiku-4-5",
        };

        /// <summary>
        /// Every picker entry, in order: the deferring default, then the aliases, then Custom.
        /// Built fresh each call so the caller owns the list it fills a combo box from.
        /// </summary>
        internal static IList<ModelChoice> BuildChoices()
        {
            var choices = new List<ModelChoice>(Aliases.Length + 2);
            choices.Add(new ModelChoice(ModelChoiceKind.Default, null, DefaultLabel));
            foreach (ModelChoice alias in Aliases)
                choices.Add(alias);
            choices.Add(new ModelChoice(ModelChoiceKind.Custom, null, CustomLabel));
            return choices;
        }

        /// <summary>
        /// Whether <paramref name="value"/> could be a model name at all: letters, digits, dots,
        /// dashes, underscores and square brackets, and nothing else.
        ///
        /// This is a COMMAND-LINE guard before it is a spelling guard. The value ends up inside a
        /// quoted <c>--model "..."</c> argument, so a stray quote in a hand-typed id or a
        /// hand-edited registry value would not produce a bad model - it would produce a
        /// different command. Square brackets are allowed because the extended-context aliases
        /// (opus[1m], sonnet[1m]) contain them and they mean nothing to CreateProcess.
        /// </summary>
        internal static bool IsWellFormedModelId(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;
            foreach (char c in value.Trim())
            {
                if (!char.IsLetterOrDigit(c) && c != '.' && c != '-' && c != '_'
                    && c != '[' && c != ']')
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Whether this is a model family OutlookAI has heard of. Answers nothing about what the
        /// model can do - see the class comment. An alias is not a model id and is not claimed
        /// to be one; only the free-text box is checked against this.
        /// </summary>
        internal static bool IsKnownFamily(string modelId)
        {
            if (string.IsNullOrWhiteSpace(modelId))
                return false;
            string trimmed = modelId.Trim();
            foreach (string family in KnownFamilies)
            {
                if (trimmed.StartsWith(family, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }
}
