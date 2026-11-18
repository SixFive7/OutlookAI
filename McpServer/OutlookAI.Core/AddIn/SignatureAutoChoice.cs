using System;
using System.Collections.Generic;

namespace OutlookAI.Core.AddIn
{
    /// <summary>What the add-in's own signature choice does with a draft, and why.</summary>
    public enum SignatureAutoChoiceOutcome
    {
        /// <summary>The AI may choose among <see cref="SignatureAutoChoicePlan.Candidates"/>.</summary>
        Choose,

        /// <summary>The tick box is off.</summary>
        SettingOff,

        /// <summary>Policy disables signatures: nothing may be inserted.</summary>
        SignaturesDisabled,

        /// <summary>
        /// A policy or a deployed setting names a default (<see cref="SignatureCandidates.DefaultsForced"/>; Q156,
        /// Q156 (d)): the AI makes no choice of its own then. A person may still pick another signature for one
        /// message, as Outlook allows (Q200 (a)).
        /// </summary>
        DefaultsForced,

        /// <summary>A person put the draft's signature there, or changed it: it is theirs (Q162 (b)).</summary>
        ChosenByPerson,

        /// <summary>
        /// The draft has no signature: Outlook gave this account none for this kind of mail, and the
        /// automatic choice changes a signature, it does not add one.
        /// </summary>
        NoSignatureInDraft,

        /// <summary>The account can use one signature or none, so there is nothing to choose between.</summary>
        NothingToChooseBetween,
    }

    /// <summary>The add-in's plan for one draft: choose, among which signatures, or leave it, and why.</summary>
    public sealed class SignatureAutoChoicePlan
    {
        /// <summary>Creates the plan.</summary>
        public SignatureAutoChoicePlan(SignatureAutoChoiceOutcome outcome, IReadOnlyList<SignatureInfo> candidates)
        {
            Outcome = outcome;
            Candidates = candidates ?? Array.Empty<SignatureInfo>();
        }

        /// <summary>Choose, or why not.</summary>
        public SignatureAutoChoiceOutcome Outcome { get; }

        /// <summary>The signatures the AI may choose among; empty unless <see cref="Outcome"/> is Choose.</summary>
        public IReadOnlyList<SignatureInfo> Candidates { get; }
    }

    /// <summary>
    /// The add-in's own choice of a draft's signature (Q158, Q162 (b), decided 2026-11-17): under the
    /// tick box "Allow the AI to automatically choose the correct signature", on by default, the AI
    /// picks the signature - by the language of the mail being answered, for one - when it writes or
    /// rewrites the body and when a reply or forward opens, from the signatures the sending account can
    /// use. It never replaces a signature a person chose, and never changes a mail as it is sent. The
    /// signatures come from <see cref="SignatureSelector.Candidates"/>, the rule the MCP server's drafts
    /// use; what is here is what only the add-in decides: whether to ask the AI at all, and whether the
    /// signature in the open draft is a person's. Pure, so T1 pins it (Q163 (a)).
    /// </summary>
    public static class SignatureAutoChoice
    {
        /// <summary>
        /// The tick box's value under <c>HKCU\Software\OutlookAI</c>: a DWORD, 0 for off. Absent is on -
        /// the maintainer's "(on by default)".
        /// </summary>
        public const string SettingValueName = "AutoChooseSignature";

        /// <summary>Whether the tick box is on: anything but a stored 0.</summary>
        public static bool IsOn(int? storedValue) => storedValue != 0;

        /// <summary>
        /// Whether a person has changed the signature since Outlook or the add-in last put one in:
        /// the signature region's text now against its text then. Never, when nothing has been put
        /// there yet that the add-in saw.
        /// </summary>
        public static bool ChangedByPerson(string? regionNow, string? regionWhenLastPut)
        {
            if (regionWhenLastPut == null)
            {
                return false;
            }

            return !string.Equals(Normalize(regionNow), Normalize(regionWhenLastPut), StringComparison.Ordinal);
        }

        /// <summary>Whether to ask the AI, and among which signatures. Pure.</summary>
        public static SignatureAutoChoicePlan Plan(
            bool settingOn, SignatureCandidates candidates, bool draftHasSignature, bool changedByPerson)
        {
            if (candidates == null)
            {
                throw new ArgumentNullException(nameof(candidates));
            }

            if (!settingOn)
            {
                return Leave(SignatureAutoChoiceOutcome.SettingOff);
            }

            if (candidates.SignaturesDisabled)
            {
                return Leave(SignatureAutoChoiceOutcome.SignaturesDisabled);
            }

            if (candidates.DefaultsForced)
            {
                return Leave(SignatureAutoChoiceOutcome.DefaultsForced);
            }

            if (changedByPerson)
            {
                return Leave(SignatureAutoChoiceOutcome.ChosenByPerson);
            }

            if (!draftHasSignature)
            {
                return Leave(SignatureAutoChoiceOutcome.NoSignatureInDraft);
            }

            if (candidates.Usable.Count < 2)
            {
                return Leave(SignatureAutoChoiceOutcome.NothingToChooseBetween);
            }

            return new SignatureAutoChoicePlan(SignatureAutoChoiceOutcome.Choose, candidates.Usable);
        }

        /// <summary>
        /// The signature among <paramref name="signatures"/> that the AI's answer names: case-insensitive,
        /// with the quotes and the full stop a model may put around a name taken off. Null when it names
        /// none - the answer is the model's, so it is matched, never trusted.
        /// </summary>
        public static SignatureInfo? FindNamed(IReadOnlyList<SignatureInfo> signatures, string? answer)
        {
            if (signatures == null || answer == null || answer.Trim().Length == 0)
            {
                return null;
            }

            string trimmed = answer.Trim().Trim('"', '\'', '.');
            foreach (SignatureInfo signature in signatures)
            {
                if (string.Equals(signature.Name, trimmed, StringComparison.OrdinalIgnoreCase))
                {
                    return signature;
                }
            }

            return null;
        }

        private static SignatureAutoChoicePlan Leave(SignatureAutoChoiceOutcome outcome) =>
            new SignatureAutoChoicePlan(outcome, Array.Empty<SignatureInfo>());

        private static string Normalize(string? regionText) => (regionText ?? string.Empty).Trim();
    }
}
