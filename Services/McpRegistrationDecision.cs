using System;
using System.Collections.Generic;

namespace OutlookAI.Services
{
    /// <summary>
    /// WHEN OUTLOOK MAY CHANGE THE USER-SCOPE MCP REGISTRATION BY ITSELF, AND WHEN IT HAS TO
    /// STOP AND ASK.
    ///
    /// The rule the whole file exists to enforce: <b>never act on inferred intent</b>. The
    /// add-in edits <c>~/.claude.json</c>, a file it does not own and the user may well have
    /// hand-edited, so every state in which "what the user wants" is genuinely unknown ends in
    /// a question rather than a guess. Only three states are unambiguous enough to act on
    /// silently, and they are listed below.
    ///
    /// Two inputs, and both are deliberately coarse:
    ///  - the STORED intent — the tri-state DWORD under <c>HKCU\Software\OutlookAI\Mcp</c>:
    ///    1 on, 0 off, ABSENT never decided;
    ///  - the ENTRY — what <c>mcpServers.outlookai</c> currently is, reduced to
    ///    <see cref="EntryState"/>. "Ours" means its <c>command</c> already resolves to the
    ///    installed mail server; an entry naming ANYTHING else is Foreign and is never
    ///    adopted, never overwritten and never removed on our own initiative.
    ///
    /// <para>The matrix (Unreadable is always "leave it completely alone"):</para>
    /// <list type="table">
    /// <item><description>undecided + no entry ⇒ ASK (first run)</description></item>
    /// <item><description>undecided + entry is OURS ⇒ adopt silently, persist ON — an upgrade
    /// of a working install must not be interrogated</description></item>
    /// <item><description>undecided + FOREIGN entry ⇒ ASK (replace it, or leave it alone)</description></item>
    /// <item><description>ON + entry is OURS ⇒ keep it correct (a correct entry is a no-op)</description></item>
    /// <item><description>ON + FOREIGN entry ⇒ repoint it — an explicit ON is a standing
    /// instruction to keep the entry naming the installed server, and healing exactly that
    /// drift (an older install path, a build output) is what the toggle promises</description></item>
    /// <item><description>ON + no entry ⇒ ASK (register it again, or turn the setting off)</description></item>
    /// <item><description>OFF + no entry ⇒ nothing at all</description></item>
    /// <item><description>OFF + entry is OURS ⇒ ASK (remove it, or turn the setting on)</description></item>
    /// <item><description>OFF + FOREIGN entry ⇒ nothing at all — it is not ours to delete</description></item>
    /// </list>
    ///
    /// <para>
    /// Pure by construction — no registry, no filesystem, no Outlook, no COM, no UI — so the
    /// T1 suite pins the shipped code rather than a re-implementation (the file is LINKED into
    /// the test project, exactly like <see cref="McpConfigEditor"/>). Framework-neutral for the
    /// same reason: it compiles into the net48 add-in and into the .NET 10 test host. Public
    /// for that reason too, and only that one — the theory signatures that walk the matrix have
    /// to be able to name these types, the same bargain <c>ComposeSurface</c> makes; nothing
    /// outside the add-in consumes any of it, and the assembly is <c>ComVisible(false)</c>.
    /// </para>
    /// </summary>
    public static class McpRegistrationDecision
    {
        /// <summary>What <c>mcpServers.outlookai</c> currently is, in the only four flavours that matter.</summary>
        public enum EntryState
        {
            /// <summary>No <c>outlookai</c> member at all (or no <c>mcpServers</c> object).</summary>
            Absent,

            /// <summary>A stdio entry whose <c>command</c> resolves to our installed mail server.</summary>
            Ours,

            /// <summary>
            /// An <c>outlookai</c> member that is NOT ours: a command pointing somewhere else
            /// (a wrapper script the user chose), a remote entry, or a value that is not even
            /// an object. Never adopted, never overwritten, never deleted unless the user says so.
            /// </summary>
            Foreign,

            /// <summary>
            /// The configuration could not be read as JSON (or read back empty). Nothing is
            /// known, so nothing is asked and nothing is written.
            /// </summary>
            Unreadable,
        }

        /// <summary>Which question to put to the user; <see cref="None"/> when none is warranted.</summary>
        public enum PromptKind
        {
            None,

            /// <summary>Nothing decided, nothing registered: offer to register for all projects.</summary>
            FirstRun,

            /// <summary>Nothing decided, and an <c>outlookai</c> entry points elsewhere: replace it, or leave it.</summary>
            ForeignEntry,

            /// <summary>Registration is on, but the entry is gone: put it back, or turn the setting off.</summary>
            EntryMissing,

            /// <summary>Registration is off, yet an entry of ours exists: remove it, or turn the setting on.</summary>
            EntryUnexpected,
        }

        /// <summary>What the reconcile may do without asking.</summary>
        public enum RegistrationAction
        {
            /// <summary>Touch nothing. The only action allowed while a question is outstanding.</summary>
            None,

            /// <summary>Make <c>mcpServers.outlookai</c> name the installed server (no-op when it already does).</summary>
            Register,

            /// <summary>Persist the opt-in ON first, then <see cref="Register"/>.</summary>
            AdoptAndRegister,

            /// <summary>Remove our entry. Only ever reached for <see cref="EntryState.Ours"/>.</summary>
            Remove,
        }

        /// <summary>The verdict: at most one question, at most one action, never both.</summary>
        public sealed class Decision
        {
            public Decision(PromptKind prompt, RegistrationAction action, bool deferred)
            {
                Prompt = prompt;
                Action = action;
                Deferred = deferred;
            }

            /// <summary>The question to ask, or <see cref="PromptKind.None"/>.</summary>
            public PromptKind Prompt { get; private set; }

            /// <summary>What to do. Always <see cref="RegistrationAction.None"/> when a question is outstanding.</summary>
            public RegistrationAction Action { get; private set; }

            /// <summary>
            /// True when a question was warranted but cannot be put to anyone right now — an
            /// Outlook running in the background with no window, or one that has already asked
            /// this session. Nothing is written and nothing is remembered, so the next
            /// interactive session asks instead.
            /// </summary>
            public bool Deferred { get; private set; }

            /// <summary>True when this verdict changes nothing on disk. Every prompt verdict is one.</summary>
            public bool ChangesNothing
            {
                get { return Action == RegistrationAction.None; }
            }
        }

        private static readonly Decision Nothing =
            new Decision(PromptKind.None, RegistrationAction.None, false);

        /// <summary>
        /// The whole decision, from the stored intent and what is actually in the config.
        ///
        /// <paramref name="canAskUser"/> is the host's answer to "could I put a dialog in
        /// front of a human right now?" — false for a background Outlook with no visible
        /// window (the headless autostart the mail server relies on), and false once this
        /// session has already asked. A question that cannot be asked becomes
        /// <see cref="Decision.Deferred"/>, never a guess and never a write.
        ///
        /// <paramref name="intentJustDeclared"/> is set when the user has just said what they
        /// want — ticked the box in OutlookAI Settings, or answered one of these very prompts.
        /// It suppresses the two questions that only exist because the world disagreed with a
        /// stored preference: with the preference freshly declared there is nothing left to
        /// disagree about, so "on" registers and "off" removes, immediately.
        /// </summary>
        public static Decision Decide(
            int? storedPreference, EntryState entry, bool canAskUser, bool intentJustDeclared)
        {
            // A file we could not read tells us nothing, so it earns neither a question we
            // could not word truthfully nor a write that would destroy it.
            if (entry == EntryState.Unreadable)
                return Nothing;

            if (!storedPreference.HasValue)
            {
                // Never decided. The one state that may be resolved without asking is an entry
                // that is ALREADY ours: it can only have come from an earlier version of this
                // add-in (or from the user registering the very same server by hand), so
                // adopting it changes nothing the user can observe.
                if (entry == EntryState.Ours)
                    return new Decision(PromptKind.None, RegistrationAction.AdoptAndRegister, false);

                return Ask(entry == EntryState.Foreign ? PromptKind.ForeignEntry : PromptKind.FirstRun, canAskUser);
            }

            if (storedPreference.Value != 0)
            {
                // ON. An entry that is missing contradicts that, and only the user can say
                // whether it was removed on purpose. Anything else is kept pointing at the
                // installed server, which is the drift-healing the toggle promises.
                if (entry == EntryState.Absent && !intentJustDeclared)
                    return Ask(PromptKind.EntryMissing, canAskUser);

                return new Decision(PromptKind.None, RegistrationAction.Register, false);
            }

            // OFF. An entry of ours that is nevertheless present contradicts that — someone
            // added it outside Outlook — so ask before deleting anything.
            if (entry == EntryState.Ours)
            {
                return intentJustDeclared
                    ? new Decision(PromptKind.None, RegistrationAction.Remove, false)
                    : Ask(PromptKind.EntryUnexpected, canAskUser);
            }

            // Absent: already how it should be. Foreign: not ours, so not ours to delete —
            // removing it is exactly the damage the "only adopt what is already ours" rule
            // exists to prevent, and it would silently undo a "leave it alone" answer.
            return Nothing;
        }

        private static Decision Ask(PromptKind prompt, bool canAskUser)
        {
            return canAskUser
                ? new Decision(prompt, RegistrationAction.None, false)
                : new Decision(PromptKind.None, RegistrationAction.None, true);
        }

        /// <summary>
        /// Whether the stored opt-in reads as ON, for the settings dialog's tick box.
        ///
        /// Undecided falls back to <paramref name="ourEntryAlreadyPresent"/> so the dialog and
        /// the reconcile agree about a machine that is about to be adopted silently. It is the
        /// same evidence <see cref="Decide"/> uses, and deliberately NOT "any entry exists":
        /// a foreign entry means the tick box shows off, which is the truth — we did not
        /// register it and we are about to ask about it.
        /// </summary>
        public static bool ResolveOptIn(int? storedPreference, bool ourEntryAlreadyPresent)
        {
            if (storedPreference.HasValue)
                return storedPreference.Value != 0;
            return ourEntryAlreadyPresent;
        }

        // ===== Is anyone actually there to be asked? =====

        /// <summary>
        /// Where <c>ComposeSurface</c> parks a window it must keep off a human's screen.
        /// Mirrored (not referenced) because the add-in is net48 and cannot reference the
        /// .NET 10 server assembly; the two constants must stay in step.
        /// </summary>
        public const int ParkX = -32000;

        /// <summary>The other half of <see cref="ParkX"/>.</summary>
        public const int ParkY = -32000;

        /// <summary>
        /// One top-level window of the Outlook process, reduced to what the "is a human
        /// looking at this?" rule needs — a plain record so the rule is testable without a
        /// window manager, the same shape as <c>ComposeSurface.WindowState</c>.
        /// </summary>
        public sealed class OutlookWindow
        {
            public OutlookWindow(bool visible, bool minimized, int left, int top, int right, int bottom)
            {
                Visible = visible;
                Minimized = minimized;
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            /// <summary>
            /// Win32 <c>IsWindowVisible</c>. TRUE for a minimized window and TRUE for a
            /// DWM-cloaked one — which is what makes it the right predicate here: the user's
            /// own minimized Outlook still counts as a session that can be asked a question.
            /// </summary>
            public bool Visible { get; private set; }

            /// <summary>
            /// Win32 <c>IsIconic</c>. Load-bearing, and the one place this rule is deliberately
            /// finer than <c>ComposeSurface.CountUserVisibleWindows</c>: a MINIMIZED window
            /// reports a rectangle in the same far-off-screen corner a PARKED one does, and the
            /// two mean opposite things here. Minimized is the user's own Outlook, sitting on
            /// their taskbar, perfectly able to receive a question; parked is a window the mail
            /// server hid from them on purpose.
            /// </summary>
            public bool Minimized { get; private set; }

            public int Left { get; private set; }

            public int Top { get; private set; }

            public int Right { get; private set; }

            public int Bottom { get; private set; }
        }

        /// <summary>
        /// Whether any of these windows is one a human could actually see — the same test
        /// <c>ComposeSurface.CountUserVisibleWindows</c> applies, for the same reason.
        ///
        /// This is the headless guard. The product deliberately autostarts Outlook in the
        /// background for agent sessions (D17/D33): no window, tray icon only, and the mail
        /// server depends on that instance staying responsive. A modal dialog there would be
        /// invisible to everyone and would wedge the reconcile that raised it, so a false here
        /// means DEFER — ask nothing, write nothing, and leave every stored value exactly as
        /// it was found.
        ///
        /// A window that is invisible, collapsed to nothing, or parked in the off-screen
        /// corner the invisible compose surface uses is not a window anyone can see. A
        /// MINIMIZED one is — it is on the user's taskbar, and Windows gives it the same
        /// far-off-screen rectangle as a parked window, which is exactly why it is recognised
        /// before the rectangle is ever looked at.
        /// </summary>
        public static bool AnyWindowAHumanCanSee(IEnumerable<OutlookWindow> windows)
        {
            if (windows == null)
                return false;

            foreach (OutlookWindow w in windows)
            {
                if (w == null || !w.Visible)
                    continue;
                if (w.Minimized)
                    return true;
                if (w.Right <= w.Left || w.Bottom <= w.Top)
                    continue;
                if (w.Left <= ParkX / 2 && w.Top <= ParkY / 2)
                    continue;

                return true;
            }

            return false;
        }
    }
}
