using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// Asks the user, in Outlook, what to do about the Claude Code registration — the one
    /// place a question from <see cref="McpRegistrationService"/> becomes something a human
    /// can answer.
    ///
    /// Three rules shape everything here, and they matter more than the wording:
    ///
    ///  1. NEVER ask when nobody is looking. Outlook is deliberately autostarted in the
    ///     background for agent sessions — no window, tray icon only — and the mail server
    ///     depends on that instance staying responsive. A modal dialog there would be
    ///     invisible to everyone and would wedge the reconcile that raised it. So the last
    ///     thing checked before showing anything is whether this process owns a window a
    ///     human could actually see; if it does not, the question is dropped for this session
    ///     and nothing is written. A later interactive session asks instead.
    ///
    ///  2. NEVER block Outlook starting. The reconcile that raises the question runs on a
    ///     background thread during startup; this marshals onto the UI thread and then waits
    ///     for startup to settle before putting anything on screen.
    ///
    ///  3. Ask at most once per Outlook session. <see cref="McpRegistrationService"/> hands a
    ///     question over exactly once; if it goes unanswered — dismissed, or never shown —
    ///     nothing happens until Outlook next starts, or until the user opens OutlookAI
    ///     Settings and decides there.
    /// </summary>
    internal static class McpRegistrationPrompt
    {
        /// <summary>
        /// How long to let Outlook finish starting before the first look. Long enough that a
        /// normal launch has painted its window (which is why the "is anyone there?" test is
        /// made HERE and not in the startup reconcile, where no window exists yet).
        /// </summary>
        private const int FirstLookDelayMs = 5000;

        /// <summary>Gap between later looks, for a background Outlook the user may still promote.</summary>
        private const int LaterLookDelayMs = 60000;

        /// <summary>Total looks before the question is dropped for this session (~5 minutes).</summary>
        private const int MaxLooks = 6;

        private static Timer _timer;
        private static int _looks;
        private static bool _shuttingDown;

        /// <summary>Makes this the destination for registration questions. UI thread, at startup.</summary>
        internal static void Install()
        {
            _shuttingDown = false;
            McpRegistrationService.PromptHost = OnPromptRequested;
        }

        /// <summary>Stops asking anything, and disposes the timer. UI thread, at shutdown.</summary>
        internal static void Shutdown()
        {
            _shuttingDown = true;
            McpRegistrationService.PromptHost = null;
            StopTimer();
        }

        /// <summary>
        /// Called by the reconcile, on whatever thread it runs on. Returns immediately: it only
        /// posts the work to the UI thread, because the reconcile holds its own lock while
        /// calling this and Outlook's startup must not wait on a dialog.
        /// </summary>
        private static void OnPromptRequested(McpRegistrationDecision.PromptKind prompt)
        {
            if (_shuttingDown || prompt == McpRegistrationDecision.PromptKind.None)
                return;

            try
            {
                var ui = Globals.ThisAddIn == null ? null : Globals.ThisAddIn.UiMarshalControl;
                if (ui == null || ui.IsDisposed || !ui.IsHandleCreated)
                    return;

                ui.BeginInvoke((Action)(() => Schedule(prompt)));
            }
            catch (Exception ex)
            {
                // A question that cannot be marshaled is simply not asked this session.
                System.Diagnostics.Debug.WriteLine("MCP prompt marshal: " + ex.Message);
            }
        }

        /// <summary>
        /// Arms the delayed look. UI thread. Which question gets asked is settled at that
        /// later moment, not now, so the state it describes is the state that is still true.
        /// </summary>
        private static void Schedule(McpRegistrationDecision.PromptKind prompt)
        {
            if (_shuttingDown || prompt == McpRegistrationDecision.PromptKind.None || _timer != null)
                return;

            try
            {
                _looks = 0;
                _timer = new Timer { Interval = FirstLookDelayMs };
                _timer.Tick += OnLook;
                _timer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt schedule: " + ex.Message);
                StopTimer();
            }
        }

        private static void OnLook(object sender, EventArgs e)
        {
            if (_shuttingDown)
            {
                StopTimer();
                return;
            }

            try
            {
                _looks++;

                // RULE 1. A background Outlook has no window anyone can see, so there is
                // nobody to ask: look again in case the user promotes it, and otherwise let
                // the question lapse without touching a single stored value.
                if (!AnyoneIsLooking())
                {
                    if (_looks >= MaxLooks)
                        StopTimer();
                    else
                        _timer.Interval = LaterLookDelayMs;
                    return;
                }

                StopTimer();

                // The world may have moved in the seconds since the question was raised — the
                // user could have ticked the box in OutlookAI Settings, or the CLI could have
                // written the entry. Re-check rather than ask something already answered.
                var prompt = McpRegistrationService.PendingPrompt();
                if (prompt == McpRegistrationDecision.PromptKind.None)
                    return;

                Ask(prompt);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt look: " + ex.Message);
                StopTimer();
            }
        }

        private static void StopTimer()
        {
            var timer = _timer;
            _timer = null;
            if (timer == null)
                return;

            try
            {
                timer.Stop();
                timer.Tick -= OnLook;
                timer.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt timer: " + ex.Message);
            }
        }

        // ===== The question =====

        private static void Ask(McpRegistrationDecision.PromptKind prompt)
        {
            DialogResult answer;
            try
            {
                answer = MessageBox.Show(
                    OwnerWindow(),
                    BuildText(prompt),
                    "OutlookAI",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    // "No" is the default in all four questions, and for one reason: in every
                    // one of them it is the answer that leaves the user's Claude Code
                    // configuration exactly as it is. A stray Enter must not edit a file the
                    // user did not come here to edit.
                    MessageBoxDefaultButton.Button2);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt show: " + ex.Message);
                return;
            }

            // Dismissed rather than answered: nothing is recorded and nothing is written, and
            // it is not asked again until Outlook next starts.
            if (answer != DialogResult.Yes && answer != DialogResult.No)
                return;

            bool saidYes = answer == DialogResult.Yes;
            bool register = YesMeansRegister(prompt) ? saidYes : !saidYes;

            try
            {
                McpRegistrationService.ApplyUserChoice(register);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt apply: " + ex.Message);
            }

            // Keep an open settings dialog honest about what just changed.
            SettingsDialog.RefreshIfOpen();
        }

        /// <summary>
        /// Whether "Yes" means "register it". True for the three questions that offer to put
        /// the entry in; false for the one that offers to take it out, where Yes is "remove
        /// it" and the setting stays off.
        /// </summary>
        private static bool YesMeansRegister(McpRegistrationDecision.PromptKind prompt)
        {
            return prompt != McpRegistrationDecision.PromptKind.EntryUnexpected;
        }

        private static string BuildText(McpRegistrationDecision.PromptKind prompt)
        {
            var sb = new StringBuilder();

            switch (prompt)
            {
                case McpRegistrationDecision.PromptKind.ForeignEntry:
                    sb.AppendLine("Claude Code's personal configuration already has a server called \"outlookai\" in "
                                  + "it, and it does not point at the mail server OutlookAI installed:");
                    sb.AppendLine();
                    sb.AppendLine("    " + RegisteredCommand());
                    sb.AppendLine();
                    sb.AppendLine("OutlookAI has not touched it. Should it take that entry over?");
                    sb.AppendLine();
                    sb.AppendLine("Yes  -  replace it with OutlookAI's mail server, in all your projects.");
                    sb.AppendLine("No   -  leave it exactly as it is.");
                    break;

                case McpRegistrationDecision.PromptKind.EntryMissing:
                    sb.AppendLine("OutlookAI's mail server is set to be available in all your Claude Code projects, "
                                  + "but its entry is no longer in Claude Code's personal configuration - something "
                                  + "outside Outlook removed it.");
                    sb.AppendLine();
                    sb.AppendLine("Nothing has been changed. What would you like?");
                    sb.AppendLine();
                    sb.AppendLine("Yes  -  register it again.");
                    sb.AppendLine("No   -  turn the setting off and leave the configuration alone.");
                    break;

                case McpRegistrationDecision.PromptKind.EntryUnexpected:
                    sb.AppendLine("OutlookAI's mail server is registered in Claude Code's personal configuration, but "
                                  + "the \"all my projects\" setting is off - something outside Outlook added it.");
                    sb.AppendLine();
                    sb.AppendLine("Nothing has been changed. What would you like?");
                    sb.AppendLine();
                    sb.AppendLine("Yes  -  remove the entry, and keep the setting off.");
                    sb.AppendLine("No   -  keep the entry, and turn the setting on.");
                    break;

                default:
                    sb.AppendLine("OutlookAI can let Claude Code search, read and draft your Outlook mail. To do that, "
                                  + "its mail server has to be registered in Claude Code's personal configuration.");
                    sb.AppendLine();
                    sb.AppendLine("Nothing has been registered yet. Would you like that now?");
                    sb.AppendLine();
                    sb.AppendLine("Yes  -  make the mail server available in all your Claude Code projects.");
                    sb.AppendLine("No   -  leave Claude Code's configuration alone. You can still add it to individual "
                                  + "projects later.");
                    break;
            }

            sb.AppendLine();
            sb.Append("You can change this whenever you like in OutlookAI Settings, on the Mail ribbon.");
            return sb.ToString();
        }

        /// <summary>What is registered right now, as recorded by the reconcile that asked.</summary>
        private static string RegisteredCommand()
        {
            try
            {
                string command = McpRegistrationService.GetSnapshot().RegisteredCommand;
                return string.IsNullOrEmpty(command) ? "(an entry with no command)" : command;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt command: " + ex.Message);
                return "(an entry that could not be read)";
            }
        }

        // ===== Is anyone there? =====

        /// <summary>
        /// Whether this Outlook has a window a human could see. The add-in runs INSIDE
        /// OUTLOOK.EXE, so it asks about its own process rather than hunting for one — and it
        /// applies the same test the mail server's compose surface uses
        /// (<c>ComposeSurface.CountUserVisibleWindows</c>): Win32 <c>IsWindowVisible</c>, which
        /// stays true for a minimized or DWM-cloaked window, minus windows collapsed to nothing
        /// and minus ones parked in the off-screen corner the invisible compose surface uses.
        /// </summary>
        private static bool AnyoneIsLooking()
        {
            try
            {
                return McpRegistrationDecision.AnyWindowAHumanCanSee(SnapshotOwnWindows());
            }
            catch (Exception ex)
            {
                // Unable to tell: treat it as "nobody", which defers rather than interrupts.
                System.Diagnostics.Debug.WriteLine("MCP prompt visibility: " + ex.Message);
                return false;
            }
        }

        private static List<McpRegistrationDecision.OutlookWindow> SnapshotOwnWindows()
        {
            var windows = new List<McpRegistrationDecision.OutlookWindow>();
            uint own = NativeMethods.GetCurrentProcessId();

            // Held in a local for the duration of the (synchronous) call so the delegate
            // cannot be collected while native code is calling it back.
            NativeMethods.EnumWindowsProc callback = (handle, lparam) =>
            {
                uint pid;
                NativeMethods.GetWindowThreadProcessId(handle, out pid);
                if (pid == own)
                {
                    NativeMethods.RECT r;
                    if (!NativeMethods.GetWindowRect(handle, out r))
                        r = new NativeMethods.RECT();

                    windows.Add(new McpRegistrationDecision.OutlookWindow(
                        NativeMethods.IsWindowVisible(handle),
                        NativeMethods.IsIconic(handle),
                        r.Left,
                        r.Top,
                        r.Right,
                        r.Bottom));
                }

                return true;
            };

            NativeMethods.EnumWindows(callback, IntPtr.Zero);
            GC.KeepAlive(callback);
            return windows;
        }

        /// <summary>
        /// Outlook's own main window, so the question appears in front of it rather than
        /// behind it. Null when there is none — the same <c>MainWindowHandle</c> probe the mail
        /// server uses to call an Outlook headless, and a null owner is harmless anyway.
        /// </summary>
        private static IWin32Window OwnerWindow()
        {
            try
            {
                IntPtr main;
                using (var self = System.Diagnostics.Process.GetCurrentProcess())
                {
                    main = self.MainWindowHandle;
                }

                return main == IntPtr.Zero ? null : new WindowHandle(main);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt owner: " + ex.Message);
                return null;
            }
        }

        private sealed class WindowHandle : IWin32Window
        {
            private readonly IntPtr _handle;

            internal WindowHandle(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle
            {
                get { return _handle; }
            }
        }

        private static class NativeMethods
        {
            internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            [StructLayout(LayoutKind.Sequential)]
            internal struct RECT
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

            [DllImport("user32.dll")]
            internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsWindowVisible(IntPtr hWnd);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsIconic(IntPtr hWnd);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

            [DllImport("kernel32.dll")]
            internal static extern uint GetCurrentProcessId();
        }
    }
}
