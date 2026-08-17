using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using OutlookAI.Services;
using OutlookAI.TaskPane;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// The named mutex Inno Setup holds for the whole of an install, and the only signal
        /// the add-in has that it is being replaced underneath itself.
        ///
        /// THE SAME NAME IS IN Installer.iss AS SetupMutex. Renaming one side evaporates the
        /// guard with no error at all: the add-in initialises during a silent auto-update,
        /// spins up the updater and warm-up processes, and the installer tears them down
        /// mid-flight - exactly the failure this exists to prevent. The two are therefore
        /// compared mechanically by .github/scripts/check-pinned-constants.ps1 rather than
        /// left to a comment on one side.
        /// </summary>
        internal const string InstallerMutexName = "OutlookAISetup";

        private Outlook.Inspectors _inspectors;
        private Outlook.Explorers _explorers;
        private readonly List<Outlook.Explorer> _hookedExplorers = new List<Outlook.Explorer>();
        private Ribbon _ribbon;

        internal Microsoft.Office.Core.IRibbonUI RibbonUI { get; set; }

        // Created on Outlook's main UI thread during startup; lets code that may run on a
        // non-UI thread (the COMAddIn.Object automation surface — out-of-process COM calls
        // arrive on RPC worker threads) marshal UI work onto the UI thread. Never shown.
        private System.Windows.Forms.Control _uiMarshalControl;

        internal System.Windows.Forms.Control UiMarshalControl
        {
            get { return _uiMarshalControl; }
        }

        internal static void ReleaseCom(object obj)
        {
            if (obj != null)
            {
                try { Marshal.ReleaseComObject(obj); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("ReleaseCom: " + ex.Message); }
            }
        }

        private static bool ComEqual(object a, object b)
        {
            if (a == b) return true;
            if (a == null || b == null) return false;
            IntPtr punkA = IntPtr.Zero, punkB = IntPtr.Zero;
            try
            {
                punkA = Marshal.GetIUnknownForObject(a);
                punkB = Marshal.GetIUnknownForObject(b);
                return punkA == punkB;
            }
            finally
            {
                if (punkA != IntPtr.Zero) Marshal.Release(punkA);
                if (punkB != IntPtr.Zero) Marshal.Release(punkB);
            }
        }

        internal CustomTaskPane FindPaneForWindow(object window)
        {
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                object paneWindow = null;
                try
                {
                    paneWindow = pane.Window;
                    if (ComEqual(paneWindow, window))
                        return pane;
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("FindPaneForWindow: " + ex.Message); }
                finally
                {
                    ReleaseCom(paneWindow);
                }
            }
            return null;
        }

        private void InvalidateRibbonToggle()
        {
            if (_ribbon != null)
            {
                _ribbon.InvalidateAll("btnAICompose");
                _ribbon.InvalidateAll("btnAIInline");
            }
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            InvalidateRibbonToggle();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // If the installer is mid-run it holds the InstallerMutexName mutex; skip all
            // startup work. Outlook can be reopened during a silent update install, and
            // initializing here would re-trigger the updater and spin up work that the
            // installer immediately tears down to swap the add-in files. The add-in
            // loads cleanly on the next Outlook restart.
            bool installerRunning;
            try
            {
                System.Threading.Mutex mutex;
                installerRunning = System.Threading.Mutex.TryOpenExisting(InstallerMutexName, out mutex);
                mutex?.Dispose();
            }
            catch
            {
                installerRunning = false;
            }
            if (installerRunning)
                return;

            try
            {
                _uiMarshalControl = new System.Windows.Forms.Control();
                var forceHandle = _uiMarshalControl.Handle; // force creation on the UI thread
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UI marshal control: " + ex.Message);
                _uiMarshalControl = null;
            }

            ClaudeService.WarmUp();
            UpdateService.Start();
            ThemeService.StartWatching();

            // Keep the user's Outlook tuning (search / caching / OST headroom) applied.
            // Registry-only, idempotent, and never throws out of its public surface, but
            // guard anyway: tuning must never be able to break add-in startup.
            try { OutlookTuningService.ReconcileOnStartup(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("Tuning reconcile: " + ex.Message); }

            // Where the registration reconcile below sends a question it cannot answer on its
            // own. Installed BEFORE the reconcile starts, so a first run has somewhere to ask;
            // nothing is shown from here — the prompt waits for startup to settle and appears
            // only if this Outlook has a window a human can see (it deliberately does not, when
            // it was autostarted in the background for an agent session).
            try { TaskPane.McpRegistrationPrompt.Install(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP prompt install: " + ex.Message); }

            // Keep Claude Code's MCP registration pointing at the installed mail server, and
            // heal it when it drifts (a developer build output, or a path from an older
            // install). Off the UI thread, unlike the tuning reconcile above: this one reads
            // and may rewrite a file in the user profile, and Outlook startup must not wait
            // on disk. It touches no COM and no Outlook object model, so running it on a
            // worker thread does not disturb the add-in's COM ownership rules.
            try
            {
                System.Threading.Tasks.Task.Run(() =>
                {
                    try { McpRegistrationService.Reconcile(); }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP registration reconcile: " + ex.Message); }
                });
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP registration reconcile start: " + ex.Message); }

            _inspectors = this.Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;

            try
            {
                _explorers = this.Application.Explorers;
                _explorers.NewExplorer += Explorers_NewExplorer;

                var explorer = this.Application.ActiveExplorer();
                if (explorer != null && !HookExplorer(explorer))
                    ReleaseCom(explorer);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Explorer event hookup failed: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try { TaskPane.McpRegistrationPrompt.Shutdown(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP prompt shutdown: " + ex.Message); }

            try { SettingsDialog.CloseIfOpen(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("Settings close on shutdown: " + ex.Message); }

            try { _uiMarshalControl?.Dispose(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("UI marshal dispose: " + ex.Message); }
            _uiMarshalControl = null;

            UpdateService.Stop();
            ThemeService.StopWatching();
            ClaudeService.Shutdown();

            if (_inspectors != null)
                _inspectors.NewInspector -= Inspectors_NewInspector;

            if (_explorers != null)
                _explorers.NewExplorer -= Explorers_NewExplorer;

            foreach (var explorer in _hookedExplorers)
            {
                try
                {
                    ((Outlook.ExplorerEvents_10_Event)explorer).InlineResponse -= Explorer_InlineResponse;
                    ((Outlook.ExplorerEvents_10_Event)explorer).InlineResponseClose -= Explorer_InlineResponseClose;
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("Unhook explorer: " + ex.Message); }
                ReleaseCom(explorer);
            }
            _hookedExplorers.Clear();

            ReleaseCom(_inspectors);
            _inspectors = null;
            ReleaseCom(_explorers);
            _explorers = null;
        }

        private bool IsExplorerHooked(Outlook.Explorer explorer)
        {
            foreach (var hooked in _hookedExplorers)
            {
                if (ComEqual(hooked, explorer))
                    return true;
            }
            return false;
        }

        private bool HookExplorer(Outlook.Explorer explorer)
        {
            try
            {
                if (IsExplorerHooked(explorer))
                    return false;
                var events = (Outlook.ExplorerEvents_10_Event)explorer;
                events.InlineResponse += Explorer_InlineResponse;
                events.InlineResponseClose += Explorer_InlineResponseClose;

                // Mirror the Inspector close pattern: when the Explorer window closes, dispose its
                // inline pane (stops the pane's timer + unsubscribes it from theme changes), unhook
                // the events, drop it from _hookedExplorers, and release the RCW. Unlike inspector
                // panes, the inline pane does not own the explorer RCW, so we always release here.
                Outlook.ExplorerEvents_10_CloseEventHandler closeHandler = null;
                closeHandler = () =>
                {
                    try
                    {
                        events.InlineResponse -= Explorer_InlineResponse;
                        events.InlineResponseClose -= Explorer_InlineResponseClose;
                        events.Close -= closeHandler;
                        RemovePaneForWindow(explorer);
                        bool wasTracked = UnhookExplorer(explorer);
                        // Release only if we removed it from _hookedExplorers here, so we never
                        // double-release an explorer the shutdown path already released.
                        if (wasTracked)
                            ReleaseCom(explorer);
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine("Explorer close cleanup: " + ex.Message); }
                };
                events.Close += closeHandler;

                _hookedExplorers.Add(explorer);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("HookExplorer failed: " + ex.Message);
                return false;
            }
        }

        private bool UnhookExplorer(Outlook.Explorer explorer)
        {
            bool removed = false;
            for (int i = _hookedExplorers.Count - 1; i >= 0; i--)
            {
                if (ComEqual(_hookedExplorers[i], explorer))
                {
                    _hookedExplorers.RemoveAt(i);
                    removed = true;
                }
            }
            return removed;
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            // _hookedExplorers takes ownership when hooked (released at shutdown); otherwise
            // release this RCW so an already-hooked / failed-hook explorer doesn't leak.
            if (!HookExplorer(explorer))
                ReleaseCom(explorer);
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            try
            {
                var events = (Outlook.InspectorEvents_10_Event)inspector;
                bool handled = false;

                Outlook.InspectorEvents_10_ActivateEventHandler activateHandler = null;
                Outlook.InspectorEvents_10_CloseEventHandler closeHandler = null;

                activateHandler = () =>
                {
                    if (handled) return;
                    handled = true;
                    events.Activate -= activateHandler;
                    ShowTaskPaneForInspector(inspector);
                };
                closeHandler = () =>
                {
                    handled = true;
                    events.Activate -= activateHandler;
                    events.Close -= closeHandler;

                    // The pane (if present) owns the inspector RCW and releases it on dispose;
                    // release here only when no pane took ownership, to avoid over-releasing.
                    if (!RemovePaneForWindow(inspector))
                        ReleaseCom(inspector);
                };

                events.Activate += activateHandler;
                events.Close += closeHandler;
            }
            catch
            {
                if (!ShowTaskPaneForInspector(inspector))
                    ReleaseCom(inspector);
            }
        }

        private bool ShowTaskPaneForInspector(Outlook.Inspector inspector)
        {
            Outlook.MailItem mailItem = null;
            try
            {
                var existingPane = FindPaneForWindow(inspector);
                if (existingPane != null)
                {
                    existingPane.Visible = true;
                    return true;
                }

                object rawItem = inspector.CurrentItem;
                mailItem = rawItem as Outlook.MailItem;
                if (mailItem == null)
                {
                    ReleaseCom(rawItem);
                    return false;
                }
                if (mailItem.Sent)
                    return false;

                var taskPaneControl = new AITaskPane(isInlineResponse: false, inspector: inspector);
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
                customTaskPane.Width = taskPaneControl.PreferredHostWidth;
                customTaskPane.VisibleChanged += TaskPane_VisibleChanged;
                customTaskPane.Visible = true;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ShowTaskPaneForInspector error: " + ex.Message);
                return false;
            }
            finally
            {
                ReleaseCom(mailItem);
            }
        }

        private bool RemovePaneForWindow(object window)
        {
            try
            {
                var pane = FindPaneForWindow(window);
                if (pane != null)
                {
                    pane.VisibleChanged -= TaskPane_VisibleChanged;
                    this.CustomTaskPanes.Remove(pane);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("RemovePaneForWindow: " + ex.Message);
            }
            return false;
        }

        private void Explorer_InlineResponse(object item)
        {
            Outlook.Explorer explorer = null;
            try
            {
                if (!(item is Outlook.MailItem))
                    return;

                explorer = this.Application.ActiveExplorer();
                if (explorer == null)
                    return;

                var existingPane = FindPaneForWindow(explorer);
                if (existingPane != null)
                {
                    var ctrl = existingPane.Control as AITaskPane;
                    ctrl?.ResetForNewEmail();
                    existingPane.Visible = true;
                    return;
                }

                var taskPaneControl = new AITaskPane(isInlineResponse: true);
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", explorer);
                customTaskPane.Width = taskPaneControl.PreferredHostWidth;
                customTaskPane.VisibleChanged += TaskPane_VisibleChanged;
                customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InlineResponse error: " + ex.Message);
            }
            finally
            {
                ReleaseCom(explorer);
            }
        }

        private void Explorer_InlineResponseClose()
        {
            Outlook.Explorer explorer = null;
            try
            {
                explorer = this.Application.ActiveExplorer();
                if (explorer == null)
                    return;

                var pane = FindPaneForWindow(explorer);
                if (pane != null)
                    pane.Visible = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InlineResponseClose error: " + ex.Message);
            }
            finally
            {
                ReleaseCom(explorer);
            }
        }

        public void ToggleTaskPane(object context)
        {
            try
            {
                var existingPane = FindPaneForWindow(context);
                if (existingPane != null)
                {
                    existingPane.Visible = !existingPane.Visible;
                    return;
                }

                var asInspector = context as Outlook.Inspector;
                var asExplorer = context as Outlook.Explorer;

                if (asInspector != null)
                {
                    Outlook.MailItem mailItem = null;
                    try
                    {
                        object rawItem = asInspector.CurrentItem;
                        mailItem = rawItem as Outlook.MailItem;
                        if (mailItem == null)
                        {
                            ReleaseCom(rawItem);
                        }
                        else if (!mailItem.Sent)
                        {
                            var taskPaneControl = new AITaskPane(isInlineResponse: false, inspector: asInspector);
                            var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", asInspector);
                            customTaskPane.Width = taskPaneControl.PreferredHostWidth;
                            customTaskPane.VisibleChanged += TaskPane_VisibleChanged;
                            customTaskPane.Visible = true;
                        }
                    }
                    finally
                    {
                        ReleaseCom(mailItem);
                    }
                }
                else if (asExplorer != null)
                {
                    Outlook.MailItem inlineItem = null;
                    try
                    {
                        object rawInline = asExplorer.ActiveInlineResponse;
                        inlineItem = rawInline as Outlook.MailItem;
                        if (inlineItem == null)
                        {
                            ReleaseCom(rawInline);
                        }
                        else
                        {
                            var taskPaneControl = new AITaskPane(isInlineResponse: true);
                            var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", asExplorer);
                            customTaskPane.Width = taskPaneControl.PreferredHostWidth;
                            customTaskPane.VisibleChanged += TaskPane_VisibleChanged;
                            customTaskPane.Visible = true;
                        }
                    }
                    finally
                    {
                        ReleaseCom(inlineItem);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ToggleTaskPane error: " + ex.Message);
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Ribbon();
            return _ribbon;
        }

        // Exposes a small automation object on COMAddIn.Object so out-of-process callers
        // (unattended verification, future tooling) can open/close the settings dialog and
        // read tuning state without driving the ribbon UI.
        private AddInAutomation _automation;

        protected override object RequestComAddInAutomationService()
        {
            if (_automation == null)
                _automation = new AddInAutomation();
            return _automation;
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
