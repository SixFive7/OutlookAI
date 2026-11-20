using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    /// <summary>
    /// COM automation surface exposed via COMAddIn.Object (RequestComAddInAutomationService).
    /// Lets out-of-process callers (unattended verification, future tooling) open and close
    /// the OutlookAI settings dialog, read tuning state, and read and press a compose window's
    /// AI Assistant toggle without driving the ribbon UI.
    /// Calls marshal onto Outlook's main STA thread, where all our UI lives.
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAddInAutomation
    {
        void OpenSettings();
        void CloseSettings();
        bool IsSettingsOpen();
        bool GetRestartNeeded();

        /// <summary>
        /// The compose ribbon's AI Assistant toggle for the open compose window showing the item with
        /// <paramref name="entryId"/>, as the ribbon's own getPressed callback answers Office: whether
        /// that window's pane shows. False when no such window is open.
        /// </summary>
        bool IsAIAssistantPressed(string entryId);

        /// <summary>That window's toggle pressed, as a click presses it: its pane shown, or hidden.</summary>
        void ToggleAIAssistant(string entryId);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class AddInAutomation : IAddInAutomation
    {
        // Out-of-process COM calls to a managed object arrive on RPC worker threads, NOT on
        // Outlook's UI thread — all UI work must be marshaled via the add-in's UI-thread
        // control or the dialog would be created without a message pump (a zombie window).
        private static void OnUiThread(Action action)
        {
            var ui = Globals.ThisAddIn?.UiMarshalControl;
            try
            {
                if (ui != null && !ui.IsDisposed && ui.InvokeRequired)
                {
                    ui.Invoke(action);
                    return;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Automation marshal: " + ex.Message);
                return;
            }
            action();
        }

        public void OpenSettings()
        {
            OnUiThread(() => TaskPane.SettingsDialog.ShowSettings());
        }

        public void CloseSettings()
        {
            OnUiThread(() => TaskPane.SettingsDialog.CloseIfOpen());
        }

        public bool IsSettingsOpen()
        {
            return TaskPane.SettingsDialog.IsOpen;
        }

        public bool GetRestartNeeded()
        {
            return Services.OutlookTuningService.GetRestartNeeded();
        }

        // Both through the ribbon's own callbacks, handed the window as Office hands it - a control whose
        // Context is the inspector - so they bind exactly as Office's calls do. A callback that throws, as
        // getPressed did until K2's trace found it, leaves the answer false.
        public bool IsAIAssistantPressed(string entryId)
        {
            bool pressed = false;
            OnUiThread(() => WithInspectorShowing(entryId, inspector =>
            {
                Ribbon ribbon = Globals.ThisAddIn.RibbonCallbacks;
                pressed = ribbon != null && ribbon.GetAIAssistantPressed(new WindowOnlyRibbonControl(inspector));
            }));
            return pressed;
        }

        public void ToggleAIAssistant(string entryId)
        {
            OnUiThread(() => WithInspectorShowing(entryId, inspector =>
                Globals.ThisAddIn.RibbonCallbacks?.OnAIAssistantToggle(new WindowOnlyRibbonControl(inspector), false)));
        }

        // The open inspector whose item has entryId, handed to work; nothing happens when none is open. Every
        // reference this takes is released once, after work.
        private static void WithInspectorShowing(string entryId, Action<Outlook.Inspector> work)
        {
            if (string.IsNullOrEmpty(entryId))
                return;

            Outlook.Inspectors inspectors = Globals.ThisAddIn.Application.Inspectors;
            try
            {
                int count = inspectors.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Inspector inspector = inspectors[i];
                    object item = null;
                    try
                    {
                        item = inspector.CurrentItem;
                        var mail = item as Outlook.MailItem;
                        if (mail != null && string.Equals(mail.EntryID, entryId, StringComparison.OrdinalIgnoreCase))
                        {
                            work(inspector);
                            return;
                        }
                    }
                    finally
                    {
                        ThisAddIn.ReleaseCom(item);
                        ThisAddIn.ReleaseCom(inspector);
                    }
                }
            }
            finally
            {
                ThisAddIn.ReleaseCom(inspectors);
            }
        }
    }

    /// <summary>
    /// The ribbon control an automation call names a window through: the window as its Context, which
    /// is all the ribbon's callbacks read of Office's own.
    /// </summary>
    internal sealed class WindowOnlyRibbonControl : Microsoft.Office.Core.IRibbonControl
    {
        private readonly object _window;

        internal WindowOnlyRibbonControl(object window)
        {
            _window = window;
        }

        public string Id => string.Empty;

        public object Context => _window;

        public string Tag => string.Empty;
    }
}
