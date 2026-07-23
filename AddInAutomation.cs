using System;
using System.Runtime.InteropServices;

namespace OutlookAI
{
    /// <summary>
    /// COM automation surface exposed via COMAddIn.Object (RequestComAddInAutomationService).
    /// Lets out-of-process callers (unattended verification, future tooling) open and close
    /// the OutlookAI settings dialog and read tuning state without driving the ribbon UI.
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
    }
}
