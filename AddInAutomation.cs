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
        public void OpenSettings()
        {
            TaskPane.SettingsDialog.ShowSettings();
        }

        public void CloseSettings()
        {
            TaskPane.SettingsDialog.CloseIfOpen();
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
