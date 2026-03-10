using System;
using Microsoft.Office.Tools;
using OutlookAI.Services;
using OutlookAI.TaskPane;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Pre-warm a Claude CLI process so the first request is fast
            ClaudeService.WarmUp();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Kill any idle pre-warmed process
            ClaudeService.Shutdown();
        }

        public void ShowTaskPane()
        {
            try
            {
                var inspector = this.Application.ActiveInspector();
                if (inspector == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please open an email first.",
                        "AI Assistant",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // Check if task pane already exists for this inspector.
                // When re-showing a hidden pane, reset it so stale results from
                // a previous email don't bleed into the new session.
                foreach (CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Window == inspector)
                    {
                        if (!pane.Visible)
                        {
                            var existingControl = pane.Control as AITaskPane;
                            existingControl?.ResetForNewEmail();
                        }
                        pane.Visible = !pane.Visible;
                        return;
                    }
                }

                // Create new task pane
                var taskPaneControl = new AITaskPane();
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
                customTaskPane.Width = 280;
                customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error: {ex.Message}",
                    "AI Assistant",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
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