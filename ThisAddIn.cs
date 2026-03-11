using System;
using Microsoft.Office.Tools;
using OutlookAI.Services;
using OutlookAI.TaskPane;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors _inspectors;
        private Outlook.Explorer _explorer;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ClaudeService.WarmUp();

            // Auto-show task pane when a compose inspector opens
            _inspectors = this.Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;

            // Auto-show task pane for inline replies in the Explorer.
            // Wrapped in try-catch so a failure here doesn't prevent the
            // Inspector hookup above from working.
            try
            {
                _explorer = this.Application.ActiveExplorer();
                if (_explorer != null)
                {
                    ((Outlook.ExplorerEvents_10_Event)_explorer).InlineResponse += Explorer_InlineResponse;
                    ((Outlook.ExplorerEvents_10_Event)_explorer).InlineResponseClose += Explorer_InlineResponseClose;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Explorer event hookup failed: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ClaudeService.Shutdown();

            if (_inspectors != null)
                _inspectors.NewInspector -= Inspectors_NewInspector;

            if (_explorer != null)
            {
                try
                {
                    ((Outlook.ExplorerEvents_10_Event)_explorer).InlineResponse -= Explorer_InlineResponse;
                    ((Outlook.ExplorerEvents_10_Event)_explorer).InlineResponseClose -= Explorer_InlineResponseClose;
                }
                catch { }
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            try
            {
                // CurrentItem is often not available yet when NewInspector fires.
                // Defer task pane creation to the Activate event which fires once
                // the inspector window is fully loaded.
                var events = (Outlook.InspectorEvents_10_Event)inspector;

                Outlook.InspectorEvents_10_ActivateEventHandler activateHandler = null;
                Outlook.InspectorEvents_10_CloseEventHandler closeHandler = null;

                activateHandler = () => ShowTaskPaneForInspector(inspector);
                closeHandler = () =>
                {
                    // Release event subscriptions and the captured inspector
                    // reference so the COM RCW can be garbage-collected.
                    events.Activate -= activateHandler;
                    events.Close -= closeHandler;
                };

                events.Activate += activateHandler;
                events.Close += closeHandler;
            }
            catch
            {
                // Fallback: try creating the task pane immediately
                ShowTaskPaneForInspector(inspector);
            }
        }

        private void ShowTaskPaneForInspector(Outlook.Inspector inspector)
        {
            try
            {
                // Activate fires on every focus gain; only add the pane once.
                foreach (CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Window == inspector)
                        return;
                }

                if (!(inspector.CurrentItem is Outlook.MailItem mailItem) || mailItem.Sent)
                    return;

                var taskPaneControl = new AITaskPane(isInlineResponse: false, inspector: inspector);
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
                customTaskPane.Width = 280;
                customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ShowTaskPaneForInspector error: " + ex.Message);
            }
        }

        private void Explorer_InlineResponse(object item)
        {
            try
            {
                if (!(item is Outlook.MailItem))
                    return;

                // Reuse existing explorer task pane if one was already created
                foreach (CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Window == _explorer)
                    {
                        var ctrl = pane.Control as AITaskPane;
                        ctrl?.ResetForNewEmail();
                        pane.Visible = true;
                        return;
                    }
                }

                var taskPaneControl = new AITaskPane(isInlineResponse: true);
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", _explorer);
                customTaskPane.Width = 280;
                customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InlineResponse error: " + ex.Message);
            }
        }

        private void Explorer_InlineResponseClose()
        {
            try
            {
                foreach (CustomTaskPane pane in this.CustomTaskPanes)
                {
                    if (pane.Window == _explorer)
                    {
                        pane.Visible = false;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InlineResponseClose error: " + ex.Message);
            }
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
