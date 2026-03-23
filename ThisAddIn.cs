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
        private Outlook.Inspectors _inspectors;
        private Outlook.Explorers _explorers;
        private readonly List<Outlook.Explorer> _hookedExplorers = new List<Outlook.Explorer>();

        internal Microsoft.Office.Core.IRibbonUI RibbonUI { get; set; }

        internal static void ReleaseCom(object obj)
        {
            if (obj != null)
            {
                try { Marshal.ReleaseComObject(obj); } catch { }
            }
        }

        internal CustomTaskPane FindPaneForWindow(object window)
        {
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                object paneWindow = pane.Window;
                bool match = (paneWindow == window);
                ReleaseCom(paneWindow);
                if (match)
                    return pane;
            }
            return null;
        }

        private void InvalidateRibbonToggle()
        {
            try { RibbonUI?.InvalidateControl("btnAICompose"); } catch { }
            try { RibbonUI?.InvalidateControl("btnAIInline"); } catch { }
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            InvalidateRibbonToggle();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ClaudeService.WarmUp();
            UpdateService.Start();

            // Auto-show task pane when a compose inspector opens
            _inspectors = this.Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;

            // Hook inline response events on all current and future Explorer windows
            try
            {
                _explorers = this.Application.Explorers;
                _explorers.NewExplorer += Explorers_NewExplorer;

                var explorer = this.Application.ActiveExplorer();
                if (explorer != null)
                    HookExplorer(explorer);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Explorer event hookup failed: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            UpdateService.Stop();
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
                catch { }
                ReleaseCom(explorer);
            }
            _hookedExplorers.Clear();

            ReleaseCom(_inspectors);
            _inspectors = null;
            ReleaseCom(_explorers);
            _explorers = null;
        }

        private void HookExplorer(Outlook.Explorer explorer)
        {
            try
            {
                ((Outlook.ExplorerEvents_10_Event)explorer).InlineResponse += Explorer_InlineResponse;
                ((Outlook.ExplorerEvents_10_Event)explorer).InlineResponseClose += Explorer_InlineResponseClose;
                _hookedExplorers.Add(explorer);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("HookExplorer failed: " + ex.Message);
                ReleaseCom(explorer);
            }
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            HookExplorer(explorer);
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

                activateHandler = () =>
                {
                    ShowTaskPaneForInspector(inspector);
                    // Only auto-show once per inspector lifecycle.
                    // The user can toggle via the ribbon button after this.
                    events.Activate -= activateHandler;
                };
                closeHandler = () =>
                {
                    // Release remaining event subscriptions so the COM
                    // RCW can be garbage-collected.
                    events.Activate -= activateHandler;
                    events.Close -= closeHandler;
                    ReleaseCom(inspector);
                };

                events.Activate += activateHandler;
                events.Close += closeHandler;
            }
            catch
            {
                // Fallback: try creating the task pane immediately.
                // No Close handler was set up, so release the inspector
                // unless ShowTaskPaneForInspector stored it in a task pane
                // (AITaskPane.DisposeCustomResources handles that case).
                ShowTaskPaneForInspector(inspector);
                ReleaseCom(inspector);
            }
        }

        private void ShowTaskPaneForInspector(Outlook.Inspector inspector)
        {
            Outlook.MailItem mailItem = null;
            try
            {
                // If a pane already exists for this inspector (Outlook recycles
                // Inspector objects), ensure it is visible for the new composition.
                var existingPane = FindPaneForWindow(inspector);
                if (existingPane != null)
                {
                    existingPane.Visible = true;
                    return;
                }

                mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem == null || mailItem.Sent)
                    return;

                var taskPaneControl = new AITaskPane(isInlineResponse: false, inspector: inspector);
                var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", inspector);
                customTaskPane.Width = 280;
                customTaskPane.VisibleChanged += TaskPane_VisibleChanged;
                customTaskPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ShowTaskPaneForInspector error: " + ex.Message);
            }
            finally
            {
                ReleaseCom(mailItem);
            }
        }

        private void Explorer_InlineResponse(object item)
        {
            Outlook.Explorer explorer = null;
            try
            {
                if (!(item is Outlook.MailItem))
                    return;

                // Find the explorer that owns this inline response
                explorer = this.Application.ActiveExplorer();
                if (explorer == null)
                    return;

                // Reuse existing explorer task pane if one was already created
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
                customTaskPane.Width = 280;
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
                ReleaseCom(item);
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

        /// <summary>
        /// Toggles the AI Assistant task pane for the given ribbon context
        /// (Inspector or Explorer window).  If a pane already exists for
        /// the window its visibility is toggled; otherwise a new pane is created.
        /// </summary>
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

                // No pane exists yet — create one.
                var asInspector = context as Outlook.Inspector;
                var asExplorer = context as Outlook.Explorer;

                if (asInspector != null)
                {
                    Outlook.MailItem mailItem = null;
                    try
                    {
                        mailItem = asInspector.CurrentItem as Outlook.MailItem;
                        if (mailItem != null && !mailItem.Sent)
                        {
                            var taskPaneControl = new AITaskPane(isInlineResponse: false, inspector: asInspector);
                            var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", asInspector);
                            customTaskPane.Width = 280;
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
                        inlineItem = asExplorer.ActiveInlineResponse as Outlook.MailItem;
                        if (inlineItem != null)
                        {
                            var taskPaneControl = new AITaskPane(isInlineResponse: true);
                            var customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "AI Assistant", asExplorer);
                            customTaskPane.Width = 280;
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
