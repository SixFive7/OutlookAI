using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private readonly List<Office.IRibbonUI> _ribbonUIs = new List<Office.IRibbonUI>();

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Explorer")
                return GetResourceText("OutlookAI.Ribbon.xml");

            return string.Empty;
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbonUIs.Add(ribbonUI);
            Globals.ThisAddIn.RibbonUI = ribbonUI;
        }

        internal void InvalidateAll(string controlId)
        {
            foreach (var ui in _ribbonUIs)
            {
                try { ui.InvalidateControl(controlId); } catch { }
            }
        }

        public void OnAIAssistantToggle(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.ToggleTaskPane(control.Context);
        }

        public bool GetAIAssistantPressed(Office.IRibbonControl control)
        {
            var pane = Globals.ThisAddIn.FindPaneForWindow(control.Context);
            return pane != null && pane.Visible;
        }

        private static string GetResourceText(string resourceName)
        {
            using (Stream stream = typeof(Ribbon).Assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return string.Empty;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}
