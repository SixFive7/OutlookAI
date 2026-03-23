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
        private Office.IRibbonUI _ribbonUI;

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Explorer")
                return GetResourceText("OutlookAI.Ribbon.xml");

            return null;
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
            Globals.ThisAddIn.RibbonUI = ribbonUI;
        }

        public void OnAIAssistantToggle(Office.IRibbonControl control, bool pressed) =>
            Globals.ThisAddIn.ToggleTaskPane(control.Context);

        public bool GetAIAssistantPressed(Office.IRibbonControl control)
        {
            var context = control.Context;
            foreach (CustomTaskPane pane in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (pane.Window == context)
                    return pane.Visible;
            }
            return false;
        }

        private static string GetResourceText(string resourceName)
        {
            using (Stream stream = typeof(Ribbon).Assembly.GetManifestResourceStream(resourceName))
                return stream == null ? null : new StreamReader(stream).ReadToEnd();
        }
    }
}
