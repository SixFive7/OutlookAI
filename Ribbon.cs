using System;
using System.IO;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookAI
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            {
                return GetResourceText("OutlookAI.Ribbon.xml");
            }
            return null;
        }

        public void OnAIAssistantClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowTaskPane();
        }

        private static string GetResourceText(string resourceName)
        {
            using (Stream stream = typeof(Ribbon).Assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
