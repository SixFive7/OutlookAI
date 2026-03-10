using System.IO;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookAI
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID) =>
            ribbonID == "Microsoft.Outlook.Mail.Compose" ? GetResourceText("OutlookAI.Ribbon.xml") : null;

        public void OnAIAssistantClick(Office.IRibbonControl control) =>
            Globals.ThisAddIn.ShowTaskPane();

        private static string GetResourceText(string resourceName)
        {
            using (Stream stream = typeof(Ribbon).Assembly.GetManifestResourceStream(resourceName))
                return stream == null ? null : new StreamReader(stream).ReadToEnd();
        }
    }
}
