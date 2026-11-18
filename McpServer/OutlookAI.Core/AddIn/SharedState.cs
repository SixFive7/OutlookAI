using System;
using System.IO;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// OutlookAI's own per-user folder, <c>%LOCALAPPDATA%\OutlookAI</c> (v3.MD section 0.5.2), which the
    /// mail server and the add-in share: the audit log, the Outlook session registry, Outlook's start
    /// log, attachment downloads, signature backups, the add-in's temporary HTML and the server's private
    /// .NET runtime all live under it. Each of them built the path for itself until the code-sharing
    /// refactor's stage S6 (Q154).
    /// </summary>
    public static class SharedState
    {
        /// <summary>The folder: <c>%LOCALAPPDATA%\OutlookAI</c> of the user this process runs as.</summary>
        public static string Root =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI");
    }
}
