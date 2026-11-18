using System;
using System.Windows.Forms;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// Hands work to a control's UI thread without waiting for it (survey group A12): the one way the
    /// pane and OutlookAI Settings take an event that may arrive on another thread - ThemeService raises
    /// its change from SystemEvents and its registry watcher, UpdateService from its poll and the worker a
    /// manual check runs on, PromptStore from whoever saved - and touch their controls. Fire-and-forget,
    /// so a busy UI thread cannot stall the thread that raised the event, and a control already gone, or
    /// going, is simply skipped. Five handlers spelled this out for themselves until the code-sharing
    /// refactor's stage S4.
    /// </summary>
    internal static class UiThread
    {
        /// <summary>
        /// Queues <paramref name="work"/> on <paramref name="control"/>'s UI thread, unless the control is
        /// disposed, being disposed (<paramref name="disposing"/>, the owner's own flag), or has no handle.
        /// </summary>
        internal static void Post(Control control, bool disposing, Action work)
        {
            if (disposing || control.IsDisposed || !control.IsHandleCreated)
                return;
            try { control.BeginInvoke(work); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }
    }
}
