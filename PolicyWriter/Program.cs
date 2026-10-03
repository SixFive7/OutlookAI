using System;
using OutlookAI.Services;

namespace OutlookAI.PolicyWriter
{
    /// <summary>
    /// OutlookAI.PolicyWriter.exe. Everything it does is <see cref="PolicyWriterRun.Run"/>; this only
    /// makes sure that whatever happens, the process ends with one of <see cref="PolicyWriterExit"/>'s
    /// codes, which is all the add-in can read back from a process it started through UAC.
    /// </summary>
    internal static class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                return PolicyWriterRun.Run(args, new WindowsPolicyHost(), Console.Error);
            }
            catch (Exception ex)
            {
                try
                {
                    Console.Error.WriteLine("FAILED: " + ex.GetType().Name + ": " + ex.Message);
                }
                catch (Exception)
                {
                    // No console to report to; the exit code still says it.
                }
                return PolicyWriterExit.Failed;
            }
        }
    }
}
