namespace OutlookAI.RuntimeSetup
{
    /// <summary>
    /// Setup runs this after it has copied its files, hidden, on every install and every silent
    /// update, and reads the exit code (<see cref="PrivateRuntime.ExitCode"/>): 0 when the server's
    /// runtime is in place - it was already, or it is now - and anything else when it is not, with the
    /// log saying why. One optional argument, <c>--from &lt;folder&gt;</c>: a folder - setup passes its
    /// own - where Microsoft's runtime package may already be, under the name Microsoft gives it
    /// (<see cref="PrivateRuntime.PackageFileName"/>); it is held to the same length and SHA-512 as a
    /// download would be. Everything else is <see cref="PrivateRuntime"/>'s: this reads the arguments
    /// and nothing more.
    /// </summary>
    internal static class Program
    {
        private static int Main(string[] args)
        {
            string? from = null;
            if (args.Length == 2 && args[0] == "--from" && args[1].Length > 0)
                from = args[1];
            else if (args.Length != 0)
                return (int)PrivateRuntime.ExitCode.BadArguments;

            return (int)PrivateRuntime.Run(from);
        }
    }
}
