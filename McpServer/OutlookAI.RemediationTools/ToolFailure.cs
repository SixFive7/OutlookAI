namespace OutlookAI.RemediationTools;

/// <summary>
/// How the tool's last-resort FATAL line names a failure: the whole chain of inner exceptions, not
/// only the outermost. Added 2026-09-27 on OAI-UNINDEXED, where <c>corpus-probe</c> ended in
/// <c>FATAL: InvalidOperationException: corpus undated probe failed.</c> and nothing else - the STA
/// runner wraps whatever the session threw (<see cref="ComStaRunner"/>), and the cause, which was the
/// only thing worth reading, was dropped on the floor. Pure, so T1 pins it.
/// </summary>
public static class ToolFailure
{
    /// <summary>How deep the chain is followed before it is cut, so a cycle cannot loop.</summary>
    public const int MaxDepth = 8;

    /// <summary>
    /// <c>Type: message</c> for the exception and each inner one in turn, joined by
    /// <c>" &lt;- caused by: "</c>; an <see cref="AggregateException"/>'s first inner exception is
    /// followed like any other.
    /// </summary>
    public static string Describe(Exception ex)
    {
        ArgumentNullException.ThrowIfNull(ex);
        var parts = new List<string>();
        Exception? current = ex;
        for (int depth = 0; current != null && depth < MaxDepth; depth++)
        {
            parts.Add(current.GetType().Name + ": " + current.Message);
            current = current is AggregateException aggregate && aggregate.InnerExceptions.Count > 0
                ? aggregate.InnerExceptions[0]
                : current.InnerException;
        }

        if (current != null)
        {
            parts.Add("...");
        }

        return string.Join(" <- caused by: ", parts);
    }
}
