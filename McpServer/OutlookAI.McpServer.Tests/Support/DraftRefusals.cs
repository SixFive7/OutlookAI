using System.Globalization;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// What the refusal of a draft whose subject is too long has to say, checked once for both tiers that check it:
/// in process (<c>T1/DraftValidationTests</c>) and over stdio (<c>T3/DraftOptionsCiToolShapeTests</c>), which
/// each wrote the same five assertions out (the code-sharing refactor's S5a, the survey's T15).
/// </summary>
internal static class DraftRefusals
{
    /// <summary>The refusal <paramref name="message"/> of a subject <paramref name="supplied"/> characters long.</summary>
    internal static void AssertOverlongSubjectMessage(string message, int supplied)
    {
        // 1. It names the argument at fault - a draft call carries several strings.
        Assert.Contains("subject", message, StringComparison.OrdinalIgnoreCase);

        // 2. The LIMIT, as a number, derived from the constant the service enforces.
        Assert.Contains(
            "max " + MailService.SubjectCharsCap.ToString(CultureInfo.InvariantCulture) + " characters",
            message,
            StringComparison.Ordinal);

        // 3. What was WRONG with what was supplied, including how long it actually was.
        //    Without this the model cannot tell a near miss from an order-of-magnitude
        //    one, and cannot compute how much to cut.
        Assert.Contains(
            supplied.ToString(CultureInfo.InvariantCulture) + " characters supplied",
            message,
            StringComparison.Ordinal);

        // 4. That nothing was written, and that the SAME call with a shorter subject
        //    works - so the retry is obvious and needs no user round trip.
        Assert.Contains("Nothing was created or changed", message, StringComparison.Ordinal);
        Assert.Contains("call again", message, StringComparison.OrdinalIgnoreCase);
    }
}
