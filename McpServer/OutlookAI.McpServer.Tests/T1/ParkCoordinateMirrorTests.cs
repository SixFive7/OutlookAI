using OutlookAI.Core.Com;
using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// THE ONE CROSS-ASSEMBLY CONSTANT, AND THE ONLY PLACE ANYTHING CHECKS IT.
///
/// <c>ComposeSurface</c> (net10.0, mail server) parks a window it must keep off a human's
/// screen at <c>(ParkX, ParkY)</c>. <c>McpRegistrationDecision</c> (net48, add-in) has to
/// recognise that same corner, because "is there a window a human could be asked a question
/// in?" is exactly the question of whether every visible window is a parked one. The add-in
/// cannot reference the server assembly - it is net48 and the server is .NET 10 - so the
/// constant is MIRRORED, and the add-in's copy said so in a comment and left it there.
///
/// A comment is not a mechanism. This file is: the test project references the server
/// assembly AND compiles the add-in's decision file as a linked source, so it is the one
/// compilation in the repository that can see both numbers at once.
///
/// <para>
/// What drift costs, in the two directions:
/// </para>
/// <list type="bullet">
/// <item><description>Server parks FURTHER out than the add-in expects: the add-in's
/// half-way test still catches it, so nothing breaks - until the add-in's copy is the one
/// that moves.</description></item>
/// <item><description>Server parks NEARER in than the add-in's threshold: a parked window
/// reads as an ordinary one, so a modal question is put in front of a headless Outlook
/// nobody can see - invisible to everyone, and it wedges the reconcile that raised
/// it.</description></item>
/// <item><description>Add-in parks its threshold further out than the server ever parks:
/// every window reads as parked, the question is never asked at all, and MCP registration
/// silently never gets set up.</description></item>
/// </list>
/// </summary>
public class ParkCoordinateMirrorTests
{
    [Fact]
    public void AddInParkCoordinateMatchesTheServerConstantItMirrors()
    {
        Assert.Equal(ComposeSurface.ParkX, McpRegistrationDecision.ParkX);
        Assert.Equal(ComposeSurface.ParkY, McpRegistrationDecision.ParkY);
    }

    /// <summary>
    /// The detection threshold is <c>ParkX / 2</c> on both sides - each derives it from its
    /// own copy of the park coordinate rather than writing -16000 down, so the two stay
    /// related even while the value moves. Pinned because the halving is the part a reader
    /// is most likely to "simplify" into a literal.
    /// </summary>
    [Fact]
    public void AWindowParkedByTheServerIsNotOneTheAddInWouldAskAQuestionIn()
    {
        var parked = new McpRegistrationDecision.OutlookWindow(
            visible: true,
            minimized: false,
            left: ComposeSurface.ParkX,
            top: ComposeSurface.ParkY,
            right: ComposeSurface.ParkX + 1100,
            bottom: ComposeSurface.ParkY + 800);

        Assert.False(McpRegistrationDecision.AnyWindowAHumanCanSee(new[] { parked }));
    }

    /// <summary>
    /// The other half of the same rule, so the test above cannot be satisfied by a predicate
    /// that simply always says "no": an ordinary on-screen window of the same size still
    /// counts as somewhere a question can be put.
    /// </summary>
    [Fact]
    public void AnOrdinaryWindowStillCountsAsSomewhereAQuestionCanBeAsked()
    {
        var onScreen = new McpRegistrationDecision.OutlookWindow(
            visible: true, minimized: false, left: 100, top: 100, right: 1200, bottom: 900);

        Assert.True(McpRegistrationDecision.AnyWindowAHumanCanSee(new[] { onScreen }));
    }
}
