using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Soak fix 22 (D49). The window-selection rules behind the invisible compose surface.
/// <para>
/// These two functions are the SAFETY boundary of the whole feature: promotion moves and
/// hides real top-level windows of a running OUTLOOK.EXE, and the user may well be looking
/// at one of them. The rules are therefore pure and pinned here rather than left implicit
/// in the COM path, and both are expressed against Win32 <c>IsWindowVisible</c> - which is
/// TRUE for a minimized window and TRUE for a DWM-cloaked one, so "the user's Outlook is
/// minimized" can never be mistaken for "this window is mine to hide".
/// </para>
/// </summary>
public class ComposeSurfaceTests
{
    private static ComposeSurface.WindowState W(int handle, bool visible)
        => new ComposeSurface.WindowState(new IntPtr(handle), visible);

    [Fact]
    public void Park_TakesEveryInvisibleWindow_AndNoVisibleOne()
    {
        IReadOnlyList<IntPtr> parked = ComposeSurface.SelectWindowsToPark(new[]
        {
            W(1, visible: false),
            W(2, visible: true),
            W(3, visible: false),
        });

        Assert.Equal(new[] { new IntPtr(1), new IntPtr(3) }, parked);
    }

    [Fact]
    public void Park_NeverTouchesAWindowTheUserCanSee()
    {
        // A minimized or cloaked Outlook still reports IsWindowVisible == true, so it
        // arrives here as visible - and must be left completely alone. Hiding it would
        // take the user's own Outlook off their taskbar.
        Assert.Empty(ComposeSurface.SelectWindowsToPark(new[] { W(10, true), W(11, true) }));
    }

    [Fact]
    public void Park_ToleratesNoWindowsAtAll()
    {
        Assert.Empty(ComposeSurface.SelectWindowsToPark(Array.Empty<ComposeSurface.WindowState>()));
        Assert.Empty(ComposeSurface.SelectWindowsToPark(null!));
    }

    [Fact]
    public void Hide_TakesOnlyWhatBecameVisibleDuringThePromotion()
    {
        var before = new[] { W(1, false), W(2, true) };
        var after = new[] { W(1, true), W(2, true) };

        Assert.Equal(new[] { new IntPtr(1) }, ComposeSurface.SelectWindowsToHide(before, after));
    }

    [Fact]
    public void Hide_CoversAWindowTheActivationCreated()
    {
        // Activate may materialise a window that did not exist at entry. The rule is a set
        // difference over VISIBLE handles precisely so a brand-new handle is covered.
        var before = new[] { W(2, true) };
        var after = new[] { W(2, true), W(99, true) };

        Assert.Equal(new[] { new IntPtr(99) }, ComposeSurface.SelectWindowsToHide(before, after));
    }

    [Fact]
    public void Hide_NeverTakesAWindowTheUserAlreadyHadOpen()
    {
        // The whole windowed-Outlook case: every window was visible before, so promotion
        // has nothing to hide and the user's session is untouched.
        var before = new[] { W(1, true), W(2, true), W(3, true) };
        var after = new[] { W(1, true), W(2, true), W(3, true) };

        Assert.Empty(ComposeSurface.SelectWindowsToHide(before, after));
    }

    [Fact]
    public void Hide_IgnoresWindowsThatWentAwayOrStayedInvisible()
    {
        var before = new[] { W(1, false), W(2, true) };
        var after = new[] { W(1, false) }; // 2 closed, 1 still invisible

        Assert.Empty(ComposeSurface.SelectWindowsToHide(before, after));
    }

    [Fact]
    public void Hide_ToleratesMissingSnapshots()
    {
        Assert.Empty(ComposeSurface.SelectWindowsToHide(null!, null!));
        Assert.Equal(
            new[] { new IntPtr(5) },
            ComposeSurface.SelectWindowsToHide(null!, new[] { W(5, true) }));
    }

    [Fact]
    public void ParkCoordinatesAreOffAnyRealVirtualScreen()
    {
        // Parking is what stops Activate() from painting the compose window where a human
        // sees it. -32000 is the classic "minimized/off-screen" corner and is far outside
        // any plausible monitor arrangement, including negative secondary displays.
        Assert.True(ComposeSurface.ParkX <= -32000);
        Assert.True(ComposeSurface.ParkY <= -32000);
    }

    [Fact]
    public void ParkAndHide_ArePartitioned_SoNoWindowIsBothOwnedAndUserOwned()
    {
        // Property check across every arrangement of three windows: a handle selected for
        // parking is by construction invisible at entry, and a handle selected for hiding
        // is by construction NOT visible at entry - so the two rules can never disagree
        // about who owns a window, which is the invariant the safety argument rests on.
        for (int mask = 0; mask < 8; mask++)
        {
            var before = new[]
            {
                W(1, (mask & 1) != 0),
                W(2, (mask & 2) != 0),
                W(3, (mask & 4) != 0),
            };
            var after = new[] { W(1, true), W(2, true), W(3, true) };

            HashSet<IntPtr> visibleBefore = new HashSet<IntPtr>(
                before.Where(w => w.Visible).Select(w => w.Handle));

            foreach (IntPtr h in ComposeSurface.SelectWindowsToPark(before))
            {
                Assert.DoesNotContain(h, visibleBefore);
            }

            foreach (IntPtr h in ComposeSurface.SelectWindowsToHide(before, after))
            {
                Assert.DoesNotContain(h, visibleBefore);
            }
        }
    }
}
