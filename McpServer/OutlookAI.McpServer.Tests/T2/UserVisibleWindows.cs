using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// What <see cref="ComposeSurface.CountUserVisibleWindows"/> counted, window by window, so that a test names
/// the windows it found on screen as it started and those it left: each one's class (rctrl_renwnd32 is
/// Outlook's own - an Explorer or an Inspector, a compose window among them), its title with anything but
/// the window's kind reduced to the test tag and a hash - never a subject - its process, its rectangle,
/// whether DWM cloaks it or it is minimised, and when it was read. The windows come from
/// <see cref="ComposeSurface.UserVisibleWindowHandles"/>, the count's own filter; each listing says when
/// the count taken a moment before and the list disagree.
/// </summary>
internal static class UserVisibleWindows
{
    /// <summary>One counted window, as read.</summary>
    internal sealed record Window(
        IntPtr Handle,
        string ClassName,
        string Title,
        int ProcessId,
        int Left,
        int Top,
        int Right,
        int Bottom,
        bool Cloaked,
        bool Minimized,
        DateTime ReadUtc);

    /// <summary>The windows the count counts now, each read at once.</summary>
    internal static IReadOnlyList<Window> Take()
    {
        DateTime at = DateTime.UtcNow;
        List<Window> windows = new List<Window>();
        foreach (IntPtr handle in ComposeSurface.UserVisibleWindowHandles())
        {
            windows.Add(Read(handle, at));
        }

        return windows;
    }

    /// <summary>
    /// The windows at one moment, one line each, with the count taken just before them: what a test found
    /// on screen as it started. Windows an earlier test left can already be showing then, and a test whose
    /// count goes from 3 to 3 passes without naming them. Starts with a line break, as
    /// <see cref="Describe"/> does.
    /// </summary>
    internal static string List(string moment, IReadOnlyList<Window> windows, int counted)
    {
        StringBuilder text = new StringBuilder();
        text.Append(CultureInfo.InvariantCulture, $"{Environment.NewLine}user-visible windows {moment}: {windows.Count} listed ({counted} counted)");
        if (windows.Count != counted)
        {
            text.Append($"{Environment.NewLine}  NOTE: the count and the list taken just after it disagree - a window came or went between them, or the two filters have drifted apart");
        }

        foreach (Window w in windows)
        {
            text.Append(Environment.NewLine);
            text.Append("  ");
            text.Append(Format(w));
        }

        return text.ToString();
    }

    /// <summary>
    /// The windows after a step, one line each, those whose handle was not there before marked NEW - with
    /// the two counts beside them and a note when a count and its list disagree. Starts with a line break,
    /// to follow an assertion's own message.
    /// </summary>
    internal static string Describe(IReadOnlyList<Window> before, IReadOnlyList<Window> after, int countBefore, int countAfter)
    {
        StringBuilder text = new StringBuilder();
        text.Append(CultureInfo.InvariantCulture, $"{Environment.NewLine}user-visible windows before: {before.Count} listed ({countBefore} counted); after: {after.Count} listed ({countAfter} counted)");
        if (before.Count != countBefore || after.Count != countAfter)
        {
            text.Append($"{Environment.NewLine}  NOTE: a count and the list taken just after it disagree - a window came or went between them, or the two filters have drifted apart");
        }

        HashSet<IntPtr> known = new HashSet<IntPtr>(before.Select(w => w.Handle));
        foreach (Window w in after)
        {
            text.Append(Environment.NewLine);
            text.Append(known.Contains(w.Handle) ? "  was  " : "  NEW  ");
            text.Append(Format(w));
        }

        foreach (Window w in before.Where(b => !after.Any(a => a.Handle == b.Handle)))
        {
            text.Append(Environment.NewLine);
            text.Append("  gone ");
            text.Append(Format(w));
        }

        return text.ToString();
    }

    /// <summary>
    /// A title with nothing but its window's kind left readable: Outlook ends a title with the kind after
    /// the last " - " ("Message (HTML)", "Outlook"); the rest - a subject, a folder, a store - becomes the
    /// test tag when it carries one, and a hash of it either way.
    /// </summary>
    internal static string RedactTitle(string title)
    {
        if (string.IsNullOrEmpty(title))
        {
            return "(no title)";
        }

        string kind = string.Empty;
        string rest = title;
        int cut = title.LastIndexOf(" - ", StringComparison.Ordinal);
        if (cut >= 0)
        {
            string tail = title.Substring(cut + 3);
            if (Regex.IsMatch(tail, @"^[A-Za-z][A-Za-z ()]{0,39}$", RegexOptions.CultureInvariant))
            {
                kind = tail;
                rest = title.Substring(0, cut);
            }
        }

        string tag = LiveOutlookTestMailer.CarriesSubjectTag(rest) ? LiveOutlookTestMailer.SubjectTag + " " : string.Empty;
        string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(rest))).Substring(0, 8);
        string reduced = string.Create(CultureInfo.InvariantCulture, $"{tag}#{hash} ({rest.Length} chars)");
        return kind.Length == 0 ? reduced : $"{reduced} - {kind}";
    }

    private static string Format(Window w)
    {
        return string.Create(
            CultureInfo.InvariantCulture,
            $"0x{w.Handle.ToInt64():X} class={w.ClassName} title=\"{RedactTitle(w.Title)}\" pid={w.ProcessId} rect=({w.Left},{w.Top})-({w.Right},{w.Bottom}) cloaked={w.Cloaked} minimized={w.Minimized} at={w.ReadUtc:HH:mm:ss.fff}Z");
    }

    private static Window Read(IntPtr handle, DateTime at)
    {
        StringBuilder className = new StringBuilder(256);
        _ = ScreenCapture.NativeMethods.GetClassNameW(handle, className, className.Capacity);
        int length = ScreenCapture.NativeMethods.GetWindowTextLengthW(handle);
        StringBuilder title = new StringBuilder(Math.Max(length, 0) + 1);
        _ = ScreenCapture.NativeMethods.GetWindowTextW(handle, title, title.Capacity);
        _ = ScreenCapture.NativeMethods.GetWindowThreadProcessId(handle, out uint pid);
        _ = ScreenCapture.NativeMethods.GetWindowRect(handle, out ScreenCapture.NativeMethods.RECT r);
        bool cloaked = ScreenCapture.NativeMethods.DwmGetWindowAttribute(handle, ScreenCapture.NativeMethods.DWMWA_CLOAKED, out int c, sizeof(int)) == 0 && c != 0;
        bool minimized = ScreenCapture.NativeMethods.IsIconic(handle);
        return new Window(handle, className.ToString(), title.ToString(), (int)pid, r.Left, r.Top, r.Right, r.Bottom, cloaked, minimized, at);
    }
}
