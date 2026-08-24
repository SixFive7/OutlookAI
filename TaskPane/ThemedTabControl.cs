using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// The one bit of arithmetic every hand-laid-out surface in this add-in shares: a 96-DPI
    /// design pixel turned into a real one, using the FONT the control actually got rather than
    /// any DPI the process happens to believe it is running at.
    ///
    /// Why the font and not the DPI: these forms set <see cref="Form.AutoScaleMode"/> to None
    /// on purpose - letting WinForms scale on top of measured layout applies the display scaling
    /// twice - so the only honest signal that a display is at 125% or 150% is that Segoe UI 9pt
    /// came back 19 or 22 pixels tall instead of 15. It is also what makes the offline render
    /// harness meaningful: setting the form's Font to 11.25pt reproduces a 125% display exactly,
    /// with no manifest, no DPI awareness and no second monitor.
    /// </summary>
    internal static class UiScale
    {
        /// <summary>Segoe UI 9pt line height at 96 DPI, which every design value is written against.</summary>
        internal const int DesignLineHeight = 15;

        internal static int ScaledFor(Font font, int designPixels)
        {
            int line = font == null ? DesignLineHeight : Math.Max(1, font.Height);
            return (int)Math.Round(designPixels * (line / (double)DesignLineHeight));
        }
    }

    /// <summary>
    /// A real <see cref="TabControl"/> that paints its own strip.
    ///
    /// WHY IT HAS TO PAINT ITSELF. A tab control's strip background, tab items and page frame
    /// are drawn by the Windows visual style, and BackColor does not reach any of them: in
    /// dark mode that is a light grey band and a light grey frame around a dark page.
    /// DrawMode.OwnerDrawFixed on its own does not fix it either - it hands over the tab
    /// ITEMS and carries on drawing the band beside the last tab and the frame round the
    /// page. So the whole client area is ours (ControlStyles.UserPaint) and every colour
    /// comes out of <see cref="ThemeService"/>, which means light and dark are the same code
    /// path and cannot drift apart.
    ///
    /// WHAT IS STILL THE REAL CONTROL: everything that is not paint. The native control lays
    /// the tabs out and answers GetTabRect, hit-tests the mouse, moves the selection on the
    /// arrow keys and on Ctrl+Tab, and reports itself to a screen reader as a PageTabList.
    /// That is exactly what a row of buttons could not do, and none of it is re-implemented
    /// here.
    ///
    /// WHAT THE PAINT SAYS. The selected tab is filled with the page's own colour, has no
    /// bottom edge, and runs a couple of pixels past the page border so the two are visibly
    /// one surface; it also carries an accent bar along its top, because fill-versus-fill is
    /// a contrast cue and a high-contrast theme can flatten it. Every other tab starts lower,
    /// is filled with a shaded button face (see <see cref="RestingFill"/>), is bounded on all
    /// four sides, and is written in secondary text - it sits behind the page rather than
    /// on it.
    ///
    /// SIZING IS THE NATIVE CONTROL'S, DELIBERATELY. <see cref="TabControl.ItemSize"/> is never
    /// set, so WinForms never sends TCM_SETITEMSIZE and each tab is exactly as wide as its own
    /// caption needs. The alternative - equal-width tabs sized to the widest caption - wastes
    /// the difference on every other tab, and five tabs of the widest caption's width overflow
    /// the window at its minimum size. An overflowing strip is not a cosmetic problem: the
    /// native control answers it with an OS-drawn up-down scroller that UserPaint does NOT
    /// cover, so it comes back light grey in dark mode. Multiline is the same trap from the
    /// other side - the paint below assumes ONE row, because <c>pageTop</c> is the bottom of
    /// the tallest tab rect. Keep the captions short and the strip stays one row.
    ///
    /// AND THE ONE THING USERPAINT BREAKS THAT NOTHING WARNS YOU ABOUT: the native control never
    /// hears about the font. <c>Control.OnHandleCreated</c> sends WM_SETFONT only when
    /// ControlStyles.UserPaint is FALSE - a user-painted control is assumed to draw its own text,
    /// which this one does. But text is not all the font decides here: the native half still
    /// MEASURES the tabs, and with no WM_SETFONT it measures them with the default shell font
    /// forever. Measured symptom, before <see cref="PushFontToNativeStrip"/> existed: identical
    /// 61x21 tab rects at 100%, 125% AND 150%, so at 150% a 25px-tall, 113px-wide "Claude Code"
    /// was painted into a 21x94 tab and came back ellipsised. Sending the font ourselves is the
    /// fix, and it is also why the fixed ItemSize this control used to carry appeared to work -
    /// it was papering over exactly this.
    /// </summary>
    internal sealed class ThemedTabControl : TabControl
    {
        private const int WM_SETFONT = 0x0030;

        /// <summary>
        /// TCM_SETPADDING, sent straight at the window. Deliberately NOT the managed
        /// <c>TabControl.Padding</c> property, whose setter runs the framework's UpdateSize() -
        /// a one-pixel Size round trip on a Dock=Fill control inside a TableLayoutPanel, which
        /// re-enters layout from inside a font change and hangs the form. Measured: the harness
        /// wedged and had to be killed. The message on its own does exactly the wanted thing.
        /// </summary>
        private const int TCM_SETPADDING = 0x1300 + 43;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr handle);

        /// <summary>The GDI font the native strip is measuring with. Ours to create and to free.</summary>
        private IntPtr _stripFont = IntPtr.Zero;

        /// <summary>Guards the size nudge below from re-entering through a font change.</summary>
        private bool _pushingFont;

        /// <summary>The tab under the pointer, or -1. Ours to track: nothing native paints.</summary>
        private int _hot = -1;

        internal ThemedTabControl()
        {
            // UserPaint is the whole trick. AllPaintingInWmPaint stops the erase-background
            // flicker that painting a strip on every resize would otherwise produce, and
            // ResizeRedraw is needed because the page frame is measured from Width/Height.
            SetStyle(ControlStyles.UserPaint
                     | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.OptimizedDoubleBuffer
                     | ControlStyles.ResizeRedraw, true);

            // Text-sized tabs, one row. See the class comment for why this is not Fixed.
            SizeMode = TabSizeMode.Normal;
            // Says what is going on to anything that inspects the control. It does NOT make the
            // tabs equal width - that is TCS_FIXEDWIDTH, which is SizeMode above.
            DrawMode = TabDrawMode.OwnerDrawFixed;
            Multiline = false;
        }

        private int Scale(int designPixels)
        {
            return UiScale.ScaledFor(Font, designPixels);
        }

        /// <summary>
        /// Hands the native strip the font it is supposed to measure tabs with. See the class
        /// comment: WinForms deliberately withholds WM_SETFONT from a UserPaint control, and this
        /// one still needs it because the tab RECTS come from the native side.
        /// </summary>
        private void PushFontToNativeStrip()
        {
            if (_pushingFont || !IsHandleCreated || Font == null)
                return;

            _pushingFont = true;
            try
            {
                IntPtr previous = _stripFont;
                IntPtr created = IntPtr.Zero;
                try
                {
                    created = Font.ToHfont();
                    _stripFont = created;
                    // Sent BEFORE the old handle is freed, so the control is never briefly
                    // pointing at a deleted GDI object.
                    SendMessage(Handle, WM_SETFONT, created, (IntPtr)1);

                    // Padding is the other half of a tab's width, and it does not scale either:
                    // WinForms sends TCM_SETPADDING once with a flat (6,3) and the native control
                    // keeps it. At 150% that left a measured ZERO pixels of slack around
                    // "Outlook" - the caption exactly filled its own tab - so it is scaled here
                    // with everything else.
                    SendMessage(Handle, TCM_SETPADDING, IntPtr.Zero, MakeLParam(Scale(6), Scale(3)));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Tab strip font: " + ex.Message);
                }
                finally
                {
                    if (previous != IntPtr.Zero && previous != created)
                    {
                        try { DeleteObject(previous); }
                        catch (Exception ex) { Debug.WriteLine("Tab strip font free: " + ex.Message); }
                    }
                }

                // The tab rects have just changed size, and TabControl caches the page area it
                // derives from them. A resize is what clears that cache, and a one-pixel round
                // trip is how the framework's own ItemSize setter forces one.
                try
                {
                    Size was = Size;
                    Size = new Size(was.Width + 1, was.Height);
                    Size = was;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Tab strip resize: " + ex.Message);
                }
            }
            finally
            {
                _pushingFont = false;
            }
        }

        private static IntPtr MakeLParam(int low, int high)
        {
            return (IntPtr)(((high & 0xFFFF) << 16) | (low & 0xFFFF));
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            PushFontToNativeStrip();
        }

        /// <summary>
        /// Where the page starts, in client pixels: the bottom of the tab row. Measured from
        /// every tab rather than assumed, and exposed so a caller can check the strip is still
        /// the single row this paint code assumes.
        /// </summary>
        internal int PageTop()
        {
            int top = 0;
            try
            {
                for (int i = 0; i < TabCount; i++)
                    top = Math.Max(top, GetTabRect(i).Bottom);
            }
            catch (ArgumentOutOfRangeException)
            {
                // The tab count changed underneath the loop.
            }
            return top;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            try
            {
                PaintStrip(e.Graphics);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Tab strip paint: " + ex.Message);
            }
        }

        private void PaintStrip(Graphics g)
        {
            Color body = ThemeService.Background;
            Color face = ThemeService.ButtonFace;
            Color edge = ThemeService.Border;

            using (var brush = new SolidBrush(body))
                g.FillRectangle(brush, ClientRectangle);

            int count = TabCount;
            if (count == 0 || Width < 2 || Height < 2)
                return;

            var rects = new Rectangle[count];
            for (int i = 0; i < count; i++)
                rects[i] = GetTabRect(i);

            int selected = SelectedIndex;

            // One row, so the row's bottom is the tallest tab's bottom. The native control
            // raises the SELECTED tab a couple of pixels in Normal size mode, which is why every
            // unselected tab is then drawn down to this line rather than to its own bottom: it
            // closes the hairline gap that difference would otherwise leave above the page.
            int pageTop = 0;
            for (int i = 0; i < count; i++)
                pageTop = Math.Max(pageTop, rects[i].Bottom);

            // The page body, bordered all the way round. The selected tab breaks into it.
            using (var pen = new Pen(edge))
                g.DrawRectangle(pen, Rectangle.FromLTRB(0, pageTop, Width - 1, Height - 1));

            for (int i = 0; i < count; i++)
            {
                if (i != selected)
                    PaintUnselectedTab(g, rects[i], pageTop, i, face, edge);
            }

            if (selected >= 0 && selected < count)
                PaintSelectedTab(g, rects[selected], pageTop, body, edge);
        }

        private void PaintUnselectedTab(Graphics g, Rectangle rect, int pageTop, int index,
                                        Color face, Color edge)
        {
            Rectangle tab = Rectangle.FromLTRB(
                rect.Left, rect.Top + Scale(3), rect.Right, Math.Max(rect.Bottom, pageTop));
            if (tab.Width < 2 || tab.Height < 2)
                return;

            Color resting = RestingFill(face);
            Color fill = index == _hot ? Blend(resting, ThemeService.Accent, HoverAccentBlend) : resting;
            using (var brush = new SolidBrush(fill))
                g.FillRectangle(brush, tab);
            using (var pen = new Pen(edge))
                g.DrawRectangle(pen, new Rectangle(tab.X, tab.Y, tab.Width - 1, tab.Height - 1));

            DrawCaption(g, index, tab, ThemeService.SecondaryText);
        }

        private void PaintSelectedTab(Graphics g, Rectangle rect, int pageTop, Color body, Color edge)
        {
            int overlap = Scale(2);
            Rectangle tab = Rectangle.FromLTRB(
                rect.Left, rect.Top, rect.Right, pageTop + overlap);
            if (tab.Width < 2 || tab.Height < 2)
                return;

            using (var brush = new SolidBrush(body))
                g.FillRectangle(brush, tab);

            // Three sides only. The missing bottom edge is what joins the tab to the page.
            using (var pen = new Pen(edge))
            {
                g.DrawLine(pen, tab.Left, tab.Top, tab.Right - 1, tab.Top);
                g.DrawLine(pen, tab.Left, tab.Top, tab.Left, tab.Bottom - 1);
                g.DrawLine(pen, tab.Right - 1, tab.Top, tab.Right - 1, tab.Bottom - 1);
            }

            int bar = Math.Max(2, Scale(2));
            using (var brush = new SolidBrush(ThemeService.Accent))
                g.FillRectangle(brush, new Rectangle(tab.Left + 1, tab.Top + 1, tab.Width - 2, bar));

            Rectangle caption = Rectangle.FromLTRB(tab.Left, tab.Top + bar, tab.Right, pageTop);
            DrawCaption(g, SelectedIndex, caption, ThemeService.Text);

            // Keyboard focus, which nothing else would show now that the strip is ours.
            // ShowFocusCues is what keeps it off after a mouse click: Windows only wants
            // focus rectangles once somebody has used the keyboard, and it tracks that.
            if (Focused && ShowFocusCues && caption.Width > Scale(8) && caption.Height > Scale(6))
            {
                ControlPaint.DrawFocusRectangle(
                    g,
                    Rectangle.Inflate(caption, -Scale(3), -Scale(2)),
                    ThemeService.Text,
                    body);
            }
        }

        private void DrawCaption(Graphics g, int index, Rectangle bounds, Color colour)
        {
            if (index < 0 || index >= TabCount || bounds.Width < 2 || bounds.Height < 2)
                return;
            string caption = TabPages[index].Text;
            if (string.IsNullOrEmpty(caption))
                return;

            TextRenderer.DrawText(g, caption, Font, bounds, colour,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis
                | TextFormatFlags.NoPrefix);
        }

        /// <summary>
        /// HOW MUCH ACCENT A HOVERED UNSELECTED TAB PICKS UP - AND IT IS A TASTE VALUE, CHOSEN
        /// BY EYE. That is the whole justification, and saying so is the point of this comment.
        ///
        /// <para>
        /// Its sibling <see cref="RestingFill"/> carries measured colour distances because it
        /// had to: that one answers a question with a right answer - "is this tab visibly
        /// behind the page in BOTH themes?" - and the light palette's ten-unit gap is what
        /// forced the number. Hover has no such question. It has to read as a response to the
        /// pointer and it must not read as selection, and everything between those two bounds
        /// is preference. 0.22 was picked inside that range and then looked at.
        /// </para>
        ///
        /// <para>
        /// So do not go looking for the derivation, and do not manufacture one. A measurement
        /// taken after the fact would produce a number that LOOKS derived while still being the
        /// number somebody chose, which is worse than this comment: the next reader would
        /// believe it, and would then be afraid to change a value that is theirs to change.
        /// Changing it changes what the product looks like; that is a decision for whoever
        /// owns the look, not an audit finding.
        /// </para>
        /// </summary>
        private const double HoverAccentBlend = 0.22;

        /// <summary>
        /// What an unselected tab is filled with: the button face pushed towards the border
        /// colour. Measured reason for the push - in the light palette Background is
        /// (250,249,248) and ButtonFace is (240,240,240), ten units apart, which is not
        /// enough on its own to say "this tab is behind the page". In dark they are 23 apart
        /// and the push only helps. Both themes therefore go through the same rule rather
        /// than through two hand-picked colours.
        /// </summary>
        private static Color RestingFill(Color face)
        {
            return Blend(face, ThemeService.Border, 0.35);
        }

        /// <summary>A step from one colour towards another, 0 to 1. Theme-agnostic.</summary>
        private static Color Blend(Color from, Color to, double amount)
        {
            double a = amount < 0 ? 0 : (amount > 1 ? 1 : amount);
            return Color.FromArgb(
                (int)Math.Round(from.R + (to.R - from.R) * a),
                (int)Math.Round(from.G + (to.G - from.G) * a),
                (int)Math.Round(from.B + (to.B - from.B) * a));
        }

        protected override void OnSelectedIndexChanged(EventArgs e)
        {
            base.OnSelectedIndexChanged(e);
            // The selection moved, so which tab is drawn as the front one moved with it.
            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            SetHot(TabAt(e.Location));
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            SetHot(-1);
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            Invalidate();
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            Invalidate();
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            PushFontToNativeStrip();
            Invalidate();
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (_stripFont != IntPtr.Zero)
                {
                    DeleteObject(_stripFont);
                    _stripFont = IntPtr.Zero;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Tab strip font dispose: " + ex.Message);
            }
            base.Dispose(disposing);
        }

        private void SetHot(int index)
        {
            if (_hot == index)
                return;
            _hot = index;
            Invalidate();
        }

        private int TabAt(Point point)
        {
            try
            {
                for (int i = 0; i < TabCount; i++)
                {
                    if (GetTabRect(i).Contains(point))
                        return i;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                // The tab count changed underneath the loop. Nothing is hot.
            }
            return -1;
        }
    }
}
