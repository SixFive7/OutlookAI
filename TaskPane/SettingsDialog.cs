using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// "OutlookAI Settings" — a small theme-aware dialog over OutlookTuningService: master and
    /// per-group toggles, the current effective registry values, a restart-needed indicator,
    /// and a flag line for values group policy has taken back. Opened modeless from the
    /// Explorer ribbon button (single instance) and from the COM automation hook; all calls
    /// arrive on Outlook's UI thread.
    /// </summary>
    public class SettingsDialog : Form
    {
        private static SettingsDialog _open;

        private readonly CheckBox chkMaster;
        private readonly CheckBox chkSearch;
        private readonly CheckBox chkCaching;
        private readonly CheckBox chkOst;
        private readonly CheckBox chkGlobalMcp;
        private readonly GroupBox grpSearch;
        private readonly GroupBox grpCaching;
        private readonly GroupBox grpOst;
        private readonly GroupBox grpClaude;
        private readonly Label lblHeader;
        private readonly Label lblSearchValues;
        private readonly Label lblSearchWarning;
        private readonly Label lblCachingValues;
        private readonly Label lblOstValues;
        private readonly Label lblRestart;
        private readonly Label lblGpo;
        private readonly Label lblGlobalMcpHelp;
        private readonly Label lblMcp;
        private readonly Button btnAddProject;
        private readonly Button btnCopyCommand;
        private readonly Button btnApply;
        private readonly Button btnClose;

        private bool _updating;
        private bool _disposedCustom;

        /// <summary>
        /// Whether the two conditional status lines belong on screen. Held here rather than
        /// read back from <see cref="Control.Visible"/>, which answers "is it on screen right
        /// now" — false for every child while the form itself has not been shown. The layout
        /// runs once inside the constructor, before that, and asking the control there would
        /// reserve no room for a line that is about to appear and open the Claude Code group
        /// underneath it.
        /// </summary>
        private bool _showRestart;
        private bool _showGpo;

        /// <summary>
        /// What a manual registration would name, refreshed by <see cref="RefreshFromState"/>.
        /// Cached in a field so the theme handler can redraw the status line without probing
        /// the disk again.
        /// </summary>
        private string _preferredCommand = "";

        /// <summary>
        /// The server's real path. The copy button uses this rather than the portable
        /// <c>${LOCALAPPDATA}</c> spelling on purpose: PowerShell expands <c>${NAME}</c>
        /// itself — quoted or not — so a copied command carrying that form would arrive at
        /// the CLI with the path blanked out. Claude Code expands it when it READS the
        /// config, which is why the file the button writes can use it and a shell command
        /// cannot.
        /// </summary>
        private string _resolvedServerPath = "";

        internal static bool IsOpen
        {
            get { return _open != null && !_open.IsDisposed; }
        }

        internal static void ShowSettings()
        {
            // Callers must be on the UI thread (ribbon callbacks are; the COM automation
            // surface marshals via the add-in's UI-thread control before calling this).
            SettingsDialog dlg = null;
            try
            {
                if (IsOpen)
                {
                    _open.Activate();
                    return;
                }
                dlg = new SettingsDialog();
                dlg.FormClosed += (s, e) => { if (ReferenceEquals(_open, dlg)) _open = null; };
                dlg.Show();
                dlg.Activate();
                _open = dlg;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ShowSettings: " + ex.Message);
                // Never leave a half-shown zombie registered as "open".
                if (ReferenceEquals(_open, dlg))
                    _open = null;
                try { dlg?.Dispose(); } catch { }
            }
        }

        internal static void CloseIfOpen()
        {
            try
            {
                if (IsOpen)
                    _open.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("CloseIfOpen: " + ex.Message);
            }
        }

        /// <summary>
        /// Repaints an open dialog from stored state. Called when something OUTSIDE it changed
        /// the registration — the startup prompt being answered — so the tick box and the
        /// status line can never sit there contradicting what was just chosen. UI thread; a
        /// no-op when the dialog is closed. Never throws.
        /// </summary>
        internal static void RefreshIfOpen()
        {
            try
            {
                if (IsOpen)
                    _open.RefreshFromState();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("RefreshIfOpen: " + ex.Message);
            }
        }

        public SettingsDialog()
        {
            SuspendLayout();

            Name = "OutlookAISettingsForm";
            Text = "OutlookAI Settings";
            Font = new Font("Segoe UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            // The width is fixed; the height is whatever the laid-out content turns out to
            // need, and PerformDialogLayout sets it before this constructor returns.
            ClientSize = new Size(DialogWidth, 0);

            // Positions and sizes are NOT set here: PerformDialogLayout owns every coordinate
            // in this dialog, so there is exactly one place that decides how tall a label has
            // to be and where the next control starts.

            lblHeader = new Label
            {
                Name = "lblHeader",
                Text = "OutlookAI keeps these Outlook settings applied: fast local search, a fully cached " +
                       "mailbox (sync slider = All), and enough OST size headroom for it.",
            };

            chkMaster = new CheckBox
            {
                Name = "chkMaster",
                Text = "Manage Outlook tuning",
            };

            // --- Search group ---
            grpSearch = new GroupBox
            {
                Name = "grpSearch",
                Text = "Search",
            };
            chkSearch = new CheckBox
            {
                Name = "chkSearch",
                Text = "Keep local search tuning applied",
            };
            lblSearchValues = new Label
            {
                Name = "lblSearchValues",
            };
            lblSearchWarning = new Label
            {
                Name = "lblSearchWarning",
                Text = "Turning this off restores Outlook's online search: slower, capped results, and " +
                       "'show me' results may no longer match what the agent finds.",
            };
            grpSearch.Controls.Add(chkSearch);
            grpSearch.Controls.Add(lblSearchValues);
            grpSearch.Controls.Add(lblSearchWarning);

            // --- Full caching group ---
            grpCaching = new GroupBox
            {
                Name = "grpCaching",
                Text = "Full caching (sync slider = All)",
            };
            chkCaching = new CheckBox
            {
                Name = "chkCaching",
                Text = "Keep full Cached Mode sync applied",
            };
            lblCachingValues = new Label
            {
                Name = "lblCachingValues",
            };
            grpCaching.Controls.Add(chkCaching);
            grpCaching.Controls.Add(lblCachingValues);

            // --- OST headroom group ---
            grpOst = new GroupBox
            {
                Name = "grpOst",
                Text = "OST size headroom",
            };
            chkOst = new CheckBox
            {
                Name = "chkOst",
                Text = "Keep raised OST size limits applied (100 GB max)",
            };
            lblOstValues = new Label
            {
                Name = "lblOstValues",
            };
            grpOst.Controls.Add(chkOst);
            grpOst.Controls.Add(lblOstValues);

            // Both of these are normally hidden, and the layout gives a hidden one no room at
            // all: the old fixed coordinates reserved 62px for them whether they were on
            // screen or not, which is the empty band that used to sit under the OST group.
            lblRestart = new Label
            {
                Name = "lblRestart",
                Text = "Restart Outlook to apply pending changes.",
                Visible = false,
            };

            lblGpo = new Label
            {
                Name = "lblGpo",
                Visible = false,
            };

            // --- Claude Code group: where the mail server is registered, and its state ---
            grpClaude = new GroupBox
            {
                Name = "grpClaude",
                Text = "Mail server in Claude Code",
            };
            chkGlobalMcp = new CheckBox
            {
                Name = "chkGlobalMcp",
                Text = "Make available in all my Claude Code projects",
            };
            lblGlobalMcpHelp = new Label
            {
                Name = "lblGlobalMcpHelp",
                Text = "Registers the mail server in your personal Claude Code configuration, so every " +
                       "project you open can use it. Turning this off removes that entry again.",
            };
            // Always visible: "connected and pointing at the right place" is worth stating,
            // not just its absence.
            lblMcp = new Label
            {
                Name = "lblMcp",
            };
            btnAddProject = new Button
            {
                Name = "btnAddProject",
                Text = "Add to a specific project…",
            };
            btnCopyCommand = new Button
            {
                Name = "btnCopyCommand",
                Text = "Copy CLI command",
            };
            grpClaude.Controls.Add(chkGlobalMcp);
            grpClaude.Controls.Add(lblGlobalMcpHelp);
            grpClaude.Controls.Add(lblMcp);
            grpClaude.Controls.Add(btnAddProject);
            grpClaude.Controls.Add(btnCopyCommand);

            btnApply = new Button
            {
                Name = "btnApply",
                Text = "Apply now",
            };
            btnClose = new Button
            {
                Name = "btnClose",
                Text = "Close",
            };

            Controls.Add(lblHeader);
            Controls.Add(chkMaster);
            Controls.Add(grpSearch);
            Controls.Add(grpCaching);
            Controls.Add(grpOst);
            Controls.Add(lblRestart);
            Controls.Add(lblGpo);
            Controls.Add(grpClaude);
            Controls.Add(btnApply);
            Controls.Add(btnClose);

            AcceptButton = btnClose;
            CancelButton = btnClose;

            chkMaster.CheckedChanged += OnToggleChanged;
            chkSearch.CheckedChanged += OnToggleChanged;
            chkCaching.CheckedChanged += OnToggleChanged;
            chkOst.CheckedChanged += OnToggleChanged;
            // Deliberately NOT OnToggleChanged: this one owns a different service, and
            // toggling it must not drag the Outlook tuning reconcile along with it.
            chkGlobalMcp.CheckedChanged += OnGlobalMcpChanged;
            btnAddProject.Click += OnAddProject;
            btnCopyCommand.Click += OnCopyCommand;
            btnApply.Click += (s, e) =>
            {
                OutlookTuningService.ReconcileFromUi();
                // Same button, same promise: re-check the mail-server registration too, so a
                // drift the user just fixed (installing the runtime, say) is picked up without
                // restarting Outlook. Never throws out of its public surface.
                McpRegistrationService.Reconcile();
                RefreshFromState();
            };
            btnClose.Click += (s, e) => Close();

            ResumeLayout(false);

            ApplyTheme();
            ThemeService.ThemeChanged += OnThemeChanged;

            RefreshFromState();
        }

        // ===== Layout =====
        //
        // The dialog is laid out in code, top to bottom, exactly as it always was — with one
        // rule changed: a wrapped label's height is MEASURED from its text instead of being
        // written down. The shipped dialog wrote it down, and wrote down a value with no
        // slack whatsoever: 30px for the Claude Code help text, which needs exactly two 15px
        // lines at 96 DPI. On a display scaled to 125% the same text needs three 20px lines
        // (60px), so its last line — "...removes that entry again." — had nowhere to render
        // and was cut off. Measuring, and flowing everything below from the measured bottom,
        // is what stops that happening again the next time the wording, the font or the
        // display scale changes.

        private const int DialogWidth = 470;
        private const int FormMargin = 12;
        private const int HeaderTop = 10;
        private const int GroupPadX = 10;        // left/right inset of a group's contents
        private const int GroupFirstRow = 20;    // first row inside a group, clear of its caption
        private const int RowGap = 4;            // between rows
        private const int HelpTextGap = 2;       // a help line hugs the control it explains
        private const int HelpTextIndent = 26;   // aligned with a check box's caption
        private const int GroupSpacing = 6;      // between one group and the next
        private const int GroupBottomPad = 8;    // below the last row inside a group
        private const int GroupButtonPad = 12;   // ...when that last row is buttons
        private const int ButtonGap = 8;         // around a row of buttons
        private const int DialogBottomPad = 10;

        /// <summary>
        /// Lays the dialog out from the top down and sizes the form to the result. Cheap, and
        /// safe to call as often as the content changes — which is the point: it runs after
        /// every refresh and every theme change, so a longer status line, a reworded help
        /// text or a scaled display makes the dialog taller instead of cutting text off.
        /// Never throws; a layout that failed would leave the last good one on screen.
        /// </summary>
        private void PerformDialogLayout()
        {
            SuspendLayout();
            try
            {
                // Absolute child coordinates and a scrolled viewport do not mix, so start from
                // the top. Only ever relevant in the scrolling case below.
                if (AutoScroll)
                    AutoScrollPosition = new Point(0, 0);

                int contentHeight = FlowControls(DialogWidth);

                // Growing to fit the text is only a fix while the buttons stay reachable, so
                // the dialog never grows past the screen it will open on. Beyond that it
                // scrolls — and the flow runs again, narrower by the scroll bar, so a vertical
                // one cannot summon a horizontal one.
                Rectangle workingArea = IsHandleCreated
                    ? Screen.FromControl(this).WorkingArea
                    : Screen.PrimaryScreen.WorkingArea;
                int chrome = Math.Max(0, Height - ClientSize.Height);
                int maxClientHeight = Math.Max(300, workingArea.Height - chrome - 2 * FormMargin);

                if (contentHeight > maxClientHeight)
                {
                    AutoScroll = true;
                    FlowControls(DialogWidth - SystemInformation.VerticalScrollBarWidth);
                    ClientSize = new Size(DialogWidth, maxClientHeight);
                }
                else
                {
                    AutoScroll = false;
                    ClientSize = new Size(DialogWidth, contentHeight);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings layout: " + ex.Message);
            }
            finally
            {
                ResumeLayout(true);
            }
        }

        /// <summary>
        /// Positions every control for a dialog <paramref name="layoutWidth"/> pixels wide and
        /// returns the client height that arrangement needs. Pure geometry: it decides nothing
        /// about what is shown, it only honours it — a hidden status line takes no space.
        /// </summary>
        private int FlowControls(int layoutWidth)
        {
            int inner = layoutWidth - 2 * FormMargin;
            int groupInner = inner - 2 * GroupPadX;
            int y = HeaderTop;

            // The last argument to PlaceLabel is a FLOOR, not a height: the size that label
            // had when the dialog was hand-arranged, kept so the familiar proportions survive
            // and a shorter status line does not make the whole dialog jump. Text that needs
            // more than the floor gets more.

            y = PlaceLabel(lblHeader, FormMargin, y, inner, 32) + RowGap;

            chkMaster.Location = new Point(FormMargin, y);
            chkMaster.Size = new Size(inner, RowHeight(20));
            y = chkMaster.Bottom + ButtonGap;

            // --- Search ---
            int row = PlaceCheck(chkSearch, GroupFirstRow, groupInner) + RowGap;
            row = PlaceLabel(lblSearchValues, GroupPadX, row, groupInner, 74) + RowGap;
            row = PlaceLabel(lblSearchWarning, GroupPadX, row, groupInner, 32);
            y = PlaceGroup(grpSearch, y, inner, row + GroupBottomPad) + GroupSpacing;

            // --- Full caching ---
            row = PlaceCheck(chkCaching, GroupFirstRow, groupInner) + RowGap;
            row = PlaceLabel(lblCachingValues, GroupPadX, row, groupInner, 125);
            y = PlaceGroup(grpCaching, y, inner, row + GroupBottomPad) + GroupSpacing;

            // --- OST headroom ---
            row = PlaceCheck(chkOst, GroupFirstRow, groupInner) + RowGap;
            row = PlaceLabel(lblOstValues, GroupPadX, row, groupInner, 40);
            y = PlaceGroup(grpOst, y, inner, row + GroupBottomPad) + GroupSpacing;

            // The two conditional status lines cost nothing while they are off, which is what
            // closes the gap between the OST group and the one below it. They are positioned
            // either way, so switching one on can never flash it at the top-left corner
            // before the next layout catches up.
            int statusBottom = PlaceLabel(lblRestart, FormMargin, y, inner, 18);
            if (_showRestart)
                y = statusBottom + RowGap;
            statusBottom = PlaceLabel(lblGpo, FormMargin, y, inner, 30);
            if (_showGpo)
                y = statusBottom + RowGap;

            // --- Mail server in Claude Code ---
            row = PlaceCheck(chkGlobalMcp, GroupFirstRow, groupInner) + HelpTextGap;
            // THE defect this layout exists for: this label's text needs two lines at 96 DPI
            // and three on a scaled display, and it used to be given a flat 30px either way.
            row = PlaceLabel(lblGlobalMcpHelp, HelpTextIndent, row,
                             inner - HelpTextIndent - GroupPadX, 30) + RowGap;
            row = PlaceLabel(lblMcp, GroupPadX, row, groupInner, 64) + ButtonGap;

            // Two buttons sharing a row: matched in size, grown for whichever caption needs
            // the most room, and never wider than half the group between them.
            Size add = ButtonSize(btnAddProject, 176, 26);
            Size copy = ButtonSize(btnCopyCommand, 176, 26);
            var projectButton = new Size(
                Math.Min((groupInner - ButtonGap) / 2, Math.Max(add.Width, copy.Width)),
                Math.Max(add.Height, copy.Height));
            btnAddProject.Size = projectButton;
            btnCopyCommand.Size = projectButton;
            btnAddProject.Location = new Point(GroupPadX, row);
            btnCopyCommand.Location = new Point(btnAddProject.Right + ButtonGap, row);
            y = PlaceGroup(grpClaude, y, inner, btnAddProject.Bottom + GroupButtonPad) + ButtonGap;

            btnApply.Size = ButtonSize(btnApply, 84, 26);
            btnClose.Size = ButtonSize(btnClose, 80, 26);
            btnClose.Location = new Point(layoutWidth - FormMargin - btnClose.Width, y);
            btnApply.Location = new Point(btnClose.Left - ButtonGap - btnApply.Width, y);

            return btnClose.Bottom + DialogBottomPad;
        }

        /// <summary>
        /// Puts <paramref name="label"/> at (<paramref name="x"/>, <paramref name="y"/>),
        /// <paramref name="width"/> wide and as tall as its text needs there — never shorter
        /// than <paramref name="minHeight"/>, the height the dialog shipped with, so labels
        /// whose text already fits keep their familiar proportions and the dialog does not
        /// resize itself every time a status line happens to get shorter. Returns the Y just
        /// below it, which is what the next control flows from.
        /// </summary>
        private static int PlaceLabel(Label label, int x, int y, int width, int minHeight)
        {
            label.Location = new Point(x, y);
            label.Size = new Size(width, Math.Max(minHeight, MeasureLabelHeight(label, width)));
            return label.Bottom;
        }

        /// <summary>
        /// How tall <paramref name="label"/>'s current text is once wrapped at
        /// <paramref name="width"/>, measured with the font it will actually paint with rather
        /// than the font someone had in mind while writing the coordinates down.
        /// </summary>
        private static int MeasureLabelHeight(Label label, int width)
        {
            if (width <= 0)
                return 0;
            try
            {
                // An unbounded height asks "how tall, wrapped at this width?". The two extra
                // pixels are slack: the shipped help label needed exactly the 30px it was
                // given, and a label with no slack is one font update away from clipping.
                Size needed = TextRenderer.MeasureText(
                    label.Text ?? "",
                    label.Font,
                    new Size(width, int.MaxValue),
                    TextFormatFlags.WordBreak);
                return needed.Height + 2;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings measure: " + ex.Message);
                return label.Height;
            }
        }

        /// <summary>Places a group's check box on its own row and returns the Y below it.</summary>
        private int PlaceCheck(CheckBox check, int y, int width)
        {
            check.Location = new Point(GroupPadX, y);
            check.Size = new Size(width, RowHeight(18));
            return check.Bottom;
        }

        private static int PlaceGroup(GroupBox group, int y, int width, int height)
        {
            group.Location = new Point(FormMargin, y);
            group.Size = new Size(width, height);
            return group.Bottom;
        }

        /// <summary>
        /// A single-line row's height: what it was designed as, or the font's line height
        /// where that is taller — the case on a scaled display, where the font grows and
        /// these coordinates do not.
        /// </summary>
        private int RowHeight(int designHeight)
        {
            return Math.Max(designHeight, Font.Height + 2);
        }

        /// <summary>
        /// A button's designed size, grown wherever its caption no longer fits inside it.
        /// </summary>
        private static Size ButtonSize(Button button, int designWidth, int designHeight)
        {
            Size preferred;
            try
            {
                preferred = button.GetPreferredSize(Size.Empty);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings button size: " + ex.Message);
                preferred = Size.Empty;
            }
            return new Size(Math.Max(designWidth, preferred.Width), Math.Max(designHeight, preferred.Height));
        }

        private void OnToggleChanged(object sender, EventArgs e)
        {
            if (_updating)
                return;
            try
            {
                OutlookTuningService.SetMasterEnabled(chkMaster.Checked);
                OutlookTuningService.SetGroupEnabled(OutlookTuningService.GroupSearch, chkSearch.Checked);
                OutlookTuningService.SetGroupEnabled(OutlookTuningService.GroupCaching, chkCaching.Checked);
                OutlookTuningService.SetGroupEnabled(OutlookTuningService.GroupOst, chkOst.Checked);
                OutlookTuningService.ReconcileFromUi();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings toggle: " + ex.Message);
            }
            RefreshFromState();
        }

        // The "all my projects" toggle. Ticking or unticking it IS the user declaring their
        // intent, so it applies immediately rather than waiting for the next Outlook start —
        // a user who unticks it expects the entry gone now — and it never re-opens the
        // question the startup prompt asks: they just answered it.
        private void OnGlobalMcpChanged(object sender, EventArgs e)
        {
            if (_updating)
                return;
            try
            {
                Cursor = Cursors.WaitCursor;
                McpRegistrationService.ApplyUserChoice(chkGlobalMcp.Checked);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP toggle: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            RefreshFromState();
        }

        // Project scope: writes/merges .mcp.json in a folder the user picks. Never throws
        // out of the click handler - a failure is something to read, not a crash dialog.
        private void OnAddProject(object sender, EventArgs e)
        {
            try
            {
                string folder;
                using (var picker = new FolderBrowserDialog())
                {
                    picker.Description = "Choose the project folder that should get the OutlookAI mail server.";
                    picker.ShowNewFolderButton = false;
                    if (picker.ShowDialog(this) != DialogResult.OK)
                        return;
                    folder = picker.SelectedPath;
                }

                string configPath, error;
                bool ok;
                try
                {
                    Cursor = Cursors.WaitCursor;
                    ok = McpRegistrationService.TryRegisterInProject(folder, out configPath, out error);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

                if (!ok)
                {
                    MessageBox.Show(this, error, "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine("Written to:");
                sb.AppendLine(configPath);
                sb.AppendLine();
                sb.AppendLine("The first time you open Claude Code in that folder it will ask you to approve "
                              + "this server. That prompt is Claude Code's own security check and only you can "
                              + "answer it — the add-in deliberately does not.");
                sb.AppendLine();
                sb.Append(".mcp.json is normally committed to source control. ");
                if (McpConfigEditor.ContainsEnvironmentReference(_preferredCommand))
                {
                    sb.Append("Because the entry points at ${LOCALAPPDATA} rather than a fixed path, it is "
                              + "portable: teammates who have OutlookAI installed get a working mail server, "
                              + "and teammates who do not simply see a failed-connection warning for this one "
                              + "server — nothing else breaks.");
                }
                else
                {
                    sb.Append("This entry names a fixed path on this machine (the mail server is not in the "
                              + "default install location here), so it will not resolve on a teammate's "
                              + "machine — they would see a failed-connection warning for this one server.");
                }

                MessageBox.Show(this, sb.ToString(), "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Add to project: " + ex.Message);
                MessageBox.Show(this, ex.Message, "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // The same registration as a command, for people who would rather use the CLI.
        private void OnCopyCommand(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_resolvedServerPath))
                    return;

                string command = "claude mcp add --scope project outlookai -- \"" + _resolvedServerPath + "\"";
                Clipboard.SetText(command);
                MessageBox.Show(
                    this,
                    "Copied to the clipboard. Run it from the project folder:" + Environment.NewLine
                        + Environment.NewLine + command + Environment.NewLine + Environment.NewLine
                        + "Use --scope user instead of --scope project to cover every project, which is what "
                        + "the tick box above does." + Environment.NewLine + Environment.NewLine
                        + "This names the real path so it works in any shell. The \"Add to a specific project…\" "
                        + "button writes the portable ${LOCALAPPDATA} form instead, which is the one to keep if "
                        + "the file goes into source control.",
                    "OutlookAI",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // The clipboard belongs to whatever grabbed it last; losing that race is not
                // worth a crash dialog.
                System.Diagnostics.Debug.WriteLine("Copy command: " + ex.Message);
            }
        }

        // Reads only what the last reconcile recorded, so opening the dialog never touches
        // Claude Code's config file. "Apply now" is what re-runs the reconcile.
        private void RefreshMcpLine()
        {
            try
            {
                var reg = McpRegistrationService.GetSnapshot();
                string text;
                switch (reg.Status)
                {
                    case McpRegistrationService.StatusOk:
                        text = "Registered for all your projects.";
                        break;
                    case McpRegistrationService.StatusHealed:
                        text = "Registration was missing or pointing elsewhere and has been repaired.";
                        break;
                    case McpRegistrationService.StatusDisabled:
                        text = "Not registered for all your projects. Individual projects with a .mcp.json are unaffected.";
                        break;
                    case McpRegistrationService.StatusRemoved:
                        text = "Removed from your personal Claude Code configuration. Individual projects with a .mcp.json are unaffected.";
                        break;
                    case McpRegistrationService.StatusNoClaude:
                        text = "Claude Code was not found on this machine, so there is nothing to register with.";
                        break;
                    case McpRegistrationService.StatusNoServer:
                        text = "The mail server is not installed alongside the add-in; any existing registration was left unchanged.";
                        break;
                    case McpRegistrationService.StatusNoRuntime:
                        text = "The mail server needs the .NET 10 runtime, which is not installed. Get it from "
                               + McpRegistrationService.DotnetRuntimeDownloadUrl + " and restart Outlook.";
                        break;
                    case McpRegistrationService.StatusParseFailed:
                        text = "Claude Code's configuration could not be read, so it was left untouched.";
                        break;
                    case McpRegistrationService.StatusAwaitingChoice:
                        // The detail carries what was actually found, and there is no useful
                        // shorter way to say it: nothing has been changed, and why.
                        text = string.IsNullOrEmpty(reg.Detail)
                            ? "Waiting for you to choose whether to register the mail server. Nothing has been changed."
                            : reg.Detail;
                        break;
                    default:
                        text = "Registration state unknown"
                               + (string.IsNullOrEmpty(reg.Detail) ? "." : (" — " + reg.Detail));
                        break;
                }

                if (!string.IsNullOrEmpty(reg.RegisteredCommand) &&
                    (reg.Status == McpRegistrationService.StatusOk || reg.Status == McpRegistrationService.StatusHealed))
                {
                    text += Environment.NewLine + reg.RegisteredCommand;
                }
                else if (!string.IsNullOrEmpty(_preferredCommand))
                {
                    // Labelled: an unadorned path under "not registered" reads like a claim
                    // that it IS registered.
                    text += Environment.NewLine + "Server: " + _preferredCommand;
                }

                lblMcp.Text = text;
                lblMcp.ForeColor = IsMcpProblem(reg.Status) ? ThemeService.StatusError : ThemeService.SecondaryText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP line refresh: " + ex.Message);
            }
        }

        private static bool IsMcpProblem(string status)
        {
            return status == McpRegistrationService.StatusNoRuntime
                || status == McpRegistrationService.StatusParseFailed
                || status == McpRegistrationService.StatusError;
        }

        private void RefreshFromState()
        {
            // The Claude Code group first and in its own guarded block: a tuning snapshot
            // that fails must not leave the registration controls unpainted.
            _updating = true;
            try
            {
                _preferredCommand = McpRegistrationService.ResolvePreferredCommand();
                _resolvedServerPath = McpRegistrationService.ResolveInstalledServerPath() ?? "";
                chkGlobalMcp.Checked = McpRegistrationService.GetSnapshot().GlobalRegistrationEnabled;
                bool haveServer = !string.IsNullOrEmpty(_preferredCommand);
                btnAddProject.Enabled = haveServer;
                btnCopyCommand.Enabled = haveServer && !string.IsNullOrEmpty(_resolvedServerPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP settings refresh: " + ex.Message);
            }
            finally
            {
                _updating = false;
            }

            RefreshMcpLine();

            OutlookTuningService.TuningSnapshot snap;
            try
            {
                snap = OutlookTuningService.GetSnapshot();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Settings refresh: " + ex.Message);
                // The Claude Code line above was still repainted, so lay out around it.
                PerformDialogLayout();
                return;
            }

            _updating = true;
            try
            {
                chkMaster.Checked = snap.MasterEnabled;
                chkSearch.Checked = snap.SearchEnabled;
                chkCaching.Checked = snap.CachingEnabled;
                chkOst.Checked = snap.OstEnabled;

                chkSearch.Enabled = snap.MasterEnabled;
                chkCaching.Enabled = snap.MasterEnabled;
                chkOst.Enabled = snap.MasterEnabled;

                lblSearchValues.Text = BuildGroupText(snap, OutlookTuningService.GroupSearch);
                lblCachingValues.Text = BuildGroupText(snap, OutlookTuningService.GroupCaching);
                lblOstValues.Text = BuildGroupText(snap, OutlookTuningService.GroupOst);

                _showRestart = snap.RestartNeeded;
                lblRestart.Visible = _showRestart;

                if (snap.PolicyConflicts.Count > 0)
                {
                    var names = new StringBuilder();
                    foreach (var v in snap.Values)
                    {
                        if (!v.BackedOff)
                            continue;
                        if (names.Length > 0)
                            names.Append(", ");
                        names.Append(v.Entry.ValueName);
                    }
                    lblGpo.Text = "Managed by your organization's policy (left unchanged): " + names;
                    _showGpo = true;
                }
                else
                {
                    _showGpo = false;
                }

                lblGpo.Visible = _showGpo;
            }
            finally
            {
                _updating = false;
            }

            // Last, and always: every label above may have just changed length, and the two
            // status lines may have just appeared or gone.
            PerformDialogLayout();
        }

        private static string BuildGroupText(OutlookTuningService.TuningSnapshot snap, string groupId)
        {
            var sb = new StringBuilder();
            foreach (var v in snap.Values)
            {
                if (v.Entry.GroupId != groupId)
                    continue;
                string prefix = "";
                if (groupId == OutlookTuningService.GroupCaching)
                    prefix = v.Entry.IsPolicyHive ? "policy: " : "user: ";
                string live = v.Live.HasValue ? v.Live.Value.ToString() : "(not set)";
                string status;
                if (v.BackedOff)
                    status = "(policy override, left unchanged)";
                else if (v.InSync)
                    status = "(OK)";
                else if (!v.GroupEnabled)
                    status = "(not managed)";
                else
                    status = "-> " + v.Desired + " on apply";
                sb.AppendLine(prefix + v.Entry.ValueName + " = " + live + "  " + status);
            }
            return sb.ToString().TrimEnd();
        }

        // ===== Theming (mirrors the AITaskPane pattern) =====

        private void OnThemeChanged(object sender, EventArgs e)
        {
            // ThemeService may raise this on a non-UI (SystemEvents / registry watcher) thread.
            if (_disposedCustom || IsDisposed || !IsHandleCreated)
                return;
            try { BeginInvoke((Action)ApplyTheme); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void ApplyTheme()
        {
            BackColor = ThemeService.Background;
            ForeColor = ThemeService.Text;

            lblHeader.ForeColor = ThemeService.SecondaryText;
            lblSearchWarning.ForeColor = ThemeService.SecondaryText;
            lblRestart.ForeColor = ThemeService.StatusError;
            lblGpo.ForeColor = ThemeService.SecondaryText;
            lblGlobalMcpHelp.ForeColor = ThemeService.SecondaryText;
            // Colour depends on the state, so let the refresh own it rather than pinning a
            // colour here that the next refresh would immediately overwrite.
            RefreshMcpLine();

            foreach (var grp in new[] { grpSearch, grpCaching, grpOst, grpClaude })
                grp.ForeColor = ThemeService.Text;

            foreach (var chk in new[] { chkMaster, chkSearch, chkCaching, chkOst, chkGlobalMcp })
                chk.ForeColor = ThemeService.Text;

            foreach (var lbl in new[] { lblSearchValues, lblCachingValues, lblOstValues })
                lbl.ForeColor = ThemeService.Text;

            foreach (var btn in new[] { btnApply, btnClose, btnAddProject, btnCopyCommand })
            {
                if (ThemeService.IsDarkMode)
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderColor = ThemeService.Border;
                    btn.BackColor = ThemeService.ButtonFace;
                    btn.ForeColor = ThemeService.ButtonText;
                }
                else
                {
                    btn.FlatStyle = FlatStyle.Standard;
                    btn.UseVisualStyleBackColor = true;
                    btn.ForeColor = ThemeService.ButtonText;
                }
            }

            // RefreshMcpLine above rewrote the status line, and a theme switch is also the
            // moment a font substitution would land: re-measure rather than assume.
            PerformDialogLayout();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposedCustom)
            {
                _disposedCustom = true;
                ThemeService.ThemeChanged -= OnThemeChanged;
            }
            base.Dispose(disposing);
        }
    }
}
