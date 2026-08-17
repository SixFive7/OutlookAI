using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// "OutlookAI Settings" - one resizable, tabbed window holding everything the add-in lets a
    /// user change. Modeless, single instance, opened from the Explorer ribbon button and from
    /// the COM automation hook; every call arrives on Outlook's UI thread.
    ///
    /// FIVE TABS, and what is on them:
    ///   Outlook     - the master switch and the three Outlook tuning groups, plus the restart
    ///                 and group-policy status lines.
    ///   Claude Code - which model requests run on, and where the mail server is registered.
    ///   Prompts     - the five prompt sections every request is assembled from.
    ///   Buttons     - the quick buttons the compose sidebar shows, in order.
    ///   Updates     - the version line, the last check's error, and "Check for updates".
    ///
    /// FOUR DECISIONS SHAPE THE FILE:
    ///
    ///  1. TWO COMMIT MODELS, ON PURPOSE, AND THE FOOTER SAYS SO. A tick box IS the decision, so
    ///     the tuning boxes and the Claude Code toggle write the moment they are clicked, exactly
    ///     as they always have. So does picking a model from the list - a chosen list entry is
    ///     the same shape of decision as a ticked box. Text being typed is not a decision: a
    ///     half-typed prompt is not an instruction anybody meant to send, so the Prompts and
    ///     Buttons tabs buffer into drafts and reach the registry only on "Apply now". The
    ///     custom model id sits between the two and is treated as text: it writes when the box
    ///     is left, when Enter is pressed, on Apply now, and on close, never per keystroke.
    ///     That is why the footer is Apply now + Close and NOT Apply/Cancel: with instant-apply
    ///     tick boxes one tab away, a global Cancel would be a lie about what it undoes. Closing
    ///     with unsaved prompt edits asks once, through <see cref="ConfirmDiscard"/>.
    ///
    ///  2. ENTER TYPES A NEWLINE. <see cref="Form.AcceptButton"/> is deliberately null - with a
    ///     default button, Enter in any of the six multi-line editors would press it instead of
    ///     starting a new line. Escape still closes, through CancelButton, and goes through the
    ///     same unsaved-changes question as the window's X.
    ///
    ///  3. IT DOCKS INSTEAD OF COUNTING PIXELS. The predecessor of this window was a fixed-size
    ///     dialog that owned every child coordinate, and re-laying it out reset its scroll
    ///     position - which is why its version line had to prove the layout had really changed
    ///     before it was allowed to run. A resizable window cannot count pixels anyway, so the
    ///     layout is dock/anchor plus TableLayoutPanel throughout. Wrapped labels still have to
    ///     be MEASURED - see <see cref="ReflowWrappedLabels"/> - and every pixel constant is a
    ///     96-DPI design value put through <see cref="Scaled"/>, so a display at 125% or 150%
    ///     moves the whole layout instead of clipping the last line of a label.
    ///
    ///  4. THEMING WALKS THE TREE. <see cref="ApplyThemeTo"/> recurses over the real control
    ///     tree and colours by type, so a control cannot fall off a hand-maintained list and
    ///     paint light-on-light in dark mode. The three colour ROLES a type cannot imply (body,
    ///     secondary, warning) are registered by <see cref="NewLabel"/>, the factory that makes
    ///     the label, so forgetting to register is not a thing you can do.
    ///
    /// The prompts and buttons half of the window lives in SettingsDialog.Prompts.cs.
    /// </summary>
    public partial class SettingsDialog : Form
    {
        private static SettingsDialog _open;

        // ===== Shell =====

        private readonly TableLayoutPanel _root;
        private readonly ThemedTabControl _pageTabs;
        private readonly TabPage _tabOutlook;
        private readonly TabPage _tabClaude;
        private readonly TabPage _tabPrompts;
        private readonly TabPage _tabButtons;
        private readonly TabPage _tabUpdates;

        /// <summary>
        /// The themed panel inside each tab page - the surface the user actually sees behind the
        /// controls. Held so <see cref="ApplyTheme"/> colours it explicitly rather than trusting
        /// a TabPage to pass its own BackColor down.
        /// </summary>
        private readonly List<Control> _pageSurfaces = new List<Control>();

        private readonly Button btnApply;
        private readonly Button btnClose;

        // ===== Outlook tab =====

        private readonly Label lblHeader;
        private readonly CheckBox chkMaster;
        private readonly GroupBox grpSearch;
        private readonly CheckBox chkSearch;
        private readonly Label lblSearchValues;
        private readonly Label lblSearchWarning;
        private readonly GroupBox grpCaching;
        private readonly CheckBox chkCaching;
        private readonly Label lblCachingValues;
        private readonly GroupBox grpOst;
        private readonly CheckBox chkOst;
        private readonly Label lblOstValues;
        private readonly Label lblRestart;
        private readonly Label lblGpo;

        // ===== Claude Code tab =====

        private readonly GroupBox grpModel;
        private readonly Label lblModelHelp;
        private readonly ComboBox cmbModel;
        private readonly Label lblModelIdCaption;
        private readonly TextBox txtModel;
        private readonly Label lblModelNote;
        private readonly Label lblModelState;

        private readonly GroupBox grpClaude;
        private readonly CheckBox chkGlobalMcp;
        private readonly Label lblGlobalMcpHelp;
        private readonly Label lblMcp;
        private readonly Button btnAddProject;
        private readonly Button btnCopyCommand;

        // ===== Updates tab =====

        private readonly GroupBox grpVersion;
        private readonly Label lblVersion;
        private readonly Label lblUpdateError;
        private readonly Button btnCheckUpdates;

        /// <summary>
        /// Keeps "checked 4m ago" honest and picks up a check finishing. The rate is
        /// <see cref="UpdateService.VersionLineTickMs"/>, shared with the sidebar's indicator so
        /// the two cannot be tuned apart; <see cref="RefreshVersionLine"/> makes an unchanged
        /// tick free, so this costs nothing while nothing is happening.
        /// </summary>
        private readonly Timer _versionTimer;

        private bool _updating;
        private bool _disposedCustom;
        private bool _reflowing;
        private bool _closingWithoutAsking;

        /// <summary>
        /// Set between the user clicking "Check for updates" and that check completing. Held
        /// separately from <see cref="UpdateService.IsChecking"/> so the button stays disabled
        /// across the whole round trip, including a check that ended the instant it began -
        /// a developer build, or one already running when the click landed.
        /// </summary>
        private bool _checkInFlight;

        /// <summary>
        /// Whether the update-error line belongs on screen. Kept as a field rather than read
        /// back from <see cref="Control.Visible"/> because <see cref="RefreshVersionLine"/> has
        /// to answer "did this tick change the geometry", and a line appearing or disappearing
        /// is the biggest geometry change there is.
        /// </summary>
        private bool _showUpdateError;

        /// <summary>
        /// What a manual registration would name, refreshed by <see cref="RefreshFromState"/>.
        /// Cached in a field so the theme handler can redraw the status line without probing
        /// the disk again.
        /// </summary>
        private string _preferredCommand = "";

        /// <summary>
        /// The server's real path. The copy button uses this rather than the portable
        /// <c>${LOCALAPPDATA}</c> spelling on purpose: PowerShell expands <c>${NAME}</c>
        /// itself - quoted or not - so a copied command carrying that form would arrive at
        /// the CLI with the path blanked out. Claude Code expands it when it READS the
        /// config, which is why the file the button writes can use it and a shell command
        /// cannot.
        /// </summary>
        private string _resolvedServerPath = "";

        // ===== Single instance, modeless =====

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
                    if (_open.WindowState == FormWindowState.Minimized)
                        _open.WindowState = FormWindowState.Normal;
                    _open.Activate();
                    return;
                }
                dlg = new SettingsDialog();
                var opened = dlg;
                dlg.FormClosed += (s, e) => { if (ReferenceEquals(_open, opened)) _open = null; };
                dlg.Show();
                dlg.Activate();
                _open = dlg;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ShowSettings: " + ex.Message);
                // Never leave a half-shown zombie registered as "open".
                if (ReferenceEquals(_open, dlg))
                    _open = null;
                try { if (dlg != null) dlg.Dispose(); } catch { }
            }
        }

        /// <summary>
        /// Closes the window if it is open, WITHOUT asking about unsaved prompt edits. Reached
        /// from Outlook shutting down and from the COM automation surface tidying up - neither
        /// is a moment at which a modal question could be answered. The user's own Close, the
        /// window's X and Escape do ask; they do not come through here.
        /// </summary>
        internal static void CloseIfOpen()
        {
            try
            {
                if (IsOpen)
                    _open.CloseWithoutAsking();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("CloseIfOpen: " + ex.Message);
            }
        }

        /// <summary>
        /// Repaints an open window from stored state. Called when something OUTSIDE it changed
        /// the registration - the startup prompt being answered - so the tick box and the
        /// status line can never sit there contradicting what was just chosen. UI thread; a
        /// no-op when the window is closed. Never throws.
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
                Debug.WriteLine("RefreshIfOpen: " + ex.Message);
            }
        }

        public SettingsDialog()
        {
            SuspendLayout();

            Name = "OutlookAISettingsForm";
            Text = "OutlookAI Settings";
            Font = new Font("Segoe UI", 9F);
            // Every size in this file is a 96-DPI design value put through Scaled(), which reads
            // the font the form actually has. Letting WinForms scale on top of that would apply
            // the display scaling twice.
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            // In the taskbar, unlike the fixed dialog this replaces. It is now an editing surface
            // people leave open next to a compose window, and an ownerless modeless form that is
            // not in the taskbar has no way back once Outlook is clicked - other than the ribbon
            // button, which does bring it forward.
            ShowInTaskbar = true;
            StartPosition = FormStartPosition.CenterScreen;

            // NO AcceptButton. See rule 2 in the class comment: a default button would swallow
            // Enter out of every multi-line editor in this window.
            AcceptButton = null;

            // --- Outlook tab ---
            lblHeader = NewLabel(
                "OutlookAI keeps these Outlook settings applied: fast local search, a fully cached "
                + "mailbox (sync slider = All), and enough OST size headroom for it.",
                LabelRole.Secondary, wrap: true);

            chkMaster = NewCheck("chkMaster", "Manage Outlook tuning");

            chkSearch = NewCheck("chkSearch", "Keep local search tuning applied");
            lblSearchValues = NewLabel("", LabelRole.Body, wrap: true);
            lblSearchValues.Name = "lblSearchValues";
            lblSearchWarning = NewLabel(
                "Turning this off restores Outlook's online search: slower, capped results, and "
                + "'show me' results may no longer match what the agent finds.",
                LabelRole.Secondary, wrap: true);
            grpSearch = NewGroup("grpSearch", "Search",
                                 chkSearch, lblSearchValues, lblSearchWarning);

            chkCaching = NewCheck("chkCaching", "Keep full Cached Mode sync applied");
            lblCachingValues = NewLabel("", LabelRole.Body, wrap: true);
            lblCachingValues.Name = "lblCachingValues";
            grpCaching = NewGroup("grpCaching", "Full caching (sync slider = All)",
                                  chkCaching, lblCachingValues);

            // The cap is read from the tuning service rather than typed in here: the desired
            // numbers live in the registry and are meant to be tunable, so a hand-written
            // "100 GB" would go on claiming a limit the product had stopped applying.
            chkOst = NewCheck("chkOst",
                "Keep raised OST size limits applied (" + OutlookTuningService.DescribeOstMaxSize() + " max)");
            lblOstValues = NewLabel("", LabelRole.Body, wrap: true);
            lblOstValues.Name = "lblOstValues";
            grpOst = NewGroup("grpOst", "OST size headroom", chkOst, lblOstValues);

            // Both of these are normally hidden. In a docked layout a hidden control genuinely
            // costs nothing: the layout engine skips it, so there is no empty band to close.
            lblRestart = NewLabel("Restart Outlook to apply pending changes.",
                                  LabelRole.Warning, wrap: true);
            lblRestart.Name = "lblRestart";
            lblRestart.Visible = false;

            lblGpo = NewLabel("", LabelRole.Secondary, wrap: true);
            lblGpo.Name = "lblGpo";
            lblGpo.Visible = false;

            // --- Claude Code tab: the model ---
            // First on the tab because it is the setting a user is most likely to have opened
            // this window to change; the mail-server group below is one-time setup that mostly
            // reports its own state.
            lblModelHelp = NewLabel(
                "Which model OutlookAI asks for. Left on the default it asks for nothing in "
                + "particular, so Claude Code uses whatever it is set up to use - your own "
                + "choice, and any model released later, without an OutlookAI update.",
                LabelRole.Secondary, wrap: true);
            lblModelHelp.Name = "lblModelHelp";

            cmbModel = new ComboBox
            {
                Name = "cmbModel",
                // DropDownList, not DropDown: the free-text case has its own box below, and a
                // combo that is both a list and an editor makes "what did I actually choose"
                // impossible to answer from a glance.
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                // Names, not model ids, so the list cannot go stale - see ModelCatalog.
                FormattingEnabled = false,
                // Owner-drawn because this is the ONE control in this window that ignores the
                // BackColor the theme walk gives it: the area showing the selected entry is
                // painted by the OS, which in dark mode drew a white field with white text on
                // it. Measured, not guessed - see ReadModelPickerPixels in the render harness.
                DrawMode = DrawMode.OwnerDrawFixed,
            };
            foreach (ModelChoice choice in ModelCatalog.BuildChoices())
                cmbModel.Items.Add(choice);

            lblModelIdCaption = NewLabel("Model id", LabelRole.Body, wrap: false);
            lblModelIdCaption.Name = "lblModelIdCaption";

            txtModel = new TextBox
            {
                Name = "txtModel",
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                // Only live for the Custom entry, so an id cannot be typed into a box that is
                // not the thing being sent.
                Enabled = false,
            };

            // Hidden unless there is something true to say - see RefreshModelLines. It never
            // claims a model cannot do something; it says OutlookAI cannot vouch for it.
            lblModelNote = NewLabel("", LabelRole.Warning, wrap: true);
            lblModelNote.Name = "lblModelNote";
            lblModelNote.Visible = false;

            // The honesty line: the argument that will actually be sent, in full.
            lblModelState = NewLabel("", LabelRole.Secondary, wrap: true);
            lblModelState.Name = "lblModelState";

            grpModel = NewGroup("grpModel", "Model",
                                lblModelHelp, cmbModel, lblModelIdCaption, txtModel,
                                lblModelNote, lblModelState);

            // --- Claude Code tab: the mail server ---
            chkGlobalMcp = NewCheck("chkGlobalMcp", "Make available in all my Claude Code projects");
            lblGlobalMcpHelp = NewLabel(
                "Registers the mail server in your personal Claude Code configuration, so every "
                + "project you open can use it. Turning this off removes that entry again.",
                LabelRole.Secondary, wrap: true);
            lblGlobalMcpHelp.Name = "lblGlobalMcpHelp";
            // Always visible: "connected and pointing at the right place" is worth stating,
            // not just its absence. Dynamic, because its colour says whether it is a problem
            // and RefreshMcpLine owns that.
            lblMcp = NewLabel("", LabelRole.Dynamic, wrap: true);
            lblMcp.Name = "lblMcp";
            btnAddProject = NewButton("Add to a specific project…", "btnAddProject");
            btnCopyCommand = NewButton("Copy CLI command", "btnCopyCommand");
            grpClaude = NewGroup("grpClaude", "Mail server in Claude Code",
                                 chkGlobalMcp, lblGlobalMcpHelp, lblMcp,
                                 NewButtonRow("claudeButtons", btnAddProject, btnCopyCommand));

            // --- Updates tab ---
            // Text comes from UpdateService, so this line and the sidebar's always agree.
            lblVersion = NewLabel("", LabelRole.Secondary, wrap: true);
            lblVersion.Name = "lblVersion";
            // The sidebar has to hide the reason behind a link for want of space; here it fits,
            // so it is simply on screen when there is one.
            lblUpdateError = NewLabel("", LabelRole.Warning, wrap: true);
            lblUpdateError.Name = "lblUpdateError";
            lblUpdateError.Visible = false;
            btnCheckUpdates = NewButton("Check for updates", "btnCheckUpdates");
            grpVersion = NewGroup("grpVersion", "Version and updates",
                                  lblVersion, lblUpdateError,
                                  NewButtonRow("updateButtons", btnCheckUpdates));

            // --- Prompts and Buttons tabs (SettingsDialog.Prompts.cs) ---
            BuildPromptControls();

            // --- Tabs ---
            _tabOutlook = NewPage("Outlook", "tabOutlook",
                NewScroller("outlookPage", NewStack("outlookStack",
                    lblHeader, chkMaster, grpSearch, grpCaching, grpOst, lblRestart, lblGpo)));
            _tabClaude = NewPage("Claude Code", "tabClaude",
                NewScroller("claudePage", NewStack("claudeStack", grpModel, grpClaude)));
            _tabPrompts = NewPage("Prompts", "tabPrompts", BuildSectionsPage());
            _tabButtons = NewPage("Buttons", "tabButtons", BuildButtonsPage());
            _tabUpdates = NewPage("Updates", "tabUpdates",
                NewScroller("updatesPage", NewStack("updatesStack", grpVersion)));

            _pageTabs = new ThemedTabControl
            {
                Name = "pageTabs",
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
            };
            _pageTabs.TabPages.Add(_tabOutlook);
            _pageTabs.TabPages.Add(_tabClaude);
            _pageTabs.TabPages.Add(_tabPrompts);
            _pageTabs.TabPages.Add(_tabButtons);
            _pageTabs.TabPages.Add(_tabUpdates);

            // --- Footer ---
            btnApply = NewButton("Apply now", "btnApply");
            btnClose = NewButton("Close", "btnClose");

            _root = new TableLayoutPanel
            {
                Name = "root",
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = Padding.Empty,
            };
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // Rows are added by index rather than by Controls.Add order, so nothing here depends
            // on the docking order rule that trips up Dock=Top/Bottom stacks. The footer is a row
            // of its own, OUTSIDE the tab control: Apply now and Close speak for the whole window,
            // so they must not move or repaint when the page changes.
            _root.Controls.Add(_pageTabs, 0, 0);
            _root.Controls.Add(BuildFooter(), 0, 1);

            Controls.Add(_root);

            // Escape closes, through the same unsaved-changes question as the window's X.
            CancelButton = btnClose;

            chkMaster.CheckedChanged += OnToggleChanged;
            chkSearch.CheckedChanged += OnToggleChanged;
            chkCaching.CheckedChanged += OnToggleChanged;
            chkOst.CheckedChanged += OnToggleChanged;
            // Deliberately NOT OnToggleChanged: this one owns a different service, and
            // toggling it must not drag the Outlook tuning reconcile along with it.
            chkGlobalMcp.CheckedChanged += OnGlobalMcpChanged;
            cmbModel.SelectedIndexChanged += OnModelChoiceChanged;
            cmbModel.DrawItem += OnDrawModelChoice;
            // Three ways to finish typing an id, and none of them is a keystroke: leaving the
            // box, pressing Enter in it, and Apply now (plus closing the window, in
            // OnFormClosing). TextChanged only re-reads the note, which writes nothing.
            txtModel.Leave += (s, e) => CommitCustomModel();
            txtModel.KeyDown += OnModelIdKeyDown;
            txtModel.TextChanged += OnModelIdEdited;
            btnAddProject.Click += OnAddProject;
            btnCopyCommand.Click += OnCopyCommand;
            btnCheckUpdates.Click += OnCheckForUpdates;
            btnApply.Click += OnApplyNow;
            btnClose.Click += (s, e) => Close();
            _pageTabs.SelectedIndexChanged += OnPageChanged;
            WirePromptEvents();

            ResumeLayout(false);

            ApplyMetrics();
            LoadFromStore(0);

            ApplyTheme();
            ThemeService.ThemeChanged += OnThemeChanged;

            // Only ever re-lays the window out when the line actually changed, which on most
            // ticks it has not: "checked 4m ago" turns over once a minute.
            _versionTimer = new Timer { Interval = UpdateService.VersionLineTickMs };
            _versionTimer.Tick += (s, e) =>
            {
                if (RefreshVersionLine())
                    RelayoutAfterTextChange();
            };
            _versionTimer.Start();

            RefreshFromState();
        }

        // ===== Construction helpers =====

        private enum LabelRole
        {
            Body,
            Secondary,
            Warning,

            /// <summary>
            /// The owner paints this one. Used where the colour carries meaning that a theme
            /// switch must not flatten - the mail-server status line, which turns red on a
            /// problem, and the footer status, which turns green on a save.
            /// </summary>
            Dynamic,
        }

        // Colour roles a control's type cannot imply. Populated by NewLabel, so a label that
        // exists is a label that gets themed.
        private readonly List<Label> _bodyLabels = new List<Label>();
        private readonly List<Label> _secondaryLabels = new List<Label>();
        private readonly List<Label> _warningLabels = new List<Label>();

        /// <summary>Every label whose height comes from measuring its wrapped text.</summary>
        private readonly List<Label> _wrapped = new List<Label>();

        /// <summary>
        /// Creates a label AND registers it for theming and, when it wraps, for measurement.
        /// Registration happens here rather than in a list at the bottom of the file so that a
        /// label cannot exist without being themed - which was the failure mode of the fixed
        /// dialog this window replaces, whose four hardcoded arrays were a list somebody had to
        /// remember to update.
        /// </summary>
        private Label NewLabel(string text, LabelRole role, bool wrap)
        {
            var label = new Label
            {
                Text = text,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
            };

            if (wrap)
            {
                // A starting bound, replaced by the measured one on the first layout. Without
                // it the first pass reports one very long line.
                label.MaximumSize = new Size(Scaled(360), 0);
                _wrapped.Add(label);
            }

            switch (role)
            {
                case LabelRole.Secondary:
                    _secondaryLabels.Add(label);
                    break;
                case LabelRole.Warning:
                    _warningLabels.Add(label);
                    break;
                case LabelRole.Dynamic:
                    break;
                default:
                    _bodyLabels.Add(label);
                    break;
            }

            return label;
        }

        private static CheckBox NewCheck(string name, string text)
        {
            return new CheckBox
            {
                Name = name,
                Text = text,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
            };
        }

        private static Button NewButton(string text)
        {
            return NewButton(text, "btn" + text.Replace(" ", ""));
        }

        private static Button NewButton(string text, string name)
        {
            return new Button
            {
                Name = name,
                Text = text,
                AutoSize = false,
                // A caption that no longer fits ends in "..." rather than half a letter, which
                // is what a narrow window at 150% scaling would otherwise give.
                AutoEllipsis = true,
                UseVisualStyleBackColor = true,
            };
        }

        /// <summary>
        /// A single-column stack: the container shape every wrapped label in this window lives
        /// in, because <see cref="ReflowWrappedLabels"/> asks a label's PARENT how wide it is and
        /// that answer is only the label's own width when the parent has one column.
        /// </summary>
        private static TableLayoutPanel NewStack(string name, params Control[] rows)
        {
            var stack = new TableLayoutPanel
            {
                Name = name,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 0,
                Margin = Padding.Empty,
            };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            foreach (Control row in rows)
                AddRow(stack, row);
            return stack;
        }

        /// <summary>
        /// Appends one control as its own AutoSize row. Explicit rather than relying on
        /// GrowStyle, so the row count and the styles cannot drift apart.
        /// </summary>
        private static void AddRow(TableLayoutPanel stack, Control child)
        {
            stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            stack.RowCount = stack.RowStyles.Count;
            stack.Controls.Add(child, 0, stack.RowCount - 1);
        }

        /// <summary>
        /// A group box whose height comes from its contents. Dock=Fill inside an AutoSize row
        /// plus AutoSize on the box itself: the row asks the box how tall it wants to be, and
        /// the box asks its single-column stack.
        /// </summary>
        private static GroupBox NewGroup(string name, string caption, params Control[] rows)
        {
            var group = new GroupBox
            {
                Name = name,
                Text = caption,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
            };
            TableLayoutPanel inner = NewStack(name + "Inner", rows);
            inner.Dock = DockStyle.Fill;
            group.Controls.Add(inner);
            return group;
        }

        /// <summary>A left-to-right row of buttons, as tall as the buttons are.</summary>
        private static FlowLayoutPanel NewButtonRow(string name, params Button[] buttons)
        {
            var row = new FlowLayoutPanel
            {
                Name = name,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty,
            };
            foreach (Button button in buttons)
                row.Controls.Add(button);
            return row;
        }

        /// <summary>
        /// A scrolling viewport around a stack. AutoScroll rather than "make everything fit":
        /// tall content on a small laptop at 150% does not fit, and shrinking it to make it fit
        /// is the opposite of useful. Dock=Top on the stack means it is exactly as wide as the
        /// viewport, so a vertical scroll bar appearing cannot summon a horizontal one.
        /// </summary>
        private static Panel NewScroller(string name, Control content)
        {
            var page = new Panel
            {
                Name = name,
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Margin = Padding.Empty,
            };
            page.Controls.Add(content);
            return page;
        }

        /// <summary>
        /// One tab page, wrapping <paramref name="content"/> in a themed panel. Two reasons for
        /// the panel: a TabPage only honours its own BackColor while UseVisualStyleBackColor is
        /// false, and a plain Panel's colour is not something any part of WinForms is entitled to
        /// reinterpret. What the user sees behind the controls is the panel.
        /// </summary>
        private TabPage NewPage(string text, string name, Control content)
        {
            var surface = new Panel
            {
                Name = name + "Surface",
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
            };
            surface.Controls.Add(content);
            _pageSurfaces.Add(surface);

            var page = new TabPage(text)
            {
                Name = name,
                // False on purpose: with the visual style background the page paints light grey
                // in dark mode and no BackColor can reach it.
                UseVisualStyleBackColor = false,
                Padding = Padding.Empty,
                Margin = Padding.Empty,
            };
            page.Controls.Add(surface);
            return page;
        }

        private Control BuildFooter()
        {
            var footer = new TableLayoutPanel
            {
                Name = "footer",
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty,
            };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // RightToLeft flow: the FIRST control added sits furthest right, so this reads
            // "Apply now  Close" from the left once it is on screen.
            var commit = new FlowLayoutPanel
            {
                Name = "commitButtons",
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty,
            };
            commit.Controls.Add(btnClose);
            commit.Controls.Add(btnApply);

            footer.Controls.Add(_lblStatus, 0, 0);
            footer.Controls.Add(commit, 1, 0);
            return footer;
        }

        // ===== Metrics =====
        //
        // Design values are 96-DPI pixels for Segoe UI 9pt, whose line height is 15. Scaled()
        // moves them with the font the form actually got, which is how a display at 125% or
        // 150% ends up with a bigger layout instead of a clipped one. ApplyMetrics runs from the
        // constructor and again whenever the font changes, so nothing is baked in.

        private int Scaled(int designPixels)
        {
            return UiScale.ScaledFor(Font, designPixels);
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            // Also the hook the offline layout harness uses to render this window at the font a
            // 125% or 150% display would give it.
            ApplyMetrics();
        }

        private void ApplyMetrics()
        {
            if (_root == null)
                return;

            SuspendLayout();
            try
            {
                int pad = Scaled(10);
                int gap = Scaled(6);
                int rowHeight = Scaled(27);

                // The minimum has to stay inside the screen. A window whose MINIMUM height is
                // taller than the work area cannot have its Close button reached at all, and
                // that is exactly what a 150% display on a small laptop produces from a
                // comfortable 96-DPI number.
                Rectangle work = IsHandleCreated
                    ? Screen.FromControl(this).WorkingArea
                    : Screen.PrimaryScreen.WorkingArea;
                // Before the handle exists there is no chrome to measure, so estimate it. It is
                // only ever used to keep the window inside the work area.
                int chromeW = Math.Max(Width - ClientSize.Width, Scaled(16));
                int chromeH = Math.Max(Height - ClientSize.Height, Scaled(39));
                int roomW = Math.Max(Scaled(320), work.Width - chromeW);
                int roomH = Math.Max(Scaled(240), work.Height - chromeH);

                int minClientW = Math.Min(Scaled(640), roomW);
                int minClientH = Math.Min(Scaled(430), roomH);
                MinimumSize = new Size(minClientW + chromeW, minClientH + chromeH);
                ClientSize = new Size(
                    Math.Max(minClientW, Math.Min(Scaled(800), roomW)),
                    Math.Max(minClientH, Math.Min(Scaled(620), roomH)));

                _root.Padding = new Padding(pad, gap, pad, gap);

                foreach (string stackName in new[] { "outlookStack", "claudeStack", "updatesStack" })
                {
                    Control found = FindByName(this, stackName);
                    if (found != null)
                        found.Padding = new Padding(pad, gap, pad, gap);
                }

                foreach (GroupBox group in new[] { grpSearch, grpCaching, grpOst, grpModel, grpClaude, grpVersion })
                    group.Margin = new Padding(0, 0, 0, gap);

                // The picker is a fixed, scaled width rather than docked: full width on a wide
                // window would make a one-line choice look like a paragraph, and the drop-down
                // is given extra room of its own so a long label is never the reason an entry
                // is unreadable.
                cmbModel.Width = Scaled(360);
                cmbModel.DropDownWidth = Scaled(440);
                cmbModel.Margin = new Padding(0, 0, 0, gap);
                lblModelHelp.Margin = new Padding(0, 0, 0, gap);
                lblModelIdCaption.Margin = new Padding(0, 0, 0, 0);
                txtModel.Width = Scaled(280);
                txtModel.Margin = new Padding(0, Scaled(2), 0, gap);
                lblModelNote.Margin = new Padding(0, 0, 0, gap);
                lblModelState.Margin = new Padding(0, 0, 0, 0);

                lblHeader.Margin = new Padding(0, 0, 0, gap);
                chkMaster.Margin = new Padding(0, 0, 0, gap);
                lblRestart.Margin = new Padding(0, gap, 0, 0);
                lblGpo.Margin = new Padding(0, gap, 0, 0);

                // Indented to sit under the tick box's caption, the way it always has. The reflow
                // subtracts the margin, so an indented label wraps at the width it really has.
                lblGlobalMcpHelp.Margin = new Padding(Scaled(20), 0, 0, gap);
                lblMcp.Margin = new Padding(0, 0, 0, gap);
                lblUpdateError.Margin = new Padding(0, gap, 0, 0);

                SizeButton(btnAddProject, Scaled(176), rowHeight);
                SizeButton(btnCopyCommand, Scaled(176), rowHeight);
                SizeButton(btnCheckUpdates, Scaled(176), rowHeight);
                foreach (Button button in new[] { btnAddProject, btnCopyCommand, btnCheckUpdates })
                    button.Margin = new Padding(0, gap, gap, 0);

                foreach (Button button in new[] { btnApply, btnClose })
                {
                    SizeButton(button, Scaled(88), rowHeight);
                    button.Margin = new Padding(gap, gap, 0, 0);
                }

                ApplyPromptMetrics(pad, gap, rowHeight);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings metrics: " + ex.Message);
            }
            finally
            {
                ResumeLayout(true);
            }
        }

        /// <summary>
        /// A button at its designed width, grown wherever its caption no longer fits inside it.
        /// The grow-to-fit half is what stops "Add to a specific project…" turning into
        /// "Add to a speci..." the moment a font substitution or a scaled display lands.
        /// </summary>
        private static void SizeButton(Button button, int designWidth, int height)
        {
            int width = designWidth;
            try
            {
                width = Math.Max(designWidth, button.GetPreferredSize(Size.Empty).Width);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings button size: " + ex.Message);
            }
            button.Size = new Size(width, height);
        }

        private static Control FindByName(Control parent, string name)
        {
            foreach (Control child in parent.Controls)
            {
                if (child.Name == name)
                    return child;
                Control deeper = FindByName(child, name);
                if (deeper != null)
                    return deeper;
            }
            return null;
        }

        // ===== Wrapped-label measurement =====
        //
        // The lesson of the fixed dialog's clipped help text: a wrapped label's height is
        // MEASURED, never written down, because the same sentence needs two lines at 96 DPI and
        // three at 120. Here the measuring is done by the label itself - AutoSize with a bounded
        // MaximumSize.Width makes Label.GetPreferredSize run TextRenderer.MeasureText with
        // WordBreak, which is the same measurement, and it feeds the AutoSize rows above it. All
        // this has to do is keep that bound equal to the width actually available, which changes
        // every time the user drags the window edge.

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            ReflowWrappedLabels();
        }

        private void ReflowWrappedLabels()
        {
            if (_reflowing || _disposedCustom || IsDisposed)
                return;

            _reflowing = true;
            try
            {
                foreach (Label label in _wrapped)
                {
                    Control parent = label.Parent;
                    if (parent == null)
                        continue;

                    // Every wrapped label lives in a single-column container by construction,
                    // so the parent's client width IS the width this label has to wrap into.
                    int available = parent.ClientSize.Width
                                    - parent.Padding.Horizontal
                                    - label.Margin.Horizontal;
                    if (available < Scaled(80))
                        available = Scaled(80);

                    if (label.MaximumSize.Width != available)
                        label.MaximumSize = new Size(available, 0);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings reflow: " + ex.Message);
            }
            finally
            {
                _reflowing = false;
            }
        }

        /// <summary>
        /// Re-measures and re-lays the window out after text changed underneath it. Cheap, and
        /// never throws - a layout that failed leaves the last good one on screen.
        /// </summary>
        private void RelayoutAfterTextChange()
        {
            if (_disposedCustom || IsDisposed)
                return;
            try
            {
                ReflowWrappedLabels();
                PerformLayout();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings layout: " + ex.Message);
            }
        }

        private void OnPageChanged(object sender, EventArgs e)
        {
            // A tab page gets its real size only when it becomes the selected one, and a wrapped
            // label's height is measured from the width it has. Re-measure now, so the page the
            // user is about to look at is not laid out for the width it had a moment ago.
            ReflowWrappedLabels();
        }

        /// <summary>
        /// Selects a tab from code. The user's own tab clicks and Ctrl+Tab do not come through
        /// here - they are the tab control's business - so this only exists for the moments the
        /// window has to move the user itself: a validation failure, which is always about a
        /// button, and Add, which has just made one.
        /// </summary>
        private void ShowPage(TabPage page)
        {
            try
            {
                _pageTabs.SelectedTab = page;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings page: " + ex.Message);
            }
        }

        // ===== Commit =====

        /// <summary>
        /// "Apply now". Three jobs, and the order matters: the instant-apply half is reconciled
        /// first because it never fails visibly, and the buffered half goes last because a
        /// rejected button name opens a message box and moves the user to the Buttons tab, which
        /// has to be the thing they are left looking at.
        /// </summary>
        private void OnApplyNow(object sender, EventArgs e)
        {
            // A model id still in the box counts as asked for: "Apply now" is the one button on
            // this window that means "do the thing I typed", so it lands before the refresh
            // below reads the stored value back.
            CommitCustomModel();

            // Same button, same promise as before: re-check the Outlook tuning AND the
            // mail-server registration, so a drift the user just fixed (installing the runtime,
            // say) is picked up without restarting Outlook. Neither throws out of its public
            // surface.
            OutlookTuningService.ReconcileFromUi();
            McpRegistrationService.Reconcile();
            RefreshFromState();

            if (HasChanges())
                Commit();
        }

        private void CloseWithoutAsking()
        {
            _closingWithoutAsking = true;
            try
            {
                Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings close: " + ex.Message);
                _closingWithoutAsking = false;
            }
        }

        /// <summary>
        /// True when it is all right to throw the prompt and button drafts away. Asked once, by
        /// whichever of Close, Escape, the X or Alt+F4 got there first - never twice, because
        /// everything that asks then closes through <see cref="CloseWithoutAsking"/>. The tuning
        /// tick boxes are not in the question: they wrote when they were clicked, so there is
        /// nothing about them left to discard.
        /// </summary>
        private bool ConfirmDiscard()
        {
            if (!HasChanges())
                return true;

            DialogResult answer;
            try
            {
                answer = MessageBox.Show(
                    this,
                    "Your changes to the prompts and buttons have not been saved. Close anyway?",
                    "OutlookAI",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    // Keeping the window open is the answer that loses nothing.
                    MessageBoxDefaultButton.Button2);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings discard prompt: " + ex.Message);
                return true;
            }

            return answer == DialogResult.Yes;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_closingWithoutAsking && !ConfirmDiscard())
            {
                e.Cancel = true;
                return;
            }

            // Not part of the discard question, and deliberately on the other side of it: the
            // model box is instant-apply text, so closing the window is the last chance to take
            // what is in it. A focus change would normally have committed it already; this is
            // for the case where the window closes while the box still has focus.
            CommitCustomModel();

            base.OnFormClosing(e);
        }

        // ===== Outlook tuning =====

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
                Debug.WriteLine("Settings toggle: " + ex.Message);
            }
            RefreshFromState();
        }

        // The "all my projects" toggle. Ticking or unticking it IS the user declaring their
        // intent, so it applies immediately rather than waiting for Apply now - a user who
        // unticks it expects the entry gone now - and it never re-opens the question the startup
        // prompt asks: they just answered it.
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
                Debug.WriteLine("MCP toggle: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            RefreshFromState();
        }

        // ===== Model =====
        //
        // Instant-apply, like everything else on this tab: a picked list entry IS the decision,
        // so it writes now. The exception is the custom id, which is typed - see rule 1 in the
        // class comment. Nothing here is ever buffered into the prompt drafts, so nothing here
        // is part of the unsaved-changes question.

        private void OnModelChoiceChanged(object sender, EventArgs e)
        {
            if (_updating)
                return;

            var choice = cmbModel.SelectedItem as ModelChoice;
            if (choice == null)
                return;

            if (choice.Kind == ModelChoiceKind.Custom)
            {
                // Picking "Custom" stores NOTHING. It opens the box; what is typed there is the
                // decision, and it is not made yet. Whatever was stored before stays stored
                // until it is, so the pane keeps working on the model it already had.
                txtModel.Enabled = true;
                RefreshModelLines();
                RelayoutAfterTextChange();
                try
                {
                    txtModel.Focus();
                    txtModel.SelectAll();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Model focus: " + ex.Message);
                }
                return;
            }

            txtModel.Enabled = false;
            // Alias, or the deferring default whose Alias is null and therefore clears the value.
            ClaudeService.SetConfiguredModel(choice.Alias);
            RefreshModelLines();
            RelayoutAfterTextChange();
        }

        /// <summary>
        /// Paints one picker entry, in the closed control and in the dropped-down list alike.
        ///
        /// This exists because <see cref="ApplyThemeTo"/> cannot reach a ComboBox: it takes the
        /// BackColor it is given and then lets the OS draw the entry anyway, which in dark mode
        /// produced a white field. Themed by ROLE like everything else - the field colour for an
        /// entry, the accent for the one the mouse or keyboard is on.
        /// </summary>
        private void OnDrawModelChoice(object sender, DrawItemEventArgs e)
        {
            try
            {
                bool highlighted = (e.State & DrawItemState.Selected) == DrawItemState.Selected
                                   && (e.State & DrawItemState.ComboBoxEdit) != DrawItemState.ComboBoxEdit;
                Color back = highlighted ? ThemeService.Accent : ThemeService.TextBoxBackground;
                // White on both accent shades, and the only place in this window that needs a
                // colour the theme does not carry: an accent is a background, and neither
                // Text nor ButtonText is guaranteed to be legible on one.
                Color fore = highlighted ? Color.White : ThemeService.Text;

                using (var brush = new SolidBrush(back))
                    e.Graphics.FillRectangle(brush, e.Bounds);

                if (e.Index >= 0 && e.Index < cmbModel.Items.Count)
                {
                    TextRenderer.DrawText(
                        e.Graphics,
                        Convert.ToString(cmbModel.Items[e.Index]),
                        e.Font ?? Font,
                        e.Bounds,
                        fore,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter
                        | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
                }

                e.DrawFocusRectangle();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Model item draw: " + ex.Message);
            }
        }

        private void OnModelIdEdited(object sender, EventArgs e)
        {
            if (_updating)
                return;
            // Typing only re-reads the advice. Nothing is written, and nothing is claimed about
            // a half-typed id beyond whether it could be one at all.
            RefreshModelLines();
            RelayoutAfterTextChange();
        }

        private void OnModelIdKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;
            // Enter is the other way to say "that is the id". Handled here rather than left to a
            // default button, because this form deliberately has none, and an unhandled Enter in
            // a single-line box makes Windows beep.
            e.Handled = true;
            e.SuppressKeyPress = true;
            CommitCustomModel();
        }

        /// <summary>
        /// Stores what the custom box holds, if it holds a usable model id. Called when the box
        /// loses focus, on Enter, from "Apply now", and from the window closing - never per
        /// keystroke, because "claude-op" is not a model anybody meant to choose.
        ///
        /// A BLANK BOX DOES NOT CLEAR THE STORED MODEL. Clearing is what the first list entry is
        /// for, one click away and labelled; emptying a text box is at least as likely to be
        /// somebody halfway through retyping.
        /// </summary>
        private void CommitCustomModel()
        {
            if (_updating || _disposedCustom || IsDisposed)
                return;

            var choice = cmbModel.SelectedItem as ModelChoice;
            if (choice == null || choice.Kind != ModelChoiceKind.Custom)
                return;

            string typed = txtModel.Text == null ? "" : txtModel.Text.Trim();
            if (typed.Length == 0 || !ModelCatalog.IsWellFormedModelId(typed))
            {
                // The note already says why nothing was stored.
                RefreshModelLines();
                RelayoutAfterTextChange();
                return;
            }

            // Unchanged means untouched: writing it again would throw away a perfectly good
            // pre-warmed process for nothing.
            if (!string.Equals(typed, ClaudeService.ConfiguredModel, StringComparison.Ordinal))
                ClaudeService.SetConfiguredModel(typed);

            RefreshModelLines();
            RelayoutAfterTextChange();
        }

        /// <summary>
        /// Puts the stored choice on screen: the matching list entry, or Custom with the id in
        /// the box when it is not one of the aliases. Absent means the first entry, which is the
        /// shipped default and stores nothing.
        /// </summary>
        private void LoadModelFromStore()
        {
            string stored = ClaudeService.ConfiguredModel;

            _updating = true;
            try
            {
                int index = 0;
                if (stored != null)
                {
                    // Custom unless an alias claims it. Ordinal-ignore-case: the aliases are
                    // lower case and a hand-edited registry value need not be.
                    index = cmbModel.Items.Count - 1;
                    for (int i = 0; i < cmbModel.Items.Count; i++)
                    {
                        var choice = cmbModel.Items[i] as ModelChoice;
                        if (choice == null || choice.Kind != ModelChoiceKind.Alias)
                            continue;
                        if (string.Equals(choice.Alias, stored, StringComparison.OrdinalIgnoreCase))
                        {
                            index = i;
                            break;
                        }
                    }
                }

                cmbModel.SelectedIndex = index;
                bool custom = index == cmbModel.Items.Count - 1;
                txtModel.Enabled = custom;
                if (custom && stored != null)
                    txtModel.Text = stored;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Model load: " + ex.Message);
            }
            finally
            {
                _updating = false;
            }

            RefreshModelLines();
        }

        /// <summary>
        /// The two lines under the picker: what will actually be sent, and the one honest thing
        /// there is to say about a hand-typed id.
        ///
        /// The state line quotes the ARGUMENT, not a friendly paraphrase, because the whole
        /// point of the default is that no argument is sent and a paraphrase would hide exactly
        /// that. The note appears only for the custom box - an alias is not a model id and is
        /// not checked against the known families, because Claude Code resolves it and the
        /// families list has nothing to say about it.
        /// </summary>
        private void RefreshModelLines()
        {
            try
            {
                string stored = ClaudeService.ConfiguredModel;
                string sending = stored == null
                    ? "OutlookAI sends no --model argument, so Claude Code picks."
                    : "OutlookAI sends: --model \"" + stored + "\"";

                string note = null;
                string pending = null;
                var choice = cmbModel.SelectedItem as ModelChoice;
                if (choice != null && choice.Kind == ModelChoiceKind.Custom)
                {
                    string typed = txtModel.Text == null ? "" : txtModel.Text.Trim();
                    if (typed.Length == 0)
                        note = null;
                    else if (!ModelCatalog.IsWellFormedModelId(typed))
                        note = ModelCatalog.MalformedModelNote;
                    else if (!ModelCatalog.IsKnownFamily(typed))
                        note = ModelCatalog.UnverifiedModelNote;

                    // Says out loud that a typed id has not landed yet. Without it the box and
                    // the line below it read as a contradiction - the box showing one model, the
                    // line naming another - when in fact one is a decision and one is not made.
                    if (ModelCatalog.IsWellFormedModelId(typed)
                        && !string.Equals(typed, stored, StringComparison.Ordinal))
                    {
                        pending = "Press Enter, or click outside the box, to start using \""
                                  + typed + "\" - for now, "
                                  + (stored == null
                                      ? "no --model argument is sent and Claude Code picks."
                                      : "OutlookAI still sends --model \"" + stored + "\".");
                    }
                }

                SetLabel(lblModelState, pending == null ? sending : pending);

                SetLabel(lblModelNote, note == null ? "" : note);
                lblModelNote.Visible = note != null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Model lines: " + ex.Message);
            }
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
                              + "answer it - the add-in deliberately does not.");
                sb.AppendLine();
                sb.Append(".mcp.json is normally committed to source control. ");
                if (McpConfigEditor.ContainsEnvironmentReference(_preferredCommand))
                {
                    sb.Append("Because the entry points at ${LOCALAPPDATA} rather than a fixed path, it is "
                              + "portable: teammates who have OutlookAI installed get a working mail server, "
                              + "and teammates who do not simply see a failed-connection warning for this one "
                              + "server - nothing else breaks.");
                }
                else
                {
                    sb.Append("This entry names a fixed path on this machine (the mail server is not in the "
                              + "default install location here), so it will not resolve on a teammate's "
                              + "machine - they would see a failed-connection warning for this one server.");
                }

                MessageBox.Show(this, sb.ToString(), "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Add to project: " + ex.Message);
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
                Debug.WriteLine("Copy command: " + ex.Message);
            }
        }

        // Reads only what the last reconcile recorded, so opening the window never touches
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
                               + (string.IsNullOrEmpty(reg.Detail) ? "." : (" - " + reg.Detail));
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
                Debug.WriteLine("MCP line refresh: " + ex.Message);
            }
        }

        private static bool IsMcpProblem(string status)
        {
            return status == McpRegistrationService.StatusNoRuntime
                || status == McpRegistrationService.StatusParseFailed
                || status == McpRegistrationService.StatusError;
        }

        // ===== Version and updates =====

        /// <summary>
        /// Repaints the version line and the update-error line from <see cref="UpdateService"/>,
        /// and returns whether the window now needs laying out again - which is NOT the same as
        /// "the text changed". New text that wraps to the same number of lines occupies exactly
        /// the same box, so it is simply painted and nothing below it moves.
        ///
        /// The gate matters less than it used to and is kept anyway. On the fixed dialog this
        /// window replaces, a re-layout reset the scrolled viewport to the top, and the version
        /// line changes once a minute all on its own as "checked 4m ago" becomes 5m - so a free
        /// re-layout meant the dialog scrolled itself back to the top every minute while the user
        /// was reading it. Here the line has a tab to itself and the layout is docked, but the
        /// tick still runs once a second for the life of the window and there is no reason to
        /// spend a full measure-and-arrange pass on a string that did not move. Never throws.
        /// </summary>
        private bool RefreshVersionLine()
        {
            if (_disposedCustom || IsDisposed)
                return false;
            try
            {
                string line = UpdateService.VersionLine();
                string error = UpdateService.LastError;
                bool showError = !string.IsNullOrEmpty(error);
                string errorText = showError ? "Last update check failed: " + error : "";

                // Cheap and idempotent, so it is settled on every tick rather than only when
                // the text moved: a check started from the sidebar disables this button too.
                btnCheckUpdates.Enabled = !_checkInFlight && !UpdateService.IsChecking;

                // A line appearing or disappearing always moves everything under it.
                bool relayout = _showUpdateError != showError;

                relayout |= SetMeasuredText(lblVersion, line);
                relayout |= SetMeasuredText(lblUpdateError, errorText);

                _showUpdateError = showError;
                lblUpdateError.Visible = showError;
                return relayout;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Version line refresh: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Sets <paramref name="label"/>'s text and answers whether that changed how much room
        /// it needs - measured the same way the label itself will measure it, at the wrap width
        /// it currently has, so the two cannot disagree about whether a re-layout is due.
        /// </summary>
        private static bool SetMeasuredText(Label label, string text)
        {
            if (label.Text == text)
                return false;
            int width = label.MaximumSize.Width > 0 ? label.MaximumSize.Width : Math.Max(1, label.Width);
            int before = MeasureWrapped(label, label.Text, width);
            label.Text = text;
            return MeasureWrapped(label, text, width) != before;
        }

        /// <summary>
        /// How tall <paramref name="text"/> is once wrapped at <paramref name="width"/>, measured
        /// with the font it will actually paint with rather than the font someone had in mind
        /// while writing a number down.
        /// </summary>
        private static int MeasureWrapped(Label label, string text, int width)
        {
            if (width <= 0)
                return 0;
            try
            {
                return TextRenderer.MeasureText(
                    text ?? "",
                    label.Font,
                    new Size(width, int.MaxValue),
                    TextFormatFlags.WordBreak).Height;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings measure: " + ex.Message);
                return label.Height;
            }
        }

        // "Check for updates": the UpdateService poll, on demand. async void is what an event
        // handler is, and it swallows everything - a failed check belongs on the error line
        // above the button, not in a crash dialog.
        private async void OnCheckForUpdates(object sender, EventArgs e)
        {
            if (_checkInFlight)
                return;
            _checkInFlight = true;
            if (RefreshVersionLine())
                RelayoutAfterTextChange();

            try
            {
                await UpdateService.CheckNowAsync();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Manual update check: " + ex.Message);
            }
            finally
            {
                // The await can outlive the window - RefreshVersionLine is a no-op once it has
                // gone, and the flag is only read from here.
                _checkInFlight = false;
                if (RefreshVersionLine())
                    RelayoutAfterTextChange();
            }
        }

        // ===== State =====

        private void RefreshFromState()
        {
            // The Claude Code tab first and in its own guarded block: a tuning snapshot that
            // fails must not leave the registration controls unpainted.
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
                Debug.WriteLine("MCP settings refresh: " + ex.Message);
            }
            finally
            {
                _updating = false;
            }

            // Its own guarded block for the same reason: a registry read that failed must not
            // leave the model picker showing a choice nobody made. It manages _updating itself.
            LoadModelFromStore();

            RefreshMcpLine();
            // Result ignored on purpose: this method re-lays the window out at the end either way.
            RefreshVersionLine();

            OutlookTuningService.TuningSnapshot snap;
            try
            {
                snap = OutlookTuningService.GetSnapshot();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings refresh: " + ex.Message);
                // The Claude Code tab above was still repainted, so lay out around it.
                RelayoutAfterTextChange();
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

                lblRestart.Visible = snap.RestartNeeded;

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
                    lblGpo.Visible = true;
                }
                else
                {
                    lblGpo.Visible = false;
                }
            }
            finally
            {
                _updating = false;
            }

            // Last, and always: every label above may have just changed length, and the two
            // status lines may have just appeared or gone.
            RelayoutAfterTextChange();
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

        // ===== Theming =====

        private void OnThemeChanged(object sender, EventArgs e)
        {
            // ThemeService raises this from SystemEvents and from its own registry watcher, so
            // it can arrive on a thread that has no business touching controls.
            if (_disposedCustom || IsDisposed || !IsHandleCreated)
                return;
            try { BeginInvoke((Action)ApplyTheme); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void ApplyTheme()
        {
            if (_disposedCustom || IsDisposed)
                return;

            SuspendLayout();
            try
            {
                BackColor = ThemeService.Background;
                ForeColor = ThemeService.Text;

                // By TYPE, over the real control tree: nothing can fall off a list, because
                // there is no list.
                ApplyThemeTo(this);

                // By ROLE, which a type cannot imply. Registered by NewLabel at creation.
                foreach (Label label in _bodyLabels)
                    label.ForeColor = ThemeService.Text;
                foreach (Label label in _secondaryLabels)
                    label.ForeColor = ThemeService.SecondaryText;
                foreach (Label label in _warningLabels)
                    label.ForeColor = ThemeService.StatusError;

                // The two dynamic lines keep their meaning across a theme switch: the
                // mail-server line is red only while there is a problem, and the footer status
                // is green after a save and red after a failed one.
                RefreshMcpLine();
                _lblStatus.ForeColor = StatusColour();

                // The tab strip and the page frame are painted, not coloured, so the flip has to
                // reach the paint: the surfaces are set explicitly and the strip is invalidated.
                // Deliberately NOT setting _pageTabs.BackColor - TabControl overrides it to
                // return SystemColors.Control and ignores the setter, which is one more face of
                // the same trap and the reason the strip is painted at all.
                _pageTabs.ForeColor = ThemeService.Text;
                foreach (Control surface in _pageSurfaces)
                    surface.BackColor = ThemeService.Background;
                _pageTabs.Invalidate();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Settings theme: " + ex.Message);
            }
            finally
            {
                ResumeLayout(true);
            }

            // A theme switch is also when a font substitution would land, so re-measure rather
            // than assume the wrapped labels still fit.
            ReflowWrappedLabels();
        }

        private void ApplyThemeTo(Control parent)
        {
            foreach (Control control in parent.Controls)
            {
                var textBox = control as TextBox;
                if (textBox != null)
                {
                    textBox.BackColor = ThemeService.TextBoxBackground;
                    textBox.ForeColor = ThemeService.Text;
                    textBox.BorderStyle = ThemeService.IsDarkMode
                        ? BorderStyle.FixedSingle
                        : BorderStyle.Fixed3D;
                }

                // A ComboBox ignores BackColor entirely while it is painted by the visual style,
                // so dark mode also has to take the painting off it. FlatStyle.Flat is what
                // makes the colour it is given the colour it draws.
                var combo = control as ComboBox;
                if (combo != null)
                {
                    combo.BackColor = ThemeService.TextBoxBackground;
                    combo.ForeColor = ThemeService.Text;
                    combo.FlatStyle = ThemeService.IsDarkMode ? FlatStyle.Flat : FlatStyle.Standard;
                }

                var listBox = control as ListBox;
                if (listBox != null)
                {
                    listBox.BackColor = ThemeService.TextBoxBackground;
                    listBox.ForeColor = ThemeService.Text;
                    listBox.BorderStyle = ThemeService.IsDarkMode
                        ? BorderStyle.FixedSingle
                        : BorderStyle.Fixed3D;
                }

                var group = control as GroupBox;
                if (group != null)
                    group.ForeColor = ThemeService.Text;

                var check = control as CheckBox;
                if (check != null)
                    check.ForeColor = ThemeService.Text;

                // A TabPage is the one control in this window that will not inherit a BackColor:
                // it reads its own, and only while UseVisualStyleBackColor is false.
                var tabPage = control as TabPage;
                if (tabPage != null)
                {
                    tabPage.UseVisualStyleBackColor = false;
                    tabPage.BackColor = ThemeService.Background;
                    tabPage.ForeColor = ThemeService.Text;
                }

                var button = control as Button;
                if (button != null)
                {
                    if (ThemeService.IsDarkMode)
                    {
                        button.FlatStyle = FlatStyle.Flat;
                        button.FlatAppearance.BorderColor = ThemeService.Border;
                        button.BackColor = ThemeService.ButtonFace;
                        button.ForeColor = ThemeService.ButtonText;
                        button.UseVisualStyleBackColor = false;
                    }
                    else
                    {
                        button.FlatStyle = FlatStyle.Standard;
                        button.UseVisualStyleBackColor = true;
                        button.ForeColor = ThemeService.ButtonText;
                    }
                }

                ApplyThemeTo(control);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposedCustom)
            {
                _disposedCustom = true;
                // Static event: a subscription left behind roots this window for the life of the
                // process, and there is one of these per Outlook session.
                ThemeService.ThemeChanged -= OnThemeChanged;
                try
                {
                    if (_versionTimer != null)
                    {
                        _versionTimer.Stop();
                        _versionTimer.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Settings dispose: " + ex.Message);
                }
            }
            base.Dispose(disposing);
        }
    }
}
