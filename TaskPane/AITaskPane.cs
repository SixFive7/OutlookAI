using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane
{
    public partial class AITaskPane : UserControl
    {
        private readonly bool _isInlineResponse;
        private Outlook.Inspector _owningInspector;
        private readonly Timer _versionTimer;
        private ToolTip _toolTip;
        private bool _disposed;

        /// <summary>
        /// The one font every quick button shares. Owned by the pane, not by the buttons: the
        /// button set is rebuilt whenever the user edits it, and a font per button would leak one
        /// handle per button per rebuild for the life of the process.
        /// </summary>
        private Font _quickButtonFont;

        /// <summary>
        /// Guards <see cref="LayoutPane"/> against itself. It resizes the very group whose
        /// SizeChanged brings it back, so without this the first pass would recurse.
        /// </summary>
        private bool _layingOut;

        /// <summary>Set when a pass was re-entered, meaning its own result changed its input.</summary>
        private bool _layoutDirty;

        /// <summary>Quick-action group width the last layout pass worked from.</summary>
        private int _layoutWidth = -1;

        // Iterative editing state
        private readonly List<EditTurn> _editHistory = new List<EditTurn>();
        private bool _freshDraft;
        private bool _isProcessing;

        // D38: "Select the best signature" needs at least one installed signature.
        private bool _signaturesAvailable;

        // Debug: 7 clicks within 3 seconds to enable
        private bool _debug;
        private int _debugClickCount;
        private DateTime _debugFirstClick;
        private readonly StringBuilder _debugLog = new StringBuilder();
        private bool _lastStatusError;

        /// <summary>
        /// Set between "check for updates" being clicked and that check completing. Held apart
        /// from <see cref="UpdateService.IsChecking"/> so the link stays off across the whole
        /// round trip, including a check that ended the instant it began — a developer build,
        /// or one already running when the click landed.
        /// </summary>
        private bool _checkInFlight;

        public AITaskPane(bool isInlineResponse = false, Outlook.Inspector inspector = null)
        {
            _isInlineResponse = isInlineResponse;
            _owningInspector = inspector;
            InitializeComponent();
            // Order matters twice over: the quick buttons ask _toolTip for their tip as they are
            // built, so SetupTooltips comes first, and ApplyTheme has to see the buttons that
            // RebuildQuickActions created or the pane opens with unthemed buttons in dark mode.
            SetupTooltips();
            RebuildQuickActions();
            ApplyTheme();
            ThemeService.ThemeChanged += OnThemeChanged;
            PromptStore.Changed += OnPromptsChanged;
            // A width change reaches the buttons only through here: the group is anchored, so
            // widening the pane widens it, and how many buttons fit on a row follows from that.
            grpQuickActions.SizeChanged += grpQuickActions_SizeChanged;
            this.FontChanged += AITaskPane_FontChanged;
            RefreshSignatureAvailability();
            lblVersion.Click += lblVersion_Click;
            lblVersion.DoubleClick += lblVersion_Click;

            _versionTimer = new Timer();
            _versionTimer.Interval = 1000;
            _versionTimer.Tick += (s, ev) => UpdateVersionLabel();
            _versionTimer.Start();
            UpdateVersionLabel();
        }

        // The wording moved into UpdateService so this indicator and the one in OutlookAI
        // Settings cannot describe the same update state in two different ways.
        private void UpdateVersionLabel()
        {
            if (_disposed || IsDisposed) return;

            // Debug mode takes the version label over as its copy button, so its text is left
            // alone from then on. The timer is stopped by that point, but a manual check still
            // comes back through here.
            if (!_debug)
                lblVersion.Text = UpdateService.VersionLine();
            lnkUpdateError.Visible = UpdateService.LastError != null;
            // Off while a check is running, wherever it was started from — including one
            // started in the settings dialog.
            lnkCheckUpdates.Enabled = !_checkInFlight && !UpdateService.IsChecking;
        }

        // "check for updates": the ten-minute poll, on demand, and the same trigger the
        // "Version and updates" group in OutlookAI Settings offers. async void is what an event
        // handler is, and it swallows everything — the outcome of a check belongs on the version
        // line above, not in a crash dialog.
        private async void lnkCheckUpdates_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_checkInFlight)
                return;
            _checkInFlight = true;
            UpdateVersionLabel();
            try
            {
                await UpdateService.CheckNowAsync();
            }
            catch (Exception ex)
            {
                DebugLog("CheckNow", null, ex.Message);
            }
            finally
            {
                // The await can outlive the pane — Outlook disposes it when the inspector
                // closes. UpdateVersionLabel is a no-op by then.
                _checkInFlight = false;
                UpdateVersionLabel();
            }
        }

        private void lblVersion_Click(object sender, EventArgs e)
        {
            if (_debug) return;

            var now = DateTime.Now;
            if (_debugClickCount == 0 || (now - _debugFirstClick).TotalSeconds > 3)
            {
                _debugClickCount = 0;
                _debugFirstClick = now;
            }

            _debugClickCount++;
            if (_debugClickCount >= 7)
            {
                _debug = true;
                _debugLog.Clear();
                _versionTimer.Stop();
                lblVersion.Text = "Debug enabled (click to copy)";
                lblVersion.Click -= lblVersion_Click;
                lblVersion.Click += (s, ev) =>
                {
                    if (_debugLog.Length > 0)
                        Clipboard.SetText(_debugLog.ToString());
                };
            }
        }

        private void DebugLog(string label, dynamic doc = null, string extra = null)
        {
            if (!_debug) return;
            try
            {
                _debugLog.AppendLine($"=== {label} at {DateTime.Now:HH:mm:ss} ===");

                if (doc != null)
                {
                    var content = doc.Content;
                    _debugLog.AppendLine($"Content.End = {content.End}");
                    ThisAddIn.ReleaseCom(content);

                    var bmks = doc.Bookmarks;
                    bmks.ShowHidden = true;
                    foreach (var bmName in new[] { "_MailAutoSig", "_MailOriginal" })
                    {
                        if (bmks.Exists(bmName))
                        {
                            var bmk = bmks[bmName];
                            var range = bmk.Range;
                            int start = range.Start, end = range.End;
                            string text = range.Text ?? "";
                            ThisAddIn.ReleaseCom(range);
                            ThisAddIn.ReleaseCom(bmk);
                            if (text.Length > 200) text = text.Substring(0, 200) + "...";
                            _debugLog.AppendLine($"  {bmName}: [{start}, {end}] = {text}");
                        }
                        else
                        {
                            _debugLog.AppendLine($"  {bmName}: NOT FOUND");
                        }
                    }
                    ThisAddIn.ReleaseCom((object)bmks);

                    int draftEnd = FindDraftEnd(doc);
                    var draftRange = doc.Range(0, draftEnd);
                    string draft = draftRange.Text ?? "";
                    ThisAddIn.ReleaseCom(draftRange);
                    if (draft.Length > 300) draft = draft.Substring(0, 300) + "...";
                    _debugLog.AppendLine($"  Draft [0, {draftEnd}]: {draft}");
                }

                if (extra != null)
                    _debugLog.AppendLine($"  {extra}");

                _debugLog.AppendLine();
            }
            catch (Exception ex)
            {
                _debugLog.AppendLine($"  DEBUG ERROR: {ex.Message}");
            }
        }

        private void lnkUpdateError_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var error = UpdateService.LastError;
            if (error != null)
                MessageBox.Show(error, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public void ResetForNewEmail()
        {
            txtPrompt.Text = "";
            lblStatus.Visible = false;
            _editHistory.Clear();
            _freshDraft = false;
            RefreshSignatureAvailability();
        }

        /// <summary>
        /// The signature button is enabled only while at least one signature is
        /// installed (cheap directory listing; re-checked when the pane is reused).
        /// </summary>
        private void RefreshSignatureAvailability()
        {
            _signaturesAvailable = SignatureStore.AnySignatureInstalled();
            btnSelectSignature.Enabled = _signaturesAvailable && !_isProcessing;
        }

        // === Quick actions: built from the store, laid out by measurement ===

        // The design-time units the layout is written in, at the 96 DPI / default-font baseline
        // the Designer's own literals assume. LayoutPane scales them by UiScale, because nothing
        // computed at runtime goes through the AutoScaleMode.Font pass the Designer's do.
        private const int PaneMargin = 10;             // pane and group inner left margin
        private const int GroupTopInset = 22;          // first row, clear of the group caption
        private const int QuickActionsDesignTop = 40;  // the group's own Y, set in the Designer
        private const int GroupBottomInset = 8;        // below a group's last control
        private const int ControlGap = 5;              // between neighbours, and between groups
        private const int MinQuickButtonWidth = 70;    // narrowest quick button; sets the columns
        private const int MinQuickButtonHeight = 28;
        private const int QuickButtonPadding = 10;     // button chrome around its caption

        /// <summary>Measuring a caption the way a Button paints it: wrapped, not clipped.</summary>
        private const TextFormatFlags QuickButtonMeasureFlags = TextFormatFlags.WordBreak;

        /// <summary>
        /// Replaces the quick buttons with one per entry in the user's saved button set, then lays
        /// the pane out around however many that turned out to be. Zero is a legitimate answer -
        /// the user is allowed to delete every button - and it renders as a group holding nothing
        /// but the signature action.
        ///
        /// The old buttons are disposed and unhooked rather than just removed: this runs again
        /// every time the button set is edited, for the whole life of a compose window.
        /// </summary>
        private void RebuildQuickActions()
        {
            RebuildQuickActions(PromptStore.GetButtons());
        }

        /// <summary>
        /// The rebuild itself, over a given set. Split from the store read above so the layout can
        /// be driven over button counts and caption lengths that whatever is in HKCU right now does
        /// not happen to cover - zero buttons, twenty of them, and a name at the store's
        /// 64-character cap all have to come out laid out rather than clipped.
        /// </summary>
        private void RebuildQuickActions(IList<PromptButton> buttons)
        {
            flowQuickActions.SuspendLayout();
            try
            {
                var stale = new List<Control>();
                foreach (Control ctrl in flowQuickActions.Controls)
                    stale.Add(ctrl);
                flowQuickActions.Controls.Clear();
                foreach (Control ctrl in stale)
                {
                    var btn = ctrl as Button;
                    if (btn != null)
                        btn.Click -= QuickAction_Click;
                    // Drop the tip too: ToolTip keeps its own map of the controls it serves, and
                    // this pane's ToolTip outlives every button set it ever showed.
                    if (_toolTip != null)
                        _toolTip.SetToolTip(ctrl, null);
                    ctrl.Dispose();
                }

                if (_quickButtonFont == null)
                    _quickButtonFont = new Font("Segoe UI", 8F);

                int tabIndex = 0;
                foreach (PromptButton button in buttons ?? new List<PromptButton>())
                {
                    var btn = new Button();
                    // Wrap where the caption can be broken, ellipsis where it cannot. A name at the
                    // store's 64-character cap with no spaces in it does not fit any button this
                    // pane can offer at any width, so the honest end of that road is "MMMM..." plus
                    // a tooltip, not a caption sliced off mid-glyph.
                    btn.AutoEllipsis = true;
                    btn.Font = _quickButtonFont;
                    btn.Text = button.Name;
                    // The instruction, and also the label this turn goes into the edit history
                    // under - one string, so what was sent and what is recorded cannot diverge.
                    btn.Tag = button.Prompt;
                    btn.TabIndex = tabIndex++;
                    btn.UseVisualStyleBackColor = true;
                    btn.Click += QuickAction_Click;
                    flowQuickActions.Controls.Add(btn);
                    // The tip IS the prompt. It used to be a second hand-written copy of the same
                    // wording, which could only ever drift once the prompt became editable.
                    if (_toolTip != null)
                        _toolTip.SetToolTip(btn, button.Prompt);
                }
            }
            finally
            {
                flowQuickActions.ResumeLayout(false);
            }

            // New buttons start with the Designer's light-mode look, so they need the current
            // palette before they are ever painted.
            ApplyThemeToButtons(grpQuickActions);
            LayoutPane();
            if (_isProcessing)
                SetUIEnabled(false);
        }

        /// <summary>
        /// Sizes the quick-action group to the buttons it holds and slides everything below it
        /// down to match. Nothing here is a literal position: the group's height falls out of how
        /// many rows the buttons wrapped onto, and grpInstruction and lblStatus are placed from
        /// that measured bottom. The old fixed geometry clipped a seventh button and overlapped
        /// anything taller than the two rows it was drawn for.
        ///
        /// For the six shipped buttons this reproduces the previous layout exactly (group 122 tall,
        /// instruction at y=167, status at y=342), which is the check that the constants above are
        /// the same ones the literals encoded.
        /// </summary>
        private void LayoutPane()
        {
            if (_layingOut)
            {
                // Re-entered from inside a pass. Do not lay out on top of a half-finished pass;
                // note that the width moved and let the driver below run another one.
                _layoutDirty = true;
                return;
            }

            _layingOut = true;
            try
            {
                // A pass can invalidate its own input: making the group taller can push the pane
                // past its height, and the scrollbar that then appears takes 17px off the client
                // width, which narrows this anchored group and changes how many buttons fit a row.
                // That arrives re-entrantly, so it is recorded rather than acted on, and answered
                // by running again. Capped, because a scrollbar that is needed at one width and
                // not at the other would otherwise alternate forever.
                for (int pass = 0; pass < 4; pass++)
                {
                    _layoutDirty = false;
                    LayoutQuickActions();
                    if (!_layoutDirty)
                        break;
                }
            }
            finally
            {
                _layingOut = false;
                _layoutDirty = false;
            }
        }

        private void LayoutQuickActions()
        {
            float scale = UiScale;
            int pad = Scaled(PaneMargin, scale);
            int gap = Scaled(ControlGap, scale);
            int minWidth = Scaled(MinQuickButtonWidth, scale);

            // Recorded so grpQuickActions_SizeChanged can tell this group being resized BY this
            // method - its height, which changes nothing here - from it being resized by anything
            // else, which means its width, which changes everything.
            _layoutWidth = grpQuickActions.ClientSize.Width;

            int inner = _layoutWidth - 2 * pad;
            if (inner < minWidth)
                inner = minWidth;

            // The panel is one gap wider than the content it holds. Every button carries a right
            // margin to space it from its neighbour, and FlowLayoutPanel counts that margin when
            // deciding where to wrap - so a panel exactly `inner` wide would fit one button fewer
            // per row than there is room for. The extra width is margin, so nothing draws in it.
            int panelWidth = inner + gap;
            flowQuickActions.SuspendLayout();
            try
            {
                SizeQuickButtons(inner, gap, minWidth, scale);
                flowQuickActions.Location = new Point(pad, Scaled(GroupTopInset, scale));
                // Asking the layout engine for the height at a known width rather than letting
                // AutoSize find it: this answers now, with no dependency on when the next layout
                // pass happens to run.
                flowQuickActions.Size = new Size(
                    panelWidth,
                    flowQuickActions.GetPreferredSize(new Size(panelWidth, 0)).Height);
            }
            finally
            {
                flowQuickActions.ResumeLayout(true);
            }

            // No gap added here: the bottom margin of the last button row is already in the panel's
            // height, and with no buttons at all there is nothing to space away from.
            btnSelectSignature.Location = new Point(pad, flowQuickActions.Bottom);
            btnSelectSignature.Width = inner;
            grpQuickActions.Height = btnSelectSignature.Bottom + Scaled(GroupBottomInset, scale);

            grpInstruction.Location = new Point(pad, grpQuickActions.Bottom + gap);
            lblStatus.Location = new Point(pad, grpInstruction.Bottom + gap);
        }

        /// <summary>
        /// Widths and one shared height for the current buttons. Widths snap to a column grid so
        /// rows pack evenly: a short caption gets one column (the same 70px the six shipped
        /// buttons always had), and a caption too long for that spans two or three rather than
        /// being cut off. The height is the tallest caption's wrapped height, applied to all of
        /// them, so no row comes out ragged.
        /// </summary>
        private void SizeQuickButtons(int inner, int gap, int minWidth, float scale)
        {
            int columns = (inner + gap) / (minWidth + gap);
            if (columns < 1)
                columns = 1;
            int columnWidth = (inner - (columns - 1) * gap) / columns;
            if (columnWidth < 1)
                columnWidth = 1;

            int chrome = Scaled(QuickButtonPadding, scale);
            var margin = new Padding(0, 0, gap, gap);
            int height = Scaled(MinQuickButtonHeight, scale);

            foreach (Control ctrl in flowQuickActions.Controls)
            {
                var btn = ctrl as Button;
                if (btn == null)
                    continue;

                Size needed = MeasureCaption(btn, Math.Max(1, inner - chrome));
                int columnsUsed = 1;
                while (columnsUsed < columns &&
                       needed.Width > SpanWidth(columnsUsed, columnWidth, gap) - chrome)
                    columnsUsed++;

                btn.Margin = margin;
                // The span is the ceiling, not the width. A caption that overflows one column by a
                // few pixels - which is what a display-scaled font does to a caption sized for 96
                // DPI - gets those pixels rather than a whole extra column it would not fill.
                btn.Width = Math.Min(
                    SpanWidth(columnsUsed, columnWidth, gap),
                    Math.Max(columnWidth, needed.Width + chrome));

                // Re-measure at the width it actually got: a caption that wrapped onto a second
                // line there needs a taller button, and a 64-character name (the store's cap)
                // does exactly that.
                int wrapped = MeasureCaption(btn, Math.Max(1, btn.Width - chrome)).Height + chrome;
                if (wrapped > height)
                    height = wrapped;
            }

            foreach (Control ctrl in flowQuickActions.Controls)
                ctrl.Height = height;
        }

        private Size MeasureCaption(Button btn, int availableWidth)
        {
            return TextRenderer.MeasureText(
                btn.Text ?? string.Empty,
                btn.Font ?? Font,
                new Size(availableWidth, int.MaxValue),
                QuickButtonMeasureFlags);
        }

        private static int SpanWidth(int columns, int columnWidth, int gap)
        {
            return columns * columnWidth + (columns - 1) * gap;
        }

        private static int Scaled(int value, float scale)
        {
            return (int)Math.Round(value * scale);
        }

        /// <summary>
        /// The factor AutoScaleMode.Font applied to every literal in the Designer, read back off
        /// one of them. Values computed at runtime never went through that pass, so they have to be
        /// scaled by hand or a 125% display keeps 96 DPI paddings around text that grew by a
        /// quarter.
        ///
        /// It is read back rather than recomputed because it cannot be recomputed afterwards:
        /// ContainerControl sets AutoScaleDimensions to the dimensions it just scaled TO, so the
        /// ratio it reports once the scaling has happened is always 1. The probe is
        /// grpQuickActions.Top because that is a Designer literal nothing else ever assigns -
        /// LayoutPane moves everything BELOW that group and never the group itself.
        ///
        /// AutoScrollPosition has to come back out of it. This pane scrolls, and a scrolled
        /// ScrollableControl reports its children's positions relative to the scrolled viewport, so
        /// a pane the user had scrolled down would otherwise report a smaller scale than the one
        /// actually in force - which is exactly what it did: 20 long-named buttons scrolled the pane
        /// far enough to make a 150% display measure as 100% and lay itself out with 96 DPI gaps.
        /// Floored at 1; shrinking below the shipped geometry is not a case worth having.
        /// </summary>
        private float UiScale
        {
            get
            {
                float top = grpQuickActions.Top - AutoScrollPosition.Y;
                float scale = top / QuickActionsDesignTop;
                return scale < 1f ? 1f : scale;
            }
        }

        private void grpQuickActions_SizeChanged(object sender, EventArgs e)
        {
            // Only a width change is news. LayoutPane sets this group's height itself, and
            // answering that with another layout would mean every pass triggering the next.
            if (grpQuickActions.ClientSize.Width == _layoutWidth)
                return;
            LayoutPane();
        }

        private void AITaskPane_FontChanged(object sender, EventArgs e)
        {
            // The pane's font changing IS the display-scaling event: AutoScaleMode.Font rescales
            // the Designer's literals off the back of it, and UiScale has to be read again for the
            // computed half to keep up.
            LayoutPane();
        }

        /// <summary>
        /// The user changed their buttons or prompts somewhere else - the settings dialog, opened
        /// from the Explorer ribbon - so this pane is stale. Every open compose window has its own
        /// pane and they all get this, because PromptStore.Changed is static and process-wide.
        ///
        /// Three things about that event: it is static, so <see cref="DisposeCustomResources"/>
        /// MUST unhook it or every compose window ever opened stays rooted for the life of the
        /// process; <paramref name="sender"/> is always null and says nothing; and it can arrive
        /// on whatever thread did the save, so the rebuild is marshalled the same fire-and-forget
        /// way <see cref="OnThemeChanged"/> marshals a theme switch.
        /// </summary>
        private void OnPromptsChanged(object sender, EventArgs e)
        {
            if (_disposed || this.IsDisposed || !this.IsHandleCreated)
                return;
            try { this.BeginInvoke((Action)RebuildQuickActions); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        // === Button click handlers ===

        /// <summary>
        /// Every quick button, whatever the user called it, arrives here. The instruction rides
        /// on the button's <see cref="Control.Tag"/> rather than being looked up by caption: the
        /// caption is the user's to change, and a lookup would go stale the moment the store is
        /// edited between the pane being built and the click landing.
        /// </summary>
        private async void QuickAction_Click(object sender, EventArgs e)
        {
            try
            {
                var btn = sender as Button;
                await ProcessAction(btn == null ? null : btn.Tag as string);
            }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnDraft_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtPrompt.Text))
                {
                    ShowStatus("Please enter instructions for the email you want to draft.", true);
                    return;
                }
                _editHistory.Clear();
                _freshDraft = true;
                await ProcessAction(ClaudeService.BuildDraftLabel(txtPrompt.Text));
            }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnEditDraft_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtPrompt.Text))
                {
                    ShowStatus("Please enter instructions for editing the draft.", true);
                    return;
                }
                await ProcessAction(ClaudeService.BuildDraftLabel(txtPrompt.Text));
            }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnEditSelection_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtPrompt.Text))
                {
                    ShowStatus("Please enter instructions for editing the selection.", true);
                    return;
                }
                string selectedText = GetSelectedText();
                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    ShowStatus("Please select text in the email editor first.", true);
                    return;
                }
                await ProcessAction(ClaudeService.BuildDraftLabel(txtPrompt.Text), selectedText);
            }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnSelectSignature_Click(object sender, EventArgs e)
        {
            try { await ProcessSelectSignature(); }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        // === D38: Select the best signature ===

        /// <summary>
        /// Gathers the draft context (recipients + draft body + quoted thread), lets
        /// the AI pick the best installed signature (skipped when only one exists),
        /// and applies it through the SAME _MailAutoSig bookmark machinery that
        /// anchors every other pane action - draft text and quoted thread are never
        /// touched, only the signature region is replaced (or inserted above the
        /// quote / at the end when none exists yet).
        /// </summary>
        private async Task ProcessSelectSignature()
        {
            if (_isProcessing) return;
            _isProcessing = true;
            object preAsyncDoc = null;
            try
            {
                SetUIEnabled(false);

                var signatures = SignatureStore.ListSignatures();
                if (signatures.Count == 0)
                {
                    RefreshSignatureAvailability();
                    ShowStatus("No signatures are installed in Outlook.", true);
                    return;
                }

                string draftText;
                string threadText;

                object doc = null;
                try
                {
                    doc = GetWordDocument();
                    if (doc == null)
                    {
                        ShowStatus("Could not access email editor.", true);
                        return;
                    }

                    dynamic wordDoc = doc;
                    var bmks = wordDoc.Bookmarks;
                    bmks.ShowHidden = true;
                    ThisAddIn.ReleaseCom((object)bmks);
                    DebugLog("SelectSignature BEFORE read", wordDoc);

                    draftText = ReadDraftText(wordDoc);
                    threadText = ReadThreadText(wordDoc);
                    preAsyncDoc = doc;
                    doc = null;
                }
                finally
                {
                    ThisAddIn.ReleaseCom(doc);
                }

                string recipientsText = GetRecipientSummary();

                SignatureStore.SignatureOption chosen;
                if (signatures.Count == 1)
                {
                    // One signature installed: the choice is trivial - no AI call.
                    chosen = signatures[0];
                }
                else
                {
                    ShowStatus("Selecting the best signature...", false);
                    string answer = await ClaudeService.SelectSignatureAsync(
                        signatures, draftText, threadText, recipientsText);
                    chosen = SignatureStore.FindByName(signatures, answer);
                    if (chosen == null)
                    {
                        string shown = answer ?? "";
                        if (shown.Length > 60) shown = shown.Substring(0, 60) + "...";
                        InvokeOnUI(() => ShowStatus("AI answered \"" + shown + "\", which matches no installed signature.", true));
                        return;
                    }
                }

                var capturedDoc = preAsyncDoc;
                InvokeOnUI(() =>
                {
                    if (_disposed) return;
                    if (ApplySignatureToDocument(capturedDoc, chosen.FilePath))
                        ShowStatus("Applied signature \"" + chosen.Name + "\".", false);
                });
            }
            catch (Exception ex)
            {
                InvokeOnUI(() =>
                {
                    string msg = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    ShowStatus(msg, true);
                });
            }
            finally
            {
                ThisAddIn.ReleaseCom(preAsyncDoc);
                _isProcessing = false;
                if (!_disposed && !IsDisposed)
                    SetUIEnabled(true);
            }
        }

        /// <summary>
        /// The signature-region bookmark dance (ADDITIVE reuse of the pane's proven
        /// _MailAutoSig machinery - same pattern as Outlook's own Insert &gt; Signature
        /// switcher): replace the _MailAutoSig region when it exists, else insert
        /// above _MailOriginal (quoted thread), else at the end of the document; then
        /// recreate the _MailAutoSig bookmark over the inserted content so the
        /// draft/signature/thread split keeps working. The draft region [0, draftEnd)
        /// and the quoted thread are never modified.
        /// </summary>
        private bool ApplySignatureToDocument(object capturedDoc, string signatureFilePath)
        {
            object doc = capturedDoc;
            dynamic bookmarks = null;
            try
            {
                if (doc == null)
                {
                    ShowStatus("Could not access email editor.", true);
                    return false;
                }

                if (signatureFilePath == null || !System.IO.File.Exists(signatureFilePath))
                {
                    ShowStatus("The signature's file no longer exists.", true);
                    return false;
                }

                dynamic wordDoc = doc;
                bookmarks = wordDoc.Bookmarks;
                bookmarks.ShowHidden = true;
                DebugLog("ApplySignature BEFORE", wordDoc);

                int insertAt;
                if (bookmarks.Exists("_MailAutoSig"))
                {
                    // Replace: drop the marker, then the old signature content itself.
                    var bmk = bookmarks["_MailAutoSig"];
                    var range = bmk.Range;
                    insertAt = range.Start;
                    bmk.Delete();
                    range.Delete();
                    ThisAddIn.ReleaseCom(range);
                    ThisAddIn.ReleaseCom(bmk);
                }
                else if (bookmarks.Exists("_MailOriginal"))
                {
                    // No signature region yet: insert directly ABOVE the quoted thread.
                    var bmk = bookmarks["_MailOriginal"];
                    var range = bmk.Range;
                    insertAt = range.Start;
                    ThisAddIn.ReleaseCom(range);
                    ThisAddIn.ReleaseCom(bmk);
                }
                else
                {
                    // Plain new draft: end of document (before the final paragraph mark).
                    var content = wordDoc.Content;
                    insertAt = Math.Max(0, (int)content.End - 1);
                    ThisAddIn.ReleaseCom(content);
                }

                var contentBefore = wordDoc.Content;
                int endBefore = contentBefore.End;
                ThisAddIn.ReleaseCom(contentBefore);

                var insertRange = wordDoc.Range(insertAt, insertAt);
                insertRange.InsertFile(signatureFilePath, Type.Missing, false, false, false);
                ThisAddIn.ReleaseCom(insertRange);

                var contentAfter = wordDoc.Content;
                int endAfter = contentAfter.End;
                ThisAddIn.ReleaseCom(contentAfter);
                int newEnd = insertAt + Math.Max(0, endAfter - endBefore);

                // Recreate the marker over the inserted content so Outlook and the
                // pane's draft/signature/thread split keep working on this draft.
                var newRange = wordDoc.Range(insertAt, newEnd);
                bookmarks.Add("_MailAutoSig", newRange);
                ThisAddIn.ReleaseCom(newRange);

                DebugLog("ApplySignature AFTER", wordDoc);

                if (endAfter <= endBefore)
                {
                    ShowStatus("The signature file inserted no content.", true);
                    return false;
                }

                return true;
            }
            catch (COMException ex)
            {
                Debug.WriteLine("ApplySignatureToDocument COM error: " + ex.Message);
                ShowStatus("Could not apply signature: " + ex.Message, true);
                return false;
            }
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine("ApplySignatureToDocument error: " + ex.Message);
                ShowStatus("Could not apply signature: " + ex.Message, true);
                return false;
            }
            finally
            {
                ThisAddIn.ReleaseCom((object)bookmarks);
            }
        }

        /// <summary>The compose item behind the pane (Inspector item or the inline response).</summary>
        private object GetCurrentMailItem()
        {
            try
            {
                if (_disposed) return null;
                if (!_isInlineResponse)
                    return _owningInspector?.CurrentItem;

                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null) return null;
                try
                {
                    return ((dynamic)explorer).ActiveInlineResponse;
                }
                finally
                {
                    ThisAddIn.ReleaseCom(explorer);
                }
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
        }

        /// <summary>
        /// Compact "Name &lt;address&gt;" summary of the draft's recipients (capped) as
        /// context for the signature choice. Empty when unavailable - the choice then
        /// rests on draft/thread language alone.
        /// </summary>
        private string GetRecipientSummary()
        {
            object item = null;
            object recipients = null;
            try
            {
                item = GetCurrentMailItem();
                if (item == null) return "";

                recipients = ((dynamic)item).Recipients;
                int count = ((dynamic)recipients).Count;
                var sb = new StringBuilder();
                for (int i = 1; i <= count && i <= 20; i++)
                {
                    object recipient = null;
                    try
                    {
                        recipient = ((dynamic)recipients)[i];
                        string name = null;
                        string address = null;
                        try { name = ((dynamic)recipient).Name as string; } catch (COMException) { }
                        try { address = ((dynamic)recipient).Address as string; } catch (COMException) { }
                        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(address))
                            continue;

                        if (sb.Length > 0) sb.AppendLine();
                        sb.Append(name ?? "");
                        if (!string.IsNullOrWhiteSpace(address) && !string.Equals(address, name, StringComparison.OrdinalIgnoreCase))
                            sb.Append(" <").Append(address).Append(">");
                    }
                    finally
                    {
                        ThisAddIn.ReleaseCom(recipient);
                    }
                }

                return sb.ToString();
            }
            catch (COMException)
            {
                return "";
            }
            catch (InvalidCastException)
            {
                return "";
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return "";
            }
            finally
            {
                ThisAddIn.ReleaseCom(recipients);
                ThisAddIn.ReleaseCom(item);
            }
        }

        // === Core processing ===

        /// <summary>
        /// Runs one edit. <paramref name="actionLabel"/> is the resolved instruction, whatever it
        /// came from - a quick button's stored prompt, or what the user typed wrapped by
        /// <see cref="ClaudeService.BuildDraftLabel"/> - and the SAME string is what the edit
        /// history records, so a prompt edited later cannot rewrite what a past turn asked for.
        /// </summary>
        private async Task ProcessAction(string actionLabel, string selectedText = null)
        {
            if (_isProcessing) return;
            _isProcessing = true;
            object preAsyncDoc = null;
            // Consume the one-shot "fresh draft" signal exactly once per invocation, up
            // front, so an early return or read error can't leave it set and then silently
            // blank the draft on the next unrelated action.
            bool freshDraft = _freshDraft;
            _freshDraft = false;
            try
            {
                SetUIEnabled(false);
                ShowStatus("Processing...", false);
                string draftText;
                string signatureText;
                string threadText;

                object doc = null;
                try
                {
                    doc = GetWordDocument();
                    if (doc == null)
                    {
                        ShowStatus("Could not access email editor.", true);
                        return;
                    }

                    dynamic wordDoc = doc;
                    var bmks = wordDoc.Bookmarks;
                    bmks.ShowHidden = true;
                    ThisAddIn.ReleaseCom((object)bmks);
                    DebugLog($"ProcessAction({actionLabel}) BEFORE read", wordDoc);

                    // Read signature/thread fresh on every action. Caching them across
                    // the pane lifetime sent stale context to the AI when Outlook reused
                    // an inspector window for a different mail item.
                    signatureText = ReadSignatureText(wordDoc);
                    threadText = ReadThreadText(wordDoc);

                    draftText = ReadDraftText(wordDoc);
                    preAsyncDoc = doc;
                    doc = null;
                }
                finally
                {
                    ThisAddIn.ReleaseCom(doc);
                }

                if (freshDraft)
                {
                    draftText = "";
                }

                DebugLog("Sending to Claude", extra:
                    $"draftText ({draftText.Length} chars), sigText ({signatureText.Length} chars), threadText ({threadText.Length} chars)");

                string result = await ClaudeService.ProcessEmailAsync(
                    actionLabel, _editHistory,
                    draftText, signatureText, threadText, selectedText);

                DebugLog("Claude returned", extra: $"result ({result.Length} chars): {(result.Length > 300 ? result.Substring(0, 300) + "..." : result)}");

                var capturedDoc = preAsyncDoc;
                InvokeOnUI(() =>
                {
                    if (_disposed) return;
                    if (WriteDraftToDocument(result, capturedDoc))
                    {
                        _editHistory.Add(new EditTurn
                        {
                            Label = actionLabel,
                            SelectedText = selectedText,
                            Result = result
                        });

                        ShowStatus("Done!", false);
                    }
                });
            }
            catch (Exception ex)
            {
                InvokeOnUI(() =>
                {
                    string msg = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    ShowStatus(msg, true);
                });
            }
            finally
            {
                // Always release the captured doc and clear the processing/UI state,
                // even if the pane was disposed during the await (so the UI can never
                // latch permanently disabled with _isProcessing stuck true).
                ThisAddIn.ReleaseCom(preAsyncDoc);
                _isProcessing = false;
                if (!_disposed && !IsDisposed)
                    SetUIEnabled(true);
            }
        }

        // === Word Object Model helpers ===

        private object GetWordDocument()
        {
            try
            {
                if (_disposed) return null;
                if (!_isInlineResponse)
                    return _owningInspector?.WordEditor;

                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null) return null;
                try
                {
                    return ((dynamic)explorer).ActiveInlineResponseWordEditor;
                }
                finally
                {
                    ThisAddIn.ReleaseCom(explorer);
                }
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
        }

        private int FindDraftEnd(dynamic doc)
        {
            var bookmarks = doc.Bookmarks;
            try
            {
                if (bookmarks.Exists("_MailAutoSig"))
                {
                    var bmk = bookmarks["_MailAutoSig"];
                    var range = bmk.Range;
                    int pos = range.Start;
                    ThisAddIn.ReleaseCom(range);
                    ThisAddIn.ReleaseCom(bmk);
                    return pos;
                }
                if (bookmarks.Exists("_MailOriginal"))
                {
                    var bmk = bookmarks["_MailOriginal"];
                    var range = bmk.Range;
                    int pos = range.Start;
                    ThisAddIn.ReleaseCom(range);
                    ThisAddIn.ReleaseCom(bmk);
                    return pos;
                }
            }
            finally
            {
                ThisAddIn.ReleaseCom((object)bookmarks);
            }
            var content = doc.Content;
            int end = content.End;
            ThisAddIn.ReleaseCom(content);
            return end;
        }

        private string ReadDraftText(dynamic doc)
        {
            int draftEnd = FindDraftEnd(doc);

            var range = doc.Range(0, draftEnd);
            string text = range.Text ?? "";
            ThisAddIn.ReleaseCom(range);

            return text.Trim('\r', '\n');
        }

        private string ReadSignatureText(dynamic doc)
        {
            var bookmarks = doc.Bookmarks;
            try
            {
                if (!bookmarks.Exists("_MailAutoSig"))
                    return "";

                var bmk = bookmarks["_MailAutoSig"];
                var range = bmk.Range;
                string text = range.Text ?? "";
                ThisAddIn.ReleaseCom(range);
                ThisAddIn.ReleaseCom(bmk);

                return text.TrimEnd('\r', '\n');
            }
            finally
            {
                ThisAddIn.ReleaseCom((object)bookmarks);
            }
        }

        private string ReadThreadText(dynamic doc)
        {
            var bookmarks = doc.Bookmarks;
            try
            {
                if (!bookmarks.Exists("_MailOriginal"))
                    return "";

                var bmk = bookmarks["_MailOriginal"];
                var range = bmk.Range;
                string text = range.Text ?? "";
                ThisAddIn.ReleaseCom(range);
                ThisAddIn.ReleaseCom(bmk);

                return text.TrimEnd('\r', '\n');
            }
            finally
            {
                ThisAddIn.ReleaseCom((object)bookmarks);
            }
        }

        private bool WriteDraftToDocument(string newDraftText, object capturedDoc = null)
        {
            object doc = capturedDoc;
            bool ownDoc = false;
            dynamic bookmarks = null;
            try
            {
                if (doc == null)
                {
                    doc = GetWordDocument();
                    ownDoc = true;
                }
                if (doc == null)
                {
                    ShowStatus("Could not access email editor.", true);
                    return false;
                }

                dynamic wordDoc = doc;
                bookmarks = wordDoc.Bookmarks;
                bookmarks.ShowHidden = true;
                DebugLog("WriteDraft BEFORE", wordDoc);

                string boundaryBookmark = null;
                var contentRange = wordDoc.Content;
                int draftEnd = contentRange.End;
                ThisAddIn.ReleaseCom(contentRange);
                int origBmkEnd = -1;

                if (bookmarks.Exists("_MailAutoSig"))
                    boundaryBookmark = "_MailAutoSig";
                else if (bookmarks.Exists("_MailOriginal"))
                    boundaryBookmark = "_MailOriginal";

                if (boundaryBookmark != null)
                {
                    var bmk = bookmarks[boundaryBookmark];
                    var bmkRange = bmk.Range;
                    draftEnd = bmkRange.Start;
                    origBmkEnd = bmkRange.End;
                    ThisAddIn.ReleaseCom(bmkRange);
                    bmk.Delete();
                    ThisAddIn.ReleaseCom(bmk);
                }

                bool textReplaced = false;
                int newDraftEnd = draftEnd;
                try
                {
                    var range = wordDoc.Range(0, draftEnd);
                    range.Text = newDraftText + "\r\n";
                    newDraftEnd = range.End;
                    textReplaced = true;
                    ThisAddIn.ReleaseCom(range);

                    if (boundaryBookmark != null && origBmkEnd >= 0)
                    {
                        int newBmkEnd = origBmkEnd + (newDraftEnd - draftEnd);
                        if (newBmkEnd < newDraftEnd) newBmkEnd = newDraftEnd;
                        var restoreRange = wordDoc.Range(newDraftEnd, newBmkEnd);
                        bookmarks.Add(boundaryBookmark, restoreRange);
                        ThisAddIn.ReleaseCom(restoreRange);
                    }
                }
                catch
                {
                    // We deleted the boundary bookmark up front; if the write/recreate failed,
                    // restore it so the signature/thread marker (and the context it anchors)
                    // isn't silently lost. Re-throw so the normal error handling still runs.
                    if (boundaryBookmark != null && origBmkEnd >= 0 && !bookmarks.Exists(boundaryBookmark))
                    {
                        try
                        {
                            int rs = textReplaced ? newDraftEnd : draftEnd;
                            int re = textReplaced ? origBmkEnd + (newDraftEnd - draftEnd) : origBmkEnd;
                            if (re < rs) re = rs;
                            var restore = wordDoc.Range(rs, re);
                            bookmarks.Add(boundaryBookmark, restore);
                            ThisAddIn.ReleaseCom(restore);
                        }
                        catch { }
                    }
                    throw;
                }

                DebugLog("WriteDraft AFTER", wordDoc);
                return true;
            }
            catch (COMException ex)
            {
                Debug.WriteLine("WriteDraftToDocument COM error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine("WriteDraftToDocument error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
            finally
            {
                ThisAddIn.ReleaseCom((object)bookmarks);
                if (ownDoc) ThisAddIn.ReleaseCom(doc);
            }
        }

        // === Selection support ===

        private string GetSelectedText()
        {
            object doc = null;
            object app = null;
            object sel = null;
            try
            {
                doc = GetWordDocument();
                if (doc == null) return null;
                app = ((dynamic)doc).Application;
                sel = ((dynamic)app).Selection;
                string text = ((dynamic)sel).Text as string;
                if (!string.IsNullOrEmpty(text) && text.EndsWith("\r"))
                    text = text.Substring(0, text.Length - 1);
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
            finally
            {
                ThisAddIn.ReleaseCom(sel);
                ThisAddIn.ReleaseCom(app);
                ThisAddIn.ReleaseCom(doc);
            }
        }

        // === UI helpers ===

        /// <summary>
        /// Tips for the fixed controls. The quick buttons get theirs in
        /// <see cref="RebuildQuickActions"/> instead, straight from the prompt each one sends -
        /// the six that used to be listed here were a second copy of that wording, and a copy is
        /// exactly what cannot survive the prompts becoming editable.
        /// </summary>
        private void SetupTooltips()
        {
            _toolTip = new ToolTip();
            _toolTip.SetToolTip(btnDraft, "Draft a new email from scratch based on your instruction.\nClears any previous AI draft.");
            _toolTip.SetToolTip(btnEditDraft, "Edit the current draft based on your instruction.\nPreserves conversation history for iterative refinement.");
            _toolTip.SetToolTip(btnEditSelection, "Edit only the selected text based on your instruction.\nLeaves the rest of the draft unchanged.");
            _toolTip.SetToolTip(btnSelectSignature, "Let the AI pick the best of your installed signatures for this email\n(matching the language of the draft, thread, and recipients) and apply it.\nYour draft text and the quoted thread stay untouched.");
            _toolTip.SetToolTip(lnkCheckUpdates, "Look for a newer version now, instead of waiting for the next\nautomatic check. OutlookAI checks every 10 minutes on its own.");
        }

        private void InvokeOnUI(Action action)
        {
            if (_disposed || this.IsDisposed || !this.IsHandleCreated)
                return;

            try
            {
                if (this.InvokeRequired)
                    this.Invoke(action);
                else
                    action();
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void OnThemeChanged(object sender, EventArgs e)
        {
            // ThemeService may raise this on a non-UI (SystemEvents) thread. Fire-and-forget
            // marshal via BeginInvoke so a busy UI thread can't stall the SystemEvents thread.
            if (_disposed || this.IsDisposed || !this.IsHandleCreated)
                return;
            try { this.BeginInvoke((Action)ApplyTheme); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void ApplyTheme()
        {
            // Two-way: apply the current ThemeService palette (dark or light) so the pane
            // tracks a runtime theme switch. The Designer binds BackColor/lblTitle/lblVersion
            // to ThemeService once at construction, so they must be re-applied here too.
            this.BackColor = ThemeService.Background;
            this.ForeColor = ThemeService.Text;
            lblTitle.ForeColor = ThemeService.Accent;
            lblVersion.ForeColor = ThemeService.SecondaryText;
            lnkUpdateError.LinkColor = ThemeService.LinkError;
            lnkCheckUpdates.LinkColor = ThemeService.Accent;
            lnkCheckUpdates.DisabledLinkColor = ThemeService.SecondaryText;
            if (lblStatus.Visible)
                lblStatus.ForeColor = _lastStatusError ? ThemeService.StatusError : ThemeService.StatusSuccess;

            foreach (var grp in new[] { grpQuickActions, grpInstruction })
            {
                grp.ForeColor = ThemeService.Text;
            }

            foreach (var txt in new[] { txtPrompt })
            {
                txt.BackColor = ThemeService.TextBoxBackground;
                txt.ForeColor = ThemeService.Text;
            }

            foreach (Control ctrl in this.Controls)
                ApplyThemeToButtons(ctrl);
        }

        private void ApplyThemeToButtons(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Button btn)
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
                        // Restore the native visual-styled light appearance (Designer default).
                        btn.FlatStyle = FlatStyle.Standard;
                        btn.UseVisualStyleBackColor = true;
                        btn.ForeColor = ThemeService.ButtonText;
                    }
                }
                if (ctrl.HasChildren)
                    ApplyThemeToButtons(ctrl);
            }
        }

        private void ShowStatus(string message, bool isError)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = isError ? ThemeService.StatusError : ThemeService.StatusSuccess;
            _lastStatusError = isError;
            lblStatus.Visible = true;
        }

        /// <summary>
        /// Locks the pane down for the duration of a request. Found by walking the tree rather
        /// than by naming fields: the quick buttons are not fields any more, and a named list
        /// would leave every one of them clickable during a request - where the _isProcessing gate
        /// then drops the second click silently, with no status line to say why nothing happened.
        /// The update links are deliberately not in scope; a check for updates is unrelated to a
        /// draft being rewritten.
        /// </summary>
        private void SetUIEnabled(bool enabled)
        {
            SetInputEnabled(this, enabled);
            btnSelectSignature.Enabled = enabled && _signaturesAvailable;
        }

        private void SetInputEnabled(Control parent, bool enabled)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Button || ctrl is TextBox)
                    ctrl.Enabled = enabled;
                if (ctrl.HasChildren)
                    SetInputEnabled(ctrl, enabled);
            }
        }

        partial void DisposeCustomResources()
        {
            _disposed = true;
            ThemeService.ThemeChanged -= OnThemeChanged;
            // Static event, so this is not optional: left hooked, it roots this pane - and the
            // compose window behind it - for the life of the Outlook process, once per window
            // ever opened.
            PromptStore.Changed -= OnPromptsChanged;
            grpQuickActions.SizeChanged -= grpQuickActions_SizeChanged;
            this.FontChanged -= AITaskPane_FontChanged;
            _versionTimer?.Stop();
            _versionTimer?.Dispose();
            _toolTip?.Dispose();
            // Shared by every quick button, so it outlives each individual button and is the
            // pane's to dispose. The buttons themselves go with the control tree.
            _quickButtonFont?.Dispose();
            _quickButtonFont = null;
            var inspector = _owningInspector;
            _owningInspector = null;
            ThisAddIn.ReleaseCom(inspector);
        }
    }
}
