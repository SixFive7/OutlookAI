using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI.TaskPane
{
    /// <summary>
    /// The Prompts and Buttons half of <see cref="SettingsDialog"/>: the editing surface over
    /// <see cref="PromptStore"/>. Two tabs - the four prompt sections every request is assembled
    /// from, and the quick buttons the compose sidebar shows, in order.
    ///
    /// IT BUFFERS, and that is the one thing that makes it different from the rest of the window.
    /// The tuning tick boxes write on every click because a tick box IS the decision. Text being
    /// typed is not: a half-typed prompt is not an instruction anybody meant to send, so
    /// everything here is a draft (<see cref="ButtonDraft"/> and <see cref="_sectionDraft"/>),
    /// compared against the baseline that was loaded, and nothing reaches the registry until
    /// "Apply now". That is also why no destructive-looking action asks for confirmation:
    /// deleting a built-in button or restoring the shipped set costs nothing until it is
    /// committed, and closing without applying is the undo. The one confirmation is closing with
    /// unsaved changes.
    ///
    /// One thing it deliberately does NOT do: subscribe to <see cref="PromptStore.Changed"/>.
    /// That event is static, so a subscription that outlives the window roots it for the life of
    /// the process, and this window is the only thing that writes prompts anyway. The compose
    /// panes are the subscribers; saving through the store is how they hear about it, and it is
    /// the only channel - several panes can be live at once, in other windows, and this window
    /// never touches one.
    /// </summary>
    public partial class SettingsDialog
    {
        // ===== The sections, in the order they are shown =====

        private static readonly PromptSection[] Sections =
        {
            PromptSection.Preamble,
            PromptSection.ReplyRules,
            PromptSection.SignatureRule,
            PromptSection.SignatureSelection,
        };

        /// <summary>
        /// How many sections there are - asked of <see cref="Sections"/>, never stated. Two
        /// constants encoding one fact is a trap with two ends: a fifth section added to the
        /// array without bumping the count vanishes from the window, and a bumped count without
        /// the array entry throws IndexOutOfRangeException the moment the tab loads. Declared
        /// AFTER the array on purpose - static initialisers run in textual order, so the other
        /// way round this reads 0.
        /// </summary>
        private static readonly int SectionCount = Sections.Length;

        // ===== The button-detail grid: one fact, three statements =====

        /// <summary>
        /// Rows in the button-detail grid: label, name box, label, prompt editor, state line,
        /// reset button. Stated once so the grid's <c>RowCount</c> and the loop that gives every
        /// row a style cannot disagree about how many there are.
        /// </summary>
        private const int ButtonDetailRowCount = 6;

        /// <summary>
        /// WHICH ROW OF THAT GRID IS THE PROMPT EDITOR - and it is one constant because three
        /// statements have to agree about it: the row that gets <c>Percent(100F)</c> (the one
        /// row that grows with the window), the cell <c>_txtPrompt</c> is added to, and the
        /// assertion below that the two ended up being the same row.
        ///
        /// <para>
        /// They used to be three hand-written 3s, and the failure that costs was SILENT.
        /// Insert a row above the editor, update the count and the add, forget the row style,
        /// and the dialog still opens, still lays out and still works - the editor simply stops
        /// filling and becomes a small box that scrolls, while some label above it takes all the
        /// height instead. Nothing throws, nothing logs, and no test sees it, because this is
        /// layout and layout is not tested here.
        /// </para>
        /// </summary>
        private const int PromptEditorRow = 3;

        /// <summary>Prompt for a button created by Add. Non-empty, because an empty one is rejected.</summary>
        private const string NewButtonPrompt =
            "Rewrite the draft. Keep the meaning and the tone unchanged.";

        private const string NewButtonName = "New button";

        // ===== Working state: drafts, and the baseline they are compared against =====

        private readonly List<ButtonDraft> _buttons = new List<ButtonDraft>();
        private readonly List<ButtonDraft> _baseline = new List<ButtonDraft>();
        private readonly string[] _sectionDraft = new string[SectionCount];
        private readonly string[] _sectionBaseline = new string[SectionCount];

        // ===== Controls =====

        private TableLayoutPanel _listSide;
        private ListBox _lstButtons;
        private Button _btnAdd;
        private Button _btnRemove;
        private Button _btnMoveUp;
        private Button _btnMoveDown;
        private Button _btnRestoreButtons;
        private TextBox _txtName;
        private TextBox _txtPrompt;
        private Label _lblButtonState;
        private Button _btnResetButton;

        private readonly TextBox[] _sectionEditors = new TextBox[SectionCount];
        private readonly Label[] _sectionLint = new Label[SectionCount];
        private readonly Label[] _sectionState = new Label[SectionCount];
        private readonly Button[] _btnResetSection = new Button[SectionCount];
        private readonly RowStyle[] _sectionEditorRow = new RowStyle[SectionCount];

        /// <summary>The footer line. Says what the BUFFERED half of the window is waiting on.</summary>
        private Label _lblStatus;

        /// <summary>
        /// What the status line is currently saying, so a theme switch can repaint it in the
        /// right colour without matching on its text.
        /// </summary>
        private StatusKind _status;

        /// <summary>Set while the code is writing to controls, so change handlers stay quiet.</summary>
        private bool _loading;

        // ===== Construction =====

        /// <summary>
        /// Creates the leaf controls the two pages and the footer are built from. Called before
        /// the pages themselves, because the footer's status label is assembled with the shell.
        /// </summary>
        private void BuildPromptControls()
        {
            _lstButtons = new ListBox
            {
                Name = "lstButtons",
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                SelectionMode = SelectionMode.One,
            };

            _btnAdd = NewButton("Add");
            _btnRemove = NewButton("Remove");
            _btnMoveUp = NewButton("Move up");
            _btnMoveDown = NewButton("Move down");
            _btnRestoreButtons = NewButton("Restore default buttons");

            _txtName = new TextBox
            {
                Name = "txtName",
                Dock = DockStyle.Top,
                MaxLength = PromptDefaults.MaxButtonNameLength,
            };

            _txtPrompt = NewEditor("txtPrompt");
            _lblButtonState = NewLabel("", LabelRole.Secondary, wrap: true);
            _btnResetButton = NewButton("Reset to default");

            _lblStatus = new Label
            {
                Name = "lblStatus",
                Dock = DockStyle.Fill,
                AutoSize = false,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
            };
        }

        private void WirePromptEvents()
        {
            _lstButtons.SelectedIndexChanged += OnButtonSelected;
            _txtName.TextChanged += OnNameEdited;
            _txtName.KeyDown += OnSingleLineKeyDown;
            _txtPrompt.TextChanged += OnPromptEdited;
            _btnAdd.Click += OnAddButton;
            _btnRemove.Click += OnRemoveButton;
            _btnMoveUp.Click += (s, e) => MoveSelected(-1);
            _btnMoveDown.Click += (s, e) => MoveSelected(1);
            _btnRestoreButtons.Click += OnRestoreDefaultButtons;
            _btnResetButton.Click += OnResetButtonPrompt;
        }

        private TableLayoutPanel BuildButtonsPage()
        {
            // One column at this level, which is what makes the wrapped help label below
            // measurable: ReflowWrappedLabels asks a label's PARENT how wide it is, and that
            // answer is only the label's own width in a single-column layout.
            var page = new TableLayoutPanel
            {
                Name = "buttonsPage",
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = Padding.Empty,
            };
            page.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            page.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            page.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label help = NewLabel(
                "These are the quick buttons in the compose sidebar, in this order. A button IS "
                + "its name: renaming one makes a different button, and renaming a built-in one "
                + "turns it into a custom button that no longer picks up improvements to the "
                + "text OutlookAI ships.",
                LabelRole.Secondary, wrap: true);

            var body = new TableLayoutPanel
            {
                Name = "buttonsBody",
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34F));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66F));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Left: the ordered list, with its actions under it.
            var listSide = new TableLayoutPanel
            {
                Name = "listSide",
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = Padding.Empty,
            };
            listSide.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            listSide.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            listSide.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            listSide.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _listSide = listSide;

            // A 2x2 grid rather than a wrapping flow: the four captions are wider than this
            // column, and a flow panel measured without a width constraint would lay them out
            // in one row and clip them. Grid cells share the column instead, so they get
            // narrower on a narrow window and AutoEllipsis takes care of the caption.
            var listActions = new TableLayoutPanel
            {
                Name = "listActions",
                // Dock=Top plus AutoSize, not Dock=Fill: a Fill grid is handed the whole row and
                // then hands the leftover to its LAST row, which made "Move up" and "Move down"
                // twice as tall as "Add" and "Remove". AutoSize means there is no leftover.
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 2,
                Margin = Padding.Empty,
            };
            listActions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            listActions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            listActions.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            listActions.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            listActions.Controls.Add(_btnAdd, 0, 0);
            listActions.Controls.Add(_btnRemove, 1, 0);
            listActions.Controls.Add(_btnMoveUp, 0, 1);
            listActions.Controls.Add(_btnMoveDown, 1, 1);

            listSide.Controls.Add(_lstButtons, 0, 0);
            listSide.Controls.Add(listActions, 0, 1);
            listSide.Controls.Add(_btnRestoreButtons, 0, 2);

            // Right: name, prompt, and what this button currently is.
            var detail = new TableLayoutPanel
            {
                Name = "buttonDetail",
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = ButtonDetailRowCount,
                Margin = Padding.Empty,
            };
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            for (int i = 0; i < ButtonDetailRowCount; i++)
                detail.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // The one row that grows with the window. Everything else is AutoSize.
            detail.RowStyles[PromptEditorRow] = new RowStyle(SizeType.Percent, 100F);

            detail.Controls.Add(NewLabel("Button name", LabelRole.Body, wrap: false), 0, 0);
            detail.Controls.Add(_txtName, 0, 1);
            detail.Controls.Add(NewLabel("Prompt sent to the model", LabelRole.Body, wrap: false), 0, 2);
            detail.Controls.Add(_txtPrompt, 0, PromptEditorRow);
            detail.Controls.Add(_lblButtonState, 0, 4);
            detail.Controls.Add(_btnResetButton, 0, 5);

            // ASKS THE GRID WHERE THE EDITOR ACTUALLY LANDED, which is not the same question as
            // where it was put. TableLayoutPanel silently relocates a control whose cell is
            // already taken, so inserting a row above the editor without moving PromptEditorRow
            // bumps it to the next free cell - and the 100% style stays on the row it was told
            // about, not the row the editor is in. That is the exact silent break the constant
            // above cannot catch on its own.
            //
            // Debug.Assert, so it is compiled out of the shipped Release build: a throw here
            // would trade a cosmetic layout fault for a settings dialog that will not open, and
            // the reader who needs telling is the one who just inserted the row and is running
            // the dialog to see what it looks like.
            Debug.Assert(detail.GetRow(_txtPrompt) == PromptEditorRow,
                "The prompt editor is not in PromptEditorRow, so the row set to fill the window "
                + "is not the editor's row. The editor will stop expanding.");

            body.Controls.Add(listSide, 0, 0);
            body.Controls.Add(detail, 1, 0);

            page.Controls.Add(help, 0, 0);
            page.Controls.Add(body, 0, 1);
            return page;
        }

        private Panel BuildSectionsPage()
        {
            // AutoScroll rather than "make everything fit": four multi-line editors do not fit
            // a laptop screen and shrinking them to make them fit is the opposite of useful.
            var page = new Panel
            {
                Name = "sectionsPage",
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Margin = Padding.Empty,
            };

            var stack = new TableLayoutPanel
            {
                Name = "sectionsStack",
                // Dock=Top for the width, AutoSize for the height: the stack is exactly as tall
                // as its groups need and exactly as wide as the viewport, so a vertical scroll
                // bar appearing cannot summon a horizontal one.
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = SectionCount + 1,
                Margin = Padding.Empty,
            };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            for (int i = 0; i <= SectionCount; i++)
                stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            stack.Controls.Add(NewLabel(
                "Every request is assembled from these blocks. Only your changes are stored, so "
                + "a block you leave alone keeps tracking the text OutlookAI ships. You may edit "
                + "all of it, including the safety wording - the warnings below an editor are "
                + "advice, and they never stop a save.",
                LabelRole.Secondary, wrap: true), 0, 0);

            for (int i = 0; i < SectionCount; i++)
                stack.Controls.Add(BuildSectionGroup(i), 0, i + 1);

            page.Controls.Add(stack);
            return page;
        }

        private GroupBox BuildSectionGroup(int index)
        {
            PromptSection section = Sections[index];

            var group = new GroupBox
            {
                Name = "grpSection" + section,
                Text = SectionTitle(section),
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
            };

            var inner = new TableLayoutPanel
            {
                Name = "section" + section,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 4,
                Margin = Padding.Empty,
            };
            inner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // The editor is the only row with a height of its own; everything else grows from
            // its text, which is what keeps this readable at 125% and 150%.
            _sectionEditorRow[index] = new RowStyle(SizeType.Absolute, EditorHeight(SectionLines(section)));
            inner.RowStyles.Add(_sectionEditorRow[index]);
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inner.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _sectionEditors[index] = NewEditor("txtSection" + section);
            _sectionLint[index] = NewLabel("", LabelRole.Warning, wrap: true);
            _sectionLint[index].Visible = false;
            _sectionState[index] = NewLabel("", LabelRole.Secondary, wrap: false);
            _sectionState[index].TextAlign = ContentAlignment.MiddleLeft;

            // Held in the array as well as in the layout, because ApplyMetrics has to size it -
            // a button created here and then forgotten keeps the WinForms default 75x23 and
            // shows "Resto..." at every display scale. Same shape of mistake as a control that
            // never gets themed, so it gets the same answer: register it where it is made.
            Button reset = NewButton("Restore default");
            _btnResetSection[index] = reset;
            int captured = index;
            reset.Click += (s, e) => RestoreSectionDefault(captured);
            _sectionEditors[index].TextChanged += (s, e) => OnSectionEdited(captured);

            var actions = new FlowLayoutPanel
            {
                Name = "sectionActions" + section,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty,
            };
            actions.Controls.Add(reset);
            actions.Controls.Add(_sectionState[index]);

            inner.Controls.Add(NewLabel(SectionPurpose(section), LabelRole.Secondary, wrap: true), 0, 0);
            inner.Controls.Add(_sectionEditors[index], 0, 1);
            inner.Controls.Add(_sectionLint[index], 0, 2);
            inner.Controls.Add(actions, 0, 3);

            group.Controls.Add(inner);
            return group;
        }

        private TextBox NewEditor(string name)
        {
            var editor = new TextBox
            {
                Name = name,
                Dock = DockStyle.Fill,
                Multiline = true,
                // The whole reason AcceptButton has to stay null on this form.
                AcceptsReturn = true,
                // False, so Tab still moves to the next control rather than typing a tab that
                // the store would reject in a name and that means nothing in a prompt.
                AcceptsTab = false,
                WordWrap = true,
                ScrollBars = ScrollBars.Vertical,
            };
            editor.KeyDown += OnEditorKeyDown;
            return editor;
        }

        /// <summary>
        /// Ctrl+A in a multi-line text box. WinForms does not wire it up itself, and selecting a
        /// whole prompt is the first thing anybody does before replacing one.
        /// </summary>
        private static void OnEditorKeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.A)
                return;
            var editor = sender as TextBox;
            if (editor == null)
                return;
            editor.SelectAll();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        // ===== Metrics =====

        /// <summary>Height for an editor showing <paramref name="lines"/> lines of text.</summary>
        private int EditorHeight(int lines)
        {
            int line = Font == null ? UiScale.DesignLineHeight : Math.Max(1, Font.Height);
            return lines * (line + 1) + Scaled(8);
        }

        /// <summary>
        /// The sizes the two prompt pages need, from <see cref="SettingsDialog.ApplyMetrics"/> so
        /// there is one place that decides what a row is worth at this display scale.
        /// </summary>
        private void ApplyPromptMetrics(int pad, int gap, int rowHeight)
        {
            // The four list actions and the restore button all fill their cell, so their
            // width comes from the column and only the height is set here.
            foreach (Button b in new[] { _btnAdd, _btnRemove, _btnMoveUp, _btnMoveDown })
            {
                b.Dock = DockStyle.Fill;
                b.MinimumSize = new Size(0, rowHeight);
                b.Margin = new Padding(0, gap, gap, 0);
            }
            _btnRemove.Margin = new Padding(0, gap, 0, 0);
            _btnMoveDown.Margin = new Padding(0, gap, 0, 0);

            _btnRestoreButtons.Dock = DockStyle.Fill;
            _btnRestoreButtons.MinimumSize = new Size(0, rowHeight);
            _btnRestoreButtons.Margin = new Padding(0, gap, 0, 0);

            _btnResetButton.Size = new Size(Scaled(150), rowHeight);
            _btnResetButton.Margin = new Padding(0, gap, 0, 0);

            _listSide.Margin = new Padding(0, 0, pad, 0);
            _lblStatus.MinimumSize = new Size(0, rowHeight);

            // Both prompt pages sit directly on the tab surface, so they carry their own inset -
            // the settings pages get theirs from the stack inside their scroller.
            Control buttonsPage = FindByName(this, "buttonsPage");
            if (buttonsPage != null)
                buttonsPage.Padding = new Padding(pad, gap, pad, gap);
            Control sectionsStack = FindByName(this, "sectionsStack");
            if (sectionsStack != null)
                sectionsStack.Padding = new Padding(pad, gap, pad, gap);

            for (int i = 0; i < SectionCount; i++)
            {
                if (_sectionEditorRow[i] != null)
                    _sectionEditorRow[i].Height = EditorHeight(SectionLines(Sections[i]));
                if (_sectionEditors[i] != null)
                    _sectionEditors[i].Margin = new Padding(0, gap, 0, gap);
                if (_btnResetSection[i] != null)
                {
                    _btnResetSection[i].Size = new Size(Scaled(150), rowHeight);
                    _btnResetSection[i].Margin = new Padding(0, 0, gap, 0);
                }
            }
        }

        // ===== Loading and saving =====

        /// <summary>
        /// Fills the drafts from storage and takes a baseline copy. The baseline is what "has
        /// anything changed" is answered against, so the footer knows whether closing would lose
        /// anything.
        /// </summary>
        private void LoadFromStore(int select)
        {
            _loading = true;
            try
            {
                _buttons.Clear();
                _baseline.Clear();
                foreach (PromptButton button in PromptStore.GetButtons())
                {
                    _buttons.Add(new ButtonDraft(button.Name, button.Prompt));
                    _baseline.Add(new ButtonDraft(button.Name, button.Prompt));
                }

                for (int i = 0; i < SectionCount; i++)
                {
                    string text = ToEditorText(PromptStore.GetSection(Sections[i]));
                    _sectionDraft[i] = text;
                    _sectionBaseline[i] = text;
                    _sectionEditors[i].Text = text;
                }
            }
            catch (Exception ex)
            {
                // A read that failed leaves whatever was already loaded on screen rather than
                // an empty form claiming the user has no buttons.
                Debug.WriteLine("Prompt load: " + ex.Message);
            }
            finally
            {
                _loading = false;
            }

            RebuildButtonList(select);
            RefreshDetailEditors();
            RefreshDetailState();
            for (int i = 0; i < SectionCount; i++)
                RefreshSectionState(i);
            RefreshCommitState();
        }

        /// <summary>
        /// Writes the drafts. Buttons are validated FIRST and nothing at all is written if that
        /// fails, so a rejected name cannot leave half the form applied. Returns whether
        /// everything asked for landed.
        /// </summary>
        private bool Commit()
        {
            var proposed = new List<PromptButton>(_buttons.Count);
            foreach (ButtonDraft draft in _buttons)
                proposed.Add(new PromptButton(draft.Name, draft.Prompt));

            PromptValidationResult validation = PromptStore.ValidateButtons(proposed);
            if (!validation.Succeeded)
            {
                // Validation only ever rejects buttons, so put the user in front of them.
                ShowPage(_tabButtons);
                SetStatus("Nothing was saved.", StatusKind.Failed);
                MessageBox.Show(this, validation.Message, "OutlookAI",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            var problems = new List<string>();

            for (int i = 0; i < SectionCount; i++)
            {
                // Only what changed: every write raises PromptStore.Changed, and a pane
                // rebuilding four times over for three blocks nobody touched is waste.
                if (PromptDefaults.SameText(_sectionDraft[i], PromptStore.GetSection(Sections[i])))
                    continue;
                if (!PromptStore.SetSection(Sections[i], _sectionDraft[i]))
                    problems.Add("\"" + SectionTitle(Sections[i]) + "\" could not be saved.");
            }

            if (ButtonsChanged())
            {
                if (MatchesShippedButtons())
                {
                    // Back to exactly what OutlookAI ships. Say that to the store rather than
                    // writing the six names out: the stored order is authoritative when it is
                    // present, so writing it would pin today's six and a seventh shipped later
                    // would never appear.
                    if (!PromptStore.RestoreDefaultButtons())
                        problems.Add("The buttons could not be restored to their defaults.");
                }
                else
                {
                    PromptValidationResult saved = PromptStore.SaveButtons(proposed);
                    if (!saved.Succeeded)
                        problems.Add(saved.Message);
                }
            }

            if (problems.Count > 0)
            {
                SetStatus("Some changes were not saved.", StatusKind.Failed);
                MessageBox.Show(this, string.Join(Environment.NewLine, problems.ToArray()),
                                "OutlookAI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            AdoptDraftAsBaseline();
            SetStatus("Saved.", StatusKind.Saved);
            RefreshCommitState();
            return true;
        }

        /// <summary>
        /// The drafts are now what storage holds, so they become the baseline. Deliberately not
        /// a re-read: re-assigning every editor's Text would throw away the caret and the scroll
        /// position of whatever the user was in the middle of, and there is nothing to learn -
        /// the store was just told exactly this.
        /// </summary>
        private void AdoptDraftAsBaseline()
        {
            _baseline.Clear();
            foreach (ButtonDraft draft in _buttons)
                _baseline.Add(new ButtonDraft(draft.Name, draft.Prompt));
            for (int i = 0; i < SectionCount; i++)
                _sectionBaseline[i] = _sectionDraft[i];
        }

        private bool ButtonsChanged()
        {
            if (_baseline.Count != _buttons.Count)
                return true;
            for (int i = 0; i < _buttons.Count; i++)
            {
                if (!string.Equals(_baseline[i].Name, _buttons[i].Name, StringComparison.Ordinal))
                    return true;
                if (!PromptDefaults.SameText(_baseline[i].Prompt, _buttons[i].Prompt))
                    return true;
            }
            return false;
        }

        private bool SectionsChanged()
        {
            for (int i = 0; i < SectionCount; i++)
            {
                if (!PromptDefaults.SameText(_sectionBaseline[i], _sectionDraft[i]))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Whether anything BUFFERED is waiting to be applied. The tuning tick boxes are never
        /// part of this: they wrote when they were clicked.
        /// </summary>
        private bool HasChanges()
        {
            return ButtonsChanged() || SectionsChanged();
        }

        /// <summary>
        /// Whether the drafted buttons are, exactly, the six OutlookAI ships: same names in the
        /// same order with the same casing, and the shipped prompt for each. Ordinal on purpose -
        /// a user who lowercased a name asked for that spelling, and handing them back the
        /// canonical casing would be a change they did not make.
        /// </summary>
        private bool MatchesShippedButtons()
        {
            IList<string> names = PromptDefaults.ButtonNames;
            if (_buttons.Count != names.Count)
                return false;

            for (int i = 0; i < names.Count; i++)
            {
                if (!string.Equals(_buttons[i].Name, names[i], StringComparison.Ordinal))
                    return false;
                string shipped;
                if (!PromptDefaults.TryGetButtonPrompt(names[i], out shipped))
                    return false;
                if (!PromptDefaults.SameText(_buttons[i].Prompt, shipped))
                    return false;
            }
            return true;
        }

        // ===== Buttons tab behaviour =====

        private void RebuildButtonList(int select)
        {
            _loading = true;
            try
            {
                _lstButtons.BeginUpdate();
                _lstButtons.Items.Clear();
                foreach (ButtonDraft draft in _buttons)
                    _lstButtons.Items.Add(DisplayName(draft));
                _lstButtons.EndUpdate();

                if (_buttons.Count > 0)
                {
                    int index = select;
                    if (index < 0)
                        index = 0;
                    if (index > _buttons.Count - 1)
                        index = _buttons.Count - 1;
                    _lstButtons.SelectedIndex = index;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Prompt list: " + ex.Message);
            }
            finally
            {
                _loading = false;
            }
        }

        /// <summary>
        /// What the list shows for a draft. An empty name is shown as a placeholder rather than
        /// as a blank row, because a blank row is a button the user cannot select to fix.
        /// </summary>
        private static string DisplayName(ButtonDraft draft)
        {
            string name = draft.Name == null ? "" : draft.Name.Trim();
            return name.Length == 0 ? "(no name yet)" : draft.Name;
        }

        private ButtonDraft Selected()
        {
            int index = _lstButtons.SelectedIndex;
            if (index < 0 || index >= _buttons.Count)
                return null;
            return _buttons[index];
        }

        private void OnButtonSelected(object sender, EventArgs e)
        {
            if (_loading)
                return;
            RefreshDetailEditors();
            RefreshDetailState();
        }

        /// <summary>Pushes the selected draft into the editors. Only on a selection change.</summary>
        private void RefreshDetailEditors()
        {
            _loading = true;
            try
            {
                ButtonDraft draft = Selected();
                _txtName.Text = draft == null ? "" : draft.Name;
                _txtPrompt.Text = draft == null ? "" : ToEditorText(draft.Prompt);
            }
            finally
            {
                _loading = false;
            }
        }

        /// <summary>
        /// Everything about the selection that is not its text: which actions are available, and
        /// what this button currently IS. Safe to call while the user is typing - it never writes
        /// to an editor.
        /// </summary>
        private void RefreshDetailState()
        {
            ButtonDraft draft = Selected();
            int index = _lstButtons.SelectedIndex;
            bool have = draft != null;

            _txtName.Enabled = have;
            _txtPrompt.Enabled = have;
            _btnRemove.Enabled = have;
            _btnMoveUp.Enabled = have && index > 0;
            _btnMoveDown.Enabled = have && index >= 0 && index < _buttons.Count - 1;

            if (!have)
            {
                _btnResetButton.Enabled = false;
                SetLabel(_lblButtonState, _buttons.Count == 0
                    ? "There are no buttons. Add one, or restore the defaults."
                    : "Select a button to edit its name and its prompt.");
                return;
            }

            bool builtIn = PromptDefaults.IsDefaultButtonName(draft.Name);
            string shipped;
            bool edited = !PromptDefaults.TryGetButtonPrompt(draft.Name, out shipped)
                          || !PromptDefaults.SameText(draft.Prompt, shipped);

            // Enabled only for a built-in NAME. Resetting a custom button has no shipped text to
            // fall back to - the store would drop its prompt and leave the name pointing at
            // nothing - so it is not offered at all.
            _btnResetButton.Enabled = builtIn;

            if (!builtIn)
            {
                SetLabel(_lblButtonState,
                    "Custom button. There is no shipped text behind it, so it cannot be reset - "
                    + "and it will not pick up improvements to OutlookAI's own prompts.");
            }
            else if (edited)
            {
                SetLabel(_lblButtonState,
                    "Built-in button with an edited prompt. Reset to default puts the shipped "
                    + "text back and lets it track future improvements again.");
            }
            else
            {
                SetLabel(_lblButtonState,
                    "Built-in button, unchanged. Its prompt is whatever OutlookAI ships, so it "
                    + "improves when OutlookAI does.");
            }
        }

        private void OnNameEdited(object sender, EventArgs e)
        {
            if (_loading)
                return;
            ButtonDraft draft = Selected();
            if (draft == null)
                return;

            int index = _lstButtons.SelectedIndex;
            draft.Name = _txtName.Text;

            _loading = true;
            try
            {
                // Replacing an item can drop the selection, so it is put back under the guard.
                _lstButtons.Items[index] = DisplayName(draft);
                _lstButtons.SelectedIndex = index;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Prompt rename: " + ex.Message);
            }
            finally
            {
                _loading = false;
            }

            // A rename can turn a built-in button into a custom one, which changes both the
            // advice and whether Reset is available.
            RefreshDetailState();
            RefreshCommitState();
        }

        private void OnPromptEdited(object sender, EventArgs e)
        {
            if (_loading)
                return;
            ButtonDraft draft = Selected();
            if (draft == null)
                return;

            draft.Prompt = _txtPrompt.Text;
            RefreshDetailState();
            RefreshCommitState();
        }

        private void OnSingleLineKeyDown(object sender, KeyEventArgs e)
        {
            // Enter in the name box has nowhere to go - there is no default button, by design -
            // and an unhandled Enter in a single-line text box makes Windows beep.
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void OnAddButton(object sender, EventArgs e)
        {
            var draft = new ButtonDraft(UniqueNewName(), NewButtonPrompt);
            _buttons.Add(draft);
            RebuildButtonList(_buttons.Count - 1);
            RefreshDetailEditors();
            RefreshDetailState();
            RefreshCommitState();

            // Straight into the name, all of it selected: the first thing anybody does with a
            // new button is name it.
            try
            {
                ShowPage(_tabButtons);
                _txtName.Focus();
                _txtName.SelectAll();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Prompt add focus: " + ex.Message);
            }
        }

        private string UniqueNewName()
        {
            for (int suffix = 1; suffix < 1000; suffix++)
            {
                string candidate = suffix == 1 ? NewButtonName : NewButtonName + " " + suffix;
                bool taken = false;
                foreach (ButtonDraft draft in _buttons)
                {
                    if (string.Equals(draft.Name, candidate, StringComparison.OrdinalIgnoreCase))
                    {
                        taken = true;
                        break;
                    }
                }
                if (!taken)
                    return candidate;
            }
            return NewButtonName;
        }

        private void OnRemoveButton(object sender, EventArgs e)
        {
            int index = _lstButtons.SelectedIndex;
            if (index < 0 || index >= _buttons.Count)
                return;

            // No confirmation, on purpose: nothing is written until Apply now, so closing
            // without applying is the undo. Removing every button is legal and means "no quick
            // buttons".
            _buttons.RemoveAt(index);
            RebuildButtonList(index);
            RefreshDetailEditors();
            RefreshDetailState();
            RefreshCommitState();
        }

        private void MoveSelected(int delta)
        {
            int index = _lstButtons.SelectedIndex;
            int target = index + delta;
            if (index < 0 || index >= _buttons.Count || target < 0 || target >= _buttons.Count)
                return;

            ButtonDraft draft = _buttons[index];
            _buttons[index] = _buttons[target];
            _buttons[target] = draft;

            RebuildButtonList(target);
            RefreshDetailEditors();
            RefreshDetailState();
            RefreshCommitState();
        }

        private void OnRestoreDefaultButtons(object sender, EventArgs e)
        {
            _buttons.Clear();
            foreach (PromptButton button in PromptDefaults.CreateButtons())
                _buttons.Add(new ButtonDraft(button.Name, button.Prompt));

            RebuildButtonList(0);
            RefreshDetailEditors();
            RefreshDetailState();
            RefreshCommitState();
        }

        private void OnResetButtonPrompt(object sender, EventArgs e)
        {
            ButtonDraft draft = Selected();
            if (draft == null)
                return;

            string shipped;
            if (!PromptDefaults.TryGetButtonPrompt(draft.Name, out shipped))
                return;

            draft.Prompt = shipped;
            RefreshDetailEditors();
            RefreshDetailState();
            RefreshCommitState();
        }

        // ===== Prompts tab behaviour =====

        private void OnSectionEdited(int index)
        {
            if (_loading)
                return;
            _sectionDraft[index] = _sectionEditors[index].Text;
            RefreshSectionState(index);
            RefreshCommitState();
        }

        private void RestoreSectionDefault(int index)
        {
            string shipped = ToEditorText(PromptDefaults.GetSection(Sections[index]));
            _sectionDraft[index] = shipped;

            _loading = true;
            try
            {
                _sectionEditors[index].Text = shipped;
            }
            finally
            {
                _loading = false;
            }

            RefreshSectionState(index);
            RefreshCommitState();
        }

        /// <summary>
        /// The advisory warnings and the edited marker for one section. The warnings come from
        /// <see cref="PromptLint"/> and are exactly that - advice. Nothing here can stop a save.
        /// </summary>
        private void RefreshSectionState(int index)
        {
            PromptSection section = Sections[index];
            string text = _sectionDraft[index] == null ? "" : _sectionDraft[index];

            bool edited = !PromptDefaults.SameText(text, PromptDefaults.GetSection(section));
            SetLabel(_sectionState[index], edited ? "Edited" : "Unchanged");

            if (!PromptLint.IsChecked(section))
                return;

            IList<string> warnings = PromptLint.Warn(section, text);
            if (warnings.Count == 0)
            {
                SetLabel(_sectionLint[index], "");
                _sectionLint[index].Visible = false;
                return;
            }

            var joined = new List<string>(warnings.Count);
            foreach (string warning in warnings)
                joined.Add("Warning: " + warning);

            SetLabel(_sectionLint[index], string.Join(Environment.NewLine, joined.ToArray()));
            _sectionLint[index].Visible = true;
        }

        // ===== Status =====

        /// <summary>
        /// Keeps the footer line honest about the BUFFERED half of the window. "Apply now" itself
        /// is never disabled: it also re-runs the Outlook and Claude Code reconciles, which are
        /// worth doing whether or not a prompt was touched.
        /// </summary>
        private void RefreshCommitState()
        {
            if (HasChanges())
                SetStatus("Unsaved changes on the Prompts and Buttons tabs.", StatusKind.Neutral);
            else if (_status != StatusKind.Saved)
                SetStatus("No unsaved changes.", StatusKind.Neutral);
        }

        private enum StatusKind
        {
            Neutral,
            Saved,
            Failed,
        }

        private void SetStatus(string text, StatusKind kind)
        {
            _status = kind;
            _lblStatus.Text = text;
            _lblStatus.ForeColor = StatusColour();
        }

        private Color StatusColour()
        {
            switch (_status)
            {
                case StatusKind.Saved:
                    return ThemeService.StatusSuccess;
                case StatusKind.Failed:
                    return ThemeService.StatusError;
                default:
                    return ThemeService.SecondaryText;
            }
        }

        // ===== Section descriptions =====

        private static string SectionTitle(PromptSection section)
        {
            switch (section)
            {
                case PromptSection.Preamble:
                    return "Always sent";
                case PromptSection.ReplyRules:
                    return "Reply rules";
                case PromptSection.SignatureRule:
                    return "Signature rule";
                case PromptSection.SignatureSelection:
                    return "Signature selection";
                default:
                    return section.ToString();
            }
        }

        private static string SectionPurpose(PromptSection section)
        {
            switch (section)
            {
                case PromptSection.Preamble:
                    return "Sent with every writing request: who the model is, that the draft and "
                           + "the quoted thread below it are untrusted content rather than "
                           + "instructions, what shape the answer has to be in, and the language, "
                           + "tone and no-trace-of-AI rules.";
                case PromptSection.ReplyRules:
                    return "Added only when the draft has a quoted thread under it. Saying this "
                           + "when there is no thread invents one.";
                case PromptSection.SignatureRule:
                    return "Added only when the draft already carries a signature. Saying this "
                           + "when there is none loses the sign-off entirely.";
                case PromptSection.SignatureSelection:
                    return "The instruction half of the prompt that picks which signature fits a "
                           + "draft. The signature list, the recipients, the draft and the thread "
                           + "are appended after it.";
                default:
                    return "";
            }
        }

        /// <summary>
        /// Lines of text an editor shows before it scrolls, chosen from how long the shipped text
        /// for that section is. A design value in LINES rather than in pixels, so it means the
        /// same thing at every display scale.
        /// </summary>
        private static int SectionLines(PromptSection section)
        {
            switch (section)
            {
                case PromptSection.Preamble:
                    // Counted from the shipped text, not written down: the no-trace-of-AI rule
                    // is the LAST line of the preamble and the one a user is most likely to have
                    // come here to read, so the box shows the whole block rather than scrolling
                    // that line out of sight. The hand-counted 12 this replaces would have
                    // silently re-created exactly that fault the first time a 13th line was
                    // added to PromptDefaults.Preamble.
                    return LineCount(PromptDefaults.Preamble);
                case PromptSection.ReplyRules:
                    return 3;
                case PromptSection.SignatureRule:
                    return 3;
                case PromptSection.SignatureSelection:
                    // Deliberately fewer than the shipped text's line count: it is the longest
                    // block by some way, it is the one section whose head is its whole point
                    // (role, then untrusted-content warning), and a box tall enough for all of
                    // it would push the rest of the tab off the window on a laptop screen.
                    return 10;
                default:
                    return 6;
            }
        }

        /// <summary>Lines in a block of prompt text, however its line endings are spelled.</summary>
        private static int LineCount(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 1;
            return PromptDefaults.Normalize(text).Split('\n').Length;
        }

        // ===== Small helpers =====

        /// <summary>
        /// CRLF for a text box. Stored text can hold bare LF - the defaults are CRLF, but a
        /// registry value someone edited by hand need not be - and a multi-line TextBox draws a
        /// bare LF as a box character instead of a line break.
        /// </summary>
        private static string ToEditorText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";
            return text.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        /// <summary>Sets a label's text only when it differs, so nothing re-lays out for nothing.</summary>
        private static void SetLabel(Label label, string text)
        {
            if (label.Text != text)
                label.Text = text;
        }

        /// <summary>
        /// One button as the user is editing it. Mutable, unlike <see cref="PromptButton"/>,
        /// which is the immutable value handed to the store at save time. A draft is not a
        /// button yet: its name may be empty and its prompt may be half typed, and neither is
        /// something the store would accept.
        /// </summary>
        private sealed class ButtonDraft
        {
            internal ButtonDraft(string name, string prompt)
            {
                Name = name == null ? "" : name;
                Prompt = prompt == null ? "" : prompt;
            }

            internal string Name { get; set; }

            internal string Prompt { get; set; }
        }
    }
}
