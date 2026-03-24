using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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
        private readonly Outlook.Inspector _owningInspector;
        private readonly Timer _versionTimer;

        // Iterative editing state
        private readonly List<EditTurn> _editHistory = new List<EditTurn>();
        private string _signatureText;   // Cached plain text from _MailAutoSig bookmark
        private string _threadText;      // Cached plain text from _MailOriginal bookmark
        private bool _contextCaptured;   // Whether sig/thread have been read for this email
        private bool _freshDraft;        // When true, send empty draft (no previous AI content)

        // Debug: 7 clicks on version label to enable, accumulates log, copies to clipboard
        private bool _debug;
        private int _debugClickCount;
        private readonly StringBuilder _debugLog = new StringBuilder();

        public AITaskPane(bool isInlineResponse = false, Outlook.Inspector inspector = null)
        {
            _isInlineResponse = isInlineResponse;
            _owningInspector = inspector;
            InitializeComponent();
            ApplyTheme();
            SetupTooltips();
            lblVersion.Click += lblVersion_Click;
            lblVersion.DoubleClick += lblVersion_Click;

            _versionTimer = new Timer();
            _versionTimer.Interval = 1000; // 1 second
            _versionTimer.Tick += (s, ev) => UpdateVersionLabel();
            _versionTimer.Start();
            UpdateVersionLabel();
        }

        private void UpdateVersionLabel()
        {
            var version = "v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var lastChecked = UpdateService.LastChecked;
            var error = UpdateService.LastError;
            var status = UpdateService.Status;

            string suffix;
            if (status != null && status != "up to date")
                suffix = status;
            else if (lastChecked == null)
                suffix = error != null ? null : "checking\u2026";
            else
            {
                var ago = DateTime.Now - lastChecked.Value;
                if (ago.TotalSeconds < 60)
                    suffix = "checked just now";
                else if (ago.TotalMinutes < 60)
                    suffix = $"checked {(int)ago.TotalMinutes}m ago";
                else if (ago.TotalHours < 24)
                    suffix = $"checked {(int)ago.TotalHours}h ago";
                else
                    suffix = $"checked {(int)ago.TotalDays}d ago";
            }

            lblVersion.Text = suffix != null ? $"{version} - {suffix}" : version;
            lnkUpdateError.Visible = error != null;
        }

        private void lblVersion_Click(object sender, EventArgs e)
        {
            if (_debug) return;
            _debugClickCount++;
            if (_debugClickCount >= 7)
            {
                _debug = true;
                _debugLog.Clear();
                _versionTimer.Stop();
                lblVersion.Text = "Debug enabled";
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
                    _debugLog.AppendLine($"Content.End = {doc.Content.End}");
                    doc.Bookmarks.ShowHidden = true;
                    foreach (var bmName in new[] { "_MailAutoSig", "_MailOriginal" })
                    {
                        if (doc.Bookmarks.Exists(bmName))
                        {
                            var bmk = doc.Bookmarks[bmName];
                            int start = bmk.Range.Start, end = bmk.Range.End;
                            string text = bmk.Range.Text ?? "";
                            if (text.Length > 200) text = text.Substring(0, 200) + "...";
                            _debugLog.AppendLine($"  {bmName}: [{start}, {end}] = {text}");
                            ThisAddIn.ReleaseCom(bmk);
                        }
                        else
                        {
                            _debugLog.AppendLine($"  {bmName}: NOT FOUND");
                        }
                    }

                    int draftEnd = FindDraftEnd(doc);
                    var range = doc.Range(0, draftEnd);
                    string draft = range.Text ?? "";
                    if (draft.Length > 300) draft = draft.Substring(0, 300) + "...";
                    _debugLog.AppendLine($"  Draft [0, {draftEnd}]: {draft}");
                    ThisAddIn.ReleaseCom(range);
                }

                if (extra != null)
                    _debugLog.AppendLine($"  {extra}");

                _debugLog.AppendLine();

                // Clipboard requires STA thread — marshal to UI thread
                string logSnapshot = _debugLog.ToString();
                if (this.InvokeRequired)
                    this.Invoke((Action)(() => Clipboard.SetText(logSnapshot)));
                else
                    Clipboard.SetText(logSnapshot);
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

        /// <summary>
        /// Call this when the task pane becomes visible for a new email.
        /// </summary>
        public void ResetForNewEmail()
        {
            txtPrompt.Text = "";
            lblStatus.Visible = false;
            _editHistory.Clear();
            _signatureText = null;
            _threadText = null;
            _contextCaptured = false;
            _freshDraft = false;
        }

        // === Button click handlers ===

        private async void btnProofread_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Proofread);
        }

        private async void btnRevise_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Revise);
        }

        private async void btnShorten_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Shorten);
        }

        private async void btnLengthen_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Lengthen);
        }

        private async void btnFormal_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Formal);
        }

        private async void btnFriendly_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Friendly);
        }

        private async void btnDraft_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPrompt.Text))
            {
                ShowStatus("Please enter instructions for the email you want to draft.", true);
                return;
            }
            // Draft Email = fresh start — clear history, re-read context, send empty draft
            _editHistory.Clear();
            _contextCaptured = false;
            _freshDraft = true;
            await ProcessAction(ClaudeService.ActionType.Draft, txtPrompt.Text);
        }

        private async void btnEditDraft_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPrompt.Text))
            {
                ShowStatus("Please enter instructions for editing the draft.", true);
                return;
            }
            await ProcessAction(ClaudeService.ActionType.Draft, txtPrompt.Text);
        }

        private async void btnEditSelection_Click(object sender, EventArgs e)
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
            await ProcessAction(ClaudeService.ActionType.Draft, txtPrompt.Text, selectedText);
        }

        // === Core processing ===

        private async Task ProcessAction(ClaudeService.ActionType action, string prompt = "", string selectedText = null)
        {
            SetUIEnabled(false);
            ShowStatus("Processing...", false);

            try
            {
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
                        SetUIEnabled(true);
                        return;
                    }

                    dynamic wordDoc = doc;
                    wordDoc.Bookmarks.ShowHidden = true;
                    DebugLog($"ProcessAction({action}) BEFORE read", wordDoc);

                    // Capture signature and thread context on first interaction
                    if (!_contextCaptured)
                    {
                        _signatureText = ReadSignatureText(wordDoc);
                        _threadText = ReadThreadText(wordDoc);
                        _contextCaptured = true;
                    }

                    signatureText = _signatureText;
                    threadText = _threadText;

                    // Read current draft (always re-read to capture manual edits)
                    draftText = ReadDraftText(wordDoc);
                }
                finally
                {
                    ThisAddIn.ReleaseCom(doc);
                }

                // Fresh draft: send empty text so Claude drafts from scratch
                if (_freshDraft)
                {
                    _freshDraft = false;
                    draftText = "";
                }

                DebugLog("Sending to Claude", extra:
                    $"draftText ({draftText.Length} chars), sigText ({signatureText.Length} chars), threadText ({threadText.Length} chars)");

                string result = await ClaudeService.ProcessEmailAsync(
                    action, prompt, _editHistory,
                    draftText, signatureText, threadText, selectedText);

                DebugLog("Claude returned", extra: $"result ({result.Length} chars): {(result.Length > 300 ? result.Substring(0, 300) + "..." : result)}");

                InvokeOnUI(() =>
                {
                    if (WriteDraftToDocument(result))
                    {
                        _editHistory.Add(new EditTurn
                        {
                            Action = action,
                            Instruction = prompt,
                            SelectedText = selectedText,
                            Result = result
                        });

                        ShowStatus("Done!", false);
                    }
                    SetUIEnabled(true);
                });
            }
            catch (Exception ex)
            {
                InvokeOnUI(() =>
                {
                    string msg = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    ShowStatus(msg, true);
                    SetUIEnabled(true);
                });
            }
        }

        // === Word Object Model helpers ===

        private object GetWordDocument()
        {
            try
            {
                if (!_isInlineResponse)
                    return _owningInspector?.WordEditor;

                // Inline response: use Explorer.ActiveInlineResponseWordEditor (Outlook 2016+).
                // Accessed via late binding for compatibility with older Outlook PIAs.
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
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Finds the draft boundary: _MailAutoSig.Start ?? _MailOriginal.Start ?? Content.End
        /// </summary>
        private int FindDraftEnd(dynamic doc)
        {
            if (doc.Bookmarks.Exists("_MailAutoSig"))
            {
                var bmk = doc.Bookmarks["_MailAutoSig"];
                int pos = bmk.Range.Start;
                ThisAddIn.ReleaseCom(bmk);
                return pos;
            }
            if (doc.Bookmarks.Exists("_MailOriginal"))
            {
                var bmk = doc.Bookmarks["_MailOriginal"];
                int pos = bmk.Range.Start;
                ThisAddIn.ReleaseCom(bmk);
                return pos;
            }
            return doc.Content.End;
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
            if (!doc.Bookmarks.Exists("_MailAutoSig"))
                return "";

            var bmk = doc.Bookmarks["_MailAutoSig"];
            string text = bmk.Range.Text ?? "";
            ThisAddIn.ReleaseCom(bmk);

            return text.TrimEnd('\r', '\n');
        }

        private string ReadThreadText(dynamic doc)
        {
            if (!doc.Bookmarks.Exists("_MailOriginal"))
                return "";

            var bmk = doc.Bookmarks["_MailOriginal"];
            string text = bmk.Range.Text ?? "";
            ThisAddIn.ReleaseCom(bmk);

            return text.TrimEnd('\r', '\n');
        }

        private bool WriteDraftToDocument(string newDraftText)
        {
            object doc = null;
            try
            {
                doc = GetWordDocument();
                if (doc == null)
                {
                    ShowStatus("Could not access email editor.", true);
                    return false;
                }

                dynamic wordDoc = doc;
                wordDoc.Bookmarks.ShowHidden = true;
                DebugLog("WriteDraft BEFORE", wordDoc);

                // Determine which bookmark is the draft boundary and save its extent.
                // Delete it before writing so it doesn't absorb our text.
                string boundaryBookmark = null;
                int draftEnd = wordDoc.Content.End;
                int origBmkEnd = -1;

                if (wordDoc.Bookmarks.Exists("_MailAutoSig"))
                    boundaryBookmark = "_MailAutoSig";
                else if (wordDoc.Bookmarks.Exists("_MailOriginal"))
                    boundaryBookmark = "_MailOriginal";

                if (boundaryBookmark != null)
                {
                    var bmk = wordDoc.Bookmarks[boundaryBookmark];
                    draftEnd = bmk.Range.Start;
                    origBmkEnd = bmk.Range.End;
                    bmk.Delete(); // removes bookmark marker, not the text
                    ThisAddIn.ReleaseCom(bmk);
                }

                var range = wordDoc.Range(0, draftEnd);
                range.Text = newDraftText + "\r\n";
                int newDraftEnd = range.End;
                ThisAddIn.ReleaseCom(range);

                // Re-create the boundary bookmark at the adjusted position
                if (boundaryBookmark != null && origBmkEnd >= 0)
                {
                    int newBmkEnd = origBmkEnd + (newDraftEnd - draftEnd);
                    var bmkRange = wordDoc.Range(newDraftEnd, newBmkEnd);
                    wordDoc.Bookmarks.Add(boundaryBookmark, bmkRange);
                    ThisAddIn.ReleaseCom(bmkRange);
                }

                DebugLog("WriteDraft AFTER", wordDoc);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("WriteDraftToDocument error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
            finally
            {
                ThisAddIn.ReleaseCom(doc);
            }
        }

        // === Selection support ===

        private string GetSelectedText()
        {
            object doc = null;
            try
            {
                doc = GetWordDocument();
                if (doc == null) return null;
                string text = ((dynamic)doc).Application.Selection.Text as string;
                // Word appends a trailing paragraph mark to selections
                if (!string.IsNullOrEmpty(text) && text.EndsWith("\r"))
                    text = text.Substring(0, text.Length - 1);
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }
            catch
            {
                return null;
            }
            finally
            {
                ThisAddIn.ReleaseCom(doc);
            }
        }

        // === UI helpers ===

        private void SetupTooltips()
        {
            var tip = new ToolTip();
            tip.SetToolTip(btnProofread, "Fix any spelling, grammar, and punctuation errors.\nKeep the tone, meaning, and structure unchanged.");
            tip.SetToolTip(btnRevise, "Improve clarity, flow, and word choice.\nPreserve the original meaning and tone.");
            tip.SetToolTip(btnShorten, "Make the email more concise.\nRemove filler and redundancy while keeping all key points.");
            tip.SetToolTip(btnLengthen, "Expand the email with more detail, context, or explanation.\nKeep the same tone and intent.");
            tip.SetToolTip(btnFormal, "Rewrite in a more formal, professional tone.\nKeep the same content and meaning.");
            tip.SetToolTip(btnFriendly, "Rewrite in a warmer, more conversational tone.\nKeep the same content and meaning.");
            tip.SetToolTip(btnDraft, "Draft a new email from scratch based on your instruction.\nClears any previous AI draft.");
            tip.SetToolTip(btnEditDraft, "Edit the current draft based on your instruction.\nPreserves conversation history for iterative refinement.");
            tip.SetToolTip(btnEditSelection, "Edit only the selected text based on your instruction.\nLeaves the rest of the draft unchanged.");
        }

        private void InvokeOnUI(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated)
                return;

            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        private void ApplyTheme()
        {
            if (!ThemeService.IsDarkMode)
                return;

            // Main background
            this.ForeColor = ThemeService.Text;

            // Group boxes
            foreach (var grp in new[] { grpQuickActions, grpInstruction })
            {
                grp.ForeColor = ThemeService.Text;
            }

            // Text boxes
            foreach (var txt in new[] { txtPrompt })
            {
                txt.BackColor = ThemeService.TextBoxBackground;
                txt.ForeColor = ThemeService.Text;
            }

            // Buttons
            foreach (Control ctrl in this.Controls)
                ApplyThemeToButtons(ctrl);

            // Status label will be themed via ShowStatus
        }

        private void ApplyThemeToButtons(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Button btn)
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderColor = ThemeService.Border;
                    btn.BackColor = ThemeService.ButtonFace;
                    btn.ForeColor = ThemeService.ButtonText;
                }
                if (ctrl.HasChildren)
                    ApplyThemeToButtons(ctrl);
            }
        }

        private void ShowStatus(string message, bool isError)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = isError ? ThemeService.StatusError : ThemeService.StatusSuccess;
            lblStatus.Visible = true;
        }

        private void SetUIEnabled(bool enabled)
        {
            btnProofread.Enabled = enabled;
            btnRevise.Enabled = enabled;
            btnShorten.Enabled = enabled;
            btnLengthen.Enabled = enabled;
            btnFormal.Enabled = enabled;
            btnFriendly.Enabled = enabled;
            btnDraft.Enabled = enabled;
            btnEditDraft.Enabled = enabled;
            btnEditSelection.Enabled = enabled;
            txtPrompt.Enabled = enabled;
        }

        partial void DisposeCustomResources()
        {
            _versionTimer?.Stop();
            _versionTimer?.Dispose();
            ThisAddIn.ReleaseCom(_owningInspector);
            // Don't call ClaudeService.Shutdown() here -- it kills the shared
            // warm process that other panes still need. ThisAddIn_Shutdown
            // handles final cleanup.
        }
    }
}
