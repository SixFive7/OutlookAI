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

        // Iterative editing state
        private readonly List<EditTurn> _editHistory = new List<EditTurn>();
        private bool _freshDraft;
        private bool _isProcessing;

        // Debug: 7 clicks within 3 seconds to enable
        private bool _debug;
        private int _debugClickCount;
        private DateTime _debugFirstClick;
        private readonly StringBuilder _debugLog = new StringBuilder();
        private bool _lastStatusError;

        public AITaskPane(bool isInlineResponse = false, Outlook.Inspector inspector = null)
        {
            _isInlineResponse = isInlineResponse;
            _owningInspector = inspector;
            InitializeComponent();
            ApplyTheme();
            ThemeService.ThemeChanged += OnThemeChanged;
            SetupTooltips();
            lblVersion.Click += lblVersion_Click;
            lblVersion.DoubleClick += lblVersion_Click;

            _versionTimer = new Timer();
            _versionTimer.Interval = 1000;
            _versionTimer.Tick += (s, ev) => UpdateVersionLabel();
            _versionTimer.Start();
            UpdateVersionLabel();
        }

        private void UpdateVersionLabel()
        {
            if (_disposed || IsDisposed) return;

            var version = "v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var lastChecked = UpdateService.LastChecked;
            var error = UpdateService.LastError;
            var status = UpdateService.Status;

            string suffix;
            if (status != null && status != "up to date")
                suffix = status;
            else if (lastChecked == null)
                suffix = error != null ? null : "checking…";
            else
            {
                var ago = DateTime.Now - lastChecked.Value;
                if (ago < TimeSpan.Zero) ago = TimeSpan.Zero;
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
        }

        // === Button click handlers ===

        private async void btnProofread_Click(object sender, EventArgs e)
        {
            try { await ProcessAction(ClaudeService.ActionType.Proofread); }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnRevise_Click(object sender, EventArgs e)
        {
            try { await ProcessAction(ClaudeService.ActionType.Revise); }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnShorten_Click(object sender, EventArgs e)
        {
            try { await ProcessAction(ClaudeService.ActionType.Shorten); }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnLengthen_Click(object sender, EventArgs e)
        {
            try { await ProcessAction(ClaudeService.ActionType.Lengthen); }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnFormal_Click(object sender, EventArgs e)
        {
            try { await ProcessAction(ClaudeService.ActionType.Formal); }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        private async void btnFriendly_Click(object sender, EventArgs e)
        {
            try { await ProcessAction(ClaudeService.ActionType.Friendly); }
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
                await ProcessAction(ClaudeService.ActionType.Draft, txtPrompt.Text);
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
                await ProcessAction(ClaudeService.ActionType.Draft, txtPrompt.Text);
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
                await ProcessAction(ClaudeService.ActionType.Draft, txtPrompt.Text, selectedText);
            }
            catch (Exception ex) { ShowStatus(ex.Message, true); }
        }

        // === Core processing ===

        private async Task ProcessAction(ClaudeService.ActionType action, string prompt = "", string selectedText = null)
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
                    DebugLog($"ProcessAction({action}) BEFORE read", wordDoc);

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
                    action, prompt, _editHistory,
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
                            Action = action,
                            Instruction = prompt,
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

        private void SetupTooltips()
        {
            _toolTip = new ToolTip();
            _toolTip.SetToolTip(btnProofread, "Fix any spelling, grammar, and punctuation errors.\nKeep the tone, meaning, and structure unchanged.");
            _toolTip.SetToolTip(btnRevise, "Improve clarity, flow, and word choice.\nPreserve the original meaning and tone.");
            _toolTip.SetToolTip(btnShorten, "Make the email more concise.\nRemove filler and redundancy while keeping all key points.");
            _toolTip.SetToolTip(btnLengthen, "Expand the email with more detail, context, or explanation.\nKeep the same tone and intent.");
            _toolTip.SetToolTip(btnFormal, "Rewrite in a more formal, professional tone.\nKeep the same content and meaning.");
            _toolTip.SetToolTip(btnFriendly, "Rewrite in a warmer, more conversational tone.\nKeep the same content and meaning.");
            _toolTip.SetToolTip(btnDraft, "Draft a new email from scratch based on your instruction.\nClears any previous AI draft.");
            _toolTip.SetToolTip(btnEditDraft, "Edit the current draft based on your instruction.\nPreserves conversation history for iterative refinement.");
            _toolTip.SetToolTip(btnEditSelection, "Edit only the selected text based on your instruction.\nLeaves the rest of the draft unchanged.");
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
            _disposed = true;
            ThemeService.ThemeChanged -= OnThemeChanged;
            _versionTimer?.Stop();
            _versionTimer?.Dispose();
            _toolTip?.Dispose();
            var inspector = _owningInspector;
            _owningInspector = null;
            ThisAddIn.ReleaseCom(inspector);
        }
    }
}
