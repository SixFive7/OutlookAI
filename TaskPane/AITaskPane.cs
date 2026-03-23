using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
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
        private string _htmlPrefix;      // HTML from start up to and including <body...> tag
        private string _signatureHtml;   // Signature HTML block (preserved across draft writes)
        private string _threadHtml;      // HTML from reply boundary to end (includes </body></html>)
        private bool _threadCaptured;
        private string _firstTurnEmailContent; // Full plain text sent on first turn only

        public AITaskPane(bool isInlineResponse = false, Outlook.Inspector inspector = null)
        {
            _isInlineResponse = isInlineResponse;
            _owningInspector = inspector;
            InitializeComponent();
            ApplyTheme();

            // Selection-based buttons require WordEditor, only available for Inspector windows
            if (_isInlineResponse)
            {
                btnDraftSelection.Enabled = false;
                btnCustomSelection.Enabled = false;
            }

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
            txtDraftPrompt.Text = "";
            txtCustomPrompt.Text = "";
            lblStatus.Visible = false;
            _editHistory.Clear();
            _htmlPrefix = null;
            _signatureHtml = null;
            _threadHtml = null;
            _threadCaptured = false;
            _firstTurnEmailContent = null;
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
            if (string.IsNullOrWhiteSpace(txtDraftPrompt.Text))
            {
                ShowStatus("Please enter instructions for the email you want to draft.", true);
                return;
            }
            // Draft Email = fresh start — clear history and re-capture email structure
            _editHistory.Clear();
            _threadCaptured = false;
            _firstTurnEmailContent = null;
            await ProcessAction(ClaudeService.ActionType.Draft, txtDraftPrompt.Text);
        }

        private async void btnEditDraft_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDraftPrompt.Text))
            {
                ShowStatus("Please enter instructions for editing the draft.", true);
                return;
            }
            // Edit Draft = iterative — keep history, continue conversation
            await ProcessAction(ClaudeService.ActionType.Draft, txtDraftPrompt.Text);
        }

        private async void btnDraftSelection_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDraftPrompt.Text))
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
            await ProcessAction(ClaudeService.ActionType.Draft, txtDraftPrompt.Text, selectedText);
        }

        private async void btnCustom_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCustomPrompt.Text))
            {
                ShowStatus("Please enter instructions for the custom action.", true);
                return;
            }
            await ProcessAction(ClaudeService.ActionType.Custom, txtCustomPrompt.Text);
        }

        private async void btnCustomSelection_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCustomPrompt.Text))
            {
                ShowStatus("Please enter instructions for the custom action.", true);
                return;
            }
            string selectedText = GetSelectedText();
            if (string.IsNullOrWhiteSpace(selectedText))
            {
                ShowStatus("Please select text in the email editor first.", true);
                return;
            }
            await ProcessAction(ClaudeService.ActionType.Custom, txtCustomPrompt.Text, selectedText);
        }

        // === Core processing ===

        private async Task ProcessAction(ClaudeService.ActionType action, string prompt = "", string selectedText = null)
        {
            SetUIEnabled(false);
            ShowStatus("Processing...", false);

            try
            {
                // Capture email structure on first interaction
                if (!_threadCaptured)
                {
                    string htmlBody = GetEmailHtmlBody();
                    CaptureEmailStructure(htmlBody);
                    _threadCaptured = true;
                    // Store full plain text for the first turn's context
                    _firstTurnEmailContent = GetEmailPlainBody();
                }

                // Get current draft from the live email (captures any manual edits)
                string currentDraft = ExtractCurrentDraftText();

                // Email content is only sent on the first turn (when history is empty)
                string emailContent = _editHistory.Count == 0 ? _firstTurnEmailContent : null;

                // Signature context — sent every turn so Claude knows not to add a sign-off
                string signatureText = !string.IsNullOrEmpty(_signatureHtml)
                    ? HtmlToPlainText(_signatureHtml) : null;

                string result = await ClaudeService.ProcessEmailAsync(
                    action, prompt, _editHistory, emailContent, currentDraft, signatureText, selectedText);

                InvokeOnUI(() =>
                {
                    if (WriteDraftToEmail(result))
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

        // === Email access ===

        private Outlook.MailItem GetCurrentMailItem(out Outlook.Explorer explorer)
        {
            explorer = null;
            if (_isInlineResponse)
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                object rawInline = explorer?.ActiveInlineResponse;
                var mail = rawInline as Outlook.MailItem;
                if (mail == null)
                    ThisAddIn.ReleaseCom(rawInline);
                return mail;
            }

            // Use the owning inspector rather than ActiveInspector so that
            // clicking a button in a background window processes the right email.
            object rawItem = _owningInspector?.CurrentItem;
            var mailItem = rawItem as Outlook.MailItem;
            if (mailItem == null)
                ThisAddIn.ReleaseCom(rawItem);
            return mailItem;
        }

        private string GetEmailHtmlBody()
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            try
            {
                mail = GetCurrentMailItem(out explorer);
                return mail?.HTMLBody ?? "";
            }
            catch (Exception ex)
            {
                Debug.WriteLine("GetEmailHtmlBody error: " + ex.Message);
                return "";
            }
            finally
            {
                ThisAddIn.ReleaseCom(mail);
                ThisAddIn.ReleaseCom(explorer);
            }
        }

        private string GetEmailPlainBody()
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            try
            {
                mail = GetCurrentMailItem(out explorer);
                return mail?.Body ?? "";
            }
            catch (Exception ex)
            {
                Debug.WriteLine("GetEmailPlainBody error: " + ex.Message);
                return "";
            }
            finally
            {
                ThisAddIn.ReleaseCom(mail);
                ThisAddIn.ReleaseCom(explorer);
            }
        }

        // === HTML boundary detection and draft extraction ===

        private void CaptureEmailStructure(string htmlBody)
        {
            if (string.IsNullOrWhiteSpace(htmlBody))
            {
                _htmlPrefix = "<html><body>";
                _signatureHtml = "";
                _threadHtml = "</body></html>";
                return;
            }

            int bodyTagEnd = FindBodyTagEnd(htmlBody);
            if (bodyTagEnd < 0)
            {
                _htmlPrefix = "<html><body>";
                _signatureHtml = "";
                _threadHtml = "</body></html>";
                return;
            }

            _htmlPrefix = htmlBody.Substring(0, bodyTagEnd);

            int boundary = FindDraftBoundary(htmlBody);
            if (boundary >= 0)
            {
                // Everything from boundary onward is the thread (preserved on each write)
                _threadHtml = htmlBody.Substring(boundary);
            }
            else
            {
                // No thread boundary — this is a new email (not a reply)
                int bodyClose = htmlBody.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                _threadHtml = bodyClose >= 0 ? htmlBody.Substring(bodyClose) : "</body></html>";
            }

            // Extract signature from the draft area (between body tag and thread boundary).
            // Outlook places the signature before appendonsend/divRplyFwdMsg.
            int draftEnd = boundary >= 0 ? boundary : htmlBody.Length;
            string draftArea = htmlBody.Substring(bodyTagEnd, draftEnd - bodyTagEnd);
            int sigStart = FindSignatureStart(draftArea);
            if (sigStart >= 0)
            {
                _signatureHtml = draftArea.Substring(sigStart);
            }
            else
            {
                _signatureHtml = "";
            }
        }

        private string ExtractCurrentDraftText()
        {
            string htmlBody = GetEmailHtmlBody();
            if (string.IsNullOrWhiteSpace(htmlBody))
                return "";

            int bodyTagEnd = FindBodyTagEnd(htmlBody);
            if (bodyTagEnd < 0)
                return HtmlToPlainText(htmlBody);

            int boundary = FindDraftBoundary(htmlBody);
            if (boundary >= 0)
            {
                // Extract only the draft portion (before the reply boundary)
                string draftAreaHtml = htmlBody.Substring(bodyTagEnd, boundary - bodyTagEnd);

                // Exclude the signature from the draft text sent to Claude
                if (!string.IsNullOrEmpty(_signatureHtml))
                {
                    int sigIdx = FindSignatureStart(draftAreaHtml);
                    if (sigIdx >= 0)
                        draftAreaHtml = draftAreaHtml.Substring(0, sigIdx);
                }

                return HtmlToPlainText(draftAreaHtml);
            }

            // No boundary found — send everything (safe default, as discussed)
            int bodyClose = htmlBody.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            string allContentHtml = bodyClose >= 0
                ? htmlBody.Substring(bodyTagEnd, bodyClose - bodyTagEnd)
                : htmlBody.Substring(bodyTagEnd);
            return HtmlToPlainText(allContentHtml);
        }

        private bool WriteDraftToEmail(string plainTextDraft)
        {
            Outlook.MailItem mail = null;
            Outlook.Explorer explorer = null;
            try
            {
                mail = GetCurrentMailItem(out explorer);
                if (mail == null)
                {
                    ShowStatus("Could not find active email window.", true);
                    return false;
                }

                string draftHtml = PlainTextToHtml(plainTextDraft);
                mail.HTMLBody = _htmlPrefix + draftHtml + _signatureHtml + _threadHtml;
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("WriteDraftToEmail error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
            finally
            {
                ThisAddIn.ReleaseCom(mail);
                ThisAddIn.ReleaseCom(explorer);
            }
        }

        // === Selection support ===

        private string GetSelectedText()
        {
            if (_isInlineResponse)
                return null;

            object doc = null;
            try
            {
                doc = _owningInspector?.WordEditor;
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

        // === HTML utilities ===

        private static int FindBodyTagEnd(string html)
        {
            int bodyStart = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (bodyStart < 0) return -1;
            int bodyEnd = html.IndexOf('>', bodyStart);
            return bodyEnd >= 0 ? bodyEnd + 1 : -1;
        }

        /// <summary>
        /// Finds the boundary between the user's draft area and the quoted thread.
        /// Returns the index of the opening tag of the boundary element, or -1 if not found.
        /// Checks for appendonsend first (preserves Outlook signature insertion),
        /// then divRplyFwdMsg (reply/forward header), then border-top separator.
        /// </summary>
        private static int FindDraftBoundary(string html)
        {
            // appendonsend — Outlook's signature insertion point (preserves signature on send)
            int idx = FindTagWithAttribute(html, "id", "appendonsend");
            if (idx >= 0) return idx;

            // divRplyFwdMsg — reply/forward header block
            idx = FindTagWithAttribute(html, "id", "divRplyFwdMsg");
            if (idx >= 0) return idx;

            // border-top separator (older Outlook versions)
            idx = html.IndexOf("border-top:solid #E1E1E1", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                // Walk back to find the opening '<' of this element
                for (int i = idx - 1; i >= 0; i--)
                {
                    if (html[i] == '<') return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Finds the start of the email signature within a draft area HTML fragment.
        /// Looks for common Outlook signature markers: id containing "signature",
        /// or class "MsoSignature". Returns the index of the opening tag, or -1.
        /// </summary>
        private static int FindSignatureStart(string html)
        {
            // Outlook desktop: <div id="Signature">, <div id="signature_...">, etc.
            int idx = FindTagWithAttributeContaining(html, "id", "signature");
            if (idx >= 0) return idx;

            // Outlook web / some configurations: class="MsoSignature"
            idx = FindTagWithAttributeContaining(html, "class", "MsoSignature");
            if (idx >= 0) return idx;

            // Some Outlook versions use id="ms-outlook-mobile-signature"
            idx = FindTagWithAttributeContaining(html, "id", "ms-outlook");
            if (idx >= 0) return idx;

            return -1;
        }

        private static int FindTagWithAttributeContaining(string html, string attr, string substring)
        {
            string prefix = attr + "=\"";
            int searchFrom = 0;
            while (searchFrom < html.Length)
            {
                int attrIdx = html.IndexOf(prefix, searchFrom, StringComparison.OrdinalIgnoreCase);
                if (attrIdx < 0) return -1;
                int valueStart = attrIdx + prefix.Length;
                int valueEnd = html.IndexOf('"', valueStart);
                if (valueEnd < 0) return -1;
                string value = html.Substring(valueStart, valueEnd - valueStart);
                if (value.IndexOf(substring, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Walk back to the opening '<'
                    for (int i = attrIdx - 1; i >= 0; i--)
                    {
                        if (html[i] == '<') return i;
                    }
                }
                searchFrom = valueEnd + 1;
            }
            return -1;
        }

        private static int FindTagWithAttribute(string html, string attr, string value)
        {
            string pattern = attr + "=\"" + value + "\"";
            int idx = html.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return -1;
            // Walk back to find the opening '<'
            for (int i = idx - 1; i >= 0; i--)
            {
                if (html[i] == '<') return i;
            }
            return -1;
        }

        private static string HtmlToPlainText(string html)
        {
            if (string.IsNullOrEmpty(html)) return "";
            // Replace block-level closers with newlines before stripping tags
            string text = Regex.Replace(html, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"</p>", "\n\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"</div>", "\n", RegexOptions.IgnoreCase);
            // Strip all remaining HTML tags
            text = Regex.Replace(text, @"<[^>]+>", "");
            // Decode HTML entities
            text = WebUtility.HtmlDecode(text);
            // Collapse excessive blank lines
            text = Regex.Replace(text, @"\n{3,}", "\n\n");
            return text.Trim();
        }

        private static string PlainTextToHtml(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            string encoded = WebUtility.HtmlEncode(text);
            // Split on double newlines into paragraphs
            var paragraphs = Regex.Split(encoded, @"\r?\n\r?\n");
            var sb = new StringBuilder();
            foreach (var para in paragraphs)
            {
                if (string.IsNullOrWhiteSpace(para)) continue;
                string content = para.Replace("\r\n", "<br>").Replace("\n", "<br>");
                sb.Append("<p style=\"margin:0\">").Append(content).Append("</p>");
            }
            return sb.ToString();
        }

        // === UI helpers ===

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
            foreach (var grp in new[] { grpQuickActions, grpDraft, grpCustom })
            {
                grp.ForeColor = ThemeService.Text;
            }

            // Text boxes
            foreach (var txt in new[] { txtDraftPrompt, txtCustomPrompt })
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
            btnDraftSelection.Enabled = enabled && !_isInlineResponse;
            txtDraftPrompt.Enabled = enabled;
            btnCustom.Enabled = enabled;
            btnCustomSelection.Enabled = enabled && !_isInlineResponse;
            txtCustomPrompt.Enabled = enabled;
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
