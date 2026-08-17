namespace OutlookAI.TaskPane
{
    partial class AITaskPane
    {
        private System.ComponentModel.IContainer components = null;

        partial void DisposeCustomResources();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                DisposeCustomResources();
                components?.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpQuickActions = new System.Windows.Forms.GroupBox();
            this.flowQuickActions = new System.Windows.Forms.FlowLayoutPanel();
            this.btnSelectSignature = new System.Windows.Forms.Button();
            this.grpInstruction = new System.Windows.Forms.GroupBox();
            this.txtPrompt = new System.Windows.Forms.TextBox();
            this.btnDraft = new System.Windows.Forms.Button();
            this.btnEditDraft = new System.Windows.Forms.Button();
            this.btnEditSelection = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lnkUpdateError = new System.Windows.Forms.LinkLabel();
            this.lnkCheckUpdates = new System.Windows.Forms.LinkLabel();
            this.grpQuickActions.SuspendLayout();
            this.flowQuickActions.SuspendLayout();
            this.grpInstruction.SuspendLayout();
            this.SuspendLayout();

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = Services.ThemeService.Accent;
            this.lblTitle.Location = new System.Drawing.Point(10, 10);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(150, 21);
            this.lblTitle.Text = "AI Writing Assistant";

            // grpQuickActions
            // Only the WIDTH is set here: it anchors to the pane, so it has to start out right.
            // The height is whatever the buttons inside end up needing and is computed by
            // LayoutPane(), which also places everything below this group. The 56 below is what
            // that comes to with no quick buttons at all, so the value here is never a lie.
            this.grpQuickActions.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.grpQuickActions.Controls.Add(this.flowQuickActions);
            this.grpQuickActions.Controls.Add(this.btnSelectSignature);
            this.grpQuickActions.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            // Named, not a literal: UiScale reads this back to recover the factor
            // AutoScaleMode.Font applied, so the two have to be the same number.
            this.grpQuickActions.Location = new System.Drawing.Point(PaneMargin, QuickActionsDesignTop);
            this.grpQuickActions.Name = "grpQuickActions";
            this.grpQuickActions.Size = new System.Drawing.Size(240, 56);
            this.grpQuickActions.TabIndex = 0;
            this.grpQuickActions.TabStop = false;
            this.grpQuickActions.Text = "Quick Actions (Edit Current Email)";

            // flowQuickActions
            // Host for the quick-action buttons. There are no buttons here because there is no
            // fixed set of them any more: the pane builds one per entry in the user's saved button
            // list (PromptStore) every time that list changes. Wrapping the buttons onto as many
            // rows as they need is this panel's job, which is what lets a caption of any length
            // and any button count lay out without a hard-coded grid to outgrow.
            this.flowQuickActions.AutoSize = false;
            this.flowQuickActions.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.flowQuickActions.Margin = new System.Windows.Forms.Padding(0);
            this.flowQuickActions.Name = "flowQuickActions";
            this.flowQuickActions.Padding = new System.Windows.Forms.Padding(0);
            this.flowQuickActions.TabIndex = 0;
            this.flowQuickActions.WrapContents = true;

            // btnSelectSignature
            // Location and width come from LayoutPane(): it sits directly under the last row of
            // quick buttons, and there is no telling from here how many rows that is.
            this.btnSelectSignature.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.btnSelectSignature.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnSelectSignature.Location = new System.Drawing.Point(PaneMargin, GroupTopInset);
            this.btnSelectSignature.Name = "btnSelectSignature";
            this.btnSelectSignature.Size = new System.Drawing.Size(220, 26);
            this.btnSelectSignature.TabIndex = 1;
            this.btnSelectSignature.Text = "Select the best signature";
            this.btnSelectSignature.UseVisualStyleBackColor = true;
            this.btnSelectSignature.Click += new System.EventHandler(this.btnSelectSignature_Click);

            // grpInstruction
            this.grpInstruction.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.grpInstruction.Controls.Add(this.txtPrompt);
            this.grpInstruction.Controls.Add(this.btnDraft);
            this.grpInstruction.Controls.Add(this.btnEditDraft);
            this.grpInstruction.Controls.Add(this.btnEditSelection);
            this.grpInstruction.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            // Y comes from LayoutPane(), measured off the bottom of the quick-action group. The
            // literal that used to be here was derived from that group's old fixed height, so a
            // seventh button clipped and a taller group overlapped this one.
            this.grpInstruction.Location = new System.Drawing.Point(PaneMargin, 0);
            this.grpInstruction.Name = "grpInstruction";
            this.grpInstruction.Size = new System.Drawing.Size(240, 170);
            this.grpInstruction.TabIndex = 1;
            this.grpInstruction.TabStop = false;
            this.grpInstruction.Text = "Instruction";

            // txtPrompt
            this.txtPrompt.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.txtPrompt.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.txtPrompt.Location = new System.Drawing.Point(10, 22);
            this.txtPrompt.Multiline = true;
            this.txtPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtPrompt.Size = new System.Drawing.Size(220, 50);

            // btnDraft
            this.btnDraft.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.btnDraft.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDraft.Location = new System.Drawing.Point(10, 78);
            this.btnDraft.Size = new System.Drawing.Size(220, 26);
            this.btnDraft.Text = "Draft new email";
            this.btnDraft.UseVisualStyleBackColor = true;
            this.btnDraft.Click += new System.EventHandler(this.btnDraft_Click);

            // btnEditDraft
            this.btnEditDraft.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.btnEditDraft.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnEditDraft.Location = new System.Drawing.Point(10, 108);
            this.btnEditDraft.Size = new System.Drawing.Size(220, 26);
            this.btnEditDraft.Text = "Edit current draft";
            this.btnEditDraft.UseVisualStyleBackColor = true;
            this.btnEditDraft.Click += new System.EventHandler(this.btnEditDraft_Click);

            // btnEditSelection
            this.btnEditSelection.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.btnEditSelection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnEditSelection.Location = new System.Drawing.Point(10, 138);
            this.btnEditSelection.Size = new System.Drawing.Size(220, 26);
            this.btnEditSelection.Text = "Edit selection only";
            this.btnEditSelection.UseVisualStyleBackColor = true;
            this.btnEditSelection.Click += new System.EventHandler(this.btnEditSelection_Click);

            // lblStatus
            this.lblStatus.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 8F);
            // Y comes from LayoutPane(), same reason as grpInstruction above.
            this.lblStatus.Location = new System.Drawing.Point(PaneMargin, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(240, 32);
            this.lblStatus.Visible = false;

            // lnkUpdateError
            this.lnkUpdateError.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lnkUpdateError.Font = new System.Drawing.Font("Segoe UI", 7F);
            this.lnkUpdateError.LinkColor = Services.ThemeService.LinkError;
            this.lnkUpdateError.Name = "lnkUpdateError";
            this.lnkUpdateError.Size = new System.Drawing.Size(260, 14);
            this.lnkUpdateError.TabStop = false;
            this.lnkUpdateError.Text = "update error";
            this.lnkUpdateError.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lnkUpdateError.Visible = false;
            this.lnkUpdateError.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkUpdateError_LinkClicked);

            // lnkCheckUpdates
            this.lnkCheckUpdates.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lnkCheckUpdates.Font = new System.Drawing.Font("Segoe UI", 7F);
            this.lnkCheckUpdates.LinkColor = Services.ThemeService.Accent;
            this.lnkCheckUpdates.DisabledLinkColor = Services.ThemeService.SecondaryText;
            this.lnkCheckUpdates.Name = "lnkCheckUpdates";
            this.lnkCheckUpdates.Size = new System.Drawing.Size(260, 14);
            this.lnkCheckUpdates.TabStop = false;
            this.lnkCheckUpdates.Text = "check for updates";
            this.lnkCheckUpdates.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lnkCheckUpdates.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkCheckUpdates_LinkClicked);

            // lblVersion
            this.lblVersion.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblVersion.Font = new System.Drawing.Font("Segoe UI", 7.5F);
            this.lblVersion.ForeColor = Services.ThemeService.SecondaryText;
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(260, 18);
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // AITaskPane
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = Services.ThemeService.Background;
            // These three are docked Bottom, and for those the order below is what decides
            // which sits lowest: the LAST one added ends up outermost, hard against the bottom
            // edge, and earlier ones stack upwards from it. So the footer reads, top to bottom,
            // the update error (only when there is one), then the action, then the version —
            // the action next to the version it acts on, with the transient notice floating
            // above the pair rather than pushing them apart.
            //
            // That is also why the quick-action buttons are built inside grpQuickActions rather
            // than added to this collection at runtime: appending to this.Controls after the
            // footer would silently reorder it.
            this.Controls.Add(this.lnkUpdateError);
            this.Controls.Add(this.lnkCheckUpdates);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.grpQuickActions);
            this.Controls.Add(this.grpInstruction);
            this.Controls.Add(this.lblStatus);
            this.Name = "AITaskPane";
            this.Size = new System.Drawing.Size(260, 500);
            this.flowQuickActions.ResumeLayout(false);
            this.grpQuickActions.ResumeLayout(false);
            this.grpInstruction.ResumeLayout(false);
            this.grpInstruction.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.GroupBox grpQuickActions;
        private System.Windows.Forms.FlowLayoutPanel flowQuickActions;
        private System.Windows.Forms.Button btnSelectSignature;
        private System.Windows.Forms.GroupBox grpInstruction;
        private System.Windows.Forms.TextBox txtPrompt;
        private System.Windows.Forms.Button btnDraft;
        private System.Windows.Forms.Button btnEditDraft;
        private System.Windows.Forms.Button btnEditSelection;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.LinkLabel lnkUpdateError;
        private System.Windows.Forms.LinkLabel lnkCheckUpdates;
    }
}
