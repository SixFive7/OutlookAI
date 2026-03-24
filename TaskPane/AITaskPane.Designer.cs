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
                components?.Dispose();
                DisposeCustomResources();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpQuickActions = new System.Windows.Forms.GroupBox();
            this.btnProofread = new System.Windows.Forms.Button();
            this.btnRevise = new System.Windows.Forms.Button();
            this.btnShorten = new System.Windows.Forms.Button();
            this.btnLengthen = new System.Windows.Forms.Button();
            this.btnFormal = new System.Windows.Forms.Button();
            this.btnFriendly = new System.Windows.Forms.Button();
            this.grpInstruction = new System.Windows.Forms.GroupBox();
            this.txtPrompt = new System.Windows.Forms.TextBox();
            this.btnDraft = new System.Windows.Forms.Button();
            this.btnEditDraft = new System.Windows.Forms.Button();
            this.btnEditSelection = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lnkUpdateError = new System.Windows.Forms.LinkLabel();
            this.grpQuickActions.SuspendLayout();
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
            this.grpQuickActions.Controls.Add(this.btnProofread);
            this.grpQuickActions.Controls.Add(this.btnRevise);
            this.grpQuickActions.Controls.Add(this.btnShorten);
            this.grpQuickActions.Controls.Add(this.btnLengthen);
            this.grpQuickActions.Controls.Add(this.btnFormal);
            this.grpQuickActions.Controls.Add(this.btnFriendly);
            this.grpQuickActions.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpQuickActions.Location = new System.Drawing.Point(10, 40);
            this.grpQuickActions.Name = "grpQuickActions";
            this.grpQuickActions.Size = new System.Drawing.Size(240, 95);
            this.grpQuickActions.TabIndex = 0;
            this.grpQuickActions.TabStop = false;
            this.grpQuickActions.Text = "Quick Actions (Edit Current Email)";

            // btnProofread
            this.btnProofread.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnProofread.Location = new System.Drawing.Point(10, 22);
            this.btnProofread.Size = new System.Drawing.Size(70, 28);
            this.btnProofread.Text = "Proofread";
            this.btnProofread.UseVisualStyleBackColor = true;
            this.btnProofread.Click += new System.EventHandler(this.btnProofread_Click);

            // btnRevise
            this.btnRevise.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnRevise.Location = new System.Drawing.Point(85, 22);
            this.btnRevise.Size = new System.Drawing.Size(70, 28);
            this.btnRevise.Text = "Revise";
            this.btnRevise.UseVisualStyleBackColor = true;
            this.btnRevise.Click += new System.EventHandler(this.btnRevise_Click);

            // btnShorten
            this.btnShorten.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnShorten.Location = new System.Drawing.Point(160, 22);
            this.btnShorten.Size = new System.Drawing.Size(70, 28);
            this.btnShorten.Text = "Shorten";
            this.btnShorten.UseVisualStyleBackColor = true;
            this.btnShorten.Click += new System.EventHandler(this.btnShorten_Click);

            // btnLengthen
            this.btnLengthen.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnLengthen.Location = new System.Drawing.Point(10, 55);
            this.btnLengthen.Size = new System.Drawing.Size(70, 28);
            this.btnLengthen.Text = "Lengthen";
            this.btnLengthen.UseVisualStyleBackColor = true;
            this.btnLengthen.Click += new System.EventHandler(this.btnLengthen_Click);

            // btnFormal
            this.btnFormal.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnFormal.Location = new System.Drawing.Point(85, 55);
            this.btnFormal.Size = new System.Drawing.Size(70, 28);
            this.btnFormal.Text = "Formal";
            this.btnFormal.UseVisualStyleBackColor = true;
            this.btnFormal.Click += new System.EventHandler(this.btnFormal_Click);

            // btnFriendly
            this.btnFriendly.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnFriendly.Location = new System.Drawing.Point(160, 55);
            this.btnFriendly.Size = new System.Drawing.Size(70, 28);
            this.btnFriendly.Text = "Friendly";
            this.btnFriendly.UseVisualStyleBackColor = true;
            this.btnFriendly.Click += new System.EventHandler(this.btnFriendly_Click);

            // grpInstruction
            this.grpInstruction.Controls.Add(this.txtPrompt);
            this.grpInstruction.Controls.Add(this.btnDraft);
            this.grpInstruction.Controls.Add(this.btnEditDraft);
            this.grpInstruction.Controls.Add(this.btnEditSelection);
            this.grpInstruction.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpInstruction.Location = new System.Drawing.Point(10, 140);
            this.grpInstruction.Name = "grpInstruction";
            this.grpInstruction.Size = new System.Drawing.Size(240, 170);
            this.grpInstruction.TabIndex = 1;
            this.grpInstruction.TabStop = false;
            this.grpInstruction.Text = "Instruction";

            // txtPrompt
            this.txtPrompt.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.txtPrompt.Location = new System.Drawing.Point(10, 22);
            this.txtPrompt.Multiline = true;
            this.txtPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtPrompt.Size = new System.Drawing.Size(220, 50);

            // btnDraft
            this.btnDraft.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDraft.Location = new System.Drawing.Point(10, 78);
            this.btnDraft.Size = new System.Drawing.Size(220, 26);
            this.btnDraft.Text = "Draft new email";
            this.btnDraft.UseVisualStyleBackColor = true;
            this.btnDraft.Click += new System.EventHandler(this.btnDraft_Click);

            // btnEditDraft
            this.btnEditDraft.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnEditDraft.Location = new System.Drawing.Point(10, 108);
            this.btnEditDraft.Size = new System.Drawing.Size(220, 26);
            this.btnEditDraft.Text = "Edit current draft";
            this.btnEditDraft.UseVisualStyleBackColor = true;
            this.btnEditDraft.Click += new System.EventHandler(this.btnEditDraft_Click);

            // btnEditSelection
            this.btnEditSelection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnEditSelection.Location = new System.Drawing.Point(10, 138);
            this.btnEditSelection.Size = new System.Drawing.Size(220, 26);
            this.btnEditSelection.Text = "Edit selection only";
            this.btnEditSelection.UseVisualStyleBackColor = true;
            this.btnEditSelection.Click += new System.EventHandler(this.btnEditSelection_Click);

            // lblStatus
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblStatus.Location = new System.Drawing.Point(10, 315);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(240, 20);
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
            this.Controls.Add(this.lnkUpdateError);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.grpQuickActions);
            this.Controls.Add(this.grpInstruction);
            this.Controls.Add(this.lblStatus);
            this.Name = "AITaskPane";
            this.Size = new System.Drawing.Size(260, 500);
            this.grpQuickActions.ResumeLayout(false);
            this.grpInstruction.ResumeLayout(false);
            this.grpInstruction.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.GroupBox grpQuickActions;
        private System.Windows.Forms.Button btnProofread;
        private System.Windows.Forms.Button btnRevise;
        private System.Windows.Forms.Button btnShorten;
        private System.Windows.Forms.Button btnLengthen;
        private System.Windows.Forms.Button btnFormal;
        private System.Windows.Forms.Button btnFriendly;
        private System.Windows.Forms.GroupBox grpInstruction;
        private System.Windows.Forms.TextBox txtPrompt;
        private System.Windows.Forms.Button btnDraft;
        private System.Windows.Forms.Button btnEditDraft;
        private System.Windows.Forms.Button btnEditSelection;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.LinkLabel lnkUpdateError;
    }
}
