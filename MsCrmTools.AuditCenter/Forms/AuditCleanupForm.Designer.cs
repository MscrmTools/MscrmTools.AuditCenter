namespace MsCrmTools.AuditCenter.Forms
{
    partial class AuditCleanupForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblTables = new System.Windows.Forms.Label();
            this.clbTables = new System.Windows.Forms.CheckedListBox();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.lblFilter = new System.Windows.Forms.Label();
            this.lblFrom = new System.Windows.Forms.Label();
            this.dtpFrom = new System.Windows.Forms.DateTimePicker();
            this.lblTo = new System.Windows.Forms.Label();
            this.dtpTo = new System.Windows.Forms.DateTimePicker();
            this.lblUtcNote = new System.Windows.Forms.Label();
            this.chkConfirm = new System.Windows.Forms.CheckBox();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.btnPreview = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.lblWarning = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // lblWarning
            //
            this.lblWarning.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(244)))), ((int)(((byte)(214)))));
            this.lblWarning.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblWarning.Location = new System.Drawing.Point(12, 12);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Padding = new System.Windows.Forms.Padding(6);
            this.lblWarning.Size = new System.Drawing.Size(616, 56);
            this.lblWarning.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblWarning.TabIndex = 0;
            this.lblWarning.Text = "This operation submits an irreversible Dataverse Bulk Delete job that permanently " +
                "removes audit records for the selected table(s) within the specified UTC changed" +
                "-date range. It runs asynchronously as a system job.";
            //
            // lblFilter
            //
            this.lblFilter.AutoSize = true;
            this.lblFilter.Location = new System.Drawing.Point(12, 82);
            this.lblFilter.Name = "lblFilter";
            this.lblFilter.Size = new System.Drawing.Size(32, 13);
            this.lblFilter.TabIndex = 1;
            this.lblFilter.Text = "Filter";
            //
            // txtFilter
            //
            this.txtFilter.Location = new System.Drawing.Point(50, 79);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(578, 20);
            this.txtFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFilter.TabIndex = 2;
            //
            // lblTables
            //
            this.lblTables.AutoSize = true;
            this.lblTables.Location = new System.Drawing.Point(12, 108);
            this.lblTables.Name = "lblTables";
            this.lblTables.Size = new System.Drawing.Size(126, 13);
            this.lblTables.TabIndex = 3;
            this.lblTables.Text = "Audit-enabled table(s)";
            //
            // clbTables
            //
            this.clbTables.CheckOnClick = true;
            this.clbTables.FormattingEnabled = true;
            this.clbTables.IntegralHeight = false;
            this.clbTables.Location = new System.Drawing.Point(12, 124);
            this.clbTables.Name = "clbTables";
            this.clbTables.Size = new System.Drawing.Size(616, 190);
            this.clbTables.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.clbTables.TabIndex = 4;
            //
            // lblFrom
            //
            this.lblFrom.AutoSize = true;
            this.lblFrom.Location = new System.Drawing.Point(12, 330);
            this.lblFrom.Name = "lblFrom";
            this.lblFrom.Size = new System.Drawing.Size(84, 13);
            this.lblFrom.TabIndex = 5;
            this.lblFrom.Text = "From day (UTC)";
            this.lblFrom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            //
            // dtpFrom
            //
            this.dtpFrom.CustomFormat = "yyyy-MM-dd";
            this.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpFrom.Location = new System.Drawing.Point(102, 326);
            this.dtpFrom.Name = "dtpFrom";
            this.dtpFrom.Size = new System.Drawing.Size(150, 20);
            this.dtpFrom.TabIndex = 6;
            this.dtpFrom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            //
            // lblTo
            //
            this.lblTo.AutoSize = true;
            this.lblTo.Location = new System.Drawing.Point(268, 330);
            this.lblTo.Name = "lblTo";
            this.lblTo.Size = new System.Drawing.Size(122, 13);
            this.lblTo.TabIndex = 7;
            this.lblTo.Text = "To day (UTC, inclusive)";
            this.lblTo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            //
            // dtpTo
            //
            this.dtpTo.CustomFormat = "yyyy-MM-dd";
            this.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpTo.Location = new System.Drawing.Point(396, 326);
            this.dtpTo.Name = "dtpTo";
            this.dtpTo.Size = new System.Drawing.Size(150, 20);
            this.dtpTo.TabIndex = 8;
            this.dtpTo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            //
            // lblUtcNote
            //
            this.lblUtcNote.AutoSize = true;
            this.lblUtcNote.ForeColor = System.Drawing.SystemColors.GrayText;
            this.lblUtcNote.Location = new System.Drawing.Point(12, 356);
            this.lblUtcNote.Name = "lblUtcNote";
            this.lblUtcNote.Size = new System.Drawing.Size(616, 13);
            this.lblUtcNote.TabIndex = 9;
            this.lblUtcNote.Text = "Whole UTC days, matched on \'Changed Date\' (createdon): from 00:00 of the \'From\' da" +
                "y up to (but excluding) 00:00 of the day after the \'To\' day.";
            this.lblUtcNote.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            //
            // chkConfirm
            //
            this.chkConfirm.Location = new System.Drawing.Point(12, 379);
            this.chkConfirm.Name = "chkConfirm";
            this.chkConfirm.Size = new System.Drawing.Size(616, 24);
            this.chkConfirm.TabIndex = 10;
            this.chkConfirm.Text = "I understand this permanently deletes the matching audit records and cannot be un" +
                "done.";
            this.chkConfirm.UseVisualStyleBackColor = true;
            this.chkConfirm.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            //
            // lblResult
            //
            this.lblResult.ForeColor = System.Drawing.Color.DarkGreen;
            this.lblResult.Location = new System.Drawing.Point(12, 409);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(616, 96);
            this.lblResult.TabIndex = 11;
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            //
            // btnSubmit
            //
            this.btnSubmit.Location = new System.Drawing.Point(472, 515);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(156, 28);
            this.btnSubmit.TabIndex = 12;
            this.btnSubmit.Text = "Submit bulk delete job";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            this.btnSubmit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            //
            // btnPreview
            //
            this.btnPreview.Location = new System.Drawing.Point(12, 515);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(160, 28);
            this.btnPreview.TabIndex = 14;
            this.btnPreview.Text = "Preview matching records";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            this.btnPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            //
            // btnClose
            //
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(390, 515);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(76, 28);
            this.btnClose.TabIndex = 13;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            //
            // AuditCleanupForm
            //
            this.AcceptButton = this.btnSubmit;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(640, 555);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.chkConfirm);
            this.Controls.Add(this.lblUtcNote);
            this.Controls.Add(this.dtpTo);
            this.Controls.Add(this.lblTo);
            this.Controls.Add(this.dtpFrom);
            this.Controls.Add(this.lblFrom);
            this.Controls.Add(this.clbTables);
            this.Controls.Add(this.lblTables);
            this.Controls.Add(this.txtFilter);
            this.Controls.Add(this.lblFilter);
            this.Controls.Add(this.lblWarning);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(656, 594);
            this.Name = "AuditCleanupForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Audit cleanup by table and date range";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblTables;
        private System.Windows.Forms.CheckedListBox clbTables;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.Label lblFilter;
        private System.Windows.Forms.Label lblFrom;
        private System.Windows.Forms.DateTimePicker dtpFrom;
        private System.Windows.Forms.Label lblTo;
        private System.Windows.Forms.DateTimePicker dtpTo;
        private System.Windows.Forms.Label lblUtcNote;
        private System.Windows.Forms.CheckBox chkConfirm;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Label lblWarning;
    }
}
