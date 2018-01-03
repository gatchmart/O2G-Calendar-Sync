namespace Outlook_Calendar_Sync {
    partial class DebuggingForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing ) {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.Data_RTB = new System.Windows.Forms.RichTextBox();
            this.FileSelect_CB = new System.Windows.Forms.ComboBox();
            this.Load_BTN = new System.Windows.Forms.Button();
            this.Compare_SC = new System.Windows.Forms.SplitContainer();
            this.FileSelect2_CB = new System.Windows.Forms.ComboBox();
            this.Data2_RTB = new System.Windows.Forms.RichTextBox();
            this.Load2_BTN = new System.Windows.Forms.Button();
            this.Compare_CB = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.Compare_SC)).BeginInit();
            this.Compare_SC.Panel1.SuspendLayout();
            this.Compare_SC.Panel2.SuspendLayout();
            this.Compare_SC.SuspendLayout();
            this.SuspendLayout();
            // 
            // Data_RTB
            // 
            this.Data_RTB.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Data_RTB.BackColor = System.Drawing.SystemColors.Window;
            this.Data_RTB.Location = new System.Drawing.Point(4, 37);
            this.Data_RTB.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Data_RTB.Name = "Data_RTB";
            this.Data_RTB.ReadOnly = true;
            this.Data_RTB.Size = new System.Drawing.Size(1337, 606);
            this.Data_RTB.TabIndex = 0;
            this.Data_RTB.Text = "";
            // 
            // FileSelect_CB
            // 
            this.FileSelect_CB.FormattingEnabled = true;
            this.FileSelect_CB.Location = new System.Drawing.Point(4, 4);
            this.FileSelect_CB.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileSelect_CB.Name = "FileSelect_CB";
            this.FileSelect_CB.Size = new System.Drawing.Size(324, 24);
            this.FileSelect_CB.TabIndex = 1;
            // 
            // Load_BTN
            // 
            this.Load_BTN.Location = new System.Drawing.Point(336, 1);
            this.Load_BTN.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Load_BTN.Name = "Load_BTN";
            this.Load_BTN.Size = new System.Drawing.Size(100, 28);
            this.Load_BTN.TabIndex = 2;
            this.Load_BTN.Text = "Load File";
            this.Load_BTN.UseVisualStyleBackColor = true;
            this.Load_BTN.Click += new System.EventHandler(this.Load_BTN_Click);
            // 
            // Compare_SC
            // 
            this.Compare_SC.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Compare_SC.Location = new System.Drawing.Point(12, 12);
            this.Compare_SC.Name = "Compare_SC";
            // 
            // Compare_SC.Panel1
            // 
            this.Compare_SC.Panel1.Controls.Add(this.Compare_CB);
            this.Compare_SC.Panel1.Controls.Add(this.FileSelect_CB);
            this.Compare_SC.Panel1.Controls.Add(this.Data_RTB);
            this.Compare_SC.Panel1.Controls.Add(this.Load_BTN);
            // 
            // Compare_SC.Panel2
            // 
            this.Compare_SC.Panel2.Controls.Add(this.FileSelect2_CB);
            this.Compare_SC.Panel2.Controls.Add(this.Data2_RTB);
            this.Compare_SC.Panel2.Controls.Add(this.Load2_BTN);
            this.Compare_SC.Panel2Collapsed = true;
            this.Compare_SC.Size = new System.Drawing.Size(1345, 647);
            this.Compare_SC.SplitterDistance = 630;
            this.Compare_SC.TabIndex = 3;
            // 
            // FileSelect2_CB
            // 
            this.FileSelect2_CB.FormattingEnabled = true;
            this.FileSelect2_CB.Location = new System.Drawing.Point(4, 4);
            this.FileSelect2_CB.Margin = new System.Windows.Forms.Padding(4);
            this.FileSelect2_CB.Name = "FileSelect2_CB";
            this.FileSelect2_CB.Size = new System.Drawing.Size(324, 24);
            this.FileSelect2_CB.TabIndex = 4;
            // 
            // Data2_RTB
            // 
            this.Data2_RTB.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Data2_RTB.BackColor = System.Drawing.SystemColors.Window;
            this.Data2_RTB.Location = new System.Drawing.Point(4, 37);
            this.Data2_RTB.Margin = new System.Windows.Forms.Padding(4);
            this.Data2_RTB.Name = "Data2_RTB";
            this.Data2_RTB.ReadOnly = true;
            this.Data2_RTB.Size = new System.Drawing.Size(703, 606);
            this.Data2_RTB.TabIndex = 3;
            this.Data2_RTB.Text = "";
            // 
            // Load2_BTN
            // 
            this.Load2_BTN.Location = new System.Drawing.Point(336, 1);
            this.Load2_BTN.Margin = new System.Windows.Forms.Padding(4);
            this.Load2_BTN.Name = "Load2_BTN";
            this.Load2_BTN.Size = new System.Drawing.Size(100, 28);
            this.Load2_BTN.TabIndex = 5;
            this.Load2_BTN.Text = "Load File";
            this.Load2_BTN.UseVisualStyleBackColor = true;
            this.Load2_BTN.Click += new System.EventHandler(this.Load2_BTN_Click);
            // 
            // Compare_CB
            // 
            this.Compare_CB.AutoSize = true;
            this.Compare_CB.Location = new System.Drawing.Point(443, 6);
            this.Compare_CB.Name = "Compare_CB";
            this.Compare_CB.Size = new System.Drawing.Size(87, 21);
            this.Compare_CB.TabIndex = 3;
            this.Compare_CB.Text = "Compare";
            this.Compare_CB.UseVisualStyleBackColor = true;
            this.Compare_CB.CheckedChanged += new System.EventHandler(this.Compare_CB_CheckedChanged);
            // 
            // DebuggingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1369, 671);
            this.Controls.Add(this.Compare_SC);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "DebuggingForm";
            this.Text = "Debugging";
            this.Load += new System.EventHandler(this.DebuggingForm_Load);
            this.Compare_SC.Panel1.ResumeLayout(false);
            this.Compare_SC.Panel1.PerformLayout();
            this.Compare_SC.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Compare_SC)).EndInit();
            this.Compare_SC.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox Data_RTB;
        private System.Windows.Forms.ComboBox FileSelect_CB;
        private System.Windows.Forms.Button Load_BTN;
        private System.Windows.Forms.SplitContainer Compare_SC;
        private System.Windows.Forms.ComboBox FileSelect2_CB;
        private System.Windows.Forms.RichTextBox Data2_RTB;
        private System.Windows.Forms.Button Load2_BTN;
        private System.Windows.Forms.CheckBox Compare_CB;
    }
}