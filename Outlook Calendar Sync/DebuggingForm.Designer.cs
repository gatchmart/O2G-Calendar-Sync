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
            this.SuspendLayout();
            // 
            // Data_RTB
            // 
            this.Data_RTB.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Data_RTB.BackColor = System.Drawing.SystemColors.Window;
            this.Data_RTB.Location = new System.Drawing.Point(12, 41);
            this.Data_RTB.Name = "Data_RTB";
            this.Data_RTB.ReadOnly = true;
            this.Data_RTB.Size = new System.Drawing.Size(1003, 657);
            this.Data_RTB.TabIndex = 0;
            this.Data_RTB.Text = "";
            // 
            // FileSelect_CB
            // 
            this.FileSelect_CB.FormattingEnabled = true;
            this.FileSelect_CB.Location = new System.Drawing.Point(12, 12);
            this.FileSelect_CB.Name = "FileSelect_CB";
            this.FileSelect_CB.Size = new System.Drawing.Size(244, 21);
            this.FileSelect_CB.TabIndex = 1;
            // 
            // Load_BTN
            // 
            this.Load_BTN.Location = new System.Drawing.Point(262, 10);
            this.Load_BTN.Name = "Load_BTN";
            this.Load_BTN.Size = new System.Drawing.Size(75, 23);
            this.Load_BTN.TabIndex = 2;
            this.Load_BTN.Text = "Load File";
            this.Load_BTN.UseVisualStyleBackColor = true;
            this.Load_BTN.Click += new System.EventHandler(this.Load_BTN_Click);
            // 
            // DebuggingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1027, 710);
            this.Controls.Add(this.Load_BTN);
            this.Controls.Add(this.FileSelect_CB);
            this.Controls.Add(this.Data_RTB);
            this.Name = "DebuggingForm";
            this.Text = "Debugging";
            this.Load += new System.EventHandler(this.DebuggingForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox Data_RTB;
        private System.Windows.Forms.ComboBox FileSelect_CB;
        private System.Windows.Forms.Button Load_BTN;
    }
}