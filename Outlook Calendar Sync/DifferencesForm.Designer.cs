namespace Outlook_Calendar_Sync {
    partial class DifferencesForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing ) {
            if ( disposing && ( components != null ) ) {
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Outlook_LV = new System.Windows.Forms.ListBox();
            this.Google_LV = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.outlook_BTN = new System.Windows.Forms.Button();
            this.Google_BTN = new System.Windows.Forms.Button();
            this.Ignore_BTN = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Outlook_LV);
            this.groupBox1.Location = new System.Drawing.Point(10, 46);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(345, 308);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Outlook\'s Appointment";
            // 
            // Outlook_LV
            // 
            this.Outlook_LV.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Outlook_LV.FormattingEnabled = true;
            this.Outlook_LV.Location = new System.Drawing.Point(5, 18);
            this.Outlook_LV.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Outlook_LV.Name = "Outlook_LV";
            this.Outlook_LV.Size = new System.Drawing.Size(336, 277);
            this.Outlook_LV.TabIndex = 0;
            // 
            // Google_LV
            // 
            this.Google_LV.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Google_LV.FormattingEnabled = true;
            this.Google_LV.Location = new System.Drawing.Point(5, 18);
            this.Google_LV.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Google_LV.Name = "Google_LV";
            this.Google_LV.Size = new System.Drawing.Size(336, 277);
            this.Google_LV.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.Google_LV);
            this.groupBox2.Location = new System.Drawing.Point(359, 46);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(345, 308);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Google\'s Appointment";
            // 
            // outlook_BTN
            // 
            this.outlook_BTN.Location = new System.Drawing.Point(10, 359);
            this.outlook_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.outlook_BTN.Name = "outlook_BTN";
            this.outlook_BTN.Size = new System.Drawing.Size(137, 27);
            this.outlook_BTN.TabIndex = 2;
            this.outlook_BTN.Text = "Keep Outlook\'s Version";
            this.outlook_BTN.UseVisualStyleBackColor = true;
            this.outlook_BTN.Click += new System.EventHandler(this.outlook_BTN_Click);
            // 
            // Google_BTN
            // 
            this.Google_BTN.Location = new System.Drawing.Point(567, 360);
            this.Google_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Google_BTN.Name = "Google_BTN";
            this.Google_BTN.Size = new System.Drawing.Size(137, 27);
            this.Google_BTN.TabIndex = 3;
            this.Google_BTN.Text = "Keep Google\'s Version";
            this.Google_BTN.UseVisualStyleBackColor = true;
            this.Google_BTN.Click += new System.EventHandler(this.Google_BTN_Click);
            // 
            // Ignore_BTN
            // 
            this.Ignore_BTN.Location = new System.Drawing.Point(289, 360);
            this.Ignore_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Ignore_BTN.Name = "Ignore_BTN";
            this.Ignore_BTN.Size = new System.Drawing.Size(137, 27);
            this.Ignore_BTN.TabIndex = 4;
            this.Ignore_BTN.Text = "Ignore Changes";
            this.Ignore_BTN.UseVisualStyleBackColor = true;
            this.Ignore_BTN.Click += new System.EventHandler(this.Ignore_BTN_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(9, 7);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(476, 37);
            this.label1.TabIndex = 5;
            this.label1.Text = "Two events with the same ID are different. Please review the differences and sele" +
    "ct which event\'s version to keep or ignore the changes.";
            // 
            // DifferencesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(713, 398);
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Ignore_BTN);
            this.Controls.Add(this.Google_BTN);
            this.Controls.Add(this.outlook_BTN);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "DifferencesForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Two Events are Different.";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox Outlook_LV;
        private System.Windows.Forms.ListBox Google_LV;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button outlook_BTN;
        private System.Windows.Forms.Button Google_BTN;
        private System.Windows.Forms.Button Ignore_BTN;
        private System.Windows.Forms.Label label1;
    }
}