namespace Outlook_Calendar_Sync
{
    partial class InitialLoadForm
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
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.Connect_LBL = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.Select_LBL = new System.Windows.Forms.Label();
            this.Initial_LBL = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.Done_LBL = new System.Windows.Forms.Label();
            this.Connect_GB = new System.Windows.Forms.GroupBox();
            this.Connected_LBL = new System.Windows.Forms.Label();
            this.Connect_BTN = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.Initial_GB = new System.Windows.Forms.GroupBox();
            this.Cancel_BTN = new System.Windows.Forms.Button();
            this.Start_BTN = new System.Windows.Forms.Button();
            this.Status_TB = new System.Windows.Forms.RichTextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label10 = new System.Windows.Forms.Label();
            this.Next_BTN = new System.Windows.Forms.Button();
            this.Select_GB = new System.Windows.Forms.GroupBox();
            this.Remove_BTN = new System.Windows.Forms.Button();
            this.Add_BTN = new System.Windows.Forms.Button();
            this.OutlookCal_CB = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.GoogleCal_CB = new System.Windows.Forms.ComboBox();
            this.Pair_LB = new System.Windows.Forms.ListBox();
            this.label6 = new System.Windows.Forms.Label();
            this.Previous_BTN = new System.Windows.Forms.Button();
            this.InitialSyncer_BW = new System.ComponentModel.BackgroundWorker();
            this.Done_GB = new System.Windows.Forms.GroupBox();
            this.Close_BTN = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.Connect_GB.SuspendLayout();
            this.Initial_GB.SuspendLayout();
            this.Select_GB.SuspendLayout();
            this.Done_GB.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(438, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Welcome to Outlook - Google Calendar Sync";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Outlook_Calendar_Sync.Properties.Resources.synchronization_arrows;
            this.pictureBox1.Location = new System.Drawing.Point(456, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(96, 96);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // Connect_LBL
            // 
            this.Connect_LBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Connect_LBL.Location = new System.Drawing.Point(17, 41);
            this.Connect_LBL.Name = "Connect_LBL";
            this.Connect_LBL.Size = new System.Drawing.Size(113, 23);
            this.Connect_LBL.TabIndex = 2;
            this.Connect_LBL.Text = "Connect to Google";
            this.Connect_LBL.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(128, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(27, 19);
            this.label2.TabIndex = 3;
            this.label2.Text = "->";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(263, 38);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 19);
            this.label3.TabIndex = 5;
            this.label3.Text = "->";
            // 
            // Select_LBL
            // 
            this.Select_LBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Select_LBL.Location = new System.Drawing.Point(149, 41);
            this.Select_LBL.Name = "Select_LBL";
            this.Select_LBL.Size = new System.Drawing.Size(125, 23);
            this.Select_LBL.TabIndex = 4;
            this.Select_LBL.Text = "Select Calendars";
            this.Select_LBL.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Initial_LBL
            // 
            this.Initial_LBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Initial_LBL.Location = new System.Drawing.Point(296, 41);
            this.Initial_LBL.Name = "Initial_LBL";
            this.Initial_LBL.Size = new System.Drawing.Size(70, 23);
            this.Initial_LBL.TabIndex = 6;
            this.Initial_LBL.Text = "Initial Sync";
            this.Initial_LBL.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(372, 38);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(27, 19);
            this.label5.TabIndex = 7;
            this.label5.Text = "->";
            // 
            // Done_LBL
            // 
            this.Done_LBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Done_LBL.Location = new System.Drawing.Point(405, 41);
            this.Done_LBL.Name = "Done_LBL";
            this.Done_LBL.Size = new System.Drawing.Size(43, 23);
            this.Done_LBL.TabIndex = 8;
            this.Done_LBL.Text = "Done!";
            this.Done_LBL.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Connect_GB
            // 
            this.Connect_GB.Controls.Add(this.Connected_LBL);
            this.Connect_GB.Controls.Add(this.Connect_BTN);
            this.Connect_GB.Controls.Add(this.label7);
            this.Connect_GB.Location = new System.Drawing.Point(13, 67);
            this.Connect_GB.Name = "Connect_GB";
            this.Connect_GB.Size = new System.Drawing.Size(437, 254);
            this.Connect_GB.TabIndex = 9;
            this.Connect_GB.TabStop = false;
            this.Connect_GB.Text = " ";
            // 
            // Connected_LBL
            // 
            this.Connected_LBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Connected_LBL.Location = new System.Drawing.Point(135, 134);
            this.Connected_LBL.Name = "Connected_LBL";
            this.Connected_LBL.Size = new System.Drawing.Size(156, 23);
            this.Connected_LBL.TabIndex = 2;
            this.Connected_LBL.Text = "Connected!!";
            this.Connected_LBL.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.Connected_LBL.Visible = false;
            // 
            // Connect_BTN
            // 
            this.Connect_BTN.Location = new System.Drawing.Point(151, 60);
            this.Connect_BTN.Name = "Connect_BTN";
            this.Connect_BTN.Size = new System.Drawing.Size(117, 30);
            this.Connect_BTN.TabIndex = 1;
            this.Connect_BTN.Text = "Connect to Google";
            this.Connect_BTN.UseVisualStyleBackColor = true;
            this.Connect_BTN.Click += new System.EventHandler(this.Connect_BTN_Click);
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(6, 17);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(420, 23);
            this.label7.TabIndex = 0;
            this.label7.Text = "First thing we need to do is to connect this add-in to Google\'s Calender.";
            this.label7.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Initial_GB
            // 
            this.Initial_GB.Controls.Add(this.Cancel_BTN);
            this.Initial_GB.Controls.Add(this.Start_BTN);
            this.Initial_GB.Controls.Add(this.Status_TB);
            this.Initial_GB.Controls.Add(this.progressBar1);
            this.Initial_GB.Controls.Add(this.label10);
            this.Initial_GB.Location = new System.Drawing.Point(579, 352);
            this.Initial_GB.Name = "Initial_GB";
            this.Initial_GB.Size = new System.Drawing.Size(437, 254);
            this.Initial_GB.TabIndex = 13;
            this.Initial_GB.TabStop = false;
            this.Initial_GB.Text = " ";
            this.Initial_GB.Visible = false;
            // 
            // Cancel_BTN
            // 
            this.Cancel_BTN.Enabled = false;
            this.Cancel_BTN.Location = new System.Drawing.Point(349, 216);
            this.Cancel_BTN.Name = "Cancel_BTN";
            this.Cancel_BTN.Size = new System.Drawing.Size(75, 23);
            this.Cancel_BTN.TabIndex = 4;
            this.Cancel_BTN.Text = "Cancel";
            this.Cancel_BTN.UseVisualStyleBackColor = true;
            this.Cancel_BTN.Click += new System.EventHandler(this.Cancel_BTN_Click);
            // 
            // Start_BTN
            // 
            this.Start_BTN.Location = new System.Drawing.Point(9, 215);
            this.Start_BTN.Name = "Start_BTN";
            this.Start_BTN.Size = new System.Drawing.Size(75, 23);
            this.Start_BTN.TabIndex = 3;
            this.Start_BTN.Text = "Start";
            this.Start_BTN.UseVisualStyleBackColor = true;
            this.Start_BTN.Click += new System.EventHandler(this.Start_BTN_Click);
            // 
            // Status_TB
            // 
            this.Status_TB.Location = new System.Drawing.Point(9, 74);
            this.Status_TB.Name = "Status_TB";
            this.Status_TB.ReadOnly = true;
            this.Status_TB.Size = new System.Drawing.Size(415, 117);
            this.Status_TB.TabIndex = 2;
            this.Status_TB.Text = "";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(9, 44);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(415, 23);
            this.progressBar1.TabIndex = 1;
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(6, 17);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(420, 23);
            this.label10.TabIndex = 0;
            this.label10.Text = "Now it\'s time to perform the initial sync of the calendar pairs you\'ve created.";
            this.label10.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Next_BTN
            // 
            this.Next_BTN.Enabled = false;
            this.Next_BTN.Location = new System.Drawing.Point(456, 269);
            this.Next_BTN.Name = "Next_BTN";
            this.Next_BTN.Size = new System.Drawing.Size(96, 23);
            this.Next_BTN.TabIndex = 10;
            this.Next_BTN.Text = "Next";
            this.Next_BTN.UseVisualStyleBackColor = true;
            this.Next_BTN.Click += new System.EventHandler(this.Next_BTN_Click);
            // 
            // Select_GB
            // 
            this.Select_GB.Controls.Add(this.Remove_BTN);
            this.Select_GB.Controls.Add(this.Add_BTN);
            this.Select_GB.Controls.Add(this.OutlookCal_CB);
            this.Select_GB.Controls.Add(this.label8);
            this.Select_GB.Controls.Add(this.label4);
            this.Select_GB.Controls.Add(this.GoogleCal_CB);
            this.Select_GB.Controls.Add(this.Pair_LB);
            this.Select_GB.Controls.Add(this.label6);
            this.Select_GB.Location = new System.Drawing.Point(579, 67);
            this.Select_GB.Name = "Select_GB";
            this.Select_GB.Size = new System.Drawing.Size(437, 254);
            this.Select_GB.TabIndex = 11;
            this.Select_GB.TabStop = false;
            // 
            // Remove_BTN
            // 
            this.Remove_BTN.Location = new System.Drawing.Point(320, 216);
            this.Remove_BTN.Name = "Remove_BTN";
            this.Remove_BTN.Size = new System.Drawing.Size(106, 23);
            this.Remove_BTN.TabIndex = 7;
            this.Remove_BTN.Text = "Remove Pair";
            this.Remove_BTN.UseVisualStyleBackColor = true;
            this.Remove_BTN.Click += new System.EventHandler(this.Remove_BTN_Click);
            // 
            // Add_BTN
            // 
            this.Add_BTN.Location = new System.Drawing.Point(320, 187);
            this.Add_BTN.Name = "Add_BTN";
            this.Add_BTN.Size = new System.Drawing.Size(106, 23);
            this.Add_BTN.TabIndex = 6;
            this.Add_BTN.Text = "Add Pair";
            this.Add_BTN.UseVisualStyleBackColor = true;
            this.Add_BTN.Click += new System.EventHandler(this.Add_BTN_Click);
            // 
            // OutlookCal_CB
            // 
            this.OutlookCal_CB.FormattingEnabled = true;
            this.OutlookCal_CB.Location = new System.Drawing.Point(111, 218);
            this.OutlookCal_CB.Name = "OutlookCal_CB";
            this.OutlookCal_CB.Size = new System.Drawing.Size(152, 21);
            this.OutlookCal_CB.TabIndex = 5;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 221);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(97, 13);
            this.label8.TabIndex = 4;
            this.label8.Text = "Outlook Calendars:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 194);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Google Calenders:";
            // 
            // GoogleCal_CB
            // 
            this.GoogleCal_CB.FormattingEnabled = true;
            this.GoogleCal_CB.Location = new System.Drawing.Point(111, 191);
            this.GoogleCal_CB.Name = "GoogleCal_CB";
            this.GoogleCal_CB.Size = new System.Drawing.Size(152, 21);
            this.GoogleCal_CB.TabIndex = 2;
            // 
            // Pair_LB
            // 
            this.Pair_LB.FormattingEnabled = true;
            this.Pair_LB.Location = new System.Drawing.Point(9, 44);
            this.Pair_LB.Name = "Pair_LB";
            this.Pair_LB.Size = new System.Drawing.Size(417, 134);
            this.Pair_LB.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(6, 17);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(420, 23);
            this.label6.TabIndex = 0;
            this.label6.Text = "Next you\'ll need to add pairs of calendars to sync.";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Previous_BTN
            // 
            this.Previous_BTN.Enabled = false;
            this.Previous_BTN.Location = new System.Drawing.Point(456, 298);
            this.Previous_BTN.Name = "Previous_BTN";
            this.Previous_BTN.Size = new System.Drawing.Size(96, 23);
            this.Previous_BTN.TabIndex = 12;
            this.Previous_BTN.Text = "Back";
            this.Previous_BTN.UseVisualStyleBackColor = true;
            this.Previous_BTN.Click += new System.EventHandler(this.Previous_BTN_Click);
            // 
            // InitialSyncer_BW
            // 
            this.InitialSyncer_BW.WorkerReportsProgress = true;
            this.InitialSyncer_BW.WorkerSupportsCancellation = true;
            this.InitialSyncer_BW.DoWork += new System.ComponentModel.DoWorkEventHandler(this.InitialSyncer_BW_DoWork);
            this.InitialSyncer_BW.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.InitialSyncer_BW_ProgressChanged);
            this.InitialSyncer_BW.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.InitialSyncer_BW_RunWorkerCompleted);
            // 
            // Done_GB
            // 
            this.Done_GB.Controls.Add(this.Close_BTN);
            this.Done_GB.Controls.Add(this.label11);
            this.Done_GB.Location = new System.Drawing.Point(11, 352);
            this.Done_GB.Name = "Done_GB";
            this.Done_GB.Size = new System.Drawing.Size(437, 254);
            this.Done_GB.TabIndex = 14;
            this.Done_GB.TabStop = false;
            this.Done_GB.Text = " ";
            // 
            // Close_BTN
            // 
            this.Close_BTN.Location = new System.Drawing.Point(153, 116);
            this.Close_BTN.Name = "Close_BTN";
            this.Close_BTN.Size = new System.Drawing.Size(117, 30);
            this.Close_BTN.TabIndex = 1;
            this.Close_BTN.Text = "Close";
            this.Close_BTN.UseVisualStyleBackColor = true;
            this.Close_BTN.Click += new System.EventHandler(this.Close_BTN_Click);
            // 
            // label11
            // 
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(6, 17);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(420, 38);
            this.label11.TabIndex = 0;
            this.label11.Text = "All Done!! You\'ve completed the initial setup for the Outlook Google Calendar Syn" +
    "c.\r\n";
            this.label11.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // InitialLoadForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1112, 679);
            this.Controls.Add(this.Done_GB);
            this.Controls.Add(this.Select_GB);
            this.Controls.Add(this.Connect_GB);
            this.Controls.Add(this.Previous_BTN);
            this.Controls.Add(this.Initial_GB);
            this.Controls.Add(this.Next_BTN);
            this.Controls.Add(this.Done_LBL);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Initial_LBL);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Select_LBL);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Connect_LBL);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Name = "InitialLoadForm";
            this.Text = "Initial Setup";
            this.Load += new System.EventHandler(this.InitialLoadForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.Connect_GB.ResumeLayout(false);
            this.Initial_GB.ResumeLayout(false);
            this.Select_GB.ResumeLayout(false);
            this.Select_GB.PerformLayout();
            this.Done_GB.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label Connect_LBL;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label Select_LBL;
        private System.Windows.Forms.Label Initial_LBL;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label Done_LBL;
        private System.Windows.Forms.GroupBox Connect_GB;
        private System.Windows.Forms.Button Connect_BTN;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label Connected_LBL;
        private System.Windows.Forms.Button Next_BTN;
        private System.Windows.Forms.GroupBox Select_GB;
        private System.Windows.Forms.Button Remove_BTN;
        private System.Windows.Forms.Button Add_BTN;
        private System.Windows.Forms.ComboBox OutlookCal_CB;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox GoogleCal_CB;
        private System.Windows.Forms.ListBox Pair_LB;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button Previous_BTN;
        private System.Windows.Forms.GroupBox Initial_GB;
        private System.Windows.Forms.Button Cancel_BTN;
        private System.Windows.Forms.Button Start_BTN;
        private System.Windows.Forms.RichTextBox Status_TB;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label10;
        private System.ComponentModel.BackgroundWorker InitialSyncer_BW;
        private System.Windows.Forms.GroupBox Done_GB;
        private System.Windows.Forms.Button Close_BTN;
        private System.Windows.Forms.Label label11;
    }
}