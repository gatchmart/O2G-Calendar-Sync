namespace Outlook_Calendar_Sync {
    partial class SyncerForm {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SyncerForm));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.End_DTP = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.Start_DTP = new System.Windows.Forms.DateTimePicker();
            this.Sync_BTN = new System.Windows.Forms.Button();
            this.calendarUpdate_WORKER = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.googleCal_CB = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.outlookCal_CB = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.End_DTP);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.Start_DTP);
            this.groupBox1.Location = new System.Drawing.Point(10, 101);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(246, 93);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Sync by Date";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(6, 19);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(2);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(140, 17);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "Sync within date range?";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "End Date:";
            // 
            // End_DTP
            // 
            this.End_DTP.CustomFormat = "MM/dd/yyyy";
            this.End_DTP.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.End_DTP.Location = new System.Drawing.Point(70, 66);
            this.End_DTP.Name = "End_DTP";
            this.End_DTP.Size = new System.Drawing.Size(97, 20);
            this.End_DTP.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Start Date:";
            // 
            // Start_DTP
            // 
            this.Start_DTP.CustomFormat = "MM/dd/yyyy";
            this.Start_DTP.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.Start_DTP.Location = new System.Drawing.Point(70, 41);
            this.Start_DTP.Name = "Start_DTP";
            this.Start_DTP.Size = new System.Drawing.Size(97, 20);
            this.Start_DTP.TabIndex = 0;
            // 
            // Sync_BTN
            // 
            this.Sync_BTN.Location = new System.Drawing.Point(10, 201);
            this.Sync_BTN.Name = "Sync_BTN";
            this.Sync_BTN.Size = new System.Drawing.Size(75, 23);
            this.Sync_BTN.TabIndex = 4;
            this.Sync_BTN.Text = "Start Sync";
            this.Sync_BTN.UseVisualStyleBackColor = true;
            this.Sync_BTN.Click += new System.EventHandler(this.Sync_BTN_Click);
            // 
            // calendarUpdate_WORKER
            // 
            this.calendarUpdate_WORKER.WorkerReportsProgress = true;
            this.calendarUpdate_WORKER.DoWork += new System.ComponentModel.DoWorkEventHandler(this.calendarUpdate_WORKER_DoWork);
            this.calendarUpdate_WORKER.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.calendarUpdate_WORKER_ProgressChanged);
            this.calendarUpdate_WORKER.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.calendarUpdate_WORKER_RunWorkerCompleted);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(8, 236);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(2);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(248, 19);
            this.progressBar1.TabIndex = 5;
            // 
            // googleCal_CB
            // 
            this.googleCal_CB.FormattingEnabled = true;
            this.googleCal_CB.Location = new System.Drawing.Point(112, 51);
            this.googleCal_CB.Margin = new System.Windows.Forms.Padding(2);
            this.googleCal_CB.Name = "googleCal_CB";
            this.googleCal_CB.Size = new System.Drawing.Size(144, 21);
            this.googleCal_CB.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(10, 11);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(250, 38);
            this.label3.TabIndex = 7;
            this.label3.Text = "Please select which Google and Outlook Calendar to sync with each other.\r\n";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 54);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Google Calendars:";
            // 
            // outlookCal_CB
            // 
            this.outlookCal_CB.FormattingEnabled = true;
            this.outlookCal_CB.Location = new System.Drawing.Point(112, 76);
            this.outlookCal_CB.Margin = new System.Windows.Forms.Padding(2);
            this.outlookCal_CB.Name = "outlookCal_CB";
            this.outlookCal_CB.Size = new System.Drawing.Size(144, 21);
            this.outlookCal_CB.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 78);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(97, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Outlook Calendars:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(200, 203);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(56, 19);
            this.button1.TabIndex = 11;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // SyncerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(278, 267);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.Sync_BTN);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.outlookCal_CB);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.googleCal_CB);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SyncerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "O2G Calendar Sync";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SyncerForm_FormClosing);
            this.Load += new System.EventHandler(this.SyncerForm_Load);
            this.Shown += new System.EventHandler(this.SyncerForm_Shown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker Start_DTP;
        private System.Windows.Forms.Button Sync_BTN;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker End_DTP;
        private System.ComponentModel.BackgroundWorker calendarUpdate_WORKER;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox googleCal_CB;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox outlookCal_CB;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button button1;
    }
}