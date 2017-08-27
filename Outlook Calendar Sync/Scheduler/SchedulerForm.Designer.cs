namespace Outlook_Calendar_Sync.Scheduler
{
    partial class SchedulerForm
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
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.UpdateTask_BTN = new System.Windows.Forms.Button();
            this.Save_BTN = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Precedence_CB = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.SilentSync_CB = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.Event_CB = new System.Windows.Forms.ComboBox();
            this.Time_TB = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.OutlookCal_CB = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.GoogleCal_CB = new System.Windows.Forms.ComboBox();
            this.RemoveTask_BTN = new System.Windows.Forms.Button();
            this.Calendars_LB = new System.Windows.Forms.ListBox();
            this.AddTask_BTN = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(11, 9);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(581, 22);
            this.label3.TabIndex = 12;
            this.label3.Text = "This is where you can schedule tasks to automatically sync calendar pairs.";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.UpdateTask_BTN);
            this.groupBox1.Controls.Add(this.Save_BTN);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.RemoveTask_BTN);
            this.groupBox1.Controls.Add(this.Calendars_LB);
            this.groupBox1.Controls.Add(this.AddTask_BTN);
            this.groupBox1.Location = new System.Drawing.Point(12, 34);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(580, 335);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Scheduled Tasks";
            // 
            // UpdateTask_BTN
            // 
            this.UpdateTask_BTN.Location = new System.Drawing.Point(360, 299);
            this.UpdateTask_BTN.Name = "UpdateTask_BTN";
            this.UpdateTask_BTN.Size = new System.Drawing.Size(100, 25);
            this.UpdateTask_BTN.TabIndex = 5;
            this.UpdateTask_BTN.Text = "Update Task";
            this.UpdateTask_BTN.UseVisualStyleBackColor = true;
            this.UpdateTask_BTN.Click += new System.EventHandler(this.UpdateTask_BTN_Click);
            // 
            // Save_BTN
            // 
            this.Save_BTN.Location = new System.Drawing.Point(7, 299);
            this.Save_BTN.Name = "Save_BTN";
            this.Save_BTN.Size = new System.Drawing.Size(100, 25);
            this.Save_BTN.TabIndex = 4;
            this.Save_BTN.Text = "Save Schedule";
            this.Save_BTN.UseVisualStyleBackColor = true;
            this.Save_BTN.Click += new System.EventHandler(this.Save_BTN_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.Precedence_CB);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.SilentSync_CB);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.Event_CB);
            this.groupBox2.Controls.Add(this.Time_TB);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.OutlookCal_CB);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.GoogleCal_CB);
            this.groupBox2.Location = new System.Drawing.Point(7, 148);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(564, 145);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Schedule a Task";
            // 
            // Precedence_CB
            // 
            this.Precedence_CB.Enabled = false;
            this.Precedence_CB.FormattingEnabled = true;
            this.Precedence_CB.Items.AddRange(new object[] {
            "Outlook",
            "Google"});
            this.Precedence_CB.Location = new System.Drawing.Point(391, 111);
            this.Precedence_CB.Name = "Precedence_CB";
            this.Precedence_CB.Size = new System.Drawing.Size(121, 21);
            this.Precedence_CB.TabIndex = 25;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(302, 114);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(68, 13);
            this.label9.TabIndex = 24;
            this.label9.Text = "Precedence:";
            // 
            // SilentSync_CB
            // 
            this.SilentSync_CB.AutoSize = true;
            this.SilentSync_CB.Location = new System.Drawing.Point(305, 91);
            this.SilentSync_CB.Name = "SilentSync_CB";
            this.SilentSync_CB.Size = new System.Drawing.Size(85, 17);
            this.SilentSync_CB.TabIndex = 23;
            this.SilentSync_CB.Text = "Silent Sync?";
            this.SilentSync_CB.UseVisualStyleBackColor = true;
            this.SilentSync_CB.CheckedChanged += new System.EventHandler(this.SilentSync_CB_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(288, 71);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(271, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "4) Select if you want silent syncing and the precedence.";
            // 
            // Event_CB
            // 
            this.Event_CB.FormattingEnabled = true;
            this.Event_CB.Items.AddRange(new object[] {
            "Automatically",
            "Only on Startup",
            "Hourly",
            "Daily",
            "Weekly",
            "Custom Interval"});
            this.Event_CB.Location = new System.Drawing.Point(26, 111);
            this.Event_CB.Name = "Event_CB";
            this.Event_CB.Size = new System.Drawing.Size(153, 21);
            this.Event_CB.TabIndex = 21;
            this.Event_CB.SelectedIndexChanged += new System.EventHandler(this.Event_CB_SelectedIndexChanged);
            // 
            // Time_TB
            // 
            this.Time_TB.Enabled = false;
            this.Time_TB.Location = new System.Drawing.Point(391, 38);
            this.Time_TB.Name = "Time_TB";
            this.Time_TB.Size = new System.Drawing.Size(57, 20);
            this.Time_TB.TabIndex = 20;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(302, 41);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(83, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Time in minutes:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(285, 20);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(227, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "3) Enter the interval (only for custom time span)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 95);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(211, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "2) Select how often you want them to sync.";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "1) Select the calendars to sync.";
            // 
            // OutlookCal_CB
            // 
            this.OutlookCal_CB.FormattingEnabled = true;
            this.OutlookCal_CB.Location = new System.Drawing.Point(126, 64);
            this.OutlookCal_CB.Name = "OutlookCal_CB";
            this.OutlookCal_CB.Size = new System.Drawing.Size(152, 21);
            this.OutlookCal_CB.TabIndex = 9;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(23, 67);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(97, 13);
            this.label8.TabIndex = 8;
            this.label8.Text = "Outlook Calendars:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 40);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(94, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Google Calenders:";
            // 
            // GoogleCal_CB
            // 
            this.GoogleCal_CB.FormattingEnabled = true;
            this.GoogleCal_CB.Location = new System.Drawing.Point(126, 37);
            this.GoogleCal_CB.Name = "GoogleCal_CB";
            this.GoogleCal_CB.Size = new System.Drawing.Size(152, 21);
            this.GoogleCal_CB.TabIndex = 6;
            // 
            // RemoveTask_BTN
            // 
            this.RemoveTask_BTN.Location = new System.Drawing.Point(466, 299);
            this.RemoveTask_BTN.Name = "RemoveTask_BTN";
            this.RemoveTask_BTN.Size = new System.Drawing.Size(100, 25);
            this.RemoveTask_BTN.TabIndex = 2;
            this.RemoveTask_BTN.Text = "Remove Task";
            this.RemoveTask_BTN.UseVisualStyleBackColor = true;
            this.RemoveTask_BTN.Click += new System.EventHandler(this.RemoveTask_BTN_Click);
            // 
            // Calendars_LB
            // 
            this.Calendars_LB.FormattingEnabled = true;
            this.Calendars_LB.Location = new System.Drawing.Point(7, 20);
            this.Calendars_LB.Name = "Calendars_LB";
            this.Calendars_LB.Size = new System.Drawing.Size(564, 121);
            this.Calendars_LB.TabIndex = 0;
            this.Calendars_LB.SelectedIndexChanged += new System.EventHandler(this.Calendars_LB_SelectedIndexChanged);
            // 
            // AddTask_BTN
            // 
            this.AddTask_BTN.Location = new System.Drawing.Point(254, 299);
            this.AddTask_BTN.Name = "AddTask_BTN";
            this.AddTask_BTN.Size = new System.Drawing.Size(100, 25);
            this.AddTask_BTN.TabIndex = 1;
            this.AddTask_BTN.Text = "Add Task";
            this.AddTask_BTN.UseVisualStyleBackColor = true;
            this.AddTask_BTN.Click += new System.EventHandler(this.AddTask_BTN_Click);
            // 
            // SchedulerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 378);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "SchedulerForm";
            this.Text = "Task Scheduler";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox Calendars_LB;
        private System.Windows.Forms.Button RemoveTask_BTN;
        private System.Windows.Forms.Button AddTask_BTN;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox OutlookCal_CB;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox GoogleCal_CB;
        private System.Windows.Forms.TextBox Time_TB;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox Event_CB;
        private System.Windows.Forms.Button Save_BTN;
        private System.Windows.Forms.CheckBox SilentSync_CB;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox Precedence_CB;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button UpdateTask_BTN;
    }
}