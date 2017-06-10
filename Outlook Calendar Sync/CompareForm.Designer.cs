namespace Outlook_Calendar_Sync {
    partial class CompareForm {
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
            this.listView1 = new System.Windows.Forms.ListView();
            this.eventSubject = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.eventAction = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.checkAll_BTN = new System.Windows.Forms.Button();
            this.uncheckAll_BTN = new System.Windows.Forms.Button();
            this.cancel_BTN = new System.Windows.Forms.Button();
            this.submit_BTN = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.CheckBoxes = true;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.eventSubject,
            this.eventAction});
            this.listView1.FullRowSelect = true;
            this.listView1.GridLines = true;
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView1.Location = new System.Drawing.Point(9, 10);
            this.listView1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(354, 216);
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // eventSubject
            // 
            this.eventSubject.Text = "Event Title";
            this.eventSubject.Width = 260;
            // 
            // eventAction
            // 
            this.eventAction.Text = "Action";
            this.eventAction.Width = 200;
            // 
            // checkAll_BTN
            // 
            this.checkAll_BTN.Location = new System.Drawing.Point(10, 231);
            this.checkAll_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.checkAll_BTN.Name = "checkAll_BTN";
            this.checkAll_BTN.Size = new System.Drawing.Size(75, 21);
            this.checkAll_BTN.TabIndex = 1;
            this.checkAll_BTN.Text = "Check All";
            this.checkAll_BTN.UseVisualStyleBackColor = true;
            this.checkAll_BTN.Click += new System.EventHandler(this.checkAll_BTN_Click);
            // 
            // uncheckAll_BTN
            // 
            this.uncheckAll_BTN.Location = new System.Drawing.Point(89, 231);
            this.uncheckAll_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.uncheckAll_BTN.Name = "uncheckAll_BTN";
            this.uncheckAll_BTN.Size = new System.Drawing.Size(86, 22);
            this.uncheckAll_BTN.TabIndex = 2;
            this.uncheckAll_BTN.Text = "Uncheck All";
            this.uncheckAll_BTN.UseVisualStyleBackColor = true;
            this.uncheckAll_BTN.Click += new System.EventHandler(this.uncheckAll_BTN_Click);
            // 
            // cancel_BTN
            // 
            this.cancel_BTN.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cancel_BTN.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel_BTN.Location = new System.Drawing.Point(9, 276);
            this.cancel_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cancel_BTN.Name = "cancel_BTN";
            this.cancel_BTN.Size = new System.Drawing.Size(56, 27);
            this.cancel_BTN.TabIndex = 3;
            this.cancel_BTN.Text = "Cancel";
            this.cancel_BTN.UseVisualStyleBackColor = true;
            this.cancel_BTN.Click += new System.EventHandler(this.cancel_BTN_Click);
            // 
            // submit_BTN
            // 
            this.submit_BTN.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.submit_BTN.Location = new System.Drawing.Point(305, 276);
            this.submit_BTN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.submit_BTN.Name = "submit_BTN";
            this.submit_BTN.Size = new System.Drawing.Size(56, 27);
            this.submit_BTN.TabIndex = 4;
            this.submit_BTN.Text = "Submit";
            this.submit_BTN.UseVisualStyleBackColor = true;
            this.submit_BTN.Click += new System.EventHandler(this.submit_BTN_Click);
            // 
            // CompareForm
            // 
            this.AcceptButton = this.submit_BTN;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancel_BTN;
            this.ClientSize = new System.Drawing.Size(370, 313);
            this.ControlBox = false;
            this.Controls.Add(this.submit_BTN);
            this.Controls.Add(this.cancel_BTN);
            this.Controls.Add(this.uncheckAll_BTN);
            this.Controls.Add(this.checkAll_BTN);
            this.Controls.Add(this.listView1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "CompareForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Compare Form";
            this.Load += new System.EventHandler(this.CompareForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader eventSubject;
        private System.Windows.Forms.ColumnHeader eventAction;
        private System.Windows.Forms.Button checkAll_BTN;
        private System.Windows.Forms.Button uncheckAll_BTN;
        private System.Windows.Forms.Button cancel_BTN;
        private System.Windows.Forms.Button submit_BTN;
    }
}