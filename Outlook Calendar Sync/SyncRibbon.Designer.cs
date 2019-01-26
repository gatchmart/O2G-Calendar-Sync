namespace Outlook_Calendar_Sync {
    partial class SyncRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SyncRibbon()
            : base( Globals.Factory.GetRibbonFactory() ) {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Debug_BTN = this.Factory.CreateRibbonButton();
            this.AboutBtn = this.Factory.CreateRibbonButton();
            this.Sync_BTN = this.Factory.CreateRibbonButton();
            this.Scheduler_BTN = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "O2G Calendar Sync";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Sync_BTN);
            this.group1.Items.Add(this.Scheduler_BTN);
            this.group1.Items.Add(this.AboutBtn);
            this.group1.Items.Add(this.Debug_BTN);
            this.group1.Label = "Cal Sync";
            this.group1.Name = "group1";
            // 
            // Debug_BTN
            // 
            this.Debug_BTN.Label = "Debug";
            this.Debug_BTN.Name = "Debug_BTN";
            this.Debug_BTN.Visible = false;
            this.Debug_BTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Debug_BTN_Click);
            // 
            // AboutBtn
            // 
            this.AboutBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AboutBtn.Image = global::Outlook_Calendar_Sync.Properties.Resources.question_mark_button;
            this.AboutBtn.Label = "About";
            this.AboutBtn.Name = "AboutBtn";
            this.AboutBtn.ShowImage = true;
            this.AboutBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutBtn_Click);
            // 
            // Sync_BTN
            // 
            this.Sync_BTN.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Sync_BTN.Image = global::Outlook_Calendar_Sync.Properties.Resources.synchronization_arrows;
            this.Sync_BTN.Label = "Sync";
            this.Sync_BTN.Name = "Sync_BTN";
            this.Sync_BTN.ShowImage = true;
            this.Sync_BTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sync_BTN_Click);
            // 
            // Scheduler_BTN
            // 
            this.Scheduler_BTN.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Scheduler_BTN.Image = global::Outlook_Calendar_Sync.Properties.Resources.calendar;
            this.Scheduler_BTN.Label = "Scheduler";
            this.Scheduler_BTN.Name = "Scheduler_BTN";
            this.Scheduler_BTN.ShowImage = true;
            this.Scheduler_BTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Settings_BTN_Click);
            // 
            // SyncRibbon
            // 
            this.Name = "SyncRibbon";
            this.RibbonType = "Microsoft.Outlook.Appointment, Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail" +
    ".Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SyncRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Sync_BTN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Scheduler_BTN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Debug_BTN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutBtn;
    }

    partial class ThisRibbonCollection {
        internal SyncRibbon SyncRibbon
        {
            get { return this.GetRibbon<SyncRibbon>(); }
        }
    }
}
