using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;
using Settings = Outlook_Calendar_Sync.Properties.Settings;

namespace Outlook_Calendar_Sync {
    public partial class SyncerForm : Form {

        public SyncRibbon Ribbon { get; set; }

        private delegate void SetProgressCallback( int progress );

        private CalendarList m_googleCalendars;
        private readonly Syncer m_syncer;
        private bool m_multiThreaded = false;

        public SyncerForm() {
            InitializeComponent();
            m_syncer = Syncer.Instance;
        }

        public void StartUpdate( List<CalendarItem> list ) {
            if ( m_multiThreaded )
                calendarUpdate_WORKER.RunWorkerAsync( list );
            else {
                m_syncer.SubmitChanges( list, calendarUpdate_WORKER );
                MessageBox.Show( "Synchronization has been completed." );
            }
        }

        private void Sync_BTN_Click( object sender, EventArgs e ) {

            var finalList = m_syncer.GetFinalList( checkBox1.Checked, Start_DTP.Value, End_DTP.Value );

            var compare = new CompareForm();
            compare.SetParent( this );
            compare.LoadData( finalList );
            compare.Show( this );
            
        }

        private void SyncerForm_Load( object sender, EventArgs e ) {
            // Get the list of Google Calendars and load them into googleCal_CB
            m_googleCalendars = GoogleSync.Syncer.PullCalendars();

            foreach ( var calendarListEntry in m_googleCalendars.Items )
                googleCal_CB.Items.Add( calendarListEntry.Summary );
                
            // Get the list of Outlook Calendars and load them into the outlookCal_CB
            var folders = OutlookSync.Syncer.PullCalendars();

            foreach ( var folder in folders ) { 
                outlookCal_CB.Items.Add( folder.Name );
            }
        }

        private void checkBox1_CheckedChanged( object sender, EventArgs e ) {
            Start_DTP.Enabled = checkBox1.Checked;
            End_DTP.Enabled = checkBox1.Checked;
        }

        #region BackgroundWorker Methods

        private void calendarUpdate_WORKER_DoWork( object sender, System.ComponentModel.DoWorkEventArgs e ) {
            List<CalendarItem> list = (List<CalendarItem>)e.Argument;
            m_syncer.SubmitChanges( list, calendarUpdate_WORKER );
        }

        private void calendarUpdate_WORKER_ProgressChanged( object sender, System.ComponentModel.ProgressChangedEventArgs e ) {
            SetProgress( e.ProgressPercentage );
        }

        private void calendarUpdate_WORKER_RunWorkerCompleted( object sender, System.ComponentModel.RunWorkerCompletedEventArgs e ) {
            MessageBox.Show( "Synchronization has been completed." );
        }

        private void SetProgress( int progress ) {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if ( progressBar1.InvokeRequired ) {
                var d = new SetProgressCallback( SetProgress );
                Invoke( d, new[] { progress } );
            } else {
                progressBar1.Value = progress;
            }
        }
        #endregion BackgroundWorker Methods

        private void button1_Click( object sender, EventArgs e ) {
            if ( File.Exists( Application.UserAppDataPath + "\\" + "calendarItems.txt" ) ) {
                File.Delete( Application.UserAppDataPath + "\\" + "calendarItems.txt" );
                Settings.Default.IsInitialLoad = true;
                Settings.Default.Save();
                MessageBox.Show( this, "Reset Initial Load", "Reset", MessageBoxButtons.OK );
            }
        }


        private void SyncerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if ( e.CloseReason == CloseReason.UserClosing )
            {
                e.Cancel = true;
                Hide();

                m_syncer.Action = 0;
                m_syncer.PerformActionToAll = false;
            }
        }

    }
}
