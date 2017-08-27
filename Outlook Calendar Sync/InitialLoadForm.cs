using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Google;
using Google.Apis.Calendar.v3.Data;
using Outlook_Calendar_Sync.Scheduler;
using Settings = Outlook_Calendar_Sync.Properties.Settings;

namespace Outlook_Calendar_Sync {
    public partial class InitialLoadForm : Form {
        private delegate void SetProgressCallback( int progress );

        private CalendarList m_googleFolders;
        private readonly Syncer m_syncer;
        private List<OutlookFolder> m_outlookFolders;
        private int m_step;
        private readonly bool m_multiThreaded;

        private readonly Font m_nonBold;
        private readonly Font m_bolded;

        public InitialLoadForm() {
            InitializeComponent();
            m_step = 0;

            m_nonBold = new Font( Connect_LBL.Font, FontStyle.Regular );
            m_bolded = new Font( Connect_LBL.Font, FontStyle.Bold );
            m_syncer = Syncer.Instance;
            m_multiThreaded = Settings.Default.MultiThreaded;
        }

        private void Connect_BTN_Click( object sender, EventArgs e ) {
            try {
                if ( GoogleSync.Syncer.PerformAuthentication() ) {
                    // Get the list of Google Calendars and load them into googleCal_CB
                    m_googleFolders = GoogleSync.Syncer.PullCalendars();

                    foreach ( var calendarListEntry in m_googleFolders.Items )
                        GoogleCal_CB.Items.Add( calendarListEntry.Summary );

                    // Get the list of Outlook Calendars and load them into the outlookCal_CB
                    m_outlookFolders = OutlookSync.Syncer.PullCalendars();

                    foreach ( var folder in m_outlookFolders )
                        OutlookCal_CB.Items.Add( folder.Name );

                    Connected_LBL.Text = "Connected!!";
                    Connected_LBL.BackColor = Color.Green;
                    Connected_LBL.Visible = true;
                    Connect_BTN.Enabled = false;
                    Next_BTN.Enabled = true;

                } else {
                    Log.Write( "GoogleSync.Syncer.PerformAuthentication returned false." );
                    Connected_LBL.Text = "Failed to Connect";
                    Connected_LBL.BackColor = Color.Red;
                    Connected_LBL.Visible = true;
                }
            } catch ( GoogleApiException error ) {
                Log.Write( error );
                Connected_LBL.Text = "Failed to Connect";
                Connected_LBL.BackColor = Color.Red;
                Connected_LBL.Visible = true;
            }
        }

        private void Next_BTN_Click( object sender, EventArgs e ) {
            switch ( m_step ) {
                case 0:
                    Connect_GB.Visible = false;
                    Connect_GB.Enabled = false;
                    Connect_LBL.Font = m_nonBold;

                    Select_GB.Visible = true;
                    Select_GB.Enabled = true;
                    Select_LBL.Font = m_bolded;

                    m_step++;
                    Next_BTN.Enabled = false;
                    break;

                case 1:
                    Select_GB.Visible = false;
                    Select_GB.Enabled = false;
                    Select_LBL.Font = m_nonBold;

                    Initial_GB.Visible = true;
                    Initial_GB.Enabled = true;
                    Initial_LBL.Font = m_bolded;

                    m_step++;
                    Next_BTN.Enabled = false;
                    Previous_BTN.Enabled = true;
                    Scheduler.Scheduler.Instance.Save( false );
                    break;

                case 2:
                    Initial_GB.Visible = false;
                    Initial_GB.Enabled = false;
                    Initial_LBL.Font = m_nonBold;

                    Done_GB.Visible = true;
                    Done_GB.Enabled = true;
                    Done_LBL.Font = m_bolded;

                    Next_BTN.Enabled = false;
                    break;
            }
        }

        private void Previous_BTN_Click( object sender, EventArgs e ) {
            switch ( m_step ) {
                case 2:
                    Select_GB.Visible = true;
                    Select_GB.Enabled = true;
                    Select_LBL.Font = m_bolded;

                    Initial_GB.Visible = false;
                    Initial_GB.Enabled = false;
                    Initial_LBL.Font = m_nonBold;

                    m_step--;
                    Next_BTN.Enabled = true;
                    Previous_BTN.Enabled = false;
                    break;

                case 3:
                    Initial_GB.Visible = true;
                    Initial_GB.Enabled = true;
                    Initial_LBL.Font = m_bolded;
                    Initial_GB.BringToFront();

                    Done_GB.Visible = false;
                    Done_GB.Enabled = false;
                    Done_LBL.Font = m_nonBold;

                    Next_BTN.Enabled = true;
                    m_step--;
                    break;
            }
        }

        private void Add_BTN_Click( object sender, EventArgs e ) {
            if ( GoogleCal_CB.SelectedIndex >= 0 && OutlookCal_CB.SelectedIndex >= 0 ) {
                var pair = new SyncPair();
                var googleCal = m_googleFolders.Items[GoogleCal_CB.SelectedIndex];
                var outlookCal = m_outlookFolders[OutlookCal_CB.SelectedIndex];

                pair.GoogleName = googleCal.Summary;
                pair.GoogleId = googleCal.Id;
                pair.OutlookName = outlookCal.Name;
                pair.OutlookId = outlookCal.EntryID;

                if ( !Pair_LB.Items.Contains( pair ) ) {
                    Pair_LB.Items.Add( pair );
                    Scheduler.Scheduler.Instance.AddTask( new SchedulerTask { Event = SchedulerEvent.Automatically, Pair = pair} );
                    Next_BTN.Enabled = true;
                }
            } else
                MessageBox.Show( this, "You need to select a calendar from both lists.", "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error );
        }

        private void Remove_BTN_Click( object sender, EventArgs e ) {
            if ( Pair_LB.SelectedIndex >= 0 && Pair_LB.SelectedIndex < Pair_LB.Items.Count ) {
                Pair_LB.Items.RemoveAt( Pair_LB.SelectedIndex );

                if ( Pair_LB.Items.Count < 1 )
                    Next_BTN.Enabled = false;
            }
        }

        private void Cancel_BTN_Click( object sender, EventArgs e ) {
            InitialSyncer_BW.CancelAsync();
        }

        private void Start_BTN_Click( object sender, EventArgs e ) {
            Syncer.Instance.StatusUpdate = SetStatus;

            if ( m_multiThreaded ) {
                Cancel_BTN.Enabled = true;
                InitialSyncer_BW.RunWorkerAsync( Pair_LB.Items );
            } else {
                List<SyncPair> pairs = Pair_LB.Items.Cast<SyncPair>().ToList();

                m_syncer.SynchornizePairs( pairs, InitialSyncer_BW );
            }

            Next_BTN.Enabled = true;
        }

        private void InitialLoadForm_Load( object sender, EventArgs e ) {
            Size = new Size( 580, 370 );
            Connect_GB.Location = new Point( 13, 67 );

            Select_GB.Location = new Point( 13, 67 );
            Select_GB.Visible = false;
            Select_GB.Enabled = false;

            Initial_GB.Location = new Point( 13, 67 );
            Initial_GB.Visible = false;
            Initial_GB.Enabled = false;

            Done_GB.Location = new Point( 13, 67 );
            Done_GB.Visible = false;
            Done_GB.Enabled = false;
        }

        #region BackgroundWorker Methods

        private void InitialSyncer_BW_DoWork( object sender, System.ComponentModel.DoWorkEventArgs e ) {
            List<SyncPair> list = (List<SyncPair>) e.Argument;
            m_syncer.SynchornizePairs( list, InitialSyncer_BW );
        }

        private void InitialSyncer_BW_ProgressChanged( object sender,
            System.ComponentModel.ProgressChangedEventArgs e ) {
            SetProgress( e.ProgressPercentage );
        }

        private void InitialSyncer_BW_RunWorkerCompleted( object sender,
            System.ComponentModel.RunWorkerCompletedEventArgs e ) {
            MessageBox.Show( "Synchronization has been completed." );
        }

        private void SetProgress( int progress ) {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if ( progressBar1.InvokeRequired ) {
                var d = new SetProgressCallback( SetProgress );
                Invoke( d, new[] {progress} );
            } else {
                progressBar1.Value = progress;
            }
        }

        #endregion BackgroundWorker Methods

        private void Close_BTN_Click( object sender, EventArgs e ) {
            Syncer.Instance.StatusUpdate = null;
            Settings.Default.IsInitialLoad = false;
            Settings.Default.Save();
            Scheduler.Scheduler.Instance.ActivateThread();
            Close();
        }

        private void SetStatus( string text ) {
            Status_TB.AppendText( text + "\n" );
            Status_TB.ScrollToCaret();
        }
    }
}