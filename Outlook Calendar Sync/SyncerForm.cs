using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;

namespace Outlook_Calendar_Sync {
    public partial class SyncerForm : Form {

        private delegate void SetProgressCallback( int progress );

        private CalendarList m_googleCalendars;
        private bool m_multiThreaded = false;

        public SyncerForm() {
            InitializeComponent();
        }

        public void StartUpdate( List<CalendarItem> list ) {
            if ( m_multiThreaded )
                calendarUpdate_WORKER.RunWorkerAsync( list );
            else
                SubmitChanges( list );
        }

        private void Sync_BTN_Click( object sender, EventArgs e ) {
            List<CalendarItem> outlookList;
            List<CalendarItem> googleList;

            if ( checkBox1.Checked ) {
                outlookList = OutlookSync.Syncer.PullListOfAppointmentsByDate( Start_DTP.Value, End_DTP.Value );
                googleList = GoogleSync.Syncer.PullListOfAppointmentsByDate( Start_DTP.Value, End_DTP.Value );
            }
            else {
                outlookList = OutlookSync.Syncer.PullListOfAppointments();
                googleList = GoogleSync.Syncer.PullListOfAppointments();
            }

            // Check to see what events need to be added to google from outlook
            var finalList = CompareLists( outlookList, googleList );

            var compare = new CompareForm();
            compare.SetParent( this );
            compare.LoadData( finalList );
            compare.Show( this );
            
        }

        private List<CalendarItem> CompareLists( List<CalendarItem> outlookList, List<CalendarItem> googleList ) {
            var finalList = new List<CalendarItem>();

            foreach ( var calendarItem in outlookList ) {
                if ( !googleList.Contains( calendarItem ) ) {

                    if ( Archiver.Instance.Contains( calendarItem.ID ) ) {
                        if (
                            MessageBox.Show(
                                "It appears the calendar event '" + calendarItem.Subject +
                                "' was deleted from Google. Would you like to remove it from Outlook also?", "Delete Event?",
                                MessageBoxButtons.YesNo ) == DialogResult.Yes ) {
                            calendarItem.Action |= CalendarItemAction.OutlookDelete;
                            finalList.Add( calendarItem );
                        }
                    }
                    else {

                        if ( calendarItem.Recurrence != null ) {
                            if ( calendarItem.IsFirstOccurence  ) {
                                calendarItem.Action |= CalendarItemAction.GoogleAdd;
                                finalList.Add( calendarItem );
                            }
                        }
                        else {
                            calendarItem.Action |= CalendarItemAction.GoogleAdd;
                            finalList.Add( calendarItem );
                        }
                    }
                }
                else {
                    var item = googleList.Find( x => x.ID.Equals( calendarItem.ID ) );
                    item.Action |= CalendarItemAction.ContentsEqual;

                    if ( !item.Equals( calendarItem ) ) {

                        var result = DifferencesForm.Show( calendarItem, item );

                        if ( result == DialogResult.Yes ) {
                            calendarItem.Action |= CalendarItemAction.GoogleUpdate;
                            finalList.Add( calendarItem );
                        }
                        else if ( result == DialogResult.No ) {
                            item.Action |= CalendarItemAction.OutlookUpdate;
                            finalList.Add( item );
                        }
                    }
                }

            }

            foreach ( var calendarItem in googleList ) {
                if ( !outlookList.Contains( calendarItem ) ) {
                    if ( Archiver.Instance.Contains( calendarItem.ID ) ) {
                        if (
                            MessageBox.Show(
                                "It appears the calendar event '" + calendarItem.Subject +
                                "' was deleted from Outlook. Would you like to remove it from Google also?",
                                "Delete Event?",
                                MessageBoxButtons.YesNo ) == DialogResult.Yes ) {
                            calendarItem.Action |= CalendarItemAction.GoogleDelete;
                            finalList.Add( calendarItem );
                        }
                    }
                    else {
                        calendarItem.Action |= CalendarItemAction.OutlookAdd;
                        finalList.Add( calendarItem );
                    }
                }
            }

            return finalList;
        }

        private void SyncerForm_Load( object sender, EventArgs e ) {
            // Get the list of Google Calendars and load them into googleCal_CB
            
            m_googleCalendars = GoogleSync.Syncer.PullCalendars();

            foreach ( var calendarListEntry in m_googleCalendars.Items )
                googleCal_CB.Items.Add( calendarListEntry.Summary );
                
            // Get the list of Outlook Calendars and load them into the outlookCal_CB
            var folders = OutlookSync.Syncer.PullCalendars();

            foreach ( var folder in folders ) { 
                outlookCal_CB.Items.Add( folder );
            }

        }

        private void checkBox1_CheckedChanged( object sender, EventArgs e ) {
            Start_DTP.Enabled = checkBox1.Checked;
            End_DTP.Enabled = checkBox1.Checked;
        }

        #region BackgroundWorker Methods
        private void SubmitChanges( List<CalendarItem> items ) {
            int currentCount = 0;

            foreach ( var calendarItem in items ) {
                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookAdd ) )
                    OutlookSync.Syncer.AddAppointment( calendarItem );

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleAdd ) )
                    GoogleSync.Syncer.AddAppointment( calendarItem );

                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookUpdate ) )
                    OutlookSync.Syncer.UpdateAppointment( calendarItem );

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleUpdate ) )
                    GoogleSync.Syncer.UpdateAppointment( calendarItem );

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleDelete ) )
                    GoogleSync.Syncer.DeleteAppointment( calendarItem );

                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookDelete ) )
                    OutlookSync.Syncer.DeleteAppointment( calendarItem );

                currentCount++;
                var progress = (int)( currentCount / (float)items.Count * 100 );
                calendarUpdate_WORKER.ReportProgress( progress );
            }

            Archiver.Instance.Save();

            calendarUpdate_WORKER.ReportProgress( 100 );
        }

        private void calendarUpdate_WORKER_DoWork( object sender, System.ComponentModel.DoWorkEventArgs e ) {
            List<CalendarItem> list = (List<CalendarItem>)e.Argument;
            SubmitChanges( list );
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
                //Invoke( d, new[] { progress } );
            } else {
                progressBar1.Value = progress;
            }
        }
        #endregion BackgroundWorker Methods

        private void button1_Click( object sender, EventArgs e ) {
            OutlookSync.Syncer.PullListOfAppointments();
            OutlookSync.Syncer.PullCalendars();
        }
    }
}
