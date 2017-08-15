using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Application = System.Windows.Forms.Application;

namespace Outlook_Calendar_Sync {
    
    public class Syncer {

        public static Syncer Instance => _instance ?? ( _instance = new Syncer() );
        private static Syncer _instance;

        public delegate void WriteToStatus( string text );

        public WriteToStatus StatusUpdate;

        public bool PerformActionToAll { get; set; }

        public int Action { get; set; }

        public Folder Folder { get; set; }

        private bool m_syncingPairs;

        public Syncer() {
            Action = 0;
            PerformActionToAll = false;
            m_syncingPairs = false;
        }

        public List<CalendarItem> GetFinalList(bool byDate = false, DateTime start = default( DateTime ), DateTime end = default( DateTime )) {
            List<CalendarItem> outlookList;
            List<CalendarItem> googleList;

            if ( byDate )
            {
                outlookList = OutlookSync.Syncer.PullListOfAppointmentsByDate( start, end );
                googleList = GoogleSync.Syncer.PullListOfAppointmentsByDate( start, end );
            } else
            {
                outlookList = OutlookSync.Syncer.PullListOfAppointments();
                googleList = GoogleSync.Syncer.PullListOfAppointments();
            }

            // Check to see what events need to be added to google from outlook
            var finalList = CompareLists( outlookList, googleList );

#if DEBUG
            WriteToLog( outlookList, "Outlook List Log.rtf" );
            WriteToLog( googleList, "Google List Log.rtf" );
            WriteToLog( finalList, "Final List Log.rtf" );
#endif
            return finalList;
        }

        private List<CalendarItem> CompareLists(List<CalendarItem> outlookList, List<CalendarItem> googleList)
        {
            var finalList = new List<CalendarItem>();

            foreach ( var calendarItem in outlookList )
            {
                if ( !googleList.Contains( calendarItem ) )
                {

                    if ( Archiver.Instance.Contains( calendarItem.ID ) )
                    {
                        if (
                            MessageBox.Show(
                                "It appears the calendar event '" + calendarItem.Subject +
                                "' was deleted from Google. Would you like to remove it from Outlook also?", "Delete Event?",
                                MessageBoxButtons.YesNo ) == DialogResult.Yes )
                        {

                            calendarItem.Action |= CalendarItemAction.OutlookDelete;
                            finalList.Add( calendarItem );
                        }
                    } else
                    {

                        if ( calendarItem.Recurrence != null )
                        {
                            if ( calendarItem.IsFirstOccurence )
                            {
                                calendarItem.Action |= CalendarItemAction.GoogleAdd;
                                finalList.Add( calendarItem );
                            }
                        } else
                        {
                            calendarItem.Action |= CalendarItemAction.GoogleAdd;
                            finalList.Add( calendarItem );
                        }
                    }
                } else
                {
                    var item = googleList.Find( x => x.ID.Equals( calendarItem.ID ) );
                    item.Action |= CalendarItemAction.ContentsEqual;

                    if ( !item.Equals( calendarItem ) )
                    {

                        if ( PerformActionToAll )
                        {
                            if ( Action != 0 )
                            {
                                calendarItem.Action |= (CalendarItemAction)Action;
                                finalList.Add( calendarItem );
                            }
                        } else
                        {
                            var result = DifferencesForm.Show( calendarItem, item );

                            // Save Outlook Version Once
                            if ( result == DialogResult.Yes )
                            {
                                calendarItem.Action |= CalendarItemAction.GoogleUpdate;
                                finalList.Add( calendarItem );

                                // Save Outlook Version for All
                            } else if ( result == DialogResult.OK )
                            {
                                calendarItem.Action |= CalendarItemAction.GoogleUpdate;
                                finalList.Add( calendarItem );

                                Action = (int)CalendarItemAction.GoogleUpdate;
                                PerformActionToAll = true;

                                // Save Google Version Once
                            } else if ( result == DialogResult.No )
                            {
                                item.Action |= CalendarItemAction.OutlookUpdate;
                                finalList.Add( item );

                                // Save Google Version for All
                            } else if ( result == DialogResult.None )
                            {
                                item.Action |= CalendarItemAction.OutlookUpdate;
                                finalList.Add( item );

                                Action = (int)CalendarItemAction.OutlookUpdate;
                                PerformActionToAll = true;

                                // Ignore All
                            } else if ( result == DialogResult.Ignore )
                            {
                                PerformActionToAll = true;
                            }
                        }
                    }
                }

            }

            foreach ( var calendarItem in googleList )
            {
                if ( !outlookList.Contains( calendarItem ) )
                {
                    if ( Archiver.Instance.Contains( calendarItem.ID ) )
                    {
                        if (
                            MessageBox.Show(
                                "It appears the calendar event '" + calendarItem.Subject +
                                "' was deleted from Outlook. Would you like to remove it from Google also?",
                                "Delete Event?",
                                MessageBoxButtons.YesNo ) == DialogResult.Yes )
                        {
                            calendarItem.Action |= CalendarItemAction.GoogleDelete;
                            finalList.Add( calendarItem );
                        }
                    } else
                    {
                        calendarItem.Action |= CalendarItemAction.OutlookAdd;
                        finalList.Add( calendarItem );
                    }
                }
            }

            return finalList;
        }

        public void SubmitChanges( List<CalendarItem> items, BackgroundWorker worker, float pairProgress = 1 )
        {
            int currentCount = 0;

            foreach ( var calendarItem in items )
            {
                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookAdd ) ) {
                    StatusUpdate?.Invoke( "- Adding " + calendarItem.Subject + " to Outlook." );

                    OutlookSync.Syncer.AddAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleAdd ) ) {
                    StatusUpdate?.Invoke( "- Adding " + calendarItem.Subject + " to Google." );

                    GoogleSync.Syncer.AddAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookUpdate ) ) {
                    StatusUpdate?.Invoke( "- Updating " + calendarItem.Subject + " in Outlook." );

                    OutlookSync.Syncer.UpdateAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleUpdate ) ) {
                    StatusUpdate?.Invoke( "- Updating " + calendarItem.Subject + " in Google." );

                    GoogleSync.Syncer.UpdateAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleDelete ) ) {
                    StatusUpdate?.Invoke( "- Deleting " + calendarItem.Subject + " from Outlook." );

                    GoogleSync.Syncer.DeleteAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookDelete ) ) {
                    StatusUpdate?.Invoke( "- Deleting " + calendarItem.Subject + " from Google." );
                    OutlookSync.Syncer.DeleteAppointment( calendarItem );
                }

                currentCount++;
                var progress = (int)( currentCount / (float)items.Count * 100 * pairProgress);
                if ( worker.WorkerReportsProgress )
                    worker.ReportProgress( progress );
            }

            if ( !m_syncingPairs ) {
                Archiver.Instance.Save();
                if ( worker.WorkerReportsProgress )
                    worker.ReportProgress( 100 );
                StatusUpdate?.Invoke( "- Sync has been completed." );
            }
        }

        public void SynchornizePairs( SyncPair pair )
        {
            SynchornizePairs( new List<SyncPair> { pair }, new BackgroundWorker());
        }

        public void SynchornizePairs( List<SyncPair> pairs, BackgroundWorker worker )
        {
            StatusUpdate?.Invoke( "- Starting Sync" );

            float count = 0;
            m_syncingPairs = true;
            foreach ( SyncPair pair in pairs ) {
                StatusUpdate?.Invoke( "- Starting Sync for " + pair.OutlookName + " and " + pair.GoogleName );

                OutlookSync.Syncer.SetOutlookWorkingFolder( pair.OutlookId );
                GoogleSync.Syncer.SetGoogleWorkingFolder( pair.GoogleId );
                Archiver.Instance.CurrentPair = pair;

                var finalList = GetFinalList();
                var progress = ++count / pairs.Count;

                var compare = new CompareForm();
                compare.SetCalendars( pair );
                compare.LoadData( finalList );

                if ( compare.ShowDialog() == DialogResult.OK )
                    SubmitChanges( compare.Data, worker, progress );

                StatusUpdate?.Invoke( "- Sync Completed for " + pair.OutlookName + " and " + pair.GoogleName );
            }
            m_syncingPairs = false;
            Archiver.Instance.Save();
            if ( worker.WorkerReportsProgress )
                worker.ReportProgress( 100 );
            StatusUpdate?.Invoke( "- Sync has been completed." );
        }

#if DEBUG
        private void WriteToLog(List<CalendarItem> items, string file)
        {
            StringBuilder builder = new StringBuilder();

            foreach ( var calendarItem in items )
                builder.AppendLine( calendarItem.ToString() );

            if ( !Directory.Exists( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                    "\\OutlookGoogleSync\\" ) )
                Directory.CreateDirectory( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                           "\\OutlookGoogleSync" );

            File.WriteAllText( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + file, builder.ToString() );
        }
#endif
    }
}
