using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using Outlook_Calendar_Sync.Enums;
using System.Windows.Forms;

namespace Outlook_Calendar_Sync {
    
    /// <summary>
    /// The Syncer is responsible for pulling, comparing, and pushing the changes to the calendars.
    /// </summary>
    public class Syncer {

        /// <summary>
        /// This is the instance of the Syncer to be used. 
        /// </summary>
        public static Syncer Instance => _instance ?? ( _instance = new Syncer() );
        private static Syncer _instance;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        public delegate void WriteToStatus( string text );

        /// <summary>
        /// This delegate is used to send status messages
        /// </summary>
        public WriteToStatus StatusUpdate;

        /// <summary>
        /// Will the same action be performed to all events with changes?
        /// </summary>
        public bool PerformActionToAll { get; set; }

        /// <summary>
        /// This is the action to be performed when PerformActionToAll is true
        /// </summary>
        public CalendarItemAction Action { get; set; }

        /// <summary>
        /// Do you want the sync to use a sync token
        /// </summary>
        public bool IsUsingSyncToken { get; set; } 
        
        private bool m_syncingPairs;                    // Are we syncing pairs?
        private bool m_silentSync;                      // Performing a silent sync?
        private Precedence m_precedence;                // What is the precedence for the silent sync
        private readonly OutlookSync m_outlookSync;     // A reference to the outlook syncer
        private readonly GoogleSync m_googleSync;       // A reference to the google syncer
        private readonly Archiver m_archiver;           // A reference to the archiver

        /// <summary>
        /// Default constructor, sets all variables to their default values
        /// </summary>
        public Syncer() {
            Action = CalendarItemAction.Nothing;
            PerformActionToAll = false;
            IsUsingSyncToken = false;
            m_syncingPairs = false;

            m_outlookSync = OutlookSync.Syncer;
            m_googleSync = GoogleSync.Syncer;
            m_archiver = Archiver.Instance;
            m_precedence = Precedence.None;
            m_silentSync = false;
        }

        /// <summary>
        /// Performs the actual pulling of events and appointments, compares them, and returns a list of the ones with differences.
        /// </summary>
        /// <param name="byDate">Do you want to restrict the query by date?</param>
        /// <param name="start">The start date of the restriction</param>
        /// <param name="end">The end date of the restriction</param>
        /// <returns>A list of calendar items that need to be updated or added.</returns>
        public List<CalendarItem> GetFinalList(bool byDate = false, DateTime start = default( DateTime ), DateTime end = default( DateTime )) {
            List<CalendarItem> outlookList;
            List<CalendarItem> googleList;

            if ( byDate )
            {
                outlookList = m_outlookSync.PullListOfAppointmentsByDate( start, end );
                googleList = m_googleSync.PullListOfAppointmentsByDate( start, end );
            } else 
            {
                outlookList = m_outlookSync.PullListOfAppointments();
                googleList = IsUsingSyncToken
                    ? m_googleSync.PullListOfAppointmentsBySyncToken()
                        : m_googleSync.PullListOfAppointments();
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

        /// <summary>
        /// Compares the two lists of calendar items and finds the differences in the lists.
        /// </summary>
        /// <param name="outlookList">The outlook list</param>
        /// <param name="googleList">The google list</param>
        /// <returns>A list of calendar items with the appropriate changes specified.</returns>
        private List<CalendarItem> CompareLists(List<CalendarItem> outlookList, List<CalendarItem> googleList )
        {
            var finalList = new List<CalendarItem>();

            foreach ( var calendarItem in outlookList )
            {
                // Check to see if Item is not in google list
                if ( !googleList.Contains( calendarItem ) )
                {
                    // It's not, check the archiver. Maybe it was deleted in google.
                    if ( m_archiver.Contains( calendarItem.ID ) )
                    {
                        // It appears to have been deleted in the google cal.
                        // Do you want to delete it from outlook?
                        if (
                            MessageBox.Show(
                                "It appears the calendar event '" + calendarItem.Subject +
                                "' was deleted from Google. Would you like to remove it from Outlook also?", "Delete Event?",
                                MessageBoxButtons.YesNo ) == DialogResult.Yes )
                        {
                            // Yes, delete it
                            calendarItem.Action |= CalendarItemAction.OutlookDelete;
                            finalList.Add( calendarItem );
                        }
                      
                    } else
                    { // It is not it google's list or the archiver, add it to google

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
                { // It is in googles list. Find them differences.

                    var item = googleList.Find( x => x.ID.Equals( calendarItem.ID ) );
                    // This forces the item to check for differences
                    item.Action |= CalendarItemAction.ContentsEqual;

                    // Get the differences
                    if ( !item.Equals( calendarItem ) )
                    {
                        // Should we do the same action to every calendar item?
                        if ( PerformActionToAll )
                        {
                            // Check and make sure there is an action to perform
                            if ( Action != CalendarItemAction.Nothing )
                            {
                                calendarItem.Action |= Action;
                                finalList.Add( calendarItem );
                            }
                        } else
                        { // We are not performing the same action to every item, yet.

                            // Are we performing a silentSync?
                            if ( m_silentSync )
                            {
                                // Yep, setup the action to be performed
                                PerformActionToAll = true;
                                Action = m_precedence == Precedence.Outlook
                                    ? CalendarItemAction.GoogleUpdate
                                    : m_precedence == Precedence.Google
                                        ? CalendarItemAction.OutlookUpdate
                                        : CalendarItemAction.Nothing;
                            } else
                            { // No we are not performing a silent sync

                                // Check to see if the only thing that needs to updated is the calid
                                if ( item.Changes == CalendarItemChanges.CalId &&
                                     calendarItem.Changes == CalendarItemChanges.Nothing )
                                {
                                    calendarItem.Action = CalendarItemAction.OutlookUpdate;
                                } else
                                {
                                    var result = (DifferencesFormResults)DifferencesForm.Show( calendarItem, item );

                                    // Save Outlook Version Once
                                    if ( result == DifferencesFormResults.KeepOutlookSingle )
                                    {
                                        calendarItem.Action |= CalendarItemAction.GoogleUpdate;
                                        finalList.Add( calendarItem );

                                        // Save Outlook Version for All
                                    } else if ( result == DifferencesFormResults.KeepOutlookAll )
                                    {
                                        calendarItem.Action |= CalendarItemAction.GoogleUpdate;
                                        finalList.Add( calendarItem );

                                        Action = CalendarItemAction.GoogleUpdate;
                                        PerformActionToAll = true;

                                        // Save Google Version Once
                                    } else if ( result == DifferencesFormResults.KeepGoogleSingle )
                                    {
                                        item.Action |= CalendarItemAction.OutlookUpdate;
                                        finalList.Add( item );

                                        // Save Google Version for All
                                    } else if ( result == DifferencesFormResults.KeepGoogleAll )
                                    {
                                        item.Action |= CalendarItemAction.OutlookUpdate;
                                        finalList.Add( item );

                                        Action = CalendarItemAction.OutlookUpdate;
                                        PerformActionToAll = true;

                                        // Ignore All
                                    } else if ( result == DifferencesFormResults.IgnoreAll )
                                    {
                                        PerformActionToAll = true;
                                        Action = CalendarItemAction.Nothing;
                                    }
                                }
                            }
                        }
                    }
                }

            }

            foreach ( var calendarItem in googleList )
            {
                if ( !outlookList.Contains( calendarItem ) )
                {
                    if ( m_archiver.Contains( calendarItem.ID ) )
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

        /// <summary>
        /// Submits the changes requested to the Outlook and Google calendars
        /// </summary>
        /// <param name="items">The list of calendar items with the required changes</param>
        /// <param name="worker">A background worker</param>
        /// <param name="pairProgress">The completed progress when submitting changes to multiple pairs</param>
        public void SubmitChanges( List<CalendarItem> items, BackgroundWorker worker, float pairProgress = 1 )
        {
            int currentCount = 0;

            foreach ( var calendarItem in items )
            {
                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookAdd ) ) {
                    StatusUpdate?.Invoke( "- Adding " + calendarItem.Subject + " to Outlook." );

                    m_outlookSync.AddAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleAdd ) ) {
                    StatusUpdate?.Invoke( "- Adding " + calendarItem.Subject + " to Google." );

                    m_googleSync.AddAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookUpdate ) ) {
                    StatusUpdate?.Invoke( "- Updating " + calendarItem.Subject + " in Outlook." );

                    m_outlookSync.UpdateAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleUpdate ) ) {
                    StatusUpdate?.Invoke( "- Updating " + calendarItem.Subject + " in Google." );

                    m_googleSync.UpdateAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.GoogleDelete ) ) {
                    StatusUpdate?.Invoke( "- Deleting " + calendarItem.Subject + " from Outlook." );

                    m_googleSync.DeleteAppointment( calendarItem );
                }

                if ( calendarItem.Action.HasFlag( CalendarItemAction.OutlookDelete ) ) {
                    StatusUpdate?.Invoke( "- Deleting " + calendarItem.Subject + " from Google." );
                    m_outlookSync.DeleteAppointment( calendarItem );
                }

                currentCount++;
                var progress = (int)( currentCount / (float)items.Count * 100 * pairProgress);
                if ( worker.WorkerReportsProgress )
                    worker.ReportProgress( progress );
            }

            if ( !m_syncingPairs ) {
                m_archiver.Save();
                if ( worker.WorkerReportsProgress )
                    worker.ReportProgress( 100 );
                StatusUpdate?.Invoke( "- Sync has been completed." );
            }
        }

        /// <summary>
        /// Synchronizes a SyncPair.
        /// </summary>
        /// <param name="pair">The pair to sync</param>
        /// <param name="precedence">The precendence used when performing silent syncing</param>
        /// <param name="silentSync">Perform a silent sync?</param>
        public void SynchornizePairs( SyncPair pair, Precedence precedence = Precedence.None, bool silentSync = false )
        {
            m_precedence = precedence;
            m_silentSync = silentSync;

            SynchornizePairs( new List<SyncPair> { pair }, new BackgroundWorker());

            m_silentSync = false;
            m_precedence = Precedence.None;
            PerformActionToAll = false;
        }

        /// <summary>
        /// Synchronizes a list of sync pairs
        /// </summary>
        /// <param name="pairs">The list of pairs to sync</param>
        /// <param name="worker">A background work used to perform the sync</param>
        public void SynchornizePairs( List<SyncPair> pairs, BackgroundWorker worker )
        {
            StatusUpdate?.Invoke( "- Starting Sync" );

            float count = 0;
            m_syncingPairs = true;

            foreach ( SyncPair pair in pairs ) {
                StatusUpdate?.Invoke( $"- Starting Sync for  {pair.OutlookName} and {pair.GoogleName}" );

                m_outlookSync.SetOutlookWorkingFolder( pair.OutlookId );
                m_googleSync.SetGoogleWorkingFolder( pair.GoogleId );
                m_archiver.CurrentPair = pair;

                var finalList = GetFinalList();
                var progress = ++count / pairs.Count;

                if ( finalList.Count > 0 )
                {
                    var compare = new CompareForm();
                    compare.SetCalendars( pair );
                    compare.LoadData( finalList );

                    if ( !m_silentSync )
                    {
                        if ( compare.ShowDialog() == DialogResult.OK )
                            SubmitChanges( compare.Data, worker, progress );
                    }
                    else
                    {
                        SubmitChanges( finalList, worker, progress );
                    }
                }
                else
                {
                    MessageBox.Show( $"There are no differences between the {pair.GoogleName} Google Calender and the {pair.OutlookName} Outlook Calender.", "No Events", MessageBoxButtons.OK,
                        MessageBoxIcon.Information );
                    StatusUpdate?.Invoke( "The was no items to sync." );
                }

                StatusUpdate?.Invoke( $"- Sync Completed for {pair.OutlookName} and {pair.GoogleName}" );
            }

            m_syncingPairs = false;
            PerformActionToAll = false;
            m_archiver.Save();

            if ( worker.WorkerReportsProgress )
                worker.ReportProgress( 100 );

            StatusUpdate?.Invoke( "- Sync has been completed." );
        }

        /// <summary>
        /// Searches through the specified SyncPair and finds events that were deleted.
        /// </summary>
        /// <param name="pair">The pair to search</param>
        /// <returns>A list of IDs of the deleted events</returns>
        public List<string> FindDeletedEvents( SyncPair pair )
        {
            var list = new List<string>();

            m_outlookSync.SetOutlookWorkingFolder( pair.OutlookId );
            m_archiver.CurrentPair = pair;

            var outlookAppts = m_outlookSync.PullListOfAppointments();
            var archlist = m_archiver.GetListForSyncPair( pair );

            foreach ( var id in archlist )
            {
                if ( outlookAppts.Find( x => x.ID == id ) == null )
                    list.Add( id );
            }

            return list;
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
