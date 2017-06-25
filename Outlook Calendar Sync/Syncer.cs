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

        public bool PerformActionToAll { get; set; }

        public int Action { get; set; }

        public Folder Folder { get; set; }

        public Syncer() {
            Action = 0;
            PerformActionToAll = false;
        }

        public void PerformInitalLoad()
        {
            var outlookList = OutlookSync.Syncer.PullListOfAppointments();
            var googleList = GoogleSync.Syncer.PullListOfAppointments();

            var finalList = CompareLists( outlookList, googleList );

            var compare = new CompareForm();
            compare.LoadData( finalList );
            compare.Show();
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

        public void SubmitChanges(List<CalendarItem> items, BackgroundWorker worker )
        {
            int currentCount = 0;

            foreach ( var calendarItem in items )
            {
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
                worker.ReportProgress( progress );
            }

            Archiver.Instance.Save();

            worker.ReportProgress( 100 );
        }

#if DEBUG
        private void WriteToLog(List<CalendarItem> items, string file)
        {
            StringBuilder builder = new StringBuilder();

            foreach ( var calendarItem in items )
                builder.AppendLine( calendarItem.ToString() );

            File.WriteAllText( Application.UserAppDataPath + "\\" + file, builder.ToString() );
        }
#endif
    }
}
