using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Outlook_Calendar_Sync.Properties;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;

namespace Outlook_Calendar_Sync {

    struct OutlookFolder {
        public string Name;
        public string EntryID;
    }

    class OutlookSync {

        public static OutlookSync Syncer => m_instance ?? ( m_instance = new OutlookSync() );

        private static OutlookSync m_instance;

        public Application Application { get; set; }

        private readonly Folder m_folder;

        public OutlookSync() {
            Application = new Application();
            m_folder = Application.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar ) as Folder;
        }

        public void AddAppointment( CalendarItem item ) {
            try {
                var newEvent = item.GetOutlookAppointment( (AppointmentItem) Application.CreateItem( OlItemType.olAppointmentItem ) );
                newEvent.Save();
                Archiver.Instance.Add( item.ID );

                Marshal.ReleaseComObject( newEvent );

            } catch ( Exception ex ) {
                MessageBox.Show( "Outlook Sync: The following error occurred: " + ex.Message );
            } 
        }

        /// <summary>
        /// Pull the list of calendars from Outlook
        /// </summary>
        /// <returns>List of string names.</returns>
        public List<OutlookFolder> PullCalendars() {
            var list = new List<OutlookFolder> { new OutlookFolder { Name = m_folder.Name, EntryID = m_folder.EntryID } };

            foreach ( Folder f in m_folder.Folders )
                list.Add( new OutlookFolder { Name = f.Name, EntryID = f.EntryID } );

            return list;
        }

        public List<CalendarItem> PullListOfAppointments() { 
            var calList = new List<CalendarItem>();

            try {

                Items item = m_folder.Items;

                foreach ( var i in item ) {
                    var cal = new CalendarItem();
                    cal.LoadFromOutlookAppointment( (AppointmentItem)i );
                    if ( !calList.Exists( x => x.ID.Equals( cal.ID ) ) )
                        calList.Add( cal );
                }

            } catch ( NullReferenceException ) {
                Debug.WriteLine( "'folder' was null" );
            }

            return calList;
        }

        public List<CalendarItem> PullListOfAppointmentsByDate( DateTime startDate, DateTime endDate ) {
            var calList = new List<CalendarItem>();

            try {

                Items item = GetAppointmentsInRange( m_folder, startDate, endDate );

                foreach ( var i in item ) {
                    var cal = new CalendarItem();
                    cal.LoadFromOutlookAppointment( (AppointmentItem)i );
                    if ( !calList.Exists( x => x.ID.Equals( cal.ID ) ) )
                        calList.Add( cal );
                }

            } catch ( NullReferenceException ) {
                Debug.WriteLine( "'folder' was null" );
            }

            return calList;
        }

        public CalendarItem FindEvent( string gid ) {
            CalendarItem c = null;

            foreach ( var i in m_folder.Items ) {
                var cal = new CalendarItem();
                cal.LoadFromOutlookAppointment( (AppointmentItem)i );
                if ( cal.ID.Equals( gid ) )
                    c = cal;
            }

            return c;
        }

        public void UpdateAppointment( CalendarItem ev ) {

            if ( ev.Recurrence != null && Resources.UpdateRecurrance.Equals( "true" ) ) {
                //throw new NotImplementedException("Recurrance update is not fully incorported yet. This will not effect single event updates.");

                // TODO: Find out why this item in not found when using recurring events
                //var filter = "[Start] >= '" + DateTime.Parse( ev.Start ).ToString( "g" ) + "' AND [End] <= '" +
                //             DateTime.Parse( ev.End ).ToString( "g" ) + "'";

#pragma warning disable 162
                var filter = ( ev.Action.HasFlag( CalendarItemAction.GeneratedId ) )
                    ? ( "[Subject]='" + ev.Subject + "'" )
                    : "[ID] = '" + ev.ID + "'";

                Items items = m_folder.Items;
                //items.IncludeRecurrences = true;
                items.Sort( "[Start]", Type.Missing );

                Items item = items.Restrict( filter );

                foreach ( AppointmentItem appointmentItem in item ) {
                    if ( ev.Recurrence != null ) {
                        AppointmentItem i = null;
                        if ( appointmentItem.RecurrenceState == OlRecurrenceState.olApptNotRecurring ) {
                            appointmentItem.GetRecurrencePattern();
                            i = appointmentItem;
                        } else
                            i = appointmentItem.GetRecurrencePattern().GetOccurrence( DateTime.Parse( ev.Start ) );

                        ev.GetOutlookAppointment( i );
                        i.Save();
                    }

                    ev.Action &= ~CalendarItemAction.GeneratedId;
                    ev.Action &= ~CalendarItemAction.OutlookUpdate;
                }
#pragma warning restore 162
            } else {
                // Check to see if the CalendarItem has a copy of the Outlook AppointmentItem
                if ( ev.ContainsOutlookAppointmentItem ) {
                    ev.Action &= ~CalendarItemAction.GeneratedId;
                    ev.Action &= ~CalendarItemAction.OutlookUpdate;
                    var i = ev.GetOutlookAppointment();
                    i.Save();
                } else {
                    var id = ( ev.Action.HasFlag( CalendarItemAction.GeneratedId ) )
                    ? ( "[Subject]='" + ev.Subject + "'" )
                    : "[ID] = '" + ev.ID + "'";

                    Items items = m_folder.Items;
                    items.Sort( "[Subject]", Type.Missing );

                    Items item = items.Restrict( id );
                    foreach ( AppointmentItem appointmentItem in item ) {
                        if ( appointmentItem != null ) {
                            ev.Action &= ~CalendarItemAction.GeneratedId;
                            ev.Action &= ~CalendarItemAction.OutlookUpdate;
                            ev.GetOutlookAppointment( appointmentItem );
                            appointmentItem.Save();
                        }
                    }
                }
               
            }

        }

        public void DeleteAppointment( CalendarItem ev ) {
             var items = m_folder.Items.Restrict( "[ID] = '" + ev.ID + "'"  );
            foreach ( AppointmentItem appointmentItem in items ) {
                appointmentItem.Delete();
            }

            Archiver.Instance.Delete( ev.ID );
        }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns>Outlook.Items</returns>
        private Items GetAppointmentsInRange( Folder folder, DateTime startTime, DateTime endTime ) {
            string filter = "[Start] >= '"
                + startTime.ToString( "g" )
                + "' AND [End] <= '"
                + endTime.ToString( "g" ) + "'";
            Debug.WriteLine( filter );

            try {
                Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort( "[Start]", Type.Missing );
                Items restrictItems = calItems.Restrict( filter );
                if ( restrictItems.Count > 0 ) {
                    return restrictItems;
                } else {
                    return null;
                }
            } catch { return null; }
        }

    }
}
