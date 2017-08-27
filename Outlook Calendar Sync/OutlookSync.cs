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

    class OutlookSync
    {

        public static OutlookSync Syncer => m_instance ?? ( m_instance = new OutlookSync() );

        private static OutlookSync m_instance;

        public Application Application { get; set; }

        private List<OutlookFolder> m_folderList;
        private DateTime m_lastUpdate;
        private Folder m_folder;

        public void Init( Application application )
        {
            Application = application;
            m_folder = Application.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar ) as Folder;
            m_lastUpdate = DateTime.MinValue;
            m_folderList = null;
        }

        public void AddAppointment( CalendarItem item )
        {
            try
            {
                var newEvent = item.GetOutlookAppointment(
                    (AppointmentItem) Application.CreateItem( OlItemType.olAppointmentItem ) );
                newEvent.Move( m_folder );
                newEvent.Save();
                Archiver.Instance.Add( item.ID );

                Marshal.ReleaseComObject( newEvent );

            } catch ( Exception ex )
            {
                MessageBox.Show( "Outlook Sync: The following error occurred: " + ex.Message );
            }
        }

        /// <summary>
        /// Pull the list of calendars from Outlook
        /// </summary>
        /// <returns>List of string names.</returns>
        public List<OutlookFolder> PullCalendars()
        {
            if ( m_lastUpdate == DateTime.MinValue || m_lastUpdate < DateTime.Now.Subtract( TimeSpan.FromMinutes( 30 ) ) )
            {
                var oldFolder = m_folder.EntryID;
                SetOutlookWorkingFolder( "", true );

                if ( m_folderList != null )
                {
                    m_folderList.Clear();
                    m_folderList = null;
                }

                m_folderList = new List<OutlookFolder> { new OutlookFolder { Name = m_folder.Name, EntryID = m_folder.EntryID } };

                foreach ( Folder f in m_folder.Folders )
                    m_folderList.Add( new OutlookFolder { Name = f.Name, EntryID = f.EntryID } );

                SetOutlookWorkingFolder( oldFolder );
                m_lastUpdate = DateTime.Now;
            }

            return m_folderList;
        }

        public List<CalendarItem> PullListOfAppointments()
        {
            var calList = new List<CalendarItem>();

            try
            {

                Items item = m_folder.Items;

                foreach ( var i in item )
                {
                    var cal = new CalendarItem();
                    cal.LoadFromOutlookAppointment( (AppointmentItem) i );
                    if ( !calList.Exists( x => x.ID.Equals( cal.ID ) ) )
                        calList.Add( cal );
                }

            } catch ( NullReferenceException )
            {
                Log.Write( "'folder' was null" );
            }

            return calList;
        }

        public List<CalendarItem> PullListOfAppointmentsByDate( DateTime startDate, DateTime endDate )
        {
            var calList = new List<CalendarItem>();

            try
            {

                Items item = GetAppointmentsInRange( m_folder, startDate, endDate );

                foreach ( var i in item )
                {
                    var cal = new CalendarItem();
                    cal.LoadFromOutlookAppointment( (AppointmentItem) i );
                    if ( !calList.Exists( x => x.ID.Equals( cal.ID ) ) )
                        calList.Add( cal );
                }

            } catch ( NullReferenceException )
            {
                Log.Write( "'folder' was null" );
            }

            return calList;
        }

        public CalendarItem FindEvent( string gid )
        {
            CalendarItem c = null;

            foreach ( var i in m_folder.Items )
            {
                var cal = new CalendarItem();
                cal.LoadFromOutlookAppointment( (AppointmentItem) i );
                if ( cal.ID.Equals( gid ) )
                    c = cal;
            }

            return c;
        }

        public CalendarItem FindEventByEntryId( string entryId )
        {

            try
            {
                var appt = (AppointmentItem)Application.Session.GetItemFromID( entryId, m_folder.Store.StoreID );
                if ( appt != null )
                {
                    var calEvent = new CalendarItem();
                    calEvent.LoadFromOutlookAppointment( appt );
                    return calEvent;
                }
            } catch ( COMException ex )
            {
                Log.Write( ex );
            }

            return null;
        }

        public void UpdateAppointment( CalendarItem ev )
        {

            if ( ev.Recurrence != null && Resources.UpdateRecurrance.Equals( "true" ) )
            {

                // Create the filter to find the event. This is done by either the subject and start date or the ID
                var filter = ( ev.Action.HasFlag( CalendarItemAction.GeneratedId ) )
                    ? ( "[Subject]='" + ev.Subject + "'" )
                    : "[ID] = '" + ev.ID + "'";

                /*
                Items items = m_folder.Items;
                //items.IncludeRecurrences = true;
                items.Sort( "[Start]", Type.Missing );

                Items item = items.Restrict( filter );
                */

                var appointmentItem = m_folder.Items.Find( filter );

                //foreach ( AppointmentItem appointmentItem in item ) {
                if ( ev.Recurrence != null )
                {
                    AppointmentItem i = null;
                    if ( appointmentItem.RecurrenceState == 0 )
                    {
                        appointmentItem.GetRecurrencePattern();
                        i = appointmentItem;
                    } else
                        i = appointmentItem.GetRecurrencePattern().GetOccurrence( DateTime.Parse( ev.Start ) );

                    ev.GetOutlookAppointment( i );
                    i.Save();
                    //}

                    ev.Action &= ~CalendarItemAction.GeneratedId;
                    ev.Action &= ~CalendarItemAction.OutlookUpdate;
                }
            } else
            {
                // Check to see if the CalendarItem has a copy of the Outlook AppointmentItem
                if ( ev.ContainsOutlookAppointmentItem )
                {
                    ev.Action &= ~CalendarItemAction.GeneratedId;
                    ev.Action &= ~CalendarItemAction.OutlookUpdate;
                    var i = ev.GetOutlookAppointment();
                    i.Save();
                } else
                {
                    var id = ( ev.Action.HasFlag( CalendarItemAction.GeneratedId ) )
                        ? ( "[Subject]='" + ev.Subject + "'" )
                        : "[ID] = '" + ev.ID + "'";

                    Items items = m_folder.Items;
                    items.Sort( "[Subject]", Type.Missing );

                    Items item = items.Restrict( id );
                    foreach ( AppointmentItem appointmentItem in item )
                    {
                        if ( appointmentItem != null )
                        {
                            ev.Action &= ~CalendarItemAction.GeneratedId;
                            ev.Action &= ~CalendarItemAction.OutlookUpdate;
                            ev.GetOutlookAppointment( appointmentItem );
                            appointmentItem.Save();
                        }
                    }
                }

            }

            if ( !Archiver.Instance.Contains( ev.ID ) )
                Archiver.Instance.Add( ev.ID );

        }

        public void DeleteAppointment( CalendarItem ev )
        {
            var items = m_folder.Items.Restrict( "[ID] = '" + ev.ID + "'" );
            foreach ( AppointmentItem appointmentItem in items )
            {
                appointmentItem.Delete();
            }

            Archiver.Instance.Delete( ev.ID );
        }

        public void SetOutlookWorkingFolder( string entryId, bool defaultFolder = false )
        {
            if ( defaultFolder )
                m_folder = Application.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar ) as Folder;
            else
                m_folder = Application.Session.GetFolderFromID( entryId ) as Folder;
        }

    public MAPIFolder GetDefaultMapiFolder()
        {
            return Application.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar );
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
            Log.Write( filter );

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
