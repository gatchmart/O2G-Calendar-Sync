using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Outlook_Calendar_Sync.Enums;
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

        public static OutlookSync Syncer => _instance ?? ( _instance = new OutlookSync() );

        private static OutlookSync _instance;

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
                
                var newEvent = item.GetOutlookAppointment( (AppointmentItem) Application.CreateItem( OlItemType.olAppointmentItem ) );
                newEvent.Move( m_folder );
                newEvent.Save();
                Archiver.Instance.Add( item.ID );

                Log.Write( $"Added {item} to Outlooks Calendar {m_folder.Name}" );

                Marshal.ReleaseComObject( newEvent );

            } catch ( Exception ex )
            {
                Log.Write( ex );
                MessageBox.Show( "Outlook Sync: The following error occurred: " + ex.Message );
            }
        }

        /// <summary>
        /// Pull the list of calendars from Outlook
        /// </summary>
        /// <returns>List of string names.</returns>
        public List<OutlookFolder> PullCalendars()
        {
            try
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

                    Log.Write( "Pulled a list of Outlook Calendars" );
                }

                return m_folderList;
            } catch ( Exception ex )
            {
                Log.Write( ex );
                MessageBox.Show( "There was an error when trying to pull a list of Calendars available in Outlook.",
                    "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }

            return null;
        }

        public List<CalendarItem> PullListOfAppointments()
        {
            var calList = new List<CalendarItem>();
            try
            {
                Log.Write( $"Pulling a full list of appointments from Outlook calendar, {m_folder.Name}." );
                Items item = m_folder.Items;

                foreach ( var i in item )
                {
                    var cal = new CalendarItem();
                    cal.LoadFromOutlookAppointment( (AppointmentItem) i );
                    if ( !calList.Exists( x => x.ID.Equals( cal.ID ) ) )
                        calList.Add( cal );
                }

                Log.Write( "Completed pulling Outlook appointments." );

            } catch ( NullReferenceException ex )
            {
                Log.Write( ex );
            }

            return calList;
        }

        public List<CalendarItem> PullListOfAppointmentsByDate( DateTime startDate, DateTime endDate )
        {
            var calList = new List<CalendarItem>();

            try
            {
                Log.Write( $"Pulling a full list of appointments from Outlook calendar, {m_folder.Name}, by date. " );
                Items item = GetAppointmentsInRange( m_folder, startDate, endDate );

                foreach ( var i in item )
                {
                    var cal = new CalendarItem();
                    cal.LoadFromOutlookAppointment( (AppointmentItem) i );
                    if ( !calList.Exists( x => x.ID.Equals( cal.ID ) ) )
                        calList.Add( cal );
                }
                Log.Write( "Completed pulling Outlook appointments." );
            } catch ( NullReferenceException ex )
            {
                Log.Write( ex );
            }

            return calList;
        }

        public CalendarItem FindEvent( string gid )
        {
            Log.Write( "Looking up an Outlook appointment using the ID" );
            foreach ( var i in m_folder.Items )
            {
                var cal = new CalendarItem();
                cal.LoadFromOutlookAppointment( (AppointmentItem) i );
                if ( cal.ID.Equals( gid ) )
                {
                    Log.Write( $"Found an Outlook appointment with ID, {gid}" );
                    return cal;
                }
            }

            Log.Write( $"Unable to find an Outlook appointment with ID, {gid}" );
            return null;
        }

        public CalendarItem FindEventByEntryId( string entryId )
        {

            try
            {
                Log.Write( "Looking up an Outlook appointment using the EntryId" );
                var appt = (AppointmentItem)Application.Session.GetItemFromID( entryId, m_folder.Store.StoreID );
                if ( appt != null )
                {
                    var calEvent = new CalendarItem();
                    calEvent.LoadFromOutlookAppointment( appt );

                    Log.Write( $"Found an Outlook appointment with EntryID, {entryId}" );
                    return calEvent;
                }
            } catch ( COMException ex )
            {
                Log.Write( ex );
            }
            Log.Write( $"Unable to find an Outlook appointment with EntryID, {entryId}" );
            return null;
        }

        public void UpdateAppointment( CalendarItem ev )
        {
            try
            {
                if ( ev.Recurrence != null && Settings.Default.UpdateRecurrance )
                {
                    Log.Write( $"Updating recurring Outlook appointment, {ev.Subject}" );
                    // Create the filter to find the event. This is done by either the subject and start date or the ID
                    var filter = ( ev.Action.HasFlag( CalendarItemAction.GeneratedId ))
                        ? ( "[Subject]='" + ev.Subject + "'" )
                        : "[ID] = '" + ev.ID + "'";

                    var appointmentItems = m_folder.Items.Restrict( filter );
                    appointmentItems.IncludeRecurrences = true;
                    appointmentItems.Sort( "[Start]", Type.Missing );

                    foreach ( AppointmentItem appointmentItem in appointmentItems )
                    {
                        if ( appointmentItem.RecurrenceState != OlRecurrenceState.olApptMaster )
                        {
                            var a = appointmentItem.Parent as AppointmentItem;
                            ev.GetOutlookAppointment( a );
                            continue;
                        }

                        ev.GetOutlookAppointment( appointmentItem );
                        appointmentItem.Save();

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
            } catch ( Exception ex )
            {
                Log.Write( ex );
            }

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
