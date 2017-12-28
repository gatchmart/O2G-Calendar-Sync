﻿using System;
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

        public Application CurrentApplication { get; set; }

        private List<OutlookFolder> m_folderList;
        private DateTime m_lastUpdate;
        private Folder m_folder;

        public void Init( Application application )
        {
            CurrentApplication = application;
            m_folder = CurrentApplication.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar ) as Folder;
            m_lastUpdate = DateTime.MinValue;
            m_folderList = null;
        }

        public void AddAppointment( CalendarItem item )
        {
            try
            {
                // Create the new appointmentItem and move it into the correct folder. (Ensure to grab the new event after it is moved since "move"
                // will create a copy of the item in the new folder and if you just save it, Outlook will create two copies of the appointment.)
                var tempEvent = item.GetOutlookAppointment( (AppointmentItem)CurrentApplication.CreateItem( OlItemType.olAppointmentItem ) );
                var newEvent = tempEvent.Move( m_folder );
                newEvent.Save();
                item.OutlookEntryId = newEvent.EntryID;
                Archiver.Instance.Add( item.ID );

                Log.Write( $"Added {item} to Outlooks Calendar {m_folder.Name}" );

                Marshal.ReleaseComObject( newEvent );
                Marshal.ReleaseComObject( tempEvent );

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
                    // Save the current wokring folder and set the new working folder to the default calendar.
                    var oldFolder = m_folder.EntryID;
                    SetOutlookWorkingFolder( "", true );

                    // Delete the old list
                    if ( m_folderList != null )
                    {
                        m_folderList.Clear();
                        m_folderList = null;
                    }

                    // Go through all the folders and find the calendars
                    m_folderList = new List<OutlookFolder>();
                    foreach ( MAPIFolder folder in CurrentApplication.Session.Folders )
                        GetFolders( folder );

                    // Restore the previous working folder and set the new update time
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
                var appt = (AppointmentItem)CurrentApplication.Session.GetItemFromID( entryId, m_folder.Store.StoreID );
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

                    //// Create the filter to find the event. This is done by either the subject and start date or the ID
                    var filter = ( ev.Action.HasFlag( CalendarItemAction.GeneratedId ) )
                        ? ( "[Subject]='" + ev.Subject + "'" )
                        : "[ID] = '" + ev.ID + "'";

                    var appointmentItems = m_folder.Items.Restrict( filter );
                    appointmentItems.IncludeRecurrences = true;
                    appointmentItems.Sort( "[Start]", Type.Missing );

                    foreach ( AppointmentItem appointmentItem in appointmentItems )
                    {
                        if ( appointmentItem.RecurrenceState == OlRecurrenceState.olApptException )
                            continue;

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

                        if ( string.IsNullOrEmpty( ev.OutlookEntryId ) )
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
                        else
                        {
                            AppointmentItem item =
                                CurrentApplication.Session.GetItemFromID( ev.OutlookEntryId, m_folder.Store.StoreID );

                            if ( item != null )
                            {
                                ev.Action &= ~CalendarItemAction.GeneratedId;
                                ev.Action &= ~CalendarItemAction.OutlookUpdate;
                                ev.GetOutlookAppointment( item );
                                item.Save();
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
                m_folder = CurrentApplication.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar ) as Folder;
            else
                m_folder = CurrentApplication.Session.GetFolderFromID( entryId ) as Folder;
        }

        public MAPIFolder GetDefaultMapiFolder()
        {
            return CurrentApplication.Session.GetDefaultFolder( OlDefaultFolders.olFolderCalendar );
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

        private void GetFolders( MAPIFolder folder )
        {
            foreach ( MAPIFolder child in folder.Folders )
            {
                if ( child.DefaultItemType == OlItemType.olAppointmentItem )
                {
                    m_folderList.Add( new OutlookFolder
                    {
                        Name = child.Name,
                        EntryID = child.EntryID
                    } );
                    Log.Write(
                        $"Found Outlook Folder: {child.Name}, with EntryID: {child.EntryID}" );

                    if ( child.Folders.Count != 0 )
                        GetFolders( child );
                }
            }
        }

    }
}
