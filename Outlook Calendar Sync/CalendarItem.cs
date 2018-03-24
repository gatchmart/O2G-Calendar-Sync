using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Serialization;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using Outlook_Calendar_Sync.Enums;

namespace Outlook_Calendar_Sync {

    /// <summary>
    /// The CalendarItem class is an intermediate class between the Google Event class and the Outlook AppointmentItem interface.
    /// This class allows you to easily convert between AppointmentItems and Events and, provides a common storage class for the data.
    /// </summary>
    [Serializable]
    public sealed class CalendarItem : IEquatable<CalendarItem> {
        // The date time format string used to properly format the start and end dates
        internal const string DateTimeFormatString = "yyyy-MM-ddTHH:mm:sszzz";

        internal const int DEFAULT_REMINDER_TIME = 30;

        #region Properties

        /// <summary>
        /// The current action needed to be perform on this event.
        /// </summary>
        public CalendarItemAction Action { get; set; }

        /// <summary>
        /// The current changes to the event
        /// </summary>
        public CalendarItemChanges Changes { get; set; }

        /// <summary>
        /// The start date and time
        /// </summary>
        public string Start { get; set; }

        /// <summary>
        /// The end date and time
        /// </summary>
        public string End { get; set; }

        /// <summary>
        /// The location of the event
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// The body/description of the event
        /// </summary>
        public string Body { get; set; }

        /// <summary>
        /// The subject/summery of the event
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// The name of the current time zone
        /// </summary>
        public string StartTimeZone { get; set; }

        /// <summary>
        /// The name of the current time zone
        /// </summary>
        public string EndTimeZone { get; set; }

        /// <summary>
        /// The start time of the event in its correct time zone. (Used when comparing events)
        /// </summary>
        public string StartTimeInTimeZone { get; set; }

        /// <summary>
        /// The end time of the event in its correct time zone. (Used when comparing events)
        /// </summary>
        public string EndTimeInTimeZone { get; set; }

        /// <summary>
        /// This is the amount of time before the event to show a reminder in minutes.
        /// </summary>
        public int ReminderTime { get; set; }

        /// <summary>
        /// The list of attendees for this event
        /// </summary>
        public List<Attendee> Attendees { get; set; }

        /// <summary>
        /// The Recurrence object that is only used if the CalendarItem repeats
        /// </summary>
        public Recurrence Recurrence { get; set; }

        public Identifier CalendarItemIdentifier { get; set; }

        /// <summary>
        /// Is this event an all day event?
        /// </summary>
        public bool IsAllDayEvent { get; set; }

        /// <summary>
        /// Is this instance the first appointment in a recurrance pattern?
        /// </summary>
        public bool IsFirstOccurence => Recurrence != null && Recurrence.GetPatternStartTimeWithHours().Equals( Start );

        /// <summary>
        /// Is this instance the last appointment in a recurrance pattern?
        /// </summary>
        public bool IsLastOccurence => Recurrence != null && Recurrence.GetPatternEndTimeWithHours().Equals( End );

        #endregion

        private bool m_isUsingDefaultReminders;

        private string m_source;

        public CalendarItem() {
            Start = "";
            End = "";
            Location = "";
            Body = "";
            Subject = "";
            ReminderTime = 0;
            Attendees = new List<Attendee>();
            Recurrence = null;
            m_isUsingDefaultReminders = false;
            Action = CalendarItemAction.Nothing;
            Changes = CalendarItemChanges.Nothing;
            IsAllDayEvent = false;
            CalendarItemIdentifier = new Identifier();
        }

        /// <summary>
        /// Gets the Outlook AppointmentItem representation of this CalendarItem
        /// </summary>
        /// <param name="item">The AppointmentItem to edit</param>
        /// <returns>An Outlook AppointmentItem representation of this CalendarItem</returns>
        public AppointmentItem GetOutlookAppointment( AppointmentItem item ) {
            try
            {
                item.Start = DateTime.Parse( Start );
                item.End = DateTime.Parse( End );
                item.Location = Location;
                item.Body = Body;
                item.Subject = Subject;
                item.AllDayEvent = IsAllDayEvent;

                Microsoft.Office.Interop.Outlook.TimeZone startTz = OutlookSync.Syncer.CurrentApplication.TimeZones[TimeZoneConverter.IanaToWindows( StartTimeZone )];
                Microsoft.Office.Interop.Outlook.TimeZone endTz = OutlookSync.Syncer.CurrentApplication.TimeZones[TimeZoneConverter.IanaToWindows( EndTimeZone )];

                item.StartTimeZone = startTz;
                item.EndTimeZone = endTz;

                if ( !m_isUsingDefaultReminders )
                {
                    item.ReminderMinutesBeforeStart = ReminderTime;
                    item.ReminderSet = true;
                }

                if ( Recurrence != null )
                {
                    var pattern = item.GetRecurrencePattern();
                    Recurrence.GetOutlookPattern( ref pattern, IsAllDayEvent );
                }

                foreach ( var attendee in Attendees )
                {
                    var recipt = item.Recipients.Add( attendee.Name );
                    bool result = recipt.Resolve();

                    if ( !result )
                    {
                        ContactItem contact = new ContactItem( attendee.Name, attendee.Email );
                        result = contact.CreateContact();
                    }

                    result &= recipt.Resolve();
                    if ( result )
                    {
                        recipt.AddressEntry.Address = attendee.Email;
                        recipt.Type = attendee.Required
                            ? (int)OlMeetingRecipientType.olRequired
                            : (int)OlMeetingRecipientType.olOptional;
                    }
                }

                return item;
            } catch ( COMException ex )
            {
                Log.Write( ex );
                MessageBox.Show(
                    "CalendarItem: There has been an error when trying to create a outlook appointment item from a CalendarItem.", "Unknown Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }

            return null;
        }

        /// <summary>
        /// Gets the Google Event representation of this CalendarItem
        /// </summary>
        /// <returns>Google Event of the CalendarItem</returns>
        public Event GetGoogleCalendarEvent() {
            var st = IsAllDayEvent
                ? new EventDateTime
                {
                    Date = Start.Substring( 0, 10 )
                }
                : new EventDateTime
                {
                    DateTime = DateTime.Parse( Start ),
                    TimeZone = StartTimeZone
                };

            var en = IsAllDayEvent
                ? new EventDateTime
                {
                    Date = End.Substring( 0, 10 )
                }
                : new EventDateTime
                {
                    DateTime = DateTime.Parse( End ),
                    TimeZone = EndTimeZone
                };

            var e = new Event
            {
                Summary = Subject,
                Description = Body,
                Start = st,
                End = en,
                Location = Location
            };

            if ( !string.IsNullOrEmpty( CalendarItemIdentifier.GoogleId ) )
            {
                e.Id = CalendarItemIdentifier.GoogleId.ToLower();
            }

            if ( !string.IsNullOrEmpty( CalendarItemIdentifier.GoogleICalUId ) )
            {
                e.ICalUID = CalendarItemIdentifier.GoogleICalUId;
            }

            if ( !m_isUsingDefaultReminders )
            {
                if ( e.Reminders == null )
                    e.Reminders = new Event.RemindersData();

                e.Reminders.UseDefault = false;
                e.Reminders.Overrides = new List<EventReminder> {
                    new EventReminder {
                        Method = "popup",
                        Minutes = ReminderTime
                    }
                };
            }

            if ( Recurrence != null )
                e.Recurrence = new string[] { Recurrence.GetGoogleRecurrenceString() };

            var att = new EventAttendee[Attendees.Count];
            for ( int i = 0; i < Attendees.Count; i++ )
            {
                var a = Attendees[i];
                att[i] = new EventAttendee
                {
                    Email = a.Email,
                    Optional = !a.Required,
                    DisplayName = a.Name
                };
            }

            e.Attendees = att;

            return e;
        }

        /// <summary>
        /// Fills this CalendarItem with the data from a Google Event object
        /// </summary>
        /// <param name="ev">The Google Event to use</param>
        public void LoadFromGoogleEvent( Event ev )
        {
            m_source = "Google";
            Start = ev.Start.DateTimeRaw ?? ev.Start.Date;
            End = ev.End.DateTimeRaw ?? ev.End.Date;
            Location = ev.Location;
            Body = ev.Description;
            Subject = ev.Summary;

            // Ensure time zone is properly setup.
            StartTimeZone = ev.Start.TimeZone ?? TimeZoneConverter.WindowsToIana( TimeZoneInfo.Local.Id );
            EndTimeZone = ev.End.TimeZone ?? TimeZoneConverter.WindowsToIana( TimeZoneInfo.Local.Id );
            StartTimeInTimeZone = Start;
            EndTimeInTimeZone = End;

            var id = Archiver.Instance.FindIdentifier( ev.Id );
            if ( id == null )
                CalendarItemIdentifier.GoogleId = ev.Id;
            else
                CalendarItemIdentifier = id;

            if ( string.IsNullOrEmpty( CalendarItemIdentifier.GoogleICalUId ) )
                CalendarItemIdentifier.GoogleICalUId = ev.ICalUID;

            IsAllDayEvent = ( ev.Start.DateTimeRaw == null && ev.End.DateTimeRaw == null );

            if ( ev.Reminders.Overrides != null )
            {
                EventReminder reminder = ev.Reminders.Overrides.FirstOrDefault( x => x.Method == "popup" );

                ReminderTime = reminder?.Minutes ?? DEFAULT_REMINDER_TIME;
            }
            else
            {
                m_isUsingDefaultReminders = true;
                ReminderTime = DEFAULT_REMINDER_TIME;
            }

            if ( ev.Recurrence != null )
                Recurrence = new Recurrence( ev.Recurrence[0], this );

            if ( ev.Attendees != null )
            {
                foreach ( var eventAttendee in ev.Attendees )
                {
                    if ( string.IsNullOrEmpty( eventAttendee.DisplayName ) || string.IsNullOrEmpty( eventAttendee.Email ) )
                        continue;

                    Attendees.Add( new Attendee
                    {
                        Name = eventAttendee.DisplayName ?? "",
                        Email = eventAttendee.Email ?? "",
                        Required = !( eventAttendee.Optional ?? true )
                    } );
                }
            }

            CalendarItemIdentifier.EventHash = EventHasher.GetHash( this );
        }

        /// <summary>
        /// Fills this CalendarItem with the data from an Outlook AppointmentItem object
        /// </summary>
        /// <param name="item">The Outlook AppointmentItem to use</param>
        /// <param name="createID">Specify if you need to create and ID.</param>
        public void LoadFromOutlookAppointment( AppointmentItem item, bool createID = true ) {
            m_source = "Outlook";

            Start = item.Start.ToString( DateTimeFormatString );
            End = item.End.ToString( DateTimeFormatString );
            Body = item.Body;
            Subject = item.Subject;
            Location = item.Location;
            m_isUsingDefaultReminders = !item.ReminderSet;
            ReminderTime = m_isUsingDefaultReminders ? DEFAULT_REMINDER_TIME : item.ReminderMinutesBeforeStart;
            StartTimeZone = TimeZoneConverter.WindowsToIana( item.StartTimeZone.ID );
            EndTimeZone = TimeZoneConverter.WindowsToIana( item.EndTimeZone.ID );
            IsAllDayEvent = item.AllDayEvent;
            StartTimeInTimeZone = item.StartInStartTimeZone.ToString( DateTimeFormatString );
            EndTimeInTimeZone = item.EndInEndTimeZone.ToString( DateTimeFormatString );
            
            string entryId = null;
            string globalId = null;
            bool useParent = false;

            // This ensures that if we grab a occurence of a recurring appointment we use the proper global ID.
            // You must use the MasterAppointment's global ID since that is what I track.
            if ( item.IsRecurring )
            {
                if ( item.RecurrenceState != OlRecurrenceState.olApptMaster )
                {
                    if ( item.Parent is AppointmentItem parent )
                    {
                        entryId = parent.EntryID;
                        globalId = parent.GlobalAppointmentID;
                        useParent = true;
                    }
                }
            }

            // Outlook is fucking stupid and changes the GlobalAppointmentID everytime it restarts but doesn't change the EntryID so use one or the other.
            var id = Archiver.Instance.FindIdentifier( useParent ? entryId : item.EntryID ) ?? Archiver.Instance.FindIdentifier( useParent ? globalId : item.GlobalAppointmentID );
            if ( id == null )
            {
                CalendarItemIdentifier.OutlookEntryId = useParent ? entryId : item.EntryID;
                CalendarItemIdentifier.OutlookGlobalId = useParent ? globalId : item.GlobalAppointmentID;
            }
            else
                CalendarItemIdentifier = id;

            if ( string.IsNullOrEmpty( CalendarItemIdentifier.GoogleId ) )
            {
                if ( createID )
                {
                    CalendarItemIdentifier.GoogleId = GuidCreator.Create();
                }
            }

            // Check if the event is recurring
            if ( item.IsRecurring )
            {
                var recure = item.GetRecurrencePattern();
                Recurrence = new Recurrence( recure );
                //Recurrence.AdjustRecurrenceOutlookPattern( item.Start, item.End );
            }

            // Add attendees
            if ( !string.IsNullOrEmpty( item.OptionalAttendees ) )
            {
                if ( item.OptionalAttendees.Contains( ";" ) )
                {
                    var attendees = item.OptionalAttendees.Split( ';' );
                    foreach ( var attendee in attendees )
                    {
                        ContactItem contact = ContactItem.GetContactItem( attendee );
                        Attendees.Add( contact != null
                            ? new Attendee( contact.Name, contact.Email, false )
                            : new Attendee( attendee, false ) );
                    }
                }
                else
                {
                    ContactItem contact = ContactItem.GetContactItem( item.OptionalAttendees );
                    Attendees.Add( contact != null
                        ? new Attendee( contact.Name, contact.Email, true )
                        : new Attendee( item.OptionalAttendees, true ) );
                }
                    
            }

            // Grab the required attendees.
            if ( !string.IsNullOrEmpty( item.RequiredAttendees ) )
            {
                if ( item.RequiredAttendees.Contains( ";" ) )
                {
                    var attendees = item.RequiredAttendees.Split( ';' );
                    foreach ( var attendee in attendees )
                    {
                        ContactItem contact = ContactItem.GetContactItem( attendee );
                        Attendees.Add( contact != null
                            ? new Attendee( contact.Name, contact.Email, true )
                            : new Attendee( attendee, true ) );
                    }
                }
                else
                {
                    ContactItem contact = ContactItem.GetContactItem( item.RequiredAttendees );
                    Attendees.Add( contact != null
                        ? new Attendee( contact.Name, contact.Email, true )
                        : new Attendee( item.RequiredAttendees, true ) );
                }
            }

            CalendarItemIdentifier.EventHash = EventHasher.GetHash( this );
        }

        public bool Equals( CalendarItem other )
        {
            if ( other == null )
                return false;

            var idEqual =
                EventHasher.Equals( CalendarItemIdentifier.EventHash, other.CalendarItemIdentifier.EventHash );

            if ( !idEqual && !string.IsNullOrEmpty( CalendarItemIdentifier.GoogleId ) &&
                    !string.IsNullOrEmpty( other.CalendarItemIdentifier.GoogleId ) )
                idEqual |= CalendarItemIdentifier.GoogleId.Equals( other.CalendarItemIdentifier.GoogleId );

            if ( !idEqual && !string.IsNullOrEmpty( CalendarItemIdentifier.OutlookEntryId ) &&
                    !string.IsNullOrEmpty( other.CalendarItemIdentifier.OutlookEntryId ) )
                idEqual |= CalendarItemIdentifier.OutlookEntryId.Equals( other.CalendarItemIdentifier.OutlookEntryId );

            return idEqual;
        }

        public bool IsContentsEqual( CalendarItem item )
        {
            Changes = CalendarItemChanges.Nothing;
            GetCalendarDifferences( item );
            return Changes == CalendarItemChanges.Nothing;
        }

        public override string ToString() {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine( "----------------" + Subject + "----------------" );
            builder.AppendLine( "\tStart: " + Start );
            builder.AppendLine( "\tEnd: " + End );
            builder.AppendLine( "\tStart Time Zone: " + StartTimeZone );
            builder.AppendLine( "\tEnd Time Zone: " + EndTimeZone );
            builder.AppendLine( "\tLocation: " + Location );
            builder.AppendLine( "\tBody: " + Body );
            builder.AppendLine( CalendarItemIdentifier.ToString() );
            builder.AppendLine( "\tReminder Time: " + ReminderTime );
            builder.AppendLine( "\tUsing Default Reminder: " + ( m_isUsingDefaultReminders ? "Yes" : "No" ) );
            builder.AppendLine( "\tIs All Day Event: " + ( IsAllDayEvent ? "Yes" : "No" ) );

            if ( Attendees != null && Attendees.Count > 0 )
            {
                builder.AppendLine( "\tAttendees: " );

                foreach ( var attendee in Attendees )
                    builder.AppendLine( "\t\t" + attendee.Name + ", " + attendee.Email + ", " +
                                        ( attendee.Required ? "Required" : "Not Required" ) );
            }

            if ( Recurrence != null )
            {
                builder.AppendLine( "\tRecurrence: " );
                builder.AppendLine( Recurrence.ToString() );
            }

            if ( Changes != CalendarItemChanges.Nothing )
            {
                builder.Append( "\tChanges: " );
                if ( Changes.HasFlag( CalendarItemChanges.StartDate ) )
                    builder.Append( "Change Start Date | " );

                if ( Changes.HasFlag( CalendarItemChanges.EndDate ) )
                    builder.Append( "Change End Date | " );

                if ( Changes.HasFlag( CalendarItemChanges.Location ) )
                    builder.Append( "Change Location | " );

                if ( Changes.HasFlag( CalendarItemChanges.Body ) )
                    builder.Append( "Change Body | " );

                if ( Changes.HasFlag( CalendarItemChanges.Subject ) )
                    builder.Append( "Change Subject | " );

                if ( Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                    builder.Append( "Change Start Time Zone | " );

                if ( Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                    builder.Append( "Change End Time Zone | " );

                if ( Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                    builder.Append( "Change Reminder Time | " );

                if ( Changes.HasFlag( CalendarItemChanges.Attendees ) )
                    builder.Append( "Change Attendees | " );

                if ( Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                    builder.Append( "Change Recurrence | " );

                builder.Remove( builder.Length - 2, 2 );
                builder.AppendLine();
            }

            if ( Action != CalendarItemAction.Nothing )
            {
                builder.Append( "\tAction: " );
                if ( Action.HasFlag( CalendarItemAction.ContentsEqual ) )
                    builder.Append( "Action Contents Equal | " );
                if ( Action.HasFlag( CalendarItemAction.GeneratedId ) )
                    builder.Append( "Action Generated ID | " );
                if ( Action.HasFlag( CalendarItemAction.GoogleAdd ) )
                    builder.Append( "Action Google Add | " );
                if ( Action.HasFlag( CalendarItemAction.GoogleDelete ) )
                    builder.Append( "Action Google Delete | " );
                if ( Action.HasFlag( CalendarItemAction.GoogleUpdate ) )
                    builder.Append( "Action Google Update | " );
                if ( Action.HasFlag( CalendarItemAction.OutlookAdd ) )
                    builder.Append( "Action Outlook Add | " );
                if ( Action.HasFlag( CalendarItemAction.OutlookDelete ) )
                    builder.Append( "Action Outlook Delete | " );
                if ( Action.HasFlag( CalendarItemAction.OutlookUpdate ) )
                    builder.Append( "Action Outlook Update | " );
                builder.Remove( builder.Length - 2, 2 );
                builder.AppendLine();
            }

            builder.AppendLine( "----------------" + Subject + "----------------" );

            return builder.ToString();
        }

        public string GetHasherString()
        {
            var builder = new StringBuilder();

            builder.Append( Start );

            builder.Append( "Start: " + Start );
            builder.Append( "End: " + End );
            builder.Append( "Start Time Zone: " + StartTimeZone );
            builder.Append( "End Time Zone: " + EndTimeZone );
            builder.Append( "Location: " + Location );
            builder.Append( "Body: " + Body );
            //builder.Append( CalendarItemIdentifier.ToString() );
            builder.Append( "Reminder Time: " + ReminderTime );
            builder.Append( "Using Default Reminder: " + ( m_isUsingDefaultReminders ? "Yes" : "No" ) );

            if ( Attendees != null && Attendees.Count > 0 )
            {
                builder.Append( "Attendees: " );
                foreach ( var attendee in Attendees )
                    builder.Append( attendee.Name + ", " + attendee.Email + ", " +
                                        ( attendee.Required ? "Required" : "Not Required" ) );
            }

            if ( Recurrence != null )
            {
                builder.Append( "Recurrence: " );
                builder.Append( Recurrence.GetHasherString() );
            }

            if ( Changes != CalendarItemChanges.Nothing )
            {
                builder.Append( "Changes: " );
                if ( Changes.HasFlag( CalendarItemChanges.StartDate ) )
                    builder.Append( "Change Start Date | " );

                if ( Changes.HasFlag( CalendarItemChanges.EndDate ) )
                    builder.Append( "Change End Date | " );

                if ( Changes.HasFlag( CalendarItemChanges.Location ) )
                    builder.Append( "Change Location | " );

                if ( Changes.HasFlag( CalendarItemChanges.Body ) )
                    builder.Append( "Change Body | " );

                if ( Changes.HasFlag( CalendarItemChanges.Subject ) )
                    builder.Append( "Change Subject | " );

                if ( Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                    builder.Append( "Change Start Time Zone | " );

                if ( Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                    builder.Append( "Change End Time Zone | " );

                if ( Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                    builder.Append( "Change Reminder Time | " );

                if ( Changes.HasFlag( CalendarItemChanges.Attendees ) )
                    builder.Append( "Change Attendees | " );

                if ( Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                    builder.Append( "Change Recurrence | " );

                builder.Remove( builder.Length - 2, 2 );
            }

            if ( Action != CalendarItemAction.Nothing )
            {
                builder.Append( "Action: " );
                if ( Action.HasFlag( CalendarItemAction.ContentsEqual ) )
                    builder.Append( "Action Contents Equal | " );
                if ( Action.HasFlag( CalendarItemAction.GeneratedId ) )
                    builder.Append( "Action Generated ID | " );
                if ( Action.HasFlag( CalendarItemAction.GoogleAdd ) )
                    builder.Append( "Action Google Add | " );
                if ( Action.HasFlag( CalendarItemAction.GoogleDelete ) )
                    builder.Append( "Action Google Delete | " );
                if ( Action.HasFlag( CalendarItemAction.GoogleUpdate ) )
                    builder.Append( "Action Google Update | " );
                if ( Action.HasFlag( CalendarItemAction.OutlookAdd ) )
                    builder.Append( "Action Outlook Add | " );
                if ( Action.HasFlag( CalendarItemAction.OutlookDelete ) )
                    builder.Append( "Action Outlook Delete | " );
                if ( Action.HasFlag( CalendarItemAction.OutlookUpdate ) )
                    builder.Append( "Action Outlook Update | " );
                builder.Remove( builder.Length - 2, 2 );
            }

            return builder.ToString();
        }

        /// <summary>
        /// Gets the differences between two CalendarItems
        /// </summary>
        /// <param name="other">The other CalendarItem to compare</param>
        private void GetCalendarDifferences( CalendarItem other )
        {

            var s = DateTime.Parse( !Start.Equals( StartTimeInTimeZone ) ? StartTimeInTimeZone : Start ).ToUniversalTime();
            var e = DateTime.Parse( !End.Equals( EndTimeInTimeZone ) ? EndTimeInTimeZone : End ).ToUniversalTime();
            var ss = DateTime.Parse( !other.Start.Equals( other.StartTimeInTimeZone ) ? other.StartTimeInTimeZone : other.Start ).ToUniversalTime();
            var ee = DateTime.Parse( !other.End.Equals( other.EndTimeInTimeZone ) ? other.EndTimeInTimeZone : other.End ).ToUniversalTime();

            if ( !s.Equals( ss ) )
                Changes |= CalendarItemChanges.StartDate;

            if ( !e.Equals( ee ) )
                Changes |= CalendarItemChanges.EndDate;

            if ( !Subject.Equals( other.Subject ) )
                Changes |= CalendarItemChanges.Subject;

            if ( !string.IsNullOrEmpty( Location ) || !string.IsNullOrEmpty( other.Location ) )
                if ( !IgnoreSpaceAndNewLineEquals( Location, other.Location ) )
                    Changes |= CalendarItemChanges.Location;

            if ( !string.IsNullOrEmpty( Body ) || !string.IsNullOrEmpty( other.Body ) )
                if ( !IgnoreSpaceAndNewLineEquals( Body, other.Body ) )
                    Changes |= CalendarItemChanges.Body;

            if ( !string.IsNullOrEmpty( StartTimeZone ) || !string.IsNullOrEmpty( other.StartTimeZone ) )
                if ( !IgnoreNullEquals( StartTimeZone, other.StartTimeZone ) )
                    Changes |= CalendarItemChanges.StartTimeZone;

            if ( !string.IsNullOrEmpty( EndTimeZone ) || !string.IsNullOrEmpty( other.EndTimeZone ) )
                if ( !IgnoreNullEquals( EndTimeZone, other.EndTimeZone ) )
                    Changes |= CalendarItemChanges.EndTimeZone;

            if ( ReminderTime >= 0 && other.ReminderTime >= 0 )
                if ( ReminderTime != 1080 || other.ReminderTime != 1080 )
                    if ( ReminderTime != other.ReminderTime )
                        Changes |= CalendarItemChanges.ReminderTime;

            if ( Attendees.Count > 0 && other.Attendees.Count > 0 )
                if ( !Attendees.All( other.Attendees.Contains ) )
                    Changes |= CalendarItemChanges.Attendees;

            if ( Recurrence != null && other.Recurrence != null )
            {
                if ( Recurrence.Equals( other.Recurrence ) )
                    Changes |= CalendarItemChanges.Recurrence;
            }
            else if ( Recurrence != null || other.Recurrence != null )
                Changes |= CalendarItemChanges.Recurrence;

            other.Changes = Changes;
        }

        /// <summary>
        /// Perform a string equals ignoring spaces and new lines
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        private bool IgnoreSpaceAndNewLineEquals( string s1, string s2 )
        {
            if ( s1 == null || s2 == null )
                return false;

            var normS1 = Regex.Replace( s1, @"\s+", "" );
            var normS2 = Regex.Replace( s2, @"\s+", "" );

            return string.Equals( normS1, normS2, StringComparison.OrdinalIgnoreCase );
        }

        private bool IgnoreNullEquals( string s1, string s2 )
        {
            if ( s1 == null || s2 == null )
                return false;

            return s1.Equals( s2 );
        }

    }
}
