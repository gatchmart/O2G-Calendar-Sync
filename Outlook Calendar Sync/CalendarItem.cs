using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using TimeZone = System.TimeZone;

namespace Outlook_Calendar_Sync {

    public sealed class Attendee : IEquatable<Attendee> {
        public string Name { get; set; }
        public string Email { get; set; }
        public bool Required { get; set; }

        public Attendee() {
            Name = "";
            Email = "";
            Required = false;
        }

        public Attendee( string name = "", string email = "", bool required = false ) {
            Name = name;
            Email = email;
            Required = required;
        }

        /// <summary>
        /// Creates and Attendee
        /// </summary>
        /// <param name="outlookAttendee">The Outlook attendee string</param>
        /// <param name="required">is the attendee required</param>
        public Attendee( string outlookAttendee, bool required = false ) {
            // Gregory Atchley-Martin ( gatchmart@gmail.com )
            var strs = outlookAttendee.Split( '(' );

            Name = strs[0].Trim();
            Email = ( strs.Length > 1 ) ? strs[1].TrimEnd( ')' ).Trim() : "";
            Required = required;
        }

        public bool Equals( Attendee other ) {
            return Name.Equals( other.Name ) && Email.Equals( other.Email ) && Required == other.Required;
        }
    }

    [Flags]
    public enum CalendarItemAction {
        Nothing = 0,
        GoogleUpdate = 1,
        OutlookUpdate = 2,
        GoogleAdd = 4,
        OutlookAdd = 8,
        GeneratedId = 16,
        ContentsEqual = 32,
        GoogleDelete = 64,
        OutlookDelete = 128
    }
    
    [Flags]
    public enum CalendarItemChanges {
        Nothing = 0,
        StartDate = 1,
        EndDate = 2,
        Location = 4,
        Body = 8,
        Subject = 16,
        StartTimeZone = 32,
        EndTimeZone = 64,
        ReminderTime = 128,
        Attendees = 256,
        Recurrence = 512,
        CalId = 1024
    }

    public sealed class CalendarItem : IEquatable<CalendarItem> {
        // The date time format string used to properly format the start and end dates
        internal const string DateTimeFormatString = "yyyy-MM-ddTHH:mm:sszzz";

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
        /// The unique Google event ID
        /// </summary>
        public string ID { get; set; }

        /// <summary>
        /// This ID is used for recurring events in Google. It will allow you to delete events.
        /// </summary>
        public string iCalID { get; set; }

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

        public bool IsAllDayEvent { get; set; }

        public bool IsFirstOccurence => Recurrence != null && Recurrence.GetPatternStartTimeWithHours().Equals( Start );

        public bool IsLastOccurence => Recurrence != null && Recurrence.GetPatternEndTimeWithHours().Equals( End );

        public bool ContainsOutlookAppointmentItem => m_outlookAppointment != null;

        #endregion

        private AppointmentItem m_outlookAppointment;

        private bool m_isUsingDefaultReminders;

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
        }

        public AppointmentItem GetOutlookAppointment() {
            return GetOutlookAppointment( m_outlookAppointment );
        }

        /// <summary>
        /// Gets the Outlook AppointmentItem representation of this CalendarItem
        /// </summary>
        /// <param name="item">The AppointmentItem to edit</param>
        /// <returns>An Outlook AppointmentItem representation of this CalendarItem</returns>
        public AppointmentItem GetOutlookAppointment( AppointmentItem item ) {
            try {
                item.Start = DateTime.Parse( Start );
                item.End = DateTime.Parse( End );
                item.Location = Location;
                item.Body = Body;
                item.Subject = Subject;
                item.AllDayEvent = IsAllDayEvent;

                UserProperties prop = item.UserProperties;
                var p = prop.Find( "ID", true );
                var d = prop.Find( "iCalID", true );

                // Check to see if we found either user property, if not add it
                if ( p == null )
                    p = prop.Add( "ID", OlUserPropertyType.olText );

                if ( d == null )
                    d = prop.Add( "iCalID", OlUserPropertyType.olText );

                // Finally set the UserProperty values
                p.Value = ID;
                d.Value = iCalID;

                if ( !m_isUsingDefaultReminders ) {
                    item.ReminderMinutesBeforeStart = ReminderTime;
                    item.ReminderSet = true;
                }

                if ( Recurrence != null ) {
                    var pattern = item.GetRecurrencePattern();
                    Recurrence.GetOutlookPattern( ref pattern );
                }

                foreach ( var attendee in Attendees ) {
                    var recipt = item.Recipients.Add( attendee.Name );
                    recipt.AddressEntry.Address = attendee.Email;
                    recipt.Type = attendee.Required
                        ? (int) OlMeetingRecipientType.olRequired
                        : (int) OlMeetingRecipientType.olOptional;
                }

                return item;
            } catch ( COMException ex ) {
                Console.WriteLine(ex);
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
                ? new EventDateTime {
                    Date = Start.Substring( 0, 10 )
                }
                : new EventDateTime {
                    DateTime = DateTime.Parse( Start ),
                    TimeZone = StartTimeZone
                };

            var en = IsAllDayEvent
                ? new EventDateTime {
                    Date = End.Substring( 0, 10 )
                }
                : new EventDateTime {
                    DateTime = DateTime.Parse( End ),
                    TimeZone = EndTimeZone
                };

            var e = new Event {
                Summary = Subject, Description = Body, Start = st, End = en, Location = Location
            };

            if ( !string.IsNullOrEmpty( ID ) )
                e.Id = ID;

            if ( !m_isUsingDefaultReminders ) {
                if ( e.Reminders == null )
                    e.Reminders = new Event.RemindersData();

                e.Reminders.UseDefault = false;
                e.Reminders.Overrides = new List<EventReminder> {
                    new EventReminder {
                        Method = "email",
                        Minutes = ReminderTime
                    }
                };
            }

            if ( Recurrence != null ) 
                e.Recurrence = new string[] { Recurrence.GetGoogleRecurrenceString() };

            var att = new EventAttendee[Attendees.Count];
            for ( int i = 0; i < Attendees.Count; i++ ) {
                var a = Attendees[i];
                att[i] = new EventAttendee {
                    Email = a.Email, Optional = !a.Required, DisplayName = a.Name
                };
            }

            e.Attendees = att;

            return e;
        }

        /// <summary>
        /// Fills this CalendarItem with the data from a Google Event object
        /// </summary>
        /// <param name="ev">The Google Event to use</param>
        public void LoadFromGoogleEvent( Event ev ) {
            Start = ev.Start.DateTimeRaw ?? ev.Start.Date;
            End = ev.End.DateTimeRaw ?? ev.End.Date;
            Location = ev.Location;
            Body = ev.Description;
            Subject = ev.Summary;
            StartTimeZone = ev.Start.TimeZone;
            EndTimeZone = ev.End.TimeZone;
            ID = ev.Id;
            iCalID = ev.ICalUID;

            IsAllDayEvent = ( ev.Start.DateTimeRaw == null && ev.End.DateTimeRaw == null );
           
            if ( ev.Reminders.Overrides != null )
                ReminderTime = ev.Reminders.Overrides.First( x => x.Method == "email" || x.Method == "popup" ).Minutes ?? 0;
            else {
                m_isUsingDefaultReminders = true;
                ReminderTime = 30;
            }

            if ( ev.Recurrence != null ) 
                Recurrence = new Recurrence( ev.Recurrence[0], this );

            if ( ev.Attendees != null ) {
                foreach ( var eventAttendee in ev.Attendees ) {
                    Attendees.Add( new Attendee {
                        Name = eventAttendee.DisplayName, Email = eventAttendee.Email, Required = !( eventAttendee.Optional ?? true )
                    } );
                }
            }
        }

        /// <summary>
        /// Fills this CalendarItem with the data from an Outlook AppointmentItem object
        /// </summary>
        /// <param name="item">The Outlook AppointmentItem to use</param>
        /// <param name="createID">Specify if you need to create and ID.</param>
        public void LoadFromOutlookAppointment( AppointmentItem item, bool createID = true ) {
            // Store a copy of the Outlook Appointment
            m_outlookAppointment = item;

            Start = item.Start.ToString( DateTimeFormatString );
            End = item.End.ToString( DateTimeFormatString );
            Body = item.Body;
            Subject = item.Subject;
            Location = item.Location;
            ReminderTime = item.ReminderMinutesBeforeStart;
            m_isUsingDefaultReminders = !item.ReminderSet;
            StartTimeZone = TimeZoneConverter.WindowsToIana( item.StartTimeZone.ID );
            EndTimeZone = TimeZoneConverter.WindowsToIana( item.EndTimeZone.ID );
            IsAllDayEvent = item.AllDayEvent;

            // Try to find the ID and iCalID from the UserProperties
            var idProp = item.UserProperties.Find( "ID", true );
            var icalProp = item.UserProperties.Find( "iCalID", true );

            // Check to make sure they are not null and set their values
            if ( idProp != null && icalProp != null ) {
                ID = idProp.Value;
                iCalID = icalProp.Value;
            } else if ( idProp != null ) {
                ID = idProp.Value;
            } else { 
                // If both UserProperties are null create an ID
                if ( createID ) {
                    Action |= CalendarItemAction.OutlookUpdate;
                    Action |= CalendarItemAction.GeneratedId;
                    ID = Guid.NewGuid().ToString().Replace( "-", "" );
                }
            }

            // Check if the event is recurring
            if ( item.IsRecurring ) {
                var recure = item.GetRecurrencePattern();
                Recurrence = new Recurrence( recure );
                Recurrence.AdjustRecurrenceOutlookPattern( item.Start, item.End );
            }

            // Add attendees
            if ( !string.IsNullOrEmpty( item.OptionalAttendees ) ) {
                if ( item.OptionalAttendees.Contains( ";" ) ) {
                    var attendees = item.OptionalAttendees.Split( ';' );
                    foreach ( var attendee in attendees )
                        Attendees.Add( new Attendee( attendee, false ) );
                } else
                    Attendees.Add( new Attendee( item.OptionalAttendees, false ) );
            }

            // Grab the required attendees.
            if ( !string.IsNullOrEmpty( item.RequiredAttendees ) ) {
                if ( item.RequiredAttendees.Contains( ";" ) ) {
                    var attendees = item.RequiredAttendees.Split( ';' );
                    foreach ( var attendee in attendees )
                        Attendees.Add( new Attendee( attendee, true ) );
                } else
                    Attendees.Add( new Attendee( item.RequiredAttendees, true ) );
            }
        }

        public bool Equals( CalendarItem other ) {
            if ( !Action.HasFlag( CalendarItemAction.ContentsEqual ) && !string.IsNullOrEmpty( ID ) && !string.IsNullOrEmpty( other.ID ) )
                return ID.Equals( other.ID );

            Changes = CalendarItemChanges.Nothing;

            GetCalendarDifferences( other );

            return Changes == CalendarItemChanges.Nothing;
        }

        /// <summary>
        /// Gets the differences between two CalendarItems
        /// </summary>
        /// <param name="other">The other CalendarItem to compare</param>
        public void GetCalendarDifferences( CalendarItem other ) {

            var s = DateTime.Parse( Start ).ToUniversalTime();
            var e = DateTime.Parse( End ).ToUniversalTime();
            var ss = DateTime.Parse( other.Start ).ToUniversalTime();
            var ee = DateTime.Parse( other.End ).ToUniversalTime();

            if ( !s.Equals( ss ) )
                Changes |= CalendarItemChanges.StartDate;

            if ( !e.Equals( ee ) )
                Changes |= CalendarItemChanges.EndDate;

            if ( !Subject.Equals( other.Subject ))
                Changes |= CalendarItemChanges.Subject;

            if ( !string.IsNullOrEmpty( Location ) && !string.IsNullOrEmpty( other.Location ) )
                if ( !Location.Equals( other.Location ) )
                    Changes |= CalendarItemChanges.Location;

            if ( !string.IsNullOrEmpty( Body ) && !string.IsNullOrEmpty( other.Body ) )
                if ( !Body.Equals( other.Body ) )
                    Changes |= CalendarItemChanges.Body;

            if ( !string.IsNullOrEmpty( StartTimeZone ) && !string.IsNullOrEmpty(other.StartTimeZone ) )
                if ( !StartTimeZone.Equals( other.StartTimeZone ) )
                    Changes |= CalendarItemChanges.StartTimeZone;

            if ( !string.IsNullOrEmpty( EndTimeZone ) && !string.IsNullOrEmpty( other.EndTimeZone ) )
                if ( !EndTimeZone.Equals( other.EndTimeZone ) )
                    Changes |= CalendarItemChanges.EndTimeZone;

            // TODO: Update to ignore default outlook reminder of 18 hours.
            if ( ReminderTime >= 0 && other.ReminderTime >= 0 )
                if ( ReminderTime != other.ReminderTime )
                    Changes |= CalendarItemChanges.ReminderTime;

            if ( Attendees.Count > 0 && other.Attendees.Count > 0 )
                if ( !Attendees.All( other.Attendees.Contains ) )
                    Changes |= CalendarItemChanges.Attendees;

            if ( Recurrence != null && other.Recurrence != null ) {
                if ( Recurrence.GetGoogleRecurrenceString().Equals( other.Recurrence.GetGoogleRecurrenceString() ) )
                    Changes |= CalendarItemChanges.Recurrence;
            } else if ( Recurrence != null || other.Recurrence != null)
                Changes |= CalendarItemChanges.Recurrence;

            if ( !iCalID.Equals( other.iCalID ) )
                Changes |= CalendarItemChanges.CalId;

            other.Changes = Changes;
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
            builder.AppendLine( "\tID: " + ID );
            builder.AppendLine( "\tiCalID: " + iCalID );
            builder.AppendLine( "\tReminder Time: " + ReminderTime );
            builder.AppendLine( "\tUsing Default Reminder: " + ( m_isUsingDefaultReminders ? "Yes" : "No" ) );

            if ( Attendees != null && Attendees.Count > 0 ) {
                builder.AppendLine( "\tAttendees:" );

                foreach ( var attendee in Attendees )
                    builder.AppendLine( "\t\t" + attendee.Name + ", " + attendee.Email + ", " +
                                        ( attendee.Required ? "Required" : "Not Required" ) );
            }

            if ( Recurrence != null ) {
                builder.AppendLine( "\tRecurrence:" );
                builder.AppendLine( Recurrence.ToString() );
            }

            if ( Changes != CalendarItemChanges.Nothing ) {
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

            if ( Action != CalendarItemAction.Nothing ) {
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

    }
}
