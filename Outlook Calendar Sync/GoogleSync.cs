using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util;
using Google.Apis.Util.Store;

namespace Outlook_Calendar_Sync {

    /// <summary>
    /// The Google Syncer is the connection between the application and the Google Calendar API. This is a Singleton class so you need to use GoogleSync.Syncer to access it.
    /// </summary>
    public class GoogleSync : IDisposable {

        /// <summary>
        /// The Singleton instance of GoogleSync
        /// </summary>
        public static GoogleSync Syncer => _instance ?? ( _instance = new GoogleSync() );
        private static GoogleSync _instance;

        private string m_currentCalendar;


        private readonly string[] m_scopes = { CalendarService.Scope.Calendar };
        private readonly string m_workingDirectory =
            Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\";

        private const string ApplicationName = "Outlook Google Calendar Sync";
        private const int DEFAULT_CANCEL_TIME_OUT = 60000;
        private CalendarService m_service;
        private CalendarList m_calendarList;
        private DateTime m_lastUpdate;

        /// <summary>
        /// The default constructor for GoogleSync. This should never be called directly. You should always use GoogleSync.Syncer
        /// </summary>
        public GoogleSync() {
            m_currentCalendar = "primary";
            m_lastUpdate = DateTime.MinValue;
            m_calendarList = null;
        }

        /// <summary>
        /// Add an appointment to the selected Google calendar.
        /// </summary>
        /// <param name="item">The CalendarItem to add</param>
        /// <returns>nothing</returns>
        public void AddAppointment( CalendarItem item ) {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                var e = item.GetGoogleCalendarEvent();
                m_service.Events.Insert( e, m_currentCalendar ).Execute();

                Archiver.Instance.Add( e.Id );
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to add an event to google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        /// <summary>
        /// Pull a complete list of appointments for the set calendar
        /// </summary>
        /// <returns>List of CalendarItems</returns>
        public List<CalendarItem> PullListOfAppointments() {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                List<CalendarItem> items = new List<CalendarItem>();

                // Iterate over the events in the specified calendar
                string pageToken = null;
                do
                {
                    EventsResource.ListRequest list = m_service.Events.List( m_currentCalendar );
                    list.PageToken = pageToken;
                    Events events = list.Execute();
                    List<Event> i = events.Items.ToList();

                    foreach ( var @event in i )
                    {
                        var cal = new CalendarItem();
                        cal.LoadFromGoogleEvent( @event );
                        if ( !items.Exists( x => x.ID.Equals( cal.ID ) ) )
                            items.Add( cal );
                    }

                    pageToken = events.NextPageToken;
                } while ( pageToken != null );

                return items;
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to pull a list of events from google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
                return null;
            }
        }

        /// <summary>
        /// Pulls a list of appointments for the set calendar between a given date range.
        /// </summary>
        /// <param name="startDate">The start of the date range</param>
        /// <param name="endDate">The end of the date range</param>
        /// <returns>List of CalendarItems with the startDate and endDate</returns>
        public List<CalendarItem> PullListOfAppointmentsByDate( DateTime startDate, DateTime endDate ) {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                List<CalendarItem> items = new List<CalendarItem>();

                // Iterate over the events in the specified calendar
                string pageToken = null;
                do
                {
                    EventsResource.ListRequest list = m_service.Events.List( m_currentCalendar );
                    list.TimeMin = startDate;
                    list.TimeMax = endDate;
                    list.PageToken = pageToken;

                    Events events = list.Execute();
                    List<Event> i = events.Items.ToList();

                    foreach ( var @event in i )
                    {
                        var cal = new CalendarItem();
                        cal.LoadFromGoogleEvent( @event );
                        if ( !items.Exists( x => x.ID.Equals( cal.ID ) ) )
                            items.Add( cal );
                    }

                    pageToken = events.NextPageToken;
                } while ( pageToken != null );

                return items;
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to pull a list of events from google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
                return null;
            }
        }

        /// <summary>
        /// Gets an event using the event id
        /// </summary>
        /// <param name="id">The ID of the Event</param>
        /// <returns>A CalendarItem of the event if found or null if not.</returns>
        public CalendarItem PullAppointmentById( string id ) {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                var item = m_service.Events.Get( m_currentCalendar, id ).Execute();
                if ( item != null )
                {
                    var calEvent = new CalendarItem();
                    calEvent.LoadFromGoogleEvent( item );

                    return calEvent;
                }

                return null;
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to pull a list of events from google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
                return null;
            }
        }

        /// <summary>
        /// Pulls the list of calendars and allows a force refresh of the calendars.
        /// </summary>
        /// <param name="forceRefresh">If you want to force a refresh set to true</param>
        /// <returns>A list of the user's Google calendars</returns>
        public CalendarList PullCalendars( bool forceRefresh ) {
            if ( forceRefresh )
                m_lastUpdate = DateTime.MinValue;

            return PullCalendars();
        }

        /// <summary>
        /// Pulls the list of calendars.
        /// </summary>
        /// <returns>A list of the user's Google calendars</returns>
        public CalendarList PullCalendars() {
            try
            {

                // Force the list to be updated initially and then every 30 minutes.
                if ( m_lastUpdate == DateTime.MinValue || m_lastUpdate < DateTime.Now.Subtract( TimeSpan.FromMinutes( 30 ) ) )
                {
                    if ( m_service == null )
                        PerformAuthentication();

                    m_calendarList = m_service.CalendarList.List().Execute();
                    m_lastUpdate = DateTime.Now;
                }

                return m_calendarList;
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to pull a list of calendars from google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
                return null;
            } catch ( TokenResponseException ex )
            {
                Debug.WriteLine( ex );
                return null;
            }
        }

        /// <summary>
        /// Updates an appointment using the specified CalendarItem
        /// </summary>
        /// <param name="ev">The updated appointment</param>
        public void UpdateAppointment( CalendarItem ev ) {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                m_service.Events.Update( ev.GetGoogleCalendarEvent(), m_currentCalendar, ev.ID ).Execute();

            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to update an event on google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        /// <summary>
        /// Deletes all appointments that share the same ID provided by CalendarItem
        /// </summary>
        /// <param name="ev">The CalendarItem to be deleted</param>
        public void DeleteAppointment( CalendarItem ev ) {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                if ( ev.Recurrence != null )
                {
                    string pageToken = null;
                    do
                    {
                        var instancesRequest = m_service.Events.Instances( m_currentCalendar, ev.ID );
                        instancesRequest.PageToken = pageToken;

                        var events = instancesRequest.Execute();
                        pageToken = events.NextPageToken;

                        var list = events.Items.ToList();
                        foreach ( var @event in list )
                            m_service.Events.Delete( m_currentCalendar, @event.Id ).Execute();

                    } while ( pageToken != null );
                } else
                    m_service.Events.Delete( m_currentCalendar, ev.ID ).Execute();

                Archiver.Instance.Delete( ev.ID );
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to delete an event from google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        /// <summary>
        /// Deletes an appointment using just the ID
        /// </summary>
        /// <param name="id">The ID of the appointment</param>
        public void DeleteAppointment( string id ) {
            try
            {
                if ( m_service == null )
                    PerformAuthentication();

                m_service.Events.Delete( m_currentCalendar, id ).Execute();

                Archiver.Instance.Delete( id );
            } catch ( GoogleApiException ex )
            {
                Debug.Write( ex );
                MessageBox.Show( "There was an error when trying to delete an event from google.", "Error!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        /// <summary>
        /// Sets the current working folder (calendar) for all additions, deletions, updates, and queries.
        /// </summary>
        /// <param name="folder">The ID of the calendar</param>
        /// <param name="defaultFoler">Do you want to reset to the default folder? If you do you can leave folder blank.</param>
        public void SetGoogleWorkingFolder( string folder, bool defaultFoler = false ) {
            m_currentCalendar = defaultFoler ? "primary" : folder;
        }

        /// <summary>
        /// 
        /// </summary> 
        public void Dispose() {
            m_service?.Dispose();
        }

        /// <summary>
        /// Performs the initial authentication and the subsequient reauthentication
        /// </summary>
        /// <returns></returns>
        public bool PerformAuthentication() {
            if ( m_service != null )
                return true;

            try
            {
                m_currentCalendar = "primary";
                UserCredential credential;

                using ( var stream =
                    new FileStream( m_workingDirectory + "client_secret.json", FileMode.Open, FileAccess.Read ) )
                {
                    var credPath = Path.Combine( m_workingDirectory, ".credentials\\Outlook-Google-Sync.json" );
                    var cancel = new CancellationTokenSource( DEFAULT_CANCEL_TIME_OUT );

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load( stream ).Secrets,
                            m_scopes,
                            "user",
                            cancel.Token,
                            new FileDataStore( credPath, true ) )
                        .Result;

                    if ( credential.Token.IsExpired( SystemClock.Default ) )
                    {
                        GoogleWebAuthorizationBroker.ReauthorizeAsync( credential, cancel.Token ).RunSynchronously();
                    }

                    Debug.WriteLine( "Credential file saved to: " + credPath );
                }

                // Create Google Calendar API service.
                m_service = new CalendarService( new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                } );

                return true;
            } catch ( GoogleApiException ex )
            {
                Debug.WriteLine( ex );
                MessageBox.Show(
                    "There has been an error when trying to authenticate the user. Please review the error log for more information.",
                    "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return false;
            } catch ( AggregateException ex )
            {
                Debug.WriteLine( ex );

                if ( ex.InnerException != null && ex.InnerException.GetType() == typeof( TaskCanceledException ) )
                    MessageBox.Show( "The authorization timed out. Please try again.", "Oh no", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning );
                else
                    MessageBox.Show(
                        "Access to the calendar was denied by Google. You may need to reauthenticate this application.",
                        "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return false;
            }
        }
    }
}
