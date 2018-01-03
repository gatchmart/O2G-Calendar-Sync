using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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
using Outlook_Calendar_Sync.Enums;
using Outlook_Calendar_Sync.Properties;
using Outlook_Calendar_Sync.Scheduler;

namespace Outlook_Calendar_Sync {

    /// <summary>
    /// The Google Syncer is the connection between the application and the Google Calendar API. This is a Singleton class so you need to use GoogleSync.Syncer to access it.
    /// </summary>
    public sealed class GoogleSync : IDisposable {

        /// <summary>
        /// The Singleton instance of GoogleSync
        /// </summary>
        public static GoogleSync Syncer => _instance ?? ( _instance = new GoogleSync() );
        private static GoogleSync _instance;

        /// <summary>
        /// The retry action is used when trying an action
        /// </summary>
        public RetryTask Retry;

        private readonly string[] m_scopes = { CalendarService.Scope.Calendar };
        private readonly string m_workingDirectory =
            Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
            "\\OutlookGoogleSync\\.credentials\\Outlook-Google-Sync.json";

        private const string APPLICATION_NAME = "Outlook Google Calendar Sync";
        private const int DEFAULT_CANCEL_TIME_OUT = 60000;

        private string m_currentCalendar;
        private string m_previousCalendar;
        private TokenResponse m_credentialToken;
        private CalendarService m_service;
        private CalendarList m_calendarList;
        private DateTime m_lastUpdate;
        private readonly FileDataStore m_fileDataStore;

        /// <summary>
        /// The default constructor for GoogleSync. This should never be called directly. You should always use GoogleSync.Syncer
        /// </summary>
        public GoogleSync() {
            m_currentCalendar = "primary";
            m_lastUpdate = DateTime.MinValue;
            m_calendarList = null;
            m_fileDataStore = new FileDataStore( m_workingDirectory, true );
        }

        /// <summary>
        /// Add an appointment to the selected Google calendar.
        /// </summary>
        /// <param name="item">The CalendarItem to add</param>
        /// <returns>nothing</returns>
        public void AddAppointment( CalendarItem item ) {
            try
            {
                PerformAuthentication();

                Event newEvent;
                if ( string.IsNullOrEmpty( item.CalendarItemIdentifier.GoogleICalUId ) )
                    newEvent = m_service.Events.Insert( item.GetGoogleCalendarEvent(), m_currentCalendar ).Execute();
                else
                    newEvent = m_service.Events.Import( item.GetGoogleCalendarEvent(), m_currentCalendar ).Execute();

                Log.Write( $"Added {item.Subject} Appointment to Google" );

                var oldId = item.CalendarItemIdentifier;
                item.CalendarItemIdentifier = new Identifier( newEvent.Id, newEvent.ICalUID, oldId.OutlookEntryId, oldId.OutlookGlobalId );

                Archiver.Instance.UpdateIdentifier( oldId, item.CalendarItemIdentifier );

                Retry?.Successful();
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to add an event to google. Review the log file to get more information.", item, RetryAction.Add );
            }
        }

        /// <summary>
        /// Pull a complete list of appointments for the set calendar and retreives the sync token.
        /// </summary>
        /// <returns>List of CalendarItems</returns>
        public List<CalendarItem> PullListOfAppointmentsBySyncToken() {
            try
            {
                PerformAuthentication();

                Log.Write( $"Pulling a list of Google Appointments from {m_currentCalendar} with sync token." );

                EventsResource.ListRequest list = m_service.Events.List( m_currentCalendar );

                string syncToken = m_fileDataStore.GetAsync<string>( m_currentCalendar + " - Sync Token" ).Result;
                if ( !string.IsNullOrEmpty( syncToken ) )
                    list.SyncToken = syncToken;

                List<CalendarItem> items = PullListOfAppointments( list, ref syncToken );

                m_fileDataStore.StoreAsync( m_currentCalendar + " - Sync Token", syncToken );

                return items;
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to pull a list of events from google." );
                return null;
            }
        }

        /// <summary>
        /// Pull a complete list of appointments for the set calendar
        /// </summary>
        /// <returns>List of CalendarItems</returns>
        public List<CalendarItem> PullListOfAppointments()
        {
            try
            {
                PerformAuthentication();

                EventsResource.ListRequest list = m_service.Events.List( m_currentCalendar );
                string token = "";

                Log.Write( $"Pulling a list of Google Appointments from {m_currentCalendar}." );

                var items = PullListOfAppointments( list, ref token );
                return items;
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to pull a list of events from the Google calendar, " + m_currentCalendar );
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
                PerformAuthentication();

                List<CalendarItem> items = new List<CalendarItem>();

                Log.Write( $"Pulling a list of Google Appointments by date from {m_currentCalendar}." );

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
                        if ( @event.Status.Equals( "cancelled" ) )
                            continue;

                        var cal = new CalendarItem();
                        cal.LoadFromGoogleEvent( @event );
                        if ( !items.Exists( x => x.CalendarItemIdentifier.GoogleId.Equals( cal.CalendarItemIdentifier.GoogleId ) ) )
                            items.Add( cal );
                    }

                    pageToken = events.NextPageToken;
                } while ( pageToken != null );

                return items;
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to pull a list of events from google." );
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
                PerformAuthentication();

                Log.Write( $"Looking for a Google Appointment from {m_currentCalendar} with ID {id}." );

                var item = m_service.Events.Get( m_currentCalendar, id ).Execute();
                if ( item != null )
                {
                    var calEvent = new CalendarItem();
                    calEvent.LoadFromGoogleEvent( item );

                    Log.Write( $"Found Google appointment with ID {id}." );

                    return calEvent;
                }

                Log.Write( $"Could not find Google appointment with ID {id}." );

                return null;
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to pull a list of events from google." );
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

            Log.Write( $"Performed forced refresh of Google calendars." );

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
                    PerformAuthentication();

                    m_calendarList = m_service.CalendarList.List().Execute();
                    m_lastUpdate = DateTime.Now;
                }

                Log.Write( "Pulled the list of Google calendars." );

                return m_calendarList;
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to pull a list of calendars from google." );
                return null;
            } catch ( TokenResponseException ex )
            {
                Log.Write( ex );
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
                PerformAuthentication();

                Event gEvent = ev.GetGoogleCalendarEvent();
                if ( gEvent.Id.StartsWith( "_" ) )
                    gEvent.Id = gEvent.Id.Remove( 0, 1 );

                var newEvent = m_service.Events.Update( gEvent, m_currentCalendar, gEvent.Id ).Execute();

                var oldId = ev.CalendarItemIdentifier;
                ev.CalendarItemIdentifier = new Identifier( newEvent.Id, newEvent.ICalUID, oldId.OutlookEntryId, oldId.OutlookGlobalId );

                Archiver.Instance.UpdateIdentifier( oldId, ev.CalendarItemIdentifier );

                Log.Write( $"Updated Google appointment, {ev.Subject}." );
                Retry?.Successful();

            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to update an event on google.", ev, RetryAction.Update );
            }
        }

        /// <summary>
        /// Deletes all appointments that share the same ID provided by CalendarItem
        /// </summary>
        /// <param name="ev">The CalendarItem to be deleted</param>
        public void DeleteAppointment( CalendarItem ev ) {
            try
            {
                PerformAuthentication();

                if ( ev.Recurrence != null )
                {
                    string pageToken = null;
                    do
                    {
                        var instancesRequest = m_service.Events.Instances( m_currentCalendar, ev.CalendarItemIdentifier.GoogleId );
                        instancesRequest.PageToken = pageToken;

                        var events = instancesRequest.Execute();
                        pageToken = events.NextPageToken;

                        var list = events.Items.ToList();
                        foreach ( var @event in list )
                            m_service.Events.Delete( m_currentCalendar, @event.Id ).Execute();

                        Log.Write( $"Deleted Google recurring appointment {ev.Subject}." );

                    } while ( pageToken != null );
                } else
                {
                    m_service.Events.Delete( m_currentCalendar, ev.CalendarItemIdentifier.GoogleId ).Execute();
                    Log.Write( $"Deleted Google appointment {ev.Subject}." );
                }

                Archiver.Instance.Delete( ev.CalendarItemIdentifier );
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to delete an event from google.", ev, RetryAction.Delete );
            }
        }

        /// <summary>
        /// Deletes an appointment using just the ID
        /// </summary>
        /// <param name="id">The ID of the appointment</param>
        public void DeleteAppointment( Identifier id ) {
            try
            {
                PerformAuthentication();

                m_service.Events.Delete( m_currentCalendar, id.GoogleId ).Execute();
                Log.Write( $"Deleted Google appointment with ID, {id}." );

                Archiver.Instance.Delete( id );
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                HandleException( ex, "There was an error when trying to delete an event from google.", new CalendarItem{ CalendarItemIdentifier = id }, RetryAction.DeleteById );
            }
        }

        /// <summary>
        /// Sets the current working folder (calendar) for all additions, deletions, updates, and queries.
        /// </summary>
        /// <param name="folder">The ID of the calendar</param>
        public void SetGoogleWorkingFolder( string folder )
        {
            Log.Write( $"Set Google working folder to, { folder }" );
            m_previousCalendar = m_currentCalendar;
            m_currentCalendar = folder;
        }

        /// <summary>
        /// Resets the working folder back to it's previous setting
        /// </summary>
        /// <param name="defaultFoler">Do you want to reset to the default folder? If you do you can leave folder blank.</param>
        public void ResetGoogleWorkingFolder( bool defaultFoler = false )
        {
            Log.Write( $"Set Google working folder to, { ( defaultFoler ? "primary" : m_previousCalendar ) }" );
            m_currentCalendar = defaultFoler ? "primary" : m_previousCalendar;
            m_previousCalendar = null;
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
            if ( m_service != null && m_credentialToken != null && !m_credentialToken.IsExpired( SystemClock.Default ) )
                return true;

            try
            {
                UserCredential credential;
                byte[] secrets = Resources.client_secret;

                Log.Write( "Performing Google authentication" );

                using ( var stream = new MemoryStream( secrets ) )
                {
                    var cancel = new CancellationTokenSource( DEFAULT_CANCEL_TIME_OUT );

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load( stream ).Secrets,
                            m_scopes,
                            "user",
                            cancel.Token,
                            m_fileDataStore )
                        .Result;

                    m_credentialToken = credential.Token;

                    Log.Write( "Credential file saved to: " + m_workingDirectory );
                }

                // Create Google Calendar API service.
                m_service = new CalendarService( new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = APPLICATION_NAME
                } );

                return true;
            } catch ( GoogleApiException ex )
            {
                Log.Write( ex );
                MessageBox.Show(
                    "There has been an error when trying to authenticate the user. Please review the error log for more information.",
                    "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return false;
            } catch ( AggregateException ex )
            {
                Log.Write( ex );

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

        private List<CalendarItem> PullListOfAppointments( EventsResource.ListRequest list, ref string syncToken )
        {
            // Iterate over the events in the specified calendar
            List<CalendarItem> items = new List<CalendarItem>();
            string pageToken = null;
            Events events = null;
            do
            {
                list.PageToken = pageToken;
                events = list.Execute();
                List<Event> i = events.Items.ToList();

                foreach ( var @event in i )
                {
                    if ( @event.Status.Equals( "cancelled" ) )
                        continue;

                    var cal = new CalendarItem();
                    cal.LoadFromGoogleEvent( @event );
                    if ( !items.Exists( x => x.CalendarItemIdentifier.GoogleId.Equals( cal.CalendarItemIdentifier.GoogleId ) ) )
                        items.Add( cal );
                }

                pageToken = events.NextPageToken;
            } while ( pageToken != null );
            syncToken = events.NextSyncToken;

                return items;
        }

        private void HandleException( GoogleApiException ex, string errorMsg, CalendarItem item = null, RetryAction action = 0 )
        {
            if ( ex.HttpStatusCode == HttpStatusCode.BadRequest || ex.HttpStatusCode == HttpStatusCode.InternalServerError )
            {
                MessageBox.Show( errorMsg, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error );
            } else if ( ex.HttpStatusCode == HttpStatusCode.Unauthorized )
            {
                // Reauthenticate
                m_service.Dispose();
                m_service = null;
                PerformAuthentication();

                CreateRetry( item, action );
            } else if ( ex.HttpStatusCode == HttpStatusCode.Conflict )
            {
                if ( item != null )
                    item.CalendarItemIdentifier.GoogleId = GuidCreator.Create();

                CreateRetry( item, action );
            } else if ( ex.HttpStatusCode == HttpStatusCode.Forbidden )
            {
                CreateRetry( item, action );
            }
        }

        private void CreateRetry( CalendarItem item, RetryAction action )
        {
            if ( Retry == null )
            {
                var retry = new RetryTask( item, m_currentCalendar, action );
                Scheduler.Scheduler.Instance.AddRetry( retry );
            } else
            {
                Retry.RetryFailed();
                Retry = null;
            }
        }
    }

    
}
