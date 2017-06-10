using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace Outlook_Calendar_Sync {
    class GoogleSync {

        public static GoogleSync Syncer => _instance ?? ( _instance = new GoogleSync() );

        private static GoogleSync _instance;
        public static string CurrentCalendar { get; set; }

        private readonly string[] m_scopes = { CalendarService.Scope.Calendar };
        private const string ApplicationName = "Google Calendar API .NET Quickstart";

        private readonly CalendarService m_service;

        public GoogleSync() {
            CurrentCalendar = "primary";
            UserCredential credential;
            var dir = Directory.GetCurrentDirectory();

            using ( var stream =
                new FileStream( dir + "\\client_secret.json", FileMode.Open, FileAccess.Read ) ) {
                string credPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal );
                credPath = Path.Combine( credPath, ".credentials/calendar-dotnet-quickstart.json" );

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load( stream ).Secrets,
                    m_scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore( credPath, true ) ).Result;
                Debug.WriteLine( "Credential file saved to: " + credPath );
            }

            // Create Google Calendar API service.
            m_service = new CalendarService( new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            } );
        }

        public void AddAppointment( CalendarItem item ) {
            var e = item.GetGoogleCalendarEvent();
            m_service.Events.Insert( e, CurrentCalendar ).Execute();
            Archiver.Instance.Add( e.Id );
        }

        public List<CalendarItem> PullListOfAppointments() {
            List<CalendarItem> items = new List<CalendarItem>();

            // Iterate over the events in the specified calendar
            string pageToken = null;
            do {
                EventsResource.ListRequest list = m_service.Events.List( CurrentCalendar );
                list.PageToken = pageToken;
                Events events = list.Execute();
                List<Event> i = events.Items.ToList();

                foreach ( var @event in i ) {
                    var cal = new CalendarItem();
                    cal.LoadFromGoogleEvent( @event );
                    if ( !items.Exists( x => x.ID.Equals( cal.ID ) ) )
                        items.Add( cal );
                }

                pageToken = events.NextPageToken;
            } while ( pageToken != null );

            return items;
        }

        public List<CalendarItem> PullListOfAppointmentsByDate( DateTime startDate, DateTime endDate ) {
            List<CalendarItem> items = new List<CalendarItem>();

            // Iterate over the events in the specified calendar
            string pageToken = null;
            do {
                EventsResource.ListRequest list = m_service.Events.List( CurrentCalendar );
                list.TimeMin = startDate;
                list.TimeMax = endDate;
                list.PageToken = pageToken;
                Events events = list.Execute();
                List<Event> i = events.Items.ToList();

                foreach ( var @event in i ) {
                    var cal = new CalendarItem();
                    cal.LoadFromGoogleEvent( @event );
                    if ( !items.Exists( x => x.ID.Equals( cal.ID ) ) )
                        items.Add( cal );
                }

                pageToken = events.NextPageToken;
            } while ( pageToken != null );

            return items;
        }

        public CalendarList PullCalendars() {
            CalendarList calendarList = m_service.CalendarList.List().Execute();

            return calendarList; 
        }

        public void UpdateAppointment( CalendarItem ev ) {
            m_service.Events.Update( ev.GetGoogleCalendarEvent(), CurrentCalendar, ev.ID  ).Execute();
        }

        public void DeleteAppointment( CalendarItem ev ) {
            // TODO: Deal with reoccuring events
            m_service.Events.Delete( CurrentCalendar, ev.ID ).Execute();
            Archiver.Instance.Delete( ev.ID );
        }

        public CalendarItem FindItem( string googleId, DateTime startDate, DateTime endDate ) {
            var items = PullListOfAppointmentsByDate( startDate, endDate );
            return items.FirstOrDefault( calendarItem => calendarItem.ID.Equals( googleId ) );
        }

    }
}
