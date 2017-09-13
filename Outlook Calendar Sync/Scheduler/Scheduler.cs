using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Xml.Serialization;
using Outlook_Calendar_Sync.Enums;
using Outlook = Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace Outlook_Calendar_Sync.Scheduler
{

    public enum SchedulerEvent
    {
        Automatically = 0,
        OnStartup = 1,
        Hourly = 2,
        Daily = 3,
        Weekly = 4,
        CustomTime = 5,
        Manually = 6
    }

    public class Scheduler : IEnumerable<SchedulerTask>
    {
        public static Scheduler Instance => _instance ?? ( _instance = new Scheduler() );
        private static Scheduler _instance;

        public int Count => m_tasks.Count;

        private readonly string TasksDataFilePath =
            Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" +
            "schedulerTasks.xml";

        private readonly string AutoSyncDataFilePath =
            Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" +
            "autoSync.xml";

        private readonly string RetryDataFilePath =
            Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" +
            "retryData.xml";

        private List<SchedulerTask> m_tasks;
        private List<AutoSyncEvent> m_autoSyncEvents;
        private List<RetryTask> m_retryList;
        private Queue<int> m_retryDeleteQueue;
        private readonly Thread m_thread;

        public Scheduler()
        {
            if ( !File.Exists( TasksDataFilePath ) )
                m_tasks = new List<SchedulerTask>();

            if ( !File.Exists( AutoSyncDataFilePath ) )
                m_autoSyncEvents = new List<AutoSyncEvent>();

            if ( !File.Exists( RetryDataFilePath ) )
                m_retryList = new List<RetryTask>();

            m_retryDeleteQueue = new Queue<int>();

            Load();

            m_thread = new Thread( Tick );
            m_thread.IsBackground = true;

            if ( m_tasks.Count > 0 )
                m_thread.Start();

            foreach ( var task in m_tasks )
            {
                if ( task.Event == SchedulerEvent.OnStartup )
                {
                    Syncer.Instance.SynchornizePairs( task.Pair );
                    task.LastRunTime = DateTime.Now;
                }
            }
        }

        #region List Modifiers

        public void AddTask( SchedulerTask task ) {
            lock ( m_tasks )
            {
                m_tasks.Add( task );
            }
        }

        public void RemoveTask( SchedulerTask task ) {
            lock ( m_tasks )
            {
                m_tasks.Remove( task );
            }
        }

        public void UpdateTask( SchedulerTask task, int index ) {
            lock ( m_tasks )
            {
                if ( index >= 0 && index < m_tasks.Count && task != null )
                    m_tasks[index] = task;
            }
        }

        public void AddRetry( RetryTask task ) {
            lock ( m_retryList )
            {
                m_retryList.Add( task );
            }
        }

        public void RemoveRetry( RetryTask task ) {
            lock ( m_retryList )
            {
                m_retryList.Remove( task );
            }
        }

        public SchedulerTask this[int key] {
            get { return m_tasks[key]; }
            set
            {
                if ( m_tasks != null ) m_tasks[key] = value;
            }
        }

        public IEnumerator<SchedulerTask> GetEnumerator() {
            return m_tasks.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return GetEnumerator();
        }

        #endregion

        #region EventHandlers

        public void Item_Add( object item ) {
            lock ( m_autoSyncEvents )
            {
                Outlook.AppointmentItem aitem = item as Outlook.AppointmentItem;
                if ( aitem != null )
                {
                    Outlook.MAPIFolder calender = aitem.Parent as Outlook.MAPIFolder;
                    if ( calender != null )
                    {
                        var tasks = m_tasks.FindAll( x => x.Event == SchedulerEvent.Automatically &&
                                                          x.Pair.OutlookId.Equals( calender.EntryID ) );

                        foreach ( var schedulerTask in tasks )
                        {
                            m_autoSyncEvents.Add( new AutoSyncEvent
                            {
                                Action = CalendarItemAction.GoogleAdd,
                                EntryId = aitem.EntryID,
                                Pair = schedulerTask.Pair
                            } );
                        }
                    }
                }
            }
        }

        public void Item_Change( object item ) {
            lock ( m_autoSyncEvents )
            {
                Outlook.AppointmentItem aitem = item as Outlook.AppointmentItem;
                if ( aitem != null )
                {
                    Outlook.MAPIFolder calender = aitem.Parent as Outlook.MAPIFolder;
                    if ( calender != null )
                    {
                        var tasks = m_tasks.FindAll( x => x.Event == SchedulerEvent.Automatically &&
                                                          x.Pair.OutlookId.Equals( calender.EntryID ) );

                        foreach ( var schedulerTask in tasks )
                        {
                            m_autoSyncEvents.Add( new AutoSyncEvent
                            {
                                Action = CalendarItemAction.GoogleUpdate,
                                EntryId = aitem.EntryID,
                                Pair = schedulerTask.Pair
                            } );
                        }
                    }
                }
            }
        }

        public void Item_Remove() {
            lock ( m_autoSyncEvents )
            {
                var ase = new AutoSyncEvent { Action = CalendarItemAction.GoogleDelete };
                if ( !m_autoSyncEvents.Contains( ase ) )
                    m_autoSyncEvents.Add( ase );
            }
        } 

        #endregion

        private void Tick()
        {
            try
            {
                // Setup an initial delay before starting to sync
                Thread.Sleep( 5000 );
                while ( true )
                {
                    // Lock m_tasks so the main thread doesn't mess with it while we are looping through it.
                    lock ( m_tasks )
                    {


                        // Loop through all the tasks and perform syncs as apporiate.
                        foreach ( var schedulerTask in m_tasks )
                        {
                            if ( schedulerTask.Event == SchedulerEvent.Daily )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromDays( 1 ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence,
                                        schedulerTask.SilentSync );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.Hourly )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromHours( 1 ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence,
                                        schedulerTask.SilentSync );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.Weekly )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromDays( 7 ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence,
                                        schedulerTask.SilentSync );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.CustomTime )
                            {
                                if ( schedulerTask.LastRunTime <
                                     DateTime.Now.Subtract( TimeSpan.FromMinutes( schedulerTask.TimeSpan ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence,
                                        schedulerTask.SilentSync );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.Automatically )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromMinutes( 1 ) ) )
                                {
                                    Syncer.Instance.IsUsingSyncToken = true;
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence,
                                        schedulerTask.SilentSync );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                    Syncer.Instance.IsUsingSyncToken = false;
                                }
                            }

                        }
                    }

                    // Lock m_autoSyncEvents to ensure no other threads modify it as we iterate through it
                    lock ( m_autoSyncEvents )
                    {
                        foreach ( var autoSyncEvent in m_autoSyncEvents )
                        {
                            if ( autoSyncEvent.Action == CalendarItemAction.GoogleAdd )
                            {
                                // Setup the proper working folders and the CurrentPair
                                OutlookSync.Syncer.SetOutlookWorkingFolder( autoSyncEvent.Pair.OutlookId );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( autoSyncEvent.Pair.GoogleId );
                                Archiver.Instance.CurrentPair = autoSyncEvent.Pair;

                                // Find the Outlook appointment using its ID.
                                var calendarItem = OutlookSync.Syncer.FindEventByEntryId( autoSyncEvent.EntryId );
                                if ( calendarItem != null )
                                    GoogleSync.Syncer.AddAppointment( calendarItem );

                                // Reset the default working folders.
                                OutlookSync.Syncer.SetOutlookWorkingFolder( "", true );
                                GoogleSync.Syncer.ResetGoogleWorkingFolder();
                            } else if ( autoSyncEvent.Action == CalendarItemAction.GoogleUpdate )
                            {
                                // Setup the proper working folders and the CurrentPair
                                OutlookSync.Syncer.SetOutlookWorkingFolder( autoSyncEvent.Pair.OutlookId );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( autoSyncEvent.Pair.GoogleId );
                                Archiver.Instance.CurrentPair = autoSyncEvent.Pair;

                                // Find the Outlook appointment using its ID.
                                var calendarItem = OutlookSync.Syncer.FindEventByEntryId( autoSyncEvent.EntryId );
                                if ( calendarItem != null )
                                    GoogleSync.Syncer.UpdateAppointment( calendarItem );

                                // Reset the default working folders.
                                OutlookSync.Syncer.SetOutlookWorkingFolder( "", true );
                                GoogleSync.Syncer.ResetGoogleWorkingFolder();
                            } else if ( autoSyncEvent.Action == CalendarItemAction.GoogleDelete )
                            {
                                // Find the deleted appointments in Outlook
                                foreach ( var task in m_tasks )
                                {
                                    var events = Syncer.Instance.FindDeletedEvents( task.Pair );

                                    OutlookSync.Syncer.SetOutlookWorkingFolder( task.Pair.OutlookId );
                                    GoogleSync.Syncer.SetGoogleWorkingFolder( task.Pair.GoogleId );
                                    Archiver.Instance.CurrentPair = task.Pair;

                                    foreach ( var evnt in events )
                                        GoogleSync.Syncer.DeleteAppointment( evnt );

                                    // Reset the default working folders.
                                    OutlookSync.Syncer.SetOutlookWorkingFolder( "", true );
                                    GoogleSync.Syncer.ResetGoogleWorkingFolder();
                                }
                            }
                        }

                        // Save the archiver and clear the auto sync events since they have been performed.
                        Archiver.Instance.Save();
                        m_autoSyncEvents.Clear();
                    }

                    // Lock the m_retryList to prevent data corruption
                    lock ( m_retryList )
                    {
                        for (int i = 0; i < m_retryList.Count; i++)
                        {
                            var retry = m_retryList[i];
                            if ( retry.Eligible() && retry.LastRun <
                                 DateTime.Now.Subtract( TimeSpan.FromMinutes( retry.Delay ) ) )
                            {
                                switch ( retry.Action )
                                {
                                    case RetryAction.Add:
                                        GoogleSync.Syncer.Retry = retry;
                                        GoogleSync.Syncer.SetGoogleWorkingFolder( retry.Calendar );

                                        GoogleSync.Syncer.AddAppointment( retry.CalendarItem );

                                        GoogleSync.Syncer.Retry = null;
                                        GoogleSync.Syncer.ResetGoogleWorkingFolder();
                                        break;
                                    case RetryAction.Update:
                                        GoogleSync.Syncer.Retry = retry;
                                        GoogleSync.Syncer.SetGoogleWorkingFolder( retry.Calendar );

                                        GoogleSync.Syncer.UpdateAppointment( retry.CalendarItem );

                                        GoogleSync.Syncer.Retry = null;
                                        GoogleSync.Syncer.ResetGoogleWorkingFolder();
                                        break;
                                    case RetryAction.Delete:
                                        GoogleSync.Syncer.Retry = retry;
                                        GoogleSync.Syncer.SetGoogleWorkingFolder( retry.Calendar );

                                        GoogleSync.Syncer.DeleteAppointment( retry.CalendarItem );

                                        GoogleSync.Syncer.Retry = null;
                                        GoogleSync.Syncer.ResetGoogleWorkingFolder();
                                        break;
                                    case RetryAction.DeleteById:
                                        GoogleSync.Syncer.Retry = retry;
                                        GoogleSync.Syncer.SetGoogleWorkingFolder( retry.Calendar );

                                        GoogleSync.Syncer.DeleteAppointment( retry.CalendarItem.ID );

                                        GoogleSync.Syncer.Retry = null;
                                        GoogleSync.Syncer.ResetGoogleWorkingFolder();
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException();
                                }
                            } else if ( !retry.Eligible() )
                                RemoveRetry( retry );
                        }
                    }

                    // Sleep the thread for 30 seconds.
                    Thread.Sleep( 30000 );
                }

            } catch ( Exception ex )
            {
                Log.Write( ex );
            }
        }

        /// <summary>
        /// Activate the sync thread if not already activated
        /// </summary>
        public void ActivateThread()
        {
            if ( !m_thread.IsAlive )
            {
                m_thread.Start();
                Log.Write( "Activated Scheduler Thread." );
            }
        }

        public void AboutThread()
        {
            if ( m_thread.IsAlive )
            {
                m_thread.Abort();
                Log.Write( "Aborted Scheduler Thread" );
            }
        }

        /// <summary>
        /// Saves the list of scheduled tasks.
        /// </summary>
        public void Save( bool runThread = true ) {
            try
            {
                if ( m_tasks != null && m_tasks.Count > 0 )
                {
                    var serializer = new XmlSerializer( typeof( List<SchedulerTask> ) );
                    var writer = new StreamWriter( TasksDataFilePath );
                    serializer.Serialize( writer, m_tasks );
                    writer.Close();

                    if ( !m_thread.IsAlive && runThread )
                        m_thread.Start();

                } else if ( File.Exists( TasksDataFilePath ) )
                    File.Delete( TasksDataFilePath );

                if ( m_autoSyncEvents != null && m_autoSyncEvents.Count > 0 )
                {
                    var serializer = new XmlSerializer( typeof( List<AutoSyncEvent> ) );
                    var writer = new StreamWriter( AutoSyncDataFilePath );
                    serializer.Serialize( writer, m_autoSyncEvents );
                    writer.Close();

                } else if ( File.Exists( AutoSyncDataFilePath ) )
                    File.Delete( AutoSyncDataFilePath );

                if ( m_retryList != null && m_retryList.Count > 0 )
                {
                    var serializer = new XmlSerializer( typeof( List<RetryTask> ) );
                    var writer = new StreamWriter( RetryDataFilePath );
                    serializer.Serialize( writer, m_retryList );
                    writer.Close();

                } else if ( File.Exists( RetryDataFilePath ) )
                    File.Delete( RetryDataFilePath );

            } catch ( Exception ex )
            {
                Log.Write( ex );
            }
        }

        /// <summary>
        /// Loads the list of scheduled tasks
        /// </summary>
        private void Load()
        {
            try
            {
                XmlSerializer serializer;
                FileStream reader;

                Log.Write( "Loading Scheduler Data." );

                if ( File.Exists( TasksDataFilePath ) )
                {
                    Log.Write( "Found Scheduler Tasks, loading them in." );
                    serializer = new XmlSerializer( typeof( List<SchedulerTask> ) );
                    reader = new FileStream( TasksDataFilePath, FileMode.Open );

                    if ( m_tasks != null )
                    {
                        m_tasks.Clear();
                        m_tasks = null;
                    }

                    m_tasks = (List<SchedulerTask>)serializer.Deserialize( reader );

                    reader.Close();
                    Log.Write( "Completed loading in scheduler tasks." );
                }

                if ( File.Exists( AutoSyncDataFilePath ) )
                {
                    Log.Write( "Found Scheduler Tasks, loading them in." );
                    serializer = new XmlSerializer( typeof( List<AutoSyncEvent> ) );
                    reader = new FileStream( AutoSyncDataFilePath, FileMode.Open );
                    if ( m_autoSyncEvents != null )
                    {
                        m_autoSyncEvents.Clear();
                        m_autoSyncEvents = null;
                    }

                    m_autoSyncEvents = (List<AutoSyncEvent>) serializer.Deserialize( reader );

                    reader.Close();
                    Log.Write( "Completed loading in scheduler tasks." );
                }

                if ( File.Exists( RetryDataFilePath ) )
                {
                    Log.Write( "Found Scheduler Tasks, loading them in." );
                    serializer = new XmlSerializer( typeof( List<RetryTask> ) );
                    reader = new FileStream( RetryDataFilePath, FileMode.Open );
                    if ( m_retryList != null )
                    {
                        m_retryList.Clear();
                        m_retryList = null;
                    }

                    m_retryList = (List<RetryTask>) serializer.Deserialize( reader );

                    reader.Close();
                    Log.Write( "Completed loading in scheduler tasks." );
                }

            } catch ( Exception ex )
            {
                Log.Write( ex );
            }
        }

    }
}

