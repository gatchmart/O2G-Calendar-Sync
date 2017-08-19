using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace Outlook_Calendar_Sync {

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

        private readonly string TasksDateFilePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + "schedulerTasks.xml";
        private readonly string AutoSyncDateFilePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + "autoSync.xml";
        private List<SchedulerTask> m_tasks;
        private List<AutoSyncEvent> m_autoSyncEvents;
        private readonly Thread m_thread;

        public Scheduler()
        {
            if ( File.Exists( TasksDateFilePath ) )
                Load();
            else
                m_tasks = new List<SchedulerTask>();

            if ( !File.Exists( AutoSyncDateFilePath ) )
                m_autoSyncEvents = new List<AutoSyncEvent>();

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

        public void AddTask( SchedulerTask task )
        {
            lock ( m_tasks )
            {
                m_tasks.Add( task );
            }
        }

        public void RemoveTask( SchedulerTask task )
        {
            lock ( m_tasks )
            {
                m_tasks.Remove( task );

                if ( m_tasks.Count == 0 && m_thread.IsAlive )
                    m_thread.Abort();
            }
        }

        public void UpdateTask( SchedulerTask task, int index )
        {
            lock ( m_tasks )
            {
                if ( index >= 0 && index < m_tasks.Count && task != null )
                    m_tasks[index] = task;
            }
        }

        public SchedulerTask this[int key]
        {
            get { return m_tasks[key]; }
            set
            {
                if ( m_tasks != null ) m_tasks[key] = value;
            }
        }

        public void Item_Add( object item )
        {
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

        public void Item_Change( object item )
        {
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

        public void Item_Remove()
        {
            lock ( m_autoSyncEvents )
            {
                var ase = new AutoSyncEvent {Action = CalendarItemAction.GoogleDelete};
                if ( !m_autoSyncEvents.Contains( ase ) )
                    m_autoSyncEvents.Add( ase );
            }
        }

        private void Tick()
        {
            try
            {
                if ( m_tasks != null && m_tasks.Count > 0 )
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
                                        Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence, schedulerTask.SilentSync );
                                        schedulerTask.LastRunTime = DateTime.Now;
                                    }
                                } else if ( schedulerTask.Event == SchedulerEvent.Hourly )
                                {
                                    if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromHours( 1 ) ) )
                                    {
                                        Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence, schedulerTask.SilentSync );
                                        schedulerTask.LastRunTime = DateTime.Now;
                                    }
                                } else if ( schedulerTask.Event == SchedulerEvent.Weekly )
                                {
                                    if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromDays( 7 ) ) )
                                    {
                                        Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence, schedulerTask.SilentSync );
                                        schedulerTask.LastRunTime = DateTime.Now;
                                    }
                                } else if ( schedulerTask.Event == SchedulerEvent.CustomTime )
                                {
                                    if ( schedulerTask.LastRunTime <
                                         DateTime.Now.Subtract( TimeSpan.FromMinutes( schedulerTask.TimeSpan ) ) )
                                    {
                                        Syncer.Instance.SynchornizePairs( schedulerTask.Pair, schedulerTask.Precedence, schedulerTask.SilentSync );
                                        schedulerTask.LastRunTime = DateTime.Now;
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
                                    GoogleSync.Syncer.SetGoogleWorkingFolder( "", true );
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
                                    GoogleSync.Syncer.SetGoogleWorkingFolder( "", true );
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
                                        GoogleSync.Syncer.SetGoogleWorkingFolder( "", true );
                                    }
                                }
                            }

                            // Save the archiver and clear the auto sync events since they have been performed.
                            Archiver.Instance.Save();
                            m_autoSyncEvents.Clear();
                        }

                        // Sleep the thread for 30 seconds.
                        Thread.Sleep( 30000 );
                    }
                }
            } catch ( Exception ex )
            {
                Debug.WriteLine( ex );
            }
        }

        /// <summary>
        /// Saves the list of scheduled tasks.
        /// </summary>
        public void Save( bool runThread = true )
        {
            try
            {
                if ( m_tasks != null && m_tasks.Count > 0 )
                {
                    var serializer = new XmlSerializer( typeof( List<SchedulerTask> ) );
                    var writer = new StreamWriter( TasksDateFilePath );
                    serializer.Serialize( writer, m_tasks );
                    writer.Close();

                    if ( !m_thread.IsAlive && runThread )
                        m_thread.Start();

                } else if ( File.Exists( TasksDateFilePath ) )
                    File.Delete( TasksDateFilePath );

                if ( m_autoSyncEvents != null && m_autoSyncEvents.Count > 0 )
                {
                    var serializer = new XmlSerializer( typeof( List<AutoSyncEvent> ) );
                    var writer = new StreamWriter( AutoSyncDateFilePath );
                    serializer.Serialize( writer, m_autoSyncEvents );
                    writer.Close();

                } else if ( File.Exists( AutoSyncDateFilePath ) )
                    File.Delete( AutoSyncDateFilePath );

            } catch ( Exception ex )
            {
                Debug.WriteLine( ex );
            }
        }

        /// <summary>
        /// Activate the sync thread if not already activated
        /// </summary>
        public void ActivateThread()
        {
            if ( !m_thread.IsAlive )
                m_thread.Start();
        }

        /// <summary>
        /// Loads the list of scheduled tasks
        /// </summary>
        private void Load()
        {
            try
            {
                var serializer = new XmlSerializer( typeof( List<SchedulerTask> ) );
                var reader = new FileStream( TasksDateFilePath, FileMode.Open );

                if ( m_tasks != null )
                {
                    m_tasks.Clear();
                    m_tasks = null;
                }

                m_tasks = (List<SchedulerTask>) serializer.Deserialize( reader );

                reader.Close();

                if ( File.Exists( AutoSyncDateFilePath ) )
                {
                    serializer = new XmlSerializer( typeof( List<AutoSyncEvent> ) );
                    reader = new FileStream( AutoSyncDateFilePath, FileMode.Open );
                    if ( m_autoSyncEvents != null )
                    {
                        m_autoSyncEvents.Clear();
                        m_autoSyncEvents = null;
                    }

                    m_autoSyncEvents = (List<AutoSyncEvent>) serializer.Deserialize( reader );

                    reader.Close();
                }

            } catch ( Exception ex )
            {
                Debug.WriteLine( ex );
            }
        }

        public IEnumerator<SchedulerTask> GetEnumerator()
        {
            return m_tasks.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    [Serializable]
    public class SchedulerTask
    {
        /// <summary>
        /// The calendar pair to sync
        /// </summary>
        public SyncPair Pair;

        /// <summary>
        /// When should it sync
        /// </summary>
        public SchedulerEvent Event;

        /// <summary>
        /// The delay between syncs in minutes. This is only used with SchedulerEvent.CustomTime
        /// </summary>
        public int TimeSpan;

        /// <summary>
        /// This tells us the last time the pair was synced
        /// </summary>
        public DateTime LastRunTime;

        /// <summary>
        /// Allows you to set which calendar takes precedence over the other.
        /// This comes in handy if you want a silent sync.
        /// 0 = Ignore differences
        /// 1 = Outlook has precedence
        /// 2 = Google has precedence
        /// </summary>
        public int Precedence;

        /// <summary>
        /// Allows you to set if the sync will prompt the user when changes have occurred.
        /// Use this with the precedence property
        /// </summary>
        public bool SilentSync;

        public SchedulerTask()
        {
            Pair = null;
            Event = SchedulerEvent.Manually;
            TimeSpan = 0;
            LastRunTime = DateTime.MinValue;
            Precedence = 0;
            SilentSync = false;
        }

        public override string ToString()
        {
            return string.Format( "{0} <=> {1}", Pair.GoogleName, Pair.OutlookName );
        }
    }

    [Serializable]
    public class AutoSyncEvent
    {
        public SyncPair Pair;
        public string EntryId;
        public CalendarItemAction Action;
    }
}
