using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Xml.Serialization;
using Google.Apis.Util;
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

        private readonly string DateFilePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + "schedulerTasks.xml";
        private List<SchedulerTask> m_tasks;
        private List<AutoSyncEvent> m_autoSyncEvents;
        private readonly Thread m_thread;

        public Scheduler()
        {
            if ( File.Exists( DateFilePath ) )
                Load();
            else
                m_tasks = new List<SchedulerTask>();

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
            m_tasks.Add( task );
        }

        public void RemoveTask( SchedulerTask task )
        {
            m_tasks.Remove( task );

            if ( m_tasks.Count == 0 && m_thread.IsAlive )
                m_thread.Abort();

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
            Outlook.AppointmentItem aitem = item as Outlook.AppointmentItem;
            if ( aitem != null )
            {
                Outlook.MAPIFolder calender = aitem.Parent as Outlook.MAPIFolder;
                if ( calender != null )
                {
                    var tasks = m_tasks.FindAll( x => x.Event == SchedulerEvent.Automatically && x.Pair.OutlookId.Equals( calender.EntryID) );

                    foreach ( var schedulerTask in tasks )
                    {
                        m_autoSyncEvents.Add( new AutoSyncEvent { Action = CalendarItemAction.GoogleAdd, EntryId = aitem.EntryID, Pair = schedulerTask.Pair } );
                    }
                }
            }
        }

        public void Item_Change( object item )
        {
            Outlook.AppointmentItem aitem = item as Outlook.AppointmentItem;
            if ( aitem != null )
            {
                Outlook.MAPIFolder calender = aitem.Parent as Outlook.MAPIFolder;
                if ( calender != null )
                {
                    var tasks = m_tasks.FindAll( x => x.Event == SchedulerEvent.Automatically && x.Pair.OutlookId.Equals( calender.EntryID ) );

                    foreach ( var schedulerTask in tasks )
                    {
                        m_autoSyncEvents.Add( new AutoSyncEvent { Action = CalendarItemAction.GoogleUpdate, EntryId = aitem.EntryID, Pair = schedulerTask.Pair } );
                    }
                }
            }
        }

        public void Item_Remove()
        {
            
        }

        private void Tick()
        {
            if ( m_tasks != null && m_tasks.Count > 0 )
            {
                Thread.Sleep( 5000 );
                while ( true )
                {
                    lock ( m_tasks )
                    {
                        foreach ( var schedulerTask in m_tasks )
                        {
                            if ( schedulerTask.Event == SchedulerEvent.Daily )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromDays( 1 ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.Hourly )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromHours( 1 ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.Weekly )
                            {
                                if ( schedulerTask.LastRunTime < DateTime.Now.Subtract( TimeSpan.FromDays( 7 ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            } else if ( schedulerTask.Event == SchedulerEvent.CustomTime )
                            {
                                if ( schedulerTask.LastRunTime <
                                     DateTime.Now.Subtract( TimeSpan.FromMinutes( schedulerTask.TimeSpan ) ) )
                                {
                                    Syncer.Instance.SynchornizePairs( schedulerTask.Pair );
                                    schedulerTask.LastRunTime = DateTime.Now;
                                }
                            }

                        }
                    }

                    lock ( m_autoSyncEvents )
                    {
                        foreach ( var autoSyncEvent in m_autoSyncEvents )
                        {
                            if ( autoSyncEvent.Action == CalendarItemAction.GoogleAdd )
                            {
                                OutlookSync.Syncer.SetOutlookWorkingFolder( autoSyncEvent.Pair.OutlookId );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( autoSyncEvent.Pair.GoogleId );
                                Archiver.Instance.CurrentPair = autoSyncEvent.Pair;

                                var calendarItem = OutlookSync.Syncer.FindEventByEntryId( autoSyncEvent.EntryId );
                                if ( calendarItem != null )
                                    GoogleSync.Syncer.AddAppointment( calendarItem );

                                OutlookSync.Syncer.SetOutlookWorkingFolder( "", true );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( "", true );
                            } else if ( autoSyncEvent.Action == CalendarItemAction.GoogleUpdate )
                            {
                                OutlookSync.Syncer.SetOutlookWorkingFolder( autoSyncEvent.Pair.OutlookId );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( autoSyncEvent.Pair.GoogleId );
                                Archiver.Instance.CurrentPair = autoSyncEvent.Pair;

                                var calendarItem = OutlookSync.Syncer.FindEventByEntryId( autoSyncEvent.EntryId );
                                if ( calendarItem != null )
                                    GoogleSync.Syncer.UpdateAppointment( calendarItem );

                                OutlookSync.Syncer.SetOutlookWorkingFolder( "", true );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( "", true );
                            } else if ( autoSyncEvent.Action == CalendarItemAction.GoogleDelete )
                            {
                                OutlookSync.Syncer.SetOutlookWorkingFolder( autoSyncEvent.Pair.OutlookId );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( autoSyncEvent.Pair.GoogleId );
                                Archiver.Instance.CurrentPair = autoSyncEvent.Pair;

                                var calendarItem = OutlookSync.Syncer.FindEventByEntryId( autoSyncEvent.EntryId );
                                if ( calendarItem != null )
                                    GoogleSync.Syncer.DeleteAppointment( calendarItem );

                                OutlookSync.Syncer.SetOutlookWorkingFolder( "", true );
                                GoogleSync.Syncer.SetGoogleWorkingFolder( "", true );
                            }
                        }

                        m_autoSyncEvents.Clear();
                    }

                    Thread.Sleep( 30000 );
                }
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
                    var writer = new StreamWriter( DateFilePath );
                    serializer.Serialize( writer, m_tasks );
                    writer.Close();

                    if ( !m_thread.IsAlive && runThread )
                        m_thread.Start();

                } else if ( File.Exists( DateFilePath ) )
                    File.Delete( DateFilePath );

            } catch ( Exception ex )
            {
                Debug.WriteLine( ex );
            }
        }

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
                var reader = new FileStream( DateFilePath, FileMode.Open);

                if ( m_tasks != null )
                {
                    m_tasks.Clear();
                    m_tasks = null;
                }

                m_tasks = (List<SchedulerTask>)serializer.Deserialize( reader );

                reader.Close();

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

        public SchedulerTask()
        {
            Pair = null;
            Event = SchedulerEvent.Manually;
            TimeSpan = 0;
            LastRunTime = DateTime.MinValue;
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
