using System;
using Outlook_Calendar_Sync.Enums;

namespace Outlook_Calendar_Sync.Scheduler {
    [Serializable]
    public class SchedulerTask {
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
        /// This is the next run time for the task. This is in minutes
        /// </summary>
        public int NextRunTime;

        /// <summary>
        /// Allows you to set which calendar takes precedence over the other.
        /// This comes in handy if you want a silent sync.
        /// </summary>
        public Precedence Precedence;

        /// <summary>
        /// Allows you to set if the sync will prompt the user when changes have occurred.
        /// Use this with the precedence property
        /// </summary>
        public bool SilentSync;

        private const int Max_Delay = 10;

        public SchedulerTask() {
            Pair = null;
            Event = SchedulerEvent.Manually;
            TimeSpan = 0;
            LastRunTime = DateTime.MinValue;
            NextRunTime = 1;
            Precedence = Precedence.None;
            SilentSync = false;
        }

        public override string ToString() {
            return string.Format( "{0} <=> {1}", Pair.GoogleName, Pair.OutlookName );
        }

        public void IncreaseDelay()
        {
            if ( NextRunTime < Max_Delay )
                NextRunTime += 1;
        }

        public void ResetDelay()
        {
            NextRunTime = 1;
        }
    }
}
