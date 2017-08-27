using System;

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
        /// Allows you to set which calendar takes precedence over the other.
        /// This comes in handy if you want a silent sync.
        /// </summary>
        public Precedence Precedence;

        /// <summary>
        /// Allows you to set if the sync will prompt the user when changes have occurred.
        /// Use this with the precedence property
        /// </summary>
        public bool SilentSync;

        public SchedulerTask() {
            Pair = null;
            Event = SchedulerEvent.Manually;
            TimeSpan = 0;
            LastRunTime = DateTime.MinValue;
            Precedence = Precedence.None;
            SilentSync = false;
        }

        public override string ToString() {
            return string.Format( "{0} <=> {1}", Pair.GoogleName, Pair.OutlookName );
        }
    }
}
