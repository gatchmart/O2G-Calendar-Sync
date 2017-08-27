using System;

namespace Outlook_Calendar_Sync.Scheduler {
    [Serializable]
    public class AutoSyncEvent : IEquatable<AutoSyncEvent> {
        public SyncPair Pair;
        public string EntryId;
        public CalendarItemAction Action;

        public AutoSyncEvent() {
            Pair = null;
            EntryId = null;
            Action = CalendarItemAction.Nothing;
        }

        public bool Equals( AutoSyncEvent other ) {
            if ( ReferenceEquals( null, other ) ) return false;
            if ( ReferenceEquals( this, other ) ) return true;
            return Equals( Pair, other.Pair ) && string.Equals( EntryId, other.EntryId ) && Action == other.Action;
        }

        public override bool Equals( object obj ) {
            if ( ReferenceEquals( null, obj ) ) return false;
            if ( ReferenceEquals( this, obj ) ) return true;
            if ( obj.GetType() != this.GetType() ) return false;
            return Equals( (AutoSyncEvent)obj );
        }

        public override int GetHashCode() {
            unchecked
            {
                var hashCode = ( Pair != null ? Pair.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( EntryId != null ? EntryId.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ (int)Action;
                return hashCode;
            }
        }
    }
}
