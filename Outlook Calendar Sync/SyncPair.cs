using System;

namespace Outlook_Calendar_Sync
{
    [Serializable]
    public class SyncPair  {
        public string GoogleName;
        public string GoogleId;
        public string OutlookName;
        public string OutlookId;

        public override string ToString() {
            return GoogleName + " <=> " + OutlookName;
        }

        public override bool Equals(object o)
        {
            return o.ToString().Equals( ToString() );
        }

        protected bool Equals( SyncPair other ) {
            return string.Equals( GoogleName, other.GoogleName ) && string.Equals( GoogleId, other.GoogleId ) && string.Equals( OutlookName, other.OutlookName ) && string.Equals( OutlookId, other.OutlookId );
        }

        public bool IsEmpty() {
            return string.IsNullOrEmpty( GoogleName ) && string.IsNullOrEmpty( GoogleId ) &&
                   string.IsNullOrEmpty( OutlookName ) && string.IsNullOrEmpty( OutlookId );
        }

        public override int GetHashCode() {
            unchecked {
                var hashCode = ( GoogleName != null ? GoogleName.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( GoogleId != null ? GoogleId.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( OutlookName != null ? OutlookName.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( OutlookId != null ? OutlookId.GetHashCode() : 0 );
                return hashCode;
            }
        }

        
    }
}
