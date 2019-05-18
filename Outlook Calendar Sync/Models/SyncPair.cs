using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity.Core.Metadata.Edm;

namespace Outlook_Calendar_Sync
{
    [Serializable]
    public class SyncPair  {

        [Key]
        public int SyncPairId { get; set; }
        public string GoogleName { get; set; }
        public string GoogleId { get; set; }
        public string OutlookName { get; set; }
        public string OutlookId { get; set; }

        public List<Identifier> Identifiers { get; set; }

        public SyncPair()
        {
            this.Identifiers = new List<Identifier>();
        }

        public override string ToString() {
            return GoogleName + " <=> " + OutlookName;
        }

        public override bool Equals(object o)
        {
            return o.ToString().Equals( ToString() );
        }

        protected bool Equals( SyncPair other ) {
            return string.Equals( GoogleName, other.GoogleName ) && string.Equals( GoogleId, other.GoogleId ) && string.Equals( OutlookName, other.OutlookName ) && string.Equals( OutlookId, other.OutlookId ) && SyncPairId == other.SyncPairId;
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
