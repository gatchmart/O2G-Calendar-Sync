using System;

namespace Outlook_Calendar_Sync
{
    [Serializable]
    public class Identifier : IEquatable<Identifier>
    {
        public string GoogleId { get; set; }
        public string GoogleICalUId { get; set; }
        public string OutlookEntryId { get; set; }
        public string OutlookGlobalId { get; set; }

        public Identifier()
        {
            GoogleId = "";
            GoogleICalUId = "";
            OutlookEntryId = "";
            OutlookGlobalId = "";
        }

        public Identifier( string gid, string guid, string oeid, string osid )
        {
            GoogleId = gid;
            GoogleICalUId = guid;
            OutlookEntryId = oeid;
            OutlookGlobalId = osid;
        }

        public override string ToString()
        {
            return
                $"\tIdentifier:\n\t\tGoogleId: {GoogleId}\n\t\tGoogle iCalUID: {GoogleICalUId}\n\t\tOutlook EntryID: {OutlookEntryId}";
        }
        
        public bool PartialCompare( string id )
        {
            return GoogleId.Equals( id ) || GoogleICalUId.Equals( id ) || OutlookEntryId.Equals( id ) ||
                   OutlookGlobalId.Equals( id );
        }

        public bool Equals( Identifier other )
        {
            if ( ReferenceEquals( null, other ) ) return false;
            if ( ReferenceEquals( this, other ) ) return true;
            return string.Equals( GoogleId, other.GoogleId ) && string.Equals( GoogleICalUId, other.GoogleICalUId ) && string.Equals( OutlookEntryId, other.OutlookEntryId ) && string.Equals( OutlookGlobalId, other.OutlookGlobalId );
        }

        public override bool Equals( object obj )
        {
            if ( ReferenceEquals( null, obj ) ) return false;
            if ( ReferenceEquals( this, obj ) ) return true;
            if ( obj.GetType() != this.GetType() ) return false;
            return Equals( (Identifier) obj );
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = ( GoogleId != null ? GoogleId.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( GoogleICalUId != null ? GoogleICalUId.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( OutlookEntryId != null ? OutlookEntryId.GetHashCode() : 0 );
                hashCode = ( hashCode * 397 ) ^ ( OutlookGlobalId != null ? OutlookGlobalId.GetHashCode() : 0 );
                return hashCode;
            }
        }
    }
}
