using System;

namespace Outlook_Calendar_Sync
{
    /// <summary>
    /// The Addtendee class is used to store addtendee data for appointments. This class is used internally in the plug-in.
    /// </summary>
    [Serializable]
    public sealed class Attendee : IEquatable<Attendee> {
        public string Name { get; set; }
        public string Email { get; set; }
        public bool Required { get; set; }

        public Attendee() {
            Name = "";
            Email = "";
            Required = false;
        }

        public Attendee( string name = "", string email = "", bool required = false ) {
            Name = name;
            Email = email;
            Required = required;
        }

        /// <summary>
        /// Creates and Attendee
        /// </summary>
        /// <param name="outlookAttendee">The Outlook attendee string</param>
        /// <param name="required">is the attendee required</param>
        public Attendee( string outlookAttendee, bool required = false ) {
            // Gregory Atchley-Martin ( gatchmart@gmail.com )
            var strs = outlookAttendee.Split( '(' );

            Name = strs[0].Trim();
            Email = ( strs.Length > 1 ) ? strs[1].TrimEnd( ')' ).Trim() : "";
            Required = required;
        }

        public bool Equals( Attendee other )
        {
            return other != null && Name.Equals( other.Name );
        }

        public override string ToString() {
            return $"Name: {Name}, E-Mail: {Email}, Required: {( Required ? "Yes" : "No" )}";
        }
    }

}