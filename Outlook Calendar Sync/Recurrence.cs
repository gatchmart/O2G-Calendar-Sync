using System;
using System.Globalization;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using Outlook_Calendar_Sync.Enums;

namespace Outlook_Calendar_Sync {

    /// <summary>
    /// This class is a middle class for the Google and Outlook recurrance patterns. It can convert between the two and is
    /// used through the add-in as the recurrance storage class.
    /// </summary>
    [Serializable]
    public class Recurrence {

        /// <summary>
        /// Returns or sets an RecurrenceType constant specifying the frequency of occurrences for the recurrence pattern
        /// </summary>
        public RecurrenceType Type { get; set; }

        /// <summary>
        /// Returns or sets an DaysOfWeek constant representing the mask for the days of the week on which the recurring appointment or task occurs.
        /// </summary>
        public DaysOfWeek DaysOfTheWeekMask { get; set; }

        /// <summary>
        /// Returns or sets an Integer (int in C#) value indicating the day of the month on which the recurring appointment or task occurs.
        /// </summary>
        public int DayOfMonth { get; set; }

        /// <summary>
        /// Returns or sets an Integer (int in C#) value indicating the duration (in minutes) of the RecurrencePattern. 
        /// </summary>
        public int Duration { get; set; }

        /// <summary>
        /// Returns or sets an Integer (int in C#) value specifying the number of units of a given recurrence type between occurrences.
        /// </summary>
        /// <remarks>
        /// For example, setting the Interval property to 2 and the RecurrenceType property to Weekly would cause the pattern to occur every second week.
        /// When RecurrenceType is set to YearNth or Year, the Interval property indicates the number of years between occurrences. For example,
        ///  Interval equals 1 indicates the recurrence is every year, Interval equals 2 indicates the recurrence is every 2 years, and so on.
        /// </remarks>
        public int Interval { get; set; }

        /// <summary>
        /// Returns or sets an Integer (int in C#) value specifying the count for which the recurrence pattern is valid for a given interval.
        /// </summary>
        /// <remarks>
        /// This property is only valid for recurrences of the olRecursMonthNth and olRecursYearNth type and allows the definition of a recurrence pattern 
        /// that is only valid for the Nth occurrence, such as "the 2nd Sunday in March" pattern. The count is set numerically: 1 for the first, 2 for
        ///  the second, and so on through 5 for the last. Values greater than 5 will generate errors when the pattern is saved.
        /// </remarks>
        public int Instance { get; set; }

        /// <summary>
        /// Returns or sets an Integer (int in C#) value indicating the number of occurrences of the recurrence pattern. Read/write.
        /// </summary>
        /// <remarks>
        /// This property allows the definition of a recurrence pattern that is only valid for the specified number of subsequent occurrences.
        /// For example, you can set this property to 10 for a formal training course that will be held on the next ten Thursday evenings.
        /// This property must be coordinated with other properties when setting up a recurrence pattern. If the PatternEndDate property or
        ///  the Occurrences property is set, the pattern is considered to be finite and the NoEndDate property is False. 
        /// If neither PatternEndDate nor Occurrences is set, the pattern is considered infinite and NoEndDate is True.
        /// </remarks>
        public int Occurrences { get; set; }

        /// <summary>
        /// The end time for the event
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// The start time for the event
        /// </summary>
        public string StartTime { get; set; }

        /// <summary>
        /// Returns or sets an Integer (int in C#) value indicating which month of the year is valid for the specified recurrence pattern.
        /// </summary>
        /// <remarks>
        /// The value can be a number from 1 through 12. For example, setting this property to 5 and the RecurrenceType property to Yearly would cause this recurrence pattern to occur every May. 
        /// This property is only valid for recurrence patterns whose RecurrenceType property is set to Yearly or YearNth.
        /// </remarks>
        public int MonthOfYear { get; set; }

        /// <summary>
        /// The pattern start DateTime
        /// </summary>
        public string PatternStartDate { get; set; }

        /// <summary>
        /// The pattern end DateTime
        /// </summary>
        public string PatternEndDate { get; set; }

        /// <summary>
        /// Returns a Boolean (bool in C#) value that indicates True if the recurrence pattern has no end date.
        /// </summary>
        public bool NoEndDate { get; set; }

        /// <summary>
        /// Default contructor used for Serialization
        /// </summary>
        public Recurrence()
        {
            DayOfMonth = 0;
            DaysOfTheWeekMask = 0;
            Duration = 0;
            EndTime = "";
            StartTime = "";
            Instance = 0;
            Interval = 0;
            MonthOfYear = 0;
            NoEndDate = false;
            Occurrences = 0;
            PatternStartDate = "";
            PatternEndDate = "";
            Type = 0;
        }

        /// <summary>
        /// Creates a Recurrence object
        /// </summary>
        /// <param name="rrule">The Google recurrence string</param>
        /// <param name="calItem">The CalendarItem currently representing the Google Event</param>
        public Recurrence( string rrule, CalendarItem calItem ) {
            try {
                // "RRULE:FREQ=WEEKLY;UNTIL=20160415T200000Z;BYDAY=WE,FR"
                var items = rrule.Replace( "RRULE:", "" ).Split( ';' );

                // Pull all the information from the RRULE string
                foreach ( var item in items ) {
                    var split = item.Split( '=' );
                    switch ( split[0] ) {
                        case "FREQ":

                            if ( split[1] == "DAILY" )
                                Type = RecurrenceType.Daily;
                            else if ( split[1] == "WEEKLY" )
                                Type = RecurrenceType.Weekly;
                            else if ( split[1] == "MONTHLY" )
                                Type = RecurrenceType.Monthly;
                            else if ( split[1] == "YEARLY" )
                                Type = RecurrenceType.Yearly;

                            break;

                        case "UNTIL":
                            var date = ( calItem.IsAllDayEvent && split[1].Length == 8 )
                                ? DateTime.ParseExact( split[1], "yyyyMMdd", CultureInfo.InvariantCulture )
                                : DateTime.ParseExact( split[1], "yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture );
                            PatternEndDate = date.ToString( "yyyy-MM-dd" );
                            EndTime = date.ToString( "HH:mm:sszzz" );
                            break;

                        case "BYDAY":
                            var days = split[1].Split( ',' );

                            foreach ( var dday in days ) {
                                // This ensures we grab the right day of the month
                                var day = dday;
                                if ( day.Length == 3 ) {
                                    Instance = int.Parse( day.Substring( 0, 1 ) );
                                    day = day.Remove( 0, 1 );
                                } else if ( day.Length == 4 ) {
                                    Instance = int.Parse( day.Substring( 0, 2 ) );
                                    day = day.Remove( 0, 2 );
                                }

                                if ( day.Equals( "MO" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Monday;
                                else if ( day.Equals( "TU" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Tuesday;
                                else if ( day.Equals( "WE" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Wednesday;
                                else if ( day.Equals( "TH" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Thursday;
                                else if ( day.Equals( "FR" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Friday;
                                else if ( day.Equals( "SA" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Saturday;
                                else if ( day.Equals( "SU" ) )
                                    DaysOfTheWeekMask |= DaysOfWeek.Sunday;
                            }

                            break;

                        case "BYMONTH":
                            MonthOfYear = int.Parse( split[1] );
                            break;

                        case "INTERVAL":
                            Interval = int.Parse( split[1] );
                            break;
                        case "COUNT":
                            Occurrences = int.Parse( split[1] );
                            break;
                    } // switch
                } // foreach

                if ( Type == RecurrenceType.Daily && Interval == 0 )
                    Interval = 1;

                if ( Type == RecurrenceType.Weekly && Interval == 0 )
                    Interval = 1;

                if ( Type == RecurrenceType.Monthly && Interval == 0 ) {
                    Interval = 1;
                    if ( Instance == 0 )
                        DayOfMonth = DateTime.Parse( calItem.Start ).Day;
                    else
                        Type = RecurrenceType.MonthNth;
                }

                if ( Type == RecurrenceType.Yearly ) {
                    var date = DateTime.Parse( calItem.Start );
                    DayOfMonth = date.Day;
                    Interval = 12;
                    MonthOfYear = date.Month;
                }

                var startDate = ( calItem.IsAllDayEvent && calItem.Start.Length == 10 )
                    ? DateTime.ParseExact( calItem.Start, "yyyy-MM-dd", CultureInfo.InvariantCulture )
                    : DateTime.ParseExact( calItem.Start, "yyyy-MM-ddTHH:mm:sszzz", CultureInfo.InvariantCulture );
                PatternStartDate = startDate.ToString( "yyyy-MM-dd" );
                StartTime = startDate.ToString( "HH:mm:sszzz" );

                if ( Type == RecurrenceType.Weekly && DaysOfTheWeekMask == 0 )
                {
                    DaysOfTheWeekMask = (DaysOfWeek) ( 1 << (int) startDate.DayOfWeek );
                }

                // If there is no specified end date set NoEndDate to true
                if ( string.IsNullOrEmpty( PatternEndDate ) )
                    NoEndDate = true;

                if ( string.IsNullOrEmpty( EndTime ) && !calItem.IsAllDayEvent && calItem.End.Length > 10 )
                    EndTime = calItem.End.Substring( 11 );

            } catch ( Exception ex ) {
                Log.Write( "An Error occured when trying to parse Google Recurrence data, " + ex.Message );
            }
        }

        /// <summary>
        /// Creates a Recurrence object
        /// </summary>
        /// <param name="pattern">The Outlook RecurrencePattern to use</param>
        public Recurrence( RecurrencePattern pattern ) {
            DayOfMonth = pattern.DayOfMonth;
            DaysOfTheWeekMask = (DaysOfWeek)pattern.DayOfWeekMask;
            Duration = pattern.Duration;
            EndTime = pattern.EndTime.ToString( "HH:mm:sszzz" );
            StartTime = pattern.StartTime.ToString( "HH:mm:sszzz" );
            Instance = pattern.Instance;
            Interval = pattern.Interval;
            MonthOfYear = pattern.MonthOfYear;
            NoEndDate = pattern.NoEndDate;
            Occurrences = pattern.Occurrences;
            PatternStartDate = pattern.PatternStartDate.ToString( "yyyy-MM-dd" );
            PatternEndDate = pattern.PatternEndDate.ToString( "yyyy-MM-dd" );
            Type = (RecurrenceType)pattern.RecurrenceType;
        }

        public void GetOutlookPattern( ref RecurrencePattern pattern, bool isAllDay ) {
            switch ( Type ) {
                case RecurrenceType.Daily:
                    pattern.RecurrenceType = OlRecurrenceType.olRecursDaily;

                    AddOutlookRecurrenceData( ref pattern, isAllDay );
                    break;
                case RecurrenceType.Weekly:
                    pattern.RecurrenceType = OlRecurrenceType.olRecursWeekly;
                    pattern.DayOfWeekMask = (OlDaysOfWeek) DaysOfTheWeekMask;

                    AddOutlookRecurrenceData( ref pattern, isAllDay );
                    break;
                case RecurrenceType.Monthly:
                    pattern.RecurrenceType = OlRecurrenceType.olRecursMonthly;
                    if ( DayOfMonth != 0 )
                        pattern.DayOfMonth = DayOfMonth;

                    AddOutlookRecurrenceData( ref pattern, isAllDay );
                    break;
                case RecurrenceType.MonthNth:
                    pattern.RecurrenceType = OlRecurrenceType.olRecursMonthNth;
                    pattern.DayOfWeekMask = (OlDaysOfWeek)DaysOfTheWeekMask;

                    AddOutlookRecurrenceData( ref pattern, isAllDay );
                    if ( Instance != 0 )
                        pattern.Instance = Instance;
                    break;
                case RecurrenceType.Yearly:
                    pattern.RecurrenceType = OlRecurrenceType.olRecursYearly;
                    if ( DayOfMonth != 0 )
                        pattern.DayOfMonth = DayOfMonth;
                    if ( MonthOfYear != 0 )
                        pattern.MonthOfYear = MonthOfYear;

                    AddOutlookRecurrenceData( ref pattern, isAllDay );
                    break;
                case RecurrenceType.YearNth:
                    pattern.RecurrenceType = OlRecurrenceType.olRecursYearNth;
                    pattern.DayOfWeekMask = (OlDaysOfWeek)DaysOfTheWeekMask;

                    AddOutlookRecurrenceData( ref pattern, isAllDay );
                    if ( Instance != 0 )
                        pattern.Instance = Instance;
                    break;
            }
        }

        public string GetGoogleRecurrenceString() {
            var builder = new StringBuilder();

            // RRULE:FREQ=WEEKLY;UNTIL=20160415T200000Z;BYDAY=WE,FR
            builder.Append( "RRULE:" );

            builder.Append( "FREQ=" );

            switch ( Type ) {
                case RecurrenceType.Daily:
                    builder.Append( "DAILY;" );
                    break;
                case RecurrenceType.Weekly:
                    builder.Append( "WEEKLY;" );
                    break;
                case RecurrenceType.Monthly:
                    builder.Append( "MONTHLY;" );
                    break;
                case RecurrenceType.MonthNth:
                    builder.Append( "MONTHLY;" );
                    break;
                case RecurrenceType.Yearly:
                    builder.Append( "YEARLY;" );
                    break;
                case RecurrenceType.YearNth:
                    builder.Append( "YEARLY;" );
                    break;
            }

            if ( PatternEndDate != null && !NoEndDate ) {
                builder.Append( "UNTIL=" );
                var date = DateTime.ParseExact( $"{PatternEndDate}T{EndTime}", "yyyy-MM-ddTHH:mm:sszzz", CultureInfo.InvariantCulture );
                builder.Append( date.ToUniversalTime().ToString( "yyyyMMddTHHmmssZ" ) + ";" );
            }

            if ( DaysOfTheWeekMask != 0 ) {
                builder.Append( "BYDAY=" );

                if ( Instance != 0 )
                    builder.Append( Instance );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Sunday ) )
                    builder.Append( "SU," );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Monday ) )
                    builder.Append( "MO," );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Tuesday ) )
                    builder.Append( "TU," );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Wednesday ) )
                    builder.Append( "WE," );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Thursday ) )
                    builder.Append( "TH," );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Friday ) )
                    builder.Append( "FR," );

                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Saturday ) )
                    builder.Append( "SA," );

                builder.Remove( builder.Length - 1, 1 );
                builder.Append( ";" );
            }

            if ( Type == RecurrenceType.Yearly ) {
                builder.Append( "BYMONTH=" );
                builder.Append( MonthOfYear + ";" );
            }

            if ( Interval > 1 && Type != RecurrenceType.Yearly )
                builder.Append( "INTERVAL=" + Interval + ";" );

            builder.Remove( builder.Length - 1, 1 );
            return builder.ToString();
        }

        public string GetPatternStartTimeWithHours() {
            return PatternStartDate + "T" + StartTime;
        }

        public string GetPatternEndTimeWithHours()
        {
            return PatternEndDate + "T" + EndTime;
        }

        public void AdjustRecurrenceOutlookPattern( DateTime start, DateTime end ) {
            var s = start.ToString( "HH:mm:sszzz" );
            var e = end.ToString( "HH:mm:sszzz" );

            PatternStartDate = PatternStartDate.Remove( PatternStartDate.IndexOf( "T" ) );
            PatternStartDate += "T" + s;

            PatternEndDate = PatternEndDate.Remove( PatternEndDate.IndexOf( "T" ) );
            PatternEndDate += "T" + e;

        }

        public override string ToString() {
            StringBuilder builder = new StringBuilder();

            builder.AppendLine( "\t\tPattern Start Date: " + PatternStartDate );
            builder.AppendLine( "\t\tPattern End Date: " + PatternEndDate );
            builder.AppendLine( "\t\tDuration: " + Duration );
            builder.AppendLine( "\t\tOccurrences: " + Occurrences );
            builder.AppendLine( "\t\tInterval: " + Interval );
            builder.AppendLine( "\t\tInstance: " + Instance );
            builder.AppendLine( "\t\tStart Time: " + StartTime );
            builder.AppendLine( "\t\tEnd Time: " + EndTime );
            builder.AppendLine( "\t\tMonth of Year: " + MonthOfYear );
            builder.AppendLine( "\t\tDay of Month: " + DayOfMonth );
            builder.AppendLine( "\t\tNo End Date: " + ( NoEndDate ? "Yes" : "No" ) );

            if ( Type == RecurrenceType.Daily )
                builder.AppendLine( "\t\tDaily Recurrence" );
            else if ( Type == RecurrenceType.MonthNth )
                builder.AppendLine( "\t\tMonthNth Recurrence" );
            else if ( Type == RecurrenceType.Monthly)
                builder.AppendLine( "\t\tMonthly Recurrence" );
            else if ( Type == RecurrenceType.Weekly )
                builder.AppendLine( "\t\tWeekly Recurrence" );
            else if ( Type == RecurrenceType.YearNth )
                builder.AppendLine( "\t\tYearNth Recurrence" );
            else if ( Type == RecurrenceType.Yearly )
                builder.AppendLine( "\t\tYearly Recurrence" );

            if ( DaysOfTheWeekMask != 0 ) {
                builder.Append( "\t\t" );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Monday ) )
                    builder.Append( "Monday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Tuesday ) )
                    builder.Append( "Tuesday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Wednesday ) )
                    builder.Append( "Wednesday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Thursday ) )
                    builder.Append( "Thursday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Friday ) )
                    builder.Append( "Friday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Saturday ) )
                    builder.Append( "Saturday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Sunday ) )
                    builder.Append( "Sunday | " );
                builder.Remove( builder.Length - 2, 2 );
            }

            return builder.ToString();
        }

        public string GetHasherString()
        {
            StringBuilder builder = new StringBuilder();

            builder.Append( "Pattern Start Date: " + PatternStartDate );
            builder.Append( "Pattern End Date: " + PatternEndDate );
            builder.Append( "Duration: " + Duration );
            builder.Append( "Occurrences: " + Occurrences );
            builder.Append( "Interval: " + Interval );
            builder.Append( "Instance: " + Instance );
            builder.Append( "Start Time: " + StartTime );
            builder.Append( "End Time: " + EndTime );
            builder.Append( "Month of Year: " + MonthOfYear );
            builder.Append( "Day of Month: " + DayOfMonth );
            builder.Append( "No End Date: " + ( NoEndDate ? "Yes" : "No" ) );

            if ( Type == RecurrenceType.Daily )
                builder.Append( "Daily Recurrence" );
            else if ( Type == RecurrenceType.MonthNth )
                builder.Append( "MonthNth Recurrence" );
            else if ( Type == RecurrenceType.Monthly )
                builder.Append( "Monthly Recurrence" );
            else if ( Type == RecurrenceType.Weekly )
                builder.Append( "Weekly Recurrence" );
            else if ( Type == RecurrenceType.YearNth )
                builder.Append( "YearNth Recurrence" );
            else if ( Type == RecurrenceType.Yearly )
                builder.Append( "Yearly Recurrence" );

            if ( DaysOfTheWeekMask != 0 )
            {
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Monday ) )
                    builder.Append( "Monday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Tuesday ) )
                    builder.Append( "Tuesday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Wednesday ) )
                    builder.Append( "Wednesday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Thursday ) )
                    builder.Append( "Thursday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Friday ) )
                    builder.Append( "Friday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Saturday ) )
                    builder.Append( "Saturday | " );
                if ( DaysOfTheWeekMask.HasFlag( DaysOfWeek.Sunday ) )
                    builder.Append( "Sunday | " );
                builder.Remove( builder.Length - 2, 2 );
            }

            return builder.ToString();
        }

        public bool Equals( Recurrence other ) {
            bool result = true;

            result &= Type == other.Type;
            result &= DaysOfTheWeekMask == other.DaysOfTheWeekMask;
            result &= DayOfMonth == other.DayOfMonth;
            result &= Duration == other.Duration;
            result &= Interval == other.Interval;
            result &= Instance == other.Instance;
            result &= Occurrences == other.Occurrences;
            result &= MonthOfYear == other.MonthOfYear;
            result &= NoEndDate && other.NoEndDate;

            if ( !string.IsNullOrEmpty( EndTime ) && !string.IsNullOrEmpty( other.EndTime ) )
                result &= EndTime.Equals( other.EndTime );

            if ( !string.IsNullOrEmpty( StartTime ) && !string.IsNullOrEmpty( other.StartTime ) )
                result &= StartTime.Equals( other.StartTime );

            if ( !string.IsNullOrEmpty( PatternEndDate ) && !string.IsNullOrEmpty( other.PatternEndDate ) )
                result &= PatternEndDate.Equals( other.PatternEndDate );

            if ( !string.IsNullOrEmpty( PatternStartDate ) && !string.IsNullOrEmpty( other.PatternStartDate ) )
                result &= PatternStartDate.Equals( other.PatternStartDate );

            return result;
        }

        private void AddOutlookRecurrenceData( ref RecurrencePattern pattern, bool isAllDay ) {
            pattern.Duration = isAllDay ? 1440 : Duration;

            if ( EndTime != null )
                pattern.EndTime = DateTime.Parse( EndTime );

            if ( StartTime != null )
                pattern.StartTime = DateTime.Parse( StartTime );

            if ( Interval != 0 )
                pattern.Interval = Interval;

            pattern.NoEndDate = NoEndDate;

            if ( Occurrences != 0 )
                pattern.Occurrences = Occurrences;

            if ( PatternStartDate != null )
                pattern.PatternStartDate = DateTime.Parse( PatternStartDate );

            if ( PatternEndDate != null && !NoEndDate )
                pattern.PatternEndDate = DateTime.Parse( PatternEndDate );
        }

    }
}
