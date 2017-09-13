using System;

namespace Outlook_Calendar_Sync.Enums
{
    [Serializable]
    public enum RecurrenceType {
        Daily = 0,
        Weekly = 1,
        Monthly = 2,
        MonthNth = 3,
        Yearly = 5,
        YearNth = 6
    }

}