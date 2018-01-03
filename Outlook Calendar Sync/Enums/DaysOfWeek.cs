using System;

namespace Outlook_Calendar_Sync.Enums
{
    [Flags]
    [Serializable]
    public enum DaysOfWeek {
        Sunday = 1,
        Monday = 2,
        Tuesday = 4,
        Wednesday = 8,
        Thursday = 16,
        Friday = 32,
        Saturday = 64
    }

}