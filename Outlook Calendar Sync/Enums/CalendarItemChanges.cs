using System;

namespace Outlook_Calendar_Sync.Enums
{
    [Flags]
    [Serializable]
    public enum CalendarItemChanges {
        Nothing = 0,
        StartDate = 1,
        EndDate = 2,
        Location = 4,
        Body = 8,
        Subject = 16,
        StartTimeZone = 32,
        EndTimeZone = 64,
        ReminderTime = 128,
        Attendees = 256,
        Recurrence = 512,
        CalId = 1024
    }
}