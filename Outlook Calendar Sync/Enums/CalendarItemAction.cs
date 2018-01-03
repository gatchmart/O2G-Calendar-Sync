using System;

namespace Outlook_Calendar_Sync.Enums
{
    [Flags]
    [Serializable]
    public enum CalendarItemAction {
        Nothing = 0,
        GoogleUpdate = 1,
        OutlookUpdate = 2,
        GoogleAdd = 4,
        OutlookAdd = 8,
        GeneratedId = 16,
        ContentsEqual = 32,
        GoogleDelete = 64,
        OutlookDelete = 128
    }
}