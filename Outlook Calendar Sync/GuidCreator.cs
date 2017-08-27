using System;

namespace Outlook_Calendar_Sync {
    public static class GuidCreator {
        public static string Create()
        {
            return Guid.NewGuid().ToString().Replace( "-", "" );
        }
    }
}
