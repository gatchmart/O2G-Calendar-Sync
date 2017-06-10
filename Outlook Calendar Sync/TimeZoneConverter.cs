﻿using System;
using System.Linq;

namespace Outlook_Calendar_Sync {
    internal sealed class TimeZoneConverter {
        // This will return the Windows zone that matches the IANA zone, if one exists.
        public static string IanaToWindows( string ianaZoneId ) {
            var utcZones = new[] { "Etc/UTC", "Etc/UCT", "Etc/GMT" };
            if ( utcZones.Contains( ianaZoneId, StringComparer.Ordinal ) )
                return "UTC";

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;

            // resolve any link, since the CLDR doesn't necessarily use canonical IDs
            var links = tzdbSource.CanonicalIdMap
                .Where( x => x.Value.Equals( ianaZoneId, StringComparison.Ordinal ) )
                .Select( x => x.Key );

            // resolve canonical zones, and include original zone as well
            var possibleZones = tzdbSource.CanonicalIdMap.ContainsKey( ianaZoneId )
                ? links.Concat( new[] { tzdbSource.CanonicalIdMap[ianaZoneId], ianaZoneId } )
                : links;

            // map the windows zone
            var mappings = tzdbSource.WindowsMapping.MapZones;
            var item = mappings.FirstOrDefault( x => x.TzdbIds.Any( possibleZones.Contains ) );
            if ( item == null ) return null;
            return item.WindowsId;
        }

        // This will return the "primary" IANA zone that matches the given windows zone.
        // If the primary zone is a link, it then resolves it to the canonical ID.
        public static string WindowsToIana( string windowsZoneId ) {
            if ( windowsZoneId.Equals( "UTC", StringComparison.Ordinal ) )
                return "Etc/UTC";

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;
            var tzi = TimeZoneInfo.FindSystemTimeZoneById( windowsZoneId );
            if ( tzi == null ) return null;
            var tzid = tzdbSource.MapTimeZoneId( tzi );
            if ( tzid == null ) return null;
            return tzdbSource.CanonicalIdMap[tzid];
        }
    }
}
