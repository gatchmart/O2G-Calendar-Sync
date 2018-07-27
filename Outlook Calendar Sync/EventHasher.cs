using System;
using System.Security.Cryptography;
using System.Text;

namespace Outlook_Calendar_Sync
{
    public static class EventHasher
    {
        public static string GetHash( CalendarItem evnt )
        {
            return CreateHash( evnt.GetHasherString() );
        }

        public static string GetHash( string evnt )
        {
            return CreateHash( evnt );
        }

        public static string GetHash( object item )
        {
            return CreateHash( item.ToString() );
        }

        // Verify a hash against a string.
        public static bool CompareHash( string input, string hash )
        {
            // Hash the input.
            string hashOfInput = CreateHash( input );

            // Create a StringComparer an compare the hashes.
            StringComparer comparer = StringComparer.OrdinalIgnoreCase;

            return 0 == comparer.Compare( hashOfInput, hash );
        }

        public static bool Equals( string hash1, string hash2 )
        {
            // Create a StringComparer an compare the hashes.
            StringComparer comparer = StringComparer.OrdinalIgnoreCase;

            return 0 == comparer.Compare( hash1, hash2 );
        }

        private static string CreateHash( string evnt )
        {
            var mySha256 = SHA256.Create();
            var hashBytes = mySha256.ComputeHash( Encoding.UTF8.GetBytes( evnt ) );

            var builder = new StringBuilder();

            foreach ( var t in hashBytes )
            {
                builder.Append( t.ToString( "x2" ) );
            }

            return builder.ToString();
        }
    }
}
