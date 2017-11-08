using System;
using System.Diagnostics;
using System.IO;

namespace Outlook_Calendar_Sync {
    public class Log : IDisposable
    {
        protected static Log Instance => _instance ?? ( _instance = new Log() );
        private static Log _instance;

        private readonly string m_logFilePath =
            Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
            "\\OutlookGoogleSync\\Logs\\";

        private readonly StreamWriter m_writer;

#if DEBUG
        public static EventHandler<string> RefreshStream;
        public static string CurrentFileName;
#endif

        public Log()
        {

            if ( !Directory.Exists( m_logFilePath ) )
                Directory.CreateDirectory( m_logFilePath );
            else
                ClearOldLogs();

            CurrentFileName = m_logFilePath + "Log - " + DateTime.Now.ToString( "yyyy-MM-dd HHmm" ) + ".txt";

            m_writer = new StreamWriter( CurrentFileName, false );

        }

        private void ClearOldLogs()
        {
            foreach ( var file in Directory.EnumerateFiles( m_logFilePath ) )
            {
                int start = file.IndexOf( "Log - " ) + 6;
                string dateStr = file.Substring( start, file.Length - start - 9 );
                string[] pieces = dateStr.Split( '-' );
                var date = new DateTime( int.Parse( pieces[0] ), int.Parse( pieces[1] ), int.Parse( pieces[2] ) );

                if ( date < DateTime.Today.Subtract( TimeSpan.FromDays( 30 ) ) )
                    File.Delete( file );
            }
        }

        public static bool Write( string str )
        {
            return Instance.WriteLn( str );
        }

        public static bool Write( Exception ex ) {
            return Instance.WriteLn( ex );
        }

        private bool WriteLn( string str )
        {
            try
            {
#if DEBUG
                Debug.WriteLine( str );
                RefreshStream?.Invoke( this, DateTime.Now.ToShortTimeString() + " - " + str + "\n" );
#endif
                m_writer.WriteLine( DateTime.Now.ToShortTimeString() + " - " + str );
                m_writer.Flush();

                return true;
            } catch ( Exception e )
            {
                Debug.Write( e );
            }

            return false;
        }

        private bool WriteLn( Exception ex ) {
            try
            {
#if DEBUG
                Debug.WriteLine( ex );
                RefreshStream?.Invoke( this, DateTime.Now.ToShortTimeString() + " - " + ex + "\n" );
#endif
                m_writer.WriteLine( DateTime.Now.ToShortTimeString() + " - " + ex );
                m_writer.Flush();

                return true;
            } catch ( Exception e )
            {
                Debug.Write( e );
            }

            return false;
        }

        public void Dispose() {
            m_writer?.Close();
            m_writer?.Dispose();
        }
    }
}
