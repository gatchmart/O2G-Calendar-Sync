using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Outlook_Calendar_Sync {

    internal class Archiver {

        private string m_filePath = Application.UserAppDataPath + "\\" + "calendarItems.txt";

        public static Archiver Instance => _instance ?? ( _instance = new Archiver() );

        private static Archiver _instance;

        private List<string> m_data;

        public Archiver() {
            Load();
        }

        public void Load() { 

            if ( File.Exists( m_filePath ) ) {
                var fileContents = File.ReadAllText( m_filePath );
                var re = Path.GetFullPath( m_filePath );

                m_data = JsonConvert.DeserializeObject<List<string>>( fileContents ) ?? new List<string>();
            } else
                m_data = new List<string>();

        }

        public void Save() {
            var data = JsonConvert.SerializeObject( m_data );
            File.WriteAllText( m_filePath, data );
        }

        public void Add( string id ) {
            m_data.Add( id );
        }

        public void Delete( string id ) {
            if ( Contains( id ) )
                m_data.Remove( id );
        }

        public bool Contains( string id ) {
            return m_data.Contains( id );
        }
    }
}