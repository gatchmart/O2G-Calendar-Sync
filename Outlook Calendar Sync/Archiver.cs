using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace Outlook_Calendar_Sync {

    internal class Archiver {

        private readonly string m_filePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + "calendarItems.xml";

        public static Archiver Instance => _instance ?? ( _instance = new Archiver() );

        public SyncPair CurrentPair;

        private static Archiver _instance;

        private SerializableDictionary<SyncPair, List<Identifier>> m_data;

        public Archiver() {
            Load();
        }

        /// <summary>
        /// Loads the XML data file
        /// </summary>
        public void Load()
        {
            
            if ( !Directory.Exists( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" ) )
            {
                Directory.CreateDirectory( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                           "\\OutlookGoogleSync" );
                Log.Write( "Created OutlookGoogleSync directory." );
            }

            if ( File.Exists( m_filePath ) )
            {
                Log.Write( "Starting loading Archiver data file..." );
                var serializer = new XmlSerializer( typeof( SerializableDictionary<SyncPair, List<Identifier>> ) );
                var reader = new FileStream( m_filePath, FileMode.Open );
                if ( m_data != null )
                {
                    m_data.Clear();
                    m_data = null;
                }

                m_data = (SerializableDictionary<SyncPair, List<Identifier>>)serializer.Deserialize( reader );

                reader.Close();
                Log.Write( "Completed loading Archiver data file" );

            }
            else
            {
                Log.Write( "No Archiver data file found creating an empty list" );
                m_data = new SerializableDictionary<SyncPair, List<Identifier>>();
            }

        }

        public void Save()
        {
            if ( m_data != null && m_data.Count > 0 )
            {
                Log.Write( "Writing to Archiver data file..." );
                
                var serializer = new XmlSerializer( typeof( SerializableDictionary<SyncPair, List<Identifier>> ) );
                var writer = new StreamWriter( m_filePath );
                serializer.Serialize( writer, m_data );
                writer.Close();

                Log.Write( "Finished writing Archiver data file." );
            }
            else if ( File.Exists( m_filePath ) )
                File.Delete( m_filePath );
        }

        public List<Identifier> GetListForSyncPair( SyncPair pair )
        {
            return m_data.ContainsKey( pair ) ? m_data[pair] : null;
        }

        public void Add( Identifier id ) {
            if ( m_data.ContainsKey( CurrentPair ) )
                m_data[CurrentPair].Add( id );
            else
                m_data.Add( CurrentPair, new List<Identifier> { id } );
        }

        public void Delete( Identifier id ) {
            if ( Contains( id ) )
                m_data[CurrentPair].Remove( id );
        }

        public bool Contains( Identifier id ) {
            return m_data.ContainsKey( CurrentPair ) && m_data[CurrentPair].Contains( id );
        }

        /// <summary>
        /// Updates an Identifier
        /// </summary>
        /// <param name="oldId">The previous CalendarItem Identifier</param>
        /// <param name="newId">The new CalendarItem Identifier</param>
        /// <returns>true is the update was successful</returns>
        public void UpdateIdentifier( Identifier oldId, Identifier newId )
        {
            if ( m_data.ContainsKey( CurrentPair ) )
            {
                if ( m_data[CurrentPair].Contains( oldId ) )
                    m_data[CurrentPair].Remove( oldId );

                m_data[CurrentPair].Add( newId );
            }
            else
                m_data.Add( CurrentPair, new List<Identifier> { newId } );
        }

        public Identifier FindIdentifier( string id )
        {
            return m_data.SelectMany( pair => pair.Value ).FirstOrDefault( identifier => identifier.PartialCompare( id ) );
        }

    }
}