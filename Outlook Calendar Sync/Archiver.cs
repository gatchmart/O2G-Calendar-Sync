using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Outlook_Calendar_Sync {

    internal class Archiver {

        private readonly string m_filePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + "calendarItems.xml";

        public static Archiver Instance => _instance ?? ( _instance = new Archiver() );

        public SyncPair CurrentPair;

        private static Archiver _instance;

        private Dictionary<SyncPair, List<string>> m_data;

        public Archiver() {
            Load();
        }

        /// <summary>
        /// Loads the XML data file
        /// </summary>
        public void Load() {
            if ( !Directory.Exists( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                    "\\OutlookGoogleSync\\" ) )
                Directory.CreateDirectory( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                           "\\OutlookGoogleSync" );

            if ( File.Exists( m_filePath ) ) {
                var reader = new XmlDocument();
                reader.Load( m_filePath );
                XmlNodeList nodes = reader.GetElementsByTagName( "Calendar" );
                m_data = new Dictionary<SyncPair, List<string>>();

                foreach ( XmlNode node in nodes )
                {
                    var pair = new SyncPair();
                    var list = new List<string>();

                    switch ( node.Name )
                    {
                        case "SyncPair":
                            foreach ( XmlNode childNode in node.ChildNodes )
                            {
                                switch ( childNode.Name )
                                {
                                    case "GoogleName":
                                        pair.GoogleName = childNode.InnerText;
                                        break;

                                    case "GoogleID":
                                        pair.GoogleId = childNode.InnerText;
                                        break;

                                    case "OutlookName":
                                        pair.OutlookName = childNode.InnerText;
                                        break;

                                    case "OutlookID":
                                        pair.OutlookId = childNode.InnerText;
                                        break;
                                }
                            }
                            break;

                        case "Events":
                            list.AddRange( from XmlNode childNode in node.ChildNodes where childNode.Name.Equals( "ID" ) select childNode.InnerText );

                            break;
                    }


                    if ( !pair.IsEmpty() )
                        m_data.Add( pair, list );
                }

            } else
                m_data = new Dictionary<SyncPair, List<string>>();

        }

        public void Save() {
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "\t",
                NewLineChars = Environment.NewLine,
                NewLineHandling = NewLineHandling.Replace,
                CloseOutput = true
            };

            XmlWriter writer = XmlWriter.Create( m_filePath, settings );

            writer.WriteStartDocument();
            writer.WriteStartElement( "Calendars" );
            foreach ( var entry in m_data )
            {
                writer.WriteStartElement( "Calendar" );

                // Write SyncPair Data
                writer.WriteStartElement( "SyncPair" );
                writer.WriteElementString( "GoogleName", entry.Key.GoogleName );
                writer.WriteElementString( "GoogleID", entry.Key.GoogleId );
                writer.WriteElementString( "OutlookName", entry.Key.OutlookName );
                writer.WriteElementString( "OutlookID", entry.Key.OutlookId );
                writer.WriteEndElement();

                // Write all event IDs for this calendar.
                writer.WriteStartElement( "Events" );
                foreach ( var eventId in entry.Value )
                {
                    writer.WriteElementString( "ID", eventId );
                }
                writer.WriteEndElement(); // Events

                writer.WriteEndElement(); // Calendar
            }
            writer.WriteEndElement();
            writer.WriteEndDocument();

            writer.Close();
        }

        public void Add( string id ) {
            if ( m_data.ContainsKey( CurrentPair ) )
                m_data[CurrentPair].Add( id );
            else
                m_data.Add( CurrentPair, new List<string> { id } );
        }

        public void Delete( string id ) {
            if ( Contains( id ) )
                m_data[CurrentPair].Remove( id );
        }

        public bool Contains( string id ) {
            return m_data.ContainsKey( CurrentPair ) && m_data[CurrentPair].Contains( id );
        }
    }
}