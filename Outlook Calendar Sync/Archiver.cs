using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using System.Data.SQLite;
using Outlook_Calendar_Sync.Models;

namespace Outlook_Calendar_Sync {

    internal class Archiver {

        private readonly string m_filePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" + "calendarItems.xml";

        private readonly string m_path = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\Archive.db";
        //private readonly string m_connectionString = $"Data Source={m_path};Version=3;";

        public static Archiver Instance => _instance ?? ( _instance = new Archiver() );

        public SyncPair CurrentPair;

        private static Archiver _instance;

        private Dictionary<SyncPair, List<Identifier>> m_data;

        private ArchiverContext m_context;

        public Archiver() {

            if (!File.Exists(m_path))
            {
                using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
                {
                    connection.Open();

                    var pair = "PRAGMA foreign_keys = off; BEGIN TRANSACTION;" +
                               "CREATE TABLE SyncPairs (" +
                               "SyncPairId INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, " +
                               "GoogleName STRING(1, 160), " +
                               "GoogleId STRING(1, 160), " +
                               "OutlookName STRING(1, 160), " +
                               "OutlookId STRING(1, 160) " +
                               "); COMMIT TRANSACTION; PRAGMA foreign_keys = on;";

                    var sel = "PRAGMA foreign_keys = off;BEGIN TRANSACTION; " +
                              "CREATE TABLE Identifiers (" +
                              "Id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, "+
                              "SyncPair INTEGER REFERENCES SyncPairs ON DELETE CASCADE NOT NULL, "+
                              "GoogleId STRING(160), "+
                              "GoogleICalUId STRING(160), " +
                              "OutlookEntryId STRING(200), " +
                              "OutlookGlobalId STRING(200), " +
                              "EventHash STRING(64) " +
                              "); COMMIT TRANSACTION;PRAGMA foreign_keys = on;";

                    var cmd = new SQLiteCommand(pair, connection);
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = sel;
                    cmd.ExecuteNonQuery();

                    connection.Close();
                }
            }

            //m_context = new ArchiverContext( $"Data Source={m_path};" );
            m_data = new Dictionary<SyncPair, List<Identifier>>();

            Load();
        }

        /// <summary>
        /// Loads the XML data file
        /// </summary>
        public void Load()
        {
            using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
            {
                connection.Open();
                var syncCmd = new SQLiteCommand("SELECT SyncPairId, GoogleName, GoogleId, OutlookName, OutlookId FROM SyncPairs;", connection);
                var syncReader = syncCmd.ExecuteReader();

                while (syncReader.Read())
                {
                    var pair = new SyncPair();
                    pair.SyncPairId = syncReader.GetInt32(0);
                    pair.GoogleName = syncReader.IsDBNull(1) ? "" : syncReader.GetString(1);
                    pair.GoogleId = syncReader.IsDBNull(2) ? "" : syncReader.GetString(2);
                    pair.OutlookName = syncReader.IsDBNull(3) ? "" : syncReader.GetString(3);
                    pair.OutlookId = syncReader.IsDBNull(4) ? "" : syncReader.GetString(4);

                    var cmd = new SQLiteCommand("SELECT Id, GoogleId, GoogleICalUId, OutlookEntryId, OutlookGlobalId, EventHash FROM Identifiers WHERE SyncPair=" + pair.SyncPairId, connection);
                    var reader = cmd.ExecuteReader();
                    var identifiers = new List<Identifier>();

                    while (reader.Read())
                    {
                        var id = new Identifier();
                        id.Id = reader.GetInt32(0);
                        id.SyncPairId = pair.SyncPairId;
                        id.GoogleId = reader.IsDBNull(1) ? "" : reader.GetString(1);
                        id.GoogleICalUId = reader.IsDBNull(2) ? "" : reader.GetString(2);
                        id.OutlookEntryId = reader.IsDBNull(2) ? "" : reader.GetString(3);
                        id.OutlookGlobalId = reader.IsDBNull(4) ? "" : reader.GetString(4);
                        id.EventHash = reader.IsDBNull(5) ? "" : reader.GetString(5);

                        id.SyncPair = pair;

                        identifiers.Add(id);
                    }

                    m_data.Add(pair, identifiers);
                }

                connection.Close();
            }

            //if ( !Directory.Exists( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) + "\\OutlookGoogleSync\\" ) )
            //{
            //    Directory.CreateDirectory( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
            //                               "\\OutlookGoogleSync" );
            //    Log.Write( "Created OutlookGoogleSync directory." );
            //}

            //if ( File.Exists( m_filePath ) )
            //{
            //    if ( File.Exists( m_filePath + ".bak" ) )
            //    {
            //        File.Delete( m_filePath + ".bak" );
            //        Log.Write( "Deleted the old Archiver backup" );
            //    }

            //    File.Copy( m_filePath, m_filePath + ".bak" );
            //    Log.Write( "Created backup of Archiver Data" );

            //    Log.Write( "Starting loading Archiver data file..." );
            //    var serializer = new XmlSerializer( typeof( SerializableDictionary<SyncPair, List<Identifier>> ) );
            //    var reader = new FileStream( m_filePath, FileMode.Open );
            //    if ( m_data != null )
            //    {
            //        m_data.Clear();
            //        m_data = null;
            //    }

            //    m_data = (SerializableDictionary<SyncPair, List<Identifier>>)serializer.Deserialize( reader );

            //    reader.Close();
                Log.Write( "Completed loading Archiver data file" );

            //}
            //else
            //{
            //    Log.Write( "No Archiver data file found creating an empty list" );
            //    m_data = new SerializableDictionary<SyncPair, List<Identifier>>();
            //}

        }

        public void Save()
        {

            if ( m_data != null && m_data.Count > 0 )
            {
                //Log.Write( "Writing to Archiver data file..." );
                
                //var serializer = new XmlSerializer( typeof( SerializableDictionary<SyncPair, List<Identifier>> ) );
                //var writer = new StreamWriter( m_filePath );
                //serializer.Serialize( writer, m_data );
                //writer.Close();

                //using ( var db = new SQLiteConnection( $"Data Source={m_path};Version=3" ) )
                //{
                //    db.Open();
                //    foreach ( var pairs in m_data )
                //    {
                //        var cmd = new SQLiteCommand( $"SELECT Id FROM SyncPairs WHERE GoogleName='{pairs.Key.GoogleName}' AND GoogleId='{pairs.Key.GoogleId}' AND OutlookName='{pairs.Key.OutlookName}' AND OutlookId='{pairs.Key.OutlookId}';", db);
                //        var id = cmd.ExecuteScalar().ToString();

                //        if ( string.IsNullOrEmpty( id ) )
                //        {
                //            cmd.CommandText =
                //                "INSERT INTO SyncPairs (GoogleName, GoogleId, OutlookName, OutlookId) VALUES(@GoogleName, @GoogleId, @OutlookName, @OutlookId); ";
                //            cmd.Parameters.Add( "@GoogleName", DbType.String, 160 ).Value = pairs.Key.GoogleName;
                //            cmd.ExecuteNonQuery();
                //        }

                //    }
                //    db.Close();
                //}

                Log.Write( "Finished writing Archiver data file." );
            }
            else if ( File.Exists( m_filePath ) )
                File.Delete( m_filePath );
        }

        public List<Identifier> GetListForSyncPair( SyncPair pair )
        {
            //return m_context.Identifiers.Where( x => x.SyncPair.Equals( pair ) ).ToList();
            return m_data.ContainsKey( pair ) ? m_data[pair] : null;
        }

        public void Add( Identifier id )
        {
            //if ( m_context.Pairs.Find( CurrentPair.SyncPairId ) != null )
            //    m_context.Identifiers.Add( id );
            //else
            //{
            //    m_context.Pairs.Add( CurrentPair );
            //    m_context.Identifiers.Add( id );
            //}

            if (m_data.ContainsKey(CurrentPair))
            {
                using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
                {
                    connection.Open();
                    var cmd = new SQLiteCommand("INSERT INTO Identifiers (SyncPair, GoogleId, GoogleICalUId, OutlookEntryId, OutlookGlobalId, EventHash ) " +
                                                "VALUES(@SyncPair, @GoogleId, @GoogleICalUId, @OutlookEntryId, @OutlookGlobalId, @EventHash); ", connection);
                    cmd.Parameters.Add("@SyncPair", DbType.Int32).Value = id.SyncPairId;
                    cmd.Parameters.Add("@GoogleId", DbType.String, 160).Value = id.GoogleId;
                    cmd.Parameters.Add("@GoogleICalUId", DbType.String, 160).Value = id.GoogleICalUId;
                    cmd.Parameters.Add("@OutlookEntryId", DbType.String, 140).Value = id.OutlookEntryId;
                    cmd.Parameters.Add("@OutlookGlobalId", DbType.String, 112).Value = id.OutlookGlobalId;
                    cmd.Parameters.Add("@EventHash", DbType.String, 64).Value = id.EventHash;

                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "SELECT last_insert_rowid()";
                    var i = (int)cmd.ExecuteScalar();

                    id.Id = i;

                    connection.Close();
                }

                m_data[CurrentPair].Add(id);
            }
            else
            {
                using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
                {
                    connection.Open();

                    var cmd = new SQLiteCommand("INSERT INTO SyncPairs (GoogleName, GoogleId, OutlookName, OutlookId) VALUES ( " +
                                                "@GoogleName, @GoogleId, @OutlookName, @OutlookId); ", connection);
                    cmd.Parameters.Add("@GoogleName", DbType.String, 160).Value = CurrentPair.GoogleName;
                    cmd.Parameters.Add("@GoogleId", DbType.String, 160).Value = CurrentPair.GoogleId;
                    cmd.Parameters.Add("@OutlookName", DbType.String, 160).Value = CurrentPair.OutlookName;
                    cmd.Parameters.Add("@OutlookId", DbType.String, 160).Value = CurrentPair.OutlookId;

                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "SELECT last_insert_rowid()";
                    var i = cmd.ExecuteScalar();
                    CurrentPair.SyncPairId = Convert.ToInt32(i);

                    cmd.CommandText =
                        "INSERT INTO Identifiers (SyncPair, GoogleId, GoogleICalUId, OutlookEntryId, OutlookGlobalId, EventHash ) " +
                        "VALUES(@SyncPair, @GoogleId, @GoogleICalUId, @OutlookEntryId, @OutlookGlobalId, @EventHash);";
                    cmd.Parameters.Clear();

                    cmd.Parameters.Add("@SyncPair", DbType.Int32).Value = CurrentPair.SyncPairId;
                    cmd.Parameters.Add("@GoogleId", DbType.String, 160).Value = id.GoogleId;
                    cmd.Parameters.Add("@GoogleICalUId", DbType.String, 160).Value = id.GoogleICalUId;
                    cmd.Parameters.Add("@OutlookEntryId", DbType.String, 140).Value = id.OutlookEntryId;
                    cmd.Parameters.Add("@OutlookGlobalId", DbType.String, 112).Value = id.OutlookGlobalId;
                    cmd.Parameters.Add("@EventHash", DbType.String, 64).Value = id.EventHash;

                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "SELECT last_insert_rowid()";
                    i = cmd.ExecuteScalar();
                    id.Id = Convert.ToInt32(i);

                    id.SyncPair = CurrentPair;

                    connection.Close();
                }

                m_data.Add(CurrentPair, new List<Identifier> { id });
            }
                
        }

        public void Delete( Identifier id ) {
            if (Contains(id))
            {
                m_data[CurrentPair].Remove(id);

                using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
                {
                    connection.Open();

                    var cmd = new SQLiteCommand("DELETE FROM Identifiers WHERE Id=@Id", connection);
                    cmd.Parameters.Add("@Id", DbType.Int32).Value = id.Id;

                    cmd.ExecuteNonQuery();

                    connection.Close();
                }
            }
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
            if (m_data.ContainsKey(CurrentPair))
            {
                if (m_data[CurrentPair].Contains(oldId))
                {
                    m_data[CurrentPair].Remove(oldId);

                    newId.Id = oldId.Id;
                    using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
                    {
                        connection.Open();
                        var cmd = new SQLiteCommand(
                            "UPDATE Identifiers SET GoogleId = @GoogleId, GoogleICalUId = @GoogleICalUId, OutlookEntryId = @OutlookEntryId, " +
                            "OutlookGlobalId = @OutlookGlobalId, EventHash = @EventHash WHERE Id = @Id ", connection);

                        cmd.Parameters.Add("@Id", DbType.Int32).Value = newId.Id;
                        cmd.Parameters.Add("@GoogleId", DbType.String, 160).Value = newId.GoogleId;
                        cmd.Parameters.Add("@GoogleICalUId", DbType.String, 160).Value = newId.GoogleICalUId;
                        cmd.Parameters.Add("@OutlookEntryId", DbType.String, 140).Value = newId.OutlookEntryId;
                        cmd.Parameters.Add("@OutlookGlobalId", DbType.String, 112).Value = newId.OutlookGlobalId;
                        cmd.Parameters.Add("@EventHash", DbType.String, 64).Value = newId.EventHash;

                        cmd.ExecuteNonQuery();

                        connection.Close();
                    }

                }
                else
                {
                    using (var connection = new SQLiteConnection($"Data Source={m_path};Version=3"))
                    {
                        connection.Open();
                        var cmd = new SQLiteCommand("INSERT INTO Identifiers (SyncPair, GoogleId, GoogleICalUId, OutlookEntryId, OutlookGlobalId, EventHash ) " +
                            "VALUES(@SyncPair, @GoogleId, @GoogleICalUId, @OutlookEntryId, @OutlookGlobalId, @EventHash);", connection);
                        
                        cmd.Parameters.Add("@SyncPair", DbType.Int32).Value = CurrentPair.SyncPairId;
                        cmd.Parameters.Add("@GoogleId", DbType.String, 160).Value = newId.GoogleId;
                        cmd.Parameters.Add("@GoogleICalUId", DbType.String, 160).Value = newId.GoogleICalUId;
                        cmd.Parameters.Add("@OutlookEntryId", DbType.String, 140).Value = newId.OutlookEntryId;
                        cmd.Parameters.Add("@OutlookGlobalId", DbType.String, 112).Value = newId.OutlookGlobalId;
                        cmd.Parameters.Add("@EventHash", DbType.String, 64).Value = newId.EventHash;

                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "SELECT last_insert_rowid()";
                        var i = cmd.ExecuteScalar();
                        newId.Id = Convert.ToInt32(i);

                        connection.Close();
                    }
                    
                }

                newId.SyncPair = CurrentPair;
                m_data[CurrentPair].Add(newId);
            }
            else
                Add( newId );
        }

        public Identifier FindIdentifier( string id )
        {
            //return m_context.Identifiers.FirstOrDefault( i => i.PartialCompare( id ) );
            return m_data.SelectMany( pair => pair.Value ).FirstOrDefault( identifier => identifier.PartialCompare( id ) );
        }

        /// <summary>
        /// This method will try to get a SyncPair with the matching identifiers.
        /// returns null if one does not exist.
        /// </summary>
        /// <param name="googldId"></param>
        /// <param name="outlookId"></param>
        /// <returns></returns>
        public SyncPair TryGetSyncPair(string googldId, string outlookId)
        {
            return m_data.Keys.FirstOrDefault(x => x.GoogleId.Equals(googldId) && x.OutlookId.Equals(outlookId));
        }

    }
}