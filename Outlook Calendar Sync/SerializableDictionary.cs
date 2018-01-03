using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace Outlook_Calendar_Sync
{
    [XmlRoot("Dictionary")]
    [Serializable]
    public class SerializableDictionary<TKey, TValue> : Dictionary<TKey, TValue>, IXmlSerializable
    {
        public XmlSchema GetSchema()
        {
            // This method should always return null
            return null;
        }

        public void ReadXml( XmlReader reader )
        {
            var keySerializer = new XmlSerializer( typeof(TKey) );
            var valueSerializer = new XmlSerializer( typeof(TValue) );
            bool wasEmpty = reader.IsEmptyElement;
            reader.Read();
            if ( wasEmpty )
                return;

            while ( reader.NodeType != XmlNodeType.EndElement )
            {
                reader.ReadStartElement("Item");
                reader.ReadStartElement("Key");
                TKey key = (TKey) keySerializer.Deserialize( reader );
                reader.ReadEndElement();
                reader.ReadStartElement("Value");
                TValue value = (TValue) valueSerializer.Deserialize( reader );
                reader.ReadEndElement();
                this.Add( key, value );
                reader.ReadEndElement();
                reader.MoveToContent();
#if DEBUG
                Log.Write( $"Read {key}, {value} from dictionary." );
#endif
            }

            reader.ReadEndElement();
        }

        public void WriteXml( XmlWriter writer )
        {
            var keySerializer = new XmlSerializer( typeof( TKey ) );
            var valueSerializer = new XmlSerializer( typeof( TValue ) );
            foreach ( TKey key in this.Keys )
            {
                writer.WriteStartElement( "Item" );
                writer.WriteStartElement( "Key" );
                keySerializer.Serialize( writer, key );
                writer.WriteEndElement();
                writer.WriteStartElement( "Value" );
                TValue value = this[key];
                valueSerializer.Serialize( writer, value );
                writer.WriteEndElement();
                writer.WriteEndElement();
#if DEBUG
                Log.Write( $"Wrote {key}, {value} from dictionary." );
#endif
            }
        }
    }
}
