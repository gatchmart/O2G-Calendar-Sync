using System;
using System.Windows.Forms;

namespace Outlook_Calendar_Sync
{
    public partial class DetailsForm : Form
    {
        public DetailsForm()
        {
            InitializeComponent();
        }

        public void SetData( string data )
        {
            var rootItem = new TreeNode("Event");

            // For every property, add a list view entry
            var lines = data.Split( '\n' );
            for ( int i = 1; i < lines.Length-1; i++ )
            { 
                // If the line is empty or starts with a dash ignore it
                if ( lines[i].StartsWith( "-" ) || lines[i].StartsWith( "\r" ))
                    continue;

                // Get the content on the line seperated by a colon
                var item = lines[i].Trim( new char[] { '\t', '\n' } );
                var childNode = new TreeNode( item );

                // We need to expand on the data if its an 'Attendee', 'Identifier', or a 'Reocurrance'
                if ( item.StartsWith( "Attendees" ) || item.StartsWith( "Identifier" ) || item.StartsWith( "Recurrence" ) )
                {
                    int y = i;
                    while ( lines[++y].StartsWith( "\t\t" ) )
                    {
                        childNode.Nodes.Add( lines[y].Trim( new char[] {'\r', '\t', '\n'} ) );
                    }

                    i +=  y - i > 1 ? y - i - 1 : 0;
                }

                rootItem.Nodes.Add( childNode );
            }

            treeView1.Nodes.Add( rootItem );
            rootItem.ExpandAll();
        }

    }
}
 