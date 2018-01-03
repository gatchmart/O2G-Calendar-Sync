using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace Outlook_Calendar_Sync
{
    public partial class DebuggingForm : Form
    {
#if DEBUG
        private readonly string m_basePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                             "\\OutlookGoogleSync\\";

        private Dictionary<string, string> m_fileContents;

        public DebuggingForm()
        {
            InitializeComponent();
        }

        private void Load_BTN_Click( object sender, EventArgs e )
        {
            if ( FileSelect_CB.SelectedIndex >= 0 && FileSelect_CB.SelectedIndex < m_fileContents.Count )
            {
                var select = FileSelect_CB.SelectedItem.ToString();
                Data_RTB.Text = m_fileContents[select];
            }
        }

        private void DebuggingForm_Load( object sender, EventArgs e )
        {
            m_fileContents = new Dictionary<string, string>();

            var files = Directory.GetFiles( m_basePath );
            var logFile = Directory.GetFiles( m_basePath + "Logs\\" );

            var logFileKey = "View Log Stream";
            FileSelect_CB.Items.Add( logFileKey );
            FileSelect2_CB.Items.Add( logFileKey );
            m_fileContents.Add( logFileKey, "" );

            Log.RefreshStream += delegate( object o, string args )
            {
                m_fileContents[logFileKey] += args;

                if ( FileSelect_CB.SelectedItem.ToString().Equals( logFileKey ) )
                    Data_RTB.Text = m_fileContents[logFileKey];
                if ( FileSelect2_CB.SelectedItem.ToString().Equals( logFileKey ) )
                    Data2_RTB.Text = m_fileContents[logFileKey];
            };

            foreach ( var file in files )
            {
                var name = Path.GetFileName( file );
                m_fileContents.Add( name, File.ReadAllText( file ) );
                FileSelect_CB.Items.Add( name );
                FileSelect2_CB.Items.Add( name );
            }

            foreach ( var s in logFile )
            {
                if ( s.Equals( Log.CurrentFileName ) )
                    continue;

                var name = Path.GetFileName( s );
                m_fileContents.Add( name, File.ReadAllText( s ) );
                FileSelect_CB.Items.Add( name );
                FileSelect2_CB.Items.Add( name );
            }
        }

        private void Load2_BTN_Click( object sender, EventArgs e ) {
            if ( Compare_SC.Panel2Collapsed )
                return;

            if ( FileSelect2_CB.SelectedIndex >= 0 && FileSelect2_CB.SelectedIndex < m_fileContents.Count )
            {
                var select = FileSelect2_CB.SelectedItem.ToString();
                Data2_RTB.Text = m_fileContents[select];
            }
        }

        private void Compare_CB_CheckedChanged( object sender, EventArgs e ) {
            Compare_SC.Panel2Collapsed = !Compare_CB.Checked;
        }
#endif
    }
}
