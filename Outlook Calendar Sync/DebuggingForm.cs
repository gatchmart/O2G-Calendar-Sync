using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Outlook_Calendar_Sync.Properties;

namespace Outlook_Calendar_Sync
{
    public partial class DebuggingForm : Form
    {
#if DEBUG
        private readonly string m_basePath = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) +
                                             "\\OutlookGoogleSync\\";
        private const string LOG_FILE_KEY = "View Log Stream";

        private Dictionary<string, string> m_fileContents;

        private delegate void UpdateIfNeeded();

        private UpdateIfNeeded del;

        public DebuggingForm()
        {
            InitializeComponent();
        }

        private void Load_BTN_Click( object sender, EventArgs e )
        {
            if ( FileSelect_CB.SelectedIndex >= 0 && FileSelect_CB.SelectedIndex < m_fileContents.Count )
            {
                var select = FileSelect_CB.SelectedItem.ToString();
                
                // Grab the initial log stream data 
                if ( select.Equals( LOG_FILE_KEY ) && m_fileContents[select].Length == 0 )
                    m_fileContents[select] = Log.Instance.m_builder.ToString();

                Data_RTB.Text = m_fileContents[select];
            }
        }

        private void DebuggingForm_Load( object sender, EventArgs e )
        {
            m_fileContents = new Dictionary<string, string>();

            var files = Directory.GetFiles( m_basePath );
            var logFile = Directory.GetFiles( m_basePath + "Logs\\" );

            FileSelect_CB.Items.Add( LOG_FILE_KEY );
            FileSelect2_CB.Items.Add( LOG_FILE_KEY );
            m_fileContents.Add( LOG_FILE_KEY, "" );

            Log.RefreshStream += UpdateStream;
            del = delegate
            {
                if (FileSelect_CB.SelectedItem != null)
                    if ( FileSelect_CB.SelectedItem.ToString().Equals( LOG_FILE_KEY ) )
                        Data_RTB.Text = m_fileContents[LOG_FILE_KEY];

                if (FileSelect2_CB.SelectedItem != null)
                    if ( FileSelect2_CB.SelectedItem.ToString().Equals( LOG_FILE_KEY ) )
                        Data2_RTB.Text = m_fileContents[LOG_FILE_KEY];
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

        private void UpdateStream( object sender, string args )
        {
            m_fileContents[LOG_FILE_KEY] = args;

            this.Invoke( del );
        }

        private void Load2_BTN_Click( object sender, EventArgs e ) {
            if ( Compare_SC.Panel2Collapsed )
                return;

            if ( FileSelect2_CB.SelectedIndex >= 0 && FileSelect2_CB.SelectedIndex < m_fileContents.Count )
            {
                var select = FileSelect2_CB.SelectedItem.ToString();
               
                // Grab the initial log stream data 
                if ( select.Equals( LOG_FILE_KEY ) && m_fileContents[select].Length == 0 )
                    m_fileContents[select] = Log.Instance.m_builder.ToString();

                Data2_RTB.Text = m_fileContents[select];
            }
        }

        private void Compare_CB_CheckedChanged( object sender, EventArgs e ) {
            Compare_SC.Panel2Collapsed = !Compare_CB.Checked;
        }

        private void DebuggingForm_FormClosing( object sender, FormClosingEventArgs e )
        {
            // ReSharper disable once DelegateSubtraction
            if ( Log.RefreshStream != null ) Log.RefreshStream -= UpdateStream;
        }


#endif
    }
}
