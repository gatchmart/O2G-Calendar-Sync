using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;

namespace Outlook_Calendar_Sync
{
    public partial class SchedulerForm : Form
    {

        private CalendarList m_googleFolders;
        private List<OutlookFolder> m_outlookFolders;
        private Scheduler m_scheduler;

        public SchedulerForm()
        {
            InitializeComponent();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            if ( GoogleSync.Syncer.PerformAuthentication() )
            {
                // Get the list of Google Calendars and load them into googleCal_CB
                m_googleFolders = GoogleSync.Syncer.PullCalendars();

                foreach ( var calendarListEntry in m_googleFolders.Items )
                    GoogleCal_CB.Items.Add( calendarListEntry.Summary );

                // Get the list of Outlook Calendars and load them into the outlookCal_CB
                m_outlookFolders = OutlookSync.Syncer.PullCalendars();

                foreach ( var folder in m_outlookFolders )
                    OutlookCal_CB.Items.Add( folder.Name );
            }

            m_scheduler = Scheduler.Instance;

            foreach ( var task in m_scheduler )
            {
                Calendars_LB.Items.Add( task.ToString() );
            }

        }

        private void RemoveTask_BTN_Click( object sender, EventArgs e ) {
            if ( Calendars_LB.SelectedIndex >= 0 && Calendars_LB.SelectedIndex <= m_scheduler.Count )
            {
                m_scheduler.RemoveTask( m_scheduler[Calendars_LB.SelectedIndex] );
                Calendars_LB.Items.RemoveAt( Calendars_LB.SelectedIndex );
            }
        }

        private void AddTask_BTN_Click( object sender, EventArgs e ) {
            SchedulerTask task = new SchedulerTask();

            if ( Event_CB.SelectedIndex >= 0 && Event_CB.SelectedIndex <= 6 )
                task.Event = (SchedulerEvent) Event_CB.SelectedIndex;

            if ( GoogleCal_CB.SelectedIndex >= 0 && OutlookCal_CB.SelectedIndex >= 0 )
            {
                var pair = new SyncPair();
                var googleCal = m_googleFolders.Items[GoogleCal_CB.SelectedIndex];
                var outlookCal = m_outlookFolders[OutlookCal_CB.SelectedIndex];

                pair.GoogleName = googleCal.Summary;
                pair.GoogleId = googleCal.Id;
                pair.OutlookName = outlookCal.Name;
                pair.OutlookId = outlookCal.EntryID;

                task.Pair = pair;

            } else
                MessageBox.Show( this, "You need to select a calendar from both lists.", "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error );

            if ( task.Event == SchedulerEvent.CustomTime )
                task.TimeSpan = int.Parse( Time_TB.Text );

            task.LastRunTime = DateTime.MinValue;

            m_scheduler.AddTask( task );
            Calendars_LB.Items.Add( task.ToString() );
        }

        private void Save_BTN_Click( object sender, EventArgs e )
        {
            m_scheduler.Save();
        }
    }
}
