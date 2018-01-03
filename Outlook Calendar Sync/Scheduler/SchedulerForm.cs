using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;
using Outlook_Calendar_Sync.Enums;

namespace Outlook_Calendar_Sync.Scheduler
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

            if ( SilentSync_CB.Checked )
            {
                task.SilentSync = true;
                task.Precedence = (Precedence)Precedence_CB.SelectedIndex;
            }

            task.LastRunTime = DateTime.MinValue;

            m_scheduler.AddTask( task );
            Calendars_LB.Items.Add( task.ToString() );
        }

        private void Save_BTN_Click( object sender, EventArgs e )
        {
            m_scheduler.Save();
            MessageBox.Show( this, "The Scheduled Events have been saved!", "Scheduled", MessageBoxButtons.OK,
                MessageBoxIcon.Information );
        }

        private void SilentSync_CB_CheckedChanged( object sender, EventArgs e )
        {
            Precedence_CB.Enabled = SilentSync_CB.Checked;
        }

        private void Event_CB_SelectedIndexChanged( object sender, EventArgs e )
        {
            Time_TB.Enabled = Event_CB.SelectedIndex == 5;
        }

        private void Calendars_LB_SelectedIndexChanged( object sender, EventArgs e ) {
            if ( Calendars_LB.SelectedIndex < m_scheduler.Count && Calendars_LB.SelectedIndex >= 0 )
            {
                var task = m_scheduler[Calendars_LB.SelectedIndex];
                if ( (int) task.Event < Event_CB.Items.Count )
                    Event_CB.SelectedIndex = (int) task.Event;

                GoogleCal_CB.SelectedIndex = GoogleCal_CB.Items.IndexOf( task.Pair.GoogleName );
                OutlookCal_CB.SelectedIndex = OutlookCal_CB.Items.IndexOf( task.Pair.OutlookName );

                if ( task.TimeSpan > 0 && task.Event == SchedulerEvent.CustomTime )
                    Time_TB.Text = task.TimeSpan.ToString();

                SilentSync_CB.Checked = task.SilentSync;

                if ( task.SilentSync )
                    Precedence_CB.SelectedIndex = (int)task.Precedence;

            } else
            {
                Event_CB.SelectedText = "";
                GoogleCal_CB.SelectedText = "";
                OutlookCal_CB.SelectedText = "";
                Time_TB.Text = "";
                SilentSync_CB.Checked = false;
                Precedence_CB.SelectedText = "";
            }
        }

        private void UpdateTask_BTN_Click( object sender, EventArgs e )
        {
            try
            {
                if ( Calendars_LB.SelectedIndex >= 0 && Calendars_LB.SelectedIndex < m_scheduler.Count )
                {
                    SchedulerTask task = m_scheduler[Calendars_LB.SelectedIndex];

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
                        MessageBox.Show( this, "You need to select a calendar from both lists.", "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error );

                    if ( task.Event == SchedulerEvent.CustomTime )
                        task.TimeSpan = int.Parse( Time_TB.Text );

                    if ( SilentSync_CB.Checked )
                    {
                        task.SilentSync = true;
                        task.Precedence = (Precedence)Precedence_CB.SelectedIndex;
                    }

                    m_scheduler.UpdateTask( task, Calendars_LB.SelectedIndex );
                }
            } catch ( ArgumentOutOfRangeException ex )
            {
                Log.Write( ex );
                MessageBox.Show( this,
                    "There was an invalid index selected in the application. Unable to update the task.", ProductName,
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
            }

        }
    }
}
