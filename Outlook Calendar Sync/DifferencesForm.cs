using System;

using System.Windows.Forms;

namespace Outlook_Calendar_Sync {
    public partial class DifferencesForm : Form {

        public DifferencesForm() {
            InitializeComponent();
        }

        public static DialogResult Show( CalendarItem outlook, CalendarItem google ) {
            using ( var form = new DifferencesForm() ) {
                form.LoadData( outlook, google );
                return form.ShowDialog();
            }
        }

        public void LoadData( CalendarItem outlook, CalendarItem google ) {

            //////////////////////////////////////////////////////////////////////////
            // Outlook Changes
            //////////////////////////////////////////////////////////////////////////

            if ( outlook.Changes.HasFlag( CalendarItemChanges.StartDate ) )
                Outlook_LV.Items.Add( new ListViewItem( new [] { "Start Date: " + DateTime.Parse( outlook.Start ).ToString( "R" ) } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.EndDate ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "End Date: " + DateTime.Parse( outlook.End ).ToString( "R" )  } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Location ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Location: " + outlook.Location } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Body ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Body: " + outlook.Body } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Subject ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Subject: " + outlook.Subject } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Start Time Zone: " + outlook.StartTimeZone } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "End Time Zone: " + outlook.EndTimeZone } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Reminder Time: " + outlook.ReminderTime } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Attendees ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Attendees: " + outlook.Attendees } ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                Outlook_LV.Items.Add( new ListViewItem( new[] { "Recurrence" } ) );

            //////////////////////////////////////////////////////////////////////////
            // Google Changes
            //////////////////////////////////////////////////////////////////////////

            if ( google.Changes.HasFlag( CalendarItemChanges.StartDate ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Start Date: " + DateTime.Parse( google.Start ).ToString( "R" ) } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.EndDate ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "End Date: " + DateTime.Parse( google.End ).ToString( "R" ) } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Location ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Location: " + google.Location } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Body ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Body: " + google.Body } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Subject ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Subject: " + google.Subject } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Start Time Zone: " + google.StartTimeZone } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "End Time Zone: " + google.EndTimeZone } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Reminder Time: " + google.ReminderTime } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Attendees ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Attendees: " + google.Attendees } ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                Google_LV.Items.Add( new ListViewItem( new[] { "Recurrence" } ) );


        }

        private void outlook_BTN_Click( object sender, EventArgs e ) {
            DialogResult = DialogResult.Yes;
            Dispose();
        }

        private void Ignore_BTN_Click( object sender, EventArgs e ) {
            DialogResult = DialogResult.Cancel;
            Dispose();
        }

        private void Google_BTN_Click( object sender, EventArgs e ) {
            DialogResult = DialogResult.No;
            Dispose();
        }
    }
}
