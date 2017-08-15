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

            if ( outlook.Changes.HasFlag( CalendarItemChanges.CalId ) )
                AppendOutlook_RTBText( "iCal ID needs to be updated." );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.StartDate ) )
                AppendOutlook_RTBText( "Start Date: " + DateTime.Parse( outlook.Start ).ToString( "R" ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.EndDate ) )
                AppendOutlook_RTBText( "End Date: " + DateTime.Parse( outlook.End ).ToString( "R" ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Location ) )
                AppendOutlook_RTBText( "Location: " + outlook.Location );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Body ) )
                AppendOutlook_RTBText( "Body: " + outlook.Body );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Subject ) )
                AppendOutlook_RTBText( "Subject: " + outlook.Subject );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                AppendOutlook_RTBText( "Start Time Zone: " + outlook.StartTimeZone );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                AppendOutlook_RTBText( "End Time Zone: " + outlook.EndTimeZone );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                AppendOutlook_RTBText( "Reminder Time: " + outlook.ReminderTime );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Attendees ) )
                AppendOutlook_RTBText( "Attendees: " + outlook.Attendees );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                AppendOutlook_RTBText( "Recurrence:\n" + outlook.Recurrence );

            if ( outlook.Changes == CalendarItemChanges.Nothing )
                throw new Exception("Why is the differences form being shown if there are no differences.\nOutlook:\n" + outlook + "\nGoogle\n" + google );

            //////////////////////////////////////////////////////////////////////////
            // Google Changes
            //////////////////////////////////////////////////////////////////////////

            if ( google.Changes.HasFlag( CalendarItemChanges.CalId ) )
                AppendGoogle_RTBText( "iCal ID needs to be updated." );

            if ( google.Changes.HasFlag( CalendarItemChanges.StartDate ) )
                AppendGoogle_RTBText( "Start Date: " + DateTime.Parse( google.Start ).ToString( "R" ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.EndDate ) )
                AppendGoogle_RTBText( "End Date: " + DateTime.Parse( google.End ).ToString( "R" ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Location ) )
                AppendGoogle_RTBText( "Location: " + google.Location );

            if ( google.Changes.HasFlag( CalendarItemChanges.Body ) )
                AppendGoogle_RTBText( "Body: " + google.Body );

            if ( google.Changes.HasFlag( CalendarItemChanges.Subject ) )
                AppendGoogle_RTBText( "Subject: " + google.Subject );

            if ( google.Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                AppendGoogle_RTBText( "Start Time Zone: " + google.StartTimeZone );

            if ( google.Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                AppendGoogle_RTBText( "End Time Zone: " + google.EndTimeZone);

            if ( google.Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                AppendGoogle_RTBText( "Reminder Time: " + google.ReminderTime );

            if ( google.Changes.HasFlag( CalendarItemChanges.Attendees ) )
                AppendGoogle_RTBText( "Attendees: " + google.Attendees );

            if ( google.Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                AppendGoogle_RTBText( "Recurrence\n" + google.Recurrence );

            if ( google.Changes == CalendarItemChanges.Nothing )
                throw new Exception( "Why is the differences form being shown if there are no differences.\nOutlook:\n" + outlook + "\nGoogle\n" + google );

        }

        private void outlook_BTN_Click( object sender, EventArgs e ) {
            DialogResult = All_CB.Checked ? DialogResult.OK : DialogResult.Yes;
            Dispose();
        }

        private void Ignore_BTN_Click( object sender, EventArgs e ) {
            DialogResult = All_CB.Checked ? DialogResult.Ignore : DialogResult.Cancel;
            Dispose();
        }

        private void Google_BTN_Click( object sender, EventArgs e ) {
            DialogResult = All_CB.Checked ? DialogResult.None : DialogResult.No;
            Dispose();
        }

        private void AppendOutlook_RTBText( string text ) {
            Outlook_RTB.AppendText( text + "\n" );
        }

        private void AppendGoogle_RTBText( string text ) {
            Google_RTB.AppendText( text + "\n" );
        }
    }
}
