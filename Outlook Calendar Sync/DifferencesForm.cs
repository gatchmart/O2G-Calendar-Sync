using System;
using System.Windows.Forms;
using Outlook_Calendar_Sync.Enums;

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

        public void LoadData( CalendarItem outlook, CalendarItem google )
        {

            OutlookSubject_LBL.Text = outlook.Subject;
            GoogleSubject_LBL.Text = google.Subject;

            //////////////////////////////////////////////////////////////////////////
            // Outlook Changes
            //////////////////////////////////////////////////////////////////////////

            if ( outlook.Changes.HasFlag( CalendarItemChanges.CalId ) )
                AppendOutlook_RTBText( $"iCal ID needs to be updated:\n\t{outlook.iCalID}" );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.StartDate ) )
                AppendOutlook_RTBText( "Start Date:\n\t" + DateTime.Parse( outlook.Start ).ToString( "R" ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.EndDate ) )
                AppendOutlook_RTBText( "End Date:\n\t" + DateTime.Parse( outlook.End ).ToString( "R" ) );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Location ) )
                AppendOutlook_RTBText( "Location:\n\t" + outlook.Location );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Body ) )
                AppendOutlook_RTBText( "Body:\n\t" + outlook.Body );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Subject ) )
                AppendOutlook_RTBText( "Subject:\n\t" + outlook.Subject );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                AppendOutlook_RTBText( "Start Time Zone:\n\t" + outlook.StartTimeZone );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                AppendOutlook_RTBText( "End Time Zone:\n\t" + outlook.EndTimeZone );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                AppendOutlook_RTBText( "Reminder Time:\n\t" + outlook.ReminderTime );

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Attendees ) )
            {
                AppendOutlook_RTBText( "Attendees:" );
                foreach ( var outlookAttendee in outlook.Attendees )
                    AppendOutlook_RTBText( "\n\t" + outlookAttendee );
            }

            if ( outlook.Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                AppendOutlook_RTBText( "Recurrence:\n" + outlook.Recurrence );

            if ( outlook.Changes == CalendarItemChanges.Nothing )
                throw new Exception("Why is the differences form being shown if there are no differences.\nOutlook:\n" + outlook + "\nGoogle\n" + google );

            //////////////////////////////////////////////////////////////////////////
            // Google Changes
            //////////////////////////////////////////////////////////////////////////

            if ( google.Changes.HasFlag( CalendarItemChanges.CalId ) )
                AppendGoogle_RTBText( $"iCal ID needs to be updated.\n\t{google.iCalID}" );

            if ( google.Changes.HasFlag( CalendarItemChanges.StartDate ) )
                AppendGoogle_RTBText( "Start Date:\n\t" + DateTime.Parse( google.Start ).ToString( "R" ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.EndDate ) )
                AppendGoogle_RTBText( "End Date:\n\t" + DateTime.Parse( google.End ).ToString( "R" ) );

            if ( google.Changes.HasFlag( CalendarItemChanges.Location ) )
                AppendGoogle_RTBText( "Location:\n\t" + google.Location );

            if ( google.Changes.HasFlag( CalendarItemChanges.Body ) )
                AppendGoogle_RTBText( "Body:\n\t" + google.Body );

            if ( google.Changes.HasFlag( CalendarItemChanges.Subject ) )
                AppendGoogle_RTBText( "Subject:\n\t" + google.Subject );

            if ( google.Changes.HasFlag( CalendarItemChanges.StartTimeZone ) )
                AppendGoogle_RTBText( "Start Time Zone:\n\t" + google.StartTimeZone );

            if ( google.Changes.HasFlag( CalendarItemChanges.EndTimeZone ) )
                AppendGoogle_RTBText( "End Time Zone:\n\t" + google.EndTimeZone);

            if ( google.Changes.HasFlag( CalendarItemChanges.ReminderTime ) )
                AppendGoogle_RTBText( "Reminder Time:\n\t" + google.ReminderTime );

            if ( google.Changes.HasFlag( CalendarItemChanges.Attendees ) )
            {
                AppendGoogle_RTBText( "Attendees:" );
                foreach ( var googleAttendee in google.Attendees )
                    AppendGoogle_RTBText("\n\t" + googleAttendee );
            }

            if ( google.Changes.HasFlag( CalendarItemChanges.Recurrence ) )
                AppendGoogle_RTBText( "Recurrence\n" + google.Recurrence );

            if ( google.Changes == CalendarItemChanges.Nothing )
                throw new Exception( "Why is the differences form being shown if there are no differences.\nOutlook:\n" + outlook + "\nGoogle\n" + google );

        }

        private void outlook_BTN_Click( object sender, EventArgs e )
        {
            var result = All_CB.Checked
                ? DifferencesFormResults.KeepOutlookAll
                : DifferencesFormResults.KeepOutlookSingle;
            DialogResult = (DialogResult) result;
            Dispose();
        }

        private void Ignore_BTN_Click( object sender, EventArgs e )
        {
            var result = All_CB.Checked 
                ? DifferencesFormResults.IgnoreAll 
                : DifferencesFormResults.IgnoreSingle;
            DialogResult = (DialogResult) result;
            Dispose();
        }

        private void Google_BTN_Click( object sender, EventArgs e )
        {
            var result = All_CB.Checked
                ? DifferencesFormResults.KeepGoogleAll
                : DifferencesFormResults.KeepGoogleSingle;
            DialogResult = (DialogResult) result;
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
