﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Outlook_Calendar_Sync {
    public partial class CompareForm : Form {

        private List<CalendarItem> m_data;
        private SyncerForm m_parent;

        public CompareForm() {
            InitializeComponent();
        }

        public void SetParent( SyncerForm p ) {
            m_parent = p;
        }

        public void LoadData( List<CalendarItem> items ) {

            m_data = items;
            foreach ( var calendarItem in items ) {
                listView1.Items.Add(
                    new ListViewItem( new[] {calendarItem.Subject, GetActionString( calendarItem.Action )} ) );
            }

            listView1.AutoResizeColumns( ColumnHeaderAutoResizeStyle.ColumnContent );
            eventAction.Width = eventAction.Width < 260 ? 260 : eventAction.Width;
            eventSubject.Width = eventSubject.Width < 200 ? 200 : eventSubject.Width;
        }

        private string GetActionString( CalendarItemAction action ) {
            string output = "";

            if ( action.HasFlag( CalendarItemAction.GoogleAdd ) )
                output += "Add to Google | ";

            if ( action.HasFlag( CalendarItemAction.GoogleUpdate ) )
                output += "Update Google | ";

            if ( action.HasFlag( CalendarItemAction.GoogleDelete ) )
                output += "Delete from Google | ";

            if ( action.HasFlag( CalendarItemAction.OutlookAdd ) )
                output += "Add to Outlook | ";

            if ( action.HasFlag( CalendarItemAction.OutlookUpdate ) )
                output += "Update Outlook | ";

            if ( action.HasFlag( CalendarItemAction.OutlookDelete ) )
                output += "Delete from Outlook | ";

            return output.TrimEnd( new [] { ' ', '|'} );
        }

        private void submit_BTN_Click( object sender, EventArgs e ) {
            var cal = ( from int item in listView1.CheckedIndices select m_data[item] ).ToList();

            m_parent.StartUpdate( cal );
            Close();
        }

        private void cancel_BTN_Click( object sender, EventArgs e ) {
            Close();
        }

        private void checkAll_BTN_Click( object sender, EventArgs e ) {
            foreach ( ListViewItem item in listView1.Items )
                item.Checked = true;
        }

        private void uncheckAll_BTN_Click( object sender, EventArgs e ) {
            foreach ( ListViewItem item in listView1.Items )
                item.Checked = false;
        }

        private void CompareForm_Load( object sender, EventArgs e ) {
            
        }
    }
}
