using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook_Calendar_Sync.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_Calendar_Sync {
    public partial class ThisAddIn
    {

        private Syncer m_syncer;
        private Scheduler.Scheduler m_scheduler;
        private List<Outlook.Items> m_items = new List<Outlook.Items>();

        private void ThisAddIn_Startup( object sender, System.EventArgs e ) {
            OutlookSync.Syncer.Init( Application );
            ( (Outlook.ApplicationEvents_11_Event) Application ).Quit += ThisAddIn_Quit;

            m_syncer = Syncer.Instance;
            m_scheduler = Scheduler.Scheduler.Instance;

            foreach ( Outlook.MAPIFolder folder in Application.Session.Folders )
                GetFolders( folder );

            //var fo = Application.Session.Folders;
            //foreach ( Outlook.Folder f in fo )
            //{
            //    foreach ( Outlook.Folder f2 in f.Folders )
            //    {
            //        if ( f2.FolderPath.Contains( "Calendar" ) )
            //        {
            //            Log.Write( f2.FolderPath );
            //            var items = f2.Items;
            //            m_items.Add( items );
            //            items.ItemChange += Outlook_ItemChange;
            //            items.ItemAdd += Outlook_ItemAdd;
            //            items.ItemRemove += Outlook_ItemRemove;

            //            foreach ( Outlook.Folder f2Folder in f2.Folders )
            //            {
            //                Log.Write( f2Folder.FolderPath );
            //                var items2 = f2Folder.Items;
            //                m_items.Add( items2 );
            //                items2.ItemChange += Outlook_ItemChange;
            //                items2.ItemAdd += Outlook_ItemAdd;
            //                items2.ItemRemove += Outlook_ItemRemove;
            //            }
            //        }
            //    }
            //}

        }

        private void Outlook_ItemAdd( object item )
        {
            m_scheduler.Item_Add( item );
        }

        private void Outlook_ItemChange( object item )
        {
            m_scheduler.Item_Change( item );
        }

        private void Outlook_ItemRemove()
        {
            m_scheduler.Item_Remove();
        }

        private void ThisAddIn_Quit()
        {
            m_scheduler.AboutThread();
            m_scheduler.Save( false );
            Archiver.Instance.Save();
        }

        private void ThisAddIn_Shutdown( object sender, System.EventArgs e ) {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void GetFolders( Outlook.MAPIFolder folder )
        {
            foreach ( Outlook.MAPIFolder child in folder.Folders )
            {
                if ( child.DefaultItemType == Outlook.OlItemType.olAppointmentItem )
                {
                    m_items.Add( child.Items );
                    child.Items.ItemChange += Outlook_ItemChange;
                    child.Items.ItemAdd += Outlook_ItemAdd;
                    child.Items.ItemRemove += Outlook_ItemRemove;

                    Log.Write( $"Added EventHandlers for the {child.Name} folder." );

                    if ( child.Folders.Count != 0 )
                        GetFolders( child );
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler( ThisAddIn_Startup );
            this.Shutdown += new System.EventHandler( ThisAddIn_Shutdown );
        }

        #endregion
    }
}
