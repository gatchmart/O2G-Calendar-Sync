using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_Calendar_Sync {
    public partial class ThisAddIn
    {

        private Syncer m_syncer;
        private Scheduler.Scheduler m_scheduler;
        private List<Outlook.Items> m_items;
        private List<Outlook.Folder> m_folders;

        private void ThisAddIn_Startup( object sender, System.EventArgs e ) {
            OutlookSync.Syncer.Init( Application );
            ( (Outlook.ApplicationEvents_11_Event) Application ).Quit += ThisAddIn_Quit;

            m_syncer = Syncer.Instance;
            m_scheduler = Scheduler.Scheduler.Instance;

            m_items = new List<Outlook.Items>();
            m_folders = new List<Outlook.Folder>();

            foreach ( Outlook.Folder folder in Application.Session.Folders )
                GetFolders( folder );
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
            m_scheduler.AbortThread();
            m_scheduler.Save( false );
            Archiver.Instance.Save();

            foreach ( var mapiFolder in m_folders )
                Marshal.ReleaseComObject( mapiFolder );

            foreach ( var item in m_items )
                Marshal.ReleaseComObject( item );

            m_folders.Clear();
            m_items.Clear();
        }

        private void ThisAddIn_Shutdown( object sender, System.EventArgs e ) {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void GetFolders( Outlook.Folder folder )
        {
            foreach ( Outlook.Folder child in folder.Folders )
            {
                if ( child.DefaultItemType == Outlook.OlItemType.olAppointmentItem )
                {
                    var items = child.Items;
                    m_folders.Add( child );
                    m_items.Add( items );
                    items.ItemChange += Outlook_ItemChange;
                    items.ItemAdd += Outlook_ItemAdd;
                    items.ItemRemove += Outlook_ItemRemove;

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
