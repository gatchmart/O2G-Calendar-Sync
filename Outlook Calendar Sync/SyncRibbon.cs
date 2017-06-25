using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Outlook_Calendar_Sync.Properties;

namespace Outlook_Calendar_Sync {
    public partial class SyncRibbon {
        private SyncerForm m_syncerForm;

        private void SyncRibbon_Load( object sender, RibbonUIEventArgs e ) {
            m_syncerForm = new SyncerForm();
            m_syncerForm.Ribbon = this;

            if ( Settings.Default.IsInitialLoad ) {
                Syncer.Instance.PerformInitalLoad();

                Settings.Default.IsInitialLoad = false;
                Settings.Default.Save();
            }
                
        }

        private void Sync_BTN_Click( object sender, RibbonControlEventArgs e ) {
            m_syncerForm.Show();
        }
    }
}
