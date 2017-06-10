using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace Outlook_Calendar_Sync {
    public partial class SyncRibbon {

        private void SyncRibbon_Load( object sender, RibbonUIEventArgs e ) {
            
        }

        private void Sync_BTN_Click( object sender, RibbonControlEventArgs e ) {
            SyncerForm form = new SyncerForm();
            form.Show();
        }
    }
}
