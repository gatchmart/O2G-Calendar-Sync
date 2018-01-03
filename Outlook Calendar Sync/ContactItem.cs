using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_Calendar_Sync {
    public class ContactItem {

        public string Name { get; set; }

        public string Email { get; set; }

        public string Phone { get; set; }

        public ContactItem(string name, string email = null, string phone = null ) {
            Name = name;
            Email = email;
            Phone = phone;
        }

        public bool CreateContact() {
            Outlook.ContactItem contact =
                OutlookSync.Syncer.CurrentApplication.CreateItem( Outlook.OlItemType.olContactItem ) as Outlook.ContactItem;

            if ( contact == null )
                return false;

            contact.FirstName = Name;

            if ( Email != null )
                contact.Email1Address = Email;

            if ( Phone != null )
                contact.HomeTelephoneNumber = Phone;

            contact.Save();

            return true;
        }

        public static ContactItem GetContactItem( string name )
        {
            var strs = name.Split( '(' );

            name = strs[0].Trim();
            Outlook.MAPIFolder contactFolder =
                OutlookSync.Syncer.CurrentApplication.Session.GetDefaultFolder( Outlook.OlDefaultFolders
                    .olFolderContacts );
            Outlook.Items contactItems = contactFolder.Items;

            Outlook.ContactItem contact = (Outlook.ContactItem) contactItems.Find( $"[FirstName]={name}" );

            if ( contact == null )
                return null;

            ContactItem item = new ContactItem( contact.FirstName, contact.Email1Address, contact.PrimaryTelephoneNumber );
            return item;
        }
        
    }
}