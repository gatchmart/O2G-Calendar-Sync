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
                OutlookSync.Syncer.Application.CreateItem( Outlook.OlItemType.olContactItem ) as Outlook.ContactItem;

            if ( contact == null )
                return false;

            contact.FirstName = Name;

            if ( Email != null )
                contact.Email1Address = Email;

            if ( Phone != null )
                contact.HomeTelephoneNumber = Phone;

            //contact.Display( false );
            contact.Save();

            return true;
        }
        
    }
}