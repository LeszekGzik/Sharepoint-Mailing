using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Sharepoint_Mailing.model
{
    class OutlookMailer
    {
        Outlook.Application app;

        public OutlookMailer()
        {
            app = new Outlook.Application();
        }

        //rozsyła wszystkie maile podane w argumencie w postaci listy <adres, wiadomość>
        public void sendToAll(String subject, UserList userList)
        {
            Outlook.MailItem mailItem;
            foreach (String userName in userList.Items.Keys)
            {
                mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = subject;
                mailItem.To = userList.get(userName).Address;
                mailItem.Body = userList.getErrorString(userName);
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
                mailItem.Display(false);
                mailItem.Send();
            }
        }

        public void sendMail(String subject, String address, String message)
        {
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.To = address;
            mailItem.Body = message;
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
            mailItem.Display(false);
            mailItem.Send();
        }
    }
}
