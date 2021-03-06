﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Sharepoint_Mailing.model
{
    //klasa do wysyłania e-maili za pośrednictwem Outlooka
    class OutlookMailer
    {
        Outlook.Application app;

        public OutlookMailer()
        {
            app = new Outlook.Application();
        }

        //wysyła mail z podanym tematem, treścią i załącznikami do wszystkich użytkowników w userLiście
        public void sendToAll(String subject, UserList userList, String message, params String[] attachments)
        {
            Outlook.MailItem mailItem;
            foreach (String userName in userList.Items.Keys)
            {
                mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = subject;
                mailItem.To = userList.get(userName).Address;
                mailItem.Body = message;
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
                foreach (String filePath in attachments)
                {
                    mailItem.Attachments.Add(filePath, Outlook.OlAttachmentType.olByValue, 1, "Report.xlsx");
                }
                mailItem.Display(false);
                mailItem.Send();
            }
        }

        //wysyła mail z podanym tematem i treścią pod wybrany adres
        //(outdated)
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

        //wysyła mail z podanym tematem, treścią i załącznikami pod wybrany adres
        public void sendMail(String subject, String address, String message, params String[] attachments)
        {
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.To = address;
            mailItem.Body = message;
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
            foreach(String filePath in attachments)
            {
                mailItem.Attachments.Add(filePath, Outlook.OlAttachmentType.olByValue, 1, "Report.xlsx");
            }
            mailItem.Display(false);
            mailItem.Send();
        }
    }
}
