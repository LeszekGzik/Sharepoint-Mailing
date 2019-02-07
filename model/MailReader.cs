using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Sharepoint_Mailing.model
{
    public class MailReader
    {
        protected String filePath, fileName;
        protected Excel.Application app;
        protected Excel.Workbook workbook;
        protected Excel._Worksheet worksheet;

        public string FileName { get => fileName; set => fileName = value; }

        public MailReader(String filePath)
        {
            this.filePath = filePath;
            FileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
            app = new Excel.Application();
            workbook = app.Workbooks.Open(filePath);
        }

        //zamyka plik excelowy
        public void close()
        {
            workbook.Close();
        }

        //przyjmuje listę <user, tekst wiadomości> i zwraca <adres, wiadomość>
        public Dictionary<String, String> getMailingList1(Dictionary<String, String> messageList)
        {
            Dictionary<String, String> mailingList = new Dictionary<string, string>();
            worksheet = workbook.Sheets[1];
            Excel.Range column1 = worksheet.Columns[1];

            foreach (String user in messageList.Keys)
            {
                int row = column1.Find(user).Row;
                String name = worksheet.Cells[2][row].Value.ToString();
                String address = worksheet.Cells[3][row].Value.ToString();
                mailingList.Add(
                    address,
                    "Dear " +
                    name +
                    ",\n\n" +
                    messageList[user] +
                    "\n" +
                    "Thank you for your cooperation.");
            }

            return mailingList;
        }

        //przyjmuje listę <user, tekst wiadomości> i zwraca <adres approvera, wiadomość do approvera>
        public Dictionary<String, String> getMailingList23(Dictionary<String, String> messageList)
        {
            Dictionary<String, String> mailingList = new Dictionary<string, string>();
            worksheet = workbook.Sheets[1];

            foreach (String user in messageList.Keys)
            {
                int row = worksheet.Columns[1].Find(user).Row;
                String name = worksheet.Cells[5][row].Value.ToString();
                String address = worksheet.Cells[6][row].Value.ToString();
                if(mailingList.Keys.Contains(address))
                {
                    mailingList[address] += ("\n" + messageList[user]);
                }
                else
                {
                    mailingList.Add(
                        address, 
                        "Dear " + 
                        name + 
                        ",\n\n" + 
                        messageList[user]);
                }
            }

            List<String> addresses = mailingList.Keys.ToList();


            foreach (String address in addresses)
            {
                mailingList[address] += "\nThank you for your cooperation.";
            }

            return mailingList;
        }
    }
}
