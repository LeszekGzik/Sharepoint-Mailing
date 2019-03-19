using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Sharepoint_Mailing.model
{
    //odczytuje listę mailingową (imiona i nazwiska, adresy, stream liderzy) z pliku Excelowego
    public class MailReader
    {
        protected String filePath, fileName;
        protected Excel.Application app;
        protected Excel.Workbook workbook;
        protected Excel._Worksheet worksheet;

        public string FileName { get => fileName; set => fileName = value; }

        //konstruktor przyjmujący ścieżkę do pliku jako parametr
        public MailReader(String filePath)
        {
            this.filePath = filePath;
            FileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
            app = new Excel.Application();
            workbook = app.Workbooks.Open(filePath);
            worksheet = workbook.Sheets[1];
        }

        //zamyka plik
        public void close()
        {
            workbook.Close();
        }

        //odnajduje podanego usera w pliku i zwraca jego pełne imię i nazwisko
        public String getFullName(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[2][row].Value.ToString();
        }

        //odnajduje podanego usera w pliku i zwraca jego adres
        public String getAddress(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[3][row].Value.ToString();
        }

        //odnajduje podanego usera w pliku i zwraca jego stream
        public String getStream(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[4][row].Value.ToString();
        }

        //odnajduje podanego usera w pliku i zwraca jego lidera
        public String getLeadName(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[5][row].Value.ToString();
        }

        //odnajduje podanego usera w pliku i zwraca adres jego lidera
        public String getLeadAddress(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[6][row].Value.ToString();
        }
    }
}
