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
            worksheet = workbook.Sheets[1];
        }

        //zamyka plik excelowy
        public void close()
        {
            workbook.Close();
        }

        public String getFullName(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[2][row].Value.ToString();
        }

        public String getAddress(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[3][row].Value.ToString();
        }

        public String getStream(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[4][row].Value.ToString();
        }

        public String getLeadName(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[5][row].Value.ToString();
        }

        public String getLeadAddress(String userName)
        {
            Excel.Range userColumn = worksheet.Columns[1];
            int row = userColumn.Find(userName).Row;
            return worksheet.Cells[6][row].Value.ToString();
        }
    }
}
