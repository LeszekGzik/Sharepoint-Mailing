using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Sharepoint_Mailing.IO
{
    public class ExcelWriter
    {
        public String fileName;
        protected Excel.Application app;
        protected Excel.Workbook workbook;
        protected Excel._Worksheet worksheet;

        public ExcelWriter(String fileName)
        {
            this.fileName = fileName;
            app = new Excel.Application();
            workbook = app.Workbooks.Add();
            worksheet = workbook.Worksheets.Item[1];
        }

        public void writeHeaders()
        {
            //first row
            worksheet.Cells[1, 7] = "Missing lines of information";
            worksheet.Cells[1, 13] = "Approver";
            worksheet.Range[worksheet.Cells[1, 7], worksheet.Cells[1, 12]].Merge();
            worksheet.Range[worksheet.Cells[1, 13], worksheet.Cells[1, 15]].Merge();

            //second row
            worksheet.Cells[2, 1] = "Date";
            worksheet.Cells[2, 2] = "File name";
            worksheet.Cells[2, 3] = "User";
            worksheet.Cells[2, 4] = "User name";
            worksheet.Cells[2, 5] = "E-Mail Address";
            worksheet.Cells[2, 6] = "File tab";
            worksheet.Cells[2, 7] = "Incident Number";
            worksheet.Cells[2, 8] = "Comments";
            worksheet.Cells[2, 9] = "Approver";
            worksheet.Cells[2, 10] = "Comment";
            worksheet.Cells[2, 11] = "Key User Approval/Comment";
            worksheet.Cells[2, 12] = "Approval in incident (Yes/No)";
            worksheet.Cells[2, 13] = "Stream";
            worksheet.Cells[2, 14] = "Stream Lead Name";
            worksheet.Cells[2, 15] = "Stream Lead E-Mail Address";
        }

        public void save()
        {
            workbook.SaveAs("./" + fileName);
            workbook.Close();
            app.Quit();
        }
    }
}
