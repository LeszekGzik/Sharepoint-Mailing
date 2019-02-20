﻿using Sharepoint_Mailing.model;
using System;
using System.Collections.Generic;
using System.IO;
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
        int row;

        public int Row { get => row; set => row = value; }

        public ExcelWriter(String filePath)
        {
            this.fileName = filePath;
            app = new Excel.Application();
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Add();
            worksheet = workbook.Worksheets.Item[1];
        }

        public void importHeaders(String template)
        {
            Excel.Workbook tempBook = app.Workbooks.Open(template);
            Excel.Worksheet tempSheet = tempBook.Worksheets[1];
            tempSheet.Copy(workbook.Worksheets[1]);
            worksheet = workbook.Worksheets.Item[1];
            tempBook.Close();
            Row = 3;
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
            Row = 3;
        }

        internal void delete()
        {
            app.Quit();
            File.Delete(fileName);
        }

        public void writeReport(ReportRowsList list)
        {
            foreach(String key in list.getKeys())
            {
                ReportRow rr = list.get(key);
                worksheet.Cells[Row, 1] = rr.Date;
                worksheet.Cells[Row, 2] = rr.FileName;
                worksheet.Cells[Row, 3] = rr.Name;
                worksheet.Cells[Row, 4] = rr.FullName;
                worksheet.Cells[Row, 5] = rr.Address;
                worksheet.Cells[Row, 6] = rr.FileTab;
                if (rr.IncidentNumber != 0) { worksheet.Cells[Row, 7] = rr.IncidentNumber; }
                if (rr.Comments != 0) { worksheet.Cells[Row, 8] = rr.Comments; }
                if (rr.Approver != 0) { worksheet.Cells[Row, 9] = rr.Approver; }
                if (rr.Comment != 0) { worksheet.Cells[Row, 10] = rr.Comment; }
                if (rr.KeyUserApproval != 0) { worksheet.Cells[Row, 11] = rr.KeyUserApproval; }
                if (rr.ApprovalInIncident != 0) { worksheet.Cells[Row, 12] = rr.ApprovalInIncident; }
                worksheet.Cells[Row, 13] = rr.Stream;
                worksheet.Cells[Row, 14] = rr.StreamLeadName;
                worksheet.Cells[Row, 15] = rr.StreamLeadAddress;
                Row++;
            }
        }

        public void writeErrors(User user)
        {
            foreach(String key in user.getErrorKeys())
            {
                Error err = user.getError(key);
                worksheet.Cells[Row, 1] = err.Date;
                worksheet.Cells[Row, 2] = err.File;
                worksheet.Cells[Row, 3] = user.Name;
                worksheet.Cells[Row, 4] = user.FullName;
                worksheet.Cells[Row, 5] = user.Address;
                worksheet.Cells[Row, 6] = err.Tab;
                switch(err.Column)
                {
                    case "Incident Number":
                        worksheet.Cells[Row, 7] = err.Count;
                        break;
                    case "Comments":
                        worksheet.Cells[Row, 8] = err.Count;
                        break;
                    case "Approver":
                        worksheet.Cells[Row, 9] = err.Count;
                        break;
                    case "Comment":
                        worksheet.Cells[Row, 10] = err.Count;
                        break;
                    case "Key User Approval/Comment":
                        worksheet.Cells[Row, 11] = err.Count;
                        break;
                    case "Approval in incident (Yes/No)":
                        worksheet.Cells[Row, 12] = err.Count;
                        break;
                }
                worksheet.Cells[Row, 13] = user.Stream;
                worksheet.Cells[Row, 14] = user.StreamLeadName;
                worksheet.Cells[Row, 15] = user.StreamLeadAddress;
                Row++;
            }
        }

        public void writeErrors(UserList users)
        {
            foreach(String userName in users.getKeys())
            {
                writeErrors(users.get(userName));
            }
        }

        public void save()
        {
            Excel.Range range = worksheet.Range[worksheet.Cells[4,1], worksheet.Cells[Row-1,15]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThin;
            workbook.SaveAs(fileName);
            workbook.Close();
        }
    }
}
