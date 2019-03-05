using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Sharepoint_Mailing.model
{
    public class ExcelReader
    {
        String sheetName;
        protected String filePath, fileName;
        protected Excel.Application app;
        protected Excel.Workbook workbook;
        protected Excel._Worksheet worksheet;
        private int rowsTotal;
        private int emptyRowsTotal;
        private int allSheetsRowsTotal;

        Dictionary<String, int> errorList1, errorList2, errorList3;
        private int columnsTotal;

        public string SheetName { get => sheetName; set => sheetName = value; }
        public string FileName { get => fileName; set => fileName = value; }
        public Dictionary<string, int> ErrorList1 { get => errorList1; set => errorList1 = value; }
        public Dictionary<string, int> ErrorList2 { get => errorList2; set => errorList2 = value; }
        public Dictionary<string, int> ErrorList3 { get => errorList3; set => errorList3 = value; }
        public int EmptyRowsTotal { get => emptyRowsTotal; set => emptyRowsTotal = value; }
        public int ColumnsTotal { get => columnsTotal; set => columnsTotal = value; }
        public int RowsTotal { get => rowsTotal; set => rowsTotal = value; }
        public int AllSheetsRowsTotal { get => allSheetsRowsTotal; set => allSheetsRowsTotal = value; }

        public ExcelReader() { }

        public ExcelReader(String filePath)
        {
            this.filePath = filePath;
            FileName = filePath.Substring(filePath.LastIndexOf("\\")+1);
            app = new Excel.Application();
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Open(filePath);
            emptyRowsTotal = 0;
            allSheetsRowsTotal = 0;
            allSheetsRowsTotal += workbook.Sheets["SE16N_LOG"].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            allSheetsRowsTotal += workbook.Sheets["SM20"].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            allSheetsRowsTotal += workbook.Sheets["CDHDR_CDPOS"].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            allSheetsRowsTotal += workbook.Sheets["DBTABLOG"].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

        //otwiera arkusz o podanym tytule
        public void openSheet(String sheet)
        {
            SheetName = sheet;
            worksheet = workbook.Sheets[sheet];
            Excel.Range last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            ColumnsTotal = last.Column;
            RowsTotal = last.Row;
        }

        //zamyka plik excelowy
        public void close()
        {
            workbook.Close();
        }

        //wyszukuje wszystkie brakujące wartości, po czym wypełnia błędami 3 listy zawarte w readerze (errorList1, 2 i 3), odpowiednio dla kroków 1-3
        public UserList findMissingCells()
        {

            UserList users = new UserList();

            Excel.Range range = worksheet.Cells.Find("Incident_Number");
            int incidentColumn = range.Column;
            int row = range.Row;

            range = worksheet.Rows[row].Find("Comments");
            int commentsColumn = range.Column;

            range = worksheet.Rows[row].Find("User");
            int userColumn = range.Column;

            range = worksheet.Cells.Find("Approver");
            int approverColumn = range.Column;

            range = worksheet.Rows[row].Find("Comment")[2];
            int commentColumn = range.Column;

            range = worksheet.Cells.Find("Key_User_Approval");
            int keyUserColumn = range.Column;

            range = worksheet.Rows[row].Find("Approval_in_incident_Yes_No");
            int approvalColumn = range.Column;

            range = worksheet.Rows[row].Find("CHG_Date");
            int dateColumn = range.Column;

            for (int i = row + 1; i < RowsTotal; i++)
            {
                bool error = false;
                String column = "";

                //step1
                Excel.Range cell = worksheet.Cells[incidentColumn][i];
                if (cell.Value == null || cell.Value.ToString().Equals(""))
                {
                    error = true;
                    column = "Incident Number";
                }
                else
                {
                    cell = worksheet.Cells[commentsColumn][i];
                    if (cell.Value == null || cell.Value.ToString().Equals(""))
                    {
                        error = true;
                        column = "Comments";
                    }
                    else
                    {
                        //step2
                        cell = worksheet.Cells[approverColumn][i];
                        if (cell.Value == null || cell.Value.ToString().Equals(""))
                        {
                            error = true;
                            column = "Approver";
                        }
                        else
                        {
                            cell = worksheet.Cells[commentColumn][i];
                            if (cell.Value == null || cell.Value.ToString().Equals(""))
                            {
                                error = true;
                                column = "Comment";
                            }
                            else
                            {
                                //step3
                                cell = worksheet.Cells[keyUserColumn][i];
                                if (cell.Value == null || cell.Value.ToString().Equals(""))
                                {
                                    error = true;
                                    column = "Key User Approval/Comment";
                                }
                                else
                                {
                                    cell = worksheet.Cells[approvalColumn][i];
                                    if (cell.Value == null || cell.Value.ToString().Equals(""))
                                    {
                                        error = true;
                                        column = "Approval in incident (Yes/No)";
                                    }
                                }
                            }
                        }
                    }
                }

                //tworzenie listy
                if (error)
                {
                    cell = worksheet.Cells[userColumn][i];
                    String userName = cell.Value.ToString();
                    cell = worksheet.Cells[dateColumn][i];
                    String date = cell.Value.ToString();
                    users.add(userName, "Consultant");
                    users.addError(userName, fileName, sheetName, column, date);
                }
            }

            return users;
        }
    }
}
