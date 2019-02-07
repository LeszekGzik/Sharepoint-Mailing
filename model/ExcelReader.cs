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

            Excel.Range range = worksheet.Cells.Find("Incident Number");
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

            range = worksheet.Cells.Find("Key User Approval");
            int keyUserColumn = range.Column;

            range = worksheet.Rows[row].Find("Approval in incident");
            int approvalColumn = range.Column;

            for (int i = row + 1; i < RowsTotal; i++)
            {
                bool error1 = false;
                bool error2 = false;
                bool error3 = false;

                //step1
                Excel.Range cell = worksheet.Cells[incidentColumn][i];
                if (cell.Value == null || cell.Value.ToString().Equals(""))
                {
                    error1 = true;
                }
                else
                {
                    cell = worksheet.Cells[commentsColumn][i];
                    if (cell.Value == null || cell.Value.ToString().Equals(""))
                    {
                        error1 = true;
                    }
                    else
                    {
                        //step2
                        cell = worksheet.Cells[approverColumn][i];
                        if (cell.Value == null || cell.Value.ToString().Equals(""))
                        {
                            error2 = true;
                        }
                        else
                        {
                            cell = worksheet.Cells[commentColumn][i];
                            if (cell.Value == null || cell.Value.ToString().Equals(""))
                            {
                                error2 = true;
                            }
                            else
                            {
                                //step3
                                cell = worksheet.Cells[keyUserColumn][i];
                                if (cell.Value == null || cell.Value.ToString().Equals(""))
                                {
                                    error3 = true;
                                }
                                else
                                {
                                    cell = worksheet.Cells[approvalColumn][i];
                                    if (cell.Value == null || cell.Value.ToString().Equals(""))
                                    {
                                        error3 = true;
                                    }
                                }
                            }
                        }
                    }
                }

                //tworzenie listy
                if (error1)
                {
                    cell = worksheet.Cells[userColumn][i];
                    String userName = cell.Value.ToString();
                    users.add(userName, "Consultant");
                    users.addError(userName, fileName, sheetName);
                }
                else if (error2)
                {
                    cell = worksheet.Cells[userColumn][i];
                    String userName = "@" + cell.Value.ToString();
                    users.add(userName, "Approver");
                    users.addError(userName, fileName, sheetName);
                }
                else if (error3)
                {
                    cell = worksheet.Cells[userColumn][i];
                    String userName = "@" + cell.Value.ToString();
                    users.add(userName, "Approver");
                    users.addError(userName, fileName, sheetName);
                }
            }

            return users;
        }
    }
}
