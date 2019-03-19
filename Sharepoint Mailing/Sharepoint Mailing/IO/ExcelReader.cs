using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Sharepoint_Mailing.model
{
    //klasa służąca do czytania arkuszy excelowych i sprawdzania ich pod kątem niewypełnionych (błędnych) wierszy
    public class ExcelReader
    {
        String sheetName;
        protected String filePath, fileName;
        protected Excel.Application app;
        protected Excel.Workbook workbook;
        protected Excel._Worksheet worksheet;
        private int rowsTotal;

        public string SheetName { get => sheetName; set => sheetName = value; }
        public string FileName { get => fileName; set => fileName = value; }
        public int RowsTotal { get => rowsTotal; set => rowsTotal = value; }

        public ExcelReader() { }

        //konstruktor przyjmujący ścieżkę do pliku jako parametr
        public ExcelReader(String filePath)
        {
            this.filePath = filePath;
            FileName = filePath.Substring(filePath.LastIndexOf("\\")+1);    //pozyskaj nazwę pliku ze ścieżki
            app = new Excel.Application();                                  //otwórz Excela
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Open(filePath, false, true);           //otwórz plik jako read-only
            System.Threading.Thread.Sleep(1000);                    //sekundowy wait stanowi zabezpieczenie przed błędnym odczytem pliku
        }

        //otwiera arkusz o podanym tytule
        public void openSheet(String sheet)
        {
            SheetName = sheet;
            worksheet = workbook.Sheets[sheet];
            RowsTotal = worksheet.UsedRange.Rows.Count; //zapisz liczbę wierszy w arkuszu
        }

        //zamyka plik excelowy
        public void close()
        {
            workbook.Close();
        }

        //wyszukuje wszystkie brakujące wartości w aktualnej zakładce i tworzy na ich podstawie userListę
        public UserList findMissingCells()
        {
            UserList users = new UserList();

            //odnajdywanie indeksów kolumn z poszukiwanymi wartościami
            Excel.Range range = worksheet.Cells.Find("Incident_Number");
            if(range==null)
            {
                range = worksheet.Cells.Find("Incident Number");
            }
            int incidentColumn = range.Column;
            int row = range.Row;    //określenie wiersza z nagłówkami

            range = worksheet.Rows[row].Find("Comments");
            int commentsColumn = range.Column;

            range = worksheet.Rows[row].Find("User");
            int userColumn = range.Column;

            range = worksheet.Cells.Find("Approver");
            int approverColumn = range.Column;

            range = worksheet.Rows[row].Find("Comment")[2];
            int commentColumn = range.Column;

            range = worksheet.Cells.Find("Key_User_Approval");
            if (range == null)
            {
                range = worksheet.Cells.Find("Key User Approval/Comment");
            }
            int keyUserColumn = range.Column;

            range = worksheet.Rows[row].Find("Approval_in_incident_Yes_No");
            if (range == null)
            {
                range = worksheet.Cells.Find("Approval in incident (Yes/No)");
            }
            int approvalColumn = range.Column;

            range = worksheet.Rows[row].Find("CHG_Date");
            if (range == null)
            {
                range = worksheet.Cells.Find("Date");
                if (range == null)
                {
                    range = worksheet.Cells.Find("DBTABLOG-LOGDATE");
                }
            }
            int dateColumn = range.Column;

            //pętla dla każdego wiersza, poczynając od wiersza pod nagłówkami, aż do końca arkusza
            for (int i = row + 1; i < RowsTotal; i++)
            {
                bool error = false;
                String column = "";

                Excel.Range cell = worksheet.Cells[incidentColumn][i];
                if (cell.Value == null || cell.Value.ToString().Equals(".") || cell.Value.ToString().Equals(""))
                {
                    error = true;
                    column = "Incident Number";
                }
                else
                {
                    cell = worksheet.Cells[commentsColumn][i];
                    if (cell.Value == null || cell.Value.ToString().Equals(".") || cell.Value.ToString().Equals(""))
                    {
                        error = true;
                        column = "Comments";
                    }
                    else
                    {
                        cell = worksheet.Cells[approverColumn][i];
                        if (cell.Value == null || cell.Value.ToString().Equals(".") || cell.Value.ToString().Equals(""))
                        {
                            error = true;
                            column = "Approver";
                        }
                        else
                        {
                            cell = worksheet.Cells[commentColumn][i];
                            if (cell.Value == null || cell.Value.ToString().Equals(".") || cell.Value.ToString().Equals(""))
                            {
                                error = true;
                                column = "Comment";
                            }
                            else
                            {
                                cell = worksheet.Cells[keyUserColumn][i];
                                if (cell.Value == null || cell.Value.ToString().Equals(".") || cell.Value.ToString().Equals(""))
                                {
                                    error = true;
                                    column = "Key User Approval/Comment";
                                }
                                else
                                {
                                    cell = worksheet.Cells[approvalColumn][i];
                                    if (cell.Value == null || cell.Value.ToString().Equals(".") || cell.Value.ToString().Equals(""))
                                    {
                                        error = true;
                                        column = "Approval in incident (Yes/No)";
                                    }
                                }
                            }
                        }
                    }
                }

                //jeśli znalazłeś pustą komórkę:
                if (error)
                {
                    cell = worksheet.Cells[userColumn][i]; //odnajdź komórkę z nazwą użytkownika
                    if (cell.Value != null)
                    {
                        String userName = cell.Value.ToString();    //odczytaj nazwę użytkownika
                        cell = worksheet.Cells[dateColumn][i];      //odszukaj komórkę z datą
                        String date = cell.Value.ToString();        //odczytaj datę
                        users.add(userName, "Consultant");          //dodaj usera do userlisty
                        users.addError(userName, fileName, sheetName, column, date);    //dodaj błąd do usera
                    }
                }
            }

            return users;
        }
    }
}
