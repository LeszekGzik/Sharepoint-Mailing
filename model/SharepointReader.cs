using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
//using Microsoft.SharePoint.Client;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Sharepoint_Mailing.model
{
    public class SharepointReader : ExcelReader
    {
        //public ClientContext clientContext { get; set; }

        public SharepointReader(String URL)
        {
            this.filePath = HttpUtility.HtmlEncode(URL);
            this.FileName = filePath.Substring(filePath.LastIndexOf("/") + 1);
            app = new Excel.Application();
            workbook = app.Workbooks.Open(filePath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        }
    }
}
