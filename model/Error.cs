using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    public class Error
    {
        String file;
        String tab;
        String column;
        int count;
        String date;

        public string File { get => file; set => file = value; }
        public string Tab { get => tab; set => tab = value; }
        public int Count { get => count; set => count = value; }
        public string Column { get => column; set => column = value; }
        public string Date { get => date; set => date = value; }

        public Error(String _file, String _tab, String _column, String _date)
        {
            File = _file;
            Tab = _tab;
            Column = _column;
            Date = _date;
            Count = 1;
        }

        public Error(String _file, String _tab, String _column, String _date, int _count)
        {
            File = _file;
            Tab = _tab;
            Column = _column;
            Date = _date;
            Count = count;
        }

        public void increment(int howMany)
        {
            Count += howMany;
        }
    }
}
