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
        int count;

        public string File { get => file; set => file = value; }
        public string Tab { get => tab; set => tab = value; }
        public int Count { get => count; set => count = value; }

        public Error(String _file, String _tab)
        {
            File = _file;
            Tab = _tab;
            Count = 1;
        }

        public Error(String _file, String _tab, int _count)
        {
            File = _file;
            Tab = _tab;
            Count = count;
        }

        public void increment(int howMany)
        {
            Count += howMany;
        }
    }
}
