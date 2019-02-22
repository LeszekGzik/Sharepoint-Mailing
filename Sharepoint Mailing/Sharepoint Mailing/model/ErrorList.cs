using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    public class ErrorList
    {
        Dictionary<String, Error> items;

        public Dictionary<string, Error> Items { get => items; set => items = value; }

        public ErrorList()
        {
            Items = new Dictionary<string, Error>();
        }

        public Error get(String key)
        {
            return Items[key];
        }

        public void addError(String file, String tab, String column, String date)
        {
            String key = file + ";" + tab + ";" + column;
            if(Items.Keys.Contains(key))
            {
                Items[key].increment(1);
            }
            else
            {
                Items.Add(key, new Error(file, tab, column, date));
            }
        }

        public void addErrors(String file, String tab, String column, String date, int count)
        {
            String key = file + ";" + tab + ";" + column;
            if (Items.Keys.Contains(key))
            {
                Items[key].increment(count);
            }
            else
            {
                Items.Add(key, new Error(file, tab, column, date, count));
            }
        }

        public void addError(Error _err)
        {
            addErrors(_err.File, _err.Tab, _err.Column, _err.Date, _err.Count);
        }
    }
}
