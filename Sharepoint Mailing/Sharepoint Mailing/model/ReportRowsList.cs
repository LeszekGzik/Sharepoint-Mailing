using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    public class ReportRowsList
    {
        public Dictionary<String,ReportRow> Items;

        public ReportRowsList(UserList userList)
        {
            Items = new Dictionary<String,ReportRow>();
            foreach(String name in userList.getKeys())
            {
                User user = userList.get(name);
                foreach(String key in user.getErrorKeys())
                {
                    Error err = user.getError(key);
                    String shortKey = name + ";" + key.Substring(0, key.LastIndexOf(';'));
                    if (Items.Keys.Contains(shortKey))
                    {
                        Items[shortKey].addToColumn(err.Count, err.Column);
                    }
                    else
                    {
                        Items.Add(shortKey, new ReportRow(user, err.Tab, err.File, err.Date));
                        Items[shortKey].addToColumn(err.Count, err.Column);
                    }
                }
            }
        }

        public ReportRow get(String key)
        {
            return Items[key];
        }

        public List<String> getKeys()
        {
            return Items.Keys.ToList();
        }
    }
}
