using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    public class User
    {
        String name;
        String address;
        String role;
        ErrorList errors;
        int totalRows;
        String fullName;

        public string Name { get => name; set => name = value; }
        public string Address { get => address; set => address = value; }
        public string Role { get => role; set => role = value; }
        public ErrorList Errors { get => errors; set => errors = value; }
        public int TotalRows { get => totalRows; set => totalRows = value; }
        public string FullName { get => fullName; set => fullName = value; }

        public User()
        {
            Errors = new ErrorList();
            TotalRows = 0;
        }

        public User(String _name, String _address, String _role)
        {
            Name = _name;
            Address = _address;
            Role = _role;
            TotalRows = 0;
        }

        public void addError(String file, String tab)
        {
            Errors.addError(file, tab);
        }

        public void addError(Error error)
        {
            Errors.addError(error);
        }

        public Error getError(String key)
        {
            return Errors.get(key);
        }

        public List<String> getErrorKeys()
        {
            return Errors.Items.Keys.ToList();
        }

        public void sumErrors(User anotherUser)
        {
            foreach(String key in anotherUser.getErrorKeys())
            {
                errors.addError(anotherUser.getError(key));
            }
        }

        internal string getErrorString()
        {
            String errorString = "";
            foreach(String key in getErrorKeys())
            {
                Error e = getError(key);
                errorString += ("User " + FullName + " has " + e.Count + " rows to fill in tab " + e.Tab + " in file " + e.File + ".\n");
            }
            return errorString;
        }
    }
}
