using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    //klasa przechowująca listę wielu użytkowników
    public class UserList
    {
        Dictionary<String,User> items;

        public Dictionary<String, User> Items { get => items; set => items = value; }

        public UserList()
        {
            Items = new Dictionary<String, User>();
        }

        public List<String> getKeys()
        {
            return Items.Keys.ToList();
        }

        //dodaje użytkownika o podanej nazwie do listy
        public void add(String name)
        {
            if (!Items.Keys.Contains(name)) {
                User user = new User();
                user.Name = name;
                Items.Add(name, user);
            }
        }

        //dodaje użytkownika o podanej roli i nazwie do listy
        public void add(String name, String role)
        {
            if (!Items.Keys.Contains(name))
            {
                User user = new User();
                user.Name = name;
                user.Role = role;
                Items.Add(name, user);
            }
        }

        //dodaje użytkownika do listy
        public void add(User user)
        {
            Items.Add(user.Name, user);
        }

        //zwraca usera o podanej nazwie
        public User get(String name)
        {
            return Items[name];
        }

        //dodaje błąd do użytkownika o podanej nazwie
        public void addError(String name, String file, String tab, String column, String date)
        {
            Items[name].addError(file, tab, column, date);
        }

        //scala tą userlistę z inną userlistą
        public UserList sum(UserList anotherList)
        {
            foreach(String anotherUserName in anotherList.Items.Keys) //pętla dla każdego usera w drugiej liście
            {
                if (Items.Keys.Contains(anotherUserName))
                {
                    get(anotherUserName).sumErrors(anotherList.get(anotherUserName)); //jeśli ta lista zawiera już tego usera, dodaj do niego nowe błędy
                }
                else
                {
                    Items.Add(anotherUserName, anotherList.get(anotherUserName));     //jeśli nie, dodaj usera do listy
                }
            }
            return this;
        }

        //zwraca string z błędami wszystkich użytkowników
        public String getErrorString()
        {
            String errorString = "";

            foreach(String userName in Items.Keys)
            {
                errorString += get(userName).getErrorString();
            }

            return errorString;
        }

        //zwraca string z błędami użytkownika o danej nazwie
        public String getErrorString(String userName)
        {
            return get(userName).getErrorString();
        }

        //pobiera z mailreadera pełne imiona i nazwiska wszystkich userów
        public void getFullNames(MailReader reader)
        {
            foreach(String userName in Items.Keys)
            {
                get(userName).FullName = reader.getFullName(userName);
            }
        }

        //pobiera z mailreadera adresy, streamy i dane liderów dla wszystkich userów
        public void getAddresses(MailReader reader)
        {
            foreach (String userName in Items.Keys)
            {
                get(userName).Address = reader.getAddress(userName);
                get(userName).Stream = reader.getStream(userName);
                get(userName).StreamLeadName = reader.getLeadName(userName);
                get(userName).StreamLeadAddress = reader.getLeadAddress(userName);
            }
        }

        //zwraca true jeśli userlista zawiera użytkownika o danej nazwie
        public Boolean contains(String fullName)
        {
            foreach(String userName in Items.Keys)
            {
                if (get(userName).FullName.Equals(fullName))
                {
                    return true;
                }
            }
            return false;
        }

        public String keyOf(String fullName)
        {
            foreach (String userName in Items.Keys)
            {
                if (get(userName).FullName.Equals(fullName))
                {
                    return userName;
                }
            }
            return "null";
        }

        //scala userów którzy mają ten sam Full Name
        public UserList mergeExcessUsers()
        {
            UserList newList = new UserList();

            foreach(String userName in Items.Keys)
            {
                String fullName = get(userName).FullName;
                if (newList.contains(fullName)) {
                    String newUserName = newList.keyOf(fullName);
                    foreach(String error in Items[userName].getErrorKeys())
                    {
                        newList.get(newUserName).addError(Items[userName].getError(error));
                    }
                }
                else
                {
                    newList.add(get(userName));
                }
            }

            return newList;
        }
    }
}
