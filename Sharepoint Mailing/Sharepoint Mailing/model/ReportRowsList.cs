using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    //lista wierszy do wpisania do raportu
    public class ReportRowsList
    {
        public Dictionary<String,ReportRow> Items;

        //konstruktor tworzący listę na podstawie userlisty
        public ReportRowsList(UserList userList)
        {
            Items = new Dictionary<String,ReportRow>();
            foreach(String name in userList.getKeys())      //pętla po wszystkich userach
            {
                User user = userList.get(name);
                foreach(String key in user.getErrorKeys())  //pętla po wszystkich błędach dla danego usera
                {
                    Error err = user.getError(key);
                    String shortKey = name + ";" + key.Substring(0, key.LastIndexOf(';')); //generowanie unikatowego klucza w postaci "username;plik;zakładka"
                    if (Items.Keys.Contains(shortKey))
                    {
                        Items[shortKey].addToColumn(err.Count, err.Column); //jeśli wiersz o takim kluczu już istnieje, dodaj do niego więcej błędów
                    }
                    else
                    {
                        Items.Add(shortKey, new ReportRow(user, err.Tab, err.File, err.Date)); //jeśli nie, utwórz nowy wiersz
                        Items[shortKey].addToColumn(err.Count, err.Column);
                    }
                }
            }
        }

        //zwraca wiersz o podanym kluczu
        public ReportRow get(String key)
        {
            return Items[key];
        }

        //zwraca listę kluczy występujących na liście
        public List<String> getKeys()
        {
            return Items.Keys.ToList();
        }
    }
}
