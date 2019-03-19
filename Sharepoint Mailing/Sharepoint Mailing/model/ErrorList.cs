using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    //lista przechowująca wiele błędów
    public class ErrorList
    {
        Dictionary<String, Error> items;

        public Dictionary<string, Error> Items { get => items; set => items = value; }

        public ErrorList()
        {
            Items = new Dictionary<string, Error>();
        }

        //zwraca błąd o podanym kluczu
        public Error get(String key)
        {
            return Items[key];
        }

        //dodaje nowy błąd do listy (jeśli już jest na liście to inkrementuje liczbę jego wystąpień)
        public void addError(String file, String tab, String column, String date)
        {
            //generowanie unikatowego klucza błędu w postaci "plik;zakładka;kolumna"
            String key = file + ";" + tab + ";" + column;
            if(Items.Keys.Contains(key))
            {
                Items[key].increment(1);    //jeśli już istnieje, zwiększ wystąpienia o 1
            }
            else
            {
                Items.Add(key, new Error(file, tab, column, date)); //jeśli nie istnieje, dodaj
            }
        }

        //dodaje nowy błąd do listy (jeśli już jest na liście to inkrementuje liczbę jego wystąpień o COUNT)
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

        //dodaje istniejący błąd do listy
        public void addError(Error _err)
        {
            addErrors(_err.File, _err.Tab, _err.Column, _err.Date, _err.Count);
        }
    }
}
