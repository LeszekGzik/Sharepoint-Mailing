﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    //klasa przechowująca dane jednego użytkownika i jego listę błędów
    public class User
    {
        String name;
        String address;
        String role;
        ErrorList errors;
        int totalRows;
        String fullName;
        String stream;
        String streamLeadName;
        String streamLeadAddress;


        public string Name { get => name; set => name = value; }
        public string Address { get => address; set => address = value; }
        public string Role { get => role; set => role = value; }
        public ErrorList Errors { get => errors; set => errors = value; }
        public int TotalRows { get => totalRows; set => totalRows = value; }
        public string FullName { get => fullName; set => fullName = value; }
        public string Stream { get => stream; set => stream = value; }
        public string StreamLeadName { get => streamLeadName; set => streamLeadName = value; }
        public string StreamLeadAddress { get => streamLeadAddress; set => streamLeadAddress = value; }

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

        //dodaje błąd do listy usera
        public void addError(String file, String tab, String column, String date)
        {
            Errors.addError(file, tab, column, date);
        }

        //dodaje błąd do listy usera
        public void addError(Error error)
        {
            Errors.addError(error);
        }

        //zwraca błąd z listy o podanym kluczu
        public Error getError(String key)
        {
            return Errors.get(key);
        }

        //zwraca pełną listę kluczy z listy błędów
        public List<String> getErrorKeys()
        {
            return Errors.Items.Keys.ToList();
        }

        //sumuje błędy z błędami innego użytkownika
        public void sumErrors(User anotherUser)
        {
            foreach(String key in anotherUser.getErrorKeys())
            {
                errors.addError(anotherUser.getError(key));
            }
        }

        //zwraca string zawierający spis wszystkich błędów dla danego usera
        internal string getErrorString()
        {
            String errorString = "";
            foreach(String key in getErrorKeys())
            {
                Error e = getError(key);
                errorString += ("User " + FullName + " has " + e.Count + " rows to fill in column " + e.Column + " in tab " + e.Tab + " in file " + e.File + ".\n");

            }
            return errorString;
        }
    }
}
