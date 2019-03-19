using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_Mailing.model
{
    //klasa przechowująca pojedynczy wiersz do wpisania do raportu
    public class ReportRow
    {
        int incidentNumber;
        int comments;
        int approver;
        int comment;
        int keyUserApproval;
        int approvalInIncident;
        int totalErrors;

        String name;
        String address;
        String role;
        int totalRows;
        String fullName;
        String stream;
        String streamLeadName;
        String streamLeadAddress;

        String fileName;
        String fileTab;
        String date;

        public string Name { get => name; set => name = value; }
        public string Address { get => address; set => address = value; }
        public string Role { get => role; set => role = value; }
        public int TotalRows { get => totalRows; set => totalRows = value; }
        public string FullName { get => fullName; set => fullName = value; }
        public string Stream { get => stream; set => stream = value; }
        public string StreamLeadName { get => streamLeadName; set => streamLeadName = value; }
        public string StreamLeadAddress { get => streamLeadAddress; set => streamLeadAddress = value; }
        public string FileName { get => fileName; set => fileName = value; }
        public string FileTab { get => fileTab; set => fileTab = value; }
        public string Date { get => date; set => date = value; }
        public int IncidentNumber { get => incidentNumber; set => incidentNumber = value; }
        public int Comments { get => comments; set => comments = value; }
        public int Approver { get => approver; set => approver = value; }
        public int Comment { get => comment; set => comment = value; }
        public int KeyUserApproval { get => keyUserApproval; set => keyUserApproval = value; }
        public int ApprovalInIncident { get => approvalInIncident; set => approvalInIncident = value; }
        public int TotalErrors { get => totalErrors; set => totalErrors = value; }

        public ReportRow(User user, String _tab, String _file, String _date)
        {
            Name = user.Name;
            Address = user.Address;
            Role = user.Role;
            TotalRows = user.TotalRows;
            FullName = user.FullName;
            Stream = user.Stream;
            StreamLeadName = user.StreamLeadName;
            StreamLeadAddress = user.StreamLeadAddress;

            FileName = _file;
            FileTab = _tab;
            Date = _date;

            IncidentNumber = 0;
            Comments = 0;
            Approver = 0;
            Comment = 0;
            KeyUserApproval = 0;
            ApprovalInIncident = 0;
            TotalErrors = 0;
        }

        //dodaje NUMBER do liczby błędów w danej kolumnie
        public void addToColumn(int number, String column)
        {
            switch(column)
            {
                case "Incident Number":
                    IncidentNumber += number;
                    break;
                case "Comments":
                    Comments += number;
                    break;
                case "Approver":
                    Approver += number;
                    break;
                case "Comment":
                    Comment += number;
                    break;
                case "Key User Approval/Comment":
                    KeyUserApproval += number;
                    break;
                case "Approval in incident (Yes/No)":
                    ApprovalInIncident += number;
                    break;
            }
            TotalErrors += number;
        }
    }
}
