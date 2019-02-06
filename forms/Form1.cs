using Sharepoint_Mailing.model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Sharepoint_Mailing
{
    public partial class Form1 : Form
    {
        Dictionary<String, String> messageList1, messageList2, messageList3;
        ExcelReader excelReader;
        MailReader mailReader;
        OutlookMailer outlookMailer;

        public Form1()
        {
            InitializeComponent();
            outlookMailer = new OutlookMailer();
            loadConfig();
        }

        private void buttonCheck_Click(object sender, EventArgs e)
        {
            String errorString = "";
            foreach(DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
                if (chk.Value == chk.TrueValue)
                {
                    errorString += runCheckOnFile(textBoxFilePath.Text + "\\" + row.Cells[0].Value.ToString());
                }
            }

            if (checkBoxMail.Checked)
            {
                sendReport(errorString);
            }

            MessageBox.Show(errorString, "Errors found");
        }

        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog()==DialogResult.OK)
            {
                textBoxFilePath.Text = folderBrowserDialog1.SelectedPath;
                showFiles();
            }
        }

        //wypełnia dataGridView plikami w wybranym folderze
        private void showFiles()
        {
            DirectoryInfo dir = new DirectoryInfo(textBoxFilePath.Text);
            FileInfo[] files = dir.GetFiles("*.xls*");
            dataGridView1.Rows.Clear();
            foreach (FileInfo file in files)
            {
                dataGridView1.Rows.Add(file.Name);
            }
        }
        
        private void buttonCheckAndRemind_Click(object sender, EventArgs e)
        {
            String errorString = "";
            messageList1 = new Dictionary<string, string>();
            messageList2 = new Dictionary<string, string>();
            messageList3 = new Dictionary<string, string>();
            mailReader = new MailReader(textBoxEmailPath.Text);
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
                if (chk.Value == chk.TrueValue)
                {
                    errorString += runCheckAndRemindOnFile(textBoxFilePath.Text + "\\" + row.Cells[0].Value.ToString());
                }
            }
            outlookMailer.sendToAll("ZSOX Sharepoint Reminder", mailReader.getMailingList1(messageList1));
            outlookMailer.sendToAll("ZSOX Sharepoint Reminder", mailReader.getMailingList23(messageList2));
            outlookMailer.sendToAll("ZSOX Sharepoint Reminder", mailReader.getMailingList23(messageList3));

            if (checkBoxMail.Checked)
            {
                sendReport(errorString);
            }

            mailReader.close();
        }

        //przeprowadza sprawdzenie wszystkich zakładek w pliku i wypełnia messageLists powiadomieniami; zwraca listę wszystkich błędów + podsumowanie w postaci stringa
        private string runCheckAndRemindOnFile(string filePath)
        {
            Console.WriteLine("File check starting: " + filePath + "...");

            String errorString = "";
            excelReader = new ExcelReader(filePath);
            errorString += runCheckAndRemindOnTab("SE16N_LOG");
            errorString += runCheckAndRemindOnTab("SM20");
            errorString += runCheckAndRemindOnTab("CDHDR_CDPOS");
            errorString += runCheckAndRemindOnTab("DBTABLOG");
            errorString += (excelReader.EmptyRowsTotal + "/" + excelReader.AllSheetsRowsTotal + " rows missing in total.\n\n");
            excelReader.close();

            Console.WriteLine("File check finished: " + filePath);

            return errorString;
        }

        //przeprowadza sprawdzenie wszystkich zakładek w pliku; zwraca listę wszystkich błędów + podsumowanie w postaci stringa
        public String runCheckOnFile(String filePath)
        {
            Console.WriteLine("File check starting: " + filePath + "...");

            String errorString = "";
            excelReader = new ExcelReader(filePath);
            errorString += runCheckOnTab("SE16N_LOG");
            errorString += runCheckOnTab("SM20");
            errorString += runCheckOnTab("CDHDR_CDPOS");
            errorString += runCheckOnTab("DBTABLOG");
            errorString += (excelReader.EmptyRowsTotal + "/" + excelReader.AllSheetsRowsTotal + " rows missing in total.\n\n");
            excelReader.close();

            Console.WriteLine("File check finished: " + filePath);

            return errorString;
        }

        //przeprowadza sprawdzenie na podanej zakładce; dopisuje liczbę znalezionych błędów do odpowiednich messageLists; oraz zwraca listę wszystkich błędów w postaci stringa
        public String runCheckAndRemindOnTab(String tab)
        {
            Console.WriteLine("   Tab check starting: " + tab + "...");

            String errorString = "";
            excelReader.openSheet(tab);
            excelReader.findMissingCells();

            //step1
            Dictionary<string, int> errorList = excelReader.ErrorList1;
            foreach (String user in errorList.Keys)
            {
                errorString += ("User " + user + " has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + " in file " + excelReader.FileName + "\n");
                if (messageList1.Keys.Contains(user))
                {
                    messageList1[user] += ("You have " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
                else
                {
                    messageList1.Add(user, "You have " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
            }

            //step2
            errorList = excelReader.ErrorList2;
            foreach (String user in errorList.Keys)
            {
                errorString += (user + "'s APPROVER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + " in file " + excelReader.FileName + "\n");
                if (messageList2.Keys.Contains(user))
                {
                    messageList2[user] += ("You have " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
                else
                {
                    messageList2.Add(user, "You have " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
            }

            //step3
            errorList = excelReader.ErrorList3;
            foreach (String user in errorList.Keys)
            {
                errorString += (user + "'s KEY USER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + " in file " + excelReader.FileName + "\n");
                if (messageList3.Keys.Contains(user))
                {
                    messageList3[user] += ("A key user has " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
                else
                {
                    messageList3.Add(user, "A key user has " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
            }

            Console.WriteLine("   Tab check finished: " + tab);

            return errorString;
        }

        //wykonuje pełne sprawdzenie jednej zakładki w obecnym pliku w ExcelReaderze; zwraca listę wszystkich błędów w postaci stringa
        public String runCheckOnTab(String tab)
        {
            Console.WriteLine("   Tab check starting: " + tab + "...");

            excelReader.openSheet(tab);
            excelReader.findMissingCells();
            String errorString = "";
            Dictionary<string, int> errorList = excelReader.ErrorList1;

            foreach (String user in errorList.Keys)
            {
                errorString += ("User " + user + " has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + " in file " + excelReader.FileName + "\n");
            }

            errorList = excelReader.ErrorList2;
            foreach (String user in errorList.Keys)
            {
                errorString += (user + "'s APPROVER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + " in file " + excelReader.FileName + "\n");
            }

            errorList = excelReader.ErrorList3;
            foreach (String user in errorList.Keys)
            {
                errorString += (user + "'s KEY USER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + " in file " + excelReader.FileName + "\n");
            }

            Console.WriteLine("   Tab check finished: " + tab);

            return errorString;
        }

        private void buttonOpenFileEmail_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxEmailPath.Text = openFileDialog1.FileName;
            }
        }

        private void checkBoxMail_CheckedChanged(object sender, EventArgs e)
        {
            textBoxControllerEmail.Enabled = checkBoxMail.Checked;
        }

        //wczytuje ostatnią konfigurację z config.xml, jeśli takowy istnieje
        private void loadConfig()
        {
            if (File.Exists("config.xml"))
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("config.xml");
                XmlNodeList nodelist = doc.SelectNodes("//config/*");
                foreach (XmlElement node in nodelist)
                {
                    if (node.Name.Equals("mailingPath"))
                    {
                        textBoxEmailPath.Text = node.InnerText;
                    }
                    else if (node.Name.Equals("recentPath"))
                    {
                        textBoxFilePath.Text = node.InnerText;
                        showFiles();
                    }
                    else if (node.Name.Equals("mailMe"))
                    {
                        if(node.InnerText.Equals("True"))
                        {
                            checkBoxMail.Checked = true;
                        }
                    }
                    else if (node.Name.Equals("mailTo"))
                    {
                        textBoxControllerEmail.Text = node.InnerText;
                    }
                }
            }
        }

        private void checkBoxAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)row.Cells[1];
                cell.Value = checkBoxAll.Checked;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            saveConfig();
        }

        //zapisuje ostatnią konfigurację do config.xml
        private void saveConfig()
        {
            XmlDocument doc = new XmlDocument();
            XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.AppendChild(docNode);

            XmlElement configNode = doc.CreateElement("config");
            doc.AppendChild(configNode);

            XmlElement childNode = doc.CreateElement("mailingPath");
            childNode.InnerText = textBoxEmailPath.Text;
            configNode.AppendChild(childNode);

            childNode = doc.CreateElement("recentPath");
            childNode.InnerText = textBoxFilePath.Text;
            configNode.AppendChild(childNode);

            childNode = doc.CreateElement("mailMe");
            childNode.InnerText = checkBoxMail.Checked.ToString();
            configNode.AppendChild(childNode);

            childNode = doc.CreateElement("mailTo");
            childNode.InnerText = textBoxControllerEmail.Text;
            configNode.AppendChild(childNode);

            doc.Save("config.xml");
        }

        //wysyła zbiorczy, kompletny raport na adres podany w textboksie
        private void sendReport(String message)
        {
            outlookMailer.sendMail("ZSOX Sharepoint check results from day " + DateTime.Now.ToShortDateString(), textBoxControllerEmail.Text, message);
        }
    }
}