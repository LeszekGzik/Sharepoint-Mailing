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
        ExcelReader excelReader, mailReader;
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
            excelReader = new ExcelReader(textBoxFilePath.Text);
            errorString += runCheckOnTab("SE16N_LOG");
            errorString += runCheckOnTab("SM20");
            errorString += runCheckOnTab("CDHDR_CDPOS");
            errorString += runCheckOnTab("DBTABLOG");
            errorString += ("\n" + excelReader.EmptyRowsTotal + "/" + excelReader.AllSheetsRowsTotal + " rows missing in total.");

            if (checkBoxMail.Checked)
            {
                sendReport(errorString);
            }

            MessageBox.Show(errorString, "Errors found");
            excelReader.close();
        }

        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                textBoxFilePath.Text = openFileDialog1.FileName;
            }
        }

        private void buttonCheckAndRemind_Click(object sender, EventArgs e)
        {
            String errorString = "";
            messageList1 = new Dictionary<string, string>();
            messageList2 = new Dictionary<string, string>();
            messageList3 = new Dictionary<string, string>();
            excelReader = new ExcelReader(textBoxFilePath.Text);
            mailReader = new ExcelReader(textBoxEmailPath.Text);
            errorString += runCheckAndRemindOnTab("SE16N_LOG");
            errorString += runCheckAndRemindOnTab("SM20");
            errorString += runCheckAndRemindOnTab("CDHDR_CDPOS");
            errorString += runCheckAndRemindOnTab("DBTABLOG");
            errorString += ("\n" + excelReader.EmptyRowsTotal + "/" + excelReader.AllSheetsRowsTotal + " rows missing in total.");

            outlookMailer.sendToAll("Reminder for file " + excelReader.FileName, mailReader.getMailingList1(messageList1));
            outlookMailer.sendToAll("Reminder for file " + excelReader.FileName, mailReader.getMailingList23(messageList2));
            outlookMailer.sendToAll("Reminder for file " + excelReader.FileName, mailReader.getMailingList23(messageList3));

            if (checkBoxMail.Checked)
            {
                sendReport(errorString);
            }

            excelReader.close();
            mailReader.close();
        }

        //przeprowadza sprawdzenie na podanej zakładce; dopisuje liczbę znalezionych błędów do messageList
        public String runCheckAndRemindOnTab(String tab)
        {
            excelReader.openSheet(tab);
            excelReader.findMissingCells();
            String errorString = "";

            //step1
            Dictionary<string, int> errorList = excelReader.ErrorList1;
            foreach (String user in errorList.Keys)
            {
                errorString += ("User " + user + " has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + "\n");
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
                errorString += (user + "'s APPROVER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + "\n");
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
                errorString += (user + "'s KEY USER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + "\n");
                if (messageList3.Keys.Contains(user))
                {
                    messageList3[user] += ("A key user has " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
                else
                {
                    messageList3.Add(user, "A key user has " + errorList[user] + " rows left to fill in file " + excelReader.FileName + " in tab " + excelReader.SheetName + ".\n");
                }
            }

            return errorString;
        }

        public String runCheckOnTab(String tab)
        {
            excelReader.openSheet(tab);
            excelReader.findMissingCells();
            String errorString = "";
            Dictionary<string, int> errorList = excelReader.ErrorList1;

            foreach (String user in errorList.Keys)
            {
                errorString += ("User " + user + " has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + "\n");
            }

            errorList = excelReader.ErrorList2;
            foreach (String user in errorList.Keys)
            {
                errorString += (user + "'s APPROVER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + "\n");
            }

            errorList = excelReader.ErrorList3;
            foreach (String user in errorList.Keys)
            {
                errorString += (user + "'s KEY USER has " + errorList[user] + " rows to fill in tab " + excelReader.SheetName + "\n");
            }

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

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            saveConfig();
        }

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

        private void sendReport(String message)
        {
            outlookMailer.sendMail("Check results from file " + excelReader.FileName, textBoxControllerEmail.Text, message);
        }
    }
}