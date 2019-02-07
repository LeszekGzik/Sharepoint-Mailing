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
            mailReader = new MailReader(textBoxEmailPath.Text);
            UserList userList = new UserList();
            foreach(DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
                if (chk.Value == chk.TrueValue)
                {
                    userList.sum(runCheckOnFile(textBoxFilePath.Text + "\\" + row.Cells[0].Value.ToString()));
                }
            }
            userList.getFullNames(mailReader);
            userList = userList.mergeExcessUsers();
            String errorString = userList.getErrorString();

            if (checkBoxMail.Checked)
            {
                sendReport(errorString);
            }

            MessageBox.Show(errorString);
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
            mailReader = new MailReader(textBoxEmailPath.Text);
            UserList userList = new UserList();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
                if (chk.Value == chk.TrueValue)
                {
                    userList.sum(runCheckOnFile(textBoxFilePath.Text + "\\" + row.Cells[0].Value.ToString()));
                }
            }
            userList.getFullNames(mailReader);
            userList = userList.mergeExcessUsers();
            userList.getAddresses(mailReader);

            String errorString = userList.getErrorString();

            if (checkBoxMail.Checked)
            {
                sendReport(errorString);
            }

            outlookMailer.sendToAll("ZSOX Sharepoint Reminder", userList);
        }

        //przeprowadza sprawdzenie wszystkich zakładek w pliku; zwraca listę wszystkich błędów + podsumowanie w postaci stringa
        public UserList runCheckOnFile(String filePath)
        {
            Console.WriteLine("File check starting: " + filePath + "...");

            UserList users = new UserList();
            excelReader = new ExcelReader(filePath);
            users = users.sum(runCheckOnTab("SE16N_LOG"));
            users = users.sum(runCheckOnTab("SM20"));
            users = users.sum(runCheckOnTab("CDHDR_CDPOS"));
            users = users.sum(runCheckOnTab("DBTABLOG"));
            excelReader.close();

            Console.WriteLine("File check finished: " + filePath);

            return users;
        }

        //wykonuje pełne sprawdzenie jednej zakładki w obecnym pliku w ExcelReaderze; zwraca listę wszystkich błędów w postaci stringa
        public UserList runCheckOnTab(String tab)
        {
            Console.WriteLine("   Tab check starting: " + tab + "...");
            excelReader.openSheet(tab);
            UserList userList = excelReader.findMissingCells();
            Console.WriteLine("   Tab check finished: " + tab);

            return userList;
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
                        if (!node.InnerText.Equals(""))
                        {
                            textBoxFilePath.Text = node.InnerText;
                            showFiles();
                        }
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