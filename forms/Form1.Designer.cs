namespace Sharepoint_Mailing
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.textBoxFilePath = new System.Windows.Forms.TextBox();
            this.buttonOpenFile = new System.Windows.Forms.Button();
            this.labelFilePath = new System.Windows.Forms.Label();
            this.buttonRunCheck = new System.Windows.Forms.Button();
            this.buttonRunCheckAndRemind = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.labelEmailPath = new System.Windows.Forms.Label();
            this.buttonOpenFileEmail = new System.Windows.Forms.Button();
            this.textBoxEmailPath = new System.Windows.Forms.TextBox();
            this.labelControllerEmail = new System.Windows.Forms.Label();
            this.textBoxControllerEmail = new System.Windows.Forms.TextBox();
            this.checkBoxMail = new System.Windows.Forms.CheckBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ColumnFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnCheckBox = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.checkBoxAll = new System.Windows.Forms.CheckBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabelFiles = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusLabelTabs = new System.Windows.Forms.ToolStripStatusLabel();
            this.buttonDebug = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxFilePath
            // 
            this.textBoxFilePath.Enabled = false;
            this.textBoxFilePath.Location = new System.Drawing.Point(47, 51);
            this.textBoxFilePath.Name = "textBoxFilePath";
            this.textBoxFilePath.Size = new System.Drawing.Size(376, 20);
            this.textBoxFilePath.TabIndex = 0;
            // 
            // buttonOpenFile
            // 
            this.buttonOpenFile.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenFile.Image")));
            this.buttonOpenFile.Location = new System.Drawing.Point(429, 49);
            this.buttonOpenFile.Name = "buttonOpenFile";
            this.buttonOpenFile.Size = new System.Drawing.Size(28, 23);
            this.buttonOpenFile.TabIndex = 1;
            this.buttonOpenFile.UseVisualStyleBackColor = true;
            this.buttonOpenFile.Click += new System.EventHandler(this.buttonOpenFile_Click);
            // 
            // labelFilePath
            // 
            this.labelFilePath.AutoSize = true;
            this.labelFilePath.Location = new System.Drawing.Point(47, 32);
            this.labelFilePath.Name = "labelFilePath";
            this.labelFilePath.Size = new System.Drawing.Size(60, 13);
            this.labelFilePath.TabIndex = 2;
            this.labelFilePath.Text = "Folder path";
            // 
            // buttonRunCheck
            // 
            this.buttonRunCheck.Location = new System.Drawing.Point(121, 468);
            this.buttonRunCheck.Name = "buttonRunCheck";
            this.buttonRunCheck.Size = new System.Drawing.Size(75, 23);
            this.buttonRunCheck.TabIndex = 4;
            this.buttonRunCheck.Text = "Run Check";
            this.buttonRunCheck.UseVisualStyleBackColor = true;
            this.buttonRunCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // buttonRunCheckAndRemind
            // 
            this.buttonRunCheckAndRemind.Location = new System.Drawing.Point(261, 468);
            this.buttonRunCheckAndRemind.Name = "buttonRunCheckAndRemind";
            this.buttonRunCheckAndRemind.Size = new System.Drawing.Size(121, 23);
            this.buttonRunCheckAndRemind.TabIndex = 4;
            this.buttonRunCheckAndRemind.Text = "Run Check + Remind";
            this.buttonRunCheckAndRemind.UseVisualStyleBackColor = true;
            this.buttonRunCheckAndRemind.Click += new System.EventHandler(this.buttonCheckAndRemind_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // labelEmailPath
            // 
            this.labelEmailPath.AutoSize = true;
            this.labelEmailPath.Location = new System.Drawing.Point(49, 330);
            this.labelEmailPath.Name = "labelEmailPath";
            this.labelEmailPath.Size = new System.Drawing.Size(87, 13);
            this.labelEmailPath.TabIndex = 7;
            this.labelEmailPath.Text = "Email list file path";
            // 
            // buttonOpenFileEmail
            // 
            this.buttonOpenFileEmail.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenFileEmail.Image")));
            this.buttonOpenFileEmail.Location = new System.Drawing.Point(432, 348);
            this.buttonOpenFileEmail.Name = "buttonOpenFileEmail";
            this.buttonOpenFileEmail.Size = new System.Drawing.Size(28, 23);
            this.buttonOpenFileEmail.TabIndex = 6;
            this.buttonOpenFileEmail.UseVisualStyleBackColor = true;
            this.buttonOpenFileEmail.Click += new System.EventHandler(this.buttonOpenFileEmail_Click);
            // 
            // textBoxEmailPath
            // 
            this.textBoxEmailPath.Location = new System.Drawing.Point(50, 350);
            this.textBoxEmailPath.Name = "textBoxEmailPath";
            this.textBoxEmailPath.Size = new System.Drawing.Size(376, 20);
            this.textBoxEmailPath.TabIndex = 5;
            // 
            // labelControllerEmail
            // 
            this.labelControllerEmail.AutoSize = true;
            this.labelControllerEmail.Location = new System.Drawing.Point(49, 379);
            this.labelControllerEmail.Name = "labelControllerEmail";
            this.labelControllerEmail.Size = new System.Drawing.Size(79, 13);
            this.labelControllerEmail.TabIndex = 9;
            this.labelControllerEmail.Text = "Report address";
            // 
            // textBoxControllerEmail
            // 
            this.textBoxControllerEmail.Enabled = false;
            this.textBoxControllerEmail.Location = new System.Drawing.Point(49, 395);
            this.textBoxControllerEmail.Name = "textBoxControllerEmail";
            this.textBoxControllerEmail.Size = new System.Drawing.Size(410, 20);
            this.textBoxControllerEmail.TabIndex = 8;
            // 
            // checkBoxMail
            // 
            this.checkBoxMail.AutoSize = true;
            this.checkBoxMail.Location = new System.Drawing.Point(49, 421);
            this.checkBoxMail.Name = "checkBoxMail";
            this.checkBoxMail.Size = new System.Drawing.Size(188, 17);
            this.checkBoxMail.TabIndex = 10;
            this.checkBoxMail.Text = "Mail me a report when you\'re done";
            this.checkBoxMail.UseVisualStyleBackColor = true;
            this.checkBoxMail.CheckedChanged += new System.EventHandler(this.checkBoxMail_CheckedChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnFileName,
            this.ColumnCheckBox});
            this.dataGridView1.Location = new System.Drawing.Point(50, 78);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(410, 197);
            this.dataGridView1.TabIndex = 11;
            // 
            // ColumnFileName
            // 
            this.ColumnFileName.HeaderText = "File name";
            this.ColumnFileName.Name = "ColumnFileName";
            this.ColumnFileName.ReadOnly = true;
            this.ColumnFileName.Width = 300;
            // 
            // ColumnCheckBox
            // 
            this.ColumnCheckBox.FalseValue = "false";
            this.ColumnCheckBox.HeaderText = "Check";
            this.ColumnCheckBox.Name = "ColumnCheckBox";
            this.ColumnCheckBox.TrueValue = "true";
            this.ColumnCheckBox.Width = 50;
            // 
            // checkBoxAll
            // 
            this.checkBoxAll.AutoSize = true;
            this.checkBoxAll.Location = new System.Drawing.Point(389, 281);
            this.checkBoxAll.Name = "checkBoxAll";
            this.checkBoxAll.Size = new System.Drawing.Size(70, 17);
            this.checkBoxAll.TabIndex = 13;
            this.checkBoxAll.Text = "Check all";
            this.checkBoxAll.UseVisualStyleBackColor = true;
            this.checkBoxAll.CheckedChanged += new System.EventHandler(this.checkBoxAll_CheckedChanged);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabelFiles,
            this.statusLabelTabs});
            this.statusStrip1.Location = new System.Drawing.Point(0, 528);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(477, 22);
            this.statusStrip1.TabIndex = 14;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusLabelFiles
            // 
            this.statusLabelFiles.Name = "statusLabelFiles";
            this.statusLabelFiles.Size = new System.Drawing.Size(83, 17);
            this.statusLabelFiles.Text = "Files done: 0/0";
            // 
            // statusLabelTabs
            // 
            this.statusLabelTabs.Name = "statusLabelTabs";
            this.statusLabelTabs.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.statusLabelTabs.Size = new System.Drawing.Size(85, 17);
            this.statusLabelTabs.Text = "Tabs done: 0/0";
            // 
            // buttonDebug
            // 
            this.buttonDebug.Location = new System.Drawing.Point(12, 290);
            this.buttonDebug.Name = "buttonDebug";
            this.buttonDebug.Size = new System.Drawing.Size(75, 23);
            this.buttonDebug.TabIndex = 15;
            this.buttonDebug.Text = "DEBUG";
            this.buttonDebug.UseVisualStyleBackColor = true;
            this.buttonDebug.Click += new System.EventHandler(this.buttonDebug_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 550);
            this.Controls.Add(this.buttonDebug);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.checkBoxAll);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.checkBoxMail);
            this.Controls.Add(this.labelControllerEmail);
            this.Controls.Add(this.textBoxControllerEmail);
            this.Controls.Add(this.labelEmailPath);
            this.Controls.Add(this.buttonOpenFileEmail);
            this.Controls.Add(this.textBoxEmailPath);
            this.Controls.Add(this.buttonRunCheckAndRemind);
            this.Controls.Add(this.buttonRunCheck);
            this.Controls.Add(this.labelFilePath);
            this.Controls.Add(this.buttonOpenFile);
            this.Controls.Add(this.textBoxFilePath);
            this.Name = "Form1";
            this.Text = "Sharepoint Mailing";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFilePath;
        private System.Windows.Forms.Button buttonOpenFile;
        private System.Windows.Forms.Label labelFilePath;
        private System.Windows.Forms.Button buttonRunCheck;
        private System.Windows.Forms.Button buttonRunCheckAndRemind;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label labelEmailPath;
        private System.Windows.Forms.Button buttonOpenFileEmail;
        private System.Windows.Forms.TextBox textBoxEmailPath;
        private System.Windows.Forms.Label labelControllerEmail;
        private System.Windows.Forms.TextBox textBoxControllerEmail;
        private System.Windows.Forms.CheckBox checkBoxMail;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.CheckBox checkBoxAll;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnFileName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnCheckBox;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabelFiles;
        private System.Windows.Forms.ToolStripStatusLabel statusLabelTabs;
        private System.Windows.Forms.Button buttonDebug;
    }
}

