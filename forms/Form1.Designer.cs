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
            this.buttonCheck = new System.Windows.Forms.Button();
            this.buttonCheckAndRemind = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.labelEmailPath = new System.Windows.Forms.Label();
            this.buttonOpenFileEmail = new System.Windows.Forms.Button();
            this.textBoxEmailPath = new System.Windows.Forms.TextBox();
            this.labelControllerEmail = new System.Windows.Forms.Label();
            this.textBoxControllerEmail = new System.Windows.Forms.TextBox();
            this.checkBoxMail = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // textBoxFilePath
            // 
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
            this.labelFilePath.Size = new System.Drawing.Size(47, 13);
            this.labelFilePath.TabIndex = 2;
            this.labelFilePath.Text = "File path";
            // 
            // buttonCheck
            // 
            this.buttonCheck.Location = new System.Drawing.Point(118, 244);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new System.Drawing.Size(75, 23);
            this.buttonCheck.TabIndex = 4;
            this.buttonCheck.Text = "Check";
            this.buttonCheck.UseVisualStyleBackColor = true;
            this.buttonCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // buttonCheckAndRemind
            // 
            this.buttonCheckAndRemind.Location = new System.Drawing.Point(271, 244);
            this.buttonCheckAndRemind.Name = "buttonCheckAndRemind";
            this.buttonCheckAndRemind.Size = new System.Drawing.Size(108, 23);
            this.buttonCheckAndRemind.TabIndex = 4;
            this.buttonCheckAndRemind.Text = "Check and Remind";
            this.buttonCheckAndRemind.UseVisualStyleBackColor = true;
            this.buttonCheckAndRemind.Click += new System.EventHandler(this.buttonCheckAndRemind_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // labelEmailPath
            // 
            this.labelEmailPath.AutoSize = true;
            this.labelEmailPath.Location = new System.Drawing.Point(47, 97);
            this.labelEmailPath.Name = "labelEmailPath";
            this.labelEmailPath.Size = new System.Drawing.Size(87, 13);
            this.labelEmailPath.TabIndex = 7;
            this.labelEmailPath.Text = "Email list file path";
            // 
            // buttonOpenFileEmail
            // 
            this.buttonOpenFileEmail.Image = ((System.Drawing.Image)(resources.GetObject("buttonOpenFileEmail.Image")));
            this.buttonOpenFileEmail.Location = new System.Drawing.Point(429, 114);
            this.buttonOpenFileEmail.Name = "buttonOpenFileEmail";
            this.buttonOpenFileEmail.Size = new System.Drawing.Size(28, 23);
            this.buttonOpenFileEmail.TabIndex = 6;
            this.buttonOpenFileEmail.UseVisualStyleBackColor = true;
            this.buttonOpenFileEmail.Click += new System.EventHandler(this.buttonOpenFileEmail_Click);
            // 
            // textBoxEmailPath
            // 
            this.textBoxEmailPath.Location = new System.Drawing.Point(47, 116);
            this.textBoxEmailPath.Name = "textBoxEmailPath";
            this.textBoxEmailPath.Size = new System.Drawing.Size(376, 20);
            this.textBoxEmailPath.TabIndex = 5;
            // 
            // labelControllerEmail
            // 
            this.labelControllerEmail.AutoSize = true;
            this.labelControllerEmail.Location = new System.Drawing.Point(47, 155);
            this.labelControllerEmail.Name = "labelControllerEmail";
            this.labelControllerEmail.Size = new System.Drawing.Size(79, 13);
            this.labelControllerEmail.TabIndex = 9;
            this.labelControllerEmail.Text = "Report address";
            // 
            // textBoxControllerEmail
            // 
            this.textBoxControllerEmail.Enabled = false;
            this.textBoxControllerEmail.Location = new System.Drawing.Point(47, 171);
            this.textBoxControllerEmail.Name = "textBoxControllerEmail";
            this.textBoxControllerEmail.Size = new System.Drawing.Size(410, 20);
            this.textBoxControllerEmail.TabIndex = 8;
            // 
            // checkBoxMail
            // 
            this.checkBoxMail.AutoSize = true;
            this.checkBoxMail.Location = new System.Drawing.Point(47, 197);
            this.checkBoxMail.Name = "checkBoxMail";
            this.checkBoxMail.Size = new System.Drawing.Size(188, 17);
            this.checkBoxMail.TabIndex = 10;
            this.checkBoxMail.Text = "Mail me a report when you\'re done";
            this.checkBoxMail.UseVisualStyleBackColor = true;
            this.checkBoxMail.CheckedChanged += new System.EventHandler(this.checkBoxMail_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 302);
            this.Controls.Add(this.checkBoxMail);
            this.Controls.Add(this.labelControllerEmail);
            this.Controls.Add(this.textBoxControllerEmail);
            this.Controls.Add(this.labelEmailPath);
            this.Controls.Add(this.buttonOpenFileEmail);
            this.Controls.Add(this.textBoxEmailPath);
            this.Controls.Add(this.buttonCheckAndRemind);
            this.Controls.Add(this.buttonCheck);
            this.Controls.Add(this.labelFilePath);
            this.Controls.Add(this.buttonOpenFile);
            this.Controls.Add(this.textBoxFilePath);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFilePath;
        private System.Windows.Forms.Button buttonOpenFile;
        private System.Windows.Forms.Label labelFilePath;
        private System.Windows.Forms.Button buttonCheck;
        private System.Windows.Forms.Button buttonCheckAndRemind;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label labelEmailPath;
        private System.Windows.Forms.Button buttonOpenFileEmail;
        private System.Windows.Forms.TextBox textBoxEmailPath;
        private System.Windows.Forms.Label labelControllerEmail;
        private System.Windows.Forms.TextBox textBoxControllerEmail;
        private System.Windows.Forms.CheckBox checkBoxMail;
    }
}

