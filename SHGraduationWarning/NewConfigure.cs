﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;


namespace SHGraduationWarning
{
    public partial class NewConfigure : FISCA.Presentation.Controls.BaseForm
    {
        public Aspose.Words.Document Template { get; private set; }
        public int SubjectLimit { get; private set; }
        public string ConfigName { get; private set; }
        public bool WithSchoolYearScore { get; private set; }
        public bool WithPrevSemesterScore { get; private set; }

        public NewConfigure()
        {
            InitializeComponent();
            checkBoxX1.CheckedChanged += new EventHandler(SetupTemplate);
            checkBoxX2.CheckedChanged += new EventHandler(UploadTemplate);
        }
        private void SetupTemplate(object sender, EventArgs e)
        {
            if (checkBoxX1.Checked)
            {
                Template = new Aspose.Words.Document(new MemoryStream(Properties.Resources.普通高中個人畢業預警通知書樣版));
            }
        }
        private void UploadTemplate(object sender, EventArgs e)
        {
            if (checkBoxX2.Checked)
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "上傳樣板";
                dialog.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        Template = new Aspose.Words.Document(dialog.FileName);
                       // List<string> fields = new List<string>(Template.MailMerge.GetFieldNames());
                    }
                    catch
                    {
                        MessageBox.Show("樣板開啟失敗");
                        checkBoxX2.Checked = false;
                    }
                }
                else
                    checkBoxX2.Checked = false;
            }
        }

        private void checkReady(object sender, EventArgs e)
        {
            bool ready = true;
            if (txtName.Text == "")
                ready = false;
            else
                ConfigName = txtName.Text;
            if (!checkBoxX1.Checked && !checkBoxX2.Checked)
            {
                ready = false;
            }
            btnSubmit.Enabled = ready;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 儲存檔案
            
            string reportName = "普通高中個人畢業預警通知書樣板";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                Document document = new Document(new MemoryStream(Properties.Resources.普通高中個人畢業預警通知書樣版));
                document.Save(path, Aspose.Words.SaveFormat.Docx);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        Document document = new Document(new MemoryStream(Properties.Resources.普通高中個人畢業預警通知書樣版));
                        document.Save(path, Aspose.Words.SaveFormat.Docx);

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 儲存檔案
            string inputReportName = "畢業預警通知單合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                Document document = new Document(new MemoryStream(Properties.Resources.畢業預警通知單合併欄位總表));
                document.Save(path, Aspose.Words.SaveFormat.Docx);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        Document document = new Document(new MemoryStream(Properties.Resources.畢業預警通知單合併欄位總表));
                        document.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);                  

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }
    }
}
