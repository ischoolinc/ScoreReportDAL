using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using K12.Data;
using System.IO;
using DevComponents.DotNetBar.Controls;
using Aspose.Words;
using System.Xml.Linq;

namespace SHGraduationWarning
{
    public partial class ConfigForm : FISCA.Presentation.Controls.BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        private Dictionary<string, List<string>> _ExamSubjects = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> _ExamSubjectFull = new Dictionary<string, List<string>>();

        private List<Configure> _Configures = new List<Configure>();
        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        public ConfigForm()
        {
            InitializeComponent();
            List<ExamRecord> exams = new List<ExamRecord>();
            BackgroundWorker bkw = new BackgroundWorker();
            bkw.DoWork += delegate
            {
                bkw.ReportProgress(1);
                bkw.ReportProgress(10);

                bkw.ReportProgress(20);

                bkw.ReportProgress(80);
                _Configures = _AccessHelper.Select<Configure>();
                bkw.ReportProgress(100);

            };
            bkw.WorkerReportsProgress = true;
            bkw.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                //circularProgress1.Value = e.ProgressPercentage;
            };
            bkw.RunWorkerCompleted += delegate
            {
                cboConfigure.Items.Clear();
                foreach (var item in _Configures)
                {
                    cboConfigure.Items.Add(item);
                }
                cboConfigure.Items.Add(new Configure() { Name = "新增" });

                if (_Configures.Count > 0)
                {
                    cboConfigure.SelectedIndex = 0;
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            };
            bkw.RunWorkerAsync();
        }

        public Configure Configure { get; private set; }


        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Configure = new Configure();
                    Configure.Name = dialog.ConfigName;
                    Configure.Template = dialog.Template;
                    ;
                    _Configures.Add(Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;

                    Configure.Encode();
                    Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = true;
                    Configure = _Configures[cboConfigure.SelectedIndex];
                    if (Configure.Template == null)
                        Configure.Decode();


                }
                else
                {
                    Configure = null;

                }
            }
        }


        private void SaveTemplate(object sender, EventArgs e)
        {
            if (Configure == null) return;
            Configure.Active = true;
            Configure.Encode();
            
            Configure.Save();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 儲存檔案            
            string reportName = "個人畢業預警通知單合併欄位總表";

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
                document.Save(path, SaveFormat.Docx);
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
                        document.Save(sd.FileName, SaveFormat.Docx);

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
            if (Configure == null) return;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    this.Configure.Template = new Aspose.Words.Document(dialog.FileName);
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.Configure == null) return;
            #region 儲存檔案            
            string reportName = "畢業預警報表樣板(" + this.Configure.Name + ")";

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
             
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                this.Configure.Template.Save(stream, Aspose.Words.SaveFormat.Docx);             
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        this.Configure.Template.Save(stream, Aspose.Words.SaveFormat.Docx);
                        stream.Flush();
                        stream.Close();

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

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _Configures.Remove(Configure);
                if (Configure.UID != "")
                {
                    Configure.Deleted = true;
                    Configure.Save();
                }
                var conf = Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.Template = Configure.Template;
                conf.Encode();
                conf.Save();
                _Configures.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }


        private void ConfigForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            SaveTemplate(null, null);
            this.DialogResult = DialogResult.OK;
        }
    }
}
