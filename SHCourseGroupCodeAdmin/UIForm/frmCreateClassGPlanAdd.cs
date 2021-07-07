using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using SHCourseGroupCodeAdmin.DAO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateClassGPlanAdd : BaseForm
    {
        string GPlanName = "";
        string GPCode = "";
        DataAccess _da = new DataAccess();
        BackgroundWorker _bgWorker;
        List<string> _ErrorList;
        List<string> _SelGroupCodeList;

        public frmCreateClassGPlanAdd()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _SelGroupCodeList = new List<string>();
            _ErrorList = new List<string>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程規劃表 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (_ErrorList.Count == 0)
            {
                MsgBox.Show("產生完成");
            }

            btnCreate.Enabled = true;
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _ErrorList.Clear();
            // 有傳入 Group Code 再執行
            if (_SelGroupCodeList.Count > 0)
            {
                foreach (string gpCode in _SelGroupCodeList)
                {

                    string errMsg = _da.WriteToGPlanByGroupCode(gpCode);
                    if (!string.IsNullOrEmpty(errMsg))
                        _ErrorList.Add(errMsg);
                    _bgWorker.ReportProgress(50);
                }
            }
            _bgWorker.ReportProgress(100);
        }

        public void SetGPlanCode(string code)
        {
            GPCode = code;
        }
     

        private void btnCreate_Click(object sender, EventArgs e)
        {
            btnCreate.Enabled = false;
            _SelGroupCodeList.Clear();            
            _SelGroupCodeList.Add(GPCode);
            btnCreate.Enabled = false;
            _bgWorker.RunWorkerAsync();
        }

        private void frmCreateClassGPlanAdd_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            lblName.Text = "課程規劃名稱：" + GPlanName;
        }

        public void SetPlanName(string name)
        {
            GPlanName = name;
        }
    }
}
