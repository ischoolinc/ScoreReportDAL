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
using System.IO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateClassGPlan : BaseForm
    {
        DataAccess _da = new DataAccess();
        string _SelectGroupName = "";
        // 已有課程規劃表
        bool hasGPlan = false;

        List<string> _SelGroupCodeList;
        List<string> _ErrorList;
        BackgroundWorker _bgWorker;

        public frmCreateClassGPlan()
        {
            InitializeComponent();
            _SelGroupCodeList = new List<string>();
            _ErrorList = new List<string>();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;


        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("建立班級課程規劃表 產生完成");

            if (_ErrorList.Count == 0)
            {
                MsgBox.Show("建立完成");
            }

            btnCreate.Enabled = true;

        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("建立班級課程規劃表 產生中...", e.ProgressPercentage);
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

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            // 取得畫面上選擇群科班名稱
            _SelectGroupName = cbxGroupName.Text;

            hasGPlan = false;
            string gpName = _da.GetGPlanNameByGroupName(_SelectGroupName);

            bool hasGPName = _da.CheckHasGPlan(gpName);
            if (hasGPName)
            {
                hasGPlan = true;
                MsgBox.Show(gpName + " 班級課程規劃表內已存在，無法建立。");
                return;
                //MsgBox.Show(gpName + " 班級課程規劃表內已存在，按「是」新增科目與更新學分數，請問是否執行？", "資料已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            }

            _SelGroupCodeList.Clear();
            string gpCode = _da.GetGroupCodeByName(_SelectGroupName);
            _SelGroupCodeList.Add(gpCode);
            btnCreate.Enabled = false;
            _bgWorker.RunWorkerAsync();

        }

        private void frmCreateClassGPlan_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            cbxGroupName.Enabled = false;
            btnCreate.Enabled = false;

            // 取得課程代碼大表資料並填入群科班選項
            _da.LoadMOEGroupCodeDict();

            foreach (string name in _da.GetGroupNameList())
            {
                cbxGroupName.Items.Add(name);
            }

            if (cbxGroupName.Items.Count > 0)
            {
                cbxGroupName.SelectedIndex = 0;
                btnCreate.Enabled = true;
            }

            cbxGroupName.Enabled = true;

        }
    }
}
