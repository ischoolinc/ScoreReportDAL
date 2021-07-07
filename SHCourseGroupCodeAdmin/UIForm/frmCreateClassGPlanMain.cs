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
    public partial class frmCreateClassGPlanMain : BaseForm
    {
        DataAccess _da = new DataAccess();
        // 選擇的群科班代碼
        string SelectGroupCode = "";

        string SelectGroupName = "";

        // 課程規劃表名稱
        string GPlanName = "";

        BackgroundWorker _bwWorker;

        // 使用課程規劃表
        List<GPlanData> _GPlanDataList;

        public frmCreateClassGPlanMain()
        {
            InitializeComponent();
            _bwWorker = new BackgroundWorker();
            _GPlanDataList = new List<GPlanData>();
            _bwWorker.DoWork += _bwWorker_DoWork;
            _bwWorker.ProgressChanged += _bwWorker_ProgressChanged;
            _bwWorker.RunWorkerCompleted += _bwWorker_RunWorkerCompleted;
            _bwWorker.WorkerReportsProgress = true;
        }

        private void _bwWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            btnNext.Enabled = true;
            if (_GPlanDataList.Count > 0)
            {
                // 已有課程規劃表需要調整
                frmCreateClassGPlanHasData fHasData = new frmCreateClassGPlanHasData();
                fHasData.SetGroupName(SelectGroupName);
                fHasData.ShowDialog();
            }
            else
            {
                // 沒有相對課程規劃表需要新增
                frmCreateClassGPlanAdd fAdd = new frmCreateClassGPlanAdd();
                fAdd.SetPlanName(GPlanName);
                fAdd.SetGPlanCode(SelectGroupCode);
                fAdd.ShowDialog();
            }
        }

        private void _bwWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取課程規劃表...", e.ProgressPercentage);
        }

        private void _bwWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bwWorker.ReportProgress(1);
            _GPlanDataList.Clear();
            _da.LoadMOEGroupCodeDict();
            GPlanName = _da.GetGPlanNameByCode(SelectGroupCode);
            _GPlanDataList = _da.GetGPlanDataListByMOECode(SelectGroupCode);
            _bwWorker.ReportProgress(100);
        }

        private void frmCreateClassGPlanMain_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;

            btnNext.Enabled = false;

            // 取得課程代碼大表資料並填入群科班選項
            _da.LoadMOEGroupCodeDict();

            foreach (string name in _da.GetGroupNameList())
            {
                cbxGroupName.Items.Add(name);
            }

            if (cbxGroupName.Items.Count > 0)
            {
                cbxGroupName.SelectedIndex = 0;
                btnNext.Enabled = true;
            }

            cbxGroupName.Enabled = true;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            btnNext.Enabled = false;

            SelectGroupName = cbxGroupName.Text;
            // 取得群科班代碼
            SelectGroupCode = _da.GetGroupCodeByName(SelectGroupName);

            _bwWorker.RunWorkerAsync();

        }
    }
}
