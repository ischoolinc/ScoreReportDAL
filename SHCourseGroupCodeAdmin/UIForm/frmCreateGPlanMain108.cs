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
    public partial class frmCreateGPlanMain108 : BaseForm
    {
        BackgroundWorker _bgWorker;
        DataAccess _da;
        List<GPlanInfo108> _GPlanInfo108List;

        public frmCreateGPlanMain108()
        {
            InitializeComponent();
            _da = new DataAccess();
            _bgWorker = new BackgroundWorker();
            _GPlanInfo108List = new List<GPlanInfo108>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取課程代碼大表與課程規劃表...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            foreach (GPlanInfo108 data in _GPlanInfo108List)
            {
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = data;
                dgData.Rows[rowIdx].Cells[colEntrySchoolYear.Index].Value = data.EntrySchoolYear;
                dgData.Rows[rowIdx].Cells[colGroupName.Index].Value = data.GDCName;
                dgData.Rows[rowIdx].Cells[colChangeDesc.Index].Value = data.Status;
                dgData.Rows[rowIdx].Cells[colUpdateSetup.Index].Value = "設定";
            }

            GPlanDataCount();
            ControlEnable(true);
        }

        /// <summary>
        /// 資料統計
        /// </summary>
        private void GPlanDataCount()
        {
            int AddCount = 0, UpdateCount = 0, NoChangeCount = 0;
            
            foreach(DataGridViewRow drv in dgData.Rows)                
            {
                if (drv.IsNewRow)
                    continue;

                string Status = drv.Cells[colChangeDesc.Index].Value.ToString();

                if (Status == "新增")
                    AddCount++;

                if (Status == "更新")
                    UpdateCount++;

                if (Status == "無變動")
                    NoChangeCount = 0;
            }

            lblGroupCount.Text = "群科班數" + _GPlanInfo108List.Count + "筆";
            lblAddCount.Text = "新增" + AddCount + "筆";
            lblUpdateCount.Text = "更新" + UpdateCount + "筆";
            lblNoChangeCount.Text = "無變動" + NoChangeCount + "筆";
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _GPlanInfo108List = _da.GPlanInfo108List();
            _bgWorker.ReportProgress(30);
            
            // 解析課程代碼大表 XML
            foreach(GPlanInfo108 data in _GPlanInfo108List)
            {
                data.ParseMOEXml();

                data.ParseRefGPContentXml();

                data.CheckData();

                if (string.IsNullOrEmpty(data.RefGPID))
                    data.Status = "新增";
            }
            
            _bgWorker.ReportProgress(50);
            _bgWorker.ReportProgress(100);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCreateGPlanMain108_Load(object sender, EventArgs e)
        {
            ControlEnable(false);
            _bgWorker.RunWorkerAsync();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {

        }

        private void ControlEnable(bool value)
        {
            dgData.Enabled = btnCreate.Enabled = btnQueryAndSet.Enabled = value;
        }

        private void btnQueryAndSet_Click(object sender, EventArgs e)
        {
            frmCreateGPlanQueryAndSetup108 fgq = new frmCreateGPlanQueryAndSetup108();
            fgq.ShowDialog();
        }

        private void dgData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                if (e.ColumnIndex == colUpdateSetup.Index)
                {
                    GPlanInfo108 data = dgData.Rows[e.RowIndex].Tag as GPlanInfo108;
                    if (data != null)
                    {
                        frmCreateGPlanItemSetup108 fgpd = new frmCreateGPlanItemSetup108();
                        fgpd.SetGPlanInfo(data);

                        if (fgpd.ShowDialog() == DialogResult.OK)
                        {
                            GPlanInfo108 newData = fgpd.GetGPlanInfo();
                            dgData.Rows[e.RowIndex].Tag = newData;
                            dgData.Rows[e.RowIndex].Cells[colChangeDesc.Index].Value = newData.Status;
                        }
                    }

                    GPlanDataCount();
                }
            }
        }
    }
}
