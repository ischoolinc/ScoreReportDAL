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
    public partial class frmCreateGPlanMainBatch : BaseForm
    {
        BackgroundWorker _bgWorker;
        DataAccess _da;
        List<GPlanInfo108> _GPlanInfo108List;

        public frmCreateGPlanMainBatch()
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取課程代碼總表與課程規劃表...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dgData.Rows.Clear();
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取完成。");
            _GPlanInfo108List = (from data in _GPlanInfo108List orderby data.OrderByInt, data.EntrySchoolYear descending, data.GDCName select data).ToList();

            foreach (GPlanInfo108 data in _GPlanInfo108List)
            {
           
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = data;
                dgData.Rows[rowIdx].Cells[colEntrySchoolYear.Index].Value = data.EntrySchoolYear;
                dgData.Rows[rowIdx].Cells[colGroupName.Index].Value = data.GDCName;
                dgData.Rows[rowIdx].Cells[colChangeDesc.Index].Value = data.Status;
                dgData.Rows[rowIdx].Cells[colGpName.Index].Value = data.RefGPName;
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

            foreach (DataGridViewRow drv in dgData.Rows)
            {
                if (drv.IsNewRow)
                    continue;

                string Status = drv.Cells[colChangeDesc.Index].Value.ToString();

                if (Status == "新增")
                    AddCount++;

                if (Status == "更新")
                    UpdateCount++;

                if (Status == "無變動")
                    NoChangeCount++;
            }

            lblGroupCount.Text = "群科班數" + _GPlanInfo108List.Count + "筆";
            lblAddCount.Text = "新增" + AddCount + "筆";
            lblUpdateCount.Text = "更新" + UpdateCount + "筆";
            lblNoChangeCount.Text = "無變動" + NoChangeCount + "筆";
        }


        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _GPlanInfo108List = _da.GPlanInfoOld108List();
            _bgWorker.ReportProgress(30);

            // 解析課程代碼大表 XML
            foreach (GPlanInfo108 data in _GPlanInfo108List)
            {
                data.ParseMOEXml();

                data.ParseRefGPContentXml();

                data.CheckData();

                foreach (chkSubjectInfo subj in data.chkSubjectInfoList)
                {
                    subj.GDCCode = data.GDCCode;
                }

                if (string.IsNullOrEmpty(data.RefGPID))
                    data.Status = "新增";

                if (data.calSubjDiffCount() > 0)
                    data.Status = "更新";

                data.ParseOrderByInt();
            }

            _bgWorker.ReportProgress(50);
            _bgWorker.ReportProgress(100);
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            ControlEnable(false);

            try
            {
                List<GPlanInfo108> insertDataList = new List<GPlanInfo108>();
                List<GPlanInfo108> updateDataList = new List<GPlanInfo108>();

                foreach (GPlanInfo108 data in _GPlanInfo108List)
                {
                    if (data.Status == "新增")
                        insertDataList.Add(data);

                    if (data.Status == "更新")
                        updateDataList.Add(data);
                }

                // 新增資料
                List<string> insertSQLList = new List<string>();
                if (insertDataList.Count > 0)
                {
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();

                    try
                    {
                        foreach (GPlanInfo108 data in insertDataList)
                        {
                            string sql = "" +
                                " INSERT INTO graduation_plan(" +
" name " +
" ,content " +
" ,moe_group_code )  " +
" VALUES( " +
" '" + data.GDCName + "' " +
" ,'" + data.MOEXml.ToString() + "' " +
" ,'" + data.GDCCode + "' " +
" ); ";
                            insertSQLList.Add(sql);
                        }
                        uh.Execute(insertSQLList);
                        MsgBox.Show("新增" + insertSQLList.Count + "筆課程規劃表");
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show("新增資料發生錯誤：" + ex.Message);
                    }
                }

                // 更新資料
                List<string> updateSQLList = new List<string>();
                if (updateDataList.Count > 0)
                {
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
                    
                    try
                    {
                        foreach (GPlanInfo108 data in updateDataList)
                        {
                            
                            if (!string.IsNullOrEmpty(data.RefGPID))
                            {
                                string sql = "UPDATE graduation_plan SET content = '" + data.RefGPContent + "' WHERE id = " + data.RefGPID + ";";

                                updateSQLList.Add(sql);
                            }
                        }

                        // 更新資料
                        uh.Execute(updateSQLList);
                        MsgBox.Show("更新" + updateSQLList.Count + "筆課程規劃表");
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show("更新資料發生錯誤" + ex.Message);
                    }
                }

                if (insertSQLList.Count > 0 || updateSQLList.Count > 0)
                {
                    ControlEnable(false);
                    _bgWorker.RunWorkerAsync();
                }

            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存發生錯誤：" + ex.Message);
            }

            ControlEnable(true);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void ControlEnable(bool value)
        {
            dgData.Enabled = btnCreate.Enabled = btnQueryAndSet.Enabled = value;
        }


        private void btnQueryAndSet_Click(object sender, EventArgs e)
        {
            btnQueryAndSet.Enabled = false;
            frmCreateGPlanQueryAndSetup108 fgq = new frmCreateGPlanQueryAndSetup108();
            fgq.SetGPlanInfos(_GPlanInfo108List);

            if (fgq.ShowDialog() == DialogResult.OK)
            {
                Dictionary<string, GPlanInfo108> dataDict = fgq.GetGPlanInfoDicts();
                foreach (DataGridViewRow drv in dgData.Rows)
                {
                    if (drv.IsNewRow)
                        continue;

                    GPlanInfo108 data = drv.Tag as GPlanInfo108;
                    if (data != null)
                    {
                        if (dataDict.ContainsKey(data.GDCCode))
                        {
                            dataDict[data.GDCCode].ParseStatus();
                            drv.Tag = dataDict[data.GDCCode];
                            drv.Cells[colChangeDesc.Index].Value = dataDict[data.GDCCode].Status;
                        }
                    }
                }

                GPlanDataCount();
            }

            btnQueryAndSet.Enabled = true;
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
                            newData.ParseStatus();
                            dgData.Rows[e.RowIndex].Tag = newData;
                            dgData.Rows[e.RowIndex].Cells[colChangeDesc.Index].Value = newData.Status;
                        }
                    }

                    GPlanDataCount();
                }
            }
        }

        private void frmCreateGPlanMainBatch_Load(object sender, EventArgs e)
        {
            ControlEnable(false);
            _bgWorker.RunWorkerAsync();
        }
    }
}
