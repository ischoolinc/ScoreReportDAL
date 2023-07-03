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
using System.Xml.Linq;
using FISCA.LogAgent;
using System.IO;

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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取課程代碼總表與課程規劃表...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            int currentSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);

            dgData.Rows.Clear();
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取完成。");
            _GPlanInfo108List = (from data in _GPlanInfo108List where int.Parse(data.EntrySchoolYear) > (currentSchoolYear - 3) orderby data.OrderByInt, data.EntrySchoolYear descending, data.GDCName select data).ToList();

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
            int AddCount = 0, UpdateCount = 0, NoChangeCount = 0, UpdateLevelCount = 0;

            foreach (DataGridViewRow drv in dgData.Rows)
            {
                if (drv.IsNewRow)
                    continue;

                string Status = drv.Cells[colChangeDesc.Index].Value.ToString();

                if (Status == "新增")
                    AddCount++;

                if (Status == "更新")
                    UpdateCount++;

                if (Status == "級別更新")
                    UpdateLevelCount++;

                if (Status == "無變動")
                    NoChangeCount++;
            }

            lblGroupCount.Text = "群科班數" + _GPlanInfo108List.Count + "筆";
            lblAddCount.Text = "新增" + AddCount + "筆";
            lblUpdateCount.Text = "更新" + UpdateCount + "筆";
            lblUpdateLevelCount.Text = "級別更新" + UpdateLevelCount + "筆";
            lblNoChangeCount.Text = "無變動" + NoChangeCount + "筆";
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _GPlanInfo108List = _da.GPlanInfo108List();
            _bgWorker.ReportProgress(30);

            List<string> newGPNameList = new List<string>();

            // 處理綜合型高中學術學程1年級不分群不分班群，群科班代碼包含這串 M111960
            // 請空田值
            Global._GPlanInfo108MDict.Clear();
            foreach (GPlanInfo108 data in _GPlanInfo108List)
            {
                if (string.IsNullOrEmpty(data.RefGPID))
                {
                    newGPNameList.Add(data.GDCName);
                }
                data.ParseMOEXml();
                data.ParseRefGPContentXml();

                if (data.GDCCode.IndexOf("M111960") > -1)
                {
                    // 取前面學年度
                    string key = data.GDCCode.Substring(0, 3);
                    if (!Global._GPlanInfo108MDict.ContainsKey(key))
                        Global._GPlanInfo108MDict.Add(key, data);
                }
            }

            Dictionary<string, string> hadGPNameIDDict = _da.GetGPNameIDByNameList(newGPNameList);

            _bgWorker.ReportProgress(50);

            // 解析課程代碼大表 XML
            foreach (GPlanInfo108 data in _GPlanInfo108List)
            {
                //data.ParseMOEXml();

                //data.ParseRefGPContentXml();

                data.CheckData();

                foreach (chkSubjectInfo subj in data.chkSubjectInfoList)
                {
                    subj.GDCCode = data.GDCCode;
                }

                // 檢查系統內可能沒有群科班代碼或代碼不同，但是組合課程規劃表明稱會相同，用新的更新舊的
                if (hadGPNameIDDict.ContainsKey(data.GDCName))
                {
                    data.RefGPID = hadGPNameIDDict[data.GDCName];
                }

                if (string.IsNullOrEmpty(data.RefGPID))
                    data.Status = "新增";

                if (data.calSubjDiffCount() > 0 && !string.IsNullOrEmpty(data.RefGPID))
                    data.Status = "更新";

                if (data.needUpdateEntryYear)
                    data.Status = "更新";

                data.ParseOrderByInt();
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
            ControlEnable(false);
            try
            {
                List<GPlanInfo108> insertDataList = new List<GPlanInfo108>();
                List<GPlanInfo108> updateDataList = new List<GPlanInfo108>();

                foreach (GPlanInfo108 data in _GPlanInfo108List)
                {
                    if (data.Status == "新增")
                        insertDataList.Add(data);

                    if (data.Status == "更新" || data.Status == "級別更新")
                        updateDataList.Add(data);
                }

                if (insertDataList.Count == 0 && updateDataList.Count == 0)
                {
                    MsgBox.Show("沒有新增或更新資料。");
                    ControlEnable(true);
                    return;
                }

                // 更新資料
                List<string> updateSQLList = new List<string>();
                if (updateDataList.Count > 0)
                {
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();

                    try
                    {
                        StringBuilder sbUpd = new StringBuilder();
                        sbUpd.AppendLine("更新課程規劃表資料：");

                        foreach (GPlanInfo108 data in updateDataList)
                        {
                            XElement GPlanXml = new XElement("GraduationPlan");
                            foreach (chkSubjectInfo subj in data.chkSubjectInfoList)
                            {
                                if (subj.ProcessStatus == "更新" || subj.ProcessStatus == "新增")
                                {
                                    foreach (XElement elm in subj.MOEXml)
                                    {
                                        GPlanXml.Add(elm);
                                    }
                                }

                                if (subj.ProcessStatus == "級別更新")
                                {
                                    foreach (XElement element in subj.GPlanXml)
                                    {
                                        Utility.CalculateSubjectLevel(element);
                                        GPlanXml.Add(element);
                                    }
                                }

                                if (subj.ProcessStatus == "略過")
                                {
                                    foreach (XElement elm in subj.GPlanXml)
                                    {
                                        GPlanXml.Add(elm);
                                    }
                                }
                            }

                            // 依課程代碼,科目名稱排序
                            List<XElement> orderList = (from elmData in GPlanXml.Elements("Subject") orderby elmData.Attribute("課程代碼").Value ascending, elmData.Attribute("SubjectName").Value select elmData).ToList();
                            //  List<XElement> orderList = GPlanXml.Elements("Subject").OrderBy(x => x.Attribute("課程代碼").Value).ToList();

                            // 重整 idx                                
                            int rowIdx = 0;
                            string tmpCode = "";
                            foreach (XElement elm in orderList)
                            {
                                string codeA = elm.Attribute("課程代碼").Value + elm.Attribute("SubjectName").Value;
                                if (codeA != tmpCode)
                                {
                                    rowIdx++;
                                    tmpCode = codeA;
                                }

                                if (elm.Element("Grouping") != null)
                                {
                                    elm.Element("Grouping").SetAttributeValue("RowIndex", rowIdx);
                                }
                            }

                            GPlanXml.ReplaceAll(orderList);
                            if (data.GDCCode.Length > 3)
                            {
                                GPlanXml.SetAttributeValue("EntryYear", data.GDCCode.Substring(0, 3));

                                // 舊有
                                GPlanXml.SetAttributeValue("SchoolYear", data.GDCCode.Substring(0, 3));
                            }

                            // 檢查是否有自訂科目及群組設定，找到後加找到的放進去
                            if (data.RefGPContentXml != null)
                            {
                                if (data.RefGPContentXml.Element("使用者自訂科目") != null)
                                {
                                    XElement elmUD = new XElement(data.RefGPContentXml.Element("使用者自訂科目"));
                                    GPlanXml.Add(elmUD);
                                }

                                if (data.RefGPContentXml.Element("CourseGroupSetting") != null)
                                {
                                    XElement courseGroupSetting = new XElement(data.RefGPContentXml.Element("CourseGroupSetting"));
                                    GPlanXml.Add(courseGroupSetting);
                                }
                            }



                            data.RefGPContent = GPlanXml.ToString();


                            if (!string.IsNullOrEmpty(data.RefGPID))
                            {

                                string sql = "UPDATE graduation_plan SET content = '" + data.RefGPContent + "',moe_group_code='" + data.GDCCode + "' WHERE id = " + data.RefGPID + ";";

                                updateSQLList.Add(sql);
                                // log
                                sbUpd.AppendLine("課程規劃表名稱：" + data.GDCName + "，群科班代碼：" + data.GDCCode + "。");
                            }
                        }

                        // 更新資料
                        uh.Execute(updateSQLList);
                        // log data
                        ApplicationLog.Log("課程規劃表.更新課程規劃表", sbUpd.ToString());

                        MsgBox.Show("更新" + updateSQLList.Count + "筆課程規劃表");
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show("更新資料發生錯誤" + ex.Message);
                    }
                }


                // 新增資料
                List<string> insertSQLList = new List<string>();
                if (insertDataList.Count > 0)
                {
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();

                    try
                    {
                        StringBuilder sbIns = new StringBuilder();
                        sbIns.AppendLine("新增課程規劃表資料：");

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

                            // log
                            sbIns.AppendLine("課程規劃表名稱：" + data.GDCName + "，群科班代碼：" + data.GDCCode + "。");
                        }
                        uh.Execute(insertSQLList);

                        //// debug
                        //foreach (string str in insertSQLList)
                        //{
                        //    try
                        //    {
                        //        uh.Execute(str);
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        using (StreamWriter sw = new StreamWriter(Application.StartupPath + "\\debug.txt", false))
                        //        {
                        //            sw.WriteLine(str);
                        //            sw.WriteLine(ex.Message);
                        //        }
                        //    }
                        //}

                        // log data
                        ApplicationLog.Log("課程規劃表.新增課程規劃表", sbIns.ToString());

                        MsgBox.Show("新增" + insertSQLList.Count + "筆課程規劃表");
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show("新增資料發生錯誤：" + ex.Message);
                    }
                }


                if (insertSQLList.Count > 0 || updateSQLList.Count > 0)
                {
                    try
                    {
                        FISCA.Features.Invoke("GraduationPlanSyncAllBackground");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

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

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(dgData.CurrentCell.ColumnIndex + ":" + dgData.CurrentCell.RowIndex);
        }

        private void dgData_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                menuCalculateSubjectLevel.Show(Cursor.Position.X, Cursor.Position.Y);
            }
        }

        private void menuItemMultiCalculateSubjectLevel_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgData.SelectedRows)
            {
                GPlanInfo108 gplanInfo = (GPlanInfo108)row.Tag;
                gplanInfo.Status = "級別更新";
                row.Cells[colChangeDesc.Index].Value = gplanInfo.Status;
                foreach (chkSubjectInfo subjectInfo in gplanInfo.chkSubjectInfoList)
                {
                    subjectInfo.ProcessStatus = "級別更新";
                }
            }
        }
    }
}
