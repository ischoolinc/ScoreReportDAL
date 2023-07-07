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
using Aspose.Cells;
using System.IO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateClassGPlanSetDetail : BaseForm
    {
        GPlanData _GPlanData = null;
        string MOEGroupNmae = "";
        XElement MOEXml = null;
        DataAccess da = new DataAccess();

        Workbook _wb;
        Dictionary<string, int> _ColIdxDict;

        public frmCreateClassGPlanSetDetail()
        {
            InitializeComponent();
            _ColIdxDict = new Dictionary<string, int>();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //int rowIdx = 0;
            //foreach (chkSubjectInfo subj in _GPlanData.chkSubjectInfoList)
            //{
            //    if (dgData.Rows[rowIdx].Cells["處理方式"].Value != null)
            //    {
            //        subj.ProcessStatus = dgData.Rows[rowIdx].Cells["處理方式"].Value.ToString();
            //    }

            //    rowIdx++;
            //}

            _GPlanData.chkSubjectInfoList.Clear();
            foreach(DataGridViewRow drv in dgData.Rows)
            {
                chkSubjectInfo subj = drv.Tag as chkSubjectInfo;
                if (subj != null)
                {
                    if (drv.Cells["處理方式"].Value != null)
                        subj.ProcessStatus = drv.Cells["處理方式"].Value.ToString();                    
                }

                _GPlanData.chkSubjectInfoList.Add(subj);
            }

            this.DialogResult = DialogResult.OK;
        }

        public void SetGPlanData(GPlanData data)
        {
            _GPlanData = data;
        }

        public GPlanData GetGPlanData()
        {
            return _GPlanData;
        }

        public void SetMOENameAndXml(string name, XElement xmlData)
        {
            MOEGroupNmae = name;
            MOEXml = xmlData;
        }

        private void frmCreateClassGPlanSetDetail_Load(object sender, EventArgs e)
        {
            dgData.Enabled = false;

            // 載入資料欄位
            LoadDataGridViewColumns();

            // 載入資料
            LoadData();
            dgData.Enabled = true;
        }

        private void LoadDataGridViewColumns()
        {
            try
            {
                DataGridViewTextBoxColumn tbDiffStatus = new DataGridViewTextBoxColumn();
                tbDiffStatus.Name = "差異狀態";
                tbDiffStatus.Width = 100;
                tbDiffStatus.HeaderText = "差異狀態";
                tbDiffStatus.ReadOnly = true;

                DataGridViewComboBoxColumn cbProcStatus = new DataGridViewComboBoxColumn();
                cbProcStatus.Name = "處理方式";
                cbProcStatus.Width = 90;
                cbProcStatus.Width = 120;
                cbProcStatus.HeaderText = "處理方式";

                DataTable dtPitems = new DataTable();
                dtPitems.Columns.Add("ITEM");
                dtPitems.Columns.Add("VALUE");

                List<string> items = new List<string>();
                items.Add("更新");
                items.Add("略過");
                items.Add("刪除");
                items.Add("新增");


                foreach (string str in items)
                {
                    DataRow dr = dtPitems.NewRow();
                    dr["ITEM"] = str;
                    dr["VALUE"] = str;
                    dtPitems.Rows.Add(dr);
                }

                cbProcStatus.DataSource = dtPitems;
                cbProcStatus.DisplayMember = "ITEM";
                cbProcStatus.ValueMember = "VALUE";


                DataGridViewTextBoxColumn tbDomain = new DataGridViewTextBoxColumn();
                tbDomain.Name = "領域";
                tbDomain.Width = 70;
                tbDomain.HeaderText = "領域";
                tbDomain.ReadOnly = true;

                DataGridViewTextBoxColumn tbScoreType = new DataGridViewTextBoxColumn();
                tbScoreType.Name = "分項類別";
                tbScoreType.Width = 90;
                tbScoreType.HeaderText = "分項類別";
                tbScoreType.ReadOnly = true;

                DataGridViewTextBoxColumn tbSubjectName = new DataGridViewTextBoxColumn();
                tbSubjectName.Name = "科目名稱";
                tbSubjectName.Width = 150;
                tbSubjectName.HeaderText = "科目名稱";
                tbSubjectName.ReadOnly = true;

                DataGridViewTextBoxColumn tbOfficialSubjectName = new DataGridViewTextBoxColumn();
                tbOfficialSubjectName.Name = "報部科目名稱";
                tbOfficialSubjectName.Width = 150;
                tbOfficialSubjectName.HeaderText = "報部科目名稱";
                tbOfficialSubjectName.ReadOnly = true;


                DataGridViewTextBoxColumn tbRequiredBy = new DataGridViewTextBoxColumn();
                tbRequiredBy.Name = "校訂部定";
                tbRequiredBy.Width = 90;
                tbRequiredBy.HeaderText = "校訂部定";
                tbRequiredBy.ReadOnly = true;

                DataGridViewTextBoxColumn tbIsRequired = new DataGridViewTextBoxColumn();
                tbIsRequired.Name = "必選修";
                tbIsRequired.Width = 90;
                tbIsRequired.HeaderText = "必選修";
                tbIsRequired.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS11 = new DataGridViewTextBoxColumn();
                tbGS11.Name = "1上";
                tbGS11.Width = 60;
                tbGS11.HeaderText = "1上";
                tbGS11.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS12 = new DataGridViewTextBoxColumn();
                tbGS12.Name = "1下";
                tbGS12.Width = 60;
                tbGS12.HeaderText = "1下";
                tbGS12.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS21 = new DataGridViewTextBoxColumn();
                tbGS21.Name = "2上";
                tbGS21.Width = 60;
                tbGS21.HeaderText = "2上";
                tbGS21.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS22 = new DataGridViewTextBoxColumn();
                tbGS22.Name = "2下";
                tbGS22.Width = 60;
                tbGS22.HeaderText = "2下";
                tbGS22.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS31 = new DataGridViewTextBoxColumn();
                tbGS31.Name = "3上";
                tbGS31.Width = 60;
                tbGS31.HeaderText = "3上";
                tbGS31.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS32 = new DataGridViewTextBoxColumn();
                tbGS32.Name = "3下";
                tbGS32.Width = 60;
                tbGS32.HeaderText = "3下";
                tbGS32.ReadOnly = true;

                DataGridViewTextBoxColumn tbOpenStatus = new DataGridViewTextBoxColumn();
                tbOpenStatus.Name = "開課方式";
                tbOpenStatus.Width = 90;
                tbOpenStatus.HeaderText = "開課方式";
                tbOpenStatus.ReadOnly = true;

                DataGridViewTextBoxColumn tbCourseCode = new DataGridViewTextBoxColumn();
                tbCourseCode.Name = "課程代碼";
                tbCourseCode.Width = 300;
                tbCourseCode.HeaderText = "課程代碼";
                tbCourseCode.ReadOnly = true;

                dgData.Columns.Add(tbDiffStatus);
                dgData.Columns.Add(cbProcStatus);
                dgData.Columns.Add(tbDomain);
                dgData.Columns.Add(tbScoreType);
                dgData.Columns.Add(tbOfficialSubjectName);
                dgData.Columns.Add(tbSubjectName);
                dgData.Columns.Add(tbRequiredBy);
                dgData.Columns.Add(tbIsRequired);
                dgData.Columns.Add(tbGS11);
                dgData.Columns.Add(tbGS12);
                dgData.Columns.Add(tbGS21);
                dgData.Columns.Add(tbGS22);
                dgData.Columns.Add(tbGS31);
                dgData.Columns.Add(tbGS32);
                dgData.Columns.Add(tbOpenStatus);
                dgData.Columns.Add(tbCourseCode);

                //// 因為自動排序有些問題，先將關閉
                //foreach(DataGridViewColumn col in dgData.Columns )
                //{
                //    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private void LoadData()
        {
            lblGroupName.Text = MOEGroupNmae;
            lblGPName.Text = _GPlanData.Name;
            lblDiffCount.Text = _GPlanData.calSubjDiffCount() + "";
            lblUpdateCount.Text = _GPlanData.calSubjUpdateCount() + "";
            lblDelCount.Text = _GPlanData.calSubjDelCount() + "";
            lblNoChangeCount.Text = _GPlanData.calSubjNoChangeCount() + "";

            dgData.Rows.Clear();

            foreach (chkSubjectInfo subj in _GPlanData.chkSubjectInfoList)
            {
                int rowIdx = dgData.Rows.Add();
                try
                {
                    Console.WriteLine("subj.ProcessStatus:" + subj.ProcessStatus);
                    dgData.Rows[rowIdx].Tag = subj;
                    dgData.Rows[rowIdx].Cells["差異狀態"].Value = string.Join(",", subj.DiffStatusList.ToArray());
                    dgData.Rows[rowIdx].Cells["處理方式"].Value = subj.ProcessStatus;
                    dgData.Rows[rowIdx].Cells["領域"].Value = subj.Domain;
                    dgData.Rows[rowIdx].Cells["分項類別"].Value = subj.Entry;
                    dgData.Rows[rowIdx].Cells["科目名稱"].Value = subj.SubjectName;
                    dgData.Rows[rowIdx].Cells["報部科目名稱"].Value = subj.OfficialSubjectName;

                    dgData.Rows[rowIdx].Cells["校訂部定"].Value = subj.RequiredBy;
                    dgData.Rows[rowIdx].Cells["必選修"].Value = subj.isRequired;


                    // 缺少解析方式 學分
                    if (subj.DiffStatusList.Contains("缺"))
                    {
                        if (subj.credit_period != null && subj.credit_period.Length == 6)
                        {
                            dgData.Rows[rowIdx].Cells["1上"].Value = subj.credit_period.Substring(0, 1);
                            dgData.Rows[rowIdx].Cells["1下"].Value = subj.credit_period.Substring(1, 1);
                            dgData.Rows[rowIdx].Cells["2上"].Value = subj.credit_period.Substring(2, 1);
                            dgData.Rows[rowIdx].Cells["2下"].Value = subj.credit_period.Substring(3, 1);
                            dgData.Rows[rowIdx].Cells["3上"].Value = subj.credit_period.Substring(4, 1);
                            dgData.Rows[rowIdx].Cells["3下"].Value = subj.credit_period.Substring(5, 1);
                        }
                    }
                    else
                    {
                        foreach (XElement elm in subj.GPlanXml)
                        {
                            if (elm.Attribute("GradeYear").Value == "1" && elm.Attribute("Semester").Value == "1")
                                dgData.Rows[rowIdx].Cells["1上"].Value = elm.Attribute("Credit").Value;

                            if (elm.Attribute("GradeYear").Value == "1" && elm.Attribute("Semester").Value == "2")
                                dgData.Rows[rowIdx].Cells["1下"].Value = elm.Attribute("Credit").Value;

                            if (elm.Attribute("GradeYear").Value == "2" && elm.Attribute("Semester").Value == "1")
                                dgData.Rows[rowIdx].Cells["2上"].Value = elm.Attribute("Credit").Value;


                            if (elm.Attribute("GradeYear").Value == "2" && elm.Attribute("Semester").Value == "2")
                                dgData.Rows[rowIdx].Cells["2下"].Value = elm.Attribute("Credit").Value;


                            if (elm.Attribute("GradeYear").Value == "3" && elm.Attribute("Semester").Value == "1")
                                dgData.Rows[rowIdx].Cells["3上"].Value = elm.Attribute("Credit").Value;

                            if (elm.Attribute("GradeYear").Value == "3" && elm.Attribute("Semester").Value == "2")
                                dgData.Rows[rowIdx].Cells["3下"].Value = elm.Attribute("Credit").Value;

                        }

                    }


                    dgData.Rows[rowIdx].Cells["開課方式"].Value = subj.OpenStatus;
                    dgData.Rows[rowIdx].Cells["課程代碼"].Value = subj.CourseCode;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void lnkViewGroupDetail_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkViewGroupDetail.Enabled = false;
            if (_GPlanData != null)
            {
                List<MOECourseCodeInfo> CourseData = da.GetCourseGroupCodeListByGroupCode(_GPlanData.MOEGroupCode);

                // 填值到 Excel
                _wb = new Workbook(new MemoryStream(Properties.Resources.課程代碼表樣版));
                Worksheet wst = _wb.Worksheets[0];

                _ColIdxDict.Clear();

                // 讀取欄位與索引            
                for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
                {
                    _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
                }

                int rowIdx = 1;
                foreach (MOECourseCodeInfo data in CourseData)
                {
                    wst.Cells[rowIdx, GetColIndex("群組代碼")].PutValue(data.group_code);
                    wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.course_code);
                    wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.subject_name);
                    wst.Cells[rowIdx, GetColIndex("入學年")].PutValue(data.entry_year);
                    wst.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.require_by);
                    wst.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.is_required);
                    wst.Cells[rowIdx, GetColIndex("課程類型")].PutValue(data.course_type);
                    wst.Cells[rowIdx, GetColIndex("群別")].PutValue(data.group_type);
                    wst.Cells[rowIdx, GetColIndex("科別")].PutValue(data.subject_type);
                    wst.Cells[rowIdx, GetColIndex("班群")].PutValue(data.class_type);
                    wst.Cells[rowIdx, GetColIndex("授課學期學分/節數")].PutValue(data.credit_period);
                    wst.Cells[rowIdx, GetColIndex("授課開課方式")].PutValue(data.open_type);
                    wst.Cells[rowIdx, GetColIndex("課程屬性")].PutValue(data.course_attr);
                    rowIdx++;
                }

                wst.AutoFitColumns();


                if (_wb != null)
                {
                    Utility.ExprotXls("課程代碼表", _wb);
                }

            }          


            lnkViewGroupDetail.Enabled = true;
        }

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }


        private void dgData_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                chkSubjectInfo subj = dgData.Rows[e.RowIndex].Tag as chkSubjectInfo;

                if (subj != null && subj.DiffMessageList.Count > 0)
                    dgData.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = string.Join(",", subj.DiffMessageList.ToArray());

            }
        }

        private void DataCount()
        {
            int diffCount = 0, updateCount = 0, delCount = 0, noChangeCount = 0;

            foreach (DataGridViewRow dr in dgData.Rows)
            {
                if (dr.Cells["處理方式"].Value != null)
                {
                    if (dr.Cells["處理方式"].Value.ToString() != "略過")
                        diffCount++;

                    if (dr.Cells["處理方式"].Value.ToString() == "略過")
                        noChangeCount++;

                    if (dr.Cells["處理方式"].Value.ToString() == "刪除")
                        delCount++;

                    if (dr.Cells["處理方式"].Value.ToString() == "更新")
                        updateCount++;
                }
            }

            lblDiffCount.Text = diffCount + "";
            lblUpdateCount.Text = updateCount + "";
            lblDelCount.Text = delCount + "";
            lblNoChangeCount.Text = noChangeCount + "";
        }

        private void dgData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex == 1)
            {
                DataCount();
            }
        }

        private void dgData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex == 1)
            {
                DataCount();
            }
        }

        private void dgData_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           
        }
    }
}
