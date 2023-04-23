using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using SHSemsSubjectCheckEdit.DAO;

namespace SHSemsSubjectCheckEdit.UIForm
{
    public partial class frmSemsSubjectLevelDuplicate : BaseForm
    {
        BackgroundWorker bgWorkerLoad;

        List<string> GradeYearItems;
        List<string> ClassNameItems;
        List<string> SubjectNameItmes;
        string SelectedGradeYear = "全部";
        string SelectedClassName = "全部";
        string SelectedSubjectName = "全部";
        List<StudSubjectInfo> StudSubjectLevelDuplicateList;
        List<StudSubjectInfo> UpdateStudSubjectList;

        // 檢查有問題資料        
        Dictionary<string, StudSubjectInfo> hasErrorSubjectInfoDict;

        public frmSemsSubjectLevelDuplicate()
        {
            bgWorkerLoad = new BackgroundWorker();
            bgWorkerLoad.DoWork += BgWorkerLoad_DoWork;
            bgWorkerLoad.RunWorkerCompleted += BgWorkerLoad_RunWorkerCompleted;
            bgWorkerLoad.ProgressChanged += BgWorkerLoad_ProgressChanged;
            bgWorkerLoad.WorkerReportsProgress = true;
            GradeYearItems = new List<string>();
            ClassNameItems = new List<string>();
            SubjectNameItmes = new List<string>();

            StudSubjectLevelDuplicateList = new List<StudSubjectInfo>();
            UpdateStudSubjectList = new List<StudSubjectInfo>();
            hasErrorSubjectInfoDict = new Dictionary<string, StudSubjectInfo>();

            InitializeComponent();
        }

        private void BgWorkerLoad_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績科目級別重複查詢中...", e.ProgressPercentage);
        }

        private void BgWorkerLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            // 載入畫面
            if (hasErrorSubjectInfoDict.Values.Count > 0)
            {
                // 填入資料
                DataToDataGridView();

                // 填入選項
                LoadGradeYearClassNameAndSubjectName();
            }
            ControlEnable(true);
        }

        private void ControlEnable(bool value)
        {
            btnQuery.Enabled = btnUpdateSubjectLevel.Enabled = comboGradeYear.Enabled = checkAll.Enabled = comboClass.Enabled = comboSubject.Enabled = value;
        }


        private void BgWorkerLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            // 清空資料            
            hasErrorSubjectInfoDict.Clear();
            GradeYearItems.Clear();
            ClassNameItems.Clear();
            SubjectNameItmes.Clear();
            StudSubjectLevelDuplicateList.Clear();

            bgWorkerLoad.ReportProgress(rpInt);
            // 依畫面選學年度取得資料
            for (int gr = 1; gr <= 3; gr++)
            {
                StudSubjectLevelDuplicateList.AddRange(DataAccess.GetSemsSubjectLevelDuplicateByGradeYear(gr));
                rpInt += 10;
                bgWorkerLoad.ReportProgress(rpInt);
            }

            rpInt = 70;
            bgWorkerLoad.ReportProgress(rpInt);

            // 填入錯誤資料
            foreach (StudSubjectInfo ssi in StudSubjectLevelDuplicateList)
            {
                AddErrorSubjectInfoDict(ssi);
            }

            foreach (StudSubjectInfo ssi in hasErrorSubjectInfoDict.Values)
            {
                if (!GradeYearItems.Contains(ssi.ClassGradeYear))
                    GradeYearItems.Add(ssi.ClassGradeYear);
                if (!ClassNameItems.Contains(ssi.ClassName))
                    ClassNameItems.Add(ssi.ClassName);
                if (!SubjectNameItmes.Contains(ssi.SubjectName))
                    SubjectNameItmes.Add(ssi.SubjectName);
            }

            rpInt = 100;
            bgWorkerLoad.ReportProgress(rpInt);

        }

        private void AddErrorSubjectInfoDict(StudSubjectInfo ssi)
        {
            string key = ssi.SchoolYear + "_" + ssi.Semester + "_" + ssi.StudentID + "_" + ssi.SubjectName + "_" + ssi.SubjectLevel;
            if (!hasErrorSubjectInfoDict.ContainsKey(key))
                hasErrorSubjectInfoDict.Add(key, ssi);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            ControlEnable(false);
            comboGradeYear.Items.Clear();
            comboGradeYear.Text = "";
            comboClass.Items.Clear();
            comboClass.Text = "";
            comboSubject.Items.Clear();
            comboSubject.Text = "";
            dgData.Rows.Clear();
            lblMsg.Text = "共0筆";
            bgWorkerLoad.RunWorkerAsync();
        }

        private void btnUpdateSubjectLevel_Click(object sender, EventArgs e)
        {
            if (dgData.SelectedRows.Count < 1)
            {
                MsgBox.Show("沒有資料");
                return;
            }

            // 取得需要更新資料
            foreach (DataGridViewRow drv in dgData.SelectedRows)
            {
                StudSubjectInfo ssi = drv.Tag as StudSubjectInfo;
                if (ssi != null)
                    UpdateStudSubjectList.Add(ssi);
            }
        }

        private void frmSemsSubjectLevelDuplicate_Load(object sender, EventArgs e)
        {
            comboGradeYear.DropDownStyle = ComboBoxStyle.DropDownList;
            comboClass.DropDownStyle = ComboBoxStyle.DropDownList;
            comboSubject.DropDownStyle = ComboBoxStyle.DropDownList;

            // 載入欄位名稱
            LoadDataGridViewColumns();

            // 載入預設值
            SetFormDefaultLoadValue();
        }

        // 畫面預設值
        private void SetFormDefaultLoadValue()
        {
            btnUpdateSubjectLevel.Enabled = comboGradeYear.Enabled = comboClass.Enabled = comboSubject.Enabled = checkAll.Enabled = false;

            comboClass.Items.Clear();
            comboClass.Text = "";
            comboSubject.Items.Clear();
            comboSubject.Text = "";
            dgData.Rows.Clear();
            lblMsg.Text = "共0筆";
            checkAll.Checked = false;
        }

        private void checkAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgData.Rows)
            {
                row.Selected = checkAll.Checked;
            }
        }

        // 處理畫面年級、班級、科別選項
        private void LoadGradeYearClassNameAndSubjectName()
        {
            // 年級
            comboGradeYear.Items.Clear();
            if (GradeYearItems.Count > 0)
            {
                comboGradeYear.Items.Add("全部");
                comboGradeYear.Items.AddRange(GradeYearItems.ToArray());
                comboGradeYear.SelectedIndex = 0;
            }

            // 班級名稱
            comboClass.Items.Clear();
            if (ClassNameItems.Count > 0)
            {
                comboClass.Items.Add("全部");
                comboClass.Items.AddRange(ClassNameItems.ToArray());
                comboClass.SelectedIndex = 0;
            }

            // 科目名稱
            if (SubjectNameItmes.Count > 0)
            {
                comboSubject.Items.Add("全部");
                comboSubject.Items.AddRange(SubjectNameItmes.ToArray());
                comboSubject.SelectedIndex = 0;
            }

        }

        // 載入欄位名稱
        private void LoadDataGridViewColumns()
        {
            dgData.Columns.Clear();

            DataGridViewTextBoxColumn tbGradeYear = new DataGridViewTextBoxColumn();
            tbGradeYear.Name = "年級";
            tbGradeYear.Width = 30;
            tbGradeYear.HeaderText = "年級";
            tbGradeYear.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbGradeYear.ReadOnly = true;

            DataGridViewTextBoxColumn tbSchoolYear = new DataGridViewTextBoxColumn();
            tbSchoolYear.Name = "學年度";
            tbSchoolYear.Width = 40;
            tbSchoolYear.HeaderText = "學年度";
            tbSchoolYear.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSchoolYear.ReadOnly = true;

            DataGridViewTextBoxColumn tbSemester = new DataGridViewTextBoxColumn();
            tbSemester.Name = "學期";
            tbSemester.Width = 30;
            tbSemester.HeaderText = "學期";
            tbSemester.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSemester.ReadOnly = true;

            DataGridViewTextBoxColumn tbGGradeYear = new DataGridViewTextBoxColumn();
            tbGGradeYear.Name = "成績年級";
            tbGGradeYear.Width = 30;
            tbGGradeYear.HeaderText = "成績年級";
            tbGGradeYear.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbGGradeYear.ReadOnly = true;


            DataGridViewTextBoxColumn tbStudentNumber = new DataGridViewTextBoxColumn();
            tbStudentNumber.Name = "學號";
            tbStudentNumber.Width = 100;
            tbStudentNumber.HeaderText = "學號";
            tbStudentNumber.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbStudentNumber.ReadOnly = true;

            DataGridViewTextBoxColumn tbClassName = new DataGridViewTextBoxColumn();
            tbClassName.Name = "班級";
            tbClassName.Width = 80;
            tbClassName.HeaderText = "班級";
            tbClassName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbClassName.ReadOnly = true;

            DataGridViewTextBoxColumn tbSeatNo = new DataGridViewTextBoxColumn();
            tbSeatNo.Name = "座號";
            tbSeatNo.Width = 30;
            tbSeatNo.HeaderText = "座號";
            tbSeatNo.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSeatNo.ReadOnly = true;

            DataGridViewTextBoxColumn tbStudentName = new DataGridViewTextBoxColumn();
            tbStudentName.Name = "姓名";
            tbStudentName.Width = 100;
            tbStudentName.HeaderText = "姓名";
            tbStudentName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbStudentName.ReadOnly = true;

            DataGridViewTextBoxColumn tbSubjectName = new DataGridViewTextBoxColumn();
            tbSubjectName.Name = "科目";
            tbSubjectName.Width = 110;
            tbSubjectName.HeaderText = "科目";
            tbSubjectName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSubjectName.ReadOnly = true;

            DataGridViewTextBoxColumn tbSubjectLevel = new DataGridViewTextBoxColumn();
            tbSubjectLevel.Name = "科目級別";
            tbSubjectLevel.Width = 30;
            tbSubjectLevel.HeaderText = "科目級別";
            tbSubjectLevel.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSubjectLevel.ReadOnly = true;

            DataGridViewTextBoxColumn tbSubjectLevelNew = new DataGridViewTextBoxColumn();
            tbSubjectLevelNew.Name = "新科目級別";
            tbSubjectLevelNew.Width = 30;
            tbSubjectLevelNew.HeaderText = "新科目級別";
            tbSubjectLevelNew.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSubjectLevelNew.ReadOnly = true;

            DataGridViewTextBoxColumn tbRequiredBy = new DataGridViewTextBoxColumn();
            tbRequiredBy.Name = "校部定";
            tbRequiredBy.Width = 40;
            tbRequiredBy.HeaderText = "校部定";
            tbRequiredBy.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbRequiredBy.ReadOnly = true;

            DataGridViewTextBoxColumn tbRequired = new DataGridViewTextBoxColumn();
            tbRequired.Name = "必選修";
            tbRequired.Width = 40;
            tbRequired.HeaderText = "必選修";
            tbRequired.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbRequired.ReadOnly = true;

            DataGridViewTextBoxColumn tbCredit = new DataGridViewTextBoxColumn();
            tbCredit.Name = "學分";
            tbCredit.Width = 30;
            tbCredit.HeaderText = "學分";
            tbCredit.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbCredit.ReadOnly = true;


            dgData.Columns.Add(tbGradeYear);
            dgData.Columns.Add(tbSchoolYear);
            dgData.Columns.Add(tbSemester);
            dgData.Columns.Add(tbGGradeYear);
            dgData.Columns.Add(tbStudentNumber);
            dgData.Columns.Add(tbClassName);
            dgData.Columns.Add(tbSeatNo);
            dgData.Columns.Add(tbStudentName);
            dgData.Columns.Add(tbSubjectName);
            dgData.Columns.Add(tbSubjectLevel);
            dgData.Columns.Add(tbSubjectLevelNew);
            dgData.Columns.Add(tbRequiredBy);
            dgData.Columns.Add(tbRequired);
            dgData.Columns.Add(tbCredit);

        }

        private void dgData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboGradeYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedGradeYear = comboGradeYear.Text;
            // 填入資料
            DataToDataGridView();
        }

        private void comboClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedClassName = comboClass.Text;
            // 填入資料
            DataToDataGridView();
        }

        private void comboSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedSubjectName = comboSubject.Text;
            // 填入資料
            DataToDataGridView();
        }

        private void DataToDataGridView()
        {
            // 填入資料
            dgData.Rows.Clear();
            int rowCount = 0;
            foreach (StudSubjectInfo stud in hasErrorSubjectInfoDict.Values)
            {
                // 判斷是否新增
                bool isAdd = false;

                // 判斷班級
                if (SelectedGradeYear == "全部" && SelectedClassName == "全部" && SelectedSubjectName == "全部")
                    isAdd = true;
                else if (SelectedGradeYear == "全部" && SelectedClassName == "全部" && SelectedSubjectName != "全部")
                {
                    if (SelectedSubjectName == stud.SubjectName)
                        isAdd = true;
                    else
                        isAdd = false;
                }
                else if (SelectedGradeYear == "全部" && SelectedClassName != "全部" && SelectedSubjectName != "全部")
                {
                    if (SelectedClassName == stud.ClassName && SelectedSubjectName == stud.SubjectName)
                        isAdd = true;
                    else
                        isAdd = false;
                }
                else if (SelectedGradeYear == "全部" && SelectedClassName != "全部" && SelectedSubjectName == "全部")
                {
                    if (SelectedClassName == stud.ClassName)
                        isAdd = true;
                    else
                        isAdd = false;

                }
                else if (SelectedGradeYear != "全部" && SelectedClassName == "全部" && SelectedSubjectName == "全部")
                {
                    if (SelectedGradeYear == stud.ClassGradeYear)
                        isAdd = true;
                    else
                        isAdd = false;
                }
                else if (SelectedGradeYear != "全部" && SelectedClassName != "全部" && SelectedSubjectName != "全部")
                {
                    if (SelectedGradeYear == stud.ClassGradeYear && SelectedClassName == stud.ClassName && SelectedSubjectName == stud.SubjectName) isAdd = true;
                    else
                        isAdd = false;
                }


                if (isAdd)
                {
                    int rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Tag = stud;
                    dgData.Rows[rowIdx].Cells["年級"].Value = stud.ClassGradeYear;
                    dgData.Rows[rowIdx].Cells["學年度"].Value = stud.SchoolYear;
                    dgData.Rows[rowIdx].Cells["學期"].Value = stud.Semester;
                    dgData.Rows[rowIdx].Cells["成績年級"].Value = stud.GradeYear;
                    dgData.Rows[rowIdx].Cells["學號"].Value = stud.StudentNumber;
                    dgData.Rows[rowIdx].Cells["班級"].Value = stud.ClassName;
                    dgData.Rows[rowIdx].Cells["座號"].Value = stud.SeatNo;
                    dgData.Rows[rowIdx].Cells["姓名"].Value = stud.Name;
                    dgData.Rows[rowIdx].Cells["科目"].Value = stud.SubjectName;
                    dgData.Rows[rowIdx].Cells["科目級別"].Value = stud.SubjectLevel;

                    if (stud.RequiredBy == "部訂")
                        dgData.Rows[rowIdx].Cells["校部定"].Value = "部定";
                    else
                        dgData.Rows[rowIdx].Cells["校部定"].Value = stud.RequiredBy;

                    dgData.Rows[rowIdx].Cells["必選修"].Value = stud.Required;

                    dgData.Rows[rowIdx].Cells["學分"].Value = stud.Credit;

                    rowCount++;
                }

            }

            lblMsg.Text = "共" + rowCount + "筆";
        }

        private void dgData_SelectionChanged(object sender, EventArgs e)
        {
            lblMsg.Text = "共" + dgData.Rows.Count + "筆，已選" + dgData.SelectedRows.Count + "筆。";
        }
    }
}
