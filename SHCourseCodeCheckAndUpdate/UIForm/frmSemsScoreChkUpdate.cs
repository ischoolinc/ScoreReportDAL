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
using SHCourseCodeCheckAndUpdate.DAO;

namespace SHCourseCodeCheckAndUpdate.UIForm
{
    public partial class frmSemsScoreChkUpdate : BaseForm
    {
        BackgroundWorker bgWorkerLoad;
        BackgroundWorker bgWorkerUpdate;
        List<StudSubjectScoreInfo> StudSubjectScoreInfoList;
        Dictionary<string, List<StudSubjectScoreInfo>> UpdateSemsScoreDict;
        string SchoolYear = "", Semester = "", GradeYear = "";

        int UpdateCount = 0;


        public frmSemsScoreChkUpdate()
        {
            InitializeComponent();
            bgWorkerLoad = new BackgroundWorker();
            bgWorkerLoad.DoWork += BgWorkerLoad_DoWork;
            bgWorkerLoad.RunWorkerCompleted += BgWorkerLoad_RunWorkerCompleted;
            bgWorkerLoad.ProgressChanged += BgWorkerLoad_ProgressChanged;
            bgWorkerLoad.WorkerReportsProgress = true;
            bgWorkerUpdate = new BackgroundWorker();
            bgWorkerUpdate.DoWork += BgWorkerUpdate_DoWork;
            bgWorkerUpdate.RunWorkerCompleted += BgWorkerUpdate_RunWorkerCompleted;
            bgWorkerUpdate.ProgressChanged += BgWorkerUpdate_ProgressChanged;
            bgWorkerUpdate.WorkerReportsProgress = true;
            StudSubjectScoreInfoList = new List<StudSubjectScoreInfo>();
            UpdateSemsScoreDict = new Dictionary<string, List<StudSubjectScoreInfo>>();
        }

        private void BgWorkerUpdate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("更新學期成績課程代碼中...", e.ProgressPercentage);
        }

        private void BgWorkerUpdate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MsgBox.Show("更新" + UpdateCount + "筆");
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            ClearTempData();
        }

        private void ClearTempData()
        {
            StudSubjectScoreInfoList.Clear();
            UpdateSemsScoreDict.Clear();
            dgData.Rows.Clear();
            lblMsg.Text = "共0筆";
            btnUpdate.Enabled = false;
        }

        private void BgWorkerUpdate_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerUpdate.ReportProgress(5);

            UpdateSemsScoreDict.Clear();
            // 依學期成績系統編號，整理需要更新資料，因為成績是一包XML組成
            foreach(StudSubjectScoreInfo ss in StudSubjectScoreInfoList)
            {
                if (!UpdateSemsScoreDict.ContainsKey(ss.SemsSubjID))
                    UpdateSemsScoreDict.Add(ss.SemsSubjID, new List<StudSubjectScoreInfo>());

                UpdateSemsScoreDict[ss.SemsSubjID].Add(ss);
            }

            // 更新資料
            UpdateCount = DataAccess.UpdateSubjectScoreCourseCode(UpdateSemsScoreDict);

            bgWorkerUpdate.ReportProgress(100);
        }

        private void BgWorkerLoad_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績課程代碼檢查中...", e.ProgressPercentage);
        }

        private void BgWorkerLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DataToDataGridView(StudSubjectScoreInfoList);

            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            ControlEnable(true);
        }

        private void DataToDataGridView(List<StudSubjectScoreInfo> dataList)
        {
            // 填入資料
            dgData.Rows.Clear();
            int rowCount = 0;
            foreach (StudSubjectScoreInfo stud in dataList)
            {
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = stud;
                dgData.Rows[rowIdx].Cells["學年度"].Value = stud.SchoolYear;
                dgData.Rows[rowIdx].Cells["學期"].Value = stud.Semester;
                dgData.Rows[rowIdx].Cells["年級"].Value = stud.GradeYear;
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
                dgData.Rows[rowIdx].Cells["學期課程代碼"].Value = stud.SS_CourseCode;
                dgData.Rows[rowIdx].Cells["課規課程代碼"].Value = stud.GP_CourseCode;
                dgData.Rows[rowIdx].Cells["使用課規"].Value = stud.GPName;
                dgData.Rows[rowIdx].Cells["學生狀態"].Value = stud.Status;
                rowCount++;
            }

            lblMsg.Text = "共" + rowCount + "筆";
        }

        private void BgWorkerLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerLoad.ReportProgress(5);
            // 取得學生學期成績與學生課程規劃比對後，有差異的資料。
            StudSubjectScoreInfoList = DataAccess.GetSubjectScoreBySchoolYearSems(SchoolYear, Semester, GradeYear);
            bgWorkerLoad.ReportProgress(100);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            QueryData();
        }

        private void SetUIDefault()
        {
            // 學年度
            int sc = 0;
            int.TryParse(K12.Data.School.DefaultSchoolYear, out sc);

            for (int i = sc - 2; i <= sc; i++)
            {
                cbxSchoolYear.Items.Add(i + "");
            }

            cbxSchoolYear.Text = sc + "";
            cbxSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;

            // 學期
            cbxSemester.Items.Add("1");
            cbxSemester.Items.Add("2");
            cbxSemester.Text = K12.Data.School.DefaultSemester;
            cbxSemester.DropDownStyle = ComboBoxStyle.DropDownList;

            // 年級
            for (int i = 1; i <= 3; i++)
                cbxGradeYear.Items.Add(i + "");

            cbxGradeYear.Text = "1";
            cbxGradeYear.DropDownStyle = ComboBoxStyle.DropDownList;

            lblMsg.Text = "共0筆";
            btnUpdate.Enabled = false;
        }

        private void QueryData()
        {
            // 畫面上無法變更
            ControlEnable(false);

            // 取得畫面資料
            SchoolYear = cbxSchoolYear.Text;
            Semester = cbxSemester.Text;
            GradeYear = cbxGradeYear.Text;

            // 執行
            bgWorkerLoad.RunWorkerAsync();
        }

        private void frmSemsScoreChkUpdate_Load(object sender, EventArgs e)
        {
            // 畫面設定
            SetUIDefault();

            // 取得欄位
            LoadDataGridViewColumns();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            UpdateData();
        }

        private void LoadDataGridViewColumns()
        {
            dgData.Columns.Clear();
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

            DataGridViewTextBoxColumn tbGradeYear = new DataGridViewTextBoxColumn();
            tbGradeYear.Name = "年級";
            tbGradeYear.Width = 30;
            tbGradeYear.HeaderText = "年級";
            tbGradeYear.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbGradeYear.ReadOnly = true;

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

            DataGridViewTextBoxColumn tbCoCourseCode = new DataGridViewTextBoxColumn();
            tbCoCourseCode.Name = "學期課程代碼";
            tbCoCourseCode.Width = 210;
            tbCoCourseCode.HeaderText = "學期課程代碼";
            tbCoCourseCode.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbCoCourseCode.ReadOnly = true;

            DataGridViewTextBoxColumn tbGpCourseCode = new DataGridViewTextBoxColumn();
            tbGpCourseCode.Name = "課規課程代碼";
            tbGpCourseCode.Width = 210;
            tbGpCourseCode.HeaderText = "課規課程代碼";
            tbGpCourseCode.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbGpCourseCode.ReadOnly = true;

            DataGridViewTextBoxColumn tbUseGpName = new DataGridViewTextBoxColumn();
            tbUseGpName.Name = "使用課規";
            tbUseGpName.Width = 200;
            tbClassName.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            tbUseGpName.HeaderText = "使用課規";
            tbUseGpName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbUseGpName.ReadOnly = true;

            DataGridViewTextBoxColumn tbStatus = new DataGridViewTextBoxColumn();
            tbStatus.Name = "學生狀態";
            tbStatus.Width = 80;
            tbStatus.HeaderText = "學生狀態";
            tbStatus.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbStatus.ReadOnly = true;


            dgData.Columns.Add(tbSchoolYear);
            dgData.Columns.Add(tbSemester);
            dgData.Columns.Add(tbGradeYear);
            dgData.Columns.Add(tbStudentNumber);
            dgData.Columns.Add(tbClassName);
            dgData.Columns.Add(tbSeatNo);
            dgData.Columns.Add(tbStudentName);
            dgData.Columns.Add(tbSubjectName);
            dgData.Columns.Add(tbSubjectLevel);
            dgData.Columns.Add(tbRequiredBy);
            dgData.Columns.Add(tbRequired);
            dgData.Columns.Add(tbCredit);
            dgData.Columns.Add(tbCoCourseCode);
            dgData.Columns.Add(tbGpCourseCode);
            dgData.Columns.Add(tbUseGpName);
            dgData.Columns.Add(tbStatus);

        }

        private void ControlEnable(bool value)
        {
            cbxSchoolYear.Enabled = cbxSemester.Enabled = cbxGradeYear.Enabled = btnQuery.Enabled = btnUpdate.Enabled = dgData.Enabled = value;
        }

        private void UpdateData()
        {

            if (dgData.Rows.Count == 0)
            {
                MsgBox.Show("沒有資料，無法更新。");
                return;
            }

            ControlEnable(false);

            bgWorkerUpdate.RunWorkerAsync();

        }
    }
}
