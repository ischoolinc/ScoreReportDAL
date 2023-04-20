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
using SHSemsSubjectCheckEdit.DAO;

namespace SHSemsSubjectCheckEdit.UIForm
{
    public partial class frmSemsSubjectNameCheckEdit : BaseForm
    {
        BackgroundWorker bgWorker;
        string SchoolYear = "";
        // 依學年度取得所有學生學期科目成績
        List<StudSubjectInfo> StudSubjectInfoList;        

        // 整理有問題資料使用
        Dictionary<string, StudCheckDataInfo> StudCheckDataDict;

        // 課程規劃表對照
        Dictionary<string, GPlanInfo> GPlanDict;

        List<string> ClassNameItems;
        List<string> SubjectNameItmes;
        string SelectedClassName = "全部";
        string SelectedSubjectName = "全部";

        // 檢查有問題資料
        //List<StudSubjectInfo> hasErrorSubjectInfoList;
        Dictionary<string, StudSubjectInfo> hasErrorSubjectInfoDict;

        // 刪除資料
        List<StudSubjectInfo> DelStudSubjectList;

        // 更新資料
        List<StudSubjectInfo> UpdateStudSubjectList;

        public frmSemsSubjectNameCheckEdit()
        {
            InitializeComponent();
            bgWorker = new BackgroundWorker();
            StudSubjectInfoList = new List<StudSubjectInfo>();
            StudCheckDataDict = new Dictionary<string, StudCheckDataInfo>();
            //hasErrorSubjectInfoList = new List<StudSubjectInfo>();
            hasErrorSubjectInfoDict = new Dictionary<string, StudSubjectInfo>();
            GPlanDict = new Dictionary<string, GPlanInfo>();
            ClassNameItems = new List<string>();
            SubjectNameItmes = new List<string>();
            DelStudSubjectList = new List<StudSubjectInfo>();
            UpdateStudSubjectList = new List<StudSubjectInfo>();

            bgWorker.DoWork += BgWorker_DoWork;
            bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
            bgWorker.ProgressChanged += BgWorker_ProgressChanged;
            bgWorker.WorkerReportsProgress = true;
        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期科目成績查詢中...", e.ProgressPercentage);
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            // 載入畫面
            if (hasErrorSubjectInfoDict.Values.Count > 0)
            {
                // 填入資料
                DataToDataGridView();

                // 填入選項
                LoadClassNameAndSubjectName();
            }
            ControlEnable(true);
        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            StudCheckDataDict.Clear();
            hasErrorSubjectInfoDict.Clear();
            ClassNameItems.Clear();
            SubjectNameItmes.Clear();
            StudSubjectInfoList.Clear();
            GPlanDict.Clear();

            bgWorker.ReportProgress(rpInt);
            // 依畫面選學年度取得資料
            for (int gr = 1; gr <= 3; gr++)
            {
                StudSubjectInfoList.AddRange(DataAccess.GetSemsSubjectBySchoolYear(SchoolYear, gr));
                rpInt += 10;
                bgWorker.ReportProgress(rpInt);
            }

            bgWorker.ReportProgress(rpInt);
            foreach (StudSubjectInfo ssi in StudSubjectInfoList)
            {
                if (!StudCheckDataDict.ContainsKey(ssi.StudentID))
                {
                    StudCheckDataInfo ss = new StudCheckDataInfo();
                    ss.StudentID = ssi.StudentID;
                    StudCheckDataDict.Add(ssi.StudentID, ss);
                }

                if (!StudCheckDataDict[ssi.StudentID].SubjectNameList.Contains(ssi.SubjectName))
                    StudCheckDataDict[ssi.StudentID].SubjectNameList.Add(ssi.SubjectName);

                // 如果有設定指定學年科目名稱，不檢查
                if (!string.IsNullOrEmpty(ssi.SYSubjectName))
                    continue;

                // 檢查同一學期科目名稱相同
                if (ssi.Semester == "1")
                {
                    //檢查相同科目名稱1
                    if (!StudCheckDataDict[ssi.StudentID].SameSubjectNameDict1.ContainsKey(ssi.SubjectName))
                    {
                        StudCheckDataDict[ssi.StudentID].SameSubjectNameDict1.Add(ssi.SubjectName, new List<StudSubjectInfo>());
                    }
                    StudCheckDataDict[ssi.StudentID].SameSubjectNameDict1[ssi.SubjectName].Add(ssi);

                    // 檢查科目名稱相同，上下學期都有使用
                    string subjKey = ssi.SubjectName + "_" + ssi.Required + "_" + ssi.RequiredBy + "_" + ssi.Credit;

                    if (!StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict1.ContainsKey(ssi.SubjectName))
                        StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict1.Add(ssi.SubjectName, new Dictionary<string, List<StudSubjectInfo>>());

                    if (!StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict1[ssi.SubjectName].ContainsKey(subjKey))
                        StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict1[ssi.SubjectName].Add(subjKey, new List<StudSubjectInfo>());

                    StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict1[ssi.SubjectName][subjKey].Add(ssi);
                }

                if (ssi.Semester == "2")
                {
                    //檢查相同科目名稱2
                    if (!StudCheckDataDict[ssi.StudentID].SameSubjectNameDict2.ContainsKey(ssi.SubjectName))
                    {
                        StudCheckDataDict[ssi.StudentID].SameSubjectNameDict2.Add(ssi.SubjectName, new List<StudSubjectInfo>());
                    }
                    StudCheckDataDict[ssi.StudentID].SameSubjectNameDict2[ssi.SubjectName].Add(ssi);

                    // 檢查科目名稱相同，上下學期都有使用
                    string subjKey = ssi.SubjectName + "_" + ssi.Required + "_" + ssi.RequiredBy + "_" + ssi.Credit;

                    if (!StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict2.ContainsKey(ssi.SubjectName))
                        StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict2.Add(ssi.SubjectName, new Dictionary<string, List<StudSubjectInfo>>());

                    if (!StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict2[ssi.SubjectName].ContainsKey(subjKey))
                        StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict2[ssi.SubjectName].Add(subjKey, new List<StudSubjectInfo>());

                    StudCheckDataDict[ssi.StudentID].StudSubjectInfoDict2[ssi.SubjectName][subjKey].Add(ssi);
                }

            }
            bgWorker.ReportProgress(rpInt);

            // 檢查資料
            foreach (StudCheckDataInfo scd in StudCheckDataDict.Values)
            {
                foreach (string key in scd.SameSubjectNameDict1.Keys)
                {
                    if (scd.SameSubjectNameDict1[key].Count > 1)
                    {
                        foreach (StudSubjectInfo ssi in scd.SameSubjectNameDict1[key])
                            AddErrorSubjectInfoDict(ssi);
                    }
                }

                foreach (string key in scd.SameSubjectNameDict2.Keys)
                {
                    if (scd.SameSubjectNameDict2[key].Count > 1)
                        foreach (StudSubjectInfo ssi in scd.SameSubjectNameDict2[key])
                            AddErrorSubjectInfoDict(ssi);
                }

                // 檢查上下學期都有，科目名稱相同，但是校部定、必選修、學分數有不同

                foreach (string subjName in scd.SubjectNameList)
                {
                    if (scd.StudSubjectInfoDict1.ContainsKey(subjName) && scd.SameSubjectNameDict2.ContainsKey(subjName))
                    {
                        foreach (string subjKey in scd.StudSubjectInfoDict1[subjName].Keys)
                        {
                            if (!scd.StudSubjectInfoDict2[subjName].ContainsKey(subjKey))
                            {
                                // 可能有問題填入
                                foreach (StudSubjectInfo ssi1 in scd.StudSubjectInfoDict1[subjName][subjKey])
                                {
                                    AddErrorSubjectInfoDict(ssi1);
                                }

                                foreach (string subjKey2 in scd.StudSubjectInfoDict2[subjName].Keys)
                                {
                                    foreach (StudSubjectInfo ssi2 in scd.StudSubjectInfoDict2[subjName][subjKey2])
                                    {
                                        AddErrorSubjectInfoDict(ssi2);
                                    }
                                }

                            }
                        }
                    }
                }
            }
            rpInt = 70;
            bgWorker.ReportProgress(rpInt);

            List<string> gpids = new List<string>();
            // 取得使用課程規畫表對照
            foreach (StudSubjectInfo ssi in hasErrorSubjectInfoDict.Values)
            {
                if (!gpids.Contains(ssi.GPID))
                    gpids.Add(ssi.GPID);
            }
            GPlanDict = DataAccess.GetGPlanDictByIDs(gpids);

            rpInt = 90;
            bgWorker.ReportProgress(rpInt);
            // 填入資料
            foreach (StudSubjectInfo ssi in hasErrorSubjectInfoDict.Values)
            {
                // 填入課規資料
                if (GPlanDict.ContainsKey(ssi.GPID))
                {
                    string subjKey = ssi.SubjectName + "_" + ssi.SubjectLevel;
                    if (GPlanDict[ssi.GPID].SubjectsDict.ContainsKey(subjKey))
                    {
                        ssi.GPRequired = GPlanDict[ssi.GPID].SubjectsDict[subjKey].Required;
                        ssi.GPRequiredBy = GPlanDict[ssi.GPID].SubjectsDict[subjKey].RequiredBy;
                        ssi.GPCredit = GPlanDict[ssi.GPID].SubjectsDict[subjKey].Credit;
                        ssi.GPSYSubjectName = GPlanDict[ssi.GPID].SubjectsDict[subjKey].SubjectNameByYear;

                    }                    
                }

                if (!ClassNameItems.Contains(ssi.ClassName))
                    ClassNameItems.Add(ssi.ClassName);

                if (!SubjectNameItmes.Contains(ssi.SubjectName))
                    SubjectNameItmes.Add(ssi.SubjectName);
            }

            rpInt = 100;
            bgWorker.ReportProgress(rpInt);
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
            comboClass.Items.Clear();
            comboClass.Text = "";
            comboSubject.Items.Clear();
            comboSubject.Text = "";
            dgData.Rows.Clear();
            SchoolYear = cboSchoolYear.Text;
            lblMsg.Text = "共0筆";
            bgWorker.RunWorkerAsync();
        }

        private void ControlEnable(bool value)
        {
            btnQuery.Enabled = btnUpdateFromGPlan.Enabled = btnDel.Enabled = cboSchoolYear.Enabled = checkAll.Enabled = comboClass.Enabled = comboSubject.Enabled = value;
        }

        private void frmSemsSubjectNameCheckEdit_Load(object sender, EventArgs e)
        {
            // 載入欄位名稱
            LoadDataGridViewColumns();

            // 載入學期成績學年度
            LoadSemsScoreSchoolYear();
            cboSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;

            // 載入預設值
            SetFormDefaultLoadValue();

        }

        // 載入學期成績學年度
        private void LoadSemsScoreSchoolYear()
        {
            cboSchoolYear.Items.Clear();
            cboSchoolYear.Items.AddRange(DataAccess.GetSemsScoreSchoolYear().ToArray());
        }

        private void SetFormDefaultLoadValue()
        {
            btnUpdateFromGPlan.Enabled = btnDel.Enabled = comboClass.Enabled = comboSubject.Enabled = checkAll.Enabled = false;

            comboClass.Items.Clear();
            comboClass.Text = "";
            comboSubject.Items.Clear();
            comboSubject.Text = "";
            dgData.Rows.Clear();
            lblMsg.Text = "共0筆";
            checkAll.Checked = false;
        }

        // 載入欄位名稱
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

            DataGridViewTextBoxColumn tbSYSubjectName = new DataGridViewTextBoxColumn();
            tbSYSubjectName.Name = "指定學年科目名稱";
            tbSYSubjectName.Width = 150;
            tbSYSubjectName.HeaderText = "指定學年科目名稱";
            tbSYSubjectName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSYSubjectName.ReadOnly = true;

            DataGridViewTextBoxColumn tbRequiredBy1 = new DataGridViewTextBoxColumn();
            tbRequiredBy1.Name = "課規校部定";
            tbRequiredBy1.Width = 40;
            tbRequiredBy1.HeaderText = "課規校部定";
            tbRequiredBy1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbRequiredBy1.ReadOnly = true;

            DataGridViewTextBoxColumn tbRequired1 = new DataGridViewTextBoxColumn();
            tbRequired1.Name = "課規必選修";
            tbRequired1.Width = 40;
            tbRequired1.HeaderText = "課規必選修";
            tbRequired1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbRequired1.ReadOnly = true;

            DataGridViewTextBoxColumn tbCredit1 = new DataGridViewTextBoxColumn();
            tbCredit1.Name = "課規學分";
            tbCredit1.Width = 30;
            tbCredit1.HeaderText = "課規學分";
            tbCredit1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbCredit1.ReadOnly = true;

            DataGridViewTextBoxColumn tbSYSubjectName1 = new DataGridViewTextBoxColumn();
            tbSYSubjectName1.Name = "課規指定學年科目名稱";
            tbSYSubjectName1.Width = 150;
            tbSYSubjectName1.HeaderText = "課規指定學年科目名稱";
            tbSYSubjectName1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tbSYSubjectName1.ReadOnly = true;

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
            dgData.Columns.Add(tbRequiredBy1);
            
            dgData.Columns.Add(tbRequired);
            dgData.Columns.Add(tbRequired1);
            
            dgData.Columns.Add(tbCredit);
            dgData.Columns.Add(tbCredit1);
            
            dgData.Columns.Add(tbSYSubjectName);            
            dgData.Columns.Add(tbSYSubjectName1);
            
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            DelStudSubjectList.Clear();
            if (dgData.SelectedRows.Count < 1)
            {
                MsgBox.Show("沒有資料");
                return;
            }

            // 取得需要刪除資料
            foreach(DataGridViewRow drv in dgData.SelectedRows)
            {
                StudSubjectInfo ssi = drv.Tag as StudSubjectInfo;
                if (ssi != null)
                    DelStudSubjectList.Add(ssi);
            }
        }

        private void btnUpdateFromGPlan_Click(object sender, EventArgs e)
        {
            if (dgData.SelectedRows.Count < 1)
            {
                MsgBox.Show("沒有資料");
                return;
            }
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
                if (SelectedClassName == "全部" && SelectedSubjectName == "全部")
                    isAdd = true;
                else if (SelectedClassName == "全部" && SelectedSubjectName != "全部")
                {
                    if (SelectedSubjectName == stud.SubjectName)
                        isAdd = true;
                    else
                        isAdd = false;
                }
                else if (SelectedClassName != "全部" && SelectedSubjectName == "全部")
                {
                    if (SelectedClassName == stud.ClassName)
                        isAdd = true;
                    else
                        isAdd = false;

                }
                else if (SelectedClassName != "全部" && SelectedSubjectName != "全部")
                {
                    if (SelectedClassName == stud.ClassName && SelectedSubjectName == stud.SubjectName) isAdd = true;
                    else
                        isAdd = false;
                }


                if (isAdd)
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

                    if (stud.RequiredBy != stud.GPRequiredBy)
                        dgData.Rows[rowIdx].Cells["校部定"].Style.BackColor = Color.Yellow;

                    dgData.Rows[rowIdx].Cells["必選修"].Value = stud.Required;
                    if (stud.Required != stud.GPRequired)
                        dgData.Rows[rowIdx].Cells["必選修"].Style.BackColor = Color.Yellow;

                    dgData.Rows[rowIdx].Cells["學分"].Value = stud.Credit;
                    if (stud.Credit != stud.GPCredit)
                        dgData.Rows[rowIdx].Cells["學分"].Style.BackColor = Color.Yellow;

                    dgData.Rows[rowIdx].Cells["指定學年科目名稱"].Value = stud.SYSubjectName;
                    if (stud.SYSubjectName != stud.GPSYSubjectName)
                        dgData.Rows[rowIdx].Cells["指定學年科目名稱"].Style.BackColor = Color.Yellow;

                    if (stud.GPRequiredBy == "部訂")
                        dgData.Rows[rowIdx].Cells["課規校部定"].Value = "部定";
                    else
                        dgData.Rows[rowIdx].Cells["課規校部定"].Value = stud.GPRequiredBy;
                    dgData.Rows[rowIdx].Cells["課規必選修"].Value = stud.GPRequired;
                    dgData.Rows[rowIdx].Cells["課規學分"].Value = stud.GPCredit;
                    dgData.Rows[rowIdx].Cells["課規指定學年科目名稱"].Value = stud.GPSYSubjectName;

                    rowCount++;
                }

            }

            lblMsg.Text = "共" + rowCount + "筆";
        }

        private void checkAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgData.Rows)
            {
                row.Selected = checkAll.Checked;
            }
        }

        // 處理畫面班級與科目選項
        private void LoadClassNameAndSubjectName()
        {
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

        private void dgData_SelectionChanged(object sender, EventArgs e)
        {
            lblMsg.Text = "共" + dgData.Rows.Count + "筆，已選" + dgData.SelectedRows.Count + "筆。";
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SchoolYear != cboSchoolYear.Text)
                SetFormDefaultLoadValue();

            SchoolYear = cboSchoolYear.Text;
        }
    }
}
