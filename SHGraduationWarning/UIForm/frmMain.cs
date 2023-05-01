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
using DevComponents.DotNetBar.Controls;
using FISCA.Presentation.Controls;
using SHGraduationWarning.DAO;
using System.Web.Script.Serialization;
using System.Net.NetworkInformation;

namespace SHGraduationWarning.UIForm
{
    public partial class frmMain : BaseForm
    {
        string AllStr = "全部";
        // 載入預設資料
        BackgroundWorker bgWorkerLoadDefault;
        // 載入畢業資格檢查資料
        BackgroundWorker bgwDataGWLoad;
        // 載入資料合理
        BackgroundWorker bgwDataChkEditLoad;

        List<ClassDeptInfo> ClassDeptInfoList;
        string SelectedDeptName = "";
        string SelectedClassName = "";
        Dictionary<string, List<string>> DeptNameDict;

        int TabControlSelectedIndex = 0;

        List<StudSubjectInfo> StudSubjectInfoList;

        // 課程規劃表對照
        Dictionary<string, GPlanInfo> GPlanDict;

        // 檢查有問題資料        
        Dictionary<string, StudSubjectInfo> hasErrorSubjectInfoDict;

        public frmMain()
        {
            GPlanDict = new Dictionary<string, GPlanInfo>();
            hasErrorSubjectInfoDict = new Dictionary<string, StudSubjectInfo>();
            bgWorkerLoadDefault = new BackgroundWorker();
            bgWorkerLoadDefault.DoWork += BgWorkerLoadDefault_DoWork;
            bgWorkerLoadDefault.RunWorkerCompleted += BgWorkerLoadDefault_RunWorkerCompleted;
            bgWorkerLoadDefault.ProgressChanged += BgWorkerLoadDefault_ProgressChanged;
            bgWorkerLoadDefault.WorkerReportsProgress = true;

            bgwDataGWLoad = new BackgroundWorker();
            bgwDataGWLoad.DoWork += BgwDataGWLoad_DoWork;
            bgwDataGWLoad.RunWorkerCompleted += BgwDataGWLoad_RunWorkerCompleted;
            bgwDataGWLoad.ProgressChanged += BgwDataGWLoad_ProgressChanged;
            bgwDataGWLoad.WorkerReportsProgress = true;

            bgwDataChkEditLoad = new BackgroundWorker();
            bgwDataChkEditLoad.DoWork += BgwDataChkEditLoad_DoWork;
            bgwDataChkEditLoad.RunWorkerCompleted += BgwDataChkEditLoad_RunWorkerCompleted;
            bgwDataChkEditLoad.ProgressChanged += BgwDataChkEditLoad_ProgressChanged;
            bgwDataChkEditLoad.WorkerReportsProgress = true;

            StudSubjectInfoList = new List<StudSubjectInfo>();

            DeptNameDict = new Dictionary<string, List<string>>();
            ClassDeptInfoList = new List<ClassDeptInfo>(); ;
            SelectedDeptName = SelectedClassName = AllStr;

            InitializeComponent();
        }

        private void BgwDataChkEditLoad_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("資料合理檢查中...", e.ProgressPercentage);
        }

        private void BgwDataChkEditLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");

            if (StudSubjectInfoList.Count > 0)
            {
                dgDataChkEdit.Rows.Clear();
                foreach (StudSubjectInfo ss in StudSubjectInfoList)
                {
                    int rowIdx = dgDataChkEdit.Rows.Add();
                    dgDataChkEdit.Rows[rowIdx].Tag = ss;
                    dgDataChkEdit.Rows[rowIdx].Cells["學號"].Value = ss.StudentNumber;
                    dgDataChkEdit.Rows[rowIdx].Cells["班級"].Value = ss.ClassName;
                    dgDataChkEdit.Rows[rowIdx].Cells["座號"].Value = ss.SeatNo;
                    dgDataChkEdit.Rows[rowIdx].Cells["科別"].Value = ss.DeptName;
                    dgDataChkEdit.Rows[rowIdx].Cells["姓名"].Value = ss.Name;
                    dgDataChkEdit.Rows[rowIdx].Cells["學年度"].Value = ss.SchoolYear;
                    dgDataChkEdit.Rows[rowIdx].Cells["學期"].Value = ss.Semester;
                    dgDataChkEdit.Rows[rowIdx].Cells["成績年級"].Value = ss.GradeYear;
                    dgDataChkEdit.Rows[rowIdx].Cells["科目名稱"].Value = ss.SubjectName;
                    dgDataChkEdit.Rows[rowIdx].Cells["科目級別"].Value = ss.SubjectLevel;
                    if (ss.IsSubjectLevelChanged)
                        dgDataChkEdit.Rows[rowIdx].Cells["科目級別"].Style.BackColor = Color.Yellow;
                    else
                        dgDataChkEdit.Rows[rowIdx].Cells["科目級別"].Style.BackColor = Color.White;
                    dgDataChkEdit.Rows[rowIdx].Cells["建議級別"].Value = ss.SubjectLevelNew;
                    dgDataChkEdit.Rows[rowIdx].Cells["分項"].Value = ss.Entry;
                    dgDataChkEdit.Rows[rowIdx].Cells["領域"].Value = ss.Domain;
                    dgDataChkEdit.Rows[rowIdx].Cells["學分"].Value = ss.Credit;
                    dgDataChkEdit.Rows[rowIdx].Cells["必選修"].Value = ss.Required;
                    dgDataChkEdit.Rows[rowIdx].Cells["校部定"].Value = ss.RequiredBy;
                    dgDataChkEdit.Rows[rowIdx].Cells["使用課規"].Value = ss.GPName;
                    dgDataChkEdit.Rows[rowIdx].Cells["指定學年科目名稱"].Value = ss.SchoolYearSubjectName;
                    dgDataChkEdit.Rows[rowIdx].Cells["問題說明"].Value = string.Join(",", ss.ErrorMsgList.ToArray());

                }
            }

            lblMsg.Text = "共" + dgDataChkEdit.Rows.Count + "筆";
            ControlEnable(true);
        }

        private void BgwDataChkEditLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            StudSubjectInfoList.Clear();
            GPlanDict.Clear();

            bgwDataChkEditLoad.ReportProgress(rpInt);
            for (int i = 1; i <= 3; i++)
                StudSubjectInfoList.AddRange(DataAccess.GetSemsSubjectLevelDuplicateByGradeYear(i));

            foreach (StudSubjectInfo ss in StudSubjectInfoList)
            {
                AddErrorSubjectInfoDict(ss);
            }

            rpInt = 70;
            bgwDataChkEditLoad.ReportProgress(rpInt);

            List<string> gpids = new List<string>();
            // 取得使用課程規畫表對照
            foreach (StudSubjectInfo ssi in hasErrorSubjectInfoDict.Values)
            {
                if (!gpids.Contains(ssi.CoursePlanID))
                    gpids.Add(ssi.CoursePlanID);
            }
            GPlanDict = DataAccess.GetGPlanDictByIDs(gpids);

            // 填入資料
            foreach (StudSubjectInfo ssi in hasErrorSubjectInfoDict.Values)
            {
                ssi.IsSubjectLevelChanged = true;
                // 填入課規資料
                if (GPlanDict.ContainsKey(ssi.CoursePlanID))
                {
                    ssi.GPName = GPlanDict[ssi.CoursePlanID].Name;
                    // 使用年級+學期+科目+級別當key
                    string subjKey = ssi.GradeYear + "_" + ssi.Semester + "_" + ssi.SubjectName;
                    if (GPlanDict[ssi.CoursePlanID].SubjectsDict.ContainsKey(subjKey))
                    {                        
                        ssi.GPRequired = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].Required;
                        ssi.GPRequiredBy = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].RequiredBy;
                        ssi.GPCredit = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].Credit;
                        ssi.SubjectLevelNew = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].SubjectLevel;
                        ssi.GPSYSubjectName = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].SubjectNameByYear;
                        ssi.GPDomain = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].Domain;
                        ssi.GPEntry = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].Entry;
                        ssi.GPCourseCode = GPlanDict[ssi.CoursePlanID].SubjectsDict[subjKey].CourseCode;

                        if (ssi.SubjectLevel == ssi.SubjectLevelNew)
                            ssi.IsSubjectLevelChanged = false;
                    }
                }

            }


            rpInt = 100;
            bgwDataChkEditLoad.ReportProgress(rpInt);

        }

        private void BgwDataGWLoad_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業預計檢查中...", e.ProgressPercentage);
        }

        private void BgwDataGWLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void BgwDataGWLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void BgWorkerLoadDefault_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("預設資料載入中...", e.ProgressPercentage);
        }

        private void BgWorkerLoadDefault_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            // 載入班級科別資訊
            LoadClassDeptToForm();
            ControlEnable(true);
        }

        private void BgWorkerLoadDefault_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            bgWorkerLoadDefault.ReportProgress(rpInt);

            // 讀取班級科別資訊            
            ClassDeptInfoList = DataAccess.GetClassDeptList();
            DeptNameDict.Clear();
            foreach (ClassDeptInfo cd in ClassDeptInfoList)
            {
                if (!DeptNameDict.ContainsKey(cd.DeptName))
                    DeptNameDict.Add(cd.DeptName, new List<string>());

                DeptNameDict[cd.DeptName].Add(cd.ClassName);
            }

            rpInt = 10;
            bgWorkerLoadDefault.ReportProgress(rpInt);


            rpInt = 100;
            bgWorkerLoadDefault.ReportProgress(rpInt);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //清除班級科別資訊
        private void ClearClassDept()
        {
            ClassDeptInfoList.Clear();
            comboDept.Text = "";
            comboDept.Items.Clear();
            comboClass.Text = "";
            comboClass.Items.Clear();
        }

        // 載入科別班級資訊
        private void LoadClassDeptToForm()
        {
            this.SuspendLayout();
            // 科別、班級
            comboDept.Items.Add(AllStr);
            comboClass.Items.Add(AllStr);
            foreach (string name in DeptNameDict.Keys)
            {
                comboDept.Items.Add(name);
                foreach (string cname in DeptNameDict[name])
                    comboClass.Items.Add(cname);
            }

            comboDept.Text = AllStr;
            comboClass.Text = AllStr;
            this.ResumeLayout();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

            ControlEnable(false);
            ClearClassDept();
            comboDept.DropDownStyle = ComboBoxStyle.DropDownList;
            comboClass.DropDownStyle = ComboBoxStyle.DropDownList;

            // 載入畢業預警欄位
            LoadDgGWColumns();

            // 載入資料合理性檢查欄位
            LoadDgDataChkColumns();

            // 讀取班級科別資訊
            bgWorkerLoadDefault.RunWorkerAsync();

        }

        private void ControlEnable(bool value)
        {
            btnQuery.Enabled = comboDept.Enabled = comboClass.Enabled = btnDel.Enabled = btnUpdate.Enabled = value;
        }

        private void comboDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 清除班級，填入相對值
            comboClass.Text = "";
            SelectedDeptName = comboDept.Text;

            if (comboClass.Text != AllStr)
                comboClass.Items.Clear();

            if (SelectedDeptName == AllStr)
            {
                if (comboDept.Items.Count == 0)
                {
                    LoadClassDeptToForm();
                }
            }
            else
            {
                if (DeptNameDict.ContainsKey(SelectedDeptName))
                {
                    foreach (string name in DeptNameDict[SelectedDeptName])
                    {
                        comboClass.Items.Add(name);
                    }
                    if (comboClass.Items.Count > 0)
                        comboClass.SelectedIndex = 0;
                }
            }

        }

        // 載入畢業預警檢查欄位
        private void LoadDgGWColumns()
        {
            try
            {
                string textColumnStrig = @"
                    [
                        {
                            ""HeaderText"": ""學號"",
                            ""Name"": ""學號"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""班級"",
                            ""Name"": ""班級"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""座號"",
                            ""Name"": ""座號"",
                            ""Width"": 30,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""姓名"",
                            ""Name"": ""姓名"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""總學分數"",
                            ""Name"": ""總學分數"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""必修學分"",
                            ""Name"": ""必修學分"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""部定必修"",
                            ""Name"": ""部定必修"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""校訂必修"",
                            ""Name"": ""校訂必修"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""選修學分"",
                            ""Name"": ""選修學分"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""專業及實習"",
                            ""Name"": ""專業及實習"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""實習科目"",
                            ""Name"": ""實習科目"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""部定均修習"",
                            ""Name"": ""部定均修習"",
                            ""Width"": 40,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""必修均修習"",
                            ""Name"": ""必修均修習"",
                            ""Width"": 40,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""部定及格率"",
                            ""Name"": ""部定及格率"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""校訂必修及格率"",
                            ""Name"": ""校訂必修及格率"",
                            ""Width"": 80,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""是否畢業"",
                            ""Name"": ""是否畢業"",
                            ""Width"": 40,
                            ""ReadOnly"": true
                        },
                        {
                            ""HeaderText"": ""未達畢業標準說明"",
                            ""Name"": ""未達畢業標準說明"",
                            ""Width"": 140,
                            ""ReadOnly"": true
                        }
                    ]            
                
";
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<DataGridViewTextBoxColumnInfo> jsonObjArray = serializer.Deserialize<List<DataGridViewTextBoxColumnInfo>>(textColumnStrig);
                foreach (DataGridViewTextBoxColumnInfo jObj in jsonObjArray)
                {
                    DataGridViewTextBoxColumn dgt = new DataGridViewTextBoxColumn();
                    dgt.Name = jObj.Name;
                    dgt.Width = jObj.Width;
                    dgt.HeaderText = jObj.HeaderText;
                    dgt.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgt.ReadOnly = jObj.ReadOnly;
                    dgDataGW.Columns.Add(dgt);
                }

                // 加入學生功能按鈕
                DataGridViewButtonXColumn btnCol1 = new DataGridViewButtonXColumn();
                btnCol1.Name = "學生通知單";
                btnCol1.HeaderText = "學生通知單";
                btnCol1.Text = "列印";
                btnCol1.Width = 100;
                btnCol1.UseColumnTextForButtonValue = true;
                btnCol1.Click += BtnCol1_Click;
                dgDataGW.Columns.Add(btnCol1);

                // 加入班級功能按鈕
                DataGridViewButtonXColumn btnCol2 = new DataGridViewButtonXColumn();
                btnCol2.Name = "班級通知單";
                btnCol2.HeaderText = "班級通知單";
                btnCol2.Text = "列印";
                btnCol2.Width = 100;
                btnCol2.UseColumnTextForButtonValue = true;
                btnCol2.Click += BtnCol2_Click;
                dgDataGW.Columns.Add(btnCol2);

                // 加入學生待處理
                DataGridViewButtonXColumn btnCol3 = new DataGridViewButtonXColumn();
                btnCol3.Name = "學生待處理";
                btnCol3.HeaderText = "學生待處理";
                btnCol3.Text = "加入";
                btnCol3.Width = 80;
                btnCol3.UseColumnTextForButtonValue = true;
                btnCol3.Click += BtnCol3_Click;
                dgDataGW.Columns.Add(btnCol3);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        // 載入資料合理檢查欄位
        private void LoadDgDataChkColumns()
        {
            try
            {
                string textColumnStrig = @"
                [
                {
                    ""HeaderText"": ""學號"",
                    ""Name"": ""學號"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""科別"",
                    ""Name"": ""科別"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                }
                ,
                {
                    ""HeaderText"": ""班級"",
                    ""Name"": ""班級"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""座號"",
                    ""Name"": ""座號"",
                    ""Width"": 30,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""姓名"",
                    ""Name"": ""姓名"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""學年度"",
                    ""Name"": ""學年度"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""學期"",
                    ""Name"": ""學期"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""成績年級"",
                    ""Name"": ""成績年級"",
                    ""Width"": 40,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""科目名稱"",
                    ""Name"": ""科目名稱"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""科目級別"",
                    ""Name"": ""科目級別"",
                    ""Width"": 30,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""建議級別"",
                    ""Name"": ""建議級別"",
                    ""Width"": 30,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""領域"",
                    ""Name"": ""領域"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""分項"",
                    ""Name"": ""分項"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""校部定"",
                    ""Name"": ""校部定"",
                    ""Width"": 60,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""必選修"",
                    ""Name"": ""必選修"",
                    ""Width"": 60,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""學分"",
                    ""Name"": ""學分"",
                    ""Width"": 30,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""指定學年科目名稱"",
                    ""Name"": ""指定學年科目名稱"",
                    ""Width"": 100,
                    ""ReadOnly"": false
                },
                {
                    ""HeaderText"": ""使用課規"",
                    ""Name"": ""使用課規"",
                    ""Width"": 100,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""問題說明"",
                    ""Name"": ""問題說明"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                }
                ]           
                
";
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<DataGridViewTextBoxColumnInfo> jsonObjArray = serializer.Deserialize<List<DataGridViewTextBoxColumnInfo>>(textColumnStrig);
                foreach (DataGridViewTextBoxColumnInfo jObj in jsonObjArray)
                {
                    DataGridViewTextBoxColumn dgt = new DataGridViewTextBoxColumn();
                    dgt.Name = jObj.Name;
                    dgt.Width = jObj.Width;
                    dgt.HeaderText = jObj.HeaderText;
                    dgt.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgt.ReadOnly = jObj.ReadOnly;
                    dgDataChkEdit.Columns.Add(dgt);
                }


                // 加入刪除勾選
                DataGridViewCheckBoxColumn chkCol1 = new DataGridViewCheckBoxColumn();
                chkCol1.Name = "刪除";
                chkCol1.HeaderText = "刪除";
                chkCol1.Width = 50;
                chkCol1.TrueValue = "是";
                chkCol1.FalseValue = "否";
                chkCol1.IndeterminateValue = "否";
                dgDataChkEdit.Columns.Add(chkCol1);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnCol3_Click(object sender, EventArgs e)
        {
            MsgBox.Show("學生待處理");
        }

        private void BtnCol2_Click(object sender, EventArgs e)
        {
            MsgBox.Show("班級通知單");
        }

        private void BtnCol1_Click(object sender, EventArgs e)
        {
            MsgBox.Show("學生通知單");
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow drv in dgDataChkEdit.Rows)
            {
                if (drv.Cells["刪除"].Value != null)
                    if (drv.Cells["刪除"].Value.ToString() == "是")
                    {
                        MsgBox.Show("刪除");
                    }
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {

        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            // 畢業預警
            if (TabControlSelectedIndex == 0)
            {

            }

            // 資料合理檢查
            if (TabControlSelectedIndex == 1)
            {
                ControlEnable(false);
                bgwDataChkEditLoad.RunWorkerAsync();
            }

        }

        private void tabControl1_SelectedTabChanged(object sender, DevComponents.DotNetBar.TabStripTabChangedEventArgs e)
        {
            TabControlSelectedIndex = tabControl1.SelectedTabIndex;

        }

        private void AddErrorSubjectInfoDict(StudSubjectInfo ssi)
        {
            string key = ssi.SchoolYear + "_" + ssi.Semester + "_" + ssi.StudentID + "_" + ssi.SubjectName + "_" + ssi.SubjectLevel;
            if (!hasErrorSubjectInfoDict.ContainsKey(key))
                hasErrorSubjectInfoDict.Add(key, ssi);
        }
    }
}
