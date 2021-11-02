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

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateCourseByGPlan108_C : BaseForm
    {
        List<string> _ClassIDList = new List<string>();
        DataAccess da;
        string _SchoolYear = "";
        string _Semester = "";
        List<CClassCourseInfo> CClassCourseInfoList;
        BackgroundWorker _bwWorker;
        List<string> _errClassList = new List<string>();

        Dictionary<string, SubjectCourseInfo> _SubjectCourseInfoDict;

        public frmCreateCourseByGPlan108_C(List<string> ClassIDs)
        {
            InitializeComponent();
            da = new DataAccess();
            _ClassIDList = ClassIDs;
            CClassCourseInfoList = new List<CClassCourseInfo>();
            _SubjectCourseInfoDict = new Dictionary<string, SubjectCourseInfo>();
            _bwWorker = new BackgroundWorker();
            _bwWorker.WorkerReportsProgress = true;
            _bwWorker.DoWork += _bwWorker_DoWork;
            _bwWorker.ProgressChanged += _bwWorker_ProgressChanged;
            _bwWorker.RunWorkerCompleted += _bwWorker_RunWorkerCompleted;
        }

        private void _bwWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取資料中 ...", e.ProgressPercentage);
        }

        private void _bwWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MsgBox.Show("班級：" + string.Join(",", _errClassList.ToArray()) + "，使用課程規劃非108適用，無法產生。");
            }
            else
            {
                ControlEnable(true);              
                FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取完成");
            }
        }

        private void _bwWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _errClassList.Clear();
            _bwWorker.ReportProgress(1);
            CClassCourseInfoList = da.GetCClassCourseInfoList(_ClassIDList);

            // 檢查課程規劃表
            foreach (CClassCourseInfo data in CClassCourseInfoList)
            {
                if (data.RefGPlanXML == null)
                    _errClassList.Add(data.ClassName);
            }

            if (_errClassList.Count > 0)
                e.Cancel = true;

            Dictionary<string, List<string>> classStudentIDList = da.GetClassStudentDict(_ClassIDList);

            _SubjectCourseInfoDict.Clear();

            // 整理目前學年度學期年級，跨班開課
            foreach (CClassCourseInfo data in CClassCourseInfoList)
            {
                // 加入班級學生
                if (classStudentIDList.ContainsKey(data.ClassID))
                {
                    data.RefStudentIDList = classStudentIDList[data.ClassID];
                }

                if (data.RefGPlanXML != null)
                {
                    foreach (XElement subjElm in data.RefGPlanXML.Elements("Subject"))
                    {
                        if (data.GradeYear == subjElm.Attribute("GradeYear").Value && _Semester == subjElm.Attribute("Semester").Value)
                        {
                            if (subjElm.Attribute("開課方式").Value == "跨班")
                            {
                                data.OpenSubjectSourceList.Add(subjElm);
                                string subjName = subjElm.Attribute("SubjectName").Value;
                                if (!data.SubjectBDict.ContainsKey(subjName))
                                    data.SubjectBDict.Add(subjName, false);


                                // 處理科目
                                string openSems = "0";

                                if (data.GradeYear == "3" && _Semester == "2")
                                {
                                    openSems = "6";
                                }
                                else if (data.GradeYear == "3" && _Semester == "1")
                                {
                                    openSems = "5";
                                }
                                else if (data.GradeYear == "2" && _Semester == "2")
                                {
                                    openSems = "4";
                                }
                                else if (data.GradeYear == "2" && _Semester == "1")
                                {
                                    openSems = "3";
                                }
                                else if (data.GradeYear == "1" && _Semester == "2")
                                {
                                    openSems = "2";
                                }
                                else if (data.GradeYear == "1" && _Semester == "1")
                                {
                                    openSems = "1";
                                }
                                else { }

                                string subjKey = openSems + "_" + subjName;

                                if (!_SubjectCourseInfoDict.ContainsKey(subjKey))
                                {
                                    SubjectCourseInfo sci = new SubjectCourseInfo();
                                    sci.SubjectName = subjName;
                                    sci.SubjectXML = subjElm;
                                    sci.SchoolYear = _SchoolYear;
                                    sci.Semester = _Semester;
                                    sci.CourseCount = 0;
                                    sci.ClassNameDict = new Dictionary<string, string>();
                                    sci.ClassStudentIDDict = new Dictionary<string, List<string>>();
                                    sci.OpenSemester = openSems;
                                    _SubjectCourseInfoDict.Add(subjKey, sci);
                                }

                                // 班級
                                if (!_SubjectCourseInfoDict[subjKey].ClassNameDict.ContainsKey(data.ClassName))
                                {
                                    _SubjectCourseInfoDict[subjKey].ClassNameDict.Add(data.ClassName, data.ClassID);
                                }

                                // 班級學生
                                if (!_SubjectCourseInfoDict[subjKey].ClassStudentIDDict.ContainsKey(data.ClassID))
                                {
                                    _SubjectCourseInfoDict[subjKey].ClassStudentIDDict.Add(data.ClassID, data.RefStudentIDList);
                                }
                            }
                        }

                    }
                }
            }




            _bwWorker.ReportProgress(100);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ControlEnable(bool value)
        {
            cboSchoolYear.Enabled = cboSemester.Enabled = btnCreate.Enabled = value;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            frmCreateCourseByGPlan108_C_Detail fcc = new frmCreateCourseByGPlan108_C_Detail();
            fcc.SetSchoolYearSemester(_SchoolYear, _Semester);
            fcc.SetSubjectCourseInfoDict(_SubjectCourseInfoDict);

            if (fcc.ShowDialog() == DialogResult.OK)
            {
                this.Close();
            }
        }

        private void frmCreateCourseByGPlan108_C_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            // 預設值
            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            cboSemester.DropDownStyle = ComboBoxStyle.DropDownList;
            cboSemester.Text = K12.Data.School.DefaultSemester;
            int sy;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
            {
                for (int i = sy - 3; i <= (sy + 3); i++)
                {
                    cboSchoolYear.Items.Add(i);
                }
            }
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;

            cboSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;

            _Semester = cboSemester.Text;
            _SchoolYear = cboSchoolYear.Text;

            ControlEnable(false);
            _bwWorker.RunWorkerAsync();
        }
    }
}
