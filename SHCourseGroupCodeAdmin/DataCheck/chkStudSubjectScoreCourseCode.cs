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
using Aspose.Cells;
using SHCourseGroupCodeAdmin.DAO;
using System.IO;


namespace SHCourseGroupCodeAdmin.DataCheck
{
    public partial class chkStudSubjectScoreCourseCode : BaseForm
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        int _GradeYear = 1;
        List<StudentSubjectInfoChk> _StudentSubjectInfoChkList;
        Dictionary<string, int> _ColIdxDict;

        public chkStudSubjectScoreCourseCode()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _wb = new Workbook();
            _StudentSubjectInfoChkList = new List<StudentSubjectInfoChk>();
            _ColIdxDict = new Dictionary<string, int>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;

        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生學期科目成績課程代碼檢查 產生完成");

            btnRun.Enabled = true;

            if (_wb != null)
            {
                Utility.ExprotXls("學生學期科目成績課程代碼檢查", _wb);
            }
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生學期科目成績課程代碼檢查 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);

            _StudentSubjectInfoChkList = da.GetStudentSubjectInfoListByGradeYear(_GradeYear);

            _bgWorker.ReportProgress(30);

            // 取得課程代碼大表科目內容
            List<MOECourseCodeInfo> MOECoursedList = da.GetCourseGroupCodeList();

            //比對大表使用
            Dictionary<string, MOECourseCodeInfo> MOECoursedMapDict = new Dictionary<string, MOECourseCodeInfo>();

            Dictionary<string, List<MOECourseCodeInfo>> CourseGroupCodeDict = da.GetCourseGroupCodeDict();


            foreach (MOECourseCodeInfo data in MOECoursedList)
            {
                // 群組代碼 + 科目名稱 +部校訂+必選修
                string key = data.group_code + "_" + data.subject_name + "_" + data.require_by + "_" + data.is_required;
                if (!MOECoursedMapDict.ContainsKey(key))
                    MOECoursedMapDict.Add(key, data);
            }

            _bgWorker.ReportProgress(40);

            // 比對資料，找出課程代碼
            foreach (StudentSubjectInfoChk stud in _StudentSubjectInfoChkList)
            {
                foreach (SubjectInfoChk subj in stud.SubjectInfoChkList)
                {
                    string key = stud.gdc_code + "_" + subj.SubjectName + "_" + subj.RequireBy + "_" + subj.IsRequired;

                    if (MOECoursedMapDict.ContainsKey(key))
                    {
                        subj.course_code = MOECoursedMapDict[key].course_code;
                        subj.GroupCode = stud.gdc_code;
                        subj.credit_period = MOECoursedMapDict[key].credit_period;
                    }
                }

            }

            // 取得學分對照表
            Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

            List<string> errMesList = new List<string>();
            List<string> errItem = new List<string>();

            // 處理沒有比對到原因
            foreach (StudentSubjectInfoChk stud in _StudentSubjectInfoChkList)
            {
                foreach (SubjectInfoChk subj in stud.SubjectInfoChkList)
                {
                    if (string.IsNullOrEmpty(subj.course_code))
                    {
                        errMesList.Clear();
                        errItem.Clear();
                        errItem.Add("科目名稱");
                        errItem.Add("部定校訂");
                        errItem.Add("必修選修");
                        errItem.Add("學分數");
                        subj.Memo = "";

                        if (CourseGroupCodeDict.ContainsKey(stud.gdc_code))
                        {

                            if (subj.CheckCreditPass(mappingTable))
                            {
                                errItem.Remove("學分數");
                            }

                            foreach (MOECourseCodeInfo mm in CourseGroupCodeDict[stud.gdc_code])
                            {
                                if (subj.SubjectName == mm.subject_name && subj.IsRequired == mm.is_required)
                                {
                                    errItem.Remove("科目名稱");
                                    errItem.Remove("必修選修");
                                    break;
                                }
                            }

                            foreach (MOECourseCodeInfo mm in CourseGroupCodeDict[stud.gdc_code])
                            {
                                if (subj.SubjectName == mm.subject_name && subj.RequireBy == mm.require_by)
                                {
                                    errItem.Remove("科目名稱");
                                    errItem.Remove("部定校訂");
                                    break;
                                }
                            }

                            foreach (MOECourseCodeInfo mm in CourseGroupCodeDict[stud.gdc_code])
                            {
                                if (subj.SubjectName == mm.subject_name)
                                {
                                    errItem.Remove("科目名稱");
                                    break;
                                }
                            }
                        }
                        else
                        {
                            errMesList.Add("群科班代碼無法對照");
                        }


                        if (errItem.Count > 0)
                        {
                            errMesList.Add(string.Join("、", errItem.ToArray()) + " 無法對照");
                        }

                        subj.Memo = string.Join(",", errMesList.ToArray());
                    }
                }
            }





            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.學期科目成績課程代碼檢查樣版));
            Worksheet wst = _wb.Worksheets["學期科目成績"];
            wst.Name = _GradeYear + "年級";
            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            int rowIdx = 1;
            foreach (StudentSubjectInfoChk data in _StudentSubjectInfoChkList)
            {
                foreach (SubjectInfoChk si in data.SubjectInfoChkList)
                {
                    wst.Cells[rowIdx, GetColIndex("學生系統編號")].PutValue(data.StudentID);
                    wst.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                    wst.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                    wst.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                    wst.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                    wst.Cells[rowIdx, GetColIndex("學年度")].PutValue(si.SchoolYear);
                    wst.Cells[rowIdx, GetColIndex("學期")].PutValue(si.Semester);
                    wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(si.SubjectName);
                    wst.Cells[rowIdx, GetColIndex("科目級別")].PutValue(si.SubjectLevel);
                    wst.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(si.RequireBy);
                    wst.Cells[rowIdx, GetColIndex("必修選修")].PutValue(si.IsRequired);
                    wst.Cells[rowIdx, GetColIndex("學分數")].PutValue(si.Credit);
                    wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(si.course_code);
                    wst.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(si.credit_period);
                    wst.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(si.GroupCode);
                    wst.Cells[rowIdx, GetColIndex("說明")].PutValue(si.Memo);
                    rowIdx++;
                }
            }

            wst.AutoFitColumns();

            _bgWorker.ReportProgress(100);
        }

        private void chkStudSubjectScoreCourseCode_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            try
            {
                iptGradeYear.Value = 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            _GradeYear = iptGradeYear.Value;
            _bgWorker.RunWorkerAsync();
        }

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }
    }
}
