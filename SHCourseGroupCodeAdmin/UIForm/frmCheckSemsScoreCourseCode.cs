using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHCourseGroupCodeAdmin.DAO;
using FISCA.Presentation.Controls;
using System.Xml.Linq;
using Aspose.Cells;
using System.IO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCheckSemsScoreCourseCode : BaseForm
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        string _GradeYear = "1";
        int _SchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
        int _Semester = int.Parse(K12.Data.School.DefaultSemester);

        List<string> _grList = new List<string>();

        Dictionary<string, int> _ColIdxDict;

        List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoHasDataList = new List<rptStudSemsScoreCodeChkInfo>();
        List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoErrorList = new List<rptStudSemsScoreCodeChkInfo>();
        List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoNoList = new List<rptStudSemsScoreCodeChkInfo>();

        public frmCheckSemsScoreCourseCode()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _wb = new Workbook();
            _ColIdxDict = new Dictionary<string, int>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            cboGradeYear.Enabled = btnRun.Enabled = iptSchoolYear.Enabled = iptSemester.Enabled = true;

            // 檢查 所選學年期和系統學年期不同，查詢目標學生 所選學年期的學期對照表。
            if (!(_SchoolYear == int.Parse(K12.Data.School.DefaultSchoolYear) && _Semester == int.Parse(K12.Data.School.DefaultSemester)))
            {
                DataTable notHasSemsHistoryStudents = da.GetStudentListNotHasSemsHistory(_GradeYear, _SchoolYear, _Semester);

                //若查詢到 學期對照表 不完整的學生，則跳出提示。
                if (notHasSemsHistoryStudents.Rows.Count > 0)
                {
                    NotHasSemsHistoryStudents nhshs = new NotHasSemsHistoryStudents(notHasSemsHistoryStudents, _GradeYear, _SchoolYear, _Semester);

                    if (nhshs.ShowDialog() == DialogResult.Cancel)
                    {
                        _wb = null;  //取消就不印報表
                        return;
                    }
                }
            }



            if (_wb != null)
            {
                Utility.ExprotXls("學期成績檢核課程代碼", _wb);
            }
        }



        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績檢核課程代碼中...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            da.LoadMOEGroupCodeDict();

            StudSemsScoreCodeChkInfoHasDataList.Clear();
            StudSemsScoreCodeChkInfoErrorList.Clear();
            StudSemsScoreCodeChkInfoNoList.Clear();

            /// https://3.basecamp.com/4399967/buckets/15765350/todos/4245362818#__recording_4286998512
            /// 判斷學年度/學期 是否和當前學年度學期一樣，若是，成績年級=現在年級 處理，若否，成績年級=學期對照表上的年級

            // 取得學期成績
            List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoList = new List<rptStudSemsScoreCodeChkInfo>();

            // 判斷學年度/學期 是否和當前學年度學期一樣
            if (_SchoolYear == int.Parse(K12.Data.School.DefaultSchoolYear) && _Semester == int.Parse(K12.Data.School.DefaultSemester))
            {
                // 相同
                StudSemsScoreCodeChkInfoList = da.GetStudentSemsScoreInfo(_GradeYear, _SchoolYear, _Semester);
            }
            else
            {
                //  不同

                // 取得有學期對照表(年級)的學生 當時的學期成績
                StudSemsScoreCodeChkInfoList = da.GetStudentSemsScoreInfoByExistingSemsHistory(_GradeYear, _SchoolYear, _Semester);
            }


            Dictionary<string, List<string>> chkHasCourseCodeDict = new Dictionary<string, List<string>>();

            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoList)
            {
                if (!string.IsNullOrEmpty(data.CourseCode))
                {
                    if (!chkHasCourseCodeDict.ContainsKey(data.StudentID))
                        chkHasCourseCodeDict.Add(data.StudentID, new List<string>());

                    chkHasCourseCodeDict[data.StudentID].Add(data.CourseCode);
                }

                if (data.ErrorMsgList.Count > 0)
                {
                    StudSemsScoreCodeChkInfoErrorList.Add(data);
                }
                else
                {
                    StudSemsScoreCodeChkInfoHasDataList.Add(data);
                }
            }
            _bgWorker.ReportProgress(40);

            // 處理應修未修
            
            // 取得課程大表資料
            Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = da.GetCourseGroupCodeDict();            

            List<DataRow> haGDCCodeStudents = new List<DataRow>();

            if (_SchoolYear == int.Parse(K12.Data.School.DefaultSchoolYear) && _Semester == int.Parse(K12.Data.School.DefaultSemester))
            {
                haGDCCodeStudents = da.GetHasGDCCodeStudent(_GradeYear);
            }
            else
            {
                haGDCCodeStudents = da.GetHasGDCCodeAndSemHistoryStudent(_GradeYear, _SchoolYear, _Semester);
            }

            // 這學期需要開課(應要結算到學期成績)
            List<MOECourseCodeInfo> thisCousreCodeList = new List<MOECourseCodeInfo>();

            foreach (DataRow dr in haGDCCodeStudents)
            {
                string sid = dr["student_id"] + "";
                string gdc_code = dr["gdc_code"] + "";

                thisCousreCodeList.Clear();
                // 取得大表
                if (MOECourseDict.ContainsKey(gdc_code))
                {
                    // 年級
                    string gr = dr["grade_year"] + ""; ;
                    // 成績年級
                    string sgr = "";

                    if (_SchoolYear == int.Parse(K12.Data.School.DefaultSchoolYear) && _Semester == int.Parse(K12.Data.School.DefaultSemester))
                    {
                        // 所選學年期=當前學年期，現在年級=成績年級
                        sgr = gr;
                    }
                    else
                    {
                        //若所選學年期 !=當前學年期，則視學期對照表的年級為成績年級
                        sgr = dr["his_grade_year"] + "";
                    }

                    int idx = -1;

                    if (sgr == "1" && _Semester == 1)
                        idx = 0;

                    if (sgr == "1" && _Semester == 2)
                        idx = 1;

                    if (sgr == "2" && _Semester == 1)
                        idx = 2;

                    if (sgr == "2" && _Semester == 2)
                        idx = 3;

                    if (sgr == "3" && _Semester == 1)
                        idx = 4;

                    if (sgr == "3" && _Semester == 2)
                        idx = 5;

                    foreach (MOECourseCodeInfo couInfo in MOECourseDict[gdc_code])
                    {
                        char[] cp = couInfo.open_type.ToArray();
                        if (idx != -1 && idx < cp.Length)
                        {
                            // 需要開課
                            if (cp[idx] != '-')
                                thisCousreCodeList.Add(couInfo);
                        }
                    }

                    foreach (MOECourseCodeInfo coInfo in thisCousreCodeList)
                    {
                        if (chkHasCourseCodeDict.ContainsKey(sid))
                        {
                            if (chkHasCourseCodeDict[sid].Contains(coInfo.course_code))
                                continue;
                        }

                        rptStudSemsScoreCodeChkInfo data = new rptStudSemsScoreCodeChkInfo();
                        data.StudentID = sid;
                        data.ClassName = dr["class_name"] + "";
                        data.SeatNo = dr["seat_no"] + "";
                        data.GradeYear = gr;
                        data.SemsGradeYear = sgr;
                        data.StudentNumber = dr["student_number"] + "";
                        data.StudentName = dr["student_name"] + "";
                        data.CourseCode = coInfo.course_code;
                        data.credit_period = coInfo.credit_period;
                        data.gdc_code = gdc_code;
                        data.IsRequired = coInfo.is_required;
                        data.RequiredBy = coInfo.require_by;
                        data.SubjectName = coInfo.subject_name;
                        data.ScoreType = coInfo.score_type;
                        //data.GraduationPlanName = 
                        StudSemsScoreCodeChkInfoNoList.Add(data);
                    }
                }
            }


            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.學期成績檢核課程代碼樣板));
             Worksheet wstSC = _wb.Worksheets["檢查學期成績課程代碼"];
            wstSC.Name = _GradeYear + "年級_檢查學期成績課程代碼";

            Worksheet wstSCError = _wb.Worksheets["檢查學期成績課程代碼(有差異)"];
            wstSCError.Name = _GradeYear + "年級_檢查學期成績課程代碼(有差異)";

            Worksheet wstSCNo = _wb.Worksheets["檢查學生應修課程代碼未修"];
            wstSCNo.Name = _GradeYear + "年級_檢查學生應修課程代碼未修";

            int rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wstSC.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstSC.Cells[0, co].StringValue, co);
            }

            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoHasDataList)
            {
                wstSC.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSC.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSC.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSC.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSC.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstSC.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstSC.Cells[rowIdx, GetColIndex("成績年級")].PutValue(data.SemsGradeYear);
                wstSC.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSC.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstSC.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSC.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSC.Cells[rowIdx, GetColIndex("分項類別")].PutValue(data.ScoreType);
                wstSC.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstSC.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSC.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSC.Cells[rowIdx, GetColIndex("課程規劃表名稱")].PutValue(data.GraduationPlanName);
                rowIdx++;
            }

            wstSC.AutoFitColumns();

            rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wstSCError.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstSCError.Cells[0, co].StringValue, co);
            }
            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoErrorList)
            {
                wstSCError.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSCError.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSCError.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSCError.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSCError.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstSCError.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstSCError.Cells[rowIdx, GetColIndex("成績年級")].PutValue(data.SemsGradeYear);
                wstSCError.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSCError.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstSCError.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSCError.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSCError.Cells[rowIdx, GetColIndex("分項類別")].PutValue(data.ScoreType);
                wstSCError.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstSCError.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSCError.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSCError.Cells[rowIdx, GetColIndex("課程規劃表名稱")].PutValue(data.GraduationPlanName);
                if (data.ErrorMsgList.Contains("課程規劃表 不同"))
                {
                    wstSCError.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + "");
                }
                else
                {
                    // 科目名稱無法對照
                    if (data.ErrorMsgList.Contains("科目名稱與級別"))
                    {
                        wstSCError.Cells[rowIdx, GetColIndex("說明")].PutValue("沒有此科目名稱與級別");
                    }
                    else
                    {
                        // 科目名稱有對到其他有問題
                        wstSCError.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + " 不同");
                    }
                }
                rowIdx++;
            }
            wstSCError.AutoFitColumns();

            rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wstSCNo.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstSCNo.Cells[0, co].StringValue, co);
            }
            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoNoList)
            {
                wstSCNo.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSCNo.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSCNo.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSCNo.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSCNo.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSCNo.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSCNo.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSCNo.Cells[rowIdx, GetColIndex("分項類別")].PutValue(data.ScoreType);
                wstSCNo.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSCNo.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSCNo.Cells[rowIdx, GetColIndex("課程規劃表名稱")].PutValue(data.GraduationPlanName);
                rowIdx++;
            }
            wstSCNo.AutoFitColumns();


            if (StudSemsScoreCodeChkInfoErrorList.Count == 0)
            {
                _wb.Worksheets.RemoveAt(wstSCError.Name);
            }

            if (StudSemsScoreCodeChkInfoNoList.Count == 0)
                _wb.Worksheets.RemoveAt(wstSCNo.Name);

            _bgWorker.ReportProgress(100);


        }

        private void frmCheckSemsScoreCourseCode_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            try
            {
                // 載入年級
                _grList = da.GetClassGradeYear();
                cboGradeYear.Items.Add("全部");
                foreach (string gr in _grList)
                    cboGradeYear.Items.Add(gr);

                cboGradeYear.Text = "全部";


                iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
                iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

                // 說明文字
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("說明：");
                sb.AppendLine("1.學生狀態：一般、延修。");
                sb.AppendLine("2.讀取選擇學年度、學期、年級，學生學期成績及學期對照表，透過學生指定課程規畫表、群科班代碼與課程規畫表比對");
                sb.AppendLine("若所選擇的學年度、學期與系統當前的學年度、學期不同，以學期對照表上的年級，找出當時的年級。");
                sb.AppendLine("並以科目名稱+科目級別，比對出課程代碼。");
                sb.AppendLine("");
                sb.AppendLine("3.工作表:檢查學生應修課程代碼未修，依課程代碼大表為主，與工作表:檢查學期成績課程代碼，透過課程代碼進行比對，該學年度學期沒有修課科目會被列出。");

                txtDesc.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = cboGradeYear.Enabled = iptSchoolYear.Enabled = iptSemester.Enabled = false;
            //_GradeYear = iptGradeYear.Value;
            _SchoolYear = int.Parse(iptSchoolYear.Text);
            _Semester = int.Parse(iptSemester.Text);

            if (cboGradeYear.Text == "全部")
                _GradeYear = string.Join(",", _grList.ToArray());
            else
                _GradeYear = cboGradeYear.Text;

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
