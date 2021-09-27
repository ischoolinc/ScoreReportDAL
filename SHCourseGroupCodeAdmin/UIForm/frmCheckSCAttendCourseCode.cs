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
    public partial class frmCheckSCAttendCourseCode : BaseForm
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        string _StrGradeYear = "1";
        int _SchoolYear = 1;
        int _Semester = 1;
        List<string> _grList = new List<string>();

        Dictionary<string, int> _ColIdxDict;

        List<rptSCAttendCodeChkInfo> SCAttendCodeChkInfoHasDataList = new List<rptSCAttendCodeChkInfo>();
        List<rptSCAttendCodeChkInfo> SCAttendCodeChkInfoErrorList = new List<rptSCAttendCodeChkInfo>();
        List<rptSCAttendCodeChkInfo> SCAttendCodeChkInfoNoList = new List<rptSCAttendCodeChkInfo>();

        public frmCheckSCAttendCourseCode()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _wb = new Workbook();
            _ColIdxDict = new Dictionary<string, int>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("修課檢核課程代碼中...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            btnRun.Enabled = true;

            if (_wb != null)
            {
                Utility.ExprotXls("修課檢核課程代碼", _wb);
            }
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            da.LoadMOEGroupCodeDict();

            SCAttendCodeChkInfoHasDataList.Clear();
            SCAttendCodeChkInfoErrorList.Clear();
            SCAttendCodeChkInfoNoList.Clear();

            // 取得使用者設定學年度學期
            List<rptSCAttendCodeChkInfo> SCAttendCodeChkInfoList = da.GetGetStudentCourseInfoBySchoolYearSemester(_SchoolYear, _Semester, _StrGradeYear);

            Dictionary<string, List<string>> chkHasCourseCodeDict = new Dictionary<string, List<string>>();

            foreach (rptSCAttendCodeChkInfo data in SCAttendCodeChkInfoList)
            {
                if (!string.IsNullOrEmpty(data.CourseCode))
                {
                    if (!chkHasCourseCodeDict.ContainsKey(data.StudentID))
                        chkHasCourseCodeDict.Add(data.StudentID, new List<string>());

                    chkHasCourseCodeDict[data.StudentID].Add(data.CourseCode);
                }

                if (data.ErrorMsgList.Count > 0)
                {
                    SCAttendCodeChkInfoErrorList.Add(data);
                }
                else
                {
                    SCAttendCodeChkInfoHasDataList.Add(data);
                }
            }
            _bgWorker.ReportProgress(40);

            // 取得課程大表資料
            Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = da.GetCourseGroupCodeDict();
            List<DataRow> haGDCCodeStudents = da.GetHasGDCCodeStudent(_StrGradeYear);

            // 這學期需要開課
            List<MOECourseCodeInfo> thisCousreCodeList = new List<MOECourseCodeInfo>();

            // 有群科班代碼學生，每筆代表一位學生
            foreach (DataRow dr in haGDCCodeStudents)
            {
                string sid = dr["student_id"] + "";
                string gdc_code = dr["gdc_code"] + "";
                thisCousreCodeList.Clear();
                // 取得大表
                if (MOECourseDict.ContainsKey(gdc_code))
                {
                    // 學生年級學期
                    string gr = dr["grade_year"] + "";
                    int idx = -1;

                    if (gr == "1" && _Semester == 1)
                        idx = 0;

                    if (gr == "1" && _Semester == 2)
                        idx = 1;

                    if (gr == "2" && _Semester == 1)
                        idx = 2;
                    if (gr == "2" && _Semester == 2)
                        idx = 3;

                    if (gr == "3" && _Semester == 1)
                        idx = 4;

                    if (gr == "3" && _Semester == 2)
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

                        rptSCAttendCodeChkInfo data = new rptSCAttendCodeChkInfo();
                        data.StudentID = sid;
                        data.ClassName = dr["class_name"] + "";
                        data.SeatNo = dr["seat_no"] + "";
                        data.GradeYear = gr;
                        data.StudentNumber = dr["student_number"] + "";
                        data.StudentName = dr["student_name"] + "";
                        data.CourseCode = coInfo.course_code;
                        data.credit_period = coInfo.credit_period;
                        data.open_type = coInfo.open_type;
                        data.gdc_code = gdc_code;
                        data.IsRequired = coInfo.is_required;
                        data.RequiredBy = coInfo.require_by;
                        data.SubjectName = coInfo.subject_name;

                        //if (data.IsRequired == "必修")
                        SCAttendCodeChkInfoNoList.Add(data);
                    }


                }

            }


            //// 需要調整 ---
            //// 取得課程大表資料
            //Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = da.GetCourseGroupCodeDict();

            //// 
            //List<DataRow> haGDCCodeStudents = da.GetHasGDCCodeStudent();

            //foreach (DataRow dr in haGDCCodeStudents)
            //{
            //    string sid = dr["student_id"] + "";
            //    string gdc_code = dr["gdc_code"] + "";
            //    if (MOECourseDict.ContainsKey(gdc_code))
            //    {
            //        foreach (MOECourseCodeInfo Mo in MOECourseDict[gdc_code])
            //        {
            //            // 已修
            //            if (chkHasCourseCodeDict.ContainsKey(sid))
            //            {
            //                if (chkHasCourseCodeDict[sid].Contains(Mo.course_code))
            //                    continue;
            //            }

            //            rptSCAttendCodeChkInfo data = new rptSCAttendCodeChkInfo();
            //            data.StudentID = sid;
            //            data.ClassName = dr["class_name"] + "";
            //            data.SeatNo = dr["seat_no"] + "";
            //            data.StudentNumber = dr["student_number"] + "";
            //            data.StudentName = dr["student_name"] + "";
            //            data.CourseCode = Mo.course_code;
            //            data.credit_period = Mo.credit_period;
            //            data.gdc_code = gdc_code;
            //            data.IsRequired = Mo.is_required;
            //            data.RequiredBy = Mo.require_by;
            //            data.SubjectName = Mo.subject_name;


            //            SCAttendCodeChkInfoNoList.Add(data);
            //        }
            //    }
            //}

            // ----


            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.修課檢核課程代碼樣板));
            Worksheet wstSC = _wb.Worksheets["檢查修課學生課程代碼"];
            //wstSC.Name = _GradeYear + "年級_檢查修課學生課程代碼";
            wstSC.Name = "檢查修課學生課程代碼";

            Worksheet wstSCError = _wb.Worksheets["檢查修課學生課程代碼(有差異)"];
            //wstSCError.Name = _GradeYear + "年級_檢查修課學生課程代碼(有差異)";
            wstSCError.Name = "檢查修課學生課程代碼(有差異)";

            Worksheet wstSCNo = _wb.Worksheets["檢查學生應修課程代碼未修"];
            //wstSCNo.Name = _GradeYear + "年級_檢查學生應修課程代碼未修";
            wstSCNo.Name = "檢查學生應修課程代碼未修";


            int rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wstSC.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstSC.Cells[0, co].StringValue, co);
            }

            foreach (rptSCAttendCodeChkInfo data in SCAttendCodeChkInfoHasDataList)
            {
                wstSC.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstSC.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstSC.Cells[rowIdx, GetColIndex("年級")].PutValue(data.GradeYear);
                wstSC.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSC.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSC.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSC.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSC.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wstSC.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSC.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstSC.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSC.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSC.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstSC.Cells[rowIdx, GetColIndex("節數")].PutValue(data.Period);
                wstSC.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSC.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSC.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstSC.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
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
            foreach (rptSCAttendCodeChkInfo data in SCAttendCodeChkInfoErrorList)
            {
                wstSCError.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstSCError.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstSCError.Cells[rowIdx, GetColIndex("年級")].PutValue(data.GradeYear);
                wstSCError.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSCError.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSCError.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSCError.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSCError.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wstSCError.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSCError.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstSCError.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSCError.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSCError.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstSCError.Cells[rowIdx, GetColIndex("節數")].PutValue(data.Period);
                wstSCError.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSCError.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSCError.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstSCError.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
                if (data.ErrorMsgList.Contains("群科班代碼無法對照"))
                {
                    wstSCError.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + "");
                }
                else
                {
                    // 科目名稱無法對照
                    if (data.ErrorMsgList.Contains("科目名稱"))
                    {
                        wstSCError.Cells[rowIdx, GetColIndex("說明")].PutValue("科目名稱無法對照");
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
            foreach (rptSCAttendCodeChkInfo data in SCAttendCodeChkInfoNoList)
            {
                wstSCNo.Cells[rowIdx, GetColIndex("學年度")].PutValue(_SchoolYear);
                wstSCNo.Cells[rowIdx, GetColIndex("學期")].PutValue(_Semester);
                wstSCNo.Cells[rowIdx, GetColIndex("年級")].PutValue(data.GradeYear);
                wstSCNo.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSCNo.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSCNo.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSCNo.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSCNo.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSCNo.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSCNo.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSCNo.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSCNo.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSCNo.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstSCNo.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
                rowIdx++;
            }
            wstSCNo.AutoFitColumns();


            if (SCAttendCodeChkInfoErrorList.Count == 0)
            {
                _wb.Worksheets.RemoveAt(wstSCError.Name);
            }

            if (SCAttendCodeChkInfoNoList.Count == 0)
                _wb.Worksheets.RemoveAt(wstSCNo.Name);

            _bgWorker.ReportProgress(100);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCheckSCAttendCourseCode_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            try
            {
                iptSchoolYear.IsInputReadOnly = true;
                iptSemester.IsInputReadOnly = true;
                // 先註解之後再使用
                //iptSchoolYear.Value = iptSchoolYear.MinValue;
                //iptSemester.Value = iptSemester.MinValue;
                //int sy, ss;
                //if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
                //{
                //    if (sy <= iptSchoolYear.MaxValue && sy >= iptSchoolYear.MinValue)
                //        iptSchoolYear.Value = sy;
                //}

                //if (int.TryParse(K12.Data.School.DefaultSemester, out ss))
                //{
                //    if (ss <= iptSemester.MaxValue && ss >= iptSemester.MinValue)
                //        iptSemester.Value = ss;
                //}


                // 2021/9/24 討論先固定110學年度第1學期
                iptSchoolYear.MinValue = iptSchoolYear.MaxValue = 110;
                iptSemester.MinValue = iptSemester.MaxValue = 1;

                iptSchoolYear.Value = 110;
                iptSemester.Value = 1;

                // 載入年級
                _grList = da.GetClassGradeYear();
                cboGradeYear.Items.Add("全部");
                foreach (string gr in _grList)
                    cboGradeYear.Items.Add(gr);

                cboGradeYear.Text = "全部";
                cboGradeYear.DropDownStyle = ComboBoxStyle.DropDownList;

                // 說明文字
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("說明：");
                sb.AppendLine("學生狀態：一般、延修。");
                sb.AppendLine("讀取選擇學年度、學期、年級，學生修課紀錄，透過學生群科班代碼與課程代碼大表群科班代碼比對，");
                sb.AppendLine("以科目名稱 +校訂部定+必選修，比對出課程代碼。");
                sb.AppendLine("");
                sb.AppendLine("工作表:檢查學生應修課程代碼未修，依課程代碼大表為主，與工作表:檢查修課學生課程代碼，透過課程代碼進行比對，該學年度學期沒有修課科目會被列出。");


                txtDesc.Text = sb.ToString();

            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
                //Console.WriteLine(ex.Message);
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            _SchoolYear = iptSchoolYear.Value;
            _Semester = iptSemester.Value;
            if (cboGradeYear.Text == "全部")
                _StrGradeYear = string.Join(",", _grList.ToArray());
            else
                _StrGradeYear = cboGradeYear.Text;

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
