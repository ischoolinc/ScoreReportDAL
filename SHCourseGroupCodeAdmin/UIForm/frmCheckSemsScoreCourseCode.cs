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
        int _GradeYear = 1;

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
            iptGradeYear.Enabled = btnRun.Enabled = true;

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

            List<rptStudSemsScoreCodeChkInfo> StudSemsScoreCodeChkInfoList = da.GetStudentSemsScoreInfoByGradeYear(_GradeYear);

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
            // 取得課程大表資料
            Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = da.GetCourseGroupCodeDict();

            List<DataRow> haGDCCodeStudents = da.GetHasGDCCodeStudentByGradeYear(_GradeYear);

            foreach (DataRow dr in haGDCCodeStudents)
            {
                string sid = dr["student_id"] + "";
                string gdc_code = dr["gdc_code"] + "";
                if (MOECourseDict.ContainsKey(gdc_code))
                {
                    foreach (MOECourseCodeInfo Mo in MOECourseDict[gdc_code])
                    {
                        // 已修
                        if (chkHasCourseCodeDict.ContainsKey(sid))
                        {
                            if (chkHasCourseCodeDict[sid].Contains(Mo.course_code))
                                continue;
                        }

                        rptStudSemsScoreCodeChkInfo data = new rptStudSemsScoreCodeChkInfo();
                        data.StudentID = sid;
                        data.ClassName = dr["class_name"] + "";
                        data.SeatNo = dr["seat_no"] + "";
                        data.StudentNumber = dr["student_number"] + "";
                        data.StudentName = dr["student_name"] + "";
                        data.CourseCode = Mo.course_code;
                        data.credit_period = Mo.credit_period;
                        data.gdc_code = gdc_code;
                        data.IsRequired = Mo.is_required;
                        data.RequiredBy = Mo.require_by;
                        data.SubjectName = Mo.subject_name;
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
                wstSC.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSC.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstSC.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSC.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSC.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstSC.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSC.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
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
            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoErrorList)
            {
                wstSCError.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSCError.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSCError.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSCError.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSCError.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstSCError.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstSCError.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSCError.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstSCError.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSCError.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSCError.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstSCError.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSCError.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
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
            foreach (rptStudSemsScoreCodeChkInfo data in StudSemsScoreCodeChkInfoNoList)
            {
                wstSCNo.Cells[rowIdx, GetColIndex("學號")].PutValue(data.StudentNumber);
                wstSCNo.Cells[rowIdx, GetColIndex("班級")].PutValue(data.ClassName);
                wstSCNo.Cells[rowIdx, GetColIndex("座號")].PutValue(data.SeatNo);
                wstSCNo.Cells[rowIdx, GetColIndex("姓名")].PutValue(data.StudentName);
                wstSCNo.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstSCNo.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstSCNo.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wstSCNo.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstSCNo.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstSCNo.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
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
                iptGradeYear.Value = 1;
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
            btnRun.Enabled = iptGradeYear.Enabled = false;
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
