using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using SHCourseGroupCodeAdmin.DAO;
using System.IO;

namespace SHCourseGroupCodeAdmin.DataCheck
{
    public partial class chkCheckCourseCode : FISCA.Presentation.Controls.BaseForm
    {
        BackgroundWorker _bgWorker;
        Workbook _wb;
        DataAccess da = new DataAccess();
        int _GradeYear = 1;
        List<CourseInfoChk> _CourseInfoChkList;
        Dictionary<string, int> _ColIdxDict;

        public chkCheckCourseCode()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _wb = new Workbook();
            _CourseInfoChkList = new List<CourseInfoChk>();
            _ColIdxDict = new Dictionary<string, int>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;

        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程開課檢查 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程開課檢查 產生完成");
            FuncEnable(true);
            if (_wb != null)
            {
                Utility.ExprotXls("課程開課檢查", _wb);
            }
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _CourseInfoChkList.Clear();
            // 透過年級取得學生修課資料
            _CourseInfoChkList = da.GetCourseCheckInfoListByGradeYear(_GradeYear);

            _bgWorker.ReportProgress(30);

            // 取得課程代碼大表科目內容
            Dictionary<string, List<MOECourseCodeInfo>> MOECoursedDict = da.GetCourseGroupCodeDict();
            _bgWorker.ReportProgress(40);


            List<string> errMesList = new List<string>();
            List<string> errItem = new List<string>();


            // 資料比對
            foreach (CourseInfoChk ci in _CourseInfoChkList)
            {
                errMesList.Clear();
                errItem.Clear();
                errItem.Add("科目名稱");
                errItem.Add("部定校訂");
                errItem.Add("必修選修");

                // 比對群組代碼
                if (MOECoursedDict.ContainsKey(ci.GroupCode))
                {
                    // 資料比對
                    foreach (MOECourseCodeInfo mi in MOECoursedDict[ci.GroupCode])
                    {
                        if (ci.SubjectName == mi.subject_name && ci.IsRequired == mi.is_required && ci.RequireBy == mi.require_by)
                        {
                            ci.course_code = mi.course_code;
                            ci.credit_period = mi.credit_period;
                            errItem.Remove("科目名稱");
                            errItem.Remove("部定校訂");
                            errItem.Remove("必修選修");
                            break;
                        }
                    }

                    foreach (MOECourseCodeInfo mi in MOECoursedDict[ci.GroupCode])
                    {
                        if (ci.course_code == "" && ci.SubjectName == mi.subject_name && ci.IsRequired == mi.is_required )
                        {                          
                            errItem.Remove("科目名稱");           
                            errItem.Remove("必修選修");
                            break;
                        }
                    }

                    foreach (MOECourseCodeInfo mi in MOECoursedDict[ci.GroupCode])
                    {
                        if (ci.course_code == "" && ci.SubjectName == mi.subject_name && ci.RequireBy == mi.require_by)
                        {                           
                            errItem.Remove("科目名稱");
                            errItem.Remove("部定校訂");
                            break;
                        }
                    }

                    foreach (MOECourseCodeInfo mi in MOECoursedDict[ci.GroupCode])
                    {
                        if (ci.course_code == "" && ci.SubjectName == mi.subject_name )
                        {                           
                            errItem.Remove("科目名稱");
                            break;
                        }
                    }


                    if (errItem.Count > 0)
                    {
                        errMesList.Add(string.Join("、", errItem.ToArray()) + " 無法對照");
                    }

                }
                else
                {
                    errMesList.Add("群科班代碼無法對照");
                }
                ci.Memo = string.Join(",", errMesList.ToArray());
            }

            _bgWorker.ReportProgress(70);
            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.課程開課檢查樣版));
            Worksheet wst = _wb.Worksheets[0];
            wst.Name = _GradeYear + "年級";
            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
            }

            int rowIdx = 1;
            foreach (CourseInfoChk data in _CourseInfoChkList)
            {
                wst.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wst.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wst.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wst.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wst.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequireBy);
                wst.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.IsRequired);
                wst.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wst.Cells[rowIdx, GetColIndex("節數")].PutValue(data.Period);
                wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.course_code);
                wst.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wst.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.GroupCode);
                wst.Cells[rowIdx, GetColIndex("說明")].PutValue(data.Memo);
                rowIdx++;
            }

            wst.AutoFitColumns();
            _bgWorker.ReportProgress(100);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            FuncEnable(false);
            _GradeYear = iptGradeYear.Value;
            _bgWorker.RunWorkerAsync();
        }

        private void FuncEnable(bool value)
        {
            iptGradeYear.Enabled = btnRun.Enabled = value;
        }

        private void chkCheckCourseCode_Load(object sender, EventArgs e)
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

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }
    }
}
