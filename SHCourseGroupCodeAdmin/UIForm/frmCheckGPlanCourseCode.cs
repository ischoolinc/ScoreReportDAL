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
using Aspose.Cells;
using System.IO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCheckGPlanCourseCode : BaseForm
    {
        DataAccess da = new DataAccess();
        BackgroundWorker _bwWorker;
        Workbook _wb;
        int _SchoolYear = 0, _Semester = 1;
        Dictionary<string, int> _ColIdxDict;

        public frmCheckGPlanCourseCode()
        {
            InitializeComponent();
            _bwWorker = new BackgroundWorker();

            _bwWorker.DoWork += _bwWorker_DoWork;
            _ColIdxDict = new Dictionary<string, int>();
            _bwWorker.RunWorkerCompleted += _bwWorker_RunWorkerCompleted;
            _bwWorker.ProgressChanged += _bwWorker_ProgressChanged;
            _bwWorker.WorkerReportsProgress = true;
        }

        private void _bwWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生檢核資料 ...", e.ProgressPercentage);
        }

        private void _bwWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            ControlEnable(true);
            if (_wb != null)
            {
                Utility.ExprotXls("開課檢核課程代碼", _wb);
            }
        }

        private void _bwWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bwWorker.ReportProgress(1);
            // 取得課程規劃表與採用班級群科班設定
            List<rptGPlanClassChkInfo> RptGPlanClassChkInfoList = da.GetRptGPlanClassChkInfo(_SchoolYear, _Semester);

            List<rptGPlanClassChkInfo> RptGPlanClassChkPass = new List<rptGPlanClassChkInfo>();
            List<rptGPlanClassChkInfo> RptGPlanClassChkError = new List<rptGPlanClassChkInfo>();

            foreach (rptGPlanClassChkInfo data in RptGPlanClassChkInfoList)
            {
                if (data.ErrorMsgList.Count > 0)
                {
                    RptGPlanClassChkError.Add(data);
                }
                else
                {
                    RptGPlanClassChkPass.Add(data);
                }
            }
            _bwWorker.ReportProgress(20);

            // 取得課程大表資料
            Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = da.GetCourseGroupCodeDict();

            // 取得學分對照表
            Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

            // 取得學年度學期有開課且有修課學生的課程比對課程代碼大表用
            Dictionary<string, List<chkCourseInfo>> hasCourseAttendDict = da.GetCourseHasAttendBySchoolYearSemester(_SchoolYear, _Semester);

            List<string> errItem = new List<string>();

            // 依課程代碼表開課
            List<rptGPlanCourseChkInfo> RptMOECourseChkInfoListPass = new List<rptGPlanCourseChkInfo>();
            List<rptGPlanCourseChkInfo> RptMOECourseChkInfoListError = new List<rptGPlanCourseChkInfo>();
            try
            {
                foreach (string key in MOECourseDict.Keys)
                {
                    foreach (MOECourseCodeInfo data in MOECourseDict[key])
                    {
                        errItem.Clear();
                        errItem.Add("科目名稱");
                        errItem.Add("部定校訂");
                        errItem.Add("必修選修");
                        errItem.Add("學分數");

                        int grYear = _SchoolYear - int.Parse(data.entry_year) + 1;
                        // 學生年級學期

                        int idx = -1;

                        if (grYear == 1 && _Semester == 1)
                            idx = 0;

                        if (grYear == 1 && _Semester == 2)
                            idx = 1;

                        if (grYear == 2 && _Semester == 1)
                            idx = 2;
                        if (grYear == 2 && _Semester == 2)
                            idx = 3;

                        if (grYear == 3 && _Semester == 1)
                            idx = 4;

                        if (grYear == 3 && _Semester == 2)
                            idx = 5;


                        char[] cp = data.open_type.ToArray();
                        if (idx != -1 && idx < cp.Length)
                        {
                            // 需要開課
                            if (cp[idx] != '-')
                            {
                                rptGPlanCourseChkInfo m_data = new rptGPlanCourseChkInfo();
                                m_data.SchoolYear = _SchoolYear + "";
                                m_data.Semester = _Semester + "";
                                m_data.GradeYear = grYear + "";
                                m_data.entry_year = data.entry_year;
                                m_data.CourseCode = data.course_code;
                                m_data.credit_period = data.credit_period;
                                m_data.gdc_code = data.group_code;
                                m_data.open_type = data.open_type;
                                m_data.SubjectName = data.subject_name;
                                m_data.RequiredBy = data.require_by;
                                m_data.isRequired = data.is_required;


                                // 學分數
                                char[] ret = data.credit_period.ToCharArray();
                                if (idx > -1 && idx < ret.Count())
                                {
                                    string x = ret[idx] + "";
                                    if (mappingTable.ContainsKey(x))
                                        m_data.Credit = mappingTable[x];
                                }


                                if (hasCourseAttendDict.ContainsKey(m_data.gdc_code))
                                {
                                    foreach (chkCourseInfo c_info in hasCourseAttendDict[m_data.gdc_code])
                                    {
                                        if (m_data.SubjectName == c_info.SubjectName && m_data.isRequired == c_info.required && m_data.RequiredBy == c_info.required_by && m_data.Credit == c_info.credit)
                                        {
                                            m_data.CourseID = c_info.CourseID;
                                            m_data.CourseName = c_info.CourseName;
                                            errItem.Clear();
                                            break;
                                        }

                                    }

                                    foreach (chkCourseInfo c_info in hasCourseAttendDict[m_data.gdc_code])
                                    {
                                        if (m_data.SubjectName == c_info.SubjectName && m_data.isRequired == c_info.required )
                                        {
                                            errItem.Remove("科目名稱");
                                            errItem.Remove("必修選修");
                                            break;
                                        }
                                    }

                                    foreach (chkCourseInfo c_info in hasCourseAttendDict[m_data.gdc_code])
                                    {
                                        if (m_data.SubjectName == c_info.SubjectName && m_data.RequiredBy == c_info.required_by)
                                        {
                                            errItem.Remove("科目名稱");
                                            errItem.Remove("部定校訂");
                                            break;
                                        }
                                    }

                                    foreach (chkCourseInfo c_info in hasCourseAttendDict[m_data.gdc_code])
                                    {
                                        if (m_data.SubjectName == c_info.SubjectName)
                                        {
                                            errItem.Remove("科目名稱");                                            
                                            break;
                                        }
                                    }                                    
                                    

                                    if (errItem.Count > 0)
                                    {
                                        foreach (string err in errItem)
                                            m_data.ErrorMsgList.Add(err);
                                    }


                                }
                                else
                                {
                                    m_data.ErrorMsgList.Add("群科班代碼 不同");
                                }

                                if (m_data.ErrorMsgList.Count == 0)
                                    RptMOECourseChkInfoListPass.Add(m_data);
                                else
                                    RptMOECourseChkInfoListError.Add(m_data);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }



            // 取得課程檢查資料
            List<rptGPlanCourseChkInfo> RptGPlanCourseChkInfoList = da.GetRptGPlanCourseChkInfo(_SchoolYear, _Semester);
            List<rptGPlanCourseChkInfo> RptGPlanCourseChkInfoPass = new List<rptGPlanCourseChkInfo>();
            List<rptGPlanCourseChkInfo> RptGPlanCourseChkInfoError = new List<rptGPlanCourseChkInfo>();

            foreach (rptGPlanCourseChkInfo data in RptGPlanCourseChkInfoList)
            {
                if (data.ErrorMsgList.Count > 0)
                {
                    RptGPlanCourseChkInfoError.Add(data);
                }
                else
                {
                    // RptGPlanCourseChkInfoPass.Add(data);

                    // 學生年級學期
                    string gr = data.GradeYear;
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



                    char[] cp = data.open_type.ToArray();
                    if (idx != -1 && idx < cp.Length)
                    {
                        // 需要開課
                        if (cp[idx] != '-')
                        {
                            RptGPlanCourseChkInfoPass.Add(data);
                        }
                        else
                        {
                            data.ErrorMsgList.Add("不需要開課。");
                            RptGPlanCourseChkInfoError.Add(data);
                            Console.WriteLine("test");
                        }
                        //thisCousreCodeList.Add(couInfo);
                    }


                }
            }
            _bwWorker.ReportProgress(50);

            // 輸出報表 
            _wb = new Workbook(new MemoryStream(Properties.Resources.課程規劃表開課檢核樣版));

            Worksheet wstMOEChkPass = _wb.Worksheets["檢查依課程代碼表開課"];
            Worksheet wstMOEChkErr = _wb.Worksheets["檢查依課程代碼表開課(有差異)"];
            Worksheet wstGPChkPass = _wb.Worksheets["檢查課程規劃與採用班級群科班"];

            Worksheet wstGPChkErr = _wb.Worksheets["檢查課程規劃與採用班級群科班(有差異)"];

            Worksheet wstGPChkCoursePass = _wb.Worksheets["檢查課程課程代碼_依所屬班級"];

            Worksheet wstGPChkCourseErr = _wb.Worksheets["檢查課程課程代碼_依所屬班級(有差異)"];

            int rowIdx = 1;
            // 讀取欄位與索引            
            for (int co = 0; co <= wstMOEChkPass.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstMOEChkPass.Cells[0, co].StringValue, co);
            }

            foreach (rptGPlanCourseChkInfo data in RptMOECourseChkInfoListPass)
            {
                wstMOEChkPass.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("年級")].PutValue(data.GradeYear);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.isRequired);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstMOEChkPass.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
                rowIdx++;
            }

            _ColIdxDict.Clear();
            rowIdx = 1;
            // 讀取欄位與索引            
            for (int co = 0; co <= wstMOEChkErr.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstMOEChkErr.Cells[0, co].StringValue, co);
            }

            foreach (rptGPlanCourseChkInfo data in RptMOECourseChkInfoListError)
            {
                wstMOEChkErr.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("年級")].PutValue(data.GradeYear);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.isRequired);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstMOEChkErr.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
                //wstMOEChkErr.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + "  ");
                if (data.ErrorMsgList.Contains("群科班代碼 不同"))
                {
                    wstMOEChkErr.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + "");
                }
                else
                {
                    // 科目名稱無法對照
                    if (data.ErrorMsgList.Contains("科目名稱"))
                    {
                        wstMOEChkErr.Cells[rowIdx, GetColIndex("說明")].PutValue("沒有此科目名稱");
                    }
                    else
                    {
                        // 科目名稱有對到其他有問題
                        wstMOEChkErr.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + " 不同");
                    }
                }

                rowIdx++;
            }


            rowIdx = 1;
            _ColIdxDict.Clear();
            // 讀取欄位與索引            
            for (int co = 0; co <= wstGPChkPass.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstGPChkPass.Cells[0, co].StringValue, co);
            }


            foreach (rptGPlanClassChkInfo data in RptGPlanClassChkPass)
            {
                wstGPChkPass.Cells[rowIdx, GetColIndex("課程規劃表名稱")].PutValue(data.GPName);
                wstGPChkPass.Cells[rowIdx, GetColIndex("課程規劃表群科班名稱")].PutValue(data.GPMOEName);
                wstGPChkPass.Cells[rowIdx, GetColIndex("課程規劃表群組代碼")].PutValue(data.GPMOECode);
                wstGPChkPass.Cells[rowIdx, GetColIndex("採用班級")].PutValue(data.ClassName);
                wstGPChkPass.Cells[rowIdx, GetColIndex("班級設定群科班名稱")].PutValue(data.ClassGDCName);
                wstGPChkPass.Cells[rowIdx, GetColIndex("班級設定群科班代碼")].PutValue(data.ClassGDCCode);
                rowIdx++;
            }

            _ColIdxDict.Clear();
            for (int co = 0; co <= wstGPChkErr.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstGPChkErr.Cells[0, co].StringValue, co);
            }

            rowIdx = 1;
            foreach (rptGPlanClassChkInfo data in RptGPlanClassChkError)
            {
                wstGPChkErr.Cells[rowIdx, GetColIndex("課程規劃表名稱")].PutValue(data.GPName);
                wstGPChkErr.Cells[rowIdx, GetColIndex("課程規劃表群科班名稱")].PutValue(data.GPMOEName);
                wstGPChkErr.Cells[rowIdx, GetColIndex("課程規劃表群組代碼")].PutValue(data.GPMOECode);
                wstGPChkErr.Cells[rowIdx, GetColIndex("採用班級")].PutValue(data.ClassName);
                wstGPChkErr.Cells[rowIdx, GetColIndex("班級設定群科班名稱")].PutValue(data.ClassGDCName);
                wstGPChkErr.Cells[rowIdx, GetColIndex("班級設定群科班代碼")].PutValue(data.ClassGDCCode);
                wstGPChkErr.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + "  ");
                rowIdx++;
            }

            _bwWorker.ReportProgress(70);

            _ColIdxDict.Clear();
            for (int co = 0; co <= wstGPChkCoursePass.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstGPChkCoursePass.Cells[0, co].StringValue, co);
            }

            rowIdx = 1;
            foreach (rptGPlanCourseChkInfo data in RptGPlanCourseChkInfoPass)
            {
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("所屬班級")].PutValue(data.ClassName);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.isRequired);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstGPChkCoursePass.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);

                rowIdx++;
            }


            _ColIdxDict.Clear();
            for (int co = 0; co <= wstGPChkCourseErr.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wstGPChkCourseErr.Cells[0, co].StringValue, co);
            }

            rowIdx = 1;
            foreach (rptGPlanCourseChkInfo data in RptGPlanCourseChkInfoError)
            {
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("學年度")].PutValue(data.SchoolYear);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("學期")].PutValue(data.Semester);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("課程名稱")].PutValue(data.CourseName);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("所屬班級")].PutValue(data.ClassName);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(data.SubjectName);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("科目級別")].PutValue(data.SubjectLevel);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(data.RequiredBy);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("必修選修")].PutValue(data.isRequired);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("學分數")].PutValue(data.Credit);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(data.CourseCode);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(data.credit_period);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("開課方式")].PutValue(data.open_type);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(data.gdc_code);
                wstGPChkCourseErr.Cells[rowIdx, GetColIndex("說明")].PutValue(string.Join(",", data.ErrorMsgList.ToArray()) + "  不同");
                rowIdx++;
            }


            wstGPChkPass.AutoFitColumns();
            wstGPChkErr.AutoFitColumns();
            wstGPChkCoursePass.AutoFitColumns();
            wstGPChkCourseErr.AutoFitColumns();

            // 沒有錯誤移除工作表
            if (RptGPlanClassChkError.Count == 0)
            {
                _wb.Worksheets.RemoveAt("檢查課程規劃與採用班級群科班(有差異)");
            }
            if (RptGPlanCourseChkInfoError.Count == 0)
            {
                _wb.Worksheets.RemoveAt("檢查課程課程代碼_依所屬班級(有差異)");
            }

            if (RptMOECourseChkInfoListError.Count == 0)
            {
                _wb.Worksheets.RemoveAt("檢查依課程代碼表開課(有差異)");
            }
            _bwWorker.ReportProgress(100);
        }

        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }

        private void frmCheckGPlanCourseCode_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            ControlEnable(false);

            // 載入學年度、學期選項
            int schoolYear;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out schoolYear))
            {
                for (int sc = (schoolYear - 5); sc <= (schoolYear + 5); sc++)
                {
                    cbxSchoolYear.Items.Add(sc);
                }
                cbxSchoolYear.Text = K12.Data.School.DefaultSchoolYear;

                cbxSemester.Items.Add("1");
                cbxSemester.Items.Add("2");

                cbxSemester.Text = K12.Data.School.DefaultSemester;
            }

            // 說明文字
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("說明：");
            sb.AppendLine("1.檢查依課程代碼表為主，檢查課程是否正確開出，課程內需要有修課學生，課程會用學生群科班，再比對群科班內：科目名稱＋部定校訂＋必選修＋學分數，相同課程放在工作表[檢查依課程代碼表開課]，差異放在工作表[檢查依課程代碼表開課(有差異)]。");
            sb.AppendLine("2.檢查課程規劃與採用班級群科班是否與班級群科班設定相同，相同放在工作表[檢查課程規劃與採用班級群科班]，差異放在工作表[檢查課程規劃與採用班級群科班(有差異)]。");
            sb.AppendLine("3.檢查課程內所屬班級開課的課程代碼大表比對後差異，相同放在工作表[檢查課程課程代碼_依所屬班級]，差異放在工作表[檢查課程課程代碼_依所屬班級(有差異)]。");
            

            txtDesc.Text = sb.ToString();

            ControlEnable(true);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            ControlEnable(false);
            int sy, ss;
            if (int.TryParse(cbxSchoolYear.Text, out sy) && int.TryParse(cbxSemester.Text, out ss))
            {
                _SchoolYear = sy;
                _Semester = ss;
            }
            _bwWorker.RunWorkerAsync();
        }

        private void ControlEnable(bool value)
        {
            cbxSchoolYear.Enabled = cbxSemester.Enabled = btnRun.Enabled = value;
        }

    }
}
