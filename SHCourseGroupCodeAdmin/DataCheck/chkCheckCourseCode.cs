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
            da.LoadMOEGroupCodeDict();
            _CourseInfoChkList.Clear();
            // 透過年級取得學生修課資料
            _CourseInfoChkList = da.GetCourseCheckInfoListByGradeYear(_GradeYear);

            _bgWorker.ReportProgress(30);

            // 取得課程代碼大表科目內容
            Dictionary<string, List<MOECourseCodeInfo>> MOECoursedDict = da.GetCourseGroupCodeDict();
            _bgWorker.ReportProgress(40);

            // 取得課程規劃ID與gdc_code對應，由學生來過濾
            Dictionary<string, List<string>> CPIdGdcCodeDict = da.GetCPIdGdcCodeDict(_GradeYear);

            //取得有使用的課程規劃大表
            Dictionary<string, GPlanInfo> GPlanInfoDict = da.GetGPlanInfoDictByGPID(CPIdGdcCodeDict.Keys.ToList());

            Dictionary<string, MOECourseCodeInfo> tmpMoeDict = new Dictionary<string, MOECourseCodeInfo>();
            List<string> errMesList = new List<string>();
            List<string> errItem = new List<string>();

            // 取得學分對照表
            Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

            // 課程規劃表比對
            foreach (string id in GPlanInfoDict.Keys)
            {
                GPlanInfo data = GPlanInfoDict[id];

                // 使用課程規劃表ID group_code
                if (CPIdGdcCodeDict.ContainsKey(data.ID))
                {
                    foreach (string gdc_code in CPIdGdcCodeDict[data.ID])
                    {
                        errMesList.Clear();
                        bool hasGdcCode = false;

                        // 取得課程代碼大表
                        if (MOECoursedDict.ContainsKey(gdc_code))
                        {
                            tmpMoeDict.Clear();

                            foreach (MOECourseCodeInfo mdata in MOECoursedDict[gdc_code])
                            {
                                string key = mdata.subject_name + "_" + mdata.require_by + "_" + mdata.is_required;

                                if (!tmpMoeDict.ContainsKey(key))
                                    tmpMoeDict.Add(key, mdata);
                            }
                            hasGdcCode = true;
                        }
                        else
                        {
                            errMesList.Add("群科班代碼無法對照");
                            hasGdcCode = false;

                            foreach (GPCourseInfo gpCo in data.CourseInfoList)
                            {
                                gpCo.Memo = string.Join(",", errMesList.ToArray());
                            }
                        }

                        if (hasGdcCode)
                        {
                            // 掃課程規劃表資料比對
                            foreach (GPCourseInfo gpCo in data.CourseInfoList)
                            {
                                errMesList.Clear();
                                errItem.Clear();
                                errItem.Add("科目名稱");
                                errItem.Add("部定校訂");
                                errItem.Add("必修選修");
                                gpCo.Memo = "";
                                // 使用科目名稱+部校訂+必選修
                                if (tmpMoeDict.ContainsKey(gpCo.tmpKey))
                                {
                                    MOECourseCodeInfo md = tmpMoeDict[gpCo.tmpKey];
                                    gpCo.GroupCode = md.group_code;
                                    gpCo.CourseCode = md.course_code;
                                    gpCo.GroupName = da.GetGroupNameByCode(gpCo.GroupCode);
                                    gpCo.credit_period = md.credit_period;
                                  
                                    errItem.Remove("科目名稱");
                                    errItem.Remove("部定校訂");
                                    errItem.Remove("必修選修");
                                }
                                else
                                {
                                    // 比對不到原因
                                    foreach (MOECourseCodeInfo mm in tmpMoeDict.Values)
                                    {
                                        if (gpCo.SubjectName == mm.subject_name && gpCo.Required == mm.is_required)
                                        {
                                            errItem.Remove("科目名稱");
                                            errItem.Remove("必修選修");
                                            break;
                                        }
                                    }

                                    foreach (MOECourseCodeInfo mm in tmpMoeDict.Values)
                                    {
                                        if (gpCo.SubjectName == mm.subject_name && gpCo.RequiredBy == mm.require_by)
                                        {
                                            errItem.Remove("科目名稱");
                                            errItem.Remove("部定校訂");
                                            break;
                                        }
                                    }

                                    foreach (MOECourseCodeInfo mm in tmpMoeDict.Values)
                                    {
                                        if (gpCo.SubjectName == mm.subject_name)
                                        {
                                            errItem.Remove("科目名稱");
                                            break;
                                        }
                                    }
                                }

                                if (errItem.Count > 0)
                                {
                                    errMesList.Add(string.Join("、", errItem.ToArray()) + " 無法對照");
                                }
                                if (gpCo.CourseCode == "")
                                    gpCo.Memo = string.Join(",", errMesList.ToArray());
                            }
                        }

                    }
                }
            }
            
            // 已開修課資料比對
            foreach (CourseInfoChk ci in _CourseInfoChkList)
            {
                errMesList.Clear();
                errItem.Clear();
                errItem.Add("科目名稱");
                errItem.Add("部定校訂");
                errItem.Add("必修選修");
                errItem.Add("學分數");

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
                            ci.entry_year = mi.entry_year;
                            if (ci.CheckCreditPass(mappingTable))
                            {
                                errItem.Remove("學分數");
                            }
                            errItem.Remove("科目名稱");
                            errItem.Remove("部定校訂");
                            errItem.Remove("必修選修");
                            break;
                        }
                    }

                    foreach (MOECourseCodeInfo mi in MOECoursedDict[ci.GroupCode])
                    {
                        if (ci.course_code == "" && ci.SubjectName == mi.subject_name && ci.IsRequired == mi.is_required)
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
                        if (ci.course_code == "" && ci.SubjectName == mi.subject_name)
                        {
                            errItem.Remove("科目名稱");
                            break;
                        }
                    }


                    if (errItem.Count > 0)
                    {
                        errMesList.Add(string.Join("、", errItem.ToArray())+ " 不同。");
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
            Worksheet wst = _wb.Worksheets["已開課課程"];
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


            // 處理課程規劃表比對結果
            Worksheet wst2 = _wb.Worksheets["課程規劃表"];

            _ColIdxDict.Clear();

            // 讀取欄位與索引            
            for (int co = 0; co <= wst2.Cells.MaxDataColumn; co++)
            {
                _ColIdxDict.Add(wst2.Cells[0, co].StringValue, co);
            }

            rowIdx = 1;
            // 產生資料，先過濾重複資料
            Dictionary<string, GPCourseInfo> tmpCourseInfoDcit = new Dictionary<string, GPCourseInfo>();

            foreach (string id in GPlanInfoDict.Keys)
            {
                GPlanInfo data = GPlanInfoDict[id];

                foreach (GPCourseInfo gpCo in data.CourseInfoList)
                {
                    string key = gpCo.GPName + "_" + gpCo.tmpKey + "_" + gpCo.CourseCode;
                    if (!tmpCourseInfoDcit.ContainsKey(key))
                        tmpCourseInfoDcit.Add(key, gpCo);
                }
            }

            // 產生到 Excel
            foreach (string key in tmpCourseInfoDcit.Keys)
            {
                GPCourseInfo gpCo = tmpCourseInfoDcit[key];
                wst2.Cells[rowIdx, GetColIndex("課程規劃名稱")].PutValue(gpCo.GPName);
                wst2.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(gpCo.SubjectName);
                wst2.Cells[rowIdx, GetColIndex("部定校訂")].PutValue(gpCo.RequiredBy);
                wst2.Cells[rowIdx, GetColIndex("必修選修")].PutValue(gpCo.Required);
                wst2.Cells[rowIdx, GetColIndex("學分")].PutValue(gpCo.Credit);
                wst2.Cells[rowIdx, GetColIndex("群科班代碼")].PutValue(gpCo.GroupCode);
                wst2.Cells[rowIdx, GetColIndex("群科班名稱")].PutValue(gpCo.GroupName);
                wst2.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(gpCo.CourseCode);
                wst2.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(gpCo.credit_period);
                wst2.Cells[rowIdx, GetColIndex("說明")].PutValue(gpCo.Memo);
                rowIdx++;
            }

            wst2.AutoFitColumns();

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
